#include "excel.h"

#include <chrono>
#include <filesystem>

#include "util.h"

namespace {

// Excel can block forever in a modal dialog (locked file, repair prompt, ...).
// A file is given this long before the Excel process is terminated. The
// XLPT_TIMEOUT environment variable (seconds) can override the default.
std::chrono::seconds FileTimeout() {
    wchar_t buffer[32] = {};
    const DWORD length = GetEnvironmentVariableW(L"XLPT_TIMEOUT", buffer, 32);
    if (length > 0 && length < 32) {
        wchar_t* end = nullptr;
        const long value = wcstol(buffer, &end, 10);
        if (end != buffer && value >= 5 && value <= 3600) {
            return std::chrono::seconds(value);
        }
    }
    return std::chrono::seconds(180);
}

long SaveFormatFor(const std::wstring& path) {
    const std::wstring extension = util::ToLower(std::filesystem::path(path).extension().wstring());
    if (extension == L".xlsx") return 51;  // xlOpenXMLWorkbook
    if (extension == L".xlsm") return 52;  // xlOpenXMLWorkbookMacroEnabled
    if (extension == L".xlsb") return 50;  // xlExcel12
    if (extension == L".xltx") return 54;  // xlOpenXMLTemplate
    if (extension == L".xltm") return 53;  // xlOpenXMLTemplateMacroEnabled
    if (extension == L".xlt")  return 17;  // xlTemplate
    if (extension == L".xlam") return 55;  // xlAddIn
    return 56;                             // xlExcel8 (.xls)
}

}  // namespace

ExcelApp::~ExcelApp() {
    Stop();
}

void ExcelApp::Start() {
    if (app_) {
        return;
    }

    CLSID clsid{};
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        throw com::Error(hr, L"Microsoft Excel is not installed or not registered on this computer");
    }

    IDispatch* instance = nullptr;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                          reinterpret_cast<void**>(&instance));
    if (FAILED(hr) || instance == nullptr) {
        if (instance != nullptr) {
            instance->Release();
        }
        throw com::Error(FAILED(hr) ? hr : E_FAIL,
                         L"Microsoft Excel could not be started. Is Excel installed?");
    }

    app_.Attach(instance);
    Configure();

    processId_ = 0;
    try {
        const long window = com::GetLong(app_.Get(), L"Hwnd");
        if (window != 0) {
            DWORD pid = 0;
            GetWindowThreadProcessId(reinterpret_cast<HWND>(static_cast<intptr_t>(window)), &pid);
            processId_ = pid;
        }
    } catch (...) {
        processId_ = 0;
    }
    watchdog_.Start(processId_);
}

void ExcelApp::Stop() {
    StopInternal(true);
}

void ExcelApp::StopInternal(bool quitExcel) {
    if (app_) {
        if (quitExcel) {
            watchdog_.Arm(std::chrono::seconds(30));
            try {
                com::ArgList args(0);
                com::CallVoid(app_.Get(), L"Quit", args);
            } catch (...) {
            }
            watchdog_.Disarm();
        }
        app_.Reset();
    }
    watchdog_.Stop();
    processId_ = 0;
}

void ExcelApp::Configure() {
    auto setBool = [this](const wchar_t* name, bool value) {
        try {
            com::PutBool(app_.Get(), name, value);
        } catch (...) {
        }
    };
    setBool(L"Visible", false);
    setBool(L"DisplayAlerts", false);
    setBool(L"ScreenUpdating", false);
    setBool(L"EnableEvents", false);
    setBool(L"Interactive", false);
    setBool(L"AskToUpdateLinks", false);
    setBool(L"DisplayStatusBar", false);
    setBool(L"EnableAnimations", false);
    setBool(L"AlertBeforeOverwriting", false);
    try {
        com::PutLong(app_.Get(), L"AutomationSecurity", 3);  // msoAutomationSecurityForceDisable
    } catch (...) {
    }
    try {
        com::PutLong(app_.Get(), L"FeatureInstall", 0);  // msoFeatureInstallNone
    } catch (...) {
    }
    try {
        // AutoRecover writes a copy on every save; pointless for a batch run.
        com::ComPtr autoRecover = com::GetObject(app_.Get(), L"AutoRecover");
        com::PutBool(autoRecover.Get(), L"Enabled", false);
    } catch (...) {
    }
}

FileResult ExcelApp::ProcessFile(const std::wstring& path, const std::wstring& password,
                                 bool encrypt, const std::filesystem::path& tempDir) {
    FileResult result;
    result.action = FileAction::Failed;

    if (!app_) {
        try {
            Start();
        } catch (const com::Error& error) {
            result.detail = error.Message();
            return result;
        }
    }

    watchdog_.Arm(FileTimeout());

    com::ComPtr workbook;
    bool opened = false;
    bool closeFailed = false;
    std::wstring failure;
    std::filesystem::path tempPath;

    try {
        com::ComPtr workbooks = com::GetObject(app_.Get(), L"Workbooks");

        com::ArgList openArgs(15);
        openArgs.SetString(1, path);       // Filename
        openArgs.SetLong(2, 0);            // UpdateLinks: never update
        openArgs.SetBool(3, false);        // ReadOnly
        openArgs.SetMissing(4);            // Format
        openArgs.SetString(5, password);   // Password (ignored if not encrypted)
        openArgs.SetString(6, L"");        // WriteResPassword
        openArgs.SetBool(7, true);         // IgnoreReadOnlyRecommended
        openArgs.SetMissing(8);            // Origin
        openArgs.SetMissing(9);            // Delimiter
        openArgs.SetBool(10, false);       // Editable
        openArgs.SetBool(11, false);       // Notify
        openArgs.SetMissing(12);           // Converter
        openArgs.SetBool(13, false);       // AddToMru
        openArgs.SetMissing(14);           // Local
        openArgs.SetMissing(15);           // CorruptLoad

        workbook = com::CallObject(workbooks.Get(), L"Open", openArgs);
        opened = true;

        // Excel only accepts this once a workbook is open.
        try {
            com::PutLong(app_.Get(), L"Calculation", -4135);  // xlCalculationManual
        } catch (...) {
        }

        const bool hasPassword = com::GetBool(workbook.Get(), L"HasPassword");
        if (encrypt && hasPassword) {
            result.action = FileAction::SkippedAlreadyEncrypted;
        } else if (!encrypt && !hasPassword) {
            result.action = FileAction::SkippedNotEncrypted;
        } else {
            if (!encrypt) {
                // Clear the passwords before saving, otherwise Excel keeps them.
                com::PutString(workbook.Get(), L"Password", L"");
                com::PutString(workbook.Get(), L"WritePassword", L"");
            }

            // Excel takes roughly twice as long to overwrite the file it has
            // open, so write a sibling temp file and swap it in afterwards.
            const std::filesystem::path original(path);
            if (!tempDir.empty()) {
                tempPath = tempDir / original.filename();
            } else {
                tempPath = original;
                tempPath += L".xlpt-tmp";
                tempPath += original.extension();
            }

            com::ArgList saveArgs(6);
            saveArgs.SetString(1, tempPath.wstring());        // Filename
            saveArgs.SetLong(2, SaveFormatFor(path));         // FileFormat
            saveArgs.SetString(3, encrypt ? password : L"");  // Password
            saveArgs.SetString(4, L"");                       // WriteResPassword
            saveArgs.SetBool(5, false);                       // ReadOnlyRecommended
            saveArgs.SetBool(6, false);                       // CreateBackup
            com::CallVoid(workbook.Get(), L"SaveAs", saveArgs);

            com::ArgList closeArgs(1);
            closeArgs.SetBool(1, false);
            com::CallVoid(workbook.Get(), L"Close", closeArgs);
            opened = false;

            bool good = true;
            if (util::IsOoxml(original)) {
                const util::Container container = util::DetectContainer(tempPath);
                good = encrypt ? container == util::Container::Cfb
                               : container == util::Container::Zip;
            }
            if (!good) {
                util::RemoveQuiet(tempPath);
                failure = L"the saved file does not look right - the original was left untouched";
            } else if (!MoveFileExW(tempPath.c_str(), original.c_str(),
                                    MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH |
                                        MOVEFILE_COPY_ALLOWED)) {
                const DWORD code = GetLastError();
                util::RemoveQuiet(tempPath);
                failure = L"could not replace the file: " + util::Win32Text(code);
            } else {
                result.action = encrypt ? FileAction::Encrypted : FileAction::Decrypted;
            }
        }
    } catch (const com::Error& error) {
        failure = error.Message();
    }

    if (opened && workbook) {
        try {
            com::ArgList closeArgs(1);
            closeArgs.SetBool(1, false);
            com::CallVoid(workbook.Get(), L"Close", closeArgs);
        } catch (...) {
            closeFailed = true;
        }
    }

    watchdog_.Disarm();

    if (watchdog_.Triggered()) {
        result.action = FileAction::Failed;
        result.detail = L"Excel stopped responding and was closed";
        if (!failure.empty()) {
            result.detail += L" (" + failure + L")";
        }
        StopInternal(false);
        return result;
    }

    if (closeFailed) {
        Stop();  // leave Excel in a clean state before the next file
        result.action = FileAction::Failed;
        result.detail = failure.empty() ? L"the workbook could not be closed" : failure;
        return result;
    }

    if (!failure.empty()) {
        if (!tempPath.empty()) {
            util::RemoveQuiet(tempPath);
        }
        result.action = FileAction::Failed;
        result.detail = failure;
        return result;
    }

    return result;
}
