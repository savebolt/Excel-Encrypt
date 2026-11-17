#include "word.h"

#include <tlhelp32.h>

#include <chrono>
#include <filesystem>
#include <set>

#include "util.h"

namespace {

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

std::set<DWORD> WordProcessIds() {
    std::set<DWORD> ids;
    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snapshot == INVALID_HANDLE_VALUE) {
        return ids;
    }
    PROCESSENTRY32W entry{};
    entry.dwSize = sizeof(entry);
    if (Process32FirstW(snapshot, &entry)) {
        do {
            if (_wcsicmp(entry.szExeFile, L"WINWORD.EXE") == 0) {
                ids.insert(entry.th32ProcessID);
            }
        } while (Process32NextW(snapshot, &entry));
    }
    CloseHandle(snapshot);
    return ids;
}

long SaveFormatFor(const std::wstring& path) {
    const std::wstring extension = util::ToLower(std::filesystem::path(path).extension().wstring());
    if (extension == L".docx") return 12;  // wdFormatXMLDocument
    if (extension == L".docm") return 13;  // wdFormatXMLDocumentMacroEnabled
    if (extension == L".dotx") return 14;  // wdFormatXMLTemplate
    if (extension == L".dotm") return 15;  // wdFormatXMLTemplateMacroEnabled
    if (extension == L".dot")  return 1;   // wdFormatTemplate
    return 0;                              // wdFormatDocument (.doc)
}

}  // namespace

WordApp::~WordApp() {
    Stop();
}

void WordApp::Start() {
    if (app_) {
        return;
    }

    CLSID clsid{};
    HRESULT hr = CLSIDFromProgID(L"Word.Application", &clsid);
    if (FAILED(hr)) {
        throw com::Error(hr, L"Microsoft Word is not installed or not registered on this computer");
    }

    const std::set<DWORD> before = WordProcessIds();

    IDispatch* instance = nullptr;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                          reinterpret_cast<void**>(&instance));
    if (FAILED(hr) || instance == nullptr) {
        if (instance != nullptr) {
            instance->Release();
        }
        throw com::Error(FAILED(hr) ? hr : E_FAIL,
                         L"Microsoft Word could not be started. Is Word installed?");
    }

    app_.Attach(instance);
    Configure();

    // Word starts a fresh WINWORD.EXE for every automation instance, so the
    // process that appeared while we connected is ours to close.
    processId_ = 0;
    const std::set<DWORD> after = WordProcessIds();
    for (DWORD id : after) {
        if (before.find(id) == before.end()) {
            processId_ = id;
            break;
        }
    }
    watchdog_.Start(processId_);
}

void WordApp::Stop() {
    StopInternal(true);
}

void WordApp::StopInternal(bool quitWord) {
    if (app_) {
        if (quitWord) {
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

void WordApp::Configure() {
    auto setBool = [this](const wchar_t* name, bool value) {
        try {
            com::PutBool(app_.Get(), name, value);
        } catch (...) {
        }
    };
    setBool(L"Visible", false);
    setBool(L"ScreenUpdating", false);
    try {
        com::PutLong(app_.Get(), L"DisplayAlerts", 0);  // wdAlertsNone
    } catch (...) {
    }
    try {
        com::PutLong(app_.Get(), L"AutomationSecurity", 3);  // force-disable macros
    } catch (...) {
    }
    try {
        com::ComPtr options = com::GetObject(app_.Get(), L"Options");
        com::PutBool(options.Get(), L"UpdateLinksAtOpen", false);
    } catch (...) {
    }
}

FileResult WordApp::ProcessFile(const std::wstring& path, const std::wstring& password,
                                bool encrypt) {
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

    com::ComPtr document;
    bool opened = false;
    bool closeFailed = false;
    std::wstring failure;
    std::filesystem::path tempPath;

    try {
        com::ComPtr documents = com::GetObject(app_.Get(), L"Documents");

        com::ArgList openArgs(16);
        openArgs.SetString(1, path);       // FileName
        openArgs.SetBool(2, false);        // ConfirmConversions
        openArgs.SetBool(3, false);        // ReadOnly
        openArgs.SetBool(4, false);        // AddToRecentFiles
        openArgs.SetString(5, password);   // PasswordDocument
        openArgs.SetString(6, L"");        // PasswordTemplate
        openArgs.SetBool(7, false);        // Revert
        openArgs.SetString(8, L"");        // WritePasswordDocument
        openArgs.SetString(9, L"");        // WritePasswordTemplate
        openArgs.SetMissing(10);           // Format
        openArgs.SetMissing(11);           // Encoding
        openArgs.SetBool(12, false);       // Visible
        openArgs.SetBool(13, false);       // OpenAndRepair
        openArgs.SetMissing(14);           // DocumentDirection
        openArgs.SetBool(15, true);        // NoEncodingDialog
        openArgs.SetMissing(16);           // XMLTransform

        document = com::CallObject(documents.Get(), L"Open", openArgs);
        opened = true;

        const bool hasPassword = com::GetBool(document.Get(), L"HasPassword");
        if (encrypt && hasPassword) {
            result.action = FileAction::SkippedAlreadyEncrypted;
        } else if (!encrypt && !hasPassword) {
            result.action = FileAction::SkippedNotEncrypted;
        } else {
            if (!encrypt) {
                // Clear both passwords, otherwise Word keeps them on save.
                com::PutString(document.Get(), L"Password", L"");
                com::PutString(document.Get(), L"WritePassword", L"");
            }

            // Word ignores the Password argument when saving over the same
            // file, so write to a sibling temp file and swap it in after a
            // successful, verified save.
            const std::filesystem::path original(path);
            tempPath = original;
            tempPath += L".xlpt-tmp";
            tempPath += original.extension();

            com::ArgList saveArgs(17);
            saveArgs.SetString(1, tempPath.wstring());        // FileName
            saveArgs.SetLong(2, SaveFormatFor(path));         // FileFormat
            saveArgs.SetBool(3, false);                       // LockComments
            saveArgs.SetString(4, encrypt ? password : L"");  // Password
            saveArgs.SetBool(5, false);                       // AddToRecentFiles
            saveArgs.SetString(6, L"");                       // WritePassword
            saveArgs.SetBool(7, false);                       // ReadOnlyRecommended
            saveArgs.SetMissing(8);                           // EmbedTrueTypeFonts
            saveArgs.SetMissing(9);                           // SaveNativePictureFormat
            saveArgs.SetMissing(10);                          // SaveFormsData
            saveArgs.SetMissing(11);                          // SaveAsAOCELetter
            saveArgs.SetMissing(12);                          // Encoding
            saveArgs.SetMissing(13);                          // InsertLineBreaks
            saveArgs.SetMissing(14);                          // AllowSubstitutions
            saveArgs.SetMissing(15);                          // LineEnding
            saveArgs.SetMissing(16);                          // AddBiDiMarks
            saveArgs.SetMissing(17);                          // CompatibilityMode
            com::CallVoid(document.Get(), L"SaveAs2", saveArgs);

            com::ArgList closeArgs(1);
            closeArgs.SetLong(1, 0);  // wdDoNotSaveChanges
            com::CallVoid(document.Get(), L"Close", closeArgs);
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
                                    MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
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

    if (opened && document) {
        try {
            com::ArgList closeArgs(1);
            closeArgs.SetLong(1, 0);
            com::CallVoid(document.Get(), L"Close", closeArgs);
        } catch (...) {
            closeFailed = true;
        }
    }

    watchdog_.Disarm();

    if (watchdog_.Triggered()) {
        result.action = FileAction::Failed;
        result.detail = L"Word stopped responding and was closed";
        if (!failure.empty()) {
            result.detail += L" (" + failure + L")";
        }
        StopInternal(false);
        return result;
    }

    if (closeFailed) {
        Stop();
        result.action = FileAction::Failed;
        result.detail = failure.empty() ? L"the document could not be closed" : failure;
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
