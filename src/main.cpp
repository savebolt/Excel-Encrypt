#include <windows.h>
#include <shellapi.h>
#include <shobjidl.h>

#include <atomic>
#include <chrono>
#include <cwchar>
#include <filesystem>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <utility>
#include <vector>

#include "excel.h"
#include "pdf.h"
#include "util.h"
#include "word.h"

namespace fs = std::filesystem;
using util::Color;
using util::Paint;

namespace {

// Keep  - leave name.ext.bak next to the file (default)
// Temp  - create the backup, delete it again when the file was processed OK
// None  - do not create backups at all
enum class BackupMode { Keep, Temp, None };

struct Options {
    fs::path folder;
    bool recursive = false;
    BackupMode backups = BackupMode::Keep;
    util::FileTypes types;
    int jobs = 0;  // 0 = choose automatically
};

std::wstring BackupLabel(BackupMode mode) {
    switch (mode) {
        case BackupMode::Keep: return L"ON (keep .bak)";
        case BackupMode::Temp: return L"TEMP (delete on success)";
        default:               return L"OFF";
    }
}

BackupMode NextBackupMode(BackupMode mode) {
    switch (mode) {
        case BackupMode::Keep: return BackupMode::Temp;
        case BackupMode::Temp: return BackupMode::None;
        default:               return BackupMode::Keep;
    }
}

struct BatchStats {
    int done = 0;
    int skipped = 0;
    int failed = 0;
};

enum class OneShot { None, Encrypt, Decrypt };

struct CommandLineOptions {
    Options options;
    OneShot oneShot = OneShot::None;
    std::wstring password;
    bool assumeYes = false;
    bool showHelp = false;
    bool bad = false;
    std::wstring error;
};

void PrintSeparator() {
    util::WriteLine(Paint(L"   " + std::wstring(60, L'-'), Color::Dim));
}

void PrintBanner() {
    util::WriteLine();
    util::WriteLine(Paint(L"   " + std::wstring(60, L'='), Color::Cyan));
    util::WriteLine(Paint(L"    Excel Password Tool   v1.3", Color::Bold));
    util::WriteLine(L"    Encrypt / decrypt Excel, Word and PDF files with a password");
    util::WriteLine(Paint(L"   " + std::wstring(60, L'='), Color::Cyan));
}

std::wstring TypeSummary(const util::FileTypes& types) {
    std::wstring text;
    auto add = [&text](const wchar_t* name) {
        if (!text.empty()) text += L" + ";
        text += name;
    };
    if (types.excel) add(L"Excel");
    if (types.word) add(L"Word");
    if (types.pdf) add(L"PDF");
    return text.empty() ? L"none" : text;
}

void CountTypes(const std::vector<fs::path>& files, int& excel, int& word, int& pdf) {
    excel = word = pdf = 0;
    for (const fs::path& file : files) {
        switch (util::TypeOf(file)) {
            case util::FileType::Excel: ++excel; break;
            case util::FileType::Word:  ++word;  break;
            case util::FileType::Pdf:   ++pdf;   break;
            default: break;
        }
    }
}

void PrintStatus(const Options& options, const util::ScanResult& scan) {
    int excel = 0, word = 0, pdf = 0;
    CountTypes(scan.files, excel, word, pdf);

    util::WriteLine(L"   Folder      : " + options.folder.wstring());
    util::WriteLine(L"   Files       : " + std::to_wstring(scan.files.size()) +
                    L"   (Excel " + std::to_wstring(excel) +
                    L", Word " + std::to_wstring(word) +
                    L", PDF " + std::to_wstring(pdf) + L")");
    util::WriteLine(L"   Types       : " + TypeSummary(options.types));
    util::WriteLine(L"   Subfolders  : " + std::wstring(options.recursive ? L"yes" : L"no"));
    util::WriteLine(L"   Backups     : " + BackupLabel(options.backups));
    for (const std::wstring& error : scan.errors) {
        util::WriteLine(Paint(L"   ! " + error, Color::Red));
    }
    PrintSeparator();
}

void PrintMenu(const Options& options) {
    util::WriteLine(L"   [1] Encrypt all files  (set a password)");
    util::WriteLine(L"   [2] Decrypt all files  (remove the password)");
    util::WriteLine(L"   [3] Change folder");
    util::WriteLine(L"   [4] Include subfolders : " +
                    std::wstring(options.recursive ? L"ON" : L"OFF") + L"   (toggle)");
    util::WriteLine(L"   [5] Backups            : " + BackupLabel(options.backups) +
                    L"   (cycle)");
    util::WriteLine(L"   [6] List files");
    util::WriteLine(L"   [7] File types         : " + TypeSummary(options.types) + L"   (change)");
    util::WriteLine(L"   [8] Browse for folder...");
    util::WriteLine(L"   [0] Exit");
    util::WriteLine();
}

std::wstring DescribeFile(const fs::path& file) {
    std::error_code ec;
    const std::uintmax_t size = fs::file_size(file, ec);
    const std::wstring sizeText = ec ? L"?" : util::FormatSize(size);

    const util::FileType type = util::TypeOf(file);
    if (type == util::FileType::Pdf) {
        return sizeText + L"   PDF";
    }
    if (util::IsOoxml(file)) {
        const util::Container container = util::DetectContainer(file);
        if (container == util::Container::Zip) {
            return sizeText + L"   open XML, not encrypted";
        }
        if (container == util::Container::Cfb) {
            return sizeText + L"   open XML, ENCRYPTED";
        }
        return sizeText + L"   unknown file content";
    }
    if (type == util::FileType::Word) {
        return sizeText + L"   legacy .doc, encryption cannot be detected from outside";
    }
    return sizeText + L"   legacy .xls, encryption cannot be detected from outside";
}

void ListFiles(const Options& options, const util::ScanResult& scan) {
    if (scan.files.empty()) {
        util::WriteLine(L"\n   No matching files found.\n");
        return;
    }
    util::WriteLine();
    const int total = static_cast<int>(scan.files.size());
    const int digits = static_cast<int>(std::to_wstring(total).size());
    for (int index = 0; index < total; ++index) {
        const fs::path& file = scan.files[static_cast<std::size_t>(index)];
        std::wstring number = std::to_wstring(index + 1);
        if (static_cast<int>(number.size()) < digits) {
            number.insert(0, static_cast<std::size_t>(digits) - number.size(), L' ');
        }
        std::wstring name = util::Fit(util::RelativeForDisplay(file, options.folder), 44);
        if (name.size() < 44) {
            name.append(44 - name.size(), L' ');
        }
        util::WriteLine(L"   " + number + L") " + name + L"   " + DescribeFile(file));
    }
    util::WriteLine();
}

std::wstring AskNewPassword() {
    for (int attempt = 1; attempt <= 3; ++attempt) {
        std::wstring first;
        std::wstring second;
        if (!util::ReadPassword(L"\n   Password : ", first)) {
            return {};
        }
        if (first.empty()) {
            util::WriteLine(Paint(L"   The password cannot be empty.", Color::Red));
            continue;
        }
        if (first.size() > 255) {
            util::WriteLine(Paint(L"   Excel passwords are limited to 255 characters.", Color::Red));
            continue;
        }
        if (!util::ReadPassword(L"   Confirm  : ", second)) {
            return {};
        }
        if (first != second) {
            util::WriteLine(Paint(L"   Passwords do not match.", Color::Red));
            continue;
        }
        return first;
    }
    util::WriteLine(L"   Cancelled.");
    return {};
}

std::wstring AskExistingPassword() {
    std::wstring password;
    if (!util::ReadPassword(L"\n   Password: ", password)) {
        return {};
    }
    if (password.empty()) {
        util::WriteLine(Paint(L"   A password is required to decrypt.", Color::Red));
        return {};
    }
    return password;
}

// Office writes its own scratch files next to the target, so two workers
// saving into the same folder can collide. Each worker gets a private folder.
fs::path MakeWorkerTempDir(int workerIndex) {
    wchar_t buffer[MAX_PATH] = {};
    const DWORD length = GetTempPathW(MAX_PATH, buffer);
    const fs::path base = length > 0 ? fs::path(std::wstring(buffer, length))
                                     : fs::temp_directory_path();
    const fs::path dir = base / (L"ExcelPasswordTool-" +
                                 std::to_wstring(GetCurrentProcessId()) + L"-" +
                                 std::to_wstring(workerIndex));
    std::error_code ec;
    fs::remove_all(dir, ec);
    fs::create_directories(dir, ec);
    if (!fs::is_directory(dir, ec)) {
        return {};  // fall back to a sibling temp file
    }
    return dir;
}

int WorkerCount(std::size_t files, int requested, int cap) {
    if (files == 0) {
        return 0;
    }
    int workers = requested > 0 ? requested : static_cast<int>((files + 2) / 3);
    if (requested <= 0 && workers > cap) {
        workers = cap;
    }
    if (workers < 1) {
        workers = 1;
    }
    if (workers > static_cast<int>(files)) {
        workers = static_cast<int>(files);
    }
    return workers;
}

BatchStats RunBatch(const Options& options, const std::vector<fs::path>& files,
                    const std::wstring& password, bool encrypt) {
    BatchStats stats;
    std::vector<std::pair<std::wstring, std::wstring>> failures;
    std::mutex outMutex;

    std::vector<fs::path> excelFiles;
    std::vector<fs::path> wordFiles;
    std::vector<fs::path> pdfFiles;
    for (const fs::path& file : files) {
        switch (util::TypeOf(file)) {
            case util::FileType::Excel: excelFiles.push_back(file); break;
            case util::FileType::Word:  wordFiles.push_back(file);  break;
            case util::FileType::Pdf:   pdfFiles.push_back(file);   break;
            default: break;
        }
    }

    const int total = static_cast<int>(files.size());
    const int digits = static_cast<int>(std::to_wstring(total).size());
    const auto started = std::chrono::steady_clock::now();
    std::atomic<int> completed{0};

    util::WriteLine();

    enum class Outcome { Done, Skipped, Failed };

    auto processOne = [&](const fs::path& file, util::FileType type,
                          ExcelApp& excel, WordApp& word, PdfTool& pdf,
                          const fs::path& tempDir) {
        const std::wstring label = util::RelativeForDisplay(file, options.folder);

        auto report = [&](Outcome outcome, const std::wstring& text, Color color,
                          const std::wstring& detail = std::wstring()) {
            std::lock_guard<std::mutex> lock(outMutex);
            const int index = ++completed;
            std::wstring counter = std::to_wstring(index);
            if (static_cast<int>(counter.size()) < digits) {
                counter.insert(0, static_cast<std::size_t>(digits) - counter.size(), L' ');
            }
            util::Write(L"   [" + counter + L"/" + std::to_wstring(total) + L"] " +
                        util::Fit(label, 44) + L"   ");
            util::WriteLine(Paint(text, color));
            if (!detail.empty()) {
                util::WriteLine(L"         " + Paint(detail, Color::Red));
            }
            switch (outcome) {
                case Outcome::Done:    ++stats.done;    break;
                case Outcome::Skipped: ++stats.skipped; break;
                case Outcome::Failed:
                    ++stats.failed;
                    failures.emplace_back(label, detail);
                    break;
            }
        };

        // For OOXML the container tells us the state without opening the file.
        if (util::IsOoxml(file)) {
            const util::Container container = util::DetectContainer(file);
            if (encrypt && container == util::Container::Cfb) {
                report(Outcome::Skipped, L"skipped (already encrypted)", Color::Yellow);
                return;
            }
            if (!encrypt && container == util::Container::Zip) {
                report(Outcome::Skipped, L"skipped (not encrypted)", Color::Yellow);
                return;
            }
        }

        bool createdBackup = false;
        if (options.backups != BackupMode::None) {
            std::wstring backupError;
            if (!util::EnsureBackup(file, createdBackup, backupError)) {
                report(Outcome::Failed, L"FAILED", Color::Red, backupError);
                return;
            }
        }

        FileResult result;
        if (type == util::FileType::Excel) {
            result = excel.ProcessFile(file.wstring(), password, encrypt, tempDir);
        } else if (type == util::FileType::Word) {
            result = word.ProcessFile(file.wstring(), password, encrypt, tempDir);
        } else {
            result = pdf.ProcessFile(file.wstring(), password, encrypt);
        }

        switch (result.action) {
            case FileAction::Encrypted:
            case FileAction::Decrypted: {
                if (options.backups == BackupMode::Temp) {
                    fs::path backup = file;
                    backup += L".bak";
                    util::RemoveQuiet(backup);
                }
                report(Outcome::Done,
                       result.action == FileAction::Encrypted ? L"encrypted" : L"decrypted",
                       Color::Green);
                break;
            }
            case FileAction::SkippedAlreadyEncrypted:
            case FileAction::SkippedNotEncrypted: {
                if (createdBackup) {
                    fs::path backup = file;
                    backup += L".bak";
                    util::RemoveQuiet(backup);
                }
                report(Outcome::Skipped,
                       result.action == FileAction::SkippedAlreadyEncrypted
                           ? L"skipped (already encrypted)"
                           : L"skipped (not encrypted)",
                       Color::Yellow);
                break;
            }
            case FileAction::Failed:
                report(Outcome::Failed, L"FAILED", Color::Red, result.detail);
                break;
        }
    };

    std::vector<std::thread> threads;
    std::atomic<int> workerIndex{0};

    auto spawn = [&](std::vector<fs::path>& group, util::FileType type, int cap) {
        const int workers = WorkerCount(group.size(), options.jobs, cap);
        if (workers == 0) {
            return;
        }
        auto next = std::make_shared<std::atomic<std::size_t>>(0);
        for (int i = 0; i < workers; ++i) {
            threads.emplace_back([&, type, next]() {
                CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
                const fs::path tempDir = MakeWorkerTempDir(workerIndex.fetch_add(1));
                ExcelApp excel;
                WordApp word;
                PdfTool pdf;
                for (;;) {
                    const std::size_t index = next->fetch_add(1);
                    if (index >= group.size()) {
                        break;
                    }
                    processOne(group[index], type, excel, word, pdf, tempDir);
                }
                excel.Stop();
                word.Stop();
                std::error_code ec;
                fs::remove_all(tempDir, ec);
                CoUninitialize();
            });
        }
    };

    // Excel, Word and PDF run at the same time; each type also gets several
    // workers because Office spends most of its time encrypting, which is CPU
    // bound and scales well across instances.
    spawn(excelFiles, util::FileType::Excel, 3);
    spawn(wordFiles, util::FileType::Word, 2);
    spawn(pdfFiles, util::FileType::Pdf, 1);

    for (std::thread& thread : threads) {
        thread.join();
    }

    const double seconds =
        std::chrono::duration<double>(std::chrono::steady_clock::now() - started).count();
    wchar_t elapsed[64] = {};
    swprintf(elapsed, 64, L"%.1f s", seconds);

    util::WriteLine();
    PrintSeparator();
    std::wstring summary = L"   ";
    summary += encrypt ? L"Encrypted: " : L"Decrypted: ";
    summary += std::to_wstring(stats.done);
    summary += L"   Skipped: " + std::to_wstring(stats.skipped);
    summary += L"   Failed: " + std::to_wstring(stats.failed);
    summary += L"   Time: " + std::wstring(elapsed);
    util::WriteLine(summary);
    PrintSeparator();

    if (!failures.empty()) {
        util::WriteLine();
        util::WriteLine(Paint(L"   Files that failed:", Color::Red));
        for (const auto& failure : failures) {
            util::WriteLine(L"     - " + failure.first + L": " + failure.second);
        }
    }
    return stats;
}

void DoEncrypt(const Options& options, const util::ScanResult& scan) {
    if (scan.files.empty()) {
        util::WriteLine(L"\n   No matching files found.\n");
        return;
    }
    util::WriteLine();
    util::WriteLine(L"   " + std::to_wstring(scan.files.size()) +
                    L" file(s) will be encrypted in place.");
    if (options.backups == BackupMode::Keep) {
        util::WriteLine(Paint(L"   A .bak copy is kept the first time; it stays unencrypted.",
                              Color::Yellow));
    } else if (options.backups == BackupMode::Temp) {
        util::WriteLine(Paint(L"   A .bak copy is made first and deleted again when the file succeeds.",
                              Color::Yellow));
    }

    const std::wstring password = AskNewPassword();
    if (password.empty()) {
        return;
    }
    if (!util::Confirm(L"\n   Encrypt now? Type YES to continue: ", false)) {
        util::WriteLine(L"   Cancelled.");
        return;
    }
    RunBatch(options, scan.files, password, true);
}

void DoDecrypt(const Options& options, const util::ScanResult& scan) {
    if (scan.files.empty()) {
        util::WriteLine(L"\n   No matching files found.\n");
        return;
    }
    util::WriteLine();
    util::WriteLine(L"   " + std::to_wstring(scan.files.size()) +
                    L" file(s) will be decrypted in place.");
    if (options.backups == BackupMode::Keep) {
        util::WriteLine(Paint(L"   A .bak copy is kept the first time.", Color::Yellow));
    } else if (options.backups == BackupMode::Temp) {
        util::WriteLine(Paint(L"   A .bak copy is made first and deleted again when the file succeeds.",
                              Color::Yellow));
    }

    const std::wstring password = AskExistingPassword();
    if (password.empty()) {
        return;
    }
    if (!util::Confirm(L"\n   Decrypt now? Type YES to continue: ", false)) {
        util::WriteLine(L"   Cancelled.");
        return;
    }
    RunBatch(options, scan.files, password, false);
}

// Native Windows folder picker (IFileOpenDialog with FOS_PICKFOLDERS).
std::wstring PickFolder(const fs::path& initial) {
    std::wstring result;

    IFileOpenDialog* dialog = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
                                  IID_PPV_ARGS(&dialog));
    if (FAILED(hr) || dialog == nullptr) {
        return result;
    }

    DWORD flags = 0;
    dialog->GetOptions(&flags);
    dialog->SetOptions(flags | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
    dialog->SetTitle(L"Select the folder with the files to protect");

    if (!initial.empty()) {
        IShellItem* start = nullptr;
        if (SUCCEEDED(SHCreateItemFromParsingName(initial.c_str(), nullptr,
                                                  IID_PPV_ARGS(&start)))) {
            dialog->SetFolder(start);
            start->Release();
        }
    }

    if (SUCCEEDED(dialog->Show(GetConsoleWindow()))) {
        IShellItem* item = nullptr;
        if (SUCCEEDED(dialog->GetResult(&item))) {
            PWSTR path = nullptr;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &path))) {
                result = path;
                CoTaskMemFree(path);
            }
            item->Release();
        }
    }

    dialog->Release();
    return result;
}

void ChangeFolder(Options& options) {
    util::WriteLine();
    util::WriteLine(L"   Current folder: " + options.folder.wstring());
    std::wstring input;
    if (!util::ReadLine(L"   New folder (empty = cancel): ", input)) {
        return;
    }
    input = util::Trim(input);
    if (input.empty()) {
        return;
    }
    if (input.size() >= 2 && input.front() == L'"' && input.back() == L'"') {
        input = input.substr(1, input.size() - 2);
    }

    std::error_code ec;
    const fs::path folder = fs::absolute(fs::path(input), ec);
    if (ec || !fs::is_directory(folder, ec)) {
        util::WriteLine(Paint(L"   Not a folder: " + input, Color::Red));
        util::Pause();
        return;
    }
    options.folder = folder;
    util::WriteLine(Paint(L"   Folder set to " + folder.wstring(), Color::Green));
}

void ChangeTypes(Options& options) {
    util::WriteLine();
    util::WriteLine(L"   Current: " + TypeSummary(options.types));
    util::WriteLine(L"   e = Excel, w = Word, p = PDF, a = all (combine, e.g. \"ew\")");
    std::wstring input;
    if (!util::ReadLine(L"   Types: ", input)) {
        return;
    }
    input = util::ToLower(util::Trim(input));
    if (input.empty()) {
        return;
    }
    if (input == L"a" || input == L"all") {
        options.types = util::FileTypes{};
        return;
    }
    util::FileTypes types{false, false, false};
    for (wchar_t c : input) {
        if (c == L'e') types.excel = true;
        else if (c == L'w') types.word = true;
        else if (c == L'p') types.pdf = true;
    }
    if (!types.Any()) {
        util::WriteLine(Paint(L"   Nothing selected.", Color::Red));
        util::Pause();
        return;
    }
    options.types = types;
}

int RunInteractive(Options options) {
    PrintBanner();
    for (;;) {
        const util::ScanResult scan =
            util::ScanFiles(options.folder, options.recursive, options.types);
        PrintStatus(options, scan);
        PrintMenu(options);

        std::wstring choice;
        if (!util::ReadLine(L"   Select: ", choice)) {
            util::WriteLine();
            return 0;
        }
        choice = util::ToLower(util::Trim(choice));

        if (choice == L"1") {
            DoEncrypt(options, scan);
            util::Pause();
        } else if (choice == L"2") {
            DoDecrypt(options, scan);
            util::Pause();
        } else if (choice == L"3") {
            ChangeFolder(options);
        } else if (choice == L"4") {
            options.recursive = !options.recursive;
        } else if (choice == L"5") {
            options.backups = NextBackupMode(options.backups);
        } else if (choice == L"6") {
            ListFiles(options, scan);
            util::Pause();
        } else if (choice == L"7") {
            ChangeTypes(options);
        } else if (choice == L"8") {
            const std::wstring picked = PickFolder(options.folder);
            if (!picked.empty()) {
                options.folder = fs::path(picked);
                util::WriteLine(Paint(L"   Folder set to " + options.folder.wstring(),
                                      Color::Green));
            }
        } else if (choice == L"0" || choice == L"q") {
            return 0;
        } else if (!choice.empty()) {
            util::WriteLine(Paint(L"   Unknown option.", Color::Red));
        }
    }
}

int RunOnce(const Options& options, bool encrypt, std::wstring password, bool assumeYes) {
    std::error_code ec;
    if (!fs::is_directory(options.folder, ec)) {
        util::WriteLine(Paint(L"   Folder not found: " + options.folder.wstring(), Color::Red));
        return 2;
    }

    const util::ScanResult scan =
        util::ScanFiles(options.folder, options.recursive, options.types);
    for (const std::wstring& error : scan.errors) {
        util::WriteLine(Paint(L"   ! " + error, Color::Red));
    }
    util::WriteLine(L"   Folder: " + options.folder.wstring());
    if (scan.files.empty()) {
        util::WriteLine(L"   No matching files found.");
        return 0;
    }

    if (password.empty()) {
        if (!util::ReadPassword(L"   Password: ", password) || password.empty()) {
            util::WriteLine(Paint(L"   A password is required.", Color::Red));
            return 2;
        }
    }
    if (password.size() > 255) {
        util::WriteLine(Paint(L"   Passwords are limited to 255 characters.", Color::Red));
        return 2;
    }

    if (!assumeYes) {
        const std::wstring question =
            L"   " + std::wstring(encrypt ? L"Encrypt " : L"Decrypt ") +
            std::to_wstring(scan.files.size()) + L" file(s) in place? Type YES: ";
        if (!util::Confirm(question, false)) {
            util::WriteLine(L"   Cancelled.");
            return 1;
        }
    }

    const BatchStats stats = RunBatch(options, scan.files, password, encrypt);
    return stats.failed == 0 ? 0 : 1;
}

void PrintHelp() {
    util::WriteLine();
    util::WriteLine(Paint(L"   Excel Password Tool 1.3", Color::Bold));
    util::WriteLine(L"   Encrypts Excel, Word and PDF files in a folder with one password, or");
    util::WriteLine(L"   removes the password again. Excel and Word are driven through COM,");
    util::WriteLine(L"   so those files get their real \"Encrypt with password\" protection.");
    util::WriteLine(L"   PDFs use the standard PDF AES-256 security handler.");
    util::WriteLine();
    util::WriteLine(L"   Usage: ExcelPasswordTool.exe [folder] [options]");
    util::WriteLine();
    util::WriteLine(L"     -d, --folder <path>    Folder to process (default: current folder)");
    util::WriteLine(L"     -e, --encrypt          Encrypt all matching files (no menu)");
    util::WriteLine(L"     -x, --decrypt          Decrypt all matching files (no menu)");
    util::WriteLine(L"     -p, --password <text>  Password (otherwise asked, hidden)");
    util::WriteLine(L"     -r, --recursive        Include subfolders");
    util::WriteLine(L"         --types <list>     e=Excel w=Word p=PDF, e.g. \"ewp\" or \"all\"");
    util::WriteLine(L"         --no-backup        Do not keep .bak copies (default: keep)");
    util::WriteLine(L"         --temp-backup      Make a .bak first, delete it when the file succeeds");
    util::WriteLine(L"         --jobs <n>         Parallel workers per file type (1-8, default auto)");
    util::WriteLine(L"     -y, --yes              Skip the confirmation question");
    util::WriteLine(L"     -h, --help             Show this help");
    util::WriteLine();
    util::WriteLine(L"   Excel: .xlsx .xlsm .xlsb .xls .xltx .xltm .xlt .xlam");
    util::WriteLine(L"   Word : .docx .docm .dotx .dotm .doc .dot");
    util::WriteLine(L"   PDF  : .pdf");
    util::WriteLine();
    util::WriteLine(L"   Examples:");
    util::WriteLine(L"     ExcelPasswordTool.exe");
    util::WriteLine(L"     ExcelPasswordTool.exe \"D:\\Reports\"");
    util::WriteLine(L"     ExcelPasswordTool.exe -e -d \"D:\\Reports\" -p secret -y");
    util::WriteLine(L"     ExcelPasswordTool.exe -x --types ewp -d \"D:\\Reports\" -p secret");
    util::WriteLine();
    util::WriteLine(L"   Exit codes: 0 ok, 1 some files failed, 2 usage or password error.");
    util::WriteLine();
}

CommandLineOptions ParseCommandLine() {
    CommandLineOptions result;

    std::error_code ec;
    result.options.folder = fs::current_path(ec);
    if (ec) {
        result.options.folder = L".";
    }

    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv == nullptr) {
        return result;
    }

    bool wantEncrypt = false;
    bool wantDecrypt = false;

    auto needValue = [&](int& index, const std::wstring& option) {
        if (index + 1 >= argc) {
            result.bad = true;
            result.error = L"Missing value after " + option;
            return false;
        }
        ++index;
        return true;
    };

    for (int index = 1; index < argc; ++index) {
        const std::wstring argument = argv[index];
        if (argument == L"-h" || argument == L"--help") {
            result.showHelp = true;
        } else if (argument == L"-d" || argument == L"--folder") {
            if (needValue(index, argument)) {
                result.options.folder = argv[index];
            }
        } else if (argument == L"-e" || argument == L"--encrypt") {
            wantEncrypt = true;
        } else if (argument == L"-x" || argument == L"--decrypt") {
            wantDecrypt = true;
        } else if (argument == L"-p" || argument == L"--password") {
            if (needValue(index, argument)) {
                result.password = argv[index];
            }
        } else if (argument == L"-r" || argument == L"--recursive") {
            result.options.recursive = true;
        } else if (argument == L"--no-backup") {
            result.options.backups = BackupMode::None;
        } else if (argument == L"--temp-backup") {
            result.options.backups = BackupMode::Temp;
        } else if (argument == L"--jobs") {
            if (needValue(index, argument)) {
                wchar_t* end = nullptr;
                const long value = wcstol(argv[index], &end, 10);
                if (end == argv[index] || value < 1 || value > 8) {
                    result.bad = true;
                    result.error = L"--jobs needs a number from 1 to 8";
                } else {
                    result.options.jobs = static_cast<int>(value);
                }
            }
        } else if (argument == L"--types") {
            if (needValue(index, argument)) {
                const std::wstring value = util::ToLower(argv[index]);
                if (value == L"all") {
                    result.options.types = util::FileTypes{};
                } else {
                    util::FileTypes types{false, false, false};
                    for (wchar_t c : value) {
                        if (c == L'e') types.excel = true;
                        else if (c == L'w') types.word = true;
                        else if (c == L'p') types.pdf = true;
                    }
                    if (!types.Any()) {
                        result.bad = true;
                        result.error = L"--types needs a combination of e, w, p or \"all\"";
                    } else {
                        result.options.types = types;
                    }
                }
            }
        } else if (argument == L"-y" || argument == L"--yes") {
            result.assumeYes = true;
        } else if (!argument.empty() && argument.front() != L'-') {
            result.options.folder = argument;
        } else {
            result.bad = true;
            result.error = L"Unknown option: " + argument;
        }
    }
    LocalFree(argv);

    if (wantEncrypt && wantDecrypt) {
        result.bad = true;
        result.error = L"Use either --encrypt or --decrypt, not both.";
    } else if (wantEncrypt) {
        result.oneShot = OneShot::Encrypt;
    } else if (wantDecrypt) {
        result.oneShot = OneShot::Decrypt;
    }

    if (!result.bad) {
        std::error_code absoluteError;
        const fs::path absolute = fs::absolute(result.options.folder, absoluteError);
        if (!absoluteError) {
            result.options.folder = absolute;
        }
    }
    return result;
}

}  // namespace

int main() {
    util::InitConsole();

    const HRESULT comResult = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(comResult)) {
        util::WriteLine(Paint(L"   ERROR: COM could not be initialized.", Color::Red));
        return 2;
    }

    int exitCode = 0;
    {
        CommandLineOptions parsed = ParseCommandLine();
        if (parsed.showHelp) {
            PrintHelp();
        } else if (parsed.bad) {
            util::WriteLine(Paint(L"   ERROR: " + parsed.error, Color::Red));
            util::WriteLine();
            PrintHelp();
            exitCode = 2;
        } else if (parsed.oneShot == OneShot::None) {
            exitCode = RunInteractive(parsed.options);
        } else {
            exitCode = RunOnce(parsed.options, parsed.oneShot == OneShot::Encrypt,
                               parsed.password, parsed.assumeYes);
        }
    }

    CoUninitialize();
    return exitCode;
}
