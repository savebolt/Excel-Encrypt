#include "util.h"

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>

#include <algorithm>
#include <conio.h>
#include <cwchar>
#include <cwctype>
#include <io.h>
#include <iostream>

namespace util {
namespace {

bool g_colors = false;

const wchar_t* CodeFor(Color color) {
    switch (color) {
        case Color::Green:  return L"\x1b[32m";
        case Color::Red:    return L"\x1b[31m";
        case Color::Yellow: return L"\x1b[33m";
        case Color::Cyan:   return L"\x1b[36m";
        case Color::Bold:   return L"\x1b[1m";
        case Color::Dim:    return L"\x1b[2m";
        case Color::Default:
        default:            return L"";
    }
}

}  // namespace

std::wstring ToLower(std::wstring text) {
    std::transform(text.begin(), text.end(), text.begin(),
                   [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    return text;
}

std::string ToUtf8(const std::wstring& text) {
    if (text.empty()) {
        return {};
    }
    const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                                         nullptr, 0, nullptr, nullptr);
    if (size <= 0) {
        return {};
    }
    std::string result(static_cast<std::size_t>(size), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                        result.data(), size, nullptr, nullptr);
    return result;
}

std::wstring FromUtf8(const std::string& text) {
    if (text.empty()) {
        return {};
    }
    const int size = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                                         nullptr, 0);
    if (size <= 0) {
        return {};
    }
    std::wstring result(static_cast<std::size_t>(size), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                        result.data(), size);
    return result;
}

std::wstring Trim(const std::wstring& text) {
    const std::size_t begin = text.find_first_not_of(L" \t\r\n");
    if (begin == std::wstring::npos) {
        return {};
    }
    const std::size_t end = text.find_last_not_of(L" \t\r\n");
    return text.substr(begin, end - begin + 1);
}

std::wstring Fit(const std::wstring& text, std::size_t width) {
    if (text.size() <= width) {
        return text;
    }
    if (width <= 3) {
        return text.substr(0, width);
    }
    const std::size_t left = (width - 3) / 2;
    const std::size_t right = width - 3 - left;
    return text.substr(0, left) + L"..." + text.substr(text.size() - right);
}

std::wstring FormatSize(std::uintmax_t bytes) {
    wchar_t buffer[64] = {};
    if (bytes < 1024) {
        swprintf(buffer, 64, L"%llu B", static_cast<unsigned long long>(bytes));
    } else if (bytes < 1024 * 1024) {
        swprintf(buffer, 64, L"%.1f KB", static_cast<double>(bytes) / 1024.0);
    } else if (bytes < 1024ull * 1024 * 1024) {
        swprintf(buffer, 64, L"%.1f MB", static_cast<double>(bytes) / (1024.0 * 1024.0));
    } else {
        swprintf(buffer, 64, L"%.2f GB", static_cast<double>(bytes) / (1024.0 * 1024.0 * 1024.0));
    }
    return buffer;
}

std::wstring Win32Text(unsigned long code) {
    wchar_t* buffer = nullptr;
    const DWORD length = FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
            FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr, static_cast<DWORD>(code), 0, reinterpret_cast<LPWSTR>(&buffer), 0, nullptr);
    std::wstring text = (length > 0 && buffer != nullptr)
                            ? std::wstring(buffer, length)
                            : L"error " + std::to_wstring(code);
    if (buffer != nullptr) {
        LocalFree(buffer);
    }
    while (!text.empty() &&
           (text.back() == L'\r' || text.back() == L'\n' || text.back() == L' ')) {
        text.pop_back();
    }
    return text;
}

void InitConsole() {
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);

    HANDLE output = GetStdHandle(STD_OUTPUT_HANDLE);
    DWORD mode = 0;
    if (output != INVALID_HANDLE_VALUE && GetConsoleMode(output, &mode)) {
        if (SetConsoleMode(output, mode | ENABLE_VIRTUAL_TERMINAL_PROCESSING)) {
            g_colors = true;
        }
    }
}

std::wstring Paint(const std::wstring& text, Color color) {
    if (!g_colors || color == Color::Default) {
        return text;
    }
    std::wstring result = CodeFor(color);
    result += text;
    result += L"\x1b[0m";
    return result;
}

void Write(const std::wstring& text) {
    const std::string utf8 = ToUtf8(text);
    std::cout.write(utf8.data(), static_cast<std::streamsize>(utf8.size()));
    std::cout.flush();
}

void WriteLine(const std::wstring& text) {
    Write(text + L"\n");
}

bool ReadLine(const std::wstring& prompt, std::wstring& out) {
    out.clear();
    Write(prompt);
    std::string line;
    if (!std::getline(std::cin, line)) {
        return false;
    }
    if (!line.empty() && line.back() == '\r') {
        line.pop_back();
    }
    out = FromUtf8(line);
    return true;
}

bool ReadPassword(const std::wstring& prompt, std::wstring& out) {
    out.clear();
    Write(prompt);

    if (_isatty(_fileno(stdin)) == 0) {
        std::string line;
        if (!std::getline(std::cin, line)) {
            return false;
        }
        if (!line.empty() && line.back() == '\r') {
            line.pop_back();
        }
        out = FromUtf8(line);
        return true;
    }

    for (;;) {
        const wint_t ch = _getwch();
        if (ch == WEOF) {
            WriteLine();
            return false;
        }
        if (ch == 3) {  // Ctrl+C
            WriteLine();
            return false;
        }
        if (ch == L'\r' || ch == L'\n') {
            WriteLine();
            return true;
        }
        if (ch == L'\b') {
            if (!out.empty()) {
                out.pop_back();
                Write(L"\b \b");
            }
            continue;
        }
        if (ch == 0 || ch == 0xE0) {  // function / arrow keys arrive as two codes
            _getwch();
            continue;
        }
        out.push_back(static_cast<wchar_t>(ch));
        Write(L"*");
    }
}

bool Confirm(const std::wstring& prompt, bool defaultYes) {
    std::wstring answer;
    if (!ReadLine(prompt, answer)) {
        return false;
    }
    answer = ToLower(Trim(answer));
    if (answer.empty()) {
        return defaultYes;
    }
    return answer == L"y" || answer == L"yes";
}

void Pause() {
    std::wstring ignored;
    ReadLine(L"\n  Press Enter to continue...", ignored);
}

bool IsExcelFile(const std::filesystem::path& file) {
    const std::wstring extension = ToLower(file.extension().wstring());
    return extension == L".xlsx" || extension == L".xlsm" || extension == L".xlsb" ||
           extension == L".xls" || extension == L".xltx" || extension == L".xltm" ||
           extension == L".xlt" || extension == L".xlam";
}

bool IsWordFile(const std::filesystem::path& file) {
    const std::wstring extension = ToLower(file.extension().wstring());
    return extension == L".docx" || extension == L".docm" || extension == L".dotx" ||
           extension == L".dotm" || extension == L".doc" || extension == L".dot";
}

bool IsPdfFile(const std::filesystem::path& file) {
    return ToLower(file.extension().wstring()) == L".pdf";
}

bool IsOoxml(const std::filesystem::path& file) {
    const std::wstring extension = ToLower(file.extension().wstring());
    return extension == L".xlsx" || extension == L".xlsm" || extension == L".xlsb" ||
           extension == L".xltx" || extension == L".xltm" || extension == L".xlam" ||
           extension == L".docx" || extension == L".docm" ||
           extension == L".dotx" || extension == L".dotm";
}

FileType TypeOf(const std::filesystem::path& file) {
    if (IsExcelFile(file)) return FileType::Excel;
    if (IsWordFile(file)) return FileType::Word;
    if (IsPdfFile(file)) return FileType::Pdf;
    return FileType::Other;
}

ScanResult ScanFiles(const std::filesystem::path& folder, bool recursive,
                     const FileTypes& types) {
    namespace fs = std::filesystem;
    ScanResult result;

    std::error_code ec;
    if (!fs::is_directory(folder, ec)) {
        result.errors.push_back(L"Folder not found: " + folder.wstring());
        return result;
    }

    auto add = [&result, &types](const fs::directory_entry& entry) {
        std::error_code entryError;
        if (!entry.is_regular_file(entryError)) {
            return;
        }
        const std::wstring name = entry.path().filename().wstring();
        if (name.size() >= 2 && name[0] == L'~' && name[1] == L'$') {
            return;  // Excel/Word lock file
        }
        const FileType type = TypeOf(entry.path());
        const bool wanted = (type == FileType::Excel && types.excel) ||
                            (type == FileType::Word && types.word) ||
                            (type == FileType::Pdf && types.pdf);
        if (!wanted) {
            return;
        }
        result.files.push_back(entry.path());
    };

    if (recursive) {
        fs::recursive_directory_iterator it(folder, fs::directory_options::skip_permission_denied, ec);
        fs::recursive_directory_iterator end;
        while (it != end) {
            add(*it);
            it.increment(ec);
            if (ec) {
                result.errors.push_back(L"Stopped reading subfolders: " + FromUtf8(ec.message()));
                break;
            }
        }
    } else {
        fs::directory_iterator it(folder, fs::directory_options::skip_permission_denied, ec);
        fs::directory_iterator end;
        while (it != end) {
            add(*it);
            it.increment(ec);
            if (ec) {
                result.errors.push_back(L"Stopped reading folder: " + FromUtf8(ec.message()));
                break;
            }
        }
    }

    std::sort(result.files.begin(), result.files.end(),
              [](const fs::path& left, const fs::path& right) {
                  return ToLower(left.filename().wstring()) <
                         ToLower(right.filename().wstring());
              });
    return result;
}

ScanResult ScanExcelFiles(const std::filesystem::path& folder, bool recursive) {
    FileTypes types;
    types.excel = true;
    types.word = false;
    types.pdf = false;
    return ScanFiles(folder, recursive, types);
}

Container DetectContainer(const std::filesystem::path& file) {
    HANDLE handle = CreateFileW(file.c_str(), GENERIC_READ,
                                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (handle == INVALID_HANDLE_VALUE) {
        return Container::Unknown;
    }

    unsigned char signature[8] = {};
    DWORD read = 0;
    const BOOL ok = ReadFile(handle, signature, sizeof(signature), &read, nullptr);
    CloseHandle(handle);
    if (ok == FALSE || read < 4) {
        return Container::Unknown;
    }

    if (signature[0] == 0xD0 && signature[1] == 0xCF &&
        signature[2] == 0x11 && signature[3] == 0xE0 &&
        signature[4] == 0xA1 && signature[5] == 0xB1 &&
        signature[6] == 0x1A && signature[7] == 0xE1) {
        return Container::Cfb;
    }
    if (signature[0] == 'P' && signature[1] == 'K' &&
        signature[2] == 0x03 && signature[3] == 0x04) {
        return Container::Zip;
    }
    return Container::Unknown;
}

std::wstring RelativeForDisplay(const std::filesystem::path& file,
                                const std::filesystem::path& root) {
    std::error_code ec;
    const std::filesystem::path relative = std::filesystem::relative(file, root, ec);
    if (ec || relative.empty()) {
        return file.filename().wstring();
    }
    const std::wstring text = relative.wstring();
    if (text.rfind(L"..", 0) == 0) {
        return file.wstring();
    }
    return text;
}

bool EnsureBackup(const std::filesystem::path& file, bool& created, std::wstring& error) {
    namespace fs = std::filesystem;
    created = false;
    error.clear();

    fs::path backup = file;
    backup += L".bak";

    std::error_code ec;
    if (fs::exists(backup, ec)) {
        return true;  // keep the oldest copy
    }

    if (!fs::copy_file(file, backup, fs::copy_options::none, ec)) {
        error = L"could not create backup: " + FromUtf8(ec.message());
        return false;
    }
    created = true;
    return true;
}

void RemoveQuiet(const std::filesystem::path& file) {
    std::error_code ec;
    std::filesystem::remove(file, ec);
}

}  // namespace util
