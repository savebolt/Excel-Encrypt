#pragma once

#include <cstdint>
#include <filesystem>
#include <string>
#include <vector>

namespace util {

enum class Container { Unknown, Zip, Cfb };

enum class Color { Default, Green, Red, Yellow, Cyan, Bold, Dim };

std::wstring ToLower(std::wstring text);
std::string ToUtf8(const std::wstring& text);
std::wstring FromUtf8(const std::string& text);
std::wstring Trim(const std::wstring& text);
std::wstring Fit(const std::wstring& text, std::size_t width);
std::wstring FormatSize(std::uintmax_t bytes);
std::wstring Win32Text(unsigned long code);

void InitConsole();
std::wstring Paint(const std::wstring& text, Color color);
void Write(const std::wstring& text);
void WriteLine(const std::wstring& text = std::wstring());
bool ReadLine(const std::wstring& prompt, std::wstring& out);
bool ReadPassword(const std::wstring& prompt, std::wstring& out);
bool Confirm(const std::wstring& prompt, bool defaultYes);
void Pause();

struct ScanResult {
    std::vector<std::filesystem::path> files;
    std::vector<std::wstring> errors;
};

enum class FileType { Excel, Word, Pdf, Other };

struct FileTypes {
    bool excel = true;
    bool word = true;
    bool pdf = true;
    bool Any() const { return excel || word || pdf; }
};

bool IsExcelFile(const std::filesystem::path& file);
bool IsWordFile(const std::filesystem::path& file);
bool IsPdfFile(const std::filesystem::path& file);
bool IsOoxml(const std::filesystem::path& file);
FileType TypeOf(const std::filesystem::path& file);
ScanResult ScanFiles(const std::filesystem::path& folder, bool recursive,
                     const FileTypes& types);
ScanResult ScanExcelFiles(const std::filesystem::path& folder, bool recursive);
Container DetectContainer(const std::filesystem::path& file);
std::wstring RelativeForDisplay(const std::filesystem::path& file,
                                const std::filesystem::path& root);

bool EnsureBackup(const std::filesystem::path& file, bool& created, std::wstring& error);
void RemoveQuiet(const std::filesystem::path& file);

}  // namespace util
