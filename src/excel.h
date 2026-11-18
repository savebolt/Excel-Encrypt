#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>

#include <filesystem>
#include <string>

#include "com.h"
#include "result.h"
#include "watchdog.h"

// Drives Microsoft Excel through COM so the password protection written to the
// file is the real "Encrypt with password" feature of Excel itself.
class ExcelApp {
public:
    ExcelApp() = default;
    ~ExcelApp();

    ExcelApp(const ExcelApp&) = delete;
    ExcelApp& operator=(const ExcelApp&) = delete;

    void Start();
    void Stop();
    bool Running() const { return static_cast<bool>(app_); }

    FileResult ProcessFile(const std::wstring& path, const std::wstring& password, bool encrypt,
                           const std::filesystem::path& tempDir = {});

private:
    void Configure();
    void StopInternal(bool quitExcel);

    com::ComPtr app_;
    Watchdog watchdog_;
    DWORD processId_ = 0;
};
