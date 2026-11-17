#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>

#include <string>

#include "com.h"
#include "result.h"
#include "watchdog.h"

// Drives Microsoft Word through COM so the password written to the file is
// Word's own "Encrypt with password" (File > Info > Protect Document).
class WordApp {
public:
    WordApp() = default;
    ~WordApp();

    WordApp(const WordApp&) = delete;
    WordApp& operator=(const WordApp&) = delete;

    void Start();
    void Stop();
    bool Running() const { return static_cast<bool>(app_); }

    FileResult ProcessFile(const std::wstring& path, const std::wstring& password, bool encrypt);

private:
    void Configure();
    void StopInternal(bool quitWord);

    com::ComPtr app_;
    Watchdog watchdog_;
    DWORD processId_ = 0;
};
