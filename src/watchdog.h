#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>

#include <chrono>
#include <condition_variable>
#include <mutex>
#include <thread>

// Kills a worker process if a guarded operation takes too long. Excel can show
// a modal dialog that blocks COM forever (for example when a file is locked);
// this lets the tool give up instead of hanging.
class Watchdog {
public:
    Watchdog() = default;
    ~Watchdog();

    Watchdog(const Watchdog&) = delete;
    Watchdog& operator=(const Watchdog&) = delete;

    void Start(DWORD processId);
    void Arm(std::chrono::milliseconds timeout);
    void Disarm();
    bool Triggered();
    void Stop();

private:
    void Loop();

    std::mutex mutex_;
    std::condition_variable wake_;
    std::thread thread_;
    DWORD processId_ = 0;
    bool stop_ = false;
    bool armed_ = false;
    bool triggered_ = false;
    std::chrono::steady_clock::time_point deadline_{};
};
