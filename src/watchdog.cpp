#include "watchdog.h"

Watchdog::~Watchdog() {
    Stop();
}

void Watchdog::Start(DWORD processId) {
    Stop();
    {
        std::lock_guard<std::mutex> lock(mutex_);
        processId_ = processId;
        stop_ = false;
        armed_ = false;
        triggered_ = false;
    }
    thread_ = std::thread(&Watchdog::Loop, this);
}

void Watchdog::Stop() {
    {
        std::lock_guard<std::mutex> lock(mutex_);
        stop_ = true;
        armed_ = false;
    }
    wake_.notify_all();
    if (thread_.joinable()) {
        thread_.join();
    }
    std::lock_guard<std::mutex> lock(mutex_);
    processId_ = 0;
}

void Watchdog::Arm(std::chrono::milliseconds timeout) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!thread_.joinable()) {
        return;
    }
    triggered_ = false;
    armed_ = true;
    deadline_ = std::chrono::steady_clock::now() + timeout;
    wake_.notify_all();
}

void Watchdog::Disarm() {
    std::lock_guard<std::mutex> lock(mutex_);
    armed_ = false;
    wake_.notify_all();
}

bool Watchdog::Triggered() {
    std::lock_guard<std::mutex> lock(mutex_);
    return triggered_;
}

void Watchdog::Loop() {
    std::unique_lock<std::mutex> lock(mutex_);
    for (;;) {
        if (stop_) {
            return;
        }
        if (!armed_) {
            wake_.wait(lock, [this] { return stop_ || armed_; });
            continue;
        }
        if (wake_.wait_until(lock, deadline_, [this] { return stop_ || !armed_; })) {
            continue;
        }

        // Deadline reached while still armed: the guarded call is stuck.
        armed_ = false;
        triggered_ = true;
        const DWORD processId = processId_;
        lock.unlock();
        if (processId != 0) {
            HANDLE process = OpenProcess(PROCESS_TERMINATE | SYNCHRONIZE, FALSE, processId);
            if (process != nullptr) {
                TerminateProcess(process, 1);
                CloseHandle(process);
            }
        }
        lock.lock();
    }
}
