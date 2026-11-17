#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>
#include <ole2.h>
#include <oleauto.h>

#include <cstddef>
#include <exception>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace com {

std::string ToUtf8(const std::wstring& text);
std::wstring HResultText(HRESULT hr);

class Error : public std::exception {
public:
    Error(HRESULT code, const std::wstring& message);
    const char* what() const noexcept override;
    HRESULT Code() const noexcept { return code_; }
    const std::wstring& Message() const noexcept { return message_; }

private:
    HRESULT code_;
    std::wstring message_;
    std::string text_;
};

class ComPtr {
public:
    ComPtr() = default;
    explicit ComPtr(IDispatch* object) : object_(object) {}
    ~ComPtr();

    ComPtr(const ComPtr&) = delete;
    ComPtr& operator=(const ComPtr&) = delete;

    ComPtr(ComPtr&& other) noexcept;
    ComPtr& operator=(ComPtr&& other) noexcept;

    void Reset();
    IDispatch* Get() const { return object_; }
    IDispatch* operator->() const { return object_; }
    explicit operator bool() const { return object_ != nullptr; }
    IDispatch** Put();
    void Attach(IDispatch* object);

private:
    IDispatch* object_ = nullptr;
};

// Holds the arguments of one late-bound call. Positions are given in the
// natural (left to right) order of the COM method; the class reverses them
// internally as IDispatch::Invoke requires.
class ArgList {
public:
    explicit ArgList(std::size_t count);
    ~ArgList();

    ArgList(const ArgList&) = delete;
    ArgList& operator=(const ArgList&) = delete;

    UINT Count() const { return static_cast<UINT>(values_.size()); }
    VARIANT* Data() { return values_.data(); }

    void SetMissing(std::size_t position);
    void SetBool(std::size_t position, bool value);
    void SetLong(std::size_t position, long value);
    void SetString(std::size_t position, const std::wstring& value);

private:
    VARIANT* Slot(std::size_t position);
    std::vector<VARIANT> values_;
};

HRESULT InvokeRaw(IDispatch* object, const wchar_t* name, WORD flags,
                  ArgList* args, VARIANT* result, std::wstring* errorText);

std::wstring GetString(IDispatch* object, const wchar_t* name);
long GetLong(IDispatch* object, const wchar_t* name);
bool GetBool(IDispatch* object, const wchar_t* name);
ComPtr GetObject(IDispatch* object, const wchar_t* name);

void PutBool(IDispatch* object, const wchar_t* name, bool value);
void PutLong(IDispatch* object, const wchar_t* name, long value);
void PutString(IDispatch* object, const wchar_t* name, const std::wstring& value);

ComPtr CallObject(IDispatch* object, const wchar_t* name, ArgList& args);
void CallVoid(IDispatch* object, const wchar_t* name, ArgList& args);

}  // namespace com
