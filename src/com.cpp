#include "com.h"

#include <cstring>

namespace com {
namespace {

std::wstring Describe(const wchar_t* name, const std::wstring& detail) {
    std::wstring text(name);
    text += L" failed";
    if (!detail.empty()) {
        text += L": ";
        text += detail;
    }
    return text;
}

VARIANT InvokeGet(IDispatch* object, const wchar_t* name) {
    VARIANT value;
    VariantInit(&value);
    std::wstring errorText;
    const HRESULT hr = InvokeRaw(object, name, DISPATCH_PROPERTYGET, nullptr, &value, &errorText);
    if (FAILED(hr)) {
        VariantClear(&value);
        throw Error(hr, Describe(name, errorText.empty() ? HResultText(hr) : errorText));
    }
    return value;
}

void ThrowIfFailed(HRESULT hr, const wchar_t* name, const std::wstring& errorText) {
    if (SUCCEEDED(hr)) {
        return;
    }
    throw Error(hr, Describe(name, errorText.empty() ? HResultText(hr) : errorText));
}

}  // namespace

std::string ToUtf8(const std::wstring& text) {
    if (text.empty()) {
        return {};
    }
    const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                                         nullptr, 0, nullptr, nullptr);
    if (size <= 0) {
        return {};
    }
    std::string utf8(static_cast<std::size_t>(size), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                        utf8.data(), size, nullptr, nullptr);
    return utf8;
}

std::wstring HResultText(HRESULT hr) {
    wchar_t* buffer = nullptr;
    const DWORD length = FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr, static_cast<DWORD>(hr), MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        reinterpret_cast<wchar_t*>(&buffer), 0, nullptr);

    std::wstring text;
    if (length != 0 && buffer != nullptr) {
        text.assign(buffer, length);
        LocalFree(buffer);
        while (!text.empty() && (text.back() == L'\r' || text.back() == L'\n' || text.back() == L' ')) {
            text.pop_back();
        }
    }

    wchar_t code[16] = {};
    swprintf(code, 16, L"0x%08lX", static_cast<unsigned long>(hr));
    if (text.empty()) {
        text = code;
    } else {
        text += L" (";
        text += code;
        text += L")";
    }
    return text;
}

Error::Error(HRESULT code, const std::wstring& message)
    : code_(code), message_(message), text_(ToUtf8(message)) {}

const char* Error::what() const noexcept {
    return text_.c_str();
}

ComPtr::~ComPtr() {
    Reset();
}

ComPtr::ComPtr(ComPtr&& other) noexcept : object_(other.object_) {
    other.object_ = nullptr;
}

ComPtr& ComPtr::operator=(ComPtr&& other) noexcept {
    if (this != &other) {
        Reset();
        object_ = other.object_;
        other.object_ = nullptr;
    }
    return *this;
}

void ComPtr::Reset() {
    if (object_ != nullptr) {
        object_->Release();
        object_ = nullptr;
    }
}

IDispatch** ComPtr::Put() {
    Reset();
    return &object_;
}

void ComPtr::Attach(IDispatch* object) {
    Reset();
    object_ = object;
}

ArgList::ArgList(std::size_t count) : values_(count) {
    for (VARIANT& value : values_) {
        VariantInit(&value);
    }
}

ArgList::~ArgList() {
    for (VARIANT& value : values_) {
        VariantClear(&value);
    }
}

VARIANT* ArgList::Slot(std::size_t position) {
    if (position == 0 || position > values_.size()) {
        throw std::out_of_range("COM argument position out of range");
    }
    return &values_[values_.size() - position];
}

void ArgList::SetMissing(std::size_t position) {
    VARIANT* slot = Slot(position);
    VariantClear(slot);
    V_VT(slot) = VT_ERROR;
    V_ERROR(slot) = DISP_E_PARAMNOTFOUND;
}

void ArgList::SetBool(std::size_t position, bool value) {
    VARIANT* slot = Slot(position);
    VariantClear(slot);
    V_VT(slot) = VT_BOOL;
    V_BOOL(slot) = value ? VARIANT_TRUE : VARIANT_FALSE;
}

void ArgList::SetLong(std::size_t position, long value) {
    VARIANT* slot = Slot(position);
    VariantClear(slot);
    V_VT(slot) = VT_I4;
    V_I4(slot) = value;
}

void ArgList::SetString(std::size_t position, const std::wstring& value) {
    VARIANT* slot = Slot(position);
    VariantClear(slot);
    V_VT(slot) = VT_BSTR;
    V_BSTR(slot) = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
    if (V_BSTR(slot) == nullptr) {
        V_VT(slot) = VT_EMPTY;
        throw std::bad_alloc();
    }
}

HRESULT InvokeRaw(IDispatch* object, const wchar_t* name, WORD flags,
                  ArgList* args, VARIANT* result, std::wstring* errorText) {
    if (object == nullptr) {
        return E_POINTER;
    }

    DISPID id = DISPID_UNKNOWN;
    LPOLESTR names[] = {const_cast<LPOLESTR>(name)};
    HRESULT hr = object->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) {
        if (errorText != nullptr) {
            *errorText = HResultText(hr);
        }
        return hr;
    }

    DISPPARAMS params{};
    if (args != nullptr) {
        params.rgvarg = args->Data();
        params.cArgs = args->Count();
    }
    DISPID putId = DISPID_PROPERTYPUT;
    if ((flags & DISPATCH_PROPERTYPUT) != 0) {
        params.rgdispidNamedArgs = &putId;
        params.cNamedArgs = 1;
    }

    VARIANT localResult;
    VariantInit(&localResult);
    EXCEPINFO exception{};
    UINT argumentError = 0;
    hr = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &params,
                        result != nullptr ? result : &localResult, &exception, &argumentError);

    if (FAILED(hr) && errorText != nullptr) {
        if (exception.pfnDeferredFillIn != nullptr) {
            exception.pfnDeferredFillIn(&exception);
        }
        if (exception.bstrDescription != nullptr) {
            *errorText = exception.bstrDescription;
        } else if (exception.bstrSource != nullptr) {
            *errorText = exception.bstrSource;
        } else {
            *errorText = HResultText(hr);
        }
    }

    SysFreeString(exception.bstrSource);
    SysFreeString(exception.bstrDescription);
    SysFreeString(exception.bstrHelpFile);
    VariantClear(&localResult);
    return hr;
}

std::wstring GetString(IDispatch* object, const wchar_t* name) {
    VARIANT value = InvokeGet(object, name);
    std::wstring text;
    if (V_VT(&value) == VT_BSTR && V_BSTR(&value) != nullptr) {
        text.assign(V_BSTR(&value), SysStringLen(V_BSTR(&value)));
    } else if (V_VT(&value) != VT_EMPTY && V_VT(&value) != VT_NULL) {
        VARIANT converted;
        VariantInit(&converted);
        if (SUCCEEDED(VariantChangeType(&converted, &value, 0, VT_BSTR)) && V_BSTR(&converted) != nullptr) {
            text.assign(V_BSTR(&converted), SysStringLen(V_BSTR(&converted)));
        }
        VariantClear(&converted);
    }
    VariantClear(&value);
    return text;
}

long GetLong(IDispatch* object, const wchar_t* name) {
    VARIANT value = InvokeGet(object, name);
    long number = 0;
    if (V_VT(&value) == VT_I4) {
        number = V_I4(&value);
    } else {
        VARIANT converted;
        VariantInit(&converted);
        if (SUCCEEDED(VariantChangeType(&converted, &value, 0, VT_I4))) {
            number = V_I4(&converted);
        }
        VariantClear(&converted);
    }
    VariantClear(&value);
    return number;
}

bool GetBool(IDispatch* object, const wchar_t* name) {
    VARIANT value = InvokeGet(object, name);
    bool flag = false;
    if (V_VT(&value) == VT_BOOL) {
        flag = V_BOOL(&value) != VARIANT_FALSE;
    } else {
        VARIANT converted;
        VariantInit(&converted);
        if (SUCCEEDED(VariantChangeType(&converted, &value, 0, VT_BOOL))) {
            flag = V_BOOL(&converted) != VARIANT_FALSE;
        }
        VariantClear(&converted);
    }
    VariantClear(&value);
    return flag;
}

ComPtr GetObject(IDispatch* object, const wchar_t* name) {
    VARIANT value = InvokeGet(object, name);
    if (V_VT(&value) != VT_DISPATCH || V_DISPATCH(&value) == nullptr) {
        VariantClear(&value);
        throw Error(E_FAIL, Describe(name, L"did not return a COM object"));
    }
    ComPtr result;
    result.Attach(V_DISPATCH(&value));
    V_VT(&value) = VT_EMPTY;
    VariantClear(&value);
    return result;
}

void PutBool(IDispatch* object, const wchar_t* name, bool value) {
    ArgList args(1);
    args.SetBool(1, value);
    std::wstring errorText;
    ThrowIfFailed(InvokeRaw(object, name, DISPATCH_PROPERTYPUT, &args, nullptr, &errorText), name, errorText);
}

void PutLong(IDispatch* object, const wchar_t* name, long value) {
    ArgList args(1);
    args.SetLong(1, value);
    std::wstring errorText;
    ThrowIfFailed(InvokeRaw(object, name, DISPATCH_PROPERTYPUT, &args, nullptr, &errorText), name, errorText);
}

void PutString(IDispatch* object, const wchar_t* name, const std::wstring& value) {
    ArgList args(1);
    args.SetString(1, value);
    std::wstring errorText;
    ThrowIfFailed(InvokeRaw(object, name, DISPATCH_PROPERTYPUT, &args, nullptr, &errorText), name, errorText);
}

ComPtr CallObject(IDispatch* object, const wchar_t* name, ArgList& args) {
    VARIANT value;
    VariantInit(&value);
    std::wstring errorText;
    const HRESULT hr = InvokeRaw(object, name, DISPATCH_METHOD, &args, &value, &errorText);
    if (FAILED(hr)) {
        VariantClear(&value);
        throw Error(hr, Describe(name, errorText.empty() ? HResultText(hr) : errorText));
    }
    if (V_VT(&value) != VT_DISPATCH || V_DISPATCH(&value) == nullptr) {
        VariantClear(&value);
        throw Error(E_FAIL, Describe(name, L"did not return a COM object"));
    }
    ComPtr result;
    result.Attach(V_DISPATCH(&value));
    V_VT(&value) = VT_EMPTY;
    VariantClear(&value);
    return result;
}

void CallVoid(IDispatch* object, const wchar_t* name, ArgList& args) {
    std::wstring errorText;
    const HRESULT hr = InvokeRaw(object, name, DISPATCH_METHOD, &args, nullptr, &errorText);
    ThrowIfFailed(hr, name, errorText);
}

}  // namespace com
