#include "pdf.h"

#include <windows.h>

#include <qpdf/Constants.h>
#include <qpdf/QPDF.hh>
#include <qpdf/QPDFExc.hh>
#include <qpdf/QPDFWriter.hh>

#include <filesystem>
#include <string>

#include "util.h"

namespace {

std::string Utf8(const std::wstring& text) {
    if (text.empty()) {
        return {};
    }
    const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(),
                                         static_cast<int>(text.size()),
                                         nullptr, 0, nullptr, nullptr);
    if (size <= 0) {
        return {};
    }
    std::string out(static_cast<std::size_t>(size), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                        &out[0], size, nullptr, nullptr);
    return out;
}

// Opens the file with no password. A password-protected PDF raises
// qpdf_e_password; an owner-password-only PDF opens and reports isEncrypted().
bool Inspect(const std::string& file, bool& encrypted, std::wstring& error) {
    encrypted = false;
    try {
        QPDF qpdf;
        qpdf.processFile(file.c_str());
        encrypted = qpdf.isEncrypted();
        return true;
    } catch (const QPDFExc& e) {
        if (e.getErrorCode() == qpdf_e_password) {
            encrypted = true;
            return true;
        }
        error = util::FromUtf8(e.what());
        return false;
    } catch (const std::exception& e) {
        error = util::FromUtf8(e.what());
        return false;
    }
}

}  // namespace

FileResult PdfTool::ProcessFile(const std::wstring& path, const std::wstring& password,
                                bool encrypt) {
    FileResult result;
    result.action = FileAction::Failed;

    const std::string input = Utf8(path);
    const std::string secret = Utf8(password);

    if (secret.size() > 127) {
        result.detail = L"PDF passwords are limited to 127 bytes";
        return result;
    }

    bool encrypted = false;
    std::wstring error;
    if (!Inspect(input, encrypted, error)) {
        result.detail = error.empty() ? L"the PDF could not be read" : error;
        return result;
    }

    if (encrypt && encrypted) {
        result.action = FileAction::SkippedAlreadyEncrypted;
        return result;
    }
    if (!encrypt && !encrypted) {
        result.action = FileAction::SkippedNotEncrypted;
        return result;
    }

    std::filesystem::path temp = path;
    temp += L".xlpt-tmp";
    const std::string output = Utf8(temp.wstring());

    try {
        QPDF qpdf;
        qpdf.processFile(input.c_str(), secret.c_str());

        QPDFWriter writer(qpdf, output.c_str());
        if (encrypt) {
            // AES-256 (revision 6): the modern standard every reader supports.
            writer.setR6EncryptionParameters(secret.c_str(), secret.c_str(),
                                             true,   // accessibility
                                             true,   // extract
                                             true,   // assemble
                                             true,   // annotate and form
                                             true,   // form filling
                                             true,   // modify other
                                             qpdf_r3p_full,
                                             true);  // encrypt metadata
        } else {
            // qpdf keeps the source encryption unless told otherwise.
            writer.setPreserveEncryption(false);
        }
        writer.write();
    } catch (const QPDFExc& e) {
        util::RemoveQuiet(temp);
        result.detail = (e.getErrorCode() == qpdf_e_password)
                            ? L"wrong password"
                            : util::FromUtf8(e.what());
        return result;
    } catch (const std::exception& e) {
        util::RemoveQuiet(temp);
        result.detail = util::FromUtf8(e.what());
        return result;
    }

    // Make sure the file we are about to put in place really is what we expect.
    bool nowEncrypted = false;
    std::wstring verifyError;
    if (!Inspect(output, nowEncrypted, verifyError) || nowEncrypted != encrypt) {
        util::RemoveQuiet(temp);
        result.detail = L"the saved PDF does not look right - the original was left untouched";
        return result;
    }

    std::error_code ec;
    if (std::filesystem::file_size(temp, ec) == 0 || ec) {
        util::RemoveQuiet(temp);
        result.detail = L"the saved PDF is empty - the original was left untouched";
        return result;
    }

    if (!MoveFileExW(temp.c_str(), path.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        const DWORD code = GetLastError();
        util::RemoveQuiet(temp);
        result.detail = L"could not replace the file: " + util::Win32Text(code);
        return result;
    }

    result.action = encrypt ? FileAction::Encrypted : FileAction::Decrypted;
    return result;
}
