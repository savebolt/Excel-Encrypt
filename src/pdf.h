#pragma once

#include <string>

#include "result.h"

// Password-protects PDF files with the standard PDF security handler
// (AES-256, revision 6), so they open in Adobe Reader, Chrome, Edge and any
// other PDF viewer with nothing but the password.
//
// The qpdf engine is linked into the tool itself - no external program.
class PdfTool {
public:
    FileResult ProcessFile(const std::wstring& path, const std::wstring& password, bool encrypt);
};
