#pragma once

#include <string>

enum class FileAction {
    Failed,
    Encrypted,
    Decrypted,
    SkippedAlreadyEncrypted,
    SkippedNotEncrypted
};

struct FileResult {
    FileAction action = FileAction::Failed;
    std::wstring detail;
};
