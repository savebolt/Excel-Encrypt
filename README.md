# Excel Password Tool

A small C++ console program for Windows that applies Excel's real
**"Encrypt with password"** to every workbook in a folder, and can remove the
password again.

It drives the installed Microsoft Excel through COM, so the files are encrypted
by Excel itself (ECMA-376, AES + SHA-512). Files open normally in Excel,
Office, and other tools that understand Excel encryption.

## Requirements

- Windows 10/11
- Desktop **Microsoft Excel 2010 or newer** installed (COM automation)
- To build: MinGW-w64 g++ (MSYS2 UCRT64) or Visual Studio 2019+, CMake optional

## Build

Easiest (uses `g++` from PATH, or `C:\msys64\ucrt64\bin`):

```bat
build.bat
```

With CMake:

```bat
cmake -S . -B build
cmake --build build --config Release
```

With MSVC (Developer Command Prompt):

```bat
cl /std:c++17 /EHsc /O2 src\*.cpp /Fe:ExcelPasswordTool.exe ole32.lib oleaut32.lib uuid.lib shell32.lib
```

The result is one file: `ExcelPasswordTool.exe` (no runtime dependencies besides
Excel).

## Usage

Run `ExcelPasswordTool.exe` with no arguments for the menu. The default folder
is the current working directory (when started from Explorer: the exe folder).

```
   ============================================================
    Excel Password Tool   v1.0
    Encrypt / decrypt Excel files with a password
   ============================================================
   Folder      : C:\Users\You\Documents\Reports
   Excel files : 12
   Subfolders  : no
   Backups     : yes (plain .bak copies)
   ------------------------------------------------------------
   [1] Encrypt all Excel files  (set a password)
   [2] Decrypt all Excel files  (remove the password)
   [3] Change folder
   [4] Include subfolders : OFF   (toggle)
   [5] Backups            : ON    (toggle)
   [6] List files
   [0] Exit
```

Menu 1 asks for the password twice, then for confirmation. Menu 2 asks for the
existing password. The password is hidden while typing.

### Command line (scriptable, same engine)

```bat
ExcelPasswordTool.exe                          :: menu, current folder
ExcelPasswordTool.exe "D:\Reports"             :: menu, folder preselected
ExcelPasswordTool.exe -e -d "D:\Reports" -p "secret" -y
ExcelPasswordTool.exe -x -d "D:\Reports" -p "secret" -y
```

| Option | Meaning |
| --- | --- |
| `-d, --folder <path>` | Folder to process (default: current folder) |
| `-e, --encrypt` | Encrypt all Excel files, no menu |
| `-x, --decrypt` | Decrypt all Excel files, no menu |
| `-p, --password <text>` | Password (otherwise asked, hidden) |
| `-r, --recursive` | Include subfolders |
| `--no-backup` | Do not create `.bak` copies |
| `-y, --yes` | Skip the confirmation question |
| `-h, --help` | Show help |

Exit codes: `0` success, `1` some files failed, `2` usage/password error.

Supported files: `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.xltx`, `.xltm`, `.xlt`,
`.xlam`. Temporary `~$...` lock files and non-Excel files are ignored.

## What it does to a file

**Encrypt**

1. Opens the workbook with Excel.
2. Saves it in place with `SaveAs(Filename, FileFormat, Password)`.
3. Encryption is the workbook-level *password to open* (not sheet protection).

**Decrypt**

1. Opens the workbook with the given password.
2. Clears `Workbook.Password` and `Workbook.WritePassword`.
3. Saves it in place without a password.

Notes:

- Already encrypted files are skipped when encrypting. For OOXML formats this is
  detected from the container (encrypted OOXML is an OLE/CFB file, plain OOXML
  is a ZIP), so no password is needed for that check.
- Unencrypted OOXML files are skipped when decrypting.
- Legacy `.xls` files are always OLE containers, so their state is checked via
  Excel (`Workbook.HasPassword`). An `.xls` encrypted with a *different*
  password makes the tool report a failure instead of guessing.
- Excel passwords are limited to 255 characters. One password is used for the
  whole run.

## Backups

Backups are **on by default**: the first time a file is processed,
`name.ext.bak` is created next to it (only if that `.bak` does not already
exist, so the oldest copy is kept).

**The `.bak` copy is not encrypted.** Turn backups off with menu `[5]` or
`--no-backup` if plain copies must not exist.

## Safety and limitations

- Files are changed **in place**. Try it on copies of important files first.
- Close Excel before running; a workbook open elsewhere opens read-only and the
  save fails with *"Cannot access read-only document"* (nothing is changed).
- Excel rewrites the whole file on `SaveAs`. That is how the feature works:
  digital signatures are invalidated, and unusual content may be re-saved by
  Excel. The `.bak` file is the untouched original.
- Macro workbooks keep their macros (`SaveFormat 52`), but macros are disabled
  during automation (`AutomationSecurity = forceDisable`).
- If Excel opens a modal dialog that blocks COM (rare), a watchdog terminates
  **only the automation instance** after 180 seconds and reports the file as
  failed. Override with the `XLPT_TIMEOUT` environment variable (seconds,
  5–3600), e.g. `set XLPT_TIMEOUT=30`.
- OneDrive/SharePoint folders: let Excel/OneDrive finish syncing between runs.

## Troubleshooting

| Message | Cause / fix |
| --- | --- |
| `Microsoft Excel is not installed or not registered` | No desktop Excel installed, or Office needs repair |
| `Open failed: The password you supplied is not correct` | Wrong decrypt password, or the file is already encrypted with another password |
| `SaveAs failed: Cannot access read-only document` | File is open in Excel or marked read-only |
| `Excel stopped responding and was closed` | A dialog blocked automation; the timeout killed the helper instance. Check the file, then retry |
| `the saved file does not look right` | Post-save signature check failed; restore the `.bak` and investigate |

## How it works

- Late-bound COM (`IDispatch` / `GetIDsOfNames` / `Invoke`) — no `#import`, so it
  builds with both MinGW g++ and MSVC.
- `Application` is started invisible with alerts, events, screen updating and
  interaction disabled, then reused for the whole batch.
- Per file: `Workbooks.Open(..., Password)` → optional password clear →
  `Workbook.SaveAs(...)` → `Workbook.Close(false)`.
- A watchdog thread (`src/watchdog.*`) kills the private Excel process if a COM
  call exceeds the timeout, so the tool can never hang forever.

## Project layout

```
ExcelPasswordTool/
├─ build.bat            g++ build
├─ CMakeLists.txt       CMake build
├─ ExcelPasswordTool.exe
└─ src/
   ├─ main.cpp          menu, CLI parsing, batch loop
   ├─ excel.h/.cpp      Excel COM wrapper, encrypt/decrypt logic
   ├─ com.h/.cpp        COM smart pointer, late binding, error text
   ├─ util.h/.cpp       console I/O, file scanning, signature detection
   └─ watchdog.h/.cpp   process timeout guard
```
