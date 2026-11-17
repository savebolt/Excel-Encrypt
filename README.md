# Excel Password Tool

A small C++ console program for Windows that applies real **"Encrypt with
password"** protection to every Excel, Word and PDF file in a folder, and can
remove the password again.

- **Excel** and **Word** are driven through COM, so those files get exactly the
  protection the Office apps themselves write (ECMA-376, AES + SHA-512).
- **PDF** uses the standard PDF security handler (AES-256, revision 6) through
  the embedded [qpdf](https://qpdf.readthedocs.io/) engine.

In every case the result opens in the normal application with nothing but the
password - no special tool required.

## Requirements

- Windows 10/11
- Desktop **Microsoft Excel** and/or **Microsoft Word** (2010 or newer) for
  those file types. PDF support needs no Office install.
- To build: MinGW-w64 g++ (MSYS2 UCRT64) or CMake. The static qpdf library is
  already in `third_party/qpdf/`.

## Build

Easiest (uses `g++` from PATH, or `C:\msys64\ucrt64\bin`):

```bat
build.bat
```

With CMake:

```bat
cmake -S . -B build -G "MinGW Makefiles"
cmake --build build
```

The result is one file: `ExcelPasswordTool.exe` (~8 MB, no runtime
dependencies besides Office for Excel/Word).

## Usage

Run `ExcelPasswordTool.exe` with no arguments for the menu. The default folder
is the current working directory (when started from Explorer: the exe folder).

```
   ============================================================
    Excel Password Tool   v1.1
    Encrypt / decrypt Excel, Word and PDF files with a password
   ============================================================
   Folder      : C:\Users\You\Documents\Reports
   Files       : 12   (Excel 5, Word 4, PDF 3)
   Types       : Excel + Word + PDF
   Subfolders  : no
   Backups     : yes (plain .bak copies)
   ------------------------------------------------------------
   [1] Encrypt all files  (set a password)
   [2] Decrypt all files  (remove the password)
   [3] Change folder
   [4] Include subfolders : OFF   (toggle)
   [5] Backups            : ON    (toggle)
   [6] List files
   [7] File types         : Excel + Word + PDF   (change)
   [0] Exit
```

Menu 1 asks for the password twice, then for confirmation. Menu 2 asks for the
existing password. The password is hidden while typing. Menu 7 limits the run
to a combination of `e` (Excel), `w` (Word) and `p` (PDF).

### Command line (scriptable, same engine)

```bat
ExcelPasswordTool.exe                                   :: menu, current folder
ExcelPasswordTool.exe "D:\Reports"                      :: menu, folder preselected
ExcelPasswordTool.exe -e -d "D:\Reports" -p "secret" -y
ExcelPasswordTool.exe -x -d "D:\Reports" -p "secret" -y
ExcelPasswordTool.exe -e --types p -d "D:\Reports" -p "secret" -y
```

| Option | Meaning |
| --- | --- |
| `-d, --folder <path>` | Folder to process (default: current folder) |
| `-e, --encrypt` | Encrypt all matching files, no menu |
| `-x, --decrypt` | Decrypt all matching files, no menu |
| `-p, --password <text>` | Password (otherwise asked, hidden) |
| `-r, --recursive` | Include subfolders |
| `--types <list>` | `e` Excel, `w` Word, `p` PDF - e.g. `ewp` or `all` |
| `--no-backup` | Do not create `.bak` copies |
| `-y, --yes` | Skip the confirmation question |
| `-h, --help` | Show help |

Exit codes: `0` success, `1` some files failed, `2` usage/password error.

Supported files:

| Type | Extensions |
| --- | --- |
| Excel | `.xlsx` `.xlsm` `.xlsb` `.xls` `.xltx` `.xltm` `.xlt` `.xlam` |
| Word | `.docx` `.docm` `.dotx` `.dotm` `.doc` `.dot` |
| PDF | `.pdf` |

Temporary `~$...` lock files and non-matching files are ignored.

## What it does to a file

**Excel / Word - encrypt**

1. Opens the document with the Office app.
2. Saves it with the password set (Word saves to a sibling temp file first,
   because Word ignores the password when overwriting the same path).
3. The result is the workbook/document-level *password to open* (not sheet or
   edit protection).

**Excel / Word - decrypt**

1. Opens the document with the given password.
2. Clears the open and write passwords.
3. Saves it without a password.

**PDF - encrypt**

1. Reads the PDF with qpdf.
2. Writes it back with AES-256 (revision 6) and the password as both the user
   and owner password, so permissions stay fully open.

**PDF - decrypt**

1. Opens the PDF with the password.
2. Writes it back with encryption removed.

Notes:

- Already encrypted files are skipped when encrypting. For OOXML formats this is
  detected from the container (encrypted OOXML is an OLE/CFB file, plain OOXML
  is a ZIP), so no password is needed for that check. PDFs are checked with
  qpdf.
- Unencrypted files are skipped when decrypting.
- Legacy `.xls` / `.doc` files are always OLE containers, so their state is
  checked through the Office app (`HasPassword`). A file encrypted with a
  *different* password is reported as a failure instead of guessing.
- Office passwords are limited to 255 characters; PDF passwords to 127 bytes.
- One password is used for the whole run.

## Backups

Backups are **on by default**: the first time a file is processed,
`name.ext.bak` is created next to it (only if that `.bak` does not already
exist, so the oldest copy is kept).

**The `.bak` copy is not encrypted.** Turn backups off with menu `[5]` or
`--no-backup` if plain copies must not exist.

## Safety and limitations

- Files are changed **in place**. Try it on copies of important files first.
- Close Excel/Word before running; a document open elsewhere opens read-only
  and the save fails (nothing is changed).
- Office rewrites the whole file on save. That is how the feature works:
  digital signatures are invalidated, and unusual content may be re-saved by
  the app. The `.bak` file is the untouched original.
- Macro documents keep their macros, but macros are disabled during automation
  (`AutomationSecurity = forceDisable`).
- If an Office app opens a modal dialog that blocks COM (rare), a watchdog
  terminates **only the automation instance** after 180 seconds and reports the
  file as failed. Override with the `XLPT_TIMEOUT` environment variable
  (seconds, 5-3600), e.g. `set XLPT_TIMEOUT=30`.
- OneDrive/SharePoint folders: let the sync client finish between runs.

## Troubleshooting

| Message | Cause / fix |
| --- | --- |
| `Microsoft Excel is not installed or not registered` | No desktop Excel installed, or Office needs repair |
| `Microsoft Word is not installed or not registered` | No desktop Word installed, or Office needs repair |
| `Open failed: The password you supplied is not correct` | Wrong decrypt password, or the file is already encrypted with another password |
| `SaveAs failed: Cannot access read-only document` | File is open in the app or marked read-only |
| `Excel/Word stopped responding and was closed` | A dialog blocked automation; the timeout killed the helper instance. Check the file, then retry |
| `the saved file does not look right` | Post-save signature check failed; restore the `.bak` and investigate |
| `wrong password` (PDF) | The PDF is encrypted with a different password |

## How it works

- Late-bound COM (`IDispatch` / `GetIDsOfNames` / `Invoke`) - no `#import`, so it
  builds with both MinGW g++ and MSVC.
- `Application` is started invisible with alerts, events, screen updating and
  interaction disabled, then reused for the whole batch.
- Excel: `Workbooks.Open(..., Password)` -> `Workbook.SaveAs(...)` ->
  `Workbook.Close(false)`.
- Word: `Documents.Open(..., PasswordDocument)` -> `Document.SaveAs2(...)` to a
  temp file -> verified -> atomically replaces the original.
- PDF: qpdf's `QPDF` / `QPDFWriter` API, linked statically into the exe.
- A watchdog thread (`src/watchdog.*`) kills the private Office process if a COM
  call exceeds the timeout, so the tool can never hang forever.

## Project layout

```
ExcelPasswordTool/
├─ build.bat            g++ build
├─ CMakeLists.txt       CMake build
├─ ExcelPasswordTool.exe
├─ src/
│  ├─ main.cpp          menu, CLI parsing, batch loop
│  ├─ excel.h/.cpp      Excel COM wrapper
│  ├─ word.h/.cpp       Word COM wrapper
│  ├─ pdf.h/.cpp        PDF encryption via qpdf
│  ├─ com.h/.cpp        COM smart pointer, late binding, error text
│  ├─ util.h/.cpp       console I/O, file scanning, signature detection
│  ├─ result.h          shared FileResult type
│  └─ watchdog.h/.cpp   process timeout guard
└─ third_party/qpdf/    static qpdf 12.4.2 (Apache-2.0) + libjpeg + zlib
```

## Third-party code

`third_party/qpdf/` contains a statically built **qpdf 12.4.2** (Apache License
2.0, see `third_party/qpdf/LICENSE.txt`), built with the native crypto provider
so no OpenSSL/GnuTLS is required, plus static **libjpeg** and **zlib**.
