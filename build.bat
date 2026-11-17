@echo off
setlocal
cd /d "%~dp0"

set "CXX=g++"
where g++ >nul 2>nul
if errorlevel 1 (
    if exist "C:\msys64\ucrt64\bin\g++.exe" (
        set "CXX=C:\msys64\ucrt64\bin\g++.exe"
        set "PATH=C:\msys64\ucrt64\bin;%PATH%"
    ) else (
        echo g++ was not found.
        echo Install MSYS2 and its UCRT64 g++ toolchain, or build with CMake / Visual Studio.
        exit /b 1
    )
)

%CXX% -std=c++20 -O2 -Wall -Wextra -pthread ^
    -I third_party\qpdf\include ^
    src\main.cpp src\com.cpp src\excel.cpp src\word.cpp src\pdf.cpp src\util.cpp src\watchdog.cpp ^
    -o ExcelPasswordTool.exe ^
    -static -lole32 -loleaut32 -luuid -lshell32 -ladvapi32 ^
    third_party\qpdf\lib\libqpdf.a third_party\qpdf\lib\libjpeg.a third_party\qpdf\lib\libz.a

if errorlevel 1 (
    echo.
    echo Build failed.
    exit /b 1
)

echo.
echo Built ExcelPasswordTool.exe
endlocal
