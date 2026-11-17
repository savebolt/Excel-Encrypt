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

%CXX% -std=c++17 -O2 -Wall -Wextra -pthread ^
    src\main.cpp src\com.cpp src\excel.cpp src\util.cpp src\watchdog.cpp ^
    -o ExcelPasswordTool.exe ^
    -static -lole32 -loleaut32 -luuid -lshell32

if errorlevel 1 (
    echo.
    echo Build failed.
    exit /b 1
)

echo.
echo Built ExcelPasswordTool.exe
endlocal
