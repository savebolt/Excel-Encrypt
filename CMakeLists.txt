cmake_minimum_required(VERSION 3.16)
project(ExcelPasswordTool CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_executable(ExcelPasswordTool
    src/main.cpp
    src/com.cpp
    src/excel.cpp
    src/util.cpp
    src/watchdog.cpp
)

target_include_directories(ExcelPasswordTool PRIVATE src)
target_link_libraries(ExcelPasswordTool PRIVATE ole32 oleaut32 uuid shell32)

if (MSVC)
    target_compile_options(ExcelPasswordTool PRIVATE /W4 /EHsc /utf-8)
else()
    target_compile_options(ExcelPasswordTool PRIVATE -Wall -Wextra)
    find_package(Threads REQUIRED)
    target_link_libraries(ExcelPasswordTool PRIVATE Threads::Threads)
    if (MINGW)
        target_link_options(ExcelPasswordTool PRIVATE -static)
    endif()
endif()
