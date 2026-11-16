@echo off
if exist "D:\devkit\bin\g++.exe" set "PATH=D:\devkit\bin;%PATH%"

where g++ >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] g++ not found. Please add w64devkit\bin to your PATH.
    pause
    exit /b 1
)

echo Compiling DocPulseConverter.exe (C++20 / Win32 / WIC)...
g++ -std=c++20 -O3 -march=skylake -ffunction-sections -fdata-sections -Wl,--gc-sections -s -municode -mwindows DocPulseConverter.cpp -o DocPulseConverter.exe -ldwmapi -lcomdlg32 -lshell32 -lole32 -lwindowscodecs -luuid


if %errorlevel% equ 0 (
    echo.
    echo [SUCCESS] DocPulseConverter.exe compiled cleanly!
    for %%I in (DocPulseConverter.exe) do echo Final size: %%~zI bytes
) else (
    echo.
    echo [ERROR] Compilation failed.
)
pause
