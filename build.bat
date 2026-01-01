@echo off
echo ========================================
echo NewFileAndFolderShellExtension - Build
echo ========================================
echo.

REM Find Visual Studio installation
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"

if not exist "%VSWHERE%" (
    echo ERROR: Visual Studio installer not found!
    echo Please install Visual Studio 2022 with C++ tools.
    pause
    exit /b 1
)

REM Get VS installation path
for /f "usebackq tokens=*" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do (
    set "VSPATH=%%i"
)

if not defined VSPATH (
    echo ERROR: Visual Studio C++ tools not found!
    echo Please install Visual Studio 2022 with "Desktop development with C++" workload.
    pause
    exit /b 1
)

echo Found Visual Studio at: %VSPATH%
echo.

REM Setup Visual Studio environment (suppress verbose output)
call "%VSPATH%\VC\Auxiliary\Build\vcvars64.bat" >nul

echo.
echo Compiling shell extension DLL...
echo.

REM Compile
cl /std:c++17 /O2 /EHsc /nologo /c /D_USRDLL /D_WINDLL /DUNICODE /D_UNICODE NewFileAndFolderShellExtension.cpp DllMain.cpp

set COMPILE_SUCCESS=%ERRORLEVEL%

if %COMPILE_SUCCESS% NEQ 0 (
    echo.
    echo ========================================
    echo Compilation: FAILED
    echo ========================================
    pause
    exit /b 1
)

echo.
echo Linking...
echo.

REM Link
link /DLL /OUT:NewFileAndFolderShellExtension.dll /DEF:NewFileAndFolderShellExtension.def /nologo NewFileAndFolderShellExtension.obj DllMain.obj ole32.lib oleaut32.lib shlwapi.lib shell32.lib user32.lib gdi32.lib

set LINK_SUCCESS=%ERRORLEVEL%

echo.
echo ========================================
if %LINK_SUCCESS% EQU 0 (
    echo Build: SUCCESSFUL
    echo   Output: NewFileAndFolderShellExtension.dll
) else (
    echo Build: FAILED
)
echo ========================================
echo.

if %LINK_SUCCESS% EQU 0 (
    echo To install, run as Administrator:
    echo   install.bat
    echo.
    echo Or manually:
    echo   regsvr32 NewFileAndFolderShellExtension.dll
    echo.
)

REM Clean up object files
del /q *.obj 2>nul

pause
