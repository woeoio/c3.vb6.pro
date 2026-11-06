@echo off
rem crashprobe build: VCVARSALL env var first, then vswhere, then known fallback.
rem Builds BOTH bitnesses from the same crashprobe.cpp:
rem   crashprobe.exe    (x64) -- debug 64-bit targets
rem   crashprobe32.exe  (x86) -- debug 32-bit targets (C3 default output)
rem The debugger must match the target bitness; tools\crashprobe\crashprobe.ps1
rem reads the target PE machine field and picks the right binary automatically.
rem NOTE: keep this file ASCII-only -- cmd.exe reads UTF-8 Chinese as ANSI.
setlocal
if not defined VCVARSALL (
    set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
    if not exist "%VSWHERE%" set "VSWHERE=%ProgramFiles%\Microsoft Visual Studio\Installer\vswhere.exe"
    if exist "%VSWHERE%" (
        for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do (
            set "VCVARSALL=%%i\VC\Auxiliary\Build\vcvarsall.bat"
        )
    )
)
if not defined VCVARSALL if exist "D:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat" (
    set "VCVARSALL=D:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat"
)
if not defined VCVARSALL (
    echo [crashprobe] VCVARSALL not found; set it to your vcvarsall.bat path and retry.
    exit /b 2
)
cd /d "%~dp0"

rem ---- x64 ----
call "%VCVARSALL%" x64 >nul 2>&1
cl /nologo /EHsc /W3 crashprobe.cpp /Fo:crashprobe_x64.obj /link /out:crashprobe.exe dbghelp.lib psapi.lib
if errorlevel 1 goto :fail

rem ---- x86 ----
rem The second call resets PATH/INCLUDE/LIB to the x86 toolset (vcvarsall supports it).
call "%VCVARSALL%" x86 >nul 2>&1
cl /nologo /EHsc /W3 crashprobe.cpp /Fo:crashprobe32.obj /link /out:crashprobe32.exe dbghelp.lib psapi.lib
if errorlevel 1 goto :fail

echo [crashprobe] built: crashprobe.exe (x64) + crashprobe32.exe (x86)
exit /b 0

:fail
echo [crashprobe] build FAILED
exit /b 1
