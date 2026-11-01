@echo off
REM ============================================================
REM build.bat - portable build for C3
REM
REM NOTE: for day-to-day development prefer ./build.sh -- same build, but it
REM       drives ninja from bash instead of cmd.exe.  cmd.exe spawns child
REM       processes at ~82 ms each once vcvars has padded the environment,
REM       versus ~17 ms from bash, so a no-op build costs 9.3 s here and
REM       ~0.5 s there.  build.bat stays the portable / CI path.
REM
REM Locates the Visual Studio C++ toolset via vswhere (no hardcoded
REM install path), sets up the x64 environment, then runs the existing
REM build. Usage:
REM     build.bat                (builds .build, the default)
REM     build.bat --clean        (reconfigure + build .build)
REM     build.bat --dir some\dir (build a different build dir)
REM
REM Ninja is always taken from the Visual Studio installation - never from
REM PATH, so no third-party ninja can be picked up. The VS instance is
REM therefore resolved with the optional "CMake tools for Visual Studio"
REM component required as well: if the newest VS with the C++ toolset does
REM not ship Ninja, the newest VS that does is used instead and a NOTICE is
REM printed, so the toolchain choice is never silent. If no VS installation
REM ships Ninja at all, this fails loudly instead of letting CMake report a
REM bare "ninja not found".
REM ============================================================
setlocal EnableExtensions
cd /d "%~dp0"

set "BDIR=.build"
set "CLEAN=0"
:parseargs
if "%~1"=="" goto :found_bdir
if /i "%~1"=="--clean" set "CLEAN=1" & shift & goto :parseargs
if /i "%~1"=="--dir" goto :get_dir
echo [build] unknown argument: %~1
exit /b 1
:get_dir
REM shift happens outside a block on purpose: %~1 inside a parenthesized block
REM is expanded once before the block runs, so a shift in there left BDIR set
REM to the literal string --dir. The error check must not use "& exit" either -
REM outside a block that exit would run unconditionally.
shift
if not "%~1"=="" goto :set_dir
echo [build] --dir needs an argument
exit /b 1
:set_dir
set "BDIR=%~1"
shift
goto :parseargs
:found_bdir

REM --- locate vswhere (32-bit ProgramFiles is the documented location) ---
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" set "VSWHERE=%ProgramFiles%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" goto :no_vswhere

REM --- newest VS shipping BOTH the x64 C++ toolchain AND Ninja ---
REM Ninja travels with the optional "CMake tools for Visual Studio" component.
REM Requiring only the VC toolset can resolve a VS installation that has the
REM compiler but no Ninja, which is what made "ninja not found" appear
REM intermittently on machines holding more than one Visual Studio.
set "VSNINJA="
for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -requires Microsoft.VisualStudio.Component.VC.CMake.Project -property installationPath`) do set "VSNINJA=%%i"

REM --- newest VS with the C++ toolchain, Ninja or not ---
set "VSTOOL="
for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSTOOL=%%i"

if not defined VSNINJA goto :no_ninja_available
set "VSINST=%VSNINJA%"
if not defined VSTOOL goto :vs_ready
if /i not "%VSNINJA%"=="%VSTOOL%" goto :notice_ninja
goto :vs_ready
:notice_ninja
echo [build] NOTICE   : newest VS with the C++ toolset ships no Ninja; using the newest one that does
goto :vs_ready
:no_ninja_available
set "VSINST=%VSTOOL%"
if not defined VSINST goto :no_vswhere
goto :no_ninja
:vs_ready
if not exist "%VSINST%\VC\Auxiliary\Build\vcvars64.bat" goto :no_vswhere
echo [build] Visual Studio: %VSINST%

REM --- VS ships the Ninja generator; put it on PATH so --clean -G Ninja works ---
set "PATH=%VSINST%\Common7\IDE\CommonExtensions\Microsoft\CMake\Ninja;%PATH%"

call "%VSINST%\VC\Auxiliary\Build\vcvars64.bat"
if errorlevel 1 goto :vcvars_failed

REM --- resolve cmake (prefer PATH, fall back to the copy VS ships with) ---
set "CMAKE_EXE="
for /f "delims=" %%p in ('where.exe cmake 2^>nul') do set "CMAKE_EXE=%%p"
if not defined CMAKE_EXE (
    if exist "%VSINST%\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe" (
        set "CMAKE_EXE=%VSINST%\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe"
    )
)
if not defined CMAKE_EXE goto :no_cmake
echo [build] CMake       : %CMAKE_EXE%

if %CLEAN%==1 goto :do_build

REM --- a stale build dir can record a ninja.exe that no longer exists (VS
REM upgraded, repaired or moved). CMake fails cryptically in that case, so
REM report it and name the remedy instead. ---
set "CACHED_MAKE=__none__"
for /f "tokens=1,2 delims==" %%a in ('findstr /b "CMAKE_MAKE_PROGRAM:" "%BDIR%\CMakeCache.txt" 2^>nul') do set "CACHED_MAKE=%%b"
if "%CACHED_MAKE%"=="__none__" goto :do_build
if exist "%CACHED_MAKE%" goto :do_build
echo [build] ERROR: %BDIR% was configured with "%CACHED_MAKE%" which no longer exists
echo [build]           (Visual Studio was upgraded, repaired or moved)
echo [build]           re-run: build.bat --clean
exit /b 1

:do_build
if %CLEAN%==1 (
    "%CMAKE_EXE%" -S . -B "%BDIR%" -G Ninja
    if errorlevel 1 goto :cmake_failed
)

"%CMAKE_EXE%" --build "%BDIR%"
if errorlevel 1 goto :cmake_failed
exit /b 0

:no_ninja
echo [build] ERROR: the selected Visual Studio does not ship Ninja:
echo [build]     %VSINST%
echo [build]     Install the "CMake tools for Visual Studio" component into that
echo [build]     installation (or into a VS 2019+ installation that has it),
echo [build]     then retry. Ninja is taken from VS only - never from PATH.
exit /b 1
:no_vswhere
echo [build] ERROR: no Visual Studio C++ toolset found via vswhere
exit /b 1
:vcvars_failed
echo [build] ERROR: vcvars64.bat failed
exit /b 1
:no_cmake
echo [build] ERROR: cmake not found (not on PATH and not bundled with Visual Studio)
exit /b 1
:cmake_failed
echo [build] ERROR: cmake failed
exit /b 1
