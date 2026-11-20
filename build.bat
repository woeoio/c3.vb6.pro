@echo off
REM ============================================================
REM build.bat - portable build for C3
REM
REM NOTE: for day-to-day development prefer ./build.sh -- same build, but it
REM       drives ninja from bash instead of cmd.exe.  cmd.exe spawns child
REM       processes at ~82 ms each once vcvars has padded the environment,
REM       versus ~17 ms from bash, so a no-op build costs ~9.3 s here and
REM       ~0.5 s there.  build.bat stays the portable / CI path.
REM
REM Usage:
REM     build.bat                (builds .build, the default)
REM     build.bat --clean        (reconfigure + build .build)
REM     build.bat --dir some\dir (build a different build dir)
REM     build.bat --no-vswhere   (skip vswhere, force filesystem scan -- self-test)
REM     build.bat --print-vs     (print what would be used and exit)
REM
REM Toolchain discovery is deliberately redundant, because "works on my
REM machine" is exactly the failure this script had:
REM   1. %C3_VS_INSTALL%            -- explicit override (portable / unusual layout)
REM   2. vswhere.exe                -- authoritative, handles non-default roots
REM   3. standard install roots     -- VS2022/2019/2017 x {Enterprise,Professional,
REM                                    Community,BuildTools} on %ProgramFiles%,
REM                                    %ProgramFiles(x86)% and %SystemDrive%
REM   4. registry HKCU/HKLM ...\VisualStudio\SxS\VS7 -- legacy installs
REM   5. where.exe devenv / MSBuild -- last resort: derive the install root
REM
REM Every path above may itself be missing: on non-interactive logons (services,
REM some CI agents, this repo's automation) %ProgramFiles(x86)% and even
REM %ProgramFiles% can be EMPTY, and the old script turned that into a bare
REM "no Visual Studio C++ toolset found" with no hint.  So nothing here relies
REM on a single environment variable being present.
REM
REM Ninja: VS ships it under Common7\IDE\CommonExtensions\Microsoft\CMake\Ninja.
REM That component is optional, so we fall back to ninja.exe on PATH, and only
REM fail loudly when neither exists.  %C3_NINJA% overrides both.
REM ============================================================
setlocal EnableExtensions
cd /d "%~dp0"

set "BDIR=.build"
set "CLEAN=0"
set "SKIP_VSWHERE=0"
set "PRINT_ONLY=0"
:parseargs
if "%~1"=="" goto :found_bdir
if /i "%~1"=="--clean" set "CLEAN=1" & shift & goto :parseargs
if /i "%~1"=="--no-vswhere" set "SKIP_VSWHERE=1" & shift & goto :parseargs
if /i "%~1"=="--print-vs" set "PRINT_ONLY=1" & shift & goto :parseargs
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

REM ============================================================
REM 1) explicit override
REM ============================================================
set "VSINST="
if defined C3_VS_INSTALL (
    if exist "%C3_VS_INSTALL%\VC\Auxiliary\Build\vcvars64.bat" set "VSINST=%C3_VS_INSTALL%"
)

REM ============================================================
REM 2) vswhere
REM ============================================================
set "VSWHERE="
if "%SKIP_VSWHERE%"=="0" call :find_vswhere
if not defined VSINST if defined VSWHERE call :pick_via_vswhere

REM ============================================================
REM 3) standard install roots -- %ProgramFiles% may be EMPTY in a stripped
REM    environment (observed on this very machine), so this alone is not enough.
REM ============================================================
if not defined VSINST call :scan_roots

REM ============================================================
REM 3b) every fixed drive: VS does NOT have to live on C:.  On this machine
REM     it is on D:, with %ProgramFiles% empty -- which is exactly why the old
REM     script reported "no toolset" while vcvars64.bat was sitting on disk.
REM     Fixed disks only (DriveType=3) so a dead network mapping cannot hang us.
REM ============================================================
if not defined VSINST call :scan_drives

REM ============================================================
REM 4) registry: ...\VisualStudio\SxS\VS7 holds "<ver>" = "<root>" (VS2015 era,
REM    still present on some machines); harmless when absent.
REM ============================================================
if not defined VSINST call :scan_registry

if not defined VSINST goto :no_vs

if not exist "%VSINST%\VC\Auxiliary\Build\vcvars64.bat" goto :no_vs
echo [build] Visual Studio: %VSINST%

REM ============================================================
REM Ninja: %C3_NINJA% -> VS-bundled -> PATH.
REM ============================================================
set "NINJA="
if defined C3_NINJA if exist "%C3_NINJA%" set "NINJA=%C3_NINJA%"
if not defined NINJA if exist "%VSINST%\Common7\IDE\CommonExtensions\Microsoft\CMake\Ninja\ninja.exe" set "NINJA=%VSINST%\Common7\IDE\CommonExtensions\Microsoft\CMake\Ninja\ninja.exe"
if not defined NINJA for /f "delims=" %%p in ('where.exe ninja 2^>nul') do if not defined NINJA set "NINJA=%%p"
if not defined NINJA goto :no_ninja
echo [build] Ninja       : %NINJA%
for %%p in ("%NINJA%") do set "NINJA_DIR=%%~dp"
set "PATH=%NINJA_DIR%%PATH%"

REM ============================================================
REM CMake: PATH first, then the copy VS bundles.
REM ============================================================
set "CMAKE_EXE="
for /f "delims=" %%p in ('where.exe cmake 2^>nul') do if not defined CMAKE_EXE set "CMAKE_EXE=%%p"
if not defined CMAKE_EXE if exist "%VSINST%\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe" set "CMAKE_EXE=%VSINST%\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe"
if not defined CMAKE_EXE goto :no_cmake
echo [build] CMake       : %CMAKE_EXE%

if "%PRINT_ONLY%"=="1" exit /b 0

call "%VSINST%\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 goto :vcvars_failed

if not exist "%BDIR%\CMakeCache.txt" set "CLEAN=1"
if "%CLEAN%"=="1" goto :do_build

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

REM ============================================================
REM Subroutines
REM ============================================================

:find_vswhere
REM Documented location first, but never assume the environment has the vars:
REM an empty %ProgramFiles(x86)% must degrade to the %SystemDrive% form, not to
REM a broken path like "\Microsoft Visual Studio\Installer\vswhere.exe".
for %%d in ("%ProgramFiles(x86)%" "%ProgramFiles%" "%SystemDrive%\Program Files (x86)" "%SystemDrive%\Program Files") do (
    if not defined VSWHERE call :probe_vswhere "%%~d"
)
if not defined VSWHERE for /f "delims=" %%p in ('where.exe vswhere 2^>nul') do if not defined VSWHERE set "VSWHERE=%%p"
exit /b 0

:probe_vswhere
if defined VSWHERE exit /b 0
if "%~1"=="" exit /b 0
if exist "%~1\Microsoft Visual Studio\Installer\vswhere.exe" set "VSWHERE=%~1\Microsoft Visual Studio\Installer\vswhere.exe"
exit /b 0

:pick_via_vswhere
REM -latest + -products * covers Community/Professional/Enterprise/BuildTools and
REM preview installs.  We ask only for the C++ toolset; Ninja is resolved
REM separately so that a VS without the optional CMake component still builds.
set "VS_TOOL="
for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VS_TOOL=%%i"
if not defined VS_TOOL exit /b 0
if exist "%VS_TOOL%\VC\Auxiliary\Build\vcvars64.bat" set "VSINST=%VS_TOOL%"
exit /b 0

:scan_roots
for %%r in ("%SystemDrive%" "%ProgramFiles%" "%ProgramFiles(x86)%") do call :scan_one_volume "%%~r"
exit /b 0

:scan_drives
for /f "skip=1 tokens=1" %%v in ('wmic logicaldisk where "DriveType=3" get DeviceID 2^>nul') do call :scan_one_volume "%%v"
if not defined VSINST for %%d in (C D E F G H) do call :scan_one_volume "%%d:"
exit /b 0

REM A VS install root may sit directly under the drive or under either
REM "Program Files" spelling -- on this machine it is
REM D:\Program Files (x86)\Microsoft Visual Studio\2019\Community while
REM %ProgramFiles(x86)% is empty, so both shapes must be probed on every volume.
:scan_one_volume
if defined VSINST exit /b 0
if "%~1"=="" exit /b 0
call :scan_one_base "%~1\Microsoft Visual Studio"
call :scan_one_base "%~1\Program Files\Microsoft Visual Studio"
call :scan_one_base "%~1\Program Files (x86)\Microsoft Visual Studio"
exit /b 0

:scan_one_base
if defined VSINST exit /b 0
for %%y in (2022 2019 2017) do call :scan_one_year "%~1\%%y"
exit /b 0

:scan_one_year
if defined VSINST exit /b 0
for %%e in (Enterprise Professional Community BuildTools) do call :scan_one_edition "%~1\%%e"
exit /b 0

:scan_one_edition
if defined VSINST exit /b 0
if exist "%~1\VC\Auxiliary\Build\vcvars64.bat" set "VSINST=%~1"
exit /b 0

:scan_registry
REM VS2017+ writes the install root under ...\VisualStudio\SxS\VS7 as
REM "<major>.0" = "<root>"; it is the only hint left when the Installer
REM directory itself is gone (portable/extracted VS, stripped images).
for %%k in (HKLM HKCU) do call :scan_registry_key "%%k"
exit /b 0

:scan_registry_key
if defined VSINST exit /b 0
for /f "tokens=2,*" %%a in ('reg query "%~1\SOFTWARE\Microsoft\VisualStudio\SxS\VS7" 2^>nul ^| findstr /r /c:"REG_SZ"') do call :probe_vs_root "%%~b"
if defined VSINST exit /b 0
for /f "tokens=2,*" %%a in ('reg query "%~1\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\SxS\VS7" 2^>nul ^| findstr /r /c:"REG_SZ"') do call :probe_vs_root "%%~b"
exit /b 0

:probe_vs_root
if defined VSINST exit /b 0
if "%~1"=="" exit /b 0
if exist "%~1\VC\Auxiliary\Build\vcvars64.bat" set "VSINST=%~1"
exit /b 0

REM ============================================================
REM Errors
REM ============================================================

:no_vs
echo [build] ERROR: no Visual Studio C++ toolset (x64) found.
echo [build]        Looked for: %%C3_VS_INSTALL%%, vswhere, standard install roots,
echo [build]        HKCU/HKLM ...\VisualStudio\SxS\VS7.
echo [build]        Fix one of these:
echo [build]          * install the "Desktop development with C++" workload, or
echo [build]          * set C3_VS_INSTALL to the folder that contains
echo [build]            VC\Auxiliary\Build\vcvars64.bat, e.g.
echo [build]            set C3_VS_INSTALL=C:\Program Files\Microsoft Visual Studio\2022\BuildTools
exit /b 1
:no_ninja
echo [build] ERROR: ninja.exe not found.
echo [build]        Looked for: %%C3_NINJA%%, VS-bundled Ninja, ninja on PATH.
echo [build]        Install the "CMake tools for Visual Studio" component (it
echo [build]        ships Ninja), or put ninja.exe on PATH, or set C3_NINJA to
echo [build]        its full path.
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
