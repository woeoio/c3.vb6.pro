@echo off
rem bench.bat - generic benchmark runner. ASCII only.
rem   %%1 = C3 exe path, %%2 = extra args, %%3 = clean flag (1 = wipe out)
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat" x64 >nul 2>&1
if errorlevel 1 exit /b 1
cd /d C:\Users\vi\Desktop\c3.vb6.pro\publish\demos\optbench
if "%3"=="1" if exist out rmdir /s /q out
set "EXTRA=%~2"
"%1" OptBench.vbp --output-dir "C:\Users\vi\Desktop\c3.vb6.pro\publish\demos\optbench\out" --incremental --no-warn 3001,3003 %EXTRA%
exit /b %ERRORLEVEL%
