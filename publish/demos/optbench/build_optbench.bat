@echo off
rem build_optbench.bat - build optbench demo with all optimizations. ASCII only.
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat" x64 >nul 2>&1
if errorlevel 1 exit /b 1
cd /d C:\Users\vi\Desktop\c3.vb6.pro\publish\demos\optbench
"C:\Users\vi\Desktop\c3.vb6.pro\.build\C3.exe" OptBench.vbp --output-dir "C:\Users\vi\Desktop\c3.vb6.pro\publish\demos\optbench\out" --incremental --no-warn 3001,3003 %*
exit /b %ERRORLEVEL%
