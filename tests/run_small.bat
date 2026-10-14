@echo off
set VCVARS=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat
set C3=C:\Users\vi\Desktop\c3.vb6.pro\.build\C3.exe
call "%VCVARS%" x64 >nul 2>&1
if errorlevel 1 (
    echo [ERROR] vcvarsall failed
    exit /b 1
)
cd /d C:\Users\vi\Desktop\c3.vb6.pro\tests
echo === Compile test_class.vbp (default mode) ===
"%C3%" "test_class.vbp" --output-dir C:\Users\vi\Desktop\c3.vb6.pro\output 2>&1
echo EXIT: %ERRORLEVEL%
