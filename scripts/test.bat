@echo off
REM ============================================================
REM  test.bat - 运行回归测试
REM  用法: test [all|run|compile|syntax|smoke|gui] [verbose]
REM    test          = 全部测试
REM    test run      = 仅运行测试(编译+运行+输出校验)
REM    test compile  = 仅编译测试
REM    test syntax   = 仅语法测试
REM    test smoke    = 仅冒烟测试(tests\smoke.bas, 构建后快速验证产物可用)
REM    test gui      = 仅 GUI 表 (M2 数据驱动用例, 同样会随 vbp 分类在 CI 里跑)
REM  一键基线(M1, 同时被 CI 调用): powershell -File tests\baseline.ps1
REM ============================================================

REM 路径由脚本自身位置推导, 不依赖固定盘符
cd /d "%~dp0.."
if errorlevel 1 (
    echo [ERROR] 无法进入项目根目录
    exit /b 1
)

set "CAT=%1"
if "%CAT%"=="" set "CAT=all"

REM ai/030 sec.10: local runs use the content-addressed obj store (RTL is not recompiled
REM per case: measured 282s -> 36s on a 15-case bas shard). The compiler's own default is
REM untouched -- this entry point passes the switch explicitly. Set C3_NO_OBJCACHE=1 to
REM reproduce exactly what GA does.
set "INC=-Incremental"
if not "%C3_NO_OBJCACHE%"=="" set "INC="

REM Set C3_TEST_OUTDIR to build/verify into a private directory. The suite otherwise uses
REM the repo default output\, which several people share on this working tree at once --
REM concurrent runs stepping on each other's artifacts produced false reds before.
set "OD="
if not "%C3_TEST_OUTDIR%"=="" set "OD=-OutputDirectory %C3_TEST_OUTDIR%"

if "%2"=="" (
    powershell -ExecutionPolicy Bypass -File "%~dp0..\tests\run_tests.ps1" -Category %CAT% %INC% %OD%
) else (
    powershell -ExecutionPolicy Bypass -File "%~dp0..\tests\run_tests.ps1" -Category %CAT% -Verbose:%2 %INC% %OD%
)
