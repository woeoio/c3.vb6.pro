param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# C30-10 第二十刀：fx 判据的薄封装（本体在 tc_fx_px.ps1，按 -Mode 分派；
# run_tests 的 PixelAssert 只传 -Exe，所以按老约定一模式一脚本）。
& (Join-Path $PSScriptRoot 'tc_fx_px.ps1') -Exe $Exe -Mode fx -OutDir $OutDir -Sec $Sec
exit $LASTEXITCODE