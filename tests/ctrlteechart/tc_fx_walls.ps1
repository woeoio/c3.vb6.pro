param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = ""
)
# C30-10 第二十刀：walls 判据的薄封装（本体在 tc_fx_px.ps1，按 -Mode 分派；
# run_tests 的 PixelAssert 只传 -Exe，所以按老约定一模式一脚本）。
& (Join-Path $PSScriptRoot 'tc_fx_px.ps1') -Exe $Exe -Mode walls -OutDir $OutDir
exit $LASTEXITCODE