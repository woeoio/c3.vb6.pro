param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# C30-10 第十八刀：donut 判据的薄封装（本体在 tc_type_px.ps1，按 -Mode 分派）。
# 拆成独立脚本是为了跟 run_tests.ps1 的 PixelAssert 约定一致（它只传 -Exe）。
& (Join-Path $PSScriptRoot 'tc_type_px.ps1') -Exe $Exe -Mode donut -OutDir $OutDir -Sec $Sec
exit $LASTEXITCODE