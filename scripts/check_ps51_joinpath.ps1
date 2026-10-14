# check_ps51_joinpath.ps1 - 结构性哨兵（第 36 道 [STATIC]）：全仓 .ps1 里不许出现「三个以上位置参数」的 Join-Path。
#
# 为什么要钉这一族（2026-10-08，从 origin/ferock/0.10.7 那笔 5735aff6 捞回来的真修）：
# Windows PowerShell 5.1 的 Join-Path 只有两个位置参数（-Path / -ChildPath），第三段必须走 -AdditionalChildPath；
# 把三段直接写成位置参数，在 5.1 下抛 ParameterBindingException —— 而它通常被塞在双引号内插 "$(...)" 里，
# 于是**只剩一行噪声、那一段静默变空**。本机实测（改前的 tests/run_tests.ps1::Get-MsvcToolset）：
#   5.1   IncludeSegs=1  LibSegs=0  EmptySegs=8   Include=[...MSVC\14.29.30133\include;;;;;]
#   pwsh7 IncludeSegs=6  LibSegs=3  EmptySegs=0
# 后果就是 ferock 那条 commit message 写的：INCLUDE/LIB 里 SDK 段全丢 ⇒ cl.exe C1083 打不开 stddef.h/windows.h
# ⇒ 谁用 5.1 跑回归，编译组整片假红。CI 与日常都用 pwsh 7，所以这缺陷**只在"有人用 5.1"那一刻现形** ——
# 正是只能靠结构判据钉住的那一族（真跑覆盖面不到）。改法统一：显式内插 "$dir\$child\$leaf"，两壳同解。

param(
    [string]$Root = ""
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }
$Root = (Resolve-Path $Root).Path

$files = @(Get-ChildItem -LiteralPath (Join-Path $Root "tests") -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue) +
         @(Get-ChildItem -LiteralPath (Join-Path $Root "scripts") -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue)

# 「扫到 0 份」等于哨兵瞎了 —— 这类判据的空过形状在这本账里出现过不止一次，先把自己兜住。
if ($files.Count -lt 40) {
    Write-Host ("ps51_joinpath 自证失败：只扫到 " + $files.Count + " 份 .ps1（预期 >= 40）—— 路径/过滤器坏了")
    exit 2
}

$off = @()
foreach ($f in $files) {
    $parseErrs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$parseErrs)
    if ($parseErrs -and $parseErrs.Count -gt 0) {
        $off += ("解析失败 " + $f.Name + " : " + $parseErrs[0].Message)
        continue
    }
    $cmds = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
    foreach ($c in $cmds) {
        if ($c.GetCommandName() -ne 'Join-Path') { continue }
        $pos = 0
        $pendingName = $false
        for ($k = 1; $k -lt $c.CommandElements.Count; $k++) {   # 0 = 命令名本身
            $e = $c.CommandElements[$k]
            if ($e -is [System.Management.Automation.Language.CommandParameterAst]) {
                $pendingName = $true
                continue
            }
            if ($pendingName) { $pendingName = $false; continue }   # 具名参数的值，不算位置参数
            $pos++
        }
        if ($pos -gt 2) {
            $rel = $f.FullName.Substring($Root.Length + 1) -replace '\\', '/'
            $txt = $c.Extent.Text
            if ($txt.Length -gt 88) { $txt = $txt.Substring(0, 88) + '...' }
            $off += ($rel + ":" + $c.Extent.StartLineNumber + "  位置参数 " + $pos + " 段 -> " + $txt)
        }
    }
}

Write-Host ("ps51_joinpath: 扫了 " + $files.Count + " 份 .ps1，违规 " + $off.Count + " 处")
foreach ($o in ($off | Select-Object -First 12)) { Write-Host ("  " + $o) }
if ($off.Count -gt 0) {
    Write-Host '改法：写成显式内插 "$dir\$child\$leaf"（5.1 与 pwsh 7 同解）；别改成 -AdditionalChildPath —— 那是 7 才有的名字。'
    exit 1
}
Write-Host "ps51_joinpath OK（没有三参数以上的 Join-Path）"
exit 0
