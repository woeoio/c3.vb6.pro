# check_no_raw_tab_in_ps1.ps1 - 结构性哨兵（第 94 道 [STATIC]）：tests/ 与 scripts/ 下的 .ps1 里不许出现裸制表符 (0x09)。
#
# 为什么要钉这一族（2026-10-21，第一百六十刀「判据」的实物）：
# 给 run_tests.ps1 打补丁的脚本多半是 Python 写的，而那一边用的是双引号字符串 ——
# 在 Python 的双引号串里写反斜杠 t（chr(92) 加字母 t），落到文件里是一枚真制表符 (0x09) 加 "est_..."，不是「反斜杠 + test_...」。
# PowerShell 侧恰好相反：双引号串里的 \t 就是字面的反斜杠+t（路径分隔符 + 一个字母），制表符得写 `t 或 [char]9。
# 于是那一次四条夹具路径的开头都变成「反斜杠被制表符顶掉」：文件不存在，而后果分两种，两种都坏 ——
#   运行期那两例（Add-BasTest）连名字都没出现在 CI 日志里 —— 用例静默消失，判据看着绿其实一条没跑；
#   发码那两例（Test-EmitcShape / Test-EmitcAbsent）报红，但红话是「missing 全部针」与「还在场:（空）」——
#   看着像发码回归，实际是命令行里那个源文件压根没打开。
# 全仓 .ps1 今天 0 个裸制表符（真需要制表符的地方一律走 `t），所以这条判据不设豁免名单，
# 出现即红；红话点名到哪一行，比让下一个人去猜「是发码变了还是路径坏了」便宜得多。
param(
    [string]$Root = ""
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }
$Root = (Resolve-Path $Root).Path

$files = @(Get-ChildItem -LiteralPath (Join-Path $Root "tests") -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue) +
         @(Get-ChildItem -LiteralPath (Join-Path $Root "scripts") -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue)

# 「扫到 0 份」等于哨兵瞎了 —— 先把自己兜住（同 ps51_joinpath 那条自证）。
if ($files.Count -lt 40) {
    Write-Host ("no_raw_tab 自证失败：只扫到 " + $files.Count + " 份 .ps1（预期 >= 40）—— 路径/过滤器坏了")
    exit 2
}

$off = @()
foreach ($f in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $idx = 0
    while ($idx -lt $bytes.Length) {
        $at = [Array]::IndexOf($bytes, [byte]9, $idx)
        if ($at -lt 0) { break }
        $head = [System.Text.Encoding]::UTF8.GetString($bytes, 0, $at)
        $line = 1 + ([regex]::Matches($head, "`n")).Count
        $rel = $f.FullName.Substring($Root.Length + 1) -replace '\\', '/'
        $off += ($rel + ":" + $line + "  裸制表符 (0x09)")
        $idx = $at + 1
    }
}

Write-Host ("no_raw_tab: 扫了 " + $files.Count + " 份 .ps1，裸制表符 " + $off.Count + " 处")
foreach ($o in ($off | Select-Object -First 12)) { Write-Host ("  " + $o) }
if ($off.Count -gt 0) {
    Write-Host '改法：路径里要的是反斜杠加字面 t；真要制表符写反引号 t 或 [char]9。补丁脚本一律用 chr(9)、chr(92) 拼字节，别在 Python 的双引号串里写反斜杠 t。'
    exit 1
}
Write-Host "no_raw_tab OK（没有裸制表符）"
exit 0
