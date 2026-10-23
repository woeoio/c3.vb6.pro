# check_owned_temp_semantics.ps1 - 账 #463 §B164：VB 那一面的**参考实现**（VBScript）跑一遍
#
# 为什么要有这一道：回收刀要动的那一族「语句里现造的带权临时量」，在 **VB 家族那一面**
#   是什么？只有三问，而它们必须由**参考实现**回答，不能由编译器自己的推理想当然：
#     ① 它是**值**还是**别名**？      ⇒ 「现造」这个说法站不站得住
#     ② 它是**什么档**？              ⇒ C 侧该声明成什么档
#     ③ 赋值是不是**拷贝**？          ⇒ 生命周期（语句级）站不站得住
#   参考件 = scripts/ref_b164_semantics.vbs（**纯 ASCII**：cscript 按机器 ANSI 码页读 .vbs，
#   UTF-8 的中文注释在 zh-CN 机器上会被 GBK 解码吃掉后面的 ASCII ⇒ 下一行凭空消失；
#   在 CI 机器（cp1252）上又是另一副样子 ⇒ 纯 ASCII 才是「每台机器读到的字节一样」）。
#
# ⚠ 这一道**只答 VB 那一面**。**谁负责放掉**（C 层的 free/clear）是**另一面**，
#   那一面的权威 = src/rtl/core/vb6rtl/*.c 的实读（scripts/check_owned_temp_keeps.ps1 推导，
#   落在 scripts/owned_temp_keeps_164.txt 的 [consumes] 节）。两面分开记：
#   不让「VB 语义」替「C 所有权」背书。
#
# 跑法: powershell -File scripts/check_owned_temp_semantics.ps1
param(
    [string]$Root = ""
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$Root = (Resolve-Path $Root).Path
$vbs = Join-Path $Root 'scripts\ref_b164_semantics.vbs'

$fail = @()
if (-not (Test-Path -LiteralPath $vbs)) {
    Write-Host ("MISSING " + $vbs)
    exit 2
}
# ① 纯 ASCII（每台机器读到的字节一样 ⇒ 这一道的读数跨机器可比）
$bytes = [System.IO.File]::ReadAllBytes($vbs)
$nonAscii = @($bytes | Where-Object { $_ -gt 127 })
if ($nonAscii.Count -gt 0) {
    $fail += ("参考件里有 " + $nonAscii.Count + " 个非 ASCII 字节（cscript 按机器 ANSI 码页读 .vbs ⇒ 跨机器读数会不一样）")
}
# ② 跑参考实现
$cscript = Get-Command cscript.exe -ErrorAction SilentlyContinue
if (-not $cscript) {
    # 不许静默绿：这一格压根没判 ⇒ 写进读数里，并按红算
    Write-Host "VBS-SEMANTICS NO-CSCRIPT (找不到 cscript.exe ⇒ 这一格没判)"
    exit 3
}
$outFile = Join-Path ([System.IO.Path]::GetTempPath()) ("vb164_" + [guid]::NewGuid().ToString("N") + ".txt")
$errFile = [System.IO.Path]::ChangeExtension($outFile, ".err")
$p = Start-Process -FilePath $cscript.Source -ArgumentList @('//NoLogo', $vbs) -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput $outFile -RedirectStandardError $errFile
$out = ''
if (Test-Path -LiteralPath $outFile) { $out = (Get-Content -LiteralPath $outFile -Raw -ErrorAction SilentlyContinue) }
$pass = @([regex]::Matches($out, '(?m)^PASS ')).Count
$bad = @([regex]::Matches($out, '(?m)^FAIL ')).Count
if ($p.ExitCode -ne 0) { $fail += ("参考实现退出码 " + $p.ExitCode + "（0 = 全过）") }
if ($bad -gt 0) { $fail += ("参考实现里有 " + $bad + " 条 FAIL") }
if ($out -notmatch 'ALL-PASS') { $fail += '参考实现没有 ALL-PASS 那一行' }
# ④ **求值次序**那三条必须在（它们是回收刀那几道次序门的**外部权威**：
#    "&" 的两个操作数左到右、调用实参左到右、链式拼接左到右）
#    —— 少了任何一条 ⇒ 那些门就只剩 C 侧自己的假设 ⇒ 钉住。
#    账 #473 再补 `+` 当拼接那一站的三条（值 / 两操作数次序 / 链式次序）：那一站
#    现在也回收右操作数 ⇒ 同样要外部权威证明「值不动、次序左到右」。
foreach ($needle in @('concat-operand-order', 'call-argument-order', 'concat-chain-order',
                      'add-concat-value', 'add-concat-operand-order', 'add-concat-chain-order')) {
    if ($out -notmatch ('(?m)^PASS ' + [regex]::Escape($needle) + ' ')) {
        $fail += ("参考实现里少了一条求值次序的读数：" + $needle + "（次序门失去外部权威）")
    }
}
Remove-Item -LiteralPath $outFile, $errFile -Force -ErrorAction SilentlyContinue

if ($fail.Count -gt 0) {
    Write-Host "VBS-SEMANTICS FAIL"
    foreach ($f in $fail) { Write-Host ("  - " + $f) }
    exit 1
}
Write-Host ("VBS-SEMANTICS OK (参考件纯 ASCII / cscript 退出码 0 / PASS " + $pass +
            " 条 / FAIL 0 条 / ALL-PASS 在)")
exit 0
