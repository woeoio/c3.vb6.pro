# 账 #236 的结构性哨兵 (只扫夹具源码, 不起 cl 也不跑 exe)
#
# 为什么单开这一道: 门禁里每枚 GUI 夹具都是"起进程 -> 等它自己关掉", 兜底是
# Invoke-TestExe 的 -RunTimeoutSec (默认 60 秒) —— 超时那条路会把进程**杀掉**再判 FAIL,
# 所以症状不是"卡住整个门", 而是"这片里莫名红一条、而且 .out 有内容"。门 #359 就是这样:
# resalpha 那枚夹具里 `If done Then Exit Sub` 走在 `tick = tick + 1` 前面, 第一拍把 done
# 置了之后每一拍都提前返回, 阈值再也够不着 => 窗口永远不关。发码全对、属性全对、
# 两条 needle 也都打出来了 —— 坏的是**夹具自己的收线**。
#
# 规则 (改坏了会红, 不是装饰):
#   F1  扫到的 *_Timer 处理器数 >= 20 (夹具族还在长; 数出 0 说明 glob 或编码坏了, 那比红更糟)
#   F2  处理器里凡是「靠计数阈值收线」的 (If <v> >= <n> Then Unload Me / End / x.Enabled = False),
#       那次自增 <v> = <v> + 1 必须出现在**第一条 Exit Sub 之前** —— 否则任何提前返回都会把
#       计数器冻住, 阈值形同虚设 (账 #236 的实物就是这个形状)
#   F3  F2 的适用面不许是空的: 至少 1 枚处理器是阈值收线形 (针面若哪天全没了, 这道哨兵等于没电)
#
# 用法:  pwsh -File scripts\check_fixture_timer_close.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$thresholdClose = '^\s*If\s+(\w+)\s*(?:>=|>)\s*(\d+)\s+Then\s+(Unload Me|End|\w+\.Enabled = False)\s*$'

$handlers = @()
foreach ($f in (Get-ChildItem -LiteralPath (Join-Path $root "tests") -Recurse -File -Filter *.frm)) {
    $lines = [System.IO.File]::ReadAllText($f.FullName) -split "`r?`n"
    $i = 0
    while ($i -lt $lines.Count) {
        if ($lines[$i] -match '^\s*(?:Private |Public )?Sub\s+(\w+)_Timer\(\)\s*$') {
            $name = $Matches[1] + "_Timer"
            $body = @()
            $j = $i + 1
            while ($j -lt $lines.Count -and $lines[$j] -notmatch '^\s*End Sub\s*$') {
                $body += $lines[$j]
                $j++
            }
            $handlers += [pscustomobject]@{ File = $f.FullName.Substring($root.Length + 1);
                                            Name = $name; Body = $body }
            $i = $j
        }
        $i++
    }
}

# ---- F1: 覆盖面 ----
if ($handlers.Count -lt 20) {
    $bad += ("F1 only " + $handlers.Count + " *_Timer handlers scanned under tests/ (floor 20) ->" +
             " glob or encoding broke, the rule below would be checking nothing")
}

# ---- F2: 阈值收线的处理器, 自增必须在第一条提前返回之前 ----
$thrHandlers = @()
foreach ($h in $handlers) {
    $var = ""
    foreach ($line in $h.Body) {
        if ($line -match $thresholdClose) { $var = $Matches[1]; break }
    }
    if ($var -eq "") { continue }
    $thrHandlers += $h
    $incRe = '^\s*' + [regex]::Escape($var) + '\s*=\s*' + [regex]::Escape($var) + '\s*\+\s*1\s*$'
    $incAt = -1
    $exitAt = -1
    for ($k = 0; $k -lt $h.Body.Count; $k++) {
        $t = $h.Body[$k]
        if ($incAt -lt 0 -and $t -match $incRe) { $incAt = $k }
        if ($exitAt -lt 0 -and $t -match '\bExit\s+Sub\b') { $exitAt = $k }
    }
    if ($incAt -lt 0) {
        $bad += ("F2 " + $h.File + " " + $h.Name + " closes on the threshold of '" + $var +
                 "' but never increments it in this handler")
    } elseif ($exitAt -ge 0 -and $incAt -gt $exitAt) {
        $bad += ("F2 " + $h.File + " " + $h.Name + ": '" + $var + " = " + $var + " + 1' sits at line " +
                 ($incAt + 1) + " but an early Exit Sub comes first (line " + ($exitAt + 1) +
                 ") -> once that return fires the counter freezes and the threshold is unreachable;" +
                 " the exe then only dies to the 60s run-timeout kill")
    }
}

# ---- F3: 适用面非空 ----
if ($thrHandlers.Count -lt 1) {
    $bad += ("F3 no threshold-closing *_Timer handler matched at all (count " + $thrHandlers.Count +
             ") -> the F2 rule is dead, update its shape or the sentinel is decoration")
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS Fixture timer close: handlers " + $handlers.Count +
                " threshold-closing " + $thrHandlers.Count + " violations 0") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
