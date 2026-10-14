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
# 账 #213 (第三十刀) 在同一道哨兵里加的是**另一问**: 「post 出去的消息被泵异步消化,
# 什么时候才许读它的效果?」门 #332 (pbsub 五条通知读成全 0) 与 #253 (modal hops 3->1)
# 都是这一问在 CI 上按负载红的实物, 两台本地编译器却稳定 —— 因为本地一拍之内就消化完了。
# tabwalk 相 3 是同一族的第三处: 三步方向键每步只留**恰好一拍**余量去读落点。
# 修法不是逐处补 sleep, 是把「等到效果真的出现再读」写成一条能被数的口径:
#   W1  发过 PostMessage 的 *_Timer 处理器 = **一张登记名册** (现 4 枚, 逐枚对上);
#       新冒出发 post 的处理器却没登记 = 红, 登记了却扫不到 = 红 (glob/编码坏掉时同样是红)
#   W2  名册里凡「post 之后还要读效果」的那几枚, 必须带着它那一族的**等待口径**, 且等待行
#       在被保护的那次**读取之前** (顺序错了等于没等: modal 的间隔闸 / pbsub 的到齐就读 /
#       tabwalk 的落点轮询各一条, 逐枚按字面认)
#   W3  名册里刻意**不等**的那一枚 (ModalDlg 只 post WM_CLOSE、之后什么都不读) 必须仍然
#       「post 是处理器最后一条语句」—— 一旦有人在它后面补一句读法, 它就落进 W2 的适用面
#   W4  tabwalk 那一枚的形状细节: 每步只发一对键 (`Call PostMessage(GetFocus(), WM_KEYDOWN`
#       恰 1 处)、每步只自增一次 (`gAkStep = gAkStep + 1` 恰 1 处)、保险丝是**有限**的
#       (`Private Const AK_WAIT_MAX As Long = 40` 逐字在), 而且那条证人**有人读** ——
#       夹具打 `TW-SETTLE late=`、runner 钉 `TW-SETTLE late=0`, 两头缺一即红
#       (账 #226 那一课: 只打不判的读数等于没打)
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

# ---- W1: 发过 post 的处理器是一张名册, 逐枚登记它「等到效果出现」的口径 ----
# 名册 = 这一问唯一的权威。第三种形 (新加一枚往 Timer 里 post 消息、却按固定拍号读结果) 不许
# 悄悄进来: 没登记就红。登记的那一句被删掉、或被挪到读取**之后** (等于没等) 也红。
$reg = @(
    [pscustomobject]@{ Rel = "tests\modal\ModalDlg.frm";   Name = "t2_Timer";
                       Settle = ""; Read = ""; NoRead = "WhereIs(" },
    [pscustomobject]@{ Rel = "tests\modal\ModalMain.frm";  Name = "tMain_Timer";
                       Settle = "(Timer - gLastBeat) < PHASE_MIN_SEC"; Read = "st = WhereIs(f)"; NoRead = "" },
    [pscustomobject]@{ Rel = "tests\pbsub\PbForm.frm";     Name = "tmrStep_Timer";
                       Settle = "ElseIf allIn Or mStep >= 20 Then"; Read = 'Debug.Print "PB-CNT'; NoRead = "" },
    [pscustomobject]@{ Rel = "tests\tabwalk\WalkForm.frm"; Name = "tWalk_Timer";
                       Settle = "If gAkWait < AK_WAIT_MAX Then Exit Sub"; Read = "st = WhereIs(GetFocus())"; NoRead = "" }
)

function Posting-Handlers([array]$hs) {
    $r = @()
    foreach ($h in $hs) {
        if (@($h.Body | Where-Object { $_.Contains("PostMessage(") }).Count -gt 0) { $r += $h }
    }
    return $r
}
function First-Line([array]$body, [string]$lit) {
    for ($k = 0; $k -lt $body.Count; $k++) { if ($body[$k].Contains($lit)) { return $k } }
    return -1
}

$posting = @(Posting-Handlers $handlers)
$regKeys = @($reg | ForEach-Object { $_.Rel + "::" + $_.Name })
$postKeys = @($posting | ForEach-Object { $_.File + "::" + $_.Name })
foreach ($k in $postKeys) {
    if ($regKeys -notcontains $k) {
        $bad += ("W1 a *_Timer handler posts messages but is not on the register: " + $k +
                 " -> add its settle caliber to `$reg` in this script, do not read a posted" +
                 " effect on a fixed beat count")
    }
}
foreach ($g in $reg) {
    $key = $g.Rel + "::" + $g.Name
    if ($postKeys -notcontains $key) {
        $bad += ("W1 registered handler no longer posts (or vanished): " + $key)
        continue
    }
    $h = @($posting | Where-Object { ($_.File + "::" + $_.Name) -eq $key })[0]
    if ($g.Settle -ne "") {
        $sAt = First-Line $h.Body $g.Settle
        $rAt = First-Line $h.Body $g.Read
        if ($sAt -lt 0) { $bad += ("W2 " + $key + " lost its settle line: '" + $g.Settle + "'") }
        if ($rAt -lt 0) { $bad += ("W2 " + $key + " no longer contains the read it guards: '" + $g.Read + "'") }
        if ($sAt -ge 0 -and $rAt -ge 0 -and $sAt -gt $rAt) {
            $bad += ("W2 " + $key + ": the settle test sits at line " + ($sAt + 1) + " but the effect is" +
                     " read earlier (line " + ($rAt + 1) + ") -> the wait guards nothing")
        }
    } elseif ($g.NoRead -ne "") {
        # 登记为「post 完不读任何效果」的那一枚：一旦它开始读，就必须换一种形登记
        if (@($h.Body | Where-Object { $_.Contains($g.NoRead) }).Count -gt 0) {
            $bad += ("W2 " + $key + " is registered as post-without-readback but its body now reads" +
                     " window state ('" + $g.NoRead + "') -> register a settle caliber for it")
        }
    }
}

# ---- W3: tabwalk 那一枚的形状 —— 每步一发、每步一自增、保险丝有限 ----
$twReg = @($reg | Where-Object { $_.Name -eq "tWalk_Timer" })
if ($twReg.Count -eq 1) {
    $twH = @($posting | Where-Object { ($_.File + "::" + $_.Name) -eq ($twReg[0].Rel + "::tWalk_Timer") })
    if ($twH.Count -eq 1) {
        foreach ($pin in @(@("Call PostMessage(GetFocus(), WM_KEYDOWN", 1), @("gAkStep = gAkStep + 1", 1),
                           @("If gAkStep < 4 Then", 1), @("If gAkStep < 3 Then", 1))) {
            $n = @($twH[0].Body | Where-Object { $_.Contains($pin[0]) }).Count
            if ($n -ne $pin[1]) {
                $bad += ("W3 tWalk_Timer: '" + $pin[0] + "' appears " + $n + " times (want exactly " +
                         $pin[1] + ") -> a step that posts twice skips a station, a step that advances" +
                         " twice reads a key nobody posted")
            }
        }
    }
    $twSrc = ""
    if (Test-Path -LiteralPath (Join-Path $root "tests\tabwalk\WalkForm.frm")) {
        $twSrc = [System.IO.File]::ReadAllText((Join-Path $root "tests\tabwalk\WalkForm.frm"))
    }
    $m = [regex]::Match($twSrc, 'Private Const AK_WAIT_MAX As Long = (\d+)')
    if (-not $m.Success) {
        $bad += "W3 WalkForm.frm has no finite fuse constant (AK_WAIT_MAX) -> an unwaited settle can hang the fixture to the 60s run-timeout kill"
    } else {
        $bound = [int]$m.Groups[1].Value
        # 3 keys x bound beats x 50ms must stay well inside Invoke-TestExe's 60s budget
        if ($bound -lt 5 -or $bound -gt 100) {
            $bad += ("W3 AK_WAIT_MAX = " + $bound + " (want 5..100): below 5 the settle can give up" +
                     " before the pump ever digests a posted key, above 100 the worst case" +
                     " (" + $bound + " beats x 3 steps x 50ms) eats the 60s run timeout")
        }
    }
}

# ---- W4: 证人打了就得有人读 (账 #226: 只打不判的读数等于没打) ----
$twFile = Join-Path $root "tests\tabwalk\WalkForm.frm"
$runner = Join-Path $root "tests\run_tests.ps1"
if (Test-Path -LiteralPath $twFile) {
    $fx = [System.IO.File]::ReadAllText($twFile)
    $rn = ""
    if (Test-Path -LiteralPath $runner) { $rn = [System.IO.File]::ReadAllText($runner) }
    $prints = $fx.Contains('Log1 "TW-SETTLE late=')
    $pins = $rn.Contains('"TW-SETTLE late=0"')
    if ($prints -and -not $pins) {
        $bad += "W4 WalkForm prints TW-SETTLE late= but run_tests.ps1 never pins it -> the settle witness is decoration"
    }
    if ($pins -and -not $prints) {
        $bad += "W4 run_tests.ps1 pins TW-SETTLE late=0 but WalkForm stopped printing it -> that needle can never pass"
    }
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS Fixture timer close: handlers " + $handlers.Count +
                " threshold-closing " + $thrHandlers.Count + " posting " + $posting.Count +
                " registered " + $reg.Count + " violations 0") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
