# check_quit_pump_invariant.ps1 - 账 #259 的结构性哨兵（只扫 RTL 源码，不起 cl）
#
# 为什么单开这一道：VB6 里最后一个窗体卸载 ⇒ 进程结束。本仓的做法是 vb6_Forms_Unregister 在
# g_formCount 归零时 PostQuitMessage 一条 WM_QUIT，而**一个线程只有一份 quit**：谁把它抽出来
# 谁就负责投回去。DoEvents 与模态 Show 的循环都是内层泵，抄走它的后果是外层主循环永远等不到
# 退出信号 —— 窗口一个不剩、进程驻留。门 #395 唯一红就是这个形状在 CI 上的样子（combofocus_x86
# 七行判据全打完、60s 不退、CPU 62ms）；饿机器上是概率性，本地把形状做硬之后是每次。
#
# 规则（改坏了会红，不是装饰）：
#   P1  vb6forms.c 里的泵数 = 4：主循环两条（OLE 探针那条 + 正常那条）+ 模态一条 + DoEvents 一条。
#       数字变了必须来说清哪条是外层、哪条是内层。
#   P2  vb6_RePostQuitIfTaken 提及 = 3（定义 + **恰好两个**内层泵各调一次）。
#   P3  DoEvents 那条泵：抽到 WM_QUIT ⇒ 先 hand-back、再 break，中间不许先 Translate/Dispatch。
#       扫窗懒配到第一个**顶格 }**、判据前**先剥注释** —— 那条臂里现在写着第二条规则的说明，
#       定长窗口或不剥注释都会把「别人合理加的那段」读成假红（同族坑见 §C）。
#   P6  同一条 quit 被抽到**第二次**（`>= 2`）⇒ vb6_End()：调用者停在 DoEvents 忙等里没有外层泵时，
#       VB6 的口径是「最后一个窗体卸载 ⇒ 程序结束」；`C3_NO_QUIT_IN_DOEVENTS` 是退回纯投递的开关。
#   P4  模态那条 while 的收尾紧跟一条 hand-back。
#   P5  PostQuitMessage 在 vb6forms.c / vb6rtl_system.c 里 = 2（登记处那一条投 + 助手那一条补投），
#       并且**主循环那条泵里没有** hand-back —— 外层是消费者，给它加投回等于把退出信号弄丢两次。
#
# 用法:   powershell -File scripts\check_quit_pump_invariant.ps1
# 退出码: 0 = 全绿；1 = 红

param()

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()
$f1 = Join-Path $root "src\rtl\core\vb6forms\vb6forms.c"
$f2 = Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_system.c"
if (-not (Test-Path -LiteralPath $f1)) { $bad += ("missing " + $f1) }
if (-not (Test-Path -LiteralPath $f2)) { $bad += ("missing " + $f2) }

$t1 = ""
$t2 = ""
if ($bad.Count -eq 0) {
    $t1 = [System.IO.File]::ReadAllText($f1)
    $t2 = [System.IO.File]::ReadAllText($f2)
}

function Count-Of([string]$text, [string]$pattern) {
    return @([regex]::Matches($text, $pattern)).Count
}

if ($bad.Count -eq 0) {
    # ---- P1: 泵数 ----
    $nGet = Count-Of $t1 'GetMessage\(&msg, NULL, 0, 0\)'
    $nPeek = Count-Of $t1 'PeekMessage\(&msg, NULL, 0, 0, PM_REMOVE\)'
    if ($nGet -ne 3) { $bad += ("P1 GetMessage pumps in vb6forms.c = " + $nGet + " (want 3: 两条主循环 + 模态一条)") }
    if ($nPeek -ne 1) { $bad += ("P1 PeekMessage pumps in vb6forms.c = " + $nPeek + " (want 1: DoEvents)") }

    # ---- P2: hand-back 助手 ----
    $nRe = Count-Of $t1 'vb6_RePostQuitIfTaken\s*\('
    if ($nRe -ne 3) { $bad += ("P2 vb6_RePostQuitIfTaken mentioned " + $nRe + " times (want definition + exactly 2 inner pumps)") }

    # ---- P3: DoEvents 那条内层泵 —— 抽到 quit 就投回去并停泵 ----
    # 扫窗懒配到**第一个顶格 }**（定长窗口会把别人后来加进那条臂的分支截掉 = 假红，
    # 与 check_control_dc 那条同族坑同一形状）；判据前先剥注释（那条臂里本来就写着中文说明）。
    $mDE = [regex]::Match($t1, 'int vb6_DoEvents\(void\)\s*\{[\s\S]*?\r?\n\}')
    if (-not $mDE.Success) { $bad += 'P3 vb6_DoEvents body not found' }
    else {
        $codeDE = ($mDE.Value -replace '//[^\r\n]*', '') -replace '/\*[\s\S]*?\*/', ''
        $iQ = $codeDE.IndexOf('msg.message == WM_QUIT')
        if ($iQ -lt 0) { $bad += 'P3 DoEvents no longer has a WM_QUIT guard at all' }
        else {
            $after = $codeDE.Substring($iQ)
            $iRe = $after.IndexOf('vb6_RePostQuitIfTaken(&msg)')
            $iBr = $after.IndexOf('break;')
            $iTr = $after.IndexOf('TranslateMessage')
            if ($iRe -lt 0 -or $iBr -lt 0 -or $iRe -gt $iBr) {
                $bad += 'P3 DoEvents stopped handing the quit back before it breaks (a swallowed quit = the process never exits)'
            }
            if ($iTr -ge 0 -and ($iTr -lt $iRe -or $iTr -lt $iBr)) {
                $bad += 'P3 DoEvents dispatches messages before dealing with the quit (the quit gets eaten by an inner pump)'
            }
            # ---- P6: 同一条 quit 被抽到第二次 = 没有外层泵在等它 ⇒ 程序结束（3DMenu 那轮定的口径）----
            $nEnd = @([regex]::Matches($codeDE, 'vb6_End\s*\(\s*\)')).Count
            if ($nEnd -ne 1) {
                $bad += ('P6 DoEvents calls vb6_End() ' + $nEnd + ' times on the second take (want 1: a caller stuck in a DoEvents busy-wait has no outer pump to exit through)')
            }
            if ($codeDE -notmatch '>=\s*2') {
                $bad += 'P6 the second-take rule no longer counts takes (>= 2) - the first take must still just break'
            }
            if ($codeDE -notmatch 'C3_NO_QUIT_IN_DOEVENTS') {
                $bad += 'P6 the second-take exit lost its escape hatch C3_NO_QUIT_IN_DOEVENTS (the A/B switch the 259 contract needs)'
            }
        }
    }

    # ---- P4: 模态那条循环的收尾 ----
    $mStart = $t1.IndexOf('while (IsWindow((HWND)hwnd) && GetMessage(&msg, NULL, 0, 0))')
    $mTail = $t1.IndexOf('模态循环退出', $mStart)
    if ($mStart -lt 0 -or $mTail -lt 0) { $bad += 'P4 modal loop or its exit trace not found' }
    else {
        $seg = $t1.Substring($mStart, $mTail - $mStart)
        $n4 = @([regex]::Matches($seg, 'vb6_RePostQuitIfTaken\(&msg\);')).Count
        if ($n4 -ne 1) {
            $bad += ('P4 the modal loop hands the quit back ' + $n4 + ' times between the loop and its exit trace (want 1)')
        }
    }

    # ---- P5: 投 quit 的地方只有两处，且外层泵不投 ----
    $nPost1 = Count-Of $t1 'PostQuitMessage\('
    $nPost2 = Count-Of $t2 'PostQuitMessage\('
    if ($nPost1 + $nPost2 -ne 2) {
        $bad += ("P5 PostQuitMessage sites = " + ($nPost1 + $nPost2) + " (want 2: 登记处投 + 助手补投) -> vb6forms.c=" + $nPost1 + " vb6rtl_system.c=" + $nPost2)
    }
    $mMain = [regex]::Match($t1, 'int vb6_MessageLoop\(void\)\s*\{[\s\S]*?\r?\n\}')
    if (-not $mMain.Success) { $bad += "P5 vb6_MessageLoop body not found" }
    elseif ($mMain.Value -match 'vb6_RePostQuitIfTaken') {
        $bad += "P5 the OUTERMOST pump hands the quit back - it is the consumer; re-posting there would loop forever"
    }
}

if ($bad.Count -eq 0) {
    Write-Host "PASS quit-pump invariant: 4 pumps, 2 inner hand-backs, outer pump consumes" -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
