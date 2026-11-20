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
#   P5  C30-9 改口径：**投 quit 的地方只剩一处**（vb6forms.c 里 hand-back 助手）——
#       登记处那条已经换成可撤销的「挂起意头」：
#         · 旧的 2 处（登记处当场 PostQuitMessage + 助手补投）实测有缺陷：`Unload Me` 紧跟
#           `Main.Show`（frmlogin.frm:191-192 登录转场）时那条 WM_QUIT 撤不回来，主窗体刚
#           建好就被队列里那条 quit 判死（退出码 0、窗口闪没）。
#         · 新形状 = Unregister 置 `g_formQuitPending`、Register 一有窗体就清、两个主泵的
#           while 顶部每轮重问 `vb6_Forms_QuitPending()`，真正投 quit 的只有助手。
#       所以这一格现在钉四件事（少一件就会退回那个退出码 0）：
#         P5a 剥注释后 `PostQuitMessage(` 只剩 vb6forms.c 1 处、vb6rtl_system.c 0 处；
#         P5b 意头的三处（QuitPending 读 / RequestQuit 只置位 / Register 清）都在；
#         P5c 两个主泵的 while 都是「先问意头再 GetMessage」，泵后补一条合成 WM_QUIT；
#         P5d 发码侧 WM_DESTROY 调 vb6_Forms_RequestQuit 且**不再**发裸 PostQuitMessage。
#       ⚠ 计数一律先剥注释：本刀在 vb6rtl_system.c 的说明文字里引了旧写法（`PostQuitMessage(0)`），
#       文本计数会把它当成两枚调用点（GA 实测 P5 = 3，全红）—— 修的是尺子的形状，不是把说明改哑。
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

# 剥注释（与 check_com_server_error_frame 的 Strip-Comments185 同形，但多一条字符字面量状
# 态）：字符串字面量原样保留 —— 这道门数的是**C 代码**里的调用点，不是说明文字里的引用。
# 发码那个 .inc 更绕一层：调用点住在 emitLine("...") 的字符串里，所以这里必须把字符串当内容抄
# 下来、而把 // 与 /* */ 当注释丢掉。
#
# 为什么必须认字符字面量 `L'"'`（实测 2026-11-16，这道门第一次跑 P5a 就栽在这）：
#   vb6rtl_system.c:208 `if (*cmdLine == L'"') {` 里的那枚 " 不是字符串的开头 —— 不认字符字面量，
#   扫描器从 208 行起就“在字符串里”，随后 664/681 两个块注释的 /* 与 */ 全被当内容抄下来，
#   P5a 于是读到 vb6rtl_system.c=2（实际是 0，那两处是说明文字）。一行差一个状态，整道门读错。
function Strip-Cmt([string]$text) {
    $out = New-Object System.Text.StringBuilder
    $i = 0
    $n = $text.Length
    $inBlock = $false
    $inLine = $false
    $inStr = $false
    $inChr = $false
    while ($i -lt $n) {
        $c = $text[$i]
        $d = if ($i + 1 -lt $n) { $text[$i + 1] } else { [char]0 }
        if ($inLine) { if ($c -eq "`n") { $inLine = $false; [void]$out.Append($c) }; $i++; continue }
        if ($inBlock) { if ($c -eq "*" -and $d -eq "/") { $inBlock = $false; $i += 2; continue }
                        if ($c -eq "`n") { [void]$out.Append($c) }; $i++; continue }
        if ($inStr -or $inChr) {
            $q = if ($inStr) { '"' } else { "'" }
            [void]$out.Append($c)
            if ($c -eq "\") { if ($i + 1 -lt $n) { [void]$out.Append($text[$i+1]) }; $i += 2; continue }
            if ($c -eq $q) { if ($inStr) { $inStr = $false } else { $inChr = $false } }
            $i++; continue
        }
        if ($c -eq "/" -and $d -eq "*") { $inBlock = $true; $i += 2; continue }
        if ($c -eq "/" -and $d -eq "/") { $inLine = $true; $i += 2; continue }
        if ($c -eq '"') { $inStr = $true; [void]$out.Append($c); $i++; continue }
        if ($c -eq "'") { $inChr = $true; [void]$out.Append($c); $i++; continue }
        [void]$out.Append($c)
        $i++
    }
    return $out.ToString()
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

    # ---- P5 (C30-9): 投 quit 只剩助手一处；登记的退出意头必须**可撤销** ----
    $c1 = Strip-Cmt $t1
    $c2 = Strip-Cmt $t2
    $nPost1 = Count-Of $c1 'PostQuitMessage\('
    $nPost2 = Count-Of $c2 'PostQuitMessage\('
    if ($nPost1 -ne 1 -or $nPost2 -ne 0) {
        $bad += ("P5a PostQuitMessage sites (comments stripped) = vb6forms.c:" + $nPost1 + " vb6rtl_system.c:" + $nPost2 +
                 " (want 1/0: 只有 hand-back 助手真投 quit, 登记处置挂起意头 -> 旧口径 2 处就是退出码 0 那一格)")
    }
    # P5b: 意头的三处主人都在 vb6rtl_system.c，且各自只有一枚。
    $nQp = Count-Of $c2 'int vb6_Forms_QuitPending\(void\)'
    $nRq = Count-Of $c2 'void vb6_Forms_RequestQuit\(void\)'
    if ($nQp -ne 1) { $bad += ("P5b vb6_Forms_QuitPending definition = " + $nQp + " (want 1: 两个主泵每轮问的就是它)") }
    if ($nRq -ne 1) { $bad += ("P5b vb6_Forms_RequestQuit definition = " + $nRq + " (want 1: 发码侧退出意头的唯一出口)") }
    if ($c2 -notmatch 'void vb6_Forms_RequestQuit\(void\)\s*\{\s*g_formQuitPending = 1;\s*\}') {
        $bad += 'P5b vb6_Forms_RequestQuit no longer just raises the pending intent (g_formQuitPending = 1) - 投一条撤不回的 quit 就是 C30-9 那个退出码 0'
    }
    $mReg = [regex]::Match($c2, 'void vb6_Forms_Register\([\s\S]*?\r?\n\}')
    if (-not $mReg.Success) { $bad += 'P5b vb6_Forms_Register body not found' }
    elseif ($mReg.Value -notmatch 'g_formQuitPending\s*=\s*0;') {
        $bad += 'P5b vb6_Forms_Register no longer clears the pending quit - 同一事件里 `Unload Me` 紧跟 `Main.Show` 会被上一发意头判死'
    }
    $mUnreg = [regex]::Match($c2, 'void vb6_Forms_Unregister\([\s\S]*?\r?\n\}')
    if (-not $mUnreg.Success) { $bad += 'P5b vb6_Forms_Unregister body not found' }
    else {
        if ($mUnreg.Value -notmatch 'g_formQuitPending\s*=\s*1;') {
            $bad += 'P5b vb6_Forms_Unregister no longer raises the pending quit when the last form goes - 关窗后没人告诉泵该收工'
        }
        if ($mUnreg.Value -notmatch 'g_msgLoopDepth\s*>\s*0') {
            $bad += 'P5b the `g_msgLoopDepth > 0` gate on the pending quit is gone - 泵没跑时置位无人消费, 反而判死 Sub Main 里 `Load/Unload` 之后的 Show'
        }
    }
    # P5c: 两个主泵的 while 都是「先问意头、再 GetMessage」，泵后补一条合成的 WM_QUIT。
    $nGate = Count-Of $c1 'while \(!vb6_Forms_QuitPending\(\) && GetMessage\(&msg, NULL, 0, 0\)\)'
    if ($nGate -ne 2) {
        $bad += ("P5c main pumps checking the pending quit = " + $nGate + " (want 2: 两条主循环; DoEvents 那条是内层泵, 判意头就把内层当外层了)")
    }
    if ($c1 -notmatch 'if \(vb6_Forms_QuitPending\(\)\) \{ msg\.message = WM_QUIT;') {
        $bad += 'P5c the synthesized WM_QUIT after the pump is gone - 泵因意头而停时 msg 没被赋值过, 栈上垃圾会被当成退出码交回去'
    }
    # P5d: 发码那一侧的形状主人 (启动窗体 WM_DESTROY 只许发 RequestQuit)。
    $genP = Join-Path $root "src\backend\detail\module\cgen_form_wndproc_dispatch.inc"
    if (-not (Test-Path -LiteralPath $genP)) { $bad += ("P5d missing " + $genP) }
    else {
        $cg = Strip-Cmt ([System.IO.File]::ReadAllText($genP))
        $nGenRq = Count-Of $cg 'vb6_Forms_RequestQuit\(\)'
        $nGenPost = Count-Of $cg 'PostQuitMessage\('
        if ($nGenRq -ne 1) {
            $bad += ("P5d emit sites of vb6_Forms_RequestQuit() in cgen_form_wndproc_dispatch.inc = " + $nGenRq + " (want 1: 启动窗体 WM_DESTROY)")
        }
        if ($nGenPost -ne 0) {
            $bad += ("P5d the emit side posts a raw PostQuitMessage again (" + $nGenPost + ") - 那条 quit 撤不回来, 登录转场会被它判死")
        }
    }
    $mMain = [regex]::Match($t1, 'int vb6_MessageLoop\(void\)\s*\{[\s\S]*?\r?\n\}')
    if (-not $mMain.Success) { $bad += "P5 vb6_MessageLoop body not found" }
    elseif ($mMain.Value -match 'vb6_RePostQuitIfTaken') {
        $bad += "P5 the OUTERMOST pump hands the quit back - it is the consumer; re-posting there would loop forever"
    }
}

if ($bad.Count -eq 0) {
    Write-Host "PASS quit-pump invariant: 4 pumps, 2 inner hand-backs, 1 raw quit (helper), pending-intent cancellable, outer pump consumes" -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
