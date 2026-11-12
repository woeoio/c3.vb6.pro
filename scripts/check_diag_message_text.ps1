# check_diag_message_text.ps1 - 诊断消息文本的形状哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: FIX-010b/c 改的是**诊断消息**本身 —— 它不进产物, 所以形状门
# (只哈希 --emit-c 的 stdout) 根本看不见它; 而这两条消息恰恰是用户唯一能看见的
# 那一面: TypeLib 加载失败时打的是**哪条路径**与**哪个 HRESULT**, rc.exe 找不到时
# 告诉不告诉用户**怎么补**。没有针 => 下一个人改回十进制 / 打回未展开的路径,
# 没人会发现 (本地与 CI 上这条消息都没有稳定的触发路径: 本机有 rc.exe, CI 的
# VS 镜像也带 Windows SDK => 两边都打不出这条消息, 端到端针不可能存在)。
#
# 规则 (改坏了会红, 不是装饰):
#   R1  TypeLib 加载失败必须打**展开后**的路径 (wideToAcpNarrow(tlbPathW));
#       不许打注册表原值 (那种带 %CommonProgramFiles% 的 %VAR%) —— 用户会读成
#       "路径写错", 而实际错因常是 32/64 位不匹配。
#   R2  HRESULT 必须打 8 位十六进制 ("0x%08X"); 不许 std::to_string(hr)
#       (那是十进制带负号: 0x-2147312566, 读不出是哪个错误码)。
#   R3  三条 rc.exe 诊断每条都要说**怎么补** (Windows SDK / PATH)。
#   R4  rc.exe 的发现仍只有一处 (findTool("rc")) —— 消息不许各写各的。
#
# 用法:  powershell -File scripts\check_diag_message_text.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$tl = Join-Path $root "src\com\typelib_parser.cpp"
$dl = Join-Path $root "src\driver\driver_link.cpp"
foreach ($f in @($tl, $dl)) {
    if (-not (Test-Path -LiteralPath $f)) { $bad += ("M0 missing file: " + $f) }
}

if ($bad.Count -eq 0) {
    $t = [System.IO.File]::ReadAllText($tl)
    $d = [System.IO.File]::ReadAllText($dl)

    # --- R1: TypeLib 失败消息必须打展开后的路径 ---
    $m = [regex]::Match($t, '"Failed to load TypeLib: "')
    if (-not $m.Success) {
        $bad += 'R1 the "Failed to load TypeLib" diagnostic is gone'
    } else {
        $seg = $t.Substring($m.Index, [Math]::Min(320, $t.Length - $m.Index))
        if ($seg -notmatch 'wideToAcpNarrow\(tlbPathW\)') {
            $bad += 'R1 the TypeLib failure message does not print the EXPANDED path (wideToAcpNarrow(tlbPathW))'
        }
        if ($seg -match '"Failed to load TypeLib: "\s*\+\s*tlbPath\b') {
            $bad += 'R1 the TypeLib failure message still prints the raw registry value (tlbPath)'
        }
    }

    # --- R2: HRESULT 必须 8 位十六进制 ---
    if ($t -match 'std::to_string\(hr\)') {
        $bad += 'R2 the HRESULT is still rendered in decimal (std::to_string(hr)) - use 0x%08X'
    }
    if ($t -notmatch '0x%08X') {
        $bad += 'R2 no 8-digit hex HRESULT rendering (0x%08X) found'
    }

    # --- R3: 三条 rc.exe 诊断每条都要说怎么补 ---
    $rcLines = @([regex]::Matches($d, '"C3: rc\.exe not found[^\n]*'))
    if ($rcLines.Count -ne 3) {
        $bad += ("R3 expected 3 rc.exe diagnostics, found " + $rcLines.Count)
    }
    foreach ($one in $rcLines) {
        $s = $one.Value
        if (($s -notmatch 'Windows SDK') -or ($s -notmatch 'PATH')) {
            $bad += ('R3 an rc.exe diagnostic does not say how to fix it: ' + $s)
        }
    }

    # --- R4: rc.exe 的发现仍只有一处 ---
    $finds = @([regex]::Matches($d, 'findTool\("rc"\)'))
    if ($finds.Count -ne 1) {
        $bad += ('R4 rc.exe discovery is not in one place (findTool("rc") x' + $finds.Count + ')')
    }
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS diagnostic-message shapes: expanded TypeLib path + 8-hex HRESULT + rc.exe fix hint") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
