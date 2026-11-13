# check_mouse_enter_index.ps1 - 账 #610 第345刀的结构哨兵
#
# 判据: MouseEnter / MouseLeave 的**数组档**派发必须与处理器的**定义侧**同源。
#
# 为什么这道门非有不可:
#   形状门 (只哈希 --emit-c 的 stdout) 看得见"发码变了"，但**看不见**"定义与派发是否
#   自洽"。本刀修的正是这种自洽性:
#     修前 定义 static void ..._MouseEnter(int16_t* INDEX);   ← ByRef 指针 (VB6 默认)
#          派发 { extern void ..._MouseEnter(int16_t); ..._MouseEnter(3); }  ← 传值 + 编译期常量
#   两种形参表 ⇒ MSVC 只给 C4133 警告、链接照过 ⇒ **运行期**把整数 3 当指针解引用 ⇒ 崩。
#   这类错位在形状门眼里是"一次正常的发码变化"，只有逐字钉住派发形态才抓得到。
#
# 口径:
#   R1 派发 lambda 在 (改名 evCallBlockUcIndex)
#   R2 传值/传址的判据必须走权威 eventParamByAddress, 且 isUserControl 逐字为 false
#      (问的是**过程那一路**: 处理器定义读 isByVal, 与 makeParamCType 同一答案)
#   R3 数组下标必须从 GetPropW(hwnd, "VB6_CtrlArrIndex") - 1 读 (运行时下标),
#      不得再用编译期常量 info.arrayIndex (40 个 Label 一律答 39)
#   R4 派发 lambda 体内不得出现 std::to_string(info.arrayIndex)
#   R5 旧名 evCallBlockUcByVal 不得残留
#   R6 传址/传值两档都要在源码里 (byAddr ? "int16_t*" : "int16_t")
#   R7 宿主 cpp 必须 include 权威头 common/event_param_abi.hpp
#
# 用法: powershell -File scripts/check_mouse_enter_index.ps1   (PASS => exit 0)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$incPath = Join-Path $Root 'src/backend/detail/module/cgen_form_wndproc_subclass.inc'
$cppPath = Join-Path $Root 'src/backend/cgen_util_ctrl.cpp'

$problems = @()

if (-not (Test-Path $incPath)) {
    Write-Host ("FAIL A: 找不到 " + $incPath)
    exit 1
}
$src = Get-Content $incPath -Raw

# ---- R1: 派发 lambda 的名字 ----
if (-not $src.Contains('evCallBlockUcIndex')) {
    $problems += "R1: 找不到 evCallBlockUcIndex (MouseEnter/MouseLeave 的数组档派发 lambda)"
}
# 两处调用点都要用新名
foreach ($suffix in @('"_MouseEnter"', '"_MouseLeave"')) {
    $hit = $src.Contains('evCallBlockUcIndex(' + $suffix + ')')
    if (-not $hit) {
        $problems += "R1: 没有 evCallBlockUcIndex(" + $suffix + ") 的调用点"
    }
}

# ---- R2: 判据必须走权威 eventParamByAddress, isUserControl 逐字 false ----
if (-not $src.Contains('eventParamByAddress(')) {
    $problems += "R2: 派发侧没有调权威 eventParamByAddress => 传值/传址各答一套 ⇒ 与定义侧可能不同源"
}
if (-not $src.Contains('/*isUserControl=*/false')) {
    $problems += "R2: eventParamByAddress 的 isUserControl 实参必须逐字是 /*isUserControl=*/false (过程那一路)"
}

# ---- R4: 派发 lambda 体内不得用编译期常量下标 ----
$lines = @($src -split "`r?`n")
$lamStart = -1
$lamEnd = -1
for ($k = 0; $k -lt $lines.Count; $k++) {
    $hit = $lines[$k].Contains('auto evCallBlockUcIndex = [&](const char* suffix)')
    if ($hit -and $lamStart -lt 0) { $lamStart = $k }
}
if ($lamStart -ge 0) {
    for ($k = $lamStart; $k -lt $lines.Count; $k++) {
        $hit = $lines[$k].Contains('};')
        if ($hit) { $lamEnd = $k; break }
    }
}
if ($lamStart -lt 0) {
    $problems += "R4: 定位不到 evCallBlockUcIndex 的 lambda 头 (R1 已经报过)"
} else {
    if ($lamEnd -lt 0) { $lamEnd = $lines.Count }
    $body = ''
    for ($k = $lamStart; $k -le $lamEnd; $k++) { $body += $lines[$k] + "`n" }
    # ---- R3: 数组下标必须是**运行时**读回 (按 lambda 作用域判!) ----
    # 全局查会被 evCallBlock / emitMouseMoveCall 里同形的 GetPropW 串命中 =>
    # 把本条挪进作用域才抓得到「这一条臂改回编译期常量」。
    if (-not $body.Contains('GetPropW(hwnd, L\"VB6_CtrlArrIndex\") - 1')) {
        $problems += 'R3: 派发 lambda 体内没有从 GetPropW 读运行时下标 (VB6_CtrlArrIndex - 1)'
    }
    if ($body.Contains('std::to_string(info.arrayIndex)')) {
        $problems += "R4: 派发 lambda 体内仍有 std::to_string(info.arrayIndex) => 编译期常量下标 (40 个 Label 一律答 39)"
    }
    # ---- R6: 两档都要在 (传址 / 传值) ----
    if (-not $body.Contains('"int16_t*"') -or -not $body.Contains('"int16_t"')) {
        $problems += "R6: 派发 lambda 里没有同时出现 int16_t* 与 int16_t 两档声明 => ByVal/ByRef 形态没分"
    }
    if (-not $body.Contains('"&vb6_evindex"') -or -not $body.Contains('"vb6_evindex"')) {
        $problems += "R6: 派发 lambda 里没有同时出现 &vb6_evindex 与 vb6_evindex 两个实参形态"
    }
}

# ---- R5: 旧名不得残留 ----
if ($src.Contains('evCallBlockUcByVal')) {
    $problems += "R5: 旧名 evCallBlockUcByVal 还有残留"
}

# ---- R7: 宿主必须 include 权威头 ----
if (-not (Test-Path $cppPath)) {
    $problems += "R7: 找不到宿主 cpp " + $cppPath
} else {
    $cpp = Get-Content $cppPath -Raw
    if (-not $cpp.Contains('common/event_param_abi.hpp')) {
        $problems += "R7: 宿主 cgen_util_ctrl.cpp 没有 include 权威头 common/event_param_abi.hpp"
    }
}

if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host "OK check_mouse_enter_index: R1..R7 (dispatch matches the definition side, runtime index, authority ABI)"
exit 0
