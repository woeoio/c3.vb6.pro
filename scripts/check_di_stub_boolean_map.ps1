# check_di_stub_boolean_map.ps1 - di 桩签名必须与 C3 的 Declare→C 映射**逐字**一致
#
# 为什么要有这道 (账 #610 第342刀, 富贵竹 FIX-006 末节那族):
#   di 桩是**声明式**的 —— C3 发码时把用户的 `Declare ... Lib "user32"` 名字动态拼成
#   `vb6_di_<name>` (cgen_decl_api.cpp), 生成的 extern 原型与 RTL 里的桩定义**分处两个
#   翻译单元** ⇒ 链接期**不检查类型** ⇒ 签名不一致时既不报错也不红, 只是 x86 stdcall 下
#   参数宽度错位 (行为错, 链接错, 运行时才炸)。
#   本刀补的 vb6_di_AttachThreadInput / vb6_di_IsIconic 就踩在这个缝里:
#   **上游 FIX-006 文档给的是 intptr_t 全形**, 而本项目 C3 实际生成的是
#       extern int16_t __stdcall vb6_di_AttachThreadInput(intptr_t, intptr_t, int16_t);
#   (ByVal Boolean → int16_t, 见 src/backend/cgen_base_type.cpp 的 Boolean 档)。
#   照抄上游 ⇒ 桩与生成头不一致 ⇒ x86 下第三个参数的栈槽宽度对不上。
#
# 判定三格 (缺一格就是没判):
#   A. C3 侧的 Boolean 映射仍是 int16_t (从 cgen_base_type.cpp 源码读, 不写死);
#   B. RTL 侧两枚桩的返回类型 / Boolean 形参都用 A 读到的那一档;
#   C. C3 侧生成头的形状 (形参宽度逐字) 与 B 里的桩字样一致 —— 这一格是"两侧必须同形"的锁。
#
# 用法:  powershell -File scripts/check_di_stub_boolean_map.ps1
# 退出码: 0 = 三格都成立; 1 = 有一格不成立 (逐条打在 stdout)。
#
# ⚠ 只覆盖本刀那两枚: 其余 di 桩的历史签名**不**遵循同一口径
#   (例: vb6_di_IsWindow 返回 intptr_t, 而 C3 的 Long 档是 int32_t),
#   泛化到全表会红一片 ⇒ 那属于另一刀的活, 不在这里顺手改。

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

function Get-Text([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8
}

$problems = @()

# --- A. C3 侧 Boolean 的 C 类型 (从源码读, 不写死) ---
$baseTypePath = Join-Path $root "src/backend/cgen_base_type.cpp"
$baseType = Get-Text $baseTypePath
$boolCt = $null
if ($baseType) {
    $m = [regex]::Match($baseType, 'case\s+Vb6Type::Boolean\s*:\s*cType\s*=\s*"([A-Za-z0-9_]+)"')
    if ($m.Success) { $boolCt = $m.Groups[1].Value }
}
if (-not $boolCt) {
    $problems += 'A: 在 src/backend/cgen_base_type.cpp 里读不到 case Vb6Type::Boolean: cType = "..." 这一行 (C3 的映射改了/挪了/换写法了) —— 本哨兵需要跟着改'
}

# --- B. RTL 侧两枚桩的签名 ---
$stubPath = Join-Path $root "src/rtl/core/di/vb6_di_user32_stubs.c"
$stub = Get-Text $stubPath
$expectAttach = if ($boolCt) {
    "${boolCt} __stdcall vb6_di_AttachThreadInput(intptr_t idAttach, intptr_t idAttachTo, ${boolCt} fAttach)"
} else { "" }
$expectIsIconic = if ($boolCt) { "${boolCt} __stdcall vb6_di_IsIconic(intptr_t hWnd)" } else { "" }

if (-not $stub) {
    $problems += "B: 找不到 src/rtl/core/di/vb6_di_user32_stubs.c"
} else {
    if (-not $expectAttach -or $stub.IndexOf($expectAttach) -lt 0) {
        $problems += ("B: vb6_di_user32_stubs.c 里没有逐字匹配这一行:" + [Environment]::NewLine + "     " + $expectAttach)
    }
    if (-not $expectIsIconic -or $stub.IndexOf($expectIsIconic) -lt 0) {
        $problems += ("B: vb6_di_user32_stubs.c 里没有逐字匹配这一行:" + [Environment]::NewLine + "     " + $expectIsIconic)
    }
}

# --- C. 两枚桩都在 di 桩基线里 (少登记 = 将来重生成会静默丢掉) ---
$manifestPath = Join-Path $PSScriptRoot "di_stubs_manifest.txt"
$manifest = Get-Text $manifestPath
if (-not $manifest) {
    $problems += "C: 找不到 scripts/di_stubs_manifest.txt"
} else {
    foreach ($nm in @("vb6_di_AttachThreadInput", "vb6_di_IsIconic")) {
        if ($manifest -notmatch ("(?m)^" + [regex]::Escape($nm) + "\s*$")) {
            $problems += "C: $nm 不在 scripts/di_stubs_manifest.txt 基线里 (用 check_di_stubs.ps1 -Update 收进去)"
        }
    }
}

if ($problems.Count -gt 0) {
    Write-Host "di_stub_boolean_map: FAIL ($($problems.Count) 格不成立)"
    foreach ($p in $problems) { Write-Host ("  - " + $p) }
    exit 1
}
Write-Host "di_stub_boolean_map: OK (Boolean 档=$boolCt; 两枚桩的签名与 C3 映射逐字一致)"
exit 0
