# check_decl_alias_guard.ps1 - 账 #305 第一百九十八刀 族A 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 「一枚 DLL 导出的 LastDllError 快照包装发在哪、按什么去重」这一问原本
# 只有一把钥匙 (diGuard = 导出名), 而这条钥匙同时被**非别名路** (`Declare Sub RtlMoveMemory
# Lib "kernel32"`) 与**别名路** (`Declare Sub CopyMemory ... Alias "RtlMoveMemory"`) 共用 ——
# VB6 语料里两条写在同一个模块里是常态 (实物 = VbEclipse MagneticWnd.cls:120/:182)。先发那条
# 占下 guard 之后, 别名路那**整块** (extern + 全工程唯一一份 static __inline + #define) 被
# 预处理器吃掉, 调用点却仍按别名直发 `vb6_lwa_RtlMoveMemory` ⇒ cl 只报 C4013 (隐式声明是警告、
# 压根不进错误日志), 到链接期才露出 LNK2019。修 = 别名路的包装按**自己的 C 名**去重
# (`VB6_LWA_<导出名>_DEFINED`)。哨兵钉的是: 这把新钥匙只有一个主人铸、两条路的 C 名不许并成
# 一份、夹具与两头发码针齐件。
#
# 规则 (改坏了会红, 不是装饰):
#   R1  新钥匙只有一个主人: `VB6_LWA_` 在 cgen_decl_api.cpp 恰好 1 处 (拼 guard 名那一行);
#       `#ifndef " + lwaGuard` 与 `#define " + lwaGuard` 各恰好 1 —— 成对, 少一行就是裸发,
#       多模块工程里同一个别名声明会落在多份头里, 同一个 .c 立刻撞 C2084 (已经有函数体)。
#   R2  别名路不许再回「蹭 diGuard」那一形: 单行 `if (!isStaticDecl && aliasedCall)
#       emitLdlWrapper();` 恰好 0; `emitLdlWrapper();` 全文件恰好 2 = 别名路 1 + 非别名路 1。
#   R3  Fix 217 那两套 C 名仍不相交: `("vb6_lwa_" + sanitizedExport)` 与
#       `("vb6_lw_" + cFuncIdent)` 各恰好 1; diGuard 仍按导出名拼 (恰好 1 处)。
#   R4  夹具与真跑判据齐件: k3136 三件套在; run_tests.ps1 里夹具名恰好 9 处
#       (头注 1 + 两架构用例名 2 + 四份路径 + 两枚发码格名 = 9)。
#   R5  两头发码针各恰好 1 处: 正针 = 新 guard 两行 + 别名路包装那份签名 (带 pDest/ByteLen,
#       不是按名那条的 Destination/Length); 反面 = 三条「修法走歪」的形状。
# 用法: powershell -File scripts\check_decl_alias_guard.ps1   (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$script:bad = @()

function Text-Of([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { $script:bad += ("缺文件 " + $rel); return "" }
    return (Get-Content -Raw -Encoding UTF8 $p)
}

function Count-Of([string]$hay, [string]$needle) {
    if (-not $hay) { return 0 }
    return @([regex]::Matches($hay, [regex]::Escape($needle))).Count
}

$apiRel = "src\backend\decl\cgen_decl_api.cpp"
$rtRel  = "tests\run_tests.ps1"
$api = Text-Of $apiRel
$rt  = Text-Of $rtRel

# ---- R1 新钥匙只有一个主人 ----
$nKey = Count-Of $api "VB6_LWA_"
if ($nKey -ne 1) { $script:bad += ("R1 VB6_LWA_ 在 cgen_decl_api.cpp = $nKey 处 (要 1: 只有这一处拼那把钥匙)") }
$nIf   = Count-Of $api '#ifndef " + lwaGuard'
$nDef  = Count-Of $api '#define " + lwaGuard'
if ($nIf -ne 1)  { $script:bad += ("R1 #ifndef 那一行 = $nIf 处 (要 1: 少了它就是裸发 static __inline ⇒ C2084)") }
if ($nDef -ne 1) { $script:bad += ("R1 #define 那一行 = $nDef 处 (要 1: 少了它同一份 .c 里两条头各发一遍)") }

# ---- R2 不许回退成「别名路整块被 diGuard 吃掉」 ----
if ((Count-Of $api "if (!isStaticDecl && aliasedCall) emitLdlWrapper();") -ne 0) {
    $script:bad += "R2 旧形回来了 (别名路的包装又塞进 diGuard 块里 = 撞名那条 Declare 会把它整块吃掉)"
}
$nEmit = Count-Of $api "emitLdlWrapper();"
if ($nEmit -ne 2) { $script:bad += ("R2 emitLdlWrapper(); 调用 = $nEmit 处 (要 2: 别名路 1 + 非别名路 1)") }

# ---- R3 两套 C 名仍不相交 ----
if ((Count-Of $api '("vb6_lwa_" + sanitizedExport)') -ne 1) {
    $script:bad += "R3 别名路的命名出口不是一处 (vb6_lwa_ + 导出名)"
}
if ((Count-Of $api '("vb6_lw_" + cFuncIdent)') -ne 1) {
    $script:bad += "R3 非别名路的命名出口不是一处 (vb6_lw_ + VB 名)"
}
if ((Count-Of $api "std::string diGuard = (isStaticDecl") -ne 1) {
    $script:bad += "R3 diGuard 那把按导出名的钥匙不再只有一处拼"
}

# ---- R4 夹具与真跑判据齐件 ----
foreach ($f in @("tests\k3136_lwa_guard\K3136.vbp", "tests\k3136_lwa_guard\mdl3136.bas")) {
    if (-not (Test-Path -LiteralPath (Join-Path $root $f))) { $script:bad += ("R4 缺夹具 " + $f) }
}
$nCells = Count-Of $rt "k3136_lwa_guard"
if ($nCells -ne 9) { $script:bad += ("R4 run_tests.ps1 里夹具名 = $nCells 处 (要 9: 头注 1 + 用例名 2 + 路径 4 + 发码格名 2)") }

# ---- R5 两头发码针 ----
foreach ($c in @('#ifndef VB6_LWA_RtlMoveMemory_DEFINED',
                 '#define VB6_LWA_RtlMoveMemory_DEFINED',
                 'static __inline void __stdcall vb6_lwa_RtlMoveMemory(void* pDest, void* pSrc, intptr_t ByteLen) {',
                 '#define CopyMemory vb6_lw_RtlMoveMemory',
                 '#define RtlMoveMemory vb6_lwa_RtlMoveMemory',
                 'vb6_lwa_RtlMoveMemory(void* Destination')) {
    $k = Count-Of $rt $c
    if ($k -ne 1) { $script:bad += ("R5 发码针 " + $c + " = $k 处 (要恰好 1)") }
}

if ($script:bad.Count -eq 0) {
    $summary = "PASS key=$nKey guard-pair=$nIf/$nDef emit-calls=$nEmit legacy-shape=0 cells=$nCells"
    Write-Host $summary -ForegroundColor Green
    exit 0
}
foreach ($b in $script:bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
