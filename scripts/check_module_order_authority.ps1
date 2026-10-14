# check_module_order_authority.ps1 - 账 §B97 的结构哨兵 (第 41 道；只扫 src/，不起 cl)
#
# 为什么要这一道 (§B97 那一族的形状):
#   「多模块工程里外部模块名单」这枚数据往下决定三处输出的**行序** ——
#     · 聚合 .c 的跨模块 #include 段   (cgen_base_generate_c_open.inc)
#     · .h 的 crossmod #include 段     (cgen_base_generate_crossmod.inc)
#     · 入口点里 vb6_mod_<X>_init() 的调用序 (= 各模块全局的初始化先后, cgen_base_generate_entry.inc)
#   它以前是 std::unordered_set，由符号表那枚 unordered_map 的遍历造出 ⇒ 同一份 .vbp
#   换一台编译器就换一种启动序 (实测 tests/cc_id/Id.vbp 两台各交一种、且都不是工程序)。
#   现在权威 = driver 按 modules_ 的**下标序**交下来的有序 vector（那是 runParse 末尾「类模块
#   前移」那次 stable_sort 之后的序，见 driver_frontend.cpp:327-339 —— 要保的是**确定**，不是
#   照抄 .vbp 的字面行序，这句别说过头）。
#   这一道盯的是「别再让顺序由哈希桶决定」与「别再开第二份 init 发射循环」两件事。
#
# 规则 (改坏了会红):
#   P1  SymbolTable::getExternalModuleNames() 的声明必须是 std::vector<std::string>，
#       全仓不许再出现 `std::unordered_set<std::string> getExternalModuleNames`。
#   P2  后端存这份名单的成员 (cgen_state.inc 的 externalModules_) 必须是 vector。
#   P3  driver 必须经 extModulesInModuleVectorOrder 造名单 (定义 1 次 + 调用 1 次)，
#       且本地变量不许退回 `std::unordered_set<std::string> externalModules`。
#   P4  入口点里发射模块 init 的循环**只能有一枚** (emitExtInits 那个 lambda)，
#       调用点必须仍是 5 处 —— 5 处是入口点模板的形状 (WinMain/main × 启动判定分支)，
#       再长就是有人把循环内联回去了 (五份同形正是本账要收掉的东西)。
#   P5  两处 #include 发射循环各迭代**形参**那份有序名单，各恰好 1 次 (不许多不许少)。
#
# 用法: powershell -File scripts/check_module_order_authority.ps1   (PASS => exit 0)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$fails = @()

function RawOf([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return (Get-Content -LiteralPath $p -Raw -Encoding UTF8)
}
function CountOf([string]$text, [string]$needle) {
    if ($null -eq $text) { return -1 }
    return @([regex]::Matches($text, [regex]::Escape($needle))).Count
}

$HPP  = 'src\semantics\symbol_table.hpp'
$CPP  = 'src\semantics\symbol_table.cpp'
$ST   = 'src\backend\detail\util\cgen_state.inc'
$DRV  = 'src\driver\detail\driver_codegen_module_loop.inc'
$ENTRY = 'src\backend\detail\base\cgen_base_generate_entry.inc'
$COPEN = 'src\backend\detail\base\cgen_base_generate_c_open.inc'
$XMOD = 'src\backend\detail\base\cgen_base_generate_crossmod.inc'

# ---- P1 ----
$hpp = RawOf $HPP
$cpp = RawOf $CPP
if ($null -eq $hpp -or $null -eq $cpp) { $fails += 'P1 missing ' + $HPP }
else {
    if ((CountOf $hpp 'std::vector<std::string> getExternalModuleNames() const;') -ne 1) {
        $fails += 'P1 getExternalModuleNames must be declared as std::vector<std::string> (an ordered list) - the unordered_set version made the emitted #include / module-init order a function of the compiler binary'
    }
    if ((CountOf $hpp 'std::unordered_set<std::string> getExternalModuleNames') -ne 0) {
        $fails += 'P1 getExternalModuleNames is back to unordered_set'
    }
    if ((CountOf $cpp 'std::vector<std::string> SymbolTable::getExternalModuleNames() const {') -ne 1) {
        $fails += 'P1 the definition no longer returns an ordered vector'
    }
    if ((CountOf $cpp 'std::sort(result.begin(), result.end());') -ne 1) {
        $fails += 'P1 getExternalModuleNames lost its own sort - it must at least be self-deterministic'
    }
}

# ---- P2 ----
$st = RawOf $ST
if ($null -eq $st) { $fails += 'P2 missing ' + $ST }
else {
    if ((CountOf $st 'std::vector<std::string> externalModules_;') -ne 1) {
        $fails += 'P2 CCodeGen::externalModules_ must be a std::vector<std::string> - three emitting sites walk it in order'
    }
    if ((CountOf $st 'std::unordered_set<std::string> externalModules_;') -ne 0) {
        $fails += 'P2 externalModules_ is back to unordered_set'
    }
}

# ---- P3 ----
$drv = RawOf $DRV
if ($null -eq $drv) { $fails += 'P3 missing ' + $DRV }
else {
    if ((CountOf $drv 'auto extModulesInModuleVectorOrder = [&](size_t self) {') -ne 1) {
        $fails += 'P3 the driver must build the module list through extModulesInModuleVectorOrder (modules_ index order is the single authority)'
    }
    if ((CountOf $drv 'extModulesInModuleVectorOrder(i)') -ne 1) {
        $fails += 'P3 extModulesInModuleVectorOrder is not called exactly once in the per-module loop'
    }
    if ((CountOf $drv 'std::unordered_set<std::string> externalModules') -ne 0) {
        $fails += 'P3 the driver list went back to a local unordered_set'
    }
    if ((CountOf $drv 'classTypeParams.empty()') -lt 1) {
        $fails += 'P3 the generic-template exclusion vanished from the module list (dangling #include risk, see G4)'
    }
}

# ---- P4 ----
$entry = RawOf $ENTRY
if ($null -eq $entry) { $fails += 'P4 missing ' + $ENTRY }
else {
    if ((CountOf $entry 'for (const auto& extMod : externalModules_) {') -ne 1) {
        $fails += 'P4 entry.inc must hold exactly ONE loop that emits module inits (the emitExtInits lambda) - five copies were the original smell'
    }
    if ((CountOf $entry 'emitExtInits();') -ne 5) {
        $fails += 'P4 emitExtInits() must be called at each of the 5 entry-point templates (count is the census of those shapes)'
    }
}

# ---- P5 ----
foreach ($pair in @(@($COPEN, '#include'), @($XMOD, '#include'))) {
    $txt = RawOf $pair[0]
    if ($null -eq $txt) { $fails += ('P5 missing ' + $pair[0]); continue }
    if ((CountOf $txt 'for (const auto& extMod : externalModules) {') -ne 1) {
        $fails += ('P5 ' + $pair[0] + ' must walk the ordered parameter list exactly once when emitting ' + $pair[1] + ' lines')
    }
}

if ($fails.Count -gt 0) {
    $fails | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host ('OK check_module_order_authority: P1..P5 (module list = ordered vector in modules_ index order;' +
            ' init loops=1 x calls=5; include loops=1+1)')
exit 0
