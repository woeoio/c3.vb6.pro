# check_legacy_slot_key_owners.ps1 - legacy 接口槽表的**去重键**只许问 D2 那张词汇表
#
# 起因 (2026-10-23, 账 #305 新账 #489): VB6 风格接口 (legacy) 铸 vtbl 那两趟 ——
#   ① 形状/字段那一趟 (driver_codegen_dup_module_vars.inc 的 seen + backend 的 slot 表)
#   ② 实现绑定那一趟 (cgen_com.cpp 的 harvest / implOfSlot)
# 从前两侧各按**裸成员名**去重。同名属性的 Get 与 Let 于是共用一席，留哪一档由各自的
# 声明序决定: 实测 (探针 .build/k2739_emit2.txt) 接口侧按 Let 的形状铸出
# `void (*Label)(void* me, BSTR v)`，实现侧却把 **Get 的函数**填进同一格 (硬套 setter 的
# 强转) = **静默调错函数，一条 cl 诊断都不出**；同时 Get 那一档的形状在表里消失 ⇒
# 第一百七十八刀那位读者问不到 (retType=void 当场拒接)，属性读落回晚绑定 helper 的首枚
# void* 槽 = cl C2172。修法不是新立答案: tB 式接口早就有那张槽键词汇表 (ifaceSlotKey /
# D2)，legacy 这两趟改成问它 ⇒ 一个名字两个档位。
# 这条线一旦有人再按裸名去重 (或自己拼 `_put` 之外的第二套字段名)，同一件事就又有了两位
# 主人 ⇒ 本哨兵按"位数"钉死，加读者要回台账改数。
#
# 规则 (改坏了会红, 不是装饰):
#   P1  ifacePropSlotPrefix( 恰好 3 处且全在 interface_sig.hpp = 定义 1 + 两位转调者各 1。
#   P2  "get_" / "put_" / "putref_" 那三枚前缀字面量各恰好 1 处, 且都住在 interface_sig.hpp
#       —— 那张词汇表从前住过两处 (ifaceSlotKeyNamed 与 ifaceSlotPrefix), 本刀合一。
#   P3  ifaceSlotKeyNamed( 恰好 8 处 = 定义 1 + ifaceSlotKey 本体转调 1 + legacy 那五站 + 第 186 刀带参写槽查询 1
#       (去重键 1 [driver inc] / 槽键与字段 1 [cgen_iface_vtbl] / 实现绑定 1 [cgen_com] /
#        属性读那一问 + 属性写那一问 = 2 [cgen_util_classtype 两位读者])。
#   P4  ifaceDeclProcKind( 恰好 4 处 = 定义 1 + ifaceSlotKey 本体 1 + legacy 两站各 1。
#   P5  legacy 槽键只许铸一次: `s.slotKey = ifaceSlotKeyNamed(` 恰好 1 处 (cgen_iface_vtbl.cpp)。
#   P6  字段名的第二档只许这一处: `+ "_put"` 与 `+ "_putref"` 各恰好 1 处且都在
#       cgen_iface_vtbl.cpp (Get/Sub/Function 保持 cIdent(书写名) ⇒ 全部存量产物零移动)。
#   P7  裸名那一把键只许当**原料**、不许当键: `ifaceSlotKey(ifLower, written)` 恰好 1 处且
#       必须嵌在 ifaceSlotKeyNamed 那一发里 (driver_codegen_dup_module_vars.inc)。
#   P8  实现绑定那一趟必须带档: cgen_com.cpp 里 `ASTNodeKind::PropertyDecl, prop.propKind`
#       恰好 1 处 (harvest 的第三条臂)。
#
# 用法: powershell -File scripts\check_legacy_slot_key_owners.ps1 [-Root <副本目录>]  (PASS = exit 0)

param([string]$Root = "")

$ErrorActionPreference = "Stop"
if ($Root -eq "") { $Root = Split-Path -Parent $PSScriptRoot }

$srcDir = Join-Path $Root "src"
if (-not (Test-Path -LiteralPath $srcDir)) { Write-Host ("FAIL missing src dir: " + $srcDir); exit 1 }

# 只认**代码行**: 以 // 、/* 、* 开头的整行不算, 其余行裁掉行内 // 之后才算。
$code = @{}
foreach ($f in Get-ChildItem -LiteralPath $srcDir -Recurse -Include *.cpp,*.h,*.hpp,*.inc -File) {
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($line in [System.IO.File]::ReadAllLines($f.FullName)) {
        $t = $line.TrimStart()
        if ($t.StartsWith("//") -or $t.StartsWith("/*") -or $t.StartsWith("*")) { continue }
        $idx = $line.IndexOf("//")
        if ($idx -ge 0) { $line = $line.Substring(0, $idx) }
        $keep.Add($line)
    }
    $code[$f.FullName] = $keep
}

function CountWhere([string]$needle, [string]$fileName) {
    $n = 0
    foreach ($kv in $code.GetEnumerator()) {
        if ($fileName -and ($kv.Key -notlike ("*" + $fileName))) { continue }
        foreach ($line in $kv.Value) {
            if ($line.Contains($needle)) { $n++ }
        }
    }
    return $n
}

$fail = @()

$p1 = CountWhere "ifacePropSlotPrefix(" "interface_sig.hpp"
$p1all = CountWhere "ifacePropSlotPrefix(" ""
if ($p1 -ne 3 -or $p1all -ne 3) {
    $fail += ("P1 前缀表那一处主人: 期望 3 = 定义 + 两位转调者, 实到 本文件=" + $p1 + " 全仓=" + $p1all)
}

foreach ($pair in @(@('"get_"', 'return "get_"'), @('"put_"', 'return "put_"'), @('"putref_"', 'return "putref_"'))) {
    $c = CountWhere $pair[1] ""
    $ch = CountWhere $pair[1] "interface_sig.hpp"
    if ($c -ne 1 -or $ch -ne 1) {
        $fail += ("P2 前缀字面量 " + $pair[0] + ": 期望全仓 1 处且住在 interface_sig.hpp, 实到 全仓=" + $c + " 本文件=" + $ch)
    }
}

$p3 = CountWhere "ifaceSlotKeyNamed(" ""
if ($p3 -ne 8) {
    $fail += ("P3 槽键主人位数: 期望 8 = 定义 + ifaceSlotKey 转调 + legacy 那五站 + 第 186 刀带参写槽查询, 实到 " + $p3)
}
foreach ($fn in @("driver_codegen_dup_module_vars.inc", "cgen_iface_vtbl.cpp", "cgen_com.cpp")) {
    $c = CountWhere "ifaceSlotKeyNamed(" $fn
    if ($c -ne 1) { $fail += ("P3 legacy 那一站 " + $fn + ": 只许问一次槽键主人, 实到 " + $c) }
}
if ((CountWhere "ifaceSlotKeyNamed(" "cgen_util_classtype.cpp") -ne 3) {
    $fail += "P3 那一份文件: 读侧/写侧/带参写槽查询三位读者各问一次槽键主人, 别处再问就是第二位主人"
}

$p4 = CountWhere "ifaceDeclProcKind(" ""
if ($p4 -ne 4) {
    $fail += ("P4 「这一档是 Get/Let/Set」那一问位数: 期望 4 = 定义 + ifaceSlotKey 本体 + legacy 两站, 实到 " + $p4)
}

$p5 = CountWhere "s.slotKey = ifaceSlotKeyNamed(" ""
if ($p5 -ne 1) { $fail += ("P5 legacy 槽键的铸处: 期望恰好 1, 实到 " + $p5) }
if ((CountWhere "s.slotKey = ifaceSlotKeyNamed(" "cgen_iface_vtbl.cpp") -ne 1) {
    $fail += "P5 legacy 槽键铸处住错了文件 - 应当在 cgen_iface_vtbl.cpp"
}

$p6a = CountWhere '+ "_put"' "cgen_iface_vtbl.cpp"
$p6b = CountWhere '+ "_putref"' "cgen_iface_vtbl.cpp"
if ($p6a -ne 1 -or $p6b -ne 1) {
    $fail += ('P6 字段名第二档: 期望 + "_put" =1 / + "_putref" =1 且都在 cgen_iface_vtbl.cpp, 实到 ' + $p6a + ' / ' + $p6b)
}
if ((CountWhere '+ "_put"' "") -gt 1 -or (CountWhere '+ "_putref"' "") -gt 1) {
    $fail += 'P6 字段名那一档在别处又长了第二位主人'
}

$p7 = CountWhere "ifaceSlotKey(ifLower, written)" ""
$p7n = CountWhere "ifaceSlotKeyNamed(ifaceSlotKey(ifLower, written)" "driver_codegen_dup_module_vars.inc"
if ($p7 -ne 1 -or $p7n -ne 1) {
    $fail += ("P7 裸名那一把键: 期望只当原料 (全仓 1 处且嵌在 ifaceSlotKeyNamed 那一发里), 实到 裸=" + $p7 + " 嵌=" + $p7n)
}

$p8 = CountWhere "ASTNodeKind::PropertyDecl, prop.propKind" "cgen_com.cpp"
if ($p8 -ne 1) { $fail += ("P8 实现绑定那一趟的档: 期望 cgen_com.cpp 恰好 1 处, 实到 " + $p8) }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ("FAIL " + $m) }
    exit 1
}
Write-Host ("PASS 前缀表主人=" + $p1all + "(interface_sig.hpp) 槽键主人=" + $p3 + "(legacy 四站各 1) 档位问句=" + $p4 + " 槽键铸处=" + $p5 + " 字段第二档=" + $p6a + "+" + $p6b + " 裸名原料=" + $p7 + " 实现绑定带档=" + $p8)
exit 0
