#!/usr/bin/env powershell
# check_host_pseudo_table.ps1 - 结构哨兵: 宿主伪成员只有一个权威表 (账 #159)
#
# 背景 (实测): UserControl.hWnd / hDC / ScaleMode / Enabled 这些成员在 C 侧是 RTL
# 全局量 (src/rtl/core/vb6rtl/vb6rtl_userctl.h 的 extern int32_t / int16_t / void* /
# BSTR), 而发码的类型 oracle 认不得它们 ⇒ 答 Variant ⇒ 比较发成
#   vb6_VarCmpLongNe(&vb6_UserControl_hWnd, 0)
# —— 拿 8 字节 void* 的**地址**当 vb6_VARIANT* (RTL 签名第一形参, 16 字节) 递进
# RTL ⇒ 读到的是越界垃圾, 恒假; CStr 那一路落进 _Generic 的
# `default: vb6_VariantObject` (vb6rtl_variant.h:215) ⇒ 打空串。同一个决定 (哪个
# 成员叫什么 / 是什么类型 / 能不能裸写 / 赋值要不要拆) 此前抄在五份清单里, 覆盖面
# 还彼此不一致 (写 ScaleWidth 答 Long, 写 UserControl.ScaleWidth 才答 Long, 写
# UserControl.hWnd 两边都不答)。
#
# 现在只有一张表: src/common/host_pseudo.hpp 的 kHostPseudoRows, 五个消费点
# (裸名发射 / 两条类型 oracle / 赋值拆数值 / 语义层的裸写放行) 都只问它。
# 表从 src/backend/cgen_util_com.cpp 搬到 common 是账 #219 那一刀：语义层不许再自己抄一串
# 成员名 (那样每增一行 HPF_BARE 都要两处同改)，它只能问这张表。
#
# 本哨兵钉三件事:
#   1) 表里每一行的 vb6_<对象>_<拼写> 必须在 vb6rtl_userctl.h 真有其人;
#   2) 标量行的类型必须与那一行的 extern 声明对上 (Long↔int32_t, Integer/Boolean↔
#      int16_t, LongPtr↔void*, String↔BSTR) —— 表说错类型就是下一个装箱误读;
#   3) 那五份旧清单不许再被抄回来 (被禁的旧形状见 $deny)。
# 输出一律 ASCII (控制台是 GBK, 中文读数在重定向文件里不可 grep)。文件必须带
# UTF-8 BOM: PS 5.1 无 BOM 时按 ANSI 读, 行尾中文字节会吃掉换行 ⇒ param() 被并进
# 注释、参数全空、ParserError。
# 与 check_uc_scale_units.ps1 / check_di_stubs.ps1 同一类静态用例, 由 run_tests.ps1
# 的 [STATIC] 项调用。
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$TableFile = ""      # 负控用: 把表指到副本上跑
)
$ErrorActionPreference = "Continue"

$tblPath = if ($TableFile) { $TableFile } else { Join-Path $Root "src\common\host_pseudo.hpp" }
$hdrPath = Join-Path $Root "src\rtl\core\vb6rtl\vb6rtl_userctl.h"
$viol = @()
$hdrRaw = ""
if (-not (Test-Path $tblPath)) { Write-Host "FAIL table file missing: $tblPath"; exit 1 }
if (-not (Test-Path $hdrPath)) { Write-Host "FAIL rtl header missing: $hdrPath"; exit 1 }
$hdrRaw = Get-Content $hdrPath -Raw
if (-not $hdrRaw) { Write-Host "FAIL rtl header empty: $hdrPath"; exit 1 }

# ---------- 1) 读表 ----------
$rowRe = '\{"(usercontrol|propertypage|extender|ambient)",\s*"([a-z]+)",\s*"([A-Za-z]+)",\s*Vb6Type::(\w+),\s*(HPF_[A-Z| ]+)\}'
$rows = @(Select-String -Path $tblPath -Pattern $rowRe)
if ($rows.Count -lt 40) {
    $viol += ("TABLE-SHRUNK: parsed only {0} rows (a deleted or reformatted table would make checks below spin)" -f $rows.Count)
}

# ---------- 2) 读 RTL 声明 (指针全局 / struct 全局 / 值全局 / 函数) ----------
$symType = @{}
foreach ($m in [regex]::Matches($hdrRaw, 'extern\s+([A-Za-z_][\w]*\s*\*)\s*(vb6_(?:UserControl|PropertyPage|Extender|Ambient)_[A-Za-z]+)\s*;')) {
    $symType[$m.Groups[2].Value] = ($m.Groups[1].Value -replace '\s+', '')
}
foreach ($m in [regex]::Matches($hdrRaw, 'extern\s+(struct\s+[A-Za-z_]\w*)\s+(vb6_(?:UserControl|PropertyPage|Extender|Ambient)_[A-Za-z]+)\s*;')) {
    $symType[$m.Groups[2].Value] = "struct"
}
foreach ($m in [regex]::Matches($hdrRaw, 'extern\s+([A-Za-z_][\w]*)\s+(vb6_(?:UserControl|PropertyPage|Extender|Ambient)_[A-Za-z]+)\s*;')) {
    if (-not $symType.ContainsKey($m.Groups[2].Value)) { $symType[$m.Groups[2].Value] = $m.Groups[1].Value }
}
$fn = @{}
foreach ($m in [regex]::Matches($hdrRaw, '(?:static\s+inline\s+)?[A-Za-z_][\w\*]*\s+(vb6_(?:UserControl|PropertyPage|Extender|Ambient)_[A-Za-z]+)\s*\(')) {
    $fn[$m.Groups[1].Value] = 1
}

$expect = @{ Long = "int32_t"; Integer = "int16_t"; Boolean = "int16_t"; Byte = "uint8_t";
             String = "BSTR"; LongPtr = "void*"; Single = "float"; Double = "double" }
$objPascal = @{ usercontrol = "UserControl"; propertypage = "PropertyPage";
               extender = "Extender"; ambient = "Ambient" }

$scalarRows = 0
foreach ($r in $rows) {
    $g = [regex]::Match($r.Line, $rowRe)
    $obj = $g.Groups[1].Value; $name = $g.Groups[2].Value
    $rtl = $g.Groups[3].Value; $type = $g.Groups[4].Value
    $sym = "vb6_" + $objPascal[$obj] + "_" + $rtl
    if ($type -ne "Unknown") { $scalarRows++ }
    # 2a) 符号必须在场 (全局或函数) —— 收了没人声明的就是发一个 C2065
    if (-not ($symType.ContainsKey($sym) -or $fn.ContainsKey($sym))) {
        $viol += ("ROW-SYMBOL-MISSING: {0}.{1} -> {2} not declared in vb6rtl_userctl.h" -f $obj, $name, $sym)
        continue
    }
    # 2b) 标量行的类型必须与 RTL 声明对上
    if ($expect.ContainsKey($type) -and $symType.ContainsKey($sym)) {
        $got = $symType[$sym]
        if ($got -ne $expect[$type]) {
            $viol += ("ROW-TYPE-MISMATCH: {0} table says {1} (expects {2}) but rtl declares {3}" -f $sym, $type, $expect[$type], $got)
        }
    }
}

# ---------- 3) 旧形状禁止再被抄回来 ----------
$deny = @(
    # 五份旧清单的**定义本体** (注释里提名字是允许的, 故只禁声明形态)
    @{ File = "src\backend\cgen_util_com.cpp"; Pat = 'std::pair<const char\*, const char\*> kUserControlCanon' },
    @{ File = "src\backend\detail\expr\cgen_expr_ident_builtin.inc"; Pat = 'kUserControlHostMembers|kPropertyPageHostMembers' },
    @{ File = "src\backend\detail\stmt\cgen_assign_host_pseudo.inc"; Pat = 'kNumericHostMembers\[\]' },
    # 两条硬编码的类型回退 (裸名那一路 + 限定名那一路)
    @{ File = "src\backend\cgen_util_type.cpp"; Pat = 'lower == "scalewidth"|memLowerCz' },
    # 成员名字面量只许出现在那张表里: 消费点再按名字判成员 = 抄了第二份
    @{ File = "src\backend\detail\expr\cgen_expr_ident_builtin.inc"; Pat = '"(scalewidth|scaleheight|scalemode|containerhwnd|enabled|autoredraw|hdc|hwnd)"' },
    @{ File = "src\backend\detail\stmt\cgen_assign_host_pseudo.inc"; Pat = '"(scalewidth|scaleheight|scalemode|containerhwnd|enabled|autoredraw|hdc|hwnd)"' },
    # 语义层那一头同样禁用成员名字面量（账 #219 删掉的就是这一份）
    @{ File = "src\semantics\semantic_analyzer_util.cpp"; Pat = '"(changed|scalewidth|scaleheight|scalemode|containerhwnd|enabled|autoredraw|hdc|hwnd)"' }
)
foreach ($d in $deny) {
    $p = Join-Path $Root $d.File
    if (-not (Test-Path $p)) { $viol += ("MISSING " + $d.File); continue }
    foreach ($m in (Select-String -Path $p -Pattern $d.Pat)) {
        $viol += ("DENY {0}:{1}: {2}" -f $d.File, $m.LineNumber, $m.Line.Trim())
    }
}

# ---------- 4) 权威必须在场 (否则"禁旧形状"退化成没人管) ----------
$must = @(
    @{ File = "src\common\host_pseudo.hpp"; Pat = 'inline const HostPseudoRow kHostPseudoRows\[\]' },
    @{ File = "src\common\host_pseudo.hpp"; Pat = 'inline bool hostPseudoBareEligible' },
    @{ File = "src\backend\cgen_util_com.cpp"; Pat = 'bool CCodeGen::hostPseudoValueType' },
    @{ File = "src\backend\cgen_util_com.cpp"; Pat = 'bool CCodeGen::hostPseudoBareName' },
    @{ File = "src\backend\cgen_util_com.cpp"; Pat = 'bool CCodeGen::hostPseudoIsNumeric' },
    @{ File = "src\backend\detail\util\cgen_helpers.inc"; Pat = 'hostPseudoValueType' },
    @{ File = "src\backend\cgen_util_type.cpp"; Pat = 'hostPseudoValueType\(' },
    @{ File = "src\backend\detail\expr\cgen_expr_ident_builtin.inc"; Pat = 'hostPseudoBareName\(' },
    @{ File = "src\backend\detail\stmt\cgen_assign_host_pseudo.inc"; Pat = 'hostPseudoIsNumeric\(' },
    # 第五个消费点 (账 #219): 语义层的裸写放行必须问表，不许把成员名抄回 semantics。
    @{ File = "src\semantics\semantic_analyzer_util.cpp"; Pat = 'hostPseudoBareEligible\(' }
)
foreach ($m in $must) {
    $p = Join-Path $Root $m.File
    if (-not (Test-Path $p)) { $viol += ("MISSING " + $m.File); continue }
    if (-not (Select-String -Quiet -Path $p -Pattern $m.Pat)) {
        $viol += ("AUTHORITY-ABSENT: {0} has no {1}" -f $m.File, $m.Pat)
    }
}

# ---------- 5) 普查读数 (不判红, 给下一轮留证据) ----------
$files = Get-ChildItem -Path @((Join-Path $Root "src\backend"), (Join-Path $Root "src\common"), (Join-Path $Root "src\semantics")) -Recurse -Include *.cpp,*.inc,*.hpp
$cen = @($files | Select-String -Pattern 'vb6_(UserControl|PropertyPage|Extender|Ambient)_[A-Za-z]+' |
         Where-Object { $_.Path -notmatch 'host_pseudo\.hpp$' -and $_.Line -notmatch '^\s*//' })
Write-Host ("census: rows={0} scalar={1} hand-written-host-symbols-outside-table={2}" -f $rows.Count, $scalarRows, $cen.Count)
foreach ($c in $cen) {
    Write-Host ("   site: {0}:{1}" -f (Split-Path $c.Path -Leaf), $c.LineNumber)
}

if ($viol.Count -gt 0) {
    Write-Host ("FAIL: host-pseudo-table sentinel, {0} issue(s)" -f $viol.Count)
    foreach ($v in $viol) { Write-Host ("  " + $v) }
    exit 1
}
Write-Host ("OK: one host-pseudo table ({0} rows, {1} scalar); no old list came back" -f $rows.Count, $scalarRows)
exit 0
