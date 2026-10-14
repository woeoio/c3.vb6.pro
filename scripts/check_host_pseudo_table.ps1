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
#   4) 账 #278 §B115/§B120: 宿主符号的拼法必须**问表**，而产物里同一个事实只许有
#      一个答复；§B120 再把「装配」与「是不是宿主伪对象」这两问各收成一处 (从前分别是 4 份与 5 份)。
#      这一段跑两次 --emit-c (只走前端, 不起 cl)。
#
# 为什么 4) 值得单独钉 (实测): cgen_with.cpp 以前把 `vb6_<对象>_<成员>` 手抄成两枚字面量,
# 而 RTL 把 UserControl / PropertyPage 两种大小写拼写**都声明且定义了** (vb6rtl_com.c:
# 1072/1073, vb6rtl_userctl.h:181/182) ⇒ 抄错的那一枚不是编不过，是编向**另一枚全局**。
# 改前同一份产物里一个事实两个答复:
#   With PropertyPage      ->  void* _vb6_with_0 = (void*)vb6_PropertyPage_hwnd
#   n = PropertyPage.hWnd  ->  n = vb6_PropertyPage_hWnd
# 判据两头: 那两行必须落在同一个符号上，且旧的小写拼写在 src/ (RTL 除外) 不许再有人引用。
# 今天两枚都还是 NULL (§B118: .pag 的宿主全局全仓 0 个写者) ⇒ 这一条钉的是将来接写者时
# 不许只接一头。
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
# Latin-1: 产物/诊断按字节 1:1 取，ASCII 片段(符号名、VB3001)不受 GBK 中文影响 (同 emit_manifest.ps1)
$lat2 = [System.Text.Encoding]::GetEncoding(28591)
$hdrRaw = ""
if (-not (Test-Path $tblPath)) { Write-Host "FAIL table file missing: $tblPath"; exit 1 }
if (-not (Test-Path $hdrPath)) { Write-Host "FAIL rtl header missing: $hdrPath"; exit 1 }
$hdrRaw = Get-Content $hdrPath -Raw
if (-not $hdrRaw) { Write-Host "FAIL rtl header empty: $hdrPath"; exit 1 }

# ---------- 1) 读表 ----------
$rowRe = '\{"(UserControl|PropertyPage|Extender|Ambient)",\s*"([a-z]+)",\s*"([A-Za-z]+)",\s*Vb6Type::(\w+),\s*(HPF_[A-Z| ]+)\}'
# Select-String 默认**不区分大小写** (§B120 实测: 把表里一行 obj 列改回小写, 这条不加就还是 54 行,
# 于是那行躲过的是"逐行对 RTL 声明"这道检查, 报出来的却是 ROW-SYMBOL-MISSING —— 红是红了, 说的不是这件事)
$rows = @(Select-String -Path $tblPath -Pattern $rowRe -CaseSensitive)
# 恰好 62 行: 从前的 "< 40" 会让"某一行从正则里掉出去"静默通过 —— 而每行都要逐条对 RTL 声明,
# 少认一行 = 那一行没人查 (账 #278 §B120 把 obj 列改成 PascalCase 时正是这个形状)。
# 改这张表要同批改这道针 (加一行 -> 63, 删一行 -> 61)。
# 54 -> 61: ComCtlsDemo 那一刀往表里补了 7 行宿主伪成员 (RTL 侧同步补了真身):
#   UserControl 的 AccessKeys / Appearance / DrawStyle / ContainedControls / EventsFrozen
#   (后两枚此前既不在表也不在 RTL ⇒ C2065, 实测 CoolBar/ToolBar 那一族);
#   Extender 的 Cancel / vb6_Default。逐行的 RTL 声明核对照旧跑, 故只是把数目改对。
# 61 -> 62: Fix 255 补 `UserControl.BackStyle` (LabelW.ctl:936/942 的
#   `BackStyle = UserControl.BackStyle` / `UserControl.BackStyle = Value`) ——
#   表里没有这一行时 hostPseudoRtlSymbol 会按 `vb6_<对象>_<成员名>` 硬拼出
#   vb6_UserControl_BackStyle, 而 RTL 没这枚全局 ⇒ C2065 ×2 (LabelW.c 1608/1618)。
#   同批补了 RTL 的 extern int32_t 与定义 (长 1 行的 .h/.c 两侧), 故数目 +1。
if ($rows.Count -ne 62) {
    $viol += ("TABLE-ROWS: parsed {0} rows (want exactly 62 - a deleted or reformatted row makes the per-row RTL checks spin)" -f $rows.Count)
}

# ---------- 2) 读 RTL 声明 (指针全局 / struct 全局 / 值全局 / 函数) ----------
# Fix 214: 拼写段放宽到 [A-Za-z_]+ —— VB6 关键字成员经 cIdent 转义后 RTL 名里会带下划线
# (Extender.Default → vb6_Extender_**vb6_**Default)。原来的 [A-Za-z]+ 匹配不到它,
# 表里那一行会被误判成 ROW-SYMBOL-MISSING。放宽只影响"能不能认出名字", 不放松类型比对。
$symType = @{}
foreach ($m in [regex]::Matches($hdrRaw, 'extern\s+([A-Za-z_][\w]*\s*\*)\s*(vb6_(?:UserControl|PropertyPage|Extender|Ambient)_[A-Za-z_]+)\s*;')) {
    $symType[$m.Groups[2].Value] = ($m.Groups[1].Value -replace '\s+', '')
}
foreach ($m in [regex]::Matches($hdrRaw, 'extern\s+(struct\s+[A-Za-z_]\w*)\s+(vb6_(?:UserControl|PropertyPage|Extender|Ambient)_[A-Za-z_]+)\s*;')) {
    $symType[$m.Groups[2].Value] = "struct"
}
foreach ($m in [regex]::Matches($hdrRaw, 'extern\s+([A-Za-z_][\w]*)\s+(vb6_(?:UserControl|PropertyPage|Extender|Ambient)_[A-Za-z_]+)\s*;')) {
    if (-not $symType.ContainsKey($m.Groups[2].Value)) { $symType[$m.Groups[2].Value] = $m.Groups[1].Value }
}
$fn = @{}
foreach ($m in [regex]::Matches($hdrRaw, '(?:static\s+inline\s+)?[A-Za-z_][\w\*]*\s+(vb6_(?:UserControl|PropertyPage|Extender|Ambient)_[A-Za-z]+)\s*\(')) {
    $fn[$m.Groups[1].Value] = 1
}

$expect = @{ Long = "int32_t"; Integer = "int16_t"; Boolean = "int16_t"; Byte = "uint8_t";
             String = "BSTR"; LongPtr = "void*"; Single = "float"; Double = "double" }
# 账 #278 §B120: 从前这里有一份 usercontrol->UserControl 的小映射 —— 那正是"对象那一段没有
# 权威"的样子。表的 obj 列现在就存 RTL 那一拼，下面直接用 $obj 装配，于是"逐行对 RTL"这道检查
# 变成对**权威自己**的检查。

$scalarRows = 0
$chanRows = 0
foreach ($r in $rows) {
    $g = [regex]::Match($r.Line, $rowRe)
    $obj = $g.Groups[1].Value; $name = $g.Groups[2].Value
    $rtl = $g.Groups[3].Value; $type = $g.Groups[4].Value
    $flags = $g.Groups[5].Value
    # 账 #278 §B105: HPF_CHANNEL 行的契约不是 `vb6_<对象>_<rtl>`，而是 channel 字段点名的那条
    # 码头（`Controls` 走 vb6_UC_Controls()，不是恒 NULL 的 vb6_PropertyPage_Controls 空桩）。
    # 所以这一档换一条判据：码头符号必须在 RTL 真有人声明，且这类行不许带值类型。
    if ($flags -match 'HPF_CHANNEL') {
        $cm = [regex]::Match($r.Line, 'HPF_CHANNEL,\s*"([A-Za-z_]\w*)"')
        if (-not $cm.Success) {
            $viol += ("CHANNEL-ROW-NODOCK: {0}.{1} flags HPF_CHANNEL but names no dock symbol" -f $obj, $name)
            continue
        }
        $dock = $cm.Groups[1].Value
        $chanRows++
        if ($type -ne "Unknown") {
            $viol += ("CHANNEL-ROW-TYPE: {0}.{1} is answered by dock {2} yet table claims {3}" -f $obj, $name, $dock, $type)
        }
        $decl = @(Get-ChildItem -Path (Join-Path $Root "src\rtl") -Recurse -Include *.h |
                  Select-String -Pattern ("\b" + [regex]::Escape($dock) + "\s*\("))
        if ($decl.Count -lt 1) {
            $viol += ("CHANNEL-DOCK-MISSING: {0}.{1} -> {2}() is not declared in any src/rtl header" -f $obj, $name, $dock)
        }
        continue
    }
    $sym = "vb6_" + $obj + "_" + $rtl
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
    @{ File = "src\semantics\semantic_analyzer_util.cpp"; Pat = '"(changed|scalewidth|scaleheight|scalemode|containerhwnd|enabled|autoredraw|hdc|hwnd)"' },
    # 账 #278 §B105: 码头的符号名只许写在表里；发码侧再硬编码一次 vb6_UC_Controls() = 第二份权威
    @{ File = "src\backend\detail\expr\cgen_expr_ident_builtin.inc"; Pat = '"vb6_UC_Controls\(\)"' }
    ,@{ File = "src\backend\stmt\cgen_with.cpp"; Pat = '"vb6_(UserControl|PropertyPage)_h[A-Za-z]*"' }  # 账 #278 §B115
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
    @{ File = "src\semantics\semantic_analyzer_util.cpp"; Pat = 'hostPseudoBareEligible\(' },
    # 账 #278 §B105: HPF_CHANNEL 那一档只经 hostPseudoChannel 出口流动 —— 语义层放行与发码
    # 选码头都问它；表里删掉这句 = 两头退回各自硬编码。
    @{ File = "src\common\host_pseudo.hpp"; Pat = 'inline bool hostPseudoChannel' },
    @{ File = "src\semantics\semantic_analyzer_util.cpp"; Pat = 'hostPseudoChannel\(' },
    @{ File = "src\backend\detail\expr\cgen_expr_ident_builtin.inc"; Pat = 'hostPseudoChannel\(' }
)
foreach ($m in $must) {
    $p = Join-Path $Root $m.File
    if (-not (Test-Path $p)) { $viol += ("MISSING " + $m.File); continue }
    if (-not (Select-String -Quiet -Path $p -Pattern $m.Pat)) {
        $viol += ("AUTHORITY-ABSENT: {0} has no {1}" -f $m.File, $m.Pat)
    }
}

# ---------- 6) 账 #278 §B115: 一个事实一个答复 (发码针, 只走前端) ----------
$withRel = "src\backend\stmt\cgen_with.cpp"
$withPath = Join-Path $Root $withRel
if (-not (Test-Path $withPath)) {
    $viol += ("MISSING " + $withRel)
} else {
    # §B120: 这一处的"问表"由第 7) 段的 ASSEMBLY-SITE 钉 (hostPseudoRtlSymbol)，不再钉旧出口
    # canonicalHostPseudoMember —— 装配上收之后那一问已经不在这里做了，留着就是把判据钉在历史上。
}
# 旧的小写拼写只是 RTL 的兼容别名: src/ (RTL 那一头除外) 再引用它 = 第二份答案
$oldHits = @(Get-ChildItem -Path (Join-Path $Root "src") -Recurse -Include *.cpp, *.inc, *.hpp, *.h |
             Where-Object { $_.FullName -notmatch '\\rtl\\' } |
             Select-String -Pattern 'vb6_PropertyPage_hwnd\b' |
             Where-Object { $_.Line -notmatch '^\s*(//|\*)' })
foreach ($o in $oldHits) {
    $viol += ("OLD-SPELLING: " + $o.Path.Substring($Root.Length + 1) + ":" + $o.LineNumber)
}
$c3x = Join-Path $Root ".build\C3.exe"
$pagRel = "tests\dochost\dhWithHost.pag"
if (-not (Test-Path $c3x)) {
    $viol += ("NEEDLE: no compiler at " + $c3x + " (this check reads the emitted C)")
} elseif (-not (Test-Path (Join-Path $Root $pagRel))) {
    $viol += ("NEEDLE: missing fixture " + $pagRel)
} else {
    $tmpC = Join-Path ([System.IO.Path]::GetTempPath()) ("hostpseudo_" + [guid]::NewGuid().ToString("N") + ".c")
    $tmpE = [System.IO.Path]::ChangeExtension($tmpC, ".err")
    $proc = Start-Process -FilePath $c3x -ArgumentList @('"' + (Join-Path $Root $pagRel) + '"', '--emit-c') -WorkingDirectory $Root -RedirectStandardOutput $tmpC -RedirectStandardError $tmpE -NoNewWindow -Wait -PassThru
    $emit = $lat2.GetString([System.IO.File]::ReadAllBytes($tmpC))
    $errt = $lat2.GetString([System.IO.File]::ReadAllBytes($tmpE))
    Remove-Item $tmpC, $tmpE -Force -ErrorAction SilentlyContinue
    if ($proc.ExitCode -ne 0) { $viol += ("NEEDLE: fixture exits " + $proc.ExitCode) }
    $mWith = [regex]::Match($emit, '(?m)^\s*void\* _vb6_with_\d+ = \(void\*\)(vb6_PropertyPage_h\w+)')
    $mRead = [regex]::Match($emit, '(?m)^\s*n = (vb6_PropertyPage_h\w+);')
    if (-not $mWith.Success) { $viol += 'NEEDLE-WITH: the With-object head is not in the emit' }
    if (-not $mRead.Success) { $viol += 'NEEDLE-READ: the qualified PropertyPage.hWnd head is not in the emit' }
    $twoMsg = "TWO-ANSWERS: With head says {0} but the qualified head says {1}"
    # -cne: 这条钉的就是**只差大小写**的两枚全局，-ne (PowerShell 默认不区分大小写) 看不见它
    if ($mWith.Success -and $mRead.Success -and $mWith.Groups[1].Value -cne $mRead.Groups[1].Value) {
        $viol += ($twoMsg -f $mWith.Groups[1].Value, $mRead.Groups[1].Value)
    }
    if (@([regex]::Matches($emit, 'vb6_PropertyPage_hWnd')).Count -lt 2) {
        $viol += 'NEEDLE-SAME: fewer than two heads land on vb6_PropertyPage_hWnd'
    }
    # 反面证人: 放开的是那张表里那两个位，不是「凡是文档名都合法」—— 裸位的 VBA 还得响
    $n3001 = @([regex]::Matches($errt, 'VB3001')).Count
    $bareMsg = "NEEDLE-BARE: VB3001 count {0} on the fixture (want exactly 1 - the VBA value-position head)"
    if ($n3001 -ne 1) { $viol += ($bareMsg -f $n3001) }
}

# ---------- 7) 账 #278 §B120: 命名契约与"是不是宿主伪对象"各只有一个住所 ----------
$hp = Join-Path $Root "src\common\host_pseudo.hpp"
$com = Join-Path $Root "src\backend\cgen_util_com.cpp"
$hlp = Join-Path $Root "src\backend\detail\util\cgen_helpers.inc"
function Count-InFile([string]$path, [string]$pat) {
    if (-not (Test-Path $path)) { return -1 }
    return @([regex]::Matches($lat2.GetString([System.IO.File]::ReadAllBytes($path)), $pat)).Count
}
$cKnown = Count-InFile $hp 'inline bool hostPseudoObjectKnown'
$cSymDef = Count-InFile $hp 'inline bool hostPseudoObjectSymbol'
$cAsmDef = Count-InFile $com 'std::string CCodeGen::hostPseudoRtlSymbol'
$cAsmDecl = Count-InFile $hlp 'std::string hostPseudoRtlSymbol'
if ($cKnown -ne 1) { $viol += ("OBJECT-KNOWN-DEF " + $cKnown + " (want exactly 1)") }
if ($cSymDef -ne 1) { $viol += ("OBJECT-SYMBOL-DEF " + $cSymDef + " (want exactly 1)") }
if ($cAsmDef -ne 1) { $viol += ("ASSEMBLY-DEF " + $cAsmDef + " (want exactly 1)") }
if ($cAsmDecl -ne 1) { $viol += ("ASSEMBLY-DECL " + $cAsmDecl + " (want exactly 1)") }
# "是不是宿主伪对象"从前在发码侧抄了五份名单：再出现手抄的四档并列 = 又一份权威
$listHits = @(Get-ChildItem -Path (Join-Path $Root "src\backend") -Recurse -Include *.cpp, *.inc |
              Select-String -Pattern 'objLower\w* == "usercontrol"' -CaseSensitive)
foreach ($l in $listHits) {
    $viol += ("MEMBER-LIST: " + $l.Path.Substring($Root.Length + 1) + ":" + $l.LineNumber)
}
# 四条发码路各自问那一个装配出口 (少一路 = 那一路退回现拼；多一路 = 一句两个答案)
$asmSites = @(
    @{ F = "src\backend\stmt\cgen_with.cpp"; N = 1 },
    @{ F = "src\backend\detail\stmt\cgen_assign_host_pseudo.inc"; N = 1 },
    @{ F = "src\backend\detail\expr\cgen_expr_ident_builtin.inc"; N = 1 },
    @{ F = "src\backend\detail\expr\cgen_expr_member_m22_module.inc"; N = 1 }
)
foreach ($s in $asmSites) {
    $sp = Join-Path $Root $s.F
    $c = Count-InFile $sp 'hostPseudoRtlSymbol\('
    if ($c -ne $s.N) {
        $viol += ("ASSEMBLY-SITE: " + $s.F + " asks the assembly exit " + $c + " times (want " + $s.N + ")")
    }
}
# 大小写那枚发码针 (§B120 的另一头)：VB 不区分大小写，源码把宿主对象写成小写也必须落在 RTL 符号上。
# 从前这里是把**源码拼写**抄进 C 标识符 ⇒ `usercontrol.hDC` 发成 vb6_usercontrol_hDC (没人声明它)。
$ctlCase = Join-Path $Root "tests\dochost\dhWithCase.ctl"
if (Test-Path $c3x) {
    if (-not (Test-Path $ctlCase)) {
        $viol += "NEEDLE-CASE: missing fixture tests/dochost/dhWithCase.ctl"
    } else {
        $tmp2 = Join-Path ([System.IO.Path]::GetTempPath()) ("hp11_" + [guid]::NewGuid().ToString("N") + ".c")
        $pr2 = Start-Process -FilePath $c3x -ArgumentList @('"' + $ctlCase + '"', '--emit-c') -WorkingDirectory $Root -RedirectStandardOutput $tmp2 -RedirectStandardError ([System.IO.Path]::ChangeExtension($tmp2, ".err")) -NoNewWindow -Wait -PassThru
        $emit2 = $lat2.GetString([System.IO.File]::ReadAllBytes($tmp2))
        Remove-Item $tmp2, ([System.IO.Path]::ChangeExtension($tmp2, ".err")) -Force -ErrorAction SilentlyContinue
        if ($pr2.ExitCode -ne 0) { $viol += ("NEEDLE-CASE: fixture exits " + $pr2.ExitCode) }
        $wantCase = @('\(void\*\)vb6_UserControl_hWnd', 'n = vb6_UserControl_hDC')
        foreach ($w in $wantCase) {
            if (@([regex]::Matches($emit2, $w)).Count -lt 1) {
                $viol += ("NEEDLE-CASE: emit has no " + $w + " (the object segment came from the source spelling)")
            }
        }
        if (@([regex]::Matches($emit2, '\x22?vb6_usercontrol_')).Count -ne 0) {
            $viol += "NEEDLE-CASE: emit still carries vb6_usercontrol_* (nobody declares that)"
        }
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
