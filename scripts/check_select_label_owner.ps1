# check_select_label_owner.ps1 -- 账 #604 (第321刀): 混合标签 Variant Select Case 的标签箱谁来放
#
# 病灶 (修复前实测, .local/k321): 319 刀给混合标签发的每一枚标签箱是
#     vb6_VARIANT _vb6_select_lbl_N = vb6_VariantString(vb6_BSTR_FromStr(L"..."));
# 而那枚临时 VARIANT 从不 VariantClear => 每执行一次这条 Select, 每枚串字面量标签漏一只 BSTR。
# 20 万轮 x 3 枚字面标签的 PagefileUsage 增量: BASE x64 +31014912 B / x86 +15118336 B;
# 修复后两台都是 0 (夹具 tests/b604_select_label_owner.bas, 判据 S604-leak-ok)。
#
# 口径 (唯一主人): 「这枚标签箱有没有拿下所有权」只问 variantRhsOwnsValue (账 #305 §B164/§B167
# 那位主人) —— 它说 owns 的 (字面量 / 现造的调用 / 对象档取了引用) 在 if 链之后 vb6_VariantClear;
# 它说 borrowed 的 (裸串变量 vb6_VariantString(s) / Variant 槽的浅别名 vb6_VariantFromValue(v))
# 一个都不许清 —— 清下去放的就是别人还活着的那只 BSTR / 那个 Dispatch。
#
# 规则 (每条都对应一个会让本哨兵红的假改动, 见 .local/k321/nc_321.py):
#   R1 标签箱声明只此一处, 且装箱仍只问 wrapVariantValue (不另开打包器);
#   R2 所有权问句只此一处, 且问的是 variantRhsOwnsValue (不许按文本/按表另答);
#   R3 清收循环与清收那一行各只一处, 整份文件的 VariantClear 只此一枚;
#   R4 本文件里不许出现自己的串/变体打包器或 BSTR 释放器 (那会把 borrowed 档一起放掉);
#   R5 位置序: 声明 -> if 链 -> 清收 (清收落进 Case 体就等于没清);
#   R6 扫描地板 (文件长度 + VarCmp 站点数), 扫描瞎了不许报干净;
#   R7 夹具 / 期望 / 本哨兵自己都已在 run_tests.ps1 登记。
# 本文件里不写反斜杠字面量: 路径走 [IO.Path]::DirectorySeparatorChar。
# 输出 ASCII (控制台 GBK, 中文读数重定向后不可 grep)。文件必须 UTF-8 BOM + CRLF。
#
# 用法: powershell -File scripts\check_select_label_owner.ps1 [-RepoRoot <树根>]   (PASS = exit 0)

param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$fail = @()

$sep = [IO.Path]::DirectorySeparatorChar
function Load([string]$rel) {
    $p = Join-Path $RepoRoot ($rel.replace('/', $sep))
    if (-not (Test-Path $p)) { Write-Output ('FAIL: missing ' + $p); exit 1 }
    return [IO.File]::ReadAllText($p).Replace("`r`n", "`n")
}
# 只认**代码文本**: 每行砍掉 `//` 之后的部分。本哨兵自己的注释里就写着病灶那两形
# (VariantString / VariantFromValue), 不砍的话 R4 对着一句注释报红。
function CodeOf([string]$txt) {
    $out = New-Object System.Text.StringBuilder
    foreach ($ln in $txt.Split([char]10)) {
        $cut = $ln.IndexOf('//')
        if ($cut -ge 0) {
            [void]$out.Append($ln.Substring(0, $cut))
        } else {
            [void]$out.Append($ln)
        }
        [void]$out.Append([char]10)
    }
    return $out.ToString()
}
function C([string]$lit, [string]$txt) {
    return ([regex]::Matches($txt, [regex]::Escape($lit))).Count
}
function Want([string]$rule, [int]$got, [int]$expect, [string]$why) {
    if ($got -ne $expect) {
        $script:fail += ($rule + ' = ' + $got + ' (want ' + $expect + '): ' + $why)
    }
}
function Before([string]$a, [string]$b, [string]$txt, [string]$rule, [string]$why) {
    $ia = $txt.IndexOf($a); $ib = $txt.IndexOf($b)
    if ($ia -lt 0 -or $ib -lt 0 -or $ia -ge $ib) {
        $script:fail += ($rule + ': position order broken (' + $ia + ' vs ' + $ib + '): ' + $why)
    }
}

$SEL = Load 'src/backend/stmt/cgen_select.cpp'
$RUN = Load 'tests/run_tests.ps1'
$FX  = Load 'tests/b604_select_label_owner.bas'
$COD = CodeOf $SEL

# ---- R1 标签箱声明只此一处, 装箱只问 wrapVariantValue ----
Want 'R1 label declaration' (C ('labelDecls.push_back("vb6_VARIANT " + lbl + " = " + wrapVariantValue(ast, expr) + ";");') $COD) 1 '标签箱只许在 labelVariantAddr 里造一次'
Want 'R1 boxing authority' (C 'wrapVariantValue(ast, expr)' $COD) 1 '装箱仍只问 wrapVariantValue 那一处 (另开打包器 = 第二份档位表)'
Want 'R1 decls emitted once' (C 'for (auto& decl : labelDecls) c_.emitLine(decl);' $COD) 1 '声明只许在 if 链之前发一次'

# ---- R2 所有权问句 ----
Want 'R2 ownership question' (C 'if (variantRhsOwnsValue(ast, expr)) labelClears.push_back(lbl);' $COD) 1 'owns/borrowed 只许问 variantRhsOwnsValue 一次'
Want 'R2 clears list' (C 'std::vector<std::string> labelClears;' $COD) 1 '清收名单只许一枚'

# ---- R3 清收循环与清收行 ----
Want 'R3 clear loop' (C 'for (auto& lblClear : labelClears) {' $COD) 1 '清收循环只许一处'
Want 'R3 clear line' (C ('c_.emitLine("vb6_VariantClear(&" + lblClear + ");");') $COD) 1 '清收那一行只许拼一处'
Want 'R3 one VariantClear in file' (C 'vb6_VariantClear(' $COD) 1 '本文件里 VariantClear 只此一枚 (另开一站 = 第二个释放主人)'

# ---- R4 本文件不许自造打包器/释放器 ----
Want 'R4 no own VariantString' (C 'vb6_VariantString(' $COD) 0 '标签箱的打包归 wrapVariantValue, 这里不许另发'
Want 'R4 no own VariantFromValue' (C 'vb6_VariantFromValue(' $COD) 0 '同上: 这里不许另发浅别名箱'
Want 'R4 no BSTR_Free' (C 'vb6_BSTR_Free(' $COD) 0 'VARIANT 箱要用 VariantClear; 直接 BSTR_Free 会把别的档 (Dispatch/x86 DECIMAL) 漏掉'

# ---- R5 位置序 ----
Before 'for (auto& decl : labelDecls) c_.emitLine(decl);' 'c_.emitLine("if (" + clauseConds[ci] + ") {");' $COD 'R5 decls before chain' '声明必须在 if 链之前 (否则第二枚标签的声明落在 } 与 else if 之间 = 非法 C)'
Before 'c_.emitLine("if (" + clauseConds[ci] + ") {");' 'for (auto& lblClear : labelClears) {' $COD 'R5 chain before clears' '清收必须在 if 链(含 Case 体)之后 —— 放到链前就是边比边放, 比的还是已释放的块'
Before 'if (!firstVariant) {' 'for (auto& lblClear : labelClears) {' $COD 'R5 clears after chain close' '清收要落在最后一条 Case 的闭括号之后'

# ---- R6 扫描地板 ----
if ($SEL.Length -lt 8000) { $fail += ('R6 scan floor: cgen_select.cpp is ' + $SEL.Length + ' chars (want >=8000) -- 扫描瞎了却报干净') }
if ((C 'vb6_VarCmp' $COD) -lt 3) { $fail += ('R6 VarCmp floor: ' + (C 'vb6_VarCmp' $COD) + ' site(s) (want >=3) -- 混合档比较出口被搬走或改名了') }
if ((C 'selectCaseLabelClass' $COD) -lt 2) { $fail += 'R6 label-class authority: selectCaseLabelClass 的定义与调用都要在' }

# ---- R7 登记面 ----
Want 'R7 x64 cell' (C ('Add-BasTest "b604_select_label_owner"') $RUN) 1 '默认架构那一片'
Want 'R7 x86 cell' (C ('Add-BasTest "b604_select_label_owner_x86"') $RUN) 1 'x86 那一片 (两台各自的 PagefileUsage 读数不同: BASE 31MB / 15MB)'
Want 'R7 fixture leak arms' (C ('Debug.Print "S604-leak-ok=') $FX) 2 '夹具里 Y/N 两条读数都要在 (只钉 Y 的话把判据删掉也算绿)'
Want 'R7 fixture done witness' (C ('Debug.Print "S604-DONE"') $FX) 1 '夹具末尾的 DONE 证人'
Want 'R7 sentinel registered' (C 'check_select_label_owner.ps1' $RUN) 1 '本哨兵在 run_tests.ps1 里被调用一次'
Want 'R7 sentinel function' (C 'function Test-SelectLabelOwner' $RUN) 1 '那道静态检查自己有一处定义'
Want 'R7 sentinel called' ([regex]::Matches($RUN, '(?m)^\s*if \(Enter-StaticShard\) \{ Test-SelectLabelOwner \}\s*$').Count) 1 '那条裸调用自己有一行 (没有它, 定义只是死码)'

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Output ('FAIL ' + $m) }
    Write-Output ('FAIL check_select_label_owner: ' + $fail.Count + ' rule(s) broken')
    exit 1
}
Write-Output ('PASS check_select_label_owner: decls=' + (C 'labelDecls.push_back(' $COD) +
    ' owns=' + (C 'variantRhsOwnsValue(ast, expr)' $COD) +
    ' clears=' + (C 'vb6_VariantClear(' $COD) +
    ' cmps=' + (C 'vb6_VarCmp' $COD) +
    ' cells=' + (C 'Add-BasTest "b604_select_label_owner' $RUN) +
    ' needles=' + (C 'Debug.Print "S604-' $FX))
exit 0
