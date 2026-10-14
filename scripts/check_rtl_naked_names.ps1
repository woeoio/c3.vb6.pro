# check_rtl_naked_names.ps1 - 账 #220 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: RTL 里曾写着 `const int32_t B = 1; const int32_t BF = 2;` (外加头文件里的
# 两行 extern), 用途是让 parser 原样发射的 Picture.Line 语法旗标 `, B` / `, BF` 有个落脚处。
# 生成的模块 C 会 #include 那批 RTL 头, 而用户模块级变量在 C 里也是**裸名** —— 于是
# `Public B As Long` 直接撞成 C2373 重定义 + C2166 给 const 赋值, 连 exe 都出不来
# (改前探针实测: BUILD-RC=1 / 5 条诊断 / no exe)。这类名字与用户名字空间是**共享**的:
# 头里 extern 的在编译期撞, .c 里非 static 定义的在链接期撞 (LNK2005), 只有 static 的不撞。
#
# 口径: RTL 不许导出「裸名 = VB6 合法标识符」的文件作用域数据全局。旗标改由 parser 在
# Line 的 style 位置折成字面量 (parser_expr_postfix.cpp 一处), RTL 不再需要名字。
#
# 规则 (改坏了会红, 不是装饰):
#   N1  B / BF 这两枚裸名全局在 RTL 里必须彻底没有 (定义 0 + extern 0)
#   N2  非 static 的裸名文件作用域数据全局 = 一份钉死的名单 (多一枚就红; 名单要缩小
#       必须在同一次提交里改这里 —— 例如账 #219 把 Changed 归到 vb6_PropertyPage_Changed)
#   N3  折旗标只许一个地方: style 位守卫 1 / B,C,F 三个字母位各 1 / 交出的字面量 1
#   N4  折出来的必须是位掩码本身 (intValue = styleBits 一处), 且旗标名字不许进 AST (0)
#   N5  RTL 解释的必须是同三位 (style & 1 / & 2 / & 4 各 >=1) 且 vb6_ControlLine 定义 1 份
#       —— parser 与 RTL 是**一张表**的两头, 谁单独改编号就红
#   N6  Line 的改道必须有表有码头: 声明 1 + 定义 1 + 码头至少查一次
#       (只接一头就是账 #143 那条"编得过、跑了、什么都没发生")
#   值口径 (账 #221): B=1 矩形 / C=2 椭圆 / F=4 填充 ⇒ BF=5、CF=6。
#   订正: 账 #220 那一版沿用被删的两枚 RTL 全局写的是 B=1 / BF=2, 那一版既没有 C 也没有 F 的
#   落脚点, 本档随 #221 一起收成字母位。
# 规则扫的是标量/指针档 (intN_t / long / char / double / BSTR / VARIANT / VB6_* 等)。
#
# 扫的是标量/指针档 (intN_t / long / char / double / BSTR / VARIANT / VB6_* 等)。
#
# 用法:  pwsh -File scripts\check_rtl_naked_names.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$rtlDir = Join-Path $root 'src\rtl'
$parserRel = Join-Path $root 'src\parser\parser_expr_postfix.cpp'
if (-not (Test-Path -LiteralPath $rtlDir)) { Write-Host 'FAIL src\rtl missing' -ForegroundColor Red; exit 1 }
if (-not (Test-Path -LiteralPath $parserRel)) { Write-Host 'FAIL parser_expr_postfix.cpp missing' -ForegroundColor Red; exit 1 }

$types = '(?:int8_t|int16_t|int32_t|int64_t|uint8_t|uint16_t|uint32_t|uint64_t|int|long|short|unsigned\s+long|char|double|float|void|BSTR|VARIANT|VARIANTARG|HRESULT|BOOL|BYTE|WORD|DWORD|LPARAM|WPARAM|SIZE_T|size_t|HINSTANCE|HMODULE|HFONT|HDC|HWND|HBITMAP|HPEN|HBRUSH|HANDLE|UINT|INT|LPWSTR|WCHAR|GUID|CLSID|REFCLSID|ITypeLib|ITypeInfo|IDispatch|IUnknown|VB6_[A-Za-z_]\w*|vb6_[A-Za-z_]\w*)'
$globalPat = '^(?:extern\s+)?(?:static\s+)?(?:const\s+)?' + $types + '\s*(?:\*+\s*)?([A-Za-z_]\w*)\s*(?:\[[^\]]*\])?\s*(?:=[^;]*)?;'
$allowPat = '^(vb6_|VB6_|c3_|C3_)'

$rtlText = @{}
Get-ChildItem -Path $rtlDir -Recurse -Include *.c, *.h, *.inc | ForEach-Object {
    $rtlText[$_.FullName] = [System.IO.File]::ReadAllText($_.FullName)
}

# N1: B / BF / Changed 这些裸名不许回到 RTL
$n1 = 0
foreach ($t in $rtlText.Values) {
    $n1 += @([regex]::Matches($t, '^(?:extern\s+)?(?:static\s+)?(?:const\s+)?(?:int32_t|uint32_t|int16_t|int)\s+\*?\s*(?:B|BF|Changed)\s*(?:\[[^\]]*\])?\s*(?:=[^;]*)?;', 'Multiline')).Count
}
if ($n1 -ne 0) { $bad += ('N1 naked B/BF/Changed globals back in RTL ' + $n1 + ' times (want 0 - the flag is folded by the parser, and PropertyPage.Changed has vb6_PropertyPage_Changed)') }

# N2: 非 static 的裸名文件作用域数据全局, 名单钉死
$found = @{}
foreach ($kv in $rtlText.GetEnumerator()) {
    foreach ($m in [regex]::Matches($kv.Value, $globalPat, 'Multiline')) {
        $decl = $m.Groups[0].Value
        $name = $m.Groups[1].Value
        if ($decl -match '^static\b') { continue }
        if ($name -match $allowPat) { continue }
        if (-not $found.ContainsKey($name)) { $found[$name] = 0 }
        $found[$name] += 1
    }
}
# 名单历史（账 #219）：这一格原本只钉住 4 枚，本机重测真实是 12 枚 —— 旧类型表没有 `int`、没有小写 `vb6_` 结构体档，
#       也不扫 `.inc`，于是 `vb6_HostObjRec g_ho[N]`、`vb6_UCControls* g_uc_current`、`vb6_UCRecord g_uc_recs[N]` 全看不见
#       （钉着一张漏提的图 = 名单里那四枚一直在替八枚隐身）。本刀把 10 枚 UC/axsite 内部的裸名换成 vb6_ 前缀
#       （`g_hoCount`→`vb6_ucHoCount`、`g_hInstance`→`vb6_hInstance` 等），名单缩到下面这三枚。
#       这三枚是 cgen 与 RTL 的**契约名**（`cgen_util_dllentry_exports.inc` 直接把 `g_vb6_coclasses` / `g_vb6_coclassCount`
#       发进 dll_entry.c，`cgen_localdecl.cpp` 发 `g_vb6_optionCompareText`），改名要两头同批 + 发码 A/B 归因，不在这一刀里；
#       它们自带 vb6_ 印记，用户手打出来的概率近零。名单要再缩小必须在同一次提交里改这里。
$pin = @('g_vb6_coclassCount', 'g_vb6_coclasses', 'g_vb6_optionCompareText')
$extra = @($found.Keys | Where-Object { $pin -notcontains $_ })
$missing = @($pin | Where-Object { -not $found.ContainsKey($_) })
if ($extra.Count -gt 0) { $bad += ('N2 new naked RTL global: ' + ($extra -join ', ') + ' - VB lets a module name a variable that, so this collides') }
if ($missing.Count -gt 0) { $bad += ('N2 pinned census no longer matches: ' + ($missing -join ', ') + ' disappeared - update the pin in the same commit that removes it') }

# N3 + N4: 折旗标那一处 —— 账 #221 起按**字母位**折 (B=1 / C=2 / F=4, 所以 BF=5、CF=6),
# 不再枚举两个词; 名字一律不进 AST 这一条没变 (那正是撞车的成因)。
$ptxt = [System.IO.File]::ReadAllText($parserRel)
$guard = @([regex]::Matches($ptxt, 'if\s*\(\s*trailingIdx\s*==\s*1')).Count
$bitB = @([regex]::Matches($ptxt, "ch\s*==\s*'b'\s*\)\s*\{\s*styleBits\s*\|=\s*1")).Count
$bitC = @([regex]::Matches($ptxt, "ch\s*==\s*'c'\s*\)\s*\{\s*styleBits\s*\|=\s*2")).Count
$bitF = @([regex]::Matches($ptxt, "ch\s*==\s*'f'\s*\)\s*\{\s*styleBits\s*\|=\s*4")).Count
$litInt = @([regex]::Matches($ptxt, 'LiteralKind::Integer,\s*bitText')).Count
$hand = @([regex]::Matches($ptxt, 'intValue\s*=\s*styleBits')).Count
$idName = @([regex]::Matches($ptxt, 'IdentifierExpr[^;]*styleBits')).Count
if ($guard -ne 1) { $bad += ('N3 style-slot guard appears ' + $guard + ' times (want exactly 1 - one fold point)') }
if ($bitB -ne 1) { $bad += ('N3 the B bit is folded ' + $bitB + ' times (want 1)') }
if ($bitC -ne 1) { $bad += ('N3 the C bit is folded ' + $bitC + ' times (want 1) - 账 #220 那一版没这一形') }
if ($bitF -ne 1) { $bad += ('N3 the F bit is folded ' + $bitF + ' times (want 1) - 同上') }
if ($litInt -ne 1) { $bad += ('N3 literal handoff ' + $litInt + ' times (want 1)') }
if ($hand -ne 1) { $bad += ('N4 intValue=styleBits appears ' + $hand + ' times (want 1 - the folded value must BE the bit mask)') }
if ($idName -ne 0) { $bad += ('N3 the flag name is back in the AST ' + $idName + ' times (want 0) - that is what created the collision') }

# N5: RTL 解释的必须是**同三位** —— parser 与 RTL 是一张表, 不是两边各自编号
$rtlCtrl = [System.IO.File]::ReadAllText((Join-Path $root 'src\rtl\core\vb6forms\vb6forms_ctrl.c'))
foreach ($bit in @('1', '2', '4')) {
    $one = @([regex]::Matches($rtlCtrl, '\(\s*style\s*&\s*' + $bit + '\)')).Count
    if ($one -lt 1) { $bad += ('N5 RTL never interprets style bit ' + $bit + ' - parser and RTL must be one table') }
}
$lineFn = @([regex]::Matches($rtlCtrl, 'void vb6_ControlLine\(')).Count
if ($lineFn -ne 1) { $bad += ('N5 vb6_ControlLine defined ' + $lineFn + ' times (want 1)') }

# N6: Line 的改道必须有表有码头 (缺一头就是"编得过、跑了、什么都没画", 本线踩过三次)
$helpersTxt = [System.IO.File]::ReadAllText((Join-Path $root 'src\backend\detail\util\cgen_helpers.inc'))
$utilTxt = [System.IO.File]::ReadAllText((Join-Path $root 'src\backend\cgen_util_ctrl.cpp'))
$withmTxt = [System.IO.File]::ReadAllText((Join-Path $root 'src\backend\detail\expr\cgen_expr_call_callee_withm.inc'))
$cvDecl = @([regex]::Matches($helpersTxt, 'std::string controlCanvasMethod')).Count
$cvDef = @([regex]::Matches($utilTxt, 'CCodeGen::controlCanvasMethod')).Count
$cvUse = @([regex]::Matches($withmTxt, 'controlCanvasMethod\(')).Count
if ($cvDecl -ne 1) { $bad += ('N6 canvas table declared ' + $cvDecl + ' times (want 1)') }
if ($cvDef -ne 1) { $bad += ('N6 canvas table defined ' + $cvDef + ' times (want 1)') }
if ($cvUse -lt 1) { $bad += ('N6 the withm docket never consults the canvas table (' + $cvUse + ') - Line falls back to the silent COM path') }
# 成员侧那一头也必须给标记: 只接调用侧 = 成员先被兜底当 COM 属性取走, 码头根本收不到标记
# (这一刀第一次就红在这一条上 —— 表与码头都写好了, 产物里照旧是 ComGetObjectProp(L"Line")+Item)
$memberTxt = [System.IO.File]::ReadAllText((Join-Path $root 'src\backend\detail\expr\cgen_expr_member_form_builtin.inc'))
$cvMember = @([regex]::Matches($memberTxt, 'controlCanvasMethod\(')).Count
if ($cvMember -lt 1) { $bad += 'N6 the member side never marks canvas methods - Line is read as a COM property and the docket never sees a marker' }

# N7: 语义层放行"裸写的文档成员"这一格只许一处 (账 #219) —— 谓词 定义 1 / 声明 1 / 调用 1,
#     调用点必须带 !memberObjCtx_ (限定符位由 isDocumentHostObject 管, 两支不许重叠);
#     而名字本身不归语义层管 —— 它只许问那张宿主伪成员表 (hostPseudoBareEligible),
#     自己抄一份成员名清单就是本账删掉的那第二个权威。

$semUtil = [System.IO.File]::ReadAllText((Join-Path $root 'src\semantics\semantic_analyzer_util.cpp'))
$semHdr = [System.IO.File]::ReadAllText((Join-Path $root 'src\semantics\semantic_analyzer.hpp'))
$semExpr = [System.IO.File]::ReadAllText((Join-Path $root 'src\semantics\semantic_analyzer_expr.cpp'))
$bareDef = @([regex]::Matches($semUtil, 'bool SemanticAnalyzer::isDocumentBarePseudoMember')).Count
$bareDec = @([regex]::Matches($semHdr, 'bool isDocumentBarePseudoMember')).Count
$bareUse = @([regex]::Matches($semExpr, 'isDocumentBarePseudoMember\(node\.name\)')).Count
$bareGate = @([regex]::Matches($semExpr, '!memberObjCtx_\s*&&\s*isDocumentBarePseudoMember')).Count
$bareTbl = @([regex]::Matches($semUtil, 'hostPseudoBareEligible\(')).Count
$bareName = 0
foreach ($t in @($semUtil, $semHdr, $semExpr)) { $bareName += @([regex]::Matches($t, 
    '"(changed|hdc|hwnd|scalewidth|scaleheight|scalemode|containerhwnd|enabled|autoredraw)"')).Count }
if ($bareDef -ne 1) { $bad += ('N7 isDocumentBarePseudoMember defined ' + $bareDef + ' times (want 1)') }
if ($bareDec -ne 1) { $bad += ('N7 isDocumentBarePseudoMember declared ' + $bareDec + ' times (want 1)') }
if ($bareUse -ne 1) { $bad += ('N7 bare-document-member exemption used ' + $bareUse + ' times (want exactly 1)') }
if ($bareGate -ne 1) { $bad += ('N7 the exemption is not gated by !memberObjCtx_ (found ' + $bareGate + ', want 1) - the two exemptions must stay disjoint') }
if ($bareTbl -ne 1) { $bad += ('N7 semantics asks the host-pseudo table ' + $bareTbl + ' times (want 1) - 0 means the name list is back in semantics, which is the second authority this account deleted') }
if ($bareName -ne 0) { $bad += ('N7 semantics copies member-name literals again (' + $bareName + ') - the only authority for those names is kHostPseudoRows') }

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
    exit 1
}
Write-Host ('PASS RTL naked-name guard: B/BF ' + $n1 + ', census ' + ($found.Keys.Count) +
    ' pinned names, fold bits ' + $bitB + '+' + $bitC + '+' + $bitF + ', handoff ' + $litInt + '+' + $hand +
    ', flag names in AST ' + $idName + ', RTL ' + $lineFn + '/bits, ' +
    'canvas ' + $cvDecl + '+' + $cvDef + '+' + $cvUse + '+' + $cvMember + ' ' +
    'bareDoc ' + $bareDef + '/' + $bareDec + '/' + $bareUse + '/' + $bareGate + '/tbl' + $bareTbl + '/names' + $bareName)
exit 0
