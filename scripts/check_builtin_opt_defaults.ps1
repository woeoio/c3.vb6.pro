# check_builtin_opt_defaults.ps1 —— 「内置出口省略实参时补哪一档」那张表与 RTL / 语义名单对账（账 #305 §B153，第四十一刀）
#
# 为什么要有这一道：收表之前，这个决定在 src/backend/detail/expr/cgen_expr_call_builtin_fixup.inc 里
# 有 **28 位主人**（一枚内置函数一段 `if (callee == "vb6_X") { ... argList += ", 1"; }`），而答案本身
# 住在两处别人写的东西里 —— VB6 的形参表（哪一档可省、缺省值是什么）与 RTL 的 C 声明（一共几档）。
# 三份从不汇合 ⇒ 三个方向都会走散，且都贵：
#   少一档 = 发出来的实参比 RTL 声明少 ⇒ C2198「参数太少」（实物：§B149 的 DeleteSetting 两参形，
#          以及 086 / 110l / 158j 那一笔一条一个 bug 的补洞史）。
#   多一档 = C2197「实参太多」，更坏是位置错接（§B121 那格 InStr 三参形把 needle 接到 start 槽，实测段错误）。
#   补错值 = 静默答错（Mid 省略 length 曾补 0 ⇒ czUI 的 HexToColor("#243447") 解析成纯黑，整屏颜色错）。
# 所以这道针四个方向都要问：表 ↔ RTL 的档数、表 ↔ 语义名单的可选档、表 ↔ 发码侧不许有第二份抄法、
# 留在文件里那些非补档的段必须逐名可枚举（新增一形硬编码补档就红）。
#
# 判据面（任何一条不齐就红）：
#   R0 表里每一枚 callee 都必须在 RTL 里有声明（认不出 = 死名/拼错的名字躺在表里）。
#   R1 每行的数组长度 == RTL 声明的形参个数（不齐 = 补档会多补或少补）。
#   R2 每一行至少一枚非 nullptr 的缺省档（全 nullptr = 死行）。
#   R3 非 nullptr 的档必须构成数组的**后缀**（中间有洞 = 省略前一档时后面补不出来，实参整体错位）。
#   R4 表里登记过的名字，不许在同一个文件里另有 `callee == "vb6_X"` 且往 argList 追加实参的段（第二份抄法）。
#   R5 语义侧 src/semantics/builtin/builtin_funcs.inc 若给这枚名字登记了形参表，它的**可选档集合**必须
#      等于表里非 nullptr 的档集合（不齐 = 两份名单从不汇合；例外要写理由，见 $leadingDefault）。
#   R6 census 不许读空：rows >= 25、每行引用的数组**都真的存在**（不要求数组数 >= 行数 —— 同一张表被几家共用是合法的，见下面 §1 的注）、RTL 形参数至少读出 300 枚、语义形参表至少读出 80 行。
#   R7 留在发码侧按实参个数做决定的那些名字（改写 callee / 重排实参 / 截断，不是补档）必须**恰好**是
#   R8 表里一枚出口只许一行（重复行会让前一档的缺省答案静默失效，本机负控实测）。
#      下面这一份名单 —— 多一名 = 新写了一形硬编码，少一名 = 名单过期了。
#
# 用法: powershell -File scripts\check_builtin_opt_defaults.ps1   (PASS = exit 0)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$fixInc = Join-Path $Root 'src\backend\detail\expr\cgen_expr_call_builtin_fixup.inc'
$semInc = Join-Path $Root 'src\semantics\builtin\builtin_funcs.inc'
foreach ($p in @($fixInc, $semInc)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL: 读不到 ' + $p); exit 1 }
}

function Strip-CComments([string]$t) {
    $t = [regex]::Replace($t, '/\*[\s\S]*?\*/', ' ')
    $t = [regex]::Replace($t, '(?m)//.*$', ' ')
    return $t
}

# 取一段花括号平衡的块体 (start 指向块首 '{' 的位置)
function Get-BlockBody([string]$t, [int]$start) {
    $depth = 0
    for ($j = $start; $j -lt $t.Length; $j++) {
        $ch = $t[$j]
        if ($ch -eq '{') { $depth++ }
        elseif ($ch -eq '}') {
            $depth--
            if ($depth -eq 0) { return $t.Substring($start + 1, $j - $start - 1) }
        }
    }
    return $null
}

$fix = Strip-CComments ([IO.File]::ReadAllText($fixInc))

# ---- 1) 那张表: 行 + 每行的档 ----
$rows = @{}
$rowCount = @{}
foreach ($m in [regex]::Matches($fix, '\{\s*"(vb6_[A-Za-z0-9_]+)"\s*,\s*(kOD_[A-Za-z0-9_]+)\s*,\s*sizeof\(')) {
    $nm = $m.Groups[1].Value
    $rows[$nm] = $m.Groups[2].Value
    if ($rowCount.ContainsKey($nm)) { $rowCount[$nm] = $rowCount[$nm] + 1 } else { $rowCount[$nm] = 1 }
}
$arrays = @{}
foreach ($m in [regex]::Matches($fix, 'static const char\* const (kOD_[A-Za-z0-9_]+)\s*\[\]\s*=\s*\{(?<b>[^}]*)\};')) {
    $slots = @()
    foreach ($e in ($m.Groups['b'].Value -split ',')) {
        $e = $e.Trim()
        if ($e -eq '') { continue }
        $slots += $e
    }
    $arrays[$m.Groups[1].Value] = $slots
}
if ($rows.Count -lt 25) { Write-Host ('FAIL R6: 表只读出 ' + $rows.Count + ' 行 (<25) —— 扫描读空, 不是一条绿着的判据'); exit 1 }
# ⚠ 这里曾经写的是 `$arrays.Count -lt $rows.Count`，**比错了两个量**：`$arrays` 数的是**去重后
#   的数组个数**，`$rows` 数的是**调用方个数**。表里**一行一枚私数组**本来只是巧合，不是规矩 ——
#   账 #428 把 `vb6_FormatCurrency` / `vb6_FormatPercent` 登记到**既有的 `kOD_FormatNum`** 上
#   （三家 C 签名逐字同形，共用一张表才是对的），于是 rows 31 / arrays 29，那条判据立刻红了。
#   ⇒ 同一张表被几家共用是**合法且应当鼓励**的写法，不能拿来当「扫描读空」的信号。
#   真正的「读空」信号改成两件事：数组本身读不出多少条（<25），以及**每行引用的数组必须真的存在**
#   （下面逐行判的那条才是权威，R6 行找不到数组）。这里先做一次前置的整体断言，免得逐行那条
#   在 0 行时静默通过。
$dangling = @($rows.Values | Where-Object { -not $arrays.ContainsKey($_) } | Sort-Object -Unique)
if ($dangling.Count -gt 0) {
    Write-Host ('FAIL R6: ' + $dangling.Count + ' 枚数组被引用但没定义: ' + ($dangling -join ', ')) -ForegroundColor Red
    exit 1
}
if ($arrays.Count -lt 25) { Write-Host ('FAIL R6: 只读出 ' + $arrays.Count + ' 枚数组 (<25) —— 数组扫描读空'); exit 1 }

# ---- 2) RTL: 每一枚 vb6_X( 的形参个数 (取第一条能认的声明) ----
# 认法与 check_rtl_proto_arity.ps1 的 $HeadRx 同形: 返回类型 + 名字 + '('; 跨行的那几条按「认不出」跳过。
$HeadRx = [regex]'^(?:[A-Za-z_][A-Za-z0-9_]*\s+|[A-Za-z_][A-Za-z0-9_]*\*+\s*|\s)*(vb6_[A-Za-z0-9_]+)\s*\('
$arity = @{}
foreach ($f in @(Get-ChildItem -Path (Join-Path $Root 'src\rtl') -Recurse -File -Include *.c,*.h)) {
    $t = Strip-CComments ([IO.File]::ReadAllText($f.FullName))
    foreach ($ln in ($t -split "`n")) {
        $ln = $ln.TrimEnd()
        $mm = $HeadRx.Match($ln)
        if (-not $mm.Success) { continue }
        $nm = $mm.Groups[1].Value
        if ($arity.ContainsKey($nm)) { continue }
        $i = $ln.IndexOf('(', $ln.IndexOf($nm))
        if ($i -lt 0) { continue }
        $depth = 0; $commas = 0
        for ($j = $i; $j -lt $ln.Length; $j++) {
            $ch = $ln[$j]
            if ($ch -eq '(') { $depth++ }
            elseif ($ch -eq ')') { $depth--; if ($depth -eq 0) { break } }
            elseif ($ch -eq ',' -and $depth -eq 1) { $commas++ }
        }
        if ($depth -ne 0) { continue }
        $closeAt = -1; $d2 = 0
        for ($j = $i; $j -lt $ln.Length; $j++) {
            if ($ln[$j] -eq '(') { $d2++ } elseif ($ln[$j] -eq ')') { $d2--; if ($d2 -eq 0) { $closeAt = $j; break } }
        }
        if ($closeAt -lt 0) { continue }
        $inner = $ln.Substring($i + 1, $closeAt - $i - 1)
        if ($inner.Trim() -match '^(void)?$') { $arity[$nm] = 0 } else { $arity[$nm] = $commas + 1 }
    }
}
if ($arity.Count -lt 300) { Write-Host ('FAIL R6: RTL 形参数只数出 ' + $arity.Count + ' 枚 (<300) —— 扫描读空'); exit 1 }

# ---- 3) 语义名单: builtin 的形参表与可选档 ----
$sem = @{}
foreach ($m in [regex]::Matches((Strip-CComments ([IO.File]::ReadAllText($semInc))),
                               'addBuiltinFuncWithParams\("(?<n>[A-Za-z0-9_]+)"[^;]*?\);')) {
    $ps = @([regex]::Matches($m.Value, '\{\s*"[A-Za-z0-9_]+"\s*,\s*[^{}]+?,\s*(?:true|false)\s*,\s*(?<o>true|false)\s*\}'))
    if ($ps.Count -eq 0) { continue }
    $opt = @()
    for ($k = 0; $k -lt $ps.Count; $k++) { if ($ps[$k].Groups['o'].Value -eq 'true') { $opt += $k } }
    $sem[$m.Groups['n'].Value] = @{ n = $ps.Count; opt = $opt }
}
if ($sem.Count -lt 80) { Write-Host ('FAIL R6: 语义形参表只读出 ' + $sem.Count + ' 行 (<80) —— 扫描读空'); exit 1 }

# 例外: VB6 手册里**第一档本身可省**的那一族 —— 缺省答案在表里, 而语义名单把第一档标成必填
# (isOptional=false)。那不是表错, 是语义那本名单的这一格与 VB6 不齐 (已量到; 它今天不影响发码,
# 因为通用 padding 对 builtin 压根不吃 isOptional)。走例外必须写清手册依据。
$leadingDefault = @{
    'vb6_Dir'           = 'VB6: Dir[(pathname[, attributes])] —— 第一档可省 (空 = 继续上次搜索)'
    'vb6_CurDir'        = 'VB6: CurDir[(drive)] —— 第一档可省 (缺省取当前盘)'
    'vb6_LoadPictureEx' = '语义侧登记的名字是 LoadPicture, 与本行的 RTL 出口名不同, 无从按名对账'
}

# R7 的名单: 留在发码侧按实参个数做决定、但**不是补档**的那些名字 (改写 callee / 重排实参 / 截断 /
# 整串替换), 逐名一条理由。
$notPadding = @{
    'vb6_LoadResString'  = '资源三枚共用: 按 args 逐档拼 BSTR 表, 整串重建 argList'
    'vb6_LoadResData'    = '同上'
    'vb6_LoadResPicture' = '同上'
    'vb6_LoadForm'       = '按首参形状改成 &vb6_arr_<名> + 下标 (换实参本体, 不是补档)'
    'vb6_UnloadForm'     = '同上'
    'vb6_UBound'         = '首参可能是 ParamArray / Variant 载体: 会改 callee 名与实参本体。它那一档 "补 dimension=1" 与表同口径但**刻意不收进表** —— PA 那支把 argList 重建成一枚实参, 表再补第二档就是 C2197'
    'vb6_LBound'         = '同上'
    'vb6_InStr'          = '双形态 (start 在前 / 字符串在前) 要**重排**实参并按 compare 换 callee = vb6_InStrC'
    'vb6_InStrB'         = '与 InStr 同一段'
    'vb6_ErrRaise'       = '反向那一格: VB6 最多 5 参而 RTL 只吃 3 参 ⇒ 截断'
    'vb6_StrConv'        = '首参是已知 Byte 数组时改 callee = vb6_StrConvFromByteArray'
    'vb6_StrPtr'         = '首参是已知 Byte 数组时改 callee = vb6_SafeArrayDataPtr'
}

$viol = @()

# R4 / R7: 逐段看留在文件里的 `if (... callee == "vb6_X" ...)`
foreach ($m in [regex]::Matches($fix, 'if\s*\(\s*\(?\s*callee == "(vb6_[A-Za-z0-9_]+)"')) {
    $nm = $m.Groups[1].Value
    $open = $fix.IndexOf('{', $m.Index)
    if ($open -lt 0) { continue }
    $body = Get-BlockBody $fix $open
    if ($null -eq $body) { continue }
    $pads = ([regex]::Matches($body, 'argList \+=')).Count
    if ($pads -gt 0) {
        if ($rows.ContainsKey($nm)) { $viol += ('R4 第二份抄法: ' + $nm + ' ( —— 表里登记过了, 发码侧还往 argList 追加 (硬编码 ' + $pads + ' 处)') }
        elseif (-not $notPadding.ContainsKey($nm)) { $viol += ('R4 表外补档: ' + $nm + ' ( —— 新写了一形硬编码补档, 应该收进 kBuiltinOptDefaults 那张表') }
    } else {
        # 一段只做「改 callee / 重排实参 / 截断」而不补档的段: 不管这枚名字在不在表里, 都必须
        # 在 $notPadding 里写清理由 —— 一行两处分本身是允许的 (ErrRaise 既补到 3 档又把 VB6 的
        # 第 4、5 档截掉; StrConv 既补 localeID 又按首参换 callee), 但它必须是**登记过**的第二处。
        if (-not $notPadding.ContainsKey($nm)) {
            $viol += ('R7 名单过期/新增段: ' + $nm + ' ( —— 这一格按实参个数做的事没在 $notPadding 里写理由')
        }
    }
}

foreach ($nm in @($rows.Keys | Sort-Object)) {
    # R8: 一枚出口一行 —— 同一枚名字有两行时, 上面那张哈希表只会留下后写的那一行,
    #     于是前一档的缺省答案**静默失效** (本机负控实测: 加一行重复的 DatePart, R1 一声不出)。
    if ($rowCount[$nm] -gt 1) { $viol += ('R8 重复行: ' + $nm + ' ( —— 表里有 ' + $rowCount[$nm] + ' 行, 只有最后一行生效') }
    $arrName = $rows[$nm]
    if (-not $arrays.ContainsKey($arrName)) { $viol += ('R6 行找不到数组: ' + $nm + ' -> ' + $arrName); continue }
    $slots = $arrays[$arrName]

    # R0 / R1: 与 RTL 的形参表对齐
    if (-not $arity.ContainsKey($nm)) {
        $viol += ('R0 死名: ' + $nm + ' ( —— RTL 里没有这枚声明, 表里躺着一条认不出来的行')
    } elseif ($arity[$nm] -ne $slots.Count) {
        $viol += ('R1 档数不齐: ' + $nm + ' ( 表 ' + $slots.Count + ' vs RTL ' + $arity[$nm])
    }

    # R2 / R3: 至少一枚缺省档, 且缺省档是后缀
    $nonNull = @()
    for ($k = 0; $k -lt $slots.Count; $k++) { if ($slots[$k] -ne 'nullptr') { $nonNull += $k } }
    if ($nonNull.Count -eq 0) { $viol += ('R2 死行: ' + $nm + ' ( —— 一档缺省值都没有') }
    if ($nonNull.Count -gt 0) {
        for ($k = $nonNull[0]; $k -lt $slots.Count; $k++) {
            if ($slots[$k] -eq 'nullptr') { $viol += ('R3 中间有洞: ' + $nm + ' ( 第 ' + $k + ' 档是 nullptr 而它后面还有缺省值'); break }
        }
    }

    # R5: 与语义名单的可选档集合对齐
    $vb = $nm.Substring(4)
    if ($sem.ContainsKey($vb) -and -not $leadingDefault.ContainsKey($nm)) {
        $s = $sem[$vb]
        $want = ($nonNull -join ',')
        $got = ($s.opt -join ',')
        if ($s.n -ne $slots.Count) { $viol += ('R5 形参个数不齐: ' + $nm + ' ( 语义 ' + $s.n + ' vs 表 ' + $slots.Count) }
        elseif ($got -ne $want) { $viol += ('R5 可选档不齐: ' + $nm + ' ( 语义 opt={' + $got + '} vs 表 non-null={' + $want + '}') }
    }
}

# R7 的另一半: $notPadding 里写了理由而文件里已经没这一段的, 也算过期
foreach ($nm in @($notPadding.Keys | Sort-Object)) {
    if (-not [regex]::IsMatch($fix, ('callee == "' + [regex]::Escape($nm) + '"'))) {
        $viol += ('R7 名单里躺着已消失的段: ' + $nm + ' ( —— 段没了就把这一格撤掉, 别让例外名单变成化石')
    }
}

Write-Host ('census: rows=' + $rows.Count + ' arrays=' + $arrays.Count +
            ' rtlArityNames=' + $arity.Count + ' semParamTables=' + $sem.Count +
            ' leadingDefault=' + $leadingDefault.Count + ' notPadding=' + $notPadding.Count)

if ($viol.Count -gt 0) {
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    Write-Host ('FAIL builtin_opt_defaults: ' + $viol.Count + ' 条违规 —— 修法是把表与 RTL/语义名单对账 (改档/改名/撤段), 不是往例外名单里加名字; 真要走例外请连带写清 VB6 手册依据并重跑 §B153 那把语料 A/B (401 份 --emit-c 逐行不变)')
    exit 1
}
Write-Host 'PASS builtin_opt_defaults (R0..R8)'
exit 0
