# check_rtl_proto_arity.ps1 —— 「递几枚实参」这一件事, 四个住所必须给同一个答案
#   (账 #240 钉 RTL 的头与体; 账 #278 §B72 补上发码表、类型 oracle 与产物那两头)
#
# 症状不在本机：vb6_OleCon_Init 的体是 10 参、头里那份原型还停在 6 参，而 cgen 发的是 10 个实参。
# VS2019 的 cl 在 C 模式下压根不诊断「实参过多」（本机两架构 rc=0、出 exe），runner 上那台新 cl
# 报 4 行 error C2197（10-6=4 枚多余实参，一枚一行），门就红在 olecon / olecon_x86 两条上。
# 所以这一族不能靠"本机编一遍"来验，只能对着源码比：头的参数个数 == 体的参数个数。
#
# 判据面：
#   R1 两头都有的名字：声明侧的个数集合必须等于定义侧的个数集合，否则红并列出 文件:行 与两边的数
#      （账 #278 §B128：认「这是一条声明」之前先把注释剥干净 —— `//` 之外还要剥 `/* … */`，
#      因为 RTL 头里大量写法是 `void f(…);   /* 说明 */`，尾巴不是 ';' 就当没声明，头/体两头一起隐身）
#   R2 census 地板：扫到的 .h/.c 文件数 >= 100，且「两头都有」的签名对数 >= 900
#      （实测：§B128 之前 112 份 / 1108 对，之后 127 份 / 1267 对 —— 多出来的那 8 对从前成对隐身）
#      —— 路径写错或正则被改坏时，哨兵不许变成"绿着的空转"
#
#   R3（账 #278 §B72）发码那几张控件方法表（cgen_util_ctrl.cpp 的 controlExit("vb6_…", 个数, outArgc)）
#      三面都得对上：名字在 RTL 头里真有原型且个数相同（查不到原型也红 —— 没有声明就没有担保）、
#      cgen_util_type.cpp 里那张运行时参数表若有同名行则个数相同、行数恰好 23（第十二刀 7 行 +
#      第十三刀的 Winsock 那 8 行 + 第十四刀并进零实参那张表的 8 行（两枚集合 Clear +
#      CommonDialog 那六枚 Show*）；
#   R4 产物那一头：跑一次 --emit-c（只走前端，不起 cl）数 tests/ctrlzero/ZeroForm.frm 里实际发出
#      的调用递了几枚实参，与表里的数对（夹具跑出 22 枚：零/一/换算三张表 6 枚 + Winsock 8 枚
#      + 集合 Clear 2 枚 + CommonDialog Show* 6 枚）
#      —— 这一面守的正是「RTL 加了形参、发码仍递旧的个数」
#      那条本机只 warning C4020、runner 上新 cl 才升 error C2197 的形状（§B72 立项的理由）；
#   R5 那些出口名在 src/backend 别处再出现成字符串字面量 = 同一个事实两份答案 ⇒ 红
#      （两处例外：表自己那份文件，与类型 oracle —— 后者由 R3 的 TABLE-VS-TYPEORACLE 对账）。
#      （比的是「名字"」与「名字("」两种字面量开头 —— 码头习惯把左括号拼进同一条串，
#      只比「整条等于名字」会漏掉一半站点：第十三刀的实测数就因此少报过。）
# 边界（宁可漏报不误报，故此处只比个数）：
#   * 只有**第 0 列开始**的行才算签名 —— 调用点都缩进在函数体里，这一条同时把它们排干净；
#   * 参数表里出现认不出的形态（数组 / 函数指针 / 默认值 / 串） ⇒ 跳过该条，不计入也不报红；
#   * 不比类型拼写：同一函数的头与体写成 const X* 与 X* 是合法的，硬比只会造噪声。

param(
    [string]$Root = ''      # 负控用: 把整棵树指到副本上跑
)
$ErrorActionPreference = 'Stop'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }
$Rtl  = Join-Path $Root 'src\rtl'

$HeadRx = [regex]'^[A-Za-z_][A-Za-z0-9_]*(?:[\s\*]+[A-Za-z_][A-Za-z0-9_]*)*?[\s\*]+(vb6_[A-Za-z0-9_]+)\s*\('
$ItemRx = [regex]'^([A-Za-z_][A-Za-z0-9_]*\s*\*{0,3}\s*)?[A-Za-z_][A-Za-z0-9_]*$'

# 账 #278 §B128: 注释剥离的唯一出口。三种都要处理：`//` 到行尾、成对的 `/* … */`、
# 以及"本行起了 /* 却在这一行外面才关"的那种（跨行块 —— 从 /* 起整段切掉）。
# 从前这里只剥 `//`，而续行才剥成对的块注释 ⇒ 单行声明带行尾块注释时 $tail 以 `*/` 结尾，
# R1 的「声明」判定（尾巴是 ';'）当场失配，头/体两头同时隐身。
function Strip-Cmt([string]$S) {
    $t = ($S -replace '/\*.*?\*/', '') -replace '//.*$', ''
    $ix = $t.IndexOf('/*')
    if ($ix -ge 0) { $t = $t.Substring(0, $ix) }
    return $t
}

function Get-Balance([string]$S) {
    $n = 0
    foreach ($ch in $S.ToCharArray()) {
        if ($ch -eq '(') { $n += 1 }
        elseif ($ch -eq ')') { $n -= 1 }
    }
    return $n
}

function Get-Arity([string]$Params) {
    $p = ($Params -replace '\s+', ' ').Trim()
    if ($p -eq '' -or $p -eq 'void') { return 0 }
    if ($p -eq '...') { return -1 }
    $depth = 0
    $cur = ''
    $items = @()
    foreach ($ch in $p.ToCharArray()) {
        if ($ch -eq '(' -or $ch -eq '[') { $depth += 1 }
        elseif ($ch -eq ')' -or $ch -eq ']') { $depth -= 1 }
        if ($ch -eq ',' -and $depth -eq 0) { $items += $cur; $cur = '' } else { $cur += $ch }
    }
    $items += $cur
    foreach ($it in $items) {
        $t = (($it -replace '\bconst\s+', ' ') -replace '\b(unsigned|signed)\s+', ' ').Trim()
        if ($t -eq 'void' -or $t -eq '...') { continue }
        if ($t -match '[\[\]()="\'']') { return -1 }
        if (-not $ItemRx.IsMatch($t)) { return -1 }
    }
    return $items.Count
}

if (-not (Test-Path -LiteralPath $Rtl)) {
    Write-Host ("FAIL: RTL 目录不在: " + $Rtl)
    exit 1
}
$decls = @{}
$defs  = @{}
$files = @(Get-ChildItem -LiteralPath $Rtl -Recurse -File -Include *.h, *.c)
foreach ($f in $files) {
    $isHeader = ($f.Extension -eq '.h')
    $lines = [IO.File]::ReadAllLines($f.FullName)
    $i = 0
    while ($i -lt $lines.Count) {
        $ln = $lines[$i]
        if ($ln -eq '' -or $ln -match '^[ \t#/]') { $i += 1; continue }
        $m = $HeadRx.Match($ln)
        if (-not $m.Success) { $i += 1; continue }
        $buf = Strip-Cmt $ln
        $k = $i
        while ((Get-Balance $buf) -gt 0 -and ($k + 1 -lt $lines.Count)) {
            $k += 1
            $buf += ' ' + (Strip-Cmt $lines[$k]).Trim()
        }
        $open  = $buf.IndexOf('(')
        $close = $buf.LastIndexOf(')')
        if ($open -lt 0 -or $close -le $open) { $i += 1; continue }
        $params = $buf.Substring($open + 1, $close - $open - 1)
        $tail   = $buf.Substring($close + 1).Trim()
        $kind = ''
        if ($isHeader) {
            if ($tail.EndsWith(';')) { $kind = 'decl' }
        } else {
            $nxt = ''
            if ($k + 1 -lt $lines.Count) { $nxt = $lines[$k + 1].Trim() }
            if ($tail.StartsWith('{') -or $nxt.StartsWith('{')) { $kind = 'def' }
        }
        if ($kind -ne '') {
            $ar = Get-Arity (Strip-Cmt $params)
            if ($ar -ge 0) {
                $name = $m.Groups[1].Value
                $bag = $defs
                if ($kind -eq 'decl') { $bag = $decls }
                if (-not $bag.ContainsKey($name)) { $bag[$name] = @() }
                $rel = $f.FullName.Substring($Root.Length + 1)
                $bag[$name] += , @($ar, $rel, ($i + 1))
            }
        }
        $i = $k + 1
    }
}

$viol = @()
$compared = 0
foreach ($name in $defs.Keys) {
    if (-not $decls.ContainsKey($name)) { continue }
    $compared += 1
    $da = @($decls[$name] | ForEach-Object { $_[0] } | Sort-Object -Unique)
    $fa = @($defs[$name]  | ForEach-Object { $_[0] } | Sort-Object -Unique)
    if (($da -join ',') -ne ($fa -join ',')) {
        $d0 = $decls[$name][0]
        $f0 = $defs[$name][0]
        $viol += ("{0}: 头 {1} 参 ({2}:{3}) 对不上 体 {4} 参 ({5}:{6})" -f `
            $name, ($da -join '/'), $d0[1], $d0[2], ($fa -join '/'), $f0[1], $f0[2])
    }
}
if ($files.Count -lt 100) { $viol += ("RTL 文件数={0} (<100) ⇒ 路径或通配被改坏" -f $files.Count) }
if ($compared -lt 900)   { $viol += ("两头都有的签名对数={0} (<900) ⇒ 哨兵自废" -f $compared) }

# ---------- R3 / R4 / R5 (账 #278 §B72): 表 ↔ RTL ↔ 类型 oracle ↔ 产物 ----------
$tblPath = Join-Path $Root 'src\backend\cgen_util_ctrl.cpp'
$typPath = Join-Path $Root 'src\backend\cgen_util_type.cpp'
if (-not (Test-Path -LiteralPath $tblPath)) {
    $viol += ('R3 发码表不在: ' + $tblPath)
} else {
    $tblTxt = [IO.File]::ReadAllText($tblPath)
    $tbl = @([regex]::Matches($tblTxt, 'controlExit\("(vb6_[A-Za-z0-9_]+)",\s*(\d+),\s*outArgc\)'))
    if ($tbl.Count -ne 23) { $viol += ('TABLE-ROWS: controlExit 答了 ' + $tbl.Count + ' 行 (恰好 23)') }
    $typArity = @{}
    if (Test-Path -LiteralPath $typPath) {
        $typTxt = [IO.File]::ReadAllText($typPath)
        foreach ($tm in [regex]::Matches($typTxt, '\{"(vb6_[A-Za-z0-9_]+)",\s*\{([^}]*)\}\}')) {
            $body = $tm.Groups[2].Value.Trim()
            $n = 1
            if ($body -eq '') { $n = 0 }
            foreach ($ch in $body.ToCharArray()) { if ($ch -eq ',') { $n += 1 } }
            $typArity[$tm.Groups[1].Value] = $n
        }
    }
    $tblNames = @{}
    foreach ($e in $tbl) {
        $nm = $e.Groups[1].Value
        $ar = [int]$e.Groups[2].Value
        $tblNames[$nm] = $ar
        if ($decls.ContainsKey($nm)) {
            $da = @($decls[$nm] | ForEach-Object { $_[0] } | Sort-Object -Unique)
            if (($da -join ',') -ne ('' + $ar)) {
                $viol += ('TABLE-VS-RTL: ' + $nm + ' 表里说递 ' + $ar + ' 枚, RTL 头里是 ' + ($da -join '/') + ' 参')
            }
        } elseif ($defs.ContainsKey($nm)) {
            $viol += ('TABLE-VS-RTL: ' + $nm + ' 只在 RTL 有定义没有声明 ⇒ 没人能替发码担保个数')
        } else {
            $viol += ('TABLE-VS-RTL: ' + $nm + ' 在 RTL 里查不到原型')
        }
        if ($typArity.ContainsKey($nm) -and $typArity[$nm] -ne $ar) {
            $viol += ('TABLE-VS-TYPEORACLE: ' + $nm + ' 表 ' + $ar + ' 参, 类型行 ' + $typArity[$nm] + ' 项')
        }
    }
    # R5: 名字只许住在表里(加上按名字对账的那张类型表)
    $spell = @(Get-ChildItem -Path (Join-Path $Root 'src\backend') -Recurse -Include *.cpp, *.inc, *.hpp |
                Where-Object { $_.FullName -notmatch 'cgen_util_ctrl\.cpp$' -and
                               $_.FullName -notmatch 'cgen_util_type\.cpp$' } |
                Select-String -CaseSensitive -Pattern '"(vb6_ClearList|vb6_SetControlFocus|vb6_Slider_ClearSel|vb6_ControlTextHeight|vb6_ControlTextWidth|vb6_ScaleUnitX|vb6_ScaleUnitY|vb6_ImageList_ClearImages|vb6_StatusBar_ClearPanels|vb6_CdShowOpen|vb6_CdShowSave|vb6_CdShowColor|vb6_CdShowFont|vb6_CdShowPrinter|vb6_CdShowAbout|vb6_Ws_Close|vb6_Ws_Listen|vb6_Ws_Connect|vb6_Ws_Accept|vb6_Ws_Bind|vb6_Ws_SendData|vb6_Ws_GetData|vb6_Ws_PeekData)[("]' |
                Where-Object { $_.Line -notmatch '^\s*(//|\*)' })
    foreach ($s in $spell) {
        $viol += ('NAME-COPIED: ' + $s.Path.Substring($Root.Length + 1) + ':' + $s.LineNumber)
    }
    # R4: 产物那一头 —— 实际递了几枚, 只吃 --emit-c, 不起 cl
    $c3 = Join-Path $Root '.build\C3.exe'
    $fx = Join-Path $Root 'tests\ctrlzero\ZeroForm.frm'
    if (-not (Test-Path -LiteralPath $c3)) {
        $viol += ('EMIT: 没有编译器 ' + $c3 + ' (这一面读的是发码)')
    } elseif (-not (Test-Path -LiteralPath $fx)) {
        $viol += 'EMIT: 夹具 tests/ctrlzero/ZeroForm.frm 不在'
    } else {
        $lat = [System.Text.Encoding]::GetEncoding(28591)
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ('ctrlarity_' + [guid]::NewGuid().ToString('N') + '.c')
        $tmpE = [IO.Path]::ChangeExtension($tmp, '.err')
        $pr = Start-Process -FilePath $c3 -ArgumentList @('"' + $fx + '"', '--emit-c') -WorkingDirectory $Root -RedirectStandardOutput $tmp -RedirectStandardError $tmpE -NoNewWindow -Wait -PassThru
        $emit = ''
        if (Test-Path -LiteralPath $tmp) { $emit = $lat.GetString([IO.File]::ReadAllBytes($tmp)) }
        Remove-Item $tmp, $tmpE -Force -ErrorAction SilentlyContinue
        if ($pr.ExitCode -ne 0) { $viol += ('EMIT: 夹具退出码 ' + $pr.ExitCode) }
        $exercised = 0
        foreach ($nm in $tblNames.Keys) {
            $ix = $emit.IndexOf($nm + '(')
            if ($ix -lt 0) { continue }
            $d = 0
            $nArgs = 0
            $seen = $false
            for ($j = $ix + $nm.Length + 1; $j -lt $emit.Length; $j++) {
                $ch = $emit[$j]
                if ($ch -eq '(') { $d += 1 }
                elseif ($ch -eq ')') { if ($d -eq 0) { break }; $d -= 1 }
                elseif ($ch -eq ',' -and $d -eq 0) { $nArgs += 1; $seen = $true }
                elseif ($ch -ne ' ' -and $ch -ne "`t" -and $ch -ne "`r" -and $ch -ne "`n") { $seen = $true }
            }
            $got = 0
            if ($seen) { $got = $nArgs + 1 }
            if ($got -ne $tblNames[$nm]) {
                $viol += ('EMIT-VS-TABLE: ' + $nm + ' 产物里递了 ' + $got + ' 枚, 表里说 ' + $tblNames[$nm] + ' 枚')
            }
            $exercised += 1
        }
        if ($exercised -ne 22) {
            $viol += ('EMIT-CENSUS: 夹具只跑出 ' + $exercised + ' 枚出口 (恰好 22) ⇒ 夹具或发码被改坏')
        }
        Write-Host ('table_rows=' + $tbl.Count + ' exercised=' + $exercised + ' type_rows=' + $typArity.Count)
    }
}

Write-Host ("rtl_files={0} decl_names={1} def_names={2} compared={3}" -f `
    $files.Count, $decls.Count, $defs.Count, $compared)
if ($viol.Count -gt 0) {
    Write-Host ("FAIL: 头与体的签名不齐 {0} 处" -f $viol.Count)
    foreach ($v in $viol) { Write-Host ("  " + $v) }
    exit 1
}
Write-Host ("PASS: {0} 对签名逐条对上" -f $compared)
exit 0
