# check_rtl_proto_arity.ps1 —— RTL 的声明与定义必须同一张签名（参数个数）（账 #240）
#
# 症状不在本机：vb6_OleCon_Init 的体是 10 参、头里那份原型还停在 6 参，而 cgen 发的是 10 个实参。
# VS2019 的 cl 在 C 模式下压根不诊断「实参过多」（本机两架构 rc=0、出 exe），runner 上那台新 cl
# 报 4 行 error C2197（10-6=4 枚多余实参，一枚一行），门就红在 olecon / olecon_x86 两条上。
# 所以这一族不能靠"本机编一遍"来验，只能对着源码比：头的参数个数 == 体的参数个数。
#
# 判据面：
#   R1 两头都有的名字：声明侧的个数集合必须等于定义侧的个数集合，否则红并列出 文件:行 与两边的数
#   R2 census 地板：扫到的 .h/.c 文件数 >= 100，且「两头都有」的签名对数 >= 900
#      （实测 112 份文件 / 1108 对）—— 路径写错或正则被改坏时，哨兵不许变成"绿着的空转"
#
# 边界（宁可漏报不误报，故此处只比个数）：
#   * 只有**第 0 列开始**的行才算签名 —— 调用点都缩进在函数体里，这一条同时把它们排干净；
#   * 参数表里出现认不出的形态（数组 / 函数指针 / 默认值 / 串） ⇒ 跳过该条，不计入也不报红；
#   * 不比类型拼写：同一函数的头与体写成 const X* 与 X* 是合法的，硬比只会造噪声。

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Rtl  = Join-Path $Root 'src\rtl'

$HeadRx = [regex]'^[A-Za-z_][A-Za-z0-9_]*(?:[\s\*]+[A-Za-z_][A-Za-z0-9_]*)*?[\s\*]+(vb6_[A-Za-z0-9_]+)\s*\('
$ItemRx = [regex]'^([A-Za-z_][A-Za-z0-9_]*\s*\*{0,3}\s*)?[A-Za-z_][A-Za-z0-9_]*$'

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
        $buf = ($ln -replace '//.*$', '')
        $k = $i
        while ((Get-Balance $buf) -gt 0 -and ($k + 1 -lt $lines.Count)) {
            $k += 1
            $buf += ' ' + (($lines[$k] -replace '//.*$', '') -replace '/\*.*?\*/', '').Trim()
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
            $ar = Get-Arity (($params -replace '//.*$', '') -replace '/\*.*?\*/', '')
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

Write-Host ("rtl_files={0} decl_names={1} def_names={2} compared={3}" -f `
    $files.Count, $decls.Count, $defs.Count, $compared)
if ($viol.Count -gt 0) {
    Write-Host ("FAIL: 头与体的签名不齐 {0} 处" -f $viol.Count)
    foreach ($v in $viol) { Write-Host ("  " + $v) }
    exit 1
}
Write-Host ("PASS: {0} 对签名逐条对上" -f $compared)
exit 0
