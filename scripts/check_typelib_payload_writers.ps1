# 类型库符号 payload 写点哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道 (2026-11 共享化重构的守门人):
#   类型库符号 (coclass / 接口成员 / 枚举 / 模块函数 / GlobalNs) 以前每个模块各造一份,
#   539 模块 × 3042 符号 ≈ 12 GB。重构后全仓只有**一份**共享基座 (Driver::typelibTable_),
#   各模块 SymbolTable 以非拥有只读回退方式共享它。
#
#   这份共享之所以安全, 全靠一个前提: **payload 字段是只读的**。审计确认它们的写点
#   全仓只有注入块一处; 分析与发码阶段只读。将来任何人只要在别处写一笔 (例如 cgen 里
#   顺手改一下 comMethods, 或跨模块阶段把某个共享符号的 comIidStr 覆盖掉), 共享表立刻
#   变成"跨模块串值"——A 模块的写会让 B 模块读到错的东西, 而且没有崩溃、没有诊断,
#   只有发码悄悄变味。这种 bug 极难查, 所以把这个前提钉成断言。
#
# 规则:
#   S0  注入块 (Driver::buildSharedTypelibSymbols) 必须存在且能被大括号配对定位
#   S1  payload 字段名**程序化**取自注入块 (不手抄): 所有 `<symVar>->FIELD =` /
#       `->FIELD[...] =` / `->FIELD.push_back(...)` 的 FIELD 名; 注入块里调用了
#       insertComMethod ⇒ 补上 comMethods (它由该助手写入)。
#   S2  核心 payload (com*/eventNames/memberNames/comMethods) 的写点全仓只许落在
#       "注入权威区" = insertComMethod 函数体 + buildSharedTypelibSymbols 函数体。
#   S3  const 族 (hasConstValue / constIntValue / constType) 是**例外**: 工程自有常量
#       在声明期合法写它们。因此只允许声明期文件 (semantics/* 与 builtin/*.inc)、
#       driver_crossmod.cpp 的 extSym 拷贝、以及注入权威区; 其余任何位置写一律红。
#   S4  isBuiltin 不纳入本哨兵: 它是声明期标志, 且 ClassMemberParamsMemo 等局部结构
#       也有同名字段 (纯文本静态扫描会误报)。
#   S5  (2026-11 载荷共享) 写门只认 Symbol::W() —— 字段写形态可以是
#       `sym->FIELD =` 或 `sym->W().FIELD =`; W() 是 COW 门, 判据口径不变。
#   S6  (2026-11 载荷共享) 读门 Symbol::P() 不许出现在写点左侧。共享之后
#       `sym->P().memberNames = …` 会改到别的模块共享的那一份, 必须走 W()。
#
# 用法:  powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check_typelib_payload_writers.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$injRel = "src\driver\driver_semantics.cpp"
$injPath = Join-Path $root $injRel
if (-not (Test-Path -LiteralPath $injPath)) {
    Write-Host ("FAIL S0 missing " + $injRel) -ForegroundColor Red
    exit 1
}
$src = [System.IO.File]::ReadAllText($injPath)

# ---- 定位函数体 (大括号配对; 跳过字符串与 // 注释) ----
function Find-FuncBody([string]$text, [string]$pattern) {
    $m = [regex]::Match($text, $pattern)
    if (-not $m.Success) { return $null }
    $i = $text.IndexOf('{', $m.Index + $m.Length - 1)
    if ($i -lt 0) { return $null }
    $depth = 0
    $inStr = $false
    $j = $i
    while ($j -lt $text.Length) {
        $c = $text[$j]
        if ($inStr) {
            if ($c -eq '\') { $j += 2; continue }
            if ($c -eq '"') { $inStr = $false }
            $j++
            continue
        }
        if ($c -eq '"') { $inStr = $true; $j++; continue }
        if ($c -eq '/' -and ($j + 1) -lt $text.Length -and $text[$j + 1] -eq '/') {
            while ($j -lt $text.Length -and $text[$j] -ne "`n") { $j++ }
            continue
        }
        if ($c -eq '{') { $depth++ }
        elseif ($c -eq '}') {
            $depth--
            if ($depth -eq 0) {
                return [pscustomobject]@{ Start = $i; End = $j; Text = $text.Substring($i, $j - $i + 1) }
            }
        }
        $j++
    }
    return $null
}

$build = Find-FuncBody $src 'void\s+Driver::buildSharedTypelibSymbols\s*\(\s*\)'
$helper = Find-FuncBody $src 'void\s+insertComMethod\s*\([^)]*\)'
if (-not $build) {
    Write-Host "FAIL S0 cannot locate Driver::buildSharedTypelibSymbols body" -ForegroundColor Red
    exit 1
}
if (-not $helper) {
    Write-Host "FAIL S0 cannot locate insertComMethod body" -ForegroundColor Red
    exit 1
}

function Get-LineNo([string]$text, [int]$idx) {
    return (($text.Substring(0, $idx) -split "`n").Count)
}
$allowed = @(
    @{ File = $injRel.ToLower(); From = (Get-LineNo $src $helper.Start); To = (Get-LineNo $src $helper.End) },
    @{ File = $injRel.ToLower(); From = (Get-LineNo $src $build.Start); To = (Get-LineNo $src $build.End) }
)

# ---- S1: 从注入块程序化提取 payload 字段名 ----
# 只认**类型库符号对象本身**的字段: 先找出注入块里 `auto <v> = std::make_unique<Symbol>(...)`
# 的局部名 (sym/enumSym/enumTypeSym ...), 再取 `<v>->FIELD` 的写形态。
# 这样 sig.realName / pi.type 这些 ComMethodSig/ParameterInfo 的局部字段不会被误算进来。
$body = $build.Text
$symVars = @([regex]::Matches($body, 'auto\s+(\w+)\s*=\s*std::make_unique<\s*Symbol\s*>') |
             ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
if ($symVars.Count -eq 0) {
    Write-Host "FAIL S1 no 'auto X = std::make_unique<Symbol>' found in the injection block" -ForegroundColor Red
    exit 1
}
$names = @()
$varAlt = (($symVars | ForEach-Object { [regex]::Escape($_) }) -join '|')
foreach ($mm in [regex]::Matches($body, '\b(' + $varAlt + ')\s*->\s*(?:W\(\)\s*\.\s*)?([A-Za-z_]\w*)\s*(?:\[[^\]]*\]\s*)?(?:=(?!=)|\+=|-=|\|=|&=)')) {
    $names += $mm.Groups[2].Value
}
foreach ($mm in [regex]::Matches($body, '\b(' + $varAlt + ')\s*->\s*(?:W\(\)\s*\.\s*)?([A-Za-z_]\w*)\s*\.\s*(?:push_back|emplace_back|insert|emplace|push)\s*\(')) {
    $names += $mm.Groups[2].Value
}
if ($body -match 'insertComMethod\s*\(') { $names += 'comMethods' }
$names = @($names | Sort-Object -Unique)

$constFamily = @('hasConstValue', 'constIntValue', 'constType', 'constStringValue', 'constFloatValue', 'constBoolValue')
$core = @($names | Where-Object { $constFamily -notcontains $_ -and $_ -ne 'isBuiltin' })
$constSeen = @($names | Where-Object { $constFamily -contains $_ })

if ($core.Count -lt 15) {
    $bad += ("S1 extracted only " + $core.Count + " core payload fields from the injection block (want >= 15); names=" + ($names -join ','))
}

# ---- 扫描全仓写点 ----
$srcDir = Join-Path $root "src"
$files = Get-ChildItem -LiteralPath $srcDir -Recurse -File |
    Where-Object { $_.Extension -in ".cpp", ".hpp", ".inc", ".c", ".h" }

# const 族允许的声明期文件 (只按**文件名**判, 并显式列出目录约束)
$constAllowedFiles = @(
    "src\semantics\semantic_analyzer_register.cpp",
    "src\semantics\semantic_analyzer_decl_type.cpp",
    "src\semantics\builtin\builtin_consts.inc",
    "src\semantics\builtin\builtin_funcs.inc",
    "src\driver\driver_crossmod.cpp",
    # Fix 047 跨模块枚举成员预登记 (fresh msym/alias, 不是共享层符号)
    "src\driver\driver_semantics.cpp"
)

function New-WriteRegex([string[]]$fields) {
    $alt = (($fields | ForEach-Object { [regex]::Escape($_) }) -join '|')
    # 2026-11 载荷共享重构后, payload 的写形态多了一层权威入口: `sym->W().FIELD`
    # (Symbol::W() = COW 写门)。这里把 W() 记为**可选前缀**, 判据本身不放宽 ——
    # 仍然是"哪个字段被写、写在哪个文件"的同一套口径。
    $pat = "(?:->|\.)\s*(?:W\(\)\s*\.\s*)?($alt)\s*(?:\[[^\]]*\]\s*)?(?:=(?!=)|\+=|-=|\|=|&=)" +
           "|(?:->|\.)\s*(?:W\(\)\s*\.\s*)?($alt)\s*\.\s*(?:push_back|emplace_back|insert|emplace|push)\s*\("
    return [regex]::new($pat, [System.Text.RegularExpressions.RegexOptions]::Compiled)
}
# S6: 读门 P() 绝不许出现在写点左侧 —— 写必须走 W() (否则共享载荷被原地改,
# 消费模块的眼里的类成员表会跟着源模块变; 无崩溃、无诊断, 只有发码悄悄变味)。
function New-ReadAsWriteRegex([string[]]$fields) {
    if ($fields.Count -eq 0) { return $null }
    $alt = (($fields | ForEach-Object { [regex]::Escape($_) }) -join '|')
    $pat = "(?:->|\.)\s*P\(\)\s*\.\s*($alt)\s*(?:\[[^\]]*\]\s*)?(?:=(?!=)|\+=|-=|\|=|&=)" +
           "|(?:->|\.)\s*P\(\)\s*\.\s*($alt)\s*\.\s*(?:push_back|emplace_back|insert|emplace|push)\s*\("
    return [regex]::new($pat, [System.Text.RegularExpressions.RegexOptions]::Compiled)
}
$coreRe = $null
$constRe = $null
if ($core.Count -gt 0) { $coreRe = New-WriteRegex $core }
if ($constSeen.Count -gt 0) { $constRe = New-WriteRegex $constSeen }
$readAsWriteRe = New-ReadAsWriteRegex $core

$coreHits = @()
$constHits = @()
$readAsWriteHits = @()
foreach ($file in $files) {
    $rel = $file.FullName.Substring($root.Length + 1)
    $relLower = $rel.ToLower()
    $lines = [System.IO.File]::ReadAllText($file.FullName) -split "`r?`n"
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $raw = $lines[$i]
        $cut = $raw.IndexOf('//')
        $code = if ($cut -ge 0) { $raw.Substring(0, $cut) } else { $raw }
        if ($code.Trim().Length -eq 0) { continue }
        $ln = $i + 1
        $isAllowed = $false
        foreach ($a in $allowed) {
            if ($a.File -eq $relLower -and $ln -ge $a.From -and $ln -le $a.To) { $isAllowed = $true; break }
        }
        if ($null -ne $coreRe) {
            foreach ($mm in $coreRe.Matches($code)) {
                $fld = if ($mm.Groups[1].Success) { $mm.Groups[1].Value } else { $mm.Groups[2].Value }
                if (-not $isAllowed) { $coreHits += ("{0}:{1}  {2}" -f $rel, $ln, $fld) }
            }
        }
        if ($null -ne $constRe) {
            foreach ($mm in $constRe.Matches($code)) {
                $fld = if ($mm.Groups[1].Success) { $mm.Groups[1].Value } else { $mm.Groups[2].Value }
                if (-not $isAllowed -and ($constAllowedFiles -notcontains $rel)) {
                    $constHits += ("{0}:{1}  {2}" -f $rel, $ln, $fld)
                }
            }
        }
        if ($null -ne $readAsWriteRe) {
            foreach ($mm in $readAsWriteRe.Matches($code)) {
                $fld = if ($mm.Groups[1].Success) { $mm.Groups[1].Value } else { $mm.Groups[2].Value }
                $readAsWriteHits += ("{0}:{1}  {2}" -f $rel, $ln, $fld)
            }
        }
    }
}

# ---- 已知的"工程 Class 符号"payload 写点 (冻结名单) ----
# 为什么不是"全仓只许注入块一处": Symbol 是同一个结构体, 工程 Class 也用它那批
# com*/memberNames/eventNames 字段 —— 声明期 (semantic_analyzer)、继承合并
# (driver_classchain)、跨模块副本 (driver_crossmod 的 fresh extSym)、DLL 元数据
# 回填 (driver_codegen_dll_*) 本来就合法地写它们。它们写的都是 kind==Class 的
# **模块自有**符号, 不是共享层里的类型库符号。
# 所以哨兵的口径是:
#   ① 注入权威区以外的**新文件**写核心 payload → 红;
#   ② 下列文件的核心写点**条数**被冻结, 多一条/少一条 → 红 (新增必须人工确认
#      写的是工程 Class, 不是查表得到的共享类型库符号)。
# 注意 cgen_util_dllentry_collect.inc 那条写的是局部 info 结构, 不是 Symbol。
$knownCore = @{
    "src\driver\driver_classchain.cpp"                     = 1
    # 2026-11 载荷共享: 3 → 2。少的那一条是 `extSym->memberNames = srcSym->memberNames`
    # (11 张类成员表之一), 现在整份载荷由 `extSym->sharePayloadFrom(*srcSym)` 共享,
    # 不再逐字段深拷贝; eventNames / comClsidStr 仍是普通字段, 照旧显式拷。
    "src\driver\driver_crossmod.cpp"                       = 2
    "src\driver\detail\driver_codegen_dll_sync.inc"        = 7
    "src\driver\detail\driver_codegen_dll_typelib.inc"     = 9
    "src\semantics\semantic_analyzer.cpp"                  = 5
    "src\backend\detail\util\cgen_util_dllentry_collect.inc" = 1
}

$coreByFile = @{}
foreach ($h in $coreHits) {
    $f = ($h -split ':')[0]
    if (-not $coreByFile.ContainsKey($f)) { $coreByFile[$f] = @() }
    $coreByFile[$f] += $h
}
foreach ($f in $coreByFile.Keys) {
    $n = @($coreByFile[$f]).Count
    if (-not $knownCore.ContainsKey($f)) {
        $bad += ("S2 NEW payload writer file (outside injection authority): " + $f + " x" + $n)
        $bad += $coreByFile[$f]
    } elseif ($knownCore[$f] -ne $n) {
        $bad += ("S2 payload write count drift in " + $f + ": expected " + $knownCore[$f] + ", got " + $n)
        $bad += $coreByFile[$f]
    }
}
foreach ($f in $knownCore.Keys) {
    if (-not $coreByFile.ContainsKey($f)) {
        $bad += ("S2 frozen payload writer disappeared from " + $f + ": expected " + $knownCore[$f] + ", got 0 (re-pin after review)")
    }
}

if ($constHits.Count -gt 0) {
    $bad += ("S3 const-family write sites outside declaration-time files = " + $constHits.Count)
    $bad += $constHits
}

if ($readAsWriteHits.Count -gt 0) {
    $bad += ("S6 payload written through the read door P() = " + $readAsWriteHits.Count + " (写必须走 W())")
    $bad += $readAsWriteHits
}

if ($bad.Count -eq 0) {
    $frozen = 0
    foreach ($v in $knownCore.Values) { $frozen += $v }
    Write-Host ("PASS typelib payload writers: 核心字段 " + $core.Count + " 个 / const 族 " + $constSeen.Count +
                " 个; 注入权威区 = " + $injRel + "; 冻结的工程 Class 写点 = " + $frozen + " 处/" + $knownCore.Count + " 文件") -ForegroundColor Green
    Write-Host ("      core=" + ($core -join ','))
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
