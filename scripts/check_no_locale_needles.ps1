# check_no_locale_needles.ps1 - 登记针里不许出现「跟着区域设置走」的格式化字面量
#
# 为什么单开这一道: 这一类**不会**让本机红 —— 它是**跨机器**才显形的假红，而且显形方式
#   最难查: 门在 CI 那台 en-US 上红、在本机 zh-CN 上绿，读数还「看着像对的」。
#   本仓已经为这件事红过两次（账 #433 第一跑、账 #436 第一跑，都是**我自己写的夹具**）:
#     FormatDateTime(d, 0) = namedFormat 0 是 **vbShortDate = 跟着区域设置走**的一档
#       本机 zh-CN 读 "2026/1/5"、CI en-US 读 "1/5/2026" ⇒ 钉字面量必假红
#     CStr(True) 的字面文本各语言不同 ⇒ Boolean 也不能直接 CStr 钉，要比成数字 1/0
#   所以这不是「某个人写错一次」，是**一整类**会反复咬人的坑，值得钉成门。
#
# 判定不看注释、**看做法**：一根针里出现了区域形状的字面量（日期斜杠/连字符、时间冒号、
#   AM-PM 记号），它**必须**由下面两种**区域无关**的做法之一产生，否则报红:
#     ① 显式格式串 `Format$(v, "yyyy/mm/dd")` / `Format(v, "hh:mm:ss")` —— 自己写死要哪个形状，
#        不问系统（记号大小写也跟着串里写的那一种，§B201）。
#     ② 数值取值器拼装 `Year(v) & "/" & Month(v)` —— 全是数字，与区域无关。
#   反过来 `FormatDateTime(v, 0)` / `CStr(<Date>)` / `Now` 造出来的东西**不算**（那正是会跟着
#   区域设置走的一族），两种做法都找不到就报红。
#
# 配对方式：**按登记语句解析**，不按行距猜。
#   `Add-BasTest "n" "$Tests\f.bas" $nNeedles` / `Test-Vbp 'n' "$Tests\f.bas" $nExpected -Arch ...`
#   ⇒ (用例名, 夹具路径, 针变量名)，再取那个变量的字面量表。
# ⚠ 行距猜测会把针算到**上一个**登记块头上（第一版就这么把本仓自己的 b424/b425/b426 针
#   全算到 b184_fieldw 名下，26 条假红）—— 登记块有两种写法，路径有的在针下面、有的在上面。
#
# 用法:  powershell -File scripts\check_no_locale_needles.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$harness = Join-Path $root 'tests\run_tests.ps1'
if (-not (Test-Path -LiteralPath $harness)) { Write-Host ('FAIL missing ' + $harness) -ForegroundColor Red; exit 1 }
$raw = [IO.File]::ReadAllLines($harness)

# ---- 1) 建 (fixture, needleVar) 表：解析每一条登记语句 ----
$regRe = [regex]'(?i)(?:Add-(?:Bas|Vbp)Test|Test-Vbp)\s+[''"]([^''"]+)[''"]\s+[''"]\$Tests\\([^''"]+)[''"]\s+\$(\w+)'
$regs = @()
foreach ($l in $raw) {
    $t = ($l -replace '#.*$', '')
    $m = $regRe.Match($t)
    if ($m.Success) { $regs += [pscustomobject]@{ Name = $m.Groups[1].Value; Fixture = $m.Groups[2].Value; Var = $m.Groups[3].Value } }
}
if ($regs.Count -lt 20) {
    Write-Host ('FAIL the registration scan only found ' + $regs.Count + ' registrations (floor 20) -> the parse broke; ' +
                'a gate that silently checks almost nothing is worse than no gate') -ForegroundColor Red
    exit 1
}

# ---- 2) 取每个针变量的字面量表（整段扫描，遇到下一个 `= @(` 就换人）----
$varLits = @{}
$cur = ''
foreach ($l in $raw) {
    $t = ($l -replace '#.*$', '')
    $a = [regex]::Match($t, '^\s*\$(\w+)\s*=\s*@\(')
    if ($a.Success) {
        $cur = $a.Groups[1].Value
        if (-not $varLits.ContainsKey($cur)) { $varLits[$cur] = @() }
        # ⚠ **单行写法** `$zzNeedles = @("V=2026/1/5")`：字面量就在 `= @(` 那一行上。
        #   只 `continue` 而不收 ⇒ 这一行的针**全部丢失** ⇒ 门对单行声明的登记完全失明
        #   （负控就是这么溜过去的：注入的那根针压根没被扫到）。
        $after = $t.Substring($t.IndexOf('@(') + 2)
        foreach ($m in [regex]::Matches($after, '"([^"]*)"')) { $varLits[$cur] += $m.Groups[1].Value }
        if ($t -match '\)\s*$') { $cur = '' }      # 单行写完就收口
        continue
    }
    if ($t -match '^\s*\)\s*$' -or $t -match '^\s*@\(') { if ($cur -ne '') { $cur = '' }; continue }
    if ($cur -ne '') {
        foreach ($m in [regex]::Matches($t, '"([^"]*)"')) { $varLits[$cur] += $m.Groups[1].Value }
    }
}

# 区域形状
$shapeRes = @('\d{4}/\d{1,2}/\d{1,2}', '\d{4}-\d{1,2}-\d{1,2}', '\d{1,2}:\d{2}:\d{2}',
              '\d{1,2}:\d{2}\s*[APap]\b', '\b\d{1,2}:\d{2}\s*(?:AM|PM)\b')
# 区域无关的两种做法（先剥注释再找：注释里写 Format$ 不算做法）
$safeBuilder = '(?i)\bFormat\$?\s*\(|\bFormat(?:Number|Currency|Percent)\s*\(|\b(?:Year|Month|Day|Hour|Minute|Second)\s*\('

$bad = @(); $checked = 0; $srcCache = @{}
foreach ($r in $regs) {
    if (-not $varLits.ContainsKey($r.Var)) { continue }
    $hits = @()
    foreach ($lit in $varLits[$r.Var]) {
        foreach ($p in $shapeRes) { if ($lit -match $p) { $hits += $lit; break } }
    }
    if ($hits.Count -eq 0) { continue }
    if (-not $srcCache.ContainsKey($r.Fixture)) {
        $p2 = Join-Path $root (Join-Path 'tests' $r.Fixture)
        $b = ''
        if (Test-Path -LiteralPath $p2) {
            if ([IO.Path]::GetExtension($p2) -ieq '.vbp') {
                # ⚠ .vbp 是**工程文件**，VB 源码在它列出来的 Form/Module/Class 里。
                #   拿 .vbp 的文本当源码找做法 ⇒ 永远找不到 ⇒ 本仓每一个 vbp 用例都假红
                #   （第一版就这么把 b417 / b419 / b426 三家全判红）。要跟着工程去读成员文件。
                $projDir = Split-Path -Parent $p2
                $vbpTxt = [IO.File]::ReadAllText($p2)
                foreach ($mm in [regex]::Matches($vbpTxt, '(?im)^\s*(?:Form|Module|Class|Base)\s*=\s*"?([^"\r\n]+?)"?\s*$')) {
                    # ⚠ 工程文件里一行是「**VB 名; 文件名**」两个字段: `Module=B426Main; B426Main.bas`
                    #   / `Form=frmMain; B426Main.frm`。整段当路径 ⇒ 文件不存在 ⇒ 又回到假红。
                    $parts = $mm.Groups[1].Value -split ';'
                    $fileName = $parts[$parts.Count - 1].Trim()
                    if ($fileName -eq '') { continue }
                    $member = Join-Path $projDir $fileName
                    if (Test-Path -LiteralPath $member) { $b += [IO.File]::ReadAllText($member) + "`n" }
                }
                if ($b -eq '') { $b = [IO.File]::ReadAllText($p2) }
            } else {
                $b = [IO.File]::ReadAllText($p2)
            }
        }
        $srcCache[$r.Fixture] = [regex]::Replace($b, '(?m)''.*$', ' ')
    }
    $code = $srcCache[$r.Fixture]
    $checked++
    if ($code -notmatch $safeBuilder) {
        $bad += ('fixture ' + $r.Fixture + ' (case ' + $r.Name + ') pins locale-shaped literal(s): ' +
                 (($hits | Select-Object -Unique) -join ', ') +
                 ' -- but its source builds them with NEITHER explicit format string NOR numeric date getters. ' +
                 'FormatDateTime(v,0) / CStr(<Date>) / Now all follow the system locale -> false-pink elsewhere.')
    }
}

if ($bad.Count -eq 0) {
    Write-Host ('PASS no_locale_needles: registrations ' + $regs.Count + ' | fixtures with locale-shaped literals ' +
                $checked + ' | all locale-proof | violations 0') -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1