# check_event_handler_names.ps1 - 账 #190 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: VB6 的标识符大小写不敏感, C 敏感。事件臂里调用的那个 C 函数名以前是拿
# **控件的设计期拼写**现拼的 (`cProcName(ctrl.controlName + "_Click")`), 可过程定义发的是
# **Sub 自己的拼写** —— 只要有人改了控件名没改过程名 (VB6 完全合法, 两枚照样配一对), 产物就是
# "引用一个没人定义的函数": 链接期 LNK2019 (实测 Charts 2020/ucChartBar 的 Form1 两枚, 正好 2 个
# 无法解析的外部符号)。这种红**只在链接期现形**, 单看 --emit-c 的形状完全看不出坏了。
#
# 口径现在收在一处: CCodeGen::eventHandlerFn (src/backend/cgen_util_ctrl.cpp)
#   —— 存在性与名字都从同一个 lookup 出来, 命中的是**符号**, 发的就是符号自己的名字。
#
# 规则 (改坏了会红, 不是装饰):
#   E1  src/backend 与 src/semantics 里**现拼**处理器名 = 0 处 (非注释行出现 cProcName(<某> + "_...) 即红)
#   E2  权威存在、只定义一次, 且必须"从命中的符号取名" (cProcName(sym->name) 那一手) + 留着空串出口
#   E3  调用点 >= 40 处 (43 处是本轮普查到的实际数量; 只留定义、把规则写回本地也算红)
#
# 用法:  pwsh -File scripts\check_event_handler_names.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$authRel = "src\backend\cgen_util_ctrl.cpp"
$auth = Join-Path $root $authRel
if (-not (Test-Path -LiteralPath $auth)) {
    Write-Host ("FAIL E0 authority file missing: " + $authRel) -ForegroundColor Red
    exit 1
}

$dirs = @("src\backend", "src\semantics") | ForEach-Object { Join-Path $root $_ }
$files = @()
foreach ($d in $dirs) {
    if (Test-Path -LiteralPath $d) {
        $files += Get-ChildItem -LiteralPath $d -Recurse -File |
                  Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" }
    } else { $bad += ("E0 missing dir: " + $d) }
}

# E1: 现拼 (注释行除外 —— 注释里引用这个形状正是为了说明别再写它)
$patHand = 'cProcName\([A-Za-z_.]+ \+ "_'
$hand = 0
$handWhere = @()
$calls = 0
foreach ($f in $files) {
    $text = [System.IO.File]::ReadAllText($f.FullName)
    $lines = $text -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $t = $lines[$i].Trim()
        if ($t.StartsWith("//") -or $t.StartsWith("*")) { continue }
        if ([regex]::IsMatch($lines[$i], $patHand)) {
            $hand++
            $handWhere += ($f.Name + ":" + ($i + 1))
        }
    }
    if ($f.FullName -ne $auth) {
        $calls += @([regex]::Matches($text, 'eventHandlerFn\s*\(')).Count
    }
}
if ($hand -ne 0) {
    $bad += ("E1 hand-composed handler names = " + $hand + " (must be 0; use eventHandlerFn) -> " +
             (($handWhere | Select-Object -First 6) -join " | "))
}

# E2: 权威的形状
$h = [System.IO.File]::ReadAllText($auth)
$defs = @([regex]::Matches($h, 'std::string\s+CCodeGen::eventHandlerFn\s*\('))
if ($defs.Count -ne 1) {
    $bad += ("E2 eventHandlerFn defined " + $defs.Count + " times in the authority (want 1)")
}
if ($h -notmatch 'symTab_\s*\.\s*lookup\(ctrlName \+ suffix\)') {
    $bad += 'E2 the authority no longer asks the symbol table (存在性与名字必须出自同一次 lookup)'
}
if ($h -notmatch 'cProcName\(sym->name') {
    $bad += 'E2 the authority does not take the name FROM the symbol - that is exactly the bug this guards'
}
$blk = [regex]::Match($h, 'eventHandlerFn[\s\S]{0,700}?\r?\n\}')
if (-not $blk.Success -or $blk.Value -notmatch 'return\s+std::string\(\)\s*;') {
    $bad += 'E2 the authority lost its empty-string exit ("没有处理器" 必须交空串, 调用方据此不装臂)'
}

# E3: 调用点数量
if ($calls -lt 40) {
    $bad += ("E3 only " + $calls + " call sites into the authority (want >= 40) - some arm went back to hand-composing")
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS event-handler-name authority: 0 hand-composed sites, " + $calls +
                " call sites into CCodeGen::eventHandlerFn") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
