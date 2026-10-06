# check_ctrl_array_members.ps1 - 账 #189 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: VB6 里 `arr.Count / arr.LBound / arr.UBound` 问的是**数组本身**, 不是某一枚
# 控件的属性。发码侧以前只有 LBound/UBound 在成员读取那一路各写了一条 if, Count 漏了 ⇒ 掉进
# COM 兜底, 发成 vb6_ComGetIntProp(vb6_hwnd_<数组名>, L"Count") —— 而数组控件压根没有
# vb6_hwnd_<数组名> 这个变量 (实测 Charts 2020/ucChartBar 的 Form2: error C2065)。
# 与账 #157 同族: 句柄类表达式必须走 vb6_arr_*, 不许凭空拼 vb6_hwnd_。
#
# 口径现在收在一处: CCodeGen::ctrlArrayMetaMemberExpr (src/backend/cgen_util_ctrl.cpp)。
#
# 规则 (改坏了会红, 不是装饰):
#   A1  src/backend 与 src/semantics 里, 三个 RTL 出口 (vb6_CtrlArr_GetCount/LBound/UBound) 只许
#       出现在权威那一个文件里 —— 别处再拼一份就是又开了一条平行体系
#   A2  权威函数存在、只定义一次, 三条分支 (count/lbound/ubound) 各一条, 且"不是这三条"要返空串
#   A3  手写分支 (memLower == "lbound" 那一形) 归零; 权威至少有 1 个调用点 (只留定义不算)
#
# 用法:  pwsh -File scripts\check_ctrl_array_members.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$authRel = "src\backend\cgen_util_ctrl.cpp"
$auth = Join-Path $root $authRel
if (-not (Test-Path -LiteralPath $auth)) {
    Write-Host ("FAIL A0 authority file missing: " + $authRel) -ForegroundColor Red
    exit 1
}

$dirs = @("src\backend", "src\semantics") | ForEach-Object { Join-Path $root $_ }
$files = @()
foreach ($d in $dirs) {
    if (Test-Path -LiteralPath $d) {
        $files += Get-ChildItem -LiteralPath $d -Recurse -File |
                  Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" }
    } else { $bad += ("A0 missing dir: " + $d) }
}

# A1: 三个 RTL 出口只许在权威文件里出现
$patRtl = 'vb6_CtrlArr_(GetCount|LBound|UBound)\('
$leak = @()
foreach ($f in $files) {
    if ($f.FullName -eq $auth) { continue }
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($t, $patRtl)) { $leak += ($f.Name + ":" + $m.Value) }
}
if ($leak.Count -ne 0) {
    $bad += ("A1 RTL accessor hand-concatenated outside the authority = " + $leak.Count +
             " -> " + (($leak | Select-Object -First 6) -join " | "))
}

# A2: 权威的形状
$h = [System.IO.File]::ReadAllText($auth)
$defs = @([regex]::Matches($h, 'std::string\s+CCodeGen::ctrlArrayMetaMemberExpr\s*\('))
if ($defs.Count -ne 1) {
    $bad += ("A2 ctrlArrayMetaMemberExpr defined " + $defs.Count + " times in the authority (want 1)")
}
foreach ($br in @("count", "lbound", "ubound")) {
    $n = @([regex]::Matches($h, 'memberLower\s*==\s*"' + $br + '"')).Count
    if ($n -ne 1) { $bad += ("A2 branch for '" + $br + "' present " + $n + " times (want 1)") }
}
$blk = [regex]::Match($h, 'ctrlArrayMetaMemberExpr[\s\S]{0,900}?\r?\n\}')
if (-not $blk.Success -or $blk.Value -notmatch 'return\s+""\s*;') {
    $bad += 'A2 the authority lost its empty-string exit ("not one of these three" must return "")'
}

# A3: 手写分支归零 + 调用点 >= 1
# 注意只扫 lbound/ubound: `memLower == "count"` 在 Collections 那两路 (vb6_Forms_Count /
# COM 集合 Count) 是**正当**的, 一并扫会变成冤案红 —— 本条钉的是数组那三条兄弟分支别再散着写。
$patHand = 'memLower\s*==\s*"(lbound|ubound)"'
$hand = 0
$handWhere = @()
$calls = 0
foreach ($f in $files) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    if ($f.FullName -ne $auth) {
        foreach ($m in [regex]::Matches($t, $patHand)) {
            $hand++
            $handWhere += ($f.Name + ":" + $m.Value)
        }
    }
    $calls += @([regex]::Matches($t, 'ctrlArrayMetaMemberExpr\s*\(')).Count
}
if ($hand -ne 0) {
    $bad += ("A3 hand-written member branches = " + $hand + " (must be 0) -> " +
             (($handWhere | Select-Object -First 6) -join " | "))
}
# 调用点计数里要扣掉权威自己的定义与声明 (各 1): 只剩定义 = 没人接, 也算红
if ($calls -lt 3) {
    $bad += ("A3 only " + $calls + " mentions of the authority (definition + declaration + >=1 call site wanted)")
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS ctrl-array member authority: RTL exits confined to " + $authRel +
                ", " + $calls + " mentions, 0 hand-written branches") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
