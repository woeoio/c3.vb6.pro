# check_builtin_com_sigs.ps1 - C30-11 第十二刀: 内建 Scripting 签名表
#
# 背景: TypeLibParser::registerBuiltinScriptingTypelib() 原本只登类名（默认接口是空的），
# 于是「这台机器没有 scrrun 类型库」时 `Dim x As FileSystemObject` 的成员只能运行期 by-name：
# 行为对，但编译期不查成员存在性、不查参数个数。第十二刀把成员签名也登进去，
# 表是生成物 src\com\builtin_scripting_sigs.inc（从本机 scrrun.dll 用 --dump-com-sigs 吐出）。
#
# 判据三条:
#   R1 生成物文件在、带生成物标记、9 个 coclass 一个不少；
#   R2 每个类的成员数（类表那一列）与它 kBM_ 数组的实际行数一致（防止写表时错位）；
#   R3 **动态**: 同一份探针在「默认」与「C3_TEST_NO_SCRRUN_TLB=1」两种编译下发出的 C
#      必须逐字节相同。这一条才是这张表的意义所在 —— 有库/无库编译结果必须一模一样。
#      （没有 .build\C3.exe 时 R3 记 SKIP，不算失败: 静态哨兵仍然生效。）
#
# 为什么要 R3 而不是只数成员: 签名表差一档（比如 Boolean 写成 Long）不会让任何静态检查红，
# 只会让「没库的机器」发出另一发调用（ComCallBool ⇄ ComCallInt），运行期表现为
# VarType/CStr 与有库时不同 —— 只有逐字节对 emit 才捉得住。
#
# Usage: pwsh -File scripts\check_builtin_com_sigs.ps1   (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = New-Object System.Collections.ArrayList

function Fail([string]$msg) { [void]$bad.Add($msg); Write-Host ("  x " + $msg) }

# ---------------------------------------------------------------- R1
$incRel = 'src\com\builtin_scripting_sigs.inc'
$inc = Join-Path $root $incRel
if (-not (Test-Path -LiteralPath $inc)) {
    Write-Host ("FAIL missing " + $incRel)
    exit 1
}
$txt = [System.IO.File]::ReadAllText($inc)
if ($txt -notmatch '本文件是\*\*生成物\*\*') { Fail "R1 表头缺『生成物』标记（会被人当手写文件改）" }
if ($txt -notmatch 'dump-com-sigs') { Fail "R1 表头没写来源命令（--dump-com-sigs），无从复现" }
if ($txt -notmatch 'C3_TEST_NO_SCRRUN_TLB') { Fail "R1 表头没写验收口径（两腿 emit 逐字节相同）" }

$classPat = '(?m)^\s*\{\s*"(\w+)",\s*"Scripting\.([\w.]+)",\s*"(\w+)",\s*"(\{[0-9A-Fa-f-]+\})",\s*(true|false),\s*(true|false),\s*"([^"]*)",\s*"([^"]*)",\s*(\d+),\s*kBM_(\w+)\s*\}'
$classRows = [regex]::Matches($txt, $classPat)
if ($classRows.Count -ne 9) {
    Fail ("R1 类表行数 = " + $classRows.Count + "，应为 9（FileSystemObject/Dictionary/Drive/Drives/File/Files/Folder/Folders/TextStream）")
}

# ---------------------------------------------------------------- R2
$total = 0
foreach ($m in $classRows) {
    $cls = $m.Groups[1].Value
    $arr = $m.Groups[10].Value
    $declared = [int]$m.Groups[9].Value
    if ($arr -ne $cls) { Fail ("R2 " + $cls + " 的数组名是 kBM_" + $arr + "，与类名不一致") }
    $blockPat = '(?s)static const BuiltinMember kBM_' + [regex]::Escape($arr) + '\[\]\s*=\s*\{(.*?)\n\};'
    $blk = [regex]::Match($txt, $blockPat)
    if (-not $blk.Success) { Fail ("R2 找不到 kBM_" + $arr + " 数组"); continue }
    $rows = [regex]::Matches($blk.Groups[1].Value, '(?m)^\s*\{ "')
    if ($rows.Count -ne $declared) {
        Fail ("R2 " + $cls + ": 类表写 " + $declared + " 个成员, 数组实际 " + $rows.Count + " 行")
    }
    if ($m.Groups[5].Value -ne 'true' -or $m.Groups[6].Value -ne 'true') {
        Fail ("R2 " + $cls + ": dual/dispatch 都不是 true —— scrrun 真库两样都是 true，照抄才对")
    }
    if ($m.Groups[4].Value -notmatch '^\{[0-9A-Fa-f]{8}-') { Fail ("R2 " + $cls + ": IID 不像 IID") }
    $total += $declared
}
if ($total -ne 109) {
    Fail ("R2 成员总数 = " + $total + "，基线 109（scrrun.dll 的 9 个类实测量）。变了要么真库换了版本，要么表被改坏 —— 都要人看一眼")
}

# ---------------------------------------------------------------- R3
$c3 = Join-Path $root '.build\C3.exe'
if (-not (Test-Path -LiteralPath $c3)) {
    Write-Host "  (R3 SKIP: 没有 .build\C3.exe)"
} else {
    $probe = Join-Path ([System.IO.Path]::GetTempPath()) ("c3_com_sigs_" + [Guid]::NewGuid().ToString('N') + ".bas")
    $src = @'
Option Explicit

Sub Main()
    Dim fso As FileSystemObject
    Dim d As Dictionary
    Dim f As File
    Dim ts As TextStream
    Dim b As Boolean
    Dim n As Long
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = CreateObject("Scripting.Dictionary")
    b = fso.FolderExists("C:\Windows")
    b = fso.FileExists("C:\Windows\notepad.exe")
    s = fso.GetFileName("C:\Windows\notepad.exe")
    n = fso.GetSpecialFolder(2).Size
    Set f = fso.GetFile("C:\Windows\notepad.exe")
    b = f.Attributes
    d.Add "a", 1
    n = d.Count
    b = d.Exists("a")
    Set ts = fso.OpenTextFile("C:\Windows\win.ini", 1)
    b = ts.AtEndOfStream
    s = ts.ReadLine
    Debug.Print b; s; n
End Sub
'@
    [System.IO.File]::WriteAllText($probe, $src, (New-Object System.Text.UTF8Encoding($false)))
    function Emit([bool]$noScrrun) {
        if ($noScrrun) { $env:C3_TEST_NO_SCRRUN_TLB = '1' } else { Remove-Item Env:C3_TEST_NO_SCRRUN_TLB -ErrorAction SilentlyContinue }
        $out = & $c3 $probe --emit-c 2>&1 | Out-String
        Remove-Item Env:C3_TEST_NO_SCRRUN_TLB -ErrorAction SilentlyContinue
        return $out
    }
    $a = Emit $false
    $b2 = Emit $true
    Remove-Item -LiteralPath $probe -ErrorAction SilentlyContinue
    if ($a.Length -eq 0 -or $b2.Length -eq 0) {
        Fail "R3 探针 emit 为空（编译失败？）"
    } elseif ($a -ceq $b2) {
        Write-Host ("  (R3 两腿 emit 逐字节相同, " + (($a -split "`n").Count) + " 行)")
    } else {
        $la = $a -split "`r?`n"; $lb = $b2 -split "`r?`n"
        $idx = 0
        for ($i = 0; $i -lt [Math]::Min($la.Count, $lb.Count); $i++) { if ($la[$i] -ne $lb[$i]) { $idx = $i; break } }
        Fail ("R3 两腿 emit 不同, 首差行 " + ($idx + 1) + ": 有库=[" + $la[$idx].Trim() + "] 无库=[" + $lb[$idx].Trim() + "]")
    }
}

if ($bad.Count -gt 0) {
    Write-Host ("FAIL " + $bad.Count + " 项")
    exit 1
}
Write-Host ("PASS 内建 Scripting 签名表: 9 类 / " + $total + " 成员, 两腿 emit 一致")
exit 0
