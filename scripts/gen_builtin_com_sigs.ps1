# gen_builtin_com_sigs.ps1 - 生成 src\com\builtin_scripting_sigs.inc (C30-11 第十二刀)
#
# 这张表是**生成物**，用途见 TypeLibParser::registerBuiltinScriptingTypelib()：
# 没装 scrrun 类型库的机器上，`Dim x As FileSystemObject` 也要拿到与有库时完全相同的编译期类型。
#
# 两步:
#   1) 用编译器把本机真库吐成 TSV（只读诊断，不改任何代码）:
#        .build\C3.exe <任一.bas> --dump-com-sigs Scripting > %TEMP%\c3_scripting_sigs.tsv
#   2) 跑本脚本:  pwsh -File scripts\gen_builtin_com_sigs.ps1 [-Tsv <路径>]
#
# 验收口径（不要跳过）: pwsh -File scripts\check_builtin_com_sigs.ps1
#   它会把同一份探针在「默认」与「C3_TEST_NO_SCRRUN_TLB=1」两种编译下各发一遍，
#   要求两侧 C **逐字节相同**。表改错一档（Boolean 写成 Long）只有这条捉得住。
#
# 注意（踩过的坑）: 成员表若用 OrderedDictionary 存，PowerShell 的 `.Count`/`.Keys`
# 会被**同名成员**遮住 —— Scripting.Dictionary 自己有 Count/Keys 成员，于是
# `$dict.Count` 拿到的是那个成员而不是元素个数。所以这里另外留一份有序 List。
param(
    [string]$Tsv = (Join-Path $env:TEMP 'c3_scripting_sigs.tsv')
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not (Test-Path -LiteralPath $Tsv)) {
    Write-Host ("FAIL 找不到 TSV: " + $Tsv)
    Write-Host "     先跑: .build\C3.exe <任一.bas> --dump-com-sigs Scripting > " + $Tsv
    exit 1
}

$lines = Get-Content -LiteralPath $Tsv -Encoding UTF8
$keep = @('FileSystemObject','Dictionary','Drive','Drives','File','Files','Folder','Folders','TextStream')
$inherited = @('QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke')

$classes = [ordered]@{}
$cur = $null
$curMember = $null

foreach ($ln in $lines) {
    $f = $ln -split "`t"
    if ($f[0] -eq 'CC') {
        $cls = $f[1]
        $cur = $null; $curMember = $null
        if ($keep -notcontains $cls) { continue }
        if ($classes.Contains($cls)) { continue }     # 重复的 coclass（被 import 带进来的）取第一份
        $cur = [ordered]@{
            cls = $cls; progId = $f[2]; iface = $f[4]; iid = ''
            dual = $false; dispatch = $false; dfltName = ''; dfltReal = ''
            dict = @{}; list = New-Object System.Collections.ArrayList
        }
        $classes[$cls] = $cur
        continue
    }
    if (-not $cur) { continue }
    if ($f[0] -eq 'IF' -and $f[1] -eq $cur.cls) {
        $cur.iface = $f[2]; $cur.iid = $f[3]
        foreach ($part in $f[4..($f.Count - 1)]) {
            if ($part -like 'dual=*') { $cur.dual = ($part -eq 'dual=1') }
            elseif ($part -like 'dispatch=*') { $cur.dispatch = ($part -eq 'dispatch=1') }
            elseif ($part -like 'default=*') { $cur.dfltName = $part.Substring(8) }
        }
        continue
    }
    if ($f[0] -eq 'MB' -and $f[1] -eq $cur.cls) {
        $curMember = $null
        if ($inherited -contains $f[2]) { continue }
        if ($cur.dict.ContainsKey($f[2])) { $curMember = $cur.dict[$f[2]]; continue }
        $curMember = [ordered]@{ real = $f[2]; kind = $f[3]; ret = $f[4]; memid = [int]$f[5]; vtbl = [int]$f[6]; params = New-Object System.Collections.ArrayList }
        $cur.dict[$f[2]] = $curMember
        [void]$cur.list.Add($curMember)
        continue
    }
    if ($f[0] -eq 'PA' -and $curMember -and $f[1] -eq $cur.cls) {
        [void]$curMember.params.Add(@($f[3], $f[4], $f[5], ($f[6] -eq 'opt=1')))
        continue
    }
}

$sb = [System.Text.StringBuilder]::new()
function W($s) { [void]$sb.AppendLine($s) }
$tlbPath = (($lines | Where-Object { $_ -like 'TL*' } | Select-Object -First 1) -split "`t")[3]

W '// ============================================================'
W '// 本文件是**生成物** —— 不要手改。'
W '//'
W '// 用途: TypeLibParser::registerBuiltinScriptingTypelib() 的内建签名表。'
W '//       内建兜底原本只登类名（默认接口是空的），于是没有 scrrun 类型库的机器上成员'
W '//       只能运行期 by-name；这张表把本机真库（scrrun.dll）的成员签名原样搬过来。'
W '//'
W '// 来源: `C3.exe <任一.bas> --dump-com-sigs Scripting` 的输出，'
W '//       再由 scripts\gen_builtin_com_sigs.ps1 机械转换（重新生成的办法写在那个脚本头部）。'
W '//       生成时间: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
W '//       源库: ' + $tlbPath
W '//'
W '// 已剔除: IUnknown/IDispatch 那 7 个继承成员（按名字调用不到）、以及本运行期不实现的'
W '//         Encoder coclass。clsidStr 也**故意**留空 —— 内建表不该进 ComLib=/Object='
W '//         那几条免注册表路（见函数内注释）。'
W '//'
W '// 验收: scripts\check_builtin_com_sigs.ps1 —— 同一份探针在「默认」与'
W '//       C3_TEST_NO_SCRRUN_TLB=1 两种编译下，发出的 C 必须逐字节相同。'
W '// ============================================================'
W ''
W 'struct BuiltinParam { const char* name; Vb6Type type; ComParamDir dir; bool opt; };'
W 'struct BuiltinMember { const char* real; ComMemberKind kind; Vb6Type ret; int32_t memid; int32_t vtbl; int nparams; const BuiltinParam* params; };'
W 'struct BuiltinClass { const char* cls; const char* progId; const char* iface; const char* iid; bool dual; bool dispatch; const char* dfltName; const char* dfltReal; int nmembers; const BuiltinMember* members; };'
W ''

foreach ($c in $classes.Values) {
    $idx = 0
    $memberRows = New-Object System.Collections.ArrayList
    foreach ($m in $c.list) {
        $pname = 'nullptr'; $pcount = 0
        if ($m.params.Count -gt 0) {
            $pname = "kBP_$($c.cls)_$idx"
            W "static const BuiltinParam $pname[] = {"
            foreach ($p in $m.params) {
                $pn = ($p[0] -replace '\\', '\\').Replace('"', '\"')
                $b = if ($p[3]) { 'true' } else { 'false' }
                W "    { `"$pn`", Vb6Type::$($p[1]), ComParamDir::$($p[2]), $b },"
            }
            W '};'
            $pcount = $m.params.Count
        }
        $real = $m.real.Replace('"', '\"')
        [void]$memberRows.Add("    { `"$real`", ComMemberKind::$($m.kind), Vb6Type::$($m.ret), $($m.memid), $($m.vtbl), $pcount, $pname },")
        $idx++
    }
    W "static const BuiltinMember kBM_$($c.cls)[] = {"
    foreach ($r in $memberRows) { W $r }
    W '};'
    W ''
}

W 'static const BuiltinClass kBuiltinClasses[] = {'
foreach ($c in $classes.Values) {
    $dr = ''
    foreach ($m in $c.list) { if ($m.real -eq $c.dfltName) { $dr = $m.real } }
    if (-not $dr) { foreach ($m in $c.list) { if ($m.real -ieq $c.dfltName) { $dr = $m.real } } }
    $c.dfltReal = $dr
    $dual = if ($c.dual) { 'true' } else { 'false' }
    $disp = if ($c.dispatch) { 'true' } else { 'false' }
    W ("    { `"$($c.cls)`", `"$($c.progId)`", `"$($c.iface)`", `"$($c.iid)`", $dual, $disp, `"$($c.dfltName)`", `"$dr`", $($c.list.Count), kBM_$($c.cls) },")
}
W '};'

New-Item -ItemType Directory -Force -Path 'src\com' | Out-Null
Set-Content -Path 'src\com\builtin_scripting_sigs.inc' -Value $sb.ToString() -Encoding UTF8
$total = 0
foreach ($c in $classes.Values) { $total += $c.list.Count }
Write-Output ('生成完成: src\com\builtin_scripting_sigs.inc, 9 类 / ' + $total + ' 成员, 源库=' + $tlbPath)
