# 规则（改坏了会红）：
#   R1  通用销毁里必须**按运行期秩**问一句 `vb6_SaIs1D(` —— 命中数必须是 1。
#       0 = 那一句被撤掉了（又回到「一律当 1D」）；>1 = 有人在别处现问一遍秩。
#   R2  两条腿各一枚，形状钉死：`vb6_SafeArrayDestroy1D((vb6_SafeArray1D*)arr)` 与
#       `vb6_SafeArrayDestroyND((vb6_SafeArrayND*)arr)` 各 1。少一条就是某一种秩没人销。
#   R3  发码侧只许有一处发 `vb6_SA_Destroy(me->`（cgen_com.cpp 的出厂圈）。
#       变成 0 说明出厂圈改发别的名字（那要看新名字有没有同样问秩）；变成 2 说明又长出一处。
#   R4  夹具三形必须在：多维 Long / 多维 String / 一维 String（最后一枚是"别把 1D 也改坏"的对照）。
#   R5  判据已进门禁：b223_nddestroy_ 前缀的用例恰 4 格（两台 × 两枚 ND 类）。
#
# 用法:  powershell -File scripts\check_sa_destroy_rank_aware.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
# C3_CHECK_ROOT 只为**负控**存在: 把整棵树拷到别处、改一处、再指回来跑, 真树一个字节不动
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$rtHdr   = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_runtime.h'
$comC    = Join-Path $root 'src\backend\module\cgen_com.cpp'
$fxLong  = Join-Path $root 'tests\b223_nddestroy\NdLongCls.cls'
$fxStr   = Join-Path $root 'tests\b223_nddestroy\NdStrCls.cls'
$fxOne   = Join-Path $root 'tests\b223_nddestroy\OneDCls.cls'
$runner  = Join-Path $root 'tests\run_tests.ps1'

foreach ($p in @($rtHdr, $comC, $fxLong, $fxStr, $fxOne, $runner)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}

function Count-InFile($path, $pattern) {
    $t = [System.IO.File]::ReadAllText($path)
    return ([regex]::Matches($t, $pattern)).Count
}

$fails = @()
function Expect($tag, $got, $want, $note) {
    if ($got -ne $want) {
        $script:fails += ('{0} = {1} (want {2}) -- {3}' -f $tag, $got, $want, $note)
        Write-Host ('FAIL {0} = {1}, want {2}  ({3})' -f $tag, $got, $want, $note) -ForegroundColor Red
    } else {
        Write-Host ('ok   {0} = {1}' -f $tag, $got)
    }
}

Expect 'R1_rank_question_sites' (Count-InFile $rtHdr 'vb6_SaIs1D\(') 1 '通用销毁必须按运行期秩问一句（多维的存储也可能是 1D*，秩只能问运行期）'
Expect 'R2_destroy1d_leg'      (Count-InFile $rtHdr 'vb6_SafeArrayDestroy1D\(\(vb6_SafeArray1D\*\)arr\)') 1 '1D 那一腿'
Expect 'R2_destroynd_leg'      (Count-InFile $rtHdr 'vb6_SafeArrayDestroyND\(\(vb6_SafeArrayND\*\)arr\)') 1 'ND 那一腿'
Expect 'R3_emit_sites'         (Count-InFile $comC  'vb6_SA_Destroy\(me->') 1 '出厂圈只许一处发这枚通用销毁'
Expect 'R4_fixture_nd_long'    (Count-InFile $fxLong 'nds\(1 To 2, 3\) As Long') 1 '多维 Long 那一形'
Expect 'R4_fixture_nd_string'  (Count-InFile $fxStr  'nds\(1 To 2, 3\) As String') 1 '多维 String 那一形（唯一会漏元素的那档）'
Expect 'R4_fixture_1d_control' (Count-InFile $fxOne  'Private f\(3\) As String') 1 '一维对照形'
Expect 'R5_runner_cases'       (Count-InFile $runner '"b223_nddestroy_') 4 '判据两台 × 两枚 ND 类'

if ($fails.Count -gt 0) {
    Write-Host ('SA_DESTROY_RANK_AWARE RED (' + $fails.Count + ')') -ForegroundColor Red
    exit 1
}
Write-Host 'SA_DESTROY_RANK_AWARE GREEN'
exit 0
