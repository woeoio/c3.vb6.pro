# 规则（改坏了会红）：
#   R1  出站回写只有一位主人：`void vb6_ByRefScalarWriteBack(` 定义 1 处（vb6rtl_com.c）、
#       声明 1 处（vb6rtl_variant.h）。
#   （负控里 R1_def 那一份会**连带** R4 变红：定义被改名后 R4 的函数体切片取不到 —— 那不是假红。）
#   R2  服务端只在收尾之前调一次，且顺序钉死：那一句的行号必须小于
#       `vb6_VariantSlotRelease(&((vb6_VarSlot*)coercedArgs)` 那一句。
#   R3  入槽那一臂必须按 pointee 的类型搬值：`vb6_VariantBorrowIntoSlot` 函数体内
#       `case VT_I4/VT_I2/VT_R8/VT_R4/VT_BOOL/VT_UI1:` 恰 6 条 —— 少一条就是某一档又去读指针了。
#   R4  出站回写认同一组六档：`vb6_ByRefScalarWriteBack` 函数体内同样恰 6 条。
#   R5  夹具三件都在：类（`Bump(ByRef v As Long)` 里 `v = v + 1`）、CoClass 块、vbp 的 Class= 行。
#   R6  判据已进门禁：`b231_byref_scalar` 恰 2 格（x64 + x86），探针 `BYREFARG=` / `BYVALARG=` 各 1 处 dispatch。
#
# 用法:  powershell -File scripts\check_byref_scalar_writeback.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
# C3_CHECK_ROOT 只为**负控**存在: 把整棵树拷到别处、改一处、再指回来跑, 真树一个字节不动
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$comC   = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_com.c'
$varH   = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_variant.h'
$srvC   = Join-Path $root 'src\rtl\core\vb6comserver\vb6comserver_obj.c'
$fxCls  = Join-Path $root 'tests\b231_byrefparam\ByRefCls.cls'
$fxBas  = Join-Path $root 'tests\b231_byrefparam\DllBlocks.bas'
$fxVbp  = Join-Path $root 'tests\b231_byrefparam\B231k.vbp'
$runner = Join-Path $root 'tests\run_tests.ps1'
$probe  = Join-Path $root 'tests\tools\disp_probe.c'

foreach ($p in @($comC, $varH, $srvC, $fxCls, $fxBas, $fxVbp, $runner, $probe)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}

function Count-InFile($path, $pattern) {
    $t = [System.IO.File]::ReadAllText($path)
    return ([regex]::Matches($t, $pattern)).Count
}
function Count-InFunc($path, $fnMarker, $pattern) {
    $t = [System.IO.File]::ReadAllText($path)
    $i = $t.IndexOf($fnMarker)
    if ($i -lt 0) { return -1 }
    $blk = $t.Substring($i)
    $e = $blk.IndexOf('' + [char]10 + '}')
    if ($e -lt 0) { return -2 }
    $blk = $blk.Substring(0, $e)
    return ([regex]::Matches($blk, $pattern)).Count
}
function Line-Of($path, $pattern) {
    $lines = [System.IO.File]::ReadAllLines($path)
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match $pattern) { return $i + 1 }
    }
    return -1
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

Expect 'R1_def_sites'  (Count-InFile $comC 'void vb6_ByRefScalarWriteBack\(') 1 '回写主人只许一处定义'
Expect 'R1_decl_sites' (Count-InFile $varH 'void vb6_ByRefScalarWriteBack\(') 1 '声明一处'
Expect 'R2_call_sites' (Count-InFile $srvC 'vb6_ByRefScalarWriteBack\(') 1 '服务端只调一次'
$lnCall = Line-Of $srvC 'vb6_ByRefScalarWriteBack\('
$lnRel  = Line-Of $srvC 'vb6_VariantSlotRelease\(&\(\(vb6_VarSlot\*\)coercedArgs\)'
if (($lnCall -gt 0) -and ($lnRel -gt 0) -and ($lnCall -lt $lnRel)) { $orderOk = 1 } else { $orderOk = 0 }
Expect 'R2_writeback_before_release' $orderOk 1 ("回写须在收尾之前 (call=" + $lnCall + " release=" + $lnRel + ")")
Expect 'R3_inbound_switch'  (Count-InFunc $comC 'void vb6_VariantBorrowIntoSlot(' 'case VT_(I4|I2|R8|R4|BOOL|UI1):') 6 '入槽那一臂六档各一条'
Expect 'R4_outbound_switch' (Count-InFunc $comC 'void vb6_ByRefScalarWriteBack(' 'case VT_(I4|I2|R8|R4|BOOL|UI1):') 6 '出站回写认同一组六档'
Expect 'R5_fixture_class'   (Count-InFile $fxCls 'Public Function Bump\(ByRef v As Long\) As Long') 1 '夹具那一枚方法'
Expect 'R5_fixture_coclass' (Count-InFile $fxBas 'Implementation\("ByRefCls"\)') 1 'CoClass 块'
Expect 'R5_fixture_vbp'     (Count-InFile $fxVbp 'Class=ByRefCls; ByRefCls.cls') 1 'vbp 的 Class= 行'
Expect 'R6_runner_cases'    (Count-InFile $runner '"b231_byref_scalar') 2 '判据两台'
Expect 'R6_probe_byref'     (Count-InFile $probe 'strncmp\(argv\[ai\], "BYREFARG=", 9\)') 1 '探针 BYREF 臂'
Expect 'R6_probe_byval'     (Count-InFile $probe 'strncmp\(argv\[ai\], "BYVALARG=", 9\)') 1 '探针按值对照臂'

if ($fails.Count -gt 0) {
    Write-Host ('BYREF_SCALAR_WRITEBACK RED (' + $fails.Count + ')') -ForegroundColor Red
    exit 1
}
Write-Host 'BYREF_SCALAR_WRITEBACK GREEN'
exit 0
