# 规则（改坏了会红）：
#   R1  出站回写只有一个主人：`void vb6_ArraySlotWriteBack(` 定义 1 处（vb6rtl_com.c）、
#       声明 1 处（vb6rtl_variant.h）。0 = 被撤了；2 = 又长出一份。
#   R2  服务端只在**收尾之前**调它一次，且顺序钉死：那一句的行号必须小于
#       `vb6_VariantSlotRelease(&((vb6_VarSlot*)coercedArgs)` 那一句 —— 反了就变成
#       「先回收描述符再回写」（读到已释放的载体）。
#   R3  入槽那一臂必须认 VT_BYREF：`(src->vt & VT_BYREF) ? (src->pparray ? *src->pparray : NULL) : src->parray`
#       恰 1 处。0 = BYREF 形又被当成 SAFEARRAY* 用（槽留 Empty、每发漏一个载体）。
#   R4  回写只认 BYREF|ARRAY 那一形：`if ((client->vt & VT_ARRAY) == 0 || (client->vt & VT_BYREF) == 0) return;` 恰 1 处。
#   R5  夹具三件都在：类（两个形参、ReDim 后填 (i+1)*7）、CoClass 块、vbp 的 Class= 行。
#   R6  判据已进门禁：`b230_arrout_writeback` 恰 2 格（x64 + x86），且探针两条臂都在
#       （`ARRARG=` 与 `ARRARGV=` 各 1 处 dispatch）。
#
# 用法:  powershell -File scripts\check_array_slot_writeback.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
# C3_CHECK_ROOT 只为**负控**存在: 把整棵树拷到别处、改一处、再指回来跑, 真树一个字节不动
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$comC   = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_com.c'
$varH   = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_variant.h'
$srvC   = Join-Path $root 'src\rtl\core\vb6comserver\vb6comserver_obj.c'
$fxCls  = Join-Path $root 'tests\b230_arrout\ArrOutCls.cls'
$fxBas  = Join-Path $root 'tests\b230_arrout\DllBlocks.bas'
$fxVbp  = Join-Path $root 'tests\b230_arrout\B230k.vbp'
$runner = Join-Path $root 'tests\run_tests.ps1'
$probe  = Join-Path $root 'tests\tools\disp_probe.c'

foreach ($p in @($comC, $varH, $srvC, $fxCls, $fxBas, $fxVbp, $runner, $probe)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}

function Count-InFile($path, $pattern) {
    $t = [System.IO.File]::ReadAllText($path)
    return ([regex]::Matches($t, $pattern)).Count
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

Expect 'R1_def_sites'      (Count-InFile $comC 'void vb6_ArraySlotWriteBack\(') 1 '回写主人只许一处定义'
Expect 'R1_decl_sites'     (Count-InFile $varH 'void vb6_ArraySlotWriteBack\(') 1 '声明一处'
Expect 'R2_call_sites'     (Count-InFile $srvC 'vb6_ArraySlotWriteBack\(') 1 '服务端只调一次（不含注释以外的第二处）'
$lnCall = Line-Of $srvC 'vb6_ArraySlotWriteBack\('
$lnRel  = Line-Of $srvC 'vb6_VariantSlotRelease\(&\(\(vb6_VarSlot\*\)coercedArgs\)'
if (($lnCall -gt 0) -and ($lnRel -gt 0) -and ($lnCall -lt $lnRel)) { $orderOk = 1 } else { $orderOk = 0 }
Expect 'R2_writeback_before_release' $orderOk 1 ("回写须在收尾之前 (call=" + $lnCall + " release=" + $lnRel + ")")
Expect 'R3_byref_deref'    (Count-InFile $comC '\(src->vt & VT_BYREF\) \? \(src->pparray \? \*src->pparray : NULL\) : src->parray') 1 '入槽那一臂必须认 VT_BYREF'
Expect 'R4_byref_only'     (Count-InFile $comC 'if \(\(client->vt & VT_ARRAY\) == 0 \|\| \(client->vt & VT_BYREF\) == 0\) return;') 1 '回写只认 BYREF|ARRAY'
Expect 'R5_fixture_class'  (Count-InFile $fxCls 'Public Function Fill\(ByVal n As Long, a\(\) As Long\) As Long') 1 '夹具那一枚方法'
Expect 'R5_fixture_coclass'(Count-InFile $fxBas 'Implementation\("ArrOutCls"\)') 1 'CoClass 块'
Expect 'R5_fixture_vbp'    (Count-InFile $fxVbp 'Class=ArrOutCls; ArrOutCls.cls') 1 'vbp 的 Class= 行'
Expect 'R6_runner_cases'   (Count-InFile $runner '"b230_arrout_writeback') 2 '判据两台'
Expect 'R6_probe_byref'    (Count-InFile $probe 'strncmp\(argv\[ai\], "ARRARG=", 7\)') 1 '探针 BYREF 臂'
Expect 'R6_probe_plain'    (Count-InFile $probe 'strncmp\(argv\[ai\], "ARRARGV=", 8\)') 1 '探针 plain 对照臂'

if ($fails.Count -gt 0) {
    Write-Host ('ARRAY_SLOT_WRITEBACK RED (' + $fails.Count + ')') -ForegroundColor Red
    exit 1
}
Write-Host 'ARRAY_SLOT_WRITEBACK GREEN'
exit 0
