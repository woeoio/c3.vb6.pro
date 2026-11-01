# 规则（改坏了会红）：
#   R1  CLSID 的补齐必须**逐表逐符号**扫（两趟：先收集非空、再填空值）。两趟各命中 1 处。
#       0 = 有人把这条腿撤了（View / PopupMenu 那类又会被 dll_entry 重新 mint 一枚）；
#       >1 = 有人在别处又实现了一遍按名补齐。
#   R2  这条腿**不许**住进「默认接口 IID 还空着」那道闸里：闸内那条 comClsidStr 拷贝必须已经删掉
#       （命中数 0）。住回去 = 已经有默认成员的那几枚整块被跳过（正是本刀修的病灶）。
#   R3  只许填空值：补齐语句旁边必须带 `comClsidStr.empty()` 守卫（命中 ≥1），
#       没有它就是「覆盖别人已经写好的身份」—— 那会把已有身份搬走。
#   R4  值只有一个来源：`generateUuid` 兜底在 driver 侧仍只有 1 处（类型库那一趟），
#       服务器表侧不许新长出一枚自己的 mint。
#
# 用法:  powershell -File scripts\check_clsid_sync_all_copies.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
# C3_CHECK_ROOT 只为**负控**存在: 把整棵树拷到别处、改一处、再指回来跑, 真树一个字节不动
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$sync = Join-Path $root 'src\driver\detail\driver_codegen_dll_sync.inc'
$tlb  = Join-Path $root 'src\driver\detail\driver_codegen_dll_typelib.inc'
foreach ($p in @($sync, $tlb)) {
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

Expect 'R1a_collect_pass'   (Count-InFile $sync 'knownClsidByName\[s2->name\] = s2->comClsidStr;') 1 '第一趟：收集非空身份'
Expect 'R1b_fill_pass'      (Count-InFile $sync 's2->comClsidStr = it2->second;') 1 '第二趟：按名填空值'
Expect 'R2_old_guarded_copy' (Count-InFile $sync 'sym->comClsidStr = srcSym->comClsidStr;') 0 '旧的「IID 闸内」那条拷贝必须已删'
Expect 'R3a_fill_op_only_empty' (Count-InFile $sync 'if \(!s2->comClsidStr\.empty\(\)\) continue;') 1 '补齐那一轮只接空值'
Expect 'R3b_collect_only_filled' (Count-InFile $sync 'if \(!s2->comClsidStr\.empty\(\)\) knownClsidByName') 1 '收集那一轮只收非空值'
Expect 'R4_tlb_fallback'    (Count-InFile $tlb  'TypeLibBuilder::generateUuid\(clsSym\.name\)') 1 '身份兜底仍只有类型库那一处'

if ($fails.Count -gt 0) {
    Write-Host ('CLSID_SYNC_ALL_COPIES RED (' + $fails.Count + ')') -ForegroundColor Red
    exit 1
}
Write-Host 'CLSID_SYNC_ALL_COPIES GREEN'
exit 0
