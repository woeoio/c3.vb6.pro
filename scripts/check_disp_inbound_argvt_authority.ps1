# check_disp_inbound_argvt_authority.ps1 - 入站 IDispatch「VARIANT → 形参声明档」那一问只许有一位主人
#   (账 #305 §B219，第二百二十刀)
#
# 改前这一问的答案住在**两处半**：服务端适配器里那半张手抄名单（只折 VT_I2/I1/UI1/UI2/BOOL/
# EMPTY→VT_I4 与 VT_R4→VT_R8，名单之外一律不管），和发码那 16 处读臂（按声明档裸取联合体成员，
# 从不看标签）。两半各答一半的后果是实测出来的（两台相同，见 ai/022 第二百二十刀的账）：
# 文本递进数值档 ⇒ 交回那枚 BSTR 指针的低 32 位；数递进文本档 ⇒ 把标量当 BSTR 解 ⇒ 原生 AV，
# **客户端进程直接死掉**，一条错误都没有。
#
# 现在标签那一问只住 vb6_ComArgNormalize（换算产物落服务端自己的 scratch，收尾那一臂回收），
# 档位那一问只住 cgen_base_type.cpp 的 cDispatchArgVt，两者之间靠方法表里新增的 paramVts 传话。
# 本门钉的是「不许再长出第二份」，行为面由 tests\b220_args 那两格真跑钉（两台各一份针面）。
#
# 规则（改坏了会红）：
#   R1  服务端那一趟不许自己调 VariantChangeType —— 换算只许走 vb6_ComArgNormalize。
#       命中数必须是 0；>0 说明那半张手抄名单回来了（名单之外的档就又开始裸取了）。
#   R2  唯一出口的形状：定义 1 枚(vb6rtl_com.c) + 声明 1 枚(vb6rtl_variant.h) + 调用 1 处
#       (vb6comserver_obj.c)。任何一处变成 2 枚，就是把「换算」又抄了一遍（§B142 双胞胎同形）。
#   R3  发码读臂一律不看标签：collect.inc 里 "->vt ==" 命中数必须是 0（改前 Boolean 那一臂有 1 处，
#       与服务端那半张表同问 —— 两半各答一半就是这么来的）。
#   R4  档位答案只住 cDispatchArgVt：定义 1 处 + 站点 1 处(tables.inc)。站点从 1 变多 ⇒ 有人在
#       别处现拼档位名单。
#   R5  VB6_ARGVT_SKIP 这个「本刀不问」的口径必须有主人（define 1 处）；表里字段桥那两条交
#       NULL, 0 = 不问，命中数 = 2（变成 0 说明有人给字段桥也套上换算，那一腿交的是整槽 vb6_VARIANT）。
#   R6  方法描述表那两格新字段（paramVts / paramVtCount）各 1 枚；少一枚就是发码与 RTL 不同步。
#
# 用法:  powershell -File scripts\check_disp_inbound_argvt_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
# C3_CHECK_ROOT 只为**负控**存在: 把整棵树拷到别处、改一处、再指回来跑, 真树一个字节不动
# (这条线过去做负控是就地改+md5 还原; 这里选了更稳的那一头)。
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$srvDir   = Join-Path $root 'src\rtl\core\vb6comserver'
$srvObj   = Join-Path $root 'src\rtl\core\vb6comserver\vb6comserver_obj.c'
$srvHdr   = Join-Path $root 'src\rtl\core\vb6comserver\vb6comserver.h'
$comC     = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_com.c'
$varHdr   = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_variant.h'
$collect  = Join-Path $root 'src\backend\detail\util\cgen_util_dllentry_collect.inc'
$tables   = Join-Path $root 'src\backend\detail\util\cgen_util_dllentry_tables.inc'
$baseType = Join-Path $root 'src\backend\cgen_base_type.cpp'

foreach ($p in @($srvObj, $srvHdr, $comC, $varHdr, $collect, $tables, $baseType)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}

function Count-Pattern($path, $pattern, $recursive) {
    if ($recursive) {
        $files = Get-ChildItem -Path $path -Recurse -File -Include *.c, *.h
    } else {
        $files = @(Get-Item -LiteralPath $path)
    }
    $n = 0
    foreach ($f in $files) {
        $t = [System.IO.File]::ReadAllText($f.FullName)
        $n += ([regex]::Matches($t, $pattern)).Count
    }
    return $n
}

$fails = @()
function Expect($tag, $got, $want, $note) {
    if ($got -ne $want) {
        $script:fails += ('{0} = {1} (want {2}) -- {3}' -f $tag, $got, $want, $note)
        Write-Host ('FAIL {0} = {1}, want {2}  ({3})' -f $tag, $got, $want, $note) -ForegroundColor Red
    } else {
        Write-Host ('ok   {0} = {1}' -f $tag, $got) -ForegroundColor Green
    }
}

# R1 —— 服务端那一趟一次也不许自己换算
Expect 'R1 server-side VariantChangeType' (Count-Pattern $srvDir 'VariantChangeType\s*\(' $true) 0 `
    '换算只许走 vb6_ComArgNormalize; 这里>0 = 那半张手抄名单回来了'

# R2 —— 唯一出口的定义/声明/调用各一枚
Expect 'R2 vb6_ComArgNormalize defined'   (Count-Pattern $comC   '(?m)^HRESULT vb6_ComArgNormalize' $false) 1 `
    '定义只许一枚, 在 vb6rtl_com.c'
Expect 'R2 vb6_ComArgNormalize declared'  (Count-Pattern $varHdr 'HRESULT vb6_ComArgNormalize' $false) 1 `
    '声明只许一枚, 在 vb6rtl_variant.h'
Expect 'R2 vb6_ComArgNormalize called'    (Count-Pattern $srvObj 'vb6_ComArgNormalize\(src' $false) 1 `
    '服务端调用点只许一处(进门那一趟)'

# R3 —— 发码读臂不看标签
Expect 'R3 codegen arg reader tests vt'   (Count-Pattern $collect '->vt\s*==' $false) 0 `
    '读臂一律裸取联合体成员; 标签那一问住服务端那一趟'

# R4 —— 档位答案只住一位主人
Expect 'R4 cDispatchArgVt defined'        (Count-Pattern $baseType '(?m)^std::string CCodeGen::cDispatchArgVt' $false) 1 `
    '声明档的答案只住 cgen_base_type.cpp 这一处'
Expect 'R4 cDispatchArgVt site'           (Count-Pattern $tables   'cDispatchArgVt\(' $false) 1 `
    '表那一处是唯一消费者; 变多 = 别处又开始现拼档位'

# R5 —— 「本刀不问」的口径与字段桥那两条
Expect 'R5 VB6_ARGVT_SKIP defined'        (Count-Pattern $varHdr '#define VB6_ARGVT_SKIP' $false) 1 `
    'SKIP 这一档必须只有一个定义处'
Expect 'R5 field-bridge rows hands-off'   (Count-Pattern $tables 'NULL, 0 \},' $false) 2 `
    'Public 字段的 Get/Let 两行交整槽, 不许套换算'

# R6 —— 表的新字段
Expect 'R6 desc has paramVts'             (Count-Pattern $srvHdr 'paramVts;' $false) 1 '方法描述表的声明档数组字段'
Expect 'R6 desc has paramVtCount'         (Count-Pattern $srvHdr 'paramVtCount;' $false) 1 '它的长度字段(越界不许读)'

if ($fails.Count -gt 0) {
    Write-Host ('RED: ' + $fails.Count + ' guard(s) broken') -ForegroundColor Red
    exit 1
}
Write-Host 'PASS: 入站「VARIANT → 声明档」只有一位主人 (R1-R6)' -ForegroundColor Green
exit 0
