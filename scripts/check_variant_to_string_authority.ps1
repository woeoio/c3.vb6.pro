# check_variant_to_string_authority.ps1 - 「VARIANT -> 目标档」只有一位主人，布尔转字符串按 VB 语言口径
#   (账 #305 §B224，第二百三十八刀)
#
# 起因：同一个进程里两种答案。oleaut32 的 VariantChangeType(VT_BOOL -> VT_BSTR) 交 "-1"/"0"
# (.local/k3684_ole.exe 实测，len=2)，而 RTL 自己的 CStr/Format 交 "True"/"False"。用户在真 VB6 6.0 上
# 跑的服务端探针给了判决：EchoStr(ByVal s As String) 收 VT_BOOL(True) 答 "[True]"、收 False 答
# "[False]" (ai/probes/bool_string A1/A2) ⇒ 服务端入站那一格与 VB 语言同口径，不是 Automation 的口径。
#
# 规则（改坏了会红）：
#   R1  全仓不许再有以 VT_BSTR 为目的档的 VariantChangeType 直调 —— 换算只许问 vb6_ComChangeType。
#   R2  唯一出口：定义 1 枚(vb6rtl_com.c) + 声明 2 枚(vb6rtl_variant.h 给 RTL；vb6com_internal.h 给
#       vb6com —— 那一族不 include vb6rtl.h，与 vb6_VariantToDouble 同一惯例)。
#   R3  消费者五处：vb6_ComArgNormalize(服务端入站) / vb6_ComVariantIntoSlot 的两条 default 臂
#       (COM 元素与整枚 VARIANT 各一处) / vb6_DecimalToBSTR / vb6_ComUnpackBSTR。谁退回直调都会被 R1
#       抓住，这一条只数「消费者还在不在」。
#   R4  布尔那一支的形状：VT_TYPEMASK 判档 + L"True"/L"False" 两个常量，且只许在主人自己的文件里。
#   R5  判据面：run_tests.ps1 的 b220 两格各有 True 与 False 两针，且改前的 -1 拼法只在 Absent 里。
#
# 用法: powershell -File scripts\check_variant_to_string_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
# C3_CHECK_ROOT 只为负控存在（拷一棵树、改一处、指回来跑；真树零改动）。
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$srcDir  = Join-Path $root 'src'
$comC    = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_com.c'
$rtlC    = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl.c'
$packC   = Join-Path $root 'src\rtl\core\vb6com\vb6com_pack.c'
$varHdr  = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_variant.h'
$comHdr  = Join-Path $root 'src\rtl\core\vb6com\vb6com_internal.h'
$harness = Join-Path $root 'tests\run_tests.ps1'

foreach ($p in @($srcDir, $comC, $rtlC, $packC, $varHdr, $comHdr, $harness)) {
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

# R1 —— 目的档 VT_BSTR 的换算不许绕过唯一出口
Expect 'R1 direct VariantChangeType to VT_BSTR' (Count-Pattern $srcDir 'VariantChangeType\s*\([^;\r\n]*VT_BSTR' $true) 0 `
    '目的档 BSTR 只许问 vb6_ComChangeType (布尔那一档口径不同: True/False vs -1/0)'

# R2 —— 唯一出口的定义与两份声明
Expect 'R2 vb6_ComChangeType defined'  (Count-Pattern $comC '(?m)^HRESULT vb6_ComChangeType' $false) 1 `
    '定义只许一枚, 在 vb6rtl_com.c'
Expect 'R2 declared for vb6rtl'        (Count-Pattern $varHdr 'HRESULT vb6_ComChangeType\(' $false) 1 `
    'RTL 侧声明, 在 vb6rtl_variant.h'
Expect 'R2 declared for vb6com'        (Count-Pattern $comHdr 'HRESULT vb6_ComChangeType\(' $false) 1 `
    'vb6com 不 include vb6rtl.h, 声明在 vb6com_internal.h (与 vb6_VariantToDouble 同惯例)'

# R3 —— 消费者四处都还在
Expect 'R3 consumers in vb6rtl_com.c'  (Count-Pattern $comC  'vb6_ComChangeType\(' $false) 4 `
    '定义 1 + 入站归一化 1 + 两条 default 臂 (元素 / 整枚 VARIANT) 2'
Expect 'R3 consumer in vb6rtl.c'       (Count-Pattern $rtlC  'vb6_ComChangeType\(' $false) 1 `
    'vb6_DecimalToBSTR 那一发'
Expect 'R3 consumer in vb6com_pack.c'  (Count-Pattern $packC 'vb6_ComChangeType\(' $false) 1 `
    'vb6_ComUnpackBSTR 那一发'

# R4 —— 布尔那一支的形状
Expect 'R4 bool arm tests VT_TYPEMASK' (Count-Pattern $comC '\(src->vt & VT_TYPEMASK\) == VT_BOOL' $false) 1 `
    '认得 VT_BOOL | VT_BYREF 这一形'
Expect 'R4 True literal'               (Count-Pattern $comC 'L"True"' $false) 1 `
    'VB 语言口径'
Expect 'R4 False literal'              (Count-Pattern $comC 'L"False"' $false) 1 `
    'VB 语言口径'

# R5 —— 判据面（b220 两格各一枚 True/False，改前 -1 只在 Absent 里）
$h = [System.IO.File]::ReadAllText($harness)
Expect 'R5 needle True'   ([regex]::Matches($h, 'val=<True> len=4')).Count 2 `
    'x64/x86 各一枚'
Expect 'R5 needle False'  ([regex]::Matches($h, 'val=<False> len=5')).Count 2 `
    'x64/x86 各一枚'
Expect 'R5 absent old -1' ([regex]::Matches($h, 'VTARG=EchoStr:b giving=VT_BOOL hr=0x00000000 vt=8 val=<-1>')).Count 2 `
    '改前的 -1 拼法只许留在 Absent 名单里 (两格各一枚)'
Expect 'R5 False arm wired' ([regex]::Matches($h, 'VTARG=EchoStr:B')).Count 4 `
    '每格: 正针 1 + Tail 参数 1'

if ($fails.Count -gt 0) {
    Write-Host ('RED: ' + $fails.Count + ' guard(s) broken') -ForegroundColor Red
    exit 1
}
Write-Host 'PASS: VARIANT->目标档只有一位主人 (布尔答 True/False)' -ForegroundColor Green
exit 0
