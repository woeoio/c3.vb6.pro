# check_var_arr_inbound.ps1 — 账 #610 第344刀: VT_ARRAY|VT_VARIANT 入站档的静态哨兵。
#
# 为什么要有它 (形状门看不见的那部分):
#   · 形状门只哈希 --emit-c 的 stdout ⇒ **RTL 内部**改了什么它一个字都不管;
#   · 这一刀的运行期证人 (服务端读到四枚元素的真值) 只有在真 COM 客户端在场时才有;
#     静态哨兵是**兜底**: 有人把逐枚那一支删掉/改回 memcpy 时立刻红, 不用等真跑;
#   · 「逐枚」和「整块 memcpy」的区别全在**几个关键字**上 ⇒ 必须逐字钉住。
#
# 判据:
#   R1  闸门处必须**先**问 `base == VT_VARIANT` 再决定走白名单 (VT_VARIANT 不问白名单)
#   R2  逐枚分支里必须有 VariantCopy (深拷贝) —— 少了它就是搬对端的指针 ⇒ 悬垂/双释放
#   R3  逐枚分支里必须有消耗式转换 (steal=1) + VariantClear (释放那份私有副本)
#   R4  **不得**出现整块 memcpy 掉 VT_VARIANT 的形态 (memcpy 只许留在非 variant 那一支)
#   R5  描述符的 elemSize 必须来自 vb6_SaElemSize (与写入步长同源 ⇒ 不错位)
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot)
)
$ErrorActionPreference = "Stop"
$problems = @()

$Root = (Resolve-Path $Root).Path
$srcPath = Join-Path $Root "src\rtl\core\vb6rtl\vb6rtl_com.c"
if (-not (Test-Path $srcPath)) {
    Write-Output "FAIL: 找不到 src/rtl/core/vb6rtl/vb6rtl_com.c"
    exit 1
}
$src = Get-Content $srcPath -Raw
$lines = @($src -split "`r?`n")

function Find-Line([string]$needle) {
    # ⚠ 循环变量不能叫 $i (PowerShell 函数默认读写父作用域 ⇒ 会污染调用方的 $i)
    $fl = 0
    while ($fl -lt $lines.Count) {
        $hit = $lines[$fl].Contains($needle)
        if ($hit) { return $fl }
        $fl++
    }
    return -1
}

# R1: 闸门处先问 VT_VARIANT
$iIsVar = Find-Line 'isVarElem = (base == VT_VARIANT);'
$iAllow = Find-Line 'allow = vb6_ComSaVtBlittable(et);'
if ($iIsVar -lt 0 -or $iAllow -lt 0) {
    $problems += "R1: 找不到闸门处的 isVarElem 判定或白名单那一问"
} elseif ($iAllow -lt $iIsVar) {
    $problems += "R1: 白名单那一问排在 isVarElem 判定**之前** ⇒ VT_VARIANT 元素档仍落在白名单外 (整枚数组转 Empty)"
}

# R2/R3: 逐枚分支的四件套
$iCopy = Find-Line 'VariantCopy(&elem, srcElem);'
$iConv = Find-Line 'vb6_ComVariantIntoSlot(&elem, 1)'
$iClear = Find-Line 'VariantClear(&elem);'
if ($iCopy -lt 0) {
    $problems += "R2: 逐枚分支里没有 VariantCopy ⇒ 搬的是**对端拥有**的指针 ⇒ 对端 free 槽阵列后 ours 悬垂 + 双释放"
}
if ($iConv -lt 0) {
    $problems += "R3: 逐枚分支里没有消耗式转换 vb6_ComVariantIntoSlot(&elem, 1)"
}
if ($iClear -lt 0) {
    $problems += "R3: 逐枚分支里没有 VariantClear(&elem) ⇒ 那份私有副本没人释放 (每元素漏一枚)"
}
if ($iCopy -ge 0 -and $iConv -ge 0 -and $iClear -ge 0) {
    if (-not ($iCopy -lt $iConv -and $iConv -lt $iClear)) {
        $problems += "R3: 三步的先后不对 (应 VariantCopy → 消耗式转换 → VariantClear)"
    }
}

# R4: memcpy 只许留在非 variant 那一支 (else 里)
$iMemcpy = Find-Line 'memcpy(data, srcData, (size_t)n * (size_t)esz);'
$iElse = Find-Line '    } else {'
if ($iMemcpy -lt 0) {
    $problems += "R4: 找不到 memcpy 那一支 (非 variant 元素档仍要整块搬)"
} elseif ($iElse -ge 0 -and $iMemcpy -gt $iElse) {
    # memcpy 落在 else 之后 ⇒ 只在非 variant 档走 ⇒ 对
    $ok = $true
} else {
    $problems += "R4: memcpy 落在 else **之前** ⇒ VT_VARIANT 元素档被整块搬 (搬的是对端的指针)"
}

# R5: elemSize 来自 vb6_SaElemSize (与写入步长同源)
$iElemSize = Find-Line 'esz = vb6_SaElemSize(et);'
$iWrite = Find-Line 'vb6_VARIANT* dstElem = (vb6_VARIANT*)((char*)data + i * esz);'
if ($iElemSize -lt 0 -or $iWrite -lt 0) {
    $problems += "R5: 找不到 esz 的来源或写入步长那一行 (步长必须与描述符 elemSize 同源)"
}

if ($problems.Count -eq 0) {
    Write-Output "OK: VT_ARRAY|VT_VARIANT 入站档 (R1..R5) 全部通过"
    exit 0
} else {
    Write-Output "FAIL: VT_ARRAY|VT_VARIANT 入站档哨兵"
    foreach ($p in $problems) { Write-Output "  - $p" }
    exit 1
}
