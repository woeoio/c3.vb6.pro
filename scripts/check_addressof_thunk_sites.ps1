# check_addressof_thunk_sites.ps1 - 账 #184 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: `AddressOf` 交出去的必须是**桩地址**, 不是本体地址。x86 下本体是
# __cdecl 而 OS/COM 按 __stdcall 回调 ⇒ 每次回调把栈少弹 N*4 字节。实测 VBFlexGrid 的
# SUBCLASSPROC (6 形参 = 24 字节) 因此起窗即堆损坏 0xC0000374 (本地 3/4 崩), 而那份 demo
# 当时完全在门禁之外 ⇒ 缺陷藏了很久。
#
# 口径现在收在一处: CCodeGen::addressOfTargetCName (src/backend/module/cgen_delegate.cpp)。
# 这条哨兵拦的是"又加一个把 (void*) 直接拼 cProcName 的 AddressOf 出口" —— 那种写法在 x64
# 上一点症状都没有, 只有 x86 运行期才炸, 所以必须静态钉住。
#
# 规则 (改坏了会红, 不是装饰):
#   R1  src/backend 里 `"(void*)" + cProcName` 形式的出口 = 0 处 (那是旧写法)
#   R2  `"(void*)" + addressOfTargetCName` 的调用点 >= 2 处 (AddressOf 两条形状都接上了)
#   R3  标记趟在发码之前被调用, 且只调一次 (stage 3.5c)
#   R4  桩的原型点与定义点各恰好一处 (decl pass / epilogue)
#
# 用法:  pwsh -File scripts\check_addressof_thunk_sites.ps1
# 退出码: 0 = 全绿; 1 = 红 (每条失败单独打一行)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function Get-Text([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    return [System.IO.File]::ReadAllText($path)
}

$backendDir = Join-Path $root "src\backend"
$files = Get-ChildItem -LiteralPath $backendDir -Recurse -File |
         Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" }

$rawCount = 0
$rawWhere = @()
$wrapped = 0
foreach ($f in $files) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($t, '"\(void\*\)"\s*\+\s*cProcName')) {
        $rawCount++
        $line = 1 + (($t.Substring(0, $m.Index)) -split "`n").Count
        $rawWhere += ($f.Name + ":" + $line)
    }
    $wrapped += ([regex]::Matches($t, '"\(void\*\)"\s*\+\s*addressOfTargetCName')).Count
}
if ($rawCount -ne 0) {
    $bad += ("R1 raw `"(void*)`" + cProcName AddressOf 出口 = {0} 处 (应为 0): {1}" -f $rawCount, ($rawWhere -join ", "))
}
if ($wrapped -lt 2) { $bad += ("R2 addressOfTargetCName 调用点 = {0} 处 (应 >= 2)" -f $wrapped) }

$compile = Get-Text (Join-Path $root "src\driver\driver_compile.cpp")
if ($null -eq $compile) {
    $bad += "R3 driver_compile.cpp 不在位"
} else {
    $n = ([regex]::Matches($compile, 'markAddressOfCallbacks\(\)\s*;')).Count
    if ($n -ne 1) { $bad += ("R3 markAddressOfCallbacks() 调用次数 = {0} (应为 1)" -f $n) }
    $iMark = $compile.IndexOf("markAddressOfCallbacks();")
    $iGen = $compile.IndexOf("runCodeGeneration(effectiveOpts")
    if ($iGen -lt 0) { $iGen = $compile.IndexOf("runCodeGeneration(options") }
    if ($iMark -lt 0 -or $iGen -lt 0 -or $iMark -gt $iGen) {
        $bad += ("R3 标记趟必须早于发码 (mark@{0} gen@{1})" -f $iMark, $iGen)
    }
}

$declPass = Get-Text (Join-Path $root "src\backend\detail\base\cgen_base_generate_decl_pass.inc")
$epilogue = Get-Text (Join-Path $root "src\backend\detail\base\cgen_base_generate_epilogue.inc")
if ($null -eq $declPass -or ([regex]::Matches($declPass, 'emitAddressOfThunkDecls\(module\)')).Count -ne 1) {
    $bad += "R4 decl pass 必须恰好一处 emitAddressOfThunkDecls(module)"
}
if ($null -eq $epilogue -or ([regex]::Matches($epilogue, 'emitAddressOfThunkDefs\(\)')).Count -ne 1) {
    $bad += "R4 epilogue 必须恰好一处 emitAddressOfThunkDefs()"
}

if ($bad.Count -eq 0) {
    Write-Output ("OK addrof-thunk-sites: raw=0 wrapped={0} decl+epilogue 各一处" -f $wrapped)
    exit 0
}
$bad | ForEach-Object { Write-Output ("FAIL " + $_) }
exit 1
