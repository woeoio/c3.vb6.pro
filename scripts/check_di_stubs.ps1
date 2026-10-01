# check_di_stubs.ps1 - DI 桩(**定义集**)的 census 哨兵: 只许增, 不许静默减
#
# 为什么单开这一道 (2026-09-30, 门 #233/#234 的教训):
#   scripts/gen_di_stubs.ps1 是**按"跑它那次会话的引用面"发桩**的 —— 换一个工程会话重跑,
#   上一会话引用过的桩会被静默丢掉。本批实测丢 9 个 (comdlg32/winspool 那一族 +
#   ChooseColorA/GetUserNameW/WinHelpA), 连带丢掉 vb6_di_win32_stubs.c 的
#   `#pragma comment(lib,"winspool.lib")`; 丢的时候没有任何东西会响, 只有某个工程恰好
#   Declare 到那个符号时才在链接期以 LNK2019 暴露 (Charts 2020 的 vb6_di_ChooseColorA 就是)。
#
# 为什么是"对基线"而不是"从代码推":
#   引用侧推不出来 —— 发码是把用户的 `Declare ... Lib "x"` 名字**动态**拼成
#   `vb6_di_<name>` (cgen_decl_api.cpp), 静态扫不到; 真正需要的桩 = 全部受测工程 Declare 的并集。
#   定义侧可以精确算出来, 于是把"上次清点过的定义集"冻结成 scripts/di_stubs_manifest.txt:
#     · 少一个  = 红 (这正是重生成丢掉桩的形态);
#     · 多一个  = 提示, 用 -Update 收进基线 (加桩是好事, 只是要人过一眼)。
#
# 用法:
#   pwsh -File scripts\check_di_stubs.ps1            # 对表 (CI/门禁用这条)
#   pwsh -File scripts\check_di_stubs.ps1 -Update    # 把当前定义集写成新基线
# 退出码: 0 = 没掉; 1 = 掉了 (清单打在 stdout)。

param(
    [switch]$Update,
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $PSScriptRoot "di_stubs_manifest.txt"

function Get-SourceText {
    param([string]$Path)
    $t = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    # 去注释 (先块后行): 注释里提到的名字不算定义
    $t = $t -replace '(?s)/\*.*?\*/', ' '
    $t = $t -replace '//[^\r\n]*', ' '
    return $t
}

# --- 定义集: src/rtl/core/di/**/*.c 里 `... vb6_di_X(...) {` ---
$defined = New-Object 'System.Collections.Generic.HashSet[string]'
$diDir = Join-Path $root "src/rtl/core/di"
Get-ChildItem -LiteralPath $diDir -Recurse -Filter *.c | ForEach-Object {
    $t = Get-SourceText $_.FullName
    foreach ($m in [regex]::Matches($t, '\b(vb6_di_[A-Za-z0-9_]+)\s*\([^;{}]*\)\s*\{')) {
        [void]$defined.Add($m.Groups[1].Value)
    }
}
# 枚举/宏里的 vb6_di_* 是**值**不是桩 (vb6rtl_date.c 的 DateInterval: vb6_di_year/…),
# 不参与清点 —— 它们不在 di/ 目录里, 上面的扫法本来也扫不到, 这条注释是给下一个人看的。
$current = @($defined | Sort-Object)

if ($Update -or -not (Test-Path -LiteralPath $manifestPath)) {
    $hdr = @(
        "# DI 桩定义集基线 (scripts/check_di_stubs.ps1 的判据; -Update 重写)",
        "# 生成自: src/rtl/core/di/**/*.c 的函数定义; 少一个就是重生成丢了桩, 必须红。",
        "# 加桩请连同对应家族的 #pragma comment(lib, ...) 一起加。",
        ("# 共 {0} 个。" -f $current.Count)
    )
    Set-Content -LiteralPath $manifestPath -Value ($hdr + $current) -Encoding UTF8
    Write-Host ("基线已写入: {0} ({1} 个定义)" -f $manifestPath, $current.Count)
    exit 0
}

$baseline = @(Get-Content -LiteralPath $manifestPath -Encoding UTF8 |
              Where-Object { $_ -and -not $_.StartsWith('#') } |
              ForEach-Object { $_.Trim() } |
              Where-Object { $_ })
$cur = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($c in $current) { [void]$cur.Add($c) }
$base = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($b in $baseline) { [void]$base.Add($b) }

$dropped = @($baseline | Where-Object { -not $cur.Contains($_) } | Sort-Object)
$added = @($current | Where-Object { -not $base.Contains($_) } | Sort-Object)

Write-Host ("DI 桩 census: 基线 {0} 个, 当前 {1} 个, 掉 {2} 个, 新增 {3} 个" -f `
            $baseline.Count, $current.Count, $dropped.Count, $added.Count)
if ($added.Count -gt 0 -and $Verbose) {
    Write-Host "新增 (确认无误后用 -Update 收进基线):" -ForegroundColor Cyan
    foreach ($a in $added) { Write-Host "  $a" -ForegroundColor Cyan }
}
if ($dropped.Count -gt 0) {
    Write-Host ""
    Write-Host "以下 DI 桩**从 RTL 里消失了** (重生成 di 桩最典型的丢法):" -ForegroundColor Red
    foreach ($d in $dropped) { Write-Host "  $d" -ForegroundColor Red }
    Write-Host ""
    Write-Host "修法: 从上次的生成结果里把定义搬回来 (或改生成器为全语料并集); 若该族还要 .lib," -ForegroundColor Yellow
    Write-Host "      同文件的 #pragma comment(lib, ...) 也要一起搬 (DeviceCapabilitiesA→winspool.lib)。" -ForegroundColor Yellow
    Write-Host "      确认是**故意删**的, 再跑 -Update 重写基线。" -ForegroundColor Yellow
    exit 1
}
Write-Host "OK" -ForegroundColor Green
exit 0
