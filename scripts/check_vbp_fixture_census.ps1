# check_vbp_fixture_census.ps1 - 「每个 .vbp 夹具都必须先入册」哨兵
#
# 为什么单开这一道 (2026-10-06, 隐患留档的后半截):
#   dbgdlg 的 `vb6_di_PageSetupDlgA` 缺桩是**真红**, 可它在门禁里连"红"都算不上 ——
#   因为 dbgdlg 这个夹具压根不在 tests/run_tests.ps1 的任何清单里, 门禁从不编它。
#   "编不过的东西, 在门禁里等于不存在" (#188 那句的同源形态)。单把一个 dbgdlg 补进去
#   只修了这一枚; 这一类(夹具建好却没人编)照样能再发生。所以这里把"入册"本身变成判据。
#
# 判据: tests/**/*.vbp 每一个都必须满足其一 ——
#   ① 被 tests/run_tests.ps1 以 `$Tests\<相对路径>` 字面引用 (即真被门禁编/跑);
#   ② 出现在 scripts/vbp_fixtures_unregistered.txt 里, 且写明理由 (刻意不登记/别处已覆盖)。
#   两者都不满足 = 红。加夹具的人必须做一次显式选择, 不许"建完就算完"。
#
# 为什么按"run_tests.ps1 原文里有没有这个字面路径"判, 而不是真去跑门禁:
#   夹具的登记方式就是那一行字面路径 (全 298 处 .vbp 出现里, 只有 1 处是 Join-Path 动态拼的,
#   且那一处的源夹具另有字面登记)。静态判据够准、够快, 也不依赖工具链。
#
# 范围 (两处刻意圈外, 别以为漏了):
#   · 只看 **git 已跟踪**的 .vbp —— tests\VBFlexGridDemo\bisect1..7.vbp 是本机 bisect 时留下的
#     散件 (已被 .gitignore 忽略), 拿它们报红只是噪声。CI 上工作区就是 checkout 出来的,
#     两者等价。
#   · 不看 tests_github\ —— 那边的 t2_cases\ 是 tests\ 同名夹具的**镜像集**, 另由 ci_t0.yml
#     那条线驱动, 列在这里只会把同一份夹具数两遍。
#
# 用法:
#   powershell -File scripts\check_vbp_fixture_census.ps1            # 对表 (CI/门禁用这条)
#   powershell -File scripts\check_vbp_fixture_census.ps1 -Update    # 把当前未登记集写成新基线
# 退出码: 0 = 全部入册; 1 = 有夹具两头不靠 (或基线里有已消失的条目)。
#   注: -Update 只写"当前未登记的那个集合", 理由一栏留给人工 —— 新进来的条目会标 TODO。

param(
    [switch]$Update,
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"
$root      = Split-Path -Parent $PSScriptRoot
$testsDir  = Join-Path $root "tests"
$runner    = Join-Path $testsDir "run_tests.ps1"
$listPath  = Join-Path $PSScriptRoot "vbp_fixtures_unregistered.txt"

if (-not (Test-Path $runner)) { Write-Host "[ERR] 找不到 $runner" -ForegroundColor Red; exit 1 }
$runnerText = Get-Content -LiteralPath $runner -Raw -Encoding UTF8

# --- 全量夹具 (相对 tests\ 的路径, 反斜杠) ---
# 优先 git 已跟踪集合 (排除本机散件); git 不可用时退回文件系统扫描。
$all = @()
$fromGit = $false
try {
    $prevDir = (Get-Location).Path
    Set-Location -LiteralPath $root
    $tracked = & git ls-files -- tests 2>$null
    Set-Location -LiteralPath $prevDir
    if ($LASTEXITCODE -eq 0 -and $tracked) {
        $fromGit = $true
        foreach ($t in $tracked) {
            $t = $t.Trim()
            if ($t -notmatch '\.vbp$') { continue }
            if ($t -notmatch '^tests[/\\]') { continue }
            $all += $t.Substring(6).Replace('/', '\')     # 去掉 'tests\' / 'tests/'
        }
    }
} catch { $fromGit = $false }
if (-not $fromGit) {
    Write-Host "提示: 未取到 git 跟踪集合 (git 不可用?), 退回文件系统扫描 —— 本机散件可能被算进来。" -ForegroundColor Yellow
    Get-ChildItem -LiteralPath $testsDir -Recurse -Filter *.vbp -File | ForEach-Object {
        $rel = $_.FullName.Substring($testsDir.Length + 1).Replace('/', '\')
        $all += $rel
    }
}
$all = @($all | Sort-Object -Unique)

# --- 已入册: run_tests.ps1 里出现过 `\<相对路径>` ---
$registered = @($all | Where-Object { $runnerText.Contains('\' + $_) })
$candidates = @($all | Where-Object { -not $runnerText.Contains('\' + $_) })

# --- 基线: `相对路径` 可选接 ` # 理由` ---
$baseline = @{}
if (Test-Path $listPath) {
    foreach ($ln in (Get-Content -LiteralPath $listPath)) {
        $t = $ln.Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        $name = ($t -split '\s+#')[0].Trim()
        if ($name) { $baseline[$name] = $t }
    }
}

if ($Update) {
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add("# .vbp 夹具未登记基线 (scripts/check_vbp_fixture_census.ps1 的判据; -Update 重写)")
    $lines.Add("# 格式: <相对 tests\ 的路径>   # <理由>")
    $lines.Add("# 出现在这里 = 承认它**不在** tests\run_tests.ps1 门禁的编译面上。门禁只认两类理由:")
    $lines.Add("#   (a) 刻意不列  —— 编不过 / 是已知红, 列进去就是把已知的红当基线;")
    $lines.Add("#   (b) 别处已覆盖 —— 由别的 runner 覆盖 (tests\run_vbman_test.ps1 / tests\regress_all.ps1)。")
    $lines.Add('# 凡**能确定编得过**的一律登记进 run_tests.ps1 的 Test-VbpBuild —— 那只是「不跑」, 不是「不编」。')
    $lines.Add('# -Update 只重排上面这几行与末尾的计数; 每条的理由原样保留, 新的带 TODO (请当天补上)。')
    $lines.Add("# 共 " + $candidates.Count + " 个。")
    foreach ($c in $candidates) {
        if ($baseline.ContainsKey($c)) { $lines.Add($baseline[$c]) }
        else { $lines.Add($c + "   # TODO: 说明为何不登记 (或直接登记进 run_tests.ps1)") }
    }
    Set-Content -LiteralPath $listPath -Value $lines -Encoding UTF8
    Write-Host ("已重写基线: " + $candidates.Count + " 个 -> " + $listPath) -ForegroundColor Green
    exit 0
}

# --- 对表 ---
$unaccounted = @($candidates | Where-Object { -not $baseline.ContainsKey($_) })
$ghost       = @($baseline.Keys | Where-Object { -not ($all -contains $_) })
$stale       = @($baseline.Keys | Where-Object { $registered -contains $_ })
$todo        = @($baseline.Keys | Where-Object { $baseline[$_] -match '#\s*TODO' })

if ($Verbose) {
    Write-Host ("夹具 " + $all.Count + " 个: 已入册 " + $registered.Count + ", 未登记 " + $candidates.Count)
}

$bad = $false
if ($unaccounted.Count -gt 0) {
    $bad = $true
    Write-Host ("以下 .vbp 夹具既没登记进 run_tests.ps1, 也没写进基线 (共 " + $unaccounted.Count + " 个):") -ForegroundColor Red
    foreach ($u in $unaccounted) { Write-Host ("  " + $u) -ForegroundColor Red }
    Write-Host "  处置: 该编的请登记进 tests\run_tests.ps1; 刻意不编的请跑 -Update 或手工补进基线并写明理由。" -ForegroundColor Yellow
}
if ($ghost.Count -gt 0) {
    $bad = $true
    Write-Host ("基线里有 " + $ghost.Count + " 条指向已不存在的夹具 (清理掉):") -ForegroundColor Red
    foreach ($g in $ghost) { Write-Host ("  " + $g) -ForegroundColor Red }
}
if ($stale.Count -gt 0 -and $Verbose) {
    Write-Host ("提示: 基线里 " + $stale.Count + " 条其实已经登记进 run_tests.ps1 了, 可以从基线删掉:") -ForegroundColor DarkGray
    foreach ($s in $stale) { Write-Host ("  " + $s) -ForegroundColor DarkGray }
}
if ($todo.Count -gt 0) {
    Write-Host ("提示: 基线里还有 " + $todo.Count + " 条理由写着 TODO, 请补上真实理由。") -ForegroundColor Yellow
}

if ($bad) { Write-Host "FAIL" -ForegroundColor Red; exit 1 }
Write-Host ("夹具 census: 共 " + $all.Count + " 个, 入册 " + $registered.Count + " 个, 基线挂账 " + $baseline.Count + " 个, 两头不靠 0 个") -ForegroundColor Green
Write-Host "OK"
exit 0
