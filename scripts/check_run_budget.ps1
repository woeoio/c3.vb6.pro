# check_run_budget.ps1 - 逐例跑步预算的结构性哨兵 (只扫 tests/run_tests.ps1, 不起 cl 也不跑 exe)
#
# 为什么单开这一道 (新账 §B140): 「这一枚用例该给几秒」这个决定以前只有全局那一个数, 而真弹模态框的
# GUI 夹具天生吃墙钟余量 —— 门 #484 与门 #491 的 `Tests (vbp #1)` 两次红的是同一形状 (那片工件里
# DlApp.out 停在 DL10、而 DlApp.err 不存在 = 超时被杀, 不是崩)。修法不是把全局 60 抬大 (那等于把所有
# 真挂死的用例一起放出去), 而是加一张**具名逐例**的表。表这种东西一旦能加, 就会长出三种坏法:
# 键写错 (静默按全局走 = "只接了一半")、有人在自己那处就地乘全局数 (第二条答案)、
# 或者干脆把全局数抬了 (判据还在, 但挂死再也看不见)。这三条都能结构性地数出来, 所以钉成哨兵。
#
# 规则 (改坏了会红, 不是装饰):
#   B1  权威只有一处: 解析函数恰好被提 3 次 (定义 + 两个消费点), 表变量恰好 3 次
#       (表本身 + ContainsKey + 取值)。绕过解析函数直接读表 = 数就变
#   B2  表里每个键都必须对得上一枚**真实注册**的夹具 (runner 里找得到 <键>.vbp 或 <键>.bas) ——
#       写错一个键不会报错, 只会静默按全局走, 这正是最难发现的那一类
#   B3  全局默认值仍是 60 (`[int]$RunTimeoutSec = 60,` 逐字在) —— 抬全局就是本账禁止的那种"修红"
#   B4  两个等待点都问权威: Invoke-TestExe 用 `$budgetSec` 那一处恰 1、-Parallel 用逐枚 stamped
#       数那一处恰 1、逐枚问权威那一趟恰 1; 旧的"就地乘全局"两种形状必须 0 处
#   B5  每个预算值必须落在 (全局, 600] —— 不大于全局的数 = 静默收紧; 超过 600 会把一整片的墙钟吃掉,
#       而挂死的用例本来就该被杀
#   B6  表的适用面不许是空的: 至少 1 枚具名夹具 (与 check_fixture_timer_close 的 F3 同一口径 ——
#       哪天 GUI 夹具真不再吃余量, 该红一次让人来决定是撤表还是留着)
#   B7  「起窗自关」探针(RTL vb6forms_ctrl.c)一发就走: 命中标志必须既有写者也有读者, 且上膛
#       那一步要把它归零 —— 这一格决定 B1 那张表要吃多大 (台账 §B157: 命中后没人退, 一枚框
#       被投了 259 发取消, 同一枚 ShowSave 的墙钟从 0.22 s 推到 4.64 s)
#
# 用法:  pwsh -File scripts/check_run_budget.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$runner = Join-Path (Join-Path $root 'tests') 'run_tests.ps1'
if (-not (Test-Path -LiteralPath $runner)) {
    Write-Host ('FAIL runner missing: ' + $runner) -ForegroundColor Red
    exit 1
}
$txt = [System.IO.File]::ReadAllText($runner)
$bad = @()
$TBL = '$RunBudgetSec'
$FN = 'Get-RunBudgetSec'

# ---- B1: 一处权威 ----
$nFn = [regex]::Matches($txt, $FN).Count
if ($nFn -ne 3) {
    $bad += ('B1 the resolver is named ' + $nFn + ' times (want the definition + the TWO wait paths)')
}
$nTbl = [regex]::Matches($txt, [regex]::Escape($TBL)).Count
if ($nTbl -ne 3) {
    $bad += ('B1 the budget table is named ' + $nTbl + ' times (want the table itself + ContainsKey +' +
             ' the read inside the resolver) -> someone reads it behind the resolver')
}

# ---- 逐行把表里的 (键, 秒) 抓出来：先把引号剥掉再配，免得在这份脚本里跟转义较劲 ----
$entries = @()
$mm = [regex]::Match($txt, [regex]::Escape($TBL) + '\s*=\s*@\{[\s\S]*?\r?\n\}')
if (-not $mm.Success) {
    $bad += 'B2 the named-budget table was not found in the runner'
} else {
    $q = [string][char]39
    foreach ($ln in ($mm.Value -split "`r?`n")) {
        $c = $ln.Trim().Replace($q, ' ').Replace('"', ' ').Trim()
        if ($c -match '^(\w+)\s*=\s*(\d+)') {
            $entries += [pscustomobject]@{ Key = $Matches[1]; Sec = [int]$Matches[2] }
        }
    }
}

# ---- B6: 适用面非空 ----
if ($entries.Count -lt 1) {
    $bad += ('B6 the named-budget table holds ' + $entries.Count + ' entries -> B2/B5 would be checking' +
             ' nothing; if GUI fixtures genuinely stopped eating the wall clock, delete the table AND' +
             ' this rule together')
}

# ---- B2: 每个键都必须在表外真实出现过一次（写错的键是静默无效，这一句把它变成红）----
# 表外出现的形式有两种：夹具源文件名（<键>.vbp / <键>.bas）或注册时写的字面名（"rs_gbk_crlf"
# 这种在循环里拼出来的变体也算）⇒ 这里只问「这个键名在 runner 里不是孤零零的」，
# 因为拼错成 DlAp 时它是打不到任何别处的（\b 在 DlApp 里不成立）。
$body = $txt.Remove($mm.Index, $mm.Length)
foreach ($en in $entries) {
    $pat = '\b' + [regex]::Escape($en.Key) + '\b'
    if (-not [regex]::IsMatch($body, $pat, 'IgnoreCase')) {
        $bad += ('B2 key ' + $en.Key + ' appears nowhere outside the table itself -> nothing ever asks' +
                 ' for that name, the entry silently never fires and the fixture keeps the global number')
    }
}

# ---- B3: 全局默认值不动 ----
$globSec = 60
if (-not $txt.Contains('[int]$RunTimeoutSec = 60,')) {
    $bad += 'B3 the global default is no longer 60s -> raising it globally ships every genuinely hung fixture past the watchdog'
}

# ---- B4: 两个等待点都问权威, 旧形状不许回来 ----
# ⚠ PowerShell 会把 @( @(..), @(..) ) 这种嵌套数组字面量**摊平**成一条扁平数组 (踩过两次)，
#   所以下面每一格都用 pscustomobject 写，规则读起来也直白
# ⚠ 要拼接的那条模式先拼成变量：PS 5.1 不接受在 @( .. ) 里的 hashtable 字面量**值**上写 ` + ` 串
#   (整块报 6 个语法错，报错位置还都指向块首，看着像后面全坏了)
$stampPat = 'BudgetMs\s*=\s*\(' + $FN + '\s+\(\[IO\.Path\]::GetFileNameWithoutExtension\(\$it\.Source\)\)'
$pairs = @(
    [pscustomobject]@{ Pat = 'WaitForExit\(\$budgetSec \* 1000\)'; Want = 1; Why = 'Invoke-TestExe must ask the resolver' },
    [pscustomobject]@{ Pat = 'WaitForExit\(\$it\.BudgetMs\)';        Want = 1; Why = 'the -Parallel path must consume the per-item stamp' },
    [pscustomobject]@{ Pat = $stampPat;                             Want = 1; Why = 'the stamping loop that asks the resolver per item' },
    [pscustomobject]@{ Pat = 'WaitForExit\(\$RunTimeoutSec \* 1000\)'; Want = 0; Why = 'the old in-place global multiply in Invoke-TestExe' },
    [pscustomobject]@{ Pat = 'runTimeoutMs\s*=\s*\$RunTimeoutSec \* 1000'; Want = 0; Why = 'the old in-place global multiply in the parallel path' }
)
foreach ($p in $pairs) {
    $n = [regex]::Matches($txt, $p.Pat).Count
    if ($n -ne $p.Want) {
        $bad += ('B4 ' + $p.Why + ': shape found ' + $n + ' times (want ' + $p.Want + ')')
    }
}

# ---- B5: 预算落在 (全局, 600] ----
foreach ($en in $entries) {
    if ($en.Sec -le $globSec -or $en.Sec -gt 600) {
        $bad += ('B5 ' + $en.Key + ' = ' + $en.Sec + 's (want >' + $globSec + ' and <=600): a number at or' +
                 ' under the global is a silent tightening, one over 600 eats the shard wall clock while a' +
                 ' hung fixture should be killed')
    }
}

# ---- B7: 探针一发就走 (同一格账: §B140 的预算被这个机制吃掉的) ----
$rtlProbe = Join-Path $root ('src' + [char]92 + 'rtl' + [char]92 + 'core' + [char]92 + 'vb6forms' + [char]92 + 'vb6forms_ctrl.c')
if (-not (Test-Path -LiteralPath $rtlProbe)) {
    $bad += ('B7 probe source not found: ' + $rtlProbe)
} else {
    $ptxt = [System.IO.File]::ReadAllText($rtlProbe)
    $ppairs = @(
        [pscustomobject]@{ Pat = 'InterlockedExchange(&vb6_CdProbeFired, 0)'; Want = 1;
            Why = 'the arm must clear the shot flag - without that the second Show* sees the first shot and posts nothing' },
        [pscustomobject]@{ Pat = 'InterlockedExchange(&vb6_CdProbeFired, 1)'; Want = 1;
            Why = 'the enum callback still has to mark the shot as delivered' },
        [pscustomobject]@{ Pat = 'InterlockedCompareExchange(&vb6_CdProbeFired, 0, 0)'; Want = 1;
            Why = 'the poll loop must READ the flag - that read is the only exit that ends a delivered shot' },
        [pscustomobject]@{ Pat = 'BOOL found = FALSE'; Want = 0;
            Why = 'the dead exit condition is back (nobody ever wrote it, so the thread kept firing)' },
        [pscustomobject]@{ Pat = 'for (int i = 0; i < 400; i++)'; Want = 0;
            Why = 'a tick count is not a time unit (Sleep(10) measures 15.6ms idle, longer under load)' }
    )
    foreach ($p in $ppairs) {
        # 这几条都是**字面形状** ⇒ 走 Escape 而不是正则: '(int i = 0; i < 400; i++)' 当正则喂进去会
        # 因为 `))+` 是嵌套量词而抛 ArgumentException, 报出来的是脚本自己的错, 不是这一格的判定
        $n = [regex]::Matches($ptxt, [regex]::Escape($p.Pat)).Count
        if ($n -ne $p.Want) {
            $bad += ('B7 ' + $p.Why + ': shape found ' + $n + ' times (want ' + $p.Want + ')')
        }
    }
}

if ($bad.Count -eq 0) {
    Write-Host ('PASS run budget: one resolver, ' + $entries.Count + ' named entries, every key matched a' +
                ' real fixture, global still ' + $globSec + 's, both wait paths ask the authority,' +
                ' probe exits on its own shot') -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1
