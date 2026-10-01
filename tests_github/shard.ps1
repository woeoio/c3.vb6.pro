# shard.ps1 - tests_github 三个门禁脚本共用的分片算子
#   用法: . (Join-Path $PSScriptRoot 'shard.ps1')   然后   $q = Select-ShardSlice -Items $q ...
#
# 为什么单开一个文件: 切片的算式原本在 run_t0 / run_t1 / run_t2 (两份) 里各抄了一遍, 而
# 分片总数是 CI 矩阵上随时可调的旋钮 —— 四份手抄的算式改一处漏三处, 下次调 -ShardTotal
# 就会有一层静默漏跑或重跑。收成一处算子, 脚本只调用、不再自己写算法。
#
# 语义 (与 tests\run_tests.ps1 的 bas/vbp 分片同一口径):
#   -ShardTotal <= 1  = 不分片, 原样返回完整清单 (今天的执行路径, 输出逐字不变)。
#   -ShardTotal > 1   = 按**清单次序**切 Ceiling(N/total) 的连续段, 每个 runner 只跑第 Shard 段。
#   调用方必须先对**完整**清单跑条数守卫, 再来切执行集 —— 守卫里的条数是脚本自带的契约,
#   不能因为分片而只对 1/N 的清单校验。
#
# 幂等/可复现: 只依赖清单次序与 Shard/ShardTotal, 不依赖时钟、随机数或环境变量。

# 结构哨兵: 并行段"输入非空、一个 worker 都没回来"必须红。门禁最不能容忍的是看起来绿了
# 其实一条没跑 —— 实测 run_t1.ps1 在 Windows PowerShell 5.1 下: ForEach-Object -Parallel 不
# 存在, 叠加脚本顶部的 SilentlyContinue, 64 条 bas 任务静默变成 "0 worker(s), pass= fail="
# 而退出码仍是 0。三个脚本的串行/并行分岔写法各不相同 (t1 根本没有串行兜底), 所以统一挡在
# 这里, 而不是每处再抄一遍 if。
function Assert-ParallelRan {
    [CmdletBinding()]
    param(
        [object[]]$Items,
        [object]$Results,
        [string]$Label = 'parallel section'
    )
    $n = 0
    if ($null -ne $Items) { $n = @($Items).Count }
    if ($n -gt 0 -and @($Results).Count -eq 0) {
        Write-Host "[FATAL] $Label : $n 条任务, 但 0 个并行 worker 返回 (本层需要 PowerShell 7; 当前 PS $($PSVersionTable.PSVersion))" -ForegroundColor Red
        exit 1
    }
}

function Select-ShardSlice {
    [CmdletBinding()]
    param(
        [object[]]$Items,
        [int]$Shard,
        [int]$ShardTotal,
        [string]$Label = 'item'
    )
    if ($ShardTotal -le 1) { return @($Items) }
    if ($Shard -lt 1 -or $Shard -gt $ShardTotal) {
        Write-Host "[ERROR] Shard 需在 1..ShardTotal (当前 $Shard/$ShardTotal)" -ForegroundColor Red
        exit 1
    }
    $all = @($Items | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0) {
        Write-Host "  ($Label shard ${Shard}/${ShardTotal}: 0 项)" -ForegroundColor Cyan
        return @()
    }
    $len = [int][Math]::Ceiling($all.Count / [double]$ShardTotal)
    $start = ($Shard - 1) * $len
    # 片数多于条目时靠后的片是空手。必须真的返回空集: $all[$start..$end] 在 start 越界时
    # 给的是 $null, 而 @($null) 是"含一个 $null 元素"的数组, 下游会把它当一条用例去跑。
    if ($start -ge $all.Count) {
        Write-Host "  ($Label shard ${Shard}/${ShardTotal}: 0 项)" -ForegroundColor Cyan
        return @()
    }
    $end = [int][Math]::Min($start + $len - 1, $all.Count - 1)
    $slice = @($all[$start..$end])
    Write-Host "  ($Label shard ${Shard}/${ShardTotal}: $($slice.Count) 项)" -ForegroundColor Cyan
    return $slice
}
