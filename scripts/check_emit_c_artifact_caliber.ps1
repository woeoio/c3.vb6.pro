# check_emit_c_artifact_caliber.ps1 - 账 #267: --emit-c 的产物字节只有一条路，而且不过控制台代码页。
#
# 为什么钉形状：C3 把发码写成中间目录里的 .h/.c 交 cl 编 (命令行带 /utf-8)，同时 --emit-c 又把同一份东西
# 打到 stdout 给人接。老写法走 std::cout，而 stdout 被 ConsoleUtf8Buf 接管 —— 管道/文件那一支按控制台
# 代码页转字节，于是本机 (936) 交 GBK、CI 交 UTF-8：「这次发码长什么样」有了两个答案，其中一个随机器变，
# 跨机的发码清单因此永远对不齐。修法是把产物那一路收成一个出口 (encoding.hpp 的 writeStdoutRaw，直接
# fwrite 内部 UTF-8)。行为本身由 tests/run_tests.ps1 的 Test-EmitcByteCaliber 管 (两档代码页各跑一趟，
# 和中间目录那两份逐字节比)；这条哨兵管的是「别再退回 cout」与「出口/消费者名单没长」。
$ErrorActionPreference = 'Continue'
if (-not $PSScriptRoot) { exit 2 }
$Root = Split-Path -Parent $PSScriptRoot
$fail = 0
function Report([string]$m) { Write-Host $m; $script:fail++ }

$inc = Join-Path $Root "src\driver\detail\driver_codegen_module_loop.inc"
$enc = Join-Path $Root "src\common\encoding.hpp"
foreach ($p in @($inc, $enc)) { if (-not (Test-Path $p)) { Report ("MISSING " + $p) } }
if ($fail -gt 0) { exit 1 }
$incText = [System.IO.File]::ReadAllText($inc)
$encText = [System.IO.File]::ReadAllText($enc)
$sources = @(Get-ChildItem (Join-Path $Root "src") -Recurse -File -Include *.cpp, *.hpp, *.inc)

# C1: emitC 那一段必须走 writeStdoutRaw；全仓不许再有 std::cout << cgen. (那一份是给终端的)。
$i = $incText.IndexOf("if (options.emitC)")
if ($i -lt 0) {
    Report "C1 --emit-c 的发码分支在 driver_codegen_module_loop.inc 里找不到了 (挪地方要连这条一起改)"
} else {
    $seg = $incText.Substring($i, [Math]::Min(1400, $incText.Length - $i))
    if ($seg -notmatch 'writeStdoutRaw\(') { Report "C1 emitC 分支没走 writeStdoutRaw (产物又回 cout 了)" }
    if ($seg -notmatch 'std::cout\.flush') { Report "C1 emitC 分支没先 cout.flush (诊断会与产物错位)" }
}
$leaked = @($sources | Where-Object { (Get-Content -Raw $_.FullName) -match 'std::cout << cgen\.' })
if ($leaked.Count -gt 0) { Report ("C1 还有 std::cout << cgen. 的写法: " + (($leaked | ForEach-Object { $_.Name }) -join ',')) }

# C2: 唯一出口的定义必须落 fwrite，不许自己又去转码。
$d = $encText.IndexOf("inline void writeStdoutRaw")
if ($d -lt 0) { Report "C2 encoding.hpp 里 writeStdoutRaw 的定义没了" }
else {
    $body = $encText.Substring($d, [Math]::Min(700, $encText.Length - $d))
    if ($body -notmatch 'std::fwrite') { Report "C2 writeStdoutRaw 不再用 std::fwrite 落字节" }
    if ($body -match 'std::cout|WideCharToMultiByte') { Report "C2 writeStdoutRaw 里又出现 cout / 代码页转换" }
}

# C3: 出口调用点恰好 1 处 (定义不算)。今天实测 = driver_codegen_module_loop.inc 一处。
$calls = @($sources | Select-String -Pattern 'writeStdoutRaw\(' |
    Where-Object { $_.Line -notmatch 'inline void writeStdoutRaw' } |
    ForEach-Object { ($_.Path.Substring($Root.Length + 1) -replace '\\', '/') + ':' + $_.LineNumber })
if ($calls.Count -ne 1) { Report ("C3 writeStdoutRaw 的调用点应恰好 1 处，实测 " + $calls.Count + " 处: " + ($calls -join ',')) }

# C4: options.emitC 的消费者恰好 2 处，且名单就是这两份文件 (发码那一路 + runLinker 早退)。
$cons = @($sources | Select-String -Pattern 'options\.emitC' |
    ForEach-Object { $_.Path.Substring($Root.Length + 1) -replace '\\', '/' })
$want = @('src/driver/detail/driver_codegen_module_loop.inc', 'src/driver/driver_link.cpp')
$gotNames = @($cons | Sort-Object -Unique)
if (@($cons).Count -ne 2 -or (Compare-Object $want $gotNames).Count -gt 0) {
    Report ("C4 options.emitC 的消费者不再是「2 处 / 这两份文件」: 行数=" + @($cons).Count + " 名单=" + ($gotNames -join ','))
}

if ($fail -eq 0) {
    Write-Host "emit_c_artifact_caliber OK: 产物出口 1 处 / options.emitC 消费者 2 处 / 无 cout 回潮"
    exit 0
}
Write-Host ("emit_c_artifact_caliber FAILED: " + $fail + " 条")
exit 1
