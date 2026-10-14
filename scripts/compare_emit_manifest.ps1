# compare_emit_manifest.ps1 - 把一份发码清单与 checked-in 的期望清单逐行比；不同就非零退出。
#
# 为什么单独一个脚本、而且是唯一一份比较逻辑：CI 那一步（emit_manifest.ps1 出完清单后调它）与本机
# 复算/负控/重新登记必须走同一段代码，否则「CI 上绿、本机换个说法就红」这种口径分家又会长出来。
# 背景（台账 §B73，账 #245 + #267）：门 #418 实测 CI 与本机两台的全语料清单 395/395 行 sha256 相同，
# 于是「这次发码长什么样」第一次可以判红，而不是只留档。
#
# 用法:
#   powershell -File scripts/compare_emit_manifest.ps1 -Manifest <清单> [-Expect <期望>]
#   powershell -File scripts/compare_emit_manifest.ps1 -Manifest <清单> [-Expect <期望>] -Bless [-AllowVanish]
# 退出码: 0 相同 / 1 不一致（或 -Bless 被退场那条挡住）/ 2 文件或参数缺失
param(
    [string]$Manifest = "",
    [string]$Expect = "",
    [switch]$Bless,
    [switch]$AllowVanish
)
$ErrorActionPreference = 'Continue'
if (-not $Manifest) { Write-Host "用法: -Manifest <清单路径>"; exit 2 }
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Expect) { $Expect = Join-Path $repoRoot "emit-manifest.expected.txt" }
if (-not (Test-Path $Manifest)) { Write-Host ("MISSING " + $Manifest); exit 2 }
if (-not (Test-Path $Expect)) { Write-Host ("MISSING " + $Expect); exit 2 }

# 只比**数据行**。'#' 开头的是身份注脚（c3-exe 哈希、pe-lnk、runner 名、inputs= 之类），
# 天生随机器/随构建变 —— 留在清单里给人看，不参与判定。
function DataRows([string]$path) {
    @(Get-Content -LiteralPath $path -Encoding UTF8 | Where-Object { $_ -match '\S' -and $_ -notmatch '^#' })
}
function HeaderRows([string]$path) {
    $all = @(Get-Content -LiteralPath $path -Encoding UTF8)
    $stop = 0
    foreach ($ln in $all) {
        if ($ln -match '\S' -and $ln -notmatch '^#') { break }
        $stop++
    }
    @($all[0..([Math]::Max($stop - 1, 0))])
}
$got = DataRows $Manifest
# 自证覆盖率：清单自己的注脚里写着 `inputs=<N>`（emit_manifest.ps1 数的是它递给 C3.exe 的输入条数）。
# 判定只比 '#' 以外的行 —— 万一有人（或某次改动）把一**整行数据**折成注释，两边会一起少掉同一行、
# 于是"全对"而覆盖悄悄缩水。注脚与数据行各数各的，对不上就是这里出了问题。
$declared = 0
foreach ($h in @(Get-Content -LiteralPath $Manifest -Encoding UTF8 | Where-Object { $_ -match '^#\s*arch=.*inputs=(\d+)' })) {
    $declared = [int]([regex]::Match($h, 'inputs=(\d+)').Groups[1].Value)
}
if ($declared -gt 0 -and $declared -ne $got.Count) {
    Write-Host ("发码清单自相矛盾: 注脚 inputs=" + $declared + " 而数据行 " + $got.Count +
                " —— 有输入没进判定（别把数据行改成 '#' 开头）")
    exit 1
}
# 行尾那段就是输入路径（compare 的键），也是「登记时谁的顺序不动」的依据。
function PathOfRow([string]$r) { $i = $r.LastIndexOf(' '); if ($i -gt 0) { return $r.Substring($i + 1) } return $r }

$exp = DataRows $Expect

if ($Bless) {
    $hdr = @(HeaderRows $Expect)
    # 【登记不许吃注释行】HeaderRows 只认**头部**那一块（到第一条数据行为止），而每刀的「登记来源」是写在清单**末尾**的（门 #501 实测：这一档原先会静默吃掉末尾两条来源注脚，22 行注释变 20 行，还照样报「已登记」）。
    # 所以末尾那段单独收下来，落盘前两头各数一遍：注释总数对不上就拒绝登记，而不是少一条也照样整写。
    $expLines = @(Get-Content -LiteralPath $Expect -Encoding UTF8)
    $hdrCmt = @($hdr | Where-Object { $_ -match '^#' }).Count
    $tailCmt = @()
    for ($i = $hdr.Count; $i -lt $expLines.Count; $i++) {
        if ($expLines[$i] -match '^#') { $tailCmt += [string]$expLines[$i] }
    }
    $cmtAll = @($expLines | Where-Object { $_ -match '^#' }).Count
    if (($hdrCmt + $tailCmt.Count) -ne $cmtAll) {
        Write-Host ('拒绝登记：期望清单共 ' + $cmtAll + ' 行注释，头部 ' + $hdrCmt + ' + 末尾 ' +$tailCmt.Count + ' 对不上 —— 登记会把对不上的那几行吃掉')
        exit 2
    }
    # 登记**保持现有行的次序**，新出现的输入按清单自己的顺序追加在末尾。
    # 为什么：语料枚举的次序与这张清单的历史次序实测差 68 个位置（同一批路径、两种排法），
    # 照清单顺序整写一遍 = git diff 里 398 行全动，而真正变的只有 7 行 —— 别人review 时看不出
    # 哪里是真变化（门 #434 之后本线手 splice 那 29 行，就是为了躲这个）。
    $newByPath = [ordered]@{}
    foreach ($r in $got) { $newByPath[(PathOfRow $r)] = $r }
    $ordered = New-Object System.Collections.Generic.List[string]
    $kept = @{}
    foreach ($r in $exp) {
        $p = PathOfRow $r
        if ($newByPath.Contains($p)) { $ordered.Add([string]$newByPath[$p]); $kept[$p] = $true }
    }
    $vanished = @($exp | ForEach-Object { PathOfRow $_ } | Where-Object { -not $kept.ContainsKey($_) })
    $appended = @($newByPath.Keys | Where-Object { -not $kept.ContainsKey($_) })
    if ($vanished.Count -gt 0 -and -not $AllowVanish) {
        Write-Host ('拒绝登记：有 ' + $vanished.Count + ' 份输入在期望清单里、在这份清单里没有（' +
                    (($vanished | Select-Object -First 6) -join ', ') + '）')
        Write-Host '  这通常是**夹具被删/改名**或**语料枚举口径变了**，不是发码变了；用登记把它抹平 = 悄悄缩小覆盖面。'
        Write-Host '  确认是有意为之再加 -AllowVanish。'
        exit 1
    }
    foreach ($p in $appended) { $ordered.Add([string]$newByPath[$p]) }
    # 编码**不能交给 Set-Content**：Windows PowerShell 5.1 的 -Encoding UTF8 带 BOM、pwsh 7 的不带，
    # 而 CI 那一步跑的就是 pwsh 7。这张期望清单是被 check_manifest_coverage.ps1 的 K2 钉成
    # 「BOM + 纯 CRLF」的（没 BOM 时 5.1 会把中文注脚按 ANSI 读坏），所以这里自己写死，
    # 两个版本的 PowerShell 出来必须是同一串字节。
    $text = ((@($hdr) + $ordered + $tailCmt) -join "`r`n") + "`r`n"
    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Resolve-Path $Expect).Path, $text, $enc)
    $changed = @($exp | Where-Object { $kept.ContainsKey((PathOfRow $_)) -and ([string]$newByPath[(PathOfRow $_)]) -ne $_ }).Count
    Write-Host ('已登记 ' + $Expect + '：注脚 ' + ($hdrCmt + $tailCmt.Count) + ' 行（头部 ' + $hdrCmt + ' + 末尾 ' + $tailCmt.Count + '）+ 数据行 ' + $ordered.Count +
                ' 行（改写 ' + $changed + ' / 新增 ' + $appended.Count + ' / 退场 ' + $vanished.Count + '）')
    exit 0
}

# 键取行尾那段输入路径（语料顺序变了也不该误报），值取前面所有哈希/计数字段
function AsMap($rows) {
    $h = [ordered]@{}
    foreach ($r in $rows) {
        $i = $r.LastIndexOf(' ')
        if ($i -gt 0) { $h[$r.Substring($i + 1)] = $r.Substring(0, $i) }
    }
    return $h
}
$me = AsMap $exp
$mg = AsMap $got
$missing = @($me.Keys | Where-Object { -not $mg.Contains($_) })
$extra = @($mg.Keys | Where-Object { -not $me.Contains($_) })
$diff = @($me.Keys | Where-Object { $mg.Contains($_) -and $me[$_] -ne $mg[$_] })
$same = @($me.Keys | Where-Object { $mg.Contains($_) -and $me[$_] -eq $mg[$_] }).Count

Write-Host ("发码清单比较: 期望 " + $exp.Count + " 行 / 实得 " + $got.Count +
            " 行 / 相同 " + $same + " / 哈希不同 " + $diff.Count +
            " / 缺席 " + $missing.Count + " / 多出 " + $extra.Count)
foreach ($k in ($diff | Select-Object -First 10)) {
    Write-Host ("  哈希不同: " + $k)
    Write-Host ("    期望 " + ([string]$me[$k]).Substring(0, [Math]::Min(78, ([string]$me[$k]).Length)))
    Write-Host ("    实得 " + ([string]$mg[$k]).Substring(0, [Math]::Min(78, ([string]$mg[$k]).Length)))
}
foreach ($k in (@($missing + $extra) | Select-Object -First 8)) { Write-Host ("  行集合变了: " + $k) }

if ($diff.Count -eq 0 -and $missing.Count -eq 0 -and $extra.Count -eq 0 -and $exp.Count -eq $got.Count) {
    Write-Host "emit-manifest 与期望逐行相同"
    exit 0
}

# ---- 红了就把「该怎么办」直接打在报错里 ----
# 为什么要把手册写进这段代码而不只写在文档里：这道门撞红的多半不是本线的作者，而是往语料里加了夹具、
# 或改了一处发码的协作者（门 #434 = 新夹具没登记；head 15e119f3 = 整型提升改了 7 份发码）。
# 他们看到的只有「Emit manifest (shape oracle) 挂了」这一句，而登记这件事本机要装 MSVC、要跑二十分钟 ——
# 不写清一条命令，门红就从「保护」变成「卡住别人」。三种成因分开说，因为**处置完全不同**。
$handbook = @()
if ($extra.Count -gt 0) {
    $handbook += ('新增 ' + $extra.Count + ' 份输入没有登记哈希（**不是发码回归**，新夹具本来就没有旧期望可比）：' +
                  (($extra | Select-Object -First 6) -join ', '))
}
if ($diff.Count -gt 0) {
    $handbook += ('存量 ' + $diff.Count + ' 份发码真的变了（哈希不同）：逐行看过差异，确认是你这次改动想要的形状再登记')
}
if ($missing.Count -gt 0) {
    $handbook += ('登记里有、这次清单里没有的 ' + $missing.Count + ' 份：' + (($missing | Select-Object -First 6) -join ', ') +
                  ' —— 先问「是夹具被删/改名，还是语料枚举口径变了」，**别用登记把它抹平**（那等于悄悄缩小覆盖面）')
}
$handbook += ('登记（不用本机装 MSVC：吃 CI 那台二进制自己交回来的数）： `pwsh scripts/rebless_emit_manifest.ps1 -Bless`' +
              ' ；先不带 -Bless 跑一遍就是只看不写。CI 那份清单也可自己取： `git fetch <remote> ci/emit-manifest`')
foreach ($line in $handbook) { Write-Host ("  处置: " + $line) }
if ($env:GITHUB_ACTIONS -eq 'true') {
    # 门读数进 check-run 的注解区（台账 §B92 同一条通道）：别人不必翻整篇日志也能看见这三行。
    foreach ($line in $handbook) {
        $flat = ($line -replace '\r?\n', ' ') -replace '::', ':'
        Write-Host ('::error::shape-oracle: ' + $flat)
    }
}
Write-Host "emit-manifest 与期望不一致（发码形状变了，或语料清单要重新登记）"
exit 1
