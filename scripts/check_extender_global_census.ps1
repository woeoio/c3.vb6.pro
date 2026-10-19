# check_extender_global_census.ps1 - 发码侧的 Extender 全局普查 (Fix 296 收口)
#
# 3b 把 push/pop 里那 15 枚 vb6_Extender_* 全局的存取整个拆掉了, 理由是"生成码里一处都
# 不再读/写它们"。这个脚本把那句话**变成可执行的判据**, 免得它以后悄悄变回假话。
#
# 它查的不是源码, 是**实发 C**。源码里写 `Extender.` 的地方很少, 但真正要证明的是
# cgen 之后那 465 份发码里没有一枚裸全局。所以:
#   1. 先扫语料源码里带 `Extender.` 的文件, 反查出它们各自的**宿主输入**(宿主才进清单);
#   2. 只对这几个宿主跑 --emit-c(比整表 465 份快得多, 而且够精确);
#   3. 在实发 C 上普查那 15 枚全局。
#
# **阳性对照是这个脚本的关键**。普查报 0 很容易, 空转也报 0 —— 如果 cgen 哪天不再发
# Extender 相关代码, 普查会"通过"但什么也没证明。所以额外断言:全仓发码里通道调用
# (vb6_UC_GetExtenderMember*/SetExtenderMember) 的总数 **必须 > 0**。
#
# 顺带报一下**第二套机制**(只报不断言): `UserControl.Extender.X` 走的是
# `vb6_UserControl_Extender.X`(一个全局**结构体**), 不是那 15 枚全局。它发码真的在读
# (czUI.ctl 7 处), 所以 3b 那套"直接删掉"对它不适用; 而且它同样由 push/pop 存取
# (uc_host.c:225 存 / :258 回灌 / :290 还原 / :351 快照) —— 形状和那 15 枚一模一样。
# 这是**已知的第二条向量**, 还没治。哪天要治, 走 4a/3b 的路(接进实例通道), 别删。
#
# 用法: powershell -File scripts\check_extender_global_census.ps1   (PASS = exit 0)
#       powershell -File scripts\check_extender_global_census.ps1 -KeepEmitted

param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [switch]$KeepEmitted
)
$ErrorActionPreference = "Stop"
$bad = @()

# ---------------------------------------------------------------- 1. 那 15 枚名字
# 跟 check_extender_push_pop.ps1 用**同一个**正则从 RTL 声明里推, 不写死清单, 两边不会漂。
# 注意 `vb6_UserControl_Extender` 不匹配 `vb6_Extender_` 前缀 —— 正好把两套机制分开。
$names = New-Object System.Collections.Generic.HashSet[string]
$rxVar = '(?m)^\s*(?:extern\s+)?[A-Za-z_]\w*\s*\*?\s*(vb6_Extender_[A-Za-z_]\w*)\s*(?:=[^;]*)?;'
foreach ($df in @((Join-Path $Root "src\rtl\core\vb6rtl\vb6rtl_com.c"),
                 (Join-Path $Root "src\rtl\core\vb6rtl\vb6rtl_userctl.h"))) {
    if (-not (Test-Path -LiteralPath $df)) { $bad += "missing $df"; continue }
    $txt = [System.IO.File]::ReadAllText($df)
    foreach ($m in [regex]::Matches($txt, $rxVar)) { [void]$names.Add($m.Groups[1].Value) }
}
if ($names.Count -lt 10) { $bad += ("census found only " + $names.Count + " Extender globals - the name census is broken") }
if ($names.Contains("vb6_UserControl_Extender")) {
    $bad += "the name census swallowed vb6_UserControl_Extender - that is the OTHER mechanism and must not be in this list"
}

# ---------------------------------------------------------------- 2. 哪些宿主
$c3 = Join-Path $Root ".build\C3.exe"
if (-not (Test-Path -LiteralPath $c3)) {
    Write-Host ("MISSING " + $c3 + " - build first (the RTL is compiled per test, so rebuild after touching src/rtl)")
    exit 2
}
$testsRoot = Join-Path $Root "tests"
if (-not (Test-Path -LiteralPath $testsRoot)) { Write-Host ("MISSING " + $testsRoot); exit 2 }

$touched = @(Get-ChildItem -Path $testsRoot -Recurse -Include *.frm, *.ctl, *.bas -File |
             Select-String -Pattern "Extender\." -ErrorAction SilentlyContinue |
             ForEach-Object { $_.Path } | Sort-Object -Unique)

# .bas 自己就是清单输入; .frm/.ctl 要反查哪个 .vbp 把它收进去(Control= / UserControl=)
$vbps = @(Get-ChildItem -Path $testsRoot -Recurse -Include *.vbp -File)
$vbpText = @{}
foreach ($v in $vbps) { $vbpText[$v.FullName] = [System.IO.File]::ReadAllText($v.FullName) }

$hosts = New-Object System.Collections.Generic.HashSet[string]
$loose  = @()
foreach ($f in $touched) {
    if ($f -match '\.bas$') { [void]$hosts.Add($f); continue }
    $leaf = [System.IO.Path]::GetFileName($f)
    $owner = @($vbps | Where-Object { $vbpText[$_.FullName] -match ([regex]::Escape($leaf)) })
    if ($owner.Count -eq 0) { $loose += $leaf; continue }   # 散装夹具, 不进任何宿主
    foreach ($o in $owner) { [void]$hosts.Add($o.FullName) }
}

if ($hosts.Count -eq 0) {
    Write-Host ("check_extender_global_census: no Extender-touching host inputs found (touched=" + $touched.Count + ") - census is broken")
    exit 1
}

# ---------------------------------------------------------------- 3. 发码 + 普查
$rxGlobal = '(?<![\w])(' + (($names | Sort-Object | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')(?![\w])'

# 双向防空转。正阳性对照保证"通道还在发"; 这里再保证"全局探测本身是活的" ——
# 若 $rxGlobal 哪天写坏了(转义、字符类、分组), 普查会一路报 0 然后**假装通过**。
# 所以拿一份含全部 15 枚名字的合成文本走一遍同一条正则, 必须一枚不少地抓到。
$selfProbe = ($names | Sort-Object | ForEach-Object { "x = " + $_ + " ;" }) -join "`r`n"
$probeHit = @([regex]::Matches($selfProbe, $rxGlobal) | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
if ($probeHit.Count -ne $names.Count) {
    $bad += ("the global-detection regex is broken: a synthetic line containing all " + $names.Count +
             " names matched only " + $probeHit.Count + " - the census would pass vacuously")
}

$rxChan   = 'vb6_UC_(?:GetExtenderMemberNum|GetExtenderMemberStr|GetExtenderMemberObj|SetExtenderMember|ExtenderField)'
$rxPseudo = 'vb6_UserControl_Extender\.'

$tmpDir = Join-Path $Root ("output\_tmp\extcensus_" + [System.Diagnostics.Process]::GetCurrentProcess().Id)
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

# C3 建会话目录要靠 TEMP/TMP, 指向仓库内的 output\_tmp —— 全仓约定, 不设就是
# "cannot create session directory" 然后 --emit-c 悄悄吐 0 字节。别指望调用方设。
$c3Temp = Join-Path $tmpDir "c3temp"
New-Item -ItemType Directory -Force -Path $c3Temp | Out-Null
$env:TEMP = $c3Temp
$env:TMP  = $c3Temp

$totalChan = 0
$totalPseudo = 0
Write-Host ("Emitting " + $hosts.Count + " Extender-touching host input(s) out of " + $vbps.Count + " .vbp...")
foreach ($h in ($hosts | Sort-Object)) {
    $rel = $h.Substring($Root.Length + 1)
    $out = Join-Path $tmpDir (($rel -replace '[\\/]', '_') + ".c")
    $err = Join-Path $tmpDir (($rel -replace '[\\/]', '_') + ".err")
    $p = Start-Process -FilePath $c3 -ArgumentList @('"' + $h + '"', "--emit-c") -WorkingDirectory $Root `
        -RedirectStandardOutput $out -RedirectStandardError $err -NoNewWindow -Wait -PassThru
    if (-not (Test-Path -LiteralPath $out)) {
        $why = ""
        if (Test-Path -LiteralPath $err) { $why = ((Get-Content -LiteralPath $err | Select-Object -Last 3) -join " / ") }
        $bad += ("$rel : emit produced no output (rc=" + $p.ExitCode + ")" + $(if ($why) { " : " + $why } else { "" }))
        continue
    }

    $t = [System.IO.File]::ReadAllText($out)
    # 发码 0 字节 = C3 根本没跑成(常见: TEMP/TMP 没指到 output\_tmp)。绝不能当成
    # "这个工程没有 Extender" —— 那会让普查假绿, 而下面的阳性对照只会报一句"0 通道"。
    if ($t.Length -eq 0) {
        $why = ""
        if (Test-Path -LiteralPath $err) { $why = ((Get-Content -LiteralPath $err | Select-Object -Last 3) -join " / ") }
        $bad += ("$rel : emitted 0 bytes (rc=" + $p.ExitCode + ")" + $(if ($why) { " : " + $why } else { "" }))
        continue
    }
    $g = [regex]::Matches($t, $rxGlobal)
    $c = [regex]::Matches($t, $rxChan).Count
    $u = [regex]::Matches($t, $rxPseudo).Count
    $totalChan += $c
    $totalPseudo += $u

    if ($g.Count -gt 0) {
        foreach ($grp in ($g | Group-Object { $_.Groups[1].Value } | Sort-Object Name)) {
            $bad += ("$rel : " + $grp.Name + " x" + $grp.Count + " - bare Extender global left in emitted C (must go through the instance channel)")
        }
    }
    $mech = @()
    if ($c -gt 0) { $mech += ("channel x" + $c) }
    if ($u -gt 0) { $mech += ("UserControl_Extender x" + $u) }
    if ($mech.Count -eq 0) { $mech += @("none(this project emits no Extender access)") }
    Write-Host ("  {0,-44} globals={1,-3} {2}" -f $rel, $g.Count, ($mech -join "  "))
}

# ---------------------------------------------------------------- 4. 阳性对照
# 空转的普查也会报 globals=0。所以要求全仓至少有一处真的走了通道, 否则"通过"是假的。
if ($totalChan -le 0) {
    $bad += "positive control failed: 0 channel calls in all emitted C - the census would pass vacuously (cgen may have stopped emitting Extender access entirely)"
}

if ($loose.Count -gt 0) {
    Write-Host ("  note: " + $loose.Count + " Extender-touching file(s) belong to no .vbp (loose fixtures, not compiled hosts): " + ($loose -join ", "))
}
Write-Host ("  channel calls across all emitted C = " + $totalChan + " (positive control)")
Write-Host ("  UserControl_Extender sites = " + $totalPseudo + "  <- the SECOND mechanism, still a global, not fixed by 3b (report only)")

if (-not $KeepEmitted) { Remove-Item -Recurse -Force $tmpDir -ErrorAction SilentlyContinue }
else { Write-Host ("  emitted C kept in " + $tmpDir) }

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
    exit 1
}
Write-Host ("PASS (" + $names.Count + " Extender globals absent from every emitted C; positive control satisfied)")
exit 0