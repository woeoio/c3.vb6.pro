# check_builtin_global_objects.ps1 - census: 内置全局对象的裸名两层必须答同一份名单（账 #278 §B108）
#
# 起因（2026-10-09，探针 `.build/b112_probe` 实测）：`src/semantics/builtin/builtin_funcs.inc`
# 的 `addObj` 名单是"哪些裸名是 VB6 内置全局对象"在语义层的唯一登记处，而发码侧
# (`src/backend/detail/expr/cgen_expr_ident_dispatch.inc` 的内置对象段) 各自硬编码拦名。
# 两份名单曾经不一致：`Printers` 只在发码侧有 (`vb6_Printers_Collection()`)、语义侧漏登记 ⇒
#   Option Explicit 的模块：每条合法引用多配一条 VB3001；
#   宽松模块：更实 —— 裸名落进"隐式 Variant 局部"那格，产物里多出
#     `vb6_VARIANT Printers = vb6_VariantEmpty();` 且 ForEach 的集合被包一层
#     `vb6_VariantToObjectVal(...)` —— 名字被一枚空 Variant 挤掉。
#
# 本哨兵钉三层，缺一层都会红：
#   V1 语义放行：同一份探针工程里裸 `Printers`（一份 Option Explicit、一份宽松）都不许报 VB3001。
#   V2 产物形状：`vb6_Printers_Collection()` 必须出现两次（两条模块各一次），且**不许**出现
#      隐式局部 `vb6_VARIANT Printers`、也**不许**出现 `vb6_VariantToObjectVal(vb6_Printers_Collection`。
#   V3 反面证人：同一工程里放一枚真不存在的名 (`bgoNoSuchNameAnywhere`) ⇒ 必须**仍然**报 VB3001
#      （只钉 V1 的哨兵会替"干脆什么都不报"背书 —— 那是同一个洞的另一种坏法）。
#   S1 跨层契约（这条才是本 bug 的结构抓法）：发码侧内置对象段拦的每一个裸名都必须在
#      `addObj` 名单里。段 = 文件里第一个 `vbokonly`（内置常量段）之前的部分；两个刻意豁免的
#      名字列在 $kExempt 里并写明理由（它们不是"内置全局对象"那一格）。两处采集都**跳过注释行**
#      （负控实测：不跳的话把 `addObj("Printers")` 注释掉哨兵照样绿 —— 那种哨兵不算护栏）。
#   S2 名单不许换家：`addObj` 登记点恰好 9 处且就是那 9 枚名字（加一格要写在这里，别散到别处）。
#
# 输出 ASCII（控制台是 GBK，中文读数重定向后不可 grep）。文件必须 UTF-8 BOM + CRLF：
# PS 5.1 读无 BOM 的 .ps1 按 ANSI，行尾中文字节会吃掉换行 ⇒ ParserError 而退出码仍 0。
#
# 用法: powershell -File scripts\check_builtin_global_objects.ps1   (PASS = exit 0)

param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Exe = ""      # 负控用: 拿另一枚 C3.exe 跑同一份夹具
)
$ErrorActionPreference = "Stop"
$root = $Root
$exe = if ($Exe) { $Exe } else { Join-Path $root ".build\C3.exe" }
if (-not (Test-Path -LiteralPath $exe)) {
    Write-Host ("FAIL missing " + $exe + "  (先跑 scripts/build.bat 产出 C3.exe)")
    exit 1
}

$bad = @()

# 只认**非注释行**里的登记点。这条不是洁癖：负控实测过 —— 纯文本正则会连
# `// addObj("Printers")` 一起算进去，那种哨兵对"有人把这行注释掉了"完全无感（红不了）。
function Get-Registrations([string]$text, [string]$pattern) {
    $names = @()
    foreach ($m in [regex]::Matches($text, $pattern)) {
        $ls = $text.LastIndexOf("`n", $m.Index)
        if ($ls -lt 0) { $ls = 0 } else { $ls = $ls + 1 }
        if ($text.Substring($ls, $m.Index - $ls).Contains('//')) { continue }
        $names += $m.Groups[1].Value.ToLower()
    }
    return $names
}
$work = Join-Path $env:TEMP ("c3_builtinglobal_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $work -Force | Out-Null

try {
    # ---------- 夹具：一份 Option Explicit 的标准模块 + 一份宽松的标准模块，都按裸名用 Printers ----------
    $modExp = @(
        'Attribute VB_Name = "BgoExplicit"',
        'Option Explicit',
        '',
        'Public Function EpCount() As Long',
        '    Dim q As Variant',
        '    EpCount = 0',
        '    For Each q In Printers',
        '        EpCount = EpCount + 1',
        '    Next',
        'End Function',
        '',
        'Public Function EpBogus() As Long',
        '    Dim w As Long',
        '    w = bgoNoSuchNameAnywhere',
        '    EpBogus = w',
        'End Function',
        ''
    )
    $modLoose = @(
        'Attribute VB_Name = "BgoLoose"',
        '',
        'Public Function LpCount() As Long',
        '    Dim r As Variant',
        '    LpCount = 0',
        '    For Each r In Printers',
        '        LpCount = LpCount + 1',
        '    Next',
        'End Function',
        ''
    )
    $vbp = @(
        'Type=Exe',
        'Module=BgoExplicit; BgoExplicit.bas',
        'Module=BgoLoose; BgoLoose.bas',
        'ExeName32="bgo.exe"',
        'Startup="BgoExplicit"',
        ''
    )
    Set-Content -LiteralPath (Join-Path $work "BgoExplicit.bas") -Value $modExp -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $work "BgoLoose.bas")    -Value $modLoose -Encoding ASCII
    $proj = Join-Path $work "bgo.vbp"
    Set-Content -LiteralPath $proj -Value $vbp -Encoding ASCII

    $outC = Join-Path $work "emit.c"
    $errFile = Join-Path $work "emit.err"
    # 用 OS 级重定向取字节（--emit-c 的产物走 stdout, 诊断走 stderr）
    $p = Start-Process -FilePath $exe -ArgumentList @("`"$proj`"", "--emit-c") `
                       -NoNewWindow -Wait -PassThru `
                       -RedirectStandardOutput $outC -RedirectStandardError $errFile
    if ($p.ExitCode -ne 0) { $bad += ("V0 compiler exited " + $p.ExitCode + " on the probe project") }

    $emit = ""
    if (Test-Path -LiteralPath $outC) { $emit = Get-Content -LiteralPath $outC -Raw }
    $errTxt = ""
    if (Test-Path -LiteralPath $errFile) { $errTxt = Get-Content -LiteralPath $errFile -Raw }

    # ---------- V1: 裸 Printers 不许再报 VB3001 ----------
    if ($errTxt -match "VB3001[^\r\n]*Printers") {
        $bad += "V1 bare Printers still reports VB3001 (Printers is a VB6 builtin global object)"
    }

    # ---------- V2: 产物两头钉 —— 走真出口, 且没被隐式局部挤掉 ----------
    $nColl = ([regex]::Matches($emit, 'vb6_Printers_Collection\(\)')).Count
    if ($nColl -ne 2) {
        $bad += ("V2 emitted C has " + $nColl + " vb6_Printers_Collection() sites (want exactly 2: one per module)")
    }
    if ($emit -match 'vb6_VARIANT\s+Printers\b') {
        $bad += "V2 emitted C still declares an implicit Variant local named Printers (the object got pushed out)"
    }
    if ($emit -match 'vb6_VariantToObjectVal\(\s*vb6_Printers_Collection') {
        $bad += "V2 emitted C still wraps the collection in vb6_VariantToObjectVal (relic of the implicit local)"
    }

    # ---------- V3: 真不存在的那个名字必须继续报 ----------
    if ($errTxt -notmatch "VB3001[^\r\n]*bgoNoSuchNameAnywhere") {
        $bad += "V3 NEGATIVE CONTROL FAILED: bgoNoSuchNameAnywhere no longer reports VB3001 - the gate would be blanket-off"
    }

    # ---------- S1: 跨层契约 —— 发码侧拦的裸名都得在语义侧的名单里 ----------
    $semFile = Join-Path $root "src\semantics\builtin\builtin_funcs.inc"
    $cgenFile = Join-Path $root "src\backend\detail\expr\cgen_expr_ident_dispatch.inc"
    foreach ($f in @($semFile, $cgenFile)) {
        if (-not (Test-Path -LiteralPath $f)) { $bad += ("S missing file " + $f) }
    }
    if ((Test-Path -LiteralPath $semFile) -and (Test-Path -LiteralPath $cgenFile)) {
        $sem = Get-Content -LiteralPath $semFile -Raw
        $cgen = Get-Content -LiteralPath $cgenFile -Raw

        $addObjNames = @(Get-Registrations $sem 'addObj\("([A-Za-z][A-Za-z0-9]*)"')
        if (@($addObjNames | Sort-Object -Unique).Count -ne $addObjNames.Count) {
            $bad += ("S2 addObj registers a name twice (" + ($addObjNames -join ",") + ") - one row per object")
        }

        # 内置对象段 = 第一个 vbokonly（内置常量段）之前
        $cut = $cgen.IndexOf('vbokonly')
        if ($cut -lt 0) { $bad += "S1 cannot find the intrinsic-constant anchor (vbokonly) in the dispatch file" }
        $seg = $cgen.Substring(0, [Math]::Max(0, $cut))
        $cgenNames = @(Get-Registrations $seg 'if \(lower == "([a-z][a-z0-9]*)"')
        $cgenNames = @($cgenNames | Sort-Object -Unique)

        # 两个刻意豁免: 不是"内置全局对象"那一格，各有自己的出口
        $kExempt = @('me', 'parent')
        $missing = @($cgenNames | Where-Object { ($kExempt -notcontains $_) -and ($addObjNames -notcontains $_) })
        if ($missing.Count -gt 0) {
            $bad += ("S1 codegen intercepts bare name(s) " + ($missing -join ",") +
                     " that the semantics list addObj does not register -> VB3001 noise + implicit local")
        }

        # ---------- S2: 名单只在 addObj 一处，且恰好这 9 枚 ----------
        $wantAddObj = @('app', 'clipboard', 'console', 'debug', 'err', 'forms', 'printer', 'printers', 'screen')
        if ($addObjNames.Count -ne $wantAddObj.Count) {
            $bad += ("S2 addObj registers " + $addObjNames.Count + " names (" + ($addObjNames -join ",") +
                     ") - want exactly " + $wantAddObj.Count + "; a new builtin object must be listed here too")
        } else {
            $diff = @($wantAddObj | Where-Object { $addObjNames -notcontains $_ })
            $extra = @($addObjNames | Where-Object { $wantAddObj -notcontains $_ })
            if ($diff.Count -gt 0 -or $extra.Count -gt 0) {
                $bad += ("S2 addObj set drift: missing=" + ($diff -join ",") + " unexpected=" + ($extra -join ","))
            }
        }
        $callSites = $addObjNames.Count
        if ($callSites -ne $wantAddObj.Count) {
            $bad += ("S2 addObj registration sites = " + $callSites + " (want exactly " + $wantAddObj.Count + ")")
        }
    }
} finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
    exit 1
}
Write-Host "PASS builtin_global_objects (V1 no VB3001, V2 2 exits + no implicit local, V3 bogus name still reported, S1 codegen names covered, S2 9 addObj sites)"
exit 0
