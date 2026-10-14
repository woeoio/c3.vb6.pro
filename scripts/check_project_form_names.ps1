# check_project_form_names.ps1 - census: 工程级窗体名的名单只有一个建造点（账 #278 §B110）
#
# 起因（2026-10-09，语料 tests/c29timer/TmForm.frm:131 的 `Unload TmForm2`）：VB6 里**窗体名**在
# 裸名位就是它的默认实例。发码侧一直是这么答的 —— `cgen_expr_ident_symbol.inc` 拿
# `knownFormModuleNames_` 查，命中就发 `vb6_form_hwnd_<名>()`。但语义层那份"工程级已有的名字"
# 表（Public 过程 / Public Const / 模块名）里没有这一格，而模块名那份**只在限定符位**认
# （`TmForm2.Visible` 不报、`Unload TmForm2` 报）。后果不是"只是噪声"：
#   * Option Explicit 的模块：每条合法写法多配一条 VB3001；
#   * 宽松模块：裸名落进"隐式 Variant 局部"那一格 —— 实测产物从
#     `vb6_UnloadForm(vb6_form_hwnd_PfSibling())` 变成
#     `vb6_UnloadForm(vb6_VariantToObjectVal(PfSibling))` + 一行 `vb6_VARIANT PfSibling = … 隐式变量`，
#     也就是**卸载一枚空的 Variant 而不是那枚窗体**（与 §B108 那条"零产物差异 ≠ 无害"同一形状）。
# 修法 = 第四份工程级名单，**与发码侧同源**：两边都走 Driver::collectFormModuleNames()。
#
# 本哨兵钉两头，缺一头都会红：
#   V1 正例放行：Option Explicit 的窗体里裸写兄弟窗体名 ⇒ 不许报 VB3001。
#   V2 反面证人：同一枚窗体里放一枚真不存在的名 (`pfNoSuchNameAnywhere`) ⇒ 必须**仍然**报 VB3001。
#   V3 宽松窗体名那一格走真出口：产物里 `vb6_form_hwnd_PfSibling(` 至少 2 处（一份 Option Explicit
#      的 .frm + 一份宽松的 .bas 各一次），且不许出现隐式局部 `vb6_VARIANT PfSibling`、
#      也不许出现 `vb6_VariantToObjectVal(PfSibling` 那层包装。
#   S1 只有一个建造点：`collectFormModuleNames` 恰好 1 处定义 + 1 处声明（driver.hpp），语义那一侧
#      的文件里同串恰好 2 次（定义 + 建表调用），发码那一侧恰好 1 次调用；
#      旧那份就地扫描 (`formModuleNames.insert`) 不许回来；语义层消费点 `projFormNames_.count` 恰好 1；
#      setter 1 声明 + 1 调用。
#
# 口径里那条**自我排除**（窗体模块里指着自己是名字不算放行）不是随手加的：发码侧那条支路写的就是
# `lower != knownFormName_`，两层必须同一个答案 —— 放行它等于把一条发码侧不认的写法判成合法。
#
# 输出 ASCII（控制台是 GBK，中文读数重定向后不可 grep）。文件必须 UTF-8 BOM + CRLF：
# PS 5.1 读无 BOM 的 .ps1 按 ANSI，行尾中文字节会吃掉换行 ⇒ ParserError 而退出码仍 0。
#
# 用法: powershell -File scripts\check_project_form_names.ps1   (PASS = exit 0)

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
$work = Join-Path $env:TEMP ("c3_formnames_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $work -Force | Out-Null

# 只认**非注释部分**里的登记点：逐行先切掉 `//` 之后再匹配（实测 LastIndexOf("`n") 反推行首
# 会把上一整行当成前缀算进来，把好行误判成注释）。
function Collect-Matches([string]$text, [string]$pattern) {
    $found = @()
    foreach ($line in ($text -split "`r?`n")) {
        $cut = $line.IndexOf('//')
        $body = if ($cut -ge 0) { $line.Substring(0, $cut) } else { $line }
        foreach ($m in [regex]::Matches($body, $pattern)) { $found += $m }
    }
    return $found
}

try {
    function Frm($nm, $body) {
        @(
            'VERSION 5.00',
            "Begin VB.Form $nm ",
            "   Caption         =   `"$nm`"",
            '   ClientHeight    =   3000',
            '   ClientWidth     =   4500',
            '   ScaleHeight     =   3000',
            '   ScaleWidth      =   4500',
            'End',
            "Attribute VB_Name = `"$nm`"",
            'Option Explicit',
            '',
            $body
        )
    }
    $main = Frm "PfMain" @(
        'Private Sub Form_Load()',
        '    Unload PfSibling',
        '    Dim bogus As Long',
        '    bogus = pfNoSuchNameAnywhere',
        'End Sub',
        ''
    )
    $sib = Frm "PfSibling" @(
        'Private Sub Form_Load()',
        '    Debug.Print "sibling"',
        'End Sub',
        ''
    )
    # 宽松标准模块（**没有** Option Explicit）：同一枚窗体名按裸名用
    $loose = @(
        'Attribute VB_Name = "PfLoose"',
        '',
        'Public Sub PfGo()',
        '    Unload PfSibling',
        'End Sub',
        ''
    )
    $vbp = @(
        'Type=Exe',
        'Form=PfMain.frm',
        'Form=PfSibling.frm',
        'Module=PfLoose; PfLoose.bas',
        'ExeName32="pf.exe"',
        'Startup="PfMain"',
        ''
    )
    Set-Content -LiteralPath (Join-Path $work "PfMain.frm")    -Value $main   -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $work "PfSibling.frm") -Value $sib    -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $work "PfLoose.bas")   -Value $loose  -Encoding ASCII
    $proj = Join-Path $work "pf.vbp"
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

    # ---------- V1: 裸写的兄弟窗体名不许报 ----------
    if ($errTxt -match "VB3001[^\r\n]*PfSibling") {
        $bad += "V1 bare project form name PfSibling still reports VB3001 (VB6: a form name is its default instance)"
    }
    # ---------- V2: 真不存在的那个名字必须继续报 ----------
    if ($errTxt -notmatch "VB3001[^\r\n]*pfNoSuchNameAnywhere") {
        $bad += "V2 NEGATIVE CONTROL FAILED: pfNoSuchNameAnywhere no longer reports VB3001 - the gate would be blanket-off"
    }
    # ---------- V3: 宽松那一档必须走真出口, 不许留隐式局部 ----------
    $nHwnd = @([regex]::Matches($emit, 'vb6_form_hwnd_PfSibling\(')).Count
    if ($nHwnd -lt 2) {
        $bad += ("V3 emitted C has only " + $nHwnd + " vb6_form_hwnd_PfSibling() sites (want >= 2: the .frm and the loose .bas)")
    }
    if ($emit -match 'vb6_VARIANT\s+PfSibling\b') {
        $bad += "V3 emitted C still declares an implicit Variant local named PfSibling"
    }
    if ($emit -match 'vb6_VariantToObjectVal\(\s*PfSibling\s*\)') {
        $bad += "V3 the unload still goes through vb6_VariantToObjectVal(PfSibling) - that unloads an empty Variant, not the form"
    }

    # ---------- S1: 名单只有一个建造点 ----------
    $sem  = Join-Path $root "src\driver\driver_semantics.cpp"
    $scan = Join-Path $root "src\driver\detail\driver_codegen_typedfield_scan.inc"
    $hdr  = Join-Path $root "src\driver\driver.hpp"
    $util = Join-Path $root "src\semantics\semantic_analyzer_util.cpp"
    $anH  = Join-Path $root "src\semantics\semantic_analyzer.hpp"
    foreach ($f in @($sem, $scan, $hdr, $util, $anH)) {
        if (-not (Test-Path -LiteralPath $f)) { $bad += ("S missing file " + $f) }
    }
    if ((Test-Path -LiteralPath $sem) -and (Test-Path -LiteralPath $scan) -and
        (Test-Path -LiteralPath $hdr) -and (Test-Path -LiteralPath $util) -and
        (Test-Path -LiteralPath $anH)) {
        $def   = @(Collect-Matches ([System.IO.File]::ReadAllText($sem))  'std::unordered_set<std::string> Driver::collectFormModuleNames').Count
        $decl  = @(Collect-Matches ([System.IO.File]::ReadAllText($hdr))  'std::unordered_set<std::string> collectFormModuleNames').Count
        $useSem  = @(Collect-Matches ([System.IO.File]::ReadAllText($sem))  'collectFormModuleNames\(\)').Count
        $useCgen = @(Collect-Matches ([System.IO.File]::ReadAllText($scan)) 'collectFormModuleNames\(\)').Count
        $oldScan = @(Collect-Matches ([System.IO.File]::ReadAllText($scan)) 'formModuleNames\.insert').Count
        $cons  = @(Collect-Matches ([System.IO.File]::ReadAllText($util)) 'projFormNames_\.count').Count
        $setD  = @(Collect-Matches ([System.IO.File]::ReadAllText($anH)) 'void setProjectFormNames\(').Count
        $setC  = @(Collect-Matches ([System.IO.File]::ReadAllText($sem)) 'analyzer->setProjectFormNames\(').Count
        if ($def -ne 1)  { $bad += ("S1 collectFormModuleNames defined " + $def + " times (want exactly 1)") }
        if ($decl -ne 1) { $bad += ("S1 collectFormModuleNames declared " + $decl + " times in driver.hpp (want 1)") }
        if ($useSem -ne 2 -or $useCgen -ne 1) {
            $bad += ("S1 consumers: semantics=" + $useSem + " (want 2 = the definition + the semantics call site) codegen=" +
                     $useCgen + " (want 1) - the list must have ONE builder for both layers")
        }
        if ($oldScan -ne 0) {
            $bad += ("S1 the old inline scan is back in the codegen pass (formModuleNames.insert = " + $oldScan +
                     ", want 0) - that is a second builder of the same fact")
        }
        if ($cons -ne 1) { $bad += ("S1 semantics asks projFormNames_ " + $cons + " times (want exactly 1 consumer)") }
        if ($setD -ne 1 -or $setC -ne 1) {
            $bad += ("S1 setProjectFormNames decl=" + $setD + " call=" + $setC + " (want 1 + 1)")
        }
    }
} finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
    exit 1
}
Write-Host "PASS project_form_names (V1 bare form name silent, V2 bogus name still reported, V2 2 real exits + no implicit local, S1 one builder / one consumer)"
exit 0
