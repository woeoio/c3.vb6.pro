# check_format_date_arg_arm.ps1 - Format 首参那一臂的 Date 档只许问唯一权威 (账 #430 / §B204)
#
# 这一臂原来**手抄**了一张"VB 类型 → 装箱出口"的名单, 其中 Date 那一行写的是
#   vb6_VariantDouble((double)(...))
# 而唯一权威 boxToVariant (src/backend/cgen_util_ctrl.cpp:2827) 对同一问的答案是
#   vb6_VariantDate(...)
# 两位主人两个答案, 且权威那条从这一站**从没被走到**。跟 #428/#431 同一条教训:
# 表的格子就是口径, 抄表的人不会知道自己抄错了。
#
# 实测后果(判据 tests/b432 的 B432-18): 一枚 Dim d As Date 局部发成 double, 交进
# vb6_Format 永远贴 VT_R8, 可同一枚 VarType(d) 答 7 —— "这枚值真是日期"在运行期
# 根本读不出来。第二格是**故意钉下的错答案**: BASE 答 True/False, 改后答 True/True。
#
# R1  Format 首参链里的 Date 档走 boxToVariant 权威(源码里那一臂必须调它)
# R2  这一臂**不许**再出现 vb6_VariantDouble((double) 那句手抄的 Date 装箱
# R3  权威 boxToVariant 自己确有 Date 档且答 vb6_VariantDate
# R4  **只动 Date 一臂**: Long/Integer/Double/Single/String/Byte 那几档必须原样还在
#     (它们与权威也各不同, 但都是已登记的既存形状; 且数值臂要留给真数值
#      Format(43831,"mmm") 那一族。别顺手塌成一句 boxToVariant —— 那会一次移六档)
# R5  判据两头都在: tests/b432 的 B432-18 两格都钉在
#
# 用法: powershell -File scripts\check_format_date_arg_arm.ps1   (PASS = exit 0)

param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = "Stop"
$bad = @()

$inc = Join-Path $Root "src\backend\detail\expr\cgen_expr_call_conv_output.inc"
$cpp = Join-Path $Root "src\backend\cgen_util_ctrl.cpp"
$fix = Join-Path $Root "tests\b432\B432Main.bas"
foreach ($f in @($inc, $cpp, $fix)) { if (-not (Test-Path -LiteralPath $f)) { $bad += "missing $f" } }

if ($bad.Count -eq 0) {
    $src = [System.IO.File]::ReadAllText($inc)
    $co  = [System.IO.File]::ReadAllText($cpp)
    $fx  = [System.IO.File]::ReadAllText($fix)

    # 只看 Format 首参那一段链, 别把文件别处也算进来
    $i = $src.IndexOf('if (callee == "vb6_Format"')
    $chain = ""
    if ($i -ge 0) {
        $j = $src.IndexOf("argList.clear();", $i)
        if ($j -lt 0) { $j = $src.Length }
        $chain = $src.Substring($i, $j - $i)
    } else { $bad += "cannot locate the Format first-arg chain in cgen_expr_call_conv_output.inc" }

    if ($chain) {
        # ⚠ **先剥注释再扫**: 本刀在那一臂上方留了一段解释"原来写的是什么"的注释, 里面
        #   **字面**引着 vb6_VariantDouble 那句手抄装箱 —— 不剥注释的话 R2 会把自己的
        #   说明当成违规(踩过一次)。哨兵扫的是代码, 不是文档。
        $chainCode = [regex]::Replace($chain, '(?s)/\*.*?\*/', ' ')
        $chainCode = [regex]::Replace($chainCode, '(?m)//.*$', ' ')

        # ---- R1: Date 档那一臂必须调 boxToVariant ----
        $dateArm = [regex]::Match($chainCode, '(?s)argType\s*==\s*Vb6Type::Date\s*\)\s*\{(.*?)\}')
        if (-not $dateArm.Success) {
            $bad += "R1: no `argType == Vb6Type::Date` arm found in the Format first-arg chain"
        } elseif ($dateArm.Groups[1].Value -notmatch 'boxToVariant\s*\(') {
            $bad += "R1: the Format first-arg Date arm does NOT call boxToVariant -- it still has its own hand-copied boxing list (#430)"
        }
        # ---- R2: 那句手抄的 Date 装箱不许还在 ----
        if ($dateArm.Success -and $dateArm.Groups[1].Value -match 'vb6_VariantDouble') {
            $bad += "R2: the Format first-arg Date arm still emits vb6_VariantDouble -- a Dim-As-Date local would keep boxing as VT_R8"
        }
        # ---- R4: 旁证 —— 其余档位必须原样(防止有人把整条链塌成一句 boxToVariant) ----
        foreach ($keep in @('vb6_VariantLong', 'vb6_VariantDouble', 'vb6_VariantString', 'vb6_VariantInt')) {
            if ($chainCode -notmatch [regex]::Escape($keep)) {
                $bad += ("R4: " + $keep + " is gone from the Format first-arg chain -- #430 only covers the Date arm; collapsing the whole chain moves six shapes at once")
            }
        }
        # 负控: 把旧写法塞回去必须被 R1/R2 抓到
        $neg = $chainCode -replace '(?s)(argType\s*==\s*Vb6Type::Date\s*\)\s*\{)[^}]*\}', '$1 args[0] = "vb6_VariantDouble((double)(" + args[0] + "))"; }'
        if ($neg -notmatch 'vb6_VariantDouble') {
            $bad += "negative control is vacuous: reinstating the old Date arm is not detected"
        } else {
            $na = [regex]::Match($neg, '(?s)argType\s*==\s*Vb6Type::Date\s*\)\s*\{(.*?)\}')
            if ($na.Success -and $na.Groups[1].Value -match 'boxToVariant\s*\(') {
                $bad += "negative control is vacuous: the reinstated Date arm still passes the boxToVariant test"
            }
            if (-not ($na.Groups[1].Value -match 'vb6_VariantDouble')) {
                $bad += "negative control is vacuous: reinstating the old Date arm does not trip the vb6_VariantDouble ban"
            }
        }
    }

    # ---- R3: 权威自己确有 Date 档 ----
    # boxToVariant 的真身是 `CCodeGen::boxToVariant`(台账里记的 cgen_util_ctrl.cpp:2828
    # 只是那枚 return 所在的行号, 不是函数头 —— 按函数头定位才不会跟着行号漂)。
    $box = [regex]::Match($co, '(?s)CCodeGen::boxToVariant\s*\(.*?\n\}')
    if (-not $box.Success) {
        $bad += "R3: cannot locate CCodeGen::boxToVariant in cgen_util_ctrl.cpp"
    } else {
        $boxCode = [regex]::Replace($box.Value, '(?s)/\*.*?\*/', ' ')
        $boxCode = [regex]::Replace($boxCode, '(?m)//.*$', ' ')
        if ($boxCode -notmatch 'Vb6Type::Date\s*:\s*\r?\n\s*return\s+"vb6_VariantDate\(') {
            $bad += "R3: boxToVariant's Date arm does not answer vb6_VariantDate(...) -- the sole authority and this station must agree"
        }
    }

    # ---- R5: 判据两头都在 ----
    if ($fx -notmatch 'B432-18') { $bad += "R5: tests/b432 no longer pins B432-18 (the #430 witness)" }
    if ($fx -notmatch 'Format\(dt\)') { $bad += "R5: tests/b432 no longer exercises Format(<Dim As Date local>)" }
    if ($fx -notmatch 'Format\(v\)')   { $bad += "R5: tests/b432 no longer exercises Format(<Variant slot>)" }
}

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
    exit 1
}
Write-Host "PASS (Format first-arg Date arm asks the sole authority boxToVariant -> vb6_VariantDate; the other arms untouched; b432 B432-18 still pinned)"
exit 0