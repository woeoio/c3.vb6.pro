# check_extender_push_pop.ps1 - Extender 的每个**全局变量**都必须在 push/pop/snapshot 里存取 (账 #294)
#
# 起因（实测 ComCtlsDemo）: `vb6_uc_push`/`vb6_uc_pop` 过去只存取
# `vb6_Extender_Left/Top/Width/Height` 四个槽位，其余 11 个 `vb6_Extender_*` 全是裸全局,
# 出栈没人还原 ⇒ **控件之间互相污染**。实测 ToolBar 在 ReadProperties 里
# `Extender.Align = vbAlignTop`（ToolBar.ctl:879），之后 StatusBar 读到的就是 1，
# 而不是它 .frm 里的 Align=2 (AlignBottom)，于是走错分支 —— 只因为 StatusBar 的
# 设计期高度恰好等于自然高度，肉眼看不出来。静默、跨控件、GA 抓不到。
#
# 钉住两件事：
#   1. 成员集合 —— 从 RTL 声明处**自动列出**所有 vb6_Extender_* 全局变量（不写死清单），
#      所以以后新增成员同样受管；
#   2. 存取对称 —— 每个成员都要在 vb6_uc_push 的保存段、vb6_uc_pop 的还原段、
#      以及 vb6_uc_snapshot 的"没有当前记录"兜底段里各出现一次。
#
# 注意只捞**变量**不捞函数: vb6_Extender_SetFocus / _DragIcon / _Container / _ZOrder
# 都是 RTL 函数，本来就不该进出栈。
#
# 用法: powershell -File scripts\check_extender_push_pop.ps1   (PASS = exit 0)

param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = "Stop"
$bad = @()

$names = New-Object System.Collections.Generic.HashSet[string]
$rxVar = '(?m)^\s*(?:extern\s+)?[A-Za-z_]\w*\s*\*?\s*(vb6_Extender_[A-Za-z_]\w*)\s*(?:=[^;]*)?;'
foreach ($df in @((Join-Path $Root "src\rtl\core\vb6rtl\vb6rtl_com.c"),
                 (Join-Path $Root "src\rtl\core\vb6rtl\vb6rtl_userctl.h"))) {
    if (-not (Test-Path -LiteralPath $df)) { $bad += "missing $df"; continue }
    $txt = [System.IO.File]::ReadAllText($df)
    foreach ($m in [regex]::Matches($txt, $rxVar)) { [void]$names.Add($m.Groups[1].Value) }
}
if ($names.Count -lt 10) { $bad += ("census found only " + $names.Count + " Extender globals - census is broken") }

$hostFile = Join-Path $Root "src\rtl\core\vb6forms\uc\uc_host.c"
if (-not (Test-Path -LiteralPath $hostFile)) { $bad += "missing $hostFile" }
else {
    $src = [System.IO.File]::ReadAllText($hostFile)
    function Get-Body([string]$text, [string]$sig) {
        $i = $text.IndexOf($sig); if ($i -lt 0) { return "" }
        $j = $text.IndexOf("`n}", $i); if ($j -lt 0) { return $text.Substring($i) }
        return $text.Substring($i, $j - $i)
    }
    $push = Get-Body $src "void vb6_uc_push("
    $pop  = Get-Body $src "void vb6_uc_pop("
    $snap = Get-Body $src "static void vb6_uc_snapshot("
    if (-not $push) { $bad += "cannot locate vb6_uc_push" }
    if (-not $pop)  { $bad += "cannot locate vb6_uc_pop" }
    if (-not $snap) { $bad += "cannot locate vb6_uc_snapshot" }
    if ($push -and $pop -and $snap) {
        $extNames = @($names | Sort-Object)
        # --- Fix 296 / 3b: Extender 那 15 枚全局**不许**再出现在 push/pop/snapshot ---
        # 生成码侧全部走实例通道(vb6_UC_ExtenderField -> 当前实例记录 vb6_UCRec.ext*),
        # 生成码里一处都不再读/写这些全局, 所以这套存取是纯开销, 而且它正是「控件之间互相
        # 污染」的来源。**这一条是新不变式本身** —— 原来那条「每枚都必须存取」正好相反。
        foreach ($n in $extNames) {
            if ($push -match ('(?<![\w])' + [regex]::Escape($n) + '(?![\w])')) {
                $bad += "$n still saved in vb6_uc_push (Extender state must not round-trip through globals)"
            }
            if ($pop -match ('(?<![\w])' + [regex]::Escape($n) + '(?![\w])')) {
                $bad += "$n still restored in vb6_uc_pop (Extender state must not round-trip through globals)"
            }
            if ($snap -match ('(?<![\w])' + [regex]::Escape($n) + '(?![\w])')) {
                $bad += "$n still copied by vb6_uc_snapshot (the struct fields are gone; that write was never consumed)"
            }
        }
        # 负控: 把一行存取塞回去必须转红(否则这条检查形同虚设)
        $negPush = $push + "`r`n    saved->extTag = (void*)vb6_Extender_Tag;"
        if ($negPush -notmatch '(?<![\w])vb6_Extender_Tag(?![\w])') {
            $bad += "sentinel is vacuous: reintroducing an Extender save in push is not detected"
        }
        # --- 另一族: Scale / Font / hDC 等 vb6_UserControl_* 全局**仍在**存取 ---
        # 它们与 Extender 无关(实测 FontCombo.ctl DrawString -> hDC、VBFlexGrid.ctl
        # NewFont -> Font 都要读), 3b 一根没碰。这里钉住, 免得"顺手清理"扩大化。
        $keepSaves = @('vb6_UserControl_ScaleWidth','vb6_UserControl_ScaleHeight',
                       'vb6_UserControl_ScaleMode','vb6_UserControl_hDC',
                       'vb6_UserControl_ContainerHwnd','vb6_UserControl_Enabled',
                       'vb6_UserControl_Font','vb6_Ambient_Font')
        foreach ($n in $keepSaves) {
            if ($push -notmatch ('(?<![\w])' + [regex]::Escape($n) + '(?![\w])')) {
                $bad += "$n no longer saved in vb6_uc_push (3b must not touch the Scale/Font/hDC family)"
            }
            if ($pop -notmatch ('(?<![\w])' + [regex]::Escape($n) + '(?![\w])')) {
                $bad += "$n no longer restored in vb6_uc_pop (3b must not touch the Scale/Font/hDC family)"
            }
        }
        # --- 对称性: 剩下那一族全局, push 存了哪些 pop 就必须还原哪些, 反之亦然 ---
        # 这条是 3b 真正该有的护栏。踩过的坑: pop 里
        #   `vb6_UserControl_Width = saved->extWidth;`
        # 借的是 **Extender** 的存档位, push 那边存的却是 `saved->extWidth =
        # vb6_Extender_Width`(另一枚)。3b 把 Extender 那套拆掉后, 这个借位断了 ——
        # pop 读到已删除的字段, 整个 RTL 编译不过, 82 道 compile 里 10 道 UserControl/
        # form 类全红。逐个名字查是查不出来的(两边拼写都不一样), 只能按**集合**比。
        $fam = 'vb6_(?:UserControl|Ambient)_\w+'
        $pushSaves = @([regex]::Matches($push, ('saved->\w+\s*=\s*\(?[^;]*?\b(' + $fam + ')\b')) | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $popRestores = @([regex]::Matches($pop, ('\b(' + $fam + ')\b\s*=\s*\(?[^;]*?saved->\w+')) | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        foreach ($g in $pushSaves) {
            if ($popRestores -notcontains $g) { $bad += "$g saved in push but never restored in pop (asymmetric -> leaks into the next control)" }
        }
        foreach ($g in $popRestores) {
            if ($pushSaves -notcontains $g) { $bad += "$g restored in pop but never saved in push (asymmetric, or borrowed a field that no longer exists)" }
        }
        # 负控: 只在 push 侧加一枚, 必须转红
        $neg = $pushSaves + 'vb6_UserControl_NegOnly'
        if (($neg | Sort-Object -Unique) -notcontains 'vb6_UserControl_NegOnly') {
            $bad += "sentinel is vacuous: a push-only global is not detected as asymmetric"
        }
        # --- 不许有悬空字段引用: saved->/s-> 上那 15 个字段已经删了 ---
        foreach ($m in [regex]::Matches(($push + "`r`n" + $pop + "`r`n" + $snap), '\b(?:saved|s)\s*->\s*(ext[A-Z]\w*)')) {
            $bad += "dangling field reference '$($m.Groups[1].Value)' on a vb6_UCSaved (Fix 296 / 3b deleted all 15 ext* fields)"
        }
    }
    # --- vb6_UCSaved 不得再声明那 15 个 ext* 字段(Clean Slate) ---
    # ⚠ 别把 `struct vb6_UserControl_Extender_Type ext;`(Fix 133u 的 Extender.Visible/Height
    #   结构体成员)一起当成被删的那些 —— 它与 15 枚全局无关, 必须留着。
    $ihFile = Join-Path $Root "src\rtl\core\vb6forms\vb6forms_uc_internal.h"
    if (-not (Test-Path -LiteralPath $ihFile)) { $bad += "missing $ihFile" }
    else {
        $ih = [System.IO.File]::ReadAllText($ihFile)
        $i0 = $ih.IndexOf('typedef struct vb6_UCSaved {')
        $i1 = $ih.IndexOf('} vb6_UCSaved;')
        if ($i0 -lt 0 -or $i1 -lt 0) { $bad += "cannot locate the vb6_UCSaved struct" }
        else {
            $savedBody = $ih.Substring($i0, $i1 - $i0)
            # 注释里会**故意**引用被删掉的字段名(解释为什么删), 所以先剥注释再扫字段
            $savedCode = $savedBody
            $savedCode = [regex]::Replace($savedCode, '(?s)/\*.*?\*/', ' ')
            $savedCode = [regex]::Replace($savedCode, '(?m)//.*$', ' ')
            foreach ($f2 in @('extLeft','extTop','extWidth','extHeight','extAlign','extDragMode',
                              'extHelpContextID','extWhatsThisHelpID','extVisible','extVb6Default',
                              'extCancel','extTag','extToolTipText','extContainer','extDragIcon')) {
                if ($savedCode -match ('\b' + $f2 + '\b')) {
                    $bad += "vb6_UCSaved still declares $f2 (Clean Slate removed all 15 ext* fields)"
                }
            }
            if ($savedCode -notmatch 'vb6_UserControl_Extender_Type\s+ext\b') {
                $bad += "vb6_UCSaved lost its 'ext' member of type vb6_UserControl_Extender_Type (Fix 133u Extender.Visible/Height -- NOT one of the 15)"
            }
        }
    }
}

# 同一个普查, 换一个消费点: vb6_UC_ExtenderField 的**全局兜底**那一段。
# 它必须把普查里每一枚都兜住 —— 兜不住的那几枚, 在**没有当前记录**时
# vb6_UC_SetExtenderMember 撞 `if (!fp) return;`, 写**静默丢**(改之前那枚全局落地,
# 改之后没了)。这条也是 3b 的前置: push/pop 一旦不再镜像全局, 兜底就是它们唯一的退路。
$ctlFile = Join-Path $Root "src\rtl\core\vb6forms\uc\uc_controls.c"
if (-not (Test-Path -LiteralPath $ctlFile)) { $bad += "missing $ctlFile" }
else {
    $ctls = [System.IO.File]::ReadAllText($ctlFile)
    # 必须点名**定义** (结尾带 `{`); 用 `void* vb6_UC_ExtenderField(` 会先撞上
    # getter 里那三行 extern 声明 (uc_controls.c:356/366/374), 取到空body,
    # 于是 15 枚**全都**报"没有兜底" —— 假红。
    $sigFb = 'void* vb6_UC_ExtenderField(void* p, const wchar_t* name, int32_t* kind) {'
    $i = $ctls.IndexOf($sigFb)
    if ($i -lt 0) {
        $bad += "cannot locate the vb6_UC_ExtenderField definition (signature must include the full parameter list and the opening brace)"
    } else {
        $j = $ctls.IndexOf("`n}", $i); if ($j -lt 0) { $j = $ctls.Length }
        $efBody = $ctls.Substring($i, $j - $i)
        # 记录段用 L"Width" 之类, 兜底段用 vb6_Extender_Width —— 正好分得开
        foreach ($n in ($names | Sort-Object)) {
            if ($efBody -notmatch ('\b' + [regex]::Escape($n) + '\b')) {
                $bad += "$n has no global fallback in vb6_UC_ExtenderField (write silently dropped when no record is current)"
            }
        }
        # 负控: 抠掉一枚兜底行必须转红。改名必须换成**不含原名**的一串, 否则子串匹配
        # 照样判过, 负控就恒红而什么也没钉。
        $negFb = $efBody -replace [regex]::Escape('vb6_Extender_Left'), 'REMOVED_LEFT'
        $stillCovers = $true
        foreach ($n in ($names | Sort-Object)) {
            if ($negFb -notmatch ('\b' + [regex]::Escape($n) + '\b')) { $stillCovers = $false; break }
        }
        if ($stillCovers) { $bad += "sentinel is vacuous: deleting a fallback row still passes" }
    }
}
if ($bad.Count -eq 0) {
    Write-Host ("PASS (" + $names.Count + " Extender globals absent from push/pop/snapshot (Fix 296 / 3b new invariant); Scale/Font/hDC family still saved+restored; vb6_UCSaved carries no ext* field)")
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
exit 1