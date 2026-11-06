#!/usr/bin/env powershell
# check_form_icon.ps1 - 结构哨兵: 窗体默认图标只有一位主人 (账 #305 第303刀)
#
# 背景 (实测): 参考那台 VB6 exe 里 RT_GROUP_ICON(14) id=1 + RT_ICON(3) ids 30001..30003
# 就是 VB6 的默认窗体图标 (16 份语料 exe 逐字节同); 没有 Icon 属性的窗体, VB6 运行时把它
# WM_SETICON(ICON_BIG) 交给窗口而**类图标仍为 0** (GCLP_HICON=0, WM_GETICON(ICON_BIG)!=0)。
# C3 此前两半都没有: 产物里 0 枚图标资源, 窗口也没人设图标 -> 标题栏画的是 shell 兜底图。
#
# 现在的形状 (两处, 各自只有一位主人):
#   driver: src/driver/default_form_icon.hpp 存字节 (group 1 + 3 枚 image),
#           driver_link.cpp 的 embedDefaultFormIcon 在**链接后**用 UpdateResourceW 注入;
#           工程自己带 RT_GROUP_ICON 或产物是 DLL 就不注入。
#   RTL:    vb6forms.c 的 vb6_formApplyDefaultIcon 在窗体创建时 LoadIconW + WM_SETICON。
# 本哨兵钉: 两处的 group id 必须同值; 注入 4 条资源; 设图标只有一个站点;
# 资源归属在 BAS 入口 (hInstance == NULL) 时落回 GetModuleHandleW(NULL)。
# 与 check_di_stubs.ps1 同一类静态用例, 由 run_tests.ps1 的 [STATIC] 项调用。
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot)
)
$ErrorActionPreference = "Continue"
$viol = @()

function Read-Rel([string]$rel) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path $p)) { return $null }
    return (Get-Content -Raw -LiteralPath $p)
}

$hdrRel = "src/driver/default_form_icon.hpp"
$drvRel = "src/driver/driver_link.cpp"
$invRel = "src/driver/res_inventory.cpp"
$invHdr = "src/driver/res_inventory.hpp"
$rtlRel = "src/rtl/core/vb6forms/vb6forms.c"

$hdr = Read-Rel $hdrRel
$drv = Read-Rel $drvRel
$inv = Read-Rel $invRel
$invH = Read-Rel $invHdr
$rtl = Read-Rel $rtlRel
foreach ($pair in @(@($hdrRel,$hdr), @($drvRel,$drv), @($invRel,$inv), @($invHdr,$invH), @($rtlRel,$rtl))) {
    if ($null -eq $pair[1]) { $viol += ("MISSING " + $pair[0]) }
}

# --- 1) driver 那张字节表: group 1 + 三枚 image (ids / 字节数固定) ---
$hdrGroupId = -1
$rtlGroupId = -1
if ($hdr) {
    $m = [regex]::Match($hdr, "kDefaultFormIconGroupId\s*=\s*(\d+)\s*;")
    if ($m.Success) { $hdrGroupId = [int]$m.Groups[1].Value }
    else { $viol += "R1 default_form_icon.hpp: 找不到 kDefaultFormIconGroupId" }
    foreach ($pat in @("kDefaultFormIconGroup\[48\]",
                       "kDefaultFormIconImage30001\[304\]",
                       "kDefaultFormIconImage30002\[744\]",
                       "kDefaultFormIconImage30003\[296\]",
                       "kDefaultFormIconImages\[3\]")) {
        if (-not [regex]::IsMatch($hdr, $pat)) { $viol += ("R1 default_form_icon.hpp: 缺 " + $pat) }
    }
    # 表里三枚 id 与数组名必须成对出现 (抄新一组字节却漏改 id = 注入成孤儿资源)
    foreach ($id in @(30001, 30002, 30003)) {
        if (-not [regex]::IsMatch($hdr, ("\{\s*" + $id + "\s*,\s*\d+\s*,\s*kDefaultFormIconImage" + $id + "\s*\}"))) {
            $viol += ("R1 default_form_icon.hpp: 表项 " + $id + " 不成对")
        }
    }
}

# --- 2) RTL: 同一个 id + 创建窗体时设一次 + 资源归属落回 exe ---
if ($rtl) {
    $m = [regex]::Match($rtl, "#define\s+VB6_DEFAULT_FORM_ICON_ID\s+(\d+)")
    if ($m.Success) { $rtlGroupId = [int]$m.Groups[1].Value }
    else { $viol += "R2 vb6forms.c: 找不到 VB6_DEFAULT_FORM_ICON_ID" }
    if (-not [regex]::IsMatch($rtl, "static void vb6_formApplyDefaultIcon\(HWND hwnd, void\* hInstance\)")) {
        $viol += "R2 vb6forms.c: 缺 vb6_formApplyDefaultIcon 定义"
    }
    # BAS 的 Sub Main 入口不调 vb6_SetAppInstance => hInstance 是 NULL; 不许直接把它交给 LoadIconW
    if (-not [regex]::IsMatch($rtl, "if \(!hi\) hi = GetModuleHandleW\(NULL\);")) {
        $viol += "R2 vb6forms.c: 没有 hInstance==NULL 时落回 GetModuleHandleW(NULL) (LoadIconW(NULL, 1) 恒失败)"
    }
    $nSet = ([regex]::Matches($rtl, "SendMessageW\(hwnd, WM_SETICON, ICON_BIG")).Count
    if ($nSet -ne 1) { $viol += ("R2 vb6forms.c: WM_SETICON 站点 = " + $nSet + " (want 1)") }
    $nCall = ([regex]::Matches($rtl, "vb6_formApplyDefaultIcon\(hwnd, hInstance\);")).Count
    if ($nCall -ne 1) { $viol += ("R2 vb6forms.c: 调用点 = " + $nCall + " (want 1: vb6_CreateFormWindowB)") }
}
if ($hdrGroupId -ge 0 -and $rtlGroupId -ge 0 -and $hdrGroupId -ne $rtlGroupId) {
    $viol += ("R2 group id 两份权威不同值: driver=" + $hdrGroupId + " rtl=" + $rtlGroupId)
}
if ($hdrGroupId -ne 1) { $viol += ("R2 group id = " + $hdrGroupId + " (want 1: VB6 默认窗体图标)") }

# --- 3) 链接后注入: 两处闸门 + 写资源只有两处站点 + 失败必回滚 ---
if ($drv) {
    # 第328刀把「注哪一份字节」与「怎么写资源」分开了：一条注入腿（两发 UpdateResourceW）+
    # 两条来路（默认表 / 工程 .frx）。定义与调用各只许一处，否则第二位主人又长出来。
    $nDef = ([regex]::Matches($drv, "static bool embedFormIcon\(const std::string& exePath, const std::string& iconFrmUtf8,")).Count
    if ($nDef -ne 1) { $viol += ("R3 driver_link.cpp: embedFormIcon 定义 = " + $nDef + " (want 1)") }
    $nUse = ([regex]::Matches($drv, "embedFormIcon\(msvcOpts\.outputFile, iconFrmUtf8, options\.verbose\);")).Count
    if ($nUse -ne 1) { $viol += ("R3 driver_link.cpp: embedFormIcon 调用 = " + $nUse + " (want 1)") }
    foreach ($pair in @("static void defaultFormIconParts", "static bool embedFormIconImages", "static bool embedFormIconGroup")) {
        $n = ([regex]::Matches($drv, [regex]::Escape($pair))).Count
        if ($n -ne 1) { $viol += ("R3 driver_link.cpp: " + $pair + " = " + $n + " (want 1)") }
    }
    # 两条 UpdateResourceW: RT_ICON(3) 那一发 + RT_GROUP_ICON(14) 那一发
    # (Begin/End 是同族名字, 这里按词界排除, 免得把开合两句算进去)
    $nUpd = ([regex]::Matches($drv, "(?<!Begin)(?<!End)UpdateResourceW\(")).Count
    if ($nUpd -ne 2) { $viol += ("R3 driver_link.cpp: UpdateResourceW = " + $nUpd + " (want 2: RT_ICON + RT_GROUP_ICON)") }
    if (-not [regex]::IsMatch($drv, "if \(hasFormModule && !userHasIcon\)")) {
        $viol += "R3 driver_link.cpp: 注入闸门 (hasFormModule && !userHasIcon) 不见了"
    }
    if (-not [regex]::IsMatch($drv, "!options\.isDll")) {
        $viol += "R3 driver_link.cpp: DLL 不该注入默认窗体图标 (闸门 !options.isDll 不见了)"
    }
    if (-not [regex]::IsMatch($drv, "EndUpdateResourceW\(hUpd, ok \? FALSE : TRUE\)")) {
        $viol += "R3 driver_link.cpp: 注入失败必须 EndUpdateResourceW(discard) 回滚"
    }
}

# --- 4) 「工程自带图标」那一读只有一个产地 ---
if ($inv) {
    if (-not [regex]::IsMatch($inv, "kRtGroupIcon\s*=\s*14;")) { $viol += "R4 res_inventory.cpp: 缺 kRtGroupIcon = 14" }
    if (-not [regex]::IsMatch($inv, "typeOrd == kRtGroupIcon\) hasIconGroup = true;")) {
        $viol += "R4 res_inventory.cpp: hasIconGroup 的产地不见了"
    }
}
if ($invH -and -not [regex]::IsMatch($invH, "bool\s+hasIconGroup")) {
    $viol += "R4 res_inventory.hpp: 缺 hasIconGroup 字段"
}

# --- 5) 第328刀: 「工程自己带图标」那一问只许有一位主人 ---
$pfhRel = "src/driver/project_form_icon.hpp"
$pfcRel = "src/driver/project_form_icon.cpp"
$cmRel = "CMakeLists.txt"
$pfh = Read-Rel $pfhRel
$pfc = Read-Rel $pfcRel
$cm = Read-Rel $cmRel
foreach ($pair in @(@($pfhRel,$pfh), @($pfcRel,$pfc), @($cmRel,$cm))) {
    if ($null -eq $pair[1]) { $viol += ("MISSING " + $pair[0]) }
}
if ($pfc) {
    foreach ($sig in @("bool loadFormIconFromFrm(", "bool buildFormIconResources(", "bool isControlBeginLine(")) {
        $n = ([regex]::Matches($pfc, [regex]::Escape($sig))).Count
        if ($n -ne 1) { $viol += ("R5 project_form_icon.cpp: " + $sig + " = " + $n + " (want 1)") }
    }
    # 帧号起始值只许问 default_form_icon.hpp 那一位权威（在 .cpp 里出现一次 = 没有第二处手抄 30001）
    $nBase = ([regex]::Matches($pfc, "kFormIconImageIdBase")).Count
    if ($nBase -ne 1) { $viol += ("R5 project_form_icon.cpp: kFormIconImageIdBase = " + $nBase + " (want 1)") }
    # 停块规则必须真被扫描循环调用（前缀匹配会把 `BeginProperty` 当控件块 ⇒ 窗体自己的 Icon 永远读不到）
    $nStop = ([regex]::Matches($pfc, "isControlBeginLine\(line, s\)")).Count
    if ($nStop -ne 1) { $viol += ("R5 project_form_icon.cpp: 控件块停块那一站 = " + $nStop + " (want 1)") }
}
if ($hdr) {
    $nId = ([regex]::Matches($hdr, "kFormIconImageIdBase\s*=\s*30001;")).Count
    if ($nId -ne 1) { $viol += "R5 default_form_icon.hpp: 帧号起始值那一位权威不见了" }
}
if ($drv) {
    $nLoad = ([regex]::Matches($drv, "loadFormIconFromFrm\(iconFrmUtf8, ico, why\)")).Count
    if ($nLoad -ne 1) { $viol += ("R5 driver_link.cpp: 取图标那一站 = " + $nLoad + " (want 1)") }
    $nBuild = ([regex]::Matches($drv, "buildFormIconResources\(ico, images, ids, group\)")).Count
    if ($nBuild -ne 1) { $viol += ("R5 driver_link.cpp: 拍资源那一站 = " + $nBuild + " (want 1)") }
    $nRead = ([regex]::Matches($drv, "!iconFormName_\.empty\(\) \? iconFormName_ : startupObject_")).Count
    if ($nRead -ne 1) { $viol += ("R5 driver_link.cpp: IconForm 那一问的读者 = " + $nRead + " (want 1)") }
}
$nSet = ([regex]::Matches((Read-Rel "src/driver/driver_compile.cpp"), "iconFormName_ = project\.iconForm;")).Count
if ($nSet -ne 1) { $viol += ("R5 driver_compile.cpp: iconFormName_ 的产地 = " + $nSet + " (want 1: 别在第二处再读 .vbp)") }
if ($cm) {
    $nSrc = ([regex]::Matches($cm, "src/driver/project_form_icon\.cpp")).Count
    if ($nSrc -ne 1) { $viol += ("R5 CMakeLists.txt: project_form_icon.cpp 入册 = " + $nSrc + " (want 1: 不入册就静默不参与构建)") }
}

if ($viol.Count -gt 0) {
    Write-Host ("FAIL form_icon ({0} 处)" -f $viol.Count)
    $viol | ForEach-Object { Write-Host ("  " + $_) }
    exit 1
}
Write-Host "OK form_icon"
exit 0
