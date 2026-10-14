# check_statusbar_panel_hm.ps1 —— 状态条面板宽那两处答案只许有一套（账 #206 §B41 第一格 / §B129）
#
# 立账的原委：`Panels(i).Width` 在**真 VB6 设计器**手里有两个名字、两种单位 ——
# 设计块里那一行是 `Object.Width`，值是 himetric(0.01mm)；VB 代码里读写的是缇。
# 本仓从前只认无前缀的 `Width`（`frm_parser.cpp:81` 保留点号全名 ⇒ 那一行永远查不到），
# 所以别人工程的面板几何整格丢掉：实测语料 5623 份 .frm 里 356 份 / 1634 处这么写，
# 而本仓三份状态条夹具用的都是自己造的点号形 ⇒ 全语料 A/B 对这一族永远沉默（零覆盖）。
# 这个哨兵钉住修法之后的形状，防的是：
#   ① 那条前缀认读被改掉（发码又看不见设计值了）；
#   ② himetric 与 像素/缇 之间的折算被抄第二份（§B129 那份名单到第二十四刀只剩 1 份文件 / 5 个非
#      注释行，且全在权威本人 vb6forms.c 里 —— 第二十二刀并掉 ax_load.c，第二十四刀并掉 olecon /
#      picture_prop / ax_site_ext 三处；名单**往下减**是本职，再长一份就是 #234「拿 DC 抄两遍」、
#      #235「画笔色两份存储」那个形状）；
#   ③ 弹簧档的 MinWidth 又变成"加项"（夹具实测 e3=493 而 cw=467，多出来正好一枚 MinWidth）。
#   ④ 面板几何那一档的单位回到两份答案（第十八刀把像素存进 e->width，而 VB 代码那一头读写的是
#      缇 ⇒ 同一格两种单位，正是 §B41 剩下的那一格）。第二十二刀之后的口径只有一套：
#      存储=缇、排版时经 vb6_TwipToX 折像素、getter 交排版结果再经 vb6_XToTwipX 折回缇。
#
# 判据面：
#   R1 产物：tests/sbhm/SbHm.vbp 的 --emit-c 里 —— SetPanelWidthHm 恰好 2 处（两枚固定档面板）、
#      实参字面量 `1500` / `2646` 各 1 处、AddPanel 恰好 3、SetPanelAutoSize 恰好 1（弹簧那枚不给
#      Width ⇒ 它必须**不**出现在 Hm 里）、且无前缀那枚 SetPanelWidth( 在产物里恰好 0 处
#      （设计期这一趟只走 Hm 出口；运行期赋值走的是另一张表，不归这条判据管）。
#      第二十二刀起再钉两枚新出口：SetPanelMinWidthHm 恰好 3（三枚面板都写了设计 MinWidth）、
#      GetPanelLeft 恰好 3（HM14 那三处读）—— 缺一个落点就是"设计值整格不见"/"读回空"。
#   R2 源码：src/rtl 里含 2540 的**非注释行**必须逐文件对上名单（第二十四刀起 1 份 / 5 行，见下），
#      多出一份文件即红。数之前先把注释剥干净 —— 从前这里只剥每行的 `//`，**块注释的续行**因此被
#      当代码数（第二十二刀那版名单里的 14 行有两行是 `/* … 2540 … */` 的续行，补上剥法后同一棵
#      树是 12 行；第十七刀定的"剥注释要跨行"这一条对本哨兵同样成立）。vb6forms.c 里那 5 行必须
#      是权威本人：公式档 vb6_HimetricToPxOnDc、浮点对 vb6_HimetricToPxF / vb6_PxToHimetricF、
#      缇档一对 vb6_TwipsToHimetric / vb6_HimetricToTwips（屏 DPI 的两枚 vb6_HimetricToPxX/Y 是
#      转调，本身不含 2540）。状态条那一族只许**转调**权威：本地一次乘除都不许出现（该文件里
#      MulDiv 0 处、HimetricToPxX 0 处 —— 后者是第十八刀那一版"把像素存进缇档"的形状）。
#   R3 源码：vb6forms_statusbar.c 里 `want[i] = e->minWidth` 这一形必须 0 处（下限只答一次），
#      而兜底那一行恰好 1 处、且它折像素问的也是 vb6_TwipToX。
#   R4 防空转：扫到的 RTL 份数 >= 100、产物字节数 >= 4000、两张名单自己的长度也钉死。
#   R5（账 #300 同轮加）：`sbFinishByKey` 三处的两个出口名不许相同，且新出口 `vb6_StatusBar_GetPanelIndex`
#      必须在 header / RTL 定义 / 发码点三处各出现恰好一次（少一处就是 LNK2019 或 C2065 那一族）。
#   R7（账 #206 第二格同轮加）：一枚成员要过四道登记才"真做得动"（memberobj 名字表 / DISPID /
#      派发 case / 发码 recognizer）+ RTL 定义 + header 声明。Panel.Left 这六格各恰好 1 处，
#      设计块那行 MinWidth 的三处同钉 —— **header 那一格是实测漏过的一格**：产物照跑、日志一声不报。
#   R8（账 #303 同轮加）：`fixedTotal += want[i];` 恰好 2 处 —— 固定档与内容档各从剩余空间里扣一次。
#      0 处是本账（弹簧把内容档那格宽吃两次、末格越过客户区），3 处是第十八刀撤掉的"MinWidth 当加项"复发。
#   R9（账 #298 第二十四刀加）：并表之后换算只剩一处权威，所以钉的是**调用点** —— olecon 2+4+4、
#      picture_prop 2（仍按 DC 的 DPI）、ax_site_ext 4（浮点对），外加三条旧形状墓碑各 0 处
#      （picture_prop 的 MulDiv、olecon 的 `* 15L *`、ax_site_ext 的 `const double k`）。
#
# 出口只走 --emit-c（只到发码，不起 cl），所以这条判据是秒级的；负控用 -Root 把整棵树指到副本上跑。

param(
    [string]$Root = ''
)
$ErrorActionPreference = 'Stop'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }

$viol = @()

# R2 名单：himetric / 像素 / 缇 之间的折算在 RTL 里的全部住处（账 #226 同族：一份决定抄几遍）。
# 这份名单是**读数**不是愿望。第二十四刀（账 #298 §B129 收尾）之后 = **1 份文件 / 5 个非注释行**，
# 全在 vb6forms.c：公式本身（`vb6_HimetricToPxOnDc` 一枚 int + `vb6_HimetricToPxF` /
# `vb6_PxToHimetricF` 一对 float）+ 缇档那一对（与 DPI 无关）。以前各写各的三处 —— olecon 的两个
# 方向、picture_prop 按 DC 的 DPI、ax_site_ext 把 DPI 写死成 96 —— 现在一律只转调。
# ⚠ 计数口径本刀起变了：扫描改成**跨行剥块注释**（账 #278 第十七刀在 arity 哨兵上办掉的同一族，
#   这是那份盲区的第二格）。按新口径重算的**改前**读数是 12 行 / 4 份，不是旧名单上的 14 / 5 ——
#   多出来的两条一直是块注释续行被数成代码（olecon 那句"反过来用"、picture_prop 那句英文说明）。
$HmFiles = [ordered]@{
    'src/rtl/core/vb6forms/vb6forms.c'                   = 5
}
$HmWantTotal = 5
$HmAuthority = 'vb6_HimetricToPxOnDc'

# ---------- R1: 产物形状 ----------
$C3 = Join-Path $Root '.build\C3.exe'
if (-not (Test-Path $C3)) {
    Write-Host "FAIL statusbar_panel_hm: C3.exe not found ($C3)"; exit 1
}
$proj = Join-Path $Root 'tests\sbhm\SbHm.vbp'
if (-not (Test-Path $proj)) {
    Write-Host "FAIL statusbar_panel_hm: fixture project missing ($proj)"; exit 1
}
$prev = $ErrorActionPreference
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $C3
$psi.Arguments = '"' + $proj + '" --emit-c'
$psi.UseShellExecute = $false
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$p = [System.Diagnostics.Process]::Start($psi)
$ms = New-Object System.IO.MemoryStream
$p.StandardOutput.BaseStream.CopyTo($ms)
$null = $p.StandardError.ReadToEnd()
$p.WaitForExit()
$ErrorActionPreference = $prev
$bytes = $ms.ToArray()
$emit = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)

function Count-Of([string]$needle) {
    if ($needle -eq '') { return 0 }
    return ([regex]::Matches($emit, [regex]::Escape($needle))).Count
}

$hm = Count-Of 'vb6_StatusBar_SetPanelWidthHm('
if ($hm -ne 2) { $viol += "R1-HM-CALL want exactly 2, got $hm" }
if ((Count-Of ', 1, 1500);') -ne 1) { $viol += 'R1-ARG-1500 want exactly 1 (panel1 Object.Width verbatim)' }
if ((Count-Of ', 2, 2646);') -ne 1) { $viol += 'R1-ARG-2646 want exactly 1 (panel2 Object.Width verbatim)' }
if ((Count-Of 'vb6_StatusBar_AddPanel(') -ne 3) { $viol += 'R1-ADD-PANEL want exactly 3' }
if ((Count-Of 'vb6_StatusBar_SetPanelAutoSize(') -ne 1) { $viol += 'R1-AUTOSIZE want exactly 1 (spring panel only)' }
if ((Count-Of 'vb6_StatusBar_SetPanelWidth((void*)') -ne 0) { $viol += 'R1-NO-PLAIN-SET want 0 (design path goes through the *Hm exit)' }
if ($bytes.Length -lt 4000) { $viol += ('R4-EMIT-BYTES floor 4000, got ' + $bytes.Length) }

# 账 #300 (§B131): Panels(<数字>).Index 与 Panels("键").Index 必须各走各的出口。
# 改前那一格发的是 *ByKey(hwnd, 1) —— 整数进 wchar_t* 槽, 真跑 0xC0000005。
if ((Count-Of 'vb6_StatusBar_GetPanelIndex((void*)') -ne 2) {
    $viol += 'R1-INDEX-BYIDX want exactly 2 (the two numeric subscripts, Panels(2)/Panels(9))'
}
if ((Count-Of 'vb6_StatusBar_GetPanelIndexByKey((void*)') -ne 1) {
    $viol += 'R1-INDEX-BYKEY want exactly 1 (the key subscript only, Panels("fx2"))'
}

# 第二十二刀新增的两枚发码点：设计块那行 MinWidth 走 *Hm 出口（夹具三枚面板都写了它，读数 3），
# 而 Panel.Left 现在有自己的出口（HM14 那三处读，读数 3）。这两条计数是"接线在产物里真落了地"的
# 存在性证人 —— 缺发码点 = 设计值整格不见（第十八刀那一格的形状），缺 recognizer = 读回空。
$hmMin = Count-Of 'vb6_StatusBar_SetPanelMinWidthHm('
if ($hmMin -ne 3) { $viol += ('R1-HM-MINWIDTH want exactly 3 (design block writes MinWidth on all three panels), got ' + $hmMin) }
$plEmit = Count-Of 'vb6_StatusBar_GetPanelLeft('
if ($plEmit -ne 3) { $viol += ('R1-LEFT-EMIT want exactly 3 (the three Panels(i).Left reads in HM14), got ' + $plEmit) }

# ---------- R2 + R3: 源码面 ----------
$rtlRoot = Join-Path $Root 'src\rtl'
$perFile = @{}
$rtlFiles = 0
Get-ChildItem -Path $rtlRoot -Recurse -File -Include *.c, *.h, *.inc | ForEach-Object {
    $rtlFiles++
    $txt = [System.IO.File]::ReadAllText($_.FullName)
    $n = 0
    $inBlock = $false
    foreach ($ln in ($txt -replace "`r`n", "`n").Split("`n")) {
        # 剥注释要**跨行**：块注释的续行（` * 说明 ...`）里出现常数，从前会被数成代码 ——
        # 这正是账 #278 第十七刀在 arity 哨兵上办掉的那一族（"同一个判断抄几份"抄在判据自己身上）。
        $line = $ln
        if ($inBlock) {
            $ix = $line.IndexOf('*/')
            if ($ix -lt 0) { continue }
            $line = $line.Substring($ix + 2)
            $inBlock = $false
        }
        while ($true) {
            $o = $line.IndexOf('/*')
            if ($o -lt 0) { break }
            $c = $line.IndexOf('*/', $o + 2)
            if ($c -lt 0) { $line = $line.Substring(0, $o); $inBlock = $true; break }
            $line = $line.Substring(0, $o) + $line.Substring($c + 2)
        }
        $dc = $line.IndexOf('//')
        if ($dc -ge 0) { $line = $line.Substring(0, $dc) }
        if ($line.IndexOf('2540') -ge 0) { $n++ }
    }
    if ($n -gt 0) {
        $rel = ($_.FullName.Substring($Root.Length).Replace('\', '/').TrimStart('/'))
        $perFile[$rel] = $n
    }
}
foreach ($k in $HmFiles.Keys) {
    $want = $HmFiles[$k]
    $got = 0
    if ($perFile.ContainsKey($k)) { $got = $perFile[$k] }
    if ($got -ne $want) { $viol += "R2-HM-SITES $k want exactly $want, got $got" }
}
foreach ($k in $perFile.Keys) {
    if (-not $HmFiles.Contains($k)) { $viol += "R2-HM-NEW-SITE $k carries an extra himetric conversion (the list is down to 1 file: vb6forms.c)" }
}
$sum = 0
foreach ($v in $perFile.Values) { $sum += $v }
if ($sum -ne $HmWantTotal) { $viol += "R2-HM-TOTAL want $HmWantTotal, got $sum" }

$authPath = Join-Path $rtlRoot 'core\vb6forms\vb6forms.c'
$authTxt = [System.IO.File]::ReadAllText($authPath)
# 公式本人的定义必须逐字住在 vb6forms.c（int 那一枚 + float 那一对）；屏 DPI 与某枚 DC 的 DPI
# 都只是它的调用方 —— 这一条钉的是"公式只有一份"，第二十四刀（账 #298 收尾）之后是硬要求。
$authDef = ([regex]::Matches($authTxt, [regex]::Escape('int vb6_HimetricToPxOnDc(int hm, int dpi) { return MulDiv(hm, dpi, 2540); }'))).Count
if ($authDef -ne 1) { $viol += "R2-AUTHORITY $HmAuthority definition want exactly 1 formula line, got $authDef" }
$pxX = ([regex]::Matches($authTxt, [regex]::Escape('int vb6_HimetricToPxX(int hm) { return vb6_HimetricToPxOnDc(hm, vb6_DpiX()); }'))).Count
$pxY = ([regex]::Matches($authTxt, [regex]::Escape('int vb6_HimetricToPxY(int hm) { return vb6_HimetricToPxOnDc(hm, vb6_DpiY()); }'))).Count
if ($pxX -ne 1) { $viol += ('R2-AUTHORITY the screen-DPI wrapper vb6_HimetricToPxX must ask OnDc exactly once, got ' + $pxX) }
if ($pxY -ne 1) { $viol += ('R2-AUTHORITY the screen-DPI wrapper vb6_HimetricToPxY must ask OnDc exactly once, got ' + $pxY) }
$tfDef = ([regex]::Matches($authTxt, [regex]::Escape('float vb6_HimetricToPxF(float hm, float dpi)'))).Count
$tfInv = ([regex]::Matches($authTxt, [regex]::Escape('float vb6_PxToHimetricF(float px, float dpi)'))).Count
if ($tfDef -ne 1 -or $tfInv -ne 1) {
    $viol += ('R2-AUTHORITY the float pair for OLE container coords must be 1+1, got ' + $tfDef + '+' + $tfInv)
}
# 缇档那一 pair 也必须是权威本人（第二十二刀新增：存储换缇之后，设计期那一档与 getter 那一档
# 都转调它们，本族不再自己乘除）。少了其中一枚 = 某一侧改回本地写法 = 第二份答案。
$twipDef = ([regex]::Matches($authTxt, [regex]::Escape('long vb6_TwipsToHimetric(long tw)'))).Count
$hm2twDef = ([regex]::Matches($authTxt, [regex]::Escape('long vb6_HimetricToTwips(long hm)'))).Count
if ($twipDef -ne 1) { $viol += "R2-AUTHORITY vb6_TwipsToHimetric definition want exactly 1, got $twipDef" }
if ($hm2twDef -ne 1) { $viol += "R2-AUTHORITY vb6_HimetricToTwips definition want exactly 1, got $hm2twDef" }
# axsite/ax_load.c 从名单上退掉的条件就是"它只转调" —— 这里钉住那一格，防它哪天又自己写一遍
$axTxt = [System.IO.File]::ReadAllText((Join-Path $rtlRoot 'core\vb6forms\axsite\ax_load.c'))
$axDeleg = ([regex]::Matches($axTxt, 'vb6_TwipsToHimetric\(')).Count
if ($axDeleg -ne 1) { $viol += ('R2-MERGED axsite/ax_load.c must delegate to the authority exactly once, got ' + $axDeleg) }

$sbPath = Join-Path $rtlRoot 'core\vb6forms\vb6forms_statusbar.c'
$sbTxt = [System.IO.File]::ReadAllText($sbPath)
# 状态条这一族问权威的三道折，各恰好 2 处（固定档 + 弹簧下限 / Width + Left / 两枚 *Hm 设计出口）。
# 注：这些数含注释里出现的次数吗？不含 —— 一律带左括号，写成 `(` 才算一次调用。
$sbTwip = ([regex]::Matches($sbTxt, 'vb6_TwipToX\(')).Count
$sbBack = ([regex]::Matches($sbTxt, 'vb6_XToTwipX\(')).Count
$sbHm = ([regex]::Matches($sbTxt, 'vb6_HimetricToTwips\(')).Count
if ($sbTwip -ne 2) { $viol += ('R2-CALLER statusbar must ask vb6_TwipToX exactly twice (fixed tile + spring floor), got ' + $sbTwip) }
if ($sbBack -ne 2) { $viol += ('R2-CALLER statusbar must ask vb6_XToTwipX exactly twice (Width + Left getter), got ' + $sbBack) }
if ($sbHm -ne 2) { $viol += ('R2-CALLER statusbar must ask vb6_HimetricToTwips exactly twice (the two *Hm design exits), got ' + $sbHm) }
# 本地一次乘除都不许出现；而第十八刀那一版"把像素直接存进缇档"的形状必须归零 —— 这条是它的墓碑。
$sbMul = 0
$sbPx = ([regex]::Matches($sbTxt, 'vb6_HimetricToPxX\(')).Count
foreach ($ln in (($sbTxt -replace "`r`n", "`n").Split("`n"))) {
    $code = ($ln -replace '/\*.*?\*/', '') -replace '//.*$', ''
    $sbMul += ([regex]::Matches($code, 'MulDiv')).Count
}
if ($sbMul -ne 0) { $viol += ('R2-LOCAL-MATH statusbar re-does the conversion ' + $sbMul + ' time(s) instead of asking the authority') }
if ($sbPx -ne 0) { $viol += ('R2-OLD-SHAPE vb6_HimetricToPxX( is called ' + $sbPx + ' time(s) in statusbar - 存储档已换成缇，那一版形状回来了') }
$hdrPath = Join-Path $rtlRoot 'core\vb6forms\vb6forms_window.h'
if (([regex]::Matches(([System.IO.File]::ReadAllText($hdrPath)), 'vb6_HimetricToPxX')).Count -ne 1) {
    $viol += 'R2-DECL header must declare the authority exactly once'
}

$sbLines = ($sbTxt -replace "`r`n", "`n").Split("`n")
$seed = 0
$flo = 0
$floAuth = 0
foreach ($ln in $sbLines) {
    if ($ln -match 'want\[i\]\s*=\s*e->minWidth') { $seed++ }
    if ($ln -match 'take\s*<\s*floorPx') { $flo++ }
    if ($ln.Contains('int floorPx = e->minWidth > 0 ? vb6_TwipToX(e->minWidth) : 0;')) { $floAuth++ }
}
if ($seed -ne 0) { $viol += "R3-SPRING-SEED want[i]=e->minWidth must be 0 (MinWidth is a floor, not an addend), got $seed" }
if ($flo -ne 1) { $viol += "R3-SPRING-FLOOR want exactly 1 floor line (take < floorPx), got $flo" }
if ($floAuth -ne 1) { $viol += ('R3-SPRING-FLOOR-AUTH the floor must fold its 缇 into pixels through vb6_TwipToX, exactly once, got ' + $floAuth) }

# ---------- R5: 数字/键两条出路不许同名（账 #300 的结构性面） ----------
# `sbFinishByKey($byIdx, $byKey)` 的两个参数一旦写成同一个名字，数字下标那一形就会把
# 整数递进 `const wchar_t*` 槽（#300 的真实形状：Index 两格都填 *ByKey ⇒ 启动期 AV，
# 而键下标那一形是对的 ⇒ 症状按"下标写数字还是键"分家，存量针永远看不见）。
$comPath = Join-Path $Root 'src\backend\cgen_util_com.cpp'
$comTxt = [System.IO.File]::ReadAllText($comPath)
$pm = [regex]::Matches($comTxt, 'sbFinishByKey\(\s*"([A-Za-z0-9_]+)"\s*,\s*"([A-Za-z0-9_]+)"')
if ($pm.Count -ne 3) { $viol += ('R5-PAIRS want exactly 3 sbFinishByKey call sites, got ' + $pm.Count) }
foreach ($m in $pm) {
    if ($m.Groups[1].Value -eq $m.Groups[2].Value) {
        $viol += ('R5-SAME-EXIT ' + $m.Groups[1].Value + ' answers both the numeric and the key form')
    }
}
# 新出口三处必须同步（header 声明 / RTL 定义 / 发码点）—— 少一处就是 LNK 或 C2065 那一族
$tri = @(
    (Join-Path $rtlRoot 'core\vb6forms\vb6forms_prop_ctrl.h'),
    (Join-Path $rtlRoot 'core\vb6forms\vb6forms_statusbar.c'),
    $comPath
)
foreach ($f in $tri) {
    $c = ([regex]::Matches([System.IO.File]::ReadAllText($f), 'vb6_StatusBar_GetPanelIndex\b')).Count
    if ($c -ne 1) { $viol += ('R5-TRIPLE ' + (Split-Path -Leaf $f) + ' must carry vb6_StatusBar_GetPanelIndex exactly 1 time, got ' + $c) }
}

# ---------- R6: 账 #299 (§B130) —— Panels 的集合对象只许 memberobj 那一头答 ----------
# 立这一格的原委（全部实测）：改前 `Set po =` / 模块级 `As Object` / `For Each` / `With`
# 四种"把成员当对象用"的写法，头一律发成 vb6_ComGetObjectProp(<HWND>, L"Panels")
# = 拿 HWND 当 IDispatch 用（C29-8b 当年给 TreeView Nodes 写下的同一句症状），
# `For Each` 连循环都没进 —— BASE 那台跑同一份夹具是 fe=0 idx=0 txt=0，
# 第二十刀之后是 fe=3 idx=6 txt=5。
# 第二十一刀把**三条**发射路（值语境 / 对象语境 / 默认成员下标的 binder）收成同一处出口，
# 第二十五刀补上第四条（With 的头）—— 四条现在都问 memberCollectionObjectExpr，所以这里钉的是**读数**不是愿望。
$gaPath = Join-Path $Root 'src\backend\detail\expr\cgen_expr_member_generic_access.inc'
$stPath = Join-Path $Root 'src\backend\detail\util\cgen_state.inc'
$moPath = Join-Path $rtlRoot 'core\vb6forms\vb6forms_memberobj.c'
$hdPath = Join-Path $rtlRoot 'core\vb6forms\vb6forms_prop_ctrl.h'
$gaTxt = [System.IO.File]::ReadAllText($gaPath)
$stTxt = [System.IO.File]::ReadAllText($stPath)
$moTxt = [System.IO.File]::ReadAllText($moPath)
$hdTxt = [System.IO.File]::ReadAllText($hdPath)
# com_bind 那条默认成员下标的发射路（第二十一刀起也问同一处出口）
$cbTxt = [System.IO.File]::ReadAllText((Join-Path $Root 'src\backend\detail\expr\cgen_expr_call_com_bind.inc'))
# With 的头 = 第四条发射路（第二十五刀起也问同一处出口）
$cwTxt = [System.IO.File]::ReadAllText((Join-Path $Root 'src\backend\stmt\cgen_with.cpp'))

# R6-A 产物：四形全部由 memberobj 那枚入口造出来（第二十一刀之后 Set / 模块级 Object /
# With / For Each 走的是同一处出口），而"拿 HWND 当 IDispatch"那一形必须归零 ——
# 改前三形各有 1 处、实参那一路还有 4 处（合计 8 处 stale head），这一条就是它的墓碑。
$coll = Count-Of 'vb6_StatusBar_Panels((void*)vb6_hwnd_StatusBar1)'
$collItem = Count-Of 'vb6_ComCallObject(vb6_StatusBar_Panels('
$collEnum = Count-Of 'vb6_ForEach_Init(vb6_StatusBar_Panels('
$stale = Count-Of 'vb6_ComGetObjectProp(vb6_hwnd_StatusBar1, L"Panels")'
# 第二十五刀的墓碑：With 的头没消费标记时，体内第一条语句会把它当自己的写目标 ——
$leakW = Count-Of 'vb6_ComSetProp(vb6_hwnd_StatusBar1, L"Panels"'
if ($coll -ne 13) { $viol += ('R6-CREATOR-EMIT want exactly 13 (10 走 Item + 1 走枚举 + 2 走 With 头), got ' + $coll) }
if ($leakW -ne 0) { $viol += ('R6-WITH-LEAK-TOMBSTONE ' + $leakW + ' read(s) inside a With block are emitted as a property WRITE again') }
if ($collItem -ne 10) { $viol += ('R6-ITEM-VIA-CREATOR want exactly 10, got ' + $collItem) }
if ($collEnum -ne 1) { $viol += ('R6-FOREACH-VIA-CREATOR want exactly 1, got ' + $collEnum) }
if ($stale -ne 0) {
    $viol += ('R6-OVERLAP ' + $stale + ' head(s) still use the HWND as an IDispatch - ' +
              '成员集合的创建式只许 memberCollectionObjectExpr 一处答')
}

# R6-B 认"这串表达式是集合对象"的三张名单必须同时认这一枚前缀。
# 少一处的后果是**实测过的**：只补拦子不补名单，`Panels.Count = 3` 发成
# `vb6_StatusBar_Panels((void*)vb6_hwnd_SB1).Count` ⇒ cl C2224（void* 上点成员）。
$gaN = ([regex]::Matches($gaTxt, 'find\("vb6_StatusBar_Panels\("\)')).Count
if ($gaN -ne 2) { $viol += ('R6-LIST generic_access must carry the prefix at both sites, got ' + $gaN) }
$stK = ([regex]::Matches($stTxt, '"vb6_StatusBar_Panels\(\(void\*\)vb6_"')).Count
if ($stK -ne 1) { $viol += ('R6-LIST statusBarNameOfExpr kHeads want exactly 1, got ' + $stK) }
$stC = ([regex]::Matches($stTxt, '"vb6_StatusBar_Panels\("')).Count
if ($stC -ne 1) { $viol += ('R6-LIST isControlCollectionExpr want exactly 1, got ' + $stC) }

# R6-C 出口本身只许有一处**定义**，并且三条发射路都问它（值语境 / 对象语境 / 默认成员下标）。
$cgN = ([regex]::Matches($comTxt, 'memberCollectionObjectExpr\(')).Count
$cbN = ([regex]::Matches($cbTxt, 'memberCollectionObjectExpr\(')).Count
$cwN = ([regex]::Matches($cwTxt, 'memberCollectionObjectExpr\(')).Count
$defN = ([regex]::Matches($stTxt, 'std::string memberCollectionObjectExpr\(')).Count
if ($defN -ne 1) { $viol += ('R6-EXIT definition must exist exactly once in cgen_state.inc, got ' + $defN) }
if ($cgN -ne 2) { $viol += ('R6-EXIT-CALLS cgen_util_com must ask the exit twice (value+object), got ' + $cgN) }
if ($cbN -ne 1) { $viol += ('R6-EXIT-CALLS com-bind indexer must ask the exit once, got ' + $cbN) }
if ($cwN -ne 1) { $viol += ('R6-EXIT-CALLS the With header must ask the exit once (fourth road), got ' + $cwN) }

# R6-D 一表到底只许一处写：**创建式字面量**在非注释行里只许出现在 cgen_state.inc。
# 这六枚是仓里"成员集合由谁造"这张表的全部行；再有一处在别的文件里出现，就是第二份答案
# (#234 拿 DC 抄两遍、#235 画笔色两份存储、#229/#88 同一族的复发)。
$creatorLits = @(
    'vb6_TreeView_Nodes((void*)', 'vb6_Toolbar_Buttons((void*)', 'vb6_StatusBar_Panels((void*)',
    'vb6_ImageList_ListImages((void*)', 'vb6_ListView_ListItems((void*)', 'vb6_ListView_ColumnHeaders((void*)')
$creatorWant = @{
    'vb6_TreeView_Nodes((void*)' = 2; 'vb6_Toolbar_Buttons((void*)' = 1
    'vb6_StatusBar_Panels((void*)' = 2; 'vb6_ImageList_ListImages((void*)' = 2
    'vb6_ListView_ListItems((void*)' = 2; 'vb6_ListView_ColumnHeaders((void*)' = 2 }
$creatorFiles = @{}
$creatorTot = @{}
foreach ($L in $creatorLits) { $creatorTot[$L] = 0 }
Get-ChildItem -Path (Join-Path $Root 'src\backend') -Recurse -File -Include *.cpp, *.inc, *.h | ForEach-Object {
    $rel = ($_.FullName.Substring($Root.Length).Replace('\', '/').TrimStart('/'))
    foreach ($ln in ([System.IO.File]::ReadAllText($_.FullName) -replace "`r`n", "`n").Split("`n")) {
        $s = $ln.Trim()
        if ($s.StartsWith('//')) { continue }
        foreach ($L in $creatorLits) {
            $n = ([regex]::Matches($s, [regex]::Escape($L))).Count
            if ($n -gt 0) {
                $creatorTot[$L] += $n
                if (-not $creatorFiles.ContainsKey($L)) { $creatorFiles[$L] = @{} }
                $creatorFiles[$L][$rel] = 1
            }
        }
    }
}
foreach ($L in $creatorLits) {
    $got = $creatorTot[$L]
    $want = $creatorWant[$L]
    if ($got -ne $want) { $viol += ("R6-ONE-MAP $L code-line count want exactly $want, got $got") }
    foreach ($f in $creatorFiles[$L].Keys) {
        if ($f -ne 'src/backend/detail/util/cgen_state.inc') {
            $viol += ("R6-ONE-MAP-SITE $L is also written in $f - the table must live in one place")
        }
    }
}

# R6-E RTL 那一头：入口定义 / 声明 / Item 按 Key 那一档，各恰好一次
if (([regex]::Matches($moTxt, 'vb6_StatusBar_Panels\(void\* hwnd\)')).Count -ne 1) {
    $viol += 'R6-CREATOR vb6forms_memberobj.c must define the collection entry exactly once'
}
if (([regex]::Matches($hdTxt, 'vb6_StatusBar_Panels\(void\* hwnd\)')).Count -ne 1) {
    $viol += 'R6-DECL vb6forms_prop_ctrl.h must declare it exactly once'
}
if (([regex]::Matches($moTxt, 'VB6_MEMCK_PANELS\)\s+idx = vb6_StatusBar_GetPanelIndexByKey')).Count -ne 1) {
    $viol += 'R6-ITEM-BY-KEY memColl Item must answer the PANELS key form (else Panels("k") = Nothing)'
}
# #300 那格留下的第三份答案不许回到面板这一族：inline `p->index` 全仓现存 6 处（别的族），
# 面板那一处已改成问 getter。这个数只许降不许升。
$inline = ([regex]::Matches($moTxt, 'memSetI4\(out, p->index\);')).Count
if ($inline -gt 6) {
    $viol += ('R6-THIRD-ANSWER inline p->index answers grew beyond the 6 that exist: ' + $inline)
}
if (([regex]::Matches($moTxt, 'memSetI4\(out, vb6_StatusBar_GetPanelIndex\(p->owner, p->index\)\);')).Count -ne 1) {
    $viol += 'R6-INDEX-AUTHORITY Panel.Index must ask the one getter exactly once'
}

# ---------- R7: 账 #206 第二格 —— 新登记的 Panel.Left 与 *Hm 设计出口，落点一处不许缺 ----------
# 一枚成员从"名字认得、读数交空"到"真做得动"要过四道登记（memberobj 名字表 / DISPID / 派发
# case / 发码 recognizer），加 RTL 定义与 header 声明两格 —— 缺哪一格的症状都不一样，而
# **实测过的最阴那一格是 header**：第二十二刀发码点与定义都写齐了、声明漏了，产物照跑
# （调用返回 void、实参按 ABI 到位），C3 的构建日志里根本没有 cl 警告这一族 ⇒ 只有数落点能发现。
# 账 #174 就是同一形状（SelectedControls 发成隐式声明）。
$plNames = ([regex]::Matches($moTxt, 'L"Left"')).Count
$plDisp = ([regex]::Matches($moTxt, '#define VB6_MEMD_PANEL_LEFT 9')).Count
$plCase = ([regex]::Matches($moTxt, 'case VB6_MEMD_PANEL_LEFT:')).Count
$plMoCall = ([regex]::Matches($moTxt, 'vb6_StatusBar_GetPanelLeft\(p->owner, p->index\)')).Count
$plDef = ([regex]::Matches($sbTxt, 'int32_t vb6_StatusBar_GetPanelLeft\(void\* hwnd, int32_t index\)')).Count
$plDecl = ([regex]::Matches($hdTxt, 'vb6_StatusBar_GetPanelLeft\(void\* hwnd, int32_t index\);')).Count
$plRecog = ([regex]::Matches($comTxt, 'sbGet\("vb6_StatusBar_GetPanelLeft"\)')).Count
if ($plNames -ne 1) { $viol += ('R7-NAMES kPanelNames must register L"Left" exactly once, got ' + $plNames) }
if ($plDisp -ne 1) { $viol += ('R7-DISPID VB6_MEMD_PANEL_LEFT must be defined exactly once, got ' + $plDisp) }
if ($plCase -ne 1) { $viol += ('R7-DISPATCH memInvokePanel must answer VB6_MEMD_PANEL_LEFT exactly once, got ' + $plCase) }
if ($plMoCall -ne 1) { $viol += ('R7-MO-CALL memberobj must call the Left getter exactly once, got ' + $plMoCall) }
if ($plDef -ne 1) { $viol += ('R7-DEF vb6forms_statusbar.c must define vb6_StatusBar_GetPanelLeft exactly once, got ' + $plDef) }
if ($plDecl -ne 1) { $viol += ('R7-DECL vb6forms_prop_ctrl.h must declare vb6_StatusBar_GetPanelLeft exactly once, got ' + $plDecl) }
if ($plRecog -ne 1) { $viol += ('R7-RECOGNIZER cgen_util_com must map Panel.Left to that exit exactly once, got ' + $plRecog) }
# 设计块那行 MinWidth 的三处（发码点 / RTL 定义 / header 声明）—— 同上，缺第三处是静默的
$saTxt = [System.IO.File]::ReadAllText((Join-Path $Root 'src\backend\detail\module\cgen_form_ctrl_style_apply.inc'))
$minEmit = ([regex]::Matches($saTxt, 'vb6_StatusBar_SetPanelMinWidthHm\(\(void\*\)"')).Count
$minDef = ([regex]::Matches($sbTxt, 'void vb6_StatusBar_SetPanelMinWidthHm\(void\* hwnd, int32_t index, int32_t hm\)')).Count
$minDecl = ([regex]::Matches($hdTxt, 'vb6_StatusBar_SetPanelMinWidthHm\(void\* hwnd, int32_t index, int32_t hm\);')).Count
if ($minEmit -ne 1) { $viol += ('R7-MINW-EMIT the design block must emit SetPanelMinWidthHm exactly once, got ' + $minEmit) }
if ($minDef -ne 1) { $viol += ('R7-MINW-DEF RTL must define SetPanelMinWidthHm exactly once, got ' + $minDef) }
if ($minDecl -ne 1) { $viol += ('R7-MINW-DECL header must declare SetPanelMinWidthHm exactly once, got ' + $minDecl) }
# 单位这一格用代码形状钉（不靠注释）：排版那一趟必须把 e->width **折**一次才递给 SB_SETPARTS
# —— 这一行存在就说明存储档不是像素（像素档是第十八刀那一版，它这里直接用 e->width）。
$unitFold = ([regex]::Matches($sbTxt, 'want\[i\] = e->width > 0 \? vb6_TwipToX\(e->width\)')).Count
if ($unitFold -ne 1) { $viol += ('R7-UNIT the layout must fold e->width through vb6_TwipToX exactly once (storage is not pixels), got ' + $unitFold) }

# ---------- R8: 账 #303 (§B132) —— 要占位的那两档各从剩余空间里扣一次，不多不少 ----------
# 病名（第二十二刀把 getter 改成交排版结果才露面的存量）：`SbLayout` 第一遍把内容档的实测宽写进
# want[i] 却**没累加进 fixedTotal** ⇒ 弹簧把同一格宽吃两次 ⇒ 末格右边界越过客户区
# （夹具实测 sum=6630 而 bar=6000，差的 630 缇 = 42 像素 = 那枚 sbrTime 面板的文本宽）。
# 钉成"恰好 2 次"的两个方向都是真形状：0 次 = 本账；3 次（把弹簧档也扣上）= 第十八刀撤掉的
# 那个"MinWidth 当加项"复发。判据面 = tests/ctrlstatusbar 的 SB36-TILES-BAR（BASE 上 False）。
$tileAcc = ([regex]::Matches($sbTxt, 'fixedTotal \+= want\[i\];')).Count
if ($tileAcc -ne 2) { $viol += ('R8-TILE-RESERVE want exactly 2 reserved tiles (fixed + contents), got ' + $tileAcc) }

# ---------- R9: 账 #298 (§B129) 并表收尾 —— 那三处以前各写各的，现在只许转调 ----------
# 并表之后"这一族有几处换算"这个问题只能有一个答案（vb6forms.c 里那 5 行，R2 已经钉死），
# 所以这里钉的是**调用点**：数对了说明没人偷偷把旧写法搬回来。三条墓碑 = 旧形状本身。
# 值面（写进台账 §B129，别只信这条哨兵）：picture_prop 是**同一条公式搬走**（构造上同值，且它
# 仍按 DC 的 DPI 而不是屏 DPI）；olecon 的 px->hm 两处 96 DPI 下逐值相同（1440/96=15 整除）并
# 顺带消掉 32 位溢出；olecon 的 hm->px 那一处由"两次截断"变成"一次就近"，实测差值只有 0 或 +1 px；
# ax_site_ext 的 DPI 从写死 96 换成真实 DPI —— 本机就是 96，那条只有形状能证。
$olcTxt = [System.IO.File]::ReadAllText((Join-Path $rtlRoot 'core\vb6forms\vb6forms_olecon.c'))
$ppTxt = [System.IO.File]::ReadAllText((Join-Path $rtlRoot 'core\vb6forms\vb6forms_picture_prop.c'))
$axsTxt = [System.IO.File]::ReadAllText((Join-Path $rtlRoot 'core\vb6forms\axsite\ax_site_ext.c'))
$olcPx = ([regex]::Matches($olcTxt, 'vb6_HimetricToPxX\(')).Count + ([regex]::Matches($olcTxt, 'vb6_HimetricToPxY\(')).Count
$olcHm = ([regex]::Matches($olcTxt, 'vb6_TwipsToHimetric\(')).Count
$olcTw = ([regex]::Matches($olcTxt, 'vb6_XToTwipX\(')).Count + ([regex]::Matches($olcTxt, 'vb6_YToTwipY\(')).Count
$ppOnDc = ([regex]::Matches($ppTxt, 'vb6_HimetricToPxOnDc\(')).Count
$axsF = ([regex]::Matches($axsTxt, 'vb6_HimetricToPxF\(')).Count + ([regex]::Matches($axsTxt, 'vb6_PxToHimetricF\(')).Count
if ($olcPx -ne 2) { $viol += ('R9-OLECON-SIZE olcSizeToObject must ask the hm->px authorities (X+Y) exactly twice, got ' + $olcPx) }
if ($olcHm -ne 4) { $viol += ('R9-OLECON-EXTENT the two SetExtent sites must fold px->hm through vb6_TwipsToHimetric 4 times, got ' + $olcHm) }
if ($olcTw -ne 4) { $viol += ('R9-OLECON-EXTENT px->twip must go through the X/Y authorities 4 times, got ' + $olcTw) }
if ($ppOnDc -ne 2) { $viol += ('R9-PICTURE-ONDCC Picture non-stretch must ask the formula authority twice (X+Y), got ' + $ppOnDc) }
if ($axsF -ne 4) { $viol += ('R9-AXSITE the OLE container transform must ask the float pair 4 times (2 ax x 2 dir), got ' + $axsF) }
$ppMul = ([regex]::Matches($ppTxt, 'MulDiv')).Count
$olcOld = ([regex]::Matches($olcTxt, '\* 15L \*')).Count
$axsOld = ([regex]::Matches($axsTxt, 'const double k')).Count
if ($ppMul -ne 0) { $viol += ('R9-TOMBSTONE picture_prop computes ' + $ppMul + ' conversion(s) itself again - the formula lives in vb6forms.c only') }
if ($olcOld -ne 0) { $viol += ('R9-TOMBSTONE olecon multiplied by a hardcoded 15 again (' + $olcOld + ' site(s)) - that 15 was the DPI=96 assumption') }
if ($axsOld -ne 0) { $viol += ('R9-TOMBSTONE ax_site_ext recomputes the himetric factor itself again (' + $axsOld + ') - DPI must come from vb6_DpiX/Y') }

$oneMapTot = 0
foreach ($L in $creatorLits) { $oneMapTot += $creatorTot[$L] }
if ($oneMapTot -lt 11) { $viol += ("R4-R6-ONEMAP the one-map census found $oneMapTot sites, floor 11 = it is idling") }

# ---------- R4: 防空转 ----------
if ($rtlFiles -lt 100) { $viol += "R4-RTL-FILES floor 100, got $rtlFiles" }
if ($HmFiles.Count -ne 1) { $viol += 'R4-LIST the allowlist itself changed size (must stay 1 file: vb6forms.c)' }
if ($HmWantTotal -ne 5) { $viol += 'R4-TOTAL the pinned total changed (must stay 5)' }
if ($pm.Count -lt 1) { $viol += 'R4-R5-CENSUS the pair scan found nothing = the sentinel is idling' }
if (($gaN + $stK + $stC + $cgN) -lt 5) { $viol += 'R4-R6-CENSUS the panels-list scan found nothing = the sentinel is idling' }
if (($coll + $collItem + $collEnum + $defN + $cbN + $cwN) -lt 15) { $viol += 'R4-R6-CENSUS2 the panels emit/exit scan found nothing = the sentinel is idling' }
if (($plNames + $plDisp + $plCase + $plMoCall + $plDef + $plDecl + $plRecog + $minEmit + $minDef + $minDecl + $unitFold) -lt 6) {
    $viol += 'R4-R7-CENSUS the Left/design-MinWidth registration scan found nothing = the sentinel is idling'
}
if (($sbTwip + $sbBack + $sbHm + $axDeleg + $twipDef + $hm2twDef) -lt 3) {
    $viol += 'R4-R2-CENSUS the authority-call scan found nothing = the sentinel is idling'
}

if ($viol.Count -eq 0) {
    Write-Host ("PASS static_sentinel_statusbar_panel_hm: emit Hm=" + $hm + "/" + $hmMin +
                " add=3 autosz=1 plain=0 left=$plEmit / himetric sites $sum in $($perFile.Count) files" +
                " (statusbar asks $sbTwip+$sbBack+$sbHm, local math $sbMul, old shape $sbPx, tiles reserved $tileAcc; callers olc=$olcPx+$olcHm+$olcTw pp=$ppOnDc axs=$axsF)" +
                " / rtl_files=$rtlFiles / panels coll=$coll item=$collItem " +
                "enum=$collEnum stale=$stale withwrite=$leakW exit=$defN+$cgN+$cbN+$cwN lists=$gaN+$stK+$stC " +
                "onemap=$oneMapTot inline=$inline")
    exit 0
}
Write-Host 'FAIL statusbar_panel_hm:'
foreach ($v in $viol) { Write-Host "  - $v" }
exit 1
