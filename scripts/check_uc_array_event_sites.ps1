# check_uc_array_event_sites.ps1 —— 控件数组的 UC 事件臂：一处权威 × 四条路（账 #222）
#
# 这一族以前有三种形参表同时存在于同一枚产物里（模块级按 VB 默认发 ByRef 的原型、
# prelude 逐元素再发一份**零形参**的前向声明、定义那一趟按事件 ABI 发 ByVal），
# 而 thunk / sink / attach 又按元素重复 24 遍同名符号 —— 实测 ucProgressCircular
# 因此编不过（Form1.c 24 条 C2198 + 1 条 C2084，两架构同形）。
#
# 修完之后形状靠这三条守住：
#   ① 元素键只有一个出口 CCodeGen::ctrlElemKey()，四条路都问它；
#   ② 事件 ABI 只在一个地方翻（applyEventHandlerAbi），登记处理器时立刻翻，
#      所以模块级声明与定义同源 —— 不许再出现第二条翻法或第二份前向声明；
#   ③ 「sink 供了哪一枚事件」是一张表（ucSinkEvents_），WM_LBUTTONUP 那两条 arm
#      都问它，少问一条就是一次点击双发。
# 每条都是钉死计数：多了少了都红（加路要连着改判据，不许静默漂）。

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$viol = @()

function Read-Src([string]$rel) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path -LiteralPath $p)) { $script:viol += ("缺文件: " + $rel); return '' }
    return [IO.File]::ReadAllText($p)
}
function Count-Of([string]$text, [string]$needle) {
    if ($needle -eq '') { return 0 }
    return ([regex]::Matches($text, [regex]::Escape($needle))).Count
}

$ctrl   = Read-Src 'src\backend\cgen_util_ctrl.cpp'
$decl   = Read-Src 'src\backend\decl\cgen_decl_proc.cpp'
$helpers= Read-Src 'src\backend\detail\util\cgen_helpers.inc'
$state  = Read-Src 'src\backend\detail\util\cgen_state.inc'
$prelude= Read-Src 'src\backend\detail\module\cgen_form_prelude.inc'
$create = Read-Src 'src\backend\detail\module\cgen_form_create_controls.inc'
$wndp   = Read-Src 'src\backend\detail\module\cgen_form_wndproc_subclass.inc'

# ---- E1: 元素键 = 一个定义 + 一条声明 + 四条调用路 ----
$e1def = Count-Of $ctrl   'std::string CCodeGen::ctrlElemKey('
$e1dec = Count-Of $helpers 'static std::string ctrlElemKey('
$e1call = (Count-Of $prelude 'ctrlElemKey(child.controlName, child.index)') +
          (Count-Of $create  'ctrlElemKey(ctrl.controlName, ctrl.index)') +
          (Count-Of $wndp    'ucSinkEvents_.count(ctrlElemKey(')
if ($e1def -ne 1)  { $viol += ("E1: ctrlElemKey 的定义应为 1 处，读到 " + $e1def) }
if ($e1dec -ne 1)  { $viol += ("E1: ctrlElemKey 的声明应为 1 处，读到 " + $e1dec) }
if ($e1call -ne 4) { $viol += ("E1: 问元素键的路应为 4 条（prelude / create / wndproc 两条 arm），读到 " + $e1call) }

# ---- E2: 事件 ABI 只许一处翻 ----
$e2def = Count-Of $decl 'void CCodeGen::applyEventHandlerAbi('
$e2dec = Count-Of $helpers 'void applyEventHandlerAbi(SubDecl& node);'
$e2call= (Count-Of $decl 'applyEventHandlerAbi(node)') + (Count-Of $decl 'applyEventHandlerAbi(sub)')
if ($e2def -ne 1)  { $viol += ("E2: applyEventHandlerAbi 的定义应为 1 处，读到 " + $e2def) }
if ($e2dec -ne 1)  { $viol += ("E2: applyEventHandlerAbi 的声明应为 1 处，读到 " + $e2dec) }
if ($e2call -ne 2) { $viol += ("E2: 调用点应为 2 处（声明趟 + 登记趟），读到 " + $e2call) }
# 登记之后**立刻**翻 —— 翻在 visit 里就赶不上模块级声明那一趟
if ((Count-Of $prelude 'prepareEventHandlerProc(module, handlerLower)') -ne 1) {
    $viol += 'E2: prelude 必须在登记处理器后立刻调 prepareEventHandlerProc 翻 ABI（一处）'
}
if ((Count-Of $decl 'bool CCodeGen::prepareEventHandlerProc(') -ne 1) {
    $viol += 'E2: prepareEventHandlerProc 的定义应为 1 处'
}

# ---- E3: 不许有第二份前向声明（旧形状） ----
if ((Count-Of $prelude 'frmPrefix + fnOrig + "(" + sig') -ne 0) {
    $viol += 'E3: prelude 里又出现了逐元素的前向声明（旧形状，24 枚元素各发一份就是 C2084/C2198）'
}
if ((Count-Of $prelude 'emitLine("static void " + thunkName') -ne 1) {
    $viol += 'E3: prelude 应只发一条 thunk 定义（逐元素命名靠 elemSfx，不靠重复发同名）'
}

# ---- E4: sink 供了哪一枚事件 = 一张表，两条 arm 都问它 ----
$e4dec = Count-Of $state 'std::set<std::string> ucSinkEvents_;'
$e4ins = Count-Of $prelude 'ucSinkEvents_.insert(attachKey'
$e4arm = Count-Of $wndp '!ucSinkEvents_.count(ctrlElemKey('
if ($e4dec -ne 1)  { $viol += ("E4: ucSinkEvents_ 的声明应为 1 处，读到 " + $e4dec) }
if ($e4ins -ne 1)  { $viol += ("E4: ucSinkEvents_ 的登记应为 1 处，读到 " + $e4ins) }
if ($e4arm -ne 2)  { $viol += ("E4: WM_LBUTTONUP 那两条 arm 都要问表，读到 " + $e4arm) }
# gate 写在"要不要装这一档"的那两处判据里 (顶层 + 容器子控件各一条)，
# 发码那一条 arm 只有一处 (if (info.hasClick))。少一条 gate = 一次点击双发。
$e4gate = (Count-Of $wndp '.count(ctrlElemKey(ctrl.controlName, ctrl.index) + ".click")') +
           (Count-Of $wndp '.count(ctrlElemKey(child.controlName, child.index) + ".click")')
$e4arm  = Count-Of $wndp 'if (info.hasClick) {'
if ($e4gate -ne 2) { $viol += ("E4: 两条 hasClick 判据都要问 sink 表，读到 " + $e4gate) }
if ($e4arm -ne 1)  { $viol += ("E4: WM_LBUTTONUP 的发码 arm 应为 1 处，读到 " + $e4arm) }

# ---- E5: 事件形参的 C 类型只有一处权威 (mapTypeRef) = 门 #350 那条回归的护栏 ----
# prelude 以前自带一张 kTypeMapPre (七档 + default int32_t), 与 emitEventSink 的回调 typedef
# 容器侧处理器原型两处**不同源**: 缺 Variant 一档 => thunk 声明 int32_t 而 typedef 与处理器
# 都是 vb6_VARIANT => charts_ucTreeMaps 两架构 C2440; Integer 那一档表给 int32_t、mapTypeRef
# 给 int16_t => VBFlexGrid 的 Button/Shift/Cancel 形参全按错宽度接 (发送侧按 typedef 发)。
# 现在表已撤, 类型名现拼一枚 SimpleTypeRef 喂 mapTypeRef。钉死两头: 旧表不许回来、新出口只一处。
if ((Count-Of $prelude 'kTypeMapPre') -ne 0) {
    $viol += 'E5: prelude 里又出现了第二张事件形参类型表 kTypeMapPre (须与 typedef/处理器同源)'
}
$e5type = Count-Of $prelude 'mapTypeRef(&tyRefPre)'
$e5node = Count-Of $prelude 'SimpleTypeRef tyRefPre('
if ($e5type -ne 1) { $viol += ('E5: prelude 问 mapTypeRef 的路应为 1 条，读到 ' + $e5type) }
if ($e5node -ne 1) { $viol += ('E5: prelude 现拼 SimpleTypeRef 应为 1 处，读到 ' + $e5node) }

# ---- C: 账 #226 的 click 落点 —— 三份权威 + 顺序 (布局式初始化) ----
# vb6_UserControlDesc 是**按位置**初始化式, 所以『新槽追加到末尾』不是风格问题,
# 而是错位一个指针宽就运行期 AV 的那种问题 (rev22 那段教训)。这里钉两头:
# RTL 的结构体末尾必须是 click, cgen 的初始化式末槽也必须是 click。
$hdrTxt = Read-Src 'src\rtl\core\vb6forms\vb6forms_controls.h'
$uhoTxt = Read-Src 'src\rtl\core\vb6forms\uc\uc_host_window.c'
$iClick = $hdrTxt.IndexOf('(*click)(void* me);')
$iResize = $hdrTxt.IndexOf('(*designResize)(void* me, const char* ctrlName);')
$iClose  = $hdrTxt.IndexOf('} vb6_UserControlDesc;')
if ($iClick -lt 0 -or $iResize -lt 0 -or $iClose -lt 0) {
    $viol += 'C1: 结构体里找不到 click / designResize / 收尾括号 (权威变了先看清再改)'
} elseif (-not ($iResize -lt $iClick -and $iClick -lt $iClose)) {
    $viol += 'C1: click 槽必须是 vb6_UserControlDesc 的**最后一枚**成员 (顺序=布局)'
}
# 运行期: 必须在 MouseUp 转调之后才转 Click (VB6 的顺序)
$nHandoff = (Count-Of $uhoTxt 'r->desc->click(r->me)')
if ($nHandoff -ne 1) { $viol += ('C2: RTL 的 click 转调应为 1 处，读到 ' + $nHandoff) }
$iUp = $uhoTxt.IndexOf('hook(r->me, button, 0, sx, sy);')
$iCl = $uhoTxt.IndexOf('r->desc->click(r->me)')
if ($iUp -lt 0 -or $iCl -lt 0 -or -not ($iUp -lt $iCl)) {
    $viol += 'C2: click 转调必须排在 MouseUp 转调之后 (VB6: 抬起之后才发 Click)'
}
# 发码: 末槽 + 封装各一处, 且末槽在 designResize 之后
$genTxt = Read-Src 'src\backend\module\cgen_form.cpp'
$nWrap = (Count-Of $genTxt '_ucHostClick(void* me) {')
$nSlot = (Count-Of $genTxt '_ucHostClick  /*')
if ($nWrap -ne 2) { $viol += ('C3: ucHostClick 的封装应为 2 条 (有/无处理器两支)，读到 ' + $nWrap) }
if ($nSlot -ne 1) { $viol += ('C3: desc 初始化式里的 click 末槽应为 1 条，读到 ' + $nSlot) }
$gRes = $genTxt.IndexOf('_ucHostDesignResize"))')
$gCli = $genTxt.IndexOf('_ucHostClick  /*')
if ($gRes -lt 0 -or $gCli -lt 0 -or -not ($gRes -lt $gCli)) {
    $viol += 'C3: 发码里 click 必须排在 designResize 之后 (与结构体同序)'
}
# ---- C4: 账 #227 —— desc 的鼠标/点击那一族, 每一槽都必须有落点 ----
# cgen 是按位置把这张表填满的, 所以 RTL 少转调哪一格都不会有编译症状, 只会让那一格
# 事件永不出响: click 是账 #226, dblClick 是账 #227 (同一个形状栽了两次)。
# 于是这里钉**逐槽非零**, 而不是只钉某一条转调的行数。
foreach ($s4 in @('mouseDown', 'mouseUp', 'mouseMove', 'click', 'dblClick')) {
    $n4 = Count-Of $uhoTxt ('desc->' + $s4)
    if ($n4 -lt 1) {
        $viol += ('C4: 宿主没有转调 desc->' + $s4 + ' 这一槽 (读到 ' + $n4 + ') - 那一格事件永远不出声')
    }
}
# 而且 dblClick 的落点必须独立在 WM_LBUTTONDBLCLK 那一档里: 挂进鼠标那一档就会把
# MouseDown/MouseUp 一起双发 (物理双击 Windows 发的是 DOWN/UP/DBLCLK/UP)。
$i4case = $uhoTxt.IndexOf('case WM_LBUTTONDBLCLK:')
$i4call = $uhoTxt.IndexOf('r->desc->dblClick(r->me)')
if ($i4case -lt 0 -or $i4call -lt 0 -or -not ($i4case -lt $i4call)) {
    $viol += 'C4: dblClick 转调必须在 case WM_LBUTTONDBLCLK 那一档里 (不挂鼠标那一段)'
}
$c4n = Count-Of $uhoTxt 'r->desc->dblClick(r->me)'
if ($c4n -ne 1) { $viol += ('C4: dblClick 的转调应为 1 处，读到 ' + $c4n) }

# ---- 普查（不判红，给下一轮留证据） ----
$cenSink = Count-Of $prelude 'ucEventAttach_[attachKey]'
$cenSfx  = Count-Of $prelude 'elemSfx'
Write-Host ("census: key(def/dec/calls)={0}/{1}/{2} abi(def/dec/calls)={3}/{4}/{5} sink-table(dec/ins/gate/arm)={6}/{7}/{8}/{9} attach-store={10} elemSfx={11} evt-type-exits(mapTypeRef/SimpleTypeRef)={12}/{13}" -f `
    $e1def, $e1dec, $e1call, $e2def, $e2dec, $e2call, $e4dec, $e4ins, $e4gate, $e4arm, $cenSink, $cenSfx, $e5type, $e5node)
Write-Host ("census-click: handoff={0} wrap={1} slot={2} struct-last-click={3} dblclick-dispatch(now)={4} dbl-click-calls={5}" -f `
    $nHandoff, $nWrap, $nSlot, ($iClick -lt $iClose), (Count-Of $uhoTxt 'desc->dblClick'), $c4n)

if ($viol.Count -gt 0) {
    Write-Host ("FAIL: uc-array-event 哨兵, {0} 条" -f $viol.Count)
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    exit 1
}
Write-Host ("PASS: 元素键 1+1+4 / ABI 1+1+2+prep / 无第二份前向声明 / sink 表 1+1 且 gate {0} 条、arm {1} 条 / click 落点 1+2+1 且两份权威同序 / 事件形参类型一处权威 / 鼠标族五槽都有落点" -f $e4gate, $e4arm)
exit 0
