#include "backend/cgen.hpp"
#include <cstdlib>
#include <cstdio>
#include "common/owned_temp_keeps.hpp"  // 账 #464 §B164: 回收刀的**唯一权威**（那张表由 RTL 推导 ⇒ 哨兵双向对账）
#include "common/float_literal.hpp"  // 账 #188: 浮点字面量的单一出口
#include "common/int_literal.hpp"   // 账 #194: 整数字面量的单一出口
#include "common/canvas_drawing.hpp"  // 账 #232①: 画布动词的单一出口 (名字 / 接收者 / 实参形状)
#include "common/event_param_abi.hpp" // 账 #610 第345刀: 事件形参「按不按地址递」的唯一权威
#include "common/form_pseudo.hpp"     // 账 #278 §B109: 属性名→读函数的单一出口 (两层都问它)
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_util_ctrl.cpp: 控件属性映射 (读/写函数名 + hWnd 参数 + 默认属性) + Variant 包装 ---


// ============================================================
// P7.5: 控件属性 → RTL读取函数名映射
// ============================================================

// ============================================================
// P20-42: 控件属性 → 返回类型
//
// 为什么需要这张表: inferExprType 对 `对象.成员` 的兜底是 lookupModule(memberName),
// 只按成员**裸名**查模块符号, 于是和模块级/内置符号同名的属性会被顶掉。
// 实测 `SSTab1.Tab` 撞上内置函数 `Tab` (vb6_Tab 返回 BSTR) → 判成 Vb6Type::String
// → 字符串拼接不套 vb6_CStr(vb6_VariantFromValue(...)) → int32_t 当 BSTR 解引用 → AV。
//
// 只对"会撞名"的属性登记即可: 其余属性继续走原有兜底 (lookupModule 找不到就返回
// Variant, 由 _Generic 安全包装)。这里登记的都按 Long —— RTL 侧这些 getter 就是
// int32_t, C 层口径一致。
// ============================================================
Vb6Type CCodeGen::controlPropType(FrmControlType ctrlType, const std::string& propName) const {
    std::string p = propName;
    std::transform(p.begin(), p.end(), p.begin(), ::tolower);

    // ---- 通用属性 (所有控件都适用) ----
    //
    // **必须先于** SSTab 那段, 而且必须存在 —— 这不是 SSTab 一个控件的问题:
    // `Left` / `Right` / `Mid` 全是 VB6 内置函数 (cgen_expr_ident_builtin.inc),
    // 其中 Left/Right/Mid 返回 **String**。于是 `lblPage0.Left` 被判成 String →
    // 生成码不套数值转换 → int 当 BSTR 解引用 → 0xC0000005。
    // 实测 SSTab 夹具: TS1 的 SSTab1.Tabs 正常, TS22 的 lblPage0.Left 崩。
    // 这几个 RTL getter 的 C 返回类型都是 int, 故一律 Long。
    // 账 #124: `Enabled` / `Visible` 在 VB6 那边是 **Boolean** —— 以前和左/上/宽/高一起
    // 按 Long 登记, 于是 `CStr(cb.Enabled)` 折成 vb6_CStrLong 打出 -1、装箱成 VT_I4,
    // 而 VB6 打 "True"、装 VT_BOOL(11)。两个 RTL getter 的 C 返回型是 int 无妨:
    // 所有消费面都按"VB6 的 -1/0"取值, 需要窄化的地方 (boxToVariant / wrapVariantValue)
    // 自带 (int16_t) 显式收窄。
    if (p == "left" || p == "top" || p == "width" || p == "height") {
        return Vb6Type::Long;
    }
    // 账 #231: 这一档原来住在 inferExprType 里 (C29-1a 那份自带名单), 搬进来之后
    // 控件属性的类型**只有这张表**。名单逐条照搬、答案一字未改。
    // 为什么仍要问 "不是 Unknown": Unknown = 工程外的自定义 OCX / UC 实例 (VBFlexGrid
    // 那一族), 它们的属性面由类型库说话, `Shape` / `BorderWidth` 这类名字在 OCX 上
    // 未必是同一件事 —— C29-1a 当年就是靠这道闸只放行内建控件的。
    if (ctrlType != FrmControlType::Unknown) {
        // Shape/PictureBox 的 Shape/FillStyle/… 与 Line 的四条端点坐标: RTL getter
        // 全是 int32_t。BorderStyle 在 Shape/Line 上是"画笔线型"(0..6)、在其他控件上
        // 是窗口边框样式 —— 两条读法在 getControlPropReadFn 里分家, 类型这侧同档。
        if (p == "shape" || p == "fillstyle" || p == "borderwidth" || p == "borderstyle"
            || p == "fillcolor" || p == "bordercolor"
            || p == "x1" || p == "y1" || p == "x2" || p == "y2") {
            return Vb6Type::Long;
        }
    }
    if (p == "visible" || p == "enabled") {
        return Vb6Type::Boolean;
    }
    // C30-3: MaskEdBox 的属性面。Text/ClipText/Mask/PromptChar 都走 RTL 的 BSTR
    // （String）；PromptInclude 在 VB6 那边是 Boolean；Sel* 与编辑控件同口径。
    // 没有这一格时，`s = m.ClipText` 会被当成 Variant 走打包路 → 拿到的是空串
    // （与 TabStrip 链式属性那条老坑同型）。
    if (ctrlType == FrmControlType::MaskEdBox) {
        if (p == "text" || p == "cliptext" || p == "mask"
            || p == "promptchar" || p == "seltext") return Vb6Type::String;
        if (p == "promptinclude") return Vb6Type::Boolean;
        if (p == "selstart" || p == "sellength") return Vb6Type::Long;
    }
    // C30-4: MSComm 的属性面。Settings / Input 走 RTL 的 BSTR（String）；PortOpen / dtrenable /
    // RTSEnable / EOFEnable / CTS·DSR·CDHolding 在 VB6 那边是 Boolean；其余是 Long。
    // 没登记的症状与 MaskEdBox 同型：`s = MSComm1.Input` 被当 Variant 装箱成空串。
    if (ctrlType == FrmControlType::MSComm) {
        if (p == "settings" || p == "input") return Vb6Type::String;
        // C30-4b：C3 扩展里的字符串那几档（PortList / CommDevice / InputBytes）
        if (p == "commdevice" || p == "portlist" || p == "inputbytes") return Vb6Type::String;
        if (p == "portopen" || p == "dtrenable" || p == "rtsenable" || p == "eofenable"
            || p == "ctsholding" || p == "dsrholding" || p == "cdholding") return Vb6Type::Boolean;
        if (p == "commport" || p == "inputmode" || p == "inputlen"
            || p == "inbuffercount" || p == "outbuffercount"
            || p == "inbuffersize" || p == "outbuffersize"
            || p == "rthreshold" || p == "sthreshold" || p == "handshaking"
            || p == "commevent"
            || p == "txtotal" || p == "rxtotal"
            || p == "portcount" || p == "readintervaltimeout"
            || p == "readtotaltimeout" || p == "writetimeout") return Vb6Type::Long;
    }
    // C30-5: Adodc 的属性面。ConnectionString / RecordSource 走 RTL 的 BSTR（String）；
    // 其余都是枚举/数值（Long）。不登记的症状：`s = Adodc1.RecordSource` 被当 Variant 装箱成空串。
    if (ctrlType == FrmControlType::Adodc) {
        if (p == "connectionstring" || p == "recordsource") return Vb6Type::String;
        if (p == "lastrowcount" || p == "opened") return Vb6Type::Long;   // recordset 不登记（对象，与 Data 同口径）
    }
    // C30-5: DataGrid 的属性面。Text / DataSource 是 String；ColumnHeaders/Allow* 是 Boolean；
    // Row/Col/RowCount/ColCount/RowHeight/DefColWidth 是 Long。
    if (ctrlType == FrmControlType::DataGrid) {
        if (p == "text" || p == "datasource") return Vb6Type::String;
        if (p == "columnheaders" || p == "allowaddnew" || p == "allowupdate"
            || p == "allowdelete") return Vb6Type::Boolean;
        if (p == "row" || p == "col" || p == "rowcount" || p == "colcount"
            || p == "rowheight" || p == "defcolwidth") return Vb6Type::Long;
    }
    // C30-6: MSHFlexGrid / MSFlexGrid 的属性面。
    // String：Text / TextMatrix(r,c) / TextArray(i) / FormatString / DataSource；
    // Boolean：七条开关位；其余（含 ColWidth/RowHeight/ColAlignment 的**元素值**）是 Long。
    // 不登记的症状与 Adodc 同：`s = grd.TextMatrix(1,2)` 被当 Variant 装箱成空串，
    // 表格画得出来但读回全假。
    if (ctrlType == FrmControlType::MshFlexGrid) {
        if (p == "text" || p == "textmatrix" || p == "textarray"
            || p == "formatstring" || p == "datasource") return Vb6Type::String;
        if (p == "wordwrap" || p == "editable" || p == "focusrect"
            || p == "allowuserresizing" || p == "mergecells" || p == "redraw") return Vb6Type::Boolean;
        if (p == "rows" || p == "cols" || p == "fixedrows" || p == "fixedcols"
            || p == "row" || p == "col" || p == "rowsel" || p == "colsel"
            || p == "toprow" || p == "leftcol" || p == "colwidth" || p == "rowheight"
            || p == "colalignment" || p == "rowheightmin" || p == "gridlines"
            || p == "scrollbars" || p == "selectionmode" || p == "fillstyle"
            || p == "backcolor" || p == "forecolor" || p == "fillcolor"
            || p == "cellbackcolor" || p == "cellforecolor" || p == "mouserow"
            || p == "mousecol" || p == "cellheight" || p == "cellwidth"
            || p == "cellleft" || p == "celltop" || p == "sort") return Vb6Type::Long;
    }
    // C30-2: TabStrip 的标量面。四个 Client* 与三条行为位的 RTL getter 都是 int32_t;    // MultiRow/HotTracking/ShowTips 在 VB6 那边是 Boolean。
    // ⚠ Tabs(i).Caption 这类**链式成员**不在本表（这一问只看“控件.属性”一形），
    // 所以 `"cap=" & TabStrip1.Tabs(n).Caption` 这种接法仍会被当 Variant 装箱成空串；
    // 先读进 String 变量再用就对了（tools/tabstrip_manual 的写法）。
    if (ctrlType == FrmControlType::TabStrip) {
        if (p == "clientleft" || p == "clienttop"
            || p == "clientwidth" || p == "clientheight") return Vb6Type::Long;
        if (p == "multirow" || p == "hottracking" || p == "showtips") return Vb6Type::Boolean;
    }
    // C29-SL-e: `ToolTipText` / `Tag` 是 VB6 的 **String** 通用属性，但两条 RTL getter 原本
    // 声明成 `void*`，而装箱那一步用的是 C11 `_Generic`（`vb6rtl_variant.h` 的表：
    // "其他指针 (void*/class*/type*) -> VariantObject"）⇒ 裸指针被当**对象**装箱。
    // 实测（探针工程 .build/sltt，改之前）：`CStr(Slider1.ToolTipText)` 打**空**、
    // `TypeName(...)` 答 **"Object"**、`If Slider1.Tag = "tag1"` 答**假**（那一条比
    // ToolTipText 更绕：比较面先装箱再 CStr，等于拿 "" 去比），而 `Len(...)` 与
    // `InStr(...)` 这些**直接拿指针**的面反而是对的 —— 同一枚属性两种答案，就是没登记的证状。
    // 登记成 String 之后所有消费面统一按 BSTR 取值（与 `vb6_GetControlText` 那条同档），
    // RTL 侧的返回型也一起改成 `wchar_t*`，让**万一**还剩的装箱落到 VT_BSTR。
    // 同一段里 `MouseIcon` 的 getter 也是 `void*`，那条**本来就是对象**（FlexGrid 的
    // `CellPicture` 同型），装箱成对象是对的。同一句把 `Hwnd` 也一起算成对象是错的 ——
    // VB6 那边 `hWnd` / `hDC` 是**句柄数值**（`TypeName(Form1.hWnd)` 答 "Long"），
    // 订正与读数见账 §B136 / 下面那两行。
    if (p == "tooltiptext" || p == "tag") {
        return Vb6Type::String;
    }
    // 账 #154（C29-SL-q）: `FontName` 以前**根本没登记**，于是同一枚属性两种答案 ——
    // `Len(Text1.FontName)` 直接拿指针、答得对（探针里 "Consolas" 读出 8），而
    // `Debug.Print ... & Text1.FontName` 先装箱再 CStr、打出来是空串。看着像"字体名没设上"，
    // 其实窗口里早就换上了（同一枚控件的 FontPixelHeight 证人当场翻案）。
    // CommonDialog 那一条的 getter 也回字符串（vb6_CdGetFontName，ChooseFont 的字段），同一档。
    if (p == "fontname") {
        return Vb6Type::String;
    }
    // 账 #83(a) 的读侧那半（与 #124 / #136 同一族）：`vb6_GetTabStop` 答的就是 VB6 的 -1/0，
    // 但以前没登记 ⇒ `CStr(Check1.TabStop)` 打的是 `-1`/`0`、`If ... Then` 靠装箱碰巧对。
    // 登记成 Boolean 之后打印面就是 True/False（探针 `.build/sltab` 里 VarType 读回 3 = Integer 就是没登记的证状）。
    if (p == "tabstop") {
        return Vb6Type::Boolean;
    }

    // 账 §B136（第二十八刀）: 句柄属性登记成**数值**，不是对象。这两枚以前根本没登记 ⇒ 类型答
    // Unknown ⇒ 消费面一律走装箱那条路，而 `vb6_VariantFromValue` 的 `_Generic` 把**任何指针**都投给
    // `vb6_VariantObject`（rtl 那张表自己写的最后一行），于是 `If a.hWnd = b.hWnd` 发成
    // `vb6_VarCmpEq(&对象档, &对象档)` —— 而 `vb6_VarToDouble_internal` 只认 VT_I2/I4/R4/R8/BOOL/BSTR、
    // `default: return 0.0` ⇒ **两枚不同窗口的句柄恒答相等**（`<>` 恒假；探针 `.build/b606_inline`：
    // 单独打 `TF(a.hwnd <> b.hwnd)` 与并进 And 链同答 False，先 CLng 落进声明过的 Long 才答 True）。
    // 档位取 LongPtr = 与 RTL 的 `void*` / `intptr_t` 同宽（账 #180 已定「句柄按指针宽度」这一档），
    // 也比 VB6 的 Long 更适合 x64；登记之后比较面走原生 `==`/`!=`（与「声明过的 LongPtr 变量」同形）。
    if (p == "hwnd" || p == "hdc") {
        return Vb6Type::LongPtr;
    }

    if (ctrlType == FrmControlType::SSTab) {
        if (p == "tabs" || p == "tab" || p == "taborientation" || p == "tabstyle"
            || p == "tabsperrow") {
            return Vb6Type::Long;
        }
        // 账 #128: WordWrap 在 VB6 是 Boolean，而 `vb6_SSTab_GetWordWrap` 给的就是 -1/0
        // （实测见 029 §#128）⇒ 与 #124 那两条同档，`CStr` 打 True/False、装箱走 VT_BOOL。
        if (p == "wordwrap") return Vb6Type::Boolean;
    }
    // C29-OLE: OLEType/OLETypeAllowed/SizeMode/AutoActivate 是 Long 语义 (枚举/布尔);
    // Class/SourceDoc/SourceItem 是 String (RTL getter 返回 wchar_t*, 包装层会转 BSTR)。
    if (ctrlType == FrmControlType::OLE) {
        if (p == "oletype" || p == "oletypeallowed" || p == "sizemode"
            || p == "autoactivate" || p == "displayasicon" || p == "autoverbmenu"
            || p == "borderstyle") {
            return Vb6Type::Long;
        }
        if (p == "class" || p == "sourcedoc" || p == "sourceitem") {
            return Vb6Type::String;
        }
    }
    if (ctrlType == FrmControlType::TreeView) {
        // C29-8a: 同理 —— vb6_TreeView_Get* 全是 int32_t (布尔按 VB6 的 -1/0 给)。
        // 不登记就走 inferExprType 的兜底, 那条按成员裸名查符号, 判成 Variant/String
        // 都不匹配 C 侧的 int32_t (SSTab1.Tab 撞内置 Tab 那次 AV 的同族)。
        if (p == "linestyle" || p == "indentation") {
            return Vb6Type::Long;
        }
        // IconSize 是 XtremeSuiteControls.TreeView 的扩展 (像素), RTL getter 给 int32_t。
        // 原生 SysTreeView32 没有这一格 —— RTL 拿它当"预留图标列"的宽度 (挂 16x16
        // ImageList)，见 vb6forms_treeview.c 的 SetIconSize。
        if (p == "iconsize") {
            return Vb6Type::Long;
        }
        // 账 #128: 这三条 getter 走 `vb6_TvBitAsVbBool` ⇒ 答的就是 VB6 的 -1/0，
        // 所以它们与 #124 那两条同档（以前跟着"全是 int32_t"那一句一起吃 Long）。
        if (p == "checkboxes" || p == "hottracking" || p == "hideselection") {
            return Vb6Type::Boolean;
        }
    }
    if (ctrlType == FrmControlType::DTPicker) {
        // C29-DT-a: 同一条纪律 —— vb6_DTP_Get* 除 CustomFormat 外全是 int32_t
        // (布尔按 VB6 的 -1/0 给); CustomFormat 的 getter 返回 wchar_t* ⇒ String。
        if (p == "format" || p == "updown"
            || p == "calendarbackcolor" || p == "calendarforecolor"
            || p == "calendartrailingforecolor" || p == "calendartitlebackcolor"
            || p == "calendartitleforecolor" || p == "idealwidth") {
            return Vb6Type::Long;
        }
        // 账 #128: CheckBox 的 getter 是 `vb6_DtpBitAsVbBool` ⇒ -1/0，VB6 那边也是 Boolean。
        if (p == "checkbox") return Vb6Type::Boolean;
        if (p == "customformat") return Vb6Type::String;
        // C29-DT-b: 三条 Date 型属性。登记成 Date 而不是让它落到兜底 —— Date 在 C 层就是
        // double（`Dim d As Date` 实测发成 `double d`），不登记的话 inferExprType 那条按成员
        // 裸名查符号的兜底会判成 Variant/String，与 RTL 侧的 double 返回值不匹配（TreeView
        // 那批同款坑，见上方注释）。HasDate 是本项目的扩展读数，Long。
        if (p == "value" || p == "mindate" || p == "maxdate") return Vb6Type::Date;
        if (p == "hasdate") return Vb6Type::Long;
    }
    if (ctrlType == FrmControlType::MonthView) {
        // C29-MV-a: 同一口径 —— vb6_MV_Get* 的 getter 全是 int32_t（布尔按 VB6 的 -1/0 给，
        // 色值是 COLORREF 那个 32 位），漏登记就会落到兜底那条按成员裸名查符号的路，
        // 判成 Variant/String 就跟 C 层不匹配（SSTab1.Tab 那次 AV 的同族）。
        // Value / SelStart / SelEnd 是 Date 型，由 MV-b 那格登记。
        if (p == "maxselcount" || p == "backcolor" || p == "forecolor"
            || p == "titlebackcolor" || p == "titleforecolor" || p == "trailingforecolor"
            || p == "monthbackcolor" || p == "minreqwidth" || p == "minreqheight"
            || p == "monthcount") {
            return Vb6Type::Long;
        }
        // 账 #128: 这三条就是样式位读数，getter 实测答 -1/0（ShowToday 还是**反**的那一位）
        // ⇒ 归 Boolean，与 #124 / SL-b 的 SelectRange 同一口径。
        if (p == "multiselect" || p == "showweeknumbers" || p == "showtoday") {
            return Vb6Type::Boolean;
        }
        // C29-MV-b：Date 那三格与 DTPicker 同一条口径（C 层就是裸 double，不装箱）。
        if (p == "value" || p == "selstart" || p == "selend") return Vb6Type::Date;
    }
    // C30-UD-a: UpDown（原生 msctls_updown32）。值面四条 + 行为面两条都是 int32_t
    // （getter 全是 int32_t，不登记就落到兜底那条按裸名查符号的路，判成 Variant/String）。
    if (ctrlType == FrmControlType::UpDown) {
        if (p == "min" || p == "max" || p == "value" || p == "increment"
            || p == "alignment" || p == "orientation") {
            return Vb6Type::Long;
        }
        // Wrap / SyncBuddy 在 VB6 是 Boolean ⇒ 按 #124 那条口径登记（getter 给 -1/0）。
        if (p == "wrap" || p == "syncbuddy") return Vb6Type::Boolean;
        // BuddyControl 是控件名字符串（类型库里是 VB_UserDefine 的引用位，运行时读回为空）
        if (p == "buddycontrol") return Vb6Type::String;
    }
    if (ctrlType == FrmControlType::Slider) {
        // C29-SL-a: vb6_Slider_Get* 全是 int32_t（两条证人读数也是数值），不登记就落到
        // 兜底那条按成员裸名查符号的路，判成 Variant/String 跟 C 层不匹配（SSTab1.Tab 那次
        // AV 的同族）。值面那几条（Min/Max/Value/Small·LargeChange/Sel*）由 SL-b 登记。
        if (p == "orientation" || p == "tickfrequency"
            || p == "travelisvert" || p == "tickpresent"
            || p == "min" || p == "max" || p == "value"
            || p == "smallchange" || p == "largechange"
            || p == "selstart" || p == "selend"
            // C29-SL-h: 三条都是数值（TickStyle 在类型库里是 VT_USERDEFINED 的那张枚举，
            // VB6 侧读回来就是 0..3 这一个数，与 Orientation 同档）。
            || p == "tickstyle" || p == "getnumticks" || p == "channeltop"
            // C29-SL-i: VB6 选区那一对的第二条（SelStart + SelLength，两条 VT_I4）。
            || p == "sellength") {
            return Vb6Type::Long;
        }
        // SelectRange 在 VB6 是 Boolean ⇒ 按 #124 那条口径登记（getter 给的就是 -1/0），
        // 这样 `CStr(sld.SelectRange)` 打 True/False、装箱走 VT_BOOL。
        if (p == "selectrange") return Vb6Type::Boolean;
        // 账 #148 的判据证人同样是 -1/0 ⇒ 同一档。
        if (p == "tooltipregistered") return Vb6Type::Boolean;
        // C29-SL-k：`Text` 在类型库里是 VT_BSTR（那颗气泡里的串，**不是**窗口标题），
        // `BubbleText` 读的是 tooltip 宿主里的工具文本 ⇒ 两条都是 String 档。
        if (p == "text" || p == "bubbletext") return Vb6Type::String;
        if (p == "bubblevisible") return Vb6Type::Boolean;      // -1/0
        if (p == "textposition" || p == "bubbletop") return Vb6Type::Long;
    }
    if (ctrlType == FrmControlType::RichTextBox) {
        // C29-RT-a: 同一口径。vb6_RTB_Get* 除 SelText 外全是 int32_t（布尔按 VB6 的 -1/0 给，
        // ScrollBars 是枚举、两条量程是数值）；SelText 的 getter 返回 wchar_t* ⇒ String。
        // Text 不在这里 —— 它走通用那条 vb6_GetControlText（与 TextBox 同一格）。
        if (p == "selstart" || p == "sellength" || p == "maxlength"
            || p == "scrollbars" || p == "vscrollrange"
            || p == "hscrollrange") {
            return Vb6Type::Long;
        }
        // 账 #128: ReadOnly 读样式位、WordWrap 是自存但**写侧就归化成 -1/0**（setter 里
        // `*(int32_t*)h = on ? -1 : 0`）⇒ 两条答的都是 VB6 的 True/False。
        if (p == "readonly" || p == "wordwrap") return Vb6Type::Boolean;
        if (p == "seltext") return Vb6Type::String;
        // C29-RT-b: 格式面。布尔四条 + 对齐 + 三个缩进 + 色值 = int32_t ⇒ Long；
        // 字号 getter 是 float（与通用那条 FontSize 同一口径）；FontName 的 getter 回 BSTR。
        if (p == "selbold" || p == "selitalic" || p == "selunderline" || p == "selstrikethru"
            || p == "selcolor" || p == "selalignment" || p == "selindent"
            || p == "selrightindent" || p == "selhangingindent") {
            return Vb6Type::Long;
        }
        if (p == "selfontsize") return Vb6Type::Single;
        if (p == "selfontname") return Vb6Type::String;
        // C29-RT-c: TextRTF 的 getter 回 wchar_t* ⇒ String（判成数值就是把指针当数读，
        // SSTab1.Tab 那一族同型）。LoadFile / SaveFile / Find 是**方法**，不在属性表里，
        // 走 cgen_expr_call_callee_withm.inc 那条改道。
        if (p == "textrtf") return Vb6Type::String;
    }
    if (ctrlType == FrmControlType::ListView) {
        // 账 #128-b: 这六位在 VB6 是 Boolean，RTL 侧现在也真的答 -1/0（存进去就归一化，见
        // vb6forms_listview.c）。以前它们**根本没登记** ⇒ 走 inferExprType 的兜底、被当成
        // 数值装箱，`CStr(ListView1.GridLines)` 打的是 1 而不是 True。
        // 数值那几位（View / SortKey / SortOrder / LabelEdit / Sorted）刻意**不跟着登记**：
        // 它们要么是枚举要么口径待拍，继续走今天那条装箱路，发码一字不动。
        if (p == "gridlines" || p == "fullrowselect" || p == "multiselect"
            || p == "checkboxes" || p == "hidecolumnheaders" || p == "allowcolumnreorder") {
            return Vb6Type::Boolean;
        }
    }
    if (ctrlType == FrmControlType::OptionButton) {
        // 账 #128-b: OptionButton.Value 单开了一对专桩（读映射成 -1/0、写把非 0 折回
        // BST_CHECKED —— BM_SETCHECK 不吃 -1），所以这里才敢登 Boolean。
        // CheckBox.Value **不动** —— VB6 那边它是三态 Integer(0/1/2)，本来就是数值档。
        if (p == "value") return Vb6Type::Boolean;
    }
    if (ctrlType == FrmControlType::Winsock) {
        // C29-WS-a: 与 C 层签名对齐是这里的唯一目的 —— 六条 getter 回 int32_t ⇒ Long，
        // 两条回 BSTR ⇒ String（判成数值就是把指针当数读，SSTab1.Tab 那一族同型缺陷）。
        if (p == "protocol" || p == "state" || p == "localport" || p == "remoteport"
            || p == "bytesreceived" || p == "bytetransferred") {
            return Vb6Type::Long;
        }
        if (p == "localip" || p == "remotehost") return Vb6Type::String;
    }
    if (ctrlType == FrmControlType::Toolbar) {
        // C29-5a: 同一口径 —— 这四条的 RTL getter 都是 int32_t, 判成 Variant/String
        // 就跟 C 层不匹配 (SSTab1.Tab 那次 AV 的同族)。
        if (p == "showtips" || p == "textstyle" || p == "allowcustomize" || p == "align") {
            return Vb6Type::Long;
        }
        // `tb1.Buttons.Count` 的读函数在 cgen_util_com.cpp 里改道, 那条路不查控件属性表,
        // 于是成员名 "count" 单独进来 —— 不登记成 Long 就被装箱, `Count = 3` 恒假
        // (实测: 值是对的 3, 打出来却是空串, 比较也全 N)。
        if (ctrlType == FrmControlType::Toolbar && p == "count") return Vb6Type::Long;
    }
    if (ctrlType == FrmControlType::DriveListBox || ctrlType == FrmControlType::DirListBox
        || ctrlType == FrmControlType::FileListBox) {
        // 账 #231 (原 C29-1b 那份名单): 不登记 ⇒ 判成 Variant ⇒
        // `File1.FileName = File1.List(0)` 这类比较走 vb6_VarCmpEq 而不是 vb6_StrCmp ——
        // 右边 (RTL 声明 void*) 装箱成 VT_UNKNOWN, 于是同一条读数 x64 为真、x86 为假。
        // VB6 里这几条就是 String, 类型该在这里落地, 不在用例里绕。
        if (p == "drive" || p == "path" || p == "pattern" || p == "filename"
            || p == "list") {
            return Vb6Type::String;
        }
    }
    if (ctrlType == FrmControlType::CommonDialog) {
        // 账 #231 (原 D6 / C29-9 那两份名单): 这枚控件没有外观, 读的全是对话框字段。
        // 字符串那七条里 `FontName` 已由通用档答 String (与控件字体同档, 见 #154 那段)，
        // 这里补剩下六条。
        if (p == "filter" || p == "filename" || p == "filetitle" || p == "dialogtitle"
            || p == "initdir" || p == "defaultext") {
            return Vb6Type::String;
        }
        // CancelError 在 VB6 是 Boolean、Color 是 OLE_COLOR —— 这里**照旧答 Long**：
        // 这一刀只做"两份名单合一", 一条答案都不改; 改口径要另开账（否则发码会动）。
        if (p == "flags" || p == "cancelerror" || p == "color" || p == "min"
            || p == "max" || p == "copies" || p == "fontsize") {
            return Vb6Type::Long;
        }
    }
    return Vb6Type::Unknown;
}

// 账 #278 §B109: common 那张表不许认识 FrmControlType（它在 project/，common 不得向上依赖），
// 所以这里把接收者类别折成表用的类别位。**只折表里要用到的那四类**：其余类型一律 0,
// 等价于原来"通用段对所有类型都答"的行为。
static uint32_t formPseudoKindBits(FrmControlType ctrlType) {
    uint32_t bits = vb6c3::FPQ_NONE;
    switch (ctrlType) {
        case FrmControlType::Form:          bits |= vb6c3::FPQ_FORM;          break;
        case FrmControlType::CommonDialog:  bits |= vb6c3::FPQ_COMMON_DIALOG; break;
        case FrmControlType::Shape:         bits |= vb6c3::FPQ_SHAPE;         break;
        case FrmControlType::Line:          bits |= vb6c3::FPQ_LINE;          break;
        // C30-UD-a: UpDown 的 Alignment 是自己那一档（UDS_ALIGNLEFT/RIGHT），
        // 不能落进通用 `vb6_GetAlignment` / `vb6_SetAlignment`。
        case FrmControlType::UpDown:        bits |= vb6c3::FPQ_UPDOWN;        break;
        default: break;
    }
    return bits;
}

std::string CCodeGen::getControlPropReadFn(FrmControlType ctrlType, const std::string& propName) const {
    std::string propLower = propName;
    std::transform(propLower.begin(), propLower.end(), propLower.begin(), ::tolower);

    // P11.8 / P13.x / D6 / C29-9 / C29-SL-p / C29-1a / Fix 056: 属性名 → 读函数**只有一处登记**,
    // 在 src/common/form_pseudo.hpp 的 kFormPseudoRows（账 #278 §B109）。原通用段与窗体臂的
    // 每一行都搬进了那张表, 顺序未改（通用行在前、formOnly 行在后）；三条类型例外
    // （CommonDialog 的 Font* / Shape·Line 的 BorderStyle）由行的 excl 位带出。
    // 为什么搬：语义层判"这名字是不是未声明的标识符"必须问同一个事实, 拷两份就会
    // "产物对、诊断错"（VB3001 噪声 5 条）—— 与 host_pseudo.hpp（账 #159/#219）同族。
    // 表里没登记的仍走下面各类型的 switch 臂（那些是"只有这枚控件答"的成员, 与本表不同一格）。
    {
        std::string tableFn = formPseudoReadFn(propLower, formPseudoKindBits(ctrlType));
        if (!tableFn.empty()) return tableFn;
    }

    switch (ctrlType) {
    case FrmControlType::MaskEdBox:   // C30-3
        // Text/ClipText 都必须走 RTL 掩码引擎：直接 GetWindowText 拿到的是带提示符的
        // 显示串，不是 VB6 语义下的 Text。Mask/PromptChar/PromptInclude 是这枚控件独有。
        if (propLower == "text")          return "vb6_MaskEd_GetText";
        if (propLower == "cliptext")      return "vb6_MaskEd_GetClipText";
        if (propLower == "mask")          return "vb6_MaskEd_GetMask";
        if (propLower == "promptchar")    return "vb6_MaskEd_GetPromptChar";
        if (propLower == "promptinclude") return "vb6_MaskEd_GetPromptInclude";
        if (propLower == "selstart")      return "vb6_GetSelStart";
        if (propLower == "sellength")     return "vb6_GetSelLength";
        if (propLower == "seltext")       return "vb6_GetSelText";
        if (propLower == "visible")       return "vb6_GetControlVisible";
        if (propLower == "enabled")       return "vb6_GetControlEnabled";
        break;
    case FrmControlType::TextBox:
        if (propLower == "text") return "vb6_GetControlText";
        if (propLower == "multiline") return "vb6_GetMultiLine";
        if (propLower == "scrollbars") return "vb6_GetScrollBars";
        if (propLower == "maxlength") return "vb6_GetMaxLength";
        if (propLower == "passwordchar") return "vb6_GetPasswordChar";
        if (propLower == "locked") return "vb6_GetLocked";
        if (propLower == "selstart") return "vb6_GetSelStart";
        if (propLower == "sellength") return "vb6_GetSelLength";
        if (propLower == "seltext") return "vb6_GetSelText";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::ListBox:
    case FrmControlType::ComboBox:
        if (propLower == "text") return "vb6_GetControlText";
        if (propLower == "listcount") return "vb6_GetListCount";
        if (propLower == "listindex") return "vb6_GetListIndex";
        if (propLower == "list") return "vb6_GetListItem";
        if (propLower == "selected") return "vb6_GetSelected";
        if (propLower == "itemdata") return "vb6_GetItemData";
        if (propLower == "newindex") return "vb6_GetNewIndex";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::Frame:
        if (propLower == "caption") return "vb6_GetControlText";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        if (propLower == "align") return "vb6_GetControlAlign";  // P13.5b: 停靠
        break;
    case FrmControlType::Label:
        if (propLower == "caption") return "vb6_GetControlText";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::CommandButton:
        if (propLower == "caption") return "vb6_GetControlText";
        if (propLower == "default") return "vb6_GetDefaultButton";
        if (propLower == "cancel") return "vb6_GetCancelButton";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::CheckBox:
        // CheckBox.Value 在 VB6 是**三态 Integer**(0/1/2)，继续走直接搬 BM_GETCHECK 的那对。
        if (propLower == "value") return "vb6_GetCheckValue";
    case FrmControlType::OptionButton:
        // 账 #128-b: OptionButton.Value 在 VB6 是 Boolean，单开一对专桩。
        // 注意这里**不能再让两个 case 共用同一段函数体**：第一版就把专桩写在了共用体里，
        // 结果连 CheckBox 也一起被换掉，灰态(2)直接被吃掉、变成 -1（本地自检抚住的）。
        if (propLower == "value") return "vb6_GetOptionValue";
        if (propLower == "caption") return "vb6_GetControlText";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::Form:
        // 账 #278 §B109: 窗体臂那 11 行已整张搬进 kFormPseudoRows（formOnly 那批），
        // 上面表里查过就轮不到这里 —— 本臂刻意留空。
        break;
    case FrmControlType::WebBrowser:
        if (propLower == "url" || propLower == "locationurl") return "vb6_WebViewGetUrl";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::HScrollBar:
    case FrmControlType::VScrollBar:
        if (propLower == "value") return "vb6_GetScrollValue";
        if (propLower == "min") return "vb6_GetScrollMin";
        if (propLower == "max") return "vb6_GetScrollMax";
        if (propLower == "largechange") return "vb6_GetLargeChange";
        if (propLower == "smallchange") return "vb6_GetSmallChange";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::Timer:
        if (propLower == "interval") return "vb6_GetTimerInterval";
        if (propLower == "enabled") return "vb6_GetTimerEnabled";
        break;
    case FrmControlType::PictureBox:
        // 账 #192: 画笔光标 CurrentX / CurrentY —— VB6 只在"画得上去"的那几枚上有
        // (Form / PictureBox / UserControl / PropertyPage / Printer)，所以**不给通用行**：
        // 给成通用的话 `List1.CurrentX` 也会答一个 0，那是伪造成功。RTL 早就有
        // (vb6forms_widget_prop.c 的 vb6_GetCurrentX/Y，按 HWND 存窗口属性 VB6_CurrentX/Y)，
        // 后端这一档一直没登记 ⇒ 真工程里 `With Picture1 : .CurrentX` 撞 cgen_expr_with.cpp
        // 那条 tempVar + "." + 成员名 的兜底 = C2039 (实测 ucTreeMaps PropPagFMR.c 74/75/76)。
        if (propLower == "currentx") return "vb6_GetCurrentX";
        if (propLower == "currenty") return "vb6_GetCurrentY";
        // 账 #196: `.hDC` 的出口。跟 currentx 一样**不给通用行** —— VB6 只在画得上去的
        // 那几枚上有 hDC，给成通用 ⇒ `List1.hDC` 也答一个数就是伪造成功。
        if (propLower == "hdc") return "vb6_GetControlHDC";
        // 账 #197: ScaleMode 的读法与几何换算必须是**同一处** (vb6_WindowScaleModeSelf)，
        // 否则程序读到一个数、量出来按另一个数走。写侧成对登记。
        if (propLower == "scalemode") return "vb6_WindowScaleModeSelf";
        if (propLower == "autoredraw") return "vb6_GetAutoRedraw";
        if (propLower == "caption") return "vb6_GetControlText";
        if (propLower == "picture") return "vb6_GetControlPicture";
        /* ucTreeView: PictureBox.Image (AutoRedraw 记忆位图的 StdPicture 视图)。
           读表缺这一枚时 `.Image` 落到 With 块兜底 `_vb6_with_N.Image` —— HWND 是
           结构体指针 ⇒ C2231 (Form6.c:463)。RTL 侧 vb6_FormGetImagePicture 早就有
           (与窗体 Image 同一实现, 按 VB6_AutoRedrawBmp 窗口属性取图)。 */
        if (propLower == "image") return "vb6_FormGetImagePicture";
        if (propLower == "autosize") return "vb6_GetPictureAutoSize";
        if (propLower == "align") return "vb6_GetControlAlign";  // P13.5b: 停靠
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::Image:
        if (propLower == "picture") return "vb6_GetControlPicture";
        if (propLower == "stretch") return "vb6_GetImageStretch";  // P17.2
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::ProgressBar:  // P20-38: 只读回来必须走 RTL,
        // 否则会落到 vb6_ComGetStringProp 这条 COM 占位路径上, 静默返回空。
        if (propLower == "min") return "vb6_GetProgressBarMin";
        if (propLower == "max") return "vb6_GetProgressBarMax";
        if (propLower == "value") return "vb6_GetProgressBarValue";
        if (propLower == "orientation") return "vb6_GetProgressBarOrientation";
        if (propLower == "scrolling") return "vb6_GetProgressBarScrolling";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::ImageList:  // P20-39: 读写都走原生 RTL, 不落 COM 占位路径
        // ListImages 集合本身 (Count/Add/Remove/Item) 由 cgen_util_com.cpp 的
        // resolveComValue 拦截改道, 这里只管 ImageList 自身的标量属性。
        if (propLower == "imagewidth") return "vb6_GetImageListImageWidth";
        if (propLower == "imageheight") return "vb6_GetImageListImageHeight";
        break;
    case FrmControlType::StatusBar:  // P20-40: 只读回来必须走 RTL, 否则落 COM 占位路径静默答错
        if (propLower == "align") return "vb6_StatusBar_GetAlign";
        if (propLower == "style") return "vb6_StatusBar_GetStyle";
        if (propLower == "simpletext") return "vb6_StatusBar_GetSimpleText";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::TabStrip:  // C30-2
        // Tabs / SelectedItem 是**集合/对象**，不走这张表（见 cgen 的集合改道）；
        // 这里只有标量面：四枚 Client* + 三条行为位。
        if (propLower == "clientleft")   return "vb6_TabStrip_GetClientLeft";
        if (propLower == "clienttop")    return "vb6_TabStrip_GetClientTop";
        if (propLower == "clientwidth")  return "vb6_TabStrip_GetClientWidth";
        if (propLower == "clientheight") return "vb6_TabStrip_GetClientHeight";
        if (propLower == "multirow")     return "vb6_TabStrip_GetMultiRow";
        if (propLower == "hottracking")  return "vb6_TabStrip_GetHotTracking";
        if (propLower == "showtips")     return "vb6_TabStrip_GetShowTips";
        break;
    case FrmControlType::SSTab:  // P20-42
        // TabCaption(i)/TabVisible(i) 是**带下标的索引属性**, 不是标量;
        // 走不了这张表 (表的入参只有 propLower), 见 cgen_expr_member_form_builtin.inc 的 TODO。
        if (propLower == "tabs")           return "vb6_SSTab_GetTabs";
        if (propLower == "tab")            return "vb6_SSTab_GetTab";
        if (propLower == "taborientation") return "vb6_SSTab_GetTabOrientation";
        if (propLower == "tabstyle")       return "vb6_SSTab_GetTabStyle";
        if (propLower == "tabsperrow")     return "vb6_SSTab_GetTabsPerRow";
        if (propLower == "wordwrap")       return "vb6_SSTab_GetWordWrap";
        // P20-44 颜色族: 缺省 Ambient(宿主容器)/品红, 见 vb6forms_sstab.c。
        // 不登记会落 vb6_ComGetStringProp 裸 HWND 占位路径静默答空 (同 ProgressBar 纪律)。
        if (propLower == "maskcolor")        return "vb6_SSTab_GetMaskColor";
        if (propLower == "tabbackcolor")     return "vb6_SSTab_GetTabBackColor";
        if (propLower == "tabselbackcolor")  return "vb6_SSTab_GetTabSelBackColor";
        if (propLower == "tabselforecolor")  return "vb6_SSTab_GetTabSelForeColor";
        if (propLower == "visible")        return "vb6_GetControlVisible";
        if (propLower == "enabled")        return "vb6_GetControlEnabled";
        break;
    case FrmControlType::ListView:  // C29-7
        // ListItems / ColumnHeaders 是**集合**, 走成员对象那条路
        // (cgen_util_com.cpp 的 resolveComValue / resolveComMarkerForPack 拦截改道);
        // 这张表只管 ListView 自身的标量属性。
        // 不登记就会落到 vb6_ComGetStringProp 这个 COM 占位路径上静默答空 ——
        // 与 ProgressBar / StatusBar 那两批同一条纪律。
        if (propLower == "view")               return "vb6_ListView_GetView";
        if (propLower == "gridlines")          return "vb6_ListView_GetGridLines";
        if (propLower == "fullrowselect")      return "vb6_ListView_GetFullRowSelect";
        if (propLower == "multiselect")        return "vb6_ListView_GetMultiSelect";
        if (propLower == "checkboxes")         return "vb6_ListView_GetCheckBoxes";
        if (propLower == "hidecolumnheaders")  return "vb6_ListView_GetHideColumnHeaders";
        if (propLower == "allowcolumnreorder") return "vb6_ListView_GetAllowColumnReorder";
        if (propLower == "labeledit")          return "vb6_ListView_GetLabelEdit";
        if (propLower == "sorted")             return "vb6_ListView_GetSorted";
        if (propLower == "sortkey")            return "vb6_ListView_GetSortKey";
        if (propLower == "sortorder")          return "vb6_ListView_GetSortOrder";
        if (propLower == "visible")            return "vb6_GetControlVisible";
        if (propLower == "enabled")            return "vb6_GetControlEnabled";
        break;
    case FrmControlType::OLE:  // C29-OLE: OLEType/Class/SizeMode 等走 RTL (真 OLE 容器)
        // Object 属性走晚绑定 (cgen_util_com.cpp 那条), 不在这张表。
        if (propLower == "class")           return "vb6_OleCon_GetClass";
        if (propLower == "oletype")         return "vb6_OleCon_GetOleType";
        if (propLower == "oletypeallowed")  return "vb6_OleCon_GetOLETypeAllowed";
        if (propLower == "sizemode")        return "vb6_OleCon_GetSizeMode";
        if (propLower == "displayasicon")   return "vb6_OleCon_GetDisplayAsIcon";
        if (propLower == "autoactivate")    return "vb6_OleCon_GetAutoActivate";
        if (propLower == "autoverbmenu")    return "vb6_OleCon_GetAutoVerbMenu";
        if (propLower == "borderstyle")     return "vb6_OleCon_GetBorderStyle";
        if (propLower == "sourcedoc")       return "vb6_OleCon_GetSourceDoc";
        if (propLower == "sourceitem")      return "vb6_OleCon_GetSourceItem";
        if (propLower == "visible")         return "vb6_GetControlVisible";
        if (propLower == "enabled")         return "vb6_GetControlEnabled";
        break;
    case FrmControlType::Winsock:  // C29-WS-a: Winsock2 复刻；状态一切以 RTL 实例表为准
        if (propLower == "protocol")       return "vb6_Ws_GetProtocol";
        if (propLower == "state")          return "vb6_Ws_GetState";
        if (propLower == "localport")      return "vb6_Ws_GetLocalPort";
        if (propLower == "localip")        return "vb6_Ws_GetLocalIP";
        if (propLower == "remotehost")     return "vb6_Ws_GetRemoteHost";
        if (propLower == "remoteport")     return "vb6_Ws_GetRemotePort";
        if (propLower == "bytesreceived")  return "vb6_Ws_GetBytesReceived";
        if (propLower == "bytetransferred") return "vb6_Ws_GetByteTransferred";
        break;
    case FrmControlType::MSComm:  // C30-4: 串口；一切状态在 RTL 实例表里
        if (propLower == "commport")      return "vb6_MsComm_GetCommPort";
        if (propLower == "settings")      return "vb6_MsComm_GetSettings";
        if (propLower == "portopen")      return "vb6_MsComm_GetPortOpen";
        if (propLower == "input")         return "vb6_MsComm_GetInput";
        if (propLower == "inputmode")     return "vb6_MsComm_GetInputMode";
        if (propLower == "inputlen")      return "vb6_MsComm_GetInputLen";
        if (propLower == "inbuffercount") return "vb6_MsComm_GetInBufferCount";
        if (propLower == "outbuffercount") return "vb6_MsComm_GetOutBufferCount";
        if (propLower == "inbuffersize")  return "vb6_MsComm_GetInBufferSize";
        if (propLower == "outbuffersize") return "vb6_MsComm_GetOutBufferSize";
        if (propLower == "rthreshold")    return "vb6_MsComm_GetRThreshold";
        if (propLower == "sthreshold")    return "vb6_MsComm_GetSThreshold";
        if (propLower == "handshaking")   return "vb6_MsComm_GetHandshaking";
        if (propLower == "dtrenable")     return "vb6_MsComm_GetDTREnable";
        if (propLower == "rtsenable")     return "vb6_MsComm_GetRTSEnable";
        if (propLower == "eofenable")     return "vb6_MsComm_GetEOFEnable";
        if (propLower == "commevent")     return "vb6_MsComm_GetCommEvent";
        if (propLower == "ctsholding")    return "vb6_MsComm_GetCTSHolding";
        if (propLower == "dsrholding")    return "vb6_MsComm_GetDSRHolding";
        if (propLower == "cdholding")     return "vb6_MsComm_GetCDHolding";
        // C3 扩展（不是 VB6 属性）：夹具证人 —— 真写/真收了多少字节。
        if (propLower == "txtotal")       return "vb6_MsComm_GetTxTotal";
        if (propLower == "rxtotal")       return "vb6_MsComm_GetRxTotal";
        // C30-4b：超时 / 任意设备名 / 端口枚举 / 二进制（VB6 没有的几件）
        if (propLower == "commdevice")    return "vb6_MsComm_GetCommDevice";
        if (propLower == "readintervaltimeout") return "vb6_MsComm_GetReadIntervalTimeout";
        if (propLower == "readtotaltimeout")    return "vb6_MsComm_GetReadTotalTimeout";
        if (propLower == "writetimeout")   return "vb6_MsComm_GetWriteTimeout";
        if (propLower == "portcount")      return "vb6_MsComm_PortCount";
        if (propLower == "portlist")       return "vb6_MsComm_GetPortList";
        if (propLower == "inputbytes")     return "vb6_MsComm_GetInputBytes";
        break;
    case FrmControlType::Adodc:  // C30-5: ACE 底座（ODBC 层）；属性全在 RTL 实例表里
        if (propLower == "connectionstring") return "vb6_Adodc_GetConnectionString";
        if (propLower == "recordsource")  return "vb6_Adodc_GetRecordSource";
        if (propLower == "commandtype")   return "vb6_Adodc_GetCommandType";
        if (propLower == "cursorlocation") return "vb6_Adodc_GetCursorLocation";
        if (propLower == "cursortype")    return "vb6_Adodc_GetCursorType";
        if (propLower == "locktype")      return "vb6_Adodc_GetLockType";
        if (propLower == "mode")          return "vb6_Adodc_GetMode";
        if (propLower == "maxrecords")    return "vb6_Adodc_GetMaxRecords";
        if (propLower == "bofaction")     return "vb6_Adodc_GetBOFAction";
        if (propLower == "eofaction")     return "vb6_Adodc_GetEOFAction";
        if (propLower == "lastrowcount")  return "vb6_Adodc_GetLastRowCount";
        if (propLower == "opened")        return "vb6_Adodc_GetOpened";
        // Recordset：真 IDispatch（引擎的 VB6_MEMK_RS），晚绑定链原样走
        if (propLower == "recordset")     return "vb6_Adodc_Recordset";
        break;
    case FrmControlType::DataGrid:  // C30-5: 自绘表格
        if (propLower == "datasource")    return "vb6_DataGrid_GetDataSourceName";
        if (propLower == "rowcount")      return "vb6_DataGrid_GetRowCount";
        if (propLower == "colcount")      return "vb6_DataGrid_GetColCount";
        if (propLower == "row")           return "vb6_DataGrid_GetRow";
        if (propLower == "col")           return "vb6_DataGrid_GetCol";
        if (propLower == "text")          return "vb6_DataGrid_GetText";
        if (propLower == "columnheaders") return "vb6_DataGrid_GetColumnHeaders";
        if (propLower == "allowaddnew")   return "vb6_DataGrid_GetAllowAddNew";
        if (propLower == "allowupdate")   return "vb6_DataGrid_GetAllowUpdate";
        if (propLower == "allowdelete")   return "vb6_DataGrid_GetAllowDelete";
        if (propLower == "rowheight")     return "vb6_DataGrid_GetRowHeight";
        if (propLower == "defcolwidth")   return "vb6_DataGrid_GetDefColWidth";
        break;
    case FrmControlType::MshFlexGrid:  // C30-6: 字符串矩阵表格（get）
        // 带下标的 TextMatrix(r,c) / TextArray(i) / ColWidth(i) 三条走
        // cgen_expr_call_callee_member.inc 的索引属性分支，不在这张表里。
        if (propLower == "text")          return "vb6_MshFlexGrid_GetText";
        if (propLower == "rows")          return "vb6_MshFlexGrid_GetRows";
        if (propLower == "cols")          return "vb6_MshFlexGrid_GetCols";
        if (propLower == "fixedrows")     return "vb6_MshFlexGrid_GetFixedRows";
        if (propLower == "fixedcols")     return "vb6_MshFlexGrid_GetFixedCols";
        if (propLower == "row")           return "vb6_MshFlexGrid_GetRow";
        if (propLower == "col")           return "vb6_MshFlexGrid_GetCol";
        if (propLower == "rowsel")        return "vb6_MshFlexGrid_GetRowSel";
        if (propLower == "colsel")        return "vb6_MshFlexGrid_GetColSel";
        if (propLower == "toprow")        return "vb6_MshFlexGrid_GetTopRow";
        if (propLower == "leftcol")       return "vb6_MshFlexGrid_GetLeftCol";
        if (propLower == "gridlines")     return "vb6_MshFlexGrid_GetGridLines";
        if (propLower == "scrollbars")    return "vb6_MshFlexGrid_GetScrollBars";
        if (propLower == "selectionmode") return "vb6_MshFlexGrid_GetSelectionMode";
        if (propLower == "fillstyle")     return "vb6_MshFlexGrid_GetFillStyle";
        if (propLower == "wordwrap")      return "vb6_MshFlexGrid_GetWordWrap";
        if (propLower == "editable")      return "vb6_MshFlexGrid_GetEditable";
        if (propLower == "focusrect")     return "vb6_MshFlexGrid_GetFocusRect";
        if (propLower == "allowuserresizing") return "vb6_MshFlexGrid_GetAllowUserResizing";
        if (propLower == "mergecells")    return "vb6_MshFlexGrid_GetMergeCells";
        if (propLower == "redraw")        return "vb6_MshFlexGrid_GetRedraw";
        if (propLower == "backcolor")     return "vb6_MshFlexGrid_GetBackColor";
        if (propLower == "forecolor")     return "vb6_MshFlexGrid_GetForeColor";
        if (propLower == "fillcolor")     return "vb6_MshFlexGrid_GetFillColor";
        if (propLower == "cellbackcolor") return "vb6_MshFlexGrid_GetCellBackColor";
        if (propLower == "cellforecolor") return "vb6_MshFlexGrid_GetCellForeColor";
        if (propLower == "mouserow")      return "vb6_MshFlexGrid_GetMouseRow";
        if (propLower == "mousecol")      return "vb6_MshFlexGrid_GetMouseCol";
        if (propLower == "cellheight")    return "vb6_MshFlexGrid_GetCellHeight";
        if (propLower == "cellwidth")     return "vb6_MshFlexGrid_GetCellWidth";
        if (propLower == "cellleft")      return "vb6_MshFlexGrid_GetCellLeft";
        if (propLower == "celltop")       return "vb6_MshFlexGrid_GetCellTop";
        if (propLower == "formatstring")  return "vb6_MshFlexGrid_GetFormatString";
        if (propLower == "datasource")    return "vb6_MshFlexGrid_GetDataSource";
        if (propLower == "rowheightmin")  return "vb6_MshFlexGrid_GetRowHeightMin";
        if (propLower == "sort")          return "vb6_MshFlexGrid_GetSort";
        break;
    case FrmControlType::Data:  // C29-Data: ODBC 后端
        if (propLower == "recordset")    return "vb6_Data_RecordsetObj";  /* 真 IDispatch, 链走晚绑定 */
        if (propLower == "databasename") return "vb6_Data_GetDatabaseName";
        if (propLower == "recordsource") return "vb6_Data_GetRecordSource";
        if (propLower == "connect")      return "vb6_Data_GetConnect";
        if (propLower == "bof")          return "vb6_Data_BOF";
        if (propLower == "eof")          return "vb6_Data_EOF";
        if (propLower == "recordcount")  return "vb6_Data_RecordCount";
        break;
    case FrmControlType::Menu:  // P20-36
        if (propLower == "caption") return "vb6_GetMenuCaption";
        if (propLower == "checked") return "vb6_GetMenuChecked";
        if (propLower == "enabled") return "vb6_GetMenuEnabled";
        if (propLower == "visible") return "vb6_GetMenuVisible";
        break;
    // C29-1b: 文件系统三控件的专有成员。列表成员走 ListBox/ComboBox 那一族 helper
    // (它们内部按窗口类分流 LB_* / CB_*), 所以 Drive 的组合框与 Dir/File 的列表框
    // 共用同一批读函数, 不需要为这三类另写一套。
    case FrmControlType::DriveListBox:
        if (propLower == "drive") return "vb6_DriveListBoxDrive";
        if (propLower == "text") return "vb6_GetControlText";
        if (propLower == "listcount") return "vb6_GetListCount";
        if (propLower == "listindex") return "vb6_GetListIndex";
        if (propLower == "list") return "vb6_GetListItem";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::DirListBox:
        if (propLower == "path") return "vb6_DirListBoxPath";
        if (propLower == "listcount") return "vb6_GetListCount";
        if (propLower == "listindex") return "vb6_GetListIndex";
        if (propLower == "list") return "vb6_GetListItem";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::FileListBox:
        if (propLower == "path") return "vb6_FileListBoxPath";
        if (propLower == "pattern") return "vb6_FileListBoxPattern";
        if (propLower == "filename") return "vb6_FileListBoxFileName";
        if (propLower == "listcount") return "vb6_GetListCount";
        if (propLower == "listindex") return "vb6_GetListIndex";
        if (propLower == "list") return "vb6_GetListItem";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    // C29-8a: TreeView 的标量属性面。之前这类零格 ⇒ 全部落到"未知属性"的通用兜底
    // vb6_ComGetStringProp(裸 HWND, L"CheckBoxes") —— 读回空串、写进去静默丢 (029 §九)。
    // 这四条的真值在窗口样式位上 (RTL 读 GWL_STYLE)，Indentation 的缇值存窗口属性。
    // Style / LabelEdit / Sorted / PathSeparator / Nodes 一族留 C29-8b (等成员对象机制)。
    case FrmControlType::TreeView:
        if (propLower == "linestyle") return "vb6_TreeView_GetLineStyle";
        if (propLower == "indentation") return "vb6_TreeView_GetIndentation";
        if (propLower == "iconsize") return "vb6_TreeView_GetIconSize";
        if (propLower == "checkboxes") return "vb6_TreeView_GetCheckBoxes";
        if (propLower == "hottracking") return "vb6_TreeView_GetHotTracking";
        if (propLower == "hideselection") return "vb6_TreeView_GetHideSelection";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    // C29-DT-a: DTPicker 的标量属性面 (原生 SysDateTimePick32)。登记之前这枚控件走的是
    // "第三方 OCX 按 COM 后期绑定"那一组 (MSComCtl2 在工程里没引用类型库时连符号都查不到)
    // ⇒ 属性读回空、写进去静默丢。这里除 CustomFormat 外全按 int32_t 走，与类型登记表同批。
    case FrmControlType::DTPicker:
        if (propLower == "format") return "vb6_DTP_GetFormat";
        if (propLower == "customformat") return "vb6_DTP_GetCustomFormat";
        if (propLower == "checkbox") return "vb6_DTP_GetCheckBox";
        if (propLower == "updown") return "vb6_DTP_GetUpDown";
        if (propLower == "calendarbackcolor") return "vb6_DTP_GetCalendarBackColor";
        if (propLower == "calendarforecolor") return "vb6_DTP_GetCalendarForeColor";
        if (propLower == "calendartrailingforecolor") return "vb6_DTP_GetCalendarTrailingForeColor";
        if (propLower == "calendartitlebackcolor") return "vb6_DTP_GetCalendarTitleBackColor";
        if (propLower == "calendartitleforecolor") return "vb6_DTP_GetCalendarTitleForeColor";
        // C3 扩展（只读）：控件自己算的"装得下当前格式"宽度，判据用它把"格式真选中了吗"
        // 从自家人读数换成控件侧读数。VB6 没有这条，写侧刻意不登记。
        if (propLower == "idealwidth") return "vb6_DTP_IdealWidth";
        // C29-DT-b 的读侧（Date 三条 + 扩展 HasDate）。
        if (propLower == "value") return "vb6_DTP_GetValue";
        if (propLower == "hasdate") return "vb6_DTP_HasDate";
        if (propLower == "mindate") return "vb6_DTP_GetMinDate";
        if (propLower == "maxdate") return "vb6_DTP_GetMaxDate";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    // C29-MV-a: MonthView 读侧。样式三位 + MaxSelCount + 五色，全部直读控件；
    // 三条 C3 扩展读数（MinReqWidth / MinReqHeight / MonthCount）是**控件侧**证据，
    // 判据靠它们把"样式位写进去了"升级成"控件真按那位在画"。
    case FrmControlType::MonthView:
        if (propLower == "multiselect") return "vb6_MV_GetMultiSelect";
        if (propLower == "showweeknumbers") return "vb6_MV_GetShowWeekNumbers";
        if (propLower == "showtoday") return "vb6_MV_GetShowToday";
        if (propLower == "maxselcount") return "vb6_MV_GetMaxSelCount";
        if (propLower == "backcolor") return "vb6_MV_GetBackColor";
        if (propLower == "forecolor") return "vb6_MV_GetForeColor";
        if (propLower == "titlebackcolor") return "vb6_MV_GetTitleBackColor";
        if (propLower == "titleforecolor") return "vb6_MV_GetTitleForeColor";
        if (propLower == "trailingforecolor") return "vb6_MV_GetTrailingForeColor";
        if (propLower == "monthbackcolor") return "vb6_MV_GetMonthBackColor";
        if (propLower == "minreqwidth") return "vb6_MV_MinReqWidth";
        if (propLower == "minreqheight") return "vb6_MV_MinReqHeight";
        if (propLower == "monthcount") return "vb6_MV_GetMonthCount";
        // C29-MV-b 的读侧（Date 三条）。
        if (propLower == "value") return "vb6_MV_GetValue";
        if (propLower == "selstart") return "vb6_MV_GetSelStart";
        if (propLower == "selend") return "vb6_MV_GetSelEnd";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    // C29-SL-a: Slider 读侧（原生 msctls_trackbar32）。登记之前这枚控件整个走
    // "第三方 OCX 按 COM 后期绑定"那一组，实测**连窗口都没建**（frm_parser_util 那张
    // 按子串匹配的表里没有 "slider" ⇒ Unknown ⇒ 创建流程跳过），属性读回全空、写进去静默丢。
    // 两条 C3 扩展读数（TravelIsVert / TickPresent）是**控件侧证人**，判据靠它们把
    // "样式位写进去了"升级成"控件真按那一档在走"（channel 矩形那条是假证人，实测过）。
    // C30-UD-a: UpDown 读侧（原生 msctls_updown32）。全部直问控件：range/pos/accel
    // 都有回读消息，行为面四条就读窗口的样式位 —— 没有一条是自存的。
    case FrmControlType::UpDown:
        if (propLower == "min") return "vb6_UpDown_GetMin";
        if (propLower == "max") return "vb6_UpDown_GetMax";
        if (propLower == "value") return "vb6_UpDown_GetValue";
        if (propLower == "increment") return "vb6_UpDown_GetIncrement";
        if (propLower == "wrap") return "vb6_UpDown_GetWrap";
        if (propLower == "alignment") return "vb6_UpDown_GetAlignment";
        if (propLower == "orientation") return "vb6_UpDown_GetOrientation";
        if (propLower == "syncbuddy") return "vb6_UpDown_GetSyncBuddy";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    case FrmControlType::Slider:
        if (propLower == "orientation") return "vb6_Slider_GetOrientation";
        if (propLower == "tickfrequency") return "vb6_Slider_GetTickFrequency";
        if (propLower == "travelisvert") return "vb6_Slider_TravelIsVert";
        if (propLower == "tickpresent") return "vb6_Slider_TickPresent";
        // C29-SL-b 的读侧（值面）。全部直问控件，只有 TickFrequency 因原生问不出而自存。
        if (propLower == "min") return "vb6_Slider_GetMin";
        if (propLower == "max") return "vb6_Slider_GetMax";
        if (propLower == "value") return "vb6_Slider_GetValue";
        if (propLower == "smallchange") return "vb6_Slider_GetSmallChange";
        if (propLower == "largechange") return "vb6_Slider_GetLargeChange";
        if (propLower == "selectrange") return "vb6_Slider_GetSelectRange";
        if (propLower == "selstart") return "vb6_Slider_GetSelStart";
        if (propLower == "selend") return "vb6_Slider_GetSelEnd";
        // C29-SL-h: TickStyle 四档 + GetNumTicks（VB6 只读那一面），加一条判据证人 ChannelTop。
        if (propLower == "tickstyle") return "vb6_Slider_GetTickStyle";
        if (propLower == "getnumticks") return "vb6_intrSlider_GetNumTicks";
        if (propLower == "channeltop") return "vb6_Slider_ChannelTop";
        // 账 #148 的判据证人（**不是 VB6 属性**）：tooltip 宿主里到底有没有这枚控件的工具。
        if (propLower == "tooltipregistered") return "vb6_ToolTipRegistered";
        // C29-SL-k: VB6 的 Slider.Text 是**气泡里那句串**（类型库 VT_BSTR，0x0010），
        // 不是窗口标题 —— 这一格把它从通用的 vb6_GetControlText 那支抢过来自己实现。
        // 三条 Bubble* 是判据证人（C3 扩展）：问的是 tooltip 宿主自己，不是我们的窗口属性。
        if (propLower == "text") return "vb6_Slider_GetText";
        if (propLower == "textposition") return "vb6_Slider_GetTextPosition";
        if (propLower == "bubblevisible") return "vb6_Slider_BubbleVisible";
        if (propLower == "bubbletop") return "vb6_Slider_BubbleTop";
        if (propLower == "bubbletext") return "vb6_Slider_BubbleText";
        // C29-SL-i: VB6 选区那一对的第二条（SelStart + SelLength，类型库 dispid 0x0007/0x0008）。
        if (propLower == "sellength") return "vb6_Slider_GetSelLength";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    // C29-RT-a: RichTextBox 读侧。Text 走通用那条（与 TextBox 同一格 vb6_GetControlText），
    // BorderStyle 也走通用那条窗口边框读数 —— 原生这枚控件的边框在 WS_EX_CLIENTEDGE 上，
    // 通用 getter 认它（实测见夹具 RT5/RT6 两条）。
    case FrmControlType::RichTextBox:
        if (propLower == "text") return "vb6_GetControlText";
        if (propLower == "selstart") return "vb6_RTB_GetSelStart";
        if (propLower == "sellength") return "vb6_RTB_GetSelLength";
        if (propLower == "seltext") return "vb6_RTB_GetSelText";
        // C29-RT-c: 整串 RTF = EM_STREAMOUT + SF_RTF。
        if (propLower == "textrtf") return "vb6_RTB_GetTextRTF";
        if (propLower == "readonly") return "vb6_RTB_GetReadOnly";
        if (propLower == "maxlength") return "vb6_RTB_GetMaxLength";
        if (propLower == "scrollbars") return "vb6_RTB_GetScrollBars";
        if (propLower == "wordwrap") return "vb6_RTB_GetWordWrap";
        // C3 扩展读数（不是 VB6 属性）：控件自己的滚动量程 —— "滚动条活没活"的硬证人。
        if (propLower == "vscrollrange") return "vb6_RTB_GetVScrollRange";
        if (propLower == "hscrollrange") return "vb6_RTB_GetHScrollRange";
        // C29-RT-b 的读侧（格式面九条 + 字体名/字号）。混合态一律回"没有"那一头（False / 0 / 空串），
        // 判据靠"只涂一半再问跨界"钉住它（RT33-RT48）。
        if (propLower == "selbold") return "vb6_RTB_GetSelBold";
        if (propLower == "selitalic") return "vb6_RTB_GetSelItalic";
        if (propLower == "selunderline") return "vb6_RTB_GetSelUnderline";
        if (propLower == "selstrikethru") return "vb6_RTB_GetSelStrikethru";
        if (propLower == "selcolor") return "vb6_RTB_GetSelColor";
        if (propLower == "selfontname") return "vb6_RTB_GetSelFontName";
        if (propLower == "selfontsize") return "vb6_RTB_GetSelFontSize";
        if (propLower == "selalignment") return "vb6_RTB_GetSelAlignment";
        if (propLower == "selindent") return "vb6_RTB_GetSelIndent";
        if (propLower == "selrightindent") return "vb6_RTB_GetSelRightIndent";
        if (propLower == "selhangingindent") return "vb6_RTB_GetSelHangingIndent";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    // C29-5a: Toolbar 的标量属性面。改之前这枚控件连窗口都没有 (被"ImageList || Toolbar
    // 走 CoCreateInstance"那一组扣住)，读一个 tb1.Visible 就是 C2065: vb6_hwnd_tb1 未声明。
    // ShowTips / TextStyle / AllowCustomize 的真值在 GWL_STYLE 上，Align 存窗口属性。
    case FrmControlType::Toolbar:
        if (propLower == "showtips") return "vb6_Toolbar_GetShowTips";
        if (propLower == "textstyle") return "vb6_Toolbar_GetTextStyle";
        if (propLower == "allowcustomize") return "vb6_Toolbar_GetAllowCustomize";
        if (propLower == "align") return "vb6_Toolbar_GetAlign";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    // D6 / C29-9: CommonDialog —— 属性袋挂在那枚自注册的不可见窗口上，
    // 读写口全部走 vb6_Cd* （原生 comdlg32，不再经 MSComDlg.OCX）。
    case FrmControlType::CommonDialog:
        if (propLower == "filter") return "vb6_CdGetFilter";
        if (propLower == "filename") return "vb6_CdGetFileName";
        if (propLower == "filetitle") return "vb6_CdGetFileTitle";
        if (propLower == "dialogtitle") return "vb6_CdGetDialogTitle";
        if (propLower == "initdir") return "vb6_CdGetInitDir";
        if (propLower == "defaultext") return "vb6_CdGetDefaultExt";
        if (propLower == "fontname") return "vb6_CdGetFontName";
        if (propLower == "flags") return "vb6_CdGetFlags";
        if (propLower == "cancelerror") return "vb6_CdGetCancelError";
        if (propLower == "color") return "vb6_CdGetColor";
        if (propLower == "min") return "vb6_CdGetMin";
        if (propLower == "max") return "vb6_CdGetMax";
        if (propLower == "copies") return "vb6_CdGetCopies";
        if (propLower == "fontsize") return "vb6_CdGetFontSize";
        if (propLower == "filterindex") return "vb6_CdGetFilterIndex";
        break;
    case FrmControlType::Shape:  // P20-35
        if (propLower == "shape") return "vb6_GetShapeType";        if (propLower == "borderwidth") return "vb6_GetShapeBorderWidth";
        if (propLower == "borderstyle") return "vb6_GetShapeBorderStyle";
        if (propLower == "fillstyle") return "vb6_GetShapeFillStyle";
        if (propLower == "bordercolor") return "vb6_GetShapeBorderColor";
        if (propLower == "fillcolor") return "vb6_GetShapeFillColor";
        if (propLower == "visible") return "vb6_GetControlVisible";
        break;
    case FrmControlType::Line:  // P20-35
        if (propLower == "x1") return "vb6_GetLineX1";
        if (propLower == "y1") return "vb6_GetLineY1";
        if (propLower == "x2") return "vb6_GetLineX2";
        if (propLower == "y2") return "vb6_GetLineY2";
        if (propLower == "borderwidth") return "vb6_GetLineBorderWidth";
        if (propLower == "borderstyle") return "vb6_GetLineBorderStyle";
        if (propLower == "bordercolor") return "vb6_GetLineColor";
        if (propLower == "visible") return "vb6_GetControlVisible";
        break;
    default:
        // 所有可见控件通用属性
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        break;
    }
    return "";  // 未知属性
}


std::string CCodeGen::getControlPropWriteFn(FrmControlType ctrlType, const std::string& propName) const {
    std::string propLower = propName;
    std::transform(propLower.begin(), propLower.end(), propLower.begin(), ::tolower);

        // P11.8: Common properties for all visible controls (checked before switch)
    if (propLower == "left") return "vb6_SetControlLeft";
    if (propLower == "top") return "vb6_SetControlTop";
    if (propLower == "width") return "vb6_SetControlWidth";
    if (propLower == "height") return "vb6_SetControlHeight";
    // P13.1: Font properties (all visible controls with text)
    // D6 / C29-9: 同上 —— CommonDialog 的 Font* 走它自己的属性袋（写侧）。
    if (propLower == "fontname" && ctrlType != FrmControlType::CommonDialog) return "vb6_SetControlFontName";
    if (propLower == "fontsize" && ctrlType != FrmControlType::CommonDialog) return "vb6_SetControlFontSize";
    if (propLower == "fontbold") return "vb6_SetControlFontBold";
    if (propLower == "fontitalic") return "vb6_SetControlFontItalic";
    if (propLower == "fontunderline") return "vb6_SetControlFontUnderline";
    if (propLower == "fontstrikethrough") return "vb6_SetControlFontStrikethrough";
    // P13.2: Color properties (all visible controls)
    if (propLower == "forecolor") return "vb6_SetControlForeColor";
    if (propLower == "backcolor") return "vb6_SetControlBackColor";
    // P13.5: Alignment (all text controls)
    // C30-UD-a: UpDown 除外 —— 它的 Alignment 是 UDS_ALIGNLEFT/RIGHT 两位（自己那格），
    // 不能被这条通用 ES_* 写路径抢走。实测：不除它，写 0 后读回 1，且它会拿旧样式重写
    // GWL_STYLE，把刚立上的 UDS_WRAP 一起抹掉（夹具 UD10 就是那条红的证人）。
    if (propLower == "alignment" && ctrlType != FrmControlType::UpDown) return "vb6_SetAlignment";
    // P13.6: TabIndex/TabStop (all visible controls)
    if (propLower == "tabindex") return "vb6_SetTabIndex";
    if (propLower == "tabstop") return "vb6_SetTabStop";
    if (propLower == "causesvalidation") return "vb6_SetCausesValidation";
    // P13.8: ToolTipText (all visible controls)
    if (propLower == "tooltiptext") return "vb6_SetToolTipText";
    // P13.9: Tag (all controls)
    if (propLower == "tag") return "vb6_SetControlTag";
    // P13.7: MousePointer/MouseIcon (all visible controls)
    if (propLower == "mousepointer") return "vb6_SetMousePointer";
    if (propLower == "mouseicon") return "vb6_SetMouseIcon";
    // P13.10: BorderStyle (all visible controls)
    // C29-1a: Shape/Line 的 BorderStyle 是"画笔线型"(0..6), 与窗口边框样式(0/1)同名
    // 不同物 —— 让这两个类型走下面各自的 case, 否则读回来的永远是窗口那套。
    if (propLower == "borderstyle" && ctrlType != FrmControlType::Shape
        && ctrlType != FrmControlType::Line) return "vb6_SetBorderStyle";

    switch (ctrlType) {
    case FrmControlType::MaskEdBox:   // C30-3
        if (propLower == "text")          return "vb6_MaskEd_SetText";
        if (propLower == "cliptext")      return "vb6_MaskEd_SetClipText";
        if (propLower == "mask")          return "vb6_MaskEd_SetMask";
        if (propLower == "promptchar")    return "vb6_MaskEd_SetPromptChar";
        if (propLower == "promptinclude") return "vb6_MaskEd_SetPromptInclude";
        if (propLower == "selstart")      return "vb6_SetSelStart";
        if (propLower == "sellength")     return "vb6_SetSelLength";
        if (propLower == "seltext")       return "vb6_SetSelText";
        if (propLower == "visible")       return "vb6_SetControlVisible";
        if (propLower == "enabled")       return "vb6_SetControlEnabled";
        break;
    case FrmControlType::TextBox:
        if (propLower == "text") return "vb6_SetControlText";
        if (propLower == "multiline") return "vb6_SetMultiLine";
        if (propLower == "scrollbars") return "vb6_SetScrollBars";
        if (propLower == "maxlength") return "vb6_SetMaxLength";
        if (propLower == "passwordchar") return "vb6_SetPasswordChar";
        if (propLower == "locked") return "vb6_SetLocked";
        if (propLower == "selstart") return "vb6_SetSelStart";
        if (propLower == "sellength") return "vb6_SetSelLength";
        if (propLower == "seltext") return "vb6_SetSelText";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::ListBox:
    case FrmControlType::ComboBox:
        if (propLower == "text") return "vb6_SetControlText";
        if (propLower == "listindex") return "vb6_SetListIndex";
        if (propLower == "list") return "vb6_SetListItem";
        if (propLower == "selected") return "vb6_SetSelected";
        if (propLower == "itemdata") return "vb6_SetItemData";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::Frame:
        if (propLower == "caption") return "vb6_SetControlText";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        if (propLower == "align") return "vb6_SetControlAlign";  // P13.5b: 停靠
        break;
    case FrmControlType::Label:
        if (propLower == "caption") return "vb6_SetControlText";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::CommandButton:
        if (propLower == "caption") return "vb6_SetControlText";
        if (propLower == "default") return "vb6_SetDefaultButton";
        if (propLower == "cancel") return "vb6_SetCancelButton";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::CheckBox:
        if (propLower == "value") return "vb6_SetCheckValue";
    case FrmControlType::OptionButton:
        if (propLower == "value") return "vb6_SetOptionValue";
        if (propLower == "caption") return "vb6_SetControlText";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::Form:
        if (propLower == "caption") return "vb6_SetControlText";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        if (propLower == "scalemode") return "vb6_SetScaleMode";  // 账 #197: Me.ScaleMode 写得动
        // 读侧 form_pseudo 表早就有 windowstate → vb6_GetWindowState；
        // 写侧没登记 ⇒ `FormMain.WindowState = vbNormal` 落 Module.var 回退，
        // 发出裸 `WindowState = 0;`（整编实测 C2065 ×4：Interface.c 2、
        // Tools_ScreenVideo.c 1、Tools_ScreenVideoPrefs.c 1）。
        if (propLower == "windowstate") return "vb6_SetWindowState";
        // OLEDropMode 是窗体/控件的 VB6 属性, 但写侧从来没登记过 ⇒
        // `FormMain.OLEDropMode = vbOLEDropNone` 落 Module.var 回退, 发出裸
        // `OLEDropMode = 0;`(整编实测 C2065 ×2: Processor.c 1176/1185)。
        // 配对出口走 RTL 的 OLE 放置目标开关 (0 → Revoke, 非 0 → Register)。
        if (propLower == "oledropmode") return "vb6_SetOLEDropMode";
        // 账 #233: 与上面读侧成对 —— 赋值语句发 C 时**只问这张表**, 没登记就把读函数当左值。
        if (propLower == "currentx") return "vb6_SetCurrentX";
        if (propLower == "currenty") return "vb6_SetCurrentY";
        if (propLower == "drawwidth") return "vb6_Form_DrawSetWidth";
        break;
    case FrmControlType::WebBrowser:
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::HScrollBar:
    case FrmControlType::VScrollBar:
        if (propLower == "value") return "vb6_SetScrollValue";
        if (propLower == "min") return "vb6_SetScrollMin";
        if (propLower == "max") return "vb6_SetScrollMax";
        if (propLower == "largechange") return "vb6_SetLargeChange";
        if (propLower == "smallchange") return "vb6_SetSmallChange";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::ProgressBar:
        if (propLower == "min") return "vb6_SetProgressBarMin";
        if (propLower == "max") return "vb6_SetProgressBarMax";
        if (propLower == "value") return "vb6_SetProgressBarValue";
        if (propLower == "orientation") return "vb6_SetProgressBarOrientation";
        if (propLower == "scrolling") return "vb6_SetProgressBarScrolling";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::ImageList:  // P20-39
        if (propLower == "imagewidth") return "vb6_SetImageListImageWidth";
        if (propLower == "imageheight") return "vb6_SetImageListImageHeight";
        break;
    case FrmControlType::StatusBar:  // P20-40
        if (propLower == "align") return "vb6_StatusBar_SetAlign";
        if (propLower == "style") return "vb6_StatusBar_SetStyle";
        if (propLower == "simpletext") return "vb6_StatusBar_SetSimpleText";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::TabStrip:  // C30-2
        if (propLower == "multirow")    return "vb6_TabStrip_SetMultiRow";
        if (propLower == "hottracking") return "vb6_TabStrip_SetHotTracking";
        if (propLower == "showtips")    return "vb6_TabStrip_SetShowTips";
        break;
    case FrmControlType::SSTab:  // P20-42
        if (propLower == "tabs")           return "vb6_SSTab_SetTabs";
        if (propLower == "tab")            return "vb6_SSTab_SetTab";
        if (propLower == "taborientation") return "vb6_SSTab_SetTabOrientation";
        if (propLower == "tabstyle")       return "vb6_SSTab_SetTabStyle";
        if (propLower == "tabsperrow")     return "vb6_SSTab_SetTabsPerRow";
        if (propLower == "wordwrap")       return "vb6_SSTab_SetWordWrap";
        // P20-44 颜色族 (写): 不登记会落 vb6_ComSetProp 裸 HWND 泛化写 → 静默丢。
        if (propLower == "maskcolor")        return "vb6_SSTab_SetMaskColor";
        if (propLower == "tabbackcolor")     return "vb6_SSTab_SetTabBackColor";
        if (propLower == "tabselbackcolor")  return "vb6_SSTab_SetTabSelBackColor";
        if (propLower == "tabselforecolor")  return "vb6_SSTab_SetTabSelForeColor";
        if (propLower == "visible")        return "vb6_SetControlVisible";
        if (propLower == "enabled")        return "vb6_SetControlEnabled";
        break;
    case FrmControlType::ListView:  // C29-7
        if (propLower == "view")               return "vb6_ListView_SetView";
        if (propLower == "gridlines")          return "vb6_ListView_SetGridLines";
        if (propLower == "fullrowselect")      return "vb6_ListView_SetFullRowSelect";
        if (propLower == "multiselect")        return "vb6_ListView_SetMultiSelect";
        if (propLower == "checkboxes")         return "vb6_ListView_SetCheckBoxes";
        if (propLower == "hidecolumnheaders")  return "vb6_ListView_SetHideColumnHeaders";
        if (propLower == "allowcolumnreorder") return "vb6_ListView_SetAllowColumnReorder";
        if (propLower == "labeledit")          return "vb6_ListView_SetLabelEdit";
        if (propLower == "sorted")             return "vb6_ListView_SetSorted";
        if (propLower == "sortkey")            return "vb6_ListView_SetSortKey";
        if (propLower == "sortorder")          return "vb6_ListView_SetSortOrder";
        if (propLower == "visible")            return "vb6_SetControlVisible";
        if (propLower == "enabled")            return "vb6_SetControlEnabled";
        break;
    case FrmControlType::OLE:  // C29-OLE 写表
        if (propLower == "oletypeallowed") return "vb6_OleCon_SetOLETypeAllowed";
        if (propLower == "sizemode")       return "vb6_OleCon_SetSizeMode";
        if (propLower == "displayasicon")  return "vb6_OleCon_SetDisplayAsIcon";
        if (propLower == "autoactivate")   return "vb6_OleCon_SetAutoActivate";
        if (propLower == "autoverbmenu")   return "vb6_OleCon_SetAutoVerbMenu";
        if (propLower == "borderstyle")    return "vb6_OleCon_SetBorderStyle";
        break;
    case FrmControlType::Timer:
        if (propLower == "interval") return "vb6_SetTimerInterval";
        if (propLower == "enabled") return "vb6_SetTimerEnabled";
        break;
    case FrmControlType::PictureBox:
        // 账 #192: 与读表成对（只给读侧的话 `.CurrentX = 0` 会落到 HWND 结构体字段上）。
if (propLower == "currentx") return "vb6_SetCurrentX";
        if (propLower == "currenty") return "vb6_SetCurrentY";
        if (propLower == "scalemode") return "vb6_SetScaleMode";  // �� #197: �����ɶ�
        if (propLower == "autoredraw") return "vb6_SetAutoRedraw";  // P23-1: With img .AutoRedraw
        if (propLower == "caption") return "vb6_SetControlText";
        if (propLower == "picture") return "vb6_SetControlPicture";
        if (propLower == "autosize") return "vb6_SetPictureAutoSize";
        if (propLower == "align") return "vb6_SetControlAlign";  // P13.5b: Picture1.Align 停靠
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::Image:
        if (propLower == "picture") return "vb6_SetControlPicture";
        if (propLower == "stretch") return "vb6_SetImageStretch";  // P17.2
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::Winsock:  // C29-WS-a 写表。LocalPort / LocalIP **没有写口**：
        // VB6 那两格是"运行期只读"的（LocalPort 由 Bind 决定、LocalIP 由系统定），
        // 发一条"写得动但其实什么都不改"的 setter 比不发更难查（RT-a 的 ScrollBars 同口径）。
        if (propLower == "protocol")   return "vb6_Ws_SetProtocol";
        if (propLower == "remotehost") return "vb6_Ws_SetRemoteHost";
        if (propLower == "remoteport") return "vb6_Ws_SetRemotePort";
        break;
    case FrmControlType::MSComm:  // C30-4 写表。只读那几条（CommEvent / InBufferCount 读侧、
        // CTS·DSR·CDHolding）没有写口 —— 发一条"写得动但不改"的 setter 比不发更难查。
        if (propLower == "commport")      return "vb6_MsComm_SetCommPort";
        if (propLower == "settings")      return "vb6_MsComm_SetSettings";
        if (propLower == "portopen")      return "vb6_MsComm_SetPortOpen";
        if (propLower == "output")        return "vb6_MsComm_SetOutput";
        if (propLower == "inputmode")     return "vb6_MsComm_SetInputMode";
        if (propLower == "inputlen")      return "vb6_MsComm_SetInputLen";
        if (propLower == "inbuffercount") return "vb6_MsComm_SetInBufferCount";
        if (propLower == "outbuffercount") return "vb6_MsComm_SetOutBufferCount";
        if (propLower == "inbuffersize")  return "vb6_MsComm_SetInBufferSize";
        if (propLower == "outbuffersize") return "vb6_MsComm_SetOutBufferSize";
        if (propLower == "rthreshold")    return "vb6_MsComm_SetRThreshold";
        if (propLower == "sthreshold")    return "vb6_MsComm_SetSThreshold";
        if (propLower == "handshaking")   return "vb6_MsComm_SetHandshaking";
        if (propLower == "dtrenable")     return "vb6_MsComm_SetDTREnable";
        if (propLower == "rtsenable")     return "vb6_MsComm_SetRTSEnable";
        if (propLower == "eofenable")     return "vb6_MsComm_SetEOFEnable";
        // C30-4b：超时 / 设备名 / 二进制写
        if (propLower == "commdevice")    return "vb6_MsComm_SetCommDevice";
        if (propLower == "readintervaltimeout") return "vb6_MsComm_SetReadIntervalTimeout";
        if (propLower == "readtotaltimeout")    return "vb6_MsComm_SetReadTotalTimeout";
        if (propLower == "writetimeout")   return "vb6_MsComm_SetWriteTimeout";
        if (propLower == "outputbytes")    return "vb6_MsComm_SetOutputBytes";
        break;
    case FrmControlType::Adodc:  // C30-5 写表（Recordset / LastRowCount / Opened 是只读）
        if (propLower == "connectionstring") return "vb6_Adodc_SetConnectionString";
        if (propLower == "recordsource")  return "vb6_Adodc_SetRecordSource";
        if (propLower == "commandtype")   return "vb6_Adodc_SetCommandType";
        if (propLower == "cursorlocation") return "vb6_Adodc_SetCursorLocation";
        if (propLower == "cursortype")    return "vb6_Adodc_SetCursorType";
        if (propLower == "locktype")      return "vb6_Adodc_SetLockType";
        if (propLower == "mode")          return "vb6_Adodc_SetMode";
        if (propLower == "maxrecords")    return "vb6_Adodc_SetMaxRecords";
        if (propLower == "bofaction")     return "vb6_Adodc_SetBOFAction";
        if (propLower == "eofaction")     return "vb6_Adodc_SetEOFAction";
        break;
    case FrmControlType::DataGrid:  // C30-5 写表（RowCount/ColCount/Text 只读）
        if (propLower == "datasource")    return "vb6_DataGrid_SetDataSourceName";
        if (propLower == "row")           return "vb6_DataGrid_SetRow";
        if (propLower == "col")           return "vb6_DataGrid_SetCol";
        if (propLower == "columnheaders") return "vb6_DataGrid_SetColumnHeaders";
        if (propLower == "allowaddnew")   return "vb6_DataGrid_SetAllowAddNew";
        if (propLower == "allowupdate")   return "vb6_DataGrid_SetAllowUpdate";
        if (propLower == "allowdelete")   return "vb6_DataGrid_SetAllowDelete";
        if (propLower == "rowheight")     return "vb6_DataGrid_SetRowHeight";
        if (propLower == "defcolwidth")   return "vb6_DataGrid_SetDefColWidth";
        break;
    case FrmControlType::MshFlexGrid:  // C30-6 写表（Text 写当前格；TextMatrix/ColWidth 走索引属性）
        if (propLower == "text")          return "vb6_MshFlexGrid_SetText";
        if (propLower == "rows")          return "vb6_MshFlexGrid_SetRows";
        if (propLower == "cols")          return "vb6_MshFlexGrid_SetCols";
        if (propLower == "fixedrows")     return "vb6_MshFlexGrid_SetFixedRows";
        if (propLower == "fixedcols")     return "vb6_MshFlexGrid_SetFixedCols";
        if (propLower == "row")           return "vb6_MshFlexGrid_SetRow";
        if (propLower == "col")           return "vb6_MshFlexGrid_SetCol";
        if (propLower == "rowsel")        return "vb6_MshFlexGrid_SetRowSel";
        if (propLower == "colsel")        return "vb6_MshFlexGrid_SetColSel";
        if (propLower == "toprow")        return "vb6_MshFlexGrid_SetTopRow";
        if (propLower == "leftcol")       return "vb6_MshFlexGrid_SetLeftCol";
        if (propLower == "gridlines")     return "vb6_MshFlexGrid_SetGridLines";
        if (propLower == "scrollbars")    return "vb6_MshFlexGrid_SetScrollBars";
        if (propLower == "selectionmode") return "vb6_MshFlexGrid_SetSelectionMode";
        if (propLower == "fillstyle")     return "vb6_MshFlexGrid_SetFillStyle";
        if (propLower == "wordwrap")      return "vb6_MshFlexGrid_SetWordWrap";
        if (propLower == "editable")      return "vb6_MshFlexGrid_SetEditable";
        if (propLower == "focusrect")     return "vb6_MshFlexGrid_SetFocusRect";
        if (propLower == "allowuserresizing") return "vb6_MshFlexGrid_SetAllowUserResizing";
        if (propLower == "mergecells")    return "vb6_MshFlexGrid_SetMergeCells";
        if (propLower == "redraw")        return "vb6_MshFlexGrid_SetRedraw";
        if (propLower == "backcolor")     return "vb6_MshFlexGrid_SetBackColor";
        if (propLower == "forecolor")     return "vb6_MshFlexGrid_SetForeColor";
        if (propLower == "fillcolor")     return "vb6_MshFlexGrid_SetFillColor";
        if (propLower == "cellbackcolor") return "vb6_MshFlexGrid_SetCellBackColor";
        if (propLower == "cellforecolor") return "vb6_MshFlexGrid_SetCellForeColor";
        if (propLower == "formatstring")  return "vb6_MshFlexGrid_SetFormatString";
        if (propLower == "datasource")    return "vb6_MshFlexGrid_SetDataSource";
        if (propLower == "rowheightmin")  return "vb6_MshFlexGrid_SetRowHeightMin";
        if (propLower == "sort")          return "vb6_MshFlexGrid_SetSort";
        break;
    case FrmControlType::Data:  // C29-Data 写表 (三属性先存后 Refresh 用)
        if (propLower == "databasename") return "vb6_Data_SetDatabaseName";
        if (propLower == "recordsource") return "vb6_Data_SetRecordSource";
        if (propLower == "connect")      return "vb6_Data_SetConnect";
        break;
    case FrmControlType::Menu:  // P20-36
        if (propLower == "caption") return "vb6_SetMenuCaption";
        if (propLower == "checked") return "vb6_SetMenuChecked";
        if (propLower == "enabled") return "vb6_SetMenuEnabled";
        if (propLower == "visible") return "vb6_SetMenuVisible";
        break;
    // C29-1b: 写侧同理 —— Path / Pattern / Drive 一赋就重刷列表 (RTL setter 内部
    // 调 Refresh), 这正是 VB6 三控件联动的机制。
    case FrmControlType::DriveListBox:
        if (propLower == "drive") return "vb6_DriveListBoxSetDrive";
        if (propLower == "text") return "vb6_SetControlText";
        if (propLower == "listindex") return "vb6_SetListIndex";
        break;
    case FrmControlType::DirListBox:
        if (propLower == "path") return "vb6_DirListBoxSetPath";
        if (propLower == "listindex") return "vb6_SetListIndex";
        break;
    case FrmControlType::FileListBox:
        if (propLower == "path") return "vb6_FileListBoxSetPath";
        if (propLower == "pattern") return "vb6_FileListBoxSetPattern";
        if (propLower == "filename") return "vb6_FileListBoxSetFileName";
        if (propLower == "listindex") return "vb6_SetListIndex";
        break;
    // C29-8a: TreeView 写侧。四条样式位的 setter 会连带 SWP_FRAMECHANGED + 重绘
    // (复选框位改的是每个节点的度量)，所以运行期赋值立刻见效，不是"只改了张表"。
    case FrmControlType::TreeView:
        if (propLower == "linestyle") return "vb6_TreeView_SetLineStyle";
        if (propLower == "indentation") return "vb6_TreeView_SetIndentation";
        if (propLower == "iconsize") return "vb6_TreeView_SetIconSize";
        if (propLower == "checkboxes") return "vb6_TreeView_SetCheckBoxes";
        if (propLower == "hottracking") return "vb6_TreeView_SetHotTracking";
        if (propLower == "hideselection") return "vb6_TreeView_SetHideSelection";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    // C29-DT-a: DTPicker 写侧 (与读侧同一批)。Format / CheckBox / UpDown 的 setter 会连带
    // SWP_FRAMECHANGED + 重绘 (格式变了显示区宽度就得重算)，CustomFormat 发 DTM_SETFORMATW。
    case FrmControlType::DTPicker:
        if (propLower == "format") return "vb6_DTP_SetFormat";
        if (propLower == "customformat") return "vb6_DTP_SetCustomFormat";
        if (propLower == "checkbox") return "vb6_DTP_SetCheckBox";
        if (propLower == "updown") return "vb6_DTP_SetUpDown";
        if (propLower == "calendarbackcolor") return "vb6_DTP_SetCalendarBackColor";
        if (propLower == "calendarforecolor") return "vb6_DTP_SetCalendarForeColor";
        if (propLower == "calendartrailingforecolor") return "vb6_DTP_SetCalendarTrailingForeColor";
        if (propLower == "calendartitlebackcolor") return "vb6_DTP_SetCalendarTitleBackColor";
        if (propLower == "calendartitleforecolor") return "vb6_DTP_SetCalendarTitleForeColor";
        // C29-DT-b 的写侧。HasDate 刻意不给写口（它是原生 GDT_NONE 那一态的读数，
        // 要"清空"请用 CheckBox 那枚勾选框，写它会把读数与观感拆成两张皮）。
        if (propLower == "value") return "vb6_DTP_SetValue";
        if (propLower == "hasdate") return "vb6_DTP_SetHasDate";
        if (propLower == "mindate") return "vb6_DTP_SetMinDate";
        if (propLower == "maxdate") return "vb6_DTP_SetMaxDate";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    // C29-MV-a: MonthView 写侧 (与读侧同一批)。样式那三位的 setter 是"尽力而为"—— 它们是
    // 创建参数，运行期落不落地由夹具的读数说；MonthRows/MonthColumns 刻意不给写口（多月
    // 平铺在原生里等于"窗口多大"，运行期改它得连带挪窗，那是另一格的事）。
    case FrmControlType::MonthView:
        if (propLower == "multiselect") return "vb6_MV_SetMultiSelect";
        if (propLower == "showweeknumbers") return "vb6_MV_SetShowWeekNumbers";
        if (propLower == "showtoday") return "vb6_MV_SetShowToday";
        if (propLower == "maxselcount") return "vb6_MV_SetMaxSelCount";
        if (propLower == "backcolor") return "vb6_MV_SetBackColor";
        if (propLower == "forecolor") return "vb6_MV_SetForeColor";
        if (propLower == "titlebackcolor") return "vb6_MV_SetTitleBackColor";
        if (propLower == "titleforecolor") return "vb6_MV_SetTitleForeColor";
        if (propLower == "trailingforecolor") return "vb6_MV_SetTrailingForeColor";
        if (propLower == "monthbackcolor") return "vb6_MV_SetMonthBackColor";
        // C29-MV-b 的写侧（Date 三条；两端表改一端由 RTL 读回整张再发回去）。
        if (propLower == "value") return "vb6_MV_SetValue";
        if (propLower == "selstart") return "vb6_MV_SetSelStart";
        if (propLower == "selend") return "vb6_MV_SetSelEnd";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    // C29-SL-a: Slider 写侧。Orientation 实测**运行期改有效**（写 TBS_VERT + SetWindowPos 换帧
    // 之后滑块位移轴跟着换，见 vb6forms_slider.c 文件头第 2 条）；两条证人读数刻意不给写口。
    // C30-UD-a: UpDown 写侧。Wrap/Alignment/Orientation/SyncBuddy 四条都属于
    // 「可运行期改」的样式位（RTL 里写完样式位再 SetWindowPos 重排一次）。
    case FrmControlType::UpDown:
        if (propLower == "min") return "vb6_UpDown_SetMin";
        if (propLower == "max") return "vb6_UpDown_SetMax";
        if (propLower == "value") return "vb6_UpDown_SetValue";
        if (propLower == "increment") return "vb6_UpDown_SetIncrement";
        if (propLower == "wrap") return "vb6_UpDown_SetWrap";
        if (propLower == "alignment") return "vb6_UpDown_SetAlignment";
        if (propLower == "orientation") return "vb6_UpDown_SetOrientation";
        if (propLower == "syncbuddy") return "vb6_UpDown_SetSyncBuddy";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    case FrmControlType::Slider:
        if (propLower == "orientation") return "vb6_Slider_SetOrientation";
        if (propLower == "tickfrequency") return "vb6_Slider_SetTickFrequency";
        // C29-SL-b 的写侧（值面）。
        if (propLower == "min") return "vb6_Slider_SetMin";
        if (propLower == "max") return "vb6_Slider_SetMax";
        if (propLower == "value") return "vb6_Slider_SetValue";
        if (propLower == "smallchange") return "vb6_Slider_SetSmallChange";
        if (propLower == "largechange") return "vb6_Slider_SetLargeChange";
        if (propLower == "selectrange") return "vb6_Slider_SetSelectRange";
        if (propLower == "selstart") return "vb6_Slider_SetSelStart";
        if (propLower == "selend") return "vb6_Slider_SetSelEnd";
        // C29-SL-h: TickStyle 运行期写有效（实测写位 + 换帧之后的每一条形与创建时带那一位逐字
        // 相同）。GetNumTicks / ChannelTop 两条刻意不给写口 —— 前者 VB6 就是只读。
        if (propLower == "tickstyle") return "vb6_Slider_SetTickStyle";
        // C29-SL-i: SelLength 写侧 = 起点不动、终点 = 起点 + 长度（远端超量程由控件夹住，实测）。
        if (propLower == "sellength") return "vb6_Slider_SetSelLength";
        // C29-SL-k: Text 写的是气泡串（同时把宿主里那条工具的文本换掉）；TextPosition 自存
        // 一枚 0/1（原生没有这条消息），摆气泡时用它定上/下。
        if (propLower == "text") return "vb6_Slider_SetText";
        if (propLower == "textposition") return "vb6_Slider_SetTextPosition";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    // C29-RT-a: RichTextBox 写侧。ScrollBars 刻意**没有**写口 —— 它是滚动条子窗口的创建参数，
    // 事后 SetWindowLong 只有外观、量程停在默认（探针量出来的，见 029 §九 本格）；发一条
    // "写了但没用"的 setter 比不发更难查，所以运行期一律读回设计期那个值。
    case FrmControlType::RichTextBox:
        if (propLower == "text") return "vb6_SetControlText";
        if (propLower == "selstart") return "vb6_RTB_SetSelStart";
        if (propLower == "sellength") return "vb6_RTB_SetSelLength";
        if (propLower == "seltext") return "vb6_RTB_SetSelText";
        if (propLower == "readonly") return "vb6_RTB_SetReadOnly";
        if (propLower == "maxlength") return "vb6_RTB_SetMaxLength";
        if (propLower == "wordwrap") return "vb6_RTB_SetWordWrap";
        // C29-RT-b 的写侧（与读侧同一批十一条）。
        if (propLower == "selbold") return "vb6_RTB_SetSelBold";
        if (propLower == "selitalic") return "vb6_RTB_SetSelItalic";
        if (propLower == "selunderline") return "vb6_RTB_SetSelUnderline";
        if (propLower == "selstrikethru") return "vb6_RTB_SetSelStrikethru";
        if (propLower == "selcolor") return "vb6_RTB_SetSelColor";
        if (propLower == "selfontname") return "vb6_RTB_SetSelFontName";
        if (propLower == "selfontsize") return "vb6_RTB_SetSelFontSize";
        if (propLower == "selalignment") return "vb6_RTB_SetSelAlignment";
        if (propLower == "selindent") return "vb6_RTB_SetSelIndent";
        if (propLower == "selrightindent") return "vb6_RTB_SetSelRightIndent";
        if (propLower == "selhangingindent") return "vb6_RTB_SetSelHangingIndent";
        // C29-RT-c: 写 TextRTF = EM_STREAMIN(SF_RTF)；RTL 那边先全选再灌 ⇒ 语义是"换掉内容"。
        if (propLower == "textrtf") return "vb6_RTB_SetTextRTF";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    // C29-5a: Toolbar 写侧 (与读侧同一批四条 + 通用两条)。
    case FrmControlType::Toolbar:
        if (propLower == "showtips") return "vb6_Toolbar_SetShowTips";
        if (propLower == "textstyle") return "vb6_Toolbar_SetTextStyle";
        if (propLower == "allowcustomize") return "vb6_Toolbar_SetAllowCustomize";
        if (propLower == "align") return "vb6_Toolbar_SetAlign";
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    // D6 / C29-9: CommonDialog 写侧（取消由 RTL 按 CancelError 决定报不报 32755）。
    case FrmControlType::CommonDialog:
        if (propLower == "filter") return "vb6_CdSetFilter";
        if (propLower == "filename") return "vb6_CdSetFileName";
        if (propLower == "filetitle") return "vb6_CdSetFileTitle";
        if (propLower == "dialogtitle") return "vb6_CdSetDialogTitle";
        if (propLower == "initdir") return "vb6_CdSetInitDir";
        if (propLower == "defaultext") return "vb6_CdSetDefaultExt";
        if (propLower == "fontname") return "vb6_CdSetFontName";
        if (propLower == "flags") return "vb6_CdSetFlags";
        if (propLower == "cancelerror") return "vb6_CdSetCancelError";
        if (propLower == "color") return "vb6_CdSetColor";
        if (propLower == "min") return "vb6_CdSetMin";
        if (propLower == "max") return "vb6_CdSetMax";
        if (propLower == "copies") return "vb6_CdSetCopies";
        if (propLower == "fontsize") return "vb6_CdSetFontSize";
        if (propLower == "filterindex") return "vb6_CdSetFilterIndex";
        break;
    case FrmControlType::Shape:  // P20-35
        if (propLower == "shape") return "vb6_SetShapeType";
        if (propLower == "borderwidth") return "vb6_SetShapeBorderWidth";
        if (propLower == "borderstyle") return "vb6_SetShapeBorderStyle";
        if (propLower == "fillstyle") return "vb6_SetShapeFillStyle";
        if (propLower == "bordercolor") return "vb6_SetShapeBorderColor";
        if (propLower == "fillcolor") return "vb6_SetShapeFillColor";
        if (propLower == "visible") return "vb6_SetControlVisible";
        break;
    case FrmControlType::Line:  // P20-35
        if (propLower == "x1") return "vb6_SetLineX1";
        if (propLower == "y1") return "vb6_SetLineY1";
        if (propLower == "x2") return "vb6_SetLineX2";
        if (propLower == "y2") return "vb6_SetLineY2";
        if (propLower == "borderwidth") return "vb6_SetLineBorderWidth";
        if (propLower == "borderstyle") return "vb6_SetLineBorderStyle";
        if (propLower == "bordercolor") return "vb6_SetLineColor";
        if (propLower == "visible") return "vb6_SetControlVisible";
        break;
    default:
        if (propLower == "visible") return "vb6_SetControlVisible";
        if (propLower == "enabled") return "vb6_SetControlEnabled";
        break;
    }
    return "";  // 未知属性
}

// 应用 .frm 设计期样式属性 (BorderStyle / Alignment)。
// 说明: 这些属性不会改变窗口类默认外观，必须在控件创建后调用对应 RTL setter；
//       顶层控件与容器(Frame/PictureBox)子控件都需要应用。
void CCodeGen::emitDesignerStyleProps(const FrmControl& ctrl, const std::string& hwndExpr) {
    std::string hw = "(void*)" + hwndExpr;

    auto bsIt = ctrl.properties.find("BorderStyle");
    if (bsIt != ctrl.properties.end()) {
        switch (ctrl.controlType) {
        case FrmControlType::Label:
        case FrmControlType::Frame:
        case FrmControlType::PictureBox:
        case FrmControlType::Image:
        case FrmControlType::TextBox:
        case FrmControlType::RichTextBox:   // C29-RT-a: 原生这枚的边框就在 WS_EX_CLIENTEDGE 上
            c_.emitLine("vb6_SetBorderStyle(" + hw + ", " + std::to_string((int)bsIt->second.intValue) + ");");
            break;
        default:
            break;
        }
    }

    auto alIt = ctrl.properties.find("Alignment");
    if (alIt != ctrl.properties.end() &&
        (ctrl.controlType == FrmControlType::Label || ctrl.controlType == FrmControlType::TextBox)) {
        c_.emitLine("vb6_SetAlignment(" + hw + ", " + std::to_string((int)alIt->second.intValue) + ");");
    }

    // Fix <c3-menu3d-labelbg>: Label 设计期 BackStyle 此前**两条创建路都不发**。
    //
    // VB6 语义: BackStyle=1 (Opaque, 默认) 时 Label 用 BackColor 填底; BackStyle=0
    // (Transparent) 时**完全不画背景**, 文字直接坐在父窗底色/图片上。
    // 窗口层这边 Label 是 STATIC 窗口类, 它的 WM_PAINT 会先拿 WM_CTLCOLORSTATIC
    // 返回的刷子填一块底 —— 不接管就是类背景刷 (COLOR_WINDOW = 纯白)。
    //
    // 实测 (3DMenu): Label3 / LblSub 在 .frm 里都是 `BackStyle = 0 'Transparent`,
    // 且**都没有 BackColor** (VB6 里 BackStyle=0 时 BackColor 不参与绘制, 设计期
    // 就不写)。于是它拿不到 VB6_BackColorSet 这个哨兵, 一路走到 DefWindowProcW,
    // 在窗体左下角留下一块 241x33 的纯白方块 (用户可见)。
    //
    // 发射口径: 只在 .frm **显式写过** BackStyle 时发 (VB6 默认 1, 没写的控件逐字节
    // 维持今天的行为 —— 与 emitDesignerStateProps 的"反面才发"同一条纪律)。
    // 两个值都发: 0 → 空刷透出父窗; 1 → 显式还原不透明, 让这条 API 的两个分支
    // 都可达 (否则 val!=0 分支是死代码)。
    auto bksIt = ctrl.properties.find("BackStyle");
    if (bksIt != ctrl.properties.end() && ctrl.controlType == FrmControlType::Label
        && bksIt->second.type == FrmValueType::Integer) {
        c_.emitLine("vb6_SetLabelBackStyle(" + hw + ", "
                    + std::to_string((int)bksIt->second.intValue) + ");  /* design BackStyle */");
    }
}

// 账 #125: 设计期 Enabled / Visible / Value 三件此前**两条创建路都没打到窗口上**
// (实测: 生成码里 BM_SETCHECK / EnableWindow / ShowWindow 各 0 次, 运行期一律读回默认值)。
// 只在 .frm **显式写过**这一项时才发: 没写的控件逐字节维持今天的行为。
// True/可见/未勾是 Win32 创建出来的默认观感, 所以三条都只在"反面"发一条。
void CCodeGen::emitDesignerStateProps(const FrmControl& ctrl, const std::string& hwndExpr) {
    switch (ctrl.controlType) {
    case FrmControlType::Timer:   // 无窗口控件, Enabled 由 C29-T 那条路管
    case FrmControlType::Menu:    // 菜单项走 EnableMenuItem, 不是窗口状态
    case FrmControlType::ImageList:
        return;
    default:
        break;
    }
    const std::string hw = "(void*)" + hwndExpr;

    auto enIt = ctrl.properties.find("Enabled");
    if (enIt != ctrl.properties.end() && enIt->second.type == FrmValueType::Integer
        && enIt->second.intValue == 0) {
        c_.emitLine("vb6_SetControlEnabled(" + hw + ", 0);  /* design Enabled=False */");
    }

    auto viIt = ctrl.properties.find("Visible");
    if (viIt != ctrl.properties.end() && viIt->second.type == FrmValueType::Integer
        && viIt->second.intValue == 0) {
        c_.emitLine("vb6_SetControlVisible(" + hw + ", 0);  /* design Visible=False */");
    }

    auto vaIt = ctrl.properties.find("Value");
    if (vaIt != ctrl.properties.end() && vaIt->second.type == FrmValueType::Integer) {
        const int v = (int)vaIt->second.intValue;
        if (ctrl.controlType == FrmControlType::CheckBox) {
            // VB6: 0=Unchecked 1=Checked 2=Grayed —— 0 就是创建默认, 不发
            if (v == 1 || v == 2) {
                c_.emitLine("vb6_SetCheckValue(" + hw + ", " + std::to_string(v)
                            + ");  /* design Value=" + std::to_string(v) + " */");
            }
        } else if (ctrl.controlType == FrmControlType::OptionButton && v != 0) {
            c_.emitLine("vb6_SetOptionValue(" + hw + ", 1);  /* design Value=True(账 #128-b 专桩) */");
        }
    }

    // 账 (TeeChart Chart display 夹具): HScrollBar/VScrollBar 的设计期 Min/Max/Value 在
    // 两条创建路上都**一条不发** —— 控件窗口的量程停在默认 0..0，运行期
    // `HScroll4.Value = 32` 被 SetScrollPos 夹回 0（实测 C3_EVTRACE：old=0 new=0），
    // 于是 VB6 里必然发生的 Change 一声也没有，Panel.MarginRight 整条丢失。
    // 顺序必须是 Min → Max → Value（量程立不起来之前写 Value 无意义）。
    // 只在非默认值时发，与上面 Value 的"反面才发"同一条纪律。
    if (ctrl.controlType == FrmControlType::HScrollBar || ctrl.controlType == FrmControlType::VScrollBar) {
        auto mnIt = ctrl.properties.find("Min");
        if (mnIt != ctrl.properties.end() && mnIt->second.type == FrmValueType::Integer
            && mnIt->second.intValue != 0) {
            c_.emitLine("vb6_SetScrollMin(" + hw + ", " + std::to_string(mnIt->second.intValue)
                        + ");  /* design Min */");
        }
        auto mxIt = ctrl.properties.find("Max");
        if (mxIt != ctrl.properties.end() && mxIt->second.type == FrmValueType::Integer
            && mxIt->second.intValue != 32767) {
            c_.emitLine("vb6_SetScrollMax(" + hw + ", " + std::to_string(mxIt->second.intValue)
                        + ");  /* design Max */");
        }
        auto vaIt2 = ctrl.properties.find("Value");
        if (vaIt2 != ctrl.properties.end() && vaIt2->second.type == FrmValueType::Integer
            && vaIt2->second.intValue != 0) {
            c_.emitLine("vb6_SetScrollValue(" + hw + ", " + std::to_string(vaIt2->second.intValue)
                        + ");  /* design Value */");
        }
    }
}

// 账 #142: 设计期的 `ToolTipText` / `Tag` 此前**两条创建路都没发**（实测：`.frm` 里写了
// `ToolTipText = "dtip"`，运行期读回来是空串 —— 属性行解析到了、就是没人发射）。
// 两条口径：
//   · **只在显式写了非空值时才发**。空串与"没写过"在 VB6 里读数同为 `""`，而运行期那一步会把
//     值存进窗口属性**并往共享 tooltip 控件注册一条工具** —— 多发一条 = 改产物去做一件语义上
//     零的事。与 #125 那三条"反面才发"是同一条纪律。
//   · 值面**不转小写**（`Symbol::toLower` 是给标识符用的，转了就改坏用户写的文本），只补
//     C 字面量要的两个转义：反斜杠先走，双引号后走。
//   · **只认 `FrmValueType::String`**。真实的 VB6 .frm 会把资源引用写进这两条属性
//     （`ToolTipText=   "frmTest.frx":0000`），解析器把那种行归成 `FrxReference` 而不是
//     `String` ⇒ 这里直接跳过。顺带挡掉两件事：把 `frmTest.frx":0000` 当字面文本发出去
//     （转义后是个假字符串），以及以后有人给 `FrxReference` 补 `.frx` 解析时把类型比较
//     写成 `== FrmValueType::FrxReference`（那会把引用当文本灌进 `SysAllocString`）。
void CCodeGen::emitDesignerStringProps(const FrmControl& ctrl, const std::string& hwndExpr) {
    const std::string hw = "(void*)" + hwndExpr;
    auto emitOne = [&](const char* key, const char* fn) {
        auto it = ctrl.properties.find(key);
        if (it == ctrl.properties.end()) return;
        if (it->second.type != FrmValueType::String) return;
        std::string raw = it->second.rawText;   // .frm 里的原始文本，含首尾双引号
        if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
            raw = raw.substr(1, raw.size() - 2);
        if (raw.empty()) return;
        std::string esc;
        for (size_t i = 0; i < raw.size(); i++) {
            char ch = raw[i];
            if (ch == '"' || ch == '\\') esc += '\\';
            esc += ch;
        }
        c_.emitLine(std::string(fn) + "(" + hw + ", L\"" + esc + "\");  /* design "
                    + key + " */");
    };
    emitOne("ToolTipText", "vb6_SetToolTipText");
    emitOne("Tag", "vb6_SetControlTag");
}

// 账 #154（C29-SL-q，从 `ai/内置控件/Slider 控件（滑杆）.md` §4 那条例子顺带量出来的）:
// 设计期的 `FontName` / `FontSize` 此前**两条创建路都没打到窗口上** —— 探针 `.build/slfont` 里
// 一枚写 Consolas+14 的文本框运行时读回 空串/8.25，另一枚写 20 的读回 8.25、窗口像素高度还是默认的 11。
// 解析侧是留着这两条的（`frm_parser.cpp:388` 原样存进 `ctrl.properties`），缺的只有发码这一步
// ⇒ 与 #125（Enabled/Visible/Value）、#142（ToolTipText/Tag）同一形状：运行期赋值一直是好的。
// 只在 `.frm` **显式写过**时发，没写的控件逐字节维持今天的行为。
// 发序是 先 Size 后 Name：两支 setter 都是"读当前 LOGFONT、只改自己那一格"，而 Name 那支在窗口
// 本来没有字体时会先造一枚 `-13`（≈10pt）的底子 —— 把 Size 放在前面就不会被它盖掉。
void CCodeGen::emitDesignerFontProps(const FrmControl& ctrl, const std::string& hwndExpr) {
    switch (ctrl.controlType) {
    case FrmControlType::Timer:         // 无窗口控件：字体没有落脚的地方
    case FrmControlType::Menu:          // 菜单项的字体是 owner-draw 的事，不是窗口字体
    case FrmControlType::ImageList:     // 不是窗口
    case FrmControlType::CommonDialog:  // 它的 Font* 是 ChooseFont 的 LOGFONT 字段（D6 口径），
                                        // 已由上面那条 cdStr/cdInt 分支发过，这里再发就是两遍
        return;
    default:
        break;
    }
    const std::string hw = "(void*)" + hwndExpr;

    auto szIt = ctrl.properties.find("FontSize");
    if (szIt != ctrl.properties.end()) {
        double pt = -1.0;
        if (szIt->second.type == FrmValueType::Float) pt = szIt->second.floatValue;
        else if (szIt->second.type == FrmValueType::Integer) pt = (double)szIt->second.intValue;
        // 0 与负数不是字号（VB6 的设计期也不会写这种数），发了只会把窗口打成"默认字体"那一档。
        if (pt > 0.0) {
            c_.emitLine("vb6_SetControlFontSize(" + hw + ", "
                        + floatFixed6Literal(pt) + ");  /* design FontSize */");
        }
    }

    auto nmIt = ctrl.properties.find("FontName");
    if (nmIt != ctrl.properties.end() && nmIt->second.type == FrmValueType::String) {
        std::string raw = nmIt->second.rawText;   // .frm 里的原始文本，含首尾双引号
        if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
            raw = raw.substr(1, raw.size() - 2);
        if (!raw.empty()) {
            std::string esc;
            for (size_t i = 0; i < raw.size(); i++) {
                char ch = raw[i];
                if (ch == '"' || ch == '\\') esc += '\\';
                esc += ch;
            }
            c_.emitLine("vb6_SetControlFontNameW(" + hw + ", L\"" + esc
                        + "\");  /* design FontName */");
        }
    }
}

// 账 #160: 设计期的 `TabIndex` 下发到窗口。RTL 侧那一对 `vb6_GetTabIndex` / `vb6_SetTabIndex`
// （`vb6forms_style.c:86/93`，存成窗口属性 `VB6_TabIndex`， getter 没存过时答 0）早就齐了，
// 缺的只有"创建时没人发"这一刀 ⇒ `.frm` 里写着 `TabIndex = 4`，运行期 `Ctl.TabIndex` 读回 0。
// 口径两条：① `.frm` 写了就照写的发（**每个父窗自己从 0 编号**，VB6 就是这样的 —— 框架里的
// 控件不占窗体那一串的号）；② 没写用**声明序号**兜底（VB6 存盘时从不省这一行，省了的都是
// 手写夹具；兜成 0 会让好几枚控件同时声称自己是 0，兜成声明序至少是个全序）。
// 与 #125 / #142 / #154 / #83 同一形状：**两条创建路都要调**，只接一头就是本线踩过几次的那声不响。
void CCodeGen::emitDesignerTabIndexProp(const FrmControl& ctrl, const std::string& hwndExpr,
                                        long declarationIndex) {
    switch (ctrl.controlType) {
    case FrmControlType::Timer:         // 无窗口控件：号没地方存
    case FrmControlType::Menu:          // 菜单项不参与 tab 序
    case FrmControlType::ImageList:     // 不是窗口
    case FrmControlType::CommonDialog:  // 不是窗口
    case FrmControlType::Data:          // 不是窗口
    case FrmControlType::Unknown:
        return;
    default:
        break;
    }
    long tabIndex = declarationIndex;
    auto it = ctrl.properties.find("TabIndex");
    if (it != ctrl.properties.end() && it->second.type == FrmValueType::Integer) {
        tabIndex = (long)it->second.intValue;
    }
    c_.emitLine("vb6_SetTabIndex((void*)" + hwndExpr + ", " + std::to_string(tabIndex)
                + ");  /* 账 #160: design TabIndex */");
}

// 账 #197（§B32）: 设计期写下的 `ScaleMode` 以前从来没有落到窗口上 —— 读的一侧早就齐了
// (vb6_GetScaleMode / vb6_WindowScaleModeSelf / vb6_ContainerScaleMode，缺省 1=缇)，
// 而写的一侧**全仓 0 个调用者** ⇒ 谁读 ScaleMode 都答缺省。这一档决定的是
// **量出来的数对不对**（ScaleWidth/ScaleHeight、控件几何、文字量纲全按它折算，见账 #175/#177），
// 不是某一枚控件的外观。语料普查 93 份设计块共 19 处 ScaleMode —— UserControl 10 处
// （已由 .ctl 注册那条路接走，不在这儿）、PictureBox 6 处、Form 3 处，全是 VB6 真有此属性的型 ⇒
// 口径收成「设计块写了就发」：不按值筛（声明 1 与没声明在产物里要分得开），也不按控件型再开一张
// 白名单（Frame 那一类根本不写这一行，写了就是给人读的）。
void CCodeGen::emitDesignerScaleModeProp(const FrmControl& ctrl, const std::string& hwndExpr) {
    auto it = ctrl.properties.find("ScaleMode");
    if (it == ctrl.properties.end() || it->second.type != FrmValueType::Integer) return;
    c_.emitLine("vb6_SetScaleMode((void*)" + hwndExpr + ", "
                + std::to_string((int)it->second.intValue)
                + ");  /* 账 #197: design ScaleMode */");
}

// C29-SL-g: 这个控件的 **焦点事件**（`GotFocus` / `LostFocus`）是不是已经由原生通知送进来了。
// 与上一条同型的问题在焦点这一档更隐蔽：仓里同一句判据被抄成了**三份表**（发 arm 的
// `cgen_form_wndproc_subclass.inc`、装子类的 `cgen_form_frame_menu.inc`、拆子类的
// `cgen_form_wndproc_dispatch.inc`），三份里都写着同一份四类白名单
// （PictureBox/Frame/Label/Image），而 TextBox/ComboBox/ListBox/命令按钮那批走的是另一条路：
// 父窗的 `WM_COMMAND` 通知码（见 dispatch 那一趟开头：EN_SETFOCUS=256 / EN_KILLFOCUS=512、
// CBN_=3/4、LBN_=4/5、BN_=6/7）。**Combo 那两位在账 #158 之前一直写成 1024/2048** ——
// 那是 CBEM_*（发给 ComboBoxEx 窗口的消息号，永远不会作为 notification code 出现），
// 所以 ComboBox 的焦点通知从能编事件那天起就没触发过（订正见 dispatch 那一趟的注记）。
// 于是剩下的所有窗口态控件 —— **Slider 首当其冲**，
// 还有 ListView / TreeView / DTPicker / MonthView / RichTextBox / 两个滚动条 / 文件系统三件套 ——
// 的 `_GotFocus` / `_LostFocus` 是编得过、永不触发的死处理器（轨道条那两条实测过：原生只往父窗发
// `WM_HSCROLL`，焦点变化一律以 `WM_SETFOCUS` / `WM_KILLFOCUS` 到**控件自己**的过程中，
// 见 029 §C29-SL-d-0 第 1 条与 `.build/slprobe/slmeasure9.c`）。
// 这张排除表就是"已经有 WM_COMMAND 那一条来源"的那六类 —— 剩下的都交给子类化那一档，
// 同一次焦点变化不会有两处发。
bool CCodeGen::controlFocusFromNativeNotify(FrmControlType ctrlType) {
    switch (ctrlType) {
    case FrmControlType::TextBox:
    case FrmControlType::MaskEdBox:   // C30-3: EDIT 类，EN_SETFOCUS/EN_KILLFOCUS 走 WM_COMMAND
    case FrmControlType::ComboBox:
    case FrmControlType::ListBox:
    case FrmControlType::CommandButton:
    case FrmControlType::CheckBox:
    case FrmControlType::OptionButton:
        return true;
    default:
        return false;
    }
}

// 账 #158（按钮那一半）: BUTTON 类要挂上 `BS_NOTIFY` 才会把焦点变化作为
// `WM_COMMAND` 的 `BN_SETFOCUS=6` / `BN_KILLFOCUS=7` 报给父窗。裸码探针
// `.build/bnnotify/bnnotify.c`（同一父窗两枚按钮，只差这一位）实测：挂了的那枚，
// 无论程序化 `SetFocus` 还是对话框管理器（`IsDialogMessage` + VK_TAB）都把 6/7 送到父窗；
// 没挂的那枚一条都不送（而 `BN_CLICKED=0` 不需要这一位，所以按钮的 `_Click` 一直是通的）。
// 两条创建路共用这一处判断（顶层那条在 cgen_form_ctrl_style_apply.inc，容器子控件
// 那条在 cgen_form_frame_menu.inc）——与 #83(a) 的 `WS_TABSTOP`、#83(b2) 的
// `WS_EX_CONTROLPARENT` 同一族：判据只写一遍，两处各调一次。
long CCodeGen::controlButtonNotifyStyleBit(const FrmControl& ctrl) const {
    constexpr long kBsNotify = 0x00004000L;  // BS_NOTIFY（SDK 头里的实测值 = 16384）
    switch (ctrl.controlType) {
    case FrmControlType::CommandButton:
    case FrmControlType::CheckBox:
    case FrmControlType::OptionButton:
        return kBsNotify;
    default:
        return 0L;
    }
}

// C29-SL-a/h: Slider 的创建样式位 —— **两条创建路共用这一处**（顶层那条在
// cgen_form_ctrl_style_apply.inc 的 Slider 分支，容器子控件那条走下面的
// controlTypeStyleBits，账 #83 说的就是这两张表会各写各的）。
// 取值与 vb6forms_slider.c 里的 getter 同一档：
//   · TBS_AUTOTICKS(0x0001) 一律挂上：VB6 的 Slider 默认就画刻度。
//   · Orientation 0=水平 / 1=垂直 → TBS_VERT(0x0002)。
//   · TickStyle 四档 → TBS_TOP(0x0004) / TBS_BOTH(0x0008) / TBS_NOTICKS(0x0010)，
//     0=sldBottomRight 就是那三位全清（TBS_BOTTOM==TBS_RIGHT==0）。四档的数值是
//     从 OCX 自带的类型库读出来的（探针 .build/slprobe/sltlb.cpp），不是猜的；
//     每一档在控件侧的证人实测于 .build/slprobe/slmeasure12.c。
long CCodeGen::sliderStyleBits(const FrmControl& ctrl) const {
    constexpr long kTbsAutoTicks = 0x0001L;
    constexpr long kTbsVert      = 0x0002L;
    constexpr long kTbsTop       = 0x0004L;   // == TBS_LEFT
    constexpr long kTbsBoth      = 0x0008L;
    constexpr long kTbsNoTicks   = 0x0010L;
    constexpr long kTbsPosMask   = kTbsTop | kTbsBoth | kTbsNoTicks;

    long style = kTbsAutoTicks;
    auto slOr = ctrl.properties.find("Orientation");
    if (slOr != ctrl.properties.end() && (int)slOr->second.intValue != 0) style |= kTbsVert;
    auto slTs = ctrl.properties.find("TickStyle");
    if (slTs != ctrl.properties.end()) {
        long pos = 0;
        switch ((int)slTs->second.intValue) {
            case 1: pos = kTbsTop; break;
            case 2: pos = kTbsBoth; break;
            case 3: pos = kTbsNoTicks; break;
            default: pos = 0; break;   // 0 与越界值都落"下半/右半"那一档（与 setter 同口径）
        }
        style = (style & ~kTbsPosMask) | pos;
    }
    return style;
}

// C30-UD-a: UpDown 的创建样式位。顶层与容器子控件两条路共用这一处（账 #83：不许每路各写一张）。
// 常量取自 commctrl.h：UDS_WRAP 0x1 / UDS_SETBUDDYINT 0x2 / UDS_ALIGNRIGHT 0x4 /
// UDS_ALIGNLEFT 0x8 / UDS_ARROWKEYS 0x20 / UDS_HORZ 0x40 / UDS_HOTTRACK 0x100。
// 两条默认值按 VB6 文档口径：Alignment = 1（右）、Orientation = 0（竖）。
long CCodeGen::updownStyleBits(const FrmControl& ctrl) const {
    constexpr long kUdsWrap       = 0x0001L;
    constexpr long kUdsSetBuddy   = 0x0002L;
    constexpr long kUdsAlignRight = 0x0004L;
    constexpr long kUdsAlignLeft  = 0x0008L;
    constexpr long kUdsArrowKeys  = 0x0020L;
    constexpr long kUdsHorz       = 0x0040L;
    constexpr long kUdsHotTrack   = 0x0100L;

    // ARROWKEYS：VB6 里上下键能在 UpDown 上改值；HOTTRACK：悬停高亮（VB6 的观感）。
    long style = kUdsArrowKeys | kUdsHotTrack;
    auto ud = [&](const char* k) -> const FrmValue* {
        auto it = ctrl.properties.find(k);
        return it == ctrl.properties.end() ? nullptr : &it->second;
    };
    const FrmValue* v = ud("Wrap");
    if (v && v->intValue) style |= kUdsWrap;
    v = ud("Alignment");
    if (v) style |= (v->intValue != 0) ? kUdsAlignRight : kUdsAlignLeft;
    else   style |= kUdsAlignRight;
    v = ud("Orientation");
    if (v && v->intValue != 0) style |= kUdsHorz;
    v = ud("SyncBuddy");
    if (v && v->intValue) style |= kUdsSetBuddy;
    return style;
}

// C29-SL-l（账 #143）: 把一条“控件对象表达式”折成 (控件类型, 那枚窗口的 C 表达式)。
// 三种形态都要认，因为控件方法的发码有两条码头：
//   · 表达式路给 `vb6_hwnd_List1`（裸槽，visit(IndexOrCallExpr) 的 marker）；
//   · 语句路给裸小写名 `list1`（`List1.Clear` 那族已有的几支就按这个查）；
//   · 控件数组给 `vb6_CtrlArr_GetAt(&vb6_hwnd_txtSearch, (Index))` —— GetAt 回的就是那枚
//     HWND（`vb6forms_ctrlarr.c` 里存的就是句柄），所以整条表达式直接当窗口用。
// 认不出（不是控件、或者槽后面还挂着成员/下标）就回 false，让调用方落回原来的兜底。
bool CCodeGen::formCtrlSlot(const std::string& objExpr, FrmControlType& outType,
                            std::string& outHwndExpr) const {
    static const std::string kSlot = "vb6_hwnd_";
    static const std::string kArr  = "vb6_CtrlArr_GetAt(&vb6_hwnd_";
    std::string bare;
    if (objExpr.compare(0, kArr.size(), kArr) == 0) {
        size_t e = objExpr.find(',', kArr.size());
        if (e == std::string::npos) return false;
        bare = objExpr.substr(kArr.size(), e - kArr.size());
        outHwndExpr = objExpr;
    } else if (objExpr.compare(0, kSlot.size(), kSlot) == 0 && objExpr.size() > kSlot.size()) {
        if (objExpr.find_first_of("[(.", kSlot.size()) != std::string::npos) return false;
        bare = objExpr.substr(kSlot.size());
        outHwndExpr = objExpr;
    } else if (objExpr.find_first_of("[(.") == std::string::npos) {
        bare = objExpr;
    } else {
        return false;
    }
    std::string lower = Symbol::toLower(bare);
    auto it = knownFormControls_.find(lower);
    if (it == knownFormControls_.end()) return false;
    outType = it->second;
    if (outHwndExpr.empty()) {
        auto org = knownFormControlOriginalNames_.find(lower);
        outHwndExpr = kSlot + cIdent(org != knownFormControlOriginalNames_.end()
                                        ? org->second : bare);
    }
    return true;
}

// 账 #278 §B72: 表交名字，也交「这条出口在 C 里收几枚实参」—— 两者必须同一行回答。
// 以前只有名字，个数住在各条码头的参数串里，于是 RTL 与发码点之间只剩「头 vs 体」这一头
// 有人对账 (账 #240 的 check_rtl_proto_arity.ps1)：谁给某枚出口加一枚形参、只改头与体，
// 发码仍递旧的个数，本机只是 warning C4020、新 cl 才升成 error C2197 —— 门红在别人刚登记
// 的用例上。有了这一行，同一道哨兵可以三面一起问：表 ↔ RTL 原型 ↔ 产物里实际的调用。
static std::string controlExit(const char* rtlName, int argc, int* outArgc) {
    if (outArgc) *outArgc = argc;
    return rtlName;
}

// C29-SL-l（账 #143）: 控件的**零实参方法**名表（命中回 C 里的函数名，否则空串）。
// 焦点面只给"真能拿焦点"的那批窗口型控件 —— Label / Image / Shape / Line / Timer / Menu /
// Data / OLE / CommonDialog / Winsock / ImageList / StatusBar / ProgressBar 刻意不接：
// 真 VB6 在那里是 raise 一个错误号，而本项目还没有运行期错误面，"什么都不做"比伪造成功诚实。
std::string CCodeGen::controlZeroArgMethod(FrmControlType ctrlType,
                                          const std::string& memberLower,
                                          int* outArgc) const {
    if (memberLower == "setfocus") {
        switch (ctrlType) {
            case FrmControlType::CommandButton:
            case FrmControlType::TextBox:
            case FrmControlType::CheckBox:
            case FrmControlType::OptionButton:
            case FrmControlType::ListBox:
            case FrmControlType::ComboBox:
            case FrmControlType::PictureBox:
            case FrmControlType::HScrollBar:
            case FrmControlType::VScrollBar:
            case FrmControlType::DriveListBox:
            case FrmControlType::DirListBox:
            case FrmControlType::FileListBox:
            case FrmControlType::Slider:
            case FrmControlType::TreeView:
            case FrmControlType::ListView:
            case FrmControlType::Toolbar:
            case FrmControlType::SSTab:
            case FrmControlType::DTPicker:
            case FrmControlType::MonthView:
            case FrmControlType::RichTextBox:
                return controlExit("vb6_SetControlFocus", 1, outArgc);
            default:
                return "";
        }
    }
    if (memberLower == "clearsel" && ctrlType == FrmControlType::Slider)
        return controlExit("vb6_intrSlider_ClearSel", 1, outArgc);
    // ⚠ `.Cls` **不登记在本表**：它属于画布家族 (Cls / Print / Line)，出口名只许出自
    // src/common/canvas_drawing.hpp 那张表 (哨兵 form_draw_state S9 + rtl_proto_arity
    // R3/R5 都钉着这一条)。With 块里的 `.Cls` 走 `controlCanvasMethod`，见
    // cgen_expr_with.cpp 的 FormControl 臂 (Fix 263)。
    // 账 #192: ListBox / ComboBox 的 Clear。VB6 里它是方法而不是属性，且只有这两枚
    // 有清空语义（TreeView/ListView 的清是各自那一族，另有出口）。
    if (memberLower == "clear"
        && (ctrlType == FrmControlType::ListBox || ctrlType == FrmControlType::ComboBox))
        return controlExit("vb6_ClearList", 1, outArgc);
    // 账 #278 §B72 第十四刀: 集合那一族的 Clear。两枚的**接收者表达式不一样** —— ImageList 的
    // 槽位是 `vb6_com_<名>`（它是宿主对象不是窗口），StatusBar 才是 `vb6_hwnd_<名>` —— 所以表管
    // "叫什么、递几枚"，"递给谁"仍由码头拼。这一刀之前，这两枚名字在两条码头各抄一遍。
    if (memberLower == "clear" && ctrlType == FrmControlType::ImageList)
        return controlExit("vb6_ImageList_ClearImages", 1, outArgc);
    if (memberLower == "clear" && ctrlType == FrmControlType::StatusBar)
        return controlExit("vb6_StatusBar_ClearPanels", 1, outArgc);
    // C30-2: TabStrip 的 Tabs.Clear。
    if (memberLower == "clear" && ctrlType == FrmControlType::TabStrip)
        return controlExit("vb6_TabStrip_TabsClear", 1, outArgc);
    // CommonDialog 那六枚: 从前两条码头各自做一遍**字符串手术**
    // （`"vb6_CdShow" + toupper(m[4]) + m.substr(5)`）—— 命名规则住在调用点，正是 §B115 那句
    // "拼法只许来自表"要禁的形状；而且它对 `mCd[4]` 的依赖意味着名单加一枚就得信拼法还对。
    if (ctrlType == FrmControlType::CommonDialog) {
        if (memberLower == "showopen")    return controlExit("vb6_CdShowOpen", 1, outArgc);
        if (memberLower == "showsave")    return controlExit("vb6_CdShowSave", 1, outArgc);
        if (memberLower == "showcolor")   return controlExit("vb6_CdShowColor", 1, outArgc);
        if (memberLower == "showfont")    return controlExit("vb6_CdShowFont", 1, outArgc);
        if (memberLower == "showprinter") return controlExit("vb6_CdShowPrinter", 1, outArgc);
        if (memberLower == "showabout")   return controlExit("vb6_CdShowAbout", 1, outArgc);
    }
    return "";
}

// 账 #196（§B31 剩下的一半）: 控件的**一个实参方法**名表 —— VB6 的 TextHeight/TextWidth。
// 与 controlZeroArgMethod 同一套规矩：**表只交名字，实参由码头拼**（三条码头都只有成员名与
// 接收者，实参表在调用点手里）。档位同样**刻意不给通用行** —— VB6 只在画得上去的那几枚上
// 有文字量（Form / PictureBox / UserControl / PropertyPage / Printer），给成通用 ⇒
// `List1.TextHeight("x")` 也答一个数就是伪造成功（同 #192 的 CurrentX、#196 的 hDC）。
// UserControl 那一档早就有（#177/#178 按实例那对），Printer 有 vb6_Printer_*，这里补的是
// 窗体与 PictureBox —— 语料物证 ucTreeMaps 的 PropPagFMR.pag:265 `With Picture1 : .TextHeight(Text)`。
std::string CCodeGen::controlOneArgMethod(FrmControlType ctrlType,
                                          const std::string& memberLower,
                                          int* outArgc) const {
    // C30-UD-b: UpDown 的判据方法 SimDelta(delta) —— 造一条**真** UDN_DELTAPOS 递给父窗
    // （与 Slider.SimNotify / DTPicker.SimChange 同先例，非 VB6 语义，不进手册方法表）。
    // 它正好是“一个整型实参”，落在本表：带括号那形由 cgen_expr_call_com_bind.inc 拼成
    // `vb6_UpDown_SimDelta((void*)hwnd, delta)`。**不带括号的语句形没接管** —— 语句路不查
    // 这张表，`ud1.SimDelta 1` 会退化成“取属性 + 丢结果”，所以夹具只用带括号那形。
    if (ctrlType == FrmControlType::UpDown && memberLower == "simdelta")
        return controlExit("vb6_UpDown_SimDelta", 2, outArgc);
    if (memberLower != "textheight" && memberLower != "textwidth") return "";
    switch (ctrlType) {
        case FrmControlType::Form:
        case FrmControlType::PictureBox:
            return memberLower == "textheight"
                 ? controlExit("vb6_ControlTextHeight", 2, outArgc)
                 : controlExit("vb6_ControlTextWidth", 2, outArgc);
        default:
            return "";
    }
}

// 账 #196 第三条: 控件的**单位换算方法**名表 —— VB6 的 ScaleX/ScaleY(x, fromScale, toScale)。
// 与 controlOneArgMethod 同一套规矩：表只交名字、实参由码头拼；档位同样**刻意不给通用行** ——
// VB6 只有"自己有 ScaleMode 的那些对象"才有这一对（Form / PictureBox / UserControl / PropertyPage /
// Printer），给成通用行就等于允许 `List1.ScaleX(...)` 也答一个数（伪造成功，同 #192/#196 那条口径）。
// 名字不带宿主前缀是故意的：这四形接收者（显式控件、`Me.`、With 块里那枚、UC/页里裸写）要的换算
// 只吃那两个显式的 from/to，实现只有一份 `vb6_ScaleUnitX/Y`（UC 那一档另有一层同名转手，
// 因为宿主伪成员表的命名契约是 `vb6_<Host>_<Member>`）。
std::string CCodeGen::controlScaleMethod(FrmControlType ctrlType,
                                         const std::string& memberLower,
                                         int* outArgc) const {
    if (memberLower != "scalex" && memberLower != "scaley") return "";
    switch (ctrlType) {
        case FrmControlType::Form:
        case FrmControlType::PictureBox:
            return memberLower == "scalex"
                 ? controlExit("vb6_ScaleUnitX", 3, outArgc)
                 : controlExit("vb6_ScaleUnitY", 3, outArgc);
        default:
            return "";
    }
}

// 账 #278 §B72 第十三刀: Winsock 那一族的名字表（八枚成员 → RTL 出口 + 那条出口收几枚实参）。
// 这一族从前把「有哪些成员」与「各自叫什么」抄在两处 —— 表达式码头 (withm) 里一张 kWsKnown
// 名单外加八条分支里的字面量，语句码头 (cgen_call) 里又手写三枚。名单与名字分家的下场就是
// 本账反复量到的那一味：加一档要改三处，漏一处的症状是编得过、跑起来一声不响。
// 个数取自 RTL 原型（vb6forms_prop_ctrl.h:821-828，第一枚一律是 void* hwnd）：
//   close / listen / connect 1、accept 2、bind 3、senddata 2、getdata / peekdata 4。
// 与前几张表同一套规矩：**表交名字与个数，实参由码头拼** —— 出参取址（&变量）、(int32_t)
// 强制、L"" 兜底那些形状是调用点的知识，不搬进表里。
std::string CCodeGen::controlWinsockMethod(const std::string& memberLower,
                                           int* outArgc) const {
    if (memberLower == "close")    return controlExit("vb6_Ws_Close", 1, outArgc);
    if (memberLower == "listen")   return controlExit("vb6_Ws_Listen", 1, outArgc);
    if (memberLower == "connect")  return controlExit("vb6_Ws_Connect", 1, outArgc);
    if (memberLower == "accept")   return controlExit("vb6_Ws_Accept", 2, outArgc);
    if (memberLower == "bind")     return controlExit("vb6_Ws_Bind", 3, outArgc);
    if (memberLower == "senddata") return controlExit("vb6_Ws_SendData", 2, outArgc);
    if (memberLower == "getdata")  return controlExit("vb6_Ws_GetData", 4, outArgc);
    if (memberLower == "peekdata") return controlExit("vb6_Ws_PeekData", 4, outArgc);
    return "";
}
// 账 #221 = C29-PL-a（语料物证 Charts 2020/ucProgressCircular 的 ppProgressCircular.pag
// 那 8 条 `Picture1.Line` / `Picture2.Line`）: 画表面方法的名字表。
// 这一族以前**没有表** —— parser 的 Fix 102 把坐标对折进实参表之后，注释写着"由后端按控件
// 类型发射"，但后端从来没接这一刀，于是整条调用落到通用 COM 兜底
// （`ComGetObjectProp(hwnd, L"Line")` 再对它取 `Item`，两跳都被 RTL 登记成"认识但什么都不做"），
// 症状是**编得过、跑得起、一笔不画**。档位与 TextHeight/ScaleX 两族同样只给 Form 与 PictureBox。
// 账 #232②: 这张表从此是画布家族**名字**的唯一出口 —— cls / print / line 三档。
// 以前 cls 与 print 不在表里，而是三条码头各自硬编码了一遍 `== "print" || == "cls"`
// 且每一处都只认 PictureBox（成员侧打标记、表达式码头、语句码头各抄一份），于是
// `Me.Cls` / `Me.Print "AB"` 两形压根没人接 —— 实测发成
// `vb6_ComCall(vb6_hwnd_<窗体>, L"Cls", NULL, 0)` 与
// `vb6_ComCallObject(vb6_ComGetObjectProp(同一 HWND, L"Print"), L"Item", …)`：
// 窗体句柄不是 IDispatch ⇒ 编得过、跑得起、一笔不画（本线第四次栽在同一味上）。
// 窗体自己那枚接收者是认得的（`cgen_form_ctrl_registry.inc` 把窗体名也登记进
// knownFormControls_，类型 Form），缺的只有这两行。
// line 那一档今天三族接收者都在表里（账 #232① 把 Form/Printer 两行搬进来的），但 Form 那一行是
// owner=DRAW（12 参 `vb6_Form_Line`，带 Step 相对位）⇒ 归绘图码头答，从这条 METHOD 出口问它是空串。
//
// 账 #232①-c: 表的**本体**（动词名 / 哪些接收者有这一档 / 这一档归谁发码）搬到
// `src/common/canvas_drawing.hpp`，因为语义层（裸写的 `Cls` / `PSet` 折成 `Me.<动词>`）也要问同一张表，
// 而 common 是两层唯一都能问的家（与 host_pseudo.hpp / float_literal.hpp 同族）。
// 这里只留**两份** AST 侧才有的知识：`FrmControlType` 到接收者旗标的映射，以及"这条码头问的是
// METHOD 那一族"。表里 owner=DRAW 的行（`Me.PSet` / `Me.Line` 在窗体上的 12 参那一档）由绘图码头答，
// 从这条出口问必须是空串 —— 两条码头各答一半，正是本账 ② 之前的形状。
// 比对口径原样保留（表按拿进来的字符串逐字比，不在这个出口里擅自 toLower —— 两条码头给的拼写
// 本来就不一样，改它是另一刀，得单独有用量）。
std::string CCodeGen::controlCanvasMethod(FrmControlType ctrlType,
                                          const std::string& memberLower) const {
    uint32_t receiver = 0;
    switch (ctrlType) {
        case FrmControlType::Form:       receiver = CR_CANVAS_FORM; break;
        case FrmControlType::PictureBox: receiver = CR_CANVAS_PBOX; break;
        default: return "";   // 没有画布的控件: 本表一律不答
    }
    return canvasMethodEntry(memberLower, receiver);
}


// C29-SL-n（账 #141）: 「这枚控件要不要子类化」的唯一一份判据 —— 内容与
// `cgen_form_wndproc_subclass.inc` 汇总 info.hasXxx 那一趟逐条对应（改一边就得改另一边，
// 否则又回到"arm 发了、没人 install"那一形）。装的那趟在 `cgen_form_frame_menu.inc`、
// 拆的那趟在 `cgen_form_wndproc_dispatch.inc`，两边各自抄了一份，实测**两份都漏了
// `_DblClick` 与 `_Paint`** ⇒ 只挂这两条处理器之一的控件，子类过程与消息臂都生成得好好的，
// 一次也没被 install（处理器编得过、永不触发）。
// 账 #189 (B29②): 控件数组的**整体成员** —— VB6 里 `arr.Count / arr.LBound / arr.UBound`
// 问的是数组本身, 不是某一枚控件的属性。以前只有 LBound/UBound 在成员读取那一路各写了一条 if,
// Count 漏了 ⇒ 掉进 COM 兜底, 发成 vb6_ComGetIntProp(vb6_hwnd_arr1, L"Count")，
// 而数组控件根本没有 vb6_hwnd_<数组名> 这个变量 (C2065，实测 Charts 2020/ucChartBar Form2)。
// 与账 #157 同族：句柄类表达式必须走 vb6_arr_*，不许凭空拼 vb6_hwnd_。
std::string CCodeGen::ctrlArrayMetaMemberExpr(const std::string& arrName,
                                              const std::string& memberLower) const {
    std::string arg = "&vb6_arr_" + cIdent(arrName);
    if (memberLower == "count") return "vb6_CtrlArr_GetCount(" + arg + ")  /* ctrl array Count */";
    if (memberLower == "lbound") return "vb6_CtrlArr_LBound(" + arg + ")  /* ctrl array LBound */";
    if (memberLower == "ubound") return "vb6_CtrlArr_UBound(" + arg + ")  /* ctrl array UBound */";
    return "";
}

// 账 #190 (B29⑤): 控件事件处理器的 C 函数名 —— 唯一出口。
// VB6 的标识符大小写不敏感, 而 C 敏感: 过程定义发的是 **Sub 自己的拼写**
// (`vb6_Form_<模块>_<Sub名>`), 所以臂里调用的也必须是那一个名字。以前这一手是拿
// **控件的设计期名**现拼的 (`cProcName(ctrl.controlName + "_Click")`), 于是只要有人改了控件名
// 而没改过程名 (VB6 里完全合法, 两枚照样配一对), 产物就是"引用一个没人定义的函数" ——
// 链接期 LNK2019 (实测 Charts 2020/ucChartBar 的 Form1: 控件 ChkAxisY / 过程 ChkAxisy_Click,
// 控件 cboLabelsPositions / 过程 CboLabelsPositions_Click, 正好 2 个无法解析的外部符号)。
// 存在性那一步 (symTab_.lookup) 本来就是大小写无关的 —— 缺的只是"命中之后按谁的名字发"。
// 返回空串 = 这个事件没有处理器, 调用方**不要**装这条臂。
std::string CCodeGen::eventHandlerFn(const std::string& ctrlName,
                                     const std::string& suffix) const {
    auto* sym = symTab_.lookup(ctrlName + suffix);
    if (!sym) return std::string();
    return cProcName(sym->name, AccessLevel::Private);
}

bool CCodeGen::controlNeedsSubclass(const FrmControl& ctrl) const {
    auto has = [&](const char* ev) {
        return symTab_.lookup(ctrl.controlName + std::string(ev)) != nullptr;
    };
    // 焦点：只有一枚控件**没有**另一条原生来源（父窗 WM_COMMAND 那批码）时才由子类过程供，
    // 否则同一次焦点变化会两边各发一次（C29-SL-g）。
    if (!controlFocusFromNativeNotify(ctrl.controlType)
        && (has("_GotFocus") || has("_LostFocus"))) return true;
    if (has("_MouseEnter") || has("_MouseLeave") || has("_MouseHover")
        || has("_MouseDown") || has("_MouseUp") || has("_MouseMove")
        || has("_KeyPress") || has("_KeyDown") || has("_KeyUp")
        || has("_Validate")) return true;
    // C29-SL-n: 这两条以前只有发 arm 那份认，装/拆两份都漏。
    if (has("_DblClick")) return true;
    // VB6 里有绘制表面的控件才有 Paint 语义，本线只接 PictureBox（Fix 185 的口径，与 arm 那趟一致）。
    if (ctrl.controlType == FrmControlType::PictureBox && has("_Paint")) return true;
    // 同上：`_Click` 只在没有原生 Click 来源时才由子类过程补（C29-SL-d）。
    if (!controlClickFromNativeNotify(ctrl.controlType) && has("_Click")) return true;
    return false;
}

// C29-SL-o（账 #83 的 Slider 半边）: Slider 的设计期值面 —— 两条创建路共用这一处。
// 参数顺序不是随便排的，是 .build/slprobe/slmeasure6.c 量出来的：改 range 会把 pos 顶到新下限、
// 把 page 从 20 重算成 18、把 selstart 跟到新下限 ⇒ range → line/page → pos → 刻度 → Sel
// 这个序，任何一端先动都会把后面的吃掉。-999 哨兵 = .frm 没写过 ⇒ 那一条消息不发（保持控件默认），
// 而 Init 本身照发 —— 发码形状稳定，emitc 断言要看它（同 vb6_TreeView_Init 那条）。
void CCodeGen::emitSliderDesignTimeInit(const FrmControl& ctrl, const std::string& hwndExpr) {
    auto slProp = [&](const char* key) -> std::string {
        auto it = ctrl.properties.find(key);
        if (it == ctrl.properties.end()) return "-999";
        return intLiteralText(static_cast<int64_t>((int)it->second.intValue), true);
    };
    // C29-SL-i: VB6 那一面的选区是 **SelStart + SelLength**（类型库 dispid 0x0007/0x0008），
    // SelEnd 是我们 SL-b 按原生 TBM_SETSELEND 自己加的名字。所以 .frm 里只写了前两条时
    // 这里折一次：终点 = 起点 + 长度 —— 否则真 VB6 工程的 SelLength 在设计期就被整条丢掉
    // （运行期那条写口一直是好的，缺口只在创建那一趟）。
    // 只在**两条都写了**的时候折：起点没写就不知道拿什么当锚，宁可不折也不猜一个起点出来下发。
    std::string slEndArg = slProp("SelEnd");
    if (slEndArg == "-999") {
        auto ssIt = ctrl.properties.find("SelStart");
        auto slIt = ctrl.properties.find("SelLength");
        if (ssIt != ctrl.properties.end() && slIt != ctrl.properties.end()) {
            slEndArg = intLiteralText(static_cast<int64_t>((int)ssIt->second.intValue
                                      + (int)slIt->second.intValue), true);
        }
    }
    c_.emitLine("vb6_Slider_Init((void*)" + hwndExpr + ", "
                + slProp("Min") + ", " + slProp("Max") + ", " + slProp("Value") + ", "
                + slProp("SmallChange") + ", " + slProp("LargeChange") + ", "
                + slProp("TickFrequency") + ", " + slProp("SelStart") + ", "
                + slEndArg + ", " + slProp("SelectRange") + ");");
}

// C30-UD-a: UpDown 设计期那一趟。参数序与 -999 哨兵口径只在 RTL 那一处定，这里只负责转述。
// BuddyControl 在 .frm 里写的是**控件名字符串**（FrmValue 的 rawText 带引号）⇒ 这里折成
// 那枚控件的句柄槽。折不出来（名字空 / 写了个不存在的控件）就传 NULL：宁可不挂 buddy，
// 也不猜一个句柄出去。
void CCodeGen::emitUpDownDesignTimeInit(const FrmControl& ctrl, const std::string& hwndExpr) {
    auto udProp = [&](const char* key) -> std::string {
        auto it = ctrl.properties.find(key);
        if (it == ctrl.properties.end()) return "-999";
        return intLiteralText(static_cast<int64_t>((int)it->second.intValue), true);
    };
    std::string buddy = "NULL";
    auto bcIt = ctrl.properties.find("BuddyControl");
    if (bcIt != ctrl.properties.end()) {
        std::string nm = bcIt->second.rawText;
        if (nm.size() >= 2 && nm.front() == '"' && nm.back() == '"') {
            nm = nm.substr(1, nm.size() - 2);
        }
        if (!nm.empty()) buddy = "(void*)vb6_hwnd_" + cIdent(nm);
    }
    c_.emitLine("vb6_UpDown_Init((void*)" + hwndExpr + ", "
                + udProp("Min") + ", " + udProp("Max") + ", " + udProp("Value") + ", "
                + udProp("Increment") + ", " + udProp("Wrap") + ", "
                + udProp("Alignment") + ", " + udProp("Orientation") + ", "
                + udProp("SyncBuddy") + ", " + buddy + ");");
}

// 账 #83(a)（C29-SL-r）: VB6 的 `TabStop` 默认 True,而本项目**两条创建路以前都不立 WS_TABSTOP**
// —— 探针 `.build/sltab` 实测：顶层按钮（`.frm` 没写 TabStop）读回 0、Frame 里的按钮与文本框也读回 0，
// 只有"`.frm` 写了 `TabStop = 0`"那一枚碰巧对（因为它要的就是 0）。RTL 那边其实一直按"默认 True"写的
// （`vb6_GetTabStop` 里 `!hwnd` 就回 -1），缺的只是创建时把这一位立上。
// 读侧就是 `GetWindowLong(GWL_STYLE) & WS_TABSTOP` ⇒ 问的是窗口自己，我们没有另存一份。
// 排除的是拿不到焦点的那几类；`Unknown`（uc 实例与没登记的 OCX）也不立 —— 那些可能压根没有窗口。
// 账 #164：`PictureBox` 也进排除表。裸码实测 `.build/picstop/picstop.c`：同一棵里
// `STATIC` 挂上 `WS_TABSTOP` 就会被对话框管理器当成一站（`B1 → Static'PIC' → B2`），
// 摘掉这一位就变成 `B1 → B2` —— 机制全在这一位上，不用碰派发。
// 上面那个"写了照发"的分支仍优先：`.frm` 真写了 `TabStop` 就按写的来。
long CCodeGen::controlTabStopStyleBit(const FrmControl& ctrl) const {
    constexpr long kWsTabStop = 0x00010000L;
    auto tsIt = ctrl.properties.find("TabStop");
    if (tsIt != ctrl.properties.end() && tsIt->second.type == FrmValueType::Integer) {
        return tsIt->second.intValue != 0 ? kWsTabStop : 0L;
    }
    switch (ctrl.controlType) {
    case FrmControlType::Label:
    case FrmControlType::Image:
    case FrmControlType::Shape:
    case FrmControlType::Line:
    case FrmControlType::PictureBox:  // 账 #164: VB6 的 PictureBox 拿不到焦点，不该进 tab 序
    case FrmControlType::Frame:
    case FrmControlType::Timer:
    case FrmControlType::Menu:
    case FrmControlType::Data:
    case FrmControlType::OLE:
    case FrmControlType::ImageList:
    case FrmControlType::CommonDialog:
    case FrmControlType::Form:
    case FrmControlType::MDIForm:
    case FrmControlType::Unknown:
        return 0L;
    default:
        return kWsTabStop;
    }
}

// 账 #167：容器清单的**唯一出口**。这一份以前在 backend 里有七处抄本 ——
//   创建侧三条（发子控件的递归、顶层那趟入口、装 `vb6_ForwardChildCommands` 那一趟）都含 SSTab，
//   而派发侧四条（`walkCmdChildren142` 两处、`walkFocusChildren158` 两处）只认 Frame/PictureBox。
//   ⇒ 任何把控件放进 SSTab 页里的工程，派发侧的 DFS 编号整体错位：实测（`.build/ss167`，
//   SSTab 排在 Frame 前面）给 `cmdInTab` 焦点，跑的是 `cmdInFrame` 的 `_GotFocus`
//   （`SS167-tabGot=0/frameGot=1`）—— 不是「没人发」，是「发给别人」。
//   通用式：判据抄第二遍就会漂，漂的方向还不唯一（这边漏 SSTab、那边漏 `_DblClick`/#141、
//   那边漏 `_Paint`）⇒ 收成一处、census 一遍、再留一条能红的针。
// 账 #83(b2): 容器窗口挂 `WS_EX_CONTROLPARENT`，对话框管理器才肯走进它；
// 两条创建路都要吃这个出口（容器嵌容器走第二条路）—— "只接一头"是本线踩过多次的那一声不响。
bool CCodeGen::controlIsContainerType(const FrmControl& ctrl) const {
    switch (ctrl.controlType) {
    case FrmControlType::Frame:
    case FrmControlType::PictureBox:
    case FrmControlType::SSTab:
        return true;
    default:
        return false;
    }
}

bool CCodeGen::formTreeHasEventLoadingUsers(const FrmControl& root) const {
    for (const auto& child : root.children) {
        switch (child.controlType) {
        case FrmControlType::CheckBox:
        case FrmControlType::OptionButton:
        case FrmControlType::HScrollBar:
        case FrmControlType::VScrollBar:
            return true;
        default:
            break;
        }
        // 容器里的子控件也算（TeeChart 示例 8 枚 HScrollBar 全在 Frame8 里）。
        if (!child.children.empty() && formTreeHasEventLoadingUsers(child)) return true;
    }
    return false;
}

long CCodeGen::controlContainerExStyleBit(const FrmControl& ctrl) const {
    // ⚠ 账 #165 的根因就在这个数上：`WS_EX_CONTROLPARENT` 在 SDK 头里是 **0x00010000**
    //   （winuser.h:2855），而 0x00040000 是 `WS_EX_APPWINDOW`（同表往下 2857 行）。以前这里
    //   手抄成了后者 ⇒ 那两轮的读数只证明"想发的数落到窗口上了"，证不了"那是管理器认的那一位"。
    //   通用式：**手抄常量一律去 SDK 头对值**（或直接引符号），证人行打印符号对应的值。
    constexpr long kWsExControlParent = 0x00010000L;  // WS_EX_CONTROLPARENT（winuser.h 实测值）
    // 残差 #5 的另一半：VB6 的 ThunderRT6TextBox 实测 ex=0x204（WS_EX_CLIENTEDGE | NOPARENTNOTIFY），
    // 边框是**有 3D 内凹**的（左/上 236 亮边 + 下/右 131 暗边、客户区 87x29 比控件小 2px/边）；只去掉
    // WS_BORDER 会连这条内凹一起丢掉，于是「深框没了、浅边也没了」。WS_EX_CLIENTEDGE=0x200（winuser.h）。
    // NOPARENTNOTIFY 只影响父窗通知、与外形无关，这里不立（少一位就少一条看不见的副作用）。
    constexpr long kWsExClientEdge = 0x00000200L;     // WS_EX_CLIENTEDGE（winuser.h）
    if (controlIsContainerType(ctrl)) return kWsExControlParent;
    if (ctrl.controlType == FrmControlType::TextBox) return kWsExClientEdge;
    return 0L;
}

// 账 #157: 为什么这一份留在发码期算，声明处的注释有交代。这里只做**选择**，并把选中那枚的句柄
// 变量名交给窗体的 WM_CREATE 发一句 `vb6_Form_SetInitialFocus`。
// 选择口径 = VB6：`TabIndex` 最小、且拿得到焦点（`controlTabStopStyleBit` 那一族排除 +
// 显式 `TabStop = False` 不算）、设计期没被藏起来 / 没被禁用的那枚；同序号按创建顺序取先。
void CCodeGen::emitFormInitialFocus(const FrmControl& formNode) {
    const FrmControl* best = nullptr;
    long bestTabIndex = 0;
    long bestOrder = 0;
    long order = 0;

    std::function<void(const FrmControl&)> walk;
    walk = [&](const FrmControl& node) {
        for (const auto& ctrl : node.children) {
            const long orderHere = order++;
            bool takesFocus = controlTabStopStyleBit(ctrl) != 0;
            auto visIt = ctrl.properties.find("Visible");
            if (visIt != ctrl.properties.end() && visIt->second.type == FrmValueType::Integer
                && visIt->second.intValue == 0) takesFocus = false;
            auto enIt = ctrl.properties.find("Enabled");
            if (enIt != ctrl.properties.end() && enIt->second.type == FrmValueType::Integer
                && enIt->second.intValue == 0) takesFocus = false;
            long tabIndex = 0;
            auto tiIt = ctrl.properties.find("TabIndex");
            if (tiIt != ctrl.properties.end() && tiIt->second.type == FrmValueType::Integer)
                tabIndex = tiIt->second.intValue;
            if (takesFocus && (!best || tabIndex < bestTabIndex
                               || (tabIndex == bestTabIndex && orderHere < bestOrder))) {
                best = &ctrl;
                bestTabIndex = tabIndex;
                bestOrder = orderHere;
            }
            // Frame / PictureBox 里的子控件也在这枚窗体的 tab 序里，所以容器本身被排除掉
            // 之后仍要继续往里走。
            walk(ctrl);
        }
    };
    walk(formNode);
    if (!best) return;

    // 句柄表达式必须走 `ctrlHwndExprForInit` —— 控件数组（如 txtSearch(0)）的句柄在
    // `vb6_arr_<名>` 里，硬写 `vb6_hwnd_<名>_0` 会引用一个不存在的全局（NewTab 实测：
    // error C2065 未声明的标识符 'vb6_hwnd_txtSearch_0'）。
    c_.emitLine("vb6_Form_SetInitialFocus((void*)hwnd, (void*)" + ctrlHwndExprForInit(*best)
                + ");  /* 账 #157: 显示时把焦点交给 " + best->controlName
                + "（TabIndex=" + std::to_string(bestTabIndex) + "） */");
}

// 控件类型的 Win32 样式位。取值与 cgen_form_ctrl_style_apply.inc 保持一致。
// 顶层控件与容器子控件共用，避免容器内子控件缺失类型样式。
long CCodeGen::controlTypeStyleBits(const FrmControl& ctrl) const {
    constexpr long kWsBorder   = 0x00800000L;
    constexpr long kBsPush     = 0x00000000L;
    constexpr long kBsAutoChk  = 0x00000003L;
    constexpr long kBsAutoRad  = 0x00000009L;
    constexpr long kBsGroupBox = 0x00000007L;
    constexpr long kSsLeft     = 0x00000000L;
    constexpr long kSsNotify   = 0x00000100L;  // SS_NOTIFY: Label 接收鼠标消息 (tooltip/Click)
    constexpr long kEsAutoH    = 0x00000080L;
    constexpr long kLbsNotify  = 0x00000001L;
    constexpr long kCbsDrop    = 0x00000002L;
    constexpr long kSsBitmap   = 0x0000000EL;
    constexpr long kSsCenterImg= 0x00000200L;
    constexpr long kBsPushLike = 0x00001000L;
    constexpr long kEsMulti    = 0x00000004L;
    constexpr long kEsAutoV    = 0x00000040L;
    constexpr long kWsHscroll  = 0x00100000L;
    constexpr long kWsVscroll  = 0x00200000L;

    long style = 0;
    switch (ctrl.controlType) {
        case FrmControlType::CommandButton: {
            style |= kBsPush;
            auto it = ctrl.properties.find("Style");
            if (it != ctrl.properties.end() && it->second.intValue == 1) style |= kBsPushLike;
            break;
        }
        case FrmControlType::TextBox: {
            // VB6 ThunderRT6TextBox: style 无 WS_BORDER、ex 有 WS_EX_CLIENTEDGE；
            // 原生 EDIT 立 WS_BORDER 会被主题 EDIT 画成 1px `100,100,100` 深框 (见顶层同处)
            style |= kEsAutoH;
            auto mlIt = ctrl.properties.find("MultiLine");
            if (mlIt != ctrl.properties.end() && mlIt->second.intValue != 0) style |= kEsMulti | kEsAutoV;
            auto sbIt = ctrl.properties.find("ScrollBars");
            if (sbIt != ctrl.properties.end()) {
                int sb = (int)sbIt->second.intValue;
                // 账 #108: 与顶层那条同一处错、同一处修法 —— VB6 是 1 水平 / 2 垂直
                // (证人：tests/VBFlexGridDemo/InputForm.frm 里 VB6 自己存的
                //  `ScrollBars = 2  'Vertical`)。
                if (sb == 1 || sb == 3) style |= kWsHscroll;
                if (sb == 2 || sb == 3) style |= kWsVscroll;
            }
            break;
        }
        // C30-3: MaskEdBox —— 单行 EDIT（与 TextBox 单行那格同一套样式）。掩码不管样式，
        // 掩码由 RTL 的 vb6_MaskEd_* 那层接管（显示串固定位对齐，不需要 ES_CENTER 之类）。
        case FrmControlType::MaskEdBox:
            style |= kEsAutoH;
            break;
        case FrmControlType::VScrollBar: {
            // SBS_VERT：缺这个方向位时 SCROLLBAR 默认走 SBS_HORZ(=0)，微调按钮区整块空白            // （实测 33x31 区 992/1023 纯白、无箭头）；VB6 VScrollBar 实测含此位。
            style |= 0x0001L;
            break;
        }
        case FrmControlType::HScrollBar:
            break;  // SBS_HORZ = 0
        // C29-RT-a: 与顶层那条创建样式**同一口径**（容器子控件走的就是这条路，账 #83）。
        // 枚举按 VB6 文档：0 无 / 1 水平 / 2 垂直 / 3 两者 —— 与上面的 TextBox 那一格现在一致
        // （账 #108 之前 TextBox 把 1/2 用反了，两条控件同名不同向）。
        case FrmControlType::RichTextBox: {
            style |= kEsMulti;
            auto sbIt = ctrl.properties.find("ScrollBars");
            if (sbIt != ctrl.properties.end()) {
                int sb = (int)sbIt->second.intValue;
                if (sb == 1 || sb == 3) style |= kWsHscroll;
                if (sb == 2 || sb == 3) style |= kWsVscroll;
                if (sb != 0) style |= 0x00002000L;   // ES_DISABLENOSCROLL（与顶层那条同）
            }
            break;
        }
        case FrmControlType::Label:
            style |= kSsLeft | kSsNotify;
            break;
        case FrmControlType::CheckBox: {
            style |= kBsAutoChk;
            auto it = ctrl.properties.find("Style");
            if (it != ctrl.properties.end() && it->second.intValue == 1) style |= kBsPushLike;
            break;
        }
        case FrmControlType::OptionButton: {
            style |= kBsAutoRad;
            auto it = ctrl.properties.find("Style");
            if (it != ctrl.properties.end() && it->second.intValue == 1) style |= kBsPushLike;
            break;
        }
        case FrmControlType::Frame:
            style |= kBsGroupBox;
            break;
        case FrmControlType::ListBox: {
            style |= kLbsNotify | kWsBorder | kWsVscroll;
            auto sortIt = ctrl.properties.find("Sorted");
            if (sortIt != ctrl.properties.end() && sortIt->second.intValue != 0) style |= 0x0002L;
            auto msIt = ctrl.properties.find("MultiSelect");
            if (msIt != ctrl.properties.end()) {
                if (msIt->second.intValue == 1) style |= 0x0008L;
                else if (msIt->second.intValue == 2) style |= 0x0800L;
            }
            break;
        }
        case FrmControlType::ComboBox: {
            auto stIt = ctrl.properties.find("Style");
            if (stIt != ctrl.properties.end()) {
                if (stIt->second.intValue == 1) style |= 0x0001L;       // CBS_SIMPLE
                else if (stIt->second.intValue == 2) style |= 0x0003L;  // CBS_DROPDOWNLIST
                else style |= kCbsDrop;
            } else {
                style |= kCbsDrop;
            }
            // Fix 147: VB6 的 ComboBox 下拉列表带垂直滚动条 — 项数超过可见区时
            // 靠它滚动查看全部项. 此前漏了 WS_VSCROLL, 下拉只能看到前几项,
            // 用户以为"选项不全" (实测 32 个主题只能看到 9 个且无法滚动).
            style |= kWsVscroll;
            style |= kWsBorder;
            auto sortIt = ctrl.properties.find("Sorted");
            if (sortIt != ctrl.properties.end() && sortIt->second.intValue != 0) style |= 0x0100L;
            break;
        }
        // FIX-010 (富贵竹): PictureBox / Image 建的是 STATIC 类窗口，不带 SS_NOTIFY 的
        // STATIC 在 WM_NCHITTEST 上答 HTTRANSPARENT，鼠标消息被交给父窗口 ⇒ 挂在它们
        // 身上的子类化 WndProc 一条 WM_LBUTTONDOWN/UP 都收不到。Label 一直带着这一位，
        // 所以「同一窗体里 Label_Click 活着、Image_MouseDown 全哑」正是这一位的对照。
        // 这里是**容器子控件**那条创建路 (cgen_form_frame_menu.inc 调本函数)；顶层那条路
        // 在 cgen_form_ctrl_style_apply.inc 有自己的第二份表，两处必须同口径。
        case FrmControlType::PictureBox:
            style |= kSsBitmap | kSsCenterImg | kWsBorder | kSsNotify;
            break;
        case FrmControlType::Image:
            style |= kSsBitmap | kSsCenterImg | kSsNotify;
            break;
        // C29-SL-h: 此前这条路**压根没有 Slider 这一格** —— 容器里的滑杆连 TBS_AUTOTICKS 都没
        // 挂上（账 #83 那条"顶层专有项没铺到第二条创建路"的一个具体落点）。折算共用
        // sliderStyleBits，不留第二份表。
        case FrmControlType::Slider:
            style |= sliderStyleBits(ctrl);
            break;
        // C30-UD-a: UpDown 也和 Slider 同型 —— 容器里那枚以前连 ARROWKEYS 都没有，
        // 折算共用 updownStyleBits，不留第二份表。
        case FrmControlType::UpDown:
            style |= updownStyleBits(ctrl);
            break;
        default:
            break;
    }
    return style;
}

bool CCodeGen::controlTypeClearsCaption(const FrmControl& ctrl) const {
    return ctrl.controlType == FrmControlType::PictureBox ||
           ctrl.controlType == FrmControlType::Image ||
           // C29-1a: Shape / Line 是自绘控件, 窗口文字没有任何视觉效果, 但留着控件名
           // 当 caption 会让子类化/工具提示那几条路把它当有文本的控件看待。
           ctrl.controlType == FrmControlType::Shape ||
           ctrl.controlType == FrmControlType::Line;
}

// C29-1a: Line 的窗口矩形就是四个端点的包围盒 (单位 = 容器缇值, 与 .frm 存的一致)。
// 两条创建路 (顶层 / 容器子控件) 都调这里, 免得一边算对一边算成默认的 2000x300。
void CCodeGen::lineRectFromEndpoints(const FrmControl& ctrl,
                                     int& l, int& t, int& w, int& h) {
    auto getTw = [&ctrl](const char* key) -> int {
        auto it = ctrl.properties.find(key);
        return (it != ctrl.properties.end()) ? (int)it->second.intValue : 0;
    };
    int x1 = getTw("X1"), y1 = getTw("Y1");
    int x2 = getTw("X2"), y2 = getTw("Y2");
    l = (x1 < x2) ? x1 : x2;
    t = (y1 < y2) ? y1 : y2;
    w = (x1 < x2) ? x2 - x1 : x1 - x2;
    h = (y1 < y2) ? y2 - y1 : y1 - y2;
    if (w < 1) w = 1;
    if (h < 1) h = 1;
}

// C29-1a: 设计期外观属性落到控件上。值为 0 的那些 (Shape=0、FillStyle=0、BorderStyle=0)
// 也照发 —— RTL 侧把这类枚举存成 val+1, 所以 0 不再等于"没设过"。
// C29-1a: 设计期初始化用的句柄表达式 —— 控件数组 (如 ShapeLamp(0)/ShapeLamp(1)) 的
// 句柄在 vb6_arr_<名> 里, 硬写 vb6_hwnd_<名> 会打到空句柄上 (SetProp 静默失败)。
std::string CCodeGen::ctrlHwndExprForInit(const FrmControl& ctrl) const {
    std::string lower = Symbol::toLower(ctrl.controlName);
    if (knownControlArrays_.count(lower)) {
        return "vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", "
             + std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")";
    }
    return "vb6_hwnd_" + cIdent(ctrl.controlName);
}

// C29-SL-d: 这个控件的 `Click` 是不是已经由**原生通知**送进来了。
// 子类化那条路（`cgen_form_wndproc_subclass.inc`）本来没有 Click 这一档 —— 补上之后
// 必须把"已经有别的 Click 来源"的类型排除掉，否则同一次点击会调两次 handler：
// 一条来自原生通知（下表第二列），一条来自控件自己的 WM_LBUTTONUP。
// 逐条来源（都在本文件/cgen_form_wndproc_create.inc/cgen_form_wndproc_dispatch.inc 里）：
//   CommandButton/CheckBox/OptionButton -> WM_COMMAND BN_CLICKED
//   ListBox/FileListBox/DirListBox/DriveListBox/ComboBox -> WM_COMMAND *_SELCHANGE
//   SSTab -> WM_NOTIFY TCN_SELCHANGE（带 PreviousTab 实参，形参表都不一样）
//   Toolbar -> WM_COMMAND（ButtonClick 那一档，见 cgen_form_wndproc_create.inc）
//   Menu -> 菜单命令那条路（cgen_form_menu.cpp），且 Menu 压根不子类化
// 反过来，Label/Image/PictureBox/Frame/TextBox/ScrollBar/Slider 这些**没有**任何
// 原生 Click 通知的，才由子类化那一档补上。
// 账 #222: 控件数组元素的键 —— 一枚数组的 24 枚元素共用一枚事件处理器, 但每枚要自己的
// thunk / sink / 挂接, 而"这枚元素"在发码里只有一个地方能拼出来: 就是这里。prelude 登记、
// 挂接那头查表、两处子类化臂问"Click 是不是已经由 sink 供给", 三处都调它 (以前挂接那侧
// 把 sink 变量名与控件名再拼一遍, 那是第二处拼名)。
std::string CCodeGen::ctrlElemKey(const std::string& ctrlName, int index) {
    std::string k = Symbol::toLower(ctrlName);
    if (index >= 0) k += "#" + std::to_string(index);
    return k;
}

// 账 #229: "`对象.成员` 的对象位是不是一枚窗体控件、什么类型" —— 这一问在类型推断里原本
// 抄成两份，一份只认 `控件名.属性`（P20-42 那条），一份只认 `控件名(i).属性`（C29-1a 那条）。
// 于是数组元素那一形从前一条前面掉下去，撞上"按成员裸名查模块符号"：`Left` 同时是返回 String
// 的 VB 内置函数 ⇒ 被判成 String ⇒ 拼接面不再套数值转换 ⇒ 裸 int 进 BSTR 槽 = 0xC0000005。
// 实测 ve_units: `qq = "Q3-array-left=" & uArr(0).Left` 两架构都崩；同一枚元素的 .Top 只是
// 绕远装箱不崩，非数组的 `uPix.Left` 也正常 —— 症状按"属性名撞不撞内置函数名"分家，很误导。
// 两条对象形态合成这一个出口。返回 false = 对象位不是窗体控件（调用方照旧走原兜底）。
bool CCodeGen::ctrlTypeOfMemberObject(const Expr* obj, FrmControlType& outType) const {
    if (!obj) return false;
    // 账 #249: `Me.<成员>` —— 窗体模块里的 MeExpr 对象位就是这枚窗体**自己**。knownFormControls_
    // 的键是窗体名（cgen_form_ctrl_registry.inc 登记的），而 MeExpr 不带名字，于是这一问以前
    // 一律答"不是控件"，成员名就掉进"按裸名查模块符号"那一档 —— `Left` 撞返回 String 的 VB
    // 内置函数 ⇒ 判成 String ⇒ `&` 拼接面不套数值转换 ⇒ 裸 int 进 BSTR 槽 = 0xC0000005
    // （实测 `Debug.Print "x=" & Me.Left` 两形五样都崩，而 .Top/.Width/.Height 只因不撞名而幸免；
    // 赋给 Long 变量也正常 —— 症状按"属性名撞不撞内置函数名"分家，与 #229 同一族）。
    // 认成 Form 之后，答案仍出自 controlPropType 那一张表，本函数不答任何具体类型。
    if (obj->kind == ASTNodeKind::MeExpr) {
        if (isFormModule_ && !knownFormName_.empty()) {
            outType = FrmControlType::Form;
            return true;
        }
        return false;
    }
    std::string name;
    if (obj->kind == ASTNodeKind::IdentifierExpr) {
        name = Symbol::toLower(static_cast<const IdentifierExpr&>(*obj).name);
    } else if (obj->kind == ASTNodeKind::IndexOrCallExpr) {
        // 控件数组的元素 (`uArr(1).Left`)：对象位是 `名字(下标)`，与单枚同一条规则。
        auto& call = static_cast<const IndexOrCallExpr&>(*obj);
        if (!call.callee || call.callee->kind != ASTNodeKind::IdentifierExpr ||
            call.positional.size() != 1) return false;
        name = Symbol::toLower(static_cast<const IdentifierExpr&>(*call.callee).name);
    } else {
        return false;
    }
    auto it = knownFormControls_.find(name);
    if (it != knownFormControls_.end()) {
        outType = it->second;
        return true;
    }
    // A cross-module form default instance (`frmMain.hWnd`) is not a control, but its
    // properties answer from the very same controlPropType table. The bare-member-name
    // fallback does not know the host, so `hWnd` came back Variant -> the comparison
    // boxed the HWND as VariantObject -> vb6_VarCmpLongEq answered 0 -> `focused`
    // stayed False and every KeyHit was swallowed (VBWolf ignored the keyboard).
    // Non-current form only: same guard as cgen_expr_member_form_builtin.inc.
    if (knownFormModuleNames_.count(name)
        && (!isFormModule_ || name != knownFormName_)) {
        outType = FrmControlType::Form;
        return true;
    }
    return false;
}
bool CCodeGen::controlClickFromNativeNotify(FrmControlType ctrlType) {
    switch (ctrlType) {
    case FrmControlType::CommandButton:
    case FrmControlType::CheckBox:
    case FrmControlType::OptionButton:
    case FrmControlType::ListBox:
    case FrmControlType::ComboBox:
    case FrmControlType::DriveListBox:
    case FrmControlType::DirListBox:
    case FrmControlType::FileListBox:
    case FrmControlType::SSTab:
    case FrmControlType::Toolbar:
    case FrmControlType::Menu:
        return true;
    default:
        return false;
    }
}

void CCodeGen::emitShapeLineProps(const FrmControl& ctrl, const std::string& hwndExpr) {
    auto emitInt = [&](const char* prop, const char* fn, int skipWhen) {
        auto it = ctrl.properties.find(prop);
        if (it == ctrl.properties.end()) return;
        int v = (int)it->second.intValue;
        if (v == skipWhen) return;
        c_.emitLine(std::string(fn) + "((void*)" + hwndExpr + ", " + std::to_string(v) + ");");
    };
    if (ctrl.controlType == FrmControlType::Shape) {
        emitInt("Shape", "vb6_SetShapeType", -1);
        emitInt("BorderWidth", "vb6_SetShapeBorderWidth", 1);
        emitInt("BorderStyle", "vb6_SetShapeBorderStyle", 1);
        emitInt("FillStyle", "vb6_SetShapeFillStyle", 1);
        emitInt("FillColor", "vb6_SetShapeFillColor", 0);
        emitInt("BorderColor", "vb6_SetShapeBorderColor", 0);
    } else if (ctrl.controlType == FrmControlType::Line) {
        emitInt("X1", "vb6_SetLineX1", -1);
        emitInt("Y1", "vb6_SetLineY1", -1);
        emitInt("X2", "vb6_SetLineX2", -1);
        emitInt("Y2", "vb6_SetLineY2", -1);
        emitInt("BorderWidth", "vb6_SetLineBorderWidth", 1);
        emitInt("BorderStyle", "vb6_SetLineBorderStyle", 1);
        emitInt("BorderColor", "vb6_SetLineColor", 0);
    }
}

// P20-36: 生成控件属性访问的HWND参数 (Menu控件用GetMenu+menuId)
std::string CCodeGen::makeCtrlHwndArg(const std::string& ctrlNameLower, FrmControlType ctrlType) const {
    // Fix 161f-extlist: 形参/局部持有的 ListView (`Sub RefillList(lv As ListView)`) ——
    // 槽变量**本身就是 HWND** (调用点传的就是 vb6_hwnd_ListView1), 不能再拼
    // vb6_hwnd_ 前缀 (那样会 C2065 `vb6_hwnd_lv` 未声明)。与设计期控件名区分开。
    if (ctrlType == FrmControlType::ListView && listViewSlotVars_.count(ctrlNameLower)) {
        return ctrlNameLower;
    }
    // P20-39: ImageList 走原生复刻, 槽里是复刻实例指针不是 HWND, 必须用 vb6_com_<Name>
    // (emitControlHandleDecls 也是这么声明的) —— 发 vb6_hwnd_<Name> 就是 C2065。
    if (ctrlType == FrmControlType::ImageList) {
        auto it0 = knownFormControlOriginalNames_.find(ctrlNameLower);
        std::string n0 = (it0 != knownFormControlOriginalNames_.end()) ? it0->second : ctrlNameLower;
        return "vb6_com_" + cIdent(n0);
    }
    if (ctrlType == FrmControlType::Menu) {
        auto menuIt = knownMenuIds_.find(ctrlNameLower);
        if (menuIt != knownMenuIds_.end()) {
            return "(void*)GetMenu((HWND)" + knownMenuFormHwnd_ + "), " + std::to_string(menuIt->second);
        }
    }
    auto origIt = knownFormControlOriginalNames_.find(ctrlNameLower);
    std::string origName = (origIt != knownFormControlOriginalNames_.end()) ? origIt->second : ctrlNameLower;
    return "vb6_hwnd_" + cIdent(origName);
}

// Fix <vbeclipse>: 控件数组属性写的 callee 实参串。
// 普通控件: vb6_CtrlArr_GetAt(&vb6_arr_X, idx) (单参);
// 菜单数组: vb6_SetMenu* 是 (HMENU, menuId, 值) 三参 —— 菜单元素没有 HWND,
// 必须发 (GetMenu((HWND)窗体), 基址+下标), 否则 C2198 "用于调用的参数太少"。
std::string CCodeGen::ctrlArrWriteCalleeArgs(const std::string& writeFn,
                                             const std::string& arrOrigName,
                                             const std::string& idxArg) const {
    if (writeFn.rfind("vb6_SetMenu", 0) == 0) {
        std::string arrLower = arrOrigName;
        std::transform(arrLower.begin(), arrLower.end(), arrLower.begin(), ::tolower);
        auto itBase = knownMenuArrayBaseIds_.find(arrLower);
        std::string idExpr = (itBase != knownMenuArrayBaseIds_.end())
            ? std::to_string(itBase->second) + " + (" + idxArg + ")"
            : idxArg;
        return "(void*)GetMenu((HWND)" + knownMenuFormHwnd_ + "), " + idExpr;
    }
    return "vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(arrOrigName) + ", " + idxArg + ")";
}

// Fix <vbeclipse>: 在其他窗体的设计器描述里找控件 (frmViewViews.tvwViews)。
const FrmControl* CCodeGen::findExternalFormControl(const std::string& formName,
                                                    const std::string& ctrlName) const {
    if (!designerFiles_ || formName.empty() || ctrlName.empty()) return nullptr;
    for (const auto& kv : *designerFiles_) {
        if (Symbol::toLower(kv.first) != Symbol::toLower(formName)) continue;
        const FrmControl* hit = nullptr;
        std::function<void(const FrmControl&)> walk = [&](const FrmControl& c) {
            if (hit) return;
            if (!c.controlName.empty()
                && Symbol::toLower(c.controlName) == Symbol::toLower(ctrlName)) {
                hit = &c;
                return;
            }
            for (const auto& ch : c.children) walk(ch);
        };
        walk(kv.second.form.formControl);
        return hit;
    }
    return nullptr;
}

const char* CCodeGen::getDefaultPropertyName(FrmControlType ctrlType) {
    switch (ctrlType) {
    case FrmControlType::TextBox:      return "Text";
    // C30-3: MaskEdBox 的默认成员也是 Text（`MaskEdBox1 = "138..."` 落到掩码上）
    case FrmControlType::MaskEdBox:    return "Text";
    case FrmControlType::Label:        return "Caption";
    case FrmControlType::CommandButton: return "Caption";
    case FrmControlType::CheckBox:     return "Value";
    case FrmControlType::OptionButton: return "Value";
    case FrmControlType::ListBox:      return "Text";
    case FrmControlType::ComboBox:     return "Text";
    case FrmControlType::Frame:        return "Caption";
    case FrmControlType::Form:         return "Caption";
    case FrmControlType::MDIForm:      return "Caption";
    case FrmControlType::PictureBox:   return "Picture";  // P17.2
    case FrmControlType::Image:        return "Picture";  // P17.2
    case FrmControlType::HScrollBar:   return "Value";  // P20-41
    case FrmControlType::VScrollBar:   return "Value";  // P20-41
    default:                           return nullptr;
    }
}


// ============================================================
// 账 #256: 默认属性的**型**（与上面那张名字表一行一行对着写）。
// 为什么另开一张而不问 controlPropType：那张表答的是「显式写出来的 `<对象>.<属性>` 读法」
// 的型，而 Text / Caption / Value 挂在默认属性上时从没在里面登记过（实测 TextBox 的 text、
// Label 的 caption 在那里都答 Unknown）。裸控件名被折成默认属性时，**打包档必须跟着折出来的
// 那个值**，所以这一问的答案就是各行 RTL getter 的 C 返回型（已逐条回头看过）：
//   vb6_GetControlText  -> wchar_t*  (String)
//   vb6_GetCheckValue   -> int       (VB6 侧 CheckBox.Value 是 0/1/2 的 Integer)
//   vb6_GetOptionValue  -> int       (账 #128-b：答的就是 VB6 的 -1/0 ⇒ Boolean)
//   vb6_GetScrollValue  -> int       (VB6 侧 ScrollBar.Value 是 Integer)
//   vb6_GetControlPicture -> void*   (对象：这一格刻意维持今天的打包路，见 comPackExpr 的白名单)
// 只有 ctrlDefaultPropOf 问它（发码那一路不读 valueType，打包那一路只读 valueType）。
// ============================================================

Vb6Type CCodeGen::defaultPropType(FrmControlType ctrlType) {
    switch (ctrlType) {
    case FrmControlType::TextBox:
    case FrmControlType::ListBox:
    case FrmControlType::ComboBox:
    case FrmControlType::Label:
    case FrmControlType::CommandButton:
    case FrmControlType::Frame:
    case FrmControlType::Form:
    case FrmControlType::MDIForm:      return Vb6Type::String;
    case FrmControlType::CheckBox:     return Vb6Type::Integer;
    case FrmControlType::OptionButton: return Vb6Type::Boolean;
    case FrmControlType::HScrollBar:
    case FrmControlType::VScrollBar:   return Vb6Type::Integer;
    default:                           return Vb6Type::Unknown;
    }
}


// ============================================================
// 账 #258 (§B87): 与 ctrlDefaultPropOf 对称的那一问 —— 「这枚裸控件名作为**对象**交出去时发什么」。
// 只许有一条答案：设计期控件 → makeCtrlHwndArg (同一条句柄出口；ListView 的槽变量、ImageList 的
// vb6_com_<名>、Menu 的 GetMenu(...) 各有各的拼法，手拼 vb6_hwnd_ 会发出 C 里根本不存在的名字)，
// WithEvents 控件变量 → 那枚变量自己 (C 侧就是持有 HWND 的变量)，都不认识才退回 vb6_hwnd_<名>。
// 消费点 = cgen_expr_ident_symbol.inc 的 suppressDefaultProp_ 那一支（With 块 / 对象形参 / Set 右值三档共用）。
// ============================================================

std::string CCodeGen::ctrlObjectRefExpr(const std::string& lower,
                                        const std::string& cNameFallback) const {
    auto itWE = knownWithEventsCtrlVars_.find(lower);
    if (itWE != knownWithEventsCtrlVars_.end()) {
        auto itOrig = knownWithEventsCtrlOrigNames_.find(lower);
        return (itOrig != knownWithEventsCtrlOrigNames_.end()) ? itOrig->second : cNameFallback;
    }
    auto itFC = knownFormControls_.find(lower);
    if (itFC != knownFormControls_.end()) return makeCtrlHwndArg(lower, itFC->second);
    return "vb6_hwnd_" + cNameFallback;
}

// ============================================================
// 账 §42: 工程内 UserControl 的**默认成员**（只读设计器代码段，不碰解析器）
//
// VB6 把默认成员写在 `.ctl` 的代码段里：
//     Public Property Get Value() As Double
//     Attribute Value.VB_UserMemId = 0
//         Value = m_Value
//     End Property
// 判据只有这两行：`VB_UserMemId = 0` 给出成员名，同名 `Property Get … As <型>` 给出返回型。
// 只认 `= 0`（VB 的 DISPID_VALUE）；`= -4` 之类是别的东西，一律不折。
// ------------------------------------------------------------
static std::string trimAscii(const std::string& s) {
    size_t b = 0, e = s.size();
    while (b < e && std::isspace((unsigned char)s[b])) ++b;
    while (e > b && std::isspace((unsigned char)s[e - 1])) --e;
    return s.substr(b, e - b);
}

static bool findUcDefaultMember(const std::string& code, std::string& propOut,
                                std::string& typeOut) {
    if (code.empty()) return false;
    std::string low = code;
    for (char& c : low) c = (char)std::tolower((unsigned char)c);

    std::string prop;
    size_t pos = 0;
    while (pos < low.size()) {
        size_t eol = low.find('\n', pos);
        if (eol == std::string::npos) eol = low.size();
        size_t start = pos;
        while (start < eol && std::isspace((unsigned char)low[start])) ++start;
        if (eol - start > 10 && low.compare(start, 10, "attribute ") == 0) {
            size_t mem = low.find(".vb_usermemid", start);
            if (mem != std::string::npos && mem < eol) {
                size_t eq = low.find('=', mem);
                if (eq != std::string::npos && eq < eol) {
                    std::string val;
                    for (size_t i = eq + 1; i < eol; ++i) {
                        const char c = low[i];
                        if (std::isdigit((unsigned char)c) || c == '-' || c == '+') val += c;
                        else if (!val.empty()) break;
                    }
                    if (val == "0") {
                        size_t nameStart = start + 10;
                        while (nameStart < mem && std::isspace((unsigned char)low[nameStart])) ++nameStart;
                        std::string name = trimAscii(code.substr(nameStart, mem - nameStart));
                        if (!name.empty()) { prop = name; break; }
                    }
                }
            }
        }
        pos = eol + 1;
    }
    if (prop.empty()) return false;
    std::string lowNeedle = "property get " + prop;
    for (char& c : lowNeedle) c = (char)std::tolower((unsigned char)c);
    // 只认**非注释行**上的声明。旧写法把整个 codeSection（含注释）都当候选，`.ctl` 注释里
    // 顺口提到 `Property Get <Prop>() As <型>` 就会把注释尾巴当型 ⇒ ucDefaultReadPlan
    // 落空 ⇒ 静默不折（C2440 原样回来）。声明行在 VB6 里写在
    // `Attribute <Prop>.VB_UserMemId = 0` **之前**，所以起点仍是 0，只加注释行过滤。
    size_t g = 0;
    while ((g = low.find(lowNeedle, g)) != std::string::npos) {
        size_t after = g + lowNeedle.size();
        if (after < low.size() && (std::isalnum((unsigned char)low[after]) || low[after] == '_')) {
            g = after;
            continue;
        }
        size_t ls = low.rfind('\n', g);
        ls = (ls == std::string::npos) ? 0 : ls + 1;
        while (ls < g && std::isspace((unsigned char)low[ls])) ++ls;
        if (ls < low.size() && low[ls] == '\'') { g = after; continue; }   // 注释行不认
        size_t eol = low.find('\n', after);
        if (eol == std::string::npos) eol = low.size();
        const std::string decl = low.substr(after, eol - after);
        size_t asPos = decl.rfind(" as ");
        if (asPos != std::string::npos) {
            std::string ty = trimAscii(decl.substr(asPos + 4));
            size_t cut = ty.find_first_of(" \t'(");
            if (cut != std::string::npos) ty = ty.substr(0, cut);
            if (!ty.empty()) {
                propOut = prop;
                typeOut = ty;
                return true;
            }
        }
        g = after;
    }
    return false;
}

// 默认成员的返回型 → COM 读数（只登记 PhotoDemon 与常见语料里出现过的型；其余不折、保持原样）
static bool ucDefaultReadPlan(const std::string& typeLower, std::string& readFn,
                              Vb6Type& vt, std::string& castTo) {
    if (typeLower == "double")  { readFn = "vb6_ComGetDoubleProp"; vt = Vb6Type::Double;  return true; }
    if (typeLower == "single")  { readFn = "vb6_ComGetDoubleProp"; vt = Vb6Type::Single;  castTo = "(float)"; return true; }
    if (typeLower == "long")    { readFn = "vb6_ComGetIntProp";    vt = Vb6Type::Long;    return true; }
    if (typeLower == "integer") { readFn = "vb6_ComGetIntProp";    vt = Vb6Type::Integer; return true; }
    if (typeLower == "boolean") { readFn = "vb6_ComGetIntProp";    vt = Vb6Type::Boolean; return true; }
    if (typeLower == "string")  { readFn = "vb6_ComGetStringProp"; vt = Vb6Type::String;  return true; }
    if (typeLower == "date")    { readFn = "vb6_ComGetDoubleProp"; vt = Vb6Type::Date;    return true; }
    return false;
}

// ============================================================
// 账 #256: 「这枚裸控件名会被折成哪一枚默认属性读数、折出来是什么型」只此一处回答
//         （声明与背景见 cgen_helpers.inc 的 CtrlDefaultProp）。
// 两个消费点：① 标识符发码那条路（cgen_expr_ident_symbol.inc）拿 readFn/prop/ctrlType
// 去发 `readFn(句柄)`；② 晚绑定实参的打包档（comPackExpr）只问 valueType。
// 两边问同一个 suppressDefaultProp_，所以「发的是值」与「按什么型装箱」不可能各说一遍。
// ============================================================

CCodeGen::CtrlDefaultProp CCodeGen::ctrlDefaultPropOf(const std::string& lower) const {
    CtrlDefaultProp r;
    if (suppressDefaultProp_) return r;   // With 块 / 对象形参那一档：交出的是控件本身（HWND）

    auto it = knownFormControls_.find(lower);
    bool isWE = false;
    if (it == knownFormControls_.end()) {
        it = knownWithEventsCtrlVars_.find(lower);   // P20-31: WithEvents 控件变量同一档语义
        if (it == knownWithEventsCtrlVars_.end()) return r;
        isWE = true;
    }
    const char* prop = getDefaultPropertyName(it->second);
    if (!prop) {
        // 账 §42: 工程内 UserControl 的默认成员。`getDefaultPropertyName` 只认 13 个内置
        // FrmControlType，工程内 UC 一律 Unknown ⇒ 这里回落到 `.ctl` 的 VB_UserMemId = 0。
        if (isWE) return r;                       // WithEvents 变量不会指向工程内 UC
        auto itTN = knownFormControlTypeNames_.find(lower);
        if (itTN == knownFormControlTypeNames_.end()) return r;
        const FrmFile* ucSpec = findUserControlSpec(itTN->second);
        if (!ucSpec) return r;
        std::string ucProp, ucType;
        if (!findUcDefaultMember(ucSpec->codeSection, ucProp, ucType)) return r;
        std::string readFn;
        Vb6Type vt = Vb6Type::Unknown;
        std::string castTo;
        if (!ucDefaultReadPlan(ucType, readFn, vt, castTo)) return r;
        r.folds = true;
        r.ctrlType = it->second;
        r.isUcDefault = true;
        r.propName = std::move(ucProp);
        r.readFn = std::move(readFn);
        r.castTo = std::move(castTo);
        r.readArgTail = ", L\"" + r.propName + "\"";   // COM getter 的第二档就是属性名
        r.valueType = vt;
        return r;
    }
    std::string readFn = getControlPropReadFn(it->second, prop);
    if (readFn.empty()) return r;        // 没有读函数 ⇒ 发码那条路本来就不会折
    r.folds = true;
    r.isWithEvents = isWE;
    r.ctrlType = it->second;
    r.prop = prop;
    r.readFn = std::move(readFn);
    r.valueType = defaultPropType(it->second);
    return r;
}

// 账 §42 续（F1）: 默认成员的读数/写值两条发码腿收成这一对，消费点只有两档接收者：
// 裸名（句柄）与控件数组元素（vb6_CtrlArr_GetAt(...)）。判据仍只住 ctrlDefaultPropOf，
// 站点不许再各拼一份 `castTo + readFn` / `ComSetProp + L"<prop>"`。
std::string CCodeGen::ctrlDefaultReadExpr(const std::string& lower,
                                          const std::string& recvExpr) const {
    CtrlDefaultProp fold = ctrlDefaultPropOf(lower);
    if (!fold.folds || fold.isWithEvents) return std::string();
    return fold.castTo + fold.readFn + "(" + recvExpr + fold.readArgTail
         + ")  /* default prop: ." + fold.propText() + " */";
}

bool CCodeGen::emitUcDefaultPropLet(const CtrlDefaultProp& fold,
                                    const std::string& recvExpr, Expr& valueExpr) {
    if (!fold.folds || !fold.isUcDefault) return false;
    emitExpr(valueExpr);
    std::string valUc = std::move(lastExpr_);
    // 装箱档问**唯一**那一问 (cgen_util_com.cpp 的 comPackFnForDefaultPropType) ——
    // 它与「裸控件名折默认属性」那一支 (comPackExpr 里的同一问) 共用同一张配对表与同一道
    // 六档闸 ⇒ 同一枚型不可能在两站给出两个打包器名 (第二百四十一刀 V11 记的真回归)。
    std::string packUc = comPackFnForDefaultPropType(fold.valueType);
    if (packUc.empty()) packUc = comPackExpr(valueExpr);
    c_.emitLine("vb6_ComSetProp(" + recvExpr + ", L\"" + fold.propText() + "\", " + packUc
                + "(" + valUc + "));  /* UC default prop let */");
    return true;
}




// ============================================================
// P8.4: Variant值包装 - 根据表达式类型推断Variant构造函数
// ============================================================

// 账 #305 §B164 (第五十五刀): 「把一枚值收进 Variant 槽」收成这一枚 —— 清旧值 + 交接新值
// 是一个决定的两半句，站点不许再各自抄。两半都要: 只抄第二半(裸赋值)每遍历一个元素漏
// 81/71 字节；只补第一半(Clear)而右值其实是借来的(数组元素那一路)就是把别人还活着的
// BSTR free 掉。两条都在本轮实测里现过形(台账 §B164 形④)。
void CCodeGen::emitVariantSlotStore(const std::string& target, const std::string& expr,
                                    bool rhsOwnsValue, bool viaTemp) {
    if (!rhsOwnsValue) {
        c_.emitLine("vb6_VariantAssign(&" + target + ", " + expr + ");");
        return;
    }
    // ucTreeView (2026-11-06): `With 链`里写字段的目标**带副作用** —— 典型
    //   With ucTreeViewPlus1.Nodes.Add("Facebook") : .Add("Ocultar").Image = 1
    // 目标发成 `vb6_Node_Add(...)->Image`。下面两半句 (Clear / 赋值) 各写一次目标文本 ⇒
    // 每写一枚字段就多调一次 Add (实测 C3 树比 VB6 多一行: "Ocultar" 出现两次)。
    // 判据: 目标文本在**最后一级成员之前**出现过 '(' (函数调用/宏)。命中就把**地址**
    // 提成具名临时量, 两半句都读它 ⇒ 目标恰好求值一次。纯左值形态 emit 逐字节不变。
    auto needsAddrTemp = [](const std::string& t) -> bool {
        size_t mem = t.rfind("->");
        if (mem == std::string::npos) return false;
        size_t lp = t.find('(');
        return lp != std::string::npos && lp < mem;
    };
    const bool hoistTarget = needsAddrTemp(target);
    std::string tgt = target;
    // 非提存路**逐字节不变**: 就是 `vb6_VariantClear(&<target>)` —— 多处发码形状针
    // (b164_slot_store_shape / b169_class_field_store_shape) 按原文逐字核这一行。
    std::string clearArg = "&" + target;
    if (hoistTarget) {
        const std::string tp = "_vtg_" + std::to_string(tempCounter_++);
        c_.emitLine("{ vb6_VARIANT* " + tp + " = &(" + target + ");");
        tgt = "*" + tp;
        clearArg = tp;
    }
    if (viaTemp) {
        const std::string tmp = "_vc_" + std::to_string(tempCounter_++);
        c_.emitLine("{ vb6_VARIANT " + tmp + " = " + expr + ";");
        c_.emitLine("  vb6_VariantClear(" + clearArg + ");");
        c_.emitLine("  " + tgt + " = " + tmp + "; }");
        if (hoistTarget) c_.emitLine("}");
        return;
    }
    c_.emitLine("vb6_VariantClear(" + clearArg + ");");
    c_.emitLine(tgt + " = " + expr + ";");
    if (hoistTarget) c_.emitLine("}");
}

// 账 #305 §B169 (第六十刀): 「把这枚右值收进一处 BSTR 槽」收成这一枚出口 —— 原先这段
// 住在 visit(AssignmentStmt&) 的 targetIsBstr 那一支里。搬出来的理由不是整洁: With 块里
// 写类的 String 数据字段那一站 (§B169 形②) 需要的正是同一串决定 (右值折算 / 自有还是借来 /
// 定长串收口 / 清旧值+拷新值), 在那里再抄一遍就是第四份答案 (与 §B164 第五十五刀把槽位写入
// 收成 emitVariantSlotStore 同一手法)。行为逐字节不变 —— 语料 emit A/B 钉这一条。
void CCodeGen::emitBstrSlotStore(const std::string& target, Expr* rhs,
                                 std::string value, const std::string& fixedLen,
                                 bool udtFixedField) {
    // Fix 038b-6: BSTR 目标 + Variant 值 → 先提取 BSTR
    // 使用 cExprIsVariant (C 字符串级) + knownVariantVars_ 检测.
    // 账 §56.40: 同 AssignmentStmt 那一处 (cgen_assign_value_sem.inc) 的改法 —— 问上
    // AST 那一半 (rhs 是现成的形参, 不是「拿不到」)。不修这一格, 「UDT 字段声明为
    // Variant」这一族 (C 文本 VB6_SA_AT(vb6_type_X, arr, i).member 不命中任何前缀表)
    // 仍然不拆箱 ⇒ 整只 vb6_VARIANT 交给 vb6_BSTR_Assign ⇒ C2440
    // (PhotoDemon pdDictionary.c 239/248: `dItem As Variant` 读进 String 变量)。
    bool valueIsVariant = isVariantSlot(value, rhs);
    bool strictVarAssign221 = false;   // 第七十三刀: 右值是 Variant 槽 ⇒ 走严格赋值出口
    std::string rawVarValue221;        // 未折成 BSTR 的那枚 Variant 表达式
    if (!valueIsVariant && rhs && rhs->kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<IdentifierExpr&>(*rhs);
        std::string idLower = id.name;
        std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
        if (knownVariantVars_.count(idLower)) valueIsVariant = true;
    }
    // Fix 089j: RHS 为类 Variant 字段 (me->mParentsColKey 等) 也视为
    // Variant 表达式 — cExprIsVariant 只查函数前缀, me-> 成员不命中,
    // 导致 PropertyGet 里 `ret = me->VarField` 不转 → C2440.
    if (!valueIsVariant && value.compare(0, 4, "me->") == 0) {
        std::string rhsMember = value.substr(4);
        std::transform(rhsMember.begin(), rhsMember.end(), rhsMember.begin(), ::tolower);
        if (classVariantMembers_.count(rhsMember)) valueIsVariant = true;
    }
    // C30-11 第七十四刀 (§B222): `s = CStr(<Variant>)` —— CStr 内建对变体走的是**宽松口**
    // vb6_CStr (Null ⇒ "")。读数 (探针 E9): 真 VB6 是 94 且**目标不变**; 这里 err=0 且目标
    // 被清成空串 —— 与第七十三刀 E6/E7/E8 是同一个错, 只是入口在 CStr 那一道站, 不在槽位折叠。
    // 形状恰好是 vb6_CStr(<变体表达式>) ⇒ 把「CStr 口径」连同那枚变体一起交给同一枚严格出口:
    // 出口内部就是 AssignMove(target, vb6_CStr(src)) ⇒ 非 Null 逐字节等价, Null 报 94 后
    // 直接返回、目标不动。只认**单实参 + 实参是变体档 + 发码确实是 vb6_CStr(**那一形;
    // 第六十六刀那两条 vb6_CStrStrict(<比较>/<拼接>) 不在此列 (前缀就不是 vb6_CStr();
    // 定长串那一档读数未量, 仍走原路。
    if (!valueIsVariant && rhs
        && rhs->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& call74 = static_cast<IndexOrCallExpr&>(*rhs);
        if (call74.callee && call74.callee->kind == ASTNodeKind::IdentifierExpr
            && call74.positional.size() == 1 && call74.named.empty()
            && value.size() > 10 && value.compare(0, 9, "vb6_CStr(") == 0
            && value.back() == ')') {
            std::string name74 = static_cast<IdentifierExpr&>(*call74.callee).name;
            std::transform(name74.begin(), name74.end(), name74.begin(), ::tolower);
            if (name74 == "cstr"
                && inferExprType(*call74.positional[0]) == Vb6Type::Variant) {
                strictVarAssign221 = true;
                rawVarValue221 = value.substr(9, value.size() - 10);
            }
        }
    }
    if (valueIsVariant) {
        // C30-11 第七十三刀 (§B221): Variant 右值 → String 槽走**严格**出口。
        // 读数 (探针 E6/E7/E8): `s = Trim(Null)` / `s = Len(Null)` / `s = Null` 在真 VB6
        // 里都是 94 且目标不变；而这里以前折成 vb6_VariantToString(Null) = "" ⇒ err 0、
        // 目标被清成空串 —— 两个错同时存在 (值丢 + 错误不报)。
        strictVarAssign221 = true;
        rawVarValue221 = value;
        value = "vb6_VariantToString(" + value + ")";
    }
    // Fix P24-10b: RHS 是裸晚绑定调用结果 (vb6_ComCallByDispid → void* 指向
    // VARIANT) 且目标是 BSTR — Fix 113 已把 ComCall 无类型结果从前缀表回退,
    // valueIsVariant 不命中, 原样交给 vb6_BSTR_Assign 会把它当 BSTR 指针
    // → 运行期 0xC0000005 (test_com_latebound `s1 = o("name")` 实测)。
    // 先 VariantFromComResult 包回 vb6_VARIANT 再提取 BSTR, 与下方 else
    // 分支的 Fix 124 valueIsBareComResult 同口径 (只认返回 void* VARIANT*
    // 的四兄弟; ComCallBSTR/Int/Double/Object 返回类型化值, 不在此列)。
    if (!valueIsVariant) {
        bool comResBstr = value.compare(0, 12, "vb6_ComCall(") == 0
            || value.compare(0, 20, "vb6_ComCallByDispid(") == 0
            || value.compare(0, 15, "vb6_ComGetProp(") == 0;
        // C30-11 第七十七刀 (§B225): 这里原来是**先转串再赋**:
        //   value = "vb6_VariantToString(vb6_VariantFromComResult(<调用>))"
        // 真机读数 (W6 `s = c(1)`, Collection.Item 装的是 Null): **94 且 s 还是 [keep]**；
        // 而 VariantToString(Null) 不 raise、交回空串 ⇒ 这里变成 err=0 + 目标被清成 ""，
        // 两个错同时存在。Null 必须在**赋值那一瞬**才判 (与第七十三刀同一个道理): 把整只
        // 变体交给 vb6_BSTR_AssignFromVariant (< 非 Null 走 vb6_CStr, 与 VariantToString 逐字节同口径)。
        // 定长串目标 (char[N]) 那一档不能套这枚出口, 仍走原来的转串路径。
        std::string comVarVal77;
        if (comResBstr) {
            comVarVal77 = "vb6_VariantFromComResult(" + value + ")";
        } else if (value.compare(0, 21, "vb6_ComGetObjectProp(") == 0) {
            // 账 (M3 身份刀): 对象读交回 IDispatch* 本体, 不能当堆上 VARIANT* 解引用。
            // 换到同一成员的值孪生 (vb6_ComGetProp) 再按字符串提取 —— 成员名一字不动。
            std::string twinObjStr = comValueTwinForObjectReadExpr(value);
            if (!twinObjStr.empty())
                comVarVal77 = "vb6_VariantFromComResult(" + twinObjStr + ")";
        }
        if (!comVarVal77.empty()) {
            if (fixedLen.empty()) {
                strictVarAssign221 = true;
                rawVarValue221 = comVarVal77;
            } else {
                value = "vb6_VariantToString(" + comVarVal77 + ")";
            }
        }
    }
    // Fix 140: RHS 是 Byte() 数组 (类成员 m_Caption 或已知字节数组变量) 且目标是
    // BSTR → 按 UTF-16LE 字节还原为字符串. 不能把数组指针当 BSTR 交给
    // vb6_BSTR_Assign (SysStringLen 会读数组前的堆元数据当长度 → 堆损坏).
    // 场景: LabelPlus.ctl `Property Get Caption() As String: Caption = m_Caption`.
    // 记账这一支改没改成 (给下面 #122 那支看): 值本身是 Byte() 数组 ⇒ inferExprType 会
    // 报 Byte, 但那已经是**转好的字符串**了, 再套一层 CStrByte 是错的 (实测把
    // LabelPlus 的 `Caption = m_Caption` 发成 vb6_CStrByte(vb6_ByteArrayToString(..)))。
    bool wrappedByteArr140 = false;
    if (!valueIsVariant) {
        std::string rhs140 = value;
        if (rhs140.compare(0, 4, "me->") == 0) rhs140 = rhs140.substr(4);
        if (rhs140.size() > 4 && rhs140[0] == '(' && rhs140[1] == '*'
            && rhs140.back() == ')') {
            rhs140 = rhs140.substr(2, rhs140.size() - 3);
        }
        std::string rhs140l = rhs140;
        std::transform(rhs140l.begin(), rhs140l.end(), rhs140l.begin(), ::tolower);
        if (rhs140.find_first_of("( .") == std::string::npos
            && (knownByteArrayVars_.count(rhs140l)
                || classByteArrayMembers_.count(rhs140l))) {
            value = "vb6_ByteArrayToString(" + value + ")";
            wrappedByteArr140 = true;
        }
    }
    // 账 #122: VB6 在**赋值那一刻**就做隐式 CStr —— `Dim s As String: s = 123` 合法，
    // 而这一格以前把值原样交给 vb6_BSTR_Assign。后果分两种，别当成同一种：
    //   整型那半 (Long/Integer/Boolean/Byte) 在 C 里只是警告 (int → BSTR 形参)，数字被
    //     当指针存进 String 变量，**第一次读/释放就崩** —— 实测修复前二进制 `s = n` 是
    //     BUILD-RC=0 而跑起来 SIGSEGV (rc=139)，所以它不是"编不过"，是静默埋雷;
    //   浮点/日期那半 (Double/Currency/Date) 才是响的 cl `error C2440 无法从 double
    //     转换为 BSTR`。
    // 这一支只在 inferExprType 认得出的**标量**类型上折 (isScalarImplicitStr159:
    // Integer/Long/Single/Double/Byte/Boolean/Currency/Date)，String/Variant/Object/
    // 未知一律不介入。正因为整型那半今天过得去，本批的护栏不是"现存工程零改动"而是
    // "改动必须逐条解释得清": 实测 7 枚存量工程里 6 枚逐字节相同，Charts 2020 只剩
    // `sDiplay = CLng(me->m_Value)` 那一行 —— 那是一处真实修复而非回归。
    // 复用形参那条路 (Fix 159-B) 的同一个 wrapToBSTR, 不另开一套折算表。
    // 上面两支 (Variant 提取 / Byte() 还原) 已经把值转成了字符串, 不能再套一层。
    bool wrapped122 = false;
    // C30-9: 对象→String 也在这一档收口。`s = rs.Fields(2)` 的右值是
    // vb6_ComCallObject(...) (IDispatch*), inferExprType 给的是对象/未知档 ⇒
    // isScalarImplicitStr159 为假, 不加这一条就发成 vb6_BSTR_Assign(&s, <对象指针>),
    // BSTR 拷贝例程拿对象指针当串读 ⇒ 0xC0000005 (memcpy 野读, 实测两处 av target)。
    // wrapToBSTR 里有同一口径的 vb6_ComCallObject 档 (取默认成员再转串)。
    const bool objRhs122 = (value.compare(0, 18, "vb6_ComCallObject(") == 0);
    if (!valueIsVariant && !wrappedByteArr140 && rhs) {
        Vb6Type vt122 = inferExprType(*rhs);
        if (isScalarImplicitStr159(vt122) || objRhs122) {
            std::string w122 = wrapToBSTR(value, *rhs);
            if (w122 != value) {
                value = w122;
                wrapped122 = true;
            }
        }
    }
    // 账 #113 / #119: 值是不是"自有的临时串" —— **字符串**字面量 / 拼接 / 返回 String 的
    // 工程函数或 Property Get (返回变量 vb6_ret_* 在入口处已按值语义拷过)。变量、形参、
    // 成员字段、数组元素都是**借来的**, 只能深拷贝。
    // #119 之前这条判据只服务模块级目标 (#113 那格为了少翻发码面刻意加了
    // `bstrTargetViaTypeQuery` 那道闸), 局部/形参/UDT String 字段三条老路一律 Assign
    // ⇒ 每句从自有临时来的赋值都漏一只堆块 (`s = Left(x,2)` 就是 64 字符 + 块头那种
    // 146 字节级的漏, 循环里就是纯漏, 旧值要等下一次赋值才被释放)。
    // 铺开的底气是这一轮把 RTL 全扫了一遍: 77 枚 `BSTR` 返回函数里只有 3 枚发的是**借来
    // 的**指针 (vb6_ErrDescription / vb6_ErrSource 直接回 vb6_err 那两个字段、
    // vb6_ho_variantToBstr 回 v->bstrVal), 而它们都只能以成员访问的形态到达
    // (`Err.Description`), 下面的 IdentifierExpr 那一支根本捡不到 ⇒ 不会被误判自有。
    // 字面量那支限定成 LiteralKind::String: 数值/布尔字面量赋给 String 目标这一格发的是
    // 裸数字 (`vb6_BSTR_Assign(&s, 123)` → cl C2440, 已另立一账), 把它标成自有只会把
    // "当场编译不响"变成"存进一只假指针", 不能顺手动。
    bool valueOwned113 = false;
    if (rhs) {
        switch (rhs->kind) {
            case ASTNodeKind::LiteralExpr:
                valueOwned113 = (static_cast<LiteralExpr&>(*rhs).literalKind
                                 == LiteralKind::String);
                break;
            case ASTNodeKind::BinaryExpr: {
                auto& vb113 = static_cast<BinaryExpr&>(*rhs);
                valueOwned113 = (vb113.op == BinaryOp::Concat || vb113.op == BinaryOp::Add);
                break;
            }
            case ASTNodeKind::IndexOrCallExpr: {
                auto& vc113 = static_cast<IndexOrCallExpr&>(*rhs);
                if (vc113.callee && vc113.callee->kind == ASTNodeKind::IdentifierExpr) {
                    auto* vSym113 = symTab_.lookup(
                        static_cast<IdentifierExpr&>(*vc113.callee).name);
                    valueOwned113 = (vSym113
                                     && (vSym113->kind == SymbolKind::Function
                                         || vSym113->kind == SymbolKind::PropertyGet));
                }
                break;
            }
            default:
                break;
        }
    }
    // #122 折出来的 `vb6_CStrXxx(...)` 是刚 alloc 的新块, 与字面量/拼接同属"自有"。
    if (strictVarAssign221 && fixedLen.empty()) {
        // 定长串那一档读数未量, 仍走原路 (只认普通 String 槽)。
        c_.emitLine("vb6_BSTR_AssignFromVariant(&" + target + ", " + rawVarValue221 + ");");
        return;
    }
    // C30-11 第九十刀: 定长串目标收到 Variant 右值 —— 两支各走一枚**原子**出口
    // (Null ⇒ 94 且目标一字不动; 其余档折成串后按 N 补空格/截右)。
    // 不套上面的 vb6_BSTR_AssignFromVariant 是因为定长串的写侧有自己的形状:
    // UDT 字段是内联 char[N] (不能当 BSTR* 用), 模块级是 BSTR 槽但多一道 fit。
    if (strictVarAssign221 && !fixedLen.empty()) {
        if (udtFixedField) {
            c_.emitLine("vb6_FixedStrAssignFromVariant(" + target + ", " + fixedLen + ", "
                        + rawVarValue221 + ");");
        } else {
            c_.emitLine("vb6_LSetFromVariant(&" + target + ", " + fixedLen + ", "
                        + rawVarValue221 + ");");
        }
        return;
    }
    if (!fixedLen.empty() && udtFixedField) {
        // 账 #118 续: 目标是 UDT 定长串字段 (C 类型 char[N], VB6 ABI 内联 ANSI) ——
        // 不能 &field 当 BSTR* 用。写侧按 N 右补空格/右截断 (与 LSet 同口径)。
        // 自有临时串那一侧顺手释放 (与上面 vb6_LSetFree 的选择同一把尺)。
        //
        // C30-11 第八十九刀: 语句级中止 —— 本语句里 raise 过就不写 char[N]。
        // 旧形只发了 `const int32_t _eN = vb6_ErrSerial();` 却不看它: `u.f = Left("abc",-1)`
        // 报 5 之后字段被抹成空格（真机：一字不动）。形与下面普通 BSTR 槽那两支同构:
        // 右值**只算一次**进临时量, 再拿算完之后的序号决定是写进去还是收回。
        const bool ownedU89 = valueOwned113 || wrapped122;
        const std::string assignFn118u =
            ownedU89 ? "vb6_FixedStrAssignFree" : "vb6_FixedStrAssign";
        if (stmtAbortGuard_.empty()) {
            c_.emitLine(assignFn118u + "(" + target + ", " + fixedLen + ", " + value + ");");
        } else {
            const std::string tmp89 = "_av" + std::to_string(stmtAbortSeq_);
            c_.emitLine("BSTR " + tmp89 + " = (" + stmtAbortGuard_ + " == vb6_ErrSerial()) ? ("
                        + value + ") : NULL;");
            if (ownedU89) {
                c_.emitLine("if (" + stmtAbortGuard_ + " == vb6_ErrSerial()) " + assignFn118u
                            + "(" + target + ", " + fixedLen + ", " + tmp89
                            + "); else vb6_BSTR_Free(" + tmp89 + ");");
            } else {
                c_.emitLine("if (" + stmtAbortGuard_ + " == vb6_ErrSerial()) " + assignFn118u
                            + "(" + target + ", " + fixedLen + ", " + tmp89 + ");");
            }
        }
    } else if (!fixedLen.empty()) {
        // 账 #118 ①: 定长串赋值的收口 —— 短则右侧补空格、长则截右 (VB6 手册 + 真机
        // CopyMemory 实测一致, 与 LSet 同口径)。必须先 fit 成一只**新** BSTR 再 move
        // 进目标: 先 Assign 再回头改长度会多留一只临时, 而裸赋值连长度都不对。
        // RHS 是自有临时串那一侧用 vb6_LSetFree 顺手把它释放 —— 尺与下面
        // AssignMove/Assign 的选择同一把, 否则每次赋值都漏一块 (账 #118 ②)。
        //
        // C30-11 第八十九刀: 同一道语句级中止门也套在这里 (`g = Left("abc",-1)` 报 5 之后
        // g 旧值不得被抹掉)。fit 出来的那只新串只有序号没变才并进目标, 否则收回。
        const bool ownedF89 = valueOwned113 || wrapped122;
        const std::string fitFn118 =
            ownedF89 ? "vb6_LSetFree" : "vb6_LSet";
        if (stmtAbortGuard_.empty()) {
            c_.emitLine("vb6_BSTR_AssignMove(&" + target + ", " + fitFn118 + "(" + value
                        + ", " + fixedLen + "));");
        } else {
            const std::string tmp89f = "_av" + std::to_string(stmtAbortSeq_);
            c_.emitLine("BSTR " + tmp89f + " = (" + stmtAbortGuard_ + " == vb6_ErrSerial()) ? ("
                        + fitFn118 + "(" + value + ", " + fixedLen + ")) : NULL;");
            c_.emitLine("if (" + stmtAbortGuard_ + " == vb6_ErrSerial()) vb6_BSTR_AssignMove(&"
                        + target + ", " + tmp89f + "); else vb6_BSTR_Free(" + tmp89f + ");");
        }
    } else if (valueOwned113 || wrapped122) {
        // C30-11 第七十九刀 (§B227): 语句级中止 —— 本语句里 raise 过就不写目标；
        // RHS 是**自有临时串**，不 move 进去就得自己释放，否则每次中止漏一块。
        if (stmtAbortGuard_.empty()) {
            c_.emitLine("vb6_BSTR_AssignMove(&" + target + ", " + value + ");");
        } else {
            // C30-11 第八十三刀：右值**只算一次**。原形把 value 在 if/else 两支各写一遍，
            // 守卫不过时 else 那一支又把右值算了一遍（`s = CStr(F(v))` 会调 F 两次）。
            // 现在先算进临时量（只在本语句还没 raise 时算），再拿序号决定是并进目标还是收回。
            const std::string tmp83 = "_av" + std::to_string(stmtAbortSeq_);
            c_.emitLine("BSTR " + tmp83 + " = (" + stmtAbortGuard_ + " == vb6_ErrSerial()) ? ("
                        + value + ") : NULL;");
            c_.emitLine("if (" + stmtAbortGuard_ + " == vb6_ErrSerial()) vb6_BSTR_AssignMove(&" + target
                        + ", " + tmp83 + "); else vb6_BSTR_Free(" + tmp83 + ");");
        }
    } else {
        // C30-11 第七十九刀 (§B227): 语句级中止 —— 借来的值只深拷贝, 不写就不写,
        // 没有要回收的东西。手抄这一发仍留在原处 (R7 名单钉的是「这一站还是手写的」)。
        //
        // C30-11 第八十三刀 (§B227 续): 借来的值也要**只算一次** —— 与上面自有支同一形。
        // 原形把 value 直接写在 if 里, 而条件是**求值之前**比的：右值自己 raise（`Asc("")`
        // 报 5、实参位的 `vb6_CStrStrict(<Null>)` 报 94）时条件早就过了 ⇒ 赋值照样发生,
        // 目标被清成 ""，而真机是「报错 + 目标不动」。现在先算进临时量, 再拿**算完之后**
        // 的序号决定写不写；借来的指针不归我们, 所以没有 else 要回收的东西。
        if (stmtAbortGuard_.empty()) {
            c_.emitLine("vb6_BSTR_Assign(&" + target + ", " + value + ");");
        } else {
            const std::string tmp83b = "_av" + std::to_string(stmtAbortSeq_);
            c_.emitLine("BSTR " + tmp83b + " = (" + stmtAbortGuard_ + " == vb6_ErrSerial()) ? ("
                        + value + ") : NULL;");
            c_.emitLine("if (" + stmtAbortGuard_ + " == vb6_ErrSerial()) vb6_BSTR_Assign(&" + target
                        + ", " + tmp83b + ");");
        }
    }
}

// 账 #305 §B164 形⑤ 站③ (第一百八十四刀): 「这一发赋值的目标是不是一枚 Variant 槽」——
// 原先这六条判据内联在 visit(AssignmentStmt&) 里, 而 LetStmt 那一站 (显式 `Let`)
// 从不问它们, 于是 `Let <变体> = "abc"` / `= <String 变量>` / `= <拼接>` 三种形
// 全发成裸赋值 ⇒ cl C2440「无法从 BSTR 转换为 vb6_VARIANT」(实测 BASE 3 发),
// 而 `Let <变体> = <变体槽>` / `Let <变体数组元素> = …` 两种形编得过但不清旧值。
// 两个站点问的是同一个问题, 答案只许住在这里; 命中后由调用方交给同一枚
// emitVariantSlotAssign (装箱 + 所有权 + 槽位语义), 站点不许自己拼形状。
bool CCodeGen::variantSlotTargetOf(const Expr* target, const std::string& targetText,
                                   bool& udtField) {
    udtField = false;
    std::string checkName = targetText;
    if (checkName.substr(0, 4) == "me->") checkName = checkName.substr(4);
    // Fix 025: (*name) 解引用形式 — ByRef Variant 参数, 剥掉外层才能查名字。
    if (checkName.size() > 4 && checkName[0] == '(' && checkName[1] == '*'
        && checkName.back() == ')') {
        checkName = checkName.substr(2, checkName.size() - 3);
    }
    std::string lower = checkName;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    if (knownVariantVars_.count(lower)) return true;
    // Fix 054: 类成员 Variant 字段 (me->m_P1)。
    if (classVariantMembers_.count(lower)) return true;
    if (target && target->kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<const IdentifierExpr&>(*target);
        std::string idLower = id.name;
        std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
        if (knownVariantVars_.count(idLower)) return true;
    }
    // Fix 154: VARIANT 数组元素 —— 目标发出来是 `VB6_SA_AT(vb6_VARIANT, …)` 族下标宏。
    if (targetText.compare(0, 7, "VB6_SA_") == 0
        && targetText.find("(vb6_VARIANT,") != std::string::npos) {
        return true;
    }
    if (!target) return false;
    // Fix 232: UDT 的 Variant 字段 (`.Image` / With 块内的 `.Image`) —— 判据是字段声明类型;
    // 字段的真 C 类型是 vb6_cls_* 时否决 (typed 对象指针槽, 不是 Variant 容器)。
    const std::string udtFldC232 = udtFieldCTypeOfTarget(target);
    if (udtFldC232.rfind("vb6_cls_", 0) != 0
        && inferUdtFieldVb6Type(target, nullptr) == Vb6Type::Variant) {
        const std::string fldTRN232 = udtFieldTypeRefNameOfTarget(target);
        if (fldTRN232.empty() || cTypeForDeclaredTypeName(fldTRN232) == "vb6_VARIANT") {
            udtField = true;
            return true;
        }
    }
    // 账 #305 §B168: `<对象>.<类的公共 Variant 字段>` / 类内 `Me.<同一字段>` —— 字段档问
    // driver 注入的工程级表 (classFieldIsVariantSlot); member 是方法/属性时否决 (归属性写那一路)。
    if (target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maT = static_cast<const MemberAccessExpr&>(*target);
        const std::string clsT = maT.object ? inferClassTypeOfExpr(*maT.object) : std::string();
        if (!clsT.empty()
            && resolveClassMemberCall(clsT, maT.memberName).empty()
            && classFieldIsVariantSlot(clsT, maT.memberName)) {
            return true;
        }
    }
    return false;
}

// 账 #305 §B169 (第六十刀): 「把这枚右值收进一枚 Variant 槽」(装箱 + 所有权 + 槽位语义)
// 收成这一枚出口。原先住在 visit(AssignmentStmt&) 的 targetIsVariant 那一支; §B168 给类
// 的公共 Variant 字段补上判据之后, With 块那一站还需要**同一串**决定 (自指右值要先收临时量,
// 借来的走 vb6_VariantAssign), 所以两条路共用这一枚, 不再各抄一遍。
void CCodeGen::emitVariantSlotAssign(const std::string& target, Expr* rhs,
                                   const std::string& value, bool udtFieldSlot) {
    // P8.4: 包装值为vb6_VARIANT, 先释放旧BSTR
    std::string wrappedValue = wrapVariantValue(rhs, value);
    // Fix <vbeclipse> rev16: Variant **数组元素**槽位走"接管"语义。
    // `Clear + 浅拷贝` 对"从别处借来的值"不足: RHS 是 ByRef Variant 实参或相邻
    // 元素时, wrapVariantValue 给出的是 vb6_VariantFromValue(...) —— 对 vb6_VARIANT
    // 是 identity **浅拷贝**, 槽与借用方共用同一只 BSTR / 同一个 Dispatch, 借用方
    // 一 Clear 槽即悬垂 (地址被后续同尺寸分配复用 ⇒ 每槽都读"最后一次写入的值")。
    // 实测 tests/ve_list (工程内 List.cls 的 m_Items) idx=0/1/2 全答 "Gamma";
    // 真工程 play78: Folder.Views.Item(0) 取回 1 字符垃圾 → 后续 NULL 解引用。
    // vb6_VariantAssign = Clear(dst) + 深拷贝(BSTR 另分配 / Dispatch AddRef),
    // 与同一个类的 m_Keys (String 数组, 走 vb6_BSTR_Assign) 同一口径。
    bool targetIsVariantArraySlot =
        (target.compare(0, 7, "VB6_SA_") == 0
         && target.find("(vb6_VARIANT,") != std::string::npos)
        // Fix 232: UDT 的 Variant 字段同属"槽位"—— `Clear(&字段)` + 浅拷贝赋值
        // 会把 wrapVariantValue 给出的借用值 (BSTR / Dispatch) 与容器共用,
        // 借用方一释放槽即悬垂 (与 Fix 154 的数组元素同一个理由)。
        || udtFieldSlot;
    if (targetIsVariantArraySlot) {
        c_.emitLine("vb6_VariantAssign(&" + target + ", " + wrappedValue + ");");
    } else if ([&] {
                   // Fix <vbeclipse> 2026-10-12: RHS 文本里是否出现目标自身 (独立标识符)。
                   // 目标不以标识符字符开头 (如 `(*o)` / `VB6_SA_AT(...)`) 时按原样走 ——
                   // 那一族要么已由上面的"槽位"分支接管, 要么不是本缺陷的形态。
                   const std::string& hay = wrappedValue;
                   const std::string& ndl = target;
                   if (ndl.empty()) return false;
                   auto idc = [](char c) {
                       return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')
                              || (c >= '0' && c <= '9') || c == '_';
                   };
                   if (!idc(ndl.front())) return false;
                   size_t p = 0;
                   while ((p = hay.find(ndl, p)) != std::string::npos) {
                       size_t e = p + ndl.size();
                       bool lb = (p == 0) || !idc(hay[p - 1]);
                       bool rb = (e >= hay.size()) || !idc(hay[e]);
                       if (lb && rb) return true;
                       p = e;
                   }
                   return false;
               }()) {
        // Fix <vbeclipse> 2026-10-12: RHS 里出现目标自己时 (`v = CDate(v)`) 不能
        // 「先 Clear 目标再赋值」—— 那样 RHS 读到的是**已被清空的目标** (VT_EMPTY),
        // 静默变成 0。实测 (ComCtlsDemo DTPicker.ctl 的 Property Let Value):
        //   Select Case VarType(DateValue)
        //     Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        //       DateValue = CDate(DateValue)
        // 原发成
        //   vb6_VariantClear(&DateValue);
        //   DateValue = vb6_VariantDate(vb6_CDate(vb6_VariantFromValue(DateValue)));  ← 读 Empty ⇒ 0
        // 于是 CreateDTPicker 结尾的 `Me.Value = PropValue` 把日期变成 0 (1899-12-30),
        // 落到 MinDate (#1/1/1900# = 2.0) 之下 ⇒ 误报
        // Err.Raise 35773 "A date was specified that does not fall within the
        // MinDate and MaxDate properties" (实测页堆下启动即退, RC=189=35773-256×139)。
        // 修法: 先把 RHS 收进栈上临时量, 再清目标并接管。**临时量刻意不 Clear** ——
        // `target = tmp` 是浅拷贝, BSTR/Dispatch 的所有权随拷贝交给目标, 与原来
        // "把内联临时值赋给目标" 的转移语义一致。
        emitVariantSlotStore(target, wrappedValue, true, true);  /* §B164 站⑤ 自指右值: 先收临时量再清+搬(形状不变) */
    } else {
        // §B164 站④ 的第二半句要问一句: 右值**本身是不是一枚活着的槽**。
        // `last = e2`(两边都是 Variant)经 wrapVariantValue 发的是 `vb6_VariantFromValue(e2)`，
        // 而 _Generic 对 vb6_VARIANT 是 identity **浅拷贝** ⇒ 两枚槽共用同一只 BSTR / 同一个
        // x86 DECIMAL 缓冲(与 Rev16 那格数组元素同一个成因)。第五十五刀给 For Each 那一站补上
        // 「先清旧值」之后，这份共用从「静默别名」变成**二次 free**: 实测 test_host_decimal 的 x86
        // 台 12 次跑里 7 次 0xC0000374(BASE 那台 0/12 —— 它只是把每一轮的旧值漏掉，谁也没 free 两次)。
        // 函数返回/集合取回那些是**内联临时量**(按值交回来)，仍走清+搬，不许多拷一次。
        // 这一问收成一处权威(第五十五刀那一形 + 本刀的借来串存储那一形), 站点不再自己抄。
        emitVariantSlotStore(target, wrappedValue,
                             variantRhsOwnsValue(rhs, value));
    }
}

// 账 #305 §B166 (第五十六刀): 被调方在统一出口销毁调用点为这一次调用现造的 ParamArray 数组。
// VB6 的口径 = 这枚数组的寿命就是这一次调用; 安全性已量过: 每个 VB 过程的 PA 实参都由调用点的
// vb6_PA_Create 新造(连 `Inner b` 那种把一枚 ParamArray 再递出去的形状也是重新装箱,
// .build/k56_p/fwd.bas 的产物), 所以这里销毁碰不到别人的指针。
void CCodeGen::emitParamArrayReclaim(const std::vector<std::unique_ptr<ParameterDecl>>& params) {
    for (const auto& prm : params) {
        if (prm && prm->isParamArray) {
            c_.emitLine("vb6_PA_Destroy(" + cIdent(prm->name) + ");");
        }
    }
}


// 账 #305 §B167 (第五十七刀): 「这枚 C 表达式交出的是刚 mint 的串, 还是别人还活着的那只缓冲」。
// 只看**顶层形状**: 内在函数(Left/CStr/…)、拼接(vb6_BSTR_Concat)、工程函数与 Property Get 的返回、
// `Public Const X As String` 折出来的 `vb6_BSTR_FromStr(L"…")` 宏 —— 都是调用 = 现造, 槽可以接管;
// `VB6_SA_AT(BSTR, …)` / `VB6_SA_ND_AT2(BSTR, …)` 是**元素存储**(宏的第一个实参就是元素 C 类型,
// 与左值那一侧 Fix 177/rev11 同一条口径)、裸标识符与字段访问是**存储** —— 都是借来的。
static bool b167TopIsCall(const std::string& s, size_t at) {
    const size_t par = s.find('(', at);
    if (par == std::string::npos) return false;
    const size_t sel = s.find_first_of(".[", at);   // `r.Name` / `me->List(i)` 里的点/括号先出现就不是调用
    return sel == std::string::npos || par < sel;
}

// 账 #305 §B167 形⑭ (第五十八刀): 「顶层是调用 = 现造的」这一腿有一条例外名单 —— RTL 里交回
// **别人那一只还活着的块**的出口。名单是 RTL 逐枚手核出来的(返回 BSTR 的出口共 90 枚, 借来的就这三枚:
// `return vb6_err.description;` / `return vb6_err.source;` / `return v->bstrVal;`, 既不 SysAllocString
// 也不 Copy), 由第 58 道哨兵的 Q6 与 RTL 双向对钉(漏一名、多一名、名单外的新借名都当场红)。
// 这三枚都不许被当成现造的: 槽一旦接管, 槽的 Clear 放掉的是 Err / 宿主自己那只块。
static const char* kBorrowedBstrExports167[] = {
    "vb6_ErrDescription",
    "vb6_ErrSource",
    "vb6_ho_variantToBstr",
};

bool CCodeGen::exprBorrowsBstrStorage(Expr* rhs, const std::string& rawC) const {
    if (!rhs) return false;
    static const std::string kSaBstr167 = "VB6_SA_AT(BSTR,";
    static const std::string kSaNdBstr167 = "VB6_SA_ND_AT2(BSTR,";
    const size_t at167 = rawC.find_first_not_of(" \t\r\n&(");
    if (at167 == std::string::npos) return false;
    if (rawC.compare(at167, kSaBstr167.size(), kSaBstr167) == 0
        || rawC.compare(at167, kSaNdBstr167.size(), kSaNdBstr167) == 0)
        return true;                                  // 串数组元素: 宏自己带元素类型, 不必再问表
    if (rhs->kind == ASTNodeKind::LiteralExpr) return false;
    if (b167TopIsCall(rawC, at167)) {
        const std::string callee167 = rawC.substr(at167, rawC.find('(', at167) - at167);
        for (const char* br167 : kBorrowedBstrExports167)
            if (callee167 == br167) return true;      // 例外: 这一枚调用交回的是借来的块
        return false;                                 // 现造的: 所有权随这一句交给槽
    }
    return inferExprType(*rhs) == Vb6Type::String;
}

bool CCodeGen::variantRhsOwnsValue(Expr* rhs, const std::string& rawC) const {
    if (rhs && rhs->kind == ASTNodeKind::IdentifierExpr) {
        // §B164 形④ (第五十五刀): 右值本身是一枚活着的槽。
        if (knownVariantVars_.count(Symbol::toLower(
                static_cast<const IdentifierExpr&>(*rhs).name)) > 0)
            return false;
    }
    if (rhs && !exprHasLateBoundComRead(rhs)) {
        // 判据只认「这确实是一枚 Variant **字段**」, 不认「源码里长得像左值」:
        //   · 晚绑定成员读 (`at.value`) 交回的是一枚**新造**的 Variant (vb6_VariantFromComResult
        //     / vb6_ComGet*Prop …), 所有权随这一句转移 —— 那不是槽;
        //   · On Error 站现造的 `_vsN` 暂存同理 (转移, 不是槽)。
        // 这两族按槽判会把所有权判反 (cvp_emitc_variant_fallback 与 b251..b255 的形状针钉的
        // 就是它们那一形)。
        //
        // 账 #305 §11.134 (第一百五十三刀): 形④ 的**字段**那一格。
        //
        // `V = mV` (类里 `Private mV As Variant`) 以前落进 true 那一支:
        // `knownVariantVars_` 只登记**本地** Variant, 而类里的字段在 C 里是
        // `me->mV`、在 AST 里是一枚 IdentifierExpr(不带 `me->`) ⇒ 一眼看上去「自有」
        // ⇒ 发 `vb6_VariantClear(&ret); ret = vb6_VariantFromValue(me->mV);`，
        // 而 `vb6_VariantFromValue` 对 vb6_VARIANT 实参是 `vb6_VariantIdentity`
        // (纯浅拷贝: 不 AddRef、不拷 BSTR) ⇒ 调用方拿到的那一枚与字段**共用**同一个对象。
        //
        // 实测 (tests/b153): `Set got = h.V` 之后 `Set got = Nothing` 就把字段还握着的那枚
        // 实例拆了 (Terminate 发生在持有者出厂**之前**), 字段里留一根悬垂指针。
        //
        // 口径: 只要右值是**一枚槽** (本地/形参/字段/UDT 成员，即 inferExprType 答 Variant
        // 的左值形)，就答 false ⇒ 走 `vb6_VariantAssign`(Clear + 深拷贝)。
        // 只有**调用**(函数/属性 Get 交回的临时量)才「自有」—— 那不是左值形，不在这里接管。
        if (rhs->kind == ASTNodeKind::IdentifierExpr) {
            // 裸名 = 本模块字段 (隐式 me->): 问类成员表, 问不到就是别的什么, 不接管。
            const std::string nm =
                Symbol::toLower(static_cast<const IdentifierExpr&>(*rhs).name);
            if (classVariantMembers_.count(nm) || classVariantMembers_.count("m_" + nm))
                return false;
        } else if (rhs->kind == ASTNodeKind::MemberAccessExpr) {
            // `Me.f` / `<类引用>.f`: 问 driver 注入的工程级字段档表。
            auto& ma = static_cast<const MemberAccessExpr&>(*rhs);
            const std::string cls = ma.object ? inferClassTypeOfExpr(*ma.object) : std::string();
            if (!cls.empty() && classFieldIsVariantSlot(cls, ma.memberName)) return false;
        }
    }
    // §B167 形⑪ (本刀): 右值是一处借来的串存储。
    return !exprBorrowsBstrStorage(rhs, rawC);
}


// ============================================================================================
// 账 #464 §B164 回收刀: 「一条语句里现造的带权临时量」提成具名临时量, 语句末按档收回
// ============================================================================================
//
// 这一刀的判据**全部**来自 src/common/owned_temp_keeps.hpp（那张表由
// scripts/check_owned_temp_keeps.ps1 从 RTL **推导** ⇒ 哨兵双向对账 ⇒ 漂不了），
// 本文件只负责**落地**（怎么把那一格提出来、怎么声明、怎么收回）。
//
// 提名的问句（owner = ownedTempNominationVerdict, 四道闸都在头里）:
//   ① 这一枚是不是**现造**的（mint 出口在 [own] 节 ⇒ 新造且归我们所有）
//   ② 消费者那一格会不会把指针**带走**（keeps=0 才提名）
//   ③ 消费者会不会**自己放掉**那一格（消费轴 ⇒ 不提名, 否则双放 ⇒ 堆损坏）
//   ④ 消费者不在 closed 族（未归因 ⇒ 停机）
//
// 档号**只由 mint 出口的返回档定**（BSTR / VARIANT）—— 不是消费者的 VB 形参档,
// 也不是源表达式的 VB 档（#461 拿形参档当 C 档 ⇒ 一枚 double 被声明成 BSTR ⇒ cl C2440;
// 账 #463 已把 [own] 表与 RTL 原型的返回档逐枚对账 = 112/112 一致）。
namespace {
// 把一段发码文本按**顶层逗号**切成实参（顶层 = 括号深度 0）。
std::vector<std::string> splitTopLevelArgs464(const std::string& s) {
    std::vector<std::string> out;
    int depth = 0;
    size_t start = 0;
    for (size_t i = 0; i < s.size(); ++i) {
        const char ch = s[i];
        if (ch == '(' || ch == '[' || ch == '{') ++depth;
        else if (ch == ')' || ch == ']' || ch == '}') --depth;
        else if (ch == ',' && depth == 0) {
            out.push_back(s.substr(start, i - start));
            start = i + 1;
        }
    }
    out.push_back(s.substr(start));
    for (auto& a : out) {
        const size_t b = a.find_first_not_of(" \t\r\n");
        const size_t e = a.find_last_not_of(" \t\r\n");
        a = (b == std::string::npos) ? std::string() : a.substr(b, e - b + 1);
    }
    return out;
}
// 「整枚实参 = 一个调用」？返回 callee 名（空 = 不是）。
std::string topLevelCallee464(const std::string& arg) {
    size_t i = 0;
    while (i < arg.size() && (std::isalnum(static_cast<unsigned char>(arg[i])) || arg[i] == '_')) ++i;
    if (i == 0 || i >= arg.size() || arg[i] != '(') return std::string();
    const std::string callee = arg.substr(0, i);
    // 括号必须**配平到末尾**（`f(a)(b)` / `f(a) + g(b)` 都不算「整枚就是一个调用」）
    int depth = 0;
    for (size_t j = i; j < arg.size(); ++j) {
        if (arg[j] == '(') ++depth;
        else if (arg[j] == ')') {
            --depth;
            if (depth == 0) return (j + 1 == arg.size()) ? callee : std::string();
        }
    }
    return std::string();
}
// 这一枚实参里**有没有调用**？（闸 ④ 用：别的实参里有调用 ⇒ 求值次序会动 ⇒ 不提名）
bool argHasCall464(const std::string& arg) {
    for (size_t i = 0; i + 1 < arg.size(); ++i) {
        if (arg[i] != '(') continue;
        // 往前找标识符（`f(` 是调用；`(x)` 是括号；`)(` 不是）
        size_t j = i;
        while (j > 0 && (std::isalnum(static_cast<unsigned char>(arg[j - 1])) || arg[j - 1] == '_')) --j;
        if (j == i) continue;                       // 前面没有标识符 ⇒ 不是调用
        return true;
    }
    return false;
}
// 这一枚文本**整枚**是不是一个「owned 且档是 BSTR」的 mint 调用？
//   （账 #468 §B164 拼接站那一族：左操作数若是这样一枚，拼接站**自己**放掉它 ⇒ 不漏）
// 判据全在 src/common/owned_temp_keeps.hpp 一个权威里：档由 mint 出口的 RTL 返回档定
//   （BSTR 档 ⇒ 拼接站发 ConcatFree 时 vb6_BSTR_Free 放的就是它 ⇒ 放一次，不双放）。
}  // namespace

bool CCodeGen::exprIsOwnedBstrMint(const std::string& text) {
    if (!currentProc_) return false;
    const std::string mint = topLevelCallee464(text);
    if (mint.empty()) return false;
    return ownedTempMintKind(mint.c_str()) == 'B';
}

// 提名**指定那一格**（拼接站的右操作数用；实参站用整张表那一版）。
// 这一格改写成具名临时量; 返回 true = 提了名。
bool CCodeGen::hoistOwnedTempSlot(std::string& arg, const std::string& consumer, int slot,
                                  const std::string& siblingText) {
    if (!currentProc_) return false;
    if (concatLeftDepth_ > 0) return false;
    // 门 ⑨: **块内条件**不许提名（声明落在块内、收回落在块外 ⇒ 引用越界）
    if (blockCondDepth_ > 0) return false;
    // 门 ⑩: 提了名 ⇒ 这一枚**提前到调用之前**求值 ⇒ 若**旁的那一枚**（拼接站的左操作数）
    //   里有调用, 两者的求值次序就动了（C 里实参次序本来是不定的, 但基线在 x64 上
    //   事实上是左到右, 提了名会变成右到左 ⇒ 行为变了）⇒ 那一格有调用就不提名。
    //   口径与闸 ⑤（实参站: 别的实参里有调用 ⇒ 不提名）同一条。
    if (!siblingText.empty() && argHasCall464(siblingText)) return false;
    // 闸 ⑦: 声明与收回必须**同一层**（这一枚临时量的作用域就在这条语句这一层）
    if (c_.indentLevel() != stmtLevel_) return false;
    if (consumer.empty()) return false;
    const std::string mint = topLevelCallee464(arg);
    if (mint.empty()) return false;
    const char verdict = ownedTempNominationVerdict(consumer.c_str(), slot, mint.c_str());
    if (!verdict) return false;
    const std::string tmp = "_own_" + std::to_string(++ownedTempCounter_);
    const bool isVariant = (verdict == 'V');
    c_.addPending((isVariant ? "vb6_VARIANT " : "BSTR ") + tmp + " = " + arg + ";");
    ownedTempsToFree_.push_back(OwnedTemp{tmp, c_.indentLevel(), isVariant ? 2 : 1});
    arg = tmp;
    return true;
}
bool CCodeGen::hoistOwnedTempArgs(std::string& argList, const std::string& callee) {
    // 闸 ⑥: 文件作用域没有「语句末」可以收回 ⇒ 不提名
    if (!currentProc_) return false;
    // 闸 ⑤: 拼接站左操作数里的一切不提名（那条子串判据会翻面 ⇒ 行为变了）
    if (concatLeftDepth_ > 0) return false;
    // 门 ⑨: **块内条件**不许提名（ElseIf 的条件：声明落在 else 块内、收回落在块外）
    if (blockCondDepth_ > 0) return false;
    // 闸 ⑦: 声明与收回必须**同一层**（这一枚临时量的作用域就在这条语句这一层）
    if (c_.indentLevel() != stmtLevel_) return false;
    if (callee.empty()) return false;
    std::vector<std::string> args = splitTopLevelArgs464(argList);
    if (args.empty()) return false;
    // 闸 ④: **只提名「别的实参都没有副作用」的那一格** ——
    //   多实参时把一枚实参提成临时量会**改动求值次序**（那一枚提前到调用之前求值）。
    //   基线里实参的求值次序本来就是编译器定的（C89 不定序；MSVC x86 从右往左），
    //   所以只有当**别的实参都是无副作用的**（纯字面量 / 纯读一个变量 / 纯成员读，
    //   也就是**不含任何调用**）时，次序改了也看不出差别 ⇒ 这才提名；
    //   其余情形（别的实参里有调用）一律不提名（宁可不收）。
    bool changed = false;
    for (size_t slot = 0; slot < args.size(); ++slot) {
        const std::string mint = topLevelCallee464(args[slot]);
        if (mint.empty()) continue;
        bool othersPure = true;
        for (size_t j = 0; j < args.size() && othersPure; ++j) {
            if (j == slot) continue;
            if (argHasCall464(args[j])) othersPure = false;
        }
        if (!othersPure) continue;
        // 判据那四道闸, 都在头里那位主人身上
        const char verdict = ownedTempNominationVerdict(callee.c_str(), (int)slot, mint.c_str());
        if (!verdict) continue;
        // 提出来: 声明走 pending（落在调用那一行之前）, 收回登记在语句末
        const std::string tmp = "_own_" + std::to_string(++ownedTempCounter_);
        const bool isVariant = (verdict == 'V');
        c_.addPending((isVariant ? "vb6_VARIANT " : "BSTR ") + tmp + " = " + args[slot] + ";");
        ownedTempsToFree_.push_back(OwnedTemp{tmp, c_.indentLevel(), isVariant ? 2 : 1});
        args[slot] = tmp;
        changed = true;
    }
    if (!changed) return false;
    argList.clear();
    for (size_t slot = 0; slot < args.size(); ++slot) {
        if (slot) argList += ", ";
        argList += args[slot];
    }
    return true;
}
// C30-10 第三十二刀: **常整数**的 VB6 类型宽度 (字面量类型, 不是常量折叠)。
// VB6 的整数字面量不是「一律 Long」: 无后缀十进制、值落在 Integer 范围 ⇒ Integer
// (VarType 2 / TypeName "Integer")。C3 的 inferExprType 对字面量一律答 Long —— 那是
// C 侧 int32 的口径 (收窄检查要用), 于是装箱站把 `1` 装成 VT_I4 (夹具 tests/dict_native
// 的 d121: int=3, 真组件 2)。这里只回答「这枚节点是不是一枚**字面量**整数、值是多少」,
// 分档仍由装箱站那一处给。
// ⚠ 刻意**不**用 tryEvalConstInt: 那一枚是常量折叠 (它把 `2 ^ 10` 也折成 1024, 而 VB6
//   的 `^` 恒返 Double), 拿它当字面量类型会把幂/位运算的档一起改掉。
// ⚠ 也不进 inferExprType: 那张表同时喂算术/收窄/格式化, 改它就是把「16 位算术 +
//   溢出检查」那一整层一起打开 (那是「算术全链」, 另排期)。
// 不认的写法 (维持现状, 台账另记): 带类型后缀的写法 (%/&/^ 等)；`&O`/`&B` 进制字面量
//   (真组件只对 `&H` 量到读数, 其余两形仍是 Long 档 —— 留白列在 ai/039 §11.12)。
//   `&H` 那一形已经接上：值早在第313刀按 16 位补码折过, 第三十三刀只把**档**接过来
//   —— 两个字都不是在编��站重算的, 而是 lexer 在 Token::radixInt16 上一次性答完、
//   parser 原样搬到 LiteralExpr::radixInt16。
static bool vb6IntConstantValue(const Expr* e, int32_t* outVal) {
    if (!e) return false;
    if (e->kind == ASTNodeKind::UnaryExpr) {
        auto& u = static_cast<const UnaryExpr&>(*e);
        if (u.op != UnaryOp::Negate || !u.operand) return false;
        int32_t inner = 0;
        if (!vb6IntConstantValue(u.operand.get(), &inner)) return false;
        if (outVal) *outVal = -inner;
        return true;
    }
    if (e->kind != ASTNodeKind::LiteralExpr) return false;
    auto& lit = static_cast<const LiteralExpr&>(*e);
    if (lit.literalKind != LiteralKind::Integer && lit.literalKind != LiteralKind::Long)
        return false;
    // 只认纯十进制书写: rawText 里除数字与空白不许有别的东西 (进制前缀 &H/&O、类型后缀
    // %/&、小数点、指数全在这里被挡下)。
    // C30-10 第三十四刀: `3%` 是 VB6 的**显式 Integer** 字面量 (账 #195 已把它认成
    // IntegerLiteral), 但文本带着 `%`, 上面那道闸把它挡在外面 ⇒ `VarType(3%)` 答 3。
    // 现在剥掉**一个**尾随 `%` 再走同一道闸; `&`/`^` 仍挡在外面 (Long / LongPtr)。
    if (lit.rawText.empty()) return false;
    std::string digits = lit.rawText;
    while (!digits.empty() && std::isspace(static_cast<unsigned char>(digits.back())))
        digits.pop_back();
    if (!digits.empty() && digits.back() == '%') digits.pop_back();
    if (digits.empty()) return false;
    for (char ch : digits) {
        if (std::isspace(static_cast<unsigned char>(ch))) continue;
        if (ch < '0' || ch > '9') return false;
    }
    const int64_t v = (lit.literalKind == LiteralKind::Integer)
                          ? static_cast<int64_t>(lit.intValue) : lit.longValue;
    if (v < -2147483648LL || v > 2147483647LL) return false;
    if (outVal) *outVal = static_cast<int32_t>(v);
    return true;
}

// C30-10 第三十二刀: 「这枚节点在 VB6 里是 **Integer** 档 (VarType 2)、而 inferExprType 答
// Long」的那两形, 收成一句判据 —— 字面量那一半见上面 vb6IntConstantValue 的话;
// `As Integer` 那一半靠 knownIntVars_ : 符号表把 Integer 记成 Long (C 侧另发 int16_t),
// 真宽度在那张表里 (inferExprType 也从它答 Integer, 见 cgen_util_type.cpp)。
bool CCodeGen::isNarrowIntExpr(const Expr* e) const {
    if (!e) return false;
    // C30-10 第五十三刀: 一元负号**先于常量折叠**判 —— 采集器 U 的 U13 `-32768` 在 VB6 里是
    // **Long**（量字面 32768 自己就装不进 Integer），而 U14 `-iMin`（Integer 变量）仍是 Integer。
    // 以前走下面的 vb6IntConstantValue 把负号折进去 ⇒ -32768 落在 int16 范围 ⇒ 错发 Integer。
    if (e->kind == ASTNodeKind::UnaryExpr) {
        const auto& unNeg = static_cast<const UnaryExpr&>(*e);
        if (unNeg.op == UnaryOp::Negate && unNeg.operand)
            return isNarrowIntExpr(unNeg.operand.get());
    }
    int32_t v = 0;
    if (vb6IntConstantValue(e, &v)) return v >= -32768 && v <= 32767;
    // C30-10 第三十三刀: 进制字面量那一半 —— 判据不在这里, 在 lexer (Token::radixInt16)。
    // 这里只问那一枚标记。范围检查是**反证**而不是判据: 走过 16 位补码折叠的值必定在
    // int16 里, 若哪天不在, 宁可退回 Long 档也不发出一个错的 (int16_t) 截断。
    // （`&HFFFFFFFF` 那种 5~8 位的写法压根没有标记 —— 它在 VB6 里是 Long 的 -1。）
    if (e->kind == ASTNodeKind::LiteralExpr) {
        auto& lit = static_cast<const LiteralExpr&>(*e);
        return lit.radixInt16 && lit.literalKind == LiteralKind::Integer
               && lit.intValue >= -32768 && lit.intValue <= 32767;
    }
    // C30-10 第四十五刀: `Not <Variant>` 的结果在 VB6 里恒为 **Integer**（用户真 VB6 日志
    // E06: `v = 5` 时 `VarType(Not v)` = 2、值 -6）。而 `Not <Boolean>` 是 **Boolean**
    // （E05: 11/False）—— 两形不能混谈，所以这里只认「操作数本身是 Variant」那一形。
    // 值那一半早就是对的（发码 `~vb6_VariantToLong(v)`），缺的只是装箱时的档。
    if (e->kind == ASTNodeKind::UnaryExpr) {
        auto& un = static_cast<const UnaryExpr&>(*e);
        if (!un.operand) return false;
        // `- <Integer 档>` → Integer（用户日志 E10/H06: `-i`(Integer) = 2/Integer）。
        // 以前发码是 `(-i)`（int16 在 C 里提成 int）→ 装箱当 Long。
        // ⚠ 已知差: `-(-32768)` 在 VB6 报错误 6，这里只会回卷 —— 编译期看不到值，
        //   要真报错得在 RTL 那条路上（本轮不开）。
        if (un.op == UnaryOp::Negate) return isNarrowIntExpr(un.operand.get());
        if (un.op == UnaryOp::Not) {
            // `Not <Variant>` / `Not <Integer 档>` → Integer（E06: -6/Integer；E08: `Not 0`
            // 也是 2/Integer）。而 `Not <Boolean>` 是 Boolean（E05: 11/False）——
            // Boolean 不在 narrow 表里，所以那一条自动落在外面。
            return inferExprType(*un.operand) == Vb6Type::Variant
                   || isNarrowIntExpr(un.operand.get());
        }
        return false;
    }
    // C30-10 第四十六刀: `+ - * \ Mod` 两侧都是 Integer 档 ⇒ 结果仍是 **Integer**
    // （用户真 VB6 日志 F01~F04 / F06 / F06b / F09 / F11 / F12 / F14 / F20 / G09 /
    // H01 / H02 全是 2/Integer）。溢出那一半不在这句判据里 —— 值由发码侧每一步
    // 的 vb6_ChkInt 查（F05/F07/F08/F19 在 VB6 是错误 6），这里只答「档」。
    // `/` 与 `^` 刻意不在内: VB6 那两条返回 Double（F10/F13）。
    if (e->kind == ASTNodeKind::BinaryExpr) {
        auto& bin = static_cast<const BinaryExpr&>(*e);
        switch (bin.op) {
        case BinaryOp::Add:
        case BinaryOp::Sub:
        case BinaryOp::Mul:
        case BinaryOp::IntDiv:
        case BinaryOp::Mod:
            return isNarrowIntExpr(bin.left.get()) && isNarrowIntExpr(bin.right.get());
        default:
            return false;
        }
    }
    // C30-10 第三十五刀: 内建返回档那一族 (Asc/Sgn/VarType/StrComp/月日年时分秒/Weekday/
    // FreeFile) —— VB6 答 Integer, 而语义层与 C 侧都答 int32。两形都要认:
    //   * 带实参的调用形 `Asc("A")` → IndexOrCallExpr(callee=IdentifierExpr)；
    //   * 零参、不带括号的 `FreeFile` → 就是一个 IdentifierExpr。后者必须先排除用户
    //     变量同名的情况(符号表里是 Variable), 否则一枚名叫 `month` 的普通变量会被
    //     白白截成 int16。
    // ⚠ 标识符这一支只能有一个出口: 以前这里先写了 `if (kind==IdentifierExpr) return
    //   knownIntVars_||knownIntConsts_;` 的提前返回，于是后面新加的内建支**永远走不到**
    //   (`v = FreeFile` 实测仍答 3)。现在三源合一、一次答完。
    if (e->kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<const IdentifierExpr&>(*e);
        const std::string lower = Symbol::toLower(id.name);
        if (knownIntVars_.count(lower) > 0) return true;
        // C30-10 第三十四刀: 用户 `Const` 的 Integer 档。两张表只在这里被问 ——
        // 没有进 knownIntVars_, 所以 inferExprType / 收窄检查 / 算术那几层一字未动。
        if (knownIntConsts_.count(lower) > 0) return true;
        if (!builtinReturnsVbIntegerName(id.name)) return false;
        // 遮蔽判据: 用户同名的变量/属性必须是 Variable/Property*, 那就**不**算内建。
        // 内置符号自己标了 isBuiltin (semantic_analyzer_builtin 那两枚 lambda), 但
        // 零参内置在 codegen 的符号表里未必查得到 (实测 `FreeFile` 裸引用查回 nullptr,
        // 而 `FreeFile()` 走调用形是对的) —— 所以 nullptr 也按内建看。
        const Symbol* sym = symTab_.lookup(id.name);
        return !sym || sym->isBuiltin;
    }
    if (e->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& call = static_cast<const IndexOrCallExpr&>(*e);
        if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr) {
            auto& cid = static_cast<const IdentifierExpr&>(*call.callee);
            if (builtinReturnsVbIntegerName(cid.name)) return true;
            // C30-10 第三十八刀: `Int`/`Fix`/`Abs` 的返回档 = **实参的档**（VB6 口径；
            // cscript 实测 `VarType(Int(5))`=2、`Fix(5)`=2、`Abs(-5)`=2，而
            // `Int(1.7)`=5、`Abs(1.5)`=5 —— 后两个本来就是对）。
            // 只对「实参本身就是 Integer 档」那一形答 true；Variant 实参靠运行时子类型
            // 传播（`Abs(v)`，v 装 5 时答 2），静态不可知，留白。
            const std::string n = Symbol::toLower(cid.name);
            if ((n == "int" || n == "fix" || n == "abs") && call.positional.size() == 1
                && call.named.empty())
                return isNarrowIntExpr(call.positional[0].get());
            // C30-10 第三十九刀: `IIf(cond, a, b)` —— VB6 返回 Variant，**子类型随被选中
            // 的那一支**（用户真 VB6 实测：`TypeName(IIf(True, 1, 2))` = Integer）。
            // 静态只在「两支都是 Integer 档」时答 Integer；一支 Long 一支 Integer 时结果
            // 取决于运行时走哪支，宁可答 false（退回现在的 Long 箱）也不发一个可能错的窄箱。
            // ⚠ 不能把 `iif` 直接塞进 builtinReturnsVbInteger：那会连
            // `IIf(True, 1&, 2&)` / `IIf(True, 1.5, 2.5)` 一起窄化。
            // Choose/Switch 同构，但本机没有 oracle（VBScript 无此三枚），先不开。
            if (n == "iif" && call.positional.size() == 3 && call.named.empty())
                return isNarrowIntExpr(call.positional[1].get())
                       && isNarrowIntExpr(call.positional[2].get());
            // C30-10 第五十五刀: Choose/Switch 一律发变体三目（选中支原样 + 越界选 Null，
            // S 段 16 行全靠它），子类型是运行期的 ⇒ 不答窄档（答 true 会让存储站发
            // `vb6_VariantInt((int16_t)(<整只 VARIANT>))` = C2440）。
            return false;
        }
        return false;
    }
    return false;
}

// C30-11 第六十四刀: Abs/Int/Fix/Hex/Oct 的「Null 进、Null 出」变体位判据。
//
// 这两站（wrapVariantValue 存储站 / boxToVariant 实参站）都要先问这一句，且在**各自的
// inferExprType switch 之前**问：那张表对 Int/Fix 答 Double、对 Hex/Oct 答 String，
// 一旦先返回 vb6_VariantDouble(...)/vb6_VariantString(...)，就再也到不了这里了
// （本刀第一版只改了 boxToVariant，实测发码一字未变，就是被 wrapVariantValue 的
//  `case Vb6Type::Double` 提前收掉了）。
//
// 判据是**文本前缀**而不是 inferExprType(实参)==Variant：`vb6_Abs(vb6_VariantToDouble(`
// 这串本身就证明上游那一站真把 Variant 槽折进来了，而那次推断对 `Dim nv As Variant`
// 的局部答不出 Variant（第五十二刀那条 AbsV 分支就是这么被挡掉的）。
static std::string mathNullFamilyV(const ASTNode* node, const std::string& cExpr) {
    if (!node || node->kind != ASTNodeKind::IndexOrCallExpr) return std::string();
    const auto& call = static_cast<const IndexOrCallExpr&>(*node);
    if (!call.callee || call.callee->kind != ASTNodeKind::IdentifierExpr) return std::string();
    const std::string fn = Symbol::toLower(
        static_cast<const IdentifierExpr&>(*call.callee).name);
    // §11.207 (第二百二十四刀): Round 带 0..1 个位数实参，下面那张单参表接不住。
    // 真机 .temp/k224p/P298 / P300 的 vI..vN 七格：变体载体上的 Round 也保子类型。
    // 作法：从 `vb6_Round(vb6_VariantToDouble(` 之后按括号配平取出第一实参正文，
    // 剩下尾巴是 ")" 位数就缺省为 0，是 ", n)" 就照抄——与表里那四枚同一个出口。
    if (fn == "round" && call.positional.size() <= 2) {
        static const std::string preR = "vb6_Round(vb6_VariantToDouble(";
        size_t sR = 0;
        while (sR < cExpr.size() && (cExpr[sR] == '(' || cExpr[sR] == ' ')) sR++;
        if (cExpr.compare(sR, preR.size(), preR) == 0) {
            const size_t pR = sR + preR.size();
            int depthR = 1;
            size_t kR = pR;
            for (; kR < cExpr.size(); ++kR) {
                if (cExpr[kR] == '(') depthR++;
                else if (cExpr[kR] == ')') { if (--depthR == 0) break; }
            }
            if (kR < cExpr.size()) {
                std::string innerR = cExpr.substr(pR, kR - pR);
                std::string tailR = cExpr.substr(kR + 1);
                if (!tailR.empty() && tailR.back() == ')') tailR.pop_back();
                if (tailR.empty()) tailR = ", 0";
                return "vb6_RoundV(" + innerR + tailR + ")";
            }
        }
        return std::string();
    }
    if (call.positional.size() != 1) return std::string();
    static const struct { const char* vb; const char* pre; const char* out; } kTbl[] = {
        {"abs", "vb6_Abs(vb6_VariantToDouble(", "vb6_AbsV("},
        {"int", "vb6_Int(vb6_VariantToDouble(", "vb6_IntV("},
        {"fix", "vb6_Fix(vb6_VariantToDouble(", "vb6_FixV("},
        {"hex", "vb6_Hex(vb6_VariantToLong(",   "vb6_HexV("},
        {"oct", "vb6_Oct(vb6_VariantToLong(",   "vb6_OctV("},
    };
    for (const auto& m : kTbl) {
        if (fn != m.vb) continue;
        const std::string pre = m.pre;
        size_t s = 0;
        while (s < cExpr.size() && (cExpr[s] == '(' || cExpr[s] == ' ')) s++;
        if (cExpr.compare(s, pre.size(), pre) == 0
            && cExpr.size() >= s + pre.size() + 2 && cExpr.back() == ')') {
            // 尾部两层右括号（一层提取、一层调用）一并剥掉。
            std::string inner = cExpr.substr(s + pre.size(),
                                             cExpr.size() - s - pre.size() - 2);
            return std::string(m.out) + inner + ")";
        }
    }
    return std::string();
}

// C30-11 第六十五刀: 比较符的 Null 传播。
//
// 读数 .temp/null/oracle_ops.vbs O5..O10：`= <> < > <= >=` 任一侧是 Null 就是 **Null**
// （含 Null = Null，改前我们答 True）。Q12 = CStr(Null="") ⇒ 94（Null 转字符串那一站）。
//
// 为什么不在比较发码处直接换成变体位：比较的 C 类型现在是 int32_t，`If` 条件、
// 赋值给 Boolean、And/Or 都按标量吃它；直接换会连带一整片调用方。
// 本刀只动“**变体上下文**”那两站（wrapVariantValue 存储站 / boxToVariant 实参站），
// 与第六十四刀同一形：已 emit 的 bool 结果**不动**，只把槽地址插进去一层守卫。
//
// 判据：① AST 是比较运算；② 发码文本是 `vb6_VariantBool(<bool>)`；
//   ③ 能从里面抽出至少一个 **Variant 槽地址**（两种形：
//   `vb6_VarCmp*(&nv, …)` 的 `&nv`，或 `vb6_VariantFromValue(nv)` 里的 nv）。
// 三条不齐就不动（宁可不变，不让发码形状凭空变）。
static bool isCmpOp65(BinaryOp op) {
    return op == BinaryOp::Eq || op == BinaryOp::Neq || op == BinaryOp::Lt
        || op == BinaryOp::Gt || op == BinaryOp::Le || op == BinaryOp::Ge;
}

// 从 `vb6_VariantFromValue(nv)` / `(&nv` 两种形里抽槽地址，最多两枚。
static void collectSlotAddrs65(const std::string& s, std::vector<std::string>& out) {
    auto pushIfNew = [&out](const std::string& a) {
        if (a.empty() || out.size() >= 2) return;
        for (const auto& e : out) if (e == a) return;
        out.push_back(a);
    };
    auto isIdentStart = [](char c) {
        return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == '_';
    };
    auto isIdentChar = [&](char c) {
        return isIdentStart(c) || (c >= '0' && c <= '9');
    };
    static const std::string kFromVal = "vb6_VariantFromValue(";
    for (size_t p = s.find(kFromVal); p != std::string::npos;
         p = s.find(kFromVal, p + 1)) {
        size_t q = p + kFromVal.size();
        if (q >= s.size() || !isIdentStart(s[q])) continue;
        size_t e = q;
        while (e < s.size() && isIdentChar(s[e])) e++;
        if (e < s.size() && s[e] == ')') pushIfNew("&" + s.substr(q, e - q));
    }
    // `(&ident,` / `(&ident)` ：只认变体比较那一族（vb6_VarCmp*）的实参位。
    for (size_t p = s.find("(&"); p != std::string::npos; p = s.find("(&", p + 1)) {
        size_t q = p + 2;
        if (q >= s.size() || !isIdentStart(s[q])) continue;
        size_t e = q;
        while (e < s.size() && isIdentChar(s[e])) e++;
        if (e < s.size() && (s[e] == ',' || s[e] == ')' || s[e] == ' '))
            pushIfNew("&" + s.substr(q, e - q));
    }
}

std::string cgNullGuardCmpV(const ASTNode* node, const std::string& cExpr) {
    if (!node || node->kind != ASTNodeKind::BinaryExpr) return std::string();
    // 幂等：实参站会被套两层（boxToVariant 之后又有一层装箱），第二层拿到的
    // 已经是 vb6_VCmpNullV(...)；再包一次会变成 VARIANT 当 int32_t → C2440（实测）。
    if (cExpr.find("vb6_VCmpNullV(") != std::string::npos) return std::string();
    const auto& bin = static_cast<const BinaryExpr&>(*node);
    if (!isCmpOp65(bin.op)) return std::string();

    size_t s = 0;
    // 只跳空格：不能跳 ( —— 裸形的首字符就是配对括号，跳掉它前缀判据就全部落空
    // （第六十七刀第一版就是这么静默不命中的）。
    while (s < cExpr.size() && cExpr[s] == ' ') s++;
    static const std::string kWrap = "vb6_VariantBool(";
    std::string inner;
    if (cExpr.compare(s, kWrap.size(), kWrap) == 0 && cExpr.back() == ')') {
        // 存储站那一支：拿到的是已经包好的 `vb6_VariantBool(<bool>)`。
        inner = cExpr.substr(s + kWrap.size(),
                             cExpr.size() - s - kWrap.size() - 1);
    } else {
        // 实参站那一支：拿到的是**裸 bool 表达式**（vb6_VariantBool 还没包），
        // 直接把它当条件交给守卫 —— 守卫自己产变体。
        inner = cExpr;
    }

    std::vector<std::string> slots;
    collectSlotAddrs65(inner, slots);
    if (slots.empty()) return std::string();
    const std::string a = slots[0];
    const std::string b = slots.size() > 1 ? slots[1] : "(vb6_VARIANT*)0";
    return "vb6_VCmpNullV(" + a + ", " + b + ", " + inner + ")";
}

namespace {
// 下面四枚只服务 cgNullGuardConcatV 的“链”解析（第六十九刀）。
bool cgIsIdentText(const std::string& s) {
    if (s.empty()) return false;
    for (char ch : s) {
        bool ok = (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z')
            || (ch >= '0' && ch <= '9') || ch == '_';
        if (!ok) return false;
    }
    return true;
}

bool cgIsVariantValuedText(const std::string& s) {
    if (cgIsIdentText(s)) return true;
    if (s.compare(0, 20, "vb6_VarConcatNullVV(") == 0) return true;
    if (s.compare(0, 19, "vb6_VarConcatNullV(") == 0) return true;
    return false;
}

// `vb6_CStr(vb6_VariantFromValue(X))` → X（X 允许是任意括号配平的 Variant 值表达式）
bool cgVariantOfCStr(const std::string& s, std::string& x) {
    static const std::string kPre = "vb6_CStr(vb6_VariantFromValue(";
    if (s.compare(0, kPre.size(), kPre) != 0) return false;
    if (s.size() < kPre.size() + 3 || s.back() != ')') return false;
    if (s.compare(s.size() - 2, 2, "))") != 0) return false;
    x = s.substr(kPre.size(), s.size() - kPre.size() - 2);
    int depth = 0;
    for (size_t i = 0; i < x.size(); ++i) {
        if (x[i] == '(') ++depth;
        else if (x[i] == ')') { if (--depth < 0) return false; }
        else if (x[i] == '"') { for (++i; i < x.size() && x[i] != '"'; ++i) {} }
    }
    return depth == 0;
}

// 把 `vb6_BSTR_ConcatFree(A, B)` 的左结合树拆成叶子表（发码的链就是这形：
// `ConcatFree(ConcatFree(a,b),c)`），叶子必须是整只 Variant 的 CStr。
bool cgFlattenConcatLeaves(const std::string& s, std::vector<std::string>& out) {
    static const std::string kCF = "vb6_BSTR_ConcatFree(";
    size_t sp = 0;
    while (sp < s.size() && s[sp] == ' ') sp++;
    const std::string body = s.substr(sp);
    if (body.compare(0, kCF.size(), kCF) == 0 && !body.empty() && body.back() == ')') {
        const std::string args = body.substr(kCF.size(), body.size() - kCF.size() - 1);
        size_t comma = std::string::npos;
        int depth = 0;
        for (size_t i = 0; i < args.size(); ++i) {
            char ch = args[i];
            if (ch == '(') ++depth;
            else if (ch == ')') { if (--depth < 0) return false; }
            else if (ch == '"') { for (++i; i < args.size() && args[i] != '"'; ++i) {} }
            else if (ch == ',' && depth == 0) { comma = i; break; }
        }
        if (comma == std::string::npos || depth != 0) return false;
        return cgFlattenConcatLeaves(args.substr(0, comma), out)
            && cgFlattenConcatLeaves(args.substr(comma + 1), out);
    }
    std::string x;
    if (!cgVariantOfCStr(body, x) || !cgIsVariantValuedText(x)) return false;
    out.push_back(x);
    return true;
}
}  // namespace

std::string cgNullGuardConcatV(const ASTNode* node, const std::string& cExpr) {
    if (!node || node->kind != ASTNodeKind::BinaryExpr) return std::string();
    // 幂等：同上，第二层拿到的已经是 vb6_VarConcatNullV(...)，不能再包。
    if (cExpr.find("vb6_VarConcatNullV(") != std::string::npos) return std::string();
    if (cExpr.find("vb6_VarConcatNullVV(") != std::string::npos) return std::string();
    const auto& bin = static_cast<const BinaryExpr&>(*node);
    if (bin.op != BinaryOp::Concat) return std::string();

    // 只认「两侧都是整枚 Variant 槽转出的 CStr」这一形；其它形（字面量/数值/嵌套拼接）
    // 一律不动 —— 那些格的真值本来就等于零期值（读数 R4..R6/R10/R12）。
    static const std::string kSlotPre = "vb6_CStr(vb6_VariantFromValue(";
    // 从 vb6_CStr(vb6_VariantFromValue(X)) 里取回 X。第六十九刀之前只认 X = 一个标识符，
    // 所以 `(Null & Null) & Null` 这种链永远不命中（左操作数是内层守卫的调用）。
    // 现在 X 允许是任意 Variant 值表达式（括号要配平）。
    auto varOf = [](const std::string& s) -> std::string {
        if (s.compare(0, kSlotPre.size(), kSlotPre) != 0) return std::string();
        if (s.size() < kSlotPre.size() + 3 || s.back() != ')') return std::string();
        if (s.compare(s.size() - 2, 2, "))") != 0) return std::string();
        std::string x = s.substr(kSlotPre.size(), s.size() - kSlotPre.size() - 2);
        int depth = 0;
        for (size_t i = 0; i < x.size(); ++i) {
            if (x[i] == '(') ++depth;
            else if (x[i] == ')') { if (--depth < 0) return std::string(); }
            else if (x[i] == '"') { for (++i; i < x.size() && x[i] != '"'; ++i) {} }
        }
        if (depth != 0) return std::string();
        return x;
    };
    auto isIdent = [](const std::string& s) -> bool {
        if (s.empty()) return false;
        for (char ch : s) {
            bool ok = (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z')
                || (ch >= '0' && ch <= '9') || ch == '_';
            if (!ok) return false;
        }
        return true;
    };

    size_t s = 0;
    // 只跳空格：不能跳 ( —— 裸形的首字符就是配对括号，跳掉它前缀判据就全部落空
    // （第六十七刀第一版就是这么静默不命中的）。
    while (s < cExpr.size() && cExpr[s] == ' ') s++;
    std::string body;
    static const std::string kVStr = "vb6_VariantString(";
    if (cExpr.compare(s, kVStr.size(), kVStr) == 0 && cExpr.back() == ')') {
        body = cExpr.substr(s + kVStr.size(), cExpr.size() - s - kVStr.size() - 1);
    } else {
        body = cExpr.substr(s);
    }

    static const std::string kConcat = "vb6_BSTR_ConcatFree(";
    if (body.compare(0, kConcat.size(), kConcat) != 0 || body.back() != ')')
        return std::string();
    std::string args = body.substr(kConcat.size(), body.size() - kConcat.size() - 1);
    // 只在「两个实参都是槽 CStr」时分：按 `, ` 切一次，且两侧都不能再含未配对括号。
    size_t comma = std::string::npos;
    int depth = 0;
    for (size_t i = 0; i < args.size(); ++i) {
        char ch = args[i];
        if (ch == '(') ++depth;
        else if (ch == ')') { if (--depth < 0) return std::string(); }
        else if (ch == '"') {  // 跳过字符串字面量
            for (++i; i < args.size() && args[i] != '"'; ++i) {}
        } else if (ch == ',' && depth == 0) { comma = i; break; }
    }
    if (comma == std::string::npos || depth != 0) return std::string();
    std::string a = args.substr(0, comma);
    std::string b = args.substr(comma + 1);
    auto trim = [](std::string& x) {
        size_t p = x.find_first_not_of(' ');
        size_t q = x.find_last_not_of(' ');
        x = (p == std::string::npos) ? std::string() : x.substr(p, q - p + 1);
    };
    trim(a);
    trim(b);
    std::vector<std::string> leaves;
    if (!cgFlattenConcatLeaves(body, leaves) || leaves.size() < 2) return std::string();
    if (leaves.size() == 2 && cgIsIdentText(leaves[0]) && cgIsIdentText(leaves[1]))
        return "vb6_VarConcatNullV(&" + leaves[0] + ", &" + leaves[1] + ", " + a + ", " + b + ")";
    // 链（≥3 叶）或叶子不是变量槽 ⇒ 整枚版逐级左结合：
    //   ((Null & Null) & Null) 与 VB6 的链完全同一个折叠顺序。
    std::string acc = leaves[0];
    for (size_t i = 1; i < leaves.size(); ++i)
        acc = "vb6_VarConcatNullVV(" + acc + ", " + leaves[i] + ")";
    return acc;
}

// C30-11 第六十七刀: `Or`/`And` 的 Null 真值表（变体上下文那两站）。
// 读数 oracle_logic.vbs：Or 只有另一侧是 True 才确定成 True、And 只有另一侧是 False 才确定成
// False，其余全是 Null；`Null Or Null`/`Null And Null` 也是 Null。
// 只认发码器自己产的那一形（实测）：
//   vb6_VariantFromValue(( vb6_VariantToLong(X) | (int32_t)(Y) ))
//   vb6_VariantFromValue(( vb6_VariantToLong(X) & (int32_t)(Y) ))
//   vb6_VariantFromValue(( (int32_t)(Y) | vb6_VariantToLong(X) ))      ← 反向
//   vb6_VariantFromValue(( vb6_VariantToLong(X) | vb6_VariantToLong(Y) )) ← 两槽
// 任一侧不是「槽」也不是「(int32_t)(…) 字面档」就不动（宁可不变）。
// 惰性求值：只传槽**地址**，VariantToLong 在 RTL 里算 —— 当实参先算会先报 94。
std::string CCodeGen::cgNullGuardLogicV(const ASTNode* node, const std::string& cExpr) const {
    if (!node || node->kind != ASTNodeKind::BinaryExpr) return std::string();
    const auto& bin = static_cast<const BinaryExpr&>(*node);
    // 第六十八刀：op 码表 —— 0=Or 1=And 2=Xor 3=Eqv 4=Imp（RTL 侧同一枚出口）
    int opCode = -1;
    switch (bin.op) {
    case BinaryOp::Or:  opCode = 0; break;
    case BinaryOp::And: opCode = 1; break;
    case BinaryOp::Xor: opCode = 2; break;
    case BinaryOp::Eqv: opCode = 3; break;
    case BinaryOp::Imp: opCode = 4; break;
    default: return std::string();
    }
    if (cExpr.find("vb6_VarLogicNullV(") != std::string::npos) return std::string();

    size_t s = 0;
    // 只跳空格：不能跳 ( —— 裸形的首字符就是配对括号，跳掉它前缀判据就全部落空
    // （第六十七刀第一版就是这么静默不命中的）。
    while (s < cExpr.size() && cExpr[s] == ' ') s++;
    static const std::string kFrom = "vb6_VariantFromValue(";
    std::string inner;
    if (cExpr.compare(s, kFrom.size(), kFrom) == 0 && cExpr.back() == ')') {
        inner = cExpr.substr(s + kFrom.size(), cExpr.size() - s - kFrom.size() - 1);
    } else {
        // 实参站/存储站拿到的可能是**裸**的 `(<左> | <右>)`（FromValue 还没包）——
        // 与第六十五/六十六刀同一个坑：两条形都得认，否则永远不命中。
        inner = cExpr.substr(s);
    }
    // 剥掉外层那对全包括号
    if (inner.size() < 3 || inner.front() != '(' || inner.back() != ')')
        return std::string();
    inner = inner.substr(1, inner.size() - 2);

    auto trim2 = [](std::string x) {
        size_t p = x.find_first_not_of(' ');
        size_t q = x.find_last_not_of(' ');
        return (p == std::string::npos) ? std::string() : x.substr(p, q - p + 1);
    };

    // 深度 0 上找运算符（引号内的不算）
    auto findTop = [](const std::string& s, const std::string& ops) -> size_t {
        int depth = 0;
        for (size_t i = 0; i < s.size(); ++i) {
            char ch = s[i];
            if (ch == '(') ++depth;
            else if (ch == ')') --depth;
            else if (ch == '"') { for (++i; i < s.size() && s[i] != '"'; ++i) {} }
            else if (depth == 0 && ops.find(ch) != std::string::npos) return i;
        }
        return std::string::npos;
    };

    std::string left, right;
    int effOp = -1;
    if (inner.compare(0, 2, "~(") == 0 && inner.back() == ')') {
        // Eqv 的发码形：~(A ^ B)（外层那对全包括号已剥）
        const std::string sub = inner.substr(2, inner.size() - 3);
        const size_t p = findTop(sub, "^");
        if (p == std::string::npos) return std::string();
        left = trim2(sub.substr(0, p));
        right = trim2(sub.substr(p + 1));
        effOp = 3;
    } else {
        const size_t p = findTop(inner, "|&^");
        if (p == std::string::npos) return std::string();
        const std::string rhs = trim2(inner.substr(p + 1));
        if (inner[p] == '|') {
            const std::string l = trim2(inner.substr(0, p));
            // Imp 的发码形：(~A) | B
            if (l.size() > 4 && l.compare(0, 2, "(~") == 0 && l.back() == ')') {
                left = trim2(l.substr(2, l.size() - 3));
                right = rhs;
                effOp = 4;
            } else {
                left = l;
                right = rhs;
                effOp = 0;
            }
        } else {
            left = trim2(inner.substr(0, p));
            right = rhs;
            effOp = (inner[p] == '&') ? 1 : 2;
        }
    }
    // 文本形与 AST 运算符必须对上，否则说明歧义（宁可不命中也不发错）
    if (effOp != opCode) return std::string();

    static const std::string kToLong = "vb6_VariantToLong(";
    auto slotOfSide = [&](const std::string& x) -> std::string {
        if (x.compare(0, kToLong.size(), kToLong) != 0 || x.back() != ')')
            return std::string();
        std::string id = x.substr(kToLong.size(), x.size() - kToLong.size() - 1);
        if (id.empty()) return std::string();
        for (char ch : id) {
            bool ok = (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z')
                || (ch >= '0' && ch <= '9') || ch == '_';
            if (!ok) return std::string();
        }
        return id;
    };
    static const std::string kLit = "(int32_t)(";
    auto litOfSide = [&](const std::string& x) -> std::string {
        if (x.compare(0, kLit.size(), kLit) != 0 || x.back() != ')')
            return std::string();
        return x.substr(kLit.size(), x.size() - kLit.size() - 1);
    };

    const std::string idL = slotOfSide(left);
    const std::string idR = slotOfSide(right);
    const std::string litL = idL.empty() ? litOfSide(left) : std::string();
    const std::string litR = idR.empty() ? litOfSide(right) : std::string();
    if (idL.empty() && litL.empty()) return std::string();
    if (idR.empty() && litR.empty()) return std::string();
    const std::string slotA = idL.empty() ? "(vb6_VARIANT*)0" : ("&" + idL);
    const std::string slotB = idR.empty() ? "(vb6_VARIANT*)0" : ("&" + idR);
    const std::string lit = !litL.empty() ? litL : (!litR.empty() ? litR : "0");
    const int mask = (idL.empty() ? 0 : 1) | (idR.empty() ? 0 : 2);
    if (mask == 0) return std::string();
    // 字面量那一侧的静态类型：VB6 里 `Null Or True` = Boolean True、`Null And False`
    // = Boolean False，而 C 的 `(int32_t)((-1))` 分不出 Boolean 与整数 ⇒ 在发码时问 AST。
    int litIsBool = 0;
    {
        const auto& be = static_cast<const BinaryExpr&>(*node);
        const Expr* litSide = !litL.empty() ? be.left.get() : (!litR.empty() ? be.right.get() : nullptr);
        if (litSide) {
            const Vb6Type lt = inferExprType(const_cast<Expr&>(*litSide));
            // 0 = Long（默认）, 1 = Boolean（Null Or True ⇒ Boolean True）,
            // 2 = Integer（Null Or 5 ⇒ **Integer** 5，VB6 IDE 实测的类型）。
            // 先看字面量自己的档：inferExprType 对无后缀整数答 Long，
            // 而 VB6 的 `5` 是 Integer ⇒ 以 LiteralKind 为准。
            if (litSide->kind == ASTNodeKind::LiteralExpr) {
                switch (static_cast<const LiteralExpr&>(*litSide).literalKind) {
                case LiteralKind::Boolean: litIsBool = 1; break;
                case LiteralKind::Integer: litIsBool = 2; break;
                default: break;
                }
            } else if (litSide->kind == ASTNodeKind::UnaryExpr) {
                // `-1` 这种：AST 是一元负 + Integer 字面量，VB6 里它的档仍是 Integer
                // （Null Or -1 = Integer -1，VB6 IDE 实测 D5）。
                const auto& ue = static_cast<const UnaryExpr&>(*litSide);
                if (ue.op == UnaryOp::Negate && ue.operand
                    && ue.operand->kind == ASTNodeKind::LiteralExpr) {
                    const auto lk = static_cast<const LiteralExpr&>(*ue.operand).literalKind;
                    if (lk == LiteralKind::Boolean) litIsBool = 1;
                    else if (lk == LiteralKind::Integer) litIsBool = 2;
                }
            } else if (lt == Vb6Type::Boolean) {
                litIsBool = 1;
            } else if (lt == Vb6Type::Integer) {
                litIsBool = 2;
            }
        }
    }
    return "vb6_VarLogicNullV(" + std::to_string(opCode) + ", " + slotA + ", "
         + slotB + ", " + lit + ", " + std::to_string(mask) + ", "
         + std::to_string(litIsBool) + ")";
}

std::string CCodeGen::wrapVariantValue(ASTNode* valueNode, const std::string& cExpr) const {
    if (!valueNode) return "vb6_VariantEmpty()";
    
    // 特殊情况: Null字面量
    if (valueNode->kind == ASTNodeKind::LiteralExpr) {
        auto& lit = static_cast<LiteralExpr&>(*valueNode);
        if (lit.literalKind == LiteralKind::Null) return "vb6_VariantNull()";
        if (lit.literalKind == LiteralKind::Empty) return "vb6_VariantEmpty()";
        // Boolean 字面量**不再在这里单独成档**(2026-09-30): 原先这一支写的是
        // `vb6_VariantBool((int16_t)(…))`, 而末尾的 boxToVariant 的 Boolean 档写的是同一形
        // —— 两份表又回来了(账 #123/Fix 198 那条"同一类型两条路两种结果"的老坑)。
        // 现在字面量也落到 inferExprType → boxToVariant 这一处权威: 后者对 Boolean
        // 发的是同一个 `(int16_t)` 收窄形, 所以发码逐字节不变, 但**档位只有一份**。
        // (字面量落在 inferExprType 的白名单里, 见下面那个 kind 判断。)
    }

    // C30-11 第六十五刀: Abs/Int/Fix/Hex/Oct 的 Null 传播（必须在下面的
    // inferExprType switch 之前问，否则被 case Double/String 提前收掉）。
    {
        std::string mathV64 = mathNullFamilyV(valueNode, cExpr);
        if (!mathV64.empty()) return mathV64;
    }

    // C30-11 第六十七刀: Or/And 的 Null 真值表（同上，要在 switch 之前）。
    {
        std::string lgV67 = cgNullGuardLogicV(valueNode, cExpr);
        if (!lgV67.empty()) return lgV67;
    }

    // C30-11 第六十六刀: `&` 的双 Null 折叠（同上，要在 switch 之前；
    // 那张表对这一形答 String，答了就成 vb6_VariantString 再也没机会折 Null）。
    {
        std::string catV66 = cgNullGuardConcatV(valueNode, cExpr);
        if (!catV66.empty()) return catV66;
    }

    // C30-11 第六十五刀: 比较符的 Null 传播（同上，必须在 switch 之前；
    // 那张表对比较答 Boolean，答了就变回 Double 包）。
    {
        std::string cmpV65 = cgNullGuardCmpV(valueNode, cExpr);
        if (!cmpV65.empty()) return cmpV65;
    }

    // Fix 090d2: C 表达式顶层已是 vb6_VARIANT 时直接返回, 不再包装 —
    // 原 2998 行的检查在 inferExprType switch 之后, 对符号表/推断层
    // 认为是具体标量但实际生成 vb6_VARIANT 的表达式 (如 COM 属性
    // fileInfo.Size 被推断为 Long, 实际生成
    // vb6_VariantFromComResult(vb6_ComCall(...))) 会被 switch 抢先包装
    // 成 vb6_VariantLong(vb6_VARIANT) → C2440 (cHttpServerResponse File
    // 内 fileSize = fileInfo.Size, fileSize As Variant).
    if (isVariantSlot(cExpr, nullptr)) {  // P1.3: ast 不可得 (cExpr 为形参)
        return cExpr;
    }

    // Fix 170: 右侧是 VB6 整体数组引用 `A()` (空括号) → 装箱成持有数组的 Variant。
    // 必须先于下面的 inferExprType switch: 它对 `A()` 可能给回**元素**类型 (Byte/Long),
    // 于是 vb6_VariantLong(载体指针) → C 侧只是警告, 指针被截断成 int32, 运行期才炸。
    if (valueNode->kind == ASTNodeKind::IndexOrCallExpr
        && isWholeArrayRef(static_cast<const Expr*>(valueNode))) {
        // 账 §11.180 (第一百九十七刀): 落带元素档的装箱口 —— 未分配/已 Erase 的数组
        // 装进 Variant 时真机仍答声明的元素档 (P233: `Long()` 是 8195，不是 8204)。
        // 元素档问**声明那张表** arrayElemTypes_（声明站与 ReDim 站共用的唯一一份）；
        // 查不到再回落 inferExprType 掩掉 Array 位那一档，认不出的档回落 vb6_sa_variant
        // = 8204，即改前的答案。
        Vb6Type e197w = Vb6Type::Variant;
        if (valueNode->kind == ASTNodeKind::IndexOrCallExpr) {
            const auto& ic197 = static_cast<const IndexOrCallExpr&>(*valueNode);
            if (ic197.callee && ic197.callee->kind == ASTNodeKind::IdentifierExpr) {
                const std::string nm197 = Symbol::toLower(
                    static_cast<const IdentifierExpr&>(*ic197.callee).name);
                auto it197 = arrayElemTypes_.find(nm197);
                if (it197 != arrayElemTypes_.end()) e197w = it197->second;
            }
        }
        if (e197w == Vb6Type::Variant) {
            e197w = static_cast<Vb6Type>(static_cast<uint16_t>(inferExprType(static_cast<Expr&>(*valueNode)))
                                         & ~static_cast<uint16_t>(Vb6Type::Array));
        }
        return "vb6_VariantArrayTyped((void*)" + cExpr + ", " + mapSaElemType(e197w) + ")";
    }
    
    // 账 #305 §B163 (第五十一刀): **裸数组名当右值**（`Dim v As Variant: v = mL`，无括号那一形）。
    // 这一问必须排在下面那张 inferExprType 表之前：它对同一个名字答的是**元素**档
    // （模块级数组登记在 knownLongVars_/knownBstrVars_ 里），于是发出来的是
    // `vb6_VariantLong(mL)` —— 描述符指针截进
    // int32、cl 一声不响（实测两台 3/"Long"，VB6 该 8195/"Long()"）。
    // 第二个成因是**名字撞上 VB 内在函数**：同一份语句把数组从 `loc` 改名 `loc2` 就正常装箱，
    // 叫 `loc` 就折 —— `Loc` 是返回 Long 的内在函数，inferExprType 先答了它
    // （P51E→P51H 的对照，与 §B217/#219 那族同源）。两条成因共用这一问就一起治。
    // 判据与第四十九/五十刀同一条权威 isArrayCarrierArgName（读的是 knownArrays_ 那张**名字表**，
    // 所以对撞名免疫），装箱仍走末尾 boxToVariant 已经在用的 vb6_VariantFromValue
    // （_Generic 认 vb6_SafeArray1D* 那一档，标签问描述符自己 = 第四十六刀）。
    // ⚠ 与 COM 那站同一道秩护栏：`vb6_VariantFromValue` 的 _Generic 没有 vb6_SafeArrayND* 那一档，
    //   多维载体落 default = vb6_VariantObject ⇒ 对一枚描述符发 Release = 野 Release 堆损坏
    //   （§B85 那一族；ND 归 §B151，补法是给 _Generic 加 ND 档，不是在这里绕）。
    if (valueNode->kind == ASTNodeKind::IdentifierExpr) {
        auto& idB163 = static_cast<IdentifierExpr&>(*valueNode);
        std::string lowerB163 = Symbol::toLower(idB163.name);
        if (isArrayCarrierArgName(lowerB163, cExpr)) {
            auto itB163 = arrayDimCounts_.find(lowerB163);
            if (itB163 != arrayDimCounts_.end() && itB163->second <= 1) {
                // 账 §11.180 (第一百九十七刀): 同上 —— 裸数组名装箱带元素档。
                Vb6Type e197n = Vb6Type::Variant;
                auto it197 = arrayElemTypes_.find(lowerB163);
                if (it197 != arrayElemTypes_.end()) e197n = it197->second;
                if (e197n == Vb6Type::Variant) {
                    e197n = static_cast<Vb6Type>(static_cast<uint16_t>(inferExprType(static_cast<Expr&>(*valueNode)))
                                                 & ~static_cast<uint16_t>(Vb6Type::Array));
                }
                return "vb6_VariantArrayTyped((void*)" + cExpr + ", " + mapSaElemType(e197n) + ")";
            }
        }
    }

    // 账 #305 新账 #407 / §B190 第一百七十一刀: **视图值进装箱站**（`TypeName(mv)` / 赋给
    // Variant 槽那一族）。这一问排在下面那张 inferExprType 表之前 —— 那张表对视图名答的是
    // **接口模块名**，而 C 侧那一档是值结构体，落 _Generic 的 default = vb6_VariantObject(整枚
    // 结构体) ⇒ cl C2172「实参不是指针」(实物 = 门禁在册的 tests/b189_gate，B189Drv.c(11)，
    // 改前两台同红、第一百六十六刀收线时数到的)。改的是一条从前**编不过**的路 ⇒ 存量产物
    // 按定义一个字不许动。
    {
        std::string viewBox171 = legacyViewBoxExpr(valueNode, cExpr);
        if (!viewBox171.empty()) return viewBox171;
    }

    // 账 #615 (第353刀): **UDT 成员的整数组读** (`w = h.Items`) 的装箱档 ——
    // 以前整枚落进下面那张 inferExprType 表, 而它对「成员是动态数组」答的是**元素**档 (Long)
    // ⇒ 发 `vb6_VariantLong(h.Items)` = 把描述符指针截进 int32 (cl 一声不响, 运行期才炸)
    // ⇒ VarType 答 3 而不是 8195 (VB6 那一形是数组)。
    // 这一问必须排在那张表**之前** (同 §B163 裸数组名那一问的次序理由: 那张表答的是元素档)。
    // ⚠ 只认**一维**: `vb6_VariantFromValue` 的 _Generic 没有 vb6_SafeArrayND* 那一档,
    // 多维载体落 default = vb6_VariantObject ⇒ 对一枚描述符发 Release = 野 Release 堆损坏
    // (台账 §B85 那一族) ⇒ 多维不接, 维持今天的形状。
    if (valueNode->kind == ASTNodeKind::MemberAccessExpr
        || valueNode->kind == ASTNodeKind::WithMemberExpr) {
        std::string ownerUdt615;
        std::string memberName615;
        if (valueNode->kind == ASTNodeKind::MemberAccessExpr) {
            auto& ma615 = static_cast<MemberAccessExpr&>(*valueNode);
            ownerUdt615 = inferUdtTypeOfExpr(*ma615.object);
            memberName615 = ma615.memberName;
        } else {
            auto& wm615 = static_cast<WithMemberExpr&>(*valueNode);
            memberName615 = wm615.memberName;
            if (!withObjectVars_.empty()) {
                auto itU615 = knownUdtVars_.find(Symbol::toLower(withObjectVars_.back()));
                if (itU615 != knownUdtVars_.end()) ownerUdt615 = itU615->second;
            }
        }
        if (!ownerUdt615.empty() && !memberName615.empty()) {
            MemberArrayDecl mad615;
            if (resolveMemberArrayDecl(ownerUdt615, std::string(), memberName615, mad615)
                && mad615.ok && mad615.arrayRank <= 1) {
                return "vb6_VariantFromValue(" + cExpr + ")";
            }
        }
    }

    // 使用inferExprType推断表达式类型
    if (valueNode->kind == ASTNodeKind::BinaryExpr ||
        valueNode->kind == ASTNodeKind::UnaryExpr ||
        valueNode->kind == ASTNodeKind::LiteralExpr ||
        valueNode->kind == ASTNodeKind::IdentifierExpr ||
        valueNode->kind == ASTNodeKind::IndexOrCallExpr ||
        valueNode->kind == ASTNodeKind::MemberAccessExpr) {
        Vb6Type vtype = inferExprType(static_cast<Expr&>(*valueNode));
        // C30-10 第四十一刀: 存储站与实参站对齐（两站共用同一柄尺 isNarrowIntExpr）。
        // 证据（用户真 VB6 日志 C20/C21/C22/H04/H05）：`t = Int(5)` / `Fix(5)` /
        // `Abs(-5)` / `Abs(i)` 在 VB6 里全是 Integer；而这条路以前按 inferExprType 的
        // Double 直接装 VT_R8 —— 第三十八刀只落到了实参站（本函数的 case Double 不查
        // 窄化）。判据提到 switch 之前后：case Long / case Integer 两档发出的字符串与
        // 末尾 boxToVariant 逐字相同（发码不变），只有 Int/Fix/Abs 这族由 VT_R8 变 VT_I2。
        // ⚠ `Abs(v)`（Variant 实参）依旧不在内：静态不可知，要 RTL 按运行时子类型答，
        //   那是另一刀（台账 §11.20 记着）。
        // 账 #305 §11.161 (第一百七十八刀): 裸 Currency 变量先问身份表，再走下面的
        // 类型表 —— 那张表对 Currency 只能答 Double（C 侧同为 double）。
        {
            std::string curBox178 = currencyVariantBox(static_cast<Expr*>(valueNode), cExpr);
            if (!curBox178.empty()) return curBox178;
        }
        if (isNarrowIntExpr(static_cast<const Expr*>(valueNode)))
            return "vb6_VariantInt((int16_t)(" + cExpr + "))";
        // <vbeclipse>: 这张表只留"与 _Generic 结果一致"的四档 (String/Long/Integer/Double)。
        // Boolean/Byte/Single/Date/Currency 交给末尾的 boxToVariant —— 以前两份表各写一份
        // 档位 (账 #123 改了这份的 Byte、Fix 198 改了那份的 Boolean), 于是同一类型的装箱
        // 结果取决于表达式走哪条路, DT43 的布尔就是这么漏的。
        switch (vtype) {
            case Vb6Type::String:    return "vb6_VariantString(" + cExpr + ")";
            case Vb6Type::Long:
                // C30-10 第三十二刀: 这档里有两形 VB6 答的是 **Integer** (字面量按值取型;
                // `As Integer` 变量符号表记成 Long), 不在这一站抢装 VT_I4 —— 交给末尾那个
                // 唯一权威 boxToVariant (判据 = isNarrowIntExpr, 两位共用一处)。
                if (isNarrowIntExpr(static_cast<const Expr*>(valueNode))) break;
                // §11.207 (第二百二十四刀): **整枚数组**不在此列 —— `VarType(d)` 里 d 是
                // 载体描述符, 该答 vtArray|vt<元素> (8195/8194/8197…), 强装成 VT_I4/I2 就
                // 只剩 3/2 (实测 tests/k483_redim_member_elem 的 K483-long-vt / K483-int-vt)。
                if (isWholeArrayRef(static_cast<const Expr*>(valueNode))
                    || isSafeArrayCarrier(cExpr)) break;
                // §11.207: 显式收窄 —— `vb6_Abs(l)` 这类 C 侧返回 double 的调用，不 cast
                // 就会让 _Generic 猜成 VT_R8（真机 AL/IL/FL/RL 全要求 3|Long）。
                return "vb6_VariantLong((int32_t)(" + cExpr + "))";
            // `Dim i As Integer` / CInt(...) 这一族以前也装成 VT_I4 (同一行 d121 的
            // int=3)。VB6 对 Integer 就是 VT_I2 —— 显式收窄, 不靠 _Generic 猜 short。
            // 整枚数组同一句排除（见上）。
            case Vb6Type::Integer:
                if (isWholeArrayRef(static_cast<const Expr*>(valueNode))
                    || isSafeArrayCarrier(cExpr)) break;
                return "vb6_VariantInt((int16_t)(" + cExpr + "))";
            case Vb6Type::Double:    return "vb6_VariantDouble(" + cExpr + ")";
            default: break;
        }
    }
    
    // 如果右侧已经是vb6_VARIANT类型(如函数返回Variant), 直接赋值
    // Fix 051: 排除 vb6_VariantTo* 函数 (如 vb6_VariantToObjectVal 返回 void*,
    // vb6_VariantToString 返回 BSTR), 这些不是 vb6_VARIANT 类型, 需要包装.
    if ((cExpr.find("vb6_Variant") == 0 && cExpr.find("vb6_VariantTo") != 0)
        || cExpr.find("vb6_CStr") == 0) {
        return cExpr;
    }
    
    // COM后期绑定调用: vb6_ComCall/vb6_ComCallByDispid返回VARIANT*, 需转为vb6_VARIANT
    if (cExpr.find("vb6_ComCall(") == 0 || cExpr.find("vb6_ComCallByDispid(") == 0) {
        return "vb6_VariantFromComResult(" + cExpr + ")";
    }
    
    // Array()临时变量: _arr_N is vb6_SafeArray1D*, 包装为Variant持有数组
    if (cExpr.find("_arr_") == 0) {
        return "vb6_VariantArray(" + cExpr + ")";
    }
    
    // Fix 025: 默认档 —— 现在**只**经由 boxToVariant 这一处权威发出。
    // 为什么: 装箱表以前有两份并行 (本函数的 switch + boxToVariant), 账 #123 只在
    // 这一份修了 Byte、Fix 198 只在那一份修了 Boolean, 于是"某一档对不对"取决于
    // 表达式走了哪条路 (实测 dt1.CheckBox 走实参装箱那条 → 读回 Long/3)。
    // 现在两个入口共用一处收尾, 新增档位只需要改 boxToVariant。
    return boxToVariant(static_cast<Expr*>(valueNode), cExpr);
}

// <vbeclipse>: "这个 C 表达式已经是 vb6_VARIANT 了吗" —— 结构化判定, 不再靠 ctor 名字清单。
// 名字清单是这里的第三个坑: vb6_VariantByte( 不在 cExprIsVariant 的清单里, 于是
// `VarType(VarType(by))` 第二层把已经是 VARIANT 的表达式按 Byte 档又装一遍 →
// vb6_VariantByte((uint8_t)(vb6_VariantByte(...))) → C2440 + C2198 (实测 test_bool_display B25)。
// 规则: 顶层是 vb6_Variant*/vb6_VariantFrom* 即已是 VARIANT; vb6_VariantToXxx 是**提取**,
// 返回具体类型, 不算。
static bool cExprIsVariantCarrierStr(const std::string& cExpr) {
    size_t s = cExpr.find_first_not_of(" \t\r\n(*&");
    if (s == std::string::npos) return false;
    if (cExpr.compare(s, 11, "vb6_VariantTo") == 0) return false;
    return cExpr.compare(s, 11, "vb6_Variant") == 0 ||
           cExpr.compare(s, 12, "vb6_VARIANT{") == 0;
}

bool CCodeGen::cExprIsVariantCarrier(const std::string& cExpr) const {
    return isVariantSlot(cExpr, nullptr) || cExprIsVariantCarrierStr(cExpr);  // P1.3: ast 不可得 (cExpr 为形参)
}

// Fix 198 + <vbeclipse>: **装箱的唯一权威**。按 VB 声明类型选档, 只在类型确实与
// C 表示不一致时才改写; 其余一律逐字节退回 vb6_VariantFromValue (_Generic 按 C 类型
// 选 ctor) —— 早退式的"已是 VARIANT 就不包"看着更干净, 但它会把存量码也一起改了
// (护栏实测 VbEclipse 的 VB6_SA_AT 实参少了那层恒等包装)。
// 需要显式档案的五种 (C 表示区分不出来或就是错的):
//   Boolean → VT_BOOL(11)   C 侧是 int16_t/int32_t, _Generic 会装成 VT_I4
//   Byte    → VT_UI1(17)    C 侧 uint8_t 走得到对档, 但模块级 Byte 曾被推断成 Long
//   Single  → VT_R4(4)      _Generic 的 float 档以前也升到 VT_R8 (哨兵实测 5)
//   Date    → VT_DATE(7)    C 侧就是 double, 只有 VB 类型能说话 (哨兵实测 5)
//   LongPtr → 按指针宽度的数值档  句柄 getter 交 void*, 而 _Generic 的 default 一档是
//          vb6_VariantObject —— 类型那一问 (§B136 第二十八刀) 答了句柄, 这一路不吃它
// 不显式改写的 (Integer/Long/Double/String/Object) 在 _Generic 下与 VB6 同档, 保持
// 原样以免动到存量发码。
// 账 §11.181 (第一百九十八刀): Date 数组的**元素**装进 Variant 要答 VT_DATE (VarType 7 /
// TypeName "Date"), 而元素在 C 侧就是 double (VB6_SA_AT(double, q, i)) —— _Generic 只能答
// VT_R8。这一问按声明那张表 arrayElemTypes_ (与第一百九十七刀同一份) 认出「Date 数组的
// 一枚元素」, 再把标签补上。真机 P234: C3 elem vt=7 tn=Date。
// §11.184 (第二百零一刀): 上一刀只认一维单下标, 多维那形留在原档 —— 真机
// `.temp/k201p/p243.txt` 的对照读数 (T1 标量 / T2 一维动态 / T3 一维静态 全答 7/Date,
// T4 二维静态答 5/Double) 说明差额只在**秩**上, 元素声明与 C 侧 double 槽都一样。
// 故这里也按「秩 ≥ 1 且声明是 Date 数组」认, 不再按秩分叉。
std::string CCodeGen::dateElemVariantBox(const Expr* expr, const std::string& cExpr) const {
    if (!expr || expr->kind != ASTNodeKind::IndexOrCallExpr) return std::string();
    /* 幂等: 实参站与存储站会各问一次装箱 (wrapVariantValue 末尾就交回 boxToVariant),
       第二问拿到的是已经包好的文本 —— 再包一层就成 vb6_VariantDate(vb6_VARIANT) (C2440)。 */
    if (cExpr.rfind("vb6_VariantDate(", 0) == 0) return std::string();
    const auto& ic198 = static_cast<const IndexOrCallExpr&>(*expr);
    if (!ic198.callee || ic198.callee->kind != ASTNodeKind::IdentifierExpr) return std::string();
    if (ic198.positional.empty()) return std::string();
    const std::string nm198 = Symbol::toLower(
        static_cast<const IdentifierExpr&>(*ic198.callee).name);
    auto it198 = arrayElemTypes_.find(nm198);
    if (it198 == arrayElemTypes_.end() || it198->second != Vb6Type::Date) return std::string();
    return "vb6_VariantDate((double)(" + cExpr + "))";
}

std::string CCodeGen::boxToVariant(Expr* expr, const std::string& cExpr) const {
    // C30-11 第六十五刀: 实参站的比较符 Null 传播（与存储站同一个判据）。
    {
        std::string cmpV65 = cgNullGuardCmpV(expr, cExpr);
        if (!cmpV65.empty()) return cmpV65;
    }
    // C30-11 第六十七刀: 实参站的 Or/And Null 真值表。
    {
        std::string lgV67 = cgNullGuardLogicV(expr, cExpr);
        if (!lgV67.empty()) return lgV67;
    }
    // C30-11 第六十六刀: 实参站的 `&` 双 Null 折叠（拿到的是裸 vb6_BSTR_ConcatFree(...)）。
    {
        std::string catV66 = cgNullGuardConcatV(expr, cExpr);
        if (!catV66.empty()) return catV66;
    }
    // Fix <VBFlexGridDemo>: 实参是裸 COM 调用结果 (vb6_ComCall / vb6_ComGetProp /
    // vb6_ComGetObjectProp 返回 void*, 承载一个 COM VARIANT*) 时, 必须先用
    // vb6_VariantFromComResult 解引用成 vb6_VARIANT (深拷贝并释放原 VARIANT*),
    // 绝不能用 vb6_VariantFromValue —— 后者对 void* 走 _Generic 的
    // vb6_VariantObject 档, 把 COM VARIANT* 当成 IDispatch 对象去 AddRef, 当
    // VARIANT 实际是 BSTR/数值时 *(void**)variantPtr == 0x8 触发
    // IsBadReadPtr(0x8) → 0xC0000005 "内存不能为 read"
    // (VBFlexGridDemo ReadProperties: PropClipSeparators =
    //  VarToStr(PropBag.ReadProperty("ClipSeparators", "")) 启动即崩).
    // 与 conv_cstr (Fix 121) / arg_variant (Fix 113) 既有口径一致.
    {
        // 精确匹配裸 vb6_ComCall( / vb6_ComCallByDispid( / vb6_ComGetProp(
        // (12/20/15 字符含左括号); 不误伤 vb6_ComCallInt/BSTR/Double/Object。
        // 账 (M3 身份刀): vb6_ComGetObjectProp 交回的是 **IDispatch* 本体** (vb6com_wrap.c:129),
        // 不是堆上 VARIANT*; 把它喂给 vb6_VariantFromComResult 是类型混淆 (读 vtable 首字当
        // vt、再把代码地址 free 掉)。实测 tests/ctlident 的 TypeName(ActiveForm) 即崩在此。
        // 对象指针交给 vb6_VariantFromValue 的 _Generic default 档 = vb6_VariantObject
        // (VT_DISPATCH, 且 vb6_ComIsDispatchable 对 HWND 判否, 安全)。
        size_t cs = 0;
        while (cs < cExpr.size()
               && (cExpr[cs] == '(' || cExpr[cs] == ' ' || cExpr[cs] == '\t'
                   || cExpr[cs] == '\n' || cExpr[cs] == '\r')) cs++;
        bool comRes = (cExpr.compare(cs, 12, "vb6_ComCall(") == 0)
                   || (cExpr.compare(cs, 20, "vb6_ComCallByDispid(") == 0)
                   || (cExpr.compare(cs, 15, "vb6_ComGetProp(") == 0);
        if (comRes) return "vb6_VariantFromComResult(" + cExpr + ")";
    }
    // C30-11 后记: 前期绑定 COM 方法的 Boolean 返回 (vb6_ComCallBool) —— 同一发调用、
    // 只是名字分档。VB6 的 `v = fso.FileExists(p)` 里 v 是 Variant(Boolean), 默认档会
    // 装成 VT_I4(-1)/VT_I2; 这里让它进 Boolean 那一档 (VB6_CStr / VarType 才看得对)。
    {
        size_t bs = 0;
        while (bs < cExpr.size()
               && (cExpr[bs] == '(' || cExpr[bs] == ' ' || cExpr[bs] == '\t'
                   || cExpr[bs] == '\n' || cExpr[bs] == '\r')) bs++;
        if (cExpr.compare(bs, 17, "vb6_ComCallBool(") == 0)
            return "vb6_VariantBool((int16_t)(" + cExpr + "))";
    }
    // 账 #305 新账 #407 第一百七十一刀: 与上面 wrapVariantValue 同一问、同一出口 (视图值进装箱站)。
    // 这一站也必须在 inferExprType 之前问: 它对视图名给的是接口模块名，不是 C 侧那一档。
    {
        std::string viewBox171b = legacyViewBoxExpr(expr, cExpr);
        if (!viewBox171b.empty()) return viewBox171b;
    }
    // 账 #305 §11.161 (第一百七十八刀): 实参站的同一问、同一出口（裸 Currency 变量
    // 交给 ByVal Variant 形参时 VB6 给 VT_CY —— 探针 P12/P14）。
    {
        std::string curBox178 = currencyVariantBox(expr, cExpr);
        if (!curBox178.empty()) return curBox178;
    }
    // 账 §11.181 (第一百九十八刀): 这里多问一问「这是不是 Date 数组的一枚元素」。
    // 为什么必须在装箱站问: 元素在 C 侧就是 double (VB6_SA_AT(double, q, i)), 标签丢了;
    // 真机 P234: `v = q(1)` / `TypeName(q(0))` / For Each 全部答 VT_DATE(7) / "Date"。
    {
        std::string db198 = dateElemVariantBox(expr, cExpr);
        if (!db198.empty()) return db198;
    }
    if (!expr) return "vb6_VariantFromValue(" + cExpr + ")";
    // 已经是 VARIANT 的表达式交给 _Generic 的 vb6_VariantIdentity 档恒等直传,
    // 绝不再按 VB 类型强装 (否则就是上面那条双装)。
    if (cExprIsVariantCarrier(cExpr)) return "vb6_VariantFromValue(" + cExpr + ")";
    // C30-10 第五十二刀: `Abs(Variant)` 的**变体位**（装箱站）走运行时子类型阶梅（C23）。
    // 真 VB6 GProbe AB 段: Abs(Byte)=Byte(17) / Abs(Integer)=Integer / Abs(CInt(-32768))=Long /
    // Abs(Long)=Long / Abs(Single)=Single / Abs(Double)=Double / Abs(Currency)=Currency /
    // Abs(Boolean)=Integer / Abs(Empty)=Integer 0 / Abs(Null)=Null / Abs(Date)=Date /
    // Abs("5")=Double / Abs("x")=错误 13。旧站只有 double 出口 ⇒ 13 档全堆成 VT_R8。
    //
    // **只在这一站改名，绝不在调用站改名**：`Abs` 的结果有两副消费面 ——
    // 变体位（`VarType/TypeName/赋 Variant/CStr`）过装箱站，数值位
    // （`Abs(v) < 0.001`、`Abs(v) * 2`、比较、赋 Double）**不过**这一站。
    // 上一版把改名放进调用站 fixup，实测 pclinedraw 当场红：
    // `vb6_AbsV(...) < 0.001` ⇒ cl C2440（float→vb6_VARIANT）+ C2088（struct 不能比）。
    if (expr && expr->kind == ASTNodeKind::IndexOrCallExpr) {
        const auto& call64 = static_cast<const IndexOrCallExpr&>(*expr);
        if (call64.positional.size() == 1 && call64.callee
            && call64.callee->kind == ASTNodeKind::IdentifierExpr) {
            const std::string fn64 = Symbol::toLower(
                static_cast<const IdentifierExpr&>(*call64.callee).name);
            // C30-11 第六十四刀: 这一族真值都是“Null 进、Null 出”（读数 oracle_num.vbs：
            // Abs/Int/Fix/Hex/Oct(Null) = Null）。它们以前全落在下方装箱表里，Null 到那里
            // 已被折成 0（Hex/Oct 还要先过第六十三刀新加的 vb6_VariantToLong 严格门 ⇒ 94）。
            // 处置：整只槽交给 RTL 的 *V 出口，先判 vtNull 再算。
            //
            // ⚠ 判据是**文本前缀**而不是 inferExprType：`vb6_Abs(vb6_VariantToDouble(` 这串
            //   本身就证明上游那一站真把 Variant 槽折进来了。旧版（第五十二刀的 Abs 那一分支）
            //   还要再问一句 inferExprType(实参)==Variant，而这个推断对 `Dim nv As Variant`
            //   的局部答不出 Variant ⇒ 整条路根本没走到（本刀实测：改前发码仍是
            //   vb6_VariantDouble(vb6_Abs(vb6_VariantToDouble(nv)))，Abs(Null) 给 0）。
            static const struct { const char* vb; const char* pre; const char* out; } k64[] = {
                {"abs", "vb6_Abs(vb6_VariantToDouble(", "vb6_AbsV("},
                {"int", "vb6_Int(vb6_VariantToDouble(", "vb6_IntV("},
                {"fix", "vb6_Fix(vb6_VariantToDouble(", "vb6_FixV("},
                {"hex", "vb6_Hex(vb6_VariantToLong(",   "vb6_HexV("},
                {"oct", "vb6_Oct(vb6_VariantToLong(",   "vb6_OctV("},
            };
            for (const auto& m64 : k64) {
                if (fn64 != m64.vb) continue;
                const std::string pre64 = m64.pre;
                size_t s64 = 0;
                while (s64 < cExpr.size() && (cExpr[s64] == '(' || cExpr[s64] == ' ')) s64++;
                if (cExpr.compare(s64, pre64.size(), pre64) == 0
                    && cExpr.size() >= s64 + pre64.size() + 2 && cExpr.back() == ')') {
                    // 尾部两层右括号（一层提取、一层调用）一并剥掉。
                    std::string inner64 = cExpr.substr(s64 + pre64.size(),
                                                       cExpr.size() - s64 - pre64.size() - 2);
                    return std::string(m64.out) + inner64 + ")";
                }
            }
            // §11.208 (第二百二十五刀): Round 有 0..1 个位数实参，上面那张单参表接不住
            // （它的闸门是 positional.size()==1）。真机 .temp/k224p/P298 的 vI..vN 与
            // P300 的 V1..V7 第三列：变体载体上的 Round 也保子类型（2/3/4/6/17/2/1）。
            // 作法与 §11.207 在 mathNullFamilyV 里那一支完全同形：从
            // `vb6_Round(vb6_VariantToDouble(` 之后按括号配平取第一实参正文，
            // 尾巴空 ⇒ 位数缺省 0，尾巴是 `, n)` ⇒ 照抄。
            if (fn64 == "round" && call64.positional.size() <= 2) {
                static const std::string preR64 = "vb6_Round(vb6_VariantToDouble(";
                size_t sR64 = 0;
                while (sR64 < cExpr.size() && (cExpr[sR64] == '(' || cExpr[sR64] == ' ')) sR64++;
                if (cExpr.compare(sR64, preR64.size(), preR64) == 0) {
                    const size_t pR64 = sR64 + preR64.size();
                    int depthR64 = 1;
                    size_t kR64 = pR64;
                    for (; kR64 < cExpr.size(); ++kR64) {
                        if (cExpr[kR64] == '(') depthR64++;
                        else if (cExpr[kR64] == ')') { if (--depthR64 == 0) break; }
                    }
                    if (kR64 < cExpr.size()) {
                        std::string innerR64 = cExpr.substr(pR64, kR64 - pR64);
                        std::string tailR64 = cExpr.substr(kR64 + 1);
                        if (!tailR64.empty() && tailR64.back() == ')') tailR64.pop_back();
                        if (tailR64.empty()) tailR64 = ", 0";
                        return "vb6_RoundV(" + innerR64 + tailR64 + ")";
                    }
                }
            }
        }
    }
    // C30-10 第三十二刀 格①: 常整数 / `As Integer` 变量按 VB6 的**Integer 档**装箱。
    // inferExprType 对 `1` 与 `As Integer` 变量都答 Long (C 侧 int32 的口径), 落 _Generic
    // 的 `int:` 档 ⇒ VT_I4 (VarType 3 / TypeName "Long"); VB6 答 Integer (夹具
    // tests/dict_native d121 的 int=3, 真值 2)。判据与口径见 isNarrowIntExpr 那段话。
    if (isNarrowIntExpr(expr))
        return "vb6_VariantInt((int16_t)(" + cExpr + "))";
    // 账 §11.182 (第一百九十九刀): **裸数组名进装箱站** —— 谓词实参站。
    // 九枚只读谓词 (VarType/TypeName/IsArray/…) 的实参以前落 default → _Generic 按 **C 指针**
    // 猜: 未分配 (载体 NULL) 的动态数组因此答 8204 / "Variant()" (真机 P238: 8195/"Long()"、
    // 8199/"Date()"、8200/"String()"、8209/"Byte()"、多维同一档); 模块级 Date 数组更糟 ——
    // 下面 Date 那一臂对它发 vb6_VariantDate(md) = 把描述符指针塞进 double 形参 (cl C2440, 编不过)。
    // 判据与 §B163 / §11.180 同一份名字表 (isArrayCarrierArgName) + 同一份声明元素档
    // (arrayElemTypes_); 只多一条: 出口用**借用**那一枚 (§11.182 的 vb6_VariantArrayBorrowTyped)
    // —— 谓词形参是按值 VARIANT、没人 Clear 那个临时量, 深拷副本会漏 (§11.106 同一理由)。
    // ⚠ 与 §B163 的秩护栅不同: 借用不碰所有权, 多维也安全 (真机 P238B 的 N1..N5: ReDim 之前 /
    //   Erase 之后的 `Long()` 多维同样 8195), 所以这里**不**按 arrayDimCounts_ 限秩。
    // ⚠ 卫兵: 已包过的文本 (vb6_ 前缀) 一律不接 —— 本臂接的是**裸**实参文本, 再包一层就成了
    //   (void*)(vb6_VARIANT)。
    if (expr->kind == ASTNodeKind::IdentifierExpr) {
        const std::string lower199 = Symbol::toLower(
            static_cast<const IdentifierExpr&>(*expr).name);
        if (cExpr.rfind("vb6_", 0) != 0 && isArrayCarrierArgName(lower199, cExpr)) {
            Vb6Type e199 = Vb6Type::Variant;
            auto it199 = arrayElemTypes_.find(lower199);
            if (it199 != arrayElemTypes_.end()) e199 = it199->second;
            if (e199 == Vb6Type::Variant) {
                e199 = static_cast<Vb6Type>(
                    static_cast<uint16_t>(inferExprType(*expr))
                    & ~static_cast<uint16_t>(Vb6Type::Array));
            }
            return "vb6_VariantArrayBorrowTyped((void*)(" + cExpr + "), "
                   + mapSaElemType(e199) + ")";
        }
    }
    switch (inferExprType(*expr)) {
    case Vb6Type::Integer:
        // §11.207 (第二百二十四刀): 与存储站 `case Integer` 同一行 —— 这里以前**没有**
        // Integer 臂，于是 `Int(i)` / `Int(CBool(True))` 这类 C 侧返回 double 的调用落
        // default ⇒ VT_R8（真机 II/RI/D1 全要求 2|Integer）。
        // 整枚数组除外：`VarType(d)` 该答 vtArray|vt<元素>（k483 的 8194）。
        if (isWholeArrayRef(expr) || isSafeArrayCarrier(cExpr)) break;
        return "vb6_VariantInt((int16_t)(" + cExpr + "))";
    case Vb6Type::Long:
        // §11.207 (第二百二十四刀): 与存储站同一条 —— `vb6_Abs(l)` / `vb6_Int(l)` 这类
        // C 侧返回 double 的调用，不 cast 就被 _Generic 猜成 VT_R8（真机 AL/IL/FL/RL = 3）。
        // 窄档（Integer 值）已在上面被 isNarrowIntExpr 接走，不会走到这里。
        if (isWholeArrayRef(expr) || isSafeArrayCarrier(cExpr)) break;
        return "vb6_VariantLong((int32_t)(" + cExpr + "))";
    case Vb6Type::Boolean:        // 形参是 int16_t: 显式收窄, 兼容 _Bool/int 两种 C 侧布尔表示.
        return "vb6_VariantBool((int16_t)(" + cExpr + "))";
    case Vb6Type::Byte:
        return "vb6_VariantByte((uint8_t)(" + cExpr + "))";
    case Vb6Type::Single:
        return "vb6_VariantSingle((float)(" + cExpr + "))";
    case Vb6Type::Date:
        return "vb6_VariantDate(" + cExpr + ")";
    // C30-10 第三十二刀 格②: Currency (C3 用 double 承载值, 见 vb6rtl_misc.c 的
    // vb6_CCur) 以前落 _Generic 的 `double:` 档 ⇒ VT_R8 (VarType 5 / "Double"),
    // VB6 是 VT_CY (6 / "Currency")。出口是 vb6_VariantCur (10000 倍定点)。
    case Vb6Type::Currency:
        return "vb6_VariantCur(" + cExpr + ")";
    /* ai/032 §6.9 收口: C3 扩展的无符号/小整型族 (SByte 8 / UInteger 16 / ULong 32 /
     * ULongLong 64 —— `U<x>` 恒等于「<x> 的无符号版」, 见 types.hpp:41-49)。
     *   以前这四档**没有臂**, 全落 default -> vb6_VariantFromValue, 也就是交给
     *   RTL 的 _Generic 按 C 类型猜: uint32/uint64 猜对了(VT_UI4/VT_UI8), 而
     *   SByte(int8_t=char) 命中 `char:` 档、UInteger(uint16_t=unsigned short) 命中
     *   `wchar_t:` 档, 两档都被装成 **VT_I4**。值没丢(8/16 位塞得进), 丢的是**类型
     *   身份** —— TypeName 答 "Long"、VarType 答 3。ULong/VT_UI4 那两档的注释写着
     *   「答 "Long" 才是误导」, 这两档按同一口径。
     *   ⚠ 别去给 _Generic 加 `unsigned short:` 分支 —— MSVC C 里 wchar_t 就是
     *     unsigned short, 两者是同一类型, 会被判重复关联 (C7700, 见 vb6rtl_variant.h
     *     VT_UI4 段那段话)。正确做法就是这里点名调各自的出口。
     *   顺带 LongLong/VT_I8 也一并点名(以前同样靠 _Generic 的 `long long:` 猜)。 */
    case Vb6Type::SByte:
        return "vb6_VariantSByte((int8_t)(" + cExpr + "))";
    case Vb6Type::UInteger:
        return "vb6_VariantUInteger((uint16_t)(" + cExpr + "))";
    case Vb6Type::ULong:
        return "vb6_VariantULong((uint32_t)(" + cExpr + "))";
    case Vb6Type::ULongLong:
        return "vb6_VariantULongLong((uint64_t)(" + cExpr + "))";
    case Vb6Type::LongLong:
        return "vb6_VariantLongLong((int64_t)(" + cExpr + "))";
    case Vb6Type::LongPtr:
        // 账 #230 §B137（第二十九刀）：句柄读进 Variant 槽以前落的是 _Generic 的 default
        // 一档 = vb6_VariantObject，于是**一枚 HWND 被当 IDispatch\* 装走**。实测读数（两台同形，
        // 所以不是第二十八刀带来的）：`Dim v As Variant: v = arrA(0).hWnd` 之后 VarType 答 9、
        // TypeName 借宿主表答控件名（"TextBox"）、CStr 交空串、CLng 答 0、与任何数比较恒不等。
        // 不是野调用：vb6_ComAddRefDispatch 里的 vb6_ComIsDispatchable 对 IsWindow 的句柄直接判否
        // （vb6com_invoke.c 那条注释就是为这一形写的），代价是这一档看着像对象、值面全错。
        // 按 intptr_t 走宽度：x86 = long → VT_I4，正是 VB6 对 hWnd 的答案（VarType 3 / TypeName
        // "Long"）；x64 = long long → VT_I8，不截断，与 Fix 082 的 VarPtr 档同族。
        return "vb6_VariantFromValue((intptr_t)(" + cExpr + "))";
    default: break;
    }
    return "vb6_VariantFromValue(" + cExpr + ")";
}

// ============================================================
// 账 #611: 工程内 .cls 的默认成员名 (AST 侧)。
//
// scripts/check_uc_default_member.ps1 的 R1 要求「认这条属性行的文件恰好一个」
// —— 所以 .cls 这一支也住在本文件, driver 只调它填表, 不许在别处再写一份解析。
// 与 findUcDefaultMember 同一份判据 (属性名后缀 + 值字面为 0 = DISPID_VALUE,
// -4 是 NewEnum 不折, 成员名取点号前那段保留原大小写), 差别只在输入形态:
//   · UC    —— .ctl 代码段文本, 还要同名 `Property Get ... As <型>` 定返回型;
//   · .cls  —— parser 已把属性行收成过程体/模块头的 AttributeStmt, 且不要求
//              同名 Property Get (JSonWrite.cls 的默认成员是 Property Let Item)。
// ============================================================
bool findClassDefaultMember(const Module& mod, std::string& propOut) {
    auto scanOne611 = [&propOut](const AttributeStmt& a611) {
        const std::string low611 = Symbol::toLower(a611.attrName);
        static const std::string kSuffix611 = ".vb_usermemid";
        if (low611.size() <= kSuffix611.size()
            || low611.compare(low611.size() - kSuffix611.size(),
                              kSuffix611.size(), kSuffix611) != 0) return false;
        if (!a611.value || a611.value->kind != ASTNodeKind::LiteralExpr) return false;
        std::string raw611;
        for (char c611 : static_cast<const LiteralExpr&>(*a611.value).rawText) {
            if (!std::isspace(static_cast<unsigned char>(c611))) raw611 += c611;
        }
        if (raw611 != "0") return false;   // 只认 DISPID_VALUE
        const size_t dot611 = a611.attrName.find('.');
        if (dot611 == std::string::npos || dot611 == 0) return false;
        propOut = a611.attrName.substr(0, dot611);
        return true;
    };
    for (const auto& d611 : mod.declarations) {
        if (!d611) continue;
        const StmtList* body611 = nullptr;
        if (d611->kind == ASTNodeKind::PropertyDecl) {
            body611 = &static_cast<const PropertyDecl&>(*d611).body;
        } else if (d611->kind == ASTNodeKind::FunctionDecl) {
            body611 = &static_cast<const FunctionDecl&>(*d611).body;
        } else if (d611->kind == ASTNodeKind::SubDecl) {
            body611 = &static_cast<const SubDecl&>(*d611).body;
        }
        if (!body611) continue;
        for (const auto& s611 : *body611) {
            if (!s611 || s611->kind != ASTNodeKind::AttributeStmt) continue;
            if (scanOne611(static_cast<const AttributeStmt&>(*s611))) return true;
        }
    }
    // 少数工程把属性行留在模块头 (VB_Name 那一组), 同一判据再扫一遍。
    for (const auto& a611 : mod.attributes) {
        if (a611 && scanOne611(*a611)) return true;
    }
    return false;
}
} // namespace vb6c3
