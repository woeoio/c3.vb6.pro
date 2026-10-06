#include "backend/cgen.hpp"
#include "common/float_literal.hpp"  // 账 #188: 浮点字面量的单一出口
#include "common/int_literal.hpp"   // 账 #194: 整数字面量的单一出口
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
    // C29-SL-e: `ToolTipText` / `Tag` 是 VB6 的 **String** 通用属性，但两条 RTL getter 原本
    // 声明成 `void*`，而装箱那一步用的是 C11 `_Generic`（`vb6rtl_variant.h` 的表：
    // "其他指针 (void*/class*/type*) -> VariantObject"）⇒ 裸指针被当**对象**装箱。
    // 实测（探针工程 .build/sltt，改之前）：`CStr(Slider1.ToolTipText)` 打**空**、
    // `TypeName(...)` 答 **"Object"**、`If Slider1.Tag = "tag1"` 答**假**（那一条比
    // ToolTipText 更绕：比较面先装箱再 CStr，等于拿 "" 去比），而 `Len(...)` 与
    // `InStr(...)` 这些**直接拿指针**的面反而是对的 —— 同一枚属性两种答案，就是没登记的证状。
    // 登记成 String 之后所有消费面统一按 BSTR 取值（与 `vb6_GetControlText` 那条同档），
    // RTL 侧的返回型也一起改成 `wchar_t*`，让**万一**还剩的装箱落到 VT_BSTR。
    // 同一段里 `MouseIcon` / `Hwnd` 的 getter 也是 `void*`，但那两条**本来就是对象**，
    // 装箱成对象是对的（FlexGrid 的 `CellPicture` 同型），刻意不动。
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

std::string CCodeGen::getControlPropReadFn(FrmControlType ctrlType, const std::string& propName) const {
    std::string propLower = propName;
    std::transform(propLower.begin(), propLower.end(), propLower.begin(), ::tolower);

        // P11.8: Common properties for all visible controls (checked before switch)
    if (propLower == "left") return "vb6_GetControlLeft";
    if (propLower == "top") return "vb6_GetControlTop";
    if (propLower == "width") return "vb6_GetControlWidth";
    if (propLower == "height") return "vb6_GetControlHeight";
    if (propLower == "hwnd") return "vb6_GetControlHwnd";
    // P13.1: Font properties (all visible controls with text)
    // D6 / C29-9: CommonDialog 的 FontName / FontSize 是**对话框字段**（ChooseFont 的 LOGFONT），
    // 不是控件字体 —— 这枚控件没有外观。通用那一组查在类型 switch **之前**，不挡就把
    // `CD1.FontName = "Consolas"` 静默落到字体属性上（C29-1a 那条 borderstyle 被抢走同一类碰撞）。
    if (propLower == "fontname" && ctrlType != FrmControlType::CommonDialog) return "vb6_GetControlFontName";
    if (propLower == "fontsize" && ctrlType != FrmControlType::CommonDialog) return "vb6_GetControlFontSize";
    // C29-SL-p 的判据证人（**不是 VB6 属性**，只给读侧、不给写口）：窗口真在用的字体像素高度。
    // 需要它是因为自存把"存什么读什么"做对了之后，读回来那个数已经问不出窗口了 ——
    // 与 TickPresent / TravelIsVert / ToolTipRegistered 同族（#148 那条"证人问不出东西"的教训）。
    if (propLower == "fontpixelheight" && ctrlType != FrmControlType::CommonDialog)
        return "vb6_ControlFontPixelHeight";
    if (propLower == "fontbold") return "vb6_GetControlFontBold";
    if (propLower == "fontitalic") return "vb6_GetControlFontItalic";
    if (propLower == "fontunderline") return "vb6_GetControlFontUnderline";
    if (propLower == "fontstrikethrough") return "vb6_GetControlFontStrikethrough";
    // P13.2: Color properties (all visible controls)
    if (propLower == "forecolor") return "vb6_GetControlForeColor";
    if (propLower == "backcolor") return "vb6_GetControlBackColor";
    // P13.5: Alignment (all text controls)
    if (propLower == "alignment") return "vb6_GetAlignment";
    // P13.6: TabIndex/TabStop (all visible controls)
    if (propLower == "tabindex") return "vb6_GetTabIndex";
    if (propLower == "tabstop") return "vb6_GetTabStop";
    if (propLower == "causesvalidation") return "vb6_GetCausesValidation";
    // P13.8: ToolTipText (all visible controls)
    if (propLower == "tooltiptext") return "vb6_GetToolTipText";
    // P13.9: Tag (all controls)
    if (propLower == "tag") return "vb6_GetControlTag";
    // P13.7: MousePointer/MouseIcon (all visible controls)
    if (propLower == "mousepointer") return "vb6_GetMousePointer";
    if (propLower == "mouseicon") return "vb6_GetMouseIcon";
    // P13.10: BorderStyle (all visible controls)
    // C29-1a: Shape/Line 的 BorderStyle 是"画笔线型"(0..6), 与窗口边框样式(0/1)同名
    // 不同物 —— 让这两个类型走下面各自的 case, 否则读回来的永远是窗口那套。
    if (propLower == "borderstyle" && ctrlType != FrmControlType::Shape
        && ctrlType != FrmControlType::Line) return "vb6_GetBorderStyle";
    // Fix <vbeclipse>: ScaleWidth/ScaleHeight 是**所有**控件的通用属性 (不只 Form) —
    // ucTabStrip.ctl 的 `With picButtons` 里读 ScaleWidth/ScaleHeight 落到 HWND 结构体
    // 字段上 → C2039: "ScaleWidth" 不是 "HWND__" 的成员 (×3, 同式还带出 vb6_ControlMove
    // 的 C2198). 放在 switch 之前, 与 Tag/MousePointer 同口径 (RTL: vb6forms_widget.c
    // 的 vb6_GetScaleWidth/Height, 声明在 vb6forms_prop_form.h, 由 vb6forms.h 传递可见).
    if (propLower == "scalewidth") return "vb6_GetScaleWidth";
    if (propLower == "scaleheight") return "vb6_GetScaleHeight";

    switch (ctrlType) {
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
        if (propLower == "caption") return "vb6_GetControlText";
        if (propLower == "visible") return "vb6_GetControlVisible";
        if (propLower == "enabled") return "vb6_GetControlEnabled";
        // Fix 056: Form-specific properties
        if (propLower == "windowstate") return "vb6_GetWindowState";
        if (propLower == "scalewidth") return "vb6_GetScaleWidth";
        if (propLower == "scaleheight") return "vb6_GetScaleHeight";
        if (propLower == "scalemode") return "vb6_WindowScaleModeSelf";  // 账 #197: 与写侧成对
        // 账 #233: Form 的**笔位与画笔粗细**。此前这四条住在 cgen_expr_member_form_builtin.inc
        // 的一份侧表里 —— 那份表只有读侧, 于是 `Me.DrawWidth = 3` 退化成
        // `vb6_Form_DrawGetWidth(vb6_hwnd_X) = 3;` (C2106, 实测两架构都编不过)。
        // 收进这两张表之后读写成对, 侧表删掉。forecolor **不在这里** —— 它早已由上面
        // 那条通用行答给 vb6_GetControlForeColor (侧表那一条从没命中过), 与绘图家族
        // 自带的 VB6_DrawForeColor 是两份存储, 那一问另开账。
        if (propLower == "currentx") return "vb6_GetCurrentX";   // 笔位只有一份存储 (float)
        if (propLower == "currenty") return "vb6_GetCurrentY";
        if (propLower == "drawwidth") return "vb6_Form_DrawGetWidth";
        if (propLower == "hdc") return "vb6_GetControlHDC";  // 账 #196: Form.hDC 同一处出口
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
        if (propLower == "caption") return "vb6_GetControlText";
        if (propLower == "picture") return "vb6_GetControlPicture";
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
        if (propLower == "getnumticks") return "vb6_Slider_GetNumTicks";
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
    if (propLower == "alignment") return "vb6_SetAlignment";
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
        if (propLower == "scalemode") return "vb6_SetScaleMode";  // 账 #197: 与读侧成对
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

// C29-SL-l（账 #143）: 把一条"控件对象表达式"折成 (控件类型, 那枚窗口的 C 表达式)。
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

// C29-SL-l（账 #143）: 控件的**零实参方法**名表（命中回 C 里的函数名，否则空串）。
// 焦点面只给"真能拿焦点"的那批窗口型控件 —— Label / Image / Shape / Line / Timer / Menu /
// Data / OLE / CommonDialog / Winsock / ImageList / StatusBar / ProgressBar 刻意不接：
// 真 VB6 在那里是 raise 一个错误号，而本项目还没有运行期错误面，"什么都不做"比伪造成功诚实。
std::string CCodeGen::controlZeroArgMethod(FrmControlType ctrlType,
                                          const std::string& memberLower) const {
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
                return "vb6_SetControlFocus";
            default:
                return "";
        }
    }
    if (memberLower == "clearsel" && ctrlType == FrmControlType::Slider)
        return "vb6_Slider_ClearSel";
    // 账 #192: ListBox / ComboBox 的 Clear。VB6 里它是方法而不是属性，且只有这两枚
    // 有清空语义（TreeView/ListView 的清是各自那一族，另有出口）。
    if (memberLower == "clear"
        && (ctrlType == FrmControlType::ListBox || ctrlType == FrmControlType::ComboBox))
        return "vb6_ClearList";
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
                                          const std::string& memberLower) const {
    if (memberLower != "textheight" && memberLower != "textwidth") return "";
    switch (ctrlType) {
        case FrmControlType::Form:
        case FrmControlType::PictureBox:
            return memberLower == "textheight" ? "vb6_ControlTextHeight"
                                               : "vb6_ControlTextWidth";
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
                                         const std::string& memberLower) const {
    if (memberLower != "scalex" && memberLower != "scaley") return "";
    switch (ctrlType) {
        case FrmControlType::Form:
        case FrmControlType::PictureBox:
            return memberLower == "scalex" ? "vb6_ScaleUnitX" : "vb6_ScaleUnitY";
        default:
            return "";
    }
}
// 账 #221 = C29-PL-a（语料物证 Charts 2020/ucProgressCircular 的 ppProgressCircular.pag
// 那 8 条 `Picture1.Line` / `Picture2.Line`）: 画表面方法的名字表。
// 这一族以前**没有表** —— parser 的 Fix 102 把坐标对折进实参表之后，注释写着"由后端按控件
// 类型发射"，但后端从来没接这一刀，于是整条调用落到通用 COM 兜底
// （`ComGetObjectProp(hwnd, L"Line")` 再对它取 `Item`，两跳都被 RTL 登记成"认识但什么都不做"），
// 症状是**编得过、跑得起、一笔不画**。档位与 TextHeight/ScaleX 两族同样只给 Form 与 PictureBox。
std::string CCodeGen::controlCanvasMethod(FrmControlType ctrlType,
                                          const std::string& memberLower) const {
    if (memberLower != "line") return "";
    switch (ctrlType) {
        case FrmControlType::PictureBox:
            return "vb6_ControlLine";
        default:
            // Form 那一档刻意**不给**: 现在只有 PictureBox 那一处成员侧发了标记
            // (cgen_expr_member_form_builtin.inc 的 Fix 185 那块), 给了就是"广告比应答复"。
            // 接 Form 之前先把 `Me.Line` / 窗体自绘那条码头找出来。
            return "";
    }
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

// 账 #83(a)（C29-SL-r）: VB6 的 `TabStop` 默认 True，而本项目**两条创建路以前都不立 WS_TABSTOP**
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

long CCodeGen::controlContainerExStyleBit(const FrmControl& ctrl) const {
    // ⚠ 账 #165 的根因就在这个数上：`WS_EX_CONTROLPARENT` 在 SDK 头里是 **0x00010000**
    //   （winuser.h:2855），而 0x00040000 是 `WS_EX_APPWINDOW`（同表往下 2857 行）。以前这里
    //   手抄成了后者 ⇒ 那两轮的读数只证明"想发的数落到窗口上了"，证不了"那是管理器认的那一位"。
    //   通用式：**手抄常量一律去 SDK 头对值**（或直接引符号），证人行打印符号对应的值。
    constexpr long kWsExControlParent = 0x00010000L;  // WS_EX_CONTROLPARENT（winuser.h 实测值）
    return controlIsContainerType(ctrl) ? kWsExControlParent : 0L;
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
            style |= kWsBorder | kEsAutoH;
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
        case FrmControlType::PictureBox:
            style |= kSsBitmap | kSsCenterImg | kWsBorder;
            break;
        case FrmControlType::Image:
            style |= kSsBitmap | kSsCenterImg;
            break;
        // C29-SL-h: 此前这条路**压根没有 Slider 这一格** —— 容器里的滑杆连 TBS_AUTOTICKS 都没
        // 挂上（账 #83 那条"顶层专有项没铺到第二条创建路"的一个具体落点）。折算共用
        // sliderStyleBits，不留第二份表。
        case FrmControlType::Slider:
            style |= sliderStyleBits(ctrl);
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
    if (it == knownFormControls_.end()) return false;
    outType = it->second;
    return true;
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
// P8.4: Variant值包装 - 根据表达式类型推断Variant构造函数
// ============================================================

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

    // Fix 090d2: C 表达式顶层已是 vb6_VARIANT 时直接返回, 不再包装 —
    // 原 2998 行的检查在 inferExprType switch 之后, 对符号表/推断层
    // 认为是具体标量但实际生成 vb6_VARIANT 的表达式 (如 COM 属性
    // fileInfo.Size 被推断为 Long, 实际生成
    // vb6_VariantFromComResult(vb6_ComCall(...))) 会被 switch 抢先包装
    // 成 vb6_VariantLong(vb6_VARIANT) → C2440 (cHttpServerResponse File
    // 内 fileSize = fileInfo.Size, fileSize As Variant).
    if (cExprIsVariant(cExpr)) {
        return cExpr;
    }

    // Fix 170: 右侧是 VB6 整体数组引用 `A()` (空括号) → 装箱成持有数组的 Variant。
    // 必须先于下面的 inferExprType switch: 它对 `A()` 可能给回**元素**类型 (Byte/Long),
    // 于是 vb6_VariantLong(载体指针) → C 侧只是警告, 指针被截断成 int32, 运行期才炸。
    if (valueNode->kind == ASTNodeKind::IndexOrCallExpr
        && isWholeArrayRef(static_cast<const Expr*>(valueNode))) {
        return "vb6_VariantArray((void*)" + cExpr + ")";
    }
    
    // 使用inferExprType推断表达式类型
    if (valueNode->kind == ASTNodeKind::BinaryExpr ||
        valueNode->kind == ASTNodeKind::UnaryExpr ||
        valueNode->kind == ASTNodeKind::LiteralExpr ||
        valueNode->kind == ASTNodeKind::IdentifierExpr ||
        valueNode->kind == ASTNodeKind::IndexOrCallExpr ||
        valueNode->kind == ASTNodeKind::MemberAccessExpr) {
        Vb6Type vtype = inferExprType(static_cast<Expr&>(*valueNode));
        // <vbeclipse>: 这张表只留"与 _Generic 结果一致"的四档 (String/Long/Integer/Double)。
        // Boolean/Byte/Single/Date/Currency 交给末尾的 boxToVariant —— 以前两份表各写一份
        // 档位 (账 #123 改了这份的 Byte、Fix 198 改了那份的 Boolean), 于是同一类型的装箱
        // 结果取决于表达式走哪条路, DT43 的布尔就是这么漏的。
        switch (vtype) {
            case Vb6Type::String:    return "vb6_VariantString(" + cExpr + ")";
            case Vb6Type::Long:
            case Vb6Type::Integer:   return "vb6_VariantLong(" + cExpr + ")";
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
    
    // COM后期绑定调用: vb6_ComCall返回VARIANT*, 需转为vb6_VARIANT
    if (cExpr.find("vb6_ComCall(") == 0) {
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
    return cExprIsVariant(cExpr) || cExprIsVariantCarrierStr(cExpr);
}

// Fix 198 + <vbeclipse>: **装箱的唯一权威**。按 VB 声明类型选档, 只在类型确实与
// C 表示不一致时才改写; 其余一律逐字节退回 vb6_VariantFromValue (_Generic 按 C 类型
// 选 ctor) —— 早退式的"已是 VARIANT 就不包"看着更干净, 但它会把存量码也一起改了
// (护栏实测 VbEclipse 的 VB6_SA_AT 实参少了那层恒等包装)。
// 需要显式档案的四种 (C 表示区分不出来或就是错的):
//   Boolean → VT_BOOL(11)   C 侧是 int16_t/int32_t, _Generic 会装成 VT_I4
//   Byte    → VT_UI1(17)    C 侧 uint8_t 走得到对档, 但模块级 Byte 曾被推断成 Long
//   Single  → VT_R4(4)      _Generic 的 float 档以前也升到 VT_R8 (哨兵实测 5)
//   Date    → VT_DATE(7)    C 侧就是 double, 只有 VB 类型能说话 (哨兵实测 5)
// 不显式改写的 (Integer/Long/Double/String/Object) 在 _Generic 下与 VB6 同档, 保持
// 原样以免动到存量发码。
std::string CCodeGen::boxToVariant(Expr* expr, const std::string& cExpr) const {
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
        // 精确匹配裸 vb6_ComCall( / vb6_ComGetProp( / vb6_ComGetObjectProp(
        // (12/15/21 字符含左括号); 不误伤 vb6_ComCallInt/BSTR/Double/Object/ByDispid.
        size_t cs = 0;
        while (cs < cExpr.size()
               && (cExpr[cs] == '(' || cExpr[cs] == ' ' || cExpr[cs] == '\t'
                   || cExpr[cs] == '\n' || cExpr[cs] == '\r')) cs++;
        bool comRes = (cExpr.compare(cs, 12, "vb6_ComCall(") == 0)
                   || (cExpr.compare(cs, 15, "vb6_ComGetProp(") == 0)
                   || (cExpr.compare(cs, 21, "vb6_ComGetObjectProp(") == 0);
        if (comRes) return "vb6_VariantFromComResult(" + cExpr + ")";
    }
    if (!expr) return "vb6_VariantFromValue(" + cExpr + ")";
    // 已经是 VARIANT 的表达式交给 _Generic 的 vb6_VariantIdentity 档恒等直传,
    // 绝不再按 VB 类型强装 (否则就是上面那条双装)。
    if (cExprIsVariantCarrier(cExpr)) return "vb6_VariantFromValue(" + cExpr + ")";
    switch (inferExprType(*expr)) {
    case Vb6Type::Boolean:
        // 形参是 int16_t: 显式收窄, 兼容 _Bool/int 两种 C 侧布尔表示.
        return "vb6_VariantBool((int16_t)(" + cExpr + "))";
    case Vb6Type::Byte:
        return "vb6_VariantByte((uint8_t)(" + cExpr + "))";
    case Vb6Type::Single:
        return "vb6_VariantSingle((float)(" + cExpr + "))";
    case Vb6Type::Date:
        return "vb6_VariantDate(" + cExpr + ")";
    default: break;
    }
    return "vb6_VariantFromValue(" + cExpr + ")";
}
} // namespace vb6c3
