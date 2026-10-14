#pragma once
// 宿主伪成员的唯一一张表 (kHostPseudoRows) 的家。
//
// 账 #159 把它收成一张表时住在 src/backend/cgen_util_com.cpp; 账 #219 让**语义层**也要问它
// (裸写的文档成员不许再报 VB3001)。同一个事实拷两份就是两个权威: 表里新增一行 HPF_BARE
// 而语义层没跟上, 产物照旧是正确的 C，再配一条噪声。所以表搬到 common (与
// float_literal.hpp / int_literal.hpp 同一家族: 两层都能问的那些单一出口)。
// 表本体、类型口径与那段症状记录原样搬迁, 叙事仍在这里。

#include "common/types.hpp"
#include <cctype>
#include <string>

namespace vb6c3 {

// ============================================================
// 账 #159 = C29-CH-d: 宿主伪成员 (UserControl / PropertyPage / Extender /
// Ambient) 的**唯一一张表**。
//
// 症状: 控件代码里 `UserControl.hWnd <> 0` 恒假、`CStr(UserControl.hWnd)` 打空。
// 根因不在值, 在**读法** —— 实测 (临时探针, 用完已撤): push 时 vb6_UserControl_hWnd
// 就是真 HWND。这些成员在 C 侧是 RTL 全局量 (vb6rtl_userctl.h 的 extern
// int32_t / int16_t / void* / BSTR), 而 inferExprType 认不得它们 ⇒ 答 Variant ⇒
// 比较发成 `vb6_VarCmpLongNe(&vb6_UserControl_hWnd, 0)`: 第一形参是 vb6_VARIANT*
// (16 字节), 递过去的是 8 字节 void* 的**地址** ⇒ 读到的 vt/lVal 全是越界垃圾。
// CStr 那一路是 `vb6_VariantFromValue(void*)` 落进 _Generic 的
// `default: vb6_VariantObject` (vb6rtl_variant.h:215) ⇒ 按对象装箱 ⇒ 空串。
//
// 同一个决定此前抄在五处 (全在 cgen 侧), 覆盖面还彼此不一致 (裸名答 Long 而限定名答 Variant,
// 反之亦然): cgen_util_com.cpp 的规范化表 (只有 hdc)、cgen_expr_ident_builtin.inc 的两张裸名
// 表、cgen_util_type.cpp 的两条硬编码 (裸 scalewidth/scaleheight; 限定 7 枚一律
// Long)、cgen_assign_host_pseudo.inc 的 kNumericHostMembers。本表收成一格,
// 四个消费点只问这里。
//
// 类型口径 (由 scripts/check_host_pseudo_table.ps1 逐行对 RTL 声明钉):
//   int32_t → Long    int16_t → Boolean (VB6 布尔本就 16 位, -1=True)
//   void*   → LongPtr (句柄/指针成员)   BSTR  → String
// type=Unknown 的行是**对象成员或方法** —— 维持改动前的答案 (Variant), 本账不动
// 它们的值面; HPF_METHOD 额外表示"裸名以调用形态出现", 一律不做值读。
// HPF_BARE = 允许在 .ctl/.pag 里裸写 (VB6 里等价于 <对象>.<成员>)。
//
// 不在表里的成员 = 本表刻意不收: RTL 没有对应全局 (vb6_UserControl_Left/Top、
// vb6_PropertyPage_ScaleWidth、Appearance、BorderStyle、UserControl.Parent …),
// 收了就是发一个未声明符号。那是 RTL 侧的缺口, 另立账。
// ============================================================


// 把成员名归到小写后再查表。这里自己拼而不用 Symbol::toLower: common 不得向上依赖
// semantics (拼出来的是同一件事: ASCII 小写)。
inline std::string hostPseudoLower(const std::string& s) {
    std::string o(s);
    for (char& ch : o) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
    return o;
}

enum : uint8_t { HPF_NONE = 0, HPF_BARE = 1 << 0, HPF_METHOD = 1 << 1,
                 // HPF_CHANNEL = 这一档的答案**不是** `vb6_<对象>_<rtl>`，而是一条专用码头
                 // (channel 字段给符号名)。账 #278 §B105 的 b 案：表里必须能说得出"由哪条码头回答"，
                 // 否则给 .pag 补一行 `Controls` 就是在说谎 —— vb6_PropertyPage_Controls
                 // 那枚恒 NULL 的空桩从来不该被发码 (真实答复是 vb6_UC_Controls())。
                 // 与 canvas_drawing.hpp 的 CANVAS_OWNER_METHOD/DRAW 同一形状。
                 HPF_CHANNEL = 1 << 2 };

struct HostPseudoRow {
    // 宿主伪对象名 —— 存的是 **RTL 符号里那一份拼法** (PascalCase)，查的时候两边都折小写。
    // 账 #278 §B120: 这一列从前是小写，于是 `vb6_<Obj>_<rtl>` 里对象那一段没有权威，
    // 发码侧三处各拼一次 (其中一处还把源码拼写直接抄进符号名)。
    const char* obj;     // 发成 vb6_<obj>_<rtl>
    const char* name;    // VB6 成员名 (小写)
    const char* rtl;     // RTL 侧拼写 (C 大小写敏感, 源码拼写一律归一到它)
    Vb6Type type;        // 值读类型; Unknown = 不由本表回答
    uint32_t flags;      // uint32 而非 uint8: 行里写 HPF_BARE | HPF_METHOD 会整型提升,
                         // 花括号初始化里收窄回 uint8_t 是 narrowing (MSVC 直接报错)
    // 只在 HPF_CHANNEL 那几行填：码头的 C 符号名（**不带括号**，发码侧自己加 `()`）。
    // 缺省 nullptr ⇒ 其余行沿用 `vb6_<对象>_<rtl>` 那条契约。
    const char* channel = nullptr;
};

inline const HostPseudoRow kHostPseudoRows[] = {
    // ---- UserControl (.ctl) ----
    {"UserControl", "accesskeys",      "AccessKeys",      Vb6Type::String,   HPF_NONE},
    {"UserControl", "appearance",      "Appearance",      Vb6Type::Long,     HPF_NONE},
    {"UserControl", "asyncread",       "AsyncRead",       Vb6Type::Unknown,  HPF_BARE | HPF_METHOD},
    {"UserControl", "autoredraw",      "AutoRedraw",      Vb6Type::Boolean,  HPF_NONE},
    {"UserControl", "backcolor",       "BackColor",       Vb6Type::Long,     HPF_NONE},
    // Fix 255: `UserControl.BackStyle` (ComCtlsDemo LabelW.ctl:936/942 ——
    // `BackStyle = UserControl.BackStyle` 与 `UserControl.BackStyle = Value`)。
    // 表里此前没有这一行, 而发码侧 hostPseudoRtlSymbol 在表里查不到就不再挡,
    // 直接按 `vb6_<对象>_<成员名>` 拼出 vb6_UserControl_BackStyle —— RTL 没这枚
    // 全局 ⇒ C2065 ×2 (LabelW.c 1608/1618)。补行 + 补 RTL 声明与定义 (int32_t/Long,
    // 与同一族的 BackColor/ForeColor 逐字同款, 也满足 check_host_pseudo_table 的类型配对)。
    {"UserControl", "backstyle",       "BackStyle",       Vb6Type::Long,     HPF_NONE},
    {"UserControl", "cancelasyncread", "CancelAsyncRead", Vb6Type::Unknown,  HPF_METHOD},
    {"UserControl", "cls",             "Cls",             Vb6Type::Unknown,  HPF_METHOD},
    {"UserControl", "containedcontrols", "ContainedControls", Vb6Type::Unknown, HPF_NONE},
    {"UserControl", "containerhwnd",   "ContainerHwnd",   Vb6Type::LongPtr,  HPF_BARE},
    {"UserControl", "controls",        "Controls",        Vb6Type::Unknown,  HPF_BARE | HPF_CHANNEL, "vb6_UC_Controls"},
    {"UserControl", "drawstyle",       "DrawStyle",       Vb6Type::Long,     HPF_NONE},
    {"UserControl", "enabled",         "Enabled",         Vb6Type::Boolean,  HPF_BARE},
    {"UserControl", "eventsfrozen",    "EventsFrozen",    Vb6Type::Boolean,  HPF_NONE},
    {"UserControl", "extender",        "Extender",        Vb6Type::Unknown,  HPF_NONE},
    {"UserControl", "forecolor",       "ForeColor",       Vb6Type::Long,     HPF_NONE},
    {"UserControl", "hdc",             "hDC",             Vb6Type::LongPtr,  HPF_BARE},
    {"UserControl", "height",          "Height",          Vb6Type::Long,     HPF_NONE},
    {"UserControl", "hwnd",            "hWnd",            Vb6Type::LongPtr,  HPF_BARE},
    {"UserControl", "mouseicon",       "MouseIcon",       Vb6Type::Unknown,  HPF_NONE},
    {"UserControl", "mousepointer",    "MousePointer",    Vb6Type::Long,     HPF_NONE},
    {"UserControl", "oledrag",         "OLEDrag",         Vb6Type::Unknown,  HPF_METHOD},
    {"UserControl", "oledropmode",     "OLEDropMode",     Vb6Type::Long,     HPF_NONE},
    {"UserControl", "picture",         "Picture",         Vb6Type::Unknown,  HPF_NONE},
    {"UserControl", "propertychanged", "PropertyChanged", Vb6Type::Unknown,  HPF_BARE | HPF_METHOD},
    {"UserControl", "refresh",         "Refresh",         Vb6Type::Unknown,  HPF_METHOD},
    {"UserControl", "righttoleft",     "RightToLeft",     Vb6Type::Integer,  HPF_NONE},
    {"UserControl", "scaleheight",     "ScaleHeight",     Vb6Type::Long,     HPF_BARE},
    {"UserControl", "scalemode",       "ScaleMode",       Vb6Type::Long,     HPF_BARE},
    {"UserControl", "scalewidth",      "ScaleWidth",      Vb6Type::Long,     HPF_BARE},
    {"UserControl", "scalex",          "ScaleX",          Vb6Type::Unknown,  HPF_BARE | HPF_METHOD},
    {"UserControl", "scaley",          "ScaleY",          Vb6Type::Unknown,  HPF_BARE | HPF_METHOD},
    {"UserControl", "size",            "Size",            Vb6Type::Unknown,  HPF_METHOD},
    {"UserControl", "textheight",      "TextHeight",      Vb6Type::Unknown,  HPF_METHOD},
    {"UserControl", "textwidth",       "TextWidth",       Vb6Type::Unknown,  HPF_METHOD},
    {"UserControl", "width",           "Width",           Vb6Type::Long,     HPF_NONE},    // 对象成员只登记用于**拼写规范化**; 值面一律不答 (各有专用通道: Controls 集合走
    // vb6_UC_Controls()、Parent 链由 Fix 133u 在 cgen_base.cpp 改写、Font 是
    // vb6_ComIface_Font*, 装箱比较本来就不该按标量发)。
    {"UserControl", "ambient",         "Ambient",         Vb6Type::Unknown,  HPF_NONE},
    {"UserControl", "font",            "Font",            Vb6Type::Unknown,  HPF_NONE},
    // 账 #257: ParentControls 的真实答复是 vb6_UC_ParentControls() (父容器的集合,
    // uc_controls.c), 恒 NULL 的空桩 vb6_UserControl_ParentControls 不该被发码 ——
    // `UserControl.ParentControls.Count` 以前发成 vb6_ComGetIntProp((void*)NULL, L"Count")
    // ⇒ 崩在 vb6_ComUnpackInt (StatusBar 创建期 100% 复现, StatusBar.ctl:481 就在
    // UserControl_InitProperties 里)。改道在 cgen_expr_member_generic_access.inc。
    //
    // ⚠ 刻意**不**给这行挂 HPF_CHANNEL(像 Controls 那样): check_host_pseudo_table.ps1 的
    // 行正则以 `(HPF_[A-Z| ]+)\}` 收尾, **认不出带 channel 的行** ⇒ 挂了这一行就掉出解析,
    // 行数 62→61 直接红(且那一行从此没人查)。要挂 channel 得先修那道针的正则。
    {"UserControl", "parentcontrols",  "ParentControls",  Vb6Type::Unknown,  HPF_NONE},

    // ---- PropertyPage (.pag) ----
    {"PropertyPage", "changed",          "Changed",          Vb6Type::Boolean, HPF_BARE},
    // 账 #278 §B105: .pag 里裸写 `Controls.Add/Remove/Item` 与 .ctl 同一枚码头 (vb6_UC_Controls
    // 按当前实例回落 formHwnd)。以前表里没有这一行 ⇒ 语义层问不到就配一条 VB3001，
    // 而发码那条路一直是对的 (`vb6_ComCallObject((void*)vb6_UC_Controls(), L"Add", …)`)。
    // 刻意不复用 `vb6_PropertyPage_Controls` —— 那枚是恒 NULL 的空桩，发出去是"枚举得空集"。
    {"PropertyPage", "controls",         "Controls",         Vb6Type::Unknown, HPF_BARE | HPF_CHANNEL, "vb6_UC_Controls"},
    {"PropertyPage", "hwnd",             "hWnd",             Vb6Type::LongPtr, HPF_BARE},
    {"PropertyPage", "scaleheight",      "ScaleHeight",      Vb6Type::Long,    HPF_BARE},
    {"PropertyPage", "scalemode",        "ScaleMode",        Vb6Type::Long,    HPF_BARE},
    {"PropertyPage", "selectedcontrols", "SelectedControls", Vb6Type::Unknown, HPF_BARE | HPF_METHOD},

    // ---- Extender (容器提供的扩展对象) ----
    {"Extender", "align",           "Align",           Vb6Type::Long,    HPF_NONE},
    {"Extender", "cancel",          "Cancel",          Vb6Type::Boolean, HPF_NONE},
    {"Extender", "container",       "Container",       Vb6Type::Unknown, HPF_NONE},
    {"Extender", "default",         "vb6_Default",     Vb6Type::Boolean, HPF_NONE},
    {"Extender", "drag",            "Drag",            Vb6Type::Unknown, HPF_METHOD},
    {"Extender", "dragicon",        "DragIcon",        Vb6Type::Unknown, HPF_NONE},
    {"Extender", "dragmode",        "DragMode",        Vb6Type::Long,    HPF_NONE},
    {"Extender", "height",          "Height",          Vb6Type::Long,    HPF_NONE},
    {"Extender", "helpcontextid",   "HelpContextID",   Vb6Type::Long,    HPF_NONE},
    {"Extender", "left",            "Left",            Vb6Type::Long,    HPF_NONE},
    {"Extender", "setfocus",        "SetFocus",        Vb6Type::Unknown, HPF_METHOD},
    {"Extender", "tag",             "Tag",             Vb6Type::String,  HPF_NONE},
    {"Extender", "tooltiptext",     "ToolTipText",     Vb6Type::String,  HPF_NONE},
    {"Extender", "top",             "Top",             Vb6Type::Long,    HPF_NONE},
    {"Extender", "visible",         "Visible",         Vb6Type::Boolean, HPF_NONE},
    {"Extender", "whatsthishelpid", "WhatsThisHelpID", Vb6Type::Long,    HPF_NONE},
    {"Extender", "width",           "Width",           Vb6Type::Long,    HPF_NONE},
    {"Extender", "zorder",          "ZOrder",          Vb6Type::Unknown, HPF_METHOD},

    // ---- Ambient (宿主环境) ----
    {"Ambient", "backcolor",   "BackColor",   Vb6Type::Long,     HPF_NONE},
    {"Ambient", "displayasdefault", "DisplayAsDefault", Vb6Type::Boolean,  HPF_NONE},
    {"Ambient", "displayname", "DisplayName", Vb6Type::String,   HPF_NONE},
    {"Ambient", "forecolor",   "ForeColor",   Vb6Type::Long,     HPF_NONE},
    {"Ambient", "font",        "Font",        Vb6Type::Unknown,  HPF_NONE},
    {"Ambient", "righttoleft", "RightToLeft", Vb6Type::Integer,  HPF_NONE},
    {"Ambient", "usermode",    "UserMode",    Vb6Type::Boolean,  HPF_NONE},};

inline const HostPseudoRow* hostPseudoFind(const std::string& pseudoObj,
                                    const std::string& memberName) {
    if (memberName.empty()) return nullptr;
    const std::string obj = hostPseudoLower(pseudoObj);
    const std::string mem = hostPseudoLower(memberName);
    for (const HostPseudoRow& r : kHostPseudoRows) {
        if (obj == hostPseudoLower(r.obj) && mem == r.name) return &r;
    }
    return nullptr;
}


// 这个名字是不是表里的宿主伪对象 (今天四档)。账 #278 §B120: 这一句从前在发码侧抄了
// 五份 (cgen_with 两处 / cgen_assign_host_pseudo / cgen_util_type / cgen_expr_member_obj_dispatch)。
// 表里新增一档而哪一处没跟上，那一处就答"不是" —— 同一个名字两种答案，而答案决定的是
// "这条成员访问要不要走宿主伪成员那条路"，静默。
inline bool hostPseudoObjectKnown(const std::string& name) {
    const std::string obj = hostPseudoLower(name);
    if (obj.empty()) return false;
    for (const HostPseudoRow& r : kHostPseudoRows) {
        if (obj == hostPseudoLower(r.obj)) return true;
    }
    return false;
}

// 对象那一段在 RTL 符号里的拼法 (就是本表 obj 列)。装配只许经
// CCodeGen::hostPseudoRtlSymbol 那一处，这里只交出那一段。
inline bool hostPseudoObjectSymbol(const std::string& name, std::string& outPascal) {
    const std::string obj = hostPseudoLower(name);
    for (const HostPseudoRow& r : kHostPseudoRows) {
        if (obj == hostPseudoLower(r.obj)) { outPascal = r.obj; return true; }
    }
    return false;
}

// 裸写的成员名能不能落到宿主伪成员上 (HPF_BARE)。发码侧的
// CCodeGen::hostPseudoBareName 与语义层的放行判定都只问这一句。
inline bool hostPseudoBareEligible(const std::string& pseudoObj, const std::string& memberName) {
    const HostPseudoRow* r = hostPseudoFind(pseudoObj, memberName);
    return r != nullptr && (r->flags & HPF_BARE) != 0;
}

// 这一档是不是"由专用码头回答"（HPF_CHANNEL），是的话把符号名交出去。
// 两个消费点：语义层据此在**两个位**都放行（集合名当限定符用是 VB6 的常规写法），
// 发码层据此决定发哪个码头 —— 名字到码头的映射只在这里写一次。
inline bool hostPseudoChannel(const std::string& pseudoObj, const std::string& memberName,
                              std::string& outChannel) {
    const HostPseudoRow* r = hostPseudoFind(pseudoObj, memberName);
    if (!r || (r->flags & HPF_CHANNEL) == 0 || r->channel == nullptr) return false;
    outChannel = r->channel;
    return true;
}

} // namespace vb6c3
