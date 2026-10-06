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

enum : uint8_t { HPF_NONE = 0, HPF_BARE = 1 << 0, HPF_METHOD = 1 << 1 };

struct HostPseudoRow {
    const char* obj;     // 宿主伪对象名 (小写; 发成 vb6_<Obj>_<rtl>)
    const char* name;    // VB6 成员名 (小写)
    const char* rtl;     // RTL 侧拼写 (C 大小写敏感, 源码拼写一律归一到它)
    Vb6Type type;        // 值读类型; Unknown = 不由本表回答
    uint32_t flags;      // uint32 而非 uint8: 行里写 HPF_BARE | HPF_METHOD 会整型提升,
                         // 花括号初始化里收窄回 uint8_t 是 narrowing (MSVC 直接报错)
};

inline const HostPseudoRow kHostPseudoRows[] = {
    // ---- UserControl (.ctl) ----
    {"usercontrol", "asyncread",       "AsyncRead",       Vb6Type::Unknown,  HPF_BARE | HPF_METHOD},
    {"usercontrol", "autoredraw",      "AutoRedraw",      Vb6Type::Boolean,  HPF_NONE},
    {"usercontrol", "backcolor",       "BackColor",       Vb6Type::Long,     HPF_NONE},
    {"usercontrol", "cancelasyncread", "CancelAsyncRead", Vb6Type::Unknown,  HPF_METHOD},
    {"usercontrol", "cls",             "Cls",             Vb6Type::Unknown,  HPF_METHOD},
    {"usercontrol", "containerhwnd",   "ContainerHwnd",   Vb6Type::LongPtr,  HPF_BARE},
    {"usercontrol", "controls",        "Controls",        Vb6Type::Unknown,  HPF_BARE},
    {"usercontrol", "enabled",         "Enabled",         Vb6Type::Boolean,  HPF_BARE},
    {"usercontrol", "extender",        "Extender",        Vb6Type::Unknown,  HPF_NONE},
    {"usercontrol", "forecolor",       "ForeColor",       Vb6Type::Long,     HPF_NONE},
    {"usercontrol", "hdc",             "hDC",             Vb6Type::LongPtr,  HPF_BARE},
    {"usercontrol", "height",          "Height",          Vb6Type::Long,     HPF_NONE},
    {"usercontrol", "hwnd",            "hWnd",            Vb6Type::LongPtr,  HPF_BARE},
    {"usercontrol", "mouseicon",       "MouseIcon",       Vb6Type::Unknown,  HPF_NONE},
    {"usercontrol", "mousepointer",    "MousePointer",    Vb6Type::Long,     HPF_NONE},
    {"usercontrol", "oledrag",         "OLEDrag",         Vb6Type::Unknown,  HPF_METHOD},
    {"usercontrol", "oledropmode",     "OLEDropMode",     Vb6Type::Long,     HPF_NONE},
    {"usercontrol", "picture",         "Picture",         Vb6Type::Unknown,  HPF_NONE},
    {"usercontrol", "propertychanged", "PropertyChanged", Vb6Type::Unknown,  HPF_BARE | HPF_METHOD},
    {"usercontrol", "refresh",         "Refresh",         Vb6Type::Unknown,  HPF_METHOD},
    {"usercontrol", "righttoleft",     "RightToLeft",     Vb6Type::Integer,  HPF_NONE},
    {"usercontrol", "scaleheight",     "ScaleHeight",     Vb6Type::Long,     HPF_BARE},
    {"usercontrol", "scalemode",       "ScaleMode",       Vb6Type::Long,     HPF_BARE},
    {"usercontrol", "scalewidth",      "ScaleWidth",      Vb6Type::Long,     HPF_BARE},
    {"usercontrol", "scalex",          "ScaleX",          Vb6Type::Unknown,  HPF_BARE | HPF_METHOD},
    {"usercontrol", "scaley",          "ScaleY",          Vb6Type::Unknown,  HPF_BARE | HPF_METHOD},
    {"usercontrol", "size",            "Size",            Vb6Type::Unknown,  HPF_METHOD},
    {"usercontrol", "textheight",      "TextHeight",      Vb6Type::Unknown,  HPF_METHOD},
    {"usercontrol", "textwidth",       "TextWidth",       Vb6Type::Unknown,  HPF_METHOD},
    {"usercontrol", "width",           "Width",           Vb6Type::Long,     HPF_NONE},
    // 对象成员只登记用于**拼写规范化**; 值面一律不答 (各有专用通道: Controls 集合走
    // vb6_UC_Controls()、Parent 链由 Fix 133u 在 cgen_base.cpp 改写、Font 是
    // vb6_ComIface_Font*, 装箱比较本来就不该按标量发)。
    {"usercontrol", "ambient",         "Ambient",         Vb6Type::Unknown,  HPF_NONE},
    {"usercontrol", "font",            "Font",            Vb6Type::Unknown,  HPF_NONE},
    {"usercontrol", "parentcontrols",  "ParentControls",  Vb6Type::Unknown,  HPF_NONE},

    // ---- PropertyPage (.pag) ----
    {"propertypage", "changed",          "Changed",          Vb6Type::Boolean, HPF_BARE},
    {"propertypage", "hwnd",             "hWnd",             Vb6Type::LongPtr, HPF_BARE},
    {"propertypage", "scaleheight",      "ScaleHeight",      Vb6Type::Long,    HPF_BARE},
    {"propertypage", "scalemode",        "ScaleMode",        Vb6Type::Long,    HPF_BARE},
    {"propertypage", "selectedcontrols", "SelectedControls", Vb6Type::Unknown, HPF_BARE | HPF_METHOD},

    // ---- Extender (容器提供的扩展对象) ----
    {"extender", "align",           "Align",           Vb6Type::Long,    HPF_NONE},
    {"extender", "container",       "Container",       Vb6Type::Unknown, HPF_NONE},
    {"extender", "drag",            "Drag",            Vb6Type::Unknown, HPF_METHOD},
    {"extender", "dragicon",        "DragIcon",        Vb6Type::Unknown, HPF_NONE},
    {"extender", "dragmode",        "DragMode",        Vb6Type::Long,    HPF_NONE},
    {"extender", "height",          "Height",          Vb6Type::Long,    HPF_NONE},
    {"extender", "helpcontextid",   "HelpContextID",   Vb6Type::Long,    HPF_NONE},
    {"extender", "left",            "Left",            Vb6Type::Long,    HPF_NONE},
    {"extender", "setfocus",        "SetFocus",        Vb6Type::Unknown, HPF_METHOD},
    {"extender", "tag",             "Tag",             Vb6Type::String,  HPF_NONE},
    {"extender", "tooltiptext",     "ToolTipText",     Vb6Type::String,  HPF_NONE},
    {"extender", "top",             "Top",             Vb6Type::Long,    HPF_NONE},
    {"extender", "visible",         "Visible",         Vb6Type::Boolean, HPF_NONE},
    {"extender", "whatsthishelpid", "WhatsThisHelpID", Vb6Type::Long,    HPF_NONE},
    {"extender", "width",           "Width",           Vb6Type::Long,    HPF_NONE},
    {"extender", "zorder",          "ZOrder",          Vb6Type::Unknown, HPF_METHOD},

    // ---- Ambient (宿主环境) ----
    {"ambient", "backcolor",   "BackColor",   Vb6Type::Long,     HPF_NONE},
    {"ambient", "displayname", "DisplayName", Vb6Type::String,   HPF_NONE},
    {"ambient", "forecolor",   "ForeColor",   Vb6Type::Long,     HPF_NONE},
    {"ambient", "font",        "Font",        Vb6Type::Unknown,  HPF_NONE},
    {"ambient", "righttoleft", "RightToLeft", Vb6Type::Integer,  HPF_NONE},
    {"ambient", "usermode",    "UserMode",    Vb6Type::Boolean,  HPF_NONE},
};

inline const HostPseudoRow* hostPseudoFind(const std::string& pseudoObj,
                                    const std::string& memberName) {
    if (memberName.empty()) return nullptr;
    const std::string obj = hostPseudoLower(pseudoObj);
    const std::string mem = hostPseudoLower(memberName);
    for (const HostPseudoRow& r : kHostPseudoRows) {
        if (obj == r.obj && mem == r.name) return &r;
    }
    return nullptr;
}


// 裸写的成员名能不能落到宿主伪成员上 (HPF_BARE)。发码侧的
// CCodeGen::hostPseudoBareName 与语义层的放行判定都只问这一句。
inline bool hostPseudoBareEligible(const std::string& pseudoObj, const std::string& memberName) {
    const HostPseudoRow* r = hostPseudoFind(pseudoObj, memberName);
    return r != nullptr && (r->flags & HPF_BARE) != 0;
}

} // namespace vb6c3
