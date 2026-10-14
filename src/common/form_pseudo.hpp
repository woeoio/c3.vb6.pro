#pragma once
// 窗体裸名属性表的唯一一张表 (kFormPseudoRows) 的家。
//
// 起因（账 #278 §B104 / §B109）：`.frm` 里裸写 `ScaleWidth` / `WindowState` / `ScaleHeight` 是
// VB6 的合法写法（等价于 `Me.` 打头），发码侧**一直**答得对 ——
// `cgen_expr_ident_symbol.inc` 在窗体模块里拿 `getControlPropReadFn(FrmControlType::Form, name)`
// 查这张表，命中就发 `vb6_Get…(hwnd)`。可那张表住在 backend，语义层问不到 ⇒
// "未找到标识符"那一路只认得符号表与工程级名单，于是每条合法写法都多配一条 VB3001；
// 没有 Option Explicit 时更实 —— 落进"隐式 Variant 局部"那一格，把属性读挤掉。
//
// 与 host_pseudo.hpp（账 #159/#219）同一族：同一个事实拷两份就是两个权威，表里加一行而语义层
// 没跟上，产物照旧是对的 C、诊断多一条噪声。所以表搬到 common（与 float_literal.hpp /
// int_literal.hpp 同一家族：两层都能问的那些单一出口），backend 的窗体臂与通用段改问这里，
// 语义层用同一个出口判"这名字窗体答不答"。
//
// 口径：
//   * 名字一律小写；查表前自己拼小写（common 不得向上依赖 semantics，见 hostPseudoLower 那条）。
//   * `readFn` 就是发码要调的 RTL 读函数名，**不含**实参 —— 句柄表达式由调用方拼（那是 backend
//     的事，本表只管"这个名字窗体答什么"）。
//   * `excl` = 这一行对哪几类接收者**不**成立（今天写在 `getControlPropReadFn` 通用段里的那几条
//     类型例外）。窗体不属于任何一类，所以语义层用 `FPQ_FORM` 调它时 `excl` 一律不挡。
//   * `formOnly` = 只有窗体答（今天的"窗体臂"那些行）。其余类型若也答同名属性，走的是各自
//     类型 switch 臂，本表不掺和 —— 收进来就会把"窗体的 Caption"错分给一枚文本框。
//
// 刻意不收的成员：`Left` / `Top` / `Width` / `Height` 这四形虽然在表里（通用段本来就答），
// 但 VB6 里 `Left(s, n)` 是**内置函数**，窗体坐标要写 `Me.Left` —— 那条碰撞由 backend 的
// `builtinCallWins68`（账 #68）在调用上下文里裁决，本表不重复那个决定，也不因为名字进了表
// 就允许语义层把 `Left(...)` 当属性读（语义层的放行只发生在"符号表查不到"那一路，而内置函数
// `Left` 在符号表里查得到）。

#include <cctype>
#include <cstdint>
#include <string>

namespace vb6c3 {

// 接收者类别位。backend 把自己的 FrmControlType 映射到这几个位（common 不得向上依赖
// project/frm_parser.hpp 里的那枚枚举）。
enum : uint32_t {
    FPQ_NONE          = 0,
    FPQ_FORM          = 1u << 0,   // 窗体本身（`.frm` 模块里的 Me / 裸名）
    FPQ_COMMON_DIALOG = 1u << 1,   // 没有外观：FontName / FontSize 是对话框字段
    FPQ_SHAPE         = 1u << 2,   // BorderStyle 在这里是"画笔线型"，不是窗口边框样式
    FPQ_LINE          = 1u << 3    // 同上
};

struct FormPseudoRow {
    const char* name;      // VB6 属性名（小写）
    const char* readFn;    // 发码要调的 RTL 读函数名（实参由 backend 拼）
    uint32_t excl;         // 对这些类别**不**成立
    bool formOnly;         // 只有窗体答
};

// 顺序 = 今天 `getControlPropReadFn` 的顺序（通用段在前、窗体臂在后），逐行未改：
// 通用段的 25 行原样搬进来（excl 带出原来那三条类型例外），窗体臂的 9 行标 formOnly。
// 窗体臂里的 scalewidth / scaleheight 与通用段同名同答案，不重复列。
inline const FormPseudoRow* formPseudoRows(size_t& count) {
    static const FormPseudoRow kRows[] = {
        // ---- 通用段（所有带句柄的对象）----
        {"left",              "vb6_GetControlLeft",              FPQ_NONE, false},
        {"top",               "vb6_GetControlTop",               FPQ_NONE, false},
        {"width",             "vb6_GetControlWidth",             FPQ_NONE, false},
        {"height",            "vb6_GetControlHeight",            FPQ_NONE, false},
        {"hwnd",              "vb6_GetControlHwnd",              FPQ_NONE, false},
        {"fontname",          "vb6_GetControlFontName",          FPQ_COMMON_DIALOG, false},
        {"fontsize",          "vb6_GetControlFontSize",          FPQ_COMMON_DIALOG, false},
        {"fontpixelheight",   "vb6_ControlFontPixelHeight",      FPQ_COMMON_DIALOG, false},
        {"fontbold",          "vb6_GetControlFontBold",          FPQ_NONE, false},
        {"fontitalic",        "vb6_GetControlFontItalic",        FPQ_NONE, false},
        {"fontunderline",     "vb6_GetControlFontUnderline",     FPQ_NONE, false},
        {"fontstrikethrough", "vb6_GetControlFontStrikethrough", FPQ_NONE, false},
        {"forecolor",         "vb6_GetControlForeColor",         FPQ_NONE, false},
        {"backcolor",         "vb6_GetControlBackColor",         FPQ_NONE, false},
        {"alignment",         "vb6_GetAlignment",                FPQ_NONE, false},
        {"tabindex",          "vb6_GetTabIndex",                 FPQ_NONE, false},
        {"tabstop",           "vb6_GetTabStop",                  FPQ_NONE, false},
        {"causesvalidation",  "vb6_GetCausesValidation",         FPQ_NONE, false},
        {"tooltiptext",       "vb6_GetToolTipText",              FPQ_NONE, false},
        {"tag",               "vb6_GetControlTag",               FPQ_NONE, false},
        {"mousepointer",      "vb6_GetMousePointer",             FPQ_NONE, false},
        {"mouseicon",         "vb6_GetMouseIcon",                FPQ_NONE, false},
        {"borderstyle",       "vb6_GetBorderStyle",              FPQ_SHAPE | FPQ_LINE, false},
        {"scalewidth",        "vb6_GetScaleWidth",               FPQ_NONE, false},
        {"scaleheight",       "vb6_GetScaleHeight",              FPQ_NONE, false},
        // ---- 窗体臂（只有窗体答）----
        {"caption",           "vb6_GetControlText",              FPQ_NONE, true},
        {"visible",           "vb6_GetControlVisible",           FPQ_NONE, true},
        {"enabled",           "vb6_GetControlEnabled",           FPQ_NONE, true},
        {"windowstate",       "vb6_GetWindowState",              FPQ_NONE, true},
        {"scalemode",         "vb6_WindowScaleModeSelf",         FPQ_NONE, true},  // 账 #197: 与写侧成对
        {"currentx",          "vb6_GetCurrentX",                 FPQ_NONE, true},  // 笔位只有一份存储 (float)
        {"currenty",          "vb6_GetCurrentY",                 FPQ_NONE, true},
        {"drawwidth",         "vb6_Form_DrawGetWidth",           FPQ_NONE, true},
        {"hdc",               "vb6_GetControlHDC",               FPQ_NONE, true}   // 账 #196: Form.hDC 同一处出口
    };
    count = sizeof(kRows) / sizeof(kRows[0]);
    return kRows;
}

inline std::string formPseudoLower(const std::string& s) {
    std::string o(s);
    for (char& ch : o) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
    return o;
}

// 这一行对 `kindBits` 这类接收者成不成立。
inline bool formPseudoRowApplies(const FormPseudoRow& r, uint32_t kindBits) {
    if (r.excl & kindBits) return false;
    if (r.formOnly && !(kindBits & FPQ_FORM)) return false;
    return true;
}

// 唯一出口：这类接收者读这个属性发哪个函数；没登记就交空串（= 调用方走原来的路）。
inline std::string formPseudoReadFn(const std::string& name, uint32_t kindBits) {
    size_t n = 0;
    const FormPseudoRow* rows = formPseudoRows(n);
    const std::string lk = formPseudoLower(name);
    for (size_t i = 0; i < n; ++i) {
        if (lk == rows[i].name && formPseudoRowApplies(rows[i], kindBits)) return rows[i].readFn;
    }
    return std::string();
}

// 语义层问这一句：`.frm` 里裸写这个名字，窗体答不答？答 = 这不是"未声明的标识符"。
// 只问窗体那一档（`FPQ_FORM`），因为放行只发生在窗体模块的裸名位。
inline bool formPseudoIsBare(const std::string& name) {
    return !formPseudoReadFn(name, FPQ_FORM).empty();
}

} // namespace vb6c3
