#pragma once
// 画布动词的唯一一张表 —— 一行 = 一个 (动词, 接收者)，回答"这一档有没有出口、出口叫什么、由谁发码"
// (账 #232 的 ①-c)
//
// 为什么要单开一个头，而且住在 src/common 而不是 src/backend:
// 「裸写的 Cls / PSet / Circle」与「Me.Cls / Me.PSet」是同一件事的两种写法，而折叠动作必须发生在
// **语义层**（那一刻语句已成形、符号查得到、文档种类也知道），发码动作住在 **cgen 层**。同一个决定
// 抄在两层 = 两份答案：表加一档而语义层没跟上，产物照旧编得过，只是多一条噪声（账 #217/#219 那两轮
// 就是这么定的性）。本仓对"两层都要问的单一出口"已有定死做法 —— host_pseudo.hpp /
// float_literal.hpp / int_literal.hpp 都住这里。
//
// 这张表收的是**名字**与"**这一档由哪条码头发码**"，不收实参拼装 —— 后者的形状差异是真的
// (PSet 在 Form 上要多递 hwnd 与 step/hasXY，Printer 上少一枚；Line 在 Form 上是 12 参、
// 在 PictureBox 上是 7 参)，把拼装并成一处反而要说谎。所以每一行标一个 owner:
//   CANVAS_OWNER_METHOD —— 由 CCodeGen::controlCanvasMethod 回答（三条画布码头：表达式形、语句形、打标记）
//   CANVAS_OWNER_DRAW   —— 由 withm 的 Form/Printer 绘图段回答（按 VB6 语义把 Step/省略/颜色打平成实参）
// 两个问题各自只有一个答案来源：给一档加档 = 加一行，两条码头与语义层的折叠判据都不必改。
//
// 表里没有的 (动词, 接收者) = 这一档今天没有出口。别在调用点"顺手补一个"，那正是本账 ② 清掉的形状
// (cls / print 两档以前被三条码头各自硬编码、每条都只认 PictureBox ⇒ `Me.Cls` 落 COM 兜底，
// 编得过、跑得起、一笔不画)。

#include <cstdint>
#include <string>

namespace vb6c3 {

// 接收者种类。这里**不用** FrmControlType：那是 src/project 的 AST 侧类型，common 不许向上依赖。
// "AST 类型 -> 本表旗标"的那一份映射只住在一个地方（CCodeGen::controlCanvasMethod 与绘图码头各自
// 在门口折一次，折完只把旗标传进来）。
enum : uint32_t {
    CR_CANVAS_FORM = 1u << 0,   // 窗体自己的画布
    CR_CANVAS_PBOX = 1u << 1,   // PictureBox 的画布
    CR_CANVAS_PRINTER = 1u << 2,  // Printer 对象（不是窗口，状态住在 RTL 的全局里）
};

// 这一档由哪条码头发码。
enum : uint8_t {
    CANVAS_OWNER_METHOD = 1,
    CANVAS_OWNER_DRAW = 2,
};

struct CanvasVerbRow {
    const char* verb;     // 小写名（与 Symbol::toLower 同口径：ASCII 小写）
    uint32_t receiver;    // CR_CANVAS_* 之一（一行只登记一个接收者）
    uint8_t owner;        // CANVAS_OWNER_*
    const char* entry;    // C 侧出口名
};

// 表本体。加一档就加一行 —— 两条码头与语义层的折叠判据都问这里，不许再在任何一处写第二份动词名单。
// cls 在 Form / PictureBox 上是"清画布"(vb6_ControlCls -> vb6_Form_Cls 那一条权威)，
// 而 Printer.Cls 在 VB6 里是**结束文档**，语义不同，所以那一行挂 DRAW 且名字是 EndDoc。
// line 三档齐但签名不同：PBox 7 参归 METHOD，Form 12 参（带 Step/hasXY）与 Printer 11 参归 DRAW。
inline const CanvasVerbRow kCanvasVerbRows[] = {
    {"cls",    CR_CANVAS_FORM,    CANVAS_OWNER_METHOD, "vb6_ControlCls"},
    {"cls",    CR_CANVAS_PBOX,    CANVAS_OWNER_METHOD, "vb6_ControlCls"},
    {"cls",    CR_CANVAS_PRINTER, CANVAS_OWNER_DRAW,   "vb6_Printer_EndDoc"},
    // enddoc 与 cls 在 Printer 上是**同一个出口**（VB6 里 Printer.Cls = 结束文档）：
    // 两个动词各登记一行，而不是让第二个动词自己在成员那一路拼一遍名字 —— 那样表加一档它照样不知道。
    {"enddoc", CR_CANVAS_PRINTER, CANVAS_OWNER_DRAW,   "vb6_Printer_EndDoc"},
    {"print",  CR_CANVAS_FORM,    CANVAS_OWNER_METHOD, "vb6_ControlPrint"},
    {"print",  CR_CANVAS_PBOX,    CANVAS_OWNER_METHOD, "vb6_ControlPrint"},
    {"line",   CR_CANVAS_PBOX,    CANVAS_OWNER_METHOD, "vb6_ControlLine"},
    {"line",   CR_CANVAS_FORM,    CANVAS_OWNER_DRAW,   "vb6_Form_Line"},
    {"line",   CR_CANVAS_PRINTER, CANVAS_OWNER_DRAW,   "vb6_Printer_Line"},
    {"pset",   CR_CANVAS_FORM,    CANVAS_OWNER_DRAW,   "vb6_Form_PSet"},
    {"pset",   CR_CANVAS_PRINTER, CANVAS_OWNER_DRAW,   "vb6_Printer_PSet"},
    {"circle", CR_CANVAS_FORM,    CANVAS_OWNER_DRAW,   "vb6_Form_Circle"},
    {"circle", CR_CANVAS_PRINTER, CANVAS_OWNER_DRAW,   "vb6_Printer_Circle"},
    {"point",  CR_CANVAS_FORM,    CANVAS_OWNER_DRAW,   "vb6_Form_Point"},
    {"point",  CR_CANVAS_PRINTER, CANVAS_OWNER_DRAW,   "vb6_Printer_Point"},
};

// 名字归一：common 不许向上依赖 semantics，所以这里自己拼 ASCII 小写（与 Symbol::toLower 同结果）。
inline std::string canvasVerbLower(const std::string& s) {
    std::string o(s);
    for (char& ch : o) {
        if (ch >= 'A' && ch <= 'Z') ch = static_cast<char>(ch - 'A' + 'a');
    }
    return o;
}

// 通用一问：这一 (动词, 接收者) 由指定 owner 回答时，出口叫什么。没有 = 空串。
// 返回 const char* 而不是 std::string：表是 static 数据，热路径每次构造反而多一份分配。
inline const char* canvasVerbEntry(const std::string& verb, uint32_t receiver, uint8_t owner) {
    for (const CanvasVerbRow& r : kCanvasVerbRows) {
        if (verb == r.verb && receiver == r.receiver && owner == r.owner) return r.entry;
    }
    return "";
}

// 画布码头那一条（cls / print / line-PBox）。
inline const char* canvasMethodEntry(const std::string& verb, uint32_t receiver) {
    return canvasVerbEntry(verb, receiver, CANVAS_OWNER_METHOD);
}

// 绘图码头那一条（Step/省略/颜色打平成实参的那一族）。
inline const char* canvasDrawEntry(const std::string& verb, uint32_t receiver) {
    return canvasVerbEntry(verb, receiver, CANVAS_OWNER_DRAW);
}

// 语义层的折叠判据问这一问：裸写这个名字、接收者是窗体时，有没有一档已经有人发码。
// 只要表里任何一条 (名字, FORM) 行存在就可以折 —— 不问 owner，因为折叠之后两条码头都会各自再问一遍。
inline bool canvasVerbHasFormEntry(const std::string& verb) {
    for (const CanvasVerbRow& r : kCanvasVerbRows) {
        if (verb == r.verb && r.receiver == CR_CANVAS_FORM) return true;
    }
    return false;
}

// parser 问的这一问：**这条动词的尾巴该由 parser 收进实参表**（账 #232 的 ③）。
// VB6 的画布语法里有两形不是"一个括号实参表"就能装下的：
//   PSet/Circle: `PSet (x, y)[, color]`、`Circle (x, y), radius[, ...]` —— 括号之后再跟逗号实参
//   Line:        `Line (x1, y1)-(x2, y2)[, color][, B|BF|F]` —— 紧跟第二点，再跟逗号实参
// 不收就会漏到外层表达式：`, 22` 变成一条独立的语句残桩，`-(20, 20)` 里的逗号当场
// `expected ')'`（`b499_probe3` 实测三条错误 VB2001/VB2003/VB2002）。
// 为什么这份知识住在这里而不是 parser 里：parser 那段吸收以前**自己写死** `circle` / `pset` / `line`
// 三个名字（是画布动词名单的第四份副本），而且只认带接收者那一形 ⇒ 裸写 `Circle (20,21), 22` 的尾巴
// 收不到。名单与"带不带接收者"是两件事，名字仍然只在这一个头里回答。
enum CanvasTailKind {
    CANVAS_TAIL_NONE = 0,        // 没有尾巴可收
    CANVAS_TAIL_COORD = 1,       // PSet / Circle: 括号之后再收逗号实参
    CANVAS_TAIL_TWO_POINT = 2,   // Line: 先收 `-(x2, y2)` 那枚第二点，再收逗号实参
};

inline CanvasTailKind canvasVerbTailKind(const std::string& verb) {
    if (verb == "pset" || verb == "circle") return CANVAS_TAIL_COORD;
    if (verb == "line") return CANVAS_TAIL_TWO_POINT;
    return CANVAS_TAIL_NONE;
}

} // namespace vb6c3
