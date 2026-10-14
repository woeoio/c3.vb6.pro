#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_loop.cpp: Do/While 循环语句生成 ---

void CCodeGen::visit(DoLoopStmt& node) {
    // P14.3.2: 嵌套循环栈
    std::string exitLabel = "vb6_loop_exit_" + std::to_string(labelCounter_++);
    loopStack_.push_back({ExitKind::Do, exitLabel});

    switch (node.loopKind) {
        case DoLoopKind::DoWhileLoop:
            // Fix <c3-menu3d> 2026-10-07: 预测型 Do While 的条件**必须每轮重发**。
            // 旧法 emitExpr 把条件操作数的物化语句 (如 `_vcmp_N = vb6_VariantFromValue(x)`)
            // 发在 while 之前 —— C 的 while 只重测 `vb6_VarCmpLongEq(&_vcmp_N, 0)` 这个
            // 调用, 物化语句一轮都不再跑 ⇒ 循环条件永远盯着**进循环前的快照**。
            // 3DMenu 的 `Do While Remote.KeyPress = 0: DoEvents: Loop` 就此永锁
            // (实测: 菜单落定后任何按键都无响应)。改为「条件标签 + 回跳」, 让条件
            // 全部语句每轮重执行。后测型 (Do...Loop While) 的物化本就落在 do 块内,
            // 一直是每轮重跑, 不动。
            if (node.condition) {
                std::string condLabel = "vb6_loop_cond_" + std::to_string(labelCounter_++);
                c_.emitLine(condLabel + ":;");
                emitExpr(*node.condition);
                if (isComMarker_) resolveComValue("Int");
                c_.emitLine("if (!(" + lastExpr_ + ")) goto " + exitLabel + ";");
                c_.indent();
                emitStmtList(node.body);
                c_.dedent();
                c_.emitLine("goto " + condLabel + ";");
            } else {
                c_.emitLine("while (1) {");
                c_.indent();
                emitStmtList(node.body);
                c_.dedent();
                c_.emitLine("}");
            }
            break;

        case DoLoopKind::DoUntilLoop:
            if (node.condition) {
                std::string condLabel = "vb6_loop_cond_" + std::to_string(labelCounter_++);
                c_.emitLine(condLabel + ":;");
                emitExpr(*node.condition);
                c_.emitLine("if (" + lastExpr_ + ") goto " + exitLabel + ";");
                c_.indent();
                emitStmtList(node.body);
                c_.dedent();
                c_.emitLine("goto " + condLabel + ";");
            } else {
                c_.emitLine("while (1) {");
                c_.indent();
                emitStmtList(node.body);
                c_.dedent();
                c_.emitLine("}");
            }
            break;

        case DoLoopKind::DoLoopWhile:
            c_.emitLine("do {");
            c_.indent();
            emitStmtList(node.body);
            c_.dedent();
            if (node.condition) {
                emitExpr(*node.condition);
                c_.emitLine("} while (" + lastExpr_ + ");");
            } else {
                c_.emitLine("} while (1);");
            }
            break;

        case DoLoopKind::DoLoopUntil:
            c_.emitLine("do {");
            c_.indent();
            emitStmtList(node.body);
            c_.dedent();
            if (node.condition) {
                emitExpr(*node.condition);
                c_.emitLine("} while (!(" + lastExpr_ + "));");
            } else {
                c_.emitLine("} while (1);");
            }
            break;

        case DoLoopKind::DoLoop:
            c_.emitLine("do {");
            c_.indent();
            emitStmtList(node.body);
            c_.dedent();
            c_.emitLine("} while (1);");
            break;
    }
    c_.emitLine(exitLabel + ":;  /* Exit Do target */");

    loopStack_.pop_back();
}

void CCodeGen::visit(WhileWendStmt& node) {
    // P14.3.2: 嵌套循环栈 (While...Wend 等同于 Do While...Loop)
    std::string exitLabel = "vb6_loop_exit_" + std::to_string(labelCounter_++);
    loopStack_.push_back({ExitKind::Do, exitLabel});

    // 同 DoWhileLoop 的 Fix <c3-menu3d>: While...Wend 条件每轮重发
    // (条件物化语句不许只跑一次)。
    std::string condLabel = "vb6_loop_cond_" + std::to_string(labelCounter_++);
    c_.emitLine(condLabel + ":;");
    emitExpr(*node.condition);
    c_.emitLine("if (!(" + lastExpr_ + ")) goto " + exitLabel + ";");
    c_.indent();
    emitStmtList(node.body);
    c_.dedent();
    c_.emitLine("goto " + condLabel + ";");
    c_.emitLine(exitLabel + ":;  /* Exit Do target */");

    loopStack_.pop_back();
}

} // namespace vb6c3
