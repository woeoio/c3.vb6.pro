#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_jumps.cpp: 跳转与错误处理语句生成 (GoTo/OnError/Resume/Exit/GoSub) ---

void CCodeGen::visit(GoToStmt& node) {
    // Fix 086: 与 LabelStmt 的方向副本后缀配对
    // Fix 090o: 后缀仅在「目标标签定义在某个被方向拆分的 For body 内」时使用
    // (该标签会被两份副本各定义一次, goto 需与所在副本的 _dN 定义配对)。
    // 目标若在 body 外 (函数级出口标签如 QH/EH, VB6: GoTo 跳出循环到过程尾)
    // 只定义一份, 加后缀会指向不存在的 vb6_label_X_dN → C2094 (cZipArchive 事件取消出口)。
    std::string targetLower = Symbol::toLower(node.labelName);
    bool targetInSplitBody = false;
    for (auto& splitLabels : forSplitLabelStack_) {
        if (splitLabels.count(targetLower)) {
            targetInSplitBody = true;
            break;
        }
    }
    std::string lblSuffix =
        (labelCopyIdx_ > 0 && targetInSplitBody) ? ("_d" + std::to_string(labelCopyIdx_ + 1)) : "";
    c_.emitLine("goto vb6_label_" + cIdent(node.labelName) + lblSuffix + ";");
}

void CCodeGen::visit(OnErrorStmt& node) {
    switch (node.errorKind) {
        case OnErrorKind::GoToLabel: {
            // On Error GoTo label
            // 生成 setjmp 保护点，错误发生时 longjmp 回来后跳转到对应标签
            // C代码:
            //   if (setjmp(vb6_error_jmp_buf) != 0) {
            //       goto vb6_label_ErrorHandler;
            //   }
            //   vb6_err_jmp_active = 1;
            c_.emitLine("if (setjmp(vb6_local_err_jmp) != 0) {");
            c_.indent();
            if (hasResume_) {
                c_.emitLine("vb6_err_in_handler = 0;");
            }
            c_.emitLine("goto vb6_label_" + cIdent(node.labelName) + ";");
            c_.dedent();
            c_.emitLine("}");
            c_.emitLine("vb6_error_jmp_ptr = &vb6_local_err_jmp;");
            c_.emitLine("vb6_err_jmp_active = 1;");
            c_.emitLine("vb6_error_jmp_set = 1;");
            c_.emitLine("vb6_err_resume_next = 0;");  // P14.1.2: 清除Resume Next模式
            // P14.1.2: 标记进入受保护区块
            if (hasResume_) {
                inProtectedBlock_ = true;
                currentErrorHandlerLabel_ = node.labelName;
                // resumePointCounter_不重置: 避免标签重定义
            }
            // 需要setjmp头文件
            needSetjmp_ = true;
            break;
        }
        case OnErrorKind::ResumeNext: {
            // On Error Resume Next
            c_.emitLine("vb6_err_resume_next = 1;");
            break;
        }
        case OnErrorKind::GoToZero:
            // On Error GoTo 0: 禁用错误处理
            c_.emitLine("vb6_err_resume_next = 0;");
            c_.emitLine("vb6_err_jmp_active = 0;");
            c_.emitLine("vb6_error_jmp_set = 0;");
            break;
    }
}



// P14.1.2: Resume语句 -- 在错误处理器中恢复执行
void CCodeGen::visit(ResumeStmt& node) {
    switch (node.resumeKind) {
        case ResumeKind::ResumeHere:
            c_.emitLine("vb6_err_dispatch = vb6_err_resume_point;");
            c_.emitLine("vb6_ErrClear();");
            c_.emitLine("goto vb6_err_dispatch_switch;");
            break;
        case ResumeKind::ResumeNext:
            c_.emitLine("vb6_err_dispatch = vb6_err_resume_next_point;");
            c_.emitLine("vb6_ErrClear();");
            c_.emitLine("goto vb6_err_dispatch_switch;");
            break;
        case ResumeKind::ResumeLabel:
            c_.emitLine("vb6_ErrClear();");
            c_.emitLine("goto vb6_label_" + cIdent(node.labelName) + ";");
            break;
    }
}

// P14.1.3: Error语句 -- 触发运行时错误
void CCodeGen::visit(ErrorStmt& node) {
    emitExpr(*node.errorNumber);
    std::string errNum = std::move(lastExpr_);
    c_.emitLine("vb6_RaiseError(" + errNum + ", NULL);");
}
void CCodeGen::visit(ExitStmt& node) {
    switch (node.exitKind) {
        case ExitKind::Do:
        case ExitKind::For: {
            // P14.3.2: 从循环栈查找匹配的跳出标签
            std::string targetLabel;
            for (int i = (int)loopStack_.size() - 1; i >= 0; --i) {
                if (loopStack_[i].kind == node.exitKind) {
                    targetLabel = loopStack_[i].exitLabel;
                    break;
                }
            }
            if (!targetLabel.empty()) {
                c_.emitLine("goto " + targetLabel + ";");
            } else {
                c_.emitLine("break;  /* fallback: no matching loop in stack */");
            }
            break;
        }
        case ExitKind::Sub:
            // Fix <vbeclipse>: 不能发裸 `return;` —— 那会越过过程尾部的
            // vb6_RestoreErrState()/ivref Release (错误处理状态泄漏 → 之后
            // vb6_ErrRaise 的 longjmp 打进已失效栈帧 → 野读崩)。统一走过程出口标签。
            c_.emitLine("goto vb6_proc_exit;");
            procExitLabelUsed_ = true;
            break;
        case ExitKind::Function:
            c_.emitLine("goto vb6_proc_exit;");
            procExitLabelUsed_ = true;
            break;
        case ExitKind::Property:
            // Fix 181: Property Get 有返回值, 出口处装回 vb6_ret_<name>;
            // Property Let/Set 无返回值。两者都在统一出口发 return (见 decl_proc/func/prop 尾部)。
            c_.emitLine("goto vb6_proc_exit;");
            procExitLabelUsed_ = true;
            break;
    }
}

void CCodeGen::visit(LabelStmt& node) {
    // Fix 086: For循环方向拆分会发射两份循环体, 第二份的标签加 _d2 后缀,
    // 与对应的 goto 配对, 避免同一函数内标签重定义 (C2045).
    std::string lblSuffix = (labelCopyIdx_ > 0) ? ("_d" + std::to_string(labelCopyIdx_ + 1)) : "";
    c_.emitLine("vb6_label_" + cIdent(node.labelName) + lblSuffix + ":;");
    // P14.1.2: 如果这是错误处理器标签，结束受保护区块
    if (inProtectedBlock_ && !currentErrorHandlerLabel_.empty() &&
        Symbol::toLower(node.labelName) == Symbol::toLower(currentErrorHandlerLabel_)) {
        inProtectedBlock_ = false;
    }
}

void CCodeGen::visit(GoSubStmt& node) {
    hasGoSub_ = true;
    // GoSub label: 压入返回地址 → goto label
    int retId = gosubReturnCounter_++;
    c_.emitLine("if (vb6_gosub_sp >= 32) { /* P17.4: GoSub stack overflow */ goto vb6_proc_exit; }");
    procExitLabelUsed_ = true;
c_.emitLine("vb6_gosub_stack[vb6_gosub_sp++] = " + std::to_string(retId) + ";");
    c_.emitLine("goto vb6_label_" + cIdent(node.labelName) + ";");
    c_.emitLine("vb6_gosub_ret_" + std::to_string(retId) + ":;");
}

void CCodeGen::visit(OnGoSubStmt& node) {
    hasGoSub_ = true;
    // P17.4: On x GoSub label1, label2, ... - computed GoSub
    int retId = gosubReturnCounter_++;
    // Stack overflow guard
    c_.emitLine("if (vb6_gosub_sp >= 32) { /* P17.4: GoSub stack overflow */ goto vb6_proc_exit; }");
    procExitLabelUsed_ = true;
    // Push return address
    c_.emitLine("vb6_gosub_stack[vb6_gosub_sp++] = " + std::to_string(retId) + ";");
    // Evaluate index and dispatch to selected label
    emitExpr(*node.index);
    std::string idxVar = std::move(lastExpr_);
    c_.emitLine("{");
    c_.indent();
    c_.emitLine("int _gosub_idx = " + idxVar + ";");
    c_.emitLine("if (_gosub_idx >= 1 && _gosub_idx <= " + std::to_string(node.labels.size()) + ") {");
    c_.indent();
    c_.emitLine("switch(_gosub_idx) {");
    c_.indent();
    for (size_t k = 0; k < node.labels.size(); k++) {
        c_.emitLine("case " + std::to_string(k + 1) + ": goto vb6_label_" + cIdent(node.labels[k]) + ";");
    }
    c_.dedent();
    c_.emitLine("}");
    c_.dedent();
    c_.emitLine("}");
    c_.dedent();
    c_.emitLine("}");
    // Return landing point (after On...GoSub)
    c_.emitLine("vb6_gosub_ret_" + std::to_string(retId) + ":;");
}

void CCodeGen::visit(OnGoToStmt& node) {
    // P18-A: On x GoTo label1, label2, ... - computed goto
    emitExpr(*node.index);
    std::string idxVar = std::move(lastExpr_);
    c_.emitLine("{");
    c_.indent();
    c_.emitLine("int _goto_idx = " + idxVar + ";");
    c_.emitLine("if (_goto_idx >= 1 && _goto_idx <= " + std::to_string(node.labels.size()) + ") {");
    c_.indent();
    c_.emitLine("switch(_goto_idx) {");
    c_.indent();
    for (size_t k = 0; k < node.labels.size(); k++) {
        c_.emitLine("case " + std::to_string(k + 1) + ": goto vb6_label_" + cIdent(node.labels[k]) + ";");
    }
    c_.dedent();
    c_.emitLine("}");
    c_.dedent();
    c_.emitLine("}");
    c_.dedent();
    c_.emitLine("}");
}


} // namespace vb6c3
