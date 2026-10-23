#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_stmt.cpp: 语句生成 ---

void CCodeGen::emitStmtList(StmtList& stmts, bool emitResumePoints) {
    for (auto& stmt : stmts) {
        if (!stmt) continue;
        bumpInferEpoch();   // ai/030 续: 语句边界推推断缓存 epoch (体内中途注册全落在语句边界)

        // P14.1.2: 在受保护区块中为每条语句生成resume点
        if (emitResumePoints && inProtectedBlock_ &&
            stmt->kind != ASTNodeKind::OnErrorStmt &&
            stmt->kind != ASTNodeKind::LabelStmt) {
            int pt = resumePointCounter_++;
            c_.emitLine("vb6_err_resume_point = " + std::to_string(pt) + ";");
            c_.emitLine("vb6_err_resume_next_point = " + std::to_string(pt + 1) + ";");
            c_.emitLine("vb6_resume_" + std::to_string(pt) + ":;");
            dispatchPoints_.push_back(pt);
        }

        // 账 #464 §B164: 记下**这一条语句**所在的那一层 —— 回收刀用它判「声明与收回同层」。
        // 语句走完要**还原**成外层的层号: 块语句（If/For/…）的**体**里也是语句,
        //   它们会把 stmtLevel_ 改深; 不还原的话, 块语句**之后**的那一条语句会拿
        //   「体内那一层」当自己的层号 ⇒ 回收刀在那儿整条失灵（宁可不收）。
        const int savedStmtLevel = stmtLevel_;
        stmtLevel_ = c_.indentLevel();
        // 账 #464 §B164: 记下**回收刀名单**的起点 —— 本条语句只认**自己**登记的那几枚。
        //   块语句的**体**里那些语句各自登记的归它们自己收; 本条语句的清理**不许**顺手
        //   把它们带走（早先写成 ownedTempsToFree_.clear() 就是这个坑：体内的清理先把
        //   **外层条件里**登记的那一枚抹掉 ⇒ 那一枚永远没人放 ⇒ 漏）。
        const size_t ownedMark = ownedTempsToFree_.size();
        switch (stmt->kind) {
            case ASTNodeKind::AssignmentStmt:  visit(static_cast<AssignmentStmt&>(*stmt)); break;
            case ASTNodeKind::SetStmt:         visit(static_cast<SetStmt&>(*stmt)); break;
            case ASTNodeKind::LetStmt:         visit(static_cast<LetStmt&>(*stmt)); break;
            case ASTNodeKind::IfStmt:          visit(static_cast<IfStmt&>(*stmt)); break;
            case ASTNodeKind::ForStmt:         visit(static_cast<ForStmt&>(*stmt)); break;
            case ASTNodeKind::ForEachStmt:     visit(static_cast<ForEachStmt&>(*stmt)); break;
            case ASTNodeKind::DoLoopStmt:      visit(static_cast<DoLoopStmt&>(*stmt)); break;
            case ASTNodeKind::WhileWendStmt:   visit(static_cast<WhileWendStmt&>(*stmt)); break;
            case ASTNodeKind::SelectCaseStmt:  visit(static_cast<SelectCaseStmt&>(*stmt)); break;
            case ASTNodeKind::WithStmt:        visit(static_cast<WithStmt&>(*stmt)); break;
            case ASTNodeKind::GoToStmt:        visit(static_cast<GoToStmt&>(*stmt)); break;
                        case ASTNodeKind::OnGoSubStmt:
                visit(static_cast<OnGoSubStmt&>(*stmt));
                break;
            case ASTNodeKind::OnGoToStmt:
                visit(static_cast<OnGoToStmt&>(*stmt));
                break;
            case ASTNodeKind::MidStmt:
                visit(static_cast<MidStmt&>(*stmt));
                break;
case ASTNodeKind::GoSubStmt:       visit(static_cast<GoSubStmt&>(*stmt)); break;
            case ASTNodeKind::OnErrorStmt:     visit(static_cast<OnErrorStmt&>(*stmt)); break;
    case ASTNodeKind::ResumeStmt:      visit(static_cast<ResumeStmt&>(*stmt)); break;
    case ASTNodeKind::ErrorStmt:       visit(static_cast<ErrorStmt&>(*stmt)); break;
            case ASTNodeKind::ExitStmt:        visit(static_cast<ExitStmt&>(*stmt)); break;
            case ASTNodeKind::CallStmt:        visit(static_cast<CallStmt&>(*stmt)); break;
            case ASTNodeKind::ReDimStmt:       visit(static_cast<ReDimStmt&>(*stmt)); break;
            case ASTNodeKind::EraseStmt:       visit(static_cast<EraseStmt&>(*stmt)); break;
            case ASTNodeKind::LabelStmt:       visit(static_cast<LabelStmt&>(*stmt)); break;
            case ASTNodeKind::LocalDeclStmt:   visit(static_cast<LocalDeclStmt&>(*stmt)); break;
            // ai/vb-asm-extension-spec 项2/项3: 混排 —— Asm 块出现在语句流中间。
            // (体里只有这一个 AsmStmt 的过程由 tryEmitAsmProc 早退接管, 走不到这里。)
            case ASTNodeKind::AsmStmt:         emitAsmStmtMixed(static_cast<AsmStmt&>(*stmt)); break;
            case ASTNodeKind::Block:            visit(static_cast<Block&>(*stmt)); break;
            // 文件 I/O
            case ASTNodeKind::OpenStmt:        visit(static_cast<OpenStmt&>(*stmt)); break;
            case ASTNodeKind::CloseStmt:       visit(static_cast<CloseStmt&>(*stmt)); break;
            case ASTNodeKind::PrintStmt:       visit(static_cast<PrintStmt&>(*stmt)); break;
            case ASTNodeKind::WriteStmt:       visit(static_cast<WriteStmt&>(*stmt)); break;
            case ASTNodeKind::LineInputStmt:   visit(static_cast<LineInputStmt&>(*stmt)); break;
            case ASTNodeKind::InputStmt:       visit(static_cast<InputStmt&>(*stmt)); break;
            case ASTNodeKind::GetStmt:         visit(static_cast<GetStmt&>(*stmt)); break;
            case ASTNodeKind::PutStmt:         visit(static_cast<PutStmt&>(*stmt)); break;
            case ASTNodeKind::SeekStmt:        visit(static_cast<SeekStmt&>(*stmt)); break;
            case ASTNodeKind::LockStmt:        visit(static_cast<LockStmt&>(*stmt)); break;
            case ASTNodeKind::UnlockStmt:      visit(static_cast<UnlockStmt&>(*stmt)); break;
            case ASTNodeKind::ResetStmt:        visit(static_cast<ResetStmt&>(*stmt)); break;
            case ASTNodeKind::WidthStmt:       visit(static_cast<WidthStmt&>(*stmt)); break;
            case ASTNodeKind::KillStmt:        visit(static_cast<KillStmt&>(*stmt)); break;
            case ASTNodeKind::NameStmt:        visit(static_cast<NameStmt&>(*stmt)); break;
            case ASTNodeKind::MkDirStmt:       visit(static_cast<MkDirStmt&>(*stmt)); break;
            case ASTNodeKind::RmDirStmt:       visit(static_cast<RmDirStmt&>(*stmt)); break;
            case ASTNodeKind::ChDirStmt:       visit(static_cast<ChDirStmt&>(*stmt)); break;
            case ASTNodeKind::ChDriveStmt:     visit(static_cast<ChDriveStmt&>(*stmt)); break;
            case ASTNodeKind::FileCopyStmt:    visit(static_cast<FileCopyStmt&>(*stmt)); break;
            case ASTNodeKind::RaiseEventStmt:  visit(static_cast<RaiseEventStmt&>(*stmt)); break;
            case ASTNodeKind::BeepStmt:        visit(static_cast<BeepStmt&>(*stmt)); break;
            case ASTNodeKind::DoEventsStmt:    visit(static_cast<DoEventsStmt&>(*stmt)); break;
            case ASTNodeKind::EndStmt:
                c_.emitLine("vb6_End();");
                break;
            case ASTNodeKind::StopStmt:
                c_.emitLine("__debugbreak();");  // MSVC intrinsic
                break;
            case ASTNodeKind::ReturnStmt:
                if (hasGoSub_) {
                    // VB6 GoSub Return: 弹出返回地址并跳转
                    c_.emitLine("if (vb6_gosub_sp <= 0) { /* P17.4: GoSub stack underflow */ return; }");
c_.emitLine("switch(vb6_gosub_stack[--vb6_gosub_sp]) {");
                    c_.indent();
                    for (int i = 0; i < gosubReturnCounter_; i++) {
                        c_.emitLine("case " + std::to_string(i) + ": goto vb6_gosub_ret_" + std::to_string(i) + ";");
                    }
                    c_.dedent();
                    c_.emitLine("}");
                } else if (currentProc_ && currentProc_->kind == SymbolKind::Function) {
                    c_.emitLine("return " + currentReturnVar_ + ";");
                } else {
                    c_.emitLine("return;");
                }
                break;
            // 文件I/O和杂项语句暂不处理, 后续P4阶段
            default:
                c_.emitLine("/* unhandled stmt: " + std::string(stmt->kindName()) + " */");
                break;
        }

        // Fix 161b-decl-out: Declare A 版 ByVal String 的**出参回写**。
        // VB6 语义: ByVal String 是"可写缓冲", 被调用方写入的字节要按 NUL 截断
        // 回写成新 String (GetUserName/GetModuleFileName 靠这条工作)。C3 侧实参是
        // vb6_BSTR_ToANSI() 临时副本, 故在此把副本回写回原左值 —— 必须**先于**
        // FreeANSI (下面那段) 执行。目标为 BSTR 左值 → 用 vb6_BSTR_Assign(BSTR*, ...)。
        for (auto& op : ansiOutParams_) {
            c_.emitLine("vb6_BSTR_Assign(&(" + op.targetExpr + "), "
                        "vb6_BSTR_FromANSIBuf(" + op.tempVar + ", " + op.capExpr + "));");
        }
        ansiOutParams_.clear();

        // M22: 释放Declare ANSI函数的临时char*变量 (每条语句后统一清理, 无内存泄露)
        for (auto& ansiVar : ansiTempsToFree_) {
            c_.emitLine("vb6_FreeANSI(" + ansiVar + ");");
        }
        ansiTempsToFree_.clear();

        // 账 #305 §B217 (第一百二十一刀): 收回 ByRef Variant 实参那枚临时槽**装箱多取的**
        // 那一枚引用 (发码处在 cgen_expr_call_arg_emit.inc 登记)。声明发在这条语句之前、
        // 同一层花括号里 ⇒ 在这里清正好同作用域; 层号对不上说明声明在一对已经合上的花括号里
        // (Debug.Print 那族自带 `{ ... }`), 那时**宁可不收**也不许发成引用越界的 C。
        // Clear 空槽是 no-op, 被调方写过那一格也照 VB6 的规矩一并还掉。
        for (auto& bt : boxTempsToClear_) {
            if (bt.second == c_.indentLevel()) {
                c_.emitLine("vb6_VariantClear(&" + bt.first +
                            ");  /* 账 #305 §B217: 语句末收回装箱多取的那一枚引用 */");
            }
        }
        boxTempsToClear_.clear();

        // 账 #464 §B164 回收刀: 收回那些「一条语句里现造的带权临时量」。
        // 档 1 = BSTR ⇒ vb6_BSTR_Free; 档 2 = VARIANT ⇒ vb6_VariantClear(&t)。
        // 只收**本条语句自己登记**的那几枚（ownedMark 之后的那一段）——
        //   体内那些语句登记的归它们自己收, 外层条件里登记的归外层那条语句收。
        // 层号对不上 = 声明发在一对已经合上的花括号里 ⇒ 宁可不收（与 boxTempsToClear_ 同一口径）。
        // 放掉一次的责任只许一处: 这一枚**只**由这里放（拼接站自己放掉的那一枚不进名单）。
        for (size_t k = ownedMark; k < ownedTempsToFree_.size(); ++k) {
            const OwnedTemp& ot = ownedTempsToFree_[k];
            if (ot.level != stmtLevel_) continue;
            if (ot.kind == 2) {
                c_.emitLine("vb6_VariantClear(&" + ot.name +
                            ");  /* 账 #464 §B164: 语句末收回现造的那枚变体临时量 */");
            } else {
                c_.emitLine("vb6_BSTR_Free(" + ot.name +
                            ");  /* 账 #464 §B164: 语句末收回现造的那枚串临时量 */");
            }
        }
        ownedTempsToFree_.resize(ownedMark);
        stmtLevel_ = savedStmtLevel;
    }
}

void CCodeGen::visit(Block& node) {
    emitStmtList(node.stmts);
}

void CCodeGen::visit(IfStmt& node) {
    emitExpr(*node.condition);
    if (isComMarker_) resolveComValue("Int");  // If条件通常是Boolean/整数
    // Fix 090ag: If <Variant 值> 条件 — VB6 将 Variant 按真值判定 (Empty/0/False 为假).
    // 此前对 As Variant 参数/变量的裸标识符条件生成 if (JsonValue) → C2083 (vb6_VARIANT
    // 结构体比较非法). 需显式 vb6_VariantToBool.
    if (!isComMarker_) {
        bool condIsVariant = cExprIsVariant(lastExpr_);
        if (!condIsVariant && node.condition
            && node.condition->kind == ASTNodeKind::IdentifierExpr) {
            std::string cndLower =
                Symbol::toLower(static_cast<IdentifierExpr&>(*node.condition).name);
            if (knownVariantVars_.count(cndLower)) condIsVariant = true;
        }
        if (condIsVariant && lastExpr_.find("vb6_VariantToBool(") == std::string::npos) {
            lastExpr_ = "vb6_VariantToBool(" + lastExpr_ + ")";
        }
    }
    c_.emitLine("if (" + lastExpr_ + ") {");
    c_.indent();
    emitStmtList(node.thenBody);
    c_.dedent();

    // Fix 063b: ElseIf 条件可能生成 _vcmp_ 临时变量声明,
    // C 不允许在 } 和 else if 之间出现声明语句.
    // 改用 } else { if (...) { 模式, 将 _vcmp_ 声明放入 else 块内.
    int elseIfWrapCount = 0;
    for (auto& elseif : node.elseIfs) {
        c_.emitLine("} else {");
        // 账 #464 §B164 回收刀的门 ⑨: **块内条件**不许提名。
        //   这一枚条件是在 `} else {` **之后**求值的 ⇒ 那一枚临时量的声明落在 else 块**内**,
        //   而语句末的收回行落在 else 块**外** ⇒ 收回那一行引用越界（cl C2065）。
        //   门 ⑦ 的层号拦不住它: else 体与 if 体**同层**（生成器不在 else 里再缩进）。
        //   ⇒ 这一格宁可不收（与基线同漏, 不新增漏, 也不发越界的 C）。
        ++blockCondDepth_;
        emitExpr(*elseif->condition);
        --blockCondDepth_;
        c_.emitLine("if (" + lastExpr_ + ") {");
        c_.indent();
        emitStmtList(elseif->body);
        c_.dedent();
        elseIfWrapCount++;
    }

    if (!node.elseBody.empty()) {
        c_.emitLine("} else {");
        c_.indent();
        emitStmtList(node.elseBody);
        c_.dedent();
    }

    c_.emitLine("}");  // 关闭最内层 if/else
    for (int i = 0; i < elseIfWrapCount; i++) {
        c_.emitLine("}");  // 关闭每个 else 包装块
    }
}

void CCodeGen::visit(ElseIfClause& node) {
    // 由IfStmt内部处理
}


} // namespace vb6c3
