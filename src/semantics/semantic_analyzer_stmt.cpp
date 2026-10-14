#include "semantics/semantic_analyzer.hpp"
#include <algorithm>
#include <cctype>
#include <tuple>
#include <initializer_list>
#include "semantics/semantic_analyzer_internal.h"

namespace vb6c3 {

// --- semantic_analyzer_stmt.cpp: 语句 Visitor (仅 Pass2) ---


// ============================================================
// 语句 Visitor (仅Pass2)
// ============================================================

void SemanticAnalyzer::visit(Block& node) {
    analyzeStmtList(node.stmts);
}

void SemanticAnalyzer::analyzeStmtList(StmtList& stmts) {
    for (auto& stmt : stmts) {
        if (verbose_) std::cerr << "[Sem]     stmt: " << stmt->kindName() << std::endl;
        dispatchStmt(*stmt, *this);
    }
}

void SemanticAnalyzer::visit(AssignmentStmt& node) {
    if (verbose_) std::cerr << "[Sem]       assign target=" << (node.target ? node.target->kindName() : "NULL")
                             << " value=" << (node.value ? node.value->kindName() : "NULL") << std::endl;
    if (!node.value || !node.target) {
        // 解析器应保证 target/value 都非空; 若为空则跳过分析
        return;
    }
    // Delegate 赋值绑定: op = AddressOf Proc → 签名校验 + 打桩标记
    if (node.value->kind == ASTNodeKind::AddressOfExpr &&
        node.target->kind == ASTNodeKind::IdentifierExpr) {
        auto& ident = static_cast<IdentifierExpr&>(*node.target);
        auto* varSym = symTab_.lookup(ident.name);
        if (varSym && varSym->kind == SymbolKind::Variable) {
            bindDelegateAddressOf(varSym->variableTypeName, *node.value, node.loc);
        }
    }
    Vb6Type valueType = analyzeExpr(*node.value);
    Vb6Type targetType = analyzeExpr(*node.target);

    // 类型兼容性检查
    if (targetType != Vb6Type::Unknown && valueType != Vb6Type::Unknown) {
        checkAssignment(targetType, valueType, node.loc, "assignment");
    }
}

void SemanticAnalyzer::visit(SetStmt& node) {
    Vb6Type valueType = analyzeExpr(*node.value);
    Vb6Type targetType = analyzeExpr(*node.target);

    // Set要求对象类型
    if (valueType != Vb6Type::Object && valueType != Vb6Type::Variant &&
        valueType != Vb6Type::Unknown) {
        diag_.warn(DiagnosticID::SemTypeMismatch, node.loc,
            "Set语句需要对象类型, 实际为 " + std::string(TypeSystem::typeToString(valueType)));
    }
}

void SemanticAnalyzer::visit(LetStmt& node) {
    Vb6Type valueType = analyzeExpr(*node.value);
    Vb6Type targetType = analyzeExpr(*node.target);
    checkAssignment(targetType, valueType, node.loc, "Let");
}

void SemanticAnalyzer::visit(IfStmt& node) {
    // 条件应为Boolean或可转Boolean
    Vb6Type condType = analyzeExpr(*node.condition);
    if (condType != Vb6Type::Unknown && condType != Vb6Type::Boolean &&
        condType != Vb6Type::Variant && !TypeSystem::isNumeric(condType)) {
        diag_.warn(DiagnosticID::SemTypeMismatch, node.loc,
            "If条件应为Boolean, 实际为 " + std::string(TypeSystem::typeToString(condType)));
    }

    analyzeStmtList(node.thenBody);

    for (auto& elseif : node.elseIfs) {
        dispatchStmt(*elseif, *this);
    }

    analyzeStmtList(node.elseBody);
}

void SemanticAnalyzer::visit(ElseIfClause& node) {
    Vb6Type condType = analyzeExpr(*node.condition);
    analyzeStmtList(node.body);
}

void SemanticAnalyzer::visit(ForStmt& node) {
    // 循环变量
    if (optionExplicit_) {
        auto* sym = symTab_.lookup(node.varName);
        if (!sym) {
            diag_.warn(DiagnosticID::SemUndeclaredIdentifier, node.loc,
                "未声明的变量: '" + node.varName + "' (可能来自其他模块)");
        } else {
            sym->isReferenced = true;
        }
    } else {
        markReferenced(node.varName);
    }

    // 起始/终止/步长应为数值
    Vb6Type startType = analyzeExpr(*node.start);
    Vb6Type endType = analyzeExpr(*node.end);

    if (node.step) {
        Vb6Type stepType = analyzeExpr(*node.step);
    }

    // 块级作用域 (VB6不创建块作用域, 但可以在For内声明变量)
    analyzeStmtList(node.body);
}

void SemanticAnalyzer::visit(ForEachStmt& node) {
    markReferenced(node.varName);
    Vb6Type collType = analyzeExpr(*node.collection);
    analyzeStmtList(node.body);
}

void SemanticAnalyzer::visit(DoLoopStmt& node) {
    if (node.condition) {
        Vb6Type condType = analyzeExpr(*node.condition);
    }
    analyzeStmtList(node.body);
}

void SemanticAnalyzer::visit(WhileWendStmt& node) {
    Vb6Type condType = analyzeExpr(*node.condition);
    analyzeStmtList(node.body);
}

void SemanticAnalyzer::visit(SelectCaseStmt& node) {
    Vb6Type testType = analyzeExpr(*node.testExpr);
    for (auto& caseClause : node.cases) {
        dispatchStmt(*caseClause, *this);
    }
    analyzeStmtList(node.elseCase);
}

void SemanticAnalyzer::visit(CaseClause& node) {
    for (auto& cv : node.values) {
        if (cv.value) analyzeExpr(*cv.value);
        if (cv.toValue) analyzeExpr(*cv.toValue);
    }
    analyzeStmtList(node.body);
}

void SemanticAnalyzer::visit(WithStmt& node) {
    Vb6Type objType = analyzeExpr(*node.object);
    withStack_.push_back(objType);

    analyzeStmtList(node.body);

    withStack_.pop_back();
}

void SemanticAnalyzer::visit(GoToStmt& node) {
    // GoTo标签存在性检查 (Pass2简单记录, 后续验证)
}
void SemanticAnalyzer::visit(GoSubStmt& node) {
    // P12.5: GoSub标签记录（验证在过程结束时进行）
    if (pass_ == 2) {
        gosubTargetLabels_.push_back({node.labelName, node.loc});
    }
}
// P17.4: OnGoSub label validation
void SemanticAnalyzer::visit(OnGoSubStmt& node) {
    if (pass_ == 2) {
        for (auto& label : node.labels) {
            gosubTargetLabels_.push_back({label, node.loc});
        }
    }
}

void SemanticAnalyzer::visit(OnGoToStmt& node) {
    // P18-A: OnGoTo labels validated in GoSub validation pass (same as OnGoSub)
    if (pass_ == 2) {
        for (auto& label : node.labels) {
            gosubTargetLabels_.push_back({label, node.loc});
        }
    }
}

void SemanticAnalyzer::visit(MidStmt& node) {
    // P18-A: Mid$ statement — validate expressions
    if (node.target) analyzeExpr(*node.target);
    if (node.start)  analyzeExpr(*node.start);
    if (node.hasLength && node.length) analyzeExpr(*node.length);
    if (node.value)  analyzeExpr(*node.value);
}

void SemanticAnalyzer::visit(ReturnStmt& node) {
    // Return from GoSub - 无需特殊语义检查
    // 运行时由vb6_gosub_stack处理
}

void SemanticAnalyzer::visit(OnErrorStmt& node) {
    // On Error语句无需类型检查
}

void SemanticAnalyzer::visit(ResumeStmt& node) {
    // Resume语句无需类型检查
}

void SemanticAnalyzer::visit(ErrorStmt& node) {
    if (node.errorNumber) analyzeExpr(*node.errorNumber);
}

void SemanticAnalyzer::visit(ExitStmt& node) {
    // Exit语句无需类型检查
}

void SemanticAnalyzer::visit(CallStmt& node) {
    // 账 #232①: 窗体里裸写的画布动词先折成 Me.<动词>，再按普通成员访问分析。
    // 为什么必须在 analyzeExpr 之前: 那一路会把查不到的名字落成"隐式 Variant 局部"
    // (实测发出 vb6_VARIANT Cls = vb6_VariantEmpty(); Cls();)，折完之后那条隐式声明根本不会出现。
    foldBareCanvasVerb(node.callee);
    analyzeExpr(*node.callee);
}

void SemanticAnalyzer::visit(ReDimStmt& node) {
    markReferenced(node.varName);
    for (auto& dim : node.dimensions) {
        if (dim.lower) analyzeExpr(*dim.lower);
        if (dim.upper) analyzeExpr(*dim.upper);
    }
}

void SemanticAnalyzer::visit(LabelStmt& node) {
    // 注册标签
    if (pass_ == 2) {
        std::string lower = Symbol::toLower(node.labelName);
        bool found = false;
        for (auto& label : declaredLabels_) {
            if (Symbol::toLower(label) == lower) { found = true; break; }
        }
        if (!found) {
            declaredLabels_.push_back(node.labelName);
        }
    }
}

void SemanticAnalyzer::visit(OptionStmt& node) {
    if (node.optionKind == OptionKind::Explicit) {
        optionExplicit_ = true;
    }
}

void SemanticAnalyzer::visit(LocalDeclStmt& node) {
    // 过程体内的局部声明
    if (pass_ == 2 && node.decl) {
        switch (node.decl->kind) {
            case ASTNodeKind::VariableDecl: {
                auto& varDecl = static_cast<VariableDecl&>(*node.decl);
                auto sym = std::make_unique<Symbol>(
                    SymbolKind::Variable, varDecl.name,
                    resolveTypeOrDefault(varDecl.name, varDecl.asType.get()),
                    varDecl.loc, varDecl.access
                );
                sym->isStatic = varDecl.isStatic;
                sym->isArray = !varDecl.dimensions.empty() || varDecl.isDynamicArray;
                sym->dimCount = (int32_t)varDecl.dimensions.size();  // P8.1: 多维数组维度数
                // Delegate 变量 (As <Delegate名>): 记录声明类型名 — 委托直调分支与
                // 赋值绑定都以 variableTypeName 反查委托符号 (模块级路径由
                // registerVariable 无条件记录, 此处仅委托类型需要).
                if (varDecl.asType && varDecl.asType->kind == ASTNodeKind::SimpleTypeRef &&
                    lookupDelegateSym(static_cast<SimpleTypeRef*>(varDecl.asType.get())->name)) {
                    sym->variableTypeName =
                        static_cast<SimpleTypeRef*>(varDecl.asType.get())->name;
                }
                // ai/084a M1: 类类型局部变量也记录类型名 (原先只给委托记)。
                // 仅当类型名命中成员访问级别预计算表 (= 工程内类模块) 才记,
                // 供 visit(MemberAccessExpr) 的接收者解析; 其他消费方都有 kind 守卫。
                else if (varDecl.asType && varDecl.asType->kind == ASTNodeKind::SimpleTypeRef &&
                    memberAccess_ &&
                    memberAccess_->count(Symbol::toLower(
                        static_cast<SimpleTypeRef*>(varDecl.asType.get())->name))) {
                    sym->variableTypeName =
                        static_cast<SimpleTypeRef*>(varDecl.asType.get())->name;
                }
                // tB B08c: 局部变量的 `As <类型>` 原文一律留一份 (Protected 越权判定要认接收者
                // 类)。上一条的 variableTypeName 只在委托类型时写, 那是后端消费口径 —— 给局部
                // 变量补满会改发码, 所以这里另记一个只有语义层读的字段。
                if (varDecl.asType && varDecl.asType->kind == ASTNodeKind::SimpleTypeRef) {
                    sym->srcTypeName = static_cast<SimpleTypeRef*>(varDecl.asType.get())->name;
                }
                symTab_.define(std::move(sym));
                // Dim op As Operation = AddressOf Proc — 初始化器即绑定
                if (varDecl.initializer && varDecl.asType &&
                    varDecl.asType->kind == ASTNodeKind::SimpleTypeRef) {
                    bindDelegateAddressOf(
                        static_cast<SimpleTypeRef*>(varDecl.asType.get())->name,
                        *varDecl.initializer, varDecl.loc);
                }
                break;
            }
            case ASTNodeKind::ConstDecl: {
                auto& constDecl = static_cast<ConstDecl&>(*node.decl);
                // Fix 091d: 复用 registerConstant 的完整值推导 (字面量/一元负号 →
                // Integer/Long/Double/String/Boolean). 此前局部常量被简化为
                // Variant → 生成 `const vb6_VARIANT SW_SHOWNORMAL = 1;` → C2440
                // "int → const vb6_VARIANT" (cToolsSystem.c 11/13); 且 Variant 常量
                // 参与位运算时被包装 vb6_VariantToLong(<字面量>) → C2440
                // (cDialog.c 36 BIF_USENEWUI 宏). registerConstant 内部 define 到
                // 当前 (局部) 作用域, 语义正确.
                registerConstant(constDecl);
                break;
            }
            default:
                break;
        }
    }
}
} // namespace vb6c3
