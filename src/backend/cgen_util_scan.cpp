#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_util_scan.cpp: AST 扫描辅助 (GoSub / 多维数组 / Resume / OnError) ---

// ============================================================
// 第一百零一刀 (§11.79): 晚绑定成员读的判据收成一处
//   第九十八/九十九/一百刀在 cgen_assign.cpp 里写过一遍（那三刀只用于赋值站），条件位
//   这一刀要问同一问 —— 不再抄第二份，改成成员函数，两边同读。
//   判据：接收者按**链根**下钻（MemberAccess / IndexOrCall->callee），根是标识符、
//   是登记过的**变量**（排除同工程函数名 / 本类成员），不在控件数组与伪宿主名单里，
//   且推断型是 Object（或以 knownObjectVars_ 兵底）。
// ============================================================
bool CCodeGen::lateBoundComReadExpr(Expr* e) const {
    if (!e) return false;
    Expr* rcv = nullptr;
    if (e->kind == ASTNodeKind::MemberAccessExpr) {
        rcv = static_cast<MemberAccessExpr&>(*e).object.get();
    } else if (e->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& c = static_cast<IndexOrCallExpr&>(*e);
        if (c.callee && c.callee->kind == ASTNodeKind::MemberAccessExpr)
            rcv = static_cast<MemberAccessExpr&>(*c.callee).object.get();
    }
    if (!rcv) return false;
    Expr* root = rcv;
    while (root) {
        if (root->kind == ASTNodeKind::MemberAccessExpr)
            root = static_cast<MemberAccessExpr&>(*root).object.get();
        else if (root->kind == ASTNodeKind::IndexOrCallExpr
                 && static_cast<IndexOrCallExpr&>(*root).callee)
            root = static_cast<IndexOrCallExpr&>(*root).callee.get();
        else
            break;
    }
    if (!root || root->kind != ASTNodeKind::IdentifierExpr) return false;
    const std::string rn = Symbol::toLower(static_cast<IdentifierExpr&>(*root).name);
    static const char* kHost[] = { "printer", "printers", "screen", "app",
                                  "clipboard", "me", nullptr };
    if (knownControlArrays_.count(rn)) return false;
    for (int i = 0; kHost[i]; i++)
        if (rn == kHost[i]) return false;
    if (!knownLocalVars_.count(rn) && !knownObjectVars_.count(rn)) return false;
    if (inferExprType(*rcv) == Vb6Type::Object) return true;
    return knownObjectVars_.count(rn) > 0;
}

bool CCodeGen::exprHasLateBoundComRead(Expr* e) const {
    if (!e) return false;
    if (lateBoundComReadExpr(e)) return true;
    switch (e->kind) {
        case ASTNodeKind::BinaryExpr: {
            auto& b = static_cast<BinaryExpr&>(*e);
            return exprHasLateBoundComRead(b.left.get()) || exprHasLateBoundComRead(b.right.get());
        }
        case ASTNodeKind::UnaryExpr:
            return exprHasLateBoundComRead(static_cast<UnaryExpr&>(*e).operand.get());
        case ASTNodeKind::MemberAccessExpr:
            return exprHasLateBoundComRead(static_cast<MemberAccessExpr&>(*e).object.get());
        case ASTNodeKind::DictionaryAccessExpr:
            return exprHasLateBoundComRead(static_cast<DictionaryAccessExpr&>(*e).object.get());
        case ASTNodeKind::IndexOrCallExpr: {
            auto& c = static_cast<IndexOrCallExpr&>(*e);
            if (exprHasLateBoundComRead(c.callee.get())) return true;
            for (auto& a : c.positional)
                if (exprHasLateBoundComRead(a.get())) return true;
            for (auto& na : c.named)
                if (exprHasLateBoundComRead(na.value.get())) return true;
            return false;
        }
        case ASTNodeKind::NewExpr: {
            auto& nw = static_cast<NewExpr&>(*e);
            for (auto& a : nw.args)
                if (exprHasLateBoundComRead(a.get())) return true;
            return false;
        }
        case ASTNodeKind::TypeOfExpr:
            return exprHasLateBoundComRead(static_cast<TypeOfExpr&>(*e).object.get());
        default:
            return false;
    }
}


// ============================================================
// AST辅助: 检测语句列表中是否包含GoSubStmt
// ============================================================

bool CCodeGen::hasGoSubInStmts(StmtList& stmts) const {
    for (auto& stmt : stmts) {
        if (!stmt) continue;
        if (stmt->kind == ASTNodeKind::GoSubStmt) return true;
if (stmt->kind == ASTNodeKind::OnGoSubStmt) return true;  // P17.4
        // 递归检查复合语句
        if (stmt->kind == ASTNodeKind::IfStmt) {
            auto& ifStmt = static_cast<IfStmt&>(*stmt);
            if (hasGoSubInStmts(ifStmt.thenBody)) return true;
            if (hasGoSubInStmts(ifStmt.elseBody)) return true;
            for (auto& elif : ifStmt.elseIfs) {
                if (hasGoSubInStmts(elif->body)) return true;
            }
        } else if (stmt->kind == ASTNodeKind::ForStmt) {
            if (hasGoSubInStmts(static_cast<ForStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::ForEachStmt) {
            // Fix 043a: ForEachStmt was missing — GoSub inside For Each was not detected
            if (hasGoSubInStmts(static_cast<ForEachStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::DoLoopStmt) {
            if (hasGoSubInStmts(static_cast<DoLoopStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::WhileWendStmt) {
            if (hasGoSubInStmts(static_cast<WhileWendStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::SelectCaseStmt) {
            auto& sel = static_cast<SelectCaseStmt&>(*stmt);
            for (auto& c : sel.cases) {
                if (hasGoSubInStmts(c->body)) return true;
            }
            if (hasGoSubInStmts(sel.elseCase)) return true;
        }
    }
    return false;
}


// Bug #1 fix (082h): 预扫描语句列表, 收集UBound/LBound(arr, N>1)的数组名
// 到 knownNDArraysInProc_, 这样后续 UBound(arr, 1) 也能正确使用ND版本
void CCodeGen::scanNDArraysInExpr(Expr& expr) {
    if (expr.kind == ASTNodeKind::IndexOrCallExpr) {
        auto& call = static_cast<IndexOrCallExpr&>(expr);
        // Check if this is UBound/LBound with dimension > 1
        if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr) {
            std::string name = static_cast<IdentifierExpr&>(*call.callee).name;
            std::transform(name.begin(), name.end(), name.begin(), ::tolower);
            if ((name == "ubound" || name == "lbound") && call.positional.size() >= 2) {
                // Check if dimension arg > 1 (must be a literal integer)
                auto& dimArg = call.positional[1];
                if (dimArg && dimArg->kind == ASTNodeKind::LiteralExpr) {
                    auto& lit = static_cast<LiteralExpr&>(*dimArg);
                    if (lit.literalKind == LiteralKind::Integer && lit.intValue > 1) {
                        // Register the array name
                        auto& arrArg = call.positional[0];
                        if (arrArg && arrArg->kind == ASTNodeKind::IdentifierExpr) {
                            std::string arrName = static_cast<IdentifierExpr&>(*arrArg).name;
                            std::transform(arrName.begin(), arrName.end(), arrName.begin(), ::tolower);
                            knownNDArraysInProc_.insert(arrName);
                        }
                    }
                }
            }
        }
        // Also scan sub-expressions in call args and callee
        if (call.callee) scanNDArraysInExpr(*call.callee);
        for (auto& arg : call.positional) {
            if (arg) scanNDArraysInExpr(*arg);
        }
    } else if (expr.kind == ASTNodeKind::BinaryExpr) {
        auto& bin = static_cast<BinaryExpr&>(expr);
        if (bin.left) scanNDArraysInExpr(*bin.left);
        if (bin.right) scanNDArraysInExpr(*bin.right);
    } else if (expr.kind == ASTNodeKind::UnaryExpr) {
        auto& un = static_cast<UnaryExpr&>(expr);
        if (un.operand) scanNDArraysInExpr(*un.operand);
    } else if (expr.kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<MemberAccessExpr&>(expr);
        if (ma.object) scanNDArraysInExpr(*ma.object);
    } else if (expr.kind == ASTNodeKind::LiteralExpr) {
        // No sub-expressions
    } else if (expr.kind == ASTNodeKind::IdentifierExpr) {
        // No sub-expressions
    } else if (expr.kind == ASTNodeKind::WithMemberExpr) {
        // No sub-expressions to scan
    }
}


void CCodeGen::scanNDArraysInStmts(StmtList& stmts) {
    for (auto& stmt : stmts) {
        if (!stmt) continue;
        // Scan expressions in statements
        if (stmt->kind == ASTNodeKind::AssignmentStmt) {
            auto& assign = static_cast<AssignmentStmt&>(*stmt);
            if (assign.value) scanNDArraysInExpr(*assign.value);
        } else if (stmt->kind == ASTNodeKind::CallStmt) {
            auto& call = static_cast<CallStmt&>(*stmt);
            if (call.callee) scanNDArraysInExpr(*call.callee);
        } else if (stmt->kind == ASTNodeKind::ForStmt) {
            auto& forStmt = static_cast<ForStmt&>(*stmt);
            if (forStmt.start) scanNDArraysInExpr(*forStmt.start);
            if (forStmt.end) scanNDArraysInExpr(*forStmt.end);
            if (forStmt.step) scanNDArraysInExpr(*forStmt.step);
            scanNDArraysInStmts(forStmt.body);
        } else if (stmt->kind == ASTNodeKind::ForEachStmt) {
            auto& fe = static_cast<ForEachStmt&>(*stmt);
            if (fe.collection) scanNDArraysInExpr(*fe.collection);
            scanNDArraysInStmts(fe.body);
        } else if (stmt->kind == ASTNodeKind::DoLoopStmt) {
            auto& dl = static_cast<DoLoopStmt&>(*stmt);
            if (dl.condition) scanNDArraysInExpr(*dl.condition);
            scanNDArraysInStmts(dl.body);
        } else if (stmt->kind == ASTNodeKind::WhileWendStmt) {
            auto& ww = static_cast<WhileWendStmt&>(*stmt);
            if (ww.condition) scanNDArraysInExpr(*ww.condition);
            scanNDArraysInStmts(ww.body);
        } else if (stmt->kind == ASTNodeKind::IfStmt) {
            auto& ifStmt = static_cast<IfStmt&>(*stmt);
            if (ifStmt.condition) scanNDArraysInExpr(*ifStmt.condition);
            scanNDArraysInStmts(ifStmt.thenBody);
            scanNDArraysInStmts(ifStmt.elseBody);
            for (auto& elif : ifStmt.elseIfs) {
                if (elif && elif->condition) scanNDArraysInExpr(*elif->condition);
                if (elif) scanNDArraysInStmts(elif->body);
            }
        } else if (stmt->kind == ASTNodeKind::SelectCaseStmt) {
            auto& sel = static_cast<SelectCaseStmt&>(*stmt);
            if (sel.testExpr) scanNDArraysInExpr(*sel.testExpr);
            for (auto& c : sel.cases) {
                if (c) scanNDArraysInStmts(c->body);
            }
            scanNDArraysInStmts(sel.elseCase);
        } else if (stmt->kind == ASTNodeKind::SetStmt) {
            auto& set = static_cast<SetStmt&>(*stmt);
            if (set.value) scanNDArraysInExpr(*set.value);
        } else if (stmt->kind == ASTNodeKind::LetStmt) {
            auto& let = static_cast<LetStmt&>(*stmt);
            if (let.value) scanNDArraysInExpr(*let.value);
        } else if (stmt->kind == ASTNodeKind::WithStmt) {
            auto& with = static_cast<WithStmt&>(*stmt);
            if (with.object) scanNDArraysInExpr(*with.object);
            scanNDArraysInStmts(with.body);
        } else if (stmt->kind == ASTNodeKind::Block) {
            auto& block = static_cast<Block&>(*stmt);
            scanNDArraysInStmts(block.stmts);
        }
    }
}


// P14.1.2: 检测语句列表中是否包含Resume/Resume Next语句
bool CCodeGen::hasResumeInStmts(StmtList& stmts) const {
    for (auto& stmt : stmts) {
        if (!stmt) continue;
        if (stmt->kind == ASTNodeKind::ResumeStmt) {
            auto& resume = static_cast<ResumeStmt&>(*stmt);
            if (resume.resumeKind == ResumeKind::ResumeHere || resume.resumeKind == ResumeKind::ResumeNext)
                return true;
        }
        // 递归检查复合语句
        if (stmt->kind == ASTNodeKind::IfStmt) {
            auto& ifStmt = static_cast<IfStmt&>(*stmt);
            if (hasResumeInStmts(ifStmt.thenBody)) return true;
            if (hasResumeInStmts(ifStmt.elseBody)) return true;
            for (auto& elif : ifStmt.elseIfs) {
                if (hasResumeInStmts(elif->body)) return true;
            }
        } else if (stmt->kind == ASTNodeKind::ForStmt) {
            if (hasResumeInStmts(static_cast<ForStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::ForEachStmt) {
            // Fix 043a: ForEachStmt was missing — Resume inside For Each was not detected
            if (hasResumeInStmts(static_cast<ForEachStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::DoLoopStmt) {
            if (hasResumeInStmts(static_cast<DoLoopStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::WhileWendStmt) {
            if (hasResumeInStmts(static_cast<WhileWendStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::SelectCaseStmt) {
            auto& sel = static_cast<SelectCaseStmt&>(*stmt);
            for (auto& c : sel.cases) {
                if (hasResumeInStmts(c->body)) return true;
            }
            if (hasResumeInStmts(sel.elseCase)) return true;
        }
    }
    return false;
}


// ============================================================
// P12.3: 检测语句列表中是否包含OnErrorStmt
// ============================================================

bool CCodeGen::hasOnErrorInStmts(StmtList& stmts) const {
    for (auto& stmt : stmts) {
        if (!stmt) continue;
        if (stmt->kind == ASTNodeKind::OnErrorStmt || stmt->kind == ASTNodeKind::ResumeStmt) return true;
        // 递归检查复合语句
        if (stmt->kind == ASTNodeKind::IfStmt) {
            auto& ifStmt = static_cast<IfStmt&>(*stmt);
            if (hasOnErrorInStmts(ifStmt.thenBody)) return true;
            if (hasOnErrorInStmts(ifStmt.elseBody)) return true;
            for (auto& elif : ifStmt.elseIfs) {
                if (hasOnErrorInStmts(elif->body)) return true;
            }
        } else if (stmt->kind == ASTNodeKind::ForStmt) {
            if (hasOnErrorInStmts(static_cast<ForStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::ForEachStmt) {
            // Fix 043a: ForEachStmt was missing — OnError inside For Each was not detected
            if (hasOnErrorInStmts(static_cast<ForEachStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::DoLoopStmt) {
            if (hasOnErrorInStmts(static_cast<DoLoopStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::WhileWendStmt) {
            if (hasOnErrorInStmts(static_cast<WhileWendStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::SelectCaseStmt) {
            auto& sel = static_cast<SelectCaseStmt&>(*stmt);
            for (auto& c : sel.cases) {
                if (hasOnErrorInStmts(c->body)) return true;
            }
            if (hasOnErrorInStmts(sel.elseCase)) return true;
        }
    }
    return false;
}
} // namespace vb6c3
