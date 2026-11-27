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
    // 第一百零八刀 (§11.87): 形参（`ByVal o As Object`）只进 knownObjectVars_，
    // 不进 knownLocalVars_ —— 只认后者的老判据就把「另一个过程里的 w = o.NoSuchProperty」
    // 放走了（真机 M4）。所以补一条 onlyObject 腿；但 knownObjectVars_ **全程不清空**，
    // 同名会漏到别的过程（`Thing` 是工程类时也被当 Object ⇒ b190 的第三条针变形 = 假阳），
    // 因此这一腿要求**接收者的推断型是 Object 或 Unknown**。已知型（工程类/串/数）一律不认。
    const bool isLocalVar = knownLocalVars_.count(rn) > 0;
    if (!isLocalVar && !knownObjectVars_.count(rn)) return false;
    const Vb6Type rcvType = inferExprType(*rcv);
    if (rcvType == Vb6Type::Object) return true;
    if (!isLocalVar) {
        // 接收者能解析出工程类名 ⇒ 它是字段访问，不是晚绑定（列表里那张表跨模块不清空，
        // 同名会漏进来；b190 的 `Thing->Alpha` 就是这么被误判的）。
        return inferClassTypeOfExpr(*rcv).empty();
    }
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

// 第一百九十二刀 (§11.175): 赋值站的「右值里有一次**必然爬出去**的 raise」。
//
// 与第一百九十一刀（实参站）的 exprCanRaiseStmt 不是同一问：那一站问「这枚实参可能 raise」，
// 宽着算不会错（实参本来就该在被调方之前求完）；赋值站问的是另一件事 ——「这次 raise 会不会
// 爬出**被调方**，从而中止调用方这条语句」。真机把它分成了两族（.temp/k192p/P218.bas）：
//
//   D1/D5/D7/D8/D11/D13  被调方自身无 On Error ⇒ raise 必然爬出 ⇒ 目标不动
//   D2/D10               被调方自带 On Error/Resume ⇒ 在内部就接了 ⇒ 目标该写 (x=5 / x=3)
//
// 两族在运行期状态上完全一样（vb6_ErrSerial 一视同仁），差别只在**被调方有没有处理器**，
// 所以判据取静态的：右值树里存在一枚直接调用，其被调名在本模块过程表里且那张表记着
// 「无处理器」。查不到名字的（内建、跨模块、晚绑定、下标）一律不计 —— 宽了会把 D2/D10
// 那一族误判成中止 (旧注释里那次 fdrawstate_x86 翻红就是这个)。
//
// 递归处理: D13 是「外层也无处理器、但它调的是有问题的内层」—— 那条语句的 raise 照样爬出
// 外层，所以这里按调用是否（传递地）落在无处理器过程上判，而不是只看被调方自己体内的形状。
bool CCodeGen::exprCallsUnprotectedProc(Expr* e) const {
    if (!e || procHasErrHandler_.empty()) return false;
    switch (e->kind) {
        case ASTNodeKind::IndexOrCallExpr: {
            auto& c = static_cast<IndexOrCallExpr&>(*e);
            if (c.callee && c.callee->kind == ASTNodeKind::IdentifierExpr) {
                const std::string n =
                    Symbol::toLower(static_cast<IdentifierExpr&>(*c.callee).name);
                auto it = procHasErrHandler_.find(n);
                if (it != procHasErrHandler_.end() && !it->second) return true;
            }
            for (auto& a : c.positional)
                if (exprCallsUnprotectedProc(a.get())) return true;
            for (auto& na : c.named)
                if (exprCallsUnprotectedProc(na.value.get())) return true;
            return false;
        }
        default:
            break;
    }
    // 其余节点: 用与 raiseInTree85 相同的子节点遍历（属性访问/成员读的接收者、二元两侧、
    // 一元子项、括号）。不在这里枚举全部 kind —— 只补常见的几族，其余按无风险处理。
    switch (e->kind) {
        case ASTNodeKind::MemberAccessExpr:
            return exprCallsUnprotectedProc(static_cast<MemberAccessExpr&>(*e).object.get());
        case ASTNodeKind::WithMemberExpr:
            // 隐式接收者 (.Prop): 接收者是 With 块那枚变量, 不在树里 ⇒ 无子表达式可扫。
            return false;
        case ASTNodeKind::BinaryExpr: {
            auto& b = static_cast<BinaryExpr&>(*e);
            return exprCallsUnprotectedProc(b.left.get())
                || exprCallsUnprotectedProc(b.right.get());
        }
        case ASTNodeKind::UnaryExpr:
            return exprCallsUnprotectedProc(static_cast<UnaryExpr&>(*e).operand.get());
        default:
            return false;
    }
}

// 第一百九十一刀 (§11.174): 「这枚表达式求值时可能 raise」—— 语句位置调用的实参守卫开闸判据。
// 为什么需要它: 第一百零二刀只给「实参含晚绑定 COM 读」开闸 (真机 I 族量的那一族),
// 而真机 A7/A8 (.temp/k191p/P217.bas) 量的口径与 I 族完全相同, 只是暴雷源不限于晚绑定读:
//   `Take2 FEscape(), LogTag(2)` → 坏实参之后不求值、之前的恰好一次、被调方不进入。
// 判据取**保守的宽**: 树里只要有一枚调用 (用户过程 / 内建 / 方法 / 数组下标) 或一枚可能
// 除零的算符, 就当可能 raise。提升临时量本身无副作用 (ByRef 那一形已在提升处排除),
// 宽一点只多一枚临时量, 窄一点就是错的读数。
bool CCodeGen::exprCanRaiseStmt(Expr* e) const {
    if (!e) return false;
    switch (e->kind) {
        case ASTNodeKind::IndexOrCallExpr: {
            auto& c = static_cast<IndexOrCallExpr&>(*e);
            if (exprCanRaiseStmt(c.callee.get())) return true;
            for (auto& a : c.positional)
                if (exprCanRaiseStmt(a.get())) return true;
            for (auto& na : c.named)
                if (exprCanRaiseStmt(na.value.get())) return true;
            return true;   // 调用/下标本身就可能 raise (过程体内的 raise, 下标越界 9)
        }
        case ASTNodeKind::BinaryExpr: {
            auto& b = static_cast<BinaryExpr&>(*e);
            if (b.op == BinaryOp::Div || b.op == BinaryOp::IntDiv
                || b.op == BinaryOp::Mod || b.op == BinaryOp::Pow)
                return true;
            return exprCanRaiseStmt(b.left.get()) || exprCanRaiseStmt(b.right.get());
        }
        case ASTNodeKind::UnaryExpr:
            return exprCanRaiseStmt(static_cast<UnaryExpr&>(*e).operand.get());
        case ASTNodeKind::MemberAccessExpr:
            return exprCanRaiseStmt(static_cast<MemberAccessExpr&>(*e).object.get());
        case ASTNodeKind::DictionaryAccessExpr:
            return exprCanRaiseStmt(static_cast<DictionaryAccessExpr&>(*e).object.get());
        case ASTNodeKind::NewExpr: {
            auto& nw = static_cast<NewExpr&>(*e);
            for (auto& a : nw.args)
                if (exprCanRaiseStmt(a.get())) return true;
            return true;   // New 本身会 raise (91 对象变量未设置 / 429)
        }
        case ASTNodeKind::TypeOfExpr:
            return exprCanRaiseStmt(static_cast<TypeOfExpr&>(*e).object.get());
        default:
            return false;
    }
}

// 第二百三十一刀 (§11.214): 变体目标赋值站的「右值可能 raise」。
// 真机 .temp/k225p/P325 十格全是「目标不动」：W1/W10 `w = CStr(Abs(v))`/94、W3 `w = Abs(v)`/13、
// W4 `w = Int("")`/13、W5 `w = CLng(v)`/13、W6 `w = 1 / d`/11、W7 `w = v + 1`/13。
// 与 exprCanRaiseStmt 只差一条：**加减乘也箿进**。理由是这两族走的是同一条路 ——
// `v + 1` 里的 v 是整只槽，运算前要过数值提取（vb6_VariantToLong），那就和 CLng 一样会 raise。
// 多算的代价只多一枚临时量（提升本身无副作用），少算的代价是「号对了、值被清」。
bool CCodeGen::exprCanRaiseVariantStore(Expr* e) const {
    if (!e) return false;
    switch (e->kind) {
        case ASTNodeKind::BinaryExpr: {
            auto& b = static_cast<BinaryExpr&>(*e);
            if (b.op == BinaryOp::Add || b.op == BinaryOp::Sub
                || b.op == BinaryOp::Mul || b.op == BinaryOp::Div
                || b.op == BinaryOp::IntDiv || b.op == BinaryOp::Mod
                || b.op == BinaryOp::Pow)
                return true;
            return exprCanRaiseVariantStore(b.left.get())
                || exprCanRaiseVariantStore(b.right.get());
        }
        case ASTNodeKind::UnaryExpr:
            return exprCanRaiseVariantStore(static_cast<UnaryExpr&>(*e).operand.get());
        case ASTNodeKind::IndexOrCallExpr: {
            auto& c = static_cast<IndexOrCallExpr&>(*e);
            if (exprCanRaiseVariantStore(c.callee.get())) return true;
            for (auto& a : c.positional)
                if (exprCanRaiseVariantStore(a.get())) return true;
            for (auto& na : c.named)
                if (exprCanRaiseVariantStore(na.value.get())) return true;
            return true;   // 调用/下标本身就可能 raise
        }
        case ASTNodeKind::MemberAccessExpr:
            return exprCanRaiseVariantStore(static_cast<MemberAccessExpr&>(*e).object.get());
        case ASTNodeKind::DictionaryAccessExpr:
            return exprCanRaiseVariantStore(static_cast<DictionaryAccessExpr&>(*e).object.get());
        case ASTNodeKind::NewExpr: {
            auto& nw = static_cast<NewExpr&>(*e);
            for (auto& a : nw.args)
                if (exprCanRaiseVariantStore(a.get())) return true;
            return true;
        }
        case ASTNodeKind::TypeOfExpr:
            return exprCanRaiseVariantStore(static_cast<TypeOfExpr&>(*e).object.get());
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

// 账 #305 §11.119 (第一百三十九刀): 数过程体里 With 语句的条数 (递归) —— 活载体名册的槽数上界。
// 为什么要上界而不是精确值: 名册的槽号是**发码顺序**分配的, 而「哪一帧是自有载体」要等发码
// 才算得出来 (`vb6_NewObject` / `vb6_CreateObject` / 工程函数那一形)。多给的槽没登记就是 0,
// 落点放手时空判据天然跳过 —— 少给的槽则会溢出。
int CCodeGen::countWithStmts(StmtList& stmts) const {
    int n = 0;
    for (auto& stmt : stmts) {
        if (!stmt) continue;
        switch (stmt->kind) {
            case ASTNodeKind::WithStmt:
                n += 1 + countWithStmts(static_cast<WithStmt&>(*stmt).body);
                break;
            case ASTNodeKind::IfStmt: {
                auto& ifStmt = static_cast<IfStmt&>(*stmt);
                n += countWithStmts(ifStmt.thenBody);
                n += countWithStmts(ifStmt.elseBody);
                for (auto& elif : ifStmt.elseIfs) n += countWithStmts(elif->body);
                break;
            }
            case ASTNodeKind::ForStmt:
                n += countWithStmts(static_cast<ForStmt&>(*stmt).body);
                break;
            case ASTNodeKind::ForEachStmt:
                n += countWithStmts(static_cast<ForEachStmt&>(*stmt).body);
                break;
            case ASTNodeKind::DoLoopStmt:
                n += countWithStmts(static_cast<DoLoopStmt&>(*stmt).body);
                break;
            case ASTNodeKind::WhileWendStmt:
                n += countWithStmts(static_cast<WhileWendStmt&>(*stmt).body);
                break;
            case ASTNodeKind::SelectCaseStmt: {
                auto& sel = static_cast<SelectCaseStmt&>(*stmt);
                for (auto& c : sel.cases) n += countWithStmts(c->body);
                n += countWithStmts(sel.elseCase);
                break;
            }
            default: break;
        }
    }
    return n;
}

// 账 #305 §11.116 五 (第一百四十一刀): 扫过程体, 记下每个 `LabelStmt` 定义处所在的 With 层深。
// 为什么要预先扫: `GoTo` 发码时只知道「现在在第几层 With」, 要知道「目标标签在第几层」才能算出
// 该退几层 —— 而标签可能定义在 `GoTo` 之后 (向下的跳转最常见), 发码是单趟的。
// 同一标签在方向拆分的两份循环体里各定义一次, 两处层深相同 (同一段源码), 后写覆盖无妨。
void CCodeGen::collectLabelWithDepth(StmtList& stmts, int depth,
                                    std::unordered_map<std::string, int>& out) {
    for (auto& stmt : stmts) {
        if (!stmt) continue;
        switch (stmt->kind) {
            case ASTNodeKind::LabelStmt:
                out[Symbol::toLower(static_cast<LabelStmt&>(*stmt).labelName)] = depth;
                break;
            case ASTNodeKind::WithStmt:
                collectLabelWithDepth(static_cast<WithStmt&>(*stmt).body, depth + 1, out);
                break;
            case ASTNodeKind::IfStmt: {
                auto& ifStmt = static_cast<IfStmt&>(*stmt);
                collectLabelWithDepth(ifStmt.thenBody, depth, out);
                collectLabelWithDepth(ifStmt.elseBody, depth, out);
                for (auto& elif : ifStmt.elseIfs) collectLabelWithDepth(elif->body, depth, out);
                break;
            }
            case ASTNodeKind::ForStmt:
                collectLabelWithDepth(static_cast<ForStmt&>(*stmt).body, depth, out);
                break;
            case ASTNodeKind::ForEachStmt:
                collectLabelWithDepth(static_cast<ForEachStmt&>(*stmt).body, depth, out);
                break;
            case ASTNodeKind::DoLoopStmt:
                collectLabelWithDepth(static_cast<DoLoopStmt&>(*stmt).body, depth, out);
                break;
            case ASTNodeKind::WhileWendStmt:
                collectLabelWithDepth(static_cast<WhileWendStmt&>(*stmt).body, depth, out);
                break;
            case ASTNodeKind::SelectCaseStmt: {
                auto& sel = static_cast<SelectCaseStmt&>(*stmt);
                for (auto& c : sel.cases) collectLabelWithDepth(c->body, depth, out);
                collectLabelWithDepth(sel.elseCase, depth, out);
                break;
            }
            default: break;
        }
    }
}


// ============================================================
// P12.3: 检测语句列表中是否包含OnErrorStmt
// ============================================================

// 第一百九十二刀 (§11.175): 本模块过程表的「自带 On Error 与否」—— 赋值站的右值逃逸判据用它。
// 扫全部顶层声明（Sub/Function/Property）；表值 = 过程体内是否有 OnErrorStmt/ResumeStmt。
void CCodeGen::scanProcErrHandlers(Module& module) {
    procHasErrHandler_.clear();
    for (auto& decl : module.declarations) {
        if (!decl) continue;
        if (decl->kind == ASTNodeKind::SubDecl) {
            auto& d = static_cast<SubDecl&>(*decl);
            procHasErrHandler_[Symbol::toLower(d.name)] = hasOnErrorInStmts(d.body);
        } else if (decl->kind == ASTNodeKind::FunctionDecl) {
            auto& d = static_cast<FunctionDecl&>(*decl);
            procHasErrHandler_[Symbol::toLower(d.name)] = hasOnErrorInStmts(d.body);
        } else if (decl->kind == ASTNodeKind::PropertyDecl) {
            auto& d = static_cast<PropertyDecl&>(*decl);
            procHasErrHandler_[Symbol::toLower(d.name)] = hasOnErrorInStmts(d.body);
        }
    }
}

// 第一百九十三刀 (§11.176): 过程体里有没有 `On Error Resume Next`（只认这一种，
// 不把 GoTo 那族算进来 —— 它们已经是靠 setjmp 的）。递归走复合语句，口径同
// hasOnErrorInStmts。
bool CCodeGen::hasResumeNextInStmts(StmtList& stmts) const {
    for (auto& stmt : stmts) {
        if (!stmt) continue;
        if (stmt->kind == ASTNodeKind::OnErrorStmt) {
            if (static_cast<OnErrorStmt&>(*stmt).errorKind == OnErrorKind::ResumeNext) return true;
            continue;
        }
        if (stmt->kind == ASTNodeKind::IfStmt) {
            auto& s = static_cast<IfStmt&>(*stmt);
            if (hasResumeNextInStmts(s.thenBody)) return true;
            if (hasResumeNextInStmts(s.elseBody)) return true;
            for (auto& e : s.elseIfs) if (e && hasResumeNextInStmts(e->body)) return true;
        } else if (stmt->kind == ASTNodeKind::ForStmt) {
            if (hasResumeNextInStmts(static_cast<ForStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::ForEachStmt) {
            if (hasResumeNextInStmts(static_cast<ForEachStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::DoLoopStmt) {
            if (hasResumeNextInStmts(static_cast<DoLoopStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::WhileWendStmt) {
            if (hasResumeNextInStmts(static_cast<WhileWendStmt&>(*stmt).body)) return true;
        } else if (stmt->kind == ASTNodeKind::SelectCaseStmt) {
            auto& sel = static_cast<SelectCaseStmt&>(*stmt);
            for (auto& c : sel.cases) if (c && hasResumeNextInStmts(c->body)) return true;
            if (hasResumeNextInStmts(sel.elseCase)) return true;
        } else if (stmt->kind == ASTNodeKind::WithStmt) {
            if (hasResumeNextInStmts(static_cast<WithStmt&>(*stmt).body)) return true;
        }
    }
    return false;
}

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
