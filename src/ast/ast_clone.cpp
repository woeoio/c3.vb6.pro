// vb6c3 - AST 克隆 + 类型变量替换实现 (泛型单态化基建, G2/G3)
// 见 ast_clone.hpp 顶部设计注释. 失败协议: 未支持 kind → failed_=true +
// nullptr, 由泛型器 (driver_generics) 统一报 unsupported.

#include "ast/ast_clone.hpp"
#include <algorithm>
#include <cctype>

namespace vb6c3 {

std::string toLowerStr(const std::string& s) {
    std::string out = s;
    std::transform(out.begin(), out.end(), out.begin(),
                   [](unsigned char c) { return (char)std::tolower(c); });
    return out;
}

std::string ASTCloner::substName(const std::string& name) const {
    auto rn = idRename_.find(toLowerStr(name));
    if (rn != idRename_.end()) return rn->second;
    auto it = subst_.find(toLowerStr(name));
    if (it == subst_.end()) return name;
    if (auto* s = dynamic_cast<const SimpleTypeRef*>(it->second)) return s->name;
    return name;
}

// ============================================================
// TypeRef
// ============================================================

TypeRefPtr ASTCloner::cloneTypeRef(const ASTNode* ref) {
    if (!ref) return nullptr;
    switch (ref->kind) {
    case ASTNodeKind::SimpleTypeRef: {
        auto& s = static_cast<const SimpleTypeRef&>(*ref);
        auto it = subst_.find(toLowerStr(s.name));
        if (it != subst_.end()) return cloneTypeRef(it->second);
        return std::make_unique<SimpleTypeRef>(s.loc, s.name);
    }
    case ASTNodeKind::ArrayTypeRef: {
        auto& a = static_cast<const ArrayTypeRef&>(*ref);
        std::vector<ArrayTypeRef::Dimension> dims;
        for (auto& d : a.dimensions) {
            ArrayTypeRef::Dimension nd;
            nd.lower = d.lower ? cloneExpr(d.lower.get()) : nullptr;
            nd.upper = d.upper ? cloneExpr(d.upper.get()) : nullptr;
            dims.push_back(std::move(nd));
        }
        return std::make_unique<ArrayTypeRef>(a.loc, cloneTypeRef(a.elementType.get()),
                                              std::move(dims));
    }
    case ASTNodeKind::FixedStringTypeRef: {
        auto& f = static_cast<const FixedStringTypeRef&>(*ref);
        return std::make_unique<FixedStringTypeRef>(f.loc, cloneExpr(f.length.get()));
    }
    default:
        failed_ = true;
        return nullptr;
    }
}

// ============================================================
// Expr
// ============================================================

ExprPtr ASTCloner::cloneExpr(const Expr* e) {
    if (!e) return nullptr;
    auto out = cloneExprInner(e);
    if (!out) failed_ = true;
    return out;
}

ExprPtr ASTCloner::cloneExprInner(const Expr* e) {
    if (!e) return nullptr;
    switch (e->kind) {
    case ASTNodeKind::LiteralExpr: {
        auto& l = static_cast<const LiteralExpr&>(*e);
        auto out = std::make_unique<LiteralExpr>(l.loc, l.literalKind, l.rawText);
        switch (l.literalKind) {
        case LiteralKind::Boolean:  out->boolValue   = l.boolValue;   break;
        case LiteralKind::Integer:
        case LiteralKind::LongPtr:
        case LiteralKind::Currency:
        case LiteralKind::Decimal:
        case LiteralKind::Long:     out->longValue   = l.longValue;   break;
        case LiteralKind::Single:   out->floatValue  = l.floatValue;  break;
        case LiteralKind::Double:   out->doubleValue = l.doubleValue; break;
        default: break;  // String/Date/Nothing/Empty/Null: rawText 即全部
        }
        return out;
    }
    case ASTNodeKind::IdentifierExpr: {
        auto& id = static_cast<const IdentifierExpr&>(*e);
        // 只改"旧过程名→扁名" (self-assign/self-read); 刻意不查 subst_ —
        // 变量名恰与类型参数同名 (Dim t) 时不得被改名成类型名.
        auto rn = idRename_.find(toLowerStr(id.name));
        return std::make_unique<IdentifierExpr>(
            id.loc, rn != idRename_.end() ? rn->second : id.name, id.bracketed);
    }
    case ASTNodeKind::BinaryExpr: {
        auto& b = static_cast<const BinaryExpr&>(*e);
        auto out = std::make_unique<BinaryExpr>(b.loc, b.op,
            cloneExprInner(b.left.get()), cloneExprInner(b.right.get()));
        if ((b.left && !out->left) || (b.right && !out->right)) return nullptr;
        return out;
    }
    case ASTNodeKind::UnaryExpr: {
        auto& u = static_cast<const UnaryExpr&>(*e);
        auto operand = cloneExprInner(u.operand.get());
        if (u.operand && !operand) return nullptr;
        return std::make_unique<UnaryExpr>(u.loc, u.op, std::move(operand));
    }
    case ASTNodeKind::MemberAccessExpr: {
        auto& m = static_cast<const MemberAccessExpr&>(*e);
        auto obj = cloneExprInner(m.object.get());
        if (m.object && !obj) return nullptr;
        return std::make_unique<MemberAccessExpr>(m.loc, std::move(obj), m.memberName);
    }
    case ASTNodeKind::DictionaryAccessExpr: {
        auto& m = static_cast<const DictionaryAccessExpr&>(*e);
        auto obj = cloneExprInner(m.object.get());
        if (m.object && !obj) return nullptr;
        return std::make_unique<DictionaryAccessExpr>(m.loc, std::move(obj), m.key);
    }
    case ASTNodeKind::IndexOrCallExpr: {
        auto& c = static_cast<const IndexOrCallExpr&>(*e);
        auto out = std::make_unique<IndexOrCallExpr>(c.loc, cloneExprInner(c.callee.get()));
        if (c.callee && !out->callee) return nullptr;
        for (auto& a : c.positional) {
            auto ca = cloneExprInner(a.get());
            if (!ca) return nullptr;
            out->positional.push_back(std::move(ca));
        }
        for (auto& na : c.named) {
            NamedArg cn;
            cn.name = na.name;
            cn.value = cloneExprInner(na.value.get());
            if (na.value && !cn.value) return nullptr;
            out->named.push_back(std::move(cn));
        }
        // 模板未经语义标注, 以下镜像字段恒为默认值; 拷贝仅为保真
        out->byvalOverrides = c.byvalOverrides;
        out->omittedArgs = c.omittedArgs;
        out->isDelegateCall = c.isDelegateCall;
        out->delegateTypeName = c.delegateTypeName;
        out->calleeOvlSuffix = c.calleeOvlSuffix;
        return out;
    }
    case ASTNodeKind::NewExpr: {
        auto& n = static_cast<const NewExpr&>(*e);
        auto out = std::make_unique<NewExpr>(n.loc, substName(n.className));
        // 084c: 带参构造的实参一并克隆 (泛型特化/模板克隆路径需要保真)
        for (const auto& a : n.args) {
            auto ca = cloneExprInner(a.get());
            if (a && !ca) return nullptr;
            out->args.push_back(std::move(ca));
        }
        return out;
    }
    case ASTNodeKind::TypeOfExpr: {
        auto& t = static_cast<const TypeOfExpr&>(*e);
        auto obj = cloneExprInner(t.object.get());
        if (t.object && !obj) return nullptr;
        return std::make_unique<TypeOfExpr>(t.loc, std::move(obj), substName(t.typeName));
    }
    case ASTNodeKind::AddressOfExpr: {
        auto& a = static_cast<const AddressOfExpr&>(*e);
        return std::make_unique<AddressOfExpr>(a.loc, a.funcName);
    }
    case ASTNodeKind::MeExpr:
        return std::make_unique<MeExpr>(e->loc);
    case ASTNodeKind::WithMemberExpr: {
        auto& w = static_cast<const WithMemberExpr&>(*e);
        return std::make_unique<WithMemberExpr>(w.loc, w.memberName);
    }
    default:
        return nullptr;
    }
}

// ============================================================
// Stmt
// ============================================================

std::vector<ExprPtr> ASTCloner::cloneExprVec(const std::vector<ExprPtr>& v) {
    std::vector<ExprPtr> out;
    for (auto& e : v) {
        auto c = cloneExprInner(e.get());
        if (e && !c) c = nullptr;  // 标记给上层判 failed
        out.push_back(std::move(c));
    }
    return out;
}

StmtList ASTCloner::cloneStmtList(const StmtList& list) {
    StmtList out;
    for (auto& s : list) {
        auto c = cloneStmt(s.get());
        if (!c) return out;  // failed 传播
        out.push_back(std::move(c));
    }
    return out;
}

StmtPtr ASTCloner::cloneStmt(const Stmt* s) {
    if (!s) return nullptr;
    StmtPtr out;
    switch (s->kind) {
    case ASTNodeKind::Block: {
        auto& x = static_cast<const Block&>(*s);
        out = std::make_unique<Block>(x.loc, cloneStmtList(x.stmts)); break;
    }
    case ASTNodeKind::AssignmentStmt: {
        auto& x = static_cast<const AssignmentStmt&>(*s);
        auto o = std::make_unique<AssignmentStmt>(x.loc, cloneExprInner(x.target.get()),
                                                  cloneExprInner(x.value.get()));
        if (x.target && !o->target) break;
        if (x.value && !o->value) break;
        o->isLSet = x.isLSet; o->isRSet = x.isRSet;
        out = std::move(o); break;
    }
    case ASTNodeKind::SetStmt: {
        auto& x = static_cast<const SetStmt&>(*s);
        auto o = std::make_unique<SetStmt>(x.loc, cloneExprInner(x.target.get()),
                                           cloneExprInner(x.value.get()));
        if (x.target && !o->target) break;
        if (x.value && !o->value) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::LetStmt: {
        auto& x = static_cast<const LetStmt&>(*s);
        auto o = std::make_unique<LetStmt>(x.loc, cloneExprInner(x.target.get()),
                                           cloneExprInner(x.value.get()));
        if (x.target && !o->target) break;
        if (x.value && !o->value) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::ElseIfClause: {
        auto& x = static_cast<const ElseIfClause&>(*s);
        auto o = std::make_unique<ElseIfClause>(x.loc, cloneExprInner(x.condition.get()),
                                                cloneStmtList(x.body));
        if (x.condition && !o->condition) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::IfStmt: {
        auto& x = static_cast<const IfStmt&>(*s);
        auto o = std::make_unique<IfStmt>(x.loc, cloneExprInner(x.condition.get()),
                                          cloneStmtList(x.thenBody),
                                          std::vector<std::unique_ptr<ElseIfClause>>{},
                                          cloneStmtList(x.elseBody), x.singleLine);
        if (x.condition && !o->condition) break;
        bool ok = true;
        for (auto& ei : x.elseIfs) {
            auto cl = cloneStmt(ei.get());
            if (!cl) { ok = false; break; }
            o->elseIfs.push_back(std::unique_ptr<ElseIfClause>(
                static_cast<ElseIfClause*>(cl.release())));
        }
        if (!ok) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::ForStmt: {
        auto& x = static_cast<const ForStmt&>(*s);
        auto o = std::make_unique<ForStmt>(x.loc, x.varName, cloneExprInner(x.start.get()),
                                           cloneExprInner(x.end.get()),
                                           cloneExprInner(x.step.get()),
                                           cloneStmtList(x.body));
        if ((x.start && !o->start) || (x.end && !o->end) || (x.step && !o->step)) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::ForEachStmt: {
        auto& x = static_cast<const ForEachStmt&>(*s);
        auto o = std::make_unique<ForEachStmt>(x.loc, x.varName,
                                               cloneExprInner(x.collection.get()),
                                               cloneStmtList(x.body));
        if (x.collection && !o->collection) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::DoLoopStmt: {
        auto& x = static_cast<const DoLoopStmt&>(*s);
        auto o = std::make_unique<DoLoopStmt>(x.loc, x.loopKind,
                                              cloneExprInner(x.condition.get()),
                                              cloneStmtList(x.body));
        if (x.condition && !o->condition) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::WhileWendStmt: {
        auto& x = static_cast<const WhileWendStmt&>(*s);
        auto o = std::make_unique<WhileWendStmt>(x.loc, cloneExprInner(x.condition.get()),
                                                 cloneStmtList(x.body));
        if (x.condition && !o->condition) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::CaseClause: {
        auto& x = static_cast<const CaseClause&>(*s);
        std::vector<CaseClause::CaseValue> vals;
        bool ok = true;
        for (auto& v : x.values) {
            CaseClause::CaseValue cv;
            cv.isIsClause = v.isIsClause;
            cv.relOp = v.relOp;
            cv.hasRelOp = v.hasRelOp;
            cv.value = cloneExprInner(v.value.get());
            if (v.value && !cv.value) { ok = false; break; }
            cv.toValue = cloneExprInner(v.toValue.get());
            if (v.toValue && !cv.toValue) { ok = false; break; }
            vals.push_back(std::move(cv));
        }
        if (!ok) break;
        out = std::make_unique<CaseClause>(x.loc, std::move(vals), cloneStmtList(x.body));
        break;
    }
    case ASTNodeKind::SelectCaseStmt: {
        auto& x = static_cast<const SelectCaseStmt&>(*s);
        auto o = std::make_unique<SelectCaseStmt>(x.loc, cloneExprInner(x.testExpr.get()),
                                                  std::vector<std::unique_ptr<CaseClause>>{},
                                                  cloneStmtList(x.elseCase));
        if (x.testExpr && !o->testExpr) break;
        bool ok = true;
        for (auto& c : x.cases) {
            auto cl = cloneStmt(c.get());
            if (!cl) { ok = false; break; }
            o->cases.push_back(std::unique_ptr<CaseClause>(
                static_cast<CaseClause*>(cl.release())));
        }
        if (!ok) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::WithStmt: {
        auto& x = static_cast<const WithStmt&>(*s);
        auto obj = cloneExprInner(x.object.get());
        if (x.object && !obj) break;
        out = std::make_unique<WithStmt>(x.loc, std::move(obj), cloneStmtList(x.body));
        break;
    }
    case ASTNodeKind::GoToStmt:
        out = std::make_unique<GoToStmt>(s->loc, static_cast<const GoToStmt&>(*s).labelName); break;
    case ASTNodeKind::GoSubStmt:
        out = std::make_unique<GoSubStmt>(s->loc, static_cast<const GoSubStmt&>(*s).labelName); break;
    case ASTNodeKind::ReturnStmt:
        out = std::make_unique<ReturnStmt>(s->loc); break;
    case ASTNodeKind::OnErrorStmt: {
        auto& x = static_cast<const OnErrorStmt&>(*s);
        out = std::make_unique<OnErrorStmt>(x.loc, x.errorKind, x.labelName); break;
    }
    case ASTNodeKind::ResumeStmt: {
        auto& x = static_cast<const ResumeStmt&>(*s);
        out = std::make_unique<ResumeStmt>(x.loc, x.resumeKind, x.labelName); break;
    }
    case ASTNodeKind::ErrorStmt: {
        auto& x = static_cast<const ErrorStmt&>(*s);
        auto n = cloneExprInner(x.errorNumber.get());
        if (x.errorNumber && !n) break;
        out = std::make_unique<ErrorStmt>(x.loc, std::move(n)); break;
    }
    case ASTNodeKind::OnGoToStmt: {
        auto& x = static_cast<const OnGoToStmt&>(*s);
        auto idx = cloneExprInner(x.index.get());
        if (x.index && !idx) break;
        out = std::make_unique<OnGoToStmt>(x.loc, std::move(idx), x.labels); break;
    }
    case ASTNodeKind::OnGoSubStmt: {
        auto& x = static_cast<const OnGoSubStmt&>(*s);
        auto idx = cloneExprInner(x.index.get());
        if (x.index && !idx) break;
        out = std::make_unique<OnGoSubStmt>(x.loc, std::move(idx), x.labels); break;
    }
    case ASTNodeKind::MidStmt: {
        auto& x = static_cast<const MidStmt&>(*s);
        auto t = cloneExprInner(x.target.get());
        auto st = cloneExprInner(x.start.get());
        auto ln = cloneExprInner(x.length.get());
        auto v = cloneExprInner(x.value.get());
        if ((x.target && !t) || (x.start && !st) || (x.length && !ln) || (x.value && !v)) break;
        out = std::make_unique<MidStmt>(x.loc, std::move(t), std::move(st),
                                        std::move(ln), std::move(v), x.hasLength);
        break;
    }
    case ASTNodeKind::ExitStmt:
        out = std::make_unique<ExitStmt>(s->loc, static_cast<const ExitStmt&>(*s).exitKind); break;
    case ASTNodeKind::StopStmt:
        out = std::make_unique<StopStmt>(s->loc); break;
    case ASTNodeKind::AsmStmt: {
        auto& x = static_cast<const AsmStmt&>(*s);
        auto n = std::make_unique<AsmStmt>(s->loc);
        n->lines = x.lines;
        n->naked = x.naked;
        n->clobbers = x.clobbers;
        out = std::move(n);
        break;
    }
    case ASTNodeKind::EndStmt:
        out = std::make_unique<EndStmt>(s->loc); break;
    case ASTNodeKind::CallStmt: {
        auto& x = static_cast<const CallStmt&>(*s);
        auto c = cloneExprInner(x.callee.get());
        if (x.callee && !c) break;
        out = std::make_unique<CallStmt>(x.loc, std::move(c)); break;
    }
    case ASTNodeKind::ReDimStmt: {
        auto& x = static_cast<const ReDimStmt&>(*s);
        std::vector<ReDimStmt::Dimension> dims;
        bool ok = true;
        for (auto& d : x.dimensions) {
            ReDimStmt::Dimension nd;
            nd.lower = cloneExprInner(d.lower.get());
            nd.upper = cloneExprInner(d.upper.get());
            if ((d.lower && !nd.lower) || (d.upper && !nd.upper)) { ok = false; break; }
            dims.push_back(std::move(nd));
        }
        if (!ok) break;
        auto o = std::make_unique<ReDimStmt>(x.loc, x.preserve, x.varName, std::move(dims),
                                             cloneTypeRef(x.asType.get()));
        if (x.asType && !o->asType) break;
        o->targetExpr = cloneExprInner(x.targetExpr.get());
        if (x.targetExpr && !o->targetExpr) break;
        out = std::move(o); break;
    }
    case ASTNodeKind::EraseStmt: {
        auto& x = static_cast<const EraseStmt&>(*s);
        auto o = std::make_unique<EraseStmt>(s->loc, x.varNames);
        // 账 #186: 复杂目标 (带下标的成员链) 也要跟着深拷贝, 否则克隆后静默退化成销毁整个数组
        for (auto& e : x.targets) o->targets.push_back(cloneExprInner(e.get()));
        out = std::move(o); break;
    }
    case ASTNodeKind::LabelStmt:
        out = std::make_unique<LabelStmt>(s->loc, static_cast<const LabelStmt&>(*s).labelName); break;
    case ASTNodeKind::RaiseEventStmt: {
        auto& x = static_cast<const RaiseEventStmt&>(*s);
        out = std::make_unique<RaiseEventStmt>(x.loc, x.eventName, cloneExprVec(x.args));
        break;
    }
    // ---- 文件 I/O 与杂项 (ast_stmt_io.hpp) ----
    case ASTNodeKind::OpenStmt: {
        auto& x = static_cast<const OpenStmt&>(*s);
        auto p = cloneExprInner(x.pathName.get());
        auto fn = cloneExprInner(x.fileNumber.get());
        auto rl = cloneExprInner(x.recordLength.get());
        if ((x.pathName && !p) || (x.fileNumber && !fn) || (x.recordLength && !rl)) break;
        out = std::make_unique<OpenStmt>(x.loc, std::move(p), x.mode, x.access, x.lock,
                                         std::move(fn), std::move(rl));
        break;
    }
    case ASTNodeKind::CloseStmt:
        out = std::make_unique<CloseStmt>(s->loc, cloneExprVec(static_cast<const CloseStmt&>(*s).fileNumbers)); break;
    case ASTNodeKind::GetStmt: {
        auto& x = static_cast<const GetStmt&>(*s);
        auto a = cloneExprInner(x.fileNumber.get());
        auto b = cloneExprInner(x.recordNumber.get());
        auto c = cloneExprInner(x.varName.get());
        if ((x.fileNumber && !a) || (x.recordNumber && !b) || (x.varName && !c)) break;
        out = std::make_unique<GetStmt>(x.loc, std::move(a), std::move(b), std::move(c));
        break;
    }
    case ASTNodeKind::PutStmt: {
        auto& x = static_cast<const PutStmt&>(*s);
        auto a = cloneExprInner(x.fileNumber.get());
        auto b = cloneExprInner(x.recordNumber.get());
        auto c = cloneExprInner(x.varName.get());
        if ((x.fileNumber && !a) || (x.recordNumber && !b) || (x.varName && !c)) break;
        out = std::make_unique<PutStmt>(x.loc, std::move(a), std::move(b), std::move(c));
        break;
    }
    case ASTNodeKind::InputStmt: {
        auto& x = static_cast<const InputStmt&>(*s);
        out = std::make_unique<InputStmt>(x.loc, cloneExprInner(x.fileNumber.get()),
                                          cloneExprVec(x.varList));
        break;
    }
    case ASTNodeKind::PrintStmt: {
        auto& x = static_cast<const PrintStmt&>(*s);
        auto o = std::make_unique<PrintStmt>(x.loc, cloneExprInner(x.fileNumber.get()),
                                             cloneExprVec(x.outputList));
        o->isFormPrint = x.isFormPrint;
        out = std::move(o); break;
    }
    case ASTNodeKind::WriteStmt: {
        auto& x = static_cast<const WriteStmt&>(*s);
        out = std::make_unique<WriteStmt>(x.loc, cloneExprInner(x.fileNumber.get()),
                                          cloneExprVec(x.outputList));
        break;
    }
    case ASTNodeKind::LineInputStmt: {
        auto& x = static_cast<const LineInputStmt&>(*s);
        auto a = cloneExprInner(x.fileNumber.get());
        auto b = cloneExprInner(x.varName.get());
        if ((x.fileNumber && !a) || (x.varName && !b)) break;
        out = std::make_unique<LineInputStmt>(x.loc, std::move(a), std::move(b));
        break;
    }
    case ASTNodeKind::WidthStmt: {
        auto& x = static_cast<const WidthStmt&>(*s);
        auto a = cloneExprInner(x.fileNumber.get());
        auto b = cloneExprInner(x.width.get());
        if ((x.fileNumber && !a) || (x.width && !b)) break;
        out = std::make_unique<WidthStmt>(x.loc, std::move(a), std::move(b));
        break;
    }
    case ASTNodeKind::SeekStmt: {
        auto& x = static_cast<const SeekStmt&>(*s);
        auto a = cloneExprInner(x.fileNumber.get());
        auto b = cloneExprInner(x.position.get());
        if ((x.fileNumber && !a) || (x.position && !b)) break;
        out = std::make_unique<SeekStmt>(x.loc, std::move(a), std::move(b));
        break;
    }
    case ASTNodeKind::LockStmt: {
        auto& x = static_cast<const LockStmt&>(*s);
        auto a = cloneExprInner(x.fileNumber.get());
        auto b = cloneExprInner(x.start.get());
        auto c = cloneExprInner(x.end.get());
        if ((x.fileNumber && !a) || (x.start && !b) || (x.end && !c)) break;
        out = std::make_unique<LockStmt>(x.loc, std::move(a), std::move(b), std::move(c));
        break;
    }
    case ASTNodeKind::UnlockStmt: {
        auto& x = static_cast<const UnlockStmt&>(*s);
        auto a = cloneExprInner(x.fileNumber.get());
        auto b = cloneExprInner(x.start.get());
        auto c = cloneExprInner(x.end.get());
        if ((x.fileNumber && !a) || (x.start && !b) || (x.end && !c)) break;
        out = std::make_unique<UnlockStmt>(x.loc, std::move(a), std::move(b), std::move(c));
        break;
    }
    case ASTNodeKind::ResetStmt:
        out = std::make_unique<ResetStmt>(s->loc); break;
    case ASTNodeKind::NameStmt: {
        auto& x = static_cast<const NameStmt&>(*s);
        auto a = cloneExprInner(x.oldPath.get());
        auto b = cloneExprInner(x.newPath.get());
        if ((x.oldPath && !a) || (x.newPath && !b)) break;
        out = std::make_unique<NameStmt>(x.loc, std::move(a), std::move(b));
        break;
    }
    case ASTNodeKind::FileCopyStmt: {
        auto& x = static_cast<const FileCopyStmt&>(*s);
        auto a = cloneExprInner(x.source.get());
        auto b = cloneExprInner(x.destination.get());
        if ((x.source && !a) || (x.destination && !b)) break;
        out = std::make_unique<FileCopyStmt>(x.loc, std::move(a), std::move(b));
        break;
    }
    case ASTNodeKind::KillStmt:
    case ASTNodeKind::MkDirStmt:
    case ASTNodeKind::RmDirStmt:
    case ASTNodeKind::ChDirStmt: {
        auto& x = static_cast<const KillStmt&>(*s);  // 同构: 单 pathName
        auto p = cloneExprInner(x.pathName.get());
        if (x.pathName && !p) break;
        switch (s->kind) {
        case ASTNodeKind::KillStmt:  out = std::make_unique<KillStmt>(x.loc, std::move(p)); break;
        case ASTNodeKind::MkDirStmt: out = std::make_unique<MkDirStmt>(x.loc, std::move(p)); break;
        case ASTNodeKind::RmDirStmt: out = std::make_unique<RmDirStmt>(x.loc, std::move(p)); break;
        default:                     out = std::make_unique<ChDirStmt>(x.loc, std::move(p)); break;
        }
        break;
    }
    case ASTNodeKind::ChDriveStmt: {
        auto& x = static_cast<const ChDriveStmt&>(*s);
        out = std::make_unique<ChDriveStmt>(x.loc, cloneExprInner(x.drive.get()));
        if (x.drive && !out) break;
        break;
    }
    case ASTNodeKind::BeepStmt:
        out = std::make_unique<BeepStmt>(s->loc); break;
    case ASTNodeKind::DoEventsStmt:
        out = std::make_unique<DoEventsStmt>(s->loc); break;
    case ASTNodeKind::OptionStmt:
        out = std::make_unique<OptionStmt>(s->loc, static_cast<const OptionStmt&>(*s).optionKind); break;
    case ASTNodeKind::ImplementsStmt: {
        auto& x = static_cast<const ImplementsStmt&>(*s);
        auto c = std::make_unique<ImplementsStmt>(x.loc, x.interfaceName);
        c->viaField = x.viaField;  // ai/022 B10: 委托子句要跟着拷，否则泛型/克隆路径丢 Via
        out = std::move(c); break;
    }
    case ASTNodeKind::DefTypeStmt: {
        auto& x = static_cast<const DefTypeStmt&>(*s);
        out = std::make_unique<DefTypeStmt>(x.loc, x.defKind, x.ranges); break;
    }
    case ASTNodeKind::LocalDeclStmt: {
        auto& x = static_cast<const LocalDeclStmt&>(*s);
        auto d = cloneDeclAny(x.decl.get());
        if (x.decl && !d) break;
        out = std::make_unique<LocalDeclStmt>(x.loc, std::move(d)); break;
    }
    default:
        break;  // failed
    }
    if (!out) failed_ = true;
    return out;
}

// ============================================================
// Decl (体内子集)
// ============================================================

DeclPtr ASTCloner::cloneDeclAny(const Decl* d) {
    if (!d) return nullptr;
    switch (d->kind) {
    case ASTNodeKind::VariableDecl: {
        auto& x = static_cast<const VariableDecl&>(*d);
        std::vector<VariableDecl::Dimension> dims;
        for (auto& dm : x.dimensions) {
            VariableDecl::Dimension nd;
            nd.lower = cloneExprInner(dm.lower.get());
            nd.upper = cloneExprInner(dm.upper.get());
            if ((dm.lower && !nd.lower) || (dm.upper && !nd.upper)) return nullptr;
            dims.push_back(std::move(nd));
        }
        auto out = std::make_unique<VariableDecl>(
            x.loc, x.access, x.name, x.isWithEvents, x.isStatic, x.isNew,
            cloneTypeRef(x.asType.get()), cloneExprInner(x.initializer.get()),
            std::move(dims), x.isDynamicArray);
        if ((x.asType && !out->asType) || (x.initializer && !out->initializer)) return nullptr;
        return out;
    }
    case ASTNodeKind::ConstDecl: {
        auto& x = static_cast<const ConstDecl&>(*d);
        auto out = std::make_unique<ConstDecl>(x.loc, x.access, x.name,
            cloneTypeRef(x.asType.get()), cloneExprInner(x.value.get()));
        if ((x.asType && !out->asType) || (x.value && !out->value)) return nullptr;
        return out;
    }
    case ASTNodeKind::MultiDecl: {
        auto& x = static_cast<const MultiDecl&>(*d);
        DeclList inner;
        for (auto& dd : x.declarations) {
            auto c = cloneDeclAny(dd.get());
            if (!c) return nullptr;
            inner.push_back(std::move(c));
        }
        return std::make_unique<MultiDecl>(x.loc, std::move(inner));
    }
    default:
        return nullptr;
    }
}

std::unique_ptr<ParameterDecl> ASTCloner::cloneParam(const ParameterDecl& p) {
    auto out = std::make_unique<ParameterDecl>(
        p.loc, p.name, p.isOptional, p.isByVal, p.isParamArray,
        cloneTypeRef(p.asType.get()), cloneExprInner(p.defaultValue.get()));
    if ((p.asType && !out->asType) || (p.defaultValue && !out->defaultValue))
        return nullptr;
    return out;
}

std::unique_ptr<TypeDecl> ASTCloner::cloneTypeDecl(const TypeDecl& d,
                                                   const std::string& newName) {
    std::vector<std::unique_ptr<TypeMember>> members;
    for (auto& m : d.members) {
        auto mt = std::make_unique<TypeMember>(m->loc, m->name,
                                               cloneTypeRef(m->type.get()),
                                               m->arraySize ? cloneExpr(m->arraySize.get())
                                                            : nullptr);
        mt->isArrayDynamic = m->isArrayDynamic;
        if (m->type && !mt->type) { failed_ = true; return nullptr; }
        members.push_back(std::move(mt));
    }
    return std::make_unique<TypeDecl>(d.loc, d.access, newName, std::move(members));
}

std::unique_ptr<SubDecl> ASTCloner::cloneSubDecl(const SubDecl& d,
                                                 const std::string& newName) {
    auto params = std::vector<std::unique_ptr<ParameterDecl>>();
    for (auto& p : d.params) {
        auto c = cloneParam(*p);
        if (!c) { failed_ = true; return nullptr; }
        params.push_back(std::move(c));
    }
    auto body = cloneStmtList(d.body);
    if (failed_) return nullptr;
    auto out = std::make_unique<SubDecl>(d.loc, d.access, newName, std::move(params),
                                         std::move(body), d.isStatic);
    out->implementsClauses = d.implementsClauses;
    out->isNaked = d.isNaked;
    return out;
}

std::unique_ptr<FunctionDecl> ASTCloner::cloneFunctionDecl(const FunctionDecl& d,
                                                           const std::string& newName) {
    auto params = std::vector<std::unique_ptr<ParameterDecl>>();
    for (auto& p : d.params) {
        auto c = cloneParam(*p);
        if (!c) { failed_ = true; return nullptr; }
        params.push_back(std::move(c));
    }
    auto ret = cloneTypeRef(d.returnType.get());
    if (d.returnType && !ret) { failed_ = true; return nullptr; }
    auto body = cloneStmtList(d.body);
    if (failed_) return nullptr;
    auto out = std::make_unique<FunctionDecl>(d.loc, d.access, newName, std::move(params),
                                              std::move(ret), std::move(body), d.isStatic);
    out->implementsClauses = d.implementsClauses;
    out->isNaked = d.isNaked;
    return out;
}

std::unique_ptr<PropertyDecl> ASTCloner::clonePropertyDecl(const PropertyDecl& d,
                                                           const std::string& newName) {
    auto params = std::vector<std::unique_ptr<ParameterDecl>>();
    for (auto& p : d.params) {
        auto c = cloneParam(*p);
        if (!c) { failed_ = true; return nullptr; }
        params.push_back(std::move(c));
    }
    auto ret = cloneTypeRef(d.returnType.get());
    if (d.returnType && !ret) { failed_ = true; return nullptr; }
    auto body = cloneStmtList(d.body);
    if (failed_) return nullptr;
    auto out = std::make_unique<PropertyDecl>(d.loc, d.access, d.propKind, newName,
                                              std::move(params), std::move(ret), std::move(body));
    out->isDefault = d.isDefault;
    out->implementsClauses = d.implementsClauses;
    return out;
}

// ============================================================
// Module (泛型类特化, G4)
// ============================================================

std::unique_ptr<Module> ASTCloner::cloneModule(const Module& m,
                                               const std::string& newName) {
    auto out = std::make_unique<Module>(m.loc, m.filename);
    out->filename = m.filename;
    out->moduleName = newName;
    out->packageName = m.packageName;  // ai/023 S03: 包归属随模块克隆 (泛型实例化同包)
    out->isClassModule = m.isClassModule;
    out->isFormModule = false;
    out->instancing = m.instancing;
    // classTypeParams 留空 → 特化副本作为普通类模块流过全管线

    for (auto& opt : m.options) {
        out->options.push_back(std::make_unique<OptionStmt>(opt->loc, opt->optionKind));
    }
    for (auto& dt : m.defTypes) {
        out->defTypes.push_back(std::make_unique<DefTypeStmt>(
            dt->loc, dt->defKind,
            std::vector<DefTypeStmt::LetterRange>(dt->ranges)));
    }
    for (auto& attr : m.attributes) {
        ExprPtr v;
        if (attr->attrName == "VB_Name") {
            // 类符号名以 newName 为准 (driver_frontend 从该字面量取 moduleName)
            auto lit = std::make_unique<LiteralExpr>(
                attr->loc, LiteralKind::String, "\"" + newName + "\"");
            v = std::move(lit);
        } else {
            v = cloneExprInner(attr->value.get());
            if (attr->value && !v) { failed_ = true; return nullptr; }
        }
        out->attributes.push_back(std::make_unique<AttributeStmt>(
            attr->loc, attr->attrName, std::move(v)));
    }

    for (auto& d : m.declarations) {
        DeclPtr c;
        switch (d->kind) {
        case ASTNodeKind::SubDecl:
            c = cloneSubDecl(static_cast<SubDecl&>(*d), static_cast<SubDecl&>(*d).name);
            break;
        case ASTNodeKind::FunctionDecl:
            c = cloneFunctionDecl(static_cast<FunctionDecl&>(*d),
                                  static_cast<FunctionDecl&>(*d).name);
            break;
        case ASTNodeKind::PropertyDecl:
            c = clonePropertyDecl(static_cast<PropertyDecl&>(*d),
                                  static_cast<PropertyDecl&>(*d).name);
            break;
        case ASTNodeKind::TypeDecl:
            c = cloneTypeDecl(static_cast<TypeDecl&>(*d),
                              static_cast<TypeDecl&>(*d).name);
            break;
        default:
            c = cloneDeclAny(d.get());  // VariableDecl/ConstDecl/MultiDecl
            break;
        }
        if (!c) { failed_ = true; return nullptr; }
        out->declarations.push_back(std::move(c));
    }
    return out;
}

} // namespace vb6c3
