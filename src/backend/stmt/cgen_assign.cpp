#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_assign.cpp: 赋值语句生成 (Assignment / Set / Let / Mid$) ---
// 2026-09-17 拆分：原 1367 行单文件（visit(AssignmentStmt&) 函数体 1352 行）拆为
//   本文件（伞文件）                      —— 前置说明 + visit(AssignmentStmt&) 函数骨架
//   detail/cgen_assign_stmt_special.inc   —— 语句特殊形态：Date$/Time$ / Mid$ / LSet-RSet / 控件属性 / 跨模块变量（原 12~405 行）
//   detail/cgen_assign_host_pseudo.inc    —— 宿主伪对象 / Parent 成员写入（2026-09-29 新增，VBEclipse 定向）
//   detail/cgen_assign_prop_write.inc     —— 属性写入分派：类 Property Let / 默认属性 / With 块 / 链式 COM（原 406~828 行）
//   detail/cgen_assign_com_prop.inc       —— 参数化属性与 COM 写：P25/P25b / 类内 Property Let / COM SetProp / 返回值解封（原 829~1055 行）
//   detail/cgen_assign_value_sem.inc      —— 值语义：返回值变量 / BSTR / Variant 赋值（原 1056~1362 行）
// 五个 .inc 是「函数体片段」，在 visit() 函数体内被 #include（C++ 允许），故不用 .cpp/.hpp 后缀 ——
// 它们不是独立编译单元，单独 include 会编译不过。片段局部变量原样不动，逐行未改 → 零行为改动。

void CCodeGen::emitGuardedStore(const std::string& line) {
    // 本语句有守卫 ⇒ 只有「序号没变」才写目标（变了 = 本语句里 raise 过，VB6 语句级中止）。
    if (stmtAbortGuard_.empty()) { c_.emitLine(line); return; }
    if (stmtAbortTmp_ && !stmtAbortTmpType_.empty()) {
        // 右值自己会 raise（Asc/AscB/AscW 对空串报 5）：先算进临时量，再拿序号判是否写目标。
        // 只接二元形 "<target> = <expr>;"，接不上就退回旧行为（宁可不改形状，不造错码）。
        const std::string sep = " = ";
        const size_t eq = line.find(sep);
        if (eq != std::string::npos && line.size() > 2 && line.back() == ';') {
            const std::string tgt = line.substr(0, eq);
            const std::string rhs = line.substr(eq + sep.size(), line.size() - eq - sep.size() - 1);
            const std::string tmp = "_av" + std::to_string(stmtAbortSeq_);
            // 右值**先算**，但只在本语句还没 raise 时算 —— 否则 `Asc(<Null>)` 的上游 94
            // 会被 Asc 自己那句 5 盖掉（第七十八刀撤回的正是这个）。
            c_.emitLine(stmtAbortTmpType_ + " " + tmp + " = (" + stmtAbortGuard_ + " == vb6_ErrSerial()) ? ("
                        + rhs + ") : (" + stmtAbortTmpType_ + ")0;");
            c_.emitLine("if (" + stmtAbortGuard_ + " == vb6_ErrSerial()) " + tgt + " = " + tmp + ";");
            return;
        }
    }
    c_.emitLine("if (" + stmtAbortGuard_ + " == vb6_ErrSerial()) " + line);
}

namespace {
// C30-11 第八十三刀 (§B227 续): 右值表达式树里有没有 Asc/AscB/AscW 调用。
// 只认这三枚 —— 它们 raise 在**本语句的表达式求值里**（空串直接 vb6_ErrRaiseNumber(5)），
// 不会被哪个被调方内部用 `On Error Resume Next` 吞掉，所以「序号一变 = 真中止」对它们成立。
// 不要再放宽成「任何嵌套调用」（见调用点注）：那一版把「被调方内部已处理的错误」也当成
// 中止信号，实测在 `[VBP] fdrawstate_x86` 上翻红。
bool ascCallInTree(const Expr* e) {
    if (!e) return false;
    switch (e->kind) {
        case ASTNodeKind::IndexOrCallExpr: {
            const auto& c = static_cast<const IndexOrCallExpr&>(*e);
            if (c.callee && c.callee->kind == ASTNodeKind::IdentifierExpr) {
                const std::string n =
                    Symbol::toLower(static_cast<const IdentifierExpr&>(*c.callee).name);
                if (n == "asc" || n == "ascb" || n == "ascw") return true;
            }
            if (ascCallInTree(c.callee.get())) return true;
            for (const auto& a : c.positional) if (ascCallInTree(a.get())) return true;
            for (const auto& a : c.named) if (ascCallInTree(a.value.get())) return true;
            return false;
        }
        case ASTNodeKind::BinaryExpr: {
            const auto& b = static_cast<const BinaryExpr&>(*e);
            return ascCallInTree(b.left.get()) || ascCallInTree(b.right.get());
        }
        case ASTNodeKind::UnaryExpr:
            return ascCallInTree(static_cast<const UnaryExpr&>(*e).operand.get());
        case ASTNodeKind::MemberAccessExpr:
            return ascCallInTree(static_cast<const MemberAccessExpr&>(*e).object.get());
        case ASTNodeKind::DictionaryAccessExpr:
            return ascCallInTree(static_cast<const DictionaryAccessExpr&>(*e).object.get());
        case ASTNodeKind::NewExpr: {
            const auto& n = static_cast<const NewExpr&>(*e);
            for (const auto& a : n.args) if (ascCallInTree(a.get())) return true;
            return false;
        }
        case ASTNodeKind::TypeOfExpr:
            return ascCallInTree(static_cast<const TypeOfExpr&>(*e).object.get());
        default:
            return false;
    }
}
} // namespace

void CCodeGen::visit(AssignmentStmt& node) {
    // C29-3: Let 赋值期间成员对象要取默认属性 (见 memObjLetScalar_ 的说明)。
    // `Set` 语句走 visit(SetStmt&)，不受影响。
    bool savedLetScalar = memObjLetScalar_;
    memObjLetScalar_ = true;

    // C30-11 第七十九刀 (§B227): 可能 raise 的赋值语句加一道「语句级中止」守卫。
    // 判据 = 右值是一枚调用、且它的某个实参静态类型是 Variant —— 那一步会发
    // vb6_CStrStrict(<变体>)，Null 进来就在赋值之前把 94 报掉（真机 W4b/W5b/W6b 的
    // `n = Len(...)` 全是这个形）。只给**整型标量目标**发守卫（String 目标那一族另开一刀，
    // 因为它的保存/恢复要连所有权一起算）。
    struct StmtAbortScope {
        CCodeGen* g;
        ~StmtAbortScope() {
            g->stmtAbortGuard_.clear();
            g->stmtAbortTmp_ = false;
            g->stmtAbortTmpType_.clear();
        }
    } stmtAbortScope{this};
    stmtAbortGuard_.clear();
    stmtAbortTmp_ = false;
    stmtAbortTmpType_.clear();
    {
        const Vb6Type tgtT79 = node.target ? inferExprType(*node.target) : Vb6Type::Unknown;
        const bool tgtInt79 = tgtT79 == Vb6Type::Integer || tgtT79 == Vb6Type::Long
                           || tgtT79 == Vb6Type::Byte;
        const bool tgt79179 = tgtInt79 || tgtT79 == Vb6Type::String;
        bool mayRaise79 = false;
        // C30-11 第八十三刀：String 目标那一形（`s = CStr(Asc(""))`）在**顶层**看不见 ——
        // 右值的外套是 CStr/CLng 这类转换函数，实参档是 Integer（不是 Variant）⇒ 第七十九/
        // 八十一刀那条判据一律不 fire，语句照样把 "" 写进目标（真机：5 + 目标不动）。
        // 补一条**只认 Asc 族**的树扫描，理由见 ascCallInTree 顶注。
        // ⚠ 不要放宽成「右值里任何嵌套调用」：那一版实测在 `[VBP] fdrawstate_x86` 上翻红
        // （同一分片改前两次都是 2 红：c29listview + iface 那一枚争用）。成因是判据分不清
        // 两种「序号变了」——① 错误往外传（无处理器）⇒ 真机语句级中止，目标不动；
        // ② 错误在被调方内部被 `On Error Resume Next` 吃掉 ⇒ 真机语句照执行，目标该写。
        // `vb6_ErrSerial()` 对两者一视同仁地自增 ⇒ 递归一宽，② 就被误当成 ①；要宽到那一层
        // 得先给 RTL 一枚「错误是否还在往外爬」的信号，那是另一刀的事。
        const bool ascInRhs83 =
            (tgtT79 == Vb6Type::String) && ascCallInTree(node.value.get());
        if (tgt79179 && node.value && node.value->kind == ASTNodeKind::IndexOrCallExpr) {
            auto& call79 = static_cast<IndexOrCallExpr&>(*node.value);
            // 第八十一刀: Asc/AscB/AscW 自己会在**空串**上报 5 —— 那条路不靠变体转换，
            // 所以上面的实参档判据看不见它。不把这三枚算进来的话，`n = Asc("")` 会是
            // 「先报 5、再把目标写成 0」（真机: 目标不动）。
            if (call79.callee && call79.callee->kind == ASTNodeKind::IdentifierExpr) {
                const std::string cn79 =
                    Symbol::toLower(static_cast<IdentifierExpr&>(*call79.callee).name);
                if (cn79 == "asc" || cn79 == "ascb" || cn79 == "ascw") {
                    mayRaise79 = true;
                    // 它 raise 在**调用内部**（不是前置转换）⇒ 这一档要发临时量形。
                    // 只给整型标量目标发（BSTR 目标那一半由串档自己发临时量，见
                    // 第八十三刀的 emitBstrSlotStore）。
                    if (tgtInt79) {
                        stmtAbortTmp_ = true;
                        stmtAbortTmpType_ = mapType(tgtT79);
                    }
                }
            }
            for (const auto& a79 : call79.positional) {
                if (mayRaise79) break;
                if (!a79) continue;
                const Vb6Type at79 = inferExprType(*a79);
                // Variant 实参要过 vb6_CStrStrict；`Null`/`Empty` 字面量也一样（转串就是 94）。
                // 字面量那一路 inferExprType 给的档位不稳（Null 可能落成 Empty/兜底），
                // 所以直接看语法形。
                bool lit79 = false;
                if (a79->kind == ASTNodeKind::LiteralExpr) {
                    const LiteralKind lk79 = static_cast<const LiteralExpr&>(*a79).literalKind;
                    lit79 = lk79 == LiteralKind::Null || lk79 == LiteralKind::Empty;
                }
                if (lit79 || at79 == Vb6Type::Variant || at79 == Vb6Type::Null
                    || at79 == Vb6Type::Empty) { mayRaise79 = true; break; }
            }
        }
        // 第八十三刀: 顶层判据没命中、但右值树里有 Asc 族 ⇒ String 目标也发守卫
        //（它的临时量由 emitBstrSlotStore 的两支各自发，见那里的第八十三刀注）。
        if (ascInRhs83) mayRaise79 = true;
        if (mayRaise79) {            const std::string gv79 = "_e" + std::to_string(stmtAbortSeq_++);
            c_.emitLine("const int32_t " + gv79 + " = vb6_ErrSerial();");
            stmtAbortGuard_ = gv79;
        }
    }

    // 账 #611: 赋值左值是工程类**默认属性索引** (`writer("k") = v`) —— 把 callee
    // 换成 `writer.Item`, 后面 P6.4+ 那一支 (cgen_assign_com_prop.inc) 才会按写
    // 方向形参表发 prop_let_Item; 不改写则 emitExpr 产出 `writer(_own_N)` 这个非法
    // 左值, cl C2063 + C2106。读侧在 cgen_expr_call_prelude.inc 用同一个改写器。
    // 只换 callee、不换整枚节点 (同一枚节点会被多次 emitExpr, 掏空即崩)。
    if (node.target && node.target->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& iocTgt611 = static_cast<IndexOrCallExpr&>(*node.target);
        if (ExprPtr calleeTgt611 = makeClassDefaultMemberCallee(iocTgt611)) {
            iocTgt611.callee = std::move(calleeTgt611);
        }
    }

    // 账 #305 第320刀 (账 #592): `App.Title = v` —— VB6 允许运行期改应用标题 (语料 16 行)。
    // 改前这里把读取侧的折算当左值: `vb6_App_EXEName() = vb6_BSTR_FromStr(L"...")` ⇒ C2106,
    // 整棵工程编不过。读写各一处出口: 写落 vb6_SetAppTitle (RTL 自持一份副本, 借入不交出),
    // 读走 vb6_App_Title。局部同名变量优先 (VB6 里局部遮蔽内建名), 与读取侧同一判据。
    if (node.target && node.target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maAt = static_cast<MemberAccessExpr&>(*node.target);
        if (maAt.object && maAt.object->kind == ASTNodeKind::IdentifierExpr) {
            const std::string objAt = Symbol::toLower(
                static_cast<IdentifierExpr&>(*maAt.object).name);
            const std::string memAt = Symbol::toLower(maAt.memberName);
            if (objAt == "app" && memAt == "title" && !knownLocalVars_.count(objAt)) {
                emitExpr(*node.value);
                std::string rhsAt = wrapToBSTR(std::move(lastExpr_), *node.value);
                c_.emitLine("vb6_SetAppTitle(" + rhsAt + ");");
                memObjLetScalar_ = savedLetScalar;
                return;
            }
        }
    }
#include "backend/detail/stmt/cgen_assign_host_pseudo.inc"
#include "backend/detail/stmt/cgen_assign_stmt_special.inc"
#include "backend/detail/stmt/cgen_assign_prop_write.inc"
#include "backend/detail/stmt/cgen_assign_com_prop.inc"
#include "backend/detail/stmt/cgen_assign_value_sem.inc"
    memObjLetScalar_ = savedLetScalar;
}


} // namespace vb6c3
