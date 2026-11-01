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

void CCodeGen::visit(AssignmentStmt& node) {
    // C29-3: Let 赋值期间成员对象要取默认属性 (见 memObjLetScalar_ 的说明)。
    // `Set` 语句走 visit(SetStmt&)，不受影响。
    bool savedLetScalar = memObjLetScalar_;
    memObjLetScalar_ = true;

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
