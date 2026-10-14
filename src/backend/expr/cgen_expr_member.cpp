#include "backend/cgen.hpp"
#include "common/canvas_drawing.hpp"  // 账 #232①: 画布动词表 (Printer 成员那一路问它要名字)
#include "common/host_pseudo.hpp"     // Fix 255: precheck 里限定形式的宿主集合 (PropertyPage.Controls.X) 问那张表
#include <algorithm>
#include <cctype>
#include <iostream>
#include <cstdlib>
#include <functional>
#include <cstdio>
#include "driver/detail/phase_timing.hpp"   // ai/030 续: member 分段打点 (c3StatSplit; 必须在 namespace 外)

namespace vb6c3 {

// --- cgen_expr_member.cpp: MemberAccessExpr 求值（类成员 / COM / UDT 字段访问） ---
// 2026-09-17 拆分：原 1496 行单文件（CCodeGen::visit(MemberAccessExpr&) 函数体 1481 行）拆为
//   本文件（伞文件）                          —— include + namespace + 函数签名
//   detail/cgen_expr_member_precheck.inc      —— AddressOf 模块限定函数 / Err·内建对象 / 控件数组 / 内置对象方法（原 13~131 行）
//   detail/cgen_expr_member_form_builtin.inc  —— Form 的 Me.member（Fix 023e）与 IdentifierExpr 首段：App / 窗体默认实例 / Clipboard·Screen·Printer·Forms / 控件属性（原 132~362 行）
//   detail/cgen_expr_member_obj_dispatch.inc  —— COM 前期绑定·接口·后期绑定·Variant·promoted 分派；UDT 字段与函数名引用（Fix 031/085/084z-4/086）（原 363~552 行）
//   detail/cgen_expr_member_class_module.inc  —— 类实例成员分派（优先级2）与模块名.方法（优先级3）、模块前缀解析（原 553~774 行）
//   detail/cgen_expr_member_m22_module.inc    —— M22 跨模块变量访问与函数调用结果上的方法调用（Fix 083d/084y-4）（原 775~943 行）
//   detail/cgen_expr_member_generic_access.inc—— 通用成员访问 + 链式 COM 检测 / Variant 数组元素 / UDT 字段链（Fix 090af/085）（原 944~1063 行）
//   detail/cgen_expr_member_voidptr_com.inc   —— void* 字段的 COM 化（Fix 023）与类方法链式调用（Fix 015）（原 1064~1254 行）
//   detail/cgen_expr_member_class_fallback.inc—— 类实例成员 fallback（Fix 010r-10/011r-1/014）与函数尾（原 1255~1493 行）
// 八个 .inc 是「函数体片段」，在 visit(MemberAccessExpr&) 函数体内被 #include（C++ 允许），故不用 .cpp/.hpp 后缀 ——
// 它们不是独立编译单元，单独 include 会编译不过。片段内容逐行未改（含原缩进层级）→ 零行为改动。

void CCodeGen::visit(MemberAccessExpr& node) {
    // Fix 262: 窗体的**控件 Picture 属性**上的成员访问 ——
    //   `<控件>.Picture` 的 C 面是 `vb6_GetControlPicture(hwnd)` (返回 void*,
    //   即 IPictureDisp 的 C 表示), 它的成员 (Handle/Type/Width/Height/hPal…)
    //   是那条 IDispatch 上的属性, 必须走后期绑定读。
    // 修复前 `ImagePicture.Picture.Handle` 落最末尾的 `obj + "." + member` 兜底,
    // 生成 `vb6_GetControlPicture(hwnd)  /* ctrl prop read */.Handle`
    // ⇒ C2224 "「.Handle」的左侧必须具有结构/联合类型"。
    //   实测 ComCtlsDemo PPStatusBarPanels.pag:1182/1323
    //   (`If Not ImagePicture.Picture Is Nothing Then If ImagePicture.Picture.Handle <> 0`),
    //   PPStatusBarPanels.c 1362/1548。
    // 为什么放在**函数最前**: 这条链在下面八个片段里会被若干更早的分支截走
    //   (voidptr_com 的 `objExpr + "->"` 会把 C2224 变成 C2223), 只有在入口拦才唯一。
    // RTL 侧 vb6_ComGetProp 对 HWND/宿主对象有专档 (vb6com_invoke.c:409),
    // 返回 void*, 与 `<> 0` / `Is Nothing` 这类判等天然兼容。
    if (node.object && node.object->kind == ASTNodeKind::MemberAccessExpr) {
        auto& inner262 = static_cast<MemberAccessExpr&>(*node.object);
        if (inner262.object && inner262.object->kind == ASTNodeKind::IdentifierExpr) {
            const std::string ctrl262 = Symbol::toLower(
                static_cast<IdentifierExpr&>(*inner262.object).name);
            auto itC262 = knownFormControls_.find(ctrl262);
            if (itC262 != knownFormControls_.end()
                && getControlPropReadFn(itC262->second, inner262.memberName)
                       == "vb6_GetControlPicture") {
                emitExpr(*node.object);
                std::string pic262 = std::move(lastExpr_);
                isComMarker_ = false;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                comObjExpr_.clear();
                comMemberName_.clear();
                pendingChainObj_.clear();
                lastExpr_ = "vb6_ComGetProp(" + pic262 + ", L\""
                          + node.memberName + "\")";
                return;
            }
        }
    }
    std::chrono::steady_clock::time_point _c3sMem = std::chrono::steady_clock::now();
#include "backend/detail/expr/cgen_expr_member_precheck.inc"
    c3StatSplit(13, _c3sMem);   // ai/030 续: member:precheck
#include "backend/detail/expr/cgen_expr_member_form_builtin.inc"
    c3StatSplit(14, _c3sMem);   // mForm
#include "backend/detail/expr/cgen_expr_member_obj_dispatch.inc"
    c3StatSplit(15, _c3sMem);   // mObj
#include "backend/detail/expr/cgen_expr_member_class_module.inc"
    c3StatSplit(16, _c3sMem);   // mCls
#include "backend/detail/expr/cgen_expr_member_m22_module.inc"
    c3StatSplit(17, _c3sMem);   // mM22
#include "backend/detail/expr/cgen_expr_member_generic_access.inc"
    c3StatSplit(18, _c3sMem);   // mGen
#include "backend/detail/expr/cgen_expr_member_voidptr_com.inc"
    c3StatSplit(19, _c3sMem);   // mVoid
#include "backend/detail/expr/cgen_expr_member_class_fallback.inc"
}

} // namespace vb6c3
