#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>
#include <cstdio>

namespace vb6c3 {

// --- cgen_expr_call.cpp: IndexOrCallExpr 求值（数组索引 / 函数调用统一路径） ---
// 2026-09-17 拆分：原 3627 行单文件（单函数 visit(IndexOrCallExpr&) 占函数体 3611 行）拆为
//   本文件（伞文件）                             —— 前置说明 + visit(IndexOrCallExpr&) 函数骨架
//   detail/cgen_expr_call_prelude.inc           —— 头部局部变量 + ParamArray 识别 + Variant 数组索引识别（原 13~299 行）
//   detail/cgen_expr_call_builtin_pre.inc       —— 内置函数前置拦截：IIf / VarPtr / TypeName(Me) / Choose / Array / IsMissing / Len(udt)（原 300~620 行）
//   detail/cgen_expr_call_callee_ident.inc      —— Fix 037 IdentifierExpr callee：Pattern F / F2 / G（原 621~805 行）
//   detail/cgen_expr_call_callee_member.inc     —— Fix 037 MemberAccessExpr callee：Pattern A / B / C / D（原 806~1009 行）
//   detail/cgen_expr_call_callee_withm.inc      —— Fix 060 WithMemberExpr callee + callee 发射 + WebBrowser/ListBox 方法拦截（原 1010~1167 行）
//   detail/cgen_expr_call_com_bind.inc          —— COM 后期/前期绑定检测（P6.2/P6.3/P6.4）+ callee 拆分判定（原 1168~1553 行）
//   detail/cgen_expr_call_callee_params.inc     —— callee 形参解析：calleeParams / builtin / Declare / ANSI（原 1554~1770 行）
//   detail/cgen_expr_call_arg_emit.inc          —— 位置实参发射（一）：UDT 打包 / Declare ANSI / ByRef 取地址（原 1771~2136 行）
//   detail/cgen_expr_call_arg_variant.inc       —— 位置实参发射（二）：ByVal Variant 包装 / Variant 提取（原 2137~2453 行）
//   detail/cgen_expr_call_named_args.inc        —— named 实参（含 Optional 空位填充）+ argList 组装 + ParamArray 展开（原 2454~2860 行）
//   detail/cgen_expr_call_builtin_fixup.inc     —— 内置函数调用点修正：UBound/LBound、字符串/日期/财务函数、实参个数（原 2861~3250 行）
//   detail/cgen_expr_call_opt_pad.inc           —— Optional 参数补齐 + IsMissing _has_ 标志 + 隐式 me 实参（原 pad_conv 5~81 行）
//   detail/cgen_expr_call_conv_cstr.inc         —— CStr 家族适配：按实参类型选 CStr/CStrSingle/CStrLong/CStrDbl（原 pad_conv 82~164 行）
//   detail/cgen_expr_call_conv_numeric.inc      —— 数值转换函数适配：CInt/CLng/CDbl/CByte/CSng/CBool（原 pad_conv 165~248 行）
//   detail/cgen_expr_call_conv_output.inc       —— 输出侧适配：MsgBox / Format 首参包装 + CStr 兜底（原 pad_conv 249~353 行）
//   detail/cgen_expr_call_conv_varargs_lenb.inc —— CallByName 变参打包 + LenB 类型分派（原 pad_conv 354~436 行）
//   detail/cgen_expr_call_builtin_argtype.inc   —— 内建函数实参类型适配 Fix 108c（原 pad_conv 437~546 行）
// 16 个 .inc 是「函数体片段」，在 visit(IndexOrCallExpr&) 的函数体内被 #include（C++ 允许），故不用 .cpp/.hpp 后缀 ——
// 它们不是独立编译单元，单独 include 会编译不过。片段内局部变量与 Pattern 拦截分支原样不动，逐行未改 → 零行为改动。
// 其中 arg_emit / arg_variant 两段位于 1771 行那个遍历位置实参的 for 循环内部（相对花括号深度 1），
// 单看片段自身不闭合，但头 + 全部片段 + 尾拼回后与原文件逐行一致、语义等价。
// 2026-09-19：原 pad_conv.inc（545 行）按函数体内既有的顶层分节注释再切 6 段（全部自身花括号平衡），
// 切分脚本对「片段按序拼回 == 拆分前原文件对应区间」做逐字节断言。

// 内置函数表 (builtinFuncs) + 键集合 (kBuiltinFuncNames): 与 cgen_expr_ident.cpp 同一份
// (内部链接, 各 TU 一份副本), 供下面 Pattern G 判定"内置函数优先于类成员属性"时复用。
#include "backend/detail/expr/cgen_expr_ident_builtin_table.inc"

void CCodeGen::visit(IndexOrCallExpr& node) {
    // Delegate 直调 (tB 扩展): op(a,b) → ((vb6_del_t_op)(op))(a,b), 独立短路径
    if (node.isDelegateCall) { emitDelegateCall(node); return; }
    // 重载变体 (O2): callee 是标识符且语义层选定非 head 变体时, 把符号键后缀
    // 交给 ident 发射路径消费 (visit(IdentifierExpr) 的过程分支)。
    if (!node.calleeOvlSuffix.empty() && node.callee &&
        node.callee->kind == ASTNodeKind::IdentifierExpr) {
        pendingCalleeOvl_ = node.calleeOvlSuffix;
    }
#include "backend/detail/expr/cgen_expr_call_prelude.inc"
#include "backend/detail/expr/cgen_expr_call_builtin_pre.inc"
#include "backend/detail/expr/cgen_expr_call_callee_ident.inc"
#include "backend/detail/expr/cgen_expr_call_callee_member.inc"
#include "backend/detail/expr/cgen_expr_call_callee_withm.inc"
#include "backend/detail/expr/cgen_expr_call_com_bind.inc"
#include "backend/detail/expr/cgen_expr_call_callee_params.inc"
#include "backend/detail/expr/cgen_expr_call_arg_emit.inc"
#include "backend/detail/expr/cgen_expr_call_arg_variant.inc"
#include "backend/detail/expr/cgen_expr_call_named_args.inc"
#include "backend/detail/expr/cgen_expr_call_builtin_fixup.inc"
#include "backend/detail/expr/cgen_expr_call_opt_pad.inc"
#include "backend/detail/expr/cgen_expr_call_conv_cstr.inc"
#include "backend/detail/expr/cgen_expr_call_conv_numeric.inc"
#include "backend/detail/expr/cgen_expr_call_conv_output.inc"
#include "backend/detail/expr/cgen_expr_call_conv_varargs_lenb.inc"
#include "backend/detail/expr/cgen_expr_call_builtin_argtype.inc"
}

} // namespace vb6c3
