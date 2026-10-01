#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <cstdlib>
#include <functional>
#include <cstdio>

namespace vb6c3 {

// --- cgen_expr_ident.cpp: IdentifierExpr 求值（内置函数映射、常量折叠、COM/数组/对象分派） ---
// 2026-09-17 拆分：原 744 行单文件（函数体 730 行）拆为
//   本文件（伞文件）                     —— 前置说明 + visit(IdentifierExpr&) 函数骨架
//   detail/cgen_expr_ident_dispatch.inc  —— 标识符分派与 VB6 内建常量（原 13~233 行）
//   detail/cgen_expr_ident_symbol.inc    —— 符号解析：参数/局部变量/窗体属性/枚举/COM全局（原 234~420 行）
//   detail/cgen_expr_ident_builtin.inc   —— 内置函数表 + 用户过程/Property/外部/类/跨模块（原 421~740 行）
// 三个 .inc 是「函数体片段」，在 visit() 函数体内被 #include（C++ 允许），故不用 .cpp/.hpp 后缀 ——
// 它们不是独立编译单元，单独 include 会编译不过。片段局部变量/静态表原样留在片段内，逐行未改 → 零行为改动。

// 内置函数表 (builtinFuncs) + 派生的键集合 (kBuiltinFuncNames): 提到命名空间作用域,
// 供本 TU 的 visit(IdentifierExpr&) 与 cgen_expr_call.cpp 的 visit(IndexOrCallExpr&) 共用。
#include "backend/detail/expr/cgen_expr_ident_builtin_table.inc"

void CCodeGen::visit(IdentifierExpr& node) {
#include "backend/detail/expr/cgen_expr_ident_dispatch.inc"
#include "backend/detail/expr/cgen_expr_ident_symbol.inc"
#include "backend/detail/expr/cgen_expr_ident_builtin.inc"
}

} // namespace vb6c3
