#pragma once
// vb6c3 - C代码生成器
// 将语义分析后的AST+符号表翻译为C代码，交由MSVC编译
// 设计参考: cfront, Nim, Zig早期均采用C代码生成路线
//
// 2026-09-17 拆分：原 996 行单文件拆为
//   本文件（伞头）            —— 前置说明 + CCodeGen 类骨架
//   cgen_emitter.hpp          —— CodeEmitter（代码输出辅助，独立类，原 21~60 行）
//   detail/cgen_api.inc       —— CCodeGen public 段（原 77~275 行）
//   detail/cgen_state.inc     —— CCodeGen private 状态与成员（原 278~615 行）
//   detail/cgen_helpers.inc   —— CCodeGen private 方法声明（原 616~993 行）
// 三个 .inc 是「类体片段」，在类体内被 #include（C++ 允许），故不用 .hpp 后缀 ——
// 它们不是独立头文件，单独 include 会编译不过。
// 对外入口不变：36 处仍写 #include "backend/cgen.hpp"。

#include "backend/cgen_emitter.hpp"
#include "ast/ast.hpp"
#include "ast/ast_visitor.hpp"
#include "semantics/symbol_table.hpp"
#include "semantics/type_system.hpp"
#include "semantics/interfaces_registry.hpp"  // tB Interface 契约发码 (ai/022 B04)
#include "semantics/class_chain_registry.hpp"  // tB 类继承发码 (ai/022 B07b)
#include "common/diagnostics.hpp"
#include "semantics/coclass_identity.hpp"  // CoClass 身份唯一出口 (ai/022 B13b)
#include "project/frm_parser.hpp"
#include "backend/asm_proc.hpp"  // ai/vb-asm-extension-spec: Asm 过程降级元数据
#include <array>
#include <map>
#include <string>
#include <vector>
#include <sstream>
#include <unordered_set>
#include <unordered_map>
#include <set>

namespace vb6c3 {

// ============================================================
// C代码生成器
//
// 输入: Module AST + SemanticAnalyzer的产出(符号表+类型系统)
// 输出: .h头文件 + .c源文件
//
// 架构:
//   1. generate() 主入口 → 生成 .h + .c
//   2. H文件: 类型定义、常量、变量声明、函数前向声明
//   3. C文件: #include ".h"、变量定义、函数实现
//   4. 各emit*方法按节点kind分派
// ============================================================

class CCodeGen : public ASTVisitor {
public:
#include "backend/detail/util/cgen_api.inc"

private:
#include "backend/detail/util/cgen_state.inc"
#include "backend/detail/util/cgen_helpers.inc"
};

} // namespace vb6c3
