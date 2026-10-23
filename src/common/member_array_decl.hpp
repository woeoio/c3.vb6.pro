#pragma once
#include "ast/detail/ast_fwd.hpp"   // vb6c3::ASTNode 的前向声明 (与全仓同一类型)
#include <unordered_map>
#include <string>

// 账 #476 §B190 (第一百九十刀): 「类/窗体模块的**动态数组成员**声明」的**运输**结构。
//
// 为什么需要它: 发码层拿不到「类名 → Module*」—— `clsreg_` (ClassChainRegistry) 在工程里
// 没有 Inherits / 虚修饰符 / Protected 成员时**根本不建表** (runClassChainPrepass 的早退),
// 于是独立类 (语料里绝大多数) 在发码侧查不到自己的模块。而 `classFieldKindMap_` 那张表
// (§B168/§B169) 刻意**不收数组字段** (`var.isDynamicArray` 当场 return Unknown)。
// ⇒ 类字段那一腿以前在发码侧没有任何答案, §B190 的 11 处 C2064 就是这么来的。
//
// 口径 (用户拍板): 这一份只运**语法** (`As <T>` 那一个类型引用 + 秩), **不运档位** ——
// 档位到 C 元素类型的映射一律由发码层那一枚出口 (CCodeGen::resolveMemberArrayDecl) 用
// `resolveArrayElemType` / `resolveArrayUdtElemCType` 做, 与发射类结构体字段时
// (cgen_base_generate_c_open.inc) **同一枚映射**。运输与映射分开, 才不会出现
// 「分配侧一份答案、访问侧另一份答案」(§B154 / §B220 那个病)。
//
// AST 生命期: Module AST 在整个 runCodeGeneration 期间存活 (driver 持有 unique_ptr),
// 所以这里存裸指针是安全的; 与 classVoidFieldMap_/classTypedFieldMap_ 存 C 类型文本同源。
// ⚠ 类型一定写 `vb6c3::ASTNode` (由 ast/detail/ast_fwd.hpp 给): 全局前向声明会造出
// 第二个不同类型, 发码侧赋值时 cl 报 C2440 (本分分实测)。

struct MemberArraySyntax {
    vb6c3::ASTNode* asType = nullptr;   // `Field() As <T>` 的 T; 空 = 没写 As (VB6 = Variant)
    int      rank   = 0;         // 0 = 一维
};

// 类名(小写) → 成员名(小写) → 声明语法
using ClassMemberArrayMap =
    std::unordered_map<std::string, std::unordered_map<std::string, MemberArraySyntax>>;
