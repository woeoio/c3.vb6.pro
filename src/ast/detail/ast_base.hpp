#pragma once
// ast_base.hpp - 第五节 AST 基类 + 第六节 类型引用节点
// 由 src/ast/ast.hpp 拆出（2026-09-17），内容与原文件对应区间逐字节相同。

#include "ast/detail/ast_fwd.hpp"
#include <cstdint>

namespace vb6c3 {

// ============================================================
// 第五节 AST 基类
// ============================================================

class ASTNode {
public:
    ASTNodeKind kind;
    SourceLocation loc;

    // ai/030 续 (§性能, 2026-10-14): inferExprType 的按节点缓存槽。
    // epoch 是**全局单调**计数 (CCodeGen::bumpInferEpoch): 过程入口与每条语句边界
    // 都推它 —— 任何会改推断答案的状态变更 (known*Vars_ 的体内注册 / With 栈 /
    // currentProc_) 都发生在这些边界上。0 = 未缓存。放基类让表达式全树可用
    // (语句节点多占 8 字节, 无碍)。
    uint32_t inferCacheEpoch = 0;
    uint16_t inferCacheType = 0;

    ASTNode(ASTNodeKind k, SourceLocation l) : kind(k), loc(l) {}
    virtual ~ASTNode() = default;

    // 禁止拷贝，允许移动
    ASTNode(const ASTNode&) = delete;
    ASTNode& operator=(const ASTNode&) = delete;
    ASTNode(ASTNode&&) = default;
    ASTNode& operator=(ASTNode&&) = default;

    const char* kindName() const;
};

// ============================================================
// 第六节 类型引用节点
// ============================================================

// 简单类型引用: Long, String, MyClass 等
class SimpleTypeRef : public ASTNode {
public:
    std::string name;  // 类型名 (不区分大小写)

    SimpleTypeRef(SourceLocation loc, std::string n)
        : ASTNode(ASTNodeKind::SimpleTypeRef, loc), name(std::move(n)) {}
};

// 数组类型引用: Long(), String(10), Integer(1 To 100)
class ArrayTypeRef : public ASTNode {
public:
    std::unique_ptr<ASTNode> elementType;  // 元素类型 (SimpleTypeRef 或嵌套 ArrayTypeRef)

    // 数组维度: 每维为 (lower, upper) 对; lower 默认由 Option Base 决定
    struct Dimension {
        ExprPtr lower;  // 可为 nullptr (使用 Option Base)
        ExprPtr upper;
    };
    std::vector<Dimension> dimensions;

    ArrayTypeRef(SourceLocation loc, std::unique_ptr<ASTNode> elemType,
                 std::vector<Dimension> dims)
        : ASTNode(ASTNodeKind::ArrayTypeRef, loc),
          elementType(std::move(elemType)), dimensions(std::move(dims)) {}
};

// 定长字符串类型引用: String * N
class FixedStringTypeRef : public ASTNode {
public:
    ExprPtr length;  // N 表达式

    FixedStringTypeRef(SourceLocation loc, ExprPtr len)
        : ASTNode(ASTNodeKind::FixedStringTypeRef, loc), length(std::move(len)) {}
};

using TypeRefPtr = std::unique_ptr<ASTNode>;

} // namespace vb6c3
