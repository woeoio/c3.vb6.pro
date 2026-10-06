#pragma once
// ast_expr.hpp - 第七节 表达式节点
// 由 src/ast/ast.hpp 拆出（2026-09-17），内容与原文件对应区间逐字节相同。

#include "ast/detail/ast_base.hpp"

namespace vb6c3 {

// ============================================================
// 第七节 表达式节点
// ============================================================

class Expr : public ASTNode {
public:
    Expr(ASTNodeKind k, SourceLocation loc) : ASTNode(k, loc) {}
};

// 二元表达式: a + b, x And y 等
class BinaryExpr : public Expr {
public:
    BinaryOp op;
    ExprPtr left;
    ExprPtr right;

    BinaryExpr(SourceLocation loc, BinaryOp op, ExprPtr l, ExprPtr r)
        : Expr(ASTNodeKind::BinaryExpr, loc),
          op(op), left(std::move(l)), right(std::move(r)) {}
};

// 一元表达式: -x, Not y
class UnaryExpr : public Expr {
public:
    UnaryOp op;
    ExprPtr operand;

    UnaryExpr(SourceLocation loc, UnaryOp op, ExprPtr operand)
        : Expr(ASTNodeKind::UnaryExpr, loc),
          op(op), operand(std::move(operand)) {}
};

// 字面量: 42, 3.14#, "hello", #1/1/2026#, True, Nothing 等
class LiteralExpr : public Expr {
public:
    LiteralKind literalKind;
    std::string rawText;  // 原始文本 (保留用户写法)
    union {
        int32_t intValue = 0;
        int64_t longValue;
        float floatValue;
        double doubleValue;
        bool boolValue;
    };

    LiteralExpr(SourceLocation loc, LiteralKind k, std::string raw)
        : Expr(ASTNodeKind::LiteralExpr, loc),
          // 账 #172: 联合体**整体**清零。原来只写 intValue(0) —— 那清了 4 个字节，
          // 而 Date/Double 走的是 8 字节的 doubleValue，高 4 字节从没人写过
          // ⇒ Debug 下恰好读到 0、Release/带 /RTCu 的构建读到垃圾（#1/1/1900# 发成
          // -6.277e+66）。清零不等于「算对了」，值要由 parser 折进来。
          literalKind(k), rawText(std::move(raw)), doubleValue(0) {}
};

// 标识符: x, MyVar, [带空格的名称]
class IdentifierExpr : public Expr {
public:
    std::string name;        // 标识符名称 (保留原始大小写)
    bool bracketed = false;  // 是否使用方括号 [name]

    IdentifierExpr(SourceLocation loc, std::string n, bool br = false)
        : Expr(ASTNodeKind::IdentifierExpr, loc),
          name(std::move(n)), bracketed(br) {}
};

// 成员访问: obj.member
class MemberAccessExpr : public Expr {
public:
    ExprPtr object;
    std::string memberName;

    MemberAccessExpr(SourceLocation loc, ExprPtr obj, std::string member)
        : Expr(ASTNodeKind::MemberAccessExpr, loc),
          object(std::move(obj)), memberName(std::move(member)) {}
};

// 字典访问: obj!key (等价于 obj.Fields("key"))
class DictionaryAccessExpr : public Expr {
public:
    ExprPtr object;
    std::string key;

    DictionaryAccessExpr(SourceLocation loc, ExprPtr obj, std::string k)
        : Expr(ASTNodeKind::DictionaryAccessExpr, loc),
          object(std::move(obj)), key(std::move(k)) {}
};

// 命名参数
struct NamedArg {
    std::string name;   // 参数名
    ExprPtr value;      // 参数值
};

// 索引/调用: arr(i), func(x, y), call with named args
// VB6 在解析阶段不区分数组索引和函数调用, 统一为此节点
class IndexOrCallExpr : public Expr {
public:
    ExprPtr callee;                    // 被调用者 (IdentifierExpr 或 MemberAccessExpr)
    std::vector<ExprPtr> positional;   // 位置参数
    std::vector<NamedArg> named;       // 命名参数
    std::set<size_t> byvalOverrides;  // Fix 072: ByVal 覆盖的参数索引 (VB6 允许 ByVal x 覆盖 ByRef 声明)
    std::set<size_t> omittedArgs;     // Fix 142: 省略的位置实参索引 (VB6 a, , b) — 保留槽位, 代码生成按缺省值/IsMissing 处理
    // Delegate 直调: op(5, 6) 中 op 是 As <Delegate> 变量. 语义层识别并标记,
    // cgen 按委托签名的函数指针强转调用 (isDelegateCall=true 时 delegateTypeName 有效).
    bool isDelegateCall = false;
    std::string delegateTypeName;
    // 重载选择结果 (O2): 语义层 resolveOverload 选定的变体键后缀
    // ("" = 裸键 head / 非重载, 零改动; 否则 "$ov$<fp>", cgen 据此定形 C 名与形参表).
    std::string calleeOvlSuffix;

    IndexOrCallExpr(SourceLocation loc, ExprPtr callee)
        : Expr(ASTNodeKind::IndexOrCallExpr, loc),
          callee(std::move(callee)) {}
};

// New 表达式: New ClassName
class NewExpr : public Expr {
public:
    std::string className;
    // C3 扩展 (084c): `New Cls(a, b)` 带参构造 — VB6 本体不允许 `New X(...)`，
    // 仅本工程类 (Class_Initialize 带形参) 生效; COM/外部类带实参在语义层报错。
    std::vector<ExprPtr> args;

    NewExpr(SourceLocation loc, std::string cls)
        : Expr(ASTNodeKind::NewExpr, loc), className(std::move(cls)) {}
};

// TypeOf 表达式: TypeOf obj Is ClassName
class TypeOfExpr : public Expr {
public:
    ExprPtr object;
    std::string typeName;

    TypeOfExpr(SourceLocation loc, ExprPtr obj, std::string typeName)
        : Expr(ASTNodeKind::TypeOfExpr, loc),
          object(std::move(obj)), typeName(std::move(typeName)) {}
};

// AddressOf 表达式: AddressOf funcname (用于回调)
class AddressOfExpr : public Expr {
public:
    std::string funcName;
    // 委托绑定标记: 当赋值/初始化目标为 As <Delegate> 变量且签名校验通过时,
    // 语义层填入委托类型名; cgen 据此生成/引用调用桩而非裸过程地址.
    std::string delegateTypeName;
    // 重载目标选择结果 (O2): 目标过程名所在组的变体键后缀 ("" = 裸键/非重载).
    std::string funcOvlSuffix;

    AddressOfExpr(SourceLocation loc, std::string fn)
        : Expr(ASTNodeKind::AddressOfExpr, loc), funcName(std::move(fn)) {}
};

// Me 表达式: Me (当前对象实例)
class MeExpr : public Expr {
public:
    MeExpr(SourceLocation loc)
        : Expr(ASTNodeKind::MeExpr, loc) {}
};

// With 块中的隐式成员访问: .Property (等价于 WithVar.Property)
class WithMemberExpr : public Expr {
public:
    std::string memberName;

    WithMemberExpr(SourceLocation loc, std::string member)
        : Expr(ASTNodeKind::WithMemberExpr, loc), memberName(std::move(member)) {}
};

} // namespace vb6c3
