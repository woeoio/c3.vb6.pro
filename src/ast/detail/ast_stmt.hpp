#pragma once
// ast_stmt.hpp - 第八节 语句节点（基础 / 控制流 / 错误处理 / 调用）
// 由 src/ast/ast.hpp 拆出（2026-09-17），内容与原文件对应区间逐字节相同。

#include "ast/detail/ast_expr.hpp"

namespace vb6c3 {

// ============================================================
// 第八节 语句节点
// ============================================================

class Stmt : public ASTNode {
public:
    Stmt(ASTNodeKind k, SourceLocation loc) : ASTNode(k, loc) {}
};

// 语句块 (过程体、If体、For体等)
class Block : public Stmt {
public:
    StmtList stmts;

    Block(SourceLocation loc, StmtList s)
        : Stmt(ASTNodeKind::Block, loc), stmts(std::move(s)) {}
};

// 赋值语句: x = expr (隐式Let)
class AssignmentStmt : public Stmt {
public:
    ExprPtr target;   // 左值
    ExprPtr value;    // 右值
    bool isLSet = false;  // P22: LSet statement form
    bool isRSet = false;  // P22: RSet statement form

    AssignmentStmt(SourceLocation loc, ExprPtr t, ExprPtr v)
        : Stmt(ASTNodeKind::AssignmentStmt, loc),
          target(std::move(t)), value(std::move(v)) {}
};

// Set 语句: Set obj = expr (对象赋值)
class SetStmt : public Stmt {
public:
    ExprPtr target;
    ExprPtr value;

    SetStmt(SourceLocation loc, ExprPtr t, ExprPtr v)
        : Stmt(ASTNodeKind::SetStmt, loc),
          target(std::move(t)), value(std::move(v)) {}
};

// Let 语句: Let x = expr (显式Let, 极少使用)
class LetStmt : public Stmt {
public:
    ExprPtr target;
    ExprPtr value;

    LetStmt(SourceLocation loc, ExprPtr t, ExprPtr v)
        : Stmt(ASTNodeKind::LetStmt, loc),
          target(std::move(t)), value(std::move(v)) {}
};

// ElseIf 子句
class ElseIfClause : public Stmt {
public:
    ExprPtr condition;
    StmtList body;

    ElseIfClause(SourceLocation loc, ExprPtr cond, StmtList body)
        : Stmt(ASTNodeKind::ElseIfClause, loc),
          condition(std::move(cond)), body(std::move(body)) {}
};

// If 语句: If...Then...ElseIf...Else...End If
class IfStmt : public Stmt {
public:
    ExprPtr condition;
    StmtList thenBody;
    std::vector<std::unique_ptr<ElseIfClause>> elseIfs;
    StmtList elseBody;
    bool singleLine = false;  // 单行 If...Then...Else

    IfStmt(SourceLocation loc, ExprPtr cond, StmtList then,
           std::vector<std::unique_ptr<ElseIfClause>> elseifs,
           StmtList elseBody, bool single)
        : Stmt(ASTNodeKind::IfStmt, loc),
          condition(std::move(cond)), thenBody(std::move(then)),
          elseIfs(std::move(elseifs)), elseBody(std::move(elseBody)),
          singleLine(single) {}
};

// For 语句: For i = 1 To 10 Step 2
class ForStmt : public Stmt {
public:
    std::string varName;    // 循环变量名
    ExprPtr start;          // 起始值
    ExprPtr end;            // 终止值
    ExprPtr step;           // 步长 (可为nullptr, 默认1)
    StmtList body;

    ForStmt(SourceLocation loc, std::string var, ExprPtr s, ExprPtr e,
            ExprPtr step, StmtList body)
        : Stmt(ASTNodeKind::ForStmt, loc),
          varName(std::move(var)), start(std::move(s)), end(std::move(e)),
          step(std::move(step)), body(std::move(body)) {}
};

// For Each 语句: For Each item In collection
class ForEachStmt : public Stmt {
public:
    std::string varName;
    ExprPtr collection;
    StmtList body;

    ForEachStmt(SourceLocation loc, std::string var, ExprPtr coll, StmtList body)
        : Stmt(ASTNodeKind::ForEachStmt, loc),
          varName(std::move(var)), collection(std::move(coll)), body(std::move(body)) {}
};

// Do Loop 语句 (5种变体)
class DoLoopStmt : public Stmt {
public:
    DoLoopKind loopKind;
    ExprPtr condition;  // 可为nullptr (Do...Loop 无条件变体)
    StmtList body;

    DoLoopStmt(SourceLocation loc, DoLoopKind kind, ExprPtr cond, StmtList body)
        : Stmt(ASTNodeKind::DoLoopStmt, loc),
          loopKind(kind), condition(std::move(cond)), body(std::move(body)) {}
};

// While...Wend 语句 (VB6旧式, 等价于Do While...Loop)
class WhileWendStmt : public Stmt {
public:
    ExprPtr condition;
    StmtList body;

    WhileWendStmt(SourceLocation loc, ExprPtr cond, StmtList body)
        : Stmt(ASTNodeKind::WhileWendStmt, loc),
          condition(std::move(cond)), body(std::move(body)) {}
};

// Case 子句: Case value1, value2, Case Is > 0, Case 1 To 10
class CaseClause : public Stmt {
public:
    // Case 值: 可以是单个值、Is比较、范围或多个值
    struct CaseValue {
        ExprPtr value;       // 单个值、范围的起始值, 或 Is 比较的右操作数
        ExprPtr toValue;     // To 范围的终止值 (可为nullptr)
        // `Case Is > 0`: 比较符存成 relOp, value 只装右操作数 —— VB6 的 Is 站在
        // 测试表达式的位置, 不是标识符 (账 #217; 改动前 parser 造一枚 IdentifierExpr("Is"))
        bool isIsClause = false;
        BinaryOp relOp = BinaryOp::Eq;   // 仅 hasRelOp 为真时有意义
        bool hasRelOp = false;
    };
    std::vector<CaseValue> values;  // Case val1, val2, val3
    StmtList body;

    CaseClause(SourceLocation loc, std::vector<CaseValue> vals, StmtList body)
        : Stmt(ASTNodeKind::CaseClause, loc),
          values(std::move(vals)), body(std::move(body)) {}
};

// Select Case 语句
class SelectCaseStmt : public Stmt {
public:
    ExprPtr testExpr;
    std::vector<std::unique_ptr<CaseClause>> cases;
    StmtList elseCase;  // Case Else

    SelectCaseStmt(SourceLocation loc, ExprPtr test,
                   std::vector<std::unique_ptr<CaseClause>> cases,
                   StmtList elseBody)
        : Stmt(ASTNodeKind::SelectCaseStmt, loc),
          testExpr(std::move(test)), cases(std::move(cases)),
          elseCase(std::move(elseBody)) {}
};

// With 语句: With obj ... End With
class WithStmt : public Stmt {
public:
    ExprPtr object;
    StmtList body;

    WithStmt(SourceLocation loc, ExprPtr obj, StmtList body)
        : Stmt(ASTNodeKind::WithStmt, loc),
          object(std::move(obj)), body(std::move(body)) {}
};

// GoTo 语句
class GoToStmt : public Stmt {
public:
    std::string labelName;

    GoToStmt(SourceLocation loc, std::string label)
        : Stmt(ASTNodeKind::GoToStmt, loc), labelName(std::move(label)) {}
};

// GoSub 语句 (VB6遗留, 类似GoTo但可Return)
class GoSubStmt : public Stmt {
public:
    std::string labelName;

    GoSubStmt(SourceLocation loc, std::string label)
        : Stmt(ASTNodeKind::GoSubStmt, loc), labelName(std::move(label)) {}
};

// Return 语句 (从GoSub返回)
class ReturnStmt : public Stmt {
public:
    ReturnStmt(SourceLocation loc)
        : Stmt(ASTNodeKind::ReturnStmt, loc) {}
};

// On Error 语句
class OnErrorStmt : public Stmt {
public:
    OnErrorKind errorKind;
    std::string labelName;  // GoTo label时使用

    OnErrorStmt(SourceLocation loc, OnErrorKind kind, std::string label = "")
        : Stmt(ASTNodeKind::OnErrorStmt, loc),
          errorKind(kind), labelName(std::move(label)) {}
};

// P14.1.2: Resume语句 — 在错误处理器中恢复执行
enum class ResumeKind : uint8_t {
    ResumeHere,     // Resume (回到出错点)
    ResumeNext,     // Resume Next (跳到出错点下一句)
    ResumeLabel,    // Resume label (跳到指定标签)
};

class ResumeStmt : public Stmt {
public:
    ResumeKind resumeKind;
    std::string labelName;  // ResumeLabel时使用

    ResumeStmt(SourceLocation loc, ResumeKind kind, std::string label = "")
        : Stmt(ASTNodeKind::ResumeStmt, loc),
          resumeKind(kind), labelName(std::move(label)) {}
};

// P14.1.3: Error语句 — 触发运行时错误 (等同 Err.Raise)
class ErrorStmt : public Stmt {
public:
    ExprPtr errorNumber;

    ErrorStmt(SourceLocation loc, ExprPtr errNum)
        : Stmt(ASTNodeKind::ErrorStmt, loc),
          errorNumber(std::move(errNum)) {}
};

// On...GoTo: On x GoTo label1, label2, label3
class OnGoToStmt : public Stmt {
public:
    ExprPtr index;
    std::vector<std::string> labels;

    OnGoToStmt(SourceLocation loc, ExprPtr idx, std::vector<std::string> labels)
        : Stmt(ASTNodeKind::OnGoToStmt, loc),
          index(std::move(idx)), labels(std::move(labels)) {}
};

// On...GoSub: On x GoSub label1, label2, label3
class OnGoSubStmt : public Stmt {
public:
    ExprPtr index;
    std::vector<std::string> labels;

    OnGoSubStmt(SourceLocation loc, ExprPtr idx, std::vector<std::string> labels)
        : Stmt(ASTNodeKind::OnGoSubStmt, loc),
          index(std::move(idx)), labels(std::move(labels)) {}
};

// P18-A: Mid$ statement (assignment) — Mid$(var, start, len) = expr
class MidStmt : public Stmt {
public:
    ExprPtr target;      // target variable (must be a BSTR lvalue)
    ExprPtr start;       // start position (1-based)
    ExprPtr length;      // length (optional, 0 = rest of string from start)
    ExprPtr value;       // replacement string expression
    int32_t hasLength;   // whether length argument was provided

    MidStmt(SourceLocation loc, ExprPtr tgt, ExprPtr s, ExprPtr l, ExprPtr v, int32_t hl)
        : Stmt(ASTNodeKind::MidStmt, loc),
          target(std::move(tgt)), start(std::move(s)), length(std::move(l)),
          value(std::move(v)), hasLength(hl) {}
};

// Exit 语句: Exit Sub/Function/Property/Do/For
class ExitStmt : public Stmt {
public:
    ExitKind exitKind;

    ExitStmt(SourceLocation loc, ExitKind k)
        : Stmt(ASTNodeKind::ExitStmt, loc), exitKind(k) {}
};

// Stop 语句 (进入断点中断)
class StopStmt : public Stmt {
public:
    StopStmt(SourceLocation loc)
        : Stmt(ASTNodeKind::StopStmt, loc) {}
};

// Asm ... End Asm 内联汇编块 (ai/vb-asm-extension-spec)
// v1: 行文本按原始源码直存 (不经 token 重组), 保证 `dword ptr [x]` / `.label:` 等原样;
// 参数→ABI 寄存器替换在发码/驱动侧做文本级替换。naked 为 <Naked> 修饰占位 (v1 未启用)。
class AsmStmt : public Stmt {
public:
    std::vector<std::string> lines;   // 原始汇编行 (不含 Asm / End Asm 两行)
    bool naked = false;               // <Naked> 修饰: 整函数汇编, 不生成 prologue/epilogue
    // `Asm Clobber("rbx","memory")` 声明的被踩寄存器 (小写, 原样收录; "memory" 单独保留)。
    // 与块内静态扫描出的寄存器取并集 → 生成 callee-saved 的 push/pop (x64 MASM / x86 内联)。
    std::vector<std::string> clobbers;

    AsmStmt(SourceLocation loc)
        : Stmt(ASTNodeKind::AsmStmt, loc) {}
};

// End 语句 (终止程序)
class EndStmt : public Stmt {
public:
    EndStmt(SourceLocation loc)
        : Stmt(ASTNodeKind::EndStmt, loc) {}
};

// Call 语句: Call proc(args) 或 proc args
class CallStmt : public Stmt {
public:
    ExprPtr callee;  // 通常是 IdentifierExpr 或 MemberAccessExpr

    CallStmt(SourceLocation loc, ExprPtr callee)
        : Stmt(ASTNodeKind::CallStmt, loc), callee(std::move(callee)) {}
};

// ReDim 语句: ReDim arr(10) / ReDim Preserve arr(20)
class ReDimStmt : public Stmt {
public:
    bool preserve = false;
    std::string varName;
    // ReDim 每维边界
    struct Dimension {
        ExprPtr lower;  // 可为nullptr
        ExprPtr upper;
    };
    std::vector<Dimension> dimensions;
    TypeRefPtr asType;  // As Type (可为nullptr)
    // Fix 100: 「带下标的成员链」目标 (ReDim 语句本体的左值).
    // 仅当目标是 UDT 数组元素的数组成员等复杂形态时非空, 例:
    //   ReDim m_Serie(i).PT(n)  → targetExpr = m_Serie(i).PT, varName = "m_Serie",
    //                             dimensions = {n}
    //   ReDim obj.List(j).Buf(n) → targetExpr = obj.List(j).Buf, varName = "obj.List"
    // 为 nullptr 时沿用 varName 字符串展开 (原路径, 覆盖 ReDim x()/obj.field()/
    // With 块 .field()/ByRef UDT 参数 arr.Field() 等).
    // 后端用 emitExpr(*targetExpr) 发射左值 → VB6_SA_AT(...).Field, 元素类型由
    // 末段成员的 UDT 成员信息解析 (cgen_redim.cpp resolveReDimComplexElemType).
    ExprPtr targetExpr;

    ReDimStmt(SourceLocation loc, bool pres, std::string var,
              std::vector<Dimension> dims, TypeRefPtr type)
        : Stmt(ASTNodeKind::ReDimStmt, loc),
          preserve(pres), varName(std::move(var)),
          dimensions(std::move(dims)), asType(std::move(type)) {}
};

// Erase 语句: Erase arr1, arr2
class EraseStmt : public Stmt {
public:
    std::vector<std::string> varNames;
    // 账 #186 (与 ReDimStmt 的 Fix 100 同一套机制): 下标 + 成员链那种目标
    // (`Erase m_tvFiles(lIndex).bvData`) 用字符串表达不了 —— cIdent 会把 '.' 换成 '_'
    // 且无法带下标。targets[i] 非空 ⇔ 该目标走 emitExpr 发左值
    // (VB6_SA_AT(vb6_type_TFile, m_tvFiles, lIndex).bvData)，varNames[i] 仍存**去下标的点链名**
    // (`m_tvFiles.bvData`)，供 Variant 成员那一问 (isVariantArrayTarget) 复用同一个判据。
    // 两向量按目标一一对应；简单目标 (arr / arr() / obj.Field) 的 targets[i] 为空。
    std::vector<ExprPtr> targets;

    EraseStmt(SourceLocation loc, std::vector<std::string> names)
        : Stmt(ASTNodeKind::EraseStmt, loc), varNames(std::move(names)) {}
};

// 行标签: LabelName:
class LabelStmt : public Stmt {
public:
    std::string labelName;

    LabelStmt(SourceLocation loc, std::string name)
        : Stmt(ASTNodeKind::LabelStmt, loc), labelName(std::move(name)) {}
};

// RaiseEvent 语句: RaiseEvent eventName(arg1, arg2)
class RaiseEventStmt : public Stmt {
public:
    std::string eventName;
    std::vector<ExprPtr> args;

    RaiseEventStmt(SourceLocation loc, std::string name, std::vector<ExprPtr> args)
        : Stmt(ASTNodeKind::RaiseEventStmt, loc),
          eventName(std::move(name)), args(std::move(args)) {}
};

} // namespace vb6c3
