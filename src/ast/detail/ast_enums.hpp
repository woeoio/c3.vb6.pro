#pragma once
// ast_enums.hpp - 第一节 枚举定义 + 第二节 节点类型枚举
// 由 src/ast/ast.hpp 拆出（2026-09-17），内容与原文件对应区间逐字节相同。

// 本文件是 ast 家族的第一个子头，承载原 ast.hpp 的全部 include。

#include "common/diagnostics.hpp"  // SourceLocation
#include "common/types.hpp"        // Vb6Type, CallConv, AccessLevel, ProcKind
#include <string>
#include <vector>
#include <memory>
#include <cstdint>
#include <set>

namespace vb6c3 {

// ============================================================
// 第一节 枚举定义
// ============================================================

// 二元运算符 (与Pratt解析器的14级优先级对应)
enum class BinaryOp : uint8_t {
    // 优先级1: 逻辑或
    Or, Xor,
    // 优先级2: 逻辑与
    And,
    // 优先级4: 逻辑等价
    Eqv, Imp,
    // 优先级5: 比较
    Eq, Neq, Lt, Gt, Le, Ge,
    // 优先级6: 字符串连接
    Concat,       // &
    // 优先级7: 加减
    Add, Sub,
    // 优先级8: 取模
    Mod,
    // 优先级9: 整除
    IntDiv,       // \
    // 优先级10: 乘除
    Mul, Div,
    // 优先级11: 幂 (右结合)
    Pow,          // ^
    // 优先级13: 对象比较
    Like, Is,
    // 移位 (C3 扩展, ai/032): VB.NET 的 << / >>。VB6 没有移位算符, 所以这两枚
    // 追加在枚举末尾 —— 不挪动任何既有取值, 凡是按值比较/建表的地方都不受影响。
    Shl, Shr,
};

// 一元运算符
enum class UnaryOp : uint8_t {
    Negate,   // - (一元负)
    Not,      // Not
};

// 字面量类型
enum class LiteralKind : uint8_t {
    Integer, Long, LongPtr, Single, Double, Currency, Decimal,
    String, Date, Boolean,
    Nothing, Empty, Null,
};

// 语句块类型
enum class ExitKind : uint8_t {
    Do, For, Sub, Function, Property,
};

// On Error 处理类型
enum class OnErrorKind : uint8_t {
    GoToLabel,     // On Error GoTo label
    ResumeNext,    // On Error Resume Next
    GoToZero,      // On Error GoTo 0 (关闭错误处理)
};

// Do Loop 变体
enum class DoLoopKind : uint8_t {
    DoWhileLoop,   // Do While ... Loop
    DoUntilLoop,   // Do Until ... Loop
    DoLoopWhile,   // Do ... Loop While
    DoLoopUntil,   // Do ... Loop Until
    DoLoop,        // Do ... Loop (无条件)
};

// Option 语句类型
enum class OptionKind : uint8_t {
    Explicit,      // Option Explicit
    CompareText,   // Option Compare Text
    CompareBinary, // Option Compare Binary
    BaseZero,      // Option Base 0
    BaseOne,       // Option Base 1
    PrivateModule, // Option Private Module
};

// Open 语句访问模式
enum class OpenMode : uint8_t {
    Input, Output, Append, Binary, Random,
};

// Open 语句访问权限
enum class OpenAccess : uint8_t {
    Read, Write, ReadWrite, Default,
};

// Open 语句锁类型
enum class LockType : uint8_t {
    Shared, LockRead, LockWrite, LockReadWrite, Default,
};

// DefType 语句类型 (DefInt, DefStr等)
enum class DefTypeKind : uint8_t {
    Bool, Byte, Int, Lng, Cur, Sng, Dbl, Date, Str, Obj, Var,
};

// ============================================================
// 第二节 节点类型枚举
// ============================================================

enum class ASTNodeKind : uint16_t {
    // --- 模块 ---
    Module,

    // --- 声明 ---
    SubDecl, FunctionDecl, PropertyDecl,
    TypeDecl, TypeMember,
    EnumDecl, EnumMember,
    DeclareDecl, EventDecl, DelegateDecl,
    InterfaceDecl,  // Interface ... End Interface 契约块 (tB 扩展, ai/022 D1)
    CoClassDecl,    // CoClass ... End CoClass 契约聚合块 (tB 扩展, ai/026 四节, B11/C01)
    ConstDecl, VariableDecl,
    ParameterDecl,
    MultiDecl,  // 逗号分隔的多变量声明 (Dim a, b, c As Long)

    // --- 语句 ---
    Block,
    AssignmentStmt,
    SetStmt,
    LetStmt,
    IfStmt,
    ElseIfClause,
    ForStmt,
    ForEachStmt,
    DoLoopStmt,
    WhileWendStmt,
    SelectCaseStmt,
    CaseClause,
    WithStmt,
    GoToStmt,
    GoSubStmt,
    ReturnStmt,
    OnErrorStmt,
    ResumeStmt,
    ErrorStmt,
    OnGoToStmt,
    OnGoSubStmt,
    MidStmt,  // P18-A: Mid$ statement (assignment)
    ExitStmt,
    StopStmt,
    AsmStmt,   // ai/vb-asm-extension-spec: Asm ... End Asm 内联汇编块
    EndStmt,
    CallStmt,
    ReDimStmt,
    EraseStmt,
    LabelStmt,
    RaiseEventStmt,

    // 文件I/O语句
    OpenStmt, CloseStmt, GetStmt, PutStmt,
    InputStmt, PrintStmt, WriteStmt,
    LineInputStmt, WidthStmt,
    SeekStmt, LockStmt, UnlockStmt, ResetStmt,
    NameStmt,
    FileCopyStmt, KillStmt, MkDirStmt, RmDirStmt,
    ChDirStmt, ChDriveStmt,

    // 杂项语句
    BeepStmt, DoEventsStmt,
    AttributeStmt,
    OptionStmt,
    ImplementsStmt,
    DefTypeStmt,
    LocalDeclStmt,  // 过程体内的局部声明包装 (Dim/Const/Static/Public/Private)

    // --- 表达式 ---
    BinaryExpr,
    UnaryExpr,
    LiteralExpr,
    IdentifierExpr,
    MemberAccessExpr,
    DictionaryAccessExpr,
    IndexOrCallExpr,
    NewExpr,
    TypeOfExpr,
    AddressOfExpr,
    MeExpr,
    WithMemberExpr,

    // --- 类型引用 ---
    SimpleTypeRef,
    ArrayTypeRef,
    FixedStringTypeRef,
};

} // namespace vb6c3
