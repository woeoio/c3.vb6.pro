#pragma once
// ast_stmt_io.hpp - 第八节 语句节点（文件 I/O 与其余杂项）
// 由 src/ast/ast.hpp 拆出（2026-09-17），内容与原文件对应区间逐字节相同。

#include "ast/detail/ast_stmt.hpp"

namespace vb6c3 {

// ---- 文件I/O语句 ----

// Open 语句: Open pathname For mode [Access access] [lock] As [#]filenumber [Len=reclength]
class OpenStmt : public Stmt {
public:
    ExprPtr pathName;
    OpenMode mode;
    OpenAccess access;
    LockType lock;
    ExprPtr fileNumber;
    ExprPtr recordLength;  // Len=reclength (可为nullptr)

    OpenStmt(SourceLocation loc, ExprPtr path, OpenMode m, OpenAccess a,
             LockType l, ExprPtr fn, ExprPtr recLen)
        : Stmt(ASTNodeKind::OpenStmt, loc),
          pathName(std::move(path)), mode(m), access(a), lock(l),
          fileNumber(std::move(fn)), recordLength(std::move(recLen)) {}
};

// Close 语句: Close [#]filenumber, [#]filenumber2, ...
class CloseStmt : public Stmt {
public:
    std::vector<ExprPtr> fileNumbers;  // 空列表=Close All

    CloseStmt(SourceLocation loc, std::vector<ExprPtr> fns)
        : Stmt(ASTNodeKind::CloseStmt, loc), fileNumbers(std::move(fns)) {}
};

// Get 语句: Get [#]filenumber, [recnumber], varname
class GetStmt : public Stmt {
public:
    ExprPtr fileNumber;
    ExprPtr recordNumber;  // 可为nullptr (顺序读)
    ExprPtr varName;

    GetStmt(SourceLocation loc, ExprPtr fn, ExprPtr rec, ExprPtr var)
        : Stmt(ASTNodeKind::GetStmt, loc),
          fileNumber(std::move(fn)), recordNumber(std::move(rec)),
          varName(std::move(var)) {}
};

// Put 语句: Put [#]filenumber, [recnumber], varname
class PutStmt : public Stmt {
public:
    ExprPtr fileNumber;
    ExprPtr recordNumber;
    ExprPtr varName;

    PutStmt(SourceLocation loc, ExprPtr fn, ExprPtr rec, ExprPtr var)
        : Stmt(ASTNodeKind::PutStmt, loc),
          fileNumber(std::move(fn)), recordNumber(std::move(rec)),
          varName(std::move(var)) {}
};

// Input 语句: Input #filenumber, varlist
class InputStmt : public Stmt {
public:
    ExprPtr fileNumber;
    std::vector<ExprPtr> varList;

    InputStmt(SourceLocation loc, ExprPtr fn, std::vector<ExprPtr> vars)
        : Stmt(ASTNodeKind::InputStmt, loc),
          fileNumber(std::move(fn)), varList(std::move(vars)) {}
};

// Print 语句: Print #filenumber, [outputlist]
class PrintStmt : public Stmt {
public:
    ExprPtr fileNumber;
    std::vector<ExprPtr> outputList;
    bool isFormPrint = false;  // M22-Issue6: true=Print to form surface, false=Print # to file
    // B11 第二批: 文件 Print 的分隔符, 与 outputList 一一对应 —— 0=语句尾, ';'=不换行,
    // ','=下一个 14 列区。VB6 的 Print # 数值/日期尾部空格与分隔符共同决定写盘字节,
    // 旧 AST 只留表达式、丢掉分隔符 ⇒ 多表达式语句只能"每项一行", 与 VB6 不符。
    std::vector<char> separators;

    PrintStmt(SourceLocation loc, ExprPtr fn, std::vector<ExprPtr> output)
        : Stmt(ASTNodeKind::PrintStmt, loc),
          fileNumber(std::move(fn)), outputList(std::move(output)) {}
};

// Write 语句: Write #filenumber, [outputlist]
class WriteStmt : public Stmt {
public:
    ExprPtr fileNumber;
    std::vector<ExprPtr> outputList;

    WriteStmt(SourceLocation loc, ExprPtr fn, std::vector<ExprPtr> output)
        : Stmt(ASTNodeKind::WriteStmt, loc),
          fileNumber(std::move(fn)), outputList(std::move(output)) {}
};

// Line Input 语句: Line Input #filenumber, varname
class LineInputStmt : public Stmt {
public:
    ExprPtr fileNumber;
    ExprPtr varName;

    LineInputStmt(SourceLocation loc, ExprPtr fn, ExprPtr var)
        : Stmt(ASTNodeKind::LineInputStmt, loc),
          fileNumber(std::move(fn)), varName(std::move(var)) {}
};

// Width 语句: Width #filenumber, width
class WidthStmt : public Stmt {
public:
    ExprPtr fileNumber;
    ExprPtr width;

    WidthStmt(SourceLocation loc, ExprPtr fn, ExprPtr w)
        : Stmt(ASTNodeKind::WidthStmt, loc),
          fileNumber(std::move(fn)), width(std::move(w)) {}
};

// Seek 语句: Seek [#]filenumber, position
class SeekStmt : public Stmt {
public:
    ExprPtr fileNumber;
    ExprPtr position;

    SeekStmt(SourceLocation loc, ExprPtr fn, ExprPtr pos)
        : Stmt(ASTNodeKind::SeekStmt, loc),
          fileNumber(std::move(fn)), position(std::move(pos)) {}
};

// Lock/Unlock 语句
class LockStmt : public Stmt {
public:
    ExprPtr fileNumber;
    ExprPtr start;    // 可为nullptr
    ExprPtr end;      // 可为nullptr

    LockStmt(SourceLocation loc, ExprPtr fn, ExprPtr s, ExprPtr e)
        : Stmt(ASTNodeKind::LockStmt, loc),
          fileNumber(std::move(fn)), start(std::move(s)), end(std::move(e)) {}
};

class UnlockStmt : public Stmt {
public:
    ExprPtr fileNumber;
    ExprPtr start;
    ExprPtr end;

    UnlockStmt(SourceLocation loc, ExprPtr fn, ExprPtr s, ExprPtr e)
        : Stmt(ASTNodeKind::UnlockStmt, loc),
          fileNumber(std::move(fn)), start(std::move(s)), end(std::move(e)) {}
};

// Reset 语句: Reset (关闭所有文件)
class ResetStmt : public Stmt {
public:
    ResetStmt(SourceLocation loc) : Stmt(ASTNodeKind::ResetStmt, loc) {}
};

// Name 语句: Name oldpathname As newpathname
class NameStmt : public Stmt {
public:
    ExprPtr oldPath;
    ExprPtr newPath;

    NameStmt(SourceLocation loc, ExprPtr oldP, ExprPtr newP)
        : Stmt(ASTNodeKind::NameStmt, loc),
          oldPath(std::move(oldP)), newPath(std::move(newP)) {}
};

// 简单文件系统语句 (只有路径参数)
class FileCopyStmt : public Stmt {
public:
    ExprPtr source;
    ExprPtr destination;

    FileCopyStmt(SourceLocation loc, ExprPtr src, ExprPtr dest)
        : Stmt(ASTNodeKind::FileCopyStmt, loc),
          source(std::move(src)), destination(std::move(dest)) {}
};

class KillStmt : public Stmt {
public:
    ExprPtr pathName;

    KillStmt(SourceLocation loc, ExprPtr path)
        : Stmt(ASTNodeKind::KillStmt, loc), pathName(std::move(path)) {}
};

class MkDirStmt : public Stmt {
public:
    ExprPtr pathName;

    MkDirStmt(SourceLocation loc, ExprPtr path)
        : Stmt(ASTNodeKind::MkDirStmt, loc), pathName(std::move(path)) {}
};

class RmDirStmt : public Stmt {
public:
    ExprPtr pathName;

    RmDirStmt(SourceLocation loc, ExprPtr path)
        : Stmt(ASTNodeKind::RmDirStmt, loc), pathName(std::move(path)) {}
};

class ChDirStmt : public Stmt {
public:
    ExprPtr pathName;

    ChDirStmt(SourceLocation loc, ExprPtr path)
        : Stmt(ASTNodeKind::ChDirStmt, loc), pathName(std::move(path)) {}
};

class ChDriveStmt : public Stmt {
public:
    ExprPtr drive;

    ChDriveStmt(SourceLocation loc, ExprPtr d)
        : Stmt(ASTNodeKind::ChDriveStmt, loc), drive(std::move(d)) {}
};

// ---- 杂项语句 ----

class BeepStmt : public Stmt {
public:
    BeepStmt(SourceLocation loc) : Stmt(ASTNodeKind::BeepStmt, loc) {}
};

class DoEventsStmt : public Stmt {
public:
    DoEventsStmt(SourceLocation loc) : Stmt(ASTNodeKind::DoEventsStmt, loc) {}
};

// Attribute 语句: Attribute VB_Name = "ModuleName"
class AttributeStmt : public Stmt {
public:
    std::string attrName;
    ExprPtr value;

    AttributeStmt(SourceLocation loc, std::string name, ExprPtr val)
        : Stmt(ASTNodeKind::AttributeStmt, loc),
          attrName(std::move(name)), value(std::move(val)) {}
};

// Option 语句
class OptionStmt : public Stmt {
public:
    OptionKind optionKind;

    OptionStmt(SourceLocation loc, OptionKind k)
        : Stmt(ASTNodeKind::OptionStmt, loc), optionKind(k) {}
};

// Implements 语句: Implements InterfaceName
//   [tB 扩展 B10] 末尾可带委托子句: `Implements I Via m_holder` —— 本类不再自己写这些
//   成员，契约整份转交给持有字段 m_holder 的对象。无 Via 时 viaField 为空，行为逐字节不变。
class ImplementsStmt : public Stmt {
public:
    std::string interfaceName;
    std::string viaField;  // 空 = 非委托式实现

    ImplementsStmt(SourceLocation loc, std::string name)
        : Stmt(ASTNodeKind::ImplementsStmt, loc), interfaceName(std::move(name)) {}
};

// DefType 语句: DefInt A-C, DefStr D-F
class DefTypeStmt : public Stmt {
public:
    DefTypeKind defKind;
    struct LetterRange {
        char from;
        char to;  // 相同时表示单个字母
    };
    std::vector<LetterRange> ranges;

    DefTypeStmt(SourceLocation loc, DefTypeKind k, std::vector<LetterRange> r)
        : Stmt(ASTNodeKind::DefTypeStmt, loc),
          defKind(k), ranges(std::move(r)) {}
};

// 过程体内的局部声明包装: Dim x As Long, Const k = 0, Static s As String
// VB6 允许声明出现在语句位置, 但 AST 中 Decl 不继承 Stmt
// 使用此包装节点将 Decl 存入 StmtList
class LocalDeclStmt : public Stmt {
public:
    DeclPtr decl;  // 被包装的声明

    LocalDeclStmt(SourceLocation loc, DeclPtr d)
        : Stmt(ASTNodeKind::LocalDeclStmt, loc), decl(std::move(d)) {}
};

} // namespace vb6c3
