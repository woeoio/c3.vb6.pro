#include "parser/parser.hpp"
#include <algorithm>
#include <cctype>

namespace vb6c3 {

// ============================================================
// 构造函数 + 初始化
// ============================================================

Parser::Parser(std::shared_ptr<SourceBuffer> buffer, Diagnostics& diag,
               const PreprocessOptions& ppOpts)
    : preproc_(buffer, diag, ppOpts)
    , diag_(diag)
    , buffer_(buffer)
{
    initBindingPowers();
    // 预读前两个 token (跳过 Comment)
    cur_ = fetchNextToken();
    next_ = fetchNextToken();
}

void Parser::initBindingPowers() {
    // VB6 优先级表 (按绑定力递增)。三大类: 算术 > 比较 > 逻辑 (逻辑最低)。
    // 同级的比较运算符优先级相同, 按源码从左到右计算 (即左结合)。
    //
    // 分组规则 (见 parser_expr.cpp::parseExpression 循环): 对 `a OP1 b OP2 c`,
    // OP2 会并入 OP1 的右操作数 当且仅当 l_bp(OP2) >= r_bp(OP1)。因此
    //   - 左结合要求 r_bp = l_bp + 1 (同级 l_bp 相等, 故不相并入)
    //   - 相邻两级必须满足 下一级 l_bp >= 上一级 r_bp (更紧的运算符要能并入)
    // 这也意味着相邻两级的 l_bp 只能相差 1 或 2, 且 r_bp 必须 >= 1。
    //
    // 级别 | 运算符              | l_bp | r_bp | 结合性
    // -----|--------------------|------|------|-------
    //  1   | Or                 |  2   |  3   | 左
    //  2   | Xor                |  3   |  4   | 左
    //  3   | And                |  4   |  5   | 左
    //      | Not (一元前缀)     |  -   |  5   | 前缀
    //  4   | Eqv, Imp           |  6   |  7   | 左
    //  5   | =, <>, <, >, <=, >=|  8   |  9   | 左
    //  6   | & (Concat)         | 10   | 11   | 左
    //  7   | +, -               | 12   | 13   | 左
    //  8   | Mod                | 14   | 15   | 左
    //  9   | \ (IntDiv)         | 16   | 17   | 左
    // 10   | *, /               | 18   | 19   | 左
    // 11   | ^ (Pow)            | 21   | 20   | 右 (l_bp > r_bp!)
    // 12   | - (一元负)         |  -   | 19   | 前缀
    // 13   | Like, Is           | 22   | 23   | 左
    // 注意: 绑定力的数值不重要, 只需保持相对顺序

    // 中缀运算符 → BindingPower
    struct OpBp { TokenKind kind; int l; int r; };
    OpBp ops[] = {
        // 级1: Or (逻辑运算符中最低)
        { TokenKind::Or,    2,  3 },
        // 级2: Xor — 比 Or 紧。Fix 082: 原先 Or/Xor 同为 (2,3), 导致
        // `a Or b Xor c` 误分组成 (a Or b) Xor c (探针 probe_prec3.bas 实测:
        // `1 Or 2 Xor 3` 输出 ((1|2)^3)=0, 真值 1 Or (2 Xor 3)=1)。
        // 取 (3,4) 使 Xor 的 l_bp(3) >= Or 的 r_bp(3) 而 Xor 的 r_bp(4) 又
        // 等于 And 的 l_bp(4) —— 这样无需改动 And/Not/一元负号等任何其它绑定力。
        { TokenKind::Xor,   3,  4 },
        // 级3: And
        { TokenKind::And,   4,  5 },
        // 级4: Eqv, Imp
        { TokenKind::Eqv,   6,  7 },
        { TokenKind::Imp,   6,  7 },
        // 级5: 比较运算符
        { TokenKind::Equals,      8,  9 },
        { TokenKind::NotEquals,   8,  9 },
        { TokenKind::LessThan,    8,  9 },
        { TokenKind::GreaterThan, 8,  9 },
        { TokenKind::LessEqual,   8,  9 },
        { TokenKind::GreaterEqual,8,  9 },
        // 级6: 字符串连接
        { TokenKind::Ampersand, 10, 11 },
        // 级7: 加减
        { TokenKind::Plus,   12, 13 },
        { TokenKind::Minus,  12, 13 },
        // 级8: Mod
        { TokenKind::Mod,    14, 15 },
        // 级9: 整除
        { TokenKind::BackSlash, 16, 17 },
        // 级10: 乘除
        { TokenKind::Star,   18, 19 },
        { TokenKind::Slash,  18, 19 },
        // 级11: 幂 (右结合: l_bp=21 > r_bp=20)
        { TokenKind::Caret,  21, 20 },
        // 级13: Like, Is (Is在词法器中输出为IsKeyword，两者都需要注册)
        { TokenKind::Like,     22, 23 },
        { TokenKind::Is,       22, 23 },
        { TokenKind::IsKeyword,22, 23 },
        // C3 扩展 (ai/032): 移位 (9,10) —— VB.NET 的优先级位置。
        // 取 (9,10) 的依据是两侧夹逼, 不是拍脑袋:
        //   l_bp=9  > 比较的 l_bp=8  ⇒ `a << b = c` 分组为 `(a << b) = c` (VB.NET 同)
        //   l_bp=9  < `&` 的 l_bp=10 ⇒ `a & b << c` 分组为 `(a & b) << c`
        //   r_bp=10 > 加减的 l_bp=12? 否 —— 12>10, 所以 `a << b + c` 的右操作数会
        //             继续吃掉 `+ c` ⇒ `a << (b + c)`, 与 VB.NET 的「算术高于移位」
        //             一致 (C 里也是 `a << b + c` == `a << (b + c)`)。
        { TokenKind::Shl,      9, 10 },
        { TokenKind::Shr,      9, 10 },
        // C3 扩展 (ai/032): IsNot 与 Is 同优先级 —— 它就是 Is 的否定形, 绑定范围
        // 必须完全一致, 否则 `Not obj IsNot x` 这类嵌套会分组不同。
        { TokenKind::IsNot,   22, 23 },
    };

    for (auto& op : ops) {
        bpTable_[static_cast<int>(op.kind)] = {op.l, op.r};
    }
}

// ============================================================
// Token 消费接口
// ============================================================

const Token& Parser::peek() const {
    return cur_;
}

const Token& Parser::peek2() const {
    return next_;
}

Token Parser::advance() {
    if (++advanceCount_ > MAX_ADVANCES) {
        // 安全限制: 防止无限循环消耗内存
        SourceLocation loc = currentLoc();
        diag_.error(DiagnosticID::ParseUnexpectedToken, loc,
            "解析器advance调用超过上限(" + std::to_string(MAX_ADVANCES) + "), 可能存在无限循环");
        cur_.kind = TokenKind::EndOfFile;
        return cur_;
    }
    Token tok = std::move(cur_);
    cur_ = std::move(next_);
    next_ = fetchNextToken();
    prevTok_ = tok;  // Fix 043c: 保存已消费的 token 用于空格检测
    return tok;
}

Token Parser::fetchNextToken() {
    Token tok;
    int safetyCounter = 0;
    do {
        tok = preproc_.nextToken();
        if (++safetyCounter > 1000000) {
            // 安全限制: 防止无限循环
            tok.kind = TokenKind::EndOfFile;
            break;
        }
    } while (tok.kind == TokenKind::Comment || tok.kind == TokenKind::LineContinuation);
    return tok;
}

Token Parser::expect(TokenKind kind, DiagnosticID diagId, const std::string& msg) {
    if (cur_.kind == kind) {
        return advance();
    }
    diag_.error(diagId, currentLoc(),
        msg + " (got " + std::string(Token::kindToString(cur_.kind)) + ")");
    // 不消费, 让调用者决定后续动作
    return Token{cur_.kind, cur_.text, cur_.line, cur_.column, cur_.length, {0}};
}

Token Parser::expect(TokenKind kind, const std::string& msg) {
    return expect(kind, DiagnosticID::ParseExpectedToken, msg);
}

bool Parser::match(TokenKind kind) {
    if (cur_.kind == kind) {
        advance();
        return true;
    }
    return false;
}

bool Parser::expectOrSkip(TokenKind kind, DiagnosticID diagId, const std::string& msg) {
    if (cur_.kind == kind) {
        advance();
        return true;
    }
    diag_.error(diagId, currentLoc(),
        msg + " (got " + std::string(Token::kindToString(cur_.kind)) + ")");
    return false;
}

bool Parser::check(TokenKind kind) const {
    return cur_.kind == kind;
}

bool Parser::checkAny(std::initializer_list<TokenKind> kinds) const {
    for (auto k : kinds) {
        if (cur_.kind == k) return true;
    }
    return false;
}

bool Parser::isStatementStart() const {
    switch (cur_.kind) {
        // 赋值/调用: 标识符、.member(with)、!dict
        case TokenKind::Identifier:
        case TokenKind::Dot:
        case TokenKind::Exclamation:
        // 块语句关键字
        case TokenKind::If:
        case TokenKind::For:
        case TokenKind::Do:
        case TokenKind::While:
        case TokenKind::Select:
        case TokenKind::With:
        // ai/vb-asm-extension-spec: Asm 块起始
        case TokenKind::Asm:
        // 跳转
        case TokenKind::GoTo:
        case TokenKind::GoSub:
        case TokenKind::Return:
        case TokenKind::On:
        case TokenKind::Exit:
        case TokenKind::Stop:
        case TokenKind::End:
        // 赋值关键字
        case TokenKind::Set:
        case TokenKind::Let:
        case TokenKind::Call:
        // 声明 (Dim 也可出现在过程体内)
        case TokenKind::Dim:
        case TokenKind::ReDim:
        case TokenKind::Const:
        case TokenKind::Static:
        case TokenKind::Public:
        case TokenKind::Private:
        // 声明关键字在过程体内也可出现
        case TokenKind::Sub:
        case TokenKind::Function:
        case TokenKind::Property:
        case TokenKind::Type:
        case TokenKind::Enum:
        case TokenKind::Declare:
        case TokenKind::Event:
        // 文件I/O
        case TokenKind::Open:
        case TokenKind::Close:
        case TokenKind::Get:
        case TokenKind::Put:
        case TokenKind::Input:
        case TokenKind::Output:
        case TokenKind::Print:
        case TokenKind::Write:
        case TokenKind::Seek:
        case TokenKind::Lock:
        case TokenKind::Unlock:
        case TokenKind::Name:
        case TokenKind::Line:
        case TokenKind::Width:
        case TokenKind::FileCopy:
        case TokenKind::Kill:
        case TokenKind::MkDir:
        case TokenKind::RmDir:
        case TokenKind::ChDir:
        case TokenKind::ChDrive:
        // 杂项
        case TokenKind::Beep:
        case TokenKind::DoEvents:
        case TokenKind::Attribute:
        case TokenKind::Implements:
        case TokenKind::RaiseEvent:
        // Option (过程体内极少, 但解析器应容错)
        case TokenKind::Option:
        // DefType
        case TokenKind::DefBool:
        case TokenKind::DefByte:
        case TokenKind::DefInt:
        case TokenKind::DefLng:
        case TokenKind::DefCur:
        case TokenKind::DefSng:
        case TokenKind::DefDbl:
        case TokenKind::DefDate:
        case TokenKind::DefStr:
        case TokenKind::DefObj:
        case TokenKind::DefVar:
        // Me
        case TokenKind::MeKeyword:
            return true;
        default:
            // 软关键字也可作为语句开头 (变量名赋值/调用)
            return isSoftKeyword(cur_.kind);
    }
}

bool Parser::isDeclarationStart() const {
    switch (cur_.kind) {
        case TokenKind::Dim:
        case TokenKind::Const:
        case TokenKind::Public:
        case TokenKind::Private:
        case TokenKind::Static:
        case TokenKind::Sub:
        case TokenKind::Function:
        case TokenKind::Property:
        case TokenKind::Type:
        case TokenKind::Enum:
        case TokenKind::Declare:
        case TokenKind::Event:
        case TokenKind::Friend:
        case TokenKind::Protected:  // tB 扩展 (B08a)
        case TokenKind::Overridable:       // tB 扩展 (B08b): 虚方法修饰符可起一行声明
        case TokenKind::Overrides:
        case TokenKind::NotOverridable:
        case TokenKind::Global:
            return true;
        default:
            // ai/024: `DeclareWide` (tB 兼容) —— 关键词表是精确匹配, "declarewide"
            // 不命中 "declare", 所以它以 Identifier 形态出现。这里放行, 具体解析
            // 交给 parseDeclaration() (与 `Declare` 共用 parseDeclareDecl)。
            // 判定收得很紧 (必须整词等于 declarewide), 不影响其它标识符开头的行
            // —— 那类行仍旧落回 "unexpected token at module level"。
            if (cur_.kind == TokenKind::Identifier && toLower(cur_.text) == "declarewide") {
                return true;
            }
            return false;
    }
}

// ============================================================
// 新行处理
// ============================================================

void Parser::skipNewLines() {
    while (cur_.kind == TokenKind::NewLine || cur_.kind == TokenKind::Comment || cur_.kind == TokenKind::LineContinuation) {
        advance();
    }
}

void Parser::skipStatementSeparators() {
    // 冒号与换行在 VB6 里都是语句分隔符。块语句 header 之后调用本函数即可
    // 无差别接受 `Header: body: Terminator` 单行形式, 无需每个块语句各写一份分支。
    while (cur_.kind == TokenKind::NewLine || cur_.kind == TokenKind::Comment ||
           cur_.kind == TokenKind::LineContinuation || cur_.kind == TokenKind::Colon) {
        advance();
    }
}

bool Parser::expectEndOfStatement() {
    // 接受 NewLine 或冒号
    if (cur_.kind == TokenKind::NewLine) {
        skipNewLines();
        return true;
    }
    if (match(TokenKind::Colon)) {
        skipNewLines();
        return true;
    }
    // 如果下一个是 EndOfFile, 也可以
    if (cur_.kind == TokenKind::EndOfFile) {
        return true;
    }
    // Fix 231: **块结构 token 本身就是合法的语句边界** —— 上一条语句到此为止,
    // 调用方 (块体的 parseBlockUntil) 下一轮会自己识别并 return。不许在这里报
    // "expected end of statement" 再把它 advance() 掉: 一吃就会把紧随的
    // `Sub`/`If`/`With`/语句 当新的语句解析, 级联出一整片 VB2001..VB2006。
    //
    // 为什么块 token 会"残留"到 expectEndOfStatement 面前: `Next var1, var2, …`
    // 一条 Next 关多层嵌套 For (内层在前) —— 那个唯一的 Next token 由最内层消费,
    // 逗号后其余变量排进 pendingNextVars_ 当各外层的"合成 Next" (见
    // consumeNextClause)。外层 body 的 parseBlockUntil 凭 pendingNextVars_ 在
    // **下一个块 token 处**收尾 —— 而 ForStmt 返回时 cur_ 就停在那枚块 token 上
    // (合成弹出不消费 token), 于是调用方这里看到它。
    // 实测 (tests/VBFlexGridDemo/Builds/VBFlexGrid/VBFlexGrid.ctl):
    //   6783  For iCol = 0 To …: For iRow = LBoundRows To UBoundRows
    //         …
    //   6789  Next iRow, iCol      ' 关两层
    //   6790  Else                 ' ← 残留给调用方的就是这枚 Else (层 #1)
    // 以及 6769 的 `End If` (层 #2, `Next` 之后直接收过程)。
    switch (cur_.kind) {
        case TokenKind::Else:
        case TokenKind::ElseIf:
        case TokenKind::Case:
        case TokenKind::Loop:
        case TokenKind::Next:
        case TokenKind::Wend:
            return true;
        default:
            break;
    }
    if (cur_.kind == TokenKind::End && isEndBlock()) {
        return true;
    }
    // 报错并跳过当前token, 防止无限循环
    diag_.error(DiagnosticID::ParseExpectedEndOfStatement, currentLoc(),
        "expected end of statement (newline or :)");
    advance();  // 必须前进, 否则调用方while循环可能死循环
    return false;
}

SourceLocation Parser::currentLoc() const {
    return SourceLocation{buffer_->filePath(), cur_.line, cur_.column};
}

// ============================================================
// 错误恢复
// ============================================================

void Parser::synchronize() {
    // 跳过 token 直到遇到语句开始或块终止
    while (!isStatementStart() &&
           cur_.kind != TokenKind::EndOfFile &&
           cur_.kind != TokenKind::End &&
           cur_.kind != TokenKind::Else &&
           cur_.kind != TokenKind::ElseIf &&
           cur_.kind != TokenKind::Loop &&
           cur_.kind != TokenKind::Wend &&
           cur_.kind != TokenKind::Next &&
           cur_.kind != TokenKind::Case) {
        advance();
    }
}

void Parser::skipToNextLine() {
    while (cur_.kind != TokenKind::NewLine &&
           cur_.kind != TokenKind::EndOfFile) {
        advance();
    }
    skipNewLines();
}

} // namespace vb6c3
