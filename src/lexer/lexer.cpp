#include "lexer/lexer.hpp"
#include <cctype>
#include <algorithm>

namespace vb6c3 {

Lexer::Lexer(std::shared_ptr<SourceBuffer> buffer, Diagnostics& diag,
             uint32_t begin, uint32_t end)
    : buffer_(std::move(buffer))
    , diag_(diag)
{
    // begin/end = 扫描窗口 (ai/028 V2 的插值孔)。默认 begin=0/end=0 ⇒ 整份文件,
    // 存量调用点一字不改。窗口内 offset_ 从 0 数, 但 line_/column_ 用现成的
    // getLocation 播种 ⇒ 窗口里发出的 token 位置直接落在原文件上。
    size_t total = buffer_->content().size();
    if (begin > total) begin = static_cast<uint32_t>(total);
    uint32_t stop = (end == 0 || end < begin || end > total) ? static_cast<uint32_t>(total) : end;
    base_ = begin;
    // 必须先套 string_view 再 substr: std::string::substr 返回**临时 string**, 直接赋给
    // content_ (string_view) 就是悬垂引用 —— 实测所有源文件第一格就"意外字符"。
    // buffer_ 持有这份内容 (shared_ptr), 视图的寿命因此跟着 Lexer 走。
    content_ = std::string_view(buffer_->content()).substr(begin, stop - begin);
    buffer_->getLocation(begin, line_, column_);
    initKeywords();
}

// === 字符操作 ===

char Lexer::peek() const {
    if (offset_ >= content_.size()) return '\0';
    return content_[offset_];
}

char Lexer::peekNext() const {
    if (offset_ + 1 >= content_.size()) return '\0';
    return content_[offset_ + 1];
}

char Lexer::peekNextNext() const {
    if (offset_ + 2 >= content_.size()) return '\0';
    return content_[offset_ + 2];
}

char Lexer::advance() {
    char c = content_[offset_++];
    if (c == '\n') {
        line_++;
        column_ = 1;
    } else {
        column_++;
    }
    return c;
}

bool Lexer::match(char expected) {
    if (offset_ >= content_.size() || content_[offset_] != expected) return false;
    advance();
    return true;
}

bool Lexer::isAlpha(char c) const {
    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == '_';
}

bool Lexer::isDigit(char c) const {
    return c >= '0' && c <= '9';
}

bool Lexer::isHexDigit(char c) const {
    return isDigit(c) || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
}

bool Lexer::isOctDigit(char c) const {
    return c >= '0' && c <= '7';
}

bool Lexer::isBinDigit(char c) const {
    return c == '0' || c == '1';
}

bool Lexer::isAlphaNumeric(char c) const {
    return isAlpha(c) || isDigit(c);
}

bool Lexer::isWhitespace(char c) const {
    return c == ' ' || c == '\t' || c == '\r' || c == '\f' || c == '\v';
}

bool Lexer::isLineContinuation() const {
    // VB6行续接: 行尾 _ (前面有空格, 后面只能是空白+换行)
    if (offset_ >= content_.size() || content_[offset_] != '_') return false;
    // 检查_后面是否只有空白和换行
    size_t pos = offset_ + 1;
    while (pos < content_.size()) {
        char c = content_[pos];
        if (c == '\n') return true;
        if (c == '\r') {
            // \r\n or standalone \r
            return true;
        }
        if (c != ' ' && c != '\t') return false;
        pos++;
    }
    // _ 在文件末尾也算续行
    return pos >= content_.size();
}

// === Token工厂 ===

Token Lexer::makeToken(TokenKind kind, const std::string& text, uint32_t startLine, uint32_t startCol) {
    Token tok;
    tok.kind = kind;
    tok.text = text;
    tok.line = startLine;
    tok.column = startCol;
    tok.length = static_cast<uint32_t>(text.size());
    return tok;
}

Token Lexer::errorToken(const std::string& msg, uint32_t startLine, uint32_t startCol) {
    SourceLocation loc{std::string(buffer_->filePath()), startLine, startCol};
    diag_.error(DiagnosticID::LexUnrecognizedToken, loc, msg);
    return makeToken(TokenKind::Invalid, msg, startLine, startCol);
}

// === 前瞻接口 ===
//
// 唯一的"取下一枚"入口。三处 (nextToken / peekToken / peekToken2) 都必须走它：
// 反引号插值串会把展开剩下的 token 压进 pending_，谁绕过 pending_ 直接 scanToken()
// 谁就把它跳过 —— 那是**漏 token**，比报错难查得多。
Token Lexer::takeScanned() {
    if (!pending_.empty()) {
        Token tok = std::move(pending_.front());
        pending_.pop_front();
        return tok;
    }
    return scanToken();
}

Token Lexer::nextToken() {
    if (!lookahead_.empty()) {
        Token tok = std::move(lookahead_.front());
        lookahead_.pop_front();
        return tok;
    }
    return takeScanned();
}

const Token& Lexer::peekToken() {
    if (lookahead_.empty()) {
        lookahead_.push_back(takeScanned());
    }
    return lookahead_.front();
}

const Token& Lexer::peekToken2() {
    while (lookahead_.size() < 2) {
        lookahead_.push_back(takeScanned());
    }
    return lookahead_[1];
}

// === 核心扫描 ===

Token Lexer::scanToken() {
    // 跳过空白 (不含换行)
    while (offset_ < content_.size() && isWhitespace(peek())) {
        advance();
    }

    // 文件末尾
    if (offset_ >= content_.size()) {
        return makeToken(TokenKind::EndOfFile, "", line_, column_);
    }

    uint32_t startLine = line_;
    uint32_t startCol = column_;
    char c = peek();

    // 换行
    if (c == '\n') {
        advance();
        return makeToken(TokenKind::NewLine, "\n", startLine, startCol);
    }

    // 注释: ' 或 REM
    if (c == '\'') {
        return scanComment();
    }

    // 条件编译: #If, #ElseIf, #Else, #End, #Const
    // 文件号: #1, #2 等
    // 日期字面量: #1/1/2026#
    if (c == '#') {
        advance(); // 消费 #

        // 检查 # 后是否跟纯数字 (文件号场景: As #1, Print #1, 等)
        // 文件号是 # 后跟1-3位数字, 且不是日期模式 (#mm/...#)
        if (isDigit(peek())) {
            // 向前看: 如果数字后面跟 / 或 - 则是日期, 否则是文件号
            size_t saveOff = offset_;
            uint32_t saveCol = column_;
            // 跳过数字
            while (offset_ < content_.size() && isDigit(peek())) {
                advance();
            }
            char afterDigits = peek();
            // 恢复位置
            offset_ = saveOff;
            column_ = saveCol;

            if (afterDigits == '/' || afterDigits == '-') {
                // 日期字面量: 回退 # 由scanDateLiteral处理
                offset_--;
                column_--;
                return scanDateLiteral();
            } else {
                // 文件号: 返回 Hash token, 数字由下次scanToken读取
                return makeToken(TokenKind::Hash, "#", startLine, startCol);
            }
        }

        // 非数字: 检查条件编译指令或返回Hash
        if (isAlpha(peek())) {
            size_t saveOff = offset_;
            uint32_t saveCol = column_;
            std::string text = "#";
            while (offset_ < content_.size() && (isAlpha(peek()) || isDigit(peek()) || peek() == '_')) {
                text += advance();
            }
            std::string lower = text;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            auto it = keywords_.find(lower);
            if (it != keywords_.end()) {
                return makeToken(it->second, text, startLine, startCol);
            }
            // 不是条件编译指令 (#fnum 等文件号引用)
            // 回退到 # 后, 只返回 Hash token, 后续标识符由下次 scanToken 读取
            offset_ = saveOff;
            column_ = saveCol;
            return makeToken(TokenKind::Hash, "#", startLine, startCol);
        }

        // 其他情况: 返回 Hash token (如 # 后跟空格)
        return makeToken(TokenKind::Hash, "#", startLine, startCol);
    }

    // 行续接符 _ (必须在 isAlpha 检查之前, 因为 isAlpha('_') 返回 true,
    // 否则 _ 会被当作标识符首字符, 导致行续接逻辑成为死代码)
    if (c == '_' && isLineContinuation()) {
        return scanLineContinuation();
    }

    // 标识符/关键字
    if (isAlpha(c) || c == '[') {
        if (c == '[') {
            // 方括号标识符 [name with spaces]
            advance(); // 消费 [
            std::string text = "[";
            while (offset_ < content_.size() && peek() != ']') {
                text += advance();
            }
            if (peek() == ']') {
                text += advance(); // 消费 ]
            }
            // 方括号标识符始终作为Identifier
            return makeToken(TokenKind::Identifier, text, startLine, startCol);
        }
        return scanIdentifierOrKeyword();
    }

    // 数字
    if (isDigit(c)) {
        return scanNumber();
    }

    // &H/&O/&B 十六进制/八进制/二进制, & 长整型后缀, & 字符串连接
    if (c == '&') {
        char next = peekNext();
        char lower_next = (next >= 'A' && next <= 'Z') ? (next + 32) : next;
        if (lower_next == 'h') return scanHexNumber();
        if (lower_next == 'o') return scanOctNumber();
        if (lower_next == 'b') return scanBinNumber();
        // 否则是运算符& (字符串连接) 或 Long类型后缀
        return scanOperator();
    }

    // 字符串
    if (c == '"') {
        return scanString();
    }

    // 反引号原始多行串 (C3 扩展, ai/028 V1)。VB6 里 ` 不是任何记号的开头，
    // 走 default 分支必然报"意外字符"，所以把这一格变成特性是纯加法。
    if (c == '`') {
        return scanRawString();
    }

    // 运算符和分隔符
    // C3 扩展 (ai/032): 复合赋值族 (`+=` `-=` …) 与移位族 (`<<` `>>` `<<=` `>>=`)。
    // 都是「符号 + 紧跟 =」的形状, 在各自分支里先看 `=` 再看单符号即可。
    // 存量 VB6 源码里这些相邻组合都不是合法记号 (`a + = b` 本来也解析不过), 所以
    // 整族是纯加法 —— 唯一需要留神的是 `&` (Long 类型后缀), 它的判据在
    // scanIdentifierOrKeyword 的类型后缀那一步: `x& = 5` 的 `&` 在那里就被并进
    // 标识符了, 根本到不了这里; 只有 `x &= 5` 这种 `&` 前有空白的写法才会落到
    // scanOperator, 于是「复合赋值要带空白、`x&=5` 仍是 Long 后缀」成为一条可预测的规则。
    switch (c) {
        case '+':
            advance();
            if (match('=')) return makeToken(TokenKind::PlusEq, "+=", startLine, startCol);
            return makeToken(TokenKind::Plus, "+", startLine, startCol);
        case '-':
            advance();
            if (match('=')) return makeToken(TokenKind::MinusEq, "-=", startLine, startCol);
            return makeToken(TokenKind::Minus, "-", startLine, startCol);
        case '*':
            advance();
            if (match('=')) return makeToken(TokenKind::StarEq, "*=", startLine, startCol);
            return makeToken(TokenKind::Star, "*", startLine, startCol);
        case '/':
            advance();
            if (match('=')) return makeToken(TokenKind::SlashEq, "/=", startLine, startCol);
            return makeToken(TokenKind::Slash, "/", startLine, startCol);
        case '\\':
            advance();
            if (match('=')) return makeToken(TokenKind::BackSlashEq, "\\=", startLine, startCol);
            return makeToken(TokenKind::BackSlash, "\\", startLine, startCol);
        case '^':
            advance();
            if (match('=')) return makeToken(TokenKind::CaretEq, "^=", startLine, startCol);
            return makeToken(TokenKind::Caret, "^", startLine, startCol);
        case '(': advance(); return makeToken(TokenKind::LeftParen, "(", startLine, startCol);
        case ')': advance(); return makeToken(TokenKind::RightParen, ")", startLine, startCol);
        case '.': advance(); return makeToken(TokenKind::Dot, ".", startLine, startCol);
        case ',': advance(); return makeToken(TokenKind::Comma, ",", startLine, startCol);
        case ':':
            advance();
            if (match('=')) return makeToken(TokenKind::Assign, ":=", startLine, startCol);
            return makeToken(TokenKind::Colon, ":", startLine, startCol);
        case ';': advance(); return makeToken(TokenKind::Semicolon, ";", startLine, startCol);
        case '!': advance(); return makeToken(TokenKind::Exclamation, "!", startLine, startCol);
        case '%': advance(); return makeToken(TokenKind::Percent, "%", startLine, startCol);
        case '@': advance(); return makeToken(TokenKind::AtSign, "@", startLine, startCol);
        case '$': advance(); return makeToken(TokenKind::Dollar, "$", startLine, startCol);
        case '=': advance(); return makeToken(TokenKind::Equals, "=", startLine, startCol);
        case '<':
            advance();
            // C3 扩展 (ai/032): << / <<= 。必须在 <> 与 <= 之前判 —— 三者首字符
            // 相同, 先看第二个字符是哪个才能定形。
            if (match('<')) {
                if (match('=')) return makeToken(TokenKind::ShlEq, "<<=", startLine, startCol);
                return makeToken(TokenKind::Shl, "<<", startLine, startCol);
            }
            if (match('>')) return makeToken(TokenKind::NotEquals, "<>", startLine, startCol);
            if (match('=')) return makeToken(TokenKind::LessEqual, "<=", startLine, startCol);
            return makeToken(TokenKind::LessThan, "<", startLine, startCol);
        case '>':
            advance();
            // C3 扩展 (ai/032): >> / >>= 。同上, 先判第二个字符。
            if (match('>')) {
                if (match('=')) return makeToken(TokenKind::ShrEq, ">>=", startLine, startCol);
                return makeToken(TokenKind::Shr, ">>", startLine, startCol);
            }
            if (match('=')) return makeToken(TokenKind::GreaterEqual, ">=", startLine, startCol);
            return makeToken(TokenKind::GreaterThan, ">", startLine, startCol);
        case '_':
            // 行续接符 (必须出现在行尾)
            if (isLineContinuation()) {
                return scanLineContinuation();
            }
            // 否则作为标识符的第一个字符
            return scanIdentifierOrKeyword();
        default:
            advance();
            return errorToken(std::string("意外字符: '") + c + "'", startLine, startCol);
    }
}

// === 标识符/关键字扫描 ===

Token Lexer::scanIdentifierOrKeyword() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;

    while (offset_ < content_.size() && isAlphaNumeric(peek())) {
        text += advance();
    }

    // 可能带有类型后缀: name$, name%, name&, name!, name#, name@
    // 注意: 这些后缀是标识符的一部分, 不单独返回
    if (offset_ < content_.size()) {
        char c = peek();
        if (c == '$' || c == '%' || c == '&' || c == '!' || c == '#' || c == '@') {
            // 检查后缀后面是否是非字母数字 (确认这是类型后缀而非运算符)
            char next = peekNext();
            if (!isAlpha(next) && next != '_' && !isDigit(next)) {
                text += advance(); // 消费类型后缀
            }
        }
    }

    // 查找关键字
    std::string lower = text;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    auto it = keywords_.find(lower);
    if (it != keywords_.end()) {
        // Rem 关键字: 作为注释处理, 消费到行尾 (与 ' 注释行为一致)
        if (it->second == TokenKind::REM_keyword) {
            while (offset_ < content_.size() && peek() != '\n') {
                text += advance();
            }
            return makeToken(TokenKind::Comment, text, startLine, startCol);
        }
        // C3 扩展 (ai/032): `Mod=` 复合赋值。只在 `Mod` 与 `=` **紧邻**时成立
        // (`x Mod= 2`)。带空白的 `x Mod = 2` 维持原样 —— 它与 `If x Mod 2 = 0`
        // 这类比较语法形状太近, 不做空白容忍是刻意的保守选择: 宁可少收一种写法,
        // 也不要把既有语法判错。
        if (it->second == TokenKind::Mod &&
            offset_ < content_.size() && peek() == '=') {
            advance();
            return makeToken(TokenKind::ModEq, text + "=", startLine, startCol);
        }
        return makeToken(it->second, text, startLine, startCol);
    }

    return makeToken(TokenKind::Identifier, text, startLine, startCol);
}

} // namespace vb6c3
