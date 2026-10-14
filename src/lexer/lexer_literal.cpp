#include "lexer/lexer.hpp"

namespace vb6c3 {

// === 字符串扫描 ===

Token Lexer::scanString() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;
    text += advance(); // 消费开头 "

    while (offset_ < content_.size() && peek() != '"') {
        if (peek() == '\n') {
            // VB6不允许字符串跨行 (除非行续接)
            SourceLocation loc{std::string(buffer_->filePath()), startLine, startCol};
            diag_.error(DiagnosticID::LexUnterminatedString, loc, "字符串未闭合");
            break;
        }
        text += advance();
    }

    if (peek() == '"') {
        text += advance(); // 消费结尾 "
        // VB6中 """ 表示字符串内的一个引号
        // 如果紧跟另一个", 说明是转义引号, 继续扫描
        while (offset_ < content_.size() && peek() == '"') {
            text += advance(); // 消费转义 "
            while (offset_ < content_.size() && peek() != '"') {
                text += advance();
            }
            if (offset_ < content_.size() && peek() == '"') {
                text += advance();
            }
        }
    }

    return makeToken(TokenKind::StringLiteral, text, startLine, startCol);
}

// === 反引号原始串扫描出口 (C3 扩展, ai/028 V1 + V2) ===
//
// 出口只有两种形态，且**都不是**新 token：
//   V1（串里没有 ${）: 归一成普通 StringLiteral —— 行界写成 CRLF、内嵌的 " 重新双写、
//     双写的反引号折成一枚 ⇒ token.text 与手写的 VB6 串逐字节同形，parser / semantics /
//     cgen 以及全仓 rawText 消费点 (ai/028 三-3 那四处各自剥引号折 "") 一行都不用改。
//   V2（串里有 ${expr}）: 整串展开成一串**普通 token** —— 文本段 = StringLiteral、连接 = &、
//     孔 = CStr(...) 或 Format$(..., "fmt")，外面再套一对括号；第一枚交回、其余压进 pending_。
//     于是"第二种字符串"从来不存在，未声明变量检查、按实参类型挑 vb6_CStrLong、COM 默认属性
//     解析三件事全部自动继承 (计划书 R2/R4)。
Token Lexer::scanRawString() {
    uint32_t startLine = line_, startCol = column_;
    advance(); // 消费开头的那枚反引号
    // 未闭合时只吃掉开头这一枚就退回。不退回的话整份余文会被吞进一个假字面量，
    // 一条 1007 会连带变成一片跨阶段的怪错。
    uint32_t bodyOff = offset_;
    uint32_t bodyLine = line_, bodyCol = column_;

    std::vector<RawSeg> segs;
    std::vector<RawHole> holes;
    RawSeg cur;
    bool atOpen = true;   // 起始反引号后紧跟的那一个换行按文本块惯例裁掉 (ai/028 八-1)
    bool closed = false;
    bool holeBad = false; // 孔自己报过 1008/1009 ⇒ 不要再补一条 1007

    // 每段文本记下自己的起点: V2 展开后各段是各自独立的字面量 token, 位置要指对自己那一格。
    auto markSeg = [&]() {
        if (cur.text.empty()) { cur.line = line_; cur.col = column_; }
    };

    while (offset_ < content_.size()) {
        char c = peek();
        if (c == '`') {
            if (peekNext() == '`') {   // 双写 → 一枚字面反引号 (与 VB6 的 "" 同一心智)
                markSeg();
                advance();
                advance();
                cur.text += '`';
                atOpen = false;
                continue;
            }
            advance();                 // 未成对的那枚 = 闭合
            closed = true;
            break;
        }
        // 字面 `${` 的出口是 `$${`。只在紧跟 `{` 时才折 ⇒ 文本里普通的 `$`、`$$` 不受影响。
        if (c == '$' && peekNext() == '$' && offset_ + 2 < content_.size()
            && content_[offset_ + 2] == '{') {
            markSeg();
            advance();
            advance();
            advance();
            cur.text += "${";
            atOpen = false;
            continue;
        }
        // `${` = 开孔。右界与顶层 `:` 由 scanHoleEnd 定, 之后按绝对偏移开窗子扫描。
        if (c == '$' && peekNext() == '{') {
            RawHole hole;
            hole.line = line_;
            hole.col = column_;
            markSeg();
            segs.push_back(cur);
            cur = RawSeg();
            if (!scanHoleEnd(hole)) {   // 已就地报过 1008/1009, 这里只负责退回
                closed = false;
                holeBad = true;
                segs.clear();
                break;
            }
            holes.push_back(hole);
            atOpen = false;
            continue;
        }
        if (c == '\n' || c == '\r') {
            markSeg();
            // 行界归一成 CRLF。fromFile() 已把 CRLF/孤立 CR 抹成 LF (source_manager.cpp)，
            // 所以源码行尾风格不可能影响串值；这里同时兜住不归一的 fromString() 通路。
            if (c == '\r' && peekNext() == '\n') advance();
            advance();
            if (atOpen) { atOpen = false; cur.text.clear(); continue; }
            cur.text += "\r\n";
            continue;
        }
        atOpen = false;
        markSeg();
        if (c == '"') {
            advance();
            cur.text += "\"\"";        // 重新双写, 交给下游那四处折叠
            continue;
        }
        cur.text += advance();
    }

    if (!closed) {
        if (!holeBad) {   // 孔的 1008/1009 已经报过, 不再补一条 1007
            errorAt(DiagnosticID::LexUnterminatedRawString, startLine, startCol,
                    "unterminated raw string (missing closing backtick)");
        }
        if (!holeBad) {
            // 普通的未闭合: 退回开头那一格, 让余文照常被扫 (一条 1007, 不吞源码)。
            offset_ = bodyOff;
            line_ = bodyLine;
            column_ = bodyCol;
        } else {
            // 坏孔: 丢到本行末尾。留着不丢的话, 同一行尾上那枚闭合反引号会被下一次扫描
            // 当成**新串**的开头, 一条 1009 立刻升级成 1009 + 一串"意外字符" + 假的 1007
            // (实测负例 in_n2)。丢掉整行只留"这个孔坏了"这一条词法读数。
            while (offset_ < content_.size() && content_[offset_] != '\n') advance();
        }
        return makeToken(TokenKind::Invalid, "`", startLine, startCol);
    }

    if (holes.empty()) {
        return makeToken(TokenKind::StringLiteral, "\"" + cur.text + "\"", startLine, startCol);
    }
    segs.push_back(cur);   // 尾段
    return emitRawInterp(segs, holes, startLine, startCol);
}

// 前置条件: offset_ 指着孔的 '$'。成功时 offset_ 落在闭合 '}' 之后, hole.begin/exprEnd 是
// **整份文件里**的绝对字节偏移 (下一级开窗要用), line/col 供诊断。
// 失败 (未闭合 1008 / 空孔 1009) 就地报好并回 false —— 调用方只负责退回与交付 Invalid。
bool Lexer::scanHoleEnd(RawHole& hole) {
    hole.begin = base_ + offset_ + 2;              // 跳过 "${"
    uint32_t p = offset_ + 2;
    uint32_t fmtAt = 0;                            // 顶层 ':' 的位置 (0 = 没有格式段)
    int depth = 0;

    while (p < content_.size()) {
        char ch = content_[p];
        if (ch == '\n' || ch == '\r') break;       // 孔不许跨行: VB6 的表达式本来就不跨行
        if (ch == '\'') {                          // 注释到行尾
            while (p < content_.size() && content_[p] != '\n') p++;
            continue;
        }
        if (ch == '"' || ch == '`' || ch == '#') {
            // 串 / 原始串 / 日期面量整段跳过 ⇒ 里面的 } 与 : 都不算数
            const char closer = ch;
            uint32_t q = p + 1;
            while (q < content_.size()) {
                if (content_[q] == '\n' && closer != '`') break;   // 普通串不跨行
                if (content_[q] == closer) {
                    if (q + 1 < content_.size() && content_[q + 1] == closer) { q += 2; continue; }
                    q++;
                    break;
                }
                q++;
            }
            p = q;
            continue;
        }
        if (ch == '(' || ch == '[') { depth++; p++; continue; }
        if (ch == ')' || ch == ']') { if (depth > 0) depth--; p++; continue; }
        if (ch == ':' && depth == 0 && fmtAt == 0) {
            // 顶层 ':' = 格式段的分隔符。此后的文本**不再做任何特殊处理**，取到第一个 '}'
            // 为止 (计划书 §二) —— 因为格式串里满是 `#`/`"` 这些在表达式里另有身份的字符
            // (`${n:#,##0}` 若继续按"跳过 #…# 日期面量"扫，会被第一个 '#' 一路吞到行尾)。
            fmtAt = p;
            uint32_t q = p + 1;
            while (q < content_.size() && content_[q] != '}' && content_[q] != '\n') q++;
            if (q >= content_.size() || content_[q] != '}') break;   // 交给下面的 1008
            hole.exprEnd = base_ + p;
            hole.hasFmt = true;
            for (uint32_t k = p + 1; k < q; k++) {
                // 内嵌的 " 重新双写: 格式串要作为字面量交给下游那四处折叠
                hole.fmt += (content_[k] == '"') ? std::string("\"\"") : std::string(1, content_[k]);
            }
            if (hole.exprEnd == hole.begin) {   // ${:fmt} —— 表达式位是空的
                errorAt(DiagnosticID::LexEmptyInterpHole, hole.line, hole.col,
                        "empty interpolation hole");
                return false;
            }
            offset_ = q;
            advance();                          // 消费 '}'
            return true;
        }
        if (ch == '}' && depth == 0) {
            if (p == hole.begin - base_) {      // 空孔 ${}
                errorAt(DiagnosticID::LexEmptyInterpHole, hole.line, hole.col,
                        "empty interpolation hole");
                return false;
            }
            hole.exprEnd = base_ + p;
            offset_ = p;
            advance();                             // 消费 '}'
            return true;
        }
        p++;
    }

    errorAt(DiagnosticID::LexUnterminatedInterp, hole.line, hole.col,
            "unterminated interpolation hole (missing '}')");
    return false;
}

// 孔里的表达式原文 = 一个窗口化的 Lexer 扫出来的普通 token。窗口换来三件事
// (计划书 §三.7)：诊断行列天然落在原文件上 (不用事后平移 AST)、越不出孔的右界、
// 整串一趟走完 (不回头重扫)。
std::vector<Token> Lexer::lexWindow(uint32_t begin, uint32_t end) {
    std::vector<Token> out;
    Lexer sub(buffer_, diag_, begin, end);
    while (!sub.eof()) {
        Token t = sub.nextToken();
        if (t.kind == TokenKind::EndOfFile) break;
        out.push_back(std::move(t));
    }
    return out;
}

// 展开产物: ( "文本" & CStr( expr ) & "文本" ) / ( ... & Format$( expr, "fmt" ) & ... )
// —— 与手写的 `a & CStr(x) & b` 同形, 不是"看起来像"。第一枚交回去, 其余压进 pending_。
Token Lexer::emitRawInterp(std::vector<RawSeg>& segs, const std::vector<RawHole>& holes,
                           uint32_t startLine, uint32_t startCol) {
    std::vector<Token> seq;
    seq.push_back(makeToken(TokenKind::LeftParen, "(", startLine, startCol));
    bool joined = false;

    auto putLit = [&](const RawSeg& s) {
        if (s.text.empty()) return;   // 空文本段不发 `""` ⇒ `${n}` 展开成 ( CStr( n ) )
        if (joined) seq.push_back(makeToken(TokenKind::Ampersand, "&", s.line, s.col));
        seq.push_back(makeToken(TokenKind::StringLiteral, "\"" + s.text + "\"", s.line, s.col));
        joined = true;
    };

    for (size_t i = 0; i < holes.size(); i++) {
        const RawHole& h = holes[i];
        putLit(segs[i]);
        if (joined) seq.push_back(makeToken(TokenKind::Ampersand, "&", h.line, h.col));
        seq.push_back(makeToken(TokenKind::Identifier, h.hasFmt ? "Format$" : "CStr",
                                h.line, h.col));
        seq.push_back(makeToken(TokenKind::LeftParen, "(", h.line, h.col));
        for (auto& t : lexWindow(h.begin, h.exprEnd)) seq.push_back(std::move(t));
        if (h.hasFmt) {
            seq.push_back(makeToken(TokenKind::Comma, ",", h.line, h.col));
            seq.push_back(makeToken(TokenKind::StringLiteral, "\"" + h.fmt + "\"",
                                    h.line, h.col));
        }
        seq.push_back(makeToken(TokenKind::RightParen, ")", h.line, h.col));
        joined = true;
    }
    putLit(segs[holes.size()]);
    seq.push_back(makeToken(TokenKind::RightParen, ")", startLine, startCol));

    for (size_t i = 1; i < seq.size(); i++) pending_.push_back(std::move(seq[i]));
    return seq[0];
}

void Lexer::errorAt(DiagnosticID id, uint32_t line, uint32_t col, const char* msg) {
    SourceLocation loc{std::string(buffer_->filePath()), line, col};
    diag_.error(id, loc, msg);
}

// === 日期字面量扫描 ===

Token Lexer::scanDateLiteral() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;
    text += advance(); // 消费 #

    while (offset_ < content_.size() && peek() != '#') {
        if (peek() == '\n') {
            SourceLocation loc{std::string(buffer_->filePath()), startLine, startCol};
            diag_.error(DiagnosticID::LexInvalidCharLiteral, loc, "日期字面量未闭合");
            break;
        }
        text += advance();
    }

    if (peek() == '#') {
        text += advance(); // 消费结尾 #
    }

    return makeToken(TokenKind::DateLiteral, text, startLine, startCol);
}

// === 注释扫描 ===

Token Lexer::scanComment() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;
    text += advance(); // 消费 '

    // 注释到行尾
    while (offset_ < content_.size() && peek() != '\n') {
        text += advance();
    }

    // 检查行尾是否有行续接符 _
    // 在注释中的行续接在VB6中无效, 但我们仍记录
    return makeToken(TokenKind::Comment, text, startLine, startCol);
}

// === 行续接扫描 ===

Token Lexer::scanLineContinuation() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;
    text += advance(); // 消费 _

    // 跳过空白
    while (offset_ < content_.size() && isWhitespace(peek()) && peek() != '\n') {
        text += advance();
    }

    // 跳过换行
    if (peek() == '\n') {
        text += advance();
    }

    // 跳过下一行开头的空白
    while (offset_ < content_.size() && isWhitespace(peek())) {
        advance(); // 不记录在text中
    }

    afterLineContinuation_ = true;
    return makeToken(TokenKind::LineContinuation, text, startLine, startCol);
}

// === 运算符扫描 (复杂情况) ===

Token Lexer::scanOperator() {
    uint32_t startLine = line_, startCol = column_;
    char c = peek();

    if (c == '&') {
        advance();
        // C3 扩展 (ai/032): `&=` 复合赋值。能走到这里说明这个 `&` 前面是空白或
        // 运算符 —— 紧贴标识符的 `&` (如 `x&` 的 Long 后缀) 已被
        // scanIdentifierOrKeyword 的类型后缀那一步吃掉 (它的判据是「`&` 后面不是
        // 字母/数字/下划线」, `=` 正好落在里面), 所以这里不会与 VB6 的
        // `x& = 5` 撞车。参见 lexer.cpp 运算符 switch 的同类注释。
        if (offset_ < content_.size() && peek() == '=') {
            advance();
            return makeToken(TokenKind::AmpEq, "&=", startLine, startCol);
        }
        // & 作为 Long类型后缀 : 在标识符后
        // & 作为字符串连接符 : 在表达式中间
        // 上下文决定, 这里统一返回Ampersand, 由语法分析器区分
        return makeToken(TokenKind::Ampersand, "&", startLine, startCol);
    }

    // 不应到达此处
    advance();
    return errorToken(std::string("意外运算符: '") + c + "'", startLine, startCol);
}

Token Lexer::scanNewLine() {
    uint32_t startLine = line_, startCol = column_;
    advance();
    return makeToken(TokenKind::NewLine, "\n", startLine, startCol);
}

} // namespace vb6c3
