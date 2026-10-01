#include "lexer/lexer.hpp"
#include <cstdint>

namespace vb6c3 {

// === 数字扫描 ===

namespace {

// Fix 082: 解析整数字面量 (基数 2/8/10/16) 为 uint64_t 再按补码重解释为 int64_t。
// 直接 std::stoll 会在值 > INT64_MAX (如 &H8000000000000000) 时抛异常, 而被上层
// catch(...) 吞掉后静默变成 0 —— 生成错误代码。此处显式溢出检测并返回 false。
bool parseIntLit(const std::string& s, int base, int64_t& out) {
    if (s.empty()) return false;
    const uint64_t cap = UINT64_MAX;
    uint64_t v = 0;
    for (char ch : s) {
        int d;
        if (ch >= '0' && ch <= '9')      d = ch - '0';
        else if (ch >= 'a' && ch <= 'f') d = ch - 'a' + 10;
        else if (ch >= 'A' && ch <= 'F') d = ch - 'A' + 10;
        else return false;
        if (d >= base) return false;
        if (v > (cap - d) / static_cast<uint64_t>(base)) return false; // 溢出
        v = v * static_cast<uint64_t>(base) + static_cast<uint64_t>(d);
    }
    out = static_cast<int64_t>(v); // 0x8000000000000000 -> INT64_MIN (VBA 语义)
    return true;
}

// 去掉 "&H"/"&O"/"&B" 前缀与末尾 "&"/"^" 后缀, 返回纯数字部分
std::string radixDigits(const std::string& s) {
    size_t begin = 0;
    if (s.size() >= 2 && s[0] == '&' &&
        (s[1] == 'H' || s[1] == 'h' || s[1] == 'O' || s[1] == 'o' || s[1] == 'B' || s[1] == 'b')) {
        begin = 2;
    }
    size_t end = s.size();
    if (end > begin && (s[end - 1] == '^' || s[end - 1] == '&')) end--;
    return s.substr(begin, end - begin);
}

} // namespace

// offset_ 指向 '^' 本身. 从 offset_+1 起向后看 (跳过空格/制表符): 若不是数字/&/.
// /-/+/( 之一, 则这个 '^' 是 LongPtr 后缀; 否则是幂运算符.
//
// 关键: 必须从 offset_+1 开始. 若从 offset_ 起跳空格, 第一个字符就是 '^' 自己,
// 它既不是数字也不是 '&'/'.' -> 无条件返回"后缀", 消歧完全失效, 会把
// `2 ^ 10` 误读成 LongPtrLit("2^") + 10 (Fix 082 验证探针实测到的回归).
//
// 消歧方向的选择 (验证数据, 见 tests/ 语料统计):
//   '^' 作为幂运算符在语料中 4 处 (均为 "数字 ^ 数字"), 作为 LongPtr 后缀 0 处;
//   LongPtr 字面量是 Office 2010+ / VBA7 的冷门特性 (样例工程 Common.bas 仅 5 处).
//   因此偏向幂运算符 —— 判错幂运算符会静默改坏大量常规数学代码.
//   - 含 '-'/'+': 否则 `2 ^ -3` 会误成 LongPtrLit("2^") - 3 = -1 (真值 0.125),
//     仍是合法表达式 -> 静默错误. 而 LongPtr 字面量后接二元 +/- 在样例中从不出现
//     (样例总是加括号: ((Start Xor &H8000000000000000^) + Incr)).
//   - 含 '(': 否则 `2 ^ (n+1)` 同样静默误读.
//   - 不含字母: "&H80000000^ Then" 中 ^ 后是关键字, 必须仍视为后缀. 代价是
//     `2 ^ Foo` 误判为后缀, 但那会得到 LongPtrLit 后紧跟另一个主表达式 -> 解析期
//     硬错误 (响亮), 不是静默错误.
bool Lexer::isLongPtrSuffixHere() const {
    uint32_t p = offset_ + 1;
    while (p < static_cast<uint32_t>(content_.size()) &&
           (content_[p] == ' ' || content_[p] == '\t')) {
        p++;
    }
    if (p >= static_cast<uint32_t>(content_.size())) return true; // 行尾 -> 后缀
    char c = static_cast<char>(content_[p]);
    // 这些字符能开启一个可作为幂运算符右操作数的表达式 -> '^' 是幂运算符
    if (isDigit(c) || c == '&' || c == '.' ||
        c == '-' || c == '+' || c == '(') {
        return false;
    }
    return true;
}

Token Lexer::scanNumber() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;
    bool isFloat = false;
    bool hasExponent = false;

    // 整数部分
    while (offset_ < content_.size() && isDigit(peek())) {
        text += advance();
    }

    // 小数部分
    if (peek() == '.' && !isAlpha(peekNext())) {
        // 需要区分: 123.456 (浮点) vs 123.Method (成员访问)
        char nextNext = peekNextNext();
        if (isDigit(peekNext()) || nextNext == 'e' || nextNext == 'E') {
            isFloat = true;
            text += advance(); // 消费 .
            while (offset_ < content_.size() && isDigit(peek())) {
                text += advance();
            }
        }
    }

    // 指数部分
    if (offset_ < content_.size() && (peek() == 'e' || peek() == 'E')) {
        isFloat = true;
        hasExponent = true;
        text += advance(); // 消费 E
        if (offset_ < content_.size() && (peek() == '+' || peek() == '-')) {
            text += advance();
        }
        while (offset_ < content_.size() && isDigit(peek())) {
            text += advance();
        }
    }

    // 类型后缀
    bool isLong = false;
    bool isLongPtr = false;
    if (offset_ < content_.size()) {
        char c = peek();
        switch (c) {
            case '%': text += advance(); break;
            case '&': text += advance(); isLong = true; break;
            case '^': // Fix 082: VBA7 LongPtr 后缀 (与幂运算符消歧见 isLongPtrSuffixHere)
                if (!isFloat && !hasExponent && isLongPtrSuffixHere()) {
                    text += advance();
                    isLongPtr = true;
                }
                break;
            case '!': text += advance(); isFloat = true; break;
            case '#': text += advance(); isFloat = true; break;
            case '@': text += advance(); {
                Token tok = makeToken(TokenKind::DecimalLiteral, text, startLine, startCol);
                // 解析Decimal值 (暂用double)
                try { tok.doubleValue = std::stod(text); } catch (...) {}
                return tok;
            }
        }
    }

    if (isFloat) {
        Token tok = makeToken(TokenKind::FloatLiteral, text, startLine, startCol);
        try { tok.doubleValue = std::stod(text); } catch (...) {}
        return tok;
    }

    if (isLongPtr) {
        int64_t v = 0;
        if (!parseIntLit(radixDigits(text), 10, v)) {
            return errorToken("十进制数字超出 LongPtr 范围", startLine, startCol);
        }
        Token tok = makeToken(TokenKind::LongPtrLiteral, text, startLine, startCol);
        tok.longValue = v;
        return tok;
    }

    if (isLong) {
        Token tok = makeToken(TokenKind::LongLiteral, text, startLine, startCol);
        try { tok.longValue = std::stoll(text); } catch (...) {}
        return tok;
    }

    // 无后缀十进制整数字面量: 按**数值大小**定类型, 不再一律当 Integer。
    //
    // 原先这里恒发 IntegerLiteral 并 `static_cast<int32_t>` 截断, 装不进 32 位的
    // 数值会被**回绕**而不是报错, 而负号是**另一个 token** (一元 Negate), 于是
    //   l = -2147483648         ->  l = (--2147483648)   C2105, 整工程编不过
    //   l = -9223372036854775807->  l = (--1)            值被静默改成 -1
    // 两处都是**静默错编**: 前者让合法程序编不过, 后者算出完全不同的数。
    //
    // 装得进 32 位的**保持原样**仍走 IntegerLiteral —— 那条路本来就是对的, 不动它,
    // 免得给满仓库的字面量平白加一圈文本变动。超出的走 LongLiteral (int64 存),
    // 由 cgen 视宽度补 LL 后缀。
    //
    // 再大 (超过 64 位) 直接报词法错, 不提升为 Double: Double 只有 53 位尾数,
    // 提升等于**主动**把值改错, 报出来更诚实。
    {
        int64_t v = 0;
        if (!parseIntLit(radixDigits(text), 10, v)) {
            return errorToken("十进制数字超出 64 位整数表示范围", startLine, startCol);
        }
        if (v >= INT32_MIN && v <= INT32_MAX) {
            Token tok = makeToken(TokenKind::IntegerLiteral, text, startLine, startCol);
            tok.intValue = static_cast<int32_t>(v);
            return tok;
        }
        Token tok = makeToken(TokenKind::LongLiteral, text, startLine, startCol);
        tok.longValue = v;
        return tok;
    }
}

Token Lexer::scanHexNumber() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;
    text += advance(); // 消费 &
    text += advance(); // 消费 H/h

    if (offset_ >= content_.size() || !isHexDigit(peek())) {
        return errorToken("十六进制数字格式错误", startLine, startCol);
    }

    while (offset_ < content_.size() && isHexDigit(peek())) {
        text += advance();
    }

    // Fix 082: ^ 后缀表示 LongPtr (VBA7, 指针宽度)
    if (peek() == '^' && isLongPtrSuffixHere()) {
        text += advance();
        int64_t v = 0;
        if (!parseIntLit(radixDigits(text), 16, v)) {
            return errorToken("十六进制数字超出 LongPtr 范围", startLine, startCol);
        }
        Token tok = makeToken(TokenKind::LongPtrLiteral, text, startLine, startCol);
        tok.longValue = v;
        return tok;
    }

    // & 后缀表示 Long
    if (peek() == '&') {
        text += advance();
        Token tok = makeToken(TokenKind::LongLiteral, text, startLine, startCol);
        try { tok.longValue = std::stoll(text.substr(2), nullptr, 16); } catch (...) {}
        return tok;
    }

    Token tok = makeToken(TokenKind::IntegerLiteral, text, startLine, startCol);
    try { tok.intValue = static_cast<int32_t>(std::stoll(text.substr(2), nullptr, 16)); } catch (...) {}
    return tok;
}

Token Lexer::scanOctNumber() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;
    text += advance(); // 消费 &
    text += advance(); // 消费 O/o

    while (offset_ < content_.size() && isOctDigit(peek())) {
        text += advance();
    }

    // Fix 082: ^ 后缀表示 LongPtr (VBA7)
    if (peek() == '^' && isLongPtrSuffixHere()) {
        text += advance();
        int64_t v = 0;
        if (!parseIntLit(radixDigits(text), 8, v)) {
            return errorToken("八进制数字超出 LongPtr 范围", startLine, startCol);
        }
        Token tok = makeToken(TokenKind::LongPtrLiteral, text, startLine, startCol);
        tok.longValue = v;
        return tok;
    }

    if (peek() == '&') text += advance();
    Token tok = makeToken(TokenKind::IntegerLiteral, text, startLine, startCol);
    try { tok.intValue = static_cast<int32_t>(std::stoll(text.substr(2), nullptr, 8)); } catch (...) {}
    return tok;
}

Token Lexer::scanBinNumber() {
    uint32_t startLine = line_, startCol = column_;
    std::string text;
    text += advance(); // 消费 &
    text += advance(); // 消费 B/b

    while (offset_ < content_.size() && isBinDigit(peek())) {
        text += advance();
    }

    // Fix 082: ^ 后缀表示 LongPtr (VBA7)
    if (peek() == '^' && isLongPtrSuffixHere()) {
        text += advance();
        int64_t v = 0;
        if (!parseIntLit(radixDigits(text), 2, v)) {
            return errorToken("二进制数字超出 LongPtr 范围", startLine, startCol);
        }
        Token tok = makeToken(TokenKind::LongPtrLiteral, text, startLine, startCol);
        tok.longValue = v;
        return tok;
    }

    if (peek() == '&') text += advance();
    Token tok = makeToken(TokenKind::IntegerLiteral, text, startLine, startCol);
    try { tok.intValue = static_cast<int32_t>(std::stoll(text.substr(2), nullptr, 2)); } catch (...) {}
    return tok;
}

} // namespace vb6c3
