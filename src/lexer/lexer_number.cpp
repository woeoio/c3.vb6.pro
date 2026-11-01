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
    // 账 #195: `%` 也是后缀 (VB 的 Integer 档)。此前它不在剥离表里，于是 `3%` 落进
    // 「无后缀十进制按数值大小定档」那一段，parseIntLit 看见残留的 `%` 就报
    // 「超出 64 位整数表示范围」—— 合法程序整个编不过。
    if (end > begin && (s[end - 1] == '^' || s[end - 1] == '&' || s[end - 1] == '%')) end--;
    return s.substr(begin, end - begin);
}

// 账 #305 第313刀: 无后缀进制字面量「写下的位模式 → VB 认的类型与值」那一问, 全编译器只此一处。
//
// VB6 的档位 = 装得下这个位模式的**最小有符号类型**, 于是值先按那一档取补码再交给上层:
//   <= 0xFFFF      → 16 位补码 (&HBF32 = -16590、&HFFFF = -1、&O177777 = -1)
//   <= 0xFFFFFFFF  → 32 位补码 (&HFFFFFFFF = -1 —— 旧行为, 逐字保留)
//   再大            → 交给 LongLiteral 按 int64 存 (ai/032 §6.6 那一档)
// 分界量的是**数值位宽**, 不是位数: 同一发里 &H10000 = +65536 与 &O200000 = +65536 不动,
// 而 &O100000 = -32768 要动 —— 口径由 VBScript 引擎实测 (.local/k313 两轮探针)。
//
// 缺 16 位那一档不是洁癖而是静默改坏值: `CLSID.Data2 = &HBF32` (ComCtlsDemo 与
// VBFlexGridDemo 共用的 Common.bas) 从前发成 vb6_ChkInt(48946), 运行期抛错误 6 打死进程;
// 位运算那一路更广 —— 255 Or &HFF00 在 VB 里是 -1 (掩码自己带符号扩展), 按 +65280 算是另一个数。
//
// `%` 后缀强制 Integer 档 (账 #195), 它**不许**参加上面那档折叠: 显式后缀说的是「按这个类型收」,
// 而 VB 对 `&HFFFF&` 实测交回 65535 (带 & 后缀走各自那一支, 见 tests/intsuffix 的证人)。
// allowNarrowFold 就是「这一发真的无后缀」那一问: &H 那一支的 & 后缀在更早处就返回了, 而 &O / &B
// 两支从前把 & 吃进文本却不改档位 (实测语料 0 处落在 16 位档, .local/k313/octbin.txt) ⇒ 逐字保留
// 它们今天的行为, 只把折叠让给无后缀那一形。
struct RadixTier { TokenKind kind; int64_t value; };

RadixTier radixTierOf(uint64_t magnitude, bool intSuffix, bool allowNarrowFold) {
    if (intSuffix) return {TokenKind::IntegerLiteral, static_cast<int32_t>(magnitude)};
    if (allowNarrowFold && magnitude <= 0xFFFFu) {
        return {TokenKind::IntegerLiteral,
                static_cast<int64_t>(magnitude) - ((magnitude & 0x8000u) ? 0x10000 : 0)};
    }
    if (magnitude <= 0xFFFFFFFFu) {
        return {TokenKind::IntegerLiteral, static_cast<int32_t>(static_cast<uint32_t>(magnitude))};
    }
    return {TokenKind::LongLiteral, static_cast<int64_t>(magnitude)};
}

} // namespace

// 进制字面量 (&H / &O / &B) 三条扫描支路的**唯一**发token出口: 档位与值都问上面那位主人。
// 三支从前各抄一份阶梯 (三处 `(uint64_t)v <= 0xFFFFFFFFu` + 三处 static_cast), 于是第312刀
// 量到的那一档缺口一次性存在于三处 —— 加一档要改三遍、漏一遍就是静默错值 (本刀的病灶正是漏了
// 16 位那一档)。以后要动档位只许动 radixTierOf, 别在支路里再拼比较。
Token Lexer::makeRadixToken(const std::string& text, uint32_t startLine, uint32_t startCol,
                            uint64_t magnitude, bool intSuffix, bool allowNarrowFold) {
    const RadixTier tier = radixTierOf(magnitude, intSuffix, allowNarrowFold);
    Token tok = makeToken(tier.kind, text, startLine, startCol);
    if (tier.kind == TokenKind::IntegerLiteral) {
        tok.intValue = static_cast<int32_t>(tier.value);
    } else {
        tok.longValue = tier.value;
    }
    return tok;
}

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
    bool isInteger = false;   // 账 #195: `%` = Integer 档
    if (offset_ < content_.size()) {
        char c = peek();
        switch (c) {
            case '%': text += advance(); isInteger = true; break;  // 账 #195: 以前只吃不置标志
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

    if (isInteger) {
        // 账 #195: 显式 `%` 后缀 = VB 的 Integer 档 —— 恒发 IntegerLiteral, 不参与
        // 「按数值大小升成 Long」那一路 (那一档是**无后缀**字面量的规则)。值本身装不进
        // 32 位仍是词法错 (与无后缀那一路同口径), 不静默截断。
        int64_t iv = 0;
        if (!parseIntLit(radixDigits(text), 10, iv) || iv < INT16_MIN || iv > INT16_MAX) {
            // 显式 `%` 就是 VB 的 Integer 档 (-32768..32767)。超出**必须报**, 不能像
            // int32 那样收下再让 int16 形参去截 —— 那是把值改错 (静默回绕)。
            return errorToken("Integer 字面量超出 Integer 范围 (-32768..32767)",
                              startLine, startCol);
        }
        Token tok = makeToken(TokenKind::IntegerLiteral, text, startLine, startCol);
        tok.intValue = static_cast<int32_t>(iv);
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

    bool intSuffix = false;                 // 账 #195: &HFF% = Integer 档
    if (peek() == '%') { text += advance(); intSuffix = true; }
    {
        int64_t v = 0;
        if (!parseIntLit(radixDigits(text), 16, v)) {
            return errorToken("十六进制数字超出 64 位整数表示范围", startLine, startCol);
        }
        // ai/032 §6.6 定下的阶梯 (>32 位归 LongLiteral) 与账 #305 第313刀补的 16 位补码那一档
        // 全住在 radixTierOf 那一处 —— 这里只许问它, 不许再抄一遍比较 (抄一次漏一次)。
        return makeRadixToken(text, startLine, startCol, static_cast<uint64_t>(v),
                              intSuffix, /*allowNarrowFold=*/true);
    }
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

    bool longSuffix = false;                // & 后缀: 照旧只吃字符、不改档 (逐字保留, 语料 0 处)
    if (peek() == '&') { text += advance(); longSuffix = true; }
    bool intSuffix = false;                 // 账 #195: 进制字面量也认 Integer 后缀
    if (peek() == '%') { text += advance(); intSuffix = true; }
    {
        int64_t v = 0;
        if (!parseIntLit(radixDigits(text), 8, v)) {
            return errorToken("八进制数字超出 64 位整数表示范围", startLine, startCol);
        }
        /* 阶梯与值都问那一位主人 (账 #305 第313刀), 这里不许再抄一遍。 */
        return makeRadixToken(text, startLine, startCol, static_cast<uint64_t>(v),
                              intSuffix, /*allowNarrowFold=*/!longSuffix);
    }
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

    bool longSuffix = false;                // & 后缀: 照旧只吃字符、不改档 (逐字保留, 语料 0 处)
    if (peek() == '&') { text += advance(); longSuffix = true; }
    bool intSuffix = false;                 // 账 #195: 进制字面量也认 Integer 后缀
    if (peek() == '%') { text += advance(); intSuffix = true; }
    {
        int64_t v = 0;
        if (!parseIntLit(radixDigits(text), 2, v)) {
            return errorToken("二进制数字超出 64 位整数表示范围", startLine, startCol);
        }
        /* 阶梯与值都问那一位主人 (账 #305 第313刀), 这里不许再抄一遍。 */
        return makeRadixToken(text, startLine, startCol, static_cast<uint64_t>(v),
                              intSuffix, /*allowNarrowFold=*/!longSuffix);
    }
}

} // namespace vb6c3
