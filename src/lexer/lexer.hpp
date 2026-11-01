#pragma once
// VB6词法分析器 - 手写LL(2)扫描器
// 参考: FreeBASIC lex.bas (3-token前瞻缓冲)

#include "lexer/token.hpp"
#include "common/diagnostics.hpp"
#include "common/source_manager.hpp"
#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>
#include <deque>

namespace vb6c3 {

// ai/028 V1/V2 反引号串的内部形状 (只在词法层活一会儿, 不进 AST —— 计划书 R4)。
struct RawSeg {
    std::string text;   // 已归一的文本: 行界 = CRLF, 双写的反引号已折, 内嵌 " 已重新双写
    uint32_t line = 1;
    uint32_t col = 1;
};

struct RawHole {
    uint32_t begin = 0;      // 孔内表达式的源内绝对偏移 [begin, exprEnd)
    uint32_t exprEnd = 0;
    std::string fmt;         // ':' 格式段原文 (无则空)
    bool hasFmt = false;
    uint32_t line = 1;       // '${' 那一格, 给诊断用
    uint32_t col = 1;
};

class Lexer {
public:
    // 使用整份 buffer (存量调用点一字不改)。
    // [begin,end) 形式的**窗口**只为一件事: 反引号插值串 (ai/028 V2) 要把孔里的表达式原文
    // 单独扫一遍, 而窗口右界就是闭合反引号那一侧 —— 于是"越不出串外"由类型保证, 而
    // line_/column_ 用 getLocation 播种 ⇒ 子树的位置天然落在原文件上 (不需要事后平移 AST)。
    // end=0 表示扫到文件尾。
    explicit Lexer(std::shared_ptr<SourceBuffer> buffer, Diagnostics& diag,
                   uint32_t begin = 0, uint32_t end = 0);

    // 读取下一个token, 推进位置
    Token nextToken();

    // 前瞻1个token (不推进位置)
    const Token& peekToken();

    // 前瞻2个token (不推进位置)
    const Token& peekToken2();

    // 当前位置
    uint32_t offset() const { return offset_; }
    uint32_t line() const { return line_; }
    uint32_t column() const { return column_; }

    // 是否到达文件末尾
    bool eof() const { return offset_ >= content_.size(); }

private:
    // 初始化关键字映射表
    void initKeywords();

    // 底层扫描函数
    Token scanToken();
    Token takeScanned();   // pending_ 优先于 scanToken() (见 lexer.cpp 的"前瞻接口")
    Token scanIdentifierOrKeyword();
    Token scanNumber();
    Token scanHexNumber();      // &H...
    Token scanOctNumber();      // &O...
    Token scanBinNumber();      // &B...
    Token scanString();
    Token scanRawString();          // C3 扩展: `...` 原始多行串 (ai/028 V1)

    // === ai/028 V2: 反引号串里的 ${expr} / ${expr:fmt} 插值 ===
    // 展开成**普通 token 序列** (文本段 = StringLiteral、连接 = &、孔 = CStr()/Format$()),
    // 多出来的那些 token 压进 pending_ 由 nextToken() 依次吐出 ⇒ parser / AST / cgen /
    // 语义层里不存在"第二种字符串" (R2/R4), 且未声明变量、按类型挑 vb6_CStrLong、
    // COM 默认属性解析三件事全部自动继承 (降级的目标形状与手写 `a & CStr(x) & b` 同一)。
    bool scanHoleEnd(RawHole& hole);           // current offset_ 指向 '${' 的 '$'
    std::vector<Token> lexWindow(uint32_t begin, uint32_t end);
    Token emitRawInterp(std::vector<RawSeg>& segs, const std::vector<RawHole>& holes,
                        uint32_t startLine, uint32_t startCol);
    void errorAt(DiagnosticID id, uint32_t line, uint32_t col, const char* msg);
    Token scanDateLiteral();    // #...#
    Token scanOperator();
    Token scanComment();
    Token scanNewLine();
    Token scanLineContinuation();

    // 字符操作
    char peek() const;
    char peekNext() const;       // 前瞻1个字符
    char peekNextNext() const;   // 前瞻2个字符
    char advance();              // 前进1个字符
    bool match(char expected);   // 如果当前字符匹配则前进

    // VBA7 LongPtr 字面量后缀 (^) 与 ^ 幂运算符消歧 (Fix 082)
    // 当前 offset_ 指向 '^'。从 offset_+1 起向后看 (跳过空格/制表符): 若不是
    // 数字/&/./-/+/( 之一, 则 '^' 是 LongPtr 后缀; 否则是幂运算符。
    // 详见 lexer_number.cpp 中同名函数的注释 (消歧方向的数据依据与误判代价)。
    bool isLongPtrSuffixHere() const;

    // 辅助判断
    bool isAlpha(char c) const;
    bool isDigit(char c) const;
    bool isHexDigit(char c) const;
    bool isOctDigit(char c) const;
    bool isBinDigit(char c) const;
    bool isAlphaNumeric(char c) const;
    bool isWhitespace(char c) const;
    bool isLineContinuation() const;

    // 创建token的辅助
    Token makeToken(TokenKind kind, const std::string& text, uint32_t startLine, uint32_t startCol);
    // 账 #305 第313刀: 三种进制扫描器 (&H / &O / &B) 的**唯一**出口 —— 档位与值都问
    // lexer_number.cpp 里那一位主人, 三支不许各自再抄一遍 (抄一次漏一次 = 静默错值)。
    Token makeRadixToken(const std::string& text, uint32_t startLine, uint32_t startCol,
                         uint64_t magnitude, bool intSuffix, bool allowNarrowFold);
    Token errorToken(const std::string& msg, uint32_t startLine, uint32_t startCol);

    // 状态
    std::shared_ptr<SourceBuffer> buffer_;
    Diagnostics& diag_;
    std::string_view content_;

    uint32_t offset_ = 0;
    uint32_t line_ = 1;
    uint32_t column_ = 1;
    // content_ 的起点在整份文件里的绝对偏移 (整份扫描 = 0; 插值孔的窗口 = begin)。
    // 需要绝对偏移的唯一理由: 给孔开下一级窗口时要按**原文件**的字节位置 substr。
    uint32_t base_ = 0;

    // 前瞻缓冲
    static constexpr int LOOKAHEAD = 2;
    std::deque<Token> lookahead_;
    // ai/028 V2: 插值展开多出来的 token (lookahead_ 只有 2 格, 装不下一整条降级链)
    std::deque<Token> pending_;

    // 关键字映射 (小写)
    std::unordered_map<std::string, TokenKind> keywords_;

    // 行续接后的行号映射
    bool afterLineContinuation_ = false;
};

} // namespace vb6c3
