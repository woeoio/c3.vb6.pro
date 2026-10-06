// vb6c3 - Pratt 表达式解析器
// 参考: RustASP 45行核心实现, (l_bp, r_bp) 绑定力对设计
// VB6 14级优先级, 右结合 ^ 运算符

#include "parser/parser.hpp"
#include <algorithm>
#include <cctype>

namespace vb6c3 {

// ============================================================
// Pratt 核心: parseExpression(minBp)
//
// 原理:
//   1. 解析一个 null denotation (前缀/原子表达式)
//   2. 循环: 如果下一个是中缀运算符, 且其 l_bp >= minBp
//      - 消费运算符
//      - 递归解析右操作数 (r_bp 作为新的 minBp)
//      - 构造 BinaryExpr 节点
//   3. 返回表达式
//
// 关键: 右结合运算符 (如 ^) 的 l_bp > r_bp
//   a ^ b ^ c 解析为 a ^ (b ^ c)
//   因为第二次遇到 ^ 时, l_bp(21) >= minBp(20), 进入循环
//   递归时 minBp = r_bp(20), 所以 c 被归入内层
// ============================================================

ExprPtr Parser::parseExpression() {
    return parseExpression(0);  // 最低优先级, 接受所有运算符
}

ExprPtr Parser::parseExpression(int minBp) {
    // Step 1: 解析 null denotation (前缀/原子)
    auto left = parseNullDenotation();
    if (!left) {
        return left;
    }

    // Step 2: 循环处理中缀/后缀运算符
    while (true) {
        auto bp = getBindingPower(cur_.kind);
        if (bp.l_bp < minBp) {
            break;  // 运算符优先级不够, 退出循环
        }

        // 非运算符 token (如 NewLine, EndOfFile, 标识符等):
        // 绑定力为 {0,0}, 当 minBp=0 时 0 < 0 为 false 会误入循环.
        // 若既非中缀也非后缀起始, 应直接退出, 避免 parseLeftDenotation
        // 将 left move 走后返回 nullptr 导致有效表达式丢失.
        if (bp.l_bp == 0 && bp.r_bp == 0 &&
            cur_.kind != TokenKind::Dot &&
            cur_.kind != TokenKind::LeftParen &&
            cur_.kind != TokenKind::Exclamation) {
            break;
        }

        left = parseLeftDenotation(std::move(left), minBp);
        if (!left) {
            break;
        }
    }

    return left;
}

// ============================================================
// Null Denotation: 前缀/原子表达式
// ============================================================

ExprPtr Parser::parseNullDenotation() {
    switch (cur_.kind) {
        // --- 字面量 ---
        case TokenKind::IntegerLiteral:
        case TokenKind::LongLiteral:
        case TokenKind::LongPtrLiteral:
        case TokenKind::FloatLiteral:
        case TokenKind::DecimalLiteral:
        case TokenKind::StringLiteral:
        case TokenKind::DateLiteral:
        case TokenKind::TrueKeyword:
        case TokenKind::FalseKeyword:
        case TokenKind::NothingKeyword:
        case TokenKind::EmptyKeyword:
        case TokenKind::NullKeyword:
            return parseLiteral();

        // --- 标识符 / 函数调用 ---
        case TokenKind::Identifier:
            return parseIdentifierOrCall();

        // --- 括号表达式 ---
        case TokenKind::LeftParen:
            return parseParenthesizedExpr();

        // --- 前缀运算符 ---
        case TokenKind::Minus: {
            auto loc = currentLoc();
            advance(); // consume '-'
            // 一元负: r_bp 用于控制绑定范围
            // -a + b → (-a) + b  (一元负比加号绑定更紧)
            auto operand = parseExpression(19);  // 与 * 同级 r_bp
            if (!operand) {
                diag_.error(DiagnosticID::ParseExpectedExpression, loc,
                    "expected expression after unary '-'");
                return nullptr;
            }
            return std::make_unique<UnaryExpr>(loc, UnaryOp::Negate, std::move(operand));
        }

        case TokenKind::Not: {
            auto loc = currentLoc();
            advance(); // consume 'Not'
            // Not 优先级低于 And (r_bp=5)
            auto operand = parseExpression(5);
            if (!operand) {
                diag_.error(DiagnosticID::ParseExpectedExpression, loc,
                    "expected expression after 'Not'");
                return nullptr;
            }
            return std::make_unique<UnaryExpr>(loc, UnaryOp::Not, std::move(operand));
        }

        // --- New 表达式 ---
        case TokenKind::New:
            return parseNewExpr();

        // --- TypeOf 表达式 ---
        case TokenKind::TypeOf:
            return parseTypeOfExpr();

        // --- AddressOf 表达式 ---
        case TokenKind::AddressOf:
            return parseAddressOfExpr();

        // --- Me 表达式 ---
        case TokenKind::MeKeyword:
            return parseMeExpr();

        // --- With 块中的 .Member ---
        case TokenKind::Dot:
            if (withDepth_ > 0) {
                return parseWithMemberExpr();
            }
            // 不在 With 块内的 . 是错误
            diag_.error(DiagnosticID::ParseUnexpectedToken, currentLoc(),
                "'.' outside of With block");
            advance();
            return nullptr;

        // --- 字典访问 obj!key ---
        case TokenKind::Exclamation: {
            // 这样写在表达式开头不太合理, 但容错
            auto loc = currentLoc();
            advance();
            auto key = expectName("expected identifier after '!'");
            return std::make_unique<DictionaryAccessExpr>(loc, nullptr, key.text);
        }

                // --- Input$() 函数 (P15.4) ---
        case TokenKind::Input: {
            // Input 后跟 '(' -> Input$ function; otherwise not an expression
            auto loc = currentLoc();
            advance(); // consume 'Input'
            // skip optional $ suffix
            if (cur_.kind == TokenKind::Dollar) advance();
            // must be followed by '('
            if (cur_.kind == TokenKind::LeftParen) {
                advance(); // consume '('
                auto call = std::make_unique<IndexOrCallExpr>(loc,
                    std::make_unique<IdentifierExpr>(loc, "Input$"));
                // parse argument list
                if (cur_.kind != TokenKind::RightParen) {
                    do {
                        auto arg = parseExpression();
                        call->positional.push_back(std::move(arg));
                    } while (match(TokenKind::Comma));
                }
                expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken, "expected ')'");
                return call;
            }
            diag_.error(DiagnosticID::ParseExpectedExpression, loc,
                "expected '(' after Input$ function");
            return nullptr;
        }

        // --- Command$ (P14.2.2) ---
        case TokenKind::Command: {
            auto loc = currentLoc();
            advance(); // consume 'Command'
            // skip optional $ suffix
            if (cur_.kind == TokenKind::Dollar) advance();
            // Command$ / Command is a zero-arg function - return identifier, cgen handles via builtin map
            auto name = std::string("Command");
            if (cur_.kind == TokenKind::LeftParen) {
                advance(); // consume '('
                // parse empty arg list
                expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken, "expected ')'");
            }
            return std::make_unique<IdentifierExpr>(loc, name);
        }
        default:
            // 软关键字在表达式位置 → 解析为标识符
            if (isSoftKeyword(cur_.kind)) {
                return parseIdentifierOrCall();
            }
            diag_.error(DiagnosticID::ParseExpectedExpression, currentLoc(),
                "expected expression, got " + std::string(Token::kindToString(cur_.kind)));
            return nullptr;
    }
}

// ============================================================
// Left Denotation: 中缀/后缀表达式
// ============================================================

ExprPtr Parser::parseLeftDenotation(ExprPtr left, int& minBp) {
    auto kind = cur_.kind;
    auto bp = getBindingPower(kind);

    if (bp.l_bp == 0 && bp.r_bp == 0) {
        // 不是中缀运算符, 检查后缀
        // 后缀: .member, (args), !dict
        if (kind == TokenKind::Dot || kind == TokenKind::LeftParen ||
            kind == TokenKind::Exclamation) {
            return parsePostfix(std::move(left));
        }
        // 不是中缀也不是后缀, 退出循环
        return nullptr;
    }

    // 中缀二元运算符
    auto loc = left->loc;
    auto op = tokenToBinaryOp(kind);
    advance(); // consume operator

    // 特殊处理: Is 运算符在 Case Is 中有不同语义
    // (由 parseCaseValue 单独处理, 此处按普通二元运算符)

    auto right = parseExpression(bp.r_bp);
    if (!right) {
        diag_.error(DiagnosticID::ParseExpectedExpression, currentLoc(),
            "expected expression after binary operator");
        // 尽量恢复: 返回左操作数
        return left;
    }

    return std::make_unique<BinaryExpr>(loc, op, std::move(left), std::move(right));
}

// ============================================================
// 原子表达式
// ============================================================

ExprPtr Parser::parseLiteral() {
    auto loc = currentLoc();
    auto tok = advance();

    switch (tok.kind) {
        case TokenKind::IntegerLiteral: {
            auto expr = std::make_unique<LiteralExpr>(loc, LiteralKind::Integer, tok.text);
            expr->intValue = static_cast<int32_t>(tok.intValue);
            return expr;
        }
        case TokenKind::LongLiteral: {
            auto expr = std::make_unique<LiteralExpr>(loc, LiteralKind::Long, tok.text);
            expr->longValue = tok.longValue;
            return expr;
        }
        case TokenKind::LongPtrLiteral: {
            auto expr = std::make_unique<LiteralExpr>(loc, LiteralKind::LongPtr, tok.text);
            expr->longValue = tok.longValue;
            return expr;
        }
        case TokenKind::FloatLiteral: {
            auto expr = std::make_unique<LiteralExpr>(loc, LiteralKind::Double, tok.text);
            expr->doubleValue = tok.doubleValue;
            return expr;
        }
        case TokenKind::DecimalLiteral: {
            auto expr = std::make_unique<LiteralExpr>(loc, LiteralKind::Decimal, tok.text);
            expr->doubleValue = tok.doubleValue;
            return expr;
        }
        case TokenKind::StringLiteral:
            return std::make_unique<LiteralExpr>(loc, LiteralKind::String, tok.text);
        case TokenKind::DateLiteral: {
            // 账 #172: 值在这里折成 OLE 日期序列。改前只带原文 ⇒ 发码读没写过的联合体槽。
            // 折不出形状就留着 0（不发新诊断）：语料里 VB6 认得而我认得的写法一旦出现，
            // 宁可照旧算错，也不要把一个现在编得过的工程编红。
            auto expr = std::make_unique<LiteralExpr>(loc, LiteralKind::Date, tok.text);
            double oleDate = 0.0;
            if (foldDateLiteralToOADate(tok.text, oleDate)) expr->doubleValue = oleDate;
            return expr;
        }
        case TokenKind::TrueKeyword: {
            auto expr = std::make_unique<LiteralExpr>(loc, LiteralKind::Boolean, tok.text);
            expr->boolValue = true;
            return expr;
        }
        case TokenKind::FalseKeyword: {
            auto expr = std::make_unique<LiteralExpr>(loc, LiteralKind::Boolean, tok.text);
            expr->boolValue = false;
            return expr;
        }
        case TokenKind::NothingKeyword:
            return std::make_unique<LiteralExpr>(loc, LiteralKind::Nothing, tok.text);
        case TokenKind::EmptyKeyword:
            return std::make_unique<LiteralExpr>(loc, LiteralKind::Empty, tok.text);
        case TokenKind::NullKeyword:
            return std::make_unique<LiteralExpr>(loc, LiteralKind::Null, tok.text);
        default:
            diag_.error(DiagnosticID::ParseUnexpectedToken, loc,
                "expected literal");
            return nullptr;
    }
}

ExprPtr Parser::parseIdentifierOrCall() {
    auto loc = currentLoc();
    auto nameTok = advance();
    std::string name = nameTok.text;
    // Task #40: 使用点类型后缀剥离 — 与声明侧 (parseVariableDecl/parseConstDecl
    // 的 Fix 028 stripTypeSuffix) 对齐。`dl& = ...` 的 token 文本含 &, cIdent
    // 会把 & 改成 _ → dl_ (C2065, 而声明名是 dl)。$ 后缀同样剥: 声明 `Dim s$`
    // 名为 s, 使用 s$ 不剥则成 s_ 双重不一致; builtinFuncs 的 $ 键由
    // ident_builtin 的 lookupName strip 兜底 (input 补了无 $ 双键)。
    {
        auto suf40 = stripTypeSuffix(name);
        if (!suf40.typeName.empty()) name = suf40.name;
    }
    // 泛型使用点 (tB): Foo(Of Long)(x) —— 先把 (Of …) 尾巴扁进名字,
    // 剩下的 (...) 才由 parsePostfix 当作真正的实参表.
    tryFlattenGenericName(name);
    auto expr = std::make_unique<IdentifierExpr>(loc, name);

    // 后缀处理: 可能是函数调用 arr(i) 或 func(x,y)
    return parsePostfix(std::move(expr));
}

ExprPtr Parser::parseParenthesizedExpr() {
    auto loc = currentLoc();
    advance(); // consume '('

    // VB6: 空括号 () 可能是数组声明或无参调用, 但在表达式上下文
    // 这里处理的是 (expr) 分组
    auto expr = parseExpression();
    expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
           "expected ')'");
    return expr;
}

ExprPtr Parser::parseNewExpr() {
    auto loc = currentLoc();
    advance(); // consume 'New'
    auto classTok = expectName("expected class name after 'New'");
    std::string className = classTok.text;
    // Qualified class name: New Scripting.Dictionary, New MSComctlLib.ImageList, etc.
    while (cur_.kind == TokenKind::Dot) {
        advance(); // consume '.'
        auto nextTok = expectName("expected class name after '.'");
        className += "." + nextTok.text;
    }
    // 泛型使用点 (tB): New Foo(Of Long) — 扁名化, 泛型器据此特化类模块
    tryFlattenGenericName(className);
    // C3 扩展 (084c): `New Cls(args)` 带参构造 — VB6 本体不允许 (构造只能走
    // 无参 Class_Initialize), 这里在语法上放行, 语义层按目标类 Class_Initialize
    // 的形参个数校验 (ctorParams_ 预计算表), COM/外部类带实参一律报错。
    if (cur_.kind == TokenKind::LeftParen) {
        auto ne = std::make_unique<NewExpr>(loc, className);
        advance(); // consume '('
        if (cur_.kind != TokenKind::RightParen) {
            do {
                auto arg = parseExpression();
                if (arg) ne->args.push_back(std::move(arg));
            } while (match(TokenKind::Comma));
        }
        expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken, "expected ')'");
        return ne;
    }
    return std::make_unique<NewExpr>(loc, className);
}

ExprPtr Parser::parseTypeOfExpr() {
    auto loc = currentLoc();
    advance(); // consume 'TypeOf'
    auto obj = parseExpression(23);  // TypeOf 比 Is 优先级低
    // 词法器将Is输出为IsKeyword，两者都接受
    if (cur_.kind != TokenKind::Is && cur_.kind != TokenKind::IsKeyword) {
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                    "expected 'Is' after 'TypeOf'");
    } else {
        advance();
    }
    auto typeTok = expectName("expected type name");
    // Fix 088: TypeOf obj Is ADODB.Recordset / VBMANLIB.Dictionary — 类型名可带
    // 库前缀 (点号限定). 此前只吃一个名字 token, 剩余的 ".Recordset" 会被
    // parsePostfix 挂到 TypeOfExpr 结果上, 生成 vb6_TypeOf(...,L"ADODB").Recordset
    // → C2224/C2039 (点号左侧非结构体 / 类型串不完整).
    std::string typeName = typeTok.text;
    while (cur_.kind == TokenKind::Dot) {
        advance(); // consume '.'
        auto nextTok = expectName("expected type name after '.'");
        typeName += "." + nextTok.text;
    }
    return std::make_unique<TypeOfExpr>(loc, std::move(obj), typeName);
}

ExprPtr Parser::parseAddressOfExpr() {
    auto loc = currentLoc();
    advance(); // consume 'AddressOf'
    auto funcTok = expectName("expected function name after 'AddressOf'");
    return std::make_unique<AddressOfExpr>(loc, funcTok.text);
}

ExprPtr Parser::parseMeExpr() {
    auto loc = currentLoc();
    advance(); // consume 'Me'

    auto expr = std::make_unique<MeExpr>(loc);
    return parsePostfix(std::move(expr));
}

ExprPtr Parser::parseWithMemberExpr() {
    auto loc = currentLoc();
    advance(); // consume '.'
    // 与 parsePostfix Dot 分支一致：允许硬关键字作为成员名 (如 .Type, .Loop)
    std::string memberName;
    if (canBeName(cur_.kind)) {
        memberName = advance().text;
    } else if (!cur_.text.empty() && cur_.kind != TokenKind::EndOfFile &&
               cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::Colon &&
               cur_.kind != TokenKind::LeftParen && cur_.kind != TokenKind::RightParen &&
               cur_.kind != TokenKind::Comma) {
        memberName = advance().text;
    } else {
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            std::string("expected member name after '.' (got ") +
            Token::kindToString(cur_.kind) + ")");
        memberName = "?";
    }
    auto expr = std::make_unique<WithMemberExpr>(loc, memberName);
    return parsePostfix(std::move(expr));
}

} // namespace vb6c3
