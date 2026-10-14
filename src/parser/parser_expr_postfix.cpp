// vb6c3 - Pratt 表达式解析器 — 后缀表达式 + Case 值解析
// 由 src/parser/parser_expr.cpp 拆出（2026-09-17），纯搬移、零行为改动。

#include "parser/parser.hpp"
#include "common/canvas_drawing.hpp"  // 账 #232③: 哪条动词带尾巴、尾巴什么形状，只问这张表
#include <algorithm>
#include <cctype>

namespace vb6c3 {


// ============================================================
// 后缀表达式: .member, (args), !dict
// ============================================================

ExprPtr Parser::parsePostfix(ExprPtr expr) {
    while (true) {
        switch (cur_.kind) {
    case TokenKind::Dot: {
        auto loc = currentLoc();
        // P17.1: In With context, Debug.Print .Member should parse .Member
        // as WithMemberExpr argument, not chain MemberAccessExpr
        if (withDepth_ > 0 && expr->kind == ASTNodeKind::MemberAccessExpr) {
            auto& ma = static_cast<MemberAccessExpr&>(*expr);
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& obj = static_cast<IdentifierExpr&>(*ma.object);
                std::string objL = toLower(obj.name);
                std::string memL = toLower(ma.memberName);
                if (objL == "debug" && (memL == "print" || memL == "assert")) {
                    return expr;  // stop postfix, let arg parser handle .Member
                }
            }
        }
        // Fix 043c: In With context, ".Member1 .Member2" (space before 2nd dot)
        // is a sub call: .Member1(.Member2), NOT member access .Member1.Member2.
        // Detect space by comparing prevTok_ end column with cur_ start column.
        if (withDepth_ > 0 && expr->kind == ASTNodeKind::WithMemberExpr &&
            prevTok_.line == cur_.line &&
            prevTok_.column + prevTok_.length < cur_.column) {
            return expr;  // space detected — let caller parse .Member2 as argument
        }
        // Fix 077: In With context, "SubName .Field" (space before dot) is a bare
        // sub call where .Field is a With-block member access used as argument,
        // NOT a member access SubName.Field.  E.g.:
        //   pvAppendBitsToBuffer .Mode, 4, baQrCode, lBitLen
        // Without this, parser creates MemberAccessExpr(pvAppendBitsToBuffer, Mode)
        // instead of IdentifierExpr(pvAppendBitsToBuffer) + WithMemberExpr(.Mode).
        if (withDepth_ > 0 && expr->kind == ASTNodeKind::IdentifierExpr &&
            prevTok_.line == cur_.line &&
            prevTok_.column + prevTok_.length < cur_.column) {
            return expr;  // space detected — .XXX is a With-member argument, not member access
        }
        // Fix 099: 同 Fix 077, 扩展到无括号方法调用形态 "obj.Method .Field" —
        // obj.Method 后空格跟 .Field 时, .Field 是该调用的 With 块实参
        // (VB6 sub 风格调用: 实参不带括号, 以空格分隔), 而非成员链
        // obj.Method.Rs. 否则 "Users.Decode .Rs" 被解析成
        // MemberAccessExpr(MemberAccessExpr(U,Decode),Rs) → cgen 生成
        // ComCall(ComGetObjectProp(U,L"Decode"), L"Rs", ...) — .Rs 变成
        // 对 Decode 结果的链式 COM 调用且实参丢失 → C2198+C2039
        // (Demo.bas Db2: Users.Decode .Rs, VB6 实际语义 = Users.Decode(.Rs)).
        if (withDepth_ > 0 && expr->kind == ASTNodeKind::MemberAccessExpr &&
            prevTok_.line == cur_.line &&
            prevTok_.column + prevTok_.length < cur_.column) {
            return expr;  // space detected — .Field is a With-member argument to the obj.Method call
        }
        advance(); // consume '.'
        // VB6 允许关键字作为成员名: obj.Type, obj.Loop, etc.
        // expectName 只接受 Identifier 和软关键字, 这里扩展为接受所有带文本的 token
        std::string memberName;
        if (canBeName(cur_.kind)) {
            memberName = advance().text;
        } else if (!cur_.text.empty() && cur_.kind != TokenKind::EndOfFile &&
                   cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::Colon &&
                   cur_.kind != TokenKind::LeftParen && cur_.kind != TokenKind::RightParen &&
                   cur_.kind != TokenKind::Comma) {
            // 硬关键字也可作为成员名 (如 Type, Loop, Next 等)
            memberName = advance().text;
        } else {
            diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                std::string("expected member name after '.' (got ") +
                Token::kindToString(cur_.kind) + ")");
            memberName = "?";
        }
        expr = std::make_unique<MemberAccessExpr>(
            loc, std::move(expr), memberName);
        break;
    }

            case TokenKind::LeftParen: {
                // VB6 不区分数组索引和函数调用 → IndexOrCallExpr
                auto loc = currentLoc();
                advance(); // consume '('

                auto call = std::make_unique<IndexOrCallExpr>(loc, std::move(expr));

                // 解析参数列表
                if (cur_.kind != TokenKind::RightParen) {
                    size_t argIndex = 0;
                    do {
                        skipNewLines();

                        // VB6 允许在调用时覆盖传递方式: MyFunc(ByVal arg)
                        bool hasByValOverride = false;
                        if (cur_.kind == TokenKind::ByVal) {
                            advance(); // consume ByVal
                            hasByValOverride = true;
                        } else if (cur_.kind == TokenKind::ByRef) {
                            advance(); // consume ByRef
                        }

                        // 命名参数?  name := value (允许软关键字作参数名)
                        if (canBeName(cur_.kind) &&
                            next_.kind == TokenKind::Assign) {
                            // 命名参数
                            auto nameTok = advance(); // name
                            advance(); // consume ':='
                            auto val = parseExpression();
                            call->named.push_back({nameTok.text, std::move(val)});
                        } else if (cur_.kind == TokenKind::Comma || cur_.kind == TokenKind::RightParen) {
                            // M22: 空参数占位 - VB6允许 MsgBox("hi", , "title")
                            // Fix 142: 同时记录省略索引, 供代码生成按形参缺省值填充
                            auto _ph = std::make_unique<LiteralExpr>(currentLoc(), LiteralKind::Long, "0");
                            _ph->longValue = 0;  // Union与intValue共享内存, 必须显式设置longValue
                            call->positional.push_back(std::move(_ph));
                            call->omittedArgs.insert(argIndex);
                        } else {
                            // 位置参数
                            auto arg = parseExpression();
                            call->positional.push_back(std::move(arg));
                        }

                        // Fix 072: 记录 ByVal 覆盖的参数索引
                        if (hasByValOverride) {
                            call->byvalOverrides.insert(argIndex);
                        }

                        argIndex++;
                        skipNewLines();
                    } while (match(TokenKind::Comma));
                }

                expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
                       "expected ')'");

                // Fix 102: VB6 图形方法坐标语法
                //   obj.Line (x1, y1)-(x2, y2)[, color][, BF | B | F]
                // `(x1, y1)` 已按普通实参表解析完毕; 紧随的 `-(x2, y2)` 必须在此吸收,
                // 否则外层 parseExpression 会把它当作中缀减法, 而右操作数 `(x2, y2)`
                // 的括号内含逗号 → "expected ')'", 整行解析崩坏并连锁破坏其后的
                // If/End If 配对 (Charts 2020 ppProgressCircular.pag 297/299/474).
                // 吸收后统一为 IndexOrCallExpr(callee=obj.Line,
                // 实参 = x1, y1, x2, y2[, color][, fillMode]), 由后端按控件类型发射。
                //
                // 2026-10-06: **Circle / PSet / Point 同样要吸收**, 否则它们的
                // 坐标/半径尾巴会漏到外层变成独立表达式, 而残留的 `0(...)` 桩
                // 语句既 C2064 又把半径丢掉 (实测 `Me.Circle (300,100),40` 发出
                // `vb6_Form_Circle(..., 300, 100, 0, ...)` 且下一行是残桩)。
                // 三者的尾巴形态:
                //   Circle (x, y), radius [, color] [, start] [, end] [, aspect]
                //   PSet   [Step] (x, y) [, color]
                //   Point  (x, y)                          ← 无尾巴, 无需吸收
                // ⇒ 需要吸收的是 Circle 的 `, radius[, ...]` 与 PSet 的 `, color`。
                // 账 #232③: 尾巴的名字问 canvas_drawing.hpp，而**两种写法都问**。
                // 以前只有 MemberAccessExpr 那一形进得来（`Me.Line` / `Pic1.Circle`），
                // 窗体上裸写的动词 callee 是 IdentifierExpr ⇒ `Circle (20, 21), 22` 的 `, 22`
                // 与 `Line (0, 0)-(10, 10)` 的 `-(10, 10)` 都漏到外层表达式（实测三条
                // VB2001 expected ')' / VB2003 / VB2002，整行崩坏）。折叠动作住在语义层（在 parser 之后），
                // 所以 parser 这里问的只能是**名字**，不能问符号表。
                std::string tailName;
                if (call->callee) {
                    if (call->callee->kind == ASTNodeKind::MemberAccessExpr) {
                        tailName = toLower(static_cast<MemberAccessExpr&>(*call->callee).memberName);
                    } else if (call->callee->kind == ASTNodeKind::IdentifierExpr) {
                        tailName = toLower(static_cast<IdentifierExpr&>(*call->callee).name);
                    }
                }
                const CanvasTailKind tailKind = canvasVerbTailKind(tailName);
                if (cur_.kind == TokenKind::Comma) {
                    if (tailKind == CANVAS_TAIL_COORD) {
                        // 收不动就停 (下一个 token 是语句终止符), 剩下的交给外层。
                        while (match(TokenKind::Comma)) {
                            if (cur_.kind == TokenKind::NewLine ||
                                cur_.kind == TokenKind::Colon ||
                                cur_.kind == TokenKind::EndOfFile ||
                                cur_.kind == TokenKind::RightParen) {
                                break;
                            }
                            call->positional.push_back(parseExpression());
                        }
                    }
                }
                if (cur_.kind == TokenKind::Minus && next_.kind == TokenKind::LeftParen) {
                    if (tailKind == CANVAS_TAIL_TWO_POINT) {
                        advance();  // 消费 '-'
                        advance();  // 消费 '('
                        auto x2 = parseExpression();
                        expect(TokenKind::Comma, DiagnosticID::ParseExpectedToken,
                               "expected ',' in Line (x1,y1)-(x2,y2)");
                        auto y2 = parseExpression();
                        expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
                               "expected ')' in Line (x1,y1)-(x2,y2)");
                        call->positional.push_back(std::move(x2));
                        call->positional.push_back(std::move(y2));
                        // 可选后续参数: ", color" / ", color, BF|B"
                        int trailingIdx = 0;
                        while (match(TokenKind::Comma)) {
                            if (cur_.kind == TokenKind::NewLine ||
                                cur_.kind == TokenKind::Colon ||
                                cur_.kind == TokenKind::EndOfFile) {
                                break;
                            }
                            // VB6 的 `Line` 尾参只有 color 与 style 两格，而 style 那格的
                            // B / C / F 是**语法旗标**不是名字 (账 #220)。折成数值就地定死：
                            // 名字一旦进 AST，发码就把它原样发出去，靠 RTL 里两枚裸名全局
                            // (const int32_t B / BF) 接住 —— 任何工程有个模块级变量叫 B 就撞车。
                            // 只认 style 位置 (trailingIdx 为 1，即 color 已给出)：
                            // `Line (a,b)-(c,d), B` 那一格按 VB6 是 color，用户的 B 必须照旧成立。
                            // 位口径与 RTL `vb6_ControlLine` 是**同一张表**（账 #221）：
                            // B=1 矩形、C=2 椭圆、F=4 填充，按字母逐个置位 ⇒ BF=5、CF=6。
                            // 订正一处历史: 账 #220 那一刀为了不改语义沿用了旧全局的 1/2，
                            // 于是 BF 与"C/F 两形"都没口径可依 —— 这一格把它收成字母位。
                            if (trailingIdx == 1 && cur_.kind == TokenKind::Identifier) {
                                const std::string optWord = toLower(cur_.text);
                                int styleBits = 0;
                                bool allFlagLetters = !optWord.empty();
                                for (char ch : optWord) {
                                    if (ch == 'b') { styleBits |= 1; }
                                    else if (ch == 'c') { styleBits |= 2; }
                                    else if (ch == 'f') { styleBits |= 4; }
                                    else { allFlagLetters = false; break; }
                                }
                                if (allFlagLetters) {
                                    auto locOpt = currentLoc();
                                    advance();  // 消费 B / C / F 那一串
                                    const std::string bitText = std::to_string(styleBits);
                                    auto lit = std::make_unique<LiteralExpr>(
                                        locOpt, LiteralKind::Integer, bitText);
                                    lit->intValue = styleBits;
                                    call->positional.push_back(std::move(lit));
                                    trailingIdx++;
                                    continue;
                                }
                            }
                            call->positional.push_back(parseExpression());
                            trailingIdx++;
                        }
                    }
                }

                expr = std::move(call);
                break;
            }

            case TokenKind::Exclamation: {
                // 字典访问: expr!key
                auto loc = currentLoc();
                advance(); // consume '!'
                auto key = expectName("expected identifier after '!'");
                expr = std::make_unique<DictionaryAccessExpr>(
                    loc, std::move(expr), key.text);
                break;
            }

            default:
                return expr;  // 没有更多后缀
        }
    }
}

// ============================================================
// Case 值解析 (Select Case 专用)
// ============================================================

CaseClause::CaseValue Parser::parseCaseValue() {
    CaseClause::CaseValue cv;

    // Case Is > 0 形式
    if (cur_.kind == TokenKind::IsKeyword || cur_.kind == TokenKind::Is) {
        advance(); // consume 'Is'
        cv.isIsClause = true;

        // 比较运算符
        if (checkAny({TokenKind::LessThan, TokenKind::GreaterThan,
                      TokenKind::LessEqual, TokenKind::GreaterEqual,
                      TokenKind::Equals, TokenKind::NotEquals})) {
            // VB6 的 `Is` 是 Select 的测试表达式本身, 不是标识符 (账 #217)。占位左操作数
            // 只为借一次优先级解析 (右操作数按比较符的绑定力截断, 不能吃进后面的 `Or`/`,`)
            // —— 它不进 AST: 留在树里就被语义层登记成隐式 Variant, 每形一条 VB3001 或一枚没人用的局部。
            auto placeholder = std::make_unique<IdentifierExpr>(currentLoc(), "Is");
            auto bp = getBindingPower(cur_.kind);
            auto expr = parseLeftDenotation(std::move(placeholder), bp.l_bp);
            if (expr && expr->kind == ASTNodeKind::BinaryExpr) {
                auto& bin = static_cast<BinaryExpr&>(*expr);
                cv.relOp = bin.op;
                cv.hasRelOp = true;
                cv.value = std::move(bin.right);
            }
        }
        // 无比较符的 `Case Is`: 发码侧只看 isIsClause, 值留空 (改动前放的那枚
        // IdentifierExpr("Is") 从来没人读)。
    } else {
        // 普通值或范围
        cv.value = parseExpression();

        // Case 1 To 10 范围形式
        if (match(TokenKind::To)) {
            cv.toValue = parseExpression();
        }
    }

    return cv;
}

} // namespace vb6c3
