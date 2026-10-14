// vb6c3 - 语句解析器
// VB6 块语句 + 单行语句

#include "parser/parser.hpp"

namespace vb6c3 {

// --- parser_stmt_control.cpp: 控制流语句（If / For / ForEach / DoLoop / WhileWend / SelectCase / With） ---



// ============================================================
// If 语句
// ============================================================

std::unique_ptr<IfStmt> Parser::parseIfStmt() {
    auto loc = currentLoc();
    advance(); // consume 'If'

    auto condition = parseExpression();
    expect(TokenKind::Then, DiagnosticID::ParseInvalidIf,
           "expected 'Then' after If condition");

    // 单行 If?  检查 Then 后面是否紧接语句 (无换行)
    // 跳过 Then 后的可选冒号: If x Then: stmt
    if (cur_.kind == TokenKind::Colon) {
        advance();
    }
    if (cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::EndOfFile) {
        // 单行 If...Then...[Else...]
        StmtList thenBody;
        inSingleLineIf_++;
        // Fix 082: `If cond Then Else stmt` — Then 分支可以为空 (语义: 条件为真则
        // 什么都不做)。样例工程 VisualStyles.bas:399 即此形式, 且该文件确属工程成员
        // (VBFlexGridDemo.vbp:10), 即 VB6 实际接受这种写法。必须先判空再 parseStatement(),
        // 否则 Else 会被当成语句首 token -> VB2002 "unexpected token in statement: Else"。
        if (cur_.kind != TokenKind::Else) {
            auto stmt = parseStatement();
            if (stmt) thenBody.push_back(std::move(stmt));
        }

        // 解析 Then 后续的冒号分隔语句: If x Then stmt1: stmt2: stmt3
        while (cur_.kind == TokenKind::Colon) {
            advance(); // consume ':'
            if (cur_.kind == TokenKind::Else) break;
            auto moreStmt = parseStatement();
            if (moreStmt) thenBody.push_back(std::move(moreStmt));
        }

        StmtList elseBody;
        if (match(TokenKind::Else)) {
            auto elseStmt = parseStatement();
            if (elseStmt) elseBody.push_back(std::move(elseStmt));
            // Else 也可以有冒号分隔的多语句
            while (cur_.kind == TokenKind::Colon) {
                advance(); // consume ':'
                auto moreStmt = parseStatement();
                if (moreStmt) elseBody.push_back(std::move(moreStmt));
            }
        }
        inSingleLineIf_--;

        return std::make_unique<IfStmt>(loc, std::move(condition),
            std::move(thenBody),
            std::vector<std::unique_ptr<ElseIfClause>>{},
            std::move(elseBody), true);
    }

    // 多行 If...Then
    skipNewLines();
    // 使用 parseBlockUntil 包含 End 作为停止条件
    // (不能只用 parseBlock(ElseIf, Else), 因为没有 Else 的 If 块
    //  遇到 End If 时无法终止, 导致无限循环)
    StmtList thenBody = parseBlockUntil({TokenKind::ElseIf, TokenKind::Else,
                                         TokenKind::End});

    std::vector<std::unique_ptr<ElseIfClause>> elseIfs;
    while (cur_.kind == TokenKind::ElseIf) {
        auto elseifLoc = currentLoc();
        advance(); // consume 'ElseIf'
        auto elseifCond = parseExpression();
        expect(TokenKind::Then, DiagnosticID::ParseInvalidIf,
               "expected 'Then' after ElseIf condition");
        skipNewLines();
        auto elseifBody = parseBlockUntil({TokenKind::ElseIf, TokenKind::Else,
                                           TokenKind::End});
        elseIfs.push_back(std::make_unique<ElseIfClause>(elseifLoc,
            std::move(elseifCond), std::move(elseifBody)));
    }

    StmtList elseBody;
    if (match(TokenKind::Else)) {
        skipNewLines();
        elseBody = parseBlockUntil({TokenKind::End});
    }

    // End If
    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock,
           "expected 'End If'");
    expect(TokenKind::If, DiagnosticID::ParseMismatchedBlock,
           "expected 'End If'");

    return std::make_unique<IfStmt>(loc, std::move(condition),
        std::move(thenBody), std::move(elseIfs), std::move(elseBody), false);
}

// ============================================================
// For 语句
// ============================================================

// VB6: `Next var1, var2, ...` — 一条 Next 关闭多层嵌套 For 循环 (内层在前,
// 外层在后)。真实的 Next token 由当前循环消费; 逗号后多余的变量排队进
// pendingNextVars_, 供各外层循环在检查 Next 时当作"已消费的合成 Next"弹出。
void Parser::consumeNextClause(const std::string& /*loopVar*/) {
    auto canConsumeVar = [&]() {
        return canBeName(cur_.kind) && cur_.kind != TokenKind::NewLine &&
               cur_.kind != TokenKind::Colon && cur_.kind != TokenKind::EndOfFile;
    };
    if (cur_.kind != TokenKind::Next) {
        if (!pendingNextVars_.empty()) {
            pendingNextVars_.erase(pendingNextVars_.begin());
            return;
        }
        expect(TokenKind::Next, DiagnosticID::ParseMismatchedBlock,
               "expected 'Next' to close For loop");
        return;
    }
    advance(); // consume 'Next'
    pendingNextVars_.clear(); // 真实 Next 出现即已重新同步
    if (canConsumeVar()) advance(); // 本循环的可选变量
    while (cur_.kind == TokenKind::Comma) {
        advance(); // consume ','
        if (canConsumeVar()) {
            pendingNextVars_.push_back(cur_.text);
            advance();
        }
    }
}

std::unique_ptr<ForStmt> Parser::parseForStmt() {
    auto loc = currentLoc();
    advance(); // consume 'For'

    auto varTok = expectName("expected variable name after 'For'");
    expect(TokenKind::Equals, DiagnosticID::ParseInvalidFor,
           "expected '=' in For statement");

    auto start = parseExpression();
    expect(TokenKind::To, DiagnosticID::ParseInvalidFor,
           "expected 'To' in For statement");
    auto end = parseExpression();

    ExprPtr step;
    if (match(TokenKind::Step)) {
        step = parseExpression();
    }

    // 单行 `For i = 1 To 10: stmt: Next` 与多行形式统一由 parseBlockUntil
    // 起始处的 skipStatementSeparators() 处理 (冒号 ≡ 换行)。
    auto body = parseBlockUntil({TokenKind::Next});

    // Next [var] — 必须消费 Next, 可选的循环变量
    consumeNextClause(varTok.text);

    return std::make_unique<ForStmt>(loc, varTok.text,
        std::move(start), std::move(end), std::move(step), std::move(body));
}

std::unique_ptr<ForEachStmt> Parser::parseForEachStmt() {
    auto loc = currentLoc();
    advance(); // consume 'For'
    advance(); // consume 'Each'

    auto varTok = expectName("expected variable name");
    expect(TokenKind::In, DiagnosticID::ParseInvalidFor,
           "expected 'In' in For Each statement");
    auto collection = parseExpression();

    // 单行 `For Each x In col: stmt: Next` 由 parseBlockUntil 统一处理。
    auto body = parseBlockUntil({TokenKind::Next});

    expect(TokenKind::Next, DiagnosticID::ParseMismatchedBlock,
           "expected 'Next'");
    if (canBeName(cur_.kind) && cur_.kind != TokenKind::NewLine &&
        cur_.kind != TokenKind::Colon && cur_.kind != TokenKind::EndOfFile) {
        advance();
    }

    return std::make_unique<ForEachStmt>(loc, varTok.text,
        std::move(collection), std::move(body));
}

// ============================================================
// Do Loop 语句
// ============================================================

std::unique_ptr<DoLoopStmt> Parser::parseDoLoopStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Do'

    DoLoopKind loopKind;
    ExprPtr condition;

    if (match(TokenKind::While)) {
        loopKind = DoLoopKind::DoWhileLoop;
        condition = parseExpression();
        skipNewLines();
    } else if (match(TokenKind::Until)) {
        loopKind = DoLoopKind::DoUntilLoop;
        condition = parseExpression();
        skipNewLines();
    } else {
        loopKind = DoLoopKind::DoLoop;  // 无条件 Do...Loop
        skipNewLines();
    }

    // 单行 `Do [While|Until cond]: stmt: Loop` 与多行形式统一由 parseBlockUntil
    // 起始处的 skipStatementSeparators() 处理 (冒号 ≡ 换行)。
    auto body = parseBlockUntil({TokenKind::Loop});

    expect(TokenKind::Loop, DiagnosticID::ParseMismatchedBlock,
           "expected 'Loop'");

    // Do...Loop While|Until
    if (loopKind == DoLoopKind::DoLoop) {
        if (match(TokenKind::While)) {
            loopKind = DoLoopKind::DoLoopWhile;
            condition = parseExpression();
        } else if (match(TokenKind::Until)) {
            loopKind = DoLoopKind::DoLoopUntil;
            condition = parseExpression();
        }
    }

    return std::make_unique<DoLoopStmt>(loc, loopKind, std::move(condition), std::move(body));
}

// ============================================================
// While...Wend
// ============================================================

std::unique_ptr<WhileWendStmt> Parser::parseWhileWendStmt() {
    auto loc = currentLoc();
    advance(); // consume 'While'
    auto condition = parseExpression();

    // 单行 `While cond: stmt: Wend` 与多行形式统一由 parseBlockUntil 起始处的
    // skipStatementSeparators() 处理 (冒号 ≡ 换行)。
    auto body = parseBlockUntil({TokenKind::Wend});
    expect(TokenKind::Wend, DiagnosticID::ParseMismatchedBlock,
           "expected 'Wend'");
    return std::make_unique<WhileWendStmt>(loc, std::move(condition), std::move(body));
}

// ============================================================
// Select Case
// ============================================================

std::unique_ptr<SelectCaseStmt> Parser::parseSelectCaseStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Select'
    expect(TokenKind::Case, DiagnosticID::ParseInvalidSelect,
           "expected 'Case' after 'Select'");
    auto testExpr = parseExpression();
    skipNewLines();

    std::vector<std::unique_ptr<CaseClause>> cases;
    StmtList elseCase;

    while (cur_.kind == TokenKind::Case && next_.kind != TokenKind::Else) {
        auto caseLoc = currentLoc();
        advance(); // consume 'Case'

        // 解析 Case 值列表
        std::vector<CaseClause::CaseValue> values;
        do {
            skipNewLines();
            values.push_back(parseCaseValue());
        } while (match(TokenKind::Comma));

        skipNewLines();
        // VB6 允许 Case N: statement (冒号分隔)
        match(TokenKind::Colon);
        auto body = parseBlockUntil({TokenKind::Case, TokenKind::End});
        cases.push_back(std::make_unique<CaseClause>(caseLoc,
            std::move(values), std::move(body)));
    }

    // Case Else
    if (cur_.kind == TokenKind::Case && next_.kind == TokenKind::Else) {
        advance(); // consume 'Case'
        advance(); // consume 'Else'
        skipNewLines();
        // VB6 允许 Case Else: statement (冒号分隔)
        match(TokenKind::Colon);
        elseCase = parseBlockUntil({TokenKind::End});
    }

    // End Select
    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Select'");
    expect(TokenKind::Select, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Select'");

    return std::make_unique<SelectCaseStmt>(loc, std::move(testExpr),
        std::move(cases), std::move(elseCase));
}

// ============================================================
// With 语句
// ============================================================

std::unique_ptr<WithStmt> Parser::parseWithStmt() {
    auto loc = currentLoc();
    advance(); // consume 'With'
    auto object = parseExpression();

    // 单行 `With x: .a = 1: End With` 与多行形式统一由 parseBlockUntil 起始处的
    // skipStatementSeparators() 处理 (冒号 ≡ 换行)。
    withDepth_++;
    auto body = parseBlockUntil({TokenKind::End});
    withDepth_--;

    // End With
    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock,
           "expected 'End With'");
    expect(TokenKind::With, DiagnosticID::ParseMismatchedBlock,
           "expected 'End With'");

    return std::make_unique<WithStmt>(loc, std::move(object), std::move(body));
}

} // namespace vb6c3
