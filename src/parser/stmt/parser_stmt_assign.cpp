// vb6c3 - 语句解析器
// VB6 块语句 + 单行语句

#include "parser/parser.hpp"

namespace vb6c3 {

// --- parser_stmt_assign.cpp: 赋值与声明语句（Set / Let / Call / Dim / ReDim / Const / Static / Erase / RaiseEvent + 标号与调用消歧） ---


// ============================================================
// Set / Let / Call
// ============================================================

std::unique_ptr<SetStmt> Parser::parseSetStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Set'
    // '=' 是赋值号, 不是比较运算符; 用 minBp > '='(l_bp=8) 避免表达式吃掉 '='
    auto target = parseExpression(9);
    expect(TokenKind::Equals, DiagnosticID::ParseExpectedToken,
           "expected '=' in Set statement");
    auto value = parseExpression();
    return std::make_unique<SetStmt>(loc, std::move(target), std::move(value));
}

std::unique_ptr<LetStmt> Parser::parseLetStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Let'
    // 同 Set, '=' 是赋值号
    auto target = parseExpression(9);
    expect(TokenKind::Equals, DiagnosticID::ParseExpectedToken,
           "expected '=' in Let statement");
    auto value = parseExpression();
    return std::make_unique<LetStmt>(loc, std::move(target), std::move(value));
}

std::unique_ptr<CallStmt> Parser::parseCallStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Call'
    auto callee = parseExpression();
    return std::make_unique<CallStmt>(loc, std::move(callee));
}

// ============================================================
// Dim / ReDim / Const / Static (过程体内)
// ============================================================

StmtPtr Parser::parseDimStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Dim'
    return wrapBodyDecls(loc, parseVariableDeclList(AccessLevel::Private, false));
}

// 体级声明只有一种形状: 一条声明符一条 LocalDeclStmt。parse*DeclList 在多声明符时返回
// MultiDecl, 这里就地展开 —— 以前 Dim 自己手写一遍展开 (那份副本漏了 parseVariableDecl
// 里的 WithEvents 与「后缀即类型」两步, 于是 `Dim a&, b&` 的第二枚落回 Variant),
// 而 Const/Static/Public 三条把 MultiDecl 原样交给语义层, 那儿的 switch 不认这个 kind,
// 于是一枚名字都不登记、每条使用报一条 VB3001 (账 #215)。
StmtPtr Parser::wrapBodyDecls(SourceLocation loc, DeclPtr decl) {
    if (decl->kind != ASTNodeKind::MultiDecl) {
        return std::make_unique<LocalDeclStmt>(loc, std::move(decl));
    }
    auto& multi = static_cast<MultiDecl&>(*decl);
    StmtList stmts;
    for (auto& d : multi.declarations) {
        stmts.push_back(std::make_unique<LocalDeclStmt>(loc, std::move(d)));
    }
    return std::make_unique<Block>(loc, std::move(stmts));
}

// 由点链字符串构造表达式 (账 #186 起为 ReDim / Erase **共用**的单一出口)。
// 前置 '.' 表示 With 块成员 (WithMemberExpr), 其余逐段构造 MemberAccessExpr。
ExprPtr Parser::buildDottedNameExpr(const std::string& nm, SourceLocation l) {
    if (!nm.empty() && nm[0] == '.') {
        return std::make_unique<WithMemberExpr>(l, nm.substr(1));
    }
    ExprPtr e;
    size_t pos = 0;
    for (;;) {
        size_t d = nm.find('.', pos);
        std::string seg = (d == std::string::npos) ? nm.substr(pos)
                                                   : nm.substr(pos, d - pos);
        if (!e) e = std::make_unique<IdentifierExpr>(l, seg);
        else e = std::make_unique<MemberAccessExpr>(l, std::move(e), seg);
        if (d == std::string::npos) break;
        pos = d + 1;
    }
    return e;
}

std::unique_ptr<ReDimStmt> Parser::parseReDimStmt() {
    auto loc = currentLoc();
    bool preserve = false;
    advance(); // consume 'ReDim'
    if (match(TokenKind::Preserve)) {
        preserve = true;
    }

    // 解析变量名, 支持点访问: uOutput.Buffer 和 With块: .Member
    std::string varName;
    if (match(TokenKind::Dot)) {
        varName = ".";
    }
    auto varTok = expectName("expected variable name in ReDim");
    varName += varTok.text;
    while (match(TokenKind::Dot)) {
        if (canBeName(cur_.kind)) {
            varName += "." + advance().text;
        } else if (!cur_.text.empty() && cur_.kind != TokenKind::EndOfFile &&
                   cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::Colon &&
                   cur_.kind != TokenKind::LeftParen && cur_.kind != TokenKind::RightParen &&
                   cur_.kind != TokenKind::Comma) {
            varName += "." + advance().text;
        } else {
            break;
        }
    }
    // Fix 100: 复杂目标 (带下标的成员链) — ReDim m_Serie(i).PT(n)
    // 解析策略 (无回溯): 先按原逻辑取点链名, 再解析 '(' ... ')'。
    // 若该 ')' 之后紧跟 '.', 说明刚才解析到的其实是「下标」而非 ReDim 维度 →
    // 把它包装成 IndexOrCallExpr 作为目标的一部分, 继续解析 .Member 与后续括号,
    // 直到某个 ')' 之后不是 '.' 为止, 那一组括号才是真正的 ReDim 维度。
    //
    // 返回: 括号内的维度/下标列表。To 语法在维度中表示下界, 在下标中非法
    // (仍解析, 复杂路径取 upper)。
    auto parseParenList = [this]() -> std::vector<ReDimStmt::Dimension> {
        std::vector<ReDimStmt::Dimension> list;
        do {
            ReDimStmt::Dimension dim;
            dim.lower = nullptr;
            dim.upper = parseExpression();
            if (match(TokenKind::To)) {
                dim.lower = std::move(dim.upper);
                dim.upper = parseExpression();
            }
            list.push_back(std::move(dim));
        } while (match(TokenKind::Comma));
        return list;
    };

    // 由点链字符串构造表达式 (复用于复杂目标的基名)。账 #186 起出口在 Parser::buildDottedNameExpr
    // —— Erase 与 ReDim 共用同一份，别再抄第二遍。
    auto buildNameExpr = [&](const std::string& nm, SourceLocation l) -> ExprPtr {
        return buildDottedNameExpr(nm, l);
    };

    // 把一组下标维度包装为 IndexOrCallExpr(base[i, j])
    auto wrapIndex = [](ExprPtr base, std::vector<ReDimStmt::Dimension>& idxList,
                        SourceLocation l) -> ExprPtr {
        auto call = std::make_unique<IndexOrCallExpr>(l, std::move(base));
        for (auto& d : idxList) {
            call->positional.push_back(d.lower ? std::move(d.lower) : std::move(d.upper));
        }
        return call;
    };

    expect(TokenKind::LeftParen, DiagnosticID::ParseExpectedToken,
           "expected '(' after ReDim variable");

    std::vector<ReDimStmt::Dimension> dims = parseParenList();

    expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
           "expected ')' after ReDim dimensions");

    ExprPtr targetExpr;
    if (cur_.kind == TokenKind::Dot) {
        // 上一步解析到的 '(' ... ')' 实为下标 → 转入复杂目标路径
        targetExpr = wrapIndex(buildNameExpr(varName, loc), dims, loc);
        for (;;) {
            expect(TokenKind::Dot, DiagnosticID::ParseExpectedToken,
                   "expected '.' in ReDim target");
            Token memTok = expectName("expected member name in ReDim target");
            targetExpr = std::make_unique<MemberAccessExpr>(
                loc, std::move(targetExpr), memTok.text);
            if (cur_.kind != TokenKind::LeftParen) break;
            advance();  // consume '('
            auto next = parseParenList();
            expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
                   "expected ')' in ReDim target");
            if (cur_.kind == TokenKind::Dot) {
                targetExpr = wrapIndex(std::move(targetExpr), next, loc);
                continue;
            }
            dims = std::move(next);  // 这才是真正的 ReDim 维度
            break;
        }
    }

    TypeRefPtr asType;
    if (match(TokenKind::As)) {
        asType = parseTypeRef();
    }

    auto stmt = std::make_unique<ReDimStmt>(loc, preserve, varName,
        std::move(dims), std::move(asType));
    stmt->targetExpr = std::move(targetExpr);
    return stmt;
}

StmtPtr Parser::parseConstStmtInBody() {
    auto loc = currentLoc();
    // 不需要 advance() — parseConstDeclList -> parseConstDecl 会消费 'Const'
    return wrapBodyDecls(loc, parseConstDeclList(AccessLevel::Private));
}

StmtPtr Parser::parseStaticStmtInBody() {
    auto loc = currentLoc();
    advance(); // consume 'Static'
    // Static x As Long  or  Static Sub ...
    if (cur_.kind == TokenKind::Sub) {
        auto subDecl = parseSubDecl(AccessLevel::Private, true);
        return std::make_unique<LocalDeclStmt>(loc, std::move(subDecl));
    }
    if (cur_.kind == TokenKind::Function) {
        auto funcDecl = parseFunctionDecl(AccessLevel::Private, true);
        return std::make_unique<LocalDeclStmt>(loc, std::move(funcDecl));
    }
    return wrapBodyDecls(loc, parseVariableDeclList(AccessLevel::Private, true));
}

StmtPtr Parser::parseAccessDeclInBody() {
    // Public/Private x As Long  (过程体内的声明)
    auto loc = currentLoc();
    AccessLevel access = (cur_.kind == TokenKind::Public)
        ? AccessLevel::Public : AccessLevel::Private;
    advance();
    return wrapBodyDecls(loc, parseVariableDeclList(access, false));
}

// ============================================================
// Erase / RaiseEvent
// ============================================================

std::unique_ptr<EraseStmt> Parser::parseEraseStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Erase'
    std::vector<std::string> names;
    std::vector<ExprPtr> targets;

    // 一段标识符 (与旧口径一致: 关键字位的名字也照字面收下)
    auto takeNamePiece = [this]() -> std::string {
        if (canBeName(cur_.kind)) return advance().text;
        if (!cur_.text.empty() && cur_.kind != TokenKind::EndOfFile &&
            cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::Colon &&
            cur_.kind != TokenKind::Comma && cur_.kind != TokenKind::LeftParen &&
            cur_.kind != TokenKind::RightParen) return advance().text;
        return std::string();
    };

    auto wrapSubscripts = [](ExprPtr base, std::vector<ExprPtr>& subs,
                             SourceLocation l) -> ExprPtr {
        auto call = std::make_unique<IndexOrCallExpr>(l, std::move(base));
        for (auto& s : subs) call->positional.push_back(std::move(s));
        return call;
    };

    // 括号里的下标列表 (账 #186)。`To` 在 ReDim 的维度里是下界, 在**下标**里非法。
    auto parseSubscriptList = [this](bool& toSeen) -> std::vector<ExprPtr> {
        std::vector<ExprPtr> subs;
        toSeen = false;
        while (cur_.kind != TokenKind::RightParen && cur_.kind != TokenKind::EndOfFile &&
               cur_.kind != TokenKind::NewLine) {
            subs.push_back(parseExpression());
            if (match(TokenKind::To)) { toSeen = true; parseExpression(); }
            if (!match(TokenKind::Comma)) break;
        }
        return subs;
    };

    // 点链继续留在名字里 (没有下标时的旧行为, 含 Fix 082 的空括号)
    auto finishPlainName = [this, &takeNamePiece](std::string& name) {
        while (match(TokenKind::Dot)) {
            std::string m = takeNamePiece();
            if (m.empty()) break;
            name += "." + m;
        }
    };

    // Fix 082 的另一半: 空括号可以挂在**点链的末尾** (`Erase obj.Field()` /
    // `Erase .MaxWidths()`, VBFlexGrid 6590 就是后者)。吃完整条点链之后再吃这一组括号,
    // 顺序不能反 —— 反了就留下一个裸 '(' 变成 VB2003/VB2002。
    auto dropOptionalEmptyParens = [this, &loc, &parseSubscriptList](const std::string& name) {
        if (cur_.kind != TokenKind::LeftParen) return;
        advance();  // '('
        bool to = false;
        std::vector<ExprPtr> subs = parseSubscriptList(to);
        if (!match(TokenKind::RightParen)) {
            diag_.error(DiagnosticID::ParseExpectedToken, loc,
                "Erase 目标的下标缺少 ')'");
            return;
        }
        if (!subs.empty() || to) {
            diag_.error(DiagnosticID::ParseExpectedToken, loc,
                std::string("Erase 不支持带下标的目标 (") + name
                + ") —— 只有 'Erase arr' 与 'Erase arr()' 是销毁整个数组 (indexed Erase target not supported here)");
        }
    };

    // Fix 082: VB6 允许 `Erase arr()` 的可选空括号 (Common.bas:423
    // `Erase MsgBoxHelpData()`) —— 空括号照旧丢弃。
    // 账 #186: 带**非空**下标的目标此前一律 VB2001，卡住了真工程的一句合法代码
    // (`PropPagFMR.pag:720 Erase m_tvFiles(lIndex).bvData` —— 销毁 UDT 那一格里的动态数组)。
    // EraseStmt 只能表达"名字"，所以下标必须配一棵表达式树才发得出来 (见 ast_stmt.hpp)：
    //   base(subs).member   → targets[i] 非空，发码侧 emitExpr 出 VB6_SA_AT(...) 那样的左值
    //   base(subs)          → 仍报诊断。VB6 里这形只在元素是 Variant(装着数组) 时合法，
    //                         本仓没那条通路；静默降级成"销毁整个数组"比编不过更坏。
    auto parseOneTarget = [&](std::string& nameOut, ExprPtr& exprOut) {
        std::string name;
        if (match(TokenKind::Dot)) name = ".";
        name += takeNamePiece();
        exprOut = nullptr;
        if (cur_.kind != TokenKind::LeftParen) {
            finishPlainName(name);
            dropOptionalEmptyParens(name);   // `Erase obj.Field()` / `Erase .Field()`
            nameOut = name;
            return;
        }
        advance();  // '('
        bool toSeen = false;
        std::vector<ExprPtr> subs = parseSubscriptList(toSeen);
        if (!match(TokenKind::RightParen)) {
            diag_.error(DiagnosticID::ParseExpectedToken, loc,
                "Erase 目标的下标缺少 ')'");
            nameOut = name;
            return;
        }
        if (subs.empty()) {           // `Erase arr()`
            finishPlainName(name);
            dropOptionalEmptyParens(name);   // `Erase obj.Field()` / `Erase .Field()`
            nameOut = name;
            return;
        }
        if (toSeen || cur_.kind != TokenKind::Dot) {
            diag_.error(DiagnosticID::ParseExpectedToken, loc,
                toSeen ? "Erase 的下标里不支持 To 语法 (To is not allowed in an Erase subscript; it only appears in ReDim dimensions)"
                       : std::string("Erase 的带下标目标必须是成员数组形式 (Erase arr(i).data)，'")
                         + name + "(i)' 暂不支持 —— 静默改成销毁整个数组会改变语义 (indexed Erase target not supported here)");
            nameOut = name;
            return;
        }
        ExprPtr expr = wrapSubscripts(buildDottedNameExpr(name, loc), subs, loc);
        std::string dotted = name;
        while (cur_.kind == TokenKind::Dot) {
            advance();  // '.'
            std::string m = takeNamePiece();
            if (m.empty()) {
                diag_.error(DiagnosticID::ParseExpectedToken, loc, "Erase 目标的 '.' 后缺少成员名");
                break;
            }
            expr = std::make_unique<MemberAccessExpr>(loc, std::move(expr), m);
            dotted += "." + m;
            if (cur_.kind != TokenKind::LeftParen) continue;
            advance();  // '('
            bool to2 = false;
            std::vector<ExprPtr> subs2 = parseSubscriptList(to2);
            if (!match(TokenKind::RightParen)) {
                diag_.error(DiagnosticID::ParseExpectedToken, loc,
                    "Erase 目标的下标缺少 ')'");
                break;
            }
            if (!subs2.empty()) {   // 末段带下标 = 要销毁"数组里的一格"，Erase 无此语义
                diag_.error(DiagnosticID::ParseExpectedToken, loc,
                    "Erase 不支持对成员数组再带下标 (Erase arr(i).data(j))");
                break;
            }
        }
        nameOut = dotted;
        exprOut = std::move(expr);
    };

    std::string nm;
    ExprPtr ex;
    parseOneTarget(nm, ex);
    names.push_back(std::move(nm));
    targets.push_back(std::move(ex));
    while (match(TokenKind::Comma)) {
        std::string n2;
        ExprPtr e2;
        parseOneTarget(n2, e2);
        names.push_back(std::move(n2));
        targets.push_back(std::move(e2));
    }
    auto stmt = std::make_unique<EraseStmt>(loc, std::move(names));
    stmt->targets = std::move(targets);
    return stmt;
}

std::unique_ptr<RaiseEventStmt> Parser::parseRaiseEventStmt() {
    auto loc = currentLoc();
    advance(); // consume 'RaiseEvent'
    auto nameTok = expectName("expected event name");
    std::vector<ExprPtr> args;
    if (match(TokenKind::LeftParen)) {
        if (cur_.kind != TokenKind::RightParen) {
            do {
                args.push_back(parseExpression());
            } while (match(TokenKind::Comma));
        }
        expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
               "expected ')'");
    }
    return std::make_unique<RaiseEventStmt>(loc, nameTok.text, std::move(args));
}

// ============================================================
// 标签 / 赋值 / 调用 (两可)
// ============================================================

StmtPtr Parser::parseLabelOrAssignmentOrCall() {
    auto loc = currentLoc();

    // 检查是否是标签: Name 后面紧跟冒号
    // 在语句起始位置, identifier: 只能是标签（VB6 规则）
    // 但在单行 If 内, colon 是语句分隔符 (If x Then a: b: c)
    if (inSingleLineIf_ == 0 && canBeName(cur_.kind) && next_.kind == TokenKind::Colon) {
        auto nameTok = advance();  // consume label name
        advance();                 // consume ':'
        return std::make_unique<LabelStmt>(loc, nameTok.text);
    }

    // VB6 行号标签: 语句起始处的整数字面量只可能是行号 —— 赋值/调用的左值必须是名字。
    // 冒号可选 (`100: x = 1` 与 `100 x = 1` 都合法)。与命名标签同一条路:
    // labelName 存行号文本, 后端发 `vb6_label_<cIdent(行号)>` 天然是合法 C 标签,
    // 语义层的标签存在性比对也是纯字符串, 所以只补这一处识别即可。
    if (inSingleLineIf_ == 0 && cur_.kind == TokenKind::IntegerLiteral) {
        auto numTok = advance();  // consume 行号
        if (cur_.kind == TokenKind::Colon) advance();  // 可选冒号
        return std::make_unique<LabelStmt>(loc, numTok.text);
    }

    // 解析左值/调用目标表达式
    // 在语句级, 顶层的 = 是赋值而非比较运算符。
    // 使用 minBp=9 (> = 的 l_bp=8) 阻止 = 被消费为比较运算符,
    // 同时允许 +,-,*,/,& 等运算符在目标内出现 (如 arr(i+1))。
    auto expr = parseExpression(9);

    // Fix 101: parseExpression 允许在无法构造表达式时返回 nullptr (例如 With 块外的
    //   '.Member'、一元 '-'/'Not' 后缺操作数、语句起始出现无法作为表达式前缀的 token)。
    //   此前该 nullptr 会一路传到本函数末尾被当作左值解引用 (expr->kind / ma.object->kind),
    //   在 Charts 2020 的 ppProgressCircular.pag 上触发 0xC0000005 解析期崩溃。
    //   这里先判空: parseExpression 已报过具体错误, 只需消费掉本行剩余 token,
    //   返回 nullptr (parseStatement 的调用方均能处理空语句) 即可安全恢复。
    if (!expr) {
        diag_.error(DiagnosticID::ParseExpectedExpression, loc,
            "无法解析语句起始的表达式 (token='" + cur_.text + "', kind=" +
            std::string(Token::kindToString(cur_.kind)) + "), 跳过该行");
        while (cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::Colon &&
               cur_.kind != TokenKind::EndOfFile && cur_.kind != TokenKind::End &&
               cur_.kind != TokenKind::Next && cur_.kind != TokenKind::Loop &&
               cur_.kind != TokenKind::Wend) {
            advance();
        }
        return nullptr;
    }

    // Fix 092r: Debug.Assert 的条件表达式可含顶层 '=' 比较:
    //   Debug.Assert (lSig And &HFF&) = (&H201 And &HFF&)
    // 上面 parseExpression(9) 已吃掉 `Debug.Assert (<cond>)`, 若这里再按赋值处理就会
    // 生成 `vb6_DebugAssert(<cond>) = <rhs>` → C2186 ("=" 左侧是 void, ToolsTlsThunks 4733).
    // 故先把 '=' 右侧并回断言条件, 再由调用路径生成 vb6_DebugAssert(<cond>)。
    if (cur_.kind == TokenKind::Equals && expr->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& dba092r = static_cast<IndexOrCallExpr&>(*expr);
        bool isDebugAssert092r = false;
        if (dba092r.named.empty() && dba092r.positional.size() == 1 && dba092r.callee
            && dba092r.callee->kind == ASTNodeKind::MemberAccessExpr) {
            auto& dbaMa092r = static_cast<MemberAccessExpr&>(*dba092r.callee);
            if (dbaMa092r.object && dbaMa092r.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& dbaObj092r = static_cast<IdentifierExpr&>(*dbaMa092r.object);
                isDebugAssert092r = toLower(dbaObj092r.name) == "debug"
                    && toLower(dbaMa092r.memberName) == "assert";
            }
        }
        if (isDebugAssert092r) {
            advance();  // consume '='
            auto rhs092r = parseExpression();
            auto cond092r = std::make_unique<BinaryExpr>(loc, BinaryOp::Eq,
                std::move(dba092r.positional[0]), std::move(rhs092r));
            dba092r.positional.clear();
            dba092r.positional.push_back(std::move(cond092r));
        }
    }

    // 检查是否是赋值
    if (match(TokenKind::Equals)) {
        auto value = parseExpression();  // 右值: = 是比较, 完整解析
        if (!value) {
            diag_.error(DiagnosticID::ParseExpectedExpression, loc,
                "赋值右值为空 (cur=" + std::string(Token::kindToString(cur_.kind)) + ")");
        }
        return std::make_unique<AssignmentStmt>(loc, std::move(expr), std::move(value));
    }

    // P15.6: 检测 Debug.Print/Debug.Assert (后续的-应为一元负号而非中缀减法)
    auto isDebugPrint = false;
    if (expr->kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<MemberAccessExpr&>(*expr);
        if (ma.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& obj = static_cast<IdentifierExpr&>(*ma.object);
            std::string objL = toLower(obj.name);
            std::string memL = toLower(ma.memberName);
            if (objL == "debug" && (memL == "print" || memL == "assert")) {
                isDebugPrint = true;
            }
        }
    }

    // Fix 110y: VB6 语句级「带括号调用 vs 无括号调用」歧义回退.
    // VB6 中在**语句**上下文里 `obj.Method (a) * b, (c)` 按**无括号调用**解析:
    // 第一个实参是完整表达式 `(a) * b`. 但表达式解析器见到 `Method (` 一律当作
    // 带括号调用, 于是语句变成 `BinOp(调用结果, b)` 后面再跟 `, (c)` → 生成
    //   (f((a)) * b)((c))
    // 这种畸形 C (Charts 2020 LabelPlus.ctl:1371
    //   UserControl.Size (lWidth + 1) * Screen.TwipsPerPixelX, (lHeight + 1) * Screen.TwipsPerPixelY
    //   → C2064 "项不会计算为接受 347 个参数的函数").
    // 这里在语句级按形状回退: 若最左叶是「obj.Member(恰好 1 个实参)」的调用,
    // 且语句在逗号后还要继续 (无括号调用的第二个实参), 则把该调用折叠回其实参,
    // 并以 obj.Member 作为真正的 callee.
    // 仅当外层确实是二元运算 (即"调用结果参与运算") 时回退, 避免影响正常的
    // `x = f(a) * b` 之外的单实参调用语句 (`obj.M (a)` 单独成句仍是合法调用).
    ExprPtr calleeOverride110y;
    ExprPtr leadingArg110y;
    bool debugPrintCollapse110y = false;
    if (expr->kind == ASTNodeKind::BinaryExpr) {
        ExprPtr* leafSlot = &expr;
        while ((*leafSlot)->kind == ASTNodeKind::BinaryExpr) {
            leafSlot = &static_cast<BinaryExpr&>(**leafSlot).left;
        }
        if ((*leafSlot)->kind == ASTNodeKind::IndexOrCallExpr) {
            auto& inner = static_cast<IndexOrCallExpr&>(**leafSlot);
            if (inner.callee && inner.callee->kind == ASTNodeKind::MemberAccessExpr
                && inner.named.empty() && inner.positional.size() == 1) {
                // Fix 110y2: 语句级 Debug.Print (expr1) & (expr2) — 首参括号被
                // 表达式解析器当成带括号调用 (Debug.Print((expr1))), 于是整句成为
                // BinOp(Concat, DebugPrint((expr1)), (expr2)) → vb6_DebugPrint
                // (void) 被当作 BSTR 拼进 ConcatFree → C2095. 与 Fix 110y 同源:
                // 最左叶是 obj.Member(1实参) 且其后要按表达式继续运算时, 把调用
                // 折叠回实参, 以 obj.Member 作为真正的 callee, 整个二元表达式作为
                // 无括号调用的首参 (VB6: Debug.Print (a & b) & (c & d) 打印
                // "((a&b)&(c&d))" 一个参数).
                auto& innerMa110y = static_cast<MemberAccessExpr&>(*inner.callee);
                bool innerLeafDebugPrint110y = false;
                if (innerMa110y.object && innerMa110y.object->kind == ASTNodeKind::IdentifierExpr) {
                    auto& innerObj110y = static_cast<IdentifierExpr&>(*innerMa110y.object);
                    innerLeafDebugPrint110y = toLower(innerObj110y.name) == "debug"
                        && toLower(innerMa110y.memberName) == "print";
                }
                if (cur_.kind == TokenKind::Comma || innerLeafDebugPrint110y) {
                    calleeOverride110y = std::move(inner.callee);
                    ExprPtr innerArg = std::move(inner.positional[0]);
                    *leafSlot = std::move(innerArg);
                    leadingArg110y = std::move(expr);
                    debugPrintCollapse110y = innerLeafDebugPrint110y;
                }
            }
        }
    }
    // Fix <vbeclipse>: 语句级 `obj.Method -30, -30, w, h` —— 无括号调用的首参是
    // 负数时, 表达式层把 `Method -30` 绑成中缀减法 Bin(Sub, MA, 30), 随后无括号
    // 调用路径把整个二元表达式当 callee → `(vb6_ComGetObjectProp(...) - 30)(...)`
    // 畸形 C (frmEditorBrowser `webBrowser.Move -30, ...` 实测 C2064 "347 个参数")。
    // VB6 语句级规则: 裸成员访问后跟 `-数字` 只能是调用语句的负数首参 ——
    // 减法结果无法被调用, 语句级不存在二义 (赋值/函数内不经过本路径)。
    // 折叠: callee = MemberAccess, 首参 = -(数字), 其余实参交还逗号列表。
    // 与 Fix 110y 同位但形状不同: 110y 的最左叶是带括号调用 IndexOrCallExpr,
    // 本条的最左叶是**不带括号**的 MemberAccessExpr。
    bool negFirstArgCollapse = false;
    // 注意: Fix 110y 折叠命中后 expr 已被 move 走 (leadingArg110y), 这里必须判空。
    if (expr && expr->kind == ASTNodeKind::BinaryExpr) {
        ExprPtr* leafSlot = &expr;
        ExprPtr* aboveMA = nullptr;  // 直接挂着 MemberAccessExpr 叶的二元节点
        while ((*leafSlot)->kind == ASTNodeKind::BinaryExpr) {
            aboveMA = leafSlot;
            leafSlot = &static_cast<BinaryExpr&>(**leafSlot).left;
        }
        if (aboveMA && (*leafSlot)->kind == ASTNodeKind::MemberAccessExpr) {
            auto& binNeg = static_cast<BinaryExpr&>(**aboveMA);
            if (binNeg.op == BinaryOp::Sub && binNeg.right
                && binNeg.right->kind == ASTNodeKind::LiteralExpr) {
                auto& litNeg = static_cast<LiteralExpr&>(*binNeg.right);
                if (litNeg.literalKind != LiteralKind::String
                    && litNeg.literalKind != LiteralKind::Date) {
                    auto negArg = std::make_unique<UnaryExpr>(binNeg.right->loc,
                        UnaryOp::Negate, std::move(binNeg.right));
                    calleeOverride110y = std::move(*leafSlot);  // MemberAccess 作 callee
                    *aboveMA = std::move(negArg);
                    leadingArg110y = std::move(expr);
                    negFirstArgCollapse = true;
                }
            }
        }
    }

    // Fix 110y2: Debug.Print 尾随中缀 (非逗号) 的折叠 — 直接构造无括号调用,
    // 折叠结果作为其唯一参数, 不会落入下方 isDebugPrint 判定 (此时 expr 已是
    // BinaryExpr, 而 isDebugPrint 要求 expr 是 MemberAccessExpr).
    if (debugPrintCollapse110y) {
        auto call = std::make_unique<IndexOrCallExpr>(loc, std::move(calleeOverride110y));
        call->positional.push_back(std::move(leadingArg110y));
        return std::make_unique<CallStmt>(loc, std::move(call));
    }

    // Fix <vbeclipse>: 负数首参折叠后, 行内没有逗号续参 (`obj.Method -30` 单独成句)
    // 就直接收口成 CallStmt; 有逗号则交给下方无括号调用参数列表 (它会把
    // leadingArg110y 作为首参再继续吃 `, arg...`)。
    if (negFirstArgCollapse
        && (cur_.kind == TokenKind::NewLine || cur_.kind == TokenKind::Colon
            || cur_.kind == TokenKind::EndOfFile)) {
        auto call = std::make_unique<IndexOrCallExpr>(loc, std::move(calleeOverride110y));
        call->positional.push_back(std::move(leadingArg110y));
        return std::make_unique<CallStmt>(loc, std::move(call));
    }

    // Fix 151: VB6 单行 If 的 Then 分支语句以 `Else` 终止:
    //   If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
    // Else/ElseIf 不得被当作裸调用的第一个实参 (VBFlexGrid.ctl 3047/3057)。
    auto isStmtEndKeyword151 = [this]() {
        return cur_.kind == TokenKind::Else || cur_.kind == TokenKind::ElseIf;
    };

    // VB6 无括号调用: Sub arg1, arg2 / Debug.Print "text"
    // 如果表达式后还有同一行的 token (非 NewLine/Colon/EndOfFile),
    // 且不是中缀运算符 (但前缀运算符如 - 可开始新参数), 则视为无括号调用的参数列表
    if (cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::Colon &&
        cur_.kind != TokenKind::EndOfFile && !isStmtEndKeyword151() &&
        (!isInfixOperator(cur_.kind) || isPrefixOperator(cur_.kind) || isDebugPrint)) {
        // 将表达式包装为 IndexOrCallExpr, 追加参数
        auto call = std::make_unique<IndexOrCallExpr>(loc,
            calleeOverride110y ? std::move(calleeOverride110y) : std::move(expr));
        // Fix 110y: 折叠出来的第一个实参 (见上方歧义回退)
        if (leadingArg110y) call->positional.push_back(std::move(leadingArg110y));

        // 解析参数列表
        // VB6 中逗号和分号都分隔参数 (分号是 Print 的位置修饰符)
        // MsgBox arg1, arg2, arg3        — 逗号分隔
        // Debug.Print "text"; i         — 分号分隔 (紧跟下一个参数)
        // Debug.Print "text";           — 末尾分号 (抑制换行, 无后续参数)

        // 辅助: 判断当前 token 是否可以开始一个表达式 (但不是语句结束符/中缀运算符)
        // 注意: 逗号后的 - 可以是一元负号 (前缀运算符), 需要允许
        auto canStartArg = [this]() -> bool {
            return cur_.kind != TokenKind::NewLine &&
                   cur_.kind != TokenKind::Colon &&
                   cur_.kind != TokenKind::EndOfFile &&
                   (!isInfixOperator(cur_.kind) || isPrefixOperator(cur_.kind)) &&
                   cur_.kind != TokenKind::Comma &&
                   cur_.kind != TokenKind::Semicolon;
        };

        // 第一个参数
        // Fix 073: 无括号调用路径也需记录 ByVal 覆盖到 byvalOverrides,
        // 与 parser_expr.cpp 带括号调用路径一致.
        // Fix 142: 索引以 call->positional.size() 为准 (含 Fix 110y 折叠出的
        // leadingArg 占用的槽位), 不再用独立计数器, 否则省略实参时槽位错位.
        if (cur_.kind == TokenKind::ByVal || cur_.kind == TokenKind::ByRef) {
            if (cur_.kind == TokenKind::ByVal) {
                call->byvalOverrides.insert(call->positional.size());
            }
            advance(); // 消费 ByVal/ByRef
        }
        if (canStartArg()) {
            // 命名参数?  name := value
            if (canBeName(cur_.kind) && next_.kind == TokenKind::Assign) {
                auto nameTok = advance(); // name
                advance(); // consume ':='
                auto val = parseExpression();
                call->named.push_back({nameTok.text, std::move(val)});
            } else {
                call->positional.push_back(parseExpression());
            }
        } else if (call->positional.empty() && cur_.kind == TokenKind::Comma) {
            // Fix 142: 省略首个位置实参 (Foo , x) — 保留槽位并记录索引
            // Fix 110y 补充: 若已折叠出 leadingArg (positional 非空), 则该逗号是
            //   实参分隔符, 由下方 while 循环的 match(Comma) 处理, 不能当成省略首参.
            call->omittedArgs.insert(call->positional.size());
            call->positional.push_back(std::make_unique<LiteralExpr>(
                currentLoc(), LiteralKind::Empty, "Empty"));
        }

        // 后续参数: 逗号或分号后继续
        while (true) {
            // 跳过分号 (; 在 Print 语句中是位置修饰符, 后面可能还有参数)
            bool sawSemi = false;
            while (cur_.kind == TokenKind::Semicolon) {
                advance();
                sawSemi = true;
            }
            // 逗号分隔 -> 继续解析下一个参数
            if (match(TokenKind::Comma)) {
                // ByVal/ByRef 前缀
                if (cur_.kind == TokenKind::ByVal || cur_.kind == TokenKind::ByRef) {
                    if (cur_.kind == TokenKind::ByVal) {
                        call->byvalOverrides.insert(call->positional.size());
                    }
                    advance();
                }
                if (canStartArg()) {
                    // 命名参数?  name := value
                    if (canBeName(cur_.kind) && next_.kind == TokenKind::Assign) {
                        auto nameTok = advance();
                        advance(); // ':='
                        auto val = parseExpression();
                        call->named.push_back({nameTok.text, std::move(val)});
                    } else {
                        call->positional.push_back(parseExpression());
                    }
                } else if (cur_.kind == TokenKind::Comma) {
                    // Fix 142: 省略的位置实参 (连续逗号) — 保留槽位并记录索引,
                    // 代码生成按形参缺省值填充, _has_ 标志置 0.
                    call->omittedArgs.insert(call->positional.size());
                    call->positional.push_back(std::make_unique<LiteralExpr>(
                        currentLoc(), LiteralKind::Empty, "Empty"));
                }
                continue;
            }
            // 分号后跟着表达式 -> 作为下一个参数
            if (sawSemi && canStartArg()) {
                call->positional.push_back(parseExpression());
                continue;
            }
            // 没有更多参数
            break;
        }

        return std::make_unique<CallStmt>(loc, std::move(call));
    }

    // 否则就是调用语句 (可能带括号也可能不带)
    return std::make_unique<CallStmt>(loc, std::move(expr));
}

} // namespace vb6c3
