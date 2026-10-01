// vb6c3 - 声明解析器
// Sub / Function / Property / Type / Enum / Declare / Event / Const / Variable

#include "parser/parser.hpp"

namespace vb6c3 {

// ============================================================
// 声明分发
// ============================================================

DeclPtr Parser::parseDeclaration() {
    // ai/vb-asm-extension-spec: `<Naked>` 只对 Sub/Function 有意义。角括号属性行在模块循环里
    // 就被 tryParseAngleAttr() 摘走并置 pendingNaked_; 若紧随的不是过程声明, 由调用方清掉。
    // 虚方法修饰位 (tB 扩展, ai/022 B08b): VB6 里 Overridable/Overrides/NotOverridable 与访问
    // 修饰符同位、互斥, 且访问修饰符可省 (`Overridable Sub X`) → 起手先吃一次, 吃了访问修饰符
    // 之后再吃一次, 两种书写顺序都收。
    ProcVirt virt = ProcVirt::None;
    eatVirtualModifiers(virt);
    if (cur_.kind != TokenKind::Public && cur_.kind != TokenKind::Private &&
        cur_.kind != TokenKind::Friend && cur_.kind != TokenKind::Protected &&
        cur_.kind != TokenKind::Global) {
        checkVirtualOnProcStart(virt);
    }

    // ai/024: `DeclareWide` (tB 兼容) —— 语法与 `Declare` 完全同形, 只是禁用
    // ANSI<->Unicode 转换。词法器把它切成 Identifier (关键词表是精确匹配,
    // "declarewide" 不命中 "declare"), 故在分派前拦一道。
    if (cur_.kind == TokenKind::Identifier && toLower(cur_.text) == "declarewide") {
        return parseDeclareDecl(AccessLevel::Default, true);
    }

    switch (cur_.kind) {
        case TokenKind::Sub: {
            auto d = parseSubDecl(AccessLevel::Default, false);
            if (d) d->virt = virt;
            return d;
        }
        case TokenKind::Function: {
            auto d = parseFunctionDecl(AccessLevel::Default, false);
            if (d) d->virt = virt;
            return d;
        }
        case TokenKind::Property: {
            auto d = parsePropertyDecl(AccessLevel::Default);
            if (d) d->virt = virt;
            return d;
        }
        case TokenKind::Type:     return parseTypeDecl(AccessLevel::Default);
        case TokenKind::Enum:     return parseEnumDecl(AccessLevel::Default);
        case TokenKind::Declare:  return parseDeclareDecl(AccessLevel::Default);
        case TokenKind::Event:    return parseEventDecl(AccessLevel::Default);
        case TokenKind::Delegate: return parseDelegateDecl(AccessLevel::Default);
        case TokenKind::Const:    return parseConstDeclList(AccessLevel::Default);
        case TokenKind::Dim:      return parseVariableDeclList(AccessLevel::Default, false);
        case TokenKind::Static:   return parseVariableDeclList(AccessLevel::Private, true);

        case TokenKind::Public:
        case TokenKind::Private:
        case TokenKind::Friend:
        case TokenKind::Protected:
        case TokenKind::Global: {
            AccessLevel access;
            if (cur_.kind == TokenKind::Public || cur_.kind == TokenKind::Global) {
                access = AccessLevel::Public;
            } else if (cur_.kind == TokenKind::Friend) {
                access = AccessLevel::Friend;
            } else if (cur_.kind == TokenKind::Protected) {
                access = AccessLevel::Protected;  // tB 扩展 (B08a)
            } else {
                access = AccessLevel::Private;
            }
            advance(); // consume access modifier

            eatVirtualModifiers(virt);         // `Public Overridable Sub` (B08b)
            checkVirtualOnProcStart(virt);     // 其余声明种类带虚修饰符 = 报错

            // Public Sub/Function/property/Type/Enum/Declare/Event/Const/Dim
            switch (cur_.kind) {
                case TokenKind::Sub: {
                    auto d = parseSubDecl(access, false);
                    if (d) d->virt = virt;
                    return d;
                }
                case TokenKind::Function: {
                    auto d = parseFunctionDecl(access, false);
                    if (d) d->virt = virt;
                    return d;
                }
                case TokenKind::Property: {
                    auto d = parsePropertyDecl(access);
                    if (d) d->virt = virt;
                    return d;
                }
                case TokenKind::Type:     return parseTypeDecl(access);
                case TokenKind::Enum:     return parseEnumDecl(access);
                case TokenKind::Declare:  return parseDeclareDecl(access);
                case TokenKind::Identifier:
                    // `Public DeclareWide Sub ...` (ai/024)
                    if (toLower(cur_.text) == "declarewide") {
                        return parseDeclareDecl(access, true);
                    }
                    return parseVariableDeclList(access, false);
                case TokenKind::Event:    return parseEventDecl(access);
                case TokenKind::Delegate: return parseDelegateDecl(access);
                case TokenKind::Const:    return parseConstDeclList(access);
                default:                  return parseVariableDeclList(access, false);
            }
        }

        default:
            diag_.error(DiagnosticID::ParseUnexpectedToken, currentLoc(),
                "expected declaration");
            advance();
            return nullptr;
    }
}

// 吃连续的虚方法修饰符 (B08b)。三个修饰符互斥: 写第二个就报一条, 值取最后一个 (让后面的
// 过程名解析继续走, 免得整块声明失联级联)。
void Parser::eatVirtualModifiers(ProcVirt& io) {
    for (;;) {
        ProcVirt v = ProcVirt::None;
        if (cur_.kind == TokenKind::Overridable) v = ProcVirt::Overridable;
        else if (cur_.kind == TokenKind::Overrides) v = ProcVirt::Overrides;
        else if (cur_.kind == TokenKind::NotOverridable) v = ProcVirt::NotOverridable;
        else break;
        if (io != ProcVirt::None) {
            diag_.error(DiagnosticID::ParseUnexpectedToken, currentLoc(),
                "duplicate virtual modifier (Overridable / Overrides / NotOverridable are"
                " mutually exclusive)");
        }
        io = v;
        advance();
    }
}

// 虚修饰符只能落在 Sub/Function/Property 上 (B08b)。报错后清空, 让声明本体照旧解析。
void Parser::checkVirtualOnProcStart(ProcVirt& io) {
    if (io == ProcVirt::None) return;
    if (cur_.kind == TokenKind::Sub || cur_.kind == TokenKind::Function ||
        cur_.kind == TokenKind::Property) return;
    diag_.error(DiagnosticID::ParseUnexpectedToken, currentLoc(),
        "virtual modifier is only allowed on Sub/Function/Property declarations");
    io = ProcVirt::None;
}

// ============================================================
// 成员级 Implements 尾子句 (tB 扩展, ai/022 D5, 批次 B02b)
// ============================================================
//
//   Private Sub OpenFile(s As String) Implements IStorage.Open, IStream.Read
//
// VB6 里"过程签名后跟 Implements"必然在 expectEndOfStatement() 处报 VB2003, 因此本
// 子句属"错误→可解析"的安全新增 (同 B01 属性行手法). 解出的 iface/member 名不在此处
// 校验语义, 由 checkNewStyleInterface 按登记表小写键解析.
//
// 点号拼接口径与模块级 parseImplements (Fix 083) 一致: `A.B.C` 的接口名取 `A.B`
// (工程/库限定), 成员名取末段.
void Parser::parseTrailingImplementsClauses(std::vector<ImplementsClause>& out) {
    if (cur_.kind != TokenKind::Implements) return;
    advance(); // 'Implements'

    // 畸形输入只报一条诊断: 留在行尾 NewLine 上让调用方的 expectEndOfStatement 正常通过.
    auto skipRestOfLine = [&]() {
        while (cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::EndOfFile) {
            advance();
        }
    };

    for (;;) {
        if (!canBeName(cur_.kind)) {
            diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                "expected interface member name after 'Implements'");
            skipRestOfLine();
            return;
        }
        ImplementsClause c;
        c.loc = currentLoc();
        std::vector<std::string> parts;
        parts.push_back(advance().text);
        while (cur_.kind == TokenKind::Dot && canBeName(peek2().kind)) {
            advance(); // '.'
            parts.push_back(advance().text);
        }
        if (parts.size() < 2) {
            diag_.error(DiagnosticID::ParseInvalidInterfaceMember, c.loc,
                "Member-level Implements needs a qualified name: Implements <Interface>.<Member> (got " +
                parts.front() + ")");
        } else {
            for (size_t i = 0; i + 1 < parts.size(); i++) {
                if (i) c.ifaceName += ".";
                c.ifaceName += parts[i];
            }
            c.memberName = parts.back();
            out.push_back(std::move(c));   // 畸形子句不入列表: 语义层不再二次报错
        }
        if (cur_.kind != TokenKind::Comma) break;
        advance(); // ','
    }
}

// 角括号过程属性 (ai/vb-asm-extension-spec): 目前只有 `<Naked>` 一个名字。
// 直接拿原始行文本判定 (token 级前瞻要跨两个 token, 而这里要认的只是"整行是不是 <Naked>"),
// 命中就吃掉整行 token 并把 pendingNaked_ 置位, 由紧随的 Sub/Function 声明取走。
bool Parser::tryParseAngleAttr() {
    if (cur_.kind != TokenKind::LessThan) return false;
    std::string text;
    if (buffer_) {
        std::string line(buffer_->getLine(cur_.line));
        while (!line.empty() && (line.back() == '\n' || line.back() == '\r')) line.pop_back();
        size_t b = line.find_first_not_of(" \t");
        size_t e = line.find_last_not_of(" \t");
        if (b != std::string::npos) text = line.substr(b, e - b + 1);
    }
    std::string compact;                       // 去空白 + 转小写, 只用来判定
    for (char c : text) {
        if (c == ' ' || c == '\t' || c == '\r' || c == '\n') continue;
        compact += static_cast<char>(::tolower(static_cast<unsigned char>(c)));
    }
    if (compact == "<naked>") {
        while (!check(TokenKind::NewLine) && !check(TokenKind::EndOfFile)) advance();
        pendingNaked_ = true;
        return true;
    }
    if (compact.size() >= 2 && compact.front() == '<' && compact.back() == '>') {
        diag_.error(DiagnosticID::ParseUnknownAttribute, currentLoc(),
                    "未知的角括号属性 (目前仅支持 <Naked>): " + text);
        while (!check(TokenKind::NewLine) && !check(TokenKind::EndOfFile)) advance();
        return true;
    }
    return false;   // 不是属性行 (比较运算符等), token 流原样交回
}

// ============================================================
// Sub 声明
// ============================================================

std::unique_ptr<SubDecl> Parser::parseSubDecl(AccessLevel access, bool isStatic) {
    auto loc = currentLoc();
    advance(); // consume 'Sub'

    auto nameTok = expectName("expected Sub name");
    auto typeParams = parseTypeParams();   // 泛型 (tB): Sub Foo(Of T)
    auto params = parseParameterList();
    std::vector<ImplementsClause> clauses;
    parseTrailingImplementsClauses(clauses);   // B02b
    expectEndOfStatement();

    auto body = parseBlockUntil({TokenKind::End});

    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Sub'");
    // Fix <vbeclipse>: 真实世界工程存在块终结符不匹配 (VbEclipse SchemeWinXP.cls:
    // Property Get 用 End Function 收尾)。宽容消费: 只要 End 后是名字型 token
    // 就当终结符吃掉并告警, 不让整个类模块解析脱轨。
    if (!match(TokenKind::Sub)) {
        if (cur_.kind == TokenKind::Sub || cur_.kind == TokenKind::Function
            || cur_.kind == TokenKind::Property) {
            diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, currentLoc(),
                std::string("块终结符不匹配: 'End ") + cur_.text + "' (期望 'End Sub')");
            advance();
        } else {
            diag_.error(DiagnosticID::ParseMismatchedBlock, currentLoc(),
                "expected 'End Sub'");
        }
    }

    auto d = std::make_unique<SubDecl>(loc, access, nameTok.text,
        std::move(params), std::move(body), isStatic);
    d->typeParams = std::move(typeParams);
    d->implementsClauses = std::move(clauses);   // B02b
    d->isNaked = pendingNaked_;                  // ai/vb-asm-extension-spec: <Naked>
    pendingNaked_ = false;
    curTypeParams_.clear();  // G3 护栏窗口只覆盖本模板 params+body
    return d;
}

// ============================================================
// Function 声明
// ============================================================

std::unique_ptr<FunctionDecl> Parser::parseFunctionDecl(AccessLevel access, bool isStatic) {
    auto loc = currentLoc();
    advance(); // consume 'Function'

    auto nameTok = expectName("expected Function name");
    auto typeParams = parseTypeParams();   // 泛型 (tB): Function Foo(Of T)
    auto params = parseParameterList();

    TypeRefPtr returnType;
    if (match(TokenKind::As)) {
        returnType = parseTypeRef();
    }
    std::vector<ImplementsClause> clauses;
    parseTrailingImplementsClauses(clauses);   // B02b: 子句写在 `As Type` 之后
    expectEndOfStatement();

    auto body = parseBlockUntil({TokenKind::End});

    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Function'");
    // Fix <vbeclipse>: 真实世界工程存在块终结符不匹配 (VbEclipse SchemeWinXP.cls:
    // Property Get 用 End Function 收尾)。宽容消费: 只要 End 后是名字型 token
    // 就当终结符吃掉并告警, 不让整个类模块解析脱轨。
    if (!match(TokenKind::Function)) {
        if (cur_.kind == TokenKind::Sub || cur_.kind == TokenKind::Function
            || cur_.kind == TokenKind::Property) {
            diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, currentLoc(),
                std::string("块终结符不匹配: 'End ") + cur_.text + "' (期望 'End Function')");
            advance();
        } else {
            diag_.error(DiagnosticID::ParseMismatchedBlock, currentLoc(),
                "expected 'End Function'");
        }
    }

    auto d = std::make_unique<FunctionDecl>(loc, access, nameTok.text,
        std::move(params), std::move(returnType), std::move(body), isStatic);
    d->typeParams = std::move(typeParams);
    d->implementsClauses = std::move(clauses);   // B02b
    d->isNaked = pendingNaked_;                  // ai/vb-asm-extension-spec: <Naked>
    pendingNaked_ = false;
    curTypeParams_.clear();  // G3 护栏窗口只覆盖本模板 params+As+body
    return d;
}

// ============================================================
// Property 声明
// ============================================================

std::unique_ptr<PropertyDecl> Parser::parsePropertyDecl(AccessLevel access) {
    auto loc = currentLoc();
    advance(); // consume 'Property'

    ProcKind propKind;
    if (match(TokenKind::Get)) {
        propKind = ProcKind::PropertyGet;
    } else if (match(TokenKind::Let)) {
        propKind = ProcKind::PropertyLet;
    } else if (match(TokenKind::Set)) {
        propKind = ProcKind::PropertySet;
    } else {
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            "expected 'Get', 'Let', or 'Set' after 'Property'");
        propKind = ProcKind::PropertyGet;
    }

    auto nameTok = expectName("expected Property name");
    auto typeParams = parseTypeParams();   // 泛型 (tB): Property Get Foo(Of T)
    auto params = parseParameterList();

    TypeRefPtr returnType;
    if (match(TokenKind::As)) {
        returnType = parseTypeRef();
    }
    std::vector<ImplementsClause> clauses;
    parseTrailingImplementsClauses(clauses);   // B02b: 子句写在 `As Type` 之后
    expectEndOfStatement();

    auto body = parseBlockUntil({TokenKind::End});

    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Property'");
    // Fix <vbeclipse>: 真实世界工程存在块终结符不匹配 (VbEclipse SchemeWinXP.cls:
    // Property Get 用 End Function 收尾)。宽容消费: 只要 End 后是名字型 token
    // 就当终结符吃掉并告警, 不让整个类模块解析脱轨。
    if (!match(TokenKind::Property)) {
        if (cur_.kind == TokenKind::Sub || cur_.kind == TokenKind::Function
            || cur_.kind == TokenKind::Property) {
            diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, currentLoc(),
                std::string("块终结符不匹配: 'End ") + cur_.text + "' (期望 'End Property')");
            advance();
        } else {
            diag_.error(DiagnosticID::ParseMismatchedBlock, currentLoc(),
                "expected 'End Property'");
        }
    }

    auto d = std::make_unique<PropertyDecl>(loc, access, propKind,
        nameTok.text, std::move(params), std::move(returnType), std::move(body));
    d->typeParams = std::move(typeParams);
    d->implementsClauses = std::move(clauses);   // B02b
    curTypeParams_.clear();  // G3 护栏窗口只覆盖本模板 params+As+body
    return d;
}

// ============================================================
// Type 声明 (用户自定义类型/UDT)
// ============================================================

std::unique_ptr<TypeDecl> Parser::parseTypeDecl(AccessLevel access) {
    auto loc = currentLoc();
    advance(); // consume 'Type'
    auto nameTok = expectName("expected Type name");
    auto typeParams = parseTypeParams();   // 泛型 (tB): Type Foo(Of T)
    expectEndOfStatement();

    std::vector<std::unique_ptr<TypeMember>> members;
    while (cur_.kind != TokenKind::End && cur_.kind != TokenKind::EndOfFile) {
        skipNewLines();
        if (cur_.kind == TokenKind::End) break;

        auto memberLoc = currentLoc();
        // 允许硬关键字作为 Type 成员名 (如 Next, Type 等)
        std::string memberNameStr;
        if (canBeName(cur_.kind)) {
            memberNameStr = advance().text;
        } else if (!cur_.text.empty() && cur_.kind != TokenKind::EndOfFile &&
                   cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::Colon &&
                   cur_.kind != TokenKind::LeftParen && cur_.kind != TokenKind::RightParen &&
                   cur_.kind != TokenKind::Comma && cur_.kind != TokenKind::End) {
            memberNameStr = advance().text;
        } else {
            diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                std::string("expected member name (got ") +
                Token::kindToString(cur_.kind) + ")");
            memberNameStr = "?";
        }

        // 可能有数组维度:
        //   memberName(10) As Type    — 固定大小
        //   memberName() As Type      — 动态数组
        //   memberName(1 To 8) As Type — 下界 To 上界
        //   memberName(1, 2) As Type  — 多维
        ExprPtr arraySize;
        bool isArrayDynamic = false;  // Fix 037: 标记动态数组 `()` 语法
        if (match(TokenKind::LeftParen)) {
            if (cur_.kind != TokenKind::RightParen) {
                // 解析第一个维度
                auto first = parseExpression();
                if (match(TokenKind::To)) {
                    // 1 To 8: 只保留上界 (语法检查阶段)
                    auto upper = parseExpression();
                    arraySize = std::move(upper);
                } else {
                    arraySize = std::move(first);
                }
                // 消费后续维度 (多维数组)
                while (match(TokenKind::Comma)) {
                    parseExpression(); // 解析并丢弃后续维度
                    if (match(TokenKind::To)) {
                        parseExpression();
                    }
                }
            } else {
                // 空括号 () = 动态数组, arraySize 保持 nullptr
                isArrayDynamic = true;
            }
            expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
                   "expected ')'");
        }

        TypeRefPtr type;
        if (match(TokenKind::As)) {
            type = parseTypeRef();
        }
        expectEndOfStatement();

        auto memberNode = std::make_unique<TypeMember>(memberLoc,
            memberNameStr, std::move(type), std::move(arraySize));
        memberNode->isArrayDynamic = isArrayDynamic;
        members.push_back(std::move(memberNode));
    }

    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Type'");
    expect(TokenKind::Type, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Type'");

    auto d = std::make_unique<TypeDecl>(loc, access, nameTok.text, std::move(members));
    d->typeParams = std::move(typeParams);
    curTypeParams_.clear();  // G3 护栏窗口只覆盖本模板成员类型
    return d;
}

// ============================================================
// Enum 声明
// ============================================================

std::unique_ptr<EnumDecl> Parser::parseEnumDecl(AccessLevel access) {
    auto loc = currentLoc();
    advance(); // consume 'Enum'
    auto nameTok = expectName("expected Enum name");
    expectEndOfStatement();

    std::vector<std::unique_ptr<EnumMember>> members;
    while (cur_.kind != TokenKind::End && cur_.kind != TokenKind::EndOfFile) {
        skipNewLines();
        if (cur_.kind == TokenKind::End) break;

        auto memberLoc = currentLoc();
        auto memberName = expectName("expected enum member name");

        ExprPtr value;
        if (match(TokenKind::Equals)) {
            value = parseExpression();
        }
        expectEndOfStatement();

        members.push_back(std::make_unique<EnumMember>(memberLoc,
            memberName.text, std::move(value)));
    }

    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Enum'");
    expect(TokenKind::Enum, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Enum'");

    return std::make_unique<EnumDecl>(loc, access, nameTok.text, std::move(members));
}

} // namespace vb6c3
