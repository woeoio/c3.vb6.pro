#include "parser/parser.hpp"
#include <cstdlib>
#include <cctype>
#include <string>
#include <vector>
#include <memory>

namespace vb6c3 {

// ============================================================
// 模块解析 (最高层)
// ============================================================

std::unique_ptr<Module> Parser::parseModule(bool isClassModule) {
    auto mod = std::make_unique<Module>(currentLoc(), buffer_->filePath());
    mod->isClassModule = isClassModule;
    parseModuleBody(*mod);
    return mod;
}

void Parser::parseModuleBody(Module& mod) {
    skipNewLines();

    // 类模块头部: 跳过 VERSION 1.0 CLASS ... BEGIN ... END 块
    // VB6 .cls 文件格式:
    //   VERSION 1.0 CLASS
    //   BEGIN
    //     MultiUse = -1  'True
    //     ...
    //   END
    if (mod.isClassModule) {
        // 检查是否有 VERSION 标记
        if (cur_.kind == TokenKind::Identifier &&
            toLower(cur_.text) == "version") {
            // 跳过 VERSION 1.0 CLASS 行
            skipToNextLine();
            skipNewLines();

            // 解析 BEGIN ... END 块, 提取 MultiUse 等 instancing 属性
            if (cur_.kind == TokenKind::Begin) {
                advance();  // consume BEGIN
                int depth = 1;
                while (cur_.kind != TokenKind::EndOfFile && depth > 0) {
                    if (cur_.kind == TokenKind::Begin) {
                        depth++;
                    } else if (cur_.kind == TokenKind::End) {
                        depth--;
                        if (depth == 0) {
                            advance();  // consume END
                            break;
                        }
                    } else if (depth == 1 && cur_.kind == TokenKind::Identifier) {
                        // 顶层属性: 检查是否为 MultiUse
                        std::string attrName = toLower(cur_.text);
                        if (attrName == "multiuse") {
                            advance();  // consume MultiUse
                            if (cur_.kind == TokenKind::Equals) {
                                advance();  // consume =
                                // -1 = True (MultiUse), 0 = False (Private)
                                if (cur_.kind == TokenKind::Minus) {
                                    advance();  // consume -
                                }
                                if (cur_.kind == TokenKind::IntegerLiteral ||
                                    cur_.kind == TokenKind::LongLiteral) {
                                    int val = std::atoi(cur_.text.c_str());
                                    if (val != 0) {
                                        mod.instancing = VBInstancing::MultiUse;
                                    }
                                    advance();
                                }
                            }
                            continue;  // 已消费属性, 不再 advance
                        }
                    }
                    advance();
                }
            }
            skipNewLines();
        }
    }

    while (cur_.kind != TokenKind::EndOfFile) {
        skipNewLines();
        if (cur_.kind == TokenKind::EndOfFile) break;

        // 泛型类 (tB 扩展, G4): 类模块可选头行 `Class Name(Of T[,U])`.
        // VB6 语言不存在语句位置的 Class 关键字 → 存量 .cls 零误伤;
        // 类型参数经 outerTypeParams_ 贯穿整个模块体 (护栏 + 特化克隆).
        if (mod.isClassModule && cur_.kind == TokenKind::Class &&
            peek2().kind == TokenKind::Identifier && mod.classTypeParams.empty()) {
            advance(); // 'Class'
            auto nameTok = advance(); // 类名
            auto tp = parseTypeParams();
            if (tp.empty()) {
                diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                    "泛型类头行需要 (Of T[,U]) 类型参数表: Class " + nameTok.text);
            }
            mod.classTypeParams = tp;
            if (mod.moduleName.empty()) mod.moduleName = nameTok.text;
            outerTypeParams_ = curTypeParams_;  // 提升到模块层
            classHeaderSeen_ = true;
            expectEndOfStatement();
            continue;
        }
        if (classHeaderSeen_ && mod.isClassModule &&
            cur_.kind == TokenKind::End && peek2().kind == TokenKind::Class) {
            advance(); advance();  // 'End Class'
            expectEndOfStatement();
            continue;
        }

        // Option 语句 (必须在最前面)
        if (cur_.kind == TokenKind::Option) {
            mod.options.push_back(parseOption());
            expectEndOfStatement();
            continue;
        }

        // Implements 语句
        if (cur_.kind == TokenKind::Implements) {
            mod.implements.push_back(parseImplements());
            expectEndOfStatement();
            continue;
        }

        // Inherits 语句 (tB 扩展, ai/022 D6, 批次 B07): 类模块的 `Inherits Base`.
        // 这类行过去必然落进下方的 "unexpected token at module level" 错误分支 (VB2002),
        // 因此新增分支只把"错误"变成"可解析", 存量工程逐字节不变.
        if (cur_.kind == TokenKind::Inherits) {
            mod.inherits.push_back(parseInherits());
            expectEndOfStatement();
            continue;
        }

        // Interface 契约块 (tB 扩展, ai/022 D1): 可带前置 [属性行].
        // 这类行过去必然落进下方的 "unexpected token at module level" 错误分支,
        // 因此新增分支只把"错误"变成"可解析", 存量工程逐字节不变.
        {
            bool sawAttr = false;
            std::vector<InterfaceAttr> pendingAttrs;
            while (atBracketAttrLine()) {
                sawAttr = true;
                InterfaceAttr attr;
                if (parseBracketAttrLine(attr)) pendingAttrs.push_back(std::move(attr));
                skipNewLines();
            }
            if (cur_.kind == TokenKind::Interface) {
                mod.interfaces.push_back(parseInterfaceDecl(pendingAttrs));
                expectEndOfStatement();
                continue;
            }
            // CoClass 契约聚合块 (tB 扩展, ai/026 四节 / ai/022 D44, 批次 B11/C01):
            // 同样只在"过去必然报错"的位置新增分支, 存量工程逐字节不变.
            if (cur_.kind == TokenKind::CoClass) {
                mod.coclasses.push_back(parseCoClassDecl(pendingAttrs));
                expectEndOfStatement();
                continue;
            }
            if (sawAttr) {
                if (!pendingAttrs.empty()) {
                    diag_.error(DiagnosticID::ParseUnexpectedToken, currentLoc(),
                        "Attribute line must precede an Interface or CoClass declaration");
                    skipToNextLine();
                }
                continue;  // 属性行本身已报错并越过该行
            }
        }

        // DefType 语句
        if (checkAny({TokenKind::DefBool, TokenKind::DefByte, TokenKind::DefInt,
                      TokenKind::DefLng, TokenKind::DefCur, TokenKind::DefSng,
                      TokenKind::DefDbl, TokenKind::DefDate, TokenKind::DefStr,
                      TokenKind::DefObj, TokenKind::DefVar})) {
            mod.defTypes.push_back(parseDefType());
            expectEndOfStatement();
            continue;
        }

        // 角括号过程属性行 (ai/vb-asm-extension-spec): `<Naked>`
        if (cur_.kind == TokenKind::LessThan && tryParseAngleAttr()) {
            expectEndOfStatement();
            continue;
        }

        // Attribute 语句
        if (cur_.kind == TokenKind::Attribute) {
            mod.attributes.push_back(parseAttribute());
            expectEndOfStatement();
            continue;
        }

        // 声明
        if (isDeclarationStart()) {
            auto decl = parseDeclaration();
            if (pendingNaked_) {   // `<Naked>` 后面跟的不是过程声明
                diag_.error(DiagnosticID::ParseUnknownAttribute, currentLoc(),
                            "<Naked> 只能修饰 Sub/Function");
                pendingNaked_ = false;
            }
            if (decl) {
                // 逗号分隔的多变量声明展开为独立声明
                if (decl->kind == ASTNodeKind::MultiDecl) {
                    auto* multi = static_cast<MultiDecl*>(decl.get());
                    for (auto& d : multi->declarations) {
                        mod.declarations.push_back(std::move(d));
                    }
                } else {
                    mod.declarations.push_back(std::move(decl));
                }
            }
            continue;
        }

        // 不应该出现在模块级的语句 → 报错并跳过
        diag_.error(DiagnosticID::ParseUnexpectedToken, currentLoc(),
            "unexpected token at module level: " + cur_.text);
        advance();
        skipToNextLine();
    }
}

std::unique_ptr<OptionStmt> Parser::parseOption() {
    auto loc = currentLoc();
    advance(); // consume 'Option'

    if (match(TokenKind::Explicit)) {
        return std::make_unique<OptionStmt>(loc, OptionKind::Explicit);
    }
    if (match(TokenKind::Compare)) {
        if (match(TokenKind::Text)) {
            return std::make_unique<OptionStmt>(loc, OptionKind::CompareText);
        }
        if (match(TokenKind::Binary2)) {
            return std::make_unique<OptionStmt>(loc, OptionKind::CompareBinary);
        }
        // Option Compare 无后续 → 报错
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            "expected 'Text' or 'Binary' after 'Option Compare'");
        return std::make_unique<OptionStmt>(loc, OptionKind::CompareBinary);
    }
    if (match(TokenKind::Base)) {
        if (match(TokenKind::IntegerLiteral)) {
            // 简化: 0 或 1
            return std::make_unique<OptionStmt>(loc, OptionKind::BaseZero);
        }
        // 检查是否是 1
        if (cur_.kind == TokenKind::IntegerLiteral && cur_.intValue == 1) {
            advance();
            return std::make_unique<OptionStmt>(loc, OptionKind::BaseOne);
        }
        return std::make_unique<OptionStmt>(loc, OptionKind::BaseZero);
    }
    if (match(TokenKind::Private2)) {
        if (match(TokenKind::Class)) {
            // Option Private Module — VB6 特有
            return std::make_unique<OptionStmt>(loc, OptionKind::PrivateModule);
        }
        // 不完整
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            "expected 'Module' after 'Option Private'");
        return std::make_unique<OptionStmt>(loc, OptionKind::PrivateModule);
    }

    diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
        "expected 'Explicit', 'Compare', 'Base', or 'Private' after 'Option'");
    return std::make_unique<OptionStmt>(loc, OptionKind::Explicit);
}

std::unique_ptr<ImplementsStmt> Parser::parseImplements() {
    auto loc = currentLoc();
    advance(); // consume 'Implements'
    auto nameTok = expectName("expected interface name");
    // Fix 083: VB6 允许点号限定的库/工程名 — `Implements OLEGuids.IObjectSafety`、
    // `Implements Project.IfaceName`。样例工程 VBFlexGrid.ctl(1896-1898) 三处均为
    // 此形式: 原先只吃一个标识符, expectEndOfStatement() 撞到 '.' → VB2003, 接口名
    // 再落到模块级默认分支 → VB2002 "unexpected token at module level"。
    // 用 canBeName(peek2()) 而非裸 while(Dot), 避免 `A..B` 这类畸形输入无限追加。
    // 裸名路径完全不变 (无 '.' 时不进循环)。严格超集。
    std::string fullName = nameTok.text;
    while (cur_.kind == TokenKind::Dot && canBeName(peek2().kind)) {
        advance(); // consume '.'
        auto part = expectName("expected interface name after '.'");
        fullName += "." + part.text;
    }
    auto stmt = std::make_unique<ImplementsStmt>(loc, fullName);
    // tB 扩展 (ai/022 D42, 批次 B10): 委托式实现子句 `Implements I Via m_holder`。
    // 没有 Via 时下面的分支根本不进, 存量路径逐字节不变 (VB6 里 `Via` 不是关键字,
    // 语料核查: tests/archive/publish 的 .bas/.cls/.frm/.ctl 中 \bvia\b 零命中)。
    if (cur_.kind == TokenKind::Via) {
        advance(); // consume 'Via'
        stmt->viaField = expectName("expected holder field name after 'Via'").text;
    }
    return stmt;
}

// 类继承子句 (tB 扩展, B07): 与 parseImplements 同口径吃点号限定名 (Project.IFace 那种
// 写法), 但 v1 只有一条基类名 —— 逗号列表在这里天然落到 expectEndOfStatement 的 VB2003,
// 即"单继承"的 arity 检查不需要新代码.
InheritsStmt Parser::parseInherits() {
    InheritsStmt s;
    s.loc = currentLoc();
    advance(); // consume 'Inherits'
    auto nameTok = expectName("expected base class name");
    s.baseName = nameTok.text;
    while (cur_.kind == TokenKind::Dot && canBeName(peek2().kind)) {
        advance(); // consume '.'
        auto part = expectName("expected base class name after '.'");
        s.baseName += "." + part.text;
    }
    return s;
}

std::unique_ptr<DefTypeStmt> Parser::parseDefType() {
    auto loc = currentLoc();
    DefTypeKind dk;
    switch (cur_.kind) {
        case TokenKind::DefBool: dk = DefTypeKind::Bool; break;
        case TokenKind::DefByte: dk = DefTypeKind::Byte; break;
        case TokenKind::DefInt:  dk = DefTypeKind::Int;  break;
        case TokenKind::DefLng:  dk = DefTypeKind::Lng;  break;
        case TokenKind::DefCur:  dk = DefTypeKind::Cur;  break;
        case TokenKind::DefSng:  dk = DefTypeKind::Sng;  break;
        case TokenKind::DefDbl:  dk = DefTypeKind::Dbl;  break;
        case TokenKind::DefDate: dk = DefTypeKind::Date; break;
        case TokenKind::DefStr:  dk = DefTypeKind::Str;  break;
        case TokenKind::DefObj:  dk = DefTypeKind::Obj;  break;
        case TokenKind::DefVar:  dk = DefTypeKind::Var;  break;
        default:
            diag_.error(DiagnosticID::ParseUnexpectedToken, currentLoc(),
                "expected DefType keyword");
            return nullptr;
    }
    advance(); // consume DefInt/DefStr/...

    // 解析字母范围: A-C, D, E-G
    std::vector<DefTypeStmt::LetterRange> ranges;
    do {
        auto letterTok = expectName("expected letter range");
        char from = letterTok.text.empty() ? 'A' : static_cast<char>(std::toupper(letterTok.text[0]));
        char to = from;
        if (match(TokenKind::Minus)) {
            auto endTok = expectName("expected letter");
            to = endTok.text.empty() ? 'Z' : static_cast<char>(std::toupper(endTok.text[0]));
        }
        ranges.push_back({from, to});
    } while (match(TokenKind::Comma));

    return std::make_unique<DefTypeStmt>(loc, dk, std::move(ranges));
}

std::unique_ptr<AttributeStmt> Parser::parseAttribute() {
    auto loc = currentLoc();
    advance(); // consume 'Attribute'
    // Attribute VB_Name = "ModuleName"
    auto nameTok = expectName("expected attribute name");
    std::string attrName = nameTok.text;

    // 可能有 . (如 Attribute VB_Name)
    while (match(TokenKind::Dot)) {
        auto part = expectName("expected attribute name part");
        attrName += "." + part.text;
    }

    expect(TokenKind::Equals, DiagnosticID::ParseExpectedToken,
           "expected '=' in attribute");

    auto value = parseExpression();
    // Fix <vbeclipse>: Attribute 取值可带逗号分隔的多个字面量
    // (ucPerspective.ctl: Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes")。
    // 这些附加值无语义消费方, 解析后丢弃, 不让逗号漏成 "expected end of statement"。
    while (cur_.kind == TokenKind::Comma) {
        advance();
        if (!check(TokenKind::NewLine) && !check(TokenKind::EndOfFile)) {
            parseExpression();
        }
    }
    return std::make_unique<AttributeStmt>(loc, attrName, std::move(value));
}

} // namespace vb6c3
