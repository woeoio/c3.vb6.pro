// vb6c3 - Interface / CoClass 契约块解析 (tB 扩展; 设计依据 ai/022 D1/D44 与 ai/026 四节)
//
// 语法:
//   [InterfaceId("{GUID}")] [Description("...")]      <- 属性行, 可多行
//   Interface Name [Extends ParentInterface]
//       Sub M(args)
//       Function F(args) As Type
//       Property Get|Let|Set P(args) [As Type]
//   End Interface
//
//   [CoClassId("{GUID}")] [ProgId("...")] [ComCreatable(True)] [Implementation("Cls")]
//   CoClass Name
//       [Default] Interface IfaceName        <- 引用已声明的接口, 不是内联定义
//   End CoClass
//
// 纯契约: 成员无实现体、无可见性关键字, 块内无字段/常量/枚举/事件/Declare.
//
// 属性行为什么在 parser 侧拆解而不改 lexer: 词法器把 `[`..`]` 整体扫成一个 Identifier
// token (lexer.cpp), 因为 VB6 的 `[带空格的名字]` 名称引用依赖它. 这类行过去在模块级
// 必然报 VB2002 "unexpected token at module level", 所以"错误→可解析"对存量工程零影响.

#include "parser/parser.hpp"

#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <set>
#include <string>

namespace vb6c3 {
namespace {

std::string itfAsciiLower(std::string s) {
    std::transform(s.begin(), s.end(), s.begin(),
                   [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
    return s;
}

std::string itfTrim(const std::string& s) {
    size_t b = s.find_first_not_of(" \t");
    if (b == std::string::npos) return std::string();
    size_t e = s.find_last_not_of(" \t");
    return s.substr(b, e - b + 1);
}

// 已登记的属性名 (大小写不敏感). CoClass 系列先接受语法, 由 P5/P6 批次消费与落点校验.
const std::set<std::string>& itfKnownAttrNames() {
    static const std::set<std::string> kNames = {
        "interfaceid",    "description",  "hidden",       "restricted",
        "oleautomation",  "comimport",    "comextensible", "preservesig",
        "dispid",         "default",      "source",
        "coclassid",      "comcreatable", "coclasscustomconstructor",
        "progid",         "implementation",  // ai/026 四节样本用到 (ai/022 D44 第 2 条)
        "packingalignment", // tB UDT packing attribute (Type-level)
    };
    return kNames;
}

} // namespace

// ============================================================
// 方括号属性行
// ============================================================

bool Parser::atBracketAttrLine() const {
    if (cur_.kind != TokenKind::Identifier) return false;
    if (cur_.text.size() < 3) return false;
    return cur_.text.front() == '[' && cur_.text.back() == ']';
}

// 成功: 填充 out 并消费该行剩余部分, 返回 true.
// 失败: 自行报错 + skipToNextLine 恢复, 返回 false (调用方不得再对该行重复报错).
// requireOwnLine=false 时不要求属性行独占一行: CoClass 块的 `[Default] Interface X` 写法
// (ai/026 四节样本) 让声明紧跟在同一行. Interface 那条路仍按原样要求独占 (D44 第 4 条:
// 不动它才保住 itf_n06 那条负例的既有诊断).
bool Parser::parseBracketAttrLine(InterfaceAttr& out, bool requireOwnLine) {
    out = InterfaceAttr();
    out.loc = currentLoc();
    std::string text = advance().text;  // "[Name]" / "[Name(arg)]"
    std::string inner = text.substr(1, text.size() - 2);

    auto fail = [&](const std::string& msg) {
        diag_.error(DiagnosticID::ParseUnknownAttribute, out.loc, msg);
        skipToNextLine();
        return false;
    };

    size_t lp = inner.find('(');
    out.name = itfTrim(lp == std::string::npos ? inner : inner.substr(0, lp));
    if (out.name.empty()) return fail("Attribute line has no attribute name: " + text);
    if (itfKnownAttrNames().count(itfAsciiLower(out.name)) == 0)
        return fail("Unrecognized attribute line [" + out.name + "]");

    if (lp != std::string::npos) {
        size_t rp = inner.rfind(')');
        if (rp == std::string::npos || rp < lp)
            return fail("Attribute line argument parentheses are unbalanced: " + text);
        std::string arg = itfTrim(inner.substr(lp + 1, rp - lp - 1));
        if (arg.size() >= 2 && arg.front() == '"' && arg.back() == '"') {
            std::string body = arg.substr(1, arg.size() - 2);
            for (size_t i = 0; i < body.size(); i++) {
                if (body[i] == '"' && i + 1 < body.size() && body[i + 1] == '"') i++;
                out.strValue += body[i];
            }
            out.hasStr = true;
        } else {
            char* end = nullptr;
            long long v = std::strtoll(arg.c_str(), &end, 10);
            if (end == arg.c_str() || (end && *end != '\0')) {
                // 布尔字面量: `[ComCreatable(True)]`. 整行是一个 Identifier token,
                // True/False 在这里只是文本 (D44 第 3 条), 所以在这层认.
                // 折成 numValue 1/0 而不是新增字段: C02 求解身份时 True 与 1 同义.
                std::string lower = itfAsciiLower(arg);
                if (lower == "true" || lower == "false") {
                    out.numValue = (lower == "true") ? 1 : 0;
                    out.hasNum = true;
                } else {
                    return fail("Attribute line argument must be a string, an integer or True/False: " + arg);
                }
            } else {
                out.numValue = v;
                out.hasNum = true;
            }
        }
    }
    if (requireOwnLine) expectEndOfStatement();
    return true;
}

// ============================================================
// 接口成员签名 (复用 Sub/Function/Property 节点, body 恒空)
// ============================================================

DeclPtr Parser::parseInterfaceMemberDecl() {
    auto loc = currentLoc();
    auto genericGuard = [&](const std::string& procName) {
        auto typeParams = parseTypeParams();
        if (!typeParams.empty()) {
            diag_.error(DiagnosticID::ParseInvalidInterfaceMember, loc,
                "Interface member does not support generic type parameters (Of T): " + procName);
        }
    };

    switch (cur_.kind) {
        case TokenKind::Sub: {
            advance(); // 'Sub'
            auto nameTok = expectName("expected Sub name in Interface");
            genericGuard(nameTok.text);
            auto params = parseParameterList();
            expectEndOfStatement();
            curTypeParams_.clear();
            return std::make_unique<SubDecl>(loc, AccessLevel::Public, nameTok.text,
                                             std::move(params), StmtList(), false);
        }
        case TokenKind::Function: {
            advance(); // 'Function'
            auto nameTok = expectName("expected Function name in Interface");
            genericGuard(nameTok.text);
            auto params = parseParameterList();
            TypeRefPtr returnType;
            if (match(TokenKind::As)) {
                returnType = parseTypeRef();
            } else {
                diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                    "expected 'As type' in Interface Function " + nameTok.text);
            }
            expectEndOfStatement();
            curTypeParams_.clear();
            return std::make_unique<FunctionDecl>(loc, AccessLevel::Public, nameTok.text,
                                                  std::move(params), std::move(returnType),
                                                  StmtList(), false);
        }
        case TokenKind::Property: {
            advance(); // 'Property'
            ProcKind propKind = ProcKind::PropertyGet;
            if (match(TokenKind::Get)) {
                propKind = ProcKind::PropertyGet;
            } else if (match(TokenKind::Let)) {
                propKind = ProcKind::PropertyLet;
            } else if (match(TokenKind::Set)) {
                propKind = ProcKind::PropertySet;
            } else {
                diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                    "expected 'Get', 'Let', or 'Set' after 'Property'");
            }
            auto nameTok = expectName("expected Property name in Interface");
            genericGuard(nameTok.text);
            auto params = parseParameterList();
            TypeRefPtr returnType;
            if (match(TokenKind::As)) {
                returnType = parseTypeRef();
            }
            expectEndOfStatement();
            curTypeParams_.clear();
            return std::make_unique<PropertyDecl>(loc, AccessLevel::Public, propKind,
                                                  nameTok.text, std::move(params),
                                                  std::move(returnType), StmtList());
        }
        default:
            return nullptr;
    }
}

// ============================================================
// Interface ... End Interface
// ============================================================

std::unique_ptr<InterfaceDecl> Parser::parseInterfaceDecl(
    std::vector<InterfaceAttr>& pendingAttrs) {
    auto loc = currentLoc();
    advance(); // 'Interface'

    auto d = std::make_unique<InterfaceDecl>(loc, std::string(), std::string());
    d->attributes = std::move(pendingAttrs);
    pendingAttrs.clear();

    if (!canBeName(cur_.kind)) {
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            "expected Interface name");
    } else {
        d->name = advance().text;
    }
    if (cur_.kind == TokenKind::LeftParen) {
        diag_.error(DiagnosticID::ParseInvalidInterfaceMember, currentLoc(),
            "Interface does not support generic type parameters (Of T): " + d->name);
        skipToNextLine();
        return d;
    }
    if (match(TokenKind::Extends)) {
        if (!canBeName(cur_.kind)) {
            diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                "expected parent interface name after 'Extends'");
        } else {
            // 限定名 A.B 的拼接口径与模块级 Implements (Fix 083) 一致
            d->extendsName = advance().text;
            while (cur_.kind == TokenKind::Dot && canBeName(peek2().kind)) {
                advance(); // '.'
                d->extendsName += "." + advance().text;
            }
        }
    }
    expectEndOfStatement();

    std::vector<InterfaceAttr> memberAttrs;
    while (cur_.kind != TokenKind::EndOfFile) {
        skipNewLines();
        if (cur_.kind == TokenKind::EndOfFile) break;
        if (cur_.kind == TokenKind::End && peek2().kind == TokenKind::Interface) break;
        if (cur_.kind == TokenKind::End && peek2().kind == TokenKind::EndOfFile) break;

        if (cur_.kind == TokenKind::End) {
            // 成员写了 `End Sub/Function/Property` = 接口里出现实现体
            diag_.error(DiagnosticID::ParseInvalidInterfaceMember, currentLoc(),
                "Interface member must not contain an implementation body (End " +
                    peek2().text + " instead of End Interface)");
            advance(); // 'End'
            advance(); // 'Sub'/'Function'/...
            expectEndOfStatement();
            continue;
        }

        if (atBracketAttrLine()) {
            InterfaceAttr a;
            if (parseBracketAttrLine(a)) memberAttrs.push_back(std::move(a));
            continue;  // 整行已消费 (失败路径已 skipToNextLine)
        }

        if (cur_.kind == TokenKind::Overridable || cur_.kind == TokenKind::Overrides ||
            cur_.kind == TokenKind::NotOverridable) {  // tB 扩展 (B08b): 契约块无实现, 谈不上覆盖
            diag_.error(DiagnosticID::ParseInvalidInterfaceMember, currentLoc(),
                "Interface member must not carry a virtual modifier (a contract has no"
                " implementation to dispatch): " + cur_.text);
            advance();
            continue;
        }

        if (cur_.kind == TokenKind::Public || cur_.kind == TokenKind::Private ||
            cur_.kind == TokenKind::Friend || cur_.kind == TokenKind::Static ||
            cur_.kind == TokenKind::Protected) {  // tB 扩展 (B08a)
            diag_.error(DiagnosticID::ParseInvalidInterfaceMember, currentLoc(),
                "Interface member must not carry an access modifier (contract members are public): " + cur_.text);
            advance();
            continue;
        }

        if (cur_.kind == TokenKind::Sub || cur_.kind == TokenKind::Function ||
            cur_.kind == TokenKind::Property) {
            InterfaceMember m;
            m.attributes.swap(memberAttrs);
            m.decl = parseInterfaceMemberDecl();
            if (m.decl) d->members.push_back(std::move(m));
            continue;
        }

        diag_.error(DiagnosticID::ParseInvalidInterfaceMember, currentLoc(),
            "Interface block accepts only Sub/Function/Property signatures, not: " + cur_.text);
        skipToNextLine();
    }

    if (!memberAttrs.empty()) {
        diag_.error(DiagnosticID::ParseUnknownAttribute, memberAttrs.back().loc,
            "Attribute line [" + memberAttrs.back().name + "] is not followed by an interface member");
    }

    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock, "expected 'End Interface'");
    expect(TokenKind::Interface, DiagnosticID::ParseMismatchedBlock,
           "expected 'End Interface'");
    return d;
}

// ============================================================
// CoClass ... End CoClass (tB 扩展; ai/026 四节 + ai/022 D44, 批次 B11/C01)
//
// C01 只负责"落到 AST": 块名与类型库 coclass 的优先级、[Implementation] 指向是否存在、
// [Default] 是否 ∈ 接口集合、宿主能否放块 —— 整片校验留给 C03, 身份求解留给 C02,
// `As`/`New`/`CreateObject` 的消费留给 C05. 产物对既有管线零影响 (Module::coclasses 无人读).
// ============================================================

std::unique_ptr<CoClassDecl> Parser::parseCoClassDecl(
    std::vector<InterfaceAttr>& pendingAttrs) {
    auto loc = currentLoc();
    advance(); // 'CoClass'

    auto d = std::make_unique<CoClassDecl>(loc, std::string());
    d->attributes = std::move(pendingAttrs);
    pendingAttrs.clear();

    if (!canBeName(cur_.kind)) {
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            "expected CoClass name");
    } else {
        d->name = advance().text;
    }
    expectEndOfStatement();

    std::vector<InterfaceAttr> refAttrs;
    while (cur_.kind != TokenKind::EndOfFile) {
        skipNewLines();
        if (cur_.kind == TokenKind::EndOfFile) break;
        if (cur_.kind == TokenKind::End && peek2().kind == TokenKind::CoClass) break;
        if (cur_.kind == TokenKind::End && peek2().kind == TokenKind::EndOfFile) break;

        if (atBracketAttrLine()) {
            InterfaceAttr a;
            // 同行可接声明: `[Default] Interface ICircle`
            if (parseBracketAttrLine(a, /*requireOwnLine=*/false)) {
                // 属性行归谁：026 四节的样本把身份四件套写在**首个引用之前**，而 [Default]
                // 只能贴在某个契约条目上 —— 于是按"名字 + 位置 + 是否同行"三条合起来判，
                // 不依赖空行 (空行不是语法).
                const bool entryOnly = (itfAsciiLower(a.name) == "default");
                const bool inlineRef = (cur_.kind == TokenKind::Interface);
                if (entryOnly || inlineRef || !d->ifaces.empty()) {
                    refAttrs.push_back(std::move(a));
                } else {
                    d->attributes.push_back(std::move(a));
                }
            }
            continue;
        }

        if (cur_.kind == TokenKind::Interface) {
            CoClassIfaceRef ref;
            ref.loc = currentLoc();
            advance(); // 'Interface'
            if (!canBeName(cur_.kind)) {
                diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                    "expected interface name after 'Interface' in CoClass block");
                skipToNextLine();
                continue;
            }
            // 限定名 A.B 的拼接口径与 Extends / 成员级 Implements 一致
            ref.ifaceName = advance().text;
            while (cur_.kind == TokenKind::Dot && canBeName(peek2().kind)) {
                advance(); // '.'
                ref.ifaceName += "." + advance().text;
            }
            expectEndOfStatement();
            for (auto& a : refAttrs) {
                // [Default] 折成 isDefault 标志, 不再留在 attributes 里 (免得 C03 有两处真相)
                if (itfAsciiLower(a.name) == "default") ref.isDefault = true;
                else ref.attributes.push_back(a);
            }
            refAttrs.clear();
            d->ifaces.push_back(std::move(ref));
            continue;
        }

        diag_.error(DiagnosticID::ParseUnexpectedToken, currentLoc(),
            "CoClass block accepts only attribute lines and 'Interface <name>' references, not: " +
                cur_.text);
        skipToNextLine();
    }

    if (!refAttrs.empty()) {
        diag_.error(DiagnosticID::ParseUnknownAttribute, refAttrs.back().loc,
            "Attribute line [" + refAttrs.back().name + "] is not followed by an interface reference");
    }

    expect(TokenKind::End, DiagnosticID::ParseMismatchedBlock, "expected 'End CoClass'");
    expect(TokenKind::CoClass, DiagnosticID::ParseMismatchedBlock,
           "expected 'End CoClass'");
    return d;
}

} // namespace vb6c3
