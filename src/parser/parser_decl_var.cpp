// vb6c3 - 声明解析器 — Declare / Event / Delegate / Const / Variable 声明 + 参数与类型引用
// 由 src/parser/parser_decl.cpp 拆出（2026-09-17），纯搬移、零行为改动。

#include "parser/parser.hpp"
#include "semantics/generics_registry.hpp"

namespace vb6c3 {


// ============================================================
// Declare 声明 (外部函数声明)
// ============================================================

// ai/024: `isWide` = 该声明由 `DeclareWide` 引入 (tB 兼容), 语义是"禁用
// ANSI<->Unicode 转换"。两者语法完全同形, 只是起始记号不同 (`Declare` 关键字 vs
// 标识符 `DeclareWide`), 故共用本函数。
std::unique_ptr<DeclareDecl> Parser::parseDeclareDecl(AccessLevel access, bool isWide) {
    auto loc = currentLoc();
    advance(); // consume 'Declare' / 'DeclareWide'

    bool isPtrSafe = false;
    if (match(TokenKind::Identifier) && toLower(cur_.text) == "ptrsafe") {
        isPtrSafe = true;
        advance(); // consume PtrSafe (实际是标识符, 词法器可能不会识别为关键字)
    }

    ProcKind procKind;
    if (match(TokenKind::Sub)) {
        procKind = ProcKind::Sub;
    } else if (match(TokenKind::Function)) {
        procKind = ProcKind::Function;
    } else {
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            "expected 'Sub' or 'Function' after 'Declare'");
        procKind = ProcKind::Sub;
    }

    auto nameTok = expectName("expected Declare name");

    // 与变量声明同口径 (Fix 028 / Task #40): 剥掉类型后缀 ($%&!#@),
    // 否则符号名带着 & 进符号表, 同名调用点被使用侧剥成一裸名, 两头对不上。
    auto suffixInfo = stripTypeSuffix(nameTok.text);
    const std::string& declareName = suffixInfo.name;

    expect(TokenKind::Lib, DiagnosticID::ParseExpectedToken,
           "expected 'Lib' in Declare statement");
    auto libTok = expect(TokenKind::StringLiteral, "expected library name string");

    std::string aliasName;
    if (match(TokenKind::Alias)) {
        auto aliasTok = expect(TokenKind::StringLiteral, "expected alias name string");
        aliasName = aliasTok.text;
    }

    // 调用约定: 默认 StdCall, 可选 CDecl
    CallConv callingConv = CallConv::StdCall;
    if (match(TokenKind::CDecl)) {
        callingConv = CallConv::CDecl;
    }

    auto params = parseParameterList();

    TypeRefPtr returnType;
    if (procKind == ProcKind::Function && match(TokenKind::As)) {
        returnType = parseTypeRef();
    }
    // Fix <vbeclipse> 2026-10-07: `Declare Function Foo& Lib …` 无 `As Type` 时,
    // 类型后缀即返回类型 (&→Long %→Integer $→String …), 与变量声明同一映射。
    // 不补这口的旧结局: DeclareFunc 符号落 Variant → inferExprType(调用) = Variant
    // → cgen 收窄检查把 API 调用结果 (x86 上句柄 / COLORREF, 值域 0x80000000..
    // 0xFFFFFFFF 属常态) 包进 vb6_ChkLong → 高位值误判 Overflow(Error 6)。
    // 实测 3DMenu TransBltNow 每个 BitBlt/SetBkColor/CreateCompatibleDC 结果
    // 都被包, 首次进入即崩; 原始 VB6 x86 把 32 位句柄原样存入 Long, 从不报 6。
    if (!returnType && procKind == ProcKind::Function && !suffixInfo.typeName.empty()) {
        returnType = std::make_unique<SimpleTypeRef>(loc, suffixInfo.typeName);
    }

    return std::make_unique<DeclareDecl>(loc, access, procKind,
        declareName, libTok.text, aliasName, callingConv, isPtrSafe,
        std::move(params), std::move(returnType), isWide);
}

// ============================================================
// Event 声明
// ============================================================

std::unique_ptr<EventDecl> Parser::parseEventDecl(AccessLevel access) {
    auto loc = currentLoc();
    advance(); // consume 'Event'
    auto nameTok = expectName("expected Event name");
    auto params = parseParameterList();
    expectEndOfStatement();
    return std::make_unique<EventDecl>(loc, access, nameTok.text, std::move(params));
}

// ============================================================
// Delegate 声明 (tB 扩展: 具名函数指针类型)
// ============================================================

std::unique_ptr<DelegateDecl> Parser::parseDelegateDecl(AccessLevel access) {
    auto loc = currentLoc();
    advance(); // consume 'Delegate'

    ProcKind procKind;
    if (match(TokenKind::Sub)) {
        procKind = ProcKind::Sub;
    } else if (match(TokenKind::Function)) {
        procKind = ProcKind::Function;
    } else {
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            "expected 'Sub' or 'Function' after 'Delegate'");
        procKind = ProcKind::Sub;
    }

    auto nameTok = expectName("expected Delegate name");

    // 调用约定: 默认 StdCall, 可选 CDecl (位置同 Declare: 参数表之前)
    CallConv callingConv = CallConv::StdCall;
    if (match(TokenKind::CDecl)) {
        callingConv = CallConv::CDecl;
    }

    auto params = parseParameterList();

    TypeRefPtr returnType;
    if (procKind == ProcKind::Function) {
        if (match(TokenKind::As)) {
            returnType = parseTypeRef();
        } else {
            diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                "expected 'As type' in Delegate Function declaration");
        }
    }
    expectEndOfStatement();

    return std::make_unique<DelegateDecl>(loc, access, procKind,
        nameTok.text, callingConv, std::move(params), std::move(returnType));
}

// ============================================================
// Const 声明
// ============================================================

std::unique_ptr<ConstDecl> Parser::parseConstDecl(AccessLevel access) {
    auto loc = currentLoc();
    advance(); // consume 'Const'

    auto nameTok = expectName("expected Const name");

    // Fix 028: 剥离 VB6 类型后缀 ($%&!#@)
    auto suffixInfo = stripTypeSuffix(nameTok.text);
    const std::string& constName = suffixInfo.name;

    TypeRefPtr asType;
    if (match(TokenKind::As)) {
        asType = parseTypeRef();
    }

    expect(TokenKind::Equals, DiagnosticID::ParseExpectedToken,
           "expected '=' in Const declaration");

    auto value = parseExpression();

    return std::make_unique<ConstDecl>(loc, access, constName,
        std::move(asType), std::move(value));
}

// Const 列表: Const A = 1, B = 2, C As Long = 3
DeclPtr Parser::parseConstDeclList(AccessLevel access) {
    auto first = parseConstDecl(access);
    if (cur_.kind != TokenKind::Comma) {
        return first;
    }
    DeclList decls;
    decls.push_back(std::move(first));
    while (match(TokenKind::Comma)) {
        auto loc = currentLoc();
        auto nameTok = expectName("expected Const name");
        // Fix 028: 剥离 VB6 类型后缀 ($%&!#@)
        auto suffixInfo = stripTypeSuffix(nameTok.text);
        const std::string& constName = suffixInfo.name;
        TypeRefPtr asType;
        if (match(TokenKind::As)) {
            asType = parseTypeRef();
        }
        expect(TokenKind::Equals, DiagnosticID::ParseExpectedToken,
               "expected '=' in Const declaration");
        auto value = parseExpression();
        decls.push_back(std::make_unique<ConstDecl>(loc, access, constName,
            std::move(asType), std::move(value)));
    }
    return std::make_unique<MultiDecl>(decls[0]->loc, std::move(decls));
}

// ============================================================
// Variable 声明
// ============================================================

std::unique_ptr<VariableDecl> Parser::parseVariableDecl(AccessLevel access, bool isStatic) {
    auto loc = currentLoc();
    // 注意: Dim/Public/Private/Static 已由调用方消费
    // 但如果从 parseDeclaration 直接调用, Dim 尚未消费
    if (cur_.kind == TokenKind::Dim || cur_.kind == TokenKind::Public ||
        cur_.kind == TokenKind::Private || cur_.kind == TokenKind::Static ||
        cur_.kind == TokenKind::Global || cur_.kind == TokenKind::Friend ||
        cur_.kind == TokenKind::Protected) {  // tB 扩展 (B08a)
        advance();
    }

    bool isWithEvents = false;
    if (match(TokenKind::WithEvents)) {
        isWithEvents = true;
    }

    auto nameTok = expectName("expected variable name");

    // Fix 028: 剥离 VB6 类型后缀 ($%&!#@), 防止 cIdent 把 $ 改成 _ 导致
    // 声明名(BSTR/VARIANT Address_) 与使用名(Address) 不一致 → C2065。
    // 注意: 不注入后缀对应的类型, 保留默认 Variant 行为(避免引入 C2440)。
    auto suffixInfo = stripTypeSuffix(nameTok.text);
    const std::string& varName = suffixInfo.name;

    // 数组维度: dim a(1 To 10, 1 To 20) As Long
    // 或动态数组: dim a() As Long
    std::vector<VariableDecl::Dimension> dimensions;
    bool isDynamicArray = false;
    if (match(TokenKind::LeftParen)) {
        // 空括号 = 动态数组 (Dim a() As Long)
        if (cur_.kind != TokenKind::RightParen) {
            do {
                VariableDecl::Dimension dim;
                auto first = parseExpression();
                if (match(TokenKind::To)) {
                    dim.lower = std::move(first);
                    dim.upper = parseExpression();
                } else {
                    dim.upper = std::move(first);
                }
                dimensions.push_back(std::move(dim));
            } while (match(TokenKind::Comma));
        } else {
            // 空括号 () = 动态数组
            isDynamicArray = true;
        }
        expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
               "expected ')'");
    }

    bool isNew = false;
    TypeRefPtr asType;
    if (match(TokenKind::As)) {
        if (match(TokenKind::New)) {
            isNew = true;
        }
        asType = parseTypeRef();
    }

    // Task #40: VB6 类型后缀即类型声明 (`Dim dl&` ≡ `Dim dl As Long`)。
    // Fix 028 只剥名不注类型 (旧注释称避免引入 C2440), 变量落默认 Variant —
    // 结果 cDlg.cls GetLocaleMeasureSystem 的 dl/nullpos 走 Variant 链路:
    // 使用点 `dl& = GetLocaleInfo(...)` 在 cgen 里发 dl_ (cIdent 吃 &) +
    // Variant 解包, C2440/C2065 双炸。此处无 As Type 时按后缀补声明类型
    // ($→String %→Integer &→Long !→Single #→Double @→Currency, 映射与
    // stripTypeSuffix 一致), 与使用点剥后缀 (parseIdentifierOrCall) 配套。
    if (!asType && !suffixInfo.typeName.empty()) {
        asType = std::make_unique<SimpleTypeRef>(loc, suffixInfo.typeName);
    }

    ExprPtr initializer;
    if (match(TokenKind::Equals)) {
        initializer = parseExpression();
    }

    return std::make_unique<VariableDecl>(loc, access, varName,
        isWithEvents, isStatic, isNew, std::move(asType), std::move(initializer),
        std::move(dimensions), isDynamicArray);
}

// Variable 列表: Dim a, b As Long, c As String
// parseVariableDecl 会条件性地消费 Dim/Private/Public 关键字,
// 对逗号后的后续变量, 当前 token 是变量名而非关键字, 所以可以直接复用
DeclPtr Parser::parseVariableDeclList(AccessLevel access, bool isStatic) {
    auto first = parseVariableDecl(access, isStatic);
    if (cur_.kind != TokenKind::Comma) {
        return first;
    }
    DeclList decls;
    decls.push_back(std::move(first));
    while (match(TokenKind::Comma)) {
        decls.push_back(parseVariableDecl(access, isStatic));
    }
    return std::make_unique<MultiDecl>(decls[0]->loc, std::move(decls));
}

// ============================================================
// 参数列表
// ============================================================

std::vector<std::unique_ptr<ParameterDecl>> Parser::parseParameterList() {
    std::vector<std::unique_ptr<ParameterDecl>> params;

    if (!match(TokenKind::LeftParen)) {
        return params;  // 无参数列表
    }

    skipNewLines();
    if (cur_.kind != TokenKind::RightParen) {
        do {
            skipNewLines();
            params.push_back(parseParameter());
            skipNewLines();
        } while (match(TokenKind::Comma));
    }

    expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
           "expected ')'");
    return params;
}

std::unique_ptr<ParameterDecl> Parser::parseParameter() {
    auto loc = currentLoc();

    bool isOptional = false;
    bool isByVal = false;
    bool isParamArray = false;

    if (match(TokenKind::Optional)) {
        isOptional = true;
    }
    if (match(TokenKind::ByVal)) {
        isByVal = true;
    } else if (match(TokenKind::ByRef)) {
        isByVal = false;  // 显式 ByRef (默认)
    }
    if (match(TokenKind::ParamArray)) {
        isParamArray = true;
    }

    auto nameTok = expectName("expected parameter name");

    // Fix 028: 剥离 VB6 类型后缀 ($%&!#@), 见 parseVariableDecl 同样说明。
    auto suffixInfo = stripTypeSuffix(nameTok.text);
    const std::string& paramName = suffixInfo.name;

    // 数组参数: name() As Type
    bool isArrayParam = false;
    if (match(TokenKind::LeftParen)) {
        expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
               "expected ')'");
        isArrayParam = true;
    }

    TypeRefPtr asType;
    if (match(TokenKind::As)) {
        if (cur_.kind == TokenKind::Any) {
            // Declare 中的 Any 类型
            advance();
            asType = std::make_unique<SimpleTypeRef>(currentLoc(), "Any");
        } else {
            asType = parseTypeRef();
        }
    }

    // Task #40 (变量) / parseDeclareDecl 同口径: `ByVal hDestDC&` 的类型后缀
    // 即实参类型 (&→Long …)。不注入的话形参落 Variant, 调用点实参被包
    // vb6_VariantFromValue 塞进元素为 int32_t 的 C 签名 → C2172/C2440。
    // (3DMenu 的 gdi32 Declare 全套用 `ByVal x&` 写法, 实测必炸。)
    if (!asType && !suffixInfo.typeName.empty()) {
        asType = std::make_unique<SimpleTypeRef>(loc, suffixInfo.typeName);
    }

    ExprPtr defaultValue;
    if (isOptional && match(TokenKind::Equals)) {
        defaultValue = parseExpression();
    }

    // 数组参数: name() As Type 需要将类型包装为 ArrayTypeRef
    if (isArrayParam && asType && asType->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simple = static_cast<SimpleTypeRef&>(*asType);
        auto elemType = std::make_unique<SimpleTypeRef>(simple.loc, simple.name);
        asType = std::make_unique<ArrayTypeRef>(simple.loc, std::move(elemType), std::vector<ArrayTypeRef::Dimension>{});
    }

    return std::make_unique<ParameterDecl>(loc, paramName, isOptional,
        isByVal, isParamArray, std::move(asType), std::move(defaultValue));
}

// ============================================================
// 类型引用
// ============================================================

TypeRefPtr Parser::parseTypeRef() {
    auto loc = currentLoc();

    // 简单类型名: Long, String, MyClass 等
    // (C3 扩展 ai/032: SByte/UInteger/ULong/ULongLong 追加在同一条接受集里 —— 它们
    //  在语义层按类型名解析, 在词法层是软关键字, 与 LongLong/LongPtr 完全同构)
    if (cur_.kind == TokenKind::Identifier || cur_.kind == TokenKind::Boolean ||
        cur_.kind == TokenKind::Byte || cur_.kind == TokenKind::Integer ||
        cur_.kind == TokenKind::Long || cur_.kind == TokenKind::LongLong ||
        cur_.kind == TokenKind::LongPtr || cur_.kind == TokenKind::Single ||
        cur_.kind == TokenKind::Double || cur_.kind == TokenKind::Currency ||
        cur_.kind == TokenKind::Decimal || cur_.kind == TokenKind::Date ||
        cur_.kind == TokenKind::Object || cur_.kind == TokenKind::String ||
        cur_.kind == TokenKind::Variant ||
        cur_.kind == TokenKind::SByte || cur_.kind == TokenKind::UInteger ||
        cur_.kind == TokenKind::ULong || cur_.kind == TokenKind::ULongLong) {
        auto tok = advance();
        auto typeRef = std::make_unique<SimpleTypeRef>(loc, tok.text);

        // 限定类型名: Scripting.Dictionary, MSComctlLib.ImageList 等
        if (cur_.kind == TokenKind::Dot) {
            std::string qualified = tok.text;
            while (cur_.kind == TokenKind::Dot) {
                advance(); // consume '.'
            // VB6 允许关键字作为限定类型名组件: ADODB.Error, ADODB.Command 等
            // 与 parsePostfix() Dot 分支一致, 使用 canBeName + 文本检查
            if (canBeName(cur_.kind)) {
                qualified += "." + advance().text;
            } else if (!cur_.text.empty() && cur_.kind != TokenKind::EndOfFile &&
                       cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::Colon &&
                       cur_.kind != TokenKind::LeftParen && cur_.kind != TokenKind::RightParen &&
                       cur_.kind != TokenKind::Comma) {
                // 硬关键字也可作为限定类型名组件 (如 Error, Command, Type 等)
                qualified += "." + advance().text;
            } else {
                break;
            }
            }
            typeRef = std::make_unique<SimpleTypeRef>(loc, qualified);
        }

        // 泛型使用点 (tB): As Foo(Of Long) —— 在数组下标判定之前拦截.
        // 守卫 (下一 token 必须是 Identifier "Of") 使 T()/T(10) 数组形态不受影响.
        {
            std::string baseName = static_cast<SimpleTypeRef*>(typeRef.get())->name;
            if (tryFlattenGenericName(baseName)) {
                typeRef = std::make_unique<SimpleTypeRef>(loc, baseName);
            }
        }

        // 数组类型: Long(), String(10)
        if (match(TokenKind::LeftParen)) {
            std::vector<ArrayTypeRef::Dimension> dims;
            if (cur_.kind != TokenKind::RightParen) {
                do {
                    ArrayTypeRef::Dimension dim;
                    auto first = parseExpression();
                    if (match(TokenKind::To)) {
                        dim.lower = std::move(first);
                        dim.upper = parseExpression();
                    } else {
                        dim.upper = std::move(first);
                    }
                    dims.push_back(std::move(dim));
                } while (match(TokenKind::Comma));
            }
            expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
                   "expected ')'");
            return std::make_unique<ArrayTypeRef>(loc, std::move(typeRef), std::move(dims));
        }

        // 定长字符串: String * N
        if (tok.kind == TokenKind::String && match(TokenKind::Star)) {
            auto len = parseExpression();
            return std::make_unique<FixedStringTypeRef>(loc, std::move(len));
        }

        return typeRef;
    }

    // 括号标识符类型: [My Type]
    if (cur_.kind == TokenKind::LeftBracket) {
        advance();
        auto nameTok = expectName("expected type name");
        expect(TokenKind::RightBracket, DiagnosticID::ParseExpectedToken,
               "expected ']'");
        return std::make_unique<SimpleTypeRef>(loc, nameTok.text);
    }

    diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
        "expected type name");
    return std::make_unique<SimpleTypeRef>(loc, "Variant");
}

// ============================================================
// 泛型 (tB 扩展, G1) — 声明侧类型参数表 + 使用侧扁名化
// ============================================================
// 设计 (与用户确认的计划冻结版一致):
//   - "Of" 非词法关键字 (避免动关键字表影响存量), 以 Identifier 文本上下文
//     识别; 触发形态唯一: '(' 后紧跟 Identifier "Of" (大小写无关), 与调用
//     实参表 (x) / 数组 T() / T(10) 天然无歧义.
//   - 使用点 Foo(Of Long) 在 parse 期即改写为扁名 (编码见 makeFlatGenericName),
//     下游 (语义/跨模块/cgen) 只见普通名字, 泛型仅活在泛型器 (G2) 视角内.
//   - $ 不是合法 VB 标识符字符 → 扁名不可能撞真实名 (cIdent 统一转 '_').

std::string Parser::makeFlatGenericName(const std::string& base,
                                        const std::vector<std::string>& args) {
    // 纯下划线编码 (G2 调试结论): '$' 方案会让符号键与 cIdent 产物两种形态在
    // 各按名查找点系统性错位. 单一实现见 semantics/generics_registry.hpp
    // (driver 物化器 / analyzer 推断共用).
    return genMakeFlat(base, args);
}

std::vector<std::string> Parser::parseTypeParams() {
    std::vector<std::string> out;
    // 起点 = 类模块类型参数 (泛型类成员体内仍可见 T, 供 Box(Of T) 护栏命中);
    // .bas/普通类模块 outerTypeParams_ 恒空 → 与原清空行为逐字节等价.
    curTypeParams_ = outerTypeParams_;
    if (!check(TokenKind::LeftParen)) return out;
    if (peek2().kind != TokenKind::Identifier || toLower(peek2().text) != "of")
        return out;
    advance(); // '('
    advance(); // 'Of'
    do {
        auto t = expectName("expected type parameter name");
        curTypeParams_.push_back(toLower(t.text));
        out.push_back(t.text);
    } while (match(TokenKind::Comma));
    expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
           "expected ')' after type parameters");
    return out;
}

std::string Parser::parseGenericArgFlat() {
    // 泛型实参 = 类型名 (与 parseTypeRef 同一接受集, 含关键字类型 Long/String…)
    // + 可选点号限定 + 可选嵌套 (Of …) 递归扁平. 数组实参 (T()) v1 不支持.
    if (!(cur_.kind == TokenKind::Identifier || cur_.kind == TokenKind::Boolean ||
          cur_.kind == TokenKind::Byte || cur_.kind == TokenKind::Integer ||
          cur_.kind == TokenKind::Long || cur_.kind == TokenKind::LongLong ||
          cur_.kind == TokenKind::LongPtr || cur_.kind == TokenKind::Single ||
          cur_.kind == TokenKind::Double || cur_.kind == TokenKind::Currency ||
          cur_.kind == TokenKind::Decimal || cur_.kind == TokenKind::Date ||
          cur_.kind == TokenKind::Object || cur_.kind == TokenKind::String ||
          cur_.kind == TokenKind::Variant ||
          // C3 扩展 (ai/032): 与 parseTypeRef 保持同一接受集
          cur_.kind == TokenKind::SByte || cur_.kind == TokenKind::UInteger ||
          cur_.kind == TokenKind::ULong || cur_.kind == TokenKind::ULongLong)) {
        diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
            "expected type name in generic type argument list");
        return "";
    }
    std::string name = advance().text;
    while (cur_.kind == TokenKind::Dot) {
        advance();
        auto nextTok = expectName("expected qualified type name");
        name += "." + nextTok.text;
    }
    tryFlattenGenericName(name);  // 嵌套: Bar(Of Long) → Bar$gen1$Long
    return name;
}

bool Parser::tryFlattenGenericName(std::string& ioName) {
    if (!check(TokenKind::LeftParen)) return false;
    if (peek2().kind != TokenKind::Identifier || toLower(peek2().text) != "of")
        return false;
    auto loc = currentLoc();
    (void)loc;
    advance(); // '('
    advance(); // 'Of'
    std::vector<std::string> args;
    do {
        std::string a = parseGenericArgFlat();
        if (a.empty()) {  // 出错恢复: 跳到 ')'
            while (!check(TokenKind::RightParen) &&
                   !check(TokenKind::EndOfFile) && !check(TokenKind::NewLine))
                advance();
            match(TokenKind::RightParen);
            return false;
        }
        args.push_back(a);
    } while (match(TokenKind::Comma));
    expect(TokenKind::RightParen, DiagnosticID::ParseExpectedToken,
           "expected ')' after generic type arguments");

    // v1 护栏: 模板体内把类型参数当泛型实参 (Box(Of T)) 明确拒绝 — 使用点在
    // parse 期即折成扁名, T 已被折进串里, 克隆期的 subst (作用于 AST 类型
    // 引用) 再也替换不到它; 放行会静默产出对不存在类型的引用.
    for (auto& a : args) {
        for (auto& tp : curTypeParams_) {
            if (toLower(a) == tp) {
                diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
                    "泛型模板体内不能以类型参数 '" + a + "' 作泛型类型实参 (v1 不支持)");
                return false;
            }
        }
    }

    std::string flat = makeFlatGenericName(ioName, args);
    // 扁名整体小写化: VB 名字大小写不敏感, C 名大小写敏感 — 保留书写大小写会
    // 把同一实例化特化成两个类型 (Box_G1_Long vs box_g1_long, C2079 实测).
    // 小写扁名天然按"同一实例化"去重; 与用户名的撞车由泛型器物化时显式检测.
    flat = toLower(flat);
    GenericUse use;
    use.base = ioName;
    use.args = args;
    flatGenerics_[flat] = std::move(use);
    ioName = flat;
    return true;
}

} // namespace vb6c3
