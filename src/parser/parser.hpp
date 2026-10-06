#pragma once
// vb6c3 - Visual Basic 6.0 Compiler
// 递归下降解析器 + Pratt 表达式解析
// 参考: FreeBASIC parser.bis, RustASP Pratt (l_bp, r_bp) 设计, VB6.g4 ANTLR

#include "preprocessor/preprocessor.hpp"
#include "ast/ast.hpp"
#include "common/diagnostics.hpp"
#include <memory>
#include <string>
#include <vector>
#include <unordered_map>

namespace vb6c3 {

// ============================================================
// Pratt 优先级表
// VB6 运算符优先级 (从低到高, 14级)
// l_bp = 左绑定力, r_bp = 右绑定力
// 左结合: l_bp < r_bp (如 +: 5,6)
// 右结合: l_bp > r_bp (如 ^: 12,11)
// ============================================================
struct BindingPower {
    int l_bp;  // 左绑定力 (null denotation 时为 0)
    int r_bp;  // 右绑定力
};

class Parser {
public:
    // 使用预处理器 (含条件编译支持)
    Parser(std::shared_ptr<SourceBuffer> buffer, Diagnostics& diag,
           const PreprocessOptions& ppOpts = PreprocessOptions());

    // 主入口: 解析整个模块
    // isClassModule: true=.cls类模块, false=.bas标准模块/.frm窗体模块
    std::unique_ptr<Module> parseModule(bool isClassModule = false);

    // 诊断查询
    bool hasErrors() const { return diag_.hasErrors(); }

    // ============================================================
    // 泛型 (tB 扩展, G1): 使用点扁平时登记的结构
    //   使用点 Foo(Of Long) 在 parse 期即改写为扁名 (VB 标识符不含 '$',
    //   故扁名不可能与真实名冲突), 同时把 flat → (模板基名, 实参表) 记入
    //   flatGenerics_, 供 G2 泛型器物化 (免字符串反解码).
    //   编码: Base + "$gen" + <arity> + "$" + join(args, "$$"),
    //   实参自身可为嵌套扁名 (递归自描述).
    // ============================================================
    struct GenericUse {
        std::string base;               // 模板基名 (源码原大小写)
        std::vector<std::string> args;  // 类型实参 (各自已扁平)
    };
    // key = 扁名小写; 同 key 覆盖写 (内容同源必然一致, 幂等)
    const std::unordered_map<std::string, GenericUse>& flatGenerics() const {
        return flatGenerics_;
    }
    static std::string makeFlatGenericName(const std::string& base,
                                           const std::vector<std::string>& args);

private:
    // ============================================================
    // Token 消费接口
    // ============================================================

    // 前瞻 (不消费)
    const Token& peek() const;
    const Token& peek2() const;

    // 消费当前 token 并返回 (自动跳过 Comment)
    Token advance();

    // 从词法器获取下一个非 Comment token
    Token fetchNextToken();

    // 消费当前 token, 断言类型匹配; 不匹配则报错并返回当前 token
    Token expect(TokenKind kind, DiagnosticID diagId, const std::string& msg);
    Token expect(TokenKind kind, const std::string& msg);

    // 如果当前 token 匹配则消费, 否则返回 false
    bool match(TokenKind kind);

    // 如果当前 token 匹配则消费, 否则报错
    // 返回是否成功匹配
    bool expectOrSkip(TokenKind kind, DiagnosticID diagId, const std::string& msg);

    // 当前 token 类型判断
    bool check(TokenKind kind) const;
    bool checkAny(std::initializer_list<TokenKind> kinds) const;

    // 判断是不是语句开始 token
    bool isStatementStart() const;
    bool isDeclarationStart() const;

    // 新行处理: VB6 新行是语句分隔符
    // 跳过零或多个 NewLine (同时跳过 Comment)
    void skipNewLines();
    // 期望至少一个 NewLine (语句终止), 变体: 也接受冒号
    bool expectEndOfStatement();

    // 标识符解析: 接受 Identifier 或软关键字 (VB6 允许 Name/String/Step 等作为标识符)
    bool isSoftKeyword(TokenKind kind) const;
    bool canBeName(TokenKind kind) const;
    Token expectName(const std::string& msg);
    // 标签目标: 命名标签 或 VB6 **行号标签** (`GoTo 100` / `Resume 200` / `On Error GoTo 300`)
    std::string expectLabelTarget(const std::string& msg);

    // 获取当前源码位置
    SourceLocation currentLoc() const;

    // ============================================================
    // 模块解析 (最高层)
    // ============================================================
    void parseModuleBody(Module& mod);
    std::unique_ptr<OptionStmt> parseOption();
    std::unique_ptr<ImplementsStmt> parseImplements();
    InheritsStmt parseInherits();  // 类继承子句 `Inherits Base` (tB 扩展, ai/022 B07)
    std::unique_ptr<DefTypeStmt> parseDefType();
    std::unique_ptr<AttributeStmt> parseAttribute();

    // ============================================================
    // 声明解析 (parser_decl.cpp)
    // ============================================================
    DeclPtr parseDeclaration();
    // 虚方法修饰位 (tB 扩展, ai/022 B08b): 见 parser_decl.cpp 的 parseDeclaration 头注释
    void eatVirtualModifiers(ProcVirt& io);
    void checkVirtualOnProcStart(ProcVirt& io);
    DeclPtr parseVariableDeclList(AccessLevel access, bool isStatic);
    DeclPtr parseConstDeclList(AccessLevel access);
    std::unique_ptr<SubDecl> parseSubDecl(AccessLevel access, bool isStatic);
    std::unique_ptr<FunctionDecl> parseFunctionDecl(AccessLevel access, bool isStatic);
    std::unique_ptr<PropertyDecl> parsePropertyDecl(AccessLevel access);
    std::unique_ptr<TypeDecl> parseTypeDecl(AccessLevel access);
    std::unique_ptr<EnumDecl> parseEnumDecl(AccessLevel access);
    // isWide = 由 `DeclareWide` 引入 (tB 兼容, ai/024): 禁用 ANSI<->Unicode 转换
    std::unique_ptr<DeclareDecl> parseDeclareDecl(AccessLevel access, bool isWide = false);
    std::unique_ptr<EventDecl> parseEventDecl(AccessLevel access);
    std::unique_ptr<DelegateDecl> parseDelegateDecl(AccessLevel access);
    // Interface 契约块 (tB 扩展, parser_interface.cpp): `Interface Name [Extends P] ... End Interface`
    // pendingAttrs = 声明上方累积的 [Xxx] 属性行 (由 parseModuleBody 交出)
    std::unique_ptr<InterfaceDecl> parseInterfaceDecl(std::vector<InterfaceAttr>& pendingAttrs);
    // CoClass 契约聚合块 (tB 扩展, ai/026 四节 / ai/022 D44, 批次 B11/C01; 同在 parser_interface.cpp):
    // `CoClass Name ... End CoClass`, 块体只收属性行与 `Interface <名>` 引用行. C01 不校验、不发码.
    std::unique_ptr<CoClassDecl> parseCoClassDecl(std::vector<InterfaceAttr>& pendingAttrs);
    DeclPtr parseInterfaceMemberDecl();
    // 成员级 `Implements I.M[, I.N]` 尾子句 (tB 扩展, ai/022 D5, B02b):
    // 在过程签名之后、行尾之前调用; 无 Implements 时不消费任何 token.
    void parseTrailingImplementsClauses(std::vector<ImplementsClause>& out);
    bool atBracketAttrLine() const;
    // requireOwnLine=false: 属性行后同行还可以接声明 (CoClass 块的 `[Default] Interface X`)
    bool parseBracketAttrLine(InterfaceAttr& out, bool requireOwnLine = true);
    std::unique_ptr<ConstDecl> parseConstDecl(AccessLevel access);
    std::unique_ptr<VariableDecl> parseVariableDecl(AccessLevel access, bool isStatic);

    // 参数列表
    std::vector<std::unique_ptr<ParameterDecl>> parseParameterList();
    std::unique_ptr<ParameterDecl> parseParameter();

    // 类型引用
    TypeRefPtr parseTypeRef();

    // 泛型 (tB 扩展, G1)
    // 声明侧: proc/Type 名后的 (Of T[, U]) 类型参数表; 非该形态则零消耗返回空.
    std::vector<std::string> parseTypeParams();
    // 使用侧: 名字后紧跟 (Of A[, B]) 时消费尾巴, ioName 改为扁名并登记
    // flatGenerics_. 守卫: '(' 的下一个 token 必须是 Identifier "Of" (大小写
    // 无关), 与数组下标 T() / 实参表 (x) 天然区分.
    bool tryFlattenGenericName(std::string& ioName);
    // 泛型实参位上的类型名: 基础名(可点号限定) + 递归 (Of ...) 扁平.
    std::string parseGenericArgFlat();

    // ============================================================
    // 语句解析 (parser_stmt.cpp)
    // ============================================================
    StmtPtr parseStatement();
    StmtList parseBlock(TokenKind endKind1, TokenKind endKind2 = TokenKind::EndOfFile);
    StmtList parseBlockUntil(std::initializer_list<TokenKind> endKinds);

    // 块语句
    std::unique_ptr<IfStmt> parseIfStmt();
    StmtPtr parseForOrForEach();
    std::unique_ptr<ForStmt> parseForStmt();
    std::unique_ptr<ForEachStmt> parseForEachStmt();
    // `Next a, b, c` 逗号列表: 真实 Next 归当前循环, 多余变量排队供外层消费
    void consumeNextClause(const std::string& loopVar);
    std::vector<std::string> pendingNextVars_;
    std::unique_ptr<DoLoopStmt> parseDoLoopStmt();
    std::unique_ptr<WhileWendStmt> parseWhileWendStmt();
    std::unique_ptr<SelectCaseStmt> parseSelectCaseStmt();
    std::unique_ptr<WithStmt> parseWithStmt();

    // ai/vb-asm-extension-spec: `<Naked>` 角括号过程属性。角括号属性行 (VB6 家族没有这条路)
    // 只承载 Naked 一个名字; 解析后暂存, 由紧随其后的 Sub/Function 声明取走。
    bool pendingNaked_ = false;
    bool tryParseAngleAttr();   // 命中并消费 `<Naked>` 返回 true; 否则不动 token 流

    // 单行语句
    StmtPtr parseOnStmt();
    std::unique_ptr<OnErrorStmt> parseOnErrorStmt();
    std::unique_ptr<ResumeStmt> parseResumeStmt();
    std::unique_ptr<ErrorStmt> parseErrorStmt();
    StmtPtr parseMidStmt();  // P18-A
    std::unique_ptr<OnGoToStmt> parseOnGoToStmt();
    std::unique_ptr<OnGoSubStmt> parseOnGoSubStmt();
    std::unique_ptr<ExitStmt> parseExitStmt();
    std::unique_ptr<GoToStmt> parseGoToStmt();
    std::unique_ptr<GoSubStmt> parseGoSubStmt();
    std::unique_ptr<ReturnStmt> parseReturnStmt();
    std::unique_ptr<ReDimStmt> parseReDimStmt();
    std::unique_ptr<EraseStmt> parseEraseStmt();
    // 账 #186: 点链字符串 → 表达式树 (ReDim 与 Erase **共用这一处**，别再抄第二份)。
    // 前置 '.' 表示 With 块成员 (WithMemberExpr)，其余逐段 MemberAccessExpr。
    ExprPtr buildDottedNameExpr(const std::string& nm, SourceLocation loc);
    std::unique_ptr<RaiseEventStmt> parseRaiseEventStmt();
    std::unique_ptr<EndStmt> parseEndStmt();
    std::unique_ptr<StopStmt> parseStopStmt();
    // ai/vb-asm-extension-spec: Asm ... End Asm (原始行捕获, 不做 VB 语法解析)
    StmtPtr parseAsmStmt();
    std::unique_ptr<SetStmt> parseSetStmt();
    std::unique_ptr<LetStmt> parseLetStmt();
    std::unique_ptr<CallStmt> parseCallStmt();
    StmtPtr parseDimStmt();
    StmtPtr parseConstStmtInBody();
    StmtPtr parseStaticStmtInBody();
    StmtPtr parseAccessDeclInBody();
    // 体级声明的单一约定: 一条声明符一条 LocalDeclStmt (账 #215)
    StmtPtr wrapBodyDecls(SourceLocation loc, DeclPtr decl);

    // 行标签/赋值/调用 (两可: label: 或 x = 1 或 proc args)
    StmtPtr parseLabelOrAssignmentOrCall();

    // 文件 I/O 语句
    std::unique_ptr<OpenStmt> parseOpenStmt();
    std::unique_ptr<CloseStmt> parseCloseStmt();
    std::unique_ptr<GetStmt> parseGetStmt();
    std::unique_ptr<PutStmt> parsePutStmt();
    std::unique_ptr<InputStmt> parseInputStmt();
    std::unique_ptr<PrintStmt> parsePrintStmt();
    std::unique_ptr<WriteStmt> parseWriteStmt();
    std::unique_ptr<LineInputStmt> parseLineInputStmt();
    std::unique_ptr<WidthStmt> parseWidthStmt();
    std::unique_ptr<SeekStmt> parseSeekStmt();
    std::unique_ptr<LockStmt> parseLockStmt();
    std::unique_ptr<UnlockStmt> parseUnlockStmt();
    std::unique_ptr<NameStmt> parseNameStmt();
    std::unique_ptr<FileCopyStmt> parseFileCopyStmt();
    std::unique_ptr<KillStmt> parseKillStmt();
    std::unique_ptr<MkDirStmt> parseMkDirStmt();
    std::unique_ptr<RmDirStmt> parseRmDirStmt();
    std::unique_ptr<ChDirStmt> parseChDirStmt();
    std::unique_ptr<ChDriveStmt> parseChDriveStmt();

    // 杂项语句
    StmtPtr parseBeepOrDoEvents();
    std::unique_ptr<AttributeStmt> parseAttributeInBody();

    // ============================================================
    // 表达式解析 (parser_expr.cpp) — Pratt 解析器
    // ============================================================
    ExprPtr parseExpression();
    ExprPtr parseExpression(int minBp);        // Pratt 核心: 最小绑定力
    ExprPtr parseNullDenotation();             // 前缀/原子表达式
    ExprPtr parseLeftDenotation(ExprPtr left, int& minBp);  // 中缀/后缀

    // 原子表达式辅助
    ExprPtr parseLiteral();
    ExprPtr parseIdentifierOrCall();
    ExprPtr parseParenthesizedExpr();
    ExprPtr parseNewExpr();
    ExprPtr parseTypeOfExpr();
    ExprPtr parseAddressOfExpr();
    ExprPtr parseMeExpr();
    ExprPtr parseWithMemberExpr();

    // 后缀: .member, !dict, (args)
    ExprPtr parsePostfix(ExprPtr expr);

    // Case 值解析 (Is > 0, 1 To 10, 等变体)
    CaseClause::CaseValue parseCaseValue();

    // ============================================================
    // 运算符优先级查找
    // ============================================================
    BindingPower getBindingPower(TokenKind kind) const;

    // TokenKind → BinaryOp 映射
    BinaryOp tokenToBinaryOp(TokenKind kind) const;

    // 判断 token 是否是中缀运算符
    bool isInfixOperator(TokenKind kind) const;

    // 判断 token 是否是前缀运算符
    bool isPrefixOperator(TokenKind kind) const;

    // ============================================================
    // 错误恢复
    // ============================================================
    // 跳过 token 直到遇到语句开始符号或同步点
    void synchronize();
    // 跳过到行尾
    void skipToNextLine();

    // ============================================================
    // 辅助
    // ============================================================
    // 标识符比较 (VB6 不区分大小写)
    bool identifierEquals(const std::string& a, const std::string& b) const;
    // Token 文本转小写
    std::string toLower(const std::string& s) const;

    // 判断 End 后面跟着什么 (End If, End Sub, End Function, ...)
    // 返回 true 如果 End + next 构成 End XXX 块终止符
    bool isEndBlock() const;

    // Fix 028: VB6 标识符类型后缀 ($ % & ! # @) 处理
    // 词法器 scanIdentifierOrKeyword() 在类型后缀字符后跟非字母数字时,
    // 会把后缀并入标识符文本 (如 'Dim x$' 得到 text="x$")。
    // 这导致 cIdent() 把 $ 替换为 _ 后, 声明名为 'x_' 但使用名为 'x'
    // (使用处后缀不并入), 形成 C2065 "未声明的标识符" 错误。
    // stripTypeSuffix 剥离标识符末尾的类型后缀, 返回剥离后的名字
    // 和后缀对应的 VB6 类型名(空表示无后缀)。Fix 028 暂不进行类型注入,
    // 保留 Variant 默认类型(避免引入 Variant→BSTR/Long 的 C2440)。
    struct TypeSuffixStrip { std::string name; std::string typeName; };
    TypeSuffixStrip stripTypeSuffix(const std::string& text) const;

private:
    Preprocessor preproc_;
    Diagnostics& diag_;
    std::shared_ptr<SourceBuffer> buffer_;

    // 泛型 (tB 扩展, G1): 使用点扁名 → 结构登记表 (public flatGenerics() 暴露)
    std::unordered_map<std::string, GenericUse> flatGenerics_;
    // 泛型 (tB 扩展, G3) 护栏: 正在解析的模板声明的类型参数名 (小写).
    // tryFlattenGenericName 据此拒绝"体内以类型参数作泛型实参" (v1 不支持),
    // 因扁名在 parse 期固化, 克隆期的 AST 类型替换无法再触及被折进串里的 T.
    std::vector<std::string> curTypeParams_;
    // 泛型类 (G4): 类头行 (Of T) 参数的模块层驻留 (成员声明的 curTypeParams_ 起点)
    std::vector<std::string> outerTypeParams_;
    bool classHeaderSeen_ = false;   // 本类模块已见 `Class X(Of T)` 头行

    // 当前 token (lookahead 缓冲的第一个)
    Token cur_;
    // 前瞻 token
    Token next_;
    // 上一个已消费的 token (Fix 043c: 用于检测 .Member .Member 之间的空格)
    Token prevTok_;

    // With 嵌套深度 (>0 时 .member 为 WithMemberExpr)
    int withDepth_ = 0;

    // 单行 If 上下文: >0 时禁止标签检测 (colon是语句分隔符, 非标签冒号)
    int inSingleLineIf_ = 0;

    // 安全限制: advance调用计数
    size_t advanceCount_ = 0;
    static constexpr size_t MAX_ADVANCES = 10'000'000;

    // 优先级表 (初始化一次)
    std::unordered_map<int, BindingPower> bpTable_;

    void initBindingPowers();
};

// 账 #172: VB6 日期字面量折算成 OLE 自动化日期序列（epoch 1899-12-30 = 0.0）。
// 改前这里**没人算过**：AST 节点只带原文，发码侧照 doubleValue 打出去 ⇒ 读到的是没写过的
// 联合体高 4 字节（Debug 恰好 0.0、Release 是 -6.277e+66 这类垃圾，见账 #172）。
// 认得的形状：#M/D/Y#、#D-M-Y#、#M/D/Y H:N[:S][ AM|PM]#、#H:N[:S][ AM|PM]#；两年份按 VB6 规则
// （<50 → 2000s，≥50 → 1900s）。不认得就返回 false、out 不动，由调用方决定退路。
bool foldDateLiteralToOADate(const std::string& raw, double& out);

} // namespace vb6c3
