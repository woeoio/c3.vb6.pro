#pragma once
// 诊断系统 - 错误/警告/信息报告

#include <string>
#include <string_view>
#include <vector>
#include <sstream>
#include <cstdint>
#include <set>

namespace vb6c3 {

// 源码位置
struct SourceLocation {
    std::string filename;
    uint32_t line = 1;
    uint32_t column = 1;

    std::string toString() const {
        return filename + "(" + std::to_string(line) + "," + std::to_string(column) + ")";
    }
};

// 诊断级别
enum class DiagnosticLevel : uint8_t {
    Note,
    Warning,
    Error,
    Fatal,
};

// 诊断ID编号空间
// 1xxx: 词法阶段
// 2xxx: 语法阶段
// 3xxx: 语义阶段
// 4xxx: 代码生成阶段
// 5xxx: RTL/链接阶段
enum class DiagnosticID : uint16_t {
    // 词法 (1xxx)
    LexUnexpectedChar = 1001,
    LexUnterminatedString = 1002,
    LexInvalidNumber = 1003,
    LexInvalidCharLiteral = 1004,
    LexUnrecognizedToken = 1005,
    LexFileEncodingError = 1006,
    // C3 扩展: 反引号原始多行串 (ai/028 V1)。文案 = ASCII (D12 硬约束)。
    LexUnterminatedRawString = 1007,   // `...` 里找不到闭合的那枚反引号
    // C3 扩展: 串内插值 ${expr} / ${expr:fmt} (ai/028 V2)。这两条也在词法层 ——
    // 插值整串在出口就展开成普通 token, 后面的语法/语义问题由既有的诊断负责 (R2/R4)。
    LexUnterminatedInterp = 1008,      // 孔没等到闭合的 '}' (含"孔跨行"这种写法)
    LexEmptyInterpHole = 1009,         // ${} / ${:fmt} —— 不静默当文本

    // 语法 (2xxx)
    ParseExpectedToken = 2001,
    ParseUnexpectedToken = 2002,
    ParseExpectedEndOfStatement = 2003,
    ParseExpectedExpression = 2004,
    ParseMismatchedBlock = 2005,
    ParseInvalidIf = 2006,
    ParseInvalidFor = 2007,
    ParseInvalidSelect = 2008,
    ParseDuplicateLabel = 2009,
    ParseUndeclaredLabel = 2010,
    ParseInvalidInterfaceMember = 2011,  // Interface 块内非法成员 (实现体/字段/可见性/事件)
    ParseUnknownAttribute = 2012,        // 无法识别的 [Xxx] 属性行
    ParseAsmBlockMalformed = 2013,       // asm扩展: Asm 块形式非法 (非块形式 / 缺 End Asm)
    ParseAsmClobberMalformed = 2014,     // asm扩展: Clobber(...) 参数非法 (非字符串 / 空表)

    // 语义 (3xxx)
    SemUndeclaredIdentifier = 3001,
    SemDuplicateDeclaration = 3002,
    SemTypeMismatch = 3003,
    SemMissingDefaultProperty = 3004,
    SemWrongNumberOfArguments = 3005,
    SemNamedArgNotFound = 3006,
    SemDuplicateNamedArg = 3007,
    SemOptionalParamAfterRequired = 3008,
    SemParamArrayMustBeLast = 3009,
    SemCantAssignToReadOnly = 3010,
    SemInvalidUseOfMe = 3011,
    SemInterfaceNotImplemented = 3012,
    SemCircularDependency = 3013,
    SemVariantOverflow = 3014,
    // Interface 契约 (tB 扩展, ai/022 B02). 文案一律 ASCII (D12).
    SemInterfaceUnknownParent = 3015,      // Extends 的父接口不存在
    SemInterfaceSlotConflict = 3016,       // 链上槽名冲突 / 接口内同名重载
    SemInterfaceSignatureMismatch = 3017,  // 实现成员签名与接口槽不符
    SemInterfaceNotSupported = 3018,       // v1 边界: 该处的 Interface 用法尚不支持
    SemInterfaceClauseUnbound = 3019,      // 成员级 Implements 子句没被任何契约比对接纳 (B02b)
    // 类继承 (tB 扩展, ai/022 D6, 批次 B07). 文案一律 ASCII (D12).
    SemInheritsUnknownBase = 3020,    // Inherits 的基名不是本工程内的类
    SemInheritsTooDeep = 3021,        // 继承链长度超上限
    SemInheritsNotSupported = 3022,   // v1 边界: 非类模块 / 泛型模板内 / 多条 Inherits
    // tB 类继承 (ai/022 D29-1, 批次 B08c): 家族外访问 Protected 成员。
    SemProtectedOutsideFamily = 3023,   // obj.<Protected 成员> 的接收者类不在当前类的家族里
    // 虚方法 (tB 扩展, ai/022 D30, 批次 B08b). 文案一律 ASCII (D12).
    SemOverrideTargetUnknown = 3024,    // Overrides 找不到同名的祖先可覆盖成员
    SemOverrideNotOverridable = 3025,   // 祖先成员存在但未标 Overridable (或显式 NotOverridable)
    SemOverrideSignatureMismatch = 3026, // Overrides 与祖先槽签名不符
    SemVirtualNotSupported = 3027,      // v1 边界: 该处的虚成员用法尚不支持 (含"需要类虚表")
    // tB 类继承 (ai/022 B09): `MyBase.<成员>` 显式基调用的可用性判定。
    SemMyBaseNotSupported = 3028,       // 无基类 / 基面上没有这个成员 / 基成员是 Private
    // tB 委托式实现 (ai/022 B10): `Implements I Via m_holder` 的可用性判定。
    SemViaTargetUnknown = 3029,         // Via 目标不是本类的对象持有字段 / 接口名不是新式 Interface
    SemViaHolderNotImplemented = 3030,  // 字段类型那个类没有实现被委托的接口 (v1 不接受再往下委托)
    // tB CoClass 块 (ai/022 D48, 批次 B11/C03a): 形状与名字校验. 文案一律 ASCII (D12).
    SemCoClassEntryInvalid = 3031,      // 契约条目: 引用不存在的接口 / 条目重复 / [Default] 标了多条
    SemCoClassDuplicate = 3032,         // 块名撞车: 重复块名 / 撞模块名 / 撞接口名
    SemCoClassNotSupported = 3033,      // v1 边界: [Implementation] 不是类模块 / EXE 工程 ComCreatable(True)
    // 084a/084c/asm 扩展 (合并期与 022 B09-B11 撞号, 顺延到 3034-3037)
    SemPrivateOutsideClass = 3034,      // 084a: 类外经 obj./Me. 访问 Private 成员 (接收者不在定义类家族内)
    SemCtorArityMismatch = 3035,        // 084c: New Cls(args) 实参个数与 Class_Initialize 形参不符
    SemAsmFormUnsupported = 3036,       // asm扩展: 该形态不支持. 当前唯一来源 = x86 `<Naked>` 过程里
                                        // 按名引用参数 (naked 无栈帧, 参数在调用者的栈上, 名字无从解析)
    SemAsmMixedBody = 3037,             // asm扩展: 类方法里的 Asm 块 (v1/v2 只支持标准模块过程)
    SemAsmOperandWidthMismatch = 3038,  // asm扩展: Asm 指令两个寄存器操作数宽度不一致 (原本要到 ml64 才报 A2022)
    // CoClass 组内激活 (tB 扩展, ai/026 五-3, ai/022 D54, B11/C05): 把一个**没有
    // [Implementation]** 的 CoClass 块名写在类型位置上。今天这形状静默当 Variant
    // (D54 实测 M1/M5: `void* a = 0;` + 晚绑定 DISPID 调用), 而"块没有实现类"编译器已经知道。
    SemCoClassTypeUnbound = 3039,
    // 混排 / 累加器别名 (asm 扩展, 合并期再与 022 撞号 3039, 顺延到 3040-3042)
    SemAsmMixedRefUnresolved = 3040,    // 混排 (项3): Asm 片段里 [X] 的 X 既不是寄存器也不是可见的 VB 变量
    SemAsmMixedRefLimit = 3041,         // 混排 (项3, x64): 单个片段引用的 VB 变量超过 4 个 (Win64 只有 4 个整型参数寄存器)
    SemAsmAccumAliasClobber = 3042,     // 项1: Asm 块里某条指令的隐含累加器寄存器 (cmpxchg/div/mul 等) 被同块
                                        //      前面的指令写坏, 且写坏前的值已无从恢复 (静态可达性启发式, 见 asm_proc.hpp)
    // 数组槽实参形态 (<vbeclipse>, 覆盖 UBound/LBound/Join 的首参): VB6 要求它是数组表达式。传字面量或
    // 非数组常量 (`UBound(vbNull)`, vbNull=1) 在 VB6 是编译期错误, 而 C3 此前把常量
    // 折成整数直接塞进 vb6_UBound 的指针形参 → 运行期解引用地址 1 → 0xC0000005。
    // 文案 ASCII (D12)。
    SemArrayArgExpected = 3043,
    // 账 #278: Implements 的两条契约诊断原本借用 VB3001 (未声明的标识符), 而 VB3001 是
    // 「一个名字解析不出来」那一族; 这里两条报的都是「接口/槽」, 主语不是标识符使用点。
    // 分家出独立号, 才可以让 VB3001 的严重级单独动 (见 §B102 ③)。文案 ASCII (D12)。
    // 缺槽那条沿用既有的 SemInterfaceNotImplemented = 3012 (新式 Interface 那条路已在用)。
    SemImplementsInterfaceNotFound = 3044,  // Implements 的接口名在本工程内找不到

    // 代码生成 (4xxx)
    CodeGenUnsupportedFeature = 4001,
    CodeGenLLVMError = 4002,
    CodeGenLinkerError = 4003,
    // Fix 195: 窗体引用的二进制资源 (.frx/.ctx/.pgx) 缺失或读不出。
    // VB6 把"无法用文本行表达的属性值"(多行文本 / 图片 / 长字符串) 存进同名 .frx,
    // 属性行只留 `Text = "Form1.frx":0000` 的偏移引用。资源缺失时这些属性的设计期
    // 取值会被**静默丢弃** (VB6 IDE 自己会写 <窗体>.log 报"文件引用无效"), 于是编出来
    // 的 exe 与 VB6 不一致却毫无提示 —— 本号就是补这条提示。
    CodeGenFormResourceMissing = 4004,

    // RTL/链接 (5xxx)
    LinkUnresolvedExternal = 5001,
    LinkDuplicateSymbol = 5002,
    LinkMissingRTL = 5003,
    // 静态库链接 (ai/024, 批次 T01). 文案一律 ASCII (沿用 022/023 纪律)。
    LinkStaticLibPath = 5004,      // 静态库文件在全部搜索根下都找不到
    LinkStaticLibFormat = 5005,    // 后端 × 归档格式不匹配 (MSVC 吃 .lib/.obj, MinGW 吃 .a/.o)
    LinkStaticLibParam = 5006,     // 静态 Declare 的参数形态不受支持 (x86 ByVal Variant, §六-8)
    LinkStaticLibAmbiguous = 5007, // 归档里多个符号与声明归一化后同名 (§五 L2)
    LinkStaticLibSymbol = 5008,    // 声明算出的引用名与归档里唯一候选不一致 (§五 L2)
    // 预处理 (6xxx)
    PreprocUndefinedConstant = 6001,
    PreprocInvalidDirective = 6002,
    PreprocConstRedefinition = 6003,

    // 工程引用/包 (ai/023 S01). 文案一律 ASCII (沿用 022/024 纪律)。
    VbpPackageNotFound = 7001,    // 包目录或 package.c3d 在全部搜索根下找不到
    VbpPackageEscape = 7002,      // 包名/版本含路径字符 (解析结果跳出搜索根)
    VbpPackageFormat = 7003,      // 清单不合法 (未知段/键、Format≠1、缺必填)
    VbpPackageConflict = 7004,    // 包名与宿主工程名冲突 (模块名冲突在 S03 符号层)
    VbpPackageFileMissing = 7005, // 清单列的文件不存在 (警告, D7 校验不拒收; 哈希比对 S05)
    VbpPackageNotExported = 7006, // 引用了包内未导出的成员 (Friend/Private 边界, ai/023 S03)
    VbpPackageHashMismatch = 7007, // 包文件 sha1/size 与清单不符 (警告, D7 校验不拒收; ai/023 S05)
    VbpPackageFriendMember = 7008, // 084a M3: 经 obj. 访问包导出类的 Friend 成员 (清单未开 Friend=True)
};

// 单条诊断信息
struct Diagnostic {
    DiagnosticLevel level;
    DiagnosticID id;
    SourceLocation location;
    std::string message;
    std::vector<SourceLocation> related; // 相关位置(如: 声明处)

    std::string toString() const {
        const char* levelStr = "";
        switch (level) {
            case DiagnosticLevel::Note:    levelStr = "note"; break;
            case DiagnosticLevel::Warning: levelStr = "warning"; break;
            case DiagnosticLevel::Error:   levelStr = "error"; break;
            case DiagnosticLevel::Fatal:   levelStr = "fatal error"; break;
        }
        std::string result = location.toString() + ": " + levelStr;
        result += " VB" + std::to_string(static_cast<uint16_t>(id)) + ": " + message;
        return result;
    }
};

// 诊断收集器
class Diagnostics {
public:
    void note(DiagnosticID id, const SourceLocation& loc, const std::string& msg) {
        diagnostics_.push_back({DiagnosticLevel::Note, id, loc, msg, {}});
    }

    void warn(DiagnosticID id, const SourceLocation& loc, const std::string& msg) {
        if (isSuppressed(id)) return;
        diagnostics_.push_back({DiagnosticLevel::Warning, id, loc, msg, {}});
        warningCount_++;
    }

    void error(DiagnosticID id, const SourceLocation& loc, const std::string& msg) {
        diagnostics_.push_back({DiagnosticLevel::Error, id, loc, msg, {}});
        errorCount_++;
    }

    void fatal(DiagnosticID id, const SourceLocation& loc, const std::string& msg) {
        diagnostics_.push_back({DiagnosticLevel::Fatal, id, loc, msg, {}});
        errorCount_++;
    }

    void addRelated(const SourceLocation& loc) {
        if (!diagnostics_.empty()) {
            diagnostics_.back().related.push_back(loc);
        }
    }

    // 查询
    bool hasErrors() const { return errorCount_ > 0; }
    bool hasWarnings() const { return warningCount_ > 0; }
    int errorCount() const { return errorCount_; }
    int warningCount() const { return warningCount_; }
    const std::vector<Diagnostic>& all() const { return diagnostics_; }

    // 输出所有诊断
    std::string toString() const {
        std::string result;
        for (const auto& d : diagnostics_) {
            result += d.toString() + "\n";
        }
        return result;
    }

    // 清空
    void clear() {
        diagnostics_.clear();
        errorCount_ = 0;
        warningCount_ = 0;
    }

    // 抑制指定ID的警告 (性能优化: 大型项目如vbman会打印上千条
    // VB3001/VB3003 宽松模式警告, 抑制后可减少日志I/O与输出膨胀)
    void suppress(DiagnosticID id) { suppressed_.insert(id); }
    bool isSuppressed(DiagnosticID id) const { return suppressed_.count(id) > 0; }

private:
    std::vector<Diagnostic> diagnostics_;
    std::set<DiagnosticID> suppressed_;
    int errorCount_ = 0;
    int warningCount_ = 0;
};

} // namespace vb6c3
