#pragma once
// 编译器驱动 - 命令行解析 + 编译流程编排

#include "preprocessor/preprocessor.hpp"
#include "ast/ast.hpp"              // 泛型 (G2): Decl/Module AST 类型
#include "semantics/symbol_table.hpp"
#include "semantics/type_system.hpp"
#include "com/typelib_parser.hpp"
#include "project/frm_parser.hpp"
#include "driver/static_lib.hpp"    // ai/024: 静态库寻址 (LibSearchPaths)
#include "backend/asm_proc.hpp"     // ai/vb-asm-extension-spec: Asm 过程降级元数据
#include "semantics/generics_registry.hpp"
#include "semantics/interfaces_registry.hpp"  // Interface 契约 (tB, B02)
#include "semantics/coclass_identity.hpp"     // CoClass 身份求解 (tB, B11/C02)
#include "semantics/class_chain_registry.hpp"  // 类继承链 (tB, B07)
#include <array>
#include <cctype>
#include <filesystem>
#include <string>
#include <vector>
#include <memory>
#include <map>
#include <unordered_map>
#include "common/encoding.hpp"

namespace vb6c3 {

class Diagnostics;
class SourceBuffer;
class Module;
class SemanticAnalyzer;
class SessionManager;

// 编译选项
struct CompileOptions {
    std::vector<std::string> sourceFiles;   // 源文件列表
    std::string outputFile;                  // 输出文件路径
    std::string outputDir;                   // 输出目录 (默认: output)

    // 目标平台
    std::string target = "win-x64";         // win-x86, win-x64, linux-x64, etc.
    std::string arch = "x64";              // x64 (default) or x86 — output binary architecture
    bool archExplicit = false;             // Task #44: 用户给过 --arch 时 --target 不再翻 arch

    // GUI模式
    std::string guiMode = "native";          // native, webview, none

    // 输出控制
    bool dumpTokens = false;
    bool dumpAST = false;
    bool dumpIR = false;
    bool dumpPreprocess = false;              // 输出预处理后的token列表
    bool dumpSymbols = false;                 // 输出符号表
    bool emitC = false;                       // 输出C代码 (.h/.c)
    bool emitLLVM = false;                   // 输出.ll文件
    bool syntaxOnly = false;                 // 只做语法检查
    bool verbose = false;

    // 条件编译
    std::vector<std::string> defines;        // -d:NAME=VALUE 或 --define NAME=VALUE

    // 优化
    int optimizationLevel = 0;               // 0=无, 1/2/3

    // 调试
    bool debugInfo = false;

    // 兼容性
    bool compatCheck = false;                // 跨平台兼容性检查

    // COM TypeLib引用 (P6.3, 前期绑定)
    std::vector<std::string> typelibRefs;    // TypeLib路径或ProgID列表
    bool autoTypelib = true;                 // 自动从源码中提取COM类型并加载TypeLib

    // ActiveX DLL (P6.6)
    bool isDll = false;                      // 编译为ActiveX DLL (而非EXE)
    std::string dllProgId;                   // DLL的ProgID前缀 (如 "MyLib")
    std::string libidStr;                    // P6.13: TypeLib的LibID (UUID格式, 空则自动生成)

    // 窗体调试 (P7)
    bool dumpFrm = false;                    // 输出.frm窗体描述解析结果
    bool extractFrx = false;                 // Fix 195: 把 .frx 设计期取值导出成 VB 代码
    bool keepTemps = false;                   // 保留中间文件 (调试用)

    // 警告抑制 (性能优化): 需要静默的诊断ID列表, 如 --no-warn 3001,3003
    std::vector<int> suppressedWarningIds;

    // 增量编译 (性能优化): 基于内容哈希的obj级缓存, 跳过未变化的.c编译
    bool incremental = false;

    // 静态库搜索根 (ai/024 E2, 批次 T01): `--libdir <dir>` 可重复。
    // 与 vbp 的 `LibDir=` 汇入同一张搜索根表; CLI 的项按 cwd 解析 (用户在命令行
    // 亲手敲的路径), 工程内的相对路径一律以 vbp 目录为基准 —— 两条基准各自诚实。
    std::vector<std::string> libDirs;

    // 附加静态库 (ai/024 M4): `--extra-lib <x>` 可重复 —— 没有对应 Declare 的链接
    // 依赖 (典型: 静态库之间互引)。寻址与静态 Declare 的 Lib 串完全同一条路。
    std::vector<std::string> extraLibs;

    // ai/023 S01: 包搜索根 (--package-root, 可重复)。
    // 与 vbp 的 Package= 同族: 相对路径以【工程目录】为基准 (E2: 不引入 cwd)。
    // 缺省根 <vbp目录>/packages 由 checkPackages 自动追加, 排在显式根之后。
    std::vector<std::string> packageRoots;

    // ai/023 S05: --check-packages 只读模式 —— 解析 Package= 引用、逐文件
    // 校验 sha1/size, 打印结果后退出 (不编译)。用于 CI 里预检包完整性。
    bool checkPackagesOnly = false;

    // 裁剪include (性能优化): 只include实际引用的外部模块, 降低cl预处理量
    bool trimIncludes = false;
};

// 编译结果
struct CompileResult {
    bool success = false;
    std::string outputFile;
    int errorCount = 0;
    int warningCount = 0;
};

// ai/vb-asm-extension-spec: 一个已发码的 Asm 过程 (过程体 = 单个 Asm 块)
//   moduleBase = 所属模块产物基名 (与 .c 同名) —— 用于命名 <base>_asm.asm
struct EmittedAsmProc {
    std::string moduleBase;
    AsmProcInfo info;
};

// 编译器驱动
class Driver {
public:
    Driver();
    ~Driver();

    // 从命令行参数编译
    CompileResult compile(int argc, char* argv[]);

    // 从选项编译
    CompileResult compile(const CompileOptions& options);

    // 解析命令行参数
    static std::pair<CompileOptions, int> parseArgs(int argc, char* argv[]);

    // 打印帮助/版本
    static void printHelp();
    static void printVersion();

private:
    std::unique_ptr<Diagnostics> diag_;
    std::vector<std::unique_ptr<Module>> modules_;  // 解析产出的AST

    // 语义分析产出 (供代码生成使用)
    std::vector<std::unique_ptr<SemanticAnalyzer>> analyzers_;

    // 窗体描述 (P7, .frm文件解析结果, 按模块名索引)
    std::map<std::string, FrmFile> frmFiles_;

    // ============================================================
    // 泛型 (tB 扩展, G2) — 单态化前端
    // ============================================================
    // 使用点登记表 (parse 期由 Parser 记录, runParser 合并到这里):
    //   key = 扁名小写 (Foo$gen1$Long), value = {模板基名, 实参扁名表}
    struct GenericUseRec { std::string base; std::vector<std::string> args; };
    std::unordered_map<std::string, GenericUseRec> genericUses_;
    // 模板登记表: lower(模板名) → 模板声明 + 宿主模块 + 类型参数名表
    struct GenericTemplateInfo {
        Module* module = nullptr;
        Decl* decl = nullptr;          // TypeDecl / SubDecl / FunctionDecl / PropertyDecl
        std::vector<std::string> typeParams;
    };
    std::unordered_map<std::string, GenericTemplateInfo> genericTemplates_;
    // 泛型类模板 (G4): lower(类名) → 模板 .cls 模块 (本体不发码, 特化=整模块克隆)
    std::unordered_map<std::string, Module*> genericClassTemplates_;
    // 模板只读视图 (analyzer 推断用; 与 genericTemplates_ 同步构建)
    GenRegistry genView_;
    // Interface 契约 (tB, B02): stage 2.7 建好的只读登记表 (小写接口名 → 展平槽表).
    // 挂 Driver 而非某个符号表: 每模块一张符号表, 而接口名是工程级唯一的.
    IfaceRegistry ifaces_;
    std::vector<std::string> ifaceOrder_;  // 登记序, 保证诊断输出确定性
    // 委托式实现 (tB, B10): stage 2.7 Pass D 建好的只读裁决表 (小写类模块名 → Via 子句).
    // 消费方 = 语义层 (据此免掉逐槽 VB3012) 与发码层 (据此转调持有对象的接口槽).
    ViaRegistry vias_;
    // 类继承 (tB, B07a): stage 2.8 建好的只读链登记表 (小写类名 → 父先己后的链).
    // 同上, 挂 Driver 而非符号表; B07b 的成员合并是唯一消费者.
    ClassChainRegistry classes_;
    std::vector<std::string> classOrder_;  // 登记序, 保证诊断输出确定性
    // ai/084a M1: 类成员访问级别预计算表 (小写类名 → 小写成员名 → 级别+定义类)。
    // runSemanticAnalysis 入口处从 AST 声明建表, 之后只读, 经 setMemberAccessTable 下发。
    // 不 visit 期查符号表的理由见 symbol_table.hpp MemberAccessTable 注释。
    MemberAccessTable memberAccessTable_;
    void buildMemberAccessTable();
    // ai/084c: 类名(小写) → Class_Initialize 形参个数 (buildMemberAccessTable 顺带扫描,
    // 经 setCtorParamCounts 下发; 表内缺席 = 非本工程类, New 带实参在语义层报错)。
    std::map<std::string, int> ctorParamCounts_;
    // 已物化扁名 (fixpoint 去重; 值为 true 即"已注入为普通声明")
    std::unordered_map<std::string, bool> genericMaterialized_;
    // cap 护栏 (计划冻结版): 总量 ≤1024, 单名嵌套深度 ≤16
    static constexpr size_t kGenericMaxInstances = 1024;
    static constexpr size_t kGenericMaxDepth = 16;

    // P6.8: VBP指定的类CLSID映射 (模块名小写 -> CLSID字符串)
    std::unordered_map<std::string, std::string> classClsidMap_;

    // CoClass 身份 (tB 扩展, ai/022 D46, 批次 B11/C02): stage 2.7 Pass E 求解一次、之后只读。
    // key = CoClass 块名小写。消费者 = 将来的 C05（`New`/`CreateObject` 编译期改写）与
    // B13/B15/B16（对外那半）；C02 只发一条 note 让它可测。
    std::unordered_map<std::string, CoClassIdentity> coclassIds_;
    // 接口 IID 表 (tB, ai/022 B13c): 同一个 Pass E 里由 `buildIfaceIdMap` 建，key = 接口名小写。
    // 三处消费者（新式接口 vtable 的 QI、dll_entry 的 IID 表、类型库）从此只读这一张表 ——
    // "一个接口在一次编译里只许一枚 GUID" 的落地形式。
    std::unordered_map<std::string, std::string> ifaceIds_;
    // vbp 的 Name= 字段（ProgID 默认值与确定性 mint 的 <Proj>，与 projectBaseName_ 分叉，见 D46）
    std::string vbpProjectName_;

    // VBP工程基名 (用于多模块工程的输出文件命名)
    std::string projectBaseName_;
    std::string projectPath32_;    // P11.1: VBP Path32 field (output dir)

    // 静态库寻址 (ai/024, 批次 T01) —— 阶段0 收集搜索根, 阶段2.x 校验 Declare 引用。
    // 基准目录 = vbp 所在目录 (单文件模式 = 源文件所在目录); 绝不引入 cwd 基准。
    LibSearchPaths staticLibPaths_;
    // 已解析成功的静态库: 原始 Lib 串 → 绝对路径 (T02 发码时按此组装链接输入)
    std::map<std::string, std::string> staticLibResolved_;
    // 附加静态库的原始串 (vbp 的 ExtraLib= ++ CLI 的 --extra-lib), 按书写顺序。
    // 与 Declare 的 Lib 串同一条寻址路径, 因此也共用同一批诊断。
    std::vector<std::string> extraLibRaw_;
    // ai/024 T02: ExtraLib 的链接输入, 与 extraLibRaw_ **同序**。元素可能是
    //   绝对路径 (工程内找到) 或 裸名 (没找到但后缀是 .lib/.a → 放行给链接器,
    //   靠 LIB 环境变量 / /LIBPATH 找)。driver_link.cpp 原样追加到链接命令行。
    std::vector<std::string> extraLibResolved_;

    // ai/024 E4 (批次 T04b): `Alias "_foo@12"` 逃生舱的 /alternatename 指令表。
    // 发码侧对 '@' 只能清洗成 '_' (非合法 C 标识符), 修饰后的内部名与真实符号
    // 对不上 → 链接命令行追加 /alternatename:<修饰内部名>=<真实符号> 桥接。
    // 元素形如 "internal=real"; 由 driver_staticlib.cpp 的 alternatenameDirective()
    // 生成 (发码侧 sanitizer 的镜像, 见该函数注释), driver_link.cpp 去重后交给
    // MsvcDriverOptions::alternatenames。
    std::vector<std::string> staticLibAlternatenames_;

    // Fix 142: VBP 的 Startup= 启动对象 ("Sub Main" 或窗体模块名).
    // 决定多模块工程中哪个模块生成进程入口点 (WinMain/main).
    std::string startupObject_;

    // ai/vb-asm-extension-spec: Asm 块过程的降级元数据 (发码期收集, 链接近前消费)。
    //   moduleBase = 该过程所属模块的产物基名 (与 .c 同名, 用于命名 <base>_asm.asm);
    //   runCodeGeneration 逐模块从 CCodeGen::asmProcs() 回收, runLinker 生成 .asm
    //   并调 ml64 → .obj → MsvcDriverOptions::extraObjects。
    std::vector<EmittedAsmProc> asmProcs_;

    // Fix 143: vbp Object= 的 OCX 文件表 (CLSID 小写去花括号 → ocx 绝对路径).
    // 第三方 OCX 控件免注册加载用.
    std::map<std::string, std::string> ocxFiles_;

    // Fix 143b: vbp Object= 原始引用列表 (CLSID 小写去花括号, ocx 绝对路径).
    // 控件实例化 CLSID 以此为准 (typelib coclass GUID ≠ 实例 CLSID).
    std::vector<std::pair<std::string, std::string>> ocxRefs_;

    // Fix <vbeclipse> 2026-10-06: vbp Object= 免注册 OCX 表 (canonical 小写绝对路径 → 相对 exe 路径).
    // driver_compile 填, runTypeLibImport 按 TypeLibResult 路径反查:
    // 命中者其全部 coclass 进免注册表, 未命中者 (普通 Reference= / auto-typelib) 不进,
    // 运行期对未声明 OCX 零变化. 路径语义与 ComLib= 一致.
    std::unordered_map<std::string, std::string> ocxCanonMap_;

    // Fix <vbeclipse> 2026-10-06: Object= OCX 组件表 {ProgID, CLSID, coclass名, 相对exe路径},
    // runTypeLibImport 收集, cgen 烘焙进产物入口点 (vb6_OcxRefRegister). 运行期
    // Me.Controls.Add 按 ProgID 命中后走 ocxCreateAny 免注册.
    std::vector<std::array<std::string, 4>> ocxLibRefs_;

    // Fix 160: vbp ComLib= 声明的组件 DLL (canonical 小写绝对路径 → 相对 exe 路径).
    // driver_compile 填, runTypeLibImport 按 TypeLibResult 路径反查:
    // 命中者其全部 coclass 进免注册表, 未命中者 (普通 Reference= / auto-typelib) 不进,
    // 运行期对未声明组件零变化.
    std::unordered_map<std::string, std::string> comLibCanonMap_;

    // Fix 160: ComLib= 组件表 {ProgID, CLSID, coclass名, 相对exe路径},
    // runTypeLibImport 收集, cgen 烘焙进产物入口点 (vb6_ComLibRegister).
    std::vector<std::array<std::string, 4>> comLibRefs_;

    // === ai/023 S03: 包导出边界 ===
    // S02 加载时 (driver_compile) 填充, S03 跨模块注入处 (driver_crossmod) 消费。
    struct PackageExportInfo {
        bool friendVisible = false;                  // 清单 [Export] Friend=True
        std::map<std::string, bool> exportedModules; // 小写 Module/Class 名 (VB_Name)
    };
    std::map<std::string, PackageExportInfo> packageExportInfos_; // key: 小写包名
    // 源文件 → 所属包 (key: normSourceKey(绝对路径); value: 小写包名)。
    // frontend 建 Module 时回填 module->packageName。
    std::map<std::string, std::string> packageOfFile_;
    // ai/023 S03: 预计算的包内被屏蔽成员表 (小写包名 → 小写成员名 → 所在模块名)。
    // driver_compile 在 S02 读包源码时用 extractProcDecls 扫描得到:
    //   非导出模块的全部过程 / 导出模块的 Friend 过程 (清单 Friend=True 时豁免)。
    // Private 过程本就不导出 (getPublicSymbols 过滤), 一并列进只为诊断完整。
    // runSemanticAnalysis 按消费者的包归属下发给各分析器, visit 期查无此名时
    // 命中即报 VB7006 —— 无实参裸调用不走 deferred 通路, 必须在 visit 期拦。
    std::map<std::string, std::map<std::string, std::string>> packageBlockedNames_;

    // ai/023 S04: 预计算的包内被屏蔽**类名**表 (小写包名 → 小写类名 → 类名)。
    // 类不走 S03 的符号注入通路 —— 消费方 codegen 靠 lookupTypeSymbol 解析,
    // 解析不到就**静默退化成后期绑定 COM** (`vb6_NewObject(L"HiddenCls")`)，
    // 比 S03 的 LNK2019 更坏 (编译通过、运行期才炸)。故非导出类必须在语义层
    // 硬拦: visit(NewExpr) 与类型引用 (Dim ... As Cls) 两处报 VB7006。
    std::map<std::string, std::map<std::string, std::string>> packageBlockedClasses_;

    // 源文件路径归一键: 弱规范化 + 小写 (Windows 同款)。S02 登记与 frontend
    // 查询两侧共用, 保证同一文件必然命中同键。
    static std::string normSourceKey(const std::string& utf8Path) {
        std::error_code ec;
        std::filesystem::path p = std::filesystem::weakly_canonical(utf8ToPath(utf8Path), ec);
        if (ec || p.empty()) p = utf8ToPath(utf8Path).lexically_normal();
        std::string s = pathToUtf8(p);
        for (auto& c : s) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
        return s;
    }

    // P23-05: VBP version info (for VS_VERSION_INFO resource)
    int verMajor_ = 1;
    int verMinor_ = 0;
    int verRevision_ = 0;
    std::string verCompanyName_;
    std::string verFileDescription_;
    std::string verLegalCopyright_;
    std::string verProductName_;
    std::string verComments_;
    std::string verLegalTrademarks_;
    std::string verOriginalFileName_;
    std::string verTitle_;
    std::string userResFile_;         // P23-03: VBP ResFile= .res path (absolute)

    // TypeLib解析器 (P6.3, COM类型导入)
    std::unique_ptr<TypeLibParser> typelibParser_;

    // 编译流水线各阶段
    bool runLexer(const CompileOptions& options);
    bool runPreprocess(const CompileOptions& options);
    bool runParser(const CompileOptions& options);
    bool runTypeLibImport(const CompileOptions& options);  // P6.3: 加载TypeLib+注册COM类型
    bool runSemanticAnalysis(const CompileOptions& options);
    bool runGenericsPrepass();  // 泛型 (tB): 模板登记 + 使用点物化 (G2)
    // Interface (tB, B02): stage 2.7 建接口契约登记表 (名字/Extends 链/展平槽表)
    // options 只被 Pass F 用一项: EXE 工程不能注册为 COM 服务器 (ai/022 D48, B11/C03a).
    bool runInterfacePrepass(const CompileOptions& options);
    // 类继承 (tB, B07a): stage 2.8 建类继承链登记表 (基名解析/环/深度/v1 边界)
    bool runClassChainPrepass();
    // 虚方法 (tB, B08b): 2.8 内两步 —— 位置合法性 (只读 modules_) / 覆盖契约与 dynamicKeys (要链)
    void checkVirtualPlacement();
    void runVirtualContractChecks();
    // 类继承 (tB, B07b): stage 3.4 把祖先自有成员并进派生类 Class 符号 (必须早于 3.5 的逐字段跨模块拷贝)
    bool mergeInheritedMembers();
    // 类虚表 (tB, B08d): stage 3.4b 排每类的有序槽表 (要读 3.4 的 inhProcs 判"本类有无入口")
    bool buildVirtualSlotTables();
    // CoClass 契约聚合 (tB, ai/022 D50, 批次 B11/C03b): stage 3.4c 判"块里列出的每个接口,
    // [Implementation] 那个类**含祖先**是否满足". 必须在 3.4 之后: 祖先自有的成员只有链表
    // (stage 2.8) 能给全, 而链表的消费序与成员合并同源. 工程无 CoClass 块时立即返回 true.
    bool runCoClassContractCheck();
    // fixpoint 单轮物化: 消费 genericUses_ 中未物化项; freshOut 收特化副本
    bool materializeGenerics(std::vector<std::pair<Module*, Decl*>>* freshOut);
    // 泛型推断 fixpoint (G3): 收请求→物化→增量分析→再跨模块, 至收敛
    bool runGenericsFixpoint();
    bool runCrossModuleResolution();  // 跨模块符号链接
    // AddressOf 取址的过程标成回调 (账 #184): stage 3.5a, 必须在跨模块链接之后
    // (要读 3.5 灌好的 isExternal/sourceModule 才认得归属模块)。
    void markAddressOfCallbacks();
    // 静态库引用校验 (ai/024 批次 T01): 遍历全部 Declare, 对静态形态的 Lib 串做
    // 寻址 + 后端格式匹配, 失败即报错。**本批次不发码** (T02 才接 codegen)。
    bool validateStaticLibRefs(const CompileOptions& options);
    bool runCodeGeneration(const CompileOptions& options, const std::string& outputDir);
    void writeErrorLog(const std::string& logPath, const std::string& stage);
    bool runLinker(const CompileOptions& options, const std::string& outputDir, const std::string& intermediatesDir, SessionManager& session);
};

} // namespace vb6c3
