#include "backend/cgen.hpp"
#include "driver/static_lib.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_decl_api.cpp: 外部 API 声明生成（DeclareDecl） ---


void CCodeGen::visit(DeclareDecl& node) {
    // 外部函数声明 (Declare Sub/Function ... Lib "xxx" [Alias "yyy"] [CDecl])
    // 去除字符串两端引号 (词法器保留引号)
    auto stripQuotes = [](const std::string& s) -> std::string {
        if (s.size() >= 2 && s.front() == '"' && s.back() == '"')
            return s.substr(1, s.size() - 2);
        return s;
    };

    // Lib名: 去引号、去.dll后缀
    std::string libName = stripQuotes(node.libName);
    if (libName.size() > 4 &&
        (libName.compare(libName.size()-4, 4, ".dll") == 0 ||
         libName.compare(libName.size()-4, 4, ".DLL") == 0)) {
        libName = libName.substr(0, libName.size()-4);
    }

    // === ai/024 T02: 静态库分支 ===
    // `.lib/.obj` 结尾的 Lib 串是**归档** (项目自带), 不是 Win32 DLL 导入库。与动态路
    // 的三处差别 (其余一字不改):
    //   1) 声明直连**真实导出符号** `MyAdd`, 而不是 `vb6_di_MyAdd` 转发桩 ——
    //      桩是给 Win32 DLL 家族准备的 (实现由 scripts/gen_di_stubs.ps1 从生成头的
    //      `/* vb6_di_lib: X */` 标记产出), 用户自己的归档里没有 `vb6_di_MyAdd`,
    //      沿用必然 LNK2019。直连的另一个好处: 引用符号由编译器按 callConv 生成
    //      (x86 stdcall 自动加 `@N`), 正是 024 §五 L0 的口径 —— 装饰名靠编译器算,
    //      不靠 C3 手写尺寸表。
    //   2) **不** emit `#pragma comment(lib,...)`: 库路径由 driver 从 LibSearchPaths
    //      解析成绝对路径后直接进链接命令行 (driver_link.cpp 的 extraLibInputs)。
    //      既避免路径里的 `\` 落进 C 字面量, 也避免今天那条
    //      `libForLink + ".lib"` 盲拼在 `mylib.lib` 上拼出 `mylib.lib.lib`
    //      (实测 LNK1104)。
    //   3) **不** emit `/* vb6_di_lib: ... */` 家族标记: 那是给桩生成器认家族用的,
    //      静态库不该被生成器当成一个新的 DLL 家族。
    // 动态路 (裸名 / .dll) 在这里 isStaticDecl=false, 下面每一处都退化成原逻辑。
    const bool isStaticDecl = (classifyLib(libName) != LibKind::Dynamic);

    // Fix 081e: Declare函数返回Long在x64下应映射为intptr_t
    // VB6 Long (32-bit) 在Declare中常用于返回句柄/指针 (如CreateEnhMetaFileW返回HDC),
    // 在x64下需要intptr_t (8字节) 才能容纳指针值
    std::string retType = (node.procKind == ProcKind::Function)
        ? mapDeclareType(node.returnType.get()) : "void";

    std::string params = makeParamList(node.params, true, isStaticDecl);

    // 调用约定
    std::string callConv = (node.callingConv == CallConv::CDecl) ? "__cdecl" : "__stdcall";

    // Alias: 去引号, 保持原始导出名 (大小写敏感, 可能含#序号前缀)
    std::string aliasName = stripQuotes(node.aliasName);

    // VB6函数名→C标识符 (用于调用点)
    std::string cFuncIdent = cIdent(node.name);

    // 导出名: Alias优先, 否则用VB6函数名
    // 注意: Windows API函数名是大小写敏感的, 需要保持原始大小写
    std::string exportedName = aliasName.empty() ? node.name : aliasName;

    // Fix 187: VB6 对 Declare 的 `ByVal <x> As String` **一律按 ANSI 编组**，与 API
    // 名字是否以 'A' 结尾无关（VB6 没有宽字符编组，`As String` 出参永远是 LPSTR）。
    // 此前只有名字以 A 结尾的才登记，于是 `GetProcAddress(h, "CallWindowProcA")` 这类
    // 无后缀 API 收到的是裸 BSTR（UTF-16），名字查找必然失败返回 0 ——
    // Charts2020 的 ucChartArea/ucChartBar/ucPieChart/ucTreeMaps/ucProgressCircular/
    // LabelPlus 自造子类化 thunk 时把 0 写进 CallWindowProcA 槽位，关窗时 user32 调该
    // thunk → `call 0` → 执行违例 0xC0000005（退出码 0xC000041D，表现为关闭转圈 ~2.7s）。
    // 调用点的转换在 cgen_expr_call_arg_emit.inc（isDeclareAnsiCall 分支）。
    // ai/024: `DeclareWide` —— 整个 ANSI 编组机制只由 knownDeclareAnsi_ 这一个集合
    // 驱动 (消费点: cgen_expr_call_callee_params.inc 的 isDeclareAnsiCall 判定,
    // 以及 cgen_expr_call_arg_emit.inc 的 vb6_BSTR_ToANSI 生成)。所以"禁用
    // ANSI<->Unicode 转换"的全部实现 = **不登记**。调用点随即把 BSTR 原样传给
    // callee (BSTR 就是宽字符指针), 无需 StrPtr。
    // 宽声明对同名函数有优先权: 若先出现普通 Declare 已登记, 这里撤掉登记。
    // (同名同时出现宽/非宽两种声明属未定义用法, 仅保证"宽的那条最终生效"。)
    {
        std::string funcLower = node.name;
        std::transform(funcLower.begin(), funcLower.end(), funcLower.begin(), ::tolower);
        if (node.isWide) {
            knownDeclareAnsi_.erase(funcLower);
        } else {
            knownDeclareAnsi_.insert(funcLower);
        }
        // Fix 161b-decl-out: Alias→调用点 的登记推迟到下方 cExportedIdent 定址之后
        // (映射里要存 vb6_di_<清洗名> 全名, 而不是裸 Alias)。
    }

    // Fix 092z-2: VB6/VBA 运行时库 (msvbvm60 等) 不生成 #pragma comment(lib, ...)
    //
    // 原因:
    //  1. 这些库只随 VB6/VBA 发行, Windows SDK / VS 均不携带 → 链接器报 fatal
    //     LNK1104 "无法打开文件 msvbvm60.lib" (本机 SysWOW64 只有 32 位
    //     msvbvm60.dll, x64 无导入库可链); VB6 运行时也只有 32 位。
    //  2. Fix 076 之后 Declare 一律生成 `extern <ret> __stdcall vb6_di_<name>(...)`
    //     + `#define <VB名> vb6_di_<name>`, 由 RTL 转发桩提供实现 —— 导入库这条路
    //     对 Declare 早已废弃。且生成物引用的符号名是 C3 内部名, 即使补上真库也
    //     解析不了 (MSVBVM60.DLL 导出的是 VarPtr / __vbaObjSetAddref / 序号)。
    //  3. 这些符号的实现见 src/rtl/core/vb6_di_stubs.c (原生实现, 不做 LoadLibrary 转发)。
    std::string libLower = libName;
    std::transform(libLower.begin(), libLower.end(), libLower.begin(), ::tolower);
    bool isVb6RuntimeLib = (libLower == "msvbvm60" || libLower == "msvbvm50" ||
                            libLower == "vbe7" || libLower == "vbe6" ||
                            libLower == "vba7" || libLower == "vba6");
    // 另一类: Windows SDK 不提供导入库的 DLL。实测 cryptdlg.dll 只有 DLL,
    // SDK 10.0.26100.0\um\x64 无 cryptdlg.lib → 同样不生成 pragma, 对应符号
    // 由 RTL 动态加载实现 (LoadLibrary + GetProcAddress), 见 vb6_di_stubs.c
    // 的 vb6_di_CertSelectCertificateW。
    bool isNoImportLib = (libLower == "cryptdlg");
    // hhctrl.ocx 是 HTML Help ActiveX 控件 (HHCTRL), 其导入库只随 32 位
    // 发行 (Windows SDK x64 不提供 hhctrl.ocx.lib) → 生成
    // #pragma comment(lib, ...) 会 LNK1104 "无法打开文件". 实际符号由
    // vb6_di_ 转发桩实现 (vb6_di_stubs.c), 无需导入库 → 不生成 pragma.
    if (libLower == "hhctrl.ocx") isNoImportLib = true;
    // olepro32.dll 是 32 位遗留库, SDK 只在 x86 目录提供 olepro32.lib (x64 SDK 无);
    // 而 OleCreatePictureIndirect / OleLoadPicture 等实际由 oleaut32.dll 在 x86/x64 上导出,
    // 故把 olepro32 的链接映射到 oleaut32, 保证 32/64 位都能链接。
    std::string libForLink = libName;
    if (libLower == "olepro32") {
        libForLink = "oleaut32";
    }
    // Task #44 (SSTabEx): winspool.drv 的 DLL 文件名带 .drv, 但 SDK 导入库叫
    // winspool.lib (um/x86、um/x64 均无 winspool.drv.lib) → 按 DLL 名拼 .lib
    // 会 LNK1104 "无法打开文件 winspool.drv.lib"。映射到 SDK 库名 (同 olepro32)。
    if (libLower == "winspool.drv") {
        libForLink = "winspool";
    }
    // ai/024 T02: 静态路**不**发 pragma (见上方 isStaticDecl 注释之二)。
    // olepro32→oleaut32 的重映射只对动态导入库有意义, 静态路自然也不适用。
    if (!isVb6RuntimeLib && !isNoImportLib && !isStaticDecl) {
        c_.emitLine("#pragma comment(lib, \"" + libForLink + ".lib\")");
    }

    // Fix 010a: 避免与Windows SDK (windows.h) 声明冲突
    //
    // 问题: vb6rtl.h 已 #include <windows.h>, 即所有Windows API函数已被声明。
    // C3生成的 __declspec(dllimport) 声明与SDK声明签名不同 (如 int32_t vs HANDLE/void*),
    // 导致 C2371 "redefinition; different basic types" 等错误。
    //
    // 解决方案: 使用C3内部唯一名称 vb6_di_<ExportedName> 作为 __declspec(dllimport) 的函数名,
    // 然后用 #define 将VB6函数名映射到该内部名称。
    //
    // 这样:
    // 1. SDK已#define的宏 (如 CopyMemory → RtlMoveMemory → memmove):
    //    #ifndef CopyMemory 为false → 不生成C3的#define → 调用点使用SDK的宏展开 → 正确
    //    (vb6_di_CopyMemory 声明存在但永远不会被调用 → 无害)
    // 2. SDK已声明为函数 (如 GetCurrentProcess):
    //    #ifndef GetCurrentProcess 为true → 生成 #define GetCurrentProcess vb6_di_GetCurrentProcess
    //    → 调用点 GetCurrentProcess() 被宏展开为 vb6_di_GetCurrentProcess() → 使用C3的导入版本
    //    (SDK的 GetCurrentProcess 声明仍在, 但不会被调用 → 无冲突, 因为名字不同)
    // 3. SDK未声明的函数 (如 archive_read_new):
    //    #ifndef 为true → #define 映射生效 → 调用 vb6_di_archive_read_new() → 正确

    // C3内部导入名: 使用导出名构造唯一标识符
    // Fix 010b: 序号导出名 (如 "#644") 含非法C标识符字符, 需清洗
    // '#' → 'ord_', 其他非字母数字/下划线字符 → '_'
    std::string sanitizedExport = exportedName;
    for (size_t i = 0; i < sanitizedExport.size(); i++) {
        char c = sanitizedExport[i];
        if (c == '#') {
            sanitizedExport[i] = '_';
        } else if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') {
            sanitizedExport[i] = '_';
        }
    }
    // 确保不以数字开头 (合法C标识符要求)
    if (!sanitizedExport.empty() && (std::isdigit(static_cast<unsigned char>(sanitizedExport[0])) || sanitizedExport[0] == '_')) {
        // 对序号导出(#nnn清洗后为_nnn), 加ord前缀使名称更清晰
        if (!exportedName.empty() && exportedName[0] == '#') {
            sanitizedExport = "ord" + sanitizedExport;  // _644 → ord_644
        } else {
            sanitizedExport = "vb6_" + sanitizedExport;
        }
    }
    // ai/024 T02: 动态路保留 `vb6_di_` 命名空间 (转发桩桥到真实 API); 静态路直接用
    // **真实导出名** —— 归档里就是这么叫的, 少一层跳转, 也少一处要生成的桩。
    std::string cExportedIdent = isStaticDecl ? sanitizedExport : ("vb6_di_" + sanitizedExport);

    // === Fix <vbeclipse> 2026-10-06: Err.LastDllError 必须是"调用那一刻的快照" ===
    //
    // VB6 语义 (MSDN ErrObject.LastDllError; 社区对 VB6 运行时的逐条逆向一致):
    //   每次 VB6 代码里由 `Declare` 发起的 DLL 调用都会
    //     1) 先把本线程 last-error 清 0   —— SetLastError(0)
    //     2) 调用 API
    //     3) 返回后**立刻** GetLastError(), 把值存进 Err.LastDllError
    //   所以 Err.LastDllError 报的是"你自己那次调用"的错误码, 而不是访问那一刻的
    //   全局 last-error —— 后者早就被中间的 RTL/打印/字符串编组/控件内部 Win32
    //   调用冲掉了 (MSDN 原文: "The LastDllError property applies only to DLL
    //   calls made from Visual Basic code"; 值只保留到下一次外部调用)。
    //
    // 落地方式: 给每个 Declare 生成一个 static __inline 真函数做这三步, 再把调用点
    // 重定向到它。**不能用"在调用点插两条语句"** —— Declare 调用可以出现在
    // `While Api() <> 0` / `IIf(Api(), a, b)` / 实参等表达式位置, 语句注入会把 C
    // 撕成 `(Api(); vb6_ErrSetLastDllError(...) != 0)` 这种非法代码 (实测)。
    // 用函数则天然在任意表达式位置成立, 且返回值类型原样保留 (含 float/double/BSTR/
    // 结构体返回), 不需要给返回类型分类。
    //
    // 静态归档路 (isStaticDecl, 项目自带 .lib) 不包: 那是 C3 的内部机制而非 VB6
    // 的 DLL 调用, 且该路调用点直连真实符号, 无重定向可言。
    //
    // 别名判据 (原来在下方 227 行, Fix 217 提到这里 —— 它现在决定 callTarget 的**命名空间**):
    //   别名路 (aliasName 与 VB 名不同): 调用点无论 SDK 宏如何都直接发 callTarget
    //     → 包装必须无条件发出 (受 diGuard 防重), 否则调用点指向不存在的函数。
    //   非别名路: 调用点发 VB 名, 由 `#ifndef <VB名>` 的 #define 重定向;
    //     若 SDK 已把该名定义为宏 (GetTempPath/GetUserName 这类 A/W 家族),
    //     #define 不生效、调用走 SDK 原路 → 此时**不能**再发包装, 否则会凭空多出
    //     一个 vb6_di_X 引用 (无桩时 LNK2019)。所以非别名路的包装跟 #define 同进同出。
    const bool aliasedCall = (!aliasName.empty() && aliasName != node.name && !isStaticDecl);

    // Fix 217: 轻量包装的 C 名必须跟"去重用的那枚 guard"**同域**。
    // 两条路的去重钥匙不同:
    //   别名路  → `#ifndef VB6_DI_<导出名>_DEFINED` (同一个 DLL 导出只发一份包装)
    //   非别名路 → `#ifndef <VB名>`            (同一个 VB 名跨模块重复声明只发一份包装)
    // 原先两条路都叫 `vb6_lw_<导出名>`, 于是"某个模块的 VB 名恰好是**别人的导出名**"
    // 就撞在同一枚 C 符号上, 而两把 guard 各管各的, 谁也挡不住谁:
    // 实测 ComCtlsDemo —— Common/VisualStyles.bas 把 ActivateVisualStyles /
    // RemoveVisualStyles 都 `Alias "SetWindowTheme"` (别名路, 发 vb6_lw_SetWindowTheme),
    // Builds/TreeView/TreeView.ctl 又自己 `Private Declare PtrSafe Function
    // SetWindowTheme Lib "uxtheme"` (非别名路, 也发 vb6_lw_SetWindowTheme)。
    // TreeView.c 依次包含两个模块头 ⇒ 同一个 `static __inline vb6_lw_SetWindowTheme`
    // 两份完整定义 ⇒ C2084 ×N。这条错只写进临时文件 (driver_compile.cpp 事后又把
    // c3-error.log 覆盖成诊断文本), 外部只看到 "C3: 编译失败 (exit code 2)",
    // 于是表现为"文件没编译完就中断, 错误不好捕捉"。
    // 别名路改挂 `vb6_lwa_` 命名空间 (a = alias), 与它的 diGuard 同域; 非别名路照旧
    // `vb6_lw_` + VB 名。两条路的 C 名从此不可能相交。
    std::string callTarget = isStaticDecl ? cExportedIdent
                            : aliasedCall  ? ("vb6_lwa_" + sanitizedExport)
                                           : ("vb6_lw_" + cFuncIdent);

    // Fix 161b-decl-out: 登记 VB名(小写) → **调用点名** (cExportedIdent 全名)。
    // 目的: VB6 `Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA"`
    // 的 VB 名恰是 SDK 的 A/W 宏名 (WinBase.h `#define GetUserName GetUserNameW`)。
    // cgen 为避 SDK 冲突生成的 `#ifndef GetUserName` 守卫为假 ⇒ 调用点被 SDK 宏
    // 改写成宽版 GetUserNameW, 而我们按 VB6 语义生成 ANSI(char*) 编组 → 窄缓冲交给
    // 宽 API (t1 实测出参只回读首字符 'A'; t6 CreateWindowEx 直接建窗失败 hwnd=0)。
    // 用 cExportedIdent (= vb6_di_GetUserNameA / vb6_di_ord_410) 作调用点, 天然绕开
    // SDK 的 A/W 宏, 且对上 RTL 转发桩 (序号别名也因此落成 vb6_di_ord_410)。
    // 静态路直接用真实导出名, 无 SDK 宏抢占问题 → 不登记。
    // (aliasedCall 的判据已上移到 callTarget 定址处 —— 它现在同时决定旧名与包装所在的
    //  diGuard 域, 见那一段的 Fix 217。)
    if (aliasedCall) {
        std::string funcLower2 = node.name;
        std::transform(funcLower2.begin(), funcLower2.end(), funcLower2.begin(), ::tolower);
        declareAliasMap_[funcLower2] = callTarget;
    }

    // 2026-09-17: 把 Lib 家族写进生成头, 供 scripts/gen_di_stubs.ps1 按 DLL 家族
    // 把转发桩拆成多个文件。此前生成器只能靠外部的"未解析符号清单"决定要产出哪些
    // 桩, 而那份清单(.temp/unresolved_syms.txt)早已不存在 → 生成器无法重跑。
    // 格式固定为单行 `/* vb6_di_lib: <libName> */`, 紧邻其后的 vb6_di_* 原型即属该家族。
    // libName 已去引号、去 .dll 后缀 (见上文), 无 Lib 时为空串, 生成器按 unknown 处理。
    // ai/024 T02: 静态路不发这个标记 —— 那会让桩生成器把一个项目自带的归档当成
    // 新的 DLL 家族去产出 vb6_di_* 桩, 而静态路根本不走桩。
    if (!isStaticDecl) {
        h_.emitLine("/* vb6_di_lib: " + libName + " */");
    }

    // Fix 010i: 同一Declare函数可能出现在多个VB6模块中 (如CoTaskMemFree)
    // 用#ifndef guard防止__declspec(dllimport)声明重定义 (C2371)
    // ai/024 T02: 静态路用**另一套** guard 名 —— 同名的动态声明 (Lib "user32" 与
    // Lib "foo.lib" 撞名) 各自独立, 不会因为共用 guard 而被静默吞掉一条。
    std::string diGuard = (isStaticDecl ? "VB6_STATICLIB_" : "VB6_DI_")
                        + sanitizedExport + "_DEFINED";
    h_.emitLine("#ifndef " + diGuard);
    h_.emitLine("#define " + diGuard);
    // Fix 076: Changed from __declspec(dllimport) to extern declaration.
    // __declspec(dllimport) creates import symbols named vb6_di_Xxx that can't be
    // resolved by Windows import libraries (which export the real API names like
    // CloseEnhMetaFile, not vb6_di_CloseEnhMetaFile). Instead, we declare them as
    // extern and provide forwarding stubs in vb6rtl.c that bridge vb6_di_Xxx → real API.
    // ai/024 T02: 静态路这条 extern 就是**真实归档符号**本身, 无桩可桥 —— 引用符号
    // 由编译器按 callConv 生成 (x86 stdcall 得 `_MyAdd@16`), 与归档成员名一致即解析成功。
    h_.emitLine("extern " + retType + " " + callConv + " " + cExportedIdent + "(" + params + ");");

    // === Fix <vbeclipse> 2026-10-06: Err.LastDllError 快照包装 ===
    // 必须在 extern vb6_di_X 之后 emit (包装体要调用它)。
    // params 形如 "intptr_t a, BSTR b", 无参时是 "void" —— 逐参取末段 token 当形参名,
    // 转发时原样传回; 返回类型原样承接 (float/double/BSTR/结构体都不需要特殊处理)。
    auto emitLdlWrapper = [&]() {
        auto paramNamesFromList = [](const std::string& pl) -> std::string {
            if (pl.empty() || pl == "void") return "";
            std::string out;
            size_t i = 0;
            while (i < pl.size()) {
                size_t comma = pl.find(',', i);
                std::string part = pl.substr(i, (comma == std::string::npos) ? std::string::npos : comma - i);
                size_t sp = part.find_last_of(" \t");
                std::string nm = (sp == std::string::npos) ? part : part.substr(sp + 1);
                while (!nm.empty() && (nm.front() == ' ' || nm.front() == '\t')) nm.erase(nm.begin());
                while (!nm.empty() && (nm.back() == ' ' || nm.back() == '\t')) nm.pop_back();
                if (!nm.empty()) { if (!out.empty()) out += ", "; out += nm; }
                if (comma == std::string::npos) break;
                i = comma + 1;
            }
            return out;
        };
        std::string fwdCall = cExportedIdent + "(" + paramNamesFromList(params) + ")";
        h_.emitLine("/* Fix <vbeclipse> 2026-10-06: Declare -> Err.LastDllError 快照");
        h_.emitLine(" * (调用前 SetLastError(0), 返回即 GetLastError(); VB6 语义) */");
        if (retType == "void") {
            h_.emitLine("static __inline void " + callConv + " " + callTarget + "(" + params + ") {");
            h_.emitLine("    SetLastError(0);");
            h_.emitLine("    " + fwdCall + ";");
            h_.emitLine("    vb6_ErrSetLastDllError((int32_t)GetLastError());");
            h_.emitLine("}");
        } else {
            h_.emitLine("static __inline " + retType + " " + callConv + " " + callTarget + "(" + params + ") {");
            h_.emitLine("    " + retType + " _r;");
            h_.emitLine("    SetLastError(0);");
            h_.emitLine("    _r = " + fwdCall + ";");
            h_.emitLine("    vb6_ErrSetLastDllError((int32_t)GetLastError());");
            h_.emitLine("    return _r;");
            h_.emitLine("}");
        }
    };

    // 别名路: 调用点无条件直发 callTarget, 所以包装也**无条件**发出。
    // 必须留在 diGuard 里面: 多模块工程里同一个 Declare 会出现在多个模块头
    // (VBFlexGridBase.h / VisualStyles.h / VTableHandle.h ...), 每个头都带同一个
    // diGuard 名; 放在 guard 外面会在同一个 .c 里撞同名 `static __inline` 重复定义
    // (实测 C2084: 'vb6_lw_RtlMoveMemory' 已经有函数体)。
    if (!isStaticDecl && aliasedCall) emitLdlWrapper();
    h_.emitLine("#endif");

    // 生成: #define <VB6名> → <调用点名> (仅当VB6名未被SDK定义为宏时)
    // #ifndef 检查处理两种情况:
    //   - SDK宏 (CopyMemory等): #ifndef为false, 跳过 → 调用使用SDK宏
    //   - SDK函数声明: #ifndef为true, 生成 → 调用重定向到C3导入版本
    //   - 无SDK定义: #ifndef为true, 生成 → 正常
    // ai/024 T02: 静态路且无 Alias 时两者同名 (`MyAdd`→`MyAdd`), 自指的 #define 无意义
    // 且会让调试时宏展开停不下来, 直接跳过。
    // Fix <vbeclipse> 2026-10-06: 非别名路的包装**跟 #define 同进同出** —— SDK 已把
    // VB 名定义成宏时两者一起跳过, 否则会凭空多出一个 vb6_di_X 引用 (无桩时 LNK2019)。
    if (cFuncIdent != callTarget) {
        h_.emitLine("#ifndef " + cFuncIdent);
        if (!isStaticDecl && !aliasedCall) emitLdlWrapper();
        h_.emitLine("#define " + cFuncIdent + " " + callTarget);
        h_.emitLine("#endif");
    }
}

} // namespace vb6c3
