// driver_link.cpp - C3 编译器驱动: 链接阶段
// 2026-09-17 从 src/driver/driver.cpp 纯搬移（逐行未改）：
//   原第 2080~2418 行

#include "driver/driver.hpp"
#include "driver/res_inventory.hpp"   // ai/029 C29-M: 让位判据要问那份 .res 里有没有 #1 清单
#include "common/diagnostics.hpp"
#include "common/encoding.hpp"
#include "ast/ast.hpp"
#include "backend/msvc_driver.hpp"
#include "driver/rtl_embedded.hpp"
#include <iostream>
#include <fstream>
#include <filesystem>
#include <cstdlib>
#include <algorithm>
#include <cctype>
#include <unordered_set>

namespace vb6c3 {

// vb6forms RTL 源文件集 (Fix 096 从 runLinker 内联清单抽出复用)
static void addFormsSources(MsvcDriverOptions& opts, const std::string& rtlDir) {
    opts.sourceFiles.push_back(rtlDir + "/vb6forms.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_ctrl.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_list.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_style.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_scroll.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_picture.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_picture_prop.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_ctrlarr.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_webview.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_widget.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_widget_prop.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_shape.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_progress.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_imagelist.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_statusbar.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_sstab.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_oledd.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_listview.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_memberobj.c");
    // ai/029 C29-8a: TreeView 标量属性面（原生 SysTreeView32）
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_treeview.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_olecon.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_data.c");
    // ai/029 C29-5a: Toolbar 的窗口与设计期按钮（原生 ToolbarWindow32）
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_toolbar.c");
    // ai/029 C29-DT-a: DTPicker 的窗口与标量属性面（原生 SysDateTimePick32）。
    // 这条是**第四处**登记：.rc + CMakeLists + rtl_embedded 那张名字表只管"解包到临时目录"，
    // 少了这里这一行，解包成功、cl 却根本不编它 ⇒ 全线 LNK2019 找不到 vb6_DTP_*（实测踩过）。
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_dtpicker.c");
    // ai/029 C29-MV-a: MonthView —— 同样是第四处登记，少这一行就是全线 LNK2019 找不到 vb6_MV_*
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_monthview.c");
    // ai/029 C29-RT-a: RichTextBox —— 第四处登记，少这行就是全线 LNK2019 找不到 vb6_RTB_*
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_richtextbox.c");
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_winsock.c");
    // ai/029 C29-SL-a: Slider —— 第四处登记，少这行就是全线 LNK2019 找不到 vb6_Slider_*
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_slider.c");
    // vb6forms_axsite.c 按功能家族拆 5 个编译单元 (2026-09-20): 伞文件本身不参与编译
    // 注意: axsite/ 下的 .c 解包后是平铺目录, 故这里写 basename 而非带子目录路径
    opts.sourceFiles.push_back(rtlDir + "/ax_site.c");
    opts.sourceFiles.push_back(rtlDir + "/ax_site_ext.c");
    opts.sourceFiles.push_back(rtlDir + "/ax_propbag.c");
    opts.sourceFiles.push_back(rtlDir + "/ax_load.c");
    opts.sourceFiles.push_back(rtlDir + "/ax_host.c");
    // Fix 112: 工程内 UserControl 实例宿主 + 宿主对象模型 (2026-09-19 按族拆 6 单元)
    // 注意: uc/ 下的 .c 解包后是平铺目录，故这里写 basename 而非带子目录路径
    opts.sourceFiles.push_back(rtlDir + "/uc_host.c");
    opts.sourceFiles.push_back(rtlDir + "/uc_host_window.c");
    opts.sourceFiles.push_back(rtlDir + "/uc_hostmodel.c");
    opts.sourceFiles.push_back(rtlDir + "/uc_controls.c");
    opts.sourceFiles.push_back(rtlDir + "/uc_collection.c");
    opts.sourceFiles.push_back(rtlDir + "/uc_debug.c");
    // Fix 120-141 移植 (2026-09-20): PropertyBag (IDispatch) 编译单元
    opts.sourceFiles.push_back(rtlDir + "/uc_propbag.c");
    // Fix 148: OCX 真宿主 (免注册 LoadLibrary + DllGetClassObject) —— NewTab 等第三方 32 位 OCX
    opts.sourceFiles.push_back(rtlDir + "/vb6forms_axcontainer.c");
}

// ============================================================
// ai/vb-asm-extension-spec: Asm 过程 → MASM (.asm) → ml64 → .obj
//   v1 (x64): 每个「函数体 = 单个 Asm 块」的过程降级为独立 MASM 过程。
//   按名引用 `[param]` → Win64 ABI 寄存器 (RCX,RDX,R8,R9), `[Function]` → RAX;
//   `'` 注释 → MASM `;`; `.name:` 局部标签 → `<proc>_<name>` (MASM 无 proc 局部标签,
//   且一个 .asm 里多个 PROC 的裸标签会撞名)。
// ============================================================

static void toLowerAscii(std::string& s) {
    for (auto& c : s) c = static_cast<char>(::tolower(static_cast<unsigned char>(c)));
}

// 把一个 Asm 过程写成 MASM PROC 体
static void emitMasmProc(std::ostream& os, const AsmProcInfo& p) {
    // 先判有没有栈参数 (第 5 个起 / 浮点溢出): 有则改用 RBP 帧指针寻址。
    // 起因: callee-saved 的 push 会移动 RSP, 使 [rsp+40] 这类偏移失效;
    // Win64 在**非叶函数**里也允许/推荐用 RBP 建帧 (即使没有 SEH)。
    bool hasStackParam = false;
    for (auto& s : asmClassifyParams(p.params, "x64"))
        if (s.cls == AsmParamClass::Stack) { hasStackParam = true; break; }

    // [name] → ABI 寄存器 / [Function] → 与返回类型同宽的返回寄存器。
    // 替换表在 asm_proc.hpp (asmBuildX64Subs) —— codegen 期的宽度校验 (3038) 用的是
    // 同一张表, 改映射两边自动一致, 不会再各写一份走偏。
    //
    // 混排 (项2/项3) 例外: cgen 已把 [X] 预替换成 [rcx] 这类**地址解引用**形态
    // (asmRewriteAddrRefs), 参数表里的 vb6_a0… 只是"这里有一个整型参数"的占位,
    // 不能再做一遍 ABI 替换 (会把 [rcx] 里的 rcx 当成另一个参数名去查)。
    std::vector<std::pair<std::string, std::string>> subs;
    if (!p.linesFinal) subs = asmBuildX64Subs(p);
    std::vector<std::string> body;
    if (p.linesFinal) {
        // 只做注释 / 括号空白 / 局部标签规整, 不动操作数
        body = asmRewriteLines(p.lines, {}, p.cName);
    } else if (hasStackParam) {
        // 有帧时栈参相对 RBP。Win64 被调方入口布局 (自 RBP 向上):
        //   [rbp+0]  已保存的 rbp
        //   [rbp+8]  返回地址 (call 压入)
        //   [rbp+16] 调用方预留的 32 字节 shadow space 起点
        //   [rbp+48] 第 5 个参数 (shadow 之上), [rbp+56] 第 6 个 …
        // 故实际偏移 = 16 + 32 + 8k = 48 + 8k。
        auto slots = asmClassifyParams(p.params, "x64");
        int k = 0;
        for (auto& s : slots) {
            if (s.cls != AsmParamClass::Stack) continue;
            std::string to = "[rbp+" + std::to_string(48 + 8 * k) + "]";
            for (auto& sub : subs)
                if (sub.first == "[" + s.name + "]") sub.second = to;
            k++;
        }
        body = asmRewriteLines(p.lines, subs, p.cName);
    } else {
        body = asmRewriteLines(p.lines, subs, p.cName);
    }

    // callee-saved 自动保存 (spec §5 第 3 条 / §7): 扫描块内实际用到的 + clobber 声明的。
    // `<Naked>` 下不生成任何保存代码 —— 用户全权负责 (含自己 ret)。
    std::vector<std::string> saved =
        p.naked ? std::vector<std::string>()
                : asmSavedRegsForArch(p.lines, p.clobbers, /*x64=*/true);

    os << "; Win64 ABI: RCX,RDX,R8,R9 = 整型参数; XMM0-3 = 浮点; RAX = 返回\n";
    if (!saved.empty()) {
        os << "; callee-saved 自动保存:";
        for (auto& r : saved) os << " " << r;
        os << " (块内使用/ clobber 声明)\n";
    }
    if (hasStackParam) os << "; 有栈参数 → 建 RBP 帧 (栈参寻址 [rbp+16+8k])\n";
    os << p.cName << " PROC\n";
    // 建帧 (有栈参数或需保存 RBP 时)。顺序: push rbp → mov rbp,rsp → 其余 callee-saved。
    bool frame = hasStackParam || (!p.naked && saved.size() &&
                 std::find(saved.begin(), saved.end(), std::string("rbp")) != saved.end());
    std::vector<std::string> pushList = saved;
    if (hasStackParam && std::find(pushList.begin(), pushList.end(), std::string("rbp")) == pushList.end())
        pushList.insert(pushList.begin(), "rbp");
    if (hasStackParam) {
        os << "    push rbp\n";
        os << "    mov  rbp, rsp\n";
        for (auto& r : pushList) if (r != "rbp") os << "    push " << r << "\n";
    } else {
        for (auto& r : pushList) os << "    push " << r << "\n";
    }
    (void)frame;

    bool lastWasRet = false;
    for (auto& line : body) {
        os << "    " << line << "\n";
        std::string tline = line; toLowerAscii(tline);
        size_t s = tline.find_first_not_of(" \t");
        lastWasRet = (s != std::string::npos && tline.compare(s, 3, "ret") == 0 &&
                      (s + 3 >= tline.size() || tline[s + 3] == ' ' || tline[s + 3] == ';'));
    }
    if (hasStackParam) {
        // 有帧: 先逆序 pop 掉 rbp 之后压的 callee-saved, 再 `leave` 复位 rsp/rbp。
        for (auto it = pushList.rbegin(); it != pushList.rend(); ++it)
            if (*it != "rbp") os << "    pop " << *it << "\n";
        os << "    leave\n";
    } else {
        for (auto it = pushList.rbegin(); it != pushList.rend(); ++it) os << "    pop " << *it << "\n";
    }
    if (!p.naked && !lastWasRet) os << "    ret\n";   // 非 Naked: 叶函数, 编译器补返回
    os << p.cName << " ENDP\n";
}

// driver_link.cpp 内的落盘 + 汇编 + 收集
static bool assembleAsmProcs(const std::vector<EmittedAsmProc>& procs,
                             const std::string& intermediatesDir, MsvcDriverOptions& msvcOpts,
                             bool verbose) {
    namespace fs = std::filesystem;
    std::string ml64 = MsvcDriver::findMl64Exe();

    // 按模块基名分组 → 一个模块一个 .asm (多 PROC 同文件)
    std::vector<std::string> order;
    std::map<std::string, std::vector<const AsmProcInfo*>> byBase;
    for (auto& ep : procs) {
        std::string key = ep.moduleBase;
        if (byBase.find(key) == byBase.end()) order.push_back(key);
        byBase[key].push_back(&ep.info);
    }

    for (auto& base : order) {
        std::string asmPath = intermediatesDir + "/" + base + "_asm.asm";
        std::string objPath = intermediatesDir + "/" + base + "_asm.obj";
        // 调试逃生舱: C3_KEEP_ASM=<目录> 时把生成的 .asm 额外拷一份过去 (排障用)。
        const char* keepAsm = std::getenv("C3_KEEP_ASM");
        {
            std::ofstream ofs = ofstreamUtf8(asmPath, std::ios::out | std::ios::trunc);
            if (!ofs) {
                std::cerr << "C3: error: 无法写入汇编文件: " << asmPath << std::endl;
                return false;
            }
            ofs << "; === C3 auto-generated (ai/vb-asm-extension-spec) ===\n";
            ofs << "; Asm 块过程降级为独立 MASM 过程 (Win64 ABI)\n";
            ofs << "_TEXT SEGMENT\n";
            for (auto* p : byBase[base]) emitMasmProc(ofs, *p);
            ofs << "_TEXT ENDS\n";
            ofs << "END\n";
        }
        if (keepAsm && keepAsm[0]) {
            std::error_code ec;
            std::string dest = std::string(keepAsm) + "/" + base + "_asm.asm";
            fs::copy_file(utf8ToPath(asmPath), utf8ToPath(dest),
                          fs::copy_options::overwrite_existing, ec);
            if (!ec) std::cerr << "C3: [C3_KEEP_ASM] " << dest << std::endl;
        }
        // ml64 /c /Fo <obj> <asm> —— 外层多包一层引号: executeCommand 走
        // "cmd /c <cmd>", cmd 在 /c 后首字符是引号时会剥首尾各一个 (同 rc.exe 的先例)。
        std::string args = "\"" + ml64 + "\" /nologo /c /Fo \"" + objPath + "\" \"" + asmPath + "\"";
        if (verbose) std::cout << "C3: ml64: " << args << std::endl;
        int ret = MsvcDriver::executeCommand("\"" + args + "\"");
        if (ret != 0 || !fs::exists(utf8ToPath(objPath))) {
            std::cerr << "C3: error: ml64 汇编失败 (" << asmPath << "), 退出码 " << ret << std::endl;
            return false;
        }
        msvcOpts.extraObjects.push_back(objPath);
    }
    return true;
}

// ============================================================
// ai/022 B17: rc.exe 的唯一发现处 (TypeLib 资源与 VS_VERSION_INFO 两处共用)。
// 旧写法只有 "WindowsSdkDir 环境变量 + C:\Program Files (x86)\Windows Kits\10" 两条路，
// 于是 SDK 装在别的盘 (本机 = D:\Windows Kits\10) 且没导出该环境变量时**静默不嵌资源** ——
// 实测产出的 DLL 连 .rsrc 段都没有，类型库注册表项写不进去，外部客户按 LIBID 找不到库。
// 顺序与 tests\run_tests.ps1 的 SDK 探测对齐: 环境变量 → Program Files → 盘符扫描 → PATH。
// 目录内取**版本号最大**的那个 (旧写法取 directory_iterator 的最后一个, 结果随枚举顺序变)。
// ============================================================

static bool versionDirGreater(const std::string& a, const std::string& b) {
    auto parse = [](const std::string& s) {
        std::vector<int> v;
        size_t i = 0;
        while (i < s.size() && v.size() < 4) {
            int n = 0;
            bool any = false;
            while (i < s.size() && s[i] >= '0' && s[i] <= '9') { n = n * 10 + (s[i] - '0'); i++; any = true; }
            if (!any) return std::vector<int>();
            v.push_back(n);
            if (i < s.size() && s[i] == '.') i++; else break;
        }
        return v;
    };
    std::vector<int> va = parse(a), vb = parse(b);
    if (va.empty() || vb.empty()) return a > b;
    return va > vb;
}

// rc.exe 与 mt.exe 同在 <SDK>\bin\<ver>\<arch>\ 下, 发现逻辑完全一致。
// **只找 x64 这一份**: mt.exe 自身的位数与目标位数无关 —— 实测 x64 的 mt.exe
// 注入 x86 的 exe 一样能把 comctl32 拉到 v6 (见 embedManifestViaMt 的实测记录);
// 而 rc.exe 的位数在 rc 侧是有意义的 (x86 目标该用 x86 那份, 别混)。
static std::string findToolInSdkBin(const std::filesystem::path& binDir, const char* tool) {
    std::error_code ec;
    if (!std::filesystem::exists(binDir, ec)) return std::string();
    std::string exeName = std::string(tool) + ".exe";
    std::vector<std::string> vers;
    for (const auto& entry : std::filesystem::directory_iterator(binDir, ec)) {
        if (!entry.is_directory()) continue;
        std::error_code ec2;
        if (std::filesystem::exists(entry.path() / "x64" / exeName, ec2)) vers.push_back(entry.path().filename().string());
    }
    if (vers.empty()) return std::string();
    std::sort(vers.begin(), vers.end(), [](const std::string& a, const std::string& b) {
        return versionDirGreater(a, b);   // 大的在前
    });
    return (binDir / vers.front() / "x64" / exeName).string();
}

static std::string findToolInPath(const char* tool) {
    std::error_code ec;
    std::string exeName = std::string(tool) + ".exe";
    // PATH 上直接有的场合 (VS 开发者提示符把 <sdk>\bin\<ver>\x64 塞进 PATH)
    if (const char* path = std::getenv("PATH"); path && path[0]) {
        std::string p = path;
        size_t start = 0;
        while (start <= p.size()) {
            size_t sep = p.find(';', start);
            std::string dir = p.substr(start, sep == std::string::npos ? std::string::npos : sep - start);
            if (!dir.empty()) {
                std::filesystem::path cand = std::filesystem::path(dir) / exeName;
                if (std::filesystem::exists(cand, ec)) return cand.string();
            }
            if (sep == std::string::npos) break;
            start = sep + 1;
        }
    }
    return std::string();
}

static std::string findTool(const char* tool) {
    std::error_code ec;
    std::string exeName = std::string(tool) + ".exe";
    std::filesystem::path toolsCopy = std::filesystem::current_path() / "tools" / exeName;
    if (std::filesystem::exists(toolsCopy, ec)) return toolsCopy.string();

    std::vector<std::filesystem::path> binDirs;
    if (const char* sdkDir = std::getenv("WindowsSdkDir"); sdkDir && sdkDir[0]) {
        std::string root = sdkDir;
        while (!root.empty() && (root.back() == '\\' || root.back() == '/')) root.pop_back();
        binDirs.emplace_back(root + "\\bin");
    }
    for (const char* envName : {"ProgramFiles(x86)", "ProgramFiles"}) {
        if (const char* base = std::getenv(envName); base && base[0]) {
            binDirs.emplace_back(std::filesystem::path(base) / "Windows Kits" / "10" / "bin");
        }
    }
    for (const char* drive : {"C:", "D:", "E:", "F:"}) {
        binDirs.emplace_back(std::filesystem::path(std::string(drive) + "\\") / "Windows Kits" / "10" / "bin");
    }
    for (const auto& dir : binDirs) {
        std::string found = findToolInSdkBin(dir, tool);
        if (!found.empty()) return found;
    }
    return findToolInPath(tool);
}

static std::string findRcExe() { return findTool("rc"); }

// mt.exe 与 rc.exe 是同一批工具, 发现方式共用。它的位数无关紧要。
static std::string findMtExe() { return findTool("mt"); }

// ============================================================
// (#44 rev2, SSTabEx 实证 2026-09-27): 用户 .res 里 RT_MANIFEST 的
// processorArchitecture 归一化。
//
// 坑的实证链: VB6 IDE 生成的 manifest 恒写 processorArchitecture="x86"
// (VB6 本身 32 位)。x64 宿主带 x86 标记的 Common-Controls 依赖 → SxS 激活
// 上下文创建失败 → 进程初始化即 APP_INIT_FAILURE (0xC0000145), 栈顶
// ntdll!NtRaiseHardError。交互桌面会弹"无法启动"框; 无头/自动化桌面连框都
// 没人点 → 表现为"启动即挂死、零窗口、Responding=True"(MainWindowHandle=0
// 时 Responding 恒真, 假绿)。SSTabEx Test.exe 两次启动都这样, 把 exe 内
// manifest 的 x86 原位改成 "*" 后窗口立即出来 —— 因果闭环。
//
// 规则: 只在【标记与目标位数冲突】时把该属性改写为 "*" (SxS 通配语义):
//   x64 目标撞 x86/ia64 改;  x86 目标撞 amd64/ia64 改; 其余原样。
// 输出写到中间目录副本, 绝不动用户原文件; 无 RT_MANIFEST / 无冲突 →
// 返回空串, 调用方直传原文件。
//
// .res 格式要点 (VisualStyleManifest.res 实测校准): 每条目 =
//   DWORD DataSize + DWORD HeaderSize + header body(HeaderSize-8) + data
//   + pad 到 4 对齐 (下一条目起点的对齐; 末条目 pad 可省)。
// HeaderSize **含开头 8 字节**(marker 条目 ds=0 hs=32, body 恰 24 字节)。
// 数值型 Type/Name = WORD 0xFFFF + WORD id (RT_MANIFEST: 0xFFFF,0x0018;
// 资源 id 1 是激活上下文约定入口)。
// ============================================================
static bool fixManifestArchAttrsInPlace(std::string& xml, const std::string& arch) {
    // 目标 arch → 需要改掉的标记集
    auto lc = [](std::string s) {
        for (auto& c : s) c = static_cast<char>(::tolower(static_cast<unsigned char>(c)));
        return s;
    };
    std::string target = lc(arch);
    const char* bad[2] = {nullptr, nullptr};
    if (target == "x64") { bad[0] = "x86"; bad[1] = "ia64"; }
    else if (target == "x86") { bad[0] = "amd64"; bad[1] = "ia64"; }
    if (!bad[0]) return false;

    // 编码探测: UTF-16LE 时 ASCII 字节后跟 0x00; 否则按字节 (ASCII/UTF-8)
    bool utf16 = false;
    for (size_t i = 1; i + 1 < xml.size() && i < 64; i += 2) {
        if (xml[i] == '\0' && xml[i - 1] != '\0') { utf16 = true; break; }
        if (xml[i] != '\0') break;
    }
    size_t step = utf16 ? 2 : 1;
    // 属性名比较**大小写不敏感** — XML 实际写的是 processorArchitecture (大写 A),
    // attr 常量已全小写, 这里把输入侧小写化; attr 本身是 ASCII, tolower 按字节即可
    auto eqAt = [&](size_t off, const char* s) {
        for (size_t i = 0; s[i]; ++i) {
            if (off + i * step + (step - 1) >= xml.size()) return false;
            char c = xml[off + i * step];
            if (static_cast<char>(::tolower(static_cast<unsigned char>(c))) != s[i]) return false;
            if (utf16 && xml[off + i * step + 1] != '\0') return false;
        }
        return true;
    };
    const std::string attr = "processorarchitecture";
    // 值读取: attr 后跳空白, '=', 跳空白, '"', 读到 '"'
    auto skipWs = [&](size_t& o) {
        while (o < xml.size() && (xml[o] == ' ' || xml[o] == '\t' || xml[o] == '\r' || xml[o] == '\n')) o += step;
    };
    std::string out;
    bool changed = false;
    size_t pos = 0;
    // 扫描**全部**出现 — manifest 里 assembly 自身 identity 与 dependency 各有一个
    // processorArchitecture, 实测两处都标 x86 (VisualStyleManifest.res), 都要改
    while (pos < xml.size()) {
        // 属性名前必有 '<' 或空白, 直接试 eqAt; 误命中风险由
        // 后续 '=' + '"' 结构校验兜住
        if (eqAt(pos, attr.c_str())) {
            size_t q = pos + attr.size() * step;
            skipWs(q);
            if (q < xml.size() && xml[q] == '=') {
                q += step;
                skipWs(q);
                if (q < xml.size() && xml[q] == '"') {
                    q += step;
                    size_t v0 = q;
                    while (q < xml.size() && xml[q] != '"') q += step;
                    if (q < xml.size()) {
                        // 提取值 (按字节)
                        std::string val;
                        for (size_t o = v0; o < q; o += step) val.push_back(xml[o]);
                        std::string vlc = lc(val);
                        if ((bad[0] && vlc == bad[0]) || (bad[1] && vlc == bad[1])) {
                            out.append(xml, pos, v0 - pos);       // 属性名..'"/值前
                            if (utf16) out.push_back('*'), out.push_back('\0');
                            else out.push_back('*');
                            changed = true;
                            pos = q;                              // 从收尾 '"' 继续, 不丢属性后文
                            continue;
                        }
                    }
                }
            }
        }
        // 未命中: 原样带过本字节
        if (utf16) {
            out.push_back(xml[pos]);
            if (pos + 1 < xml.size()) out.push_back(xml[pos + 1]);
        } else out.push_back(xml[pos]);
        pos += step;
    }
    if (changed) xml.swap(out);
    return changed;
}

// 返回归一化副本路径; 无需修改时返回空串 (调用方直传原文件)。
static std::string normalizeUserResForArch(const std::string& resPathUtf8,
                                           const std::string& arch,
                                           const std::string& interDirUtf8,
                                           bool verbose) {
    namespace fs = std::filesystem;
    std::error_code ec;
    if (resPathUtf8.empty() || !fs::exists(utf8ToPath(resPathUtf8), ec)) return {};

    std::ifstream f(utf8ToPath(resPathUtf8), std::ios::binary);
    if (!f) return {};
    std::string in((std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
    if (in.size() < 8) return {};

    auto rd16 = [&](size_t o) -> uint16_t {
        return uint16_t(uint8_t(in[o]) | (uint16_t(uint8_t(in[o + 1])) << 8));
    };
    auto rd32 = [&](size_t o) -> uint32_t {
        return uint32_t(uint8_t(in[o])) | (uint32_t(uint8_t(in[o + 1])) << 8)
             | (uint32_t(uint8_t(in[o + 2])) << 16) | (uint32_t(uint8_t(in[o + 3])) << 24);
    };

    std::string out;
    int fixedCount = 0;
    size_t off = 0;
    bool corrupt = false;
    while (off + 8 <= in.size()) {
        uint32_t ds = rd32(off), hs = rd32(off + 4);
        if (hs < 24 || (uint64_t)off + hs > in.size()
            || (uint64_t)off + hs + (uint64_t)ds > in.size()) {
            corrupt = true;   // 结构不认识: 整体放弃, 保持直传
            break;
        }
        size_t body = off + 8, dataOff = off + hs;
        bool isManifest = (rd16(body) == 0xFFFF && rd16(body + 2) == 24);
        out.append(in, off, hs);   // ds + hs + header body 原样
        if (isManifest) {
            std::string xml = in.substr(dataOff, ds);
            if (fixManifestArchAttrsInPlace(xml, arch)) {
                ++fixedCount;
                uint32_t nds = (uint32_t)xml.size();
                out[off + 0] = uint8_t(nds);
                out[off + 1] = uint8_t(nds >> 8);
                out[off + 2] = uint8_t(nds >> 16);
                out[off + 3] = uint8_t(nds >> 24);
            }
            out.append(xml);
        } else {
            out.append(in, dataOff, ds);
        }
        while (out.size() & 3) out.push_back('\0');
        off = (size_t(dataOff) + ds + 3) & ~size_t(3);
    }
    if (corrupt || fixedCount == 0) return {};

    // 副本放中间目录, 与 C3 自己生成的 .res 同居
    fs::path inter = fs::absolute(utf8ToPath(interDirUtf8), ec);
    std::string base = utf8ToPath(resPathUtf8).filename().u8string();
    std::string copyName = base + ".c3.archfix.res";
    std::string copyPath = pathToUtf8(inter / utf8ToPath(copyName));
    std::ofstream o(utf8ToPath(copyPath), std::ios::binary | std::ios::trunc);
    if (!o) return {};
    o.write(out.data(), std::streamsize(out.size()));
    if (!o) return {};
    if (verbose) {
        std::cout << "C3: user .res manifest arch normalized for " << arch
                  << " target (" << fixedCount << " attr): " << copyPath << std::endl;
    }
    return copyPath;
}

// 应用清单在**链接完之后**用 mt.exe 注入 (P20-41 rev2)。
//
// 为什么不是 rc.exe → 链接器: rc 侧只为 RT_MANIFEST 留了字符串写法
// `1 RT_MANIFEST "x"`, 而它把类型记成字符串名而非数值 (RT_MANIFEST **是 24**,
// 不是 16 —— 16 是 RT_VERSION), 链接后资源目录是 named=1 / type=88(0x58),
// Windows 激活上下文不认 → comctl32 停在 v5.82。
// mt.exe 这条实测是通的, x64/x86 目标都验过:
//   无清单对照  → comctl32 FileVersion 5.82.*.*
//   mt.exe 注入 → comctl32 FileVersion 6.10.*.*   (即 v6)
// 而且**不用把任何 .res 交给链接器**; mt.exe 自身的位数也不挑
// (用 x64 那份给 x86 的 exe 注入同样生效)。
//
// 顺带记一条 mt.exe 的脾气, 见 embedManifestViaMt 尾部的体积复核:
// 目标 PE 若完全没有资源目录, mt.exe **返回 0 却什么都不做**。
static bool embedManifestViaMt(const std::string& exePath, const std::string& manifestXml, bool verbose) {
    // 每个早退点都要喊一声: 静默不嵌 = 32/64 位出货悄悄退回 comctl32 v5.82,
    // 而这类"看起来构建成功"的失败最难查。
    if (exePath.empty() || manifestXml.empty()) {
        std::cerr << "C3: warning: no exe path / manifest to embed" << std::endl;
        return false;
    }
    if (!existsUtf8(exePath) || !existsUtf8(manifestXml)) {
        std::cerr << "C3: warning: manifest embed skipped, file missing (exe="
                  << exePath << " manifest=" << manifestXml << ")" << std::endl;
        return false;
    }

    std::string mtExe = findMtExe();
    if (mtExe.empty()) {
        std::cerr << "C3: warning: mt.exe not found, application manifest will not be embedded" << std::endl;
        return false;
    }
    if (verbose) {
        std::cout << "C3: MT exe: " << mtExe << std::endl;
    }

    // 资源 id 1 是激活上下文的约定入口。整串再包一层引号, 与 rc 那几处同款
    // (否则 cmd /c 会把 mt.exe 路径剥坏)。
    std::ostringstream mtArgs;
    mtArgs << "\"" << mtExe << "\" -manifest \"" << manifestXml
           << "\" -outputresource:\"" << exePath << ";#1\"";
    if (verbose) {
        std::cout << "C3: MT (manifest): " << mtArgs.str() << std::endl;
    }
    // mt.exe 的退出码不可靠: 目标 PE 若**完全没有资源目录**(ResourceDirRVA==0),
    // 它照样返回 0, 但一个字节都没写进去。实测 n0.exe(32 位、无 .rsrc)跑完
    // 仍是 ResourceDirRVA==0、文件体积不变。所以退出码之后必须复核体积:
    // 没变大就说明这次注入是空转, 此时 activation context 拿不到清单,
    // comctl32 会停在 v5.82 —— 必须喊出来, 不能让它假装成功。
    std::uintmax_t sizeBefore = std::filesystem::file_size(utf8ToPath(exePath));

    int ret = MsvcDriver::executeCommand("\"" + mtArgs.str() + "\"");
    if (ret != 0) {
        std::cerr << "C3: warning: mt.exe failed to embed the application manifest (exit " << ret << ")" << std::endl;
        return false;
    }

    if (std::filesystem::file_size(utf8ToPath(exePath)) == sizeBefore) {
        // 注入成功时一定会新长出 .rsrc, 体积不变几乎只可能是 mt.exe 空转。
        std::cerr << "C3: warning: mt.exe reported success but the exe did not change "
                  << "(" << sizeBefore << " bytes) -- the target probably has no resource "
                  << "directory, and no manifest got in. comctl32 will stay v5.82." << std::endl;
        return false;
    }

    if (verbose) {
        std::cout << "C3: application manifest embedded: " << exePath << std::endl;
    }
    return true;
}

bool Driver::runLinker(const CompileOptions& options, const std::string& outputDir,
                       const std::string& intermediatesDir, SessionManager& session) {
    // --emit-c mode: no linking needed
    if (options.emitC) {
        return true;
    }

    // Check MSVC availability
    if (!MsvcDriver::isMsvcAvailable()) {
        std::cerr << "C3: error: MSVC not found (install Visual Studio 2017+ with C++ workload)" << std::endl;
        std::cerr << "C3: Use --emit-c to generate C code only" << std::endl;
        return false;
    }

    // Collect generated .c files from intermediatesDir
    // MUST match baseName logic in runCodeGeneration (single-file + -o uses output stem)
    MsvcDriverOptions msvcOpts;
    for (size_t i = 0; i < modules_.size(); i++) {
        // 泛型模板类 (G4): 未发码 (见 driver_codegen_module_loop), 无 .c 可链
        if (!modules_[i]->classTypeParams.empty()) continue;
        std::string baseName;
        if (modules_.size() == 1 && !options.outputFile.empty()) {
            std::filesystem::path p(utf8ToPath(options.outputFile));
            baseName = pathToUtf8(p.stem());
        } else {
            // Fix 013: 用 module.moduleName (VB_Name) 作为基名, 与 runCodeGeneration 一致
            baseName = modules_[i]->moduleName;
        }
        std::string cPath = intermediatesDir + "/" + baseName + ".c";
        msvcOpts.sourceFiles.push_back(cPath);
    }

    // P6.6: ActiveX DLL mode, add dll_entry.c
    if (options.isDll) {
        std::string dllEntryPath = intermediatesDir + "/dll_entry.c";
        msvcOpts.sourceFiles.push_back(dllEntryPath);
    }

    // P10: Get RTL directory from session
    std::string rtlDir = session.rtlDir();
    if (rtlDir.empty()) {
        std::cerr << "C3: error: RTL runtime not available" << std::endl;
        return false;
    }
    msvcOpts.rtlDir = rtlDir;

    // P11.1+P11.2: Set intermediate directories
    msvcOpts.srcDir = intermediatesDir;   // /I for generated .h files
    msvcOpts.objDir = intermediatesDir;   // /Fo for .obj files

    // Output file path (in user's output directory, not intermediates)
    std::string outputExt = options.isDll ? ".dll" : ".exe";
    if (!options.outputFile.empty()) {
        msvcOpts.outputFile = options.outputFile;
        if (options.isDll && msvcOpts.outputFile.size() >= 4 &&
            msvcOpts.outputFile.compare(msvcOpts.outputFile.size()-4, 4, ".exe") == 0) {
            msvcOpts.outputFile.replace(msvcOpts.outputFile.size()-4, 4, ".dll");
        }
    } else if (!projectBaseName_.empty()) {
        msvcOpts.outputFile = outputDir + "/" + projectBaseName_ + outputExt;
    } else if (modules_.size() == 1) {
        std::filesystem::path p(utf8ToPath(modules_[0]->filename));
        msvcOpts.outputFile = outputDir + "/" + pathToUtf8(p.stem()) + outputExt;
    } else {
        msvcOpts.outputFile = outputDir + "/a" + outputExt;
    }

    msvcOpts.isDll = options.isDll;
    // P7: Detect GUI program
    for (auto& module : modules_) {
        if (module->isFormModule) {
            msvcOpts.isGui = true;
            break;
        }
    }
    // Fix 165: 入口点跟着启动对象走 — 见 msvc_driver.hpp 的 entryIsMain 注释。
    // 与 cgen_base_generate_entry.inc:54 用的是同一个判定 (`sub main`), 两处若改须同改。
    if (msvcOpts.isGui) {
        std::string so165 = startupObject_;
        std::transform(so165.begin(), so165.end(), so165.begin(), ::tolower);
        msvcOpts.entryIsMain = (so165 == "sub main");
    }

    // RTL 源码编译 (P10 恢复): 会话目录释放的 RTL .c 与生成代码一起编译,
    // 不再链接预编译 .lib —— 修改 RTL 源码后重编 C3.exe 即生效
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl.c");
    // vb6rtl.c 按家族拆分 (2026-09-17): 13 个族实现与主文件同批编译
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_string.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_format.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_conv.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_misc.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_system.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_compat.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_date.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_array.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_file.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_paramarray.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_financial.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_com.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6rtl_registry.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6com.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6com_invoke.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6com_pack.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6com_wrap.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6com_sink.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6com_foreach.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6com_collection.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6com_collection_enum.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_win32_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_user32_stubs.c");
    // gdiplus 二级拆分 (2026-09-20): 解包后是平铺目录, 故写 basename
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_gdiplus_draw_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_gdiplus_text_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_gdiplus_image_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_gdiplus_brush_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_gdiplus_path_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_crypto_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_com_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_net_stubs.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_shell_stubs.c");
    // Fix 160y: unknown 族 (无 vb6_di_lib 标记的杂项符号: Imm/version/TransparentBlt/msvbvm60)
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6_di_unknown_stubs.c");
    // 092z-3 + Fix 112: vb6forms 组对**所有**工程类型都链接 —— vb6rtl/vb6com 的
    // 宿主对象分派挂接点 (vb6_Host_*) 引用 vb6forms_uc, 不能只在 GUI/DLL 下链接.
    // (原先 Fix 096 的按需扫描已不需要: 一律链接.)
    addFormsSources(msvcOpts, rtlDir);
    // ExeComBridge 01: vb6comserver 组无条件链入 (原先仅 DLL).
    //   EXE 的工程类实例也要能被包装成 IDispatch (vb6_ComObject_FromInstance /
    //   vb6_FindCoClassDesc 定义在 vb6comserver_obj.c); 纯 EXE 下无调用方,
    //   不产生新的外部库依赖 (advapi32.lib 三种链接分支本就都带).
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6comserver.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6comserver_obj.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6comserver_factory.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6comserver_cp.c");
    msvcOpts.sourceFiles.push_back(rtlDir + "/vb6comserver_pci.c");
    // ExeComBridge 01: com_entry.c = EXE 模式的 coclass 表 (driver_codegen_dll_entry.inc
    //   生成于中间目录), 提供 g_vb6_coclasses / g_vb6_coclassCount (vb6comserver_obj.c
    //   的 vb6_FindCoClassDesc 引用). DLL 模式由 dll_entry.c 提供同名符号且不生成
    //   com_entry.c → 仅 EXE 链入, 无条件加会让 DLL 的 cl 报 C1083.
    if (!msvcOpts.isDll) {
        msvcOpts.sourceFiles.push_back(intermediatesDir + "/com_entry.c");
    }

    msvcOpts.verbose = options.verbose;
    msvcOpts.debugInfo = options.debugInfo;
    msvcOpts.optimizationLevel = options.optimizationLevel;
    msvcOpts.arch = options.arch;  // DualArch: pass target architecture

    // === ai/024 T02: 用户静态库 → 链接输入 ===
    // 两类来源合并:
    //   A) Declare 的 `Lib "x.lib"` 静态形态 → staticLibResolved_ (T01 已解析, 全是绝对路径)
    //   B) vbp `ExtraLib=` / CLI `--extra-lib` → extraLibResolved_ (绝对路径, 或"放行给
    //      链接器"的裸名 —— 靠 LIB 环境变量与下面的 /LIBPATH 找)
    // 发码侧对静态 Declare 只发 `extern <真实导出名>` 而**不发** `#pragma comment(lib,...)`
    // (见 cgen_decl_api.cpp 的 isStaticDecl), 所以这两份清单就是静态库进入链接的**唯一**通路。
    {
        std::unordered_set<std::string> seen;
        auto pushInput = [&](const std::string& p) {
            if (p.empty()) return;
            std::string key = p;
            for (size_t i = 0; i < key.size(); i++) {
                key[i] = static_cast<char>(::tolower(static_cast<unsigned char>(key[i])));
            }
            if (!seen.insert(key).second) return;
            msvcOpts.userLibInputs.push_back(p);
        };
        for (const auto& kv : staticLibResolved_) pushInput(kv.second);
        for (const auto& p : extraLibResolved_) pushInput(p);

        // 搜索根 → /LIBPATH:, 顺序即静态库搜索顺序:
        // vbp `LibDir=` → CLI `--libdir` → `<工程目录>/Lib` (显式总赢过隐式)。
        for (const auto& r : staticLibPaths_.roots()) msvcOpts.libSearchPaths.push_back(r);

        // ai/024 E4 (T04b): Alias "_foo@12" 逃生舱的 /alternatename 桥接指令。
        // 同名 Declare 在多模块重复出现会生成重复指令 → 按整串去重 (符号名
        // 大小写在链接器眼里有区分, 这里保序保原文)。
        {
            std::unordered_set<std::string> altSeen;
            for (const auto& a : staticLibAlternatenames_) {
                if (a.empty() || !altSeen.insert(a).second) continue;
                msvcOpts.alternatenames.push_back(a);
            }
        }

        if (options.verbose && (!msvcOpts.userLibInputs.empty() || !msvcOpts.libSearchPaths.empty())) {
            std::cout << "C3: link inputs -- user libs (" << msvcOpts.userLibInputs.size() << "):";
            for (const auto& l : msvcOpts.userLibInputs) std::cout << " " << l;
            std::cout << std::endl;
        }
    }

    // ai/030 T30-A: obj store 的根取自 %LOCALAPPDATA%\C3\objcache (跨工程/跨输出目录共享);
    // 这里给的只是**拿不到 LOCALAPPDATA 时的回退目录**，不再是缓存本体。
    msvcOpts.incremental = options.incremental;
    msvcOpts.incrementalCacheDir = outputDir + "/.c3obj";

    // P6.6: ActiveX DLL - generate .def export file (in intermediatesDir)
    if (options.isDll) {
        std::string defPath = intermediatesDir + "/activex_dll.def";
        std::ofstream defFile = ofstreamUtf8(defPath, std::ios::out | std::ios::trunc);
        if (defFile) {
            defFile << "LIBRARY\n";
            defFile << "EXPORTS\n";
            defFile << "    DllGetClassObject\n";
            defFile << "    DllCanUnloadNow\n";
            defFile << "    DllRegisterServer\n";
            defFile << "    DllUnregisterServer\n";
            defFile << "    DllMain\n";
            defFile.close();
            msvcOpts.defFile = defPath;
            if (options.verbose) {
                std::cout << "C3: Generated export definition: " << defPath << std::endl;
            }
        }
    }

    // P9: Embed TypeLib into DLL resource
    if (options.isDll && !options.dllProgId.empty()) {
        std::string tlbPath = pathToUtf8(std::filesystem::absolute(utf8ToPath(intermediatesDir + "/" + options.dllProgId + ".tlb")));
        if (existsUtf8(tlbPath)) {
            std::string absInterDir = pathToUtf8(std::filesystem::absolute(utf8ToPath(intermediatesDir)));
            std::string rcPath = absInterDir + "\\activex_dll_typelib.rc";
            {
                std::ofstream rcFile = ofstreamUtf8(rcPath, std::ios::out | std::ios::trunc);
                if (rcFile) {
                    std::string tlbPathForRc = tlbPath;
                    for (auto& c : tlbPathForRc) { if (c == '\\') c = '/'; }
                    rcFile << "1 TYPELIB \"" << tlbPathForRc << "\"\n";

                }
            }

            // Find rc.exe (共用发现处: 见 findRcExe)
            std::string rcExePath = findRcExe();

            if (!rcExePath.empty()) {
                std::string resPath = absInterDir + "\\activex_dll_typelib.res";
                std::ostringstream rcArgs;
                rcArgs << "\"" << rcExePath << "\" /r /fo \"" << resPath << "\" \"" << rcPath << "\"";
                if (options.verbose) {
                    std::cout << "C3: RC: " << rcArgs.str() << std::endl;
                }
                // executeCommand 内部拼的是 "cmd.exe /c <cmd>", 且带 CREATE_NO_WINDOW (不弹窗口)。
                // 必须再包一层引号: cmd 在「/c 之后首字符是引号」时会剥掉首尾各一个引号。
                // rcArgs 本身以 "<rc.exe 路径>" 开头 (路径含空格), 少这层外层引号时 cmd 会把
                // rc 路径剥成 C:\...\rc.exe" -> 找不到命令 -> rcRet != 0 被静默跳过,
                // TypeLib 资源不再嵌入 (实测 DLL 少 141KB)。多包一层让 cmd 剥外层、留下 rc 自己的引号。
                int rcRet = MsvcDriver::executeCommand("\"" + rcArgs.str() + "\"");
                if (rcRet == 0 && existsUtf8(resPath)) {
                    msvcOpts.typelibResFile = resPath;
                    if (options.verbose) {
                        std::cout << "C3: TypeLib resource embedded: " << resPath << std::endl;
                    }
                }
            } else {
                // 找不到 rc.exe = 类型库不嵌进 DLL ⇒ 注册表里也就没有 TypeLib 项，外部客户
                // 按 LIBID 找不到契约。这是**静默**的质量损失，所以不藏在 --verbose 后面 (B17)。
                std::cerr << "C3: rc.exe not found, TypeLib will not be embedded in DLL" << std::endl;
            }
        } else if (options.verbose) {
            std::cout << "C3: TypeLib file not found: " << tlbPath << std::endl;
        }
    }
    // P23-05: Generate VS_VERSION_INFO resource if version info is available
    if (verMajor_ > 0 || verMinor_ > 0 || !verCompanyName_.empty() || !verFileDescription_.empty()) {
        std::string absInterDir2 = pathToUtf8(std::filesystem::absolute(utf8ToPath(intermediatesDir)));
        std::string verRcPath = absInterDir2 + "\\version_info.rc";
        {
            std::ofstream rcFile = ofstreamUtf8(verRcPath, std::ios::out | std::ios::trunc);
            if (rcFile) {
                // Determine internal name from project base name or output file
                std::string internalName = projectBaseName_.empty() ? "VB6App" : projectBaseName_;
                std::string originalName = verOriginalFileName_.empty() ? (internalName + ".exe") : verOriginalFileName_;
                std::string prodName = verProductName_.empty() ? internalName : verProductName_;
                std::string fileDesc = verFileDescription_.empty() ? internalName : verFileDescription_;
                std::string company = verCompanyName_;
                std::string copyright = verLegalCopyright_;
                std::string comments = verComments_;
                std::string trademarks = verLegalTrademarks_;

                // Escape backslashes for RC string values
                auto escapeRc = [](std::string s) -> std::string {
                    std::string result;
                    for (char c : s) {
                        if (c == '\\') result += "\\\\";
                        else if (c == '"') result += "\\\"";
                        else result += c;
                    }
                    return result;
                };

                int fileVerMs = verMajor_;
                int fileVerLs = verMinor_;
                int prodVerMs = verMajor_;
                int prodVerLs = verMinor_;

                                rcFile << "\n";
                rcFile << "#pragma code_page(65001)\n";
                rcFile << "1 VERSIONINFO\n";
                rcFile << "FILEVERSION " << fileVerMs << "," << fileVerLs << ",0," << verRevision_ << "\n";
                rcFile << "PRODUCTVERSION " << prodVerMs << "," << prodVerLs << ",0," << verRevision_ << "\n";
                rcFile << "FILEFLAGSMASK 0x3fL\n";
                rcFile << "FILEFLAGS 0x0L\n";
                rcFile << "FILEOS 0x00040004L\n";
                rcFile << "FILETYPE 0x00000001L\n";
                rcFile << "FILESUBTYPE 0x00000000L\n";
                rcFile << "BEGIN\n";
                rcFile << "  BLOCK \"StringFileInfo\"\n";
                rcFile << "  BEGIN\n";
                rcFile << "    BLOCK \"080404b0\"\n";
                rcFile << "    BEGIN\n";
                rcFile << "      VALUE \"CompanyName\", \"" << escapeRc(company) << "\"\n";
                rcFile << "      VALUE \"FileDescription\", \"" << escapeRc(fileDesc) << "\"\n";
                rcFile << "      VALUE \"FileVersion\", \"" << fileVerMs << "." << fileVerLs << ".0." << verRevision_ << "\"\n";
                rcFile << "      VALUE \"InternalName\", \"" << escapeRc(internalName) << "\"\n";
                rcFile << "      VALUE \"LegalCopyright\", \"" << escapeRc(copyright) << "\"\n";
                rcFile << "      VALUE \"LegalTrademarks\", \"" << escapeRc(trademarks) << "\"\n";
                rcFile << "      VALUE \"OriginalFilename\", \"" << escapeRc(originalName) << "\"\n";
                rcFile << "      VALUE \"ProductName\", \"" << escapeRc(prodName) << "\"\n";
                rcFile << "      VALUE \"ProductVersion\", \"" << prodVerMs << "." << prodVerLs << ".0." << verRevision_ << "\"\n";
                if (!comments.empty()) {
                    rcFile << "      VALUE \"Comments\", \"" << escapeRc(comments) << "\"\n";
                }
                rcFile << "    END\n";
                rcFile << "  END\n";
                rcFile << "  BLOCK \"VarFileInfo\"\n";
                rcFile << "  BEGIN\n";
                rcFile << "    VALUE \"Translation\", 0x0804, 1200\n";
                rcFile << "  END\n";
                rcFile << "END\n";
            }
        }

        // Find rc.exe (共用发现处: 见 findRcExe)
        std::string rcExePath2 = findRcExe();

        if (!rcExePath2.empty()) {
            std::string verResPath = absInterDir2 + "\\version_info.res";
            std::ostringstream verRcArgs;
            verRcArgs << "\"" << rcExePath2 << "\" /r /fo \"" << verResPath << "\" \"" << verRcPath << "\"";
            if (options.verbose) {
                std::cout << "C3: RC (version): " << verRcArgs.str() << std::endl;
            }
            // 同 rcArgs: 必须多包一层引号, 否则 cmd 剥引号后 rc 路径被破坏 (见上文注释)
            int verRcRet = MsvcDriver::executeCommand("\"" + verRcArgs.str() + "\"");
            if (verRcRet == 0 && existsUtf8(verResPath)) {
                msvcOpts.versionInfoResFile = verResPath;
                if (options.verbose) {
                    std::cout << "C3: VS_VERSION_INFO resource compiled: " << verResPath << std::endl;
                }
            }
        } else {
            std::cerr << "C3: rc.exe not found, version info will not be embedded" << std::endl;
        }
    }

        // P20-41: 应用清单走**资源**而不是旁挂 <exe>.manifest 文件。
        // 旁挂形式有三个硬伤: ①构建系统不感知, 改清单不触发重编 ②复制/签名/分发时
        // 容易丢, 丢掉就静默退回 comctl32 v5.82 ③同目录多份 exe 会互相串。
        // 这里生成 `1 RT_MANIFEST` (资源 id 1 是激活上下文的约定入口) 编译进 exe。
        //
        // C29-M: 让位的判据从"工程给了 ResFile= 没有"改成"**那份 .res 里到底有没有 #1 清单**"。
        // 旧写法把两件事混成一件: 用户带一份只放图标/版本信息的 .res (VB6 里很常见) 就被当成
        // "他自己管清单", 于是内置那份一起让掉 ⇒ 产物**静默**退回 v5.82 (实测素材:
        // archive/vbman/src/RES/VBMANLIB.RES，5 条资源、零清单)。
        // 现在的口径 = 用户提供的清单**优先**，没有就**用内置的**；两份 #1 才是真不能发生的事
        // (加载器直接报错)，所以解析不出结论时**保守让位** —— 与旧行为一致，绝不叠加。
    {
        bool userHasManifest = false;
        if (!userResFile_.empty()) {
            ResFileProbe probe;
            if (probe.loadFromPath(userResFile_)) {
                userHasManifest = probe.hasAppManifest;
                // 走 stderr 的 `C3:` 行：note 级信息在成功的编译里不会整体打印出来，
                // stdout 那一路更不稳 ⇒ 判据要能在这轮构建里被 grep 到就得走这里。
                std::cerr << "C3: application manifest: "
                          << (userHasManifest
                                  ? "user-supplied .res carries #1 -> using theirs"
                                  : "user .res has no #1 (" + probe.why + ") -> injecting the built-in one")
                          << std::endl;
            } else {
                userHasManifest = true;    // 认不动 ⇒ 保守让位 (见上)
                std::cerr << "C3: warning: cannot parse ResFile (" << probe.why
                          << "); yielding to it instead of risking two #1 manifests. "
                          << "If the product looks like comctl32 v5, drop ResFile or add a manifest."
                          << std::endl;
            }
        }
        if (!options.isDll && !userHasManifest) {
        std::string absInterDir3 = pathToUtf8(std::filesystem::absolute(utf8ToPath(intermediatesDir)));
        std::string maniStem = projectBaseName_.empty() ? std::string("app") : projectBaseName_;
        std::string maniName = maniStem + ".c3.manifest";
        std::string maniPath = absInterDir3 + "\\" + maniName;
        std::string maniRcPath = absInterDir3 + "\\" + maniStem + ".c3.rc";
        {
            // assemblyIdentity 的 name 建议与应用同名, 这里用工程基名。
            std::string maniStemEsc = maniStem;
            std::ofstream maniFile = ofstreamUtf8(maniPath, std::ios::out | std::ios::trunc);
            if (maniFile) {
                maniFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
                maniFile << "<assembly xmlns=\"urn:schemas-microsoft-com:asm.v1\" manifestVersion=\"1.0\">\n";
                maniFile << "  <assemblyIdentity type=\"win32\" name=\"" << maniStemEsc
                         << "\" version=\"1.0.0.0\" processorArchitecture=\"*\" />\n";
                // comctl32 v6: 挂上之后 statusbar/toolbar 等通用控件才按 xp 之后的版本走
                // (注意: v6 也**不注册** msctls_status32, 它仍需 RTL 自注册, 见
                // vb6_StatusBar_RegisterClass —— 挂 v6 只解决主题与部分控件)。
                maniFile << "  <dependency>\n";
                maniFile << "    <dependentAssembly>\n";
                maniFile << "      <assemblyIdentity type=\"win32\" name=\"Microsoft.Windows.Common-Controls\""
                         << " version=\"6.0.0.0\" processorArchitecture=\"*\""
                         << " publicKeyToken=\"6595b64144ccf1df\" language=\"*\" />\n";
                maniFile << "    </dependentAssembly>\n";
                maniFile << "  </dependency>\n";
                maniFile << "</assembly>\n";
            }
        }

        // 首选: 链接完之后用 mt.exe 后注入 (见 embedManifestViaMt 的实测记录)。
        // 只有连 mt.exe 都找不着时才退回 rc.exe 交给链接器。
        std::string mtExePath = findMtExe();
        if (!mtExePath.empty() && existsUtf8(maniPath)) {
            msvcOpts.manifestXmlFile = maniPath;
            if (options.verbose) {
                std::cout << "C3: application manifest will be injected after link: " << maniPath << std::endl;
            }
        } else {
            std::ofstream maniRcFile = ofstreamUtf8(maniRcPath, std::ios::out | std::ios::trunc);
            if (maniRcFile) {
                maniRcFile << "#pragma code_page(65001)\n";
                maniRcFile << "1 RT_MANIFEST \"" << maniName << "\"\n";
            }

            std::string rcExePath3 = findRcExe();
            if (!rcExePath3.empty()) {
                std::string maniResPath = absInterDir3 + "\\" + maniStem + ".c3.res";
                std::ostringstream maniRcArgs;
                maniRcArgs << "\"" << rcExePath3 << "\" /r /fo \"" << maniResPath << "\" \"" << maniRcPath << "\"";
                if (options.verbose) {
                    std::cout << "C3: RC (manifest): " << maniRcArgs.str() << std::endl;
                }
                // 与上面两处同款: 整串再包一层引号, 否则 cmd /c 会把 rc.exe 路径剥坏。
                int maniRcRet = MsvcDriver::executeCommand("\"" + maniRcArgs.str() + "\"");
                if (maniRcRet == 0 && existsUtf8(maniResPath)) {
                    msvcOpts.manifestResFile = maniResPath;
                    if (options.verbose) {
                        std::cout << "C3: application manifest embedded: " << maniResPath << std::endl;
                    }
                    // 实测: rc.exe 这条路**恰恰是为 RT_MANIFEST 准备的写法不行** ——
                    //   ① `1 16 "x.manifest"` 会被 rc 判成 VERSIONINFO 语法错 RC2167
                    //      (RT_VERSION 也是 16), 任何数值型类型写法都被拒;
                    //   ② 唯一能过的 `1 RT_MANIFEST "x"` 被 rc 记成**字符串类型名**,
                    //      不是数值 24, 链接后资源目录里 type=88(0x58)、named=1,
                    //      Windows 激活上下文不会去找 → comctl32 停在 v5.82。
                    // 这跟 x86/x64 无关, 所以下面不再按 arch 分级告警: 走到这条退路
                    // 就已经说明机器上没有 mt.exe, 静默退回 v5.82 是可接受的,
                    // 但仍然喊一声(与 TypeLib/VERSIONINFO 两处一致, 不藏在 --verbose 后)。
                    //
                    // 存档一条被推翻过的错误结论, 免得以后有人再绕回来:x86 链接**并不是**
                    // 丢 .res。旧判断来自 tests/ctrlstatusbar 那套 PE 解析脚本把 PE32 的
                    // magic 写成了 0x107(正确是 0x10B), 于是所有 32 位产物的资源目录
                    // 一律被读成 0, 看着就像"没嵌入"。实测 32 位 link.exe 的 .rsrc 正常写入。
                    std::cerr << "C3: warning: manifest embedded via rc.exe, but that yields "
                              << "type=88 instead of RT_MANIFEST(24); comctl32 will stay v5.82. "
                              << "mt.exe was not found." << std::endl;
                }
            } else {
                // 与 TypeLib / VERSIONINFO 两处一致: 不藏在 --verbose 后面。
                std::cerr << "C3: rc.exe not found, application manifest will not be embedded" << std::endl;
            }
        }
    }
    }   // C29-M: 上面这段"要不要注内置清单"的判断整体收成一块，userHasManifest 不外泄

        // P23-03: Pass user .res file to linker
    if (!userResFile_.empty() && std::filesystem::exists(utf8ToPath(userResFile_))) {
        // (#44 rev2): x64 宿主 + 用户 manifest 的 x86 标记 = 进程初始化失败
        // (SSTabEx 实证, 见 normalizeUserResForArch 注释)。冲突时用归一化副本。
        std::string resForLink = userResFile_;
        std::string normRes = normalizeUserResForArch(userResFile_, options.arch,
                                                       intermediatesDir, options.verbose);
        if (!normRes.empty()) resForLink = normRes;
        msvcOpts.userResFile = resForLink;
        if (options.verbose) {
            std::cout << "C3: User resource file: " << resForLink << std::endl;
        }
    }

    // ai/vb-asm-extension-spec: Asm 块过程 → .asm → ml64 → .obj → 链接输入
    if (!asmProcs_.empty()) {
        if (!assembleAsmProcs(asmProcs_, intermediatesDir, msvcOpts, options.verbose)) {
            return false;
        }
    }

    MsvcDriver msvc;
    // 编译前自愈: 会话 rtl/ 下的 RTL 源文件可能已被兄弟进程的 cleanupOldSessions
    // 误删 (并发编译下 cl /MP 尚未处理的那些源文件会报 C1083). 缺什么补什么.
    if (!session.restoreMissing()) {
        std::cerr << "C3: error: RTL sources incomplete before compile (" << rtlDir << ")" << std::endl;
        return false;
    }
    if (!msvc.compileAndLink(msvcOpts)) return false;

    // 清单必须在链接成功之后注入 —— 它写的是刚产出的那个 exe。
    // (compileAndLink 内部会分派到增量路径, 所以这一侧不必再单独处理。)
    // 注意用 msvcOpts.outputFile 而不是 options.outputFile: 后者在没显式给 -o 时是空的,
    // 前者在 runLinker 里已经落到 「outputDir/<基名><扩展名>」这一档 (见 423 行起)。
    if (!msvcOpts.manifestXmlFile.empty()) {
        embedManifestViaMt(msvcOpts.outputFile, msvcOpts.manifestXmlFile, options.verbose);
    }

    return true;
}

} // namespace vb6c3
