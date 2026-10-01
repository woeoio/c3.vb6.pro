#include "backend/msvc_driver.hpp"
#include "common/encoding.hpp"

#include <iostream>
#include <fstream>
#include <sstream>
#include <filesystem>
#include <unordered_map>
#include <vector>
#include <string>
#include <cstdlib>

#ifdef _WIN32
#include <windows.h>
#endif

namespace vb6c3 {

namespace {

// opt3: FNV-1a 64位哈希 (跨运行稳定, 不依赖标准库实现)
std::string fnv1aHex(const void* data, size_t len) {
    uint64_t h = 1469598103934665603ULL;
    const unsigned char* p = static_cast<const unsigned char*>(data);
    for (size_t i = 0; i < len; ++i) {
        h ^= p[i];
        h *= 1099511628211ULL;
    }
    std::ostringstream oss;
    oss << std::hex << h;
    return oss.str();
}

std::string hashString(const std::string& s) {
    return fnv1aHex(s.data(), s.size());
}

// 文件内容哈希; 读取失败返回空串
std::string hashFile(const std::string& path) {
    std::ifstream f = ifstreamUtf8(path, std::ios::binary);
    if (!f) return "";
    uint64_t h = 1469598103934665603ULL;
    char buf[65536];
    while (f.read(buf, sizeof(buf)) || f.gcount() > 0) {
        std::streamsize n = f.gcount();
        const unsigned char* p = reinterpret_cast<const unsigned char*>(buf);
        for (std::streamsize i = 0; i < n; ++i) {
            h ^= p[i];
            h *= 1099511628211ULL;
        }
    }
    std::ostringstream oss;
    oss << std::hex << h;
    return oss.str();
}

// 解析 .c 文件中的本地头文件 #include "xxx.h"，返回绝对路径列表
// (opt3: 依赖跟踪 — 任一依赖头文件变化即触发该 .c 重编译)
// srcDir: 生成代码目录; rtlDir: RTL 源码目录 (释放的 vb6rtl.h 等)
std::vector<std::string> collectLocalIncludes(const std::string& cPath, const std::string& srcDir,
                                              const std::string& rtlDir) {
    std::vector<std::string> incs;
    std::ifstream f = ifstreamUtf8(cPath);
    if (!f) return incs;
    std::string line;
    while (std::getline(f, line)) {
        size_t hash = line.find('#');
        if (hash == std::string::npos) continue;
        size_t q1 = line.find('"', hash);
        if (q1 == std::string::npos) continue;
        size_t q2 = line.find('"', q1 + 1);
        if (q2 == std::string::npos) continue;
        std::string name = line.substr(q1 + 1, q2 - q1 - 1);
        if (name.size() < 3) continue;
        if (name.compare(name.size() - 2, 2, ".h") != 0) continue;
        // 绝对路径: 优先 srcDir (生成代码头), 不存在则查 rtlDir (RTL 头)
        std::string p = srcDir + "/" + name;
        if (!std::filesystem::exists(utf8ToPath(p)) && !rtlDir.empty()) {
            p = rtlDir + "/" + name;
        }
        incs.push_back(p);
    }
    return incs;
}

// === ai/030 T30-A: 内容寻址的 obj store ===
// 与旧方案 (cacheDir + cache.txt 索引) 的三条分别：
//  1) RTL 那一族 .c 用**固定**的一套编译档编：与用户 -O/-g 无关，且 include 路径只给
//     rtlDir。⇒ RTL obj 只由 {架构, 工具串, RTL 源与其本地 include 的内容} 决定，一份
//     缓存服务所有工程/所有产物形态；工程里放一个同名 vb6rtl.h 也抢不动 (旧路径的 /I 与
//     依赖解析都**先查 srcDir**，见 ai/030 §四)。
//  2) obj 文件名里带键哈希 ⇒ 没有索引文件，也就没有"交替编两个项目互相抹记录"、
//     "cache.txt trunc 非原子重写"这两件事。键不对就是查不到，查不到就自己编 ——
//     **错配的 obj 不可能被"用错"，只可能被"没用上"**，这是随包带缓存 (T30-D) 敢做的条件。
//  3) store 落在 %LOCALAPPDATA%\C3\objcache ⇒ 跨输出目录、跨项目共享。
std::string objStoreDir(const std::string& fallback) {
#ifdef _WIN32
    wchar_t buf[1024];
    DWORD n = GetEnvironmentVariableW(L"LOCALAPPDATA", buf,
                                      static_cast<DWORD>(sizeof(buf) / sizeof(buf[0])));
    if (n > 0 && n < sizeof(buf) / sizeof(buf[0])) {
        std::string base = pathToUtf8(std::filesystem::path(buf));
        if (!base.empty()) return base + "/C3/objcache";
    }
#endif
    return fallback;
}

// 工具串: cl.exe 全路径自带 VC 版本号 (.../MSVC/14.xx.yyyyy/bin/Hostx64/x64/cl.exe);
// 调用方环境里有 INCLUDE 时, 连 Windows SDK 的版本目录一起进 (dev shell 里跑 / 做完
// T30-C 每进程捕获一次 vcvars 环境之后, 这条总是成立)。
std::string toolsetTag(const std::string& clPath) {
    const char* inc = std::getenv("INCLUDE");
    return clPath + "|" + (inc ? inc : "");
}

// 一个 .c 真正生效的编译档。**路径类**选项一律不进键 —— 会话临时目录每次编译都换,
// 进了就永远不命中; 那些内容 (RTL 源/头) 本来就有独立哈希盖着。
std::string flagsKeyFor(bool isRtl, const MsvcDriverOptions& o) {
    std::string s = std::string(isRtl ? "rtl-fixed" : "user") + "|" + o.arch;
    if (isRtl) s += "|/Od";   // RTL 固定档: 用户的 -O 与 -g 都不改它
    else {
        s += "|/O" + std::to_string(o.optimizationLevel);
        if (o.debugInfo) s += "|/Zi";
    }
    if (o.arch == "x86") s += "|/MT";
    // C29-V6: 版本宏会改**声明面**(commctrl.h 在 _WIN32_IE>=0x0600 下给出 v6 的
    // TBBUTTONINFOW/类名宏) ⇒ 产物不同，必须进键，否则缓存会把 v5 声明下编出的 obj
    // 端给 v6 声明的编译 (ai/029 C29-5c 那条实测的正是这一族)。
    s += "|ver=" C3_V6_VERSION_DEFS;
    return s;
}

// 真正给 cl 的编译档 (与 flagsKeyFor 一一对应, 只差 /I 的路径)
std::string compileFlagsFor(bool isRtl, const MsvcDriverOptions& o,
                            const std::string& rtlDir, const std::string& srcDir) {
    std::ostringstream c;
    if (isRtl) {
        if (!rtlDir.empty()) c << " /I\"" << rtlDir << "\"";
        c << " /Od";
    } else {
        if (!rtlDir.empty()) c << " /I\"" << rtlDir << "\"";
        if (!srcDir.empty()) c << " /I\"" << srcDir << "\"";
        switch (o.optimizationLevel) {
            case 0: c << " /Od"; break;
            case 1: c << " /O1"; break;
            case 2: c << " /O2"; break;
            case 3: c << " /Ox"; break;
        }
        if (o.debugInfo) c << " /Zi";
    }
    c << " /std:c11 /DUNICODE /D_UNICODE /utf-8 /D_CRT_SECURE_NO_WARNINGS"
         " /D_CRT_NONSTDC_NO_WARNINGS /W3";
    // C29-V6: 与 msvc_driver.cpp 那条非增量路同一套版本宏 —— 两边都得有，否则"增量编出来
    // 的产物"与"非增量编出来的"声明面不同。上面 flagsKeyFor 进键的就是这一串。
    c << C3_V6_VERSION_DEFS;
    // P10 恢复: RTL 源码直接编译, /Gy 函数级链接配合 /OPT:REF 剔除未引用 RTL 代码
    c << " /Gy";
    if (o.arch == "x86") c << " /MT";
    return c.str();
}

// src 是否属于 RTL 那一族 (由 driver 从 exe 里解包出来的目录)
bool isRtlSource(const std::string& src, const std::string& rtlDir) {
    if (rtlDir.empty()) return false;
    if (src.compare(0, rtlDir.size(), rtlDir) != 0) return false;
    return src.size() > rtlDir.size() && (src[rtlDir.size()] == '/' || src[rtlDir.size()] == '\\');
}

} // namespace

// opt3: 增量编译 — 基于内容哈希的 obj 级缓存。
// 每个 .c 的缓存记录 = 源文件哈希 + 依赖头文件哈希 + 编译选项指纹。
// 命中且缓存 obj 存在 → 跳过 cl 编译，直接复用 obj 链接。
bool MsvcDriver::compileAndLinkIncremental(const MsvcDriverOptions& options) {
    std::string cl = findClExe();
    std::string arch = options.arch;

    // obj 输出目录 (新编译的 obj 先落在这里, 链接成功后复制到缓存目录)
    std::string objDir = options.objDir;
    if (objDir.empty() && !options.outputFile.empty()) {
        std::filesystem::path outP(utf8ToPath(options.outputFile));
        objDir = pathToUtf8(outP.parent_path());
    }
    if (objDir.empty()) objDir = ".";
    std::string tmpLogPath = objDir + "/_c3_msvc_out.txt";

    // 缓存目录: 全局 store (%LOCALAPPDATA%\C3\objcache), 拿不到就退回 <outputDir>/.c3obj
    std::string cacheDir = options.incrementalCacheDir;
    if (cacheDir.empty()) cacheDir = objDir;
    std::string store = objStoreDir(cacheDir);
    std::error_code ec;
    std::filesystem::create_directories(utf8ToPath(store), ec);
    const std::string toolset = toolsetTag(cl);

    // 两批: RTL 组用固定档 (跨工程共享一份), 用户码组跟着 -O/-g 走
    std::vector<std::string> reusedObjs;   // 命中的 obj = store 里的路径
    struct Pending { std::string src, objPath, storePath; };
    std::vector<Pending> rtlTodo, userTodo;
    int rtlHits = 0, userHits = 0;

    for (const auto& src : options.sourceFiles) {
        std::filesystem::path sp(utf8ToPath(src));
        std::string stem = pathToUtf8(sp.stem());
        bool isRtl = isRtlSource(src, options.rtlDir);
        // RTL 组的本地 include 只在 rtlDir 内解析: 工程里放同名头抢不动 (ai/030 §四)
        auto incs = collectLocalIncludes(src, isRtl ? std::string() : options.srcDir,
                                         options.rtlDir);
        std::string deps;
        for (auto& inc : incs) deps += hashFile(inc);
        std::string srcHash = hashFile(src);
        std::string key = srcHash.empty()
            ? std::string()
            : hashString(flagsKeyFor(isRtl, options) + "|" + toolset + "|" + srcHash +
                         "|" + hashString(deps) + "|" + stem);
        std::string storePath = key.empty() ? std::string()
                                            : store + "/" + stem + "_" + key + ".obj";
        std::error_code kec;
        if (!storePath.empty() && std::filesystem::exists(utf8ToPath(storePath), kec)) {
            reusedObjs.push_back(storePath);
            if (isRtl) rtlHits++; else userHits++;
            continue;
        }
        std::string foDir = objDir + (isRtl ? "/rtl" : "/usr");
        Pending p{src, foDir + "/" + stem + ".obj", storePath};
        (isRtl ? rtlTodo : userTodo).push_back(p);
    }
    const int totalSrcs = static_cast<int>(options.sourceFiles.size());
    const int hitCount = rtlHits + userHits;

    // ASCII 读数: 命中面走这条 (可被断言), 中文那条留给人看
    std::cerr << "C3: OBJCACHE rtl=" << rtlHits << "/" << (rtlHits + static_cast<int>(rtlTodo.size()))
              << " user=" << userHits << "/" << (userHits + static_cast<int>(userTodo.size()))
              << " total=" << hitCount << "/" << totalSrcs << std::endl;

    if (options.verbose) {
        std::cout << "C3: 增量编译: " << hitCount << "/" << totalSrcs
                  << " 个源文件命中缓存, 跳过编译 (store: " << store << ")" << std::endl;
    }

    // === 编译需要重编的 .c —— 按组各一次: 两组的编译档不同, 不能共用一条 cl 命令 ===
    auto runBatch = [&](const std::vector<Pending>& todo, const std::string& foDir,
                        bool isRtl) -> bool {
        if (todo.empty()) return true;
        std::ostringstream compileCmd;
        compileCmd << cl << compileFlagsFor(isRtl, options, options.rtlDir, options.srcDir)
                   << " /MP /c /Fo\"" << foDir << "/\"";
        for (const auto& t : todo) compileCmd << " \"" << t.src << "\"";
        std::error_code mkec;
        std::filesystem::create_directories(utf8ToPath(foDir), mkec);
        std::string rspPath = objDir + (isRtl ? "/_c3_cl_rtl.rsp" : "/_c3_cl_user.rsp");
        // Fix 196: UTF-16LE+BOM (cl/link 按系统 ANSI 代码页读 @rsp, 见 msvc_driver.hpp)
        writeMsvcResponseFile(rspPath, compileCmd.str().substr(cl.length()));
        std::string vcvarsPrefix = buildVcvarsPrefix(arch);
        // Fix <vbeclipse> D8050: 同 msvc_driver.cpp —— 每次 cl 走独立 TMP/TEMP
        // (objDir = 本次 session 目录), 断掉同 runner 并行 GUI worker 的
        // c1.exe 在共享 %TEMP% 上互踩 _CL_*.tmp 那条路。
        std::string tmpIso1 = "set \"TMP=" + objDir + "\" && set \"TEMP=" + objDir + "\" && ";
        std::string fullCmd = vcvarsPrefix + tmpIso1 + cl + " @\"" + rspPath + "\" > \"" + tmpLogPath + "\" 2>&1";
        int ret = executeCommand(fullCmd);
        std::filesystem::remove(rspPath, std::error_code());
        if (ret != 0) {
            std::string outputDirForLog;
            if (!options.outputFile.empty()) {
                std::filesystem::path outP(utf8ToPath(options.outputFile));
                outputDirForLog = pathToUtf8(outP.parent_path());
            }
            if (outputDirForLog.empty()) outputDirForLog = ".";
            std::string errorLogPath = outputDirForLog + "/c3-error.log";
            std::ifstream tmpLog = ifstreamUtf8(tmpLogPath);
            std::ofstream errLog = ofstreamUtf8(errorLogPath, std::ios::out | std::ios::trunc);
            if (tmpLog && errLog) {
                errLog << "C3: Compilation failed (exit code " << ret << ")" << std::endl;
                errLog << "=== MSVC Output ===" << std::endl;
                std::string line;
                while (std::getline(tmpLog, line)) errLog << line << "\n";
            }
            if (tmpLog) {
                tmpLog.clear();
                tmpLog.seekg(0);
                std::string line;
                while (std::getline(tmpLog, line)) std::cerr << line << std::endl;
            }
            std::cerr << "C3: 编译失败 (exit code " << ret << ")" << std::endl;
            std::cerr << "C3: 错误日志已保存: " << errorLogPath << std::endl;
            std::filesystem::remove(tmpLogPath, std::error_code());
            return false;
        }
        // 原子进 store: 同目录先拷临时名再改名 (同一分区, 不跨卷)。
        // 键就是内容哈希 ⇒ 并发放同一格是写同样字节, 谁先到位都算对 (ai/030 §六 T30-A)。
        const std::string pidTag = std::to_string(static_cast<unsigned long>(GetCurrentProcessId()));
        for (const auto& t : todo) {
            if (t.storePath.empty()) continue;
            std::error_code iec;
            if (!std::filesystem::exists(utf8ToPath(t.objPath), iec)) continue;
            if (std::filesystem::exists(utf8ToPath(t.storePath), iec)) continue;
            std::string tmp = t.storePath + "." + pidTag + ".tmp";
            std::error_code cec;
            std::filesystem::copy_file(utf8ToPath(t.objPath), utf8ToPath(tmp),
                                       std::filesystem::copy_options::overwrite_existing, cec);
            if (cec) continue;
            std::filesystem::rename(utf8ToPath(tmp), utf8ToPath(t.storePath), cec);
            if (cec) std::filesystem::remove(utf8ToPath(tmp), std::error_code());
        }
        return true;
    };

    std::vector<std::string> newObjs;   // 本批新编出来的 obj (链接输入)
    bool built = runBatch(rtlTodo, objDir + "/rtl", true);
    if (built) built = runBatch(userTodo, objDir + "/usr", false);
    for (const auto& t : rtlTodo)  newObjs.push_back(t.objPath);
    for (const auto& t : userTodo) newObjs.push_back(t.objPath);
    std::filesystem::remove(tmpLogPath, std::error_code());
    if (!built) return false;

    // === 链接 (所有 obj: 复用的 + 新编译的) ===
    // 注意: 拆分为两步后编译阶段无源文件, cl /link 不会进入链接模式(D8003),
    // 因此链接阶段直接调用 link.exe (vcvarsall 后位于 PATH)。
    std::string linkExe = "link.exe";
    std::ostringstream linkCmd;
    linkCmd << linkExe << " /NOLOGO";
    if (options.isDll) {
        linkCmd << " /DLL";
    } else if (options.isGui) {
        // Fix 165: 与 msvc_driver.cpp 同步 — /ENTRY 取决于实际生成的入口函数
        linkCmd << " /SUBSYSTEM:WINDOWS";
        if (options.entryIsMain) linkCmd << " /ENTRY:mainCRTStartup";
    } else {
        linkCmd << " /SUBSYSTEM:CONSOLE";
    }
    if (!options.outputFile.empty()) linkCmd << " /OUT:\"" << options.outputFile << "\"";
    for (auto& o : reusedObjs) linkCmd << " \"" << o << "\"";
    for (auto& o : newObjs) linkCmd << " \"" << o << "\"";
    // ai/vb-asm-extension-spec: ml64 汇编产物 (.obj) 也是链接输入
    appendExtraObjects(linkCmd, options);
    if (options.isDll) {
        if (!options.typelibResFile.empty()) linkCmd << " \"" << options.typelibResFile << "\"";
        if (!options.versionInfoResFile.empty()) linkCmd << " \"" << options.versionInfoResFile << "\"";
        if (!options.userResFile.empty()) linkCmd << " \"" << options.userResFile << "\"";
        if (!options.defFile.empty()) linkCmd << " /DEF:\"" << options.defFile << "\"";
        linkCmd << " ole32.lib oleaut32.lib uuid.lib advapi32.lib user32.lib shell32.lib gdi32.lib";
    } else if (options.isGui) {
        if (!options.typelibResFile.empty()) linkCmd << " \"" << options.typelibResFile << "\"";
        if (!options.versionInfoResFile.empty()) linkCmd << " \"" << options.versionInfoResFile << "\"";
        if (!options.manifestResFile.empty()) linkCmd << " \"" << options.manifestResFile << "\"";
        if (!options.userResFile.empty()) linkCmd << " \"" << options.userResFile << "\"";
        linkCmd << " user32.lib gdi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib advapi32.lib comctl32.lib";
    } else {
        if (!options.typelibResFile.empty()) linkCmd << " \"" << options.typelibResFile << "\"";
        if (!options.versionInfoResFile.empty()) linkCmd << " \"" << options.versionInfoResFile << "\"";
        if (!options.manifestResFile.empty()) linkCmd << " \"" << options.manifestResFile << "\"";
        if (!options.userResFile.empty()) linkCmd << " \"" << options.userResFile << "\"";
        linkCmd << " ole32.lib oleaut32.lib uuid.lib advapi32.lib user32.lib shell32.lib gdi32.lib";
    }
    // ai/024 T02: 用户静态库 (归档) 的搜索根与库文件本体 — 与 msvc_driver.cpp 的
    // cl /link 路径共用同一个追加函数, 保证增量/非增量两条路径的链接输入一致。
    appendUserLibInputs(linkCmd, options);
    if (options.arch == "x86") linkCmd << " /MACHINE:X86";
    // P24-09 把 x86 的 RTL 与用户码都按 /MT 编 (obj 里带 /DEFAULTLIB:LIBCMT)，而这条链接路
    // 是直接叫 link.exe —— 不做处理时 linker 自己按默认的 /MD 档去配 msvcrt.lib，于是
    // LNK4098 (msvcrt 与 libcmt 冲突) + LNK2019 找不到 __except_handler4_common (SEH4 是 x86
    // 独有，x64 不报)。默认路径经 cl /link 链接，cl 会替我们把这条配平 —— 所以只有增量路会翻，
    // 而 --incremental 在 ai/030 T30-B 之前从没进过任何回归，这个洞就一直没被踩过。见 030 §10.7。
    if (options.arch == "x86") linkCmd << " /NODEFAULTLIB:msvcrt.lib";
    if (options.debugInfo) linkCmd << " /DEBUG /MAP";

    std::string linkRsp = objDir + "/_c3_link_args.rsp";
    // Fix 196: UTF-16LE+BOM (link.exe 直调 @rsp 同样按 ANSI 代码页读, 实测 UTF-8 → LNK1117)
    writeMsvcResponseFile(linkRsp, linkCmd.str().substr(linkExe.length()));
    std::string vcvarsPrefix2 = buildVcvarsPrefix(arch);
    // Fix <vbeclipse> D8050: 链接同样按 session objDir 隔离 TMP/TEMP, 与 cl 一致。
    std::string tmpIso2 = "set \"TMP=" + objDir + "\" && set \"TEMP=" + objDir + "\" && ";
    std::string fullLinkCmd = vcvarsPrefix2 + tmpIso2 + linkExe + " @\"" + linkRsp + "\" > \"" + tmpLogPath + "\" 2>&1";

    if (options.verbose) {
        std::cout << "C3: 执行: " << fullLinkCmd << std::endl;
    }

    int ret2 = executeCommand(fullLinkCmd);
    if (ret2 != 0) {
        std::string outputDirForLog;
        if (!options.outputFile.empty()) {
            std::filesystem::path outP(utf8ToPath(options.outputFile));
            outputDirForLog = pathToUtf8(outP.parent_path());
        }
        if (outputDirForLog.empty()) outputDirForLog = ".";
        std::string errorLogPath = outputDirForLog + "/c3-error.log";
        std::ifstream tmpLog = ifstreamUtf8(tmpLogPath);
        std::ofstream errLog = ofstreamUtf8(errorLogPath, std::ios::out | std::ios::trunc);
        if (tmpLog && errLog) {
            errLog << "C3: Compilation failed (exit code " << ret2 << ")" << std::endl;
            errLog << "=== MSVC Output ===" << std::endl;
            std::string line;
            while (std::getline(tmpLog, line)) errLog << line << "\n";
        }
        if (tmpLog) {
            tmpLog.clear();
            tmpLog.seekg(0);
            std::string line;
            while (std::getline(tmpLog, line)) std::cerr << line << std::endl;
        }
        std::cerr << "C3: 编译失败 (exit code " << ret2 << ")" << std::endl;
        std::cerr << "C3: 错误日志已保存: " << errorLogPath << std::endl;
        std::filesystem::remove(tmpLogPath, std::error_code());
        std::filesystem::remove(linkRsp, std::error_code());
        return false;
    }
    std::filesystem::remove(tmpLogPath, std::error_code());
    std::filesystem::remove(linkRsp, std::error_code());

    if (options.verbose) {
        std::cout << "C3: 编译成功: " << options.outputFile << std::endl;
    }
    return true;
}

} // namespace vb6c3
