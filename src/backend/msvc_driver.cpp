#include "backend/msvc_driver.hpp"
#include "common/encoding.hpp"

#include <iostream>
#include <fstream>
#include <cstdlib>
#include <cstdio>
#include <sstream>
#include <filesystem>
#include <unordered_map>

#ifdef _WIN32
#include <windows.h>
#endif

namespace vb6c3 {

// Fix 210b: 保证 cl.exe/link.exe 子进程拿到一个有效的 Windows 临时目录。
// Git Bash / MSYS 会把 TEMP/TMP 导出成 POSIX 路径 (如 /tmp), cl 的 c1.dll
// 用 GetTempPath 解析后会报 fatal error D8050/D8040 ("内存不能为read" 同类)。
//
// 2026-10-01 收紧: 只在 TEMP 真是 POSIX 风格时才构造覆盖块。原来无论好坏都覆写
// 一份新块传给 CreateProcessW, 一旦块里有什么错位 (比如漏了 CREATE_UNICODE_ENVIRONMENT,
// 或者 GetEnvironmentStringsW 回来的块里带 `=C:` 那类只在父 PEB 里有意义的条目),
// 就会污染 cl/c1 的 env, 让它们**在合法 Windows Temp 目录里创建 _CL_*.tmp 都失败**
// (本地实测 C1083 Permission denied)。不覆写时子进程直接从父 PEB 继承环境, 与
// 未上 Fix 210b 之前一致, 也就不引入新风险。
namespace {
    // 判当前 TEMP 是不是 POSIX 风格 (以 '/' 开头, 无盘符)。是的话才要修。
    bool tempLooksPosix() {
        wchar_t buf[MAX_PATH + 2] = { 0 };
        DWORD n = GetEnvironmentVariableW(L"TEMP", buf, MAX_PATH + 1);
        if (n == 0 || n > MAX_PATH) {
            n = GetEnvironmentVariableW(L"TMP", buf, MAX_PATH + 1);
            if (n == 0 || n > MAX_PATH) return false;
        }
        if (n == 0) return false;
        // Windows 路径: 有 "X:" 前缀 (X 是任意字母)。POSIX 路径: '/' 开头。
        if (buf[0] == L'\\') return true;                 // \\server\share 少见但非 POSIX
        if (n >= 2 && buf[1] == L':') return false;       // 盘符: 合法 Windows
        return true;                                       // 兜底当作需要修
    }

    // 取一个合法的 Windows 临时目录: 优先 GetTempPathW, 回退到用户本地 Temp。
    std::wstring windowsTempPath() {
        wchar_t buf[MAX_PATH + 2] = { 0 };
        DWORD n = GetTempPathW(MAX_PATH + 1, buf);
        if (n > 0 && n <= MAX_PATH + 1) {
            std::wstring s = buf;
            // 必须是含盘符的 Windows 路径 (绝不可能是 POSIX /tmp 之类)
            if (s.size() >= 3 && s[1] == L':' && s[2] == L'\\') {
                return s;
            }
        }
        return std::wstring(L"C:\\Users\\Administrator\\AppData\\Local\\Temp\\");
    }

    // 构造环境块: 复制当前 env, 剔除旧 TEMP/TMP, 追加修正后的 TEMP/TMP。
    // 返回以双 \0 结尾的宽字符块 —— CreateProcessW 侧要一起带
    // CREATE_UNICODE_ENVIRONMENT 标志 (见 executeCommand 里的调用)。
    std::wstring buildEnvBlockWithWindowsTemp() {
        std::wstring block;
        std::wstring tmp = windowsTempPath();
        wchar_t* env = GetEnvironmentStringsW();
        if (env) {
            for (wchar_t* p = env; *p; ) {
                size_t len = wcslen(p);
                // 剔除现有 TEMP=/TMP= 以便覆盖
                if (_wcsnicmp(p, L"TEMP=", 5) != 0 &&
                    _wcsnicmp(p, L"TMP=", 4) != 0) {
                    block.append(p, len);
                    block.push_back(L'\0');
                }
                p += len + 1;
            }
            FreeEnvironmentStringsW(env);
        }
        block += L"TEMP="; block += tmp; block += L'\0';
        block += L"TMP=";  block += tmp; block += L'\0';
        block += L'\0'; // 块结束符
        return block;
    }

    // 只在需要时才建块; 否则返回空串, 调用方 lpEnv 传 nullptr 让子进程直接继承父 env。
    std::wstring maybeEnvBlockWithWindowsTemp() {
        if (!tempLooksPosix()) return std::wstring();
        return buildEnvBlockWithWindowsTemp();
    }

    // ======================================================================
    // §56.53: cl 命令行拆成三段 (公共编译档 / 输出档 / 链接档)
    //
    // 为什么要拆: PCH 路径下要跑**好几段** cl —— 建 PCH 一段、按前奏档分组编译
    // 一到两段、编 RTL 等其余 TU 一段、链接一段。而「一次 cl 编译+链接」的老路
    // 必须**逐字不变** (它是 PD 整编与所有回归的既有形状), 所以这里不是改写
    // 老路, 而是**抽出**共用段: 老路仍按原顺序拼一遍, PCH 路径按段拼。
    //
    // 公共编译档只有一份的原因: 建 PCH 的那次 cl 与消费它的那些 cl **必须同档**
    // —— 档不一样 (哪怕只差一个 /D 或 /I) ⇒ PCH 里存的预处理状态与 TU 自己解析
    // 出来的状态不同构 ⇒ cl 拒收, 或者更糟: 悄悄编出别的东西。
    // ======================================================================

    // 公共编译档 (与 PCH 无关的那部分: include 路径 / 优化档 / 调试档 / 语言档)
    std::string buildMsvcCompileFlags(const MsvcDriverOptions& options) {
        std::ostringstream cmd;

        // 包含路径: RTL目录
        if (!options.rtlDir.empty()) {
            cmd << " /I\"" << options.rtlDir << "\"";
        }
        if (!options.srcDir.empty()) {
            cmd << " /I\"" << options.srcDir << "\"";
        }

        // 优化级别
        switch (options.optimizationLevel) {
            case 0: cmd << " /Od"; break;
            case 1: cmd << " /O1"; break;
            case 2: cmd << " /O2"; break;
            case 3: cmd << " /Ox"; break;
        }

        // 调试信息
        if (options.debugInfo) {
            cmd << " /Zi /FS /DEBUG";
            // Fix <vbeclipse>: /Zi 无 /Fd 时 PDB 落在 **cl 的工作目录** (= 工程目录)
            // 的 vc140.pdb — 该目录常被防护软件按目录拦截新文件创建 ⇒ C1041 全灭
            // (ComCtlsDemo -g 实测)。把 PDB 钉到 exe 的输出目录 (与 /Fe 同一目录)。
            {
                // A bare directory inside quotes must not end with a separator: in a
                // response file that trailing backslash escapes the closing quote, so cl
                // swallows every following argument (D8038 "unrecognized source file
                // type" with the whole command line as one argument). Pin a concrete PDB
                // file name instead -- /Fd takes a pathname, and the stem keeps concurrent
                // builds from sharing one compiler PDB.
                std::string outPath = options.outputFile;
                size_t slash = outPath.find_last_of("/\\");
                std::string pdbDir = (slash != std::string::npos) ? outPath.substr(0, slash + 1) : std::string();
                std::string pdbStem = (slash != std::string::npos) ? outPath.substr(slash + 1) : outPath;
                size_t dot = pdbStem.find_last_of('.');
                if (dot != std::string::npos) pdbStem = pdbStem.substr(0, dot);
                cmd << " /Fd\"" << pdbDir << pdbStem << "_cl.pdb\"";
            }
        }

        // C11标准, Unicode, UTF-8源码编码, 禁用MSVC安全警告
        cmd << " /std:c11 /DUNICODE /D_UNICODE /utf-8 /D_CRT_SECURE_NO_WARNINGS /D_CRT_NONSTDC_NO_WARNINGS";
        // C29-V6: 声明面与运行面统一成 comctl v6（理由与取值见 msvc_driver.hpp 那条注释）。
        cmd << C3_V6_VERSION_DEFS;

        // 警告级别
        cmd << " /W3";

        // P10 恢复: RTL 源码与生成代码一起编译 (非 .lib 按需拉取),
        // /Gy 启用函数级链接, 链接器 /OPT:REF 可按函数剔除未引用的 RTL 代码
        cmd << " /Gy";

        // 性能优化: 多处理器并行编译 (cl.exe /MP, 默认进程数=CPU核心数)
        // 大型项目(如vbman 389个.c文件)串行编译耗时极长, /MP 让每个源文件
        // 由独立 cl 进程并行编译。
        // 账 §56.53 更正: 旧注记说「/MP 与 /GL、/Yc、/Yu 不兼容」—— **实测不成立**。
        // 真 MSVC 14.29.30159 实测矩阵 (本机, .temp/pch_compat_test*.ps1):
        //   /MP + /Yc            → rc=0, .pch 正常生成
        //   /MP + /Yu + /FI      → rc=0, PCH 状态真被套用 (哨兵宏判据)
        //   /MP + /Yu + /FI + /Fo → rc=0
        // ⇒ §56.53 保留了 /MP, 不必自己切 N 份起 N 个 cl 进程。
        cmd << " /MP";

        // P24-09: x86 RTL libraries are built with /MT (static CRT); match it to avoid LNK2038
        if (options.arch == "x86") {
            // 账 #257 第十八层后: C3_DEBUG_CRT=1 ⇒ 静态调试 CRT (/MTd + /D_DEBUG), 让 CRT 调试堆
            // (_CRTDBG_CHECK_ALWAYS_DF + RTL report hook) 抓堆越界/坏释放。默认仍 /MT。
            cmd << (getenv("C3_DEBUG_CRT") ? " /MTd /D_DEBUG" : " /MT");
        }

        return cmd.str();
    }

    // 输出档: /Fe (产物名) + /Fo (中间件目录)。/Fo 末尾用**正斜杠** ——
    // 尾反斜杠在响应文件里会把闭引号吃掉 (实测 D8003「缺少源文件名」)。
    std::string buildMsvcOutputFlags(const MsvcDriverOptions& options) {
        std::ostringstream cmd;
        if (!options.outputFile.empty()) {
            // M22-IssueB: UTF-8 path preserved through CreateProcessW UTF-16 conversion
            cmd << " /Fe\"" << options.outputFile << "\"";
            if (!options.objDir.empty()) {
                cmd << " /Fo\"" << options.objDir << "/\"";
            } else {
                std::filesystem::path outPath(utf8ToPath(options.outputFile));
                std::string objDir = pathToUtf8(outPath.parent_path());
                if (!objDir.empty()) {
                    cmd << " /Fo\"" << objDir << "/\"";
                }
            }
        }
        return cmd.str();
    }

    // 链接档 (/link 之后的全部)。与编译档分开, 是为了让 PCH 路径能把「只编不链」
    // 的那几段摘出去 (老路仍然把两段拼在一起, 命令行一字不变)。
    std::string buildMsvcLinkOptions(const MsvcDriverOptions& options) {
        std::ostringstream cmd;
        if (options.isDll) {
            // P6.6: ActiveX DLL链接
            cmd << " /link /DLL";
            if (!options.typelibResFile.empty()) {
                cmd << " \"" << options.typelibResFile << "\"";
            }
            if (!options.versionInfoResFile.empty()) {
                cmd << " \"" << options.versionInfoResFile << "\"";
            }
            if (!options.userResFile.empty()) {
                cmd << " \"" << options.userResFile << "\"";
            }
            if (!options.defFile.empty()) {
                cmd << " /DEF:\"" << options.defFile << "\"";
            }
            cmd << " ole32.lib oleaut32.lib uuid.lib advapi32.lib user32.lib shell32.lib gdi32.lib";
        } else if (options.isGui) {
            // P7: GUI程序 (Win32窗口)
            // Fix 165: /ENTRY 取决于代码生成实际发出的入口 (窗体启动=WinMain / Sub Main=main)
            cmd << " /link /SUBSYSTEM:WINDOWS";
            if (options.entryIsMain) cmd << " /ENTRY:mainCRTStartup";
            if (!options.typelibResFile.empty()) {
                cmd << " \"" << options.typelibResFile << "\"";
            }
            if (!options.versionInfoResFile.empty()) {
                cmd << " \"" << options.versionInfoResFile << "\"";
            }
            if (!options.manifestResFile.empty()) {
                cmd << " \"" << options.manifestResFile << "\"";
            }
            if (!options.userResFile.empty()) {
                cmd << " \"" << options.userResFile << "\"";
            }
            cmd << " user32.lib gdi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib advapi32.lib comctl32.lib";
        } else {
            // 控制台程序
            cmd << " /link /SUBSYSTEM:CONSOLE";
            if (!options.typelibResFile.empty()) {
                cmd << " \"" << options.typelibResFile << "\"";
            }
            if (!options.versionInfoResFile.empty()) {
                cmd << " \"" << options.versionInfoResFile << "\"";
            }
            if (!options.manifestResFile.empty()) {
                cmd << " \"" << options.manifestResFile << "\"";
            }
            if (!options.userResFile.empty()) {
                cmd << " \"" << options.userResFile << "\"";
            }
            cmd << " ole32.lib oleaut32.lib uuid.lib advapi32.lib user32.lib shell32.lib gdi32.lib";
        }

        // ai/024 T02: 用户静态库 (归档) 的搜索根与库文件本体。
        // 放在三个分支之外单点追加 —— 三种工程类型 (DLL/GUI/控制台) 的库列表不同,
        // 但静态库这一项对三者是同一件事。
        appendUserLibInputs(cmd, options);

        // P-debug: 调试构建顺带输出 .map, 便于 crashctx.py 符号化崩溃现场
        // 注: 不可用 /MAPINFO:LINES — 本工程 link.exe (14.29.30159) 仅支持
        //     /MAPINFO:{EXPORTS|PDATA}, 加了会 LNK1117 syntax error.
        //     精确到"生成 C 行号"请改用 dbghelp+PDB 或 cl /FAcs 反汇编清单.
        if (options.debugInfo) cmd << " /MAP";

        // DualArch: /MACHINE flag for x86 target (x64 is default, no explicit flag needed)
        if (options.arch == "x86") {
            cmd << " /MACHINE:X86";
        }

        // czUI fix (exe 体积): 与 msvc_driver_incremental.cpp 同步 —— /Gy 编出的函数级
        // COMDAT 要靠 /OPT:REF 剔除未引用 RTL (winsock/data/ole/...), /OPT:ICF 折叠重复。
        cmd << " /OPT:REF /OPT:ICF /MAP";

        return cmd.str();
    }

    // 一个 .c 编译出来的 .obj 在哪 (cl 的规则: /Fo 目录 + 源文件名去扩展名)
    std::string msvcObjPathFor(const std::string& src, const std::string& objDir) {
        std::filesystem::path p(utf8ToPath(src));
        std::string stem = pathToUtf8(p.stem());
        return objDir + "/" + stem + ".obj";
    }

    // 跑一次 cl: 写 rsp → 执行 → 失败时落 c3-error.log (命令 + rsp + cl 输出)。
    // §56.53: PCH 路径要跑好几段, 每段一份 rsp 与一份日志 (否则后一段会把前一段的
    // 现场覆盖掉, 报错时只剩最后一段 —— 离线无法定位是哪一段炸的)。
    bool runMsvcInvocation(const std::string& cl, const std::string& args,
                            const std::string& rspPath, const std::string& tmpLogPath,
                            const MsvcDriverOptions& options,
                            const std::string& vcvarsPrefix, const std::string& tmpIsolation,
                            bool appendErrorLog = false, bool keepLogOnSuccess = false) {
        // Fix 196: 必须写成 UTF-16LE+BOM —— cl/link 按系统 ANSI 代码页读 @rsp,
        // 原样落 UTF-8 会把中文路径解成乱码 (详见 msvc_driver.hpp 的实测矩阵)。
        writeMsvcResponseFile(rspPath, args);
        std::string fullCmd = vcvarsPrefix + tmpIsolation + cl + " @\"" + rspPath + "\" > \"" + tmpLogPath + "\" 2>&1";

        int ret = MsvcDriver::executeCommand(fullCmd);
        if (ret != 0) {
            // c3-error.log goes to output dir (user project dir), not intermediates
            std::string outputDirForLog;
            if (!options.outputFile.empty()) {
                std::filesystem::path outP(utf8ToPath(options.outputFile));
                outputDirForLog = pathToUtf8(outP.parent_path());
            }
            if (outputDirForLog.empty()) outputDirForLog = ".";
            std::string errorLogPath = outputDirForLog + "/c3-error.log";
            std::ifstream tmpLog = ifstreamUtf8(tmpLogPath);
            // §56.53: 分段编译时**追加**而不是覆盖 —— 「一次 cl 编全部」那一轮里,
            // cl 会把**所有** TU 的诊断都报出来; 分段后若第一段失败就中断, 错误面
            // 会比老路**少** (后面的组根本没编) ⇒ 判红/判绿的错误面对不上。
            // 故各段的诊断都往同一份 c3-error.log 里续写 (第一段覆盖, 后面追加)。
            std::ofstream errLog = ofstreamUtf8(errorLogPath,
                std::ios::out | (appendErrorLog ? std::ios::app : std::ios::trunc));
            if (tmpLog && errLog) {
                errLog << "C3: Compilation failed (exit code " << ret << ")" << std::endl;
                // 失败留档完整命令与响应文件: 只看 tail-N 时, cl 静默退出 (D8050/C1083
                // 这类竞态) 会把"到底用什么参数编的"一起丢掉, 离线无法重放。rsp 是
                // UTF-16LE+BOM (Fix 196), 不能按 UTF-8 回读 —— 这里写的是刚交给 cl
                // 的同一份内容 (writeMsvcResponseFile 的入参), 与磁盘上的 rsp 等价。
                errLog << "=== Command ===" << std::endl;
                errLog << fullCmd << std::endl;
                errLog << "=== Response File (" << rspPath << ") ===" << std::endl;
                errLog << args << std::endl;
                errLog << "=== MSVC Output ===" << std::endl;
                std::string line;
                while (std::getline(tmpLog, line)) {
                    errLog << line << "\n";
                }
            }
            // 将 MSVC 输出打印到 stderr
            if (tmpLog) {
                tmpLog.clear();
                tmpLog.seekg(0);
                std::string line;
                while (std::getline(tmpLog, line)) {
                    std::cerr << line << std::endl;
                }
            }
            std::cerr << "C3: 编译失败 (exit code " << ret << ")" << std::endl;
            std::cerr << "C3: 错误日志已保存: " << errorLogPath << std::endl;
            std::filesystem::remove(tmpLogPath, std::error_code());
            std::filesystem::remove(rspPath, std::error_code());
            return false;
        }
        // 编译成功: 清理临时文件。keepLogOnSuccess ⇒ 日志**留着**给调用方 (§56.53:
        //   分段编译时成功那几段的输出也要进失败留档 —— 老路是一段 cl 编全部, 那一段
        //   失败 ⇒ 688 个 TU 的诊断全在 c3-error.log 里; 分段后只有失败的那几段落盘
        //   ⇒ 成功那几段 (RTL / PCH 建库) 的警告会从记录里消失)。
        if (!keepLogOnSuccess) {
            std::filesystem::remove(tmpLogPath, std::error_code());
        }
        std::filesystem::remove(rspPath, std::error_code());
        return true;
    }
}

MsvcDriver::MsvcDriver() {}
MsvcDriver::~MsvcDriver() = default;

// 以隐藏窗口执行命令行, 返回退出码。
// 必须隐藏: C3 常被从**无控制台**的宿主拉起 (Git Bash/mintty、GUI 启动器),
// 这种父进程下 cmd.exe 子进程会被分配一个新的可见控制台 -> 屏幕上连闪黑框。
int MsvcDriver::executeCommand(const std::string& cmd) {
#ifdef _WIN32
    // Fix 210b: 只在 TEMP 是 POSIX 风格时才覆写环境块, 否则 nullptr 让子进程继承
    std::wstring envBlock = maybeEnvBlockWithWindowsTemp();
    void* lpEnv = envBlock.empty() ? nullptr : (void*)(envBlock.c_str());
    DWORD creationFlags = CREATE_NO_WINDOW;
    if (lpEnv) creationFlags |= CREATE_UNICODE_ENVIRONMENT;
    // M22-IssueB: Use CreateProcessW to pass UTF-16 command line to cmd.exe
    // This preserves Chinese/Unicode characters in file paths (e.g. /Fe"工程1.exe")
    // std::system() converts char* via CRT codepage, which corrupts UTF-8 paths
    
    // Convert UTF-8 command to wide string
    int wlen = MultiByteToWideChar(CP_UTF8, 0, cmd.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return std::system(cmd.c_str());
    std::wstring wcmd(wlen, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, cmd.c_str(), -1, &wcmd[0], wlen);
    wcmd.pop_back(); // remove trailing null from MultiByteToWideChar
    
    // Build full command: cmd.exe /c <command>
    std::wstring fullCmd = L"cmd.exe /c " + wcmd;
    
    STARTUPINFOW si = { sizeof(si) };
    // 隐藏窗口: 不设这两项时, 无控制台宿主下的 cmd.exe 会新建可见控制台窗口
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    PROCESS_INFORMATION pi = {};
    
    // CreateProcessW requires mutable command line buffer
    std::wstring mutableCmd = fullCmd;
    
    BOOL ok = CreateProcessW(
        nullptr,                // application name (nullptr = use command line)
        &mutableCmd[0],         // command line (mutable)
        nullptr,                // process security
        nullptr,                // thread security
        FALSE,                  // inherit handles
        // Fix 210b: lpEnvironment 传的是 UTF-16 块, MSDN 明确要求
        // CREATE_UNICODE_ENVIRONMENT 一起给 —— 不给时 Windows 按 ANSI 逐字节切,
        // 子进程拿到的 env 变成一堆单字母条目 (T/E/M/P/=... 各一枚), cl→c1 落
        // _CL_*.tmp 时就是 Permission denied。
        creationFlags,
        lpEnv,                  // environment (Fix 210b: 只在 POSIX TEMP 下非空)
        nullptr,                // current directory
        &si,                    // startup info
        &pi                     // process info
    );
    
    if (!ok) {
        // Fallback to std::system if CreateProcessW fails
        return std::system(cmd.c_str());
    }
    
    WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD exitCode = 1;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    
    return static_cast<int>(exitCode);
#else
    return std::system(cmd.c_str());
#endif
}

// 同 executeCommand, 但把子进程 stdout 收进 out。
// vswhere 这类「跑一下读输出」的探测用它, 替换原来的 _popen (同样会弹 cmd 窗口)。
int MsvcDriver::executeCommandCapture(const std::string& cmd, std::string& out) {
    out.clear();
#ifdef _WIN32
    std::wstring envBlock = maybeEnvBlockWithWindowsTemp();
    void* lpEnv = envBlock.empty() ? nullptr : (void*)(envBlock.c_str());
    DWORD creationFlags = CREATE_NO_WINDOW;
    if (lpEnv) creationFlags |= CREATE_UNICODE_ENVIRONMENT;
    int wlen = MultiByteToWideChar(CP_UTF8, 0, cmd.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return -1;
    std::wstring wcmd(wlen, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, cmd.c_str(), -1, &wcmd[0], wlen);
    wcmd.pop_back();

    SECURITY_ATTRIBUTES sa = {};
    sa.nLength = sizeof(sa);
    sa.bInheritHandle = TRUE;
    HANDLE hRead = nullptr, hWrite = nullptr;
    if (!CreatePipe(&hRead, &hWrite, &sa, 0)) return -1;
    SetHandleInformation(hRead, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW si = { sizeof(si) };
    si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    si.wShowWindow = SW_HIDE;
    si.hStdOutput = hWrite;
    si.hStdError = hWrite;

    PROCESS_INFORMATION pi = {};
    std::wstring mutableCmd = L"cmd.exe /c " + wcmd;
    BOOL ok = CreateProcessW(nullptr, &mutableCmd[0], nullptr, nullptr, TRUE,
                             creationFlags, lpEnv,
                             nullptr, &si, &pi);
    // 父进程必须关掉写端, 否则子进程退出后 ReadFile 也等不到 EOF
    CloseHandle(hWrite);
    if (!ok) { CloseHandle(hRead); return -1; }

    char buf[512];
    DWORD n = 0;
    while (ReadFile(hRead, buf, sizeof(buf), &n, nullptr) && n > 0) {
        out.append(buf, n);
    }
    CloseHandle(hRead);
    WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD ec = 1;
    GetExitCodeProcess(pi.hProcess, &ec);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    return static_cast<int>(ec);
#else
    FILE* pipe = popen((cmd + " 2>&1").c_str(), "r");
    if (!pipe) return -1;
    char buf[512];
    while (fgets(buf, sizeof(buf), pipe)) out += buf;
    return pclose(pipe);
#endif
}


bool MsvcDriver::compileAndLink(const MsvcDriverOptions& options) {
    if (options.sourceFiles.empty()) {
        std::cerr << "C3: 没有源文件需要编译" << std::endl;
        return false;
    }

    // opt3: 增量编译路径 (基于内容哈希的obj级缓存)
    if (options.incremental) {
        return compileAndLinkIncremental(options);
    }

    // 账 §56.53: 有 PCH 分组 ⇒ 走「建 PCH → 分组编译 → 编其余 → 链接」四段;
    // 没有 ⇒ 原来的「一次 cl 编译+链接」原样保留 (命令行一字不改)。
    if (!options.pchGroups.empty()) {
        return compileAndLinkPch(options);
    }

    std::string cl = findClExe();

    // 构建cl.exe命令行
    std::ostringstream cmd;
    cmd << cl;
    // §56.53: 命令行拆成三段 (公共编译档 / 输出档 / 链接档)。拆是为了让 PCH 路径
    // 能只取其中一段 (只编不链); 老路仍把三段拼在一起 —— **命令行一字未改**。
    cmd << buildMsvcCompileFlags(options);
    cmd << buildMsvcOutputFlags(options);

    // 源文件列表
    for (const auto& src : options.sourceFiles) {
        cmd << " \"" << src << "\"";
    }

    // ai/vb-asm-extension-spec: ml64 汇编产物作为附加链接输入 (cl 会把 .obj 转交链接器)
    appendExtraObjects(cmd, options);

    // P11.3 (reverted): RTL 以 .c 源码加入 sourceFiles 编译, 无 .lib 链接

    cmd << buildMsvcLinkOptions(options);

    if (options.verbose) {
        std::cout << "C3: 执行: " << cmd.str() << std::endl;
    }

    // P11.2: MSVC output to temp file in objDir (intermediates dir)
    std::string tmpLogDir;
    if (!options.objDir.empty()) {
        tmpLogDir = options.objDir;
    } else if (!options.outputFile.empty()) {
        std::filesystem::path outP(utf8ToPath(options.outputFile));
        tmpLogDir = pathToUtf8(outP.parent_path());
    }
    if (tmpLogDir.empty()) tmpLogDir = ".";
    std::string tmpLogPath = tmpLogDir + "/_c3_msvc_out.txt";

    // Use response file to avoid cmd.exe command line length limit (8191 chars)
    // when compiling many source files (e.g. 125+ .c files in a large project)
    std::string rspPath = tmpLogDir + "/_c3_cl_args.rsp";
    // Fix 196: 响应文件写成 UTF-16LE+BOM (writeMsvcResponseFile 内) —— cl/link 按系统
    // ANSI 代码页读 @rsp, 原样落 UTF-8 会把中文路径解成乱码 (详见 msvc_driver.hpp
    // 的实测矩阵)。
    // §56.53: 写 rsp + 执行 + 失败留档 收进 runMsvcInvocation —— PCH 路径要跑好几段,
    // 每段一份 rsp/日志 (否则后一段会把前一段的现场覆盖掉, 离线定位不到是哪一段炸的)。
    // 老路仍用同一份 rsp 名, 现场可回放。
    // P11.4: Prepend vcvarsall.bat setup if cl.exe not in PATH
    std::string vcvarsPrefix = buildVcvarsPrefix(options.arch);
    // Fix <vbeclipse> D8050: 每次 cl 调用带独立 TMP/TEMP (= 本次编译的 objDir,
    // 由 session 目录保证唯一)。GA t2 的两个 GUI worker (Charts2020 x86 +
    // VBFlexGridDemo x64) 同 runner 上并行跑 /MP 时, 两批 c1.exe 会往同一个
    // 系统 %TEMP% 落 _CL_*.tmp 互相踩 (D8050 unable to write temporary file /
    // 亦表现为 C1083 Permission denied)。给每次编译一份隔离目录, 从根上断掉
    // 跨进程冲突; session 结束时目录连带被清理, 不留残留。
    std::string tmpIsolation = "set \"TMP=" + tmpLogDir + "\" && set \"TEMP=" + tmpLogDir + "\" && ";

    if (!runMsvcInvocation(cl, cmd.str().substr(cl.length()), rspPath, tmpLogPath,
                           options, vcvarsPrefix, tmpIsolation)) {
        return false;
    }

    if (options.verbose) {
        std::cout << "C3: 编译成功: " << options.outputFile << std::endl;
    }

    return true;
}


// §56.53: PCH 路径 —— 「建 PCH → 分组编译 → 编其余 TU → 链接」四段 cl。
//
// 为什么必须分段 (这是工具链的硬约束, 不是实现偏好):
//   * 一次 cl 调用**只能用一个 PCH** (/Yc /Yu /Fp 都是调用级选项);
//   * 前奏有两副 (普通档 / 窗体档) ⇒ 至少两份 PCH;
//   * RTL 那一族 (vb6rtl_*.c / vb6com*.c / vb6forms*.c) **不能用 PCH**: 各族前奏不一致
//     (vb6rtl_* 引 vb6rtl.h, vb6com.c 引 vb6com.h, vb6forms.c 自带 WIN32_LEAN_AND_MEAN
//     ⇒ 连 windows.h 的形态都不一样) ⇒ 套上去的状态不同构;
//   * driver 自产的 dll_entry.c / com_entry.c 不是发码器 TU ⇒ 同样不用 PCH。
//
// 语义安全性 (§56.53 第一原则): PCH 头的内容 = 该组 TU 自己的前奏 (逐字, 同一处排法
// backend/cgen_pch.hpp) ⇒ 套用 PCH 后的预处理状态与不开 PCH 时**同构**; TU 里那几行
// include 变成空操作 (include guard 已定义), 之后的 #undef 块 / IID 重映射 / 接口契约
// typedef 照旧在原位发生 ⇒ **零语义变化**。
bool MsvcDriver::compileAndLinkPch(const MsvcDriverOptions& options) {
    std::string cl = findClExe();

    // 中间件目录: .obj 与 PCH 产物都落这儿
    std::string objDir = options.objDir;
    if (objDir.empty()) {
        std::filesystem::path outPath(utf8ToPath(options.outputFile));
        objDir = pathToUtf8(outPath.parent_path());
    }
    if (objDir.empty()) objDir = ".";

    // PCH 产物目录 = objDir。PCH 头用**裸名** (c3_pch_std.h): /Yc /Yu /FI 三处都必须
    // 与源文件里 #include 的**拼写**一致 (实测 C2857: 用绝对路径 ⇒ cl 报「在源文件中
    // 没有找到用 /Yc... 指定的 #include 语句」), 裸名 + /I 命中才对得上。
    std::string pchDir = objDir;

    std::string tmpLogDir = objDir;
    std::string vcvarsPrefix = buildVcvarsPrefix(options.arch);
    std::string tmpIsolation = "set \"TMP=" + tmpLogDir + "\" && set \"TEMP=" + tmpLogDir + "\" && ";

    // 公共编译档 + /I<PCH 目录>。PCH 头在 objDir, 消费它的 TU 在 srcDir: 两者相等时
    // 这个 /I 是重复项 (无害), 不相等时它才是命中路径 —— 两种情形都必须对。
    // 建 PCH 的那次 cl 与消费它的那些 cl 用**同一份** flags (档不一样 ⇒ PCH 状态不同构)。
    std::string compileFlags = buildMsvcCompileFlags(options) + " /I\"" + pchDir + "\"";

    // ---------- 1) 写 PCH 头 + PCH 源 ----------
    struct PchBuild {
        std::string headerName;   // 裸名
        std::string stem;         // 去掉 .h 的名字 (作 PCH 源/.pch/.obj 的基名)
        std::string headerPath;
        std::string sourcePath;
        std::string pchPath;
        std::string objPath;
        const std::vector<std::string>* sources = nullptr;
    };
    std::vector<PchBuild> builds;
    for (const auto& g : options.pchGroups) {
        if (g.sources.empty()) continue;
        PchBuild b;
        b.headerName = g.headerName;
        b.stem = g.headerName;
        if (b.stem.size() > 2 && b.stem.compare(b.stem.size() - 2, 2, ".h") == 0)
            b.stem = b.stem.substr(0, b.stem.size() - 2);
        b.headerPath = pchDir + "/" + g.headerName;
        b.sourcePath = pchDir + "/" + b.stem + ".c";
        b.pchPath    = pchDir + "/" + b.stem + ".pch";
        b.objPath    = pchDir + "/" + b.stem + ".obj";
        b.sources    = &g.sources;
        {
            std::ofstream ofs = ofstreamUtf8(b.headerPath, std::ios::out | std::ios::trunc);
            if (!ofs) {
                std::cerr << "C3: 无法写入 PCH 头: " << b.headerPath << std::endl;
                return false;
            }
            ofs << "// Generated by vb6c3 (C3.exe) - PCH header (precompiled prelude)\n";
            ofs << "// The include block below is byte-identical to the module .h prelude\n";
            ofs << "// (single source of truth: src/backend/cgen_pch.hpp).\n";
            for (const auto& line : g.preludeLines) ofs << line << "\n";
        }
        {
            std::ofstream ofs = ofstreamUtf8(b.sourcePath, std::ios::out | std::ios::trunc);
            if (!ofs) {
                std::cerr << "C3: 无法写入 PCH 源: " << b.sourcePath << std::endl;
                return false;
            }
            ofs << "// Generated by vb6c3 (C3.exe) - PCH source\n";
            ofs << "#include \"" << g.headerName + "\"\n";
        }
        builds.push_back(b);
    }
    if (builds.empty()) {
        std::cerr << "C3: PCH 分组全空 (无可用分组)" << std::endl;
        return false;
    }

    // 编译阶段是否有失败 (见下面第 3/4 段的注记: 分段后**不**因首个失败而中断)
    bool failed = false;
    // 成功那几段的输出 (留档用): 失败时并进 c3-error.log, 否则它们的警告会从记录里消失
    struct OkPhaseLog {
        std::string label;
        std::string logPath;
        std::string args;
    };
    std::vector<OkPhaseLog> okPhaseLogs;

    // ---------- 2) 建 PCH (每组一次 cl: /Yc) ----------
    // 建库失败 = 编译器自己的问题 (不是用户代码) ⇒ 立刻停, 别把派生噪声堆上去。
    for (const auto& b : builds) {
        std::ostringstream c;
        c << compileFlags;
        c << " /c /Yc\"" << b.headerName << "\" /Fp\"" << b.pchPath << "\"";
        c << " /Fo\"" << pchDir + "/\"";
        c << " \"" << b.sourcePath << "\"";
        if (options.verbose) {
            std::cout << "C3: [PCH] 建库 " << b.headerName << ": " << c.str() << std::endl;
        }
        std::string rsp = pchDir + "/_c3_cl_pch_" + b.stem + ".rsp";
        std::string log = pchDir + "/_c3_msvc_out_pch_" + b.stem + ".txt";
        // keepLogOnSuccess: 成功那段的输出留着 (失败时要并进 c3-error.log)
        if (!runMsvcInvocation(cl, c.str(), rsp, log, options, vcvarsPrefix, tmpIsolation,
                               false, true))
            return false;
        okPhaseLogs.push_back({"[PCH 建库 " + b.stem + "]", log, c.str()});
    }

    // ---------- 3) 按组编译 (每组一次 cl: /FI + /Yu 复用本组 PCH) ----------
    // ⚠ 某组编不过**不中断**: 老路是一次 cl 编全部 TU, cl 会把**所有** TU 的诊断都报
    //   出来; 分段后若第一组失败就停, 错误面会比老路**少** (后面的组根本没编) ⇒
    //   判红/判绿的错误面对不上。PD 整编正是失败构建, 这条直接决定 GA 的「逐码相等」
    //   能不能比。故各组都编完, 诊断续写进同一份 c3-error.log, 最后统一判失败。
    for (const auto& b : builds) {
        std::ostringstream c;
        c << compileFlags;
        c << " /c /FI\"" << b.headerName << "\" /Yu\"" << b.headerName << "\"";
        c << " /Fp\"" << b.pchPath << "\"";
        c << " /Fo\"" << objDir + "/\"";
        for (const auto& src : *b.sources) c << " \"" << src << "\"";
        if (options.verbose) {
            std::cout << "C3: [PCH] 编译 " << b.headerName << " 组 (" << b.sources->size()
                      << " 个 TU)" << std::endl;
        }
        std::string rsp = pchDir + "/_c3_cl_grp_" + b.stem + ".rsp";
        std::string log = pchDir + "/_c3_msvc_out_grp_" + b.stem + ".txt";
        bool okGrp = runMsvcInvocation(cl, c.str(), rsp, log, options, vcvarsPrefix, tmpIsolation,
                                       failed, true);
        if (!okGrp) failed = true;
        else okPhaseLogs.push_back({"[PCH 组 " + b.stem + "]", log, c.str()});
    }

    // ---------- 4) 编其余 TU (RTL + driver 自产的 .c): 不用 PCH ----------
    std::unordered_set<std::string> inGroup;
    for (const auto& b : builds)
        for (const auto& s : *b.sources) inGroup.insert(s);
    std::vector<std::string> rest;
    for (const auto& s : options.sourceFiles)
        if (!inGroup.count(s)) rest.push_back(s);
    if (!rest.empty()) {
        std::ostringstream c;
        c << compileFlags;
        c << " /c /Fo\"" << objDir + "/\"";
        for (const auto& s : rest) c << " \"" << s << "\"";
        if (options.verbose) {
            std::cout << "C3: [PCH] 编译其余 TU (" << rest.size() << " 个, 无 PCH)" << std::endl;
        }
        std::string rsp = pchDir + "/_c3_cl_rest.rsp";
        std::string log = pchDir + "/_c3_msvc_out_rest.txt";
        // 同上: 前面几组失败也要把这一段编完 (老路是一次 cl 编全部 ⇒ RTL 的诊断也在同一份日志里)
        bool okRest = runMsvcInvocation(cl, c.str(), rsp, log, options, vcvarsPrefix, tmpIsolation,
                                         failed, true);
        if (!okRest) failed = true;
        else okPhaseLogs.push_back({"[无 PCH 其余 TU]", log, c.str()});
    }

    // 失败留档补全: 把**成功**那几段的输出并进 c3-error.log。
    // 为什么: 老路是一段 cl 编全部 TU, 那一段失败 ⇒ 688 个 TU 的诊断全在日志里;
    // 分段后只有失败的那几段落盘 ⇒ 成功那几段 (PCH 建库 / RTL) 的警告会从记录里消失
    // (PD 实测: 少了 126 条警告, 全是头级与 RTL 的)。取证要全 ⇒ 补齐。
    if (failed && !okPhaseLogs.empty()) {
        std::string outputDirForLog;
        if (!options.outputFile.empty()) {
            std::filesystem::path outP(utf8ToPath(options.outputFile));
            outputDirForLog = pathToUtf8(outP.parent_path());
        }
        if (outputDirForLog.empty()) outputDirForLog = ".";
        std::string errorLogPath = outputDirForLog + "/c3-error.log";
        std::ofstream errLog = ofstreamUtf8(errorLogPath, std::ios::out | std::ios::app);
        if (errLog) {
            for (const auto& ph : okPhaseLogs) {
                errLog << "=== Phase " << ph.label << " (succeeded) ===" << std::endl;
                errLog << "=== Response File ===" << std::endl;
                errLog << ph.args << std::endl;
                errLog << "=== MSVC Output ===" << std::endl;
                std::ifstream in = ifstreamUtf8(ph.logPath);
                std::string line;
                while (std::getline(in, line)) errLog << line << "\n";
                errLog << std::endl;
            }
        }
    }

    // 编译阶段有失败 ⇒ 不进链接 (老路里 cl 有编译错误时也不会去链接)。
    if (failed) {
        std::cerr << "C3: 编译失败 (PCH 分段编译; 各段诊断已并入 c3-error.log)" << std::endl;
        return false;
    }

    // ---------- 5) 链接: 全部 .obj (含各组 PCH 自己的 .obj) ----------
    // ⚠ 必须带 PCH 的 .obj: 用 /Yu 编出来的 .obj 引用了 PCH, 链接时不带上它 ⇒
    //   LNK2011「未链接预编译对象」。实测矩阵见本文件 §56.53 注记。
    std::ostringstream c;
    c << compileFlags;
    c << buildMsvcOutputFlags(options);
    for (const auto& s : options.sourceFiles) c << " \"" << msvcObjPathFor(s, objDir) << "\"";
    for (const auto& b : builds) c << " \"" << b.objPath << "\"";
    appendExtraObjects(c, options);
    c << buildMsvcLinkOptions(options);
    if (options.verbose) {
        std::cout << "C3: [PCH] 链接: " << c.str() << std::endl;
    }
    std::string rsp = pchDir + "/_c3_cl_link.rsp";
    std::string log = pchDir + "/_c3_msvc_out_link.txt";
    if (!runMsvcInvocation(cl, c.str(), rsp, log, options, vcvarsPrefix, tmpIsolation))
        return false;

    if (options.verbose) {
        std::cout << "C3: 编译成功: " << options.outputFile << std::endl;
    }
    return true;
}
} // namespace vb6c3
