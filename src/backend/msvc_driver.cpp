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

    std::string cl = findClExe();

    // 构建cl.exe命令行
    std::ostringstream cmd;
    cmd << cl;

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
        cmd << " /Zi /DEBUG";
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
    // 由独立 cl 进程并行编译。注意 /MP 与 /GL、/Yc、/Yu 不兼容(本项目未使用)。
    cmd << " /MP";

    // P24-09: x86 RTL libraries are built with /MT (static CRT); match it to avoid LNK2038
    if (options.arch == "x86") {
        cmd << " /MT";
    }

    // Output file and .obj directory (P11.2: intermediates go to objDir)
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

    // 源文件列表
    for (const auto& src : options.sourceFiles) {
        cmd << " \"" << src << "\"";
    }

    // ai/vb-asm-extension-spec: ml64 汇编产物作为附加链接输入 (cl 会把 .obj 转交链接器)
    appendExtraObjects(cmd, options);

    // P11.3 (reverted): RTL 以 .c 源码加入 sourceFiles 编译, 无 .lib 链接

    // 链接选项
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
    // Fix 196: 必须写成 UTF-16LE+BOM —— cl/link 按系统 ANSI 代码页读 @rsp,
    // 原样落 UTF-8 会把中文路径解成乱码 (详见 msvc_driver.hpp 的实测矩阵)。
    writeMsvcResponseFile(rspPath, cmd.str().substr(cl.length()));

    // P11.4: Prepend vcvarsall.bat setup if cl.exe not in PATH
    std::string vcvarsPrefix = buildVcvarsPrefix(options.arch);
    // Fix <vbeclipse> D8050: 每次 cl 调用带独立 TMP/TEMP (= 本次编译的 objDir,
    // 由 session 目录保证唯一)。GA t2 的两个 GUI worker (Charts2020 x86 +
    // VBFlexGridDemo x64) 同 runner 上并行跑 /MP 时, 两批 c1.exe 会往同一个
    // 系统 %TEMP% 落 _CL_*.tmp 互相踩 (D8050 unable to write temporary file /
    // 亦表现为 C1083 Permission denied)。给每次编译一份隔离目录, 从根上断掉
    // 跨进程冲突; session 结束时目录连带被清理, 不留残留。
    std::string tmpIsolation = "set \"TMP=" + tmpLogDir + "\" && set \"TEMP=" + tmpLogDir + "\" && ";
    std::string fullCmd = vcvarsPrefix + tmpIsolation + cl + " @\"" + rspPath + "\" > \"" + tmpLogPath + "\" 2>&1";

    int ret = executeCommand(fullCmd);
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
        std::ofstream errLog = ofstreamUtf8(errorLogPath, std::ios::out | std::ios::trunc);
        if (tmpLog && errLog) {
            errLog << "C3: Compilation failed (exit code " << ret << ")" << std::endl;
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
    // 编译成功: 清理临时文件
    std::filesystem::remove(tmpLogPath, std::error_code());
    std::filesystem::remove(rspPath, std::error_code());

    if (options.verbose) {
        std::cout << "C3: 编译成功: " << options.outputFile << std::endl;
    }

    return true;
}

} // namespace vb6c3
