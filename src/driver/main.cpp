#include "driver/driver.hpp"
#include "driver/driver_pack.hpp"
#include "common/encoding.hpp"
#include <string>
#include <vector>
#include <iostream>
#include <algorithm>
#include <filesystem>

#ifdef _WIN32
#include <windows.h>
#include <shellapi.h>  // CommandLineToArgvW
#include <dbghelp.h>
#include <crtdbg.h>
#include <cstdio>
#include <cstdlib>

// ---------------------------------------------------------------------------
// 崩溃栈追踪 (诊断用)
//
// C3 在特定输入下会自身崩溃 (0xC0000005), 而 ASAN 构建不崩, 无法用 ASAN report
// 定位. 这里提供一个基于 dbghelp 的进程内崩溃栈打印:
//
//   set C3_CRASH_TRACE=1
//   C3.exe <vbp> ...        → 崩溃时 stderr 输出异常码/地址 + 符号化调用栈
//
// 未设置该环境变量时不会安装任何处理器, 行为与原先完全一致.
// 注意: 需要带 PDB 的构建 (RelWithDebInfo / Debug) 才能解析出 函数+文件:行.
// ---------------------------------------------------------------------------
namespace {

LONG WINAPI c3CrashHandler(EXCEPTION_POINTERS* ep) {
    const DWORD code = ep->ExceptionRecord ? ep->ExceptionRecord->ExceptionCode : 0;
    std::fprintf(stderr, "\n[C3-CRASH] exception=0x%08lX address=%p\n",
                 (unsigned long)code,
                 ep->ExceptionRecord ? ep->ExceptionRecord->ExceptionAddress : nullptr);

    HANDLE proc = GetCurrentProcess();
    SymSetOptions(SYMOPT_LOAD_LINES | SYMOPT_UNDNAME | SYMOPT_DEFERRED_LOADS);
    if (SymInitialize(proc, nullptr, TRUE)) {
        void* frames[64];
        USHORT n = CaptureStackBackTrace(0, 64, frames, nullptr);
        char symBuf[sizeof(SYMBOL_INFO) + 512];
        for (USHORT i = 0; i < n; ++i) {
            SYMBOL_INFO* si = reinterpret_cast<SYMBOL_INFO*>(symBuf);
            si->SizeOfStruct = sizeof(SYMBOL_INFO);
            si->MaxNameLen = 512;
            DWORD64 symDisp = 0;
            IMAGEHLP_LINE64 line{};
            line.SizeOfStruct = sizeof(IMAGEHLP_LINE64);
            DWORD lineDisp = 0;
            const BOOL hasSym = SymFromAddr(proc, reinterpret_cast<DWORD64>(frames[i]), &symDisp, si);
            const BOOL hasLine = SymGetLineFromAddr64(proc, reinterpret_cast<DWORD64>(frames[i]), &lineDisp, &line);
            if (hasSym && hasLine) {
                std::fprintf(stderr, "  #%02u %s+0x%llX  (%s:%lu)\n", (unsigned)i, si->Name,
                             (unsigned long long)symDisp, line.FileName, (unsigned long)line.LineNumber);
            } else if (hasSym) {
                std::fprintf(stderr, "  #%02u %s+0x%llX\n", (unsigned)i, si->Name,
                             (unsigned long long)symDisp);
            } else {
                std::fprintf(stderr, "  #%02u %p\n", (unsigned)i, frames[i]);
            }
        }
        SymCleanup(proc);
    }
    std::fflush(stderr);
    TerminateProcess(GetCurrentProcess(), static_cast<UINT>(code));
    return EXCEPTION_EXECUTE_HANDLER;
}

void installCrashTraceIfRequested() {
    char buf[8] = {0};
    if (GetEnvironmentVariableA("C3_CRASH_TRACE", buf, sizeof(buf)) > 0 && buf[0] != '0') {
        SetUnhandledExceptionFilter(c3CrashHandler);
    }
}

// C3 是**控制台**编译器: 它出错的方式只能是一条诊断 + 退出码, 不能是一个等着人点的前台模态框。
// 调试版 CRT 的默认行为恰好相反 —— assert/invalid-parameter/abort 先弹
// 「Microsoft Visual C++ Runtime Library / Debug Error!」再退出, 会把调用方(测试夹具、IDE、
// 自动化代理)的整个会话卡住, 而诊断信息一条也拿不到 (账 #116 的现场就是这样卡住的)。
// 这里把三类报告全部改道 stderr, 并去掉 abort() 的那条消息框。Release 构建里没有这些符号,
// 故整块包在 _DEBUG 下。
void silenceDebugCrtDialogs() {
    const int types[] = {_CRT_ASSERT, _CRT_ERROR, _CRT_WARN};
    for (int t : types) {
        _CrtSetReportMode(t, _CRTDBG_MODE_FILE);
        _CrtSetReportFile(t, _CRTDBG_FILE_STDERR);
    }
}

}  // namespace

// ---------------------------------------------------------------------------
// 控制台输出: 按句柄类型选路
//
//   * 句柄是**控制台**(含 ConPTY/Windows Terminal) -> UTF-8 转 UTF-16 后 WriteConsoleW:
//     控制台渲染宽字符与 chcp 无关, 也不需要动用户的代码页。
//   * 句柄是**管道/文件**(被 PowerShell 接走、或用户 `> file`) -> 写**控制台代码页**的字节:
//     cmd 重定向与 PowerShell 都按这个代码页解码 (实测 PS 5.1 与 pwsh 7 的
//     `[Console]::OutputEncoding` 默认都跟控制台代码页走, 本机 = gb2312), 写 UTF-8 给它们
//     才是乱码。装不下时 (例如 437 代码页要写中文) 退回 UTF-8, 不丢字。
//
// 老办法是"进 C3 时把控制台代码页切成 65001、退出恢复", 有两个坑: ① 判据是
// GetConsoleWindow(), 在 Windows Terminal/ConPTY 下返回 NULL ⇒ 守卫不触发, UTF-8 字节
// 落到 936 控制台上 = 乱码; ② 切换是**整个控制台**的 —— C3 随后拉起的 cl.exe/link.exe
// 吐 GBK 中文, 在 65001 下反而成了乱码。现在这个实现既不看窗口、也不动代码页。
// ---------------------------------------------------------------------------
class ConsoleUtf8Buf : public std::streambuf {
public:
    ConsoleUtf8Buf(std::streambuf* fallback, HANDLE h, bool isConsole, UINT targetCp)
        : fallback_(fallback), h_(h), isConsole_(isConsole), cp_(targetCp) {}

protected:
    int overflow(int ch) override {
        if (ch == EOF) return 0;
        buf_.push_back(static_cast<char>(ch));
        if (ch == '\n') flushBuf();
        return ch;
    }

    std::streamsize xsputn(const char* s, std::streamsize n) override {
        buf_.append(s, static_cast<size_t>(n));
        if (buf_.find('\n') != std::string::npos) flushBuf();
        return n;
    }

    int sync() override {
        flushBuf();
        if (fallback_) fallback_->pubsync();
        return 0;
    }

private:
    // UTF-8 整行 -> UTF-16 -> 目标编码; 尾部若是被截断的多字节序列就留到下次, 不丢半个汉字
    void flushBuf() {
        if (buf_.empty()) return;
        int wlen = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, buf_.data(), (int)buf_.size(),
                                       nullptr, 0);
        if (wlen <= 0) {          // 结尾是半个字符: 退到最后一个完整序列再转
            size_t keep = 0;
            for (size_t cut = 1; cut <= 3 && cut < buf_.size(); ++cut) {
                size_t n = buf_.size() - cut;
                if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, buf_.data(), (int)n, nullptr, 0) > 0) {
                    keep = n;
                    break;
                }
            }
            if (keep == 0) {      // 真转不动 (不该发生): 原样写出, 别卡着
                writeBytes(buf_.data(), buf_.size());
                buf_.clear();
                return;
            }
            std::string tail = buf_.substr(keep);
            buf_.resize(keep);
            flushBuf();
            buf_ = tail;
            return;
        }
        {
            std::wstring w(static_cast<size_t>(wlen), L'\0');
            MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, buf_.data(), (int)buf_.size(), &w[0], wlen);
            if (isConsole_) {
                size_t off = 0;
                while (off < w.size()) {
                    DWORD chunk = static_cast<DWORD>(std::min<size_t>(w.size() - off, 4096));
                    DWORD done = 0;
                    if (!WriteConsoleW(h_, w.data() + off, chunk, &done, nullptr) || done == 0) break;
                    off += done;
                }
            } else {
                writeWide(w.data(), (int)w.size());
            }
        }
        buf_.clear();
    }

    void writeWide(const wchar_t* w, int len) {
        BOOL usedDefault = FALSE;
        int need = WideCharToMultiByte(cp_, 0, w, len, nullptr, 0, nullptr, &usedDefault);
        if (need <= 0 || usedDefault) {   // 目标代码页装不下 -> UTF-8 兜底
            int n8 = WideCharToMultiByte(CP_UTF8, 0, w, len, nullptr, 0, nullptr, nullptr);
            if (n8 <= 0) return;
            std::string s8(static_cast<size_t>(n8), '\0');
            WideCharToMultiByte(CP_UTF8, 0, w, len, &s8[0], n8, nullptr, nullptr);
            writeBytes(s8.data(), s8.size());
            return;
        }
        std::string sb(static_cast<size_t>(need), '\0');
        WideCharToMultiByte(cp_, 0, w, len, &sb[0], need, nullptr, nullptr);
        writeBytes(sb.data(), sb.size());
    }

    void writeBytes(const char* s, size_t n) {
        if (fallback_) fallback_->sputn(s, (std::streamsize)n);
    }

    std::streambuf* fallback_;
    HANDLE h_;
    bool isConsole_;
    UINT cp_;
    std::string buf_;
};

// stdout/stderr 都装上: 是控制台就走宽字符, 是管道/文件就按控制台代码页转字节
static void installConsoleUtf8(std::ostream& os, DWORD stdHandleId) {
    HANDLE h = GetStdHandle(stdHandleId);
    DWORD mode = 0;
    bool isConsole = false;
    if (!h || h == INVALID_HANDLE_VALUE) return;
    isConsole = GetConsoleMode(h, &mode) != 0;
    UINT cp = GetConsoleOutputCP();
    if (!cp) cp = GetACP();
    static ConsoleUtf8Buf* coutBuf = nullptr;
    static ConsoleUtf8Buf* cerrBuf = nullptr;
    ConsoleUtf8Buf*& slot = (stdHandleId == STD_ERROR_HANDLE) ? cerrBuf : coutBuf;
    if (slot) return;
    slot = new ConsoleUtf8Buf(os.rdbuf(), h, isConsole, cp);   // 常驻: 与进程同寿
    os.rdbuf(slot);
}

// 退出前把两条流刷干净 (std::cout/cerr 的静态析构顺序不可靠)
struct ConsoleFlushAtExit {
    ~ConsoleFlushAtExit() {
        std::cout.flush();
        std::cerr.flush();
    }
};
#endif

namespace {

int runCompile(vb6c3::Driver& driver, int argc, char* argv[]) {
    // Fix 196: 路径转换异常兜底。
    //   std::filesystem 的**窄串**重载 (path(const std::string&), exists(str), absolute(str) ...)
    //   在 Windows 上按系统 ACP (中文机为 936/GBK) 解释 char*, 而 C3 内部路径一律 UTF-8。
    //   一旦那些字节不是合法 ACP 序列 —— 路径里含奇数个非 ASCII 字符时极常见, 例如
    //   "..\中文叉\Form1.frm" 的 UTF-8 字节解 GBK 会剩一个孤立前导字节 ——
    //   MultiByteToWideChar 返回 0, _Convert_narrow_to_wide 抛 filesystem_error。
    //   此前无人接住 -> std::terminate -> abort(): 用户看到的是模态「Debug Error」对话框
    //   (退出码 3), 而不是一条能看懂的错误信息。
    //   这里兜住整条管线: 至少保证编译器永远"死得有话说"(退出码 1)。
    //   根治仍是让每个入口走 utf8ToPath (见 common/encoding.hpp 的 *Utf8 包装族)。
    try {
        // ai/023 S06: --pack/--unpack 分发外壳。在进入编译管线前截获 ——
        // 容器是纯分发形态, 编译器本体永不读它 (D5)。
        for (int i = 1; i < argc; ++i) {
            std::string a = argv[i] ? argv[i] : "";
            if (a == "--pack" && i + 1 < argc) {
                return vb6c3::runPackMode(argv[i + 1]);
            }
            if (a == "--unpack" && i + 1 < argc) {
                std::string outDir;
                for (int j = i + 2; j + 1 < argc; ++j) {
                    if (argv[j] && std::string(argv[j]) == "--output-dir") outDir = argv[j + 1];
                }
                return vb6c3::runUnpackMode(argv[i + 1], outDir);
            }
        }
        auto result = driver.compile(argc, argv);

        // success=false 且 errorCount=0 = 编译未成功却无错误计数 (如 GUI 工程链接未产出
        // exe) — 必须判为失败. -h/--help/--version 走的是 compile() 里 success=true 的
        // "正常退出" 分支, 不会落到这里, 因此脚本里 c3 --version 退出码为 0.
        if (!result.success && result.errorCount == 0) {
            return 1;
        }

        return result.errorCount > 0 ? 1 : 0;
    } catch (const std::filesystem::filesystem_error& e) {
        std::cerr << "C3: 路径处理失败: " << e.what() << std::endl;
        std::cerr << "C3: 这通常意味着某个路径没有经 utf8ToPath() 转换就交给了 std::filesystem。"
                  << std::endl;
        return 1;
    } catch (const std::exception& e) {
        // Fix 196 的同一族, 只是抛点是别的 std::exception (账 #116 现场量到: 模块级
        // `Dim gT As String * 8` 触发 std::terminate -> abort() -> 模态「Debug Error」框)。
        // 编译器**任何**未接住的异常都必须落成一条能看懂的诊断 + 退出码 1: 静默崩或弹窗
        // 都会把调用方(测试夹具 / IDE / 自动化代理)的会话卡住。
        std::cerr << "C3: 编译器内部异常: " << e.what() << std::endl;
        std::cerr << "C3: 这是编译器缺陷 (ICE), 请把触发它的源码与这一行一起报上来。" << std::endl;
        return 1;
    } catch (...) {
        std::cerr << "C3: 编译器内部异常: (非 std::exception)" << std::endl;
        return 1;
    }
}

}  // namespace

int main(int argc, char* argv[]) {
#ifdef _WIN32
    installCrashTraceIfRequested();
#ifdef _DEBUG
    silenceDebugCrtDialogs();
#endif
    installConsoleUtf8(std::cout, STD_OUTPUT_HANDLE);
    installConsoleUtf8(std::cerr, STD_ERROR_HANDLE);
    static ConsoleFlushAtExit consoleFlushAtExit;   // 与进程同寿
#endif

    vb6c3::Driver driver;

#ifdef _WIN32
    // Windows 上 main() 的 argv 是 ACP(中文系统为GBK) 编码, 而 C3 内部统一按 UTF-8
    // 处理字符串。直接把 argv 交给 utf8ToPath() 会把 GBK 字节当作 UTF-8 解析,
    // 中文路径因此损坏 -> 文件打不开 -> 报 ".vbp文件中没有源文件"。
    // 这里改用 Unicode 命令行重新取参, 统一转成 UTF-8。
    int wargc = 0;
    LPWSTR* wargv = CommandLineToArgvW(GetCommandLineW(), &wargc);
    if (wargv && wargc > 0) {
        std::vector<std::string> utf8Args;
        std::vector<char*> argvUtf8;
        utf8Args.reserve(wargc);
        argvUtf8.reserve(wargc);
        for (int i = 0; i < wargc; ++i) {
            utf8Args.push_back(vb6c3::wideToUtf8(wargv[i]));
        }
        for (auto& s : utf8Args) {
            argvUtf8.push_back(&s[0]);
        }
        LocalFree(wargv);
        return runCompile(driver, wargc, argvUtf8.data());
    }
    if (wargv) LocalFree(wargv);
#endif

    return runCompile(driver, argc, argv);
}
