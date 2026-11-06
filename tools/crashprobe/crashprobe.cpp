// crashprobe - 启动型最小调试器: 复现崩溃时抓异常码 / faulting module+offset /
// 首次故障地址 / 调用栈, 并写 minidump. 用于本仓 "先复现再修" 的取证环节.
//
// 同一份源码编两个位宽 (tools/crashprobe/build.bat):
//   crashprobe.exe    x64   —— 调试 64 位目标
//   crashprobe32.exe  x86   —— 调试 32 位目标 (C3 默认产物的位宽)
// 调试器与目标**必须同位宽**: 不要拿 x64 版去调 WOW64 目标 —— 那要处理
// WOW64_CONTEXT / Wow64StackWalk / 混合 dump, 本文具刻意不做那一档,
// 位宽不符时 stack/dump 两边都可能半对半错, 比抓不到更误导.
// 用户只调包装 scripts/crashprobe.ps1, 由它读目标 PE 头自动选二进制.
//
// 用法:
//   crashprobe.exe <target.exe> <dump-path>
//   crashprobe32.exe <target.exe> <dump-path>
//   例: crashprobe32.exe .\B5b.exe .\B5b.dmp
//   (推荐) powershell -File scripts\crashprobe.ps1 -Target .\B5b.exe -DumpPath .\B5b.dmp
//
// 输出 (stdout):
//   ARCH=x64|x86
//   PID=...
//   EXCEPTION code=0xC0000005 firstChance=1 addr=...
//     accessType=0 faultAddr=...        (0=读, 1=写, 8=执行)
//   FAULT_MODULE=<dll path> base=... offset=...
//   STACK: #00 ...
//   DUMP=<dump-path>
//   EXIT_CODE=0xC0000005 (-1073741819)
//
// 退出码 = EXIT_CODE = 目标进程的退出码 (= 异常码, 若目标以异常终止).
// 便于脚本判定; 别把 0xC0000005 当普通"失败"处理 —— 那正是取证目标本身.
//
// 限制 (重要):
//   1. 仅在本机 WER 已被禁用或不可用时可靠 (HKLM\...\Windows Error Reporting\
//      Disabled=1, WerSvc=Disabled/Stopped). WER 可用时系统会接管第二机会异常,
//      拿到的可能不是目标那一发. 在 CI 上不要指望 WER 取证.
//   2. 目标进程以 DEBUG_ONLY_THIS_PROCESS 启动, 因此崩溃对话框不会弹出, 适合
//      无人值守. 抓到首个 AV 后, 在 second-chance 阶段主动 TerminateProcess,
//      退出码 = 异常码 (避免无头环境下重复 first-chance 导致挂死).
//   3. 栈只给模块内地址; 本仓 EXE 无 PDB, 需要符号请对照同名 .map.
//   4. 位宽: 本文具与目标必须同位宽 (见文件头).
//
// 构建: build.bat (会自动找 vcvarsall; 也可设 VCVARSALL 环境变量再调).
#include <windows.h>
#include <dbghelp.h>
#include <psapi.h>
#include <stdio.h>

#if defined(_M_X64)
#define CRASHPROBE_ARCH      "x64"
#define CRASHPROBE_MACHINE   IMAGE_FILE_MACHINE_AMD64
#elif defined(_M_IX86)
#define CRASHPROBE_ARCH      "x86"
#define CRASHPROBE_MACHINE   IMAGE_FILE_MACHINE_I386
#else
#error "crashprobe: only x64 and x86 targets are supported"
#endif

static void printStack(HANDLE hProcess, HANDLE hThread, PEXCEPTION_POINTERS ep) {
    STACKFRAME64 sf;
    ZeroMemory(&sf, sizeof(sf));
    CONTEXT ctxCopy = ep->ContextRecord[0];
#if defined(_M_X64)
    sf.AddrPC.Offset = ctxCopy.Rip;
    sf.AddrFrame.Offset = ctxCopy.Rbp;
    sf.AddrStack.Offset = ctxCopy.Rsp;
#else
    sf.AddrPC.Offset = ctxCopy.Eip;
    sf.AddrFrame.Offset = ctxCopy.Ebp;
    sf.AddrStack.Offset = ctxCopy.Esp;
#endif
    sf.AddrPC.Mode = AddrModeFlat;
    sf.AddrFrame.Mode = AddrModeFlat;
    sf.AddrStack.Mode = AddrModeFlat;
    SymSetOptions(SYMOPT_UNDNAME | SYMOPT_DEFERRED_LOADS | SYMOPT_LOAD_LINES);
    SymInitialize(hProcess, NULL, FALSE);
    printf("STACK:\n");
    for (int i = 0; i < 16; i++) {
        if (!StackWalk64(CRASHPROBE_MACHINE, hProcess, hThread, &sf, &ctxCopy,
                         NULL, SymFunctionTableAccess64, SymGetModuleBase64, NULL))
            break;
        if (sf.AddrPC.Offset == 0) break;
        printf("  #%02d %p\n", i, (void*)(ULONG_PTR)sf.AddrPC.Offset);
    }
}

static void printFaultModule(HANDLE hProcess, PVOID addr) {
    HMODULE mods[1024];
    DWORD needed = 0;
    if (!EnumProcessModulesEx(hProcess, mods, sizeof(mods), &needed, LIST_MODULES_ALL))
        return;
    for (DWORD i = 0; i < needed / sizeof(HMODULE); i++) {
        MODULEINFO mi;
        if (!GetModuleInformation(hProcess, mods[i], &mi, sizeof(mi))) continue;
        ULONG_PTR a = (ULONG_PTR)addr;
        ULONG_PTR base = (ULONG_PTR)mi.lpBaseOfDll;
        if (a >= base && a < base + mi.SizeOfImage) {
            wchar_t path[MAX_PATH];
            if (GetModuleFileNameExW(hProcess, mods[i], path, MAX_PATH))
                printf("FAULT_MODULE=%ls base=%p offset=0x%llX\n",
                       path, (void*)base, (unsigned long long)(a - base));
            return;
        }
    }
}

int wmain(int argc, wchar_t** argv) {
    if (argc < 3) {
        printf("usage: crashprobe[32] <target.exe> <dump-path>\n");
        return 2;
    }
    printf("ARCH=%s\n", CRASHPROBE_ARCH);
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si)); si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));
    if (!CreateProcessW(argv[1], NULL, NULL, NULL, TRUE,
                        DEBUG_ONLY_THIS_PROCESS, NULL, NULL, &si, &pi)) {
        printf("CreateProcess failed %lu\n", GetLastError());
        return 3;
    }
    printf("PID=%lu\n", pi.dwProcessId);
    fflush(stdout);

    bool captured = false;
    bool secondChanceSeen = false;
    DWORD exitCode = 0;
    for (;;) {
        DEBUG_EVENT de;
        if (!WaitForDebugEvent(&de, INFINITE)) break;
        DWORD cont = DBG_CONTINUE;

        if (de.dwDebugEventCode == EXCEPTION_DEBUG_EVENT) {
            EXCEPTION_RECORD* er = &de.u.Exception.ExceptionRecord;
            printf("EXCEPTION code=0x%08lX firstChance=%ld addr=%p\n",
                   er->ExceptionCode, de.u.Exception.dwFirstChance, er->ExceptionAddress);
            if (er->NumberParameters >= 2)
                printf("  accessType=%llu faultAddr=%p\n",
                       (unsigned long long)er->ExceptionInformation[0],
                       (void*)er->ExceptionInformation[1]);

            if (er->ExceptionCode == 0xC0000005) {
                if (!de.u.Exception.dwFirstChance && !secondChanceSeen) {
                    secondChanceSeen = true;
                    // 不再交给 WER (本机多半拿不到 dump), 直接以异常码终止, 防挂死.
                    TerminateProcess(pi.hProcess, er->ExceptionCode);
                    cont = DBG_CONTINUE;
                    ContinueDebugEvent(de.dwProcessId, de.dwThreadId, cont);
                    continue;
                }
                if (!captured && de.u.Exception.dwFirstChance) {
                    captured = true;
                    printFaultModule(pi.hProcess, er->ExceptionAddress);
                    HANDLE hThread = OpenThread(THREAD_GET_CONTEXT | THREAD_SUSPEND_RESUME,
                                                FALSE, de.dwThreadId);
                    CONTEXT ctx;
                    ZeroMemory(&ctx, sizeof(ctx));
                    ctx.ContextFlags = CONTEXT_FULL;
                    if (hThread && GetThreadContext(hThread, &ctx)) {
                        EXCEPTION_POINTERS ep;
                        ep.ExceptionRecord = er;
                        ep.ContextRecord = &ctx;
                        printStack(pi.hProcess, hThread, &ep);
                        HANDLE hf = CreateFileW(argv[2], GENERIC_WRITE, 0, NULL,
                                                CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
                        if (hf != INVALID_HANDLE_VALUE) {
                            MINIDUMP_EXCEPTION_INFORMATION mei;
                            mei.ThreadId = de.dwThreadId;
                            mei.ExceptionPointers = &ep;
                            mei.ClientPointers = FALSE;
                            MiniDumpWriteDump(pi.hProcess, pi.dwProcessId, hf,
                                              MiniDumpWithDataSegs, &mei, NULL, NULL);
                            CloseHandle(hf);
                            printf("DUMP=%ls\n", argv[2]);
                        }
                    }
                    if (hThread) CloseHandle(hThread);
                    fflush(stdout);
                    cont = DBG_EXCEPTION_NOT_HANDLED;
                }
            }
        } else if (de.dwDebugEventCode == EXIT_PROCESS_DEBUG_EVENT) {
            exitCode = de.u.ExitProcess.dwExitCode;
            ContinueDebugEvent(de.dwProcessId, de.dwThreadId, DBG_CONTINUE);
            break;
        }
        ContinueDebugEvent(de.dwProcessId, de.dwThreadId, cont);
    }
    printf("EXIT_CODE=0x%08lX (%ld)\n", exitCode, (long)exitCode);
    // 退出码 = 目标退出码 (= 异常码). 脚本据此判定, 不按"普通失败"处理.
    return (int)exitCode;
}
