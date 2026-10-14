# Account #257 layer-18 follow-up: fast-fail (0xC000037D) on-site capturer.
#
# fast-fail bypasses VEH / UEF / WER / AppVerifier logging, but an attached *debugger*
# still receives it as a first-chance EXCEPTION_DEBUG_EVENT. This script launches
# ComCtlsDemo.exe as a debuggee and, at the crash exception, grabs the WOW64 x86
# CONTEXT + stack + EBP chain into fastfail_dbg.log. RVAs are then symbolized against
# the same-batch .map (see .temp/sym.py).
#
# Layout notes (python host = 64-bit, debuggee = 32-bit/WOW64):
#   DEBUG_EVENT(MSVC x64): code@0 pid@4 tid@8 <pad@12> u@16   (union is 8-aligned!)
#   CREATE_PROCESS_DEBUG_INFO: hFile@16 hProcess@24 hThread@32 lpBaseOfImage@40(8B)
#   EXCEPTION_RECORD64:        code@16 flags@20 rec@24 addr@32(8B) nparams@40 pad@44
#                              info@48 .. dwFirstChance@168
#   EXIT_PROCESS_DEBUG_INFO:   dwExitCode@16
#   WOW64 x86 CONTEXT:         Eax@184 Ebp@188 Eip@192 SegCs@196 EFlags@200 Esp@204
import ctypes, struct, sys
from ctypes import wintypes

k = ctypes.WinDLL('kernel32', use_last_error=True)

DEBUG_PROCESS = 0x1
DEBUG_ONLY_THIS_PROCESS = 0x2
EXCEPTION_DEBUG_EVENT = 1
CREATE_PROCESS_DEBUG_EVENT = 3
EXIT_PROCESS_DEBUG_EVENT = 5
DBG_EXCEPTION_NOT_HANDLED = 0x80010001
DBG_CONTINUE = 0x00010002
THREAD_GET_CONTEXT = 0x0008
CONTEXT_X86_FULL = 0x10007

CRASH_CODES = {0xC0000374, 0xC0000375, 0xC000037D, 0xC000037E, 0xC0000409,
               0xC0000005, 0xC000001D, 0xC0000094}

EVTNAME = {1: 'EXCEPTION', 2: 'CREATE_THREAD', 3: 'CREATE_PROCESS', 4: 'EXIT_THREAD',
           5: 'EXIT_PROCESS', 6: 'LOAD_DLL', 7: 'UNLOAD_DLL', 8: 'OUTPUT_DEBUG_STRING',
           9: 'RIP'}
U = 16


class STARTUPINFO(ctypes.Structure):
    _fields_ = [('cb', wintypes.DWORD), ('lpReserved', wintypes.LPWSTR),
                ('lpDesktop', wintypes.LPWSTR), ('lpTitle', wintypes.LPWSTR),
                ('dwX', wintypes.DWORD), ('dwY', wintypes.DWORD),
                ('dwXSize', wintypes.DWORD), ('dwYSize', wintypes.DWORD),
                ('dwXCountChars', wintypes.DWORD), ('dwYCountChars', wintypes.DWORD),
                ('dwFillAttribute', wintypes.DWORD), ('dwFlags', wintypes.DWORD),
                ('wShowWindow', wintypes.WORD), ('cbReserved2', wintypes.WORD),
                ('lpReserved2', ctypes.POINTER(wintypes.BYTE)),
                ('hStdInput', wintypes.HANDLE), ('hStdOutput', wintypes.HANDLE),
                ('hStdError', wintypes.HANDLE)]


class PROCESS_INFORMATION(ctypes.Structure):
    _fields_ = [('hProcess', wintypes.HANDLE), ('hThread', wintypes.HANDLE),
                ('dwProcessId', wintypes.DWORD), ('dwThreadId', wintypes.DWORD)]


k.CreateProcessW.restype = wintypes.BOOL
k.CreateProcessW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, ctypes.c_void_p,
                             ctypes.c_void_p, wintypes.BOOL, wintypes.DWORD,
                             ctypes.c_void_p, wintypes.LPCWSTR,
                             ctypes.POINTER(STARTUPINFO), ctypes.POINTER(PROCESS_INFORMATION)]
k.WaitForDebugEvent.restype = wintypes.BOOL
k.WaitForDebugEvent.argtypes = [ctypes.c_void_p, wintypes.DWORD]
k.ContinueDebugEvent.restype = wintypes.BOOL
k.ContinueDebugEvent.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.DWORD]
k.OpenThread.restype = wintypes.HANDLE
k.OpenThread.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
k.GetThreadContext.restype = wintypes.BOOL
k.GetThreadContext.argtypes = [wintypes.HANDLE, ctypes.c_void_p]
k.ReadProcessMemory.restype = wintypes.BOOL
k.ReadProcessMemory.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p,
                                ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
k.TerminateProcess.restype = wintypes.BOOL
k.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
k.DebugBreakProcess.restype = wintypes.BOOL
k.DebugBreakProcess.argtypes = [wintypes.HANDLE]
try:
    k.GetTickCount64.restype = ctypes.c_ulonglong
    k.GetTickCount64.argtypes = []
    HAVE_TICK64 = True
except AttributeError:
    HAVE_TICK64 = False

HAVE_WOW64_CTX = True
try:
    k.Wow64GetThreadContext.restype = wintypes.BOOL
    k.Wow64GetThreadContext.argtypes = [wintypes.HANDLE, ctypes.c_void_p]
except AttributeError:
    HAVE_WOW64_CTX = False


def main():
    exe = sys.argv[1] if len(sys.argv) > 1 else r'D:\vb_yqt4qPac\ComCtlsDemo\c3out\ComCtlsDemo.exe'
    cwd = sys.argv[2] if len(sys.argv) > 2 else r'D:\vb_yqt4qPac\ComCtlsDemo\c3out'
    si = STARTUPINFO()
    si.cb = ctypes.sizeof(STARTUPINFO)
    pi = PROCESS_INFORMATION()
    ok = k.CreateProcessW(exe, None, None, None, False,
                          DEBUG_PROCESS | DEBUG_ONLY_THIS_PROCESS, None, cwd,
                          ctypes.byref(si), ctypes.byref(pi))
    if not ok:
        print('CreateProcess failed err=', ctypes.get_last_error())
        return 1
    h_proc = pi.hProcess
    img_base = [0]
    # 账 #257 第二十二层: DebugBreakProcess 打断的是**所有**线程, 先到的往往是某个
    # worker/ntdll 线程 (EBP=0、EIP 落在 64 位 ntdll), 采样就白抓了。要的是**主线程**:
    # 记下 CREATE_PROCESS 给的 tid, 打断后把不属于主线程的那个断点续掉继续等。
    main_tid = [pi.dwThreadId]
    log = open(r'D:\vb_yqt4qPac\ComCtlsDemo\c3out\fastfail_dbg.log', 'w')

    def out(s, echo=True):
        log.write(s + '\n')
        log.flush()
        if echo:
            print(s)

    out('dbg: started pid=%d wow64ctx=%s' % (pi.dwProcessId, HAVE_WOW64_CTX))
    # 账 #257 第二十一层: 卡死(无异常)时的栈采样 —— 到点 DebugBreakProcess 打断, 在随之
    # 而来的 0x80000003 上抓同一套 CONTEXT / EBP 链。第 3 参 = 毫秒 (0 = 不打断)。
    break_ms = int(sys.argv[3]) if len(sys.argv) > 3 else 0
    t0 = k.GetTickCount64() if HAVE_TICK64 else 0
    break_fired = False
    if break_ms:
        out('dbg: will DebugBreakProcess after %d ms' % break_ms)
    buf = ctypes.create_string_buffer(16 + 256)
    counts = {}
    captured = False
    n = 0
    while n < 400000:
        n += 1
        if break_ms and not break_fired and HAVE_TICK64 \
                and (k.GetTickCount64() - t0) >= break_ms:
            break_fired = True
            out('dbg: DebugBreakProcess fired at %d ms' % break_ms)
            k.DebugBreakProcess(h_proc)
        if not k.WaitForDebugEvent(ctypes.cast(buf, ctypes.c_void_p), 60000):
            out('dbg: WaitForDebugEvent timeout/err=%d' % ctypes.get_last_error())
            break
        b = bytes(buf)
        evt, pid, tid = struct.unpack_from('<III', b, 0)
        counts[evt] = counts.get(evt, 0) + 1
        if n <= 24:
            out('dbg: raw evt=%d(%s) u[0:32]=%s' %
                (evt, EVTNAME.get(evt, '?'), b[U:U + 32].hex()), echo=False)

        if evt == CREATE_PROCESS_DEBUG_EVENT:
            img_base[0] = struct.unpack_from('<Q', b, U + 24)[0]  # lpBaseOfImage
            main_tid[0] = tid  # CREATE_PROCESS 的 tid 就是主线程
            out('dbg: CREATE_PROCESS image_base=0x%08x main_tid=%d' % (img_base[0], tid), echo=False)
        elif evt == EXCEPTION_DEBUG_EVENT:
            code = struct.unpack_from('<I', b, U)[0]
            flags = struct.unpack_from('<I', b, U + 4)[0]
            addr = struct.unpack_from('<Q', b, U + 16)[0]
            npar = struct.unpack_from('<I', b, U + 24)[0]
            first = struct.unpack_from('<I', b, U + 152)[0]
            params = [hex(struct.unpack_from('<Q', b, U + 32 + 8 * i)[0])
                      for i in range(min(npar, 4))]
            if code == 0x80000003 and not break_fired:
                # WOW64 initial breakpoint / int3: let it pass through.
                k.ContinueDebugEvent(pid, tid, DBG_CONTINUE)
                continue
            # 采样断点落在这里。⚠ DebugBreakProcess **不是**给每个线程下 int3, 而是
            # 注入一枚**远程线程**去跑 DbgBreakPoint —— 所以这枚 EXCEPTION 的 tid
            # 永远是那枚注入线程, 它的 CONTEXT 没有用。要的是**主线程**的栈:
            # 下面取 CONTEXT 时按 main_tid 取 (见 h_th 那一行), 这里不续掉它。
            if code not in CRASH_CODES and not (code == 0x80000003 and break_fired):
                if counts[evt] <= 40:
                    out('dbg: EXC code=0x%08x addr=0x%x first=%d n=%d params=%s'
                        % (code, addr, first, npar, params), echo=False)
                k.ContinueDebugEvent(pid, tid, DBG_CONTINUE)
                continue
            out('dbg: EXC code=0x%08x addr=0x%x first=%d nparam=%d params=%s'
                % (code, addr, first, npar, params))
            if captured:
                k.ContinueDebugEvent(pid, tid, DBG_EXCEPTION_NOT_HANDLED)
                continue
            captured = True
            lo = int(img_base[0]) if img_base[0] else 0
            hi = lo + 0x1000000
            # 崩溃时取当事线程; DebugBreakProcess 采样时取**主线程** (见上一条注释)。
            h_th = k.OpenThread(THREAD_GET_CONTEXT, False,
                                main_tid[0] if (code == 0x80000003 and break_fired) else tid)
            ctx = (ctypes.c_ubyte * 716)()
            struct.pack_into('<I', ctx, 0, CONTEXT_X86_FULL)
            got = False
            if h_th:
                if HAVE_WOW64_CTX:
                    got = bool(k.Wow64GetThreadContext(h_th, ctypes.cast(ctx, ctypes.c_void_p)))
                if not got:
                    got = bool(k.GetThreadContext(h_th, ctypes.cast(ctx, ctypes.c_void_p)))
            if got:
                # WOW64 x86 CONTEXT offsets: Edi@156 Esi@160 Ebx@164 Edx@168 Ecx@172
                # Eax@176 Ebp@180 Eip@184 SegCs@188 EFlags@192 Esp@196 SegSs@200
                edi = struct.unpack_from('<I', ctx, 156)[0]
                esi = struct.unpack_from('<I', ctx, 160)[0]
                ebx = struct.unpack_from('<I', ctx, 164)[0]
                edx = struct.unpack_from('<I', ctx, 168)[0]
                ecx = struct.unpack_from('<I', ctx, 172)[0]
                eax = struct.unpack_from('<I', ctx, 176)[0]
                ebp = struct.unpack_from('<I', ctx, 180)[0]
                eip = struct.unpack_from('<I', ctx, 184)[0]
                segcs = struct.unpack_from('<I', ctx, 188)[0]
                eflags = struct.unpack_from('<I', ctx, 192)[0]
                esp = struct.unpack_from('<I', ctx, 196)[0]
                out('CAPTURE code=0x%08x excAddr=0x%x' % (code, addr))
                out('CAPTURE EIP=0x%08x ESP=0x%08x EBP=0x%08x' % (eip, esp, ebp))
                out('CAPTURE EAX=0x%08x EBX=0x%08x ECX=0x%08x EDX=0x%08x ESI=0x%08x EDI=0x%08x'
                    % (eax, ebx, ecx, edx, esi, edi))
                out('CAPTURE SegCs=0x%x EFlags=0x%x' % (segcs, eflags))
                out('CAPTURE img_base=0x%08x' % lo)
                if lo:
                    out('CAPTURE EIP_rva=0x%x' % (eip - lo))
                sbuf = (ctypes.c_ubyte * 2048)()
                nr = ctypes.c_size_t(0)
                if k.ReadProcessMemory(h_proc, ctypes.c_void_p(esp), sbuf, 2048,
                                       ctypes.byref(nr)):
                    out('CAPTURE --- stack dwords pointing into exe (rva) ---')
                    vals = struct.unpack_from('<512I', bytes(sbuf))
                    for i, v in enumerate(vals):
                        if lo and lo <= v < hi:
                            out('CAPTURE   esp+0x%03x = 0x%08x  rva=0x%x' % (i * 4, v, v - lo))
                else:
                    out('CAPTURE ReadProcessMemory(ESP) failed err=%d' % ctypes.get_last_error())
                out('CAPTURE --- ebp chain ---')
                cur = ebp
                for _ in range(32):
                    b2 = (ctypes.c_ubyte * 8)()
                    nr2 = ctypes.c_size_t(0)
                    if not k.ReadProcessMemory(h_proc, ctypes.c_void_p(cur), b2, 8,
                                               ctypes.byref(nr2)):
                        break
                    nxt, ret = struct.unpack_from('<II', bytes(b2))
                    if nxt <= cur:
                        break
                    tag = (' rva=0x%x' % (ret - lo)) if (lo and lo <= ret < hi) else ''
                    out('CAPTURE   ebp=0x%08x ret=0x%08x%s' % (cur, ret, tag))
                    cur = nxt
            else:
                out('CAPTURE ctx failed err=%d' % ctypes.get_last_error())
            out('CAPTURE done; terminating debuggee')
            try:
                k.TerminateProcess(h_proc, 0xC000037D)
            except Exception:
                pass
            break
        elif evt == EXIT_PROCESS_DEBUG_EVENT:
            out('dbg: exit EXIT_PROCESS code=0x%08x' % struct.unpack_from('<I', b, U)[0])
            break
        k.ContinueDebugEvent(pid, tid, DBG_CONTINUE)
    summary = sorted('%s=%d' % (EVTNAME.get(e, str(e)), c) for e, c in counts.items())
    out('dbg: events=%d %s' % (n, ' '.join(summary)))
    log.close()
    return 0


sys.exit(main())
