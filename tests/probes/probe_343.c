/* probe_343.c — Account #610 knife #343: quantitative probe for ByVal String copy leak.
 *
 * Replicates both versions of the callee body (before / after) verbatim, runs N times each in one process, counts:
 *   alloc = number of SysAllocStringLen calls
 *   free  = number of SysFreeString calls
 *   diff  = net leaked count
 *
 * BEFORE: copy hangs on the param, never freed at exit => leaks one per call.
 * AFTER: copy lives in its own variable, exit frees it only when the param still points at it.
 *
 * 探针自己劫持 SysAllocStringLen / SysFreeString 的计数 (用自定义 allocator 计数),
 * 不用 Windows API 挂钩 (免权限) —— 直接复制 C3 生成的形态, 手工实现 BSTR 分配。
 */
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>

typedef wchar_t *BSTR;

static long g_alloc = 0;
static long g_free = 0;

/* 与 C3 的 vb6_BSTR_FromBSTR / vb6_BSTR_Free 同语义 (Win32 口径) */
static BSTR p_FromBSTR(BSTR bstr) {
    UINT len = 0;
    BSTR r = NULL;
    if (!bstr) {
        r = SysAllocStringLen(L"", 0);
        g_alloc++;
        return r;
    }
    len = SysStringLen(bstr);
    r = SysAllocStringLen(NULL, len);
    g_alloc++;
    if (r && len) memcpy(r, bstr, (size_t)len * sizeof(WCHAR));
    if (r) r[len] = L'\0';
    return r;
}

static void p_Free(BSTR bstr) {
    if (bstr) {
        g_free++;
        SysFreeString(bstr);
    }
}

static BSTR p_FromStr(const wchar_t *s) {
    BSTR r = SysAllocStringLen(s, (UINT)wcslen(s));
    g_alloc++;
    return r;
}

/* ---- 修前版: 副本挂在形参上, 出口不释放 ---- */
static void before_sink(BSTR s) {
    s = p_FromBSTR(s);
    (void)s;
    /* 出口无 free  (这是修前的真实形态) */
}

/* ---- 修后版: 副本另立变量, 出口条件 free ---- */
static void after_sink(BSTR s) {
    BSTR vb6_byvalcopy_s = p_FromBSTR(s);
    s = vb6_byvalcopy_s;
    if (s == vb6_byvalcopy_s) { p_Free(s); }
}

/* ---- 反面针: 体里赋值过的形参, 出口必须跳过 (防 double free) ---- */
static void after_sink_assigned(BSTR s) {
    BSTR vb6_byvalcopy_s = p_FromBSTR(s);
    s = vb6_byvalcopy_s;
    /* AssignMove 语义: *target != source ⇒ 先 free 旧值 (= 副本), 再持有 source */
    p_Free(s);
    s = p_FromStr(L"REPLACED");      /* s 现在指向新串, 不再是副本 */
    if (s == vb6_byvalcopy_s) { p_Free(s); }  /* 出口: 必须跳过 */
}

int main(void) {
    const int N = 20000;
    int i;
    BSTR arg;

    /* warm up */
    arg = p_FromStr(L"0123456789ABCDEF0123456789ABCDEF");
    before_sink(arg);
    p_Free(arg);

    /* ---- 修前 ---- */
    g_alloc = g_free = 0;
    for (i = 0; i < N; i++) {
        BSTR a = p_FromStr(L"0123456789ABCDEF0123456789ABCDEF");
        before_sink(a);
        p_Free(a);
    }
    long ba = g_alloc, bf = g_free;
    printf("BEFORE: alloc=%ld free=%ld net_leak=%ld (per_call=%.3f)\n",
           ba, bf, ba - bf, (double)(ba - bf) / N);

    /* ---- 修后 ---- */
    g_alloc = g_free = 0;
    for (i = 0; i < N; i++) {
        BSTR a = p_FromStr(L"0123456789ABCDEF0123456789ABCDEF");
        after_sink(a);
        p_Free(a);
    }
    long aa = g_alloc, af = g_free;
    printf("AFTER : alloc=%ld free=%ld net_leak=%ld (per_call=%.3f)\n",
           aa, af, aa - af, (double)(aa - af) / N);

    /* ---- 反面针: 修后 + 体里赋值 ⇒ 净泄漏应与「修后无赋值」相同 (不双释放) ---- */
    g_alloc = g_free = 0;
    for (i = 0; i < N; i++) {
        BSTR a = p_FromStr(L"0123456789ABCDEF0123456789ABCDEF");
        after_sink_assigned(a);
        p_Free(a);
    }
    long ra = g_alloc, rf = g_free;
    printf("AFTER+ASSIGNED: alloc=%ld free=%ld net_leak=%ld (per_call=%.3f)\n",
           ra, rf, ra - rf, (double)(ra - rf) / N);

    /* 判据:
     *   修前每调用净漏 1 枚 (副本) ⇒ 证伪「不释放就漏」;
     *   修后净漏 0 枚 ⇒ 证成「出口条件 free 收掉了副本」;
     *   反面针: 体里赋值那一档的净 +1 枚是**新串**本身 (AssignMove 先 free 副本、
     *           再交出一枚新串), 不是副本 ⇒ 判据是「free 不超 alloc」(无双释放)
     *           且「该档 free 数 == 修前档 free 数 + 调用数」(副本确实被 free 过一次)。
     */
    int ok = 1;
    if (ba - bf != N) { printf("FAIL: before must leak exactly 1 per call, got %ld\n", ba - bf); ok = 0; }
    if (aa - af != 0) { printf("FAIL: after must leak 0, got %ld\n", aa - af); ok = 0; }
    if (rf > ra) { printf("FAIL: negative-case free(%ld) > alloc(%ld) => double free\n", rf, ra); ok = 0; }
    if (rf != bf + N) {
        printf("FAIL: negative-case free(%ld) must be before free(%ld) + %d, got mismatch\n",
               rf, bf, N);
        ok = 0;
    }
    printf(ok ? "PROBE343-OK\n" : "PROBE343-BAD\n");
    return ok ? 0 : 1;
}
