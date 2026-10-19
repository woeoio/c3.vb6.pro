/* ============================================================
 * ai/022 ledger 305, B188 (knife 90 follow-up, task 414): a REAL client-side reading of
 * the "way out" tail, so the fix is pinned by behaviour and not only by the shape of
 * the generated C.
 *
 * Why this exists: `IUnknown::Release` reaching zero runs desc->destroyFunc ->
 * vb6_cls_<X>_Destroy -> Class_Terminate. Before knife 90 the error frame lived in
 * ComObj_Release and therefore longjmped over the *tail* of _Destroy (field BSTRs,
 * array descriptors, vb6_Free(me)), so every release whose Class_Terminate raised
 * leaked one instance. Measured by this tool: ~37 bytes per release on x64, ~17 on x86,
 * and ~120 when the class owns one 20-character String field.
 *
 * Usage:  leak_probe.exe <dll> <raisingClsid> <cleanClsid> <N>
 * Prints, for each class, the growth of this process's committed memory over N
 * create/release cycles plus a verdict. One instance is created and dropped *before*
 * the window opens, so lazy first-instance work (typelib load and friends) stays out of
 * the measured window, then LEAK_WARMUP_CYCLES more cycles run before it opens, because the
 * raising path's first few trips allocate heap the window would otherwise be charged for.
 * The verdict - not the raw byte count - is what the gate pins:
 * a healthy build must stay inside one page, while the leak is three orders of magnitude
 * over that, so CI heap noise cannot make this red by accident.
 *
 * Late-bound readings on these classes are not asserted here: that is #412's territory.
 * ============================================================ */
#define COBJMACROS
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <psapi.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static const IID IID_IClassFactory_local =
    { 0x00000001, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };

/* One page. Deliberately generous: the pre-knife-90 leak at the N used by the gate is
 * ~175 KB on x64 and ~85 KB on x86, so this cannot hide the defect it is meant to catch. */
#define LEAK_TOLERANCE_BYTES 4096
/* Cycles burned before the window opens. Must be big enough to take every one-off allocation of
 * the path under test, and small enough that a real per-instance leak cannot hide in it. */
#define LEAK_WARMUP_CYCLES 200

typedef HRESULT(STDAPICALLTYPE *PFN_GETCLASSOBJ)(REFCLSID, REFIID, LPVOID *);

static int parse_guid(const char *s, CLSID *out) {
    WCHAR w[64];
    if (mbstowcs(w, s, 64) == (size_t)-1) return 0;
    return SUCCEEDED(CLSIDFromString(w, out));
}

static SIZE_T committed_now(void) {
    PROCESS_MEMORY_COUNTERS_EX pmc;
    pmc.cb = sizeof(pmc);
    if (!GetProcessMemoryInfo(GetCurrentProcess(), (PROCESS_MEMORY_COUNTERS *)&pmc, sizeof(pmc)))
        return (SIZE_T)-1;
    return pmc.PagefileUsage;
}

/* Create and release n instances; returns the committed-memory growth over that window.
 * A LEAK_WARMUP-cycle preamble runs first on purpose: the raising path allocates its error-object
 * strings and its stderr buffer the first few times it is taken, and that one-off heap-segment
 * growth (measured: ~7 pages) must not be read as a per-instance leak. Warm up with ONE instance
 * instead and the very first errored release lands inside the window - that is what made a fixed
 * build report delta=28672 at N=5000 while the same build read delta=0 at N=25000. */
static SIZE_T leak_cycle(PFN_GETCLASSOBJ getClassObject, const CLSID *clsid, long n) {
    SIZE_T before;
    SIZE_T after;
    IClassFactory *fac = NULL;
    IDispatch *warm = NULL;
    long i;
    HRESULT hr;

    hr = getClassObject(clsid, &IID_IClassFactory_local, (void **)&fac);
    if (FAILED(hr) || !fac) { printf("LEAK FACTORY hr=0x%08lX\n", (unsigned long)hr); return (SIZE_T)-1; }
    hr = IClassFactory_CreateInstance(fac, NULL, &IID_IDispatch, (void **)&warm);
    if (FAILED(hr) || !warm) {
        printf("LEAK WARM-UP hr=0x%08lX\n", (unsigned long)hr);
        IClassFactory_Release(fac);
        return (SIZE_T)-1;
    }
    IDispatch_Release(warm);
    for (i = 0; i < LEAK_WARMUP_CYCLES; i++) {
        IDispatch *w = NULL;
        hr = IClassFactory_CreateInstance(fac, NULL, &IID_IDispatch, (void **)&w);
        if (FAILED(hr) || !w) {
            printf("LEAK WARM-CYCLE-FAIL i=%ld hr=0x%08lX\n", i, (unsigned long)hr);
            break;
        }
        IDispatch_Release(w);
    }

    before = committed_now();
    for (i = 0; i < n; i++) {
        IDispatch *d = NULL;
        hr = IClassFactory_CreateInstance(fac, NULL, &IID_IDispatch, (void **)&d);
        if (FAILED(hr) || !d) {
            printf("LEAK CREATE-FAIL i=%ld hr=0x%08lX\n", i, (unsigned long)hr);
            break;
        }
        IDispatch_Release(d);
    }
    after = committed_now();
    IClassFactory_Release(fac);
    return after - before;
}

/* Tag and verdict print ADJACENT on purpose: the caller flattens the whole log and greps
 * for "LEAK-BOOM verdict=OK", so any number printed between them would break the match.
 * delta/n stay as the human reading. */
static void report(const char *tag, SIZE_T delta, long n) {
    if (delta == (SIZE_T)-1) {
        printf("LEAK-%s verdict=BROKEN\n", tag);
    } else if (delta <= LEAK_TOLERANCE_BYTES) {
        printf("LEAK-%s verdict=OK delta=%lu n=%ld\n", tag, (unsigned long)delta, n);
    } else {
        printf("LEAK-%s verdict=LEAK delta=%lu n=%ld\n", tag, (unsigned long)delta, n);
    }
    fflush(stdout);
}

int main(int argc, char **argv) {
    HMODULE mod;
    PFN_GETCLASSOBJ getClassObject;
    CLSID boom;
    CLSID good;
    long n;

    if (argc < 5) { printf("usage: leak_probe <dll> <raisingClsid> <cleanClsid> <N>\n"); return 2; }
    n = atol(argv[4]);
    if (n < 1) n = 1;
    mod = LoadLibraryExA(argv[1], NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
    if (!mod) { printf("LOADFAIL err=%lu\n", GetLastError()); return 2; }
    getClassObject = (PFN_GETCLASSOBJ)GetProcAddress(mod, "DllGetClassObject");
    if (!getClassObject) { printf("NOGETCLASSOBJ\n"); return 2; }
    if (!parse_guid(argv[2], &boom) || !parse_guid(argv[3], &good)) { printf("CLSIDPARSE\n"); return 2; }

    CoInitialize(NULL);
    printf("LEAK-N=%ld tolerance=%d\n", n, LEAK_TOLERANCE_BYTES);
    fflush(stdout);
    /* Raising first: if this path ever kills the client again, the clean-class control has
     * still not been printed, which is itself the reading we want on the log. */
    report("BOOM", leak_cycle(getClassObject, &boom, n), n);
    report("GOOD", leak_cycle(getClassObject, &good, n), n);
    CoUninitialize();
    printf("LEAK-DONE\n");
    fflush(stdout);
    return 0;
}
