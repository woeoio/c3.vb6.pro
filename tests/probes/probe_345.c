/* probe_345.c - Account #610 knife #345: runtime proof for the MouseEnter/MouseLeave
 * dispatch mismatch (ByVal value vs ByRef pointer) on array controls.
 *
 * Background:
 *   The generated C had, for the SAME handler:
 *     definition: static void vb6_ucSlider1_MouseEnter(int16_t* vIndex);
 *     dispatch  : { extern void vb6_ucSlider1_MouseEnter(int16_t); vb6_ucSlider1_MouseEnter(3); }
 *   i.e. the definition takes a POINTER (VB6 default ByRef) while the dispatch
 *   passed a VALUE. MSVC only warns (C4133), the binary links fine, and at
 *   runtime the callee dereferences the integer 3 as an address.
 *
 *   BEFORE: dispatch passes the literal index 3  -> callee reads *(int16_t*)3 -> access violation
 *   AFTER : dispatch passes &vb6_evindex         -> callee reads the real runtime index
 *
 * This probe replicates both dispatch forms verbatim and reports what each one
 * does. Compile with MSVC. ASCII only, on purpose: the stdout strings are the
 * verdict and must be byte-stable.
 */
#include <windows.h>
#include <stdio.h>
#include <stdint.h>

/* ---- the generated handler (VB6: Private Sub ucSlider1_MouseEnter(INDEX As Integer)) ---- */
static void handler(int16_t* vIndex) {
    /* body: Set mCtlSlider = ucSlider1(INDEX) -> reads *INDEX */
    printf("handler saw index=%d\n", (int)*vIndex);
    fflush(stdout);
}

/* runtime index as installed by the array-element setup (GetPropW - 1) */
static int16_t g_runtime_index = 2;

int main(void) {
    int ok = 1;

    /* ---- AFTER: the fixed dispatch form (address of the runtime index) ---- */
    {
        int16_t vb6_evindex = g_runtime_index;   /* GetPropW(...) - 1 */
        printf("AFTER : dispatching with &vb6_evindex (value=%d)\n", (int)vb6_evindex);
        fflush(stdout);
        handler(&vb6_evindex);
    }

    /* ---- BEFORE: the old dispatch form (the literal index passed as a value) ----
       The callee dereferences whatever it received as a pointer, so the old form
       hands it the integer 3. Reading address 3 is an access violation on both
       x86 and x64. Catch it with SEH and report it as the expected failure. */
    {
        int faulted = 0;
        SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOGPFAULTERRORBOX);
        __try {
            /* verbatim semantics of `fn(3)` where fn's real prototype is fn(int16_t*) */
            handler((int16_t *)(void *)(INT_PTR)3);
        } __except (EXCEPTION_EXECUTE_HANDLER) {
            faulted = 1;
        }
        if (faulted) {
            printf("BEFORE: FAULT as expected (the old form handed the callee a VALUE\n"
                   "        that it dereferenced as an address)\n");
        } else {
            printf("BEFORE: no fault -- UNEXPECTED, the probe could not reproduce it\n");
            ok = 0;
        }
    }

    printf(ok ? "PROBE345-OK\n" : "PROBE345-BAD\n");
    return ok ? 0 : 1;
}
