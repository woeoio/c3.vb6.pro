/* disp_probe.c - ai/022 B14: 进程内真客户端，驱动一个真编译出来的 ActiveX DLL 走 IDispatch 四件套。
 *
 * 为什么要有它：B13a-B13e 五格里，DLL 那条对外管线只被**字节级**验过（dll_entry.c 的表、
 * .tlb 里的 GUID），从来没有调用者真的拿着产出的指针走过一次 GetIDsOfNames/Invoke。
 * 本探针不查注册表：LoadLibrary + DllGetClassObject + CreateInstance(IID_IDispatch)，
 * 所以 CI 上不需要 regsvr32，也不依赖机器上装过什么组件。
 *
 * 用法：disp_probe <dll 路径> <clsid "{...}"> [另一个 iid "{...}"，看 QI 应答的是哪份指针]
 * 输出：一行一个 ASCII 读数（用例按子串断言）。
 */
#define COBJMACROS   /* 打开 IDispatch_Xxx / IClassFactory_Xxx 这类 C 形态的内联助手 */
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <oaidl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

typedef HRESULT(STDAPICALLTYPE *PFN_GETCLASSOBJ)(REFCLSID, REFIID, LPVOID *);

static const IID IID_IClassFactory_local =
    {0x00000001, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};

static int parse_guid(const char *text, GUID *out) {
    wchar_t wbuf[64];
    if (MultiByteToWideChar(CP_ACP, 0, text, -1, wbuf, 64) <= 0) return 0;
    return SUCCEEDED(CLSIDFromString(wbuf, out));
}

static void print_guid(const char *tag, const GUID *g) {
    printf("%s={%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}\n",
           tag, (unsigned long)g->Data1, g->Data2, g->Data3,
           g->Data4[0], g->Data4[1], g->Data4[2], g->Data4[3],
           g->Data4[4], g->Data4[5], g->Data4[6], g->Data4[7]);
}

/* 名字 -> dispid -> Invoke 一趟；两件事各自打一行，因为"点不到"分两种。 */
static void call_member(IDispatch *disp, const wchar_t *name, VARIANT *argv, int argc,
                        WORD wFlags, const char *tag) {
    DISPID id = DISPID_UNKNOWN;
    DISPPARAMS dp;
    VARIANT res;
    EXCEPINFO ei;
    UINT err = 0;
    LPOLESTR names[1];
    HRESULT hr;
    int i;

    names[0] = (LPOLESTR)name;
    VariantInit(&res);
    memset(&ei, 0, sizeof(ei));
    memset(&dp, 0, sizeof(dp));
    hr = IDispatch_GetIDsOfNames(disp, &IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    printf("NAMES=%s hr=0x%08lX dispid=%ld\n", tag, (unsigned long)hr, (long)id);
    if (FAILED(hr)) return;

    dp.rgvarg = argv;          /* DISPPARAMS 是逆序，调用方按逆序摆好 */
    dp.cArgs = (UINT)argc;
    if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
        /* Acct #305 SB185 (knife 77): a property write carries its own name slot. The old
         * probe never set it because it only ever called methods. */
        static DISPID s_putName = DISPID_PROPERTYPUT;
        dp.rgdispidNamedArgs = &s_putName;
        dp.cNamedArgs = 1;
    }
    hr = IDispatch_Invoke(disp, id, &IID_NULL, LOCALE_USER_DEFAULT, wFlags, &dp,
                          &res, &ei, &err);
    printf("CALL=%s hr=0x%08lX", tag, (unsigned long)hr);
    /* Acct #305 SB185 (knife 77): DISP_E_EXCEPTION only says "the server raised a VB error" --
     * the error number lives in EXCEPINFO.scode. Printing just the hr would let 0x80020009 be
     * read as a conclusion while the number it carries stayed unseen (same shape as SB144's
     * "right tag, garbage value"). VB runtime descriptions are ASCII, so they go out as bytes.
     * Ownership: the server handed this BSTR over, so the probe frees it below. */
    if (FAILED(hr)) {
        int di, dl = 0;
        printf(" scode=0x%08lX desc=", (unsigned long)ei.scode);
        if (ei.bstrDescription) {
            int asciiOnly = 1;
            dl = (int)SysStringLen(ei.bstrDescription);
            for (di = 0; di < dl && di < 40; di++)
                if (ei.bstrDescription[di] >= 0x80) { asciiOnly = 0; break; }
            /* The VB standard-description table is Chinese, and this probe prints bytes: dumping the
             * low half of those UTF-16 units produces mojibake that no needle can name. Say whether
             * the server handed back the raiser's own ASCII message, or the table's text -- the
             * error number is already pinned by scode. */
            if (asciiOnly) {
                for (di = 0; di < dl && di < 40; di++) printf("%c", (char)ei.bstrDescription[di]);
            } else {
                printf("(non-ascii)");
            }
        } else {
            printf("(null)");
        }
    }
    if (SUCCEEDED(hr) && res.vt == VT_I4) printf(" result=%ld", (long)V_I4(&res));
    else if (SUCCEEDED(hr)) printf(" resultvt=%u", (unsigned)res.vt);
    /* Acct #305 SB185 (knife 77): a class field getter answers VT_I2 / VT_I1, and "a tag came back"
     * is not the reading -- the two heads of that pin are which NUMBER landed. */
    if (SUCCEEDED(hr) && (res.vt == VT_I2 || res.vt == VT_I1 || res.vt == VT_UI1))
        printf(" num=%ld", (long)(res.vt == VT_I2 ? (long)res.iVal
                    : (res.vt == VT_I1 ? (long)res.cVal : (long)res.bVal)));
    /* Acct #305 SB144: a Decimal reply used to print only "resultvt=14" -- the tag was
     * right while the value was garbage, which is exactly why this cell hid for so long.
     * Two locale-free readings: VT_CY (fixed point, 1/10000) and VT_I8. If oleaut32
     * cannot convert the bytes we print FAIL instead of staying silent. */
    if (SUCCEEDED(hr) && res.vt == VT_DECIMAL) {
        VARIANT cy, i8;
        VariantInit(&cy); VariantInit(&i8);
        if (SUCCEEDED(VariantChangeType(&cy, &res, 0, VT_CY)) && cy.vt == VT_CY)
            printf(" dec_cy=%lld", (long long)cy.cyVal.int64);
        else
            printf(" dec_cy=FAIL");
        if (SUCCEEDED(VariantChangeType(&i8, &res, 0, VT_I8)) && i8.vt == VT_I8)
            printf(" dec_i8=%lld", (long long)i8.llVal);
        else
            printf(" dec_i8=FAIL");
        VariantClear(&cy); VariantClear(&i8);
    }
    /* Acct #305 SB159 + SB162 (knife 50): an array that came IN as a real SAFEARRAY has to be
     * readable by the callee and come BACK OUT as one. EchoVal just returns its own argument, so
     * the whole round trip is one call: inbound converts SAFEARRAY -> our descriptor, the reply
     * converts back. Before this knife both directions answered VT_EMPTY (no array arm at all),
     * and the arm that exists today must not deliver a slot whose parray still points at the
     * caller's SAFEARRAY header -- that is the cell SB162 was opened for. */
    if (SUCCEEDED(hr) && (res.vt & VT_ARRAY)) {
        SAFEARRAY* psa = res.parray;
        LONG lo = 0, hi = 0, k;
        long sum = 0;
        printf(" arrvt=%u dims=%u", (unsigned)res.vt,
               psa ? (unsigned)SafeArrayGetDim(psa) : 0u);
        if (psa) {
            SafeArrayGetLBound(psa, 1, &lo);
            SafeArrayGetUBound(psa, 1, &hi);
            printf(" lo=%ld hi=%ld", (long)lo, (long)hi);
            if ((res.vt & VT_TYPEMASK) == VT_I4) {
                for (k = lo; k <= hi; k++) {
                    LONG one = 0;
                    if (FAILED(SafeArrayGetElement(psa, &k, &one))) { sum = -1; break; }
                    sum += (long)one;
                }
                printf(" arrsum=%ld", sum);
            }
        }
    }
    printf("\n");
    if (ei.bstrDescription) { SysFreeString(ei.bstrDescription); ei.bstrDescription = NULL; }
    VariantClear(&res);
    for (i = 0; i < argc; i++) VariantClear(&argv[i]);
}

/* Acct #305 SB144: build a VT_DECIMAL the way a real client does -- through oleaut32,
 * so we never hand-write the arch-specific position of those 16 bytes (x86 has no
 * decVal member at all; x64 puts it at the head). */
static VARIANT mk_dec_i4(LONG v) {
    VARIANT src, out;
    VariantInit(&src); VariantInit(&out);
    src.vt = VT_I4; src.lVal = v;
    if (FAILED(VariantChangeType(&out, &src, 0, VT_DECIMAL)) || out.vt != VT_DECIMAL)
        printf("MKDEC=FAIL vt=%u\n", (unsigned)out.vt);
    return out;
}

static VARIANT mk_i4(LONG v) {
    VARIANT x;
    VariantInit(&x);
    x.vt = VT_I4;
    x.lVal = v;
    return x;
}

/* 账 #481 (第一百八十二刀注销腿的运行时证人, 第一百八十三刀接上): 调 Churn(n), 把 BSTR 答案本身
 * 按 ASCII 打出来。既有 call_member 只打 resultvt (VARIANT 标签) —— 这一格要看的正是**答案内容**:
 * "SatImpl"(表里查得到身份) 还是空(表已饱和, 那枚实例没登记进去)。只在命令行给了 CHURN=<n> 时才发,
 * 其余夹具的输出一个字节都不动。 */
static void call_churn(IDispatch *disp, LONG n) {
    DISPID id = DISPID_UNKNOWN;
    DISPPARAMS dp;
    VARIANT res;
    VARIANT arg;
    EXCEPINFO ei;
    UINT err = 0;
    LPOLESTR names[1];
    HRESULT hr;

    names[0] = (LPOLESTR)L"Churn";
    VariantInit(&res);
    memset(&ei, 0, sizeof(ei));
    memset(&dp, 0, sizeof(dp));
    arg = mk_i4(n);
    hr = IDispatch_GetIDsOfNames(disp, &IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    printf("CHURN_NAMES hr=0x%08lX dispid=%ld\n", (unsigned long)hr, (long)id);
    if (FAILED(hr)) return;
    dp.rgvarg = &arg;
    dp.cArgs = 1;
    hr = IDispatch_Invoke(disp, id, &IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp,
                          &res, &ei, &err);
    printf("CHURN_CALL hr=0x%08lX vt=%u\n", (unsigned long)hr, (unsigned)res.vt);
    if (SUCCEEDED(hr) && res.vt == VT_BSTR && res.bstrVal) {
        UINT i, len = SysStringLen(res.bstrVal);
        printf("CHURN_TYPE=");
        for (i = 0; i < len; i++) {
            unsigned short c = res.bstrVal[i];
            putchar(c < 128 ? (int)c : '?');
        }
        printf("\n");
        if (len == 0) printf("CHURN_EMPTY=1\n");
    } else {
        printf("CHURN_EMPTY=1\n");
    }
    if (ei.bstrDescription) SysFreeString(ei.bstrDescription);
    VariantClear(&res);
    VariantClear(&arg);
}

/* 账 #305 §B219 (第二百二十刀): 一发调用, 实参的 vt 由夹具在命令行上点出来。
 * 既有 call_member 摆的实参档是写死的 VT_I4 —— 那一枚问不出「文本递进数值档 / 数递进文本档」
 * 这一半，而这一半正是本刀的靶子 (异档进槽 = 读错联合体那一格：串档拿到 AV、数值档拿到指针)。
 * 臂形: VTARG=<成员名>:<档>[,<档>...]，档按**声明序**给 (第 0 个 = 第一个形参)，
 *   i=VT_I4(40) t=VT_I2(9) u=VT_UI1(7) f=VT_R4(12.75) r=VT_R8(12.75) b=VT_BOOL(True)
 *   B=VT_BOOL(False) s=VT_BSTR("12") x=VT_BSTR("abc") e=VT_EMPTY n=VT_NULL
 * 只在给了 VTARG= 时发，其余夹具的输出一个字节都不动。 */
static void mk_vtarg(VARIANT *out, char kind) {
    VariantInit(out);
    switch (kind) {
        case 'i': out->vt = VT_I4;   out->lVal = 40; break;
        case 't': out->vt = VT_I2;   out->iVal = 9; break;
        case 'u': out->vt = VT_UI1;  out->bVal = 7; break;
        case 'f': out->vt = VT_R4;   out->fltVal = 12.75f; break;
        case 'r': out->vt = VT_R8;   out->dblVal = 12.75; break;
        case 'b': out->vt = VT_BOOL; out->boolVal = VARIANT_TRUE; break;
        case 'B': out->vt = VT_BOOL; out->boolVal = VARIANT_FALSE; break;
        case 's': out->vt = VT_BSTR; out->bstrVal = SysAllocString(L"12"); break;
        case 'x': out->vt = VT_BSTR; out->bstrVal = SysAllocString(L"abc"); break;
        case 'n': out->vt = VT_NULL; break;
        default:  out->vt = VT_EMPTY; break;
    }
}

static const char *vtarg_kind_name(char kind) {
    switch (kind) {
        case 'i': return "VT_I4";
        case 't': return "VT_I2";
        case 'u': return "VT_UI1";
        case 'f': return "VT_R4";
        case 'r': return "VT_R8";
        case 'b': return "VT_BOOL";
        case 'B': return "VT_BOOL(False)";
        case 's': return "BSTR(12)";
        case 'x': return "BSTR(abc)";
        case 'n': return "VT_NULL";
        default:  return "VT_EMPTY";
    }
}

static void print_vtarg_value(const VARIANT *v) {
    if (v->vt == VT_BSTR && v->bstrVal) {
        UINT i, len = SysStringLen(v->bstrVal);
        printf(" val=<");
        for (i = 0; i < len; i++) {
            unsigned short c = v->bstrVal[i];
            putchar(c < 128 ? (int)c : '?');
        }
        printf("> len=%u", (unsigned)len);
    } else if (v->vt == VT_I4 || v->vt == VT_I2 || v->vt == VT_UI1) {
        printf(" val=%ld", (long)v->lVal);
    } else if (v->vt == VT_R8) {
        printf(" val=%.4f", v->dblVal);
    } else if (v->vt == VT_BOOL) {
        printf(" val=%s", v->boolVal == VARIANT_TRUE ? "True" : "False");
    }
}

static void call_vtarg(IDispatch *disp, const char *spec) {
    char member[64];
    char kinds[8];
    VARIANT rev[8];
    VARIANT res;
    DISPPARAMS dp;
    EXCEPINFO ei;
    DISPID id = DISPID_UNKNOWN;
    UINT err = 0;
    LPOLESTR names[1];
    wchar_t wname[64];
    const char *colon = strchr(spec, ':');
    const char *p;
    size_t mlen = colon ? (size_t)(colon - spec) : strlen(spec);
    int nk = 0, i;
    HRESULT hr;

    if (mlen == 0 || mlen >= sizeof(member)) { printf("VTARG=BAD %s\n", spec); return; }
    memcpy(member, spec, mlen);
    member[mlen] = '\0';
    if (colon) {
        for (p = colon + 1; *p && nk < 8; p++)
            if (*p != ',') kinds[nk++] = *p;
    }
    if (MultiByteToWideChar(CP_ACP, 0, member, -1, wname, 64) <= 0) {
        printf("VTARG=BADNAME %s\n", spec);
        return;
    }
    names[0] = wname;
    VariantInit(&res);
    memset(&ei, 0, sizeof(ei));
    memset(&dp, 0, sizeof(dp));
    hr = IDispatch_GetIDsOfNames(disp, &IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) {
        printf("VTARG=%s names hr=0x%08lX\n", spec, (unsigned long)hr);
        return;
    }
    /* DISPPARAMS 是逆序：rgvarg[0] = 最后一个形参。kinds[] 按声明序给，这里翻过来摆。 */
    for (i = 0; i < nk; i++) mk_vtarg(&rev[nk - 1 - i], kinds[i]);
    dp.rgvarg = nk ? rev : NULL;
    dp.cArgs = (UINT)nk;
    hr = IDispatch_Invoke(disp, id, &IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp,
                          &res, &ei, &err);
    printf("VTARG=%s giving=", spec);
    for (i = 0; i < nk; i++) printf("%s%s", i ? "+" : "", vtarg_kind_name(kinds[i]));
    if (nk == 0) printf("-");
    printf(" hr=0x%08lX vt=%u", (unsigned long)hr, (unsigned)res.vt);
    print_vtarg_value(&res);
    if (FAILED(hr)) {
        printf(" scode=0x%08lX argerr=%u", (unsigned long)ei.scode, (unsigned)err);
        if (ei.bstrDescription) printf(" desc=%s",
            SysStringLen(ei.bstrDescription) >= 1 && ei.bstrDescription[0] >= 0x80 ? "(non-ascii)" : "(ascii)");
    }
    printf("\n");
    if (ei.bstrDescription) SysFreeString(ei.bstrDescription);
    VariantClear(&res);
    for (i = 0; i < nk; i++) VariantClear(&rev[i]);
}

/* Acct #305 SB159 (knife 50): a real one-element-per-slot VT_ARRAY|VT_I4 built the way a
 * real client builds one -- SafeArrayCreate + SafeArrayPutElement, no hand-written layout. */
static VARIANT mk_sa_i4(const LONG* vals, int n) {
    SAFEARRAYBOUND b;
    VARIANT x;
    int i;
    VariantInit(&x);
    b.lLbound = 0;
    b.cElements = (ULONG)n;
    x.parray = SafeArrayCreate(VT_I4, 1, &b);
    if (!x.parray) {
        printf("MKSA=FAIL\n");
        return x;
    }
    for (i = 0; i < n; i++) {
        long k = i;
        if (FAILED(SafeArrayPutElement(x.parray, &k, (void*)&vals[i]))) {
            printf("MKSA=FAIL put=%d\n", i);
            break;
        }
    }
    x.vt = (VARTYPE)(VT_ARRAY | VT_I4);
    return x;
}

/* Acct #305 SB222 (knife 230): array argument arms. Spec: ARRARG=<Method>:<n> (BYREF shape,
 * what a VB6 client sends for `a() As Long`) and ARRARGV=<Method>:<n> (plain VT_ARRAY, the
 * negative control -- no write-back is expected there).
 * The pre-filled elements make "the callee overwrote it" readable without trusting the return. */
static void call_arrarg(IDispatch *disp, const char *spec, int byref) {
    char name[64];
    const char *colon;
    VARIANT a[2];
    VARIANT res;
    DISPPARAMS dp;
    EXCEPINFO ei;
    DISPID id = DISPID_UNKNOWN;
    LPOLESTR names[1];
    UINT err = 0;
    SAFEARRAY *psa, *keep;
    long n = 0, i, got = -1, rd[16];
    long v0 = -1, v1 = -1, v2 = -1, v3 = -1;
    HRESULT hr;

    colon = strchr(spec, ':');
    if (!colon || (size_t)(colon - spec) >= sizeof(name)) {
        printf("ARRARG=%s BADSPEC\n", spec);
        return;
    }
    memcpy(name, spec, (size_t)(colon - spec));
    name[colon - spec] = '\0';
    n = strtol(colon + 1, NULL, 10);
    if (n <= 0 || n > 16) {
        printf("ARRARG=%s BADN n=%ld\n", name, n);
        return;
    }
    for (i = 0; i < n; i++) rd[i] = 0;
    psa = NULL;
    a[0] = mk_sa_i4(rd, (int)n);      /* zeroed first, then the sentinel fill below */
    psa = a[0].parray;
    if (!psa) { printf("ARRARG=%s SA=FAIL\n", name); return; }
    for (i = 0; i < n; i++) {
        long k = i;
        LONG v = (LONG)(0x11111111L * (i + 1));
        SafeArrayPutElement(psa, &k, (void *)&v);
    }
    keep = psa;
    if (byref) {
        /* a[] is ByRef in VB6: the client hands over the ADDRESS of its SAFEARRAY* . */
        a[0].vt = (VARTYPE)(VT_ARRAY | VT_BYREF | VT_I4);
        a[0].pparray = &psa;
    }
    a[1] = mk_i4(n);                  /* DISPPARAMS is reversed: rgvarg[1] = first declared arg */
    VariantInit(&res);
    memset(&ei, 0, sizeof(ei));
    memset(&dp, 0, sizeof(dp));
    dp.rgvarg = a;
    dp.cArgs = 2;
    {
        wchar_t wname[64];
        if (MultiByteToWideChar(CP_ACP, 0, name, -1, wname, 64) <= 0) {
            printf("ARRARG=%s BADNAME\n", name);
            VariantClear(&a[0]);
            VariantClear(&a[1]);
            return;
        }
        names[0] = wname;
        hr = IDispatch_GetIDsOfNames(disp, &IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    }
    if (SUCCEEDED(hr))
        hr = IDispatch_Invoke(disp, id, &IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                              &dp, &res, &ei, &err);
    got = (res.vt == VT_I4) ? (long)res.lVal : -1;
    printf("ARRARG=%s:%ld mode=%s hr=0x%08lX ret=%ld", name, n, byref ? "byref" : "plain",
           (unsigned long)hr, got);
    /* Read the client's own array AFTER the call -- the whole point of the arm. */
    if (psa) {
        long* q = NULL;
        if (SUCCEEDED(SafeArrayAccessData(psa, (void **)&q)) && q) {
            ULONG cnt = psa->rgsabound[0].cElements;
            v0 = q[0];
            v1 = (cnt > 1) ? q[1] : -1;
            v2 = (cnt > 2) ? q[2] : -1;
            v3 = (cnt > 3) ? q[3] : -1;
            printf(" back n=%lu v0=%ld v1=%ld v2=%ld v3=%ld same=%d", (unsigned long)cnt,
                   v0, v1, v2, v3, (int)(psa == keep));
            SafeArrayUnaccessData(psa);
        } else {
            printf(" back=ACCESS_FAIL same=%d", (int)(psa == keep));
        }
    } else {
        printf(" back=NULL");
    }
    printf("\n");
    VariantClear(&res);
    /* BYREF: a[0]'s VariantClear is a no-op on the byref tag, so the (possibly replaced)
     * descriptor is ours to free; plain: VariantClear already destroyed it. */
    VariantClear(&a[1]);
    if (byref) { VariantClear(&a[0]); if (psa) SafeArrayDestroy(psa); }
    else { VariantClear(&a[0]); }
}

/* Acct #305 SB223 (knife 231): scalar ByRef argument arm. Spec: BYREFARG=<Method>:<n>.
 * Hands over VT_BYREF|VT_I4 pointing at a client-owned LONG (what a VB6 client sends for
 * `ByRef v As Long`) and prints both the client's value after the call and the answer. */
static void call_byrefarg(IDispatch *disp, const char *spec, int byref) {
    char name[64];
    const char *colon;
    VARIANT a[1];
    VARIANT res;
    DISPPARAMS dp;
    EXCEPINFO ei;
    DISPID id = DISPID_UNKNOWN;
    LPOLESTR names[1];
    UINT err = 0;
    LONG clientVal;
    HRESULT hr;

    colon = strchr(spec, ':');
    if (!colon || (size_t)(colon - spec) >= sizeof(name)) {
        printf("BYREFARG=%s BADSPEC\n", spec);
        return;
    }
    memcpy(name, spec, (size_t)(colon - spec));
    name[colon - spec] = '\0';
    clientVal = strtol(colon + 1, NULL, 10);
    VariantInit(&a[0]);
    if (byref) {
        a[0].vt = (VARTYPE)(VT_BYREF | VT_I4);
        a[0].piVal = &clientVal;
    } else {
        /* Our own client packs a Long argument BY VALUE (vb6_ComPackInt -> VT_I4);
         * this arm is that shape, so the two ends can be compared. */
        a[0].vt = VT_I4;
        a[0].lVal = clientVal;
    }
    VariantInit(&res);
    memset(&ei, 0, sizeof(ei));
    memset(&dp, 0, sizeof(dp));
    dp.rgvarg = a;
    dp.cArgs = 1;
    {
        wchar_t wname[64];
        if (MultiByteToWideChar(CP_ACP, 0, name, -1, wname, 64) <= 0) {
            printf("BYREFARG=%s BADNAME\n", name);
            return;
        }
        names[0] = wname;
        hr = IDispatch_GetIDsOfNames(disp, &IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    }
    if (SUCCEEDED(hr))
        hr = IDispatch_Invoke(disp, id, &IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                              &dp, &res, &ei, &err);
    printf("%s=%s hr=0x%08lX ret=%ld retvt=%d client_after=%ld\n",
           byref ? "BYREFARG" : "BYVALARG", name,
           (unsigned long)hr, (res.vt == VT_I4) ? (long)res.lVal : -1L, (int)res.vt,
           (long)clientVal);
    VariantClear(&res);
    VariantClear(&a[0]);
}


/* ai/022 B16: 薄指针的槽表形状 (与生成码 vb6_ivtbl_<I> 逐槽对齐). 薄指针本身是
 * `{ vt }` 一字段结构, 所以要**先解一层**取槽表再按槽调 —— 直接 `((probe_ivtbl*)p)->X`
 * 是拿 p 当槽表, 会读到对象后面的堆内存 (实测就是它把探针打崩的). */
typedef struct probe_ivtbl_IProbe {
    long (__stdcall *QueryInterface)(void *self, const void *riid, void **ppv);
    unsigned long (__stdcall *AddRef)(void *self);
    unsigned long (__stdcall *Release)(void *self);
    void (__stdcall *Ping)(void *self, LONG n);   /* 契约槽 0: oVft=(3+0)*指针宽 */
    LONG (__stdcall *Got)(void *self);            /* 契约槽 1: oVft=(3+1)*指针宽 */
} probe_ivtbl_IProbe;

/* ai/022 B16: 按 .tlb 里的形状**直调契约槽** —— 指针首字段是槽表, 槽 3 起是成员
 * (oVft=(3+槽号)*指针宽, CC_STDCALL), 与 TKIND_INTERFACE 那一行一一对应.
 * 本支路写死 cc_dll 的 IProbe 形状: [3]=Ping(ByVal n As Long), [4]=Got() As Long,
 * 实现是 m_got = n*2 (tests\cc_dll\CImpl.cls). 先读一次原值, 再经槽 3 写 21,
 * 再经槽 4 读回 —— 两个槽的身份与顺序都被这一次读数证明. */
static void call_vtable_slots(void *iface) {
    probe_ivtbl_IProbe *vt = *(probe_ivtbl_IProbe **)iface;
    printf("VTBL_GET_BEFORE=%ld\n", (long)vt->Got(iface));
    vt->Ping(iface, 21);
    printf("VTBL_GET_AFTER=%ld\n", (long)vt->Got(iface));
}

/* Acct #305 SB186 (knife 78): the class factory is the SECOND station that walks in from a COM
 * entry point into VB code (CreateInstance -> vb6_ComObject_Create -> factoryFunc -> Class_Initialize),
 * and IClassFactory::CreateInstance has no EXCEPINFO channel at all -- the only thing it can answer
 * with is an HRESULT. Before this knife that station had no error boundary either, so a body that
 * raised an uncaught runtime error killed the client process (exit(123), and the CREATE line below
 * never printed). Each CREATE=<clsid[,clsid...]> argument therefore asks the factory for one object
 * and prints the number it answered -- the error has to be readable in the hr, not in a scode.
 * A listed class that DOES answer a pointer is then asked for Twice(21): "the clean class still
 * constructs and its member still works" is the witness that keeps "the factory now fails for
 * everything" from passing. This block runs last, after the primary object's own readings, so an
 * early death still leaves everything before it on disk. */
static void create_arms(PFN_GETCLASSOBJ getClassObject, const char *spec) {
    char buf[512];
    char *tok;
    char *save;
    IDispatch *d;
    IClassFactory *f;
    VARIANT one[1];
    int idx = 0;
    CLSID c;
    HRESULT hr;

    strncpy(buf, spec, sizeof(buf) - 1);
    buf[sizeof(buf) - 1] = 0;
    for (tok = strtok_s(buf, ",", &save); tok != NULL; tok = strtok_s(NULL, ",", &save)) {
        char tag[64];
        char mtag[32];
        idx++;
        sprintf(tag, "CREATE=%d[%s]", idx, tok);
        sprintf(mtag, "CREATE%dMEM", idx);
        if (!parse_guid(tok, &c)) {
            printf("%s CLSID=BAD\n", tag);
            continue;
        }
        f = NULL;
        d = NULL;
        hr = getClassObject(&c, &IID_IClassFactory_local, (void **)&f);
        if (FAILED(hr) || !f) {
            printf("%s FACTORY hr=0x%08lX\n", tag, (unsigned long)hr);
            continue;
        }
        hr = IClassFactory_CreateInstance(f, NULL, &IID_IDispatch, (void **)&d);
        printf("%s hr=0x%08lX ptr=%s\n", tag, (unsigned long)hr, d ? "OK" : "NULL");
        IClassFactory_Release(f);
        if (d) {
            one[0] = mk_i4(21);          /* its own slot, set right before the call */
            call_member(d, L"Twice", one, 1, DISPATCH_METHOD, mtag);
            IDispatch_Release(d);
        }
    }
}

int main(int argc, char **argv) {
    HMODULE mod;
    PFN_GETCLASSOBJ getClassObject;
    IClassFactory *fac = NULL;
    IDispatch *disp = NULL;
    IUnknown *unk = NULL;
    IUnknown *extra = NULL;
    CLSID clsid;
    IID extraIid;
    VARIANT args[2];
    UINT cinfo = 0;
    ITypeInfo *tinfo = NULL;
    int ai;
    HRESULT hr;

    if (argc < 3) {
        printf("USAGE=<dll> <clsid> [iid]\n");
        return 2;
    }
    /* 探针的价值全在这些读数上：一旦中途崩溃，块缓冲会把已打印的行全吞掉
     * (管道/重定向下 stdout 不是行缓冲)，所以这里显式关掉缓冲。 */
    setvbuf(stdout, NULL, _IONBF, 0);
    printf("COINIT hr=0x%08lX\n", (unsigned long)CoInitializeEx(NULL, COINIT_APARTMENTTHREADED));

    mod = LoadLibraryExA(argv[1], NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
    if (!mod) {
        printf("LOAD=FAIL gle=%lu\n", (unsigned long)GetLastError());
        return 1;
    }
    printf("LOAD=OK\n");
    getClassObject = (PFN_GETCLASSOBJ)GetProcAddress(mod, "DllGetClassObject");
    if (!getClassObject) {
        printf("GETPROC=FAIL\n");
        return 1;
    }
    if (!parse_guid(argv[2], &clsid)) {
        printf("CLSID=BAD\n");
        return 1;
    }
    print_guid("CLSID", &clsid);

    hr = getClassObject(&clsid, &IID_IClassFactory_local, (void **)&fac);
    printf("GETFACTORY hr=0x%08lX ptr=%s\n", (unsigned long)hr, fac ? "OK" : "NULL");
    if (FAILED(hr) || !fac) return 1;

    hr = IClassFactory_CreateInstance(fac, NULL, &IID_IDispatch, (void **)&disp);
    printf("CREATE hr=0x%08lX ptr=%s\n", (unsigned long)hr, disp ? "OK" : "NULL");
    IClassFactory_Release(fac);
    if (FAILED(hr) || !disp) return 1;

    hr = IDispatch_QueryInterface(disp, &IID_IUnknown, (void **)&unk);
    printf("QI_IUNKNOWN hr=0x%08lX same=%s\n", (unsigned long)hr,
           (SUCCEEDED(hr) && unk == (IUnknown *)disp) ? "yes" : "no");
    if (unk) {
        IUnknown_Release(unk);
        unk = NULL;
    }

    /* 第二枚 IID 的应答面：答不答应 + 答应的是不是同一份胖指针（胖/瘦之分的实测，不断言） */
    if (argc >= 4 && parse_guid(argv[3], &extraIid)) {
        print_guid("EXTRA_IID", &extraIid);
        hr = IDispatch_QueryInterface(disp, &extraIid, (void **)&extra);
        printf("QI_EXTRA hr=0x%08lX same=%s\n", (unsigned long)hr,
               FAILED(hr) ? "no" : (extra == (IUnknown *)disp ? "yes" : "no"));
        if (extra) IUnknown_Release(extra);

        /* ai/022 B16: 走 `IClassFactory::CreateInstance(<接口 IID>)` 这条规范早期绑定路.
         * 工厂内部 QI 完立刻 Release 了包装器 (vb6comserver_factory.c), 所以这条读数同时
         * 回答两件事: ① 交出来的是**薄指针** (不是 IDispatch 包装器), ② 它**还活着** ——
         * 底座引用交还 (instanceClaimRelease) 之后, 对象保到最后一个薄引用 Release 为止. */
        {
            IClassFactory *fac2 = NULL;
            void *thin = NULL;
            hr = getClassObject(&clsid, &IID_IClassFactory_local, (void **)&fac2);
            if (SUCCEEDED(hr) && fac2) {
                hr = IClassFactory_CreateInstance(fac2, NULL, &extraIid, (void **)&thin);
                printf("CREATE_IFACE hr=0x%08lX ptr=%s\n", (unsigned long)hr,
                       thin ? "OK" : "NULL");
                IClassFactory_Release(fac2);
                if (thin) {
                    call_vtable_slots(thin);
                    /* 最后一个薄引用: 交给 IUnknown::Release 收尾. 这里用的正是
                     * 客户端的规范姿势 —— 薄指针与前 3 槽构成 IUnknown 兼容前缀,
                     * 所以它可直接当 IUnknown* 用 (B16 判据 ③ 的另一面). */
                    IUnknown_Release((IUnknown *)thin);
                }
            } else {
                printf("CREATE_IFACE hr=0x%08lX ptr=NULL\n", (unsigned long)hr);
            }
        }
    }

    hr = IDispatch_GetTypeInfoCount(disp, &cinfo);
    printf("TYPEINFOCOUNT=%u hr=0x%08lX\n", cinfo, (unsigned long)hr);
    hr = IDispatch_GetTypeInfo(disp, 0, LOCALE_USER_DEFAULT, &tinfo);
    printf("GETTYPEINFO0 hr=0x%08lX ptr=%s\n", (unsigned long)hr, tinfo ? "OK" : "NULL");
    if (tinfo) {
        TYPEATTR *ta = NULL;
        if (SUCCEEDED(ITypeInfo_GetTypeAttr(tinfo, &ta)) && ta) {
            printf("TYPEINFO0_kind=%d memcount=%u\n", (int)ta->typekind, (unsigned)ta->cFuncs);
            ITypeInfo_ReleaseTypeAttr(tinfo, ta);
        }
        ITypeInfo_Release(tinfo);
    }
    tinfo = NULL;
    hr = IDispatch_GetTypeInfo(disp, 1, LOCALE_USER_DEFAULT, &tinfo);
    printf("GETTYPEINFO1 hr=0x%08lX\n", (unsigned long)hr);
    if (tinfo) ITypeInfo_Release(tinfo);

    args[0] = mk_i4(40);
    args[1] = mk_i4(2);              /* 逆序 = Add(2, 40) */
    call_member(disp, L"Add", args, 2, DISPATCH_METHOD, "ADD");

    args[0] = mk_i4(7);
    call_member(disp, L"SetValue", args, 1, DISPATCH_METHOD, "SETVALUE");
    call_member(disp, L"GetValue", args, 0, DISPATCH_METHOD, "GETVALUE");

    /* Acct #305 SB144: the Decimal round trip, both directions. */
    call_member(disp, L"Amt", args, 0, DISPATCH_METHOD, "AMT");
    call_member(disp, L"Quarter", args, 0, DISPATCH_METHOD, "QUARTER");
    {
        VARIANT dargs[1];
        dargs[0] = mk_dec_i4(5);
        call_member(disp, L"EchoVal", dargs, 1, DISPATCH_METHOD, "ECHO");
    }
    {
        /* Acct #305 SB159/SB162: array in and back out through the same Variant slot. */
        VARIANT aargs[1];
        static const LONG av[3] = { 10, 20, 30 };
        aargs[0] = mk_sa_i4(av, 3);
        call_member(disp, L"EchoVal", aargs, 1, DISPATCH_METHOD, "ECHOARR");
    }
    /* Acct #305 SB185 (knife 77): the IDispatch server used to have no error boundary at all, so
     * a runtime error raised inside a VB body killed the CLIENT process (measured on both
     * architectures: exit code 6 for an out-of-range Public Integer field store, 5 for a plain
     * Err.Raise). These arms need a class publishing ci/cl plus the two Sub bodies; on any other
     * fixture the names just answer NAMES=<tag> hr=0x80020006 and no earlier needle moves. */
    call_member(disp, L"SeedLegal", args, 0, DISPATCH_METHOD, "SEED");
    /* Every put sets its own argument immediately before the call: a shared args[] slot that is
     * read after the previous call cleared it is a probe bug, not a product reading. */
    call_member(disp, L"ci", args, 0, DISPATCH_PROPERTYGET, "GETSEED");
    call_member(disp, L"cl", args, 0, DISPATCH_PROPERTYGET, "GETSEEDLONG");
    args[0] = mk_i4(7);
    call_member(disp, L"ci", args, 1, DISPATCH_PROPERTYPUT, "PUTLEGAL");
    call_member(disp, L"ci", args, 0, DISPATCH_PROPERTYGET, "GETLEGAL");
    args[0] = mk_i4(40000);
    call_member(disp, L"ci", args, 1, DISPATCH_PROPERTYPUT, "PUTRANGE");
    call_member(disp, L"ci", args, 0, DISPATCH_PROPERTYGET, "GETRANGE");
    args[0] = mk_i4(40000);
    call_member(disp, L"cl", args, 1, DISPATCH_PROPERTYPUT, "PUTLONG");
    call_member(disp, L"cl", args, 0, DISPATCH_PROPERTYGET, "GETLONG");
    args[0] = mk_i4(30000);
    call_member(disp, L"cl", args, 1, DISPATCH_PROPERTYPUT, "PUTLONG2");
    call_member(disp, L"cl", args, 0, DISPATCH_PROPERTYGET, "GETLONG2");
    call_member(disp, L"BoomNoHandler", args, 0, DISPATCH_METHOD, "BOOM");
    call_member(disp, L"HandledBoom", args, 0, DISPATCH_METHOD, "HANDLED");

    {
        DISPID id = DISPID_UNKNOWN;
        LPOLESTR names[1];
        names[0] = (LPOLESTR)L"NotAMember";
        hr = IDispatch_GetIDsOfNames(disp, &IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
        printf("NAMES=bogus hr=0x%08lX dispid=%ld\n", (unsigned long)hr, (long)id);
    }

    /* 账 #481: 表饱和的运行时证人 (只在给了 CHURN=<n> 时发, 其余夹具输出不变). */
    for (ai = 3; ai < argc; ai++)
        if (strncmp(argv[ai], "CHURN=", 6) == 0) call_churn(disp, (LONG)strtol(argv[ai] + 6, NULL, 10));

    /* 账 #305 §B219 (第二百二十刀): 异档实参那一族读数 (只在给了 VTARG= 时发). 排在 CREATE= 之前 ——
     * 它们打的是**同一枚已建好的** disp, 而 CREATE= 那一臂可能打死进程, 读数要先在管子里。 */
    for (ai = 3; ai < argc; ai++)
        if (strncmp(argv[ai], "VTARG=", 6) == 0) call_vtarg(disp, argv[ai] + 6);

    /* Acct #305 SB222 (knife 230): array arguments (only when asked). ARRARG= is the BYREF
     * shape a VB6 client uses for `a() As Long`; ARRARGV= is the plain-VT_ARRAY control. */
    for (ai = 3; ai < argc; ai++) {
        if (strncmp(argv[ai], "ARRARG=", 7) == 0) call_arrarg(disp, argv[ai] + 7, 1);
        else if (strncmp(argv[ai], "ARRARGV=", 8) == 0) call_arrarg(disp, argv[ai] + 8, 0);
        else if (strncmp(argv[ai], "BYREFARG=", 9) == 0) call_byrefarg(disp, argv[ai] + 9, 1);
        else if (strncmp(argv[ai], "BYVALARG=", 9) == 0) call_byrefarg(disp, argv[ai] + 9, 0);
    }

    /* Acct #305 SB186 (knife 78): the factory-side arms, last, so the primary object's readings are
     * already on the pipe before anything here can kill the process. */
    for (ai = 3; ai < argc; ai++)
        if (strncmp(argv[ai], "CREATE=", 7) == 0) create_arms(getClassObject, argv[ai] + 7);

    IDispatch_Release(disp);
    CoUninitialize();
    printf("DONE\n");
    return 0;
}
