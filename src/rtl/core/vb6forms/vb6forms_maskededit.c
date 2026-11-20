// ============================================================
// C30-3: MaskEdBox（Microsoft Masked Edit Control 6.0，MSMASK32.OCX）原生复刻
// ------------------------------------------------------------
// 落点：Win32 的 **EDIT**（不做自绘）+ 一层 RTL 掩码状态机。
// 为什么不直接给 EDIT 塞 EM_SETLIMITTEXT 了事：掩码不是"长度上限"，而是**逐位类型**，
// 每一位接受什么字符、哪些位是固定符号、哪些位可空，都只有掩码自己知道。
//
// 与 VB6 的语义口径（ai/内置控件/MaskEdBox 控件（掩码编辑框）.md）：
//   Text        含固定符号；PromptInclude=True 时还含空位提示符
//   ClipText    只含用户实际输入的字符（不含固定符号、不含提示符）—— 存库用这个
//   Mask        掩码串，回读原样（未解析的那一串）
//   PromptChar  空位提示符，默认 '_'（VB6 是 String 型，这里同样收 BSTR）
//   PromptInclude 默认 True
//
// 掩码字符表（照文档 §3 逐条立）：
//   0 数字必填   9 数字可空   # 数字/空格/正负号可空
//   L 字母必填   ? 字母可空   A 字母或数字必填   a 字母或数字可空
//   & 任意字符必填  C 任意字符可空
//   < 其后强制小写   > 其后强制大写   \ 转义：下一个字符当固定符号
//   其余字符（含 . , : ; - /）都是固定符号
//
// 编辑模型：显示串**恒为一位掩码一位字符**（固定符或用户字符或提示符），所以
// "显示下标 == 掩码位下标"这条恒等式成立 —— 光标/选区换算不用再查表。
// 键盘输入全部在子类里吃掉，不交给 EDIT 自己插（EDIT 不懂掩码，交出去就会出现
// 固定符被退格删掉、字母被插进数字位这类串味）。
// ============================================================
#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6rtl_bstr.h"

#ifdef _WIN32

#define VB6_ME_MAX      32
#define VB6_ME_MAXMASK  128

typedef enum {
    ME_LIT = 0,      // 固定符号
    ME_DIGIT_REQ,    // 0
    ME_DIGIT_OPT,    // 9
    ME_SIGN_OPT,     // #
    ME_ALPHA_REQ,    // L
    ME_ALPHA_OPT,    // ?
    ME_ALNUM_REQ,    // A
    ME_ALNUM_OPT,    // a
    ME_ANY_REQ,      // &
    ME_ANY_OPT       // C
} MeKind;

typedef struct {
    MeKind  kind;
    wchar_t lit;     // kind==ME_LIT 时的固定字符
    int     upper;   // 受 > 影响
    int     lower;   // 受 < 影响
} MeCell;

typedef struct {
    HWND      hwnd;
    MeCell    cell[VB6_ME_MAXMASK];
    wchar_t   ch[VB6_ME_MAXMASK];   // 输入位上的用户字符；0 = 空位
    int       len;                  // 掩码位数（= 显示串长度）
    wchar_t   mask[VB6_ME_MAXMASK]; // 原样掩码（Get 回读）
    wchar_t   prompt;               // 空位提示符
    int       promptInclude;
    int       guard;                // 防 SetWindowText 递归进 EN_CHANGE
    int       subbed;
    WNDPROC   prevProc;
} Vb6MaskEd;

static Vb6MaskEd g_me[VB6_ME_MAX];

static Vb6MaskEd* meFind(void* hwnd) {
    HWND h = (HWND)hwnd;
    int i;
    if (!h) return NULL;
    for (i = 0; i < VB6_ME_MAX; i++)
        if (g_me[i].hwnd == h) return &g_me[i];
    return NULL;
}

static int meIsInput(const MeCell* p) { return p->kind != ME_LIT; }

// 一位能不能收这个字符；收则把（经大小写转换后的）字符写进 *out。
static int meAccepts(const MeCell* p, wchar_t in, wchar_t* out) {
    wchar_t c = in;
    int alpha, digit;
    if (p->upper && c >= L'a' && c <= L'z') c = (wchar_t)(c - 32);
    if (p->lower && c >= L'A' && c <= L'Z') c = (wchar_t)(c + 32);
    alpha = (c >= L'A' && c <= L'Z') || (c >= L'a' && c <= L'z');
    digit = (c >= L'0' && c <= L'9');
    switch (p->kind) {
    case ME_DIGIT_REQ:
    case ME_DIGIT_OPT:   if (!digit) return 0; break;
    case ME_SIGN_OPT:    if (!digit && c != L'+' && c != L'-' && c != L' ') return 0; break;
    case ME_ALPHA_REQ:
    case ME_ALPHA_OPT:   if (!alpha) return 0; break;
    case ME_ALNUM_REQ:
    case ME_ALNUM_OPT:   if (!alpha && !digit) return 0; break;
    case ME_ANY_REQ:
    case ME_ANY_OPT:     break;
    default:             return 0;
    }
    *out = c;
    return 1;
}

// 掩码解析：`<` / `>` 是**状态量**（影响其后所有输入位），`\` 只作用于紧随的一位。
static void meParse(Vb6MaskEd* t, const wchar_t* mask) {
    int i = 0, n = 0, up = 0, low = 0;
    if (!t) return;
    if (!mask) mask = L"";
    ZeroMemory(t->cell, sizeof(t->cell));
    ZeroMemory(t->ch, sizeof(t->ch));
    while (mask[i] && n < VB6_ME_MAXMASK) {
        wchar_t c = mask[i];
        MeKind k = ME_LIT;
        if (c == L'\\' && mask[i + 1]) {          // 转义
            t->cell[n].kind = ME_LIT;
            t->cell[n].lit = mask[i + 1];
            n++; i += 2; continue;
        }
        if (c == L'>') { up = 1; low = 0; i++; continue; }
        if (c == L'<') { up = 0; low = 1; i++; continue; }
        switch (c) {
        case L'0': k = ME_DIGIT_REQ; break;
        case L'9': k = ME_DIGIT_OPT; break;
        case L'#': k = ME_SIGN_OPT;  break;
        case L'L': k = ME_ALPHA_REQ; break;
        case L'?': k = ME_ALPHA_OPT; break;
        case L'A': k = ME_ALNUM_REQ; break;
        case L'a': k = ME_ALNUM_OPT; break;
        case L'&': k = ME_ANY_REQ;   break;
        case L'C': k = ME_ANY_OPT;   break;
        default:   k = ME_LIT;       break;   // . , : ; - / 与其余字符都是固定符号
        }
        t->cell[n].kind = k;
        t->cell[n].lit = (k == ME_LIT) ? c : 0;
        t->cell[n].upper = up;
        t->cell[n].lower = low;
        n++; i++;
    }
    t->len = n;
    wcsncpy(t->mask, mask, VB6_ME_MAXMASK - 1);
    t->mask[VB6_ME_MAXMASK - 1] = 0;
}

// 渲染：includePrompt=1 → 空位出提示符（显示串）；0 → 空位直接省掉（Text 在
// PromptInclude=False 时的口径）。
static int meRender(const Vb6MaskEd* t, int includePrompt, wchar_t* out, int cap) {
    int i, n = 0;
    if (!t || !out || cap <= 0) return 0;
    for (i = 0; i < t->len && n < cap - 1; i++) {
        const MeCell* p = &t->cell[i];
        if (p->kind == ME_LIT) { out[n++] = p->lit; continue; }
        if (t->ch[i]) { out[n++] = t->ch[i]; continue; }
        if (includePrompt) out[n++] = t->prompt;
    }
    out[n] = 0;
    return n;
}

// ClipText：只交用户字符（固定符与空位都不要）。
static int meClip(const Vb6MaskEd* t, wchar_t* out, int cap) {
    int i, n = 0;
    if (!t || !out || cap <= 0) return 0;
    for (i = 0; i < t->len && n < cap - 1; i++)
        if (meIsInput(&t->cell[i]) && t->ch[i]) out[n++] = t->ch[i];
    out[n] = 0;
    return n;
}

// 把一段"人写的字符串"灌进掩码。
// 规则：走到固定符位时，若字符串这一位正好等于该固定符就吞掉它（这样
// `Text = "(010) 1234-5678"` 这种带格式的回显也认）；输入位逐位校验，不合法的
// 字符**丢弃而不是报错**（VB6 的 SetText 对非法输入就是"这位空着"）。
static void meApply(Vb6MaskEd* t, const wchar_t* s) {
    int i = 0;
    if (!t) return;
    ZeroMemory(t->ch, sizeof(t->ch));
    if (!s) return;
    for (; *s; s++) {
        wchar_t c = *s, cv;
        int j = i, litAte = 0;
        // 先过固定符：字符串这一位若正好等于固定符，就当作格式符吞掉
        while (j < t->len && !meIsInput(&t->cell[j])) {
            if (c == t->cell[j].lit) { i = j + 1; litAte = 1; break; }
            j++;
        }
        if (litAte) continue;
        // 找**第一个能收**的输入位；一个都不收就丢弃这个字符，且 i 不推进
        // —— 否则一个坏字符（例如 "ab1234" 里的 'a'）会把整条掩码烧完。
        for (; j < t->len; j++) {
            if (!meIsInput(&t->cell[j])) continue;
            if (meAccepts(&t->cell[j], c, &cv)) { t->ch[j] = cv; i = j + 1; break; }
        }
    }
}

// 当前光标所在的掩码位（显示下标与掩码位下标恒等）。
static int meCaretCell(Vb6MaskEd* t) {
    DWORD sel = 0;
    if (!t) return 0;
    SendMessageW(t->hwnd, EM_GETSEL, (WPARAM)&sel, 0);
    return (int)LOWORD(sel);
}

static void meSetCaret(Vb6MaskEd* t, int cell) {
    if (!t) return;
    if (cell < 0) cell = 0;
    if (cell > t->len) cell = t->len;
    SendMessageW(t->hwnd, EM_SETSEL, (WPARAM)cell, (LPARAM)cell);
}

// 从 from 起（含）的第一个输入位；没有则 -1。
static int meNextInput(Vb6MaskEd* t, int from) {
    int i;
    for (i = from; i >= 0 && i < t->len; i++)
        if (meIsInput(&t->cell[i])) return i;
    return -1;
}

// 把缓冲区刷到 EDIT 上（guard 期间不重入）。
static void mePaint(Vb6MaskEd* t) {
    wchar_t buf[VB6_ME_MAXMASK + 8];
    if (!t || t->guard) return;
    meRender(t, 1, buf, VB6_ME_MAXMASK + 8);
    t->guard = 1;
    SetWindowTextW(t->hwnd, buf);
    t->guard = 0;
}

// 打字：从光标位起找输入位，能收就覆写（掩码控件是**覆写**语义，不是插入）。
static void meType(Vb6MaskEd* t, wchar_t c) {
    int i = meNextInput(t, meCaretCell(t));
    wchar_t cv;
    if (i < 0) return;
    if (!meAccepts(&t->cell[i], c, &cv)) return;   // 非法：吃掉，不落盘
    t->ch[i] = cv;
    mePaint(t);
    meSetCaret(t, i + 1);
}

// 退格：从光标左边往回找第一个**有字符**的输入位，清掉它并把光标停在那儿。
static void meBackspace(Vb6MaskEd* t) {
    int i = meCaretCell(t) - 1;
    if (i >= t->len) i = t->len - 1;
    for (; i >= 0; i--) {
        if (meIsInput(&t->cell[i]) && t->ch[i]) {
            t->ch[i] = 0;
            mePaint(t);
            meSetCaret(t, i);
            return;
        }
    }
}

// Delete：清光标所在位（含前面那一串固定符之后的第一个输入位）。
static void meDelete(Vb6MaskEd* t) {
    int i = meNextInput(t, meCaretCell(t));
    if (i < 0) return;
    if (!t->ch[i]) return;
    t->ch[i] = 0;
    mePaint(t);
    meSetCaret(t, i);
}

// 粘贴：取剪贴板文本，从光标位起逐字覆写。
static void mePaste(Vb6MaskEd* t) {
    HANDLE h;
    wchar_t* p;
    int i;
    wchar_t cv;
    if (!t || !OpenClipboard(t->hwnd)) return;
    h = GetClipboardData(CF_UNICODETEXT);
    p = h ? (wchar_t*)GlobalLock(h) : NULL;
    if (p) {
        i = meNextInput(t, meCaretCell(t));
        for (; *p && i >= 0; p++) {
            if (*p == L'\r' || *p == L'\n') continue;
            if (!meAccepts(&t->cell[i], *p, &cv)) continue;
            t->ch[i] = cv;
            i = meNextInput(t, i + 1);
        }
        GlobalUnlock(h);
        mePaint(t);
        meSetCaret(t, i < 0 ? t->len : i);
    }
    CloseClipboard();
}

// 子类：**自己链**下一层，不占 cgen 的 VB6_OrigProc 属性槽（同 TabStrip 那一刀的
// 理由，账 #185：vb6_InstallControlSubclass 见槽有值就 return，而 cgen 是 Init 之后
// 才装控件子类的）。
static LRESULT CALLBACK meProc(HWND h, UINT m, WPARAM wp, LPARAM lp) {
    Vb6MaskEd* t = meFind(h);
    WNDPROC prev = t ? t->prevProc : NULL;
    if (t && !t->guard) {
        switch (m) {
        case WM_CHAR: {
            wchar_t c = (wchar_t)wp;
            if (c >= 32) { meType(t, c); return 0; }
            if (c == 8)  { meBackspace(t);  return 0; }   // VK_BACK
            if (c == 127) { meDelete(t);    return 0; }   // Ctrl+Backspace 在老编辑控件里是 DEL
            break;
        }
        case WM_KEYDOWN:
            if (wp == VK_BACK)   { meBackspace(t); return 0; }
            if (wp == VK_DELETE) { meDelete(t);    return 0; }
            break;
        case WM_PASTE:
        case WM_MBUTTONUP:     // 中键粘贴
            mePaste(t);
            return 0;
        case WM_CUT:
        case WM_CLEAR:
            return 0;          // 掩码控件不许"整段清空"，逐位退格才算
        case WM_DESTROY:
            t->hwnd = NULL;
            t->subbed = 0;
            t->prevProc = NULL;
            break;
        }
    }
    if (prev) return CallWindowProcW(prev, h, m, wp, lp);
    return DefWindowProcW(h, m, wp, lp);
}

// ===================== 对外 API =====================

// 设计期一次性：掩码 + 提示符 + PromptInclude + 初始文本。
// promptBstr 只取首字符（VB6 的 PromptChar 是 String，但只有一位有意义）。
void vb6_MaskEd_Init(void* hwnd, void* maskBstr, void* promptBstr,
                     int32_t promptInclude, void* textBstr) {
    HWND h = (HWND)hwnd;
    Vb6MaskEd* t;
    int i;
    const wchar_t* pm = (const wchar_t*)promptBstr;
    if (!h) return;
    t = meFind(hwnd);
    if (!t) {
        for (i = 0; i < VB6_ME_MAX; i++) if (!g_me[i].hwnd) { t = &g_me[i]; break; }
        if (!t) return;
        ZeroMemory(t, sizeof(*t));
        t->hwnd = h;
        t->prompt = L'_';
        t->promptInclude = 1;
    }
    if (pm && pm[0]) t->prompt = pm[0];
    t->promptInclude = promptInclude ? 1 : 0;
    meParse(t, (const wchar_t*)maskBstr);
    if (textBstr) meApply(t, (const wchar_t*)textBstr);
    mePaint(t);
    meSetCaret(t, 0);
    if (!t->subbed) {
        WNDPROC old = (WNDPROC)SetWindowLongPtrW(h, GWLP_WNDPROC, (LONG_PTR)meProc);
        if (old && old != meProc) { t->prevProc = old; t->subbed = 1; }
    }
}

// Text：PromptInclude=True 时除固定符外**还带空位提示符**；False 时空位省掉。
void* vb6_MaskEd_GetText(void* hwnd) {
    Vb6MaskEd* t = meFind(hwnd);
    wchar_t buf[VB6_ME_MAXMASK + 8];
    if (!t) return (void*)vb6_BSTR_Empty();
    meRender(t, t->promptInclude, buf, VB6_ME_MAXMASK + 8);
    return (void*)vb6_BSTR_FromStr(buf);
}

void vb6_MaskEd_SetText(void* hwnd, void* bstr) {
    Vb6MaskEd* t = meFind(hwnd);
    if (!t) return;
    meApply(t, (const wchar_t*)bstr);
    mePaint(t);
    meSetCaret(t, 0);
}

void* vb6_MaskEd_GetClipText(void* hwnd) {
    Vb6MaskEd* t = meFind(hwnd);
    wchar_t buf[VB6_ME_MAXMASK + 8];
    if (!t) return (void*)vb6_BSTR_Empty();
    meClip(t, buf, VB6_ME_MAXMASK + 8);
    return (void*)vb6_BSTR_FromStr(buf);
}

void vb6_MaskEd_SetClipText(void* hwnd, void* bstr) {
    Vb6MaskEd* t = meFind(hwnd);
    if (!t) return;
    meApply(t, (const wchar_t*)bstr);
    mePaint(t);
    meSetCaret(t, 0);
}

void* vb6_MaskEd_GetMask(void* hwnd) {
    Vb6MaskEd* t = meFind(hwnd);
    if (!t) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(t->mask);
}

// 换掩码：VB6 保留仍合法的那些位。这里保守地**重灌**（用当前 ClipText 按新掩码
// 重新对位），语义与"换完掩码再设一次 Text"一致。
void vb6_MaskEd_SetMask(void* hwnd, void* bstr) {
    Vb6MaskEd* t = meFind(hwnd);
    wchar_t keep[VB6_ME_MAXMASK + 8];
    if (!t) return;
    meClip(t, keep, VB6_ME_MAXMASK + 8);
    meParse(t, (const wchar_t*)bstr);
    meApply(t, keep);
    mePaint(t);
    meSetCaret(t, 0);
}

void* vb6_MaskEd_GetPromptChar(void* hwnd) {
    Vb6MaskEd* t = meFind(hwnd);
    wchar_t buf[2];
    if (!t) return (void*)vb6_BSTR_Empty();
    buf[0] = t->prompt; buf[1] = 0;
    return (void*)vb6_BSTR_FromStr(buf);
}

void vb6_MaskEd_SetPromptChar(void* hwnd, void* bstr) {
    Vb6MaskEd* t = meFind(hwnd);
    const wchar_t* s = (const wchar_t*)bstr;
    if (!t) return;
    t->prompt = (s && s[0]) ? s[0] : L' ';
    mePaint(t);
}

int32_t vb6_MaskEd_GetPromptInclude(void* hwnd) {
    Vb6MaskEd* t = meFind(hwnd);
    return t ? t->promptInclude : 0;
}

void vb6_MaskEd_SetPromptInclude(void* hwnd, int32_t v) {
    Vb6MaskEd* t = meFind(hwnd);
    if (!t) return;
    t->promptInclude = v ? 1 : 0;
}

#endif // _WIN32
