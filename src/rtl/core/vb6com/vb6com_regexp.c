// vb6com_regexp.c - 内置正则引擎, 顶掉 CreateObject("VBScript.RegExp") 的注册表 COM
//
// 背景：VB6 工程里 `Set re = CreateObject("VBScript.RegExp")` 走的是 vbscript.dll
// （注册表 CLSID）。本工程是"把运行时静态链接进用户程序"的路线，这种依赖外部 COM
// 组件的做法在目标机上可能压根没注册（或注册了 32/64 位不匹配），所以这里给一个
// 自带实现，并在 vb6_CreateObject 里抢先拦截该 ProgID。
//
// 语义基准 = 本机 cscript 跑真 VBScript.RegExp 的输出（读数与用例见 tests/regexp_native）：
//   · Pattern 的语法错误**延迟到第一次 Test/Execute/Replace** 才抛（赋值时不报）
//       '['        -> 5019 "正则表达式中缺少 ']'"
//       'a{3,2}'   -> 5017 "正则表达式语法错误"
//       '(?<x>a)'  -> 5017（VB 不支持命名组）
//   · '.' 不吃 \n，但吃 \r
//   · '^'/'$' 在 Multiline 下按行（"\r\n" 也成立）
//   · 空匹配在 Global 下逐个前进："a*" on "bbb" => 4 个 0:0 1:0 2:0 3:0
//   · Replace 支持 $1..$9、$&、$$、$`、$'
//   · SubMatches 未参与的组取 ""（不报错），Count 恒为组数
//
// 支持：字面量、.、[..]（含 \d\D\w\W\s\S 与 \xHH/\uHHHH/\b 退格）、
//       * + ? {n} {n,} {n,m} 及其惰性形式、捕获组、(?:)、(?=)(?!)、
//       交替 |、^ $、\b \B、后向引用 \1..\9、IgnoreCase、Multiline、Global。
// 不支持（VBScript 本身也不支持或极少用）：命名组、后向断言 (?<=)、\G、条件组。
//
// 扩展路线（文档 ai/037 里写了）：本引擎以"VBScript 兼容子集"为准，超出部分
// （如 (?<=)、\p{...}、命名组）以后可接 PCRE2 / Python re 的语义做增量，接口不变。

#include "vb6com.h"
#include "vb6com_internal.h"
#include <wchar.h>
#include <wctype.h>
#include <stdlib.h>

extern HRESULT vb6_ComChangeType(VARIANT* dest, const VARIANT* src, VARTYPE want);

// ============================================================
// 错误号（VBScript 正则引擎的真实取值，见头注的 cscript读数）
// ============================================================
#define RE_ERR_SYNTAX     5017
#define RE_ERR_MISSING_RB 5019
#define RE_ERR_MISSING_RP 5020
#define RE_ERR_TYPE       13   // 类型不匹配
#define RE_ERR_NO_MEMBER  438

// ============================================================
// 编译产物：节点数组 + 子节点下标池 + 字符集池
// ============================================================
#define RE_MAX_GROUPS 64

/* 逐节点冻结的选项位：SET 位 = “本节点自己说了算”，否则用对象上的 IgnoreCase/Multiline。
 * 内联选项 (?i:...) / (?-i:...) 就是往这几位上写；不动 ReCtx。 */
#define RE_F_IC_SET   1
#define RE_F_IC       2
#define RE_F_ML_SET   4
#define RE_F_ML       8
#define RE_F_DOT_SET  16
#define RE_F_DOT      32

enum {
    K_CHAR = 1, K_ANY, K_CLASS, K_SEQ, K_ALT, K_REP, K_GROUP,
    K_BOL, K_EOL, K_WORDB, K_NWORDB, K_BACKREF, K_LOOK, K_EMPTY,
    /* C30-10 扩展档（默认开启，不设开关；语义基准 = PCRE2 / Python re，见 ai/038）： */
    K_ATOMIC,     /* (?>...) 原子组、以及 x*+ 这类占有量词（= 原子化的贪婪 REP） */
    K_LBEHIND,    /* (?<=...) / (?<!...)：定长后顾，a = 固定长度 */
    K_STRBOL,     /* \A 串首 */
    K_STREOL,     /* \z 串尾（绝对） */
    K_STREOLZ     /* \Z 串尾（允许末尾一个 \n） */
};

typedef struct { wchar_t lo, hi; } ReRange;
#define RE_SH_MAX  8
#define RE_RG_MAX  64

typedef struct {
    int      neg;                 /* [^...] */
    int      nr;
    ReRange  r[RE_RG_MAX];
    int      nsh;                 /* 正/负简写：+1 \d  -1 \D   +2 \w  -2 \W   +4 \s  -4 \S */
    int      sh[RE_SH_MAX];
} ReSet;

typedef struct {
    int kind;
    int a, b, c, d;               /* a: 通用(字符/子节点/组号/断言类型)  b: child/min
                                     c: max(-1=∞)  d: greedy/negative */
    int fl;                       /* 冻结的选项位（RE_F_*）：内联选项 (?i:...) 落在各原子上 */
    wchar_t ch;
    int setId;
    int kid;                      /* SEQ/ALT 的 kids 起始下标, 数量在 b */
} RNode;

typedef struct {
    RNode*  n;   int nn,  nc;
    int*    kids; int nk, kc;
    ReSet*  sets; int ns, sc;
    int     ngroups;   /* 捕获组个数（不含 0） */
    wchar_t* gname[RE_MAX_GROUPS + 1];   /* 命名组的名字（下标 = 组号），malloc 的副本 */
    int     root;
} ReProg;

static void progFree(ReProg* p) {
    if (!p) return;
    int g;
    for (g = 0; g <= RE_MAX_GROUPS; g++) free(p->gname[g]);
    free(p->n); free(p->kids); free(p->sets); free(p);
}

static int progNode(ReProg* p, int kind) {
    if (p->nn == p->nc) {
        p->nc = p->nc ? p->nc * 2 : 32;
        p->n = (RNode*)realloc(p->n, sizeof(RNode) * p->nc);
    }
    memset(&p->n[p->nn], 0, sizeof(RNode));
    p->n[p->nn].kind = kind;
    p->n[p->nn].b = -1;      /* max = ∞ 缺省 */
    p->n[p->nn].setId = -1;
    return p->nn++;
}

static int progKid(ReProg* p, int node, int kid) {
    if (p->nk == p->kc) {
        p->kc = p->kc ? p->kc * 2 : 32;
        p->kids = (int*)realloc(p->kids, sizeof(int) * p->kc);
    }
    p->kids[p->nk] = kid;
    (void)node;
    return p->nk++;
}

static int progSet(ReProg* p) {
    if (p->ns == p->sc) {
        p->sc = p->sc ? p->sc * 2 : 8;
        p->sets = (ReSet*)realloc(p->sets, sizeof(ReSet) * p->sc);
    }
    memset(&p->sets[p->ns], 0, sizeof(ReSet));
    return p->ns++;
}

// ============================================================
// 解析：递归下降
// ============================================================
typedef struct {
    const wchar_t* p;
    int i, len;
    ReProg* prog;
    int err;
    int ngroups;
    int fl;              /* 当前生效的内联选项位（RE_F_*），进组时保存、出组时还原 */
} ReParse;

static int pAlt(ReParse* ps);

static int isDigitW(wchar_t c) { return c >= L'0' && c <= L'9'; }
static int isWordW(wchar_t c) {
    return isDigitW(c) || (c >= L'a' && c <= L'z') || (c >= L'A' && c <= L'Z') || c == L'_';
}
static int isSpaceW(wchar_t c) {
    return c == L' ' || c == L'\t' || c == L'\r' || c == L'\n' || c == L'\f' || c == L'\v';
}

static int hexVal(wchar_t c) {
    if (c >= L'0' && c <= L'9') return c - L'0';
    if (c >= L'a' && c <= L'f') return c - L'a' + 10;
    if (c >= L'A' && c <= L'F') return c - L'A' + 10;
    return -1;
}

static void setAddRange(ReSet* s, wchar_t lo, wchar_t hi) {
    if (s->nr < RE_RG_MAX) { s->r[s->nr].lo = lo; s->r[s->nr].hi = hi; s->nr++; }
}

static void setAddSh(ReSet* s, int sh) {
    if (s->nsh < RE_SH_MAX) s->sh[s->nsh++] = sh;
}

// 命名组查号：只在**已定义**的组里找。前向引用（\k<x> 出现在 (?<x>) 之前）不支持，
// 直接 5017 —— PCRE 允许前向引用，但 Python re 不允许，选择两边都能对上的那一半。
static int reLookupName(ReParse* ps, const wchar_t* nm, int len) {
    int g, i;
    if (len <= 0) return 0;
    for (g = 1; g <= ps->ngroups; g++) {
        const wchar_t* t = ps->prog->gname[g];
        if (!t) continue;
        for (i = 0; i < len && t[i] && t[i] == nm[i]; i++) { }
        if (i == len && t[i] == 0) return g;
    }
    return 0;
}

// 后顾只支持**定长**：算出节点的固定宽度（-1 = 变长，编译期报 5017）。
// 零宽断言贡献 0，所以 (?<=(?=a)b) 这类也能算。
static int reFixLen(const ReProg* pg, int node) {
    const RNode* n;
    if (node < 0) return -1;
    n = &pg->n[node];
    switch (n->kind) {
        case K_CHAR: case K_ANY: case K_CLASS: return 1;
        case K_EMPTY: case K_BOL: case K_EOL: case K_WORDB: case K_NWORDB:
        case K_LOOK: case K_LBEHIND: case K_STRBOL: case K_STREOL: case K_STREOLZ:
            return 0;
        case K_BACKREF: return -1;
        case K_GROUP: case K_ATOMIC: return reFixLen(pg, n->b);
        case K_REP: {
            int l;
            if (n->b != n->c) return -1;
            l = reFixLen(pg, n->a);
            if (l < 0) return -1;
            return l * n->b;
        }
        case K_SEQ: {
            int t = 0, i;
            for (i = 0; i < n->b; i++) {
                int l = reFixLen(pg, pg->kids[n->kid + i]);
                if (l < 0) return -1;
                t += l;
            }
            return t;
        }
        case K_ALT: {
            int first = -1, i;
            for (i = 0; i < n->b; i++) {
                int l = reFixLen(pg, pg->kids[n->kid + i]);
                if (l < 0) return -1;
                if (first < 0) first = l;
                else if (l != first) return -1;
            }
            return first < 0 ? 0 : first;
        }
    }
    return -1;
}

// 解析一个转义：返回 0 = 生成节点(下标在 *outNode)，1 = 往字符集里加(仅在类内)
// inClass = 1 时把结果写进 s 并返回 2（无节点）
static int parseEscape(ReParse* ps, int inClass, ReSet* s, int* outNode, wchar_t* litOut) {
    const wchar_t* p = ps->p;
    if (ps->i >= ps->len) { ps->err = RE_ERR_SYNTAX; return -1; }
    wchar_t c = p[ps->i++];
    wchar_t lit = 0;
    int sh = 0;
    switch (c) {
        case L'd': sh = 1;  break;
        case L'D': sh = -1; break;
        case L'w': sh = 2;  break;
        case L'W': sh = -2; break;
        case L's': sh = 4;  break;
        case L'S': sh = -4; break;
        case L'n': lit = L'\n'; break;
        case L'r': lit = L'\r'; break;
        case L't': lit = L'\t'; break;
        case L'f': lit = L'\f'; break;
        case L'v': lit = L'\v'; break;
        case L'0': lit = L'\0'; break;
        case L'b':
            if (inClass) { lit = L'\b'; break; }      /* 类内 \b = 退格 */
            if (outNode) *outNode = progNode(ps->prog, K_WORDB);
            return 0;
        case L'B':
            if (inClass) { lit = L'B'; break; }
            if (outNode) *outNode = progNode(ps->prog, K_NWORDB);
            return 0;
        /* ---- 扩展档（默认开启）：\A \z \Z \k<name> ---- */
        case L'A':
            if (inClass) { lit = L'A'; break; }
            if (outNode) *outNode = progNode(ps->prog, K_STRBOL);
            return 0;
        case L'z':
            if (inClass) { lit = L'z'; break; }
            if (outNode) *outNode = progNode(ps->prog, K_STREOL);
            return 0;
        case L'Z':
            if (inClass) { lit = L'Z'; break; }
            if (outNode) *outNode = progNode(ps->prog, K_STREOLZ);
            return 0;
        case L'k': {
            wchar_t open, close; int st, g2;
            if (inClass) { lit = L'k'; break; }
            open = (ps->i < ps->len) ? p[ps->i] : 0;
            if (open == L'<') close = L'>';
            else if (open == L'\'') close = L'\'';
            else if (open == L'{') close = L'}';
            else { lit = L'k'; break; }        /* 不成形就当字面量 k */
            ps->i++;
            st = ps->i;
            while (ps->i < ps->len && p[ps->i] != close) ps->i++;
            if (ps->i >= ps->len) { ps->err = RE_ERR_SYNTAX; return -1; }
            g2 = reLookupName(ps, p + st, ps->i - st);
            ps->i++;
            if (g2 <= 0) { ps->err = RE_ERR_SYNTAX; return -1; }
            if (outNode) {
                int idx = progNode(ps->prog, K_BACKREF);
                ps->prog->n[idx].a = g2;
                ps->prog->n[idx].fl = ps->fl;
                *outNode = idx;
            }
            return 0;
        }
        case L'x': {
            int h1 = (ps->i < ps->len) ? hexVal(p[ps->i]) : -1;
            int h2 = (ps->i + 1 < ps->len) ? hexVal(p[ps->i + 1]) : -1;
            if (h1 < 0 || h2 < 0) { ps->err = RE_ERR_SYNTAX; return -1; }
            ps->i += 2;
            lit = (wchar_t)((h1 << 4) | h2);
            break;
        }
        case L'u': {
            int v = 0, k;
            for (k = 0; k < 4; k++) {
                int h = (ps->i < ps->len) ? hexVal(p[ps->i]) : -1;
                if (h < 0) { ps->err = RE_ERR_SYNTAX; return -1; }
                v = (v << 4) | h;
                ps->i++;
            }
            lit = (wchar_t)v;
            break;
        }
        default:
            if (c >= L'1' && c <= L'9') {
                if (inClass) { lit = c; break; }
                if (outNode) {
                    int idx = progNode(ps->prog, K_BACKREF);
                    ps->prog->n[idx].a = (int)(c - L'0');
                    ps->prog->n[idx].fl = ps->fl;
                    *outNode = idx;
                }
                return 0;
            }
            lit = c;   /* VBScript: 未知转义 = 该字符本身（\A 实测不报错） */
            break;
    }
    if (sh) {
        if (inClass) { setAddSh(s, sh); return 0; }
        int idx = progSet(ps->prog);
        setAddSh(&ps->prog->sets[idx], sh);
        int nd = progNode(ps->prog, K_CLASS);
        ps->prog->n[nd].setId = idx;
        ps->prog->n[nd].fl = ps->fl;
        if (outNode) *outNode = nd;
        return 0;
    }
    if (inClass) { if (litOut) *litOut = lit; return 1; }
    int nd = progNode(ps->prog, K_CHAR);
    ps->prog->n[nd].ch = lit;
    ps->prog->n[nd].fl = ps->fl;
    if (outNode) *outNode = nd;
    return 0;
}

// [ ... ]
static int parseClass(ReParse* ps) {
    ReProg* pg = ps->prog;
    int sid = progSet(pg);
    ReSet* s = &pg->sets[sid];
    if (ps->i < ps->len && ps->p[ps->i] == L'^') { s->neg = 1; ps->i++; }
    int first = 1;
    for (;;) {
        if (ps->i >= ps->len) { ps->err = RE_ERR_MISSING_RB; return -1; }
        wchar_t c = ps->p[ps->i];
        if (c == L']' && !first) { ps->i++; break; }
        first = 0;
        wchar_t lo = 0;
        int loSet = 0;
        if (c == L'\\') {
            ps->i++;
            wchar_t lit = 0;
            int r = parseEscape(ps, 1, s, NULL, &lit);
            if (r < 0) return -1;
            if (r == 0) continue;         /* 简写类已并入 */
            lo = lit; loSet = 1;
        } else {
            lo = c; ps->i++; loSet = 1;
        }
        if (!loSet) continue;
        /* 范围 a-z? */
        if (ps->i + 1 < ps->len && ps->p[ps->i] == L'-' && ps->p[ps->i + 1] != L']') {
            ps->i++;
            wchar_t hi;
            if (ps->p[ps->i] == L'\\') {
                ReSet tmp; memset(&tmp, 0, sizeof(tmp));
                ps->i++;
                wchar_t lit2 = 0;
                int r = parseEscape(ps, 1, &tmp, NULL, &lit2);
                if (r < 0) return -1;
                if (r == 0) { setAddRange(s, lo, lo); setAddRange(s, L'-', L'-'); continue; }
                hi = lit2;
            } else {
                hi = ps->p[ps->i]; ps->i++;
            }
            if (hi < lo) { ps->err = RE_ERR_SYNTAX; return -1; }
            setAddRange(s, lo, hi);
        } else {
            setAddRange(s, lo, lo);
        }
    }
    int nd = progNode(pg, K_CLASS);
    pg->n[nd].setId = sid;
    pg->n[nd].fl = ps->fl;
    return nd;
}

// {n} {n,} {n,m}
static int parseBrace(ReParse* ps, int* minOut, int* maxOut) {
    int min = 0, max = -1;
    int any = 0;
    while (ps->i < ps->len && isDigitW(ps->p[ps->i])) { min = min * 10 + (ps->p[ps->i] - L'0'); ps->i++; any = 1; }
    if (!any) return 0;
    if (ps->i < ps->len && ps->p[ps->i] == L',') {
        ps->i++;
        if (ps->i < ps->len && isDigitW(ps->p[ps->i])) {
            max = 0;
            while (ps->i < ps->len && isDigitW(ps->p[ps->i])) { max = max * 10 + (ps->p[ps->i] - L'0'); ps->i++; }
        }
    } else {
        max = min;
    }
    if (ps->i >= ps->len || ps->p[ps->i] != L'}') return 0;   /* 不是量词, 当字面量 */
    ps->i++;
    if (max >= 0 && min > max) { ps->err = RE_ERR_SYNTAX; return -1; }
    *minOut = min; *maxOut = max;
    return 1;
}

static int parseAtom(ReParse* ps) {
    ReProg* pg = ps->prog;
    if (ps->i >= ps->len) return progNode(pg, K_EMPTY);
    wchar_t c = ps->p[ps->i];
    switch (c) {
        case L'(' : {
            ps->i++;
            int saveFl = ps->fl;
            int look = 0, neg = 0, noncap = 0, atomic = 0, behind = 0;
            const wchar_t* named = NULL; int namedLen = 0;
            if (ps->i < ps->len && ps->p[ps->i] == L'?') {
                wchar_t k1 = (ps->i + 1 < ps->len) ? ps->p[ps->i + 1] : 0;
                if (k1 == L'=') { look = 1; ps->i += 2; }
                else if (k1 == L'!') { look = 1; neg = 1; ps->i += 2; }
                else if (k1 == L':') { noncap = 1; ps->i += 2; }
                else if (k1 == L'>') { atomic = 1; ps->i += 2; }                  /* (?>...) */
                else if (k1 == L'<') {
                    wchar_t k2 = (ps->i + 2 < ps->len) ? ps->p[ps->i + 2] : 0;
                    if (k2 == L'=') { behind = 1; look = 1; ps->i += 3; }         /* (?<= */
                    else if (k2 == L'!') { behind = 1; look = 1; neg = 1; ps->i += 3; } /* (?<! */
                    else {                                                        /* (?<name> */
                        ps->i += 2; named = ps->p + ps->i;
                        while (ps->i < ps->len && ps->p[ps->i] != L'>') ps->i++;
                        if (ps->i >= ps->len) { ps->err = RE_ERR_MISSING_RP; return -1; }
                        namedLen = (int)((ps->p + ps->i) - named);
                        ps->i++;
                    }
                }
                else if (k1 == L'\'') {                                          /* (?'name' */
                    ps->i += 2; named = ps->p + ps->i;
                    while (ps->i < ps->len && ps->p[ps->i] != L'\'') ps->i++;
                    if (ps->i >= ps->len) { ps->err = RE_ERR_MISSING_RP; return -1; }
                    namedLen = (int)((ps->p + ps->i) - named);
                    ps->i++;
                }
                else if (k1 == L'P') {
                    wchar_t k2 = (ps->i + 2 < ps->len) ? ps->p[ps->i + 2] : 0;
                    if (k2 == L'<') {                                             /* (?P<name> */
                        ps->i += 3; named = ps->p + ps->i;
                        while (ps->i < ps->len && ps->p[ps->i] != L'>') ps->i++;
                        if (ps->i >= ps->len) { ps->err = RE_ERR_MISSING_RP; return -1; }
                        namedLen = (int)((ps->p + ps->i) - named);
                        ps->i++;
                    } else if (k2 == L'=') {                                      /* (?P=name) */
                        const wchar_t* nm; int nlen, g2, nd2;
                        ps->i += 3; nm = ps->p + ps->i;
                        while (ps->i < ps->len && ps->p[ps->i] != L')') ps->i++;
                        if (ps->i >= ps->len) { ps->err = RE_ERR_MISSING_RP; return -1; }
                        nlen = (int)((ps->p + ps->i) - nm);
                        ps->i++;
                        g2 = reLookupName(ps, nm, nlen);
                        if (g2 <= 0) { ps->err = RE_ERR_SYNTAX; return -1; }
                        nd2 = progNode(pg, K_BACKREF);
                        pg->n[nd2].a = g2;
                        return nd2;
                    } else { ps->err = RE_ERR_SYNTAX; return -1; }
                }
                /* ---- 内联选项 (?ims:...) / (?-ims:...) / (?ims) ----
                 * 冻结到每个原子上（RE_F_* 位），不动 ReCtx —— 所以 (?i:a)b 里 b 仍按对象设置。 */
                else {
                    int j = ps->i + 1, nfl = ps->fl, negf = 0, anyf = 0;
                    while (j < ps->len) {
                        wchar_t f = ps->p[j];
                        int setb, valb;
                        if (f == L'-') { negf = 1; j++; continue; }
                        if (f != L'i' && f != L'm' && f != L's') break;
                        setb = (f == L'i') ? RE_F_IC_SET : (f == L'm') ? RE_F_ML_SET : RE_F_DOT_SET;
                        valb = (f == L'i') ? RE_F_IC     : (f == L'm') ? RE_F_ML     : RE_F_DOT;
                        if (negf) nfl = (nfl & ~valb) | setb;
                        else      nfl = (nfl | valb) | setb;
                        anyf = 1; j++;
                    }
                    if (!anyf) { ps->err = RE_ERR_SYNTAX; return -1; }
                    if (j < ps->len && ps->p[j] == L':') {
                        int child2, nd2;
                        ps->i = j + 1;
                        ps->fl = nfl;
                        child2 = pAlt(ps);
                        if (child2 < 0) return -1;
                        if (ps->i >= ps->len || ps->p[ps->i] != L')') { ps->err = RE_ERR_MISSING_RP; return -1; }
                        ps->i++;
                        ps->fl = saveFl;
                        nd2 = progNode(pg, K_GROUP);
                        pg->n[nd2].a = 0;
                        pg->n[nd2].b = child2;
                        return nd2;
                    }
                    if (j < ps->len && ps->p[j] == L')') {
                        ps->i = j + 1;
                        ps->fl = nfl;          /* (?i) 作用于本组剩余部分（PCRE 也管后续分支），出组还原 */
                        return progNode(pg, K_EMPTY);
                    }
                    ps->err = RE_ERR_SYNTAX; return -1;
                }
            }
            /* VB 的 SubMatches 按**左括号出现顺序**编号 —— 外层组号必须先定，再解子模式。 */
            int g = 0;
            if (!look && !noncap && !atomic) {
                g = ++ps->ngroups;
                if (g > RE_MAX_GROUPS) { ps->err = RE_ERR_SYNTAX; return -1; }
                if (namedLen > 0) {
                    wchar_t* cp = (wchar_t*)malloc(sizeof(wchar_t) * (namedLen + 1));
                    if (cp) { memcpy(cp, named, sizeof(wchar_t) * namedLen); cp[namedLen] = 0; pg->gname[g] = cp; }
                }
            }
            int child = pAlt(ps);
            if (child < 0) return -1;
            if (ps->i >= ps->len || ps->p[ps->i] != L')') { ps->err = RE_ERR_MISSING_RP; return -1; }
            ps->i++;
            ps->fl = saveFl;
            if (look) {
                int nd;
                if (behind) {
                    int flen = reFixLen(pg, child);
                    /* 变长后顾：PCRE2 限量允许，Python 直接报错；这里报 5017（见 ai/038 限制） */
                    if (flen < 0) { ps->err = RE_ERR_SYNTAX; return -1; }
                    nd = progNode(pg, K_LBEHIND);
                    pg->n[nd].a = flen;
                    pg->n[nd].b = child;
                    pg->n[nd].d = neg;
                    return nd;
                }
                nd = progNode(pg, K_LOOK);
                pg->n[nd].b = child;
                pg->n[nd].d = neg;
                return nd;
            }
            if (atomic) {
                int nd = progNode(pg, K_ATOMIC);
                pg->n[nd].b = child;
                return nd;
            }
            {
                int nd = progNode(pg, K_GROUP);
                pg->n[nd].a = noncap ? 0 : g;
                pg->n[nd].b = child;
                return nd;
            }
        }
        case L'[' : ps->i++; return parseClass(ps);
        case L'.' : ps->i++; { int nd = progNode(pg, K_ANY); pg->n[nd].fl = ps->fl; return nd; }
        case L'^' : ps->i++; { int nd = progNode(pg, K_BOL); pg->n[nd].fl = ps->fl; return nd; }
        case L'$' : ps->i++; { int nd = progNode(pg, K_EOL); pg->n[nd].fl = ps->fl; return nd; }
        case L'\\': ps->i++; {
            int nd = -1;
            /* \Q...\E：里面全是字面量，直到 \E 或模式结尾 */
            if (ps->i < ps->len && ps->p[ps->i] == L'Q') {
                int seq, start, cnt = 0;
                ps->i++;
                seq = progNode(pg, K_SEQ);
                start = pg->nk;
                while (ps->i < ps->len) {
                    if (ps->p[ps->i] == L'\\' && ps->i + 1 < ps->len && ps->p[ps->i + 1] == L'E') { ps->i += 2; break; }
                    {
                        int ch = progNode(pg, K_CHAR);
                        pg->n[ch].ch = ps->p[ps->i];
                        pg->n[ch].fl = ps->fl;
                        ps->i++;
                        progKid(pg, seq, ch);
                        cnt++;
                    }
                }
                pg->n[seq].kid = start;
                pg->n[seq].b = cnt;
                return seq;
            }
            if (parseEscape(ps, 0, NULL, &nd, NULL) < 0) return -1;
            return nd;
        }
        case L')' : case L'|' : return progNode(pg, K_EMPTY);   /* 交给上层的空序列 */
        default: {
            ps->i++;
            int nd = progNode(pg, K_CHAR);
            pg->n[nd].ch = c;
            pg->n[nd].fl = ps->fl;
            return nd;
        }
    }
}

static int parseSeq(ReParse* ps) {
    ReProg* pg = ps->prog;
    int seq = progNode(pg, K_SEQ);
    int* tmp = NULL; int nt = 0, ct = 0;
    while (ps->i < ps->len && ps->p[ps->i] != L'|' && ps->p[ps->i] != L')') {
        int atom = parseAtom(ps);
        if (atom < 0) return -1;
        /* 量词 */
        int min = -1, max = -1, greedy = 1, poss = 0;
        if (ps->i < ps->len) {
            wchar_t q = ps->p[ps->i];
            if (q == L'*') { min = 0; max = -1; ps->i++; }
            else if (q == L'+') { min = 1; max = -1; ps->i++; }
            else if (q == L'?') { min = 0; max = 1; ps->i++; }
            else if (q == L'{') {
                int save = ps->i;
                ps->i++;
                int r = parseBrace(ps, &min, &max);
                if (r < 0) return -1;
                if (r == 0) { ps->i = save; min = -1; }   /* '{' 当字面量 */
            }
            if (min >= 0 && ps->i < ps->len) {
                if (ps->p[ps->i] == L'?') { greedy = 0; ps->i++; }
                else if (ps->p[ps->i] == L'+') { poss = 1; ps->i++; }   /* x*+ 占有量词 */
            }
        }
        if (min >= 0) {
            int rep = progNode(pg, K_REP);
            pg->n[rep].a = atom;
            pg->n[rep].b = min;
            pg->n[rep].c = max;
            pg->n[rep].d = greedy;
            atom = rep;
            /* 占有量词 = 原子化的贪婪重复（PCRE 的定义就是 ?>X*） */
            if (poss) {
                int at = progNode(pg, K_ATOMIC);
                pg->n[at].b = rep;
                atom = at;
            }
        }
        if (nt == ct) { ct = ct ? ct * 2 : 8; tmp = (int*)realloc(tmp, sizeof(int) * ct); }
        tmp[nt++] = atom;
    }
    /* 子节点必须**先收进局部数组再一次性入池**：嵌套的组/交替也会往同一个 kids 池追加，
     * 边解边入池会让外层 SEQ 的 [kid, kid+b) 区间卷进内层的子节点（那正是 18 行不一致的根）。 */
    {
        int start = pg->nk, i;
        for (i = 0; i < nt; i++) progKid(pg, seq, tmp[i]);
        pg->n[seq].kid = start;
        pg->n[seq].b = nt;
    }
    free(tmp);
    return seq;
}

static int pAlt(ReParse* ps) {
    ReProg* pg = ps->prog;
    int first = parseSeq(ps);
    if (first < 0) return -1;
    if (ps->i >= ps->len || ps->p[ps->i] != L'|') return first;
    int alt = progNode(pg, K_ALT);
    int* tmp = NULL; int nt = 0, ct = 0;
    tmp = (int*)realloc(tmp, sizeof(int) * 8); ct = 8;
    tmp[nt++] = first;
    while (ps->i < ps->len && ps->p[ps->i] == L'|') {
        ps->i++;
        int nx = parseSeq(ps);
        if (nx < 0) { free(tmp); return -1; }
        if (nt == ct) { ct *= 2; tmp = (int*)realloc(tmp, sizeof(int) * ct); }
        tmp[nt++] = nx;
    }
    {
        int start = pg->nk, i;
        for (i = 0; i < nt; i++) progKid(pg, alt, tmp[i]);
        pg->n[alt].kid = start;
        pg->n[alt].b = nt;
    }
    free(tmp);
    return alt;
}

// 语法错误时带上 VBScript 的原文
static void raisePatternError(int err, const wchar_t* pat) {
    (void)pat;
    if (err == RE_ERR_MISSING_RB)
        vb6_RaiseError(RE_ERR_MISSING_RB, (void*)SysAllocString(L"\u6b63\u5219\u8868\u8fbe\u5f0f\u4e2d\u7f3a\u5c11 ']'"));
    else
        vb6_RaiseError(RE_ERR_SYNTAX, (void*)SysAllocString(L"\u6b63\u5219\u8868\u8fbe\u5f0f\u8bed\u6cd5\u9519\u8bef"));
}

static ReProg* compilePattern(const wchar_t* pat, int* errOut) {
    ReProg* pg = (ReProg*)calloc(1, sizeof(ReProg));
    if (!pg) { *errOut = RE_ERR_SYNTAX; return NULL; }
    ReParse ps;
    ps.p = pat ? pat : L"";
    ps.i = 0;
    ps.len = (int)wcslen(ps.p);
    ps.prog = pg;
    ps.err = 0;
    ps.ngroups = 0;
    ps.fl = 0;          /* 必须显式置 0：栈上脏值会让 . / ^ / $ 随机按 DotAll 或 Multiline 跑 */
    int root = pAlt(&ps);
    if (root < 0 || ps.err) { *errOut = ps.err ? ps.err : RE_ERR_SYNTAX; progFree(pg); return NULL; }
    if (ps.i != ps.len) { *errOut = RE_ERR_SYNTAX; progFree(pg); return NULL; }  /* 多余的 ')' */
    pg->root = root;
    pg->ngroups = ps.ngroups;
    *errOut = 0;
    return pg;
}

// ============================================================
// 匹配：回溯 + 续延
// ============================================================
typedef struct Cont {
    int kind;      /* 0 = seq 剩余, 1 = 组结束, 2 = 重复一次后的续延 */
    int a, b, c;
    struct Cont* next;
} Cont;

typedef struct {
    const ReProg* pg;
    const wchar_t* s;
    int len;
    int icase, multiline;
    int caps[2 * (RE_MAX_GROUPS + 1)];
    int steps;
} ReCtx;

#define RE_STEP_LIMIT 2000000

static int mRun(ReCtx* x, int node, int pos, Cont* k);
static int mCont(ReCtx* x, Cont* k, int pos);

static int chEq(ReCtx* x, wchar_t a, wchar_t b, int ic) {
    if (a == b) return 1;
    if (!ic) return 0;
    if (towlower((wint_t)a) == towlower((wint_t)b)) return 1;
    if (towupper((wint_t)a) == towupper((wint_t)b)) return 1;
    return 0;
}

/* 逐节点生效的选项：节点自己冻结了就用它的，否则跟对象（IgnoreCase/Multiline）走。
 * 这就是 (?i:...) 不污染外层的全部机关 —— ReCtx 一个字节都没动。 */
static int nodeIC(const ReCtx* x, const RNode* n) {
    return (n->fl & RE_F_IC_SET) ? ((n->fl & RE_F_IC) != 0) : x->icase;
}
static int nodeML(const ReCtx* x, const RNode* n) {
    return (n->fl & RE_F_ML_SET) ? ((n->fl & RE_F_ML) != 0) : x->multiline;
}
static int nodeDOT(const ReCtx* x, const RNode* n) {
    (void)x;
    return (n->fl & RE_F_DOT_SET) ? ((n->fl & RE_F_DOT) != 0) : 0;
}

static int setHasSh(const ReSet* s, int want) {
    int k;
    for (k = 0; k < s->nsh; k++) if (s->sh[k] == want) return 1;
    return 0;
}

static int shMatch(int sh, wchar_t c) {
    switch (sh) {
        case 1:  return isDigitW(c);
        case 2:  return isWordW(c);
        case 4:  return isSpaceW(c);
        default: return 0;
    }
}

static int classRaw(const ReSet* s, wchar_t c) {
    int k, hit = 0;
    for (k = 0; k < s->nr; k++)
        if (c >= s->r[k].lo && c <= s->r[k].hi) { hit = 1; break; }
    if (!hit) {
        for (k = 0; k < s->nsh; k++) {
            int sh = s->sh[k];
            if (sh > 0 && shMatch(sh, c)) { hit = 1; break; }
        }
    }
    if (hit) {
        for (k = 0; k < s->nsh; k++) {
            int sh = s->sh[k];
            if (sh < 0 && shMatch(-sh, c)) return 0;      /* \D \W \S 在类内是减法 */
        }
    }
    return hit;
}

static int classMatch(ReCtx* x, const ReSet* s, wchar_t c, int ic) {
    int hit = classRaw(s, c);
    (void)x;
    if (!hit && ic) {
        wchar_t u = (wchar_t)towupper((wint_t)c);
        wchar_t l = (wchar_t)towlower((wint_t)c);
        hit = classRaw(s, u) || classRaw(s, l);
    }
    return s->neg ? !hit : hit;
}

static int isWordAt(ReCtx* x, int pos) {
    if (pos < 0 || pos >= x->len) return 0;
    return isWordW(x->s[pos]);
}

static int repTry(ReCtx* x, int node, int count, int pos, Cont* k) {
    const RNode* n = &x->pg->n[node];
    int min = n->b, max = n->c, greedy = n->d;
    int canMore = (max < 0 || count < max);
    if (greedy) {
        if (canMore) {
            Cont c; c.kind = 2; c.a = node; c.b = count + 1; c.c = pos; c.next = k;
            if (mRun(x, n->a, pos, &c)) return 1;
        }
        if (count >= min) return mCont(x, k, pos);
        return 0;
    }
    if (count >= min) { if (mCont(x, k, pos)) return 1; }
    if (canMore) {
        Cont c; c.kind = 2; c.a = node; c.b = count + 1; c.c = pos; c.next = k;
        if (mRun(x, n->a, pos, &c)) return 1;
    }
    return 0;
}

static int mCont(ReCtx* x, Cont* k, int pos) {
    if (!k) return 1;
    switch (k->kind) {
        case 0: {
            const RNode* sq = &x->pg->n[k->a];
            int idx = k->b;
            if (idx >= sq->b) return mCont(x, k->next, pos);
            int node = x->pg->kids[sq->kid + idx];
            Cont c; c.kind = 0; c.a = k->a; c.b = idx + 1; c.c = 0; c.next = k->next;
            return mRun(x, node, pos, &c);
        }
        case 1: {
            int g = k->a;
            int old = x->caps[2 * g + 1];
            x->caps[2 * g + 1] = pos;
            if (mCont(x, k->next, pos)) return 1;
            x->caps[2 * g + 1] = old;
            return 0;
        }
        case 2: {
            int node = k->a, count = k->b, startPos = k->c;
            if (pos == startPos && count >= x->pg->n[node].b)
                return mCont(x, k->next, pos);      /* 空迭代: 收手, 别死循环 */
            return repTry(x, node, count, pos, k->next);
        }
        case 3:
            k->c = pos;                              /* 原子组：只记第一次到这里的落点 */
            return 2;
        case 4:
            if (pos != k->b) return 0;               /* 后顾：必须在锚点上正好收尾 */
            return mCont(x, k->next, pos);
    }
    return 0;
}

static int mRun(ReCtx* x, int node, int pos, Cont* k) {
    const RNode* n = &x->pg->n[node];
    if (++x->steps > RE_STEP_LIMIT) return 0;
    switch (n->kind) {
        case K_EMPTY: return mCont(x, k, pos);
        case K_CHAR:
            if (pos < x->len && chEq(x, x->s[pos], n->ch, nodeIC(x, n))) return mCont(x, k, pos + 1);
            return 0;
        case K_ANY:
            if (pos < x->len && (nodeDOT(x, n) || x->s[pos] != L'\n')) return mCont(x, k, pos + 1);
            return 0;
        case K_CLASS:
            if (pos < x->len && classMatch(x, &x->pg->sets[n->setId], x->s[pos], nodeIC(x, n)))
                return mCont(x, k, pos + 1);
            return 0;
        case K_SEQ: {
            if (n->b == 0) return mCont(x, k, pos);
            Cont c; c.kind = 0; c.a = node; c.b = 1; c.c = 0; c.next = k;
            return mRun(x, x->pg->kids[n->kid], pos, &c);
        }
        case K_ALT: {
            int i;
            for (i = 0; i < n->b; i++)
                if (mRun(x, x->pg->kids[n->kid + i], pos, k)) return 1;
            return 0;
        }
        case K_REP: return repTry(x, node, 0, pos, k);
        case K_GROUP: {
            int g = n->a;
            if (g == 0) return mRun(x, n->b, pos, k);          /* (?:...) */
            int os = x->caps[2 * g], oe = x->caps[2 * g + 1];
            x->caps[2 * g] = pos;
            x->caps[2 * g + 1] = -1;
            Cont c; c.kind = 1; c.a = g; c.b = 0; c.c = 0; c.next = k;
            if (mRun(x, n->b, pos, &c)) return 1;
            x->caps[2 * g] = os;
            x->caps[2 * g + 1] = oe;
            return 0;
        }
        case K_BOL:
            if (pos == 0 || (nodeML(x, n) && x->s[pos - 1] == L'\n')) return mCont(x, k, pos);
            return 0;
        case K_EOL:
            if (pos == x->len || (nodeML(x, n) && x->s[pos] == L'\n')) return mCont(x, k, pos);
            return 0;
        case K_WORDB:
            if (isWordAt(x, pos - 1) != isWordAt(x, pos)) return mCont(x, k, pos);
            return 0;
        case K_NWORDB:
            if (isWordAt(x, pos - 1) == isWordAt(x, pos)) return mCont(x, k, pos);
            return 0;
        case K_BACKREF: {
            int g = n->a;
            int s0 = x->caps[2 * g], e0 = x->caps[2 * g + 1];
            if (g == 0 || s0 < 0 || e0 < 0) return 0;
            int len = e0 - s0;
            if (pos + len > x->len) return 0;
            int i;
            for (i = 0; i < len; i++)
                if (!chEq(x, x->s[pos + i], x->s[s0 + i], nodeIC(x, n))) return 0;
            return mCont(x, k, pos + len);
        }
        case K_LOOK: {
            int saved[2 * (RE_MAX_GROUPS + 1)];
            memcpy(saved, x->caps, sizeof(saved));
            int hit = mRun(x, n->b, pos, NULL);
            if (n->d) {                        /* 负向：命中即失败 */
                memcpy(x->caps, saved, sizeof(saved));
                if (hit) return 0;
                return mCont(x, k, pos);
            }
            if (!hit) { memcpy(x->caps, saved, sizeof(saved)); return 0; }
            return mCont(x, k, pos);
        }
        /* ---- 扩展档运行期（默认开启） ---- */
        case K_ATOMIC: {
            /* 原子组：拿子模式的**第一次**成功落点，之后再不回退进去。
             * 用续延 kind=3 记落点，而不是看 mRun 的返回值 —— 中间层会把 2 归成 1，
             * 靠返回值判断会在 "K_ALT/K_REP 包一层" 时错判。 */
            Cont c; c.kind = 3; c.a = 0; c.b = 0; c.c = -1; c.next = NULL;
            (void)mRun(x, n->b, pos, &c);
            if (c.c < 0) return 0;
            return mCont(x, k, c.c);
        }
        case K_LBEHIND: {
            int saved[2 * (RE_MAX_GROUPS + 1)];
            int L = n->a, hit = 0;
            memcpy(saved, x->caps, sizeof(saved));
            if (pos - L >= 0) {
                Cont c; c.kind = 4; c.a = 0; c.b = pos; c.c = 0; c.next = NULL;
                hit = (mRun(x, n->b, pos - L, &c) != 0);
            }
            if (n->d) {
                memcpy(x->caps, saved, sizeof(saved));
                if (hit) return 0;
                return mCont(x, k, pos);
            }
            if (!hit) { memcpy(x->caps, saved, sizeof(saved)); return 0; }
            return mCont(x, k, pos);
        }
        case K_STRBOL:  return pos == 0 ? mCont(x, k, pos) : 0;
        case K_STREOL:  return pos == x->len ? mCont(x, k, pos) : 0;
        case K_STREOLZ:
            if (pos == x->len || (pos == x->len - 1 && x->s[pos] == L'\n')) return mCont(x, k, pos);
            return 0;
    }
    return 0;
}

// 从 start 起找第一个匹配；返回 1 并把起止写进 caps[0]/caps[1] 的对应位置
static int reSearch(ReCtx* x, int start, int* ms, int* me) {
    int pos;
    if (start > x->len) return 0;
    for (pos = start; pos <= x->len; pos++) {
        int i;
        for (i = 0; i < 2 * (RE_MAX_GROUPS + 1); i++) x->caps[i] = -1;
        x->steps = 0;
        x->caps[0] = pos;
        Cont c; c.kind = 1; c.a = 0; c.b = 0; c.c = 0; c.next = NULL;
        if (mRun(x, x->pg->root, pos, &c)) {
            *ms = x->caps[0];
            *me = x->caps[1];
            return 1;
        }
    }
    return 0;
}

// 单次匹配（含全部捕获），结果写进 out
typedef struct {
    int start, end;
    int subStart[RE_MAX_GROUPS + 1];
    int subEnd[RE_MAX_GROUPS + 1];
    int nsub;
} ReMatch;

static int reExecAt(ReCtx* x, int start, ReMatch* out) {
    int ms = 0, me = 0;
    if (!reSearch(x, start, &ms, &me)) return 0;
    out->start = ms;
    out->end = me;
    out->nsub = x->pg->ngroups;
    int g;
    for (g = 0; g <= x->pg->ngroups; g++) {
        out->subStart[g] = x->caps[2 * g];
        out->subEnd[g] = x->caps[2 * g + 1];
    }
    return 1;
}

// ============================================================
// IDispatch 对象层：RegExp / MatchCollection / Match / SubMatches
// ============================================================
#include <string.h>

#define RE_DID_PATTERN     1
#define RE_DID_GLOBAL      2
#define RE_DID_IGNORECASE  3
#define RE_DID_MULTILINE   4
#define RE_DID_TEST        5
#define RE_DID_REPLACE     6
#define RE_DID_EXECUTE     7
#define RE_DID_COUNT       1
#define RE_DID_FIRSTINDEX  1
#define RE_DID_LENGTH      2
#define RE_DID_SUBMATCHES  3

#define RE_MAX_MATCHES 100000

typedef struct ReRegExp ReRegExp;
typedef struct ReMatchColl ReMatchColl;
typedef struct ReMatchObj ReMatchObj;
typedef struct ReSubMatches ReSubMatches;

struct ReRegExp {
    const IDispatchVtbl* lpVtbl;
    LONG ref;
    BSTR pattern;
    int global, icase, multiline;
    ReProg* prog;
    int dirty;
};

struct ReMatchObj {
    const IDispatchVtbl* lpVtbl;
    LONG ref;
    BSTR src;
    int start, end;
    int nsub;                        /* 捕获组个数（不含 0） */
    int subStart[RE_MAX_GROUPS + 1];
    int subEnd[RE_MAX_GROUPS + 1];
};

struct ReMatchColl {
    const IDispatchVtbl* lpVtbl;
    LONG ref;
    BSTR src;
    int count;
    ReMatchObj** items;
};

struct ReSubMatches {
    const IDispatchVtbl* lpVtbl;
    LONG ref;
    ReMatchObj* owner;               /* AddRef 持有 */
};

static HRESULT STDMETHODCALLTYPE reObj_QueryInterface(IDispatch* p, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = (void*)p;
        p->lpVtbl->AddRef(p);
        return S_OK;
    }
    return E_NOINTERFACE;
}

typedef ULONG (STDMETHODCALLTYPE* ReReleaseFn)(IDispatch*);

static ULONG STDMETHODCALLTYPE re_AddRefGeneric(IDispatch* p) { return 0; }   /* 各对象自填 */

static HRESULT STDMETHODCALLTYPE re_GetTypeInfoCount(IDispatch* p, UINT* n) {
    (void)p;
    if (n) *n = 0;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE re_GetTypeInfo(IDispatch* p, UINT a, LCID b, ITypeInfo** c) {
    (void)p; (void)a; (void)b;
    if (c) *c = NULL;
    return E_NOTIMPL;
}

// ---------- 工具 ----------
static HRESULT reArgToBstr(VARIANT* v, BSTR* out) {
    *out = NULL;
    if (!v) { *out = SysAllocString(L""); return S_OK; }
    VARTYPE vt = (VARTYPE)(V_VT(v) & VT_TYPEMASK);
    if (vt == VT_BSTR) { *out = SysAllocString(V_BSTR(v)); return S_OK; }
    if (vt == VT_EMPTY || vt == VT_NULL) { *out = SysAllocString(L""); return S_OK; }
    VARIANT tmp;
    VariantInit(&tmp);
    HRESULT hr = vb6_ComChangeType(&tmp, v, VT_BSTR);
    if (FAILED(hr)) return hr;
    *out = V_BSTR(&tmp);
    V_VT(&tmp) = VT_EMPTY;
    return S_OK;
}

static int reEnsureProg(ReRegExp* r) {
    if (!r->dirty && r->prog) return 0;
    if (r->prog) { progFree(r->prog); r->prog = NULL; }
    int err = 0;
    r->prog = compilePattern(r->pattern ? r->pattern : L"", &err);
    r->dirty = 0;
    return r->prog ? 0 : (err ? err : RE_ERR_SYNTAX);
}

static const wchar_t* reErrDesc(int err) {
    switch (err) {
        case RE_ERR_MISSING_RB: return L"\u6b63\u5219\u8868\u8fbe\u5f0f\u4e2d\u7f3a\u5c11 ']'";
        case RE_ERR_MISSING_RP: return L"\u6b63\u5219\u8868\u8fbe\u5f0f\u4e2d\u7f3a\u5c11 ')'";
        default:                return L"\u6b63\u5219\u8868\u8fbe\u5f0f\u8bed\u6cd5\u9519\u8bef";
    }
}

// 整数下标参数：VB 传过来的可能是 VT_VARIANT / VT_BYREF，逐层拆开
// 只用 oleauto.h 里真有的宏（V_VARIANT 不存在，VS2022 的 cl 会报 C2102）
static int reArgIndex(VARIANT* v, int dflt) {
    int guard = 0;
    while (v && guard++ < 4) {
        VARTYPE vt = V_VT(v);
        if (vt & VT_BYREF) {
            if (vt == (VT_VARIANT | VT_BYREF)) {
                VARIANT* inner = V_VARIANTREF(v);
                if (!inner || inner == v) return dflt;
                v = inner;
                continue;
            }
            if (vt == (VT_I2 | VT_BYREF)) return (int)*V_I2REF(v);
            if (vt == (VT_I4 | VT_BYREF)) return (int)*V_I4REF(v);
            if (vt == (VT_R8 | VT_BYREF)) return (int)*V_R8REF(v);
            if (vt == (VT_BSTR | VT_BYREF)) return *V_BSTRREF(v) ? _wtoi(*V_BSTRREF(v)) : 0;
            return dflt;
        }
        if (vt == VT_I2)   return (int)V_I2(v);
        if (vt == VT_I4)   return (int)V_I4(v);
        if (vt == VT_R8)   return (int)V_R8(v);
        if (vt == VT_BSTR) return V_BSTR(v) ? _wtoi(V_BSTR(v)) : 0;
        if (vt == VT_EMPTY || vt == VT_NULL) return 0;
        return dflt;
    }
    return dflt;
}

static ReMatchObj* reMatchObjNew(BSTR src, const ReMatch* m) {
    ReMatchObj* o = (ReMatchObj*)calloc(1, sizeof(ReMatchObj));
    if (!o) return NULL;
    o->lpVtbl = NULL;   /* 由调用方填 */
    o->ref = 1;
    o->src = SysAllocString(src ? src : L"");
    o->start = m->start;
    o->end = m->end;
    o->nsub = m->nsub;
    int g;
    for (g = 0; g <= m->nsub; g++) { o->subStart[g] = m->subStart[g]; o->subEnd[g] = m->subEnd[g]; }
    return o;
}

// Match 的 BSTR 值（调用方 SysFreeString）
static BSTR reMatchValue(ReMatchObj* o) {
    int n = o->end - o->start;
    if (n <= 0) return SysAllocString(L"");
    BSTR b = SysAllocStringLen(o->src + o->start, (UINT)n);
    return b;
}

static BSTR reMatchSubValue(ReMatchObj* o, int g) {
    if (g < 1 || g > o->nsub) { vb6_RaiseError(9, (void*)SysAllocString(L"\u4e0b\u6807\u8d8a\u754c")); return NULL; }
    int s0 = o->subStart[g], e0 = o->subEnd[g];
    if (s0 < 0 || e0 < 0) return SysAllocString(L"");     /* 未参与的组取空串（cscript 实测） */
    return SysAllocStringLen(o->src + s0, (UINT)(e0 - s0));
}

// ---------- RegExp ----------
static ULONG STDMETHODCALLTYPE re_AddRef(IDispatch* p) {
    return (ULONG)InterlockedIncrement(&((ReRegExp*)p)->ref);
}

static ULONG STDMETHODCALLTYPE re_Release(IDispatch* p) {
    ReRegExp* r = (ReRegExp*)p;
    LONG left = InterlockedDecrement(&r->ref);
    if (left == 0) {
        if (r->pattern) SysFreeString(r->pattern);
        if (r->prog) progFree(r->prog);
        free(r);
    }
    return (ULONG)left;
}

static HRESULT STDMETHODCALLTYPE re_GetIDsOfNames(IDispatch* p, REFIID riid, LPOLESTR* names,
                                                  UINT cNames, LCID lcid, DISPID* ids) {
    (void)p; (void)riid; (void)lcid;
    if (!names || !ids) return E_POINTER;
    UINT i;
    for (i = 0; i < cNames; i++) {
        const wchar_t* n = names[i];
        if      (!_wcsicmp(n, L"Pattern"))    ids[i] = RE_DID_PATTERN;
        else if (!_wcsicmp(n, L"Global"))     ids[i] = RE_DID_GLOBAL;
        else if (!_wcsicmp(n, L"IgnoreCase")) ids[i] = RE_DID_IGNORECASE;
        else if (!_wcsicmp(n, L"Multiline"))  ids[i] = RE_DID_MULTILINE;
        else if (!_wcsicmp(n, L"Test"))       ids[i] = RE_DID_TEST;
        else if (!_wcsicmp(n, L"Replace"))    ids[i] = RE_DID_REPLACE;
        else if (!_wcsicmp(n, L"Execute"))    ids[i] = RE_DID_EXECUTE;
        else { ids[i] = DISPID_UNKNOWN; return DISP_E_UNKNOWNNAME; }
    }
    return S_OK;
}

static void reAppendStr(wchar_t** buf, int* len, int* cap, const wchar_t* s, int n) {
    if (n <= 0) return;
    if (*len + n + 1 > *cap) {
        int nc = *cap ? *cap : 64;
        while (nc < *len + n + 1) nc *= 2;
        *buf = (wchar_t*)realloc(*buf, sizeof(wchar_t) * nc);
        *cap = nc;
    }
    memcpy(*buf + *len, s, sizeof(wchar_t) * n);
    *len += n;
    (*buf)[*len] = 0;
}

// Replace：展开 $1..$9 / $& / $$ / $` / $'
static BSTR reBuildReplacement(const wchar_t* rep, const wchar_t* src, int srclen, const ReMatch* m) {
    wchar_t* out = NULL;
    int len = 0, cap = 0;
    int i = 0;
    reAppendStr(&out, &len, &cap, L"", 0);
    while (rep[i]) {
        if (rep[i] != L'$' || !rep[i + 1]) { reAppendStr(&out, &len, &cap, &rep[i], 1); i++; continue; }
        wchar_t nx = rep[i + 1];
        if (nx == L'$') { reAppendStr(&out, &len, &cap, L"$", 1); i += 2; continue; }
        if (nx == L'&') { reAppendStr(&out, &len, &cap, src + m->start, m->end - m->start); i += 2; continue; }
        if (nx == L'`') { reAppendStr(&out, &len, &cap, src, m->start); i += 2; continue; }
        if (nx == L'\'') { reAppendStr(&out, &len, &cap, src + m->end, srclen - m->end); i += 2; continue; }
        if (nx >= L'0' && nx <= L'9') {
            int g = nx - L'0';
            if (g == 0) { reAppendStr(&out, &len, &cap, src + m->start, m->end - m->start); i += 2; continue; }
            if (g <= m->nsub && m->subStart[g] >= 0 && m->subEnd[g] >= 0) {
                reAppendStr(&out, &len, &cap, src + m->subStart[g], m->subEnd[g] - m->subStart[g]);
                i += 2;
                continue;
            }
            i += 2;
            continue;
        }
        reAppendStr(&out, &len, &cap, L"$", 1);
        i++;
    }
    BSTR b = SysAllocString(out ? out : L"");
    free(out);
    return b;
}

static HRESULT reDoExecute(ReRegExp* r, VARIANT* arg, VARIANT* out);
static HRESULT reDoReplace(ReRegExp* r, VARIANT* strArg, VARIANT* repArg, VARIANT* out);

static HRESULT STDMETHODCALLTYPE re_Invoke(IDispatch* p, DISPID id, REFIID riid, LCID lcid,
                                           WORD flags, DISPPARAMS* dp, VARIANT* res,
                                           EXCEPINFO* ei, UINT* argErr) {
    ReRegExp* r = (ReRegExp*)p;
    (void)riid; (void)lcid; (void)ei; (void)argErr;
    if (flags & DISPATCH_PROPERTYPUT) {
        VARIANT* v = (dp && dp->cArgs >= 1) ? &dp->rgvarg[0] : NULL;
        if (!v) return DISP_E_BADPARAMCOUNT;
        switch (id) {
            case RE_DID_PATTERN: {
                BSTR b = NULL;
                HRESULT hr = reArgToBstr(v, &b);
                if (FAILED(hr)) return hr;
                if (r->pattern) SysFreeString(r->pattern);
                r->pattern = b;
                r->dirty = 1;               /* 语法错误延迟到下次使用（与 OCX 一致） */
                return S_OK;
            }
            case RE_DID_GLOBAL:     r->global = (V_I2(v) != 0) || (V_I4(v) != 0) || (V_BOOL(v) != 0); return S_OK;
            case RE_DID_IGNORECASE: r->icase = (V_I2(v) != 0) || (V_I4(v) != 0) || (V_BOOL(v) != 0); return S_OK;
            case RE_DID_MULTILINE:  r->multiline = (V_I2(v) != 0) || (V_I4(v) != 0) || (V_BOOL(v) != 0); return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
    }
    if (flags & DISPATCH_PROPERTYGET) {
        if (id == RE_DID_PATTERN)    { if (res) { V_VT(res) = VT_BSTR; V_BSTR(res) = SysAllocString(r->pattern ? r->pattern : L""); } return S_OK; }
        if (id == RE_DID_GLOBAL)     { if (res) { V_VT(res) = VT_BOOL; V_BOOL(res) = r->global ? VARIANT_TRUE : VARIANT_FALSE; } return S_OK; }
        if (id == RE_DID_IGNORECASE) { if (res) { V_VT(res) = VT_BOOL; V_BOOL(res) = r->icase ? VARIANT_TRUE : VARIANT_FALSE; } return S_OK; }
        if (id == RE_DID_MULTILINE)  { if (res) { V_VT(res) = VT_BOOL; V_BOOL(res) = r->multiline ? VARIANT_TRUE : VARIANT_FALSE; } return S_OK; }
    }
    if (!(flags & DISPATCH_METHOD)) return DISP_E_MEMBERNOTFOUND;
    switch (id) {
        case RE_DID_TEST: {
            if (!dp || dp->cArgs < 1) return DISP_E_BADPARAMCOUNT;
            VARIANT* a = &dp->rgvarg[0];
            int pe = reEnsureProg(r);
            if (pe) { (void)ei; return vb6_ComVbHresult(pe); }
            BSTR s = NULL;
            HRESULT hr = reArgToBstr(a, &s);
            if (FAILED(hr)) return hr;
            ReCtx x;
            memset(&x, 0, sizeof(x));
            x.pg = r->prog; x.s = s; x.len = (int)wcslen(s);
            x.icase = r->icase; x.multiline = r->multiline;
            int ms, me;
            int hit = reSearch(&x, 0, &ms, &me);
            SysFreeString(s);
            if (res) { V_VT(res) = VT_BOOL; V_BOOL(res) = hit ? VARIANT_TRUE : VARIANT_FALSE; }
            return S_OK;
        }
        case RE_DID_EXECUTE: {
            if (!dp || dp->cArgs < 1) return DISP_E_BADPARAMCOUNT;
            int pe = reEnsureProg(r);
            if (pe) { (void)ei; return vb6_ComVbHresult(pe); }
            return reDoExecute(r, &dp->rgvarg[0], res);
        }
        case RE_DID_REPLACE: {
            if (!dp || dp->cArgs < 2) return DISP_E_BADPARAMCOUNT;
            int pe = reEnsureProg(r);
            if (pe) { (void)ei; return vb6_ComVbHresult(pe); }
            return reDoReplace(r, &dp->rgvarg[1], &dp->rgvarg[0], res);
        }
    }
    return DISP_E_MEMBERNOTFOUND;
}

static const IDispatchVtbl reRegExpVtbl = {
    reObj_QueryInterface, re_AddRef, re_Release,
    re_GetTypeInfoCount, re_GetTypeInfo, re_GetIDsOfNames, re_Invoke
};

void* vb6_RegExp_New(void) {
    ReRegExp* r = (ReRegExp*)calloc(1, sizeof(ReRegExp));
    if (!r) return NULL;
    r->lpVtbl = &reRegExpVtbl;
    r->ref = 1;
    r->pattern = SysAllocString(L"");
    r->global = 0;
    r->icase = 0;
    r->multiline = 0;
    r->dirty = 1;
    return (void*)r;
}

int vb6_RegExp_IsProgramId(const wchar_t* progId) {
    if (!progId) return 0;
    return _wcsicmp(progId, L"VBScript.RegExp") == 0;
}

// ---------- Match ----------
static ULONG STDMETHODCALLTYPE m_AddRef(IDispatch* p) { return (ULONG)InterlockedIncrement(&((ReMatchObj*)p)->ref); }
static ULONG STDMETHODCALLTYPE m_Release(IDispatch* p) {
    ReMatchObj* o = (ReMatchObj*)p;
    LONG left = InterlockedDecrement(&o->ref);
    if (left == 0) { if (o->src) SysFreeString(o->src); free(o); }
    return (ULONG)left;
}

static HRESULT STDMETHODCALLTYPE m_GetIDsOfNames(IDispatch* p, REFIID riid, LPOLESTR* names,
                                                 UINT cNames, LCID lcid, DISPID* ids) {
    (void)p; (void)riid; (void)lcid;
    if (!names || !ids) return E_POINTER;
    UINT i;
    for (i = 0; i < cNames; i++) {
        const wchar_t* n = names[i];
        if      (!_wcsicmp(n, L"Value"))       ids[i] = DISPID_VALUE;
        else if (!_wcsicmp(n, L"FirstIndex"))  ids[i] = RE_DID_FIRSTINDEX;
        else if (!_wcsicmp(n, L"Length"))      ids[i] = RE_DID_LENGTH;
        else if (!_wcsicmp(n, L"SubMatches"))  ids[i] = RE_DID_SUBMATCHES;
        else { ids[i] = DISPID_UNKNOWN; return DISP_E_UNKNOWNNAME; }
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE m_Invoke(IDispatch* p, DISPID id, REFIID riid, LCID lcid,
                                          WORD flags, DISPPARAMS* dp, VARIANT* res,
                                          EXCEPINFO* ei, UINT* argErr) {
    ReMatchObj* o = (ReMatchObj*)p;
    (void)riid; (void)lcid; (void)dp; (void)ei; (void)argErr; (void)flags;
    if (id == DISPID_VALUE || id == 0) {
        if (res) { V_VT(res) = VT_BSTR; V_BSTR(res) = reMatchValue(o); }
        return S_OK;
    }
    if (id == RE_DID_FIRSTINDEX) { if (res) { V_VT(res) = VT_I4; V_I4(res) = o->start; } return S_OK; }
    if (id == RE_DID_LENGTH)     { if (res) { V_VT(res) = VT_I4; V_I4(res) = o->end - o->start; } return S_OK; }
    if (id == RE_DID_SUBMATCHES) {
        /* VB 两种写法都是同一件事：m.SubMatches(i) 可能落到「属性带参」这条路，
         * 也可能先取对象再调默认成员。两条都接，否则只会读到第 0 组。 */
        if ((flags & DISPATCH_PROPERTYGET) && dp && dp->cArgs >= 1) {
            int g = reArgIndex(&dp->rgvarg[0], 0);
            BSTR v = reMatchSubValue(o, g + 1);
            if (!v) return DISP_E_EXCEPTION;
            if (res) { V_VT(res) = VT_BSTR; V_BSTR(res) = v; } else SysFreeString(v);
            return S_OK;
        }
        if (!res) return S_OK;
        ReSubMatches* sm = (ReSubMatches*)calloc(1, sizeof(ReSubMatches));
        if (!sm) return E_OUTOFMEMORY;
        sm->lpVtbl = NULL;   /* 下面填 */
        sm->ref = 1;
        sm->owner = o;
        o->lpVtbl->AddRef(p);
        extern const IDispatchVtbl reSubMatchesVtbl;
        sm->lpVtbl = &reSubMatchesVtbl;
        V_VT(res) = VT_DISPATCH;
        V_DISPATCH(res) = (IDispatch*)sm;
        return S_OK;
    }
    return DISP_E_MEMBERNOTFOUND;
}

static const IDispatchVtbl reMatchVtbl = {
    reObj_QueryInterface, m_AddRef, m_Release,
    re_GetTypeInfoCount, re_GetTypeInfo, m_GetIDsOfNames, m_Invoke
};

// ---------- SubMatches ----------
static ULONG STDMETHODCALLTYPE sm_AddRef(IDispatch* p) { return (ULONG)InterlockedIncrement(&((ReSubMatches*)p)->ref); }
static ULONG STDMETHODCALLTYPE sm_Release(IDispatch* p) {
    ReSubMatches* s = (ReSubMatches*)p;
    LONG left = InterlockedDecrement(&s->ref);
    if (left == 0) {
        if (s->owner) s->owner->lpVtbl->Release((IDispatch*)s->owner);
        free(s);
    }
    return (ULONG)left;
}

static HRESULT STDMETHODCALLTYPE sm_GetIDsOfNames(IDispatch* p, REFIID riid, LPOLESTR* names,
                                                  UINT cNames, LCID lcid, DISPID* ids) {
    (void)p; (void)riid; (void)lcid;
    if (!names || !ids) return E_POINTER;
    UINT i;
    for (i = 0; i < cNames; i++) {
        const wchar_t* n = names[i];
        if      (!_wcsicmp(n, L"Count")) ids[i] = RE_DID_COUNT;
        else if (!_wcsicmp(n, L"Item"))  ids[i] = DISPID_VALUE;
        else { ids[i] = DISPID_UNKNOWN; return DISP_E_UNKNOWNNAME; }
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE sm_Invoke(IDispatch* p, DISPID id, REFIID riid, LCID lcid,
                                           WORD flags, DISPPARAMS* dp, VARIANT* res,
                                           EXCEPINFO* ei, UINT* argErr) {
    ReSubMatches* s = (ReSubMatches*)p;
    (void)riid; (void)lcid; (void)ei; (void)argErr; (void)flags;
    if (id == RE_DID_COUNT) { if (res) { V_VT(res) = VT_I4; V_I4(res) = s->owner->nsub; } return S_OK; }
    if (id == DISPID_VALUE) {
        int g = 0;
        if (dp && dp->cArgs >= 1) g = reArgIndex(&dp->rgvarg[0], 0);
        BSTR v = reMatchSubValue(s->owner, g + 1);    /* SubMatches 是 0 基，内部捕获组 1 基 */
        if (!v) return DISP_E_EXCEPTION;
        if (res) { V_VT(res) = VT_BSTR; V_BSTR(res) = v; } else SysFreeString(v);
        return S_OK;
    }
    return DISP_E_MEMBERNOTFOUND;
}

const IDispatchVtbl reSubMatchesVtbl = {
    reObj_QueryInterface, sm_AddRef, sm_Release,
    re_GetTypeInfoCount, re_GetTypeInfo, sm_GetIDsOfNames, sm_Invoke
};

// ---------- MatchCollection ----------
static ULONG STDMETHODCALLTYPE mc_AddRef(IDispatch* p) { return (ULONG)InterlockedIncrement(&((ReMatchColl*)p)->ref); }
static ULONG STDMETHODCALLTYPE mc_Release(IDispatch* p) {
    ReMatchColl* c = (ReMatchColl*)p;
    LONG left = InterlockedDecrement(&c->ref);
    if (left == 0) {
        int i;
        for (i = 0; i < c->count; i++)
            if (c->items[i]) c->items[i]->lpVtbl->Release((IDispatch*)c->items[i]);
        free(c->items);
        if (c->src) SysFreeString(c->src);
        free(c);
    }
    return (ULONG)left;
}

static HRESULT STDMETHODCALLTYPE mc_GetIDsOfNames(IDispatch* p, REFIID riid, LPOLESTR* names,
                                                  UINT cNames, LCID lcid, DISPID* ids) {
    (void)p; (void)riid; (void)lcid;
    if (!names || !ids) return E_POINTER;
    UINT i;
    for (i = 0; i < cNames; i++) {
        const wchar_t* n = names[i];
        if      (!_wcsicmp(n, L"Count")) ids[i] = RE_DID_COUNT;
        else if (!_wcsicmp(n, L"Item"))  ids[i] = DISPID_VALUE;
        else { ids[i] = DISPID_UNKNOWN; return DISP_E_UNKNOWNNAME; }
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE mc_Invoke(IDispatch* p, DISPID id, REFIID riid, LCID lcid,
                                           WORD flags, DISPPARAMS* dp, VARIANT* res,
                                           EXCEPINFO* ei, UINT* argErr) {
    ReMatchColl* c = (ReMatchColl*)p;
    (void)riid; (void)lcid; (void)ei; (void)argErr; (void)flags;
    if (id == RE_DID_COUNT) { if (res) { V_VT(res) = VT_I4; V_I4(res) = c->count; } return S_OK; }
    if (id == DISPID_VALUE) {
        int idx = 0;
        if (dp && dp->cArgs >= 1) idx = reArgIndex(&dp->rgvarg[0], 0);
        if (idx < 0 || idx >= c->count) {
            vb6_RaiseError(9, (void*)SysAllocString(L"\u4e0b\u6807\u8d8a\u754c"));
            return DISP_E_EXCEPTION;
        }
        if (res) {
            ReMatchObj* o = c->items[idx];
            o->lpVtbl->AddRef((IDispatch*)o);
            V_VT(res) = VT_DISPATCH;
            V_DISPATCH(res) = (IDispatch*)o;
        }
        return S_OK;
    }
    return DISP_E_MEMBERNOTFOUND;
}

static const IDispatchVtbl reMatchCollVtbl = {
    reObj_QueryInterface, mc_AddRef, mc_Release,
    re_GetTypeInfoCount, re_GetTypeInfo, mc_GetIDsOfNames, mc_Invoke
};

// ---------- Execute / Replace ----------
static HRESULT reDoExecute(ReRegExp* r, VARIANT* arg, VARIANT* out) {
    if (!r->prog) return DISP_E_EXCEPTION;
    BSTR s = NULL;
    HRESULT hr = reArgToBstr(arg, &s);
    if (FAILED(hr)) return hr;
    ReCtx x;
    memset(&x, 0, sizeof(x));
    x.pg = r->prog;
    x.s = s;
    x.len = (int)wcslen(s);
    x.icase = r->icase;
    x.multiline = r->multiline;

    ReMatchColl* c = (ReMatchColl*)calloc(1, sizeof(ReMatchColl));
    if (!c) { SysFreeString(s); return E_OUTOFMEMORY; }
    c->lpVtbl = &reMatchCollVtbl;
    c->ref = 1;
    c->src = SysAllocString(s);
    int cap = 8;
    c->items = (ReMatchObj**)calloc(cap, sizeof(ReMatchObj*));
    int pos = 0;
    for (;;) {
        ReMatch m;
        if (!reExecAt(&x, pos, &m)) break;
        if (c->count == cap) {
            cap *= 2;
            c->items = (ReMatchObj**)realloc(c->items, sizeof(ReMatchObj*) * cap);
        }
        ReMatchObj* o = reMatchObjNew(s, &m);
        o->lpVtbl = &reMatchVtbl;
        c->items[c->count++] = o;
        if (!r->global) break;
        if (c->count >= RE_MAX_MATCHES) break;
        pos = (m.end == m.start) ? m.end + 1 : m.end;    /* 空匹配前进 1（cscript 实测 4 个 0:0） */
        if (pos > x.len) break;
    }
    SysFreeString(s);
    if (out) { V_VT(out) = VT_DISPATCH; V_DISPATCH(out) = (IDispatch*)c; }
    else mc_Release((IDispatch*)c);
    return S_OK;
}

static HRESULT reDoReplace(ReRegExp* r, VARIANT* strArg, VARIANT* repArg, VARIANT* out) {
    if (!r->prog) return DISP_E_EXCEPTION;
    BSTR s = NULL, rep = NULL;
    HRESULT hr = reArgToBstr(strArg, &s);
    if (FAILED(hr)) return hr;
    hr = reArgToBstr(repArg, &rep);
    if (FAILED(hr)) { SysFreeString(s); return hr; }

    ReCtx x;
    memset(&x, 0, sizeof(x));
    x.pg = r->prog;
    x.s = s;
    x.len = (int)wcslen(s);
    x.icase = r->icase;
    x.multiline = r->multiline;

    wchar_t* acc = NULL;
    int len = 0, cap = 0;
    reAppendStr(&acc, &len, &cap, L"", 0);
    int pos = 0, last = 0, guard = 0;
    for (;;) {
        ReMatch m;
        if (!reExecAt(&x, pos, &m)) break;
        reAppendStr(&acc, &len, &cap, s + last, m.start - last);
        BSTR piece = reBuildReplacement(rep, s, x.len, &m);
        reAppendStr(&acc, &len, &cap, piece, (int)wcslen(piece));
        SysFreeString(piece);
        last = m.end;
        if (!r->global) break;
        if (++guard >= RE_MAX_MATCHES) break;
        pos = (m.end == m.start) ? m.end + 1 : m.end;
        if (pos > x.len) break;
    }
    reAppendStr(&acc, &len, &cap, s + last, x.len - last);
    BSTR result = SysAllocString(acc ? acc : L"");
    free(acc);
    SysFreeString(s);
    SysFreeString(rep);
    if (out) { V_VT(out) = VT_BSTR; V_BSTR(out) = result; }
    else SysFreeString(result);
    return S_OK;
}


