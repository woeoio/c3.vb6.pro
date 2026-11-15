/* vb6forms_teechart.c —— TeeChart.TChart 的内置实现（C30-10，第一刀：窗口 + 绘制骨架）。
 *
 * 为什么要有这一刀：TeeChart5.ocx **只有 32 位**，走 ActiveX 宿主（axsite/）时
 *   --arch x64 永远链不上（tests/run_tests.ps1 那条用例因此被写成 Arch="x86"、
 *   且 RequiresFiles=TeeChart5.ocx —— CI 上只能 SKIP）。内置之后两档架构走同一条
 *   自注册窗口类 VB6_TCHART，不再碰 CoCreateInstance / 免注册清单。
 *
 * 本刀做到（P1 的第一半）：
 *   - 控件窗口 VB6_TCHART（自注册；comctl32 里没有这个类）；
 *   - 完整的一次绘制：Panel 底色/边框 + Header 居中 + Legend 边框与条目 +
 *     上/下轴线 + 数据系列（Line / Bar / Point / Area / Pie）；
 *   - 运行期模型：Series 数组（类型/颜色/标题/点集）、Panel 边距与底色、Header 文本；
 *     Series(i) **访问即建号**（TPF0 设计期流要到 P2 才有，见 ai/内置控件/TeeChart*）；
 *   - 3 枚通用入口 vb6_TChart_GetMember / SetMember / Call（发码侧唯一的新口子，
 *     契约见同目录文档 §3）。深成员（Series(i).AsPie.ExplodedSlice.Value(i) 那种）
 *     下一步接 memberobj 的 IDispatch，本刀先不接。
 *
 * 没做（如实留着，P2/P3 补）：TPF0 反序列化、AsPie/AsSurface/AsCandle 那族类型切面、
 * 3D（Aspect.View3D / Walls / Perspective）、ShowEditor / Printer / Export / Zoom。
 */

#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include <stdlib.h>
#include <math.h>
#include "vb6forms.h"
#include "vb6forms_internal.h"

#define VB6_TC_CLASS     L"VB6_TCHART"
#define VB6_TC_MAXCTL    32      /* 同进程最多几枚图表 */
#define VB6_TC_MAXSER    32      /* 每图最多几条系列 */
#define VB6_TC_MAXPTS    4096    /* 每条系列最多几个点 */

/* TeeChart 的 SeriesType 常数（真 OCX 的 scLine..scPie 取值，示例里直接写数字也要对得上） */
#define TC_SC_LINE      0
#define TC_SC_BAR       1
#define TC_SC_HBAR      2
#define TC_SC_AREA      3
#define TC_SC_POINT     4
#define TC_SC_PIE       5

typedef struct { double x, y; } TcPoint;

typedef struct {
    int    type;                 /* TC_SC_* */
    COLORREF color;
    wchar_t title[64];
    TcPoint pt[VB6_TC_MAXPTS];
    int    n;
} TcSeries;

typedef struct {
    HWND     hwnd;
    COLORREF panelColor;
    int      marginL, marginR, marginT, marginB;   /* 像素 */
    wchar_t  header[256];
    int      showLegend;
    TcSeries ser[VB6_TC_MAXSER];
    int      nser;
    int      paints;             /* 绘制次数（C3_TCHART_TRACE 用） */
    struct TcObj* obj[VB6_TC_MAXSER + 1];   /* 成员对象缓存 (0 = Series 集合) */
    struct TcObj* meth[8];                  /* 方法对象缓存（AddSeries 那几枚） */
    int      nmeth;
} TcChart;

static TcChart* g_tc[VB6_TC_MAXCTL];
static int      g_tcN = 0;
static int      g_tcTrace = -1;

static int tcTrace(void) {
    if (g_tcTrace < 0) {
        char b[16]; DWORD n = GetEnvironmentVariableA("C3_TCHART_TRACE", b, sizeof(b));
        g_tcTrace = (n > 0 && b[0] && b[0] != '0') ? 1 : 0;
    }
    return g_tcTrace;
}

/* 窗口 ↔ 状态：状态随窗口走（SetPropW），控件销毁时释放，不靠全局名册查。 */
static const wchar_t* TC_SLOT = L"VB6_TChartState";

static TcChart* tcOf(HWND h) {
    if (!h || !IsWindow(h)) return NULL;
    return (TcChart*)GetPropW(h, TC_SLOT);
}

static TcChart* tcAttach(HWND h) {
    TcChart* c;
    if (!h || !IsWindow(h)) return NULL;
    c = tcOf(h);
    if (c) return c;
    c = (TcChart*)calloc(1, sizeof(TcChart));
    if (!c) return NULL;
    c->hwnd          = h;
    c->panelColor    = RGB(255, 255, 255);
    c->marginL       = 12;
    c->marginR       = 12;
    c->marginT       = 12;
    c->marginB       = 12;
    c->showLegend    = 1;
    SetPropW(h, TC_SLOT, (HANDLE)c);
    if (g_tcN < VB6_TC_MAXCTL) g_tc[g_tcN++] = c;
    return c;
}

static TcSeries* tcSeries(TcChart* c, int idx) {
    /* Series(i) 访问即建号：P2 之前 TPF0 还没解析，示例 Form_Load 里第一句就是
       Series(0).FillSampleValues —— 不自动建号的话整条链会落在 Nothing 上。
       建号默认 Line（与 TeeChart 新建 Chart 的默认系列一致）。 */
    if (!c || idx < 0 || idx >= VB6_TC_MAXSER) return NULL;
    while (c->nser <= idx) {
        TcSeries* s = &c->ser[c->nser];
        memset(s, 0, sizeof(*s));
        s->type  = TC_SC_LINE;
        s->color = RGB(0, 128, 255);
        wsprintfW(s->title, L"Series%d", c->nser);
        c->nser++;
    }
    return &c->ser[idx];
}

/* 尾叉用：先认窗口类再取（不认窗口就不碰 —— 宿主模型每枚控件都会问到这里）。
 * 用 tcAttach 而不是 tcOf：Form_Load 里的调用**早于第一次 WM_PAINT**，
 * 那时状态还没挂上；只看 tcOf 会在 Show 之前错答 438。 */
static TcChart* tcOfClass(HWND h) {
    wchar_t cls[64];
    if (!h || !IsWindow(h)) return NULL;
    if (!GetClassNameW(h, cls, 64)) return NULL;
    if (wcscmp(cls, VB6_TC_CLASS) != 0) return NULL;
    return tcAttach(h);
}

/* 默认调色板（与 TeeChart 出厂配色同序，够 P1 的判据用） */
static COLORREF tcPalette(int i) {
    static const COLORREF p[8] = {
        RGB(0,128,255), RGB(255,0,0), RGB(0,192,0), RGB(255,192,0),
        RGB(128,0,255), RGB(0,192,192), RGB(255,0,255), RGB(128,128,0)
    };
    return p[i & 7];
}

/* ============================ 绘制 ============================ */

static void tcDrawText(HDC dc, const wchar_t* s, RECT* r, int center, COLORREF col) {
    if (!s || !*s) return;
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, col);
    DrawTextW(dc, s, -1, r, center ? (DT_CENTER | DT_VCENTER | DT_SINGLELINE)
                                   : (DT_LEFT | DT_VCENTER | DT_SINGLELINE));
}

static void tcPaint(TcChart* c) {
    PAINTSTRUCT ps;
    HDC dc;
    RECT rc, plot;
    HBRUSH br;
    HPEN   oldPen, pen;
    HGDIOBJ oldBr;
    int i, n, x0, x1, y0, y1;
    double vmin, vmax;

    if (!c) return;
    dc = BeginPaint(c->hwnd, &ps);
    if (!dc) return;
    GetClientRect(c->hwnd, &rc);

    /* Panel 底 + 外框（TeeChart 的 Panel 是一块带 1px 边的主面） */
    br = CreateSolidBrush(c->panelColor);
    oldBr = SelectObject(dc, br);
    FillRect(dc, &rc, br);
    SelectObject(dc, oldBr);
    DeleteObject(br);
    pen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));
    oldPen = (HPEN)SelectObject(dc, pen);
    MoveToEx(dc, rc.left, rc.top, NULL);
    LineTo(dc, rc.right - 1, rc.top);
    LineTo(dc, rc.right - 1, rc.bottom - 1);
    LineTo(dc, rc.left, rc.bottom - 1);
    LineTo(dc, rc.left, rc.top);

    /* 绘图区 = 客户区扣掉 Panel 四边距 */
    plot = rc;
    plot.left   += c->marginL;
    plot.right  -= c->marginR;
    plot.top    += c->marginT;
    plot.bottom -= c->marginB;
    if (plot.right <= plot.left + 2 || plot.bottom <= plot.top + 2) {
        SelectObject(dc, oldPen); DeleteObject(pen);
        EndPaint(c->hwnd, &ps);
        return;
    }

    /* Header：居中在绘图区上方（参考读数里标题就在这一带） */
    if (c->header[0]) {
        RECT hr = plot;
        hr.bottom = hr.top + 16;
        tcDrawText(dc, c->header, &hr, 1, RGB(0, 0, 0));
    }

    /* 上/下轴线：TeeChart 默认画在绘图区上下沿 */
    SelectObject(dc, pen);
    MoveToEx(dc, plot.left, plot.top, NULL);
    LineTo(dc, plot.right, plot.top);
    MoveToEx(dc, plot.left, plot.bottom, NULL);
    LineTo(dc, plot.right, plot.bottom);

    /* Y 轴量程（所有系列合起来；空系列给 0..1） */
    vmin = 1e300; vmax = -1e300;
    for (i = 0; i < c->nser; i++) {
        TcSeries* s = &c->ser[i];
        int k;
        if (s->type == TC_SC_PIE) continue;
        for (k = 0; k < s->n; k++) {
            if (s->pt[k].y < vmin) vmin = s->pt[k].y;
            if (s->pt[k].y > vmax) vmax = s->pt[k].y;
        }
    }
    if (vmax <= vmin) { vmin = 0; vmax = 1; }

    x0 = plot.left; x1 = plot.right;
    y0 = plot.top;  y1 = plot.bottom;

    /* 系列 */
    for (i = 0; i < c->nser; i++) {
        TcSeries* s = &c->ser[i];
        int k;
        COLORREF col = (s->color ? s->color : tcPalette(i));
        HPEN sp = CreatePen(PS_SOLID, 2, col);
        HPEN spOld = (HPEN)SelectObject(dc, sp);
        HBRUSH sb = CreateSolidBrush(col);
        HGDIOBJ sbOld = SelectObject(dc, sb);

        if (s->type == TC_SC_PIE) {
            /* 饼图：按 y 值占比切扇区（起始 -90°，与 TeeChart 默认同向）。
             * Pie() 的两个“终点”各自是扇区的起/止半径端点 —— 角度算错就会画出
             * 整圆或零面积，所以这里统一用 cos/sin 落在圆周上。 */
            double sum = 0, ang = -90.0;
            int cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
            int r  = ((x1 - x0) < (y1 - y0) ? (x1 - x0) : (y1 - y0)) / 2 - 2;
            if (r < 4) r = 4;
            for (k = 0; k < s->n; k++) sum += (s->pt[k].y > 0 ? s->pt[k].y : 0);
            if (sum <= 0) sum = 1;
            for (k = 0; k < s->n; k++) {
                double sweep = (s->pt[k].y > 0 ? s->pt[k].y : 0) / sum * 360.0;
                double a0 = ang * 3.14159265358979323846 / 180.0;
                double a1 = (ang + sweep) * 3.14159265358979323846 / 180.0;
                int sx = cx + (int)(r * cos(a0)), sy = cy + (int)(r * sin(a0));
                int ex = cx + (int)(r * cos(a1)), ey = cy + (int)(r * sin(a1));
                HBRUSH pb = CreateSolidBrush(tcPalette(k));
                HGDIOBJ pbOld = SelectObject(dc, pb);
                HGDIOBJ ob = SelectObject(dc, GetStockObject(NULL_PEN));
                if (sweep >= 359.9) {
                    Ellipse(dc, cx - r, cy - r, cx + r, cy + r);
                } else if (sweep > 0.01) {
                    Pie(dc, cx - r, cy - r, cx + r, cy + r, sx, sy, ex, ey);
                }
                SelectObject(dc, ob);
                SelectObject(dc, pbOld);
                DeleteObject(pb);
                ang += sweep;
            }
        } else if (s->n > 0) {
            int bw, gap, k2;
            switch (s->type) {
            case TC_SC_BAR:
                bw  = (s->n > 0) ? ((x1 - x0) / (s->n * 2)) : 8;
                gap = bw;
                if (bw < 2) bw = 2;
                for (k2 = 0; k2 < s->n; k2++) {
                    int bx = x0 + gap / 2 + k2 * (bw + gap);
                    int by = y1 - (int)((s->pt[k2].y - vmin) / (vmax - vmin) * (y1 - y0));
                    Rectangle(dc, bx, by, bx + bw, y1);
                }
                break;
            case TC_SC_POINT:
                for (k2 = 0; k2 < s->n; k2++) {
                    int px = (s->n == 1) ? (x0 + x1) / 2
                            : x0 + (int)((double)k2 / (s->n - 1) * (x1 - x0));
                    int py = y1 - (int)((s->pt[k2].y - vmin) / (vmax - vmin) * (y1 - y0));
                    Ellipse(dc, px - 3, py - 3, px + 3, py + 3);
                }
                break;
            default:   /* LINE / AREA：折线（AREA 再填充到基线） */
                if (s->type == TC_SC_AREA) {
                    POINT* poly = (POINT*)malloc(sizeof(POINT) * (s->n + 2));
                    if (poly) {
                        for (k2 = 0; k2 < s->n; k2++) {
                            poly[k2].x = (s->n == 1) ? (x0 + x1) / 2
                                       : x0 + (int)((double)k2 / (s->n - 1) * (x1 - x0));
                            poly[k2].y = y1 - (int)((s->pt[k2].y - vmin) / (vmax - vmin) * (y1 - y0));
                        }
                        poly[s->n].x = poly[s->n - 1].x; poly[s->n].y = y1;
                        poly[s->n + 1].x = poly[0].x;    poly[s->n + 1].y = y1;
                        Polygon(dc, poly, s->n + 2);
                        free(poly);
                    }
                }
                for (k2 = 0; k2 < s->n; k2++) {
                    int px = (s->n == 1) ? (x0 + x1) / 2
                            : x0 + (int)((double)k2 / (s->n - 1) * (x1 - x0));
                    int py = y1 - (int)((s->pt[k2].y - vmin) / (vmax - vmin) * (y1 - y0));
                    if (k2 == 0) MoveToEx(dc, px, py, NULL); else LineTo(dc, px, py);
                }
                break;
            }
        }
        SelectObject(dc, spOld); DeleteObject(sp);
        SelectObject(dc, sbOld); DeleteObject(sb);
    }

    /* Legend：右上角一框，条目 = 色块 + 系列标题（TeeChart 默认图例在右侧） */
    if (c->showLegend && c->nser > 0) {
        int lw = 0, lh, k;
        RECT lr;
        for (k = 0; k < c->nser; k++) {
            int tw = 0;
            SIZE sz;
            if (GetTextExtentPoint32W(dc, c->ser[k].title, (int)wcslen(c->ser[k].title), &sz))
                tw = sz.cx;
            if (tw > lw) lw = tw;
        }
        lh = c->nser * 14 + 6;
        lr.right  = x1 - 2;
        lr.left   = lr.right - (lw + 24);
        lr.top    = y0 + 2;
        lr.bottom = lr.top + lh;
        if (lr.left < x0) lr.left = x0;
        Rectangle(dc, lr.left, lr.top, lr.right, lr.bottom);
        for (k = 0; k < c->nser; k++) {
            RECT ibr, tr;
            HBRUSH cb = CreateSolidBrush(c->ser[k].color ? c->ser[k].color : tcPalette(k));
            ibr.left = lr.left + 3; ibr.right = ibr.left + 10;
            ibr.top  = lr.top + 3 + k * 14; ibr.bottom = ibr.top + 10;
            FillRect(dc, &ibr, cb);
            DeleteObject(cb);
            tr = ibr; tr.left = ibr.right + 3; tr.right = lr.right - 2;
            tcDrawText(dc, c->ser[k].title, &tr, 0, RGB(0, 0, 0));
        }
    }

    SelectObject(dc, oldPen);
    DeleteObject(pen);
    c->paints++;
    if (tcTrace()) {
        fprintf(stderr, "[TCHART] paint=%d series=%d hwnd=%p\n",
                c->paints, c->nser, (void*)c->hwnd);
        fflush(stderr);
    }
    EndPaint(c->hwnd, &ps);
}

static LRESULT CALLBACK vb6_TChartWndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    TcChart* c = tcOf(h);
    switch (m) {
    case WM_PAINT:
        c = tcAttach(h);
        tcPaint(c);
        return 0;
    case WM_ERASEBKGND:
        return 1;   /* 全在 WM_PAINT 里画，不闪 */
    case WM_SIZE:
        InvalidateRect(h, NULL, FALSE);
        return 0;
    case WM_NCDESTROY:
        if (c) {
            int i;
            RemovePropW(h, TC_SLOT);
            for (i = 0; i < g_tcN; i++) if (g_tc[i] == c) { g_tc[i] = g_tc[--g_tcN]; break; }
            free(c);
        }
        return 0;
    default:
        return DefWindowProcW(h, m, w, l);
    }
}

void vb6_RegisterTChartClass(void* hInstance) {
    static int done = 0;
    WNDCLASSW wc;
    if (done) return;
    done = 1;
    memset(&wc, 0, sizeof(wc));
    wc.lpfnWndProc   = vb6_TChartWndProc;
    wc.hInstance     = (HINSTANCE)hInstance;
    wc.hCursor       = LoadCursorW(NULL, (IDC_ARROW));
    wc.hbrBackground = NULL;       /* 自己在 WM_PAINT 里铺 Panel 底 */
    wc.lpszClassName = VB6_TC_CLASS;
    RegisterClassW(&wc);
}

/* ============================ 运行期模型 ============================ */

int32_t vb6_TChart_SeriesCount(void* hwnd) {
    TcChart* c = tcOf((HWND)hwnd);
    return c ? c->nser : 0;
}

int32_t vb6_TChart_AddSeries(void* hwnd, int32_t type) {
    TcChart* c = tcAttach((HWND)hwnd);
    TcSeries* s;
    if (!c) return -1;
    if (c->nser >= VB6_TC_MAXSER) return -1;
    s = &c->ser[c->nser];
    memset(s, 0, sizeof(*s));
    s->type  = (type >= TC_SC_LINE && type <= TC_SC_PIE) ? (int)type : TC_SC_LINE;
    s->color = tcPalette(c->nser);
    wsprintfW(s->title, L"Series%d", c->nser);
    c->nser++;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
    return c->nser - 1;
}

int32_t vb6_TChart_SeriesAddXY(void* hwnd, int32_t idx, double x, double y) {
    TcChart* c = tcOf((HWND)hwnd);
    TcSeries* s;
    if (!c) return 0;
    s = tcSeries(c, (int)idx);
    if (!s || s->n >= VB6_TC_MAXPTS) return 0;
    s->pt[s->n].x = x;
    s->pt[s->n].y = y;
    s->n++;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
    return s->n;
}

int32_t vb6_TChart_SeriesAdd(void* hwnd, int32_t idx, double y) {
    TcChart* c = tcOf((HWND)hwnd);
    TcSeries* s = c ? tcSeries(c, (int)idx) : NULL;
    return vb6_TChart_SeriesAddXY(hwnd, idx, s ? (double)s->n : 0.0, y);
}

int32_t vb6_TChart_FillSampleValues(void* hwnd, int32_t idx, int32_t count) {
    int k;
    if (count <= 0) count = 1;
    if (count > 64) count = 64;
    for (k = 0; k < count; k++)
        vb6_TChart_SeriesAddXY(hwnd, idx, k, (double)((k * 37) % 100) + 1.0);
    return count;
}

void vb6_TChart_SeriesClear(void* hwnd, int32_t idx) {
    TcChart* c = tcOf((HWND)hwnd);
    TcSeries* s = c ? tcSeries(c, (int)idx) : NULL;
    if (!s) return;
    s->n = 0;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

void vb6_TChart_SetSeriesColor(void* hwnd, int32_t idx, int32_t bgr) {
    TcChart* c = tcOf((HWND)hwnd);
    TcSeries* s = c ? tcSeries(c, (int)idx) : NULL;
    if (!s) return;
    s->color = (COLORREF)bgr;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

void vb6_TChart_SetSeriesTitle(void* hwnd, int32_t idx, const wchar_t* title) {
    TcChart* c = tcOf((HWND)hwnd);
    TcSeries* s = c ? tcSeries(c, (int)idx) : NULL;
    if (!s) return;
    wsprintfW(s->title, L"%s", title && *title ? title : L"");
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

void vb6_TChart_SetHeaderText(void* hwnd, const wchar_t* text) {
    TcChart* c = tcAttach((HWND)hwnd);
    if (!c) return;
    wsprintfW(c->header, L"%s", text ? text : L"");
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

void vb6_TChart_SetPanelColor(void* hwnd, int32_t bgr) {
    TcChart* c = tcAttach((HWND)hwnd);
    if (!c) return;
    c->panelColor = (COLORREF)bgr;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

void vb6_TChart_SetPanelMargin(void* hwnd, int32_t which, int32_t px) {
    TcChart* c = tcAttach((HWND)hwnd);
    if (!c) return;
    switch (which) {
    case 0: c->marginL = px; break;
    case 1: c->marginR = px; break;
    case 2: c->marginT = px; break;
    default: c->marginB = px; break;
    }
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

void vb6_TChart_Repaint(void* hwnd) {
    if (hwnd) InvalidateRect((HWND)hwnd, NULL, FALSE);
}

int32_t vb6_TChart_PaintCount(void* hwnd) {
    TcChart* c = tcOf((HWND)hwnd);
    return c ? c->paints : 0;
}

/* ==================== 成员对象：Series 集合 / Series（C30-10 第二刀） ====================
 *
 * 发码侧零改动：cgen 对 `TChart1.Series(0).FillSampleValues 4` 发的已经是通用晚绑定链
 * （vb6_ComGetObjectProp(hwnd,L"Series") → vb6_ComCall(默认成员) → vb6_ComCall(L"FillSampleValues")，
 * 实测见 ai/内置控件/TeeChart* §3.1）。缺的只有“接收者是控件 HWND”这件事：
 * 宿主管道已经会把 HWND 当宿主对象处理（vb6_Host_GetProp/Call），所以这里只补
 * 一条支路：GetProp("Series") 交出一枚**真 IDispatch** 集合，后续每一跳都由这张
 * vtbl 应答，不需要 cgen 特例。
 *
 * 为什么不往 vb6forms_memberobj.c 里加族：那一份是 1300 行的共用分派（文件头记了
 * 一串“次序搞错就静默吞掉”的坑），为一块新控件往里插队风险大于收益。这里自带两张
 * 名表 + 一次 Invoke 分派，形态与它一致（真 IDispatch，晚绑定/For Each 都认）。
 *
 * 生命周期：对象归图表所有（TcChart.obj[] 缓存），Release 只减计数、不归还内存，
 * 窗口销毁时一并释放 —— 避免“交出去的引用被 VariantClear 放掉、调用方还在用”这类
 * 双重释放；一枚对象 24 字节，按图缓存，不会随访问次数增长。
 */
#define TC_KIND_COLL    0
#define TC_KIND_SERIES  1
#define TC_KIND_METHOD  2   /* 根上的方法对象（AddSeries 那族） */

typedef struct TcObj {
    const IDispatchVtbl* lpVtbl;
    LONG ref;
    int  kind;      /* TC_KIND_* */
    HWND hwnd;
    int  index;     /* 系列下标（0 基）；集合为 -1 */
    wchar_t mname[64];   /* TC_KIND_METHOD：要回放的方法名 */
} TcObj;

static const IDispatchVtbl tcObjVtbl;

/* 名表顺序 = DISPID（下标 + 1），与 memberobj 同一口径。
 * 只登记扁平层真做得动的成员：Marks / Pen / Brush / AsPie 这些放进来就是
 * “看着支持、实则答错”，宁可留 438（文档 §5 已登记）。 */
static const wchar_t* kTcCollNames[] = { L"Item", L"Count", NULL };
static const wchar_t* kTcSerNames[]  = {
    L"Add", L"AddXY", L"FillSampleValues", L"Clear",
    L"Color", L"Title", L"SeriesType", L"Count", NULL };

static void tcSetEmpty(VARIANT* o) { if (o) VariantInit(o); }
static void tcSetL(VARIANT* o, int32_t v) {
    if (!o) return; VariantInit(o); o->vt = VT_I4; o->lVal = v;
}
static void tcSetD(VARIANT* o, double v) {
    if (!o) return; VariantInit(o); o->vt = VT_R8; o->dblVal = v;
}
static void tcSetStr(VARIANT* o, const wchar_t* s) {
    if (!o) return; VariantInit(o); o->vt = VT_BSTR; o->bstrVal = SysAllocString(s ? s : L"");
}
static void tcSetObj(VARIANT* o, TcObj* p) {
    if (!o) return; VariantInit(o);
    if (!p) { o->vt = VT_EMPTY; return; }
    InterlockedIncrement(&p->ref);   /* 交出去的这枚引用由调用方 VariantClear 放 */
    o->vt = VT_DISPATCH; o->pdispVal = (IDispatch*)p;
}

/* 实参取值：VB6 递进来的可能是 I4/I2/R8/BSTR/ByRef，统一交给 VariantChangeType
 * 换算，不手抄第二张表（uc_hostmodel 那两张表就是手抄的产物）。 */
static int32_t tcArgL(VARIANT* v) {
    VARIANT t; int32_t r = 0;
    if (!v) return 0;
    VariantInit(&t);
    if (SUCCEEDED(VariantChangeType(&t, v, 0, VT_I4))) r = t.lVal;
    VariantClear(&t);
    return r;
}
static double tcArgD(VARIANT* v) {
    VARIANT t; double r = 0.0;
    if (!v) return 0.0;
    VariantInit(&t);
    if (SUCCEEDED(VariantChangeType(&t, v, 0, VT_R8))) r = t.dblVal;
    VariantClear(&t);
    return r;
}

static TcObj* tcObjFor(TcChart* c, int kind, int index);
static int32_t tcRootCall(TcChart* c, void* hwnd, const wchar_t* name, int argc,
                          VARIANT** a, VARIANT* out);
static TcObj* tcMethodFor(TcChart* c, const wchar_t* name);

static HRESULT STDMETHODCALLTYPE tcQI(IDispatch* d, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = d; InterlockedIncrement(&((TcObj*)d)->ref); return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE tcAddRef(IDispatch* d) {
    return (ULONG)InterlockedIncrement(&((TcObj*)d)->ref);
}
static ULONG STDMETHODCALLTYPE tcRelease(IDispatch* d) {
    LONG r = InterlockedDecrement(&((TcObj*)d)->ref);
    /* 不归还内存：对象归图表所有（见上面生命周期那段） */
    return (ULONG)(r > 0 ? r : 0);
}
static HRESULT STDMETHODCALLTYPE tcGtiCount(IDispatch* d, UINT* n) {
    (void)d; if (n) *n = 0; return S_OK;
}
static HRESULT STDMETHODCALLTYPE tcGti(IDispatch* d, UINT i, LCID l, ITypeInfo** t) {
    (void)d; (void)i; (void)l; if (t) *t = NULL; return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE tcGetIDs(IDispatch* d, REFIID riid, LPOLESTR* names,
                                          UINT n, LCID lcid, DISPID* ids) {
    TcObj* o = (TcObj*)d;
    const wchar_t* const* tbl = (o->kind == TC_KIND_COLL) ? kTcCollNames : kTcSerNames;
    (void)riid; (void)lcid;
    if (!names || !ids) return E_POINTER;
    for (UINT k = 0; k < n; k++) {
        int i;
        if (!names[k] || !names[k][0]) { ids[k] = DISPID_VALUE; continue; }
        for (i = 0; tbl[i]; i++) if (_wcsicmp(tbl[i], names[k]) == 0) break;
        if (tbl[i]) { ids[k] = (DISPID)(i + 1); continue; }
        if (o->kind == TC_KIND_COLL && _wcsicmp(names[k], L"_NewEnum") == 0) {
            ids[k] = DISPID_NEWENUM; continue;   /* Invoke 里答 E_NOTIMPL（缺口已登记） */
        }
        return DISP_E_UNKNOWNNAME;
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE tcInvoke(IDispatch* d, DISPID id, REFIID riid, LCID lcid,
                                          WORD flags, DISPPARAMS* dp, VARIANT* res,
                                          EXCEPINFO* ei, UINT* ae) {
    TcObj* o = (TcObj*)d;
    TcChart* c = tcOf(o->hwnd);
    TcSeries* s;
    int argc = dp ? (int)dp->cArgs : 0;
    /* DISPPARAMS 的 rgvarg 是**逆序**：第一个实参在 rgvarg[argc-1]。 */
    VARIANT* a0 = (argc >= 1) ? &dp->rgvarg[argc - 1] : NULL;
    VARIANT* a1 = (argc >= 2) ? &dp->rgvarg[argc - 2] : NULL;
    (void)riid; (void)lcid; (void)ei; (void)ae;
    if (!c) return DISP_E_MEMBERNOTFOUND;
    if (res) tcSetEmpty(res);

    /* TC_KIND_METHOD：cgen 把 `TChart1.AddSeries 5` 发成
     * vb6_ComGetObjectProp(hwnd,L"AddSeries") + vb6_ComCallObject(成员, NULL, 实参, 1)
     * —— 即方法名先被当成**对象**取一次，随后对那个对象走默认成员。
     * 所以 GetProp 要能交出“方法对象”，这里回放成根上的方法。 */
    if (o->kind == TC_KIND_METHOD) {
        VARIANT* av[8];
        int i;
        if (argc > 8) argc = 8;
        for (i = 0; i < argc; i++) av[i] = &dp->rgvarg[argc - 1 - i];
        if (!tcRootCall(c, o->hwnd, o->mname, argc, av, res)) return DISP_E_MEMBERNOTFOUND;
        return S_OK;
    }

    if (o->kind == TC_KIND_COLL) {
        switch (id) {
        case DISPID_VALUE: {                   /* 默认成员 = Item(i)，0 基（TeeChart 口径） */
            int idx = (int)tcArgL(a0);
            if (idx < 0) idx = 0;
            tcObjFor(c, TC_KIND_SERIES, idx);   /* 访问即建号，与 C 入口同一策略 */
            tcSetObj(res, tcObjFor(c, TC_KIND_SERIES, idx));
            return S_OK;
        }
        case 1: {                              /* Item(i) */
            int idx = (int)tcArgL(a0);
            if (idx < 0) idx = 0;
            tcObjFor(c, TC_KIND_SERIES, idx);
            tcSetObj(res, tcObjFor(c, TC_KIND_SERIES, idx));
            return S_OK;
        }
        case 2:                                /* Count */
            tcSetL(res, c->nser);
            return S_OK;
        case DISPID_NEWENUM:
            return E_NOTIMPL;                   /* For Each 还没接（文档 §5） */
        default:
            return DISP_E_MEMBERNOTFOUND;
        }
    }

    /* TC_KIND_SERIES */
    s = tcSeries(c, o->index);
    if (!s) return DISP_E_MEMBERNOTFOUND;
    switch (id) {
    case 1:                                     /* Add(y) */
        tcSetL(res, vb6_TChart_SeriesAdd(o->hwnd, o->index, tcArgD(a0)));
        return S_OK;
    case 2:                                     /* AddXY(x,y) */
        tcSetL(res, vb6_TChart_SeriesAddXY(o->hwnd, o->index, tcArgD(a0), tcArgD(a1)));
        return S_OK;
    case 3:                                     /* FillSampleValues(n) */
        tcSetL(res, vb6_TChart_FillSampleValues(o->hwnd, o->index, (int32_t)tcArgL(a0)));
        return S_OK;
    case 4:                                     /* Clear */
        vb6_TChart_SeriesClear(o->hwnd, o->index);
        return S_OK;
    case 5:                                     /* Color */
        if (flags & DISPATCH_PROPERTYPUT) {
            vb6_TChart_SetSeriesColor(o->hwnd, o->index, (int32_t)tcArgL(a0));
            return S_OK;
        }
        tcSetL(res, (int32_t)s->color);
        return S_OK;
    case 6:                                     /* Title */
        if (flags & DISPATCH_PROPERTYPUT) {
            BSTR b = NULL; VARIANT t;
            VariantInit(&t);
            if (a0 && SUCCEEDED(VariantChangeType(&t, a0, 0, VT_BSTR))) b = t.bstrVal;
            vb6_TChart_SetSeriesTitle(o->hwnd, o->index, b ? b : L"");
            VariantClear(&t);
            return S_OK;
        }
        tcSetStr(res, s->title);
        return S_OK;
    case 7:                                     /* SeriesType */
        if (flags & DISPATCH_PROPERTYPUT) { s->type = (int)tcArgL(a0); return S_OK; }
        tcSetL(res, s->type);
        return S_OK;
    case 8:                                     /* Count = 点数 */
        tcSetL(res, s->n);
        return S_OK;
    default:
        return DISP_E_MEMBERNOTFOUND;
    }
}

static const IDispatchVtbl tcObjVtbl = {
    tcQI, tcAddRef, tcRelease, tcGtiCount, tcGti, tcGetIDs, tcInvoke
};

static TcObj* tcObjFor(TcChart* c, int kind, int index) {
    int slot;
    if (!c) return NULL;
    slot = (kind == TC_KIND_COLL) ? 0 : (index + 1);
    if (slot < 0 || slot > VB6_TC_MAXSER) return NULL;
    if (!c->obj[slot]) {
        TcObj* o = (TcObj*)calloc(1, sizeof(TcObj));
        if (!o) return NULL;
        o->lpVtbl = &tcObjVtbl;
        o->ref    = 1;
        o->kind   = kind;
        o->hwnd   = c->hwnd;
        o->index  = index;
        c->obj[slot] = o;
    }
    return c->obj[slot];
}

/* 根上的方法名（GetProp 时交方法对象出去，Invoke 时回放） */
static const wchar_t* const kTcMethodNames[] = {
    L"AddSeries", L"Repaint", L"Refresh", L"Draw", L"RemoveAllSeries", NULL };

/* 根方法分派：Host_Call 与 TC_KIND_METHOD 两条入口共用这一张表。
 * 认不出返回 0（调用方回 438 / 继续往下走）。 */
static int32_t tcRootCall(TcChart* c, void* hwnd, const wchar_t* name, int argc,
                          VARIANT** a, VARIANT* out) {
    if (!c || !name) return 0;
    if (_wcsicmp(name, L"AddSeries") == 0) {
        int32_t idx = vb6_TChart_AddSeries(hwnd, (argc >= 1) ? tcArgL(a[0]) : 0);
        if (out) tcSetL(out, idx);
        return 1;
    }
    if (_wcsicmp(name, L"Repaint") == 0 || _wcsicmp(name, L"Refresh") == 0
        || _wcsicmp(name, L"Draw") == 0) {
        vb6_TChart_Repaint(hwnd);
        return 1;
    }
    if (_wcsicmp(name, L"RemoveAllSeries") == 0) {
        c->nser = 0;
        vb6_TChart_Repaint(hwnd);
        return 1;
    }
    return 0;
}

static TcObj* tcMethodFor(TcChart* c, const wchar_t* name) {
    int i;
    if (!c || !name) return NULL;
    for (i = 0; i < c->nmeth; i++) {
        if (c->meth[i] && _wcsicmp(c->meth[i]->mname, name) == 0) return c->meth[i];
    }
    if (c->nmeth >= (int)(sizeof(c->meth) / sizeof(c->meth[0]))) return NULL;
    {
        TcObj* o = (TcObj*)calloc(1, sizeof(TcObj));
        if (!o) return NULL;
        o->lpVtbl = &tcObjVtbl;
        o->ref    = 1;
        o->kind   = TC_KIND_METHOD;
        o->hwnd   = c->hwnd;
        o->index  = -1;
        wcsncpy(o->mname, name, 63);
        c->meth[c->nmeth++] = o;
        return o;
    }
}

/* ==================== 宿主模型钩子（vb6_Host_GetProp/Call/SetProp 的尾叉） ====================
 * 形状与 vb6_MenuProxyDecode 那一族一致：认不出返回 0，让宿主模型回到它原来的
 * report438 分支 —— 对别的控件行为逐字不变。
 */
int32_t vb6_TChart_HostGetProp(void* obj, const wchar_t* name, void* outV) {
    TcChart* c = tcOfClass((HWND)obj);
    VARIANT* out = (VARIANT*)outV;
    if (!c || !name || !out) return 0;
    if (_wcsicmp(name, L"Series") == 0) {
        tcSetObj(out, tcObjFor(c, TC_KIND_COLL, -1));
        return 1;
    }
    if (_wcsicmp(name, L"SeriesCount") == 0) { tcSetL(out, c->nser); return 1; }
    if (_wcsicmp(name, L"HeaderText") == 0 || _wcsicmp(name, L"Text") == 0) {
        tcSetStr(out, c->header); return 1;
    }
    if (_wcsicmp(name, L"ChartWidth") == 0) {
        RECT r; GetClientRect(c->hwnd, &r); tcSetL(out, r.right - r.left); return 1;
    }
    if (_wcsicmp(name, L"ChartHeight") == 0) {
        RECT r; GetClientRect(c->hwnd, &r); tcSetL(out, r.bottom - r.top); return 1;
    }
    /* 方法名当对象取（见 tcInvoke 里 TC_KIND_METHOD 那段） */
    {
        int i;
        for (i = 0; kTcMethodNames[i]; i++) {
            if (_wcsicmp(name, kTcMethodNames[i]) == 0) {
                tcSetObj(out, tcMethodFor(c, kTcMethodNames[i]));
                return 1;
            }
        }
    }
    return 0;
}

int32_t vb6_TChart_HostCall(void* obj, const wchar_t* name, int32_t argc, void** argv,
                            void* outV) {
    TcChart* c = tcOfClass((HWND)obj);
    VARIANT* av[8];
    int i;
    if (!c || !name) return 0;
    if (argc > 8) argc = 8;
    for (i = 0; i < argc; i++) av[i] = (argv && argv[i]) ? (VARIANT*)argv[i] : NULL;
    return tcRootCall(c, obj, name, argc, av, (VARIANT*)outV);
}

int32_t vb6_TChart_HostSetProp(void* obj, const wchar_t* name, const void* inV) {
    TcChart* c = tcOfClass((HWND)obj);
    const VARIANT* v = (const VARIANT*)inV;
    VARIANT t;
    if (!c || !name) return 0;
    if (_wcsicmp(name, L"HeaderText") == 0 && v) {
        VariantInit(&t);
        if (SUCCEEDED(VariantChangeType(&t, v, 0, VT_BSTR))) {
            vb6_TChart_SetHeaderText(obj, t.bstrVal ? t.bstrVal : L"");
        }
        VariantClear(&t);
        return 1;
    }
    return 0;
}
