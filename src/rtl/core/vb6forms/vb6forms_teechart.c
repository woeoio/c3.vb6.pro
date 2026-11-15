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
