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
    /* C30-10 第五刀：Marks / Pen / Brush / AsPie 的标量面 */
    int      marksVisible;
    COLORREF penColor;
    int      penWidth;
    int      penVisible;
    COLORREF brushColor;
    double   pieRotation;
    int      sliceExplode[64];   /* AsPie.ExplodedSlice.Value(i) */
    int      otherSliceStyle;    /* AsPie.OtherSlice.Style */
    double   otherSliceValue;
    int      otherSliceVisible;
} TcSeries;

typedef struct {
    HWND     hwnd;
    COLORREF panelColor;
    int      marginL, marginR, marginT, marginB;   /* 像素 */
    wchar_t  header[256];
    int      showLegend;
    /* C30-10 第十刀：Zoom 的两个值（空实现里也照实存，读回才一致） */
    int      zoomEnable, zoomPercent;
    /* C30-10 第十一刀：Aspect（3D）。深度只在 View3D 时非零 */
    int      view3D, perspective, aspRot, aspElev;
    /* C30-10 第十二刀：四壁 */
    int      wallsVisible;
    COLORREF wallsColor;
    TcSeries ser[VB6_TC_MAXSER];
    int      nser;
    int      paints;             /* 绘制次数（C3_TCHART_TRACE 用） */
    struct TcObj* obj[VB6_TC_MAXSER + 1];   /* 成员对象缓存 (0 = Series 集合) */
    struct TcObj* meth[8];                  /* 方法对象缓存（AddSeries 那几枚） */
    int      nmeth;
    struct TcObj* extra[16];                /* Panel / Header / Axes / Axis / Title */
    int      nextra;
    /* 对象模型的标量面（C30-10 第三刀）。轴下标: 0=Left 1=Right 2=Top 3=Bottom */
    int      panelBorder;
    int      headerVisible;
    int      axisShow[4];
    int      axisTitleShow[4];
    int      axisAuto[4];
    double   axisMin[4], axisMax[4];
    wchar_t  axisTitle[4][64];
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
    int i;
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
    c->zoomEnable    = 0;
    c->zoomPercent   = 100;
    c->view3D        = 1;    /* 与真控件默认一致（True）；受影响判据已按实测重录，见台账 §5 */
    c->perspective   = 1;
    c->aspRot        = 30;
    c->aspElev       = 15;
    c->wallsVisible  = 1;
    c->wallsColor    = RGB(192, 192, 192);
    c->panelBorder   = 1;
    c->headerVisible = 1;
    for (i = 0; i < 4; i++) { c->axisShow[i] = 1; c->axisTitleShow[i] = 1; c->axisAuto[i] = 1; }
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
        s->penColor = s->color;
        s->brushColor = s->color;
        s->penWidth = 1;
        s->penVisible = 1;
        s->marksVisible = 1;
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

/* 3D 柱的顶面/侧面：向白/黑方向拉一段（pct>0 变亮、<0 变暗） */
static COLORREF tcShade(COLORREF c, int pct) {
    int r = GetRValue(c), g = GetGValue(c), b = GetBValue(c);
    if (pct > 0) {
        r += (255 - r) * pct / 100;
        g += (255 - g) * pct / 100;
        b += (255 - b) * pct / 100;
    } else {
        r = r * (100 + pct) / 100;
        g = g * (100 + pct) / 100;
        b = b * (100 + pct) / 100;
    }
    return RGB(r, g, b);
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
    HFONT   fnt, fntOld;
    int i, n, x0, x1, y0, y1;
    int marksDrawn = 0;   /* C30-10 第六刀：本次真的画了多少个逐点标签 */
    int headerH = 0, legendW = 0, legendLeft = 0, legendRight = 0;
    double vmin, vmax;

    if (!c) return;
    dc = BeginPaint(c->hwnd, &ps);
    if (!dc) return;
    GetClientRect(c->hwnd, &rc);

    /* 字体：TeeChart 的图例/标题用的是自己的字体（不是系统 GUI 字体）。
     * 直接沿用默认字体量出来的“Series0”宽 53px，而参照读数的图例框只有 61 宽
     * （“TeeChart display demo” 只占 129px），对不上 —— 换 Arial 8pt 后两项都对。 */
    fnt    = CreateFontW(-11, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET,
                         OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY,
                         VARIABLE_PITCH | FF_SWISS, L"Arial");
    fntOld = (HFONT)SelectObject(dc, fnt);

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

    /* ================= C30-10 第七刀：面板分带 =================
     * 真 TeeChart 不是“客户区减值当绘图区”，而是把图表分成几条带：
     *   header 带 0..26（标题文字垂直居中在这一带）、上轴带 26..53、
     *   绘图区 53..(H-37)、下轴带 25px + MarginBottom、图例在右侧帮忙挤绘图区。
     * 这些数就是 401x257 参照读数反推出来的（上轴 y=53、下轴 y=220、
     *   图例 x=310..371 y=26、标题 bbox 中心 155.5）。
     */
    if (c->header[0]) headerH = 26;
    if (c->showLegend && c->nser > 0) {
        int tw = 0, k;
        for (k = 0; k < c->nser; k++) {
            SIZE sz;
            int w = 0;
            if (c->ser[k].title[0] &&
                GetTextExtentPoint32W(dc, c->ser[k].title, (int)wcslen(c->ser[k].title), &sz))
                w = sz.cx;
            if (w > tw) tw = w;
        }
        legendW     = tw + 24;   /* 色块 10 + 缝 3 + 文字 + 右衬边（宽 61 时与参照图例 310..371 对齐） */
        legendRight = rc.right - c->marginR - 18;
        legendLeft  = legendRight - legendW;
        if (legendLeft < rc.left + 60) legendLeft = rc.left + 60;   /* 太窄时给绘图区留地面 */
    }

    plot = rc;
    plot.left   += c->marginL + 16;                       /* 左轴刻度文字位 */
    plot.top    = rc.top + headerH + 27;                  /* header 带 + 上轴带 */
    plot.right  = (legendW > 0) ? (legendLeft - 8) : (rc.right - c->marginR);
    plot.bottom = rc.bottom - 25 - c->marginB;            /* 下轴带 + MarginBottom */
    if (plot.right <= plot.left + 2 || plot.bottom <= plot.top + 2) {
        SelectObject(dc, fntOld); DeleteObject(fnt);
        SelectObject(dc, oldPen); DeleteObject(pen);
        EndPaint(c->hwnd, &ps);
        return;
    }

    /* 四壁：3D 下才画，且必须在**轴线之前**（墙是背板，不是覆盖层）。
     * 深度与柱子同一个常量，判据才能拿“墙左缘 plot.left、背板右缘 plot.right+d”当尺子。 */
    if (c->view3D && c->wallsVisible) {
        int d = (c->perspective ? 18 : 12);
        POINT back[4], lwall[4];
        HBRUSH wb = CreateSolidBrush(c->wallsColor);
        HGDIOBJ wOld = SelectObject(dc, wb);
        HGDIOBJ wpOld = SelectObject(dc, GetStockObject(NULL_PEN));
        back[0].x = plot.left + d;  back[0].y = plot.top - d;
        back[1].x = plot.right + d; back[1].y = plot.top - d;
        back[2].x = plot.right + d; back[2].y = plot.bottom - d;
        back[3].x = plot.left + d;  back[3].y = plot.bottom - d;
        lwall[0].x = plot.left;     lwall[0].y = plot.top;
        lwall[1].x = plot.left + d; lwall[1].y = plot.top - d;
        lwall[2].x = plot.left + d; lwall[2].y = plot.bottom - d;
        lwall[3].x = plot.left;     lwall[3].y = plot.bottom;
        Polygon(dc, back, 4);
        Polygon(dc, lwall, 4);
        SelectObject(dc, wpOld);
        SelectObject(dc, wOld);
        DeleteObject(wb);
    }

    /* Header：横向居中于“客户区 - 图例”那段（参照读数 centerX=155.5 = 图例左边/2） */
    if (c->header[0]) {
        RECT hr;
        hr.left   = rc.left;
        hr.right  = (legendW > 0) ? legendLeft : rc.right;
        hr.top    = rc.top + 12;
        hr.bottom = rc.top + 26;
        tcDrawText(dc, c->header, &hr, 1, RGB(0, 0, 0));
    }

    /* 上/下轴线：TeeChart 默认画在绘图区上下沿，且是**黑色**（不是面板边的灰） */
    {
        HPEN ap = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
        HPEN apOld = (HPEN)SelectObject(dc, ap);
        MoveToEx(dc, plot.left, plot.top, NULL);
        LineTo(dc, plot.right, plot.top);
        MoveToEx(dc, plot.left, plot.bottom, NULL);
        LineTo(dc, plot.right, plot.bottom);
        SelectObject(dc, apOld);
        DeleteObject(ap);
    }

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
    if (vmax <= vmin) {
        /* 全等值（含全零）：vmin 归 0、vmax 保留正值 —— 之前直接给 0..1，
         * 结果是“值 20 在 0..1 的轴上”把柱子画出绘图区（墙夹具第一版就这么红的）。 */
        vmin = 0;
        if (vmax <= 0) vmax = 1;
    }

    x0 = plot.left; x1 = plot.right;
    y0 = plot.top;  y1 = plot.bottom;

    /* 系列 */
    for (i = 0; i < c->nser; i++) {
        TcSeries* s = &c->ser[i];
        int k;
        COLORREF col = (s->color ? s->color : tcPalette(i));
        /* C30-10 第六刀：真正的笔色/笔宽来自 Pen，填充色来自 Brush。
         * Pen.Visible=False 就退回系列色（TeeChart 里那是“不描边≠不画”）。 */
        COLORREF pcol = (s->penVisible ? (s->penColor ? s->penColor : col) : col);
        int      pw   = (s->penWidth > 0 ? s->penWidth : 1);
        COLORREF fcol = (s->brushColor ? s->brushColor : col);
        HPEN sp = CreatePen(PS_SOLID, pw, pcol);
        HPEN spOld = (HPEN)SelectObject(dc, sp);
        HBRUSH sb = CreateSolidBrush(fcol);
        HGDIOBJ sbOld = SelectObject(dc, sb);

        if (s->type == TC_SC_PIE) {
            /* 饼图：按 y 值占比切扇区（起始 -90°，与 TeeChart 默认同向）。
             * Pie() 的两个“终点”各自是扇区的起/止半径端点 —— 角度算错就会画出
             * 整圆或零面积，所以这里统一用 cos/sin 落在圆周上。
             *
             * C30-10 第八刀：AsPie 的两个值真的画出来——
             *   RotationAngle 整体旋起始角；ExplodedSlice.Value(i) 把第 i 个扇区
             *   沿它自己的中线往外掰这么多**像素**（TeeChart 的口径）。 */
            double sum = 0, ang = -90.0 + s->pieRotation;
            int cx0 = (x0 + x1) / 2, cy0 = (y0 + y1) / 2;
            int r  = ((x1 - x0) < (y1 - y0) ? (x1 - x0) : (y1 - y0)) / 2 - 2;
            if (r < 4) r = 4;
            for (k = 0; k < s->n; k++) sum += (s->pt[k].y > 0 ? s->pt[k].y : 0);
            if (sum <= 0) sum = 1;
            for (k = 0; k < s->n; k++) {
                double sweep = (s->pt[k].y > 0 ? s->pt[k].y : 0) / sum * 360.0;
                double amid  = (ang + sweep / 2.0) * 3.14159265358979323846 / 180.0;
                double a0 = ang * 3.14159265358979323846 / 180.0;
                double a1 = (ang + sweep) * 3.14159265358979323846 / 180.0;
                int exp = (k < 64) ? s->sliceExplode[k] : 0;
                int cx = cx0 + (int)(exp * cos(amid));
                int cy = cy0 + (int)(exp * sin(amid));
                int sx = cx + (int)(r * cos(a0)), sy = cy + (int)(r * sin(a0));
                int ex = cx + (int)(r * cos(a1)), ey = cy + (int)(r * sin(a1));
                HBRUSH pb = CreateSolidBrush(tcPalette(k));
                HGDIOBJ pbOld = SelectObject(dc, pb);
                HGDIOBJ ob = SelectObject(dc, GetStockObject(NULL_PEN));
                if (sweep >= 359.9) {
                    Ellipse(dc, cx - r, cy - r, cx + r, cy + r);
                } else if (sweep > 0.01) {
                    /* 端点顺序得反着给：GDI 的 Pie 默认按 AD_COUNTERCLOCKWISE 从
                     * 第一个端点扫到第二个，而 y 轴向下 ⇒ 视觉上是顺时针。
                     * 顺着给 (sx,sy)->(ex,ey) 会填出 270° 的补集（症状：最后一块
                     * 颜色吃掉大半个圆，前面的颜色只剩镶边）。 */
                    Pie(dc, cx - r, cy - r, cx + r, cy + r, ex, ey, sx, sy);
                }
                SelectObject(dc, ob);
                SelectObject(dc, pbOld);
                DeleteObject(pb);
                if (getenv("TC_PIE_DEBUG")) {
                    fprintf(stderr, "[TCHART-PIE] k=%d ang=%g sweep=%g pal=%d expl=%d c=%d,%d\n",
                            k, ang, sweep, k & 7, exp, cx, cy);
                }
                ang += sweep;
            }
            /* C30-10 第九刀：饼图的 Marks 是**扇区标签**（不是逐点数值）。
             * 位置取各块中线 0.65r 处，文字用黑色（用系列色会在同色扇区上看不见）。 */
            if (s->marksVisible) {
                double a = -90.0 + s->pieRotation;
                SetBkMode(dc, TRANSPARENT);
                SetTextColor(dc, RGB(0, 0, 0));
                for (k = 0; k < s->n; k++) {
                    double sweep = (s->pt[k].y > 0 ? s->pt[k].y : 0) / sum * 360.0;
                    double amid  = (a + sweep / 2.0) * 3.14159265358979323846 / 180.0;
                    int exp = (k < 64) ? s->sliceExplode[k] : 0;
                    wchar_t lb[32];
                    int lx = (x0 + x1) / 2 + (int)((exp + r * 0.65) * cos(amid));
                    int ly = (y0 + y1) / 2 + (int)((exp + r * 0.65) * sin(amid));
                    int lw;
                    wsprintfW(lb, L"%d", (int)(s->pt[k].y + (s->pt[k].y >= 0 ? 0.5 : -0.5)));
                    lw = (int)wcslen(lb) * 7;
                    TextOutW(dc, lx - lw / 2, ly - 6, lb, (int)wcslen(lb));
                    marksDrawn++;
                    a += sweep;
                }
            }
        } else if (s->n > 0) {
            int bw, gap, k2;
            switch (s->type) {
            case TC_SC_BAR: {
                /* C30-10 第十一刀：View3D 时按 12px 深度画出顶面 + 右侧面（先画背面，
                 * 再把正面矩形压在上面）。深度的选择定成常量，判据才能拿它当尺子量。 */
                int depth = c->view3D ? 12 : 0;
                /* 透视：越靠近绘图区左侧（远端）深度越大，右侧（近端）越小，
                 * 整幅形成 16→8 的梯度（差 8px，判据量得出来）。
                 * 值映射必须按**最大**深度让位，否则最左边那根柱子的顶面又会
                 * 伸出上轴线（上一刀刚修过同一个毛病）。 */
                int maxDepth = depth + ((c->view3D && c->perspective) ? 6 : 0);
                int bh = (y1 - y0) - maxDepth;
                if (bh < 8) bh = 8;
                bw  = (s->n > 0) ? ((x1 - x0) / (s->n * 2)) : 8;
                gap = bw;
                if (bw < 2) bw = 2;
                for (k2 = 0; k2 < s->n; k2++) {
                    int bx = x0 + gap / 2 + k2 * (bw + gap);
                    int by = y1 - (int)((s->pt[k2].y - vmin) / (vmax - vmin) * bh);
                    if (depth > 0 && c->perspective) {
                        int cxm = (x0 + x1) / 2;
                        double t = (double)(bx + bw / 2 - cxm) / ((x1 - x0) / 2.0);
                        depth = 12 - (int)(6.0 * t);   /* 左端约 16、右端约 9 */
                        if (depth < 4) depth = 4;
                        if (depth > 20) depth = 20;
                    }
                    /* 值超出量程（含全等值退化）时不让柱顶爬出绘图区 */
                    if (by < y0 + depth) by = y0 + depth;
                    if (depth > 0) {
                        POINT top[4], side[4];
                        HBRUSH tb = CreateSolidBrush(tcShade(fcol, 55));
                        HBRUSH db = CreateSolidBrush(tcShade(fcol, -45));
                        HGDIOBJ ob = SelectObject(dc, GetStockObject(NULL_PEN));
                        side[0].x = bx + bw;         side[0].y = by;
                        side[1].x = bx + bw + depth; side[1].y = by - depth;
                        side[2].x = bx + bw + depth; side[2].y = y1 - depth;
                        side[3].x = bx + bw;         side[3].y = y1;
                        top[0].x = bx;               top[0].y = by;
                        top[1].x = bx + depth;       top[1].y = by - depth;
                        top[2].x = bx + bw + depth;  top[2].y = by - depth;
                        top[3].x = bx + bw;          top[3].y = by;
                        SelectObject(dc, db);
                        Polygon(dc, side, 4);
                        SelectObject(dc, tb);
                        Polygon(dc, top, 4);
                        SelectObject(dc, ob);
                        DeleteObject(tb);
                        DeleteObject(db);
                        SelectObject(dc, sb);   /* 正面仍用系列/画刷色 */
                    }
                    Rectangle(dc, bx, by, bx + bw, y1);
                }
                break;
            }
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
                    /* C30-10 第十三刀：3D 下的面积把整块向后上平移一遭（暗色）当“带”的后闸，
                     * 再画前闸 —— 前闸压后闸，看起来就是一条有厚度的带。 */
                    int d = c->view3D ? 12 : 0;
                    POINT* poly = (POINT*)malloc(sizeof(POINT) * (s->n + 2));
                    if (poly) {
                        for (k2 = 0; k2 < s->n; k2++) {
                            poly[k2].x = (s->n == 1) ? (x0 + x1) / 2
                                       : x0 + (int)((double)k2 / (s->n - 1) * (x1 - x0));
                            poly[k2].y = y1 - (int)((s->pt[k2].y - vmin) / (vmax - vmin) * (y1 - y0));
                        }
                        poly[s->n].x = poly[s->n - 1].x; poly[s->n].y = y1;
                        poly[s->n + 1].x = poly[0].x;    poly[s->n + 1].y = y1;
                        if (d > 0) {
                            HBRUSH rb = CreateSolidBrush(tcShade(fcol, -45));
                            HGDIOBJ rbOld = SelectObject(dc, rb);
                            HGDIOBJ rpOld = SelectObject(dc, GetStockObject(NULL_PEN));
                            for (k2 = 0; k2 <= s->n + 1; k2++) { poly[k2].x += d; poly[k2].y -= d; }
                            Polygon(dc, poly, s->n + 2);
                            for (k2 = 0; k2 <= s->n + 1; k2++) { poly[k2].x -= d; poly[k2].y += d; }
                            SelectObject(dc, rpOld);
                            SelectObject(dc, rbOld);
                            DeleteObject(rb);
                        }
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
        /* C30-10 第六刀：Marks.Visible=True 时逐点打值（饼图除外：那是扇区标签，
         * 口径和图例不一样，留到下一刀）。点多于 64 不打，避免糊成一片。 */
        if (s->marksVisible && s->type != TC_SC_PIE && s->n > 0 && s->n <= 64) {
            int bw2 = 0;
            if (s->type == TC_SC_BAR) {
                bw2 = (x1 - x0) / (s->n * 2);
                if (bw2 < 2) bw2 = 2;
            }
            SetBkMode(dc, TRANSPARENT);
            SetTextColor(dc, pcol);
            for (k = 0; k < s->n; k++) {
                wchar_t mb[32];
                int px, py;
                if (s->type == TC_SC_BAR) px = x0 + bw2 / 2 + k * 2 * bw2 + bw2 / 2;
                else px = (s->n == 1) ? (x0 + x1) / 2
                        : x0 + (int)((double)k / (s->n - 1) * (x1 - x0));
                py = y1 - (int)((s->pt[k].y - vmin) / (vmax - vmin) * (y1 - y0));
                wsprintfW(mb, L"%d", (int)(s->pt[k].y + (s->pt[k].y >= 0 ? 0.5 : -0.5)));
                TextOutW(dc, px - 4, py - 15, mb, (int)wcslen(mb));
                marksDrawn++;
            }
        }
        SelectObject(dc, spOld); DeleteObject(sp);
        SelectObject(dc, sbOld); DeleteObject(sb);
    }

    /* Legend：右侧一框，顶边贴在 header 带下面（参照读数 y=26、x=310..371） */
    if (legendW > 0) {
        int lh, k;
        RECT lr;
        HGDIOBJ obOld;
        lh = c->nser * 14 + 6;
        lr.left   = legendLeft;
        lr.right  = legendRight;
        lr.top    = rc.top + headerH;
        lr.bottom = lr.top + lh;
        br = CreateSolidBrush(c->panelColor);
        obOld = SelectObject(dc, br);
        FillRect(dc, &lr, br);
        SelectObject(dc, obOld);
        DeleteObject(br);
        SelectObject(dc, GetStockObject(NULL_BRUSH));
        {
            HPEN lp = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
            HPEN lpOld = (HPEN)SelectObject(dc, lp);
            Rectangle(dc, lr.left, lr.top, lr.right, lr.bottom);
            SelectObject(dc, lpOld);
            DeleteObject(lp);
        }
        SelectObject(dc, oldBr);
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
    SelectObject(dc, fntOld);
    DeleteObject(fnt);
    c->paints++;
    if (tcTrace()) {
        COLORREF trp = 0, trb = 0;
        double pieRot = 0;
        int pieExp0 = 0;
        if (c->nser > 0) {
            TcSeries* s0 = &c->ser[0];
            COLORREF c0 = (s0->color ? s0->color : tcPalette(0));
            trp = (s0->penVisible ? (s0->penColor ? s0->penColor : c0) : c0);
            trb = (s0->brushColor ? s0->brushColor : c0);
        }
        for (i = 0; i < c->nser; i++) {
            if (c->ser[i].type == TC_SC_PIE) {      /* 第一张饼图的旋转/掰开量 */
                pieRot  = c->ser[i].pieRotation;
                pieExp0 = c->ser[i].sliceExplode[0];
                break;
            }
        }
        fprintf(stderr, "[TCHART] paint=%d series=%d pen=%06X brush=%06X marks=%d pie=%g/%d hwnd=%p\n",
                c->paints, c->nser, (unsigned)trp, (unsigned)trb,
                marksDrawn, pieRot, pieExp0, (void*)c->hwnd);
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
#define TC_KIND_PANEL   3
#define TC_KIND_HEADER  4
#define TC_KIND_AXES    5   /* Axis 集合（Axis.Bottom 这类具名也走它） */
#define TC_KIND_AXIS    6   /* index = 0 Left / 1 Right / 2 Top / 3 Bottom */
#define TC_KIND_TITLE   7   /* 轴标题（index = 轴下标） */
#define TC_KIND_MARKS   8
#define TC_KIND_PEN     9
#define TC_KIND_BRUSH   10
#define TC_KIND_ASPIE   11
#define TC_KIND_OTHER   12  /* AsPie.OtherSlice */
#define TC_KIND_SLICES  13  /* AsPie.ExplodedSlice（带参 Value(i)） */
#define TC_KIND_PRINTER 14  /* 显式空实现：不打印，但不再 438 */
#define TC_KIND_EXPORT  15
#define TC_KIND_ZOOM    16
#define TC_KIND_ASPECT  17
#define TC_KIND_WALLS   18

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
    L"Color", L"Title", L"SeriesType", L"Count",
    L"Marks", L"Pen", L"Brush", L"AsPie", NULL };
/* C30-10 第三刀：Panel / Header / Axis 三层。只登记真做得动的名字。 */
static const wchar_t* kTcPanelNames[] = {
    L"Color", L"MarginLeft", L"MarginRight", L"MarginTop", L"MarginBottom",
    L"Visible", L"BorderVisible", NULL };
static const wchar_t* kTcHeaderNames[] = { L"Text", L"Caption", L"Visible", NULL };
/* Axis.Bottom / Left / Right / Top 是具名属性（DISPID 3..6），Item(i) 同集合 */
static const wchar_t* kTcAxesNames[] = {
    L"Item", L"Count", L"Bottom", L"Left", L"Right", L"Top", NULL };
static const wchar_t* kTcAxisNames[] = {
    L"Automatic", L"Minimum", L"Maximum", L"Visible", L"Title", NULL };
static const wchar_t* kTcTitleNames[] = { L"Caption", L"Text", L"Visible", NULL };
/* C30-10 第五刀：示例里最常见的四层（Series(i).Marks/Pen/Brush/AsPie） */
static const wchar_t* kTcMarksNames[]  = { L"Visible", L"ArrowLength", L"TextStyle", NULL };
static const wchar_t* kTcPenNames[]    = { L"Color", L"Width", L"Visible", NULL };
static const wchar_t* kTcBrushNames[]  = { L"Color", L"Visible", NULL };
static const wchar_t* kTcAsPieNames[]  = { L"RotationAngle", L"ExplodedSlice", L"OtherSlice", NULL };
static const wchar_t* kTcOtherNames[]  = { L"Style", L"Value", L"Visible", NULL };
static const wchar_t* kTcSlicesNames[] = { L"Value", NULL };
/* C30-10 第十刀：文档 §5 那四条未接成员 → 显式空实现（认名字、不动作） */
static const wchar_t* kTcPrinterNames[] = {
    L"BeginDoc", L"EndDoc", L"Print", L"PrintPartial",
    L"PageWidth", L"PageHeight", L"PrintOrientation", L"Copies", NULL };
static const wchar_t* kTcExportNames[] = {
    L"SaveToFile", L"CopyToClipboard", L"Email", L"Picture", L"Width", L"Height", NULL };
static const wchar_t* kTcZoomNames[] = {
    L"Enable", L"Percent", L"Direction", L"Point", L"SetMax", NULL };
/* C30-10 第十一刀：3D 面（先让值存得住、读得回，再把柱形按深度画出来） */
static const wchar_t* kTcAspectNames[] = {
    L"View3D", L"Perspective", L"Rotation", L"Elevation", L"Orthogonal", NULL };
/* C30-10 第十二刀：四壁。Back/Bottom/Right 先指回同一对象（共用可见性/颜色），
 * 分面颜色留到有人真的需要时再拆 —— 现在拆了也画不出区别。 */
static const wchar_t* kTcWallsNames[] = {
    L"Visible", L"Color", L"Back", L"Bottom", L"Right", NULL };

static const wchar_t* const* tcNames(int kind) {
    switch (kind) {
    case TC_KIND_COLL:   return kTcCollNames;
    case TC_KIND_SERIES: return kTcSerNames;
    case TC_KIND_PANEL:  return kTcPanelNames;
    case TC_KIND_HEADER: return kTcHeaderNames;
    case TC_KIND_AXES:   return kTcAxesNames;
    case TC_KIND_AXIS:   return kTcAxisNames;
    case TC_KIND_TITLE:  return kTcTitleNames;
    case TC_KIND_MARKS:  return kTcMarksNames;
    case TC_KIND_PEN:    return kTcPenNames;
    case TC_KIND_BRUSH:  return kTcBrushNames;
    case TC_KIND_ASPIE:  return kTcAsPieNames;
    case TC_KIND_OTHER:  return kTcOtherNames;
    case TC_KIND_SLICES: return kTcSlicesNames;
    case TC_KIND_PRINTER: return kTcPrinterNames;
    case TC_KIND_EXPORT:  return kTcExportNames;
    case TC_KIND_ZOOM:    return kTcZoomNames;
    case TC_KIND_ASPECT:  return kTcAspectNames;
    case TC_KIND_WALLS:   return kTcWallsNames;
    default:             return NULL;
    }
}

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
static TcObj* tcCached(TcChart* c, int kind, int index);
static int32_t tcRootCall(TcChart* c, void* hwnd, const wchar_t* name, int argc,
                          VARIANT** a, VARIANT* out);
static TcObj* tcMethodFor(TcChart* c, const wchar_t* name);
static void* tcEnumNew(TcChart* c);

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
    const wchar_t* const* all = tcNames(o->kind);
    if (all) tbl = all;
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
            /* For Each s In TChart1.Series —— 交一枚 IEnumVARIANT（C30-10 P1 第四刀）。
             * 与 memberobj 那族同一口径：枚举器每次 Next 现建一枚系列对象。 */
            tcSetObj(res, (TcObj*)tcEnumNew(c));
            return S_OK;
        default:
            return DISP_E_MEMBERNOTFOUND;
        }
    }

    /* ===== C30-10 第三刀：Panel / Header / Axis(+Title) =====
     * 这四层是像素判据的前置（真 VB6 读数里标题位置、图例边、轴线位置全由它们决定），
     * 所以先把 MakeGet/MakePut 做成能读能写，绘制跟上在下一刀。 */
    if (o->kind == TC_KIND_PANEL) {
        switch (id) {
        case 1: if (flags & DISPATCH_PROPERTYPUT) { c->panelColor = (COLORREF)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, (int32_t)c->panelColor); return S_OK;
        case 2: if (flags & DISPATCH_PROPERTYPUT) { c->marginL = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, c->marginL); return S_OK;
        case 3: if (flags & DISPATCH_PROPERTYPUT) { c->marginR = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, c->marginR); return S_OK;
        case 4: if (flags & DISPATCH_PROPERTYPUT) { c->marginT = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, c->marginT); return S_OK;
        case 5: if (flags & DISPATCH_PROPERTYPUT) { c->marginB = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, c->marginB); return S_OK;
        case 6: if (flags & DISPATCH_PROPERTYPUT) return S_OK;   /* Visible: 控件本身可见性归宿主 */
                tcSetL(res, 1); return S_OK;
        case 7: if (flags & DISPATCH_PROPERTYPUT) { c->panelBorder = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, c->panelBorder); return S_OK;
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }

    if (o->kind == TC_KIND_HEADER) {
        switch (id) {
        case 1: case 2:   /* Text / Caption */
            if (flags & DISPATCH_PROPERTYPUT) {
                VARIANT t; VariantInit(&t);
                if (a0 && SUCCEEDED(VariantChangeType(&t, a0, 0, VT_BSTR)))
                    vb6_TChart_SetHeaderText(c->hwnd, t.bstrVal ? t.bstrVal : L"");
                VariantClear(&t);
                return S_OK;
            }
            tcSetStr(res, c->header);
            return S_OK;
        case 3:   /* Visible */
            if (flags & DISPATCH_PROPERTYPUT) { c->headerVisible = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            tcSetL(res, c->headerVisible);
            return S_OK;
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }

    if (o->kind == TC_KIND_AXES) {
        switch (id) {
        case DISPID_VALUE: case 1: {    /* Item(i)：0..3 = Left/Right/Top/Bottom */
            int idx = (int)tcArgL(a0);
            if (idx < 0) idx = 0;
            if (idx > 3) idx = 3;
            tcSetObj(res, tcCached(c, TC_KIND_AXIS, idx));
            return S_OK;
        }
        case 2: tcSetL(res, 4); return S_OK;                 /* Count */
        case 3: tcSetObj(res, tcCached(c, TC_KIND_AXIS, 3)); return S_OK;   /* Bottom */
        case 4: tcSetObj(res, tcCached(c, TC_KIND_AXIS, 0)); return S_OK;   /* Left */
        case 5: tcSetObj(res, tcCached(c, TC_KIND_AXIS, 1)); return S_OK;   /* Right */
        case 6: tcSetObj(res, tcCached(c, TC_KIND_AXIS, 2)); return S_OK;   /* Top */
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }

    if (o->kind == TC_KIND_AXIS) {
        int ax = o->index;
        if (ax < 0 || ax > 3) ax = 3;
        switch (id) {
        case 1: if (flags & DISPATCH_PROPERTYPUT) { c->axisAuto[ax] = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, c->axisAuto[ax]); return S_OK;
        case 2: if (flags & DISPATCH_PROPERTYPUT) { c->axisMin[ax] = tcArgD(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetD(res, c->axisMin[ax]); return S_OK;
        case 3: if (flags & DISPATCH_PROPERTYPUT) { c->axisMax[ax] = tcArgD(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetD(res, c->axisMax[ax]); return S_OK;
        case 4: if (flags & DISPATCH_PROPERTYPUT) { c->axisShow[ax] = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, c->axisShow[ax]); return S_OK;
        case 5: tcSetObj(res, tcCached(c, TC_KIND_TITLE, ax)); return S_OK;   /* Title */
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }

    if (o->kind == TC_KIND_TITLE) {
        int ax = o->index;
        if (ax < 0 || ax > 3) ax = 3;
        switch (id) {
        case 1: case 2:   /* Caption / Text */
            if (flags & DISPATCH_PROPERTYPUT) {
                VARIANT t; VariantInit(&t);
                if (a0 && SUCCEEDED(VariantChangeType(&t, a0, 0, VT_BSTR))) {
                    wcsncpy(c->axisTitle[ax], t.bstrVal ? t.bstrVal : L"", 63);
                    c->axisTitle[ax][63] = 0;
                    InvalidateRect(c->hwnd, NULL, FALSE);
                }
                VariantClear(&t);
                return S_OK;
            }
            tcSetStr(res, c->axisTitle[ax]);
            return S_OK;
        case 3:   /* Visible */
            if (flags & DISPATCH_PROPERTYPUT) { c->axisTitleShow[ax] = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            tcSetL(res, c->axisTitleShow[ax]);
            return S_OK;
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }

    /* ===== 第十二刀：四壁（只 3D 下画）===== */
    if (o->kind == TC_KIND_WALLS) {
        switch (id) {
        case 1:
            if (flags & DISPATCH_PROPERTYPUT) { c->wallsVisible = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, c->wallsVisible);
            return S_OK;
        case 2:
            if (flags & DISPATCH_PROPERTYPUT) { c->wallsColor = (COLORREF)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, (int32_t)c->wallsColor);
            return S_OK;
        case 3: case 4: case 5:
            /* Back / Bottom / Right：暂时指回同一个对象（共用可见性/颜色） */
            if (res) tcSetObj(res, o);
            return S_OK;
        default:
            return DISP_E_MEMBERNOTFOUND;
        }
    }

    /* ===== 第十一刀：Aspect（3D 的开关与角度） ===== */
    if (o->kind == TC_KIND_ASPECT) {
        int* pv = NULL;
        switch (id) {
        case 1: pv = &c->view3D; break;
        case 2: pv = &c->perspective; break;
        case 3: pv = &c->aspRot; break;
        case 4: pv = &c->aspElev; break;
        case 5: if (res) tcSetL(res, 0); return S_OK;   /* Orthogonal：存不住，给 0 */
        default: return DISP_E_MEMBERNOTFOUND;
        }
        if (flags & DISPATCH_PROPERTYPUT) {
            *pv = (int)tcArgL(a0);
            InvalidateRect(c->hwnd, NULL, FALSE);
            return S_OK;
        }
        if (res) tcSetL(res, *pv);
        return S_OK;
    }

    /* ===== 第十刀：Printer / Export / Zoom —— 认名字、不动作，不静默 438 ===== */
    if (o->kind == TC_KIND_PRINTER || o->kind == TC_KIND_EXPORT || o->kind == TC_KIND_ZOOM) {
        if (o->kind == TC_KIND_PRINTER) {
            if (id == 5) { if (res) tcSetL(res, 793); return S_OK; }   /* PageWidth (A4@96dpi) */
            if (id == 6) { if (res) tcSetL(res, 1122); return S_OK; }  /* PageHeight */
            return S_OK;                                              /* 其余不动作 */
        }
        if (o->kind == TC_KIND_EXPORT) {
            return S_OK;   /* SaveToFile / CopyToClipboard / Email 不动作 */
        }
        switch (id) {
        case 1:            /* Zoom.Enable */
            if (flags & DISPATCH_PROPERTYPUT) { c->zoomEnable = (int)tcArgL(a0); return S_OK; }
            if (res) tcSetL(res, c->zoomEnable);
            return S_OK;
        case 2:            /* Zoom.Percent */
            if (flags & DISPATCH_PROPERTYPUT) { c->zoomPercent = (int)tcArgL(a0); return S_OK; }
            if (res) tcSetL(res, c->zoomPercent);
            return S_OK;
        default:
            return S_OK;
        }
    }

    /* ===== 第五刀：Marks / Pen / Brush / AsPie / OtherSlice / ExplodedSlice ===== */
    if (o->kind == TC_KIND_MARKS || o->kind == TC_KIND_PEN || o->kind == TC_KIND_BRUSH
        || o->kind == TC_KIND_ASPIE || o->kind == TC_KIND_OTHER || o->kind == TC_KIND_SLICES) {
        TcSeries* sr = tcSeries(c, o->index);
        if (!sr) return DISP_E_MEMBERNOTFOUND;
        if (o->kind == TC_KIND_MARKS) {
            switch (id) {
            case 1:
                if (flags & DISPATCH_PROPERTYPUT) { sr->marksVisible = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, sr->marksVisible); return S_OK;
            case 2: tcSetL(res, 8); return S_OK;    /* ArrowLength：认名字，不存值 */
            case 3: tcSetL(res, 0); return S_OK;    /* TextStyle：同上 */
            default: return DISP_E_MEMBERNOTFOUND;
            }
        }
        if (o->kind == TC_KIND_PEN) {
            switch (id) {
            case 1:
                if (flags & DISPATCH_PROPERTYPUT) { sr->penColor = (COLORREF)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, (int32_t)sr->penColor); return S_OK;
            case 2:
                if (flags & DISPATCH_PROPERTYPUT) { sr->penWidth = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, sr->penWidth); return S_OK;
            case 3:
                if (flags & DISPATCH_PROPERTYPUT) { sr->penVisible = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, sr->penVisible); return S_OK;
            default: return DISP_E_MEMBERNOTFOUND;
            }
        }
        if (o->kind == TC_KIND_BRUSH) {
            switch (id) {
            case 1:
                if (flags & DISPATCH_PROPERTYPUT) { sr->brushColor = (COLORREF)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, (int32_t)sr->brushColor); return S_OK;
            case 2: tcSetL(res, 1); return S_OK;
            default: return DISP_E_MEMBERNOTFOUND;
            }
        }
        if (o->kind == TC_KIND_ASPIE) {
            switch (id) {
            case 1:
                if (flags & DISPATCH_PROPERTYPUT) { sr->pieRotation = tcArgD(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetD(res, sr->pieRotation); return S_OK;
            case 2: tcSetObj(res, tcCached(c, TC_KIND_SLICES, o->index)); return S_OK;
            case 3: tcSetObj(res, tcCached(c, TC_KIND_OTHER, o->index)); return S_OK;
            default: return DISP_E_MEMBERNOTFOUND;
            }
        }
        if (o->kind == TC_KIND_OTHER) {
            switch (id) {
            case 1:
                if (flags & DISPATCH_PROPERTYPUT) { sr->otherSliceStyle = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, sr->otherSliceStyle); return S_OK;
            case 2: tcSetD(res, sr->otherSliceValue); return S_OK;
            case 3:
                if (flags & DISPATCH_PROPERTYPUT) { sr->otherSliceVisible = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
                tcSetL(res, sr->otherSliceVisible); return S_OK;
            default: return DISP_E_MEMBERNOTFOUND;
            }
        }
        /* TC_KIND_SLICES：`AsPie.ExplodedSlice.Value(i)`。
         * GET 只有 index 一个实参；PUT 的布局是 rgvarg[0]=新值、rgvarg[1]=下标
         * （与 memberobj 那族同一口径，弄反了症状是“写进去了但写到别的小标上”）。 */
        if (o->kind == TC_KIND_SLICES) {
            int idx = 0;
            if (id != 1) return DISP_E_MEMBERNOTFOUND;
            if (flags & DISPATCH_PROPERTYPUT) {
                int32_t val = (argc >= 1) ? tcArgL(&dp->rgvarg[0]) : 0;
                if (argc >= 2) idx = (int)tcArgL(&dp->rgvarg[1]);
                if (idx >= 0 && idx < 64) {
                    sr->sliceExplode[idx] = val;
                    InvalidateRect(c->hwnd, NULL, FALSE);
                }
                return S_OK;
            }
            idx = (int)tcArgL(a0);
            tcSetL(res, (idx >= 0 && idx < 64) ? sr->sliceExplode[idx] : 0);
            return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
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
    case 9:  tcSetObj(res, tcCached(c, TC_KIND_MARKS, o->index)); return S_OK;
    case 10: tcSetObj(res, tcCached(c, TC_KIND_PEN,   o->index)); return S_OK;
    case 11: tcSetObj(res, tcCached(c, TC_KIND_BRUSH, o->index)); return S_OK;
    case 12: tcSetObj(res, tcCached(c, TC_KIND_ASPIE, o->index)); return S_OK;
    default:
        return DISP_E_MEMBERNOTFOUND;
    }
}

static const IDispatchVtbl tcObjVtbl = {
    tcQI, tcAddRef, tcRelease, tcGtiCount, tcGti, tcGetIDs, tcInvoke
};

static TcObj* tcCached(TcChart* c, int kind, int index) {
    int i;
    if (!c) return NULL;
    for (i = 0; i < c->nextra; i++) {
        if (c->extra[i] && c->extra[i]->kind == kind && c->extra[i]->index == index)
            return c->extra[i];
    }
    if (c->nextra >= (int)(sizeof(c->extra) / sizeof(c->extra[0]))) return NULL;
    {
        TcObj* o = (TcObj*)calloc(1, sizeof(TcObj));
        if (!o) return NULL;
        o->lpVtbl = &tcObjVtbl;
        o->ref    = 1;
        o->kind   = kind;
        o->hwnd   = c->hwnd;
        o->index  = index;
        c->extra[c->nextra++] = o;
        return o;
    }
}

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
    /* C30-10 第十刀：文档 §5 那四条 —— 认名字、不动作，但不许静默 438。 */
    if (_wcsicmp(name, L"ShowEditor") == 0 || _wcsicmp(name, L"ShowEditorModal") == 0) {
        return 1;   /* 编辑器是 OCX 自带对话框，内置版不提供 */
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

/* ==================== IEnumVARIANT（For Each s In TChart1.Series） ====================
 * 与 vb6forms_memberobj.c 那族同一形态：真 IEnumVARIANT，Next 时现建一枚系列对象。
 * 枚举器是**临时对象**（不归图表所有），Release 到 0 就 free。
 */
typedef struct TcEnum {
    IEnumVARIANTVtbl* lpVtbl;
    LONG ref;
    HWND hwnd;
    int  pos;     /* 0 基 */
    int  total;
} TcEnum;

static const IEnumVARIANTVtbl tcEnumVtbl;

static void* tcEnumNew(TcChart* c) {
    TcEnum* e;
    if (!c) return NULL;
    e = (TcEnum*)calloc(1, sizeof(TcEnum));
    if (!e) return NULL;
    e->lpVtbl = (IEnumVARIANTVtbl*)&tcEnumVtbl;
    e->ref    = 1;
    e->hwnd   = c->hwnd;
    e->pos    = 0;
    e->total  = c->nser;
    return (void*)e;
}

static HRESULT STDMETHODCALLTYPE tcEnumQI(IEnumVARIANT* d, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IEnumVARIANT)) {
        *ppv = d; InterlockedIncrement(&((TcEnum*)d)->ref); return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE tcEnumAddRef(IEnumVARIANT* d) {
    return (ULONG)InterlockedIncrement(&((TcEnum*)d)->ref);
}
static ULONG STDMETHODCALLTYPE tcEnumRelease(IEnumVARIANT* d) {
    LONG r = InterlockedDecrement(&((TcEnum*)d)->ref);
    if (r <= 0) free((TcEnum*)d);
    return (ULONG)(r > 0 ? r : 0);
}
static HRESULT STDMETHODCALLTYPE tcEnumNext(IEnumVARIANT* d, ULONG celt, VARIANT* rg,
                                            ULONG* fetched) {
    TcEnum* e = (TcEnum*)d;
    TcChart* c = tcOf(e->hwnd);
    ULONG got = 0;
    if (!c || !rg) { if (fetched) *fetched = 0; return E_POINTER; }
    while (got < celt && e->pos < e->total) {
        TcObj* so = tcObjFor(c, TC_KIND_SERIES, e->pos);
        if (!so) break;
        tcSetObj(&rg[got], so);
        e->pos++;
        got++;
    }
    if (fetched) *fetched = got;
    return (got == celt) ? S_OK : S_FALSE;
}
static HRESULT STDMETHODCALLTYPE tcEnumSkip(IEnumVARIANT* d, ULONG celt) {
    TcEnum* e = (TcEnum*)d;
    e->pos += (int)celt;
    if (e->pos > e->total) e->pos = e->total;
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE tcEnumReset(IEnumVARIANT* d) {
    ((TcEnum*)d)->pos = 0;
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE tcEnumClone(IEnumVARIANT* d, IEnumVARIANT** out) {
    TcEnum* e = (TcEnum*)d;
    TcEnum* n;
    if (!out) return E_POINTER;
    n = (TcEnum*)calloc(1, sizeof(TcEnum));
    if (!n) return E_OUTOFMEMORY;
    *n = *e;
    n->ref = 1;
    *out = (IEnumVARIANT*)n;
    return S_OK;
}
static const IEnumVARIANTVtbl tcEnumVtbl = {
    (HRESULT (STDMETHODCALLTYPE*)(IUnknown*, REFIID, void**))tcEnumQI,
    (ULONG (STDMETHODCALLTYPE*)(IUnknown*))tcEnumAddRef,
    (ULONG (STDMETHODCALLTYPE*)(IUnknown*))tcEnumRelease,
    tcEnumNext, tcEnumSkip, tcEnumReset, tcEnumClone
};

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
    if (_wcsicmp(name, L"Walls") == 0)  { tcSetObj(out, tcCached(c, TC_KIND_WALLS, 0)); return 1; }
    if (_wcsicmp(name, L"Aspect") == 0) { tcSetObj(out, tcCached(c, TC_KIND_ASPECT, 0)); return 1; }
    if (_wcsicmp(name, L"Printer") == 0) { tcSetObj(out, tcCached(c, TC_KIND_PRINTER, 0)); return 1; }
    if (_wcsicmp(name, L"Export") == 0)  { tcSetObj(out, tcCached(c, TC_KIND_EXPORT, 0)); return 1; }
    if (_wcsicmp(name, L"Zoom") == 0)    { tcSetObj(out, tcCached(c, TC_KIND_ZOOM, 0)); return 1; }
    if (_wcsicmp(name, L"Panel") == 0) { tcSetObj(out, tcCached(c, TC_KIND_PANEL, 0)); return 1; }
    if (_wcsicmp(name, L"Header") == 0) { tcSetObj(out, tcCached(c, TC_KIND_HEADER, 0)); return 1; }
    if (_wcsicmp(name, L"Axis") == 0 || _wcsicmp(name, L"Axes") == 0) {
        tcSetObj(out, tcCached(c, TC_KIND_AXES, 0));
        return 1;
    }
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
