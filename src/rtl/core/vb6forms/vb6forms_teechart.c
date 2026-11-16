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

/* 账 #305 §B224 的 R1：目的档 BSTR 的唯一出口是 vb6_ComChangeType（定义在 vb6rtl_com.c）。
 * 本文件不 include vb6rtl_variant.h，与 vb6_ErrRaise 同一惯例就地声明。 */
extern HRESULT vb6_ComChangeType(VARIANT* dest, const VARIANT* src, VARTYPE want);

#define VB6_TC_CLASS     L"VB6_TCHART"
#define VB6_TC_MAXCTL    32      /* 同进程最多几枚图表 */
#define VB6_TC_MAXSER    32      /* 每图最多几条系列 */
#define VB6_TC_MAXPTS    4096    /* 每条系列最多几个点 */

/* TeeChart 的 SeriesType 常数（真 OCX 的 scLine..scPie 取值，示例里直接写数字也要对得上）
 *
 * C30-10 第十八刀：以前没接的类型号一律被当成 stLine 静默画成折线（不报错、但图是错的）。
 * 现在改成两档：下面这 9 个真画，其余一律抛 VB 错误 5（调用/参数无效）。
 * 编号按 TeeChart 的 ESeriesType 顺序（0..5 与已有实现完全对得上就是佐证）：
 *   6 scFastLine  7 scSurface  8 scCandle  9 scGantt  10 scShape
 *   11 scBubble  12 scArrow  13 scPolar  14 scRadar  15 scPolarBar
 *   16 scBezier  17 scDonut  18 scErrorBar …
 * 仍未做的（Candle 要 OHLC 四列、Surface 要二维网格）现在会明确报错，不再画错图。 */
#define TC_SC_LINE      0
#define TC_SC_BAR       1
#define TC_SC_HBAR      2
#define TC_SC_AREA      3
#define TC_SC_POINT     4
#define TC_SC_PIE       5
#define TC_SC_GANTT     9
#define TC_SC_BUBBLE    11
#define TC_SC_DONUT     17

typedef struct { double x, y; } TcPoint;

/* C30-10 第十八刀：RTL 的报错入口（vb6rtl_array.c 实现）。只声明用得到的这一个，
 * 不为了一行调用把 vb6rtl 头拖进来。 */
extern void vb6_ErrRaise(int32_t errNum, BSTR source, BSTR description);

/* 真能画的那几个类型号（其余一律报错，见 vb6_TChart_AddSeries） */
static int tcTypeSupported(int t) {
    switch (t) {
    case TC_SC_LINE: case TC_SC_BAR: case TC_SC_HBAR: case TC_SC_AREA:
    case TC_SC_POINT: case TC_SC_PIE: case TC_SC_GANTT: case TC_SC_BUBBLE:
    case TC_SC_DONUT:
        return 1;
    default:
        return 0;
    }
}

/* 气泡图的第三列（半径）。按需分配：绝大多数图表没有气泡系列，
 * 每系列常驻 16KB 太浪费（32 系列 × 32 图）。 */
#define TC_IS_PIE(t)   ((t) == TC_SC_PIE || (t) == TC_SC_DONUT)

typedef struct {
    int    type;                 /* TC_SC_* */
    COLORREF color;
    wchar_t title[64];
    TcPoint pt[VB6_TC_MAXPTS];
    double* rad;                  /* C30-10 第十八刀：气泡半径（惰性分配，见 AddBubble） */
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
    /* C30-10 第十七刀：切面里真影响几何 / 只存值的那几个 */
    int      barPct;         /* AsBar/AsHorizBar.BarWidthPercent，默认 100 */
    int      lineSmoothed;   /* AsLine.Smoothed（只存不画：平滑会改变折线点数） */
    int      candleWidth;    /* AsCandle.CandleWidth（只存：没有 OHLC 数据可画） */
    int      surfWire;       /* AsSurface.WireFrame（只存：没有网格数据可画） */
    int      dotSize;        /* AsPoint.DotSize（只存，点径目前固定 r=3） */
    int      asMiscA, asMiscB;
} TcSeries;

typedef struct {
    HWND     hwnd;
    COLORREF panelColor;
    int      marginL, marginR, marginT, marginB;   /* 像素 */
    int      effMarginR, effMarginB;               /* 钳过自动下限后的值（只给绘制用） */
    wchar_t  header[256];
    int      showLegend;
    /* C30-10 第十刀：Zoom 的两个值（空实现里也照实存，读回才一致） */
    int      zoomEnable, zoomPercent;
    /* C30-10 第十一刀：Aspect（3D）。深度只在 View3D 时非零 */
    int      view3D, perspective, aspRot, aspElev;
    /* C30-10 第十二刀：四壁 */
    int      wallsVisible;
    COLORREF wallsColor;
    /* C30-10 第十九刀：ChartDisplay 官方示例要的成员。
     * 这些值都“存得住、读得回”；真正接进绘制的只有 wallsVisible/wallsColor/panelBorder
     * 那几个老的，其余（Frame 边框、渐变、分面尺寸/透明度、BevelWidth…）目前不改画面。 */
    int      frameVisible;
    int      gradVisible;
    COLORREF gradStart, gradEnd;
    int      panelBevelOffset, panelBevelWidth;
    int      panelBevelInner, panelBevelOuter;
    int      backImageInside;
    int      chart3DPercent;
    int      wallSize[4];          /* 0=Back 1=Bottom 2=Right 3=Left（TeeChart 是百分比，默认 3） */
    int      wallFaceVisible[4];
    int      wallTransparent[4];
    COLORREF wallFaceColor[4];
    COLORREF wallBrush[4];
    TcSeries ser[VB6_TC_MAXSER];
    int      nser;
    int      paints;             /* 绘制次数（C3_TCHART_TRACE 用） */
    struct TcObj* obj[VB6_TC_MAXSER + 1];   /* 成员对象缓存 (0 = Series 集合) */
    struct TcObj* meth[8];                  /* 方法对象缓存（AddSeries 那几枚） */
    int      nmeth;
    struct TcObj* extra[32];                /* Panel / Header / Axes / Axis / Title / Frame / Gradient / 四壁分面 */
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
    /* 第十九刀：默认值照 TeeChart 文档（Frame 可见、渐变关、Chart3DPercent 50、壁厚 3） */
    c->frameVisible    = 1;
    c->gradVisible     = 0;
    c->gradStart       = RGB(255, 255, 255);
    c->gradEnd         = RGB(192, 192, 192);
    c->chart3DPercent  = 50;
    {
        int wi;
        for (wi = 0; wi < 4; wi++) {
            c->wallSize[wi]        = 3;
            c->wallFaceVisible[wi] = 1;
            c->wallTransparent[wi] = 0;
            c->wallFaceColor[wi]   = RGB(255, 255, 255);
            c->wallBrush[wi]       = RGB(255, 255, 255);
        }
    }
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
        s->barPct = 100;        /* AsBar/AsHorizBar.BarWidthPercent 默认 */
        s->candleWidth = 60;
        s->surfWire = 1;
        s->dotSize = 3;
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

/* C30-10 第十五刀：把绘制体从“只认 WM_PAINT”里抽出来 —— 导出/剪贴板要能画到任意 DC。
 * 只做机械拆分：BeginPaint/GetClientRect 留在 tcPaint，rc 由调用方给。 */
static void tcRender(TcChart* c, HDC dc, RECT rc) {
    RECT plot;
    HBRUSH br;
    HPEN   oldPen, pen;
    HGDIOBJ oldBr;
    HFONT   fnt, fntOld;
    int i, n, x0, x1, y0, y1;
    int marksDrawn = 0;   /* C30-10 第六刀：本次真的画了多少个逐点标签 */
    int headerH = 0, legendW = 0, legendLeft = 0, legendRight = 0;
    double vmin, vmax;

    if (!c || !dc) return;

    /* 字体：TeeChart 的图例/标题用的是自己的字体（不是系统 GUI 字体）。
     * 直接沿用默认字体量出来的“Series0”宽 53px，而参照读数的图例框只有 61 宽
     * （“TeeChart display demo” 只占 129px），对不上 —— 换 Arial 8pt 后两项都对。 */
    fnt    = CreateFontW(-11, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET,
                         OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY,
                         VARIABLE_PITCH | FF_SWISS, L"Arial");
    fntOld = (HFONT)SelectObject(dc, fnt);

    /* Panel 底：Gradient.Visible 时竖直线性渐变（上 StartColor → 下 EndColor），否则纯色。
     * C30-10 第二十刀：示例里 Check1/Option1..4 就管这个。 */
    if (c->gradVisible) {
        int yy, h = (rc.bottom > rc.top) ? (rc.bottom - rc.top) : 1;
        int sr = GetRValue(c->gradStart), sg = GetGValue(c->gradStart), sb = GetBValue(c->gradStart);
        int er = GetRValue(c->gradEnd),   eg = GetGValue(c->gradEnd),   eb = GetBValue(c->gradEnd);
        for (yy = rc.top; yy < rc.bottom; yy++) {
            int t = (yy - rc.top) * 100 / h;
            RECT ln;
            HBRUSH gb = CreateSolidBrush(RGB(sr + (er - sr) * t / 100,
                                             sg + (eg - sg) * t / 100,
                                             sb + (eb - sb) * t / 100));
            ln.left = rc.left; ln.right = rc.right; ln.top = yy; ln.bottom = yy + 1;
            FillRect(dc, &ln, gb);
            DeleteObject(gb);
        }
    } else {
        br = CreateSolidBrush(c->panelColor);
        oldBr = SelectObject(dc, br);
        FillRect(dc, &rc, br);
        SelectObject(dc, oldBr);
        DeleteObject(br);
    }
    /* Bevel：BevelWidth>0 时在面板内侧画斜角（左上亮、右下暗）。示例的 VScroll2 就在拖这个。 */
    if (c->panelBevelWidth > 0) {
        int bw = c->panelBevelWidth;
        int k;
        if (bw > 8) bw = 8;
        for (k = 0; k < bw; k++) {
            HPEN pl = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));
            HPEN pd = CreatePen(PS_SOLID, 1, RGB(96, 96, 96));
            HPEN po;
            po = (HPEN)SelectObject(dc, pl);
            MoveToEx(dc, rc.left + k, rc.bottom - 1 - k, NULL);
            LineTo(dc, rc.left + k, rc.top + k);
            LineTo(dc, rc.right - 1 - k, rc.top + k);
            SelectObject(dc, pd);
            MoveToEx(dc, rc.right - 1 - k, rc.top + k, NULL);
            LineTo(dc, rc.right - 1 - k, rc.bottom - 1 - k);
            LineTo(dc, rc.left + k, rc.bottom - 1 - k);
            SelectObject(dc, po);
            DeleteObject(pl);
            DeleteObject(pd);
        }
    }
    pen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));
    oldPen = (HPEN)SelectObject(dc, pen);
    /* 外框：Panel.BorderStyle(Check3) 与 Frame.Visible(Check6) 任一开着就画 */
    if (c->panelBorder || c->frameVisible) {
        MoveToEx(dc, rc.left, rc.top, NULL);
        LineTo(dc, rc.right - 1, rc.top);
        LineTo(dc, rc.right - 1, rc.bottom - 1);
        LineTo(dc, rc.left, rc.bottom - 1);
        LineTo(dc, rc.left, rc.top);
    }

    /* ================= C30-10 第七刀：面板分带 =================
     * 真 TeeChart 不是“客户区减值当绘图区”，而是把图表分成几条带：
     *   header 带 0..26（标题文字垂直居中在这一带）、上轴带 26..53、
     *   绘图区 53..(H-37)、下轴带 25px + MarginBottom、图例在右侧帮忙挤绘图区。
     * 这些数就是 401x257 参照读数反推出来的（上轴 y=53、下轴 y=220、
     *   图例 x=310..371 y=26、标题 bbox 中心 155.5）。
     */
    if (c->header[0]) headerH = 26;
    /* C30-10 第十九刀：Margin* 的自动下限（绘制用局部值，不改存进去的原值 —— 读回仍是你设的数）。
     * 凭据是 ChartDisplay 参照：它在 Load 里把 HScroll4/VScroll4 设成 32 ⇒ MarginRight/MarginBottom
     * = 38-32 = 6，而真 VB6 的图例依旧 310..371、下轴依旧 y=220（与默认 3% 一模一样）——
     * OCX 把过小的 Margin 钳在“能装下刻度文字”的下限。我们这四值以像素计、默认 12（就是 3% × 401），
     * 下限取 12：设得更大照算，设得更小当 12。 */
    {
        int mr = (c->marginR < 12) ? 12 : c->marginR;
        int mb = (c->marginB < 12) ? 12 : c->marginB;
        c->effMarginR = mr;
        c->effMarginB = mb;
    }
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
        legendRight = rc.right - c->effMarginR - 18;
        legendLeft  = legendRight - legendW;
        if (legendLeft < rc.left + 60) legendLeft = rc.left + 60;   /* 太窄时给绘图区留地面 */
    }

    plot = rc;
    plot.left   += c->marginL + 16;                       /* 左轴刻度文字位 */
    plot.top    = rc.top + headerH + 27;                  /* header 带 + 上轴带 */
    plot.right  = (legendW > 0) ? (legendLeft - 8) : (rc.right - c->effMarginR);
    plot.bottom = rc.bottom - 25 - c->effMarginB;            /* 下轴带 + MarginBottom */
    if (plot.right <= plot.left + 2 || plot.bottom <= plot.top + 2) {
        SelectObject(dc, fntOld); DeleteObject(fnt);
        SelectObject(dc, oldPen); DeleteObject(pen);
        return;
    }

    /* 四壁：3D 下才画，且必须在**轴线之前**（墙是背板，不是覆盖层）。
     * C30-10 第二十刀：从“一块灰板”改成按分面画：
     *   · Check11 管 Walls.Visible；Command4/5 管 LeftWall/BottomWall（分面 3 / 1）
     *   · Check4 管 Walls.Back.Transparent（透明就用面板底色，不再盖一层灰）
     *   · Command11 / Walls.Back.Brush.Color 改背板颜色
     *   · HScroll2 管 Walls.*.Size；HScroll1 管 Aspect.Chart3DPercent（深度 = 12 × pct/50） */
    if (c->view3D && c->wallsVisible) {
        int pct = (c->chart3DPercent > 0) ? c->chart3DPercent : 1;
        int d   = (c->perspective ? 18 : 12) * pct / 50;
        if (c->wallSize[0] == 0)      d = 0;                 /* Size=0：墙贴平 */
        else if (c->wallSize[0] > 3)  d += (c->wallSize[0] - 3);
        if (d < 0)  d = 0;
        if (d > 40) d = 40;
        if (d > 0) {
            POINT back[4], lwall[4], bwall[4];
            int fc;
            /* 背板颜色优先序：Brush.Color → Walls.Back.Color → Walls.BackColor → 墙上色 */
            if (c->wallBrush[0] != RGB(255, 255, 255))     fc = (int)c->wallBrush[0];
            else if (c->wallFaceColor[0] != RGB(255, 255, 255)) fc = (int)c->wallFaceColor[0];
            else                                            fc = (int)c->wallsColor;
            if (c->wallTransparent[0]) fc = (int)c->panelColor;   /* 透明：露面板底 */
            back[0].x = plot.left + d;  back[0].y = plot.top - d;
            back[1].x = plot.right + d; back[1].y = plot.top - d;
            back[2].x = plot.right + d; back[2].y = plot.bottom - d;
            back[3].x = plot.left + d;  back[3].y = plot.bottom - d;
            lwall[0].x = plot.left;     lwall[0].y = plot.top;
            lwall[1].x = plot.left + d; lwall[1].y = plot.top - d;
            lwall[2].x = plot.left + d; lwall[2].y = plot.bottom - d;
            lwall[3].x = plot.left;     lwall[3].y = plot.bottom;
            bwall[0].x = plot.left;     bwall[0].y = plot.bottom;
            bwall[1].x = plot.right;    bwall[1].y = plot.bottom;
            bwall[2].x = plot.right + d; bwall[2].y = plot.bottom - d;
            bwall[3].x = plot.left + d; bwall[3].y = plot.bottom - d;
            if (c->wallFaceVisible[0]) {
                HBRUSH wb = CreateSolidBrush((COLORREF)fc);
                HGDIOBJ wOld = SelectObject(dc, wb);
                HGDIOBJ wpOld = SelectObject(dc, GetStockObject(NULL_PEN));
                Polygon(dc, back, 4);
                SelectObject(dc, wpOld);
                SelectObject(dc, wOld);
                DeleteObject(wb);
            }
            if (c->wallFaceVisible[3]) {
                HBRUSH wb = CreateSolidBrush(tcShade((COLORREF)fc, -18));
                HGDIOBJ wOld = SelectObject(dc, wb);
                HGDIOBJ wpOld = SelectObject(dc, GetStockObject(NULL_PEN));
                Polygon(dc, lwall, 4);
                SelectObject(dc, wpOld);
                SelectObject(dc, wOld);
                DeleteObject(wb);
            }
            if (c->wallFaceVisible[1]) {
                HBRUSH wb = CreateSolidBrush(tcShade((COLORREF)fc, -10));
                HGDIOBJ wOld = SelectObject(dc, wb);
                HGDIOBJ wpOld = SelectObject(dc, GetStockObject(NULL_PEN));
                Polygon(dc, bwall, 4);
                SelectObject(dc, wpOld);
                SelectObject(dc, wOld);
                DeleteObject(wb);
            }
        }
    }

    /* Header：横向居中于“客户区 - 图例”那段（参照读数 centerX=155.5 = 图例左边/2） */
    if (c->header[0]) {
        RECT hr;
        /* C30-10 第十九刀：标题单独一支字体。判据参照的标题 bbox 宽 129px、高 (14..24)，
         * 而图例文字那支 Arial -11 只量出 98px —— OCX 的 Header 和图例就不是同一号字。
         * -14 是拿参考宽度反推的（98 × 14/11 ≈ 125），实测再微调。 */
        HFONT fh  = CreateFontW(-14, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET,
                                OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY,
                                VARIABLE_PITCH | FF_SWISS, L"Arial");
        HFONT fho = fh ? (HFONT)SelectObject(dc, fh) : NULL;
        hr.left   = rc.left;
        hr.right  = (legendW > 0) ? legendLeft : rc.right;
        hr.top    = rc.top + 11;
        hr.bottom = rc.top + 26;
        tcDrawText(dc, c->header, &hr, 1, RGB(0, 0, 0));
        if (fh) { SelectObject(dc, fho); DeleteObject(fh); }
    }

    /* 上/下/左/右轴线：TeeChart 默认画在绘图区四沿，且是**黑色**。
     * C30-10 第二十刀：四根都受 Axis.X.Visible（示例的 Check7/8/9/10 与 Command2/3）控制 ——
     * 以前 axisShow 只存不画，勾选框才一个个都没反应。下标: 0=Left 1=Right 2=Top 3=Bottom */
    {
        HPEN ap = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
        HPEN apOld = (HPEN)SelectObject(dc, ap);
        if (c->axisShow[2]) {
            MoveToEx(dc, plot.left, plot.top, NULL);
            LineTo(dc, plot.right, plot.top);
        }
        if (c->axisShow[3]) {
            MoveToEx(dc, plot.left, plot.bottom, NULL);
            LineTo(dc, plot.right, plot.bottom);
        }
        if (c->axisShow[0]) {
            MoveToEx(dc, plot.left, plot.top, NULL);
            LineTo(dc, plot.left, plot.bottom);
        }
        if (c->axisShow[1]) {
            MoveToEx(dc, plot.right, plot.top, NULL);
            LineTo(dc, plot.right, plot.bottom);
        }
        SelectObject(dc, apOld);
        DeleteObject(ap);
    }

    /* Y 轴量程（所有系列合起来；空系列给 0..1） */
    vmin = 1e300; vmax = -1e300;
    for (i = 0; i < c->nser; i++) {
        TcSeries* s = &c->ser[i];
        int k;
        if (TC_IS_PIE(s->type)) continue;
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

        if (TC_IS_PIE(s->type)) {
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
            /* C30-10 第十八刀：Donut 就是“饼图中间挖个洞”。几何与饼图完全同源，
             * 只是最后用 Panel 底色把内圆盖掉（洞半径固定 50%，与 TeeChart 默认同量级）。 */
            if (s->type == TC_SC_DONUT) {
                int ir = r / 2;
                HBRUSH hb = CreateSolidBrush(c->panelColor);
                HGDIOBJ hbOld = SelectObject(dc, hb);
                HGDIOBJ hpOld = SelectObject(dc, GetStockObject(NULL_PEN));
                Ellipse(dc, cx0 - ir, cy0 - ir, cx0 + ir, cy0 + ir);
                SelectObject(dc, hpOld);
                SelectObject(dc, hbOld);
                DeleteObject(hb);
                SelectObject(dc, sb);
                SelectObject(dc, sp);
            }
        } else if (s->n > 0) {
            int bw, gap, k2;
            switch (s->type) {
            case TC_SC_GANTT: {
                /* C30-10 第十八刀：甘特图 —— 每个点是一条任务：AddXY 的 x=开始、y=结束。
                 * 横轴用本系列自己的 x 定义域（[min,max] 含两列），纵向一点一条带。
                 * 这是第一个真正用到 pt.x 的图 —— 其他系列仍按点序排。 */
                double xmin = 0, xmax = 0;
                int band = (s->n > 0) ? ((y1 - y0) / s->n) : (y1 - y0);
                int bh = band / 2;
                if (bh < 2) bh = 2;
                xmin = xmax = s->pt[0].x;
                for (k2 = 0; k2 < s->n; k2++) {
                    if (s->pt[k2].x < xmin) xmin = s->pt[k2].x;
                    if (s->pt[k2].x > xmax) xmax = s->pt[k2].x;
                    if (s->pt[k2].y < xmin) xmin = s->pt[k2].y;
                    if (s->pt[k2].y > xmax) xmax = s->pt[k2].y;
                }
                if (xmax <= xmin) xmax = xmin + 1;
                for (k2 = 0; k2 < s->n; k2++) {
                    int by = y0 + k2 * band + (band - bh) / 2;
                    int gx0 = x0 + (int)((s->pt[k2].x - xmin) / (xmax - xmin) * (x1 - x0));
                    int gx1 = x0 + (int)((s->pt[k2].y - xmin) / (xmax - xmin) * (x1 - x0));
                    if (gx1 < gx0) { int t2 = gx0; gx0 = gx1; gx1 = t2; }
                    if (gx1 - gx0 < 2) gx1 = gx0 + 2;
                    Rectangle(dc, gx0, by, gx1, by + bh);
                }
                break;
            }
            case TC_SC_BUBBLE: {
                /* C30-10 第十八刀：气泡图 —— 位置由 (x,y) 定，半径来自 AddBubble 的第三列。
                 * 半径没给的按 0 处理（画成最小点），不猜。 */
                double xmin = s->pt[0].x, xmax = s->pt[0].x;
                double rmin = 0, rmax = 0;
                for (k2 = 0; k2 < s->n; k2++) {
                    if (s->pt[k2].x < xmin) xmin = s->pt[k2].x;
                    if (s->pt[k2].x > xmax) xmax = s->pt[k2].x;
                    if (s->rad) {
                        if (s->rad[k2] < rmin) rmin = s->rad[k2];
                        if (s->rad[k2] > rmax) rmax = s->rad[k2];
                    }
                }
                if (xmax <= xmin) xmax = xmin + 1;
                if (rmax <= rmin) rmax = rmin + 1;
                for (k2 = 0; k2 < s->n; k2++) {
                    int bx = x0 + (int)((s->pt[k2].x - xmin) / (xmax - xmin) * (x1 - x0));
                    int by = y1 - (int)((s->pt[k2].y - vmin) / (vmax - vmin) * (y1 - y0));
                    int rr = 3;
                    if (s->rad) {
                        rr = 3 + (int)((s->rad[k2] - rmin) / (rmax - rmin) * 9.0);
                        if (rr < 2) rr = 2;
                        if (rr > 14) rr = 14;
                    }
                    Ellipse(dc, bx - rr, by - rr, bx + rr, by + rr);
                }
                break;
            }
            case TC_SC_HBAR: {
                /* C30-10 第十七刀：横向柱真的横着画。之前 type=2 落到 default 成了折线，
                 * 看着“不报错”、其实是画错的图 —— 像素判据一量就现形。
                 * 每根柱占一条水平带，长度按值映射，宽度受 BarWidthPercent 控制。 */
                int band = (s->n > 0) ? ((y1 - y0) / s->n) : (y1 - y0);
                int bh = band * ((s->barPct > 0) ? s->barPct : 100) / 200;   /* 默认占带的一半 */
                if (bh < 2) bh = 2;
                if (bh > band) bh = band;
                for (k2 = 0; k2 < s->n; k2++) {
                    int by = y0 + k2 * band + (band - bh) / 2;
                    int bx1 = x0 + (int)((s->pt[k2].y - vmin) / (vmax - vmin) * (x1 - x0));
                    if (bx1 < x0) bx1 = x0;
                    if (bx1 > x1) bx1 = x1;
                    Rectangle(dc, x0, by, bx1, by + bh);
                }
                break;
            }
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
                bw  = bw * ((s->barPct > 0) ? s->barPct : 100) / 100;
                gap = (s->n > 0) ? ((x1 - x0) / (s->n * 2)) : 8;   /* 间距固定：只缩柱子本身 */
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
                /* C30-10 第十四刀：点系列的 3D —— 每个点画成一枚有厚度的“扁墩”：先画
                 * 暗色偏移轮寗，再压上正面。深度按点径定（常量 4，点本身就是 r=3），
                 * 跟柱形那套 12px 不是一回事：柱子有高度，点只有大小。
                 * 透视不参与 —— 4px 的偏移再缩放就只有 1~2px，量不出来也看不出。 */
                for (k2 = 0; k2 < s->n; k2++) {
                    int px = (s->n == 1) ? (x0 + x1) / 2
                            : x0 + (int)((double)k2 / (s->n - 1) * (x1 - x0));
                    int py = y1 - (int)((s->pt[k2].y - vmin) / (vmax - vmin) * (y1 - y0));
                    if (c->view3D) {
                        HBRUSH pb = CreateSolidBrush(tcShade(fcol, -45));
                        HGDIOBJ pbOld = SelectObject(dc, pb);
                        HGDIOBJ ppOld = SelectObject(dc, GetStockObject(NULL_PEN));
                        Ellipse(dc, px + 4 - 3, py - 4 - 3, px + 4 + 3, py - 4 + 3);
                        SelectObject(dc, ppOld);
                        SelectObject(dc, pbOld);
                        DeleteObject(pb);
                        SelectObject(dc, sb);   /* 正面回到系列/画刷色 */
                        SelectObject(dc, sp);
                    }
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
}

static void tcPaint(TcChart* c) {
    PAINTSTRUCT ps;
    HDC dc;
    RECT rc;
    if (!c) return;
    dc = BeginPaint(c->hwnd, &ps);
    if (!dc) return;
    GetClientRect(c->hwnd, &rc);
    tcRender(c, dc, rc);
    EndPaint(c->hwnd, &ps);
}

/* ================= C30-10 第十五刀：导出 / 剪贴板 —— 真渲染，不再是空实现 =================
 * 导出靠 tcRender 画到内存 DC，所以图形与屏幕同源（同一份绘制代码），
 * 不会出现“屏幕一套、导出另一套”的分叉。 */
typedef struct { HDC screen, mem; HBITMAP bmp; HGDIOBJ old; int w, h; } TcShot;

static int tcShotTake(TcChart* c, TcShot* s) {
    RECT r, rr;
    memset(s, 0, sizeof *s);
    if (!c) return 0;
    GetClientRect(c->hwnd, &r);
    s->w = r.right - r.left;
    s->h = r.bottom - r.top;
    if (s->w <= 0 || s->h <= 0) return 0;
    s->screen = GetDC(NULL);
    s->mem    = CreateCompatibleDC(s->screen);
    s->bmp    = CreateCompatibleBitmap(s->screen, s->w, s->h);
    if (!s->mem || !s->bmp) return 0;
    s->old = SelectObject(s->mem, s->bmp);
    rr.left = 0; rr.top = 0; rr.right = s->w; rr.bottom = s->h;
    tcRender(c, s->mem, rr);
    return 1;
}

static void tcShotFree(TcShot* s) {
    if (s->mem) { if (s->old) SelectObject(s->mem, s->old); DeleteDC(s->mem); s->mem = NULL; }
    if (s->screen) { ReleaseDC(NULL, s->screen); s->screen = NULL; }
}

/* 剪贴板：给 CF_BITMAP 必须是 DDB（DIB 段交过去会被拒绝/格式不符），
 * 所以这里用 CreateCompatibleBitmap 而不是 CreateDIBSection。
 *
 * 坑：位图仍选在内存 DC 里时 SetClipboardData **会失败**（GDI 不允许把
 * 正在用的位图交出去）。第一版就是先 Set 后 Free，结果剪贴板里什么都没有。
 * 所以顺序必须是：渲染 → 脱离 DC → Set → 成功后所有权归剪贴板，不能再删。 */
static int tcExportClipboard(TcChart* c) {
    TcShot s;
    int ok = 0;
    if (!tcShotTake(c, &s)) return 0;
    tcShotFree(&s);
    if (OpenClipboard(c->hwnd) || OpenClipboard(NULL)) {
        EmptyClipboard();
        ok = SetClipboardData(CF_BITMAP, s.bmp) ? 1 : 0;
        CloseClipboard();
    }
    if (tcTrace() > 0) {
        fprintf(stderr, "[TCHART] clipboard %dx%d ok=%d\n", s.w, s.h, ok);
        fflush(stderr);
    }
    if (!ok) DeleteObject(s.bmp);
    return ok;
}

/* JPEG：走系统自带的 gdiplus.dll，动态加载而不是加链接依赖 ——
 * 这张 .c 同时进 x86/x64 的 RTL，自己写一份基线 JPEG 编码器不值得。
 *
 * 编码器 CLSID 直接写死（{557CF401-1A04-11D3-9A73-0000F81EF32E} = image/jpeg），
 * **不走 GdipGetImageEncoders 枚举**：实测两处都不对 —— 枚举返回的 size 与
 * sizeof(ImageCodecInfo) 对不上（num=5 size=1264），按自己声明的结构体读，
 * codec[2].MimeType 读到 null、循环后直接 AV。写死一个官方 CLSID 既短又稳。 */
typedef struct {
    UINT32 GdiplusVersion;
    void*  DebugEventCallback;
    BOOL   SuppressBackgroundThread;
    BOOL   SuppressExternalCodecs;
} TcGpInput;

static const GUID kTcClsidJpeg =
    { 0x557CF401, 0x1A04, 0x11D3, { 0x9A, 0x73, 0x00, 0x00, 0xF8, 0x1E, 0xF3, 0x2E } };

typedef int (WINAPI* TcGpStartup)(ULONG_PTR*, const TcGpInput*, void*);
typedef int (WINAPI* TcGpShutdown)(ULONG_PTR);
typedef int (WINAPI* TcGpFromBmp)(HBITMAP, HPALETTE, void**);
typedef int (WINAPI* TcGpSave)(void*, const wchar_t*, const GUID*, void*);
typedef int (WINAPI* TcGpDispose)(void*);

static int tcSaveJpeg(TcChart* c, const wchar_t* path) {
    HMODULE dll;
    TcGpStartup pStartup = NULL;
    TcGpShutdown pShutdown = NULL;
    TcGpFromBmp pFromBmp = NULL;
    TcGpSave pSave = NULL;
    TcGpDispose pDispose = NULL;
    TcShot s;
    TcGpInput in;
    ULONG_PTR token = 0;
    void* img = NULL;
    int rc = 0;

    if (!path || !path[0]) return 0;
    dll = LoadLibraryW(L"gdiplus.dll");
    if (!dll) return 0;
    pStartup  = (TcGpStartup)GetProcAddress(dll, "GdiplusStartup");
    pShutdown = (TcGpShutdown)GetProcAddress(dll, "GdiplusShutdown");
    pFromBmp  = (TcGpFromBmp)GetProcAddress(dll, "GdipCreateBitmapFromHBITMAP");
    pSave     = (TcGpSave)GetProcAddress(dll, "GdipSaveImageToFile");
    pDispose  = (TcGpDispose)GetProcAddress(dll, "GdipDisposeImage");
    if (!pStartup || !pShutdown || !pFromBmp || !pSave || !pDispose) {
        FreeLibrary(dll);
        return 0;
    }
    memset(&in, 0, sizeof in);
    in.GdiplusVersion = 1;
    if (pStartup(&token, &in, NULL) != 0) { FreeLibrary(dll); return 0; }
    if (!tcShotTake(c, &s)) { pShutdown(token); FreeLibrary(dll); return 0; }
    if (pFromBmp(s.bmp, NULL, &img) == 0 && img) {
        if (pSave(img, path, &kTcClsidJpeg, NULL) == 0) rc = 1;
        pDispose(img);
    }
    tcShotFree(&s);
    DeleteObject(s.bmp);
    pShutdown(token);
    FreeLibrary(dll);
    return rc;
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

int32_t vb6_TChart_SeriesAddBubble(void* hwnd, int32_t idx, double x, double y, double r) {
    TcChart* c = tcOf((HWND)hwnd);
    TcSeries* s;
    if (!c) return 0;
    s = tcSeries(c, (int)idx);
    if (!s || s->n >= VB6_TC_MAXPTS) return 0;
    if (!s->rad) {
        s->rad = (double*)calloc(VB6_TC_MAXPTS, sizeof(double));
        if (!s->rad) return 0;
    }
    s->pt[s->n].x = x;
    s->pt[s->n].y = y;
    s->rad[s->n] = r;
    s->n++;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
    return s->n;
}

int32_t vb6_TChart_SeriesAddBubble(void* hwnd, int32_t idx, double x, double y, double r);
int32_t vb6_TChart_AddSeries(void* hwnd, int32_t type) {
    TcChart* c = tcAttach((HWND)hwnd);
    TcSeries* s;
    if (!c) return -1;
    if (c->nser >= VB6_TC_MAXSER) return -1;
    /* C30-10 第十八刀：不支持的类型号不再静默降级成折线 —— 静默降级会“不报错但画错图”，
     * 比直接报错危险得多。错误 5 = 调用/参数无效（与 VB6 对非法枚举的惯用法一致）。 */
    if (!tcTypeSupported((int)type)) {
        vb6_ErrRaise(5, SysAllocString(L"C3.TChart"),
                        SysAllocString(L"Invalid series type (supported: 0-5, 9 Gantt, 11 Bubble, 17 Donut)"));
        return -1;
    }
    s = &c->ser[c->nser];
    memset(s, 0, sizeof(*s));
    s->type  = (int)type;   /* 合法性已在上面校验过，这里不再夹回 Line */
    s->color = tcPalette(c->nser);
    /* 切面默认值必须在这里也补一遍：AddSeries 走的是这条 memset 路径，
     * 不经过 tcSeries()，否则读 .AsCandle.CandleWidth 得到 0（夹具逮过）。 */
    s->penWidth = 1;
    s->penVisible = 1;
    s->marksVisible = 1;
    s->barPct = 100;
    s->candleWidth = 60;
    s->surfWire = 1;
    s->dotSize = 3;
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
/* C30-10 第十七刀：按类型切换的切面（一个 kind 一族，成员表分开） */
#define TC_KIND_ASBAR   19
#define TC_KIND_ASLINE  20
#define TC_KIND_ASHBAR  21
#define TC_KIND_ASCANDLE 22
#define TC_KIND_ASSURF  23
#define TC_KIND_ASAREA  24
#define TC_KIND_ASPOINT 25
/* C30-10 第十九刀：ChartDisplay 官方示例牵出来的四层 */
#define TC_KIND_FRAME    26  /* 图表外框（Visible） */
#define TC_KIND_GRADIENT 27  /* Panel.Gradient */
#define TC_KIND_WALLFACE 28  /* Walls.Back/Bottom/Right/Left 各一面（index=0..3） */
#define TC_KIND_WALLBRUSH 29 /* 分面的 Brush（Color/Visible），与系列的 Brush 分开存 */

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
    L"Marks", L"Pen", L"Brush", L"AsPie",
    /* C30-10 第十七刀：按类型切换的切面。几何真生效的只有横向柱与 BarWidthPercent，
     * candle/surface 需要 OHLC / 网格数据（我们的点集只有一列 y），只开对象不绘制。 */
    L"AsBar", L"AsLine", L"AsHorizBar", L"AsCandle", L"AsSurface", L"AsArea", L"AsPoint",
    L"AddBubble",       /* C30-10 第十八刀：气泡图第三列（半径） */
    NULL };
/* 切面自己的成员表 */
static const wchar_t* kTcAsBarNames[]   = { L"BarWidthPercent", L"SideMargins", L"OffsetPercent", NULL };
static const wchar_t* kTcAsLineNames[]  = { L"Smoothed", L"Stairs", NULL };
static const wchar_t* kTcAsCandleNames[] = { L"CandleWidth", L"ShowOpen", L"ShowClose", NULL };
static const wchar_t* kTcAsSurfNames[]  = { L"WireFrame", L"DotSize", NULL };
static const wchar_t* kTcAsAreaNames[]  = { L"Transparency", L"Gradient", NULL };
static const wchar_t* kTcAsPointNames[] = { L"DotSize", L"Style", NULL };
/* C30-10 第三刀：Panel / Header / Axis 三层。只登记真做得动的名字。 */
static const wchar_t* kTcPanelNames[] = {
    L"Color", L"MarginLeft", L"MarginRight", L"MarginTop", L"MarginBottom",
    L"Visible", L"BorderVisible",
    /* 第十九刀：ChartDisplay 示例用到的那批（追加在尾部，1..7 的 DISPID 不动） */
    L"Gradient", L"BevelOffset", L"BevelWidth",
    L"BackImageInside", L"BackImageClear", L"BackImageLoad", L"BorderStyle",
    L"BevelInner", L"BevelOuter", NULL };
static const wchar_t* kTcFrameNames[]     = { L"Visible", L"Color", L"Width", L"Style", NULL };
static const wchar_t* kTcGradNames[]      = { L"Visible", L"StartColor", L"EndColor", L"Direction", NULL };
static const wchar_t* kTcWallFaceNames[]  = { L"Visible", L"Size", L"Transparent", L"Color", L"Brush", NULL };
static const wchar_t* kTcWallBrushNames[] = { L"Color", L"Visible", NULL };
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
    L"SaveToFile", L"CopyToClipboard", L"Email", L"Picture", L"Width", L"Height",
    L"SaveToJPEGFile", L"SaveToBitmapFile", NULL };
static const wchar_t* kTcZoomNames[] = {
    L"Enable", L"Percent", L"Direction", L"Point", L"SetMax", NULL };
/* C30-10 第十一刀：3D 面（先让值存得住、读得回，再把柱形按深度画出来） */
static const wchar_t* kTcAspectNames[] = {
    L"View3D", L"Perspective", L"Rotation", L"Elevation", L"Orthogonal",
    L"Chart3DPercent", NULL };
/* C30-10 第十二刀：四壁。Back/Bottom/Right 先指回同一对象（共用可见性/颜色），
 * 分面颜色留到有人真的需要时再拆 —— 现在拆了也画不出区别。 */
static const wchar_t* kTcWallsNames[] = {
    L"Visible", L"Color", L"Back", L"Bottom", L"Right",
    L"BackColor", L"Left", NULL };

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
    case TC_KIND_ASBAR:   return kTcAsBarNames;
    case TC_KIND_ASHBAR:  return kTcAsBarNames;   /* 横向柱共用同一族字段 */
    case TC_KIND_ASLINE:  return kTcAsLineNames;
    case TC_KIND_ASCANDLE: return kTcAsCandleNames;
    case TC_KIND_ASSURF:  return kTcAsSurfNames;
    case TC_KIND_ASAREA:  return kTcAsAreaNames;
    case TC_KIND_ASPOINT: return kTcAsPointNames;
    case TC_KIND_FRAME:   return kTcFrameNames;
    case TC_KIND_GRADIENT: return kTcGradNames;
    case TC_KIND_WALLFACE: return kTcWallFaceNames;
    case TC_KIND_WALLBRUSH: return kTcWallBrushNames;
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
    VARIANT* a2 = (argc >= 3) ? &dp->rgvarg[argc - 3] : NULL;
    (void)riid; (void)lcid; (void)ei; (void)ae; (void)a2;
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
        /* 第十九刀：Gradient 返子对象；其余四个存/读；BackImage* 与 BorderStyle 见下注释 */
        case 8: if (res) tcSetObj(res, tcCached(c, TC_KIND_GRADIENT, 0)); return S_OK;
        case 9: if (flags & DISPATCH_PROPERTYPUT) { c->panelBevelOffset = (int)tcArgL(a0); return S_OK; }
                tcSetL(res, c->panelBevelOffset); return S_OK;
        case 10: if (flags & DISPATCH_PROPERTYPUT) { c->panelBevelWidth = (int)tcArgL(a0); return S_OK; }
                tcSetL(res, c->panelBevelWidth); return S_OK;
        case 11: if (flags & DISPATCH_PROPERTYPUT) { c->backImageInside = (int)tcArgL(a0); return S_OK; }
                tcSetL(res, c->backImageInside); return S_OK;
        case 12: return S_OK;   /* BackImageClear：没加载过背景图，无可清 */
        case 13: return S_OK;   /* BackImageLoad(name)：不加载外部位图（台账已注明） */
        case 15: case 16:       /* BevelInner / BevelOuter：斜角样式，只存 */
                if (flags & DISPATCH_PROPERTYPUT) {
                    if (id == 15) c->panelBevelInner = (int)tcArgL(a0);
                    else          c->panelBevelOuter = (int)tcArgL(a0);
                    return S_OK;
                }
                tcSetL(res, (id == 15) ? c->panelBevelInner : c->panelBevelOuter);
                return S_OK;
        case 14:                /* BorderStyle：bsSingle=1 / bsNone=0 → 落到已有的 panelBorder */
                if (flags & DISPATCH_PROPERTYPUT) {
                    c->panelBorder = (tcArgL(a0) != 0);
                    InvalidateRect(c->hwnd, NULL, FALSE);
                    return S_OK;
                }
                tcSetL(res, c->panelBorder); return S_OK;
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }

    if (o->kind == TC_KIND_HEADER) {
        switch (id) {
        case 1: case 2:   /* Text / Caption */
            if (flags & DISPATCH_PROPERTYPUT) {
                VARIANT t; VariantInit(&t);
                if (a0 && SUCCEEDED(vb6_ComChangeType(&t, a0, VT_BSTR)))
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
                if (a0 && SUCCEEDED(vb6_ComChangeType(&t, a0, VT_BSTR))) {
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
            /* C30-10 第十九刀：Back / Bottom / Right 各给一个独立分面对象。
             * 以前四壁指回 Walls 自己 —— 但 ChartDisplay 要 Walls.Back.Brush.Color /
             * .Back.Size / .Back.Transparent，指回去会把这些名字当成 Walls 的成员。 */
            if (res) tcSetObj(res, tcCached(c, TC_KIND_WALLFACE, (int)id - 3));   /* 0=Back 1=Bottom 2=Right */
            return S_OK;
        case 6:
            if (flags & DISPATCH_PROPERTYPUT) { c->wallsColor = (COLORREF)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, (int32_t)c->wallsColor);
            return S_OK;
        case 7:
            if (res) tcSetObj(res, tcCached(c, TC_KIND_WALLFACE, 3));              /* Left */
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
        case 6:                                         /* Chart3DPercent（第十九刀） */
            if (flags & DISPATCH_PROPERTYPUT) {
                c->chart3DPercent = (int)tcArgL(a0);
                InvalidateRect(c->hwnd, NULL, FALSE);
                return S_OK;
            }
            if (res) tcSetL(res, c->chart3DPercent);
            return S_OK;
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

    /* ===== 第十九刀：Frame / Panel.Gradient / 四壁分面 =====
     * 起因：ChartDisplay 官方示例（tests\TeeChartChartDisplay）一行 TChart1.Frame.Visible
     * 就抛 438 把整个窗体打退。这里只做到“名字认得、值存得住读得回”，
     * 画面影响面在注释里逐个点名 —— 不让未实现的成员变成“看着支持、实则答错”。 */
    if (o->kind == TC_KIND_FRAME) {
        switch (id) {
        case 1:
            /* 画面上目前不单独画外框（Panel.BorderVisible/panelBorder 已在管边框） */
            if (flags & DISPATCH_PROPERTYPUT) { c->frameVisible = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, c->frameVisible);
            return S_OK;
        case 2: if (flags & DISPATCH_PROPERTYPUT) return S_OK; if (res) tcSetL(res, (int32_t)RGB(0, 0, 0)); return S_OK;
        case 3: if (flags & DISPATCH_PROPERTYPUT) return S_OK; if (res) tcSetL(res, 1); return S_OK;
        case 4: if (flags & DISPATCH_PROPERTYPUT) return S_OK; if (res) tcSetL(res, 0); return S_OK;
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }
    if (o->kind == TC_KIND_GRADIENT) {
        switch (id) {
        case 1:
            if (flags & DISPATCH_PROPERTYPUT) { c->gradVisible = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, c->gradVisible);
            return S_OK;
        case 2:
            if (flags & DISPATCH_PROPERTYPUT) { c->gradStart = (COLORREF)tcArgL(a0); return S_OK; }
            if (res) tcSetL(res, (int32_t)c->gradStart);
            return S_OK;
        case 3:
            if (flags & DISPATCH_PROPERTYPUT) { c->gradEnd = (COLORREF)tcArgL(a0); return S_OK; }
            if (res) tcSetL(res, (int32_t)c->gradEnd);
            return S_OK;
        case 4: if (flags & DISPATCH_PROPERTYPUT) return S_OK; if (res) tcSetL(res, 0); return S_OK;   /* Direction */
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }
    if (o->kind == TC_KIND_WALLFACE || o->kind == TC_KIND_WALLBRUSH) {
        int f = (o->index >= 0 && o->index < 4) ? o->index : 0;
        if (o->kind == TC_KIND_WALLBRUSH) {
            switch (id) {
            case 1:
                if (flags & DISPATCH_PROPERTYPUT) { c->wallBrush[f] = (COLORREF)tcArgL(a0); return S_OK; }
                if (res) tcSetL(res, (int32_t)c->wallBrush[f]);
                return S_OK;
            case 2: if (flags & DISPATCH_PROPERTYPUT) return S_OK; if (res) tcSetL(res, 1); return S_OK;
            default: return DISP_E_MEMBERNOTFOUND;
            }
        }
        switch (id) {
        case 1:   /* 分面可见性（C30-10 第二十刀：Check11/Command4/5 真作用） */
            if (flags & DISPATCH_PROPERTYPUT) { c->wallFaceVisible[f] = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, c->wallFaceVisible[f]);
            return S_OK;
        case 2:
            if (flags & DISPATCH_PROPERTYPUT) { c->wallSize[f] = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, c->wallSize[f]);
            return S_OK;
        case 3:
            if (flags & DISPATCH_PROPERTYPUT) { c->wallTransparent[f] = (int)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, c->wallTransparent[f]);
            return S_OK;
        case 4:
            if (flags & DISPATCH_PROPERTYPUT) { c->wallFaceColor[f] = (COLORREF)tcArgL(a0); InvalidateRect(c->hwnd, NULL, FALSE); return S_OK; }
            if (res) tcSetL(res, (int32_t)c->wallFaceColor[f]);
            return S_OK;
        case 5:
            if (res) tcSetObj(res, tcCached(c, TC_KIND_WALLBRUSH, f));
            return S_OK;
        default: return DISP_E_MEMBERNOTFOUND;
        }
    }

    /* ===== 第十刀：Printer / Export / Zoom —— 认名字、不动作，不静默 438 ===== */
    if (o->kind == TC_KIND_PRINTER || o->kind == TC_KIND_EXPORT || o->kind == TC_KIND_ZOOM) {
        if (o->kind == TC_KIND_PRINTER) {
            if (id == 5) { if (res) tcSetL(res, 793); return S_OK; }   /* PageWidth (A4@96dpi) */
            if (id == 6) { if (res) tcSetL(res, 1122); return S_OK; }  /* PageHeight */
            return S_OK;                                              /* 其余不动作 */
        }
        if (o->kind == TC_KIND_EXPORT) {
            /* 第十五刀：两个真动作落地（参数按 VARIANT 取，BSTR 路径）。
             * Email / Picture / Width / Height 仍不动作，台账注明。 */
            const wchar_t* path = (argc >= 1 && a0 && V_VT(a0) == VT_BSTR) ? V_BSTR(a0) : NULL;
            if (tcTrace() > 0) {
                fprintf(stderr, "[TCHART] export id=%d argc=%d res=%d vt=%d\n",
                        (int)id, argc, res ? 1 : 0,
                        (argc >= 1 && a0) ? (int)V_VT(a0) : -1);
                fflush(stderr);
            }
            switch (id) {
            case 2:            /* CopyToClipboard */
                if (res) tcSetL(res, tcExportClipboard(c) ? -1 : 0);
                else     tcExportClipboard(c);
                return S_OK;
            case 1:            /* SaveToFile：只认 .jpg/.jpeg，其余不落盘 */
                if (path && (wcsstr(path, L".jpg") || wcsstr(path, L".JPG"))) {
                    if (res) tcSetL(res, tcSaveJpeg(c, path) ? -1 : 0);
                }
                return S_OK;
            case 7:            /* SaveToJPEGFile */
                if (res) tcSetL(res, (path && tcSaveJpeg(c, path)) ? -1 : 0);
                return S_OK;
            default:
                return S_OK;
            }
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

    /* ===== 第十七刀：按类型切换的切面（AsBar / AsHorizBar / AsLine / AsCandle / AsSurface / AsArea / AsPoint）=====
     * 口径分开写清楚：只有 barPct（BarWidthPercent）真的改几何；其余存得住、读得回但不绘制 ——
     * candle 要 OHLC 四列、surface 要二维网格，我们现在一个系列只有一列 y。 */
    if (o->kind >= TC_KIND_ASBAR && o->kind <= TC_KIND_ASPOINT) {
        TcSeries* sr = tcSeries(c, o->index);
        int* slot = NULL;
        if (tcTrace() > 0) {
            fprintf(stderr, "[TCHART] facet kind=%d id=%d idx=%d put=%d\n",
                    o->kind, (int)id, o->index, (flags & DISPATCH_PROPERTYPUT) ? 1 : 0);
            fflush(stderr);
        }
        if (!sr) return DISP_E_MEMBERNOTFOUND;
        switch (o->kind) {
        case TC_KIND_ASBAR:
        case TC_KIND_ASHBAR:
            if (id == 1) slot = &sr->barPct;
            else if (id == 2) slot = &sr->asMiscA;
            else if (id == 3) slot = &sr->asMiscB;
            break;
        case TC_KIND_ASLINE:
            if (id == 1) slot = &sr->lineSmoothed;
            else if (id == 2) slot = &sr->asMiscA;
            break;
        case TC_KIND_ASCANDLE:
            if (id == 1) slot = &sr->candleWidth;
            else if (id == 2) slot = &sr->asMiscA;
            else if (id == 3) slot = &sr->asMiscB;
            break;
        case TC_KIND_ASSURF:
            if (id == 1) slot = &sr->surfWire;
            else if (id == 2) slot = &sr->dotSize;
            break;
        case TC_KIND_ASPOINT:
            if (id == 1) slot = &sr->dotSize;
            else if (id == 2) slot = &sr->asMiscA;
            break;
        default:   /* ASAREA：两个占位值 */
            if (id == 1) slot = &sr->asMiscA;
            else if (id == 2) slot = &sr->asMiscB;
            break;
        }
        if (!slot) return S_OK;
        if (flags & DISPATCH_PROPERTYPUT) {
            *slot = (int)tcArgL(a0);
            if (slot == &sr->barPct) InvalidateRect(c->hwnd, NULL, FALSE);
            return S_OK;
        }
        if (res) tcSetL(res, *slot);
        return S_OK;
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
            if (a0 && SUCCEEDED(vb6_ComChangeType(&t, a0, VT_BSTR))) b = t.bstrVal;
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
    case 13: tcSetObj(res, tcCached(c, TC_KIND_ASBAR,    o->index)); return S_OK;
    case 14: tcSetObj(res, tcCached(c, TC_KIND_ASLINE,   o->index)); return S_OK;
    case 15: tcSetObj(res, tcCached(c, TC_KIND_ASHBAR,   o->index)); return S_OK;
    case 16: tcSetObj(res, tcCached(c, TC_KIND_ASCANDLE, o->index)); return S_OK;
    case 17: tcSetObj(res, tcCached(c, TC_KIND_ASSURF,   o->index)); return S_OK;
    case 18: tcSetObj(res, tcCached(c, TC_KIND_ASAREA,   o->index)); return S_OK;
    case 19: tcSetObj(res, tcCached(c, TC_KIND_ASPOINT,  o->index)); return S_OK;
    case 20:                                    /* AddBubble(x,y,r) */
        tcSetL(res, vb6_TChart_SeriesAddBubble(o->hwnd, o->index, tcArgD(a0), tcArgD(a1), tcArgD(a2)));
        return S_OK;
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
    /* C30-10 第十九刀：`TChart1.Series(i)` 带变量下标时 cgen 发成 Call（常量下标走 GetProp）。
     * 以前这里只认 GetProp，ChartDisplay 里就成了 438。 */
    if (_wcsicmp(name, L"Series") == 0) {
        int idx = (argc >= 1) ? (int)tcArgL(a[0]) : 0;
        if (idx < 0 || idx >= c->nser) return 0;
        if (out) tcSetObj(out, tcObjFor(c, TC_KIND_SERIES, idx));
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
    /* 第十九刀：ChartDisplay 示例的三个根成员 */
    if (_wcsicmp(name, L"Frame") == 0)      { tcSetObj(out, tcCached(c, TC_KIND_FRAME, 0)); return 1; }
    if (_wcsicmp(name, L"LeftWall") == 0)   { tcSetObj(out, tcCached(c, TC_KIND_WALLFACE, 3)); return 1; }
    if (_wcsicmp(name, L"BottomWall") == 0) { tcSetObj(out, tcCached(c, TC_KIND_WALLFACE, 1)); return 1; }
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
        if (SUCCEEDED(vb6_ComChangeType(&t, v, VT_BSTR))) {
            vb6_TChart_SetHeaderText(obj, t.bstrVal ? t.bstrVal : L"");
        }
        VariantClear(&t);
        return 1;
    }
    return 0;
}
