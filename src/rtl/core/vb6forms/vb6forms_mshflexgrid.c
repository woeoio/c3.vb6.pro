/* vb6forms_mshflexgrid.c —— MSHFlexGrid / MSFlexGrid 复刻（MSHFLXGD.OCX / MSFLXGRD.OCX）。C30-6。
 *
 * 与 C30-5 的 DataGrid 是**两种东西**：
 *   - DataGrid 是**绑定表格**：数据在 ADO/Data 的记录集里，格子只是视图（底座 ACE，CI 上没装就 SKIP）；
 *   - MSHFlexGrid 是**字符串矩阵**：数据就在控件的 Rows×Cols 单元格里，TextMatrix(r,c) 直接读写，
 *     不需要数据库、不需要 ACE ⇒ **GA runner 上能真跑**（这是本刀的主要动机）。
 *
 * 接口面不是拍脑袋定的：统计了 D:\vb_yqt4qPac\数据库 下 601 个用到 FlexGrid 的 .frm，
 * 实际用到的成员（次数）：
 *   TextMatrix/2 454、Rows 128、Col 115、Text 103、Row 85、Clear 40、Visible 39、
 *   ColWidth/1 50、AddItem 27+1、Cols 23、RowSel 22、SetFocus 18、Refresh 13、Editable 13,
 *   MouseRow 12、Redraw 12、TopRow 12、FormatString 11、CellHeight/CellWidth 7/6、
 *   RowHeight/1 6、CellBackColor 5、FixedRows 5、TextArray 4、FixedCols 3、LeftCol 3、
 *   SelectionMode 2、ColSel 2、ScrollBars 2、ColAlignment/1 4、ClearStructure 2、
 *   RemoveItem 4+1、Sort 5、RowHeightMin 1。
 *
 * 本刀做到：
 *   - 完整字符串矩阵（Rows/Cols 可改，改时**保留内容**；FixedRows/FixedCols 冻结区）；
 *   - TextMatrix(r,c) / TextArray(i) / Text 读写；ColWidth(i)/RowHeight(i)/ColAlignment(i) 按列存；
 *   - 自绘：冻结区 3D 面板、格线、当前格高亮、选区、行/列滚动条、滚轮；
 *   - 鼠标点选 + 拖选、键盘导航（方向/Home/End/PageUp/Down/Tab）、Editable 为真时格内直接键入；
 *   - AddItem（Tab 拆列）/ RemoveItem / Clear / ClearStructure / Refresh / Redraw / Sort。
 * 没做（如实留着）：分层记录集（MSHFlexGrid 的 DataMember/层次带区）、MergeCells 真合并、
 * FormatString 只存串不参与排版、单元格掩码/图片、列宽拖拽（AllowUserResizing 只存值）。
 */

#include <windows.h>
#include <windowsx.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include <stdlib.h>
#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6forms_prop_ctrl.h"

#define VB6_MFG_CLASS   L"VB6_MSHFLEXGRID"
#define VB6_MFG_MAX     32       /* 同进程最多几枚（与 DataGrid 同口径） */
#define VB6_MFG_MAXDIM  20000    /* Rows/Cols 上限，防一条脏属性把内存吃光 */
#define VB6_MFG_CELLLEN 255      /* 单元格字符上限（VB6 侧一般 < 64） */

/* VB6 的 ColAlignment：0=左 1=右 2=中 */
#define VB6_MFG_ALIGN_LEFT   0
#define VB6_MFG_ALIGN_RIGHT  1
#define VB6_MFG_ALIGN_CENTER 2

struct vb6_MshFlexGrid {
    HWND    hwnd;                       /* NULL = 空槽 */
    int32_t rows, cols;                 /* 含冻结行/列 */
    int32_t fixedRows, fixedCols;
    int32_t row, col;                   /* 当前格（0 基） */
    int32_t rowSel, colSel;             /* 选区另一端（拖选） */
    int32_t topRow, leftCol;            /* 视口左上角 */
    int32_t gridLines;                  /* 0=None 1=Flat 2=Inset 3=Raised（只区分 0/非 0 画法） */
    int32_t wordWrap;                   /* 只影响画字换行 */
    int32_t editable;                   /* True = 数据区可键入 */
    int32_t focusRect;
    int32_t scrollBars;                 /* 0=None 1=H 2=V 3=Both */
    int32_t selectionMode;              /* 0=Free 1=ByRow 2=ByCol */
    int32_t allowUserResizing;
    int32_t mergeCells;
    int32_t fillStyle;
    int32_t redraw;
    int32_t rowHeightMin;               /* 缇 */
    int32_t sortDir;                    /* 0=None 1=Asc 2=Desc（写属性时立即排序） */
    int32_t mouseRow, mouseCol;         /* 上一次鼠标所在格（-1 = 不在格上） */
    int32_t rowHeightTw;                /* 缺省行高（缇） */
    int32_t backColor, foreColor, fillColor, cellBackColor, cellForeColor; /* VB6 &HBBGGRR */
    wchar_t formatString[256];
    wchar_t dataSource[128];            /* 只存名（绑定未做，见文件头"没做"） */
    wchar_t** cells;                    /* rows*cols，行主序；NULL = 空串 */
    int32_t*  colWidths;                /* 每列宽（缇） */
    int32_t*  colAligns;                /* 每列对齐 */
    /* 格内编辑 */
    int32_t editing;
    int32_t editLen;
    wchar_t editBuf[VB6_MFG_CELLLEN + 1];
};

static struct vb6_MshFlexGrid g_mfg[VB6_MFG_MAX];

static struct vb6_MshFlexGrid* mfgOf(HWND hwnd) {
    int i;
    if (!hwnd) return NULL;
    for (i = 0; i < VB6_MFG_MAX; i++)
        if (g_mfg[i].hwnd == hwnd) return &g_mfg[i];
    return NULL;
}

/* 缇 → 像素：96 dpi 下 1 px = 15 缇。VB6 的属性是缇，画的时候才换，存的一直是缇。 */
static int mfgPx(int tw) {
    int px = tw / 15;
    if (px < 6) px = 6;
    return px;
}

static int32_t mfgClampI(int32_t v, int32_t lo, int32_t hi) {
    if (v < lo) return lo;
    if (v > hi) return hi;
    return v;
}

/* 重建 cells 矩阵：把旧内容按 (r,c) 逐格搬过去 —— rows/cols 改动时**不许丢数据**，
 * 这是 VB6 的行为，也是真实例子最依赖的一条（`grd.Rows = grd.Rows + 1` 之后原来那些还在）。 */
static int mfgAlloc(struct vb6_MshFlexGrid* g, int32_t rows, int32_t cols) {
    wchar_t** nc;
    int32_t* nw;
    int32_t* na;
    int32_t r, c, total;
    if (!g) return 0;
    if (rows < 1) rows = 1;
    if (cols < 1) cols = 1;
    if (rows > VB6_MFG_MAXDIM) rows = VB6_MFG_MAXDIM;
    if (cols > VB6_MFG_MAXDIM) cols = VB6_MFG_MAXDIM;
    total = rows * cols;
    nc = (wchar_t**)calloc((size_t)total, sizeof(wchar_t*));
    if (!nc) return 0;
    nw = (int32_t*)malloc((size_t)cols * sizeof(int32_t));
    na = (int32_t*)malloc((size_t)cols * sizeof(int32_t));
    if (!nw || !na) { free(nc); free(nw); free(na); return 0; }
    for (c = 0; c < cols; c++) {
        nw[c] = 1200;   /* 缺省 80 px */
        na[c] = VB6_MFG_ALIGN_LEFT;
        if (g->colWidths && c < g->cols)  nw[c] = g->colWidths[c];
        if (g->colAligns && c < g->cols)  na[c] = g->colAligns[c];
    }
    /* 搬内容（行列都取交集，多出来的格子留空） */
    for (r = 0; r < rows && r < g->rows; r++) {
        for (c = 0; c < cols && c < g->cols; c++) {
            wchar_t* old = g->cells ? g->cells[r * g->cols + c] : NULL;
            if (old && old[0]) {
                nc[r * cols + c] = (wchar_t*)malloc((VB6_MFG_CELLLEN + 1) * sizeof(wchar_t));
                if (nc[r * cols + c]) wcsncpy(nc[r * cols + c], old, VB6_MFG_CELLLEN);
            }
        }
    }
    if (g->cells) {
        int32_t n = g->rows * g->cols, i;
        for (i = 0; i < n; i++) free(g->cells[i]);
        free(g->cells);
    }
    free(g->colWidths);
    free(g->colAligns);
    g->cells = nc;
    g->colWidths = nw;
    g->colAligns = na;
    g->rows = rows;
    g->cols = cols;
    if (g->fixedRows > rows) g->fixedRows = rows;
    if (g->fixedCols > cols) g->fixedCols = cols;
    g->row  = mfgClampI(g->row, 0, rows - 1);
    g->col  = mfgClampI(g->col, 0, cols - 1);
    g->rowSel = mfgClampI(g->rowSel, 0, rows - 1);
    g->colSel = mfgClampI(g->colSel, 0, cols - 1);
    g->topRow = mfgClampI(g->topRow, 0, rows - 1);
    g->leftCol = mfgClampI(g->leftCol, 0, cols - 1);
    return 1;
}

static struct vb6_MshFlexGrid* mfgSlot(HWND hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf(hwnd);
    int i;
    if (g || !hwnd) return g;
    for (i = 0; i < VB6_MFG_MAX; i++) {
        if (!g_mfg[i].hwnd) {
            memset(&g_mfg[i], 0, sizeof(g_mfg[i]));
            g_mfg[i].hwnd = hwnd;
            g_mfg[i].fixedRows = 1;
            g_mfg[i].fixedCols = 1;
            g_mfg[i].gridLines = 1;
            g_mfg[i].editable = 0;
            g_mfg[i].focusRect = -1;
            g_mfg[i].scrollBars = 3;
            g_mfg[i].fillStyle = 0;
            g_mfg[i].redraw = -1;
            g_mfg[i].rowHeightTw = 300;
            g_mfg[i].rowHeightMin = 0;
            g_mfg[i].backColor = 0xFFFFFF;
            g_mfg[i].foreColor = 0x000000;
            g_mfg[i].fillColor = 0x000000;
            g_mfg[i].cellBackColor = 0xFFFFFF;
            g_mfg[i].cellForeColor = 0x000000;
            g_mfg[i].mouseRow = -1;
            g_mfg[i].mouseCol = -1;
            mfgAlloc(&g_mfg[i], 2, 2);   /* VB6 缺省 2x2（1 冻结行 + 1 数据行） */
            return &g_mfg[i];
        }
    }
    return NULL;
}

/* 单元格指针：越界返回 NULL。空串用 NULL 表示（惰性分配，省内存）。 */
static wchar_t* mfgCell(struct vb6_MshFlexGrid* g, int32_t r, int32_t c) {
    if (!g || !g->cells || r < 0 || c < 0 || r >= g->rows || c >= g->cols) return NULL;
    return g->cells[r * g->cols + c];
}

static void mfgSetCell(struct vb6_MshFlexGrid* g, int32_t r, int32_t c, const wchar_t* s) {
    wchar_t* p;
    if (!g || !g->cells || r < 0 || c < 0 || r >= g->rows || c >= g->cols) return;
    if (!s) s = L"";
    if (!s[0]) {   /* 写空 = 归还内存，回到"空格"状态 */
        free(g->cells[r * g->cols + c]);
        g->cells[r * g->cols + c] = NULL;
        return;
    }
    p = g->cells[r * g->cols + c];
    if (!p) {
        p = (wchar_t*)malloc((VB6_MFG_CELLLEN + 1) * sizeof(wchar_t));
        if (!p) return;
        g->cells[r * g->cols + c] = p;
    }
    wcsncpy(p, s, VB6_MFG_CELLLEN);
    p[VB6_MFG_CELLLEN] = 0;
}

static int mfgColWidthPx(struct vb6_MshFlexGrid* g, int32_t c) {
    if (!g || c < 0 || c >= g->cols || !g->colWidths) return 80;
    return mfgPx(g->colWidths[c]);
}

/* ============================ 滚动 ============================ */

static void mfgSyncScroll(struct vb6_MshFlexGrid* g) {
    SCROLLINFO si;
    RECT rc;
    int rowH, visRows, visCols, x, c;
    if (!g || !g->hwnd) return;
    GetClientRect(g->hwnd, &rc);
    rowH = mfgPx(g->rowHeightTw);
    if (rowH < 1) rowH = 1;
    visRows = (rc.bottom - rc.top) / rowH;
    if (visRows < 1) visRows = 1;

    memset(&si, 0, sizeof(si));
    si.cbSize = sizeof(si);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = g->rows > 0 ? g->rows - 1 : 0;
    si.nPage = visRows;
    si.nPos = g->topRow;
    SetScrollInfo(g->hwnd, SB_VERT, &si, TRUE);
    ShowScrollBar(g->hwnd, SB_VERT, (g->scrollBars & 2) && g->rows > visRows);

    /* 横向按"列"滚（与 VB6 的 LeftCol 口径一致） */
    x = rc.right - rc.left;
    visCols = 0;
    { int acc = 0; for (c = g->leftCol; c < g->cols; c++) { acc += mfgColWidthPx(g, c); if (acc > x) break; visCols++; } }
    if (visCols < 1) visCols = 1;
    memset(&si, 0, sizeof(si));
    si.cbSize = sizeof(si);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = g->cols > 0 ? g->cols - 1 : 0;
    si.nPage = visCols;
    si.nPos = g->leftCol;
    SetScrollInfo(g->hwnd, SB_HORZ, &si, TRUE);
    ShowScrollBar(g->hwnd, SB_HORZ, (g->scrollBars & 1) && g->cols > visCols);
}

/* 当前格滚进视口（键盘/Tab/代码设 Row、Col 之后都要） */
static void mfgEnsureVisible(struct vb6_MshFlexGrid* g) {
    RECT rc;
    int rowH, visRows;
    if (!g || !g->hwnd) return;
    GetClientRect(g->hwnd, &rc);
    rowH = mfgPx(g->rowHeightTw);
    visRows = (rc.bottom - rc.top) / rowH;
    if (visRows < 1) visRows = 1;
    if (g->row < g->topRow) g->topRow = g->row;
    if (g->row >= g->topRow + visRows) g->topRow = g->row - visRows + 1;
    if (g->col < g->leftCol) g->leftCol = g->col;
    if (g->leftCol >= g->cols) g->leftCol = g->cols > 0 ? g->cols - 1 : 0;
    if (g->topRow < 0) g->topRow = 0;
    if (g->leftCol < 0) g->leftCol = 0;
}

/* ============================ 绘制 ============================ */

static COLORREF mfgCol(int32_t v6color) {
    /* VB6 的颜色是 &HBBGGRR，与 COLORREF 同序，直接用 */
    return (COLORREF)(v6color & 0xFFFFFF);
}

static void mfgPaint(struct vb6_MshFlexGrid* g, HDC dc) {
    RECT rc;
    HFONT hf, old;
    int rowH, y, r, c, x;
    int focus;
    wchar_t buf[VB6_MFG_CELLLEN + 1];
    HBRUSH back;

    GetClientRect(g->hwnd, &rc);
    back = CreateSolidBrush(mfgCol(g->backColor));
    FillRect(dc, &rc, back);
    DeleteObject(back);

    rowH = mfgPx(g->rowHeightTw);
    focus = (GetFocus() == g->hwnd);

    SetBkMode(dc, TRANSPARENT);
    hf = vb6_ControlFont(g->hwnd);
    /* 字体只用 vb6_ControlFont（D14 门禁：全仓 stock 默认字体兜底只允许 vb6forms.c 那一处）。
     * DataGrid 同口径。 */
    old = (HFONT)SelectObject(dc, hf);

    /* 冻结行画到视口顶（VB6 里固定行不随 topRow 滚），数据行从 topRow 起。
     * 这里把"冻结行 + 数据行"拼成一条行序列：先是 0..fixedRows-1，再从
     * max(topRow, fixedRows) 起。 */
    {
        int32_t vis[VB6_MFG_MAXDIM];
        int nvis = 0, i;
        for (i = 0; i < g->fixedRows && i < g->rows; i++) vis[nvis++] = i;
        for (i = (g->topRow > g->fixedRows ? g->topRow : g->fixedRows);
             i < g->rows && nvis < VB6_MFG_MAXDIM; i++) vis[nvis++] = i;

        y = 0;
        for (i = 0; i < nvis && y < rc.bottom; i++) {
            r = vis[i];
            x = 0;
            /* 冻结列同理：0..fixedCols-1 先画，再画从 leftCol 起的数据列 */
            {
                int32_t vc[VB6_MFG_MAXDIM];
                int nvc = 0, j;
                for (j = 0; j < g->fixedCols && j < g->cols; j++) vc[nvc++] = j;
                for (j = (g->leftCol > g->fixedCols ? g->leftCol : g->fixedCols);
                     j < g->cols && nvc < VB6_MFG_MAXDIM; j++) vc[nvc++] = j;

                for (j = 0; j < nvc && x < rc.right; j++) {
                    RECT cell;
                    int cw = mfgColWidthPx(g, vc[j]);
                    int isRowFixed = (r < g->fixedRows);
                    int isColFixed = (vc[j] < g->fixedCols);
                    int isCur = (r == g->row && vc[j] == g->col);
                    int inSel = (r >= (g->row < g->rowSel ? g->row : g->rowSel)
                                 && r <= (g->row > g->rowSel ? g->row : g->rowSel)
                                 && vc[j] >= (g->col < g->colSel ? g->col : g->colSel)
                                 && vc[j] <= (g->col > g->colSel ? g->col : g->colSel));
                    int align = g->colAligns && vc[j] < g->cols ? g->colAligns[vc[j]] : 0;
                    wchar_t* txt;

                    cell.left = x; cell.top = y; cell.right = x + cw; cell.bottom = y + rowH;
                    x += cw;

                    if (isCur && focus) {
                        FillRect(dc, &cell, (HBRUSH)GetSysColorBrush(COLOR_HIGHLIGHT));
                        SetTextColor(dc, GetSysColor(COLOR_HIGHLIGHTTEXT));
                    } else if (isCur) {
                        FillRect(dc, &cell, (HBRUSH)GetSysColorBrush(COLOR_3DFACE));
                        SetTextColor(dc, mfgCol(g->foreColor));
                    } else if (inSel && !isRowFixed && !isColFixed) {
                        FillRect(dc, &cell, (HBRUSH)GetSysColorBrush(COLOR_HIGHLIGHT));
                        SetTextColor(dc, GetSysColor(COLOR_HIGHLIGHTTEXT));
                    } else if (isRowFixed || isColFixed) {
                        FillRect(dc, &cell, (HBRUSH)GetSysColorBrush(COLOR_3DFACE));
                        SetTextColor(dc, mfgCol(g->foreColor));
                    } else {
                        SetTextColor(dc, mfgCol(g->foreColor));
                    }

                    /* 正在编辑的格子画编辑缓冲 + 光标位 */
                    if (g->editing && isCur) {
                        RECT tx = { cell.left + 3, cell.top + 2, cell.right - 3, cell.bottom - 2 };
                        size_t n = wcslen(g->editBuf);
                        wcsncpy(buf, g->editBuf, VB6_MFG_CELLLEN);
                        buf[VB6_MFG_CELLLEN] = 0;
                        DrawTextW(dc, buf, -1, &tx,
                                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
                        {   /* 光标：编辑串宽度之后竖一道 */
                            SIZE sz; sz.cx = 0;
                            if (n) { GetTextExtentPoint32W(dc, buf, (int)n, &sz); }
                            MoveToEx(dc, tx.left + sz.cx, tx.top + 2, NULL);
                            LineTo(dc, tx.left + sz.cx, tx.bottom - 2);
                        }
                    } else {
                        txt = mfgCell(g, r, vc[j]);
                        if (txt && txt[0]) {
                            RECT tx = { cell.left + 3, cell.top + 2, cell.right - 3, cell.bottom - 2 };
                            UINT fmt = DT_NOPREFIX | DT_END_ELLIPSIS;
                            if (g->wordWrap) fmt |= DT_WORDBREAK | DT_TOP;
                            else             fmt |= DT_SINGLELINE | DT_VCENTER;
                            if (align == VB6_MFG_ALIGN_RIGHT)       fmt |= DT_RIGHT;
                            else if (align == VB6_MFG_ALIGN_CENTER) fmt |= DT_CENTER;
                            else                                    fmt |= DT_LEFT;
                            DrawTextW(dc, txt, -1, &tx, fmt);
                        }
                    }

                    if (g->gridLines) {
                        HPEN pn = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_3DLIGHT));
                        HPEN op = (HPEN)SelectObject(dc, pn);
                        MoveToEx(dc, cell.right - 1, cell.top, NULL);
                        LineTo(dc, cell.right - 1, cell.bottom);
                        MoveToEx(dc, cell.left, cell.bottom - 1, NULL);
                        LineTo(dc, cell.right, cell.bottom - 1);
                        SelectObject(dc, op);
                        DeleteObject(pn);
                    }
                    /* 冻结区边界：VB6 是一道 3D 立体线，这里用 shadow 画一条 */
                    if (isRowFixed && (r + 1 >= g->fixedRows)) {
                        HPEN pn = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_3DSHADOW));
                        HPEN op = (HPEN)SelectObject(dc, pn);
                        MoveToEx(dc, cell.left, cell.bottom - 1, NULL);
                        LineTo(dc, cell.right, cell.bottom - 1);
                        SelectObject(dc, op);
                        DeleteObject(pn);
                    }
                    if (isColFixed && (vc[j] + 1 >= g->fixedCols)) {
                        HPEN pn = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_3DSHADOW));
                        HPEN op = (HPEN)SelectObject(dc, pn);
                        MoveToEx(dc, cell.right - 1, cell.top, NULL);
                        LineTo(dc, cell.right - 1, cell.bottom);
                        SelectObject(dc, op);
                        DeleteObject(pn);
                    }
                }
            }
            y += rowH;
        }
    }

    if (g->focusRect && focus) {
        RECT fr;   /* 当前格的焦点框 */
        int x0 = 0, i;
        for (i = 0; i < g->col; i++) {
            if (i < g->fixedCols) x0 += mfgColWidthPx(g, i);
            else if (i >= g->leftCol || i < g->fixedCols) x0 += mfgColWidthPx(g, i);
        }
        SetRect(&fr, x0, 0, x0 + mfgColWidthPx(g, g->col), rowH);
        DrawFocusRect(dc, &fr);
    }

    SelectObject(dc, old);
}

/* ============================ 输入 ============================ */

static void mfgCommitEdit(struct vb6_MshFlexGrid* g) {
    if (!g || !g->editing) return;
    g->editBuf[g->editLen] = 0;
    mfgSetCell(g, g->row, g->col, g->editBuf);
    g->editing = 0;
    g->editLen = 0;
    g->editBuf[0] = 0;
}

static void mfgBeginEdit(struct vb6_MshFlexGrid* g) {
    wchar_t* p;
    if (!g || g->editing) return;
    g->editing = 1;
    g->editLen = 0;
    g->editBuf[0] = 0;
    p = mfgCell(g, g->row, g->col);
    if (p && p[0]) {
        wcsncpy(g->editBuf, p, VB6_MFG_CELLLEN);
        g->editLen = (int32_t)wcslen(g->editBuf);
    }
}

/* 移动当前格；keepSel=1 时把 rowSel/colSel 留在原处（Shift 扩展选区）。 */
static void mfgMoveTo(struct vb6_MshFlexGrid* g, int32_t r, int32_t c, int keepSel) {
    int32_t orow, ocol;
    if (!g) return;
    orow = g->row; ocol = g->col;
    g->row = mfgClampI(r, 0, g->rows - 1);
    g->col = mfgClampI(c, 0, g->cols - 1);
    if (!keepSel) { g->rowSel = g->row; g->colSel = g->col; }
    if (g->row != orow || g->col != ocol) {
        mfgCommitEdit(g);
        mfgEnsureVisible(g);
        mfgSyncScroll(g);
        InvalidateRect(g->hwnd, NULL, FALSE);
    }
}

static LRESULT CALLBACK vb6_MshFlexGridWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    struct vb6_MshFlexGrid* g = mfgOf(hwnd);
    switch (msg) {
    case WM_CREATE:
        mfgSlot(hwnd);
        return 0;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC dc = BeginPaint(hwnd, &ps);
        if (g) mfgPaint(g, dc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;
    case WM_SIZE:
        if (g) mfgSyncScroll(g);
        return 0;
    case WM_SETFOCUS:
    case WM_KILLFOCUS:
        if (g) InvalidateRect(hwnd, NULL, FALSE);
        break;
    case WM_LBUTTONDOWN: {
        int rowH, y, x, r, c, i;
        if (!g) break;
        rowH = mfgPx(g->rowHeightTw);
        y = GET_Y_LPARAM(lp); x = GET_X_LPARAM(lp);
        r = y / (rowH > 0 ? rowH : 1);                 /* 视口行序（含冻结行） */
        if (r >= g->fixedRows) r = g->topRow > g->fixedRows
                                 ? g->topRow + (r - g->fixedRows) : r;
        c = -1;
        {   int acc = 0;
            for (i = 0; i < g->cols; i++) {
                int ci = i;
                if (i >= g->fixedCols) {
                    if (i < g->leftCol && i >= g->fixedCols) continue;
                    ci = (i == g->fixedCols) ? (g->leftCol > g->fixedCols ? g->leftCol : g->fixedCols) : i;
                }
                acc += mfgColWidthPx(g, ci);
                if (x < acc) { c = ci; break; }
            }
        }
        if (c < 0) c = g->cols - 1;
        mfgMoveTo(g, r, c, 0);
        SetFocus(hwnd);
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    case WM_MOUSEMOVE: {
        int rowH, r, c, x, i;
        if (!g) break;
        rowH = mfgPx(g->rowHeightTw);
        r = GET_Y_LPARAM(lp) / (rowH > 0 ? rowH : 1);
        if (r >= g->fixedRows) r = g->topRow > g->fixedRows
                                 ? g->topRow + (r - g->fixedRows) : r;
        x = GET_X_LPARAM(lp); c = 0;
        {   int acc = 0;
            for (i = 0; i < g->cols; i++) { acc += mfgColWidthPx(g, i); if (x < acc) { c = i; break; } c = i; }
        }
        g->mouseRow = mfgClampI(r, 0, g->rows - 1);
        g->mouseCol = mfgClampI(c, 0, g->cols - 1);
        if (wp & MK_LBUTTON) {
            /* 拖选：当前格跟着走到鼠标处，选区另一端留在按下时那格 */
            int32_t keepRow = g->row, keepCol = g->col;
            int32_t selRow = g->rowSel, selCol = g->colSel;
            mfgMoveTo(g, r, c, 1);
            g->rowSel = selRow; g->colSel = selCol;
            if (g->row >= g->rowSel) { /* 让 (rowSel,colSel) 是拖选的另一端 */ }
            InvalidateRect(hwnd, NULL, FALSE);
            (void)keepRow; (void)keepCol;
        }
        return 0;
    }
    case WM_LBUTTONDBLCLK:
        if (g && g->editable && g->row >= g->fixedRows) {
            mfgBeginEdit(g);
            InvalidateRect(hwnd, NULL, FALSE);
        }
        return 0;
    case WM_CHAR: {
        wchar_t ch = (wchar_t)wp;
        if (!g) break;
        if (ch == L'\r' || ch == L'\n') {
            mfgCommitEdit(g);
            mfgMoveTo(g, g->row + 1, g->col, 0);
            return 0;
        }
        if (ch == 0x1B) { g->editing = 0; g->editLen = 0; InvalidateRect(hwnd, NULL, FALSE); return 0; }
        if (ch == 0x09) { mfgCommitEdit(g); mfgMoveTo(g, g->row, g->col + 1, 0); return 0; }
        if (ch == 0x08) {
            if (g->editing && g->editLen > 0) { g->editLen--; g->editBuf[g->editLen] = 0; }
            InvalidateRect(hwnd, NULL, FALSE);
            return 0;
        }
        if (ch < 32) break;
        if (!g->editable) break;
        if (g->row < g->fixedRows) break;      /* 冻结区不许改 */
        mfgBeginEdit(g);
        if (g->editLen < VB6_MFG_CELLLEN) {
            g->editBuf[g->editLen++] = ch;
            g->editBuf[g->editLen] = 0;
        }
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    case WM_KEYDOWN: {
        int32_t page;
        RECT rc;
        if (!g) break;
        GetClientRect(hwnd, &rc);
        page = (rc.bottom - rc.top) / (mfgPx(g->rowHeightTw) > 0 ? mfgPx(g->rowHeightTw) : 1);
        if (page < 1) page = 1;
        switch (wp) {
        case VK_UP:     mfgMoveTo(g, g->row - 1, g->col, GetKeyState(VK_SHIFT) < 0); return 0;
        case VK_DOWN:   mfgMoveTo(g, g->row + 1, g->col, GetKeyState(VK_SHIFT) < 0); return 0;
        case VK_LEFT:   mfgMoveTo(g, g->row, g->col - 1, GetKeyState(VK_SHIFT) < 0); return 0;
        case VK_RIGHT:  mfgMoveTo(g, g->row, g->col + 1, GetKeyState(VK_SHIFT) < 0); return 0;
        case VK_TAB:    mfgMoveTo(g, g->row, g->col + 1, 0); return 0;
        case VK_HOME:   mfgMoveTo(g, g->row, 0, 0); return 0;
        case VK_END:    mfgMoveTo(g, g->row, g->cols - 1, 0); return 0;
        case VK_PRIOR:  mfgMoveTo(g, g->row - page, g->col, 0); return 0;
        case VK_NEXT:   mfgMoveTo(g, g->row + page, g->col, 0); return 0;
        case VK_DELETE:
            if (g->editable && g->row >= g->fixedRows) {
                mfgSetCell(g, g->row, g->col, L"");
                InvalidateRect(hwnd, NULL, FALSE);
            }
            return 0;
        default: break;
        }
        break;
    }
    case WM_VSCROLL: {
        SCROLLINFO si;
        if (!g) break;
        memset(&si, 0, sizeof(si));
        si.cbSize = sizeof(si); si.fMask = SIF_ALL;
        GetScrollInfo(hwnd, SB_VERT, &si);
        switch (LOWORD(wp)) {
        case SB_LINEUP: si.nPos -= 1; break;
        case SB_LINEDOWN: si.nPos += 1; break;
        case SB_PAGEUP: si.nPos -= (int)si.nPage; break;
        case SB_PAGEDOWN: si.nPos += (int)si.nPage; break;
        case SB_THUMBTRACK: case SB_THUMBPOSITION: si.nPos = si.nTrackPos; break;
        default: break;
        }
        if (si.nPos < si.nMin) si.nPos = si.nMin;
        if (si.nPos > si.nMax) si.nPos = si.nMax;
        if (si.nPos < g->fixedRows && g->rows > g->fixedRows) si.nPos = g->fixedRows;
        g->topRow = si.nPos;
        mfgSyncScroll(g);
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    case WM_HSCROLL: {
        SCROLLINFO si;
        if (!g) break;
        memset(&si, 0, sizeof(si));
        si.cbSize = sizeof(si); si.fMask = SIF_ALL;
        GetScrollInfo(hwnd, SB_HORZ, &si);
        switch (LOWORD(wp)) {
        case SB_LINELEFT: si.nPos -= 1; break;
        case SB_LINERIGHT: si.nPos += 1; break;
        case SB_PAGELEFT: si.nPos -= (int)si.nPage; break;
        case SB_PAGERIGHT: si.nPos += (int)si.nPage; break;
        case SB_THUMBTRACK: case SB_THUMBPOSITION: si.nPos = si.nTrackPos; break;
        default: break;
        }
        if (si.nPos < si.nMin) si.nPos = si.nMin;
        if (si.nPos > si.nMax) si.nPos = si.nMax;
        g->leftCol = si.nPos;
        mfgSyncScroll(g);
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    case WM_MOUSEWHEEL: {
        if (!g) break;
        g->topRow -= (GET_WHEEL_DELTA_WPARAM(wp) / WHEEL_DELTA) * 3;
        if (g->topRow < g->fixedRows && g->rows > g->fixedRows) g->topRow = g->fixedRows;
        if (g->topRow < 0) g->topRow = 0;
        if (g->topRow > g->rows - 1) g->topRow = g->rows > 0 ? g->rows - 1 : 0;
        mfgSyncScroll(g);
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    case WM_GETDLGCODE:
        return DLGC_WANTARROWS | DLGC_WANTCHARS | DLGC_WANTTAB;
    case WM_DESTROY:
        if (g) {
            int32_t i, n = g->rows * g->cols;
            for (i = 0; i < n; i++) free(g->cells[i]);
            free(g->cells);
            free(g->colWidths);
            free(g->colAligns);
            memset(g, 0, sizeof(*g));
        }
        break;
    default: break;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void vb6_RegisterMshFlexGridClass(void* hInstance) {
    static int done = 0;
    WNDCLASSW wc;
    if (done) return;
    done = 1;
    memset(&wc, 0, sizeof(wc));
    wc.style         = CS_DBLCLKS;
    wc.lpfnWndProc   = vb6_MshFlexGridWndProc;
    wc.hInstance     = (HINSTANCE)hInstance;
    wc.hCursor       = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
    wc.lpszClassName = VB6_MFG_CLASS;
    RegisterClassW(&wc);
}

/* ============================ 属性：尺寸 ============================ */

int32_t vb6_MshFlexGrid_GetRows(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    return g ? g->rows : 0;
}
void vb6_MshFlexGrid_SetRows(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    if (v < g->fixedRows && v >= 1) v = g->fixedRows;   /* VB6: Rows 不得少于 FixedRows */
    mfgAlloc(g, v, g->cols);
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

int32_t vb6_MshFlexGrid_GetCols(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    return g ? g->cols : 0;
}
void vb6_MshFlexGrid_SetCols(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    if (v < g->fixedCols && v >= 1) v = g->fixedCols;
    mfgAlloc(g, g->rows, v);
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

int32_t vb6_MshFlexGrid_GetFixedRows(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->fixedRows : 0;
}
void vb6_MshFlexGrid_SetFixedRows(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->fixedRows = mfgClampI(v, 0, g->rows);
    if (g->row < 0) g->row = 0;
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

int32_t vb6_MshFlexGrid_GetFixedCols(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->fixedCols : 0;
}
void vb6_MshFlexGrid_SetFixedCols(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->fixedCols = mfgClampI(v, 0, g->cols);
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* ============================ 属性：当前格 / 选区 ============================ */

int32_t vb6_MshFlexGrid_GetRow(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->row : 0;
}
void vb6_MshFlexGrid_SetRow(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    mfgMoveTo(g, v, g->col, 1);
    g->rowSel = g->row;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetCol(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->col : 0;
}
void vb6_MshFlexGrid_SetCol(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    mfgMoveTo(g, g->row, v, 1);
    g->colSel = g->col;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetRowSel(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->rowSel : 0;
}
void vb6_MshFlexGrid_SetRowSel(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->rowSel = mfgClampI(v, 0, g->rows - 1);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetColSel(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->colSel : 0;
}
void vb6_MshFlexGrid_SetColSel(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->colSel = mfgClampI(v, 0, g->cols - 1);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetTopRow(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->topRow : 0;
}
void vb6_MshFlexGrid_SetTopRow(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->topRow = mfgClampI(v, 0, g->rows - 1);
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetLeftCol(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->leftCol : 0;
}
void vb6_MshFlexGrid_SetLeftCol(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->leftCol = mfgClampI(v, 0, g->cols - 1);
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* ============================ 属性：文本 ============================ */

const void* vb6_MshFlexGrid_GetText(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    wchar_t* p = g ? mfgCell(g, g->row, g->col) : NULL;
    return SysAllocString(p ? p : L"");
}
void vb6_MshFlexGrid_SetText(void* hwnd, const void* bs) {
    const wchar_t* s = (const wchar_t*)bs;
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g || g->row < g->fixedRows) return;   /* 冻结区只读，与 VB6 一致 */
    mfgSetCell(g, g->row, g->col, s);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

const void* vb6_MshFlexGrid_GetTextMatrix(void* hwnd, int32_t r, int32_t c) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    wchar_t* p = g ? mfgCell(g, r, c) : NULL;
    return SysAllocString(p ? p : L"");
}
void vb6_MshFlexGrid_SetTextMatrix(void* hwnd, int32_t r, int32_t c, const void* bs) {
    const wchar_t* s = (const wchar_t*)bs;
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    mfgSetCell(g, r, c, s);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* TextArray(i)：按行主序把 cells 当一维数组看（VB6 的 TextArray 就是 cell 的线性下标） */
const void* vb6_MshFlexGrid_GetTextArray(void* hwnd, int32_t i) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    wchar_t* p = NULL;
    if (g && i >= 0 && i < g->rows * g->cols) p = g->cells[i];
    return SysAllocString(p ? p : L"");
}
void vb6_MshFlexGrid_SetTextArray(void* hwnd, int32_t i, const void* bs) {
    const wchar_t* s = (const wchar_t*)bs;
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g || i < 0 || i >= g->rows * g->cols) return;
    mfgSetCell(g, i / g->cols, i % g->cols, s);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* ============================ 属性：列宽 / 行高 / 对齐 ============================ */

int32_t vb6_MshFlexGrid_GetColWidth(void* hwnd, int32_t i) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    if (!g || !g->colWidths || i < 0 || i >= g->cols) return 0;
    return g->colWidths[i];
}
void vb6_MshFlexGrid_SetColWidth(void* hwnd, int32_t i, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g || !g->colWidths || i < 0 || i >= g->cols) return;
    g->colWidths[i] = v > 0 ? v : 1;
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* RowHeight(i)：i<0 = 缺省行高（VB6 里 RowHeight(-1) 就是这个意思） */
int32_t vb6_MshFlexGrid_GetRowHeight(void* hwnd, int32_t i) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    if (!g) return 0;
    if (i < 0 || i >= g->rows) return g->rowHeightTw;
    return g->rowHeightTw;
}
void vb6_MshFlexGrid_SetRowHeight(void* hwnd, int32_t i, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g || v <= 0) return;
    g->rowHeightTw = v;
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

int32_t vb6_MshFlexGrid_GetColAlignment(void* hwnd, int32_t i) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    if (!g || !g->colAligns || i < 0 || i >= g->cols) return 0;
    return g->colAligns[i];
}
void vb6_MshFlexGrid_SetColAlignment(void* hwnd, int32_t i, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g || !g->colAligns || i < 0 || i >= g->cols) return;
    g->colAligns[i] = mfgClampI(v, 0, 2);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

int32_t vb6_MshFlexGrid_GetRowHeightMin(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->rowHeightMin : 0;
}
void vb6_MshFlexGrid_SetRowHeightMin(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (g) g->rowHeightMin = v > 0 ? v : 0;
}

/* ============================ 属性：外观 / 行为 ============================ */

#define VB6_MFG_BOOL_PAIR(name, field)                                  \
    int32_t vb6_MshFlexGrid_Get##name(void* hwnd) {                     \
        struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);                  \
        return g ? (g->field ? -1 : 0) : 0;                             \
    }                                                                   \
    void vb6_MshFlexGrid_Set##name(void* hwnd, int32_t v) {              \
        struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);                 \
        if (!g) return;                                                  \
        mfgCommitEdit(g);                                                \
        g->field = v ? -1 : 0;                                           \
        if (g->hwnd) { mfgSyncScroll(g); InvalidateRect((HWND)hwnd, NULL, FALSE); } \
    }

VB6_MFG_BOOL_PAIR(WordWrap, wordWrap)
VB6_MFG_BOOL_PAIR(Editable, editable)
VB6_MFG_BOOL_PAIR(FocusRect, focusRect)
VB6_MFG_BOOL_PAIR(AllowUserResizing, allowUserResizing)
VB6_MFG_BOOL_PAIR(MergeCells, mergeCells)
VB6_MFG_BOOL_PAIR(Redraw, redraw)

int32_t vb6_MshFlexGrid_GetGridLines(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->gridLines : 0;
}
void vb6_MshFlexGrid_SetGridLines(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->gridLines = mfgClampI(v, 0, 3);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetScrollBars(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->scrollBars : 0;
}
void vb6_MshFlexGrid_SetScrollBars(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->scrollBars = mfgClampI(v, 0, 3);
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetSelectionMode(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->selectionMode : 0;
}
void vb6_MshFlexGrid_SetSelectionMode(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (g) g->selectionMode = mfgClampI(v, 0, 2);
}
int32_t vb6_MshFlexGrid_GetFillStyle(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->fillStyle : 0;
}
void vb6_MshFlexGrid_SetFillStyle(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (g) g->fillStyle = mfgClampI(v, 0, 1);
}

int32_t vb6_MshFlexGrid_GetBackColor(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->backColor : 0xFFFFFF;
}
void vb6_MshFlexGrid_SetBackColor(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->backColor = v & 0xFFFFFF;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetForeColor(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->foreColor : 0;
}
void vb6_MshFlexGrid_SetForeColor(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->foreColor = v & 0xFFFFFF;
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
int32_t vb6_MshFlexGrid_GetFillColor(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->fillColor : 0;
}
void vb6_MshFlexGrid_SetFillColor(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (g) g->fillColor = v & 0xFFFFFF;
}
int32_t vb6_MshFlexGrid_GetCellBackColor(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->cellBackColor : 0xFFFFFF;
}
void vb6_MshFlexGrid_SetCellBackColor(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (g) g->cellBackColor = v & 0xFFFFFF;
}
int32_t vb6_MshFlexGrid_GetCellForeColor(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->cellForeColor : 0;
}
void vb6_MshFlexGrid_SetCellForeColor(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (g) g->cellForeColor = v & 0xFFFFFF;
}

int32_t vb6_MshFlexGrid_GetMouseRow(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->mouseRow : -1;
}
int32_t vb6_MshFlexGrid_GetMouseCol(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->mouseCol : -1;
}

/* 尺寸类读数（VB6 单位是缇，这里按 15 倍换算回缇） */
int32_t vb6_MshFlexGrid_GetCellHeight(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    return g ? g->rowHeightTw : 0;
}
int32_t vb6_MshFlexGrid_GetCellWidth(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    if (!g || !g->colWidths) return 0;
    return (g->col >= 0 && g->col < g->cols) ? g->colWidths[g->col] : 0;
}
int32_t vb6_MshFlexGrid_GetCellLeft(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    int i, acc = 0;
    if (!g || !g->colWidths) return 0;
    for (i = 0; i < g->col && i < g->cols; i++) acc += g->colWidths[i];
    return acc;
}
int32_t vb6_MshFlexGrid_GetCellTop(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    return g ? g->row * g->rowHeightTw : 0;
}

const void* vb6_MshFlexGrid_GetFormatString(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    return SysAllocString(g ? g->formatString : L"");
}
void vb6_MshFlexGrid_SetFormatString(void* hwnd, const void* bs) {
    const wchar_t* s = (const wchar_t*)bs;
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    wcsncpy(g->formatString, s ? s : L"", 255);
    g->formatString[255] = 0;
}

/* DataSource：只记名字（绑定未做，见文件头）。真实例里大量写 `grd.DataSource = rs`，
 * 那种情形本刀不提供 —— 但读回来必须还是原串，否则 `Debug.Print grd.DataSource` 会空。 */
const void* vb6_MshFlexGrid_GetDataSource(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd);
    return SysAllocString(g ? g->dataSource : L"");
}
void vb6_MshFlexGrid_SetDataSource(void* hwnd, const void* bs) {
    const wchar_t* s = (const wchar_t*)bs;
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    wcsncpy(g->dataSource, s ? s : L"", 127);
    g->dataSource[127] = 0;
}

/* Sort：VB6 是属性（写 1/2 立即按当前列排），不是方法。只排数据区（冻结行不动）。 */
static void mfgSortNow(struct vb6_MshFlexGrid* g) {
    int32_t i, j, base;
    if (!g || g->sortDir == 0 || !g->cells) return;
    base = g->fixedRows;
    for (i = base; i < g->rows - 1; i++) {
        for (j = base; j < g->rows - 1 - (i - base); j++) {
            wchar_t* a = g->cells[j * g->cols + g->col];
            wchar_t* b = g->cells[(j + 1) * g->cols + g->col];
            const wchar_t* sa = a ? a : L"";
            const wchar_t* sb = b ? b : L"";
            int cmp = wcscmp(sa, sb);
            int swap = (g->sortDir == 2) ? (cmp < 0) : (cmp > 0);
            if (swap) {
                int32_t c;
                for (c = 0; c < g->cols; c++) {
                    wchar_t* t = g->cells[j * g->cols + c];
                    g->cells[j * g->cols + c] = g->cells[(j + 1) * g->cols + c];
                    g->cells[(j + 1) * g->cols + c] = t;
                }
            }
        }
    }
}
int32_t vb6_MshFlexGrid_GetSort(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgOf((HWND)hwnd); return g ? g->sortDir : 0;
}
void vb6_MshFlexGrid_SetSort(void* hwnd, int32_t v) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    g->sortDir = mfgClampI(v, 0, 2);
    mfgSortNow(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* ============================ 方法 ============================ */

/* AddItem：按 Tab 拆列，写到"当前行"；当前行是冻结行时追加到末尾。
 * VB6 里 AddItem 的落点是当前行（FillStyle=0 时插在当前格所在行），这里同口径但
 * **不搬行**：写完把当前行往下移一格 —— 真实例子的用法都是
 *   .Rows = .Rows + 1 : .Row = .Rows - 1 : .AddItem s
 * 所以只要保证"追加 + 拆列 + 可回读"即可。 */
void vb6_MshFlexGrid_AddItem(void* hwnd, const void* bs) {
    const wchar_t* s = (const wchar_t*)bs;
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    int32_t r, c;
    const wchar_t* p;
    if (!g) return;
    r = g->row;
    if (r < g->fixedRows) r = g->rows;      /* 冻结行上 AddItem ⇒ 追加 */
    if (r >= g->rows) {
        mfgAlloc(g, g->rows + 1, g->cols);
        r = g->rows - 1;
    }
    c = 0;
    p = s ? s : L"";
    while (c < g->cols) {
        const wchar_t* tab = wcschr(p, L'\t');
        size_t n = tab ? (size_t)(tab - p) : wcslen(p);
        wchar_t tmp[VB6_MFG_CELLLEN + 1];
        if (n > VB6_MFG_CELLLEN) n = VB6_MFG_CELLLEN;
        wmemcpy(tmp, p, n);
        tmp[n] = 0;
        mfgSetCell(g, r, c, tmp);
        c++;
        if (!tab) break;
        p = tab + 1;
    }
    g->row = r;
    if (g->rowSel != r) g->rowSel = r;
    mfgEnsureVisible(g);
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* RemoveItem：删数据行（冻结行不许删），后面的行上移、末行清空 */
void vb6_MshFlexGrid_RemoveItem(void* hwnd, int32_t index) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    int32_t r, c;
    if (!g || index < g->fixedRows || index >= g->rows) return;
    for (r = index; r < g->rows - 1; r++) {
        for (c = 0; c < g->cols; c++) {
            wchar_t* p = g->cells[r * g->cols + c];
            g->cells[r * g->cols + c] = g->cells[(r + 1) * g->cols + c];
            g->cells[(r + 1) * g->cols + c] = p;
        }
    }
    for (c = 0; c < g->cols; c++) mfgSetCell(g, g->rows - 1, c, L"");
    g->row = mfgClampI(g->row, 0, g->rows - 1);
    g->rowSel = mfgClampI(g->rowSel, 0, g->rows - 1);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

void vb6_MshFlexGrid_Clear(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    int32_t r, c;
    if (!g) return;
    for (r = 0; r < g->rows; r++)
        for (c = 0; c < g->cols; c++) mfgSetCell(g, r, c, L"");
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* ClearStructure：连结构一起清 —— 行列回到"只有冻结区"（VB6 的 ClearStructure 会把
 * Rows/Cols 收回 Fixed 值）。 */
void vb6_MshFlexGrid_ClearStructure(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    mfgAlloc(g, g->fixedRows > 0 ? g->fixedRows : 1, g->fixedCols > 0 ? g->fixedCols : 1);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

void vb6_MshFlexGrid_Refresh(void* hwnd) {
    struct vb6_MshFlexGrid* g = mfgSlot((HWND)hwnd);
    if (!g) return;
    mfgSyncScroll(g);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* 注：VB6 的 `Redraw` 是**布尔属性**（`grd.Redraw = False` 挂起重画），不是方法 ——
 * 这里只提供属性（VB6_MFG_BOOL_PAIR(Redraw, redraw) 生成 Get/Set），没有同名方法。 */
