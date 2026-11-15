/* vb6forms_grid.c —— DataGrid（Microsoft DataGrid Control 6.0，MSDATGRD.OCX）复刻。C30-5。
 *
 * 自绘表格，不加载 MSDATGRD.OCX。数据从**引擎**来：绑定的对象是 Data 控件那一族的
 * 隐藏引擎窗（Adodc 由 vb6_Adodc_Engine 取，Data 控件本身就是引擎），
 * 单元格取值走 vb6_Data_FieldValueAt（不动调用者当前行的任意行读）。
 *
 * 这一版做到"看得见、点得到、读得回"：
 *   - 列头（ColumnHeaders）+ 格线 + 选中格高亮；
 *   - 竖向滚动（WS_VSCROLL，行粒度）；
 *   - 点格子 = 选中该格并把**源记录集**移到那一行（VB6 的 DataGrid 就是这样带路的）；
 *   - Row / Col / RowCount / ColCount / Text 可读可写（Row/Col 越界按 VB6 钳住）。
 * 编辑（AllowUpdate 真写回、AllowAddNew/AllowDelete 真增删）**没做** —— 见 ai/029 C30-5 的"未做"。
 */

#include <windows.h>
#include <windowsx.h>   /* GET_X_LPARAM / GET_Y_LPARAM */
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6forms_prop_ctrl.h"

#define VB6_GRID_CLASS L"VB6_DATAGRID"
#define VB6_GRID_MAX   32
#define VB6_GRID_MAXCOL 32

struct vb6_Grid {
    HWND      hwnd;                 // NULL = 空槽
    HWND      src;                  // 引擎窗（Data 族 API 的入参）
    wchar_t   dataSource[128];      // 设计期写的那一串（原样回读）
    int32_t   rows, cols;
    int32_t   rowSel, colSel;       // 0 基
    int32_t   topRow;               // 视口第一行（0 基）
    int32_t   headers;              // ColumnHeaders
    int32_t   allowAddNew, allowUpdate, allowDelete;
    int32_t   rowHeightTw;          // 存**缇**（VB6 口径），画的时候 /15 换像素
    int32_t   colWidthTw;           // 同上
};

static struct vb6_Grid g_grid[VB6_GRID_MAX];

static struct vb6_Grid* gridOf(HWND hwnd) {
    int i;
    if (!hwnd) return NULL;
    for (i = 0; i < VB6_GRID_MAX; i++)
        if (g_grid[i].hwnd == hwnd) return &g_grid[i];
    return NULL;
}

static struct vb6_Grid* gridSlot(HWND hwnd) {
    struct vb6_Grid* g = gridOf(hwnd);
    int i;
    if (g || !hwnd) return g;
    for (i = 0; i < VB6_GRID_MAX; i++) {
        if (!g_grid[i].hwnd) {
            memset(&g_grid[i], 0, sizeof(g_grid[i]));
            g_grid[i].hwnd = hwnd;
            g_grid[i].headers = -1;        // VB6 缺省 True
            g_grid[i].allowAddNew = 0;
            g_grid[i].allowUpdate = -1;    // VB6 缺省 True
            g_grid[i].allowDelete = 0;
            g_grid[i].rowHeightTw = 300;   // 20 px @96dpi
            g_grid[i].colWidthTw = 1200;   // 80 px
            return &g_grid[i];
        }
    }
    return NULL;
}

// 缇 → 像素（96 dpi 下 1 px = 15 缇）
static int gridPx(int tw) {
    int px = tw / 15;
    if (px < 8) px = 8;
    return px;
}

static void gridSyncCounts(struct vb6_Grid* g) {
    if (!g) return;
    if (!g->src) { g->rows = 0; g->cols = 0; return; }
    g->rows = vb6_Data_RecordCount(g->src);
    g->cols = vb6_Data_FieldCount(g->src);
    if (g->cols > VB6_GRID_MAXCOL) g->cols = VB6_GRID_MAXCOL;
}

static void gridClampSel(struct vb6_Grid* g) {
    if (g->rowSel < 0) g->rowSel = 0;
    if (g->rowSel > g->rows - 1) g->rowSel = g->rows > 0 ? g->rows - 1 : 0;
    if (g->colSel < 0) g->colSel = 0;
    if (g->colSel > g->cols - 1) g->colSel = g->cols > 0 ? g->cols - 1 : 0;
    if (g->topRow < 0) g->topRow = 0;
}

// 点选之后把源记录集移到那一行 —— 与 VB6 一致：表格带路，Adodc1.Recordset 跟着走。
static void gridSeekSource(struct vb6_Grid* g) {
    int i;
    if (!g || !g->src) return;
    if (g->rowSel + 1 < 1) return;
    vb6_Data_MoveFirst(g->src);
    for (i = 0; i < g->rowSel; i++) vb6_Data_MoveNext(g->src);
}

static void gridPaint(struct vb6_Grid* g, HDC dc, RECT rcClient) {
    RECT rc;
    HFONT hf, old;
    int rowH, colW, headerH, y, r, c, x;
    wchar_t buf[256];
    GetClientRect(g->hwnd, &rc);
    // 底：白
    FillRect(dc, &rc, (HBRUSH)GetStockObject(WHITE_BRUSH));
    rowH = gridPx(g->rowHeightTw);
    colW = gridPx(g->colWidthTw);
    headerH = g->headers ? rowH : 0;

    SetBkMode(dc, TRANSPARENT);
    hf = vb6_ControlFont(g->hwnd);
    old = (HFONT)SelectObject(dc, hf);

    // 列头
    if (headerH > 0) {
        RECT hr = { 0, 0, rc.right, headerH };
        FillRect(dc, &hr, (HBRUSH)GetSysColorBrush(COLOR_3DFACE));
        for (c = 0; c < g->cols; c++) {
            RECT cell = { c * colW, 0, (c + 1) * colW, headerH };
            if (cell.left >= rc.right) break;
            SetRect(&cell, cell.left + 3, cell.top + 2, cell.right - 3, cell.bottom - 2);
            DrawTextW(dc, vb6_Data_FieldName(g->src, c + 1), -1, &cell,
                      DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
        }
    }

    // 数据行
    y = headerH;
    for (r = g->topRow; r < g->rows && y < rc.bottom; r++, y += rowH) {
        for (c = 0; c < g->cols; c++) {
            RECT cell = { c * colW, y, (c + 1) * colW, y + rowH };
            if (cell.left >= rc.right) break;
            if (r == g->rowSel && c == g->colSel) {
                FillRect(dc, &cell, (HBRUSH)GetSysColorBrush(COLOR_HIGHLIGHT));
                SetTextColor(dc, GetSysColor(COLOR_HIGHLIGHTTEXT));
            } else {
                SetTextColor(dc, GetSysColor(COLOR_WINDOWTEXT));
            }
            vb6_Data_FieldValueAt(g->src, r + 1, c + 1, buf, 256);
            {
                RECT tx = { cell.left + 3, cell.top + 2, cell.right - 3, cell.bottom - 2 };
                DrawTextW(dc, buf, -1, &tx,
                          DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
            }
            // 格线
            if (!(r == g->rowSel && c == g->colSel)) {
                HPEN pn = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_3DLIGHT));
                HPEN op = (HPEN)SelectObject(dc, pn);
                MoveToEx(dc, cell.right, cell.top, NULL);
                LineTo(dc, cell.right, cell.bottom);
                MoveToEx(dc, cell.left, cell.bottom, NULL);
                LineTo(dc, cell.right, cell.bottom);
                SelectObject(dc, op);
                DeleteObject(pn);
            }
        }
    }
    // 表体右边界与底部空区
    SetTextColor(dc, GetSysColor(COLOR_WINDOWTEXT));
    SelectObject(dc, old);
    (void)x;
}

static void gridUpdateScroll(struct vb6_Grid* g) {
    SCROLLINFO si;
    RECT rc;
    int rowH, headerH, visible;
    if (!g) return;
    GetClientRect(g->hwnd, &rc);
    rowH = gridPx(g->rowHeightTw);
    headerH = g->headers ? rowH : 0;
    visible = (rc.bottom - rc.top - headerH) / (rowH > 0 ? rowH : 1);
    if (visible < 1) visible = 1;
    memset(&si, 0, sizeof(si));
    si.cbSize = sizeof(si);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = g->rows > 0 ? g->rows - 1 : 0;
    si.nPage = visible;
    si.nPos = g->topRow;
    SetScrollInfo(g->hwnd, SB_VERT, &si, TRUE);
}

static LRESULT CALLBACK vb6_GridWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    struct vb6_Grid* g = gridOf(hwnd);
    switch (msg) {
    case WM_CREATE:
        gridSlot(hwnd);
        return 0;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        RECT rc;
        HDC dc = BeginPaint(hwnd, &ps);
        GetClientRect(hwnd, &rc);
        if (g) gridPaint(g, dc, rc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;   // 全在 WM_PAINT 里铺
    case WM_LBUTTONDOWN: {
        int rowH, headerH, row, col;
        if (!g) break;
        gridSyncCounts(g);
        rowH = gridPx(g->rowHeightTw);
        headerH = g->headers ? rowH : 0;
        row = g->topRow + (GET_Y_LPARAM(lp) - headerH) / (rowH > 0 ? rowH : 1);
        col = GET_X_LPARAM(lp) / gridPx(g->colWidthTw);
        if (GET_Y_LPARAM(lp) < headerH || row < 0 || row >= g->rows || col < 0 || col >= g->cols)
            break;
        g->rowSel = row;
        g->colSel = col;
        gridSeekSource(g);
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_VSCROLL: {
        SCROLLINFO si;
        if (!g) break;
        memset(&si, 0, sizeof(si));
        si.cbSize = sizeof(si);
        si.fMask = SIF_ALL;
        GetScrollInfo(hwnd, SB_VERT, &si);
        switch (LOWORD(wp)) {
        case SB_LINEUP:   si.nPos -= 1; break;
        case SB_LINEDOWN: si.nPos += 1; break;
        case SB_PAGEUP:   si.nPos -= (int)si.nPage; break;
        case SB_PAGEDOWN: si.nPos += (int)si.nPage; break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: si.nPos = si.nTrackPos; break;
        default: break;
        }
        if (si.nPos < si.nMin) si.nPos = si.nMin;
        if (si.nPos > si.nMax) si.nPos = si.nMax;
        g->topRow = si.nPos;
        gridUpdateScroll(g);
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_MOUSEWHEEL: {
        if (!g) break;
        g->topRow -= (GET_WHEEL_DELTA_WPARAM(wp) / WHEEL_DELTA) * 3;
        if (g->topRow < 0) g->topRow = 0;
        if (g->rows > 0 && g->topRow > g->rows - 1) g->topRow = g->rows - 1;
        gridUpdateScroll(g);
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_DESTROY:
        if (g) memset(g, 0, sizeof(*g));
        break;
    default: break;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void vb6_RegisterDataGridClass(void* hInstance) {
    static int done = 0;
    WNDCLASSW wc;
    if (done) return;
    done = 1;
    memset(&wc, 0, sizeof(wc));
    wc.style         = CS_DBLCLKS;
    wc.lpfnWndProc   = vb6_GridWndProc;
    wc.hInstance     = (HINSTANCE)hInstance;
    wc.hCursor       = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
    wc.lpszClassName = VB6_GRID_CLASS;
    RegisterClassW(&wc);
}

// 绑定：src = 引擎窗（Adodc 由 vb6_Adodc_Engine 交，Data 控件就是它自己）。
void vb6_DataGrid_Bind(void* hwnd, void* src) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    g->src = (HWND)src;
    g->rowSel = g->colSel = 0;
    g->topRow = 0;
    gridSyncCounts(g);
    gridUpdateScroll(g);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

void vb6_DataGrid_SetDataSourceName(void* hwnd, const wchar_t* name) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (g && name) { wcsncpy(g->dataSource, name, 127); g->dataSource[127] = 0; }
}

const void* vb6_DataGrid_GetDataSourceName(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd);
    return SysAllocString(g ? g->dataSource : L"");
}

// ============================ 属性 ============================
int32_t vb6_DataGrid_GetRowCount(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd);
    if (g) gridSyncCounts(g);
    return g ? g->rows : 0;
}
int32_t vb6_DataGrid_GetColCount(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd);
    if (g) gridSyncCounts(g);
    return g ? g->cols : 0;
}

int32_t vb6_DataGrid_GetRow(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd);
    return g ? g->rowSel : 0;
}
void vb6_DataGrid_SetRow(void* hwnd, int32_t v) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    gridSyncCounts(g);
    g->rowSel = v;
    gridClampSel(g);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}
int32_t vb6_DataGrid_GetCol(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd);
    return g ? g->colSel : 0;
}
void vb6_DataGrid_SetCol(void* hwnd, int32_t v) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    gridSyncCounts(g);
    g->colSel = v;
    gridClampSel(g);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

// Text：当前格的文本（VB6 的 DataGrid.Text 是只读快照，这里同口径）
const void* vb6_DataGrid_GetText(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd);
    wchar_t buf[256];
    buf[0] = 0;
    if (g && g->src) {
        gridSyncCounts(g);
        vb6_Data_FieldValueAt(g->src, g->rowSel + 1, g->colSel + 1, buf, 256);
    }
    return SysAllocString(buf);
}

int32_t vb6_DataGrid_GetColumnHeaders(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd);
    return g ? (g->headers ? -1 : 0) : 0;
}
void vb6_DataGrid_SetColumnHeaders(void* hwnd, int32_t v) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    g->headers = v ? -1 : 0;
    gridUpdateScroll(g);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

int32_t vb6_DataGrid_GetAllowAddNew(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd); return g ? (g->allowAddNew ? -1 : 0) : 0;
}
void vb6_DataGrid_SetAllowAddNew(void* hwnd, int32_t v) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd); if (g) g->allowAddNew = v ? -1 : 0;
}
int32_t vb6_DataGrid_GetAllowUpdate(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd); return g ? (g->allowUpdate ? -1 : 0) : 0;
}
void vb6_DataGrid_SetAllowUpdate(void* hwnd, int32_t v) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd); if (g) g->allowUpdate = v ? -1 : 0;
}
int32_t vb6_DataGrid_GetAllowDelete(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd); return g ? (g->allowDelete ? -1 : 0) : 0;
}
void vb6_DataGrid_SetAllowDelete(void* hwnd, int32_t v) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd); if (g) g->allowDelete = v ? -1 : 0;
}

int32_t vb6_DataGrid_GetRowHeight(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd); return g ? g->rowHeightTw : 300;
}
void vb6_DataGrid_SetRowHeight(void* hwnd, int32_t v) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    g->rowHeightTw = v > 0 ? v : 300;
    gridUpdateScroll(g);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}
int32_t vb6_DataGrid_GetDefColWidth(void* hwnd) {
    struct vb6_Grid* g = gridOf((HWND)hwnd); return g ? g->colWidthTw : 1200;
}
void vb6_DataGrid_SetDefColWidth(void* hwnd, int32_t v) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    g->colWidthTw = v > 0 ? v : 1200;
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

// ============================ 方法 ============================
void vb6_DataGrid_Refresh(void* hwnd) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    gridSyncCounts(g);
    gridClampSel(g);
    gridUpdateScroll(g);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

// ReBind：数据源结构变了（列数/行数）之后重读 —— 与 VB6 的 ReBind 同义。
void vb6_DataGrid_ReBind(void* hwnd) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    g->rowSel = g->colSel = 0;
    g->topRow = 0;
    vb6_DataGrid_Refresh(hwnd);
}

void vb6_DataGrid_ClearSelRows(void* hwnd) {
    struct vb6_Grid* g = gridSlot((HWND)hwnd);
    if (!g) return;
    g->rowSel = 0;
    g->colSel = 0;
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}
