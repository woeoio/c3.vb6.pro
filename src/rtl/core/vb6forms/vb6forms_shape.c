// vb6forms_shape.c - vb6forms 模块拆分: Shape/Line 窗口过程与属性 + GraphicalBtn
// 由 vb6forms.c 按控件/窗体功能家族拆分而来 (纯搬移, 零行为改动)

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include <stdio.h>
#include <stdarg.h>

#include <stdlib.h>   /* malloc, free */
#include <oleauto.h>  /* SysAllocString, BSTR */
#include <olectl.h>   /* IPicture, OleLoadPicture, OLE_HANDLE */


// ============================================================
// P20-35: Shape/Line自定义窗口类 + WM_PAINT绘制
// ============================================================

// --- Shape WndProc: 根据Shape属性绘制图形 ---
static LRESULT CALLBACK vb6_ShapeWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);

        int shapeType = (int)(INT_PTR)GetPropW(hwnd, L"VB6_ShapeType");
        int borderW  = (int)(INT_PTR)GetPropW(hwnd, L"VB6_ShapeBorderWidth");
        // C29-1a: 这两枚属性存的是 val+1 (见 vb6forms_widget_prop.c 的说明), 取出来要减回去
        HANDLE hBs = GetPropW(hwnd, L"VB6_ShapeBorderStyle");
        HANDLE hFs = GetPropW(hwnd, L"VB6_ShapeFillStyle");
        int borderS  = hBs ? (int)(INT_PTR)hBs - 1 : 1;
        int fillS    = hFs ? (int)(INT_PTR)hFs - 1 : 1;
        int32_t borderC = (int32_t)(INT_PTR)GetPropW(hwnd, L"VB6_ShapeBorderColor");
        int32_t fillC   = (int32_t)(INT_PTR)GetPropW(hwnd, L"VB6_ShapeFillColor");

        if (borderW <= 0) borderW = 1;
        if (borderS < 0) borderS = 1;  /* 旧写法把 0=Transparent 钳成 1, 下面 PS_NULL 分支成了死码 */

        /* 画笔: borderS映射 VB6: 0=Transparent,1=Solid,2=Dash,3=Dot,4=DashDot,5=DashDotDot,6=InsideSolid */
        int penStyle = PS_SOLID;
        if (borderS == 0) penStyle = PS_NULL;
        else if (borderS == 2) penStyle = PS_DASH;
        else if (borderS == 3) penStyle = PS_DOT;
        else if (borderS == 4) penStyle = PS_DASHDOT;
        else if (borderS == 5) penStyle = PS_DASHDOTDOT;
        else if (borderS == 6) penStyle = PS_INSIDEFRAME;

        COLORREF bcr = borderC ? (COLORREF)borderC : RGB(0,0,0);
        HPEN hPen = CreatePen(penStyle, borderW, bcr);
        HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

        /* 画刷: fillS映射 VB6: 0=Solid,1=Transparent,2=HorizontalLine,3=VerticalLine,4=UpwardDiagonal,5=DownwardDiagonal,6=Cross,7=DiagonalCross */
        HBRUSH hBr;
        COLORREF fcr = fillC ? (COLORREF)fillC : RGB(0,0,0);
        if (fillS == 1 || fillS == 0) {
            /* 0=Solid fill, 1=Transparent */
            hBr = (fillS == 1) ? (HBRUSH)GetStockObject(NULL_BRUSH) : CreateSolidBrush(fcr);
        } else {
            /* Hatch patterns: HS_HORIZONTAL=0, HS_VERTICAL=1, HS_FDIAGONAL=2, HS_BDIAGONAL=3, HS_CROSS=4, HS_DIAGCROSS=5 */
            int hatchMap[] = {0, 0, 0, 1, 2, 3, 4, 5};  /* fillS 2..7 -> HS_XXX */
            int hi = (fillS >= 2 && fillS <= 7) ? fillS : 2;
            hBr = CreateHatchBrush(hatchMap[hi], fcr);
        }
        HBRUSH hOldBr = (HBRUSH)SelectObject(hdc, hBr);

        /* 绘制形状: 0=Rectangle,1=Square,2=Oval,3=Circle,4=RoundedRectangle,5=RoundedSquare */
        int cx = rc.right / 2, cy = rc.bottom / 2;
        int r = (rc.right < rc.bottom) ? rc.right / 2 : rc.bottom / 2;
        int rw = rc.right / 6, rh = rc.bottom / 6;  /* 圆角大小 */

        switch (shapeType) {
            case 0: /* Rectangle */
                Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
                break;
            case 1: /* Square */
                { RECT sr = { cx - r, cy - r, cx + r, cy + r };
                  Rectangle(hdc, sr.left, sr.top, sr.right, sr.bottom); }
                break;
            case 2: /* Oval */
                Ellipse(hdc, rc.left, rc.top, rc.right, rc.bottom);
                break;
            case 3: /* Circle */
                Ellipse(hdc, cx - r, cy - r, cx + r, cy + r);
                break;
            case 4: /* Rounded Rectangle */
                RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, rw, rh);
                break;
            case 5: /* Rounded Square */
                RoundRect(hdc, cx - r, cy - r, cx + r, cy + r, rw/2, rh/2);
                break;
            default:
                Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
                break;
        }

        SelectObject(hdc, hOldPen);
        SelectObject(hdc, hOldBr);
        DeleteObject(hPen);
        if (fillS != 1) DeleteObject(hBr);  /* NULL_BRUSH is stock, don't delete */
        EndPaint(hwnd, &ps);
        return 0;
    } else if (msg == WM_DESTROY) {
        /* 清理属性 */
        RemovePropW(hwnd, L"VB6_ShapeType");
        RemovePropW(hwnd, L"VB6_ShapeBorderWidth");
        RemovePropW(hwnd, L"VB6_ShapeBorderStyle");
        RemovePropW(hwnd, L"VB6_ShapeFillStyle");
        RemovePropW(hwnd, L"VB6_ShapeBorderColor");
        RemovePropW(hwnd, L"VB6_ShapeFillColor");
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// --- Line WndProc: 绘制直线 ---
static LRESULT CALLBACK vb6_LineWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);

        /* Line属性: X1/Y1/X2/Y2 存的是**容器缇值** (与 VB6 的 .frm / 运行期赋值同一口径),
         * 控件窗口矩形 = 这四点的包围盒, 所以绘制时要先减掉包围盒原点再换算成像素。 */
        int x1t = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineX1");
        int y1t = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineY1");
        int x2t = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineX2");
        int y2t = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineY2");
        int bx1 = (x1t < x2t) ? x1t : x2t, by1 = (y1t < y2t) ? y1t : y2t;
        int x1 = vb6_TwipToX(x1t - bx1), y1 = vb6_TwipToY(y1t - by1);
        int x2 = vb6_TwipToX(x2t - bx1), y2 = vb6_TwipToY(y2t - by1);
        int borderW = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineBorderWidth");
        HANDLE hLbs = GetPropW(hwnd, L"VB6_LineBorderStyle");
        int borderS = hLbs ? (int)(INT_PTR)hLbs - 1 : 1;  /* 存的是 val+1, 见上 */
        int32_t borderC = (int32_t)(INT_PTR)GetPropW(hwnd, L"VB6_LineBorderColor");

        if (borderW <= 0) borderW = 1;
        /* 四点全为 0 (设计期没给坐标) 的退化形: 铺满控件矩形, 与旧行为一致 */
        if (x1t == 0 && y1t == 0 && x2t == 0 && y2t == 0) {
            x2 = rc.right; y2 = rc.bottom;
        }

        int penStyle = PS_SOLID;
        if (borderS == 0) penStyle = PS_NULL;
        else if (borderS == 2) penStyle = PS_DASH;
        else if (borderS == 3) penStyle = PS_DOT;
        else if (borderS == 4) penStyle = PS_DASHDOT;
        else if (borderS == 5) penStyle = PS_DASHDOTDOT;

        COLORREF cr = borderC ? (COLORREF)borderC : RGB(0,0,0);
        HPEN hPen = CreatePen(penStyle, borderW, cr);
        HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

        MoveToEx(hdc, x1, y1, NULL);
        LineTo(hdc, x2, y2);

        SelectObject(hdc, hOldPen);
        DeleteObject(hPen);
        EndPaint(hwnd, &ps);
        return 0;
    } else if (msg == WM_DESTROY) {
        RemovePropW(hwnd, L"VB6_LineX1");
        RemovePropW(hwnd, L"VB6_LineY1");
        RemovePropW(hwnd, L"VB6_LineX2");
        RemovePropW(hwnd, L"VB6_LineY2");
        RemovePropW(hwnd, L"VB6_LineBorderWidth");
        RemovePropW(hwnd, L"VB6_LineBorderStyle");
        RemovePropW(hwnd, L"VB6_LineBorderColor");
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// --- 注册VB6_Shape和VB6_Line窗口类 ---
void vb6_RegisterShapeLineClasses(void* hInstance) {
    static int registered = 0;
    if (registered) return;
    registered = 1;

    WNDCLASSW wc;
    memset(&wc, 0, sizeof(wc));
    wc.hInstance = (HINSTANCE)hInstance;
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)GetStockObject(NULL_BRUSH);  /* 透明背景 */

    /* VB6_Shape class */
    wc.lpszClassName = L"VB6_SHAPE";
    wc.lpfnWndProc = vb6_ShapeWndProc;
    RegisterClassW(&wc);

    /* VB6_Line class */
    wc.lpszClassName = L"VB6_LINE";
    wc.lpfnWndProc = vb6_LineWndProc;
    RegisterClassW(&wc);
}

// ============================================================
// P20-35: Line属性 (X1/Y1/X2/Y2/BorderColor/BorderStyle/BorderWidth)
// ============================================================

/* VB6 的 Line 没有 Left/Top/Width/Height —— 位置完全由 X1/Y1/X2/Y2 (容器缇值) 决定,
 * 控件窗口就是那四点的包围盒。赋任何一个端点都要重算包围盒并搬窗口, 否则
 * "Line1.X2 = 2000" 只挪了画线的端点、窗口还停在原处 (画到裁剪区外就看不见)。 */
static void vb6_LineApplyRect(HWND hwnd) {
    int x1 = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineX1");
    int y1 = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineY1");
    int x2 = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineX2");
    int y2 = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineY2");
    int bw = (int)(INT_PTR)GetPropW(hwnd, L"VB6_LineBorderWidth");
    if (bw < 1) bw = 1;
    int left = (x1 < x2) ? x1 : x2, top = (y1 < y2) ? y1 : y2;
    int wTw = (x1 < x2) ? x2 - x1 : x1 - x2, hTw = (y1 < y2) ? y2 - y1 : y1 - y2;
    int px = vb6_TwipToX(left), py = vb6_TwipToY(top);
    /* 细的那一维至少给到笔宽, 否则水平/垂直线被自己裁掉 */
    int pw = vb6_TwipToX(wTw), ph = vb6_TwipToY(hTw);
    if (wTw == 0) pw = bw; else if (pw < 1) pw = 1;
    if (hTw == 0) ph = bw; else if (ph < 1) ph = 1;
    SetWindowPos(hwnd, NULL, px, py, pw, ph, SWP_NOZORDER | SWP_NOACTIVATE);
    InvalidateRect(hwnd, NULL, TRUE);
}

int32_t vb6_GetLineX1(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_LineX1");
    return h ? (int32_t)(INT_PTR)h : 0;
}
void vb6_SetLineX1(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_LineX1", (HANDLE)(INT_PTR)val);
    vb6_LineApplyRect((HWND)hwnd);
}
int32_t vb6_GetLineY1(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_LineY1");
    return h ? (int32_t)(INT_PTR)h : 0;
}
void vb6_SetLineY1(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_LineY1", (HANDLE)(INT_PTR)val);
    vb6_LineApplyRect((HWND)hwnd);
}
int32_t vb6_GetLineX2(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_LineX2");
    return h ? (int32_t)(INT_PTR)h : 0;
}
void vb6_SetLineX2(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_LineX2", (HANDLE)(INT_PTR)val);
    vb6_LineApplyRect((HWND)hwnd);
}
int32_t vb6_GetLineY2(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_LineY2");
    return h ? (int32_t)(INT_PTR)h : 0;
}
void vb6_SetLineY2(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_LineY2", (HANDLE)(INT_PTR)val);
    vb6_LineApplyRect((HWND)hwnd);
}
int32_t vb6_GetLineBorderWidth(void* hwnd) {
    if (!hwnd) return 1;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_LineBorderWidth");
    return h ? (int32_t)(INT_PTR)h : 1;
}
void vb6_SetLineBorderWidth(void* hwnd, int32_t val) {
    if (!hwnd) return;
    if (val < 1) val = 1;
    if (val > 8192) val = 8192;
    SetPropW((HWND)hwnd, L"VB6_LineBorderWidth", (HANDLE)(INT_PTR)val);
    /* C29-1a: 同 Shape —— 笔宽不是 1 时虚线族 (2..5) 强制回实线 (手册那条规则) */
    if (val != 1) {
        int32_t bsC29 = vb6_GetLineBorderStyle(hwnd);
        if (bsC29 != 0 && bsC29 != 6) vb6_SetLineBorderStyle(hwnd, 1);
    }
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}
int32_t vb6_GetLineBorderStyle(void* hwnd) {
    if (!hwnd) return 1;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_LineBorderStyle");
    return h ? (int32_t)(INT_PTR)h - 1 : 1;  /* C29-1a: 存的是 val+1 (0 与"没存过"不可分辨) */
}
void vb6_SetLineBorderStyle(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_LineBorderStyle", (HANDLE)(INT_PTR)(val + 1));
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}
int32_t vb6_GetLineColor(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_LineBorderColor");
    return h ? (int32_t)(INT_PTR)h : (int32_t)RGB(0,0,0);
}
void vb6_SetLineColor(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_LineBorderColor", (HANDLE)(INT_PTR)val);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

// ============================================================
// P20-34补充: Shape FillColor/BorderColor属性
// ============================================================

int32_t vb6_GetShapeBorderColor(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_ShapeBorderColor");
    return h ? (int32_t)(INT_PTR)h : (int32_t)RGB(0,0,0);
}
void vb6_SetShapeBorderColor(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ShapeBorderColor", (HANDLE)(INT_PTR)val);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}
int32_t vb6_GetShapeFillColor(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_ShapeFillColor");
    return h ? (int32_t)(INT_PTR)h : (int32_t)RGB(0,0,0);
}
void vb6_SetShapeFillColor(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ShapeFillColor", (HANDLE)(INT_PTR)val);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

/* M12-FIX3: Graphical button (Style=1) subclass - draw picture above caption text */
static LRESULT CALLBACK vb6_GraphicalBtnSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_PAINT) {
        HANDLE hBmp = GetPropW(hwnd, L"VB6_GfxBtn_Bmp");
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);

        /* Determine button state: check state + push state + disabled */
        LRESULT checkState = SendMessageW(hwnd, BM_GETCHECK, 0, 0);
        LRESULT btnState = SendMessageW(hwnd, BM_GETSTATE, 0, 0);
        BOOL isChecked = (checkState != BST_UNCHECKED);
        BOOL isPushed = isChecked || (btnState & BST_PUSHED);
        BOOL isFocused = (btnState & BST_FOCUS) ? TRUE : FALSE;
        BOOL isEnabled = IsWindowEnabled(hwnd);

        /* Draw button background (3D raised/sunken) */
        UINT dfcState = DFCS_BUTTONPUSH;
        if (isPushed) dfcState |= DFCS_PUSHED;
        if (!isEnabled) dfcState |= DFCS_INACTIVE;
        DrawFrameControl(hdc, &rc, DFC_BUTTON, dfcState);

        /* Offset for 3D push effect */
        int pushOff = isPushed ? 1 : 0;

        /* Draw picture if available */
        if (hBmp && GetObjectType((HGDIOBJ)hBmp) == OBJ_BITMAP) {
            BITMAP bm;
            GetObjectW((HBITMAP)hBmp, sizeof(bm), &bm);

            /* Get caption text to calculate layout */
            WCHAR caption[256] = {0};
            int captionLen = GetWindowTextW(hwnd, caption, 256);

            /* Layout: picture on top, text on bottom, both centered */
            int textH = 0;
            if (captionLen > 0) {
                SIZE sz;
                HFONT hFont = vb6_ControlFont(hwnd);   // 账 #204: 与 Print/文字量同一处问法
                HFONT oldFont = NULL;
                if (hFont) oldFont = (HFONT)SelectObject(hdc, hFont);
                GetTextExtentPoint32W(hdc, caption, captionLen, &sz);
                if (oldFont) SelectObject(hdc, oldFont);
                textH = sz.cy + 4; /* 2px padding top+bottom */
            }

            int btnW = rc.right - rc.left;
            int btnH = rc.bottom - rc.top;
            int border = 4; /* inset from button edge */
            int availH = btnH - 2 * border - textH;

            /* Draw bitmap centered horizontally, top-aligned in available area */
            int imgX = pushOff + border + (btnW - 2 * border - bm.bmWidth) / 2;
            int imgY = pushOff + border + (availH - bm.bmHeight) / 2;
            if (imgY < pushOff + border) imgY = pushOff + border;

            HDC memDC = CreateCompatibleDC(hdc);
            HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, (HBITMAP)hBmp);
            if (isEnabled) {
                /* Fix <vbeclipse> 2026-10-06: 32bpp 带 alpha 的先走 AlphaBlend
                   (PNG 按钮的圆角/透明区露按钮底色, 不再是黑块); 不透明回落 BitBlt。 */
                if (!vb6_DrawBitmapAlpha(hdc, (void*)hBmp, imgX, imgY, bm.bmWidth, bm.bmHeight)) {
                    BitBlt(hdc, imgX, imgY, bm.bmWidth, bm.bmHeight, memDC, 0, 0, SRCCOPY);
                }
            } else {
                /* Disabled: draw grayed image using PATCOPY with halftone brush */
                BitBlt(hdc, imgX, imgY, bm.bmWidth, bm.bmHeight, memDC, 0, 0, SRCCOPY);
                /* Overlay with semi-transparent gray to indicate disabled state */
                HBRUSH hGray = CreateSolidBrush(GetSysColor(COLOR_GRAYTEXT));
                HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, hGray);
                PatBlt(hdc, imgX, imgY, bm.bmWidth, bm.bmHeight, PATINVERT);
                SelectObject(hdc, oldBrush);
                DeleteObject(hGray);
            }
            SelectObject(memDC, oldBmp);
            DeleteDC(memDC);

            /* Draw caption text below image, centered */
            if (captionLen > 0) {
                int textY = pushOff + imgY + bm.bmHeight + 2 - pushOff;
                if (textY + textH > btnH - border) textY = btnH - border - textH;
                RECT textRc;
                textRc.left = pushOff + border;
                textRc.top = textY;
                textRc.right = btnW - border + pushOff;
                textRc.bottom = textY + textH;
                HFONT hFont = vb6_ControlFont(hwnd);   // 账 #204: 与 Print/文字量同一处问法
                HFONT oldFont = NULL;
                if (hFont) oldFont = (HFONT)SelectObject(hdc, hFont);
                SetBkMode(hdc, TRANSPARENT);
                if (!isEnabled) SetTextColor(hdc, GetSysColor(COLOR_GRAYTEXT));
                DrawTextW(hdc, caption, captionLen, &textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                if (oldFont) SelectObject(hdc, oldFont);
            }
        } else {
            /* No bitmap: just draw caption text */
            WCHAR caption[256] = {0};
            int captionLen = GetWindowTextW(hwnd, caption, 256);
            if (captionLen > 0) {
                HFONT hFont = vb6_ControlFont(hwnd);   // 账 #204: 与 Print/文字量同一处问法
                HFONT oldFont = NULL;
                if (hFont) oldFont = (HFONT)SelectObject(hdc, hFont);
                SetBkMode(hdc, TRANSPARENT);
                if (!isEnabled) SetTextColor(hdc, GetSysColor(COLOR_GRAYTEXT));
                RECT textRc = rc;
                InflateRect(&textRc, -4, -4);
                textRc.left += pushOff;
                textRc.top += pushOff;
                DrawTextW(hdc, caption, captionLen, &textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                if (oldFont) SelectObject(hdc, oldFont);
            }
        }

        /* Draw focus rectangle if button has focus */
        if (isFocused) {
            RECT focusRc = rc;
            InflateRect(&focusRc, -3, -3);
            DrawFocusRect(hdc, &focusRc);
        }

        EndPaint(hwnd, &ps);
        return 0;
    }
    /* State-change messages: let default proc handle, then force our repaint.
       Default button wndproc directly draws to DC (bypassing WM_PAINT),
       which overwrites our custom picture+text. Force InvalidateRect after. */
    if (msg == WM_SETFOCUS || msg == WM_KILLFOCUS ||
        msg == WM_ENABLE || msg == WM_CANCELMODE ||
        (msg == BM_SETCHECK) || (msg == BM_SETSTATE)) {
        WNDPROC origProc = (WNDPROC)GetPropW(hwnd, L"VB6_GfxBtn_OrigProc");
        SendMessageW(hwnd, WM_SETREDRAW, FALSE, 0);  /* prevent flicker: disable auto-redraw */
        LRESULT result = CallWindowProcW(origProc, hwnd, msg, wp, lp);
        SendMessageW(hwnd, WM_SETREDRAW, TRUE, 0);   /* re-enable redraw */
        InvalidateRect(hwnd, NULL, FALSE);            /* our WM_PAINT draws picture+text */
        return result;
    }
    if (msg == WM_ERASEBKGND) {
        /* We handle all drawing in WM_PAINT, no need to erase background */
        return 1;
    }
    if (msg == WM_DESTROY) {
        HANDLE hBmp = GetPropW(hwnd, L"VB6_GfxBtn_Bmp");
        if (hBmp) { DeleteObject(hBmp); RemovePropW(hwnd, L"VB6_GfxBtn_Bmp"); }
        WNDPROC orig = (WNDPROC)GetPropW(hwnd, L"VB6_GfxBtn_OrigProc");
        if (orig) { RemovePropW(hwnd, L"VB6_GfxBtn_OrigProc"); SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)orig); }
        return 0;
    }
    WNDPROC origProc = (WNDPROC)GetPropW(hwnd, L"VB6_GfxBtn_OrigProc");
    if (origProc) return CallWindowProcW(origProc, hwnd, msg, wp, lp);
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void vb6_GraphicalBtn_SetImage(void* hwnd, void* hBitmap) {
    if (!hwnd || !hBitmap) return;
    HWND hw = (HWND)hwnd;
    /* Store bitmap as property */
    SetPropW(hw, L"VB6_GfxBtn_Bmp", (HANDLE)hBitmap);
    /* Subclass if not already */
    if (!GetPropW(hw, L"VB6_GfxBtn_OrigProc")) {
        WNDPROC origProc = (WNDPROC)SetWindowLongPtrW(hw, GWLP_WNDPROC, (LONG_PTR)vb6_GraphicalBtnSubclassProc);
        SetPropW(hw, L"VB6_GfxBtn_OrigProc", (HANDLE)origProc);
    }
    InvalidateRect(hw, NULL, TRUE);
}

/* ============================================================
 * 动态加载控件 (Controls.Add)
 * ============================================================ */
