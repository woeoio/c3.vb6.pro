// vb6forms_picture_prop.c - vb6forms 模块拆分: Picture 属性读写 + Image Stretch + WM_PAINT 子类化
// 由 vb6forms_picture.c 拆出 (2026-09-17, 纯搬移, 零行为改动)
// 本文件内的 static vb6_ImageSubclassProc 只被 vb6_InstallImageSubclass 使用, 同文件内可见.

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


void* vb6_GetControlPicture(void* hwnd) {
    if (!hwnd) return NULL;
    HWND hw = (HWND)hwnd;
    /* If a COM IPicture was stored via SetControlPictureFromCom, return the
     * borrowed COM pointer so callers (e.g. Clipboard.SetData) can QI/use it.
     * The caller must NOT Release this borrowed reference. */
    IPicture* pPic = (IPicture*)GetPropW(hw, L"VB6_IPicture");
    if (pPic) return (void*)pPic;
    HANDLE hProp = GetPropW(hw, L"VB6_Picture");
    return hProp;  // HBITMAP/HICON handle
}

void vb6_SetControlPicture(void* hwnd, void* hPicture) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    /* A raw handle is replacing any COM IPicture set earlier via
     * SetControlPictureFromCom. Release the retained COM reference so it
     * doesn't leak and won't be used by the subclass. */
    IPicture* pOldCom = (IPicture*)GetPropW(hw, L"VB6_IPicture");
    if (pOldCom) {
        pOldCom->lpVtbl->Release(pOldCom);
        RemovePropW(hw, L"VB6_IPicture");
    }
    /* Store the picture handle and type (default=bitmap) */
    SetPropW(hw, L"VB6_Picture", (HANDLE)hPicture);
    SetPropW(hw, L"VB6_PictureType", (HANDLE)1);  /* default: bitmap */
    
    /* Apply to the control */
    WCHAR className[256] = {0};
    GetClassNameW(hw, className, 256);
    
    if (wcsicmp(className, L"STATIC") == 0) {
        /* PictureBox uses STATIC control with SS_BITMAP/SS_ICON/SS_ENHMETAFILE style */
        if (hPicture) {
            /* Detect picture type - check if it's an icon by trying */
            LONG style = GetWindowLongW(hw, GWL_STYLE);
            style &= ~(SS_BITMAP | SS_ICON | SS_ENHMETAFILE);
            
            /* For now assume bitmap; icon detection would require more logic */
            style |= SS_BITMAP | SS_CENTERIMAGE;
            SetWindowLongW(hw, GWL_STYLE, style);
            SendMessageW(hw, STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)hPicture);
        } else {
            /* Clear picture */
            LONG style = GetWindowLongW(hw, GWL_STYLE);
            style &= ~(SS_BITMAP | SS_ICON | SS_ENHMETAFILE);
            SetWindowLongW(hw, GWL_STYLE, style);
            SendMessageW(hw, STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)NULL);
        }
        InvalidateRect(hw, NULL, TRUE);
    }
    
    /* AutoSize: if enabled, resize to fit picture */
    int autoSize = vb6_GetPictureAutoSize(hwnd);
    if (autoSize && hPicture) {
        BITMAP bm;
        HBITMAP hBmp = (HBITMAP)hPicture;
        if (GetObjectW(hBmp, sizeof(bm), &bm) != 0) {
            SetWindowPos(hw, NULL, 0, 0, bm.bmWidth, bm.bmHeight,
                SWP_NOMOVE | SWP_NOZORDER);
        }
    }
}


// Set Picture from COM IPictureDisp object (e.g. ImageList1.ListImages(i).Picture)
// IPictureDisp::get_Handle returns OLE_HANDLE (HBITMAP for bitmap type)
//
// All COM picture types are rendered directly via the retained IPicture COM
// pointer in the Image subclass (vb6_ImageSubclassProc). We do NOT convert
// metafiles to bitmaps here — that was the broken "control -> bitmap" path.
// VB6_Picture still keeps the native handle for legacy raw-handle callers.
void vb6_SetControlPictureFromCom(void* hwnd, void* pPictureDisp) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;

    /* NULL COM setter clears the picture (release COM ref + clear props). */
    if (!pPictureDisp) {
        IPicture* pOld = (IPicture*)GetPropW(hw, L"VB6_IPicture");
        if (pOld) {
            pOld->lpVtbl->Release(pOld);
            RemovePropW(hw, L"VB6_IPicture");
        }
        SetPropW(hw, L"VB6_Picture", (HANDLE)NULL);
        SetPropW(hw, L"VB6_PictureType", (HANDLE)0);
        InvalidateRect(hw, NULL, TRUE);
        return;
    }

    HRESULT hr;
    /* IPictureDisp and IPicture are separate interfaces.
     * IPictureDisp inherits IDispatch, IPicture inherits IUnknown.
     * Must QI for IID_IPicture from IPictureDisp pointer. */
    IPicture* pPic = NULL;
    IUnknown* pUnk = (IUnknown*)pPictureDisp;
    hr = pUnk->lpVtbl->QueryInterface(pUnk, &IID_IPicture, (void**)&pPic);
    if (FAILED(hr) || !pPic) return;

    OLE_HANDLE hOle = 0;
    SHORT nType = 0;
    // Get picture type: 1=Bitmap, 2=Metafile, 3=Icon, 4=Enhanced Metafile
    hr = pPic->lpVtbl->get_Type(pPic, &nType);
    if (FAILED(hr)) { pPic->lpVtbl->Release(pPic); return; }
    hr = pPic->lpVtbl->get_Handle(pPic, &hOle);
    if (FAILED(hr)) { pPic->lpVtbl->Release(pPic); return; }

    /* QI already acquired the new reference, including self-assignment.
     * Transfer that reference to the control and release exactly one old ref. */
    IPicture* pOld = (IPicture*)GetPropW(hw, L"VB6_IPicture");
    if (!SetPropW(hw, L"VB6_IPicture", (HANDLE)pPic)) {
        pPic->lpVtbl->Release(pPic);
        return;
    }
    if (pOld) pOld->lpVtbl->Release(pOld);

    /* Keep the native handle for legacy raw-handle callers (GetControlPicture
     * falls back to this when no COM pointer is stored). */
    SetPropW(hw, L"VB6_Picture", (HANDLE)(UINT_PTR)hOle);
    SetPropW(hw, L"VB6_PictureType", (HANDLE)(INT_PTR)nType);

    /* The retained QI reference keeps the native handle alive for painting. */

    /* All COM picture types route through the Image subclass for rendering.
     * Remove the default STATIC bitmap/icon/emf rendering styles so WM_PAINT is
     * fully owned by our subclass. */
    LONG style = GetWindowLongW(hw, GWL_STYLE);
    style &= ~(SS_BITMAP | SS_ICON | SS_ENHMETAFILE | SS_CENTERIMAGE);
    SetWindowLongW(hw, GWL_STYLE, style);

    vb6_InstallImageSubclass(hwnd);
    InvalidateRect(hw, NULL, TRUE);
}
int vb6_GetPictureAutoSize(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_AutoSize");
    if (hProp) return (int)(INT_PTR)hProp;
    return 0;
}

void vb6_SetPictureAutoSize(void* hwnd, int autoSize) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_AutoSize", (HANDLE)(INT_PTR)autoSize);
}
// ============================================================
// P17.2: Image.Stretch property + WM_PAINT subclass (StretchBlt)
// ============================================================

int vb6_GetImageStretch(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_Stretch");
    return hProp ? (int)(INT_PTR)hProp : 0;
}

void vb6_SetImageStretch(void* hwnd, int stretch) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    SetPropW(hw, L"VB6_Stretch", (HANDLE)(INT_PTR)stretch);
    if (stretch) {
        vb6_InstallImageSubclass(hwnd);
    }
    InvalidateRect(hw, NULL, TRUE);
}

// ============================================================
// Fix <vbeclipse> 2026-10-06: Picture 的现代 RGBA 透明绘制。
// 旧路径 (IPicture::Render / BitBlt) 不认 32bpp 的 alpha 通道 —— PNG/WebP 的
// 透明区被白底 (子类先 FillRect 白) 或底色顶掉。这里把位图取成 32bpp 自上而下
// DIB 做判定:
//   · alpha 全 0   → 整张透明, 什么都不画;
//   · alpha 全 255 → 完全不透明, 回调用方原路径 (BitBlt/Render 更快更准);
//   · 其余        → AlphaBlend(AC_SRC_ALPHA)。GDI+/OleLoadPicture 的产物是预乘
//     (PARGB); 若扫到 RGB>A 的像素说明数据是直通 alpha, 先行预乘再画。
// msimg32 按需 LoadLibrary (与 winmm/timeSetEvent 同纪律, 不新增 import lib)。
// 返回 1 = 已按 alpha 画完; 0 = 走不了 alpha 路径 (调用方回落)。
// ============================================================
typedef BOOL (WINAPI *vb6_pfnAlphaBlendT)(HDC, LONG, LONG, LONG, LONG, HDC, LONG, LONG, LONG, LONG, BLENDFUNCTION);
static vb6_pfnAlphaBlendT vb6_pAlphaBlend = NULL;
static int vb6_abProbed = 0;

static vb6_pfnAlphaBlendT vb6_AlphaBlendProc(void) {
    if (!vb6_abProbed) {
        vb6_abProbed = 1;
        HMODULE h = LoadLibraryW(L"msimg32.dll");
        if (h) vb6_pAlphaBlend = (vb6_pfnAlphaBlendT)(void*)GetProcAddress(h, "AlphaBlend");
    }
    return vb6_pAlphaBlend;
}

int vb6_DrawBitmapAlpha(void* hdcV, void* hBmpV, int dstX, int dstY, int dstW, int dstH) {
    HDC hdc = (HDC)hdcV;
    HBITMAP hBmp = (HBITMAP)hBmpV;
    BITMAP bm;
    if (!hdc || !hBmp || GetObjectW(hBmp, sizeof(bm), &bm) == 0) return 0;
    if (getenv("C3_ALPHA_TRACE"))
        fprintf(stderr, "[ALPHA] entry bpp=%d %dx%d\n", bm.bmBitsPixel, bm.bmWidth, bm.bmHeight);
    if (bm.bmBitsPixel != 32) return 0;
    int w = bm.bmWidth, h = bm.bmHeight > 0 ? bm.bmHeight : -bm.bmHeight;
    if (w <= 0 || h <= 0 || dstW <= 0 || dstH <= 0) return 0;

    BITMAPINFO bmi; memset(&bmi, 0, sizeof(bmi));
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = w;
    bmi.bmiHeader.biHeight = -h;   /* top-down */
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    size_t bytes = (size_t)w * h * 4;
    unsigned char* buf = (unsigned char*)malloc(bytes);
    if (!buf) return 0;
    HDC scr = GetDC(NULL);
    int got = GetDIBits(scr, hBmp, 0, h, buf, &bmi, DIB_RGB_COLORS);
    ReleaseDC(NULL, scr);
    if (!got) { free(buf); return 0; }

    /* alpha 统计: 全 0 / 全 255 / 部分透明; RGB>A ⇒ 直通 alpha, 需预乘 */
    int allZero = 1, allFF = 1, needPreMul = 0;
    size_t n = (size_t)w * h;
    for (size_t i = 0; i < n; i++) {
        const unsigned char* px = buf + i * 4;   /* BGRA */
        unsigned char a = px[3];
        if (a != 0) allZero = 0;
        if (a != 0xFF) allFF = 0;
        if (a < 255 && (px[0] > a || px[1] > a || px[2] > a)) needPreMul = 1;
        if (!allZero && !allFF && needPreMul) break;
    }
    if (getenv("C3_ALPHA_TRACE"))
        fprintf(stderr, "[ALPHA] %dx%d allZero=%d allFF=%d needPreMul=%d\n", w, h, allZero, allFF, needPreMul);
    if (getenv("C3_ALPHA_TRACE")) {
        const unsigned char* pl = buf + (size_t)(10 * w + 10) * 4;
        const unsigned char* pr = buf + (size_t)(10 * w + (w - 5)) * 4;
        fprintf(stderr, "[ALPHA] pxL BGR A=%d,%d,%d A=%u | pxR A=%d,%d,%d A=%u\n",
                pl[2], pl[1], pl[0], pl[3], pr[2], pr[1], pr[0], pr[3]);
    }
    if (allZero) { free(buf); return 1; }   /* 整张全透明: 什么都不画 */
    if (allFF)   { free(buf); return 0; }   /* 完全不透明: 原路径足够 */

    if (needPreMul) {
        for (size_t i = 0; i < n; i++) {
            unsigned char* px = buf + i * 4;
            unsigned a = px[3];
            px[0] = (unsigned char)(px[0] * a / 255);
            px[1] = (unsigned char)(px[1] * a / 255);
            px[2] = (unsigned char)(px[2] * a / 255);
        }
    }

    void* bits = NULL;
    HBITMAP hDib = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, &bits, NULL, 0);
    if (!hDib || !bits) { if (hDib) DeleteObject(hDib); free(buf); return 0; }
    memcpy(bits, buf, bytes);
    free(buf);

    HDC memDC = CreateCompatibleDC(hdc);
    if (!memDC) { DeleteObject(hDib); return 0; }
    HGDIOBJ old = SelectObject(memDC, hDib);
    int drawn = 0;
    vb6_pfnAlphaBlendT pAB = vb6_AlphaBlendProc();
    if (pAB) {
        BLENDFUNCTION bf;
        bf.BlendOp = AC_SRC_OVER;
        bf.BlendFlags = 0;
        bf.SourceConstantAlpha = 255;
        bf.AlphaFormat = AC_SRC_ALPHA;
        drawn = pAB(hdc, dstX, dstY, dstW, dstH, memDC, 0, 0, w, h, bf);
    }
    SelectObject(memDC, old);
    DeleteDC(memDC);
    DeleteObject(hDib);
    return drawn ? 1 : 0;
}

/* Shared painting for WM_PAINT and WM_PRINTCLIENT.
 * Draws the retained COM IPicture (metafile/bitmap/icon) or a raw
 * HBITMAP/HICON/HENHMETAFILE handle, honoring the Stretch property.
 * hdc is the paint DC (from BeginPaint) or the DC supplied by WM_PRINTCLIENT. */
static void vb6_ImagePaintHelper(HWND hwnd, HDC hdc) {
    RECT rc;
    GetClientRect(hwnd, &rc);

    /* Fix 187: 背景 = VB6_BackColor 属性色 (VB6 PictureBox 语义)。
     * 原先硬编码 COLOR_BTNFACE — BackColor 写入 (vb6_SetControlBackColor)
     * 后绘制路径读不到, 色块永远不变色。无属性时回落 BTNFACE (Image 默认)。 */
    int stretch = (int)(INT_PTR)GetPropW(hwnd, L"VB6_Stretch");
    {
        COLORREF bg187 = (COLORREF)vb6_GetControlBackColor(hwnd);
        if (bg187 & 0x80000000L) bg187 = GetSysColor(bg187 & 0xFF);
        HBRUSH hBg = CreateSolidBrush(bg187);
        FillRect(hdc, &rc, hBg);
        DeleteObject(hBg);
    }

    /* Preferred path: retained COM IPicture (from SetControlPictureFromCom).
     * Works for bitmap, icon, and metafile uniformly via IPicture::Render.
     * Fix <vbeclipse> 2026-10-06: 位图型先试 RGBA 透明绘制 (PNG/WebP 的透明区
     * 露 BackColor 而不是白底); 不透明/非位图保持原白底+Render 行为。 */
    IPicture* pPic = (IPicture*)GetPropW(hwnd, L"VB6_IPicture");
    if (pPic) {
        OLE_XSIZE_HIMETRIC hmW = 0;
        OLE_YSIZE_HIMETRIC hmH = 0;
        pPic->lpVtbl->get_Width(pPic, &hmW);
        pPic->lpVtbl->get_Height(pPic, &hmH);

        int dstW, dstH;
        if (stretch) {
            /* Stretch fills the whole client area. */
            dstW = rc.right;
            dstH = rc.bottom;
        } else {
            /* Non-stretch: draw at the picture's natural pixel size, converting
             * HIMETRIC (2540 HIMETRIC per logical inch) using the DC's DPI. */
            int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
            int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
            dstW = MulDiv((int)hmW, dpiX, 2540);
            dstH = MulDiv((int)hmH, dpiY, 2540);
            if (dstW <= 0) dstW = rc.right;
            if (dstH <= 0) dstH = rc.bottom;
        }

        /* RGBA 透明位图: 直接 AlphaBlend 到 BackColor 上 */
        OLE_HANDLE hOle206 = 0;
        SHORT nType206 = 0;
        pPic->lpVtbl->get_Type(pPic, &nType206);
        pPic->lpVtbl->get_Handle(pPic, &hOle206);
        if (nType206 == 1 /*PICTYPE_BITMAP*/ && hOle206
            && GetObjectType((HGDIOBJ)(UINT_PTR)hOle206) == OBJ_BITMAP
            && vb6_DrawBitmapAlpha(hdc, (void*)(UINT_PTR)hOle206, 0, 0, dstW, dstH)) {
            return;
        }

        /* Metafiles / QR codes typically expect a white background. */
        FillRect(hdc, &rc, (HBRUSH)GetStockObject(WHITE_BRUSH));

        /* HIMETRIC y points up, so flip the source rect: (0, hmH, hmW, -hmH). */
        pPic->lpVtbl->Render(pPic, hdc, 0, 0, dstW, dstH,
                             0, hmH, hmW, -hmH, NULL);
        return;
    }

    HANDLE hPict = GetPropW(hwnd, L"VB6_Picture");
    int picType = (int)(INT_PTR)GetPropW(hwnd, L"VB6_PictureType");
    if (!hPict) return;

    DWORD objType = GetObjectType((HGDIOBJ)hPict);
    if (objType == OBJ_ENHMETAFILE) {
        PlayEnhMetaFile(hdc, (HENHMETAFILE)hPict, &rc);
        return;
    }
    if (objType == OBJ_BITMAP) {
        HBITMAP hBmp = (HBITMAP)hPict;
        BITMAP bm;
        GetObjectW(hBmp, sizeof(bm), &bm);
        /* Fix <vbeclipse> 2026-10-06: 32bpp 带 alpha 的先走 AlphaBlend (透明区露底色) */
        if (vb6_DrawBitmapAlpha(hdc, hBmp, 0, 0,
                                stretch ? rc.right : bm.bmWidth,
                                stretch ? rc.bottom : bm.bmHeight)) {
            return;
        }
        HDC memDC = CreateCompatibleDC(hdc);
        HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, hBmp);
        if (stretch) {
            SetStretchBltMode(hdc, HALFTONE);
            SetBrushOrgEx(hdc, 0, 0, NULL);
            StretchBlt(hdc, 0, 0, rc.right, rc.bottom,
                       memDC, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY);
        } else {
            BitBlt(hdc, 0, 0, bm.bmWidth, bm.bmHeight, memDC, 0, 0, SRCCOPY);
        }
        SelectObject(memDC, oldBmp);
        DeleteDC(memDC);
        return;
    }
    if (picType == 3) {
        DrawIconEx(hdc, 0, 0, (HICON)hPict, 0, 0, 0, NULL, DI_NORMAL);
        return;
    }
}

/* Image control subclass WndProc for WM_PAINT / WM_PRINTCLIENT rendering. */
static LRESULT CALLBACK vb6_ImageSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_PAINT) {
        // 账 #185: BeginPaint/EndPaint 归**最外层**那一趟 (生成代码里的 _Paint 臂)。
        // 上层已经把 DC 放进 VB6_PaintDC 时本层只往那张 DC 上画表面 —— 否则一次绘制
        // 出现两个 PAINTSTRUCT, 而且表面会画在用户笔画**之后**把它盖掉。
        HDC given = (HDC)(INT_PTR)GetPropW(hwnd, L"VB6_PaintDC");
        PAINTSTRUCT ps;
        HDC hdc = given;
        if (!hdc) hdc = BeginPaint(hwnd, &ps);
        if (hdc) {
            if (GetEnvironmentVariableA("C3_FORMS_TRACE", NULL, 0) > 0)
                fprintf(stderr, "[C3_F187] ImagePaint hwnd=%p bg=%06X set=%d shared=%d\r\n",
                        (void*)hwnd,
                        (unsigned)(UINT_PTR)GetPropW(hwnd, L"VB6_BackColor") & 0xFFFFFFu,
                        GetPropW(hwnd, L"VB6_BackColorSet") ? 1 : 0, given ? 1 : 0);
            vb6_ImagePaintHelper(hwnd, hdc);
            if (!given) EndPaint(hwnd, &ps);
        }
        return 0;
    } else if (msg == WM_PRINTCLIENT) {
        /* Printing / theming asks us to render into the provided DC. */
        HDC hdc = (HDC)wp;
        if (hdc) vb6_ImagePaintHelper(hwnd, hdc);
        return 0;
    } else if (msg == WM_DESTROY) {
        /* Capture the original WndProc BEFORE removing props, then forward the
         * message to it. The old code removed the prop and then looked it up
         * again at the fall-through, so WM_DESTROY was lost to DefWindowProc. */
        WNDPROC origProc = (WNDPROC)GetPropW(hwnd, L"VB6_ImageOrigProc");

        /* 账 #196: `.hDC` 交出去的那张窗口 DC 归本层缓存 (VB6 是一个对象一张)，
           窗口销毁时在这里归还 —— 别在 vb6_GetControlHDC 里 ReleaseDC，那个句柄
           已经给了 VB 代码。 */
        {
            HDC hObjDC = (HDC)GetPropW(hwnd, L"VB6_ObjectDC");
            if (hObjDC) { ReleaseDC(hwnd, hObjDC); RemovePropW(hwnd, L"VB6_ObjectDC"); }
        }

        IPicture* pPic = (IPicture*)GetPropW(hwnd, L"VB6_IPicture");
        if (pPic) {
            pPic->lpVtbl->Release(pPic);
            RemovePropW(hwnd, L"VB6_IPicture");
        }
        if (origProc) {
            SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)origProc);
            RemovePropW(hwnd, L"VB6_ImageOrigProc");
            return CallWindowProcW(origProc, hwnd, msg, wp, lp);
        }
        return DefWindowProcW(hwnd, msg, wp, lp);
    }
    /* Fall through to original STATIC WndProc for all other messages. */
    WNDPROC origProc = (WNDPROC)GetPropW(hwnd, L"VB6_ImageOrigProc");
    if (origProc) return CallWindowProcW(origProc, hwnd, msg, wp, lp);
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void vb6_InstallImageSubclass(void* hwnd) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    /* Only install once */
    if (GetPropW(hw, L"VB6_ImageOrigProc")) return;
    WNDPROC origProc = (WNDPROC)SetWindowLongPtrW(hw, GWLP_WNDPROC, (LONG_PTR)vb6_ImageSubclassProc);
    if (origProc) SetPropW(hw, L"VB6_ImageOrigProc", (HANDLE)origProc);
    if (GetEnvironmentVariableA("C3_FORMS_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_F187] InstallImageSubclass hwnd=%p orig=%p installed=%d\r\n",
                (void*)hw, (void*)origProc, origProc ? 1 : 0);
}
