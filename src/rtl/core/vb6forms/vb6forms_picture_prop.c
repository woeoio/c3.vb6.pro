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

/* Fix P-BMP-3D: vb6rtl_com.c 提供的 Picture 注册表。不直接 include
 * vb6rtl_class_com.h (本 TU 无 vb6rtl_variant.h 依赖链), 签名只涉及
 * void* 与 const, 本地 extern 声明即可完成前向声明。 */
extern void vb6_PictureRegister(void* picture);
extern int  vb6_PictureIsRegistered(const void* picture);
extern void vb6_AutoRedrawRefit(void* hwnd);   /* vb6forms_widget_prop.c (Fix c3-menu3d) */


static void vb6_ImagePaintHelper(HWND hwnd, HDC hdc);

/* Fix P-BMP-3D (AutoRedraw 后备): PictureBox 设 AutoRedraw=True 时, 图片必须
 * 同步渲染进 VB6_AutoRedrawDC 记忆位图 —— 否则 `.hdc` (vb6_ControlDrawDC 交的
 * 就是这张) 恒为空, 后续 `.Picture`/BitBlt(Me.hdc, ..., ctl.hdc, ...) 合成类
 * 程序 (如 3DMenu) 全看不到图。屏幕 WM_PAINT 与记忆 DC 各自画一份。 */
static void vb6_MirrorPictureToAutoRedraw(HWND hw) {
    HDC ar = (HDC)GetPropW(hw, L"VB6_AutoRedrawDC");
    if (ar) vb6_ImagePaintHelper(hw, ar);
}

/* Fix <vbeclipse> 2026-10-07: 窗体 WM_PAINT 的 AutoRedraw 镜像 (见 vb6forms_controls.h 说明)。
 * 只在窗体持有 VB6_AutoRedrawDC 时有动作; 无 ARDC 的窗体 (AutoRedraw=False) 原样直绘。 */
void vb6_FormPaintBlitAutoRedraw(void* hwndForm, void* hdc) {
    if (!hwndForm || !hdc) return;
    HDC ar = (HDC)GetPropW((HWND)hwndForm, L"VB6_AutoRedrawDC");
    if (!ar) return;
    RECT rc;
    if (!GetClientRect((HWND)hwndForm, &rc)) return;
    BitBlt((HDC)hdc, 0, 0, rc.right, rc.bottom, ar, 0, 0, SRCCOPY);
}

/* Fix <Command2>: 控件 Picture 的 C 面必须是**真 COM IPictureDisp**。
 *
 * VB 语义 `Property Get Picture() As IPictureDisp` 交出去的是**对象**
 * (`Picture.Handle` / `.Type` / `.Width` / `.Height` 全走 IDispatch 派发)，
 * 而 `vb6_SetControlPicture` 的裸 GDI 句柄路径 (设计期 .frx / LoadPictureFromFile)
 * 只把 HBITMAP/HICON 存在 VB6_Picture 属性里。以前 vb6_GetControlPicture 直接
 * 把这个裸句柄交出去 ⇒ 下游 vb6_ComGetProp 把它当 IDispatch，
 * vb6_ComIsDispatchable 判否 ⇒ `property "Handle" not found` ⇒ 返回 0
 * ⇒ ImageList.FListImagesAdd 里 `ElseIf Picture.Handle = NULL_PTR Then
 *    Err.Raise 35607` ⇒ 未处理错误，程序整只退出 (ComCtlsDemo Command2 实测)。
 *
 * 这里用 OleCreatePictureIndirect 把裸句柄包成真 IPictureDisp 并按窗口缓存:
 *   fOwn = FALSE —— 句柄仍归控件所有 (替换/清空/WM_DESTROY 时由控件销毁)，
 *   包装对象只是借用，不会二次 DeleteObject。
 *   缓存的那份引用带**一枚保险引用**: 发码层对 `Set <控件>.Picture = Nothing`
 *   历史性地发成 `vb6_ReleaseObject((void**)(&(void*){vb6_GetControlPicture(..)}))`
 *   (借用引用被 Release)，多留一枚可吸收这次过度释放而不让缓存指针悬空；
 *   保险引用在替换/清空/销毁时与缓存引用一并归还。
 */
static IPictureDisp* vb6_WrapGdiHandleAsPictureDisp(HWND hw, HANDLE hGdi) {
    if (!hw || !hGdi) return NULL;
    IPictureDisp* pDisp = (IPictureDisp*)GetPropW(hw, L"VB6_IPictureDisp");
    if (pDisp) return pDisp;

    int picType = (int)(INT_PTR)GetPropW(hw, L"VB6_PictureType");
    DWORD objType = GetObjectType((HGDIOBJ)hGdi);

    PICTDESC pd;
    memset(&pd, 0, sizeof(pd));
    pd.cbSizeofstruct = sizeof(pd);
    if (picType == 3 /*PICTYPE_ICON*/) {
        pd.picType = PICTYPE_ICON;
        pd.icon.hicon = (HICON)hGdi;
    } else if (objType == OBJ_ENHMETAFILE) {
        pd.picType = PICTYPE_ENHMETAFILE;
        pd.emf.hemf = (HENHMETAFILE)hGdi;
    } else if (objType == OBJ_METAFILE) {
        pd.picType = PICTYPE_METAFILE;
        pd.wmf.hmeta = (HMETAFILE)hGdi;
        pd.wmf.xExt = 0;
        pd.wmf.yExt = 0;
    } else {
        pd.picType = PICTYPE_BITMAP;
        pd.bmp.hbitmap = (HBITMAP)hGdi;
    }

    static int oleInitedPic = 0;
    if (!oleInitedPic) {
        if (SUCCEEDED(OleInitialize(NULL))) oleInitedPic = 1;
        else { CoInitialize(NULL); oleInitedPic = 2; }
    }

    /* IID_IPictureDisp = {7BF80981-BF32-101A-8BBB-00AA00300CAB}
     * 与 vb6forms_picture.c 的 OleLoadPicture 同一枚本地定义 (本 TU 的头文件链
     * 不保证带出该符号)。 */
    static const GUID local_IID_IPictureDisp =
        {0x7BF80981, 0xBF32, 0x101A, {0x8B, 0xBB, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};

    IPictureDisp* pNew = NULL;
    HRESULT hr = OleCreatePictureIndirect(&pd, &local_IID_IPictureDisp, FALSE,
                                          (void**)&pNew);
    if (FAILED(hr) || !pNew) return NULL;
    /* 登进 Picture 注册表 (Fix P-BMP-3D): 这样 `Set 另一枚控件.Picture = 本控件.Picture`
     * 时 vb6_SetControlPicture 能认出这是 COM 对象而不是 GDI 句柄，正确走 FromCom。 */
    vb6_PictureRegister((void*)pNew);
    /* 缓存引用 + 保险引用 (见上方注释) */
    pNew->lpVtbl->AddRef(pNew);
    if (!SetPropW(hw, L"VB6_IPictureDisp", (HANDLE)pNew)) {
        pNew->lpVtbl->Release(pNew);
        pNew->lpVtbl->Release(pNew);
        return NULL;
    }
    return pNew;
}

/* 存一枚 IPictureDisp 到窗口 (含保险引用); 旧的那枚连保险引用一并归还。 */
static void vb6_StorePictureDisp(HWND hw, IPictureDisp* pDisp) {
    if (!hw) return;
    IPictureDisp* pOld = (IPictureDisp*)GetPropW(hw, L"VB6_IPictureDisp");
    if (pDisp) pDisp->lpVtbl->AddRef(pDisp);
    if (!SetPropW(hw, L"VB6_IPictureDisp", (HANDLE)pDisp)) {
        if (pDisp) pDisp->lpVtbl->Release(pDisp);
        return;
    }
    if (pOld) { pOld->lpVtbl->Release(pOld); pOld->lpVtbl->Release(pOld); }
}

/* 归还并清掉窗口上的 IPictureDisp 缓存 (缓存引用 + 保险引用)。 */
static void vb6_ClearPictureDisp(HWND hw) {
    if (!hw) return;
    IPictureDisp* pOld = (IPictureDisp*)GetPropW(hw, L"VB6_IPictureDisp");
    if (pOld) {
        RemovePropW(hw, L"VB6_IPictureDisp");
        pOld->lpVtbl->Release(pOld);
        pOld->lpVtbl->Release(pOld);
    }
}

void* vb6_GetControlPicture(void* hwnd) {
    if (!hwnd) return NULL;
    HWND hw = (HWND)hwnd;
    /* If a COM IPictureDisp was stored via SetControlPictureFromCom, return the
     * borrowed COM pointer so callers (e.g. Clipboard.SetData) can QI/use it.
     * The caller must NOT Release this borrowed reference. */
    IPictureDisp* pDisp = (IPictureDisp*)GetPropW(hw, L"VB6_IPictureDisp");
    if (pDisp) return (void*)pDisp;
    /* Fix <Command2>: 裸 GDI 句柄也要包成真 IPictureDisp 再交出去 ——
       VB 侧 `Picture As IPictureDisp` 是对象，交裸句柄会让下游 COM 读全部
       "not found" (实测 ImageList.FListImagesAdd → Err.Raise 35607 → 退出)。 */
    HANDLE hProp = GetPropW(hw, L"VB6_Picture");
    if (hProp) return (void*)vb6_WrapGdiHandleAsPictureDisp(hw, hProp);
    return NULL;
}

void vb6_SetControlPicture(void* hwnd, void* hPicture) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    
    /* Picture 双支持 (Fix P-BMP-3D): 本管线 LoadPicture 产出的 COM IPicture*
     * 走 COM 子类化渲染路径; 原始 GDI 句柄 (设计期 frx/资源/API) 走经典
     * STATIC 路径。二者在 vb6rtl_com.c 的注册表里可安全区分。 */
    if (hPicture && vb6_PictureIsRegistered(hPicture)) {
        vb6_SetControlPictureFromCom(hwnd, hPicture);
        return;
    }
    /* A raw handle is replacing any COM IPicture set earlier via
     * SetControlPictureFromCom. Release the retained COM reference so it
     * doesn't leak and won't be used by the subclass. */
    IPicture* pOldCom = (IPicture*)GetPropW(hw, L"VB6_IPicture");
    if (pOldCom) {
        pOldCom->lpVtbl->Release(pOldCom);
        RemovePropW(hw, L"VB6_IPicture");
    }
    /* Fix <Command2>: 句柄换了，先前懒包出来的 IPictureDisp 缓存作废
     * (它借用的是旧句柄)，连保险引用一并归还。 */
    vb6_ClearPictureDisp(hw);
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

    vb6_MirrorPictureToAutoRedraw(hw);

    /* AutoSize: if enabled, resize to fit picture */
    int autoSize = vb6_GetPictureAutoSize(hwnd);
    if (autoSize && hPicture) {
        BITMAP bm;
        HBITMAP hBmp = (HBITMAP)hPicture;
        if (GetObjectW(hBmp, sizeof(bm), &bm) != 0) {
            SetWindowPos(hw, NULL, 0, 0, bm.bmWidth, bm.bmHeight,
                SWP_NOMOVE | SWP_NOZORDER);
            vb6_AutoRedrawRefit(hwnd);   /* ARDC 还是设计期小图 → 按新客户区重建 (Fix c3-menu3d) */
            vb6_MirrorPictureToAutoRedraw(hw);   /* 上面第 132 行那次镜像画在旧小图上,
                                                    refit 清掉了 → 按新尺寸重画一遍 */
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
        vb6_ClearPictureDisp(hw);
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

    /* Fix <Command2>: 同时存一份真 IPictureDisp (QI 自传入指针)，让
     * vb6_GetControlPicture 交出去的是 IDispatch 而不是 IPicture* ——
     * 两者的 vtable 槽位不同，把 IPicture* 当 IDispatch 用会让
     * GetIDsOfNames 落到 IPicture::get_Type 上 (槽 5) 造成类型混淆。 */
    {
        static const GUID local_IID_IPictureDisp2 =
            {0x7BF80981, 0xBF32, 0x101A, {0x8B, 0xBB, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};
        IPictureDisp* pDispIn = NULL;
        HRESULT hrD = pUnk->lpVtbl->QueryInterface(pUnk, &local_IID_IPictureDisp2,
                                                   (void**)&pDispIn);
        if (SUCCEEDED(hrD) && pDispIn) {
            vb6_StorePictureDisp(hw, pDispIn);
            pDispIn->lpVtbl->Release(pDispIn);   /* 缓存里已含引用 */
        }
    }

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

    /* AutoSize (Fix P-BMP-3D): 与原始句柄路径同口径 —— PictureBox/Image 设
     * AutoSize=True 时, 指派图片后控件缩放到图片本征尺寸 (VB6 行为)。 */
    if (nType == 1 /* PICTYPE_BITMAP */ && vb6_GetPictureAutoSize(hwnd)) {
        BITMAP bm;
        if (GetObjectW((HBITMAP)(UINT_PTR)hOle, sizeof(bm), &bm) != 0) {
            SetWindowPos(hw, NULL, 0, 0, bm.bmWidth, bm.bmHeight,
                SWP_NOMOVE | SWP_NOZORDER);
            vb6_AutoRedrawRefit(hwnd);   /* ARDC 还是设计期小图 → 按新客户区重建 */
        }
    }

vb6_InstallImageSubclass(hwnd);
    InvalidateRect(hw, NULL, TRUE);

    /* Fix <c3-menu3d>: COM 路径的 AutoRedraw 镜像缺位。PictureBox 设 Visible=False
     * (3DMenu Form_Load 里 With ImgMenu(i) : .Visible=False) 时 WM_PAINT 永不触发,
     * 子类里的 arMirror 分支 (WM_PAINT 尾部) 没机会跑, 而 RuotaMenu 用
     * ImgMenu(Num).hdc (= VB6_AutoRedrawDC) 作 TransBltNow 源 → 合成出来恒空。
     * 原句柄路径 (vb6_SetControlPicture) 第 110 行已调用本镜像; COM 路径在装完
     * 子类后同样主动画一遍进记忆 DC。屏上路径 (WM_PAINT/WM_PRINTCLIENT) 照旧。 */
vb6_MirrorPictureToAutoRedraw(hw);
    {   HDC arDbg = (HDC)GetPropW(hw, L"VB6_AutoRedrawDC");
        IPicture* pDbg = (IPicture*)GetPropW(hw, L"VB6_IPicture");
        
    }
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
    /* alpha 全 0: 两种可能 —— (a) 真·全透明 PNG; (b) 32bpp BI_RGB, alpha 字节
     * 从未写入 (GDI/OleLoadPicture 产物的常态, GetDIBits 读回全 0)。后者 RGB 是
     * 实打实的 (3DMenu 图标角=FAE4E4 trsp 色实测), 按"整张透明"丢弃会让整个
     * PictureBox 的 ARDC 恒 BackColor → 环格全空 → 白窗 (Fix <c3-menu3d>)。
     * VB6 的 BitBlt 从不看 alpha —— 全 0 一律按不透明回落原路径, 透明色交给
     * 调用方 (TransBltNow 按角像素色键) 处理, 与金标准同口径。 */
    if (allZero) { free(buf); return 0; }   /* 无 alpha 信息: 回落不透明原路径 */
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
            /* Non-stretch: draw at the picture's natural pixel size. The conversion
             * formula lives in exactly one place now (vb6_HimetricToPxOnDc, zhang 298);
             * this call site keeps feeding it the **DC's** DPI, not the screen's --
             * a picture drawn into a printer/VM DC scales with that DC. */
            int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
            int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
            dstW = vb6_HimetricToPxOnDc((int)hmW, dpiX);
            dstH = vb6_HimetricToPxOnDc((int)hmH, dpiY);
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
            { static int s_h = 0; if (s_h < 40) { s_h++;
                 } }
            return;
        }

        /* Metafiles / QR codes typically expect a white background. */
        FillRect(hdc, &rc, (HBRUSH)GetStockObject(WHITE_BRUSH));

        /* HIMETRIC y points up, so flip the source rect: (0, hmH, hmW, -hmH). */
        HRESULT hrR = pPic->lpVtbl->Render(pPic, hdc, 0, 0, dstW, dstH,
                             0, hmH, hmW, -hmH, NULL);
        { static int s_h2 = 0; if (s_h2 < 40) { s_h2++;
             } }
        return;
    }

    HANDLE hPict = GetPropW(hwnd, L"VB6_Picture");
    int picType = (int)(INT_PTR)GetPropW(hwnd, L"VB6_PictureType");
    if (!hPict) {
        static int s_h3 = 0; if (s_h3 < 40) { s_h3++;
             }
        return;
    }

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
            /* Fix P-BMP-3D: 屏幕画完再补一张进 AutoRedraw 记忆位图, 保证
             * `.hdc` 来源 (VB6_AutoRedrawDC) 与所见一致 — 也覆盖 AutoRedraw
             * 在 .Picture 指派之后才建记忆 DC 的场景。 */
            HDC arMirror = (HDC)GetPropW(hwnd, L"VB6_AutoRedrawDC");
            if (arMirror && arMirror != hdc) vb6_ImagePaintHelper(hwnd, arMirror);
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
        /* Fix <Command2>: 懒包出来的 IPictureDisp 缓存也在这里归还 */
        vb6_ClearPictureDisp(hwnd);
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
