// vb6forms_picture.c - vb6forms 模块拆分: 图片/图标加载 + Picture 属性 + Image
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
// P13.13: Picture (PictureBox/Image)
// ============================================================

void* vb6_LoadPictureFromFile(const char* filePath) {
    if (!filePath || !filePath[0]) return NULL;
    /* Determine type by extension */
    const char* ext = strrchr(filePath, '.');
    if (!ext) ext = "";
    
    WCHAR wPath[MAX_PATH] = {0};
    // Fix 190: 内部字符串统一 UTF-8 (源文件转码而来), 路径同样 —— 用 CP_ACP
    // 会把中文/日文文件名转成乱码, LoadImageW 直接找不到文件。
    if (MultiByteToWideChar(CP_UTF8, 0, filePath, -1, wPath, MAX_PATH) <= 0)
        MultiByteToWideChar(CP_ACP, 0, filePath, -1, wPath, MAX_PATH);
    
    if (_stricmp(ext, ".ico") == 0 || _stricmp(ext, ".cur") == 0) {
        /* Icon/Cursor */
        if (_stricmp(ext, ".cur") == 0) {
            return (void*)LoadCursorFromFileW(wPath);
        }
        return (void*)LoadImageW(NULL, wPath, IMAGE_ICON, 0, 0, LR_LOADFROMFILE);
    } else if (_stricmp(ext, ".bmp") == 0) {
        return (void*)LoadImageW(NULL, wPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
    } else if (_stricmp(ext, ".emf") == 0 || _stricmp(ext, ".wmf") == 0) {
        /* Enhanced/Regular metafile - use GetEnhMetaFile/GetMetaFile */
        HENHMETAFILE hemf = GetEnhMetaFileW(wPath);
        return (void*)hemf;
    }
    /* Try as bitmap fallback */
    return (void*)LoadImageW(NULL, wPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
}

void* vb6_LoadPictureFromResource(void* hInstance, int resourceId, const char* type) {
    if (!hInstance) return NULL;
    HINSTANCE hInst = (HINSTANCE)hInstance;
    WCHAR wType[64] = {0};
    if (type) MultiByteToWideChar(CP_ACP, 0, type, -1, wType, 64);
    
    if (type && _stricmp(type, "ICON") == 0) {
        return (void*)LoadIconW(hInst, MAKEINTRESOURCEW(resourceId));
    } else if (type && _stricmp(type, "BITMAP") == 0) {
        return (void*)LoadBitmapW(hInst, MAKEINTRESOURCEW(resourceId));
    } else if (type && _stricmp(type, "CURSOR") == 0) {
        return (void*)LoadCursorW(hInst, MAKEINTRESOURCEW(resourceId));
    }
    /* Try bitmap as default */
    return (void*)LoadBitmapW(hInst, MAKEINTRESOURCEW(resourceId));
}

// ============================================================
// P24: UTF-8 to wide string conversion (caller must free())
// ============================================================
wchar_t* vb6_Utf8ToWide(const char* utf8) {
    if (!utf8) return NULL;
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    if (wlen <= 0) return NULL;
    wchar_t* wstr = (wchar_t*)malloc((size_t)wlen * sizeof(wchar_t));
    if (!wstr) return NULL;
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wstr, wlen);
    return wstr;
}

// ============================================================
// P24: Load picture from memory buffer
// Uses CreateStreamOnHGlobal + OleLoadPicture to support JPEG/BMP/ICO/PNG/GIF
// ============================================================

/* Fix <c3-menu3d-bg>: 本机 32 位进程里 OleLoadPicture 对 **任何 BMP** 一律回
 * E_FAIL (0x80004005) —— 64 位同一份代码成功, 连标准工具生成的 24bpp BMP 都
 * 失败 (对照探针 .temp/tools/picprobe 实测), 是 oleaut32 32 位 BMP 过滤器的
 * 环境性故障, 与被解的图无关。BMP 是设计期 .frx 背景图的主流格式, 必须自救:
 * 文件以 "BM" 开头时直接手动解析 BITMAPFILEHEADER + BITMAPINFO, 用
 * CreateDIBSection 建位图后整块拷入像素 —— 文件扫描行顺序与 DIB 一致
 * (bottom-up 或负高 top-down 都同序), 无需逐行翻转。其余格式 (JPEG/GIF/ICO/
 * WMF) 的过滤器不受影响, 仍走 OleLoadPicture (ICO 实测正常: 标题栏图标在)。 */
static HANDLE vb6_LoadBmpFromMemory(const void* data, int size) {
    const unsigned char* p = (const unsigned char*)data;
    if (size < 54 || p[0] != 'B' || p[1] != 'M') return NULL;
    DWORD bfOffBits = *(const DWORD*)(p + 10);
    const BITMAPINFO* bi = (const BITMAPINFO*)(p + 14);
    int w = bi->bmiHeader.biWidth;
    int h = bi->bmiHeader.biHeight;      /* 负值 = top-down, CreateDIBSection 同样接受 */
    int bpp = bi->bmiHeader.biBitCount;
    if (w <= 0 || h == 0 || bpp < 1 || bpp > 32) return NULL;
    int lines = (h > 0) ? h : -h;
    if (bi->bmiHeader.biCompression != BI_RGB &&
        bi->bmiHeader.biCompression != BI_BITFIELDS) return NULL;
    /* 调色板/掩码紧跟信息头; 像素从 bfOffBits 起 */
    if (bfOffBits < 14 + (DWORD)bi->bmiHeader.biSize) return NULL;
    int lineBytes = ((w * bpp + 31) / 32) * 4;
    if ((DWORD)bfOffBits + (DWORD)lineBytes * (DWORD)lines > (DWORD)size) return NULL;

    void* bits = NULL;
    HDC hdc = GetDC(NULL);
    HBITMAP hbmp = CreateDIBSection(hdc, bi, DIB_RGB_COLORS, &bits, NULL, 0);
    ReleaseDC(NULL, hdc);
    if (!hbmp || !bits) { if (hbmp) DeleteObject(hbmp); return NULL; }
    memcpy(bits, p + bfOffBits, (size_t)lineBytes * (size_t)lines);
    return (HANDLE)hbmp;
}

void* vb6_LoadPictureFromMemory(const void* data, int size) {
    if (!data || size <= 0) return NULL;

    /* BMP 走手动解析 (32 位 OleLoadPicture 对 BMP 环境性 E_FAIL, 见上) */
    {
        HANDLE hBmp = vb6_LoadBmpFromMemory(data, size);
        if (hBmp) return (void*)hBmp;
    }

    /* Ensure COM is initialized for OLE picture loading */
    static int oleInited = 0;
    if (!oleInited) {
        if (SUCCEEDED(OleInitialize(NULL))) {
            oleInited = 1;
        } else {
            /* Try CoInitialize as fallback */
            CoInitialize(NULL);
            oleInited = 2;
        }
    }

    /* Create IStream from memory via HGLOBAL */
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, (SIZE_T)size);
    if (!hMem) return NULL;
    void* pMem = GlobalLock(hMem);
    if (!pMem) { GlobalFree(hMem); return NULL; }
    memcpy(pMem, data, (size_t)size);
    GlobalUnlock(hMem);

    IStream* pStream = NULL;
    HRESULT hr = CreateStreamOnHGlobal(hMem, TRUE, &pStream);
    if (FAILED(hr) || !pStream) {
        GlobalFree(hMem);
        return NULL;
    }

    /* Load picture from stream */
    IPicture* pPicture = NULL;
    /* IID_IPicture = {7BF80980-BF32-101A-8BBB-00AA00300CAB} */
    static const GUID local_IID_IPicture = 
        {0x7BF80980, 0xBF32, 0x101A, {0x8B, 0xBB, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};
    hr = OleLoadPicture(pStream, (LONG)size, FALSE, &local_IID_IPicture, (void**)&pPicture);
    
    /* Stream is released by OleLoadPicture or we release it ourselves */
    pStream->lpVtbl->Release(pStream);
    /* Note: CreateStreamOnHGlobal with TRUE means HGLOBAL is freed when stream is released */

    if (FAILED(hr) || !pPicture) {
        return NULL;
    }

    /* Get the picture handle (HBITMAP for BMP/JPEG, HICON for ICO) */
    OLE_HANDLE hHandle = 0;
    pPicture->lpVtbl->get_Handle(pPicture, &hHandle);
    
    /* Get picture type to know if it's bitmap or icon */
    short picType = 0;
    pPicture->lpVtbl->get_Type(pPicture, &picType);
    
    /* For bitmaps, we need to copy the handle because IPicture owns it */
    HANDLE result = NULL;
    if (picType == 1) {
        /* PICTURE_TYPE_BITMAP - copy the bitmap */
        HBITMAP hSrc = (HBITMAP)(LONG_PTR)hHandle;
        if (hSrc) {
            BITMAP bm;
            if (GetObjectW(hSrc, sizeof(bm), &bm) != 0) {
                HDC hdcScreen = GetDC(NULL);
                HDC hdcSrc = CreateCompatibleDC(hdcScreen);
                HDC hdcDst = CreateCompatibleDC(hdcScreen);
                HBITMAP hDst = CreateCompatibleBitmap(hdcScreen, bm.bmWidth, bm.bmHeight);
                HBITMAP oldSrc = (HBITMAP)SelectObject(hdcSrc, hSrc);
                HBITMAP oldDst = (HBITMAP)SelectObject(hdcDst, hDst);
                BitBlt(hdcDst, 0, 0, bm.bmWidth, bm.bmHeight, hdcSrc, 0, 0, SRCCOPY);
                SelectObject(hdcSrc, oldSrc);
                SelectObject(hdcDst, oldDst);
                DeleteDC(hdcSrc);
                DeleteDC(hdcDst);
                ReleaseDC(NULL, hdcScreen);
                result = (HANDLE)hDst;
            }
        }
    } else if (picType == 3) {
        /* PICTURE_TYPE_ICON - must duplicate because IPicture owns it */
        HICON hSrc = (HICON)(LONG_PTR)hHandle;
        if (hSrc) {
            result = (HANDLE)DuplicateIcon(NULL, hSrc);
        }
    } else {
        /* Metafile - use handle directly (rare for .frx) */
        result = (HANDLE)(LONG_PTR)hHandle;
    }
    
    pPicture->lpVtbl->Release(pPicture);
    return (void*)result;
}

// ============================================================
// Load picture from memory buffer as COM IDispatch* (IPictureDisp)
// Returns IDispatch* suitable for passing to COM property/method calls
// e.g. ImageList.ListImages.Add(index, key, picture)
// Caller must Release the returned IDispatch* when done (via vb6_ReleaseObject)
// ============================================================
void* vb6_LoadPictureAsCom(const void* data, int size) {
    if (!data || size <= 0) return NULL;

    /* Ensure COM is initialized */
    static int oleInited2 = 0;
    if (!oleInited2) {
        if (SUCCEEDED(OleInitialize(NULL))) {
            oleInited2 = 1;
        } else {
            CoInitialize(NULL);
            oleInited2 = 2;
        }
    }

    /* Create IStream from memory */
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, (SIZE_T)size);
    if (!hMem) return NULL;
    void* pMem = GlobalLock(hMem);
    if (!pMem) { GlobalFree(hMem); return NULL; }
    memcpy(pMem, data, (size_t)size);
    GlobalUnlock(hMem);

    IStream* pStream = NULL;
    HRESULT hr = CreateStreamOnHGlobal(hMem, TRUE, &pStream);
    if (FAILED(hr) || !pStream) {
        GlobalFree(hMem);
        return NULL;
    }

    /* Load picture as IPictureDisp (IDispatch) */
    IPictureDisp* pPictureDisp = NULL;
    /* IID_IPictureDisp = {7BF80981-BF32-101A-8BBB-00AA00300CAB} */
    static const GUID local_IID_IPictureDisp =
        {0x7BF80981, 0xBF32, 0x101A, {0x8B, 0xBB, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};
    hr = OleLoadPicture(pStream, (LONG)size, FALSE, &local_IID_IPictureDisp, (void**)&pPictureDisp);

    pStream->lpVtbl->Release(pStream);

    if (FAILED(hr) || !pPictureDisp) {
        return NULL;
    }

    /* Return as IDispatch* — caller owns the reference and must Release */
    return (void*)pPictureDisp;
}
void* vb6_LoadIconFromMemory(const void* data, int size) {
    /* Load ICO data and return HICON by directly parsing .ico file format.
       .ico format: [2B reserved=0] [2B type=1] [2B count] [count*16B dir entries] [image data...]
       Each dir entry: [B w] [B h] [B colors] [B reserved] [W planes] [W bpp] [D dataSize] [D dataOffset]
    */
    if (!data || size <= 6) return NULL;
    const BYTE* p = (const BYTE*)data;
    /* Verify ICO magic */
    if (p[0] != 0 || p[1] != 0 || p[2] != 1 || p[3] != 0) {
        /* Not ICO format, fall back to OleLoadPicture */
        return vb6_LoadPictureFromMemory(data, size);
    }
    WORD count = (WORD)(p[4] | (p[5] << 8));
    if (count == 0) return NULL;
    /* Find best matching entry: prefer 32x32 or closest to SM_CXICON */
    int targetSize = GetSystemMetrics(SM_CXSMICON); /* small icon for title bar */
    int bestIdx = 0;
    int bestDiff = 9999;
    int bestBpp = 0;
    for (int i = 0; i < count && i < 20; i++) {
        const BYTE* entry = p + 6 + i * 16;
        int w = entry[0]; if (w == 0) w = 256;
        int h = entry[1]; if (h == 0) h = 256;
        int bpp = (int)(entry[6] | (entry[7] << 8));
        int diff = abs(w - targetSize) + abs(h - targetSize);
        /* Prefer higher bpp if same size */
        if (diff < bestDiff || (diff == bestDiff && bpp > bestBpp)) {
            bestDiff = diff;
            bestIdx = i;
            bestBpp = bpp;
        }
    }
    /* Read the best entry's data offset and size */
    const BYTE* bestEntry = p + 6 + bestIdx * 16;
    DWORD imgDataSize = (DWORD)(bestEntry[8] | (bestEntry[9] << 8) | (bestEntry[10] << 16) | (bestEntry[11] << 24));
    DWORD imgDataOffset = (DWORD)(bestEntry[12] | (bestEntry[13] << 8) | (bestEntry[14] << 16) | (bestEntry[15] << 24));
    if (imgDataOffset + imgDataSize > (DWORD)size || imgDataSize == 0) return NULL;
    /* CreateIconFromResourceEx expects the icon image data (BITMAPINFOHEADER + colors + XOR + AND) */
    const BYTE* pImgData = p + imgDataOffset;
    HICON hIcon = CreateIconFromResourceEx((PBYTE)pImgData, imgDataSize, TRUE, 0x00030000,
        targetSize, targetSize, LR_DEFAULTCOLOR);
    if (!hIcon) {
        /* Fallback: try with SM_CXICON */
        hIcon = CreateIconFromResourceEx((PBYTE)pImgData, imgDataSize, TRUE, 0x00030000,
            GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR);
    }
    return (void*)hIcon;
}
