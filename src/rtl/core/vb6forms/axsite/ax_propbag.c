// ax_propbag.c - vb6forms_axsite 拆分片：简易 IPropertyBag: 把生成代码提供的设计期属性灌给控件 (含设计期 IFont 构造)
//
// 内容 = 拆分前 vb6forms_axsite.c 第 277~280 / 287~405 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_axsite_internal.h

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#endif

#include "vb6forms_axsite_internal.h"
#include <stdio.h>
#include <stdarg.h>

#include <stdlib.h>   /* malloc, free */
#include <stddef.h>   /* offsetof */
#include <oleauto.h>  /* SysAllocString, BSTR */
#include <olectl.h>   /* IPicture, OleLoadPicture, OLE_HANDLE */

#ifdef __cplusplus
extern "C" {
#endif

/* ===== Fix 143: 第三方 OCX 控件真宿主 ===== */

/* --- 简易 IPropertyBag: 把生成代码提供的设计期属性灌给控件 --- */


static HRESULT STDMETHODCALLTYPE pb_QueryInterface(IPropertyBag* This, REFIID riid, void** ppv) {
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IPropertyBag)) {
        *ppv = This; This->lpVtbl->AddRef(This); return S_OK;
    }
    *ppv = NULL; return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE pb_AddRef(IPropertyBag* This) {
    Vb6PropBag* b = (Vb6PropBag*)This; return InterlockedIncrement(&b->ref);
}
static ULONG STDMETHODCALLTYPE pb_Release(IPropertyBag* This) {
    Vb6PropBag* b = (Vb6PropBag*)This;
    LONG r = InterlockedDecrement(&b->ref);
    if (r == 0) free(b);
    return (ULONG)r;
}
static HRESULT STDMETHODCALLTYPE pb_Read(IPropertyBag* This, LPCOLESTR pszPropName, VARIANT* pVar, IErrorLog* pErrLog) {
    (void)pErrLog;
    Vb6PropBag* b = (Vb6PropBag*)This;
    if (!pszPropName || !pVar) return E_POINTER;
    int dbg = (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0);
    if (dbg) fprintf(stderr, "[C3_BAG] Read '%ls' wantVT=%d count=%d\n",
                     pszPropName, (int)pVar->vt, b->count);
    /* .frm 里 `_ExtentX/_ExtentY` (HIMETRIC) 是容器的账, 不在生成码的属性表里;
     * VB6 一定答得出, 这里按 SetExtent 同一尺寸补上. */
    if (wcscmp(pszPropName, L"_ExtentX") == 0 || wcscmp(pszPropName, L"_ExtentY") == 0) {
        LONG v = (pszPropName[7] == L'X') ? b->extentX : b->extentY;
        VARTYPE want = pVar->vt;
        VariantInit(pVar);
        if (want == VT_I2) { pVar->vt = VT_I2; pVar->iVal = (short)v; }
        else               { pVar->vt = VT_I4; pVar->lVal = v; }
        if (dbg) fprintf(stderr, "[C3_BAG]   hit extent %ld\n", (long)v);
        return S_OK;
    }
    for (int i = 0; i < b->count; i++) {
        const Vb6OcxProp* p = &b->props[i];
        if (wcscmp(pszPropName, p->name) != 0) continue;
        if (dbg) fprintf(stderr, "[C3_BAG]   hit '%ls' ourVT=%d\n", p->name, (int)p->vt);
        /* 请求的类型与我们的存储类型可能不同 (控件按文档类型要 VT_EMPTY=任意) */
        VARTYPE want = pVar->vt;
        /* 关键: 这里**不能** VariantClear(pVar) — pVar 是控件(调用者)提供的
         * VARIANT, 其内容未定义; 若它恰好带着垃圾 vt (如 VT_BSTR), VariantClear
         * 会按该 vt 去 SysFreeString/Release 一个野指针 → 堆损坏, 异常从 WndProc
         * 逃逸, 进程被杀 (0xC000041D). 实测 NewTab01.ocx 的 TDIMode 属性读取必崩.
         * 正确做法: 用 VariantInit 清零 (不释放任何东西) 再填值. */
        VariantInit(pVar);

        /* Fix 149: 设计期字体 (BeginProperty Font / IconFont(n)).
         * 返回 oleaut32 的真 IFont (与 VB6 的 StdFont 是同一个实现), VB6 控件
         * 拿到后可直接 .Name/.Size 读写. 返回 Nothing 会让控件解引用空指针. */
        if (p->vt == VT_DISPATCH || (p->fontName && want == VT_DISPATCH)) {
            static const GUID kIID_IFont =
                {0xBEF6E003, 0xA874, 0x101A, {0x8B, 0xBA, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};
            static const GUID kIID_IFontDisp =
                {0xBEF6E002, 0xA874, 0x101A, {0x8B, 0xBA, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};
            FONTDESC fd;
            IFont* pFont = NULL;
            memset(&fd, 0, sizeof(fd));
            fd.cbSizeofstruct = sizeof(FONTDESC);
            fd.lpstrName = (LPOLESTR)(p->fontName ? p->fontName : L"Tahoma");
            fd.cySize.int64 = (LONGLONG)((p->fontSize > 0 ? p->fontSize : 9.0) * 10000.0);
            fd.sWeight = (short)(p->fontWeight > 0 ? p->fontWeight : 400);
            fd.sCharset = (short)p->fontCharset;
            fd.fItalic = (p->fontFlags & 1) ? TRUE : FALSE;
            fd.fUnderline = (p->fontFlags & 2) ? TRUE : FALSE;
            fd.fStrikethrough = (p->fontFlags & 4) ? TRUE : FALSE;
            HRESULT hrF = OleCreateFontIndirect(&fd, &kIID_IFont, (void**)&pFont);
            if (FAILED(hrF) || !pFont)
                hrF = OleCreateFontIndirect(&fd, &kIID_IFontDisp, (void**)&pFont);
            pVar->vt = VT_DISPATCH;
            pVar->pdispVal = (IDispatch*)pFont;
            if (dbg) fprintf(stderr, "[C3_BAG]   -> IFont '%ls' %.1fpt hr=0x%08lX obj=%p\n",
                             fd.lpstrName, (double)fd.cySize.int64 / 10000.0,
                             (unsigned long)hrF, (void*)pFont);
            return S_OK;
        }

        switch (want) {
        case VT_I4:
        case VT_I2:
        case VT_UI2:
        case VT_UI4:
        case VT_INT:
        case VT_UINT:
            pVar->vt = VT_I4;
            pVar->lVal = (p->vt == VT_R4) ? (long)p->fVal : p->iVal;
            return S_OK;
        case VT_R4:
        case VT_R8:
            pVar->vt = VT_R4;
            pVar->fltVal = (p->vt == VT_R4) ? p->fVal : (float)p->iVal;
            return S_OK;
        case VT_BOOL:
            pVar->vt = VT_BOOL;
            pVar->boolVal = p->iVal ? VARIANT_TRUE : VARIANT_FALSE;
            return S_OK;
        case VT_BSTR:
            pVar->vt = VT_BSTR;
            pVar->bstrVal = (p->vt == VT_BSTR) ? SysAllocString(p->sVal)
                                               : SysAllocStringLen(NULL, 12); /* 空串 */
            if (p->vt != VT_BSTR && pVar->bstrVal) pVar->bstrVal[0] = 0;
            return S_OK;
        default:
            /* 控件要 VARIANT/其它: 按我们自己存的类型给 */
            break;
        }
        switch (p->vt) {
        case VT_BSTR:
            pVar->vt = VT_BSTR;
            pVar->bstrVal = SysAllocString(p->sVal);
            return pVar->bstrVal ? S_OK : E_OUTOFMEMORY;
        case VT_R4:
            pVar->vt = VT_R4; pVar->fltVal = p->fVal; return S_OK;
        case VT_BOOL:
            pVar->vt = VT_BOOL; pVar->boolVal = p->iVal ? VARIANT_TRUE : VARIANT_FALSE; return S_OK;
        default:
            pVar->vt = VT_I4; pVar->lVal = p->iVal; return S_OK;
        }
    }
    return E_INVALIDARG;  /* 属性不存在 — 控件用默认值 (VB6 语义) */
}
static HRESULT STDMETHODCALLTYPE pb_Write(IPropertyBag* This, LPCOLESTR pszPropName, VARIANT* pVar) {
    (void)This; (void)pszPropName; (void)pVar;
    return S_OK;  /* 只读包 */
}

const IPropertyBagVtbl g_pbVtbl = {
    pb_QueryInterface, pb_AddRef, pb_Release, pb_Read, pb_Write
};


#ifdef __cplusplus
} // extern "C"
#endif
