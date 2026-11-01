// vb6forms_axcontainer.c - Fix 148: VB6 容器对象模型 (Extender / Container / Controls)
//
// VB6 UserControl 通过容器提供的对象访问宿主窗体:
//   Set iContainerUsr = UserControl.Extender.Container   ' 容器
//   Set iControls     = UserControl.Parent.Controls      ' 窗体控件集合
//   Set iCtl          = UserControl.Parent.Controls("txtDoc(0)", tabIdx)
//   iTiUsr            = UserControl.Extender.TabIndex
// 用于把设计期放在 Tab 页上的控件 SetParent 到 Tab 页 (NewTab 的 TDIMode).
//
// 此前 IOleControlSite::GetExtendedControl 返回 E_NOTIMPL, 而 VB6 控件在
// IPersistPropertyBag::Load 期间就会访问容器 (读 TDIMode 属性时) →
// MSVBVM60 内部解引用空指针 → 异常从 WndProc 逃逸 → 进程被杀 (0xC000041D).
//
// 三个对象都是轻量 IDispatch, 用 kind 区分行为; Controls.Item(name) 返回的
// 控件对象直接复用 RTL 的"宿主对象即 hwnd"约定 (hwnd 值当 IDispatch*).

#include <windows.h>
#include <oleauto.h>
#include <olectl.h>      /* Fix 149: FONTDESC / OleCreateFontIndirect (容器字体) */
#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6rtl.h"      /* Fix 148: vb6_VARIANT 定义 */
#include "vb6forms_uc_internal.h"  /* Fix 163: vb6_HostObj_At / UC_ControlsItemByName
                                    * 返回 void* —— 本 TU 此前没有它们的原型, MSVC 按
                                    * 隐式 `int` 声明编译 → x64 下 64 位指针被截成 32 位
                                    * 再赋给 void* h (315/332 两处), Form.Controls 于是
                                    * 拿到坏句柄。x86 下 pointer==int 侥幸不发作。 */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>

#define VB6_AXK_EXTENDER   1
#define VB6_AXK_CONTAINER  2
#define VB6_AXK_CONTROLS   3

typedef struct Vb6AxObj {
    IDispatchVtbl* lpVtbl;
    LONG ref;
    int32_t kind;
    void* hwndForm;
    wchar_t ctrlName[128];
} Vb6AxObj;

static const IDispatchVtbl g_axObjVtbl;
static const IDispatchVtbl g_axExtVtbl;

/* --- 名字表 (DISPID = 下标+1) --- */
#define VB6_AXDISP_SYNTH 1000   /* Fix 149: 未知成员 → 合成分派号 (Invoke 回默认值) */
static const wchar_t* kExtNames[] = {
    L"Container", L"Parent", L"Name", L"TabIndex", L"TabStop", L"Visible",
    L"Enabled", L"Left", L"Top", L"Width", L"Height", L"Tag", L"HelpContextID",
    L"Object", L"ToolTipText", L"OLEDropMode", L"WhatsThisHelpID", L"PasswordChar",
    L"ScaleMode", L"hWnd", L"Font", L"BackColor", L"ForeColor", L"Move", NULL };
/* 容器 = VB6 窗体对象。控件代码用到的: Container.ScaleMode / .Controls(...) /
 * .hWnd / .Name / .Caption / .Font 等 (ctlNewTab.ctl:12317 起) */
static const wchar_t* kConNames[] = {
    L"Controls", L"Name", L"Caption", L"hWnd", L"BackColor", L"ForeColor",
    L"Visible", L"Enabled", L"Width", L"Height", L"Left", L"Top",
    L"ScaleMode", L"ScaleWidth", L"ScaleHeight", L"ScaleLeft", L"ScaleTop",
    L"Parent", L"Font", L"Tag", L"WindowState", L"MDIChild", L"AutoRedraw",
    L"Appearance", L"BorderStyle", L"DrawWidth", L"DrawStyle", L"FillColor",
    L"FillStyle", L"MousePointer", L"MouseIcon", L"KeyPreview", L"Picture",
    L"hDC", L"ClipControls", L"ControlBox", L"MaxButton", L"MinButton",
    L"Moveable", L"Zoom", L"CurrentX", L"CurrentY", L"RightToLeft",
    L"StartUpPosition", L"Refresh", L"SetFocus", L"ZOrder", L"Move", L"Cls",
    L"Print", L"Hide", L"Show", L"PopupMenu", L"Circle", L"Line", L"PSet",
    L"Point", L"TextWidth", L"TextHeight", L"Scale", NULL };
static const wchar_t* kColNames[] = { L"Item", L"Count", L"_NewEnum", NULL };

static int axCount(const wchar_t* const* tbl) { int n = 0; while (tbl[n]) n++; return n; }

/* VB6 容器(窗体)的字体: 真 IFont (控件用 Container.Font 取默认字体) */
static void* axMakeFont(const wchar_t* name, float pt, long weight) {
    static const GUID kIID_IFont =
        {0xBEF6E003, 0xA874, 0x101A, {0x8B, 0xBA, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};
    FONTDESC fd;
    IFont* pf = NULL;
    memset(&fd, 0, sizeof(fd));
    fd.cbSizeofstruct = sizeof(FONTDESC);
    fd.lpstrName = (LPOLESTR)(name ? name : L"Tahoma");
    fd.cySize.int64 = (LONGLONG)((pt > 0 ? pt : 9.0) * 10000.0);
    fd.sWeight = (short)(weight > 0 ? weight : 400);
    if (FAILED(OleCreateFontIndirect(&fd, &kIID_IFont, (void**)&pf))) return NULL;
    return pf;
}

static int axLookup(const wchar_t* const* tbl, const wchar_t* name) {
    if (!name) return 0;
    for (int i = 0; tbl[i]; i++)
        if (_wcsicmp(tbl[i], name) == 0) return i + 1;
    return 0;
}

static void axSetEmpty(vb6_VARIANT* o) { memset(o, 0, sizeof(*o)); }
static void axSetI4(vb6_VARIANT* o, int32_t v) { axSetEmpty(o); o->vt = vb6_vtLong; o->lVal = v; }
static void axSetBool(vb6_VARIANT* o, int32_t v) { axSetEmpty(o); o->vt = vb6_vtBoolean; o->boolVal = (int16_t)(v ? -1 : 0); }
static void axSetStr(vb6_VARIANT* o, const wchar_t* s) {
    axSetEmpty(o);
    o->vt = vb6_vtBSTR;
    o->bstrVal = SysAllocString(s ? s : L"");
}
static void axSetObj(vb6_VARIANT* o, void* iface) {
    axSetEmpty(o);
    o->vt = vb6_vtDispatch;
    o->pdispVal = (void*)iface;
}

/* --- 构造 --- */
static Vb6AxObj* axNew(int32_t kind, void* hwndForm, const wchar_t* ctrlName) {
    Vb6AxObj* p = (Vb6AxObj*)calloc(1, sizeof(Vb6AxObj));
    if (!p) return NULL;
    p->lpVtbl = (kind == VB6_AXK_EXTENDER) ? &g_axExtVtbl : &g_axObjVtbl;
    p->ref = 1;
    p->kind = kind;
    p->hwndForm = hwndForm;
    if (ctrlName) { wcsncpy(p->ctrlName, ctrlName, 127); p->ctrlName[127] = 0; }
    return p;
}

/* --- IUnknown (两种 vtable 共用同一批实现) --- */
static HRESULT STDMETHODCALLTYPE ax_QI(IDispatch* This, REFIID riid, void** ppv) {
    Vb6AxObj* p = (Vb6AxObj*)This;
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = This; This->lpVtbl->AddRef(This); return S_OK;
    }
    /* VB6 的 Extender 是**聚合**对象: 不认识的 IID 转给控件本体。Codejock 的
     * PushButton 解析 BuddyControl 时正是拿 IOleControlSite::GetExtendedControl 的
     * 返回值 QI 它自己的私有 IID {C0DE2030-31B8-4F06-99A6-17A150A6FBE2}, 再 RTTI
     * 转型成 buddy 的内部对象; 拿到后 [buddy+0x280] 非 0 才把按钮并入 buddy 的边框
     * (Left -= 2, Width += 2)。以前这里恒答 E_NOINTERFACE ⇒ 内部对象 NULL ⇒ 合并
     * 块整段跳过, 按钮只有白面、宽恒 16px (参考 exe 18px, 并且 Form_Resize 里
     * `FrameBorder.Width - txtSearchBar.Width - 10 - cmdSearch.Width` 跟着差 2px)。 */
    if (p->kind == VB6_AXK_EXTENDER) {
        void* disp = vb6_OcxHost_DispByName(p->hwndForm, p->ctrlName);
        if (disp) {
            HRESULT hr = ((IUnknown*)disp)->lpVtbl->QueryInterface((IUnknown*)disp, riid, ppv);
            if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
                wchar_t g[64] = {0}; StringFromGUID2(riid, g, 64);
                fprintf(stderr, "[C3_AXOBJ]   ext QI forward %ls -> 0x%08lX ctrl='%ls'\n",
                        g, (unsigned long)hr, p->ctrlName);
                fflush(stderr);
            }
            if (hr == S_OK) return S_OK;
        }
    }
    *ppv = NULL; return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE ax_AddRef(IDispatch* This) {
    Vb6AxObj* p = (Vb6AxObj*)This; return InterlockedIncrement(&p->ref);
}
static ULONG STDMETHODCALLTYPE ax_Release(IDispatch* This) {
    Vb6AxObj* p = (Vb6AxObj*)This;
    LONG r = InterlockedDecrement(&p->ref);
    if (r <= 0) free(p);
    return (ULONG)r;
}
static HRESULT STDMETHODCALLTYPE ax_GetTypeInfoCount(IDispatch* This, UINT* p) { (void)This; *p = 0; return S_OK; }
static HRESULT STDMETHODCALLTYPE ax_GetTypeInfo(IDispatch* This, UINT i, LCID l, ITypeInfo** p) {
    (void)This; (void)i; (void)l; *p = NULL; return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE ax_GetIDsOfNames(IDispatch* This, REFIID riid, LPOLESTR* names,
                                                  UINT cNames, LCID lcid, DISPID* dispids) {
    (void)riid; (void)lcid;
    Vb6AxObj* p = (Vb6AxObj*)This;
    if (!names || !dispids) return E_POINTER;
    const wchar_t* const* tbl = (p->kind == VB6_AXK_EXTENDER) ? kExtNames
                             : (p->kind == VB6_AXK_CONTAINER) ? kConNames : kColNames;
    int axDbg = (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0);
    for (UINT i = 0; i < cNames; i++) {
        /* VB6 Extender.Index 只存在于控件数组; 非数组控件上该成员不存在.
         * Codejock 的 GetVBControlName (OCX 0x101FE830) 读完 Name 后必读 Index:
         * 只有 Index 取不到 (DISPID_UNKNOWN/-1) 时才用纯 Name 参与 BuddyControl
         * 的按名比较; 否则拼成 "Name(Index)" (Format "%s(%i)") 再比较, 于是
         * 名字再一致也匹配不上. 这里返回 DISPID_UNKNOWN 与 VB6 一致. */
        if (p->kind == VB6_AXK_EXTENDER && names[i] && _wcsicmp(names[i], L"Index") == 0) {
            dispids[i] = DISPID_UNKNOWN;
            continue;
        }
        int d = axLookup(tbl, names[i]);
        /* Fix 149: 容器/Extender 是 VB6 宿主提供的"大对象"(真身是 Form),
         * 成员成百上千, 不可能枚举完。未知成员给合成分派号 + Invoke 回默认值,
         * 而不是 DISPID_UNKNOWN —— 后者实测会让 VB6 报 "Run-time error '5':
         * Invalid procedure call or argument" (控件读 Container.ScaleMode 时)。
         * 真正需要的成员仍会打印 MISS, 便于逐个补全。 */
        dispids[i] = d ? (DISPID)d : (DISPID)VB6_AXDISP_SYNTH;
        if (axDbg && !names[i])
            fprintf(stderr, "[C3_AXOBJ] GetIDsOfNames kind=%d cNames=%u name[%u]=NULL -> synth\n",
                    p->kind, cNames, i);
        if (axDbg && !d)
            fprintf(stderr, "[C3_AXOBJ] MISS kind=%d(%s) name='%ls'\n",
                    p->kind, p->kind == VB6_AXK_EXTENDER ? "Ext" :
                            p->kind == VB6_AXK_CONTAINER ? "Con" : "Col",
                    names[i] ? names[i] : L"?");
    }
    return S_OK;
}

/* Extender 几何: 控件窗口 (windowless 则站点账) 的 Left/Top/Width/Height,
 * 折成容器单位 (窗体 ScaleMode, 默认缇) 交出去。0 = 不是本宿主托管的控件。 */
static int axExtGeometry(Vb6AxObj* p, const wchar_t* nm, int32_t* out) {
    void* disp = vb6_OcxHost_DispByName(p->hwndForm, p->ctrlName);
    if (!disp) return 0;
    int32_t r[4] = {0, 0, 0, 0}; void* form = NULL;
    HWND h = (HWND)vb6_OcxHost_WindowOf(disp);
    if (h && IsWindow(h)) {
        RECT rc; GetWindowRect(h, &rc);
        POINT p0 = { rc.left, rc.top }, p1 = { rc.right, rc.bottom };
        MapWindowPoints(NULL, (HWND)p->hwndForm, &p0, 1);
        MapWindowPoints(NULL, (HWND)p->hwndForm, &p1, 1);
        r[0] = p0.x; r[1] = p0.y; r[2] = p1.x - p0.x; r[3] = p1.y - p0.y;
        form = p->hwndForm;
    } else if (!vb6_OcxHost_GetRect(disp, r, &form)) {
        return 0;
    }
    int idx = 0, vert = 0;
    if (_wcsicmp(nm, L"Top") == 0)         { idx = 1; vert = 1; }
    else if (_wcsicmp(nm, L"Width") == 0)  { idx = 2; }
    else if (_wcsicmp(nm, L"Height") == 0) { idx = 3; vert = 1; }
    *out = (int32_t)vb6_ScalePxToUser((double)r[idx], vb6_ContainerScaleMode(form), vert);
    return 1;
}

/* --- Invoke --- */
static HRESULT STDMETHODCALLTYPE ax_Invoke(IDispatch* This, DISPID dispid, REFIID riid, LCID lcid,
                                           WORD flags, DISPPARAMS* dp, VARIANT* result,
                                           EXCEPINFO* ei, UINT* ae) {
    (void)riid; (void)lcid; (void)ei; (void)ae;
    Vb6AxObj* p = (Vb6AxObj*)This;
    int axDbg = (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0);
    if (axDbg) {
        const wchar_t* const* tb = (p->kind == VB6_AXK_EXTENDER) ? kExtNames
                                : (p->kind == VB6_AXK_CONTAINER) ? kConNames : kColNames;
        int mx = axCount(tb);
        const wchar_t* onm = ((int)dispid >= 1 && (int)dispid <= mx) ? tb[(int)dispid - 1] : L"?";
        fprintf(stderr, "[C3_AXOBJ] kind=%d(%s) invoke '%ls' dispid=%ld flags=0x%04X cArgs=%u named=%u res=%d\n",
                p->kind, p->kind == VB6_AXK_EXTENDER ? "Ext" : p->kind == VB6_AXK_CONTAINER ? "Con" : "Col",
                onm, (long)dispid, (unsigned)flags, dp ? dp->cArgs : 0, dp ? dp->cNamedArgs : 0,
                result ? 1 : 0);
        fflush(stderr);
    }
    /* 方法调用时调用方 result 常为 NULL (IDispatch::Invoke(DISPATCH_METHOD, ...)).
     * 以前在这里直接 return, 于是**所有**无返回值的成员 (Move/Refresh/SetFocus …)
     * 都被静默丢弃 (实测 Codejock PushButton 的 Extender.Move 打到这里就没了)。
     * 给一份局部 dummy 继续走, 副作用照做, 返回值丢弃即可。 */
    VARIANT dummy;
    if (!result) { memset(&dummy, 0, sizeof(dummy)); result = &dummy; }
    memset(result, 0, sizeof(*result));
    vb6_VARIANT* out = (vb6_VARIANT*)result;
    int idx = (int)dispid;

    /* 未知成员 (合成分派号): 回默认值, 不报错 */
    if (idx >= VB6_AXDISP_SYNTH) { axSetEmpty(out); return S_OK; }

    if (p->kind == VB6_AXK_EXTENDER) {
        const wchar_t* nm = (idx >= 1 && idx <= axCount(kExtNames)) ? kExtNames[idx - 1] : NULL;
        if (!nm) return DISP_E_MEMBERNOTFOUND;
        if (_wcsicmp(nm, L"Container") == 0 || _wcsicmp(nm, L"Parent") == 0) {
            Vb6AxObj* c = axNew(VB6_AXK_CONTAINER, p->hwndForm, NULL);
            if (!c) return E_OUTOFMEMORY;
            axSetObj(out, c);
            return S_OK;
        }
        if (_wcsicmp(nm, L"Name") == 0)     {
            if (axDbg) { fprintf(stderr, "[C3_AXOBJ]   ext Name -> '%ls'\n", p->ctrlName); fflush(stderr); }
            axSetStr(out, p->ctrlName); return S_OK;
        }
        if (_wcsicmp(nm, L"TabIndex") == 0) { axSetI4(out, 0); return S_OK; }
        if (_wcsicmp(nm, L"TabStop") == 0 || _wcsicmp(nm, L"Visible") == 0
            || _wcsicmp(nm, L"Enabled") == 0) { axSetBool(out, 1); return S_OK; }
        if (_wcsicmp(nm, L"Tag") == 0)      { axSetStr(out, L""); return S_OK; }
        if (_wcsicmp(nm, L"Font") == 0) {
            void* f = axMakeFont(L"Tahoma", 9.0f, 400);
            if (!f) { axSetEmpty(out); return S_OK; }
            axSetObj(out, f);
            return S_OK;
        }
        if (_wcsicmp(nm, L"BackColor") == 0 || _wcsicmp(nm, L"ForeColor") == 0) {
            axSetI4(out, 0x8000000F); return S_OK;
        }
        if (_wcsicmp(nm, L"hWnd") == 0) { axSetI4(out, (int32_t)(intptr_t)p->hwndForm); return S_OK; }
        if (_wcsicmp(nm, L"ScaleMode") == 0) { axSetI4(out, 1); return S_OK; }  /* vbTwips */
        /* 几何: 走容器 Extender 的真账 (控件窗口), 不再一律回 0.
         * Codejock PushButton 的 BuddyControl 解析正是: 读 buddy 的
         * Left/Top/Height/Width → 算好位置 → Extender.Move 贴靠。以前 Move 不在
         * 名字表、几何恒 0, 控件解析到一半就放弃 ⇒ 按钮退化成"只有白面" ;
         * 参考 exe (VB6 Extender 齐全) 画出上/右/下三边 1px #ABABAB。 */
        if (_wcsicmp(nm, L"Left") == 0 || _wcsicmp(nm, L"Top") == 0
            || _wcsicmp(nm, L"Width") == 0 || _wcsicmp(nm, L"Height") == 0) {
            int32_t v = 0;
            if (axExtGeometry(p, nm, &v)) { axSetI4(out, v); return S_OK; }
        }
        /* VB6 语言级方法: Move Left[, Top[, Width[, Height]]] (容器单位, 省略的保持) */
        if (_wcsicmp(nm, L"Move") == 0) {
            double a[4] = {0, 0, 0, 0};
            int mask = 0;
            for (int k = 0; k < 4 && dp && (UINT)k < dp->cArgs; k++) {
                VARIANT* av = &dp->rgvarg[dp->cArgs - 1 - k];   /* VB6 实参逆序 */
                short vt = V_VT(av);
                if (vt == VT_ERROR || vt == VT_EMPTY || vt == VT_NULL) continue;  /* 省略 */
                if (vt == VT_I2) a[k] = (double)V_I2(av);
                else if (vt == VT_I4) a[k] = (double)V_I4(av);
                else if (vt == VT_R4) a[k] = (double)V_R4(av);
                else if (vt == VT_R8) a[k] = V_R8(av);
                else if (vt == VT_BSTR) a[k] = wcstod(V_BSTR(av), NULL);
                else if (vt == VT_BOOL) a[k] = V_BOOL(av) ? -1.0 : 0.0;
                else continue;
                mask |= 1 << k;
            }
            if (axDbg) {
                fprintf(stderr, "[C3_AXOBJ]   ext Move mask=%d (%.0f, %.0f, %.0f, %.0f)\n",
                        mask, a[0], a[1], a[2], a[3]);
                fflush(stderr);
            }
            void* disp = vb6_OcxHost_DispByName(p->hwndForm, p->ctrlName);
            if (disp) vb6_ControlMove(disp, a[0], a[1], a[2], a[3], mask);
            axSetEmpty(out);
            return S_OK;
        }
        axSetI4(out, 0);
        return S_OK;
    }

    if (p->kind == VB6_AXK_CONTAINER) {
        const wchar_t* nm = (idx >= 1 && idx <= axCount(kConNames)) ? kConNames[idx - 1] : NULL;
        if (!nm) return DISP_E_MEMBERNOTFOUND;
        /* 对象属性 */
        if (_wcsicmp(nm, L"Controls") == 0) {
            Vb6AxObj* c = axNew(VB6_AXK_CONTROLS, p->hwndForm, NULL);
            if (!c) return E_OUTOFMEMORY;
            axSetObj(out, c);
            return S_OK;
        }
        if (_wcsicmp(nm, L"Font") == 0) {
            void* f = axMakeFont(L"Tahoma", 9.0f, 400);
            if (!f) { axSetEmpty(out); return S_OK; }
            axSetObj(out, f);
            return S_OK;
        }
        /* 字符串属性 */
        if (_wcsicmp(nm, L"Name") == 0 || _wcsicmp(nm, L"Caption") == 0) {
            wchar_t cap[256] = {0};
            if (p->hwndForm) GetWindowTextW((HWND)p->hwndForm, cap, 256);
            axSetStr(out, cap);
            return S_OK;
        }
        if (_wcsicmp(nm, L"Tag") == 0) { axSetStr(out, L""); return S_OK; }
        /* 几何: 以缇为单位 (ScaleMode=1 时的 VB6 语义) */
        if (_wcsicmp(nm, L"hWnd") == 0) { axSetI4(out, (int32_t)(intptr_t)p->hwndForm); return S_OK; }
        if (_wcsicmp(nm, L"Width") == 0 || _wcsicmp(nm, L"Height") == 0
            || _wcsicmp(nm, L"Left") == 0 || _wcsicmp(nm, L"Top") == 0
            || _wcsicmp(nm, L"ScaleWidth") == 0 || _wcsicmp(nm, L"ScaleHeight") == 0) {
            RECT r = {0, 0, 0, 0};
            if (p->hwndForm) GetWindowRect((HWND)p->hwndForm, &r);
            int32_t v = 0;
            if (_wcsicmp(nm, L"Width") == 0 || _wcsicmp(nm, L"ScaleWidth") == 0) v = vb6_XToTwipX(r.right - r.left);
            else if (_wcsicmp(nm, L"Height") == 0 || _wcsicmp(nm, L"ScaleHeight") == 0) v = vb6_YToTwipY(r.bottom - r.top);
            else if (_wcsicmp(nm, L"Left") == 0) v = vb6_XToTwipX(r.left);
            else v = vb6_YToTwipY(r.top);
            axSetI4(out, v);
            return S_OK;
        }
        /* 布尔属性 */
        if (_wcsicmp(nm, L"Visible") == 0 || _wcsicmp(nm, L"Enabled") == 0
            || _wcsicmp(nm, L"ClipControls") == 0 || _wcsicmp(nm, L"ControlBox") == 0
            || _wcsicmp(nm, L"MaxButton") == 0 || _wcsicmp(nm, L"MinButton") == 0
            || _wcsicmp(nm, L"Moveable") == 0 || _wcsicmp(nm, L"AutoRedraw") == 0) {
            axSetBool(out, 1);
            return S_OK;
        }
        if (_wcsicmp(nm, L"MDIChild") == 0 || _wcsicmp(nm, L"KeyPreview") == 0
            || _wcsicmp(nm, L"RightToLeft") == 0 || _wcsicmp(nm, L"Zoom") == 0) {
            axSetBool(out, 0);
            return S_OK;
        }
        /* 颜色 */
        if (_wcsicmp(nm, L"BackColor") == 0 || _wcsicmp(nm, L"ForeColor") == 0) {
            axSetI4(out, 0x8000000F); return S_OK;
        }
        if (_wcsicmp(nm, L"FillColor") == 0 || _wcsicmp(nm, L"MousePointer") == 0
            || _wcsicmp(nm, L"DrawStyle") == 0 || _wcsicmp(nm, L"FillStyle") == 0
            || _wcsicmp(nm, L"WindowState") == 0 || _wcsicmp(nm, L"StartUpPosition") == 0
            || _wcsicmp(nm, L"CurrentX") == 0 || _wcsicmp(nm, L"CurrentY") == 0
            || _wcsicmp(nm, L"ScaleLeft") == 0 || _wcsicmp(nm, L"ScaleTop") == 0
            || _wcsicmp(nm, L"hDC") == 0 || _wcsicmp(nm, L"TextWidth") == 0
            || _wcsicmp(nm, L"TextHeight") == 0) {
            axSetI4(out, 0); return S_OK;
        }
        /* 只读整型常量: ScaleMode=1 (Twips) / Appearance=1 / BorderStyle=2 / DrawWidth=1 */
        if (_wcsicmp(nm, L"ScaleMode") == 0)   { axSetI4(out, 1); return S_OK; }
        if (_wcsicmp(nm, L"Appearance") == 0)  { axSetI4(out, 1); return S_OK; }
        if (_wcsicmp(nm, L"BorderStyle") == 0) { axSetI4(out, 2); return S_OK; }
        if (_wcsicmp(nm, L"DrawWidth") == 0)   { axSetI4(out, 1); return S_OK; }
        /* 方法 (无返回值语义): 真做几个有用动作, 其余 no-op */
        if (_wcsicmp(nm, L"Refresh") == 0) {
            if (p->hwndForm) InvalidateRect((HWND)p->hwndForm, NULL, FALSE);
            axSetEmpty(out); return S_OK;
        }
        if (_wcsicmp(nm, L"SetFocus") == 0) {
            if (p->hwndForm) SetFocus((HWND)p->hwndForm);
            axSetEmpty(out); return S_OK;
        }
        if (_wcsicmp(nm, L"Show") == 0 || _wcsicmp(nm, L"Hide") == 0
            || _wcsicmp(nm, L"ZOrder") == 0 || _wcsicmp(nm, L"Move") == 0
            || _wcsicmp(nm, L"Cls") == 0 || _wcsicmp(nm, L"Print") == 0
            || _wcsicmp(nm, L"PopupMenu") == 0 || _wcsicmp(nm, L"Circle") == 0
            || _wcsicmp(nm, L"Line") == 0 || _wcsicmp(nm, L"PSet") == 0
            || _wcsicmp(nm, L"Point") == 0 || _wcsicmp(nm, L"Scale") == 0) {
            axSetEmpty(out); return S_OK;
        }
        /* Picture / MouseIcon / Parent 等对象属性: 无对象 */
        axSetEmpty(out);
        return S_OK;
    }

    /* CONTROLS 集合 */
    const wchar_t* nm = (idx >= 1 && idx <= axCount(kColNames)) ? kColNames[idx - 1] : NULL;
    if (!nm) return DISP_E_MEMBERNOTFOUND;
    if (_wcsicmp(nm, L"Count") == 0) {
        int32_t n = 0;
        int32_t total = vb6_HostObj_Count();
        for (int32_t i = 0; i < total; i++) {
            void* h = vb6_HostObj_At(i);
            if (h && h != p->hwndForm) n++;
        }
        axSetI4(out, n);
        return S_OK;
    }
    if (_wcsicmp(nm, L"Item") == 0) {
        /* VB6 语义: Controls(name[, tabIdx]) / Controls(index) — 实参逆序 */
        const wchar_t* wantName = NULL;
        int32_t wantIdx = -1;
        if (dp && dp->cArgs >= 1) {
            VARIANT* a0 = &dp->rgvarg[dp->cArgs - 1];      /* 第一个实参 */
            if (V_VT(a0) == VT_BSTR) wantName = V_BSTR(a0);
            else if (V_VT(a0) == VT_I4 || V_VT(a0) == VT_I2) wantIdx = (V_VT(a0) == VT_I4) ? V_I4(a0) : V_I2(a0);
        }
        if (wantName) {
            /* 按名字查 (支持 "Name(idx)") */
            void* h = vb6_UC_ControlsItemByName(NULL, wantName, -1);
            if (!h) { axSetEmpty(out); out->vt = vb6_vtEmpty; return S_OK; }
            axSetObj(out, h);   /* 宿主对象约定: hwnd 即 IDispatch* */
            return S_OK;
        }
        if (wantIdx >= 1) {
            /* 按序号 (1-based) 在窗体子控件里找第 wantIdx 个非窗体对象 */
            int32_t cur = 0, total = vb6_HostObj_Count();
            for (int32_t i = 0; i < total; i++) {
                void* h = vb6_HostObj_At(i);
                if (!h || h == p->hwndForm) continue;
                if (++cur == wantIdx) { axSetObj(out, h); return S_OK; }
            }
        }
        axSetEmpty(out);
        return S_OK;
    }
    /* _NewEnum: 交给 RTL 的宿主枚举 (控件一般不枚举容器控件, 返回空即可) */
    axSetEmpty(out);
    return S_OK;
}

static const IDispatchVtbl g_axObjVtbl = {
    ax_QI, ax_AddRef, ax_Release, ax_GetTypeInfoCount, ax_GetTypeInfo,
    ax_GetIDsOfNames, ax_Invoke
};
static const IDispatchVtbl g_axExtVtbl = {
    ax_QI, ax_AddRef, ax_Release, ax_GetTypeInfoCount, ax_GetTypeInfo,
    ax_GetIDsOfNames, ax_Invoke
};

/* --- 对外入口 --- */

/* 创建控件的 Extender (IOleControlSite::GetExtendedControl 的返回值) */
void* vb6_AxContainer_CreateExtender(void* hwndForm, const wchar_t* ctrlName) {
    return (void*)axNew(VB6_AXK_EXTENDER, hwndForm, ctrlName);
}
