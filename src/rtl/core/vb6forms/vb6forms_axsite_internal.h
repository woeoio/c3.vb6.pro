// vb6forms_axsite_internal.h - vb6forms_axsite 家族内部共享声明 (Fix 143~149 OCX 真宿主)
// 仅供 src/rtl/core/vb6forms/axsite/ 下各 .c 使用; 生成代码只 include vb6forms.h
//
// 拆分说明: vb6forms_axsite.c (1029 行) 按功能家族拆为 5 个编译单元, 各单元由 MSVC
// 独立编译成 .obj 再链接, 文件级 static 不跨文件可见 —— 因此把跨族共享的结构体
// **完整定义** 与符号声明集中在此, 定义仍留在各自的族文件里。
// 解包后是平铺目录: RTL 内部 include 一律写 basename, 不带子目录路径。

#ifndef VB6C3_VB6FORMS_AXSITE_INTERNAL_H
#define VB6C3_VB6FORMS_AXSITE_INTERNAL_H

// 平台与 COM 头：与拆分前 vb6forms_axsite.c 的文件头保持一致（勿删）。
// 本头的 Vb6AxSite / Vb6PropBag 直接使用 OLE 接口 vtable 类型
// （IOleClientSiteVtbl / IOleInPlaceSiteVtbl / IOleControlSiteVtbl /
//  IOleInPlaceSiteWindowlessVtbl / IDispatchVtbl / IPropertyBagVtbl），
// 这些类型由 <olectl.h> 提供 —— 只 include vb6forms.h 不够（它只带
// <windows.h>，而 WIN32_LEAN_AND_MEAN 下 windows.h 不含 OLE 头）。
// 缺了这几行的后果实测过：本头整段解析失败，下游 ax_host.c 级联报
// C2065 'bag' undeclared / C2223 直到 C1003 error count exceeds 100。
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
#include <stdlib.h>   /* calloc / free */
#include <stddef.h>   /* offsetof */
#include <oleauto.h>  /* SysAllocString, BSTR */
#include <olectl.h>   /* IPicture / OLE 接口 vtable */

// ===== 原 34~71 行: 站点对象结构 + 接口指针反推宏 =====

typedef struct Vb6AxSite Vb6AxSite;

struct Vb6AxSite {
    /* IOleClientSite */
    const IOleClientSiteVtbl* lpVtblClientSite;
    /* IOleInPlaceSite (通过 QI 返回同一个对象) */
    const IOleInPlaceSiteVtbl* lpVtblInPlaceSite;
    /* IOleInPlaceFrame (通过 QI 返回同一个对象) */
    const IOleInPlaceFrameVtbl* lpVtblInPlaceFrame;
    LONG ref;
    HWND hwndForm;        /* 宿主 Form 窗口句柄 */
    IOleObject* pOleObj;  /* 被托管的控件 (在SetClientSite后保存) */
    RECT rcCtrl;          /* Fix 143: 控件在窗体客户区的矩形 (GetWindowContext 用) */
    /* Fix 143d: 无窗口 (windowless) 控件宿主必需的三件套.
     * VB6 UserControl 激活时会向 site 索要这三个接口 (实测 NewTab01.ocx):
     *   IOleInPlaceSiteWindowless {40A050A0-3C31-101B-A82E-08002B2B2337}
     *   IOleControlSite          {B196B289-BAB4-101A-B69C-00AA00341D07}
     *   IDispatch (ambient 属性)
     * 缺任一个时 DoVerb 仍返回 S_OK, 但控件不进入激活态 → 窗体上什么都不画. */
    const IOleControlSiteVtbl* lpVtblControlSite;
    const IOleInPlaceSiteWindowlessVtbl* lpVtblInPlaceSiteWindowless;
    const IDispatchVtbl* lpVtblDispatch;
    int inPlaceActive;
    int captured;         /* IOleInPlaceSiteWindowless::SetCapture 状态 */
    int focused;          /* IOleInPlaceSiteWindowless::SetFocus 状态 */
    void* pInPlaceObj;    /* IOleInPlaceObjectWindowless* — 消息转发用 */
    void* pViewObj;       /* IViewObject* — 宿主 WM_PAINT 里绘制控件用 */
    wchar_t ctrlName[128];/* Fix 148: 控件实例名 (容器对象模型用) */
    void* pExtender;      /* Fix 148: Extender 缓存 (IOleControlSite::GetExtendedControl) */
};

/* Fix 143d: 从接口指针反推宿主对象.
 * COM 多重接口共用同一个对象, 但 QI 返回的是**对象内某个 vtable 字段的地址**,
 * 所以每个方法里必须减去该字段的 offset 才能拿到对象首地址
 * (原实现直接 (Vb6AxSite*)This 只在首字段 IOleClientSite 上恰好正确,
 *  其余接口 (IOleInPlaceSite/Frame 等) 全都偏移错位). */
#define SITE_OF(thisptr, field) ((Vb6AxSite*)((char*)(thisptr) - offsetof(Vb6AxSite, field)))


// ===== 原 281~286 行: 简易 IPropertyBag =====

typedef struct Vb6PropBag {
    const IPropertyBagVtbl* lpVtbl;
    LONG ref;
    const Vb6OcxProp* props;
    int count;
} Vb6PropBag;

// ===== 跨族共享符号 (原 static, 拆分后提升为外部链接) =====
// 站点 vtable: ClientSite/InPlaceSite/InPlaceFrame 定义在 ax_site.c,
// Dispatch/ControlSite/InPlaceSiteWindowless 定义在 ax_site_ext.c,
// 二者都由 ax_host.c 的 vb6_OcxHost_Create 装配到站点对象上。
extern const IOleClientSiteVtbl g_axSiteClientSiteVtbl;
extern const IOleInPlaceSiteVtbl g_axSiteInPlaceSiteVtbl;
extern const IOleInPlaceFrameVtbl g_axSiteInPlaceFrameVtbl;
extern const IDispatchVtbl g_axSiteDispatchVtbl;
extern const IOleControlSiteVtbl g_axSiteControlSiteVtbl;
extern const IOleInPlaceSiteWindowlessVtbl g_axSiteInPlaceSiteWindowlessVtbl;

// 站点 IUnknown 实现 (定义在 ax_site.c): 各接口的 QI/AddRef/Release 都转发到这里
HRESULT STDMETHODCALLTYPE axSite_CS_QueryInterface(IOleClientSite* This, REFIID riid, void** ppv);
ULONG   STDMETHODCALLTYPE axSite_CS_AddRef(IOleClientSite* This);
ULONG   STDMETHODCALLTYPE axSite_CS_Release(IOleClientSite* This);

// 已创建宿主登记表 (定义在 ax_site.c; 消息转发/绘制按窗体遍历它)
// 定义处仍用 VB6_MAX_AXSITES 定长数组, 故此处不作尺寸声明。
extern Vb6AxSite* g_axSites[];
extern int vb6_axSiteCount;
void axSiteRegister(Vb6AxSite* s);

// OCX 实例化 + 单位换算 (定义在 ax_load.c)
HRESULT vb6_ocxCreateAny(const wchar_t* ocxPath, REFCLSID rclsid, void** ppUnk);
long vb6_twipsToHimetric(long twips);

// 简易 IPropertyBag vtable (定义在 ax_propbag.c, 由 ax_host.c 装配)
extern const IPropertyBagVtbl g_pbVtbl;

#endif // VB6C3_VB6FORMS_AXSITE_INTERNAL_H
