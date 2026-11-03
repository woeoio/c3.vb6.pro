// vb6forms_ctrlarr.c - vb6forms 模块拆分: 控件数组 + MDI 窗体
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

// // 控件数组 (P7.6)
// ============================================================

void vb6_CtrlArr_Init(vb6_CtrlArr* arr) {
    int i;
    for (i = 0; i < VB6_CTRLARR_MAX; i++) {
        arr->hwnds[i] = NULL;
    }
    arr->count = 0;
    arr->lowerBound = 0;
    arr->upperBound = -1;
}

void vb6_CtrlArr_SetAt(vb6_CtrlArr* arr, int index, void* hwnd) {
    if (index < 0 || index >= VB6_CTRLARR_MAX) return;
    if (arr->hwnds[index] == NULL && hwnd != NULL) {
        arr->count++;
    } else if (arr->hwnds[index] != NULL && hwnd == NULL) {
        arr->count--;
    }
    arr->hwnds[index] = hwnd;
    /* 更新界 */
    if (hwnd != NULL) {
        if (arr->upperBound < 0 || index > arr->upperBound) arr->upperBound = index;
        if (index < arr->lowerBound) arr->lowerBound = index;
    }
}

void* vb6_CtrlArr_GetAt(const vb6_CtrlArr* arr, int index) {
    if (index < 0 || index >= VB6_CTRLARR_MAX) return NULL;
    return arr->hwnds[index];
}

int vb6_CtrlArr_GetCount(const vb6_CtrlArr* arr) {
    return arr->count;
}

int vb6_CtrlArr_LBound(const vb6_CtrlArr* arr) {
    return arr->lowerBound;
}

int vb6_CtrlArr_UBound(const vb6_CtrlArr* arr) {
    return arr->upperBound;
}

void* vb6_CtrlArr_Load(vb6_CtrlArr* arr, int index, void* hParent, void* hInstance) {
    HWND hNew, hTemplate;
    WCHAR className[256] = {0};
    WCHAR text[1024] = {0};
    RECT rc;
    DWORD style, exStyle;
    int ctrlId;

    if (index < 0 || index >= VB6_CTRLARR_MAX) return NULL;
    if (arr->hwnds[index] != NULL) return arr->hwnds[index];  /* 已存在 */

    /* 找模板: 优先index=0, 否则第一个非空 */
    hTemplate = (HWND)vb6_CtrlArr_GetAt(arr, 0);
    if (!hTemplate) {
        int i;
        for (i = 0; i < VB6_CTRLARR_MAX; i++) {
            if (arr->hwnds[i]) { hTemplate = (HWND)arr->hwnds[i]; break; }
        }
    }
    if (!hTemplate) return NULL;

    /* 从模板复制窗口属性 */
    GetClassNameW(hTemplate, className, 256);
    GetWindowTextW(hTemplate, text, 1024);
    GetWindowRect(hTemplate, &rc);
    style = (DWORD)GetWindowLongPtrA(hTemplate, GWL_STYLE);
    exStyle = (DWORD)GetWindowLongPtrA(hTemplate, GWL_EXSTYLE);

    ctrlId = vb6_NextControlId();

    hNew = CreateWindowExW(
        exStyle, className, text, style,
        rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
        (HWND)hParent, (HMENU)(intptr_t)ctrlId, (HINSTANCE)hInstance, NULL
    );

    if (hNew) {
        vb6_GeomOwnerStamp(hNew, vb6_GeomAuthoredOwner(hParent));   /* 账 #305 299: 与 vb6_CreateControl 同一口径 */

        // 账 #204: 模板的字体也从那一处出口问 —— 裸 `WM_GETFONT` 对 STATIC 那一类模板恒回 NULL，
        // 外面那句 `if (hFont) SendMessage(...)` 就会整步跳过，新窗口连一次 WM_SETFONT 都没收到。
        // 发给新窗口之后同样要存进槽位（新窗口自己也是 STATIC 那一类的话，下一读又是 NULL）。
        // ⚠ 这一条**运行路今天不可达**：`vb6_CtrlArr_Load` 全仓零调用者（发码侧还没接 `Load <数组>(n)`
        // 那一形，2026-10-05 grep 证），所以这里是**口径统一**、不是产品修复，也别拿它写判据。
        HFONT hFont = vb6_ControlFont(hTemplate);
        if (hFont) {
            SendMessage(hNew, WM_SETFONT, (WPARAM)hFont, MAKELPARAM(FALSE, 0));
            vb6_ControlFontStore(hNew, hFont);
        }
        vb6_CtrlArr_SetAt(arr, index, (void*)hNew);
    }
    return (void*)hNew;
}

void vb6_CtrlArr_Unload(vb6_CtrlArr* arr, int index) {
    if (index < 0 || index >= VB6_CTRLARR_MAX) return;
    if (arr->hwnds[index] == NULL) return;
    /* 不能卸载设计时创建的元素(index=0或其他初始元素) — VB6也是如此 */
    DestroyWindow((HWND)arr->hwnds[index]);
    vb6_CtrlArr_SetAt(arr, index, NULL);
}

// ============================================================
// MDI窗体 (P7.7)
// ============================================================

// MDI父窗体的客户窗口句柄, 存入GWLP_USERDATA
// 用Prop也可以, 但USERDATA更快

// 自动查找MDIClient窗口的回调
static BOOL CALLBACK FindMDIClientEnumProc(HWND hwnd, LPARAM lParam) {
    HWND hMDIClient = (HWND)GetPropW(hwnd, L"VB6_MDIClient");
    if (hMDIClient) {
        *(HWND*)lParam = hMDIClient;
        return FALSE;  /* 找到, 停止枚举 */
    }
    return TRUE;  /* 继续枚举 */
}

static HWND vb6_AutoFindMDIClient(void) {
    HWND hMDIClient = NULL;
    EnumWindows(FindMDIClientEnumProc, (LPARAM)&hMDIClient);
    return hMDIClient;
}
int vb6_RegisterMDIFormClass(const char* className, void* wndProc, void* hInstance, int iconResId) {
    // Fix 190: 与主窗体一致 —— 必须用 W 版注册, 否则 MDI 窗体是 ANSI 窗口,
    // 标题/子窗体标题里的非 ASCII 会被按 ACP 转换 (中韩俄文乱码)。
    wchar_t wcls[128];
    vb6_u8ToWideBuf(className, wcls, 128);
    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = (WNDPROC)wndProc;
    wc.hInstance = (HINSTANCE)hInstance;
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_APPWORKSPACE + 1);
    wc.lpszClassName = wcls;
    if (iconResId > 0) {
        wc.hIcon = LoadIconW((HINSTANCE)hInstance, MAKEINTRESOURCEW(iconResId));
    } else {
        wc.hIcon = LoadIconW(NULL, IDI_APPLICATION);
    }
    ATOM atom = RegisterClassExW(&wc);
    return (atom != 0) ? 0 : -1;
}

void* vb6_CreateMDIFormWindow(const char* className, const char* formName,
    int x, int y, int width, int height, void* hInstance) {
    int pw = vb6_TwipToX(width);
    int ph = vb6_TwipToY(height);
    DWORD mdiStyle = WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN;
    // Adjust for non-client area to get actual window size
    RECT rcM = {0, 0, pw, ph};
    AdjustWindowRectEx(&rcM, mdiStyle, FALSE, 0);
    int winW = rcM.right - rcM.left;
    int winH = rcM.bottom - rcM.top;
    int px, py;
    if (x == -1 && y == -1) {
        // M22-Issue2: CenterScreen
        px = (GetSystemMetrics(SM_CXSCREEN) - winW) / 2;
        py = (GetSystemMetrics(SM_CYSCREEN) - winH) / 2;
    } else {
        px = x;
        py = y;
    }
    // Fix 190: MDI 框架全程 W 版 (类名/标题是 RTL 内部 UTF-8, 此处转宽)
    wchar_t wclsM[128];
    vb6_u8ToWideBuf(className, wclsM, 128);
    wchar_t* wtitleM = vb6_u8ToWideDup(formName);

    HWND hwnd = CreateWindowExW(
        0, wclsM, wtitleM ? wtitleM : L"",
        mdiStyle,
        px, py,
        winW, winH,
        NULL, NULL, (HINSTANCE)hInstance, NULL);
    free(wtitleM);
    if (!hwnd) return NULL;

    /* 创建MDI客户窗口 */
    CLIENTCREATESTRUCT ccs = {0};
    ccs.hWindowMenu = NULL;   /* P7.8菜单系统实现后填充 */
    ccs.idFirstChild = 1000;  /* MDI子窗体ID起始值 */

    HWND hMDIClient = CreateWindowExW(
        0, L"MDICLIENT", NULL,
        WS_CHILD | WS_CLIPCHILDREN | WS_VSCROLL | WS_HSCROLL | MDIS_ALLCHILDSTYLES,
        0, 0, 0, 0,
        hwnd, (HMENU)0xCAC,   /* MDI客户窗口控件ID */
        (HINSTANCE)hInstance, &ccs);
    if (!hMDIClient) {
        DestroyWindow(hwnd);
        return NULL;
    }
    ShowWindow(hMDIClient, SW_SHOW);

    /* 保存MDI客户窗口句柄到父窗体属性 */
    SetPropW(hwnd, L"VB6_MDIClient", hMDIClient);

    return (void*)hwnd;
}

void* vb6_CreateMDIChildWindow(const char* className, const char* formName,
    int x, int y, int width, int height, void* hMDIClient, void* hInstance) {
    /* P7.7: 自动查找MDIClient (当hMDIClient=NULL时枚举窗口查找) */
    if (!hMDIClient) hMDIClient = vb6_AutoFindMDIClient();
    if (!hMDIClient) return NULL;  /* 没有MDI父窗体 */
    int pw = vb6_TwipToX(width);
    int ph = vb6_TwipToY(height);
    DWORD childStyle = WS_CHILD | WS_CLIPCHILDREN | WS_SYSMENU | WS_CAPTION | WS_THICKFRAME;
    RECT rcC = {0, 0, pw, ph};
    AdjustWindowRectEx(&rcC, childStyle, FALSE, 0);
    int winW = rcC.right - rcC.left;
    int winH = rcC.bottom - rcC.top;
    int px, py;
    if (x == -1 && y == -1) {
        // M22-Issue2: CenterScreen
        px = (GetSystemMetrics(SM_CXSCREEN) - winW) / 2;
        py = (GetSystemMetrics(SM_CYSCREEN) - winH) / 2;
    } else {
        px = (x == CW_USEDEFAULT) ? CW_USEDEFAULT : vb6_TwipToX(x);
        py = (y == CW_USEDEFAULT) ? CW_USEDEFAULT : vb6_TwipToY(y);
    }
    // Fix 190: MDICREATESTRUCTW —— 类名/标题必须宽字符, 否则子窗体标题乱码
    wchar_t* wclsC = vb6_u8ToWideDup(className);
    wchar_t* wtitleC = vb6_u8ToWideDup(formName);
    MDICREATESTRUCTW mcs = {0};
    mcs.szClass = wclsC;
    mcs.szTitle = wtitleC;
    mcs.x = px;
    mcs.y = py;
    mcs.cx = winW;
    mcs.cy = winH;
    /* 账 (MDI 布局刀): 子窗样式必须把**正常 MDI 子窗那一套**递进去, 不能留 0。
     * MDICLIENT 建的时候带了 MDIS_ALLCHILDSTYLES, 于是 mcs.style=0 被原样采纳 ——
     * 实测子窗样式只有 0x54000000 (CHILD|VISIBLE|CLIPSIBLINGS), 没有 CAPTION/边框/
     * MAXIMIZEBOX。后果两条: (a) AdjustWindowRectEx 按带框样式加出来的非客户区与
     * 真实窗口不符, 设计值 583x417 变成 603x465; (b) WM_MDIMAXIMIZE 对无框子窗
     * **完全不做事** (夹具 mdimax 实测: 发了消息仍 604x465, WS_MAXIMIZE 位不置)。
     * 补齐样式后 WM_MDIMAXIMIZE 才铺满 MDIClient, 且父窗后续 WM_SIZE 会自动跟。 */
    mcs.style = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
             | WS_SYSMENU | WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX;
    mcs.lParam = 0;

    HWND hChild = (HWND)SendMessageW((HWND)hMDIClient, WM_MDICREATE, 0, (LPARAM)&mcs);
    free(wclsC);
    free(wtitleC);
    return (void*)hChild;
}

/* 账 #305 第三百四十六刀: MDI 客体带重排 (Codejock 停靠条共存)。
 * CommandBars 控件排好了自己的菜单/工具栏带 (XTPDockBar 顶带) 和状态栏
 * (XTPStatusBar 底带), 但 **MDIClient 没被挪进空余带** —— 它从窗体顶铺到底,
 * 活动子窗体盖住菜单 (GUISample 实测: frmDocument @y=94 压住顶带)。这里在
 * 窗体 WM_SIZE 时把 MDIClient 收进 [顶带底, 底带顶] 的空余区, 并压到 z 序底
 * (停靠条必须留在它上面)。只接管有 XTPDockBar 的窗体, 原生 MDI 工程零影响。 */
typedef struct VB6_MdiBandCtx { HWND form; int top, bottom; BOOL xtp; } VB6_MdiBandCtx;

static BOOL CALLBACK vb6_mdi_band_cb(HWND h, LPARAM lp) {
    VB6_MdiBandCtx* c = (VB6_MdiBandCtx*)lp;
    wchar_t cls[64] = {0};
    GetClassNameW(h, cls, 63);
    if (GetParent(h) != c->form) return TRUE;
    if (wcsncmp(cls, L"XTPDockBar", 10) != 0 && wcsncmp(cls, L"XTPStatusBar", 12) != 0)
        return TRUE;
    c->xtp = TRUE;
    RECT rc; GetWindowRect(h, &rc);
    MapWindowPoints(NULL, c->form, (POINT*)&rc, 2);
    RECT rf; GetClientRect(c->form, &rf);
    int bw = rc.right - rc.left, bh = rc.bottom - rc.top;
    /* 账 (VC6 主题 MDIClient 不内缩): 只有**横跨半个客户区以上、厚度非零**的条才算
     * 顶/底带。VC6 主题下左右两枚空 XTPDockBar 是 **0 宽、从客户区顶铺到底**
     * (实测 rect 0x881@4,27 / 0x881@1532,27); 旧判据只问 rc.top<=4, 于是它们被
     * 当成顶带、把 top 取成客户区全高 881 ⇒ 下面按 h<=0 提前 return, 内缩与 z 序
     * 下压整段被跳过 (C3_MDI_TRACE: "top=881 seen=16 left=210 right=210 bottom=141")。
     * VS2015 主题下这两枚空条落在 y=104 才侥幸躲过。 */
    if (bw <= 4 || bh <= 0) return TRUE;
    if (rc.top <= 4 && bw >= rf.right / 2) {          /* 顶带: 取最深的底 */
        if (rc.bottom > c->top) c->top = rc.bottom;
    } else if (rc.bottom >= rf.bottom - 4 && bw >= rf.right / 2) {  /* 底带: 取最高 */
        if (bh > c->bottom) c->bottom = bh;
    }
    return TRUE;
}

struct Ctx352 { HWND paneMgr; };
struct Ctx367 { HWND form, client; };

static BOOL CALLBACK vb6_mdi_zorder_cb367(HWND h, LPARAM lp) {
    struct Ctx367* c = (struct Ctx367*)lp;
    wchar_t cls[64] = {0};
    GetClassNameW(h, cls, 63);
    if ((wcsncmp(cls, L"XTPDockBar", 10) == 0 || wcsncmp(cls, L"XTPStatusBar", 12) == 0
        || wcsncmp(cls, L"XTPToolBar", 10) == 0) && GetParent(h) == c->form) {
        SetWindowPos(h, HWND_TOP, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }
    return TRUE;
}

static BOOL CALLBACK vb6_mdi_panemgr_cb352(HWND h, LPARAM lp) {
    wchar_t cls[64] = {0};
    GetClassNameW(h, cls, 63);
    if (wcsncmp(cls, L"XTPDockingPaneManager", 22) == 0) {
        ((struct Ctx352*)lp)->paneMgr = h;
        return FALSE;
    }
    return TRUE;
}

/* 账 (MDI 停靠内缩刀): 左右/底部停靠窗格的实际内缩。
 * 第三百四十六刀只处理**窗体直接子窗**的顶部/底部带 (XTPDockBar / XTPStatusBar),
 * 于是左右停靠的 XTPDockingPaneTabbedContainer (ToolBox / Class View / Output)
 * 仍压在 MDIClient 上面, 文档窗被裁掉 210px。正规通路
 * `CommandBars.ShowTabWorkspace(True)` 由 OCX 自己回 S_OK + **空指针**
 * (实测 pdisp=0 ⇒ Workspace 链三处 disp=0), `Options.ClientMargin` 只管
 * 窗格与客户区边框之间 —— 两条都不内缩 MDIClient, 所以这里按窗格**实测 rect**
 * 算内缩。这是**布局兜底**, 不是完整修复: 正规通路缺的是 OCX 把
 * MDIClient 交给 DockingPaneManager 的那一步 (需要它认识非 MFC 的 frame)。 */
struct VB6_MdiDockCtx { HWND form; int left, right, bottom; };

static BOOL CALLBACK vb6_mdi_dockpane_cb(HWND h, LPARAM lp) {
    struct VB6_MdiDockCtx* c = (struct VB6_MdiDockCtx*)lp;
    wchar_t cls[64] = {0};
    GetClassNameW(h, cls, 63);
    if (GetParent(h) != c->form) return TRUE;
    /* 只认停靠容器; 浮动窗格是弹出窗 (无 WS_CHILD 且不是窗体子窗), 不参与内缩 */
    if (wcsncmp(cls, L"XTPDockingPaneTabbedContainer", 29) != 0) return TRUE;
    if (!IsWindowVisible(h)) return TRUE;
    RECT rc; GetWindowRect(h, &rc);
    MapWindowPoints(NULL, c->form, (POINT*)&rc, 2);
    RECT rcl; GetClientRect(c->form, &rcl);
    if (rc.right - rc.left <= 0 || rc.bottom - rc.top <= 0) return TRUE;
    int cw = rc.right - rc.left;
    int leftDock = (rc.left <= 4 && rc.right < rcl.right / 2);
    int rightDock = (rc.right >= rcl.right - 4 && rc.left > rcl.right / 2);
    if (leftDock && rc.right > c->left) c->left = rc.right;
    if (rightDock) {
        int inset = rcl.right - rc.left;
        if (inset > c->right) c->right = inset;
    }
    /* 底停靠: 必须**不是**左右停靠, 且横跨过半 (否则右侧窄栏的下半格会被
     * 误当成底带 —— 实测 PropertyGrid 容器 210x393@1329,501 就这么把
     * MDIClient 从 670 高压成 397)。 */
    if (!leftDock && !rightDock && cw >= rcl.right / 2 && rc.top > rcl.bottom / 2) {
        int inset = rcl.bottom - rc.top;
        if (inset > c->bottom) c->bottom = inset;
    }
    return TRUE;
}

void vb6_MDI_RelayoutClient(void* hMDIForm) {
    HWND hForm = (HWND)hMDIForm;
    if (!hForm || !IsWindow(hForm)) return;
    HWND hClient = (HWND)GetPropW(hForm, L"VB6_MDIClient");
    if (!hClient || !IsWindow(hClient)) return;
    VB6_MdiBandCtx ctx;
    ctx.form = hForm; ctx.top = 0; ctx.bottom = 0; ctx.xtp = FALSE;
    EnumChildWindows(hForm, vb6_mdi_band_cb, (LPARAM)&ctx);
    if (!ctx.xtp) return;
    RECT rc; GetClientRect(hForm, &rc);
    /* 账 (MDI 停靠内缩刀): 再按左右/底部**停靠窗格实测 rect** 内缩。停靠容器
     * (XTPDockingPaneTabbedContainer) 实测是**窗体直接子窗** (父 = MDIForm, 不在
     * pane manager 之下), 所以上面那张带表只按类名前缀过滤、看不到它们。
     * 无 pane manager / 无停靠窗格时三个内缩全 0, 与旧行为一致 (原生 MDI 工程零影响)。 */
    struct VB6_MdiDockCtx dctx = { hForm, 0, 0, 0 };
    EnumChildWindows(hForm, vb6_mdi_dockpane_cb, (LPARAM)&dctx);
    struct Ctx352 pmc = { NULL };
    EnumChildWindows(hForm, vb6_mdi_panemgr_cb352, (LPARAM)&pmc);
    if (dctx.bottom > ctx.bottom) ctx.bottom = dctx.bottom;   /* 底带取高者 */
    int leftInset = dctx.left, rightInset = dctx.right;
    int w = rc.right - leftInset - rightInset;
    int h = rc.bottom - ctx.top - ctx.bottom;
    if (w < 40) {                       /* 兜底: 绝不让内缩把客户区压没 */
        leftInset = 0; rightInset = 0; w = rc.right;
    }
    /* 带表自相矛盾 (top+bottom >= 客户区高) 时按全高摆, 但**不要 return** ——
     * 后面的 z 序下压必须照做, 否则 MDIClient 会压在整个 CommandBars 带上
     * (VC6 主题实测: 菜单/工具栏/状态栏全被它盖住)。 */
    if (h <= 0) {
        ctx.top = 0; ctx.bottom = 0; h = rc.bottom;
        if (h <= 0) return;
    }
    /* 账 (MDI 布局刀): 带内摆位与"尺寸变化"解耦。
     * 第三百四十六刀只把 MoveWindow 放在 sizeChanged353 里, 而 Codejock 的停靠条是
     * 窗体 **Show 之后**才由 OCX 建出来的 —— 那时第一次 (也是唯一一次) WM_SIZE 早已过去,
     * 于是 ctx.xtp 已真、sizeChanged353 恒假, MDIClient 永远停在 (0,0) 全客户区,
     * 压住菜单/工具栏 (实测 GUISample: Menu Bar rect y=23..50, MDIClient 同 y=23..915)。
     * 这里改成每次调用都对表: 位置不对就摆一次 (单发 MoveWindow, 很便宜); 重的
     * Codejock 回流 (FormResized / pane manager WM_SIZE / RecalcLayout) 仍只在尺寸
     * 真变时做一次 —— 逐消息做那套才会与 Codejock 打乒乓 (第三百五十三刀的实测)。 */
    {
        RECT rcC; GetWindowRect(hClient, &rcC);
        MapWindowPoints(NULL, hForm, (POINT*)&rcC, 2);
        if (rcC.left != leftInset || rcC.top != ctx.top
            || (rcC.right - rcC.left) != w || (rcC.bottom - rcC.top) != h) {
            MoveWindow(hClient, leftInset, ctx.top, w, h, TRUE);
        }
    }
    /* 第三十四至三百五十三刀: Codejock 的排布/回流只在**窗体尺寸真的变化**时做
     * 一套 (FormResized 的 SetObjectRects + DockingPaneManager 的 WM_SIZE +
     * CommandBars.RecalcLayout)。按消息逐发会打乱它的布局状态 (实测: 窗口被
     * 隐藏回小尺寸)。静态表记每个 MDI 窗体上一次的尺寸。 */
    {
        static struct { HWND h; int w, h_; } s_lastSize[8] = {0};
        static int s_nLast = 0;
        int i353 = -1;
        for (int k = 0; k < s_nLast; k++)
            if (s_lastSize[k].h == hForm) { i353 = k; break; }
        if (i353 < 0 && s_nLast < 8) {
            i353 = s_nLast++;
            s_lastSize[i353].h = hForm;
            s_lastSize[i353].w = -1; s_lastSize[i353].h_ = -1;
        }
        int sizeChanged353 = (i353 < 0)
            ? 1
            : (s_lastSize[i353].w != rc.right || s_lastSize[i353].h_ != rc.bottom);
        if (i353 >= 0) {
            s_lastSize[i353].w = rc.right;
            s_lastSize[i353].h_ = rc.bottom;
        }
        if (sizeChanged353) {
            /* MDIClient 收进带内 —— 只在尺寸真变时做一次。逐消息做会与
             * Codejock 自己的 MDIClient 重排打乒乓 (实测应用假死)。 */
            MoveWindow(hClient, leftInset, ctx.top, w, h, TRUE);
            vb6_OcxHost_FormResized(hForm, rc.right, rc.bottom);
            if (pmc.paneMgr) {
                /* 物理把控件窗口改成客户区大小再发 WM_SIZE —— Codejock 的
                 * 窗格回流由自己的 WM_SIZE 驱动, 而 0x0 隐窗没人给它改尺寸。 */
                SetWindowPos(pmc.paneMgr, NULL, 0, 0, rc.right, rc.bottom,
                             SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(pmc.paneMgr, WM_SIZE, SIZE_RESTORED,
                             MAKELPARAM(rc.right, rc.bottom));
            }
            vb6_OcxHost_InvokeAll(hForm, L"RecalcLayout");
            /* 第三百六十七刀: z 序 —— 停靠带/工具栏/状态栏 (窗体子窗) 提到
             * MDIClient 之上。此前写反 hInsertAfter 语义 (插其后 = 压其下)。
             * 停靠带窗口其实已是全宽正确位置, 只是整个被 MDIClient 压住
             * (实测: 枚举 1542x81 @0,23 全宽正确, 屏上却只见文档)。 */
            {
                struct Ctx367 { HWND form, client; } c367 = { hForm, hClient };
                EnumChildWindows(hForm, vb6_mdi_zorder_cb367, (LPARAM)&c367);
            }
        /* 第三百五十五刀: 真正的回流驱动 —— 无窗口 OCX (Codejock 一族) 的窗口
         * 消息靠容器**转发** (IOleInPlaceObjectWindowless::OnWindowMessage),
         * 而转发站 (vb6_OcxHost_ForwardMessage) 此前没有任何调用者 ⇒ 它们永远
         * 收不到 WM_SIZE, 停靠带/窗格全冻结在激活那一刻的尺寸。这里补发。 */
        /* 第三百六十八刀: ForwardMessage(WM_SIZE) 撤下 —— 它让 Codejock 重排
         * 时把 MDIClient 拉回 (0,0) 全客户区, 与带内摆位打乒乓 (实测: 摆到
         * y=81 又被拉回 y=0, 菜单带永远被盖)。 */
            /* nudge 链是同步的 —— Codejock 重排完可能把 MDIClient 挪回顶部
             * 盖住菜单带 (实测)。链尾重申一次带内位置与 z 序。 */
            RECT rcNow; GetWindowRect(hClient, &rcNow);
            MapWindowPoints(NULL, hForm, (POINT*)&rcNow, 2);
            if (rcNow.left != leftInset || rcNow.top != ctx.top
                || (rcNow.right - rcNow.left) != w || (rcNow.bottom - rcNow.top) != h)
                MoveWindow(hClient, leftInset, ctx.top, w, h, TRUE);
            SetWindowPos(hClient, HWND_BOTTOM, 0, 0, 0, 0,
                         SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        }
    }
    HWND hAbove = GetWindow(hClient, GW_HWNDPREV);
    wchar_t clsAbove[32] = {0};
    if (hAbove) {
        GetClassNameW(hAbove, clsAbove, 31);
        if (wcsncmp(clsAbove, L"XTPDockBar", 10) != 0
            && wcsncmp(clsAbove, L"XTPStatusBar", 12) != 0)
            SetWindowPos(hClient, HWND_BOTTOM, 0, 0, 0, 0,
                         SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }
}

void* vb6_GetMDIClient(void* hMDIForm) {
    if (!hMDIForm) return NULL;
    return (void*)GetPropW((HWND)hMDIForm, L"VB6_MDIClient");
}

/* 账 (MDI 布局刀): 子窗在父窗体上屏之前设的 WindowState 记在
 * VB6_PendingWindowState (值 = 状态+1, 0 表示无) 上, 父窗体 Show 之后在这里落地。 */
static BOOL CALLBACK vb6_mdi_pendingws_cb(HWND h, LPARAM lp) {
    (void)lp;
    HANDLE v = GetPropW(h, L"VB6_PendingWindowState");
    if (!v) return TRUE;
    RemovePropW(h, L"VB6_PendingWindowState");
    HWND cl = GetParent(h);
    wchar_t cls[32] = {0};
    if (cl) GetClassNameW(cl, cls, 31);
    if (!cl || _wcsicmp(cls, L"MDICLIENT") != 0) return TRUE;
    switch ((int)(intptr_t)v - 1) {
        case 2: SendMessageW(cl, WM_MDIMAXIMIZE, (WPARAM)h, 0); break;
        case 1: ShowWindow(h, SW_MINIMIZE); break;
        default: SendMessageW(cl, WM_MDIRESTORE, (WPARAM)h, 0); break;
    }
    return TRUE;
}

void vb6_MDIApplyPendingChildStates(void* hMDIClient) {
    if (!hMDIClient || !IsWindow((HWND)hMDIClient)) return;
    EnumChildWindows((HWND)hMDIClient, vb6_mdi_pendingws_cb, 0);
}

/* 账 (frmDocument 最大化): 见 vb6forms_mdi.h 的说明。标记只由 vb6com 的
 * ShowTabWorkspace 出口打, 原生 (无 TabWorkspace) 的 MDI 工程不会碰它。 */
static int s_tabWorkspaceMissing = 0;
void vb6_MDI_NoteTabWorkspaceMissing(void) { s_tabWorkspaceMissing = 1; }
int  vb6_MDI_TabWorkspaceMissing(void) { return s_tabWorkspaceMissing; }

int vb6_MDIMessageLoop(void* hAccelTable) {
    MSG msg;
    HACCEL hAccel = (HACCEL)hAccelTable;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        /* MDI加速键处理 */
        if (hAccel && TranslateAcceleratorW(msg.hwnd, hAccel, &msg)) {
            continue;
        }
        if (!TranslateMDISysAccel(vb6_GetMDIClient(GetParent(msg.hwnd)), &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    return (int)msg.wParam;
}

void vb6_MDITile(void* hMDIClient, int style) {
    SendMessageW((HWND)hMDIClient, WM_MDITILE, (WPARAM)(style ? MDITILE_VERTICAL : 0), 0);
}

void vb6_MDICascade(void* hMDIClient) {
    SendMessageW((HWND)hMDIClient, WM_MDICASCADE, 0, 0);
}

void vb6_MDIArrangeIcons(void* hMDIClient) {
    SendMessageW((HWND)hMDIClient, WM_MDIICONARRANGE, 0, 0);
}

void* vb6_MDIGetActive(void* hMDIClient) {
    return (void*)SendMessageW((HWND)hMDIClient, WM_MDIGETACTIVE, 0, 0);
}

void vb6_MDIActivate(void* hMDIClient, void* hChild) {
    SendMessageW((HWND)hMDIClient, WM_MDIACTIVATE, (WPARAM)(HWND)hChild, 0);
}
