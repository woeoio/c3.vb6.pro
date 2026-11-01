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
    mcs.style = 0;
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
    if (rc.top <= 4) {                        /* 顶带: 取最深的底 */
        if (rc.bottom > c->top) c->top = rc.bottom;
    } else if (rc.bottom >= rf.bottom - 4) {  /* 底带: 取最高 */
        int bh = rc.bottom - rc.top;
        if (bh > c->bottom) c->bottom = bh;
    }
    return TRUE;
}

struct Ctx352 { HWND paneMgr; };
static BOOL CALLBACK vb6_mdi_panemgr_cb352(HWND h, LPARAM lp) {
    wchar_t cls[64] = {0};
    GetClassNameW(h, cls, 63);
    if (wcsncmp(cls, L"XTPDockingPaneManager", 22) == 0) {
        ((struct Ctx352*)lp)->paneMgr = h;
        return FALSE;
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
    int w = rc.right;
    int h = rc.bottom - ctx.top - ctx.bottom;
    if (w <= 0 || h <= 0) return;
    /* 矩形没变就不动 (本钩每条消息都跑; Codejock 自己重排 MDIClient 后,
     * 下一条消息立即纠正)。z 序同理: 只有 MDIClient 真在停靠条之上才压。 */
    RECT rcCur; GetWindowRect(hClient, &rcCur);
    MapWindowPoints(NULL, hForm, (POINT*)&rcCur, 2);
    if (rcCur.left != 0 || rcCur.top != ctx.top
        || (rcCur.right - rcCur.left) != w || (rcCur.bottom - rcCur.top) != h)
        MoveWindow(hClient, 0, ctx.top, w, h, TRUE);
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
            vb6_OcxHost_FormResized(hForm, rc.right, rc.bottom);
            struct Ctx352 c352 = { NULL };
            EnumChildWindows(hForm, vb6_mdi_panemgr_cb352, (LPARAM)&c352);
            if (c352.paneMgr) {
                /* 物理把控件窗口改成客户区大小再发 WM_SIZE —— Codejock 的
                 * 窗格回流由自己的 WM_SIZE 驱动, 而 0x0 隐窗没人给它改尺寸。 */
                SetWindowPos(c352.paneMgr, NULL, 0, 0, rc.right, rc.bottom,
                             SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(c352.paneMgr, WM_SIZE, SIZE_RESTORED,
                             MAKELPARAM(rc.right, rc.bottom));
            }
            vb6_OcxHost_InvokeAll(hForm, L"RecalcLayout");
        /* 第三百五十五刀: 真正的回流驱动 —— 无窗口 OCX (Codejock 一族) 的窗口
         * 消息靠容器**转发** (IOleInPlaceObjectWindowless::OnWindowMessage),
         * 而转发站 (vb6_OcxHost_ForwardMessage) 此前没有任何调用者 ⇒ 它们永远
         * 收不到 WM_SIZE, 停靠带/窗格全冻结在激活那一刻的尺寸。这里补发。 */
        vb6_OcxHost_ForwardMessage(hForm, WM_SIZE, SIZE_RESTORED,
            MAKELPARAM(rc.right, rc.bottom), NULL);
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
