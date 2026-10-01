// VB6 Win32窓体运行时实现 (P7)
// 提供Win32窗口注册、创建、消息循环、控件管理等基础功能

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <psapi.h>    /* K32GetModuleInformation (崩溃栈扫描) */
#include <shellapi.h>
// C29-T: timeSetEvent / TIME_PERIODIC 的声明处。这里踩过一次坑：常量手抄成 0x02（真值是 1）
// 会让 timeSetEvent 直接失败并静默退回 SetTimer 那一档 15.6 ms 地板 => 用 SDK 头，不自己定义。
#include <mmsystem.h>

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
// Fix 190: 源码字符串 = UTF-8, 窗口层 = UTF-16
// 转码助手 (vb6_u8ToWideDup / vb6_u8ToWideBuf / vb6_wideToU8Buf) 定义在
// vb6forms_internal.h —— 它们要跨 vb6forms 的多个拆分编译单元共用。
// ============================================================

/* 账 #163/#168 的自研跳格：实现写在 `vb6_Form_SetInitialFocus` 之前那一段，而两条泵在文件里
   更靠前 ⇒ 先给个前向声明（C 的用法，不另开 .h —— 这条只给泵用，不进代码生成的接口面）。 */
static int vb6_TabNavKey(const MSG *msg);


// 全局变量
HINSTANCE g_hInstance = NULL;
static int g_nextControlId = 100;  // 控件ID从100开始 (1-99保留给菜单)
// 账 #156: 计时器 id **不能**跟着控件 id 走。控件 id 每建一枚窗体就复位一次
// (cgen_form_create_controls.inc 发 vb6_ResetControlId())，而 g_timerTable 是进程内
// 一张表、派发只按 id 找 (`vb6_SlotById`)。两者共用一个计数器 ⇒ 第二枚窗体的 Timer
// 一定拿到与第一枚相同的 id，于是：
//   1) 第一枚 Timer 若已被 Enabled=False 停掉，winmm 回调查到的那一格 running=0 ⇒ 一次都不投;
//   2) 若还活着，WM_TIMER 会投到**前一枚窗体**、跑前一枚窗体的事件过程。
// ⇒ 计时器要自己一枚永不复位的计数器。id 只在 WM_TIMER 这一路用，与控件/菜单 id 不同命名空间。
static int g_nextTimerId = 1000;

// 模态窗体状态
static HWND g_modalOwner = NULL;   // 被禁用的父窗口 (模态时)
static int g_modalResult = 0;      // 模态返回值

// Form_Unload回调类型:
// 返回0=允许关闭, 返回1=取消关闭 (对应VB6 vbCancel)
typedef int (*vb6_FormUnloadCallback)(void);

// ============================================================
// comctl32 通用控件引导
// ------------------------------------------------------------
// VB6 的 ListView/TreeView/Toolbar/StatusBar/ProgressBar/ImageList 在 VB6 侧
// 是 mscomctl.ocx 里的包装, 而它的渲染内核就是 comctl32.dll 的通用控件。
// C3 不用 OCX (本机未注册 + 32 位 inproc 进不了 x64 进程), 直接用 Win32
// 等价类复刻 —— 但通用控件类**必须**先经 InitCommonControlsEx 注册, 否则
// CreateWindowExW("msctls_progress32", ...) 返回 NULL 且不报错。
// 幂等; 唯一的副作用是加载 comctl32.dll。
// ============================================================
static int g_comCtlInited = 0;

void vb6_ComCtl_Init(void) {
#ifdef _WIN32
    if (g_comCtlInited) return;
    g_comCtlInited = 1;
    INITCOMMONCONTROLSEX ice;
    memset(&ice, 0, sizeof(ice));
    ice.dwSize = sizeof(ice);
    ice.dwICC = ICC_PROGRESS_CLASS | ICC_TAB_CLASSES | ICC_LISTVIEW_CLASSES
              | ICC_TREEVIEW_CLASSES | ICC_BAR_CLASSES | ICC_COOL_CLASSES
              | ICC_ANIMATE_CLASS | ICC_UPDOWN_CLASS | ICC_HOTKEY_CLASS
              | ICC_DATE_CLASSES | ICC_WIN95_CLASSES;
    InitCommonControlsEx(&ice);
    // C29-RT: RichTextBox 的类**不在 comctl32 里** —— RICHEDIT50W 由 Msftedit.dll 在
    // DllMain 里注册（实测：LoadLibrary 之后 GetClassInfoExW 才问得到）。这与上面那批
    // ICC_* 是两条路子：InitCommonControlsEx 管不到它。
    // 原来这里还挂了一句 "LoadLibrary(riched20.dll) 兜底" —— 本机 x64+x86 各测一遍
    // （读数：.build\richcls_out.txt）：riched20.dll 只注册 RichEdit20W，问 RICHEDIT50W
    // 回 err=1411（类不存在），而发码里的类名是写死的字面量 ⇒ 那句是个**假出口**：加载
    // 成功也照样建不出窗口，只是把"依赖缺失"伪装成"已经处理过了"。现在不兜底，改成当场
    // 喊出来（真缺 Msftedit 的场景 = Server Core / 精简镜像；带 GUI 的机器上它一直在，
    // SysWOW64 里也有 32 位那份）。VB6 自己的 RichTextBox 6.0 同样硬依赖 Msftedit。
    {
        HMODULE richMod = LoadLibraryW(L"Msftedit.dll");
        WNDCLASSEXW richCls;
        memset(&richCls, 0, sizeof(richCls));
        richCls.cbSize = sizeof(richCls);
        if (!richMod) {
            fprintf(stderr, "[C3_FORMS] Msftedit.dll not loadable (err=%lu): no RICHEDIT50W class"
                            " => every RichTextBox will have no window\n",
                    (unsigned long)GetLastError());
        } else if (!GetClassInfoExW(richMod, L"RICHEDIT50W", &richCls)) {
            fprintf(stderr, "[C3_FORMS] Msftedit.dll loaded but RICHEDIT50W is not registered"
                            " (err=%lu)\n", (unsigned long)GetLastError());
        }
    }
    // comctl32 只注册它自己那批类; msctls_status32 / msctls_toolbar32 这两个
    // v5.82 与 v6 都不注册 (实测连 dwICC=0xFFFFFFFF 全开也没用), 由各控件的
    // RTL 自己补注册 (vb6_StatusBar_RegisterClass 等)。
    vb6_StatusBar_RegisterClass();
#else
    (void)g_comCtlInited;
#endif
}

// 当前窗体的Unload回调 (每个窗体单独设置)
static vb6_FormUnloadCallback g_formUnloadCb = NULL;

// Timer回调类型
typedef void (*vb6_TimerCallback)(void);

// Timer回调表 (控件ID → 回调函数)
#define VB6_MAX_TIMERS 32
// C29-T: 每一格多带三样东西 —— key(Timer 控件自己的不可见句柄, 运行期靠它找回这一格)、
// period(当前周期) 与 winmm 那一档的句柄/标志。
struct vb6_TimerSlot {
    int   timerId;
    HWND  hwnd;       // 派发窗 = 所属窗体 (case WM_TIMER 在它上面)
    HWND  key;        // 身份窗 = Timer 控件自己的句柄
    vb6_TimerCallback callback;
    UINT  period;     // ms
    int   running;
    int   useMm;      // 1 = winmm timeSetEvent, 0 = 退回 SetTimer
    UINT  mmId;
};
static struct vb6_TimerSlot g_timerTable[VB6_MAX_TIMERS];
static int g_timerCount = 0;

// Fix 181: VB6 控件的默认字体是 **MS Sans Serif 8.25pt**，而现代 Windows 的
// DEFAULT_GUI_FONT 是 Segoe UI 9pt —— 明显更宽，于是 .frm 里按 VB6 字体排好的
// 固定宽度控件被截字 (demo: "ToolTipText"→"oolTipTex"、"Sort Desc."→"ort Des.")。
// 几何与 DPI 换算本身没错 (1215 缇 = 81px 正好放得下 VB6 字体的 11 个字符)，
// 错的是字体选择，所以只换字体、不动尺寸。
// ⚠ **必须每控件一份，不能进程内缓存共享**：`vb6forms_ctrl.c` 的
// vb6_SetControlFontFromLogFont 换字体后无条件 `DeleteObject(hOldFont)` (它只把
// stock 对象算作安全的 no-op)。共享自建字体一旦被某个改 Font 属性的控件删掉，
// 其余仍在用它的控件就拿着已销毁的 GDI 句柄 (句柄还会被复用 → 别的控件拿到同一
// 数值)，表现为跨控件随机换字形，极难归因。失败时退回 stock 对象同样安全。
static HFONT vb6_Vb6DefaultGuiFont(void) {
    LOGFONTW lf = {0};
    HDC hdc = GetDC(NULL);
    int dpiY = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
    HFONT hFont;
    if (hdc) ReleaseDC(NULL, hdc);
    lf.lfHeight = -MulDiv(825, dpiY, 7200);   /* 8.25pt → 像素 */
    lf.lfWeight = FW_NORMAL;
    // Fix 190: 保持 DEFAULT_CHARSET —— 它让 GDI 按**字符串实际字符**做字体链接,
    // 于是同一控件既能画 CJK 也能画韩文/西里尔/希腊文; 写死某个具体代码页字符集
    // (如 GB2312_CHARSET) 会让系统在字体里找不到韩文/俄文字形, 显示成方框。
    lf.lfCharSet = DEFAULT_CHARSET;
    lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
    lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
    // VB6 的 MS Sans Serif 是点阵字体, 从不抗锯齿; 开了 ClearType 会比参考图更宽更糊
    lf.lfQuality = NONANTIALIASED_QUALITY;
    lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    wcscpy(lf.lfFaceName, L"MS Sans Serif");
    hFont = CreateFontIndirectW(&lf);
    if (!hFont) hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    return hFont;
}

// ============================================================
// 缇(Twip)转换
// ============================================================

int vb6_TwipToX(int twips) {
    // 1缇 = 1/15像素 (96 DPI标准)
    // Screen.TwipsPerPixelX 通常=15
    return MulDiv(twips, vb6_DpiX(), 1440);
}

int vb6_TwipToY(int twips) {
    return MulDiv(twips, vb6_DpiY(), 1440);
}

// Fix 184: 缇/像素换算必须走**真实 DPI**。此前 TwipToX 写死 /15 (96DPI)，
// 而 vb6_GetScaleWidth / vb6_Screen_TwipsPerPixelX 用 GetDeviceCaps(LOGPIXELSX)，
// 于是在 dpiAware=true 的工程里两套口径混用：Form_Resize 拿到
// ScaleWidth = 客户区像素*12 (120DPI)，再交给 TwipToX 按 /15 落成像素，
// 控件被缩小 20% (VBFlexGridDemo: 网格 906x385 -> 729x291，可见行 23 -> 13)。
// 现在所有换算共用下面这一对 DPI 源。
int vb6_DpiX(void) {
    static int s_dpi = 0;
    if (!s_dpi) {
        HDC hdc = GetDC(NULL);
        s_dpi = hdc ? GetDeviceCaps(hdc, LOGPIXELSX) : 96;
        if (hdc) ReleaseDC(NULL, hdc);
        if (s_dpi <= 0) s_dpi = 96;
    }
    return s_dpi;
}

int vb6_DpiY(void) {
    static int s_dpi = 0;
    if (!s_dpi) {
        HDC hdc = GetDC(NULL);
        s_dpi = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
        if (hdc) ReleaseDC(NULL, hdc);
        if (s_dpi <= 0) s_dpi = 96;
    }
    return s_dpi;
}

int vb6_XToTwipX(int px) { return MulDiv(px, 1440, vb6_DpiX()); }
int vb6_YToTwipY(int px) { return MulDiv(px, 1440, vb6_DpiY()); }

// ============================================================
// 窗体框架
// ============================================================

static void vb6_formClassBg_set(const char* className, int bg);

int vb6_RegisterFormClassBg(const char* className, void* wndProc, void* hInstance,
                            int iconResId, int backColor) {
    // Fix 190: 用 W 版注册窗口类。RegisterClassExA 注册出来的类即使之后用
    // CreateWindowExW 创建, 窗口本身仍然是 ANSI 窗口 —— 控件/标题文本会被系统
    // 按 ACP 来回转换, 多语言必然乱码。类名在 RTL 内部以 UTF-8 传递, 此处转宽。
    wchar_t wcls[128];
    vb6_u8ToWideBuf(className, wcls, 128);

    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;  // 支持双击
    wc.lpfnWndProc = (WNDPROC)wndProc;
    wc.hInstance = (HINSTANCE)hInstance;
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    // czUI fix: .frm 的窗体级 BackColor — 用 .frm 颜色做类背景刷, 否则窗体永远
    // 是 BTNFACE 灰 (czForm Demo 深蓝底变灰底). backColor<0 = 未指定, 走 VB6 默认.
    if (backColor >= 0) {
        COLORREF cref = (backColor & 0x80000000L)
                            ? GetSysColor(backColor & 0xFF)
                            : (COLORREF)backColor;
        wc.hbrBackground = CreateSolidBrush(cref);
        vb6_formClassBg_set(className, (int)cref);
    } else {
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);  // VB6默认灰色背景
    }
    wc.lpszClassName = wcls;

    // czUI fix: 未指定图标时不再回退到 IDI_APPLICATION — VB6 窗体没有 Icon 属性时
    // 标题栏就是没有图标 (UserControl.Parent.Icon = Nothing, czUI 自绘标题栏因此
    // 不画图标); 类图标留空, 系统在任务栏等处自动用默认图标, 行为一致。
    if (iconResId > 0) {
        wc.hIcon = LoadIconW((HINSTANCE)hInstance, MAKEINTRESOURCEW(iconResId));
        wc.hIconSm = wc.hIcon;
    }

    if (!RegisterClassExW(&wc)) {
        if (GetEnvironmentVariableA("C3_FORMS_TRACE", NULL, 0) > 0)
            fprintf(stderr, "[C3_FORMS] RegisterClassExW cls='%ls' FAILED err=%lu\n",
                    wcls, (unsigned long)GetLastError());
        return -1;
    }
    if (GetEnvironmentVariableA("C3_FORMS_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_FORMS] RegisterClassExW cls='%ls' ok\n", wcls);
    return 0;
}

int vb6_RegisterFormClass(const char* className, void* wndProc, void* hInstance, int iconResId) {
    return vb6_RegisterFormClassBg(className, wndProc, hInstance, iconResId, -1);
}

// ---- 窗体类背景色登记 (供 UserControl Ambient.BackColor 查询) ----
#define VB6_FORMBG_MAX 64
static struct { char className[128]; int bg; } g_formBg[VB6_FORMBG_MAX];
static int g_formBgCount = 0;

static void vb6_formClassBg_set(const char* className, int bg) {
    for (int i = 0; i < g_formBgCount; i++) {
        if (strcmp(g_formBg[i].className, className) == 0) { g_formBg[i].bg = bg; return; }
    }
    if (g_formBgCount < VB6_FORMBG_MAX) {
        snprintf(g_formBg[g_formBgCount].className, sizeof(g_formBg[0].className), "%s", className);
        g_formBg[g_formBgCount].bg = bg;
        g_formBgCount++;
    }
}

int vb6_Forms_QueryClassBg(const char* className) {
    for (int i = 0; i < g_formBgCount; i++) {
        if (strcmp(g_formBg[i].className, className) == 0) return g_formBg[i].bg;
    }
    return -1;
}

void* vb6_CreateFormWindowB(const char* className, const char* formName,
    int x, int y, int width, int height, void* hInstance, void* userData,
    int borderStyle);

void* vb6_CreateFormWindow(const char* className, const char* formName,
    int x, int y, int width, int height, void* hInstance, void* userData) {
    return vb6_CreateFormWindowB(className, formName, x, y, width, height,
                                 hInstance, userData, 2 /* Sizable */);
}

void* vb6_CreateFormWindowB(const char* className, const char* formName,
    int x, int y, int width, int height, void* hInstance, void* userData,
    int borderStyle) {
    // VB6坐标是缇, 转为像素
    int pw = vb6_TwipToX(width);
    int ph = vb6_TwipToY(height);

    // 创建窗口。
    // czUI fix: 尊重 .frm 的 BorderStyle — BorderStyle=0 (None) 是无系统标题栏
    // 的无边框窗口 (czUI 自绘标题栏此前叠在系统标题栏下面, 顶部多出一截)。
    // Fix 081k: Do NOT add WS_VISIBLE here; ShowWindow is called by vb6_ShowForm after Form_Load.
    // Fix 124: WS_CLIPCHILDREN —— 窗体自身重绘(背景填充)时必须把子控件区域裁剪掉,
    // 否则窗体重绘会把已经画好的子控件整片覆盖 (表现为"控件时有时无/干脆看不见",
    // 而离屏 dump 一切正常)。
    DWORD style;
    DWORD exStyle = 0;
    switch (borderStyle) {
        case 0:  /* None */
            style = WS_POPUP | WS_CLIPCHILDREN;
            exStyle = WS_EX_APPWINDOW;   /* 仍在任务栏显示 */
            break;
        case 1: case 3:  /* Fixed Single / Fixed Dialog */
            style = WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_CLIPCHILDREN;
            break;
        case 4: case 5:  /* ToolWindow */
            style = WS_CAPTION | WS_SYSMENU | WS_CLIPCHILDREN;
            break;
        default:         /* 2 = Sizable (VB6 标准) */
            style = WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN;
            break;
    }

    // 调整窗口大小使客户区匹配指定大小
    RECT rc = {0, 0, pw, ph};
    AdjustWindowRectEx(&rc, style, FALSE, exStyle);
    int winW = rc.right - rc.left;
    int winH = rc.bottom - rc.top;

    int px, py;
    if (x == -1 && y == -1) {
        // M22-Issue2: CenterScreen -- 用窗口实际尺寸计算居中位置
        px = (GetSystemMetrics(SM_CXSCREEN) - winW) / 2;
        py = (GetSystemMetrics(SM_CYSCREEN) - winH) / 2;
    } else {
        px = (x == CW_USEDEFAULT) ? CW_USEDEFAULT : vb6_TwipToX(x);
        py = (y == CW_USEDEFAULT) ? CW_USEDEFAULT : vb6_TwipToY(y);
    }

    // Fix 190: 类名/标题在 RTL 内是 UTF-8, 窗口层要 UTF-16
    wchar_t wcls[128];
    vb6_u8ToWideBuf(className, wcls, 128);
    wchar_t* wtitle = vb6_u8ToWideDup(formName);

    HWND hwnd = CreateWindowExW(
        exStyle,
        wcls,
        wtitle ? wtitle : L"",
        style,
        px, py,
        winW, winH,
        NULL,   // 无父窗口
        NULL,   // 无菜单
        (HINSTANCE)hInstance,
        userData  // 传递给WM_CREATE
    );

    if (GetEnvironmentVariableA("C3_FORMS_TRACE", NULL, 0) > 0) {
        wchar_t pb[64] = {0};
        GetWindowTextW(hwnd, pb, 64);
        fprintf(stderr, "[C3_FORMS] CreateFormWindowB cls='%s' title='%s' -> hwnd=%p len=%d titleW='%ls' isUni=%d\n",
                className, formName ? formName : "", (void*)hwnd,
                GetWindowTextLengthW(hwnd), pb, IsWindowUnicode(hwnd));
        fflush(stderr);
    }

    free(wtitle);
    return (void*)hwnd;
}

// Fix 162d-extlist: 经典 groupbox（Fix 162c 关主题后）只画蚀刻框、不填内部；
// 而窗体带 WS_CLIPCHILDREN（Fix 124）时窗体重绘不往子控件矩形下涂底色，
// Frame 内部就成了"从未画过"的白色（实测 (255,255,255)，VB6 参考图是 240 灰）。
// VB6 观感 = Frame 内部透出窗体 BackColor → 子类化 WM_ERASEBKGND，
// 用父窗体底色填（vb6_GetControlBackColor 未显式设置时回落 BTNFACE，
// 与 VB6 默认 &H8000000F 等效）。
// Fix 163-extlist: 关掉一个控件的 comctl6 视觉样式 (SetWindowTheme(hwnd,L"",L""))。
// 为什么要关: 带视觉样式的 Button 类 (命令按钮/单选钮) 的**文字**由主题绘制器渲染 ——
// 它用主题自己的字体并在 ClearType 下描边, 结果明显比窗体标题/MS Sans Serif 点阵糊。
// ⚠ 复选框在本机实测**没有**被主题化 (原生就清晰), 但单选钮、命令按钮被主题化了
// (单选钮圆点是蓝色 = 主题标记), 二者视觉不一致。统一关样式后都退回经典 GDI
// 文本渲染 (走我们下发的 MS Sans Serif 8.25pt + NONANTIALIASED), 与标题/复选框齐平。
// ⚠ 这也正是经典 Windows 9x/VB6 的观感 —— VB6 运行时本来就禁用该控件的主题。
// uxtheme 走 LoadLibrary 动态取 (不新增 import lib, 与 Frame 的处理同路子)。
static void vb6_DisableControlTheme(HWND hwnd) {
    HMODULE themeDll = LoadLibraryW(L"uxtheme.dll");
    if (!themeDll) return;
    typedef HRESULT (WINAPI *pfnSetTheme)(HWND, LPCWSTR, LPCWSTR);
    pfnSetTheme setTheme = (pfnSetTheme)(void*)GetProcAddress(themeDll, "SetWindowTheme");
    if (setTheme) setTheme(hwnd, L"", L"");
    FreeLibrary(themeDll);
}

// Fix 162f-extlist: 取 groupbox 的标题底色 (父窗 BackColor, 未设回落 BTNFACE)。
static COLORREF vb6_GBoxTitleBg(HWND hwnd) {
    HWND parent = GetParent(hwnd);
    COLORREF bg = parent ? (COLORREF)vb6_GetControlBackColor((void*)parent)
                         : GetSysColor(COLOR_BTNFACE);
    if (bg & 0x80000000L) bg = GetSysColor(bg & 0xFF);
    return bg;
}

static LRESULT CALLBACK vb6_GroupBoxSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    WNDPROC orig = (WNDPROC)GetPropW(hwnd, L"VB6_GBox_OrigProc");
    if (msg == WM_ERASEBKGND) {
        COLORREF bg = vb6_GBoxTitleBg(hwnd);
        RECT rc;
        GetClientRect(hwnd, &rc);
        HBRUSH br = CreateSolidBrush(bg);
        if (br) { FillRect((HDC)wp, &rc, br); DeleteObject(br); }
        return 1;
    }
    // Fix 162f-extlist: 经典 groupbox (关主题后) 画标题时, 标题底用的是**系统默认
    // 白刷** (不是 WM_CTLCOLORBTN, 它只发给命令按钮; 实测 Frame 根本不发这条) ——
    // 于是标题后面留一条**纯白填充带** (用户实测 (255,255,255), 周围是 240 灰)。
    // VB6 里 Frame 标题坐在父窗 BackColor 上 ⇒ 这里自己接管绘制:
    //   ① 先把整个客户区填成父窗底色 (灰);
    //   ② 让原过程画蚀刻边框 + 标题文字 (此时标题底已是灰);
    //   ③ 再单独把**标题文字下面那条带**重新填灰 (对付它在白底上画的文字残影)。
    // 判据: 只对有非空 Caption 的 groupbox 这么干 (空标题 Frame2 无此带)。
    if (msg == WM_PAINT) {
        wchar_t cap[256] = {0};
        GetWindowTextW(hwnd, cap, 256);
        LRESULT r = orig ? CallWindowProcW(orig, hwnd, msg, wp, lp)
                         : DefWindowProcW(hwnd, msg, wp, lp);
        if (cap[0]) {
            HDC hdc = GetDC(hwnd);
            if (hdc) {
                RECT rc; GetClientRect(hwnd, &rc);
                // 标题带高度: 用当前字体算 (经典 groupbox 标题约一行高)。
                HFONT hf = (HFONT)SendMessageW(hwnd, WM_GETFONT, 0, 0);
                HFONT old = hf ? (HFONT)SelectObject(hdc, hf) : NULL;
                TEXTMETRICW tm; ZeroMemory(&tm, sizeof(tm));
                GetTextMetricsW(hdc, &tm);
                int th = tm.tmHeight + 2;
                // 标题起点: 经典 groupbox 标题缩进 ~7px (含 3px 蚀刻间距)。
                int tx = 7;
                SIZE sz = {0, 0};
                GetTextExtentPoint32W(hdc, cap, lstrlenW(cap), &sz);
                COLORREF bg = vb6_GBoxTitleBg(hwnd);
                RECT trc = { tx - 1, 0, tx + sz.cx + 1, th };
                // 先擦掉白色底 + 白底上画的文字, 再以灰底重画文字。
                HBRUSH br = CreateSolidBrush(bg);
                if (br) { FillRect(hdc, &trc, br); DeleteObject(br); }
                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(hdc, GetSysColor(COLOR_BTNTEXT));
                RECT trc2 = { tx, 0, tx + sz.cx + 2, th };
                DrawTextW(hdc, cap, -1, &trc2, DT_LEFT | DT_SINGLELINE | DT_VCENTER);
                if (old) SelectObject(hdc, old);
                ReleaseDC(hwnd, hdc);
            }
        }
        return r;
    }
    if (msg == WM_CTLCOLORBTN) {
        // Fix 162e-extlist: Frame 内的 CheckBox/OptionButton 向本 Frame 要背景刷 ——
        // 与窗体侧同口径 (见 vb6_CtlColorBtnBrush): 复选框/单选钮走空刷 (透明),
        // 嵌套 groupbox 走本 Frame 底色的实心刷。
        return vb6_CtlColorBtnBrush((HWND)lp, hwnd);
    }
    if (msg == WM_DESTROY) {
        if (orig) {
            RemovePropW(hwnd, L"VB6_GBox_OrigProc");
            SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)orig);
        }
        return 0;
    }
    return orig ? CallWindowProcW(orig, hwnd, msg, wp, lp)
                : DefWindowProcW(hwnd, msg, wp, lp);
}

// ============================================================
// 控件创建
// ============================================================

void* vb6_CreateControl(const char* win32Class, const char* controlName,
    long style, long exStyle,
    int x, int y, int width, int height,
    int id, void* hParent, void* hInstance) {
    int px = vb6_TwipToX(x);
    int py = vb6_TwipToY(y);
    int pw = vb6_TwipToX(width);
    int ph = vb6_TwipToY(height);

    // Fix 190: 标准控件类名 (Button/Edit/Static/ListBox/ComboBox/ScrollBar) 在
    // CreateWindowExW 下会创建出**Unicode 版本**的控件 —— 这才是能正确显示
    // 中文/韩文/俄文的前提 (ANSI 版 ListBox 收到 SendMessageW(LB_ADDSTRING, 宽串)
    // 会把 UTF-16 当字节串读, 直接乱码/截断)。
    // 注意: 窗口类注册 (RegisterClassExW) 与窗口创建 (CreateWindowExW) 必须同为 W,
    // 混用会被系统当成 ANSI 窗口 (见 Fix 081l 的反面教训)。
    wchar_t wcls[128];
    vb6_u8ToWideBuf(win32Class, wcls, 128);
    wchar_t* wname = vb6_u8ToWideDup(controlName);

    HWND hwnd = CreateWindowExW(
        (DWORD)exStyle,
        wcls,
        wname ? wname : L"",
        (DWORD)style,
        px, py, pw, ph,
        (HWND)hParent,
        (HMENU)(intptr_t)id,
        (HINSTANCE)hInstance,
        NULL
    );
    free(wname);

    /* 账 #163/#168：把**创建时**的 WS_TABSTOP 决定记在窗口属性上。运行期不能只看样式位 —— 
       BS_AUTORADIOBUTTON 那一族的 WS_TABSTOP 会被系统自己挪到「勾选那枚」身上（裸码探针
       `.build/cp2`：创建时两枚都不带，勾选那枚后来带着走），而 VB6 的 TabStop 是设计期属性、
       运行期不改 ⇒ 「谁是站」必须以创建时那一份为准。值写 1/2 而不是 1/0：
       SetPropW(…, 0) 等于删属性（账 #107 那一课）。 */
    if (hwnd) SetPropW(hwnd, L"VB6_TabStop", (HANDLE)(DWORD_PTR)((style & WS_TABSTOP) ? 1 : 2));

    // 设置默认字体 (VB6使用MS Sans Serif 8.25pt)
    if (hwnd) {
        // Fix 181: 原先直接用 GetStockObject(DEFAULT_GUI_FONT) —— 现代 Windows 上
        // 那是 Segoe UI 9pt, 比 VB6 的 MS Sans Serif 8.25pt 宽, 控件标题被截字。
        HFONT hFont = vb6_Vb6DefaultGuiFont();
        if (!hFont) {
            hFont = CreateFontW(
                -11, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                DEFAULT_QUALITY, FF_DONTCARE, L"MS Sans Serif"
            );
        }
        SendMessage(hwnd, WM_SETFONT, (WPARAM)hFont, MAKELPARAM(FALSE, 0));

        // Fix 162c-extlist: Frame(BS_GROUPBOX) 关掉 comctl6 主题化 —— 主题版的
        // groupbox 会用白色填掉整个内部 (子控件的 240 灰底反而成了色块), VB6
        // 参考图是经典蚀刻边框 + 透出窗体 BTNFACE 底色。classic groupbox
        // 内部透明, 观感与 VB6 一致。
        // Fix 163-extlist: **所有 BUTTON 类**都关样式 —— 命令按钮/单选钮的主题文字
        // 渲染比 MS Sans Serif 点阵糊 (见 vb6_DisableControlTheme 注释)。
        // 复选框本机未被主题化, 关掉是无害的 no-op; 关样式统一了整族观感。
        if (_stricmp(win32Class, "BUTTON") == 0) {
            vb6_DisableControlTheme(hwnd);
        }
        // Fix 162d-extlist: Frame 关主题后内部不再白填, 但也没有人涂底色了 ——
        // 子类化补上"填父窗体底色"(见 vb6_GroupBoxSubclassProc 注释)。
        if (((style & 0x0000000FL) == 0x7L) && _stricmp(win32Class, "BUTTON") == 0) {
            if (!GetPropW(hwnd, L"VB6_GBox_OrigProc")) {
                WNDPROC gOrig = (WNDPROC)SetWindowLongPtrW(hwnd, GWLP_WNDPROC,
                                                           (LONG_PTR)vb6_GroupBoxSubclassProc);
                if (gOrig) SetPropW(hwnd, L"VB6_GBox_OrigProc", (HANDLE)gOrig);
            }
        }

        /* Fix 145: ComboBox 下拉列表高度.
         * Win32 的 ComboBox 窗口高度 = 显示行 + 下拉列表高度; 而 .frm 里
         * ComboBox.Height 只是**显示区**的一行高度 (VB6 语义: 下拉部分由系统
         * 默认项数决定). 直接拿 .frm 高度当窗口高度会让下拉区 ≈ 0 —
         * 实测下拉弹出窗口只有 2px, 用户拉开只看到第一项.
         * 这里补足到 VB6 的观感: 下拉约显示 9~10 项 (VB6 ThunderRT6ComboBox
         * 实测下拉区 114px; 本机 itemH=12px → 10 项 ≈ 120px). */
        if (_stricmp(win32Class, "COMBOBOX") == 0) {
            int itemH = (int)SendMessage(hwnd, CB_GETITEMHEIGHT, 0, 0);
            if (itemH <= 0) itemH = (int)SendMessage(hwnd, CB_GETITEMHEIGHT, (WPARAM)-1, 0);
            if (itemH <= 0) itemH = 16;
            int listPx = ph - itemH;              /* .frm 高度里减去一行 = 下拉区 */
            if (listPx < itemH * 4) listPx = itemH * 10;  /* 不足则用 VB6 观感项数 */
            SetWindowPos(hwnd, NULL, 0, 0, pw, itemH + listPx,
                         SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
            if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
                RECT rr; GetWindowRect(hwnd, &rr);
                fprintf(stderr, "[C3_FIX] ComboBox itemH=%d frmH(px)=%d setH=%d actualH=%d\n",
                        itemH, ph, itemH + listPx, (int)(rr.bottom - rr.top));
            }
        }

        // Fix 187: PictureBox/Image = STATIC+SS_BITMAP, 由 RTL 自绘接管
        // (vb6_ImageSubclassProc: 先按 VB6_BackColor 属性填背景, 再画 Picture)。
        // 此前无图 PictureBox 是裸 SS_BITMAP STATIC — 无文字 STATIC 的
        // WM_CTLCOLORSTATIC 擦背景路径可走, BackColor 写了没人消费
        // (Test.exe 颜色对话框 OK 后色块不变色的第三层根因)。
        // 生成器对 PictureBox/Image 固定发 SS_BITMAP(0x0E), 用它识别;
        // Label 是 SS_LEFT|SS_NOTIFY 不含该位, 不受影响。
        if (_stricmp(win32Class, "STATIC") == 0
            && ((style & 0x0000000FL) == 0x0000000EL)) {  // SS_BITMAP
            vb6_InstallImageSubclass(hwnd);
        }
    }

    return (void*)hwnd;
}

int vb6_NextControlId(void) {
    return g_nextControlId++;
}

void vb6_ResetControlId(void) {
    g_nextControlId = 100;
}

// ============================================================
// Timer管理
// ============================================================

// winmm 按需取（同一进程只载一次）。精度那一档的关键：SetTimer 的粒度地板是系统计时
// 周期(~15.6 ms)，实测 Interval=20 只能拿到 ~34.5 ms/tick；winmm 的 timeSetEvent 配
// wResolution=1 是 ms 级。winmm 只经 LoadLibrary + GetProcAddress 取（照 GDI+ /
// comdlg32 那两条例子）=> 不新增 import lib；取不到就退回 SetTimer，功能不变、只是粗。
// WIN32_LEAN_AND_MEAN 把 mmsystem.h 挡在外面了，所以常量按本文件既有的"固定数值"写法自带
typedef UINT (WINAPI *vb6_pfnTimeSetEvent)(UINT, UINT, void (CALLBACK*)(UINT, UINT, DWORD, DWORD, DWORD), DWORD, DWORD);
typedef UINT (WINAPI *vb6_pfnTimeKillEvent)(UINT);
static vb6_pfnTimeSetEvent  vb6_pTimeSetEvent  = NULL;
static vb6_pfnTimeKillEvent vb6_pTimeKillEvent = NULL;
static int vb6_mmProbed = 0;

static void vb6_MmProbe(void) {
    if (vb6_mmProbed) return;
    vb6_mmProbed = 1;
    HMODULE h = LoadLibraryW(L"winmm.dll");
    if (!h) return;
    vb6_pTimeSetEvent  = (vb6_pfnTimeSetEvent)(void*)GetProcAddress(h, "timeSetEvent");
    vb6_pTimeKillEvent = (vb6_pfnTimeKillEvent)(void*)GetProcAddress(h, "timeKillEvent");
}

static struct vb6_TimerSlot* vb6_SlotById(int id) {
    for (int i = 0; i < g_timerCount; i++)
        if (g_timerTable[i].timerId == id) return &g_timerTable[i];
    return NULL;
}

// winmm 的回调跑在它自己的线程上，**不能**直接调 VB 的事件处理函数（生成代码里那个
// 事件体属于 UI 线程）。所以只把到期事件投回派发窗，仍走原来的
// case WM_TIMER -> vb6_DispatchTimer，语义与 SetTimer 一致，只是不再等 15.6 ms 的地板。
static void CALLBACK vb6_MmThunk(UINT mmCssId, UINT msg, DWORD user, DWORD dw1, DWORD dw2) {
    (void)mmCssId; (void)msg; (void)dw1; (void)dw2;
    struct vb6_TimerSlot* e = vb6_SlotById((int)user);
    if (e && e->running) PostMessageW(e->hwnd, WM_TIMER, (WPARAM)e->timerId, 0);
}

static void vb6_TimerStop(struct vb6_TimerSlot* e) {
    if (!e->running) return;
    if (e->useMm) {
        if (vb6_pTimeKillEvent) vb6_pTimeKillEvent(e->mmId);
        e->mmId = 0; e->useMm = 0;
    } else {
        KillTimer(e->hwnd, (UINT_PTR)e->timerId);
    }
    e->running = 0;
}

static void vb6_TimerStart(struct vb6_TimerSlot* e) {
    // VB6 口径: Interval 合法域 1..65535，0 = 不跑
    if (e->running || e->period == 0 || e->period > 65535) return;
    vb6_MmProbe();
    if (vb6_pTimeSetEvent) {
        UINT id = vb6_pTimeSetEvent(e->period, 1, vb6_MmThunk, (DWORD)e->timerId, TIME_PERIODIC);
        if (id != 0) { e->mmId = id; e->useMm = 1; e->running = 1; return; }
    }
    if (SetTimer(e->hwnd, (UINT_PTR)e->timerId, e->period, NULL)) e->running = 1;
}

// 挂一枚计时器。owner = 派发窗（窗体），key = Timer 控件自己的不可见句柄 ——
// 有了 key，运行期 `Timer1.Enabled = True` / `Timer1.Interval = 100` 才找得回这一格。
void vb6_TimerAttach(void* owner, void* key, int period, void* callback, int enabled) {
    if (g_timerCount >= VB6_MAX_TIMERS) return;
    if (period < 0) period = 0;
    if (period > 65535) period = 65535;
    struct vb6_TimerSlot* e = &g_timerTable[g_timerCount];
    g_timerCount++;
    e->timerId = g_nextTimerId++;
    e->hwnd = (HWND)owner;
    e->key = (HWND)key;
    e->callback = (vb6_TimerCallback)callback;
    e->period = (UINT)period;
    e->running = 0; e->useMm = 0; e->mmId = 0;
    if (enabled) vb6_TimerStart(e);
}

// 运行期 Enabled（VB6 的 True 是 -1，别处传来的是非零值）
void vb6_TimerSetEnabled(void* key, int enabled) {
    for (int i = 0; i < g_timerCount; i++) {
        struct vb6_TimerSlot* e = &g_timerTable[i];
        if (e->key == (HWND)key) {
            if (enabled) vb6_TimerStart(e); else vb6_TimerStop(e);
            return;
        }
    }
}

// 运行期 Interval：改了立刻按新周期重排（VB6 就是这个行为，不是"下一轮才生效"）
void vb6_TimerSetPeriod(void* key, int period) {
    if (period < 0) period = 0;
    if (period > 65535) period = 65535;
    for (int i = 0; i < g_timerCount; i++) {
        struct vb6_TimerSlot* e = &g_timerTable[i];
        if (e->key == (HWND)key) {
            int was = e->running;
            vb6_TimerStop(e);
            e->period = (UINT)period;
            if (was) vb6_TimerStart(e);
            return;
        }
    }
}

// 兼容旧入口：没有身份窗时派发窗自己当身份，建完即启。
int vb6_SetTimer(void* hwnd, int interval, void* callback) {
    if (g_timerCount >= VB6_MAX_TIMERS) return -1;
    int id = g_nextTimerId++;
    struct vb6_TimerSlot* e = &g_timerTable[g_timerCount];
    g_timerCount++;
    e->timerId = id; e->hwnd = (HWND)hwnd; e->key = (HWND)hwnd;
    e->callback = (vb6_TimerCallback)callback;
    e->period = (UINT)(interval > 65535 ? 65535 : (interval < 0 ? 0 : interval));
    e->running = 0; e->useMm = 0; e->mmId = 0;
    vb6_TimerStart(e);
    return id;
}

void vb6_KillTimer(int timerId) {
    for (int i = 0; i < g_timerCount; i++) {
        if (g_timerTable[i].timerId == timerId) {
            vb6_TimerStop(&g_timerTable[i]);
            g_timerTable[i] = g_timerTable[g_timerCount - 1];
            g_timerCount--;
            break;
        }
    }
}

// P24-Timer: WndProc中分发WM_TIMER (替代消息循环拦截)
void vb6_DispatchTimer(int timerId) {
    for (int i = 0; i < g_timerCount; i++) {
        if (g_timerTable[i].timerId == timerId) {
            if (g_timerTable[i].callback) {
                g_timerTable[i].callback();
            }
            break;
        }
    }
}

// ============================================================
// Sub Main 驻留判据 (Fix 167)
// ============================================================
// VB6 语义: `Sub Main` 返回后进程**不退出**, 运行时继续泵消息, 直到所有窗体关闭
// (或显式 End)。此前 codegen 的 Sub Main 入口模板调完 Main 直接 vb6_Exit()+return,
// 于是 `Load` 出 modeless 窗体的工程一返回就干净退出 (退出码 0) —— 表现为
// "窗口闪一下就没了", VBFlexGridDemo 正是如此。
// 判据用"本线程有没有可见窗口", 而不是另建窗体注册表: 窗体是本 RTL 在
// vb6_CreateFormWindowB 里以 RegisterClass("VB6_Form_<X>") 建的普通窗口, 归本线程所有;
// 而 `App.PrevInstance` 那一支 (激活前一个实例后返回) 只操作**别的进程**的 hwnd,
// 本线程没有窗口 → 不会误驻留。纯 .bas 控制台工程同样没有可见窗口 → 不受影响。
static BOOL CALLBACK vb6_EnumAnyVisibleWindow(HWND hwnd, LPARAM lParam) {
    if (IsWindowVisible(hwnd)) { *(int*)lParam = 1; return FALSE; }
    return TRUE;
}

int vb6_AnyThreadWindowVisible(void) {
    int found = 0;
    EnumThreadWindows(GetCurrentThreadId(), vb6_EnumAnyVisibleWindow, (LPARAM)&found);
    if (GetEnvironmentVariableA("C3_COM_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_FSM] AnyThreadWindowVisible=%d tid=%lu\n",
                found, (unsigned long)GetCurrentThreadId());
        fflush(stderr);
    }
    return found;
}

// ============================================================
// 消息循环
// ============================================================

int vb6_MessageLoop(void) {
    MSG msg;
    // czUI fix: 消息循环启动后设计器 Timer 才允许触发 (VB6 语义: Timer 事件
    // 排队等消息循环; 否则处理器在 Form_Load 前对未就绪实例运行 → AV)
    extern int vb6_uc_timersStarted;  // czUI fix (定义在 vb6forms_uc.c)
    vb6_uc_timersStarted = 1;
    extern void vb6_Forms_LoopDepth(int delta);   // Fix 188 (定义在 vb6rtl_system.c)
    vb6_Forms_LoopDepth(1);
    if (GetEnvironmentVariableA("C3_COM_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_FSM] MessageLoop enter\n"); fflush(stderr);
    }
    // P20-44: OLE 拖放无头联测 —— 环境变量在时, **处理完第一条消息之后**给每个
    // 已注册目标发一次 DragEnter+Drop, 文本 "OLE-TEST-DROP"。
    // DoDragDrop 是模态循环, 无头环境没法真拖, 只能这样驱动目标侧的事件链。
    // **必须在消息循环内 fire**: OLEDropMode 的 Register 发生在延迟的 Form_Load
    // (PostMessage 0x7FF0) 里, 循环前 g_targetCount 还是 0, 什么都 fire 不到 (实测踩过)。
    {
        wchar_t oleFlag44[8] = { 0 };
        int oleFired44 = 0;
        if (GetEnvironmentVariableW(L"C3_OLEDDB_TEST", oleFlag44, 8) > 0) {
            extern void vb6_oleDD_FireTestDropAtRegistered(void);
            while (GetMessage(&msg, NULL, 0, 0)) {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
                if (!oleFired44) {
                    oleFired44 = 1;
                    vb6_oleDD_FireTestDropAtRegistered();
                }
            }
        } else {
            // 账 #83(b)：主泵也要走对话框式键盘导航，否则**普通（非模态）窗体按 Tab 不动**。
            // 模态那条循环本来就走了（`vb6_ShowForm` 里 `IsDialogMessageW`），实测在那儿
            // VK_TAB 真跳格（029 的「C29-FS-a 之后一测」），差的只有这一条泵。
            // 落点是 `GetActiveWindow()` —— Tab 是给"用户正在打字的那枚窗体"用的，
            // 拿 msg.hwnd 当对话框会把子控件句柄当容器传进去，找不着下一站。
            // 与模态那条同一个开关 `C3_OCX_NO_DLGMSG=1` 关掉：`IsDialogMessage` 会
            // **吞掉**它处理的那条按键消息，所以 `_KeyDown`/`_KeyPress` 里想看见 VK_TAB 的
            // 用法会被这一刀改变行为（存量实测：语料里 0 处这么写）。
            int useDlgMsgMain = (GetEnvironmentVariableW(L"C3_OCX_NO_DLGMSG", NULL, 0) <= 0);
            /* 账 #163/#168：VK_TAB 由自研导航器接管（关掉 = C3_OCX_NO_TABNAV=1，退回旧行为做 A/B）。 */
            int useTabNavMain = (GetEnvironmentVariableW(L"C3_OCX_NO_TABNAV", NULL, 0) <= 0);
            while (GetMessage(&msg, NULL, 0, 0)) {
                // P24-Timer: WM_TIMER现在由WndProc分发, 消息循环不再拦截
                if (useTabNavMain && vb6_TabNavKey(&msg)) continue;
                // 账 #83(b)：主泵也接管对话框式键盘导航
                HWND act = useDlgMsgMain ? GetActiveWindow() : NULL;
                if (!act || !IsDialogMessageW(act, &msg)) {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }
            }
        }
    }
    vb6_Forms_LoopDepth(-1);   // Fix 188
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_MODAL] 主消息循环结束: msg=0x%04X hwnd=%p\n",
                msg.message, (void*)msg.hwnd);
    return (int)msg.wParam;
}

int vb6_DoEvents(void) {
    MSG msg;
    int count = 0;
    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
        // P24-Timer: WM_TIMER现在由WndProc分发, DoEvents不再拦截
        TranslateMessage(&msg);
        DispatchMessage(&msg);
        count++;
        // 安全限制: 防止无限循环 (VB6 DoEvents行为: 处理完就返回)
        if (count > 1000) break;
    }
    return count;
}

// ============================================================
// 窗体事件桥接
// ============================================================

void vb6_SetFormUserData(void* hwnd, void* userData) {
    SetWindowLongPtrA((HWND)hwnd, GWLP_USERDATA, (LONG_PTR)userData);
}

void* vb6_GetFormUserData(void* hwnd) {
    return (void*)GetWindowLongPtrA((HWND)hwnd, GWLP_USERDATA);
}

// ============================================================
// 窗体工具函数
// ============================================================

void* vb6_GetAppInstance(void) {
    return (void*)g_hInstance;
}

void vb6_SetAppInstance(void* hInstance) {
    g_hInstance = (HINSTANCE)hInstance;
}

// Fix 149 诊断: C3_CRASH_TRACE=1 时安装未处理异常过滤器, 把崩溃栈各帧的
// 「模块+偏移」写进 c3_crash.txt。没有调试器也能一眼看出异常是从
// Test.exe 自己的代码抛的, 还是逃出第三方 OCX (NewTab01.ocx) 的 VB6 代码。
static LONG WINAPI vb6_crashFilter(EXCEPTION_POINTERS* ep) {
    FILE* f = fopen("c3_crash.txt", "a");
    if (!f) return EXCEPTION_EXECUTE_HANDLER;
    fprintf(f, "=== EXCEPTION code=0x%08lX addr=%p ===\n",
            (unsigned long)ep->ExceptionRecord->ExceptionCode,
            ep->ExceptionRecord->ExceptionAddress);
    {   /* 触发地址所在模块 */
        HMODULE hm = NULL;
        wchar_t wp[MAX_PATH] = {0};
        char nm[80] = "?";
        if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                               GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                               (LPCWSTR)ep->ExceptionRecord->ExceptionAddress, &hm) && hm) {
            GetModuleFileNameW(hm, wp, MAX_PATH);
            { const wchar_t* b = wcsrchr(wp, L'\\');
              WideCharToMultiByte(CP_ACP, 0, b ? b + 1 : wp, -1, nm, sizeof(nm), NULL, NULL); }
            fprintf(f, "  FAULT %s+0x%llX\n", nm,
                    (unsigned long long)((const char*)ep->ExceptionRecord->ExceptionAddress
                                         - (const char*)hm));
        } else {
            fprintf(f, "  FAULT (unknown module) %p\n", ep->ExceptionRecord->ExceptionAddress);
        }
    }
    {   void* frames[48];
        USHORT n = RtlCaptureStackBackTrace(0, 48, frames, NULL), i;
        for (i = 0; i < n; i++) {
            HMODULE hm = NULL;
            wchar_t wp[MAX_PATH] = {0};
            char nm[80] = "?";
            if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                   GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                   (LPCWSTR)frames[i], &hm) && hm) {
                GetModuleFileNameW(hm, wp, MAX_PATH);
                { const wchar_t* b = wcsrchr(wp, L'\\');
                  WideCharToMultiByte(CP_ACP, 0, b ? b + 1 : wp, -1, nm, sizeof(nm), NULL, NULL); }
                fprintf(f, "  #%02d 0x%p  %s+0x%llX\n", i, frames[i], nm,
                        (unsigned long long)((const char*)frames[i] - (const char*)hm));
            } else {
                fprintf(f, "  #%02d 0x%p  (unknown)\n", i, frames[i]);
            }
        }
    }
    fflush(f);
    fclose(f);
    return EXCEPTION_EXECUTE_HANDLER;
}

// 堆损坏 (0xC0000374) 不经过未处理异常过滤器: ntdll 持堆锁直接终止进程。
// VEH 第一 chance 截获落栈; 此线程持堆锁, 只用无堆分配的 Win32 API
// (CreateFileA/WriteFile/wsprintfA), CRT 文件 IO 会死锁。
/* 扫描栈内存: 打印落在目标镜像代码范围内的返回地址候选 (异常派发会截断
 * EBP 链, RtlCaptureStackBackTrace 拿不到应用帧; 直接按值扫描原始栈) */
static int vb6_vehScanStackForImage(char* buf, int n, int bufsz,
                                    CONTEXT* ctx, const char* imgName,
                                    const char* imgBase, ULONG imgSize) {
    int printed = 0, i;
#ifdef _WIN64
    const ULONG_PTR* sp = (const ULONG_PTR*)ctx->Rsp;
#else
    const DWORD* sp = (const DWORD*)ctx->Esp;
#endif
    for (i = 0; i < 16384 && printed < 48 && n < bufsz - 128; i++) {
        ULONG_PTR v;
        if (IsBadReadPtr(sp + i, sizeof(ULONG_PTR))) break;
        v = sp[i];
        if (v >= (ULONG_PTR)imgBase &&
            v < (ULONG_PTR)imgBase + imgSize) {
#ifdef _WIN64
            n += wsprintfA(buf + n, "  [sp+%d] 0x%016llX -> %s+0x%llX\r\n",
                           i, (unsigned long long)v, imgName,
                           (unsigned long long)(v - (ULONG_PTR)imgBase));
#else
            n += wsprintfA(buf + n, "  [sp+%d] 0x%08lX -> %s+0x%08lX\r\n",
                           i, (unsigned long)v, imgName,
                           (unsigned long)(v - (ULONG_PTR)imgBase));
#endif
            printed++;
        }
    }
    return n;
}

static LONG WINAPI vb6_heapCorruptVEH(EXCEPTION_POINTERS* ep) {
    DWORD code = ep->ExceptionRecord->ExceptionCode;
    int isCorrupt = (code == 0xC0000374);
    int isAV = (code == 0xC0000005);
    if (!isCorrupt && !isAV)
        return EXCEPTION_CONTINUE_SEARCH;
    HANDLE h = CreateFileA("c3_crash.txt", FILE_APPEND_DATA,
                           FILE_SHARE_READ | FILE_SHARE_WRITE, NULL,
                           OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (h == INVALID_HANDLE_VALUE) return EXCEPTION_CONTINUE_SEARCH;
    {
        char buf[16384];
        int n;
        if (isCorrupt)
            n = wsprintfA(buf, "=== HEAP CORRUPTION code=0xC0000374 addr=%p ===\r\n",
                          ep->ExceptionRecord->ExceptionAddress);
        else
            n = wsprintfA(buf, "=== AV code=0xC0000005 addr=%p %s %p ===\r\n",
                          ep->ExceptionRecord->ExceptionAddress,
                          ep->ExceptionRecord->ExceptionInformation[0] ? "WRITE" : "READ",
                          (void*)ep->ExceptionRecord->ExceptionInformation[1]);
        {   /* 模块名+运行时基址+偏移 (x86 ASLR 下基址随机, 符号化必须配对基址) */
            HMODULE hm = NULL;
            wchar_t wp[MAX_PATH] = {0};
            char nm[80] = "?";
            ULONG imgSize = 0;
            if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                   GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                   (LPCWSTR)ep->ExceptionRecord->ExceptionAddress, &hm) && hm) {
                GetModuleFileNameW(hm, wp, MAX_PATH);
                { const wchar_t* b = wcsrchr(wp, L'\\');
                  WideCharToMultiByte(CP_ACP, 0, b ? b + 1 : wp, -1, nm, sizeof(nm), NULL, NULL); }
                n += wsprintfA(buf + n, "  FAULT %s base=%p off=0x%08lX\r\n", nm, (void*)hm,
                               (unsigned long)((const char*)ep->ExceptionRecord->ExceptionAddress
                                               - (const char*)hm));
            } else {
                n += wsprintfA(buf + n, "  FAULT (unknown module) %p\r\n",
                               ep->ExceptionRecord->ExceptionAddress);
            }
        }
        if (isAV) {
            CONTEXT* c = ep->ContextRecord;
#ifdef _WIN64
            n += wsprintfA(buf + n,
                "  rip=%p rsp=%p rbp=%p rax=%p rbx=%p rcx=%p rdx=%p rsi=%p rdi=%p\r\n",
                (void*)c->Rip, (void*)c->Rsp, (void*)c->Rbp, (void*)c->Rax,
                (void*)c->Rbx, (void*)c->Rcx, (void*)c->Rdx,
                (void*)c->Rsi, (void*)c->Rdi);
#else
            n += wsprintfA(buf + n,
                "  eip=%p esp=%p ebp=%p eax=%p ebx=%p ecx=%p edx=%p esi=%p edi=%p\r\n",
                (void*)c->Eip, (void*)c->Esp, (void*)c->Ebp, (void*)c->Eax,
                (void*)c->Ebx, (void*)c->Ecx, (void*)c->Edx,
                (void*)c->Esi, (void*)c->Edi);
#endif
        }
        {
            void* frames[64];
            USHORT cnt = RtlCaptureStackBackTrace(0, 64, frames, NULL), i;
            for (i = 0; i < cnt && n < (int)sizeof(buf) - 200; i++) {
                HMODULE hm = NULL;
                wchar_t wp[MAX_PATH] = {0};
                char nm[80] = "?";
                if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                       GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                       (LPCWSTR)frames[i], &hm) && hm) {
                    GetModuleFileNameW(hm, wp, MAX_PATH);
                    { const wchar_t* b = wcsrchr(wp, L'\\');
                      WideCharToMultiByte(CP_ACP, 0, b ? b + 1 : wp, -1, nm, sizeof(nm), NULL, NULL); }
                    n += wsprintfA(buf + n, "  #%02d 0x%p  %s base=%p off=0x%08lX\r\n", i, frames[i],
                                   nm, (void*)hm,
                                   (unsigned long)((const char*)frames[i] - (const char*)hm));
                } else {
                    n += wsprintfA(buf + n, "  #%02d 0x%p  (unknown)\r\n", i, frames[i]);
                }
            }
        }
        if (isAV) {
            /* 栈扫描: 找镜像范围内的返回地址 (需要 FAULT 模块的基址与大小) */
            HMODULE hm = NULL;
            if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                   GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                   (LPCWSTR)ep->ExceptionRecord->ExceptionAddress, &hm) && hm) {
                MODULEINFO mi;
                wchar_t wp[MAX_PATH] = {0};
                char nm[80] = "?";
                GetModuleFileNameW(hm, wp, MAX_PATH);
                { const wchar_t* b = wcsrchr(wp, L'\\');
                  WideCharToMultiByte(CP_ACP, 0, b ? b + 1 : wp, -1, nm, sizeof(nm), NULL, NULL); }
                if (K32GetModuleInformation(GetCurrentProcess(), hm, &mi, sizeof(mi))) {
                    n += wsprintfA(buf + n, "  -- stack scan (%s imgsize=0x%lX) --\r\n",
                                   nm, (unsigned long)mi.SizeOfImage);
                    n = vb6_vehScanStackForImage(buf, n, (int)sizeof(buf) - 256,
                                                 ep->ContextRecord, nm,
                                                 (const char*)mi.lpBaseOfDll,
                                                 mi.SizeOfImage);
                }
            }
        }
        {
            DWORD written = 0;
            WriteFile(h, buf, (DWORD)n, &written, NULL);
            FlushFileBuffers(h);
            CloseHandle(h);
        }
    }
    return EXCEPTION_CONTINUE_SEARCH;
}

static void vb6_installCrashTrace(void) {
    static int done = 0;
    if (done) return;
    done = 1;
    if (GetEnvironmentVariableW(L"C3_CRASH_TRACE", NULL, 0) > 0) {
        AddVectoredExceptionHandler(1, vb6_heapCorruptVEH);
        SetUnhandledExceptionFilter(vb6_crashFilter);
    }
    if (GetEnvironmentVariableW(L"C3_PAGEHEAP", NULL, 0) > 0) {
        /* 全堆页堆: 损坏当场变 AV(有栈), 事后 0xC0000374(无栈)。失败静默(无权限等)。 */
        static const ULONG pgAllocs = 2;  /* HEAP_PAGE_ALLOCS */
        HANDLE heaps[256];
        DWORD nHeaps = GetProcessHeaps(256, heaps);
        DWORD i;
        for (i = 0; i < nHeaps; i++)
            HeapSetInformation(heaps[i], HeapCompatibilityInformation,
                               (void*)&pgAllocs, sizeof(pgAllocs));
    }
}

// ============================================================
// 账 #163 / #168（C29-FS-e）：自研的**按 TabIndex 跳格**。
//
// 为什么自己走：`IsDialogMessage` 只沿 **z-order**（≈创建顺序）找下一枚，而 VB6 的两张表
// 都是**按父窗内的 `TabIndex`** —— 这条实测钉死：门 #234 前后同一份夹具，OS 走出来的 `TW-seq`
// 就是发码里 `vb6_CreateControl` 的先后，与 `.frm` 写的 `TabIndex` 无关。
// 容器那一半**不用自己管**：账 #165 的根因是发码把 `WS_EX_CONTROLPARENT` 抄成了
// `WS_EX_APPWINDOW`，改对之后 OS 自己就下钻（门 #234 的 `TW-in2 / TW-deep / TW-inpic` 三条一起翻 Y）。
// 但 OS 那一张表还有两处和 VB6 不一样，只能拿回自己手里：
//   ① 次序 = z-order，不是 `TabIndex`（账 #163）；
//   ② 「谁是站」看**实时**样式位，而系统会把 `WS_TABSTOP` 自己挪到单选组里勾选那枚身上
//      （账 #168：`TabStop = 0` 的 OptionButton 照样占一站，实测 `TW-orenter=Y`）。
// 口径（与建序时用的那两张表同族，不另立）：
//   A. 同一父窗内按 (`TabIndex`, z-order) 稳定排序 —— `TabIndex` 是账 #160 发到窗口属性
//      `VB6_TabIndex` 的那个数（没存过当 0）；
//   B. 能不能进站 = **创建时**立没立 `WS_TABSTOP`（`vb6_CreateControl` 里存的 `VB6_TabStop`；
//      判断本身仍是发码那一张 `controlTabStopStyleBit` 排除表，运行期只是照读，不第二处判
//      「谁拿得到焦点」）；外加 visible / enabled —— 对话框管理器本来就跳过禁用与隐藏窗口，
//      这条口径与账 #157 同族。
//   C. 容器（带 `WS_EX_CONTROLPARENT` 的那几型）自己不进站，但走到它那一格就**就地**展开它的孩子，
//      容器嵌容器同一条规则。
// 接管三种键：VK_TAB / Shift+VK_TAB 走上面那三条口径；单选组里的 VK_UP / VK_DOWN 走
// `vb6_TabNavRadioRun`（账 #168：按 `TabIndex` 走 + 到尾回绕）。其余键 —— 非单选控件的方向键、
// Alt 助记、Enter 默认按钮 —— 照旧留给 `IsDialogMessage`，爆炸半径只管这些。
// 开关：`C3_OCX_NO_TABNAV=1` 退回"完全交给 IsDialogMessage"的旧行为，做 A/B 用（与
// `C3_OCX_NO_DLGMSG` 同一个路子，否则红的时候分不出是导航器坏还是泵坏）。
// ============================================================

#define VB6_TAB_MAX 256

static int vb6_tabIsFormWindow(HWND hwnd) {
    static const wchar_t kPrefix[] = L"VB6_Form_";   /* 见 cgen_form_prelude.inc 的注册名 */
    wchar_t cls[64];
    int i;
    if (!hwnd || !GetClassNameW(hwnd, cls, 64)) return 0;
    for (i = 0; i < 9; i++) if (cls[i] != kPrefix[i]) return 0;
    return cls[9] != 0;
}

static int vb6_tabIsContainer(HWND hwnd) {
    return (GetWindowLongW(hwnd, GWL_EXSTYLE) & WS_EX_CONTROLPARENT) != 0;
}

static int vb6_tabCanTakeFocus(HWND hwnd) {
    LONG st = GetWindowLongW(hwnd, GWL_STYLE);
    HANDLE ts;
    if ((st & (WS_CHILD | WS_VISIBLE)) != (WS_CHILD | WS_VISIBLE)) return 0;
    if (!IsWindowVisible(hwnd) || !IsWindowEnabled(hwnd)) return 0;
    /* 进站与否以**创建时**那一份为准（上面 VB6_TabStop 的注释）；不是走 vb6_CreateControl
       建起来的窗口没有这份属性，退回读实时样式位。 */
    ts = GetPropW(hwnd, L"VB6_TabStop");
    if (ts) return ((int)(INT_PTR)ts) == 1;
    return (st & WS_TABSTOP) != 0;
}

static int vb6_tabIndexOf(HWND hwnd) {
    HANDLE p = GetPropW(hwnd, L"VB6_TabIndex");
    return p ? (int)(INT_PTR)p : 0;
}

/* 同一父窗里按 `TabIndex` 稳定排序（同值保持 z-order）—— 采集器与单选组共用这一处。 */
static void vb6_tabSortByIndex(HWND *arr, int n) {
    /* ⚠ 基准值必须**存进局部变量**：第一版留的是下标（`moved = i`），而 j = i-1 那一步
       `arr[i] = arr[i-1]` 正好把基准值本身覆盖掉 ⇒ 排出来的表又乱又短，跳格走两站就停。
       插入排序的这一步是它的经典陷阱，别用下标代替值。 */
    int i, j;
    HWND tmp;
    for (i = 1; i < n; i++) {
        tmp = arr[i];
        j = i - 1;
        while (j >= 0 && vb6_tabIndexOf(arr[j]) > vb6_tabIndexOf(tmp)) {
            arr[j + 1] = arr[j];
            j--;
        }
        arr[j + 1] = tmp;
    }
}

/* 账 #168：单选组内的方向键（VK_UP / VK_DOWN）。
   OS 自己**会**在组内走，但两张表都不是 VB6 那一张 ——
     ① 它按 z-order 走（探针 `checkHEAD+ARROW`：链首 A 发 VK_DOWN 到 B）；
     ② 走到链尾就**跳出容器**（探针 `checkTAIL+ARROW` 读 `SEQ=B,T1,T2,A,B,...`，
        与产品 `AK-pre=optA / down=cmdTop1` 逐字对上 —— 产品里 optB 先创建、optA 是链尾）。
   VB6 的口径 = 同容器、同型单选钮按 `TabIndex` 走，到尾回绕。
   改勾选不自己动手：给目标发 `BM_CLICK`（winuser.h:11348）—— 一发同时拿到
   「单选组自动取消别人」与 `BN_CLICKED`（= VB6 里用户改选项时那声 `_Click`）。
   判据把 optA / optB 两个 Value 一起读，钉的就是这一条：两边都 Y = 没人被自动取消。
   常量只用 SDK 符号名（BS_TYPEMASK=0x0F、BS_AUTORADIOBUTTON=0x09，winuser.h:11303/11300），
   不手抄数字 —— 上一格把 WS_EX_CONTROLPARENT 抄成 WS_EX_APPWINDOW 就是这一课。 */
static int vb6_tabIsRadio(HWND hwnd) {
    return (GetWindowLongW(hwnd, GWL_STYLE) & BS_TYPEMASK) == BS_AUTORADIOBUTTON;
}

static int vb6_TabNavRadioRun(HWND cur, int down) {
    HWND parent, h, grp[VB6_TAB_MAX];
    int n = 0, i, at = -1, nxt;
    if (!cur || !vb6_tabIsRadio(cur)) return 0;
    parent = GetParent(cur);
    if (!parent) return 0;
    for (h = GetWindow(parent, GW_CHILD); h != NULL; h = GetWindow(h, GW_HWNDNEXT)) {
        if (n >= VB6_TAB_MAX) break;
        if (!vb6_tabIsRadio(h)) continue;
        if (!IsWindowVisible(h) || !IsWindowEnabled(h)) continue;
        grp[n++] = h;
    }
    vb6_tabSortByIndex(grp, n);
    for (i = 0; i < n; i++) { if (grp[i] == cur) { at = i; break; } }
    if (at < 0) return 0;
    if (n == 1) return 1;   /* 组里只有这一枚：VB6 也是原地不动，但这声不能漏给 OS（它会跳出组走下一站） */
    nxt = (at + (down ? 1 : -1) + n) % n;
    SetFocus(grp[nxt]);
    SendMessageW(grp[nxt], BM_CLICK, 0, 0);
    return 1;
}

/* 把 parent 这一层（含容器里的孩子）按上面那三条口径依次收集进 out。 */
static void vb6_tabCollect(HWND parent, HWND *out, int *n, int depth) {
    HWND kid[VB6_TAB_MAX];
    int nk = 0, i;
    HWND h;
    if (*n >= VB6_TAB_MAX || depth > 8) return;
    for (h = GetWindow(parent, GW_CHILD); h != NULL; h = GetWindow(h, GW_HWNDNEXT)) {
        if (nk >= VB6_TAB_MAX) break;
        kid[nk++] = h;
    }
    vb6_tabSortByIndex(kid, nk);   /* 同值保持 z-order */
    for (i = 0; i < nk; i++) {
        h = kid[i];
        if (GetEnvironmentVariableW(L"C3_TABNAV_TRACE", NULL, 0) > 0) {
            fprintf(stderr, "  [collect d=%d] hwnd=%p style=%lx ex=%lx focus=%d cont=%d\n", depth, (void*)h,
                    (unsigned long)GetWindowLongW(h, GWL_STYLE),
                    (unsigned long)GetWindowLongW(h, GWL_EXSTYLE),
                    vb6_tabCanTakeFocus(h), vb6_tabIsContainer(h));
        }
        if (vb6_tabCanTakeFocus(h)) {
            out[(*n)++] = h;
            if (*n >= VB6_TAB_MAX) return;
        }
        if (vb6_tabIsContainer(h)) vb6_tabCollect(h, out, n, depth + 1);
    }
}

int vb6_Form_MoveTabFocus(void* hwndForm, int forward) {
    HWND form = (HWND)hwndForm, cur, list[VB6_TAB_MAX];
    int n = 0, i, at = -1, nxt;
    if (!form || !IsWindow(form)) return 0;
    vb6_tabCollect(form, list, &n, 0);
    if (n <= 0) return 0;
    if (GetEnvironmentVariableW(L"C3_TABNAV_TRACE", NULL, 0) > 0) {   /* 排障用，不参与判定 */
        HWND dbgCur = GetFocus();
        fprintf(stderr, "[C3_TABNAV] form=%p cur=%p fwd=%d n=%d\n", (void*)form, (void*)dbgCur, forward, n);
        for (i = 0; i < n; i++)
            fprintf(stderr, "  #%d hwnd=%p idx=%d style=%lx ex=%lx\n", i, (void*)list[i],
                    vb6_tabIndexOf(list[i]),
                    (unsigned long)GetWindowLongW(list[i], GWL_STYLE),
                    (unsigned long)GetWindowLongW(list[i], GWL_EXSTYLE));
    }
    cur = GetFocus();
    for (i = 0; i < n; i++) { if (list[i] == cur) { at = i; break; } }
    if (at < 0) {                       /* 焦点在窗体身上/在没进站的容器里：VB6 也是从首末枚开始 */
        SetFocus(forward ? list[0] : list[n - 1]);
        return 1;
    }
    if (n == 1) return 1;               /* 只有这一枚可聚焦，原地不动 */
    nxt = (at + (forward ? 1 : -1) + n) % n;
    SetFocus(list[nxt]);
    return 1;
}

/* 返回 1 = 这条按键已被接管，调用方不要再把它交给 IsDialogMessage / 不要再派发。 */
static int vb6_TabNavKey(const MSG *msg) {
    HWND focus, root;
    int vk;
    if (msg->message != WM_KEYDOWN) return 0;
    vk = (int)msg->wParam;
    if (vk != VK_TAB && vk != VK_UP && vk != VK_DOWN) return 0;
    focus = GetFocus();
    root = GetAncestor(focus ? focus : msg->hwnd, GA_ROOT);
    if (!vb6_tabIsFormWindow(root)) return 0;
    if (vk == VK_UP || vk == VK_DOWN) {
        /* 只管单选组；列表框、滚动条那一族的方向键照旧交给 IsDialogMessage。 */
        return focus ? vb6_TabNavRadioRun(focus, vk == VK_DOWN) : 0;
    }
    return vb6_Form_MoveTabFocus(root, (GetKeyState(VK_SHIFT) & 0x8000) == 0);
}

// ============================================================
// 账 #157: 编译器算好的"这枚窗体显示时该把焦点交给谁"（VB6 = TabIndex 最小那枚拿得到焦点的
// 控件，不是创建顺序 —— `.frm` 里控件的书写顺序与 TabIndex 常常相反）。存在窗体句柄上，
// **应用一次就销掉**：之后再 Show 这枚窗体，焦点该回到用户停下的地方，而不是每次都被抢回首枚 tabstop。
void vb6_Form_SetInitialFocus(void* hwnd, void* target) {
    if (!hwnd || !target) return;
    SetPropW((HWND)hwnd, L"VB6_InitFocus", (HANDLE)target);
}

static void vb6_ApplyInitialFocus(HWND hwnd) {
    HWND t = (HWND)RemovePropW(hwnd, L"VB6_InitFocus");
    if (t && IsWindow(t)) SetFocus(t);
}

void vb6_ShowForm(void* hwnd, int modal) {
    vb6_installCrashTrace();
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
        wchar_t cap[160] = {0};
        GetWindowTextW((HWND)hwnd, cap, 159);
        fprintf(stderr, "[C3_MODAL] ShowForm hwnd=%p modal=%d cap='%ls'\n", hwnd, modal, cap);
    }
    if (!hwnd) return;

    // Fix 115: 恢复 VB6 的 "先 Form_Load, 后 Show" 顺序。
    // 编译器把 Form_Load 用 PostMessageW(hwnd, 0x7FF0, 0, 0) 延迟到消息队列
    // (见 cgen_form_wndproc_create.inc 的 WM_CREATE 处理), 而这里的
    // ShowWindow/UpdateWindow 会在队列消息派发**之前**强制首次 WM_PAINT。
    // 于是 UserControl 的 Draw 会在"Form_Load 尚未添加数据系列"的状态下执行,
    // 对空数组取 m_Serie(0) → 空指针崩溃 (Charts 2020 ucChartArea 在
    // LegendAlign=LA_TOP 时必经该分支)。先把挂起的延迟 Form_Load 派发掉。
    // Fix 190: A 版消息 API 会把 Unicode 消息降级成 ANSI (WM_CHAR/文本), 全部换 W。
    {
        const UINT kDeferredFormLoad = 0x7FF0;   // 编译器生成的"延迟 Form_Load"消息
        MSG msg;
        while (PeekMessageW(&msg, (HWND)hwnd, kDeferredFormLoad, kDeferredFormLoad, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    ShowWindow((HWND)hwnd, SW_SHOWDEFAULT);
    UpdateWindow((HWND)hwnd);

    // Fix 137: VB6 启动的窗体会出现在最前并获得焦点; 我们只 ShowWindow 的话
    // 窗口常落在已有窗口后面 (双击 exe 时尤甚 —— 没有前台权限继承)。
    // 显式提到顶层并请求前台 (SetForegroundWindow 受系统限制时 BringWindowToTop
    // 至少保证同 Z 序应用内最前)。
    BringWindowToTop((HWND)hwnd);
    if (SetForegroundWindow((HWND)hwnd) == 0) {
        /* 前台锁: 挂到当前前台线程的输入队列再试一次 (经典 workaround) */
        HWND fg = GetForegroundWindow();
        DWORD fgTid = fg ? GetWindowThreadProcessId(fg, NULL) : 0;
        DWORD myTid = GetCurrentThreadId();
        if (fgTid && fgTid != myTid &&
            AttachThreadInput(myTid, fgTid, TRUE)) {
            BringWindowToTop((HWND)hwnd);
            SetForegroundWindow((HWND)hwnd);
            SetFocus((HWND)hwnd);
            AttachThreadInput(myTid, fgTid, FALSE);
        }
    }
    SetActiveWindow((HWND)hwnd);
    // 账 #157: VB6 在窗体激活之后把焦点交给第一枚 tabstop。这一步必须在激活之后 ——
    // 实测 (029 的 `s-0 第 2 条改口径`)：在 Form_Activate 里 SetFocus 会被随后的激活流程收回。
    vb6_ApplyInitialFocus((HWND)hwnd);

    if (modal) {
        int traceModal = (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0);
        // 模态窗体: 禁用所有者, 进入本地消息循环
        HWND owner = GetWindow((HWND)hwnd, GW_OWNER);
        if (owner) {
            g_modalOwner = owner;
            EnableWindow(owner, FALSE);
        }

        // 本地消息循环 (直到窗体被销毁)
        MSG msg;
        int traceMsg = traceModal;
        int msgCount = 0;
        /* Fix 144b: IsDialogMessageA 会吞掉非对话框窗口的键盘/命令消息, 可用
         * C3_OCX_NO_DLGMSG=1 关闭 (对照实验/排障). */
        int useDlgMsg = (GetEnvironmentVariableW(L"C3_OCX_NO_DLGMSG", NULL, 0) <= 0);
        /* 账 #163/#168：模态这条也一样接管 VK_TAB（与主泵同一个开关）。 */
        int useTabNav = (GetEnvironmentVariableW(L"C3_OCX_NO_TABNAV", NULL, 0) <= 0);
        while (IsWindow((HWND)hwnd) && GetMessage(&msg, NULL, 0, 0)) {
            if (traceMsg && msgCount < 80) {
                wchar_t cap[128] = {0};
                GetWindowTextW((HWND)hwnd, cap, 128);
                fprintf(stderr, "[C3_MODAL] msg[%d] 0x%04X hwnd=%p wp=%p lp=%p | ownerWin alive=%d cap='%ls'\n",
                        msgCount, msg.message, (void*)msg.hwnd,
                        (void*)msg.wParam, (void*)msg.lParam,
                        IsWindow((HWND)hwnd) ? 1 : 0, cap);
            }
            msgCount++;
            // 账 #163/#168：跳格自己按 TabIndex 走，别让对话框管理器按 z-order 认
            if (useTabNav && vb6_TabNavKey(&msg)) continue;
            // P24-Timer: WM_TIMER现在由WndProc分发, 模态循环不再拦截
            // 模态Tab键导航 (IsDialogMessage处理对话框键盘导航)
            if (!useDlgMsg || !IsDialogMessageW((HWND)hwnd, &msg)) {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
            if (msg.message == WM_QUIT && traceModal)
                fprintf(stderr, "[C3_MODAL] 收到 WM_QUIT, 模态循环结束 (hwnd=%p)\n", hwnd);
        }
        if (traceModal)
            fprintf(stderr, "[C3_MODAL] 模态循环退出: IsWindow=%d (hwnd=%p, owner=%p)\n",
                    IsWindow((HWND)hwnd) ? 1 : 0, hwnd, owner);

        // 恢复所有者窗口
        if (g_modalOwner) {
            EnableWindow(g_modalOwner, TRUE);
            SetActiveWindow(g_modalOwner);
            g_modalOwner = NULL;
        }
    }
}

void vb6_UnloadForm(void* hwnd) {
    if (!hwnd) return;
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_MODAL] UnloadForm hwnd=%p\n", hwnd);
    // P20-43: **不能直接 DestroyWindow** —— 那会跳过 WM_CLOSE, 于是 VB 代码里的
    // `Unload Me` 既不触发 Form_QueryUnload 也不触发 Form_Unload (实测 FrmEvents
    // 夹具 EV21/EV22 整个消失)。改发 WM_CLOSE: 窗体的 WM_CLOSE 分支里已有完整的
    // "QueryUnload(可取消) → Unload → DestroyWindow" 链, 语义与 VB6 一致
    // (UnloadMode=0 vbFormControlMenu 那条路径)。不会递归: WM_CLOSE 分支只会
    // DestroyWindow, 不会再发 WM_CLOSE。
    SendMessageW((HWND)hwnd, WM_CLOSE, 0, 0);
}

// M22-Issue6: 窗体表面Print
// VB6的"Print expr"语句在窗体表面绘制文本
// 维护CurrentX/CurrentY用于定位下一次输出
void vb6_Form_Print(void* hwnd, void* bstrText) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    BSTR text = (BSTR)bstrText;
    
    // Get CurrentX/CurrentY from window properties (stored as pixels)
    float currentX = vb6_GetCurrentX(hwnd);
    float currentY = vb6_GetCurrentY(hwnd);
    
    HDC hdc = GetDC(hw);
    if (!hdc) return;
    
    // Set text color and background mode (transparent for form printing)
    SetBkMode(hdc, TRANSPARENT);
    
    int len = text ? (int)SysStringLen(text) : 0;
    if (len > 0) {
        // Calculate text size for advancing CurrentX
        SIZE size;
        TEXTMETRICW tm;
        GetTextExtentPoint32W(hdc, text, len, &size);
        GetTextMetricsW(hdc, &tm);
        
        // Draw text at CurrentX, CurrentY
        TextOutW(hdc, (int)currentX, (int)currentY, text, len);
        
        // VB6 behavior: Print automatically advances to next line (newline)
        // CurrentY += line height, CurrentX reset to 0
        vb6_SetCurrentY(hwnd, currentY + (float)tm.tmHeight);
        vb6_SetCurrentX(hwnd, 0.0f);
    } else {
        // Empty Print = newline: advance CurrentY by font height, reset CurrentX
        TEXTMETRICW tm;
        GetTextMetricsW(hdc, &tm);
        vb6_SetCurrentY(hwnd, currentY + (float)tm.tmHeight);
        vb6_SetCurrentX(hwnd, 0.0f);
    }
    
    ReleaseDC(hw, hdc);
}

// ============================================================
// Form_Unload回调
// ============================================================

void vb6_SetFormUnloadCallback(void* callback) {
    g_formUnloadCb = (vb6_FormUnloadCallback)callback;
}

int vb6_QueryFormUnload(void) {
    if (g_formUnloadCb) {
        return g_formUnloadCb();
    }
    return 0;  // 无回调=允许关闭
}
