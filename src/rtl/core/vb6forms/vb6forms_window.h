#pragma once
// vb6forms_window.h - 窗体框架 / 控件创建 / Timer管理 / 消息循环 / 窗体事件桥接 / 窗体工具函数 / Form_Unload回调
// 由 vb6forms.h 伞头按固定顺序 include，不要单独使用
// 内容 = 拆分前 vb6forms.h 第 15~141 行，逐行未改

#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 窗体框架
// ============================================================

// 注册窗体窗口类 (内部调用RegisterClassEx)
// className: 窗口类名 (如 "VB6_Form1")
// wndProc: 窗口过程
// hInstance: 应用实例
// iconResId: 图标资源ID (0=默认)
// 返回: 0=成功, -1=失败
int vb6_RegisterFormClass(const char* className, void* wndProc, void* hInstance, int iconResId);
// czUI fix: 窗体级 .frm BackColor 支持 (bg<0 = 未指定走 VB6 默认)
int vb6_RegisterFormClassBg(const char* className, void* wndProc, void* hInstance, int iconResId, int backColor);
// 查询已登记窗体类的背景色 (无登记返回 -1); 供 UserControl Ambient.BackColor 使用
int vb6_Forms_QueryClassBg(const char* className);

// 创建窗体窗口
// className: 已注册的窗口类名
// formName: 窗口标题
// x, y: 窗口位置 (CW_USEDEFAULT表示系统默认)
// width, height: 客户区大小 (缇, 1缇=1/15像素)
// hInstance: 应用实例
// userData: 传递给WM_CREATE的用户数据指针
// 返回: 窗口句柄 (HWND)
void* vb6_CreateFormWindow(const char* className, const char* formName,
    int x, int y, int width, int height, void* hInstance, void* userData);
// czUI fix: 带 BorderStyle 的窗体创建 (0=None 1=FixedSingle 2=Sizable 3=FixedDialog 4/5=ToolWindow)
void* vb6_CreateFormWindowB(const char* className, const char* formName,
    int x, int y, int width, int height, void* hInstance, void* userData,
    int borderStyle);

// VB6缇(Twip)转像素
// VB6坐标单位: 1英寸=1440缇, 1像素=15缇 (96DPI标准)
int vb6_TwipToX(int twips);
int vb6_TwipToY(int twips);
// Fix 184: 唯一 DPI 源 + 反向换算 (像素 -> 缇)。RTL 内任何 px/缇 转换都必须
// 走这四个入口，禁止再写死 15。
int vb6_DpiX(void);
int vb6_HimetricToPxOnDc(int hm, int dpi);  // 账 #298: himetric -> 像素那**一个公式**的唯一住处 (DPI 由调用方递)
int vb6_HimetricToPxX(int hm);   // = OnDc(hm, vb6_DpiX())  —— 屏 DPI 那一档
int vb6_HimetricToPxY(int hm);   // = OnDc(hm, vb6_DpiY())
float vb6_HimetricToPxF(float hm, float dpi);   // OLE 容器坐标那一路 (POINTF, 不许量化成整数)
float vb6_PxToHimetricF(float px, float dpi);
// 账 #206 第二格 + 账 #298: himetric <-> 缇 (与 DPI 无关的纯单位换算), 唯一那一处。
long vb6_TwipsToHimetric(long tw);
long vb6_HimetricToTwips(long hm);
int vb6_DpiY(void);
int vb6_XToTwipX(int px);
int vb6_YToTwipY(int px);

// 账 #175: 控件坐标的单位 = 所在容器的 ScaleMode (窗体或 UserControl 的 .ctl/.frm
// 声明值)。这一对是**唯一**的像素<->容器单位换算入口; 缇 (mode 1) 那一档与上面的
// vb6_TwipToX/XToTwipX 逐字节等价。RTL 里任何"读/写控件几何"的点位都走这一对,
// 禁止再默认缇。
double vb6_ScalePxToUser(double px, int32_t mode, int vert);
int    vb6_ScaleUserToPx(double user, int32_t mode, int vert);
// 1 设备像素 = 多少该 ScaleMode 单位 (vb6rtl_com.c 的 ScaleX/ScaleY 也读这一张表,
// 账 #177: 全 RTL 只留这一份单位表)。
double vb6_ScaleUnitsPerPx(int32_t mode, int vert);
// 目标窗口的容器 ScaleMode: 容器是 UserControl 宿主 → 它的 .ctl ScaleMode;
// 否则读窗体的 VB6_ScaleMode 属性 (缺省 1=缇)。
int32_t vb6_ContainerScaleMode(void* hwndParent);
// 窗口自身的 ScaleMode (ScaleWidth/ScaleHeight 那一族读法的单位)。
int32_t vb6_WindowScaleModeSelf(void* hwnd);
// 该 HWND 是 UserControl 宿主时返回它 .ctl 声明的 ScaleMode, 否则 0。
int32_t vb6_UC_WindowScaleMode(const void* hwnd);

// ============================================================
// 控件创建
// ============================================================

// 创建子控件 (通用)
// win32Class: Win32窗口类名 (BUTTON, EDIT, STATIC, etc.)
// controlName: 控件名 (用于SetWindowText和WM_COMMAND标识)
// style: 窗口样式 (WS_CHILD | WS_VISIBLE | ...)
// exStyle: 扩展样式
// x, y, width, height: 位置和大小 (缇)
// id: 控件ID (用于WM_COMMAND)
// hParent: 父窗口句柄
// hInstance: 应用实例
// 返回: 控件窗口句柄
void* vb6_CreateControl(const char* win32Class, const char* controlName,
    long style, long exStyle,
    int x, int y, int width, int height,
    int id, void* hParent, void* hInstance);

// 控件ID分配器 (每个窗体独立ID空间)
int vb6_NextControlId(void);

// 重置控件ID计数器 (新窗体开始时调用)
void vb6_ResetControlId(void);

// ============================================================
// Timer管理
// ============================================================

// 设置定时器 (VB6 Timer控件底层实现)
// interval: 间隔毫秒 (VB6 Interval属性)
// callback: 定时器回调函数 (Timer_Timer事件)
// 返回: 定时器ID (用于vb6_KillTimer)
int vb6_SetTimer(void* hwnd, int interval, void* callback);

// C29-T: 运行期可开/停/改周期的计时器三件套。owner = 派发窗（窗体），key = Timer 控件
// 自己的不可见句柄 —— 没有 key 就没有"运行期找得回这一格"，那正是改之前的症状。
// 精度：winmm timeSetEvent（ms 级，取不到则退回 SetTimer）。
void vb6_TimerAttach(void* owner, void* key, int period, void* callback, int enabled);
void vb6_TimerSetEnabled(void* key, int enabled);
void vb6_TimerSetPeriod(void* key, int period);
// 销毁定时器
void vb6_KillTimer(int timerId);

// P24-Timer: WndProc dispatch for WM_TIMER (generated WndProc calls this)
void vb6_DispatchTimer(int timerId);

// 账 #157: 窗体显示时该把焦点交给谁（VB6 = TabIndex 最小那枚拿得到焦点的控件）。
// 发码期算好、WM_CREATE 里存一次，vb6_ShowForm 在激活之后应用并销掉。
void vb6_Form_SetInitialFocus(void* hwnd, void* target);

// ============================================================
// 消息循环
// ============================================================

// 标准VB6消息循环 (GetMessage + TranslateMessage + DispatchMessage)
// 同时处理WM_TIMER回调分发
// 返回: WM_QUIT的wParam值
int vb6_MessageLoop(void);
// Fix 167: Sub Main 返回后是否应继续驻留 (本线程仍有可见窗口)
int vb6_AnyThreadWindowVisible(void);

// DoEvents — 处理消息队列中的待处理消息
// 包括WM_TIMER回调分发
// 返回: 处理的消息数
int vb6_DoEvents(void);

// ============================================================
// 窗体事件桥接
// ============================================================

// 设置窗体用户数据 (将VB6窗体对象指针存入GWLP_USERDATA)
void vb6_SetFormUserData(void* hwnd, void* userData);

// 获取窗体用户数据
void* vb6_GetFormUserData(void* hwnd);

// ============================================================
// 窗体工具函数
// ============================================================

// 获取应用程序实例句柄
void* vb6_GetAppInstance(void);

// 设置应用程序实例句柄 (WinMain中调用)
void vb6_SetAppInstance(void* hInstance);

// 显示窗体 (vbModeless=0, vbModal=1)
// 模态时: 禁用父窗口, 进入本地消息循环直到窗体关闭
void vb6_ShowForm(void* hwnd, int modal);

// Fix <vbeclipse>: 只 Load 不 Show —— 抽干延迟 Form_Load, 窗体保持隐藏
void vb6_LoadForm(void* hwnd);

// Fix <vbeclipse> rev28: 窗体 Form_Resize 的**延后触发**口 (WM_SIZE 里排, 消息循环里跑)。
//
// 为什么需要: WM_SIZE 是 SetWindowPos/MoveWindow 的**同步** SendMessage, 停靠布局里
// 一整串嵌套 Move (ucPerspective → ucFolder → ViewArea → 视图窗体) 全在同一个调用栈里
// 跑完才返回。等它返回时, 各控件才拿到**最终**尺寸。所以 Form_Resize 必须延到本轮
// 布局收尾后再跑, 否则它按**同步中间态**去摆子控件。
// 实证 (play78 --arch x86, 探针): 直接在 WM_SIZE 里跑, frmViewViews 的
// `tvwViews.Move 0, 0, ScaleWidth, ScaleHeight` 拿到的是 0x0 ⇒ 树控件被摆成 0x0。
// 这与 UserControl 侧的 vb6_UC_QueueDesignResize (rev23) 是**同一类问题的两个面**。
//
// 与 UC 侧同款做法: PostMessage 一个 WM_APP 消息 + 窗口属性去重 (rec/结构体都不动,
// 避开"跨边界布局式初始化的字段顺序"那个坑)。cgen 侧在窗体 WndProc 里发一条
// `case VB6_FORM_FR_MSG:` 调 vb6_DrainFormResize(hwnd) 即可。
#define VB6_FORM_FR_MSG  (WM_APP + 0x61)
#define VB6_FORM_FR_PROP L"C3_FORM_FR_PENDING"
int32_t vb6_QueueFormResize(void* hwnd);
void   vb6_DrainFormResize(void* hwnd);

// Fix <vbeclipse> rev29: 窗体 Form_Resize 的**直调**通道。
// 存的是 cgen 生成的 `vb6_<Form>_Form_Resize` 的地址 (通过窗口属性关联到 HWND),
// 供 RTL 在 vb6_ControlMove 收口处同步调用 —— 不依赖 WM_SIZE, 也就绕开了
// "启动期不在消息循环里 ⇒ PostMessage 无人 Dispatch" 这个死结。
//
// 为什么不走排队 (rev29 前一版实测失败, 留档):
//   停靠视图窗体一生只收到 1 次 WM_SIZE(0x0), rev28 的 case VB6_FORM_FR_MSG
//   永远等不到; 改在 vb6_ControlMove 里 PostMessage 也不行 —— 排队那一刻还在
//   Form_Load 的同步栈里, 消息没人 Dispatch, 且实测堆损坏 0xC0000374。
// 直调为何安全: SetWindowPos 已完成, Form_Resize 读到的是**终值**; 它内部再 Move
// 子控件时由 sizeChanged 判据收敛。
#define VB6_FORM_RESIZE_PROP L"C3_FORM_RESIZE_FN"
void vb6_RegisterFormResize(void* hwnd, void* fn);
int32_t vb6_InvokeFormResize(void* hwnd);

// 卸载窗体
void vb6_UnloadForm(void* hwnd);

// M22-Issue6: 窗体表面Print (VB6的"Print expr"语句)
// hwnd: 窗体HWND, text: BSTR要输出的文本
// 使用TextOutW在窗体HDC上绘制, 维护CurrentX/CurrentY位置
void vb6_Form_Print(void* hwnd, void* bstrText);

// ============================================================
// Form_Unload回调
// ============================================================

// 设置Form_Unload回调 (WM_CLOSE时查询是否允许关闭)
// callback: 返回0=允许关闭, 返回非0=取消关闭
void vb6_SetFormUnloadCallback(void* callback);

// 查询Form_Unload (由WndProc的WM_CLOSE调用)
// 返回: 0=允许关闭, 非0=取消关闭
int vb6_QueryFormUnload(void);


#ifdef __cplusplus
}
#endif
