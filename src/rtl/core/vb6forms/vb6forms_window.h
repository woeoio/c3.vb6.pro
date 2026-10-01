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
int vb6_DpiY(void);
int vb6_XToTwipX(int px);
int vb6_YToTwipY(int px);

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
