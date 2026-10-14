#pragma once
// vb6forms_prop.h - 控件属性读写 (P7.5)：Text/Caption、Value、Visible、Enabled、位置大小、字体、前景背景色
// 由 vb6forms.h 伞头按固定顺序 include，不要单独使用
// 内容 = 拆分前 vb6forms.h 第 142~203 行，逐行未改

#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 控件属性读写 (P7.5)
// ============================================================

// Text/Caption属性 (TextBox/Label/CommandButton/Form)
// 返回BSTR (调用者负责vb6_SysFreeString释放, 或直接传给函数)
// Fix 150: 返回类型必须是 wchar_t* 而非 void* —— 代码生成的
// vb6_ComPackValue 走 C11 _Generic, void* 会落入 default 分支被当成
// 对象指针打包成 VT_DISPATCH, OLEAUT32 解引用假接口指针即崩
// (NewTab 主题下拉选择 0xC0000005 实测)。wchar_t* 才会命中
// VariantString 分支打包为 VT_BSTR。
wchar_t* vb6_GetControlText(void* hwnd);
void vb6_SetControlText(void* hwnd, void* bstr);

// Value属性 (CheckBox/OptionButton: 0=Unchecked, 1=Checked, 2=Grayed)
int vb6_GetCheckValue(void* hwnd);
void vb6_SetCheckValue(void* hwnd, int value);
// 账 #128-b: OptionButton 单开一对 —— VB6 那边它的 Value 是 Boolean(读回 -1/0)，
// 而 CheckBox 的是三态 Integer；写侧必须把非 0 折回 BST_CHECKED(1)，因为 BM_SETCHECK 不吃 -1。
int vb6_GetOptionValue(void* hwnd);
void vb6_SetOptionValue(void* hwnd, int value);

// Visible属性
int vb6_GetControlVisible(void* hwnd);
void vb6_SetControlVisible(void* hwnd, int visible);

// Enabled属性
int vb6_GetControlEnabled(void* hwnd);
void vb6_SetControlEnabled(void* hwnd, int enabled);

// C29-SL-l（账 #143）: VB6 的 `控件.SetFocus`。之前这一形从没登记过，两形（带括号与不带括号）
// 都落进 vb6_ComCall(裸 HWND, L"SetFocus", NULL, 0) —— 原生控件槽里是句柄不是 IDispatch，
// 于是编得过、链接得过、跑起来一声不响。RTL 就一句 SetFocus(hwnd)（焦点属于线程输入队列，
// 与窗口可见/激活无关，所以无头跑里也真能拿到 —— C29-SL-g 的 SimStdEvent kind=4 走同一条路，
// 实测会发 WM_SETFOCUS）。拿不到焦点时原生就是回 NULL 什么都不做，本项目不伪造、不重试。
void vb6_SetControlFocus(void* hwnd);

// 账 #171（C29-FS-h）：单选钮那一条 `BN_CLICKED` 该不该升成 VB6 的 `Click` —— 口径只留这一处。
// 实测（探针 `.build/ck171c`，两头证人：VB 侧 `Value` 与 OS 侧 `BM_GETCHECK`）：`BS_AUTORADIOBUTTON`
// 在**焦点进入**时也会替父窗发一条 `BN_CLICKED`，而那一刻它自己**并没有被勾上**；真点击（含 `BM_CLICK`）
// 那条到达时它已经是勾上的。VB6 里 `Click` 只在「用户选了这一枚」时发生 ⇒ 判据 = 通知的来路自己勾上没有。
// 与 `btnfocus` 夹具的 `BF-noclick` 同一条口径（那一格早已钉「焦点移动不许算成点击」，只是单选钮这一型
// 连通知号都是 `BN_CLICKED`，光按 `code` 筛不掉）。只管 `BS_TYPEMASK == BS_AUTORADIOBUTTON`：
// 复选框取消勾选是一次正经 `Click`，不能跟着筛。返回 0 = 这条通知不许升成 `Click`。
int vb6_RadioClickCounts(void* hwndFrom);

// P11.8: Position/Size attributes (all visible controls, in pixels)
// VB6 uses twips internally, but Win32 uses pixels; RTL handles conversion
int vb6_GetControlLeft(void* hwnd);
void vb6_SetControlLeft(void* hwnd, int left);
int vb6_GetControlTop(void* hwnd);
void vb6_SetControlTop(void* hwnd, int top);
int vb6_GetControlWidth(void* hwnd);
void vb6_SetControlWidth(void* hwnd, int width);
int vb6_GetControlHeight(void* hwnd);
void vb6_SetControlHeight(void* hwnd, int height);

// 账 #230: 控件几何的 VB 侧读数（详见 vb6forms_ctrl.c 那段注释）。四个 getter 与 Move
// 都问这一处，**没有第二份存储**。写侧三个来路：属性赋值、Move、建窗（设计值）。
enum { VB6_GEOM_LEFT = 0, VB6_GEOM_TOP = 1, VB6_GEOM_WIDTH = 2, VB6_GEOM_HEIGHT = 3 };
void vb6_GeomCacheWrite(void* hwnd, int slot, int value, int32_t mode);
int  vb6_GeomCacheRead(void* hwnd, int slot, int actualPx, int32_t mode, int fallback);
// Fix 162a: obj.Move l[,t[,w[,h]]] —— twips 进; mask 位 1=L 2=T 4=W 8=H, 缺失位不变
void vb6_ControlMove(void* hwnd, double l, double t, double w, double h, int mask);

// P11.8: hWnd attribute (read-only, returns the Win32 HWND as pointer)
void* vb6_GetControlHwnd(void* hwnd);

// P13.1: Font properties (all visible controls with text)
// FontName: returns BSTR (caller responsible for SysFreeString)
void* vb6_GetControlFontName(void* hwnd);
void vb6_SetControlFontName(void* hwnd, void* bstrName);
// C29-SL-q（账 #154）: 设计期那一条 —— 字体名是生成码里的 C 字面量，不是 BSTR。
void vb6_SetControlFontNameW(void* hwnd, const wchar_t* name);
// FontSize: returns VB6 Single (points) as float
float vb6_GetControlFontSize(void* hwnd);
void vb6_SetControlFontSize(void* hwnd, float sizePt);
// C29-SL-p 判据证人（不是 VB6 属性）：窗口真在用的字体像素高度。
// 字号那条判据只问自存的数就是自洽假绿，得同时问窗口一次（#148 那条教训）。
int vb6_ControlFontPixelHeight(void* hwnd);
// FontBold: VB6 True=-1, False=0
int vb6_GetControlFontBold(void* hwnd);
void vb6_SetControlFontBold(void* hwnd, int bold);
// FontItalic
int vb6_GetControlFontItalic(void* hwnd);
void vb6_SetControlFontItalic(void* hwnd, int italic);
// FontUnderline
int vb6_GetControlFontUnderline(void* hwnd);
void vb6_SetControlFontUnderline(void* hwnd, int underline);
// FontStrikethrough
int vb6_GetControlFontStrikethrough(void* hwnd);
void vb6_SetControlFontStrikethrough(void* hwnd, int strike);

// P13.2: ForeColor/BackColor (all visible controls)
// Returns OLE color (VB6 Long), 0x00BBGGRR format
int vb6_GetControlForeColor(void* hwnd);
void vb6_SetControlForeColor(void* hwnd, int color);
int vb6_GetControlBackColor(void* hwnd);
void vb6_SetControlBackColor(void* hwnd, int color);

// Fix 187: WM_CTLCOLORSTATIC 统一应用 — 子控件 Set 过 VB6_BackColor 时
// 应用 Fore/Back 色并返回背景刷; 未设色返回 0 (调用方走 DefWindowProcW)。
// 主窗体 WndProc (生成代码) 与 SSTab 容器子类共用。
LRESULT vb6_ApplyCtlColorStatic(HDC hdc, HWND child);

// Fix 162f-extlist: WM_CTLCOLORBTN 统一答复 (Button 类子控件要背景刷) ——
// CheckBox/OptionButton 返回 HOLLOW_BRUSH (透明); Frame(BS_GROUPBOX) 返回
// **父窗底色的实心刷** (经典 groupbox 标题的 FillRect 底, 空刷会露出自身白底)。
// 同族调用方: 生成窗体 WndProc 的 WM_CTLCOLORBTN 分支 + Frame 子类化过程。
LRESULT vb6_CtlColorBtnBrush(HWND child, HWND parent);


#ifdef __cplusplus
}
#endif
