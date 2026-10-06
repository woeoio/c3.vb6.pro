// vb6forms_ctrl.c - vb6forms 模块拆分: 控件通用属性 + 字体 + 颜色
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
// 控件属性读写 (P7.5)
// ============================================================

wchar_t* vb6_GetControlText(void* hwnd) {
    if (!hwnd) return NULL;
    HWND h = (HWND)hwnd;
    int len = GetWindowTextLengthW(h);
    if (len <= 0) {
        // 返回空BSTR
        WCHAR empty[] = {0};
        return SysAllocString(empty);
    }
    WCHAR* buf = (WCHAR*)malloc((len + 1) * sizeof(WCHAR));
    if (!buf) return NULL;
    GetWindowTextW(h, buf, len + 1);
    BSTR bstr = SysAllocString(buf);
    free(buf);
    return (void*)bstr;
}

void vb6_SetControlText(void* hwnd, void* bstr) {
    if (!hwnd) return;
    // 支持BSTR和char*两种输入
    if (bstr) {
        // 尝试作为BSTR处理 (VB6字符串)
        BSTR bs = (BSTR)bstr;
        SetWindowTextW((HWND)hwnd, bs);
    }
}

int vb6_GetCheckValue(void* hwnd) {
    if (!hwnd) return 0;
    LRESULT state = SendMessageW((HWND)hwnd, BM_GETCHECK, 0, 0);
    return (int)state;  // BST_UNCHECKED=0, BST_CHECKED=1, BST_INDETERMINATE=2
}

void vb6_SetCheckValue(void* hwnd, int value) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, BM_SETCHECK, (WPARAM)value, 0);
}

// 账 #128-b: OptionButton 的 Value 在 VB6 是 **Boolean**，而 CheckBox 的是 **三态 Integer**
// (0/1/2) —— 两边原先共用上面那一对。共用不能翻类型：把 getter 改成答 -1 会把 CheckBox 的
// 2(灰) 那档吃掉，把 setter 改成"非 0 就 -1"更糟 —— **BM_SETCHECK 只认 0/1/2**，
// 递 -1 进去按钮状态直接坏掉。所以这里单开一对：读把 BST_CHECKED 映射成 VB6 的 True(-1)，
// 写把任意非 0 折回 BST_CHECKED(1)。
int vb6_GetOptionValue(void* hwnd) {
    if (!hwnd) return 0;
    return (SendMessageW((HWND)hwnd, BM_GETCHECK, 0, 0) == BST_CHECKED) ? -1 : 0;
}

void vb6_SetOptionValue(void* hwnd, int value) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, BM_SETCHECK, (WPARAM)(value ? BST_CHECKED : BST_UNCHECKED), 0);
}

int vb6_GetControlVisible(void* hwnd) {
    if (!hwnd) return 0;
    return IsWindowVisible((HWND)hwnd) ? -1 : 0;  // VB6: True=-1
}

void vb6_SetControlVisible(void* hwnd, int visible) {
    if (!hwnd) return;
    ShowWindow((HWND)hwnd, visible ? SW_SHOW : SW_HIDE);
}

int vb6_GetControlEnabled(void* hwnd) {
    if (!hwnd) return 0;
    return IsWindowEnabled((HWND)hwnd) ? -1 : 0;  // VB6: True=-1
}

void vb6_SetControlEnabled(void* hwnd, int enabled) {
    if (!hwnd) return;
    EnableWindow((HWND)hwnd, enabled ? TRUE : FALSE);
}

// C29-SL-l（账 #143）：VB6 的 `控件.SetFocus`。原生就一句 SetFocus(hwnd) —— 它动的是**本线程
// 输入队列里的焦点**，与窗口可见/激活无关，所以无头跑里照样有效（C29-SL-g 的 SimStdEvent
// kind=4 走的就是这一句，实测会发出 WM_SETFOCUS 并点亮控件自己的 _GotFocus）。
// 拿不到焦点的那几种（控件被禁用、窗口不属于本线程）原生就是回 NULL 什么都不做，
// 本项目**不伪造、不重试** —— 真 VB6 在那里是 raise 一个错误号，而我们还没有运行期错误面。
void vb6_SetControlFocus(void* hwnd) {
    if (!hwnd) return;
    SetFocus((HWND)hwnd);
}

/* 账 #171（C29-FS-h）：单选钮的 `Click` 只在「这一枚真被选中」时才算数。见 vb6forms_prop.h 的注释。
   非单选钮一律放行 —— 这一处是那条口径的**唯一**落点，发码侧只对 OptionButton 的 arm 调它。 */
int vb6_RadioClickCounts(void* hwndFrom) {
    HWND h = (HWND)hwndFrom;
    if (!h || !IsWindow(h)) return 1;
    if ((GetWindowLongW(h, GWL_STYLE) & BS_TYPEMASK) != BS_AUTORADIOBUTTON) return 1;
    return (SendMessageW(h, BM_GETCHECK, 0, 0) == BST_CHECKED) ? 1 : 0;
}
/* 账 #175: 位置/尺寸属性的单位 = **所在容器的 ScaleMode** (VB6 语义), 不是恒为缇。
   此前这八个点位 + vb6_ControlMove 全按缇, 而 .ctl 一律声明 3=Pixel ⇒ 像素型
   UserControl 里"读来的数"和"写回去的数"彼此自洽, 却与 ScaleWidth/绘图 DC 差 15 倍。
   容器是窗体时 vb6_ContainerScaleMode 给 1 (缇), 与改动前逐字节等价。 */

// 容器 ScaleMode 的唯一取法: UserControl 宿主窗口 → 它 .ctl 声明的 ScaleMode;
// 窗体 → 窗体的 VB6_ScaleMode 属性 (缺省 1=缇); 非窗口 (NULL) → 1。
int32_t vb6_ContainerScaleMode(void* hwndParent) {
    int32_t m = vb6_UC_WindowScaleMode(hwndParent);
    if (m) return m;
    return vb6_GetScaleMode(hwndParent);
}

// 账 #175: 窗口**自身**的 ScaleMode (ScaleWidth/ScaleHeight/CurrentX 那一族读法的单位)。
int32_t vb6_WindowScaleModeSelf(void* hwnd) {
    int32_t m = vb6_UC_WindowScaleMode(hwnd);
    if (m) return m;
    return vb6_GetScaleMode(hwnd);
}

int vb6_GetControlLeft(void* hwnd) {
    if (!hwnd) return 0;
    RECT rc;
    GetWindowRect((HWND)hwnd, &rc);
    POINT pt = { rc.left, rc.top };
    ScreenToClient(GetParent((HWND)hwnd), &pt);
    return (int)vb6_ScalePxToUser((double)pt.x, vb6_ContainerScaleMode(GetParent((HWND)hwnd)), 0);
}

void vb6_SetControlLeft(void* hwnd, int left) {
    if (!hwnd) return;
    RECT rc;
    GetWindowRect((HWND)hwnd, &rc);
    POINT pt = { rc.left, rc.top };
    ScreenToClient(GetParent((HWND)hwnd), &pt);
    SetWindowPos((HWND)hwnd, NULL, vb6_ScaleUserToPx((double)left, vb6_ContainerScaleMode(GetParent((HWND)hwnd)), 0),
                 pt.y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
}

int vb6_GetControlTop(void* hwnd) {
    if (!hwnd) return 0;
    RECT rc;
    GetWindowRect((HWND)hwnd, &rc);
    POINT pt = { rc.left, rc.top };
    ScreenToClient(GetParent((HWND)hwnd), &pt);
    return (int)vb6_ScalePxToUser((double)pt.y, vb6_ContainerScaleMode(GetParent((HWND)hwnd)), 1);
}

void vb6_SetControlTop(void* hwnd, int top) {
    if (!hwnd) return;
    RECT rc;
    GetWindowRect((HWND)hwnd, &rc);
    POINT pt = { rc.left, rc.top };
    ScreenToClient(GetParent((HWND)hwnd), &pt);
    SetWindowPos((HWND)hwnd, NULL, pt.x,
                 vb6_ScaleUserToPx((double)top, vb6_ContainerScaleMode(GetParent((HWND)hwnd)), 1),
                 0, 0, SWP_NOSIZE | SWP_NOZORDER);
}

int vb6_GetControlWidth(void* hwnd) {
    if (!hwnd) return 0;
    RECT rc;
    GetWindowRect((HWND)hwnd, &rc);
    return (int)vb6_ScalePxToUser((double)(rc.right - rc.left), vb6_ContainerScaleMode(GetParent((HWND)hwnd)), 0);
}

void vb6_SetControlWidth(void* hwnd, int width) {
    if (!hwnd) return;
    RECT rc;
    GetWindowRect((HWND)hwnd, &rc);
    SetWindowPos((HWND)hwnd, NULL, 0, 0,
                 vb6_ScaleUserToPx((double)width, vb6_ContainerScaleMode(GetParent((HWND)hwnd)), 0),
                 rc.bottom - rc.top, SWP_NOMOVE | SWP_NOZORDER);
}

int vb6_GetControlHeight(void* hwnd) {
    if (!hwnd) return 0;
    RECT rc;
    GetWindowRect((HWND)hwnd, &rc);
    return (int)vb6_ScalePxToUser((double)(rc.bottom - rc.top), vb6_ContainerScaleMode(GetParent((HWND)hwnd)), 1);
}

void vb6_SetControlHeight(void* hwnd, int height) {
    if (!hwnd) return;
    RECT rc;
    GetWindowRect((HWND)hwnd, &rc);
    SetWindowPos((HWND)hwnd, NULL, 0, 0, rc.right - rc.left,
                 vb6_ScaleUserToPx((double)height, vb6_ContainerScaleMode(GetParent((HWND)hwnd)), 1),
                 SWP_NOMOVE | SWP_NOZORDER);
}

// Fix 162a-extlist: VB6 `obj.Move Left[, Top[, Width[, Height]]]` —— 语言级方法
// (Form/控件通用), 之前落 vb6_ComGetObjectProp(hwnd, L"Move") 必失败, 导致
// Form_Resize 布局全废 (extlist 的 ListView 停在设计宽, 只有 5 列可见)。
// 单位 twips (与 Left/Top/Width/Height 属性一致); mask 位 1=Left 2=Top
// 4=Width 8=Height, 未给的参数保持当前值。子控件用父客户区坐标,
// 顶层窗口 (Form) 用屏幕坐标 —— 与对应属性 setter 的坐标空间一致。
void vb6_ControlMove(void* hwnd, double L, double T, double W, double H, int mask) {
    HWND hW = (HWND)hwnd;
    if (!hW || !IsWindow(hW)) return;
    BOOL child = (GetWindowLongW(hW, GWL_STYLE) & WS_CHILD) != 0;
    RECT rc;
    GetWindowRect(hW, &rc);
    if (child) {
        POINT p0 = { rc.left, rc.top }, p1 = { rc.right, rc.bottom };
        ScreenToClient(GetParent(hW), &p0);
        ScreenToClient(GetParent(hW), &p1);
        rc.left = p0.x; rc.top = p0.y; rc.right = p1.x; rc.bottom = p1.y;
    }
    int32_t cm = vb6_ContainerScaleMode(child ? GetParent(hW) : NULL);
    int x = (mask & 1) ? vb6_ScaleUserToPx(L, cm, 0) : rc.left;
    int y = (mask & 2) ? vb6_ScaleUserToPx(T, cm, 1) : rc.top;
    int w = (mask & 4) ? vb6_ScaleUserToPx(W, cm, 0) : rc.right - rc.left;
    int h = (mask & 8) ? vb6_ScaleUserToPx(H, cm, 1) : rc.bottom - rc.top;

    /* Fix <vbeclipse> rev22: 尺寸真变了就触发该设计期子控件的 Resize 事件。
     *
     * 为什么要这一下: VB6 里 `Private Sub ViewArea_Resize()` 是 PictureBox 的
     * **Resize 事件**, 由运行时在该控件尺寸变化时自动跑。C3 的 RTL 以前没有这条
     * 转发, 于是 .ctl 写在子控件事件里的布局代码只会在"谁恰好手工调了它"时跑一次
     * —— 而那一次通常**早于布局就绪**。
     *
     * 实证 play78 (--arch x86, 探针实测): 每个 ucFolder 的 WM_SIZE 序列是
     *   设计期 (569,441) → 中间 (320,309) → 最终 (468,405)
     * `ViewArea` 每换一次尺寸, ucFolder.ctl:529 `ViewArea_Resize` 都该跑一次并按
     * **当时**的 ViewArea.Width 重摆视图窗体; 但 C3 只在 Refresh 里裸调一次, 而
     * Refresh 由 frmMain 在 Form_Load 期跑 —— 那时 rec 还没收到任何 WM_SIZE,
     * ViewArea.Width 还是设计期 8655 缇 ⇒ 五个 folder 的视图窗体**全都**只拿到
     * W=8505 (= 8655-30-2*margin), 停在同一个 567x423。
     *
     * 放在 vb6_ControlMove 而不是 uc_hostmodel_call.inc 的 Move 分派里: ViewArea /
     * ViewTabs 这些是 cgen **直调** vb6_ControlMove 的 (不过宿主分派), 那里打不到。
     * vb6_ControlMove 是所有摆位的唯一收口, 两类调用都从这里过。
     *
     * ⚠ 只在**尺寸真的变了**时触发, 否则每次 Move 都跑一遍事件 ⇒ 递归 Move 风暴
     *   (ucFolder.ViewArea_Resize 里自己就 Move 视图窗体)。*/
    int sizeChanged = ((w != rc.right - rc.left) || (h != rc.bottom - rc.top));
    SetWindowPos(hW, NULL, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
    /* Fix <vbeclipse> rev24: 设计期子控件的 `<Ctrl>_Resize` 事件在这里**排队**
     * (rev22 曾直接在这里跑 ⇒ 递归; rev23 把触发改到宿主 WM_SIZE, 但那只覆盖
     * "**宿主自己**尺寸变了"这一种 —— 见下面为什么不够)。
     *
     * 为什么宿主的 WM_SIZE 不够 (play78 --arch x86 探针实测):
     *   ucFolder 的宿主窗口在第一轮布局后**尺寸再没变过**, 所以 rev23 的 drain
     *   全部挤在启动那一段跑完 (探针: 15 次 Queue / 15 次 Drain, 行号 50-68 连成一片),
     *   而 ViewArea 的最终尺寸 (567→318/136/425/101) 是**之后**才由
     *   `UserControl_Resize` 里的 `ViewArea.Move …` 定下的 —— 那发的是
     *   **ViewArea 自己**的 WM_SIZE, 不会回到宿主 ⇒ 宿主那条排队链再没机会跑。
     *   结果 `ViewArea_Resize` 全程只读到设计期宽度 567px (= 8505 缇)。
     *
     * 为什么现在挂这里不会递归 (rev22 踩过的坑):
     *   rev22 是**直接调用** `desc->designResize(...)`, 而 `ViewArea_Resize` 内部
     *   自己就 Move 视图窗体 ⇒ 又进来 ⇒ 无限展开。
     *   现在只 `PostMessage` **排队**(rev23 的机制), 真正的调用发生在**回到消息
     *   循环之后**, 那时本轮 Move 早已全部返回, 不存在栈上的重入。
     *   `vb6_UC_QueueDesignResize` 自带窗口属性去重, 一串嵌套 Move 只排一条。
     */
    if (sizeChanged) {
        extern int32_t vb6_UC_QueueDesignResizeForCtrl(const void* hwnd);
        vb6_UC_QueueDesignResizeForCtrl(hW);
        /* Fix <vbeclipse> rev29: 目标是**窗体**时, 直接调它的 Form_Resize。
         *
         * 为什么必须"直调"而不是 PostMessage 排队 (rev29 前一版实测失败, 留档):
         *   停靠视图窗体 (frmViewViews 等) 被 `vb6_ComCall(视图窗体, L"Move", …)`
         *   摆到最终尺寸 —— 尺寸**确实变了** (探针: rect 0x0 → 318x291 → 202x364),
         *   但它**一生只收到 1 次 WM_SIZE 且那次是建窗时的 0x0** ⇒ rev28 那条
         *   `case VB6_FORM_FR_MSG` 永远等不到, Form_Resize 跑 0 次。
         *   改在 `vb6_ControlMove` 里 PostMessage 也不行: 排队那一刻还在
         *   Form_Load 的同步调用栈中、**根本不在主消息循环里**, 消息进队列没人
         *   Dispatch, 实测还堆损坏 `0xC0000374`。已回退。
         *
         * 为什么现在直调是安全的:
         *   ① 时机对 —— `SetWindowPos` 已在上面执行完, 目标窗体拿到的是**最终**尺寸,
         *      Form_Resize 里 `vb6_GetScaleWidth(hwnd)` 读到的就是终值, 不是中间态
         *      (这正是 rev23/rev28 花两版才解决的问题, 现在在 Move 收口处天然成立)。
         *   ② 不递归 —— `vb6_InvokeFormResize` 有"同窗体重入"闸; 且 Form_Resize
         *      内部 Move 的是**子控件**(另一个 HWND), 那一发进来时本窗体已出栈,
         *      靠 `sizeChanged` 判据收敛。实测栈深 ≤ 2。
         *   ③ 收窄到窗体 —— 只有 `vb6_form_load_<F>` 注册过的 HWND 才有那个属性,
         *      原生控件 (TreeView/ListView/Edit) 根本没注册 ⇒ 自动跳过。
         *      这一点很关键: rev29 前一版给**所有**控件 PostMessage 自定义消息,
         *      它们的窗口过程是 Windows 自带的, 不认那条消息, 纯浪费去重名额。*/
        {
            extern int32_t vb6_InvokeFormResize(void* hwnd);
            vb6_InvokeFormResize(hW);
        }
    }
}

// P11.8: hWnd attribute (read-only)
void* vb6_GetControlHwnd(void* hwnd) {
    return hwnd;  // Already the HWND
}

// ============================================================
// P13.1: Font properties
// ============================================================

// 账 #200：**这枚控件现在在用的字体**只从这里问一次。
// 为什么要有这一处：PictureBox / Label 那几枚是 STATIC 类，而这个类**不记字体** ——
// 一次性探针（`.build/b200probe/fontprobe.c`，裸 STATIC、谁也没子类化）实测
// `WM_SETFONT` 之后 `WM_GETFONT` 回 NULL，`STM_SETFONT`/`STM_GETFONT` 同样回 NULL。
// 于是原来那三条读法（`vb6_GetControlLogFont` 给 .FontName/.FontSize 用、Print 落笔前选字体、
// #196 的文字量）在这类窗口上**永远拿不到用户设的字体**，只能拿 DC 的默认字体画，
// 而 .FontSize 读回来还是设计值（那是另一份自存的属性）—— 两头谁都不报错。
// 所以 setter 现在把自己创建的那张 HFONT 存进 `VB6_CtrlFont`（新名字，与 #185 那条
// "一层一个窗口属性名"同纪律），这里优先问窗口、问不到再读这份自存的。
HFONT vb6_ControlFont(HWND hw) {
    HFONT h = (HFONT)SendMessageW(hw, WM_GETFONT, 0, 0);
    if (!h) h = (HFONT)GetPropW(hw, L"VB6_CtrlFont");
    return h;
}

// 账 #204: 那份自存只有这一个写口 —— 谁把字体发给窗口，谁就在这里存同一张。
// 创建期那一站原来只发不存，而 STATIC/BUTTON 那一类窗口不答 `WM_GETFONT`（见上面那段），
// 于是"从没被写过字体"的控件在出口这一头两问皆空：`.FontName` 读空串、`.FontSize` 读 0、
// 文字量按 DC 的默认字体算（实测 bName= bfs=0 bpf=0 bth=16，16 是 Segoe UI 9pt、不是 VB6 的 8.25pt）。
// 存了之后还顺带有个副作用：setter 换字体时找得到"上一张是我们造的"，那张才删得掉。
void vb6_ControlFontStore(HWND hw, HFONT hFont) {
    if (!hw || !hFont) return;
    SetPropW(hw, L"VB6_CtrlFont", (HANDLE)hFont);
}

// Helper: get LOGFONT from control's current font
static int vb6_GetControlLogFont(void* hwnd, LOGFONTW* plf) {
    if (!hwnd || !plf) return 0;
    HFONT hFont = vb6_ControlFont((HWND)hwnd);
    if (!hFont) return 0;
    return GetObjectW(hFont, sizeof(LOGFONTW), plf) > 0;
}

// Helper: create new font from modified LOGFONT and set it on control
// Also deletes the old font if it was created by us (we track via prop)
static void vb6_SetControlFontFromLogFont(void* hwnd, const LOGFONTW* plf) {
    if (!hwnd || !plf) return;
    HFONT hNewFont = CreateFontIndirectW(plf);
    if (!hNewFont) return;
    // 旧字体先按"我们存过的那张"找，找不到才退回问窗口 —— 顺序反了会双删：
    // 真记字体的那几类控件（EDIT/BUTTON…）WM_GETFONT 回来的就是我们上一轮存进去的那张。
    HFONT hOldFont = (HFONT)GetPropW((HWND)hwnd, L"VB6_CtrlFont");
    if (!hOldFont) hOldFont = vb6_ControlFont((HWND)hwnd);
    SendMessageW((HWND)hwnd, WM_SETFONT, (WPARAM)hNewFont, (LPARAM)TRUE);
    vb6_ControlFontStore((HWND)hwnd, hNewFont);
    // Force redraw
    InvalidateRect((HWND)hwnd, NULL, TRUE);
    // Delete old font only if it's not a stock font
    if (hOldFont && GetObjectType(hOldFont) == OBJ_FONT) {
        // Safe to delete non-stock fonts; stock fonts have OBJ_FONT but
        // DeleteObject on stock fonts is a no-op, so it's safe
        DeleteObject(hOldFont);
    }
}

void* vb6_GetControlFontName(void* hwnd) {
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) {
        WCHAR empty[] = {0};
        return SysAllocString(empty);
    }
    return SysAllocString(lf.lfFaceName);
}

void vb6_SetControlFontName(void* hwnd, void* bstrName) {
    if (!hwnd || !bstrName) return;
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) {
        // No existing font, create a default LOGFONT
        memset(&lf, 0, sizeof(lf));
        lf.lfHeight = -13;  // Default ~10pt
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = DEFAULT_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    }
    BSTR bs = (BSTR)bstrName;
    int len = SysStringLen(bs);
    if (len > LF_FACESIZE - 1) len = LF_FACESIZE - 1;
    memcpy(lf.lfFaceName, bs, len * sizeof(WCHAR));
    lf.lfFaceName[len] = L'\0';
    vb6_SetControlFontFromLogFont(hwnd, &lf);
}

// C29-SL-q（账 #154）: 设计期那一条走这一支 —— `.frm` 里的字体名在生成码里是一枚 C 宽字符字面量，
// 不是一枚 BSTR，而上面那支要 `SysStringLen` 量长度，喂字面量会把串尾之后的内存算进去。
// （不改用 `vb6_BSTR_FromStr` 现造一枚：那要么在发码里漏一枚串 —— 本仓刚为同类临时串开过 #119。）
void vb6_SetControlFontNameW(void* hwnd, const wchar_t* name) {
    LOGFONTW lf;
    int len;
    if (!hwnd || !name || !name[0]) return;
    if (!vb6_GetControlLogFont(hwnd, &lf)) {
        memset(&lf, 0, sizeof(lf));
        lf.lfHeight = -13;  // Default ~10pt
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = DEFAULT_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    }
    len = lstrlenW(name);
    if (len > LF_FACESIZE - 1) len = LF_FACESIZE - 1;
    memcpy(lf.lfFaceName, name, len * sizeof(WCHAR));
    lf.lfFaceName[len] = L'\0';
    vb6_SetControlFontFromLogFont(hwnd, &lf);
}

// C29-SL-p（`ai/内置控件/Slider 控件（滑杆）.md` §4 那条例子量出来的，探针 .build/slfont）:
// 点号 → 像素是**有损**的一步（96 DPI 下 1pt = 1.3333px，字体高度只能取整），所以旧写法
// 从窗口反算会把请求值量化掉：写 8 读回 8.25、写 10 读回 9.75、写 14 读回 14.25。
// 窗口表示不了的那一半按**请求值自存**（与 Slider 的 TickFrequency / TextPosition 同一族口径）。
// Set 旗标不能省：`SetPropW(hwnd, name, 0)` 等于删属性（账 #107 踩过），而 0.0f 的位就是 0 ——
// 少了这一枚，写 0 那一档会静默变回"没设过"。
static const wchar_t kFontPtProp[]    = L"VB6_FontPt";
static const wchar_t kFontPtSetProp[] = L"VB6_FontPtSet";

float vb6_GetControlFontSize(void* hwnd) {
    LOGFONTW lf;
    float pt;
    DWORD bits;
    if (!hwnd) return 0.0f;
    if (GetPropW((HWND)hwnd, kFontPtSetProp)) {
        bits = (DWORD)(DWORD_PTR)GetPropW((HWND)hwnd, kFontPtProp);
        memcpy(&pt, &bits, sizeof(pt));
        return pt;
    }
    if (!vb6_GetControlLogFont(hwnd, &lf)) return 0.0f;
    HDC hdc = GetDC(NULL);
    int dpi = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(NULL, hdc);
    if (dpi <= 0) dpi = 96;
    int heightPx = lf.lfHeight < 0 ? -lf.lfHeight : lf.lfHeight;
    return (float)heightPx * 72.0f / (float)dpi;
}

// C29-SL-p 的判据证人（**不是 VB6 属性**，与 TickPresent / TravelIsVert / ToolTipRegistered 同族）:
// 窗口现在真在用的字体像素高度。有了它，字号那条判据才是两头的 —— 自存的数读回来当然还是自存的数，
// 只有问窗口才知道这次 WM_SETFONT 到底发没发出去。
int vb6_ControlFontPixelHeight(void* hwnd) {
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) return 0;
    return lf.lfHeight < 0 ? -lf.lfHeight : lf.lfHeight;
}

void vb6_SetControlFontSize(void* hwnd, float sizePt) {
    if (!hwnd) return;
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) {
        memset(&lf, 0, sizeof(lf));
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = DEFAULT_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    }
    HDC hdc = GetDC(NULL);
    int dpi = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(NULL, hdc);
    if (dpi <= 0) dpi = 96;
    // Convert points to pixel height (negative for character height)
    lf.lfHeight = -(int)(sizePt * (float)dpi / 72.0f + 0.5f);
    vb6_SetControlFontFromLogFont(hwnd, &lf);
    {
        DWORD bits;
        memcpy(&bits, &sizePt, sizeof(bits));
        SetPropW((HWND)hwnd, kFontPtProp, (HANDLE)(DWORD_PTR)bits);
        SetPropW((HWND)hwnd, kFontPtSetProp, (HANDLE)1);
    }
}

int vb6_GetControlFontBold(void* hwnd) {
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) return 0;
    return (lf.lfWeight >= FW_BOLD) ? -1 : 0;  // VB6: True=-1
}

void vb6_SetControlFontBold(void* hwnd, int bold) {
    if (!hwnd) return;
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) {
        memset(&lf, 0, sizeof(lf));
        lf.lfHeight = -13;
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = DEFAULT_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    }
    lf.lfWeight = bold ? FW_BOLD : FW_NORMAL;
    vb6_SetControlFontFromLogFont(hwnd, &lf);
}

int vb6_GetControlFontItalic(void* hwnd) {
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) return 0;
    return lf.lfItalic ? -1 : 0;  // VB6: True=-1
}

void vb6_SetControlFontItalic(void* hwnd, int italic) {
    if (!hwnd) return;
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) {
        memset(&lf, 0, sizeof(lf));
        lf.lfHeight = -13;
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = DEFAULT_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    }
    lf.lfItalic = italic ? TRUE : FALSE;
    vb6_SetControlFontFromLogFont(hwnd, &lf);
}

int vb6_GetControlFontUnderline(void* hwnd) {
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) return 0;
    return lf.lfUnderline ? -1 : 0;  // VB6: True=-1
}

void vb6_SetControlFontUnderline(void* hwnd, int underline) {
    if (!hwnd) return;
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) {
        memset(&lf, 0, sizeof(lf));
        lf.lfHeight = -13;
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = DEFAULT_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    }
    lf.lfUnderline = underline ? TRUE : FALSE;
    vb6_SetControlFontFromLogFont(hwnd, &lf);
}

int vb6_GetControlFontStrikethrough(void* hwnd) {
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) return 0;
    return lf.lfStrikeOut ? -1 : 0;  // VB6: True=-1
}

void vb6_SetControlFontStrikethrough(void* hwnd, int strike) {
    if (!hwnd) return;
    LOGFONTW lf;
    if (!vb6_GetControlLogFont(hwnd, &lf)) {
        memset(&lf, 0, sizeof(lf));
        lf.lfHeight = -13;
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = DEFAULT_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    }
    lf.lfStrikeOut = strike ? TRUE : FALSE;
    vb6_SetControlFontFromLogFont(hwnd, &lf);
}

// ============================================================
// P13.2: ForeColor/BackColor
// ============================================================

int vb6_GetControlForeColor(void* hwnd) {
    if (!hwnd) return 0;
    // For most controls, text color is set via WM_CTLCOLOR* parent handler
    // We store foreground color as a window property
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ForeColor");
    if (hProp) return (int)(INT_PTR)hProp;
    return 0;  // Default black
}

void vb6_SetControlForeColor(void* hwnd, int color) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ForeColor", (HANDLE)(INT_PTR)color);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

int vb6_GetControlBackColor(void* hwnd) {
    if (!hwnd) return 0;
    // Fix 187: 判「Set 过」要用独立哨兵 (VB6_BackColorSet) — 黑色 0x000000
    // 经 (HANDLE) 转换是 NULL, 若只判值属性, 黑色会被当「未设置」回落到
    // GetSysColor(COLOR_BTNFACE) → 颜色对话框选黑 OK 后色块永远显示灰
    // (Test.exe 实测: BG 属性写入成功但色块像素不变, 就是这条路)。
    if (GetPropW((HWND)hwnd, L"VB6_BackColorSet"))
        return (int)(INT_PTR)GetPropW((HWND)hwnd, L"VB6_BackColor");
    return (int)(INT_PTR)GetSysColor(COLOR_BTNFACE);  // Default
}

void vb6_SetControlBackColor(void* hwnd, int color) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_BackColor", (HANDLE)(INT_PTR)color);
    // Fix 187: 「Set 过」哨兵必须独立于值本身 — color=0 (黑色) 经
    // (HANDLE)(INT_PTR)0 转换后是 NULL, 与「未设置」同构, 宿主判空检查会把
    // 黑色误判为未设置 (颜色对话框默认选黑 → OK 后色块不动的直接原因)。
    SetPropW((HWND)hwnd, L"VB6_BackColorSet", (HANDLE)1);
    // Fix 187: 背景色变化后必须失效 WM_CTLCOLORSTATIC 侧的刷子缓存
    // (vb6_ApplyCtlColorStatic 缓存在 VB6_BgBrush), 否则色块变完一次颜色后
    // 再变其他颜色永远显示第一把刷子。
    HBRUSH oldBr = (HBRUSH)GetPropW((HWND)hwnd, L"VB6_BgBrush");
    if (oldBr) {
        DeleteObject(oldBr);
        RemovePropW((HWND)hwnd, L"VB6_BgBrush");
    }
    if (GetEnvironmentVariableA("C3_FORMS_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_F187] SetBackColor hwnd=%p color=%06X\r\n",
                hwnd, (unsigned)color & 0xFFFFFFu);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

// ============================================================
// Fix 187: WM_CTLCOLORSTATIC 统一应用 (主窗体 WndProc / SSTab 容器共用)
// ------------------------------------------------------------
// VB6_BackColor/ForeColor 窗口属性此前只有 uc 宿主窗口消费
// (uc_host_window.c WM_CTLCOLORSTATIC), 主窗体直接子控件 (picBackColor 等
// STATIC 类色块) 与 SSTab 页内子控件的宿主 (生成窗体 WndProc /
// sstabSubclassProc) 均不处理 → vb6_SetControlBackColor 只写了属性 +
// InvalidateRect, 重绘路径读不到 → 色块永远默认底色 (Test.exe 颜色对话框
// OK 后 "颜色没有写入到颜色对话左边的控件里")。
// 约定: 仅当子控件 Set 过 VB6_BackColor 属性时才接管; 否则返回 0 由调用方
// break 到 DefWindowProcW 维持原生外观。
// ============================================================
LRESULT vb6_ApplyCtlColorStatic(HDC hdc, HWND child) {
    if (!hdc || !child) return 0;
    // Fix 187: 只看 Set 哨兵, 不看值 — color=0 (黑) 的值属性是 NULL, 但它是合法色。
    if (!GetPropW(child, L"VB6_BackColorSet")) return 0;  // 未显式设色 → 调用方走默认绘制
    COLORREF bg = (COLORREF)(INT_PTR)GetPropW(child, L"VB6_BackColor");
    COLORREF fg = (COLORREF)vb6_GetControlForeColor((void*)child);
    if (bg & 0x80000000L) bg = GetSysColor(bg & 0xFF);
    SetTextColor(hdc, fg);
    SetBkColor(hdc, bg);
    HBRUSH br = (HBRUSH)GetPropW(child, L"VB6_BgBrush");
    if (!br) {
        br = CreateSolidBrush(bg);
        SetPropW(child, L"VB6_BgBrush", (HANDLE)br);
    }
    return (LRESULT)br;
}

// ============================================================
// Fix 162f-extlist: WM_CTLCOLORBTN 统一答复 (按钮类子控件的背景刷)
// ------------------------------------------------------------
// Button 类 (BUTTON 窗口类) 通过 WM_CTLCOLORBTN 向**父窗**要绘制用刷子。分两种:
//   (1) CheckBox / OptionButton: VB6 语义 = **透明**。之前返回 HOLLOW_BRUSH 即可,
//       因为它们的文字直接坐在父窗底色上, 用空刷 → 不画背景 → 透出父窗灰底。
//   (2) Frame (BS_GROUPBOX): 经典 (关主题后) groupbox 的**标题**由 BUTTON 绘制器
//       先 `FillRect(标题矩形, 该刷子)` 再 `DrawText` —— 返回 HOLLOW_BRUSH 会让
//       这块矩形**什么都不填**, 于是露出 groupbox 自身窗口的底色 (经典 BUTTON 是
//       白), 表现为标题后面一条**白色填充块** (用户实测)。VB6 里 Frame 标题是坐在
//       父窗 BackColor 上的 ⇒ 这里必须返回**父窗底色的实心刷** (通常 240 灰)。
// 判据: 子控件的 GWL_STYLE & 0xF == BS_GROUPBOX(0x7)。其余 Button 类仍走空刷。
// 实心刷缓存在子控件窗口属性上 (按控件一份, 值变了由 vb6_SetControlBackColor 失效)。
// ============================================================
LRESULT vb6_CtlColorBtnBrush(HWND child, HWND parent) {
    LONG_PTR st = child ? GetWindowLongPtrW(child, GWL_STYLE) : 0;
    if ((st & 0x0000000FL) != 0x7L) {
        // 非 groupbox: CheckBox/OptionButton/命令按钮 → 透明语义 (空刷)。
        return (LRESULT)GetStockObject(HOLLOW_BRUSH);
    }
    // groupbox: 标题底 = 父窗 BackColor (未显式设色时回落 BTNFACE = 240 灰)。
    COLORREF bg = parent ? (COLORREF)vb6_GetControlBackColor((void*)parent)
                         : GetSysColor(COLOR_BTNFACE);
    if (bg & 0x80000000L) bg = GetSysColor(bg & 0xFF);
    HBRUSH br = (HBRUSH)GetPropW(child, L"VB6_GbCapBrush");
    if (br) {
        // 缓存命中: 若父窗底色已变, 需要重建 (色值存在 VB6_GbCapColor 上)。
        COLORREF cached = (COLORREF)(INT_PTR)GetPropW(child, L"VB6_GbCapColor");
        if (cached == bg) return (LRESULT)br;
        DeleteObject(br);
        RemovePropW(child, L"VB6_GbCapBrush");
    }
    br = CreateSolidBrush(bg);
    if (!br) return (LRESULT)GetStockObject(HOLLOW_BRUSH);
    SetPropW(child, L"VB6_GbCapBrush", (HANDLE)br);
    SetPropW(child, L"VB6_GbCapColor", (HANDLE)(INT_PTR)bg);
    return (LRESULT)br;
}

// ============================================================
// Fix 185: 控件级绘制入口 PictureBox.Print / PictureBox.Cls
// ============================================================
//
// DC 来源只有下面这一处口径（账 #196）：_Paint 派发时挂上的 VB6_PaintDC（BeginPaint/
// EndPaint 之间才有效，绝不能 ReleaseDC），否则回落 GetDC。VB6 允许在非 _Paint 时机
// Print，效果就是画在屏幕上、下次重绘即消失，这里保持同样的宽松度。
// 笔位不在这里（账 #239）：Print/Cls 都转调 vb6forms_draw.c 的同一份实现，
// 笔位那份全仓唯一存储（VB6_CurrentX/Y，float）由它去问。
//
// 账 #196：**「这枚控件的绘图 DC 从哪儿来」只有下面这一处口径**（与 `.hDC` 共用）。
// 区别只在句柄归谁：Print/Cls 这类内部调用用完就 ReleaseDC；而交回给 VB 代码的
// `.hDC` 必须留着 —— VB6 是一个对象一张 hDC，反复读要读回同一个值，所以那一档
// 按 HWND 缓存进窗口属性 `VB6_ObjectDC`，由 PictureBox/Image 那层自己的 WM_DESTROY
// 归还（见 vb6forms_picture_prop.c）。以前这条没处走：`.hDC` 只能撞
// `cgen_expr_with.cpp` 那条 "hwnd.成员" 兜底 = C2039（真工程物证 ucTreeMaps PropPagFMR.c:74）。

HDC vb6_ControlDrawDC(HWND hw, BOOL* pFromPaint) {
    HDC hdc = (HDC)GetPropW(hw, L"VB6_PaintDC");
    *pFromPaint = (hdc != NULL) ? TRUE : FALSE;
    if (hdc) return hdc;
    return GetDC(hw);
}

intptr_t vb6_GetControlHDC(void* hwnd) {
    if (!hwnd) return 0;
    HWND hw = (HWND)hwnd;
    BOOL fromPaint = FALSE;
    HDC hdc = vb6_ControlDrawDC(hw, &fromPaint);   // 口径只有上面那一处
    if (!hdc) return 0;
    if (fromPaint) return (intptr_t)hdc;           // 派发期那张：既不缓存也不释放
    HDC held = (HDC)GetPropW(hw, L"VB6_ObjectDC");
    if (held) { ReleaseDC(hw, hdc); return (intptr_t)held; }   // 刚才那张是白拿的
    SetPropW(hw, L"VB6_ObjectDC", (HANDLE)hdc);
    return (intptr_t)hdc;
}

// 账 #196：按 HWND 的**文字量**（TextHeight / TextWidth）。两头都要对上才成立：
// ① 量的是**这枚控件自己的字体** —— 与上面 Print 同一口径（WM_GETFONT），不是屏幕默认字体
//    （Fix 129 在 UserControl 那一族栽过的同一件事）；DC 也从同一处权威拿，字体的 DPI 才与
//    落笔的 DPI 同源。② 交回的单位是**这枚控件自己的 ScaleMode** —— 像素量完必须过
//    vb6_ScalePxToUser（账 #177 那条：缇型对象上直接交像素，比同一枚控件的 ScaleWidth 小 15 倍）。
// 语料物证：Charts 2020/ucTreeMaps 的 PropPagFMR.pag:265 `.CurrentY + .TextHeight(Text)` ——
// 量出来的数是要加到画笔光标上的，单位错了行距就错了。
static int vb6_ControlMeasureTextPx(void* hwnd, BSTR text, int wantWidth) {
    HWND hw = (HWND)hwnd;
    int len = text ? (int)SysStringLen(text) : 0;
    if (!len) return 0;
    BOOL fromPaint = FALSE;
    HDC hdc = vb6_ControlDrawDC(hw, &fromPaint);
    if (!hdc) return 0;
    HFONT hFont = vb6_ControlFont(hw);   // 账 #200: 字体只从 vb6_ControlFont 那一处问
    HFONT hOld = hFont ? (HFONT)SelectObject(hdc, hFont) : NULL;
    SIZE sz = { 0, 0 };
    GetTextExtentPoint32W(hdc, text, len, &sz);
    if (hOld) SelectObject(hdc, hOld);
    if (!fromPaint) ReleaseDC(hw, hdc);
    return wantWidth ? sz.cx : sz.cy;
}

// 返回档刻意用 **float**：VB6 的 TextHeight/TextWidth 就是 Single，RTL 这头直接交 Single
// 宽的数 ⇒ 生成 C 里 `t = picA.TextHeight(s)` 不必再靠 double→float 的隐式收窄（C4244）。
float vb6_ControlTextWidth(void* hwnd, void* bstrText) {
    if (!hwnd) return 0;
    return (float)vb6_ScalePxToUser((double)vb6_ControlMeasureTextPx(hwnd, (BSTR)bstrText, 1),
                                    vb6_WindowScaleModeSelf(hwnd), 0);
}

float vb6_ControlTextHeight(void* hwnd, void* bstrText) {
    if (!hwnd) return 0;
    return (float)vb6_ScalePxToUser((double)vb6_ControlMeasureTextPx(hwnd, (BSTR)bstrText, 0),
                                    vb6_WindowScaleModeSelf(hwnd), 1);
}

// 账 #221 = C29-PL-a: VB6 的 Picture.Line 落到原生 GDI。
//
// 为什么非补不可: 语料里那 8 条 Line 调用全发成
// `vb6_ComCallObject(vb6_ComGetObjectProp(vb6_hwnd_PictureN, L"Line"), L"Item", {...}, 6)`
// —— 先把方法名当**属性**取、再对取回的东西取 Item。而 RTL 两处都把 Line 登记成
// 「认识但什么都不做」(vb6forms_axcontainer.c 的属性位交回 Empty + uc_hostmodel_call.inc
// 的 return 1)，于是**两跳都返回成功、两跳都不落笔** —— 编得过、跑得起、画面空白。
//
// 三条口径:
//   * DC 走 vb6_ControlPrint / vb6_ControlCls **同一处** vb6_ControlDrawDC（_Paint 派发
//     挂在身上的 VB6_PaintDC 优先，取不到才 GetDC），不另开第二条取 DC 的路。
//   * 坐标按**这枚窗口自己的** ScaleMode 换算（vb6_WindowScaleModeSelf + vb6_ScaleUserToPx，
//     与 #196/#197 那一条单位表同一个来源），不自己再算一遍缇/像素。
//   * style 的位口径与 parser 折旗标那一处是**同一张表**（parser_expr_postfix.cpp 的
//     vb6_LineStyleBit 注释）: 1=B 矩形、2=C 椭圆、4=F 填充，故 BF = 1|4 = 5。
//     color 传负数 = VB6 那一面的"没写颜色"，落到控件自己的 ForeColor。
//
// 刻意没做的两头（写在账里，别当已验）: ScaleLeft/ScaleTop 的**原点偏移**没进来
// （语料的 PictureBox 都是 0），VB6 那条"Line 之后 CurrentX/CurrentY 移到终点"也没进来
// （调用点从没读回它，接进来要先定 CurrentX 的单位口径，见 #192 那一格）。
void vb6_ControlLine(void* hwnd, double x1, double y1, double x2, double y2,
                     int32_t color, int32_t style) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    int32_t mode = vb6_WindowScaleModeSelf(hw);
    int ax = vb6_ScaleUserToPx(x1, mode, 0);
    int ay = vb6_ScaleUserToPx(y1, mode, 1);
    int bx = vb6_ScaleUserToPx(x2, mode, 0);
    int by = vb6_ScaleUserToPx(y2, mode, 1);

    BOOL fromPaint = FALSE;
    HDC hdc = vb6_ControlDrawDC(hw, &fromPaint);
    if (!hdc) return;

    COLORREF col = (color < 0) ? (COLORREF)vb6_GetControlForeColor(hw) : (COLORREF)color;
    HPEN pen = CreatePen(PS_SOLID, 1, col);
    HPEN hOldPen = pen ? (HPEN)SelectObject(hdc, pen) : NULL;
    // 不填充那一档必须显式给 NULL_BRUSH: 留着上一次的画刷, Line 会顺带填出一块颜色
    HBRUSH brush = (style & 4) ? CreateSolidBrush(col) : (HBRUSH)GetStockObject(NULL_BRUSH);
    HBRUSH hOldBrush = brush ? (HBRUSH)SelectObject(hdc, brush) : NULL;

    if (style & 2) {
        Ellipse(hdc, ax, ay, bx, by);        // C: 两点是外接矩形
    } else if (style & 1) {
        Rectangle(hdc, ax, ay, bx, by);      // B: 矩形
    } else {
        POINT oldpt;
        MoveToEx(hdc, ax, ay, &oldpt);       // 默认: 线段
        LineTo(hdc, bx, by);
    }

    if (hOldPen) SelectObject(hdc, hOldPen);
    if (hOldBrush) SelectObject(hdc, hOldBrush);
    if (pen) DeleteObject(pen);
    if (brush && (style & 4)) DeleteObject(brush);
    if (!fromPaint) ReleaseDC(hw, hdc);
}

// 账 #239: 这两条以前是本族**另写的一份**实现 —— 笔位存在窗口属性 VB6_PrintX/Y 上、
// 按像素推进，既不读 pic.CurrentX/CurrentY 也不动它们（实测把笔位放到 20 像素后 Print，
// 墨仍落在第 2 行；Print 之后 CurrentY 的推进是 0，而同一枚控件自己答 TextHeight = 13）。
// Form 那一族的 vb6_Form_Print / vb6_Form_Cls 经过 #233(笔位) #234(DC) #235(色)
// #237(单位) 之后，五件都问的已经是全仓唯一的权威 ⇒ 这里直接转调：控件与窗体的
// Print/Cls 从此同一份代码，缺的那件补上，抄的那份撤掉。
void vb6_ControlCls(void* hwnd) {
    vb6_Form_Cls(hwnd);
}

void vb6_ControlPrint(void* hwnd, void* bstrText) {
    vb6_Form_Print(hwnd, bstrText);
}


// ============================================================
// C29-9 / D6: CommonDialog —— 原生 comdlg32，**不再走 MSComDlg.OCX**
// ============================================================
// 为什么必须换：那个 OCX 只有 32 位，x64 进程里 CoCreateInstance 直接失败，于是
// 今天这枚控件在 64 位下是静默空转（029 §九 给了实测读数：属性读回全空、ShowOpen
// 不出现、退出码照旧 0）。
//
// 实现口径：
//   * 属性宿主 = 一枚自注册的不可见子窗口 `VB6_COMMONDIALOG`（0x0、不带 WS_VISIBLE）。
//     有了句柄，字符串/整数属性就照 DirListBox 那一族同一套 SetPropW 存法走，cgen 的
//     "readFn(hwnd)" / "writeFn(hwnd, v)" 形状完全不用特判。
//     整数一律存 val+1：SetPropW(hwnd, name, (HANDLE)0) 与"从没设过"不可分辨，
//     而 Flags / Color / Min 的合法取值域含 0（C29-1a 那条教训）。
//   * API 只经 LoadLibrary + GetProcAddress 取（照 vb6_di_com_stubs.c 里 GDI+ 那条例子）
//     => 不给工具链加新的 import lib 依赖。
//   * VB6 的 Filter 用竖线串，原生 OPENFILENAME 要 `\0` 分隔的成对表 —— 折叠集中在
//     vb6_CdFoldFilter 一处做；读回时原样返回竖线串（VB6 的读数口径）。
//   * 取消：CancelError=True 时报 32755（VB6 的 cdlCancel），且**不改**已有读数。

#include <commdlg.h>

// 本模块的头链没引 vb6rtl_array.h，而取消路径要报 VB6 的 32755 —— 不声明就会踩
// "隐式原型"那一刀（C29-1b 刚被咬过：返回指针的函数按 int 取）。照 vb6com_internal.h 同法补一句。
extern void vb6_RaiseError(int32_t errNum, void* description);

static HMODULE vb6_ComDlgModule(void) {
    static HMODULE hMod = NULL;
    if (!hMod) hMod = LoadLibraryW(L"comdlg32.dll");
    return hMod;
}

static const wchar_t* vb6_CdDupSrc(const wchar_t* s) { return s ? s : (const wchar_t*)L""; }

static wchar_t* vb6_CdDupStr(const wchar_t* s) {
    s = vb6_CdDupSrc(s);
    size_t n = wcslen(s) + 1;
    wchar_t* b = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, n * sizeof(wchar_t));
    if (b) wcscpy_s(b, n, s);
    return b;
}

static wchar_t* vb6_CdGetStr(void* hwnd, const wchar_t* key) {
    if (!hwnd) return SysAllocString(L"");
    HANDLE h = GetPropW((HWND)hwnd, key);
    return SysAllocString(h ? (const wchar_t*)h : (const wchar_t*)L"");
}

static void vb6_CdSetStr(void* hwnd, const wchar_t* key, const wchar_t* v) {
    if (!hwnd) return;
    HANDLE old = GetPropW((HWND)hwnd, key);
    if (old) { RemovePropW((HWND)hwnd, key); HeapFree(GetProcessHeap(), 0, old); }
    SetPropW((HWND)hwnd, key, vb6_CdDupStr(v));
}

static int vb6_CdGetInt(void* hwnd, const wchar_t* key, int dflt) {
    if (!hwnd) return dflt;
    HANDLE h = GetPropW((HWND)hwnd, key);
    return h ? (int)(INT_PTR)h - 1 : dflt;
}

static void vb6_CdSetInt(void* hwnd, const wchar_t* key, int v) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, key, (HANDLE)(INT_PTR)(v + 1));
}

#define VB6_CD_STR(Name, Key)                                                \
    wchar_t* vb6_CdGet##Name(void* hwnd) { return vb6_CdGetStr(hwnd, Key); }  \
    void vb6_CdSet##Name(void* hwnd, wchar_t* v) { vb6_CdSetStr(hwnd, Key, v); }

#define VB6_CD_INT(Name, Key, Dflt)                                          \
    int vb6_CdGet##Name(void* hwnd) { return vb6_CdGetInt(hwnd, Key, Dflt); } \
    void vb6_CdSet##Name(void* hwnd, int v) { vb6_CdSetInt(hwnd, Key, v); }

VB6_CD_STR(Filter,      L"VB6_Cd_Filter")
VB6_CD_STR(FileName,    L"VB6_Cd_FileName")
VB6_CD_STR(FileTitle,   L"VB6_Cd_FileTitle")
VB6_CD_STR(DialogTitle, L"VB6_Cd_DialogTitle")
VB6_CD_STR(InitDir,     L"VB6_Cd_InitDir")
VB6_CD_STR(DefaultExt,  L"VB6_Cd_DefaultExt")
VB6_CD_STR(FontName,    L"VB6_Cd_FontName")
VB6_CD_INT(Flags,       L"VB6_Cd_Flags",       0)
// CancelError 是布尔：VB6 的 True 是 **-1**，而整数袋存的是 val+1（避开 SetPropW 存 0
// 与"从没设过"不可分辨那一坑）⇒ -1 会被存成 0、读回来变成 False。这里单独 normalize：
// 存 0/1 再 +1，读回按 VB6 口径给 0 / -1。
int vb6_CdGetCancelError(void* hwnd) {
    return vb6_CdGetInt(hwnd, L"VB6_Cd_CancelError", 0) ? -1 : 0;
}
void vb6_CdSetCancelError(void* hwnd, int v) {
    vb6_CdSetInt(hwnd, L"VB6_Cd_CancelError", v ? 1 : 0);
}

VB6_CD_INT(Color,       L"VB6_Cd_Color",       0)
VB6_CD_INT(Min,         L"VB6_Cd_Min",         0)
VB6_CD_INT(Max,         L"VB6_Cd_Max",         0)
VB6_CD_INT(Copies,      L"VB6_Cd_Copies",      1)
VB6_CD_INT(FontSize,    L"VB6_Cd_FontSize",    0)
// Fix <vbeclipse>: FilterIndex (1 基, 默认 1) —— OPENFILENAME.nFilterIndex 直通
VB6_CD_INT(FilterIndex, L"VB6_Cd_FilterIndex", 1)

#undef VB6_CD_STR
#undef VB6_CD_INT

// 自注册的不可见类。WndProc 什么都不做 —— 它只是属性袋，外加用 GetParent 拿模态父窗。
static LRESULT CALLBACK vb6_CdWndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    return DefWindowProcW(h, m, w, l);
}

// C29-T: Timer 的身份类（同样不可见、同样只当句柄用 —— 计时器真正的状态在
// vb6forms.c 的 g_timerTable 里，按这个句柄找回那一格）。
static LRESULT CALLBACK vb6_TimerWndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    return DefWindowProcW(h, m, w, l);
}

void vb6_RegisterTimerClass(void* hInstance) {
    static BOOL done = FALSE;
    if (done) return;
    WNDCLASSW wc;
    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc   = vb6_TimerWndProc;
    wc.hInstance     = (HINSTANCE)hInstance;
    wc.lpszClassName = L"VB6_TIMER";
    if (RegisterClassW(&wc)) done = TRUE;
}

void vb6_RegisterCommDialogClass(void* hInstance) {
    static BOOL done = FALSE;
    if (done) return;
    WNDCLASSW wc;
    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc   = vb6_CdWndProc;
    wc.hInstance     = (HINSTANCE)hInstance;
    wc.lpszClassName = L"VB6_COMMONDIALOG";
    if (RegisterClassW(&wc)) done = TRUE;
}

// VB6 "文本 (*.txt)|*.txt|所有文件 (*.*)|*.*" -> 原生成对表（每段以 \0 结束、整体再补一个 \0）
static void vb6_CdFoldFilter(const wchar_t* src, wchar_t* dst, size_t dstChars) {
    size_t o = 0;
    if (!src || !*src) { dst[0] = 0; dst[1] = 0; return; }
    const wchar_t* p = src;
    while (*p && o + 2 < dstChars) {
        const wchar_t* bar = wcschr(p, L'|');
        size_t len = bar ? (size_t)(bar - p) : wcslen(p);
        if (len > dstChars - o - 2) len = dstChars - o - 2;
        for (size_t i = 0; i < len; i++) dst[o++] = p[i];
        dst[o++] = 0;                       // 段结束
        if (!bar) break;                    // 落单的一段（VB6 要求描述与模式成对）
        p = bar + 1;
    }
    dst[o] = 0;                             // 表结束
}

static HWND vb6_CdOwner(void* hwnd) {
    HWND h = (HWND)hwnd;
    HWND p = h ? GetParent(h) : NULL;
    return p ? p : h;
}

static void vb6_CdCancel(void* hwnd) {
    if (vb6_CdGetCancelError(hwnd))
        vb6_RaiseError(32755, SysAllocString(L"Dialog was canceled by the user"));
}

// ai/029 C29-9b: 「起窗自关」探针 —— 让模态对话框这条判据能在无头环境里跑。
// 只有环境变量 C3_CDPROBE=1 才上膛，不设时这条路完全不存在（现有 ctrldlg 用例逐字不变）。
// 线程只认**本线程**创建的 #32770（EnumThreadWindows），所以不会去关别人的窗；发过去的
// WM_COMMAND/IDCANCEL 就是「用户点了取消」那条出口。于是「对话框真出现过」这条判据的断点
// 就是取消出口本身：框没出现 ⇒ 没人取消 ⇒ CancelError 不会报 32755。找不到就不动，
// 那种情况由套件的 -RunTimeoutSec 兜底（用例阶段最多 60 s）。
typedef struct { DWORD tid; } vb6_CdProbeArgs;

static volatile LONG vb6_CdProbeFired = 0;
static volatile LONG vb6_CdProbeStop = 0;

static BOOL CALLBACK vb6_CdProbeEnum(HWND h, LPARAM lp) {
    wchar_t cls[16] = {0};
    (void)lp;
    if (GetClassNameW(h, cls, 16) > 0 && wcscmp(cls, L"#32770") == 0) {
        PostMessageW(h, WM_COMMAND, (WPARAM)IDCANCEL, MAKELPARAM(BN_CLICKED, 0));
        InterlockedExchange(&vb6_CdProbeFired, 1);
        return FALSE;
    }
    return TRUE;
}

static DWORD WINAPI vb6_CdProbeProc(LPVOID arg) {
    vb6_CdProbeArgs a = *(vb6_CdProbeArgs*)arg;
    free(arg);
    for (int i = 0; i < 400; i++) {           // 最多等约 4 s
        BOOL found = FALSE;
        EnumThreadWindows(a.tid, vb6_CdProbeEnum, (LPARAM)&found);
        if (found || InterlockedCompareExchange(&vb6_CdProbeStop, 0, 0)) break;
        Sleep(10);
    }
    return 0;
}

static int vb6_CdProbeOn(void) {
    static int cached = -1;
    if (cached < 0) {
        wchar_t b[8];
        cached = (GetEnvironmentVariableW(L"C3_CDPROBE", b, 8) > 0) ? 1 : 0;
    }
    return cached;
}

// 返回句柄表示已上膛（调用方要 CloseHandle）；0 = 不上膛。
static HANDLE vb6_CdProbeArm(void) {
    if (!vb6_CdProbeOn()) return 0;
    InterlockedExchange(&vb6_CdProbeStop, 0);
    vb6_CdProbeArgs* a = (vb6_CdProbeArgs*)malloc(sizeof(vb6_CdProbeArgs));
    if (!a) return 0;
    a->tid = GetCurrentThreadId();
    DWORD id = 0;
    HANDLE h = CreateThread(NULL, 0, vb6_CdProbeProc, a, 0, &id);
    if (!h) free(a);
    return h;
}

// 撤膛必须真撤：上一发的线程若没命中，会一直轮到 4 s，可能把**下一发** Show* 的框关掉。
static void vb6_CdProbeDisarm(HANDLE h) {
    if (h) {
        InterlockedExchange(&vb6_CdProbeStop, 1);
        WaitForSingleObject(h, 2000);
        CloseHandle(h);
    }
}

static const wchar_t* vb6_CdBaseName(const wchar_t* full) {
    const wchar_t* s = wcsrchr(full, L'\\');
    if (!s) s = wcsrchr(full, L'/');
    return s ? s + 1 : full;
}

int vb6_CdShowFile(void* hwnd, int saveAs) {
    HMODULE m = vb6_ComDlgModule();
    if (!hwnd || !m) return 0;
    typedef BOOL (WINAPI *FnOFN)(LPOPENFILENAMEW);
    FnOFN pFn = (FnOFN)(void*)GetProcAddress(m, saveAs ? "GetSaveFileNameW" : "GetOpenFileNameW");
    if (!pFn) return 0;

    wchar_t filter[2048];
    wchar_t* filtVB = vb6_CdGetFilter(hwnd);
    vb6_CdFoldFilter(filtVB, filter, 2048);
    SysFreeString(filtVB);

    wchar_t file[1024] = {0};
    wchar_t title[1024] = {0};
    wchar_t initDir[1024] = {0};
    wchar_t defExt[64] = {0};
    wchar_t* s;
    s = vb6_CdGetFileName(hwnd);    wcsncpy_s(file, 1024, s, _TRUNCATE);  SysFreeString(s);
    s = vb6_CdGetDialogTitle(hwnd); wcsncpy_s(title, 1024, s, _TRUNCATE); SysFreeString(s);
    s = vb6_CdGetInitDir(hwnd);     wcsncpy_s(initDir, 1024, s, _TRUNCATE); SysFreeString(s);
    s = vb6_CdGetDefaultExt(hwnd);  wcsncpy_s(defExt, 64, s, _TRUNCATE);  SysFreeString(s);

    OPENFILENAMEW ofn;
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = vb6_CdOwner(hwnd);
    ofn.lpstrFilter = filter[0] ? filter : NULL;
    ofn.lpstrFile   = file;
    ofn.nMaxFile    = 1024;
    ofn.lpstrInitialDir = initDir[0] ? initDir : NULL;
    ofn.lpstrDefExt     = defExt[0] ? defExt : NULL;
    ofn.lpstrTitle      = title[0] ? title : (saveAs ? L"Save As" : L"Open");
    ofn.Flags = (DWORD)vb6_CdGetFlags(hwnd) | OFN_EXPLORER | OFN_HIDEREADONLY;
    if (saveAs) ofn.Flags |= OFN_OVERWRITEPROMPT;
    ofn.nFilterIndex = (DWORD)vb6_CdGetFilterIndex(hwnd);

    HANDLE probe = vb6_CdProbeArm();
    int shown = pFn(&ofn);
    vb6_CdProbeDisarm(probe);
    if (!shown) { vb6_CdCancel(hwnd); return 0; }
    vb6_CdSetStr(hwnd, L"VB6_Cd_FileName", file);
    vb6_CdSetStr(hwnd, L"VB6_Cd_FileTitle", vb6_CdBaseName(file));
    return 1;
}

int vb6_CdShowOpen(void* hwnd) { return vb6_CdShowFile(hwnd, 0); }
int vb6_CdShowSave(void* hwnd) { return vb6_CdShowFile(hwnd, 1); }

int vb6_CdShowColor(void* hwnd) {
    HMODULE m = vb6_ComDlgModule();
    if (!hwnd || !m) return 0;
    typedef BOOL (WINAPI *FnCC)(LPCHOOSECOLORW);
    FnCC pFn = (FnCC)(void*)GetProcAddress(m, "ChooseColorW");
    if (!pFn) return 0;
    static COLORREF g_cCustom[16] = {0};
    CHOOSECOLORW cc;
    ZeroMemory(&cc, sizeof(cc));
    cc.lStructSize  = sizeof(cc);
    cc.hwndOwner    = vb6_CdOwner(hwnd);
    cc.rgbResult    = (COLORREF)vb6_CdGetColor(hwnd);
    cc.lpCustColors = g_cCustom;
    cc.Flags        = (DWORD)vb6_CdGetFlags(hwnd) | CC_ANYCOLOR | CC_RGBINIT;
    HANDLE probe = vb6_CdProbeArm();
    int shown = pFn(&cc);
    vb6_CdProbeDisarm(probe);
    if (!shown) { vb6_CdCancel(hwnd); return 0; }
    vb6_CdSetColor(hwnd, (int)cc.rgbResult);
    return 1;
}

int vb6_CdShowFont(void* hwnd) {
    HMODULE m = vb6_ComDlgModule();
    if (!hwnd || !m) return 0;
    typedef BOOL (WINAPI *FnCF)(LPCHOOSEFONTW);
    FnCF pFn = (FnCF)(void*)GetProcAddress(m, "ChooseFontW");
    if (!pFn) return 0;
    LOGFONTW lf;
    ZeroMemory(&lf, sizeof(lf));
    wchar_t face[64] = {0};
    wchar_t* s = vb6_CdGetFontName(hwnd);
    wcsncpy_s(face, 64, s, _TRUNCATE);
    SysFreeString(s);
    wcscpy_s(lf.lfFaceName, (size_t)_countof(lf.lfFaceName), face);

    CHOOSEFONTW cf;
    ZeroMemory(&cf, sizeof(cf));
    cf.lStructSize = sizeof(cf);
    cf.hwndOwner   = vb6_CdOwner(hwnd);
    cf.lpLogFont   = &lf;
    cf.iPointSize  = vb6_CdGetFontSize(hwnd) * 10;
    cf.Flags       = (DWORD)vb6_CdGetFlags(hwnd) | CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT;
    HANDLE probe = vb6_CdProbeArm();
    int shown = pFn(&cf);
    vb6_CdProbeDisarm(probe);
    if (!shown) { vb6_CdCancel(hwnd); return 0; }
    vb6_CdSetStr(hwnd, L"VB6_Cd_FontName", lf.lfFaceName);
    vb6_CdSetFontSize(hwnd, (int)(cf.iPointSize / 10));
    return 1;
}

int vb6_CdShowPrinter(void* hwnd) {
    HMODULE m = vb6_ComDlgModule();
    if (!hwnd || !m) return 0;
    typedef BOOL (WINAPI *FnPD)(LPPRINTDLGW);
    FnPD pFn = (FnPD)(void*)GetProcAddress(m, "PrintDlgW");
    if (!pFn) return 0;
    PRINTDLGW pd;
    ZeroMemory(&pd, sizeof(pd));
    pd.lStructSize = sizeof(pd);
    pd.hwndOwner   = vb6_CdOwner(hwnd);
    pd.Flags       = (DWORD)vb6_CdGetFlags(hwnd) | PD_RETURNDC;
    pd.nCopies     = (WORD)vb6_CdGetCopies(hwnd);
    pd.nFromPage   = (WORD)vb6_CdGetMin(hwnd);
    pd.nToPage     = (WORD)vb6_CdGetMax(hwnd);
    HANDLE probe = vb6_CdProbeArm();
    int shown = pFn(&pd);
    vb6_CdProbeDisarm(probe);
    if (!shown) { vb6_CdCancel(hwnd); return 0; }
    if (pd.hDC) DeleteDC(pd.hDC);           // v1 不把 DC 交给用户（Printer 对象另立批次）
    vb6_CdSetCopies(hwnd, (int)pd.nCopies);
    return 1;
}

int vb6_CdShowAbout(void* hwnd) {
    HMODULE m = vb6_ComDlgModule();
    if (!hwnd || !m) return 0;
    typedef BOOL (WINAPI *FnSA)(HWND, LPCWSTR, LPCWSTR, HICON);
    FnSA pFn = (FnSA)(void*)GetProcAddress(m, "ShellAboutW");
    if (!pFn) return 0;
    wchar_t* t = vb6_CdGetDialogTitle(hwnd);
    BOOL ok = pFn(vb6_CdOwner(hwnd), (t && *t) ? t : (const wchar_t*)L"About",
                  (const wchar_t*)L"", NULL);
    SysFreeString(t);
    if (!ok) vb6_CdCancel(hwnd);
    return ok ? 1 : 0;
}
