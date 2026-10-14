// vb6forms_internal.h - vb6forms 模块内部共享声明 (P7 Win32 窗体运行时)
// 仅供 RTL 内部各 .c 使用; 生成代码只 include vb6forms.h
//
// 拆分说明: vb6forms.c 按控件/窗体功能家族拆为多个编译单元, 各单元由 MSVC
// 独立编译成 .obj 再链接, 文件级 static 不跨文件可见 —— 因此把跨族共享的
// 符号集中声明在此, 定义仍留在 vb6forms.c。

#ifndef VB6C3_VB6FORMS_INTERNAL_H
#define VB6C3_VB6FORMS_INTERNAL_H

#include "vb6forms.h"

#include <stdlib.h>   /* malloc (Fix 190 转码助手) */
#include <string.h>   /* memset (下面的 VB Date 换算助手) */
#include <oleauto.h>  /* VariantTimeToSystemTime / SystemTimeToVariantTime (VB Date 换算助手) */

// ============================================================
// VB Date <-> 原生 SYSTEMTIME 的共用换算（C29-DT-b 先放在 dtpicker 里，
// C29-MV-b 的 MonthView 是第二个用户 ⇒ 挪到这里共享；放法照上面那组
// static inline 助手 —— 文件级 static 不跨编译单元可见）
//
// VB 的 Date 在 C3 里就是 double 序列号（`Dim d As Date` 发成 `double d`、
// `Now` 直接回 double），所以两族控件的属性签名一律 double <-> SYSTEMTIME，
// 换算交给 oleaut32 那一对现成函数（RTL 里的直接调用先例：vb6rtl_format.c:91）。
//
// ⚠ 出界或换算失败一律回 **0**，不回负数：负数在 VB 侧是非法 Date，
// 那会把"问不出值"伪装成"一个怪值"，判据就再也分不出这两种情况。
// ============================================================
static inline void vb6_DateZero(SYSTEMTIME* st) { memset(st, 0, sizeof(*st)); }

static inline double vb6_DateToSerial(const SYSTEMTIME* st) {
    double v = 0.0;
    if (!st) return 0.0;
    if (!SystemTimeToVariantTime((LPSYSTEMTIME)st, &v) || v < 0.0) return 0.0;
    return v;
}

static inline int vb6_DateFromSerial(double serial, SYSTEMTIME* st) {
    if (!st) return 0;
    vb6_DateZero(st);
    return VariantTimeToSystemTime(serial, st) ? 1 : 0;
}

// --- 跨族共享的内部状态 (定义在 vb6forms.c) ---
// 应用实例句柄 (vb6_SetAppInstance 设置, 多处属性设置与控件创建需要)
extern HINSTANCE vb6_hInstance;

// --- 控件窗口的字体 (账 #200/#202/#204, 定义在 vb6forms_ctrl.c) ---
// 「这枚控件现在在用的字体」在本仓库只许有一处问法：先问窗口，窗口不答再读我们自存的那份。
// 必须跨文件共享的理由是**窗口类本身** —— 探针实测裸 STATIC 与裸 BUTTON(BS_GROUPBOX)
// 收到 `WM_SETFONT` 之后都不答 `WM_GETFONT`（`.build/b200probe/fontprobe.c`），
// 所以凡是"自己读窗口字体"的站点（Frame 的标题带、控件数组问模板）拿到的恒是 NULL，
// 症状不是崩而是静默按 DC 的默认字体画/量（账 #200 的 TextHeight、#202 的白带、#204 的整张 Font 面）。
// `vb6_ControlFontStore` 是**唯一**写那份自存的出口：谁把字体发给窗口，谁就同时经它存一份
// （创建期那一站尤其要紧 —— 从没被写过字体的控件今天连 `.FontName` 都读空，见账 #204）。
HFONT vb6_ControlFont(HWND hwnd);
void vb6_ControlFontStore(HWND hwnd, HFONT hFont);

// ============================================================
// Fix 190: 源码字符串 = UTF-8, 窗口层 = UTF-16
//
// C3 内部统一 UTF-8 (.frm/.bas 读入即转码), 生成的 C 也带 /utf-8 编译, 所以
// 生成代码传进来的**每一个字面量**(窗体 Caption、控件 Caption/Text、.frx
// 文本、控件名 …)都是 UTF-8 字节, 而不是系统 ACP。
//
// 旧实现把这些字节直接交给 CreateWindowExA/SetWindowTextA, 系统按 ACP 解释:
// CP936 机器上 "对比数据" 变成 "瀵规瘮鏁版嵁"; 韩/德/俄文更糟。正确做法是
// 窗口层全程走 W 系列: UTF-8 → UTF-16 只在边界转一次, 之后与语言无关。
//
// 放在头里做成 static inline, 供 vb6forms 各拆分单元共享 (文件级 static 不跨
// 编译单元可见)。
// ============================================================
static inline wchar_t* vb6_u8ToWideDup(const char* s) {
    if (!s) return NULL;
    int n = MultiByteToWideChar(CP_UTF8, 0, s, -1, NULL, 0);
    if (n <= 0) {
        /* 非法 UTF-8: 退回 ACP, 保证调用方拿到可用指针 */
        n = MultiByteToWideChar(CP_ACP, 0, s, -1, NULL, 0);
        if (n <= 0) return NULL;
        wchar_t* wa = (wchar_t*)malloc((size_t)n * sizeof(wchar_t));
        if (wa) MultiByteToWideChar(CP_ACP, 0, s, -1, wa, n);
        return wa;
    }
    wchar_t* w = (wchar_t*)malloc((size_t)n * sizeof(wchar_t));
    if (w) MultiByteToWideChar(CP_UTF8, 0, s, -1, w, n);
    return w;
}

/* 栈缓冲版: 供类名/标题等短字符串使用 (超长即截断, 不分配) */
static inline void vb6_u8ToWideBuf(const char* s, wchar_t* out, int cap) {
    if (!out || cap <= 0) return;
    out[0] = 0;
    if (!s) return;
    if (MultiByteToWideChar(CP_UTF8, 0, s, -1, out, cap) <= 0)
        MultiByteToWideChar(CP_ACP, 0, s, -1, out, cap);
    out[cap - 1] = 0;
}

/* 宽 → UTF-8 (栈缓冲): 窗口层读回的 Unicode 文本转回 RTL 内部的 UTF-8 键 */
static inline void vb6_wideToU8Buf(const wchar_t* w, char* out, int cap) {
    if (!out || cap <= 0) return;
    out[0] = 0;
    if (!w) return;
    if (WideCharToMultiByte(CP_UTF8, 0, w, -1, out, cap, NULL, NULL) <= 0)
        WideCharToMultiByte(CP_ACP, 0, w, -1, out, cap, NULL, NULL);
    out[cap - 1] = 0;
}

// ============================================================
// 绘图 DC 的唯一取法 (定义在 vb6forms_ctrl.c)。
// 账 #185/#196 把「这枚窗口的绘图 DC 从哪儿来」收成一处: WM_PAINT 派发期用宿主
// BeginPaint 后挂在窗口属性 VB6_PaintDC 上的那张, 否则 GetDC。*pFromPaint=TRUE 就
// 意味着这张**不许** ReleaseDC (派发期那张由宿主的 EndPaint 收尾)。
// 账 #234 起绘图方法家族 (vb6forms_draw.c 的 PSet/Line/Circle/Point/Cls) 也问它 ——
// 那里原本自己又写了一份同样口径, 而 check_control_dc.ps1 的名单扫不到那个文件。
// ============================================================
HDC vb6_ControlDrawDC(HWND hw, BOOL* pFromPaint);

#endif // VB6C3_VB6FORMS_INTERNAL_H
