// vb6forms_draw.h - Form / Printer 绘图方法家族声明（PSet/Line/Circle/Point/Cls）
// 2026-10-06 新增，与 vb6forms_draw.c 一一对应。
//
// 为什么单开一头：Form/Printer 的绘图面此前只有 Shape/Line 控件那条自绘路径
// （vb6forms_widget.c），用户代码在 Form_Paint 里写 Me.PSet/Line/Circle/Point 时
// 落进 COM dispatch 桩（运行时 no-op）。声明单独成头是为了与既有
// vb6forms_prop_pic.h / vb6forms_prop_form.h 的分家族口径一致。
//
// HDC 口径见 vb6forms_draw.c 文件头：优先取窗口属性 "VB6_PaintDC"（WM_PAINT
// 派发期宿主 BeginPaint 过的那张），否则 GetDC。**不新建 DC**。
#ifndef VB6FORMS_DRAW_H
#define VB6FORMS_DRAW_H

// ⚠ **不要**在这里 include "vb6forms.h": 它在末尾 include 本头, 构成循环。
// 生成的 .c/.h 已经先 include 了 vb6forms.h (那里会带齐 windows.h 与前序家族),
// 所以本头只依赖 int32_t/HWND 这类已在前序可见的类型。
// 循环的实际后果不是"编不过", 而是**每个 include 它的编译单元都拿到一份**,
// 链接期表现为 LNK2005 "vb6_Form_Circle 已经在 draw.obj 中定义" (用户代码 obj)
// —— 措辞完全误导: 真因是 RTL 侧那份 obj 与用户 obj 同时提供了符号。
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// ---- Form 版 (hwnd = 窗体句柄) ----
// 形参按 VB6 语义打平: step/hasXY 表达"Step 关键字"与"坐标是否省略"，
// hasColor 表达颜色实参是否给了 (没给 = 用 ForeColor)。
void   vb6_Form_PSet(void* hwnd, int32_t step, int32_t hasXY,
                     double x, double y, int32_t hasColor, int32_t color);
int32_t vb6_Form_Point(void* hwnd, double x, double y);
void   vb6_Form_Line(void* hwnd,
                     int32_t step1, int32_t has1, double x1, double y1,
                     int32_t step2, int32_t has2, double x2, double y2,
                     int32_t hasColor, int32_t color,
                     int32_t style);
void   vb6_Form_Circle(void* hwnd,
                       int32_t step, int32_t hasXY, double x, double y,
                       double radius,
                       int32_t hasColor, int32_t color,
                       int32_t hasStart, double startAngle,
                       int32_t hasEnd, double endAngle,
                       int32_t hasAspect, double aspect);
void   vb6_Form_Cls(void* hwnd);

// ---- Form 绘图状态读写 (cgen 的 Me.CurrentX / .ForeColor / .DrawWidth) ----
// 笔位不在这张表里 —— 账 #233: CurrentX/Y 只有一份存储, cgen 直接登记
// vb6_GetCurrentX / vb6_SetCurrentX (vb6forms_prop_form.h)。
int32_t vb6_Form_DrawGetWidth(void* hwnd);
void   vb6_Form_DrawSetWidth(void* hwnd, int32_t w);

// ---- Printer 版 (走全局 g_printerDC, 无 HWND 可挂属性 ⇒ 状态用静态变量) ----
void   vb6_Printer_PSet(int32_t step, int32_t hasXY, double x, double y,
                        int32_t hasColor, int32_t color);
int32_t vb6_Printer_Point(double x, double y);
void   vb6_Printer_Line(int32_t step1, int32_t has1, double x1, double y1,
                        int32_t step2, int32_t has2, double x2, double y2,
                        int32_t hasColor, int32_t color,
                        int32_t style);
void   vb6_Printer_Circle(int32_t step, int32_t hasXY, double x, double y,
                          double radius, int32_t hasColor, int32_t color,
                          int32_t hasStart, double startAngle,
                          int32_t hasEnd, double endAngle,
                          int32_t hasAspect, double aspect);
int32_t vb6_Printer_DrawGetCurrentX(void);
int32_t vb6_Printer_DrawGetCurrentY(void);
void   vb6_Printer_DrawSetCurrentX(int32_t x);
void   vb6_Printer_DrawSetCurrentY(int32_t y);
int32_t vb6_Printer_DrawGetForeColor(void);
void   vb6_Printer_DrawSetForeColor(int32_t c);
int32_t vb6_Printer_DrawGetWidth(void);
void   vb6_Printer_DrawSetWidth(int32_t w);

#ifdef __cplusplus
}
#endif

#endif /* VB6FORMS_DRAW_H */
