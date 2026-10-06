#pragma once
// vb6forms_prop_form.h - 窗体属性 (P20-40)、Label (P20-39)、Menu (P20-36)、按钮 Style (P20-38)、Shape/Line (P20-34/35)、PictureBox 图形属性 (P20-42)
// 由 vb6forms.h 伞头按固定顺序 include，不要单独使用
// 内容 = 拆分前 vb6forms.h 第 379~449 行，逐行未改

#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// Form属性 (P20-40)
int32_t vb6_GetKeyPreview(void* hwnd);
void vb6_SetKeyPreview(void* hwnd, int32_t val);
int32_t vb6_GetWindowState(void* hwnd);
void vb6_SetWindowState(void* hwnd, int32_t val);
int32_t vb6_GetScaleWidth(void* hwnd);
int32_t vb6_GetScaleHeight(void* hwnd);
int32_t vb6_GetControlBox(void* hwnd);
void vb6_SetControlBox(void* hwnd, int32_t val);
int32_t vb6_GetMaxButton(void* hwnd);
void vb6_SetMaxButton(void* hwnd, int32_t val);
int32_t vb6_GetMinButton(void* hwnd);
void vb6_SetMinButton(void* hwnd, int32_t val);
// Label属性 (P20-39)
int32_t vb6_GetLabelAutoSize(void* hwnd);
void vb6_SetLabelAutoSize(void* hwnd, int32_t val);
int32_t vb6_GetLabelWordWrap(void* hwnd);
void vb6_SetLabelWordWrap(void* hwnd, int32_t val);
int32_t vb6_GetLabelBackStyle(void* hwnd);
void vb6_SetLabelBackStyle(void* hwnd, int32_t val);
// Menu属性 (P20-36)
void* vb6_GetMenuCaption(void* menuHandle, int menuId);
void vb6_SetMenuCaption(void* menuHandle, int menuId, void* bstrCaption);
int32_t vb6_GetMenuChecked(void* menuHandle, int menuId);
void vb6_SetMenuChecked(void* menuHandle, int menuId, int32_t val);
int32_t vb6_GetMenuEnabled(void* menuHandle, int menuId);
void vb6_SetMenuEnabled(void* menuHandle, int menuId, int32_t val);
int32_t vb6_GetMenuVisible(void* menuHandle, int menuId);
void vb6_SetMenuVisible(void* menuHandle, int menuId, int32_t val);
// 按钮Style属性 (P20-38)
int32_t vb6_GetButtonStyle(void* hwnd);
void vb6_SetButtonStyle(void* hwnd, int32_t val);
// Shape属性 (P20-34)
int32_t vb6_GetShapeType(void* hwnd);
void vb6_SetShapeType(void* hwnd, int32_t val);
int32_t vb6_GetShapeBorderWidth(void* hwnd);
void vb6_SetShapeBorderWidth(void* hwnd, int32_t val);
int32_t vb6_GetShapeBorderStyle(void* hwnd);
void vb6_SetShapeBorderStyle(void* hwnd, int32_t val);
int32_t vb6_GetShapeFillStyle(void* hwnd);
void vb6_SetShapeFillStyle(void* hwnd, int32_t val);
int32_t vb6_GetShapeBorderColor(void* hwnd);
void vb6_SetShapeBorderColor(void* hwnd, int32_t val);
int32_t vb6_GetShapeFillColor(void* hwnd);
void vb6_SetShapeFillColor(void* hwnd, int32_t val);
// Shape/Line窗口类注册 (P20-35)
void vb6_RegisterShapeLineClasses(void* hInstance);
// Line属性 (P20-35)
int32_t vb6_GetLineX1(void* hwnd);
void vb6_SetLineX1(void* hwnd, int32_t val);
int32_t vb6_GetLineY1(void* hwnd);
void vb6_SetLineY1(void* hwnd, int32_t val);
int32_t vb6_GetLineX2(void* hwnd);
void vb6_SetLineX2(void* hwnd, int32_t val);
int32_t vb6_GetLineY2(void* hwnd);
void vb6_SetLineY2(void* hwnd, int32_t val);
int32_t vb6_GetLineBorderWidth(void* hwnd);
void vb6_SetLineBorderWidth(void* hwnd, int32_t val);
int32_t vb6_GetLineBorderStyle(void* hwnd);
void vb6_SetLineBorderStyle(void* hwnd, int32_t val);
int32_t vb6_GetLineColor(void* hwnd);
void vb6_SetLineColor(void* hwnd, int32_t val);
// PictureBox图形属性 (P20-42)
int32_t vb6_GetAutoRedraw(void* hwnd);
void vb6_SetAutoRedraw(void* hwnd, int32_t val);
int32_t vb6_GetScaleMode(void* hwnd);
void vb6_SetScaleMode(void* hwnd, int32_t val);
// 账 #196: 控件的 `.hDC` —— 见 vb6forms_ctrl.c 那条注释：两档来源（_Paint 派发期用外层
// 挂上的 VB6_PaintDC，否则窗口 DC），且**一个对象一张**，反复读回同一个句柄，销毁时归还。
intptr_t vb6_GetControlHDC(void* hwnd);
// 账 #196: 按 HWND 的文字量（TextHeight / TextWidth）。量这枚控件自己的字体、交这枚控件
// 自己的 ScaleMode 单位 —— 见 vb6forms_ctrl.c 那一段（DC 与 Print 同一处权威）。
float vb6_ControlTextWidth(void* hwnd, void* bstrText);
float vb6_ControlTextHeight(void* hwnd, void* bstrText);
float vb6_GetCurrentX(void* hwnd);
void vb6_SetCurrentX(void* hwnd, float val);
float vb6_GetCurrentY(void* hwnd);
void vb6_SetCurrentY(void* hwnd, float val);

#ifdef __cplusplus
}
#endif
