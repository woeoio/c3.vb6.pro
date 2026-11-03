#pragma once
// vb6forms_mdi.h - MDI 窗体 (P7.7)：父/子窗体、排列、激活
// 由 vb6forms.h 伞头按固定顺序 include，不要单独使用
// 内容 = 拆分前 vb6forms.h 第 489~521 行，逐行未改

#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// MDI窗体 (P7.7)
// ============================================================
// VB6 MDIForm (父窗体) + MDIChild=True的子窗体
// Win32实现: MDI客户窗口(MDICLIENT) + CreateMDIWindow

// 注册MDI父窗体窗口类
int vb6_RegisterMDIFormClass(const char* className, void* wndProc, void* hInstance, int iconResId);

// 创建MDI父窗体 (包含MDICLIENT子窗口)
void* vb6_CreateMDIFormWindow(const char* className, const char* formName,
    int x, int y, int width, int height, void* hInstance);

// 创建MDI子窗体
void* vb6_CreateMDIChildWindow(const char* className, const char* formName,
    int x, int y, int width, int height, void* hMDIClient, void* hInstance);

// 获取MDI客户窗口句柄 (从MDI父窗体获取)
void* vb6_GetMDIClient(void* hMDIForm);
/* 账 #305 第三百四十六刀: MDIClient 收进 Codejock 停靠条的空余带 (WM_SIZE 钩) */
void vb6_MDI_RelayoutClient(void* hMDIForm);

/* 账 (MDI 布局刀): MDI 父窗体 Show 之后补放子窗的设计期 WindowState。
 * 子窗在 MDIForm_Load 里 Show 时父窗体还没上屏, MDICLIENT 的 WM_MDIMAXIMIZE
 * 对"客户区不可见"的子窗不做事 —— 那时只把目标状态记在子窗属性上, 由这里落地。 */
void vb6_MDIApplyPendingChildStates(void* hMDIClient);

/* 账 (frmDocument 最大化): `CommandBars.ShowTabWorkspace(True)` 在 VB6 下返回真的
 * TabWorkspace, 活动 MDI 子窗被工作区铺满; C3 下 OCX 回空指针 (实测 pdisp=0)。
 * vb6com 的晚绑定出口在方法名 == ShowTabWorkspace 且结果为空时打这个标记,
 * vb6_ShowForm 据此按"工作区行为"把活动子窗最大化 —— 这是**替代路径**。 */
void vb6_MDI_NoteTabWorkspaceMissing(void);
int  vb6_MDI_TabWorkspaceMissing(void);

// MDI消息循环 (处理TranslateMDISysAccel)
int vb6_MDIMessageLoop(void* hAccelTable);

// MDI窗口排列
void vb6_MDITile(void* hMDIClient, int style);     // style: 0=horizontal, 1=vertical
void vb6_MDICascade(void* hMDIClient);
void vb6_MDIArrangeIcons(void* hMDIClient);

// MDI子窗体管理
void* vb6_MDIGetActive(void* hMDIClient);           // 获取活动子窗体
void vb6_MDIActivate(void* hMDIClient, void* hChild); // 激活子窗体



#ifdef __cplusplus
}
#endif
