#pragma once
// vb6rtl_userctl.h - UserControl/PropertyPage 宿主内建对象 (Fix 105)
//
// 背景: VB6 UserControl 工程 (Charts 2020 等) 的 .ctl/.pag 代码里直接引用
// UserControl.* / Ambient.* / Extender.* / PropertyPage.* 内建对象。
// 生成 C 把它们发射为裸标识符 vb6_UserControl_ScaleWidth / vb6_Ambient_Font /
// vb6_Extender_Left / vb6_PropertyPage_hwnd 等, 此前无任何声明 → C2065。
//
// 本头提供「每进程单实例」的宿主状态: 对单窗体 + 同屏多实例不精确 (多实例
// 共享一份宿主状态), 但让真实工程先可编译运行; 实例化宿主是后续特性块。
// 同时提供 StdFont 具体类 (生成代码把 UserControl.Font 强转为 vb6_cls_StdFont*
// 做 With 成员写) 与 Picture1.Line 的 B/BF 模式常量、HitResult 常量。

#include "vb6rtl_bstr.h"

#ifdef __cplusplus
extern "C" {
#endif

// --- Font 对象 (VB6 stdole.StdFont 的最小 C 形态) ---
// Fix 109: 生成代码对 UserControl.Font 既有对象用法 (裸名作为 vb6_ComIface_Font*
// 实参传给控件绘制函数), 又有成员用法 (`.Bold` / `.Size + 8` / `.Name`)。
// 因此这里给出**完整结构体**, 并把 vb6_ComIface_Font 定义为它 —— 生成代码对该
// 类型只有前向 typedef, 这里补齐定义不冲突. 成员用数值/字符串真实类型, 与
// 生成代码 bBold(int16_t) = Font.Bold; Font.Bold = -1; Font.Size + 8 一致.
typedef struct vb6_ComIface_Font {
    BSTR    Name;
    float   Size;
    int16_t Bold;
    int16_t Italic;
    int16_t Underline;
    int16_t Strikethrough;
    int32_t Weight;
    int32_t Charset;
} vb6_ComIface_Font;

// 旧名保留 (Fix 105 引入), 供既有生成代码/注释引用.
typedef vb6_ComIface_Font vb6_cls_StdFont;

// --- UserControl 宿主 ---
extern int32_t vb6_UserControl_ScaleWidth;   // 用户坐标宽度
extern int32_t vb6_UserControl_ScaleHeight;  // 用户坐标高度
extern int32_t vb6_UserControl_ScaleMode;    // 1=Twip 3=Pixel
extern void*   vb6_UserControl_hDC;          // 绘制 DC (Windowless: 容器客户区 DC)
// Fix <vbeclipse>: UserControl.HasDC (ucTab.ctl:89 `HasDC = UserControl.HasDC`).
// VB6 语义 = 当前是否有可用绘制 DC; 生成代码按**变量**读 (vb6_ret_HasDC = ...),
// 故用宏而不是函数, 与 hDC 的赋值点天然同步 (无需在 uc_host.c 各赋值处维护).
#define vb6_UserControl_HasDC (vb6_UserControl_hDC ? 1 : 0)
// 账 #180 (B19): 容器 HWND 必须是**指针宽度**。此前它是 int32_t, 而写入点递进来的是
// HWND (x64 = 64 位) ⇒ 高 32 位当场丢掉; 语料里 VBFlexGrid.ctl 把它直接当 HWND 传给
// MapWindowPoints / GetWindowLongW (实测产物 5 处), 那是把截断后的值交给窗口管理器。
// 与 vb6_UserControl_hWnd / _hDC 同档 (两者本来就是 void*)。
extern void*   vb6_UserControl_ContainerHwnd;// 容器 HWND (指针宽度)
extern int16_t vb6_UserControl_Enabled;
extern int32_t vb6_UserControl_MousePointer;
extern void*   vb6_UserControl_MouseIcon;
extern int32_t vb6_UserControl_OLEDropMode;
// Fix 154-B: UserControl 其余宿主属性 (VBFlexGrid.ctl 设计期/宿主状态成员).
// 生成代码经 M22 host-pseudo-object 路径发射为 `vb6_UserControl_<Prop>`;
// 无容器运行期读空/写 no-op 即可满足编译 (Standard EXE 里 UserControl 由宿主
// 驱动, 这些成员的真实值来自容器). 命名大小写与 cIdent 保持一致.
extern void*   vb6_UserControl_Picture;        // IPictureDisp* (Set 读写)
extern int32_t vb6_UserControl_BackColor;      // OLE_COLOR
extern int32_t vb6_UserControl_ForeColor;      // OLE_COLOR
extern int16_t vb6_UserControl_RightToLeft;    // TriState: 0/1/-1
extern void*   vb6_UserControl_ParentControls; // Controls 集合 (For Each)
extern void*   vb6_UserControl_Controls;       // Fix <vbeclipse>: UserControl.Controls (未建模, NULL)
// Fix 109: Font 是**对象指针** (生成代码把它作为 vb6_ComIface_Font* 实参传递,
// 同时对它做 `.成员` 访问 —— 成员访问由生成端改写为 '->', 见 CodeEmitter::emitLine).
extern vb6_ComIface_Font* vb6_UserControl_Font;
struct vb6_UserControl_Ambient_Type { vb6_ComIface_Font* Font; };
extern struct vb6_UserControl_Ambient_Type vb6_UserControl_Ambient;

void vb6_UserControl_Refresh(void);
// Fix <vbeclipse>: `UserControl.Line (x1,y1)-(x2,y2), [color], [mode]`
// (ucTab.ctl:231-241). 生成端把 `-` 连写的坐标对拍平为 5 个固定实参, 末参
// `, B` / `, BF` 模式常量按 Fix 102 口径原样追加 → 需变参承接 (5 参与 6 参
// 两种形态在源码里都存在). 定义见 vb6rtl_com.c (no-op, 重绘路径负责最终成像).
void vb6_UserControl_Line(double x1, double y1, double x2, double y2, int32_t color, ...);
// Fix 110z: UserControl.Size width, height — 设置控件尺寸 (单位同 ScaleMode).
// Charts 2020 LabelPlus.ctl:1371 `UserControl.Size (lWidth + 1) * ..., ...`
// 生成 vb6_UserControl_Size(w, h), 此前无声明 → C2065/C2064.
void vb6_UserControl_Size(double width, double height);
void vb6_UserControl_CancelAsyncRead(BSTR propName);
int32_t vb6_UserControl_TextWidth(BSTR text);
int32_t vb6_UserControl_TextHeight(BSTR text);

// Fix 111: UserControl 的其余内建方法 (.ctl 里常以裸名书写,
// 生成 C 端由 cgen 的宿主伪对象成员表映射为 vb6_UserControl_<Member>).
//   ScaleX/ScaleY(x, fromScale, toScale) → 单位换算
//   AsyncRead(url, asyncType, propertyName, flags) → 异步读取 (编译形态下空操作)
//   PropertyChanged(propName) → 通知容器属性已变 (触发容器端的 Changed/属性刷新)
double vb6_UserControl_ScaleX(double x, int32_t fromScale, int32_t toScale);
double vb6_UserControl_ScaleY(double x, int32_t fromScale, int32_t toScale);

// 账 #196 第三条: 上面那一对只是**转手**到这里 —— 单位换算的实现只有一份，名字不带宿主前缀，
// 因为 PictureBox / 窗体 / `Me.` / With 块里那一枚控件 / 窗体模块里裸写 这四形接收者要的是同一件事。
// 声明留在本头是因为换算的声明本来就住在这儿，而生成 C 只 include vb6rtl.h (本头由它带进来)。
double vb6_ScaleUnitX(double x, int32_t fromScale, int32_t toScale);
double vb6_ScaleUnitY(double y, int32_t fromScale, int32_t toScale);
void   vb6_UserControl_AsyncRead(BSTR url, int32_t asyncType, BSTR propertyName,
                                 int32_t flags);
void   vb6_UserControl_PropertyChanged(BSTR propName);

// Fix 133u: czUI.ctl 用到的其余 UserControl 宿主成员.
//   UserControl.hWnd    → extern void*  (控件自身窗口句柄, push 时填充)
//   UserControl.AutoRedraw → extern int16_t (重绘模式; czUI 仅判断)
//   UserControl.Cls()   → 清空控件客户区背景 (下轮 WM_PAINT 重绘)
extern void*   vb6_UserControl_hWnd;
extern int16_t vb6_UserControl_AutoRedraw;
void   vb6_UserControl_Cls(void);

// Fix 133u: UserControl.Parent (容器窗体对象). czUI.ctl 用它做全屏/恢复窗体:
//   UserControl.Parent.hWnd       → vb6_UC_ParentHwnd(void)
//   UserControl.Parent.Icon.Handle→ vb6_UC_ParentIconHandle(void)
//   UserControl.Parent.Move l,t,w,h → vb6_UC_ParentMove(l,t,w,h)
//   With UserControl.Parent       → vb6_UC_ParentObject(void) (返回窗体 HWND,
//                                   With 内 .Left/.Top/.Width/.Height 经
//                                   vb6_ComGetIntProp 以窗体 HWND 解析)
// 由 cgen 生成端文本重写 (CodeEmitter::emitLine, Fix 133u) 把生成 C 里的
// "vb6_UserControl_Parent.<成员>" 链改写为这些函数 — 见 cgen_base.cpp.
void*   vb6_UC_ParentObject(void);
void*   vb6_UC_ParentHwnd(void);
void*   vb6_UC_ParentIconHandle(void);
void    vb6_UC_ParentMove(int32_t left, int32_t top, int32_t width, int32_t height);

// --- Ambient 宿主环境 ---
extern vb6_ComIface_Font* vb6_Ambient_Font;  // 容器默认 Font
extern int16_t vb6_Ambient_UserMode;         // -1=运行期 0=设计期
extern BSTR    vb6_Ambient_DisplayName;      // 控件实例名
extern int32_t vb6_Ambient_ForeColor;
extern int32_t vb6_Ambient_BackColor;
extern int16_t vb6_Ambient_RightToLeft;   // Fix 154-B: TriState 从环境属性

// --- Extender (容器提供的扩展对象) ---
extern int32_t vb6_Extender_Left;
extern int32_t vb6_Extender_Top;
// Fix 153-B: UserControl 的 Extender 属性转发样板 (VB6 .ctl 里
//   `X = Extender.X` / `Extender.X = Value` 形式). 设计期对象, 运行期
// (Standard EXE) 恒不活动, 故只需类型正确的可读写全局槽位. 命名与
// cgen 的 host-pseudo-object M22 发射一致 (`Extender.Tag` → vb6_Extender_Tag).
extern int32_t vb6_Extender_Width;
extern int32_t vb6_Extender_Height;
// Fix 162: cgen 对宿主伪对象按 `<对象名>_<成员>` 发射, 故 `UserControl.Width` 与
// `Extender.Width` 是两个不同的标识符, 而 VB6 里是同一个值 (控件外框, 容器缇)。
// 两者由 vb6_uc_push 一起同步。
extern int32_t vb6_UserControl_Width;
extern int32_t vb6_UserControl_Height;
extern int16_t vb6_Extender_Visible;
extern BSTR    vb6_Extender_Tag;
extern BSTR    vb6_Extender_ToolTipText;
extern int32_t vb6_Extender_DragMode;
extern int32_t vb6_Extender_Align;
extern int32_t vb6_Extender_HelpContextID;
extern int32_t vb6_Extender_WhatsThisHelpID;
extern void*   vb6_Extender_Container;   // Set 对象 (Form/UserControl)
extern void*   vb6_Extender_DragIcon;    // Set 对象 (Picture)

// Fix 174: Extender / UserControl 宿主方法 (VBFlexGrid.ctl 引用:
//   UserControl.OLEDrag / Extender.Drag / Extender.SetFocus / Extender.ZOrder).
// Standard EXE 编译形态下容器不活动, 方法为最小可行实现 (见 vb6rtl_com.c).
void vb6_UserControl_OLEDrag(void);
void vb6_Extender_Drag(vb6_VARIANT* Action, int _has_Action);
void vb6_Extender_SetFocus(void);
void vb6_Extender_ZOrder(vb6_VARIANT* Position, int _has_Position);

// Fix 133u: czUI.ctl 用到 Extender 的 Visible / Height
//   (UserControl.Extender.Visible 赋值 + 判断, .Height 读取).
//   生成 C 是结构体字段访问 `vb6_UserControl_Extender.Visible`, 故提供该结构体.
//   注意: 不能命名为 extern 单独变量, 因为 cgen 会把 Extender 当作对象,
//   `.Visible` 作为字段追加. 结构体字段与生成代码天然吻合.
struct vb6_UserControl_Extender_Type {
    int16_t Visible;   // -1=可见 0=隐藏 (czUI 置位保存; 实际由宿主/property 配合)
    int32_t Height;    // 容器提供的控件外接高度 (twips/pixels 随 ScaleMode)
};
extern struct vb6_UserControl_Extender_Type vb6_UserControl_Extender;

// --- PropertyPage 设计器页 ---
extern void*   vb6_PropertyPage_hwnd;
extern void*   vb6_PropertyPage_hWnd;        // 生成代码两种拼写并存
extern int32_t vb6_PropertyPage_ScaleMode;
extern int32_t vb6_PropertyPage_ScaleHeight;
extern int16_t vb6_PropertyPage_Changed;     // PropertyPage 内建 Changed 属性
                                               // (生成 C 直接引用裸名 Changed)

// Fix 153: PropertyPage.SelectedControls(index) — 设计期属性页取"当前被编辑控件"
// 集合中的第 index 个 Control. 在 Standard EXE 运行期属性页永不加载, 该集合恒空,
// 故返回 NULL (void*) 即可. cgen 会把 `SelectedControls(0).<Prop>` 的成员访问按
// 后期绑定 COM 生成 (vb6_ComGetProp/SetProp), 而 RTL 对 NULL IDispatch* 是安全空
// 操作 (取回空 VARIANT / 写为 no-op), 与本文件其余 PropertyPage 内建同源.
// 定义为 static inline: 保证任意引用它的生成 TU 都拿到实体, 无需登记链接单元.
static inline void* vb6_PropertyPage_SelectedControls(int32_t index) {
    (void)index;
    return (void*)0;
}

// --- HitTest 常量 (VB6 HitResult) ---
#define vbHitResultOutside     0
#define vbHitResultTransparent 1
#define vbHitResultHit         2

// PropertyPage 的 Changed 只有 vb6_PropertyPage_Changed 这一个名字 (账 #219)。
// 这里曾 extern 过一枚裸名 `int16_t Changed` 给"源码里裸写 Changed"落脚 —— 注释当时说
// "项目自定义的 Changed 只会以类字段或模块限定名出现, 不会占用裸键"，**这句是错的**:
// 标准模块里 `Public Changed As Long` 在生成的模块 C 里就是裸名 (探针实测 C2371 / no exe)。

// --- AsyncProperty / Picture 类型常量 (VB6 内建, 此前缺失) ---
#define vbAsyncTypePicture     0
#define vbAsyncTypeFile        1
#define vbAsyncTypeByteArray   2
#define vbAsyncReadSynchronous 0
#define vbAsyncReadAsynchronous 2
#define vbAsyncReadForceUpdate 4
#define vbPicTypeNone          0
#define vbPicTypeBitmap        1
#define vbPicTypeMetafile      2
#define vbPicTypeIcon          3
#define vbPicTypeEMetafile     4

// Picture.Line 的 B / BF 由 parser 在 style 位置折成字面量 1/2 (账 #220) —— 这里曾
// extern 过两枚裸名 C 全局，与 VB 工程里叫 B 的模块级变量直接撞车，故删除。

#ifdef __cplusplus
}
#endif
