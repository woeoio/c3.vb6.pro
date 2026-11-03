#pragma once
// vb6forms_controls.h - 动态加载控件 (Controls.Add)
// 由 vb6forms.h 伞头按固定顺序 include，不要单独使用
// 内容 = 拆分前 vb6forms.h 第 568~579 行，逐行未改

#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 动态加载控件 (Controls.Add)
// ============================================================

// msctls_status32 的窗口类。comctl32 在本机 (v5.82 / v6) 实际**不注册**这个类,
// 由 RTL 自己补一个同名的真窗口类 (SB_* 契约 + 自绘), 已注册就跳过。
void vb6_StatusBar_RegisterClass(void);

// ============================================================
// comctl32 通用控件引导 (ProgressBar/StatusBar/Toolbar/ListView/TreeView 依赖)
// 必须在任何 CreateWindowExW 通用控件类之前调用, 否则 "msctls_progress32"
// 等类未注册 → CreateWindowExW 失败但**不报错** (GetLastError 才是 1400),
// 表现为"控件凭空消失"。由生成的 WinMain 经 vb6_Init() 调用。
// 幂等, 失败静默 (comctl32 缺失的极端环境下退化为老式控件/缺失)。
// ============================================================
void vb6_ComCtl_Init(void);

// 设置Form窗口的IDispatch指针 (Form创建后调用)
// 用于 Me.Controls.Add 等 COM 属性访问
void vb6_Form_SetDispatch(void* hwnd, void* pDispatch);

// Me.Controls.Add(progId, name) → 返回新控件的IDispatch*
// 等价于 VB6: Set ctrl = Me.Controls.Add("WMPlayer.OCX", "WMP1")
void* vb6_Form_ControlsAdd(void* hwnd, const wchar_t* progId, const wchar_t* ctrlName);

// ============================================================
// Fix <vbeclipse> 2026-10-06: 运行期 Controls.Add 免注册 OCX 表 (vbp Object= 同款机制)
// ============================================================
// cgen 在入口点烘焙: 每个 vbp Object= 引用的 OCX, 其 typelib 导入后收集的全部
// coclass {ProgID, CLSID, coclass名, 相对exe路径} 进此表。vb6_Form_ControlsAdd 按
// ProgID 命中后改走 vb6_ocxCreateAny (LoadLibrary+DllGetClassObject), 完全绕开注册表 ——
// 与设计期 vb6_OcxHost_Create 用 ocxFiles_ 免注册同款机制。未命中 → 原有注册表路径.
#define VB6_MAX_OCXREFS 256
typedef struct Vb6OcxRef {
    const wchar_t* progId;       /* "NewTabCtl.NewTab" */
    const wchar_t* clsidStr;     /* "{XXXXXXXX-...}" coclass CLSID (typelib 真实类) */
    const wchar_t* coclassName;  /* "NewTab" — progid 缺失时的末段名兜底匹配 */
    const wchar_t* fileName;     /* "NewTab01.ocx" / "bin\NewTab01.ocx" — 相对 exe, 不依赖 CWD */
} Vb6OcxRef;

// 注册运行期 OCX 免注册表 (cgen 生成, 入口点调用一次). count 超过 256 截断并打 stderr 警告.
void vb6_OcxRefRegister(const Vb6OcxRef* libs, int count);

// ============================================================
// Fix 143: 第三方 OCX 控件真宿主 (设计期 Begin NewTabCtl.NewTab 等)
// ============================================================

// 设计期属性包条目 (生成代码从 .frm 的 Begin <CoClass> 块收集)
//
// Fix 149: vt == 9 (VT_DISPATCH) 时表示该属性是**设计期字体**
// (BeginProperty Font / IconFont(n) 块), 用 font* 字段描述, Read 时经
// OleCreateFontIndirect 造一个真正的 IFont 返回.
// 背景: VB6 控件的 ReadProperties 里直接解引用字体对象
// (ctlNewTab.ctl: `If mTabData(c).IconFontName <> mTabData(c).IconFont.Name`),
// 返回 E_INVALIDARG → 默认 Nothing → 空指针解引用, 异常从窗口回调逃逸 (0xC000041D).
typedef struct Vb6OcxProp {
    const wchar_t* name;    // 属性名, 带下标原样 ("TabCaption(0)")
    unsigned short  vt;     // VARTYPE: VT_I4(3) / VT_BSTR(8) / VT_BOOL(11) / VT_R4(4) / VT_DISPATCH(9)
    long            iVal;   // VT_I4 / VT_BOOL (0/-1)
    float           fVal;   // VT_R4
    const wchar_t* sVal;    // VT_BSTR
    // --- vt == 9 (设计期字体) 字段 ---
    const wchar_t* fontName;   // 字体名 "Tahoma" / "Segoe MDL2 Assets"
    float          fontSize;   // 字号 (pt)
    long           fontWeight; // 400 / 700
    long           fontCharset;
    long           fontFlags;  // bit0 Italic, bit1 Underline, bit2 Strikethrough
} Vb6OcxProp;

// 创建并原地激活一个第三方 OCX 控件, 返回其 IDispatch* (失败返回 NULL).
//   hwndForm  宿主窗体 HWND
//   clsidStr  "{XXXXXXXX-...}" 形式的 CLSID — 首选 (来自 OCX typelib 的 coclass)
//   altClsidStr 备用 CLSID (来自 vbp Object=, 可能指向旧版本) — 首参失败时尝试; 可为 NULL
//   ocxPath   OCX 文件绝对路径 (可 NULL → 仅 CoCreateInstance, 需已注册)
//   x/y/w/h   像素矩形
//   ctrlName  控件实例名 (调试/Extender.Name 用)
//   props     设计期属性表 (经 IPersistPropertyBag::Load 灌入; 可 NULL)
// 加载顺序: <exe目录>\<文件名> (便携) → ocxPath (绝对) → CoCreateInstance (已注册) → 裸文件名.
// 激活: SetClientSite → SetExtent → PropertyBag → DoVerb(INPLACEACTIVATE|SHOW) → SetObjectRects.
void* vb6_OcxHost_Create(void* hwndForm, const wchar_t* clsidStr, const wchar_t* altClsidStr,
                         const wchar_t* ocxPath, int x, int y, int w, int h,
                         const wchar_t* ctrlName, const Vb6OcxProp* props, int propCount);

// Fix 143d: 把窗体消息转发给该窗体上的无窗口 OCX 控件 (WM_PAINT/鼠标/键盘).
// 返回 1 = 有控件处理 (调用方用 *plResult 作为 WndProc 返回值).
// 生成的窗体 WndProc 默认分支应调用它, 否则 windowless 控件永不绘制.
int32_t vb6_OcxHost_ForwardMessage(void* hwndForm, unsigned int msg,
                                   uintptr_t wParam, intptr_t lParam, intptr_t* plResult);

// 让该窗体上的无窗口 OCX 控件重绘 (窗体首次呈现 / 尺寸变化后调用)
void vb6_OcxHost_InvalidateAll(void* hwndForm);

// Fix 148: VB6 容器对象模型 (Extender / Container / Controls)
//   UserControl.Extender -> 本对象的 Extender; .Container 返回容器窗体对象;
//   容器 .Controls 返回控件集合, Controls("Name(idx)", tabIdx) 定位宿主控件.
void* vb6_AxContainer_CreateExtender(void* hwndForm, const wchar_t* ctrlName);

// Fix 143d: 把窗体上所有 OCX 控件画到给定 DC — 在 WM_PAINT 的
// BeginPaint/EndPaint 之间调用 (IViewObject::Draw).
void vb6_OcxHost_PaintAll(void* hwndForm, void* hdc);

// 账 (Outlook2013): 容器 Extender 的账 —— VB6 的 `控件.Visible / .Left/.Top/.Width/
// Height / .Move` 落在控件**窗口**上, 但 C3 生成码把它们当"原生控件"发给拿 HWND 的
// RTL 出口, 实参却是控件对象 (IDispatch*) ⇒ IsWindow() 判假、静默不动作。
// 下面三位按控件对象反查站点:
//   vb6_OcxHost_WindowOf   → 控件激活后的窗口 (windowless 返回 NULL);
//   vb6_OcxHost_GetVisible → 1 = 是托管控件 (结果写 *outVisible);
//   vb6_OcxHost_SetVisible → 1 = 已按容器语义处理 (ShowWindow / 站点标记)。
//   vb6_OcxHost_SyncRects  → 把控件窗口**当前**矩形补发给 IOleInPlaceObject::
//                            SetObjectRects (OCX 依它重排客户区); 1 = 是托管控件。
//   vb6_OcxHost_GetRect / SetRect → windowless 控件没有窗口, 容器的几何账落在站点
//                            的 rcCtrl (像素, 相对 hwndForm 客户区); 1 = 命中。
void* vb6_OcxHost_WindowOf(void* ctlDisp);
// 按 (窗体, 实例名) 反查托管控件对象 (容器 Extender 的几何/Move 用); 没有返回 NULL
void* vb6_OcxHost_DispByName(void* hwndForm, const wchar_t* name);
// 反向: 窗口(含 Codejock 内部子窗口/外层容器) → 宿主托管控件 IDispatch; 不是控件的窗口返回 NULL
void* vb6_OcxHost_DispatchOf(void* hwndOrCtl);
// 把该窗体上**无窗口**的站点控件对象补进 Controls 集合 (有窗的已由窗口枚举覆盖); 返回补入个数
int32_t vb6_OcxHost_CollectWindowless(void* hwndForm, void** out, int32_t max);
// 控件类名 (站点创建时按 CLSID 的注册 ProgID 记下); 1 = 确实是本宿主托管的 OCX 控件
int32_t vb6_OcxHost_TypeNameOf(void* hwndOrCtl, wchar_t* out, int32_t cch);
int32_t vb6_OcxHost_GetVisible(void* ctlDisp, int32_t* outVisible);
int32_t vb6_OcxHost_SetVisible(void* ctlDisp, int32_t visible);
int32_t vb6_OcxHost_SyncRects(void* ctlDisp);
/* 账 #305 第三百四十八刀: MDIForm 尺寸变化后, 给该窗体上所有 **windowless OCX**
 * (CommandBars/DockingPaneManager 这类隐形控件) 补发 SetObjectRects(全客户区)
 * + 重绘 —— 它们的排布引擎按 site 矩形走, 缺了这一发停靠带停在旧宽度。 */
void vb6_OcxHost_FormResized(void* hwndForm, int w, int h);
void vb6_OcxHost_InvokeAll(void* hwndForm, const wchar_t* method);
int32_t vb6_OcxHost_GetRect(void* ctlDisp, int32_t out[4], void** outForm);
// windowless 那条路的几何出口: 站点账本(像素)按容器 ScaleMode 折成用户单位; slot = VB6_GEOM_*
int32_t vb6_OcxHost_GetUserCoord(void* ctlDisp, int32_t slot);
int32_t vb6_OcxHost_SetRect(void* ctlDisp, int32_t x, int32_t y, int32_t w, int32_t h, int32_t mask);

// Fix <vbeclipse> 2026-10-07: 窗体 AutoRedraw=True 时, 把记忆 DC (VB6_AutoRedrawDC)
// 整块 BitBlt 到屏幕 DC — VB6 语义: AutoRedraw 窗体的所有 GDI 绘制先落记忆位图,
// WM_PAINT 再把记忆位图刷上屏 (见 3DMenu: 每 10ms 把环形图标合成进 Me.hdc, 靠窗体
// WM_PAINT 镜像可见)。无 ARDC 时不动作 (非 AutoRedraw 窗体走原生直绘)。
void vb6_FormPaintBlitAutoRedraw(void* hwndForm, void* hdc);

// 账 #305 第二百八十刀: 窗体自己的 WndProc 答「这枚窗口现在用什么字体」那一问。
// 答案住在 vb6forms_ctrl.c 那位只读主人里（自注册的窗口类答不了；DefWindowProc 也不记，见探针实测）。
void* vb6_FormAnswerGetFont(void* hwnd);

// ============================================================
// Fix 112: in-project UserControl (.ctl) instance host + form/control host object model
// ============================================================

// UserControl host descriptor — generated code emits one per .ctl module and calls vb6_UC_Register
// Fix <vbeclipse> rev18: UserControl **自有 Property 的按名桥** (晚绑定访问用)。
// 为什么需要: 工程里常见的 `Dim x As Variant: Set x = Controls.Item(id)` 之后
// `x.LeftPos = v` / `x.RightPos` 是**晚绑定**, cgen 发的是
//   vb6_ComSetProp(实例, L"LeftPos", ...) / vb6_ComGetIntProp(实例, L"LeftPos")
// 而 x 是原生 `vb6_cls_X*` 结构体 (不是 IDispatch), 宿主模型只认内建成员
// (Caption/Visible/Left/Top/Width/Height/hWnd/Refresh/Move/…), 没有名字表就只能
// 返回 Empty / 丢弃写入。实证 play78 (--arch x86): ucPerspective.Refresh 里
// `l_ucFolder.Move .LeftPos, .TopPos, …` 的四个位置读回 0 → 7 个 ucFolder 宿主全停在
// 设计期尺寸 569x441 重叠在左上角 → 停靠区完全不成形。
// cgen 为每个 .ctl 发一张本表 (名字 + 生成期写死的取值/存值 thunk), 宿主模型在
// **内建成员都没命中**时查它。
// ABI 用 void* 而不是 vb6_VARIANT*: 本头只 include <stdint.h>, 不想把 vb6rtl 拖进来;
// thunk 自己把它当 vb6_VARIANT* 用。
// get 返回非 NULL 表示命中; set 返回非 0 表示命中。
typedef struct vb6_UcPropDesc {
    const wchar_t* name;                                    // 属性名 (VB6 原名)
    void* (*get)(void* inst, void* outV);                   // NULL = 只写
    int32_t (*set)(void* inst, const void* inV);            // NULL = 只读
} vb6_UcPropDesc;

// Fix <vbeclipse> rev21: UserControl **自有 Public Sub 的按名桥** (晚绑定方法调用用)。
// 为什么需要 (与上面 props 桥同族, 但缺了它整个停靠区摆不出来):
//   工程里 `Set l_ucFolder = Controls.Item(id)` 之后调 `l_ucFolder.ShowView ViewId`
//   是**晚绑定**, cgen 发的是 `vb6_ComCall(实例, L"ShowView", …)`。而宿主模型的
//   方法分派 (uc_hostmodel_call.inc) 只有一张 **Win32 内建成员**表 (Refresh/Cls/
//   Move/Show/Hide/ZOrder/…) —— `ShowView` 是 .ctl 里自己写的 Public Sub, 不在表里
//   ⇒ 落进末尾"未知方法一律空实现"被**静默丢弃**。
//   实证 play78 (--arch x86): 5 处 `vb6_ComCall(ucFolder实例, L"ShowView", …)` 全落空
//   ⇒ ucFolder.ShowView 里那句 `l_View.View.Visible = True` 从未执行 ⇒ 所有视图窗体
//   永远 Visible=False ⇒ ucFolder.ViewArea_Resize 的 `If .Visible Then .Move …`
//   闸门永闭 ⇒ 视图窗体停在 0x0 / 设计期尺寸 ⇒ 子面板空白。
//   注意这一条与 ScaleWidth 隔离 (rev20) 无关: 那是"读到别的实例", 这是"根本没调"。
//
// 判据 (cgen 侧, 见 cgen_form_uc_methods.inc): 只发**形参全是 ByVal 标量**
// (String/Long/Integer/Boolean/Double/Single/Byte/LongPtr) 的 Public Sub;
// 带 ByRef 形参 / 对象 / Picture / Variant 的不发 —— 那些在 VB6 侧有落临时与
// VariantClear 语义, 桥里没法安全还原 (同 rev18 props 桥的取舍)。缺席 = 保持
// 改动前的"静默丢弃", 不会更糟。
//
// ABI: 形参一律 `const void* argv[]` (元素是 vb6_VARIANT*), 由 cgen 发的 thunk
// 负责按 VB6 声明类型拆箱; 返回值写 `void* outRet` (可以是 NULL —— 多数 Sub 无返回)。
// thunk 返回非 0 = 命中并已执行。
typedef int32_t (*vb6_UcMethodFn)(void* inst, int32_t argc, const void* argv[], void* outRet);

// 生成代码 (cgen_form_uc_methods.inc 发的 thunk) 拆箱 String 实参用: 从
// vb6_VARIANT 借出 BSTR, **不 VariantClear / 不 SysFreeString** —— 生命周期归
// 调用方那个实参 Variant, 与上面 props 桥的 String setter 同取舍。
// 本头只 include <stdint.h>, 不拖 vb6rtl/windows 进来 (同 vb6_UcPropDesc 的理由);
// BSTR 一律写成 `wchar_t*` (它就是 wchar_t*), 形参声明成 `const void*`,
// 调用点自己转成 vb6_VARIANT*。
// 另有一份强类型原型 (const vb6_VARIANT*) 在 vb6forms_uc_internal.h, 供 RTL 内部
// 使用 —— C 里两者类型相同, 同一 TU 同时 include 两头时也一致 (BSTR == wchar_t*)。
wchar_t* vb6_ho_variantToBstr(const void* v);

typedef struct vb6_UcMethodDesc {
    const wchar_t* name;                                   // 方法名 (VB6 原名)
    int32_t       argc;                                    // 形参个数 (含 Optional 缺省后的实参数)
    vb6_UcMethodFn fn;
} vb6_UcMethodDesc;

typedef struct vb6_UserControlDesc {
    const char* typeName;             // VB6 control type name, e.g. "ucChartBar"
    int32_t     scaleMode;            // .ctl design-time ScaleMode (1=Twip 3=Pixel)
    void*     (*create)(void);        // vb6_cls_X_New()
    void      (*init)(void* me);      // UserControl_Initialize
    void      (*paint)(void* me);     // UserControl_Paint
    void      (*resize)(void* me);    // UserControl_Resize (NULL if absent)
    void      (*show)(void* me);      // UserControl_Show (NULL if absent)
    void      (*terminate)(void* me); // UserControl_Terminate / vb6_cls_X_Destroy
    // ---- czUI fix: 鼠标事件转发 (全部可为 NULL) ----
    // VB6 运行时会把宿主窗口收到的鼠标消息转成 UserControl_MouseDown/Up/Move;
    // X/Y 已换算为控件当前 ScaleMode 单位。button: 1=左 2=右; shift: VB6 Shift。
    void      (*mouseDown)(void* me, int32_t button, int32_t shift, float x, float y);
    void      (*mouseUp)(void* me, int32_t button, int32_t shift, float x, float y);
    void      (*mouseMove)(void* me, int32_t button, int32_t shift, float x, float y);
    void      (*dblClick)(void* me);
    // ---- Fix <vbeclipse> rev18: 自有 Property 按名桥 (可为 NULL; 旧 cgen 不发这两项,
    // C 的"初始化式少于成员数"会把它们补 0 ⇒ 行为与改动前逐字节一致) ----
    const vb6_UcPropDesc* props;
    int32_t               propCount;
    // ---- Fix <vbeclipse> rev21: 自有 Public Sub 按名桥 (同为"旧 cgen 不发 ⇒ 补 0") ----
    const vb6_UcMethodDesc* methods;
    int32_t                 methodCount;
    // ---- Fix <vbeclipse> rev22: **设计期子控件的 Resize 事件**转发 (可为 NULL) ----
    //
    // VB6 里 `Private Sub ViewArea_Resize()` 是 ViewArea 这个 PictureBox 的 Resize
    // 事件, 由运行时在**该控件尺寸变化时**自动触发。RTL 以前完全没有这条转发
    // (只有上面那个 UC 自身的 resize), 于是 .ctl 写在子控件事件里的布局代码
    // 永远不跑 —— 除非谁恰好手工调了一次, 而那一次的取值可能早于布局就绪。
    //
    // 实证 play78 (--arch x86, 探针实测): ucFolder.ctl:529 `ViewArea_Resize` 里
    // 唯一那句 `View.View.Move 0+margin, 0+margin, ViewArea.Width-margin*2, …`
    // 是把视图窗体摆进 ViewArea 的**唯一**驱动。每个 ucFolder 的 WM_SIZE 序列是
    //   设计期 (569,441) → 中间 (320,309) → 最终 (468,405)
    // 而那一次手工调用发生在 Form_Load 期, 早于任何 WM_SIZE ⇒ ViewArea.Width 还是
    // 设计期 8655 缇 ⇒ 五个 folder 的视图窗体全都只拿到 W=8505, 停在同一个 567x423。
    //
    // ⚠⚠ **必须放在结构体最后**: 这个结构体是**布局式初始化** (C 的"初始化式项数
    //   少于成员数会把余下的补 0"), 所以 **RTL 与 cgen 的字段顺序必须逐字一致**。
    //   rev22 首次实现时把本槽插在 dblClick 之后 / props 之前, 而 cgen 把它发在最后
    //   ⇒ methods/methodCount 与本槽错位一整个指针宽 ⇒ `vb6_UC_OwnPropSet` 拿到
    //   垃圾 propCount/指针 ⇒ 实测 play78 直接 av read 0xa (rc=0xC0000005),
    //   栈 CreateFolder→ComSetProp→Host_SetProp→UC_OwnPropSet。
    //   **加槽一律追加到末尾**, 别按"语义分组"插到中间。
    void      (*designResize)(void* me, const char* ctrlName);

    // ---- 账 #226: UC 自身的 Click (同样**只许追加到末尾**) ----
    // VB6 的 Click 是"按下并抬起"之后发的, 落点就是宿主窗口的 WM_LBUTTONUP,
    // 由 uc_host_window.c 在 mouseUp 转调**之后**再转这一槽 (顺序与 VB6 一致)。
    // 没有这一槽时 .ctl 里的 UserControl_Click 是死码 —— Charts 六枚 UC 的
    // `RaiseEvent Click` 全写在那里, 于是容器侧的 ucX_Click 一条都收不到。
    void      (*click)(void* me);

    // ---- BackStyle = 0 'Transparent' (只许追加到末尾) ----
    // VB6 里 `.ctl` 声明 `Windowless = -1` + `BackStyle = 0` 时, 控件没有自己的
    // 窗口, 画面直接叠在容器上 ⇒ 兄弟控件 (Charts 2020 的白色卡片) 从文字底下
    // 透出来。C3 的 UC 恒是真子窗口 (WS_CLIPSIBLINGS), 兄弟窗口不会画进本窗
    // 矩形, 所以宿主得自己在离屏 DIB 里补下层: 先铺 ambient 底色, 再按 z 序把
    // 位于本窗之下且相交的兄弟窗口用 WM_PRINTCLIENT 画进来, 最后才跑
    // UserControl_Paint。0 = Opaque (VB6 缺省: 旧 cgen 不发此槽 ⇒ 补 0 ⇒ 行为
    // 与改动前逐字节一致), 1 = Transparent。
    int32_t     transparent;
} vb6_UserControlDesc;

// .ctl module self-registration (type name case-insensitive, duplicate ignored)
void vb6_UC_Register(const vb6_UserControlDesc* desc);

// Create a host child window for a UserControl instance on a form
// left/top/width/height in twips; ctrlName is the control instance name (Controls / Name)
// Returns host child HWND (NULL = type not registered)
void* vb6_UC_HostCreate(const char* typeName, int32_t left, int32_t top,
                        int32_t width, int32_t height, void* hParent, void* hInstance,
                        const char* ctrlName, int32_t index);

// host HWND -> control instance (vb6_cls_X*); NULL if not a host window
void* vb6_UC_InstanceOf(void* hwnd);
// control instance -> host HWND; NULL if unknown
void* vb6_UC_HwndOf(void* instance);
int32_t vb6_UC_IsHostHwnd(void* hwnd);

// Switch in an instance's host state (vb6_UserControl_*) before calling its public members.
void vb6_UC_Enter(void* hwnd);
// 账 #179: 成对版 —— 发码在每个 .ctl 实例方法的体首 Push、统一出口尾 Pop, 让控件代码
// 不管被谁调都跑在自己的宿主上下文里 (ScaleMode/hWnd/hDC/Font/Enabled/Extender.* 一族)。
void vb6_UC_PushInstance(void* inst);
void vb6_UC_PopInstance(void);
void vb6_UC_RefreshCurrent(void);

// ---- host object model (form/control HWND, Controls collection, Font object) ----
// vb6_com_* / vb6_CallByName / vb6_TypeName consult these before touching lpVtbl.
void vb6_HostObj_Register(void* hwnd, const char* name, const char* vbTypeName,
                          int32_t isForm, int32_t index);
int32_t vb6_Host_IsHostObject(void* obj);
// 账 #305 第二百六十二刀 (丙): 宿主模型「不认识的成员」计数 (见 uc_hostmodel.c)。
int32_t vb6_Host_UnknownMemberCount(void);
// 账 (438 语义刀): 读侧成员不存在 → 抛 438 的次数 (与上面的「真缺口」分开计数)。
int32_t vb6_Host_Error438Count(void);
const wchar_t* vb6_Host_TypeNameOf(void* obj);
int32_t vb6_Host_GetProp(void* obj, const wchar_t* name, void* outVariant);
int32_t vb6_Host_SetProp(void* obj, const wchar_t* name, const void* inVariant);
int32_t vb6_Host_Call(void* obj, const wchar_t* name, int32_t argc,
                      void** argv, void* outVariant);
void* vb6_UC_NewFont(void);
// Fix 119: 在 vb6_UC_HostCreate 之前设定该实例的字体 (.frm BeginProperty Font 块)
void  vb6_UC_SetPendingFont(void* f);
// Fix <vbeclipse> 2026-10-12 (账 #257 第十二层): 加载路径(带 bag 的顶层 UC)不跑
// UserControl_InitProperties —— 见 uc_host.c 的说明。置位方 = 窗体创建代码;
// ucHostInit 入口消费(读后清零), 由 cgen 生成的 ucHostInit 读取。
void  vb6_UC_SetSkipInitProps(int v);
/* 账 #305 第二百五十八刀: UserControl.InvisibleAtRuntime —— 置 1 时宿主记录照建, 但一个窗口都不建 */
void  vb6_UC_SetWindowless(int v);
int   vb6_UC_SkipInitPropsConsume(void);
// Fix <vbeclipse> 第十二层: 加载路径的 Show/Resize 推迟到 ReadProperties 之后再补齐
// (VB6 顺序 Initialize → ReadProperties → Show/Resize)。窗体创建代码在 UC_ReadProps
// 之后调用它; 非加载路径的实例在 HostCreate 内部就已激活, 这里会直接返回。
void  vb6_UC_ActivateHost(void* hwnd);
// ---- czUI fix: 设计器子控件实例化 (.ctl 设计面上的 TextBox 等属于每个实例) ----
// 在 ucHostInit (宿主实例初始化) 期间调用, 以当前 UC 宿主窗口为父创建真实子窗口。
// left/top/width/height 单位为缇。返回子窗口句柄 (NULL = 失败)。
void* vb6_UC_CreateDesignEdit(int32_t left, int32_t top, int32_t width, int32_t height);
// 设计器 Timer: cb(ctx) 在每次 WM_TIMER 触发 (受 vb6_SetTimerEnabled/Interval 属性控制)
void* vb6_UC_CreateDesignTimer(void (*cb)(void*), void* ctx);
// Fix <vbeclipse>: 设计器子控件 VB.PictureBox → STATIC 子窗 (视图窗体 SetParent 的容器)
void* vb6_UC_CreateDesignPicture(int32_t left, int32_t top, int32_t width, int32_t height);

// Fix <vbeclipse> rev32: 设计器 Label / Image 子控件 (ucTab 的 lblCaption / imgIcon)。
// 此前 ucHostInit 的发射循环不覆盖这两个种类 ⇒ ucTab 的两个子控件恒 NULL ⇒
// `ToolTip` setter 对 NULL 写属性直接 0xC0000005。详见 uc_host.c 里同号注释。
void* vb6_UC_CreateDesignLabel(int32_t left, int32_t top, int32_t width, int32_t height);
void* vb6_UC_CreateDesignImage(int32_t left, int32_t top, int32_t width, int32_t height);
// Fix <vbeclipse>: 设计器子控件里"工程内 UserControl" → 宿主子窗 (复用 HostCreate, 可递归);
// typeName 非已登记 UC 时返回 NULL (第三方 OCX 子控件走这条, 与不建等价)。
void* vb6_UC_CreateDesignUserControl(const char* typeName, const char* ctrlName,
                                     int32_t left, int32_t top, int32_t width, int32_t height);
// ---- czUI fix: 轻量 PropertyBag (IDispatch) ----
// 供生成的 UserControl_ReadProperties 在运行期以 VB6 语义读取设计期持久化属性,
// 使 .ctl 内部"读取后同步"逻辑 (如 toggle 的 m_AnimPos 与 Checked 同步) 得以执行。
void* vb6_UC_PropBagCreate(void);
void  vb6_UC_PropBagFree(void* bag);
void  vb6_UC_BagPutStr(void* bag, const wchar_t* name, const wchar_t* value);
void  vb6_UC_BagPutInt(void* bag, const wchar_t* name, int32_t value);
void  vb6_UC_BagPutDbl(void* bag, const wchar_t* name, double value);
void  vb6_UC_BagPutBool(void* bag, const wchar_t* name, int32_t value);
// Fix: 二进制 blob (字节数组) 槽 —— .frx 里 UserControl 的持久化 PropertyBag 块
// (InitPanels/InitBands/InitButtons/InitListImages/InitTabs...) 以原始字节喂给外层包,
// 运行期 `PropBag.Contents = <该字节数组>` 时由 uc_propbag.c 解析成扁平的
// InitXCount / InitXText<i> ... 项 (VB6 PropertyBag 序列化格式)。
void  vb6_UC_BagPutBin(void* bag, const wchar_t* name, const char* bytes, int32_t len);
// Fix 125: 字体对象身份判定 + 字段定位 (供 COM 属性读写层直接操作字体字段)
int32_t vb6_UC_IsFont(const void* p);
void*   vb6_UC_FontField(void* p, const wchar_t* name, int32_t* kind);
/* 账 #305 第二百七十六刀: 「八档 → LOGFONT」唯一主人 (fontFields = 那八档结构体本身,
   不是代理; dpi/quality 由调用方递) + `Font.hFont` 的那枚缓存句柄 (代理或八档皆可问)。
   vb6rtl_com.c 那位调用方不 include 本头, 就地 extern (与 vb6_ScalePxToUser 同款)。 */
void    vb6_UC_FontLogFont(void* fontFields, int32_t dpiY, int32_t quality, void* lfOut);
/* 账 #305 第二百七十七刀: 「磅 → 像素」(lfHeight) 那一问的唯一主人, 三处调用方都问它。 */
int32_t vb6_UC_FontHeightPx(double pt, int32_t dpiY);
/* 账 #305 第二百七十八刀: 把一枚字体的答案发到窗口上 (UC 宿主窗口 / 窗体顶层创建期那两站)。
   fontFields = NULL 就是 VB6 的默认档; 真正的发 / 存 / 删在 vb6forms_ctrl.c 那一位主人里。 */
void    vb6_UC_ApplyFontToWindow(void* hwnd, const void* fontFields);
void*   vb6_UC_FontHFont(void* p);
// Fix 168: Extender 结构体同上 (With UserControl.Extender 的 Width/Height/Align)
int32_t vb6_UC_IsExtender(const void* p);
void*   vb6_UC_ExtenderField(void* p, const wchar_t* name, int32_t* kind);
/* 账 #305 第二百八十七刀: 存完这一格之后, 「要落到宿主窗口上」由这一位主人答 (两处写手共用一条)。
   它不出现在生成代码里, 所以 §B202 那张按名点名的名单不涉及它。 */
void    vb6_UC_ExtenderVisibleApplied(void);
/* 账 #305 第二百九十二刀: Extender 几何的写手 (属性形; 方法形 `.Extender.Move` 仍刻意未接)。 */
void    vb6_UC_ExtenderGeomApplied(const wchar_t* name);

/* 账 #305 第二百八十九刀: vbContainerPosition(9)/vbContainerSize(10) 那两档要问容器的
   ScaleMode —— 交回这一位 (uc_controls.c), 问不到答 0 由调用方保留原档。
   这个名字不发进用户 TU ⇒ §B202 那张按名点名的名单不涉及它。 */
int32_t vb6_UC_ContainerScaleModeCurrent(void);
// Fix 128: 原地把 src 字体的字段拷进 dst 字体 (用于 Property Set 型字体属性,
// 避免调用其 Set 实现体里的 Refresh 破坏图表状态)
void    vb6_UC_FontAssign(void* dst, void* src);
int32_t vb6_UC_ControlsIsCollection(void* p);
// Fix 112c: built-in Collection (New Collection does not go through COM)
void* vb6_Collection_New(void);
int32_t vb6_Collection_IsCollection(void* p);
// 必须在此声明: 调用方是另一个编译单元 vb6com_foreach.c, 缺声明会被隐式当作返回
// int, 返回的枚举句柄指针在 x64 下被截断 → For Each 解引用即 0xC0000005.
void* vb6_Collection_EnumInit(void* coll);
int32_t vb6_Collection_EnumNext(void* enumPtr, void* outVariant);
int32_t vb6_UC_ControlsCount(void* coll);
// 账 #254: 接收者是窗体/UC 自身时 (VB6 的 `Form.Count`) 的同一个数 —— 两条读法共用
// 一处收集，见 uc_controls.c。
int32_t vb6_UC_ControlsCountOf(void* formHwnd);
void* vb6_UC_ControlsItem(void* coll, int32_t index);
// Fix <vbeclipse> rev14: 裸 `UserControl.Controls` 的进程级单例 (后端发射点)。
// 必须在此声明 —— 生成代码经 vb6forms.h 伞头看到它; 缺声明会走 C4013 隐式
// 声明, 返回值按 int 截断, x64 下集合指针高 32 位丢失 → 0xC0000005。
// formHwnd 留 0, 由 vb6_uc_controlsForm() 回落到"当前" UC 实例宿主窗口。
void* vb6_UC_Controls(void);
void* vb6_UC_ControlsEnumInit(void* coll);
int32_t vb6_UC_ControlsEnumNext(void* enumPtr, void* outVariant);
// 账 #305 第二百六十四刀: 同上一条的病, 在 ParentControls 这一枚上复发过一次 —— 它的声明
// 只进了内部头 vb6forms_uc_internal.h, 而生成代码经 vb6forms.h 伞头看不见那一份。
// 实测 = ComCtlsDemo 6 份产物 TU 各报一声 C4013, 交回的单例地址被按 int 折成 32 位再符号
// 扩展 (崩溃日志 av read target=0xfffffffff8cc8958, 低 32 位落在 .data) ⇒ 野接收者进晚绑定出口。
void* vb6_UC_ParentControls(void);
// Windows VARIANT <-> vb6_VARIANT conversion + cleanup (used inside vb6com_* hooks)
void vb6_Host_ToWinVariant(const void* inV, void* outV);
void vb6_Host_FromWinVariant(const void* inV, void* outV);
void vb6_Host_ClearVariant(void* v);

#ifdef __cplusplus
}
#endif
