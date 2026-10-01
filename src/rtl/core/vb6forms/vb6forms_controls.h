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
void vb6_UC_RefreshCurrent(void);

// ---- host object model (form/control HWND, Controls collection, Font object) ----
// vb6_com_* / vb6_CallByName / vb6_TypeName consult these before touching lpVtbl.
void vb6_HostObj_Register(void* hwnd, const char* name, const char* vbTypeName,
                          int32_t isForm, int32_t index);
int32_t vb6_Host_IsHostObject(void* obj);
const wchar_t* vb6_Host_TypeNameOf(void* obj);
int32_t vb6_Host_GetProp(void* obj, const wchar_t* name, void* outVariant);
int32_t vb6_Host_SetProp(void* obj, const wchar_t* name, const void* inVariant);
int32_t vb6_Host_Call(void* obj, const wchar_t* name, int32_t argc,
                      void** argv, void* outVariant);
void* vb6_UC_NewFont(void);
// Fix 119: 在 vb6_UC_HostCreate 之前设定该实例的字体 (.frm BeginProperty Font 块)
void  vb6_UC_SetPendingFont(void* f);
// ---- czUI fix: 设计器子控件实例化 (.ctl 设计面上的 TextBox 等属于每个实例) ----
// 在 ucHostInit (宿主实例初始化) 期间调用, 以当前 UC 宿主窗口为父创建真实子窗口。
// left/top/width/height 单位为缇。返回子窗口句柄 (NULL = 失败)。
void* vb6_UC_CreateDesignEdit(int32_t left, int32_t top, int32_t width, int32_t height);
// 设计器 Timer: cb(ctx) 在每次 WM_TIMER 触发 (受 vb6_SetTimerEnabled/Interval 属性控制)
void* vb6_UC_CreateDesignTimer(void (*cb)(void*), void* ctx);
// ---- czUI fix: 轻量 PropertyBag (IDispatch) ----
// 供生成的 UserControl_ReadProperties 在运行期以 VB6 语义读取设计期持久化属性,
// 使 .ctl 内部"读取后同步"逻辑 (如 toggle 的 m_AnimPos 与 Checked 同步) 得以执行。
void* vb6_UC_PropBagCreate(void);
void  vb6_UC_PropBagFree(void* bag);
void  vb6_UC_BagPutStr(void* bag, const wchar_t* name, const wchar_t* value);
void  vb6_UC_BagPutInt(void* bag, const wchar_t* name, int32_t value);
void  vb6_UC_BagPutDbl(void* bag, const wchar_t* name, double value);
void  vb6_UC_BagPutBool(void* bag, const wchar_t* name, int32_t value);
// Fix 125: 字体对象身份判定 + 字段定位 (供 COM 属性读写层直接操作字体字段)
int32_t vb6_UC_IsFont(const void* p);
void*   vb6_UC_FontField(void* p, const wchar_t* name, int32_t* kind);
// Fix 168: Extender 结构体同上 (With UserControl.Extender 的 Width/Height/Align)
int32_t vb6_UC_IsExtender(const void* p);
void*   vb6_UC_ExtenderField(void* p, const wchar_t* name, int32_t* kind);
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
void* vb6_UC_ControlsItem(void* coll, int32_t index);
// Fix <vbeclipse> rev14: 裸 `UserControl.Controls` 的进程级单例 (后端发射点)。
// 必须在此声明 —— 生成代码经 vb6forms.h 伞头看到它; 缺声明会走 C4013 隐式
// 声明, 返回值按 int 截断, x64 下集合指针高 32 位丢失 → 0xC0000005。
// formHwnd 留 0, 由 vb6_uc_controlsForm() 回落到"当前" UC 实例宿主窗口。
void* vb6_UC_Controls(void);
void* vb6_UC_ControlsEnumInit(void* coll);
int32_t vb6_UC_ControlsEnumNext(void* enumPtr, void* outVariant);
// Windows VARIANT <-> vb6_VARIANT conversion + cleanup (used inside vb6com_* hooks)
void vb6_Host_ToWinVariant(const void* inV, void* outV);
void vb6_Host_FromWinVariant(const void* inV, void* outV);
void vb6_Host_ClearVariant(void* v);

#ifdef __cplusplus
}
#endif
