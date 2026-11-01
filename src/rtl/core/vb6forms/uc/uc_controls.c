// uc_controls.c - vb6forms_uc 拆分片：Controls 集合 + Font 对象
//
// 内容 = 拆分前 vb6forms_uc.c 第 821~891 / 893~962 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// Controls 集合
// ============================================================

// Fix <vbeclipse> rev14: UserControl.Controls 的**进程级单例**。
//
// 背景: 生成代码里 .ctl 内的裸 `Controls` 是宿主集合, 由后端发射为
// `vb6_UC_Controls()`。单例的 formHwnd 留 0, 由 vb6_uc_controlsForm() 在
// 取用时回落到"当前" UC 实例宿主窗口 —— 于是同一个单例在 ucPerspective /
// ucFolder / ucTabStrip 等嵌套实例里各自指向**发起调用**的那个控件,
// 而不是绑定到创建时的那一个窗体。原先后端发射的是全局
// `vb6_UserControl_Controls`(恒 NULL) → Controls.Add/Item 全落空 → NULL 解引用。
static vb6_UCControls g_uc_controlsSingleton = { VB6_UC_CONTROLS_TAG, NULL };

void* vb6_UC_Controls(void) { return (void*)&g_uc_controlsSingleton; }

// Fix 221: 生成端对 `UserControl.ContainedControls` 拼出的名字 (见 vb6rtl_userctl.h)。
// 它按值取集合 —— `Set ContainedControls = UserControl.ContainedControls` 与
// `For Each ControlEnum In UserControl.ContainedControls` 两形。VB6 里
// UC 的 ContainedControls 与 Controls 指向同一个集合 (容器塞进来的子控件),
// 所以这里**必须**静态初始化成上面那枚单例 (取它的地址是常量表达式, 可以静态初始化),
// 而不是 NULL —— 退化成 NULL 时 `For Each` 枚举到空集 (FrameW.ctl:1240 的
// "找同名子控件" 会静默失配), Set 到 vb6_ret_* 再往下用也解不出东西。
void* vb6_UserControl_ContainedControls = (void*)&g_uc_controlsSingleton;

static void* vb6_uc_controlsForm(const void* coll) {
    if (!coll) return NULL;
    void* f = ((const vb6_UCControls*)coll)->formHwnd;
    // formHwnd==0 = 单例 (UserControl.Controls): 回落到当前 UC 实例宿主窗口。
    // vb6_ucCurrent 由 vb6_uc_push 维护 (最近进入的实例), 在 .ctl 方法执行期间
    // 恒指向发起调用的那个控件实例, 与 VB6 "UserControl.Controls" 语义一致。
    if (!f && vb6_ucCurrent) f = (void*)vb6_ucCurrent->hwnd;
    return f;
}

// 取本集合的归属宿主窗口; 单例回落再兜底到进程级 vb6_UserControl_hWnd。
// 供 Add/Remove 等需要"父窗口"的操作使用 (必须拿到非 NULL 才能建子窗口)。
void* vb6_UC_ControlsOwnHwnd(void* coll) {
    void* f = vb6_uc_controlsForm(coll);
    if (!f) f = vb6_UserControl_hWnd;
    return f;
}

// 账 #257: UserControl.ParentControls 的**进程级单例** —— 与 vb6_UC_Controls() 同族但
// 指向**父容器**, 不能共用同一枚单例(会互相踩 formHwnd)。
//
// VB6 语义: UC 上写 `UserControl.ParentControls` 拿的是**所在容器**(窗体/Frame/PictureBox)
// 的控件集合 —— ComCtlsDemo 里全是这个用法:
//   StatusBar.ctl:481   If UserControl.ParentControls.Count = 0 Then StatusBarAlignable = ...
//   CheckBoxW.ctl:1267  For Each ControlEnum In UserControl.ParentControls
// 所以归属窗口 = GetParent(当前 UC 宿主窗口); 没有父(顶层/未装填)时留 NULL ⇒ 集合为空,
// 与 `Container`/`Parent` 那一档"顶层窗体无容器"同口径(uc_hostmodel_getprop.inc)。
//
// 以前这一路发的是恒 NULL 的全局 vb6_UserControl_ParentControls ⇒
// `UserControl.ParentControls.Count` 变成 vb6_ComGetIntProp((void*)NULL, L"Count"),
// 实测崩在 vb6_ComUnpackInt (StatusBar 创建期 100%, MainForm.Show 因此静默失败)。
static vb6_UCControls g_uc_parentControlsSingleton = { VB6_UC_CONTROLS_TAG, NULL };

void* vb6_UC_ParentControls(void) {
    void* ucHwnd = NULL;
    if (vb6_ucCurrent) ucHwnd = (void*)vb6_ucCurrent->hwnd;
    if (!ucHwnd) ucHwnd = vb6_UserControl_hWnd;
    /* 账 #305 第293刀: 口径 = **窗体这一层**的控件，不是「直接父窗口的子控件」。
       实测 (ComCtlsDemo): ImageCombo1 装在 CoolBar1 里, 它要的 ImageList1/2/3 摆在
       MainForm 上; 参考版靠 `For Each c In UserControl.ParentControls` +
       `TypeName(c)="ImageList"` + `c.hImageList` 挂上图标 —— 若只交 CoolBar 的子控件,
       三枚 ImageList 一枚都进不来, 循环空转, 下拉列表就没有图标 (且全程不报错)。
       VB6 这一族的容器概念是窗体 (Extender.Container), 嵌套 UC 也一样。 */
    void* parentHwnd = ucHwnd ? (void*)GetAncestor((HWND)ucHwnd, GA_ROOT) : NULL;
    if (!parentHwnd && ucHwnd) parentHwnd = (void*)GetParent((HWND)ucHwnd);
    g_uc_parentControlsSingleton.formHwnd = parentHwnd;
    return (void*)&g_uc_parentControlsSingleton;
}

// 收集某窗体上的全部子控件 HWND (含 UserControl 宿主窗口)
static int32_t vb6_uc_collectChildren(void* formHwnd, void** out, int32_t max) {
    // 宿主窗口未知时直接返回 0 —— 不能让 GetWindow((HWND)0, GW_CHILD) 去枚举桌面
    // 顶层窗口 (那会返回与本集合无关的窗口)。
    if (!formHwnd) return 0;
    int32_t n = 0;
    for (int32_t i = 0; i < vb6_ucHoCount && n < max; i++) {
        vb6_HostObjRec* rec = &vb6_ucHo[i];
        if (rec->isForm) continue;
        if (rec->hwnd) {
            if (GetParent((HWND)rec->hwnd) == (HWND)formHwnd) out[n++] = rec->hwnd;
        } else {
            /* 账 #305 第二百八十二刀: 无窗宿主 (第二百五十八刀那户 InvisibleAtRuntime) ——
               容器关系问创建时就填好的 `UCRec->parent`(有窗那一支走的是同一条关系:
               GetParent(自己的窗口))，交回**实例指针**。VB6 本人把这三枚 ImageList 一样摆在
               容器的 Controls 里（靶工程 ToolBar/ListView/TreeView 全靠
               `For Each ControlEnum In UserControl.ParentControls` + `TypeName(...) = "ImageList"`
               找到它们再问 `.hImageList`），少交一枚 ⇒ 图标一个都不挂。 */
            vb6_UCRec* ur = vb6_uc_findByInstance(rec->instance);
            if (ur && (void*)ur->parent == formHwnd) out[n++] = rec->instance;
        }
    }
    // 兜底: 枚举真实子窗口 (vb6_CreateControl 创建的标准控件未必逐个登记)
    HWND child = GetWindow((HWND)formHwnd, GW_CHILD);
    while (child && n < max) {
        int32_t seen = 0;
        for (int32_t i = 0; i < n; i++) if (out[i] == (void*)child) { seen = 1; break; }
        if (!seen) out[n++] = (void*)child;
        child = GetWindow(child, GW_HWNDNEXT);
    }
    // 账 (Outlook2013 左面板 VisualTheme): VB6 的 Controls 含**无窗口**的 OCX 控件,
    // 窗口枚举看不见它们 —— 不补进来, `For Each Ctrl In Form.Controls` 就永远
    // 遍历不到 (CtrlCaption.VisualTheme 一类设置全部落空)。
    n += vb6_OcxHost_CollectWindowless(formHwnd, out + n, max - n);
    return n;
}

vb6_UCControls* vb6_uc_newControls(void* formHwnd) {
    vb6_UCControls* c = (vb6_UCControls*)malloc(sizeof(vb6_UCControls));
    if (c) { c->tag = VB6_UC_CONTROLS_TAG; c->formHwnd = formHwnd; }
    return c;
}

int32_t vb6_UC_ControlsIsCollection(void* p) { return vb6_uc_isControls(p); }

int32_t vb6_UC_ControlsCount(void* coll) {
    if (!vb6_uc_isControls(coll)) return 0;
    void* kids[VB6_UC_MAX_OBJ];
    return vb6_uc_collectChildren(vb6_uc_controlsForm(coll), kids, VB6_UC_MAX_OBJ);
}

/* 账 #254: 「这枚容器上有多少枚子控件」的第二个问法 —— 接收者是**窗体/UC 自己**
   (VB6 的 `Form.Count`)，手上没有集合对象。数法转调上面同一条收集，不开第二份：
   Charts 的 ClsResizer 用 `ReDim Rects(oForm.Count - 1)` 给 `For Each oCtrl In
   oForm.Controls` 铺格子，这两个数一旦不等，循环第一格就写下标越界。 */
int32_t vb6_UC_ControlsCountOf(void* formHwnd) {
    void* kids[VB6_UC_MAX_OBJ];
    return vb6_uc_collectChildren(formHwnd, kids, VB6_UC_MAX_OBJ);
}

void* vb6_UC_ControlsItem(void* coll, int32_t index) {
    if (!vb6_uc_isControls(coll)) return NULL;
    void* kids[VB6_UC_MAX_OBJ];
    int32_t n = vb6_uc_collectChildren(vb6_uc_controlsForm(coll), kids, VB6_UC_MAX_OBJ);
    if (index < 1 || index > n) return NULL;
    return kids[index - 1];
}

/* Fix 148: 按 VB6 名查找控件 — 容器 Controls("txtDoc(0)") 语义.
 * VB6 控件 (如 NewTab 的 TDIMode) 通过 UserControl.Parent.Controls(name, tabIdx)
 * 拿到宿主窗体上的控件 (用于 SetParent 重新父化到 Tab 页).
 * name 支持 "Name" 与 "Name(idx)" 两种写法; tabIdx<0 表示不限.
 * 返回控件 hwnd (未找到 NULL). */
void* vb6_UC_ControlsItemByName(void* coll, const wchar_t* name, int32_t tabIdx) {
    (void)tabIdx;
    if (!name || !*name) return NULL;
    wchar_t base[VB6_UC_NAME_LEN];
    int wantIdx = -1;
    /* 拆 "Name(idx)" */
    const wchar_t* lp = wcschr(name, L'(');
    if (lp) {
        size_t n = (size_t)(lp - name);
        if (n >= VB6_UC_NAME_LEN) n = VB6_UC_NAME_LEN - 1;
        wcsncpy(base, name, n); base[n] = 0;
        wantIdx = _wtoi(lp + 1);
    } else {
        wcsncpy(base, name, VB6_UC_NAME_LEN - 1); base[VB6_UC_NAME_LEN - 1] = 0;
    }
    HWND formHwnd = (coll && vb6_uc_isControls(coll)) ? (HWND)vb6_uc_controlsForm(coll) : NULL;
    /* 优先在已注册宿主对象表里找 (名字精确, 不受父窗口层级影响) */
    for (int32_t i = 0; i < vb6_ucHoCount; i++) {
        vb6_HostObjRec* r = &vb6_ucHo[i];
        if (r->isForm) continue;
        void* item = r->hwnd;
        /* 账 #305 第二百八十二刀: 按名查这一路以前把无窗宿主整个跳过 (`!r->hwnd) continue`)，
           与上面 collectChildren 那一档同一条病 —— 容器关系改问记录里的 parent，交实例指针。 */
        if (!item) {
            vb6_UCRec* ur = vb6_uc_findByInstance(r->instance);
            if (!ur) continue;
            if (formHwnd && (void*)ur->parent != formHwnd) continue;
            item = r->instance;
        } else if (formHwnd && GetAncestor((HWND)item, GA_ROOT) != formHwnd
                   && !IsChild(formHwnd, (HWND)item)) {
            /* 不在该窗体下 (可能是别的窗体的同名控件) — 跳过 */
            if (!IsWindow((HWND)item)) continue;
        }
        if (_wcsicmp(r->name, base) != 0) continue;
        if (wantIdx >= 0 && r->index != wantIdx) continue;
        return item;
    }
    return NULL;
}

// ---- Fix <vbeclipse> rev14: Controls 的"对象"形态 (工程内 UserControl) ----
//
// 生成代码对 `Set x = Controls.Item(...)` / `Controls.Add(...)` 有两种消费方式:
//   1) 早绑定直调 (x 声明为 ucFolder 等项目类): `vb6_ucFolder_AddView(x, ...)`
//      —— 需要的是**实例指针** vb6_cls_ucFolder*;
//   2) 晚绑定宿主分派 (x 声明为 Object/Variant)
//      —— 需要的是**宿主 HWND** (vb6_Host_Call/GetProp 才能应答)。
// 工程里同一变量两处混用 (CreateFolder 的 l_ucFolder 既直调 AddView 又走
// vb6_ComSetProp), 而 AddView 一类直调是**写死形参类型**的硬路径, 一旦收到
// HWND 即按结构体解引用 → AV。故 Add/Item 统一返回**实例指针** (找不到实例的
// 标准控件才回落 HWND); 晚绑定路径对实例是安全的 (vb6_ComIsDispatchable 判否
// → "属性/方法未找到" 静默返回, 不崩)。
void* vb6_UC_ControlsItemObj(void* coll, int32_t index) {
    void* hw = vb6_UC_ControlsItem(coll, index);
    void* inst = hw ? vb6_UC_InstanceOf(hw) : NULL;
    return inst ? inst : hw;
}

void* vb6_UC_ControlsItemObjByName(void* coll, const wchar_t* name) {
    void* hw = vb6_UC_ControlsItemByName(coll, name, -1);
    void* inst = hw ? vb6_UC_InstanceOf(hw) : NULL;
    return inst ? inst : hw;
}

// 把宿主窗口从 vb6_ucHo / vb6_ucRecs 摘除 (Remove 时调用)。
// 采用"末位填补"压缩数组; vb6_ucCurrent 若正好是被摘除项, 置 NULL (调用方
// Remove 的必然是别的控件, 不会摘到自身, 故正常情况下不发生)。
static void vb6_uc_releaseHost(void* hwnd) {
    if (!hwnd) return;
    for (int32_t i = 0; i < vb6_ucHoCount; i++) {
        if (vb6_ucHo[i].hwnd == hwnd) { vb6_ucHo[i] = vb6_ucHo[--vb6_ucHoCount]; break; }
    }
    for (int32_t i = 0; i < vb6_ucRecCount; i++) {
        if ((void*)vb6_ucRecs[i].hwnd == hwnd) {
            if (vb6_ucCurrent == &vb6_ucRecs[i]) vb6_ucCurrent = NULL;
            vb6_ucRecs[i] = vb6_ucRecs[--vb6_ucRecCount];
            break;
        }
    }
}

// Controls.Add(progId, name): 在当前 UC 宿主窗口下创建一个该类型的 UserControl
// 宿主 (VB6 `Controls.Add "VbEclipse.ucFolder", FolderId` 语义)。返回**实例指针**。
// 几何用父窗口客户区尺寸兜底 (工程随后用 Move/LeftPos 等定位), 避免 0×0 退化窗口。
void* vb6_UC_ControlsAdd(void* coll, const wchar_t* progId, const wchar_t* name) {
    if (!progId || !*progId) return NULL;
    HWND parent = (HWND)vb6_UC_ControlsOwnHwnd(coll);
    if (!parent) return NULL;
    char progU8[256];
    char nameU8[VB6_UC_NAME_LEN * 4];
    vb6_wideToU8Buf(progId, progU8, (int)sizeof(progU8));
    vb6_wideToU8Buf(name ? name : L"", nameU8, (int)sizeof(nameU8));
    HINSTANCE hInst = (HINSTANCE)GetWindowLongPtrW(parent, GWLP_HINSTANCE);
    RECT rc; GetClientRect(parent, &rc);
    int32_t w = vb6_XToTwipX(rc.right > 0 ? rc.right : 120);
    int32_t h = vb6_YToTwipY(rc.bottom > 0 ? rc.bottom : 120);
    void* hw = vb6_UC_HostCreate(progU8, 0, 0, w, h, parent, hInst, nameU8, -1);
    void* inst = hw ? vb6_UC_InstanceOf(hw) : NULL;
    return inst ? inst : hw;
}

// Controls.Remove(obj) / Controls.Remove(name): 销毁对应子控件。
// obj 可以是实例指针 (晚绑定 Set 对象) 或 HWND; name 为 VB6 控件名。
int32_t vb6_UC_ControlsRemove(void* coll, void* obj, const wchar_t* name) {
    void* hw = NULL;
    if (obj) {
        hw = vb6_UC_HwndOf(obj);                    // 工程类实例 → 宿主窗口
        if (!hw && IsWindow((HWND)obj)) hw = obj;   // 直接给的 HWND
        if (!hw) {
            /* 可能是本语境的宿主实例: 用 vb6_uc_findByInstance 的逆 →
               实例若即 vb6_ucRecs[i].me, 取其 hwnd */
            vb6_UCRec* r = vb6_uc_findByInstance(obj);
            if (r) hw = (void*)r->hwnd;
        }
    } else if (name && *name) {
        hw = vb6_UC_ControlsItemByName(coll, name, -1);
    }
    if (!hw) return 0;
    vb6_uc_releaseHost(hw);
    DestroyWindow((HWND)hw);
    return 1;
}

// 句柄首元素存指针, 必须用 intptr_t —— x64 下按 int32_t 存会截断高 32 位 (同 Collection 枚举器)
void* vb6_UC_ControlsEnumInit(void* coll) {
    intptr_t* e = (intptr_t*)malloc(sizeof(intptr_t) * 2);
    if (e) { e[0] = (intptr_t)coll; e[1] = 0; }
    return e;
}

int32_t vb6_UC_ControlsEnumNext(void* enumPtr, void* outV) {
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    intptr_t* e = (intptr_t*)enumPtr;
    if (!e) return 0;
    void* coll = (void*)e[0];
    void* kids[VB6_UC_MAX_OBJ];
    int32_t n = vb6_uc_collectChildren(vb6_uc_controlsForm(coll), kids, VB6_UC_MAX_OBJ);
    if (e[1] >= n) return 0;
    void* hwnd = kids[e[1]++];
    memset(out, 0, sizeof(*out));
    out->vt = vb6_vtDispatch;
    // Fix <vbeclipse>: 交**真 IDispatch 包装器**而非裸 HWND。`For Each ctrl In Controls`
    // 的栈 VARIANT 收尾会调 oleaut32 `VariantClear` → 对 VT_DISPATCH 无条件 Release。裸
    // HWND 被当对象指针 Release → 解引用伪 vtable → 野调用崩 (ucTabStrip.UserControl_Resize
    // 实测, PC=0xcccccc.. 未初始化)。Vb6HostWrap 有真 vtable, AddRef/Release 走引用计数,
    // VariantClear 安全; 宿主属性/方法读侧 (ComGetProp/Host_GetProp) 顶部 UnwrapHost 还原
    // 成 HWND, 语义不变。与 vb6_ComPackObject 对宿主对象的包装同口径。
    out->pdispVal = vb6_UC_WrapHostObject(hwnd);
    return 1;
}

// ============================================================
// Font 对象 (stdole.StdFont 最小实现)
// ============================================================

void* vb6_UC_NewFont(void) {
    vb6_UCFontRec* fr = (vb6_UCFontRec*)malloc(sizeof(vb6_UCFontRec));
    if (!fr) return NULL;
    vb6_ComIface_Font* f = (vb6_ComIface_Font*)calloc(1, sizeof(vb6_ComIface_Font));
    if (!f) { free(fr); return NULL; }
    f->Name = vb6_BSTR_FromStr(L"MS Sans Serif");
    f->Size = 8.25f;
    f->Charset = 0;
    f->Weight = 400;
    fr->tag = VB6_UC_FONT_TAG;
    fr->font = f;
    fr->next = vb6_ucFonts;
    vb6_ucFonts = fr;
    return f; // 与 cgen 的 vb6_ComIface_Font* 直接字段访问保持一致
}

/* 账 #305 第二百七十七刀: 「磅 → 像素」这一问的唯一主人 (lfHeight 那一格)。
 * 之前它住三处、两套口径: vb6forms.c 的 vb6_Vb6DefaultGuiFont 用 -MulDiv(825,dpiY,7200)
 * (MulDiv **四舍五入**), vb6forms_ctrl.c 的 vb6_SetControlFontSize 用 -(int)(pt*dpi/72f+0.5f)
 * (也是四舍五入), 而第二百七十六刀那位八档主人里再抄了第三份 —— 三处同公式, 于是全都比
 * VB6 大一档。证据 = 第二百七十四刀立的那把时间无关的尺 (逐枚窗口问它自己的字体再抄回 LOGFONTW 读数):
 *   参考 ComCtlsDemo.exe  69 枚窗口 (-13, 400, "MS Sans Serif")  ← 8.25pt @ 120dpi = 13.75 截断
 *   C3 编出来的同一工程  38 枚窗口 (-14, 400, ...)              ← 同一发四舍五入 = 14
 * 两台都是 system-DPI-aware (探针 .local/k277/aware.py: awareness=1、GetDpiForWindow=120、
 * 顶层客户区 844x780 ↔ 844x781 逐家相同) ⇒ 这一档差**不是** DPI 虚拟化造的, 是换算本身的口径。
 * ⚠ 只有一条读数支撑这个口径 (语料里 8.25pt 是绝对主力): 差只在 frac(pt*dpi/72) >= 0.5 时现形,
 *   96dpi 下 8.25pt = 11.0 整 ⇒ 本机 CI (96dpi) 上这刀**零可见效果**, 判据只能靠那把尺。
 * 返回**负的** lfHeight (字符高度, GDI 的 VB6 那一档), 0pt 交 0 (SP3 那枚行为钉要吃这个)。
 * dpi <= 0 一律按 96 兜底 (调用方拿不到 DC 时的旧口径, 一字未动)。 */
int32_t vb6_UC_FontHeightPx(double pt, int32_t dpiY) {
    int dpi = (dpiY > 0) ? dpiY : 96;
    return -(int)(pt * (double)dpi / 72.0);   /* 截断, 不四舍五入 —— 照 VB6 的读数 */
}

vb6_ComIface_Font* vb6_uc_fontOf(void* p) {
    if (p == &vb6_UserControl_FontObj) return &vb6_UserControl_FontObj;
    for (vb6_UCFontRec* f = vb6_ucFonts; f; f = f->next)
        if (p == f || p == f->font) return (vb6_ComIface_Font*)f->font;
    return NULL;
}

// Fix 125: 字体对象不是 COM 对象, 但生成代码会把 `With <font>: .Name = x` 编译成
// vb6_ComSetProp(字体指针, L"Name", ...) —— 对普通结构体做 IDispatch::Invoke 会走
// 垃圾 vtable 直接崩 (LabelPlus 的 `Property Set Font` 就是这么崩的), 于是
// ucProgressCircular 的 Caption1_Font/Caption2_Font (14.25/8.25) 一直无法应用。
// 这里向 COM 层暴露"身份判定 + 字段定位", 由 vb6com_invoke.c 直接读写字段。
int32_t vb6_UC_IsFont(const void* p) {
    return vb6_uc_fontOf((void*)p) != NULL;
}

/* 账 #305 第二百七十六刀: 「八档 → LOGFONT」这一问的**唯一主人**。
 * 在此之前它只有一位抄写员: vb6rtl_com.c 的 vb6_uc_measureText 就地一份 (量字用,
 * DPI 写死 96), 而 VB 层要的 `Font.hFont` 压根没人答 —— 第三百零五刀靶工程那 34 枚
 * 全零 LOGFONT 就是这么来的 (Common/Common.bas:1012 GetObjectAPI(Font.hFont, …) 拿到
 * 的是 0, 那次 GetObject 直接失败, 局部 UDT 保持全零, 递给窗口的就是 GDI 默认字体)。
 * 本刀把 measureText 那一份撤进来当第一位主人 (它那份的算法一字未改: 磅→像素仍是
 * `-(int)(pt*dpi/72+0.5)` (第二百七十七刀把这一步交给 vb6_UC_FontHeightPx, 口径改成截断)、八档的兜底仍是 8.25 / 400 / DEFAULT_CHARSET / "MS Sans Serif",
 * 连"其余字段留零"也照抄), 从此第二格答案只在调用方 (递进来的 DPI 与质量档)。
 * ⚠ 一处**刻意**的口径修正: 字重按 VB6 的读侧派生 —— Bold 为真一律 700, 与 Weight
 *   存的是哪个数无关 (StdFont 的真实存储只有 Weight, Bold 是它的视图)。旧写法
 *   `Weight ? Weight : (Bold ? 700 : 400)` 让写进 Bold 的那个 True 永远被默认的
 *   Weight=400 挡在后面, 也就是第二百七十三刀修好的那 13 处 Bold 到了这里又掉。
 *   只在 Bold 非零时答案才不同, 而语料里 Bold 非零只有两种来源: 设计期 Weight>=700
 *   (cgen_form_create_controls.inc:109 同步设过 Bold ⇒ 那档两边都给 700, 读数不变)
 *   或用户代码写 Bold ⇒ 后者正是本刀要它生效的那一种。
 * 参数按**void*** 收 (本头不含 wingdi.h 的调用方也能递栈上的 LOGFONTW), dpi/quality
 *   由调用方决定 —— 与账 #298 §B129 那条「公式只住一枚, DPI 从外面递进来」同口径。 */
void vb6_UC_FontLogFont(void* fontFields, int32_t dpiY, int32_t quality, void* lfOut) {
    const vb6_ComIface_Font* f = (const vb6_ComIface_Font*)fontFields;
    LOGFONTW* lf = (LOGFONTW*)lfOut;
    double pt;
    int dpi = (dpiY > 0) ? dpiY : 96;
    if (!lf) return;
    memset(lf, 0, sizeof(*lf));
    pt = (f && f->Size > 0) ? (double)f->Size : 8.25;
    lf->lfHeight    = vb6_UC_FontHeightPx(pt, dpi);
    lf->lfWeight    = (f && f->Bold) ? 700 : ((f && f->Weight) ? f->Weight : 400);
    lf->lfItalic    = (BYTE)((f && f->Italic) ? 1 : 0);
    lf->lfUnderline = (BYTE)((f && f->Underline) ? 1 : 0);
    lf->lfStrikeOut = (BYTE)((f && f->Strikethrough) ? 1 : 0);
    lf->lfCharSet   = (BYTE)((f && f->Charset) ? f->Charset : DEFAULT_CHARSET);
    lf->lfQuality   = (BYTE)quality;
    /* 名字那一档把**零长串**也算"没有名字" (f->Name 非 NULL 但内容空)。
     * 依据两条: ① VB6 的 StdFont.Name 不允许是空串, 所以这八档里出现空名字只能是
     *    "没指定", 而"没指定"在 VB6 交出的就是默认字族; ② 探针 (.local/k276/fontprobe.c
     *    F 行) 实测空名字交出去, GetObject 读回来的 lfFaceName 仍是空 —— GDI 不改写它,
     *    于是下游谁再拿这枚 LOGFONT 去造字体, 拿到的是 DC 默认字体而不是 VB6 的那一枚,
     *    而且是**静默**的。这一档宁可退回 "MS Sans Serif", 也不留一条静默路。
     * ⚠ 这条与本刀靶工程那 34 枚 Segoe UI **无关** (我一开始就是这么记的, 探针把它推翻:
     *    那 34 枚的名字是真从字段里读出来的 "Segoe UI", 产地在 uc_host_window.c:353)。 */
    if (f && f->Name && f->Name[0]) {
        size_t n = wcslen(f->Name);
        if (n > 31) n = 31;
        memcpy(lf->lfFaceName, f->Name, n * sizeof(wchar_t));
    } else {
        wcscpy(lf->lfFaceName, L"MS Sans Serif");
    }
}

/* 快照与当前八档是否一致 (一致 = 缓存里那枚 HFONT 还代表这八档, 不用重造)。
 * 名字比的是内容副本, 不借 Name 那枚指针: 本 RTL 往 Name 写的时候从不 SysFreeString
 * 旧串 (见 vb6_UC_FontAssign / vb6com_invoke.c 的 kind 0 那支), 借来比是踩在泄漏上。 */
static int32_t vb6_uc_fontSigSame(const vb6_ComIface_Font* f) {
    if (f->Size != f->hFontSigSize) return 0;
    if (f->Bold != f->hFontSigBold || f->Italic != f->hFontSigItalic) return 0;
    if (f->Underline != f->hFontSigUnderline || f->Strikethrough != f->hFontSigStrike) return 0;
    if (f->Weight != f->hFontSigWeight || f->Charset != f->hFontSigCharset) return 0;
    return wcsncmp(f->Name ? f->Name : L"", f->hFontSigFace, 32) == 0 ? 1 : 0;
}

static void vb6_uc_fontSigTake(vb6_ComIface_Font* f) {
    f->hFontSigSize      = f->Size;
    f->hFontSigBold      = f->Bold;
    f->hFontSigItalic    = f->Italic;
    f->hFontSigUnderline = f->Underline;
    f->hFontSigStrike    = f->Strikethrough;
    f->hFontSigWeight    = f->Weight;
    f->hFontSigCharset   = f->Charset;
    f->hFontSigFace[0] = 0;
    if (f->Name) {
        size_t n = wcslen(f->Name);
        if (n > 31) n = 31;
        memcpy(f->hFontSigFace, f->Name, n * sizeof(wchar_t));
        f->hFontSigFace[n] = 0;
    }
}

/* 造 / 取这枚伪字体的 HFONT (账 #305 第二百七十六刀, `Font.hFont` 那一条答案)。
 * 为什么是缓存而不是每次新建: 问的人只要 LOGFONT 的内容 (靶工程拿句柄去 GetObjectAPI,
 * 再自己 CreateFontIndirect 一份), 而问的频率是每建一枚控件一发 —— 每次一枚的话
 * 就是每枚控件漏一枚 GDI 对象, 而这里连伪字体本体都还欠 §B164 那一族回收。
 * 自失效: 八档一改就作废重建 (写侧没钩子, 见 vb6rtl_userctl.h 那八格新字段的注释)。
 * 所有权照 VB6 的 IFont.hFont: 句柄属于这枚字体, **不属于问它的人** ⇒
 * ⚠ 不许把它交给 WM_SETFONT —— vb6forms_ctrl.c 的 vb6_SetControlFontFromLogFont
 *   换字体后无条件 DeleteObject(旧句柄) (vb6forms.c:147 那条注释记的就是这一格),
 *   递出去就等于把缓存换成悬空句柄。本 RTL 无任何调用点这么做。
 * 质量档 = NONANTIALIASED_QUALITY, 与 vb6_Vb6DefaultGuiFont 为窗口字体定下的那条
 * 「VB6 的 MS Sans Serif 是点阵字体, 从不抗锯齿」同口径 (量字那一位调用方维持它
 * 一直以来的 0, 不跟着动 —— Fix 113k 记过: 给 TextWidth/TextHeight 换字体度量会把
 * Charts 的图例/轴标签排布推坏)。 */
static void* vb6_uc_fontEnsureHFont(vb6_ComIface_Font* f) {
    LOGFONTW lf;
    HFONT h;
    if (!f) return NULL;
    if (f->hFontCache && vb6_uc_fontSigSame(f)) return f->hFontCache;
    if (f->hFontCache) { DeleteObject((HFONT)f->hFontCache); f->hFontCache = NULL; }
    vb6_UC_FontLogFont(f, vb6_DpiY(), NONANTIALIASED_QUALITY, &lf);
    h = CreateFontIndirectW(&lf);
    if (!h) return NULL;
    vb6_uc_fontSigTake(f);
    f->hFontCache = (void*)h;
    return (void*)h;
}

void* vb6_UC_FontHFont(void* p) {
    return vb6_uc_fontEnsureHFont(vb6_uc_fontOf(p));
}

/* 账 #305 第二百七十八刀: 把一枚字体的答案**发到窗口上**。
 * 靶 = 参考 ComCtlsDemo.exe 那 29 枚 ThunderRT6UserControl 每一枚都答字体 (27 枚 -13/MS Sans Serif +
 * 2 枚 -11/Tahoma)、ThunderRT6FormDC 也答 -13/MS Sans Serif; C3 的对应两户 (27 枚 VB6_UserControlHost
 * + 顶层 VB6_Form_*) 在同一把时间无关的尺上读回**空** —— 从没发过。发过之后它们的文字量、owner-draw
 * 与 .FontName 那一族读数才落在 VB6 那枚字体上, 不是 DC 默认字体。
 * 本站只拼答案: 八档 (fontFields = NULL 即 VB6 的默认档 8.25pt / MS Sans Serif / 400) 问第二百七十六刀
 * 那位主人, DPI 递真屏、质量档递 NONANTIALIASED (与窗口字体同口径); 发出去 / 自存 / 删旧 交给
 * vb6forms_ctrl.c 里那一位既有主人 —— 那一站自带「只删我们自己造的那张」的账, 不抄第二份。
 * ⚠ 也别顺手把第二百七十六刀那枚缓存句柄交出去: WM_SETFONT 的接手方换字体时 DeleteObject 旧句柄。 */
void vb6_UC_ApplyFontToWindow(void* hwnd, const void* fontFields) {
    LOGFONTW lf;
    if (!hwnd) return;
    vb6_UC_FontLogFont((void*)fontFields, vb6_DpiY(), NONANTIALIASED_QUALITY, &lf);
    vb6_SetControlFontFromLogFont((HWND)hwnd, &lf);
}


// 返回字段地址; *kind: 0=BSTR, 1=float, 2=int16, 3=int32, 5=句柄(指针宽度); 未命中返回 NULL
void* vb6_UC_FontField(void* p, const wchar_t* name, int32_t* kind) {
    vb6_ComIface_Font* f = vb6_uc_fontOf(p);
    if (!f || !name || !kind) return NULL;
    if (_wcsicmp(name, L"Name") == 0)          { *kind = 0; return &f->Name; }
    if (_wcsicmp(name, L"Size") == 0)          { *kind = 1; return &f->Size; }
    if (_wcsicmp(name, L"Bold") == 0)          { *kind = 2; return &f->Bold; }
    if (_wcsicmp(name, L"Italic") == 0)        { *kind = 2; return &f->Italic; }
    if (_wcsicmp(name, L"Underline") == 0)     { *kind = 2; return &f->Underline; }
    if (_wcsicmp(name, L"Strikethrough") == 0) { *kind = 2; return &f->Strikethrough; }
    if (_wcsicmp(name, L"Weight") == 0)        { *kind = 3; return &f->Weight; }
    if (_wcsicmp(name, L"Charset") == 0)       { *kind = 3; return &f->Charset; }
    /* 第二百七十六刀: hFont 不是字段, 是八档的派生量 —— 交的是**缓存那一格**的地址,
       所以必须先把这格填上再返回 (造不出来就答"认不出", 由调用方去报数)。
       名字不叫 hFont 的那半句理由在同一条注释里: 用户代码的 `Font.hFont` 只会以
       vb6_ComGetProp(..., L"hFont") 的形状到达, 不会发成结构体字段直读。 */
    if (_wcsicmp(name, L"hFont") == 0) {
        if (!vb6_uc_fontEnsureHFont(f)) return NULL;
        *kind = 5;
        return &f->hFontCache;
    }
    return NULL;
}

// Fix 168: `UserControl.Extender` 同样是普通结构体 (vb6_UserControl_Extender,
// 只有 Visible/Height 两个字段), 不是 IDispatch。而 VBFlexGrid.ctl 的
//   `With UserControl.Extender: .Align / .Width / .Height` (UserControl_Resize)
// 被编译成 vb6_ComGetProp/SetProp(&vb6_UserControl_Extender, L"...") → 对结构体
// 做 Invoke 必跳进 .data 崩 (0xC0000005, 与 Fix 125 字体、Fix 166 同族)。
// 身份判定就认这一个全局实例的地址; 成员定位返回 **vb6_Extender_\* 全局槽位**
// 的地址 —— 那才是 Fix 162 起按当前实例同步 (vb6_uc_push) 的真实存储,
// 结构体字段只是 Fix 133u 时期给直接字段访问用的兼容壳。
int32_t vb6_UC_IsExtender(const void* p) {
    return p == (const void*)&vb6_UserControl_Extender;
}

// kind 编码与 vb6_UC_FontField 一致: 0=BSTR, 1=float, 2=int16, 3=int32。
// Fix 296 / 3b 前置: 补一位 **4 = 对象 (void*)**。
// Container/DragIcon 的记录字段是 `void*` (vb6forms_uc_internal.h:100-101), 此前通道给的是
// 0=BSTR —— 类型撒谎。写侧 3a 的闸门不看类型, 已经在往这条路上走; 而 ComGet/ComProp 的
// Extender 分支会把 kind 0 当 BSTR 走 `SysAllocString(对象指针)`。凡是有 kind switch 的
// 地方都必须认 4, **指针宽度**存取 —— 落进 default 的 int32 会把 64 位指针截断。
/* Fix 294 / 3a: `Extender.<成员> = v` 的实例感知写入口。
   cgen 过去对 Extender 成员发的是裸全局赋值 `vb6_Extender_X = v;` —— 而通道
   (vb6_UC_ExtenderField) 已经改指 `vb6_UCRec.ext*`, 裸写进不去, 用户代码改
   Extender.Height/Width/Left/Top 会**静默不生效**。这里按 `vb6_ucCurrent` 落进记录。
   没有当前记录时退回全局, 与 ExtenderField 的兜底保持一致。
   kind 约定同 vb6_UC_ExtenderField: 0=BSTR, 2=int16, 3=int32。 */
/* Fix 296 读侧: `Extender.<成员>` 作右值时的实例感知入口(与写侧对称)。
   与写侧刻意不同形: 读侧要参与算术/字符串表达式, 返回 void* 会把每个调用点都
   变成强制转换, 比原来更脏。所以拆成数值型 / 字符串型两个, 由调用方按成员类型选。
   kind 约定同 vb6_UC_ExtenderField: 0=BSTR / 2=int16 / 3=int32。 */
int32_t vb6_UC_GetExtenderMemberNum(const wchar_t* name) {
    extern void* vb6_UC_ExtenderField(void* p, const wchar_t* n, int32_t* k);
    void* p = (void*)&vb6_UserControl_Extender;
    int32_t kind = -1;
    void* fp = vb6_UC_ExtenderField(p, name, &kind);
    if (!fp) return 0;
    if (kind == 2) return (int32_t)(*(int16_t*)fp);
    /* 对象成员 (kind 4): 读 32 位就是**截断指针**, 交出去的是半个地址。
       调用方按类型表挑入口, 正常走不到这里; 这道是给"有人把读闸放开了"兜底。 */
    if (kind == 4) return 0;
    return (int32_t)(*(int32_t*)fp);
}

BSTR vb6_UC_GetExtenderMemberStr(const wchar_t* name) {
    extern void* vb6_UC_ExtenderField(void* p, const wchar_t* n, int32_t* k);
    void* p = (void*)&vb6_UserControl_Extender;
    int32_t kind = -1;
    void* fp = vb6_UC_ExtenderField(p, name, &kind);
    if (!fp || kind != 0) return NULL;
    return *(BSTR*)fp;
}
/* Fix 296 / 3b 最后一格: Container/DragIcon 的**读**入口 (对象成员, 返回 void*)。
   这两个的返回槽实测就是 `void* vb6_ret_Container` (ComCtlsDemo CoolBar.c:1651),
   指针宽度正好对上 —— 调用点不需要任何强制转换, 与数值/字符串两个入口同一形状。
   **只认 kind 4**: BSTR 槽 (kind 0) 当指针交出去, 调用方很可能 SysFreeString 掉它 ——
   那正是引入 kind 4 要挡住的事。 */
void* vb6_UC_GetExtenderMemberObj(const wchar_t* name) {
    extern void* vb6_UC_ExtenderField(void* p, const wchar_t* n, int32_t* k);
    void* p = (void*)&vb6_UserControl_Extender;
    int32_t kind = -1;
    void* fp = vb6_UC_ExtenderField(p, name, &kind);
    if (!fp || kind != 4) return NULL;
    return *(void**)fp;
}
/* 账 #305 第二百八十七刀: 「Extender.Visible 这一发要落到宿主窗口上」收成一位主人。
   两处写手(本文件按名的那一条 + vb6com_invoke.c 的 IDispatch Extender 臂)存完字段都转调这里。
   从前只把值存进 rc->extVisible 而**没有任何东西读它**, 于是容器侧 `SpinBox1.Visible = False`
   (它发成控件自己的 Property Let, Let 体就是 `Extender.Visible = Value`) 静默不隐藏 ——
   实测靶工程 TabStrip1 切回第一页后, 那一页的 SpinBox / ListView 两台都还留着。
   窗口那一侧早已有主人 `vb6_SetControlVisible`(原生控件那一路用的就是它), 这里只转调、不抄第二份。 */
void vb6_UC_ExtenderVisibleApplied(void) {
    vb6_UCRec* rc = vb6_ucCurrent;
    if (!rc || !rc->hwnd || !IsWindow((HWND)rc->hwnd)) return;
    vb6_SetControlVisible(rc->hwnd, rc->extVisible);
}

/* 账 #305 第二百八十九刀: `Extender` 的四档几何 (Left/Top/Width/Height) 与那一发 `Move`。
   VB6 的 Extender 几何不是第二本账: 读 = 这枚宿主窗口当前的矩形按**容器 ScaleMode** 折出来的
   数; 写 = 向容器请求把这块矩形改成那样。我方从前只把写落进 rc->ext* (建宿主时填的设计缇值,
   之后再没人刷新), 而 `Move` 那一路连应答都没有 —— 第二百八十八刀的读数: 空跑 8 秒 29 发
   method "Move" not found, 于是控件自己量出来的尺寸既进不去也读不出。
   两头都交回**已有**的那两位主人, 这里不抄第二份:
     读 = vb6_GetControlLeft/Top/Width/Height (自带 #230 的 VB 侧缓存与像素闸, 单位由
          vb6_ContainerScaleMode 答 —— 即 #175 那一位主人);
     写 = vb6_ControlMove (所有摆位的唯一收口: 按容器 ScaleMode 换像素、写 #230 缓存、
          尺寸真的变了才触发 Resize)。
   刻意不在这里判「交来的值合不合法」: 第二百八十四刀记的「排版只跑一次、跑在属性到齐之前」
   是上游那一格, 不在这一问里补兜底。 */
static int32_t vb6_ucExtGeomMask(const wchar_t* name) {
    if (!name) return 0;
    if (_wcsicmp(name, L"Left")   == 0) return 1;
    if (_wcsicmp(name, L"Top")    == 0) return 2;
    if (_wcsicmp(name, L"Width")  == 0) return 4;
    if (_wcsicmp(name, L"Height") == 0) return 8;
    return 0;
}

/* 读侧: 交出去之前按活宿主窗口把四档问一遍。读写都过 vb6_UC_ExtenderField, 所以这一问同时治
   掉「读到建宿主时的死值」与「写之前另三档还是死值」两格。没有窗口 (建宿主前 / 无窗宿主) 就
   保留记录里那个设计值 —— 与今天一致, 不凭空造数。 */
static void vb6_ucExtGeomRefresh(vb6_UCRec* rc) {
    if (!rc->hwnd || !IsWindow(rc->hwnd)) return;
    rc->extLeft   = vb6_GetControlLeft((void*)rc->hwnd);
    rc->extTop    = vb6_GetControlTop((void*)rc->hwnd);
    rc->extWidth  = vb6_GetControlWidth((void*)rc->hwnd);
    rc->extHeight = vb6_GetControlHeight((void*)rc->hwnd);
}

/* 容器 ScaleMode 的那一问, 给 vb6rtl_com.c 的 vbContainerPosition(9) / vbContainerSize(10)
   两档用。答 0 = 问不到 (没有当前记录或父窗没了), 调用方保留它今天那档。 */
/* 账 #305 第二百九十二刀: Extender 四档几何的**属性形写手**（`.Extender.Height = v`）。
   289 那批留下读侧投影、把写整块撤回 —— 当时写一落地就有两处塌缩；逐窗口归因后看清：**四格好转全部
   来自属性形**（ToolBar 32→26 / MonthView 160x161 / ListBoxW 170 / ComboBoxW 19），**两处塌缩全部来自
   方法形** `.Extender.Move`（CoolBar.ctl:996 与 ComboBoxW.ctl:1791）⇒ 两形可以分开落地：这一刀只接
   属性形，`Move` 继续没人应答（那 29 发 not found 照旧，等输入量清）。写交唯一收口 vb6_ControlMove。 */
void vb6_UC_ExtenderGeomApplied(const wchar_t* name) {
    vb6_UCRec* rc = vb6_ucCurrent;
    if (!rc || !rc->hwnd || !IsWindow((HWND)rc->hwnd)) return;
    int32_t which = vb6_ucExtGeomMask(name);
    if (!which) return;
    extern void vb6_ControlMove(void* hwnd, double l, double t, double w, double h, int mask);
    vb6_ControlMove((void*)rc->hwnd, (double)rc->extLeft, (double)rc->extTop,
                    (double)rc->extWidth, (double)rc->extHeight, (int)which);
}

/* 账 #305 第三百零一刀: Extender 四档几何的**方法形写手**（`.Extender.Move l, t, w, h`）——
   控件自己量完尺寸、来请求摆位的那一发。第二百九十五刀把它改成「响亮地丢」；当时第一次落地
   的三个塌法（CoolBar 60→37 / ComboBoxW 288↔504 来回 / ImageCombo 3672 像素）的前置都变了：
   第二百九十七刀把「几何按谁的 ScaleMode 折」收到容器主人（不再问 GetParent），第二百九十九刀
   把 authored 容器戳到窗口上，第二百九十刀补了 ReBar band 行高那一路输入。
   本条只补通道、兜底一概不抄 —— 四档原样交给 vb6_ControlMove（所有摆位的唯一收口：按容器
   ScaleMode 换像素、写 #230 缓存、尺寸真变了才排队 Resize）。VB6 的 Move 允许多种实参个数：
   给几档就只动几档（mask），一档都没有就什么也不做。
   实测靶工程：MonthView.ctl 的 `UserControl_Resize` 靠它把宿主缩到自己量出的 160x161（参考那台
   同形）；我方从前停在设计期 172x156，内层 SysMonthCal32 跟着大一圈。 */
int32_t vb6_UC_ExtenderMoveApply(double L, double T, double W, double H, int32_t mask) {
    vb6_UCRec* rc = vb6_ucCurrent;
    if (!rc || !rc->hwnd || !IsWindow((HWND)rc->hwnd)) return 0;
    if (!mask) return 0;
    extern void vb6_ControlMove(void* hwnd, double l, double t, double w, double h, int mask);
    vb6_ControlMove((void*)rc->hwnd, L, T, W, H, (int)mask);
    return 1;
}

int32_t vb6_UC_ContainerScaleModeCurrent(void) {
    vb6_UCRec* rc = vb6_ucCurrent;
    if (!rc) return 0;
    /* 账 #305 第297刀: 容器那一问只有一位主人 —— 有宿主窗口就问它 (答创建时记录的容器,
       不随 SetParent 漂); 没有窗口 (无窗宿主 / 建宿主前) 才落回记录里那个父。 */
    void* cw = rc->hwnd ? vb6_UC_ContainerWindowOf((void*)rc->hwnd) : (void*)rc->parent;
    if (!cw || !IsWindow((HWND)cw)) return 0;
    extern int32_t vb6_ContainerScaleMode(void* hwndParent);
    return vb6_ContainerScaleMode(cw);
}

void vb6_UC_SetExtenderMember(const wchar_t* name, void* value) {
    extern void* vb6_UC_ExtenderField(void* p, const wchar_t* n, int32_t* k);
    void* p = (void*)&vb6_UserControl_Extender;
    int32_t kind = -1;
    void* fp = vb6_UC_ExtenderField(p, name, &kind);
    if (!fp) return;
    /* 账 #305 第二百八十七刀: 这道早退原本还带一条 `|| !value` ⇒ **凡递来 0 的写全被静默丢掉**,
       而 kind 2 那三档布尔（Visible / vb6_Default / Cancel）递来的「假」正是 0。实测靶工程
       `SpinBox1.Visible = False` 就是这么没的：容器侧那一发走到这里, 字段没存、窗口也没动。
       第二百九十二刀把几何接到窗口之后, kind 3 也进「放开」那一档：语料里 `Property Let Left(Value As
       Long)` 的体就是 `Extender.Left = Value`, 容器侧一句 `Label1.Left = 0` 递来的正是 0。kind 0/4
       （Tag / ToolTipText / Container / DragIcon）的零值写今天仍没有读数, 刻意不动（在册）。 */
    if (!value && kind != 2 && kind != 3) return;
    switch (kind) {
        case 0: *(BSTR*)fp  = (BSTR)value; break;                 // BSTR: 搬指针
        case 2: *(int16_t*)fp = (int16_t)(intptr_t)value; break;
        // 对象 (kind 4): **原样搬指针**, 宽度不变。漏掉这一支会落进 default 的
        // *(int32_t*) —— x64 上把 64 位指针截成 32 位 (Fix 296 / 3b 前置)。
        case 4: *(void**)fp = value; break;
        default: *(int32_t*)fp = (int32_t)(intptr_t)value; break;
    }
    if (kind == 2 && _wcsicmp(name, L"Visible") == 0) vb6_UC_ExtenderVisibleApplied();
    else if (kind == 3) vb6_UC_ExtenderGeomApplied(name);   /* 账 #305 第二百九十二刀: 属性形写手 */
}
void* vb6_UC_ExtenderField(void* p, const wchar_t* name, int32_t* kind) {
    if (!vb6_UC_IsExtender(p) || !name || !kind) return NULL;
    /* Fix 294: 几何四项走**当前实例**的记录, 不再走裸全局。
       `vb6_UCRec.extLeft/extTop/extWidth/extHeight` 本来就在(建宿主时写的设计值),
       但之前**没有任何东西读它们** —— 于是 `Extender.Height` 读回来是全局槽位里的
       另一个值(或 0), `.Extender.Height = v` 写进去也只落在全局上, 与宿主实际尺寸
       脱节。这里让 COM Extender 通道直接指向记录字段: 读写都落在"这个控件自己的"
       那一份上, 天然不怕 push/pop 互相盖。
       没有当前记录时(理论上不该发生)退回全局, 保持原行为。 */
    {
        vb6_UCRec* rc = vb6_ucCurrent;
        if (rc) {
            if (vb6_ucExtGeomMask(name)) vb6_ucExtGeomRefresh(rc);   /* 账 #305 第二百八十九刀: 读侧投影 */
            if (_wcsicmp(name, L"Width")  == 0) { *kind = 3; return &rc->extWidth;  }
            if (_wcsicmp(name, L"Height") == 0) { *kind = 3; return &rc->extHeight; }
            if (_wcsicmp(name, L"Left")   == 0) { *kind = 3; return &rc->extLeft;   }
            if (_wcsicmp(name, L"Top")    == 0) { *kind = 3; return &rc->extTop;    }
            /* Fix 294 第二片: 其余成员同样走记录。名字与全局一一对应, 保证行为不变。 */
            if (_wcsicmp(name, L"Align")           == 0) { *kind = 3; return &rc->extAlign;           }
            if (_wcsicmp(name, L"DragMode")        == 0) { *kind = 3; return &rc->extDragMode;        }
            if (_wcsicmp(name, L"HelpContextID")   == 0) { *kind = 3; return &rc->extHelpContextID;   }
            if (_wcsicmp(name, L"WhatsThisHelpID") == 0) { *kind = 3; return &rc->extWhatsThisHelpID; }
            if (_wcsicmp(name, L"Visible")         == 0) { *kind = 2; return &rc->extVisible;         }
            if (_wcsicmp(name, L"vb6_Default")     == 0) { *kind = 2; return &rc->extVb6Default;      }
            if (_wcsicmp(name, L"Cancel")          == 0) { *kind = 2; return &rc->extCancel;          }
            if (_wcsicmp(name, L"Tag")             == 0) { *kind = 0; return &rc->extTag;             }
            if (_wcsicmp(name, L"ToolTipText")     == 0) { *kind = 0; return &rc->extToolTipText;     }
            // 对象成员: 字段是 void*, **不是** BSTR (Fix 296 / 3b 前置)
            if (_wcsicmp(name, L"Container")       == 0) { *kind = 4; return &rc->extContainer;       }
            if (_wcsicmp(name, L"DragIcon")        == 0) { *kind = 4; return &rc->extDragIcon;        }
        }
    }
    if (_wcsicmp(name, L"Width") == 0)           { *kind = 3; return &vb6_Extender_Width; }
    if (_wcsicmp(name, L"Height") == 0)          { *kind = 3; return &vb6_Extender_Height; }
    if (_wcsicmp(name, L"Left") == 0)            { *kind = 3; return &vb6_Extender_Left; }
    if (_wcsicmp(name, L"Top") == 0)             { *kind = 3; return &vb6_Extender_Top; }
    if (_wcsicmp(name, L"Align") == 0)           { *kind = 3; return &vb6_Extender_Align; }
    if (_wcsicmp(name, L"Visible") == 0)         { *kind = 2; return &vb6_Extender_Visible; }
    if (_wcsicmp(name, L"DragMode") == 0)        { *kind = 3; return &vb6_Extender_DragMode; }
    if (_wcsicmp(name, L"HelpContextID") == 0)   { *kind = 3; return &vb6_Extender_HelpContextID; }
    if (_wcsicmp(name, L"WhatsThisHelpID") == 0) { *kind = 3; return &vb6_Extender_WhatsThisHelpID; }
    if (_wcsicmp(name, L"Tag") == 0)             { *kind = 0; return &vb6_Extender_Tag; }
    if (_wcsicmp(name, L"ToolTipText") == 0)     { *kind = 0; return &vb6_Extender_ToolTipText; }
    /* Fix 296 / 3b 前置: 兜底段必须把 vb6_uc_push/pop 管的 15 枚**全兜住**。
       此前漏了 4 枚 (vb6_Default / Cancel / Container / DragIcon) —— 兜不住的那几枚,
       在**没有当前记录**时 vb6_UC_SetExtenderMember 撞 `if (!fp) return;`,
       写**静默丢** (改之前那枚全局落地, 改之后没了)。实测 DragIcon 33 个写点、
       Cancel 2 个。这条也是 3b 的前置: push/pop 一旦不再镜像全局, 兜底就是唯一的退路。 */
    if (_wcsicmp(name, L"vb6_Default") == 0)     { *kind = 2; return &vb6_Extender_vb6_Default; }
    if (_wcsicmp(name, L"Cancel") == 0)          { *kind = 2; return &vb6_Extender_Cancel; }
    if (_wcsicmp(name, L"Container") == 0)       { *kind = 4; return (void*)&vb6_Extender_Container; }
    if (_wcsicmp(name, L"DragIcon") == 0)        { *kind = 4; return (void*)&vb6_Extender_DragIcon; }
    return NULL;
}

// Fix 128: 把 src 字体的 8 个字段拷进 dst 字体对象 (原地覆写)。
// 用途: .frm 的 `BeginProperty Caption1_Font` 这类**Property Set** 型字体属性,
// 其 Set 实现体是 `With m_X_Font: .Name = New_Font.Name ... : Refresh` —— 直接调用
// 会在 CreateControls 阶段(Form_Load 尚未填数据)触发 Refresh, 使图表永久停在空状态
// (实测柱/面积/树/饼 只剩标题: 彩色像素 0.3~3%, 关闭后 10~59%)。
// 因此改为: 用该控件自己的 Property Get 取到内部字体对象, 直接覆写字段, 不触发 Set/Refresh。
void vb6_UC_FontAssign(void* dst, void* src) {
    vb6_ComIface_Font* d = vb6_uc_fontOf(dst);
    vb6_ComIface_Font* s = vb6_uc_fontOf(src);
    if (!d || !s || d == s) return;
    if (s->Name) d->Name = SysAllocString(s->Name);
    d->Size = s->Size;
    d->Bold = s->Bold;
    d->Italic = s->Italic;
    d->Underline = s->Underline;
    d->Strikethrough = s->Strikethrough;
    d->Weight = s->Weight;
    d->Charset = s->Charset;
}

#ifdef __cplusplus
} // extern "C"
#endif
