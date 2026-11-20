/* vb6forms_menuproxy.c - 跨模块菜单项代理 (C30-8)
 *
 * 为什么需要这东西 (江西省医保药品查询系统 实测, frmlogin.frm:169 附近):
 *     Main.Menu_Edit.Enabled = False
 *     Main.Menu_ManageAdd.Enabled = False
 * 另一窗体 (main.frm) 的菜单项被引用。VB6 的菜单项不是窗口 (FrmControlType::Menu
 * 不建 HWND)，它只是父 HMENU 里的一个**命令 ID**；而 cgen 对跨模块控件一律发
 * `vb6_form_ctrl_<Form>_<Ctrl>()` 拿句柄，菜单没有这种句柄 ⇒ 访问器原本被
 * cgen_form_ctrl_registry.inc:159 明确跳过 ⇒ LNK2019:
 *     frmlogin.obj : error LNK2019: 无法解析的外部符号 _vb6_form_ctrl_Main_Menu_Edit
 *
 * 做法: 给菜单项发一枚**代理令牌**（不是句柄）。访问器侧:
 *     void* vb6_form_ctrl_Main_Menu_Edit(void) {
 *         return vb6_MenuProxyMake((void*)GetMenu((HWND)vb6_form_hwnd_Main()), 1003);
 *     }
 * 令牌里记 (HMENU, 命令ID) 两样，正是 vb6_SetMenu* / vb6_GetMenu* 家族要的入参。
 * 再由 vb6_ComSetProp / vb6_ComGetProp 在入口认令牌，把属性名换算成那套 API。
 * 语义等价于 VB6 的 `Form.MenuItem.Prop`（运行期按名字找到菜单项再改状态）。
 *
 * 边界:
 *   - 只认 Enabled / Checked / Visible / Caption 四个属性 (与 cgen_util_ctrl.cpp
 *     的 Menu 写表同一份口径，见 controlPropWriteFn 的 FrmControlType::Menu 分支)。
 *   - 其它属性名 → 返回"没接住"(0)，由调用方继续原有 COM 路径。
 *   - 菜单栏还没建好时 GetMenu 给 NULL ⇒ 令牌照发（去重键是 (NULL,id)），
 *     所有 setter 自身都有 `if (!menuHandle) return;` 兜底，不会崩。
 *   - 表容量 256：一个工程的菜单项远少于这个数；溢出返回 NULL (=空对象)。
 */
#include <windows.h>
#include <oleauto.h>
#include <stdint.h>
#include <wchar.h>
#include <string.h>
#include <stdlib.h>

/* 同源声明: 实现在 vb6forms_widget_prop.c。这里不 include 头文件 ——
 * vb6com 侧也要用本文件的三个出口, 为了不让 COM 层反向依赖窗体层头, 统一走 extern。 */
extern void vb6_SetMenuEnabled(void* menuHandle, int menuId, int32_t val);
extern void vb6_SetMenuChecked(void* menuHandle, int menuId, int32_t val);
extern void vb6_SetMenuVisible(void* menuHandle, int menuId, int32_t val);
extern void vb6_SetMenuCaption(void* menuHandle, int menuId, void* bstrCaption);
extern int32_t vb6_GetMenuEnabled(void* menuHandle, int menuId);
extern int32_t vb6_GetMenuChecked(void* menuHandle, int menuId);
extern int32_t vb6_GetMenuVisible(void* menuHandle, int menuId);

typedef struct {
    void*   hMenu;   /* 父菜单栏 (GetMenu(窗体HWND)) */
    int32_t id;      /* 菜单命令 ID (建菜单时的 1000+ 自增) */
} Vb6MenuProxySlot;

static Vb6MenuProxySlot g_menuProxy[256];
static int32_t g_menuProxyCount = 0;

/* 取\建令牌。同一个 (hMenu,id) 只发一枚 —— 令牌地址被生成码长期持有,
 * 每调一次就新开一格会把表撑爆。 */
void* vb6_MenuProxyMake(void* hMenu, int32_t id) {
    int32_t i;
    /* id == 0 不是有效项: 叶子项拿的是 1000+ 的命令 ID, 顶层弹出菜单拿的是
     * 负数 (-1-位置)。0 这两个都不是 ⇒ 当成没找到。 */
    if (id == 0) return NULL;
    for (i = 0; i < g_menuProxyCount; i++) {
        if (g_menuProxy[i].hMenu == hMenu && g_menuProxy[i].id == id) {
            return (void*)&g_menuProxy[i];
        }
    }
    if (g_menuProxyCount >= (int32_t)(sizeof(g_menuProxy) / sizeof(g_menuProxy[0]))) {
        return NULL;
    }
    g_menuProxy[g_menuProxyCount].hMenu = hMenu;
    g_menuProxy[g_menuProxyCount].id = id;
    return (void*)&g_menuProxy[g_menuProxyCount++];
}

/* 认令牌。返回 1 表示这是菜单代理 (而不是 IDispatch/HWND)。 */
int32_t vb6_MenuProxyDecode(void* p, void** hMenu, int32_t* id) {
    int32_t i;
    if (!p) return 0;
    for (i = 0; i < g_menuProxyCount; i++) {
        if ((void*)&g_menuProxy[i] == p) {
            if (hMenu) *hMenu = g_menuProxy[i].hMenu;
            if (id) *id = g_menuProxy[i].id;
            return 1;
        }
    }
    return 0;
}

/* 写属性。value_void 是 vb6_ComPack* 出来的 VARIANT*。返回 1 = 已处理。
 * id > 0 → 命令 ID (叶子菜单项), 走 vb6_SetMenu* 家族;
 * id < 0 → 顶层弹出菜单, 位置 = -1-id, 只能按 MF_BYPOSITION 操作
 *          (MF_POPUP 项没有命令 ID)。 */
int32_t vb6_MenuProxySetProp(void* hMenu, int32_t id, const wchar_t* propName, void* value_void) {
    VARIANT* pv = (VARIANT*)value_void;
    int32_t on = 0;
    if (!propName) return 0;
    if (pv) {
        on = (pv->vt == VT_BOOL) ? (pv->boolVal != VARIANT_FALSE) : (pv->lVal != 0);
    }
    if (id < 0) {
        UINT pos = (UINT)(-1 - id);
        if (!hMenu) return 1;   /* 菜单栏还没建: 忽略赋值, 与 VB6 未加载窗体同效 */
        if (_wcsicmp(propName, L"Enabled") == 0) {
            EnableMenuItem((HMENU)hMenu, pos, MF_BYPOSITION | (on ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
            return 1;
        }
        if (_wcsicmp(propName, L"Checked") == 0) {
            CheckMenuItem((HMENU)hMenu, pos, MF_BYPOSITION | (on ? MF_CHECKED : MF_UNCHECKED));
            return 1;
        }
        if (_wcsicmp(propName, L"Caption") == 0) {
            /* 弹层标题只能用 ModifyMenuW 按位置改, 且必须带 MF_POPUP + 子菜单句柄,
             * 否则原来的子菜单会被改没。 */
            HMENU sub = GetSubMenu((HMENU)hMenu, (int)pos);
            const wchar_t* txt = (pv && pv->vt == VT_BSTR) ? pv->bstrVal : L"";
            if (sub) {
                ModifyMenuW((HMENU)hMenu, pos, MF_BYPOSITION | MF_POPUP | MF_STRING,
                            (UINT_PTR)sub, txt);
            }
            return 1;
        }
        return 0;   /* Visible: 弹出菜单按位置隐藏没有可靠 API, 不接 */
    }
    if (_wcsicmp(propName, L"Enabled") == 0) {
        vb6_SetMenuEnabled(hMenu, id, on);
        return 1;
    }
    if (_wcsicmp(propName, L"Checked") == 0) {
        vb6_SetMenuChecked(hMenu, id, on);
        return 1;
    }
    if (_wcsicmp(propName, L"Visible") == 0) {
        vb6_SetMenuVisible(hMenu, id, on);
        return 1;
    }
    if (_wcsicmp(propName, L"Caption") == 0) {
        /* vb6_SetMenuCaption 第三参是 BSTR (wchar_t*), 不是 VARIANT —— 拆一层。 */
        void* txt = NULL;
        if (pv && pv->vt == VT_BSTR) txt = (void*)pv->bstrVal;
        vb6_SetMenuCaption(hMenu, id, txt);
        return 1;
    }
    return 0;
}

/* 读属性。认得的四个回一枚 VT_I4/VT_BSTR 的 VARIANT (calloc 出来的, 与
 * vb6_ComPack* 同形 —— 调用方 vb6_ComGetIntProp 里会 vb6_ComVarClear 掉);
 * 其它属性回 NULL, 调用方那一侧 `if (!pv) return 0;` 已经兜住。 */
void* vb6_MenuProxyGetProp(void* hMenu, int32_t id, const wchar_t* propName) {
    VARIANT* pv;
    int32_t lv = 0;
    if (!propName) return NULL;
    if (_wcsicmp(propName, L"Enabled") != 0 && _wcsicmp(propName, L"Checked") != 0
        && _wcsicmp(propName, L"Visible") != 0) {
        return NULL;   /* Caption 读取暂不接 (需要另一路取串, 见文件头"边界") */
    }
    if (id < 0) {
        /* 顶层弹出菜单: 按位置读状态。 */
        MENUITEMINFOW mii;
        UINT pos = (UINT)(-1 - id);
        if (!hMenu) return NULL;
        memset(&mii, 0, sizeof(mii));
        mii.cbSize = sizeof(mii);
        mii.fMask = MIIM_STATE;
        if (!GetMenuItemInfoW((HMENU)hMenu, pos, TRUE, &mii)) return NULL;
        if (_wcsicmp(propName, L"Enabled") == 0)      lv = (mii.fState & MFS_DISABLED) ? 0 : -1;
        else if (_wcsicmp(propName, L"Checked") == 0) lv = (mii.fState & MFS_CHECKED) ? -1 : 0;
        else                                          lv = (mii.fState & MFS_DISABLED) ? 0 : -1;
    } else {
        if (_wcsicmp(propName, L"Enabled") == 0)      lv = vb6_GetMenuEnabled(hMenu, id);
        else if (_wcsicmp(propName, L"Checked") == 0) lv = vb6_GetMenuChecked(hMenu, id);
        else                                          lv = vb6_GetMenuVisible(hMenu, id);
    }
    pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    if (!pv) return NULL;
    VariantInit(pv);
    pv->vt = VT_I4;
    pv->lVal = lv;
    return (void*)pv;
}
