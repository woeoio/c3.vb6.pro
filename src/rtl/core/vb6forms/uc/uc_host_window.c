// uc_host_window.c - vb6forms_uc 拆分片：宿主窗口过程 + 类注册 + GDI+ 初始化
//
// 内容 = 拆分前 vb6forms_uc.c 第 459~565 / 567~607 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"
#include <stdio.h>  /* rev10: C3_UC_TRACE 的 fprintf */
#include <stdlib.h> /* C3_AD_TRACE 的 getenv */

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 宿主窗口过程
// ============================================================

// ---------------------------------------------------------------------------
// 透明 UserControl (`.ctl` 里 `BackStyle = 0 'Transparent'`) 的宿主合成 —— 账 #492
//
// VB6 的 Windowless UC 直接画在容器上, 文字底下的兄弟控件内容天然可见。C3 的 UC
// 一直是真子窗口, 加上 WS_CLIPSIBLINGS 之后兄弟窗口绝不会画进本窗矩形, 于是
// BackStyle=0 的宿主必须自己把"下面那一层"补进离屏 DIB:
//   1) 先铺 ambient 底色 (容器 BackColor, 与 VB6 语义一致);
//   2) 再把 z 序在本窗**之下**、且矩形相交的兄弟窗口用 WM_PRINTCLIENT 画进来。
//      兄弟按 z 序从最底往上画 (后画者压住先画者), 与本窗的遮挡关系一致。
// 实测 (tests/Charts 2020): LabelPlus3 的 BackColorOpacity=0 只画文字, 之前宿主
// 拿 ambient 灰 (Form2.BackColor=E5E5E5) 铺满整块 ⇒ 白色卡片 LabelPlus2 被盖成灰。
// ---------------------------------------------------------------------------
static void vb6_uc_fillAmbient(HDC dc, int32_t w, int32_t h, int32_t oleAmbient) {
    RECT full = { 0, 0, w, h };
    COLORREF cref = vb6_ColorToColorRef(oleAmbient);
    HBRUSH b = CreateSolidBrush(cref);
    FillRect(dc, &full, b);
    DeleteObject(b);
}

/* 账 #305 第三百刀: VB6 交给 VB 代码的那张 hDC 上**选着的字体 = 这枚控件自己的 Font**。
   我方那两张离屏位图 DC 是 CreateCompatibleDC 造的, 缺省选着的是 GDI 的 SYSTEM_FONT
   (120 DPI 下 "FrameW1" 墨迹 70x12 像素) —— 于是用户代码里不经我方出口、直接拿 DC 落笔
   的写法就整块换字: ComCtlsDemo 第二页 FrameW 的题头又粗又大, 而参考那台同一枚
   = MS Sans Serif -13 常规体 (墨迹 55x12; .local/k306/render2.py 把 SYSTEM_FONT 那一档
   单独量出来正好 70x12 才定性的)。
   ⚠ 这里只许读那一格存储 (vb6_ControlFontStored), **不许**换成 vb6_ControlFont:
   后者要 SendMessage(WM_GETFONT), 而本站是从 vb6_uc_push/paint 里进来的 —— 实发过一发,
   进程当场 0xC0000409 (fastfail)。那两格答案本来就同源 (宿主窗口的 WM_GETFONT 只答这一格)。
   没交字体 (答空) 时不动 DC, 老行为不变。 */
static HFONT vb6_uc_fontHFontFor(const void* fontPtr, int16_t bold, int16_t underline);

static COLORREF vb6_uc_resolve_color(int32_t ole) {
    return vb6_ColorToColorRef(ole);
}

/* 只读宿主窗口那一格存储 (WM_GETFONT 同源)。**不能**优先用 vb6_ucCurrent->font:
   这里可能是任意一枚 r 的 DC —— 非 AutoRedraw 的临时 DIB 在 vb6_uc_push **之前**
   就要选字体, 而 vb6_ucCurrent 还停在上一个 UC 上, 它的 font 可能是别的实例 /
   别的时刻留下的 (实测 ucSegmentedControl 的 CheckBar 因此拿到 MS Sans Serif,
   PUA 字形 (U+E16C..) 全部不画)。控件自己 `Set UserControl.Font = ...` 的逐段
   换字体走 vb6_uc_applyFontNow (那里 hdc 已经指向本实例的绘制 DC)。 */
static void vb6_uc_selectOwnFont(HDC dc, HWND hwnd) {
    HFONT hf;
    if (!dc || !hwnd) return;
    hf = vb6_ControlFontStored(hwnd);
    if (hf) SelectObject(dc, hf);
}

// ---------------------------------------------------------------------------
// ucTreeView: `UserControl.Font` / `.FontBold` / `.FontUnderline` 落到当前 paint DC
//
// VB6 里这三者是 UserControl.Font 对象 (FontBold/FontUnderline 就是它的 Bold/Underline),
// 生成 C 里却是裸全局赋值 (`vb6_UserControl_Font = ...`) —— 没有任何站点把它们选进
// hDC。ucTreeViewPlus.ctl 的 Draw 里逐节点 `Set UserControl.Font = m_SubTextFont` /
// `UserControl.FontBold = ...` 之后立刻 `DrawTextW hdc`, 所以必须**当场**改 DC 字体,
// 否则全部节点都用 ambient 字体画 (子文字不变小、粗体不粗)。
//
// HFONT 由 (字体对象, bold, underline, Size) 缓存: 一次 paint 只有两三枚字体,
// CreateFontIndirect 的次数也就两三次; 缓存里的 HFONT 不销毁 (数量被工程里的
// Font 对象数封顶, 逐个 DeleteObject 反而会撞"对象还选在 DC 里")。
// ---------------------------------------------------------------------------
typedef struct vb6_UCHfCacheEnt {
    const void* font;
    float       size;
    int16_t     bold;
    int16_t     underline;
    HFONT       hf;
} vb6_UCHfCacheEnt;
#define VB6_UC_HF_CACHE_MAX 16
static vb6_UCHfCacheEnt vb6_ucHfCache[VB6_UC_HF_CACHE_MAX];
static int32_t vb6_ucHfCacheN = 0;

static HFONT vb6_uc_fontHFontFor(const void* fontPtr, int16_t bold, int16_t underline) {
    const vb6_ComIface_Font* f = (const vb6_ComIface_Font*)fontPtr;
    LOGFONTW lf;
    int32_t i;
    vb6_ComIface_Font scratch;
    if (!f) return NULL;
    for (i = 0; i < vb6_ucHfCacheN; i++) {
        if (vb6_ucHfCache[i].font == fontPtr && vb6_ucHfCache[i].size == f->Size
            && vb6_ucHfCache[i].bold == bold && vb6_ucHfCache[i].underline == underline)
            return vb6_ucHfCache[i].hf;
    }
    scratch = *f;                      /* 只改两档, 不污染共享 Font 对象 */
    scratch.Bold = bold ? -1 : 0;
    scratch.Underline = underline ? -1 : 0;
    /* 质量档 = DEFAULT_QUALITY: VB6 的 `UserControl.Font` 是普通 StdFont, 交给 GDI 按
       系统设置平滑 (本机 ClearType)。以前这里写死 NONANTIALIASED ⇒ 节点文字是全硬边
       (实测真 VB6 "Human Resources" 一带 sat>=60 有 321 px、maxSat=112, C3 全是
       sat<=14、maxSat=9), 五窗 vivid/photo/avatar 因此系统性偏少。 */
    vb6_UC_FontLogFont(&scratch, vb6_DpiY(), DEFAULT_QUALITY, &lf);
    {
        HFONT hf = CreateFontIndirectW(&lf);
        if (hf && vb6_ucHfCacheN < VB6_UC_HF_CACHE_MAX) {
            vb6_UCHfCacheEnt* e = &vb6_ucHfCache[vb6_ucHfCacheN++];
            e->font = fontPtr; e->size = f->Size;
            e->bold = bold; e->underline = underline; e->hf = hf;
        }
        return hf;
    }
}

/* 把当前 UserControl.Font + 两个标志选进**当前**绘图 DC:
   WM_PAINT 里是 r->hdc; 事件/度量里是 AutoRedraw 持久表面 (UserControl.hDC 就是它)。
   两处都必须落到, 否则 `Set UserControl.Font` 只在 paint 里生效 —— 而 ucTreeView
   的 Draw/度量全发生在事件里, 于是主文本/子文本一直用 ambient 字体。 */
/* VB6 的 `UserControl.Font` / `FontBold` 改的是**控件窗口**的字体 (WM_GETFONT 读回的
   就是它); C3 以前只把字体选进绘图 DC, 窗口字体一直停在创建时的设计值 —— 参考产物
   Form3/Form5 读回 Segoe UI/-16/700 (最后一次 Draw 的 FontBold 状态), C3 却读回
   ambient 8.25pt/400。这里把同一份字体发到窗口; 句柄所有权仍交既有的
   `vb6_SetControlFontFromLogFont` (它自带「只删我们自己造的那张」的账, 不能把缓存
   句柄直接递给 WM_SETFONT)。相同字体不重复发 —— 否则每个节点一次 InvalidateRect。 */
static void vb6_uc_syncHostFont(HFONT hf) {
    HWND hwnd = vb6_ucCurrent ? (HWND)vb6_ucCurrent->hwnd : NULL;
    LOGFONTW want, cur;
    HFONT hCur;
    if (!hwnd || !hf) return;
    if (!GetObjectW((HGDIOBJ)hf, sizeof(want), &want)) return;
    /* 只许读那一格存储 (vb6_ControlFontStored), 不许裸问 WM_GETFONT ——
       本类自己答那一问 (case WM_GETFONT), 裸问等于第二处口径 (check_control_dc D11)。 */
    hCur = vb6_ControlFontStored(hwnd);
    if (hCur && GetObjectW((HGDIOBJ)hCur, sizeof(cur), &cur)
        && cur.lfHeight == want.lfHeight && cur.lfWeight == want.lfWeight
        && cur.lfItalic == want.lfItalic && cur.lfUnderline == want.lfUnderline
        && cur.lfCharSet == want.lfCharSet
        && wcscmp(cur.lfFaceName, want.lfFaceName) == 0) {
        return;
    }
    vb6_SetControlFontFromLogFont(hwnd, &want);
}

static void vb6_uc_applyFontNow(void) {
    HFONT hf;
    HDC dc;
    if (!vb6_ucCurrent) return;
    /* Font 属性是持久的: 事件外赋值也要写回实例记录, 下一次 push 才不会被退回 ambient。 */
    if (vb6_UserControl_Font != vb6_ucCurrent->font)
        vb6_ucCurrent->font = (vb6_ComIface_Font*)vb6_UserControl_Font;
    dc = vb6_ucCurrent->hdc ? vb6_ucCurrent->hdc : vb6_ucCurrent->adSurf.memDC;
    hf = vb6_uc_fontHFontFor(vb6_UserControl_Font, vb6_UserControl_FontBold,
                             vb6_UserControl_FontUnderline);
    if (hf && dc) SelectObject(dc, hf);
    vb6_uc_syncHostFont(hf);
}

void vb6_UC_SetFont(void* font) {
    vb6_UserControl_Font = (vb6_ComIface_Font*)font;
    vb6_uc_applyFontNow();
}

void vb6_UC_SetFontFlag(int32_t which, int32_t value) {
    if (which == 0) {
        vb6_UserControl_FontBold = value ? -1 : 0;
        if (vb6_ucCurrent) vb6_ucCurrent->fontBold = vb6_UserControl_FontBold;
    } else {
        vb6_UserControl_FontUnderline = value ? -1 : 0;
        if (vb6_ucCurrent) vb6_ucCurrent->fontUnderline = vb6_UserControl_FontUnderline;
    }
    vb6_uc_applyFontNow();
}

/* `UserControl.Extender.ZOrder v` (结构体成员函数指针的落点)。
   VB6 ZOrder: 0 = 提到最前, 1 = 送到最后。 */
void vb6_UC_ExtenderZOrder(int32_t* value, int32_t hasValue) {
    HWND hwnd = vb6_ucCurrent ? vb6_ucCurrent->hwnd : NULL;
    if (!hwnd || !IsWindow(hwnd)) return;
    SetWindowPos(hwnd, (value && hasValue && *value == 1) ? HWND_BOTTOM : HWND_TOP,
                 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
}

// ============================================================
// 账 #305 第二百九十二刀: AutoRedraw 持久表面
//
// VB6: `UserControl.AutoRedraw = True` 之后, UserControl.hDC 就是一张**常驻离屏位图**,
// 在 Resize/Refresh 里画的像素不会随下一次 WM_PAINT 消失, 重绘时容器再把它贴到屏幕上。
// C3 此前是立即模式 (hDC 只在 WM_PAINT 期间非空、Cls 空转) ⇒ ComCtlsDemo 的 FrameW
// 只在 UserControl_Resize 里画边框和标题, 永远不会出现在屏幕上 (第二页整块白板)。
// 只在 VB 代码**写过** .AutoRedraw = True 时才建表面 (adOn), 其余控件保持原路径。
// ============================================================
static void vb6_uc_adEnsure(vb6_UCRec* r) {
    if (!r || !r->hwnd) return;
    RECT rc;
    /* 宿主窗口销毁后 GetClientRect 失败而 rc 保持未初始化 —— 拿它当尺寸去
       dibDestroy/dibCreate, 会把最后一次 Draw 画好的持久表面换成一张垃圾表面,
       此后再怎么 Refresh 都是空白 (见 handoff §51: Contact List 第二次打开)。 */
    if (!IsWindow((HWND)r->hwnd) || !GetClientRect((HWND)r->hwnd, &rc)) {
        if (getenv("C3_AD_TRACE")) {
            fprintf(stderr, "[AD-skip] type=%s hwnd=%p win=%d\n",
                    r->desc ? r->desc->typeName : "?", (void*)r->hwnd,
                    (int)IsWindow((HWND)r->hwnd));
            fflush(stderr);
        }
        return;
    }
    int32_t w = rc.right - rc.left, h = rc.bottom - rc.top;
    if (w <= 0 || h <= 0) return;
    if (r->adSurf.memDC && r->adSurf.w == w && r->adSurf.h == h) return;
    vb6_uc_dibDestroy(&r->adSurf);
    vb6_uc_dibCreate(&r->adSurf, NULL, w, h);
    if (r->adSurf.memDC) {
        vb6_uc_fillAmbient(r->adSurf.memDC, w, h, r->ambientBack);
        /* VB6 的 hDC 默认 TRANSPARENT (GDI 新 DC 的默认是 OPAQUE + 白底), 且文本色
           跟着 UserControl.ForeColor。AutoRedraw 的绘制全发生在这张持久表面上,
           少这一步就会把每个字画成白底方块。 */
        SetBkMode(r->adSurf.memDC, TRANSPARENT);
        SetTextColor(r->adSurf.memDC, vb6_uc_resolve_color(vb6_UserControl_ForeColor));
    }
    if (getenv("C3_AD_TRACE")) {
        fprintf(stderr, "[AD-ensure] type=%s hwnd=%p w=%d h=%d\n",
                r->desc ? r->desc->typeName : "?", (void*)r->hwnd, (int)w, (int)h);
        fflush(stderr);
    }
}

HDC vb6_uc_autoRedrawDC(vb6_UCRec* r) {
    vb6_uc_adEnsure(r);
    if (!r || !r->adSurf.memDC) return NULL;
    vb6_uc_selectOwnFont(r->adSurf.memDC, r->hwnd);   /* 账 #305 300: 交出去的那张 DC 得带自己的字体 */
    return r->adSurf.memDC;
}

void vb6_uc_autoRedrawSet(vb6_UCRec* r, int32_t on) {
    if (!r) return;
    r->adOn = on ? 1 : 0;
    if (on) vb6_uc_adEnsure(r);
    /* 关掉 AutoRedraw **不销毁**表面: VB6 这边紧接着 `Set .Picture = .Image` 把当前
       位图交给容器显示 (FrameW DrawFrame 末尾就是这个顺序), 画面必须留下来。表面
       转成"底图", 下一次 WM_PAINT 先铺背景再把它贴上去, 立即模式语义不变。 */
}

void vb6_uc_autoRedrawCls(vb6_UCRec* r) {
    if (!r || !r->adSurf.memDC) return;
    vb6_uc_fillAmbient(r->adSurf.memDC, r->adSurf.w, r->adSurf.h, r->ambientBack);
}

/* 当前正在绘制的 UC 目标 DC (NULL = 不在绘制中)。见 vb6_UC_SetForeColor。 */
static HDC vb6_uc_activePaintDC = NULL;

/* 账 <font-color>: 真 VB6 的 `UserControl.ForeColor = v` 立即改变 UserControl.hDC
   的文本色 (控件随后 DrawTextW 画的标题/图例就吃这个色)。cgen 把这条赋值发到本函数
   (见 cgen_assign_host_pseudo.inc), 不再只写全局变量。 */
void vb6_UC_SetForeColor(int32_t ole) {
    vb6_UserControl_ForeColor = ole;
    /* 目标 DC 不止 paint 那一档: ucTreeView 的 Draw 发生在事件里, 画在 AutoRedraw
       持久表面 (UserControl.hDC) 上, 那才是此刻要改色的 DC。 */
    HDC dc = vb6_uc_activePaintDC ? vb6_uc_activePaintDC : (HDC)vb6_UserControl_hDC;
    if (dc) SetTextColor(dc, vb6_uc_resolve_color(ole));
}

/* ---------------------------------------------------------------------------
   缺陷 1 (禁用项光标): `UserControl.MousePointer = v` 的落地点。
   .ctl 在 MouseMove 里对 Disabled 项设 vbNoDrop(12)、对可用项设 vbCustom(99),
   之后自备 hCur 时再 SetCursor。C3 以前只把值写进裸全局、宿主窗口从不读,
   于是禁用项悬停仍是箭头 (实测 VB6 参考 IDC_NO=65559, C3 = 箭头)。
   现在: 存进**本实例**记录, 并立即把光标应用到宿主窗口。
   映射按 VB6 MousePointerConstants (1..15 是系统光标, 99=vbCustom 不映射)。
   --------------------------------------------------------------------------- */
static int32_t vb6_ucPaintDepth = 0;
void vb6_UC_PaintEnter(void) { vb6_ucPaintDepth++; }
void vb6_UC_PaintLeave(void) { if (vb6_ucPaintDepth > 0) vb6_ucPaintDepth--; }
int32_t vb6_UC_PaintDepth(void) { return vb6_ucPaintDepth; }

static LPCWSTR vb6_uc_cursorName(int32_t mp) {
    switch (mp) {
    case 1:  return (LPCWSTR)(intptr_t)32512; /* IDC_ARROW        vbArrow        */
    case 2:  return (LPCWSTR)(intptr_t)32515; /* IDC_CROSS        vbCrosshair    */
    case 3:  return (LPCWSTR)(intptr_t)32513; /* IDC_IBEAM        vbIbeam        */
    case 4:  return (LPCWSTR)(intptr_t)32641; /* IDC_ICON         vbIconPointer  */
    case 5:  return (LPCWSTR)(intptr_t)32646; /* IDC_SIZEALL      vbSizePointer  */
    case 6:  return (LPCWSTR)(intptr_t)32643; /* IDC_SIZENESW     vbSizeNESW     */
    case 7:  return (LPCWSTR)(intptr_t)32645; /* IDC_SIZENS       vbSizeNS       */
    case 8:  return (LPCWSTR)(intptr_t)32642; /* IDC_SIZENWSE     vbSizeNWSE     */
    case 9:  return (LPCWSTR)(intptr_t)32644; /* IDC_SIZEWE       vbSizeEW       */
    case 10: return (LPCWSTR)(intptr_t)32516; /* IDC_UPARROW      vbUpArrow      */
    case 11: return (LPCWSTR)(intptr_t)32514; /* IDC_WAIT         vbHourglass    */
    case 12: return (LPCWSTR)(intptr_t)32648; /* IDC_NO           vbNoDrop       */
    case 13: return (LPCWSTR)(intptr_t)32650; /* IDC_APPSTARTING  vbArrowHourglass */
    case 14: return (LPCWSTR)(intptr_t)32651; /* IDC_HELP         vbArrowQuestion */
    case 15: return (LPCWSTR)(intptr_t)32646; /* IDC_SIZEALL      vbSizeAll      */
    default: return NULL;                     /* 0=vbDefault / 99=vbCustom      */
    }
}

void vb6_UC_SetMousePointer(int32_t pointer) {
    vb6_UserControl_MousePointer = pointer;
    vb6_UCRec* r = vb6_ucCurrent;
    if (!r) return;
    r->mousePointer = pointer;
    LPCWSTR nm = vb6_uc_cursorName(pointer);
    if (!nm || !r->hwnd) return;
    /* 只在指针确实停在本控件的客户区时才立刻改光标 (VB6 的语义), 否则交给
       WM_SETCURSOR 在鼠标移进来时应用。 */
    POINT pt;
    if (!GetCursorPos(&pt)) return;
    RECT rc;
    if (!GetWindowRect(r->hwnd, &rc)) return;
    if (!PtInRect(&rc, pt)) return;
    SetCursor(LoadCursorW(NULL, nm));
}

static void vb6_uc_composeBelow(HWND hwnd, HDC memDC, int32_t cw, int32_t ch, int32_t ambient) {
    HWND sibs[64];
    int sibCount = 0;
    POINT self = { 0, 0 };

    vb6_uc_fillAmbient(memDC, cw, ch, ambient);
    if (!ClientToScreen(hwnd, &self)) return;

    /* 账 <transparency>: 真 VB6 的 UserControl 是 windowless, 控件的半透明填充
       (BackColorOpacity=60) 直接叠在**窗体背景**上 —— 背景图/Form.Picture 能透过来。
       C3 的 UC 是子窗口 + 离屏 DIB, DIB 以前只铺 ambient BackColor, 于是半透明只能
       和白色混 (Solid/Transparent 点了以后各控件仍是白板)。这里先把父窗体的
       AutoRedraw 持久表面 (= 窗体当前背景, 含 PaintPicture 画进去的 image.jpg)
       对应区域铺进 DIB, 再合成兄弟; 与 VB6 的累积绘制顺序一致。 */
    {
        HWND parent = GetParent(hwnd);
        HDC pdc = parent ? (HDC)GetPropW(parent, L"VB6_AutoRedrawDC") : NULL;
        if (pdc) {
            POINT pp = { 0, 0 };
            pp.x = self.x; pp.y = self.y;
            if (ScreenToClient(parent, &pp))
                BitBlt(memDC, 0, 0, cw, ch, pdc, pp.x, pp.y, SRCCOPY);
        }
    }

    /* GW_HWNDNEXT 走"本窗之下"的同层兄弟; 收集后**倒序**绘制 (最底下的先画)。*/
    for (HWND s = GetWindow(hwnd, GW_HWNDNEXT); s && sibCount < 64;
         s = GetWindow(s, GW_HWNDNEXT)) {
        sibs[sibCount++] = s;
    }
    for (int i = sibCount - 1; i >= 0; --i) {
        HWND s = sibs[i];
        POINT sp = { 0, 0 };
        RECT src;
        if (!IsWindowVisible(s)) continue;
        if (!ClientToScreen(s, &sp)) continue;
        GetClientRect(s, &src);
        int32_t dx = sp.x - self.x, dy = sp.y - self.y;
        int32_t sw = src.right - src.left, sh = src.bottom - src.top;
        if (sw <= 0 || sh <= 0) continue;
        if (dx >= cw || dy >= ch || dx + sw <= 0 || dy + sh <= 0) continue;
        int saved = SaveDC(memDC);
        if (!saved) continue;
        /* 兄弟的客户区 (0,0) 落到本 DIB 的 (dx,dy); 裁剪也只留它自己那块。*/
        SetViewportOrgEx(memDC, dx, dy, NULL);
        HRGN clip = CreateRectRgn(0, 0, sw, sh);
        if (clip) SelectClipRgn(memDC, clip);
        SendMessageW(s, WM_PRINTCLIENT, (WPARAM)memDC, PRF_CLIENT | PRF_ERASEBKGND);
        RestoreDC(memDC, saved);
        if (clip) DeleteObject(clip);
    }
}

// 把本 UC 画进 target。composeBelow=1: target 是本窗自己的窗口 DC (WM_PAINT),
// 透明 UC 需要先把下层合成进 DIB; composeBelow=0: 正被别的窗口画进它的 DIB
// (WM_PRINTCLIENT), 下层内容已经在目标 DC 里, 透明 UC 只画自己。
static void vb6_uc_paintInto(HWND hwnd, vb6_UCRec* r, HDC target,
                             int32_t cw, int32_t ch, int composeBelow) {
    if (getenv("C3_AD_TRACE")) {
        fprintf(stderr, "[AD-paint] type=%s cw=%d ch=%d paint=%p rdc=%p cb=%d\n",
                r && r->desc ? r->desc->typeName : "?", (int)cw, (int)ch,
                (void*)(r && r->desc ? (void*)r->desc->paint : NULL), (void*)(r ? (void*)r->hdc : NULL),
                (int)composeBelow);
        fflush(stderr);
    }
    if (!r || !r->desc || !r->desc->paint || cw <= 0 || ch <= 0) return;
    const int transparent = r->desc->transparent != 0;

    /* 账 #305 第二百九十二刀: AutoRedraw 表面已存在 ⇒ 画进它、再整块贴屏 (VB6 语义)。
       不重铺 ambient: 那张位图上留着上一次 Resize/Refresh 画的像素。*/
    HDC adc = (r->adOn && !transparent) ? vb6_uc_autoRedrawDC(r) : NULL;
    if (getenv("C3_AD_TRACE")) {
        fprintf(stderr, "[AD-use] type=%s adOn=%d adc=%p\n",
                r->desc->typeName, (int)r->adOn, (void*)adc);
        fflush(stderr);
    }

    vb6_UCDib dib;
    memset(&dib, 0, sizeof(dib));
    /* Fix B 2026-11-09: 「透明 UC 且 composeBelow==0」(正被别人 WM_PRINTCLIENT 合成) 这一路
       **不建离屏 DIB**。那张 DIB 是零初始化的纯黑, 而本控件在这一路只画自己那点内容
       (两个铺底 if 都不进), 末尾 `BitBlt(target, 0,0, dib.w,dib.h, dib.memDC, ...)` 就会把
       没画过的黑底整块盖回 target —— 实测 ucSlider1(4) 顶部 221x21 纯黑条 = 上层兄弟宿主
       (客户区 750..991 x 320..401) 与本控件 (750..981 x 380..471) 的交集, 正是兄弟合成时
       把这枚透明 UC 画进自己 DIB 的那一路。target 里已经是合成好的下层 (调用方也已按本窗
       矩形设好 viewport/clip), 直接画进去就是 VB6 windowless 的语义。 */
    const int directTarget = (transparent && !composeBelow);
    if (!adc && !directTarget) vb6_uc_dibCreate(&dib, target, cw, ch);
    if (!adc) vb6_uc_selectOwnFont(dib.memDC, hwnd);  /* 账 #305 300: 同上, 非 AutoRedraw 那一档 */
    const int useDib = (dib.memDC && dib.bits);
    HDC dc = adc ? adc : (useDib ? dib.memDC : target);
    /* VB6 的 UserControl.hDC 交出去时 BK 模式是 TRANSPARENT —— 控件里直接
       `DrawTextW UserControl.hDC, ...` 画标题时，参考产物(真 VB6)里字后面是
       控件自己的底色；C3 这儿沿用 GDI 默认 OPAQUE + 白 BKCOLOR，实测
       ucSegmentedControl 的 "Encendido" 等段名后面多出一个白方块。
       一个站点设一次，三条 DC 路径 (AutoRedraw / DIB / 目标) 都吃到。 */
    SetBkMode(dc, TRANSPARENT);

    /* 账 <font-color>: 真 VB6 里 `UserControl.ForeColor = v` 会落到 hDC 的文本色,
       控件里的 `DrawTextW UserControl.hDC, ...` 因此吃 ForeColor / SelectedForeColor。
       C3 以前只把它写进 vb6_UserControl_ForeColor 全局, DC 文本色永远是黑的 ——
       实测 uc12 选中页签标题 (SelectedForeColor=#CE00EF 品红) 与 uc9 "OPEN EDITOR"
       (ForeColor=vbHighlightText 白) 全被画成黑字 (问题 1)。这里每条 DC 路径都先
       按全局值设一次, 控件随后的赋值 (vb6_UC_SetForeColor) 再逐段覆盖。 */
    {
        int32_t fc = vb6_UserControl_ForeColor;
        COLORREF cr = vb6_ColorToColorRef(fc);
        SetTextColor(dc, cr);
    }
    /* 嵌套绘制 (透明 UC 的 composeBelow 会给兄弟发 WM_PRINTCLIENT) 会进到这里:
       保存外层 DC, 回来后还原, 否则外层控件随后的 ForeColor 赋值落不到自己的 DC。 */
    HDC prevPaintDC = vb6_uc_activePaintDC;
    vb6_uc_activePaintDC = dc;

    if (!adc) {
        if (transparent && composeBelow) {
            if (useDib) vb6_uc_composeBelow(hwnd, dc, cw, ch, r->ambientBack);
        } else if (!transparent) {
            vb6_uc_fillAmbient(dc, cw, ch, r->ambientBack);
        }
        /* 停写 AutoRedraw 之后那张位图当底图贴回去 (VB6 用 Picture=.Image 达到同一效果)。*/
        if (r->adSurf.memDC)
            BitBlt(dc, 0, 0, r->adSurf.w, r->adSurf.h, r->adSurf.memDC, 0, 0, SRCCOPY);
    }
    /* 透明 + 非合成 (被别人 WM_PRINTCLIENT): 一个像素都不填 —— 目标 DC 里已经是
       合成好的下层, 再填 ambient 会把兄弟内容盖掉。*/

    r->hdc = dc;
    vb6_UCSaved saved;
    vb6_uc_push(r, &saved);
    vb6_uc_trace("paint.begin", r->desc->typeName, r->me);
    r->desc->paint(r->me);
    vb6_uc_trace("paint.end", r->desc->typeName, r->me);
    vb6_uc_pop(&saved);

    int bltOk = 0;
    if (adc) {
        bltOk = (int)BitBlt(target, 0, 0, cw, ch, adc, 0, 0, SRCCOPY);
    } else if (useDib) {
        bltOk = (int)BitBlt(target, 0, 0, dib.w, dib.h, dib.memDC, 0, 0, SRCCOPY);
        // Fix 113h/123: 调试捕获 (C3_UC_DUMPDIR) —— 只在真实 WM_PAINT 路径做,
        // WM_PRINTCLIENT 是被兄弟合成调用的, 重复 dump 会淹没序号。
        const char* dumpDir113h = composeBelow ? getenv("C3_UC_DUMPDIR") : NULL;
        if (dumpDir113h && *dumpDir113h) {
            char path113h[1024];
            _snprintf(path113h, sizeof(path113h), "%s\\%s_%d_%p.bmp", dumpDir113h,
                      r->desc->typeName, (int)++vb6_ucDumpSeq, hwnd);
            path113h[sizeof(path113h) - 1] = '\0';
            vb6_uc_dibSaveBmp(&dib, path113h);
            // Fix 123: 同步输出整窗合成图 (含 z 序/相对位置)
            vb6_uc_dumpFormComposite((HWND)GetAncestor(hwnd, GA_ROOT), dumpDir113h);
        }
        vb6_uc_dibDestroy(&dib);
    }
    /* 诊断刀 (C3_AD_TRACE): 把「DIB -> 窗口表面」这一步的每个可失败点摊开 ——
       blt=0 表示 BitBlt 本身失败 (target 为 NULL / 无效); clip= 是 BeginPaint 给的
       裁剪框 (全 0 => 更新区为空, 画了也被裁掉); px= 是 BitBlt 之后从 target 回读的
       像素 (COLORREF 序 = 0x00BBGGRR, 0xFFFFFF = 纯白 = 只有 ambient)。 */
    if (getenv("C3_AD_TRACE")) {
        RECT cbox = { -1, -1, -1, -1 };
        COLORREF px = 0xFFFFFFFF;
        if (target) {
            GetClipBox(target, &cbox);
            px = GetPixel(target, 1, 1);
        }
        fprintf(stderr, "[AD-blit] type=%s blt=%d target=%p clip=%d,%d,%d,%d px=%06X adc=%p dib=%p\n",
                r && r->desc ? r->desc->typeName : "?", bltOk, (void*)target,
                (int)cbox.left, (int)cbox.top, (int)cbox.right, (int)cbox.bottom,
                (unsigned)px, (void*)adc, (void*)(useDib ? dib.memDC : NULL));
        fflush(stderr);
    }
    r->hdc = NULL;
    vb6_uc_activePaintDC = prevPaintDC;
}

static LRESULT CALLBACK vb6_uc_wndproc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    vb6_UCRec* r = (vb6_UCRec*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    switch (msg) {
        case WM_NCCREATE: {
            CREATESTRUCTW* cs = (CREATESTRUCTW*)lParam;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)cs->lpCreateParams);
            break;
        }
        case WM_SIZE: {
            if (!r || !r->ready) break;
            r->scaleWidth = (int32_t)LOWORD(lParam);
            r->scaleHeight = (int32_t)HIWORD(lParam);
            if (r->desc && r->desc->resize) {
                vb6_UCSaved saved;
                vb6_uc_push(r, &saved);
                vb6_uc_trace("resize.begin", r->desc->typeName, r->me);
                r->desc->resize(r->me);
                vb6_uc_trace("resize.end", r->desc->typeName, r->me);
                vb6_uc_pop(&saved);
            }
            /* Fix <vbeclipse> rev22: 本 UC 的尺寸变了 ⇒ 它每个**设计期子控件**的
             * 尺寸也快变了 ⇒ 该跑它们的 `<Ctrl>_Resize` 事件 (VB6 语义)。
             *
             * 为什么挂在这里而不是 vb6_ControlMove (那是所有摆位的收口):
             *   放在 Move 里会**递归**。`ViewArea.Move …` 正是 ucFolder 的
             *   `UserControl_Resize` 内部发的, 而那个过程就是经
             *   WM_SIZE → vb6_uc_push → resize → Move 这条链进来的 ⇒ Move 里再跑
             *   子控件事件 ⇒ 又 Move ⇒ 无限递归。
             *   实测 play78 (rev22 首次实现): 视图窗体尺寸确实对上了容器
             *   (frmViewHelp 253x219 / frmViewSnapshot 778x190), 但**整个窗口
             *   上下颠倒 + 文字镜像** —— 递归 Move 把各层控件反复推挤, 坐标系翻转。
             *   两害相权: 宁可尺寸停在中间值 (rev21 状态) 也不能递归。
             *
             * 为什么放在 pop **之后** 是安全的: 此时 vb6_ucCurrent 已还原成外层值
             * (通常 NULL), 所以 vb6_UC_RunDesignResize 里 "上下文内就不跑" 的
             * 判据不会误挡本调用 —— 它是唯一允许在上下文外触发的入口。
             * 且它内部有同控件重入短路, 事件体里 Move 别的控件再发 WM_SIZE 也不会
             * 无限展开 (那一层 vb6_ucCurrent 非空, 直接被挡)。*/
            /* Fix <vbeclipse> rev23: 这里**只排队**, 不直接跑 —— WM_SIZE 是
             * SetWindowPos/MoveWindow 的**同步** SendMessage, 一整串嵌套 Move 全在
             * 同一个调用栈里跑完才返回, 那时 ucFolder 的 ViewArea 还没拿到最终尺寸
             * ⇒ `ViewArea_Resize` 按设计期宽度去 Move 视图窗体 (实测恒 W=8505)。
             * 排到消息循环里跑, 各控件尺寸才是最终值。详见 vb6_UC_QueueDesignResize。*/
            vb6_UC_QueueDesignResize(r->me);
            InvalidateRect(hwnd, NULL, FALSE);
            return 0;
        }
        /* Fix <vbeclipse> rev23: 排队的子控件 Resize 事件在这里接住。
         * 必须排在消息循环里 (不能直接在 WM_SIZE 里跑) —— WM_SIZE 是
         * SetWindowPos/MoveWindow 的**同步** SendMessage, 一整串嵌套 Move 全在同
         * 一个调用栈里跑完才返回, 那时 ucFolder 的 ViewArea 还没拿到最终尺寸。*/
        case VB6_UC_DR_MSG:
            vb6_UC_DrainDesignResize(hwnd);
            return 0;
        case WM_SHOWWINDOW:
            if (r && r->ready && wParam && r->desc && r->desc->show) {
                vb6_UCSaved saved;
                vb6_uc_push(r, &saved);
                vb6_uc_trace("show.begin", r->desc->typeName, r->me);
                r->desc->show(r->me);
                vb6_uc_trace("show.end", r->desc->typeName, r->me);
                vb6_uc_pop(&saved);
            }
            return 0;
        case WM_SETCURSOR:
            /* 缺陷 1: .ctl 在 MouseMove 里设 UserControl.MousePointer (Disabled 项
               = vbNoDrop/IDC_NO)。此前那颗全局没人读 ⇒ 光标永远是箭头。这里按本实例
               记录回答: 命中客户区时用映射出的系统光标, 其余交回 DefWindowProc。 */
            if (r && r->ready && (int32_t)(lParam & 0xFFFF) == 1 /* HTCLIENT */) {
                LPCWSTR nm = vb6_uc_cursorName(r->mousePointer);
                if (nm) { SetCursor(LoadCursorW(NULL, nm)); return TRUE; }
            }
            break;
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_RBUTTONUP:
        case WM_MOUSEMOVE: {
            // czUI fix: 把宿主窗口收到的鼠标消息转成 UserControl_MouseDown/Up/Move
            // (此前完全没有转发, 控件无法点击/拖动 — czUI 标题栏拖动、按钮、开关全死)。
            if (!r || !r->ready || !r->desc || !r->me) break;
            void (*hook)(void*, int32_t, int32_t, float, float) = NULL;
            int32_t button = 0;
            switch (msg) {
                case WM_LBUTTONDOWN: hook = r->desc->mouseDown;  button = 1; break;
                case WM_RBUTTONDOWN: hook = r->desc->mouseDown;  button = 2; break;
                case WM_LBUTTONUP:   hook = r->desc->mouseUp;    button = 1; break;
                case WM_RBUTTONUP:   hook = r->desc->mouseUp;    button = 2; break;
                /* 账 #609: MouseMove 的 Button 形参在 VB6 里是"当前按住的键状态"
                 * (1=左 2=右 4=中), 不是"哪个键触发了本条消息"。这里原来写死 0,
                 * 于是 `If Button = vbLeftButton Then` 形态的拖动代码 (ucSlider 的
                 * UserControl_MouseMove) 永远进不去 —— 按在轨道上能跳值 (MouseDown
                 * 里直接调了一次 MouseMove, Button=1), 按住旋钮拖却纹丝不动
                 * (实测: 真 VB6 拖 150px 旋钮上移 187px, 旧 C3 产物 0px)。
                 * wParam 的 MK_* 就是这个状态, 直接映射。*/
                case WM_MOUSEMOVE: {
                    hook = r->desc->mouseMove;
                    if (wParam & MK_LBUTTON) button |= 1;
                    if (wParam & MK_RBUTTON) button |= 2;
                    if (wParam & MK_MBUTTON) button |= 4;
                    break;
                }
            }
            if (!hook) break;
            // 捕获鼠标, 保证按下后拖出窗口仍能收到 UP/MOVE (VB6 隐式行为)
            if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN) SetCapture(hwnd);
            else if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP) ReleaseCapture();
            float sx = (float)(short)LOWORD(lParam);
            float sy = (float)(short)HIWORD(lParam);
            // 坐标换算到控件当前 ScaleMode 单位 (账 #175: 与 ScaleWidth/Move 同一个权威)
            int32_t ucSm175 = r->desc->scaleMode;
            sx = (float)vb6_ScalePxToUser((double)sx, ucSm175, 0);
            sy = (float)vb6_ScalePxToUser((double)sy, ucSm175, 1);
            vb6_UCSaved saved;
            vb6_uc_push(r, &saved);
            // czUI fix: 回调只更新状态; 视觉刷新统一走 WM_PAINT 双缓冲
            // (直接 GetDC 画屏幕与擦除/重绘交错会造成闪烁)
            vb6_UserControl_hDC = NULL;
            hook(r->me, button, 0, sx, sy);
            vb6_uc_pop(&saved);
            // 账 #226: VB6 的 Click 是"按下并抬起"之后发的, 落点就在这条 WM_LBUTTONUP;
            // 必须在 MouseUp 转调**之后**, 顺序才与 VB6 一致。旧 cgen 不发这一槽 ⇒
            // 初始化式补 0 ⇒ 这里不转调, 行为与改动前逐字节一致。
            if (msg == WM_LBUTTONUP && r->desc->click) r->desc->click(r->me);
            InvalidateRect(hwnd, NULL, FALSE);
            UpdateWindow(hwnd);
            return 0;
        }
        case WM_LBUTTONDBLCLK: {
            // 账 #227: 这一槽 cgen 一直在填 (UC 里写 UserControl_DblClick 才有真身, 否则是空
            // stub), 但宿主从没有转调过它 —— 语料里六枚 UC 的 `RaiseEvent DblClick` 于是永远
            // 发不出去。类样式上 CS_DBLCLKS 早就立了 (下面那句 czUI fix), 消息收得到,
            // 缺的就是这一跳。
            // MouseDown/MouseUp 的转调**不**挂在这条消息上: 物理双击 Windows 发的是
            // DOWN/UP/DBLCLK/UP, Click 那一路已经由两条 UP 供过, 再挂就变三发。
            if (!r || !r->ready || !r->desc || !r->me || !r->desc->dblClick) break;
            vb6_UCSaved saved227;
            vb6_uc_push(r, &saved227);
            vb6_UserControl_hDC = NULL;
            r->desc->dblClick(r->me);
            vb6_uc_pop(&saved227);
            InvalidateRect(hwnd, NULL, FALSE);
            UpdateWindow(hwnd);
            return 0;
        }

        case WM_CAPTURECHANGED:
            break;
        /* 账 #305 第三百七刀: 宿主窗口里那批**属于容器窗体**的设计期子控件, 通知发给
           直接父窗口 = 本 UC 宿主。ComCtlsDemo 的 MainForm 就把 `Command4`("Pager Demo")
           authored 在 `ComCtlsDemo.CoolBar CoolBar1` 的 Begin 块里 —— 它的事件处理器
           (Command4_Click) 住在**窗体模块**, 而 BN_CLICKED 只送到宿主窗口; 以前这条
           走 DefWindowProc 被静默丢掉 (实测: 鼠标点与程序发 WM_COMMAND 都到过宿主,
           窗体那张 id 派发表从没看见 201)。
           VB6 的 UserControl 宿主对"不属于自己"的子控件通知就是原样上抛, 这里照做。
           不会误伤 UC 自己的设计期子控件: 那些由 CreateDesign* 建, 控件 id 恒 0
           (hMenu 传 NULL), 窗体派发表里没有 0 这一格。 */
        case WM_COMMAND:
        case WM_NOTIFY: {
            HWND parent = GetParent(hwnd);
            if (parent) return SendMessageW(parent, msg, wParam, lParam);
            break;
        }
        case WM_PAINT: {
            if (!r || !r->ready) {
                if (getenv("C3_AD_TRACE")) {
                    fprintf(stderr, "[AD-pm] type=%s SKIP not-ready\n",
                            r && r->desc ? r->desc->typeName : "?");
                    fflush(stderr);
                }
                break;
            }
            if (!r->desc || !r->desc->paint) {
                if (getenv("C3_AD_TRACE")) {
                    fprintf(stderr, "[AD-pm] type=%s SKIP no-paint\n",
                            r->desc ? r->desc->typeName : "?");
                    fflush(stderr);
                }
                break;
            }
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            RECT crc;
            GetClientRect(hwnd, &crc);
            /* 诊断刀: BeginPaint 的返回值 (NULL = 无更新区/窗口不可见 => 画了也白搭)、
               更新区矩形 (全 0 = 空更新区)、父窗口与可见性, 一次摊开。 */
            if (getenv("C3_AD_TRACE")) {
                fprintf(stderr, "[AD-pm] type=%s hdc=%p parent=%p rc=%d,%d,%d,%d vis=%d\n",
                        r->desc->typeName, (void*)hdc, (void*)GetParent(hwnd),
                        (int)ps.rcPaint.left, (int)ps.rcPaint.top,
                        (int)ps.rcPaint.right, (int)ps.rcPaint.bottom,
                        (int)IsWindowVisible(hwnd));
                fflush(stderr);
            }
            vb6_UC_PaintEnter();
            vb6_uc_paintInto(hwnd, r, hdc,
                             (int32_t)(crc.right - crc.left),
                             (int32_t)(crc.bottom - crc.top), 1);
            vb6_UC_PaintLeave();
            EndPaint(hwnd, &ps);
            return 0;
        }
        /* 账 #492: 被父窗口/兄弟窗口的 PRF_CHILDREN 合成时走这里。目标 DC 里
           (对透明 UC 来说) 已经铺好了下面那一层, 这里只把本控件自己画上去。*/
        case WM_PRINTCLIENT: {
            if (!r || !r->ready || !r->desc || !r->desc->paint) break;
            RECT crc;
            GetClientRect(hwnd, &crc);
            vb6_UC_PaintEnter();
            vb6_uc_paintInto(hwnd, r, (HDC)wParam,
                             (int32_t)(crc.right - crc.left),
                             (int32_t)(crc.bottom - crc.top), 0);
            vb6_UC_PaintLeave();
            return 0;
        }
        case WM_ERASEBKGND:
            /* 诊断刀: 擦底色这一步会把整块客户区刷成 ambient (白) —— 判据抓到的
               「宿主矩形纯白」就可能是它干的 (真因是它之后的那次 paint 没落上)。 */
            if (getenv("C3_AD_TRACE")) {
                fprintf(stderr, "[AD-erase] type=%s ready=%d ambient=%08X\n",
                        r && r->desc ? r->desc->typeName : "?", r ? (int)r->ready : 0,
                        r ? (unsigned)r->ambientBack : 0u);
                fflush(stderr);
            }
            if (r && r->ready) {
                /* 账 #492: 透明 UC 不擦背景 —— 一擦就是一块 ambient 灰, WM_PAINT
                   里的合成会覆盖它, 中间那一帧在屏幕上就是灰闪。*/
                if (!r->desc || !r->desc->transparent) {
                    // 用容器底色填充, 避免重绘闪烁 (固定白色会在深色 UI 上闪白)
                    RECT rc; GetClientRect(hwnd, &rc);
                    COLORREF cref = vb6_ColorToColorRef(r->ambientBack);
                    HBRUSH b = CreateSolidBrush(cref);
                    FillRect((HDC)wParam, &rc, b);
                    DeleteObject(b);
                }
            }
            return 1;
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC: {
            // czUI fix: 设计器子控件 (txtEmbed 等) 的 BackColor/ForeColor 落在
            // VB6_BackColor/ForeColor 窗口属性上 (vb6_SetControlBackColor), 这里
            // 应用之 — 否则深色 UI 上子编辑框是默认白底黑字。
            HWND child = (HWND)lParam;
            if (!child) break;
            HDC hdc = (HDC)wParam;
            COLORREF fg = vb6_ColorToColorRef(vb6_GetControlForeColor((void*)child));
            COLORREF bg = vb6_ColorToColorRef(vb6_GetControlBackColor((void*)child));
            SetTextColor(hdc, fg);
            SetBkColor(hdc, bg);
            // 刷子缓存进窗口属性, 避免每条消息泄漏 GDI 句柄
            HBRUSH br = (HBRUSH)GetPropW(child, L"VB6_BgBrush");
            if (!br) {
                br = CreateSolidBrush(bg);
                SetPropW(child, L"VB6_BgBrush", (HANDLE)br);
            }
            return (LRESULT)br;
        }
        /* 账 #305 第二百七十九刀: 参考那台 29 枚 UC 宿主每一枚都**本人答**这一问 (-13/MS Sans Serif),
         * C3 这一户从前答空 —— 不是没发字体 (第二百七十八刀就发了), 是 DefWindowProc 不记 WM_SETFONT
         * (探针 .local/k278/wmsetfont.c 两架构实测: 发完再问回 0) ⇒ 那一臂 44 枚一字未动。
         * 答案只从我们自己那一格拿 (所有权口径见 vb6forms_ctrl.c 里那位只读主人上面那段),
         * 所以这里刻意**不**接 WM_SETFONT。 */
        case WM_GETFONT:
            return (LRESULT)(INT_PTR)vb6_ControlFontStored(hwnd);
        case WM_DESTROY:
            // Fix <vbeclipse> rev10 (C3_UC_TRACE): 临时 —— 定位
            // "ShowPerspective 报 No perspective id found" 是否因为本控件的
            // WM_DESTROY 提前跑过 UserControl_Terminate (源码里它会
            // `Set m_Perspectives = Nothing`)。
            if (getenv("C3_UC_TRACE")) {
                fprintf(stderr, "[UC] WM_DESTROY hwnd=%p me=%p hasTerminate=%d\n",
                        (void*)hwnd, r ? r->me : NULL,
                        (r && r->desc && r->desc->terminate) ? 1 : 0);
                fflush(stderr);
            }
            if (r && r->desc && r->desc->terminate) {
                vb6_UCSaved saved;
                vb6_uc_push(r, &saved);
                r->desc->terminate(r->me);
                vb6_uc_pop(&saved);
            }
            // §51: 宿主销毁后替身记录必须摘掉 —— 留着会被 vb6_uc_findByInstance
            // 在下次打开时先扫到 (堆地址复用), 新实例的画全落到这张死表面上。
            if (r) vb6_uc_releaseHost((void*)hwnd);
            break;
        default:
            break;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// Fix 112b: GDI+ 进程级初始化. 图表控件的绘制走 GdipCreateFromHDC, 而
// Charts 2020 自带的 ManageGDIToken 在打完补丁后会 GdiplusShutdown 掉它自己
// 启动的 token —— RTL 这里保一个常驻 token, 保证绘制始终可用.
// gdiplus.h 是 C++ 专用头, 这里按桩文件的做法用 LoadLibrary + GetProcAddress.
typedef struct vb6_GdiplusStartupInput {
    uint32_t GdiplusVersion;
    void*    DebugEventCallback;
    int32_t  SuppressBackgroundThread;
    int32_t  SuppressExternalCodecs;
} vb6_GdiplusStartupInput;

void vb6_uc_gdiplusInit(void) {
    static int done = 0;
    if (done) return;
    done = 1;
    // czUI fix: 环境字体默认名 — Bag 重放 ReadProperty("Font", Ambient.Font)
    // 会把此对象设为控件字体; Name=NULL 时所有 GDI+ 文字静默消失
    // 第二百七十七刀: 这一档只要求**非空**, 族名交 VB6 自己的默认字族 (旧写法 Segoe UI) ——
    // 同一把 WM_GETFONT 尺在参考那台读回来的 69 枚就是 MS Sans Serif。
    if (!vb6_UserControl_FontObj.Name)
        vb6_UserControl_FontObj.Name = SysAllocString(L"MS Sans Serif");
    HMODULE mod = LoadLibraryA("gdiplus.dll");
    if (!mod) return;
    long (__stdcall *pStartup)(ULONG_PTR*, const void*, void*) =
        (long (__stdcall *)(ULONG_PTR*, const void*, void*))GetProcAddress(mod, "GdiplusStartup");
    if (!pStartup) return;
    vb6_GdiplusStartupInput si;
    memset(&si, 0, sizeof(si));
    si.GdiplusVersion = 1;
    ULONG_PTR token = 0;
    pStartup(&token, &si, NULL);   /* 常驻 token, 不配对 Shutdown */
}

void vb6_uc_registerClass(HINSTANCE hInst) {
    static int done = 0;
    if (done) return;
    WNDCLASSW wc;
    memset(&wc, 0, sizeof(wc));
    wc.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;  // czUI fix: UserControl_DblClick 需要
    wc.lpfnWndProc = vb6_uc_wndproc;
    wc.hInstance = hInst;
    wc.hCursor = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = NULL;
    wc.lpszClassName = L"VB6_UserControlHost";
    RegisterClassW(&wc);
    done = 1;
}

#ifdef __cplusplus
} // extern "C"
#endif
