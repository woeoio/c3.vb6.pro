// vb6forms_uc 拆分片段 uc_hostmodel_setprop (函数体)
// 内容 = 拆分前 vb6forms_uc.c 第 1392~1419 行，逐行未改
// 在函数体内 #include；路径一律用 basename

    obj = vb6_UC_UnwrapHost(obj);  // czUI fix: 包装器透明解包
    const vb6_VARIANT* v = (const vb6_VARIANT*)inV;
    if (!obj || !name) return 0;
    if (vb6_uc_isFont(obj)) {
        /* 账 #305 第二百七十六刀: 与读侧同一口径 —— 认得就写、认不出先报数再答
           (「写入被忽略」这件事本身与改前一字不差, 只是从此响亮)。
           `Font.hFont = x` 到不了这里: 它在 vb6com_invoke.c 那条 kind 5 臂就被按
           **只读**处理了 (stdole2.tlb 里 hFont 只有 propget, VB6 里那句是编译错误)。 */
        if (vb6_ho_putFontMember(obj, name, v)) return 1;
        vb6_ho_reportUnknown(L"SetProp", name, obj);
        return 1;
    }
    if (vb6_uc_isControls(obj)) return 1;
    if (vb6_uc_isColl(obj)) return 1;   // Fix 112c: Collection 只读容器

    // Fix <vbeclipse> rev18: UC 实例指针 → 宿主 HWND (与 getprop / call 同口径)
    { void* h18 = vb6_UC_HwndOf(obj); if (h18) obj = h18; }

    if (!vb6_ho_find(obj) && !vb6_uc_findByHwnd(obj) && !IsWindow((HWND)obj)) return 0;

    if (_wcsicmp(name, L"Caption") == 0 || _wcsicmp(name, L"Text") == 0) {
        BSTR s = vb6_ho_variantToBstr(v);
        if (s) SetWindowTextW((HWND)obj, s);
        InvalidateRect((HWND)obj, NULL, TRUE);
        return 1;
    }
    /* 账 #305 第三百五十一刀: Tag —— 与 cgen 的直接路径共用**同一份存储**
       (vb6forms_style.c 的 VB6_Tag 窗口属性)。以前晚绑定走宿主模型 ⇒ 落到
       reportUnknown, 而直接路径 `控件.Tag` 早就写得进; 两条路读同一枚控件得到
       两个答案。这里只把那条存储接进来, 不另立第二份表。 */
    if (_wcsicmp(name, L"Tag") == 0) {
        vb6_SetControlTag(obj, (void*)vb6_ho_variantToBstr(v));
        return 1;
    }
    if (_wcsicmp(name, L"Visible") == 0) { ShowWindow((HWND)obj, vb6_ho_variantToLong(v) ? SW_SHOW : SW_HIDE); return 1; }
    if (_wcsicmp(name, L"Enabled") == 0) { EnableWindow((HWND)obj, vb6_ho_variantToLong(v) ? TRUE : FALSE); return 1; }
    // 账 #247: 与读侧同一个出口 —— 四档各写自己那一档 (单位 = 所在容器的 ScaleMode,
    // 并把写进去的数存进 VB 侧读数缓存)。以前这一支把四档一起读成缇、再整体 MoveWindow:
    // 写 .Left 会把 Top/Width/Height 也按缇重算一遍推回窗口, 而且谁都不存, 于是同一句
    // `对象.Width = 7222` 走这里读回 7215, 走发码读回 7222。
    if (_wcsicmp(name, L"Left") == 0)   { vb6_SetControlLeft(obj, vb6_ho_variantToLong(v)); return 1; }
    if (_wcsicmp(name, L"Top") == 0)    { vb6_SetControlTop(obj, vb6_ho_variantToLong(v)); return 1; }
    if (_wcsicmp(name, L"Width") == 0)  { vb6_SetControlWidth(obj, vb6_ho_variantToLong(v)); return 1; }
    if (_wcsicmp(name, L"Height") == 0) { vb6_SetControlHeight(obj, vb6_ho_variantToLong(v)); return 1; }
    /* 账 #305 第三百五十一刀: RichTextBox.RightMargin (VB6 侧缇)。
       RichEdit 的 EM_SETMARGINS 以缇计右边距; 标准 Edit 没有 VB6 这一格,
       所以只认 RICHEDIT 类。同值重写也返回 1 (与样式位那批同一口径)。 */
    if (_wcsicmp(name, L"RightMargin") == 0) {
        wchar_t rcls[64] = {0};
        GetClassNameW((HWND)obj, rcls, 63);
        if (_wcsnicmp(rcls, L"RICHEDIT", 8) == 0) {
            int32_t tw = vb6_ho_variantToLong(v);
            if (tw < 0) tw = 0;
            SendMessageW((HWND)obj, EM_SETMARGINS, EC_RIGHTMARGIN,
                         (LPARAM)(((LPARAM)tw & 0xFFFF) << 16));
            return 1;
        }
    }
    // Fix <vbeclipse> rev18: 内建成员都没命中 → 写**本 UserControl 的自有 Property**
    // (ucFolder.LeftPos = … 这类)。注意必须放在内建档之后: `Visible`/`Caption` 等同名
    // 内建属性仍是窗口层语义。
#ifndef TVS_HASLINES
#define TVS_HASLINES    0x0002
#endif
#ifndef TVS_LINESATROOT
#define TVS_LINESATROOT 0x0004
#endif
#ifndef TVS_EDITLABELS
#define TVS_EDITLABELS  0x0008
#endif
    /* 账 #305 第三百五十一刀: 原生 **TreeView** 的外观属性写。
       ShowLines / LineStyle / LabelEdit / ShowBorder -> GWL_STYLE 真样式位;
       Appearance -> GWL_EXSTYLE 的 WS_EX_CLIENTEDGE (3D 凹边);
       CheckBoxes/HotTracking/HideSelection/Indentation 转 RTL 的既有 setter,
       不再"收下不响"。Icons 暂缓 —— 需要把 ImageList 对象可靠映射成 HIMAGELIST
       (OCX 内部状态), 本刀不做假动作, 让它继续落 reportUnknown 保持可见。 */
    {
        wchar_t cls350[64] = {0};
        GetClassNameW((HWND)obj, cls350, 63);
        if (_wcsicmp(cls350, L"SysTreeView32") == 0) {
            int32_t lv = vb6_ho_variantToLong(v);
            LONG st = GetWindowLongW((HWND)obj, GWL_STYLE);
            LONG ex = GetWindowLongW((HWND)obj, GWL_EXSTYLE);
            LONG nv = st, nex = ex;
            int handled = 0;
            if (_wcsicmp(name, L"ShowLines") == 0) {
                nv = lv ? (st | TVS_HASLINES) : (st & ~TVS_HASLINES);
                handled = 1;
            } else if (_wcsicmp(name, L"LineStyle") == 0) {
                /* 与 vb6_TreeView_SetLineStyle 同一口径: 两种取值都立 HASLINES,
                   LINESATROOT 只在非 0 时立住 (以前这里写反了)。 */
                nv = lv ? (st | TVS_HASLINES | TVS_LINESATROOT)
                        : ((st | TVS_HASLINES) & ~TVS_LINESATROOT);
                handled = 1;
            } else if (_wcsicmp(name, L"LabelEdit") == 0) {
                nv = (lv == 0) ? (st | TVS_EDITLABELS) : (st & ~TVS_EDITLABELS);
                handled = 1;
            } else if (_wcsicmp(name, L"ShowBorder") == 0) {
                nv = lv ? (st | WS_BORDER) : (st & ~WS_BORDER);
                handled = 1;
            } else if (_wcsicmp(name, L"Appearance") == 0) {
                nex = lv ? (ex | WS_EX_CLIENTEDGE) : (ex & ~WS_EX_CLIENTEDGE);
                handled = 1;
            } else if (_wcsicmp(name, L"CheckBoxes") == 0) {
                vb6_TreeView_SetCheckBoxes(obj, lv); handled = 1;
            } else if (_wcsicmp(name, L"HotTracking") == 0) {
                vb6_TreeView_SetHotTracking(obj, lv); handled = 1;
            } else if (_wcsicmp(name, L"HideSelection") == 0) {
                vb6_TreeView_SetHideSelection(obj, lv); handled = 1;
            } else if (_wcsicmp(name, L"Indentation") == 0) {
                vb6_TreeView_SetIndentation(obj, lv); handled = 1;
            }
            if (handled) {
                /* 同值重写也必须认账: 以前只在样式位有变化时 return 1,
                   `ShowLines` 值已等于当前状态就掉进 reportUnknown 误报 miss。 */
                if (nv != st) SetWindowLongW((HWND)obj, GWL_STYLE, nv);
                if (nex != ex) SetWindowLongW((HWND)obj, GWL_EXSTYLE, nex);
                if (nv != st || nex != ex) {
                    SetWindowPos((HWND)obj, NULL, 0, 0, 0, 0,
                                 SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER
                                 | SWP_FRAMECHANGED);
                    InvalidateRect((HWND)obj, NULL, TRUE);
                }
                return 1;
            }
            /* 树见过但本刀还没实现的名字, 先保持原样收下; Icons 刻意不在其中。 */
            static const wchar_t* const kTvKnown[] = {
                L"Sorted", L"SingleSel", L"FullRowSelect", L"Scroll", NULL };
            for (const wchar_t* const* k = kTvKnown; *k; ++k)
                if (_wcsicmp(name, *k) == 0) return 1;
        }
    }
    if (vb6_UC_OwnPropSet(obj, name, v)) return 1;

    /* 账 #305 第二百六十二刀 (丙): 同 Call 那条 —— 不认识的属性名以前静默忽略。 */
    vb6_ho_reportUnknown(L"SetProp", name, obj);
    return 1;   // 宿主对象: 未知属性写入一律忽略
