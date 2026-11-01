// vb6forms_uc 拆分片段 uc_hostmodel_getprop (函数体)
// 内容 = 拆分前 vb6forms_uc.c 第 1304~1387 行，逐行未改
// 在函数体内 #include；路径一律用 basename

    obj = vb6_UC_UnwrapHost(obj);  // czUI fix: 包装器透明解包
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    if (!obj || !name) return 0;
    vb6_ho_setVariantEmpty(out);

    if (vb6_uc_isControls(obj)) {
        if (_wcsicmp(name, L"Count") == 0) { vb6_ho_setVariantLong(out, vb6_UC_ControlsCount(obj)); return 1; }
        return 1;  // 其它成员当作无值, 避免 C 崩溃
    }
    if (vb6_uc_isColl(obj)) {   // Fix 112c: RTL Collection
        if (_wcsicmp(name, L"Count") == 0) { vb6_ho_setVariantLong(out, vb6_Collection_Count(obj)); return 1; }
        return 1;
    }
    if (vb6_uc_isFont(obj)) {
        /* 账 #305 第二百七十六刀: 认得就答、认不出**先报数** —— 第二百七十三刀给方法
           那一臂装的那位主人, 属性这一臂当时没跟着接 (第二百七十五刀点出的那格洞:
           靶工程 34 发 `name=hFont` 在 C3_HOST_MEMBER_MISS 名册里一条也没有)。
           行为一字不改: 认不出仍是「答了但没答」—— 出口那格早被上面的
           vb6_ho_setVariantEmpty 清零, 调用方拿到的还是同一个空值。 */
        if (vb6_ho_getFontMember(obj, name, out)) return 1;
        vb6_ho_reportUnknown(L"GetProp", name, obj);
        return 1;
    }

    // Fix <vbeclipse> rev18: UC **实例指针**归一成宿主 HWND —— 晚绑定属性读
    // (`Dim x As Variant: Set x = Controls.Item(id)` 之后 `x.LeftPos`) 传进来的是
    // `vb6_cls_X*`, 而下面整条内建链只认 HWND, 以前一律落到末尾返回 Empty。
    { void* h18 = vb6_UC_HwndOf(obj); if (h18) obj = h18; }

    vb6_UCRec* r = vb6_uc_findByHwnd(obj);
    vb6_HostObjRec* h = vb6_ho_find(obj);
    /* 账 #305 第二百八十二刀: 无窗宿主 (第二百五十八刀那户 InvisibleAtRuntime) 递进来的
       接收者**就是实例指针**, 上面那两把按 hwnd 的钥匙都开不了它的门 ⇒ 以前在这一行
       `return 0` 早退, 连 :189 那位「自有 Property」出口都走不到 (实测靶工程
       `ControlEnum.hImageList` 一条都不发 —— ImageList1/2/3 全挂在无窗那一档)。
       这一档只补「认得接收者是谁」, 已有的取值一律不动: 窗口类的伪成员 (Visible/Enabled/
       Container/Parent 那些按 HWND 问的) 对无效句柄只会答 0/NULL, 与改前「答 Empty」同形。 */
    if (!r && !h && !IsWindow((HWND)obj)) {
        h = vb6_ho_findInstance(obj);
        if (h) r = vb6_uc_findByInstance(h->instance);
    }
    if (!r && !h && !IsWindow((HWND)obj)) return 0;

    int32_t scaleMode = r ? (r->desc ? r->desc->scaleMode : 1) : 1;

    if (_wcsicmp(name, L"hwnd") == 0 || _wcsicmp(name, L"hWnd") == 0) {
        vb6_ho_setVariantLong(out, (int32_t)(intptr_t)obj); return 1;
    }
    /* 账 #271: `hWndUserControl` 是 Krool/ComCtls 那一族 UserControl 的**通行属性**
       (Animation/CheckBoxW/ComboBoxW/CommandButtonW … 每个 .ctl 都写
       `Public Property Get hWndUserControl() As LongPtr: hWndUserControl = UserControl.hWnd`)。
       CoolBar.ctl:2028 用「先问它、问不到(Err)再问 .hWnd」来区分 UC 与内建控件:
         Handle = GetWindowFromObject(Value.Object)
         If Handle = NULL_PTR Then
             On Error Resume Next: Handle = Value.hWndUserControl
             If Err.Number <> 0 Then Handle = Value.hWnd
         …  If ControlIsValid(Value) = True And Handle <> NULL_PTR Then … Else Err.Raise 380
       本 RTL 的宿主模型对**认不得的成员一律交 Empty 且不报错** (刻意如此, 见文件末尾),
       于是 `Value.hWndUserControl` 不报错 ⇒ `Err.Number = 0` ⇒ 不会退到 `.hWnd`
       ⇒ Handle 停在 0 ⇒ 第二十六层的 `Err.Raise 380`。
       VB6 里 `UserControl.hWnd` 与本模型的 `hWnd` 是**同一个窗口**, 这里按同一口径
       交回 obj 即可 (对 UC 与非 UC 控件都一样 ⇒ 与 VB6 那条 fallback 殊途同归)。 */
    if (_wcsicmp(name, L"hWndUserControl") == 0) {
        vb6_ho_setVariantLong(out, (int32_t)(intptr_t)obj); return 1;
    }
    if (_wcsicmp(name, L"Name") == 0) {
        vb6_ho_setVariantBstr(out, vb6_BSTR_FromStr(h && h->name[0] ? h->name : (r && r->ctrlName[0] ? r->ctrlName : L"")));
        return 1;
    }
    if (_wcsicmp(name, L"Index") == 0) {
        vb6_ho_setVariantLong(out, h ? h->index : (r ? r->index : -1)); return 1;
    }
    if (_wcsicmp(name, L"ScaleMode") == 0) { vb6_ho_setVariantLong(out, scaleMode); return 1; }
    /* 账 #247: Left/Top/Width/Height 在这一层**没有第二份实现** —— 读写都问
       vb6forms_ctrl.c 那四个出口 (容器 ScaleMode 的单位 + VB 侧读数缓存)。
       以前这一支自己量窗口矩形 + 硬编码缇, 于是同一句
       `对象.Width = 7222` 走发码 (vb6_SetControlWidth: 存下 7222) 与走 IDispatch
       (宿主模型: 换算到像素就丢, 且不存) 给两种答案 —— 探针实测两条路都读回 7215。 */
    if (_wcsicmp(name, L"Left") == 0)   { vb6_ho_setVariantLong(out, vb6_GetControlLeft(obj));  return 1; }
    if (_wcsicmp(name, L"Top") == 0)    { vb6_ho_setVariantLong(out, vb6_GetControlTop(obj));   return 1; }
    if (_wcsicmp(name, L"Width") == 0)  { vb6_ho_setVariantLong(out, vb6_GetControlWidth(obj)); return 1; }
    if (_wcsicmp(name, L"Height") == 0) { vb6_ho_setVariantLong(out, vb6_GetControlHeight(obj)); return 1; }
    if (_wcsicmp(name, L"ScaleWidth") == 0 || _wcsicmp(name, L"ScaleHeight") == 0) {
        int32_t w = 0, hh = 0;
        vb6_ho_clientTwips(obj, &w, &hh);
        int32_t val = (_wcsicmp(name, L"ScaleWidth") == 0) ? w : hh;
        // ScaleWidth/ScaleHeight 需按对象自身 ScaleMode 换算
        if (scaleMode == 3)   /* Fix 184: twips -> pixels via the shared DPI source */
            val = (_wcsicmp(name, L"ScaleHeight") == 0) ? vb6_TwipToY(val) : vb6_TwipToX(val);
        vb6_ho_setVariantLong(out, val); return 1;
    }
    // Fix <vbeclipse>: `Visible` 必须读**自身 WS_VISIBLE 位**, 不能用 IsWindowVisible。
    // IsWindowVisible 是 Win32 的**链式**语义 —— 自身可见但任一祖先不可见就返回 0。
    // 而 VB6 的 `对象.Visible` 只反映该对象自身的 Visible 属性, 与祖先无关。
    // 混用会让停靠面板全空: ucFolder.ViewArea_Resize 里
    // `If View.View.Visible Then View.View.Move ...` 是布局的唯一驱动, 而它
    // 全部发生在 frmMain Form_Load 期间 —— 那时 frmMain 自己还没 ShowWindow,
    // 链上层层 IsWindowVisible=0 ⇒ Move 永不执行 ⇒ 视图窗体停在 0x0。
    // (实测 play78: [VR] enter x32, 每轮每项 vis=0; 同轮 [ST] 显示
    //  ShowView 已把视图 style 由 460B0040 写成 560B0040, WS_VISIBLE=1。)
    // 顺带这也修了拖动/最小化期间的误判: 祖先临时隐藏时对象自身 Visible 不变。
    if (_wcsicmp(name, L"Visible") == 0)  {
        LONG st = GetWindowLongW((HWND)obj, GWL_STYLE);
        vb6_ho_setVariantLong(out, (st & WS_VISIBLE) ? -1 : 0); return 1;
    }
    if (_wcsicmp(name, L"Enabled") == 0)  { vb6_ho_setVariantLong(out, IsWindowEnabled((HWND)obj) ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"hDC") == 0)      { vb6_ho_setVariantLong(out, (int32_t)(intptr_t)GetDC((HWND)obj)); return 1; }
    /* 账 #254: VB6 的 `Form.Count` = 这枚窗体上的子控件枚数，与 `Controls.Count` 同一件事，
       所以数法转调集合那一条出口，不在这里填常数。以前这一档硬填 0 —— #250 把 Controls
       集合接通之后 Charts 的 ClsResizer 第一次有了流量，它 `ReDim Rects(oForm.Count - 1)`
       铺出来的格子是空的，`For Each` 第一格写 `Rects(0)` 就是错误 9（下标越界），
       启动期整个进程退出（CI 把它报成「窗口 5s 内没出现」）。 */
    if (_wcsicmp(name, L"Count") == 0)    { vb6_ho_setVariantLong(out, vb6_UC_ControlsCountOf(obj)); return 1; }
    /* 账 #257: ParentControls = **容器**的控件集合 —— 不是本 UserControl 自己的子控件。
       ComCtlsDemo 里的用法全是 UC 去**父容器**上枚举兄弟控件(找 ImageList / 判可停靠):
         CheckBoxW.ctl:1267  For Each ControlEnum In UserControl.ParentControls
         CoolBar.ctl:689     If UserControl.ParentControls.Count = 0 Then CoolBarAlignable = False
       所以取样口径必须与下面 `Container` / `Parent` 那一档**同一条关系**(GetParent),
       否则「oCtrl 属于 X.Controls」与「oCtrl.Container = X」会给出两套互相矛盾的答案。
       ComCtlsDemo StatusBar UC 初始化时读它拿到 NULL ⇒ 宿主初始化中断
       ⇒ MainForm.Show 静默失败(CI 报「窗口 5s 内没出现」)。
       顶层窗体自己没有父 ⇒ 交回空值, 与 Container 那一档同口径。 */
    if (_wcsicmp(name, L"ParentControls") == 0) {
        void* pcHost = (void*)GetParent((HWND)obj);
        if (!pcHost) { vb6_ho_setVariantEmpty(out); return 1; }
        vb6_UCControls* cpc = vb6_uc_newControls(pcHost);
        vb6_ho_setVariantObject(out, cpc); return 1;
    }
    if (_wcsicmp(name, L"Controls") == 0) {
        /* 账 #250: 这一档以前用的是 setVariantDispatch —— 那一枚**刻意**只交 Empty
           (让字体代理那一路继续走 Nothing 分支, 见 uc_hostmodel.c 的注释)。
           rev14 给"真交对象"另立了出口 setVariantObject, 但只把 `Item` / `Add` 两个
           调用点改过去, `Controls` 这一档漏了 ⇒ 上层拿到 NULL, 下游三条
           (Controls.Count / Item / For Each，都在 vb6com_foreach.c 里备好了集合分支)
           一条也进不去: 实测窗体上 `Me.Controls.Count`=0、`For Each` 零条,
           Charts 的 ClsResizer (`For Each oCtrl In oForm.Controls`) 因此静默不做。
           释放侧安全: 集合指针由 vb6_Host_IsHostObject 认得, 且 vb6_ComIsDispatchable
           守卫不会对它调 Release —— 与 Item 那一档同一个口径。 */
        vb6_UCControls* c = vb6_uc_newControls(obj);
        vb6_ho_setVariantObject(out, c); return 1;
    }
    /* 账 #257: `Container` / `Parent` 这一档以前**压根没有** ⇒ `oCtrl.Container` 交回空值，
       `oCtrl.Container.Name` 交回空串。实测（探针 probe255 模式 P）Charts 的 ClsResizer
       保存循环第一句 `If oCtrl.Container.Name = oForm.Name Then` 因此靠「空 = 空」真空通过
       （窗体自己的 `.Name` 也是空，两边一起才是那个假通过；名字那一半在同一批补）。
       口径**不是新决定**：容器成员资格在 uc_controls.c 的 `vb6_uc_collectChildren` 里本来
       就是「GetParent(控件) == 容器」，这一档问同一条关系 ⇒ 「oCtrl.Container.Name = X.Name」
       与「oCtrl 属于 X.Controls」从此是同一件事，不会两套答案互相矛盾（PictureBox 有自己的
       Controls 集合、里面的 Label 不在窗体集合里，实测 n=5，与此口径一致）。
       交对象走 #250 立的那条唯一出口 `vb6_ho_setVariantObject`，不开第二份；
       顶层窗体自己没有父 ⇒ 交回空值（VB6 里 Form 无 Container）。
       另一层同名答案记下：`vb6forms_axcontainer.c:195` 那档是 OCX 扩展器对象（自己一套
       合成 IDispatch，记录里只有 hwndForm），问的不是这一条链，本刀刻意不动它。 */
    if (_wcsicmp(name, L"Container") == 0 || _wcsicmp(name, L"Parent") == 0) {
        void* parentHwnd = (void*)GetParent((HWND)obj);
        if (!parentHwnd) { vb6_ho_setVariantEmpty(out); return 1; }
        /* 账 #271: **父是 UserControl** 时, VB6 的 Container 交回的是**该 UserControl
           的 Extender**, 不是宿主窗口 —— 这条「同一性」在 VB6 里是真的, 而且是
           Krool/NewTab 那一族控件唯一的"这个控件确实摆在我身上"判据:
             ComCtlsDemo CoolBar.ctl:3173 `Set Container = Control.Container`
                 → `CBool(Err.Number = 0 And Not Control Is Extender And Container Is Extender)`
             NewTab ctlNewTab.ctl:12401 `If iContainer_Prev.Container Is UserControl.Extender Then`
           以前一律交回父窗口 HWND ⇒ 上面两处的 `Is Extender` **恒假** ⇒
           CoolBar 的 ControlIsValid 永远 False ⇒ `Err.Raise 380`
           (ComCtlsDemo 启动即未处理错误 380 退出, 验收位"窗口 5s 内出现"因此红)。
           父是**窗体**时维持原口径 —— Charts ClsResizer 的
           `oCtrl.Container.Name = oForm.Name` 问的就是那一档 (Extender 没有 Name)。 */
        if (vb6_uc_findByHwnd(parentHwnd)) {
            vb6_ho_setVariantObject(out, (void*)&vb6_UserControl_Extender);
            return 1;
        }
        vb6_ho_setVariantObject(out, parentHwnd); return 1;
    }
    if (_wcsicmp(name, L"Caption") == 0 || _wcsicmp(name, L"Text") == 0) {
        // Fix 190: 窗口层是 Unicode —— 必须 GetWindowTextW。旧代码用 GetWindowTextA
        // 再逐字节提升为 wchar_t (= 按 Latin-1 解码 ACP 字节), 中文读回来就是乱码,
        // UserControl 里 `Parent.Caption` 一类的读取全受影响。
        wchar_t wbuf[512] = {0};
        GetWindowTextW((HWND)obj, wbuf, 511);
        vb6_ho_setVariantBstr(out, vb6_BSTR_FromStr(wbuf)); return 1;
    }
    // 字体属性 (Font / TitleFont / *Font) → 返回 Font 代理
    {
        size_t ln = wcslen(name);
        if (_wcsicmp(name, L"Font") == 0 ||
            (ln > 4 && _wcsicmp(name + ln - 4, L"Font") == 0)) {
            vb6_UCFontRec* fr = (vb6_UCFontRec*)vb6_UC_NewFont();
            vb6_ho_setVariantDispatch(out, fr); return 1;
        }
    }
    // Fix <vbeclipse> rev18: 内建成员都没命中 → 试**本 UserControl 的自有 Property**
    // (ucFolder.LeftPos/TopPos/RightPos/BottomPos 这类)。cgen 为每个 .ctl 发的
    // vb6_UcPropDesc 表就是为这一档准备的 —— 没有它, ucPerspective.Refresh 的
    // `l_ucFolder.Move .LeftPos, .TopPos, …` 会读到 0, 停靠区就摆不出来。
    if (vb6_UC_OwnPropGet(obj, name, out)) return 1;
    /* 账 #305 第293刀: 递进来的常常是**宿主包装器** 或句柄, 不是 .ctl 实例本身;
       自有属性表 (vb6_UcPropDesc) 是按实例查的 ⇒ 再用记录里的实例试一次。
       实测 (ComCtlsDemo 下拉图标): `ControlEnum.hImageList` 两把钥匙都开不了门,
       一路答 Empty ⇒ CBEM_SETIMAGELIST 收 0, 条目有字无图。 */
    if (r && r->me && r->me != obj && vb6_UC_OwnPropGet(r->me, name, out)) return 1;

    /* 账 #305 第二百六十二刀 (丙): 同 Call 那条 —— 不认识的属性名以前静默答 Empty。 */
    vb6_ho_reportUnknown(L"GetProp", name, obj);
    return 1;   // 宿主对象: 未知属性一律返回 Empty (不崩)
