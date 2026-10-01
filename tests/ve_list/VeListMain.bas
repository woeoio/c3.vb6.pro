Option Explicit

' Fix <vbeclipse> rev16 回归夹具: 工程内 List 类的 Variant 数组读写
'
' 由来: 真工程 VbEclipse 的 Classes/List.cls 用
'   Dim m_Keys() As String / Dim m_Items() As Variant
'   ReDim Preserve m_Keys(Idx) / ReDim Preserve m_Items(Idx)
' 存 BSTR / 对象, 再由 `Item(ByVal Index As Variant)` 按 **数值下标** 取回。
' 实测 play78r15.exe(真工程 x86): `Folder.Views.Item(0)` 取回的不是 "Properties"
' 而是一个 1 字符的垃圾串 → 后续 m_Views.Item(<垃圾>) 落空 → NULL 解引用
' (av read 0x4 @ vb6_View_prop_get_ViewId+0xF)。
' 夹具把这条路径缩到最小：**类成员** Variant 数组 + ReDim Preserve + 按下标取回。

Sub Main()
    Dim L As New VbList
    Dim v As Variant
    Dim s As String
    Dim idx As Long

    ' 0) 对照组: 纯局部数组 (非类成员) 的 ReDim Preserve + 取回
    Dim a() As Variant
    ReDim a(0)
    a(0) = "A"
    ReDim Preserve a(1)
    a(1) = "B"
    Debug.Print "LOCAL a0=[" & a(0) & "] a1=[" & a(1) & "]"

    Dim k() As String
    ReDim k(0)
    k(0) = "K0"
    ReDim Preserve k(1)
    k(1) = "K1"
    Debug.Print "LOCAL k0=[" & k(0) & "] k1=[" & k(1) & "]"

    ' 1) 类成员: 三个可区分元素, 判断是"整体错位一格"还是"槽 0 被覆盖"
    v = "Alpha"
    L.Add "Alpha", v
    v = "Beta"
    L.Add "Beta", v
    v = "Gamma"
    L.Add "Gamma", v

    Debug.Print "count=" & L.Count

    For idx = 0 To 2
        s = L.Item(idx)
        Debug.Print "idx=" & idx & " len=" & Len(s) & " s=[" & s & "]"
    Next idx

    ' 字面量下标 (与变量下标对照)
    Debug.Print "lit0=[" & L.Item(0) & "] lit1=[" & L.Item(1) & "] lit2=[" & L.Item(2) & "]"

    ' 2) 按 key 取
    Debug.Print "bykeyAlpha=[" & L.Item("Alpha") & "]"
    Debug.Print "bykeyGamma=[" & L.Item("Gamma") & "]"

    ' 3) 取回的 Variant 的 vt (期望 8 = VT_BSTR)
    Debug.Print "vt0=" & VarType(L.Item(0)) & " vt2=" & VarType(L.Item(2))

    ' 4) Contains: 检验 m_Keys 本身是否正确
    Debug.Print "hasAlpha=" & L.Contains("Alpha")
    Debug.Print "hasBeta=" & L.Contains("Beta")
    Debug.Print "hasGamma=" & L.Contains("Gamma")
    Debug.Print "hasNope=" & L.Contains("Nope")
    Debug.Print "indexOfGamma=" & L.IndexOf("Gamma")

    ' 5) 工程类实例经 Variant 槽往返 —— 真工程 View/List 的最小形态。
    '    真工程链: ucPerspective.AddView 里 `Dim l_View As View: Set l_View = New View`
    '    → `m_Views.Add ViewId, l_View` (ByRef Variant 装箱) → `Set l_View = Nothing`
    '    → CreateFolder 里 `m_Views.Item(ViewId)` 取回 → 交给
    '    `ucFolder.AddView(ByRef View As View)` → 里面读 `View.ViewId`。
    '    夹具把这条链逐步打点, 定位在哪一步丢的对象。
    Dim vv As VbView
    Dim back As VbView
    Set vv = New VbView
    vv.ViewId = "Props"
    L.Add "Props", vv
    Debug.Print "objcnt=" & L.Count
    Debug.Print "objvt=" & VarType(L.Item("Props"))
    Debug.Print "objnil=" & (L.Item("Props") Is Nothing)
    Set back = L.Item("Props")
    Debug.Print "objidA=[" & back.ViewId & "]"
    Set vv = Nothing
    Debug.Print "objidB=[" & back.ViewId & "]"
    TakeView L.Item("Props")
    Debug.Print "objidC=[" & back.ViewId & "]"

    ' 6) 成员与**类同名**的类经 Variant 槽往返 (真工程 View.cls / PopupMenu.cls 形态)。
    '    这类类符号会被语义层从模块作用域挤掉 → 不进 coclass 表 →
    '    运行期 vb6_FindCoClassDesc(<类名>) 返 NULL → 装箱出 vt=9/pdisp=NULL 的空壳
    '    Variant → 槽位 `Is Nothing` 恒真。夹具同时覆盖 Property 形态与 Sub 形态。
    Dim sf As VbSelf
    Dim sfBack As VbSelf
    Dim pp As VbPop
    Dim ppBack As VbPop
    Set sf = New VbSelf
    sf.VbSelf = "SelfId"
    L.Add "Self", sf
    Debug.Print "selfvt=" & VarType(L.Item("Self"))
    Debug.Print "selfnil=" & (L.Item("Self") Is Nothing)
    Set sfBack = L.Item("Self")
    Debug.Print "selfid=[" & sfBack.VbSelf & "]"
    Set sf = Nothing

    Set pp = New VbPop
    pp.VbPop
    L.Add "Pop", pp
    Debug.Print "popnil=" & (L.Item("Pop") Is Nothing)
    Set ppBack = L.Item("Pop")
    Debug.Print "pophits=" & ppBack.Hits
    Set pp = Nothing
    Debug.Print "selfid2=[" & sfBack.VbSelf & "] pophits2=" & ppBack.Hits

    ' 7) ByRef 类形参 → ByRef As Variant 槽 (真工程 ucFolder.AddView 形态)
    Dim L2 As New VbList
    Dim fwd As VbView
    Dim fwdBack As VbView
    Set fwd = New VbView
    fwd.ViewId = "FwdId"
    ForwardTo fwd, L2
    Debug.Print "fwdvt=" & VarType(L2.Item("FwdId"))
    Debug.Print "fwdnil=" & (L2.Item("FwdId") Is Nothing)
    Set fwd = Nothing
    Set fwdBack = L2.Item("FwdId")
    Debug.Print "fwdcnt=" & L2.Count & " fwdid=[" & fwdBack.ViewId & "]"

End Sub

' 对应真工程 ucFolder.AddView(ByRef View As View) 的取用侧
Private Sub TakeView(ByRef View As VbView)
    Debug.Print "takeid=[" & View.ViewId & "]"
End Sub

' Fix <vbeclipse> rev17: `As <工程类>` 的 **ByRef 形参** 交给 `ByRef ... As Variant` 槽。
' 真工程形态: ucFolder.AddView 的 `m_FolderViews.Add View.ViewId, View` —— 生成码原先是
'   vb6_List_Add(me->m_FolderViews, ..., View, -1, 0);
' 把 vb6_cls_View** 当 vb6_VARIANT* 解 (MSVC 只给 C4133 警告), 槽里存进垃圾 Variant ⇒
' ucFolder.ShowView 里 `Set l_View = .Item(i)` 得到 NULL ⇒ `l_View.ViewId` av read 0x4。
Private Sub ForwardTo(ByRef View As VbView, ByRef Dest As VbList)
    Dest.Add View.ViewId, View
End Sub
