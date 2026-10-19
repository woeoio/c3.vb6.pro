Attribute VB_Name = "TofMain"
Option Explicit

' Fix 193: `TypeOf lhs Is <项目类>` 的夹具。
'   坏的时候: `vb6_TypeOf` 是 vb6rtl_conv.c 里一个**恒返 0 的桩** (注释写着
'   "简化版, 始终返回 False") → 项目类这一位一律答"否", 连 `TypeOf raw Is
'   ShapeAct` (Raw 就声明成 ShapeAct) 都是 False。只有接口名那条路
'   (`TypeOf iv Is IShapeAct`) 走 vb6_IfaceSupports, 是好的。
'   修法 = 编译期按"声明类 + 祖先链"判定 (inferClassTypeOfExpr + ClassChainView)。
'
' 链: TofLeaf -> TofMid -> TofBase, 另有兄弟 TofSib (也 Inherits TofBase)
' 与孤立的 TofPlain (无继承、无虚槽)。
'
' 登记在案的残留边界 (不静默, 见 src/backend/expr/cgen_expr.cpp 的 Fix 193 注释):
'   `Dim b As TofBase : Set b = New TofLeaf : TypeOf b Is TofLeaf` —— VB6 按
'   **实际**类型答"是", 静态判定按**声明**类型答"否"。这是假阴性, 与改前
'   恒假同向, 不会把原本对的翻成错的; 修它需要给类加运行时类型标记 (动结构体
'   布局), 代价与收益不成比例 → 不在这批里做。因此夹具里 TOF7 只断言
'   "声明类型即实际类型"的那一半 (Set m = New TofMid, Is TofLeaf 应为 False)。

Sub Main()
    Dim leaf As TofLeaf
    Set leaf = New TofLeaf

    ' TOF1: 声明类 == 目标类
    If TypeOf leaf Is TofLeaf Then
        Debug.Print "TOF1:OK"
    Else
        Debug.Print "TOF1:FAIL (自身类应 True)"
    End If

    ' TOF2: 目标类是直接祖先
    If TypeOf leaf Is TofMid Then
        Debug.Print "TOF2:OK"
    Else
        Debug.Print "TOF2:FAIL (直接祖先应 True)"
    End If

    ' TOF3: 目标类是链根
    If TypeOf leaf Is TofBase Then
        Debug.Print "TOF3:OK"
    Else
        Debug.Print "TOF3:FAIL (链根应 True)"
    End If

    ' TOF4: 兄弟类 (同一祖先, 但不在本链上)
    If Not (TypeOf leaf Is TofSib) Then
        Debug.Print "TOF4:OK"
    Else
        Debug.Print "TOF4:FAIL (兄弟类应 False)"
    End If

    ' TOF5: 无关类
    If Not (TypeOf leaf Is TofPlain) Then
        Debug.Print "TOF5:OK"
    Else
        Debug.Print "TOF5:FAIL (无关类应 False)"
    End If

    Dim sib As TofSib
    Set sib = New TofSib

    ' TOF6: 兄弟类自身的链: 是 TofBase, 不是 TofMid/TofLeaf
    If (TypeOf sib Is TofBase) And Not (TypeOf sib Is TofMid) And Not (TypeOf sib Is TofLeaf) Then
        Debug.Print "TOF6:OK"
    Else
        Debug.Print "TOF6:FAIL (兄弟链判定错)"
    End If

    Dim mid As TofMid
    Set mid = New TofMid

    ' TOF7: 父类变量装父类实例 → Is 子类 应 False (声明类型即实际类型那一半)
    If (TypeOf mid Is TofBase) And Not (TypeOf mid Is TofLeaf) Then
        Debug.Print "TOF7:OK"
    Else
        Debug.Print "TOF7:FAIL (父类实例不应是子类)"
    End If

    Dim n As TofLeaf

    ' TOF8: 未 Set 的变量 (Nothing) 不匹配任何类型
    If Not (TypeOf n Is TofLeaf) And Not (TypeOf n Is TofBase) Then
        Debug.Print "TOF8:OK"
    Else
        Debug.Print "TOF8:FAIL (Nothing 应 False)"
    End If

    ' TOF9: Set ... = Nothing 之后同样不匹配
    Set leaf = New TofLeaf
    If TypeOf leaf Is TofLeaf Then
        Set leaf = Nothing
        If Not (TypeOf leaf Is TofLeaf) Then
            Debug.Print "TOF9:OK"
        Else
            Debug.Print "TOF9:FAIL (Set Nothing 后应 False)"
        End If
    Else
        Debug.Print "TOF9:FAIL (前置条件不成立)"
    End If

    Dim p As TofPlain
    Set p = New TofPlain

    ' TOF10: 无继承、无虚槽的普通类: 同名 True, 别的 False
    If (TypeOf p Is TofPlain) And Not (TypeOf p Is TofBase) Then
        Debug.Print "TOF10:OK"
    Else
        Debug.Print "TOF10:FAIL (普通类判定错)"
    End If

    ' TOF11: 类数组元素 (Fix 192 的元素类名登记 + Fix 193 的判定一起验)
    Dim arr(0 To 1) As TofBase
    Set arr(0) = New TofLeaf
    Set arr(1) = New TofSib
    If (TypeOf arr(0) Is TofBase) And (TypeOf arr(1) Is TofBase) Then
        Debug.Print "TOF11:OK"
    Else
        Debug.Print "TOF11:FAIL (类数组元素判定错)"
    End If

    ' TOF12: 判定结果能进 If 的分支 (不是只在 Debug.Print 里对)
    Dim cnt As Long
    cnt = 0
    Set leaf = New TofLeaf
    If TypeOf leaf Is TofBase Then cnt = cnt + 1
    If TypeOf leaf Is TofSib Then cnt = cnt + 10
    If cnt = 1 Then
        Debug.Print "TOF12:OK"
    Else
        Debug.Print "TOF12:FAIL cnt=" & cnt
    End If

    ' TOF13 (账 #446 / §B218): `TypeOf x Is Object` —— Object 是基类, 等价于"活的对象引用"。
    '   此前落到 vb6_TypeOf 那个恒返 0 的桩 ⇒ 一律答"否" (含活对象)。修后: 活引用 True,
    '   Nothing 不匹配任何类型 (含 Object) False。COM 晚绑对象 (CreateObject) 也走这一位。
    Set leaf = New TofLeaf
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If (TypeOf leaf Is Object) And (TypeOf d Is Object) And Not (TypeOf n Is Object) Then
        Debug.Print "TOF13:OK"
    Else
        Debug.Print "TOF13:FAIL (Object 基类判定错)"
    End If

    ' TOF14 (账 #447 / §B219): `TypeOf x IsNot T` —— ai/032 的 IsNot 接到 TypeOf 构造上,
    '   解析期脱糖成 Not(TypeOf x Is T) (同 binary a IsNot b)。与 TOF13 互补: 活对象
    '   IsNot Object=False, Nothing IsNot Object=True, COM 晚绑对象 IsNot Object=False。
    If Not (TypeOf leaf IsNot Object) And (TypeOf n IsNot Object) And Not (TypeOf d IsNot Object) Then
        Debug.Print "TOF14:OK"
    Else
        Debug.Print "TOF14:FAIL (TypeOf IsNot 判定错)"
    End If

    Debug.Print "TOF-DONE"
End Sub
