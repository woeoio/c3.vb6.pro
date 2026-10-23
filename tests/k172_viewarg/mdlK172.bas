Attribute VB_Name = "mdlK172"
Option Explicit

' <vbeclipse> fixture driver (账 #305 第一百七十二刀 / 新账 #4xx).
' Three stations where a legacy interface view (vb6_iface_<I> = a value struct)
' sits on the receiving end, all answered by owners that already exist:
'   站B  instance handed to a view formal  -> legacyIfaceWrapFn (knife 166)
'   站C  view Is view / view Is Nothing    -> slotIsLegacyIfaceView (knife 165)
'   站E  view into a ByRef Variant formal  -> legacyViewBoxExpr (knife 171)
' Before the fix every one of these was rejected by cl (C2440 / C2088 / C2198),
' so this whole project had no exe at all -- same as the Ray.vbp sites.
' 站A (For Each <view> In <collection>) and 站D (zero-arg view method as a
' statement) are deliberately NOT here: they need a runtime interface table and
' are booked as their own account.

Public Function ViewValue(ByVal target As IK172) As Long
    ViewValue = target.Value()
End Function

Public Function BoxValue(ByRef slot As Variant) As Long
    BoxValue = VarType(slot)
End Function

Public Sub Main()
    Dim a As K172Alpha
    Dim b As K172Beta
    Set a = New K172Alpha
    Set b = New K172Beta

    ' 站B: 类里把 Me 递进视图形参 -> 包成自己那一枚实现类的视图
    Debug.Print "K172-pushA="; a.PushInto()
    Debug.Print "K172-pushB="; b.PushInto()

    ' 站B 反面证人: 视图交给视图形参 = 值结构体直传, 不许再包一层
    Dim v As IK172
    Set v = a
    Debug.Print "K172-view2view="; ViewValue(v)

    ' 站C: 视图 Is 视图 / Is Nothing
    Dim w As IK172
    Set w = b
    If v Is w Then
        Debug.Print "K172-is-diff=1"
    Else
        Debug.Print "K172-is-diff=0"
    End If
    Dim u As IK172
    Set u = a
    If v Is u Then
        Debug.Print "K172-is-same=1"
    Else
        Debug.Print "K172-is-same=0"
    End If
    Dim nVar As IK172
    If nVar Is Nothing Then
        Debug.Print "K172-nothing=1"
    Else
        Debug.Print "K172-nothing=0"
    End If
    Set nVar = w
    If nVar Is Nothing Then
        Debug.Print "K172-after=1"
    Else
        Debug.Print "K172-after=0"
    End If

    ' 站E: 视图递进 ByRef Variant 形参那一格
    Debug.Print "K172-boxtype="; b.BoxInto()

    Debug.Print "K172-DONE"
End Sub
