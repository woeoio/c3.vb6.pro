Option Explicit

' ReDim 的两种形态（都在 PhotoDemon 里出现）：
'   1. 多数组：`ReDim a(0 To 3) As Long, b(0 To 5) As Long`（修前第二个数组不认）
'   2. 无括号成员链：`ReDim m(0).lMask.cProps(0 To 3)`（pdXCF.cls 1028 行形态：
'      .lMask 后面不跟括号，只有最后一段带维度；修前在这里断 ⇒ VB2003）

Private Type tProp
    id As Long
End Type
Private Type tMask
    cProps() As tProp
End Type
Private Type tLayer
    lMask As tMask
End Type

Public Sub Main()
    Dim a() As Long
    Dim b() As Long

    ReDim a(0 To 3) As Long, b(0 To 5) As Long
    a(3) = 7
    b(5) = 9

    Dim m As tLayer
    ReDim m.lMask.cProps(0 To 3) As tProp
    m.lMask.cProps(2).id = 42

    Debug.Print "RS=" & a(3) & "," & b(5) & "," & m.lMask.cProps(2).id
End Sub
