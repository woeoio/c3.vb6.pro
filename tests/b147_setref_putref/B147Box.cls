Option Explicit

' 写口那一位: Property Set 收对象, 存进模块级 gHold。
Public Property Set Held(ByVal v As Object)
    Set gHold = v
End Property

' 同一个名字的读口 —— 与 Set 共用一枚 dispid（COM 口径）。
Public Property Get Held() As Object
    Set Held = gHold
End Property

Private mNum As Long

' R 路反面证人: 同名 Let/Get 必须照旧走 PROPERTYPUT(4) 那一档。
Public Property Let Num(ByVal v As Long)
    mNum = v
End Property

Public Property Get Num() As Long
    Num = mNum
End Property

Public Sub Bump(ByVal d As Long)
    mNum = mNum + d
End Sub

Public Function TagOf(ByVal v As Object) As Long
    TagOf = 0
    If Not v Is Nothing Then TagOf = 1
End Function
