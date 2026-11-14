VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K612"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 账 #612: 一维动态数组成员 —— 类外 ReDim / ReDim Preserve / Erase 的目标左值
Public Digits() As Long
Public Tags() As String

' 类形参那一档 owner (registerClassFormal 把它登进 knownClassVars_)
Public Sub Grow(ByVal n As K612, ByVal k As Long)
    ReDim n.Digits(1 To k)
    n.Digits(k) = k * 10
End Sub

Public Function Sum(ByVal n As K612) As Long
    Dim i As Long
    For i = LBound(n.Digits) To UBound(n.Digits)
        Sum = Sum + n.Digits(i)
    Next i
End Function
