VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CB448"
Option Explicit
Implements IB448
Private mN As Long

Public Function Twice(ByVal n As Long) As Long Implements IB448.Twice
    mN = mN + 1
    Main.gHits48 = Main.gHits48 + 1
    Twice = n * 2 + 1
End Function

Public Function Hits() As Long
    Hits = mN
End Function

' 第一百三十三刀的判据要走「这一枚实例被拆了没有」那一问, 而视图那一面只认 Twice 一个成员,
' 所以把拆开的次数记到模块级计数器上 (W20 那一格读它)。
Private Sub Class_Terminate()
    Main.gTerm48 = Main.gTerm48 + 1
End Sub
