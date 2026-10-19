VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C445H"
Option Explicit

' 账 #305 §B218 (甲) 腿⑤ 的靶子: 一枚工程类字段 —— 持有者出厂时, 字段里那一枚必须走 Class_Terminate。
Private mInner12 As C445
' 反面针的形状: As Object 是一格 void*, 不属于工程类那一档, 出厂这一站不许碰它。
Public mLate15 As Object

Private Sub Class_Initialize()
    Main.MadeH = Main.MadeH + 1
End Sub

Private Sub Class_Terminate()
    Main.TermsH = Main.TermsH + 1
End Sub

Public Sub Hold12()
    Set mInner12 = New C445
End Sub

Public Sub HoldTwice12()
    Set mInner12 = New C445
    Set mInner12 = New C445
End Sub

Public Sub AliasInto12(ByVal o As C445)
    Set mInner12 = o
End Sub

' 账 #305 §B218 (甲) 刀3 的第二形: 工程内的 Property Get 工厂 —— 返回槽那一枚同样是交给调用方的,
' 调用方再接手时不许再 claim (与 MakeOne 那一发同一条契约, 所以要同一问句答出来)。
Public Property Get NewOne16() As C445
    Dim g As C445
    Set g = New C445
    Set NewOne16 = g
End Property

Public Sub PropFactory16()
    Dim t As C445
    Set t = NewOne16()
End Sub

' 账 #305 §B218 (甲) 刀4 的两形: 工厂挂在**实例**上 (`hh.Make17()`) 与属性读**不带括号**
' (`hh.Next18`) —— 交出来的那一枚同样自带对价, 接手这一格不许再 claim。
Public Function Make17() As C445
    Dim g As C445
    Set g = New C445
    Set Make17 = g
End Function

Public Property Get Next18() As C445
    Dim g As C445
    Set g = New C445
    Set Next18 = g
End Property

' 反面那一档 (刀4 的边界): 工程类字段是**数据字段**, 那一枚仍由持有者握着 —— 读出来必须
' claim 一枚, 否则调用方放手就把持有者手里那枚拆了。
Public mField19 As C445

Public Sub Hold19()
    Set mField19 = New C445
End Sub
