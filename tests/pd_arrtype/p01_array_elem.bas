Attribute VB_Name = "pdArrType"
Option Explicit

' 探针夹具 (R2/R3, PhotoDemon 缺口): 数组类型在**表达式类型**上必须保持一致。
'   1) 函数返回数组 `As Byte()`            —— R2: 返回类型带 Array 位, 但局部数组变量只记 isArray
'   2) 整体数组赋值 `f = a` / `b = f()`      —— R2: 两侧都必须是"数组 vs 数组"
'   3) ByRef 数组形参 `dst() As Byte` 的下标 —— R3: 形参的 Array 位在类型里, isArray 为假,
'                                               下标表达式因此退回数组类型本身而不是元素类型
'   4) UDT 数组形参的元素成员               —— R3 同族 (元素是 UDT 时同理)
' 判据: --emit-c 的 stderr 不许出现 VB3003 (类型退化/隐式转换告警)。
' 改前读数: 4 条 VB3003 (`Byte → ?` / `? → Byte` / `Integer → ?` ×2)。

Public Type Quad
    C As Byte
End Type

Public Function MakeArr() As Byte()
    Dim a() As Byte
    ReDim a(0 To 1)
    a(0) = 1
    MakeArr = a
End Function

Public Sub FillArr(ByRef dst() As Byte)
    ReDim dst(0 To 3)
    dst(0) = 7
End Sub

Public Sub FillQuad(ByRef q() As Quad)
    q(0).C = 9
End Sub

Public Sub Consume()
    Dim b() As Byte
    b = MakeArr()
    FillArr b
    Debug.Print b(0)
End Sub
