VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbHold"
Option Explicit

' <vbeclipse> 账 #212 那一面的载体：运行时错误 9 从**实例方法体内**抛出来。
' 来源是 #209/#214 —— 越界的元素访问改成抛 9 之后，真工程里的调用点全在 UC 的实例方法里
' (ucChartBar 那句 `With m_Serie(Index)` 取的就是 &(me->m_arr(k)) 当对象)，
' 而此前只量过「没处理器 ⇒ Unhandled error 9 + 进程按 9 退出」这一面。
' 下面四枚方法把两面分开问：前三枚自己写着 On Error GoTo (处理器该接住)，
' NoHandler 那枚故意不写 (9 该继续往上交回调用方)。

Private m_items(1 To 4) As Long
Private m_a2(1 To 2, 1 To 2) As Long
Private m_dyn() As Long

Private Sub Class_Initialize()
    Dim i As Long, j As Long
    For i = 1 To 4
        m_items(i) = i * 10
    Next
    For i = 1 To 2
        For j = 1 To 2
            m_a2(i, j) = i * 10 + j
        Next
    Next
End Sub

' 一维越界（账 #209 那一支）
Public Function OneDimAbove() As Long
    On Error GoTo H
    OneDimAbove = m_items(9)
    Exit Function
H:
    OneDimAbove = Err.Number
End Function

' 二维越界（账 #214 那一支）
Public Function TwoDimAbove() As Long
    On Error GoTo H
    TwoDimAbove = m_a2(3, 1)
    Exit Function
H:
    TwoDimAbove = Err.Number
End Function

' 从没 ReDim 过的成员动态数组：以前读 (NULL)->data 当场 0xC0000005
Public Function NullDyn() As Long
    On Error GoTo H
    NullDyn = m_dyn(1)
    Exit Function
H:
    NullDyn = Err.Number
End Function

' 这一枚故意没有处理器：错误该往上走到调用方那一层
Public Function NoHandler(ByVal ix As Long) As Long
    NoHandler = m_items(ix)
End Function

' 证人：三次抛出之后，这枚对象的成员照旧读得出数（实例栈没留在中间态）
Public Property Get At(ByVal ix As Long) As Long
    At = m_items(ix)
End Property
