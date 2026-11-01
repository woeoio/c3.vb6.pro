VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Button"
Option Explicit

Public Event Click(ByVal X As Long, ByVal Y As Long)
Public Event Resize(ByVal Width As Long, ByVal Height As Long)

Private m_Caption As String

Public Sub SetCaption(ByVal Value As String)
    m_Caption = Value
End Sub

Public Sub DoClick(ByVal X As Long, ByVal Y As Long)
    RaiseEvent Click(X, Y)
End Sub

Public Sub DoResize(ByVal W As Long, ByVal H As Long)
    RaiseEvent Resize(W, H)
End Sub

' Knife 256: a ByRef event param is an OUT param - the handler's write must
' reach the raiser after RaiseEvent returns (ComCtlsDemo ListView
' GetEmptyMarkup / empty-list markup text is exactly this shape).
Public Event GetEmptyMarkup(ByRef Text As String, ByRef Center As Boolean)

Public Function RenderMarkup() As String
    Dim T As String
    Dim C As Boolean
    T = ""
    C = False
    RaiseEvent GetEmptyMarkup(T, C)
    If C Then
        RenderMarkup = T & "|CENTERED"
    Else
        RenderMarkup = T & "|LEFT"
    End If
End Function

' 账 #305 308: 形参声明是 Variant 而实参是一枚标量局部 —— 这一档必须先装箱递地址，
'   处理器写进盒里的值再按实参自己的档回落。ComCtlsDemo ListView 的 GetVirtualItem
'   (Text / CallbackText 两枚 String 形参位) 就是这个形状：直递 &Text 让处理器把一枚
'   24B(x86 20B) 的槽写进 4/8 字节的标量存储里，现场是原生 AV 打死客户端。
'   两条 RaiseEvent 刻意住在同一枚函数里 —— 盒子的名字按实参序号拼，撞名就是 cl C2371。
'   V As Variant 那一发是反面证人：它本来就是槽，不许再多套一层盒。
Public Event FillText(ByRef Val As Variant)
Public Event FillNum(ByRef Val As Variant)

Public Function OutShapes() As String
    Dim S As String
    Dim L As Long
    Dim V As Variant
    S = "abc"
    L = 7
    V = "keep"
    RaiseEvent FillText(S)
    RaiseEvent FillNum(L)
    RaiseEvent FillText(V)
    OutShapes = S & "/" & L & "/" & V
End Function
