VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B190Holder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 类成员的**动态**数组: 声明给档 Long (int32_t, 4 字节一步),
' 而 ReDim 那条路若回落 vb6_sa_variant (16/24 字节一步) 就是本格要量的形。
Public Items() As Long
Public Tags() As String

' 真语料那一形 (UltimaX `TextureTool.Vertices(i).tU`): UDT 元素的成员动态数组
Public Vertices() As Vert

Public Sub InitV(ByVal n As Long)
    ReDim Vertices(1 To n)
End Sub

Public Sub Init(ByVal n As Long)
    ReDim Items(1 To n)
    ReDim Tags(1 To n)
End Sub

Public Function GetDirect(ByVal i As Long) As Long
    GetDirect = Items(i)
End Function

Public Function GetTagDirect(ByVal i As Long) As String
    GetTagDirect = Tags(i)
End Function

Public Function SumAll() As Long
    Dim i As Long
    Dim s As Long
    For i = 1 To UBound(Items)
        s = s + Items(i)
    Next i
    SumAll = s
End Function

Public Function LoAll() As Long
    LoAll = LBound(Items)
End Function

Public Function HiAll() As Long
    HiAll = UBound(Items)
End Function

Public Function LastIdx() As Long
    LastIdx = UBound(Items)
End Function
