VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K613"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 账 #613 (第351刀): 元素本身是 UDT 的那一档 —— 分配步长必须等于 sizeof(TBig),
' 修前按 variant 24 字节一步、访问按 sizeof(TBig) 64 字节走 ⇒ 越界写。
Public Type TBig
    A As Long
    B As Long
    C As Long
    D As Long
    E As Long
    F As Long
    G As Long
    H As Long
End Type

Public L() As Long
Public Big() As TBig
