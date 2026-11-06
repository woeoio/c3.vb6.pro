VERSION 5.00
Begin VB.Form E438 
   Caption         =   "E438-START"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "E438"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim v As Variant
    Dim n1 As Long

    ' 一档: On Error Resume Next 吞 438, Err.Number 要读到 438
    On Error Resume Next
    v = Me.rtfText
    n1 = Err.Number
    On Error GoTo 0

    ' 二档: On Error GoTo 必须真的跳进处理器 (反向用例: 438 不许被一起吞掉)
    On Error GoTo EH
    v = Me.rtfText
    Caption = "E438-PASS-" & n1
    Exit Sub

EH:
    Caption = "E438-GOTO-" & n1 & "-" & Err.Number
End Sub
