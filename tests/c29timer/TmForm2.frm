VERSION 5.00
Begin VB.Form TmForm2 
   Caption         =   "TmForm2"
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2400
   LinkTopic       =   "TmForm2"
   ScaleHeight     =   1200
   ScaleWidth      =   2400
   Begin VB.Timer t2 
      Enabled         =   -1   'True
      Interval        =   20
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "TmForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 账 #156: 第二枚窗体，只为自己那枚 Timer 存在。
' 计时器 id 一旦跟着"每建一枚窗体就复位"的控件 id 计数器走，这枚窗体的 Timer 就与
' 第一枚窗体的第一枚 Timer 同号 —— 派发只按 id 查进程内那张表，先建的那格先命中，
' 于是这里的 t2_Timer 一次都不会跑（第一枚窗体里读到的数就是 0），而第一枚的事件
' 过程跑得翻倍。两边各读一次，才是这一条的完整形状。
Private m2 As Long

Private Sub t2_Timer()
    m2 = m2 + 1
    TmForm.SetDlg m2
End Sub
