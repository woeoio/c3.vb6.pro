VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart builtin"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4500
   Begin TeeChart.TChart TChart1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   60
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' C30-10 夹具：内置 TChart 控件（不加载 TeeChart5.ocx）。
' 本刀只验「控件建得出窗口 + 画得出一遍」：绘制次数由 RTL 的 C3_TCHART_TRACE 打点。
Private Sub Form_Load()
    Debug.Print "TC-BUILTIN-START"
    Debug.Print "TC-WIDTH=" & TChart1.Width
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
