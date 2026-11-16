VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart walls"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6615
   Begin TeeChart.TChart TChart1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' C30-10 第二十刀：勾选框那批属性的绘制判据夹具（二）—— 四壁分面。
' 对应 ChartDisplay 的 Check4(No Back colour) / Command4(LeftWall) / Command11(Walls back colour)。
Private Sub Form_Load()
    Dim i As Integer
    TChart1.AddSeries 1
    TChart1.Series(0).Title = "S0"
    For i = 1 To 5
        TChart1.Series(0).Add i, i * 2
    Next i
    TChart1.Aspect.View3D = True
    TChart1.Walls.Visible = True
    ' 背板刷成绿色（Brush 优先于 Walls.BackColor），左壁整面关掉
    TChart1.Walls.Back.Transparent = False
    TChart1.Walls.Back.Brush.Color = vbGreen
    TChart1.LeftWall.Visible = False
    Debug.Print "TC-WALL-READY"
End Sub
