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
' C30-10 四壁判据夹具：401x257，一条柱三根等高（避免"值=vmin 高度为 0"那种退化）。
' View3D=True / Perspective=False ⇒ 深度常量 12，墙的几何可预测：
'   左壁从 x=28（plot.left）起、背板右缘 302+12=314、背板顶 27-12=15。
Private Sub Form_Load()
    TChart1.AddSeries 1
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddXY 1, 20
    TChart1.Series(0).AddXY 2, 20
    TChart1.Series(0).AddXY 3, 20
    TChart1.Aspect.View3D = True
    TChart1.Aspect.Perspective = False
    Debug.Print "TC-WALL-VIS=" & TChart1.Walls.Visible
    Debug.Print "TC-WALL-COLOR=" & TChart1.Walls.Color
    Debug.Print "TC-WALL-BACK=" & TChart1.Walls.Back.Visible
    Debug.Print "TC-WALL-READY"
End Sub
