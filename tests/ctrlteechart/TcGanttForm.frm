VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart gantt"
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
' C30-10 甘特图判据夹具：401x257，AddSeries 9（stGantt），三条任务：
'   (1→4) (2→5) (3→6)，横轴用本系列自己的 x 定义域 [1,6]。
' 判据（tc_type_px.ps1 -Mode gantt）：三条横带、最左=plot.left、最右=plot.right。
Private Sub Form_Load()
    TChart1.AddSeries 9
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddXY 1, 4
    TChart1.Series(0).AddXY 2, 5
    TChart1.Series(0).AddXY 3, 6
    TChart1.Aspect.View3D = False
    Debug.Print "TC-GT-TYPE=" & TChart1.Series(0).SeriesType
    Debug.Print "TC-GT-READY"
End Sub
