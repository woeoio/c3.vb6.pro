VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart donut"
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
' C30-10 圆环图判据夹具：401x257，AddSeries 17（stDonut），三值 30/45/25。
' 判据（tc_type_px.ps1 -Mode donut）：圆心必须是 Panel 底色（洞），
' 且半径 0.75R 处必须是系列色（环）—— 两条同时成立才叫“环”而不是“饼”。
Private Sub Form_Load()
    TChart1.AddSeries 17
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddXY 1, 30
    TChart1.Series(0).AddXY 2, 45
    TChart1.Series(0).AddXY 3, 25
    TChart1.Aspect.View3D = False
    Debug.Print "TC-DN-TYPE=" & TChart1.Series(0).SeriesType
    Debug.Print "TC-DN-READY"
End Sub
