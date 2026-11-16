VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart area3d"
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
' C30-10 面积 3D 带判据夹具：401x257，一条面积三值 20/30/25，色固定。
' 判据（tests\ctrlteechart\tc_area3d_px.ps1）量前后两闸的位置：
'   后闸（暗 45%）整体比前闸高 12px、右缘多 12px —— 就是带的厚度。
Private Sub Form_Load()
    TChart1.AddSeries 3
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddXY 1, 20
    TChart1.Series(0).AddXY 2, 30
    TChart1.Series(0).AddXY 3, 25
    TChart1.Aspect.View3D = True
    TChart1.Aspect.Perspective = False
    Debug.Print "TC-AR-VIEW=" & TChart1.Aspect.View3D
    Debug.Print "TC-AR-TYPE=" & TChart1.Series(0).SeriesType
    Debug.Print "TC-AR-READY"
End Sub
