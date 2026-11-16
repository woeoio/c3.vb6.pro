VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart point3d"
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
' C30-10 点系列 3D 判据夹具：401x257，一条点系列三值 20/30/25，色固定。
' 判据（tests\ctrlteechart\tc_point3d_px.ps1）量“暗色偏移圆”相对正面圆的位置：
'   后墩上缘比正面高 4px、最高行的起点右移 4px —— 就是扁墩的厚度。
Private Sub Form_Load()
    TChart1.AddSeries 4
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddXY 1, 20
    TChart1.Series(0).AddXY 2, 30
    TChart1.Series(0).AddXY 3, 25
    TChart1.Aspect.View3D = True
    TChart1.Aspect.Perspective = False
    Debug.Print "TC-PT-VIEW=" & TChart1.Aspect.View3D
    Debug.Print "TC-PT-TYPE=" & TChart1.Series(0).SeriesType
    Debug.Print "TC-PT-READY"
End Sub
