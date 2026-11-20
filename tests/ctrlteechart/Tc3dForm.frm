VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart 3d"
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
' C30-10 3D 判据夹具：客户区 401x257，一条柱三层值 10/20/30，色固定便于按像素认。
' 判据（tests\ctrlteechart\tc_3d_px.ps1）量两件事：
'   侧面比正面右缘多出 12px（深度常量）、顶面比柱顶高 12px（方向朝上）。
Private Sub Form_Load()
    TChart1.AddSeries 1
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddXY 1, 10
    TChart1.Series(0).AddXY 2, 20
    TChart1.Series(0).AddXY 3, 30
    TChart1.Aspect.View3D = True
    TChart1.Aspect.Perspective = False
    Debug.Print "TC-3D-VIEW=" & TChart1.Aspect.View3D
    Debug.Print "TC-3D-PERSP=" & TChart1.Aspect.Perspective
    Debug.Print "TC-3D-ROT=" & TChart1.Aspect.Rotation
    Debug.Print "TC-3D-ELEV=" & TChart1.Aspect.Elevation
    Debug.Print "TC-3D-READY"
End Sub
