VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart persp"
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
' C30-10 透视判据夹具：与 Tc3dApp 同几何同数据，只把 Perspective 打开。
' 判据（tests\ctrlteechart\tc_persp_px.ps1）量三根柱子的侧面色带宽：
' 左边（远端）应比右边（近端）厚 —— 这就是“透视”与“平行投影”的像素差别。
Private Sub Form_Load()
    TChart1.AddSeries 1
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    ' 三根柱子等高：值等于 vmin 的柱子高度为 0、根本没侧面，量不到带宽（踩过）。
    TChart1.Series(0).AddXY 1, 20
    TChart1.Series(0).AddXY 2, 20
    TChart1.Series(0).AddXY 3, 20
    TChart1.Aspect.View3D = True
    TChart1.Aspect.Perspective = True
    Debug.Print "TC-PS-VIEW=" & TChart1.Aspect.View3D
    Debug.Print "TC-PS-PERSP=" & TChart1.Aspect.Perspective
    Debug.Print "TC-PS-READY"
End Sub
