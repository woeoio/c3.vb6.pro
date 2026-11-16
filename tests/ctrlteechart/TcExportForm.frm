VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart export"
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
' C30-10 导出判据夹具：401x257，一条柱三值（20/30/25），View3D=False（色块干净）。
' Form_Load 里真导出 JPEG + 真拷剪贴板；判据（tc_export_px.ps1）回头看文件与剪贴板。
Private Sub Form_Load()
    TChart1.AddSeries 1
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddXY 1, 20
    TChart1.Series(0).AddXY 2, 30
    TChart1.Series(0).AddXY 3, 25
    TChart1.Aspect.View3D = False
    TChart1.Export.SaveToJPEGFile "tc_export.jpg"
    TChart1.Export.CopyToClipboard
    Debug.Print "TC-EXP-DONE"
End Sub
