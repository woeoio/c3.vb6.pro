VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart pie"
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
' C30-10 像素判据夹具（饼图掰开）：图表客户区同样 401x257。
' 四等份扇区 ⇒ 每块 90°，第 0 块从 -90° 扫到 0°（右上象限），中线 -45°。
' 只掰第 0 块 18px：判据就看它比其余三块的可见半径多出约 18px。
Private Sub Form_Load()
    Dim i As Long
    TChart1.AddSeries 5
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    For i = 1 To 4
        TChart1.Series(0).AddXY i, 25
    Next i
    TChart1.Series(0).AsPie.ExplodedSlice.Value(0) = 18
    Debug.Print "TP-EXPLODE=" & TChart1.Series(0).AsPie.ExplodedSlice.Value(0)
    Debug.Print "TP-ROT=" & TChart1.Series(0).AsPie.RotationAngle
    Debug.Print "TP-READY"
End Sub
