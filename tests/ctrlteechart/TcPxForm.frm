VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart px"
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
' C30-10 像素判据夹具：图表客户区 401x257（与 tests\TeeChartChartDisplay 的真 VB6 参照同几何）。
' 内容刻意与参照对齐：同一个标题文字、一条柱一条线、图例在右。
' Marks 全关 —— 逐点标签会落进判据的扫描带里，把轴线点数搅脏。
Private Sub Form_Load()
    Dim i As Long
    TChart1.Header.Caption = "TeeChart display demo"
    TChart1.AddSeries 1
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.AddSeries 0
    TChart1.Series(1).Title = "Series1"
    TChart1.Series(1).Marks.Visible = False
    For i = 1 To 10
        TChart1.Series(0).AddXY i, (i * 7) Mod 23
    Next i
    For i = 1 To 10
        TChart1.Series(1).AddXY i, 12
    Next i
    Debug.Print "TXP-READY"
End Sub
