VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart hbar"
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
' C30-10 横向柱 + 切面判据夹具：401x257，AddSeries 2（stHorizBar），**四个值**：
' 20/30/25/22。为什么四值：值等于 vmin 的那根柱长度为 0（与竖柱同构），
' 三值时只会有两根可量的横条，判据没法区分“少画了一根”和“零长柱”。
' 判据（tc_hbar_px.ps1）：
'   ① 柱从左基线长出去：最左 = plot.left(28)、最长那根到 plot.right(302)
'   ② 柱是“横条”：在 x=40 竖着扫，柱色只出现在 3 段窄带里
'   ③ BarWidthPercent=40 真改厚度：每段厚度 ≈ band*40/200 ≈ 10px
Private Sub Form_Load()
    TChart1.AddSeries 2
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddXY 1, 20
    TChart1.Series(0).AddXY 2, 30
    TChart1.Series(0).AddXY 3, 25
    TChart1.Series(0).AddXY 4, 22
    TChart1.Aspect.View3D = False
    TChart1.Series(0).AsHorizBar.BarWidthPercent = 40
    Debug.Print "TC-HB-TYPE=" & TChart1.Series(0).SeriesType
    Debug.Print "TC-HB-PCT=" & TChart1.Series(0).AsHorizBar.BarWidthPercent
    Debug.Print "TC-HB-ASBAR=" & TChart1.Series(0).AsBar.BarWidthPercent
    Debug.Print "TC-HB-CANDLE=" & TChart1.Series(0).AsCandle.CandleWidth
    Debug.Print "TC-HB-SURF=" & TChart1.Series(0).AsSurface.WireFrame
    Debug.Print "TC-HB-LINE=" & TChart1.Series(0).AsLine.Smoothed
    Debug.Print "TC-HB-READY"
End Sub
