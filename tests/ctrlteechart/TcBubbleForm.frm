VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart bubble"
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
' C30-10 气泡图判据夹具：401x257，AddSeries 11（stBubble），三枚气泡
'   AddBubble x, y, r = (1,30,5) (3,25,45) (5,20,20)。最大那枚放在**中间**：
'   放在最右时圆心就靠 plot.right，圆会被扫描窗口截掉，量出来的直径偏小（踩过）。
' 判据（tc_type_px.ps1 -Mode bubble）：气泡色像素够多、最大圆直径 ≈ 24±6。
Private Sub Form_Load()
    TChart1.AddSeries 11
    TChart1.Series(0).Title = "Series0"
    TChart1.Series(0).Marks.Visible = False
    TChart1.Series(0).Color = RGB(0, 128, 255)
    TChart1.Series(0).AddBubble 1, 30, 5
    TChart1.Series(0).AddBubble 3, 25, 45
    TChart1.Series(0).AddBubble 5, 20, 20
    TChart1.Aspect.View3D = False
    Debug.Print "TC-BB-TYPE=" & TChart1.Series(0).SeriesType
    Debug.Print "TC-BB-COUNT=" & TChart1.Series(0).Count
    Debug.Print "TC-BB-READY"
End Sub
