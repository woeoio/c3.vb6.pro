VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart fx"
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
' C30-10 第二十刀：勾选框那批属性的**绘制**判据夹具（一）—— 渐变 / 外框 / 四壁 / 2D。
' 这些就是 ChartDisplay 示例里 Check1(渐变) / Check3(边框) / Check5(3D) / Check6(Frame)
' / Check11(Walls) 控制的成员，以前只存不画，勾了没反应。
Private Sub Form_Load()
    Dim i As Integer
    TChart1.AddSeries 1
    TChart1.Series(0).Title = "S0"
    For i = 1 To 5
        TChart1.Series(0).Add i, i * 2
    Next i
    ' 渐变：红 → 蓝（判据就看上下两端的颜色）
    TChart1.Panel.Gradient.Visible = True
    TChart1.Panel.Gradient.StartColor = vbRed
    TChart1.Panel.Gradient.EndColor = vbBlue
    ' 外框与 Panel 边框都关掉（Check6 / Check3）
    TChart1.Frame.Visible = False
    TChart1.Panel.BorderStyle = 0
    ' 四壁关掉 + 回到 2D（Check11 / Check5）
    TChart1.Walls.Visible = False
    TChart1.Aspect.View3D = False
    Debug.Print "TC-FX-READY"
End Sub
