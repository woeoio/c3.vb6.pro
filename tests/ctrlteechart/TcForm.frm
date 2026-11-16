VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart builtin"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4500
   Begin TeeChart.TChart TChart1 
      Height          =   2400
      Left            =   120
      OleObjectBlob   =   "TcBlob.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   60
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' C30-10 夹具：内置 TChart 控件（不加载 TeeChart5.ocx）。
' 本刀只验「控件建得出窗口 + 画得出一遍」：绘制次数由 RTL 的 C3_TCHART_TRACE 打点。
Private Sub Form_Load()
    Dim s As Object, nF As Long
    Debug.Print "TC-BUILTIN-START"
    Debug.Print "TC-WIDTH=" & TChart1.Width
    ' TPF0（TcBlob.frx）里的设计期外观：两条系列 + 标题，Form_Load 之前就该在
    Debug.Print "TC-HEADER0=" & TChart1.Header.Caption
    Debug.Print "TC-HEADER0LEN=" & Len(TChart1.Header.Caption)
    Debug.Print "TC-DESIGNSER=" & TChart1.SeriesCount
    Debug.Print "TC-SER0TYPE=" & TChart1.Series(0).SeriesType
    Debug.Print "TC-SER0COLOR=" & TChart1.Series(0).Color
    Debug.Print "TC-SER1COLOR=" & TChart1.Series(1).Color
    TChart1.AddSeries 5
    TChart1.Series(0).FillSampleValues 4
    Debug.Print "TC-SERIES=" & TChart1.SeriesCount
    nF = 0
    For Each s In TChart1.Series
        nF = nF + 1
    Next s
    Debug.Print "TC-FOREACH=" & nF
    TChart1.Series(0).Marks.Visible = False
    Debug.Print "TC-MARKS=" & TChart1.Series(0).Marks.Visible
    TChart1.Series(0).Marks.Visible = True
    Debug.Print "TC-MARKS2=" & TChart1.Series(0).Marks.Visible
    TChart1.Series(0).Pen.Color = RGB(0, 0, 255)
    Debug.Print "TC-PEN=" & TChart1.Series(0).Pen.Color
    TChart1.Series(0).Brush.Color = RGB(1, 2, 3)
    Debug.Print "TC-BRUSH=" & TChart1.Series(0).Brush.Color
    TChart1.Series(2).AsPie.ExplodedSlice.Value(0) = 15
    Debug.Print "TC-SLICE=" & TChart1.Series(2).AsPie.ExplodedSlice.Value(0)
    Debug.Print "TC-ROT=" & TChart1.Series(2).AsPie.RotationAngle
    ' 文档 §5 那四条未接成员：显式空实现（认名字、不动作，不许 438）
    TChart1.ShowEditor
    Debug.Print "TC-ED=ok"
    TChart1.Printer.BeginDoc
    TChart1.Printer.EndDoc
    Debug.Print "TC-PR=" & TChart1.Printer.PageWidth
    TChart1.Export.SaveToFile "tc-export.png"
    Debug.Print "TC-EX=ok"
    TChart1.Zoom.Enable = False
    Debug.Print "TC-ZM=" & TChart1.Zoom.Enable
    TChart1.Panel.Color = RGB(255, 240, 240)
    Debug.Print "TC-PANEL=" & TChart1.Panel.Color
    TChart1.Header.Caption = "内置图表"
    Debug.Print "TC-HEADER=" & TChart1.Header.Caption
    Debug.Print "TC-HEADERLEN=" & Len(TChart1.Header.Caption)
    TChart1.Axis.Bottom.Title.Caption = "月份"
    Debug.Print "TC-AXISTITLE=" & TChart1.Axis.Bottom.Title.Caption
    Debug.Print "TC-AXTITLELEN=" & Len(TChart1.Axis.Bottom.Title.Caption)
    Debug.Print "TC-MARGINR=" & TChart1.Panel.MarginRight
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
