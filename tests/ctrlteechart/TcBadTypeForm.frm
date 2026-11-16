VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TChart badtype"
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
' C30-10 不支持类型的报错判据夹具：AddSeries 8（Candle，未支持）必须抛错。
' 以前它静默降级成折线 —— 不报错但图是错的，比报错危险。
' 判据（tc_type_px.ps1 -Mode badtype）：Err.Number 必须是 5，且 Count 仍是 0
' （报错后不许悄悄建出一条系列）。
Private Sub Form_Load()
    On Error Resume Next
    TChart1.AddSeries 8
    Debug.Print "TC-BAD-ERR=" & Err.Number
    Err.Clear
    Debug.Print "TC-BAD-COUNT=" & TChart1.Series.Count
    TChart1.AddSeries 17
    Debug.Print "TC-BAD-OK=" & TChart1.Series(0).SeriesType
    Debug.Print "TC-BAD-READY"
End Sub
