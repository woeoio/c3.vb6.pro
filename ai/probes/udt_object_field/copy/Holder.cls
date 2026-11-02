VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Holder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 与上一级目录那份唯一的不同：计数与打印走 CopyMain（这一份工程里没有名为 Main 的模块）。
Public Tag As Long

Private Sub Class_Terminate()
    CopyMain.gTerm = CopyMain.gTerm + 1
    CopyMain.Say "TERM " & Me.Tag
End Sub
