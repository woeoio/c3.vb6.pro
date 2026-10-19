VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C445"
Option Explicit
Public N As Long

' 两枚计数器都记在 Main 那张全局上: Made = 出厂过的创建次数, Terms = 走完 Class_Terminate 的出厂次数。
' 这一对就是 (甲) 那把行为尺: 每格窗口先记下起点, 再打差值, 免得各格互相把对方的数算进自己头上。
Private Sub Class_Initialize()
    Main.Made = Main.Made + 1
    N = 7
End Sub

Private Sub Class_Terminate()
    Main.Terms = Main.Terms + 1
End Sub
