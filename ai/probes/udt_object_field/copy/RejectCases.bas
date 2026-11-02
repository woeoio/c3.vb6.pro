Attribute VB_Name = "RejectCases"
Option Explicit

' 这四格是**故意可能不许编译**的形状。默认没有加进 UdtCopyProbe.vbp —— 一次加一格，编译，
' 把 IDE 的原话贴在对应编号上（「不许编译」本身就是 C3 要的答案）；能编译的那格把输出贴回来。
'
'   R1  UDT 形参能不能 ByVal
'   R2  Function 能不能返回 UDT
'   R3  ReDim Preserve 缩小含对象字段的 UDT 数组时，被丢掉那几格的对象归谁放（这格是运行期读数）
'   R4  Let <对象字段> = <对象> 是哪一条错误
'
' 用法：VB6 里 Project → Add File 选这份，编译；报错就把那句话原样抄下来，然后按编号继续。
' （R1/R2/R4 各是一句独立声明，可以单独复制进一个空模块逐条试。）

Private Type RObject
    O As Object
End Type

' ---- R1 ----
Private Sub TakeByVal(ByVal u As RObject)
    u.O = Nothing
End Sub

Public Sub RunR1()
    Dim u As RObject
    Set u.O = CopyMain.NewH()
    TakeByVal u
    CopyMain.Say "R1-done"
End Sub

' ---- R2 ----
Public Function GiveUDT() As RObject
    Dim u As RObject
    Set u.O = CopyMain.NewH()
    GiveUDT = u
End Function

Public Sub RunR2()
    Dim v As RObject
    v = GiveUDT()
    CopyMain.Say "R2-done"
End Sub

' ---- R3（能编译才有运行期读数）----
Public Sub RunR3()
    Dim arr() As RObject
    ReDim arr(3)
    Dim i As Long
    For i = 0 To 3
        Set arr(i).O = CopyMain.NewH()
    Next i
    CopyMain.Say "R3-before-shrink"
    ReDim Preserve arr(1)
    CopyMain.Say "R4-after-shrink zero-alive=" & (Not (arr(0).O Is Nothing))
End Sub

' ---- R4 ----
Public Sub RunR4()
    Dim u As RObject
    Dim h As Holder
    Set h = CopyMain.NewH()
    Let u.O = h
    CopyMain.Say "R5-done"
End Sub
