Option Explicit

' 账 #305 §11.155 (第一百七十二刀): String 与数值之间的隐式换算。
'
' 赋值站 (以前是裸强转):
'   B175-1 = Long   ← "12"      (vb6_CLng, BSTR 臂 → vb6_NumVal 那位主人)
'   B175-2 = Integer← "34"      (vb6_CInt)
'   B175-3 = Byte   ← "200"     (Byte 没有 BSTR 臂 ⇒ vb6_CByte(vb6_CDbl(..)))
'   B175-4 = Double ← "1.5"
' 实参站 (属性读当 ByRef 实参, 数值属性 → String 形参):
'   B175-5 = Long 属性 7  → "7" & "0" = "70" ⇒ 写回 70
'   B175-6 = Double 属性 1.5 → "1.5" & "5" = "1.55" ⇒ 写回 1.55
' 对照 (同档, 一字不改):
'   B175-7 = Long 属性 → Long 形参
'   B175-8 = Double 属性 → Double 形参
'
' 在册 (本刀不碰): String 属性 → 数值形参仍是按值 (要两枚临时量 + 写回那枚 BSTR),
' 见 §11.155 欠项清单。
Public Sub PushNum(ByRef n As Long)
    n = n + 5
End Sub

Public Sub DblTwice(ByRef d As Double)
    d = d * 2
End Sub

Public Sub AppendZero(ByRef s As String)
    s = s & "0"
End Sub

Public Sub AppendFive(ByRef s As String)
    s = s & "5"
End Sub

Public Sub Main()
    Dim n As Long
    n = "12"
    Debug.Print "B175-1=" & n

    Dim i As Integer
    i = "34"
    Debug.Print "B175-2=" & i

    Dim by As Byte
    by = "200"
    Debug.Print "B175-3=" & by

    Dim d As Double
    d = "1.5"
    Debug.Print "B175-4=" & d

    Dim bx As B175Box
    Set bx = New B175Box

    bx.N = 7
    AppendZero bx.N
    Debug.Print "B175-5=" & bx.N

    bx.D = 1.5
    AppendFive bx.D
    Debug.Print "B175-6=" & bx.D

    bx.N = 10
    PushNum bx.N
    Debug.Print "B175-7=" & bx.N

    bx.D = 2
    DblTwice bx.D
    Debug.Print "B175-8=" & bx.D

    Debug.Print "B175-DONE"
End Sub
