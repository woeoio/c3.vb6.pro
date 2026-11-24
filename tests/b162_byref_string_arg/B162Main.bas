Option Explicit

Public Sub Main()
    Dim box As B162Box
    Dim o As Object
    Dim s As String
    Dim l As Long

    Set box = New B162Box
    Set o = box

    ' S1: ByRef String，体里 x = x & "-ok"
    s = "hi"
    o.Bump s
    Debug.Print "B162-S1=" & Len(s) & "/" & s

    ' S2: ByRef String，体里整个换掉
    s = "x"
    o.Swap s
    Debug.Print "B162-S2=" & Len(s) & "/" & s

    ' S3: ByRef String，体里清空
    s = "abc"
    o.Clear s
    Debug.Print "B162-S3=" & Len(s) & "/[" & s & "]"

    ' S4: ByVal String 对照 —— 调用方必须原样
    s = "keep"
    o.BumpVal s
    Debug.Print "B162-S4=" & Len(s) & "/" & s

    ' S5: 体里 x = x（不改）—— 走一遍入槽/出站，串必须还是好的
    s = "loop"
    o.NoChange s
    Debug.Print "B162-S5=" & Len(s) & "/" & s

    ' S6: 空串进，ByRef 出来带上尾巴
    s = ""
    o.Bump s
    Debug.Print "B162-S6=" & Len(s) & "/" & s

    ' S7: 数值那六档同一夹具里对照，必须一字不变
    l = 41
    o.BumpL l
    Debug.Print "B162-S7=" & l

    Debug.Print "B162-DONE"
    Set box = Nothing
    Set o = Nothing
End Sub
