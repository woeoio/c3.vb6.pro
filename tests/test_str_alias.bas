' 账 #113: String **形参/变量**存进模块变量必须是拷贝，不是指针别名。
' 修复前那句 `gS = s;` 存的是调用方那只 BSTR 的地址，调用方一改写自己的变量
' （vb6_BSTR_Assign 释放旧块）模块变量就跟着悬垂 —— WS-c 的 Error 事件里读出过
' 别人用过的堆块（"WS38=" / "Y"）。
' 本用例的问法都是**当场问值**，而且先制造堆扰动：不扰动的话，那块已释放的内存
' 往往还留着原字节，悬垂也照样"读对"，判据就是假绿。
Option Explicit

Dim gS As String

Sub Take(ByVal s As String)
    gS = s
End Sub

Function MakeIt() As String
    MakeIt = "FROM-FUNC"
End Function

Sub Churn(n As Long)
    Dim i As Long
    Dim junk As String
    junk = ""
    For i = 1 To n
        junk = junk & "0123456789ABCDEF"
        If Len(junk) > 64 Then junk = "x"
    Next i
    If Len(junk) < 0 Then Debug.Print "never"
End Sub

Sub Main()
    Dim l As String
    Dim l2 As String

    l = "PAYLOAD-0123456789"
    Take l                       ' 形参 -> 模块变量
    l = "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"   ' 释放掉 gS 还指着的那块
    Churn 300
    Debug.Print "S113-A=" & gS
    Debug.Print "S113-B=" & Len(gS)

    l2 = "SECOND-ONE"
    gS = l2                      ' 局部变量 -> 模块变量
    l2 = "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
    Churn 300
    Debug.Print "S113-C=" & gS
    gS = l2
    Debug.Print "S113-D=" & Len(gS)

    gS = "LITERAL-OK"
    Debug.Print "S113-E=" & gS
    gS = "A" & "B" & "C"
    Debug.Print "S113-F=" & gS
    gS = MakeIt()
    Debug.Print "S113-G=" & gS
    Debug.Print "S113-DONE"
End Sub
