' P24-10b: 后期绑定COM默认属性调用 obj(args) -> vb6_ComCallByDispid(obj, 0, ...)
' 与 test_com_default_prop 互补: 那是前期绑定 (Dim d As Dictionary, CLSID 签名路径,
' 按 comDefaultMemberRealName 调 vb6_ComCall); 这里是 Dim o As Object (无类型库签名),
' 按 VB6 语义直接以 DISPID_VALUE=0 调默认成员。

Option Explicit

' ---- vbeclipse knife 50 (case #305 B159): module-level arrays <-> real COM SAFEARRAY.
' Values sit at the index each case reads back; b159Mu() is never ReDim-ed on purpose.
Private b159L() As Long
Private b159B() As Byte
Private b159X(1 To 3) As Long
Private b159St() As String
Private b159Bl() As Boolean
Private b159Sg() As Single
Private b159Mu() As Long

Sub Main()
    Dim passCount As Long
    passCount = 0

    ' 测试1: 基础读 - Dim o As Object (后期绑定)
    ' o("name") 应等价于 o.Item("name") (DISPID_VALUE)
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "name", "Alice"
    o.Add "age", 30

    Dim s1 As String
    s1 = o("name")
    If s1 = "Alice" Then
        passCount = passCount + 1
        Debug.Print "LB-1:OK o(""name"") = Alice"
    Else
        Debug.Print "LB-1:FAIL o(""name"") expected Alice, got "; s1
    End If

    ' 测试2: 数值类型读 (Variant = 30)
    Dim n1 As Long
    n1 = o("age")
    If n1 = 30 Then
        passCount = passCount + 1
        Debug.Print "LB-2:OK o(""age"") = 30"
    Else
        Debug.Print "LB-2:FAIL o(""age"") expected 30, got "; n1
    End If

    ' 测试3: 二进制比较 (If o("k") = "v" Then) - 结果进二进制路径, 需先解包
    If o("name") = "Alice" Then
        passCount = passCount + 1
        Debug.Print "LB-3:OK o(""name"") == ""Alice"" in If"
    Else
        Debug.Print "LB-3:FAIL o(""name"") != Alice"
    End If

    ' 测试4: 字符串拼接 (字符串上下文)
    Dim s2 As String
    s2 = "v=" & o("name")
    If s2 = "v=Alice" Then
        passCount = passCount + 1
        Debug.Print "LB-4:OK concat v=Alice"
    Else
        Debug.Print "LB-4:FAIL concat got "; s2
    End If

    ' 测试5: 传入字符串参数 (Len -> BSTR 形参)
    Dim n2 As Long
    n2 = Len(o("name"))
    If n2 = 5 Then
        passCount = passCount + 1
        Debug.Print "LB-5:OK Len(o(""name"")) = 5"
    Else
        Debug.Print "LB-5:FAIL Len expected 5, got "; n2
    End If


    ' ---- knife 50: write an array through IDispatch, read it back, ask what came ----
    ReDim b159L(0 To 3)
    b159L(0) = 7002
    ReDim b159B(0 To 2)
    b159B(1) = 66
    b159X(2) = 7022
    ReDim b159St(0 To 1)
    b159St(0) = "s0"
    ReDim b159Bl(0 To 1)
    b159Bl(0) = True
    ReDim b159Sg(0 To 1)
    b159Sg(0) = 1.25
    Dim b159Loc() As Long
    ReDim b159Loc(0 To 1)
    b159Loc(0) = 7100
    Dim b159Va As Variant
    b159Va = b159L
    o.Add "ml", b159L
    B159Show "ml", o, 0
    o.Add "mb", b159B
    B159Show "mb", o, 1
    o.Add "mx", b159X
    B159Show "mx", o, 2
    o.Add "bl", b159Bl
    B159Show "bl", o, 0
    o.Add "sg", b159Sg
    B159Show "sg", o, 0
    o.Add "loc", b159Loc
    B159Show "loc", o, 0
    o.Add "mu", b159Mu
    B159Show "mu", o, 0
    o.Add "ms", b159St
    B159Show "ms", o, 0
    o.Add "va", b159Va
    B159Show "va", o, 0
    On Error GoTo 0

    Debug.Print "P24-10b: "; passCount; "/5 passed"
End Sub

' Reads one key back into a Variant slot and prints vt/type/UBound/element.
' UBound and the element read are error-guarded: the deliberate negative cases answer
' Empty there (VB6 raises 9 as well), and a bare error would abort the whole fixture.
Private Sub B159Show(tag As String, d As Object, idx As Long)
    Dim r As Variant
    r = d(tag)
    Dim s As String
    s = "VB-b159-" & tag & "=" & CStr(VarType(r)) & "/" & TypeName(r)
    On Error Resume Next
    s = s & "/" & CStr(UBound(r))
    If Err.Number <> 0 Then
        s = s & "(e" & CStr(Err.Number) & ")"
        Err.Clear
    End If
    s = s & "/" & CStr(r(idx))
    If Err.Number <> 0 Then
        s = s & "(e" & CStr(Err.Number) & ")"
        Err.Clear
    End If
    Debug.Print s
End Sub
