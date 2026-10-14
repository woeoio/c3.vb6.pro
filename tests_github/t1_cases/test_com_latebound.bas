' P24-10b: 后期绑定COM默认属性调用 obj(args) -> vb6_ComCallByDispid(obj, 0, ...)
' 与 test_com_default_prop 互补: 那是前期绑定 (Dim d As Dictionary, CLSID 签名路径,
' 按 comDefaultMemberRealName 调 vb6_ComCall); 这里是 Dim o As Object (无类型库签名),
' 按 VB6 语义直接以 DISPID_VALUE=0 调默认成员。

Option Explicit

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

    Debug.Print "P24-10b: "; passCount; "/5 passed"
End Sub