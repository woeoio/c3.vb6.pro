Option Explicit

' 账 #417 (2026-11-07): URL 图片加载链的最小夹具。
' 链: CreateObject("Microsoft.XMLHTTP") -> .Open/.send -> readyState=4 -> Status=200
'     -> .responseBody (VT_ARRAY|VT_UI1) -> Byte() 赋值。
' 用法: UrlPicMain.exe <url>   (缺省 http://127.0.0.1:8765/probe.png)
' 每一站都打印返回值; 断言在 tests/run_tests.ps1 的 urlpic_local / urlpic_net。

Private Function SafeUbound(ByRef a() As Byte) As Long
    On Error GoTo bad
    SafeUbound = UBound(a)
    Exit Function
bad:
    SafeUbound = -999
End Function

Private Function SafeLbound(ByRef a() As Byte) As Long
    On Error GoTo bad
    SafeLbound = LBound(a)
    Exit Function
bad:
    SafeLbound = -999
End Function

Private Function SafeElem(ByRef a() As Byte, ByVal i As Long) As Long
    On Error GoTo bad
    SafeElem = CLng(a(i))
    Exit Function
bad:
    SafeElem = -999
End Function

Private Function SafeVarType(ByRef v As Variant) As Long
    On Error GoTo bad
    SafeVarType = CLng(VarType(v))
    Exit Function
bad:
    SafeVarType = -999
End Function

Private Function SafeVarElem(ByRef v As Variant, ByVal i As Long) As Long
    On Error GoTo bad
    SafeVarElem = v(i)
    Exit Function
bad:
    SafeVarElem = -999
End Function

Sub Main()
    Dim url As String
    Dim o As Object
    Dim polls As Long
    Dim objOk As Long

    url = Command$
    If Len(url) = 0 Then url = "http://127.0.0.1:8765/probe.png"
    Debug.Print "URL=" & url

    On Error Resume Next

    Set o = CreateObject("Microsoft.XMLHTTP")
    objOk = 0
    If Not (o Is Nothing) Then objOk = 1
    Debug.Print "S1 create obj=" & objOk & " err=" & Err.Number
    Err.Clear

    If objOk = 0 Then
        Debug.Print "S2..S8 SKIP (create failed)"
        Exit Sub
    End If

    o.Open "GET", url, True
    Debug.Print "S2 open err=" & Err.Number
    Err.Clear

    o.send
    Debug.Print "S3 send err=" & Err.Number
    Err.Clear

    Do While o.readyState <> 4
        DoEvents
        polls = polls + 1
        If polls > 5000 Then Exit Do
    Loop
    Debug.Print "S4 readyState=" & o.readyState & " polls=" & polls & " err=" & Err.Number
    Err.Clear

    Debug.Print "S5 status=" & o.Status & " err=" & Err.Number
    Err.Clear

    Dim v As Variant
    v = o.responseBody
    Debug.Print "S6a responseBody VarType=" & SafeVarType(v) & " e0=" & SafeVarElem(v, 0) & " e1=" & SafeVarElem(v, 1) & " err=" & Err.Number
    Err.Clear

    Dim b() As Byte
    b() = o.responseBody
    Debug.Print "S6b assign err=" & Err.Number & " lb=" & SafeLbound(b) & " ub=" & SafeUbound(b)
    Err.Clear

    Debug.Print "S7 b0=" & SafeElem(b, 0) & " b1=" & SafeElem(b, 1) & " b2=" & SafeElem(b, 2) & " b3=" & SafeElem(b, 3)
    Debug.Print "S8 DONE"
End Sub
