Option Explicit

' §11.224 (第二百四十二刀): vbNullString 的**空指针语义**沿赋值/返回/形参/变体槽保真。
' 判据是真机 .temp/k232p/P348（VB6Mini）。这里只钉「StrPtr 是不是 0」（写成 True/False），
' 因为地址本身每次不同；反方向那几格（Left/Space/Trim 给的是真空串）一起钉住，防止有人
' 把「空」与「空指针」当同一件事。
' 规则一句话：NULL 只从 vbNullString 那一族（赋值、函数返回、String 形参透传、变体槽、
' Error(0)）流过去；字符串函数算出来的空串是**真 BSTR**。
Private Function RetNull() As String
    RetNull = vbNullString
End Function

Private Function RetEmpty() As String
    RetEmpty = ""
End Function

Private Function PassThrough(s As String) As String
    PassThrough = s
End Function

Public Sub Main()
    Dim s As String, t As String

    s = vbNullString
    Debug.Print "N1=" & CStr(StrPtr(s) = 0) & "/" & CStr(Len(s)) & "/" & CStr(LenB(s))

    s = ""
    Debug.Print "N2=" & CStr(StrPtr(s) = 0)

    s = "abc"
    s = vbNullString
    Debug.Print "N3=" & CStr(StrPtr(s) = 0) & "/" & CStr(Len(s))

    s = vbNullString
    t = "x" & s
    Debug.Print "N4=[" & t & "]/" & CStr(StrPtr(t) = 0)

    s = vbNullString
    Debug.Print "N5=" & CStr(s = "") & "/" & CStr(Len(s & "y"))

    Debug.Print "N6=" & CStr(StrPtr(RetNull()) = 0) & "/" & CStr(StrPtr(RetEmpty()) = 0)

    Debug.Print "N7=" & CStr(StrPtr(PassThrough(vbNullString)) = 0) & "/" & CStr(StrPtr(PassThrough("")) = 0)

    s = RetNull()
    Debug.Print "N8=" & CStr(StrPtr(s) = 0)

    Dim v As Variant
    v = vbNullString
    Debug.Print "N9=" & CStr(VarType(v)) & "/" & CStr(StrPtr(CStr(v)) = 0)
    v = ""
    Debug.Print "N10=" & CStr(VarType(v)) & "/" & CStr(StrPtr(CStr(v)) = 0)

    s = Error(0)
    Debug.Print "N11=" & CStr(StrPtr(s) = 0)

    ' 反方向：字符串函数给的是真空串，不是空指针
    s = Left("abc", 0)
    Debug.Print "N12=" & CStr(StrPtr(s) = 0) & "/" & CStr(Len(s))
    s = Mid("abc", 2, 0)
    Debug.Print "N13=" & CStr(StrPtr(s) = 0)
    s = Space(0)
    Debug.Print "N14=" & CStr(StrPtr(s) = 0)
    s = Trim("")
    Debug.Print "N15=" & CStr(StrPtr(s) = 0)
End Sub