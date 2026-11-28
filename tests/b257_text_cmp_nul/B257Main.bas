Option Explicit
Option Compare Text

' §11.234 (第二百五十二刀): Option Compare Text 档的字符串比较口径。
' 夹具 = 真机探针 .temp/k234p/P368 + P369 的同一批语句（Print #f 换成 Debug.Print）。
' 文本档与 Binary 档**不同**的两件事：
'   1) 内嵌 NUL 被剔掉再比（`"a" = "a" & Chr$(0)` True），不是当普通字符；
'   2) 一侧长度 0 而另一侧非 0 一律不等（`Chr$(0) = ""` False），连剔完相等也不认。
' 相同的一件事：空格不补，`"a" = "a "` 仍是 False。
Public Sub Main()
    Dim z As String
    z = Chr(0)

    ' --- P368: 文本档主读数
    Debug.Print "T1|" & CStr("a" = "A")
    Debug.Print "T2|" & CStr("a" = "a ")
    Debug.Print "T3|" & CStr("a" = "a" & z)
    Debug.Print "T4|" & CStr("" = z)
    Debug.Print "T5|" & CStr("a" & z = "A")
    Debug.Print "T6|" & CStr("a" < "a ")
    Debug.Print "T7|" & CStr(StrComp("a", "A", 1))
    Debug.Print "T8|" & CStr(StrComp("a", "a ", 1))
    Debug.Print "T9|" & CStr(StrComp("a", "a" & z, 1))
    Debug.Print "T10|" & CStr(Len(z))
    Debug.Print "T11|" & CStr("ABC" < "abd")
    Debug.Print "T12|" & CStr("a" & z < "A")
    Debug.Print "T13|" & CStr(z = " ")

    Dim s As String
    s = "a" & z
    Select Case s
        Case "A"
            Debug.Print "T14|varNul=HIT"
        Case Else
            Debug.Print "T14|varNul=MISS"
    End Select

    s = "a "
    Select Case s
        Case "A"
            Debug.Print "T15|varPad=HIT"
        Case Else
            Debug.Print "T15|varPad=MISS"
    End Select

    ' --- P369: "空串 vs 带 NUL 的串"到底按什么判
    Debug.Print "U1|" & CStr("" = "")
    Debug.Print "U2|" & CStr(z = "")
    Debug.Print "U3|" & CStr("a" = z)
    Debug.Print "U4|" & CStr("a" = z & "a")
    Debug.Print "U5|" & CStr("ab" = "a" & z)
    Debug.Print "U6|" & CStr("ab" = "a" & z & "b")
    Debug.Print "U7|" & CStr("" = z & z)
    Debug.Print "U8|" & CStr(StrComp("", z, 1))
    Debug.Print "U9|" & CStr(StrComp(z, "", 1))
    Debug.Print "U10|" & CStr("a" < z)
    Debug.Print "U11|" & CStr("A" = "a")
    Debug.Print "U12|" & CStr(Len("a" & z))
    Debug.Print "U13|" & CStr("" = " ")
    Debug.Print "U14|" & CStr(" " = "")
End Sub