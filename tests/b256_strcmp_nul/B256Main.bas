Option Explicit

' §11.233 (第二百五十一刀): 字符串二元比较与 Select Case 的 VB6 口径 ——
' 按两侧真实长度逐字符比（不补空格）、内嵌 NUL 是普通字符（不是串尾）。
' 夹具 = 真机探针 .temp/k233p/P364b + P365 的同一批语句（Print #f 换成 Debug.Print）。
Public Sub Main()
    Dim s As String
    Dim z As String

    z = Chr(0)

    Debug.Print "C1|" & CStr("a" = "a ")
    Debug.Print "C2|" & CStr("a " = "a")
    Debug.Print "C3|" & CStr("a" = "a" & Chr(0))
    Debug.Print "C4|" & CStr("a" & Chr(0) = "a")
    Debug.Print "C5|" & CStr("" = Chr(0))
    Debug.Print "C6|" & CStr(Len(Chr(0)))
    Debug.Print "C7|" & CStr("a" < "a ")
    Debug.Print "C8|" & CStr("a" & Chr(0) < "a")
    Debug.Print "C9|" & CStr("abc" > "ab")
    Debug.Print "C10|" & CStr(Chr(0) & "a" = "a")
    Debug.Print "C11|" & CStr(StrComp("a", "a ", 0))
    Debug.Print "C12|" & CStr("A" = "a")
    Debug.Print "C13|" & CStr("a" & Chr(0) & Chr(0) = "")
    Debug.Print "C14|" & CStr(Mid$("abc", 1, 0) = "")
    Debug.Print "C15|" & CStr(Chr(0) = Chr(0))
    Debug.Print "C16|" & CStr("   " = "")
    Debug.Print "C17|" & CStr(Chr(0) = " ")

    ' 变量形式（不是两侧都折成新串的那一支）
    s = "a" & z
    Debug.Print "C18|" & CStr(s = "a") & "|len=" & CStr(Len(s))
    Dim t As String
    t = "a"
    Debug.Print "C19|" & CStr(s = t) & "|" & CStr(t = s)
    Debug.Print "C20|" & CStr(s > t) & "|" & CStr(t < s)
    s = "a"
    Debug.Print "C21|" & CStr(s = "a" & Chr(0))

    Select Case "a "
        Case "a"
            Debug.Print "D1|pad=HIT"
        Case Else
            Debug.Print "D1|pad=MISS"
    End Select

    Select Case ""
        Case Chr(0)
            Debug.Print "D3|emptyVsNul=HIT"
        Case Else
            Debug.Print "D3|emptyVsNul=MISS"
    End Select

    ' 变量选择器：NUL 结尾 / 空格结尾两类（真机 P367 D5/D6）
    Dim u As String
    u = "a" & z
    Select Case u
        Case "a"
            Debug.Print "D5|varNul=HIT"
        Case Else
            Debug.Print "D5|varNul=MISS"
    End Select

    Dim w As String
    w = "a "
    Select Case w
        Case "a"
            Debug.Print "D6|varPad=HIT"
        Case Else
            Debug.Print "D6|varPad=MISS"
    End Select

    Select Case "abc"
        Case "ab" To "abd"
            Debug.Print "D4|range=HIT"
        Case Else
            Debug.Print "D4|range=MISS"
    End Select
End Sub