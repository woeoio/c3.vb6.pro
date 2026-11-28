Option Explicit

' §11.235 (第二百五十三刀): 搜索面（InStr / InStrRev / Like）遇内嵌 NUL ——
' §11.233/234 只量了比较面。夹具 = 真机探针 .temp/k235p/P370 + P371 的同一批语句。
' 口径：NUL 是普通字符（搜索不截断、`?`/`#`/`[...]`/`*` 都按位置吃字符），
' 但 "空串做 haystack/needle" 的老口径（InStr("","") = 0、InStrRev("abc","") = 3）不变。
Public Sub Main()
    Dim z As String
    z = Chr(0)

    ' --- P370: InStr / InStrRev 遇 NUL
    Debug.Print "V1|" & CStr(InStr("a" & z, "a"))
    Debug.Print "V2|" & CStr(InStr("a" & z, z))
    Debug.Print "V3|" & CStr(InStr("a" & z & "b", "b"))
    Debug.Print "V4|" & CStr(InStr(z, "a"))
    Debug.Print "V5|" & CStr(InStrRev("a" & z & "b", "b"))
    Debug.Print "V6|" & CStr(InStrRev("a" & z & "b", z))
    Debug.Print "V7|" & CStr(Len("a" & z & "b"))
    Debug.Print "V8|" & CStr(("a" & z & "b") Like "ab")
    Debug.Print "V9|" & CStr(("a" & z) Like "a?")
    Debug.Print "V10|" & CStr(InStr("abc", ""))
    Debug.Print "V11|" & CStr(InStr("", ""))
    Debug.Print "V12|" & CStr(InStrRev("abc", ""))
    Debug.Print "V13|" & CStr(InStr(1, "A" & z, "a", 1))
    Debug.Print "V14|" & CStr(InStr(z & z & "x", "x"))
    Debug.Print "V15|" & CStr(InStrRev("a" & z & "a", z))

    ' --- P371: Like 的各档通配符遇 NUL（模式一律加括号：`Like` 与 `&` 的优先级是另一刀）
    Debug.Print "W1|" & CStr(("a" & z & "b") Like "a?b")
    Debug.Print "W2|" & CStr(("a" & z) Like "a")
    Debug.Print "W3|" & CStr(("a" & z) Like "a*")
    Debug.Print "W4|" & CStr(z Like "?")
    Debug.Print "W5|" & CStr(("a" & z) Like "a[!b]")
    Debug.Print "W6|" & CStr((z & "5") Like "#5")
    Debug.Print "W7|" & CStr(("a" & z & "b") Like "a#b")
    Debug.Print "W8|" & CStr("abc" Like "a*")
    Debug.Print "W9|" & CStr(("a" & z) Like ("a" & z))
    Debug.Print "W10|" & CStr(("a" & z & "b") Like "a[!x]b")
    Debug.Print "W11|" & CStr((z) Like "[]")
    Debug.Print "W12|" & CStr(("x" & z) Like ("*" & z))
    Debug.Print "W13|" & CStr("ab" Like ("a" & "b"))
End Sub