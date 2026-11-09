Option Explicit
Option Compare Binary

' Option Compare Binary（PhotoDemon 的 Modules\*.bas 里常见）。
' 词法把 "binary" 统一归 TokenKind::Binary（Open ... For Binary 也用它），
' 修前 parseOption 只认 Binary2 ⇒ VB2001 expected 'Text' or 'Binary'。
' 真 VB6：`"abc" = "ABC"` 在 Binary 下是 False。

Public Sub Main()
    Dim a As String
    Dim b As String
    a = "abc"
    b = "ABC"
    Debug.Print "OCB=" & (a = b) & "," & Len(a)
End Sub
