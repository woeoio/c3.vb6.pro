Option Explicit

' <vbeclipse> regression fixture (knife 93, ledger 305 #416).
' A string fed to a strict conversion (CByte/CInt/CLng/CSng/CDbl/CCur/CBool and the
' ai/032 unsigned tier) must get ONE answer about the text. Before this knife C3
' answered silently with 0 for non-numeric text and swallowed trailing junk
' (CInt("abc")=0, CInt("")=0, CInt("123abc")=123); VB6 raises 13 Type mismatch for
' all three. Range stays error 6 -- that answer already has its own authority
' (vb6_OvfChk) and is pinned here so this knife cannot quietly move it.
' Val() is the OTHER authority and stays lenient on purpose.

Sub Main()
    On Error Resume Next
    Dim s As String
    Dim v As Variant
    Dim w As Variant
    Dim b As Boolean
    Dim r As Long

    ' --- three shapes VB6 answers with 13
    Err.Clear
    r = CInt("abc")
    Debug.Print "B416-01=" & Err.Number
    Err.Clear
    r = CInt("")
    Debug.Print "B416-02=" & Err.Number
    Err.Clear
    s = "123abc"
    r = CInt(s)
    Debug.Print "B416-03=" & Err.Number

    ' --- overflow keeps its own number (6), not 13
    Err.Clear
    r = CInt("70000")
    Debug.Print "B416-04=" & Err.Number

    ' --- accepted shapes still answer the value
    r = CLng("&HFF")
    Debug.Print "B416-05=" & r
    r = CLng("123")
    Debug.Print "B416-06=" & r
    r = CByte("&HFF")
    Debug.Print "B416-07=" & r
    r = CLng(" 42 ")
    Debug.Print "B416-08=" & r

    ' --- Val() untouched: still the lenient one
    Debug.Print "B416-09=" & Val("123abc")
    Debug.Print "B416-10=" & Val("abc")

    ' --- Boolean asks the boolean question, not the numeric one
    b = CBool("True")
    Debug.Print "B416-11=" & b
    Err.Clear
    b = CBool("abc")
    Debug.Print "B416-12=" & Err.Number
    Err.Clear
    r = CLng("True")
    Debug.Print "B416-13=" & Err.Number

    ' --- the Variant route asks the SAME question (that is the point of the knife)
    v = "abc"
    Err.Clear
    r = CLng(v)
    Debug.Print "B416-14=" & Err.Number
    v = "42"
    r = CLng(v)
    Debug.Print "B416-15=" & r
    ' Empty is not text: it keeps answering 0 and must NOT start raising.
    Err.Clear
    r = CLng(w)
    Debug.Print "B416-16=" & r & "/" & Err.Number

    Debug.Print "B416-DONE"
End Sub
