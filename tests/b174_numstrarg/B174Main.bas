Option Explicit

' <vbeclipse> regression fixture (knife 65, ledger 305 B174).
' "A numeric actual argument handed to a VB string formal" is folded the way VB6 folds it
' (implicit CStr). That rule used to have TWO codegen stations: Fix 175 (asks the formal's
' declared type, shape = wrapToBSTR) and Fix 108c (a 13-name whitelist with its own shapes,
' applied by text search once the argument list was already one string). Knife 65 keeps
' Fix 175 as the single station and closes the hole that station could not reach: intrinsics
' that deliberately carry NO formal table -- InStr, whose optional leading start would
' misalign every position. Pinned here too: the memory-width family (Len / LenB) must NOT
' fold (VB6 answers the byte width of the value, not CStr of it), and an unsigned actual has
' to fold as well (that was the whitelist's own coverage, now in the shared type gate).

Private Type U74
    UV As Variant
End Type

Private Sub ShowNum(ByVal s As String)
    Debug.Print "B74-SN=" & s
End Sub

Private Sub Shapes()
    Dim n As Long
    n = 12
    Dim d As Double
    d = 1.5
    Dim b As Boolean
    b = True
    Dim s As String
    s = "withform"
    Dim v As Variant
    v = "withform"
    Dim q As ULong
    q = 4000000000
    Dim u As U74
    u.UV = "withform"
    Debug.Print "B74-1=" & InStr(n, "1")
    Debug.Print "B74-2=" & Len(n)
    Debug.Print "B74-3=" & InStr(d, "5")
    Debug.Print "B74-4=" & InStr(b, "-1")
    Debug.Print "B74-5=" & Left(n, 1)
    Debug.Print "B74-6=" & Mid(n, 2, 1)
    Debug.Print "B74-7=" & Replace(n, "1", "9")
    Debug.Print "B74-8=" & StrComp(n, "12")
    Debug.Print "B74-9=" & Trim(n)
    Debug.Print "B74-10=" & InStr(s, "t")
    Debug.Print "B74-11=" & InStr(v, "t")
    Debug.Print "B74-12=" & InStr(u.UV, "t")
    Debug.Print "B74-13=" & Len(s)
    Debug.Print "B74-14=" & Len(v)
    Debug.Print "B74-15=" & Left(q, 10)
    Debug.Print "B74-16=" & InStr(2, s, "t")
    Debug.Print "B74-17=" & LenB(n)
    Debug.Print "B74-18=" & (LenB(u) > LenB(n))
    Call ShowNum(n)
End Sub

Public Sub Main()
    Shapes
    Debug.Print "B74-DONE"
End Sub
