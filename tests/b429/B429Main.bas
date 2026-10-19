Option Explicit

' <vbeclipse> regression fixture (knife 99, ledger 305 #428).
' FormatDateTime died on its very first call in an x64 build.  vb6_FormatDateTime lives in
' vb6rtl_system.c and is named by the builtin table, so cgen emits it straight into the USER
' translation unit -- but no RTL header declared it, and C then assumes "returns int".  On x64
' that folds the returned BSTR pointer to 32 bits before vb6_BSTR_AssignMove dereferences it =>
' 0xC0000005 (probe .build/k423_proj -> k424_run.txt: the run stops right after the first
' marker; k426_new_x64_run.txt vs k426_new_x86_run.txt: same source, x87-agnostic x86 ran all
' seven calls).  BASE (the compiler before this knife) is red the same way, so this is a
' pre-existing defect, not a move by knife 98.
' A census of "names cgen can emit as a literal" against "names the RTL headers declare" left
' exactly one gap (.build/k428_emit_gap.py -> k428_emit_gap.txt: emitted 882, declared 1587,
' gap = 1), which is why this fixture comes with a sentinel over that whole question instead of
' a rule about this one name.
' What gets pinned is deliberately locale-free: the five named formats go through
' GetDateFormatW / GetTimeFormatW with LOCALE_USER_DEFAULT, so their STRINGS differ between this
' machine (zh-CN) and a CI runner (en-US) -- the oracle readings (.build/k463_fdt2.vbs ->
' k463_fdt2_out.txt) confirm the shapes but cannot be written into a regression.  So: VarType
' says a real string came back, no arm is empty, general date starts with the short date and
' contains the long time, long date differs from short date, long time differs from short time,
' and the digits the calendar says are in there.  Reaching B429-DONE is itself the witness that
' the x64 leg survives the return.

Sub Main()
    Dim d1 As Date
    Dim s0 As String
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim s4 As String
    Dim v As Variant

    d1 = CDate("2026/1/5 14:30:45")
    s0 = FormatDateTime(d1, 0)
    s1 = FormatDateTime(d1, 1)
    s2 = FormatDateTime(d1, 2)
    s3 = FormatDateTime(d1, 3)
    s4 = FormatDateTime(d1, 4)

    ' --- the return really is a VB string, in every arm (0 = 1 = 8 = vbString)
    Debug.Print "B429-01=" & VarType(s0)
    Debug.Print "B429-02=" & VarType(s1) & "/" & VarType(s2) & "/" & VarType(s3) & "/" & VarType(s4)

    ' --- no arm answers empty
    Debug.Print "B429-03=" & (Len(s0) > 0) & "/" & (Len(s1) > 0) & "/" & (Len(s2) > 0)
    Debug.Print "B429-04=" & (Len(s3) > 0) & "/" & (Len(s4) > 0)

    ' --- relations between the arms, not their spellings
    Debug.Print "B429-05=" & (Left(s0, Len(s2)) = s2)
    Debug.Print "B429-06=" & (InStr(s0, s3) > 0)
    Debug.Print "B429-07=" & (s1 <> s2)
    Debug.Print "B429-08=" & (s3 <> s4)
    Debug.Print "B429-09=" & (Len(s0) > Len(s2))

    ' --- the digits the calendar says have to be inside the two date arms
    Debug.Print "B429-10=" & (InStr(s1, CStr(DatePart("yyyy", d1))) > 0)
    Debug.Print "B429-11=" & (InStr(s2, CStr(DatePart("yyyy", d1))) > 0)
    Debug.Print "B429-12=" & (InStr(s2, CStr(DatePart("d", d1))) > 0)

    ' --- the same arm reached from a Double instead of from a Date
    Debug.Print "B429-13=" & (FormatDateTime(CDbl(46027.6046875), 2) = s2)

    ' --- the second formal is optional and its default is vbGeneralDate (0)
    Debug.Print "B429-14=" & (FormatDateTime(d1) = s0)

    ' --- the same BSTR into a Variant slot rather than a String slot
    v = FormatDateTime(d1, 2)
    Debug.Print "B429-15=" & VarType(v) & "/" & (CStr(v) = s2)

    ' --- VarType of the CALL itself.  This is the witness that the user TU now knows the return
    ' type: with no prototype C assumes int, the packing station picks the integer arm and the
    ' answer is 3 (VT_I4) -- measured on the BASE machine, .build/k467_x86_run.txt reads 3/3.
    ' With the declaration it is 8 (VT_BSTR) in both legs, so this line moves on x86 too and the
    ' x64 leg no longer dies before it.
    Debug.Print "B429-16=" & (VarType(FormatDateTime(d1, 2)) = 8) & "/" & (VarType(FormatDateTime(0, 2)) = 8)

    ' --- serial 0 takes the "read the local clock" branch: still a string, still not empty
    Debug.Print "B429-17=" & (Len(FormatDateTime(0, 2)) > 0)

    Debug.Print "B429-DONE"
End Sub
