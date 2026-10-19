' Fix 431 (ledger 305 #431, knife 102): DateSerial's RETURN TAG is one cell in the intrinsic table.
'
' Caliber: same as knife 101 - every needle is relational, locale-free, or pinned to an ASCII format
' pattern. The box is zh-CN and the CI runner is en-US, so date SPELLINGS are never pinned; where a
' spelled answer is needed the value is formatted with an explicit "yyyy-mm-dd" pattern.
'
' Before this knife src/semantics/builtin/builtin_funcs.inc wrote Vb6Type::Double for DateSerial while
' the same table writes Vb6Type::Date for Now / Date / Time / DateAdd / DateValue / TimeSerial /
' TimeValue / FileDateTime. The RTL returns double for all of them, so that one cell is what decides
' which VARIANT tag the value carries when it enters a slot. Measured with .build/k607 (both arches
' identical) against the on-box engine (.build/k607_oracle_out.txt): VarType 5 / TypeName "Double" /
' IsDate False / IsNumeric True / CStr [46027] / Len 5 - six answers inverted; the engine is
' 7 / Date / True / False / [2026/1/5] / 8.
'
' B432-05/06/07/08/09/14 are witnesses, not fixes: the shapes that already agreed with the engine must
' still agree. B432-09 is the numeric Format arm - the RTL comment that justified it claimed DateSerial
' returned VT_R8; the arm itself is kept for genuine numbers and only its rationale changed.
' B432-17 pins the other half of this knife: the narrowing station (cgen_util_type.cpp narrowCheckAssign,
' ledger #248) only routed Single / Double / Currency through vb6_FltToLng, so a Date handed to a Long
' target fell into "fits, no check" and vb6_ChkLong(int64_t) truncated at the call boundary. B432-12 is
' the same station on a Double source and is a witness - it must stay green, both arches.
Option Explicit

Private Function SlotOf(ByVal v As Variant) As String
    SlotOf = CStr(VarType(v)) & "/" & TypeName(v) & "/" & CStr(IsDate(v))
End Function

Private Function DateArg(ByVal p As Date) As String
    DateArg = CStr(VarType(p)) & "/" & CStr(IsDate(p))
End Function

Sub Main()
    Dim v As Variant
    Dim dt As Date
    Dim s As String
    Dim n As Long
    Dim m As Long
    Dim k As Long

    v = DateSerial(2026, 1, 5)
    dt = DateSerial(2026, 1, 5)

    Debug.Print "B432-01=" & CStr(VarType(v) = 7) & "/" & CStr(TypeName(v) = "Date") & "/" & _
        CStr(IsDate(v)) & "/" & CStr(IsNumeric(v) = False)
    Debug.Print "B432-02=" & CStr(Len(v) = 8) & "/" & CStr(CDbl(v) = 46027)
    Debug.Print "B432-03=" & CStr(CStr(v) = CStr(CDate(46027))) & "/" & CStr(CStr(dt) = CStr(CDate(46027)))
    Debug.Print "B432-04=" & CStr(SlotOf(DateSerial(2026, 1, 5)) = "7/Date/True") & "/" & _
        CStr(DateArg(DateSerial(2026, 1, 5)) = "7/True")
    Debug.Print "B432-05=" & CStr(VarType(DateSerial(2026, 1, 5) + TimeSerial(14, 30, 45)) = 7) & "/" & _
        CStr(CDbl(DateSerial(2026, 1, 5) + TimeSerial(14, 30, 45)) = 46027.6046875)
    Debug.Print "B432-06=" & CStr(VarType(DateSerial(2026, 1, 6) - DateSerial(2026, 1, 5)) = 5) & "/" & _
        CStr((DateSerial(2026, 1, 6) - DateSerial(2026, 1, 5)) = 1)
    Debug.Print "B432-07=" & CStr(DateSerial(2026, 1, 5) = 46027) & "/" & _
        CStr(DateSerial(2026, 1, 5) = CDate(46027))
    Debug.Print "B432-08=" & CStr(Format(DateSerial(2020, 1, 1), "mmm") = "Jan") & "/" & _
        CStr(Format(DateSerial(2026, 1, 5), "yyyy-mm-dd") = "2026-01-05")
    Debug.Print "B432-09=" & CStr(Format(43831, "mmm") = "Jan") & "/" & _
        CStr(Format(43831, "yyyy-mm-dd") = "2020-01-01")
    Debug.Print "B432-10=" & CStr(Year(v) = 2026) & "/" & CStr(Month(v) = 1) & "/" & _
        CStr(Day(v) = 5) & "/" & CStr(Weekday(v) = 2)

    s = DateSerial(2026, 2, 30)
    Debug.Print "B432-11=" & CStr(s = CStr(CDate(46083))) & "/" & _
        CStr(CDbl(DateSerial(2026, 2, 30)) = 46083)

    n = DateSerial(2026, 1, 5) + 0.7
    m = CLng(DateSerial(2026, 1, 5) + 0.7)
    Debug.Print "B432-12=" & CStr(n = 46028) & "/" & CStr(n = m) & "/" & CStr(m = 46028)

    v = DateSerial(2026, 1, 5) + 1
    Debug.Print "B432-13=" & CStr(VarType(v) = 7) & "/" & CStr(CStr(v) = CStr(CDate(46028)))

    Debug.Print "B432-14=" & CStr(VarType(Now) = 7) & "/" & CStr(VarType(Date) = 7) & "/" & _
        CStr(VarType(TimeSerial(1, 2, 3)) = 7) & "/" & CStr(VarType(DateAdd("d", 1, CDate(46027))) = 7)
    Debug.Print "B432-15=" & CStr(DateDiff("d", DateSerial(2026, 1, 5), DateSerial(2026, 1, 6)) = 1) & "/" & _
        CStr(CStr(DateSerial(2026, 1, 5)) = CStr(dt))
    Debug.Print "B432-16=" & CStr(FormatDateTime(DateSerial(2026, 1, 5), 0) = FormatDateTime(CDate(46027), 0)) & "/" & _
        CStr(FormatDateTime(DateSerial(2026, 1, 5), 2) = FormatDateTime(CDate(46027), 2))

    ' The narrowing station itself, measured on a value that ALREADY answered Date before this knife:
    ' CDate() is Vb6Type::Date in the table, so BASE handed the double straight to vb6_ChkLong(int64_t)
    ' and C truncated at the call boundary (46027). This needle is BASE-red / NEW-green on its own.
    k = CDate(46027) + 0.7
    Debug.Print "B432-17=" & CStr(k = 46028) & "/" & CStr(k = CLng(CDate(46027) + 0.7))

    ' The junction with knife 101, and the LABEL WITNESS that knife 102 did NOT close but
    ' knife 104 DID. First clause: a date that first sits in a Variant slot reaches Format's
    ' no-format-string path (the general-date outlet) and answers like CStr -- that half was
    ' BASE-red / NEW-green at knife 102. Second clause: a Dim-As-Date local handed straight to
    ' Format used to be boxed by the hand-copied table in ledger #430, so it kept answering
    ' with the number -- that half was pinned WRONG on purpose ("whoever fixes #430 must turn
    ' it True").
    ' >>> knife 104 (#430) landed: the Format first-arg Date arm now asks the sole authority
    '     boxToVariant (vb6_VariantDate) instead of its own copied list (vb6_VariantDouble),
    '     so the second clause flipped to True. This cell is the witness that it really is so.
    '     Readings: before = True/False, after = True/True (both arches, byte-identical).
    ' (measured: .build/k638_base_run.txt vs k638_new_run.txt U09; k636_charts_diff.txt)
    v = DateSerial(2026, 1, 5)
    Debug.Print "B432-18=" & CStr(Format(v) = CStr(CDate(46027))) & "/" & _
        CStr(Format(dt) = CStr(CDate(46027)))

    Debug.Print "B432-DONE"
End Sub
