' Fix 430 (ledger 305 #427, knife 100): Format's date face must answer w / ww / q / ttttt / ddddd.
'
' Caliber: every needle here is RELATIONAL or locale-free. The five new arms go through
' GetDateFormatW / GetTimeFormatW with LOCALE_USER_DEFAULT, and the CI runner is en-US while this
' box is zh-CN, so no date SPELLING is pinned. What is pinned is which QUESTION each token answers:
'   w     weekday number   = DatePart("w")        = Weekday(d, vbSunday)
'   ww    week of year      = DatePart("ww")       (and it really reads FirstDayOfWeek /
'                                                   FirstWeekOfYear -- B430-19/20/21)
'   q     quarter           = DatePart("q")
'   ttttt long time         = FormatDateTime(d, vbLongTime)
'   ddddd short date        = FormatDateTime(d, vbShortDate)
' dddd (four letters) is still the long weekday name: the old day arm read cnt >= 4 and swallowed
' the fifth d, which is why the corpus's "ddddd ttttt" (36 sites) printed "Monday ttttt".
'
' Before this knife w / ww / q / ttttt fell out of the gate (vb6_fmtIsDateFormat only knew
' y/d/h/m/s/n plus the AM/PM designators) and the VariantTime came back as a bare double:
' Format(d1,"w") answered 46027.6046875 (probe .build/k504_proj -> k504_x64_run.txt).
' B430-24/25 are the blast-radius witnesses: the corpus writes unit text BARE in numeric format
' strings ("#,##0 Twip", "T0", "#### Byte" - 15 sites, .build/k496_census.txt), so the numeric arm
' still must print the NUMBER, not a date. The gate therefore admits a lone w only when the value
' really carries VT_DATE; q / ww / ttttt are admitted wherever, and a single t stays a literal
' (that is what keeps the T in the ISO-ish pattern of B430-22/23).
'
' There is NO runtime oracle for VB6's Format on this box (no VBScript Format, and oleaut32's
' VarTokenizeFormatString returns E_INVALIDARG for every string here), so the expected values come
' from the documented token table plus these internal relations - not from a VB6 measurement.
Sub Main()
    Dim d1 As Date, d2 As Date, d3 As Date, e1 As Date
    Dim s1 As String
    d1 = #1/5/2026 2:30:45 PM#
    d2 = #12/31/2026 9:07:03 AM#
    d3 = #1/1/2026 8:00:00 AM#
    e1 = #1/5/2026#

    ' 01..05: none of the five may answer the bare VariantTime any more
    s1 = Format(d1, "w")
    Debug.Print "B430-01=" & (s1 <> CStr(d1)) & "/" & (InStr(s1, ".") = 0)
    Debug.Print "B430-02=" & (InStr(Format(d1, "ww"), ".") = 0)
    Debug.Print "B430-03=" & (InStr(Format(d1, "q"), ".") = 0)
    Debug.Print "B430-04=" & (InStr(Format(d1, "ttttt"), ".") = 0)
    Debug.Print "B430-05=" & (InStr(Format(d1, "ddddd"), ".") = 0)

    ' 06..15: each token answers the same question as its own DatePart / FormatDateTime twin
    Debug.Print "B430-06=" & (Format(d1, "w") = CStr(DatePart("w", d1)))
    Debug.Print "B430-07=" & (Format(d1, "w") = CStr(Weekday(d1, vbSunday)))
    Debug.Print "B430-08=" & (Format(d2, "w") = CStr(DatePart("w", d2)))
    Debug.Print "B430-09=" & (Format(d1, "ww") = CStr(DatePart("ww", d1)))
    Debug.Print "B430-10=" & (Format(d2, "ww") = CStr(DatePart("ww", d2)))
    Debug.Print "B430-11=" & (Format(d1, "q") = CStr(DatePart("q", d1)))
    Debug.Print "B430-12=" & (Format(d2, "q") = CStr(DatePart("q", d2)))
    Debug.Print "B430-13=" & (Format(d1, "ttttt") = FormatDateTime(d1, vbLongTime))
    Debug.Print "B430-14=" & (Format(d1, "ddddd") = FormatDateTime(d1, vbShortDate))
    Debug.Print "B430-15=" & (Format(e1, "ddddd") = FormatDateTime(e1, vbShortDate))

    ' 16..18: the neighbours must not have been dragged along; the invariant-English name table stays
    Debug.Print "B430-16=" & (Format(d1, "dddd") = "Monday") & "/" & (Format(d2, "ddd") = "Thu") & "/" & (Format(d2, "dddd") = "Thursday")
    Debug.Print "B430-17=" & (Format(d1, "ddddd") <> Format(d1, "dddd"))
    Debug.Print "B430-18=" & (Format(d1, "ddddd ttttt") = Format(d1, "ddddd") & " " & Format(d1, "ttttt"))

    ' 19..21: the two formals are really read (d3 = 1/1/2026, a Thursday: 1 vs 3 split the year tail)
    Debug.Print "B430-19=" & (Format(d3, "ww", vbMonday, vbFirstFourDays) = CStr(DatePart("ww", d3, vbMonday, vbFirstFourDays)))
    Debug.Print "B430-20=" & (Format(d3, "ww", vbMonday, vbFirstFullWeek) <> Format(d3, "ww", vbMonday, vbFirstJan1))
    Debug.Print "B430-21=" & (Format(d3, "ww", vbUseSystemDayOfWeek, vbUseSystem) = CStr(DatePart("ww", d3, vbUseSystemDayOfWeek, vbUseSystem)))

    ' 22..23: locale-free literals -- the w/ww inside a y-bearing string, and the literal T of an ISO pattern
    Debug.Print "B430-22=" & (Format(d1, "yyyy w ww") = "2026 " & CStr(DatePart("w", d1)) & " " & CStr(DatePart("ww", d1)))
    Debug.Print "B430-23=" & (Format(d1, "yyyy-mm-ddTHH:mm:ss") = "2026-01-05T14:30:45")

    ' 24..25: the numeric arm still prints numbers (unit text carried by 0/#, no date)
    Debug.Print "B430-24=" & (Format(1234, "T0") = "1234") & "/" & (Format(1234, "0 Twip") = "1234")
    Debug.Print "B430-25=" & (Format(1234, "#,##0") = "1,234") & "/" & (Format(1234, "0") = "1234")

    ' 26..27: the boxing tag of the two new locale arms (a BSTR, not the integer arm)
    Debug.Print "B430-26=" & (VarType(Format(d1, "ddddd")) = 8) & "/" & (VarType(Format(d1, "ttttt")) = 8)
    Debug.Print "B430-27=" & (VarType(Format(d1, "w")) = 8) & "/" & (VarType(Format(d1, "q")) = 8)

    Debug.Print "B430-DONE"
End Sub
