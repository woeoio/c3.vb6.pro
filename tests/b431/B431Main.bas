' Fix 431 (ledger 305 #429 items 1 and 4, knife 101): "general date" must have ONE home in the RTL.
'
' Caliber: every needle is RELATIONAL or locale-free. Each arm ends up in GetDateFormatW /
' GetTimeFormatW with LOCALE_USER_DEFAULT (this box is zh-CN, the CI runner is en-US), so no date
' SPELLING is pinned - only the SHAPE of the answer (colon or no colon, trailing space or not,
' short + " " + long or not) and which station produced it.
'
' Before this knife the same value got THREE answers (probe .build/k541_proj -> k541_x64_run.txt,
' both arches identical; engine itself = .build/k559_olefdt3.c -> k559_out.txt):
'   FormatDateTime(#1/5/2026#, vbGeneralDate) = "2026/1/5 "        len 9  (orphan trailing space)
'   Format(#1/5/2026#, "General Date")        = "2026/1/5 0:00:00" len 16
'   CStr(#1/5/2026#)                          = "2026/1/5"         len 8  (the only engine match)
' Mechanism: vb6_FormatDateTime wrote buf[len - 1] = ' ' UNCONDITIONALLY while GetDateFormatW
' returns a length that INCLUDES the terminator, so that write turned the terminator into a space -
' only the time segment was gated by the fraction test. conv.c kept a second hand-written fraction
' rule (1e-9, Fix 175) beside it, so the gate had three owners and three answers.
'
' The engine's gate is measured point by point in .build/k564_oleband.c -> k564_out.txt: the time
' segment is dropped when the fraction of a day rounds to ZERO WHOLE SECONDS, i.e. below half a
' second - 50/200/400/499/500 ms answer date-only, 600 ms answers with a time. That is neither
' "fraction != 0" (system.c) nor "fraction < 1e-9" (conv.c). B431-07/08/09 pin the band on both
' sides, B431-14 the midnight carry-over that sinks both rival rules at once.
' B431-12/13 are the blast-radius witnesses: the numeric Format arm and knife 100's ddddd / ttttt
' twins must not have been dragged into the shared outlet.
Option Explicit

Sub Main()
    Dim d0 As Date, d1 As Date
    Dim q1 As Date, q2 As Date, q3 As Date, m1 As Date
    Dim v0 As Variant, v1 As Variant
    Dim g0 As String, s0 As String, l0 As String
    Dim g1 As String, s1 As String, t1 As String

    d0 = #1/5/2026#
    d1 = #1/5/2026 2:30:45 PM#
    q1 = d0 + 0.25 / 86400
    q2 = d0 + 0.6 / 86400
    q3 = d0 + 1 / 86400 / 1000000000
    m1 = #1/6/2026# - 0.001 / 86400
    v0 = d0
    v1 = d1

    g0 = FormatDateTime(d0, vbGeneralDate)
    s0 = FormatDateTime(d0, vbShortDate)
    l0 = FormatDateTime(d0, vbLongDate)
    g1 = FormatDateTime(d1, vbGeneralDate)
    s1 = FormatDateTime(d1, vbShortDate)
    t1 = FormatDateTime(d1, vbLongTime)

    ' 01/02: no orphan space, no phantom 00:00:00 for a whole day
    Debug.Print "B431-01=" & (Right$(g0, 1) <> " ") & "/" & (InStr(g0, ":") = 0)
    Debug.Print "B431-02=" & (g0 = s0) & "/" & (Len(g0) = Len(s0))
    ' 03: the three stations that used to disagree now answer one string (zero fraction)
    Debug.Print "B431-03=" & (Format(d0, "General Date") = g0) & "/" & (CStr(d0) = g0) & "/" & (CStr(v0) = g0)
    ' 04/05: a real fraction keeps the time, composed as short + ONE space + long time
    Debug.Print "B431-04=" & (InStr(g1, ":") > 0) & "/" & (Len(g1) > Len(s1))
    Debug.Print "B431-05=" & (g1 = s1 & " " & t1) & "/" & (Left$(g1, Len(s1) + 1) = s1 & " ")
    ' 06: same agreement with a fraction present, all three stations
    Debug.Print "B431-06=" & (Format(d1, "General Date") = g1) & "/" & (CStr(d1) = g1) & "/" & (CStr(v1) = g1)
    ' 07/08/09: the engine's half-second band (250 ms -> date only, 600 ms -> time, 1 tick -> date only)
    Debug.Print "B431-07=" & (FormatDateTime(q1, vbGeneralDate) = s0) & "/" & (InStr(FormatDateTime(q1, vbGeneralDate), ":") = 0)
    Debug.Print "B431-08=" & (InStr(FormatDateTime(q2, vbGeneralDate), ":") > 0) & "/" & (FormatDateTime(q2, vbGeneralDate) = s0 & " " & FormatDateTime(q2, vbLongTime))
    Debug.Print "B431-09=" & (FormatDateTime(q3, vbGeneralDate) = s0) & "/" & (Format(q3, "General Date") = s0)
    ' 14: the sharpest cell of the engine's rule - a fraction ONE MILLISECOND short of midnight rounds UP
    '        across midnight, so the rounded time-of-day is zero: the whole segment is dropped AND the date
    '        carries to 1/6. That one row sinks both rival rules (system.c's != 0 prints a time here, and a
    '        sub-second tolerance prints one too). 15 is the same rounding from the rendering side.
    Debug.Print "B431-14=" & (FormatDateTime(m1, vbGeneralDate) = FormatDateTime(#1/6/2026#, vbShortDate)) & "/" & (InStr(FormatDateTime(m1, vbGeneralDate), ":") = 0)
    Debug.Print "B431-15=" & (FormatDateTime(q2, vbLongTime) <> FormatDateTime(d0, vbLongTime)) & "/" & (FormatDateTime(q3, vbLongTime) = FormatDateTime(d0, vbLongTime))
    ' 10/11: the other named arms are still their own answers (long date is not general date,
    '        short date never carries a time, long time is not a date)
    Debug.Print "B431-10=" & (l0 <> g0) & "/" & (InStr(l0, ":") = 0)
    Debug.Print "B431-11=" & (InStr(s0, ":") = 0) & "/" & (Len(s0) > 0) & "/" & (t1 <> s0)
    ' 12: blast radius - the numeric Format arm still prints the NUMBER
    Debug.Print "B431-12=" & (Format(1234, "0") = "1234") & "/" & (CStr(1234) = "1234") & "/" & (CStr(0.5) = "0.5")
    ' 13: blast radius - knife 100's ddddd / ttttt twins still answer their own twins
    Debug.Print "B431-13=" & (Format(d1, "ddddd") = s1) & "/" & (Format(d1, "ttttt") = t1)
    Debug.Print "B431-DONE"
End Sub
