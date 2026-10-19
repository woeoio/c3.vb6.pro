Option Explicit

' <vbeclipse> regression fixture (knife 98, ledger 305 #426).
' Format's date face answered three wrong questions:
'   (1) h / hh always ran through the 12-hour expression (st.wHour % 12), so a pattern with no
'       AM/PM designator lost the afternoon.  The corpus uses "hh:mm:ss" and
'       "yyyy-mm-dd hh:mm:ss" as TIMESTAMPS -- 70 and 101 sites of 1693 dated Format calls
'       (.build/k392_fmt_census.py) -- so those programs were writing half an afternoon into
'       their logs.  VB caliber: h/hh is 24-hour and only a designator in the pattern switches
'       it to 12-hour.
'   (2) the AM/PM designator itself was eaten one character at a time: 'A' fell to the literal
'       branch and 'M' became the month, so Format(d,"AM/PM") echoed "A1/P1" (the 1 and the 7 in
'       the two probe readings below are the two MONTH numbers -- witness that the letter was
'       read as a month token, .build/k391_run.txt).
'   (3) a pattern that is nothing but a designator ("A/P", "AMPM") never reached the date branch
'       at all, because vb6_fmtIsDateFormat only looked for y/d/h/m/s/n, and printed the raw
'       serial 46027.6046875.
' One owner now answers "is there a designator, how long is it, which shape, and does the hour
' run 12- or 24-hour": vb6_fmtAmPmLenAt + vb6_fmtHasAmPm + vb6_fmtAppendAmPm.
' Caliber of the letters: invariant English "AM"/"PM"/"A"/"P", lower-cased only when the pattern
' wrote the token in lower case -- same rank as this function's own January/Sunday tables, i.e.
' NOT locale-dependent.
' WARNING on provenance: there is NO runtime oracle for Format on this machine (VBScript has no
' Format; oleaut32 VarTokenizeFormatString returns E_INVALIDARG for every pattern here, ruler
' .build/k388_varfmt.c; .NET Strings.Format uses its own token table, .build/k390_vbnet.txt).
' The expectations below come from the documented VB6 table plus how the corpus writes these
' patterns -- deliberately NOT recorded as "measured against the runtime".
' B426-14 is the witness group for the mm ambiguity (month vs minute) that the designator branch
' must not disturb; B426-17 is the witness that a numeric pattern still goes the numeric way.

Sub Main()
    Dim d1 As Date
    Dim d2 As Date
    Dim d3 As Date
    Dim d4 As Date
    Dim d5 As Date
    Dim d6 As Date

    d1 = CDate("2026/1/5 14:30:45")
    d2 = CDate("2026/7/4 0:05:00")
    d3 = CDate("2026/1/5 12:00:00")
    d4 = CDate("2026/1/5 0:00:00")
    d5 = CDate("2026/1/5 23:00:00")
    d6 = CDate("2026/1/5 13:00:00")

    ' --- arm (1): no designator in the pattern => 24-hour
    Debug.Print "B426-01=" & Format(d1, "hh:mm:ss") & "|" & Format(d2, "hh:mm:ss")
    Debug.Print "B426-02=" & Format(d1, "h:nn") & "|" & Format(d2, "h:nn")
    Debug.Print "B426-03=" & Format(d1, "yyyy-mm-dd hh:mm:ss") & "|" & Format(d2, "yyyy-mm-dd hh:mm:ss")

    ' --- arms (2) + (3): the designator renders, and it is what switches the hour to 12
    Debug.Print "B426-04=" & Format(d1, "hh:mm AM/PM") & "|" & Format(d2, "hh:mm AM/PM")
    Debug.Print "B426-05=" & Format(d1, "h:mm:ss a/p") & "|" & Format(d2, "h:mm:ss a/p")
    Debug.Print "B426-06=" & Format(d1, "h:nn AMPM") & "|" & Format(d2, "h:nn AMPM")
    Debug.Print "B426-07=" & Format(d1, "hhmmss AMPM") & "|" & Format(d2, "hhmmss AMPM")
    Debug.Print "B426-08=" & Format(d1, "h:mm ampm") & "|" & Format(d2, "h:mm ampm")
    Debug.Print "B426-09=" & Format(d1, "AM/PM") & "|" & Format(d2, "AM/PM")
    Debug.Print "B426-10=" & Format(d1, "A/P") & "|" & Format(d2, "A/P")
    Debug.Print "B426-11=" & Format(d1, "am/pm") & "|" & Format(d2, "am/pm")
    Debug.Print "B426-12=" & Format(d1, "d-mmm-yy h:nn AMPM") & "|" & Format(d2, "d-mmm-yy h:nn AMPM")
    Debug.Print "B426-13=" & Format(d1, "yyyy-mm-dd hh:mm:ss AM/PM") & "|" & Format(d2, "yyyy-mm-dd hh:mm:ss AM/PM")

    ' --- witnesses: month vs minute vs hour, none of which this knife may move
    Debug.Print "B426-14=" & Format(d1, "yyyy-mm-dd") & "|" & Format(d1, "mm/dd/yyyy") & "|" & Format(d1, "hh:mm")

    ' --- the two 12-hour edges: noon and midnight keep 12, and stay 24-hour without a designator
    Debug.Print "B426-15=" & Format(d3, "hh:mm AM/PM") & "|" & Format(d4, "hh:mm AM/PM") & "|" & Format(d3, "hh:mm") & "|" & Format(d4, "hh:mm")
    Debug.Print "B426-16=" & Format(d5, "h:mm AMPM") & "|" & Format(d6, "hh")

    ' --- witness: a pattern with no date letter at all still goes the numeric way
    '     (both shapes are locale-free on purpose: no grouping character, so the answer is the
    '     same on any runner regardless of its decimal/thousands digits)
    Debug.Print "B426-17=" & Format(1234.5, "0.00") & "|" & Format(-1234.5, "0.00")

    Debug.Print "B426-DONE"
End Sub
