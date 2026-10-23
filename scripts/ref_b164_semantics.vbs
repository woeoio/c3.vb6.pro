Option Explicit
' =====================================================================
' ref_b164_semantics.vbs - reference semantics for the owned-temp reclaim line
' (research artifact for ledger #462; ASCII-only on purpose)
'
' WHY ASCII-ONLY: cscript reads a .vbs with the machine ANSI codepage. UTF-8
' Chinese comments get decoded as GBK on a zh-CN box (a GBK lead byte eats the
' following ASCII byte => the next line silently disappears) and as cp1252 on a
' CI runner. So the file is pure ASCII => every machine reads the same bytes.
'
' WHAT IT ANSWERS (the VB-family side only):
'   1) is a freshly built temp a VALUE or an ALIAS?  => "freshly built" claim
'   2) what TYPE does it carry?                      => which C type to declare
'   3) is it statement-scoped?                       => "free at statement end"
' The ownership side (who frees it) is NOT answered here: that is read from
' the RTL sources (output/_tmp/census/rtl_free_scan.py). Two sides, two
' references - do not let VB semantics vouch for C ownership.
'
' NOTE: VBScript has no IIf (that is VBA/VB6 only), so the branch-pick case
' is written as an explicit If. VB6's IIf evaluates both branches and keeps
' one - the same value semantics.
'
' RUN: cscript //NoLogo ref_b164_semantics.vbs   (exit 0 = all pass)
' =====================================================================

Dim nFail
nFail = 0

Sub Check(label, got, want)
    If CStr(got) = CStr(want) Then
        WScript.Echo "PASS " & label & " = " & CStr(got)
    Else
        WScript.Echo "FAIL " & label & " got=" & CStr(got) & " want=" & CStr(want)
        nFail = nFail + 1
    End If
End Sub

Dim a, b, t, s, v, v2

' ---- 1) value semantics: a built temp is a VALUE, not an alias ----------
a = "x"
b = a & "y"            ' the shape we are about to touch
a = "zzz"
Call Check("concat-value-not-alias", b, "xy")

a = "hello"
t = Left(a, 2)
a = "zzz"
Call Check("left-value-not-alias", t, "he")

a = 42
Call Check("cstr-value", CStr(a), "42")
t = CStr(a)
a = 7
Call Check("cstr-value-not-alias", t, "42")
Call Check("typename-value", TypeName(42), "Integer")

' ---- 2) types: what the consumer receives -------------------------------
Call Check("len-of-concat", Len("ab" & "cd"), 4)
Call Check("len-of-cstr", Len(CStr(1234)), 4)
Call Check("cdbl-value", CDbl("3.5"), 3.5)

' ---- 3) branch pick (VB6 IIf shape, written as If here) ------------------
a = "left"
b = "right"
If True Then
    t = a
Else
    t = b
End If
Call Check("branch-true", t, "left")
If False Then
    t = a
Else
    t = b
End If
Call Check("branch-false", t, "right")

' ---- 4) assignment copies: the right-hand value survives on its own -----
a = "x"
b = a & "y"
s = b
a = "zzz"
b = "qqq"
Call Check("assign-copies", s, "xy")

' ---- 5) variant boxing: CStr(v) is a new string, the slot is untouched ---
v2 = 12
Call Check("variant-cstr", CStr(v2), "12")
v2 = 13
Call Check("variant-cstr-after-change", CStr(v2), "13")

' ---- 6) numeric/string round trip (which C slot a value lands in) -------
Call Check("cstr-of-cdbl", CStr(CDbl("2.25")), "2.25")
Call Check("len-of-cstr-of-cdbl", Len(CStr(CDbl("2.25"))), 4)

' ---- 7) EVALUATION ORDER (the authority the reclaim knife's guards need) --
' The reclaim knife moves a minted value to the FRONT of the statement (it names
' it and frees it at the statement end).  That is only allowed if the order of
' the pieces around it cannot be observed.  This section measures the VB-family
' order with a logging function, so the guards have an EXTERNAL authority
' instead of a C-side assumption:
'   (a) the two operands of "&"      => left first, then right
'   (b) the arguments of a call      => left first, then right
' Both matter: (a) is what the concat station's left/right guards are about,
' (b) is what the "the other arguments have no side effect" guard is about.
Dim gLog

Function LogIt(tag, val)
    gLog = gLog & tag
    LogIt = val
End Function

Sub TwoArgs(x, y)
End Sub

gLog = ""
s = LogIt("L", "a") & LogIt("R", "b")
Call Check("concat-operand-order", gLog, "LR")

gLog = ""
Call TwoArgs(LogIt("1", "x"), LogIt("2", "y"))
Call Check("call-argument-order", gLog, "12")

gLog = ""
s = LogIt("A", "a") & LogIt("B", "b") & LogIt("C", "c")
Call Check("concat-chain-order", gLog, "ABC")

' ---- 8) "+" used as concatenation (the Fix 092i station) -----------------
' Same VB-family value and order semantics as "&": String + String is a new
' value, and the operands are evaluated left to right.  The reclaim knife now
' names the RIGHT operand of this station too (ledger #473), so both facts are
' pinned here: the value must not move, and the order must stay left-to-right
' whenever the left operand has an observable call.
a = "ab"
Call Check("add-concat-value", a + CStr(42), "ab42")

gLog = ""
s = LogIt("L", "a") + LogIt("R", "b")
Call Check("add-concat-operand-order", gLog, "LR")

gLog = ""
s = LogIt("A", "a") + LogIt("B", "b") + LogIt("C", "c")
Call Check("add-concat-chain-order", gLog, "ABC")

If nFail > 0 Then
    WScript.Echo "FAILURES = " & nFail
    WScript.Quit 1
Else
    WScript.Echo "ALL-PASS"
    WScript.Quit 0
End If
