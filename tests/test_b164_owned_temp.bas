Attribute VB_Name = "B164OwnedTemp"
' Ledger #464 (B164) regression fixture: "a freshly built owned temp in one statement".
'
' What this knife does (the whole point of the fixture):
'   A minted BSTR/VARIANT that is passed to a call as a WHOLE argument and that nobody
'   keeps was leaked. The knife now names it (`_own_N`), and frees it at the end of the
'   statement, in the right KIND (BSTR -> vb6_BSTR_Free, VARIANT -> vb6_VariantClear).
'
' The three gates that decide whether a site is nominated (all in
' src/common/owned_temp_keeps.hpp, derived from the RTL by check_owned_temp_keeps.ps1):
'   1. the mint must be a FRESH value we own (not borrowed, not an alias, not one of the
'      four exits whose return value is one of their own parameters);
'   2. the consumer must not KEEP the pointer, must not be in the closed family, and must
'      not FREE the argument itself (that would be a double free);
'   3. the type comes from the mint exit's own RTL return type, never from the consumer's
'      VB parameter type (that mistake is what broke #461 with cl C2440).
'
' The readings below must match on x64 and x86 - the values must NOT move.
Option Explicit

Public Sub Main()
    Dim a As String
    Dim b As String
    Dim x As String
    Dim n As Long
    Dim i As Long
    Dim bad10 As Long
    Dim bad12 As Long

    a = "ab"
    b = "cd"
    n = 42

    ' (1) whole-argument mint: Len(a & b) -> the concat result is named and freed
    Debug.Print "T1-len=" & Len(a & b)
    ' (2) the concat's RIGHT operand: "a" & CStr(n) -> the CStr result is named and freed
    Debug.Print "T2-right=" & (a & CStr(n))
    ' (3) two-argument call, the other argument is a literal (pure) -> nominated
    Debug.Print "T3-strcmp=" & StrComp(a & b, "abcd")
    ' (4) two-argument call, the mint is in slot 0, the other argument is a literal
    Debug.Print "T4-left=" & Left(a & b, 2)
    ' (5) the mint sits INSIDE the concat's left operand -> the concat station's own
    '     judgment reads that text ("does it contain vb6_BSTR_Concat?") => naming it
    '     would flip that judgment => the value must stay as it is (readings must not move)
    x = Len(a & b) & "!"
    Debug.Print "T5-trap=" & x
    ' (6) the mint is the whole right-hand side of an assignment -> that is the
    '     assignment station's business (it moves the value), not ours
    x = a & b
    Debug.Print "T6-assign=" & x
    ' (7) chained concats
    Debug.Print "T7-chain=" & (a & b & "!" & CStr(n))
    ' (8) a mint inside a bigger expression
    Debug.Print "T8-mix=" & (Len(a & b) + Len(CStr(n)))
    ' (9) the concat station's LEFT operand is itself a freshly built owned temp
    '     => the station frees it itself (that is what ledger #469 §B164 adds).
    '     The value must not move; and if it were freed twice this line would crash.
    x = "ab" & b
    Debug.Print "T9-concatleft=" & x
    ' (10) the `+` station used as concatenation (Fix 092i) must reclaim its RIGHT
    '      operand too, exactly like the `&` station (ledger #473 §B164). The left
    '      operand is a plain variable here (no call) so gate 10 lets it through.
    x = a + CStr(n)
    Debug.Print "T10-addright=" & x
    ' (10b) stress the new reclaim: a double free or use-after-free would crash or
    '       corrupt the value inside this loop (2000 statement scopes, same shapes).
    bad10 = 0
    For i = 1 To 2000
        x = a + CStr(i)
        If x <> ("ab" & CStr(i)) Then bad10 = bad10 + 1
    Next i
    Debug.Print "T10-loop-bad=" & bad10
    ' (11) gate 10 negative control: the LEFT operand of `+` contains a call, so the
    '      right mint must NOT be hoisted (the emit shape must keep both calls inline).
    '      The numeric tags avoid a second thing being hoisted (BSTR literal args of a
    '      call are minted too), so the shape pin below is about the concat station only.
    '      NOTE on the log: C3/MSVC evaluates the two arguments of vb6_BSTR_Concat
    '      right-to-left, so this run reads "21". The VBS reference says the VB family
    '      order is left-to-right ("12"); that divergence predates this knife and is
    '      recorded separately. Pinning "21" here only says the reclaim knife did not
    '      move it -- it must be re-pinned when the order itself is fixed.
    mOrder164 = ""
    x = Mark164(1) + CStr(MarkNum164(2, 7))
    Debug.Print "T11-value=" & x
    Debug.Print "T11-log=" & mOrder164
    ' (12) negative control for the COUPLING of the two `+`-station rules: the LEFT
    '      operand is itself a concat (an owned BSTR mint), so gate 10 sees a call in
    '      the left text and refuses to hoist the RIGHT mint. The station still frees
    '      the left exactly once (ConcatFree) and the right stays inline, exactly like
    '      T11. This pins the current gate on purpose: when the gate learns to tell a
    '      pure mint from a real call, this negative control must be re-pinned to the
    '      reclaimed shape. A double free of the left would crash / change the value.
    x = (a & b) + CStr(n)
    Debug.Print "T12-both=" & x
    bad12 = 0
    For i = 1 To 500
        x = (a & b) + CStr(i)
        If x <> ("abcd" & CStr(i)) Then bad12 = bad12 + 1
    Next i
    Debug.Print "T12-loop-bad=" & bad12

    Debug.Print "B164-DONE"
End Sub

Private mOrder164 As String

Private Function Mark164(ByVal tag As Long) As String
    mOrder164 = mOrder164 & tag
    Mark164 = "A"
End Function

Private Function MarkNum164(ByVal tag As Long, ByVal v As Long) As Long
    mOrder164 = mOrder164 & tag
    MarkNum164 = v
End Function
