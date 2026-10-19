' Acct #305 para-B175 (2026-10-14 knife 66): a numeric actual argument in a BSTR formal
' of an intrinsic that has NO typed parameter table.
'
' Why these eight names: they are registered in builtin_funcs.inc by addBuiltinFunc
' (return type only, no parameter table), so the Fix 175 station cannot ask "is this
' position a string formal" of calleeParams and used to hand the raw int32_t / double
' straight to a wchar_t* parameter.  cl only warns C4047 (outside -W3), the build
' succeeds, and the x64 run AVs on that line -- the same lesion class as para-B174
' (InStr) but on the registry family, which the real VB6 corpus writes with numeric
' actuals 461 times (para-B175 census: GetSetting default arm 369, SaveSetting
' setting arm 78).
'
' VB6 coerces there implicitly (CStr), so the pinned shape is: the actual is folded and
' the call answers as if the programmer had written the string.  ASCII comments only
' (C3 reads sources as ANSI/GBK).  Registry keys are created and removed again, and
' removed at the top too, so a crashed run cannot poison the next one.
Option Explicit

Private Const NB_APP As String = "C3ProbeB175"

' Emit-only witness, never called at run time (InputBox is modal): the OMITTED middle
' argument must keep its placeholder instead of being folded into the string "0".
' Real-corpus shape = ppProgressCircular.pag:351
'   Opacity = InputBox("enter the value ...", , 100)
' BASE emits 0L there (no title written -> the app title); the first cut of this knife
' wired the title slot into the fold and turned that 0L into vb6_CStrLong(0L), i.e. the
' caption "0".  The guard now also asks whether the user wrote this argument at all.
Private Sub OmitWitness()
    Dim w As Long
    w = InputBox("witness omitted title", , 100)
    Debug.Print w
End Sub

Public Sub Main()
    Dim nLen As Long
    Dim nSec As Long
    Dim nSet As Long
    Dim nForm As Long
    Dim dDef As Double
    Dim sKey As String
    Dim v As Variant
    Dim s7 As String
    Dim s8 As String

    nLen = 12
    nSec = 2
    nSet = 42
    nForm = 2
    dDef = 1.5
    sKey = "K1"

    DeleteSetting NB_APP, "Sec", sKey
    DeleteSetting NB_APP, CStr(nSec), sKey

    ' R1: the default arm is a BSTR formal -- VB6 answers CStr(default) when the key
    ' is absent; before this knife that slot got a raw 12 used as a pointer.
    Debug.Print "R1-default-long=" & GetSetting(NB_APP, "Sec", sKey, nLen)
    ' R2: same arm, Double actual
    Debug.Print "R2-default-dbl=" & GetSetting(NB_APP, "Sec", sKey, dDef)
    ' R3: the appname arm (a name that does not exist) -- proves slot 0 folds too
    Debug.Print "R3-app-numeric=" & GetSetting(nLen, "Sec", sKey, "d3")
    ' R4: write side with a numeric setting, then read it back
    SaveSetting NB_APP, "Sec", sKey, nSet
    Debug.Print "R4-write-read=" & GetSetting(NB_APP, "Sec", sKey, "none")
    ' R5: numeric SECTION name on the write side, read back through the literal "2"
    SaveSetting NB_APP, nSec, sKey, nLen
    Debug.Print "R5-num-section=" & GetSetting(NB_APP, "2", sKey, "none")
    ' R6: GetAllSettings with a numeric section -- array shape stays VB6's (8204)
    v = GetAllSettings(NB_APP, nSec)
    Debug.Print "R6-gas-vt=" & VarType(v)
    Debug.Print "R6-gas-cell=" & CStr(v(0, 1)) & "|" & CStr(v(0, 2))
    DeleteSetting NB_APP, CStr(nSec), sKey
    DeleteSetting NB_APP, "Sec", sKey
    ' R7/R8: a numeric Format style must behave exactly like the written-out string;
    ' Len() keeps the witness locale-robust while R7 pins the two arms are the same.
    s7 = Format(1234.567, nForm)
    s8 = Format(1234.567, "2")
    Debug.Print "R7-fmt-eq=" & (s7 = s8)
    Debug.Print "R8-fmt-len=" & Len(s7)
    Debug.Print "NB-DONE"
End Sub
