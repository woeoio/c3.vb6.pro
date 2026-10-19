Attribute VB_Name = "B176VarSlot"
' Ledger B176 (knife 67) regression fixture: a Variant actual handed to a formal of
' one of the builtins that have NO typed parameter table. Two shapes, both measured:
'   shape 1  Variant -> BSTR formal = the raw vb6_VARIANT goes into a wchar_t* slot
'            -> cl C2440 on BOTH arches, no exe (16 such errors on this file).
'   shape 2  Variant -> int32_t formal = the runtime C-signature table keys its rows
'            by the arity VB WROTE while this arm asks by the EMITTED position, so a
'            3-arg InStr shifts a whole slot: the string goes into 'int32_t start',
'            cl only warns, and at run time start is a truncated pointer -> InStr
'            answers 0 silently.
' Both shapes are now answered by one authority (builtinNoParamTableFormalCType).

' The dialog-popping forms (MsgBox / InputBox / SendKeys) and SavePicture live in the
' never-called witness sub at the bottom: their shapes are pinned by Test-EmitcShape /
' Test-EmitcAbsent in run_tests.ps1. The readings below must match on x64 and x86.
Option Explicit

Private Const BV_APP As String = "C3ProbeB176"

Public Sub Main()
    Dim vKey As Variant
    Dim vSec As Variant
    Dim vSet As Variant
    Dim vDef As Variant
    Dim vS4 As Variant
    Dim vS5 As Variant
    Dim vStyle As Variant
    Dim dNum As Variant
    Dim s1 As String
    Dim got As Variant

    vKey = "K1"
    vSec = "Sec"
    vSet = "roundtrip"
    vDef = "unused-default"
    vS4 = 4
    vS5 = 5
    vStyle = "0"
    dNum = 1234.567
    s1 = "abc7def"

    ' shape 2: InStr start slot takes a Variant (before: folded to a string)
    Debug.Print "V1-start4=" & InStr(vS4, s1, "7")
    Debug.Print "V2-start5=" & InStr(vS5, s1, "7")
    Debug.Print "V3-start2=" & InStr(2, s1, "7")
    ' witness: a TWO-arg InStr has no start slot -- the haystack must still unwrap
    Debug.Print "V4-twoarg=" & InStr(vKey, "text")

    ' shape 1: Variant into a BSTR formal (registry family, real round trip)
    DeleteSetting BV_APP, vSec, vKey
    SaveSetting BV_APP, vSec, vKey, vSet
    got = GetSetting(BV_APP, vSec, vKey, vDef)
    Debug.Print "V5-roundtrip=" & got
    Debug.Print "V6-vartype=" & VarType(got)
    ' a different key must miss -> the default comes back: proves the key slot
    Debug.Print "V7-miss=" & GetSetting(BV_APP, vSec, "no-such-key-zz", vSet)
    DeleteSetting BV_APP, vSec, vKey

    ' shape 1: Format style slot takes a Variant (before: cl C2440)
    Debug.Print "V8-fmt=" & Format(dNum, vStyle)
    Debug.Print "V9-fmt-lit=" & Format(dNum, "0")
    Debug.Print "V10-fmt-len=" & Len(Format(dNum, vStyle))

    Debug.Print "BV-DONE"
End Sub

Private Sub VariantWitness()
    ' Never called: these shapes exist only for the emit-c needles in run_tests.ps1
    ' (MsgBox / InputBox pop a dialog, SendKeys types into the foreground window,
    '  SavePicture needs a bitmap -- none of them may run).
    Dim vP As Variant
    Dim vBtn As Variant
    Dim vTitle As Variant
    Dim vKeys As Variant
    Dim vWait As Variant
    Dim vName As Variant
    Dim vAttr As Variant
    Dim pic As Object
    Dim vS3 As Variant
    Dim vNeedle As Variant
    Dim vCmp As Variant
    Dim w1 As String
    vP = "hi"
    vBtn = 1
    vTitle = "t"
    vKeys = "a"
    vWait = 0
    vName = "f.txt"
    vAttr = 32
    vS3 = 3
    vNeedle = "7"
    vCmp = 0
    w1 = "abc7def"
    MsgBox vP, vBtn, vTitle
    Debug.Print InputBox(vP, vTitle, vP)
    SendKeys vKeys, vWait
    SetAttr vName, vAttr
    SavePicture pic, vName
    Debug.Print InStr(vS3, w1, vNeedle)
    Debug.Print InStr(vS3, w1, "7", vCmp)
End Sub
