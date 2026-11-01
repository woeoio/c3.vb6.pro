' Knife 318 (ledger #594): Select Case over a user function that returns Variant,
' in a SINGLE-MODULE project. Single-module projects emit bare C names vb6_<proc>,
' and the driver pre-scan used to register only vb6_<module>_<proc> / vb6_pv_<proc>
' in variantReturnFuncs. So cExprIsVariant answered "not a slot" and the Select Case
' subject was stored as a plain scalar:
'     BSTR _vb6_select_0 = vb6_VPick();      -> cl C2440 (vb6_VARIANT -> BSTR)
' Fix: the driver also registers the bare vb6_<proc> names whenever the project has
' exactly one code module. Expected output, per loop round, is:
'     B594-VS-1=abc
'     B594-VS-2=2
'     B594-VS-3=pv
Option Explicit

Public Function VPick() As Variant
    VPick = "abc"
End Function

Public Function NPick() As Variant
    NPick = 2
End Function

Private Function PPick() As Variant
    PPick = "pv"
End Function

Sub Main()
    Dim i As Long
    For i = 1 To 3
        Select Case VPick()
            Case "abc"
                Debug.Print "B594-VS-1=abc"
            Case Else
                Debug.Print "B594-VS-1=else"
        End Select
        Select Case NPick()
            Case 2
                Debug.Print "B594-VS-2=2"
            Case Else
                Debug.Print "B594-VS-2=else"
        End Select
        Select Case PPick()
            Case "pv"
                Debug.Print "B594-VS-3=pv"
            Case Else
                Debug.Print "B594-VS-3=else"
        End Select
    Next i
    Debug.Print "B594-DONE"
End Sub
