Option Explicit

' 11.206 (knife 223): Round is half-to-even (banker's), and numdecimalplaces is 0..22.
' Oracle .temp/k223p/P290..P292 (VB6):
'   Round(0.5)=0 | Round(2.5)=2 | Round(4.5)=4 | Round(-0.5)=0 | Round(-2.5)=-2
'   Round(0.125,2)=0.12 | Round(0.375,2)=0.38 | Round(2.675,2)=2.68 | Round(1.005,2)=1
'   Round(2.5,-1) / Round(1.5,23) => error 5 ; Round(1.5,22) is fine
' The "multiply then halve-to-even" shape reproduces both 2.675 -> 2.68 and 1.005 -> 1,
' which a decimal-string rounding cannot do (it would give 1.01 for 1.005).
'
' Still open (next knife, type identity): VB6 answers Integer/Long/Single/Byte/Currency
' for the matching argument subtype and Null/Integer for Null/Empty; C3 answers Double.
'
' Cells are plain numbers ("." decimal point is locale-independent here).
Public Sub Main()
    Debug.Print "R1=[" & CStr(Round(0.5)) & "]"
    Debug.Print "R2=[" & CStr(Round(2.5)) & "]"
    Debug.Print "R3=[" & CStr(Round(4.5)) & "]"
    Debug.Print "R4=[" & CStr(Round(-0.5)) & "] len=" & CStr(Len(CStr(Round(-0.5))))
    Debug.Print "R5=[" & CStr(Round(-2.5)) & "]"
    Debug.Print "R6=[" & CStr(Round(0.125, 2)) & "]"
    Debug.Print "R7=[" & CStr(Round(0.375, 2)) & "]"
    Debug.Print "R8=[" & CStr(Round(2.675, 2)) & "]"
    Debug.Print "R9=[" & CStr(Round(1.005, 2)) & "]"
    Debug.Print "RA=[" & CStr(Round(1.2345, 3)) & "]"
    Debug.Print "RB=[" & CStr(Round(0.05, 1)) & "]"
    Debug.Print "RC=[" & CStr(Round(0.15, 1)) & "]"
    Debug.Print "RD=[" & CStr(Round(1235.5)) & "]"
    Debug.Print "RE=[" & CStr(Round(2.5000000000000004, 0)) & "]"
    Debug.Print "RF=[" & CStr(Round(2.4999999999999996, 0)) & "]"
    Debug.Print "RG=[" & CStr(Round(CCur(2.5))) & "]"

    On Error Resume Next
    Dim v As Variant
    v = Empty: Err.Clear: v = Round(2.5, -1)
    Debug.Print "S1 err=" & CStr(Err.Number)
    v = Empty: Err.Clear: v = Round(1.5, 23)
    Debug.Print "S2 err=" & CStr(Err.Number)
    v = Empty: Err.Clear: v = Round(1.5, 22)
    Debug.Print "S3 err=" & CStr(Err.Number) & " v=[" & CStr(v) & "]"
End Sub