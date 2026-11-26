Option Explicit

' §11.191 (knife 208): a Date compared with a Variant. Oracle: VB6Mini /make
' .temp/k208p/P249.bas -> p249.txt, eight lines:
'   Z1 = 1  (Variant holding text "2020-12-31": parsed as a date, d is greater)
'   Z2 = 1  (Variant holding text equal to d)
'   Z3 = 0  (Variant holding 45000: plain serial, not equal to 44259)
'   Z4 = 0 err=0  (Variant holding "abc": answer false, NO error raised)
'   Z5 = 0 err=0  (Variant holding Null: answer false, NO error raised)
'   Z6 = 1  (Variant holding Empty: behaves as 0)
'   Z7 = 1  (Date ARRAY ELEMENT < Variant holding a date)
'   Z8 = 1  (Variant on the left, Date on the right)
' Before this knife the Date was boxed as VT_R8, so the RTL applied its
' "number < string" rule (Z1/Z2/Z8 answered 0) and Null was unpacked as 0
' (Z5 answered 1).
Public Sub Main()
    Dim d As Date
    Dim v As Variant
    Dim q() As Date

    On Error Resume Next
    ReDim q(0 To 1)
    d = CDate("2021-03-04")
    q(0) = d

    v = "2020-12-31"
    Err.Clear
    If d > v Then
        Debug.Print "Z1 gt=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Z1 gt=0 err=" & CStr(Err.Number)
    End If

    v = "2021-03-04"
    Err.Clear
    If d = v Then
        Debug.Print "Z2 eq=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Z2 eq=0 err=" & CStr(Err.Number)
    End If

    v = 45000
    Err.Clear
    If d = v Then
        Debug.Print "Z3 num=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Z3 num=0 err=" & CStr(Err.Number)
    End If

    v = "abc"
    Err.Clear
    If d > v Then
        Debug.Print "Z4 bad=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Z4 bad=0 err=" & CStr(Err.Number)
    End If

    v = Null
    Err.Clear
    If d > v Then
        Debug.Print "Z5 null=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Z5 null=0 err=" & CStr(Err.Number)
    End If

    v = Empty
    Err.Clear
    If d > v Then
        Debug.Print "Z6 empty=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Z6 empty=0 err=" & CStr(Err.Number)
    End If

    v = CDate("2021-03-05")
    Err.Clear
    If q(0) < v Then
        Debug.Print "Z7 elem=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Z7 elem=0 err=" & CStr(Err.Number)
    End If

    v = "2021-03-04"
    Err.Clear
    If v = d Then
        Debug.Print "Z8 rev=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Z8 rev=0 err=" & CStr(Err.Number)
    End If
End Sub