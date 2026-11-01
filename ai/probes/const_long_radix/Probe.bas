Attribute VB_Name = "Probe"
Option Explicit

' ai/probes/const_long_radix - the one reading only real VB6 can give (C3 account #305, knife 313).
'
' Already answered by the same-machine VBScript engine (kept below as a control):
'   a bare &H8000 is -32768  (the 16-bit two's-complement tier).
' NOT answerable by VBScript - it has no typed Const:
'   when the same literal sits in an "As Long" type context, does VB6 widen it
'   BY VALUE (-32768) or BY BITS (32768)?
' The real corpus has 58 sites of the shape  Const ... As Long = &Hxxxx  in that band
' (czUI MEM_RELEASE, DT_END_ELLIPSIS, CLR_INVALID, H_MAX = &HFFFF + 1,
'  Optional ToPage As Long = &HFFFF ...), and the two answers differ at runtime there.
'
' How to run: open ConstRadixProbe.vbp in VB6, press F5, and paste the NUMBERED lines back.
' (It also prints to the immediate window.) If any of lines 2-8 does not COMPILE, paste the
' IDE's message over that number instead - "it does not compile" is itself the answer.

Public Const A_L8000 As Long = &H8000             ' q2
Public Const B_LFFFF As Long = &HFFFF             ' q3
Public Const C_LFFFF1 As Long = &HFFFF + 1        ' q4
Public Const D_L8000S As Long = &H8000&           ' q5
Public Const E_I8000 As Integer = &H8000          ' q7
Public Const F_L32B As Long = &HFFFFFFF0          ' q8

Public Function GRet() As Long                    ' q9: return-type context
    GRet = &H8000
End Function

Public Function Fmt(ByVal v As Long) As String    ' q10: ByVal Long parameter
    Fmt = CStr(v)
End Function

Public Function Opt(Optional ByVal t As Long = &HFFFF) As Long   ' q12: omitted-argument default
    Opt = t
End Function

Sub Main()
    Dim s As String
    Dim x As Long

    x = &H8000
    s = "1)  bare &H8000                  = " & CStr(&H8000) & vbCrLf
    s = s & "2)  Const A As Long = &H8000   = " & CStr(A_L8000) & vbCrLf
    s = s & "3)  Const B As Long = &HFFFF   = " & CStr(B_LFFFF) & vbCrLf
    s = s & "4)  Const C As Long = &HFFFF+1 = " & CStr(C_LFFFF1) & vbCrLf
    s = s & "5)  Const D As Long = &H8000&  = " & CStr(D_L8000S) & vbCrLf
    s = s & "6)  Dim x As Long: x = &H8000  = " & CStr(x) & vbCrLf
    s = s & "7)  Const E As Integer         = " & CStr(E_I8000) & vbCrLf
    s = s & "8)  Const F As Long = &HFFFFFFF0 = " & CStr(F_L32B) & vbCrLf
    s = s & "9)  Function() As Long         = " & CStr(GRet()) & vbCrLf
    s = s & "10) passed to ByVal Long       = " & Fmt(&H8000) & vbCrLf
    s = s & "11) x And &HF0000   (x from 6) = " & CStr(x And &HF0000) & vbCrLf
    s = s & "12) Optional As Long = &HFFFF  = " & CStr(Opt()) & vbCrLf
    s = s & "13) VarType(A from 2)          = " & CStr(VarType(A_L8000)) & vbCrLf
    s = s & "14) VarType(bare &H8000)       = " & CStr(VarType(&H8000)) & vbCrLf
    s = s & "15) witness: x = &H8000& then  = "
    x = &H8000&
    s = s & CStr(x) & vbCrLf
    Debug.Print s
    MsgBox s, vbOKOnly, "const_long_radix probe"
End Sub
