Option Explicit

' §11.215 (第二百三十二刀): CStr(<整只 Variant 槽>) 的 Null 面 —— 不论消费者是谁都是 94，
' 且值消费得一个字不剩。真机读数 .temp/k225p/P327（V1\V5、D1\D4、N1、C1\C3 全是 94）。
Public Sub Main()
    Dim v As Variant
    Dim w As Variant
    Dim s As String
    Dim d As Double
    Dim n As Long
    On Error Resume Next

    v = Null: Err.Clear: w = "keep": w = CStr(v): Debug.Print "V1=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = Empty: Err.Clear: w = "keep": w = CStr(v): Debug.Print "V2=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = "x": Err.Clear: w = "keep": w = CStr(v): Debug.Print "V3=[" & w & "]/" & Err.Number
    v = 1.5: Err.Clear: w = "keep": w = CStr(v): Debug.Print "V4=[" & w & "]/" & Err.Number
    Err.Clear: w = "keep": w = CStr(Null): Debug.Print "V5=[" & w & "]/" & Err.Number

    v = Null: Err.Clear: d = -99: d = CStr(v): Debug.Print "D1=[" & d & "]/" & Err.Number
    v = Empty: Err.Clear: d = -99: d = CStr(v): Debug.Print "D2=[" & d & "]/" & Err.Number
    v = "2.5": Err.Clear: d = -99: d = CStr(v): Debug.Print "D3=[" & d & "]/" & Err.Number
    Err.Clear: d = -99: d = CStr(Null): Debug.Print "D4=[" & d & "]/" & Err.Number

    v = Null: Err.Clear: s = "keep": s = CStr(v): Debug.Print "S1=[" & s & "]/" & Err.Number
    v = Empty: Err.Clear: s = "keep": s = CStr(v): Debug.Print "S2=[" & s & "]/" & Err.Number

    v = Null: Err.Clear: n = -99: n = CStr(v): Debug.Print "N1=[" & n & "]/" & Err.Number
    v = Empty: Err.Clear: n = -99: n = CStr(v): Debug.Print "N2=[" & n & "]/" & Err.Number
    v = "7": Err.Clear: n = -99: n = CStr(v): Debug.Print "N3=[" & n & "]/" & Err.Number

    v = Null: Err.Clear: s = "keep": s = "a" & CStr(v): Debug.Print "C1=[" & s & "]/" & Err.Number
    v = Null: Err.Clear: s = "keep": s = Len(CStr(v)): Debug.Print "C2=[" & s & "]/" & Err.Number
    v = Null: Err.Clear: w = "keep": w = CStr(v) & "": Debug.Print "C3=[" & w & "]/" & Err.Number
    On Error GoTo 0
End Sub