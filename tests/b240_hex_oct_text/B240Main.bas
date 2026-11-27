Option Explicit

' §11.216 (第二百三十三刀): Hex/Oct 收到「文字档」的整只 Variant 槽 —— 文本能折数就折，
' 折不成就 13，且目标一个字不动。真机读数 .temp/k225p/P328（十四格）:
'   H1 "x" ⇒ keep/13、H3 "" ⇒ keep/13、H2 "12" ⇒ C、O2 "9" ⇒ 11、H4/L2 Null ⇒ Null 不报错。
Public Sub Main()
    Dim v As Variant
    Dim w As Variant
    Dim s As String
    Dim n As Long
    On Error Resume Next

    v = "x": Err.Clear: w = "keep": w = Hex(v): Debug.Print "H1=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = "12": Err.Clear: w = "keep": w = Hex(v): Debug.Print "H2=[" & w & "]/" & Err.Number
    v = "": Err.Clear: w = "keep": w = Hex(v): Debug.Print "H3=[" & w & "]/" & Err.Number
    v = Null: Err.Clear: w = "keep": w = Hex(v): Debug.Print "H4=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = 12.7: Err.Clear: w = "keep": w = Hex(v): Debug.Print "H5=[" & w & "]/" & Err.Number
    v = True: Err.Clear: w = "keep": w = Hex(v): Debug.Print "H6=[" & w & "]/" & Err.Number
    v = "x": Err.Clear: w = "keep": w = Oct(v): Debug.Print "O1=[" & w & "]/" & Err.Number
    v = "9": Err.Clear: w = "keep": w = Oct(v): Debug.Print "O2=[" & w & "]/" & Err.Number

    v = "x": Err.Clear: s = "keep": s = Hex(v): Debug.Print "S1=[" & s & "]/" & Err.Number
    v = "x": Err.Clear: n = -99: n = Hex(v): Debug.Print "N1=[" & n & "]/" & Err.Number
    v = "12": Err.Clear: n = -99: n = Hex(v): Debug.Print "N2=[" & n & "]/" & Err.Number
    v = "x": Err.Clear: s = "keep": s = "a" & Hex(v): Debug.Print "C1=[" & s & "]/" & Err.Number

    Err.Clear: w = "keep": w = Hex("x"): Debug.Print "L1=[" & w & "]/" & Err.Number
    Err.Clear: w = "keep": w = Hex(Null): Debug.Print "L2=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    On Error GoTo 0
End Sub