Option Explicit

' 11.214 (knife 231): a raise inside the value expression aborts the whole statement --
' a Variant target keeps its old value. Real machine (.temp/k225p/P325):
'   W1 w=CStr(Abs(Null)) -> [keep]/94 ; W3 w=Abs("x") -> [keep]/13 ; W4 w=Int("") -> [keep]/13
'   W5 w=CLng("x") -> [keep]/13 ; W6 w=1/0 -> [keep]/11 ; W7 w=v+1 -> [keep]/13
'   W10 w=CStr(Abs("x")) -> [keep]/13 ; W9 w=Hex(Null) -> []/0 (Null, no error)
' Before: everything but the late-bound COM reads fell through to a plain store -> "" (or 0).
Public Sub Main()
    Dim v As Variant
    Dim w As Variant
    Dim s As String
    Dim d As Double

    On Error Resume Next

    v = Null: Err.Clear: w = "keep": w = CStr(Abs(v))
    Debug.Print "Q1=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = "x": Err.Clear: w = "keep": w = Abs(v)
    Debug.Print "Q2=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = "": Err.Clear: w = "keep": w = Int(v)
    Debug.Print "Q3=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = "x": Err.Clear: w = "keep": w = CLng(v)
    Debug.Print "Q4=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    d = 0: Err.Clear: w = "keep": w = 1 / d
    Debug.Print "Q5=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = "x": Err.Clear: w = "keep": w = v + 1
    Debug.Print "Q6=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = Null: Err.Clear: w = "keep": w = Hex(v)
    Debug.Print "Q7=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = "x": Err.Clear: w = "keep": w = CStr(Abs(v))
    Debug.Print "Q8=[" & w & "]/" & Err.Number & "/" & IsNull(w)

    ' 没有抛错的那几格必须照写（守卫不许把它们也挡住）
    v = Null: Err.Clear: w = "keep": w = Null
    Debug.Print "Q9=[" & w & "]/" & Err.Number & "/" & IsNull(w)
    v = "2.5": Err.Clear: w = "keep": w = CDbl(v)
    Debug.Print "Q10=[" & w & "]/" & Err.Number & "/" & IsNull(w)

    ' 另外两档目标（String / Double）的守卫是既有行为，这里一并钉住
    v = Null: Err.Clear: s = "keep": s = CStr(v): Debug.Print "Q11=[" & s & "]/" & Err.Number
    v = "x": Err.Clear: s = "keep": s = CLng(v): Debug.Print "Q12=[" & s & "]/" & Err.Number
    d = 0: Err.Clear: s = "keep": s = 1 / d: Debug.Print "Q13=[" & s & "]/" & Err.Number
    v = "x": Err.Clear: d = -99: d = CLng(v): Debug.Print "Q14=[" & d & "]/" & Err.Number

    On Error GoTo 0
End Sub