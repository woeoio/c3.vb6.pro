Option Explicit

' §11.218 (第二百三十五刀): `Err.Source = v` / `Err.Description = v` 的写面。
' 真机读数 .temp/k226p/p330.txt（工程名 P330，本夹具的工程名是 B242Res）:
'   A1 两格赋值后读回、Number 仍是 0；A2 `Err.Raise 5`（省掉源与描述）**不清格**（[s1]/[d1]）；
'   A3 `Err.Clear` 之后全空；A4/A5 只赋一格时另一格仍空；A6 raise 之后覆盖 source、描述不动；
'   A7 赋空串就是空串；A8 运行期 13 之后赋描述，号与源都还在；A9 raise 缺参时源补工程名、
'   描述保留；A10 两格都赋。
Public Sub Main()
    On Error Resume Next

    Err.Source = "s1"
    Err.Description = "d1"
    Debug.Print "A1=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    Err.Raise 5
    Debug.Print "A2=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    Err.Clear
    Debug.Print "A3=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number

    Err.Description = "only desc"
    Debug.Print "A4=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    Err.Clear
    Err.Source = "only src"
    Debug.Print "A5=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    Err.Clear

    Err.Raise 5, "rs", "rd"
    Err.Source = "over"
    Debug.Print "A6=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    Err.Clear

    Err.Source = ""
    Err.Description = ""
    Debug.Print "A7=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    Err.Clear

    Dim n As Long
    n = CLng("x")
    Err.Description = "after13"
    Debug.Print "A8=" & Err.Number & "/[" & Err.Source & "]/[" & Err.Description & "]"
    Err.Clear

    Err.Description = "x"
    Err.Raise 5
    Debug.Print "A9=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    Err.Clear

    Err.Source = "ss"
    Err.Description = "dd"
    Debug.Print "A10=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    On Error GoTo 0
End Sub