Option Explicit

' §11.219 (第二百三十六刀): `Err.HelpFile` / `Err.HelpContext` 的读写两面，以及五参 `Err.Raise`。
' 真机读数 .temp/k227p/p331.txt（工程名 P331，本夹具的工程名是 B243Res）:
'   B1 赋值后能读、Number 仍是 0；B2 `Err.Raise 5`（缺参）保留两格；B3 `Err.Clear` 才清；
'   B4 `Err.Raise 5,"s1","d1","hf",7` 直接把后两格设进去（改前被截成三参 ⇒ 空串/0）；
'   B6 只给 context 时 file 仍空；B8 运行期 13 之后只改 context，号与另一格都不动。
Public Sub Main()
    On Error Resume Next

    Err.HelpFile = "h1"
    Err.HelpContext = 42
    Debug.Print "B1=[" & Err.HelpFile & "]/" & Err.HelpContext & "/" & Err.Number

    Err.Raise 5
    Debug.Print "B2=[" & Err.HelpFile & "]/" & Err.HelpContext & "/" & Err.Number
    Err.Clear
    Debug.Print "B3=[" & Err.HelpFile & "]/" & Err.HelpContext & "/" & Err.Number

    Err.Raise 5, "s1", "d1", "hf", 7
    Debug.Print "B4=[" & Err.HelpFile & "]/" & Err.HelpContext & "/" & Err.Number & "/[" & Err.Source & "]/[" & Err.Description & "]"
    Err.Clear
    Debug.Print "B5=[" & Err.HelpFile & "]/" & Err.HelpContext & "/" & Err.Number

    Err.HelpContext = 9
    Debug.Print "B6=[" & Err.HelpFile & "]/" & Err.HelpContext & "/" & Err.Number
    Err.Clear

    Err.HelpFile = ""
    Err.HelpContext = 0
    Debug.Print "B7=[" & Err.HelpFile & "]/" & Err.HelpContext & "/" & Err.Number
    Err.Clear

    Dim n As Long
    n = CLng("x")
    Err.HelpContext = 3
    Debug.Print "B8=" & Err.Number & "/[" & Err.HelpFile & "]/" & Err.HelpContext
    Err.Clear

    Debug.Print "B9=[" & Err.HelpFile & "]/" & Err.HelpContext & "/" & Err.Number
    On Error GoTo 0
End Sub