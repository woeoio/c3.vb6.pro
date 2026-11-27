Option Explicit

' 11.207 (knife 224): Abs / Int / Fix / Round answer the *argument's* numeric subtype.
' Oracle .temp/k224p/P298 / P300 / P302 (VB6):
'   Integer->2  Long->3  Single->4  Currency->6  Byte->17  Double->5  Date->7
'   Boolean argument -> Integer (2), value -1
'   Variant carrier: Int/Fix of a Long carrier is still Long, Byte still Byte, Empty 2.
' Still open after this knife (next knife): Round's *Variant carrier* exit
'   (Round(v) where v holds Integer should answer 2, C3 still answers 5) — the static
'   column and Int/Fix's carrier column are the ones closed here.
'
' All cells are VarType numbers / small integers / "(-1)", so nothing is locale-dependent.
Public Sub Main()
    Dim i As Integer
    Dim l As Long
    Dim s As Single
    Dim c As Currency
    Dim b As Byte
    Dim d As Double
    Dim t As Date
    Dim v As Variant
    i = 3: l = 70000: s = 2.5: c = 2.5: b = 7: d = 2.5: t = CDate(44259.62)

    Debug.Print "S1=" & CStr(VarType(Abs(i))) & CStr(VarType(Abs(l))) & CStr(VarType(Abs(s))) & CStr(VarType(Abs(c))) & CStr(VarType(Abs(b))) & CStr(VarType(Abs(d)))
    Debug.Print "S2=" & CStr(VarType(Int(i))) & CStr(VarType(Int(l))) & CStr(VarType(Int(s))) & CStr(VarType(Int(c))) & CStr(VarType(Int(b))) & CStr(VarType(Int(d)))
    Debug.Print "S3=" & CStr(VarType(Fix(i))) & CStr(VarType(Fix(l))) & CStr(VarType(Fix(s))) & CStr(VarType(Fix(c))) & CStr(VarType(Fix(b))) & CStr(VarType(Fix(d)))
    Debug.Print "S4=" & CStr(VarType(Round(i))) & CStr(VarType(Round(l))) & CStr(VarType(Round(s))) & CStr(VarType(Round(c))) & CStr(VarType(Round(b))) & CStr(VarType(Round(d)))

    ' Boolean 实参 → Integer（真机 D1..D3 = 2/2/2，值 -1）
    Debug.Print "T1=" & CStr(VarType(Int(CBool(True)))) & CStr(VarType(Fix(CBool(True)))) & CStr(VarType(Round(CBool(True))))
    Debug.Print "T2=[" & CStr(Int(CBool(True))) & "]"

    ' Date 实参仍是 Date
    Debug.Print "U1=" & CStr(VarType(Int(t))) & CStr(VarType(Round(t)))

    ' 变体载体：Int/Fix 保子类型（Round 那一列是下一刀的账）
    v = l: Debug.Print "V1=" & CStr(VarType(Int(v))) & CStr(VarType(Fix(v)))
    v = b: Debug.Print "V2=" & CStr(VarType(Int(v))) & CStr(VarType(Fix(v)))
    v = Empty: Debug.Print "V3=" & CStr(VarType(Int(v))) & CStr(VarType(Fix(v)))
    v = Null: Debug.Print "V4=" & CStr(VarType(Int(v))) & CStr(VarType(Fix(v)))
    v = s: Debug.Print "V5=" & CStr(VarType(Int(v))) & CStr(VarType(Fix(v)))

    ' 值本身不变（Integer 载体的 Int/Fix 不提升成 Long）
    v = CInt(-5): Debug.Print "W1=[" & CStr(Int(v)) & "][" & CStr(Fix(v)) & "]"
    Debug.Print "W2=[" & CStr(Int(l)) & "][" & CStr(Round(2.5)) & "]"
End Sub