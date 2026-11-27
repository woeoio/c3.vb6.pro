Option Explicit

' 11.208 (knife 225): Round on a *Variant carrier* keeps the carrier's subtype.
' Oracle .temp/k224p/P298 (vI..vN) and P300 (V1..V7 column 3): Integer 2 / Long 3 /
' Single 4 / Currency 6 / Byte 17 / Date 7 / Boolean 2 / Empty 2 / Null 1, and the
' banker's rounding of 11.206 still holds through the variant door.
'
' The reroute lives in boxToVariant (the k64 station); before it, the emitted C was
' vb6_Round(vb6_VariantToDouble(v), 0) -> _Generic -> VT_R8 (5) for every carrier.
'
' Cells are VarType numbers / small integers / "True"/"False" / fixed decimals, so
' nothing here depends on the locale.
Public Sub Main()
    Dim v As Variant

    v = CInt(3):      Debug.Print "N1=" & VarType(Round(v)) & "/" & Round(v)
    v = CLng(70000):  Debug.Print "N2=" & VarType(Round(v)) & "/" & Round(v)
    v = CSng(2.5):    Debug.Print "N3=" & VarType(Round(v)) & "/" & Round(v)
    v = CCur(2.5):    Debug.Print "N4=" & VarType(Round(v)) & "/" & Round(v)
    v = CByte(7):     Debug.Print "N5=" & VarType(Round(v)) & "/" & Round(v)
    v = CDate(44259.62): Debug.Print "N6=" & VarType(Round(v)) & "/" & Day(Round(v))
    v = CBool(True):  Debug.Print "N7=" & VarType(Round(v)) & "/" & Round(v)
    v = Empty:        Debug.Print "N8=" & VarType(Round(v)) & "/" & Round(v)

    v = Null:         Debug.Print "N9=" & IsNull(Round(v)) & "/" & VarType(Round(v))

    ' 值仍然是 11.206 的半到偶；位数 0..22 的位域门不变
    v = 2.5:          Debug.Print "M1=" & Round(v) & "/" & Round(v, 0) & "/" & Round(v, 2)
    v = 0.5:          Debug.Print "M2=" & Round(v)
    v = 2.675:        Debug.Print "M3=" & Round(v, 2)
    v = "2.9":        Debug.Print "M4=" & VarType(Round(v)) & "/" & Round(v)
    v = CInt(-5):     Debug.Print "M5=" & Round(v) & "/" & Round(v, 1)

    ' 直接档（不经载体）仍然是实参子类型 —— 11.207 那一刀不许被本刀碰坏
    Debug.Print "K1=" & VarType(Round(CInt(3))) & VarType(Round(CLng(3))) & VarType(Round(CCur(3)))
End Sub