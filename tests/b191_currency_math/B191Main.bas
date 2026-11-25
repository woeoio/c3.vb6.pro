Option Explicit

' 账 #305 §11.166 (第一百八十三刀): Currency 经 Abs/Int/Fix/Round 保持 Currency 身份。
'
' 真 VB6 探针 .temp/k183p (P191.bas) 读数（四列: VarType | TypeName | CStr | CLng(x*10000)）:
'   M1  Abs(c)        -> 6 Currency 10.1235      M2  Int(c)   -> 6 Currency 10
'   M3  Fix(c)        -> 6 Currency 10           M4  Round(c,2)-> 6 Currency 10.12
'   M5  Sgn(c)        -> 2 Integer 1             M8  -c       -> 6 Currency -10.1235
'   M6  Abs(-c)       -> 6 Currency 10.1235      M7  Int(-c)  -> 6 Currency -11
'   M9  Fix(-c)       -> 6 Currency -10          M10 Round(-c,2)-> 6 Currency -10.12
'   M11 Int(v)  (v 装 Currency) -> 6 Currency 10
'   M12 Abs(v)                  -> 6 Currency 10.1235
'   M13 c \ 2         -> 3 Long 5                M14 c Mod 3  -> 3 Long 1
'   M15 c ^ 2         -> 5 Double 102.48525225
'
' 改前 C3: M1..M4/M6/M7/M9/M10/M11 全是 5|Double —— 值本来就对，缺的只是**身份**:
' 装箱答 VT_R8、TypeName 答 Double。M12 已对 (AbsV 早就保子类型)，M11 的 IntV 不保。

Public Sub Main()
    Dim c As Currency
    Dim v As Variant
    Dim n As Long

    c = 10.123456
    ShowV "M1", Abs(c)
    ShowV "M2", Int(c)
    ShowV "M3", Fix(c)
    ShowV "M4", Round(c, 2)
    ShowV "M5", Sgn(c)
    ShowV "M8", -c

    c = -10.123456
    ShowV "M6", Abs(c)
    ShowV "M7", Int(c)
    ShowV "M9", Fix(c)
    ShowV "M10", Round(c, 2)

    c = 10.123456
    v = c
    ShowV "M11", Int(v)
    ShowV "M12", Abs(v)

    ShowV "M13", c \ 2
    ShowV "M14", c Mod 3
    ShowV "M15", c ^ 2

    Debug.Print "B191-DONE"
End Sub

Private Sub ShowV(ByVal tag As String, ByVal v As Variant)
    Debug.Print tag & "=" & VarType(v) & "|" & TypeName(v) & "|" & _
                  CStr(v) & "|" & CStr(CLng(v * 10000))
End Sub
