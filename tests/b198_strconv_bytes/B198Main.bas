Option Explicit

' 账 #305 §11.173 (第一百九十刀): StrConv 的档位语义。
'
' 改前 (病): `vbUnicode` / `vbFromUnicode` 在 RTL 里是 no-op; `999` 被 `& 0x3C`
' 拆成 vbWide + vbHiragana 走了 LCMapStringEx, 交出全角串; `vbUnicode + vbUpperCase`
' 既不报错也没有大小写转换 (三个错各自一档)。
'
' 判据以 TwinBASIC 为准 —— C3 的 String 是 Unicode, 这两个字节重解释档跟 VB 的 ANSI
' 语义本来就有真实差异; VB6Mini 在这两档上自相矛盾 (见 .temp/k189p/P213.bas)。
' TB 读数: .temp/k189p/TB_StrConv.bas (zh-CN / cp936 那台)。
'
' 钉 18 格 + DONE。没钉的几格与理由 (都记在 §11.173 在册):
'   S6   双重往返: TB 的字节串把奇数尾字节藏在隐式长度里, C3 的 String 没有那个字段;
'   S7   逐字符码点跟代码页走 (zh-CN 是 63,0 / en-US 是 196,0), 只钉 Len;
'   X1/X3 Err 5 已钉, 值那半边要等「语句中止」闸 (raise 之后调用方仍拿到空串);
'   X4   Null 形参: 94 已对, 但会多打一行空值, 同一个闸。

Public Sub Main()
    On Error Resume Next

    Dim junk As String

    ' ---- 字节重解释两档 ----
    Dump "S1", StrConv("ABC", vbUnicode)
    Dump "S2", StrConv("ABC", vbFromUnicode)
    P "S3", Len(StrConv("ABC", vbUnicode))
    P "S4", Len(StrConv("", vbUnicode))
    P "S5", Len(StrConv("", vbFromUnicode))

    ' 非 ASCII: vbUnicode 把 C4 00 按代码页解码 (zh-CN: 非法前导 -> '?' 再跟 NUL),
    ' 所以只钉长度 (两台代码页不同, 长度都是 2)。
    P "S7L", Len(StrConv("Ä", vbUnicode))
    Dump "S8", StrConv("Ä", vbFromUnicode)
    P "S9", Len(StrConv("Ä", vbUnicode))
    P "S10", Len(StrConv("Ä", vbFromUnicode))

    ' ---- 大小写 / 首字母 ----
    Dump "C1", StrConv("abc", vbUpperCase)
    Dump "C2", StrConv("ABC", vbLowerCase)
    Dump "C3", StrConv("hello world", vbProperCase)
    Dump "C4", StrConv("abc", 0)
    Dump "C5", StrConv("aBc dEf", vbProperCase)

    ' ---- 非法档: 只钉 Err ----
    junk = StrConv("abc", vbUnicode + vbUpperCase)
    P "X1E", Err.Number
    Err.Clear
    P "X2", Err.Number
    junk = StrConv("abc", 999)
    P "X3E", Err.Number
    Err.Clear

    Dump "X5", StrConv(Empty, vbUpperCase)

    ' ---- LCID 形参 ----
    Dump "L1", StrConv("abc", vbUpperCase, 1033)

    Debug.Print "B198-DONE"
End Sub

Private Sub Dump(ByVal tag As String, ByVal s As String)
    Dim i As Long, codes As String
    codes = ""
    For i = 1 To Len(s)
        codes = codes & CStr(AscW(Mid(s, i, 1))) & ","
    Next i
    Debug.Print tag & "=[len " & Len(s) & " :" & codes & "]"
End Sub

Private Sub P(ByVal tag As String, ByVal n As Long)
    Debug.Print tag & "=[" & n & "]"
End Sub
