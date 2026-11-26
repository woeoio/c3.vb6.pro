Option Explicit

' 账 #305 §11.172 (第一百八十九刀): IsNumeric / IsDate 的判定面。
'
' 病: `vb6_IsNumeric` 的 switch 少 VT_EMPTY 一档 —— 落到 default 答 False，
' 而真机 `IsNumeric(Empty)` 是 True（Empty 的缺省值就是 0，数值语境里与 0 同用；
' 与 Null 不同：`IsNumeric(Null)` 真机 False）。其余 63 格改前就已经逐字相同。
'
' 真 VB6 读数:
'   P214 (.temp/k189p/p214.txt, 42 格): N1..N30 文本档 (N14 不入夹具, 见下) + D1..D12 日期档
'   P215 (.temp/k189p/p215.txt, 24 格): E1..E16 字面量档 + V1..V8 整只变体槽档
' 本夹具把 65 格合成一份，两台（x64/x86）都钉。
' ⚠ N14 IsNumeric("$1") 不入夹具：货币符号跟机器 locale 走（本机 zh-CN ¥ 答 False，
'   CI en-US $ 答 True），真值不唯一（§11.172 在册第二条）。
' ⚠ E13/E15 用 CDate("12:00:00") / DateSerial(1900,1,1) 而不是 #12:00:00 PM#：
'   C3 的日期字面量解析不收纯时间形（§11.172 在册第一条），与这两个判据无关。

Public Sub Main()
    On Error Resume Next

    Dim v As Variant

    ' ---- 文本档: 数值 ----
    Show "N1", IsNumeric("1"): ShowE "N1"
    Show "N2", IsNumeric(" 1 "): ShowE "N2"
    Show "N3", IsNumeric("1.5"): ShowE "N3"
    Show "N4", IsNumeric(".5"): ShowE "N4"
    Show "N5", IsNumeric("5."): ShowE "N5"
    Show "N6", IsNumeric("+1"): ShowE "N6"
    Show "N7", IsNumeric("-1"): ShowE "N7"
    Show "N8", IsNumeric("1e3"): ShowE "N8"
    Show "N9", IsNumeric("1E+3"): ShowE "N9"
    Show "N10", IsNumeric("1d3"): ShowE "N10"
    Show "N11", IsNumeric("&H10"): ShowE "N11"
    Show "N12", IsNumeric("&O17"): ShowE "N12"
    Show "N13", IsNumeric("1,000"): ShowE "N13"
    Show "N15", IsNumeric("1%"): ShowE "N15"
    Show "N16", IsNumeric("1.2.3"): ShowE "N16"
    Show "N17", IsNumeric(""): ShowE "N17"
    Show "N18", IsNumeric(" "): ShowE "N18"
    Show "N19", IsNumeric("abc"): ShowE "N19"
    Show "N20", IsNumeric("1 "): ShowE "N20"
    Show "N21", IsNumeric(" 1"): ShowE "N21"
    Show "N22", IsNumeric("1 2"): ShowE "N22"
    Show "N23", IsNumeric(Empty): ShowE "N23"
    Show "N24", IsNumeric(123): ShowE "N24"
    Show "N25", IsNumeric(True): ShowE "N25"
    Show "N26", IsNumeric("1,5"): ShowE "N26"
    Show "N27", IsNumeric("(1)"): ShowE "N27"
    Show "N28", IsNumeric("1e"): ShowE "N28"
    Show "N29", IsNumeric("--1"): ShowE "N29"
    Show "N30", IsNumeric("0x10"): ShowE "N30"

    ' ---- 文本档: 日期 ----
    Show "D1", IsDate("2026-11-26"): ShowE "D1"
    Show "D2", IsDate("2026/11/26"): ShowE "D2"
    Show "D3", IsDate("11/26/2026"): ShowE "D3"
    Show "D4", IsDate("2026-13-01"): ShowE "D4"
    Show "D5", IsDate("2026-02-30"): ShowE "D5"
    Show "D6", IsDate(""): ShowE "D6"
    Show "D7", IsDate("13:00"): ShowE "D7"
    Show "D8", IsDate("2026-11-26 13:00"): ShowE "D8"
    Show "D9", IsDate(Empty): ShowE "D9"
    Show "D10", IsDate(0): ShowE "D10"
    Show "D11", IsDate("abc"): ShowE "D11"
    Show "D12", IsDate(#11/26/2026#): ShowE "D12"

    ' ---- 整只变体槽: 字面量与子类型 ----
    Show "E1", IsNumeric(Empty): ShowE "E1"
    Show "E2", IsDate(Empty): ShowE "E2"
    Show "E3", IsNumeric(Null): ShowE "E3"
    Show "E4", IsDate(Null): ShowE "E4"
    Show "E5", IsNumeric(True): ShowE "E5"
    Show "E6", IsNumeric(False): ShowE "E6"
    Show "E7", IsDate(#11/26/2026#): ShowE "E7"
    Show "E8", IsNumeric(#11/26/2026#): ShowE "E8"
    Show "E9", IsNumeric(CCur(1.5)): ShowE "E9"
    Show "E10", IsNumeric(CDec(1.5)): ShowE "E10"
    Show "E11", IsNumeric(CByte(7)): ShowE "E11"
    Show "E12", IsNumeric(CSng(1.5)): ShowE "E12"
    Show "E13", IsDate(CDate("12:00:00")): ShowE "E13"
    Show "E14", IsDate(1): ShowE "E14"
    Show "E15", IsDate(DateSerial(1900, 1, 1)): ShowE "E15"
    Show "E16", IsNumeric(""): ShowE "E16"

    ' ---- 整只变体槽: 变量档 ----
    v = Empty
    Show "V1", IsNumeric(v): ShowE "V1"
    Show "V2", IsDate(v): ShowE "V2"
    v = Null
    Show "V3", IsNumeric(v): ShowE "V3"
    Show "V4", IsDate(v): ShowE "V4"
    v = "12"
    Show "V5", IsNumeric(v): ShowE "V5"
    v = "abc"
    Show "V6", IsNumeric(v): ShowE "V6"
    v = #11/26/2026#
    Show "V7", IsDate(v): ShowE "V7"
    Show "V8", IsNumeric(v): ShowE "V8"

    Debug.Print "B197-DONE"
End Sub

Private Sub Show(ByVal tag As String, ByVal b As Boolean)
    Debug.Print tag & "=[" & CStr(b) & "]"
End Sub

Private Sub ShowE(ByVal tag As String)
    If Err.Number <> 0 Then
        Debug.Print tag & "=[ERR " & Err.Number & "]"
        Err.Clear
    End If
End Sub
