Option Explicit

' 账 #305 §11.165 (第一百八十二刀): Currency 的值域闸门与溢出 Err 6。
'
' 真 VB6 探针 .temp/k182p (P188/P189.bas) 读数:
'   O2  c = CCur("922337203685477.5808")   -> Err 6
'   O3  c = CCur("1000") : c = c * <上限>   -> Err 6
'   E3  Err.Clear 后                          -> 0
' 值域内 (double 能精确承载的) 证人: N1/N2/N3/N5 全部 Err 0 且 CStr 逐字相同。
'
' 改前 C3: O1/O2/O3 全是 Err 0，值静默回绕成 -922337203685477.5808
' ((int64_t) 越界在 C 里是未定义行为)。
'
' 在册: 上限 922337203685477.5807 与 .5808 在 double 里是**同一个数**，
' 因此 O1 (正好上限、真机 Err 0) 无法与 O2 区分 —— C3 两格都报 Err 6；
' 同理 999999999999.9999 (16 位有效数字) 在 double 里落成 1000000000000。
' 要逐一还原需要 Currency 改 int64 承载 (另一刀)。

Public Sub Main()
    Dim c As Currency
    On Error Resume Next

    Err.Clear
    c = CCur("1234567.8912"):       Debug.Print "N1=" & Err.Number & "|" & CStr(c)
    c = CCur("-1234567.8912"):      Debug.Print "N2=" & Err.Number & "|" & CStr(c)
    c = CCur("1000000000000"):      Debug.Print "N3=" & Err.Number & "|" & CStr(c)
    c = 123456789.1234:             Debug.Print "N5=" & Err.Number & "|" & CStr(c)
    c = 10.123456:                  Debug.Print "N6=" & Err.Number & "|" & CStr(c)

    Err.Clear
    c = 0
    c = CCur("922337203685477.5808")
    Debug.Print "E1=" & Err.Number
    Err.Clear
    c = CCur("1000")
    c = c * 922337203685477.5807
    Debug.Print "E2=" & Err.Number
    Err.Clear
    Debug.Print "E3=" & Err.Number
    On Error GoTo 0

    Debug.Print "B185-DONE"
End Sub
