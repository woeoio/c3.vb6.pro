Option Explicit

' §11.222 (第二百四十刀): Err.Description 的默认表 —— 从 11 枚补到 87 枚，
' 外加三条边缘：未登记的号给兜底句、负数号给 "Automation error"、号 0 / >65535 时 raise 不成立改报 5。
' 判据全部来自真机一键扫（.temp/k230p）：
'   P335: For n = 1 To 4000: Err.Clear: Err.Raise n: ? Err.Description
'         —— 4000 个号里只有 87 个不是兜底句；
'   P336/P337: 9999 / -1 / 0 / 65536 与显式描述/显式空串；
'   P338: 非法号时 HelpFile/HelpContext 两格**原样留着**（新给的 hf2/99 不生效）。
' 本夹具只钉能拿真机读数对齐的格；描述文本是 VB6 运行期资源的 en-US 文本（CI 同为 en-US）。
Public Sub Main()
    On Error Resume Next

    ' 一号一族各取一枚
    Err.Clear
    Err.Raise 3
    Debug.Print "D1=" & CStr(Err.Number) & "/" & Err.Description

    Err.Clear
    Err.Raise 28
    Debug.Print "D2=" & CStr(Err.Number) & "/" & Err.Description

    Err.Clear
    Err.Raise 53
    Debug.Print "D3=" & CStr(Err.Number) & "/" & Err.Description

    Err.Clear
    Err.Raise 91
    Debug.Print "D4=" & CStr(Err.Number) & "/" & Err.Description

    Err.Clear
    Err.Raise 424
    Debug.Print "D5=" & CStr(Err.Number) & "/" & Err.Description

    Err.Clear
    Err.Raise 481
    Debug.Print "D6=" & CStr(Err.Number) & "/" & Err.Description

    Err.Clear
    Err.Raise 746
    Debug.Print "D7=" & CStr(Err.Number) & "/" & Err.Description

    ' 未登记号：兜底句（不是空串）
    Err.Clear
    Err.Raise 9999
    Debug.Print "D8=" & CStr(Err.Number) & "/" & Err.Description

    ' 负数号：Automation error；给了描述就原样留
    Err.Clear
    Err.Raise -1
    Debug.Print "D9=" & CStr(Err.Number) & "/" & Err.Description

    Err.Clear
    Err.Raise -1, "s3", "d3"
    Debug.Print "D10=" & CStr(Err.Number) & "/" & Err.Description & "/" & Err.Source

    ' 显式空串是「有值」：不被兜底句替掉
    Err.Clear
    Err.Raise 9999, "s6", ""
    Debug.Print "D11=" & CStr(Err.Number) & "/[" & Err.Description & "]"

    ' 号 0：raise 不成立 ⇒ 改报 5，显式 d0 丢掉，先前那两格 help 值留着
    Err.Clear
    Err.HelpFile = "hf0"
    Err.HelpContext = 77
    Err.Raise 0, "s0", "d0"
    Debug.Print "D12=" & CStr(Err.Number) & "/" & Err.Description & "/hf=" & Err.HelpFile & "/hc=" & CStr(Err.HelpContext) & "/src=" & Err.Source

    ' 号 65536 同理；新给的 hf2/99 不生效，读回仍是 hf1/88
    Err.Clear
    Err.HelpFile = "hf1"
    Err.HelpContext = 88
    Err.Raise 65536, "s2", "d2", "hf2", 99
    Debug.Print "D13=" & CStr(Err.Number) & "/" & Err.Description & "/hf=" & Err.HelpFile & "/hc=" & CStr(Err.HelpContext) & "/src=" & Err.Source
End Sub