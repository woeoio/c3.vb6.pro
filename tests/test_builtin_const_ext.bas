Option Explicit

' 内建常量表补项：PhotoDemon 剩余 VB3001 里的 8 条来自两枚缺表常量
'   vbCFFiles (= 15, CF_HDROP)                      -> pdClipboardMain.cls 等 7 处
'   vbAsyncStatusCodeEndDownloadData (= 6)          -> pdDownload.ctl 1 处
' 表里没有它们时，Option Explicit 下每处引用各报一条 VB3001 并落成隐式 Variant。

Public Sub Main()
    Dim f As Long
    Dim s As Long

    f = vbCFFiles
    s = vbAsyncStatusCodeEndDownloadData

    Debug.Print "BCE-CF=" & f
    Debug.Print "BCE-ASYNC=" & s
    Debug.Print "BCE PASS"
End Sub
