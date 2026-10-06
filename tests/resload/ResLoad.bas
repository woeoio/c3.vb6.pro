Option Explicit

' Fix <vbeclipse> 2026-10-06: LoadRes* 实装回归 — .res 资源段加载。
' res.res 由 res.rc 编译 (rc.exe /fo test.res res.rc), 内容:
'   字符串表 101 / BITMAP 102 (GDI+ 32bpp) / ICON 103 / BIN1 CUSTOM (8 字节)
'   IMG1 PNG / IMG2 WEBP (PNG 字节挂自定义类型, 验签名分流)。

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Sub Main()
    On Error Resume Next
    Dim s As String
    s = LoadResString(101)
    Debug.Print "STR=" & s & "|" & TF(s = "ResString-OK")
    Err.Clear

    Dim p As Object
    Set p = LoadResPicture(102, 0)          ' vbResBitmap: RT_BITMAP 裸 DIB
    Debug.Print "BMP=" & TF(Not (p Is Nothing)) & "/" & Err.Number
    Err.Clear

    Set p = LoadResPicture(103, 1)          ' vbResIcon: RT_GROUP_ICON 组解析
    Debug.Print "ICO=" & TF(Not (p Is Nothing)) & "/" & Err.Number
    Err.Clear

    Set p = LoadResPicture("IMG1", "PNG")   ' 字符串类型名 → 字节签名 → WIC 解码
    Debug.Print "PNG=" & TF(Not (p Is Nothing)) & "/" & Err.Number
    Err.Clear

    Set p = LoadResPicture("IMG2", "WEBP")  ' WebP 字节走 WIC (PNG 字节占位: 同签名路径)
    Debug.Print "WEBP=" & TF(Not (p Is Nothing)) & "/" & Err.Number
    Err.Clear

    Dim d() As Byte
    d = LoadResData("BIN1", "CUSTOM")
    Debug.Print "BIN=" & (UBound(d) + 1) & "/" & d(0)
    Err.Clear

    Dim miss As String
    miss = LoadResString(999)               ' 找不到 → 错误 326
    Debug.Print "MISS=" & Err.Number
    Err.Clear
End Sub
