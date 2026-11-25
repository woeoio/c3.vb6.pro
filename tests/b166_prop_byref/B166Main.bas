Option Explicit

' 账 #305 §11.147 (第一百六十五刀) / §11.148 (第一百六十六刀): 属性读当 ByRef 实参 ——
' 读-改-写。两族问同一句:
'   P* = 晚绑定目标 (`o As Object`)，实参按 VT_BYREF 发；
'   Q* = 早绑定目标 (`b As B166Box`，直接调 vb6_B166Box_Inc)。
' P8 / Q3 是**反面证人**: String 属性暂不开 (临时量要管 BSTR 所有权)，仍按值。
Public Sub Main()
    Dim b As B166Box
    Dim s As String
    Dim o As Object

    Set b = New B166Box
    Set o = b

    b.Tag = 40
    o.Inc b.Tag
    Debug.Print "B166-P1=" & b.Tag
    o.Inc b.Tag
    Debug.Print "B166-P2=" & b.Tag

    b.INum = 20
    o.IncI b.INum
    Debug.Print "B166-P3=" & b.INum

    b.DNum = 30
    o.IncD b.DNum
    Debug.Print "B166-P4=" & b.DNum

    b.YNum = 3
    o.SetY b.YNum
    Debug.Print "B166-P5=" & b.YNum

    b.FNum = False
    o.FlipB b.FNum
    Debug.Print "B166-P6=" & b.FNum

    b.Tm = DateSerial(2020, 1, 1)
    o.AddDay b.Tm, 2
    ' 用数值取值器打印 (Year/Month/Day)，不用 Format 的日期格式串 —— 后者跟系统区域设置，
    ' 换一台机器就换一种写法（CI 的 locale 门就是这么红的）。
    Debug.Print "B166-P7=" & Year(b.Tm) & "," & Month(b.Tm) & "," & Day(b.Tm)

    b.Label2 = "ab"
    o.Append b.Label2, "cd"
    Debug.Print "B166-P8=" & b.Label2

    s = "xy"
    o.Append s, "z"
    Debug.Print "B166-P9=" & s

    ' ---- 早绑定接收者那一族 (§11.148) ----
    b.Tag = 40
    b.Inc b.Tag
    Debug.Print "B166-Q1=" & b.Tag

    b.DNum = 30
    b.IncD b.DNum
    Debug.Print "B166-Q2=" & b.DNum

    b.Label2 = "ab"
    b.Append b.Label2, "cd"
    Debug.Print "B166-Q3=" & b.Label2

    Debug.Print "B166-DONE"
End Sub
