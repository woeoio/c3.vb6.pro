Option Explicit

' 回归: VB6 的浮点→整数取整是**银行家舍入** (四舍六入五取偶), 不是半值远离零。
' 依据不只是手册: 真 VB6 的 Corporate Hierarchy 参考图里
' `NSpacing = 34 * 1.25 = 42.5` 落在 42、`lMBottom = 10 * 1.25 = 12.5` 落在 12
' (树线列 x=80/122/164 vs C3 的 82/125/168), 与 round() 的 43/13 不符。
' 改前读数: RB-1=3 RB-3=3 RB-5=-3 (半值远离零); 真 VB6 = 2 / 2 / -2。

Sub Main()
    Dim l As Long

    l = 2.5
    Debug.Print "RB-1=" & l
    l = 3.5
    Debug.Print "RB-2=" & l
    Debug.Print "RB-3=" & CLng(2.5)
    Debug.Print "RB-4=" & CLng(3.5)
    Debug.Print "RB-5=" & CLng(-2.5)
    Debug.Print "RB-6=" & CLng(-3.5)
    Debug.Print "RB-7=" & CInt(2.5)
    Debug.Print "RB-8=" & CInt(3.5)
    Debug.Print "RB-DONE"
End Sub
