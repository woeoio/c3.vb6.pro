' 账 #172: VB6 日期字面量 #...# 的**编译期折算**。
' 改前：parser 只把原文挂进 AST，发码侧读的是联合体里从没被写过的那 8 字节
'       ⇒ Debug 恰好发 0.0、Release / 带 /RTCu 的构建读到垃圾（VBFlexGrid 的 MinDate 发出
'       -6.2774359784998866e+66）。两头都钉：值等于 DateSerial 的运行期算法，且带 PM 的那条
'       只比整日期多出不到一天。
Option Explicit

Public Sub Main()
    Debug.Print "A-eq2=" & CStr(CDbl(#1/1/1900#) = 2#)
    Debug.Print "B-epoch=" & CStr(#1/1/1970# = DateSerial(1970, 1, 1))
    Debug.Print "C-maxday=" & CStr(#12/31/9999# = DateSerial(9999, 12, 31))
    Debug.Print "D-plus31=" & CStr(#1/1/1900# + 31 = #2/1/1900#)
    Debug.Print "E-pm-after=" & CStr(CDbl(#12/31/9999 11:59:59 PM#) > CDbl(#12/31/9999#))
    ' 11:59:59 PM = 当天第 86399 秒 ⇒ 比整日期多 0.9999884259…：两头都要钉
    ' (>0.99 才算真折进了时间，<1 才算没滚到第二天)。
    Debug.Print "F-pm-in-day=" & CStr(CDbl(#12/31/9999 11:59:59 PM#) - CDbl(#12/31/9999#) > 0.99)
    Debug.Print "M-pm-not-next=" & CStr(CDbl(#12/31/9999 11:59:59 PM#) - CDbl(#12/31/9999#) < 1#)
    Debug.Print "G-leap=" & CStr(Year(#2/29/2000#) & "-" & Month(#2/29/2000#) & "-" & Day(#2/29/2000#) = "2000-2-29")
    Debug.Print "H-y99=" & CStr(#1/1/99# = DateSerial(1999, 1, 1))
    Debug.Print "I-y49=" & CStr(#1/1/49# = DateSerial(2049, 1, 1))
    Debug.Print "J-dash-dmy=" & CStr(#1-2-1999# = DateSerial(1999, 2, 1))

    Dim d As Date
    d = #1/2/2020#
    Select Case d
        Case #1/1/1900# To #12/31/9999 11:59:59 PM#
            Debug.Print "K-case-hit=True"
        Case Else
            Debug.Print "K-case-hit=False"
    End Select

    Debug.Print "L-serial=" & CStr(CDbl(#1/2/2020#))
    Debug.Print "DL-DONE"
End Sub
