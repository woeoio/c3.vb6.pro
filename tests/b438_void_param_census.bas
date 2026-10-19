Attribute VB_Name = "ModB438"
Option Explicit
' 账 #438（§B214）= 「(void) 形参」那一族的普查结论。三条都是**否证**，但都得钉住，
'   否则下一次普查会重新把它们举成新账，然后有人去「修」一个没坏的东西。
Sub Main()
    ' ① vb6_DateDiff 的 (void)firstWeekOfYear 是**实测惰性**、不是丢掉了一位。
    '    尺 .build/kfwy.vbs：6 个间隔(w/ww/y/q/m/d) x 5 组跨年边界，fwy=1/2/3 **读数全同**。
    '    w / ww 两格读数与 fwy **无关**，故 (void) 是对的 —— §B197 的两条公式只吃 f=firstDayOfWeek。
    Debug.Print "w_f1=" & DateDiff("w", DateSerial(2000, 1, 1), DateSerial(2000, 12, 31), vbSunday, 1)
    Debug.Print "w_f2=" & DateDiff("w", DateSerial(2000, 1, 1), DateSerial(2000, 12, 31), vbSunday, 2)
    Debug.Print "w_f3=" & DateDiff("w", DateSerial(2000, 1, 1), DateSerial(2000, 12, 31), vbSunday, 3)
    Debug.Print "ww_f1=" & DateDiff("ww", DateSerial(2000, 1, 1), DateSerial(2000, 12, 31), vbSunday, 1)
    Debug.Print "ww_f2=" & DateDiff("ww", DateSerial(2000, 1, 1), DateSerial(2000, 12, 31), vbSunday, 2)
    Debug.Print "ww_f3=" & DateDiff("ww", DateSerial(2000, 1, 1), DateSerial(2000, 12, 31), vbSunday, 3)
    ' 跨年那组也照钉（边界最容易出偏差）
    Debug.Print "wwb_1=" & DateDiff("ww", DateSerial(2020, 12, 28), DateSerial(2021, 1, 4), vbSunday, 1)
    Debug.Print "wwb_2=" & DateDiff("ww", DateSerial(2020, 12, 28), DateSerial(2021, 1, 4), vbSunday, 2)

    ' ② vb6_IIfLong / vb6_IIfDouble 的 (void)falsepart 是**多余**的，下一行就用到了。
    '    这三个出口**发码真的走得到**（cgen_expr_call_builtin_pre.inc 会发 vb6_IIfLong/Double/BSTR），
    '    而 IIf 是 VB6 高频件 —— 普查前它**一处回归都没有**，这里补上。
    Debug.Print "iifL_t=" & IIf(1 > 0, 11, 22)          ' 11
    Debug.Print "iifL_f=" & IIf(1 > 2, 11, 22)          ' 22
    Debug.Print "iifD_t=" & IIf(1 > 0, 1.5, 2.5)        ' 1.5
    Debug.Print "iifD_f=" & IIf(1 > 2, 1.5, 2.5)        ' 2.5
    Debug.Print "iifS_t=" & IIf(1 > 0, "yes", "no")     ' yes
    Debug.Print "iifS_f=" & IIf(1 > 2, "yes", "no")     ' no
    ' 零与负数两格：VB6 的条件是「**非零即真**」，所以 -5 是真 ⇒ 交 7。
    ' ⚠ 这里我第一版把注释写成「-5 该交 8」，量出来是 7 —— **注释错了、代码是对的**。
    '   真正的风险方向是反过来：把负数当假。所以这格钉的是「非零（含负）即真」。
    Debug.Print "iifL_0=" & IIf(0, 7, 8)                ' 8   0 是假
    Debug.Print "iifL_n=" & IIf(-5, 7, 8)               ' 7   -5 非零 ⇒ 真

    Debug.Print "B438-DONE"
End Sub