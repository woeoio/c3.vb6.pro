Attribute VB_Name = "ModB437"
Option Explicit

' 账 #437（§B213，第一百一十三刀）= FormatNumber / FormatCurrency / FormatPercent 三家那
'   5 位签名上**剩下那两枚**被 `(void)` 丢掉的位。§B212 收了 useParens，incLeading 与
'   groupDigits 当时**一个字没动**，所以接着量。
'
' 两枚都是**静默答错**（不编译失败），运行期口径（cscript 同机代理，尺 `.build/klead.vbs`
' / `kedge.vbs`）三家**完全一致**：
'
'   IncludeLeadingDigits：显式 False ⇒ 整数部分**恰为 "0"** 时去掉那个 0。
'     FormatNumber(0.5,2,False)=`.50`   (0,2,False)=`.00`   (-0.5,2,False)=`-.50`
'     (0,0,False)=**空串**            省略 ⇒ `0.50`（默认带前导 0）
'     FormatPercent(0,0,False)=**`%`**  FormatPercent(0.005,2,False)=`.50%`
'
'   GroupThousands：显式 False ⇒ 不插千分位；省略/True ⇒ 每 3 位插 ','。
'     FormatNumber(1234.5,2,True,False,False)=`1234.50`   `FormatPercent(12.345,2,True,False,True)=`1,234.50%`
'
' ⚠ 两个把实现写错就交不出来的边（都实测过，不是猜的）：
'   ① percent 的 buf 是 `"0%"` / `"1234.50%"`，下一位是 `'%'` **不是**小数点或结束。
'      「整数部分恰好是 0」这个判据若只认 0 与 '.' ⇒ `"0%"` 被当成「不止一位」⇒ 交 `0%`
'      而不是运行期的 `%`；若千分位按「数到 '.' 为止」来算整数段 ⇒ `%` 被当成一位数字
'      ⇒ 逗号插错位置。所以两位主人都按**数字跑**来看，不按 '.' 找。
'   ② `FormatNumber(0,0,False)` 是**空串**，不是 `0` —— 去掉 0 就是去掉 0，不留痕。

Sub Main()
    ' --- IncludeLeadingDigits（FormatNumber，字面量可钉）---
    Debug.Print "inc_T=" & FormatNumber(0.5, 2, True)                ' 0.50   默认形态不许变
    Debug.Print "inc_def=" & FormatNumber(0.5, 2)                   ' 0.50   省略 = 默认带 0
    Debug.Print "inc_F=" & FormatNumber(0.5, 2, False)               ' .50    去掉前导 0
    Debug.Print "inc_F_neg=" & FormatNumber(-0.5, 2, False)         ' -.50   负号留着
    Debug.Print "inc_F_zero=" & FormatNumber(0, 2, False)            ' .00
    Debug.Print "inc_F_empty=[" & FormatNumber(0, 0, False) & "]"    ' 空串
    Debug.Print "inc_F_odd=" & FormatNumber(1234.5, 2, False)         ' 1,234.50 整数段不止一位 ⇒ 前导 0 无关（groupDigits 省略 = 默认分组）

    ' --- GroupThousands（FormatNumber）---
    Debug.Print "grp_T=" & FormatNumber(1234.5, 2, True, False, True)    ' 1,234.50
    Debug.Print "grp_def=" & FormatNumber(1234.5, 2)                     ' 1,234.50 省略 = 分组
    Debug.Print "grp_F=" & FormatNumber(1234.5, 2, True, False, False)   ' 1234.50
    Debug.Print "grp_F_neg=" & FormatNumber(-1234.5, 2, True, False, False)  ' -1234.50
    Debug.Print "grp_big=" & FormatNumber(1000000, 2, True, False, True)  ' 1,000,000.00

    ' --- 两枚在 FormatPercent 上也生效（buf 里带 '%' ⇒ ①那一族最容易写错）---
    Debug.Print "pct_inc_T=" & FormatPercent(0.005, 2, True)                     ' 0.50%
    Debug.Print "pct_inc_F=" & FormatPercent(0.005, 2, False)                    ' .50%
    Debug.Print "pct_inc_F_pct=[" & FormatPercent(0, 0, False) & "]"              ' %
    Debug.Print "pct_grp_T=" & FormatPercent(12.345, 2, True, False, True)        ' 1,234.50%
    Debug.Print "pct_grp_F=" & FormatPercent(12.345, 2, True, False, False)       ' 1234.50%
    Debug.Print "pct_grp_def=" & FormatPercent(12.345, 2)                         ' 1,234.50%

    ' --- FormatCurrency：两枚都生效，但**只钉形状**，不钉货币符号 ---
    '     （符号跟着区域设置走，与 §B142 / §B212 同一条纪律）
    Debug.Print "cur_inc_T=" & StripCur(FormatCurrency(0.5, 2, True))          ' 0.50
    Debug.Print "cur_inc_F=" & StripCur(FormatCurrency(0.5, 2, False))         ' .50
    Debug.Print "cur_grp_T=" & StripCur(FormatCurrency(1234.5, 2, True, False, True))   ' 1,234.50
    Debug.Print "cur_grp_F=" & StripCur(FormatCurrency(1234.5, 2, True, False, False))  ' 1234.50
    Debug.Print "cur_grp_neg=" & StripCur(FormatCurrency(-1234.5, 2, True, False, False)) ' -1234.50

    ' --- 回归：§B212 那一族不许被本刀带坏 ---
    ' ⚠ §B212 的夹具只钉了 `InStr(s,"(")>0` 的**形状**，于是「符号该在括号里还是括号外」
    '   整整错了一刀而没人看见（引擎是 `(¥1,234.57)`）。这里把**位置**连同千分位一起钉。
    Debug.Print "k_def=" & FormatNumber(-1234.5678, 2)                       ' -1,234.57
    Debug.Print "k_false=" & FormatNumber(-1234.5678, 2, True, False)        ' -1,234.57
    Debug.Print "k_true=" & FormatNumber(-1234.5678, 2, True, True)          ' (1,234.57)
    Debug.Print "k_skip=" & FormatNumber(-1234.5678, 2, , True)               ' (1,234.57)
    Debug.Print "k_pct_def=" & FormatPercent(0.1234)                         ' 12.34%
    Debug.Print "k_pct_neg=" & FormatPercent(-1234.5678, 2)                 ' -123,456.78%
    Debug.Print "k_pct_true=" & FormatPercent(-1234.5678, 2, True, True)    ' (123,456.78%)
    Debug.Print "k_pct_ng=" & FormatPercent(-1234.5678, 2, True, True, False) ' (123456.78%)
    ' 符号在括号**里面**（§B212 写成外面了）。剥掉符号后连括号一起比对。
    Debug.Print "k_cur_par=" & StripCur(FormatCurrency(-1234.5678, 2, True, True))   ' (1,234.57)
    Debug.Print "k_cur_ng=" & StripCur(FormatCurrency(-1234.5678, 2, True, True, False)) ' (1234.57)

    Debug.Print "B437-DONE"
End Sub

' 把**所有**非「数字 + 标点」的字符丢掉（首当其冲就是货币符号），留下数字与
' 括号/逗号/小数点/百分号 —— 括号**位置**照样钉得住，符号本身不依赖区域设置。
Private Function StripCur(s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "," Or ch = "-" _
           Or ch = "(" Or ch = ")" Or ch = "%" Then
            out = out & ch
        End If
    Next i
    StripCur = out
End Function