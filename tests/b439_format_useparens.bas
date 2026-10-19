Attribute VB_Name = "ModB439"
Option Explicit

' 账 #439（§B212，第一百一十二刀）= FormatNumber / FormatCurrency / FormatPercent 三家，
'   由「新开那条普查线」的可选实参那一轨量出来的**两笔账**。
'
' ① **编不过那一族**：`FormatCurrency(x)` / `FormatPercent(x)` 发成 1 个实参进 5 参的 RTL 出口
'    = **cl C2198 实参太少**，整份文件编不过。病根 = `kOD_FormatNum` 那张「省略实参填什么」的
'    默认档表**只登记了 FormatNumber 一家**，而三家的 C 签名与可选实参**完全同形**。
'
' ② **读数那一族**：三家出口第一行都写着 `(void)incLeading; (void)useParens; (void)groupDigits;`
'    —— UseParentheses 这一位被**明确丢掉**，于是负数永远交 `-1,234.57`，而运行期在它为真时交
'    `(1,234.57)`。收成一位主人 `vb6_fmtApplyUseParens`，三家共用。
'
' ⚠ 本刀真正的坑在②的**默认值**：省略档原先也填 `-1`，而 VB6 的 `True` **恒为 -1** ⇒
'   「省略」与「显式写 True」在 C 侧**同码**，两边分不开：
'     把 -1 当默认 ⇒ 显式 True 被吃掉（实测 -1,234.57 vs 运行期 (1,234.57)）；
'     把 -1 当真   ⇒ **省略档全部变括号**，把绝大多数正常调用一起改坏。
'   VB6 里 UseParentheses 的默认是 **False**，所以发码那一档改成 **0** —— 0 = 不要括号、
'   -1/1 = 要括号，两边才分得开。位 1/2/4 的默认（NumDigits=-1 / IncludeLeadingDigits=True /
'   GroupDigits=True）都还是 -1，**没动**。
'
' 运行期口径（cscript 同机代理，尺 .build/kopt.vbs / kpar.vbs / kpar2.vbs）三家**完全一致**：
'   省略 ⇒ 负号；True ⇒ 括号且无负号；0 ⇒ 负号；**正数 + True ⇒ 两样都没有**（括号只裹负数）。
'
' ⚠ 夹具**不钉货币符号**：`FormatCurrency` 的符号那一枚**跟着区域设置走**（本仓写死 U+00A5），
'   与 §B142「区域设置那一格照打不钉」同一条纪律 ⇒ 那一族只钉「括号/负号」的**形状**。

Sub Main()
    ' --- ① 省略可选实参：三家都得编得过（改前 FormatCurrency/FormatPercent 是 C2198）---
    Debug.Print "def-1arg=" & OkLen(FormatNumber(1234.5678))
    Debug.Print "cur-1arg=" & OkLen(FormatCurrency(1234.5678))
    Debug.Print "pct-1arg=" & OkLen(FormatPercent(0.1234))
    Debug.Print "cur-2arg=" & OkLen(FormatCurrency(1234.5678, 2))

    ' --- ② UseParentheses 四档（FormatNumber，字面量可钉）---
    Debug.Print "def=" & FormatNumber(-1234.5678, 2)                     ' -1,234.57   省略 = 负号
    Debug.Print "false=" & FormatNumber(-1234.5678, 2, True, False)      ' -1,234.57   显式 False = 负号
    Debug.Print "true=" & FormatNumber(-1234.5678, 2, True, True)        ' (1,234.57)  显式 True  = 括号
    Debug.Print "pos=" & FormatNumber(1234.5678, 2, True, True)         ' 1,234.57    正数不裹括号
    ' 省略中间那一位（UseParentheses 落在第 4 位）
    Debug.Print "skip=" & FormatNumber(-1234.5678, 2, , True)            ' (1,234.57)

    ' --- 三家口径一致：只钉形状，不钉货币符号 ---
    Debug.Print "cur-def-parens=" & Shape(FormatCurrency(-1234.5678, 2))         ' False
    Debug.Print "cur-true-parens=" & Shape(FormatCurrency(-1234.5678, 2, True, True))  ' True
    Debug.Print "pct-true-parens=" & Shape(FormatPercent(-1234.5678, 2, True, True))   ' True

    Debug.Print "B439-DONE"
End Sub

' 该调用必须**编得过**（①那一族的判据）；长度 > 0 即说明编过且有内容。
Private Function OkLen(s As String) As Long
    OkLen = IIf(Len(s) > 0, 1, 0)
End Function

' 只答「有没有括号」，不答具体文本 —— 货币符号/千分位那些是区域设置与实现细节。
Private Function Shape(s As String) As String
    Shape = IIf(InStr(s, "(") > 0, "True", "False")
End Function