Option Explicit

' 账 #305 §11.168 (第一百八十五刀): 「数值 + 单位文本」——谁先进格式串, 谁赢。
'
' 语料现状: #,##0 Twip 3 处 / Settings 4 / T0 2 / #### Byte 1 / #,# bytes 4 / 0 KB 一族,
' 全是**裸写**的单位文本 (没有一处用引号包起来)。改前 C3 把这些字母交给日期臂,
' 交出来是乱码: Format(1234,"#### Byte")="#### B03te"、Format(1024,"0 KB")="0 5B"。
'
' 真 VB6 读数 (.temp/k184p/P200/P201/P202/P203, 每格逐字抄自 p2xx.txt):
'   U1  #### Byte        -> 1234 Byt       'e' 在段末是指数记号, 被吞
'   U2  #,# bytes        -> 1,234 bytes    'e' 后面还有字 => 字面量
'   U3  #,##0 Twip       -> 1,234 Twip
'   U4  0 KB             -> 1024 KB
'   U5  #,##0 Settings   -> 1,234 Settings
'   U6  0 AM             -> 1234 AM        0 先出现 => 上午/下午记号也只是字面量
'   U7  0 A/P            -> 1234 A/P
'   U8  0.00 d           -> 1234.00 d
'   U9  #0 mm            -> 1234 mm        'mm' 在 0 之后还是字面量
'   U10 0 Bytes          -> 1234 Bytes
'   U11 0 e              -> "1234 "        空格留, e 吞
'   U12 0 "e"            -> 1234 e         引号内是字面量
'   U13 0\e              -> 1234e          反斜杠转义
'   U14 0 ttttt          -> 1234 ttttt
'   U15 0 ww             -> 1234 ww
'   U16 dd 0             -> 18 0           字母先出现 => 日期档
'   U17 dd0              -> 180
'   U18 yyyy 0           -> 1903 0
'   U19 00:00            -> 12:34          冒号是强信号 (0 先出现也走时间档)
'   U21 mm               -> 05             通篇没有数值占位符
'
' U16..U21 是反向证人: 本刀不许把日期档的既有读法一起改掉。

Public Sub Main()
    Show "U1", Format(1234, "#### Byte")
    Show "U2", Format(1234, "#,# bytes")
    Show "U3", Format(1234, "#,##0 Twip")
    Show "U4", Format(1024, "0 KB")
    Show "U5", Format(1234, "#,##0 Settings")
    Show "U6", Format(1234, "0 AM")
    Show "U7", Format(1234, "0 A/P")
    Show "U8", Format(1234, "0.00 d")
    Show "U9", Format(1234, "#0 mm")
    Show "U10", Format(1234, "0 Bytes")
    Show "U11", Format(1234, "0 e")
    Show "U12", Format(1234, "0 ""e""")
    Show "U13", Format(1234, "0\e")
    Show "U14", Format(1234, "0 ttttt")
    Show "U15", Format(1234, "0 ww")
    Show "U16", Format(1234, "dd 0")
    Show "U17", Format(1234, "dd0")
    Show "U18", Format(1234, "yyyy 0")
    Show "U19", Format(1234, "00:00")
    Show "U20", Format(1234, "mm")

    Debug.Print "B193-DONE"
End Sub

Private Sub Show(ByVal tag As String, ByVal s As String)
    Debug.Print tag & "=[" & s & "]"
End Sub
