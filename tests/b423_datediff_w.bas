Attribute VB_Name = "ModB423"
Option Explicit

' 账 #423 (§B197，第一百零八刀) = DateDiff 的「ww」与「w」两条臂。
'
' 病根 = `vb6_parse_interval` 是 DateAdd / DateDiff / DatePart **三家共用**的一问，而
' `w` 与 `ww` 的含义**按 builtin 分岔**:
'     DatePart("w",  …) = 星期序号(与周首无关)    DateDiff("w",  …) = **整段七天**
'     DatePart("ww", …) = 一年里第几周            DateDiff("ww", …) = 半开区间 (a,b] 里
'                                                   落在第 f 天那一格的条数(方向带符号)
' 共用一枚枚举 ⇒ 改 parse_interval 会把另外两家的**正确答案**一起弄坏
' (实测 DatePart 的 w/ww 今天都是 2、都对)。所以**枚举那一问不动**，DateDiff 自己认这一格。
'
' 改前本仓这一格发的是 `(d2-d1)/7` 而 d1/d2 是**截断过的整日**、且落进 `vb6_di_weekday`
' 的 `d2 - d1` ⇒ 答的是**天数** (1/5→1/12 答 7、1/5→2/5 答 31)；运行期两格都答 1 / 4。
'
' 运行期口径 (cscript 同机代理, 尺 .build/k423_oracle.vbs, 逐格量死):
'   ① "ww" = floor((b-f)/7) - floor((a-f)/7)，f 参与 ⇒ 改周首**会动**这一格
'   ② "w"  = trunc((b-a)/7)，**与 f 无关** ⇒ 改周首**不许动**这一格
'
' 承面 = **0 处** (真语料 DateDiff 76 处全是三实参形、区间只有 d/s/m/n/y) —— 纸面缺口，
' 但它与 §B207/§B208 是同一格问句(「这个间隔在**这一家**里要什么」)的第三家，故一并补齐。

Sub Main()
    ' --- ② "w" = 整段七天，与周首无关 ---
    Debug.Print "w-7d=" & DateDiff("w", #1/5/2026#, #1/12/2026#)          ' 1
    Debug.Print "w-31d=" & DateDiff("w", #1/5/2026#, #2/5/2026#)          ' 4
    Debug.Print "w-rev=" & DateDiff("w", #1/12/2026#, #1/5/2026#)          ' -1
    Debug.Print "w-same=" & DateDiff("w", #1/5/2026#, #1/5/2026#)           ' 0
    ' 周首**不许动**这一格：f=2 与 f=7 同答
    Debug.Print "w-f2=" & DateDiff("w", #1/5/2026#, #2/5/2026#, 2, 1)      ' 4
    Debug.Print "w-f7=" & DateDiff("w", #1/5/2026#, #2/5/2026#, 7, 1)      ' 4

    ' --- ① "ww" = 半开区间里落在第 f 天那一格的条数，周首**参与** ---
    Debug.Print "ww-7d=" & DateDiff("ww", #1/5/2026#, #1/12/2026#, 1, 1)   ' 1
    Debug.Print "ww-rev=" & DateDiff("ww", #1/12/2026#, #1/5/2026#, 1, 1)   ' -1
    Debug.Print "ww-31d=" & DateDiff("ww", #1/5/2026#, #2/5/2026#, 1, 1)   ' 4
    Debug.Print "ww-f2=" & DateDiff("ww", #1/5/2026#, #2/5/2026#, 2, 1)    ' 4
    Debug.Print "ww-f7=" & DateDiff("ww", #1/5/2026#, #2/5/2026#, 7, 1)    ' 4

    ' --- 负控: 共用解析器那一家**不许动** (本刀只改 DateDiff) ---
    Debug.Print "NC-ok=" & (IIf(DatePart("w", #1/5/2026#, 1, 1) = 2 _
                              And DatePart("ww", #1/5/2026#, 1, 1) = 2 _
                              And DateAdd("w", 1, #1/5/2026#) = #1/6/2026#, 1, 0))  ' 1

    Debug.Print "B423-DONE"
End Sub