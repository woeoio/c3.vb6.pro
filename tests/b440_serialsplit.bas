Attribute VB_Name = "ModB440"
Option Explicit

' 账 #440 (§B209，第一百零七刀) = 账 #420 (§B194) + 账 #421 (§B195) 一次结掉:
' 「一枚序列号拆成日期 + 日内秒」这一问原本**两处各答一半**。
'
' ① §B194 (#420) 一天末尾那一格只把日内秒归零、**日期那一半不进位**:
'    46023.9999988426 (= 2026/1/1 23:59:59.9) → H/M/S=0/0/0 且 **Day=2**; 改前答 Day=1。
' ② §B195 (#421) epoch 之前的零头与运行期**反号**: 运行期取**绝对值**, 本仓原先 `+1.0`:
'    -657433.395833333 → 运行期 9:30, 改前本仓 14:30 (同一瞬间的两种拆法)。
'    且两半**符号对称**: -46023.9999988426 → 运行期 1773/12/28 (日期也进了一格)。
'
' 两格都由 cscript 同机代理逐格量死 (尺 .build/k421_grid.vbs / k420_oracle.vbs /
' k420b_grid.vbs): 主网格 12 格 + 进位边界网格 8 格 = **20 格零失配**。
' ⚠ 日期那一半用 `(int32_t)date` **向零截断**(不是 floor) —— 运行期就是这一套, 改前就一直对;
'   反号的只有**零头**那一半。两半现在共用一位主人 `vb6_split_serial`, 不许再各抄一遍。

Sub Main()
    ' --- ① 一天末尾进位 (#420) ---
    Debug.Print "roll-hms=" & Hms(46023.9999988426)        ' 0/0/0
    Debug.Print "roll-day=" & Day(46023.9999988426)         ' 2
    Debug.Print "neghi-day=" & Day(-46023.9999988426)       ' 28 (负号那一侧也进)

    ' --- ② 零头取绝对值 (#421) ---
    Debug.Print "neg-hms=" & Hms(-657433.395833333)         ' 9/30/0
    Debug.Print "neg-day=" & Day(-657433.395833333) & "/" & Month(-657433.395833333) & "/" & Year(-657433.395833333)
    Debug.Print "neg4-hms=" & Hms(-1.25)                    ' 6/0/0
    Debug.Print "neg6-hms=" & Hms(-0.75)                    ' 18/0/0

    ' --- 对照: 正数与整数值两半**不许动** (§B193 那条 +0.5 进位口径) ---
    Debug.Print "pos-hms=" & Hms(46027.6046875)            ' 14/30/45
    Debug.Print "whole-hms=" & Hms(-657433#)               ' 0/0/0
    Debug.Print "frac6-hms=" & Hms(0.99999305555)           ' 23/59/59 (边界另一侧)

    Debug.Print "B440-DONE"
End Sub

Private Function Hms(v As Double) As String
    Hms = Hour(v) & "/" & Minute(v) & "/" & Second(v)
End Function