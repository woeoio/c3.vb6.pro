Option Explicit

' 账 #606（第324刀）夹具: 一枚 Case 标签**本身就是整枚 Variant 槽**时，承载档必须跟着换成整枚槽。
'
' 修复前四档主语全部发**编不过的 C**（实测 error C2088: 「==」对于 struct 非法）:
'   Variant 主语 -> int32_t _vb6_select_N = vb6_VariantToLong((*v));  再  == (*w)
'   Long    主语 -> int32_t _vb6_select_N = (*n);                     再  == (*w)
'   String  主语 -> BSTR _vb6_select_N = (*t);                        再  vb6_StrCmp(_vb6_select_N, (*w))
'   Double  主语 -> float _vb6_select_N = (*d);                       再  (float) == (float)((*w))
' 三条整档承载都把主语折成一个标量，再拿标量去比那枚结构体。VB6 那一侧本来就是按 Case 项
' 逐项 VarCmp，所以修法 = 承载档由标签回答: 有槽标签就整枚槽 + 逐项 VarCmp，主语不是槽的
' 那三档按 wrapVariantValue 唯一出口装箱（借来的不收，同 §B164 那位主人）。
'
' 每档三条读数: 命中字面量标签 / 命中槽标签 / 两支都不命中（走 Case Else）。
' 末两档 hitIs / hitTo 钉的是**槽标签落在关系形与范围形**上的那两形（第324刀当时只在
' 一次性探针 .local/p606/P606e.bas 量过，本刀升格进来；改夹具=动清单那一行，随本刀一起登记）。
' 读数一律不带方括号: 套件用 PowerShell 的 -like 比针，[ABC] 会被当字符类。

Private Function hitV(v As Variant, w As Variant) As String
    Select Case v
        Case 7
            hitV = "lit"
        Case w
            hitV = "slot"
        Case Else
            hitV = "none"
    End Select
End Function

Private Function hitL(n As Long, w As Variant) As String
    Select Case n
        Case 7
            hitL = "lit"
        Case w
            hitL = "slot"
        Case Else
            hitL = "none"
    End Select
End Function

Private Function hitS(t As String, w As Variant) As String
    Select Case t
        Case "s7"
            hitS = "lit"
        Case w
            hitS = "slot"
        Case Else
            hitS = "none"
    End Select
End Function

Private Function hitD(d As Double, w As Variant) As String
    Select Case d
        Case 1.5
            hitD = "lit"
        Case w
            hitD = "slot"
        Case Else
            hitD = "none"
    End Select
End Function

Private Function hitIs(n As Long, w As Variant) As String
    Select Case n
        Case Is > w
            hitIs = "gt"
        Case Else
            hitIs = "no"
    End Select
End Function

Private Function hitTo(d As Double, w As Variant) As String
    Select Case d
        Case 1 To w
            hitTo = "in"
        Case Else
            hitTo = "no"
    End Select
End Function


Sub Main()
    Debug.Print "S606-v-lit=" & hitV(7, 9)
    Debug.Print "S606-v-slot=" & hitV(5, 5)
    Debug.Print "S606-v-none=" & hitV(5, 6)
    Debug.Print "S606-l-lit=" & hitL(7, 9)
    Debug.Print "S606-l-slot=" & hitL(5, 5)
    Debug.Print "S606-l-none=" & hitL(5, 6)
    Debug.Print "S606-s-lit=" & hitS("s7", "q")
    Debug.Print "S606-s-slot=" & hitS("q", "q")
    Debug.Print "S606-s-none=" & hitS("q", "r")
    Debug.Print "S606-d-lit=" & hitD(1.5, 2.5)
    Debug.Print "S606-d-slot=" & hitD(0.5, 0.5)
    Debug.Print "S606-d-none=" & hitD(0.5, 2.5)
    Debug.Print "S606-is-gt=" & hitIs(9, 5)
    Debug.Print "S606-is-no=" & hitIs(3, 5)
    Debug.Print "S606-to-in=" & hitTo(3, 5)
    Debug.Print "S606-to-no=" & hitTo(7, 5)
    Debug.Print "S606-DONE"
End Sub
