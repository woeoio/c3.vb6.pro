Option Explicit

' 账 #305 §11.153 (第一百七十一刀): Currency 属性在拼串里按 Long 格式化的病。
'   Q1 = 10.5 写进去, 读回来拼串 —— 病版本打 "10" (CStrLong 截成 int64)
'   Q2 = 读-改-写一轮 (b.C1 = b.C1 + 1) ⇒ 11.5
'   Q3 = Double 属性同一形 —— 对照组, 一直是 10.5 (它走 CStrDbl)
'   Q4 = 模块级 Function ... As Currency 的返回 —— 同一站, 补上返回值那一形
'   Q5 = CStr(b.C1) —— 显式 CStr 也要同一枚
'   Q6 = 局部 Currency 变量 —— 改前就对 (走 Double 档), 钉住不许回退
'   Q7 = With 块里的 .C1 —— 读那一形也要走同一枚

Public Function Cur() As Currency
    Cur = 20.25
End Function

Public Sub Main()
    Dim b As B171Box
    Set b = New B171Box

    b.C1 = 10.5
    Debug.Print "B171-Q1=" & b.C1

    b.C1 = b.C1 + 1
    Debug.Print "B171-Q2=" & b.C1

    b.D1 = 10.5
    Debug.Print "B171-Q3=" & b.D1

    Debug.Print "B171-Q4=" & Cur()

    Debug.Print "B171-Q5=" & CStr(b.C1)

    Dim c As Currency
    c = 7.25
    Debug.Print "B171-Q6=" & c

    With b
        Debug.Print "B171-Q7=" & .C1
    End With

    Debug.Print "B171-DONE"
End Sub
