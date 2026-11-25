Option Explicit

' 账 #305 §11.152 (第一百七十刀): 属性读当 ByRef 实参 —— 型不一致那一格的**换算**。
'   T1 = Long 属性 → Integer 形参   (vb6_ChkInt 进来, vb6_ChkLong 回去)
'   T2 = Byte 属性 → Integer 形参
'   T3 = Integer 属性 → Long 形参
'   T4 = Integer 属性 → Byte 形参   (vb6_ChkByte 进来)
'   T5 = Single 属性 → Double 形参
'   T6 = Double 属性 → Long 形参    (进来走 vb6_FltToLng = VB6 的取整: 3.7 折 4, 不是截成 3)
'   T7 = Currency 属性 → Currency 形参 —— 证人: 本刀仍按值 (在册)
Public Sub Main()
    Dim b As B170Box

    Set b = New B170Box

    b.L1 = 40
    b.IncI b.L1
    Debug.Print "B170-T1=" & b.L1

    b.B1 = 7
    b.IncI b.B1
    Debug.Print "B170-T2=" & b.B1

    b.I1 = 20
    b.IncL b.I1
    Debug.Print "B170-T3=" & b.I1

    b.I1 = 20
    b.IncB b.I1
    Debug.Print "B170-T4=" & b.I1

    b.S1 = 1.5
    b.IncD b.S1
    Debug.Print "B170-T5=" & b.S1

    b.D1 = 3.7
    b.IncL b.D1
    Debug.Print "B170-T6=" & b.D1

    ' Currency 的拼串格式化另有一个旧病 (按 Long 发, 丢小数) —— 这里用比较避开它,
    ' 只问「写回没写回」。
    b.C1 = 10.5
    b.IncC b.C1
    If b.C1 = 10.5 Then
        Debug.Print "B170-T7=unchanged"
    Else
        Debug.Print "B170-T7=changed"
    End If

    Debug.Print "B170-DONE"
End Sub
