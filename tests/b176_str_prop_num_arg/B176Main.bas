Option Explicit

' 账 #305 §11.156 (第一百七十三刀): String 属性读当 **数值** ByRef 实参。
'
' 改前这一格根本没走进登记道 (comPropReadWriteArg 见到 String 属性就 return false),
' 落到早绑定兜底: `vb6_IncL((&(int32_t){vb6_b_prop_get_S1(b)}));` —— 复合字面量的
' 地址, 被调方的写回落在临时量上就没了 (实测 b.S1 = "12", IncL 之后还是 12)。
' 改后两枚临时量 (BSTR 源 + 数值) + 写回时现造一枚 BSTR 给 Let (它归调用方退)。
'
' T1 = "12"  → Long    形参 (+1) ⇒ "13"  (走 vb6_CLng)
' T2 = "40"  → Integer 形参 (+1) ⇒ "41"  (走 vb6_CInt; CStr 档问声明型, 不按 C 型猜)
' T3 = "1.5" → Double  形参 (×2) ⇒ "3"   (走 vb6_CDbl)
' T4 = "10"  → Byte    形参 (+1) ⇒ "11"  (Byte 没有 BSTR 臂 ⇒ 先折 double 再收宽)
' T5 = 同档对照: Long 属性 → Long 形参 (+5) ⇒ 12
Public Sub IncL(ByRef x As Long)
    x = x + 1
End Sub

Public Sub IncI(ByRef x As Integer)
    x = x + 1
End Sub

Public Sub IncB(ByRef x As Byte)
    x = x + 1
End Sub

Public Sub DblTwice(ByRef d As Double)
    d = d * 2
End Sub

Public Sub AddFive(ByRef n As Long)
    n = n + 5
End Sub

Public Sub Main()
    Dim b As B176Box
    Set b = New B176Box

    b.S1 = "12"
    IncL b.S1
    Debug.Print "B176-1=" & b.S1

    b.S1 = "40"
    IncI b.S1
    Debug.Print "B176-2=" & b.S1

    b.S1 = "1.5"
    DblTwice b.S1
    Debug.Print "B176-3=" & b.S1

    b.S1 = "10"
    IncB b.S1
    Debug.Print "B176-4=" & b.S1

    b.N = 7
    AddFive b.N
    Debug.Print "B176-5=" & b.N

    Debug.Print "B176-DONE"
End Sub
