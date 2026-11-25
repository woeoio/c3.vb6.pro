Option Explicit

' 账 #305 §11.158 (第一百七十五刀): Currency 是**定点 4 位** (VT_CY = 值×10000)。
'
' 改前赋值站直接存原值 —— 探针 .temp/k174/p1 的 P5 实测 `c = 10.123456` 打
' 10.123456, 而 VB6 存 10.1235 (第 5 位起圈掉)。C3 用 double 承载 Currency,
' 所以「圈」必须显式发, 且只能问 vb6_CCur 这一位主人 (它也是 CCur() 的出口)。
'
' 五处声明站 (局部 / 模块级 / 形参 / 类字段 / UDT 字段) 的 C 型与 Double 同型,
' 光按 C 类型串登记认不出 Currency ⇒ 新增 knownCurrencyVars_ (口径同 Fix 175 的
' knownDateVars_)。
'
' B178-11 是反向证人: Double 目标**不圈** (同一个 10.123456 照原样打出)。
Public Sub Main()
    Dim c As Currency
    c = 10.123456
    Debug.Print "B178-1=" & c

    c = -1.234567
    Debug.Print "B178-2=" & c

    c = 2.5
    Debug.Print "B178-3=" & c

    c = 0.00004
    Debug.Print "B178-4=" & c

    SetModCur 10.123456
    Debug.Print "B178-5=" & ModCur()

    Dim pc As Currency
    SetByRefCur pc
    Debug.Print "B178-6=" & pc

    Debug.Print "B178-7=" & CurFn()

    Dim u As B178Cur
    u.Cur = 10.123456
    Debug.Print "B178-8=" & u.Cur

    Dim bx As B178Box
    Set bx = New B178Box
    bx.Cur = 10.123456
    Debug.Print "B178-9=" & bx.Cur

    Dim lc As Currency
    Let lc = 10.123456
    Debug.Print "B178-10=" & lc

    Dim d As Double
    d = 10.123456
    Debug.Print "B178-11=" & d

    Debug.Print "B178-DONE"
End Sub

Private Type B178Cur
    Cur As Currency
End Type

Private m_cur As Currency

Private Function ModCur() As Currency
    ModCur = m_cur
End Function

Private Sub SetModCur(ByVal v As Currency)
    m_cur = v
End Sub

Private Sub SetByRefCur(ByRef c As Currency)
    c = 10.123456
End Sub

Private Function CurFn() As Currency
    CurFn = 10.123456
End Function
