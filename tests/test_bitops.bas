Option Explicit

' <vbeclipse> 账 #216 = 数值位运算的结果类型收成 VB6 的口径。
' 权威现在只有一处：TypeSystem::bitwiseResult / TypeSystem::logicalNotResult，
' 语义层 (semantic_analyzer_expr) 与发码层 (CCodeGen::inferExprType) 都调它 ——
' 改前发码那一份 oracle 把 And/Or/Xor/Not 无条件推成 Boolean，于是位运算数一进字符串
' 上下文就打 True/False：
'   改前实测 (`.build/pcprobe/o_base.out`)：O1-lit=True / O7-cstr-var=True / O9-clipar=Tr /
'                                          O12-not=True / O15-mixed=True
'   改后同一台探针 (`.build/pcprobe/o_new.out`)：51 / 51 / 51 / -6 / -1 —— 与真 VB6 一致
' 关键的一点是**发码那头早就是对的**（Fix 039 把两侧转 Long 再做 |/&/^），
' 所以这条不是"算错数"而是"问错类型"：同一个表达式先赋给 Long 变量再打印（BF-assigned）
' 改前改后都是 51 —— 这也是为什么它在真工程里潜伏至今。
' 两头都要钉（缺一头就是假绿）：
'   · 数值侧按位运算数答（BF-or/BF-xor/BF-and/BF-not/BF-int/BF-mixed/BF-cstr/BF-left）
'   · 两侧都是 Boolean 时**仍然**是 Boolean（BF-bool-or / BF-bool-and），
'     以及条件位上数值位运算照旧当真假用（BF-cond）—— 拦"一路改回去全推成数值"那个反向错。
'   · Byte/Integer 那一档答 Integer、Eqv/Imp 同归位运算（BF-byte*/BF-eqv/BF-imp/BF-dblnot）
'     —— 这两条是两份 oracle 原本各写一份、答案不一样的地方，收成一处后必须钉住。

Public Sub Main()
    Dim a As Long, b As Long
    Dim i1 As Integer, i2 As Integer
    Dim t As Boolean, f As Boolean
    Dim acc As Long
    Dim bt As Boolean
    Dim by1 As Byte

    a = 34
    b = 51
    i1 = 1
    i2 = 2
    t = True
    f = False
    by1 = 85

    Debug.Print "BF-or-lit=" & (34 Or 51)
    Debug.Print "BF-or-var=" & (a Or b)
    Debug.Print "BF-xor=" & (a Xor b)
    Debug.Print "BF-and=" & (a And b)
    Debug.Print "BF-not=" & (Not 5)
    Debug.Print "BF-int-or=" & (i1 Or i2)
    Debug.Print "BF-mixed=" & (t Or 0)
    Debug.Print "BF-cstr=" & CStr(a Or b)
    Debug.Print "BF-left=" & Left(a Or b, 2)
    Debug.Print "BF-byte-and=" & (by1 And 240)
    Debug.Print "BF-byte-or=" & (by1 Or 15)
    Debug.Print "BF-byte-not=" & (Not by1)
    Debug.Print "BF-int-not=" & (Not i1)
    Debug.Print "BF-dblnot=" & (Not 2#)
    Debug.Print "BF-eqv=" & (5 Eqv 3)
    Debug.Print "BF-imp=" & (5 Imp 3)

    acc = (a Or b)
    Debug.Print "BF-assigned=" & acc
    Debug.Print "BF-bool-or=" & (t Or f)
    Debug.Print "BF-bool-and=" & (t And f)
    bt = (t And f)
    Debug.Print "BF-boolbox=" & bt
    If (a And 2) <> 0 Then
        Debug.Print "BF-cond=hit"
    Else
        Debug.Print "BF-cond=miss"
    End If
    If (t And f) Then
        Debug.Print "BF-boolcond=hit"
    Else
        Debug.Print "BF-boolcond=miss"
    End If
    Debug.Print "BF-sum=" & ((a Or b) + (a And b))
    Debug.Print "BF-DONE"
End Sub
