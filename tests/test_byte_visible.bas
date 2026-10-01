' 账 #123 夹具: 局部/形参/局部 Const 的 As Byte 在类型推断里以前**完全不可见**。
'
' 根因: As Byte 的 C 型是 uint8_t, 而 local/param 的登记分支只匹配
'   BSTR | int32_t | int16_t | VBABOOL | intptr_t | vb6_VARIANT ⇒ 局部 Byte 进不了任何
'   knownXxxVars_ ⇒ inferExprType 答 Unknown, isDefinitelyVariantExpr 又回退去查符号表。
' 后果实测三件 (修复前):
'   1) `s = 局部Byte` 发 vb6_BSTR_Assign(&s, bt) —— 数字被当 BSTR 指针存, 跑起来 SIGSEGV;
'   2) `(bt = 65)` 发 vb6_VarCmpLongEq(&bt, 65) —— 该助手签名是 (vb6_VARIANT*, int32_t),
'      拿 1 字节对象的地址当 16 字节 VARIANT 传 ⇒ **今天就在给错答案** (实测返回 False);
'   3) 装箱那形顺手修了一条**预存**的错: wrapVariantValue 的 Byte 档以前发
'      vb6_VariantLong ⇒ VT_I4=3, 而 VB6 要 VT_UI1=17 —— 这条**模块级 Byte 与 CByte() 早就中**
'      (实测修复前 `v = 模块级Byte` / `v = CByte(67)` 都读回 3), 局部 Byte 反而因为不可见
'      掉进 _Generic 而凑对 (17)。所以本格把三形一起对齐到 17。
'
' 读数一律不带方括号 (套件用 PowerShell -like 比针, [ABC] 在那里是字符类 —— 门 #178 的教训)。

Option Explicit

Dim gb As Byte

Function PackIt(b As Byte) As Long
    Dim v As Variant
    v = b
    PackIt = VarType(v)
End Function

' ByVal Byte 形参: 形参侧的登记与局部同口径
Function CmpByVal(ByVal b As Byte) As String
    If b = 200 Then CmpByVal = "EQ200" Else CmpByVal = "NE200"
End Function

' ByRef Byte 形参: 写穿之后调用方读回, 且写进去的值仍按 Byte 比较
Sub FlipByRef(ByRef b As Byte)
    b = 201
End Sub

Sub Main()
    Dim bt As Byte
    Dim v As Variant
    Dim s As String
    Const CB As Byte = 70

    bt = 65: gb = 66

    ' --- 比较面 (以前静默给错答案) ---
    Debug.Print "B1-cmp-eq-true="; (bt = 65)
    Debug.Print "B2-cmp-eq-false="; (bt = 66)
    Debug.Print "B3-cmp-lt="; (bt < 100)
    Debug.Print "B4-cmp-module-eq="; (gb = 66)
    Debug.Print "B5-cmp-const-eq="; (CB = 70)

    ' --- 赋值到 String 面 (以前 SIGSEGV) ---
    s = bt
    Debug.Print "B6-str-local="; s
    s = gb
    Debug.Print "B7-str-module="; s
    s = CB
    Debug.Print "B8-str-const="; s
    Debug.Print "B9-str-len="; Len(s)

    ' --- Variant 装箱面 (三形都要 VT_UI1=17) ---
    v = bt
    Debug.Print "B10-pack-local="; VarType(v)
    v = gb
    Debug.Print "B11-pack-module="; VarType(v)
    v = CB
    Debug.Print "B12-pack-const="; VarType(v)
    v = CByte(67)
    Debug.Print "B13-pack-cbyte="; VarType(v)
    Debug.Print "B14-pack-param="; PackIt(68)
    Debug.Print "B15-packed-val="; CByte(v)

    ' --- 形参面 ---
    Debug.Print "B16-param-byval-eq="; CmpByVal(200)
    Debug.Print "B17-param-byval-ne="; CmpByVal(65)
    bt = 5
    FlipByRef bt
    Debug.Print "B18-byref-writethru="; bt
    Debug.Print "B19-byref-cmp="; (bt = 201)

    ' --- 不许被牵连的面: 算术 / Debug.Print / 与 Long 混算 ---
    Dim n As Long
    n = 1000
    Debug.Print "B20-arith="; bt + 3
    Debug.Print "B21-print="; gb
    Debug.Print "B22-mixed-lt="; (gb < n)
    Debug.Print "B23-mixed-sum="; gb + n

    ' --- 反复装箱不泄漏/不损坏 (所有权判错当场就崩) ---
    Dim i As Long
    For i = 1 To 20000
        v = bt + 1
    Next i
    Debug.Print "B24-loop-pack="; VarType(v)
    Debug.Print "B25-loop-val="; CByte(v)
    Debug.Print "B26-DONE"
End Sub
