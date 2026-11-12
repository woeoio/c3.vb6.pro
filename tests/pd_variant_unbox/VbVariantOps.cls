VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbVariantOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 夹具: **Variant 表达式 → 具体型** 的六个消费点。每一处的 C 文本都不命中任何
' 「Variant 槽」前缀表, 此前一律不拆箱 ⇒ 整只 vb6_VARIANT 结构体递给
' BSTR/double/int32_t 形参 = C2440。
'   R1 Like  两操作数   → vb6_VariantToString   (PhotoDemon Evaluator.c 474)
'   R2 Mod   操作数     → vb6_VariantToLong     (PhotoDemon Evaluator.c 494)
'   R3 ^     操作数     → vb6_VariantToDouble   (PhotoDemon Evaluator.c 514)
'   R4 ParamArray 实参  → vb6_VariantToLong     (PhotoDemon pdSpinner.c 919 / TextSupport.c 357)
'   R5 形参 (Long)      → vb6_VariantToLong
'   R6 赋值 (Double)    → vb6_VariantToDouble

Public Function GetS() As Variant
    GetS = "abc"
End Function

Public Function GetD() As Variant
    GetD = 1.5
End Function

Public Function GetL() As Variant
    GetL = 7
End Function

' R1
Public Sub DoLike()
    Dim r As Boolean
    r = (GetS() Like "a*")
End Sub

' R2
Public Sub DoMod()
    Dim r As Long
    r = GetL() Mod 3
End Sub

' R3
Public Sub DoPow()
    Dim r As Double
    r = GetD() ^ 2
End Sub

' R4
Public Sub TakesPA(ParamArray args())
End Sub

Public Sub CallPA()
    TakesPA GetL(), GetL()
End Sub

' R5
Public Sub TakesLong(ByVal n As Long)
End Sub

Public Sub CallTakesLong()
    TakesLong GetL()
End Sub

' R6
Public Sub AssignToDouble()
    Dim d As Double
    d = GetD()
End Sub
