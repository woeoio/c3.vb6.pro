' Account #611 knife 348: Array() element variant-kind fidelity.
' VB6 semantics: Array(True/False) is VT_BOOL(11), Null=1, Empty=0,
' Nothing=9 (VT_DISPATCH with a null pointer), a nested Array() is
' VT_ARRAY|VT_VARIANT(8204), and a date literal is VT_DATE(7).
' Before this knife every one of them went through vb6_ArraySetLong (VT_I4=3),
' and Nothing / a nested array pushed a pointer into an int32 parameter.
' C30-10 knife 36: an Integer-tier element (small literal / &H<=0xFFFF / 3% /
' `As Integer` variable / user Const / the Integer-returning builtins) is now
' boxed as VT_I2, which is what VB6 says: `Array(7)(0)` is VarType 2,
' TypeName "Integer" (cscript cross-check: AL-long=2, AL-i4=2). So AL-long and
' AL-i4 below answer 2, not 3.
' C30-11 第六十刀: 认得出档的实参（对象 / Nothing / Byte / Currency / Single / Date 值）
' 不再走"默认 Long"：真 VB6 的 Array(obj)(0) 是 VT_DISPATCH(9)、Array(CByte(65))(0) 是
' 17、CCur→6、CSng→4、CDate→7。AL-obj / AL-obj-is / AL-nothing-is / AL-byte / AL-cy /
' AL-sng / AL-date2 就是这七条读数。
' The plain BSTR/Double arms below are the zero-regression guard:
' their emitted setter must stay byte-for-byte what it was.
Option Explicit

Public Sub Main()
    Dim a As Variant
    Dim v As Variant
    a = Array(True, False, Null, Empty, Nothing)
    v = a(0): Debug.Print "AL-bt="; VarType(v)
    v = a(1): Debug.Print "AL-bf="; VarType(v)
    v = a(2): Debug.Print "AL-null="; VarType(v)
    v = a(3): Debug.Print "AL-empty="; VarType(v)
    v = a(4): Debug.Print "AL-nothing="; VarType(v)

    Dim b As Variant
    Dim w As Variant
    b = Array(Array(1, 2), 3)
    v = b(0): Debug.Print "AL-nested="; VarType(v)
    w = v(1): Debug.Print "AL-nested-val="; w
    v = b(1): Debug.Print "AL-long="; VarType(v)

    Dim d As Variant
    d = Array(#1/2/2000#)
    v = d(0): Debug.Print "AL-date="; VarType(v)

    Dim c As Variant
    c = Array(7, "abc", 1.5)
    v = c(0): Debug.Print "AL-i4="; VarType(v)
    v = c(1): Debug.Print "AL-bstr="; VarType(v)
    v = c(2): Debug.Print "AL-r8="; VarType(v)

    ' C30-11 第六十刀: 认得出档的实参不再落进"默认 vb6_ArraySetLong"。
    ' 真值取自 cscript (.temp/arr/oracle_arr2.vbs): CByte→17 / CCur→6 / CSng→4 / CDate→7；
    ' 对象档按 VB6 语义: Array(obj)(0) 是 VT_DISPATCH(9), IsObject=True（VBScript 传对象会取
    ' 默认属性，所以对象那两格不能拿 cscript 当真值；理由同 ai/039 §11.34）。
    Dim o As Object
    Dim e As Variant
    Dim f As Variant
    Set o = CreateObject("Scripting.Dictionary")
    e = Array(o)
    v = e(0): Debug.Print "AL-obj="; VarType(v)
    Debug.Print "AL-obj-is="; IsObject(e(0))
    f = Array(Nothing)
    Debug.Print "AL-nothing-is="; IsObject(f(0))
    e = Array(CByte(65))
    v = e(0): Debug.Print "AL-byte="; VarType(v)
    e = Array(CCur(1.2345))
    v = e(0): Debug.Print "AL-cy="; VarType(v)
    e = Array(CSng(1.5))
    v = e(0): Debug.Print "AL-sng="; VarType(v)
    e = Array(CDate("2020-01-02"))
    v = e(0): Debug.Print "AL-date2="; VarType(v)

    Debug.Print "AL-DONE"
End Sub
