Option Explicit

' 账 #611 第349刀: 多维数组经真外部 COM (Scripting.Dictionary) 往返。
' 修前: 出站 vb6_ComPackArrayOut 的 ND 支只收 vb6_ComSaVtBlittable 元素档,
'       vb6_sa_variant 直接 return => VT_EMPTY; 入站 vb6_ComVariantIntoSlot /
'       vb6_VariantArrayIntoSlot 又只认 cDims==1。留证探针
'       probe_DictArrayRoundTrip.bas 的修前读数 = v1d=8204 / v2d=0。
' 修后: ND variant 逐枚带权装箱 + 唯一 ND 转换出口 (vb6_ComNDArrayIntoSlot),
'       维界 / 列主序 / 元素档全保真。
' 本机 sa_dim_probe 实测的维序口径: SafeArrayCreate 的 rg[d] 正序对应
' SafeArrayGetBound(dim=d+1), 线性布局 d=0 最快变化, 与 C3 vb6_SafeArrayND
' 同序 => 不重排。下面 NC-b1/NC-b2/NC-e00/NC-e12 就是钉这一条的。
Sub Main()
    Dim d As Object
    Dim a2(0 To 1, 0 To 2) As Variant
    Dim a3(0 To 1, 0 To 1, 0 To 1) As Variant
    Dim b2(0 To 1, 0 To 1) As Long
    Dim a1 As Variant
    Dim v As Variant
    Dim w As Variant
    Dim x As Variant
    Dim y As Variant
    Set d = CreateObject("Scripting.Dictionary")

    ' --- 1D Variant 零回归护栏 (FIX-011/012 那条既有通路不许被碰坏) ---
    a1 = Array("x", "y")
    d.Add "one", a1
    y = d("one")
    Debug.Print "NC-1d=" & VarType(y) & "/" & UBound(y, 1)

    ' --- 2D Variant: 维界 + 列主序 + 混合元素档 ---
    a2(0, 0) = "k1":  a2(1, 0) = "v1"
    a2(0, 1) = 42:    a2(1, 1) = True
    a2(0, 2) = Null:  a2(1, 2) = "v3"
    d.Add "two", a2
    v = d("two")
    Debug.Print "NC-vt=" & VarType(v)
    Debug.Print "NC-b1=" & LBound(v, 1) & "/" & UBound(v, 1)
    Debug.Print "NC-b2=" & LBound(v, 2) & "/" & UBound(v, 2)
    Debug.Print "NC-e00=" & v(0, 0)
    Debug.Print "NC-e10=" & v(1, 0)
    Debug.Print "NC-e01=" & VarType(v(0, 1)) & "/" & v(0, 1)
    Debug.Print "NC-e11=" & VarType(v(1, 1))
    Debug.Print "NC-e02=" & VarType(v(0, 2))
    Debug.Print "NC-e12=" & v(1, 2)

    ' --- 3D Variant: 秩数泛化 (不是只把 2 写死) ---
    a3(0, 0, 0) = "a": a3(1, 1, 1) = "h"
    d.Add "three", a3
    w = d("three")
    Debug.Print "NC-3vt=" & VarType(w)
    Debug.Print "NC-3b3=" & LBound(w, 3) & "/" & UBound(w, 3)
    Debug.Print "NC-3e=" & w(0, 0, 0) & w(1, 1, 1)

    ' --- 2D Long: 既有 blittable memcpy 支零回归护栏 ---
    b2(0, 0) = 11: b2(1, 0) = 22
    b2(0, 1) = 33: b2(1, 1) = 44
    d.Add "long2d", b2
    x = d("long2d")
    Debug.Print "NC-lvt=" & VarType(x)
    Debug.Print "NC-le=" & x(1, 1)
    Debug.Print "NC-DONE"
End Sub
