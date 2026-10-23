Attribute VB_Name = "mdlMain"
Option Explicit

' 第一百六十六刀 / 新账 #464: 一枚 legacy 接口被两枚类 Implements。
' 改前每枚实现类各发一份同名 vb6_vtbl_IMul / vb6_iface_IMul / vb6_iface_IMul_wrap,
' 于是这个夹具在同一次编译里就是红的 (typedef C2011 + wrap 同名重定义)。
' 改后表由接口自己定一份, wrap 带实现者名, 少写的那一席发 438 桩。
Sub Main()
    Dim a As IMul
    Dim b As IMul
    Set a = New Alpha
    Set b = New Beta
    Debug.Print "A-ONE=" & a.One()
    Debug.Print "A-TWO=" & a.Two()
    Debug.Print "A-THREE=" & a.Three()
    Debug.Print "B-ONE=" & b.One()
    Debug.Print "B-TWO=" & b.Two()
    On Error Resume Next
    Dim r As Long
    r = b.Three()
    Debug.Print "B-THREE-ERR=" & Err.Number
    On Error GoTo 0
    Debug.Print "A-TNAME=" & TypeName(a)
    Debug.Print "B-TNAME=" & TypeName(b)
    Debug.Print "MULTI-DONE"
End Sub
