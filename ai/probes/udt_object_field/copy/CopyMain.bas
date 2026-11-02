Attribute VB_Name = "CopyMain"
Option Explicit

' ai/probes/udt_object_field/copy - 问的是**整枚 UDT 赋值**（t2 = t1）时，里面那枚对象字段归谁。
'
' C3 这边发的是逐字节 memcpy（vb6_udtcpy_*），所以「拷贝到底带不带 AddRef」直接决定第317刀押后的
' 那一格能不能做：带 ⇒ 装订 / 退旧 / 出厂 / 读出四格是一条链，只做装订会把共享指针变成双重释放；
' 不带 ⇒ VB6 自己就是别名，我方那格「不 claim」反而是保真。
'
' 读数还是 INIT / TERM 与编号标记的先后。整段贴回即可。

Public gAll As String
Public gInit As Long
Public gTerm As Long

Private Type TObject
    O As Object
End Type

Private Type TClass
    O As Holder
End Type

Private Type TVariant
    O As Variant
End Type

Private Type TString
    S As String
End Type

Private Type TInner
    O As Object
End Type

Private Type TOuter
    N As TInner
    S As String
End Type

Public Sub Say(ByVal s As String)
    gAll = gAll & s & vbCrLf
    Debug.Print s
End Sub

Public Function NewH() As Holder
    gInit = gInit + 1
    Dim h As Holder
    Set h = New Holder
    h.Tag = gInit
    Say "INIT " & h.Tag
    Set NewH = h
End Function

' ---- A: 拷贝是否**自己持有一枚引用**（源放手之后还活着吗）----
Private Sub CaseA()
    Dim t1 As TObject
    Dim t2 As TObject
    Set t1.O = NewH()
    t2 = t1
    Say "A2-after-copy sameIs=" & (t2.O Is t1.O)
    Set t1.O = Nothing
    Say "A3-after-source-clear"
End Sub

' ---- A5: 目标本来就握着一枚，整枚赋值要不要先把它放掉 ----
Private Sub CaseA5()
    Dim a As TObject
    Dim b As TObject
    Set a.O = NewH()
    Set b.O = NewH()
    Say "A5-both-set"
    b = a
    Say "A6-after-struct-overwrite"
End Sub

' ---- G2/G3: 另外两种声明档的同一问（工程类 / Variant 装对象）----
Private Sub CaseG2()
    Dim a As TClass
    Dim b As TClass
    Set a.O = NewH()
    b = a
    Say "G2-sameIs=" & (b.O Is a.O)
    Set a.O = Nothing
    Say "G3-after-source-clear"
End Sub

Private Sub CaseG3()
    Dim a As TVariant
    Dim b As TVariant
    Set a.O = NewH()
    b = a
    Say "G4-sameIs=" & (b.O Is a.O)
    Set a.O = Nothing
    Say "G5-after-source-clear"
End Sub

' ---- F: 同一条拷贝路径上的字符串字段（同一位主人管的另一半）----
Private Sub CaseF()
    Dim a As TString
    Dim b As TString
    a.S = "AAA"
    b = a
    a.S = "BBB"
    Say "F1-b-string=" & b.S
End Sub

' ---- I: 嵌套 UDT（内层才是对象字段那一格）----
Private Sub CaseI()
    Dim a As TOuter
    Dim b As TOuter
    Set a.N.O = NewH()
    a.S = "outer"
    b = a
    Say "I1-copied-nested"
End Sub

' ---- E2: 数组元素之间整枚赋值 ----
Private Sub CaseE2()
    Dim arr() As TObject
    ReDim arr(1)
    Set arr(0).O = NewH()
    arr(1) = arr(0)
    Say "E2-element-copy sameIs=" & (arr(1).O Is arr(0).O)
    Set arr(0).O = Nothing
    Say "E3-after-clear-element-zero"
End Sub

Sub Main()
    gAll = ""
    gInit = 0
    gTerm = 0
    Say "P0-begin"

    CaseA
    Say "A4-after-sub"
    CaseA5
    Say "A7-after-sub"
    CaseG2
    Say "G6-after-sub"
    CaseG3
    Say "G7-after-sub"
    CaseF
    Say "F2-after-sub"
    CaseI
    Say "I2-after-sub"
    CaseE2
    Say "E4-after-sub"

    Say "Z-COUNT init=" & gInit & " term=" & gTerm
    Say "Z-DONE"

    Dim fn As Integer
    fn = FreeFile
    Open App.Path & "\probe_copy_out.txt" For Output As #fn
    Print #fn, gAll
    Close #fn
    MsgBox "init=" & gInit & " term=" & gTerm & vbCrLf & _
           "把 probe_copy_out.txt 或立即窗口整段贴回来（顺序就是读数）。", _
           vbOKOnly, "udt_object_field / copy"
End Sub
