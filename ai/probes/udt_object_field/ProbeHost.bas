Attribute VB_Name = "ProbeHost"
Option Explicit

' ai/probes/udt_object_field - C3 账 #600 的运行期读数（只需 VB6，或 twinBASIC 也可，注明即可）。
'
' 这一份只碰**字段级**的动作（Set / 覆写 / 置空 / 出过程 / Erase / ByRef），
' 整枚 UDT 赋值（t2 = t1）那一族住在 copy/ 那一个工程里 —— 因为那句在 VB6 里有可能根本不许编译，
' 那本身就是答案，不该把这一份一起拖倒。
'
' 读数 = INIT / TERM 与编号标记的**先后顺序**。请整段贴回（立即窗口或 probe_out.txt 都行）。
' 少了哪一段、或者中途崩了，也请照实说「跑到 X 之后没了」。

Public gAll As String
Public gInit As Long
Public gTerm As Long

Private Type TObject           ' 声明档一：泛对象
    O As Object
End Type

Private Type TClass            ' 声明档二：工程类
    O As Holder
End Type

Private Type TVariant          ' 声明档三：Variant 里装对象
    O As Variant
End Type

Private Type TTwo              ' 同一枚对象被两个字段指着
    A As Object
    B As Object
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

' ---- B: 覆写字段时，旧的那一枚是否当场被放手 ----
Private Sub CaseB()
    Dim t As TObject
    Say "B1-begin"
    Set t.O = NewH()
    Say "B2-assigned"
    Set t.O = NewH()
    Say "B3-after-overwrite"
End Sub

' ---- C: Set <字段> = Nothing 是否当场退 ----
Private Sub CaseC()
    Dim t As TObject
    Say "C1-begin"
    Set t.O = NewH()
    Say "C2-assigned"
    Set t.O = Nothing
    Say "C3-after-nothing"
End Sub

' ---- D: 局部 UDT 出过程，字段里那一枚归谁放（三种声明档各问一次）----
Private Sub CaseD1()
    Dim t As TObject
    Set t.O = NewH()
    Say "D1-in-scope"
End Sub

Private Sub CaseD2()
    Dim t As TClass
    Set t.O = NewH()
    Say "D2-in-scope"
End Sub

Private Sub CaseD3()
    Dim t As TVariant
    Set t.O = NewH()
    Say "D3-in-scope"
End Sub

' ---- E: 动态数组的 Erase / 重新 ReDim 对元素字段做了什么 ----
Private Sub CaseE()
    Dim arr() As TObject
    ReDim arr(1)
    Say "E1-after-redim"
    Set arr(0).O = NewH()
    Set arr(1).O = NewH()
    Say "E2-two-elements-set"
    Erase arr
    Say "E3-after-erase"
    ReDim arr(2)
    Set arr(0).O = NewH()
    Say "E4-element-set-again"
End Sub

' ---- H: 从字段读出来交给一枚普通对象变量，字段随后置空 ----
Private Sub CaseH()
    Dim t As TObject
    Dim g As Object
    Say "H1-begin"
    Set t.O = NewH()
    Set g = t.O
    Say "H2-read-out"
    Set t.O = Nothing
    Say "H3-field-cleared"
    Say "H4-tag-through-g=" & g.Tag
    Set g = Nothing
    Say "H5-after-release"
End Sub

' ---- M: 两个字段指着同一枚对象，放手几次 ----
Private Sub CaseM()
    Dim t As TTwo
    Dim h As Holder
    Set h = NewH()
    Set t.A = h
    Set t.B = h
    Say "M1-two-fields-one-object"
    Set h = Nothing
    Say "M2-after-extra-release"
End Sub

' ---- J: ByRef 传 UDT 给子过程，子过程里置空 ----
Private Sub CalleeJ(ByRef u As TObject)
    Say "J1-inside-callee"
    Set u.O = Nothing
    Say "J2-after-clear-inside"
End Sub

Private Sub CaseJ()
    Dim t As TObject
    Set t.O = NewH()
    Say "J3-before-call"
    CalleeJ t
    Say "J4-after-call caller-field-is-Nothing=" & (t.O Is Nothing)
    Set t.O = NewH()
    Say "J5-refilled"
End Sub

' ---- N: 没赋值那一档的控制组（新 Dim 出来的字段是不是干净的）----
Private Sub CaseN()
    Dim t As TObject
    Say "N1-fresh-is-Nothing=" & (t.O Is Nothing)
    Set t.O = NewH()
    Say "N2-set"
End Sub

Sub Main()
    gAll = ""
    gInit = 0
    gTerm = 0
    Say "P0-begin"

    CaseB
    Say "B4-after-sub"
    CaseC
    Say "C4-after-sub"
    CaseD1
    Say "D1-after-sub"
    CaseD2
    Say "D2-after-sub"
    CaseD3
    Say "D3-after-sub"
    CaseE
    Say "E5-after-sub"
    CaseH
    Say "H6-after-sub"
    CaseM
    Say "M3-after-sub"
    CaseJ
    Say "J6-after-sub"
    CaseN
    Say "N3-after-sub"

    Say "Z-COUNT init=" & gInit & " term=" & gTerm
    Say "Z-DONE"

    Dim fn As Integer
    fn = FreeFile
    Open App.Path & "\probe_out.txt" For Output As #fn
    Print #fn, gAll
    Close #fn
    MsgBox "init=" & gInit & " term=" & gTerm & vbCrLf & _
           "把 probe_out.txt 或立即窗口整段贴回来（顺序就是读数）。", vbOKOnly, "udt_object_field"
End Sub
