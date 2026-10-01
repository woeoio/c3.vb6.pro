' <vbeclipse>: 内置函数"VB6 合法最小实参形"合规模具。这一族此前是静默红:
'   Format(x)        -> vb6_Format(VARIANT, BSTR) 只发 1 参 → cl C2198 参数太少
'   Filter(a, "x")   -> 同上 (且首参没登记 Variant|Array 时 Variant 实参还 C2440)
'   GetAttr/SetAttr  -> RTL 只有定义没有声明 → cl C4013 隐式 int
' 本夹具的价值在**编译通过** (任何一条 C2198/C2440 会让整条用例红), 所以会弹窗的
' MsgBox/InputBox 两形放进恒假分支 (只进代码生成, 不执行); 其余运行期一律
' On Error Resume Next 兜住 (Seek/Dir 与机器状态有关)。needle 只取 locale 无关的读数。
Option Explicit

Public Sub Main()
    On Error Resume Next
    Dim s As String
    Dim v As Variant
    Dim n As Long
    Dim d As Date
    Dim a(3) As String
    s = "Hello World"
    a(0) = "abc"
    a(1) = "xyz"
    Dim guard As Boolean
    guard = (Len(s) < 0)
    If guard Then
        v = InputBox("p")
        v = InputBox("p", "t")
        v = MsgBox("m")
        v = MsgBox("m", 0)
    End If
    s = Command
    s = Environ("PATH")
    n = FreeFile
    n = Seek(1)
    s = Format(Now)
    Debug.Print "AR-fmt="; Format(2 + 2)
    Debug.Print "AR-fmt2="; Format(2 + 2, "0.0")
    s = String(3, 65)
    s = Space(3)
    v = Split(s)
    v = Split(s, " ")
    s = Join(a)
    s = Join(a, " ")
    v = Filter(a, "x")
    v = Filter(a, "x", True)
    s = Replace(s, "a", "b")
    s = Replace(s, "a", "b", 1)
    n = InStrRev(s, "l")
    n = StrComp(s, s)
    n = UBound(a)
    n = LBound(a)
    n = Len(s)
    d = CDate("2026-1-1")
    n = Year(d)
    s = Dir("c:\")
    Dim tp As String
    tp = Environ("TEMP") & "\c3_arity_probe.txt"
    Open tp For Output As #1
    Print #1, "x"
    Close #1
    n = GetAttr(tp)
    SetAttr tp, n
    Debug.Print "AR-attr="; (GetAttr(tp) = n)
    Kill tp
    v = IIf(True, 1, 2)
    v = Choose(1, "a")
    Debug.Print "AR-join1=["; Join(a); "]"
    Debug.Print "AR-filter=["; Join(Filter(a, "ab"), "|"); "]"
    Debug.Print "AR-DONE"
End Sub
