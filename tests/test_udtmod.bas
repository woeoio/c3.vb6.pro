Attribute VB_Name = "test_udtmod"
' 第一百一十二刀 (§11.94): 模块级 UDT 变量的静态初始化
' 修前: `Public uv As UV`（UV 里含 Variant 成员）发成 `vb6_type_UV uv = 0;` —— 结构体不能
'       用 `= 0` 初始化, cl 直接 C2074, 整个工程编不过（真机 VB6 里这是再普通不过的声明）。
'       根因: UDT 判定用 lookupDotted 查名字, 而 `Public Type` 只登记在**模块作用域**,
'       查不到就退 defaultValue(Unknown) = "0"。
' 修后: 模块作用域兜底 + `vb6_type_` 前缀兜底 ⇒ `= {0}`（全零 = Variant Empty / "" / 0）。
Option Explicit

Public Type UV
    f As Variant
End Type

Public Type Nest
    inner As UV
    s As String
    n As Long
End Type

Public uv As UV
Public nst As Nest
Private pv As UV

Public Sub Main()
    Dim e As Long
    Dim c As Object

    ' ① 零初始化: Variant 成员 = Empty（VarType 0 / IsEmpty True / IsNull False）
    Debug.Print "UDTMOD-1 VarType=" & CStr(VarType(uv.f)) & " IsEmpty=" & CStr(IsEmpty(uv.f)) & " IsNull=" & CStr(IsNull(uv.f))

    ' ② 能写能读（Variant 成员不是只读的死槽）
    uv.f = 7
    Debug.Print "UDTMOD-2 uv.f=" & CStr(uv.f) & " VarType=" & CStr(VarType(uv.f))

    ' ③ 嵌套 UDT + String / Long 成员也全零
    Debug.Print "UDTMOD-3 nst.inner.f=[" & CStr(nst.inner.f) & "] nst.s=[" & nst.s & "] nst.n=" & CStr(nst.n)

    ' ④ 写进嵌套成员再读回
    nst.inner.f = "x"
    nst.s = "abc"
    nst.n = 42
    Debug.Print "UDTMOD-4 " & CStr(nst.inner.f) & "|" & nst.s & "|" & CStr(nst.n)

    ' ⑤ Private 模块级同形（同一处发射, 只是链接性不同）
    Debug.Print "UDTMOD-5 IsEmpty(pv.f)=" & CStr(IsEmpty(pv.f))

    ' ⑥ 晚绑定 Set 进模块级 UDT 的 Variant 字段（§11.93 Q2 的模块级版本）
    '    注: 这里必须用一个**真会报 438** 的对象 (Scripting.Dictionary)；本机 Collection 的
    '    未知名成员现在静默返回空（见 §11.94 六、登记的一处独立缺口），拿它当靶子测不到守卫。
    Set c = CreateObject("Scripting.Dictionary")
    c.Add "k", "vv"
    uv.f = 111
    Err.Clear
    On Error Resume Next
    Set uv.f = c.NoSuchProperty
    e = Err.Number
    On Error GoTo 0
    ' 真机 probe111 Q2（同形, 槽在过程级）: uv.f=111 err=438
    Debug.Print "UDTMOD-6 v=" & CStr(uv.f) & " err=" & CStr(e)

    ' ⑦ 控制格: 成功的 Set 必须落槽
    Err.Clear
    On Error Resume Next
    Set uv.f = c
    e = Err.Number
    On Error GoTo 0
    Debug.Print "UDTMOD-7 isobj=" & CStr(IsObject(uv.f)) & " err=" & CStr(e)

    Debug.Print "UDTMOD-DONE"
End Sub
