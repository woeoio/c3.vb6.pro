Attribute VB_Name = "Module1"
' 账 #118: 定长 String (String * N) 的 VB6 语义。
'
' 真 VB6 口径 (仓库自带手册 Dim/Private/Static 语句的同一句 + MSDN 英文本 + 真机
' CopyMemory 实测, 三处一致), 两段别混:
'   · **赋值**: 不足 N 右侧补空格, 超过 N 截右 (与 LSet 同口径);
'     `Len(定长串)` **恒为 N**, 与当前装着几个字符无关。
'   · **初始化**: 未赋值时用 Chr$(0) 填充 (不是空格)。这一点只影响"赋值前读到什么",
'     Debug.Print 上看不见 (NUL 不可见), 故本用例不钉; 发码侧 vb6_BSTR_FixedSTR 已按 0 填充。
'
' 覆盖落点: 模块级 / 局部 / UDT 成员 / Function 返回 / Property Get 返回 / ByVal 定长形参 /
'           借用侧 (变量→变量) / 循环内反复收口。
'
' ⚠ **刻意不钉的一格**: 调用点实参侧的收口。VB6 里 `S "gh"` 传进 `ByVal x As String * 4`
'   时实参**也**被补到 4。收口必须落在调用点才有正确的所有权语义 (在过程体内补就要么漏
'   入参那块、要么释放调用方借来的串), 那需要"被调过程的形参定长表", 本批没做。
'   所以第 ⑦ 段只钉"形参在**过程体内**是字符串、Len 等于声明长度", 不钉实参补齐 ——
'   下面 `FS-i=[...]` 那行照实打出来, 但不做判据 (改完调用点收口, 它自然变成 "gh  ")。
Option Explicit

Private gT As String * 5

Type Rec
    f As String * 8
End Type

Function F() As String * 5
    F = "ab"
End Function

Property Get P() As String * 4
    P = "z"
End Property

Sub S(ByVal x As String * 4)
    ' 此前这里是 vb6_DebugWriteLong((int32_t)x) —— 把 BSTR 形参当**数值**打 (x64 还截指针)
    Debug.Print "FS-i-len=" & Len(x)
    Debug.Print "FS-i=[" & x & "]"
End Sub

Sub Main()
    Dim l As String * 5
    Dim r As Rec
    Dim v As String
    Dim i As Long

    ' ① 模块级赋短串 → 右侧补空格到 N; Len 恒为 N
    gT = "ab"
    Debug.Print "FS-a-len=" & Len(gT)
    Debug.Print "FS-a=[" & gT & "]"

    ' ② 模块级赋长串 → 截右到 N
    gT = "abcdefgh"
    Debug.Print "FS-b-len=" & Len(gT)
    Debug.Print "FS-b=[" & gT & "]"

    ' ③ 局部同款 (局部本来就初始化成定长, 但赋值一样要收口)
    l = "cd"
    Debug.Print "FS-c-len=" & Len(l)
    Debug.Print "FS-c=[" & l & "]"
    l = "cdefghij"
    Debug.Print "FS-d-len=" & Len(l)
    Debug.Print "FS-d=[" & l & "]"

    ' ④ UDT 定长成员
    r.f = "ef"
    Debug.Print "FS-e-len=" & Len(r.f)
    Debug.Print "FS-e=[" & r.f & "]"
    r.f = "efghijklmn"
    Debug.Print "FS-f-len=" & Len(r.f)
    Debug.Print "FS-f=[" & r.f & "]"

    ' ⑤ Function 返回值槽 (Function F() As String * 5)
    Debug.Print "FS-g-len=" & Len(F())
    Debug.Print "FS-g=[" & F() & "]"

    ' ⑥ Property Get 返回值槽 (Property Get P() As String * 4)
    Debug.Print "FS-h-len=" & Len(P)
    Debug.Print "FS-h=[" & P & "]"

    ' ⑦ ByVal 定长形参
    S "gh"

    ' ⑧ 借用侧 (变量→变量) 也要收口, 且不能存成别名 (改 v 不该改 l)
    v = "zz"
    l = v
    Debug.Print "FS-j-len=" & Len(l)
    Debug.Print "FS-j=[" & l & "]"
    v = "QQ"
    Debug.Print "FS-j2=[" & l & "]"

    ' ⑨ 循环内反复收口 (自有临时串): 不崩 + 末值正确
    For i = 1 To 200
        l = CStr(i)
    Next i
    Debug.Print "FS-k-len=" & Len(l)
    Debug.Print "FS-k=[" & l & "]"

    Debug.Print "FIXEDSTR-DONE"
End Sub
