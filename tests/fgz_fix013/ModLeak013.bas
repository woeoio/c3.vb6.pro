Attribute VB_Name = "ModLeak013"
Option Explicit

' 账 #610 第343刀: ByVal String 形参的**副本泄漏**（每调用泄一枚串）的正式夹具。
'
' 背景: 账 #610 第341刀（FIX-012）让被调方先拷副本（VB6 语义: 被调方拿到自己的副本），
'   但副本**在过程退出时不释放** ⇒ 每调用泄一枚 BSTR。
'   本刀修法: 副本另立一个变量, 过程出口只在「形参仍是那枚副本」时才 free。
'
' 四档:
'   A: Sub + ByVal String 形参, 体里不赋值 ⇒ 副本在出口被 free ⇒ 泄 0。
'   B: Sub + ByVal String 形参, 体里**赋值** ⇒ 体里那次 vb6_BSTR_AssignMove 已 free 旧值
'      ⇒ 形参已不是副本 ⇒ 出口跳过 ⇒ **不二次释放**（这一档是防 double free 的反面针）。
'   C: Function + ByVal String 形参 ⇒ **保守档**: 本刀不发 free（返回值可能就是该形参,
'      free 它会让返回值悬垂）⇒ 与修前逐字同形（零回归护栏）。
'   D: Sub + ByVal String 形参 + **Exit Sub** ⇒ 出口标签之后的释放段必须也走到
'      （Exit Sub 的 goto 落在标签上 ⇒ 标签之后的代码继续执行 ⇒ 副本仍被释放）。

Public Sub LeakSinkA(ByVal s As String)
    Dim n As Long
    n = Len(s)
End Sub

Public Sub LeakSinkB(ByVal s As String)
    s = "REPLACED"
End Sub

Public Function LeakFuncC(ByVal s As String) As String
    LeakFuncC = s
End Function

Public Sub LeakSinkD(ByVal s As String)
    If Len(s) > 0 Then
        Exit Sub
    End If
    s = "NEVER"
End Sub

Public Sub Main()
    Dim i As Long
    Dim r As String
    ' 循环足够多次: 修前每调用泄一枚 BSTR ⇒ 20 万次 x 32 字符 ≈ 数 MB ⇒ 工作集可量;
    ' 修后副本在出口释放 ⇒ 工作集平。CI 上只看「跑通 + 值正确」, 量化在本地探针做。
    For i = 1 To 200000
        LeakSinkA "0123456789ABCDEF0123456789ABCDEF"
    Next i
    LeakSinkB "0123456789ABCDEF0123456789ABCDEF"
    LeakSinkD "0123456789ABCDEF0123456789ABCDEF"
    r = LeakFuncC("hello")
    If r <> "hello" Then
        Debug.Print "LEAK13-BAD-C=" & r
        Exit Sub
    End If
    Debug.Print "LEAK13-OK=" & r
End Sub
