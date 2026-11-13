Attribute VB_Name = "ModDiProbe"
Option Explicit

' 账 #610 第342刀（富贵竹 FIX-006 末节 LNK2019 那一族）的正式夹具。
'
' 上游原文 —— "编译清零后链接报: OpenDatabase x59、ZoomWindow x4、
' vb6_di_AttachThreadInput、vb6_di_IsIconic" ⇒ 本刀补 AttachThreadInput / IsIconic
' 两枚 di 桩（Win32 家族）。
'
' ⚠ 这两枚是**声明式**桩: C3 发码时把用户的 `Declare ... Lib "user32"` 名字动态拼成
' `vb6_di_<name>`（cgen_decl_api.cpp）⇒ 静态清单扫不出来, 只有真链一次才会以
' LNK2019 现形。所以本夹具的作用就是**把那两个符号链出来**。
'
' ⚠ 别在这两行 Declare 上加 Alias —— 漏 Alias / SDK 宏名会走另一条路径
' (见「最小夹具必须复刻真实 Declare 上下文」那条纪律)。本例刻意用最朴素的形式。
'
' 判据: 修前 (RTL 无这两枚桩) 链接报 LNK2019 vb6_di_AttachThreadInput /
'       vb6_di_IsIconic ⇒ 根本没有 exe; 修后 链通 + 跑得起来 (退出码 0)。

Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Boolean) As Boolean
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Boolean

Public Function ProbeAttachThreadInput(ByVal idAttach As Long, ByVal idAttachTo As Long) As Boolean
    ProbeAttachThreadInput = AttachThreadInput(idAttach, idAttachTo, True)
End Function

Public Function ProbeIsIconic(ByVal hWnd As Long) As Boolean
    ProbeIsIconic = IsIconic(hWnd)
End Function

Public Function ProbeBoth(ByVal hWnd As Long) As Boolean
    Dim r As Boolean
    r = AttachThreadInput(1, 2, False)
    If IsIconic(hWnd) Then
        r = True
    End If
    ProbeBoth = r
End Function

Public Sub Main()
    ' 假线程 ID / 空 HWND ⇒ 两个 API 都安全地返回 FALSE (不崩, 无副作用)。
    ' 判据是"链得出 exe 且跑得起来", 不是返回值本身。
    Dim r As Boolean
    r = ProbeBoth(0)
    If r Then
        Debug.Print "unexpected-true"
    End If
    Debug.Print "DI-OK"
End Sub
