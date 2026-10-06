' test_lastdllerror.bas — Err.LastDllError 语义回归 (Fix <vbeclipse> 2026-10-06)
'
' VB6 语义 (MSDN ErrObject.LastDllError + 社区对 VB6 运行时的逆向, 见 tek-tips
' thread 222-475113 "GetLastError 在 VB 里恒为 0"): 每次由 VB6 代码里的 Declare
' 发起的 DLL 调用都会
'     1) 先把本线程 last-error 清 0   (SetLastError(0))
'     2) 调用 API
'     3) 返回后**立刻** GetLastError(), 存进 Err.LastDllError
' 由此三条可观测后果, 本夹具逐条钉住:
'     A. Err.LastDllError 是"调用那一刻的快照" —— 中间做多少 VB 字符串/打印操作
'        (它们都会踩全局 last-error) 都不改它。                              → LDL-survive
'     B. 调用前先清 0 ⇒ API 成功且不设错误码时报 0, 而不是沿用上一个错误。   → LDL-ok
'     C. Declare 形式的 GetLastError() 恒为 0 (清 0 之后没有 API 再设置它)。  → LDL-raw
'     D. Err.Clear 清空 Err 对象全部属性 ⇒ 快照也归 0 (文档属性表: LastDLLError 0)。→ LDL-cleared
' 改前 cgen_expr_member_precheck.inc / cgen_expr_with.cpp 把 Err.LastDllError 直接
' 映成**访问那一刻**的 GetLastError() —— A/B/C 三条全反 (A 读到被打印冲掉的垃圾值,
' B 读到上一个错误的残留, C 读到真错误码)。
'
' 同时覆盖两种调用点重定向形态 (包装函数在发码头里的落点不同):
'     GetModuleHandleXPA  —— 有 Alias 且与 VB 名不同 ⇒ 走 declareAliasMap_ (别名路)
'     GetCurrentProcessId —— 无 Alias ⇒ 走 `#ifndef <VB名>` 的 #define (非别名路)
' LDL-loop 钉的是"Declare 调用出现在 Do While 条件这种表达式位置"必须逐轮求值
' (曾经的错误实现会在调用点插语句, 把条件撕成非法的 `(Api(); ...() != 0)`)。
Declare Function GetModuleHandleXPA Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function GetLastError Lib "kernel32" () As Long

Public Sub Main()
    Dim s As String
    Dim i As Long
    Dim h As Long
    Dim n As Long
    Dim x As Long

    ' A 的前半: 一次必然失败的调用 → 快照 = 126 (ERROR_MOD_NOT_FOUND)
    h = GetModuleHandleXPA("NoSuchModuleLDL12345")
    Debug.Print "LDL-ret=" & h
    Debug.Print "LDL-fail=" & Err.LastDllError

    ' A 的后半: 200 次字符串拼接 + 打印之后, 快照必须原地不动
    For i = 1 To 200
        s = s & CStr(i)
    Next i
    Debug.Print "LDL-len=" & Len(s)
    Debug.Print "LDL-survive=" & Err.LastDllError

    ' B: 成功调用 (GetCurrentProcessId 不设错误码) → 0
    x = GetCurrentProcessId()
    Debug.Print "LDL-ok=" & Err.LastDllError

    ' C: Declare 形式的 GetLastError() 恒为 0
    Debug.Print "LDL-raw=" & GetLastError()

    ' 表达式位置: 条件里的 Declare 调用要逐轮求值 (跑 4 轮, 第 4 轮 Exit)
    n = 0
    Do While GetCurrentProcessId() <> 0
        n = n + 1
        If n > 3 Then Exit Do
    Loop
    Debug.Print "LDL-loop=" & n

    ' D: Err.Clear 清空 Err 对象**全部**属性 —— 含 LastDllError。
    ' VB6/VBA 文档 "Clear Method (Err Object)" 的 Clear 后属性表逐项写着 LastDLLError 0。
    ' 先造一个非 0 快照 (126), Clear 之后读必须归 0。
    h = GetModuleHandleXPA("NoSuchModuleLDL12345")
    Debug.Print "LDL-preclear=" & Err.LastDllError
    Err.Clear
    Debug.Print "LDL-cleared=" & Err.LastDllError

    Debug.Print "LDL-DONE"
End Sub
