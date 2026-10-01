' Fix 161b-decl-out 夹具: Declare A 版 ByVal String 出参之二 (GetModuleFileNameA)。
' 判据: path=[<完整 exe 路径>]; 修复前只回读到空 (n 有值但 buf 未回写)。

Option Explicit

Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" _
    (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long

Public Sub Main()
    ' 同一条读数用**定长缓冲**再问一遍（账 #91 的另一半：String * N 当出参缓冲）。
    ' 定长串的 C 型是 BSTR，此前"回写"只能写到堆上那只可变串里；这里钉住两条同时成立。
    Dim bufF As String * 260
    Dim nF As Long
    nF = GetModuleFileName(0, bufF, 259)
    Debug.Print "nF="; nF
    If nF > 3 Then
        Debug.Print "gmn-fixed-ok=Y"
    Else
        Debug.Print "gmn-fixed-ok=N"
    End If

    Dim buf As String
    Dim n As Long
    buf = String$(260, 0)
    n = GetModuleFileName(0, buf, 260)
    Debug.Print "n="; n
    Debug.Print "path=["; Left$(buf, n); "]"
    If n > 3 Then
        Debug.Print "gmn-path-ok=Y"
    Else
        Debug.Print "gmn-path-ok=N"
    End If
End Sub
