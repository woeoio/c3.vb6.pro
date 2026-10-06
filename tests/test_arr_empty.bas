' <vbeclipse> 数组"空"与"缺少数组"两形对表 (用户在真 VB6 里实测确认):
'   Array()            -> 空数组: UBound=-1 / LBound=0 / UBound-LBound+1=0 / For Each 零次
'   UBound(<未分配>)   -> 运行时错误 9 (工程靠 On Error 接住, 见 List.cls)
' 常量实参那一形 (`UBound(vbNull)`) 是**编译期**错误, 见 tests/arr_neg/n01_ubound_vbnull.bas。
Option Explicit

' <vbeclipse> 账 #209: 元素访问那一半 (下面 Ea* 七枚私有过程共用一枚 UDT)。
Private Type EaPair
    A As Long
    B As Long
End Type

Public Sub Main()
    Dim a()
    a = Array()
    Debug.Print "EA-ub="; UBound(a)
    Debug.Print "EA-lb="; LBound(a)
    Debug.Print "EA-n="; UBound(a) - LBound(a) + 1
    Debug.Print "EA-isarr="; IsArray(a)
    Debug.Print "EA-join=["; Join(a, ","); "]"

    Dim n As Long
    n = 0
    Dim x As Variant
    For Each x In a
        n = n + 1
    Next
    Debug.Print "EA-fe="; n

    Dim v As Variant
    v = Array()
    Debug.Print "EA-vub="; UBound(v)

    Dim w()
    w = Array(7, 8, 9)
    Debug.Print "EA-3ub="; UBound(w)
    Debug.Print "EA-e2="; w(2)

    Dim s() As String
    ReDim s(1 To 3)
    Debug.Print "EA-rb-lb="; LBound(s)
    Debug.Print "EA-rb-ub="; UBound(s)

    ' <vbeclipse> 账 #209: 动态数组的**元素**访问在"未分配 / 越界"两形下 VB6 都是运行时
    ' 错误 9; 本刀之前这里裸读描述符 (NULL+0xc) -> 原生 0xC0000005
    ' (实测 Charts 2020 ucChartBar demo 点 Random, 三次同偏移 0x1abf2)。
    EaNullRead 1
    EaNullWrite 2
    EaAbove 3
    EaBelow 4
    EaUdtWith 5
    EaAfterErase 6
    EaLoop 7

    Dim u() As Long
    On Error GoTo Miss
    Debug.Print "EA-never="; UBound(u)
    Debug.Print "EA-DONE"
    Exit Sub
Miss:
    Debug.Print "EA-err="; Err.Number
    Debug.Print "EA-DONE"
End Sub

' ============================================================
' <vbeclipse> 账 #209: 元素访问的"未分配 / 越界"两形。
' 每枚形状单独一个过程: On Error 处理器打完 Err.Number 就退出, 不让 Err 状态
' 漏给下一枚 (和上面 Main 里 Miss: 那一段同一个写法)。
' EA-ea-in / EA-ea-loop 那两枚是**负控** —— 在范围内的读数必须照旧出来,
' 检查若误报就会红在它们身上。
' ============================================================

Private Sub EaNullRead(ByVal n As Long)
    Dim u() As Long
    Dim got As Long
    On Error GoTo H
    got = u(0)
    Debug.Print "EA-ea-nullread=NOERR "; got
    Exit Sub
H:
    Debug.Print "EA-ea-nullread="; Err.Number
End Sub

Private Sub EaNullWrite(ByVal n As Long)
    Dim u() As Long
    On Error GoTo H
    u(0) = 5
    Debug.Print "EA-ea-nullwrite=NOERR"
    Exit Sub
H:
    Debug.Print "EA-ea-nullwrite="; Err.Number
End Sub

Private Sub EaAbove(ByVal n As Long)
    Dim m() As Long
    Dim got As Long
    ReDim m(2 To 5)
    m(2) = 20
    m(3) = 30
    m(5) = 50
    On Error GoTo H
    Debug.Print "EA-ea-in=" & m(2) & "/" & m(3) & "/" & m(5)
    got = m(6)
    Debug.Print "EA-ea-above=NOERR "; got
    Exit Sub
H:
    Debug.Print "EA-ea-above="; Err.Number
End Sub

Private Sub EaBelow(ByVal n As Long)
    Dim m() As Long
    Dim got As Long
    ReDim m(2 To 5)
    On Error GoTo H
    got = m(1)
    Debug.Print "EA-ea-below=NOERR "; got
    Exit Sub
H:
    Debug.Print "EA-ea-below="; Err.Number
End Sub

Private Sub EaUdtWith(ByVal n As Long)
    Dim s() As EaPair
    ReDim s(0 To 1)
    With s(0)
        .A = 11
        .B = 22
    End With
    Debug.Print "EA-ea-udt-in=" & s(0).A & "/" & s(0).B & "/" & s(1).A & "/" & s(1).B
    On Error GoTo H
    With s(2)
        .A = 1
    End With
    Debug.Print "EA-ea-udt=NOERR"
    Exit Sub
H:
    Debug.Print "EA-ea-udt="; Err.Number
End Sub

Private Sub EaAfterErase(ByVal n As Long)
    Dim q() As Long
    Dim got As Long
    ReDim q(0 To 3)
    q(2) = 8
    Debug.Print "EA-ea-before=" & q(2)
    Erase q
    On Error GoTo H
    got = q(2)
    Debug.Print "EA-ea-erase=NOERR "; got
    Exit Sub
H:
    Debug.Print "EA-ea-erase="; Err.Number
End Sub

Private Sub EaLoop(ByVal n As Long)
    Dim v() As Long
    Dim i As Long
    Dim sum As Long
    ReDim v(1 To 4)
    For i = 1 To 4
        v(i) = i * 10
    Next i
    For i = LBound(v) To UBound(v)
        sum = sum + v(i)
    Next i
    Debug.Print "EA-ea-loop="; sum
End Sub
