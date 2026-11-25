Option Explicit

' 账 #305 §11.160 (第一百七十七刀): ByVal Currency 形参在**调用点**就折到 4 位。
'
' 探针 .temp/k177p（P177.bas，用 VB6 编译成 exe 后写文件）读数:
'   A1 1.234567 -> 1.2346      A4 0.00005 -> 0.0001
'   B1 Double 1.234567 -> 1.2346   B2 10.123456 -> 10.1235
'   E1 Variant 1.234567 -> 1.2346  F1 String "1.234567" -> 1.2346
'   H1 ByRef x = x*1.0001 -> 1.2346（折在被调方赋值站，不在调用点）
' 改前 C3 把 double 原值直接交给 C 侧同为 double 的形参（C 类型相同 ⇒ 不发换算），
' 于是 A1 打 1.234567、B2 打 10.123456。
'
' 判据两列: 第一列 = CStr(CLng(x * 10000))（存得准就是整数），
'           第二列 = CStr(x)（未折的尾巴一眼可见 —— 只钉第一列会被 CLng 的四舍五入遮住，
'           与 §11.159 的 A7 同一教训）。
Public Sub Main()
    ' A. 字面量直接交给 ByVal Currency 形参
    TakeByVal2 1.234567, "A1"
    TakeByVal2 1.2345, "A2"
    TakeByVal2 1.23455, "A3"
    TakeByVal2 0.00005, "A4"
    TakeByVal2 3.00005, "A5"
    TakeByVal2 (-1.234567), "A6"

    ' B. Double 变量
    Dim d As Double
    d = 1.234567: TakeByVal2 d, "B1"
    d = 10.123456: TakeByVal2 d, "B2"

    ' C. Currency 变量（已 4 位，不该被折第二次）
    Dim c As Currency
    c = 1.2345: TakeByVal2 c, "C1"
    c = 1.234567: TakeByVal2 c, "C2"

    ' D. 表达式 / 返回 Currency 的函数
    TakeByVal2 1.2 + 0.034567, "D1"
    TakeByVal2 RetCur(1.234567), "D2"

    ' E. Variant（运行期转 Currency 再折）
    Dim v As Variant
    v = 1.234567: TakeByVal2 v, "E1"
    v = 1.2345: TakeByVal2 v, "E2"

    ' F. String（String→Currency 那一站本来就折，不重复折）
    TakeByVal2 "1.234567", "F1"
    TakeByVal2 "3.00005", "F2"

    ' G. 整数档：交给 Currency 是精确的，不该动
    TakeByVal2 42, "G1"
    Dim n As Integer
    n = 7: TakeByVal2 n, "G2"

    ' H. ByRef：调用点不折，进入后赋值才折
    Dim r As Currency
    r = 1.2345
    BumpByRef r
    ShowC "H1", r

    Debug.Print "PROBE-DONE"
End Sub

Private Sub ShowC(ByVal tag As String, ByVal x As Currency)
    Debug.Print tag & "=" & CStr(CLng(x * 10000)) & " | " & CStr(x)
End Sub

Private Sub TakeByVal2(ByVal x As Currency, ByVal tag As String)
    ShowC tag, x
End Sub

Private Function RetCur(ByVal d As Double) As Currency
    RetCur = d
End Function

Private Sub BumpByRef(ByRef x As Currency)
    x = x * 1.0001
End Sub
