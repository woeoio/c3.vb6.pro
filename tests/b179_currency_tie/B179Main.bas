Option Explicit

' 账 #305 §11.159 (第一百七十六刀): Currency 折数**不许两次舍入**。
'
' 改前 vb6_CCur 是 `round(v * 10000.0) / 10000.0` —— 乘积本身先被舍进 double。
' 探针 .temp/k176/p1 在真 VB6 里量到的六格恰好都落在「乘积等于 .5」上:
'
'   A5  3.00005  VB6=30000  改前 C3=30001
'   A8  2.00005  VB6=20000  改前 C3=20001
'   A10 -3.00005 VB6=-30000 改前 C3=-30001
'   A12 2.00025  VB6=20002  改前 C3=20003
'   A13 4.00005  VB6=40000  改前 C3=40001
'   A7  0.00005  VB6=1      改前 C3=1      (真值本来就在半分之上的那一格)
'
' A7 与其余五格的差别正是本条的病征: 不是「取整方向变了一半」, 而是**舍的是乘积
' 而不是原值**。判据用 `CStr(CLng(c * 10000))` —— 存得准就是整数。
Public Sub Main()
    Dim c As Currency

    c = 10.123456:  ShowN "B179-1", c
    c = -1.234567:  ShowN "B179-2", c
    c = 2.5:        ShowN "B179-3", c
    c = 0.00004:    ShowN "B179-4", c

    c = 3.00005:    ShowN "B179-5", c
    c = 1.00005:    ShowN "B179-6", c
    c = 0.00005:    ShowN "B179-7", c
    c = 2.00005:    ShowN "B179-8", c
    c = 1.00015:    ShowN "B179-9", c
    c = -3.00005:   ShowN "B179-10", c
    c = 1.23455:    ShowN "B179-11", c
    c = 2.00025:    ShowN "B179-12", c
    c = 4.00005:    ShowN "B179-13", c

    Debug.Print "B179-DONE"
End Sub

Private Sub ShowN(ByVal tag As String, ByVal x As Currency)
    Debug.Print tag & "=" & CStr(CLng(x * 10000))
End Sub
