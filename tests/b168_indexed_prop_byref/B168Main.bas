Option Explicit

' 账 #305 §11.150 (第一百六十八刀): 带索引的属性读当 ByRef 实参 —— 读-改-写。
'   S1/S2 = 单索引，早绑定 / 晚绑定接收者
'   S3    = 两个索引实参 (Let 的取值形参在最后)
'   S4    = String 档带索引 (临时量的 BSTR 归调用方退)
'   S5    = 索引本身是表达式
Public Sub Main()
    Dim b As B168Box
    Dim o As Object

    Set b = New B168Box
    Set o = b

    b.Slot(1) = 1
    b.Slot(2) = 10

    b.Bump b.Slot(2)
    Debug.Print "B168-S1=" & b.Slot(2)

    o.Bump b.Slot(2)
    Debug.Print "B168-S2=" & b.Slot(2)

    b.Slot(1) = 1
    b.Slot(2) = 10
    b.Bump b.Pair(1, 2)
    Debug.Print "B168-S3=" & b.Pair(1, 2)

    b.Tag(1) = "ab"
    b.App b.Tag(1)
    Debug.Print "B168-S4=" & b.Tag(1)

    b.Slot(3) = 30
    b.Bump b.Slot(1 + 1 + 1)
    Debug.Print "B168-S5=" & b.Slot(3)

    Debug.Print "B168-DONE"
End Sub
