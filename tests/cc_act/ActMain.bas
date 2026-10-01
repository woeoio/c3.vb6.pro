Attribute VB_Name = "ActMain"
Option Explicit

' ai/022 B11/C05 = the in-project activation of a CoClass name: `As Circle`, `New Circle`
' and `CreateObject("ActApp.Circle")` all bind to the block's [Implementation] class.
' CC2/CC3 put the group name in a field, a parameter and a return type as well, because
' those are separate positions in the AST that the rewrite has to reach.

Private m_saved As Circle

Function MakeCircle() As Circle
    Dim t As Circle
    Set t = New Circle
    Set MakeCircle = t
End Function

' CC2 的槽: **ByRef 组名形参**。参数表把工程类名折成 Vb6Type::Variant (类名只留在
' typeRefName), 而定义侧发的是 `vb6_cls_ShapeAct** c` —— 调用点必须传**类指针的地址**。
' <vbeclipse> 回归: 修 cc_act 时曾把这个槽误当 Variant 槽装箱成 vb6_VARIANT*, ABI 不符,
' 被调方 c.Move 1 拿 Variant 当类指针解引用 → 段错误 (CC1 之后 rc=139)。
Sub UseCircle(c As Circle)
    c.Move 1
End Sub

' CC10 的槽: **真 ByRef Variant 形参** (typeRefName 为空)。同一个类实参走这里**必须**
' 仍然装箱 —— VB6 语义是非 Variant 实参拷进临时 Variant 再传址; 若也被"修"成直传裸
' 类指针, 被调方按 vb6_VARIANT 解析它 → 读垃圾。CC9/CC10 这一对是防"过度修复"的夹子。
Sub UseVariant(ByRef v As Variant)
    If v Is Nothing Then
        Debug.Print "CC10:FAIL nothing"
    Else
        Debug.Print "CC10:OK type=" & TypeName(v)
    End If
End Sub

Sub Main()
    Dim c As Circle
    Set c = New Circle
    c.Move 2
    If c.Area() = 6# Then
        Debug.Print "CC1:OK"
    Else
        Debug.Print "CC1:FAIL area=" & c.Area()
    End If

    Set m_saved = MakeCircle()
    m_saved.Move 4
    UseCircle m_saved
    If m_saved.Area() = 12# Then
        Debug.Print "CC2:OK"
    Else
        Debug.Print "CC2:FAIL area=" & m_saved.Area()
    End If

    Dim c2 As Circle
    Set c2 = c
    c2.Move 10
    If c.Area() = 26# Then
        Debug.Print "CC3:OK"
    Else
        Debug.Print "CC3:FAIL area=" & c.Area()
    End If

    Dim c4 As Circle
    Set c4 = CreateObject("ActApp.Circle")
    c4.Move 5
    If c4.Area() = 12# Then
        Debug.Print "CC4:OK"
    Else
        Debug.Print "CC4:FAIL area=" & c4.Area()
    End If

    Dim o5 As Object
    Set o5 = CreateObject("ActApp.Ring")
    If o5 Is Nothing Then
        Debug.Print "CC5:FAIL nothing"
    Else
        Debug.Print "CC5:OK"
    End If

    Dim r As Ring
    Set r = New Ring
    r.Move 1
    If r.Area() = -1# Then
        Debug.Print "CC6:OK"
    Else
        Debug.Print "CC6:FAIL area=" & r.Area()
    End If

    Dim c7 As Circle
    Set c7 = CreateObject("actapp.circle")
    c7.Move 1
    If c7.Area() = 4# Then
        Debug.Print "CC7:OK"
    Else
        Debug.Print "CC7:FAIL area=" & c7.Area()
    End If

    If c.Area() = 26# Then
        Debug.Print "CC8:OK"
    Else
        Debug.Print "CC8:FAIL area=" & c.Area()
    End If

    Dim s As ShapeAct
    Set s = c
    If s.Area() = 26# And s.Side() = 13# Then
        Debug.Print "CC9:OK"
    Else
        Debug.Print "CC9:FAIL area=" & s.Area() & " side=" & s.Side()
    End If

    UseVariant c

    Debug.Print "CC-DONE"
End Sub
