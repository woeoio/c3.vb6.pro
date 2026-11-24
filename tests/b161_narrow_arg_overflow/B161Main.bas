Option Explicit

Public Sub Main()
    Dim box As B161Box
    Dim o As Object

    Set box = New B161Box
    Set o = box
    box.I = 7
    box.Y = 9
    box.L = 11

    On Error GoTo EH

    ' I1: Integer 属性合法写入
    o.I = 4000
    Debug.Print "B161-I1=ok " & box.I

    ' I2: Integer 属性越界 (VB6 错误 6，目标不变)
    o.I = 40000
    Debug.Print "B161-I2=no-error val=" & box.I
    GoTo I3
EH:
    Debug.Print "B161-I2=err" & Err.Number & " val=" & box.I
    Resume I3
I3:

    ' I3: Integer 属性负向越界
    On Error GoTo EH2
    o.I = -40000
    Debug.Print "B161-I3=no-error val=" & box.I
    GoTo Y1
EH2:
    Debug.Print "B161-I3=err" & Err.Number & " val=" & box.I
    Resume Y1
Y1:

    ' Y1: Byte 属性合法写入
    On Error GoTo EH3
    o.Y = 200
    Debug.Print "B161-Y1=ok " & box.Y

    ' Y2: Byte 属性越界
    o.Y = 300
    Debug.Print "B161-Y2=no-error val=" & box.Y
    GoTo L1
EH3:
    Debug.Print "B161-Y2=err" & Err.Number & " val=" & box.Y
    Resume L1
L1:

    ' L1: Long 属性照旧（宽档不受本刀影响）
    On Error GoTo 0
    o.L = 2000000000
    Debug.Print "B161-L1=ok " & box.L

    Debug.Print "B161-DONE"
    Set box = Nothing
    Set o = Nothing
End Sub
