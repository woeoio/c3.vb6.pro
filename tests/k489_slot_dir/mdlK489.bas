Attribute VB_Name = "mdlK489"
Option Explicit

' Target shape for 489 + the 474-F write half: the interface host declares
' `Property Let Label` BEFORE `Property Get Label`. Before 489 the minted table
' had ONE slot whose shape came from Let while the bound function came from Get
' (silent mismatch). Before the write half, `x.Label = ...` emitted
' vb6_IK489_prop_let_Label(x, ...) = the view struct into a vb6_cls_IK489* slot
' (cl C2440, no exe). Now both directions dispatch through the view vtable.
Sub Main()
    Dim x As IK489
    Dim n As Long
    Dim s As String
    Set x = New K489Holder
    x.Label = "abc"
    x.Label = "zzz"
    s = x.Label
    n = x.Num
    Debug.Print "K489-s=" & s
    Debug.Print "K489-n=" & n
    Debug.Print "K489-DONE"
End Sub
