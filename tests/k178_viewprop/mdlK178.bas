Option Explicit

Public Rc As Long
Public Rn As Long
Public Beta As Long
Public BetaN As Long
Public LetHits As Long
Public t As String
Public w As Long

Public Sub Main()
    Dim x As IK178
    Dim s As String
    Dim n As Long
    Set x = New K178Alpha
    s = x.Label
    n = x.Num
    w = x.Twice(4)
    Debug.Print x.Label
    With x
        s = .Label
        n = .Num
    End With
    t = x.Label
    Set x = New K178Beta
    s = x.Label
    n = x.Num
    Debug.Print "K178-s=" & s
    Debug.Print "K178-n=" & n
    Debug.Print "K178-rc=" & Rc
    Debug.Print "K178-rn=" & Rn
    Debug.Print "K178-beta=" & Beta
    Debug.Print "K178-betan=" & BetaN
    Debug.Print "K178-w=" & w
    Debug.Print "K178-let=" & LetHits
    Debug.Print "K178-DONE"
End Sub
