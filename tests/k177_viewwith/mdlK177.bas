Attribute VB_Name = "mdlK177"
Option Explicit

' <vbeclipse> fixture driver (knife 177 / account 471 station D-prime):
' a With block whose RECEIVER is a legacy interface VIEW. Before the
' fix the With temp was declared void* and initialised with (void*)view,
' i.e. the whole value struct into a pointer slot = cl C2440, which
' killed the whole block before any member was even looked at.

Public Alpha As Long
Public Beta As Long

Public Sub Main()
    Dim x As IK177
    Set x = New K177Alpha
    With x
        .Ping
        .Ping
        Dim t1 As Long
        t1 = .Twice(3)
    End With
    Set x = New K177Beta
    With x
        .Ping
    End With
    Dim t2 As Long
    t2 = x.Twice(5)
    Debug.Print "K177-alpha="; Alpha
    Debug.Print "K177-beta="; Beta
    Debug.Print "K177-t1="; t1
    Debug.Print "K177-t2="; t2
    Debug.Print "K177-DONE"
End Sub
