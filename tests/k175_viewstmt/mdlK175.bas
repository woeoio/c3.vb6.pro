Attribute VB_Name = "mdlK175"
Option Explicit

' <vbeclipse> fixture driver (knife 175 / account 471 station D):
' a zero-arg method call written in STATEMENT position on a legacy
' interface VIEW. Before the fix this fell to vb6_ComCall(view, ...)
' = the whole struct into a void* slot = cl C2172 (24 sites in the
' external corpus).

Public Alpha As Long
Public Beta As Long

Public Sub Main()
    Dim x As IK175
    Set x = New K175Alpha
    x.Ping
    x.Ping
    Set x = New K175Beta
    x.Ping
    Dim t As Long
    t = x.Twice(5)
    Debug.Print "K175-alpha="; Alpha
    Debug.Print "K175-beta="; Beta
    Debug.Print "K175-t="; t
    Debug.Print "K175-DONE"
End Sub
