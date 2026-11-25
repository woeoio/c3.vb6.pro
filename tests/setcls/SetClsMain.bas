Attribute VB_Name = "SetClsMain"
' ?????? (§11.91): Set ???? Variant ??? -- ?? P ?(p109/P1..P4)
'   P1 = ?????? (C: c->f)     ?? c.f=111 err=438
'   P2 = ????????? (C: me->f) ?? c.f=111 err=438
'   P3/P4 = ???:??? Set ?????? (isobj=True err=0)
Option Explicit

Private Sub PE(ByVal tag As String, ByVal what As String, ByVal errNum As Long)
    Debug.Print tag & " " & what & " err=" & CStr(errNum)
End Sub

Sub Main()
    On Error Resume Next
    Dim o As Object
    Dim c As SetClsSlot
    Dim e As Long
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "vv"
    Set c = New SetClsSlot

    ' P1: ???,?????????? (C: c->f)
    c.f = 111
    Err.Clear
    Set c.f = o.NoSuchProperty
    e = Err.Number
    PE "P1-lateoutside", "c.f=" & CStr(c.f), e

    ' P2: ???,Set ????? (C: me->f)
    c.f = 111
    Err.Clear
    c.SetMissing o
    e = Err.Number
    PE "P2-lateinside", "c.f=" & CStr(c.f), e

    ' P3: ??? -- ?????? Set ????
    c.f = 111
    Err.Clear
    Set c.f = o
    e = Err.Number
    PE "P3-ctloutside", "isobj=" & CStr(IsObject(c.f)), e

    ' P4: ??? -- ?????? Set ????
    c.f = 111
    Err.Clear
    c.SetObj o
    e = Err.Number
    PE "P4-ctlinside", "isobj=" & CStr(IsObject(c.f)), e

    Debug.Print "P-DONE"
End Sub
