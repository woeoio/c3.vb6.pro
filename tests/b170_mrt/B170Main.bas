Attribute VB_Name = "B170Main"
Option Explicit

Sub Main()
    Dim h As New B170Holder
    Dim k As New B170Holder2
    Dim s As String
    s = "withform"
    h.PV = s
    h.PZ = s
    h.Command = 7
    k.PV = s
    Debug.Print "B170-1=" & CStr(h.PV) & "|" & Len(CStr(h.PV))
    Debug.Print "B170-2=" & CStr(h.PZ) & "|" & Len(CStr(h.PZ))
    Debug.Print "B170-3=" & CStr(k.PV) & "|" & Len(CStr(k.PV))
    Debug.Print "B170-4=" & CStr(h.Command) & "|" & Len(CStr(h.Command))
    With h
        Debug.Print "B170-5=" & CStr(.PV) & "|" & CStr(.PZ)
    End With
    Debug.Print "B170-6=" & h.ReadMe & "|" & h.ReadBare
    Debug.Print "B170-7=" & VarType(h.PV) & "|" & TypeName(h.PV)
    h.PV = 123
    Debug.Print "B170-8=" & CStr(h.PV)
    Debug.Print "B170-DONE"
End Sub
