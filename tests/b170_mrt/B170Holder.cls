Option Explicit

Public PV As Variant
Public PZ As Variant
Public PS As String
Public Command As Long

Public Function ReadMe() As String
    ReadMe = CStr(Me.PV) & "|" & CStr(Me.PZ)
End Function

Public Function ReadBare() As String
    ReadBare = CStr(PV) & "|" & CStr(PZ)
End Function
