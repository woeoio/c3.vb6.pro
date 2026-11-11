Attribute VB_Name = "Mod1"
Option Explicit

Public Function CountOK(ByRef srcChannels() As PdChannel) As Long
    Dim i As Long
    Dim n As Long
    For i = 0 To UBound(srcChannels)
        If Not srcChannels(i).IsChannelOK Then
            n = n + 1
        End If
        If srcChannels(i).GetChannelDIBType = 2 Then
            n = n + 1
        End If
    Next i
    CountOK = n
End Function
