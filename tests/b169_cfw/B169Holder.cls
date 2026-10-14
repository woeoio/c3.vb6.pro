Option Explicit

Public Enum B169Kind
    kText = 0
    kNumber = 1
End Enum

Public PV As Variant
Public PS As String
Public PL As Long
Public PE
Public PK As B169Kind

Public Sub StoreViaMe(v As String)
    Me.PS = v
End Sub

Public Sub StoreLitViaMe()
    Me.PS = "litme"
End Sub


Public Function ReadBack() As String
    ReadBack = Me.PS & "|" & CStr(Me.PL) & "|" & CStr(VarType(Me.PV)) & "|" & CStr(Me.PK)
End Function
