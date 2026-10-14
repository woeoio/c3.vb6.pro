Option Explicit

' <vbeclipse> regression clip (账 #305 §B168, knife 59): a Public Variant field reached
' from **outside** its own class used to be emitted as a completely unboxed raw store
' (`h->PV = 7;`) -> cl C2440 -> the whole project fails to build. The in-class form
' (`Me.PV = v`) compiles but stores the borrowed pointer and never reclaims the slot.
' The field-vs-property-vs-scalar distinction below is the negative control: PS/PL must
' keep the raw store, only the Variant slots go through the boxing station.

Public PV As Variant
Public PS As String
Public PL As Long
Public PW

Public Sub StoreViaMe(v As Variant)
    Me.PV = v
End Sub

Public Sub StoreLitViaMe()
    Me.PV = "insideme"
End Sub

Public Function ReadBack() As String
    ReadBack = CStr(Me.PV)
End Function
