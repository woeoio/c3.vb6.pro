VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Holder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ai/probes/udt_object_field - this class exists only to make "who still holds this object"
' visible: every instance prints one INIT line (from ProbeHost.NewH) and one TERM line here.
' The reading C3 account #600 needs is the ORDER of those lines against the numbered
' markers, not their content.

Public Tag As Long

Private Sub Class_Terminate()
    ProbeHost.gTerm = ProbeHost.gTerm + 1
    ProbeHost.Say "TERM " & Me.Tag
End Sub
