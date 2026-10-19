Option Explicit

' Ledger 305 knife 89 (#408) witness class B. Same shape as A with a different field name,
' so a leaked A answer against a B instance is a wrong-member read, not just a wrong value.

Public Beta As String

Public Property Get Label() As String
    Label = "B-label"
End Property
