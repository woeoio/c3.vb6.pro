Attribute VB_Name = "DialogManager"
Option Explicit

Public Sub Prompt(ByVal fi_Handle As Long)
    Load dialog_ToneMapping
    dialog_ToneMapping.fi_HandleCopy = fi_Handle
    dialog_ToneMapping.ShowDialog
End Sub
