VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B154Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Tag As Long

Private Sub Class_Initialize()
    gMade = gMade + 1
End Sub

Private Sub Class_Terminate()
    gTerms = gTerms + 1
End Sub
