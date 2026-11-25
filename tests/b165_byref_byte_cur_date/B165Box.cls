VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B165Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public B As Byte
Public C As Currency
Public T As Date

Public Sub SetB(ByRef x As Byte)
    x = 7
End Sub

Public Sub IncC(ByRef x As Currency)
    x = x + 1.25
End Sub

Public Sub NextD(ByRef x As Date)
    x = DateAdd("d", 1, x)
End Sub
