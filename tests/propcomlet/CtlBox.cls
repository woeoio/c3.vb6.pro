VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtlBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FontSize As Single

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    m_FontSize = newSize
End Property

Public Function ReadProperty(ByVal pName As String, ByVal pDefault As Variant) As Variant
    ReadProperty = pDefault
End Function

Public Sub SelfTest()
    Dim bag As Object
    With bag
        FontSize = .ReadProperty("FontSize", 10)
    End With
End Sub
