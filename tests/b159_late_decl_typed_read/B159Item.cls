VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B159Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ID As Long
Private m_Terms As Long

Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Let ID(ByVal x As Long)
    m_ID = x
End Property

Private Sub Class_Terminate()
    gTerms = gTerms + 1
End Sub
