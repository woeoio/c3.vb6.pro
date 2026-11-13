VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Items() As String
Private m_Count As Long

Public Function GetString(ByVal strIndex As Long) As String
    If strIndex < 0 Or strIndex >= m_Count Then Exit Function
    GetString = m_Items(strIndex)
End Function
