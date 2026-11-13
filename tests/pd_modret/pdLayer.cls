VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_dib As pdDIB

Friend Sub SetLayerOffsetX(ByVal v As Long)
    Debug.Print v
End Sub

Friend Function AffineTransformsActive(ByVal b As Boolean) As Boolean
    AffineTransformsActive = b
End Function

Friend Property Get GetLayerDIB() As pdDIB
    Set GetLayerDIB = m_dib
End Property
