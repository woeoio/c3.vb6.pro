VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Friend Function GetLayerByID(ByVal requestedID As Long) As pdLayer
    Set GetLayerByID = New pdLayer
End Function

Friend Function GetNumOfLayers() As Long
    GetNumOfLayers = 1
End Function

Friend Sub GetCompositedImage(ByRef dstDIB As pdDIB, Optional ByVal useFinal As Boolean = True)
    Set dstDIB = New pdDIB
End Sub
