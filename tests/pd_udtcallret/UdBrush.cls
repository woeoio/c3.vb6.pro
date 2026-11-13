VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_RGBA As RGBQuad

Friend Function GetBrushColorRGBA() As RGBQuad
    GetBrushColorRGBA = m_RGBA
End Function
