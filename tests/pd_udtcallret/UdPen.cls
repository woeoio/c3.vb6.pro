VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Color As RGBQuad

Friend Sub SetPenColorRGBA(ByRef newColor As RGBQuad)
    m_Color = newColor
End Sub
