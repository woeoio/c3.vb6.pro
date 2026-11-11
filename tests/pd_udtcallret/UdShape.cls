VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Brush As pdBrush
Private m_Pen As pdPen

' 类字段链上的零参方法 (无括号) 当 ByRef UDT 实参。
Friend Sub TestBrush()
    m_Pen.SetPenColorRGBA m_Brush.GetBrushColorRGBA
End Sub
