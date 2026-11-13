Attribute VB_Name = "Tools_Clone"
Option Explicit

Private m_ModifiedRectF As RectF

' 同名函数，形参表与 Tools_Pencil 那家**不同**（只有一枚 Optional）。
Public Function GetModifiedUpdateRectF(Optional ByVal dummy As Boolean = True) As RectF
    GetModifiedUpdateRectF = m_ModifiedRectF
End Function
