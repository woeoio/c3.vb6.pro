Attribute VB_Name = "Tools_Pencil"
Option Explicit

Private m_ModifiedRectF As RectF

' 与 Tools_Clone 同名、返回同一 UDT，但**形参表不同**（两个 Optional）。
' 裸键先到先得只留一份时，Tools_Clone.GetModifiedUpdateRectF 会借到这张表
' ⇒ 补 4 枚实参 (C2197 的另一面) 且返回 UDT 推断落空 ⇒ intptr_t 兜底 (C2440)。
Public Function GetModifiedUpdateRectF(Optional ByVal resetRectAfter As Boolean = True, Optional ByVal GetModifiedRectSinceStrokeBegan As Boolean = False) As RectF
    GetModifiedUpdateRectF = m_ModifiedRectF
End Function
