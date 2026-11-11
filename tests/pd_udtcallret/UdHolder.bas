Attribute VB_Name = "pdHolder"
Option Explicit

Private m_Img As pdImg

' 模块函数返回类实例 —— `pdHolder.GetImg.GetBoundaryRectF` 的最左一跳。
Public Function GetImg() As pdImg
    If m_Img Is Nothing Then Set m_Img = New pdImg
    Set GetImg = m_Img
End Function
