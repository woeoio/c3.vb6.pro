Attribute VB_Name = "PdChainUse"
Option Explicit

Private m_Img As PdChainImg

' PhotoDemon Layers.bas:467 的形状: 类字段链尾上「无括号零参方法（返回 UDT）」再读字段。
Public Function ChainLeft() As Long
    ChainLeft = m_Img.Selection.GetBounds.Left
End Function

Public Function ChainTop() As Long
    ChainTop = m_Img.Selection.GetBounds.Top
End Function

' 对照: 带括号那一形本来就走调用结果，形状不变。
Public Function ChainLeftParens() As Long
    ChainLeftParens = m_Img.Selection.GetBounds().Left
End Function
