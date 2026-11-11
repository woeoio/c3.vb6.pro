Attribute VB_Name = "PdChainSel"
Option Explicit

Private m_Bounds As PdChainRect

Friend Function GetBounds() As PdChainRect
    GetBounds = m_Bounds
End Function
