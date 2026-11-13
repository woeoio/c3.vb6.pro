VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' §45 (C2065 族②): 类模块内 Private Type + 该 UDT 的动态数组。
' 形状取自 PhotoDemon Classes\pdStringHash.cls: `LenB(m_hashTable(i).sKey)`。
Private Type MergedEntry
    sKey As String
    sItem As String
    idxNext As Long
End Type

Private m_hashTable() As MergedEntry

Friend Sub Grow()
    ReDim m_hashTable(0 To 3) As MergedEntry
    If LenB(m_hashTable(0).sKey) = 0 Then
        m_hashTable(0).sKey = "a"
    End If
End Sub
