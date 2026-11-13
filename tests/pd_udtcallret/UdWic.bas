Attribute VB_Name = "pdWic"
Option Explicit

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function FakeCreate Lib "ole32.dll" (ByVal pFactory As Long, ByVal pName As Long, ByRef pGuid As Any, ByVal access As Long) As Long

' 同模块私有零参函数返回 UDT —— 实参位裸名 (IdentifierExpr) 形状。
Private Function GUID_NULL() As Guid
    GUID_NULL.Data1 = 0
End Function

Public Sub TestGuid()
    Dim hr As Long
    hr = FakeCreate(0, 0, GUID_NULL, 1)
    Debug.Print CStr(hr)
End Sub
