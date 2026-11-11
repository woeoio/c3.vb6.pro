VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 注册在前的同名 GetString (3 参) —— 裸键表借到它就会多补缺省实参。
Public Function GetString(ByVal paramName As Long, _
                          Optional ByVal defaultReturn As Long = 0, _
                          Optional ByVal nameSafe As Boolean = False) As String
    GetString = vbNullString
End Function
