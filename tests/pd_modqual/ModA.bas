Attribute VB_Name = "ModA"
Option Explicit

' C2197 钉子（合格调用）：这一枚**无参**，下面 ModB 里有一枚同名但带一枚必填参的过程。
Public Function GetLevel() As Long
    GetLevel = 42
End Function
