Attribute VB_Name = "UFMain"
Option Explicit

' ai/032 §6.9: 类模块字段位的四档类型探针 (见 tu_holder.cls)。
Sub UFMain()
    Dim h As TUHolder
    Set h = New TUHolder
    h.mU = &HFFFFFFFF
    h.mUI = &HFFFF
    h.mU3 = &H123456789A
    h.mSB = -128
    h.Show
    Debug.Print "UF-DONE"
End Sub
