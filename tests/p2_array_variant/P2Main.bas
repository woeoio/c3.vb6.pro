Attribute VB_Name = "P2Main"
Private Declare Sub PAny Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal cb As Long)

Public Function ArrV() As Variant
    ArrV = Array(1, 2, 3)
End Function

Public Sub Main()
    Dim v As Variant
    v = Array(1, 2, 3)
    ' P2 A8: As Any ByRef 收到的必须是整只槽 (Variant 装数组本体时亦然)
    Call PAny(ArrV(), v, 16)
    ' P2 A1: Variant 类型 (数组本体) 参与算术 -> VarAddV + SafeArray 臂装箱
    Debug.Print Array(1, 2, 3) + 1
    ' P2 A10/A11: 类型是 Variant, 但 C 是载体时先装箱再进收 VARIANT 的出口
    Debug.Print CStr(ArrV())
    Debug.Print CStr(v)
    Debug.Print CLng(ArrV())
    ' P2 C1/file_io: Print# 的 Variant 档必须先装箱 (改前是 FileVariantText(载体) = C2440)
    Open "p2_array_variant.txt" For Output As #1
    Print #1, ArrV()
    Print #1, Array(1, 2, 3)
    Close #1
End Sub
