Option Explicit

' 账 §39: 定长串内嵌 NUL 的取子串 + 0 参 Variant 函数返回数组的下标形态。
' 两处都是 ClsTreeExplorer.cls 的真实写法 (GetLogicalDriveStrings + GetDrive(i)):
' 旧实现 Left$/Mid$/Right$ 按 wcslen 截断 (NUL 后的内容丢失), 且 GetDrive(i) 的 (i)
' 被当成调用实参按「超过形参个数」截掉 (整个数组原样返回, 探针实测 VarType=8200)。
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function GetDrive() As Variant
    Dim lngRet As Long
    Dim strDrives As String * 255
    lngRet = GetLogicalDriveStrings(Len(strDrives), strDrives)
    GetDrive = Split(Left$(strDrives, lngRet - 1), Chr$(0))
End Function

Public Property Get GetDrivesCount() As Long
    GetDrivesCount = UBound(GetDrive)
End Property

Public Sub Main()
    Dim v As Variant
    Dim strDrives As String * 255
    Dim lngRet As Long

    lngRet = GetLogicalDriveStrings(Len(strDrives), strDrives)
    ' 内嵌 NUL 在 (C:\ + Chr$(0) + ...): InStr 必须能看见它
    Debug.Print "SND-nulcmp="; (InStr(strDrives, Chr$(0)) = 4)
    ' 取子串按字符数走, 不许在 NUL 处截断
    Debug.Print "SND-leftlen="; (Len(Left$(strDrives, lngRet - 1)) = lngRet - 1)
    Debug.Print "SND-midlen="; (Len(Mid$(strDrives, 4, 3)) = 3)
    ' GetDrive(i) 是对返回数组取下标, 不是调用实参
    Debug.Print "SND-hasdrive="; (GetDrivesCount >= 1)
    v = GetDrive(0)
    Debug.Print "SND-drive0="; (Right$(v, 2) = ":\")
    Debug.Print "SND-drivevar="; (VarType(v) = vbString)
    Debug.Print "SND-DONE"
End Sub
