Attribute VB_Name = "ArrVar"
Option Explicit
' <vbeclipse> C2065 族① 判据夹子: 与 PhotoDemon 的
' `UserControls.ThemeFlyoutControls cmdFlyoutLock` 同形 —— 形参是 `ByRef ... As Variant`,
' 实参是**整枚控件数组** (13 个 Toolpanel 各一处, 整编 13 条 C2065)。
' 正文故意不碰形参: 本题只钉调用点的发码形状。
Public Sub ThemeFlyoutControls(ByRef cmdFlyoutLock As Variant)
    Dim n As Long
    n = 0
End Sub
