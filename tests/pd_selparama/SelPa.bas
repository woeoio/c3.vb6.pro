Attribute VB_Name = "SelPa"
Option Explicit
' <vbeclipse> C2065 族① 判据夹子: `Select Case` 的 Case 标签里调 ParamArray 过程。
' 与 PhotoDemon `Modules\Units.bas` 的
' `Select Case g_Language.TranslateMessage("px")` 同形 —— 标签求值要在条件里落一枚
' `SAFEARRAY* _pa_N = NULL;`。旧发码「边建条件边发 if 链」把这枚声明发进**上一条
' Case 的体内**，`} else if` 行看不到它 ⇒ 整编 12 条 C2065 (`_pa_19.._pa_30`)。
Public Function UnitId(ByVal s As String) As Long
    Select Case s
        Case "%", Tr("percent")
            UnitId = 0
        Case "px", Tr("px"), Tr("pixels")
            UnitId = 1
        Case Else
            UnitId = -1
    End Select
End Function

Private Function Tr(ByVal src As String, ParamArray extra() As Variant) As String
    Tr = src
End Function
