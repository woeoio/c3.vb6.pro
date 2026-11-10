Option Explicit

' 定长字符串常量（PhotoDemon PublicConstants.bas:28 的真实形态）：
'   Public Const MACRO_EXT As String * 3 = "pdm"
' 修前 parseTypeRef 里 String * N 的长度用 parseExpression() 全优先级，
' 把 `3 = "pdm"` 当相等比较吃掉 ⇒ "expected '=' in Const declaration (got NewLine)"。
' 真 VB6：Len=3、值为 "pdm"。

Public Const MACRO_EXT As String * 3 = "pdm"

Public Sub Main()
    Debug.Print "FSC=" & Len(MACRO_EXT) & "," & MACRO_EXT
End Sub
