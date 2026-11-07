Option Explicit

' 回归: 隐式 String → 数值 实参 (ByVal 数值形参)
' VB6 允许 `Foo "123"` / 把 String 变量交给 ByVal Long 形参, 运行期按 CLng 解析。
' C3 改前把 BSTR (wchar_t*) 直接塞进 int32_t 形参 —— 指针低位被截断成数值,
' ucTreeView 的 `RGBtoARGB(N.BadgeBackColor, 100)` 因此把徽标底色画成噪声色
' (实测 (156,45,91) / (92,53,91)), 真 VB6 是 #51b453 / #ff6961。
' 改前读数: S2L-1=<指针低位> (实测 -911689592)

Private gs As String

Sub Main()
    Dim s As String
    Dim n As Long

    n = 5485649
    s = n
    Debug.Print "S2L-1=" & EchoL(s)
    Debug.Print "S2L-2=" & EchoI("321")
    Debug.Print "S2L-3=" & EchoD("12")

    gs = "51b453"
    Debug.Print "S2L-4=" & EchoL(Val("&H" & gs))
    Debug.Print "S2L-DONE"
End Sub

Private Function EchoL(ByVal v As Long) As Long
    EchoL = v
End Function

Private Function EchoI(ByVal v As Integer) As Integer
    EchoI = v
End Function

Private Function EchoD(ByVal v As Double) As Double
    EchoD = v
End Function
