Attribute VB_Name = "PdPropLet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' pdResize.ctl:512/525/1103/1104 的同形最小夹具:
' `Property Let ResizeWidth(newWidth As Double)` 没写 ByVal ⇒ VB6 默认 ByRef,
' C 形参是 `double*`; 而 `Me.ResizeWidth = .GetDouble(...)` 的右值是值。
' 旧发码直接递值 ⇒ cl C2440「cannot convert from 'double' to 'double *'」。
Private m_w As Double

Public Property Let ResizeWidth(newWidth As Double)
    m_w = newWidth
End Property

Public Property Get ResizeWidth() As Double
    ResizeWidth = m_w
End Property

Public Sub ApplyFromBag(ByRef src As Object)
    With src
        Me.ResizeWidth = .GetDouble("width", 1920)
    End With
End Sub

Public Sub ApplyLiteral()
    Me.ResizeWidth = 42.5
End Sub
