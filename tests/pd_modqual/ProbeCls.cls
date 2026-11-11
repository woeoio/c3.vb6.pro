VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProbeCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' C2197 第三族夹具（语句级无括号调用 + 括号分组首参 + 后续实参）:
'     m_Stream.SetPosition (a + b), FILE_BEGIN
' 修前表达式层把 `M (a + b)` 当"带 1 实参的调用", 语句层再把 `, FILE_BEGIN` 挂到
' **调用结果**上, 发成 `vb6_...(me, expr, 0, 0, _own_1, 1)` ⇒ C2197。
Private m_X As Long
Private m_Y As Long

Public Sub SetPos(ByVal pos As Long, Optional ByVal origin As Long = 0)
End Sub

Public Sub Run()
    Me.SetPos (m_X + m_Y), 1
End Sub
