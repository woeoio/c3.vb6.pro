VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B171Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 账 #305 §11.153 (第一百七十一刀): Currency 属性的读在拼串里按 Long 格式化的病。
' 这一刀只修「拼串选哪枚格式化」—— 承载仍是 double (Fix 126 的老账, 在册)。

Private m_C As Currency
Private m_D As Double

Public Property Get C1() As Currency
    C1 = m_C
End Property

Public Property Let C1(ByVal v As Currency)
    m_C = v
End Property

Public Property Get D1() As Double
    D1 = m_D
End Property

Public Property Let D1(ByVal v As Double)
    m_D = v
End Property
