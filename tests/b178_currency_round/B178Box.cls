Option Explicit

Private m_c As Currency

Public Property Get Cur() As Currency
    Cur = m_c
End Property

Public Property Let Cur(ByVal v As Currency)
    m_c = v
End Property
