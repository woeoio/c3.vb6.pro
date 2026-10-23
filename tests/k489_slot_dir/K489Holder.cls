VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K489Holder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "k489 target: same-name property, Let declared FIRST"
Option Explicit
Implements IK489

Private m_label As String
Private m_puts As Long

Private Property Let IK489_Label(ByVal v As String)
    m_label = v
    m_puts = m_puts + 1
End Property

Private Property Get IK489_Label() As String
    IK489_Label = "H" & m_label & "/" & m_puts
End Property

Private Property Get IK489_Num() As Long
    IK489_Num = 489
End Property

Public Function Puts() As Long
    Puts = m_puts
End Function
