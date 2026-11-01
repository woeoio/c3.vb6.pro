VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "H3100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IK3100

Private m_hits As Long

Private Sub Class_Initialize()
    m_hits = 7
End Sub

Private Sub Class_Terminate()
    mdl3100.g_term = mdl3100.g_term + 1
End Sub

Private Sub IK3100_Ping()
    m_hits = m_hits + 1
End Sub

Public Property Get Hits() As Long
    Hits = m_hits
End Property
