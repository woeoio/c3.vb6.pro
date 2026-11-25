VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B156Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHeld As Long
Private mNote As String

Public Property Get Held() As Long
    Held = mHeld
End Property

Public Property Let Held(ByVal v As Long)
    mHeld = v
End Property

Public Property Get Note() As String
    Note = mNote
End Property

Public Property Let Note(ByVal v As String)
    mNote = v
End Property
