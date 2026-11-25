VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B158Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mV As Variant

Public Property Get V() As Variant
    V = mV
End Property

Public Property Let V(ByVal x As Variant)
    mV = x
End Property
