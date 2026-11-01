VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SatOther"
Option Explicit

' 账 #481 判据夹具: the FINAL object's class. It differs from the churned class on purpose --
' a stale table entry left by a dropped SatImpl would answer "SatImpl" here (or nothing),
' while a correct table answers "SatOther".
Implements SatIface

Public Value As Long

Private Function SatIface_Ping(ByVal n As Long) As Long
    SatIface_Ping = n
End Function
