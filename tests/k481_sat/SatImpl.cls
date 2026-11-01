VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SatImpl"
Option Explicit

' 账 #481 判据夹具: the registered project class. Churn creates and drops N of these;
' with the unregister leg in place the table never saturates, so the instance made last
' is still resolvable by address (that is what the probe's CHURN_TYPE reads).
Implements SatIface

Public Value As Long

Private Function SatIface_Ping(ByVal n As Long) As Long
    SatIface_Ping = n
End Function
