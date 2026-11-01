VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SatIface"
Option Explicit

' 账 #481 判据夹具: interface host for SatImpl (the view used by SatWorker.Churn).
Public Function Ping(ByVal n As Long) As Long
    Ping = n
End Function
