VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CB448"
Option Explicit
Implements IB448
Private mN As Long

Public Function Twice(ByVal n As Long) As Long Implements IB448.Twice
    mN = mN + 1
    Twice = n * 2 + 1
End Function

Public Function Hits() As Long
    Hits = mN
End Function
