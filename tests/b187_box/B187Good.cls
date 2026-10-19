VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Good"
Option Explicit

' The clean class of this fixture: an in-project object that this project can box. The
' boxing leg hands out a borrowed wrapper for instances of THIS class, so Twice is the
' member both legs of every arm read through.

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
