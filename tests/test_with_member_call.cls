VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True'
END
Attribute VB_Name = "WithMemberHost"
Option Explicit

' Fixture for the type-authority check on "With <a Function of this same class>".
' The bare name is neither a variable nor a field, so no by-name table can answer
' what it holds; the only owner of that question is the class's own member list.
' The two shapes below are the SAME VB sentence written inside and outside a With
' block, so they have to emit the SAME direct class-member call. ASCII on purpose:
' .cls sources are read as ANSI by the driver.

Public Left1 As Double
Public Right1 As Double

Public Function Clone() As WithMemberHost
    Set Clone = Me
End Function

Public Function ScaleFactor(ByVal k As Double) As Double
    ScaleFactor = k * 2#
End Function

Public Sub UseBothForms()
    Dim q1 As Double
    Dim q2 As Double
    With Clone
        q1 = .ScaleFactor(3#)
    End With
    q2 = Clone.ScaleFactor(3#)
    Debug.Print q1, q2
End Sub
