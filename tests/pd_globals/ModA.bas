Attribute VB_Name = "ModA"
Option Explicit

Public Type ColorRec
    r As Long
    g As Long
    b As Long
End Type

Public gX As Long
Public gArr() As Long
Public gColor As ColorRec
Private pvLocal As Long
Friend fFr As Long
