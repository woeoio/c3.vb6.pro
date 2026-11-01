VERSION 5.00
Begin VB.UserControl dhWithExt 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
End
Attribute VB_Name = "dhWithExt"
Option Explicit
' ledger 305 knife 261 (a): inside `With UserControl`, the members whose answer lives in the
' host pseudo channel must be routed there by the emitter. Before this knife the body marked
' every member as a late-bound hop onto the host HWND, so
'   .Extender          -> vb6_ComGetObjectProp(_vb6_with_N, L"Extender")   (no answer -> NULL)
'   .ScaleX / .ScaleY  -> vb6_ComCall(_vb6_with_N, L"ScaleY", ...)         (no answer -> 0)
' and the second hop then packed the *bare object* instead of the member read.
' The ComCtlsDemo reading for this shape was 18 + 5 + 2 host misses and 5 NULL receivers.

Public Sub ExtenderMove()
    With UserControl
        .Extender.Move .Extender.Left, .Extender.Top, .Extender.Width, .ScaleY(50, 3, 10)
        .Extender.Height = .ScaleY(60, 3, 10)
    End With
End Sub

Public Sub ScaleBoth()
    Dim n As Long
    With UserControl
        n = .ScaleX(10, 1, 3)
        n = .ScaleY(10, 1, 3)
    End With
End Sub
