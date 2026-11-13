Attribute VB_Name = "ModB"
Option Explicit

Public Sub UseGlobals()
    gX = 41
    gArr(0) = gX + 1
    gColor.r = gArr(0)
    Debug.Print gColor.r
End Sub

Public Sub BadRefs()
    gXX = 1
    pvLocal = 2
    fFr = 3
End Sub
