VERSION 5.00
Begin VB.Form FDForm 
   Caption         =   "FDraw"
   ClientHeight    =   3200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6400
   ScaleHeight     =   3200
   ScaleWidth      =   6400
   StartUpPosition =   3  
   Begin VB.Timer tmrF 
      Interval        =   150
      Left            =   120
      Top             =   1560
   End
End
Attribute VB_Name = "FDForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 233 = Form drawing state (DrawWidth / CurrentX / CurrentY / ScaleMode) round-trips.
' Before: the write side was never registered in the cgen write table, so
'   `Me.DrawWidth = 3` emitted `vb6_Form_DrawGetWidth(hwnd) = 3;` (C2106, both
'   arches, nothing compiled at all). And the store itself carried two
'   incompatible encodings on one window-property name: the draw family wrote
'   int32+1 while Form Print wrote a float bit pattern -- each read the other as
'   garbage. Pins here are two-sided on purpose:
'   FD01/FD02/FD03  write then read back the very numbers handed in
'   FD07/FD08       a pixel witness (draw a point, ask the same DC about it)
'   FD10            the merged store: after Print advanced the pen, the drawing
'                   side reads a sane number, not a float reinterpreted as int
' Everything runs on the first Timer tick -- a window question is asked while the
' window is alive (the dcsurf / pcline rule).
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

Private Function TF(ByVal b As Boolean) As String
    If b Then TF = "True" Else TF = "False"
End Function
' First device row in [lo,hi] that carries `want` somewhere in x in [xLo,xHi].
' Used by the two unit witnesses below; scanning a colour (not "any pixel that
' differs from the background") keeps a leftover line from answering for a new one.
Private Function FirstRowOfColor(d As LongPtr, want As Long, xLo As Long, xHi As Long, _
                                 lo As Long, hi As Long) As Long
    Dim row As Long, i As Long
    FirstRowOfColor = -1
    For row = lo To hi
        If FirstRowOfColor < 0 Then
            For i = xLo To xHi
                If GetPixel(d, i, row) = want Then FirstRowOfColor = row
            Next
        End If
    Next
End Function

Private Sub tmrF_Timer()
    Dim d As LongPtr
    Dim r As Long
    Dim px2 As Long
    Dim px3 As Long
    Dim ok1 As Boolean, ok2 As Boolean, ok3 As Boolean
    Dim ok5 As Boolean, ok10 As Boolean
    Dim ok6 As Boolean

    tmrF.Enabled = False

    Me.DrawWidth = 3
    ok1 = (Me.DrawWidth = 3)
    Me.CurrentX = 100
    Me.CurrentY = 50
    ok2 = (Me.CurrentX = 100) And (Me.CurrentY = 50)
    Me.PSet (200, 120)
    Me.PSet (300, 130)
    ok3 = (Me.CurrentX = 300) And (Me.CurrentY = 130)

    Debug.Print "FD01-drawwidth=" & TF(ok1)
    Debug.Print "FD02-curxy=" & TF(ok2)
    Debug.Print "FD03-pset2=" & TF(ok3)
    Debug.Print "FD04-RAW xy=" & CStr(Me.CurrentX) & "," & CStr(Me.CurrentY) & " dw=" & CStr(Me.DrawWidth)
    Debug.Print "FD05-sm0=" & CStr(Me.ScaleMode)

    Me.ScaleMode = vbPixels
    Debug.Print "FD06-sm1=" & CStr(Me.ScaleMode)
    Me.PSet (40, 60), vbRed
    d = GetDC(Me.hwnd)
    Debug.Print "FD07-PIXEL=" & TF(GetPixel(d, 40, 60) = vbRed)
    px2 = GetPixel(d, 41, 61)
    ok5 = (px2 <> vbRed)
    ' 235: pen color = one store only. Me.ForeColor must reach a PSet that carries no
    ' color argument; the point painted earlier with an explicit vbRed must stay red,
    ' so this cannot be won by 'repaint the whole surface blue'.
    Me.ForeColor = vbBlue
    Me.PSet (100, 120)
    px3 = GetPixel(d, 100, 120)
    ok6 = (px3 = vbBlue) And (GetPixel(d, 40, 60) = vbRed)
    Debug.Print "FD11-forecolor=" & TF(ok6)
    Debug.Print "FD12-RAW pen=" & CStr(px3) & " blue=" & CStr(vbBlue) & " first=" & CStr(GetPixel(d, 40, 60))
    r = ReleaseDC(Me.hwnd, d)
    Debug.Print "FD08-neg=" & TF(ok5) & " raw=" & CStr(px2)

    Me.CurrentX = 0
    Me.CurrentY = 30
    Print "AB"
    Debug.Print "FD09-AFTERPRINT=" & CStr(Me.CurrentX) & "," & CStr(Me.CurrentY)
    ok10 = (Me.CurrentX = 0) And (Me.CurrentY > 30) And (Me.CurrentY < 3000)
    Debug.Print "FD10-printstore=" & TF(ok10)
    ' ---- 237: Print moves the pen in this window's OWN unit, and paints where the
    ' pen says. Before, vb6_Form_Print lived outside the drawing family and answered
    ' all four questions itself: its own GetDC (never the dispatch-time one), no font
    ' selected, no colour, the user number handed straight to TextOutW as pixels, and
    ' a raw device-pixel line height stored back into the pen -- one Print in twips
    ' moved CurrentY by 16 while the same form answered TextHeight = 240.
    ' Every judge below is an equality between two paths, so no DPI constant is pinned.
    Dim curTw As Double
    Dim curPt As Double
    Dim sm0 As Long
    Dim sm1 As Long
    Dim thTw As Double
    Dim thPt As Double
    Dim rowB As Long
    Dim rowR As Long
    Dim ok13 As Boolean
    Dim ok14 As Boolean
    Dim ok16 As Boolean

    Me.ScaleMode = vbTwips
    Me.CurrentX = 0
    Me.CurrentY = 0
    Print "AB"
    curTw = Me.CurrentY
    sm0 = Me.ScaleMode
    thTw = Me.TextHeight("AB")
    ' 238: this judge is written as a plain equality ON PURPOSE. Until that account
    ' it had to be `Abs(Me.CurrentY - thTw) < 0.001`, because (Me.CurrentY = thTw)
    ' compiled to vb6_VarCmpEq(&boxed, &thTw) -- the Double local's address handed to
    ' a vb6_VARIANT* parameter -- and two equal numbers answered False.
    ok13 = (Me.CurrentY = thTw) And (Me.CurrentX = 0)
    Debug.Print "FD13-printadvance-twips=" & TF(ok13)
    Me.ScaleMode = vbPoints
    Me.CurrentX = 0
    Me.CurrentY = 0
    Print "AB"
    curPt = Me.CurrentY
    sm1 = Me.ScaleMode
    thPt = Me.TextHeight("AB")
    ok14 = (Me.CurrentY = thPt)
    Debug.Print "FD14-printadvance-points=" & TF(ok14)
    Debug.Print "FD15-RAW tw=" & CStr(thTw) & " pt=" & CStr(thPt) & _
              " curTw=" & CStr(curTw) & " curPt=" & CStr(curPt) & _
              " sm0=" & CStr(sm0) & " sm1=" & CStr(sm1)

    ' 72 points and 1 inch are the same physical distance, so the two lines must land
    ' on the same device row, an inch below the top. The long blue line stays visible
    ' to the right of the short red one, which is what the two x windows read.
    d = GetDC(Me.hwnd)
    Me.ForeColor = vbBlue
    Me.CurrentX = 0
    Me.CurrentY = 72
    Print "MMMMMMMMMMMMMMMM"
    Me.ScaleMode = vbInches
    Me.ForeColor = vbRed
    Me.CurrentX = 0
    Me.CurrentY = 1
    Print "M"
    rowB = FirstRowOfColor(d, vbBlue, 150, 200, 20, 240)
    rowR = FirstRowOfColor(d, vbRed, 0, 10, 20, 240)
    ok16 = (rowR > 50) And (rowR = rowB)
    Debug.Print "FD16-paint-units=" & TF(ok16)
    Debug.Print "FD17-RAW rowB=" & CStr(rowB) & " rowR=" & CStr(rowR)
    r = ReleaseDC(Me.hwnd, d)
    Debug.Print "FD-DONE"
    Unload Me
End Sub
