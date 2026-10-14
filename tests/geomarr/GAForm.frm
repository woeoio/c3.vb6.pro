VERSION 5.00
Begin VB.Form GAForm 
   Caption         =   "GeomArr"
   ClientHeight    =   3400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6800
   ScaleHeight     =   3400
   ScaleMode       =   1  
   ScaleWidth      =   6800
   StartUpPosition =   3  
   Begin VB.Timer tmrGA 
      Interval        =   150
      Left            =   120
      Top             =   1560
   End
   Begin VB.TextBox arrA 
      Height          =   247
      Index           =   0
      Left            =   1007
      TabIndex        =   0
      Top             =   449
      Width           =   3001
   End
   Begin VB.TextBox arrA 
      Height          =   247
      Index           =   1
      Left            =   1211
      TabIndex        =   1
      Top             =   761
      Width           =   1511
   End
   Begin VB.TextBox soloT 
      Height          =   247
      Left            =   1007
      TabIndex        =   2
      Top             =   1069
      Width           =   3001
   End
End
Attribute VB_Name = "GAForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Ledger 230 (the VB-side geometry cache) verified three write roads -- the four
' property setters, vb6_ControlMove and vb6_CreateControl -- but every judge it
' added reads a SINGLELY-NAMED control. Nothing ever asked an ARRAY ELEMENT, and
' the element road is a different host expression in the emitted code
' (vb6_CtrlArr_GetAt(&vb6_arr_<name>, i)), while the storage itself is keyed by
' (window, property). The corpus writes geometry into array elements whose design
' numbers are not a multiple of 15, so a route that quietly fell back to
' projecting the window would answer one grid-step off on every read -- and before
' this fixture the whole suite would still have been green.
'   GA01/GA02   both design-time elements read back their own .frm numbers, each
'               with a pixel witness (user32 rect + kernel32 MulDiv, never the
'               product's own conversion) so "stored" cannot mean "did not place"
'   GA03        a runtime write to an element reads back the requested number
'   GA04        Move on an element reads back all four and lands the window
'   GA05        the two elements are two windows: moving one leaves the other
'   GA05c       the two handles compared straight (no CLng round-trip) answer like
'               numbers -- a handle is a pointer value, and VB6 hands hWnd out as Long
'   GA05d       the same type answer on the two stringify roads: a handle prints its
'               digits, and CStr of a DC-shaped read stays on the Variant road
'   GA05e       the same type answer on the boxing road: a handle read into a Variant
'               slot holds a number (not a fake object), a Variant-held address compares
'               numerically, and a real Set of a control still answers VarType 9
'   GA06        the array never swallows its single-named sibling's write

Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECTAPI) As Long
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hwnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Type RECTAPI
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private dpiX As Long
Private dpiY As Long

Private Function TF(ByVal b As Boolean) As String
    If b Then TF = "True" Else TF = "False"
End Function

Private Sub ReadDpi()
    Dim d As LongPtr
    d = GetDC(0)
    dpiX = GetDeviceCaps(d, 88)   ' LOGPIXELSX
    dpiY = GetDeviceCaps(d, 90)   ' LOGPIXELSY
    ReleaseDC 0, d
End Sub

Private Function ToPx(ByVal twips As Long) As Long
    ToPx = MulDiv(twips, dpiX, 1440)
End Function

Private Function ToPy(ByVal twips As Long) As Long
    ToPy = MulDiv(twips, dpiY, 1440)
End Function

Private Function PxLeft(ByVal h As LongPtr) As Long
    Dim rc As RECTAPI
    Dim pt As POINTAPI
    GetWindowRect h, rc
    pt.x = rc.Left
    pt.y = rc.Top
    ScreenToClient GetParent(h), pt
    PxLeft = pt.x
End Function

Private Function PxTop(ByVal h As LongPtr) As Long
    Dim rc As RECTAPI
    Dim pt As POINTAPI
    GetWindowRect h, rc
    pt.x = rc.Left
    pt.y = rc.Top
    ScreenToClient GetParent(h), pt
    PxTop = pt.y
End Function

Private Function PxW(ByVal h As LongPtr) As Long
    Dim rc As RECTAPI
    GetWindowRect h, rc
    PxW = rc.Right - rc.Left
End Function

Private Function PxH(ByVal h As LongPtr) As Long
    Dim rc As RECTAPI
    GetWindowRect h, rc
    PxH = rc.Bottom - rc.Top
End Function

Private Sub tmrGA_Timer()
    tmrGA.Enabled = False
    ReadDpi

    Dim ok1 As Boolean, ok2 As Boolean, ok3 As Boolean
    Dim ok4 As Boolean, ok5 As Boolean, ok6 As Boolean
    Dim ok5c As Boolean
    Dim h5a As Long, h5b As Long
    Dim t5 As String
    Dim v11 As Variant, v12 As Variant
    Dim pv12 As LongPtr
    Dim o8 As Object

    ok1 = (arrA(0).Left = 1007) And (arrA(0).Top = 449) And (arrA(0).Width = 3001) And (arrA(0).Height = 247) _
        And (PxLeft(arrA(0).hwnd) = ToPx(1007)) And (PxTop(arrA(0).hwnd) = ToPy(449)) _
        And (PxW(arrA(0).hwnd) = ToPx(3001)) And (PxH(arrA(0).hwnd) = ToPy(247))
    Debug.Print "GA01-e0-design=" & TF(ok1) & " raw=" & arrA(0).Left & "," & arrA(0).Top & "," _
        & arrA(0).Width & "," & arrA(0).Height

    ok2 = (arrA(1).Left = 1211) And (arrA(1).Top = 761) And (arrA(1).Width = 1511) And (arrA(1).Height = 247) _
        And (PxLeft(arrA(1).hwnd) = ToPx(1211)) And (PxW(arrA(1).hwnd) = ToPx(1511))
    Debug.Print "GA02-e1-design=" & TF(ok2) & " raw=" & arrA(1).Left & "," & arrA(1).Top & "," _
        & arrA(1).Width & "," & arrA(1).Height

    arrA(1).Left = 5000
    arrA(1).Top = 444
    arrA(1).Width = 7777
    arrA(1).Height = 3001
    ok3 = (arrA(1).Left = 5000) And (arrA(1).Top = 444) And (arrA(1).Width = 7777) And (arrA(1).Height = 3001) _
        And (PxLeft(arrA(1).hwnd) = ToPx(5000)) And (PxW(arrA(1).hwnd) = ToPx(7777)) _
        And (PxTop(arrA(1).hwnd) = ToPy(444)) And (PxH(arrA(1).hwnd) = ToPy(3001))
    Debug.Print "GA03-e1-write=" & TF(ok3) & " raw=" & arrA(1).Left & "," & arrA(1).Top & "," _
        & arrA(1).Width & "," & arrA(1).Height

    arrA(1).Move 1231, 451, 3007, 247
    ok4 = (arrA(1).Left = 1231) And (arrA(1).Top = 451) And (arrA(1).Width = 3007) And (arrA(1).Height = 247) _
        And (PxLeft(arrA(1).hwnd) = ToPx(1231)) And (PxW(arrA(1).hwnd) = ToPx(3007)) _
        And (PxTop(arrA(1).hwnd) = ToPy(451))
    Debug.Print "GA04-e1-move=" & TF(ok4) & " raw=" & arrA(1).Left & "," & arrA(1).Top & "," _
        & arrA(1).Width & "," & arrA(1).Height

    ' GA05: element 0 kept its own numbers while element 1 was written twice --
    ' one cache per WINDOW, not one per array. The witness is the untouched window.
    ok5 = (arrA(0).Left = 1007) And (arrA(0).Width = 3001) _
        And (PxLeft(arrA(0).hwnd) = ToPx(1007)) And (PxW(arrA(0).hwnd) = ToPx(3001))
    Debug.Print "GA05-e0-untouched=" & TF(ok5) & " raw=" & arrA(0).Left & "," & arrA(0).Width & "," _
        & PxLeft(arrA(0).hwnd) & "/" & ToPx(1007)
    h5a = CLng(arrA(0).hwnd)
    h5b = CLng(arrA(1).hwnd)
    Debug.Print "GA05b-two-windows=" & TF(h5a <> h5b) & " raw=" & h5a & "/" & h5b & _
        " idx=" & arrA(0).Index & "/" & arrA(1).Index

    ' GA05c: the same pair asked straight, without the CLng round-trip. A handle is a
    ' pointer value and VB6 hands hWnd out as a Long, so two different windows must
    ' answer "not equal" both on its own and inside an And chain. eqself is the
    ' opposite-side witness -- a road that simply answered True would pass the first
    ' two and fail this one.
    ok5c = (arrA(0).hwnd <> arrA(1).hwnd) And (arrA(0).hwnd = arrA(0).hwnd)
    Debug.Print "GA05c-inline=" & TF(arrA(0).hwnd <> arrA(1).hwnd) & " chain=" & TF(ok5c) _
        & " eqself=" & TF(arrA(0).hwnd = arrA(0).hwnd)

    ' GA05d: the same type answer feeds the two stringify roads. Before a handle was
    ' registered as one, "x" & ctrl.hwnd printed only "x" (the pointer went into the
    ' object box and came back empty) and CStr of a DC-looking read had to stay on the
    ' Variant road -- cast that one to a number and the fixture does not even compile.
    t5 = "H" & arrA(0).hwnd
    Debug.Print "GA05d-digits=" & TF(Len(t5) > 1) & " hdcstr=" & TF(Len(CStr(soloT.hDC)) > 0)

    ' GA05e: the boxing road. A handle read assigned into a Variant slot used to reach
    ' _Generic's default arm = vb6_VariantObject, so an HWND sat in a VT_DISPATCH slot:
    ' VarType answered 9, CStr answered empty, and the slot compared unequal to the very
    ' number it holds (measured on both compilers). It must answer like the number VB6
    ' hands out -- and a real object reference beside it must stay an object, or this arm
    ' would be turning every Set into a digit.
    v11 = arrA(0).hwnd
    Debug.Print "GA05e-notobj=" & TF(VarType(v11) <> 9) & " str=" & TF(Len(CStr(v11)) > 1) _
        & " num=" & TF(CLng(v11) = h5a) & " eq=" & TF(v11 = h5a) _
        & " ne-other=" & TF(Not (v11 = h5b))
    ' The converter every Variant comparison reads through had no VT_I8 arm, so a 64-bit
    ' numeric box collapsed to 0.0 there -- reachable without any handle at all, through
    ' the address a VarPtr box carries (x64 read: VarPtr(x) <> 0 answered False).
    pv12 = VarPtr(h5a)
    v12 = pv12
    Set o8 = soloT
    Debug.Print "GA05e-vp-nz=" & TF(v12 <> 0) & " obj=" & TF(VarType(o8) = 9)

    ' GA06: the single-named sibling answers its own write too, so the array road
    ' did not take over the named-control road next to it.
    soloT.Left = 2227
    ok6 = (soloT.Left = 2227) And (PxLeft(soloT.hwnd) = ToPx(2227)) _
        And (arrA(0).Left = 1007)
    Debug.Print "GA06-solo-write=" & TF(ok6) & " raw=" & soloT.Left & "," & arrA(0).Left

    Debug.Print "GA-DONE"
    Unload Me
End Sub
