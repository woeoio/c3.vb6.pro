VERSION 5.00
Begin VB.Form frmWith 
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   3000
   ScaleMode       =   1  'Twips
   ScaleWidth      =   6000
End
Attribute VB_Name = "frmWith"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> fixture (comwith): account 260 -- inside `With <Object>`, a right-hand side
' that is EXACTLY one member read used to lose that read: the emitter packed the whole
' object into the property slot (`.CompareMode = d1.CompareMode` became
' `vb6_ComSetProp(w, L"CompareMode", vb6_ComPackValue(d1))`). Same statement written
' without With was (and stays) correct, so this is the "same decision copied twice,
' third copy never asked the one authority" shape.
' All heads use Scripting.Dictionary (late-bound, no type library) so the judgement does
' not depend on a machine's registry or on a relative Reference= path. CompareMode is the
' writable numeric property; it is set before any key is added, which is what Dictionary
' requires.
Private Sub Form_Load()
    Dim d1 As Object, d2 As Object, d3 As Object, d4 As Object
    Dim d5 As Object, d6 As Object, w As Object
    Set d1 = CreateObject("Scripting.Dictionary")
    Set d2 = CreateObject("Scripting.Dictionary")
    Set d3 = CreateObject("Scripting.Dictionary")
    Set d4 = CreateObject("Scripting.Dictionary")
    Set d5 = CreateObject("Scripting.Dictionary")
    Set d6 = CreateObject("Scripting.Dictionary")
    Set w = CreateObject("Scripting.Dictionary")
    d1.CompareMode = 1
    On Error Resume Next
    ' head 1: field/late-bound receiver, RHS is one bare member read
    Err.Clear
    With d2
        .CompareMode = d1.CompareMode
    End With
    Dim e1 As Long, c1 As Long
    e1 = Err.Number
    c1 = d2.CompareMode
    Debug.Print "CW-WITH-RAW err=" & CStr(e1) & " cm=" & CStr(c1)
    Debug.Print "CW-WITH=" & CStr(e1 = 0 And c1 = 1)
    ' head 2: same shape with a LOCAL variable as the With receiver
    Err.Clear
    With w
        .CompareMode = d1.CompareMode
    End With
    Dim e2 As Long, c2 As Long
    e2 = Err.Number
    c2 = w.CompareMode
    Debug.Print "CW-LOCAL-RAW err=" & CStr(e2) & " cm=" & CStr(c2)
    Debug.Print "CW-LOCAL=" & CStr(e2 = 0 And c2 = 1)
    ' head 3: same shape with a PARAMETER as the With receiver
    Err.Clear
    CopyInto d3, d1
    Dim e3 As Long, c3 As Long
    e3 = Err.Number
    c3 = d3.CompareMode
    Debug.Print "CW-PARAM-RAW err=" & CStr(e3) & " cm=" & CStr(c3)
    Debug.Print "CW-PARAM=" & CStr(e3 = 0 And c3 = 1)
    ' head 4: arithmetic RHS -- already correct before the fix, must not double-resolve
    Err.Clear
    With d6
        .CompareMode = d1.CompareMode + 0
    End With
    Dim e4 As Long, c4 As Long
    e4 = Err.Number
    c4 = d6.CompareMode
    Debug.Print "CW-EXPR-RAW err=" & CStr(e4) & " cm=" & CStr(c4)
    Debug.Print "CW-EXPR=" & CStr(e4 = 0 And c4 = 1)
    ' counter-witness A: the identical statement WITHOUT With (was already right)
    Err.Clear
    d4.CompareMode = d1.CompareMode
    Dim e5 As Long, c5 As Long
    e5 = Err.Number
    c5 = d4.CompareMode
    Debug.Print "CW-DIRECT-RAW err=" & CStr(e5) & " cm=" & CStr(c5)
    Debug.Print "CW-DIRECT=" & CStr(e5 = 0 And c5 = 1)
    ' counter-witness B: a literal RHS inside a With (untouched path, must stay right)
    Err.Clear
    With d5
        .CompareMode = 0
    End With
    Dim e6 As Long, c6 As Long
    e6 = Err.Number
    c6 = d5.CompareMode
    Debug.Print "CW-LIT-RAW err=" & CStr(e6) & " cm=" & CStr(c6)
    Debug.Print "CW-LIT=" & CStr(e6 = 0 And c6 = 0)
    Debug.Print "CW-DONE"
    Unload Me
End Sub

Private Sub CopyInto(ByVal target As Object, ByVal source As Object)
    On Error Resume Next
    With target
        .CompareMode = source.CompareMode
    End With
End Sub
