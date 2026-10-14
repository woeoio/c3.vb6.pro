VERSION 5.00
Begin VB.Form SbHmForm 
   Caption         =   "SbHmForm"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7000
   LinkTopic       = "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   7000
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2085
      Width           =   7000
      _ExtentX        =   12348
      _ExtentY        =   556
      Style           =   0
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   0
            Object.Width           =   1500
            MinWidth        =   26
            Text            =   "A"
            Key             =   "fx1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   0
            Object.Width           =   2646
            MinWidth        =   26
            Text            =   "BB"
            Key             =   "fx2"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            MinWidth        =   26
            Text            =   "SP"
            Key             =   "sp"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SbHmForm"
Option Explicit

' zhang 299 (B130) spelling 3: a panel kept in a MODULE-LEVEL object variable.
Private mPo As Object

' zhang #206 (B41) measure surface. Written in the shape the **real VB6 designer** saves:
' `BeginProperty Panels {GUID}` + `BeginProperty PanelN {GUID}` + the `Object.Width` line.
' Measured on 5623 third-party .frm files: those panel numbers tile the bar's `_ExtentX`
' (himetric, 0.01 mm), NOT its `Width` (twips) - and the property itself is twips in code.
' So the key name carries a prefix our ingest never looked for, and the value carries a
' unit our store never had. Before this knife the whole design value was dropped:
' fixed panels fell back to the text measure (HM01 False, RAW shows a single-digit width).
'
' What is pinned here:
'   HM01 the two fixed panels came out as himetric widths, not text widths   (the red face)
'   HM02 both went through ONE conversion (same factor, DPI-independent ratio)
'   HM03 the spring panel still absorbs the remainder (guard, True on both sides)
'   HM04 the panels themselves arrived at all                  (existence guard)
'   HM05-RAW only: Panels(i).Width is still the REQUESTED value -> B41's other face,
'        deliberately not changed by this knife (see the ledger).
Private Declare Function SbGetParts Lib "user32" Alias "SendMessageW" (ByVal hWin As LongPtr, _
    ByVal msg As Long, ByVal n As Long, ByRef parts As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWin As LongPtr, ByRef rc As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Function TF(b As Boolean) As String
    If b Then TF = "True" Else TF = "False"
End Function

Private Sub Form_Load()
    Dim pt(0 To 3) As Long
    Dim r As RECT
    Dim w1 As Long, w2 As Long, e3 As Long, cw As Long
    Dim ok1 As Boolean, ok2 As Boolean, ok3 As Boolean, ok4 As Boolean
    Dim ok5 As Boolean
    Dim ok6 As Boolean, ok7 As Boolean, ok8 As Boolean, ok9 As Boolean, ok10 As Boolean
    Dim po As Object
    Dim pv As Variant
    Dim withW As Long
    Dim tpp As Long
    Dim ok11 As Boolean, ok12 As Boolean, ok13 As Boolean
    Dim feCount As Long, feIdx As Long, feSum As Long, feWant As Long, feTxt As Long
    Dim ok14 As Boolean
    Dim withN As Long, withTxt As Long, withTxt1 As Long
    Dim ecount As Long

    SbGetParts StatusBar1.hWnd, 1030, 3, pt(0)
    w1 = pt(0)                       ' panel1 right edge == its width (it starts at 0)
    w2 = pt(1) - pt(0)               ' panel2 laid-out width
    e3 = pt(2)                       ' last edge == whole client width
    GetClientRect StatusBar1.hWnd, r
    cw = r.Right - r.Left

    Debug.Print "HM00-RAW w1=" & CStr(w1) & " w2=" & CStr(w2) & " e3=" & CStr(e3) & " cw=" & CStr(cw)

    ' HM01: 1500 hm is 57 px and 2646 hm is 100 px at 96 DPI; the texts "A"/"BB" measure
    ' far below these floors. So a pass here means the design width actually reached the RTL.
    ok1 = (w1 >= 40) And (w2 >= 80) And (w2 > w1)
    Debug.Print "HM01-HM-DESIGN-WIDTH=" & TF(ok1)

    ' HM02: one authority, so the two numbers keep the 2646:1500 ratio (integers aside).
    ok2 = Abs(w2 * 1500 - w1 * 2646) <= 3 * 2646
    Debug.Print "HM02-SINGLE-CONVERSION=" & TF(ok2)

    ' HM03/HM04: guards - they are True before this knife too, they only catch a bad fix.
    ok3 = (e3 = cw) And (cw > 0)
    Debug.Print "HM03-SPRING-TILES=" & TF(ok3)
    ok4 = (StatusBar1.Panels.Count = 3) And (StatusBar1.Panels(1).Key = "fx1") _
          And (StatusBar1.Panels(2).Text = "BB")
    Debug.Print "HM04-PANELS-ARRIVED=" & TF(ok4)

    ' RAW only: the requested-value caliber (B41's remaining face). Panel1 was given 1500 hm.
    Debug.Print "HM05-RAW cs=" & CStr(StatusBar1.Panels(1).Width) _
                & " bare=" & StatusBar1.Panels(2).Width

    ' HM06 = zhang 300 (§B131). Panels(<number>).Index used to be emitted as the *ByKey
    ' exit with an INTEGER in a wchar_t* slot -> the RTL dereferenced address 0x1 and the
    ' process died during Form_Load. Both spellings have to answer, and with the SAME
    ' number: pinning only one of them is exactly how this shipped silently (the key form
    ' was always fine, so SB14-IDXBYKEY in the neighbour fixture stayed green).
    ok5 = (StatusBar1.Panels(2).Index = 2) And (StatusBar1.Panels("fx2").Index = 2)
    Debug.Print "HM06-INDEX-BOTH-FORMS=" & TF(ok5)
    ' RAW only: an out-of-range subscript answers 0 here. VB6 raises error 9 at that point
    ' - that caliber is a separate account, so this line is deliberately not pinned to a number.
    Debug.Print "HM07-RAW idx9=" & StatusBar1.Panels(9).Index

    ' HM08 = zhang 299 (B130): `Panels` as a real collection object. Before this knife
    ' the collection was never built at all: every "use the member as an object" spelling
    ' emitted vb6_ComGetObjectProp(<HWND>, L"Panels") - a HWND used as an IDispatch, the
    ' same symptom C29-8b wrote down for TreeView Nodes - and `For Each` did not even
    ' enter its body (zero iterations, measured on the pre-knife compiler).
    ' Pinned face: the enumerator, three ways at once - iteration count, the sum of the
    ' Index each item reports, and the total text length ("A"+"BB"+"SP" = 5). Pinning
    ' only "the loop ran" would let a one-item or mis-subscripted collection through.
    ' Widths are still printed RAW (HM13..HM15 own that caliber now), but note the sum moved
    ' when the 22nd cut landed: spring panel went 0 -> its laid-out twip width.
    feCount = 0
    feIdx = 0
    feSum = 0
    feTxt = 0
    For Each pv In StatusBar1.Panels
        feCount = feCount + 1
        feIdx = feIdx + CLng(pv.Index)
        feSum = feSum + CLng(pv.Width)
        feTxt = feTxt + Len(pv.Text)
    Next
    ' want= is the SAME three reads taken through the early-bound route (the face this
    ' account opened with). Since the 21st cut it agrees with fe - printed, not judged.
    feWant = CLng(StatusBar1.Panels(1).Width)
    feWant = feWant + CLng(StatusBar1.Panels(2).Width)
    feWant = feWant + CLng(StatusBar1.Panels(3).Width)
    ok6 = (feCount = 3) And (feIdx = 6) And (feTxt = 5)
    Debug.Print "HM08-FOREACH-COLLECTION=" & TF(ok6)
    Debug.Print "HM08-RAW fe=" & CStr(feCount) & " idx=" & CStr(feIdx) & " txt=" & CStr(feTxt) _
                & " sum=" & CStr(feSum) & " want=" & CStr(feWant)

    ' HM09..HM12 = the second knife on this account (zhang 299, 21st cut). The map
    ' "which creator answers this member-collection name" used to be hand-copied into THREE
    ' emitters (value context / object context / the default-member indexer that Set-RHS and
    ' With-object-ref go through), and the StatusBar row was missing from all three - so
    ' those spellings kept handing back the HWND-as-IDispatch chain. Every judgment below is
    ' still a CROSS-READ against a witness that is NOT the RTL: w1/w2 come from SB_GETPARTS,
    ' key/index/text from the design block.
    tpp = Screen.TwipsPerPixelX       ' witness factor: pixels -> twips, the Screen route,
                                      ' NOT this repo's converter (so it stays a witness)
    Set po = StatusBar1.Panels(1)
    ok7 = (CStr(po.Width) = CStr(w1 * tpp)) And (po.Key = "fx1") And (CStr(po.Index) = "1")
    Debug.Print "HM09-OBJVAR-ROADS=" & TF(ok7)
    Set mPo = StatusBar1.Panels(2)
    ok8 = (CStr(mPo.Width) = CStr(w2 * tpp)) And (mPo.Text = "BB")
    Debug.Print "HM10-MODVAR-ROADS=" & TF(ok8)
    With StatusBar1.Panels(1)
        withW = .Width
    End With
    Set po = StatusBar1.Panels("fx2")
    ok9 = (withW = w1 * tpp) And (CStr(po.Index) = "2") And (CStr(po.Width) = CStr(w2 * tpp))
    Debug.Print "HM11-WITH-AND-KEYSUB=" & TF(ok9)

    ' HM12 = the face this account was opened with: the member read sitting in an ARGUMENT
    ' (CStr / CLng). It reads the requested width now through the same collection object, so
    ' the "argument re-resolves the receiver" half of the diagnosis turned out to be the very
    ' same missing table row - not a second defect.
    ok10 = (CStr(StatusBar1.Panels(1).Width) = CStr(w1 * tpp)) _
           And (CLng(StatusBar1.Panels(2).Width) = w2 * tpp)
    Debug.Print "HM12-ARG-ROADS=" & TF(ok10)

    ' HM13..HM15 = zhang 206 (B41) second face, 22nd cut. `Panels(i).Width` used to answer
    ' the STORED number; VB6 answers the laid-out geometry, in twips (corpus: 9 independent
    ' sites mix it with ScaleWidth / another control's Width, one divides by
    ' Screen.TwipsPerPixelX to feed Win32). Three things are pinned, each with a witness that
    ' is NOT the RTL's own store:
    '   HM13 the SPRING panel - it has no requested width at all, so it read 0 before this
    '        cut; now it must equal what the window says: client right (e3) minus the two
    '        fixed panels' right edges, converted by the Screen factor.
    '   HM14 Panel.Left - the other half of the same projection; repo-wide there was no exit
    '        for it at all, so the two corpus overlay-positioning sites read back empty.
    '   HM15 design-block MinWidth - VB6 persists it in himetric too (398 of 457 panels write
    '        MinWidth identical to Object.Width), so the number that comes back must NOT be
    '        the 26 the .frm carries.
    ok11 = (StatusBar1.Panels(3).Width > 0) _
           And (StatusBar1.Panels(3).Width = (e3 - w1 - w2) * tpp)
    Debug.Print "HM13-SPRING-WIDTH-FOLLOWS-WINDOW=" & TF(ok11)
    ok12 = (StatusBar1.Panels(1).Left = 0) _
           And (StatusBar1.Panels(2).Left = w1 * tpp) _
           And (StatusBar1.Panels(3).Left = (w1 + w2) * tpp)
    Debug.Print "HM14-PANEL-LEFT=" & TF(ok12)
    Debug.Print "HM15-RAW minw=" & StatusBar1.Panels(1).MinWidth & " frm=" & 26
    ok13 = (StatusBar1.Panels(1).MinWidth > 0) And (StatusBar1.Panels(1).MinWidth <> 26)
    Debug.Print "HM15-DESIGN-MINWIDTH-NOT-RAW=" & TF(ok13)

    ' HM16 = zhang 299 (B130), 25th cut: the FOURTH emit road -- the With header.
    ' The 21st cut put "which member-collection does this name create" in ONE place, but only
    ' three roads asked it (value / object / default-member indexer). The With header never
    ' consumed the pending COM marker at all. Measured on the pre-cut compiler, two things
    ' went wrong in one statement: the receiver stayed (void*)vb6_hwnd_StatusBar1 (a HWND used
    ' as an IDispatch answers nothing -> .Count read 0), and the leftover marker leaked into
    ' the FIRST body statement, where `withN = .Count` was emitted as
    ' vb6_ComSetProp(<hwnd>, L"Panels", ...) -- a READ turned into a property WRITE.
    ' Witnesses are not the RTL: ecount is the early-bound route, feCount the enumerator
    ' (HM08), and "BB"/"A" come from the design block. HM16-BAR-INTACT is the guard that the
    ' body did not damage the bar (the old mis-emit wrote to it).
    ecount = StatusBar1.Panels.Count
    withN = 0
    withTxt = 0
    With StatusBar1.Panels
        withN = .Count
        withTxt = Len(.Item(2).Text)
    End With
    withTxt1 = 0
    With StatusBar1.Panels
        withTxt1 = Len(.Item(1).Text)
    End With
    ok14 = (withN = 3) And (withTxt = 2) And (withTxt1 = 1) And (ecount = 3) And (feCount = 3)
    Debug.Print "HM16-WITH-COLLECTION=" & TF(ok14)
    Debug.Print "HM16-RAW with=" & CStr(withN) & " item2=" & CStr(withTxt) _
                & " item1=" & CStr(withTxt1) & " early=" & CStr(ecount) & " foreach=" & CStr(feCount)
    Debug.Print "HM16-BAR-INTACT=" & TF((StatusBar1.Panels.Count = 3) And (StatusBar1.Panels(2).Text = "BB"))

    Debug.Print "HM-DONE"
    Unload Me
End Sub
