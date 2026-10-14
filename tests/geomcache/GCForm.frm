VERSION 5.00
Begin VB.Form GCForm 
   Caption         =   "GeomCache"
   ClientHeight    =   3400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6800
   ScaleHeight     =   3400
   ScaleMode       =   1  
   ScaleWidth      =   6800
   StartUpPosition =   3  
   Begin VB.Timer tmrGC 
      Interval        =   150
      Left            =   120
      Top             =   1560
   End
   Begin VB.CheckBox chkA 
      Height          =   315
      Left            =   1007
      TabIndex        =   7
      Top             =   2800
      Width           =   1507
   End
   Begin VB.ComboBox cboA 
      Height          =   247
      Left            =   4607
      TabIndex        =   4
      Top             =   2213
      Width           =   1507
   End
   Begin VB.TextBox txtA 
      Height          =   247
      Left            =   1007
      TabIndex        =   0
      Top             =   449
      Width           =   3001
   End
   Begin VB.PictureBox picP 
      Height          =   1201
      Left            =   2000
      ScaleMode       =   1  
      TabIndex        =   1
      Top             =   1003
      Width           =   2003
      Begin VB.Label lblP 
         Height          =   211
         Left            =   311
         TabIndex        =   5
         Top             =   222
         Width           =   907
      End
   End
   Begin VB.PictureBox picX 
      Height          =   801
      Left            =   4201
      ScaleMode       =   3  
      TabIndex        =   2
      Top             =   1003
      Width           =   1401
      Begin VB.Label lblX 
         Height          =   21
         Left            =   101
         TabIndex        =   6
         Top             =   44
         Width           =   203
      End
   End
End
Attribute VB_Name = "GCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Ledger 230 = the four control geometry properties are VB-side VALUES, not a
' projection of the window rect. Reading GetWindowRect and converting back loses
' one step of the pixel grid: `txtA.Left = 5000` answered 4995, and the numbers
' the .frm designed are re-quantized the same way. A corpus scan over 105 design
' files measured 238 geometry values that are NOT a multiple of 15 (NewTab-test 86,
' ctrlslider 22, btnfocus 15, ... -- positions authored on a 120-DPI machine), so
' every one of those read back wrong before this fixture existed.
'   GC01/GC02   design numbers read back exactly (form route + container-child route)
'   GC03        placement witness: the window is still at the pixels the design
'               number implies, so the stored value did NOT move anything
'   GC04/GC05   runtime write and Move read back exactly (write side is the same
'               cache -- Move was the unverified half of ledger 193)
'   GC06        a pixel-mode container still reads pixels (the stored-twips cache
'               fails the unit gate there -- measured identical before/after)
'   GC07        switching the container's ScaleMode re-projects, switching back
'               returns the VB-side number again
'   GC08        self-heal: ComboBox is resized by the RTL at creation (dropdown
'               area), so its cached design height is stale and the read follows
'               the window instead -- the pixel gate is what makes that safe
'   GC09..GC12  the late-bound route (the RTL host model behind IDispatch) answers
'               the same numbers as the emitted one -- it used to carry a second
'               geometry implementation that never stored the VB-side value
'   GC17..GC19  a member enumerated out of Controls is that control and nobody else's,
'               and a With block over a Collection member writes that control
'   GC20..GC23  an enumerated member can say who it is (.Name / TypeName / .Index)
'               and Controls("name") finds that same window again

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

Private Function TF(ByVal b As Boolean) As String
    If b Then TF = "True" Else TF = "False"
End Function

' The window's own position in the parent's client pixels -- asked straight from
' user32, so it cannot be the cached number answering for itself.
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

' The witnesses below are built from kernel32 MulDiv + the real device DPI, i.e.
' from the same two primitives the RTL's placement uses -- but asked from outside
' the product, so the cached number cannot be answering for itself. (Me.ScaleX was
' the first thing to try here and is NOT used: a Double handed back to a VB Long
' truncates on this path, so it reads 13 where MulDiv says 14.)
Private dpiX As Long
Private dpiY As Long

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

Private Function ToTwipX(ByVal px As Long) As Long
    ToTwipX = MulDiv(px, 1440, dpiX)
End Function

Private Function ToTwipY(ByVal px As Long) As Long
    ToTwipY = MulDiv(px, 1440, dpiY)
End Function

Private Sub tmrGC_Timer()
    tmrGC.Enabled = False
    ReadDpi

    Dim ok1 As Boolean, ok2 As Boolean, ok3 As Boolean
    Dim ok4 As Boolean, ok5 As Boolean, ok6 As Boolean
    Dim ok7 As Boolean, ok8 As Boolean
    Dim pxSwitch As Long

    ok1 = (txtA.Left = 1007) And (txtA.Top = 449) And (txtA.Width = 3001) And (txtA.Height = 247)
    Debug.Print "GC01-design=" & TF(ok1) & " raw=" & txtA.Left & "," & txtA.Top & "," & txtA.Width & "," & txtA.Height

    ok2 = (lblP.Left = 311) And (lblP.Top = 222) And (lblP.Width = 907) And (lblP.Height = 211)
    Debug.Print "GC02-child=" & TF(ok2) & " raw=" & lblP.Left & "," & lblP.Top & "," & lblP.Width & "," & lblP.Height

    ok3 = (PxLeft(txtA.hwnd) = ToPx(1007)) And (PxTop(txtA.hwnd) = ToPy(449)) _
        And (PxW(txtA.hwnd) = ToPx(3001)) And (PxH(txtA.hwnd) = ToPy(247))
    Debug.Print "GC03-place=" & TF(ok3) & " raw=" & PxLeft(txtA.hwnd) & "," & PxTop(txtA.hwnd) & "," & PxW(txtA.hwnd) & "," & PxH(txtA.hwnd) & "," & ToPx(1007)

    txtA.Left = 5000
    txtA.Top = 444
    txtA.Width = 7777
    txtA.Height = 3001
    ok4 = (txtA.Left = 5000) And (txtA.Top = 444) And (txtA.Width = 7777) And (txtA.Height = 3001) _
        And (PxLeft(txtA.hwnd) = ToPx(5000)) And (PxW(txtA.hwnd) = ToPx(7777)) _
        And (PxTop(txtA.hwnd) = ToPy(444)) And (PxH(txtA.hwnd) = ToPy(3001))
    Debug.Print "GC04-write=" & TF(ok4) & " raw=" & txtA.Left & "," & txtA.Top & "," & txtA.Width & "," & txtA.Height

    txtA.Move 1231, 451, 3007, 247
    ok5 = (txtA.Left = 1231) And (txtA.Top = 451) And (txtA.Width = 3007) And (txtA.Height = 247) _
        And (PxTop(txtA.hwnd) = ToPy(451)) And (PxH(txtA.hwnd) = ToPy(247))
    Debug.Print "GC05-move=" & TF(ok5) & " raw=" & txtA.Left & "," & txtA.Top & "," & txtA.Width & "," & txtA.Height

    ok6 = (lblX.Left = PxLeft(lblX.hwnd)) And (lblX.Width = PxW(lblX.hwnd)) _
        And (PxLeft(lblX.hwnd) = ToPx(101)) And (PxW(lblX.hwnd) = ToPx(203))
    Debug.Print "GC06-pixbox=" & TF(ok6) & " raw=" & lblX.Left & "," & PxLeft(lblX.hwnd) & "," & lblX.Width & "," & ToPx(203)

    picP.ScaleMode = 3
    pxSwitch = lblP.Left
    ok7 = (pxSwitch = PxLeft(lblP.hwnd)) And (lblP.Width = PxW(lblP.hwnd))
    picP.ScaleMode = 1
    ok7 = ok7 And (lblP.Left = 311) And (lblP.Width = 907)
    Debug.Print "GC07-modesw=" & TF(ok7) & " raw=" & pxSwitch & "," & PxLeft(lblP.hwnd) & "," & lblP.Left

    ok8 = (cboA.Height <> 247) And (cboA.Height = ToTwipY(PxH(cboA.hwnd)))
    Debug.Print "GC08-stale=" & TF(ok8) & " raw=" & cboA.Height & ",247," & PxH(cboA.hwnd) & "," & ToTwipY(PxH(cboA.hwnd))

    ' GC09..GC12 -- ledger 247. The IDispatch route (the RTL host model, which is
    ' what a late-bound `With obj : .Width = ...` or a UserControl's `Parent.Width`
    ' really calls) carried its OWN geometry code: reads did GetWindowRect + a
    ' hardcoded twips conversion, writes moved all four axes at once and stored
    ' nothing. So the same property answered two different numbers depending on
    ' which route the statement took -- measured before the fix: `f.Width = 7222`
    ' read back 7222 on the emitted route only after the window had been moved,
    ' and 7215 on the late-bound one. Both routes now ask vb6forms_ctrl.c.
    Dim ok9 As Boolean, ok10 As Boolean, ok11 As Boolean, ok12 As Boolean
    Dim g As Object
    Set g = Me
    Dim gwHost As Long, ghHost As Long, gwBack As Long

    g.Width = 7222
    gwHost = CLng(g.Width)
    ok9 = (gwHost = 7222) And (Me.Width = 7222)
    Debug.Print "GC09-late-write=" & TF(ok9) & " raw=" & gwHost & "," & Me.Width

    g.Height = 3333
    ghHost = CLng(g.Height)
    ok10 = (ghHost = 3333) And (Me.Height = 3333)
    Debug.Print "GC10-late-height=" & TF(ok10) & " raw=" & ghHost & "," & Me.Height

    ok11 = (PxW(Me.hwnd) = ToPx(7222)) And (PxH(Me.hwnd) = ToPy(3333))
    Debug.Print "GC11-late-place=" & TF(ok11) & " raw=" & PxW(Me.hwnd) & "," & ToPx(7222) & "," & PxH(Me.hwnd) & "," & ToPy(3333)

    Me.Width = 6011
    gwBack = CLng(g.Width)
    ok12 = (gwBack = 6011) And (Me.Width = 6011)
    Debug.Print "GC12-direct-write=" & TF(ok12) & " raw=" & gwBack & "," & Me.Width

    ' GC13..GC14 -- ledger 250. The Controls arm of the host model asked the emitter
    ' that deliberately answers Empty, so a late-bound caller got NULL and every
    ' downstream branch (Count / Item / For Each -- all of which were already written
    ' for collections) saw an empty set. VB6: a form's Controls enumerates its own
    ' children, so both numbers below are counts of real windows (5 here: the timer,
    ' the combo, the textbox and the two picture boxes; the labels live in the boxes).
    ' What an item still cannot answer is its VB name -- standard controls created by
    ' the emitted route carry no host-side name record (measured: firstName is empty
    ' while TypeName answers Control). That half is ledger 252, not this one.
    Dim ok13 As Boolean, ok14 As Boolean
    Dim n As Long
    n = CLng(Me.Controls.Count)
    ok13 = (n >= 4) And (n <= 8)
    Debug.Print "GC13-coll-count=" & TF(ok13) & " raw=" & n

    Dim it As Long
    Dim o As Object
    it = 0
    For Each o In Me.Controls
        it = it + 1
    Next
    ok14 = (it >= 4) And (it <= 8)
    Debug.Print "GC14-foreach-iters=" & TF(ok14) & " raw=" & it

    ' GC15 -- ledger 254. VB6's Form.Count is the SAME set the Controls collection
    ' enumerates, so the two readings have to agree. The host model used to answer a
    ' hardcoded 0 here. Charts 2020's ClsResizer sizes its cache with
    ' ReDim Rects(oForm.Count - 1) and then fills it with For Each over Controls, so a
    ' 0 left it writing Rects(0) into an empty array -> runtime error 9 at startup.
    ' Read it late-bound (Set g = Me), which is the shape the resizer uses.
    ' The range is pinned TOO, not only the agreement: before ledger 250 both numbers
    ' were 0, so "fc = n" alone would read True on a compiler where nothing works.
    Dim ok15 As Boolean
    Dim fc As Long
    Set g = Me
    fc = CLng(g.Count)
    ok15 = (fc = n) And (fc >= 4) And (fc <= 8)
    Debug.Print "GC15-form-count=" & TF(ok15) & " raw=" & fc & "/" & n

    ' GC16 -- ledger 249. Reading the FORM's own Left inside a & concatenation handed
    ' the raw int straight to vb6_BSTR_Concat: the "is this object slot a control"
    ' authority did not know the MeExpr shape, so the member name was looked up as a
    ' bare module symbol, hit the String-returning VB builtin Left(), the type oracle
    ' answered String, and the concat face skipped the numeric wrap. The crash site IS
    ' the judge -- on the pre-fix compiler this line never prints (0xC0000005 mid-timer).
    ' The digits are pinned too: surviving is not the same as answering.
    ' (.Top / .Width / .Height never crashed, only because those names do not collide.)
    Dim ok16 As Boolean
    Dim s16 As String
    s16 = "L" & Me.Left
    ok16 = (Len(s16) > 2)
    Debug.Print "GC16-me-left-concat=" & TF(ok16) & "/" & s16

    ' GC17..GC19 -- ledger 255. Two knives, both on the route Charts 2020's ClsResizer
    ' walks: `For Each oCtrl In oForm.Controls` fills a Collection, and that collection
    ' is then written through `With CtrlNames(i)`.
    '   (b) each enumerated member comes back wrapped in a live COM object so the stack
    '       VARIANT can be VariantCleared. The (wrapper -> target) table was append-only
    '       while a freed block address gets handed to the next wrapper, so the scan from
    '       slot 0 always hit the FIRST stale pair and every member unwrapped to the same
    '       object: measured here before the fix all five members answered the Timer's
    '       hWnd and Left (120) and Width (0), and a Width write landed on the Timer while
    '       no control moved.
    '   (a) the With receiver was the address of the calloc'd VARIANT that the default
    '       Item returns, not the dispatch inside it; receivers are named by identity, so
    '       every member of the block fell through to "property ... not found" -- 168 of
    '       them on Charts 2020 and not one position applied.
    ' GC17 pins the members are DISTINCT (one sum only matches if each answers for
    ' itself). GC18/GC19 pin the write through With, both heads -- what the control
    ' reports and what the window itself says -- because "it only landed in a cache" is
    ' one of the ways this could still be wrong. picX is left out of the pins: it is a
    ' pixel-mode container, so its Width is not the same number in both units.
    Dim ok17 As Boolean, ok18 As Boolean, ok19 As Boolean
    Dim oC As Object
    Dim collC As Collection
    Dim iC As Long
    Dim sumW As Long, ownW As Long
    ownW = txtA.Width + picP.Width + picX.Width + cboA.Width + chkA.Width
    Set collC = New Collection
    For Each oC In Me.Controls
        sumW = sumW + CLng(oC.Width)
        collC.Add oC
    Next
    ok17 = (sumW = ownW) And (ownW > 0)
    Debug.Print "GC17-item-identity=" & TF(ok17) & " raw=" & sumW & "/" & ownW & "/" & CLng(collC.Count)

    For iC = 1 To CLng(collC.Count)
        With collC(iC)
            .Width = 1234
        End With
    Next
    ok18 = (txtA.Width = 1234) And (picP.Width = 1234) And (cboA.Width = 1234)
    Debug.Print "GC18-with-item-write=" & TF(ok18) & " raw=" & txtA.Width & "," & picP.Width & "," & cboA.Width & "," & picX.Width
    ok19 = (PxW(txtA.hwnd) = ToPx(1234)) And (PxW(picP.hwnd) = ToPx(1234))
    Debug.Print "GC19-with-item-place=" & TF(ok19) & " raw=" & PxW(txtA.hwnd) & "," & ToPx(1234) & "," & PxW(picP.hwnd)

    ' GC20..GC23 -- ledger 252. A control never carried its VB identity: the identity
    ' table (vb6_ucHo) has name/typeName/index columns and four readers that already ask
    ' it -- late-bound `.Name`, `.Index`, `TypeName(obj)` and Controls("name") -- but only
    ' FORMS were ever registered, so standard controls fell through IsWindow and answered
    ' empty name / "Control" / -1 / NULL. Measured on this form before the fix: all five
    ' members answered name=[] type=Control and Controls("picP") handed back nothing.
    ' Charts 2020 is the traffic: ClsResizer picks the font/property tables with
    ' `If TypeName(CtrlNames(i)) = FBuf(j).CtrlTypeName`, and "Control" never equals
    ' "LabelPlus.LabelClass", so not one row applied even after ledger 255 delivered the
    ' right receiver. The registration now happens at BOTH creation routes.
    ' `.Tag` through a member is still empty -- that is the late-bound string read,
    ' a separate face (family of #88), deliberately not pinned here.
    Dim ok20 As Boolean, ok21 As Boolean, ok22 As Boolean, ok23 As Boolean
    Dim m2 As Object
    Dim foundH As Long, foundT As String, foundI As Long
    foundH = 0
    foundT = ""
    foundI = -999
    For Each m2 In Me.Controls
        If CStr(m2.Name) = "txtA" Then
            foundH = CLng(m2.hWnd)
            foundT = TypeName(m2)
            foundI = CLng(m2.Index)
        End If
    Next
    ' two heads: the name has to point at the window the emitted route owns
    ok20 = (foundH <> 0) And (foundH = CLng(txtA.hwnd))
    Debug.Print "GC20-member-name=" & TF(ok20) & " raw=" & foundH & "/" & CLng(txtA.hwnd)
    ok21 = (foundT = "TextBox")
    Debug.Print "GC21-member-typename=" & TF(ok21) & " raw=[" & foundT & "]"
    ' a non-array control answers Index = -1 on both routes -- recorded so a future
    ' registration that "helpfully" passes 0 shows up as a red judge, not a silent change
    ok22 = (foundI = -1)
    Debug.Print "GC22-member-index=" & TF(ok22) & " raw=" & foundI
    Dim bk As Object
    Dim bkH As Long
    Set bk = Me.Controls("picP")
    bkH = CLng(bk.hWnd)
    ok23 = (bkH <> 0) And (bkH = CLng(picP.hwnd))
    Debug.Print "GC23-bynamed-lookup=" & TF(ok23) & " raw=" & bkH & "/" & CLng(picP.hwnd)

    ' GC24..GC27 -- ledger 257. The host model had no Container/Parent arm at all, so
    ' `oCtrl.Container` answered Empty and the save-loop line in Charts,
    ' `If oCtrl.Container.Name = oForm.Name Then`, passed only because BOTH sides were
    ' empty (probe255 mode P: container-type=[Empty], gate=True) -- a form registered
    ' itself with a NULL name, so "name = name" proved nothing. The arm now answers the
    ' same relation the Controls collection already uses for membership, so the gate
    ' means "this member is in this form's own set".
    ' Both heads pinned. True head: every member of Me.Controls that has a window names
    ' THIS form, and the name is non-empty on both sides -- plus a witness only the real
    ' window can give (its caption, and its hWnd compared against Me.hwnd), so an arm
    ' that merely echoed a string would not pass. False head: a member of picP's own set
    ' is contained by picP, so the same comparison must answer False -- that is exactly
    ' what the line filters on, and a build where every Container resolved to the form
    ' would fail it.
    Dim ok24 As Boolean, ok25 As Boolean, ok26 As Boolean, ok27 As Boolean, ok28 As Boolean
    Dim cntForm As Long, cntNone As Long, cntOther As Long
    Dim nmForm As String, nmTxt As String, capTxt As String
    Dim hTxt As Long, hSelf As Long
    Dim nmBoxMember As String, nmBoxContainer As String
    Dim pb As Object, m4 As Object
    cntForm = 0
    cntNone = 0
    cntOther = 0
    nmTxt = ""
    capTxt = ""
    hTxt = 0
    nmForm = CStr(Me.Name)
    hSelf = CLng(Me.hwnd)
    For Each oC In Me.Controls
        If TypeName(oC.Container) = "Form" Then
            If CStr(oC.Container.Name) = nmForm Then
                cntForm = cntForm + 1
            Else
                cntOther = cntOther + 1
            End If
        ElseIf TypeName(oC.Container) = "Empty" Then
            cntNone = cntNone + 1
        Else
            cntOther = cntOther + 1
        End If
        If CStr(oC.Name) = "txtA" Then
            nmTxt = CStr(oC.Container.Name)
            capTxt = CStr(oC.Container.Caption)
            hTxt = CLng(oC.Container.hWnd)
        End If
    Next
    ' no third answer, and a range too: "cntForm + cntNone + cntOther = n" alone would
    ' read True on a build where the arm is missing and every member answers Empty
    ' (the lesson of ledger 254 -- pin the identity AND the size).
    ok24 = (cntOther = 0) And (cntForm + cntNone = n) And (cntForm >= 3)
    Debug.Print "GC24-container-set=" & TF(ok24) & " raw=" & cntForm & "+" & cntNone & "+" & cntOther & "/" & n
    ok25 = (Len(nmForm) > 0) And (nmTxt = nmForm) And (nmTxt = "GCForm")
    Debug.Print "GC25-container-name=" & TF(ok25) & " raw=[" & nmForm & "]/[" & nmTxt & "]"
    ok26 = (capTxt = "GeomCache") And (hTxt <> 0) And (hTxt = hSelf)
    Debug.Print "GC26-container-window=" & TF(ok26) & " raw=[" & capTxt & "]/" & hTxt & "/" & hSelf
    nmBoxMember = ""
    nmBoxContainer = ""
    Set pb = Me.Controls("picP")
    For Each m4 In pb.Controls
        nmBoxMember = CStr(m4.Name)
        nmBoxContainer = CStr(m4.Container.Name)
    Next
    ok27 = (nmBoxMember = "lblP") And (nmBoxContainer = "picP") And (nmBoxContainer <> nmForm)
    Debug.Print "GC27-nested-container=" & TF(ok27) & " raw=[" & nmBoxMember & "]/[" & nmBoxContainer & "]/[" & nmForm & "]"
    ' The fourth head: the form itself has no container (top-level window, no parent), so
    ' the arm is provably not just handing the receiver back -- a self-reference would read
    ' "GCForm" here and GC25 would still pass. Late-bound through g (Set g = Me, GC15).
    Dim tySelf As String
    tySelf = TypeName(g.Container)
    ok28 = (tySelf = "Empty")
    Debug.Print "GC28-form-has-no-container=" & TF(ok28) & " raw=[" & tySelf & "]"

    ' GC29..GC31 -- ledger 256. A bare control name appearing in a LATE-BOUND argument
    ' position is folded into its default property (the emitted payload is a BSTR for
    ' Text/Caption, an int for CheckBox.Value), while the box around it came from a type
    ' oracle that answers "Object" for any control name. VB6 boxes VT_DISPATCH there, the
    ' stack VARIANT is VariantCleared at the end of the statement and Release is
    ' unconditional, so the payload's first bytes get read as a vtable -- 0xC0000005 right
    ' at the Add (measured with the pre-fix compiler on the same two arches this fixture
    ' runs on: the probe never reached its last line).
    ' The axis is NOT "object slot vs value slot": VB6 really does take the default
    ' property for a Variant parameter. What was wrong is that the box did not follow the
    ' value, so both rows of the new default-property type table get pinned here (String
    ' and Integer), and each reads back what the control itself says -- "it survived" is
    ' not the judgment.
    ' The third head is a witness only the window can give: the item has to be the string
    ' that is IN the control (txtA.Text read on the emitted route), not a literal copied
    ' into the judge.
    Dim ok29 As Boolean, ok30 As Boolean, ok31 As Boolean
    Dim k29 As String, k30 As String, k31 As String
    Dim ty29 As String, ty31 As String
    Dim v31 As Long
    txtA.Text = "GC29TXT"
    lblP.Caption = "GC30CAP"
    chkA.Value = 1
    collC.Add txtA, "k29"
    k29 = CStr(collC.Item("k29"))
    ty29 = TypeName(collC.Item("k29"))
    ok29 = (ty29 = "String") And (k29 = "GC29TXT") And (k29 = txtA.Text)
    Debug.Print "GC29-pack-text=" & TF(ok29) & " raw=[" & k29 & "]/[" & ty29 & "]"

    collC.Add lblP, "k30"
    k30 = CStr(collC.Item("k30"))
    ok30 = (k30 = "GC30CAP") And (k30 = lblP.Caption) And (k30 <> k29)
    Debug.Print "GC30-pack-caption=" & TF(ok30) & " raw=[" & k30 & "]"

    collC.Add chkA, "k31"
    v31 = CLng(collC.Item("k31"))
    ty31 = TypeName(collC.Item("k31"))
    k31 = CStr(chkA.Value)
    ' Integer lands in VT_I4 (vb6_ComPackInt) -- that width is today's packing口径,
    ' not this ledger's question, so the reading is pinned as-is rather than as
    ' "not Object": if the Integer box ever moves to VT_I2 this judge goes red on purpose.
    ok31 = (v31 = CLng(k31)) And (v31 = 1) And (ty31 = "Long")
    Debug.Print "GC31-pack-checkvalue=" & TF(ok31) & " raw=" & v31 & "/[" & ty31 & "]/" & k31

    Debug.Print "GC-DONE"
    Unload Me
End Sub
