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
   Begin VB.PictureBox picP
      Height          =   1680
      Left            =   3720
      ScaleHeight     =   1680
      ScaleWidth      =   1680
      TabIndex        =   3
      Top             =   60
      Width           =   1680
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

' 232-1 boundary judge: the fold may only take a name that resolves to NOTHING in
' this module. A user-written Sub named like a canvas verb has to win, otherwise the
' fix would trade "does not compile" for "silently calls the wrong function" -- the
' shape this line has been burned by before. Point is not drawn anywhere else on
' this form (the canvas Circle is FD24's job below), and the call is bare (no Me.).
Private userPoint As Long

Private Sub Point(x As Long, y As Long)
    userPoint = userPoint + 1
End Sub

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
    ' 278: this window used to start at x=150, which was calibrated to the DC's default
    ' font (16 px high -- 16 M's spanned 176 px). From the 278th knife the form paints
    ' with VB6's own font: read live off the reference process that LOGFONT is
    ' {-13 at 120 dpi, weight 400, "MS Sans Serif"} = -11 / 13 px high / 144 px of ink
    ' at 96 dpi, so 150 sat past the end of the blue line and only the row judge was
    ' left standing. Scan from x=60: still clear of the one-char red line (~9 px) and
    ' inside the blue run at 72 / 96 / 120 dpi.
    rowB = FirstRowOfColor(d, vbBlue, 60, 200, 20, 240)
    rowR = FirstRowOfColor(d, vbRed, 0, 10, 20, 240)
    ok16 = (rowR > 50) And (rowR = rowB)
    Debug.Print "FD16-paint-units=" & TF(ok16)
    Debug.Print "FD17-RAW rowB=" & CStr(rowB) & " rowR=" & CStr(rowR)
    r = ReleaseDC(Me.hwnd, d)

    ' ---- 232-2: the same two canvas entries written with an explicit receiver.
    ' Before, `Me.Cls` / `Me.Print "x"` compiled to a COM no-op on the form HWND --
    ' the canvas table answered Cls and Print for one receiver type only -- so the
    ' surface kept its ink and the pen never moved. Each judge is two-sided: an ink
    ' witness (a colour that is really there / really gone) AND the pen.
    Dim ok18 As Boolean
    Dim ok19 As Boolean
    Dim ok20 As Boolean
    Dim inkBefore As Long
    Dim inkAfter As Long
    Dim penBefore As Double
    Dim curA As Double
    Dim thA As Double

    Me.ScaleMode = vbPixels
    Me.ForeColor = vbRed
    Me.CurrentX = 0
    Me.CurrentY = 0
    Print "MMMMMMMMMMMMMMMM"
    d = GetDC(Me.hwnd)
    inkBefore = FirstRowOfColor(d, vbRed, 0, 40, 1, 120)
    Me.Cls
    inkAfter = FirstRowOfColor(d, vbRed, 0, 40, 1, 120)
    ok18 = (inkBefore > 0) And (inkAfter < 0) And (Abs(Me.CurrentY) < 0.001)
    Debug.Print "FD18-mecls=" & TF(ok18) & " raw=" & CStr(inkBefore) & "," & CStr(inkAfter)
    ' The increment is measured against the pen as it stood RIGHT BEFORE this line,
    ' not against zero: a no-op Print then reads 0 and cannot pass by inheriting an
    ' earlier line's advance (measured that way under the pre-fix compiler).
    penBefore = Me.CurrentY
    Me.Print "MMMMMMMMMMMMMMMM"
    inkAfter = FirstRowOfColor(d, vbRed, 0, 40, 1, 120)
    ok19 = (inkAfter > 0)
    curA = Me.CurrentY - penBefore
    thA = Me.TextHeight("MMMMMMMMMMMMMMMM")
    ok20 = (curA = thA)
    Debug.Print "FD19-mepaint=" & TF(ok19) & " raw=" & CStr(inkAfter)
    Debug.Print "FD20-mepen=" & TF(ok20) & " raw=" & CStr(curA) & "," & CStr(thA)

    ' 232-1 = the BARE spellings of the same two verbs (no Me.). Before, the form
    ' never became the implicit receiver: Cls and PSet (x, y) reached codegen as bare
    ' C calls. Measured on the pre-fix compiler with this very fixture: the emit held
    ' `Cls();` and `PSet(170, 130);`, Cls also drew a VB3001 "undeclared identifier",
    ' and the real build died at LNK2019 on both names (no exe at all). The semantic
    ' layer now folds them onto the form canvas, which is the route Me.Cls / Me.PSet
    ' already took: same entry, same ink.
    Dim ok21 As Boolean
    Dim ok22 As Boolean
    Me.ForeColor = vbRed
    Me.CurrentX = 0
    Me.CurrentY = 0
    Print "MMMMMMMMMMMMMMMM"
    inkBefore = FirstRowOfColor(d, vbRed, 0, 40, 1, 120)
    Cls
    inkAfter = FirstRowOfColor(d, vbRed, 0, 40, 1, 120)
    ok21 = (inkBefore > 0) And (inkAfter < 0)
    Debug.Print "FD21-barecls=" & TF(ok21) & " raw=" & CStr(inkBefore) & "," & CStr(inkAfter)
    Me.ForeColor = vbBlue
    Me.DrawWidth = 7
    PSet (170, 130)
    rowB = FirstRowOfColor(d, vbBlue, 150, 200, 120, 160)
    ok22 = (rowB > 0)
    Debug.Print "FD22-barepset=" & TF(ok22) & " raw=" & CStr(rowB)

    ' 232-3 = the bare verb WITH a tail the parser used to drop. `Circle (x, y), r`
    ' and `PSet (x, y), color` lost everything after the coordinate pair (the emit
    ' measured before this cut: `Circle(20, 21, vb6_VariantEmpty(), 22);` -- a bare
    ' undeclared call whose radius only looks absorbed), and a bare
    ' `Line (a,b)-(c,d)` did not parse at all (VB2001 expected ')' / VB2003 / VB2002,
    ' three errors off one line). Both judges are two-sided on a colour nothing else
    ' on this form uses: the scan must come back empty first, then find drawn rows.
    Dim ok24 As Boolean, ok25 As Boolean
    Dim rowG0 As Long, rowG As Long
    Dim rowM0 As Long, rowM As Long
    Me.ForeColor = vbGreen
    Me.DrawWidth = 4
    rowG0 = FirstRowOfColor(d, vbGreen, 205, 275, 55, 125)
    Circle (240, 90), 30
    rowG = FirstRowOfColor(d, vbGreen, 205, 275, 55, 125)
    ok24 = (rowG0 < 0) And (rowG > 0)
    Debug.Print "FD24-barecircle=" & TF(ok24) & " raw=" & CStr(rowG0) & "," & CStr(rowG)
    Me.ForeColor = vbMagenta
    Me.DrawWidth = 5
    rowM0 = FirstRowOfColor(d, vbMagenta, 294, 306, 35, 125)
    Line (300, 40)-(300, 120)
    rowM = FirstRowOfColor(d, vbMagenta, 294, 306, 35, 125)
    ok25 = (rowM0 < 0) And (rowM > 0)
    Debug.Print "FD25-bareline=" & TF(ok25) & " raw=" & CStr(rowM0) & "," & CStr(rowM)
    ' 232-4 = `Line -(x, y)`: the START is omitted (VB6: from CurrentX/CurrentY). Before this
    ' cut the form was rejected in the parser (VB2001 "expected 'Input' after 'Line'" plus the
    ' `-(310` read as an infix minus), so there was no build to run at all. Two-sided like
    ' FD24/FD25 on colours nothing else here uses, and each judge ALSO reads the pen: the ink
    ' has to land on the endpoint the source named, and the pen must end there (the RTL moves
    ' it to the endpoint when has2 -- vb6forms_draw.c:207).
    ' ⚠ The coordinates live in the LEFT band on purpose: this form is 6400x3200 twips, which
    ' is **426x213 px at 96 dpi** (1440 twips = 1 inch), not 640x320. The first version of
    ' these two judges used x up to 560 and CI read `raw=-1,-1 pen=560,140` on the `Me.` one:
    ' the pen moved (the call is right) but the whole segment sat past the right edge of the
    ' client area, so nothing was ever painted. A pixel judge whose window is off the window
    ' is a False that reads exactly like a broken codegen -- keep the whole segment inside.
    ' Also keep it clear of FD24's circle (205..275) and FD25's magenta line (298..302); the
    ' scans filter by colour, but overlapping ink would still be one more thing to explain.
    ' The pen half is what rejects "read the two given coordinates as the START": that shape
    ' ends the pen at (0,0) (measured: `pen=0,0` on the pre-fix dock), while its ink would
    ' still cross these windows.
    Dim ok26 As Boolean, ok27 As Boolean
    Dim rowC0 As Long, rowC As Long
    Dim rowY0 As Long, rowY As Long
    Me.ForeColor = vbCyan
    Me.DrawWidth = 3
    Me.CurrentX = 40
    Me.CurrentY = 30
    rowC0 = FirstRowOfColor(d, vbCyan, 35, 125, 25, 105)
    Line -(120, 100)
    rowC = FirstRowOfColor(d, vbCyan, 35, 125, 25, 105)
    ok26 = (rowC0 < 0) And (rowC > 0) And (Me.CurrentX = 120) And (Me.CurrentY = 100)
Debug.Print "FD26-barelinestart=" & TF(ok26) & " raw=" & CStr(rowC0) & "," & CStr(rowC) & _
              " pen=" & CStr(Me.CurrentX) & "," & CStr(Me.CurrentY)
    Me.ForeColor = vbYellow
    Me.DrawWidth = 3
    Me.CurrentX = 40
    Me.CurrentY = 120
    rowY0 = FirstRowOfColor(d, vbYellow, 35, 125, 115, 195)
    Me.Line -(120, 190)
    rowY = FirstRowOfColor(d, vbYellow, 35, 125, 115, 195)
    ok27 = (rowY0 < 0) And (rowY > 0) And (Me.CurrentX = 120) And (Me.CurrentY = 190)
Debug.Print "FD27-melinestart=" & TF(ok27) & " raw=" & CStr(rowY0) & "," & CStr(rowY) & _
              " pen=" & CStr(Me.CurrentX) & "," & CStr(Me.CurrentY)
    ' 232-5 = PictureBox.Line -(x,y): the same start-omitted form as FD26/FD27, on a PictureBox.
    ' Before, vb6_ControlLine was the only sibling not forwarded to vb6_Form_Line (账 #239) -- it
    ' kept a 7-arg body with no has bit, so Line -(x,y) drew (x,y)->(0,0) and left the pen where it
    ' was. After forwarding, has1=0 => start from picP.CurrentX/CurrentY, pen lands on the endpoint.
    ' Two-sided like FD26/FD27: a colour scan that comes back empty then finds the rows, AND the pen
    ' on the endpoint. The pen half is what rejects "read the two given coordinates as the START" --
    ' that shape leaves the pen untouched (20 != 100).
    ' ⚠ picP sits at the **far right** on purpose, and this is not cosmetic: a child window paints
    ' OVER the form's DC, so the first version of this block (Left = 3360 twips ~ 280 px) covered
    ' FD25's magenta line at x ~ 300 px and turned THAT judge red on the gate (run 39232007659:
    ' FD28/FD29 green, FD25-bareline raw=-1,-1 where it had read row 38 two runs earlier).
    ' Every existing judge on this form lives at x <= 306, so picP starts at 3720 twips (~310 px)
    ' and is 1680 twips square (~112 px) -- the whole of it inside the form's 426 px client area,
    ' with the two segments kept at x,y <= 110 (picP's own client, see draw-fixture-client-area-pixels).
    Dim dp As LongPtr
    Dim ok28 As Boolean, ok29 As Boolean
    Dim rowP0 As Long, rowP As Long
    Dim rowQ0 As Long, rowQ As Long
    picP.ScaleMode = vbPixels
    dp = GetDC(picP.hwnd)
    picP.CurrentX = 20
    picP.CurrentY = 20
    rowP0 = FirstRowOfColor(dp, vbCyan, 15, 105, 15, 65)
    picP.Line -(100, 60), vbCyan
    rowP = FirstRowOfColor(dp, vbCyan, 15, 105, 15, 65)
    ok28 = (rowP0 < 0) And (rowP > 0) And (picP.CurrentX = 100) And (picP.CurrentY = 60)
    Debug.Print "FD28-piclinestart=" & TF(ok28) & " raw=" & CStr(rowP0) & "," & CStr(rowP) & _
               " pen=" & CStr(picP.CurrentX) & "," & CStr(picP.CurrentY)
    picP.CurrentX = 20
    picP.CurrentY = 70
    rowQ0 = FirstRowOfColor(dp, vbYellow, 15, 105, 65, 115)
    picP.Line -(100, 110), vbYellow
    rowQ = FirstRowOfColor(dp, vbYellow, 15, 105, 65, 115)
    ok29 = (rowQ0 < 0) And (rowQ > 0) And (picP.CurrentX = 100) And (picP.CurrentY = 110)
    Debug.Print "FD29-piclinestart2=" & TF(ok29) & " raw=" & CStr(rowQ0) & "," & CStr(rowQ) & _
               " pen=" & CStr(picP.CurrentX) & "," & CStr(picP.CurrentY)
    ' 232-6 = `Step` 前缀: 坐标相对 CurrentX/CurrentY。改前这一族在 parser 就被拒 (`Line Step ...`
    ' 报 expected 'Input' after 'Line'; `Circle Step (x,y), r` 报 expected ')' (got ,))。
    ' 两格画在窗体右侧那块**空带**上 (x 210..300 / y 140..205): 这块窗体上别的一切判据都在
    ' x <= 306, 其中 FD24 的圆止于 y=125、FD22 的点止于 y=133 —— 新的线绝不压到它们
    ' （FD18 那一课: 子窗口/新笔画到别人的扫描窗里, 只会把别人的判据弄红）。
    ' 颜色用 RGB() 而不是又一枚 vb 常量: 这块窗体上 vb 常量已经用尽, 而 RGB(255,128,0) /
    ' RGB(128,0,255) 在别处一个像素都没有 —— 两枚判据的扫描窗也因此互不相干。
    ' pen 半钉的是「Step 没生效」那一形: 相对坐标被当绝对坐标时, 落点与笔位都不在源码写的地方。
    Dim ok30 As Boolean, ok31 As Boolean
    Dim pixA0 As Long, pixA As Long
    Dim rowU0 As Long, rowU As Long
    Me.ForeColor = RGB(255, 128, 0)
    Me.DrawWidth = 1
    Me.CurrentX = 220
    Me.CurrentY = 150
    pixA0 = FirstRowOfColor(d, RGB(255, 128, 0), 255, 265, 175, 185)
    Me.PSet Step (40, 30)
    pixA = FirstRowOfColor(d, RGB(255, 128, 0), 255, 265, 175, 185)
    ok30 = (pixA0 < 0) And (pixA > 0) And (Me.CurrentX = 260) And (Me.CurrentY = 180)
    Debug.Print "FD30-psetstep=" & TF(ok30) & " raw=" & CStr(pixA0) & "," & CStr(pixA) & _
               " pen=" & CStr(Me.CurrentX) & "," & CStr(Me.CurrentY)
    Me.ForeColor = RGB(128, 0, 255)
    Me.CurrentX = 220
    Me.CurrentY = 200
    rowU0 = FirstRowOfColor(d, RGB(128, 0, 255), 235, 265, 145, 205)
    Me.Line Step (40, 0) Step -(20, -50)
    rowU = FirstRowOfColor(d, RGB(128, 0, 255), 235, 265, 145, 205)
    ok31 = (rowU0 < 0) And (rowU > 0) And (Me.CurrentX = 240) And (Me.CurrentY = 150)
    Debug.Print "FD31-linestep=" & TF(ok31) & " raw=" & CStr(rowU0) & "," & CStr(rowU) & _
               " pen=" & CStr(Me.CurrentX) & "," & CStr(Me.CurrentY)
    ' 同一族在 PictureBox 上再走一遍: 那一支码头（vb6_ControlLine）自己也要读 Step 旗标,
    ' 不是借 Form 那一支。颜色第三枚, 与 FD28/FD29/FD30/FD31 都不撞。
    Dim ok32 As Boolean
    Dim rowT0 As Long, rowT As Long
    picP.ForeColor = RGB(0, 128, 128)
    picP.CurrentX = 5
    picP.CurrentY = 100
    rowT0 = FirstRowOfColor(dp, RGB(0, 128, 128), 10, 30, 95, 112)
    picP.Line Step (10, 5) Step -(20, 0)
    rowT = FirstRowOfColor(dp, RGB(0, 128, 128), 10, 30, 95, 112)
    ok32 = (rowT0 < 0) And (rowT > 0) And (picP.CurrentX = 25) And (picP.CurrentY = 100)
    Debug.Print "FD32-piclinestep=" & TF(ok32) & " raw=" & CStr(rowT0) & "," & CStr(rowT) & _
               " pen=" & CStr(picP.CurrentX) & "," & CStr(picP.CurrentY)
    ' 305 B227: 画布动词**不写括号**那一形 (`Me.Scale` / 裸 `Scale`) = 复位自定义坐标系。
    ' 改前这两形落到 vb6_ComCall(hwnd, L"Scale", NULL, 0) —— 假 IDispatch, 编得过、链接过、
    ' 跑起来一声不响, 于是复位不生效: `after=` 那一半读到的还是设进去的 10/20 (30/40)。
    ' 两枚各钉两头: 设进去读得到 (证明 Me.Scale (x1,y1)-(x2,y2) 那一形没坏) +
    ' 复位后读回 0 (证明不写括号那一形真的到了 RTL 出口)。
    Dim ok33 As Boolean, ok34 As Boolean
    Dim sl33 As Long, st33 As Long, sl34 As Long, st34 As Long
    Me.Scale (10, 20)-(100, 100)
    sl33 = Me.ScaleLeft
    st33 = Me.ScaleTop
    Me.Scale
    ok33 = (sl33 = 10) And (st33 = 20) And (Me.ScaleLeft = 0) And (Me.ScaleTop = 0)
    Debug.Print "FD33-mescalereset=" & TF(ok33) & " raw=" & CStr(sl33) & "," & CStr(st33) & _
               " after=" & CStr(Me.ScaleLeft) & "," & CStr(Me.ScaleTop)
    Me.Scale (30, 40)-(200, 200)
    sl34 = Me.ScaleLeft
    st34 = Me.ScaleTop
    Scale
    ok34 = (sl34 = 30) And (st34 = 40) And (Me.ScaleLeft = 0) And (Me.ScaleTop = 0)
    Debug.Print "FD34-barescalereset=" & TF(ok34) & " raw=" & CStr(sl34) & "," & CStr(st34) & _
               " after=" & CStr(Me.ScaleLeft) & "," & CStr(Me.ScaleTop)
    r = ReleaseDC(picP.hwnd, dp)
    Point 7, 9
    Debug.Print "FD23-userpoint=" & CStr(userPoint)
    r = ReleaseDC(Me.hwnd, d)
    Debug.Print "FD-DONE"
    Unload Me
End Sub

' 278 B113 (ASCII only, dead code -- never called): in VB6 a form name in an Unload
' operand is the form object, not its default property (Caption). This pins the emitted
' shape so a regression that folds it back to a BSTR into the HWND slot goes red.
Public Sub UsOwnNameProbe()
    Unload FDForm
End Sub
