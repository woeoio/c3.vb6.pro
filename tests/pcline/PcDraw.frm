VERSION 5.00
Begin VB.Form PcDrawForm 
   Caption         =   "PcDraw"
   ClientHeight    =   3200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6400
   ScaleHeight     =   3200
   ScaleWidth      =   6400
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picC 
      BackColor       =   &H00FFFFFF&
      Height          =   1200
      Left            =   120
      ScaleHeight     =   1140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   4500
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1560
      Left            =   120
      ScaleHeight     =   1500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   4500
   End
   Begin VB.Timer tmrP 
      Interval        =   150
      Left            =   120
      Top             =   1560
   End
End
Attribute VB_Name = "PcDrawForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 账 #221 = C29-PL-a 的运行期判据。这条方法的缺陷天生**只有画完再问像素才看得见**：
' 改前那 8 处 Line 全发成 `ComGetObjectProp(hwnd, L"Line")` 再对它取 `Item`，两跳都在
' RTL 登记过的"认识但什么都不做"里空转 ⇒ 编得过、跑得起、一笔不画、一条诊断也不打。
' 四形每形钉**两头**：画过的那一枚像素必须等于交进去的颜色，没画过的那一枚必须不等于它
' —— 只钉前者会放过"整片刷成红"那种假绿；只钉后者会放过"根本没落笔"。
' 画与问都放在 Timer 第一拍: 窗口问题要在窗口活着的时候问 (dcsurf 那条口径) ——
' 先试在 Form_Load 里画+问, 实测 GetPixel 一律回 -1 (CLR_INVALID), 因为那时候句柄还归不出 DC。
' picC 按 ScaleMode=3(像素) 走 ⇒ 探针坐标就是设备像素, 不随 DPI 抖; 缇那一档已由
' #196/#197/#198 三格钉着, 这一格不重复问单位。椭圆按"边上有、中心没有"判:
' 切点取整在 ±1 内, 所以对 y 开一个小窗口找, 不钉死一个坐标。
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
' DC 取法照 tests/pbsub/PbForm.frm: GetDC(控件 hwnd)。picC.hDC 在这枚夹具上读出 0,
' 那是 #196 那一族的另一问, 不在本刀里改。
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

' First device row in [lo,hi] whose ink differs from the box's own background
' (a colour match would silently fail if the ink never reached this DC).
Private Function FirstInk(d As LongPtr, bg As Long, lo As Long, hi As Long) As Long
    Dim row As Long
    Dim i As Long
    Dim cnt As Long
    FirstInk = -1
    For row = lo To hi
        If FirstInk < 0 Then
            cnt = 0
            For i = 0 To 60
                If GetPixel(d, i, row) <> bg Then cnt = cnt + 1
            Next
            If cnt > 3 Then FirstInk = row
        End If
    Next
End Function

Private Sub tmrP_Timer()
    Dim d As LongPtr
    Dim r As Long
    Dim i As Long
    Dim hitEdge As Boolean
    Dim okLine As Boolean, okBox As Boolean, okFill As Boolean, okCircle As Boolean
    Dim d2 As LongPtr
    Dim wb As Long
    Dim dp As Double
    Dim dtw As Double
    Dim stk As Double
    Dim th As Double
    Dim inkTop As Long
    Dim inkNear As Long
    Dim okPen As Boolean
    Dim okCls As Boolean
    Dim okTwip As Boolean
    Dim okBlack As Boolean
    Dim okStack As Boolean

    tmrP.Enabled = False

    ' 四形都画在 picC 里, 各自占一段 x 区间, 互不重叠
    picC.Line (0, 0)-(40, 40), vbRed              ' 默认: 线段
    picC.Line (60, 0)-(100, 40), vbBlue, B        ' B: 空心框
    picC.Line (120, 0)-(160, 40), vbGreen, BF     ' BF: 实心框
    picC.Line (180, 0)-(220, 40), vbRed, C        ' C: 椭圆(不填)

    d = GetDC(picC.hwnd)
    ' 线段: (20,20) 正落在 0,0 -> 40,40 那条对角线上; (20,5) 在旁边
    okLine = (GetPixel(d, 20, 20) = vbRed) And (GetPixel(d, 20, 5) <> vbRed)
    ' 空心框: 左边框 x=60 那一条必须是蓝, 正中心 (80,20) 必须**不是**蓝(没填)
    okBox = (GetPixel(d, 60, 20) = vbBlue) And (GetPixel(d, 80, 20) <> vbBlue)
    ' 实心框: 正中心 (140,20) 必须是绿 —— 这一条同时证明 BF 与 B 折出来不是同一个数
    okFill = (GetPixel(d, 140, 20) = vbGreen)
    ' 椭圆: 左边切点 x=180 那一列在 y=15..25 之间必须有一枚红; 中心 (200,20) 必须不是红
    hitEdge = False
    For i = 15 To 25
        If GetPixel(d, 180, i) = vbRed Then hitEdge = True
    Next
    okCircle = hitEdge And (GetPixel(d, 200, 20) <> vbRed)

    Debug.Print "PL00-DC=" & CStr(d)
    Debug.Print "PL01-LINE=" & CStr(okLine)
    Debug.Print "PL02-BOX=" & CStr(okBox)
    Debug.Print "PL03-FILL=" & CStr(okFill)
    Debug.Print "PL04-CIRCLE=" & CStr(okCircle)
    Debug.Print "PL05-RAW diag=" & CStr(GetPixel(d, 20, 20))
    Debug.Print "PL06-RAW boxedge=" & CStr(GetPixel(d, 60, 20)) & " boxmid=" & CStr(GetPixel(d, 80, 20))
    Debug.Print "PL07-RAW fillmid=" & CStr(GetPixel(d, 140, 20)) & " circletop=" & CStr(GetPixel(d, 200, 0))

    ' ---- account 239: the Print pen of a PictureBox -----------------------------
    ' Before this blade picP.Print drew with its OWN pixel cursor (window props
    '  VB6_PrintX / VB6_PrintY): it neither read picP.CurrentX/CurrentY nor moved
    '  them. Measured on the pre-fix build: advance = 0 while the same box answered
    '  TextHeight = 13, and with the pen sitting at row 60 the ink still landed on
    '  row 2. Print/Cls now call the one implementation the Form family uses, so the
    '  pen is the single float store and the unit is this window's own ScaleMode.
    d2 = GetDC(picP.hwnd)
    wb = GetPixel(d2, 200, 5)
    picP.CurrentX = 0
    picP.CurrentY = 60
    picP.Print "AB"
    th = picP.TextHeight("AB")
    dp = picP.CurrentY - 60
    inkTop = FirstInk(d2, wb, 0, 40)
    inkNear = FirstInk(d2, wb, 45, 95)
    okPen = (Abs(dp - th) < 0.001) And (inkTop = -1) And (inkNear >= 45)

    picP.Cls
    okCls = (Abs(picP.CurrentY) < 0.001)

    picP.ScaleMode = vbTwips
    picP.CurrentX = 0
    picP.CurrentY = 600
    picP.Print "CD"
    dtw = picP.CurrentY - 600
    okTwip = (Abs(dtw - picP.TextHeight("CD")) < 0.001)

    ' Cls fills with the stored background. Black is the hard case: stored as a
    ' window property, 0 is indistinguishable from "never set", so the answer must
    ' come from the one getter that carries the Fix 187 set-sentinel.
    picP.BackColor = vbBlack
    picP.Cls
    okBlack = (GetPixel(d2, 200, 40) = vbBlack)

    picP.BackColor = vbWhite
    picP.ScaleMode = vbPixels
    picP.Cls
    picP.CurrentX = 0
    picP.CurrentY = 0
    picP.Print "AB"
    picP.Print "CD"
    stk = picP.CurrentY
    okStack = (Abs(stk - 2 * picP.TextHeight("AB")) < 0.001)

    Debug.Print "PL08-PEN=" & CStr(okPen)
    Debug.Print "PL09-CLSPEN=" & CStr(okCls)
    Debug.Print "PL10-TWIPADV=" & CStr(okTwip)
    Debug.Print "PL11-BLACKCLS=" & CStr(okBlack)
    Debug.Print "PL12-STACK=" & CStr(okStack)
    Debug.Print "PL13-RAW dp=" & CStr(dp) & " th=" & CStr(th) & " inkTop=" & CStr(inkTop)
    Debug.Print "PL14-RAW twip=" & CStr(dtw) & " stack=" & CStr(stk)
    Debug.Print "PL-DONE"
    r = ReleaseDC(picP.hwnd, d2)
    r = ReleaseDC(picC.hwnd, d)
    Unload Me
End Sub
