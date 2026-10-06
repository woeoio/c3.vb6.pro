VERSION 5.00
Begin VB.Form PbForm
   Caption         =   "PbSub"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Pic1
     BackColor       =   255
     Height          =   1215
     Left            =   120
     ScaleHeight     =   1155
     ScaleWidth      =   1575
     TabIndex        =   1
     Top             =   120
     Width           =   1635
   End
   Begin VB.Image Img1
     Height          =   480
     Left            =   2040
     Top             =   120
     Width           =   480
   End
   Begin VB.Label Lb1
     Caption         =   "witness"
     Height          =   255
     Left            =   2040
     TabIndex        =   2
     Top             =   720
     Width           =   1695
   End
   Begin VB.Timer tmrStep
     Enabled         =   -1  'True
     Interval        =   150
   End
End
Attribute VB_Name = "PbForm"
Option Explicit

' 账 #185 的判据夹具。
'   PictureBox / Image 由 RTL 在建窗时装一枚**自绘子类** (STATIC+SS_BITMAP: 填 BackColor、
'   画 Picture), 带事件的控件由发码装一枚**事件子类**。两层以前把原始窗口过程存在同一个
'   窗口属性名上, 又都写"属性已存在就不装" ⇒ 后装的事件层被静默丢掉: Paint / MouseDown /
'   MouseUp / Click 一次也不响 (实测 x86 与 x64 各 0 条), 而同一窗体上的 Label (也是 STATIC,
'   只是不含 SS_BITMAP ⇒ 没人跟它抢) 一直好好的 —— 这条对比就是判别力所在。
'   现在两层各用各的槽位, 画的序 = 最外层 BeginPaint/EndPaint 一次, DC 经 VB6_PaintDC 交给
'   下面那层画表面, 再抬用户的 _Paint ⇒ 用户笔画在表面之上 (VB6 口径)。
'   ⇒ 两头都要钉: 事件臂要响 (PB02..PB06), 表面不能丢 (PB01 里那颗像素必须是红色)。
' 自驱: Timer 第一拍 PostMessage 到自己家的三枚子窗, 第二拍汇总计数后自退 —— 门禁不碰真实
'   鼠标 (见账 #79 那条级联偏移的教训), 计数写成"每步几条"的自洽式。

Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long

Private mStep As Integer
Private nPaint As Integer
Private nDown As Integer
Private nUp As Integer
Private nClick As Integer
Private nImg As Integer
Private nLb As Integer

Private Sub Form_Load()
    Debug.Print "PB00-LOAD"
End Sub

Private Sub tmrStep_Timer()
    Dim lp As LongPtr
    Dim allIn As Boolean
    lp = 20 + 20 * 65536          ' lParam = (x=20, y=20), 三枚子窗都覆盖得到
    ' 门 #332 的读数: 这一形以前在**固定第三拍**读计数, 而五条通知是 PostMessage 发出去的,
    ' 饿机器上到齐要用到第 8 拍 (本地 24 趟实测, 见台账 §B45) ⇒ 饿的时候读到全 0, 判据就红了。
    ' 现在把"等多久"和"断言什么"拆开: 到齐就读, 到不了就等满 20 拍再读 —— 真丢通知仍然读成全 0 红。
    allIn = (nDown >= 1 And nUp >= 1 And nClick >= 1 And nImg >= 1 And nLb >= 1)
    If mStep = 0 Then
        Pic1.Refresh                        ' 走一遍分层绘制 (表面 + 用户笔画)
        Call PostMessage(Pic1.hwnd, &H201, 1, lp)   ' WM_LBUTTONDOWN
        Call PostMessage(Pic1.hwnd, &H202, 0, lp)   ' WM_LBUTTONUP  => MouseUp + Click
        Call PostMessage(Img1.hwnd, &H201, 1, lp)
        Call PostMessage(Img1.hwnd, &H202, 0, lp)   ' => Image 的 Click
        Call PostMessage(Lb1.hwnd, &H201, 1, lp)    ' => Label 的 MouseDown (证人)
    ElseIf allIn Or mStep >= 20 Then
        ' 计数写成自洽式: 鼠标那五条是"每发一条消息恰好一条通知"(精确值才是翻倍探测器),
        ' paint 的**条数**不由我们定 (窗口显示/遮挡都会再要一次重绘) ⇒ 只问"有没有至少一条"。
        Dim paintOk As Integer
        If nPaint >= 1 Then paintOk = 1
        Debug.Print "PB-CNT mdown=" & nDown & " mup=" & nUp & _
                    " click=" & nClick & " img=" & nImg & " lb=" & nLb & " paint_ok=" & paintOk
        Debug.Print "PB-WAIT done=" & CStr(allIn)
        tmrStep.Enabled = False
        Debug.Print "PB-DONE"
        Unload Me
    End If
    mStep = mStep + 1
End Sub

Private Sub Pic1_Paint()
    Dim d As LongPtr
    nPaint = nPaint + 1
    d = GetDC(Pic1.hwnd)
    ' 表面必须已经画好: 设计期 BackColor = 255 (红), COLORREF 是 0x0000FF = 255。
    ' 读到 16777215 (白) = 自绘那一层没画或画在用户笔画之后 —— 本条就是要拦这个。
    Debug.Print "PB01-PAINT px=" & CStr(GetPixel(d, 2, 2))
    ReleaseDC Pic1.hwnd, d
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    nDown = nDown + 1
    Debug.Print "PB02-PIC-MDOWN x=" & Int(X)
End Sub

Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    nUp = nUp + 1
    Debug.Print "PB03-PIC-MUP"
End Sub

Private Sub Pic1_Click()
    nClick = nClick + 1
    Debug.Print "PB04-PIC-CLICK"
End Sub

Private Sub Img1_Click()
    nImg = nImg + 1
    Debug.Print "PB05-IMG-CLICK"
End Sub

Private Sub Lb1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    nLb = nLb + 1
    Debug.Print "PB06-LB-MDOWN"
End Sub
