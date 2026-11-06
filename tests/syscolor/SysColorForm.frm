VERSION 5.00
Begin VB.Form SysColorForm
   Caption         =   "SysColor"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1100
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   1100
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrStep
     Enabled         =   -1  'True
     Interval        =   200
   End
   Begin VB.PictureBox T0
     BackColor       =   &H80000000&
     Height          =   120
     Left            =   90
     Top             =   90
     Width           =   120
   End
   Begin VB.PictureBox T1
     BackColor       =   &H80000001&
     Height          =   120
     Left            =   210
     Top             =   90
     Width           =   120
   End
   Begin VB.PictureBox T2
     BackColor       =   &H80000002&
     Height          =   120
     Left            =   330
     Top             =   90
     Width           =   120
   End
   Begin VB.PictureBox T3
     BackColor       =   &H80000003&
     Height          =   120
     Left            =   450
     Top             =   90
     Width           =   120
   End
   Begin VB.PictureBox T4
     BackColor       =   &H80000004&
     Height          =   120
     Left            =   570
     Top             =   90
     Width           =   120
   End
   Begin VB.PictureBox T5
     BackColor       =   &H80000005&
     Height          =   120
     Left            =   690
     Top             =   90
     Width           =   120
   End
   Begin VB.PictureBox T6
     BackColor       =   &H80000006&
     Height          =   120
     Left            =   90
     Top             =   210
     Width           =   120
   End
   Begin VB.PictureBox T7
     BackColor       =   &H80000007&
     Height          =   120
     Left            =   210
     Top             =   210
     Width           =   120
   End
   Begin VB.PictureBox T8
     BackColor       =   &H80000008&
     Height          =   120
     Left            =   330
     Top             =   210
     Width           =   120
   End
   Begin VB.PictureBox T9
     BackColor       =   &H80000009&
     Height          =   120
     Left            =   450
     Top             =   210
     Width           =   120
   End
   Begin VB.PictureBox T10
     BackColor       =   &H8000000A&
     Height          =   120
     Left            =   570
     Top             =   210
     Width           =   120
   End
   Begin VB.PictureBox T11
     BackColor       =   &H8000000B&
     Height          =   120
     Left            =   690
     Top             =   210
     Width           =   120
   End
   Begin VB.PictureBox T12
     BackColor       =   &H8000000C&
     Height          =   120
     Left            =   90
     Top             =   330
     Width           =   120
   End
   Begin VB.PictureBox T13
     BackColor       =   &H8000000D&
     Height          =   120
     Left            =   210
     Top             =   330
     Width           =   120
   End
   Begin VB.PictureBox T14
     BackColor       =   &H8000000E&
     Height          =   120
     Left            =   330
     Top             =   330
     Width           =   120
   End
   Begin VB.PictureBox T15
     BackColor       =   &H8000000F&
     Height          =   120
     Left            =   450
     Top             =   330
     Width           =   120
   End
   Begin VB.PictureBox T16
     BackColor       =   &H80000010&
     Height          =   120
     Left            =   570
     Top             =   330
     Width           =   120
   End
   Begin VB.PictureBox T17
     BackColor       =   &H80000011&
     Height          =   120
     Left            =   690
     Top             =   330
     Width           =   120
   End
   Begin VB.PictureBox T18
     BackColor       =   &H80000012&
     Height          =   120
     Left            =   90
     Top             =   450
     Width           =   120
   End
   Begin VB.PictureBox T19
     BackColor       =   &H80000013&
     Height          =   120
     Left            =   210
     Top             =   450
     Width           =   120
   End
   Begin VB.PictureBox T20
     BackColor       =   &H80000014&
     Height          =   120
     Left            =   330
     Top             =   450
     Width           =   120
   End
   Begin VB.PictureBox T21
     BackColor       =   &H80000015&
     Height          =   120
     Left            =   450
     Top             =   450
     Width           =   120
   End
   Begin VB.PictureBox T22
     BackColor       =   &H80000016&
     Height          =   120
     Left            =   570
     Top             =   450
     Width           =   120
   End
   Begin VB.PictureBox T23
     BackColor       =   &H80000017&
     Height          =   120
     Left            =   690
     Top             =   450
     Width           =   120
   End
   Begin VB.PictureBox T24
     BackColor       =   &H80000018&
     Height          =   120
     Left            =   90
     Top             =   570
     Width           =   120
   End
   Begin VB.PictureBox T25
     BackColor       =   &H80000019&
     Height          =   120
     Left            =   210
     Top             =   570
     Width           =   120
   End
   Begin VB.PictureBox T26
     BackColor       =   &H8000001A&
     Height          =   120
     Left            =   330
     Top             =   570
     Width           =   120
   End
   Begin VB.PictureBox T27
     BackColor       =   &H8000001B&
     Height          =   120
     Left            =   450
     Top             =   570
     Width           =   120
   End
   Begin VB.PictureBox T28
     BackColor       =   &H8000001C&
     Height          =   120
     Left            =   570
     Top             =   570
     Width           =   120
   End
   Begin VB.PictureBox T29
     BackColor       =   &H8000001D&
     Height          =   120
     Left            =   690
     Top             =   570
     Width           =   120
   End
   Begin VB.PictureBox T30
     BackColor       =   &H8000001E&
     Height          =   120
     Left            =   810
     Top             =   90
     Width           =   120
   End
   Begin VB.PictureBox W
     BackColor       =   &H00FFFFFF&
     Height          =   120
     Left            =   900
     Top             =   330
     Width           =   120
   End
   Begin VB.PictureBox K
     BackColor       =   &H00000000&
     Height          =   120
     Left            =   900
     Top             =   450
     Width           =   120
   End
   Begin VB.PictureBox N
     BackColor       =   &HFFFFFF00&
     Height          =   120
     Left            =   900
     Top             =   570
     Width           =   120
   End
   Begin VB.PictureBox Panel
     Height          =   780
     Left            =   90
     Top             =   800
     Width           =   760
     Begin VB.PictureBox C0
       BackColor       =   &H80000000&
       Height          =   120
       Left            =   20
       Top             =   20
       Width           =   120
     End
     Begin VB.PictureBox C1
       BackColor       =   &H80000001&
       Height          =   120
       Left            =   140
       Top             =   20
       Width           =   120
     End
     Begin VB.PictureBox C2
       BackColor       =   &H80000002&
       Height          =   120
       Left            =   260
       Top             =   20
       Width           =   120
     End
     Begin VB.PictureBox C3
       BackColor       =   &H80000003&
       Height          =   120
       Left            =   380
       Top             =   20
       Width           =   120
     End
     Begin VB.PictureBox C4
       BackColor       =   &H80000004&
       Height          =   120
       Left            =   500
       Top             =   20
       Width           =   120
     End
     Begin VB.PictureBox C5
       BackColor       =   &H80000005&
       Height          =   120
       Left            =   620
       Top             =   20
       Width           =   120
     End
     Begin VB.PictureBox C6
       BackColor       =   &H80000006&
       Height          =   120
       Left            =   20
       Top             =   140
       Width           =   120
     End
     Begin VB.PictureBox C7
       BackColor       =   &H80000007&
       Height          =   120
       Left            =   140
       Top             =   140
       Width           =   120
     End
     Begin VB.PictureBox C8
       BackColor       =   &H80000008&
       Height          =   120
       Left            =   260
       Top             =   140
       Width           =   120
     End
     Begin VB.PictureBox C9
       BackColor       =   &H80000009&
       Height          =   120
       Left            =   380
       Top             =   140
       Width           =   120
     End
     Begin VB.PictureBox C10
       BackColor       =   &H8000000A&
       Height          =   120
       Left            =   500
       Top             =   140
       Width           =   120
     End
     Begin VB.PictureBox C11
       BackColor       =   &H8000000B&
       Height          =   120
       Left            =   620
       Top             =   140
       Width           =   120
     End
     Begin VB.PictureBox C12
       BackColor       =   &H8000000C&
       Height          =   120
       Left            =   20
       Top             =   260
       Width           =   120
     End
     Begin VB.PictureBox C13
       BackColor       =   &H8000000D&
       Height          =   120
       Left            =   140
       Top             =   260
       Width           =   120
     End
     Begin VB.PictureBox C14
       BackColor       =   &H8000000E&
       Height          =   120
       Left            =   260
       Top             =   260
       Width           =   120
     End
     Begin VB.PictureBox C15
       BackColor       =   &H8000000F&
       Height          =   120
       Left            =   380
       Top             =   260
       Width           =   120
     End
     Begin VB.PictureBox C16
       BackColor       =   &H80000010&
       Height          =   120
       Left            =   500
       Top             =   260
       Width           =   120
     End
     Begin VB.PictureBox C17
       BackColor       =   &H80000011&
       Height          =   120
       Left            =   620
       Top             =   260
       Width           =   120
     End
     Begin VB.PictureBox C18
       BackColor       =   &H80000012&
       Height          =   120
       Left            =   20
       Top             =   380
       Width           =   120
     End
     Begin VB.PictureBox C19
       BackColor       =   &H80000013&
       Height          =   120
       Left            =   140
       Top             =   380
       Width           =   120
     End
     Begin VB.PictureBox C20
       BackColor       =   &H80000014&
       Height          =   120
       Left            =   260
       Top             =   380
       Width           =   120
     End
     Begin VB.PictureBox C21
       BackColor       =   &H80000015&
       Height          =   120
       Left            =   380
       Top             =   380
       Width           =   120
     End
     Begin VB.PictureBox C22
       BackColor       =   &H80000016&
       Height          =   120
       Left            =   500
       Top             =   380
       Width           =   120
     End
     Begin VB.PictureBox C23
       BackColor       =   &H80000017&
       Height          =   120
       Left            =   620
       Top             =   380
       Width           =   120
     End
     Begin VB.PictureBox C24
       BackColor       =   &H80000018&
       Height          =   120
       Left            =   20
       Top             =   500
       Width           =   120
     End
     Begin VB.PictureBox C25
       BackColor       =   &H80000019&
       Height          =   120
       Left            =   140
       Top             =   500
       Width           =   120
     End
     Begin VB.PictureBox C26
       BackColor       =   &H8000001A&
       Height          =   120
       Left            =   260
       Top             =   500
       Width           =   120
     End
     Begin VB.PictureBox C27
       BackColor       =   &H8000001B&
       Height          =   120
       Left            =   380
       Top             =   500
       Width           =   120
     End
     Begin VB.PictureBox C28
       BackColor       =   &H8000001C&
       Height          =   120
       Left            =   500
       Top             =   500
       Width           =   120
     End
     Begin VB.PictureBox C29
       BackColor       =   &H8000001D&
       Height          =   120
       Left            =   620
       Top             =   500
       Width           =   120
     End
     Begin VB.PictureBox C30
       BackColor       =   &H8000001E&
       Height          =   120
       Left            =   20
       Top             =   620
       Width           =   120
     End
   End
End
Attribute VB_Name = "SysColorForm"
Option Explicit

' 账 <VB 系统色>: 设计期颜色链路的运行时断言夹具。
'   31 枚系统色 (COLOR_SCROLLBAR(0) .. COLOR_MENUBAR(30)) 各走两档创建路:
'     · T0..T30 = 窗体的**顶层控件** (cgen_form_ctrl_style_apply.inc 那条发射路)
'     · C0..C30 = PictureBox Panel 的**容器子控件** (cgen_form_frame_menu.inc 那条发射路)
'   外加非系统色反例: W=&H00FFFFFF& (白) / K=&H00000000& (黑) / N=&HFFFFFF00& (负值但
'   **不是**系统色 —— 低字节 0 若被"只判 n<=30"的解码误翻, 会变成 GetSysColor(0) 灰)。
'   每枚 PictureBox 在 Timer 第一拍 Cls (走 vb6_Form_Cls 的 BackColor 填充), 第二拍后
'   用 GetDC/GetPixel 读中心像素, 与 GetSysColor(idx) 逐条比对。像素写法按 COLORREF:
'   R = px And &HFF (低字节), 但这里只做**等值**比较, 不拆通道。
'   判据: SC-TOP ok=31/31 / SC-CHILD ok=31/31 / SC-NONSYS ok=3/3 —— 旧产物 (裸
'   (COLORREF) 造刷) 在这三条上必红 (系统色读到 RGB(低字节,0,0) 一类残色)。

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private mStep As Integer

Private Sub Form_Load()
    Debug.Print "SC-LOAD"
End Sub

' 系统色: 期望值必须由 GetSysColor(idx) 现答 (主题色不是常量)。
Private Function CheckOne(ByVal h As LongPtr, ByVal idx As Long) As Boolean
    Dim d As LongPtr
    Dim px As Long
    Dim want As Long
    want = GetSysColor(idx)
    d = GetDC(h)
    px = GetPixel(d, 4, 4)
    ReleaseDC h, d
    If px = want Then
        CheckOne = True
    Else
        Debug.Print "SC-MISMATCH idx=" & CStr(idx) & " px=" & CStr(px) & " want=" & CStr(want)
    End If
End Function

' 非系统色反例: 期望值是 COLORREF 字面量, **不**许经过 GetSysColor。
Private Function CheckRaw(ByVal h As LongPtr, ByVal want As Long) As Boolean
    Dim d As LongPtr
    Dim px As Long
    d = GetDC(h)
    px = GetPixel(d, 4, 4)
    ReleaseDC h, d
    If px = want Then
        CheckRaw = True
    Else
        Debug.Print "SC-MISMATCH raw px=" & CStr(px) & " want=" & CStr(want)
    End If
End Function

Private Sub tmrStep_Timer()
    Dim okTop As Integer, okChild As Integer, okNon As Integer
    If mStep = 0 Then
        T0.Cls
        T1.Cls
        T2.Cls
        T3.Cls
        T4.Cls
        T5.Cls
        T6.Cls
        T7.Cls
        T8.Cls
        T9.Cls
        T10.Cls
        T11.Cls
        T12.Cls
        T13.Cls
        T14.Cls
        T15.Cls
        T16.Cls
        T17.Cls
        T18.Cls
        T19.Cls
        T20.Cls
        T21.Cls
        T22.Cls
        T23.Cls
        T24.Cls
        T25.Cls
        T26.Cls
        T27.Cls
        T28.Cls
        T29.Cls
        T30.Cls
        C0.Cls
        C1.Cls
        C2.Cls
        C3.Cls
        C4.Cls
        C5.Cls
        C6.Cls
        C7.Cls
        C8.Cls
        C9.Cls
        C10.Cls
        C11.Cls
        C12.Cls
        C13.Cls
        C14.Cls
        C15.Cls
        C16.Cls
        C17.Cls
        C18.Cls
        C19.Cls
        C20.Cls
        C21.Cls
        C22.Cls
        C23.Cls
        C24.Cls
        C25.Cls
        C26.Cls
        C27.Cls
        C28.Cls
        C29.Cls
        C30.Cls
        W.Cls
        K.Cls
        N.Cls
    ElseIf mStep >= 2 Then
        If CheckOne(T0.hwnd, 0) Then okTop = okTop + 1
        If CheckOne(T1.hwnd, 1) Then okTop = okTop + 1
        If CheckOne(T2.hwnd, 2) Then okTop = okTop + 1
        If CheckOne(T3.hwnd, 3) Then okTop = okTop + 1
        If CheckOne(T4.hwnd, 4) Then okTop = okTop + 1
        If CheckOne(T5.hwnd, 5) Then okTop = okTop + 1
        If CheckOne(T6.hwnd, 6) Then okTop = okTop + 1
        If CheckOne(T7.hwnd, 7) Then okTop = okTop + 1
        If CheckOne(T8.hwnd, 8) Then okTop = okTop + 1
        If CheckOne(T9.hwnd, 9) Then okTop = okTop + 1
        If CheckOne(T10.hwnd, 10) Then okTop = okTop + 1
        If CheckOne(T11.hwnd, 11) Then okTop = okTop + 1
        If CheckOne(T12.hwnd, 12) Then okTop = okTop + 1
        If CheckOne(T13.hwnd, 13) Then okTop = okTop + 1
        If CheckOne(T14.hwnd, 14) Then okTop = okTop + 1
        If CheckOne(T15.hwnd, 15) Then okTop = okTop + 1
        If CheckOne(T16.hwnd, 16) Then okTop = okTop + 1
        If CheckOne(T17.hwnd, 17) Then okTop = okTop + 1
        If CheckOne(T18.hwnd, 18) Then okTop = okTop + 1
        If CheckOne(T19.hwnd, 19) Then okTop = okTop + 1
        If CheckOne(T20.hwnd, 20) Then okTop = okTop + 1
        If CheckOne(T21.hwnd, 21) Then okTop = okTop + 1
        If CheckOne(T22.hwnd, 22) Then okTop = okTop + 1
        If CheckOne(T23.hwnd, 23) Then okTop = okTop + 1
        If CheckOne(T24.hwnd, 24) Then okTop = okTop + 1
        If CheckOne(T25.hwnd, 25) Then okTop = okTop + 1
        If CheckOne(T26.hwnd, 26) Then okTop = okTop + 1
        If CheckOne(T27.hwnd, 27) Then okTop = okTop + 1
        If CheckOne(T28.hwnd, 28) Then okTop = okTop + 1
        If CheckOne(T29.hwnd, 29) Then okTop = okTop + 1
        If CheckOne(T30.hwnd, 30) Then okTop = okTop + 1
        If CheckOne(C0.hwnd, 0) Then okChild = okChild + 1
        If CheckOne(C1.hwnd, 1) Then okChild = okChild + 1
        If CheckOne(C2.hwnd, 2) Then okChild = okChild + 1
        If CheckOne(C3.hwnd, 3) Then okChild = okChild + 1
        If CheckOne(C4.hwnd, 4) Then okChild = okChild + 1
        If CheckOne(C5.hwnd, 5) Then okChild = okChild + 1
        If CheckOne(C6.hwnd, 6) Then okChild = okChild + 1
        If CheckOne(C7.hwnd, 7) Then okChild = okChild + 1
        If CheckOne(C8.hwnd, 8) Then okChild = okChild + 1
        If CheckOne(C9.hwnd, 9) Then okChild = okChild + 1
        If CheckOne(C10.hwnd, 10) Then okChild = okChild + 1
        If CheckOne(C11.hwnd, 11) Then okChild = okChild + 1
        If CheckOne(C12.hwnd, 12) Then okChild = okChild + 1
        If CheckOne(C13.hwnd, 13) Then okChild = okChild + 1
        If CheckOne(C14.hwnd, 14) Then okChild = okChild + 1
        If CheckOne(C15.hwnd, 15) Then okChild = okChild + 1
        If CheckOne(C16.hwnd, 16) Then okChild = okChild + 1
        If CheckOne(C17.hwnd, 17) Then okChild = okChild + 1
        If CheckOne(C18.hwnd, 18) Then okChild = okChild + 1
        If CheckOne(C19.hwnd, 19) Then okChild = okChild + 1
        If CheckOne(C20.hwnd, 20) Then okChild = okChild + 1
        If CheckOne(C21.hwnd, 21) Then okChild = okChild + 1
        If CheckOne(C22.hwnd, 22) Then okChild = okChild + 1
        If CheckOne(C23.hwnd, 23) Then okChild = okChild + 1
        If CheckOne(C24.hwnd, 24) Then okChild = okChild + 1
        If CheckOne(C25.hwnd, 25) Then okChild = okChild + 1
        If CheckOne(C26.hwnd, 26) Then okChild = okChild + 1
        If CheckOne(C27.hwnd, 27) Then okChild = okChild + 1
        If CheckOne(C28.hwnd, 28) Then okChild = okChild + 1
        If CheckOne(C29.hwnd, 29) Then okChild = okChild + 1
        If CheckOne(C30.hwnd, 30) Then okChild = okChild + 1
        If CheckRaw(W.hwnd, &HFFFFFF) Then okNon = okNon + 1
        If CheckRaw(K.hwnd, &H0) Then okNon = okNon + 1
        If CheckRaw(N.hwnd, &HFFFF00) Then okNon = okNon + 1
        Debug.Print "SC-TOP ok=" & CStr(okTop) & "/31"
        Debug.Print "SC-CHILD ok=" & CStr(okChild) & "/31"
        Debug.Print "SC-NONSYS ok=" & CStr(okNon) & "/3"
        tmrStep.Enabled = False
        Debug.Print "SC-DONE"
        Unload Me
    End If
    mStep = mStep + 1
End Sub
