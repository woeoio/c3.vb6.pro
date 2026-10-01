VERSION 5.00
Begin VB.Form ModalMain 
   Caption         =   "ModalMain"
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3600
   LinkTopic       =   "ModalMain"
   ScaleHeight     =   1800
   ScaleWidth      =   3600
   Begin VB.CommandButton cmdY 
      Caption         =   "Y"
      Height          =   400
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   900
   End
   Begin VB.CommandButton cmdX 
      Caption         =   "X"
      Height          =   400
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   900
   End
   Begin VB.TextBox txtSecond 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtMain 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblMain 
      Caption         =   "main"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer tMain 
      Enabled         =   -1   'True
      Interval        =   40
      Left            =   2880
      Top             =   240
   End
End
Attribute VB_Name = "ModalMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 账 #157 的夹具（判据建在"事件自己驱动 + LongPtr 形参比对"这两条已证的机器上）：
' VB6 把窗体显示出来时，焦点交给**这枚窗体里 TabIndex 最小的那枚拿得到焦点的控件** ——
' 不是创建顺序、也不是"停在窗体自己身上"。这里问两件事：
'   M1 = 普通（非模态、启动）窗体显示后，焦点是不是落在 txtMain（TabIndex=1）而不是
'        lblMain（TabIndex=0，Label 拿不到焦点）；
'   D* = 模态窗体里那一整套排除（见 ModalDlg）。
' 比对一律走 HexEq/N 这两个助手（`ByVal ... As LongPtr` 形参）—— 控件的 `.hwnd` 直接
' 装箱那条路是坏的（账 #159：CStr(ctl.hwnd) 打空串、与 GetFocus() 比较恒假），
' 拿它当判据就是又一次判据自伤。
Private gState As Long
Private gA As Long
Private gB As Long
Private gSeen As String
Private gNew As Long
Private gRepeat As String
Private gHops As Long
Private gBusy As Boolean
Private gStray As Long

Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

Private Const WM_CLOSE As Long = &H10
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_TAB As Long = &H9

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Function N(ByVal v As LongPtr) As String
    N = CStr(v)
End Function

Private Function HexEq(ByVal a As LongPtr, ByVal b As LongPtr) As String
    If a = b Then
        HexEq = "Y"
    Else
        HexEq = "N(" & CStr(a) & "<>" & CStr(b) & ")"
    End If
End Function

Private Sub Log1(ByVal s As String)
    Dim h As Long
    h = FreeFile
    Open App.Path & "\modal.log" For Append As #h
    Print #h, s
    Close #h
    Debug.Print s
End Sub

' 站点名（不是"是不是某一枚"的两值判据）—— 这样一相当前面那两枚已经过的站点对上了，
' "跳错地方"与"根本没跳"分得开。比对全走 LongPtr 形参（账 #159：控件 .hwnd 的装箱坏）。
Private Function WhereIs(ByVal f As LongPtr) As String
    If HexEq(f, txtMain.hwnd) = "Y" Then
        WhereIs = "txtMain"
    ElseIf HexEq(f, txtSecond.hwnd) = "Y" Then
        WhereIs = "txtSecond"
    ElseIf HexEq(f, cmdX.hwnd) = "Y" Then
        WhereIs = "cmdX"
    ElseIf HexEq(f, cmdY.hwnd) = "Y" Then
        WhereIs = "cmdY"
    ElseIf HexEq(f, Me.hwnd) = "Y" Then
        WhereIs = "form"
    Else
        WhereIs = "other(" & CStr(f) & ")"
    End If
End Function

' 两相：tick 1 = 启动窗体的初始焦点 + 开模态（账 #157 / #156 那两格）；
' 之后各拍 = **主泵**的 Tab 导航（账 #83(b)）：读当前落点、再往它 post 一对 VK_TAB。
' ⚠ 顺序**现在钉了**（账 #163 由自研导航器接走之后）：`MW-seq` 钉的是**首次访问的次序**，
' 不是逐拍读数 —— 某一拍没动不会假红（与 WS17 / 账 #162 那一族「条数是时序不是不变量」分开）。
' VB6 的次序 = 父窗内 `TabIndex`：txtMain(1) → txtSecond(2) → cmdX(3) → cmdY(4) → 回 txtMain；
' 交给 `IsDialogMessage` 时走的是 z-order（`txtMain → cmdY → cmdX → txtSecond`）。
' ⇒ 负控有两条开关，各红各的：`C3_OCX_NO_TABNAV=1` 退回 z-order（红 `MW-seq`），
' `C3_OCX_NO_DLGMSG=1` 关掉泵里的 `IsDialogMessage`（红 `MW-new` ⇒ 回到 1）。

Private Sub tMain_Timer()
    Dim f As LongPtr
    Dim st As String
    ' 第一道闸：`Enabled = False` 之后**还会来在途的 tick**（"最多一枚"那个界是看负载的，账 #162），
    ' 而这些拍是在**模态循环里**被排空的 —— 不挡住，它们就会拿着"下一相"的身份往模态窗体
    ' 里 post VK_TAB（x86 实测照出来：模态窗体的初始焦点被挪走，D1/D2 一起红）。
    If gBusy Then
        gStray = gStray + 1
        Exit Sub
    End If
    gState = gState + 1
    If gState = 1 Then
        gBusy = True
        Log1 "M1-startup=" & HexEq(GetFocus(), txtMain.hwnd) & "/second=" & HexEq(GetFocus(), txtSecond.hwnd)
        tMain.Enabled = False
        ModalDlg.Show vbModal
        gBusy = False
        Log1 "M2-returned=Y/ticks=" & CStr(gState) & "/busy=" & CStr(gStray)
        txtMain.SetFocus
        gSeen = "txtMain,"
        tMain.Enabled = True
        Exit Sub
    End If
    If gState >= 2 And gState <= 9 Then
        f = GetFocus()
        st = WhereIs(f)
        ' tick 2 那一拍读到的就是起点（刚 SetFocus 过），不算"回头"，所以从 tick 3 起才记账
        If gState > 2 Then
            Log1 "MW" & CStr(gState - 2) & "=" & st
            If st <> "other" And st <> "form" Then
                If InStr(gSeen, st & ",") = 0 Then
                    gSeen = gSeen & st & ","
                    gNew = gNew + 1
                ElseIf gRepeat = "" Then
                    gRepeat = st
                    gHops = gNew
                End If
            End If
        End If
        Call PostMessage(f, WM_KEYDOWN, VK_TAB, 0)
        Call PostMessage(f, WM_KEYUP, VK_TAB, 0)
    End If
    If gState = 10 Then
        ' 走到底：起点 + 新访的枚数 + 第一次回头落在谁身上 + 回头前走过几格
        Log1 "MW-new=" & CStr(gNew + 1) & "/repeat=" & gRepeat & "/hops=" & CStr(gHops)
        Log1 "MW-seq=" & gSeen
        Log1 "MODAL-DONE"
        Unload Me
    End If
End Sub
