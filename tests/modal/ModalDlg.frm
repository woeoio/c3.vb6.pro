VERSION 5.00
Begin VB.Form ModalDlg 
   Caption         =   "ModalDlg"
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3600
   LinkTopic       =   "ModalDlg"
   ScaleHeight     =   1800
   ScaleWidth      =   3600
   Begin VB.CommandButton cmdB 
      Caption         =   "B"
      Height          =   400
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   900
   End
   Begin VB.TextBox txtHidden 
      Enabled         =   -1   'True
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "Off"
      Height          =   400
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      TabStop         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblHead 
      Caption         =   "head"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDis 
      Caption         =   "Dis"
      Enabled         =   0   'False
      Height          =   400
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   900
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A"
      Height          =   400
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   900
   End
   Begin VB.Timer t2 
      Enabled         =   -1   'True
      Interval        =   40
      Left            =   3120
      Top             =   240
   End
End
Attribute VB_Name = "ModalDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 账 #157 的模态那一半：这枚窗体把 VB6 那条"焦点交给 TabIndex 最小那枚**拿得到焦点的**控件"
' 的四条排除一次问完，并且让**赢家排在创建顺序的最后一位**（.frm 里 cmdA 是最后声明的那个，
' 而 TabIndex=4 既不是最小也不是最大）—— 这样"照创建顺序挑"与"照 TabIndex 挑"给出不同答案。
'   TabIndex 0 = lblHead    Label，拿不到焦点（按类型排除）
'   TabIndex 1 = cmdOff     显式写了 TabStop=False（按那枚属性排除）
'   TabIndex 2 = cmdDis     设计期 Enabled=False（禁用的拿不到焦点）
'   TabIndex 3 = txtHidden  设计期 Visible=False（藏起来的拿不到焦点）
'   TabIndex 4 = cmdA       ← 该赢的就是它
'   TabIndex 5 = cmdB       只当"别挑成第一枚创建的"的证人（它是 .frm 里第一个声明的）
' D1 是本条会翻红的那枚（改之前焦点停在窗体自己身上 ⇒ N）；D2..D6 是防伪证：把选择改成
' "第一枚创建的""最后一枚创建的""随机一枚"都会让其中某枚变红。
' 读数全走 LongPtr 形参（HexEq/NotEq），不碰控件 .hwnd 的装箱（账 #159 那条路是坏的）。
Private gState As Long

Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

Private Const WM_CLOSE As Long = &H10

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Function HexEq(ByVal a As LongPtr, ByVal b As LongPtr) As String
    If a = b Then
        HexEq = "Y"
    Else
        HexEq = "N(" & CStr(a) & "<>" & CStr(b) & ")"
    End If
End Function

Private Function NotEq(ByVal a As LongPtr, ByVal b As LongPtr) As String
    If a = b Then NotEq = "N" Else NotEq = "Y"
End Function

Private Sub Log1(ByVal s As String)
    Dim h As Long
    h = FreeFile
    Open App.Path & "\modal.log" For Append As #h
    Print #h, s
    Close #h
    Debug.Print s
End Sub

Private Sub t2_Timer()
    Dim f As LongPtr
    gState = gState + 1
    If gState = 1 Then
        f = GetFocus()
        Log1 "D1-first=" & HexEq(f, cmdA.hwnd)
        Log1 "D2-notB=" & NotEq(f, cmdB.hwnd)
        Log1 "D3-notLbl=" & NotEq(f, lblHead.hwnd)
        Log1 "D4-notOff=" & NotEq(f, cmdOff.hwnd)
        Log1 "D5-notDis=" & NotEq(f, cmdDis.hwnd)
        Log1 "D6-notHidden=" & NotEq(f, txtHidden.hwnd)
    End If
    If gState >= 4 Then
        Log1 "D7-ticks=" & TF(gState >= 4) & "/" & CStr(gState)
        Log1 "DL-closing"
        Call PostMessage(GetParent(cmdA.hwnd), WM_CLOSE, 0, 0)
    End If
End Sub
