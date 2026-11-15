VERSION 5.00
Begin VB.Form MaskForm 
   Caption         =   "MaskForm"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   LinkTopic       =   "MaskForm"
   ScaleHeight     =   3600
   ScaleWidth      =   5400
   Begin MSMask.MaskEdBox mDate 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2000
      Mask            =   "00/00/0000"
      PromptChar      =   "_"
      _ExtentX        =   3528
      _ExtentY        =   556
   End
   Begin MSMask.MaskEdBox mPhone 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2400
      Mask            =   "(999) 000-0000"
      _ExtentX        =   4233
      _ExtentY        =   556
   End
   Begin MSMask.MaskEdBox mUpper 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2400
      Mask            =   ">LLL-000000"
      _ExtentX        =   4233
      _ExtentY        =   556
   End
   Begin MSMask.MaskEdBox mFree 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1200
      Mask            =   "00"
      _ExtentX        =   2117
      _ExtentY        =   556
   End
   Begin VB.Timer tGo 
      Interval        =   60
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "MaskForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' C30-3: MaskEdBox 的无头判据（与 tests/ctrltabstrip 同构：Timer 起、Debug.Print 出）。
' 每条判据只认 ASCII 数字/字母 —— GBK 控制台下中文 needle 在别的 locale 会红。
Option Explicit

' 14) Change 事件面的证人：MaskEdBox 是 EDIT 类窗口，它的 EN_CHANGE（code=768）走
'    父窗的 WM_COMMAND —— 与 TextBox 同一条来源。mDate 被赋值多次，计数必然 > 0；
'    若 cgen 的张/派两处表里漏了 MaskEdBox，这里会是 False（事件一条不到）。
Private gChg As Long

Private Sub mDate_Change()
    gChg = gChg + 1
End Sub

Private Sub tGo_Timer()
    ' 1) 设计期掩码 + Text 写入：Text 带固定符、ClipText 只含用户字符、Mask 原样回读
    mDate.Text = "31122025"
    Debug.Print "MK1-date=" & mDate.Text & "|" & mDate.ClipText & "|" & mDate.Mask

    ' 2) 非法字符被丢弃（字母进不了数字位），余位仍是提示符
    mDate.Text = "ab1234"
    Debug.Print "MK2-bad=" & mDate.Text & "|" & mDate.ClipText

    ' 3) 电话掩码：可选位 + 固定符号（括号 / 空格 / 连字符）
    mPhone.ClipText = "0101234567"
    Debug.Print "MK3-phone=" & mPhone.Text

    ' 4) ">" 起头：其后输入位强制大写
    mUpper.ClipText = "abc123456"
    Debug.Print "MK4-upper=" & mUpper.Text

    ' 5) "<" 起头：其后输入位强制小写
    mUpper.Mask = "<LL"
    mUpper.ClipText = "AB"
    Debug.Print "MK5-lower=" & mUpper.Text

    ' 6) "?" 可选字母：数字落不进这一位，让给后面的 "0"
    mUpper.Mask = "?0"
    mUpper.ClipText = "5"
    Debug.Print "MK6-opt=" & mUpper.Text

    ' 7) "#" 接受正负号
    mUpper.Mask = "##"
    mUpper.ClipText = "+7"
    Debug.Print "MK7-sign=" & mUpper.Text

    ' 8) "\" 转义：后面的 A 当固定符号
    mUpper.Mask = "\A0"
    mUpper.ClipText = "7"
    Debug.Print "MK8-esc=" & mUpper.Text

    ' 9) "&" 任意字符必填（固定符号只在掩码里，当数据时照收）
    mUpper.Mask = "&&"
    mUpper.ClipText = "a#"
    Debug.Print "MK9-any=" & mUpper.Text

    ' 10) PromptChar：空位提示符跟着换
    mFree.PromptChar = "*"
    mFree.Mask = "00"
    Debug.Print "MK10-prompt=" & mFree.Text

    ' 11) PromptInclude=False：Text 不含提示符（长度跟着变）
    mFree.PromptInclude = False
    Debug.Print "MK11-noincl=" & mFree.Text & "|" & CStr(Len(mFree.Text))

    ' 12) ClipText 赋值回灌：Text 自动补上固定符号
    mDate.ClipText = "01012026"
    Debug.Print "MK12-clip=" & mDate.Text

    ' 13) Sel*：显示坐标（固定符也占位）
    mDate.Text = "31122025"
    mDate.SelStart = 0
    mDate.SelLength = 2
    Debug.Print "MK13-sel=" & mDate.SelText & "|" & CStr(mDate.SelStart)

    Debug.Print "MK14-chg=" & CStr(gChg > 0)
    Debug.Print "MASKED-DONE"
    Unload Me
End Sub
