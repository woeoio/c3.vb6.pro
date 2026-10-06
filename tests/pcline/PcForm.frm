VERSION 5.00
Begin VB.Form PcForm 
   Caption         =   "PcLine"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   2400
   ScaleMode       =   1  'Twips
   ScaleWidth      =   4800
   Begin VB.PictureBox picA 
      Height          =   1200
      Left            =   120
      ScaleHeight     =   1140
      ScaleWidth      =   2700
      TabIndex        =   2
      Top             =   120
      Width           =   2820
   End
End
Attribute VB_Name = "PcForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 账 #220: VB6 的 `Picture.Line (x1,y1)-(x2,y2)[, color][, B|BF]` 里, B / BF 是**语法旗标**
' 而不是名字, 且只有 style 那一格 (color 之后) 才认。以前 parser 把这两个词原样交给发码,
' 靠 RTL 的两枚裸名 C 全局接住 ⇒ 任何工程有个模块级变量叫 B 就撞车 (test_nameclash.bas 钉
' 那一头)。这一头钉的是折的位置: 同一个 picA 上四条 Line, 后两条把用户自己的 B / BF 写在
' **color** 位置 —— 那两格绝不能折成 1 / 2, 否则就是"修一处撞车、制造一处静默错值"。
' 所以 needles 两头钉: style 位给字面量 (PL-B-STYLE / PL-BF-STYLE), color 位给变量名本身。
Private B As Long
Private BF As Long

Private Sub Form_Load()
    B = 5
    BF = 6
    picA.Line (10, 10)-(50, 50), vbRed, BF
    picA.Line (0, 0)-(20, 20), vbBlue, B
    picA.Line (5, 5)-(15, 15), B
    picA.Line (6, 6)-(16, 16), BF
    Unload Me
End Sub
