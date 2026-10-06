VERSION 5.00
Begin VB.Form SmForm 
   Caption         =   "ScaleMode"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5000
   ScaleMode       =   3  'Pixel
   Begin VB.PictureBox picPix 
      Height          =   1200
      Left            =   120
      ScaleMode       =   3  'Pixel
      Top             =   240
      Width           =   1200
      Begin VB.CommandButton cmdPix 
         Caption         =   "P"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   600
      End
   End
   Begin VB.PictureBox picTwip 
      Height          =   1200
      Left            =   1560
      ScaleMode       =   1  'Twips
      Top             =   240
      Width           =   1200
      Begin VB.CommandButton cmdTwip 
         Caption         =   "T"
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   600
      End
   End
   Begin VB.Frame frN 
      Caption         =   "frN"
      Height          =   1400
      Left            =   120
      TabIndex        =   4
      Top             =   1900
      Width           =   1500
      Begin VB.PictureBox picNest 
         Height          =   600
         Left            =   120
         ScaleMode       =   3  'Pixel
         TabIndex        =   5
         Top             =   360
         Width           =   1200
         Begin VB.CommandButton cmdNest 
            Caption         =   "N"
            Height          =   240
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   480
         End
      End
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "out"
      Height          =   300
      Left            =   3120
      TabIndex        =   3
      Top             =   2000
      Width           =   900
   End
End
Attribute VB_Name = "SmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 账 #197: `VB6_ScaleMode` 这个窗口属性以前**全仓没有一个写者** —— 读的一侧早就齐了
' (vb6_GetScaleMode / vb6_WindowScaleModeSelf / vb6_ContainerScaleMode，缺省 1=缇)，
' 于是谁问 ScaleMode 都答 1；而 ScaleWidth、控件几何、文字量纲全按这一档折算 ⇒
' 这一档错不在于某一枚控件长歪，而在于**量出来的数全是错的**。
' 两头判据，缺一头就是假绿（同 #153/#154 那条口径：存请求值 + 一枚问窗口的证人）：
'   SM01 存请求值 —— 窗体、像素框、缇框三处设计期声明各自读回自己那一个数（修前三枚全答 1）。
'   SM02/SM03 问窗口 —— 两枚同尺寸的框：缇档量出的 ScaleWidth 必须比像素档大一个单位比；
'        框里那枚按钮的 Left 同理（子控件的单位 = **所在容器**的 ScaleMode，账 #175 那条）。
'        没人写、或者写成全局一份，这两条当场红：两枚会答出同一个数。比值只取 >3，不钉绝对值。
'   SM04 运行期换档 —— 像素框切到缇档后与另一枚**逐数相等**（不靠 DPI）。
'   SM05 另一枚没被写坏（RTL 按 HWND 存窗口属性，同 #192 那条双向钉）。
'   SM06/SM07 **第二条创建路**：Frame 里那枚 PictureBox 也声明了 3=Pixel —— 顶层与容器两条路
'        必须发同一句（账 #83/#151 那条"两条创建路都要打"），只接一头这里就红。
' 窗体与框都按 VB6 的真形状测：读用 `Me.ScaleMode` / `picPix.ScaleMode` / With 里那一形，
' 写用 With 块里的 `.ScaleMode = 1`（真工程 Charts 2020/ucTreeMaps 的 PropPagFMR.pag:259 就这么写）。
Private Sub Form_Load()
    Dim mForm As Long, mPix As Long, mTwip As Long, mNest As Long
    Dim wPix As Long, wTwip As Long
    Dim lPix As Long, lTwip As Long, lNest As Long

    With picPix
        mPix = .ScaleMode
    End With
    mForm = Me.ScaleMode
    mTwip = picTwip.ScaleMode
    mNest = picNest.ScaleMode
    wPix = picPix.ScaleWidth
    wTwip = picTwip.ScaleWidth
    lPix = cmdPix.Left
    lTwip = cmdTwip.Left
    lNest = cmdNest.Left

    Debug.Print "SM01-MODE=" & mForm & "/" & mPix & "/" & mTwip
    Debug.Print "SM02-UNIT=" & TF(wPix > 0 And wTwip > 3 * wPix)
    Debug.Print "SM03-CHILD=" & TF(lPix > 0 And lTwip > 3 * lPix)

    With picPix
        .ScaleMode = 1
    End With
    Debug.Print "SM04-SWITCH=" & TF(picPix.ScaleWidth = wTwip And cmdPix.Left = lTwip)
    Debug.Print "SM05-OTHER=" & TF(mTwip = 1 And picTwip.ScaleMode = 1)
    Debug.Print "SM06-NEST=" & TF(mNest = 3)
    Debug.Print "SM07-NESTCHILD=" & TF(lNest > 0 And lTwip > 3 * lNest)
    Debug.Print "SM08-RAW wPix=" & wPix & " wTwip=" & wTwip & " lPix=" & lPix _
        & " lTwip=" & lTwip & " lNest=" & lNest
    Debug.Print "SM-DONE"
    Unload Me
End Sub

Private Function TF(b As Boolean) As String
    If b Then TF = "True" Else TF = "False"
End Function
