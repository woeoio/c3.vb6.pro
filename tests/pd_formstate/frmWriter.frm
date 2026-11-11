VERSION 5.00
Begin VB.Form frmWriter 
   Caption         =   "frmWriter"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
End
Attribute VB_Name = "frmWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> C2065 判据夹子: 与 PhotoDemon 的 `FormMain.WindowState = vbMinimized`
' (Tools_ScreenVideo.frm:947 / Tools_ScreenVideoPrefs.frm:246) 同形。
' 读侧 form_pseudo 表早就有 windowstate; 写侧没登记 ⇒ 落 Module.var 回退发裸
' `WindowState = ...` (C2065)。
Private Sub Form_Load()
    frmOther.WindowState = vbMaximized
End Sub
