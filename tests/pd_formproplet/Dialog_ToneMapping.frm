VERSION 5.00
Begin VB.Form dialog_ToneMapping 
   Caption         =   "dialog_ToneMapping"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
End
Attribute VB_Name = "dialog_ToneMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> C2065 族① 判据夹子: 与 PhotoDemon `Forms\Dialog_ToneMapping.frm:368` +
' `Modules\DialogManager.bas:356` (`dialog_ToneMapping.fi_HandleCopy = fi_Handle`) 同形。
' 跨窗体**用户自定义** Property Let 不在窗体内建属性表里, 以前落 Module.var 回退
' 发出裸 `fi_HandleCopy = fi_Handle;` (整编 C2065, Dialogs.c:625)。
Private m_Handle As Long

Public Property Let fi_HandleCopy(ByVal srcHandle As Long)
    m_Handle = srcHandle
End Property

Public Sub ShowDialog()
End Sub
