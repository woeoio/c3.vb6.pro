VERSION 5.00
Begin VB.Form FormB 
   Caption         =   "FormB"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
End
Attribute VB_Name = "FormB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' C2224 复现 (PhotoDemon 整编 16 条): 接收者是**窗体值表达式**时, 成员读必须走
' 窗体属性表 (vb6_GetControlHwnd 一族), 不能拼 `call().hWnd`。
Private Function GetModalOwner() As FormA
    Set GetModalOwner = FormA
End Function

Private Function GetOwnerForm() As Form
    Set GetOwnerForm = FormA
End Function

Private Sub Form_Load()
    Dim h As Long
    h = GetModalOwner().hWnd
    h = GetOwnerForm().hWnd
    h = FormA.hWnd
    h = Forms(0&).hWnd
    Debug.Print h
End Sub
