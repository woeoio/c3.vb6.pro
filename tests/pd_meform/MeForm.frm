VERSION 5.00
Begin VB.Form MeForm 
   Caption         =   "MeForm"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Begin PdMeForm.pdFxPreviewCtl pdFxPreview 
      Height          =   1200
      Left            =   120
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "MeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> C2065 族① 判据夹子: 窗体自有过程用 `Me.` 显式限定
' (PhotoDemon 十余个效果窗体: `Me.ApplySepiaEffect GetLocalParamString(), True, pdFxPreview`)。
' 改前接收者走「未知 Form 成员 → COM dispatch」: 过程名被当属性去问窗体 HWND, 晚绑定
' 打包又把设计期 UC 控件名裸发成 C 标识符 ⇒ C2065 'pdFxPreview' (整编实测 39 条)。
' 正确形态 = 与裸名调用同一出口 (vb6_MeForm_ApplySepiaEffect), 实参走 §42 的 UC 站。

Public Function GetLocalParamString() As String
    GetLocalParamString = "pdm"
End Function

Public Sub ApplySepiaEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PdFxPreviewCtl)
End Sub

Public Sub RunMe()
    Me.ApplySepiaEffect GetLocalParamString(), True, pdFxPreview
End Sub
