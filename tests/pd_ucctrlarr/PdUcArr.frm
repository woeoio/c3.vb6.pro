VERSION 5.00
Begin VB.Form PdUcArr 
   Caption         =   "PdUcArr"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4680
   StartUpPosition =   3  'Windows Default
   Begin pductrlarr.PdUcSlider slt 
      Height          =   300
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   2400
   End
   Begin pductrlarr.PdUcSlider sltSingle 
      Height          =   300
      Left            =   120
      Top             =   900
      Width           =   2400
   End
   Begin pductrlarr.PdUcSlider slt 
      Height          =   300
      Index           =   1
      Left            =   120
      Top             =   600
      Width           =   2400
   End
End
Attribute VB_Name = "PdUcArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' <vbeclipse> F1 判据夹子: PhotoDemon FormLevels.frm 与 dialog_ToneMapping.frm 的同形 ——
' 工程内 UC 的**控件数组元素**在值上下文里就是该 UC 的默认成员
' (写 = Property Let, 读 = Property Get)。旧发码把它当普通数组元素:
' `vb6_CtrlArr_GetAt(&vb6_arr_slt, 0) = 2.2` (C2440: double → void *),
' 读侧同样 C2440。
' 注意: 本注释不许写出默认成员名, 否则会被 .ctl 侧的文本扫描误认 (见 PdUcSlider.ctl)。

Private m_LastValue As Double

Private Sub cmdReset_Click()
    slt(0) = 2.2
    slt(1) = 1#
End Sub

Public Function ReadSecond() As Double
    ReadSecond = slt(1)
    m_LastValue = ReadSecond
End Function

Public Sub SetByIndex(ByVal idx As Long, ByVal newValue As Double)
    slt(idx) = newValue
    m_LastValue = sltSingle
End Sub

' 对象上下文负控: Set 的右值要的是控件本身, 不许折成默认成员标量。
Public Function TakeObject() As Object
    Dim o As Object
    Set o = slt(0)
    Set TakeObject = o
End Function
