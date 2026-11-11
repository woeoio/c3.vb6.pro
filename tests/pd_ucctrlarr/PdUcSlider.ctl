VERSION 5.00
Begin VB.UserControl PdUcSlider 
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
End
Attribute VB_Name = "PdUcSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> F1 判据夹子: 与 PhotoDemon `Controls\pdSlider.ctl` 同形 —— 工程内
' UserControl 的默认成员声明在 .ctl 代码段 (DISPID 0 + 同名 Get 返回 Double)。
' FormLevels/dialog_ToneMapping 的控件数组元素 `sltGamma(0) = 2.2` /
' `x = tudLevels(Index)` 就是这枚默认成员的写腿与读腿。
' 注: 下面这行注释故意写出声明串 —— 钉住 findUcDefaultMember 只认非注释行、
' 且只从属性行之后认声明（旧写法从 0 起搜全文, 会把这句注释当真正的声明,
' 型取到注释尾巴 ⇒ 静默不折 ⇒ C2440 原样回来）。
' 被钉住的串: Public Property Get Value() As Double

Private m_Value As Double

Public Property Get Value() As Double
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal newValue As Double)
    m_Value = newValue
End Property
