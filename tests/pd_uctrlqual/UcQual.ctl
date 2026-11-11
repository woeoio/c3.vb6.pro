VERSION 5.00
Begin VB.UserControl UcQual 
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Begin VB.ComboBox cmbWidthUnit 
      Height          =   300
      Left            =   120
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "UcQual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> C2065 判据夹子: 与 PhotoDemon `Controls\pdResize.ctl:733`
' (`UserControl.cmbWidthUnit.ListIndex = mu_Centimeters`) 同形。
' `UserControl.<本 UC 的设计期控件>` 与 `Me.<控件>` 同义; 旧发码把它当宿主
' 伪成员拼成 `vb6_UserControl_cmbWidthUnit` (没人声明 ⇒ C2065 ×4)。
' 本题只钉**限定符解析**这一层: 接收者必须是控件的 hwnd 槽; 其后成员的
' 发码形状 (COM 还是结构体) 由控件类型另定, 不在本钉范围。
Public Function CurrentMetricIndex() As Long
    CurrentMetricIndex = UserControl.cmbWidthUnit.ListIndex
End Function

' 写侧同形 (pdResize.ctl:733/734/736/737 `UserControl.cmbWidthUnit.ListIndex = ...`):
' 只把内层接收者折成控件槽还不够 —— 外层若走 generic 会发成
' `vb6_hwnd_cmbWidthUnit.ListIndex = 3` (对 void* 取成员, C2224)。
Public Sub SetMetricIndex(ByVal newIndex As Long)
    UserControl.cmbWidthUnit.ListIndex = newIndex
End Sub

' pdResize.ctl:301 `Not UserControl.Controls(ctlChild.Name) Is ctlChild` 的同形:
' 显式限定 `UserControl.Controls(...)` = 集合默认成员 Item, 返回 Object。
' 旧发码把它当宿主伪函数, 先拼未声明的 `vb6_UserControl_Controls` 再追加实参
' ⇒ cl C2063 "not a function"。
Public Function IsChildControl(ByRef ctlChild As Object) As Boolean
    IsChildControl = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function
