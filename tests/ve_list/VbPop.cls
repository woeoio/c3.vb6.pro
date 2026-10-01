VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Fix <vbeclipse> rev17 回归夹具: 真工程 VbEclipse/Classes/PopupMenu.cls 的最小形态 ——
' **Sub 与类同名** (裸键真冲突, 区别于 VbSelf 的 Property 形态)。
' PopupMenu.cls: `Public Sub PopupMenu(ByVal hWnd As Long)`。

Private m_Hits As Long

' 与类同名 —— Sub 形态 (PopupMenu.cls)。裸键 <cls> 必须让给这个成员,
' 但类符号自己也不能消失 (否则 coclass 表缺项 → 运行期 FindCoClassDesc 返 NULL)。
Public Sub VbPop()
    m_Hits = m_Hits + 1
End Sub

Public Property Get Hits() As Long
    Hits = m_Hits
End Property
