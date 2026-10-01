VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbSelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Fix <vbeclipse> rev17 回归夹具: 真工程 VbEclipse/Classes/View.cls 的最小形态 ——
' **成员与类同名**。
' View.cls:  Public Property Get View() As Object / Public Property Set View(...)
' 该形态会让语义层把"类自符号"从模块作用域里挤掉, 而这个类就再也进不了 coclass 表:
'   运行期 vb6_FindCoClassDesc("View") = NULL
'   → vb6_ComObject_FromInstance(NULL, inst) = NULL
'   → vb6_VariantObject(NULL) 产出 vt=VT_DISPATCH / pdispVal=NULL 的 Variant
'   → 槽位读回 `Is Nothing` 恒真 → 调用方拿到 NULL → 解引用崩。
' 真工程 play78 实测: coclass 表 14 项, 缺的正是**全工程唯一两个有同名成员的类**
' View / PopupMenu (PopupMenu.cls: `Public Sub PopupMenu(...)`)。

Private m_Id As String

' 与类同名 —— Property 形态 (View.cls)
Public Property Get VbSelf() As String
    VbSelf = m_Id
End Property

Public Property Let VbSelf(ByVal NewId As String)
    m_Id = NewId
End Property
