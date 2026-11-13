VERSION 5.00
Begin VB.Form FrmMain
   Caption         =   "F4"
   ClientHeight    =   3000
   ClientWidth     =   4000
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 富贵竹 FIX-004 (账 #610 第340刀): 模块级裸 Dim = 模块私有 (与 Private 等价)。
' 与 ModG 的 Public gShared 同名 —— 修前把真全局连坐成 static 且无 extern;
' 修后本副本走影子名 vb6_shadow_gShared__FrmMain + #define 宏换绑。
Dim gShared As Long

Private Sub Form_Load()
    ModG.SetG 41
    gShared = 99
    LocalShadow
    Debug.Print "F4-GLOBAL=" & ModU.GetG()
    Debug.Print "F4-SHADOW=" & gShared
    ModQ.SetDuo 7
    Debug.Print "F4-DUO=" & ModQ.GetDuo()
    Unload Me
End Sub

' 局部同名 Dim: 块作用域遮蔽影子名 (C 合法), 且不许 #undef 掉影子宏 ——
' 否则本 TU 之后的模块级裸名引用失去换绑 → C2065。
Private Sub LocalShadow()
    Dim gShared As Long
    gShared = 5
    Debug.Print "F4-LOCAL=" & gShared
End Sub
