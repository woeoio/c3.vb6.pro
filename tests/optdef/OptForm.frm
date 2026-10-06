VERSION 5.00
Begin VB.Form OptForm
   Caption         =   "OptDef"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   LinkTopic       =   "OptForm"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "OptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 账 #194 + #195: VB 的整数类型后缀 (`12&` / `3%` / `&H10&`) 属于**词法**，既不该出现在生成 C 里,
' 语义层那份 Optional 默认值求值以前直接 return rawText, 于是真工程 (Charts 2020/ucTreeMaps
' 的 PropPagFMR.pag:740/886 `Optional ... = 0&`) 发成 `(*FontIndex) = 0&;` = C2059。
' 判据两头: ①四个默认值各按声明落地 (十六进制那枚同时证明"按数值重打"没把值改错);
' ②显式实参照样赢 —— 只钉前头那条的话, "恒取默认值"也是绿的。
Dim gL As Long
Dim gI As Integer
Dim gH As Long
Dim gZ As Long

Private Sub Take(Optional ByVal aL As Long = 12&, Optional ByVal aI As Integer = 3%, _
                 Optional ByVal aH As Long = &H10&, Optional ByVal aZ As Long = 0&)
    gL = aL
    gI = aI
    gH = aH
    gZ = aZ
End Sub

Private Sub Form_Load()
    Take
    Debug.Print "OD-RAW l=" & CStr(gL) & " i=" & CStr(gI) & " h=" & CStr(gH) & " z=" & CStr(gZ)
    Debug.Print "OD-VAL=" & CStr(gL = 12 And gI = 3 And gH = 16 And gZ = 0)
    gL = 0
    gI = 0
    gH = 0
    gZ = 0
    Take 1&, 5%, &H20&, 9&
    Debug.Print "OD-SET l=" & CStr(gL) & " i=" & CStr(gI) & " h=" & CStr(gH) & " z=" & CStr(gZ)
    Debug.Print "OD-EXPL=" & CStr(gL = 1 And gI = 5 And gH = 32 And gZ = 9)
    Debug.Print "OD-DONE"
    Unload Me
End Sub
