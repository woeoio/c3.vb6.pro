VERSION 5.00
Begin VB.Form PicForm 
   Caption         =   "PicCur"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   3000
   ScaleMode       =   1  'Twips
   ScaleWidth      =   6000
   Begin VB.PictureBox picA 
      Height          =   1200
      Left            =   120
      ScaleHeight     =   1140
      ScaleWidth      =   2700
      TabIndex        =   2
      Top             =   120
      Width           =   2820
   End
   Begin VB.PictureBox picB 
      Height          =   1200
      Left            =   3000
      ScaleHeight     =   1140
      ScaleWidth      =   2700
      TabIndex        =   3
      Top             =   120
      Width           =   2820
   End
   Begin VB.ListBox lst 
      Height          =   600
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2820
   End
End
Attribute VB_Name = "PicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 账 #192: With 块里读写的必须是**这一枚控件自己的**状态。真工程那一形是
' Charts 2020/ucTreeMaps 的 PropPagFMR.pag:256-266 `With Picture1 : .hDC / .CurrentX /
' .CurrentY / .ScaleMode / .TextHeight` 与 pag:618 `With lstFonts : Call .Clear` ——
' 后端两张表里 CurrentX/CurrentY 从没登记、零实参方法表里没有 clear，于是全部落到
' cgen_expr_with.cpp 那条 tempVar + "." + 成员名 的兜底 = C2039（"CurrentX" 不是 "HWND__" 的成员）。
' 三头判据：PC01 写了读得回（存什么读什么，走的是同一张表的两边）；PC02 **另一枚没被写坏**
' （RTL 按 HWND 存窗口属性 ⇒ 全局一份的话这里就露馅，同 #156 那条"双向钉"的口径）；
' PC03 With 形式的方法调用真的清空了（AddItem 2 条 → Clear → 0 条，且**清之前**那个数是 2，
' 拦住"Clear 其实什么都不做、ListCount 本来就 0"那种假绿）。
Private Sub Form_Load()
    Dim ok1 As Boolean, ok2 As Boolean, ok3 As Boolean
    Dim before As Long

    With picA
        .CurrentX = 3.5
        .CurrentY = 4.25
    End With
    ok1 = (CLng(picA.CurrentX * 100) = 350) And (CLng(picA.CurrentY * 100) = 425)
    ok2 = (CLng(picB.CurrentX * 100) = 0) And (CLng(picB.CurrentY * 100) = 0)

    lst.AddItem "one"
    lst.AddItem "two"
    before = lst.ListCount
    With lst
        .Clear
    End With
    ok3 = (before = 2) And (lst.ListCount = 0)

    Debug.Print "PC01-CUR=" & CStr(ok1)
    Debug.Print "PC02-SEP=" & CStr(ok2) & " aX=" & CStr(CLng(picA.CurrentX * 100)) & " bX=" & CStr(CLng(picB.CurrentX * 100))
    Debug.Print "PC03-CLEAR=" & CStr(ok3) & " before=" & CStr(before) & " after=" & CStr(lst.ListCount)
    Debug.Print "PC-DONE"
    Unload Me
End Sub
