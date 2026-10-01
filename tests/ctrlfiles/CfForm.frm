VERSION 5.00
Begin VB.Form CfForm 
   Caption         =   "CfForm"
   ClientHeight    =   3200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "CfForm"
   ScaleHeight     =   3200
   ScaleWidth      =   6000
   Begin VB.ListBox auxList 
      Height          =   600
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   1200
   End
   Begin VB.DriveListBox drvList 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.DirListBox dirList 
      Height          =   1410
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      Path            =   "C:\Windows\System32"
   End
   Begin VB.FileListBox fileList 
      Height          =   1410
      Left            =   1800
      Pattern         =   "*.dll"
      TabIndex        =   2
      Top             =   600
      Width           =   1455
      Path            =   "C:\Windows\System32"
   End
End
Attribute VB_Name = "CfForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 C29-1b: 文件系统三控件走原生 COMBOBOX / LISTBOX。
' 改之前这三类同样在 controlTypeToWin32Class 里缺格 => 控件根本没窗口, RTL 那套
' P20-37 的填充 helper 从来没拿到过句柄。ListCount > 0 就是"窗口真在 + 填进去过"
' 的双重证明 (空句柄时 LB_GETCOUNT 走不到、填充也走不到)。
' 口径: Dir 的每一项是 [名字] (VB6 的方括号约定, 双击靠它下钻); 三控件的
' Path/Pattern 一赋值就重刷列表 —— 这就是 VB6 的联动机制。

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

' 联动的三条 (VB6 手册里"简易文件浏览器"的写法)。运行期靠通知触发, 这里也直接调,
' 好在无窗口环境里把逻辑本身钉住; 通知接线由 cf_emitc_shape 断发码形状。
Private Sub drvList_Change()
    dirList.Path = drvList.Drive
End Sub

Private Sub dirList_Change()
    fileList.Path = dirList.Path
End Sub

Private Sub fileList_Click()
    auxList.AddItem fileList.FileName
End Sub

Private Sub fileList_DblClick()
    auxList.AddItem "dbl:" & fileList.FileName
End Sub

Private Sub Form_Load()

    ' --- 1. 三个控件都有窗口、都填进去过 ---
    Debug.Print "CF1=" & TF(drvList.ListCount > 0 And dirList.ListCount > 0 And fileList.ListCount > 0)
    Debug.Print "CF2=" & TF(Mid(drvList.List(0), 2, 1) = ":")
    Debug.Print "CF3=" & TF(Mid(drvList.Drive, 2, 1) = ":")

    ' --- 2. Dir 的 [名字] 约定 ---
    Debug.Print "CF4=" & TF(Mid(dirList.List(0), 1, 1) = "[")

    ' --- 3. 设计期 Path / Pattern 落位 ---
    Debug.Print "CF5=" & TF(dirList.Path = "C:\Windows\System32")
    Debug.Print "CF6=" & TF(fileList.Path = "C:\Windows\System32" And fileList.Pattern = "*.dll")
    Debug.Print "CF7=" & TF(InStr(LCase(fileList.List(0)), ".dll") > 0)

    ' --- 4. 改 Pattern 立刻重刷 (无匹配 → 空, 改回来 → 非空) ---
    fileList.Pattern = "*.zzz-not-there"
    Debug.Print "CF8=" & TF(fileList.ListCount = 0)
    fileList.Pattern = "*.dll"
    Debug.Print "CF9=" & TF(fileList.ListCount > 0)

    ' --- 5. ListIndex / FileName 回路 ---
    ' 这条比较刻意写成 VB6 的原样 (`FileName = List(0)`): 两侧的推断类型必须是 String,
    ' 否则发码落到 vb6_VarCmpEq, 而 `List(i)` 的 RTL 声明是 void* → 装箱成 VT_UNKNOWN,
    ' 同一条读数会 x64 为真、x86 为假 (Fix 194 同族的坑)。
    fileList.ListIndex = 0
    Debug.Print "CF10=" & TF(fileList.ListIndex = 0 And fileList.FileName = fileList.List(0))
    fileList_DblClick
    Debug.Print "CF11=" & TF(auxList.List(auxList.ListCount - 1) = "dbl:" & fileList.FileName)

    ' --- 6. 联动: 换目录 → 文件列表跟着换 ---
    dirList.Path = "C:\Windows"
    dirList_Change
    Debug.Print "CF12=" & TF(fileList.Path = "C:\Windows" And fileList.ListCount > 0)

    ' --- 7. 换盘 → 目录列表跟着走 (C:\ 的根目录必然有子目录) ---
    drvList.Drive = "C:"
    drvList_Change
    Debug.Print "CF13=" & TF(dirList.Path = "C:" And dirList.ListCount > 0)

    ' --- 8. 原生对照: auxList 是普通 ListBox, 读数口径必须与三控件一致 ---
    fileList_Click
    auxList.AddItem "one"
    Debug.Print "CF14=" & TF(auxList.ListCount = 3 And Mid(auxList.List(0), 1, 4) = "dbl:")

    ' --- 9. 账 #68: 窗体模块里裸 Left(...) 必须是**内置函数** (窗体的 Left 属性要写 Me.Left) ---
    ' 修复前这一句被属性抢走: 发成 vb6_GetControlLeft(hwnd) 再拼实参, 编得过而读数不是字符串。
    Debug.Print "CF15=" & TF(Left(auxList.List(0), 4) = "dbl:")
    Debug.Print "CF16=" & TF(Right("dbl:xyz", 3) = "xyz" And Left("dbl:xyz", 4) = "dbl:")


    Debug.Print "CTRLFILES-DONE"
    Unload Me
End Sub
