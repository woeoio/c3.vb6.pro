VERSION 5.00
Begin VB.Form MfgForm
   Caption         =   "MSHFlexGrid"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1
      Height          =   2880
      Left            =   240
      Top             =   360
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   5080
      _Version        =   393216
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      ColWidth        =   1200        2400        1200
      GridLines       =   1
      ScrollBars      =   3
      FormatString    =   "|<Name|<Qty"
   End
   Begin VB.Timer Timer1
      Interval        =   300
      Left            =   5880
      Top             =   360
   End
End
Attribute VB_Name = "MfgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' C30-6 判据：MSHFlexGrid 是**字符串矩阵**，与数据源无关 —— 这枚夹具不碰数据库，
' 所以 GA runner 上没有 ACE 也能真跑（这是与 C30-5 DataGrid 最大的区别）。
' 每条判据都是 ASCII 前缀 + 读数，逐字比。
Private Sub Form_Load()
    Dim s As String
    Dim r As Long
    Dim c As Long

    ' FG1: 设计期属性（.frm 里的 Rows/Cols/FixedRows/FixedCols）真灌进去了
    Debug.Print "FG1-shape=" & CStr(MSHFlexGrid1.Rows) & "|" & CStr(MSHFlexGrid1.Cols) _
        & "|" & CStr(MSHFlexGrid1.FixedRows) & "|" & CStr(MSHFlexGrid1.FixedCols)

    ' FG2: TextMatrix(r,c) 写 → 读（索引属性，两参；读侧走 GetTextMatrix，写侧走 SetTextMatrix）
    MSHFlexGrid1.TextMatrix(1, 1) = "Alice"
    Debug.Print "FG2-tm=" & MSHFlexGrid1.TextMatrix(1, 1)

    ' FG3: 改 Rows 必须**保住**已有内容（VB6 行为；真实例子全靠这一条）
    MSHFlexGrid1.Rows = 8
    Debug.Print "FG3-grow=" & CStr(MSHFlexGrid1.Rows) & "|" & MSHFlexGrid1.TextMatrix(1, 1)

    ' FG4: Row/Col + Text（当前格）—— 写读同格
    MSHFlexGrid1.Row = 2
    MSHFlexGrid1.Col = 2
    MSHFlexGrid1.Text = "Qty7"
    Debug.Print "FG4-text=" & CStr(MSHFlexGrid1.Row) & "," & CStr(MSHFlexGrid1.Col) & "|" & MSHFlexGrid1.Text

    ' FG5: AddItem 按 Tab 拆列（当前行在冻结区 ⇒ 追加到末尾）
    MSHFlexGrid1.Row = 0
    MSHFlexGrid1.AddItem "Bob" & vbTab & "Sales"
    Debug.Print "FG5-add=" & MSHFlexGrid1.TextMatrix(8, 0) & "|" & MSHFlexGrid1.TextMatrix(8, 1)

    ' FG6: ColWidth(i) / ColAlignment(i) —— 一参索引属性，读回同一串
    MSHFlexGrid1.ColWidth(1) = 2400
    MSHFlexGrid1.ColAlignment(1) = 1
    Debug.Print "FG6-col=" & CStr(MSHFlexGrid1.ColWidth(1)) & "|" & CStr(MSHFlexGrid1.ColAlignment(1))

    ' FG7: Sort 是**属性**（写 1 = 升序，立即按当前列排数据区）。造三个乱序值再排。
    MSHFlexGrid1.Rows = 6
    MSHFlexGrid1.Col = 0
    MSHFlexGrid1.TextMatrix(1, 0) = "c3"
    MSHFlexGrid1.TextMatrix(2, 0) = "a1"
    MSHFlexGrid1.TextMatrix(3, 0) = "b2"
    MSHFlexGrid1.TextMatrix(4, 0) = "d4"
    MSHFlexGrid1.TextMatrix(5, 0) = "e5"
    MSHFlexGrid1.Sort = 1
    Debug.Print "FG7-sort=" & MSHFlexGrid1.TextMatrix(1, 0) & "|" & MSHFlexGrid1.TextMatrix(3, 0)

    ' FG8: TextArray(i) 线性下标（行主序：i = r * Cols + c）
    Debug.Print "FG8-arr=" & MSHFlexGrid1.TextArray(1 * 3 + 0)

    ' FG9: Clear 只清内容，结构不动
    MSHFlexGrid1.Clear
    Debug.Print "FG9-clear=" & CStr(MSHFlexGrid1.Rows) & "|" & CStr(Len(MSHFlexGrid1.TextMatrix(1, 0)))

    ' FG10: ClearStructure 连结构一起收回到冻结区大小
    MSHFlexGrid1.ClearStructure
    Debug.Print "FG10-cs=" & CStr(MSHFlexGrid1.Rows) & "|" & CStr(MSHFlexGrid1.Cols)

    ' FG11: 读回来的一批标量面（不崩、值合理即可）
    Debug.Print "FG11-misc=" & CStr(MSHFlexGrid1.GridLines) & "|" & CStr(MSHFlexGrid1.ScrollBars) _
        & "|" & CStr(MSHFlexGrid1.MouseRow) & "|" & CStr(MSHFlexGrid1.Editable)

    ' FG12: FormatString 设计期那串能回读
    Debug.Print "FG12-fmt=" & CStr(InStr(MSHFlexGrid1.FormatString, "Name") > 0)

    ' FG13: RemoveItem 真删行（第 2 行上移）
    MSHFlexGrid1.Rows = 4
    MSHFlexGrid1.TextMatrix(1, 0) = "keep1"
    MSHFlexGrid1.TextMatrix(2, 0) = "drop2"
    MSHFlexGrid1.TextMatrix(3, 0) = "keep3"
    MSHFlexGrid1.RemoveItem 2
    Debug.Print "FG13-rm=" & CStr(MSHFlexGrid1.Rows) & "|" & MSHFlexGrid1.TextMatrix(2, 0)

    ' FG14: Cols 改大也要保住内容（列方向同一条规矩）
    MSHFlexGrid1.Cols = 5
    Debug.Print "FG14-cols=" & CStr(MSHFlexGrid1.Cols) & "|" & MSHFlexGrid1.TextMatrix(1, 0)
End Sub

Private Sub Timer1_Timer()
    Static done As Integer
    If done Then Exit Sub
    done = 1
    Debug.Print "FG-DONE"
    Unload Me
End Sub
