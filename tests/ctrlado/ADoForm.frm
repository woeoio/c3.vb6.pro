VERSION 5.00
Begin VB.Form ADoForm
   Caption         =   "Adodc / DataGrid"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Begin MSAdodcLib.Adodc Adodc1
      Height          =   375
      Left            =   240
      Top             =   3720
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   661
      _Version        =   393216
      CommandType     =   1
      CursorLocation  =   3
      CursorType      =   3
      LockType        =   3
      Mode            =   3
      RecordSource    =   ""
      ConnectionString=   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1
      Height          =   2880
      Left            =   240
      Top             =   360
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   5080
      _Version        =   393216
      ColumnHeaders   =   -1  ' True
      AllowAddNew     =   0   ' False
      AllowUpdate     =   -1  ' True
      AllowDelete     =   0   ' False
      DataSource      =   "Adodc1"
   End
   Begin VB.Timer Timer1
      Interval        =   300
      Left            =   5880
      Top             =   360
   End
End
Attribute VB_Name = "ADoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' C30-5 判据：Adodc 的底座 = ACE（ODBC 层），数据是 tests/ctrlado/data/people.csv。
' 这一版只验"读"，编辑面（AllowUpdate 真写回）在 ai/029 里挂账。
Private Sub Form_Load()
    Dim cs As String
    Dim dataDir As String
    ' 数据目录：先看夹具钩子 C3_ADO_DATA（run_tests 给绝对路径，避免依赖“拷到 exe 旁边”），
    ' 否则 App.Path & "\data"（真 VB6 工程把库放 exe 同目录的写法）。
    dataDir = Environ("C3_ADO_DATA")
    If Len(dataDir) = 0 Then dataDir = App.Path & "\data"
    cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dataDir _
       & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"""
    Adodc1.ConnectionString = cs
    Adodc1.RecordSource = "SELECT id, name, dept FROM [people.csv] ORDER BY id"
    Adodc1.Refresh
    Debug.Print "AD1-conn=" & CStr(InStr(Adodc1.ConnectionString, "ACE.OLEDB") > 0)
    Debug.Print "AD2-src=" & CStr(InStr(Adodc1.RecordSource, "people.csv") > 0)
    Debug.Print "AD3-open=" & CStr(Adodc1.Opened)
    Debug.Print "AD4-rows=" & CStr(Adodc1.Recordset.RecordCount)

    ' 表格：绑定在 .frm 里写的 DataSource = "Adodc1"，cgen 解析成 Adodc 的引擎窗
    DataGrid1.Refresh
    Debug.Print "AD5-grid=" & CStr(DataGrid1.RowCount) & "|" & CStr(DataGrid1.ColCount)
    DataGrid1.Row = 0
    DataGrid1.Col = 1
    Debug.Print "AD6-cell=" & DataGrid1.Text
    DataGrid1.Row = 2
    DataGrid1.Col = 2
    Debug.Print "AD7-cell=" & DataGrid1.Text
    Debug.Print "AD8-rs=" & Adodc1.Recordset.Fields("name").Value
End Sub

Private Sub Timer1_Timer()
    Static done As Integer
    If done Then Exit Sub
    done = 1
    ' 换 RecordSource 再 Refresh —— 验"重新开"这条（不是只重画）
    Adodc1.RecordSource = "SELECT id, name, dept FROM [people.csv] WHERE dept = 'Ops'"
    Adodc1.Refresh
    Debug.Print "AD9-filter=" & CStr(Adodc1.LastRowCount)
    DataGrid1.ReBind
    Debug.Print "AD10-grid2=" & CStr(DataGrid1.RowCount) & "|" & CStr(DataGrid1.ColCount)
    Debug.Print "AD11-val=" & Adodc1.Recordset.Fields("name").Value
    Debug.Print "AD-DONE"
    Unload Me
End Sub
