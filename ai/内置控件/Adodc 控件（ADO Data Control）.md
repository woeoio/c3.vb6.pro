# Adodc 控件（ADO Data Control）

> 部件：Microsoft ADO Data Control 6.0 (SP6)，MSADODC.OCX

## 1. 概述
ADO 版的"Data 控件"：设计期配置 ConnectionString 与 RecordSource 即可连接
Access / SQL Server / Oracle 等数据库，运行时自带导航按钮，
并为绑定控件（TextBox、DataGrid 等）提供数据源。是 Data 控件（DAO）的替代者。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `ConnectionString` | 连接串，如 "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\db.mdb;Persist Security Info=False" |
| `CommandType` | 1 - adCmdText（SQL 语句）、2 - adCmdTable（表名）、4 - adCmdStoredProc（存储过程）、0 未知 |
| `RecordSource` | 数据源内容（含义随 CommandType） |
| `Recordset` | （运行时）返回 ADO Recordset 对象，代码增删改查都走它 |
| `CursorLocation` | 2 - adUseServer 服务器游标、3 - adUseClient 客户端游标（离线、Sort/Filter 常用） |
| `CursorType` | 0 仅向前、1 键集、2 动态、3 静态（adOpenForwardOnly/Keyset/Dynamic/Static） |
| `LockType` | 1 只读、2 悲观锁、3 乐观锁（常用）、4 批乐观 |
| `BOFAction / EOFAction` | 到界时：0 移首/末条（默认）、1 停留 BOF/EOF；EOFAction 另有 2 - 自动 AddNew |
| `Mode` | 打开模式：1 只读、2 只写、3 读写 等 |
| `MaxRecords` | 限制返回记录数，0 不限 |
| `Caption / Visible / Enabled` | 控件条标题 / 可见 / 可用 |

## 3. 事件

| 事件 | 说明 |
|---|---|
| `WillMove(adReason, adStatus, pRecordset)` | 记录指针移动**前**，可设 adStatus = adStatusCancel 取消 |
| `MoveComplete(adReason, adStatus, pRecordset)` | 记录指针移动**后**（刷新"第 x/y 条"界面常用） |
| `WillChangeRecord / RecordChangeComplete` | 记录内容改变前/后 |
| `EndOfRecordset` | 到达记录集末尾 |
| `FetchProgress / FetchComplete` | 异步取数进度/完成 |
| `InfoMessage` | 数据提供者发来的提示消息 |

## 4. 方法
`Refresh` —— 重新打开记录集（改完 RecordSource 后必须调用）。
其余数据操作全部通过 `Recordset` 对象：AddNew / Update / Delete / MoveNext / Find / Sort / Filter。

## 5. 绑定
绑定控件设 `DataSource = Adodc1`、`DataField = "字段名"`；表格用 `DataGrid1.DataSource = Adodc1`。

## 6. 示例

```vb
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                          App.Path & "\demo.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM 员工 WHERE 部门='销售' ORDER BY 编号"
Adodc1.Refresh

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
        adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    lblPos.Caption = pRecordset.AbsolutePosition + 1 & " / " & pRecordset.RecordCount
End Sub

' 增删改走 Recordset
Adodc1.Recordset.AddNew
Adodc1.Recordset!姓名 = "张三"
Adodc1.Recordset.Update
```

> 坑：仅向前游标下 RecordCount 可能为 -1；要准确计数/排序/筛选请用静态游标或客户端游标（CursorLocation=3）。
## 7. C3 实现（C30-5，2026-11-15）

本仓**不复刻 MSADODC.OCX**，Adodc 是一枚自注册的无外观窗口（类 `VB6_ADODC`）+ RTL 记录集外壳。

### 7.1 底座 = ACE 的 ODBC 层（不是 OLEDB）

`ConnectionString` 里的 ACE/Jet OLEDB 串由 RTL **翻译**成等价 ODBC 串，引擎仍是
Microsoft Access Database Engine：

| 写法 | 翻译成 |
|---|---|
| `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<目录>;Extended Properties="text;HDR=Yes;..."` | `Driver={Microsoft Access Text Driver (*.txt, *.csv)};DefaultDir=<目录>;Extensions=csv,txt;HDR=Yes;FMT=Delimited` |
| 同上但 `Data Source=<文件>.mdb/.accdb` | `Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=<文件>` |
| 已经是 `Driver=...` 的 ODBC 串 | 原样交出 |

为什么不用 OLEDB：本机实测（cscript 64 位）`ADODB.Connection` 能建、连接也通（State=1），
但 `Recordset.Open` / `Connection.Execute` 一律 `0x80040154 Class not registered` ——
Office Click-to-Run 把 ACE 装在 `...\root\VFS\...\OFFICE16\ACEOLEDB.DLL`，provider 内部再要
一档 Recordset 就落空。同引擎的 **ODBC** 驱动（Access / Text）是 64 位正常注册的，
实测 `SELECT ... FROM [people.csv]` 出 3 行。32 位没有 ACE（`3706 未找到提供程序`）。

### 7.2 与 Data 控件共用引擎

每枚 Adodc 内部建一个**隐藏的 `VB6_DATA` 窗口**当引擎宿主，连接/游标/行缓存/字段读全在
`vb6forms_data.c` 一份。于是 `Adodc1.Recordset` 直接交 `vb6_Data_RecordsetObj` 的
VB6_MEMK_RS 成员对象 —— `RecordCount` / `Fields("x").Value` / `MoveFirst` / `BOF` / `EOF`
这一整条晚绑定链原样可用。

### 7.3 属性面

已实现：`ConnectionString`、`RecordSource`、`CommandType`（2 = 表名时补成 `SELECT * FROM [表]`）、
`CursorLocation`、`CursorType`、`LockType`、`Mode`、`MaxRecords`、`BOFAction`、`EOFAction`、
`Recordset`。`Refresh` 真重新开（翻译连接串 → 重开游标）。
C3 扩展两枚夹具证人：`Opened`、`LastRowCount`。

**没做**（如实挂账）：

1. **事件**（`WillMove` / `MoveComplete` / `WillChangeRecord` / `EndOfRecordset` / `FetchProgress`）——
   ADO 事件要 `IConnectionPoint` 汇，本实现不建 ADO 对象，所以一条都没发。
2. `MaxRecords` 只存值，不截断。
3. `CursorLocation/CursorType/LockType/Mode` 存值：ODBC 那一层固定 `SQL_CURSOR_STATIC` + 只读。
4. 导航按钮（控件自带的那四个箭头）没有画。
5. `.accdb/.mdb` 路径只有翻译式的代码，未实测（本机没有样本库）。

### 7.4 判据（`tests/ctrlado`，x64，`exit=0`）

数据 `tests/ctrlado/data/people.csv`（3 行），`App.Path` 旁置后由 ACE Text 驱动读：

```
AD1-conn=True     ConnectionString 回读含 ACE.OLEDB
AD2-src=True      RecordSource 回读
AD3-open=-1       Refresh 真开（VB6 True = -1）
AD4-rows=3        Recordset.RecordCount（晚绑定链）
AD5-grid=3|3      DataGrid 绑定后的 RowCount|ColCount
AD6-cell=Alice    Row=0 Col=1
AD7-cell=Sales    Row=2 Col=2
AD8-rs=Alice      Recordset.Fields("name").Value
AD9-filter=1      换 RecordSource 再 Refresh（真重开）
AD10-grid2=1|3    DataGrid.ReBind 跟着换
AD11-val=Bob      过滤后只剩 Ops 那条
```

x86：本机没有 32 位 ACE，`run_tests` 用 `RequiresCom Microsoft.ACE.OLEDB.12.0` 打 SKIP；
手工跑会看到 `AD3-open=0` 且后续触碰 `Recordset` 抛 438 —— 这是"驱动缺失"的真实现场。
