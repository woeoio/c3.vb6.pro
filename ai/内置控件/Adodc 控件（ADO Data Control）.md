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