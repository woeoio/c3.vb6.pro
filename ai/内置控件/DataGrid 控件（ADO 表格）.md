# DataGrid 控件（ADO 表格）

> 部件：Microsoft DataGrid Control 6.0 (SP6)，MSDATGRD.OCX

## 1. 概述
与 Adodc（或 ADO Recordset）配对的数据库表格控件，可直接在格内编辑、
新增、删除记录，是 VB6 后期数据库界面的标配网格。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `DataSource` | 数据源：Adodc 控件 / Data 控件 / Recordset 对象 |
| `Columns` | 列集合（Column 对象），设计期属性页或运行期配置 |
| `AllowAddNew` | 是否允许在末行新增记录 |
| `AllowUpdate` | 是否允许格内编辑 |
| `AllowDelete` | 是否允许删除记录 |
| `ColumnHeaders` | 是否显示列头 |
| `Bookmark` | 当前行书签（Variant），可保存/恢复当前行 |
| `SelBookmarks` | 选中行书签集合（配合 MultiSelect 高亮多行） |
| `RowHeight / HeadFont / DefColWidth` | 行高 / 列头字体 / 默认列宽 |

**Column 对象要点：** `Caption`（列头文字）、`DataField`（绑定字段）、
`Width`、`Visible`、`Locked`（锁定不可编辑）、`Alignment`、
`Button`（True 时格内显示下拉按钮，配合 DataCombo 做外键选择）。

## 3. 事件

| 事件 | 说明 |
|---|---|
| `BeforeColEdit(Cancel, ColIndex)` / `AfterColEdit(ColIndex)` | 单元格编辑前（可 Cancel）/后 |
| `BeforeColUpdate(Cancel, ColIndex, OldValue)` / `AfterColUpdate` | 单元格写回缓存前（可校验取消）/后 |
| `BeforeUpdate(Cancel)` / `AfterUpdate` | 整行写回数据库前/后 |
| `BeforeInsert(Cancel)` / `AfterInsert` | 新增行前/后 |
| `BeforeDelete(Cancel)` / `AfterDelete` | 删除行前/后 |
| `RowColChange(LastRow, LastCol)` | 当前单元格改变 |
| `Scroll` | 滚动时 |
| `Error` | 数据操作出错 |

## 4. 方法
`Refresh` 重绘刷新｜`ReBind` 数据源结构变化后重新绑定｜`ClearSelRows` 清除选中行高亮

## 5. 示例

```vb
Set DataGrid1.DataSource = Adodc1
DataGrid1.AllowAddNew = True: DataGrid1.AllowDelete = True

' 运行期配置列
Dim c As Column
Set c = DataGrid1.Columns("工资")
c.Caption = "月薪": c.Width = 1200: c.Alignment = dbgRight   ' 右对齐
DataGrid1.Columns("内部ID").Visible = False

' 格内校验：工资必须为数字且 > 0
Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, _
        OldValue As Variant, Cancel As Integer)
    If DataGrid1.Columns(ColIndex).DataField = "工资" Then
        If Not IsNumeric(Adodc1.Recordset!工资) Or Adodc1.Recordset!工资 <= 0 Then
            MsgBox "工资必须为正数": Cancel = True
        End If
    End If
End Sub
```

> 坑：编辑中的值先落在绑定缓存，BeforeColUpdate 里读 Recordset 字段已是新值，OldValue 才是旧值；取消后界面不会自动还原，需配合 `Adodc1.UpdateControls` 类手段或重设单元格。
## 6. C3 实现（C30-5，2026-11-15）

自绘表格（类 `VB6_DATAGRID`），不加载 MSDATGRD.OCX。数据从**引擎**来：绑定对象是
Adodc 的隐藏引擎窗（`vb6_Adodc_Engine`）或 Data 控件本身。

### 6.1 绑定

- 设计期：`.frm` 里 `DataSource = "Adodc1"`。cgen 在所有控件创建完之后统一发
  `vb6_DataGrid_Bind(格窗, 引擎窗)`，并按源控件类型分岔（Adodc 取引擎、Data 就是它自己）。
- 运行期：`Set DataGrid1.DataSource = Adodc1` 这一形**没做**（要动对象赋值那条路），
  现在**只有设计期绑定**。要换源请在 `.frm` 里写，或重新 `ReBind`。

### 6.2 画出来的东西

列头（`ColumnHeaders`）、格线、选中格高亮（`COLOR_HIGHLIGHT`）、竖向滚动条（行粒度）、
鼠标滚轮。单元格取值走 `vb6_Data_FieldValueAt(引擎, 行, 列)` —— 它**不动调用者的当前行**
（记下 curRow → FIRST+NEXT 走到目标行读一格 → 走回原行），所以画表格不会把
`Adodc1.Recordset` 挪走。点某一行会把源记录集定位到那一行（与 VB6 一致）。

### 6.3 属性 / 方法

已实现：`DataSource`、`RowCount`、`ColCount`（读）、`Row`、`Col`（读写，越界钳住）、
`Text`（当前格，只读快照）、`ColumnHeaders`、`AllowAddNew`、`AllowUpdate`、`AllowDelete`
（存值）、`RowHeight`、`DefColWidth`（缇；画时 /15 换像素）。
方法：`Refresh`（重读行列数并重绘）、`ReBind`（重置选中/滚动后重读）、`ClearSelRows`。

**没做**（如实挂账）：

1. **格内编辑**：`AllowUpdate` / `AllowAddNew` / `AllowDelete` 只存值，双击不进编辑态，
   写回数据库这一步完全没有。`BeforeColUpdate` / `BeforeUpdate` / `RowColChange` 等事件一条没发。
2. `Columns` 集合（逐列 Caption/Width/Visible/Alignment/Button）没做 —— 现在所有列同宽
   （`DefColWidth`），列头文字直接取字段名。
3. `Bookmark` / `SelBookmarks`（多选行）没做。
4. 横向滚动没有：列多了会被窗口右边裁掉。

### 6.4 判据

与 Adodc 同一份夹具（`tests/ctrlado`），表格读数：

```
AD5-grid=3|3      RowCount|ColCount
AD6-cell=Alice    Row=0 Col=1
AD7-cell=Sales    Row=2 Col=2
AD10-grid2=1|3    ReBind 后随数据源换（WHERE dept='Ops'）
```

x64 全绿；x86 因本机无 32 位 ACE 而 SKIP（见 Adodc 说明 §7.4）。
