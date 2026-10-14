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