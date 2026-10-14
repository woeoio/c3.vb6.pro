# MSHFlexGrid 控件（分层灵活表格）

> 部件：Microsoft Hierarchical FlexGrid Control 6.0 (SP6)，MSHFLXGD.OCX

## 1. 概述
功能最强的只读/半编辑网格：合并单元格、固定行列、逐格格式化、分组折叠、
排序，是**报表显示与打印**的主力。默认不可编辑（需借助 TextBox 覆盖格实现编辑）。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Rows / Cols` | 总行数 / 总列数 |
| `FixedRows / FixedCols` | 固定（表头）行 / 列数 |
| `Row / Col` | 当前单元格位置 |
| `Text` | 当前单元格文字 |
| `TextMatrix(r, c)` | 任意单元格文字（**读写报表最常用**） |
| `ColWidth(i) / RowHeight(i)` | 列宽 / 行高（缇） |
| `ColAlignment(i) / FixedAlignment(i)` | 数据列对齐 / 固定列对齐（0 左 1 右 2 中） |
| `FormatString` | 一次性定义表头：".FormatString = " |姓名|部门|工资"" |
| `DataSource` | 可直接绑 Adodc/Data 自动填表（只读展示） |
| `MergeCells` | True 开启合并；配合 `MergeRow(i)/MergeCol(i)` 合并同内容相邻格 |
| `SelectionMode` | 0 自由、1 整行、2 整列 |
| `FillStyle` | 0 只作用于当前格、1 作用于整个选区（批量着色必设 1） |
| `CellBackColor / CellForeColor / CellFontBold / CellFontSize` | 作用于当前格或选区的单元格格式 |
| `AllowUserResizing` | 0 不允许、1 拖列宽、2 拖行高、3 皆可 |
| `ScrollBars` | 0 无、1 水平、2 垂直、3 两者（默认） |
| `Redraw` | False 暂停重绘（批量填数提速），填完设回 True |
| `Sort` | 设值即排序：1/2 通用升/降、3/4 数值升/降、5/6 字符串升/降、7/8 忽略大小写升/降 |
| `WordWrap` | 单元格自动换行 |

## 3. 事件
`EnterCell` / `LeaveCell`（进出单元格）、`RowColChange`、`SelChange`、
`Scroll`、`Click`、`DblClick`、`Compare`（自定义排序比较）+ 常规鼠标/键盘事件。
（无 ColumnClick：点表头排序请在 MouseDown 里判断 Row < FixedRows。）

## 4. 方法
`AddItem(text, [index])` 加行｜`RemoveItem(index)` 删行｜`Clear` 清内容｜
`Collapse(row)` / `Expand(row)` 分组折叠/展开｜`Refresh`

## 5. 示例（填表 + 隔行着色 + 点表头排序）

```vb
With MSHFlexGrid1
    .Redraw = False
    .FormatString = " |姓名|部门|工资"
    .Rows = 1
    Do While Not Adodc1.Recordset.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Adodc1.Recordset!姓名
        .TextMatrix(.Rows - 1, 2) = Adodc1.Recordset!部门
        .TextMatrix(.Rows - 1, 3) = Adodc1.Recordset!工资
        Adodc1.Recordset.MoveNext
    Loop
    .FillStyle = flexFillRepeat
    Dim r As Long
    For r = .FixedRows To .Rows - 1 Step 2
        .Row = r: .Col = 1: .CellBackColor = &HF0F0F0
    Next
    .Redraw = True
End With

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MSHFlexGrid1.Row < MSHFlexGrid1.FixedRows Then      ' 点了表头
        MSHFlexGrid1.Col = MSHFlexGrid1.Col
        MSHFlexGrid1.Sort = 5                               ' 按当前列字符串升序
    End If
End Sub
```

> 打印提示：循环 TextMatrix 用 `Printer.Print` 配合 `Printer.CurrentX/Y` 排版；或借助 .Clip 属性整块取选区文本。