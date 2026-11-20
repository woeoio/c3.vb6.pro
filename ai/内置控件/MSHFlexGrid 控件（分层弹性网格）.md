# MSHFlexGrid 控件（分层弹性网格）

> 部件：Microsoft Hierarchical FlexGrid Control 6.0，MSHFLXGD.OCX；
> 同族的 MSFlexGrid（MSFLXGRD.OCX）在 C3 里**共用同一套实现**（MSHFlexGrid 是超集，
> 字符串矩阵那一面完全一样）。

## 1. 概述

VB6 里最常用的"表格"控件。与 DataGrid 的分界要记牢：

| | MSHFlexGrid / MSFlexGrid | DataGrid |
|---|---|---|
| 数据在哪 | **控件自己的 Rows×Cols 单元格**（字符串矩阵） | 外部记录集（Adodc / Data / ADO Recordset） |
| 典型写法 | `grd.Rows = 6 : grd.TextMatrix(1,2) = "x"` | `Set grd.DataSource = rs` |
| 依赖 | 无（不需要数据库、不需要 ACE） | 需要 ACE/ODBC 底座 |
| 编辑 | `Editable = True` 时格内直接改字符串 | 改回记录集（AllowUpdate） |

因为不依赖数据库，它是**数据录入界面**（单据、库存、明细行）的主力；数据库例程里也常
把查询结果先读成字符串再灌进来。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Rows` / `Cols` | 总行/列数（含冻结区）。改小时 VB6 也会保留还能放下的内容 |
| `FixedRows` / `FixedCols` | 冻结行/列（表头），不随滚动移动 |
| `Row` / `Col` | 当前格（0 基） |
| `RowSel` / `ColSel` | 选区另一端（配合 `Row` 得到矩形选区） |
| `Text` | 当前格的文本 |
| `TextMatrix(r, c)` | **任意格**读写（最常用的一个成员；两参索引属性） |
| `TextArray(i)` | 行主序线性下标读写（`i = r * Cols + c`） |
| `ColWidth(i)` / `RowHeight(i)` | 列宽 / 行高（缇，1 px = 15 缇）；`RowHeight(-1)` 是缺省行高 |
| `ColAlignment(i)` | 列对齐：0=左 1=右 2=中 |
| `GridLines` / `ScrollBars` | 格线 / 滚动条（0=None 1=H 2=V 3=Both） |
| `Editable` | 数据区是否允许键入（冻结区始终只读） |
| `SelectionMode` / `FillStyle` / `AllowUserResizing` / `MergeCells` | 行为位（见"未做"） |
| `FormatString` | 设计期排版串（`|<列名;对齐;宽` 形式），本版只存串不参与排版 |
| `MouseRow` / `MouseCol` | 鼠标所在格（不在格上时 -1） |
| `CellHeight/CellWidth/CellLeft/CellTop` | 当前格的几何读数 |
| `Sort` | **属性**而不是方法：写 1 升序 / 2 降序，立即按当前列排数据区 |

## 3. 事件

`Click` / `DblClick` / `MouseDown` / `MouseMove` / `MouseUp` / `KeyPress` / `KeyDown` /
`Scroll` / `EnterCell` / `LeaveCell` / `SelChange` / `RowColChange`。

真实例程最依赖的是 `Click` 里读 `Row` 再取 `TextMatrix(Row, 1)`。

## 4. 方法

`AddItem s$`（按 `vbTab` 拆列）、`RemoveItem i`、`Clear`（只清内容）、
`ClearStructure`（连结构收回到冻结区大小）、`Refresh`、`Redraw`（**属性**）。

## 5. C3 实现现状（C30-6）

实现方式：**自绘窗口类 `VB6_MSHFLEXGRID`**（`src/rtl/core/vb6forms/vb6forms_mshflexgrid.c`），
不加载 MSHFLXGD.OCX / MSFLXGRD.OCX。

### 5.1 已做

- 字符串矩阵：`Rows/Cols` 可改且**保留内容**（行列都取交集搬过去），`FixedRows/FixedCols` 冻结区；
- 读写：`TextMatrix(r,c)`、`TextArray(i)`、`Text`、`ColWidth(i)`、`RowHeight(i)`、`ColAlignment(i)`；
- 自绘：冻结区 3D 面板、格线、当前格高亮（失焦变灰）、拖选矩形、竖/横滚动条、滚轮；
- 交互：点选/拖选、方向键/Home/End/PageUp/PageDown/Tab、`Editable=True` 时**格内直接键入**
  （回车/方向键提交，Esc 取消，Delete 清格）；
- 方法：`AddItem`（Tab 拆列）/ `RemoveItem` / `Clear` / `ClearStructure` / `Refresh`；
- `Sort` 属性（1/2 立即排序，冻结行不动）。

### 5.2 未做（如实挂账）

- **分层记录集**（MSHFlexGrid 名字里的 "Hierarchical"）：`DataSource` + `DataMember`、
  带区（Band）显示 —— 本版只把 `DataSource` 当字符串存/回读，**不做绑定**；
- `MergeCells` 真合并、`AllowUserResizing` 列宽拖拽（只存值）；
- `FormatString` 的排版语义（只存串）；
- 单元格内嵌图片、掩码、`CellFont` 等富格式；
- 事件（`Click` 以外的 `EnterCell`/`RowColChange` 等）暂未派发。

## 6. 判据（tests/ctrlmshflexgrid）

夹具不碰数据库，x64/x86 都真跑（这与 C30-5 DataGrid 在无 ACE 的 CI 上 SKIP 形成对照）：

```
FG1-shape=6|3|1|1     设计期 Rows/Cols/FixedRows/FixedCols
FG2-tm=Alice          TextMatrix 写→读
FG3-grow=8|Alice      改 Rows 保留内容
FG4-text=2,2|Qty7     Row/Col/Text
FG5-add=Bob|Sales     AddItem 按 Tab 拆列
FG6-col=2400|1        ColWidth/ColAlignment
FG7-sort=a1|c3        Sort 属性
FG8-arr=a1            TextArray
FG9-clear=6|0         Clear 只清内容
FG10-cs=1|1           ClearStructure
FG11-misc=1|3|-1|False  GridLines|ScrollBars|MouseRow|Editable
FG12-fmt=True         FormatString 回读
FG13-rm=4|keep3       RemoveItem
FG14-cols=5|keep1     改 Cols 保留内容
FG-DONE
```

## 7. 真实例程用法统计（D:\vb_yqt4qPac\数据库，601 个 .frm）

属性/方法出现次数（本刀接口面的依据）：

```
TextMatrix/2 454   Rows 128   Col 115   Text 103   Row 85   Clear 40
ColWidth/1 50      AddItem 27+1   Cols 23   RowSel 22   SetFocus 18
Refresh 13         Editable 13    MouseRow 12   Redraw 12   TopRow 12
FormatString 11    CellHeight/CellWidth 7/6   RowHeight/1 6   CellBackColor 5
FixedRows 5        TextArray 4   FixedCols 3   LeftCol 3   SelectionMode 2
ColSel 2           ScrollBars 2   ColAlignment/1 4   ClearStructure 2
RemoveItem 4+1     Sort 5         RowHeightMin 1
```
