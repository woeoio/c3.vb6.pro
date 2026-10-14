# DTPicker 控件（日期时间选择器）

> 部件：Microsoft Windows Common Controls-2 6.0 (SP6)，MSCOMCT2.OCX

## 1. 概述
文本区 + 下拉月历的日期/时间选择控件，也可切换为微调按钮样式；
比 MaskEdBox 录日期更友好，杜绝非法日期。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Value` | Date 型日期值（核心进出参数） |
| `Format` | 0 - dtpShortDate 短日期（默认）、1 - dtpLongDate 长日期、2 - dtpTime 时间、3 - dtpCustom 自定义 |
| `CustomFormat` | Format=3 时生效，如 "yyyy-MM-dd HH:mm"、"yyyy'年'M'月'd'日'" |
| `CheckBox` | True 时前置复选框；不勾表示"无日期/空值"（可选日期场景） |
| `MinDate / MaxDate` | 允许选择的日期范围 |
| `UpDown` | True 改为微调按钮样式（不弹月历） |
| `CalendarBackColor / CalendarForeColor` | 下拉月历底色 / 字色 |
| `CalendarTitleBackColor / CalendarTitleForeColor` | 月历标题栏底色 / 字色 |
| `CalendarTrailingForeColor` | 非本月日期字色 |

## 3. 事件

| 事件 | 说明 |
|---|---|
| `Change` | Value 改变时触发（同步界面/重查询常用） |
| `DropDown / CloseUp` | 月历弹出 / 收起 |
| `CallbackKeyDown(KeyCode, Shift, CallbackField, NextChar)` | CustomFormat 含回调字段（如 "ddd" 自定义段）时编辑按键触发 |
| 常规 | Click、KeyDown/KeyPress/KeyUp 等 |

## 4. 示例

```vb
DTPicker1.Format = dtpCustom
DTPicker1.CustomFormat = "yyyy-MM-dd"
DTPicker1.Value = Date                       ' 默认今天
DTPicker1.MinDate = #1/1/2000#: DTPicker1.MaxDate = Date

' 拼 SQL 时务必格式化，避免系统区域差异（如 2024/5/1 vs 05/01/2024）
sql = "SELECT * FROM 订单 WHERE 日期 = #" & _
      Format(DTPicker1.Value, "yyyy-mm-dd") & "#"

Private Sub DTPicker1_Change()
    lblTip.Caption = "已选：" & Format(DTPicker1.Value, "yyyy年mm月dd日")
End Sub
```

> 坑：① CheckBox=True 且未勾选时代表空日期，读 Value 前先判断勾选状态（值可能为 Null）；② Value 是 Date 型，写库/拼 SQL 一律用 Format 转字符串；③ CustomFormat 中想原样显示中文要放在单引号里。