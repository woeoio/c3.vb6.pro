# MonthView 控件（月历面板）

> 部件：Microsoft Windows Common Controls-2 6.0 (SP6)，MSCOMCT2.OCX

## 1. 概述
常驻式月历面板（不用下拉），单击选日期、拖动选日期范围，
适合排班、考勤、日程类界面；可多个月份平铺显示。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Value` | 当前选中日期（Date 型） |
| `MultiSelect` | True 允许拖动选择日期范围 |
| `SelStart / SelEnd` | 范围选择的起 / 止日期（MultiSelect=True 时有效） |
| `MaxSelCount` | 范围选择允许的最大天数 |
| `MonthRows / MonthColumns` | 面板中月历的行 / 列数（多月平铺） |
| `ShowToday` | 底部显示"今天：xxxx-xx-xx"栏 |
| `ShowWeekNumbers` | 显示周号 |
| `BackColor / ForeColor` | 面板底色 / 日期字色 |
| `TitleBackColor / TitleForeColor` | 标题栏（年月与箭头）底色 / 字色 |
| `TrailingForeColor` | 非本月日期字色（灰显） |
| `Appearance / BorderStyle / Font / Enabled / Visible` | 常规属性 |

## 3. 事件

| 事件 | 说明 |
|---|---|
| `DateClick(DateSelected As Date)` | **核心**：单击某个日期触发，参数即所点日期 |
| `DateDblClick(DateSelected As Date)` | 双击日期 |
| 常规 | Click、DblClick、GotFocus/LostFocus、KeyDown/KeyPress/KeyUp 等 |

## 4. 示例（与 DTPicker / 列表联动）

```vb
MonthView1.Value = Date
MonthView1.ShowToday = True

Private Sub MonthView1_DateClick(ByVal DateSelected As Date)
    DTPicker1.Value = DateSelected                     ' 同步选择器
    LoadDayPlan DateSelected                           ' 加载当日日程
    lblWeek.Caption = "星期" & Mid("日一二三四五六", Weekday(DateSelected, vbMonday), 1)
End Sub

' 读取范围选择
Private Sub cmdRange_Click()
    If MonthView1.MultiSelect Then
        MsgBox Format(MonthView1.SelStart, "mm/dd") & " ~ " & _
               Format(MonthView1.SelEnd, "mm/dd") & " 共 " & _
               DateDiff("d", MonthView1.SelStart, MonthView1.SelEnd) + 1 & " 天"
    End If
End Sub
```

> 坑：SelStart/SelEnd 只在 MultiSelect=True 且用户拖选后有意义；单点时二者等于 Value。多月平铺时注意控件整体尺寸随 MonthRows×MonthColumns 变大。