# Slider 控件（滑杆）

> 部件：Microsoft Windows Common Controls 6.0 (SP6)，MSCOMCTL.OCX

## 1. 概述
轨道+滑块的滑杆控件（音量条/缩放条那种），用于在一定范围内连续取值，
比 TextBox 输数字直观，比 ScrollBar 美观。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Min / Max / Value` | 下限 / 上限 / 当前值 |
| `SmallChange` | 小步长：点箭头或按方向键一次的增量 |
| `LargeChange` | 大步长：点滑块两侧轨道一次的增量 |
| `Orientation` | 0 - sldHorizontal 水平（默认）、1 - sldVertical 垂直 |
| `TickFrequency` | 刻度间隔（按 Value 单位），如 Max=100 设 10 则 11 个刻度 |
| `SelStart / SelEnd / SelectRange` | SelectRange=True 时，SelStart~SelEnd 之间显示高亮区段（如安全范围） |
| `Enabled / Visible / ToolTipText` | 常规属性 |

## 3. 事件

| 事件 | 说明 |
|---|---|
| `Change` | Value 改变即触发（**拖拽过程中连续触发**，实时联动用） |
| `Scroll` | 用户拖动滑块/滚动操作时触发 |
| 常规 | Click、DblClick、KeyDown/KeyUp 等 |

## 4. 示例（实时缩放字号）

```vb
Private Sub Form_Load()
    With Slider1
        .Min = 8: .Max = 72: .Value = 12
        .SmallChange = 1: .LargeChange = 8
        .TickFrequency = 8
    End With
End Sub

Private Sub Slider1_Change()
    Text1.FontSize = Slider1.Value
    lblTip.Caption = "字号：" & Slider1.Value
End Sub
```

> 坑：Change 在拖拽中高频触发，若联动的是耗时操作（重查询、重绘图），改在 Scroll 结束或 MouseUp 里做，或加 Timer 节流。