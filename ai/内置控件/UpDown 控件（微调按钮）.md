# UpDown 控件（微调按钮）

> 部件：Microsoft Windows Common Controls-2 6.0 (SP6)，MSCOMCT2.OCX

## 1. 概述
贴在"伙伴控件"（通常是 TextBox）旁的小箭头按钮，点一下按步长增减数值，
用于数量、字号、年份等整数输入，免去手敲和校验。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `BuddyControl` | 伙伴控件（如 Text1），UpDown 自动贴靠其旁 |
| `BuddyProperty` | 联动的伙伴属性，一般 "Text"（也可为 "Value" 等） |
| `AutoBuddy` | True 时自动把 Tab 序中前一个控件当伙伴 |
| `BuddyAlignment` | 贴靠在伙伴的左侧还是右侧 |
| `Min / Max` | 取值下限 / 上限 |
| `Value` | 当前值 |
| `Increment` | 每点一次的步长 |
| `Orientation` | 0 - 水平排列箭头、1 - 垂直排列箭头 |
| `Wrap` | True 时到顶/到底后回绕 |
| `SyncBuddy` | True 时 Value 变化立即写回伙伴控件 |
| `Enabled / Visible` | 常规属性 |

## 3. 事件

| 事件 | 说明 |
|---|---|
| `Change` | Value 改变时触发（做范围外附加校验/联动常用） |
| `Click` | 点击箭头时触发 |

## 4. 示例

```vb
Private Sub Form_Load()
    Set UpDown1.BuddyControl = Text1        ' 伙伴：Text1
    UpDown1.BuddyProperty = "Text"
    UpDown1.Min = 1: UpDown1.Max = 99
    UpDown1.Increment = 1
    UpDown1.Value = 1
    UpDown1.SyncBuddy = True                ' 自动把值写进 Text1
End Sub

Private Sub UpDown1_Change()
    lblAmount.Caption = "单价 × " & UpDown1.Value & _
                        " = " & Format(Val(txtPrice.Text) * UpDown1.Value, "0.00")
End Sub
```

> 坑：① 伙伴 TextBox 被手敲成非数字时，UpDown 同步会出错——在 Change 或伙伴 LostFocus 里用 Val/IsNumeric 兜底；② 设计期放好位置后运行期它会自动贴靠伙伴，不必手动对齐；③ 只适合整数步长场景，小数/日期步长请改用 DTPicker 或 Slider。