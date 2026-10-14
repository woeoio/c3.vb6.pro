# TabStrip 控件（选项卡条）

> 部件：Microsoft Windows Common Controls 6.0 (SP6)，MSCOMCTL.OCX

## 1. 概述
只画"标签条"本身的选项卡控件，**不是容器**（与 SSTab 的最大区别）：
页面内容要自己用 PictureBox/Frame 摆放，并在 Click 事件里切换显示。
代价换来了：标签可带图标、可多行、布局更自由。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Tabs` | 标签集合（Tab 对象：`Caption`、`Key`、`Index`） |
| `SelectedItem` | 当前选中的 Tab 对象（用 .Index 或 .Key 判断） |
| `ClientLeft / ClientTop / ClientWidth / ClientHeight` | **关键**：标签条下方可供放页面内容的矩形区域（缇） |
| `ImageList` | 关联 ImageList 后，Tab.Image 可给标签加图标 |
| `MultiRow` | True 时标签过多自动排成多行 |
| `HotTracking` | True 时鼠标悬停标签高亮 |
| `Enabled / Visible / ToolTipText` | 常规属性 |

## 3. 事件
`Click` —— **核心**：点击标签后触发，随后读 SelectedItem 切换页面；
另有 DblClick、KeyDown/KeyPress/KeyUp、OLEDragDrop 系列等常规事件。

## 4. 与 SSTab 的取舍

| | SSTab | TabStrip |
|---|---|---|
| 容器 | 是，子控件自动随页显隐 | 否，需自己切换页面容器 |
| 标签图标 | 不支持 | 支持（ImageList） |
| 多行标签 | TabsPerRow | MultiRow |
| 适用 | 快速做设置对话框 | 主界面多页签、需图标 |

## 5. 示例（PictureBox 数组做页面）

```vb
' 窗体上放 TabStrip1 和 picPage(0..2)（PictureBox 数组，BorderStyle=0）
Private Sub Form_Load()
    TabStrip1.Tabs.Add , "t1", "常规", 1      ' 第4参为 ImageList 图标索引
    TabStrip1.Tabs.Add , "t2", "视图", 2
    TabStrip1.Tabs.Add , "t3", "高级", 3
    Dim i As Integer
    For i = 0 To 2
        picPage(i).Move TabStrip1.Left + TabStrip1.ClientLeft, _
                        TabStrip1.Top + TabStrip1.ClientTop, _
                        TabStrip1.ClientWidth, TabStrip1.ClientHeight
        picPage(i).Visible = (i = 0)
    Next
End Sub

Private Sub TabStrip1_Click()
    Dim i As Integer
    For i = 0 To 2
        picPage(i).Visible = (i = TabStrip1.SelectedItem.Index - 1)
    Next
End Sub
```

> 坑：Client* 坐标是相对 TabStrip 的，Move 页面时要加上 TabStrip.Left/Top；页面容器建议用 PictureBox（可再嵌套控件）并设 BorderStyle=0 视觉无缝。