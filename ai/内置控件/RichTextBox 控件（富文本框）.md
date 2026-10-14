# RichTextBox 控件（富文本框）

> 部件：Microsoft Rich TextBox Control 6.0 (SP6)，RICHTX32.OCX

## 1. 概述
支持富文本格式（RTF）的编辑控件：分段字体、字号、颜色、加粗斜体、对齐、
缩进、插入图片，可读写 .rtf / .txt 文件，相当于迷你 Word 编辑器。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Text` | 纯文本内容 |
| `TextRTF` | 带格式的 RTF 内容（字符串） |
| `SelStart / SelLength / SelText` | 选区起点 / 长度 / 选中文本 |
| `SelBold / SelItalic / SelUnderline / SelStrikethru` | 选区加粗 / 斜体 / 下划线 / 删除线（Null 表示混合状态） |
| `SelColor / SelFontName / SelFontSize` | 选区颜色 / 字体 / 字号 |
| `SelAlignment` | 选区段落对齐：0 左、1 中、2 右 |
| `SelIndent / SelRightIndent / SelHangingIndent` | 左缩进 / 右缩进 / 悬挂缩进 |
| `ScrollBars` | 0 无、1 水平、2 垂直、3 两者 |
| `WordWrap` | 自动换行 |
| `ReadOnly` | 只读 |
| `BorderStyle` | 0 无边框、1 单线 |
| `MaxLength` | 最大字符数（0 不限） |

> 规律：所有 Sel* 属性"有选区时改选区，无选区时改插入点之后输入的新字"。

## 3. 事件
`Change`、`SelChange`（选区/插入点变化——**更新工具栏按钮状态的关键**）、
`Click`、`DblClick`、`KeyDown/KeyPress/KeyUp`、`MouseDown/Move/Up`、
`GotFocus/LostFocus`、`OLECompleteDrag` 等。

## 4. 方法

| 方法 | 说明 |
|---|---|
| `LoadFile(path, [filetype])` | 载入文件：0 - rtfRTF、1 - rtfText |
| `SaveFile(path, [filetype])` | 保存文件，同上 |
| `Find(text, [start], [end], [flags])` | 查找，返回位置或 -1；flags：rtfWholeWord 整词、rtfMatchCase 区分大小写 |
| `SelPrint(hdc)` | 打印选区/全文（传 Printer.hDC） |
| `GetLineFromChar / GetFirstVisible` | 行号辅助 |

## 5. 示例（迷你编辑器核心代码）

```vb
' 加粗切换（Null=混合状态时视为不加粗）
Private Sub cmdBold_Click()
    RichTextBox1.SelBold = Not (RichTextBox1.SelBold = True)
    RichTextBox1.SetFocus
End Sub

' 工具栏状态随选区刷新
Private Sub RichTextBox1_SelChange()
    cmdBold.Value = (RichTextBox1.SelBold = True)
    cmdItalic.Value = (RichTextBox1.SelItalic = True)
End Sub

' 打开 / 保存
Private Sub cmdOpen_Click()
    With CommonDialog1
        .Filter = "RTF 文件 (*.rtf)|*.rtf|文本文件 (*.txt)|*.txt"
        .ShowOpen
        RichTextBox1.LoadFile .FileName, IIf(.FilterIndex = 1, 0, 1)
    End With
End Sub

' 查找
Private Sub cmdFind_Click()
    Dim p As Long
    p = RichTextBox1.Find(txtKey.Text, RichTextBox1.SelStart + RichTextBox1.SelLength)
    If p = -1 Then MsgBox "未找到" Else RichTextBox1.SelStart = p: RichTextBox1.SelLength = Len(txtKey.Text)
End Sub
```

> 坑：SelBold 等是三态（True/False/Null），判断要用 `= True` 而非直接 If；打印用 SelPrint(Printer.hDC) 后记得 Printer.EndDoc。