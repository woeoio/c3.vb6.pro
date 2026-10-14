# MaskEdBox 控件（掩码编辑框）

> 部件：Microsoft Masked Edit Control 6.0 (SP6)，MSMASK32.OCX

## 1. 概述
带"输入掩码"的文本框：用掩码字符串规定每一位允许输入什么，
非法按键直接被拒绝，适合电话、日期、编号、车牌等**固定格式**录入。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Mask` | 掩码字符串（核心），字符含义见 §3 |
| `Text` | 控件文本（含掩码中的固定符号与提示符，受 PromptInclude 影响） |
| `ClipText` | **只含用户实际输入的字符**（去掉固定符号与提示符）——存库请用这个 |
| `PromptChar` | 空位提示符，默认 "_" |
| `PromptInclude` | 决定 Text 是否包含提示符与固定符号 |
| `SelStart / SelLength / SelText` | 选区操作 |
| `ReadOnly / Enabled / BackColor / ForeColor / Font / BorderStyle` | 常规属性 |

## 3. 掩码字符表

| 字符 | 含义 |
|---|---|
| `0` | 数字 0-9，必填，不接受正负号 |
| `9` | 数字或空格，可选 |
| `#` | 数字、空格或正负号，可选 |
| `L` | 字母，必填 |
| `?` | 字母，可选 |
| `A` | 字母或数字，必填 |
| `a` | 字母或数字，可选 |
| `&` | 任意字符，必填 |
| `C` | 任意字符或空格，可选 |
| `. , : ; - /` | 日期/时间分隔符（按系统区域显示） |
| `<` / `>` | 其后字符强制小写 / 大写 |
| `\` | 转义：把下一个字符当普通符号显示 |

## 4. 事件
无专属事件：Change、Click、DblClick、GotFocus/LostFocus、
KeyDown/KeyPress/KeyUp、MouseDown/Move/Up 等常规事件。
（格式校验由掩码在按键级自动完成，无需自己写校验。）

## 5. 示例

```vb
MaskEdBox1.Mask = "(999) 000-0000"     ' 电话：区号可选、号码必填
MaskEdBox2.Mask = "00/00/0000"         ' 日期
MaskEdBox3.Mask = ">LLL-000000"        ' 如 ABC-123456，字母自动转大写

' 存库 / 回显
rs!电话 = MaskEdBox1.ClipText          ' 只存用户输入的数字
MaskEdBox1.Text = rs!电话              ' 回显时自动套上掩码格式
```

> 坑：掩码是**定长**的，长度可变的输入（如姓名、地址）请改用 TextBox + LostFocus 校验；未填满时 ClipText 可能含空格，存前 Trim 一下。