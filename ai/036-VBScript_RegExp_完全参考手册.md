# VBScript.RegExp 完全参考手册

> **方法 · 属性 · 正则语法 · 实战配方 · 避坑指南**
>
> 适用对象：`VBScript.RegExp`（COM 组件，ProgID `VBScript.RegExp` / `VBScript_RegExp_55.RegExp`）
> 适用宿主：Windows Script Host（`.vbs` / `.wsf`）、Classic ASP、VBA / VB6（引用 *Microsoft VBScript Regular Expressions 5.5*）、HTA、各类支持 COM 的脚本环境（Rhino、DIAdem、SecureCRT、QTP/UFT 等）
> 引擎版本：**IE 5.5 正则引擎**（随 IE 5.5 发布后从未升级）。即使系统装了 IE11 或 Windows 11，VBScript 用的仍是这套老引擎。
> 文档日期：2026-11-16

---

## 目录

1. [对象模型总览](#1-对象模型总览)
2. [创建对象与最小可用示例](#2-创建对象与最小可用示例)
3. [RegExp 的 4 个属性](#3-regexp-的-4-个属性)
4. [RegExp 的 3 个方法](#4-regexp-的-3-个方法)
5. [MatchCollection / Match / SubMatches](#5-matchcollection--match--submatches)
6. [替换字符串语法（$1、$&、$$）](#6-替换字符串语法1)
7. [正则语法速查（IE 5.5 方言）](#7-正则语法速查ie-55-方言)
8. [支持 / 不支持特性对照表](#8-支持--不支持特性对照表)
9. [实战配方 40 例](#9-实战配方-40-例)
10. [中文与 Unicode 处理](#10-中文与-unicode-处理)
11. [多行文本处理](#11-多行文本处理)
12. [没有回调替换？用 GetRef 自己造一个](#12-没有回调替换用-getref-自己造一个)
13. [通用工具函数库](#13-通用工具函数库)
14. [常见错误与排查](#14-常见错误与排查)
15. [性能与工程实践](#15-性能与工程实践)
16. [与其他语言正则的差异对照](#16-与其他语言正则的差异对照)
17. [完整可运行示例](#17-完整可运行示例)
18. [一页速查卡](#18-一页速查卡)
19. [参考资料与环境提示](#19-参考资料与环境提示)

---

## 1. 对象模型总览

VBScript 的正则功能由 **4 个对象**构成，全部在 `vbscript.dll` 中实现，Windows 自带，无需注册、无需安装。

```
RegExp                    ← 你唯一需要 CreateObject 的对象
 ├─ 属性: Pattern / Global / IgnoreCase / MultiLine
 ├─ 方法: Test() / Execute() / Replace()
 │
 └─ Execute() 返回 → MatchCollection      （只读集合）
                      ├─ Count
                      └─ Item(i) → Match   （只读对象）
                                    ├─ FirstIndex   0 基起始位置
                                    ├─ Length       匹配长度
                                    ├─ Value        匹配文本（默认属性）
                                    └─ SubMatches → SubMatches 集合（字符串集合）
                                                     ├─ Count
                                                     └─ Item(i)  第 i+1 个捕获组
```

| 对象 | 属性 | 方法 | 说明 |
|---|---|---|---|
| `RegExp` | `Pattern`、`Global`、`IgnoreCase`、`MultiLine` | `Test`、`Execute`、`Replace` | 可读写，可复用 |
| `MatchCollection` | `Count`、`Item`（默认） | 无 | 只读，`Execute` 的返回值快照 |
| `Match` | `FirstIndex`、`Length`、`Value`（默认）、`SubMatches` | 无 | 只读，单次匹配结果 |
| `SubMatches` | `Count`、`Item`（默认） | 无 | 只读，**只含捕获组文本**，不含整体匹配 |

> ⚠️ **三个最容易记错的点**
> 1. `SubMatches(0)` 是**第 1 个捕获组**，不是整体匹配（整体匹配用 `Match.Value`）。
> 2. `FirstIndex` 是 **0 基**，而 VBScript 的 `Mid()` 是 **1 基**，混用必须 ±1。
> 3. 没有 `Groups`、没有 `Success`、没有 `Matches` 属性——那些是 .NET / VB6 `RegExp` 包装类的东西。

---

## 2. 创建对象与最小可用示例

### 2.1 三种创建方式

```vbscript
' 方式一： late binding，最通用（ASP / WSH / VBA 都可）
Dim re
Set re = CreateObject("VBScript.RegExp")

' 方式二： 带版本号的 ProgID，等价
Set re = CreateObject("VBScript_RegExp_55.RegExp")

' 方式三： New 关键字（VBScript 5.5+ 支持；VBA 中需先勾选引用）
Dim re2
Set re2 = New RegExp
```

Classic ASP 中建议用 `Server.CreateObject("VBScript.RegExp")`（走 MTS 上下文）。

### 2.2 30 秒上手

```vbscript
Dim re, matches, m
Set re = CreateObject("VBScript.RegExp")
re.Pattern    = "\d+"       ' 正则：一个或多个数字
re.Global     = True        ' 找全部，而不是只找第一个
re.IgnoreCase = False       ' 区分大小写
re.MultiLine  = False       ' ^ $ 只匹配整串首尾

' ① 判断有没有
WScript.Echo re.Test("订单号 A12345")        ' True

' ② 提取全部
Set matches = re.Execute("订单号 A12345，金额 678 元")
WScript.Echo matches.Count                   ' 2
For Each m In matches
    WScript.Echo m.Value & " @ " & m.FirstIndex & " 长度 " & m.Length
Next
' 12345 @ 5 长度 5
' 678 @ 14 长度 3

' ③ 替换
WScript.Echo re.Replace("a1b22c333", "#")    ' a#b#c#
```

### 2.3 VBScript 字符串 ≠ 正则字符串（头号新手坑）

**VBScript 的字符串字面量里反斜杠没有转义含义**，所以正则怎么写就怎么写，**不要双写反斜杠**：

```vbscript
re.Pattern = "\d+\.\d+"      ' ✅ 正确：匹配小数
re.Pattern = "\\d+\\.\\d+"   ' ❌ 错误：这是在匹配字面量 "\d+\.\d+"
```

需要双写的只有**双引号**：

```vbscript
re.Pattern = """([^""]*)"""  ' 正则其实是 "([^"]*)"，匹配双引号内的内容
' 也可以用 Chr(34) 拼接，可读性更好：
re.Pattern = Chr(34) & "([^" & Chr(34) & "]*)" & Chr(34)
```

---

## 3. RegExp 的 4 个属性

全部**可读可写**，随时可改；改完立即对下一次 `Test/Execute/Replace` 生效。

| 属性 | 类型 | 默认值 | 作用 | 对应 JS 标志 |
|---|---|---|---|---|
| `Pattern` | String | `""` | 正则表达式本体 | `/.../` 中的内容 |
| `Global` | Boolean | `False` | 是否查找/替换**全部**匹配 | `g` |
| `IgnoreCase` | Boolean | `False` | 是否忽略大小写 | `i` |
| `MultiLine` | Boolean | `False` | `^` `$` 是否也匹配行首行尾 | `m` |

### 3.1 Pattern

* **赋值即编译**。模式非法时，错误抛在 `re.Pattern = ...` **这一行**，而不是调用方法时——排错时先看这一行。
* 默认值 `""`（空模式）能匹配任意字符串的开头位置，`Test` 恒为 `True`；配合 `Global=True` 会在每个字符缝隙产生空匹配。
* 同一个对象可以反复改 `Pattern` 复用，但每次赋值都会重新编译，**热循环里不要把 Pattern 写在循环体内**。

```vbscript
re.Pattern = "^\d{4}-\d{2}-\d{2}$"   ' 编译发生在这一行
re.Pattern = "[A-Z]+"                ' 重新编译，之前的模式被丢弃
```

### 3.2 Global

| Global | `Test` | `Execute` | `Replace` |
|---|---|---|---|
| `False`（默认） | 不受影响 | 只返回**第一个**匹配（Count ≤ 1） | 只替换**第一处** |
| `True` | 不受影响 | 返回**全部**匹配 | 替换**所有**匹配 |

```vbscript
re.Pattern = "o"
re.Global = False
WScript.Echo re.Replace("foo boo", "0")   ' f0o boo
re.Global = True
WScript.Echo re.Replace("foo boo", "0")   ' f00 b00
```

> 💡 `Test` 永远只回答"有没有"，`Global` 对它没有任何影响。

### 3.3 IgnoreCase

* `True` 时 `a` 与 `A` 等价，字符类 `[a-z]` 也会匹配大写。
* 对 **ASCII 字母**行为可靠；对非 ASCII 字符（如带重音字母、希腊/西里尔字母）的大小写折叠**不做保证**，跨语言场景请显式写出两种写法 `[a-zA-Z]` / `[а-яА-Я]`。
* VBScript **不支持**内联修饰符 `(?i)`，只能靠这个属性。

```vbscript
re.Pattern = "^color$"
re.IgnoreCase = True
WScript.Echo re.Test("CoLoR")   ' True
```

### 3.4 MultiLine

* `False`（默认）：`^` 只匹配**整个字符串**的开头，`$` 只匹配**整个字符串**的结尾。
* `True`：`^` 额外匹配 `\n` 或 `\r` **之后**的位置，`$` 额外匹配 `\n` 或 `\r` **之前**的位置。
* **它不改变 `.` 的行为**！VBScript 没有 dotall / `s` 标志，`.` 永远不匹配换行符。跨行匹配请用 `[\s\S]`。

```vbscript
Dim text
text = "apple" & vbCrLf & "banana" & vbCrLf & "cherry"

re.Pattern = "^\w+$"
re.Global = True
re.MultiLine = False
WScript.Echo re.Execute(text).Count   ' 0  ← 整串含换行，\w+ 无法覆盖

re.MultiLine = True
WScript.Echo re.Execute(text).Count   ' 3  ← 每行各一个
```

> 📌 **关于"VBScript 没有 MultiLine"的传言**
> 早期 MSDN 的《RegExp Object (VBScript)》页面只列了 3 个属性，导致网上大量文章声称 VBScript 不支持多行。
> 实际上 **IE 5.5 引擎（IRegExp2）提供了 `MultiLine`**，微软官方文档《Regular Expression Syntax (Scripting)》明确写着
> "If the **Multiline** property is set, `^` also matches the position following `\n` or `\r`"，
> 《Regular Expression Programming (Scripting)》的 Flags 表也把 JScript 的 `m` 映射到 VBScript 的 `Multiline` 属性。
> 结论：**在现代 Windows 上放心用 `re.MultiLine = True`**；只有在极古老的 5.0 引擎上才会报"对象不支持该属性或方法"。

---

## 4. RegExp 的 3 个方法

| 方法 | 签名 | 返回值 | 用途 |
|---|---|---|---|
| `Test` | `re.Test(string)` | `Boolean` | 校验：有没有匹配 |
| `Execute` | `re.Execute(string)` | `MatchCollection` | 提取：匹配了什么、在哪里 |
| `Replace` | `re.Replace(string1, string2)` | `String` | 替换：返回替换后的**新字符串** |

### 4.1 Test —— 校验

```vbscript
Boolean = re.Test(subjectString)
```

* 只要字符串中**任意位置**存在匹配就返回 `True`。
* **做整串校验时必须加锚点 `^...$`**，否则 `"abc123xyz"` 也会被 `"\d+"` 判为 True。
* 无匹配返回 `False`，不会报错；但传入 `Null` 会抛运行时错误（见 §14）。

```vbscript
Function IsMobile(ByVal s)
    Dim re : Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^1[3-9]\d{9}$"     ' 整串校验，锚点必不可少
    IsMobile = re.Test(s)
End Function

WScript.Echo IsMobile("13812345678")   ' True
WScript.Echo IsMobile("12812345678")   ' False
WScript.Echo IsMobile("tel13812345678")' False ← 有锚点才不会误判
```

### 4.2 Execute —— 提取

```vbscript
Set matches = re.Execute(subjectString)
```

* 返回 `MatchCollection`；**无匹配时返回 `Count = 0` 的集合，而不是 `Nothing`**，所以不要写 `If matches Is Nothing`。
* `Global=False` → 最多 1 个元素；`Global=True` → 全部。
* 集合是**快照**，之后修改 `re.Pattern` 不会影响已取出的集合。
* 可能产生 **零长度匹配**（`Length = 0`），遍历时通常要过滤。

```vbscript
Dim re, ms, m
Set re = CreateObject("VBScript.RegExp")
re.Pattern = "(\w+)@(\w+)\.com"
re.Global = True
Set ms = re.Execute("联系 tom@example.com 或 jerry@test.com")

WScript.Echo ms.Count                 ' 2
WScript.Echo ms(0).Value              ' tom@example.com   （Item 是默认属性）
WScript.Echo ms(0).FirstIndex         ' 3                 （0 基）
WScript.Echo ms(0).Length             ' 15
WScript.Echo ms(0).SubMatches.Count   ' 2
WScript.Echo ms(0).SubMatches(0)      ' tom
WScript.Echo ms(0).SubMatches(1)      ' example
WScript.Echo ms.Item(1).SubMatches(0) ' jerry
```

**零长度匹配的坑：**

```vbscript
re.Pattern = "\d*"      ' 注意是 * 不是 +
re.Global = True
Set ms = re.Execute("a1b22c")
' 结果 6 个匹配： ""(0) "1"(1) ""(2) "22"(3) ""(5) ""(6)
For Each m In ms
    If m.Length > 0 Then WScript.Echo m.Value   ' 只处理非空匹配
Next
```

### 4.3 Replace —— 替换

```vbscript
newString = re.Replace(subjectString, replacementText)
```

* **不修改原字符串**（VBScript 字符串不可变），必须接收返回值。
* `replacementText` 中可使用 `$1`…`$9`、`$&`、`$$`（见 §6）。
* 替换串为 `""` 即为"删除所有匹配"。
* **不支持回调函数**（JScript 的 `replace(re, fn)` 在 VBScript 里没有对应物），需要动态替换见 §12。

```vbscript
Dim re : Set re = CreateObject("VBScript.RegExp")
re.Global = True

re.Pattern = "\s+"
WScript.Echo re.Replace("  太多    空格  ", " ")     ' " 太多 空格 "

re.Pattern = "<[^>]+>"
WScript.Echo re.Replace("<p>你好<b>世界</b></p>", "") ' 你好世界

re.Pattern = "(\w+)\s+(\w+)"
WScript.Echo re.Replace("John Smith", "$2 $1")        ' Smith John

re.Pattern = "\d+"
WScript.Echo re.Replace("a1b22", "[$&]")              ' a[1]b[22]
```

---

## 5. MatchCollection / Match / SubMatches

### 5.1 MatchCollection

| 成员 | 说明 |
|---|---|
| `Count` | 匹配个数，0 表示没匹配上 |
| `Item(index)` | **默认属性**，`index` 从 `0` 到 `Count-1`，返回 `Match` 对象 |

三种遍历方式：

```vbscript
' ① For Each（推荐，最简洁）
For Each m In ms
    WScript.Echo m.Value
Next

' ② 索引循环（需要序号时用）
Dim i
For i = 0 To ms.Count - 1
    WScript.Echo i & ": " & ms.Item(i).Value
Next

' ③ 只取第一个（Global=False 时最省事）
If ms.Count > 0 Then WScript.Echo ms(0).Value
```

### 5.2 Match

| 属性 | 类型 | 说明 |
|---|---|---|
| `FirstIndex` | Long | 匹配在原串中的起始位置，**0 基** |
| `Length` | Long | 匹配的字符数（UTF-16 码元数） |
| `Value` | String | **默认属性**，匹配到的文本 |
| `SubMatches` | Collection | 捕获组文本集合 |

```vbscript
' 默认属性的便利写法
For Each m In ms
    s = s & m            ' 等价于 s = s & m.Value
Next

' 0 基 → 1 基：用 Mid 可以还原出完全相同的文本
WScript.Echo (Mid(subject, m.FirstIndex + 1, m.Length) = m.Value)   ' True
```

**用 FirstIndex 做"按位置切割"：**

```vbscript
' 把 "姓名:张三;年龄:18" 中所有的值取出来
re.Pattern = ":([^;]+)"
re.Global = True
For Each m In re.Execute("姓名:张三;年龄:18")
    WScript.Echo m.SubMatches(0)     ' 张三 / 18
Next
```

### 5.3 SubMatches

| 成员 | 说明 |
|---|---|
| `Count` | **捕获组**的个数（= 模式中 `(` 的个数，不含 `(?:`、`(?=`、`(?!`） |
| `Item(i)` | **默认属性**，第 `i+1` 个捕获组的文本；`i` 从 0 开始 |

关键规则：

1. **`SubMatches(0)` = 第 1 个捕获组**，整体匹配请用 `Match.Value`。
2. 没有捕获组时 `Count = 0`，访问 `SubMatches(0)` 会**下标越界**。
3. 未参与匹配的组（可选分支未命中）返回 **空字符串 `""`**，不是 `Null`、不是 `Empty`。
4. 组被量词重复时，**只保留最后一次迭代**的文本：`(\d)+` 匹配 `"123"` → `SubMatches(0) = "3"`。
5. **拿不到捕获组的位置和长度**（引擎不提供），只能用 `InStr(m.Value, m.SubMatches(0))` 近似定位。
6. `(?:...)`、`(?=...)`、`(?!...)` **不占用**组号。

```vbscript
re.Pattern = "(a)(x)?(b)"
Set ms = re.Execute("ab")
WScript.Echo ms(0).SubMatches.Count   ' 3
WScript.Echo ms(0).SubMatches(0)      ' "a"
WScript.Echo ms(0).SubMatches(1)      ' ""   ← 未参与，空串
WScript.Echo ms(0).SubMatches(2)      ' "b"

re.Pattern = "(?:\d{3})-(\d{4})"      ' 非捕获组不占号
Set ms = re.Execute("010-1234")
WScript.Echo ms(0).SubMatches.Count   ' 1
WScript.Echo ms(0).SubMatches(0)      ' "1234"
```

---

## 6. 替换字符串语法（$1、$&、$$）

`Replace` 的第二个参数不是普通字符串，它有自己的微型语法：

| 记号 | 含义 | 示例（模式 `\d+`，原串 `a12b`） |
|---|---|---|
| `$1` … `$9` | 第 1–9 个**捕获组**的文本 | 模式 `(\d)(\d)` → 替换 `$2$1` 得 `a21b` |
| `$&` | **整个匹配**的文本 | 替换 `[$&]` 得 `a[12]b` |
| `$$` | 一个字面量 `$` | 替换 `$$` 得 `a$b` |
| `$0` | ⚠️ **不是**整体匹配，原样输出 `$0` | 替换 `<$0>` 得 `a<$0>b` |
| `$10` 及以上 | 不支持两位数分组，`$10` = `$1` 后跟字符 `0` | — |
| `` $` `` / `$'` | JScript 中的"匹配前/后文本"，**VBScript 未文档化，不要依赖** | 用 §12 的循环方案代替 |

```vbscript
re.Global = True

re.Pattern = "(\d{4})-(\d{2})-(\d{2})"
WScript.Echo re.Replace("2026-11-16", "$2/$3/$1")   ' 11/16/2026

re.Pattern = "\w+"
WScript.Echo re.Replace("hello world", """$&""")    ' "hello" "world"

re.Pattern = "b"
WScript.Echo re.Replace("abc", "$$")                ' a$c   ← 想输出 $ 必须写 $$
```

> ⚠️ **反向引用 ≠ 替换引用**
> * 模式**内部**引用前面的组：`\1`、`\2` …（如 `\b(\w+)\s+\1\b` 找重复单词）
> * 替换串**里面**引用组：`$1`、`$2` …
> 把替换串写成 `"\1"` 是新手最常见的错误之一，它只会输出字面量 `\1`。

> ⚠️ **替换文本来自用户输入时务必转义 `$`**
> 否则用户输入的 `$&`、`$1` 会被当成引用。转义方法：`Replace(userText, "$", "$$")`（这里用的是 VBScript 内置的 `Replace` 函数，不是正则）。

---

## 7. 正则语法速查（IE 5.5 方言）

### 7.1 字符与转义

| 写法 | 含义 |
|---|---|
| `x` | 普通字符本身 |
| `\.` `\*` `\+` `\?` `\(` `\)` `\[` `\]` `\{` `\}` `\|` `\^` `\$` `\\` | 转义元字符，匹配字面量 |
| `.` | 除换行符外的任意单个字符（**没有 dotall**，跨行用 `[\s\S]`） |
| `\n` `\r` `\t` `\f` `\v` | 换行 / 回车 / 制表 / 换页 / 垂直制表 |
| `\xHH` | 十六进制 ASCII，**必须恰好 2 位**：`\x41` = `A` |
| `\uHHHH` | Unicode 码元，**必须恰好 4 位**：`\u00A9` = `©` |
| `\cX` | 控制字符，`X` 为 A–Z / a–z：`\cM` = 回车 |
| `\0nn` | 八进制转义（与反向引用有歧义，**建议改用 `\xHH`**） |
| `\1`…`\9` | 反向引用，引用前面第 n 个捕获组已匹配的文本 |

### 7.2 字符类

| 写法 | 等价 / 含义 |
|---|---|
| `\d` / `\D` | `[0-9]` / `[^0-9]`　**仅 ASCII 数字** |
| `\w` / `\W` | `[A-Za-z0-9_]` / `[^A-Za-z0-9_]`　**不含中文、不含重音字母** |
| `\s` / `\S` | `[ \t\n\r\f\v]` / 其补集 |
| `[abc]` | a 或 b 或 c |
| `[^abc]` | 除 a、b、c 外的任意字符（**包含换行**） |
| `[a-z0-9_]` | 范围可叠加 |
| `[\u4e00-\u9fa5]` | 常用汉字区间 |
| `[.-/]` | 类内多数元字符失去特殊含义；`]` `\` `^`(首位) `-`(中间) 仍需注意位置或转义 |

> 类内小技巧：`[-a-z]` 或 `[a-z-]` 中的 `-` 放在首/尾即字面量；`[\b]` 在类内表示**退格符**而非单词边界。

### 7.3 量词（贪婪 / 懒惰）

| 贪婪 | 懒惰 | 含义 |
|---|---|---|
| `*` | `*?` | 0 次或多次 |
| `+` | `+?` | 1 次或多次 |
| `?` | `??` | 0 次或 1 次 |
| `{n}` | `{n}?` | 恰好 n 次 |
| `{n,}` | `{n,}?` | 至少 n 次 |
| `{n,m}` | `{n,m}?` | n 到 m 次（**逗号后不能有空格**） |

```vbscript
' 贪婪 vs 懒惰
re.Pattern = "<b>.*</b>"    ' 匹配 "<b>one</b><b>two</b>" 整段
re.Pattern = "<b>.*?</b>"   ' 分别匹配 "<b>one</b>" 和 "<b>two</b>"
```

> ❌ 不支持占有量词 `*+` `++` `?+`，写了会报 *Unexpected quantifier*。
> ❌ 不支持 `{,m}` 这种省略下界的写法。

### 7.4 锚点与边界

| 写法 | 含义 |
|---|---|
| `^` | 串首；`MultiLine=True` 时也匹配 `\n`/`\r` 之后 |
| `$` | 串尾；`MultiLine=True` 时也匹配 `\n`/`\r` 之前 |
| `\b` | 单词边界（基于 `\w`，**对中文无效**） |
| `\B` | 非单词边界 |
| ❌ `\A` `\Z` `\z` `\G` | **不支持**，用 `^` `$` 代替 |

### 7.5 分组与断言

| 写法 | 含义 | 是否占组号 |
|---|---|---|
| `(pattern)` | 捕获组，进入 `SubMatches` | ✅ |
| `(?:pattern)` | 非捕获组，仅分组 | ❌ |
| `(?=pattern)` | 正向先行断言（右侧必须能匹配） | ❌ |
| `(?!pattern)` | 负向先行断言（右侧必须不能匹配） | ❌ |
| `a\|b` | 分支，优先级最低 | — |
| ❌ `(?<=)` `(?<!)` | **后行断言完全不支持** | — |
| ❌ `(?<name>)` `(?'name')` | **命名分组不支持** | — |
| ❌ `(?i)` `(?m)` `(?s)` | **内联修饰符不支持**，用属性 | — |
| ❌ `(?>)` | **原子组不支持** | — |
| ❌ `(?(1)yes\|no)` | **条件表达式不支持** | — |
| ❌ `(?#comment)` | **模式内注释不支持**，注释写在 VBScript 代码里 | — |

**断言的实战价值（零宽，不消耗字符）：**

```vbscript
' 密码：至少 8 位，且同时含大写、小写、数字
re.Pattern = "^(?=.*[a-z])(?=.*[A-Z])(?=.*\d).{8,}$"

' 千分位：在每个"右边还剩 3 的倍数位数字"的位置插逗号
re.Pattern = "(\d)(?=(\d{3})+(?:\.|$))"
re.Global = True
WScript.Echo re.Replace("1234567.89", "$1,")   ' 1,234,567.89

' 找出所有不以 th 开头的单词
re.Pattern = "\b(?!th)\w+\b"
```

### 7.6 运算符优先级（高 → 低）

```
\  转义
()  (?:)  (?=)  []        分组与字符类
*  +  ?  {n}  {n,}  {n,m} 量词
^  $  \b  序列             锚点与连接
|                          分支
```

所以 `m|food` 匹配 `m` 或 `food`；要匹配 `mood`/`food` 必须写 `(m|f)ood`。

---

## 8. 支持 / 不支持特性对照表

| 特性 | VBScript.RegExp | 备注 / 替代方案 |
|---|:---:|---|
| 字符类 `\d \w \s` 及取反 | ✅ | ASCII 语义 |
| 贪婪 / 懒惰量词 | ✅ | `*?` `+?` `{n,m}?` |
| 捕获组 / 非捕获组 | ✅ | `( )` / `(?: )` |
| 反向引用 `\1`–`\9` | ✅ | 模式内 |
| 替换引用 `$1`–`$9` `$&` `$$` | ✅ | `$0` 无效 |
| 正向 / 负向**先行**断言 | ✅ | `(?=)` `(?!)` |
| **后行**断言 | ❌ | 改用捕获组 + `Execute` 手工判断 |
| 命名分组 | ❌ | 用编号，常量命名 `Const G_NAME = 0` |
| 内联修饰符 `(?i)` | ❌ | `IgnoreCase` 属性 |
| 多行模式 | ✅ | `MultiLine` 属性 |
| 单行 / dotall 模式 | ❌ | 用 `[\s\S]` 代替 `.` |
| `\A` `\Z` `\z` `\G` | ❌ | 用 `^` `$` |
| 原子组 / 占有量词 | ❌ | 重写模式避免回溯 |
| 条件表达式 / 模式内注释 | ❌ | 代码里拆成多个正则 |
| Unicode 属性 `\p{Han}` | ❌ | 用 `[\u4e00-\u9fa5]` 等码位区间 |
| `\uHHHH` 码元转义 | ✅ | BMP 内可靠；增补平面是代理对 |
| 忽略大小写覆盖 Unicode | ⚠️ | 仅 ASCII 保证 |
| 回调替换（函数式 replace） | ❌ | 见 §12 `GetRef` 方案 |
| 匹配超时 / 中断 | ❌ | 自己避免灾难性回溯 |
| 捕获组的位置 / 长度 | ❌ | 只有整体 `FirstIndex`/`Length` |
| 从指定位置开始搜索 | ❌ | 先 `Mid()` 截断，再自行加偏移 |
| 反向搜索 / RightToLeft | ❌ | 反转字符串或从后遍历匹配集合 |
| `Split` / `Matches` 便捷方法 | ❌ | 见 §13 工具函数 |

---

## 9. 实战配方 40 例

> 以下模式均为 **VBScript 字符串写法**（反斜杠不双写），可直接赋给 `re.Pattern`。
> 校验类默认 `Global=False`；提取/替换类记得 `Global=True`。

### 9.1 校验类（配 `Test`，务必带 `^...$`）

| # | 用途 | 模式 |
|---|---|---|
| 1 | 中国大陆手机号 | `^1[3-9]\d{9}$` |
| 2 | 固话（区号-号码-分机） | `^0\d{2,3}-?\d{7,8}(-\d{1,6})?$` |
| 3 | 邮箱（实用版） | `^[\w.-]+@[\w-]+(\.[\w-]+)+$` |
| 4 | 18 位身份证（含日期合法性） | `^[1-9]\d{5}(19\|20)\d{2}(0[1-9]\|1[0-2])(0[1-9]\|[12]\d\|3[01])\d{3}[\dXx]$` |
| 5 | IPv4 | `^((25[0-5]\|2[0-4]\d\|1\d{2}\|[1-9]?\d)\.){3}(25[0-5]\|2[0-4]\d\|1\d{2}\|[1-9]?\d)$` |
| 6 | URL | `^(https?\|ftp)://[\w.-]+(:\d+)?(/[^\s]*)?$` |
| 7 | 日期 `yyyy-MM-dd` | `^\d{4}-(0[1-9]\|1[0-2])-(0[1-9]\|[12]\d\|3[01])$` |
| 8 | 时间 `HH:mm:ss` | `^([01]\d\|2[0-3]):[0-5]\d:[0-5]\d$` |
| 9 | 强密码（≥8 位含大小写与数字） | `^(?=.*[a-z])(?=.*[A-Z])(?=.*\d).{8,}$` |
| 10 | 用户名（字母开头，6–16 位字母数字下划线） | `^[A-Za-z][A-Za-z0-9_]{5,15}$` |
| 11 | 纯数字 / 整数 / 小数 | `^\d+$` / `^-?\d+$` / `^-?\d+(\.\d+)?$` |
| 12 | 金额（最多两位小数） | `^\d+(\.\d{1,2})?$` |
| 13 | 十六进制颜色 | `^#([0-9a-fA-F]{3}\|[0-9a-fA-F]{6})$` |
| 14 | 中文姓名（2–8 个汉字） | `^[\u4e00-\u9fa5]{2,8}$` |
| 15 | 邮政编码（中国） | `^[1-9]\d{5}$` |
| 16 | 车牌（含新能源） | `^[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领][A-Z][A-Z0-9]{5,6}$` |

```vbscript
' 校验示例
Dim re : Set re = CreateObject("VBScript.RegExp")
re.Pattern = "^((25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)\.){3}(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)$"
WScript.Echo re.Test("192.168.0.1")     ' True
WScript.Echo re.Test("256.1.1.1")       ' False
WScript.Echo re.Test("01.2.3.4")        ' False（不允许前导零）
```

### 9.2 提取类（配 `Execute` + `Global=True`）

| # | 用途 | 模式 | 取值 |
|---|---|---|---|
| 17 | 所有整数 | `\d+` | `m.Value` |
| 18 | 所有数字（含小数） | `\d+(\.\d+)?` | `m.Value` |
| 19 | 双引号内的内容 | `"([^"]*)"` | `SubMatches(0)` |
| 20 | `key="value"` 键值对 | `(\w+)\s*=\s*"([^"]*)"` | `SubMatches(0)` / `(1)` |
| 21 | HTML 超链接地址 | `<a\s[^>]*href\s*=\s*"([^"]*)"[^>]*>` | `SubMatches(0)`，配 `IgnoreCase=True` |
| 22 | HTML 标签 | `<[^>]+>` | `m.Value` |
| 23 | HTML 注释（跨行） | `<!--[\s\S]*?-->` | `m.Value` |
| 24 | 文本中的 IP | `\b(\d{1,3}\.){3}\d{1,3}\b` | `m.Value`（严格校验再套 #5） |
| 25 | 重复单词 | `\b(\w+)\s+\1\b` | `SubMatches(0)`，配 `IgnoreCase=True` |
| 26 | 连续重复字符 | `(.)\1` | `m.Value` |
| 27 | 日期片段 | `\d{4}-\d{2}-\d{2}` | `m.Value` |
| 28 | 中文片段 | `[\u4e00-\u9fa5]+` | `m.Value` |

```vbscript
' 提取所有 href
Dim re, m
Set re = CreateObject("VBScript.RegExp")
re.Pattern = "<a\s[^>]*href\s*=\s*""([^""]*)""[^>]*>"
re.IgnoreCase = True
re.Global = True
For Each m In re.Execute(html)
    WScript.Echo m.SubMatches(0)
Next

' 提取键值对
re.Pattern = "(\w+)\s*=\s*""([^""]*)"""
For Each m In re.Execute("name=""Tom"" age = ""18"" city=""BJ""")
    WScript.Echo m.SubMatches(0) & " = " & m.SubMatches(1)
Next
' name = Tom / age = 18 / city = BJ
```

### 9.3 清洗与格式化（配 `Replace` + `Global=True`）

| # | 用途 | 模式 | 替换串 |
|---|---|---|---|
| 29 | 压缩连续空白为一个空格 | `\s+` | `" "` |
| 30 | 去首尾空白（Trim） | `^\s+\|\s+$` | `""` |
| 31 | 删除所有 HTML 标签 | `<[^>]+>` | `""` |
| 32 | 删除所有空行（配 `MultiLine=True`） | `^\s*$\r?\n` | `""` |
| 33 | 日期格式转换 | `(\d{4})-(\d{2})-(\d{2})` | `$2/$3/$1` |
| 34 | 姓名倒置 | `(\w+)\s+(\w+)` | `$2 $1` |
| 35 | 给匹配加包裹 | `\w+` | `[$&]` |
| 36 | 千分位（含小数） | `(\d)(?=(\d{3})+(?:\.\|$))` | `$1,` |
| 37 | 手机号脱敏 | `(\d{3})\d{4}(\d{4})` | `$1****$2` |
| 38 | 隐藏邮箱用户名 | `([\w.-]{1,2})[\w.-]*(@)` | `$1***$2` |
| 39 | 统一换行符为 `\n` | `\r\n\|\r` | 见下方说明 |
| 40 | 敏感词打码 | `敏感词A\|敏感词B` | `**` |

```vbscript
re.Global = True

' Trim（VBScript 内置 Trim 只去空格，正则可去全部空白）
re.Pattern = "^\s+|\s+$"
WScript.Echo "[" & re.Replace("  " & vbTab & "hello" & vbCrLf, "") & "]"   ' [hello]

' 千分位
re.Pattern = "(\d)(?=(\d{3})+(?:\.|$))"
WScript.Echo re.Replace("1234567.89", "$1,")    ' 1,234,567.89
WScript.Echo re.Replace("1234567", "$1,")       ' 1,234,567

' 手机号脱敏
re.Pattern = "(\d{3})\d{4}(\d{4})"
WScript.Echo re.Replace("联系 13812345678 谢谢", "$1****$2")  ' 联系 138****5678 谢谢

' 统一换行：注意替换串里要放真实字符，不是 "\n"
re.Pattern = "\r\n|\r"
WScript.Echo re.Replace(text, vbLf)
```

> ⚠️ 替换串里的 `"\n"` 是**两个字符**（反斜杠 + n），不是换行！要换行请写 `vbLf` / `vbCrLf`。

---

## 10. 中文与 Unicode 处理

### 10.1 匹配汉字的两种写法

```vbscript
' 写法 A：\u 转义（官方支持，最常用）
re.Pattern = "[\u4e00-\u9fa5]+"

' 写法 B：ChrW 拼接（等价，动态构造或极老宿主时更保险）
re.Pattern = "[" & ChrW(&H4E00) & "-" & ChrW(&H9FA5) & "]+"
```

常用码位区间：

| 区间 | 覆盖内容 |
|---|---|
| `\u4e00-\u9fa5` | CJK 基本汉字（经典写法，20902 字） |
| `\u4e00-\u9fff` | CJK 基本区 + 后续补充（更全） |
| `\u3400-\u4dbf` | 扩展 A（生僻字，如 䶮） |
| `\uf900-\ufaff` | 兼容表意文字 |
| `\u3000-\u303f` | CJK 标点（。、《》等） |
| `\uff00-\uffef` | 全角字符（！？，：等） |

```vbscript
' 提取中文
re.Pattern = "[\u4e00-\u9fa5]+"
re.Global = True
For Each m In re.Execute("abc汉字123中文")
    WScript.Echo m.Value      ' 汉字 / 中文
Next

' 判断是否含中文
re.Pattern = "[\u4e00-\u9fa5]"
WScript.Echo re.Test("hello 世界")   ' True
```

### 10.2 必须知道的 4 条 Unicode 限制

1. **`\w` `\d` `\b` 都是 ASCII 语义**：汉字属于 `\W`，所以 `\b` 在中英文交界处会"误报"边界，**不要用 `\b` 去切中文词**。
2. **`.` 与 `Length` 按 UTF-16 码元计**：增补平面字符（如 𠮷、emoji）占 **2 个码元**，`.` 一次只匹配半个，`m.Length` 也会是 2。处理这类文本要按码元区间写模式，或先用 `AscW` 判断代理对。
3. **没有 `\p{...}` Unicode 属性**，只能用码位区间近似。
4. **`IgnoreCase` 对非 ASCII 不保证**，中文无大小写不受影响，但带重音的拉丁字母要小心。

### 10.3 编码相关的实战提示

* `.vbs` 文件含中文时，保存为 **ANSI(GBK)** 或 **UTF-16 LE with BOM** 最稳；UTF-8 无 BOM 会被 `cscript` 按 ANSI 解析导致乱码。
* 控制台输出乱码可先执行 `chcp 65001`，或改用 `WScript.Echo` 弹窗 / 写文件（`ADODB.Stream` 指定 `Charset="utf-8"`）。
* 正则本身与编码无关——VBScript 内部字符串始终是 UTF-16，乱码通常出在**读文件/输出**环节，不是正则。

---

## 11. 多行文本处理

### 11.1 三种典型需求

```vbscript
Dim text
text = "name=Tom" & vbCrLf & "age=18" & vbCrLf & "city=BJ"

' ① 逐行校验/提取：MultiLine=True
re.Pattern = "^(\w+)=([^\r\n]*)$"
re.Global = True
re.MultiLine = True
For Each m In re.Execute(text)
    WScript.Echo m.SubMatches(0) & " -> " & m.SubMatches(1)
Next
' name -> Tom / age -> 18 / city -> BJ
' 用 [^\r\n]* 而不是 .*：显式排除行尾符，跨引擎最稳妥（$ 在 \r 或 \n 之前匹配）

' ② 跨行匹配：没有 dotall，用 [\s\S]
re.MultiLine = False
re.Pattern = "name=([\s\S]*?)city="      ' 懒惰 + 任意字符
Set ms = re.Execute(text)
WScript.Echo ms(0).SubMatches(0)          ' Tom␍␊age=18␍␊

' ③ 最稳妥：先按行 Split，再逐行处理（可读性最好，也最快）
Dim lines, i, oneLine
lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
re.Pattern = "^(\w+)=([^\r\n]*)$"
re.MultiLine = False
re.Global = False
For i = 0 To UBound(lines)
    Set oneLine = re.Execute(lines(i))
    If oneLine.Count > 0 Then WScript.Echo oneLine(0).SubMatches(0)
Next
```

### 11.2 换行符备忘

| 常量 | 值 | 正则写法 |
|---|---|---|
| `vbCr` | `\r` (0x0D) | `\r` |
| `vbLf` | `\n` (0x0A) | `\n` |
| `vbCrLf` | `\r\n` | `\r\n` |
| 兼容三种行尾 | — | `\r\n\|\r\|\n`（**顺序重要**，先长后短） |

* 匹配"整行内容"：`^[^\r\n]*$` + `MultiLine=True`（显式排除行尾符，比 `^.*$` 更稳妥）。
* 匹配"空行"：`^\s*$` + `MultiLine=True`。
* 删除空行：`^\s*$\r?\n` → `""`（`\s` 含换行，注意可能吞掉相邻空行，必要时用 `^[ \t]*\r?\n`）。

---

## 12. 没有回调替换？用 GetRef 自己造一个

VBScript 的 `Replace` 只能替换成固定模板。想要"每个匹配都动态计算"（大小写转换、全角转半角、数字翻倍、按字典翻译……），标准做法是：**`Execute` 拿到全部匹配 → 按 `FirstIndex` 拼接 → 用 `GetRef` 传入回调函数**。

```vbscript
'==========================================================
' ReReplaceWith：带回调的全局替换
'   text      原字符串
'   pattern   正则
'   ignoreCase 是否忽略大小写
'   isGlobal  是否替换全部（False 则只替换第一个）
'   callback  用 GetRef("函数名") 传入，签名 Function f(m)
'==========================================================
Function ReReplaceWith(ByVal text, ByVal pattern, ByVal ignoreCase, ByVal isGlobal, ByVal callback)
    Dim re, ms, m, result, pos
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = ignoreCase
    re.Global = isGlobal
    Set ms = re.Execute(text)

    result = ""
    pos = 1                                   ' Mid 是 1 基
    For Each m In ms
        result = result & Mid(text, pos, m.FirstIndex - pos + 1) & callback(m)
        pos = m.FirstIndex + m.Length + 1     ' 0 基 → 1 基
    Next
    ReReplaceWith = result & Mid(text, pos)
End Function

'--- 回调示例 1：匹配到的单词转大写 ---
Function UpperMatch(m)
    UpperMatch = UCase(m.Value)
End Function

'--- 回调示例 2：数字翻倍 ---
Function DoubleNum(m)
    DoubleNum = CStr(CDbl(m.Value) * 2)
End Function

'--- 回调示例 3：驼峰转下划线 ---
Function ToSnake(m)
    ToSnake = "_" & LCase(m.Value)
End Function

WScript.Echo ReReplaceWith("hello world", "\w+", True, True, GetRef("UpperMatch"))
' HELLO WORLD
WScript.Echo ReReplaceWith("a1b22", "\d+", False, True, GetRef("DoubleNum"))
' a2b44
WScript.Echo ReReplaceWith("helloWorldFoo", "[A-Z]", False, True, GetRef("ToSnake"))
' hello_world_foo
```

> 💡 这个函数同时解决了另外两个痛点：
> * 替换文本里含 `$` 时不会被误解析（回调返回什么就是什么）；
> * 需要"匹配前/后文本"（JScript 的 `` $` `` / `$'`）时，`pos` 和 `Mid` 就是现成的答案。

**全角转半角**（经典无回调场景）：

```vbscript
Function FullToHalf(m)
    Dim c : c = AscW(m.Value)
    If c < 0 Then c = c + 65536      ' AscW 返回有符号整数，全角字符会是负数
    If c = &H3000 Then
        FullToHalf = " "
    Else
        FullToHalf = ChrW(c - &HFEE0)
    End If
End Function

WScript.Echo ReReplaceWith("ＡＢＣ１２３，你好", "[\uff01-\uff5e\u3000]", False, True, GetRef("FullToHalf"))
' ABC123,你好
```

---

## 13. 通用工具函数库

把下面这段存成 `regexp_lib.vbs`，用 `ExecuteGlobal` 或直接粘贴进你的脚本即可。

```vbscript
'==============================================================
' VBScript RegExp 工具库
'==============================================================

' 快速创建并配置一个 RegExp 对象
Function NewRegEx(ByVal pattern, ByVal isGlobal, ByVal ignoreCase, ByVal multiLine)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern    = pattern
    re.Global     = isGlobal
    re.IgnoreCase = ignoreCase
    re.MultiLine  = multiLine
    Set NewRegEx = re
End Function

' 是否匹配（整串校验请自行在 pattern 里加 ^ $）
Function ReIsMatch(ByVal text, ByVal pattern, ByVal ignoreCase)
    Dim re
    If IsNull(text) Then ReIsMatch = False : Exit Function
    Set re = NewRegEx(pattern, False, ignoreCase, False)
    ReIsMatch = re.Test(CStr(text))
End Function

' 正则替换
Function ReReplace(ByVal text, ByVal pattern, ByVal replacement, ByVal ignoreCase, ByVal isGlobal)
    Dim re
    If IsNull(text) Then ReReplace = "" : Exit Function
    Set re = NewRegEx(pattern, isGlobal, ignoreCase, False)
    ReReplace = re.Replace(CStr(text), replacement)
End Function

' 取第一个匹配的整体文本，无匹配返回 ""
Function ReFirst(ByVal text, ByVal pattern, ByVal ignoreCase)
    Dim re, ms
    ReFirst = ""
    If IsNull(text) Then Exit Function
    Set re = NewRegEx(pattern, False, ignoreCase, False)
    Set ms = re.Execute(CStr(text))
    If ms.Count > 0 Then ReFirst = ms(0).Value
End Function

' 取第一个匹配的第 n 个捕获组（n 从 1 开始），无则返回 ""
Function ReGroup(ByVal text, ByVal pattern, ByVal n, ByVal ignoreCase)
    Dim re, ms
    ReGroup = ""
    If IsNull(text) Then Exit Function
    Set re = NewRegEx(pattern, False, ignoreCase, False)
    Set ms = re.Execute(CStr(text))
    If ms.Count > 0 Then
        If ms(0).SubMatches.Count >= n Then ReGroup = ms(0).SubMatches(n - 1)
    End If
End Function

' 提取全部匹配文本，返回一维数组（无匹配返回空数组）
Function ReExtract(ByVal text, ByVal pattern, ByVal ignoreCase)
    Dim re, ms, m, arr, i
    If IsNull(text) Then ReExtract = Array() : Exit Function
    Set re = NewRegEx(pattern, True, ignoreCase, False)
    Set ms = re.Execute(CStr(text))
    If ms.Count = 0 Then ReExtract = Array() : Exit Function
    ReDim arr(ms.Count - 1)
    i = 0
    For Each m In ms
        arr(i) = m.Value
        i = i + 1
    Next
    ReExtract = arr
End Function

' 正则分割：先把分隔符替换成哨兵字符，再 Split
' 注意：哨兵字符不能出现在原文中；模式不要产生零长度匹配
Function ReSplit(ByVal text, ByVal pattern, ByVal ignoreCase)
    Dim marker
    marker = ChrW(1)
    If IsNull(text) Then ReSplit = Array() : Exit Function
    If InStr(text, marker) > 0 Then Err.Raise 5, , "文本中含哨兵字符，无法分割"
    ReSplit = Split(ReReplace(text, pattern, marker, ignoreCase, True), marker)
End Function

' 统计匹配次数
Function ReCount(ByVal text, ByVal pattern, ByVal ignoreCase)
    Dim re
    If IsNull(text) Then ReCount = 0 : Exit Function
    Set re = NewRegEx(pattern, True, ignoreCase, False)
    ReCount = re.Execute(CStr(text)).Count
End Function

' 转义：把用户输入变成"字面量"模式，可安全嵌入正则
Function ReEscape(ByVal s)
    ReEscape = ReReplace(CStr(s), "[\\^$.|?*+()\[\]{}]", "\$&", False, True)
End Function

' 转义替换串中的 $，避免用户输入被当成 $1/$&
Function ReEscapeReplacement(ByVal s)
    ReEscapeReplacement = Replace(CStr(s), "$", "$$")   ' 内置 Replace，非正则
End Function
```

**使用示例：**

```vbscript
WScript.Echo ReIsMatch("13812345678", "^1[3-9]\d{9}$", False)   ' True
WScript.Echo ReReplace("a1b22c", "\d+", "#", False, True)        ' a#b#c
WScript.Echo ReFirst("订单A12345", "\d+", False)                 ' 12345
WScript.Echo ReGroup("2026-11-16", "(\d{4})-(\d{2})-(\d{2})", 2, False)  ' 11
WScript.Echo Join(ReExtract("a1b22c333", "\d+", False), "|")     ' 1|22|333
WScript.Echo Join(ReSplit("a1b22c333d", "\d+", False), "|")      ' a|b|c|d
WScript.Echo ReCount("a1b22c333", "\d+", False)                  ' 3

' 安全地查找用户输入的字面量
Dim userInput : userInput = "价格(元)*100"
WScript.Echo ReIsMatch("总价 价格(元)*100 元", ReEscape(userInput), False)  ' True
```

---

## 14. 常见错误与排查

### 14.1 错误信息对照

| 错误 | 典型原因 | 修复 |
|---|---|---|
| `Microsoft VBScript regular expression error '800a139a'`<br>*Unexpected quantifier*（运行时错误 5018） | 量词前没有可重复的元素：`"*abc"`、`"+x"`、`"?a"`、`"(*)"`；或用了占有量词 `"a++"`；或 `{n,m}` 中 n>m | 转义字面量 `\*` `\+` `\?`；去掉多余量词 |
| *Syntax error in regular expression* | `[` 或 `(` 未闭合；模式以孤立 `\` 结尾；字符类范围写反 `[z-a]` | 补全括号；`\\` 表示反斜杠 |
| *对象不支持该属性或方法*（438） | 属性名拼错（`.Multiline` 拼成 `.MultiLines`）；用了 .NET 的 `.Matches`/`.Groups`；极老 5.0 引擎无 `MultiLine` | 核对 §1 成员表 |
| *下标越界*（9） | 无捕获组却访问 `SubMatches(0)`；`ms(0)` 但 `Count=0` | 先判 `Count` |
| *类型不匹配*（13）/ *Invalid use of Null*（94） | 给 `Test/Execute/Replace` 传了 `Null`（常见于数据库字段） | `If IsNull(x) Then x = ""`，或 `CStr(x & "")` |
| *变量未定义*（500，`Option Explicit` 下） | 对象变量忘了 `Set` | `Set re = CreateObject(...)` |

### 14.2 逻辑"不报错但结果不对"的 10 个坑

1. **只替换了第一个** → 忘了 `re.Global = True`。
2. **校验放行了非法串** → `Test` 是"部分匹配"，整串校验必须 `^...$`。
3. **反斜杠双写了** → VBScript 里写 `"\d"`，不是 `"\\d"`。
4. **替换串用了 `\1`** → 应该用 `$1`。
5. **`SubMatches(0)` 取到的是第 1 组而不是整体** → 整体用 `m.Value`。
6. **`FirstIndex` 当成 1 基用** → `Mid(s, m.FirstIndex + 1, m.Length)`。
7. **`.` 匹配不到换行** → 用 `[\s\S]`，`MultiLine` 解决不了这个。
8. **`^`/`$` 在多行文本上不生效** → `re.MultiLine = True`。
9. **`\w`/`\b` 对中文失效** → 中文用 `[\u4e00-\u9fa5]`，别用 `\b` 切词。
10. **出现一堆空匹配** → 模式能匹配空串（`\d*`、`a?`、`""`），改用 `+` 或过滤 `m.Length > 0`。

### 14.3 排错三板斧

```vbscript
' ① 先把模式打印出来，肉眼确认转义
WScript.Echo "Pattern = [" & re.Pattern & "]"

' ② 用最小样本二分定位：先测单个字符，再加量词，再加分组
re.Pattern = "\d"      : WScript.Echo re.Test(s)
re.Pattern = "\d+"     : WScript.Echo re.Test(s)
re.Pattern = "^\d+$"   : WScript.Echo re.Test(s)

' ③ 打印每个匹配的位置与长度，看清引擎到底匹配了什么
For Each m In re.Execute(s)
    WScript.Echo "[" & m.Value & "] pos=" & m.FirstIndex & " len=" & m.Length
Next
```

> 🔎 在线调试时，**regex101 选 ECMAScript (JavaScript) 风味**最接近，但要手动排除后行断言、命名分组、`(?i)`、`\A\Z`、dotall 这些 VBScript 没有的特性。RegexBuddy 直接内置 "VBScript" 风味，最精确。

---

## 15. 性能与工程实践

### 15.1 复用对象

```vbscript
' ❌ 每次循环都创建 + 编译
For i = 0 To 10000
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d+"
    ...
Next

' ✅ 循环外创建一次
Set re = CreateObject("VBScript.RegExp")
re.Pattern = "\d+"
re.Global = True
For i = 0 To 10000
    Set ms = re.Execute(arr(i))
    ...
Next
```

### 15.2 避免灾难性回溯（VBScript **没有超时机制**，会直接卡死）

```vbscript
' ❌ 危险：嵌套量词 + 可重叠分支
re.Pattern = "^(a+)+$"          ' 对 "aaaaaaaaaaaaaaaaaaaaaaaaX" 会指数级回溯
re.Pattern = "^(a|a)*$"
re.Pattern = "^(\w+\s?)*$"

' ✅ 改写：消除歧义，让每个字符只有一种归属
re.Pattern = "^a+$"
re.Pattern = "^\w+(\s\w+)*$"
```

**规则**：不要让"量词套量词"且内层能匹配同样的字符；分支之间尽量互斥（用首字符区分）。

### 15.3 其他优化要点

| 建议 | 原因 |
|---|---|
| 能用 `InStr` / 内置 `Replace` / `Split` 就别用正则 | 内置函数快一个数量级 |
| 尽量锚定 `^`，或先用 `InStr` 粗筛再上正则 | 减少无谓的起始位置扫描 |
| 用字符类 `[abc]` 代替分支 `(a\|b\|c)` | 字符类是查表，分支要逐个回溯 |
| 把最可能的分支放前面 | 分支是从左到右尝试 |
| 用 `(?:...)` 代替不需要的捕获组 | 少一份捕获开销 |
| 大文本先 `Split` 成行再逐行匹配 | 避免整串回溯，且便于定位 |
| 不要 `Test` 之后再 `Execute` | 等于匹配两遍，直接看 `ms.Count` |
| 谨慎使用 `[\s\S]*?` | 懒惰 + 全字符集在大文本上开销很大 |
| Classic ASP 中**不要**把 RegExp 对象放进 `Application` | 套间/线程安全问题，按请求创建即可 |

---

## 16. 与其他语言正则的差异对照

| 项目 | VBScript | .NET / C# | JavaScript (现代) | Python |
|---|---|---|---|---|
| 整体匹配 | `m.Value` | `m.Groups[0].Value` | `m[0]` | `m.group(0)` |
| 第 1 个捕获组 | `m.SubMatches(0)` | `m.Groups[1].Value` | `m[1]` | `m.group(1)` |
| 替换引用 | `$1` | `$1` / `${name}` | `$1` | `\1` / `\g<1>` |
| 整体匹配引用（替换） | `$&` | `$&` | `$&` | `\g<0>` |
| 忽略大小写 | `IgnoreCase=True` | `RegexOptions.IgnoreCase` | `/i` | `re.I` |
| 多行 | `MultiLine=True` | `RegexOptions.Multiline` | `/m` | `re.M` |
| dotall | ❌ 用 `[\s\S]` | `RegexOptions.Singleline` | `/s` | `re.S` |
| 后行断言 | ❌ | ✅ | ✅ | ✅ |
| 命名分组 | ❌ | ✅ `(?<n>)` | ✅ | ✅ `(?P<n>)` |
| 回调替换 | ❌（见 §12） | ✅ `MatchEvaluator` | ✅ | ✅ |
| 起始位置 | 0 基 `FirstIndex` | 0 基 `Index` | 0 基 `index` | 0 基 `start()` |
| 匹配超时 | ❌ | ✅ | ❌ | ❌ |

**从别处抄正则到 VBScript 时的"降级"清单：**
1. 删掉 `(?i)` `(?s)` `(?m)` → 改成属性；
2. `(?<=...)` / `(?<!...)` → 改成捕获组 + 代码判断；
3. `(?<name>...)` → 改成普通 `(...)`，用常量记录组号；
4. `\A` `\Z` → `^` `$`；
5. dotall 的 `.` → `[\s\S]`；
6. `\p{Han}` → `[\u4e00-\u9fa5]`；
7. 占有量词 `++` / 原子组 `(?>)` → 重写模式；
8. 替换串里的 `\1` → `$1`。

---

## 17. 完整可运行示例

保存为 `demo.vbs`，命令行执行 `cscript //nologo demo.vbs`。

```vbscript
Option Explicit

Dim re, ms, m, text

'---------- 1. 校验 ----------
Set re = CreateObject("VBScript.RegExp")
re.Pattern = "^1[3-9]\d{9}$"
WScript.Echo "手机号 13812345678 : " & re.Test("13812345678")
WScript.Echo "手机号 12345678901 : " & re.Test("12345678901")

'---------- 2. 提取 + 捕获组 ----------
re.Pattern = "(\w+)@([\w.]+)"
re.Global = True
Set ms = re.Execute("收件人 tom@example.com，抄送 jerry@test.org")
WScript.Echo "共找到 " & ms.Count & " 个邮箱"
For Each m In ms
    WScript.Echo "  " & m.Value & "  用户=" & m.SubMatches(0) & _
                 "  域名=" & m.SubMatches(1) & "  位置=" & m.FirstIndex
Next

'---------- 3. 替换 ----------
re.Pattern = "(\d{3})\d{4}(\d{4})"
re.Global = True
WScript.Echo re.Replace("联系电话 13812345678", "$1****$2")

re.Pattern = "(\d)(?=(\d{3})+(?:\.|$))"
WScript.Echo re.Replace("金额 1234567.89 元", "$1,")

'---------- 4. 多行 ----------
text = "name=Tom" & vbCrLf & "age=18" & vbCrLf & "city=BJ"
re.Pattern = "^(\w+)=([^\r\n]*)$"
re.MultiLine = True
For Each m In re.Execute(text)
    WScript.Echo "  " & m.SubMatches(0) & " -> " & m.SubMatches(1)
Next
re.MultiLine = False

'---------- 5. 中文 ----------
re.Pattern = "[\u4e00-\u9fa5]+"
WScript.Echo "中文片段：" & Join(ExtractAll(re, "abc汉字123中文"), " / ")

'---------- 6. 带回调的替换 ----------
WScript.Echo ReReplaceWith("hello world", "\w+", True, True, GetRef("UpperMatch"))

'---------- 7. 安全转义 ----------
Dim literal : literal = "a.b*c(d)"
re.Pattern = ReEscape(literal)
re.Global = False
WScript.Echo "字面量查找：" & re.Test("前缀 a.b*c(d) 后缀")

'================= 辅助函数 =================
Function ExtractAll(reObj, s)
    Dim col, mm, arr, i
    Set col = reObj.Execute(s)
    If col.Count = 0 Then ExtractAll = Array() : Exit Function
    ReDim arr(col.Count - 1)
    i = 0
    For Each mm In col
        arr(i) = mm.Value : i = i + 1
    Next
    ExtractAll = arr
End Function

Function UpperMatch(mm)
    UpperMatch = UCase(mm.Value)
End Function

Function ReReplaceWith(ByVal s, ByVal pattern, ByVal ic, ByVal g, ByVal cb)
    Dim r, col, mm, result, pos
    Set r = CreateObject("VBScript.RegExp")
    r.Pattern = pattern : r.IgnoreCase = ic : r.Global = g
    Set col = r.Execute(s)
    result = "" : pos = 1
    For Each mm In col
        result = result & Mid(s, pos, mm.FirstIndex - pos + 1) & cb(mm)
        pos = mm.FirstIndex + mm.Length + 1
    Next
    ReReplaceWith = result & Mid(s, pos)
End Function

Function ReEscape(ByVal s)
    Dim r : Set r = CreateObject("VBScript.RegExp")
    r.Global = True
    r.Pattern = "[\\^$.|?*+()\[\]{}]"
    ReEscape = r.Replace(s, "\$&")
End Function
```

**预期输出：**

```
手机号 13812345678 : True
手机号 12345678901 : False
共找到 2 个邮箱
  tom@example.com  用户=tom  域名=example.com  位置=4
  jerry@test.org  用户=jerry  域名=test.org  位置=23
联系电话 138****5678
金额 1,234,567.89 元
  name -> Tom
  age -> 18
  city -> BJ
中文片段：汉字 / 中文
HELLO WORLD
字面量查找：True
```

---

## 18. 一页速查卡

```vbscript
'===== 创建 =====
Set re = CreateObject("VBScript.RegExp")   ' 或 New RegExp

'===== 属性（4 个）=====
re.Pattern    = "\d+"     ' String，默认 ""，赋值即编译
re.Global     = True      ' Boolean，默认 False，全部匹配
re.IgnoreCase = True      ' Boolean，默认 False，忽略大小写
re.MultiLine  = True      ' Boolean，默认 False，^$ 匹配行首行尾

'===== 方法（3 个）=====
b   = re.Test(s)              ' Boolean，有没有（整串校验记得 ^...$）
Set ms = re.Execute(s)        ' MatchCollection，无匹配 Count=0
s2  = re.Replace(s, "$1-")    ' String，$1..$9 / $& / $$

'===== 结果 =====
ms.Count                      ' 匹配个数
ms(i) / ms.Item(i)            ' 第 i 个 Match，i 从 0 开始
m.FirstIndex                  ' 0 基起始位置
m.Length                      ' 长度
m.Value                       ' 匹配文本（默认属性）
m.SubMatches.Count            ' 捕获组个数
m.SubMatches(0)               ' 第 1 个捕获组（不是整体！）

'===== 常用语法 =====
' \d \w \s（ASCII）  . 不含换行 → 跨行用 [\s\S]
' * + ? {n} {n,m}    加 ? 变懒惰
' ^ $ \b \B          无 \A \Z
' ( ) (?: ) (?= ) (?! )      无 (?<= ) (?<! ) (?<name>) (?i)
' \1..\9 模式内引用   $1..$9 替换引用
' \xHH \uHHHH \n \r \t \f \v \cX

'===== 三大坑 =====
' 1. VBScript 字符串里反斜杠不双写： "\d" 而不是 "\\d"
' 2. SubMatches(0) 是第 1 组，整体匹配用 m.Value
' 3. FirstIndex 是 0 基，Mid() 是 1 基
```

---

## 19. 参考资料与环境提示

### 官方文档（Microsoft Learn，Scripting 旧版归档）

* [Introduction to Regular Expressions (Scripting)](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/6wzad2b2(v=vs.84))
* [Regular Expression Syntax (Scripting)](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/1400241x(v=vs.84)) —— 转义表、元字符表、优先级表的权威来源，明确记载 `\uHHHH`、`\xHH`、八进制/反向引用规则，以及 **Multiline 属性对 `^` `$` 的影响**
* [Regular Expression Programming (Scripting)](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/ee236360(v=vs.84)) —— JScript 标志 `g/i/m` ↔ VBScript `Global/IgnoreCase/Multiline` 对照表
* [Backreferences (Scripting)](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/fzhdcs5c(v=vs.84))

### 社区权威资料

* [regular-expressions.info — VBScript's Regular Expression Support](https://www.regular-expressions.info/vbscript.html) —— 引擎限制清单、`SubMatches` 与 .NET 的差异说明最完整
* [Stack Overflow — Multiline REGEX using VB Script](https://stackoverflow.com/questions/29364294/multiline-regex-using-vb-script) —— 确认 `MultiLine` 存在但文档缺失，以及 `[\s\S]` 替代 dotall
* [Stack Overflow — RegExp Replace 中的 `$&`](https://stackoverflow.com/questions/40426065/why-does-my-vbscript-regexp-replace-unexpectedly-fail-when-replacement-contains) —— `$&` 有效、`$$` 转义
* [ComputerPerformance — 错误码 800A139A](https://www.computerperformance.co.uk/error-codes/800a139a/) —— *Unexpected quantifier*

### 环境提示

* **无需安装/注册**：`vbscript.dll` 随 Windows 提供，32 位与 64 位宿主均可 `CreateObject`。
* **VBA/VB6**：可在"引用"中勾选 *Microsoft VBScript Regular Expressions 5.5* 获得早绑定与智能提示；不勾选就用 `CreateObject` 晚绑定，效果相同。
* **VBScript 已进入弃用周期**：微软已宣布 VBScript 的弃用计划——Windows 11 24H2 起以"按需功能（FOD）"形式提供且当前默认启用，**约 2027 年起将不再默认启用**，之后需手动安装 FOD，最终会移除。维护老系统时本文仍然适用；**新项目建议改用 PowerShell（`-match` / `-replace` / `[regex]`，.NET 引擎，支持后行断言、命名分组、超时）**。迁移时对照 §16 的差异表即可。

---

*本手册中的模式语义均按 IE 5.5 / ECMAScript 3 方言核对；`MultiLine`、`$&`、`\uHHHH` 等易被误传的行为已对照微软官方 Scripting 文档与 regular-expressions.info 确认。*
