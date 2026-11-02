# UDT 里的对象字段归谁（账 #600 的口径，给真 VB6 / twinBASIC 环境用）

C3 现在把整枚 UDT 赋值（`t2 = t1`）发成逐字节 `memcpy`（`vb6_udtcpy_*`），出过程也不清对象字段 ——
所以 UDT 里的对象字段是**共享指针**。第317刀当时就是把这一格押后的，理由是：
「装订（拷贝时 claim）/ 覆写退旧 / 出厂真清 / getter 交回带对价」这四格是**一条链**，
只做装订会把共享指针变成双重释放，而这条链该怎么修**完全取决于 VB6 那一侧答什么**，本机没有 VB6。

这一份就是去问那一句的。请**照实贴回顺序**，不用解释。

## 为什么是两个工程

`copy/` 那一份问的正是 `t2 = t1`（含对象字段的整枚赋值）—— 这一句在 VB6 里**有可能根本不许编译**，
而「不许编译」本身就是 C3 要的答案。所以把它单拆一份，别让一句编译错误把上一级那份
只碰字段级动作（`Set t.O = …` / 置空 / 出过程 / `Erase` / `ByRef`）的读数一起拖倒。

## 读数怎么读（只看两件事）

1. 每条 `TERM n` 落在哪两条编号标记**之间**（n = 第几枚 `INIT` 出来的实例）。
2. `Z-COUNT init=? term=?` 的差值（差 ≠ 0 = 退出前还有活着的实例）。

`A2`/`G2`/`G4`/`E2` 那几条带 `sameIs=` 的：`Is` 只比对象身份，**一枚引用和一枚 AddRef 过的引用比出来都是
True** —— 所以 `sameIs=True` 不区分「别名」和「独立持有」。区分靠的是紧随其后那一句：
`Set a.O = Nothing` 之后**有没有立刻落一条 `TERM`**。有 ⇒ VB6 交的是别名（不 AddRef）；
没有 ⇒ 拷贝自己持有一枚引用。README 下面的判据都按这个逻辑写。

## 一、`udt_object_field/UdtObjProbe.vbp`（字段级）

| 编号 | 问的是 | 读数怎么定 C3 |
|---|---|---|
| B1..B4 | 覆写 `Set t.O = 新的一枚`，旧的那枚**当场**退不退 | 立刻 `TERM 1` ⇒ C3 的「覆写退旧」必须有；不落 ⇒ 覆写不 Release（B4 才算） |
| C1..C4 | `Set t.O = Nothing` 是否当场退 | 同上，只是清空那一支 |
| D1 / D2 / D3 | 局部 UDT **出过程**时字段里那一枚归谁放，三种声明档各问一次（`As Object` / `As 工程类` / `As Variant`） | `…-after-sub` 前落 `TERM` ⇒ 出厂真清；三档**不同答案** ⇒ C3 出厂要按声明档分叉，不能一套模板 |
| E1..E5 | 动态数组 `Erase arr` / 重新 `ReDim` 对元素字段做了什么 | `E3-after-erase` 前落两条 `TERM` ⇒ `Erase` 走真清；不落 ⇒ VB6 自己就漏（C3 保持现状，别超发） |
| H1..H6 | 从字段**读出来**交给普通对象变量，随后把字段置空：对象还活着吗（`H4-tag-through-g=` 能不能读到值） | H3 前不退、`Set g = Nothing` 后才退 ⇒ getter 交回**带对价**；H4 读到值就是这条的正面证人 |
| M1..M3 | 两个字段指着**同一枚**，出厂时退几次 | `M3-after-sub` 后 `TERM 1` 只落一次 ⇒ 引用计数被 VB6 自己平衡；落两次是 VB6 的 bug，我们不复刻 |
| J1..J6 | `ByRef` 把 UDT 交给子过程，子过程里 `Set u.O = Nothing` | `J4-after-call caller-field-is-Nothing=True` ⇒ ByRef 真按地址（C3 那边同一位主人）；`=False` ⇒ VB6 传的是拷贝，则「出厂/退旧」全在 callee 一侧 |
| N1 | 控制组：新 `Dim` 出来的字段是不是干净的（应为 `True`） | `False` ⇒ 这一份读数全部作废，请连内存现场一起说 |

## 二、`udt_object_field/copy/UdtCopyProbe.vbp`（整枚 UDT 赋值 —— 这一份才是 #600 的主问）

| 编号 | 问的是 | 读数怎么定 C3 |
|---|---|---|
| A2 / A3 | `t2 = t1` 之后 `Set t1.O = Nothing`：**目标那一枚还活着吗** | A3 前**不**落 `TERM` ⇒ VB6 拷贝时 AddRef ⇒ 四格链全做（装订 claim + 覆写退旧 + 出厂真清 + getter 带对价）；A3 前落 `TERM` ⇒ VB6 本身就是别名，**现在的 memcpy 是保真的**，#600 应改口径为「按 VB6 保真」，一条都不做 |
| A4 | `CaseA` 出过程后（t1 字段已空、t2 还指着）退不退 | 落一条 ⇒ 出厂清的是「自己那一枚」，与 A3 那条合起来正好是一次 AddRef/Release 配平 |
| A5 / A6 / A7 | 目标**本来就握着**一枚，整枚赋值要不要先把它放掉（放几条 = 是不是只放对象字段而不连坐） | A6 前落 `TERM 2` ⇒ 覆写退旧在整枚赋值这一路也在；不落 ⇒ VB6 在这里漏，我方跟漏 |
| G2/G3, G4/G5 | 同一问的另外两种声明档（`As Holder` 工程类 / `As Variant`） | 与 A 档**不同**就单独标出来 —— C3 的发码按声明档分叉与否就靠这三行 |
| F1 | 同一条拷贝路径上的**字符串**字段：`b.S` 是 `AAA` 还是 `BBB` | `AAA` = 深拷贝（VB6 字符串字段自己 claim），这是给「同一条赋值语句还兼管字符串」那一半的对照 |
| I1 / I2 | **嵌套** UDT（对象字段在内层）走不走同一条 claim | I2 的 `TERM` 行为 ≠ A 档 ⇒ 发码要递归到内层，不能只看一层字段表 |
| E2 / E3 / E4 | 数组元素之间 `arr(1) = arr(0)` | 与标量那一问同判据；不同 ⇒ 数组那一支单独一格 |

## 三、`copy/RejectCases.bas`（**默认不在 .vbp 里**，一次加一格）

Project → Add File 选它，编译；编不过就把 IDE **原话**抄在对应编号上，能编的把输出贴回来。

- `R1` UDT 形参能不能 `ByVal`（`TakeByVal(ByVal u As RObject)`）
- `R2` `Function` 能不能返回 UDT（`GiveUDT() As RObject` + `v = GiveUDT()`）
- `R3` `ReDim Preserve arr(1)` **缩小**含对象字段的 UDT 数组：被丢掉那两格的对象归谁放
  （这格能编译 ⇒ 读数看 `R3-before-shrink` 与 `R4-after-shrink` 之间落几条 `TERM`）
- `R4` `Let u.O = h` 是哪一条错误

## 步骤

1. VB6 打开 `ai/probes/udt_object_field/UdtObjProbe.vbp`，**F5** 跑（不用 Make —— `App.Path` 在 IDE 里就是
   工程目录，输出落 `probe_out.txt` 在那儿）。跑完弹窗告诉你 init/term 计数。
2. VB6 打开 `ai/probes/udt_object_field/copy/UdtCopyProbe.vbp`，同样 **F5**，输出 `probe_copy_out.txt`。
3. 整段贴回两份文本（或立即窗口 `Debug.Print` 的流水，顺序一致就行）。
4. `twinBASIC` 里同样跑一遍（新建 Standard EXE，把同名三个文件加进去），读数**前面标 `TB-`** ——
   两边不一致就是另一条要记的口径。

打包两条注意（文件已按它们写好，改的时候别再踩）：
① `.vbp` 只写 `Type=Exe`，**别加 `Flavor=`**（VB6 报「包含无效关键字」）；也不需要 `Reference=` 行，
两份工程只用 VB 运行时 + 自己那枚 `Holder`。
② 任何一句**编不过 / 跑中途崩**，把原话贴在那个编号上，并说清「跑到 X 之后没了」——
报错本身就是答案（那说明 VB6 压根不允许那个形状，C3 该拒绝而不是模仿）。

## 为什么值得跑这一趟（语料实数，`D:\vb_yqt4qPac` 本机扫）

**5701 个 `Type … End Type` 块里 80 块带对象型成员**：裸 `As Object` 10 块、
COM/VB 对象型 39 块（`StdPicture` 38、`IPictureDisp` 18、`Collection` 15、`StdFont` 3、`ISubclass` 1）、
工程类型 31 块（`pdDIB` 14、`cDIB` 8、`cPalette` 3、`pdStringStack` 3、`pd2DPath` 2、`ucListView` 1；三档按块计，有重叠）。
最典型的一枚就在跑通过的样例工程里：`ComCtlsDemo\Builds\StatusBar\StatusBar.ctl:394`
`ShadowPanelStruct.Picture As IPictureDisp`（同一文件 `:374` 的 `InitPanelStruct` 也带一枚），
`VBFlexGrid.ctl:875 TCELLFMTG.Picture` 同理 —— 也就是说这些结构体是真的会被整枚赋值、
真的会出过程、真的会被 `Erase`，不是纸面形状。
