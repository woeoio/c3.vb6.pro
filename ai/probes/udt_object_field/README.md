# UDT 里的对象字段归谁（账 #600 的口径，给真 VB6 / twinBASIC 环境用）

C3 现在把整枚 UDT 赋值（`t2 = t1`）发成逐字节 `memcpy`（`vb6_udtcpy_*`），出过程也不清对象字段 ——
所以 UDT 里的对象字段是**共享指针**。第317刀当时就是把这一格押后的，理由是：
「装订（拷贝时 claim）/ 覆写退旧 / 出厂真清 / getter 交回带对价」这四格是**一条链**，
只做装订会把共享指针变成双重释放，而这条链该怎么修**完全取决于 VB6 那一侧答什么**，本机没有 VB6。

这一份就是去问那一句的。请**照实贴回顺序**，不用解释。

## 回收状态（2026-11-02）

- **`copy/` 那一份 = VB6 已答**（用户代跑，原文见文末「VB6 实测」）：主问落定 —— 整枚 UDT 赋值对对象字段**自己持有一枚引用**、覆写**先退旧**、出过程**真清**、三种声明档同形、数组元素与嵌套同形、`init=7 term=7` 配平。⇒ 四格链（装订 / 覆写退旧 / 出厂真清 / getter 带对价）按「VB6 做全套」那一支做。
- **上一级那一份（字段级）第一次跑不起来，原因钉在这儿**：模块名与入口过程同名 —— VB6 弹窗**「主程序不能是模块、类型、工程或窗体名」**，twinBASIC 报 **`TB5073: 'Main' is ambiguous`**（`UdtObjProbe.Main.Main` 过程 vs `UdtObjProbe.Main` 模块）。已把模块改名 `ProbeHost`（`Sub Main` 那句不动、`Startup="Sub Main"` 不动、`Holder.cls` 里改问 `ProbeHost.gTerm` / `ProbeHost.Say`）⇒ **请重跑这一份**。
- `copy/RejectCases.bas` 按设计**没入工程**（编不过的那几格会把整份拖倒）⇒ 想问 R1~R4 就在 IDE 里Project → Add File **一次加一格**、编译、把原话抄在编号上。

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


## 附：C3 这一侧现在的读数（本机实测，等你的 VB6 读数来对照）

两份工程用**修完账 #608 的 C3** 编出来跑（修前 `Print #fn, gAll` 交的是指针数，读不出东西）。
两份的 `Z-COUNT` 都不配平，且 `TERM` 几乎不出现 —— 这就是「共享指针 + 没人退」的形状：

| 那一带 | C3 现在的读数 |
|---|---|
| B 覆写字段 | `TERM 1` 不出现（B3、B4 都没有）⇒ 覆写不退旧 |
| C `Set t.O = Nothing` | 同样不退 ⇒ 置空不真清 |
| D1 / D2 / D3 出过程 | 三档都不退 ⇒ 出厂不清（声明档之间**没有**差别）|
| E `Erase` / 重新 `ReDim` | 不退 ⇒ 元素字段被丢在原地 |
| H 读出后清空字段 | `H4-tag-through-g=0` —— 读回来那枚连 `Tag` 都答 0；整段无 `TERM` |
| M 两字段同指一枚 | `Set h = Nothing` **立刻** `TERM 11` ⇒ 三处持有只算一枚引用 |
| J `ByRef` 里置空 | `J4-after-call caller-field-is-Nothing=True`（按地址是通的）；callee 里不退 |
| N 新 `Dim` 字段 | `N1-fresh-is-Nothing=True`（干净，控制组过）|
| A 整枚赋值 `t2 = t1` | `sameIs=True`，`Set t1.O = Nothing` 后**无** `TERM`，`A4-after-sub` 也无 ⇒ 别名 + 谁都不放 |
| A5 目标本已握着 | `A6` / `A7` 都不退 ⇒ 整枚覆写不退旧 |
| G2 / G4 另两种声明档 | 与 A 同形 ⇒ C3 不按声明档分叉 |
| F 字符串字段 | `F1-b-string=AAA` ⇒ 深拷贝这一半已经对（本刀不动它）|
| I 嵌套 UDT / E2 数组元素 | 与 A 同形（`E3`、`I2` 都不退）|

计数：字段级那份 `init=14 term=1`，整枚赋值那份 `init=7 term=0`。
## 为什么值得跑这一趟（语料实数，`D:\vb_yqt4qPac` 本机扫）

**5701 个 `Type … End Type` 块里 80 块带对象型成员**：裸 `As Object` 10 块、
COM/VB 对象型 39 块（`StdPicture` 38、`IPictureDisp` 18、`Collection` 15、`StdFont` 3、`ISubclass` 1）、
工程类型 31 块（`pdDIB` 14、`cDIB` 8、`cPalette` 3、`pdStringStack` 3、`pd2DPath` 2、`ucListView` 1；三档按块计，有重叠）。
最典型的一枚就在跑通过的样例工程里：`ComCtlsDemo\Builds\StatusBar\StatusBar.ctl:394`
`ShadowPanelStruct.Picture As IPictureDisp`（同一文件 `:374` 的 `InitPanelStruct` 也带一枚），
`VBFlexGrid.ctl:875 TCELLFMTG.Picture` 同理 —— 也就是说这些结构体是真的会被整枚赋值、
真的会出过程、真的会被 `Erase`，不是纸面形状。

## 附二：VB6 实测读数（`copy/UdtCopyProbe.vbp`，2026-11-02 用户代跑，原文照贴）

```
P0-begin
INIT 1
A2-after-copy sameIs=True
A3-after-source-clear
TERM 1
A4-after-sub
INIT 2
INIT 3
A5-both-set
TERM 3
A6-after-struct-overwrite
TERM 2
A7-after-sub
INIT 4
G2-sameIs=True
G3-after-source-clear
TERM 4
G6-after-sub
INIT 5
G4-sameIs=True
G5-after-source-clear
TERM 5
G7-after-sub
F1-b-string=AAA
F2-after-sub
INIT 6
I1-copied-nested
TERM 6
I2-after-sub
INIT 7
E2-element-copy sameIs=True
E3-after-clear-element-zero
TERM 7
E4-after-sub
Z-COUNT init=7 term=7
Z-DONE
```

逐条读法与它把 C3 定成哪一支：

| 读数 | VB6 答的 | C3 该怎么 |
|---|---|---|| `t2 = t1` 编译通过 | 这句**合法**（我原先准备的「可能不许编译」那一支不成立）| 整枚赋值必须自己处理对象字段，不能推给「拒绝编译」|| `A3` 打在 `TERM 1` **之前** | 源放手时目标仍持有 ⇒ **拷贝时 AddRef** | 装订那一格要做（现在发 memcpy ⇒ 零引用）|| `TERM 1` 落在 `A4-after-sub` **之前** | 出过程真放 | 出厂那一格要做（现在什么都不放 ⇒ 泄漏）|| `sameIs=True` | 交的是同一枚实例，不是克隆 | 装订 = AddRef，**不是**深拷贝对象|| `A5-both-set` → `TERM 3` → `A6` | 整枚覆写**先退目标自己那一枚** | 覆写退旧那一格要做（现在裸 memcpy ⇒ 目标旧的没人放）|| `A6` 之后到 `A7` 之间只有 `TERM 2` 一条 | `a.O` 与 `b.O` 同指 #2，出过程各放一次、只有归零那次打 TERM | 计数要真配平：**不能**只在覆写时退、也不能出过程退两次|| `G2/G3`（`As Holder`）、`G4/G5`（`As Variant`）与 A 档逐条同形 | 三种声明档**一套口径** | 发码不必按声明档分叉（但 Variant 那一格要按 Variant 的放法走 `VariantClear`）|| `F1-b-string=AAA` | 字符串字段深拷贝 | C3 这一半本来就对，不动 || `I1` → `TERM 6` → `I2` | **嵌套 UDT**（对象在内层）同一条链 | 拷贝要递归到内层，不能只看一层字段表|| `E2`/`E3` → `TERM 7` → `E4` | 数组元素之间整枚赋值同形 | `arr(1) = arr(0)` 那一支也要走同一位主人|| `Z-COUNT init=7 term=7` | 全程零泄漏、零双放 | 改完之后同一份夹具要读到**同一个数**（我方现在 7/0）|
对照我方正面的读数（`ai/probes/udt_object_field/README.md` 附一节，修完账 #608 才跑得出来）= `init=7 term=0`、`A3`/`A4`/`A6`/`A7` 前后**一条 TERM 都没有** ⇒ 上面四格一格都没做，且从不双放（所以是纯泄漏，不是崩溃）。
