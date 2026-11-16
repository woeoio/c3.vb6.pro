# 039 - 内置 Dictionary（设计）

> 状态：**S1/S2/S3 已完成**（设计 + 基准 + 实现 + 对照全绿）。S4（三条入口的 New 那一半）已在实现里做了，剩 S5 门禁登记与 S6 收口。
> 当前读数：`PASS dict_native 用例=54 行`（1 行 VB6-SEM 手工期望，5 行已知 C3 缺口剔除，明细见 §9）。
> 前身：C30-10 第二十二~二十五刀（内置正则引擎）验证了「ProgID 拦截 + 自带 IDispatch」这条路，
> 本文件把同一条路铺到 `Scripting.Dictionary` 上，并补上正则那刀没覆盖的**前期绑定**入口。

## 1. 为什么要内置

`CreateObject("Scripting.Dictionary")` / `New Scripting.Dictionary` 现在都落到真 `scrrun.dll`：

- 目标机必须有 Scripting Runtime。Windows 全系自带，但 Server Core 精简镜像、被组策略禁用的环境、
  Wine 下不一定有 —— 缺件时的表现是运行期 429 `ActiveX component can't create object`。
- 每一次 `Item` / `Add` / `Exists` 都是一次 IDispatch 出入，`For Each` 还要过一次 `IEnumVARIANT`。
- 我们已经在 `vb6com.c` 里维护了一张「CreateObject 产物的 coclass 名」表（账 #611）——
  只为了让 `TypeName(d) = "Dictionary"` 成立，**名字答对了但成员一个都没实现**，这个割裂本身就是债。

目标：让 `Set d = CreateObject("Scripting.Dictionary")` 与 `Dim d As New Scripting.Dictionary`
都拿到本仓自带实现，行为与真组件逐字对齐（判据见第 4 节），并在默认路径下**不再依赖 scrrun.dll**。

非目标：`FileSystemObject` / `TextStream` / `Drive` / `Folder` 等 scrrun 的其它类（那是另一件事，
需要文件系统语义与错误码面，不在本设计里）；`Dictionary` 的 `SaveToFile` 这类 scrrun 私有扩展
（scrrun 的 Dictionary 其实没有，但 `FileSystemObject` 有）一律不做。

## 2. 现状取证（代码位置，均为当前 dev 头）

| 事实 | 位置 |
| --- | --- |
| `vb6com/` 下没有字典实现，只有 `vb6com_regexp.c` 是内置组件 | `src/rtl/core/vb6com/` 清单 |
| `CreateObject` 里唯一的内置拦截是 `VBScript.RegExp` | `vb6com.c:1248` 附近（C30-10） |
| 三条建对象出口 | `vb6_CreateObject` / `vb6_NewObject` / `vb6_NewObjectClsid` |
| 运行时自带类的唯一登记处（Collection/StdFont/PropertyBag） | `vb6rtl_conv.c` `vb6_NewVb6RuntimeClass()` |
| 内建名表的主人（Collection 那条） | `vb6com_collection.c` `vb6_NewBuiltinObject()` |
| `TypeName` 把 ProgID 最后一段当 coclass 名 | `vb6com.c:1198` 起 + `vb6rtl_conv.c:578` |
| 成员调用**按名字**在运行期发 | `vb6_ComCall(disp, L"Add", …)` / `vb6_ComGetProp(disp, L"Item")` |

**关键结论（决定实现形状）**：C3 对外部 COM 的成员调用不发 vtable 偏移，一律
`GetIDsOfNames` + `Invoke` 按名字走。因此内置 Dictionary **不需要**逐位对齐 scrrun 的数值 DISPID，
只需要自己一张「名字 → DISPID」表；`DISPID_VALUE(0)` 必须留给默认成员 `Item`，
`DISPID_NEWENUM(-4)` 留给 `For Each`（与 Collection 同口径）。

## 3. 要挂的三个入口（用户点名：不能只拦 CreateObject）

| 源码写法 | 编译成 | 现在 | 内置后 |
| --- | --- | --- | --- |
| `CreateObject("Scripting.Dictionary")` | `vb6_CreateObject(L"Scripting.Dictionary")` | 真 scrrun | 命中内置（在 `vb6_CreateObject` 顶部拦，同 RegExp） |
| `New Scripting.Dictionary`（无 CLSID） | `vb6_NewObject(L"Scripting.Dictionary")` | 先问 `vb6_NewVb6RuntimeClass` → 未命中 → scrrun | 在 `vb6_NewVb6RuntimeClass` 里认出 `Scripting.Dictionary` |
| `Dim d As New Scripting.Dictionary` / `New Scripting.Dictionary`（typelib 有 CLSID） | `vb6_NewObjectClsid(clsid, L"Dictionary")` | 同上先问同一函数 | 同上（**一处登记，两条出口共用**，与 StdFont 那条教训同族） |

补充事实：`vb6_NewVb6RuntimeClass` 收到的类名取 `comSym->name`，对限定名只取最后一段
（`cgen_base_type.cpp:703`），所以登记名写 `Dictionary` 即可；为了防御手写 `Scripting.Dictionary`
这种「类名带库前缀」的写法，登记处同时认 `Dictionary` / `Scripting.Dictionary`（与 `stdole.StdFont` 同族）。

**顺手要补的同类欠账**：`New VBScript.RegExp` 现在**不**命中内置正则（拦截只在 `vb6_CreateObject`），
`Dim re As New VBScript.RegExp` 会去加载真 `vbscript.dll`。这与本设计是同一个缺口，
内置 Dictionary 这刀要一并把 RegExp 也登记进 `vb6_NewVb6RuntimeClass`，并在用例里钉一条
`Dim re As New VBScript.RegExp` 的判据（否则「内置」只是晚绑定那半张脸）。

回退开关沿用 C30-10 的口径：`C3_DICT_PREFER_MS=1` 时照旧用 Microsoft 的 scrrun（对照组用），
`C3_REGEXP_PREFER_MS=1` 同理。

## 4. 行为基准（已经钉住的部分）

夹具：`tests/dict_native/DictCases.vbs`（UTF-8 **无 BOM** + CRLF —— cscript 不认 UTF-8 BOM，
第一版带 BOM 直接报「无效字符」，这是本文件踩过的第一个坑），基准输出
`tests/dict_native/dict_oracle.txt`（**56 行**）。

复现：

```
cscript //nologo tests\dict_native\DictCases.vbs     # 重写 dict_oracle.txt
```

已确认的读数（全部来自真 scrrun，`cscript` 执行）：

| 行为 | 读数 |
| --- | --- |
| `TypeName(d)` | `Dictionary` |
| `Add` 重复键 | `457`（Binary；Text 模式大小写不同也算重复，同样 457） |
| `Remove` 缺失键 | `32811` |
| `Item` **缺失键** | **不报错，返回 Empty（VT_EMPTY）** —— 见下方口径分歧 |
| `Keys` 顺序 | 插入序；删中间一个后其余保持原序；重新 Add 的键跑到**末尾** |
| `Items` | 跟 `Keys` 同序 |
| `RemoveAll` 后 `Count` | 0 |
| `Keys` / `Items` 是数组 | `IsArray=True`、`LBound=0` |
| `For Each` | 枚举的是**键** |
| `CompareMode` 默认 | 0 |
| 有元素时改 `CompareMode` | 运行期错误 `5` |
| `Item(key) = v`（不存在则建键） | 建键成功，`Count` +1 |
| `Add` 的键为 `Null` / `Nothing` / `Empty` | 都不报错；`Exists(Null/Nothing/Empty)` 均为 True |
| 键类型（Binary） | `Add 1` 后 `Exists(1)=True`、`Exists("1")=False`；`Add CDbl(1)` 后 `Exists(1)=True`（**数值跨数值型互通，字符串不互通**） |
| 键相等（Text 模式） | `Add "Key"` 后 `Exists("key")=True`（字符串大小写不敏感）；`Add CDbl(1)` 后 `Exists("1")=False`（**Text 不等于"全部转字符串"**） |
| 尾随空格 | Binary / Text / Database 三档实测 `Exists("ab  ")` 均 False（scrrun 不裁剪空格） |
| 数组值 | 可以存数组并读回（`d("arr")(1)` 那一形另见测试） |
| 对象值 | 可以存对象，`IsObject` / `TypeName` 正常（存另一个 Dictionary 读回 `Dictionary`） |
| `Count` 的 VarType | 3（VT_I4）；`Keys` 的 VarType 8204（Variant 数组），元素类型随键 |
| `Item` 赋值到已有键 | 只改值，**不改变位置**（`keys=a,b`） |
| `For Each` 中删除当前键 | 不报错，序列仍走完（`a;b;c`），结束后 `count=2` —— **枚举是快照语义** |

### 4.0.1 键相等真值表（S1 实测，实现按此写）

| 键 | 与之相等 | 与之不等 |
| --- | --- | --- |
| `Add True` | `True`、`-1` | `1`、`"True"` |
| `Add CDate("2026-11-16 10:30")` | 同一个 `Date`、它的数值（`CDbl`） | 它的字符串形式 |
| `Add CCur(1.5)` | `CCur(1.5)`、`CDbl(1.5)` | `1` |
| `Add CInt(7)` | `CLng(7)`、`CDbl(7)` | `"7"` |
| `Add Empty` | `Empty`、`0`、`""`（**三者互为同一个键**） | `Null` |
| `Add Null` | `Null` | `Empty` |
| `Add Nothing` | `Nothing` | `0`、`""` |

读法：数值类（含 Bool/Date/Currency）跨类型互通；字符串只与字符串相等；
`Empty` 是「零值」那种特例（与 `0` 和 `""` 都撞车），`Null` 与 `Nothing` 各自独立成类。

另一个反直觉点：`CompareMode = 3` **不报错**（读回仍是 3），但判重行为等同 Text
（`Add "A"` 后再 `Add "a"` → 457）。`CompareMode = -1` 才报 5。
实现按实测走：接受 0..3，>0 一律按「大小写不敏感」处理，负数报 5。

### 4.1 口径分歧：`Item` 缺键（必须显式记账）

- **cscript（真 scrrun，晚绑定）**：`v = d.Item("nope")` **不抛错**，`v` 是 Empty，`Err.Number=0`。
  连不装 `On Error Resume Next` 也不抛（`.build/dprobe2.vbs` 实测：后续语句照常执行）。
- **VB6（文档与常识）**：`d("nope")` 抛 `32811 Element not found`。

两者不一致，原因是 VB6 走类型库（propget 语义 + 错误映射），而 VBScript 把 `d.Item("x")`
当方法调用发 `DISPATCH_METHOD`，scrrun 那一趟没 raise。

**本项目取 VB6 语义（抛 32811）**，理由：C3 编译的是 VB6 源码，真实 VB6 程序围绕 `On Error`
写容错的不少；若返回 Empty 静默继续，是**静默错值**，比报错难查。这一格因此在夹具里是
**手工期望**（`DictCases.vbs` 里对应行标注 `VB6-SEM`），与 ai/038 里 `\Z` 那格同一处理方式。
实现侧判据：`Item` 缺键 → `vb6_RaiseError(32811, …)`。

## 5. 成员面与错误码（实现目标）

| 成员 | 形态 | 语义要点 |
| --- | --- | --- |
| `Add(key, item)` | method | 键已存在 → 457（按当前 CompareMode 判重） |
| `Item(key)` | propget，**DISPID 0**（默认成员） | 缺键 → 32811（VB6 语义，见 4.1） |
| `Item(key) = v` | propput | 缺键则建键（保持插入序追加） |
| `Exists(key)` | method → Boolean | 返回 -1/0（VB6 布尔） |
| `Count` | propget | 必须**正** DISPID（`vb6_ForEach_Init` 用 `dispid>0` 判 Count 可用） |
| `Keys()` | method → Variant() | 0 基数组，插入序 |
| `Items()` | method → Variant() | 与 Keys 同序 |
| `Remove(key)` | method | 缺键 → 32811 |
| `RemoveAll()` | method | 清空，保留 CompareMode |
| `CompareMode` | propget/propput | 0=Binary，1=Text，2=Database；**非空表赋值 → 5**；非 0/1/2 → 5 |
| `_NewEnum` | `DISPID_NEWENUM(-4)` | `IEnumVARIANT`，枚举**键** |

错误码：457（键已存在）、32811（元素不存在）、5（无效过程调用/参数）、450/438（成员不存在，走通用路径）。

## 6. 键模型（实现的核心难点）

键按「类型类 + 值」比较，**不做全局字符串化**：

1. 数值类（`VT_I2/I4/I8/UI*/R4/R8/CY/DECIMAL/BOOL`）互通：比较前统一换算成 `long double`
   （`CY` 按 10000 缩放；`BOOL` 按 -1/0）。实测支持「Add CDbl(1) 后 Exists(1)=True」。
2. 字符串类（`VT_BSTR`）：Binary 模式逐字节；Text 模式按 **`CompareStringW(LOCALE_USER_DEFAULT,
   NORM_IGNORECASE)`**；Database 模式再加 `NORM_IGNOREWIDTH|NORM_IGNOREKANATYPE`。
3. `VT_EMPTY` / `VT_NULL` / `VT_DISPATCH(NULL)` 各自独立成一类（实测三者互不相等且都能当键）。
4. 日期（`VT_DATE`）：**已实测** —— 与它的数值（`CDbl`）相等、与其字符串形式不等，
   因此归入数值类（第 6 节口径正确，无需另设一类）。
5. 字符串与数值之间**永不相等**（实测 `Exists("1")=False`）。

哈希：对「类型类 + 规范化后的值」算 32 位 FNV-1a；桶内线性比对（桶数组 + 保序链表/向量），
删除只从索引摘除、保留向量里的墓碑或在压缩时重建（选型见第 7 节）。

## 7. 数据结构与实现方案

- `Vb6Dict`：`entries[]`（保序：key/val 的 VARIANT 副本 + 活标志）、`hash[]`（桶头索引）、
  `count`、`cmpMode`、`refCount`。
- 查找：哈希 → 桶链 → 用第 6 节比较器逐项比对；比较器是唯一入口（`dictKeyEq` / `dictKeyHash`），
  避免「Add 用一个算法、Exists 用另一个」这类双写漂移。
- 删除：标记墓碑；`Keys/Items/枚举` 跳过墓碑；当墓碑 > 活数 且 > 32 时压缩重建（保持相对顺序）。
- 对象模型：`IDispatch` + `IEnumVARIANT`（枚举器 vtable 与 `vb6com_collection_enum.c` 同族，
  可直接复用其骨架，但枚举值取**键**）。
- `TypeName`：内置实例直接由 `vb6_ComObject_ClassNameBStr` 那条腿答 `Dictionary`，
  不再依赖账 #611 的 ProgID 表（但表继续留着，给「优先 Microsoft 组件」那条回退路径用）。

## 8. 分步计划（每刀一个提交，判据先行）

| 步 | 内容 | 判据（必须真跑） |
| --- | --- | --- |
| S1 | **补全基准**：日期/货币/布尔/`Nothing` vs 0、Database 模式大小写、`Item` 赋值与 `Add` 的键规范化差异、`For Each` 在删除后 | `dict_oracle.txt` 扩到 ≥55 行；每行注明 cscript 读数或 `VB6-SEM` |
| ~~S1~~ | **已完成**：56 行，含键相等真值表（4.0.1）、`CompareMode=3` 语义、`For Each` 快照行为 | 已入仓 |
| ~~S2~~ | 已合并进 S3（理由见下）：`vb6com_dict.c` 的容器/比较器/桶与五个 IDispatch 成员一起落 | — |
| ~~S3~~ | **已完成**：`vb6com_dict.c`（容器 + IDispatch + IEnumVARIANT） | `run_compare.ps1` → **PASS 54 行** |
| S2 | 骨架：`vb6com_dict.c`（数据结构 + 比较器 + 哈希）+ 单测壳（C 侧自测入口，不进 VB 面） | 本地 C 自测 30 例全过（键相等矩阵） |
| S3 | IDispatch 面：`Add/Exists/Item/Count/Remove/RemoveAll/Keys/Items/CompareMode` + `_NewEnum` | C3 侧夹具 `DictMain.bas` 与 `dict_oracle.txt` 逐行一致（先只走 `CreateObject`） |
| ~~S4~~ | **已完成（与 S3 同刀）**：`vb6_CreateObject` 拦截 + `vb6_NewVb6RuntimeClass/vb6_NewBuiltinObject` 登记（`Dictionary` / `Scripting.Dictionary`）；同刀把 `RegExp` 也登记（补 `New VBScript.RegExp` 欠账） | `Dim d As New Scripting.Dictionary` 已可建；`C3_DICT_PREFER_MS=1` 回退真 scrrun |
| S5 | 门禁与文档：`[TEXT] dict_native` 进 `run_tests.ps1`（真跑两侧 diff）、`check_created_object_typename.ps1` 计数 +1、发码清单登记 | CI 22/22；`[TEXT] dict_native` 行出现在日志 |
| S6 | 收口：`ai/037`/`ai/038` 交叉引用、`ai/039` 补实测读数与「未做」清单 | 文档与代码一致（含回退开关说明） |

S1 先行的理由与正则那刀一致：**判据不许事后补**。S3 之前不接受"看着对"的实现。

## 9. 风险

1. **键相等是唯一的高危面**。`1` / `"1"` / `1.0` / `True` / `Empty` / `Nothing` / 日期 在 Binary 与
   Text 下各有一张真值表，错一格就是"存的进去、查不出来"（静默错值）。对策：比较器单点 + S2 的矩阵自测。
2. **`For Each` 与删除交互**：scrrun 在枚举过程中删除元素的语义未实测（S1 补）。若行为古怪，
   按实测对齐，不按"合理"设计。
3. **前期绑定那条腿**：`New` 走了 `vb6_NewObjectClsid` 时类名来自符号名，若某工程用
   `New Scripting.Dictionary.1`（带版本后缀）这种写法，登记名要能吃掉后缀 —— S4 要专门钉一条。
4. **回退开关不能变成第二套语义**：`C3_DICT_PREFER_MS=1` 只换实现，不换判据。

## 10. 不做

- `FileSystemObject` 全家（另一件）；
- scrrun 的 `Dictionary` 之外的类型（`TextStream` 等）；
- 任何"键全转字符串"的简化实现（会破坏第 6 节已实测的类型区分）；
- 性能优化放在功能对齐之后（先有正确性，再谈把 `For Each` 里的 VARIANT 拷贝减少）。

## 11. 实现与实测（S2/S3 落地后）

### 11.1 落地的东西

| 位置 | 内容 |
| --- | --- |
| `src/rtl/core/vb6com/vb6com_dict.c` | 容器（保序 entries + 墓碑 + 桶链）、键比较器/哈希（唯一入口）、IDispatch、IEnumVARIANT |
| `vb6com.c` | `vb6_CreateObject` 顶部拦截，回退开关 `C3_DICT_PREFER_MS=1` |
| `vb6com_collection.c` | `vb6_NewBuiltinObject` 认 `Dictionary` / `Scripting.Dictionary`，并补上 `RegExp` 的 New 那一半 |
| 五处 RTL 登记 | `CMakeLists.txt` / `c3rtl.rc(228)` / `rtl_embedded.hpp` / `rtl_embedded.cpp` / `driver_link.cpp` |
| `tests/dict_native/` | `DictCases.vbs`（oracle）+ `DictMain.bas` + `DictApp.vbp` + `run_compare.ps1` |
| `tests/run_tests.ps1` | `[TEXT] dict_native`（真跑两侧再 diff，走静态分片） |

复现：

```
pwsh -File tests\dict_native\run_compare.ps1        # 期望 PASS 54 行（54/57，剔除 5 缺口）
```

### 11.2 实测读数

```
PASS dict_native 用例=54 行 两侧逐字一致（oracle=cscript scrrun；VB6-SEM 覆盖 1 行；已知 C3 缺口剔除 5 行）
```

- **VB6-SEM 1 行**：`d10`（`Item` 缺键）—— cscript 给 Empty，本仓取 VB6 的 32811（§4.1）；
- **已知 C3 缺口 5 行**（不是 Dictionary 的语义问题，全部在 `run_compare.ps1` 的 `$C3Gap` 里点名，不静默跳过）：

| 行 | 现象 | 实质 |
| --- | --- | --- |
| `d62` `d104` | ~~`Nothing` 当键时与 `Empty` 撞键~~ | **已修**（第二十九刀）：`vb6_ComPackNothing` ⇒ VT_DISPATCH NULL |
| `d102` `d123` | ~~Currency 键被截成整数~~ | **已修**（第二十九刀）：`vb6_ComPackCur` ⇒ VT_CY（10000 倍定标） |
| `d121` | `VarType(CCur(1.5))` = 5（应 6）；`VarType(1)` = 3（应 2） | 仍在：标量类型宽度那一层（C3 用 double 表 Currency、字面量默认装 I4） |

另有一个不在基准里、但**使用者会碰到**的缺口：`d("k") = v`（默认成员带参赋值）
~~C3 目前当成"把值赋给对象变量"并告警~~ —— **已修**（第二十九刀）：发码本来就对
（`vb6_ComSetPropArg(d, 0, {key}, 1, vb6_ComPackInt(5))`），错的是语义层那条假
VB3003；现在对「目标是对 Object/Variant 变量的带参下标」不再做隐式转换检查。

### 11.4 New 那条腿的实测（2026-11-16，临时探针，不入库）

`New` 与 `As New` 两条写法分开读，因为 C3 对它们走的是不同代码：

| 写法 | Dictionary | RegExp |
| --- | --- | --- |
| `CreateObject("…")` | ✅ `TypeName=Dictionary` | ✅ `TypeName=RegExp` |
| `Set x = New …` | ✅ `TypeName=Dictionary`，Add/Count/Item/Keys 均可用 | ✅ `TypeName=RegExp`（补了 ProgID 登记后） |
| `Dim x As New …` | ⚠️ 对象能建（`Add`/`Count` 正常），但 `TypeName` 答 `Long` | ❌ 变量恒为 `Nothing`（`[C3_LATEBIND_MISS … kind=null]`） |

结论：**内置实现本身没问题，是 `As New` 那条腿的代码生成有缺口**，
且两个名字行为还不一致（字典能建、正则不能），说明它不是“外部 ProgID 统一不行”，而是另有分支。
追账项：(a) `As New` 外部 ProgID 的惰性初始化；(b) `As New` 变量的 `TypeName` 与实际对象不符。
在这两条修好之前，文档与夹具统一推荐 `Set x = New …`，不推荐 `Dim x As New …`。

### 11.5 声明成类名的槽位（本刀修的真回归）

`tests/test_com_default_prop.bas` 用的是 **`Dim d As Dictionary`**（不是 `As Object`），
这种写法编译器生成的是：

```c
vb6_ComIface_IDictionary* d = NULL;
d = (vb6_ComIface_IDictionary*)vb6_ComCreateTyped(_own_1, "{42C642C1-97E1-11CF-978F-00A02463E06F}");
```

而 `vb6_ComCreateTyped` 的老做法是「`vb6_CreateObject` 拿到 `IDispatch` → 按 scrrun/vbscript
的 typelib IID 再 QI 一次 → 释放原指针」。内置对象只实现 `IDispatch`，那一步 QI 必失败
⇒ 槽位变 `NULL` ⇒ 后面每一条成员调用都打一条 `[C3_LATEBIND_MISS]`。

**读数**：本刀合入前该用例 `latebind-miss 期望 0 条, 实到 7 条`（CI `43854293457`）；
修后本机实跑 `DP-1..DP-4` 全 OK、`P24-10: 4/4 passed`、stderr **0 条**。

修法（不动调用方、不改对象布局）：`vb6_Dict_IsInstance` / `vb6_RegExp_IsInstance`
按「首字段 vtable 指针是否为本族静态表」认亲，`vb6_ComCreateTyped` 遇到内置对象就
**跳过类型化 QI、直接交回自己的 IDispatch** —— 下游成员调用本来就全按名字走（`vb6_ComCall*`），
不需要那一层接口。

### 11.7 第二十九刀：清四条缺口（含取证）

全部先 `--emit-c` 取证，再改：

```
d.Add Nothing,   "n"   →  vb6_ComPackInt(NULL)                      ← 根因：Nothing 落到默认整数档
d.Add CCur(1.5), "cur" →  vb6_ComPackInt(vb6_CCur(1.5))             ← 根因：Currency 没有打包档
d("k") = 5             →  vb6_ComSetPropArg(d, 0, {key}, 1, …)      ← 发码正确，错在假 VB3003
Dim r As New …RegExp   →  void* r = NULL;（无惰性初始化行）          ← 仍在
```

| 缺口 | 处置 | 读数 || --- | --- | --- |
| `Nothing` 包成 Empty | 新增 `vb6_ComPackNothing`（VT_DISPATCH NULL），`comPackExpr` 对 `LiteralKind::Nothing` 走它 | 对照回填 2 行（`d62`/`d104`） |
| Currency 实参截断 | `kComPackTiers` 加 `Currency → vb6_ComPackCur`，RTL 按 `llround(v*10000)` 发 VT_CY | 对照回填 2 行（`d102`/`d123`） |
| `d("k") = v` 假 VB3003 | 语义层：目标是「Object/Variant 变量的带参下标」时不做隐式转换检查 | 编译该形 **0 条告警**（修前 1 条） |
| `As New` 外部 ProgID | **已修**（第三十刀）：局部声明那一支（`cgen_localdecl.cpp`）与模块级那一支（`cgen_decl_var.cpp`）都登记 ProgID 档，守卫发 `vb6_NewObject(L"<ProgID>")` | `[TEXT] newbinding` **4/4**（reg_exp 局部+模块级、dict 局部+模块级；stderr 0 条） |

对照总读数：**58 行逐字一致**（VB6-SEM 1 行；仅 `d121` 因标量类型宽度剔除，见 §11.2）。

### 11.8 第三十刀：`As New <外部 ProgID>` 收口（含定位过程）

前一轮猜错了落点（以为在 `cgen_decl_var.cpp`），实测证明不是：加完登记后 `--emit-c` 一字未变。
本轮用环境变量门控的诊断确认：`Dim r As New VBScript.RegExp`（过程内）**根本不进**
`visit(VariableDecl&)`，只打印出模块级的 `g`。真正的局部声明发射在 `cgen_localdecl.cpp`。

| 位置 | 处置 |
| --- | --- |
| `cgen_localdecl.cpp` | 局部 `Dim x As New <带点未解析名>` 登记 `knownNewVars_` + `knownNewProgIdVars_` |
| `cgen_decl_var.cpp` | 模块/类级同一档登记（连同 `moduleNewProgIdVars_`） |
| `cgen_decl_func/proc/prop.cpp` | 过程入口把模块级 ProgID 表带进 `knownNewProgIdVars_` |
| `cgen_expr_ident_{symbol,dispatch}.inc`（4 处） | 守卫多一档：ProgID 命中就发 `(void*)vb6_NewObject(L"<ProgID>")` |

读数（`tests/dict_native/run_newbinding.ps1`，VB6 语义固定期望）：

```
PASS newbinding 4/4（reg_exp 局部+模块级、dict 局部+模块级；stderr 0 条）
  n3 local_dict=1 typename=Long  ← 仍差一格：VB6 答 Dictionary
```

修前：四条全是空/NULL，stderr 一片 `[C3_LATEBIND_MISS … kind=null]`。
**残留一格**：局部 `Dim d As New Scripting.Dictionary` 的 `TypeName` 答 `Long`
（模块级同一名字答 `Dictionary`）—— 接口指针被 `knownLongPtrVars_` 当 LongPtr 装箱，
功能不受影响（`d.Add`/`d.Count` 全对），判据里不断言这一格、只报告。

### 11.9 与计划的差异（说明理由）

- **S2 与 S3 合并**：原计划 S2 先做"C 侧键矩阵自测"。实际做下来没有干净的运行入口 ——
  本仓的测试都是 VB 夹具，另开一个 C 单测程序要把 RTL 的 VARIANT 胶水一起拖进来，
  成本远高于收益；而键矩阵本身已经在 54 行对照里被真组件验过（每一条键相等都在里面）。
- **S4 提前**：既然 `vb6_NewBuiltinObject` 本来就要改（Registry 名字表），顺手把 `New` 那一半
  和 `RegExp` 的欠账一起做完了 —— 否则 "内置" 只覆盖晚绑定，与用户点名的那条要求不符。
