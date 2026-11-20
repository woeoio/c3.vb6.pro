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
PASS dict_native 用例=59 行 两侧逐字一致（oracle=cscript scrrun；VB6-SEM 覆盖 1 行；已知 C3 缺口剔除 0 行）
```

- **VB6-SEM 1 行**：`d10`（`Item` 缺键）—— cscript 给 Empty，本仓取 VB6 的 32811（§4.1）；
- **已知 C3 缺口 0 行**（第三十二刀把最后一格 `d121` 收回；`$C3Gap` 名单现为空，闸门仍在）：

| 行 | 现象 | 实质 |
| --- | --- | --- |
| `d62` `d104` | ~~`Nothing` 当键时与 `Empty` 撞键~~ | **已修**（第二十九刀）：`vb6_ComPackNothing` ⇒ VT_DISPATCH NULL |
| `d102` `d123` | ~~Currency 键被截成整数~~ | **已修**（第二十九刀）：`vb6_ComPackCur` ⇒ VT_CY（10000 倍定标） |
| `d121` | ~~`VarType(CCur(1.5))` = 5（应 6）；`VarType(1)` = 3（应 2）~~ | **已修**（第三十二刀）：装箱站按 VB6 类型选档 —— Currency ⇒ `vb6_VariantCur`（VT_CY）、Integer 档 ⇒ `vb6_VariantInt`（VT_I2），见 §11.10 |

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

对照总读数：**59 行逐字一致**（VB6-SEM 1 行；缺口名单已清空，见 §11.2 / §11.10）。

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

### 11.10 第三十二刀：标量宽度两格（Integer / Currency 装箱档）

`d121` 那一行是三个读数（`VarType(CCur(1.5))` / `VarType(CDbl(1.5))` / `VarType(1)`），
真组件 = `6 / 5 / 2`。修前 C3 = `5 / 5 / 3`。两根因，都在**装箱站**：

| 格 | 根因 | 处置 |
| --- | --- | --- |
| Currency | C3 用 `double` 承载 Currency 值（`vb6_CCur`），装箱落 `_Generic` 的 `double:` 档 ⇒ VT_R8 | 新增 `vb6_VariantCur(double)`（VT_CY，10000 倍定点）；`boxToVariant` 加 `case Vb6Type::Currency` |
| Integer | `inferExprType` 对字面量一律答 Long（C 侧 int32 的口径）；符号表把 `As Integer` 也记成 Long（真宽度在 `knownIntVars_`），于是 `case Long/Integer` 一律装 VT_I4 | `wrapVariantValue` 的 `case Integer` 改发 `vb6_VariantInt((int16_t)(…))`；字面量/`As Integer` 两形由 `isNarrowIntExpr` 判出、交给装箱权威 `boxToVariant` 装 VT_I2 |

判据边界（写清是为了下次不再讨论）：

- **只动装箱站**，`inferExprType` 一个字没改 —— 那张表同时喂算术/收窄/格式化，改它等于把
  「16 位算术 + 溢出检查」整层一起打开（那是「算术全链」，另排期）。
- 字面量那一半只认**纯十进制**书写（`rawText` 除数字不许有别的东西）；`&H`/`&O` 进制字面量、
  类型后缀（`%`/`&`/`^`）维持现状。
- 不用 `tryEvalConstInt`：那是常量折叠（`2 ^ 10` 也折成 1024，而 VB6 的 `^` 恒返 Double），
  拿它当字面量类型会把幂/位运算的档一起改掉。

探针读数（15 行，cscript 对照，临时 `.temp/litprobe`，不入库）：

```
p01 lit1  2/Integer  ✔（修前 3/Long）      p11 ccur       6/Currency ✔（修前 5/Double）
p02 32767 2/Integer  ✔                     p12 slot_ccur  6/Currency ✔
p03 32768 3/Long     ✔                     p13 cdbl       5/Double   ✔
p04 -1    2/Integer  ✔                     p14 ccurround  6 / 1.2346 ✔
p05 40000 3/Long     ✔                     p15 slot_mul   5/Double   ✔
p08 slot_lit  2/Integer ✔                  p06 &HFFFF  3（真组件 2）   ← 留白
p10 slot_intvar 2/Integer ✔                p09 v=1:v=v+1 3（真组件 2） ← 留白
```

两处留白（同一层但另排期）：

1. ~~**`&H` 进制字面量的宽度**~~ **已修前半（第三十三刀，见 §11.12）**：`&H` 那一形已按 VB6 的
   Integer 档装箱。**口径修正**：C3 的**值**早在第313刀就按 16 位补码折对了（`&HFFFF` = -1），
   缺的只是**档** —— §11.10 当时写的「C3 答 Long(65535)」只对了前半（那是第313刀之前的读数）。
   剩下的是 `&O`/`&B` 两形（无真组件读数，不敢按 hex 类推）。
2. **Variant 算术结果的宽度**：`v=1: v=v+1` 真组件仍 Integer，C3 答 Long（`vb6_VarAddV` 的
   结果档由 RTL 算术定）。这就是「算术全链」那一层的第一格。

影响面（发码清单，A/B 逐行看过）：49 份输入的发码真的变了，形态只有一种 ——
`vb6_VariantLong(<字面量>)` → `vb6_VariantInt((int16_t)(<字面量>))`（3 份同时含
`vb6_VariantCur`）。另有 5 份 GUI 工程（`GUISample`×3、`MsCommApp`、`DisplayProject`）
在**改前就已与 CI 清单不一致**（old/new 两枚二进制发码哈希完全相同，只与 expected 不同）
—— 本机环境漂移，不是本刀造成，登记时不得覆盖。

### 11.11 CI 反馈：三个夹具期望原本钉的是旧答案（已按真 VB6 改口）

第一轮 CI（run `43886597241`）读数：`Baseline (M1)` 绿、`Build` 绿；红的除了预期的
`Emit manifest`（49 份形状变化，等登记），还有 5 个格子 —— 全部是**夹具期望写着旧答案**，
不是新码算错。逐格核过 oracle 口径（VBScript 与 VB6 在「整数字面量装什么档」上一致：
`VarType(1)`=2、`TypeName(1)`=`Integer`）：

| 格子 | 旧期望 | 新期望 | 为什么新的是真 VB6 |
| --- | --- | --- | --- |
| `BD-variant`（`test_bodydecl`） | `Long/String/1` | `Integer/String/1` | 夹具里是 `Dim vA, vB: vA = 1` —— 小整数字面量进 Variant 就是 VT_I2 |
| `B165-long`（`test_host_decimal`） | `3\|42` | `2\|42` | 夹具里是 `Dim l As Variant: l = 42`，同一根因 |
| `B169-4`（`b169_class_field_store` 两台） | `ggg\|9\|3\|1` | `ggg\|9\|2\|1` | `With … .PV = 5`，`5` 是 Integer |
| `b168_class_var_field_shape`（在场针） | `= vb6_VariantLong(7);` | `= vb6_VariantInt((int16_t)(7));` | 字面量 7 的发码形状本来就该跟着装箱档走 |
| `b169_class_field_store_shape`（在场针） | `_vb6_with_0->PV = vb6_VariantLong(5);` | `_vb6_with_0->PV = vb6_VariantInt((int16_t)(5));` | 同上 |
| `NC-e01`（`jsonfast_nd_variant_com_roundtrip`） | `3/42` | `2/42` | 夹具里是 `a2(0, 1) = 42` —— Variant 数组元素装字面量 42 = VT_I2 |

读数自证（本机、新码）：`BD-variant=Integer/String/1`、`B165-long=2|42`、
`B169-4=ggg|9|2|1`、`NC-e01=2/42`；发码两侧各取 `--emit-c` 逐字核过在场针的新形。**没有回头改夹具源码**
（注释与读数都在 runner 里，夹具本身一字未动）—— 期望该记的是真 VB6 的答案，不是 C3 上一版
算出来的答案。

另有一格**不是期望值、而是登记面漏登**：`[STATIC] owned_temp_keeps` 的 K13/K13c 红在
`vb6_VariantCur` —— 新出口交回一枚**新造且归调用方所有**的 VARIANT，却没进
`src/common/owned_temp_keeps.hpp` 的 `kOwned`/`kOwnedVariant`（那两份是手写面，哨兵 K13 双向对账）。
补登后本地复跑 `scripts/check_owned_temp_keeps.ps1` = PASS（owned 121 / K13/K13b/K13c two-way）。

### 11.12 第三十三刀：无后缀 `&H` 的**档**（离白 1 的前半）

**病灶口径修正**：这一格原来记的是「C3 答 Long(65535)」—— 只对了前半。第313刀之后
`&HFFFF` 的**值**已按 16 位补码折成 -1（`radixTierOf`），缺的只是**类型身份**：
改前探针 `h01|lit_ffff vt=3 tn=Long val=-1`（值对、档错）。于是真组件与 C3 的差只在
`VarType`/`TypeName` 上 —— 比原估计窄一格。

判据（VB6 规则 — 与无后缀十进制同一条「装得下的最小有符号类型」口径）：

| 写法 | VB6 档 | 值 | 实证 |
| --- | --- | --- | --- |
| `&HFF` / `&H7FFF` / `&H8000` / `&HFFFF` / `&HBF32` | Integer (VT_I2) | 255 / 32767 / -32768 / -1 / -16590 | 真组件对 `&HFFFF` 的读数 = Integer(-1) |
| `&H10000`（5 位） | Long | 65536 | 位宽 > 0xFFFF，不参加 16 位折叠 |
| `&HFFFFFFFF`（8 位） | **Long** | -1 | 值也在 int16 范围内，但档是 Long —— 所以判据必须看**位宽**而不是值 |
| `&HFFFF&`（带 `&` 后缀） | Long | 65535 | `tests/intsuffix` 的证人；后缀不改 |

**落点：一位主人，两次搬运，不抄第二遍**

| 位置 | 改了什么 |
| --- | --- |
| `src/lexer/lexer_number.cpp` | `RadixTier` 多记一个 `narrow16`（只在真的走过 16 位补码折叠那一支为 true）；`makeRadixToken` 在 `tier.narrow16 && tagInt16` 时置 `tok.radixInt16` |
| `src/lexer/lexer.hpp` / `token.hpp` | `makeRadixToken` 多一个 `tagInt16` 形参（只有 `&H` 那一支传 true）；`Token` 多一枚 `radixInt16` |
| `src/parser/parser_expr.cpp` | `parseLiteral` 原样搬到 `LiteralExpr::radixInt16` |
| `src/ast/detail/ast_expr.hpp` | `LiteralExpr` 多一枚 `radixInt16` |
| `src/backend/cgen_util_ctrl.cpp` | `isNarrowIntExpr` 只问那一枚标记（不拿 rawText 重算规则） |

为什么不在发码侧重算：那等于把 `radixTierOf` 抄第二遍，而第313刀的教训正是「抄一次漏一次」。
也不进 `inferExprType`：那张表同时喂算术/收窄，改它就是把「16 位算术 + 溢出检查」整层打开。

探针读数（`.temp/litprobe/HexProbe.bas`，17 行，改前 → 改后）：

```
h01 &HFFFF      3/Long/-1  →  2/Integer/-1  ✔      h06 &H10000   3/65536  不动 ✔
h02 &H8000      3/-32768   →  2/-32768     ✔      h07 &HFFFF&   3/65535  不动 ✔（后缀证人）
h03 &H7FFF      3/32767    →  2/32767     ✔      h08 &O177777  3/-1     不动（离白）
h04 &HFF        3/255      →  2/255       ✔      h09 &O100000  3/-32768  不动（离白）
h05 &HBF32      3/-16590   →  2/-16590    ✔      h12 slot_10000 3/65536 不动 ✔
h10 slot_ffff   3/-1       →  2/-1        ✔      h13~h16 算术与转换逐字不变
```

影响面：**发码零变更**。CI run `43929396923` 的 `Emit manifest` 绿（748 份输入逐字相同），
运行期也全绿 —— 22/22 success。语料里「≤ 4 位的 `&H` 字面量进 Variant 槽」那一形真的
一处也没有（2120 处命中全是颜色、掩码、位运算与常量）。所以这一刀只有探针能看见它的效果，
这也正是它敢开的理由：宽面（161 份文件）里没有一格落在装箱站上。

同层剩下两格（已登记为离白）：

1. `&O` / `&B` 的档 —— 值已按 16 位折叠（`&O177777` = -1），档仍 Long；真组件对这一形**没有读数**，
   不按 hex 类推（`tagInt16=false` 就是那两道闸，见 `lexer_number.cpp` 两处注释）。
2. `-&HFFFF` 这类**一元负号包着**的形仍按 Long 装（`isNarrowIntExpr` 只问裸字面量节点）；
   VB6 应是 Integer(1)。要收就得给 `vb6IntConstantValue` 加一条进制的负号路 —— 目前语料 0 处，暂不开。

### 11.13 第三十四刀：用户 `Const` 的档（上一刀探针扫出来的）

上一刀收线后，同一层「标量宽度」再用一对探针（`.temp/litprobe/W2.vbs` = cscript 基准、
`W2.bas` = C3）扫了 16 形，扫出三处真缺口 —— 值都对，只有**档**不对：

| 读数 | 写法 | oracle (cscript) | C3 | 处置 |
| --- | --- | --- | --- | --- |
| w01 | `Const C1 = 1` | 2/Integer | 3/Long | **本刀修** |
| w03 | `Const C3 = &HFFFF` | 2/Integer | 3/Long | **本刀修**（吃上一刀的 `radixInt16`） |
| w07 | `VarType(Asc("A"))` | 2/Integer | 3/Long | 留白（内建返回档，另排期） |
| w18 | `VarType(Array(1,"x")(0))` | 2/Integer | 3/Long | 留白（数组字面量的元素装箱走另一条路） |

其余 12 形（CInt/Asc 以外的内建、CBool/CByte/CCur/CLng、Len/UBound/InStr、Round/Fix、`32768`）
两边逐字相同 —— 它们本来就对。（另：`CByte(300)` 在两边都抛 Overflow 错误 6，C3 那一台打的是
同一条 `Unhandled VB6 Error #6`，这形不算缺口。）

**判据（VB6 常量取型）**：`Const K = <初始化器>` 的 VB6 类型 = 初始化器字面量的类型；
写了 `As <型>` 就按那个型。于是 Integer 档共有四种写法，本刀全接：

| 写法 | 档 | 依据 |
| --- | --- | --- |
| `Const K = 1` / `= -1` | Integer | 小十进制字面量（`vb6IntConstantValue`，含一元负号） |
| `Const K = &HFFFF` | Integer | 无后缀 `&H` 且位宽 ≤ 0xFFFF（上一刀的 `radixInt16`） |
| `Const K = 3%` | Integer | `%` 后缀是显式 Integer（账 #195 的记号） |
| `Const K As Integer = 7` | Integer | 显式类型子句（`resolveArrayElemType`） |

反证（必须保持 Long）：`Const K = 32768`、`Const K = 40000`、`Const K = 1&`、`Const K As Long = 1`
—— 四条都在探针 w02/w05/w07/w09 里逐条钉住。

**落点（只喂装箱站，两层名单）**

| 位置 | 改了什么 |
| --- | --- |
| `cgen_state.inc` | 新增 `knownIntConsts_`（本过程可见）与 `moduleIntConsts_`（模块级持久，供回灌） |
| `decl/cgen_decl.cpp` | 模块级 `visit(ConstDecl&)` 末尾登记；`As Integer` 与「初始化器本身是 Integer」两条判据，后者直接问 `isNarrowIntExpr`（不在这里重算） |
| `decl/cgen_localdecl.cpp` | 局部 `Const` 同一口径登记 |
| `decl/cgen_decl_{func,proc,prop}.cpp` | 过程入口 `knownIntConsts_.clear()` + 从 `moduleIntConsts_` 回灌（与 `knownNewVars_` 同一套生命周期） |
| `cgen_util_ctrl.cpp` | `isNarrowIntExpr` 的标识符那一支多问一张表；`vb6IntConstantValue` 剥掉**一个**尾随 `%`（`&`/`^` 仍挡） |
| 同上 | 副产物：`VarType(3%)` 从 3 变 2（显式 Integer 字面量本来就该是 2） |

**为什么不进 `knownIntVars_`**：那张表同时喂 `inferExprType`、收窄检查与算术。把常量塞进去等于
把「16 位算术 + 溢出检查」一起打开（例如 `Const K = 30000: x = K * 2` 从现在的 60000 变成
溢出错误）—— 那是对的 VB6 行为，但属「算术全链」，不属于本刀。本刀只回答「装箱时这枚名字
是什么档」。

**读数**：`W3.bas`（15 行，本刀新增的 13 形 + 两个反证）：w01/w03/w04/w06/w08/w10/w12 全 2，
w02/w05/w07/w09/w11/w13/w14 全按 oracle；`Probe.bas` 15 行与 `HexProbe.bas` 17 行逐字未变；
`dict_native` PASS 59 行。发码：CI run `43933863593`（head `bda14980`）**22/22 全绿**，
`Emit manifest` 也绿 —— 748 份输入逐字相同（语料里「常量进 Variant 槽」那一形同样极少）。

同层剩下两格（已登记）：`Asc()` 的返回档、`Array()` 字面量的元素档 —— 两者都属「内建返回类型表」
那一层，需要先看那张表还喂谁（与 `inferExprType` 同源），不能拿「只改装箱」的刀法套。

### 11.14 第三十五刀：内建返回档那一族（`Asc`/`Sgn`/`VarType`/`StrComp`/月日年时分秒/`Weekday`/`FreeFile`）

先把上一刀留下的两张留白摊开量：`.temp/litprobe/W4.vbs`（cscript 基准）↔ `W4.bas`，
25 形逐行对。结论不是「两枚函数答错」，而是**一整族**内建在 VB6 里返回 Integer、C3 全按 Long 装：

| 读数 | 写法 | oracle | C3（修前） |
| --- | --- | --- | --- |
| b01 | `Asc("A")` | 2/Integer | 3/Long |
| b03 | `Sgn(-5)` | 2/Integer | 3/Long |
| b04 | `VarType(1)` | 2/Integer | 3/Long |
| b05 | `StrComp("a","a")` | 2/Integer | 3/Long |
| b08~b12 | `Month`/`Day`/`Year`/`Hour`/`Weekday` | 2/Integer | 3/Long |
| b17 | `FreeFile` | 2/Integer（VB6 文档；VBScript 无此函数，不进 oracle） | 3/Long |

**反证（都已是 3/Long，不许拉进来）**：`InStr`、`InStrRev`、`RGB`、`UBound`、`Len`
—— b06/b07/b16/b22/b23 本来就对。非整数族也全对：`IsNumeric`/`IsArray`=11 Boolean、
`TypeName`=8 String、`Timer`=4 Single、`Abs`/`Int`=5 Double、`Left`=8 String。

**为什么不动语义层 `builtin_funcs.inc`**（那里这十四行也确实写着 `Vb6Type::Long`）：
`inferExprType` 对它们答的 Long 是「C 侧 int32 口径」，同时喂算术 / 收窄 / `Print#` 包装。
把档改窄会把「Integer 参与算术时 VB6 的升格规则」整层打开（`*` 升 Long、`+` 留 Integer、
溢出错误 6）—— 与 §11.13 同一道闸：那属「算术全链」，不属本刀。

**落点**：`detail/expr/cgen_builtin_return_type_table.inc` 新增 `builtinReturnsVbInteger`（十四名一张
`unordered_set`，剥 `$` 后小写比）+ `CCodeGen::builtinReturnsVbIntegerName` 成员（声明在
`detail/util/cgen_helpers.inc`，紧挨 `isNarrowIntExpr`）；`cgen_util_ctrl.cpp::isNarrowIntExpr`
多两支：`IndexOrCallExpr` 的 callee 是标识符且在表里（`Asc("A")` 这一形）、以及裸标识符
（`FreeFile` 这一形，要求符号表里不是用户变量）。⚠ 该表是**过渡期的第二份答案**，只喂装箱站；
算链落账后应把语义层那十四行改成 Integer、删掉这个函数 —— 文件头注已写这句。

**本刀踩到并修掉的一处真 bug（值得记）**：`isNarrowIntExpr` 原先写的是
`if (kind == IdentifierExpr) return knownIntVars_||knownIntConsts_;` —— 提前 `return`，于是
后面新加的「内建标识符」支**永远走不到**。实测就是 `v = FreeFile` 仍答 3（而 `FreeFile()` 答 2），
这正是 `W5.bas` 的 f01/f02 两行之差把我逼出来的。现在标识符这一支**只留一个出口**，三源
（Integer 变量 / 用户 Const / 内建返回档）一次答完。

**读数**：`W4.bas` 25 行里 b01/b03~b12/b17 全转 2/Integer，反证六行照旧；`W5.bas`
f01/f02/f03 = 2、f04（用户 `Long` 变量）= 3；`Probe.bas` 15 行、`HexProbe.bas` 17 行逐字未变；
`W3.bas` 15 行里只 w13（`Asc`）从 3 转 2，其余不动；`dict_native` PASS 59 行。
发码：CI run `43940956600`（head `6c60a6c5`）**22/22 全绿**，`Emit manifest` 也绿 ——
748 份输入逐字相同（语料里「内建整数结果直接进 Variant 槽」同样极少）。

同层**只剩一格**：`Array(1,"x")(0)` 的元素档（b19/b20 实测仍 3/Long）。它不是「返回档」问题，
而是数组字面量的**元素装箱**走了另一条路（`cgen_expr_array.cpp`），下一刀从那里入手。

### 11.15 第三十六刀：`Array()` 字面量的元素档（用户实测补的一刀）

用户真 VB6 读数：`TypeName(Array(1, "x")(0))` 答 **`Integer`**（VarType 2）—— 与 `W2`/`W4`
里 cscript 的读数一致，把上一刀留的最后一格坐实。

**判据**：`Array(arglist)` 的每个元素保留它在 VB6 里的类型；「Integer 档」这件事**不在本处重写**，
只问装箱权威 `isNarrowIntExpr` —— 于是自动继承了前四刀的积累：小十进制字面量（含一元负号）、
`&H` 且位宽 ≤ 0xFFFF、`3%` 后缀、`As Integer` 变量、Integer 档用户 `Const`、
Integer 返回档的内建那一族（`Asc`/`Sgn`/`VarType`/月日年时分秒…）。

**落点**：`detail/expr/cgen_expr_call_builtin_pre.inc` 的 `Array` 分支。原来整数元素落
`isLongArg ⇒ vb6_ArraySetLong`（VT_I4）；现在在选 setter 之前加一闸，只往 `variantSetter`
里填 `vb6_VariantInt((int16_t)(expr))`，由本块**唯一**那条 `ArraySetVariant` 发射行带出去。

**踩到的两条硬约束（都是先红后改）**：

1. 第一版自己 `emitLine("vb6_ArraySetVariant(...)")` 直接发 ⇒ 静态哨兵 R2
   （`scripts/check_array_literal_variants.ps1`，要求该出口恰好 1 处）红成 `call sites = 2`。
   处置**不是放宽阈值**，而是复用同一个出口。
2. 改完仍红：因为那句注释里写了 `vb6_ArraySetVariant(` 字面量，被正则当成第二个调用点。
   注释里把括号去掉才绿 —— 记一笔：这类按正则数的哨兵，注释也算数。

**读数**（`W6.bas` 17 行；oracle = `W6.vbs` 13 行）：

| 读数 | 写法 | VB6 / cscript | 修前 | 修后 |
| --- | --- | --- | --- | --- |
| g01 | `Array(1,…)(0)` | 2/Integer | 3 | **2** |
| g02 | 第二个元素 `2` | 2/Integer | 3 | **2** |
| g03 | `32768` | 3/Long | 3 | 3 ✔ |
| g04 | `&HFFFF` | 2/Integer(-1) | 3 | **2** |
| g05/g06/g07 | `"x"` / `1.5` / `True` | 8 / 5 / 11 | 同 | 同 ✔ |
| g08 | `Array(i)`，`i As Integer` | 2/Integer | 2 | 2 ✔ |
| g09 | `Array(l)`，`l As Long` | 3/Long（VBScript 那行不可比：它的 `l` 是无类型 Variant） | 3 | 3 ✔ |
| g10/g11 | `3%` / `1&` | 2 / 3 | 2 / 3 | 2 / 3 ✔ |
| g12 | 用户 `Const KC = 5` | 2/Integer | 2 | 2 ✔ |
| g13 | `Asc("A")` | 2/Integer | 2 | 2 ✔ |
| g15 | `-1` | 2/Integer | 2 | 2 ✔ |
| g14 | `1 + 1` | 2/Integer | 3 | 3 ✗ 留白 |
| g16 | `Array(1)(0)` 嵌在 Array 里 | 2/Integer | 3 | 3 ✗ 留白 |

g14/g16 都不在装箱站，各自属于下层（算术全链 / 数组元素类型表），已登记。
`W4.bas` 的 b19/b20、`W3.bas` 的 w14 同步从 3 转 2；`Probe.bas` 15 行、`HexProbe.bas` 17 行
逐字未变；`dict_native` PASS 59 行。

**随之改的两处夹具（按 cscript 口径，不是「改期望让门变绿」）**：

* `tests/run_tests.ps1` 的 `jsonfast_array_literal_variants`：`AL-long` 与 `AL-i4` 3 → 2
  （`.temp/litprobe/AL.vbs` 实测就是 2、c(1)=8、c(2)=5 均未变）；
* 同一夹具 `.bas` 头注释里那句「plain Long arm 是零回归护栏」不再成立，改成记录本刀与两条 cscript 读数。

**发码**：11 份输入形状变了。本机做了 A/B（`.temp/C3_pre36.exe` = `dd6cdfbf` 那一刀 vs 新二进制，
逐输入 `--emit-c` 逐行 diff，脚本 `.temp/ab36/run.ps1`）：`test_arr_empty` 3/3 行、
`test_joinfilter` 3/3、`test_split_result_subscript` 3/3、`test_whole_array_copy` 3/3、
`P2Main` / `p2_array_variant.vbp` 各 12/12、`ArrayLiteralVariants` 4/4、三份 GUISample 各 2/2 ——
**每一行都是 `ArraySetLong(x, i, N)` → `ArraySetVariant(x, i, vb6_VariantInt((int16_t)(N)))`，无其它差异**。
确认后 `rebless_emit_manifest.ps1 -Bless` 改写 11 行（数据行 748 不变、复算绿、coverage K1..K3 绿）。
三份 GUISample 行也是真变了（`Array(..., -1)` 站点），不是本机环境漂移。

**CI**：`43944415872`（head `1b093a9d`）= 3 红（manifest 11 行 + 夹具 `jsonfast_array_literal_variants`
输出不符 + R2 哨兵）；修完 `43947217474`（head `d4c5e479`）= 只剩 manifest 那格红（已预料）；
登记后 `43949530645`（head `35eee3a2`）= **22/22 全绿**。

同层还剩两格（都不在装箱站）：`Array(1 + 1)` 元素（算术全链）、`Array(Array(1)(2))(0)` 的索引结果档
（数组元素类型表）。

### 11.16 第三十七刀：`Array()` 的整只 Variant 实参保真（修的是**丢值**，不只是档）

继续同层往下量（`W7.vbs` ↔ `W7.bas`，11 行）挖出的不是「档不对」而是**值真的丢了**：

| 读数 | 写法 | oracle(cscript) | 修前 | 修后 |
| --- | --- | --- | --- | --- |
| h05 | `Array(vi)`，`vi` 是装 Integer 的 Variant | 2/Integer | 3 | **2** |
| h06 | `Array(vs)`，`vs` 是装 `"x"` 的 Variant | 8/String | **3/Long（值变 0）** | **8/String** |
| h07 | `Array(vd)`，装 1.5 | 5/Double | 3 | **5** |
| h08 | `Array(vl)`，装 70000 | 3/Long | 3 | 3 ✔ |
| h04 | `Array(Array(1)(0))` | 2/Integer | 3 | **2** |
| h01/h02 | `1 + 1` / `32767 + 0` | 2/Integer | 3 | 3 ✗ 算术全链留白 |

**根因**：`argIsVariant` 那一支一律 `vb6_ArraySetLong(…, vb6_VariantToLong(expr))`，
靠一个 `looksLikeBSTR` 猜字符串；猜不中就把整只 Variant 压成 int32 —— 装字符串的那格
连值都没了。

**改法**：把「实参是不是整只 Variant」这条问句**提到出口之前**，命中时装
`variantSetter = argExpr`，与第三十六刀的 Integer 档共用同一条 `ArraySetVariant` 发射行
（哨兵 R2 仍 1 处）；下游那个 `else if (argIsVariant)` 臂整段删掉。`ArraySetVariant` 按值收
`vb6_VARIANT`、内部 VariantCopy ⇒ 子类型与 BSTR 引用计数都对。

**读数**：`W7.bas` 11 行里 h04/h05/h06/h07 四项转正，h08/h09/h10 不动，h01/h02 仍留白；
`W6.bas` 里只 g16（嵌在 Array 里的索引）从 3 转 2、其余 16 行未动；`W4`/`W3`/`W5` 三份逐字未变；
`Probe.bas` 15 行与 `HexProbe.bas` 17 行未变；`dict_native` PASS 59；R2 哨兵 PASS。

**发码**：CI run `43953077180`（head `a9659cf0`）**22/22 全绿**，`Emit manifest` 也绿 ——
748 份输入零变更（语料里没有 `Array(<Variant 表达式>)` 这种站点；有这一形的代码本来就很少）。
备查：本刀做了 k36/k37 两台二进制（`.temp/C3_k36.exe` / `C3_k37.exe`，用 W7 的 h05/h06 钉住
到底是哪一台），将来真需要 A/B 时可直接拿。

同层从现在起**只剩一格**：`Array(1 + 1)` 这类「Integer 参与算术后进槽」——它不在装箱站，
要的是「Integer 算术升格规则」（`+` 留 Integer、`*` 升 Long、溢出错误 6）整层，另排期。

### 11.17 第三十八刀：`Int`/`Fix`/`Abs` 的返回档 = 实参的档

算术那一格不动手，先把**挨着它的非算术位**量一遍：`.temp/litprobe/W8.vbs` ↔ `W8.bas`（19 行）。
结果分三类：

| 类别 | 读数 | 结论 |
| --- | --- | --- |
| 转换族全对 | `CInt(1.5)`=2、`CInt("7")`=2、`CByte`=17、`CLng`=3、`CSng`=4、`CDbl`=5、`CCur`=6、`CDate`=7、`CStr`=8、`CBool`=11 | 两边逐字相同，不动 |
| **缺档（本刀）** | `Int(5)`=2、`Fix(5)`=2、`Abs(-5)`=2（而 `Int(1.7)`=5、`Abs(1.5)`=5） | C3 一律 5/Double |
| 不动（算术） | `i + 0`、`i * 1` | 仍 3/Long，属算术全链 |
| 不动（运行时） | `Abs(i)`，`i` 是装 5 的 Variant | VB6=2，C3=5：子类型要在运行时才有，静态不可知 |

**判据（VB6 口径）**：`Int`/`Fix`/`Abs` 的返回类型 = **实参的类型**，不是固定 Double。

**落点**：`cgen_util_ctrl.cpp::isNarrowIntExpr` 的 `IndexOrCallExpr` 支——名字是
`int`/`fix`/`abs`、恰好 1 个位置实参、无命名实参时，**问实参自己**（`isNarrowIntExpr(arg0)`），
不另建表。于是自动继承：字面量 / `&H`≤0xFFFF / `%` 后缀 / `As Integer` 变量 / 用户 `Const` /
Integer 返回档内建。Variant 实参（`Abs(v)`）静态答不出，保留原样。

**读数**：`W8b.bas` 的 a03/a04/a05 → 2/Integer，a06（Variant 实参）仍 5，a09/a10 仍 3，
其余 13 行未变；`W4.bas` 里只有 b24（`Abs(-1)`）从 5 转 2，b25（`Int(1.7)`）仍 5；
`W6`/`W3`/`W5`/`Probe`/`HexProbe` 五份逐字未变；`dict_native` PASS 59；R2 哨兵 PASS。
发码：CI run `43958409386`（head `1da8632c`）**22/22 全绿**，`Emit manifest` 也绿 ——
748 份输入零变更（语料里没有「Int/Fix/Abs(Integer 实参) 直接进 Variant 槽」的站点）。

### 11.18 同批侦察挖到的两个真缺陷（未修，已登记）

**① `-v` / `Not v`（v 是 Variant）根本编不过。** `W8.bas` 第一次编译就死在第 28 行：
`error C2088: 「-」: 对于 struct 是非法的` 与 `C2198: vb6_VarType 参数太少`——
也就是一元 `-` 直接打在 `vb6_VARIANT` 结构上，C 侧根本没有这条路。
RTL 里也**没有** `vb6_VariantNeg` / `vb6_VariantNot` 这类出口（全仓 grep 为 0）。
修它要新开一条 RTL 面（含 `owned_temp_keeps` 与签名表登记），属「Variant 参与运算符」层，
不是本系列这种「只改装箱站」的刀法，另排期。

**② `IIf(True, 1, 2)` 答 3/Long。** VB6 的 `IIf` 返回 Variant，子类型随选中的那一支（
本例应为 Integer），C3 答 Long。VBScript 没有 `IIf`，本机拿不到 oracle，所以这条得先用真 VB6
确认再动（`IIf` 是 cgen 里的多态内置，固定成任何一档都会答错，见
`cgen_builtin_return_type_table.inc` 头注那句）。

### 11.19 第三十九刀：`IIf` 两支都是 Integer 档 → VT_I2

落点：`isNarrowIntExpr`（`cgen_util_ctrl.cpp`）的 `IndexOrCallExpr` 支。判据只有一句：
`iif` + 三个位置实参 + 无命名实参，且 `isNarrowIntExpr(第 2 支) && isNarrowIntExpr(第 3 支)`。
**不**把 `iif` 塞进 `builtinReturnsVbInteger` —— 那会把 `IIf(True, 1&, 2&)` / `IIf(True, 1.5, 2.5)`
一起窄化。

`W9.bas` 读数（`.build\k39_W9`）：i01/i02/i07/i08/i09 = 2/Integer，i03(mixed) = 3/Long，
i04(str) = 8/String，i05(dbl) = 5/Double，i06(`1&`) = 3/Long。`W8b` a19 3→2；`W6`/`W3`/`W5`
一个字没动；`dict_native` PASS 59；R2 哨兵 PASS。

CI run `43961533512` = **21/22**，唯一红的是 `Emit manifest`（`tests/test_builtin_arity.bas`
的 `v = IIf(True, 1, 2)` 发了 `vb6_VariantInt((int16_t)(vb6_IIfLong(...)))`）—— 当轮没登记，
在 §11.20 一并登记。

### 11.20 真 VB6 采集（138 行）+ 刀40/41 落地、刀42 撤回

#### 11.20.1 采集器

`C3Probe\ProbeVB6.bas`（标准 EXE 直接跑，日志落 `App.Path\c3_vb6_probe.log`），
九段 A~H 共 138 行读数：`KEY vt=<VarType> tn=<TypeName> val=<值>`，出错行落 `KEY ERROR n <描述>`。
用户回传的日志存于 `C:\Users\Administrator\.sclaw\workspace\C3Probe\c3_vb6_probe.log`。

⚠ 采集器自身两个坑（已在文件里规避）：① 会溢出的那几条必须走**变量**，VB6 会把
`30000 + 30000` 这种常量在编译期折叠并直接报错，整份日志都跑不起来；② `F07`(200*200)、
`F08`(30000*2) 两行当时漏写 `Oops` 检查，溢出发生在实参求值、**打印前**就跳走了，
所以日志里「缺行」= 该条是 ERROR 6，不是没测。

#### 11.20.2 三个被推翻的旧结论

| 旧说法 | 真 VB6 |
|---|---|
| `*` 把 Integer 升成 Long | ❌ Integer × Integer 仍是 Integer，`200*200`/`30000*2` 直接**错误 6** |
| `VarType` 返回 Integer | ❌ 返回 **Long**（C04 vt=3）。cscript 同形答 2/Integer —— VBScript 与 VB6 在这一枚上不一致，以 VB6 为准 |
| `Int`/`Fix`/`Abs` 已修好 | ⚠ 只修到「实参站」；`t = Int(5)` 的**存储站**仍答 Double |

第三十八刀的读数当年取自 `Probe(key, expr)` 的 `ByRef Variant` 实参（`wrapVariantValue` 的实参分支），
而 `t = expr` 走的是同一函数里 `switch (inferExprType)` 那条，`case Double` 不查窄化 —— 两站分叉。
这轮的对照表用**两种写法各跑一遍**才把差异显出来（30 个缺口全部来自 `t = expr` 路径）。

#### 11.20.3 刀40：`vartype` 退出 Integer 表

`cgen_builtin_return_type_table.inc` 的 `builtinReturnsVbInteger` 删掉 `"vartype"`。
依据是 C04（`t = VarType(1)` → vt=3/Long）。同一张表其余十三枚逐条被日志坐实
（Asc/AscW/Sgn/StrComp/月日年时分秒/Weekday/FreeFile → Integer；InStr/InStrRev/Len/UBound/RGB → Long）。

#### 11.20.4 刀41：存储站与实参站共用同一柄尺

`wrapVariantValue` 在 `switch (inferExprType)` **之前**先问一次 `isNarrowIntExpr`，命中就发
`vb6_VariantInt((int16_t)(...))`。与末尾 `boxToVariant` 的字符串逐字相同，所以 `case Long` /
`case Integer` 两档**发码不变**；只有 `Int`/`Fix`/`Abs` 这族从 VT_R8 变 VT_I2。

A/B 实证（`.temp\C3.old.exe`=HEAD `32148205`，`.temp\C3.new.exe`=本刀）：语料里 748 份输入
**一份都没有** `t = Int(5)` 这类站点 → 发码零变化。`dict_native` PASS 59、R2 哨兵 PASS。

#### 11.20.5 刀42（撤回）：`IIf` 混支 → `vb6_IIfVariant`

真 VB6 的口径是「选哪支给哪支的原生子类型」：D03 `IIf(True,1,40000)` = Integer、
D09 `IIf(True,Null,1)` = Null、D10 = Integer、D12~D15 的 Choose/Switch 同理。
第一版实现把「两支窄化与否不同 / 某支是 Null 字面量」的 IIf 改发 `vb6_IIfVariant(cond, box(a), box(b))`，
本机读数确实全对（30→21），但 CI `43968071126` 的 `Tests (vbp #5)` 红了：

```
tests/ucTreeView → ucTreeViewPlus.c(4415,4416): error C2440: cannot convert from 'vb6_VARIANT' to 'double'
```

根因：`IIf` 一旦是 Variant，夹在算术/数值内建里的用法会整体走 Variant 算术
（`vb6_VarAddV/vb6_VarSubV`），而**消费站**（要求 double 的那侧）不会自动取回具体值。
这层要统一收在消费站，不能靠逐点补 `vb6_VariantToDouble` 竞猜。
故撤回，补丁留档 `.temp\k42_iif_variant.patch`；撤回后 `ucTreeView` 本机 `exe=True` 编译通过。

后续这一刀的正确开法：先落「消费站把 Variant 实参取成目标档」这一条唯一出口，再让 IIf 走 Variant。

#### 11.20.6 缺口账（126 键）

30 → 21。已对齐：`VarType`、`Int/Fix/Abs`（实参站+存储站）、`IIf` 两同一型支、`Array()` 全段。
剩余 21 键按成因分五组，就是下一轮的刀序：

| 组 | 键 | VB6 | C3 |
|---|---|---|---|
| 「选中支原样」 | D12/D13 Choose、D14/D15 Switch（+ D03/D09/D10 需消费站） | Integer / Null | Long |
| Integer 算术 + 溢出错误 6 | F01~F04、F06、F06b、F09、F11、F12、F14、F20、G09、H01、H02 | Integer，`30000+30000` 报 6 | Long |
| `^` 返回 Double | F13 | Double | Long |
| Single 装箱 | A12 `1.5!`、C19 `Timer`、C28 `CSng` | VT_R4 | VT_R8 |
| `Abs(Variant)` 运行时子类型 | C23 | Integer | Double |
| 一元 `-`/`Not`（整段编不过） | E01~E10、H06 | `-v` 保子类型、`Not v`→Integer(-6) | 编译失败 |

#### 11.20.7 CI 记录

- `43968071126`（含刀42）= 20/22：`Emit manifest` + `Tests (vbp #5)`(ucTreeView)
- `36a85f19`（撤回刀42，只留 40/41）= **21/22**：仅 `Emit manifest` 红
- 登记提交 `14579ca2`：`emit-manifest.expected.txt` 1 行（第三十九刀的遗留）

### 11.21 第四十三刀：接回刀42 + `Move` 实参按 Variant 槽取 Double

刀42 被 `ucTreeView` 的 C2440 打回后，定位到的不是「IIf 该不该是 Variant」，而是**消费站缺一条取档出口**：

```c
vb6_ControlMove(hwnd, x, y, w,
    (vb6_VarSubV(vb6_VariantFromValue(A), vb6_IIfVariant(…))), 15);
//  C2440: vb6_VARIANT → double   (ucTreeViewPlus.c:4415/4416)
```

`vb6_ControlMove` 四个槽清一色 `double`，而实参是 **Variant 算术**的结果。这一形不是刀42 造的新东西：
任何 `v + 1`（v 是 Variant）喂给 typed 形参本来就会撞，只是语料里没踩到。

落点：`cgen_expr_call_callee_member.inc` 的 `Move` 特例。原来那套判据是**文本启发式**
（`find('(')==npos && find('-')==npos && find('.')==npos && find('>')==npos` + `inferExprType==Variant`），
只能接住「裸变量」，Variant 算术/选择的文本必然带括号 ⇒ 漏过。

现判据：
1. `isVariantSlot(aMVarg, nullptr)` —— 前缀表里已含 `vb6_VarAddV/SubV/MulV/DivV` 与 `vb6_IIfVariant`，
   且早已排除 `vb6_VariantTo*`，所以「这段 C 文本是整只槽」这一问只有一个出口；
2. 裸标识符 + `inferExprType==Variant`（Outlook2013 frmInbox 那种未声明类型的隐式变量，文本里没函数名可问）。

⚠ 不能把第 2 条单独推广到任意形：`Me.ScaleWidth - 200` 会被 `inferExprType` 答成 Variant，
而它明明是 `int` —— 实测 `Form5.c(240) C2440: int → vb6_VARIANT`（当时把 `vb6_VariantToDouble` 套在 int 上）。

本机验证：`ucTreeView`/`ucSlider_2022`/`ucSegmentedControl`/`VBFlexGridDemo`/`VbQRCodegen` 五份
受影响样本全 `exe=True`；`dict_native` PASS 59；R2 哨兵 PASS。

读数：仍 21 个缺口（D03/D09/D10 已回位 —— 刀42 保住了），剩的是 `Choose`/`Switch`（D12~D15）、
Integer 算术全链（13 键 + `^` 返回 Double）、Single 装箱（A12/C19/C28）、`Abs(Variant)`（C23）、
一元 `-`/`Not`（E 段 + H06）。

发码：CI run `43989114687`（head `6744bea5`）= **21/22**，仅 `Emit manifest` 红；
登记提交 `8263d2e5` 改写 10 行 → CI run `43990410456` = **22/22 全绿**。

### 11.22 第四十四刀：`Choose`/`Switch` 全窄候选支 → VT_I2（+ 两处探针失真订正）

判据（`isNarrowIntExpr` 的 `IndexOrCallExpr` 支）：`choose` 的第 1 位起、`switch` 的奇数位
**全部**是 Integer 档 ⇒ 不管运行时选哪一支，结果必是 Integer → 发 `vb6_VariantInt`。
依据：真 VB6 D12 `Choose(1,1,2)` = 2/Integer、D14 `Switch(True,1,False,2)` = 2/Integer。

**不开的一半**：D13 `Choose(1,1,40000)`、D15 `Switch(True,1,40000)`（一支宽一支窄）——
那条要求「选中支原样」，得把每支装箱、处置未选中支的 BSTR 归属；且现有发码是**嵌套三目**
（惰性求值），而 `IIf` 是悉数求值，换形会改求值时机，没有 oracle 之前不碰。发码一字不变。

发码：CI run `43992894092`（head `0146f350`）= **22/22 全绿**（`Emit manifest` 也绿 ——
语料里没有「全窄候选支」的 Choose/Switch 站点，所以发码零变化）。

#### 11.22.1 两处必须订正的读数（镜像 helper 的失真）

`Probe(key, expr)` / `P(key, val)` 的形参是 **`ByRef Variant`**。C3 在这一站把
**VT_R4 转成了 VT_R8**，于是镜像里 C19(`Timer`)、C28(`CSng(5)`) 看起来答 Double。
直接打印（不过 helper）的实测：

```
csng=4/Single
timer=4/Single
snglit=5/Double     ← 1.5! 是真缺口：发码 vb6_VariantDouble(1.5)
dbl=5/Double
```

⇒ C19/C28 不是 C3 缺口；真正的 Single 缺口只剩 `1.5!` 字面量的档（词法吃掉了 `!` 但没记下
Single 档，要词法→AST→inferExprType 一路补）。另：`ByRef Variant` 形参丢 VT_R4 本身
也是一条待查的 C3 行为（没进本刀）。

#### 11.22.2 剩余 19 键

`1.5!`（A12）、`Abs(Variant)`（C23）、D13/D15、Integer 算术全链（F01~F04/F06/F06b/F09/F11/
F12/F14/F20/G09/H01/H02 共 14）、`^` 返 Double（F13）、一元 `-`/`Not`（E01~E10 + H06）。

### 11.23 第四十六刀：Integer 算术留 Integer（逐步错误 6）

#### 11.23.1 判据（真 VB6 F 段，17 键）

直印探针 `.temp/litprobe/FProbe.bas`（不过 `ByRef Variant` helper ⇒ 没有 §11.22.1 那种失真）：

- **2/Integer**：F01~F04、F06、F06b、F09、F11、F12、F14、F20、G09、H01、H02（14 键）。
  改前一律 **3/Long** —— C3 把 `Integer + Integer` 升成了 Long。
- **错误 6**：F05（`30000+30000`）、F07（`200*200`）、F08（`30000*2`）、F19。
  改前这四枚静默算出 60000 / 40000 / 60000 这类数。
- **5/Double**：F10（`/`）、F13（`^`）。F13 改前**就是** Double ⇒ §11.22.2 那一行
  「`^` 返 Double 是缺口」**作废**（刀40 之前那次采集的 stale 行）。

#### 11.23.2 改法：同一个问句一个出口

1. `isNarrowIntExpr`（`cgen_util_ctrl.cpp`）加 `BinaryExpr` 支：`+ - * \ Mod` 且两侧窄 ⇒
   这枚节点在 VB6 里就是 Integer。`/`、`^` 刻意**不收**（真 VB6 那两枚答 Double）。
   这一枚同时喂三处：算术档、Variant 装箱档（F20）、赋值闸的「装得下不套」。
2. 发码的值那一半（`cgen_expr_binary.cpp`）：`+ - *` 两侧窄时显式抬 int32 算完再交
   `vb6_ChkInt` 收回 16 位；`\`、`Mod` 的两个操作数先过 `vb6_ChkInt`（越界的**输入**
   也要响），结果再包一层。int16 参与 C 运算会被整型提升成 int（结果悄悄变 Long），
   `(int16_t)` 硬截又会静默回卷 —— 两条都不是 VB6，所以中间量只能留 int32。
   步长逐格：`vb6_ChkInt((int32_t)(L) op (int32_t)(R))`，`L`/`R` 自己若是窄表达式则已包过
   ⇒ `2 + 3 * 4` 发成 `ChkInt(2 + ChkInt(3*4))`，每一格各响一次，与 VB6 从左到右一致。
3. `boxVariant86` 把窄操作数装成 `vb6_VariantInt((int16_t)(…))`（VT_I2）。改前走
   `vb6_VariantFromValue`，`_Generic` 把 int 装成 VT_I4 ⇒ `VarAdd` 一路升到 Long
   （F20 `arr(1) + 1` 因此以前答 3/Long）。
4. `narrowCheckAssign` 对「已经逐步套过 ChkInt」的窄算术**不再重套**：否则热循环里
   `i = i + 1` 每句两次 RTL 调用。判据问 AST（`BinaryExpr && isNarrowIntExpr`），不认文本。

#### 11.23.3 两处夹具订正（都是旧 C3 把 Integer 算术当 Long 才成立的口径）

- `tests/test_fe_slot_leak.bas` 的阈值 `8 * (500 * 40)`：`500*40 = 20000` 是 Integer，
  `8 * 20000` 在 VB6 是 Integer 乘法越界（常量表达式还是**编译期**错误 6）⇒ 改 `8&`。
  这一行以前只是靠 C3 升 Long 才算得出 160000；本刀一上就真跑 `Unhandled VB6 Error #6`
  （第一次跑就是这个堆栈），是夹具自己的问题、不是回归。
- 发码针 `f2lng_round`（`(4 * 1) + 3`）与 `b182_intformal_shape`（`fv + 1`）按新形状改。

#### 11.23.4 门与登记

- 门 **`44012205343`**（head `98557825`）：**21/22**，唯一红 = `Emit manifest (shape oracle)`
  —— 报「存量 35 份发码真的变了」。
- A/B 逐行归因（BASE = 父提交 `a09b5c78` 本机自编 `.temp/base46/.buildrel`，NEW = 本刀）：
  35 份里每份 removed == added（**无增删语句**，只有 1:1 换行），新增 **281** 行全部带
  `vb6_ChkInt` / `vb6_VariantInt`，撤下 147 行是它们被包之前的样子。抽样：
  `result = (2 + (3 * 4));` → `vb6_ChkInt((int32_t)(2) + (int32_t)(vb6_ChkInt((int32_t)(3) * (int32_t)(4))))`；
  `vb6_IntDiv(10, 3)` → `vb6_ChkInt(vb6_IntDiv(10, 3))`；`vb6_VariantFromValue(1)` → `vb6_VariantInt((int16_t)(1))`。
- 登记 `ada4c6b7`（`rebless_emit_manifest.ps1 -Bless`，吃 CI 那台自己交回来的清单）：
  748/748 行相同、`K1..K3` 集合相等。
- 验证轮 **`44013912679`**（head `ada4c6b7`）：**22/22 success**。
- 本机回归：smoke 1/1、syntax 227/227、asm 14/14、compile #1 88/88、compile #2 77/77、
  bas #1 181/181（含两条发码针）。

#### 11.23.5 剩余键（19 → 5）

`1.5!`（A12）、`Abs(Variant)`（C23）、D13/D15、一元 `-`/`Not`（E01~E10 + H06）。
`^`（F13）已从名单里划掉：实测本来是 Double。

#### 11.24 第四十七刀：`1.5!` 认 Single 档 + Single→String 的 7 位口径

真 VB6 读数来自 `.temp/litprobe/GProbe.bas` / `HProbe.bas`（用户在本机 VB6 跑，日志
`c3_vb6_probe_G.log` / `c3_vb6_probe_H.log`）。两个提交：

- `99f6e567` 类型面：`Token::singleSuffix` + lexer `case '!'` + parser 出
  `LiteralKind::Single`；`inferExprType` 认 Single→Single，`/` 走 Single 规则、`^` 恒 Double；
  Single 参与的四则按 `(float)((float)L op (float)R)` 落 float，`\` 的浮点实参走 `vb6_FltToLng`。
- `4f1e4621` 格式化面：`vb6_NumToBSTRDefault` 的 `isSingle` 支重写。

**Single→String 三条口径**（H01~H20 全量对过，不是「最短往返」）：

| 读数 | VB6 | 旧 C3 |
|---|---|---|
| `CStr(1/3)` | `0.3333333` | 同（碰巧） |
| `CStr(1234567.5)` | `1234568` | 同（碰巧） |
| `CStr(12345678!)` | `1.234568E+07` | `12345678` |
| `CStr(1e7!)` | `1E+07` | `10000000` |
| `CStr(16777216!)` | `1.677722E+07` | `16777216` |
| `CStr(0.00001!)` | `0.00001` | `1e-05` |
| `CStr(3.4E+38!)` | `3.4E+38` | `3.4e+38` |

即：① 有效数字恒 **7 位**；② 7 位舍入后 `|x| >= 1e7` 才转指数，低位一律定点；
③ 指数大写 `E` + 两位。`Str$` 另有两条口径：非负前一空格（`" 1.5"`）、`|x|<1` 去前导 0（`" .1"`）。

**Fix 117d 的启发式加门槛**：`CStr(12345678#)` 真读数是 `"12345678"`，而这个值恰好也是
float 精确值 —— 旧的无条件「float 精确 ⇒ 按 Single 打」会把它变成 `1.234568E+07`。
改成只对**非整数**float 精确值启用（噪声只来自 21.1000003814697 这一味）。

**坑（务必记住）**：`vb6rtl_conv.c` 在 `C3RTL_EMBEDDED_FILES` 里、`src/driver/c3rtl.rc` 也挂着
`OBJECT_DEPENDS`，但 **Ninja 下改 RTL 源不会重编 .rc**。必须显式 `touch src/driver/c3rtl.rc`，
否则「改完 → 重编 → 重跑」输出一字不变，会误判成「这条路径没人走」。

本机验证：HProbe 的 20 行 Single 与 VB6 逐行一致；SGProbe SG **18/19**（只剩 SG11）；
smoke 1/1、syntax 227/227、bas 181/181。

**推送经过（两个坑，记下来）**：
1. 三个提交一起推上去，**一个 CI 都没跑** —— tip 提交 `036c3baf` 的信息带 `[skip ci]`，GitHub 按它跳过**整次 push**。
2. 补一个空提交 `1b11aa4b` 敲门也不行 —— `ci.yml` 有 `paths` 白名单，空提交不改任何文件 ⇒ 不命中。
   正路是 `POST /actions/workflows/ci.yml/dispatches {ref:dev}`（= `gh workflow run ci.yml --ref dev`，
   workflow_dispatch 不受 paths 限制）。那轮 `44025035809` = **21/22**，唯一红就是发码形状门（意料之中）。
3. A/B 归因后才登记：worktree 在父提交 `aa821acf` 自编一台 `.temp/base47/.buildrel/C3.exe`，
   `tests/test_types.bas` 只 1 行变（`2 ^ 10` 变 `vb6_ChkLong(vb6_FltToLng(vb6_Pow(2, 10)))`）；
   `tests/Charts 2020/ucChartArea/Proyecto1.vbp` **85 增 85 减、1:1 换行无增删语句**，
   差异全是 Single 那一味（`(float)((float)L op (float)R)`、`\` 的实参过 `vb6_FltToLng`）。
4. 登记 `054b97e0`（`rebless_emit_manifest.ps1 -Bless`：改写 21 行，748/748 相同，K1..K3 集合相等），
   推送后门 **`44027063943`（head `054b97e0`）= 22/22 success**。

**剩余键（5 → 3 + 3 条新残）**：`Abs(Variant/Single)`（C23，含 SG11）、D13/D15、
一元 `-`/`Not`（E01~E10 + H06）；新残：Format$ 空格式那一支没透传 Single 标志
（H05/H06/H08 的 fmt 列）、`Str$` 浮点走 `vb6_Str(int32_t)` 截断（H01 打 `1`）、
Double 低位 `1e-5` 仍打 `1e-05`（只有一行读数，待补两点再定）。

#### 11.25 第四十八/四十九刀：Format$ 与 Str$ 的 Single 分出口

两刀都不动“怎么打”，只动“递进去的是什么”。前一刀（第四十七刀）已经把 7 位口径落在
`vb6_NumToBSTRDefault` 的 `isSingle` 支，可这两个内建函数从旧站拿到的是**错档的值**。

**第四十八刀（`ee83b83c`）— Format$ 首参**

`src/backend/detail/expr/cgen_expr_call_conv_output.inc` 原来是一臂：
`argType == Double || argType == Single` → `vb6_VariantDouble(...)`。Single 被装成 VT_R8，
于是 RTL 里 `vb6_NumToBSTRDefault(numVal, expr.vt == vb6_vtSingle)` 的 `isSingle`
**从这一站根本走不到**。真 VB6 探针 H05/H06/H08 的 `fmt` 列是
`1.234568E+07` / `1.677722E+07` / `1E+07`，C3 旧打 `12345678` / `16777216` / `10000000`。
改为 Double 一臂、Single 单独一臂并问权威 `boxToVariant`（同文件 Boolean 的 Fix 198、
Date 的账 #430 都是这个做法，零新形状）。

> 弯路记一笔：先猜“模块级 `As Single` 在过程入口被 `knownSingleVars_.clear()` 后没有回灌口”，
> 加了 `moduleSingleVars_`（声明处登记 + 三处过程入口回灌）。改完重编重跑 **一行未变**，
> 再做对照探针 `P48.bas`（局部 Single + 模块级 Single 各一枚）发现**两者都**发成
> `vb6_VariantDouble` —— 局部变量根本不经那条路，假设被判死。四处改动全部 `git checkout` 撤回，
> 才找到上面那一臂。**取证顺序：先造对照，再改假设。**

**第四十九刀（`461b9093`）— Str$ 实参**

`vb6_Str` 的形参一直是 `int32_t`，浮点实参在 C 侧就被截断（H01 `Str$(1.5!)` 打 `1`、
H10 `Str$(0.00001!)` 打 `0`、H14 `Str$(3.4E+38!)` 打 `-2147483648`）。RTL 新增
`vb6_StrSng(float)`：数字部分复用 `vb6_CStrSingle`（7 位口径）+ **非负数前一个空格**
+ **`|x|<1` 去前导 0**（`0.1` → ` .1`、`-0.00001` → `-.00001`）；后端在
`cgen_expr_call_builtin_fixup.inc` 里只把**静态类型确实是 Single** 的调用改 callee 名，
整数那一档一字未动。

读数（C3 跑同一份 HProbe）：**`cstr` / `str` / `fmt` 三列 × 20 行 Single 全部与 VB6 逐行一致**
（H01~H20）。

**两道静态普查门（两次红都不是产品缺陷，但都是该登记的东西没登记）**

| 门 | 报什么 | 处置 |
|---|---|---|
| `[STATIC] owned_temp_keeps` K13 | 「推导出来的 owned 1 枚发码侧那份里没有: `vb6_StrSng`」 | 登记进 `src/common/owned_temp_keeps.hpp` 的权威表（`96656bf8`） |
| `[STATIC] builtin_opt_defaults` R7 | 「名单过期/新增段: `vb6_Str`」 | 在 `$notPadding` 里补 `vb6_Str` 并写清依据（`461b9093`） |

R7 的规矩值得记住：**发码侧“按实参做决定（改 callee / 重排 / 截断）”的名字，必须恰好等于
那份带理由的名单** —— 多一个就红（新写了硬编码），少一个也红（名单过期）。这里不是补缺省档，
所以不进 opt-defaults 表，只写理由。

**另一个坑：提交信息不要内联**。`git commit -m "… Str$(Single) …"` 在 PowerShell 里会把
`$(Single)` 当子表达式展开，提交信息静默变成 `Str\ …`。本刀第一次提交就是这样坏掉的，
用 `--amend -F <文件> --force-with-lease` 修回（`c50a08d5` → `7ff613da`）。**此后一律 `-F 文件`**。

门：`44030372892`（head `ee83b83c`）= 22/22；`44038320685`（head `461b9093`）= 22/22。
本机：smoke 1/1、syntax 227/227、SG 18/19（只剩 SG11）。

**剩余键（3 + 2 条新残）**：`Abs(Single/Variant)`（C23，SG11）、D13/D15、一元 `-`/`Not`
（E01~E10 + H06）；新残：Double 的 `Str$` 仍走 `int32_t` 截断（H21 `12345678` / H22 `0`，
与 Double 低位 `1e-5` 打 `1e-05` 一起等下一批 Double 读数）。


### 11.26 第五十一刀（`Abs(Variant)`）：整刀撤回 —— 入口形状选错了

**读数（真 VB6 GProbe AB 段，Abs 之后问 VarType / TypeName / 值）**，这批读数**保留**，
下一刀直接拿来对账：

| 源 | VarType | TypeName | 值 |
|---|---|---|---|
| `CByte(5)` | 17 | Byte | 5 |
| `CInt(-5)` | 2 | Integer | 5 |
| `CLng(-70000)` | 3 | Long | 70000 |
| `CSng(-1.5)` | 4 | Single | 1.5 |
| `CDbl(-1.7)` | 5 | Double | 1.7 |
| `CCur(-2.5)` | 6 | Currency | 2.5 |
| `CBool(True)` | 2 | Integer | 1 |
| `CByte(200)` | 17 | Byte | 200 |
| `CInt(-32768)` | 3 | Long | 32768（Integer 装不下 ⇒ 升 Long） |
| `Empty` | 2 | Integer | 0 |
| `Null` | 1 | Null | — |
| `#1/2/2000#` | 7 | Date | 序列号绝对值 |
| `"5"` | 5 | Double | 5 |
| `"x"` | — | — | 错误 13 |

旧站只有 `vb6_Abs(double)` 一个出口 ⇒ 上面 13 行全塌成 VT_R8（实测 AB01~AB15 全是 5）。

**试过的做法**：RTL 加 `vb6_VARIANT vb6_AbsV(vb6_VARIANT v)`（按 vt 分档：Byte 留 Byte、
Integer 最小值升 Long、Currency 留 CY、Date 取序列号绝对值、BSTR 先 `vb6_IsNumericText`
否则 `vb6_RaiseError(13)`、Null 原样、Empty 答 Integer 0），后端 fixup 里静态类型是 Variant
时改名 `vb6_Abs` → `vb6_AbsV`，并把 Fix 108c 那一站先折出来的 `vb6_VariantToDouble(...)`
剥回原样（**只改 `args[0]` 不改 `argList` 不生效**，下游按 `argList` 发码）。

本地 AB 段 13 行全对，但门里连红三次，逐条归因：

1. `2aaa9104`：15/22 —— vbp #5~#8 跑报错、compile #1/#2 红。
   **根因：`Abs` 的结果有两副消费面** —— `VarType(Abs(v))` / `TypeName(...)` / `CStr(...)`
   要一枚**变体**，而 `Abs(v) * 2`、比较、赋 `Double` 要**标量**。fixup 那一站只看得到
   **实参**的静态类型，看不到结果往哪儿去；一律改名 = 把 `vb6_VARIANT` 当 double 算。
2. `01fd4ca5`（只把改名段 `#if 0`，RTL 与登记保留）：21/22 —— compile #1/#2 的
   `[STATIC] builtin_slot_returns` R2/R8：`cgen_variant_slot_prefixes.inc` 是**生成物**，
   RTL 多了一枚返回槽却没重跑 `scripts/gen_variant_slot_prefixes.ps1`。
3. `4b0989ca`（补跑生成脚本，表 45 名）：21/22 —— compile #1 的 `[STATIC] sa_access` A17
   （新 switch 里的 `case vb6_vtEmpty:` 让「TypeName 标量名单」数成 2 处）与
   `[STATIC] variant_i8_coverage` V2/D1/S1/S2（新 vt switch 把 I1/UI2/I8/DECIMAL 四档
   掉进 `default`；这张表要求**要么写 case、要么进带理由的豁免表**）。

**结论**：本刀的**入口形状**是错的，因此整刀撤回（`cf26609d`），树回到 `b4005273` 的发码形状
（`git diff b4005273 -- src` 只剩快照里那一行 `[own] vb6_StrSng`，那是第四十九刀的正当登记）。
下一刀要满足两条才动：
- **按消费侧分档**：变体位（`VarType`/`TypeName`/`CStr`/赋 Variant）改名走 `vb6_AbsV`，
  数值位保持 `vb6_Abs(double)` 出口 —— 或者干脆让 `vb6_AbsV` 另起一个名字、只在变体位发码；
- **新 vt switch 一次写全**：I1/UI2/I8/UI8/DECIMAL 要么写 case，要么按
  `check_variant_i8_coverage.ps1` 的豁免表写清理由（VB6 的 Variant 造不出 I8/DECIMAL，
  只有 COM 能交进来 —— 这是可以写进豁免表的正当理由）；`case vb6_vtEmpty:` 只许出现在
  TypeName 那份标量名单里，AbsV 那侧用前置 `if (v.vt == vb6_vtEmpty)` 收掉。

门：`2aaa9104` 15/22、`01fd4ca5` 21/22、`4b0989ca` 21/22、`cf26609d` **22/22**（全撤后回绿）。
本机：smoke 1/1、syntax 227/227、owned_temp_keeps(228) / builtin_slot_returns / sa_access /
variant_i8_coverage 全 PASS。

### 11.27 第五十二刀：Abs(Variant) 改在**装箱站**分档（C23 收口）

Abs(Variant) 改在**装箱站**分档（C23，AB 段 13 行对齐，pcline 不再红）

上一刀（第五十一刀）在调用站把 `vb6_Abs` 改名成 `vb6_AbsV`，本地 AB 探针 13 行全对，
门里却红 6 个作业。本机复现拿到实证（`.build\pre\tests\pcline` 的 PcDrawForm.c）：

    okCls   = (-(vb6_AbsV(vb6_GetCurrentY(vb6_hwnd_picP)) < 0.001));   ← cl C2440 + C2088
    okTwip  = (-(vb6_AbsV((dtw - vb6_ControlTextHeight(...))) < 0.001));
    okStack = (-(vb6_AbsV((stk - (2 * vb6_ControlTextHeight(...)))) < 0.001));

根因：`Abs` 的结果有**两副消费面** —— 变体位（`VarType`/`TypeName`/赋 Variant/`CStr`）
过装箱站；数值位（比较、算术、赋 Double）**不过**装箱站，直接吃 double 标量出口。
调用站看不到结果往哪儿去，一律改名 = 把一枚 `vb6_VARIANT` 当标量比。

改法（本刀）：把分档搬进**装箱站** `CCodeGen::boxToVariant`（src/backend/cgen_util_ctrl.cpp）：

  - 判据：表达式是 `IndexOrCallExpr`、被调名 Abs（大小写不敏感）、恰好 1 个位置实参、
    该实参静态类型是 Variant；
  - 实参侧 Fix 108c 已折出的 `vb6_VariantToDouble(...)` 一层剥回原样（子类型阶梯全靠 vt 说话），
    发 `vb6_AbsV(<原样 VARIANT>)`；
  - 数值位根本不经过这一站 ⇒ 照旧 `vb6_Abs(double)`，pcline 那三行恢复原形。

RTL `vb6_AbsV(vb6_VARIANT)`（src/rtl/core/vb6rtl/vb6rtl_conv.c）：按 vt 分档
Byte→Byte、Integer→Integer（-32768 升 Long）、Long→Long（溢出错误 6）、Single/Double 各自保档、
Currency 保 CY、Boolean→Integer、Date→取序列号绝对值、BSTR 先 `vb6_IsNumericText`
（否则 `vb6_RaiseError(13)`）、Null 原样、Empty 答 Integer 0；
**其余档（I1/UI2/UI4/I8/UI8/DECIMAL）不写保档臂**，落 `default` = `vb6_VariantToDouble`
取绝对值（权威换算，不是静默交 0）—— 与第四十五刀 `vb6_VarNegV` 同一取舍（无 oracle 不定口径）。

静态普查（这一刀一次对齐，避免上一刀那样一轮红一个）：
  - `check_variant_i8_coverage`：`vb6_AbsV` 进 I8 / Decimal / SByte·UInteger 三张**豁免**表，
    理由照 `vb6_VarNegV` 的写法逐条写清；
  - `check_sa_access` A17：`case vb6_vtEmpty:` 只许出现在 TypeName 那份名单 ⇒ 本刀用
    前置 `if (v.vt == vb6_vtEmpty)` 收掉；
  - `gen_variant_slot_prefixes.ps1`：RTL 多了一枚返回槽 ⇒ 派生表从 44 名补到 45 名；
  - `owned_temp_keeps`：`vb6_AbsV` 登记进 kOwned / kOwnedVariant / kRows（'V' / ret），
    快照一并 -Bless（229 行）。

验证（真 VB6 GProbe AB 段 ↔ C3 镜像探针 ABProbe_C3.bas 逐行同键）：
  AB01 17 Byte 5 / AB02 2 Integer 5 / AB03 3 Long 70000 / AB04 4 Single 1.5 /
  AB05 5 Double 1.7 / AB06 6 Currency 2.5 / AB08 2 Integer 1 / AB09 17 Byte 200 /
  AB10 3 Long 32768 / AB11 2 Integer 0 / AB12 1 Null / AB15 7 Date / AB14 5 Double 5，
  AB13（"x"）按 VB6 抛错误 13 终止（探针进程 exit 1）。
本机：`pcline\PcDraw.vbp` 编译 RC=0（上一刀在这里 C2440）；smoke 1/1；syntax 227/227；
      variant_i8_coverage / sa_access / builtin_slot_returns / owned_temp_keeps /
      builtin_opt_defaults 全 PASS。

门：ea766ea = **22/22**（manifest 未动，无需 rebless）。

### 11.28 第五十三/五十四刀：一元 `-`/`Not` 逐档子类型 + `Choose`/`Switch` 选中支原样

来源：用户跑的真 VB6 采集器 `ProbeVB6_U.bas`（47 行，日志 `.temp/litprobe/c3_vb6_probe_u.log`），
C3 侧镜像 `UProbe_C3.bas` 逐行同键。改前：U/N 段 7 行错（含 `-32768` 报错误 6、
`Not CByte(200)` 答 -201、`Not "5"` 把 BSTR 指针取反）、S 段 3 行压根编译不过（C2440）。

**一元负号**（`vb6_VarNegV` 逐档扩，src/rtl/core/vb6rtl/vb6rtl_com.c）与**一元 `Not`**
（新出口 `vb6_VarNotV`）的读数已逐行对完；`Not` 的三条口径要点：
  - Boolean ⇒ **逻辑非**（11/Boolean）；
  - Byte/Integer/Long ⇒ **位非且保档**（`Not CByte(200)` = 17/Byte 55，8 位取反）；
  - Single/Double/Currency/Date ⇒ 先按 VB6 **银行家舍入**取整再位非，答 Long
    （1.5→2、1.7→2、2.5→2 ⇒ -3）；
  - Empty ⇒ Integer -1、Null ⇒ Null、`"5"` ⇒ Long -6、`"x"` ⇒ 错误 13。

发码侧（src/backend/expr/cgen_expr.cpp）：串操作数不能直接取反（cl C2171）⇒ 先按串装箱
交给同一条阶梯；`- <Byte|Boolean>` → Integer；`Not <Byte>` → Byte；
`Not <Single|Double|Currency|Date>` → Long；`Not <Variant>` 从前一律
`(~vb6_VariantToLong(...))` 改走 `vb6_VarNotV`。

`-32768` 这一行的档在 `isNarrowIntExpr`：一元负号必须**先于常量折叠**判
（量字面 32768 自己装不进 Integer ⇒ 整个表达式是 Long）。

**Choose/Switch 走了两刀才落地，第二刀是纠偏：**
  第五十三刀一律「逐支装箱 + `vb6_VariantNull()` 兜底」⇒ S 段语义对，但把语料数值位写法全打红：
  CI 里 `VbQRCodegen`/`FlexGridX86` 双双 FAIL (compile)，本机复现
  `error C2440: 无法从 vb6_VARIANT 转换为 int64_t` / `error C2088: 「*」: 对 struct 非法`
  （`Choose(lIdx, 10, 12, 14)` 赋 Long、`Choose(...) * 8`）。
  第五十四刀改成**站点自己判**：各支同档 ⇒ 发旧标量三目（发码与改前逐字节相同）；
  异档 / 带 Null·Empty 字面量 ⇒ 逐支装箱。
  「同档」必须按 **VB6 的档**算：`1`(Integer) 与 `40000&`(Long) 在 inferExprType 里都是 Long，
  比较里要带上 `isNarrowIntExpr`（否则 S01 又答 Long）；Null/Empty 字面量的类型也是 Long，
  得单独问 LiteralKind（否则 `Choose(1, Null, 1)` 仍把 struct 与 `(1)` 摞进一个三目）。

**残留 3 行（已量到、没对，写在明面上）**：S07 `Choose(3,1,2)`、S08 `Choose(0,1,2)`、
S15 `Switch(False,1,False,2)` —— 真 VB6 是 **Null**，C3 现在答 Integer 0。
根因：这三行的候选支全是窄档 ⇒ 走标量三目，而标量三目给不出「越界交 Null」（两支同型）。
要对它们就得让**消费面知道这是个 Variant**（`inferExprType(choose/switch)` 答 Variant，
赋值/算术站出换算），那是下一刀。

harness 一处读数订正：`test_bitops` 的 `BF-byte-not` 从 `-86` 改为 **170**
（`Not <Byte>` 是 8 位取反：by1=85=0x55 ⇒ 0xAA=170，与 N03 `Not CByte(200)`=17/Byte 55 同口径）。

静态普查（一刀补齐）：`gen_variant_slot_prefixes` 补 `vb6_VarNotV(`（46 名）；
owned 表只登 kOwned / kOwnedVariant，**不进 kRows** —— 派生面说 `vb6_VarNotV|0` 的
`keeps_pointer=0`、`handoff-to-plain:vb6_VariantLong`（与 `vb6_VarNegV` 同族，登记了反而 K12 报多一行）；
`check_variant_i8_coverage` 三张豁免表（I8 / Decimal / SByte·UInteger）逐条写理由：
无 oracle ⇒ 落 `default` 走 `vb6_CLng(vb6_VariantToDouble)` 权威换算，不是静默交 0。

验证：U 段 47 行里 **44 行对上**（S 段 16 行对 13 行，见上面残留 3 行）；
语料 `VbQRCodegen-master\test\Project1.vbp`、`VBFlexGridDemo\VBFlexGridDemo.vbp`、
`pcline\PcDraw.vbp` 本机编译 RC=0；smoke 1/1；syntax 227/227；
owned_temp_keeps(230) / variant_i8_coverage / sa_access / builtin_slot_returns /
builtin_opt_defaults 全 PASS。
发码形状动了 7 份（一元 `-`/`Not` 的档位变化 + `Choose`/`Switch` 异档站点），
已按 `scripts/rebless_emit_manifest.ps1 -Bless` 从 CI 那份清单登记。

门：`94e6e7d2`（第五十三刀）红 4 个作业（manifest + bas#1 + vbp#4 + vbp#8）；
`faf5c28c`（第五十四刀）红 2 个（manifest + bas#1）；
`740af326`（登记 + 读数订正）= **22/22**。

### 11.29 第五十五刀（试探未落地）：Choose/Switch 全变体化能收下 S07/S08/S15，但消费站要一起改

试着把「同档走标量三目」那一半拆了：`Choose`/`Switch` **一律**发变体三目，同时
`inferExprType` 对它们答 Variant、`isDefinitelyVariantExpr` 的
多态内置 denylist 把 choose/switch 拿走（只留 `iif`/`if`）、`isNarrowIntExpr` 那两段删掉。

结果：
  - **探针 S 段 16/16 全对**（S07/S08/S15 都是 1/Null）—— 语义上就是这一条路；
  - FlexGrid 语料当场红两条（本机 RC=1，`error C2440`）：
      ① `MainForm.c(970/975)`：`vb6_VBFlexGrid_prop_let_Sort(..., <整只 VARIANT>)` ——
         属性 Let 的 `Long` 形参那一站**根本不换算**（标量隐式转换就行，struct 不行）；
      ② `VBFlexGrid.c(5471/5548)`：`hToolTip = <整只 VARIANT>` —— `hToolTip` 是 LongPtr，
         普通赋值站同样不换算。

所以这条路要连着改**消费面**：属性 Let 实参站、普通赋值站（至少 Long/LongPtr 两档）都得问
「源是不是 Variant」再出 `vb6_VariantToLong` / `vb6_VariantToLongPtr`；另外
`isDefinitelyVariantExpr` 里 §56.41 那条「限定形态 `VBA.Choose`/`VBA.Switch` 答否」也要一起
拿掉（发码侧现在两形都发变体，问句与发码必须同一口径，否则又回到 §B147 那一族：
oracle 说槽、`ChkLong` 当标量吃）。

这不是没量到，是**代价被量出来了**：为一个越界子项（S07/S08/S15）动全语料的数值消费面，
风险远大于收益 —— 所以本刀退回 `740af326` 的形状（同档标量 / 异档变体，S 段 13/16），
把这条路作为下一刀的入场条件记在这里。下次做的时候验收口径已定：
① 探针 S 段 16/16；② `VbQRCodegen` / `VBFlexGridDemo` / `pcline` 本机 RC=0；
③ bas 181 例 + syntax 227 + smoke 1/1 全绿，且发码形状变动逐一登记。

### 11.30 第五十五刀（落地）：Choose/Switch 一律变体三目 + 属性 Let 值站取数

按 §11.29 记的入场条件做完了。三处改动：

1. **发码**（`cgen_expr_call_builtin_pre.inc`）：删掉「同档走标量三目」那一半，
   `Choose`/`Switch` **两形都发整只 `vb6_VARIANT`** —— 逐支 `wrapVariantValue` 装箱、
   越界/无匹配交 `vb6_VariantNull()`。S07/S08/S15 那种「同一支里既要原样又要 Null」
   只有整只变体表达得出来。
2. **问句与发码同口径**：`isDefinitelyVariantExpr` 的 `isPolymorphicBuiltin` 只留 `iif`/`if`；
   账 §56.41「限定形态 `VBA.Choose`/`VBA.Switch` 答不是槽」那条豁免**撤销**；
   `isNarrowIntExpr` 删 choose/switch 两臂（子类型是运行期的，答窄档会把 struct 塞进
   `vb6_VariantInt((int16_t)(<整只 VARIANT>))`）。
3. **消费面补一站**（`src/backend/detail/stmt/cgen_assign_prop_write.inc`）：属性 Let/Set 写方向的
   **值形参**若落在数值档、而右值是整只 VARIANT，按形参档取数 —— 整型档 `vb6_VariantToLong`、
   浮点档 `vb6_VariantToDouble`，位置放在那条 ByRef 复合字面量取址**之前**（先取数、再收窄取址）。

定位过程留两条账，省下一次重犯：

  - 红点真身是**属性写站**，不是调用实参站。上一刀试过的两种「实参站兜底」都实测失败：
    放在 Fix 431 之后 → `vb6_VariantToLong(vb6_VariantToLong(x))` 双套；
    放在文件末尾 → 反向 `int32_t → vb6_VARIANT` 的 C2440（`Common.c(183/1632/1633)` 等）。
  - 表里那一发的名字是 `prop_let_` + 成员名，但**文件路径不是** `src/backend/assign/…`
    （猜错一轮），真实路径 `src/backend/detail/stmt/cgen_assign_prop_write.inc`，
    出口就是 `funcName + "(" + objExpr + ", " + valExpr + ");  /* Property Let */"` 那一行。

工具面有一条必须记住：**Ninja 不跟踪 `.inc`**。改完 `.inc` 要 touch 真正 `#include` 它的 `.cpp`；
这一回的宿主是 `src/backend/stmt/cgen_assign.cpp`（不是 `cgen_stmt.cpp`，touch 错了就是白跑一轮，
实测代价一整轮编译 + 一轮语料）。

验收（本轮按用户新规：本机只跑小量，全量交 GA）：
  - 探针 U 的 S 段 **16/16** 全对（S07/S08/S15 都是 1/Null）；
  - `VbQRCodegen` / `VBFlexGridDemo` / `pcline` 本机 **RC=0**（FG 修前是
    `MainForm.c(970/975)` 的两条 C2440：`vb6_VBFlexGrid_prop_let_Sort(..., <整只 VARIANT>)`）；
  - smoke 1/1 PASS；变体七份普查全 PASS（`variant_i8_coverage` / `variant_slot_store` /
    `prop_write_value_slot` / `variant_unbox` / `variant_cmp_boxing` /
    `variant_to_string_authority` / `manifest_coverage`）。

发码形状必定动了（两形都变变体三目 + 属性值站取数），`emit-manifest.expected.txt` 已按 CI 交回的
清单 `-Bless` 登记（改写 4 行）：`tests/test_builtin_arity.bas`、
`tests/VbQRCodegen-master/src/mdQRCodegen.bas`、`tests/VbQRCodegen-master/test/Project1.vbp`、
`tests/VBFlexGridDemo/VBFlexGridDemo.vbp`。登记后本机复算 748/748 逐行相同；
逐行看过形态：`Choose(1, "a")` 现在发
`vb6_VariantFromValue(((1)==1 ? (vb6_VariantString(vb6_BSTR_FromStr(L"a"))) : (vb6_VariantNull())))`
（`test_builtin_arity.c:109` 实测），与探针 S 段同一形状。

门（全量按用户新规走 GA，本机只留小量）：
  - 分支 PR 的 dispatch run `44091780671`（`big/variant-consumers`）= **21/22**，
    唯一红的就是 `Emit manifest (shape oracle)`（预期中的未登记）；bas #1/#2、syntax、smoke、
    asm、12 片 vbp 全绿 —— 也就是说“大改没伤别处”这一条是 GA 答的，不是本机答的；
  - 登记提交 `90f8dd5e` 推 dev 后 run `44094326716` = **22/22**（零红）。

### 11.31 第五十六刀（`762ad50e`）：Str$ 正数前导空格 + Double CStr 口径 + Str$(Double|Boolean) 出口

三项都按**真 VB6 第二轮读数**改（`run2\c3_vb6_probe_dbl2.log` 39 行 + 补测 `run2b`），不靠猜。
分支出 `dev`：`fix/c30-10-cut56a-str-dbl`。

**① `Str$` 正数前导空格（含 0）** —— `vb6rtl_string.c: vb6_Str`

| 读数 | 真 VB6 | 旧 C3 |
|---|---|---|
| `Str$(12345678&)` | `" 12345678"` | `"12345678"` |
| `Str$(0&)` | `" 0"` | `"0"` |
| `Str$(200B)` | `" 200"` | `"200"` |
| `Str$(-5&)` | `"-5"` | `"-5"`（负数不加，两边一致） |

`n >= 0 ? L" %d" : L"%d"`。这是 `Str$` 与 `CStr` 唯一的分歧点（`CStr(0)` = `"0"`、`Str$(0)` = `" 0"`）。

**② 新出口 `vb6_StrDbl` / `vb6_StrBool`**（`vb6rtl_builtin.h` + `owned_temp_keeps.hpp` + `cgen_*_fixup.inc`）

`Str$` 的数值口径 = CStr 的数字文本 + 一条**共用尾巴**（`vb6_StrFromNumText`：非负补空格、
`|x|<1` 去前导 `0`）。`vb6_StrSng` 一并改走这条尾巴，两处口径不再各写一遍。
`Str$(True)` = `" True"`（Boolean 与 Integer 同型 `int16_t`，不另开出口就被 `vb6_Str(int32_t)`
打成 `" -1"`）。发码侧按**静态档**铺分流表：

    Single → vb6_StrSng    Double/Currency → vb6_StrDbl
    Boolean → vb6_StrBool  其余整数档 → vb6_Str(int32_t) 旧路

**③ Double 的 CStr 口径**（`vb6rtl_conv.c: vb6_NumToBSTRDefault`）

真机读数定下四条，旧实现（`%.*g` 逐位搜 + `av>=1e15 || av<1e-4`）三条都不对：

1. **恒 15 位有效数字再剥尾零**。不用「最短往返」：A20 `CStr(4.94065645841247E-324)` 是 15 位，
   而最短往返在 1 位就满足（`5E-324`）—— 两者只在次正规数上分家。也正因为 15 位后剥零，
   `CStr(1234567.5#)` = `"1234567.5"`、`CStr(2.675#)` = `"2.675"`、`CStr(0.1# + 0.2#)` = `"0.3"`。
2. **定点区间 = 舍入后的十进制指数 `exp10 ∈ [-15, 14]`**：`1E-15` → `0.000000000000001`（定点）、
   `1E-16` → `1E-16`（指数）。旧门槛 `1e-4` 太低（真机到 -15 还是定点）。
3. **阈值必须判在舍入之后**：H01 `999999999999999.9` 本身 < 1E+15，先舍成 15 位有效数字
   才越界 ⇒ `"1E+15"`。旧路径判在舍入前，这一格漏判。
4. **指数 = 大写 `E` + 符号 + 至少两位**（`%+03d`）：`1E+07`、`1E-16`、`1E-300`。旧路径 `%.*g` 吐小写 `e`。

**删掉的那条启发式**（本刀唯一的减法）：

    if (!isSingle && v == (double)(float)v && v != floor(v)) isSingle = 1;   // 第四十七刀

它靠「值域」反推「这枚 Double 是不是 Single 提升来的」。D02 读数 `CStr(1234567.5#)` = `"1234567.5"`
被它命中打成 `"1234568"` —— 而 1234568 恰好也是合法的 7 位结果，**分不清**。噪声与真值在值域上
同形，只能靠发码侧的 provenance（「这一枚是 Single 提升来的」标记，账 §11.25 那一族）判，
不靠值反推。删掉后全量回归零红，说明旧启发式在语料里没有谁在依赖它的“纠偏”。

**编译器侧 39 格读数对照**（`.temp\cmp_probe.ps1`，同一份 .bas 真 VB6 与 C3 各跑一遍）：
33 格逐字对齐；余 6 格**全部是本刀范围外**，逐条归因：

| 键 | 真 VB6 | C3 | 归因 |
|---|---|---|---|
| `D02_cur_1p23456` | `1.2346` | `1.23456` | Currency 4 位半偶量化（`mapType`→double，Fix 126 的老账） |
| `D03_cur_5e-5` | `0` | `0.00005` | 同上（量化到 4 位后进位为 0） |
| `D08_var_cur_1p5` | `1.5`（`vt=6` Currency） | `1.5`（`vt=5` Double） | Variant 槽的 Currency 没保住档（同上一条的根） |
| `E01/E02/E04_sng_*` | `1E-05`/`1E-06`/`1E+07` | `0.00001`/`0.000001`/`10000000` | `ByVal Single` 形参档位丢失：进函数已是 Double，量不出 Single 的 7 位口径 |

**普查门（本刀又踩了一次 §11.25 那个坑）**

| 门 | 报什么 | 处置 |
|---|---|---|
| `[STATIC] owned_temp_keeps` K13d | 「推导出来的消费格 1 格发码侧那份里没有：`vb6_StrFromNumText|0`」 | 登记进 `kConsumes`（`0` = 无条件放掉）。**新写一枚 static RTL 助手、只要它自己 `SysFreeString` 形参，就会长出一格消费轴** —— 与它是否进用户 TU 无关，那列是对账用的 |
| `[STATIC] builtin_opt_defaults` R7 | 无（`callee == "vb6_Str"` 那一段还在原位，`$notPadding` 里已有理由） | 段内改成 switch 不改段头 ⇒ R7 看不见；理由文案已改成三档（Sng / Dbl / Bool）全写 |

`check_manifest_coverage` K1..K3 = 748/748 集合相等。**发码形状**：语料里 `Str$` 只有 2 处
（`tests/TeeChartChartDisplay/DisplayForm.frm:785/793` 的 `HScroll1/2.Value`，Integer 档），
整数路径一字未动 ⇒ 清单按理不变。本机未跑全语料 emit（20+ 分钟），这一条交门 ——
**门答了：`44144490088`（head `f2a47a9c`）= 22/22，`Emit manifest (shape oracle)` 绿**（清单逐行未变，
与「整数路径一字未动」的预测一致）。分支已推 `github`（新分支，`fix/c30-10-cut56a-str-dbl`）。
验证通道记一笔：`scripts/watch-gh-actions.ps1` 靠 `github` remote URL 里内嵌的 PAT，这个 checkout
的 URL 没嵌；`gh` 也没登录 ⇒ 实际走的是 `git credential fill`（Windows 凭据管理器里存的那枚）
取 token 再打 REST，不打印明文。

**本机验收**：`tests\run_tests.ps1 -Category all -Incremental` = **274 PASS / 0 FAIL**
（`.temp\run_tests_56a.log`）；smoke 1/1；`owned_temp_keeps` / `builtin_opt_defaults` /
`owned_temp_semantics` / `owned_temp_reclaim` / `manifest_coverage` 全 PASS。
（中途两个套件并行跑过 —— 我起的重跑与先前那个重叠，已掐掉旧的，只留带日志那轮为准。
教训：`Get-CimInstance` 只筛 `powershell.exe` 会漏掉 `pwsh.exe`，那一次“等待”就直接空转返回了。）

**剩余键（4 类，均不在本刀）**：Currency 的 4 位量化与 Variant 槽保档（同一根，第五十七刀）；
`ByVal Single` 形参档位（进函数的档位问题，不是打印问题）；`Str$(False)` 无真机读数（本轮只测了
True，按同构写 `" False"`）；Double 噪声需要发码侧 provenance（§11.25）。

### 11.32 第五十八刀：`Key`（改名）与 `HashVal` + DISPID 对齐真库（顺带修一条默认成员写的 AV）

留白普查（§7.15 of `ai/040`，109 个真库成员按名字对实现）只剩 Dictionary 两个名字没出现：
`Key`（属性写）与 `HashVal`（只读）。本刀补齐，并把 DISPID 对齐真库。

#### 一、读数先于实现（探针 `.temp/dict/`）

| 用例 | 真引擎（cscript + 真 scrrun） |
|---|---|
| 改名 `d.Key("a") = "b"` | err=0；旧键消失、新键在、值不变 |
| 旧键不存在（非空字典 / 空字典） | **32811**，字典原样不动（不顺手新建） |
| 新键已被别的键占用 | **457**，两边都不动，Count 不变 |
| 枚举顺序 | **位置不动**（`d.Add a,b,c` 后改 a→z 得 `z,b,c`） |
| 改名成同名（按 CompareMode 判等） | err=0，**空操作**（不报 457） |
| 默认（Binary）比较下改大小写 `Key("A")` | **32811**（找不到）；文本比较下可以 |
| 数值键 / 空串键 | 都能改（`1→2`、`""→"k"`），值跟着走 |
| `HashVal("a")` / `("abc")` / `("hello")` | Long：97 / 974 / 283；**不查存在性**（缺键、空字典也照算，err=0） |

`HashVal` 的字符串算法是**反推出来的**（`probe_hash.vbs` 13 条）：

```
h = 0;  对每个字符:  h = (h * 17 + c) mod 1201;   CompareMode != 0 时先取小写
```

`""`=0、`"a"`=97、`"A"`=65、`"aa"`=545、`"abc"`=974、`"abcd"`=1045、`"abcde"`=1052、
`"hello"`=283、`"zz"`=995 全部命中；文本比较下 `HashVal("a")` 与 `HashVal("A")` 都得 97
（按小写折）。

#### 二、实现（`src/rtl/core/vb6com/vb6com_dict.c`）

1. **DISPID 对齐真库 memid**：Item0 / Add1 / **Count2** / Exists3 / Items4 / **Key5** / Keys6 /
   Remove7 / RemoveAll8 / CompareMode9 / **HashVal10** / _NewEnum-4。以前是自编号
   （Exists2/Items3/Keys4/Remove5/RemoveAll6/CompareMode7/**Count8**）。改的理由不是洁癖：
   上一刀已经按真库注册了签名表（`src/com/builtin_scripting_sigs.inc`，带真 memid），
   DISPID 再自编就会出现「同一个成员两个号」，任何按 memid 走的路（外部宿主、缓存了
   DISPID 的早期绑定码）都会指到错成员。`vb6_ForEach_Init` 只要求 Count 的 dispid > 0，2 满足。
2. `Key`（属性写，dispid 5）：旧键查不到→32811；新键与旧键按 `dict_keyEq` 判等→空操作；
   新键被别人占用→457；否则**摘链→换键→重新入链**，槽位不动 ⇒ 枚举/Keys/Items 顺序不变、值不动。
3. `HashVal`（dispid 10）：上表算法，返回 VT_I4（真库签名写的是 Variant，实测 `TypeName` = Long）。

#### 三、对拍结果（`.temp/dict/run_key.ps1`、`run_reg.ps1`）

- 成员面 52 条并排比：**行为/字符串 38 条逐条一致**（K1–K8、H1–H7、M1/M2、O1/O2、U1/U2、
  S[\*]13 条、C1–C3）。
- DISPID 重排回归 21 条：18 条一致；3 条差异全部与 Dispid 无关，逐条交代见「四」。
- `C3_FSO_PREFER_MS` **只管 FSO**（`vb6com_fso.c:2066`），对 Dictionary 无效 ⇒ Dictionary 只有
  内建一条实现腿，对照轴是「cscript+真 scrrun」对「内建」。这点以前在 FSO 刀的脚本里容易被误读成两腿对照。

#### 四、没对齐 / 没修的，逐条记账

1. **非字符串键的 `HashVal` 未对齐**（本刀只给 0）。真引擎数值键另有一套：`0→0`、`1→161`、`2→985`、
   `3→196`、`-1→930`、`1.5→573`、`65536→134`。穷举 VARIANT 布局（vt 在前/值在偏移 4 或 8）×
   宽度 4/8 × 字节序 × seed 候选全不命中（`.temp/dict/derive_num.ps1`），**不猜**，直接记缺口。
   影响面：`HashVal` 是未文档化成员，语料零引用，只有「主动拿数值键调 HashVal」才会看到差别。
2. **`Join` 遇到数值元素抛 13**（旧缺口，**第五十九刀已修**，见 §11.33）：`Join(Array(1,2,3), "|")`
   真引擎给 `1|2|3`，我们报 13（`.temp/dict/JoinNum.bas`）。不是 Dictionary 的事（`Keys`/`Items` 返回的就是
   Variant 数组，字符串元素能 Join）。
3. **链式默认成员取值** `d("outer")("inner")` 两条腿都得不到值：真引擎 7，我们抛 5
   （`.temp/dict/NestProbe.bas`）。生成的 C 里那条赋值被整体折叠成 `v`（`--emit-c` 可看），
   与本次改的入口无关，属既有缺口。

#### 五、顺带修掉的一条既有 AV（`d("k") = v`）

回归探针在 `d("k") = "v"` 处 `0xC0000005`。取证三步：

1. **不是我改出来的**：把本刀的 `vb6com_dict.c` stash 掉重编，同一探针仍 AV（退出码相同）；
2. `C3_COM_TRACE=1` 的最后一行为 `getDispid("(null)") … dispid=-1`，随后 AV `read target=0x0`；
3. `--emit-c` 的生成码是 `vb6_ComSetPropArg(d, 0, …)` —— 参数化**默认成员**赋值把成员名发成 NULL
   （`cgen_util_comwrite.cpp` 的 `defMemLit25`）。

于是链条是：`vb6_getDispid(pDisp, NULL)` → 缓存遍历 `wcscmp(…, NULL)`（或错误分支
`fwprintf("%ls", NULL)`）→ 读 0x0。修法（`vb6com_invoke.c`）：

- `vb6_getDispid` 开头 `if (!name) return DISPID_UNKNOWN;`（防 NULL 进 `wcscmp`）；
- `vb6_ComSetPropArgRcv` / `vb6_ComSetRefRcv`：成员名为 NULL/空 **一律按 DISPID_VALUE(0)** 发送，
  不按名字查（默认成员在各库里叫法不同，真语义就是 0）；日志对 NULL 做兜底。
- `SetRef` 的宿主/OCX 两条分支仍要名字，用 `L"Item"` 当默认成员名（与取值腿
  `vb6_ComCallObjectDefaultRcv` 的旧口径一致）。

修后：`S2 put ok count=1` / `S3 get=v`，exit=0；回归探针 R10/R10b 与真引擎一致。

#### 六、本刀验收

编译绿；`.temp/dict/run_key.ps1` 行为/字符串差异 0；`run_reg.ps1` 只剩上述 3 条已记账差异；
门（GA）见提交信息。

### 11.33 第五十九刀：`Join` 的元素按 CStr 口径转（用户拿 VB6 立即窗口点的那条）

§11.32 记的「`Join` 遇数值元素抛 13」被用户当场验证了：VB6 立即窗口里
`?join(array(1,2,3),"|")` 与 `?vba.join(array(1,2,3),"|")` 都答 `1|2|3`。旧夹具
`tests/test_joinfilter.bas` 里那条 `JF-err=13` 是把**错的行为钉成了期望值**，本刀改回。

#### 一、读数（`.temp/join/oracle_join.vbs` 23 条 + 类型化数组 4 条 + 非数组/Null 5 条）

| 元素类型 | 真引擎 |
|---|---|
| Long / Integer / Byte | `1\|2\|3` / `-1\|-2` / `65\|66` |
| Single / Double | `1.5\|2.25` / `1.5\|1E+20`（`0.000001\|0.1` 也对） |
| Currency | `1.2345\|-2` |
| Boolean | `True\|False` |
| Empty | 空串（`Array(Empty,1)` → `\|1`） |
| **Null 元素** | **13** |
| Date | 本地化文本（`2020/1/2 3:04:05`） |
| 混排 | `a-1-True-` |
| **对象元素** | **450** |
| **数组元素** | **13** |
| 类型化数组 `Long()/Integer()/Double()/String()` | 一律转文本，不报错 |
| **非数组**（`Join(5)` / `Join("abc")` / `Join(Empty)`） | **13** |
| **二维数组** | **13** |
| 分隔符 Empty / 只给一个参数 | 拼接 / 默认空格 |
| **分隔符 Null** / **整体 Null** | **94** |

#### 二、实现（`src/rtl/core/vb6rtl/vb6rtl_compat.c`）

1. 新增 `vb6_SA_ReadJoinStrAt()`：**Join 专用**的元素读取。字符串数组直取；Variant 数组按 vt 分派
   （Empty→空、BSTR→原值、Null→13、数组→13、对象→450、其余→交给 `vb6_CStr`）；
   类型化数组按 elemType 造等价 Variant 再走 `vb6_CStr`（打印口径与 CStr 单一主人）。
   **故意不合并**进 `vb6_SA_ReadStrAt`（那是 `Filter` 的旧口径，只吃字符串/Empty）。
2. `vb6_Join()`：`!arr || cDims != 1` → 13（旧写法落入 `count<=0` 分支静默返回空串，等于把错当对）；
   元素拼接改成可增长缓冲区，现造的元素串用完即释放。

#### 三、对拍与夹具

- 23 条并排：**J01–J18 全部逐字一致**（数值/布尔/货币/字节/Empty/混排/长整型/P 维）；
  错误面 J19/J20/J22（非数组/二维/数组元素 → 13）也一致。
- `tests/test_joinfilter.bas` 补了 `JF-num/JF-mix/JF-byte/JF-cy/JF-null-err/JF-nonarr-err/JF-2d-err`，
  并用等价 VBS 逐条复核过真机值；`tests/run_tests.ps1` 的期望表同步（删 `JF-err=13`）。

#### 四、本刀发现但**未修**的（均与本刀改动无关，逐条留证）

1. **分隔符为 Null 时的 94 丢了**（J14：真引擎 94，我们 err=0 并给 `12`）。根因不在 Join：
   发码把实参转成 `vb6_VariantToString(vb6_VariantNull())`，那条转换站对 Null 静默给空串。
   真 VB6 的「Null → String」是 94；要改就得改强制转换站，影响面比 Join 大，另开一刀。
2. **`Array(<对象>)` 把对象存成 Long**（V1：`VarType=3`、`TypeName=Long`、`IsObject=False`，���出的“值”
   是指针数）。这才是 J21（`Join(Array(obj))`）的根因；Join 自己的 450 分支是通的 ——
   用 `Join(d.Items, "|")`（值为对象）实测得到 450。
3. **`Filter` 遇数值元素仍抛 13**（J23：真引擎 err=0）。真机行为是「按文本匹配、结果**保留原元素类型**」
   （`Filter(Array(1,2,21),"1",True)` → 2 个元素、`TypeName` 还是 Integer）；我们的结果数组是 BSTR 型，
   要修必须改结果数组的元素型，属结构性改动。
4. **数字做分隔符 `Join(arr, 0)` 编译不过**（C2440: double→BSTR），真引擎给 `102`。
5. 附带发现：`Array(CDate("..."))` / `Array(CByte(..))` 裸写会撞元素 setter 类型推断
   （生成 `vb6_ArraySetBSTR(..., vb6_CDate(...))` → C2440）。本刀探针一律先落 Variant 变量绕开，
   这条单独记在这里。

### 11.34 第六十刀：`Array()` 的元素打包（对象 / Byte / Date / Currency / Single）

§11.33 记的第 2 条（`Array(<对象>)` 被存成 Long，`TypeName=Long`）和第 5 条
（`Array(CDate/CByte)` C2440）是**同一处发码**：`cgen_expr_call_builtin_pre.inc` 里的
`Array()` 元素 setter 选择链是「认 Long/Double 字面量 + 猜 BSTR + 其余一律 Long」。

#### 一、读数（`.temp/arr/oracle_arr2.vbs` + `.temp/arr/ArrElem.bas`）

| 实参 | 真值 | 改动前发码 |
|---|---|---|
| `CByte(65)` | **17** | `ArraySetLong` → 3 |
| `CDate("2020-01-02")` | **7** | `ArraySetBSTR` → **C2440 编不过**（实参里含 `L"..."`，被 `looksLikeBSTR` 猜中） |
| `CCur(1.2345)` | **6** | `ArraySetLong` → 3 |
| `CSng(1.5)` | **4** | `ArraySetLong` → 3 |
| `CLng(5)` / `CInt(7)` | 3 / 2 | 3 / 2（本来就对） |
| `True` / `Empty` / `Null` | 11 / 0 / 1 | 已对（第 348 刀） |
| 混排 `Array(1,"x",True,Empty,Null)` | `2,8,11,0,1` | 已对 |
| **`Array(<对象>)`** | **9，`IsObject=True`，`TypeName=Dictionary`** | `ArraySetLong(d)` → 3，`TypeName=Long`（指针当数） |
| **`Array(Nothing)`** | **9，`IsObject=True`** | `vb6_VariantFromValue(NULL)` → 非对象档 |

对象那两格**不能拿 cscript 当真值**：VBScript 把对象实参自动取默认属性（`Array(d)` 装进去的是
`d` 的默认成员值，不是引用）。真值口径来自 VB6 语义（`Array` 存引用），并与既有实测
`Join(d.Items)`（值为对象）→ **450** 互洽。

#### 二、实现

1. **选择链**（同文件，仍在 `variantSetter` 这条唯一出口附近）：新增
   - `Nothing` 字面量：改发 `vb6_ArraySetObject(arr, i, NULL)`（原来那句注释假设 `_Generic` 会落
     `vb6_VariantObject`，实测没落）；
   - `knownObjectVars_` 里的标识符 或 `inferExprType` 答 Object/DataObject → 对象出口；
   - `inferExprType` 答 Byte/Date/Currency/Single → `vb6_VariantByte/VariantDate/VariantCur/VariantSingle`。
   **`knownObjectVars_` 是必须的**：`inferExprType` 对局部 `As Object` 答不出来（它只看 `symTab_`，
   局部变量在那里不可达）。
2. **新增 RTL 出口** `vb6_ArraySetObject()`（`vb6rtl_date.c` + `vb6rtl_array.h`）：清槽后
   `vt=VT_DISPATCH` + **一次** `vb6_ComAddRefDispatch`。
   没有发成 `vb6_ArraySetVariant(arr, i, vb6_VariantObject(p))`：那条是 `VariantCopy` 且**不清源**，
   临时值没人 `VariantClear` ⇒ 每个对象元素泄漏一份引用。
3. 副作用（正向）：`Join(Array(obj))` 现在真的走到 **450** —— 改动前对象已被压成 Long，
   450 那条分支根本不可达（§11.33 因此把它记成「Join 的缺口」，实则不是）。

#### 三、验收

- `.temp/arr/ArrElem.bas` 19 条：A1/A2/A3 对象档 9/True、A4 17、A5 7、A6 6、A7 `2,8,11,0,1`、
  A9 `Join(Array(obj))`=450、A10 4、A11 3、A15/16/17 11/0/1 —— 与真值逐条一致。
- `Array(CreateObject("..."))` → 9/True；`Array(o, 1)` → `9,2`（另一条表达式形态，也通）。
- 夹具 `tests/jsonfast_611/ArrayLiteralVariants.bas` 补 `AL-obj/AL-obj-is/AL-nothing-is/AL-byte/
  AL-cy/AL-sng/AL-date2` 七条，`run_tests.ps1` 期望表同步；本机 19 行输出与真值逐字一致。
- 哨兵 `scripts/check_array_literal_variants.ps1`：R4 的 `vb6_VariantDate(` 计数 1→2（第二个是
  `Array(<Date 值>)` 的真出口），新增 **R11** 钉对象/Byte/Cur/Single 四个出口的出现次数，
  R9/R10 补新针。PASS。

#### 四、仍遗留

- §11.33 的 #1（Null→String 该给 94 的强制转换站）、#3（`Filter` 遇到数值元素仍抛 13）、
  #4（`Join(arr, 0)` 数字分隔符 C2440）**未动**。
- 新发现：`CVar("x")` 链接失败（`LNK2019: vb6_CVar`，内建整族里没有 CVar）。

### 11.35 第六十一刀：`Filter` 的元素档（结果数组逐元素保子类型）

§11.33 记的第 3 条：`Filter` 遇数值元素抛 13、结果数组恒为 BSTR 档。真引擎不是这样：
**命中按文本比，结果逐元素保留自己的档**。

#### 一、读数（`.temp/filt/oracle_filt.vbs`，15 条）

| 条件 | 真引擎 |
|---|---|
| `Filter(Array(1,2,21),"1",True)` | n=2, LBound 0, **t=2(Integer)**, `1|21` |
| 同上 `include=False` | n=1, t=2, `2` |
| `Filter(Array(1.5,2.25,11.5),"1.5",True)` | n=2, **t=5(Double)**, `1.5|11.5` |
| `Filter(Array("ab","cd","ae"),"a",True)` | n=2, t=8, `ab|ae` |
| `Filter(Array(True,False),"True",True)` | n=1, **t=11(Boolean)** |
| `Filter(Array(1,"x1",True),"1",True)` | n=2, **t=2,8（各元素各自不变）**, `1|x1` |
| 零命中 | LBound 0 / UBound **-1** |
| 空匹配串 `""` | 全命中 |
| 文本/二进制比较 | 1 / 0 命中 |
| `Filter(5,"x")` | **13** |
| `Filter(Array(Empty,1),"",True)` | n=2, **t=0(Empty)** |
| `Filter(Array(Null,1),"1",True)` | n=1, **err=0**（Null 元素不报错、不命中） |
| `Filter(Null,"x")` | **94** |

#### 二、实现（`src/rtl/core/vb6rtl/vb6rtl_compat.c`）

1. 新增 `vb6_SA_ReadFilterText()`：元素→文本，多一个 `hasText` 出参 —— `Null` / 数组元素 /
   对象元素 `hasText=0`（不报错，只是不参与命中），其余按 CStr 口径转。**没有**复用
   `vb6_SA_ReadJoinStrAt`：Join 对 Null 元素**抛 13**（已实测），两边口径不同。
2. 新增 `vb6_SA_ReadVariantAt()`：命中元素读成一枚 Variant（类型化数组按自己档装）。
3. `vb6_Filter()` 重写：
   - 结果数组档：源是 `vb6_sa_bstr` ⇒ 仍是 BSTR 数组（既有口径，零回归）；其余 ⇒ **Variant 数组**，
     逐元素 `vb6_VariantCopy`（这条就是 F1/F3/F5/F6/F12 的根）。
   - 两遗的谓词统一：`hasText && match` 才算，空匹配串全命中，`owned` 的临时串两遗都释。
4. 顺带堵一个既有崩溃：发码对非数组实参**不一定做转换**，`Filter(5,"x")` 发成
   `vb6_Filter(5, ...)`（整数当指针）⇒ 旧写法 `source->lBound` 直接 0xC0000005（实测改动前
   F11 之后进程退出码 1）。现在 `vb6_FilterSourceUsable()` 用 SEH 探一下再解引用，野指针兜成 13。

#### 三、验收

- 15 条并排：**14 条逐字一致**（F1–F13、F15）；F11 由崩变 13（与真引擎一致）。
- 夹具 `tests/test_joinfilter.bas` 补 `JF-fnum`/`JF-fmix`/`JF-fempty`/`JF-fnull`，
  `run_tests.ps1` 期望表同步；本机输出与真引擎逐字一致。
- 零回归：`test_builtin_arity` 的 `AR-filter=[abc]`（字符串源，走 BSTR 结果档）不变。
- 门（GA）见提交信息。

#### 四、遗留

1. **`Filter(Null, …)` 真引擎给 94，我们 13**：Null 在 `vb6_VariantToSafeArray1D` 里和普通
   非数组一样化成 NULL，两者在本函数里分不开。要修得动「Variant → 数组」那一站，
   与 `Join(Null)` 该给 94（§11.33 #1 的第二半）同一格。
2. **对象元素 / 数组元素参与 `Filter` 的判定未测**：cscript 传对象会取默认属性，拿不到真值。
3. 类型化数组作源（如 `Dim l() As Long`）本机不再报错、结果给 Variant 数组，但真 VB6 口径未测。

### 11.36 第六十二刀：Null 转换站（Variant(Null) → String 一律 94）

§11.33 #1 与 §11.35 遗留 1 指向同一处：Variant 里的 Null 在被当成**字符串**用时，
真引擎报 **94**，我们静默给空串（J14 的 `Join(arr, Null)` 就是这一格的影子）。

#### 一、读数（`.temp/null/oracle_null.vbs` 27 条 + `.temp/null/NullProbe.bas`）

| 条件 | 真引擎 | 我们(改前) | 我们(改后) |
|---|---|---|---|
| `CStr(Null)` | 94 | 0 | 0（本刀未动，见遗留 0） |
| `Asc(Null)` / `Chr(Null)` | 94 / 94 | 0 / 0 | **94** / 0 |
| `CInt/CLng/CDbl(Null)` | 94 | 0 | 0 |
| `Join(arr, Null)`（分隔符） | **94** | 0 | **94** |
| `Join(Null, "\|")`（源） | 94 | 13 | 13 |
| `Split(Null, ",")` | 94 | 0 | **94** |
| `Replace(Null, …)` | 94 | 0 | **94** |
| `UBound(Null)` | 13 | 9 | 9 |
| `Len/Trim/UCase/Mid/Left/LCase(Null)` | cscript 0 | 0 | **94** |
| `Val(Null)` | cscript 13 | 0 | **94** |
| `Null & "x"` / `"x" & Null` | VB6：结果是 **Null**（cscript 是宿主口径，不当真值） | `"x"` | `"x"` |
| `Null = ""` | Null | False | False |
| `Null + 1` | Null | Null | Null ✓ |
| `IsNull(Null)` | True | True ✓ | True ✓ |
| `CStr(Empty)` | `""` | `""` ✓ | `""` ✓ |
| `Dim s As String: s = Null` | 94 | 0 | 0 |

口径说明：VB6 的通用规则是「String 形参 / String 目标收 Null ⇒ 94」。cscript 对
`Len/Trim/UCase/Left/Mid` 宽松成 0、`Val` 给 13 属 VBScript 宿主差异——这几格按 VB6 规则
统一取 94，并在探针文件头注明“不当真值”。

#### 二、实现

1. **RTL**：新增 `vb6_CStrStrict(vb6_VARIANT)`（`vb6rtl_conv.c` + `vb6rtl_builtin.h`）：
   `vt == Null` ⇒ 抛 94（`Invalid use of Null`），否则委托 `vb6_CStr`。
   特意**不改** `vb6_CStr`：`&` 与 `Debug.Print` 走的是它，两者口径另算。
2. **五个站换成严格出口**：
   - `cgen_expr_call_arg_variant.inc`：“Variant 实参 → String 形参”总口（并把它加进
     “已是 BSTR 不重复包装”的前缀表）—— 这一格就是 J14；
   - `cgen_expr_binary_util.cpp` 的 `wrapToBSTR()` 五个分支；
   - `cgen_util_classtype.cpp` 的 `toStringIfVariant()` 三个分支；
   - `cgen_assign_value_sem.inc` 六处赋值折算。

   > `CStr` 本体**不换**：`vb6_CStr` 在 RTL 里是一组按实参类型分诊的重载宏，把内建表里的
   > 名字换成只收 `vb6_VARIANT` 的 `vb6_CStrStrict`，`CStr(<bool 表达式>)` / `CStr(5)` 这类
   > 标量实参就落成 C2440（int → vb6_VARIANT）。这一点是门 `44311607151` 拿 4 个作业
   > （test_datelit / test_date_display / b181_baformal / test_redim_elem / b591_radixtier）
   > 换来的：改名后本地复现 `dl.c(12): error C2440`，回退即全绿。

   > `Like` 这一站（`toStringIfVariant`，只有 `Like` 用）**不纳入**严格口径：VB6 的
   > `Null Like "a"` 结果是 Null（不是 94），而本刀读数没量过 Like —— 没读数的格不顺手收紧。

#### 三、验收

- 探针 `NullProbe.bas` 29 条：String 面全部到位（上表加粗格）。
- 夹具 `tests/test_joinfilter.bas` 补 `JF-nullcstr=94 / JF-nullasc=94 / JF-nulldelim=94 /
  JF-nullsplit=94 / JF-nullrep=94`，并用 `JF-concat-ok=xy/0`、`JF-isnull=True/0` 钉反向（不许误伤
  宽松口径）；本机输出与真值一致。
- 零回归抽检：`test_builtin_arity` 的 `AR-join1` / `AR-filter` 不变。
- 门（GA）见提交信息。

#### 四、遗留

0. **`CStr(Null)` 仍 0**：见上“不换”一段。要动它得先把 `CStr` 的重载宏按实参类型分诊出
   “Variant 则严格”这一支，不属“换一个名字”的量级；单独一格。
1. **`Dim s As String: s = Null` 仍给 0**：这条不走上面五个站（有专用支路），编译期只给
   VB3003 警告。改它要定点到那条支路，另开一格。
2. **数值站**：`CInt/CLng/CDbl(Null)`、`Chr(Null)`、`Mid(s, Null)` 仍 0 —— 属“Null → 数值”那一站，
   影响面比 String 站大（比较/算术都用它），单独一刀量。
3. **`&` 与比较符的 Null 传播**：VB6 是 `Null & "x"` → Null、`Null = ""` → Null；
   我们现在分别给 `"x"` 与 False（VBScript 口径）。要改得动运算符发码，不能靠转换站。
4. **`Join(Null, …)` 源**仍 13（真 94）、`UBound(Null)` 仍 9（真 13）：这两格归
   “Variant → 数组”那一站，与 §11.35 遗留 1 同一处。
5. 编译器当前对 `Len(Null)` / `Join(Null, …)` / `UBound(Null)` 这类**字面量**直接报编译期错
   （VB4001 / VB3043），所以读数一律用装着 Null 的 Variant 变量取。

### 11.37 第六十三刀：数值/位置站（Variant(Null) → 数值 = 94）

§11.36 遗留 2 那一格。读数 `.temp/null/oracle_num.vbs`（40 条）+ `.temp/null/NumNull.bas`。

#### 一、读数

| 条件 | 真值 | 我们(改前) | 我们(改后) |
|---|---|---|---|
| `CInt/CLng/CDbl/CSng/CByte/CBool/CDate(Null)` | 94 | 0 | **94** |
| `CCur(Null)` | 94 | 0 | 0（遗留 1） |
| `Chr(Null)` / `Mid("abc", Null)` / `Mid(s,1,Null)` | 94 | 0 | **94** |
| `Left/Right(s, Null)` / `Space(Null)` / `String(Null, "a")` | 94 | 0 | **94** |
| `Abs/Int/Fix(Null)` | **Null** | 0 | 0（遗留 2） |
| `Hex/Oct(Null)` | **Null** | 0 | 94（遗留 3：变成报错了） |
| `Len/Trim/InStr(Null)` | cscript 给 Null，VB6 按 String 形参给 94 | 94（上一刀） | 94 ✓ |
| `1 + Null` / `Null * 2` / `-Null` | Null | Null ✓ | Null ✓ |
| `CInt(Empty)` / `CInt(12)` | 0 / 12 | 0 / 12 ✓ | 0 / 12 ✓ |

#### 二、实现（三枚闸）

1. `vb6_VariantConvNum`（`vb6rtl_builtin.h`）：`vt == vb6_vtNull` ⇒ `ErrRaise(94)`。
   `CInt/CLng/CDbl/CSng/CByte/CSByte/CUInt/CULng/CULngLng/CLngLng/CLngPtr/CCur` 的 `*V` 变体
   全走它 ⇒ 一格改完一整族。同文件补 `vb6_ErrRaise` 原型：不补的话 C11 先生成隐式
   `int(...)`，随后 `vb6rtl_array.h` 的 `void(...)` 就成了 C2371 冲突定义。
2. `vb6_CBoolV`：它自己那两条早退路（Boolean/BSTR）不走 ConvNum，所以在函数头另加同类闸。
3. `vb6_VariantToLong`（`vb6rtl.c`）：`case vb6_vtNull` ⇒ 94。这一枚是「Variant → 整数」的
   形参站/长度站共用的门：`Chr` 的 code、`Mid/Left/Right` 的 start/len、`Space/String` 的
   计数全经它，于是 M9–M15 一格全开。
4. `vb6_CDate`（`vb6rtl_conv.c`）：原本 `vtNull` 和 `vtEmpty` 一起返 0，现在只放行 Empty。

#### 三、验收

- 探针 25 条：上表加粗格到位；反向格（`CInt(Empty)=0`、`1 + Null = Null`）未被带跑。
- 夹具 `tests/test_joinfilter.bas` 补 `JF-nullcint/clng/chr/mid=94`，并用 `JF-cintempty=0/0`、
  `JF-nullplus=True/0` 钉反向；本机输出与针全一致。
- 零回归抽检（编译面）：test_joinfilter / test_builtin_arity / test_datelit / test_date_display /
  b181_baformal / test_redim_elem / ArrayLiteralVariants 全过。
- 门（GA）见提交信息。

#### 四、遗留

1. ~~**`CCur(Null)` 仍 0**~~ —— **已修（本刀续）**。根因是 `CCur` 不在 `cgen_expr_call_conv_numeric.inc`
   那道 V 后缀门的名单里（实参早已被上游包成 `vb6_VariantToDouble(...)`，所以 `firstArgIsVariant`
   是 false）。处置：名单加上 `vb6_CCur`，并把那里“剥回 double 包装”的检测（原先只服务 CBool）
   扩到 CCur。现发码 `vb6_VariantCur(vb6_CCurV(nv))`，读数 M5=94；非 Null 路未动
   （`CCur(2.5)=2.5`、`CCur("12")=12`、`CCur(1)/4=0.25` 实测）。
2. **`Abs/Int/Fix(Null)` 仍 0（真值 Null）**：它们收的是 `double`，Null 在 `vb6_VariantToDouble`
   就已丢了；要回 Null 得在“Variant 返回侧”做传播站（整个表达式要是 Variant），与 §11.36 遗留 3
   （`&`/比较符传播）同一族。
3. **`Hex/Oct(Null)` 从 0 变成 94**：真值是 Null。根因同遗留 2 —— 它们与长度站共用
   `vb6_VariantToLong`。这一格是**刻意接受的退步**（0 与 94 都错，94 至少不再静默），
   等传播站做完再改回 Null。

### 11.38 第六十四刀：Abs/Int/Fix/Hex/Oct 的 Null 传播

§11.37 遗留 2/3 那一族。真值是「Null 进、Null 出」（读数 `.temp/null/oracle_num.vbs`）。

#### 一、根因（两条，都使第五十二刀既有的 `vb6_AbsV` 白写了）

1. `boxToVariant` 里那条 Abs 分支多问了一句 `inferExprType(实参)==Variant`；实测它对
   `Dim nv As Variant` 的局部答不出 Variant ⇒ 分支不进，发码仍是
   `vb6_VariantDouble(vb6_Abs(vb6_VariantToDouble(nv)))`。
2. 更早的一道门在 `wrapVariantValue`：它自己的 `switch (inferExprType)` 对 Int/Fix 先答
   `Double`、对 Hex/Oct 先答 `String` 并**直接 return**，于是根本轮不到 `boxToVariant`。
   只改 boxToVariant 时发码一字未变，就是这个原因。

#### 二、实现

1. 抽一个共用判据 `mathNullFamilyV(node, cExpr)`（`cgen_util_ctrl.cpp`）：对 Abs/Int/Fix/Hex/Oct
   把 `vb6_X(vb6_VariantTo…(槽))` 剥回 `vb6_XV(槽)`。**判据是文本前缀**——
   `vb6_Abs(vb6_VariantToDouble(` 这串本身已证明上游真折了 Variant 槽，比那次推断硬。
2. `wrapVariantValue` 与 `boxToVariant` 都**在那张 inferExprType 表之前**先问它。
3. RTL 新增四个 `static inline vb6_VARIANT vb6_IntV/FixV/HexV/OctV`（`vb6rtl_builtin.h`）：
   `vtNull` ⇒ `vb6_VariantNull()`，否则维持原算法。⚠ 位置必须在这四枚真原型**之后**，
   否则 C11 的隐式 `int()` 声明会让它们成 C2371/C2040（本刀实测）。

#### 三、读数

| 条件 | 真值 | 改前 | 改后 |
|---|---|---|---|
| `Abs/Int/Fix/Hex/Oct(Null)`（变体位） | Null | 0（Hex/Oct 在上一刀后是 94） | **Null** |
| `Abs(-3)` / `Abs(-3.5)` | 3 / 3.5 | 同 | 3 / 3.5 ✓ |
| `Int(3.7)` / `Fix(-3.7)` | 3 / -3 | 同 | 3 / -3 ✓ |
| `Hex(255)` / `Oct(8)` | "FF" / "10" | 同 | "FF" / "10" ✓ |
| 数值位（`Abs(v) < 0.001` 这类，不过装箱站） | — | 不变 | 不变 |

#### 四、静态门登记（三面都红过，逐面归因）

1. `check_conv_strict_parse` R2：本来是**本数**（RTL 头里取 Double 的两条路各钉常数）。
   把这四枚写成头文件 `inline` 时，`vb6_IntV/FixV` 各多一个站点 ⇒ 红。
   **改法：头里只留原型，身子挪到 `vb6rtl_conv.c`（与 `vb6_AbsV` 同文件）**；
   `vb6_IntV/FixV` 走 `vb6_VariantConvNum`（文本档按 `vb6_NumVal` 严格解析，与转换站同源）。
2. `check_owned_temp_keeps`：推导器把 `vb6_FixV/IntV` 判为 **owned VARIANT**（`handoff-to-plain`），
   把 `vb6_HexV/OctV` 判为 **alias / store**（BSTR 进了 `vb6_VariantString`）。
   ⇒ 前两枚进 `ownedTempOwnedReturnExits` + `ownedTempOwnedVariantExits` 两张表，
   后两枚进 `ownedTempExceptions`（'1'/"store"，不提名）；`-Bless` 重写快照（235 行）。
3. `check_builtin_slot_returns`：派生表 `cgen_variant_slot_prefixes.inc` 依赖重跑
   `scripts/gen_variant_slot_prefixes.ps1`（50 名）。

#### 五、验收

- 探针全 30 条：M16–M19 变 `isnull=True`；M1–M15/M20–M22 与上一刀逐条相同。
- 夹具 `tests/test_joinfilter.bas` 补 `JF-nullabs/int/fix/hex/oct=True/0`，并用
  `JF-hex255=FF/0` 钉非 Null 路；本机输出与针一致。
- 门（GA）见提交信息。

#### 六、遗留

1. **`&` 与比较符的 Null 传播**（§11.36 遗留 3）：`Null & "x"` → 应为 Null（现 `"x"`）、
   `Null = ""` → 应为 Null（现 False）。要改运算符发码，且结果档要从 `BSTR`/`bool` 变 Variant，
   单独一刀。
2. **数值位的 Null**（`Dim d As Double: d = Abs(nv)`）：不经过装箱站，仍给 0；VB6 是 94
   （Null → Double）。这一格与「赋值站数值目标收 Null」同族。



### 11.39 第六十五刀：比较符的 Null 传播

上一节遗留 1 的前半（`&` 那半留下的读数见下"未做的格"）。读数 `.temp/null/oracle_ops.vbs`
（真引擎）：**`= <> < > <= >=` 只要任一侧是 Null，结果就是 Null**，包括 `Null = Null`；
`CStr(Null = "")` 则是 94（那是 Null 进 CStr 那一站）。

#### 一、为什么不在比较发码处直接改成变体位

比较的 C 类型现在是 `int32_t`：`If` 条件、赋值给 Boolean、`And/Or/Not`、`Print` 全按标量吃它。
把它整体换成 `vb6_VARIANT` 会连带一整片调用方，且 `If Null Then` 这一格**没有 VB6 读数**
（VBScript 答 false、不报错；VB6 文档口径是 94）——没读数的地方不收。

#### 二、实现（与第六十四刀同一形：变体上下文那两站各加一道守卫）

1. RTL 新增 `vb6_VARIANT vb6_VCmpNullV(vb6_VARIANT* a, vb6_VARIANT* b, int32_t cond)`
   （`vb6rtl_com.c`，原型 `vb6rtl_variant.h`）：任一侧槽是 Null ⇒ `vb6_VariantNull()`，
   否则 `vb6_VariantBool(cond ? -1 : 0)`（VB6 True 是 -1）。
2. 发码侧 `cmpNullGuardV(node, text)`（`cgen_util_ctrl.cpp`）：AST 是比较运算、且能从文本里
   抽出至少一枚 **Variant 槽地址**（`vb6_VarCmp*(&nv, …)` 的 `&nv`，或
   `vb6_VariantFromValue(nv)` 里的 `nv`）才动手；抽出两枚就都插进去，只有一枚另一侧给
   `(vb6_VARIANT*)0`。三条不齐原样返回（宁可不变）。
3. 两个站各调一次：`wrapVariantValue`（拿到的是已包好的 `vb6_VariantBool(<bool>)`，剥一层）
   与 `boxToVariant`（拿到的是**裸 bool 表达式**，直接当条件交出去）。两条形都要认，
   否则实参站那条永远不命中。

发码形状：
```
改前  r = vb6_VariantBool((int16_t)((-(vb6_StrCmp(vb6_CStr(vb6_VariantFromValue(nv)), vb6_BSTR_FromStr(L"")) == 0))));
改后  r = vb6_VCmpNullV(&nv, (vb6_VARIANT*)0, (-(vb6_StrCmp(vb6_CStr(vb6_VariantFromValue(nv)), vb6_BSTR_FromStr(L"")) == 0)));
```

#### 三、读数

| 条件 | 真值 | 改前 | 改后 |
|---|---|---|---|
| `Null = ""` | Null | True | **Null** |
| `Null <> ""` | Null | False | **Null** |
| `Null < 1` | Null | True | **Null** |
| `Null = Null` | Null | True | **Null** |
| `Empty = ""` | True | True | True ✓ |
| `"a" = "a"` | True | True | True ✓ |
| `Null & "x"` / `"x" & Null` | `"x"` | `"x"` | `"x"` ✓（真值如此，不是 Null） |

#### 四、未做的格（都有读数，留作下一刀）

| 条件 | 真值 | 现在 | 为什么没在本刀做 |
|---|---|---|---|
| `Null & Null` | Null | `""` | `&` 两侧都是槽时结果才是 Null（任一侧不是就按空串拼）；要动拼接站发码，且 `CStr(Null & Null)` 真值是 94 |
| `CStr(Null = "")` | 94 | `"True"` | 判据在 CStr 站（另一个 TU），要把 `cmpNullGuardV` 提成共享头 |
| `Null Or True` / `Null Or False` | True / Null | 94 | 逻辑位是另一族真值表（`Null And False` = False、`Null And True` = Null） |
| `Null And False` / `Null And True` | False / Null | 94 | 同上 |
| `If Null Then` | 无 VB6 读数（VBScript 答 false/0） | false | 没读数不收 |

#### 五、静态门

`vb6_VCmpNullV` 是"返回侧 owned VARIANT" ⇒ 进 `ownedTempOwnedReturnExits` +
`ownedTempOwnedVariantExits`，`cgen_variant_slot_prefixes.inc` 重新派生（51 名）；
`owned_temp_keeps_164.txt` `-Bless`（235 行 / owned 131）。
### 11.40 第六十六刀：`&` 的双 Null 折叠 + `CStr(<这两形>)` 的 94

上一节"未做的格"的前两行。读数 `.temp/null/oracle_ops3.vbs`（真引擎）：

| 条件 | 真值 |
|---|---|
| `Null & Null` | Null（TypeName 也答 Null） |
| `Null & "x"` / `"x" & Null` / `Null & 1` / `Null & Empty` | `"x"` / `"x"` / `"1"` / `""`（Null 当空串） |
| `"a" & Null & "b"` / `Null & "a" & Null` / `Null & Null & "b"` | `"ab"` / `"a"` / `"b"` |
| `Null & Null & Null` | Null |
| `CStr(Null & Null)` | 94 |
| `CStr(Null = "")` | 94（上一节记的那一格） |

**规则**：只有**两侧同时**是 Null 才折 Null；任一侧不是（含 Empty、数字、字符串）就按空串拼。
链式因为左结合 + 中间结果已是字符串，真值自动对得上（`Null & Null & "b"` = `"" & "b"` = `"b"`）；
`Null & Null & Null` 这一格仍差（我们的中间结果已经是空串，信息丢了）——记为遗留。

#### 一、实现

1. RTL `vb6_VARIANT vb6_VarConcatNullV(vb6_VARIANT* a, vb6_VARIANT* b, BSTR sa, BSTR sb)`
   （`vb6rtl_com.c`，原型 `vb6rtl_variant.h`）：两槽都是 Null ⇒ 放掉两个串、返回 `VariantNull()`；
   否则 `vb6_BSTR_Concat(sa, sb)` 后**两个串都放**，再 `vb6_VariantString(result)`。
   注意不能直接用 `vb6_BSTR_ConcatFree`：它只放左操作数（右操作数会漏），第一版就是这么写的。
2. 发码侧 `cgNullGuardConcatV(node, text)`：只认"两侧都是整枚 `vb6_CStr(vb6_VariantFromValue(<ident>))`"
   这一形（`vb6_VariantString(vb6_BSTR_ConcatFree(A, B))` 或裸 `vb6_BSTR_ConcatFree(A, B)`）；
   其它形原样不动（那些格的真值本就等于现值为：读数 R4–R6、R10、R12）。
3. CStr 站（`cgen_expr_call_conv_cstr.inc`）：`CStr(<比较>)` / `CStr(<双槽拼接>)` 先让守卫给出变体位，
   再交给 `vb6_CStrStrict`（Null ⇒ 94、其余按 Variant 档格式化，布尔仍 True/False）。
   必须放在"`CStr(String)` 是无操作"那条分支**之前** —— `CStr(<拼接>)` 的推断类型正是 String。

#### 二、两个坑（都实测过）

1. **重复套壳**：实参站会被套两层（`boxToVariant` 之后又有一层装箱），第二层拿到的已经是
   `vb6_VCmpNullV(...)`，再包一次就成"VARIANT 当 int32_t" ⇒ C2440。两个守卫都加了**幂等**检查
   （文本里已有自己的出口名就返回空）。
2. **CStr 站在另一个 TU**：`cmpNullGuardV` / `concatNullGuardV` 原来是 `cgen_util_ctrl.cpp` 里的
   `static`，CStr 站调不到 ⇒ 改名 `cgNullGuardCmpV` / `cgNullGuardConcatV` 并在 `backend/cgen.hpp`
   里声明（三个 TU 都能看见）。
3. 消费轴登记：`vb6_VarConcatNullV` 的第 2、3 格**无条件**消费（两条路都放）⇒ `kConsumes` 两行 '0'；
   第一版写 '1' 被 K13d 抓到（表 0 / 头 1 不一致）。

#### 三、读数

| 条件 | 真值 | 改前 | 改后 |
|---|---|---|---|
| `Null & Null` | Null | `""` | **Null** |
| `CStr(Null & Null)` | 94 | `""`/0 | **94** |
| `CStr(Null = "")` | 94 | `"True"`/0 | **94** |
| `Null & "a"` | `"a"` | `"a"` | `"a"` |
| `Null & 1` | `"1"` | `"1"` | `"1"` |

#### 四、遗留

1. `Null & Null & Null`（三链）→ 真值 Null，我们 `""`：中间结果已是 BSTR，槽信息丢失。要动拼接站的
   结果档（BSTR → 可 Null 的 Variant），单独一刀。
2. 逻辑位 `Null Or True` = True、`Null Or False` = Null、`Null And False` = False、
   `Null And True` = Null（现全是 94）—— 下一刀。
### 11.41 第六十七刀：`Or`/`And` 的 Null 真值表

读数 `.temp/null/oracle_logic.vbs` + `oracle_logic2.vbs`（真引擎）：

| 条件 | 真值 | | 条件 | 真值 |
|---|---|---|---|---|
| `Null Or True` | True | | `Null And False` | False |
| `Null Or False` | Null | | `Null And True` | Null |
| `True Or Null` | True | | `False And Null` | False |
| `False Or Null` | Null | | `True And Null` | Null |
| `Null Or Null` | Null | | `Null And Null` | Null |
| `Null Xor True/False` | Null | | `Null Eqv True/False` | Null |
| `Null Imp True` | True | | `Null Imp False` | Null |
| `False Imp Null` | True | | `True Imp Null` | Null |
| `Not Null` | Null（早已对） | | | |

口径与 VB6 文档一致：**Or 只有"另一侧是 True"才确定成 True，And 只有"另一侧是 False"才确定成
False，其余一律 Null**。

#### 一、实现

1. RTL `vb6_VarLogicNullV(isAnd, a, b, litBits, mask)`（`vb6rtl_com.c`）：
   mask 的 bit0/bit1 表示 a/b 是不是槽；**不是槽的那侧**用 `litBits`。
   求值顺序是这一格的要害：只传槽**地址**，`vb6_VariantToLong` 在函数里惰性算 ——
   当实参先算会在 Null 上先报 94（改前就是这么答 94 的）。
2. 发码侧 `cgNullGuardLogicV`：认四种实测形
   `vb6_VariantFromValue((vb6_VariantToLong(X) |/& (int32_t)(Y)))`、
   反向 `((int32_t)(Y) | vb6_VariantToLong(X))`、两槽 `vb6_VariantToLong(X) | vb6_VariantToLong(Y)`，
   以及**裸**形（FromValue 还没包）。某一侧既不是槽也不是 `(int32_t)(…)` 就不动。

#### 二、踩到的坑

1. **首字符是括号**：三个守卫（比较 / 拼接 / 逻辑）开头都写了
   `while (cExpr[s] == '(' || cExpr[s] == ' ') s++`，对裸形 `(<左> | <右>)` 会把配对括号吃掉，
   前缀判据全落空 ⇒ 静默不命中（本刀第一版就是这样，发码一字未变）。改成只跳空格。
2. 又一次是"两条形都得认"：存储站拿到的已包 `vb6_VariantFromValue(...)`，实参站拿裸式子。
3. 幂等：实参站会被套两层，第二层拿到的已经是 `vb6_VarLogicNullV(...)` ⇒ 返回空。

#### 三、验收

- 探针 `LogicNull.bas` 九格全对（L1 True/0、L2 Null、L3 True/0、L4 False/0、L5 Null、L6 Null、L7 Null、L8 True、L9 1）。
- 回归：`OpsNull.bas` 十四格里 Q8/Q9 也从 `-1/94`、`0/94` 变成 `True/0`、`False/0`，其余不变；
  `CatNull.bas` 六格不变。
- 夹具补 `JF-nullortrue=False/True/0`、`JF-nullorfalse=True/0`、
  `JF-nullandfalse=False/False/0`、`JF-nullandtrue=True/0`、`JF-cmplogic=False/True/0`。
- 四门本机 PASS（keeps 237 行 / 提名 193 / slot_returns 53 名 / conv_strict_parse / reclaim）。

#### 四、遗留（都有读数，暂未做）

| 条件 | 真值 | 现在 | 原因 |
|---|---|---|---|
| `Null Xor True/False`、`Null Eqv True/False` | Null | 94 | 同一族但只有"恒 Null"一列；要另开 Xor/Eqv 两条出口 |
| `Null Imp True` / `False Imp Null` | True | 94 | Imp 三态（x=False ⇒ True；y=True ⇒ True；其余 Null） |
#### 五、修正：数值侧按**真实 VB6 IDE 读数**重写（本刀收尾）

上一版按 VB6 文档的 True/False 两行做，把 `Null Or 5` 答成 Null、`Null And 0` 答成 False。
**用户 2026-11-19 在真实 VB6 IDE 里实测：`Null Or 5` = Integer 5、`Null And 0` = Integer 0**
（VBScript 同）—— 文档那两行是**不完整的**。统一规则因此改成"决定性那一侧辐到结果"：

| 运算 | 另一侧 | 结果 |
|---|---|---|
| `Or` | 非零 | **就是那一侧的值**（`Null Or 5` = 5、`Null Or True` = True） |
| `Or` | 零 | Null（`Null Or 0`、`Null Or False`） |
| `And` | 零 | **就是那一侧的值**（`Null And 0` = 0、`Null And False` = False） |
| `And` | 非零 | Null（`Null And 5`、`Null And True`） |
| 两侧都 Null | — | Null |

结果**档位跟另一侧走**（VB6 IDE 实测 `Null Or 5` 是 Integer、`Null Or True` 是 Boolean），
所以 `vb6_VarLogicNullV` 多收一个 `litIsBool`（0=Long / 1=Boolean / 2=Integer）：
槽侧看 `vt`，字面量侧由发码时问 AST 的 `LiteralKind`（`inferExprType` 对无后缀整数答 Long，
而 VB6 的 `5` 是 Integer —— 以 `LiteralKind` 为准）。为此该守卫改成 **`CCodeGen` 成员函数**
（自由函数拿不到 `inferExprType`）。

实测（`.temp/null/LogicNum.bas`）：

| 条件 | VB6 | 改后 |
|---|---|---|
| `Null Or 5` | Integer 5 | Integer 5 ✓ |
| `Null And 0` | Integer 0 | Integer 0 ✓ |
| `Null Or True` | Boolean True | Boolean True ✓ |
| `Null And False` | Boolean False | Boolean False ✓ |
| `Null Or 0` / `Null And 5` | Null | Null ✓ |

### 11.42 第六十八刀：`Xor` / `Eqv` / `Imp` 的 Null（VB6 IDE 实测）

探针 `tests/vb6probe/NullOpsProbe.bas`（用户手动在真实 VB6 IDE 跑，56 格）把这一族一次问清。

#### 一、读数

| 段 | VB6 实测 |
|---|---|
| `Xor`（X1..X7） | **任一侧 Null ⇒ Null**，`err=0`（含 `Null Xor 5`、`5 Xor Null`） |
| `Eqv`（E1..E8） | **任一侧 Null ⇒ Null**，`err=0` |
| `Imp`（I1..I8） | `Null Imp True`=**True(Boolean)**、`Null Imp 5`=**5(Integer)**、`Null Imp False/0`=Null、`Null Imp Null`=Null、`True Imp Null`=Null、`False Imp Null`=**True(Boolean)**、`5 Imp Null`=**-6(Integer)** |

#### 二、`Imp` 的口径：`x Imp y` = `(Not x) Or y`

`Imp` 不是三态表，而是先 `Not x` 再走 11.41 的 `Or` 规则（`Not Null` 仍是 Null）：

| 情形 | 推导 | 读数 |
|---|---|---|
| x Null、y 非零 | `Null Or y` ⇒ y | `Null Imp True`=True、`Null Imp 5`=5 ✓ |
| x Null、y 为零 | `Null Or 0` ⇒ Null | `Null Imp False`、`Null Imp 0` ✓ |
| y Null、~x≠0 | `(~x) Or Null` ⇒ ~x | `5 Imp Null`=~5=**-6**、`False Imp Null`=~0=-1=**True** ✓ |
| y Null、~x=0 | `0 Or Null` ⇒ Null | `True Imp Null`=~(-1)=0 ⇒ Null ✓ |

三条读数（`-6`、`True`、`Null`）只有这一种解释，反而比“另开一张三态表”更省。

#### 三、实现（不新增出口）

`vb6_VarLogicNullV` 第一枚形参由 `isAnd` 扩成 **op 码**（0=Or 1=And 2=Xor 3=Eqv 4=Imp），
已有 Or/And 的调用发码逐字节不变（仍发 `0`/`1`）。守卫 `cgNullGuardLogicV` 按 AST 的运算符
取 op 码，并认三种发码形：

| 运算 | 发码形 | 守卫怎么拆 |
|---|---|---|
| Or / And / Xor | `(L \| R)`、`(L & R)`、`(L ^ R)` | 剥外层全包括号后深度 0 找 `\| & ^` |
| Eqv | `~(A ^ B)` | 首两字符 `~(` ⇒ 剥掉再找顶层 `^` |
| Imp | `((~A) \| B)` | 顶层 `\|` 左侧是 `(~A)` ⇒ x 取 A |

文本形与 AST 运算符**必须对上**（`effOp != opCode` 就不命中），宁可不发也不发错。
非 Null 路径照旧复现原发码（`av^bv`、`~(av^bv)`、`(~av)\|bv`、`av&bv`、`av\|bv`）。

#### 四、同时修的一格：`Null Or -1` 的档位

读数 D5：`Null Or -1` 是 **Integer -1**（不是 Long）。`-1` 在 AST 里是一元负 + Integer
字面量，`inferExprType` 答 Long，所以字面量档位再加一条：一元 `Negate` 且操作数是 Integer
字面量 ⇒ Integer。

#### 五、验收

- 探针复跑（C3 自己跑同一份 .bas）：X1..X7、E1..E8、I1..I8、D1..D6 **28 格与 VB6 逐格一致**。
- 夹具新增 11 针：`JF-nullxor`/`-nullxor2`/`-nulleqv`/`-nulleqv2`、`JF-nullimp5`/`-nullimptrue`/
  `-nullimpfalse`/`-nullimpnot`/`-notimptrue`/`-notimpnull`、`JF-nullorneg`。
- 四门本机 PASS（keeps 237 行 / 提名 193 / slot_returns / conv_strict_parse / reclaim）。

### 11.43 第六十九刀：拼接链的 Null 折叠（`Null & Null & Null`）

第二批探针（`tests/vb6probe/NullFamilyProbe.bas`，用户真实 VB6 IDE）的 G 段六格：

| 格 | 条件 | VB6 实测 |
|---|---|---|
| G1 | `Null & Null & Null & Null` | **Null** |
| G3 | `(Null & Null) & (Null & Null)` | **Null** |
| G2 | `Null & Null & 1` | `"1"`（String） |
| G4 | `Null & Null & "b" & Null` | `"b"`（String） |
| G5 | `"" & Null & "" & Null` | `""`（String） |
| G6 | `Null & Empty` | `""`（String） |

即：`&` 仍按左结合二叉树逐步合并，**只在某个节点两侧同时为 Null 时**该节点折 Null，
折出来的 Null 再往上当左操作数继续走 —— 与 11.40 的两槽规则是同一条。

#### 一、根因

发码器把链**摊平成 n 元**：`Null & Null & Null` 发成
`vb6_VariantString(vb6_BSTR_ConcatFree(vb6_BSTR_ConcatFree(A, B), C))`。
而 11.40 的守卫只认“两个实参都是 `vb6_CStr(vb6_VariantFromValue(<标识符>))`”这一形，
链的左操作数是**内层拼接子树**而不是一个槽 ⇒ 永不命中（第一版就是这样，C2 仍答 `""`）。

#### 二、修法

1. 守卫改成**递归拆 `vb6_BSTR_ConcatFree` 树**收集叶子（叶子必须能取回一枚整只 Variant），
   再按左结合逐级折；叶子是两个变量槽时仍走原来的指针版（发码不变）。
2. 新增整枚接收出口 `vb6_VARIANT vb6_VarConcatNullVV(vb6_VARIANT, vb6_VARIANT)`：
   两侧都 Null ⇒ Null，否则自己 `vb6_CStr` 两串、拼接、两颗都放。
   **为什么另开一枚**：两槽版收 `vb6_VARIANT*`，而链上的操作数是内层调用的**右值**，
   取地址就得先落临时，而那枚临时里的 BSTR 没人放（漏）。
3. 叶子只认「变量槽」与「本守卫自己链出来的结果」两形；字面量档 `(int32_t)(5)`、
   BSTR 表达式一律不碰（真值已由 G2/G4/G5 覆盖）—— 它们的 C 类型不是 `vb6_VARIANT`，
   递给整枚版就是 C2440。

#### 三、登记（这一刀最贵的一步）

新增“返回 VARIANT”的出口会撞两门：`check_owned_temp_keeps` 与 `check_builtin_slot_returns`。
按规矩不能往白名单塞名字，先归因：推导把 `vb6_VarConcatNullVV` 判为 **alias**
（与已在册的 `vb6_VarConcatNullV` 同族 —— 它们的返回表达式都落到 `vb6_VariantString(...)`，
推导不把这一形认作 mint）。所以正确做法是**撤名**：两枚都不进 `ownedTempOwned*` 两份名单。
slot 表由 `gen_variant_slot_prefixes.ps1` 重生成（53 → **54** 名）。

#### 四、验收

- 探针复跑：G1..G6 六格与 VB6 逐格一致。
- 夹具补 `JF-nullcat3=True/0`、`JF-nullcat4=True/0`、`JF-nullcat4g=True/0`、
  `JF-nullcat3b=False/1/String/0`；原有 `JF-nullcat=True/0`、`JF-nullcatone=False/a/0`、
  `JF-nullcatcstr=94` 全未变。
- 四门本机 PASS（keeps 239 行 / 提名 195 / slot_returns 54 名 / conv_strict_parse / reclaim）。

#### 五、下一批（读数已齐，未做）

| 族 | VB6 实测 | 现在 |
|---|---|---|
| `Len/Trim/LTrim/RTrim/UCase/LCase(Null)` | **Null**、err 0（A1/B1/B4/B5/B6/B7） | 0/`""`、err 94 |
| `Len(Empty)` / `Trim(Empty)` | 0（Long）/ `""` | — |
| `Len(<Long 变量>)` | **4（Integer）**（A3/A5） | — |
| `Val(Null)` / `Asc(Null)` | **94** | 94（值待核） |
| `Asc(Empty)` | **5** | 待核 |
| `UBound/LBound(Null/Empty/数字)` | **13**（D1..D5） | 9 |
| `s = Null` / `s = CStr(Null)` / `s = Hex(Null)` | **94**，目标不变（S1/S4/S5、E6/E7） | err 0 |
| `d = Null` / `d = Abs(Null)` / `d = Null*1` | **94**，目标不变（N1..N4） | 0、err 0 |
| `d = <字符串表达式>` | **13**（F2） | 编译期 C2440 |
| `b = Null` / `b = (Null = "")` | **94**（F6/F7） | True、err 0 |

### 11.44 第七十刀：`UBound` / `LBound` 的变体实参（VB6 IDE 实测 → 13）

#### 一、读数（D1..D7，真 VB6 IDE）

`UBound(Null)` / `LBound(Null)` / `UBound(Empty)` / `UBound(装 5 的 Variant)` = **13 Type mismatch**；
`UBound(a)`（a 是 `1 To 3`）= 3、`LBound(a)` = 1。

#### 二、病根：不是 13 也没归到 13，是**被 9 盖掉了**

变体实参站把槽投成 `vb6_VariantToSafeArray1D(...)`，那枚宽松提取对非数组一律交回 NULL，
`vb6_UBound(NULL,1)` 于是自己抛 9 —— 所以 C3 报的是 9。

#### 三、改法与踩到的坑

1. 新出口 `vb6_VariantToSafeArray1DStrict`（`vb6rtl.c` + `vb6rtl_variant.h`）：不是数组 ⇒ 抛 13。
2. 只在 `vb6_UBound` / `vb6_LBound` 的变体实参站用它（`cgen_expr_call_arg_variant.inc` 的
   `paramIsArray` 分支按 `callee` 分流）；同一站还给 `Join` / `ReDim` 等用，那些格读数不同
   （`Join(Null)` = 94，§11.35），**不能**一起换。
3. **坑**：`vb6_ErrRaise` 在 `On Error Resume Next` 下只记号、不跳转
   （`vb6rtl.c: if (vb6_err_resume_next) return;`）。所以抛完 13 **绝不能交回 NULL** ——
   调用方 `vb6_UBound` 会再抛一个 9，把刚记下的 13 盖掉（探针第一版实测 D1..D4 全 `err=9`，
   而发码已经是 Strict，看着像“没生效”）。改为交回一枚**静态空 1D 描述符**
   （`cDims=1` + `VB6_SA1D_FFEATURES`、`uBound=-1`）：UBound/LBound 直接读边界值、不再发错，
   Err 保住 13。

#### 四、口径未动（两条负控）

- 未分配的真动态数组 `UBound(d)` 仍旧 **9**（首参不是槽，不进变体站）；
- 真数组 `UBound(a)` = 3 / `LBound(a)` = 1 不变（D5/D6）。

#### 五、登记与验收

- K5 新行直接归因、不塞白名单：`vb6_VariantToSafeArray1DStrict|0  0  <空通道>  0  open  VARIANT`
  —— 与孪生 `vb6_VariantToSafeArray1D` 同格（两者都只把入参那枚 `parray` 交回）。keeps 表 239 → **240** 行。
- 夹具：`tests/test_arr_empty.bas` 增 `UbVariantStation`，七针写入期望：
  `UB-null=13`、`UB-lnull=13`、`UB-empty=13`、`UB-scalar=13`、`UB-arr=3/0`、`UB-larr=1/0`、`UB-unalloc=9`。
- 本机：探针 D1..D7 与 VB6 逐格一致；`check_owned_temp_keeps` PASS（240 行 / 提名 196）、
  `check_builtin_slot_returns` PASS（54 名）、发码清单 750 行与 CI 逐行相同（不用重登记）。

#### 六、下一刀（已定型，未做）

`Len/Trim/LTrim/RTrim/UCase/LCase(Null)`：这一族不能只在变体站加守卫 —— 实参先被
`hoistOwnedTempArgs` 折成 `_own_N` **BSTR 临时量**，94 在造临时量那一步就抛了，`vb6_Len` 根本见不到 Null。
要另开一族变体出口（`vb6_LenV/TrimV/...` 返回 VARIANT）＋在调用站改道跳过临时量，
新出口走 `gen_variant_slot_prefixes.ps1` 自动进槽表。

### 11.45 第七十一刀：`Len`/`Trim`/`LTrim`/`RTrim`/`UCase`/`LCase(Null)` → Null

#### 一、读数（A1/B1/B4/B5/B6/B7，真 VB6 IDE）

| 格 | 调用 | 真 VB6 | C3 改前 |
|---|---|---|---|
| A1 | `Len(Null)` | **Null**、err 0 | 0、err 94 |
| B1 | `Trim(Null)` | **Null**、err 0 | `""`、err 94 |
| B4/B5 | `LTrim/RTrim(Null)` | **Null**、err 0 | `""`、err 94 |
| B6/B7 | `UCase/LCase(Null)` | **Null**、err 0 | `""`、err 94 |

这一族**不把 Null 折成串**：它把 Null 原样交回去。第六十二刀那条「String 形参收 Variant(Null) ⇒ 94」
对 `Asc`/`Replace`/`Join` 分隔符/`Split` 源是对的，对这一族不是 —— 两族在同一个实参站上，口径相反。

#### 二、病根：94 在造临时量那一步就抛了

§11.44 第六节预判的那一格被证实。改前发码是**三条语句**：

```c
BSTR _own_1 = vb6_CStrStrict(nv);
r = vb6_VariantLong(vb6_Len(_own_1));
vb6_BSTR_Free(_own_1);
```

`vb6_CStrStrict(Null)` 当场抛 94，`vb6_Len` 根本见不到那个槽。也因此**不能**学 `AbsV` 那一刀
「等到装箱站再改名」：装箱站那时只看见 `vb6_Len(_own_1)`，既认不出实参、也删不掉已经发出去的临时量语句。

#### 三、改法（四处，同一判据）

1. **RTL 六枚变体出口**（`vb6rtl_string.c` + `vb6rtl_builtin.h`）：`vb6_LenV/TrimV/LTrimV/RTrimV/UCaseV/LCaseV`，
   收 `vb6_VARIANT`、交 `vb6_VARIANT`；`vt == Null` ⇒ `vb6_VariantNull()`，其余交 base 函数的结果
   （Len 交 `VariantLong`，Trim 一族交 `VariantString`）。Empty 不经这一族（`Len(Empty)` 仍是 0，走原路）。
2. **实参站改道**（`cgen_expr_call_arg_variant.inc`）：`paramBase == String` 那一支里，
   判据成立就不折 `vb6_CStrStrict`、不造 `_own_N`，直接把整只槽交给 `*V` 出口并改 `callee`。
3. **类型 oracle 同步**（`cgen_util_type.cpp`）：`inferExprTypeImpl` 与 `isDefinitelyVariantExpr`
   的 `IndexOrCallExpr` 分支各加一格 —— 判据成立就答 Variant。
4. **For 终止值/步进折箱**（`cgen_control.cpp`）：`int32_t <var>_end = <槽>` 是 C2440；
   终止值/步进过既有那枚 `toLongIfVariant`（非 Variant 文本原样返回 ⇒ 其余语句逐字节不变）。

判据只有一处：`nullFamilyVExitCName(call)`（账 §B219）—— 被调名 ∈ 六枚 ∧ 恰一个实参 ∧
该实参的 VB 类型是 Variant（数组本体不算）。发码站与 oracle **共用**这一问：
两边不同步的代价两个方向都踩得到（oracle 说 Long ⇒ `vb6_VariantLong(vb6_LenV(nv))` C2440；
发码没改道而 oracle 说 Variant ⇒ 给裸 int32 套 `vb6_VariantToLong`）。

#### 四、踩到的坑：这四个消费面各自要一格

第一版只改「实参站 + 装箱」，`r = Len(nv)` 已经对；把消费面铺开后四形全红：

| 消费面 | 改前发码 | 处置 |
|---|---|---|
| 字符串比较 `Trim(nv) = "a"` | `vb6_StrCmp(vb6_TrimV(nv), …)` | `isDefinitelyVariantExpr` 那一格（→ `vb6_CStr(...)`） |
| 拼接 `"x" & Len(nv)` | `vb6_CStrLong(vb6_LenV(nv))` | `inferExprTypeImpl` 那一格（→ `vb6_CStr(...)`） |
| `Debug.Print` / `Print #` | `vb6_DebugWriteBSTR(VARIANT)` | 同上（→ `vb6_VariantToString` / `vb6_FileVariantText`） |
| `For i = 1 To Len(nv)` | `int32_t i_end = vb6_LenV(nv)` | For 站过 `toLongIfVariant` |

#### 五、登记与验收

- 槽表：`gen_variant_slot_prefixes.ps1` 从 RTL 派生，54 → **60** 名（六枚自动进表；
  `check_builtin_slot_returns` PASS：table=60 / rtlSlot=62 / ∩emit=58）。
- keeps：新行**直接归因**、不塞白名单 —— `vb6_LenV` 推导为 owned（与 `IntV` 同格，进两份 owned 名单），
  `TrimV/LTrimV/RTrimV/UCaseV/LCaseV` 推导为 `'1'/'store'`（BSTR 交给 `vb6_VariantString`，与 `HexV/OctV` 同理）。
  keeps 表 240 → **246** 行，提名 196 → **197**，`unclassified 0`。
- 夹具：`tests/test_joinfilter.bas` 新增 13 针并写入期望（本机实测与本表一致）：
  `JF-lenv-null/trimv-null/ltrimv-null/rtrimv-null/ucasev-null/lcasev-null=True/0`、
  `JF-lenv-tn=Null/0`、`JF-trimv-str=[ab]/0`、`JF-ucasev-str=[  AB  ]/0`、`JF-lenv-str=6/0`、
  `JF-lenv-empty=0/0`、`JF-trimv-empty=[]/0`；负控 `JF-lenv-long=94`（变体位交回的是 Null，不是 0）。
- 发码清单：整语料重跑后逐行列。本机与期望不同 **24 行**：其中 **19 行**是本刀的改道
  （b171/b172/b173×2/b174×2/b432×2/dbgdlg/bodydecl/bstr_borrow_slot/fe_slot_leak/joinfilter/
  pa_reclaim/ucSegmentedControl/ucTreeView/NullFamilyProbe/NullOpsProbe/VBFlexGridDemo），逐份用
  改前/改后两台二进制发码 diff 过，差异只有 `*V` 改道与 `_own_N` 计数位移；另 **5 行**是已记的机器侧差
  （guisample×3 + TeeChartChartDisplay + ctrlmscomm）—— 这 5 份本机改前改后逐字节相同，与这一刀无关，
  保持 CI 侧原值不抹平。
- 用例内形状断言：`b171_class_slot_arg_shape` / `b173_udt_slot_arg_shape` / `b174_num_str_arg_shape`
  三针钉的是改道前的文本（`vb6_CStrLong(vb6_Len(_own_1))`，以及 `_own_4`/`_own_6`/`_own_7` 那几枚临时量）；
  本刀少造两枚串临时量后编号前移、Len 那格变成 `CStr(VariantFromValue(LenV(...)))`，遂 GA 两个 bas 分片
  都在这三针上红（同一组三条，非分片差异）。已按改后真实发码更新，并用 Test-EmitcShape 同一判据
  （子串包含 + exit=0）复算三例全中。

#### 六、踩到的第二个坑：oracle 的“Variant”不能用 inferExprType 问

第一版判据写的是 `inferExprType(实参) == Variant`。它对 UDT 变量 / 对象 / 属性读一律**回退成 Variant**
（实参站第 77 行那条注释记的就是这个坑），于是发码没改道的形也被 oracle 判成变体位，两边不同步：

| 形 | 发码站实际 | oracle 第一版 | 后果 |
|---|---|---|---|
| `Len(bmp)`（bmp 是 UDT） | `((int32_t)sizeof(bmp))` | Variant | `vb6_VariantToLong(((int32_t)sizeof(bmp)))` |
| `RTrim$(lns(14))`（String 数组元素） | `VB6_SA_AT(BSTR, lns, 14)` 已命中 top 前缀 ⇒ 不改道 | Variant | `vb6_VariantToString(vb6_RTrim(...))` |
| `Len(ctrl.Text)`（属性读） | `vb6_Len(vb6_GetControlText(...))` | Variant | 多出 `vb6_VariantToLong(...)` 外包 |

这 26 行“假改动”带动了 47 份发码。改法（仍只一处判据）：实参节点必须是**整只槽**
（`IdentifierExpr` / `MemberAccessExpr`），类型那一半改成 `isDefinitelyVariantExpr` —— 与实参站
`isVariantSlot(文本) || isDefinitelyVariantExpr(节点)` 的**后半句同一问**。改后重发 47 份：
只剩 19 份与 base 不同，且每份的差异全是 `*V` 改道。

#### 七、下一刀

读数已齐、未做的一族仍在 §11.43 第五节：`Len(Empty)`/`Trim(Empty)`（本刀已钉住不变）、
`Len(<Long 变量>)` 的 **Integer** 档（A3/A5）、`Val(Null)`/`Asc(Null)`、`Asc(Empty)`。
`Mid/Left/Right(Null)` 与这一族同形（Null 传播），但读数未取，不跟着顺手改。

### 11.46 第七十一刀 收尾：GA 绿之后的复跑归因（探针 C3 侧读数 + With 形的编译期缺口）

#### 一、GA 收口

run `44454593300` / `e4aa570d` = **success**（`Tests (bas #1)`、`Tests (bas #2)`、`Emit manifest` 三个作业全绿）。
上一轮 `44446684530` / `df18526b` 红的只是用例内三条形状断言，已在本刀第 5 个提交里跟上改道。

#### 二、把探针在 C3 上实跑，逐格对台账

`tests/vb6probe/NullFamilyProbe.bas` 现在**整份编译不过**（F2/F3 两格 C2440，见下方第 4 条），
所以去掉那两行后实跑，得到 C3 侧现状：

| 格 | VB6 读数 | C3 实测 | 判定 |
|---|---|---|---|
| A1 `Len(Null)` | Null / 0 | Null / 0 | ✓（本刀） |
| A2 `Len(Empty)` | 0 | 0 | ✓ |
| A3/A5 `Len(<Long>)` | **4（Integer）** | 4，但 `TypeName=Long` | 值对、**返回档差一档** |
| B1..B7 / B8 | Null / `""` | Null / `""` | ✓（本刀） |
| C1 `Val(Null)` / C4 `Asc(Null)` | 94 | 94 | ✓ |
| C2 `Val(Empty)` | 0 | 0 | ✓ |
| C5 `Asc(Empty)` | **5** | 0 | ✗ |
| C6 `Mid(Null,1)` | 未取 | `""` / 94 | 待读数 |
| D1..D5 / D6..D7 | 13 / 3、1 | 13 / 3、1 | ✓ |
| E6 `s = Trim(Null)` / E7 `s = Len(Null)` | **94，目标不变** | `""` / 0 | ✗（台账改前也是 0，非本刀引入） |
| E8 `s = Null` / E9 `s = CStr(Null)` | 94（S1/S4/S5 同族） | `""` / 0 | ✗ |
| E10 `s = Hex(Null)` | 94 | 94 | ✓ |
| F1 / F5 | 0 | 0 | ✓ |
| F2 `d = Null & ""` / F3 `d = CStr(Null)` | **13** | **编译期 C2440** | ✗ |
| F6 `b = Null` | 94 | 94 | ✓ |
| F7 `b = (Null = "")` | **94** | True / 0 | ✗ |
| G1..G6 | 见 §11.43 | 与第六十九刀口径一致 | ✓ |

#### 三、新发现：`With` 体内成员是**编译期** C2440，不是运行时 94

在探针末尾临时加了 W 段（`With tw` + `Type TWV: V As Variant`），六格实测：

| 格 | 形状 | C3 发码 | 结果 |
|---|---|---|---|
| W1 | `Len(tw.V)`（UDT 成员） | `vb6_LenV(tw.V)` | ✓ Null / 0 |
| W2 | `Len(.V)`（With 体内） | `vb6_Len(_vb6_with_0->V)` | **C2440** |
| W3 | `Trim(.V)`（With 体内） | `vb6_Trim(_vb6_with_0->V)` | **C2440** |
| W4 | `Len(av(1))`（Variant 数组元素） | 原路 | 0 / **94** |
| W5 | `Len(FNullV())`（函数返回 Null） | 原路 | 0 / **94** |
| W6 | `Len(col.Item(1))`（Collection） | 原路 | 0 / **94** |

根因只有一条：`With` 体内成员在 AST 里是 **`WithMemberExpr`**，而 `nullFamilyVExitCName` 只认
`IdentifierExpr` / `MemberAccessExpr`，`isDefinitelyVariantExpr` 也没有这一 `case` ⇒ 实参站判
「不是变体」⇒ 既不造串临时量、也不改道 ⇒ 把整只 `VARIANT` 递给 `vb6_Len(BSTR)`。
**这是改前就有的**（改道只加路，不加则与改前同形状），不是本刀回归。

探针本体**没有**入仓：`NullFamilyProbe.bas` 是发码清单在册行（`bytes=25468`），加格必须与
清单重登记同一刀；扩展版留在 `.temp/k72/NullFamilyProbe.ext.bas`，改道修好那一刀再一并登记。

#### 四、下一刀的排序（按硬度，不按族）

1. **With 成员（编译不过）** —— 给 `WithMemberExpr` 补一条解析（With 帧里取目标档位），
   同时进 `nullFamilyVExitCName` 与 `isDefinitelyVariantExpr`，两处必须同一问。
2. **String 目标收 Null ⇒ 94**（E6/E7/E8/E9）：`Trim(Null)` 本刀已交真 Null，这一站再静默转空串
   就与本刀口径自相矛盾。
3. **F2/F3 的 13**（现在编译期 C2440）与 **F7**。
4. W4/W5/W6（需先取 VB6 读数）、C5 `Asc(Empty)=5`、A3/A5 的 Integer 档。

### 11.47 第七十二刀：`With` 体内成员的变体槽（编译期 C2440 → Null 传播）

#### 一、读数与病根

探针 W 段（真 VB6 读数待回，C3 侧是先量的）：`With tw`（`Type TWV: V As Variant`）体内的
`Len(.V)` 改前发成 `vb6_Len(_vb6_with_0->V)` —— 整只 `vb6_VARIANT` 进 `BSTR` 形参 = cl
**C2440**，改前就没有 exe（不是运行时 94）。同一探针里 W1 `Len(tw.V)`（UDT 成员）一直正常。

根因：`With` 体内成员在 AST 里是 **`WithMemberExpr`**，而 `nullFamilyVExitCName` 只认
`IdentifierExpr` / `MemberAccessExpr`，`isDefinitelyVariantExpr` 也没有这一 `case` ⇒
实参站判「不是变体」⇒ 既不造串临时量、也不改道。

#### 二、第一版（改全局判据）与它踩的坑

第一版直接把 `WithMemberExpr` 加进 `isDefinitelyVariantExpr`。B171 与探针都对了，但
**Charts 2020 / ucTreeMaps 两台当场编译不过**：

```
int32_t j_end = vb6_UBound(vb6_VariantToSafeArray1DStrict(_vb6_with_2->Values), 1);
                                        ^ vb6_SafeArray1D* → vb6_VARIANT = C2440
```

原因：`isDefinitelyVariantExpr` 是**全局**判据（UBound 的数组槽、装箱站、比较站都吃它），
而 `With` 成员里既有 `As Variant` 字段，也有 `SafeArray` 字段。教它认 With 成员，等于把
`.Values` 也判成变体槽。

#### 三、改法：只让 Len 族那一格认（两处，同一问）

1. `nullFamilyVExitCName`：`WithMemberExpr` **只在这一格**认 —— 档位问
   `inferUdtFieldVb6Type`（`With <UDT>`，它自己就认 With 帧 + `knownUdtVars_`）或
   `classFieldKind`（`With <类实例>` 的字段）；数组本体与「属性 Get」一律不收。
2. 发码站（`cgen_expr_call_arg_variant.inc`）：在 `argIsVariant` 那句之后加一枚
   `vexit72 = nullFamilyVExitCName(node)`，问到即视为变体槽。没有这一枚，String 支根本进不去
   （`isVariantSlot` 的文本前缀不认识 `_vb6_with_N->F`），整只 VARIANT 会被裸递给 `vb6_Len`。

口径：只影响六枚内置的单实参调用；`isDefinitelyVariantExpr` 一个字没改。

3. 类实例那一支**问投影不问底层表**：`classFieldIsVariantSlot`（它自己转问 `classFieldKind`，
   是同一个答案的唯一出口）。第一版直接调 `classFieldKind`，GA 当场红在
   `[STATIC] variant_slot_store` 的 **R8**：`cgen_util_type.cpp` 的直接问点被钉成 4 处
   （读侧三站 + §B184 写侧一问），我这一调就成了第 5 处。不改那个数（它就是用来防「读侧
   又多长一个问点」的），改问投影即归位，发码逐字节不变。

#### 四、验收

- B171：`vb6_LenV(_vb6_with_0->PZ  /* With class .PZ field variant */)`，且那枚
  `_own_2 = vb6_CStrStrict(...)` 消失（`bytes` 12727 → **12600**）。形状针新增一枚钉住这一格。
- Charts 2020 / ucTreeMaps：发码改前改后 **diff=0**（逐字节相同），编译恢复 `rc=0`。
- 探针：W2 `Len(.V)` / W3 `Trim(.V)` = **Null / 0**（改前 C2440），W1 不变。
- 门：`check_builtin_slot_returns` PASS（table=60）、`check_conv_strict_parse` PASS。
- 清单：750 行里 7 行变 = 本刀 2 行（B171 + 探针）+ 5 行已记的机器侧差；只登记那 2 行。

#### 五、留下的

- W4/W5/W6（Variant 数组元素 / 返回 Null 的函数 / `Collection.Item`）仍 94 —— 读数未取。
- `With <类实例>` 的**方法**（属性 Get）返 Null 不在本刀：字段表只认字段。

### 11.48 第七十三刀：Null 进 String 目标 ⇒ 94，且**目标不变**

#### 一、读数

探针 E 段：`s = Trim(Null)` / `s = Len(Null)` / `s = Null` = **94**，且 `s` 仍是赋前的值
（“keep”、len=4）—— 赋值根本没发生。与 §11.43 第五节 S1/S4/S5 同一族。

#### 二、改前：两个错同时存在

`emitBstrSlotStore`（String 槽收值的唯一出口）把 Variant 右值一律折成
`vb6_VariantToString`：Null 变 `""`、赋值照做 —— 既丢值，也不报 94。

#### 三、改法

1. RTL 新出口 `vb6_BSTR_AssignFromVariant(BSTR*, vb6_VARIANT)`：`vt == Null` ⇒
   `vb6_ErrRaise(94, "VBA.Conversion", "Invalid use of Null")` 后**直接 return**
   （目标一字不动）；否则 `vb6_CStr` + `AssignMove`。
   为什么不直接用 `vb6_CStrStrict`：它 raise 后交回空串，赋值方会把空串写进目标，
   变成“94 + 目标被清空”——与 VB6 不同。
2. `emitBstrSlotStore`：`valueIsVariant` 时留下未折的表达式，定长串以外的赋值改发这枚出口。
   定长串（`fixedLen` 非空）那一档读数未量，保持原路。

#### 四、验收

- 探针：E6/E7/E8 = `err=94 val=[keep]`（改前 `err=0 val=[]`）；E1/E4/E5/E10 不变。
- 发码清单：20 行变 = **15 行本刀归因**（逐份看过，全部是
  `Assign(&X, vb6_VariantToString(E))` → `AssignFromVariant(&X, E)`；dbgdlg 10 处、
  VBFlexGridDemo 9 处、Charts 2020 4 处为最大三份）+ 5 行已记的机器侧差。
- 门：`check_builtin_slot_returns` PASS（table=60）、`check_conv_strict_parse` PASS、
  `check_variant_slot_store` PASS。

#### 五、留下的（按 §11.47 的排序）

- **E9 `s = CStr(Null)`** 不在本刀：它走 CStr 内建那一站（`vb6_CStr` 宽松），要动得先决定
  “CStr(Null) 报 94 之后赋值是否跳过”——同一组读数、另一个站，另立一刀。
- F2/F3（13，现为编译期 C2440）、F7（94）、C5 `Asc(Empty)=5`、A3/A5 的 Integer 档。
- W4/W5/W6（Variant 数组元素 / 返回 Null 的函数 / `Collection.Item`）仍 94，读数未取。

#### 六、R7 的前缀误伤

R7 原本用 `emitLine("vb6_BSTR_Assign` 前缀匹配，把本刀新增的
`vb6_BSTR_AssignFromVariant` 也算成一枚“手抄裸赋值”（3 → 4）。新出口是**另一个契约**
（Null ⇒ 94 且目标不变），不是裸赋值；改判据为显式两枚（`Assign(` / `AssignMove(`），
名单仍是 3 —— 没有放宽，只是不再把不同函数名当成同一件事。

#### 七、GA 红（run 44486983023，head `ce5bcfa1`）五处，逐处归因

| 格 | 读数 | 归因 |
|---|---|---|
| `Tests (bas #1)` / `(bas #2)` | `[EMITC-SHAPE] cvp_emitc_variant_fallback ... missing` 两条旧形状 | 夹具针钉的是改前的 `Assign(&sN, VariantToString(...))`；本刀这两格正在改道 → 换针 |
| `Tests (syntax)` | `[CODEGEN-NOTE] variant_unbox_shapes ... missing` 两条旧形状 | 同一类：`Assign(&s, VariantToString(VB6_SA_AT(...).dItem))` 已改走严格出口 |
| `Tests (compile #1)` | `[STATIC] owned_temp_keeps`：K12 一枚例外行发码侧没有 / K5 漂移 `+vb6_BSTR_AssignFromVariant|1` | 新 RTL 出口是**带权形参位**（第 1 格 VARIANT 带 `store` 通道）——表与发码侧头都得有这一行，我前一刀只跑了三道门禁、没跑 KEEPS 双向对账 |
| `Emit manifest` | 相同 747 / 哈希不同 3 = guisample 三份 | 就是上一条预告的机器侧三行；按 CI 自己那份 `ci/emit-manifest`（head 同为 `ce5bcfa1`）登记，不用本机数 |

三处都是**未登记的后果**，不是改道的逻辑错：形状针与 KEEPS 表是“同一件事的另外两处记账点”，
这一刀的遗漏在于只跑了源码侧三道门禁（slot_returns / conv_strict_parse / variant_slot_store），
没跑 `check_owned_temp_keeps.ps1` 与两条形状针所在的类别。

修法（都不放宽判据）：

- 两条形状针逐字换成实测新发码（`vb6_BSTR_AssignFromVariant(&s1, vb6_VariantFromComResult(...))`、
  `... (&s, VB6_SA_AT(...).dItem)`），改由本机 `--emit-c` 逐字核对，不凭猜。
- KEEPS：`-Bless` 让推导侧把新行落进快照 `scripts/owned_temp_keeps_164.txt`
  （`vb6_BSTR_AssignFromVariant|1  1  store  0  open  VARIANT`），
  再在发码侧 `src/common/owned_temp_keeps.hpp` 同步一行（`{ "vb6_BSTR_AssignFromVariant", 1, '1', "store" }`）。
  K12 与 K5 的性质是**同一份事实的两侧**，只补一侧仍然红。
- 清单三行取 CI 的值（本机那台 `bytes` 比 CI 多 10，跨机器差的常态）。

本机复核：`check_owned_temp_keeps` PASS（surface 247 / rows 247 / nomination 197）；
`-Category syntax` 整段 PASS=227 FAIL=0（含 `variant_unbox_shapes`）；
`cvp_emitc_variant_fallback` 的针按本机发码逐字核对（bas 整段本机超 30 分钟上限，交 CI）。

### 11.49 第七十四刀：`s = CStr(Null)` ⇒ 94 且目标不变（E9）

#### 一、读数

探针 E9（`s` 先置 "keep"）：改前 `len=0 val=[] err=0`，改后 `len=4 val=[keep] err=94`。
真 VB6 是 94 —— 与 E6/E7/E8 同一族（第七十三刀），只是入口站在 CStr 那一侧。

#### 二、归因

`s = CStr(x)` 的发码是 `vb6_BSTR_AssignMove(&s, vb6_CStr(<x>))`：CStr 内建对变体走的是
**宽松口** `vb6_CStr`（`vb6_Format`，Null ⇒ ""）。所以两个错同时存在：值被清成空串（目标变了）
+ err=0（该报的 94 没报）。而这条路根本不经过 `emitBstrSlotStore` 的变体折叠 ——
右值在发码表里已是 BSTR，产业是「自有临时串」⇒ 走 AssignMove。

不能把 `vb6_CStr` 改严：`&` 与 `Debug.Print` 仍靠它的宽松口径（前者 Null 传播，后者打印 "Null"）；
也不能改成 `vb6_CStrStrict`：它 raise 后交回空串，写进目标就是「94 + 目标被清空」。

#### 三、实施

`emitBstrSlotStore` 在变体折叠那一支之前多问一句：右值是**单实参 CStr 调用**、
那一枚实参推出来是**变体档**、且发码确实是 `vb6_CStr(` ⇒ 置
`strictVarAssign221` + `rawVarValue221 = <CStr 括号里那一段>`，交给第七十三刀那枚出口
`vb6_BSTR_AssignFromVariant`。出口内部就是 `AssignMove(target, vb6_CStr(src))` ⇒
非 Null 逐字节等价，Null 报 94 后直接返回。

边界（都是刻意不动）：

- 第六十六刀那两条 `vb6_CStrStrict(<比较>)` / `vb6_CStrStrict(<拼接>)` —— 前缀就不是 `vb6_CStr(`。
- 定长串（`fixedLen` 非空）—— 读数未量，仍走原路。
- 目标不是 String（`F3 d = CStr(Null)`）—— 那是另一个站（double 目标），仍编译期 C2440，留在 F 段。

#### 四、清单

11 份变：逐份 base/new 对照（`git stash push -- src/` → rebuild → emit → pop → rebuild），
**6 份是真变**，形状全部同为 `AssignMove(&X, vb6_CStr(v))` → `AssignFromVariant(&X, v)`：
B168 1 处、GCCache 8 处、test_fe_slot_leak 1 处、test_pa_reclaim 1 处、
NullFamilyProbe 3 处、NullOpsProbe 2 处。另 5 份（guisample×3 / ctrlmscomm / TeeChart）
本机 base 与 new **逐字节相同** ⇒ 它们只是长期存在的机器侧行，不是本刀改的。

#### 五、门禁

`check_builtin_slot_returns` / `check_conv_strict_parse` / `check_variant_slot_store` /
`check_owned_temp_keeps` 四道全 PASS；`-Category syntax` PASS=227 FAIL=0；
两枚受影响用例 `test_fe_slot_leak` / `test_pa_reclaim` 实跑 rc=0。

### 11.50 第七十五刀：`d = Null & ""` / `d = CStr(Null)` ⇒ 13（F2/F3）

#### 一、读数

探针 `d` 先置 1.5：改前**编译期 C2440**（BSTR 裸赋 double，两处）；
改后 `d=1.5 err=13`（目标不变），与真 VB6 两格都是 13 一致。

#### 二、归因

Double 目标 ← String 右值这条路上**根本没有转换站**：Variant 右值有 `vb6_VariantToDouble`
（F1 `d = Empty` 就走那一条），而 String 右值（拼接、`CStr(...)`）被原样交给 double ⇒ C2440。
真 VB6 是隐式 String→Double：数字串折成数、其余（含空串）报 **13**，且**目标不变**。

#### 三、实施

- RTL 新出口 `vb6_Dbl_AssignFromBSTR(double* target, BSTR src)`：去首尾空白后 `wcstod`，
  整串吃完才算数；否则 `ErrRaise(13, "VBA.Information", "Type mismatch")` 后直接返回。
  **不用** `vb6_NumVal`：那是 Val 口径（永不报错，`"2.5abc"` 也折 2.5），而 VB6 的赋值转换是严格的。
- 发码站（`cgen_assign_value_sem.inc` 非 Variant 那一支的末尾）：右值语义档 = String、
  目标 = Double（先问 `inferExprType`，再回查 `knownDoubleVars_` —— 局部变量在语义层可能只答
  Unknown）⇒ 改发 `vb6_Dbl_AssignFromBSTR(&target, value)`。

边界（刻意不动）：Single / Currency 目标（C 类型不同）、Date 目标（已有 `knownDateVars_` +
`vb6_DateVal` 的 CDate 口径）、Variant 目标（走装箱那一支）。

#### 四、清单与 KEEPS

全量 750 行：**只有 `tests/vb6probe/NullFamilyProbe.bas` 一行变**（就是 F2/F3 那两句的发码），
另 5 份（guisample×3 / TeeChart / ctrlmscomm）是长期存在的机器侧行，本刀没动它们。
KEEPS 新行 `vb6_Dbl_AssignFromBSTR|1  0  (空通道)  0  open  BSTR` —— keeps_pointer=0，
**不是例外行**，所以 `src/common/owned_temp_keeps.hpp` 这份发码侧表不动（K12 仍然绿）。

#### 五、门禁

`check_builtin_slot_returns` / `check_conv_strict_parse` / `check_variant_slot_store` /
`check_owned_temp_keeps` 四道 PASS；`-Category syntax` PASS=227 FAIL=0。

### 11.51 Null 扩展探针读数（VB6 IDE 实跑，`ai/vb6probe/NullExtProbe.bas`）

真机一次跑完 A/W/C/D 四段（`On Error Resume Next`，每格前 `Err.Clear`，`s` 先置 `keep`）：

```
A1  v = (Null = "")  | isnull=True val=<Null> tn=Null err=0
A2  d = (Null = "")  | d=1.5 err=94
A3  s = (Null = "")  | s=[keep] err=94
A4  b = (Null = "")  | b=True err=94
A5  v = (Null <> "") | isnull=True val=<Null> tn=Null err=0
A6  v = (Null = Null) | isnull=True val=<Null> tn=Null err=0
A7  v = (Null = 0)    | isnull=True val=<Null> tn=Null err=0
A8  v = (Null > 0)    | isnull=True val=<Null> tn=Null err=0
A9  s = (Null <> "") | s=[keep] err=94
W4  s = a(0)          | s=[keep] err=94      (变体数组元素)
W4b n = Len(a(0))     | n=-1 err=94
W5  s = FNull()       | s=[keep] err=94      (返回 Null 的函数)
W5b n = Len(FNull())  | n=-1 err=94
W6  s = c(1)          | s=[keep] err=94      (Collection.Item)
W6b n = Len(c(1))     | n=-1 err=94
W7  v = a(0)          | isnull=True val=<Null> tn=Null err=0   (入口确实是 Null)
C6  s = Mid(Null,1)   | s=[keep] err=94
C7  s = Left(Null,1)  | s=[keep] err=94
C8  s = UCase(Null)   | s=[keep] err=94
D1  lv = Len(<Integer var>) | val=2 tn=Integer err=0
D2  lv = Len(<Long var>)    | val=4 tn=Integer err=0
D3  lv = Len("abc")          | val=3 tn=Long    err=0
```

四条结论（直接改写 F7 与 Len 两条的做法）：

1. **比较结果本身就是 Null**：`A1/A5/A6/A7/A8` 全 Null（连 `Null = Null` 也是 Null）⇒
   Null 传播适用于**全部**比较运算符，不限于 `=`。C3 现在发 `vb6_StrCmp(...) == 0` ⇒ True/err=0，
   所以 F7 要动的是**比较站**，不是赋值站。
2. **Null 落到目标上一律 94 且目标不变**（A2 Double / A3 String / A4 Boolean）。
   这与 F2/F3 的 13 不冲突：那是**字符串** `""` 进 Double（严格数值解析 13），
   这是**Null** 进 Double（变体通道 94）—— 两条路必须分开。
3. **W 段三入口同口径**（变体数组元素 / 返回 Null 的函数 / Collection.Item）：
   赋给 String 与 `Len(...)` 都是 94 且目标不变。第七十三刀的严格出口已覆盖这一形，
   所以 W4/W5/W6 很可能**已经是绿的**，只需复测确认；C6/C7/C8 （Mid/Left/UCase）那一道站
   则是旧口径，未改成 94。
4. **Len 的档位按实参类别分**：`Len(<Integer/Long 变量>)` ⇒ **Integer**（2 / 4），
   `Len("abc")` ⇒ **Long**（3）。所以 §11.46 那句“值对、档差一档”要改成**两档**：
   数值实参发 Integer（VT_I2），字符串实参发 Long；一刀切两向都会错一格。

### 11.52 第七十六刀：`b = Null` / `d = Null` ⇒ 94 且目标不变（A4/A2）

#### 一、读数（本机 C3）

`tests/vb6probe/NullFamilyProbe.bas` 的 F 段（改后）：

```
F1  d = Empty       | d=0   err=0
F2  d = Null & ""    | d=1.5 err=13
F3  d = CStr(Null)  | d=1.5 err=13
F4  d = 0           | d=0   err=0
F5  i = Empty       | i=0   err=0
F6  b = Null        | b=True err=94      (改前 b=False err=0)
F7  b = (Null = "")| b=True err=0       (没改，见三)
```

另加一份临时探针（`.temp/k76/f8.bas`，不入清单）补 A2 的 Double 格：

```
F8  d = Null        | d=1.5 err=94
F9  b = Null        | b=True err=94
```

对得上真机 A2/A4：**报 94 且目标不动**（改前 `target = vb6_VariantToLong(...)` 是先 raise
再写 0/False，两半都错：错成 0、还把目标清了）。

#### 二、改了什么

- RTL 两枚严格出口：`vb6_Bool_AssignFromVariant(int16_t*, vb6_VARIANT)` /
  `vb6_Dbl_AssignFromVariant(double*, vb6_VARIANT)` —— `vtNull` ⇒ `vb6_ErrRaise(94)` 后**直接 return**
  （目标一个字节都不写）；否则走原来的 `VbVariantToLong` / `VbVariantToDouble`；`Empty` ⇒ 0。
- `cgen_assign_value_sem.inc` 的 `valueIsVariant` 支首：按 `inferExprType(*node.target)` 是
  Boolean / Double 就改发上面两枚。目标类型问不出来就**维持原路**（不猜）。
- 与 13 那条路分开：字符串右值进 Double 仍走第七十五刀的 `vb6_Dbl_AssignFromBSTR`（严格解析，`""` ⇒ 13）；
  变体右值进 Double 走本刀这枚（Null ⇒ 94）。

#### 三、F7 为什么没一起改（改到一半撤回的证据）

先按 A1/A5–A8 的结论动**比较站**（`cgen_expr_binary.cpp`）：任一侧是变体就发
`vb6_VarCmpNullV(a, b, op)` 交出一整只 VARIANT，想靠前缀表让下游按目标类型走 94 通道。三条硬红：

1. `-Category syntax` 的 `varcmp_scalar_boxed` 红：数字比较也被改道了，缺 `vb6_VarCmpEq(&v, `。
2. 同批 `variant_unbox_shapes` 红：Double 目标的形状换了（那一行本刀**确实**该换，已重钉）。
3. 最硬的一条：把**有字符串一侧**的比较收窄改道之后，`tests/test_varcmp_scalar.bas` 的发码里出现
   `vb6_VCmpNullV(&v, &s, vb6_VarCmpNullV(...))` 与 `(int32_t)(vb6_VarCmpNullV(...))` ——
   这两处包装站把**比较结果当 int 用** ⇒ C2440，整份工程编不过。

所以 F7 不是「再加一条出口」：比较站一交出 VARIANT，`vb6_VCmpNullV` 的第 3 实参、`(int32_t)(...)`
这两类包装站都得先能接 VARIANT（或者改成由包装站决定内层比较怎么发）。本刀**整体回退**：
`vb6_VarCmpNullV` 的 RTL / 声明 / 前缀表 / KEEPS 行一起撤，不留死代码。F7 留给下一刀。

#### 四、登记与门禁

- KEEPS：新增 2 行（`vb6_Bool_AssignFromVariant|1 1 store` 例外行 + `vb6_Dbl_AssignFromVariant|1 0 handoff-reads:vb6_VariantToDouble`），
  `check_owned_temp_keeps.ps1` PASS（surface 250 / owned 133）。
- 发码清单：全语料 750 行里**本机 5 行动**，逐份 base/new 归因（`git stash push -- src/` 那套）全是本刀的改道文本：
  NullFamilyProbe（4 行：`d = vb6_VariantToDouble(ev)`→`vb6_Dbl_AssignFromVariant(&d, ev)`、`b = vb6_VariantToLong(nv)`→`vb6_Bool_AssignFromVariant(&b, nv)`）、
  NullOpsProbe（6 行）、test_variant_boxing（4 行）、VBFlexGridDemo/Common/Common.bas（2 行）、VBFlexGridDemo.vbp（项目级）。
  机器侧 5 行（guisample×3 / TeeChart / ctrlmscomm）**base==new**，一律不动。
- 门禁：`builtin_slot_returns` / `conv_strict_parse` / `variant_slot_store` / `owned_temp_keeps` 全 PASS；
  `-Category syntax` PASS=227 FAIL=0 SKIP=0。
- 已知噪声（下一刀顺手收）：`b = Null` 仍报编译警告 **VB3003**「无法将 Null 隐式转换为 Boolean」，
  但运行期口径已与 A4 一致；F7 仍在（`b = (Null = "")` ⇒ True/0，真机 94/True）。

### 11.53 GA run 44521823896 的 build 红：不是源码，是构建目录缓存跨镜像（已修）

#### 一、现象

手动 GA 在 `Build C3.exe` 这一格红，失败步骤是**第 8 步里的 cmake configure**（不是编译）：

```
-- The CXX compiler identification is unknown
CMake Error at CMakeLists.txt:11 (project):
  The CMAKE_CXX_COMPILER:
    C:/Program Files/Microsoft Visual Studio/18/Enterprise/VC/Tools/MSVC/14.51.36231/bin/Hostx64/x64/cl.exe
  is not a full path to an existing compiler tool.
```

同一格的前两行是 `cmake version 4.5.0` 与 `MSVC install: C:\Program Files\Microsoft Visual Studio\18\Enterprise`
—— 镜像里 VS 在，只是**工具集版本目录换了**，而 14.51.36231 那条路径已经不在了。

#### 二、归因（读日志，不猜）

```
Cache not found for input keys: c3exe-v2-766b6f97...          ← 本提交没有 C3.exe 缓存
key: c3build-v2-766b6f97...   restore-keys: c3build-v2-
Cache hit for restore-key: c3build-v2-cd3dc6b4114ad...       ← 兜底命中**旧提交**留下的构建目录
```

所以链条是：本刀的 `src/` 改动的 key 没命中 ⇒ `actions/cache` 的 restore-keys 兜底 ⇒ 恢复回来的是
另一台镜像（旧的 VS 18 / MSVC 14.51）留下的 `$RUNNER_TEMP/c3build`，它的 `CMakeCache.txt` 钉着那台镜像的
**绝对** `cl.exe` 路径 ⇒ 现镜像上该路径不存在 ⇒ cmake 4.5 在 `project()` 处直接拒。
与本次改的 RTL/发码无关，也和 14.5 与 14.29 的差异无关：**同一个坑对任何一次 src 改动都成立**，
只是这一次轮到了。

#### 三、修法（四处 configure 站一起改）

1. **编译器每次显式给值**：`$clPath = Join-Path $env:VCToolsInstallDir "bin\Hostx64\x64\cl.exe"`
   （vcvars 已导出该变量），并把 `"-DCMAKE_CXX_COMPILER=$clPath"` 写进 configure 行 —— 命令行 `-D`
   覆盖缓存项（与 launcher 那条同一个思路：不让恢复的 CMakeCache 代决定）。
2. **旧镜像的目录整份作废**：恢复回来的 `CMakeCache.txt` 里若记的编译器路径与本次不同，
   则 configure 加 `--fresh`（文档语义：丢掉 CMakeCache.txt 与 CMakeFiles，不动源码、不动 `.build`）。
   只覆盖编译器不够：`CMAKE_LINKER` / `CMAKE_AR` 同样是旧镜像的绝对路径，链接期会变成 "cannot find file"。
   路径相同则照旧走快速路。
3. 四个站：`ci.yml`（Build C3.exe）、`ci_t0.yml`（T0 构建）、`ci_vbman.yml` 两处（其中 else 支没有
   `build.ninja`，说明是半截缓存，直接 `--fresh`）。三份 YAML 都过 `yaml.safe_load`。

#### 四、本机复现与验证（不是“看着像就这么改”）

- **复现**：把一份配好的 `CMakeCache.txt` 的 `CMAKE_CXX_COMPILER` 改成 `C:/nonexistent/msvc/cl.exe`，
  按 CI 原样重配 ⇒ 一字不差的 `is not a full path to an existing compiler tool`，rc=1。
- **验证 1（只加 -D）**：把同一份缓存里的路径改成「另一台镜像的 VS 18 / 14.51.36231」，加
  `-DCMAKE_CXX_COMPILER=<本机 cl>` 重配 ⇒ `identification is MSVC 19.29...` / `Configuring done`，rc=0。
- **验证 2（整段脚本）**：把 workflow 里那十几行原样跑一遍（vcvars 注入环境 → 读 CMakeCache → 判定
  跨镜像 → `cmake @cmakeArgs ...`），输出
  `c3build cache from another image (C:\Program Files\Microsoft Visual Studio\18\...\14.51.36231\...cl.exe) -> cmake --fresh`
  然后 `Configuring done`，rc=0。
- `cmake --fresh` 本机 4.4.4 可用（文档要求 ≥ 3.24），runner 是 4.5.0。

### 11.54 第七十七刀：`s = c(1)`（Collection.Item 是 Null）⇒ 94 且目标不变（W6）

#### 一、复测读数（本机 C3，探针 `.temp/k77/wc.bas`，改前）

真机读数取自 §11.51 的 W/C 段。改前 C3：

```
W4  s = a(0)          | keep / 94   ✅ 同真机
W4b n = Len(a(0))     | n=0  / 94   ✗ 真机 n 保持 -1
W5  s = FNull()       | keep / 94   ✅
W5b n = Len(FNull())  | n=0  / 94   ✗
W6  s = c(1)          | ""   / 0    ✗✗ 真机 keep / 94
W6b n = Len(c(1))     | n=0  / 94   ✗
C6  s = Mid(Null,1)   | ""   / 94   ✗ 真机 keep
C7  s = Left(Null,1)  | ""   / 94   ✗
C8  s = UCase(Null)   | ""   / 94   ✗
```

两类缺陷：

1. **W6 连错误号都没有**：入口确属 Null（`.temp/k77/w6.bas` 的 P1：`w = c(1)` ⇒ `isnull=True`），
   P3 `s = <Null 变量>` 已经是 keep/94（第七十三刀），P4/W6 `s = c(1)` 却是 ""/0。
2. **其余六格的 94 对了，目标却被写了**：`n = Len(...)` → 0、`s = Mid/Left/UCase(Null)` → 空串。
   真机是「语句在出错处中止，赋值根本没发生」。

#### 二、W6 的归因（本刀修的）

发码：

```c
vb6_BSTR_Assign(&s, vb6_VariantToString(vb6_VariantFromComResult(vb6_ComCallByDispid(c, 0, ...))));
```

缺口在 `emitBstrSlotStore` 的 COM 裸结果那一支：它把变体**提前**折成串再赋。
`vb6_VariantToString(Null)` 不 raise、交回空串 ⇒ err=0 + 目标被清成 ""，两个错同时存在。
Null 必须在**赋值那一瞬**才判：改成把整只变体交给 `vb6_BSTR_AssignFromVariant`
（非 Null 走 `vb6_CStr`，与 `VariantToString` 同口径；定长串目标仍走原来的转串路）。
改后：

```
W6  s = c(1)          | keep / 94   ✅
P4  s = c(1)          | keep / 94   ✅（P1/P2/P3 不变）
```

门禁：四道全 PASS；`-Category syntax` PASS=227 FAIL=0。发码清单：全语料 750 行本机 8 行动
（Charts×3 / fgz_fix005a×2 / test_com_latebound / ucSegmentedControl / ucTreeView），
逐份 base/new 归因（base==expect ∧ base!=new）；机器侧 5 行本刀 **base==new**（本机发码一字未动），
保持 CI 原值。

#### 三、还欠的（下一刀）——“94 之后不许写目标”

剩下六格全是同一件事：**语句级中止**。现在发码是
`BSTR _own = vb6_CStrStrict(<Null>); n = vb6_Len(_own);` —— CStrStrict raise 后交回空串，
赋值站照写。真 VB6 是「出错即中止本语句」。要真修需引入一个单调的**错误序号**
（`vb6_ErrRaise` 自增），赋值/输出站在写目标前比一次，若本语句里 raise 过则不写。
这会给**每条赋值**加一层守卫 ⇒ 发码形状大面换；必须单独一刀、单独登记，不能混在本刀里。

#### 四、GA run 44532858724 的红：`fgz_fix005a_default_member_unwrapped`（已修）

GA 的 `Tests (vbp #9)` 报 `PASS=54 FAIL=1 SKIP=2`，注解（check-runs/annotations）点名：

```
[EMITC-SHAPE] fgz_fix005a_default_member_unwrapped ... exit=0
  missing: vb6_VariantToString(vb6_VariantFromComResult(vb6_ComCallByDispid(
```

就是本刀换掉的那个形状：那条针钉的是「默认成员结果先拆成整只变体再落槽」，而本刀把 String 槽那一支
从 `vb6_VariantToString(...)` 改成 `vb6_BSTR_AssignFromVariant(&s, ...)`。「先拆箱」这个意思没变，
变的只是落槽那一枚出口 —— 针跟着改：

```
'vb6_BSTR_AssignFromVariant(&s, vb6_VariantFromComResult(vb6_ComCallByDispid('
```

（本机发码逐字核过：`vb6_BSTR_AssignFromVariant(&s, vb6_VariantFromComResult(vb6_ComCallByDispid(d, 0, ...`。）

教训（记下来）：这类 `[EMITC-SHAPE]` 用例住在 **vbp** 类里，`-Category syntax` 看不到 ——
以后改了「写槽」那一层，本机至少要跑 `-Category vbp -VbpShard <CI 红的那一片> -VbpShardTotal 12`。
本机重跑 9/12 片：本针不再报缺；另有两格本机红（`[VBP] c29listview`、
`[COMPILE-FAIL] iface_multi_incomplete_negcontrol`）—— 同一 HEAD 上 CI 这两格是 PASS，属本机环境差
（OCX/编译器可用性），与发码无关。

### 11.55 第七十八刀（已撤回）：`Asc("")` / `Asc(Empty)` 报 5 试过，与 94 那条路揭不开

#### 一、本机读数（`.temp/k78/lena.bas`）

```
改前: L1 n = Asc("") | n=0 err=0     L2 n = Asc(Empty) | n=0 err=0
真机: 两格都是 5（Invalid procedure call or argument）
```

#### 二、改法与 GA 红

`vb6_Asc` 对空串先 `vb6_ErrRaise(5, …)` 再 `return 0`。本机探针两格都变 err=5。
但 GA `4454*****` 上 `[RUN] test_joinfilter: output mismatch`：

```
JF-nullasc=94   （期望）  →  实测 5
```

原因（发码逐字看的）：

```c
BSTR _own_6 = vb6_CStrStrict(nv2);              /* nv2 = Null ⇒ raise 94, 交回空串 */
nc = vb6_VariantInt((int16_t)(vb6_Asc(_own_6)));
```

`Asc(<Null>)` 到 `vb6_Asc` 手上时，94 已经报过了，手上的就是一个空串 ——
**两个来源在 BSTR 层不可分**，在这里再报 5 就把 94 盖掉。
“把 CStrStrict 失败时改交 NULL” 也试不得：`vb6_CStrStrict` 的返回值被下游
`Mid/Left/UCase/Len` 那一片直接当字符串用（C6–C8 就是它们），换 NULL 等于拿空指针去换
一堆未审计的解引用。

#### 三、结论

撤回本刀（`vb6_Asc` 回到“空串静默 0”，仅留注释记下来龙去脉），
`Asc("")=5` **与 W4b/W5b/W6b、C6/C7/C8 同一笔账**：都要“语句级中止”——
转换一 raise，整条语句就停，`vb6_Asc` 根本不会跑，两种来源自然分开。
那一刀落地后，本格只需把 `vb6_ErrRaise(5, …)` 加回去。

教训：RTL 层的“补一个 raise”不是局部改动 —— 它改的是**错误号**，
而上游可能已经报过更具体的错误。凡是不确定“到这一步时 Err 是什么”的格，先写探针把
整条链路（包括上游转换）的 Err 读数拿齐，再决定报不报。








### 11.56 第七十九刀：语句级中止 —— 错误序号 + 赋值站守卫（W4b/W5b/W6b、C6/C7/C8 六格一次收口）

#### 一、这一刀买的是什么

VB6 里 `On Error Resume Next` 中止的是**当前语句**：出错之前已经求值的一切保留，
但**赋值不发生**。C3 原来把「转换失败」之后的那一发赋值照样发出去，于是真机读数是

```
真机:  W4b n = Len(a(0))   → n=-1  err=94      C3 改前: n=0   err=94
       W5b n = Len(FNull())              n=-1            n=0
       W6b n = Len(c(1))                 n=-1            n=0
       C6  s = Mid(Null,1)      → s=[keep] err=94      s=""
       C7  s = Left(Null,1)              [keep]          s=""
       C8  s = UCase(Null)               [keep]          s=""
```

六格的错号早就对了（94），错的是**目标被写成了一个空值**。所以这一刀不改转换、
不改 Err：只加一句「本语句 raise 过就别写目标」。

#### 二、机制（两半，各一枚出口）

1. **RTL 侧：错误序号**（`vb6rtl.c` / `vb6rtl_array.h`）
   `int32_t vb6_err_serial`，`vb6_ErrRaise` 一进门就 `++`；`vb6_ErrSerial()` 读它。
   序号而不是 `Err.Number`：程序自己 `Err.Clear`、或者两次同号错（都是 94）都区分不开，
   「本语句有没有 raise 过」只有单调计数答得准。
2. **发码侧：一句守卫**（`cgen_state.inc` 的 `stmtAbortGuard_` + `emitGuardedStore`）
   赋值站开头（`visit(AssignmentStmt&)` 第一句之后）判一次：目标档位是
   **整型标量**（Integer/Long/Byte）或 **String**，且右值是一枚调用、它的某个实参档是
   `Variant`/`Null`/`Empty`（`Null`/`Empty` 另外直接看 `LiteralExpr::literalKind`，
   因为字面量那一路 `inferExprType` 的档位不稳）。命中就发

   ```c
   const int32_t _e7 = vb6_ErrSerial();
   ...
   if (_e7 == vb6_ErrSerial()) n = vb6_Len(_own_1);
   ```

   `StmtAbortScope`（局部 RAII）在 `visit` 任何出口清空守卫名，所以早期 `return`
   的那几十条支路不会把守卫漏给下一条语句。
   String 目标那一支走 `emitBstrSlotStore` 的两个出口：借来的值只深拷贝 ⇒ 直接不发；
   **自有临时串**（`valueOwned113 || wrapped122`）⇒ 不发就必须自己释放：

   ```c
   if (_e0 == vb6_ErrSerial()) vb6_BSTR_AssignMove(&s, vb6_Mid(_own_4, 1, -1));
   else                        vb6_BSTR_Free(vb6_Mid(_own_4, 1, -1));
   ```

   `vb6_BSTR_Assign` 那一发仍写在原处（R7 名单钉的是「这一站还是手写的」，
   改成 `emitGuardedStore` 会让 R7 数少一处，等于用改门禁来通过改动 —— 不取）。

#### 三、本机读数（`.temp/k77/wc.bas`，改后）

```
W4  s = a(0)          | s=[keep] err=94     W4b n = Len(a(0))     | n=-1 err=94
W5  s = FNull()       | s=[keep] err=94     W5b n = Len(FNull())  | n=-1 err=94
W6  s = c(1)          | s=[keep] err=94     W6b n = Len(c(1))     | n=-1 err=94
C6  s = Mid(Null,1)   | s=[keep] err=94
C7  s = Left(Null,1)  | s=[keep] err=94
C8  s = UCase(Null)   | s=[keep] err=94
```

九格全中真机口径（W4/W5/W6 上一刀已经对，这一刀没动它们）。

#### 四、门禁

- `-Category syntax`：PASS=227 FAIL=0。
- `-Category vbp -VbpShard 9 -VbpShardTotal 12`：PASS=56 FAIL=1，红的仍是本机已知
  `[VBP] c29listview`（同一 HEAD 的 CI 是 PASS，OCX/编译器可得性的本机环境差）。
- 四道静态门：`builtin_slot_returns` / `conv_strict_parse` / `variant_slot_store` /
  `bstr_borrow_slot` 全绿（R7 三次试跑后回到 3 处手抄）。
- 发码清单：78 行哈希变化，逐行看过差异 —— 只有两种新增
  `const int32_t _eN = vb6_ErrSerial();` 与包在 `if (_eN == vb6_ErrSerial())` 里的
  原赋值行（`else` 那半只有自有临时串的释放）；对应 78 份输入重新登记。

#### 五、为什么守卫发得比想象的多（523 处），以及为什么这不是白名单

判据是「右值是调用 + 某个实参档落在 Variant/Null/Empty」，而 `inferExprType` 的
**兜底答案就是 Variant**（见 `cgen_file_io.cpp` 那段注释），于是 `x = Foo(未知形参)`
这类语句也吃到守卫。语义上这是对的：VB6 里**任何**错都会中止当前语句，不只 94；
守卫的代价只是多一行声明和多一次整数比较，而被它挡住的一定是「真 raise 过」的语句。
反过来把判据收窄成「只认 94 那几族」才是白名单思路 —— 错号一换就漏。

#### 六、这一刀没做的（写清楚，别当成已完成）

- `n = Asc("")` / `n = Asc(Empty)` 仍读 `n=0 err=0`（真机 5）。机制已经就位：
  `vb6_Asc` 里补一句「空串 ⇒ `vb6_ErrRaise(5, …)`」不会再盖掉上游 94，因为
  `Asc(<Null>)` 那一路的上游转换一 raise，整条语句的赋值就不发了 —— 但**Asc 自己
  仍然会跑**（RHS 求值不会停），它照样会把 Err 从 94 改成 5。要分开得让
  `vb6_Asc` 只在「本语句还没 raise 过」时报 5，也就是把序号带进 RTL 的读侧 ——
  下一刀再说，别在这里猜。
- `Len` 两档（数值实参 ⇒ Integer，字符串 ⇒ Long）仍是 backlog：会动
  `run_tests.ps1` 的形状针（5809、8811–8813）。
- 定长串目标（`udtFixedField` / `fixedLen`）那两站没有守卫：它们的写出口是
  `vb6_FixedStrAssign*` / `vb6_LSet*`，这一刀没碰。

### 11.57 GA 44565667912 红的两格：登记时把「机器侧 5 行」换成了本机数 + 一条 runner 争用

#### 一、形状门（真错，已修）

CI 判词只有 5 行：`tests/guisample/GUISample.vbp`、`…/vb6ref/src/vc6/GUISample.vbp`、
`…/vs2015/GUISample.vbp`、`tests/TeeChartChartDisplay/DisplayProject.vbp`、
`tests/ctrlmscomm/MsCommApp.vbp`。这 5 份输入在 CI 与本机**本来就不同答**
（引用本机注册过的类型库/OCX 的那一档），所以清单里它们一直存的是 **CI 那台机器给的值**。

第七十九刀让它们的发码也变了（多出同一批守卫），`rebless_emit_manifest.ps1 -Bless`
只会用**手上这份清单**改写，于是把它们一起换成了本机数 —— 745 行对，5 行不对。

修法：从这次 run 的 `emit-manifest` 工件（`actions/artifacts/14551349391`）取 CI 自己那份，
只把 5 行按路径换回去，其余 745 行一个字不动；换完拿 CI 那份清单喂
`compare_emit_manifest.ps1` ⇒ 750/750 相同。

教训（补进清单纪律）：**-Bless 只许在没有机器侧行的改动上用**。凡本轮动到那 5 份输入，
登记后必须再用 CI 的清单复算一遍，不能只看本机 `相同 750`。
（本机现在跑 compare 会报这 5 行不同 —— 那是**本机口径差**，不是回归；
判红只能用 CI 自己那台二进制交回来的清单。）

#### 二、Tests (vbp #5) 红：`[PIXELS] uctreeview_px.ps1`（本机同一 HEAD 复跑 PASS）

CI 读数：`'Tree Explorer' 内容段: BM_CLICK 3s 未返回 + 迟到 7s 仍无窗体 (resp=False)`，
同一条详情的后半段是 `jitter[Corporate Hierarchy]=BM_CLICK(late=4954ms resp=True)`
—— 后一个按钮**迟到了 4.9 秒才响应**，说明那台机器当时被挤得很厉害（3 秒同步窗是硬判据）。

本机复现（同一 HEAD、同一 C3.exe）：手工按 `Test-VbpBuild` 的口径重建
`tests/ucTreeView/Proyecto1.vbp --arch x86`（拷 `img\` + `ejemplo.xml`/`JerarquiaEmpresa.xml`/`menu.xml`），
再跑 `tests/guisample/uctreeview_px.ps1`：**rc=0 PASS**，五段内容全开、3 轮重复全开，
一次 `BM_CLICK` 超时都没有（`fg[...] ok=True try=1`，各 `repeat[...] open sync= close=True`）。

所以归因=**runner 争用**（该 run 同时跑了 21 个 job），不是这一刀的发码。处置：不改判据、
不改超时窗，重跑；若同一格再红，再回头看是不是夹具有真的时序耦合。

#### 三、留给下一次的读法

`-Category vbp -VbpShard 5` 在本机这台机器上**不会重建** `output\build_ucTreeView`
（那目录 2026-11-10 就在，说明那一格在本机走的是缓存/跳过），所以「本机分片绿」不能当
「这一格跑过」的凭证 —— 要点名复跑像素夹具时，照上面那三行手工重建 + 单跑脚本。

### 11.58 第八十刀：`Len(<数值变量>)` 的档位收成 Integer（两档分开）

#### 一、真机口径（`.temp/k78/lena.bas`，用户真机采）

```
D1  lv = Len(<Integer>) | val=2 tn=Integer
D2  lv = Len(<Long>)    | val=4 tn=Integer
D3  lv = Len("abc")     | val=3 tn=Long
D4  lv = Len(Empty)     | val=0 tn=Long
```

值一直是 C3 早就对的（sizeof），错的是**档**：C3 原来四格全答 Long。
为什么数值那两格是 Integer 而不是 Len 声明的 Long：`Len(<数值变量>)` 在 VB6 里是
**编译期常量**（存储宽度），而 VB6 把装得下的常量定为 Integer —— 与「Len 返回 Long」
不矛盾，两档必须分开走。

#### 二、改法（两半，缺一不可）

1. **发码**（`cgen_expr_call_builtin_pre.inc`）：数值支的 `((int32_t)sizeof(x))` 收成
   `((int16_t)sizeof(x))` —— 三处：数组元素、标量数值、非字符串兜底桶。
   UDT 那一支仍是 int32_t（长度会超 32767），定长串那一支也不动（那是 String 族）。
2. **推理**（`cgen_util_type.cpp` 的 `inferExprType`，IndexOrCallExpr 分支开头）：
   `Len(<数值档实参>)` 答 **Integer**；字符串 / Variant / 对象 / 未知一概不拦（仍答 Long）。
   只改这一问，`builtin_funcs.inc` 那张表（Len = Long）不动 —— 表答的是「声明」，
   而这里答的是「这一枚常量表达式的档」，两句话并存。

光改发码不够：装箱由**推理档**决定，实测改前发 `vb6_VariantLong(((int16_t)sizeof(i1)))`
⇒ TypeName 仍答 Long（D1 正确之前那一版就是这样）。两处都改后装箱落
`vb6_VariantFromValue(((int16_t)sizeof(i1)))` 的 `_Generic` Integer 档。

#### 三、读数（改后，本机）

```
D1 val=2 tn=Integer    D2 val=4 tn=Integer
D3 val=3 tn=Long       D4 val=0 tn=Long
W4b/W5b/W6b n=-1 err=94，C6/C7/C8 keep err=94（第七十九刀的九格未被本刀动到）
```

#### 四、门禁与本机读数

- `-Category syntax`：PASS=227 FAIL=0。
- `-Category vbp -VbpShard 9`：PASS=56 FAIL=1（仍只有本机已知 `[VBP] c29listview`）。
- 四道静态门全绿。
- 形状针同步：`run_tests.ps1:5809`（b174 的 `Len(n)`）与 `8811–8813`
  （`len_emitc_width` 的 gL/gI/gB）从 `int32_t` 改成 `int16_t`；
  `8848` 的定长串 `gT` 与 `5810` 的 UDT **不动**（那两档没变，正是本刀要守住的边界）。
- 发码清单：11 份输入变化 —— 除 `test_len_width` / `test_memdim` / `b174` / `dbgdlg` /
  `NullFamilyProbe` 五份本机可复算的以外，另外 5 份是**机器侧行**
  （guisample×3 / TeeChart / ctrlmscomm）。它们的 CI 散列只能吃 CI 那台二进制交回来的数：
  先按本机数登记（闸会红在那 5 行），再从那次 run 的 `emit-manifest` 工件把 5 行换回去
  —— 与 §11.57 同一条流程，多跑一趟 GA 是这一步的固有成本。

#### 五、下一步

`n = Asc("")` ⇒ 5 仍留着（§11.56 末）：要 `vb6_Asc` 只在「本语句还没 raise 过」时报 5，
即把错误序号带进 RTL 读侧。


### 11.59 第八十一刀：`Asc("")` / `Asc(Empty)` ⇒ 5（目标不动，且不吞掉上游的 94）

#### 一、真机口径

`.temp/k81/asc.bas`（本机探针）与 `.temp/k78/lena.bas` 的 A/L 两组：

| 格 | 源 | 真机 |
|---|---|---|
| A1 | `n = Asc(<Null>)` | `n` 不动，`Err.Number=94` |
| A2 | `n = Asc("")` | `n` 不动，`Err.Number=5` |
| L1 | `n = Asc("")`（`Err.Clear` 之后） | `n` 不动，5 |
| L2 | `n = Asc(Empty)` | `n` 不动，5 |

第七十八刀在 RTL 里直接 `ErrRaiseNumber(5)` 那一版被撤回（§11.55），原因是
`Asc(<Null>)` 到 `vb6_Asc` 手上时**已经**被上游 `vb6_CStrStrict` 报过 94、交回的只是一只
空串 —— 两个来路在 BSTR 层长得一模一样，于是 5 把 94 盖掉。本刀把这两条路分开。

#### 二、三处改动（缺一不可）

1. **RTL 报 5**：`vb6_Asc`（`vb6rtl_string.c:359`）空串/NULL ⇒ `vb6_ErrRaiseNumber(5); return 0;`。
2. **赋值站的语句级守卫覆盖 Asc 族**（`cgen_assign.cpp`）：判据里加 `asc/ascb/ascw`。
   但第七十九刀那形「先比序号、再执行 store」**挡不住它** —— `Asc` 自己 raise 发生在
   store 那一行的表达式内部，号是在 store 之后才变的。所以整型标量目标改走**临时量形**：

   ```c
   const int32_t _e1 = vb6_ErrSerial();
   int32_t _av2 = (_e1 == vb6_ErrSerial()) ? (vb6_Asc(_own_2)) : (int32_t)0;
   if (_e1 == vb6_ErrSerial()) n = _av2;
   ```

   两条性质都在这一形里：① 右值先算进临时量（raise 之后就能判）；② 三元的第一支才真跑
   `vb6_Asc` —— 上游转换已经报过 94 时，这里的调用**根本不发生**，5 也就没机会盖 94。
   （第一版少了那个三元，实测 `A1` 直接答 `err=5` —— 是本刀当场量出来的红，不是推测。）
   `cgen_state.inc` 为此加 `stmtAbortTmp_` / `stmtAbortTmpType_`（`StmtAbortScope` 一并清）。
3. **目标不是整型标量那一形**：`nc = Asc(nv2)`（`nc` 是 Variant）发不出第 2 条的守卫 ——
   `test_joinfilter` 的 `JF-nullasc` 当场从 94 掉到 5。RTL 侧补一枚确定性判据：
   `vb6_CStrStrict` 报 94 之后 `vb6_ConvFailMark()`（记下当时的错误序号），`vb6_Asc` 见空串
   且该序号**仍等于**当前序号 ⇒ 判定「这只是转换失败交回来的」，不再报 5；
   `vb6_ErrClear` 把标记清掉（那是「这次错误已被程序处理完」的唯一信号）。
   `.temp/manifest` 的读侧不受影响：`vb6_err_serial` / `vb6_conv_fail_serial` 声明进
   `vb6rtl_array.h`（`check_rtl_proto_arity` 已把它算进 1780 对签名）。

#### 三、读数

- `A1 Asc(Null) n=-1 err=94`、`A2 Asc("") n=-1 err=5`、`L1/L2 = n=-1 err=5`。
- `D1–D4`（第八十刀的 Len 两档）与 `W4..W6b` / `C6..C8` 九格**一字未动**。
- `test_joinfilter`：`JF-nullasc=94`（另 20 余格 null* 不动）。

#### 四、门

- `-Category syntax`：227/227。
- `-Category vbp -VbpShard 9 -VbpShardTotal 12`：55 PASS / 2 FAIL。
  `[VBP] c29listview` 是本机已知非回归红；`[COMPILE-FAIL] iface_multi_incomplete_negcontrol`
  在同一 HEAD 手工复跑得 `VB3012` + rc=1 + 无 exe（正是该针要的），归 runner 争用。
- 静态门 8 枚全 PASS：`variant_slot_store`（R7=3 保住）、`conv_strict_parse`（cells=4）、
  `builtin_slot_returns`、`rtl_naked_names`、`rtl_proto_arity`、`owned_temp_keeps`、
  `implicit_cstr_station`、`emitted_declared`。
- 发码清单：6 行变化（`tests/Charts 2020/Proyecto1.vbp`、同目录 `ucTreeMaps/Proyecto1.vbp`、
  `tests/dbgdlg/dbgdlg.vbp`、`tests/test_str_cnum_assign.bas`、`tests/VBFlexGridDemo/Common/Common.bas`、
  `tests/VBFlexGridDemo/VBFlexGridDemo.vbp`），逐份看过：只多出守卫那两三行（+121..+328 字节）。
  同批把第八十刀按本机数登记的 5 行机器侧差（guisample×3 / TeeChart / ctrlmscomm）
  **换回 CI run `44631595167` 的工件值**，本机复算于是只剩这 5 行红（预期，CI 侧 750/750）。

#### 五、遗留

`String` 目标的形（`s = CStr(Asc(...))`）仍走第七十九刀那套 `if (...) AssignMove else Free`
的**双算**：`Asc` 会在 else 分支再跑一次（`Err.Number` 仍是 5、目标仍不动，只是多发一次）。
它与 `s = CStr(UBound(...))` 同形同口径，本刀不扩，留给「String 目标也吃临时量」那一刀。

### 11.60 第八十二刀：`AscB("")` / `AscW("")` / `AscW(Empty)` ⇒ 5（Asc 同族的尾巴）

#### 一、改前读数（同一台 `.build/C3.exe`，改前）

`.temp/k82/ascb.bas`：

| 格 | 源 | 改前 | 真机 |
|---|---|---|---|
| B1 | `w = AscW("")` | `w=0 err=0` | `w` 不动，5 |
| B2 | `b = AscB("")` | `b=0 err=0` | `b` 不动，5 |
| B3 | `w = AscW(Empty)` | `w=0 err=0` | `w` 不动，5 |
| B4 | `b = AscB(<Null>)` | `b=-1 err=94` | 94（本已对） |
| B5 | `AscW("A")` / `AscB("A")` | `65 / 65 err=0` | 同 |

归因：`vb6_AscW` / `vb6_AscB`（`vb6rtl_compat.c:31/41`）对空串是**静默 return 0** ——
第八十一刀只修了 `vb6_Asc`，这两枚同族函数各走各的身子，所以尾巴还在。
`B4` 之所以早已对，是上游 `vb6_CStrStrict(Null)` 先报了 94，这里只是没去盖它。

#### 二、改法

两枚函数都按 `vb6_Asc` 的同一口径补上：

```c
if (!s || vb6_BSTR_Len(s) == 0) {
    if (vb6_conv_fail_serial >= 0 && vb6_conv_fail_serial == vb6_err_serial) return 0;
    vb6_ErrRaiseNumber(5);
    return 0;
}
```

「转换失败交回来的空串不报 5」那一句必须照抄 —— 少了它，`B4` 会从 94 掉成 5
（第七十八刀在 `vb6_Asc` 上撤回的正是这个坑）。

**发码一侧不需要动**：第八十一刀已把 `ascb/ascw` 纳入赋值站的守卫判据，整型标量目标
走临时量三元形，所以「报 5 之后目标不写」这一半本刀免费得到 —— 读数里 B1/B2/B3 的
`w`/`b` 都停在 -1 就是这个守卫在起作用（改前是 0，因为根本没 raise）。

#### 三、改后读数

```
B1 AscW(empty)  | w=-1 n=-1 err=5
B2 AscB(empty)  | b=-1 err=5
B3 AscW(Empty)  | w=-1 err=5
B4 AscB(Null)   | b=-1 err=94
B5 AscW/AscB(A) | w=65 b=65 err=0
```

非回归：`.temp/k81/asc.bas`（`A1` 94 / `A2` 5）、`.temp/k78/lena.bas`（`L1/L2` 5、
`D1–D4` Len 两档一字未动）、`.temp/k77/wc.bas` 九格全不变。

#### 四、门

- `-Category syntax`：227/227。
- `-Category vbp -VbpShard 9 -VbpShardTotal 12`：55 PASS / 2 FAIL —— 与第八十一刀同一对
  （`c29listview` 本机已知非回归红；`iface_multi_incomplete_negcontrol` 归一 HEAD 手工复跑
  得 VB3012 + rc=1 的 runner 争用）。
- 静态门 6 枚全 PASS：`rtl_naked_names`、`rtl_proto_arity`、`conv_strict_parse`、
  `builtin_slot_returns`、`rtl_embedded`、`emitted_declared`。
- 发码清单：**一行未动**（750 行里只有那 5 行机器侧差，与上一刀登记后的本机状态相同）
  ⇒ 本刀不需要第二个「登记」提交。RTL 只在链接期进产物，不改 `.c` 文本这一层。

### 11.61 第八十三刀：String 目标的守卫真中止 —— 右值只算一次，且算完再判

#### 一、病（第七十九刀留下的两半）

第七十九刀给「可能 raise 的赋值语句」发了守卫，但 String 目标那一支的形状是
`if (<序号没变>) <写/释放>(<右值>); else <释放>(<右值>);`：

1. **条件是求值之前比的** —— 右值自己 raise（`Asc("")` 报 5、实参位的 `vb6_CStrStrict(<Null>)`
   报 94）时条件早就过了 ⇒ 赋值照样发生，目标被清成 `""`，而真机是「报错 + 目标不动」。
2. **else 那一支把右值又算了一遍** —— 中止之后还会再触发一次右值的副作用。

#### 二、改法（两处）

1. `emitBstrSlotStore`（`cgen_util_ctrl.cpp`）两攮同形：右值先算进 `BSTR _av<N>`，
   只在**本语句还没 raise 时**算（三元的第一支）；然后拿**算完之后**的序号决定
   `AssignMove`（自有）还是 `Free`（自有且中止），借来的值只 `Assign`（不回收）。
2. `cgen_assign.cpp` 补一条**只认 Asc 族**的树扫描（`ascCallInTree`）：String 目标的
   `s = CStr(Asc(""))` 这一形外层是转换函数、实参档是 Integer，顶层判据看不见它。
   Asc/AscB/AscW 的 raise 发生在**本语句的表达式求值里**，不会被被调方内部
   `On Error Resume Next` 吞掉 ⇒ 「序号一变 = 真中止」对这三枚成立。

⚠ **撤回的那一版不在这刀里**：递归扫「右值里**任何**嵌套调用」实测在 `[VBP] fdrawstate_x86`
上翻红（成因：分不清「错误往外传」与「被调方内部已吃掉」两种序号自增）⇒ 留作另刀，
要宽到那一层得先给 RTL 一枚「错误是否还在往外爬」的信号。

#### 三、读数（`.temp/k83/stra.bas`，x64，改前 `C3_base.exe` vs 改后 `.build/C3.exe`）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| T1 | `s = CStr(Asc(""))` | `s=0 err=5` | `s=keep err=5` | keep, 5 |
| T2 | `s = FBoom(Null)`（被调方内部报 94） | `s= err=94` | `s=keep err=94` | keep, 94 |
| T3 | `s = Trim(Null)` | `s=keep err=94` | `s=keep err=94` | keep, 94 |
| T4 | `n = Asc("")`（整型守卫） | `n=-1 err=5` | `n=-1 err=5` | -1, 5 |
| T5 | `s = CStr(FSeven(nv))` | `s=7 cnt=1` | `s=7 cnt=1` | 7, 1 |
| T6 | `s = "A" & "B"` | `s=AB` | `s=AB` | AB |

T1/T2 是本刀的行为面，T4/T5/T6 是零回归护栏。

#### 四、门

- `-Category syntax`：**227/227**。
- `-Category vbp -VbpShard 9 -VbpShardTotal 12`：54 PASS / 3 FAIL / 1 SKIP。三红逐条：
  `iface_multi_incomplete_negcontrol` = 本机已知 runner 争用（同一 HEAD 手工复跑得 VB3012 + rc=1）；
  `b448_arg_station_returns_transferred_ref` / `b156_local_carrier_destroy_shapes` = 该轮把 stdout
  重定向到文件后中文 needle 的解码差（**工件不是读数**：同一份 `.build/C3.exe` 手工复现两条 needle
  全在场，见 `.temp/k83/chk_shape.ps1` 的 `present=True`）⇒ 不是发码回归。
- 静态门 9 枚全 PASS：`rtl_naked_names`、`rtl_proto_arity`（1780 对）、`conv_strict_parse`、
  `builtin_slot_returns`、`emitted_declared`、`manifest_coverage`（K1..K3 inputs=750）、
  `static_sentinel_registration`（148）、`test_helper_integrity`、`rtl_embedded`。
- 发码清单：见下节「登记」。

#### 五、登记（发码清单）

本刀只改**带语句级守卫的 String 目标**那几行不字，所以候选集不从全语料猜，而是两源取并：
① 第七十九刀 / 第八十一刀登记过的 78 行（凡是带守卫的行都在里面）；② 语料里出现 `Asc/AscB/AscW`
调用的输入（本机 7 份，含它们的 `.vbp`）—— 新守卫只可能从这两个源里长出来。

逐份重发后实际变动 **23 行**（`git diff --numstat` = 23+/23-）：`cls_inh/Inh.vbp`、`dbgdlg/dbgdlg.vbp`、
`geomcache/GCCache.vbp`、`test_builtin_arity.bas`、`test_splitjoin.bas`、`test_str_cnum_assign.bas`、
`VBFlexGridDemo/VBFlexGridDemo.vbp`、`VbQRCodegen-master/src/mdQRCodegen.bas`、
`VbQRCodegen-master/test/Project1.vbp`、`b181_baformal.bas`、`b187_box/B187d.vbp`、`b189_gate/B189g.vbp`、
`b190_scope/B190.vbp`、`b190_scope/B190Main.bas`、`b433/B433.vbp`、`b433/B433Main.bas`、
`k481_sat/K481Sat.vbp`、`b598_print_numeric/B598.vbp`、`b598_print_numeric/B598Main.bas`、
`ucTreeView/Proyecto1.vbp`、`GifcontrolSource/prjNewAniGif.vbp`、`vb6probe/NullOpsProbe.bas`、
`vb6probe/NullFamilyProbe.bas`。差异形状只有两类：① 自有支从
`if (ok) AssignMove(&s, <值>) else Free(<值>)` 变成 `BSTR _avN = (ok) ? (<值>) : NULL;` +
`if (ok) AssignMove(&s, _avN) else Free(_avN)`；② 借用支从 `if (ok) Assign(&s, <值>)` 变成同样先算进
`_avN` 再 `Assign`。**不登记**的 5 行机器侧差（`guisample` 三份 `GUISample.vbp` + `TeeChartChartDisplay`
+ `ctrlmscomm`）在重算时被回退成 HEAD（CI 工件）值 —— 本机 14.29 与 CI 14.52 的发码在这 5 份上天生不同，
不能拿本机数 bless。

⚠ 本机未能跑完整语料 `emit_manifest`（`Start-Process -Wait` 在该会话里逐份派生子进程后挂住，
两次都在中途被杀），所以上面用「候选集 + 逐份重发」代替全量：候选集的两源已覆盖本刀全部触发面
（守卫存在 + Asc 族），且重发比的是与 CI 同一份归一化哈希（`// Generated by vb6c3` 行归位 +
根路径替换 + ascii256 双列）。下一步全量对齐交给 CI 的形状门。

#### 六、CI 回执（run 44675509943）

12 个 vbp 分片里 2 片红，两类读数，都不是语义回归：

- `b189_gate_same_name_slot_pair` / `b190_typename_of_object_formal_stays_runtime` 两枚
  `[EMITC-SHAPE]` 针钉的是旧形 `AssignMove(&目标, <右值>)`，被本刀的临时量行挤开。
  针按新形拆两截重钉：`? (<右值>) : NULL;`（右值原样在场 = 不许折常量，正是这两条的靶子）
  + `AssignMove(&目标, _av`（目标吃那枚临时量）。本机按 `Test-EmitcShape` 同一口径复算：
  两例 PASS，`missing=[] odd=0`。
- 同片另有几行 `[GUI] ... pixels: PASS ...` 被标注机制打成 `::error::`（读数以 png 路径收尾，
  不是 PASS/SKIP 收尾）——`Results: FAIL=1` 里不含它们，属注解口径噪声，不是用例红。

### 11.62 第八十四刀：整型目标的嵌套 Asc 形也吃临时量（`n = CLng(Asc(""))` 一族）

#### 一、病

第八十一刀只认**顶层就是** Asc 的形（`n = Asc("")`）。右值外套一层转换函数或运算符时，
顶层判据看不见它，语句照样把 0 写进目标：

- `n = CLng(Asc(""))`
- `n = Asc("") + 0`
- `n2 = Len(CStr(Asc("")))`

真机三种都是「报 5 + 目标不动」。第八十三刀把树扫描扩到了 String 目标，整型这一半当时留着
（`ascInRhs83` 前面挂了 `tgtT79 == String`）。

#### 二、改法

1. `ascInRhs83` → `ascInRhs84`，判据从「String 目标」扩到 `tgtInt79 || tgtT79 == String`。
2. 整型目标必须**显式**打开 `stmtAbortTmp_` / `stmtAbortTmpType_`：`emitGuardedStore` 只有打开它才发
   `int32_t _avN = (<守卫> == vb6_ErrSerial()) ? (<右值>) : (int32_t)0;` + `if (...) t = _avN;`；
   否则只发 `if (ok) t = <右值>;` —— 而序号是 store 那一行才变的，挡不住右值自己 raise。
   这一档**复用第八十一刀的整型临时量出口**，不新造形状。

#### 三、读数（`.temp/k84/inttarget.bas`，x64，改前 = 第八十三刀那枚 `.build/C3.exe`）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| T1 | `n = CLng(Asc(""))` | `n=0 err=5` | `n=-1 err=5` | -1, 5 |
| T2 | `n = Asc("") + 0` | `n=0 err=5` | `n=-1 err=5` | -1, 5 |
| T3 | `n = CLng(Asc("A"))` | `n=65 err=0` | `n=65 err=0` | 65, 0 |
| T4 | `n = Asc("")`（第八十一刀那格） | `n=-1 err=5` | `n=-1 err=5` | -1, 5 |
| T5 | `n2 = Len(CStr(Asc("")))` | `n2=1 err=5` | `n2=-1 err=5` | -1, 5 |
| T6 | `s = CStr(Asc(""))`（第八十三刀那格） | `s=keep err=5` | `s=keep err=5` | keep, 5 |

#### 四、门

- `-Category syntax`：**227/227**。
- 静态门 9 枚全 PASS（含 `manifest_coverage` 750/750、`rtl_proto_arity` 1780 对）。
- 本机 `-Category vbp -VbpShard 9`：55 PASS / 2 FAIL / 1 SKIP；两红是本机已知
  （`c29listview` 本机 OCX、`iface_multi_incomplete_negcontrol` 争用），同一 HEAD 上 CI 是 PASS。
- `run_tests.ps1` 里**没有任何形状针引用 `vb6_Asc*`** ⇒ 本刀不动针。

#### 五、登记（发码清单）

候选集不从全语料猜：本刀的触发面 = 右值树里有 Asc 族且目标是整型标量/String，所以取**清单里
源文件正文出现 `Asc(` / `AscB(` / `AscW(` 的 7 份**逐份重发。实际变动 **2 行**：
`tests/VbQRCodegen-master/src/mdQRCodegen.bas` 与 `tests/VbQRCodegen-master/test/Project1.vbp`。
发码逐字看过（`.temp/k84/md_old.c` vs `md_new.c`）：两处 `lDigit = vb6_ChkInt((int32_t)(vb6_Asc(_own_N)) - 48);`
变成临时量形，同一文件后面两处 `_e17/_e18` 顺移到 `_e19/_e20`（守卫序号按语句顺序发，不重排）。

⚠ 第一次只登记了 1 行，CI 的形状门回来打脸：`tests/VbQRCodegen-master/test/Project1.vbp` 也是红。
原因是我那份候选集只 grep 了 `.vbp` **文件自己的正文** —— 工程的模块（`mdQRCodegen.bas`）不在里面，
所以含 Asc 的工程整份漏掉了。下次扫工程得把它的模块一起看，别只读 `.vbp` 本体。

### 11.63 第八十五刀：除零也是「本语句内 raise」—— `n = 10 \ 0` 不再把 0 写进目标

#### 一、病

探针 `.temp/k85/raisefam.bas` 把「不是 Asc 族、但本语句内会 raise」的形扫了一遍，结果分两类：

1. **错误已经报了，目标还是被写**（本刀的对象）：
   `n = 10 \ 0` → `n=0 err=11`、`s = CStr(10 \ 0)` → `s=0 err=11`。真机都是「11 + 目标不动」。
2. **RTL 压根没报错**（另一条经，本刀不碰）：`Trim(CStr(Null))`、`Len(CStr(Null))` 读 `err=0`；
   `Mid("abc", 0, 1)`、`Chr(-1)` 也都不报错。那是「内置函数自己该 raise 而没 raise」的缺口，
   与语句级中止不是一个轴 —— 已记在待办里，别当成这刀做完了。

#### 二、改法

`ascCallInTree` → `raiseInTree85`，认两族：① Asc/AscB/AscW（第八十三/八十四刀）；
② `Div` / `IntDiv` / `Mod` 三枚除法算符。
**除数能静态看出是非零字面量时不算**（`l = 7 / 2` 不可能 raise）—— 这一条是补上的：
第一版不分字面量，语法门里 `f2lng_round` 直接红（它钉的 `l = vb6_ChkLong(...vb6_Num_Div((double)(7), (double)(2)))`
被套上了临时量）。整型目标走 `stmtAbortTmp_`，String 目标走 `emitBstrSlotStore` 的借用/自有两支，
与第八十三/八十四刀同一套出口，没有新形状。

#### 三、读数（`.temp/k85/raisefam.bas`，x64，改前 = 第八十四刀那枚二进制）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| T7 | `n = 10 \ 0` | `n=0 err=11` | `n=-1 err=11` | -1, 11 |
| T8 | `s = CStr(10 \ 0)` | `s=0 err=11` | `s=keep err=11` | keep, 11 |
| T1–T6 | `Trim(CStr(Null))` / `Mid(...,0)` / `Chr(-1)` 族 | `err=0` | `err=0`（未动） | 94 / 5 |

#### 四、门

- `-Category syntax`：**227/227**（第一版宽松判据下曾 226/227，红的是 `f2lng_round`；
  补「非零字面量不发守卫」后恢复）。
- 静态门 9 枚全 PASS（含 `manifest_coverage` 750/750、`rtl_proto_arity` 1780 对）。
- 本机没跑 vbp 分片，交给 CI —— 形状针里只有 `f2lng_round` 一处引用除法发码，已在语法门里验过。

#### 五、登记（发码清单）

本刀的触发面是「右值里有除法算符」——语料里到处都是，候选集法不再够用，所以这刀**本机全量重算**：
`.temp/k85/full_rehash85.ps1` 把 750 行逐份重发（跳过 5 行机器侧差），归一化口径与 CI 同源。
净变 **17 行**（宽松版曾 28 行，补字面量判据后有 9 行回到原值）。

### 11.64 第八十六刀：`Chr(-1)` / `Left(s,-1)` / `Mid(s,0)` 该报 5 却静默

#### 一、病（`.temp/k86/builtin.bas`，x64）

| 格 | 源 | 改前 | 真机 |
|---|---|---|---|
| T2 | `s = Chr(-1)` | `s=\uFFFF err=0` | 5，s 不动 |
| T3 | `s = Chr(70000)` | `s=<乱码> err=0` | 5，s 不动 |
| T5 | `s = Mid("abc", 0, 1)` | `s="" err=0` | 5，s 不动 |
| T7 | `s = Left("abc", -1)` | `s="" err=0` | 5，s 不动 |
| T8 | `s = Right("abc", -1)` | `s="" err=0` | 5，s 不动 |
| T9 | `n = InStr(0, "abc", "b")` | `n=2 err=0` | 5，n 不动（**本刀不做**） |

T1(`CStr(Null)`⇒94)、T4(`Chr(65)`)、T6(`Mid(...,2,1)`)、T10 本来就对，作护栏。

#### 二、改法

1. RTL（`vb6rtl_string.c`）：`vb6_Chr` 码点越界（<0 或 >65535）⇒ 5；`vb6_Left`/`vb6_Right`
   （含 B 族）负长度 ⇒ 5，0 仍是空串；`vb6_Mid`/`vb6_MidB` 起点 < 1 ⇒ 5（`len < 0` 仍是 cgen
   对「省略 length」的哨兵，不当错误）。
2. 发射侧：`raiseInTree85` 的白名单从 Asc 族扩到 `chr/chrw/chrb`、`left/leftb`、`right/rightb`、
   `mid/midb`（含 `$` 后缀）。**不扩的话 RTL 报 5 之后目标还是被写成空串** —— raise 与中止
   是两半，缺一不可。

#### 三、读数（改后）

T2/T3/T5/T7/T8 全部变成 `keep / 5`；T1/T4/T6/T9/T10 一格未动。

#### 四、门

- `-Category syntax`：**227/227**（白名单扩了 10 个名字，形状针一枚没红）。
- 静态门 9 枚全 PASS（含 `rtl_embedded` 138 枚资源与磁盘一致、`rtl_proto_arity` 1780 对）。

#### 五、登记（发码清单）

白名单一类改动的触发面比「除法」还宽（`Left/Mid/Right` 满地都是），继续全量重算 750 行：
净变 **23 行**（跳过 5 行机器侧差）。

#### 六、本刀没做的

- `InStr(<起点<1>, …)` 仍 clamp 到 1（真机报 5）—— 它不在 `VBA.Strings` 的取子串族里，
  单独一刀再动（要同时看 `InStrB`/`InStrRev` 的口径）。
- 定长串目标（`udtFixedField`/`fixedLen`）仍无守卫；`Len` 两档仍是 backlog。

#### 七、CI 回执（run 44699767997，红两份）

- `[RUN] test_joinfilter`：`JF-nullmid=5`（期望 94）。成因与本刀同形：`Mid("abc", <Null>)` 的
  起点走 `vb6_VariantToLong` 报 94，随后 Mid 自己的 5 又把它盖了。修法：`vb6_VariantToLong`
  的 Null 分支也打 `vb6_ConvFailMark()`，并新增 `vb6_ArgConvertFailed()`（标记仍等于当前序号 =
  本语句刚有一次实参转换失败）—— Mid/Left/Right/Chr（含 B 族）raise 前先问它，是就不报 5。
- `[EMITC-SHAPE] k3060_set_variant_shape`：`bt = CByte("&H" & Mid(hx,1,2))` 的右值里有 Mid，
  现在走整型临时量形 —— 针按前几刀的办法拆两截重钉。
- 本机复现：`test_joinfilter` 87/87 行全对（`missing=0 extra=0`）、K3060 四条针全在场；
  `-Category syntax` 仍 227/227；静态门 9 枚全 PASS。

### 11.65 第八十七刀：`InStr(0,…)` / `InStrRev(…,0)` 该报 5 却把起点悄悄夹成 1

#### 一、病（`.temp/k87/instr.bas`，x64，改前 = 第八十六刀那枚二进制）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| T1 | `n = InStr(0, "abc", "b")` | `n=2 err=0` | `n=-1 err=5` | -1, 5 |
| T2 | `n = InStr(-1, "abc", "b")` | `n=2 err=0` | `n=-1 err=5` | -1, 5 |
| T5 | `n = InStr(0, "abc", "")` | `n=0 err=0` | `n=-1 err=5` | -1, 5 |
| T6 | `n = InStrRev("abc", "b", 0)` | `n=2 err=0` | `n=-1 err=5` | -1, 5 |
| T8 | `n = InStrRev("abc", "b", -2)` | `n=2 err=0` | `n=-1 err=5` | -1, 5 |
| T10 | `n = InStrB(0, a, b)` | `n=2 err=0` | `n=-1 err=5` | 5 |

对照格（T3 `InStr(1,…)`⇒2、T4 `InStr(4,…)`⇒0、T7 `InStrRev(-1)`⇒2、T9 省略⇒2、T11 `InStrB(1,…)`⇒2）
改后逐字不变；T12（String 目标 `s = CStr(InStr(0,…))`）⇒ `keep/5`，T13（起点是 Null）⇒ `n=-1 err=94`
（起点转换的 94 不被 5 盖掉，走第八十六刀那枚 `vb6_ArgConvertFailed`）。

真机读数是 `.temp/k87/oracle_instr.vbs`（VBScript 的 `InStr`/`InStrRev` 与 VB6 同一套 VBA 字符串运行时；
`InStrB` 只有 VB6 有，按 InStr 同口径推）。

#### 二、改法

1. RTL（`vb6rtl_string.c`）：
   - `vb6_InStrC` 起点 < 1 ⇒ 5（`return 0`）。**校验顺序按真机**：先校起点，再判空 needle
     —— 真机 `InStr(0,"abc","")` 也是 5，不是 0。
   - `vb6_InStrB`（Byte 数组形与 BSTR 形公用入口）同一道门。
   - `vb6_InStrRev`：`-1` 是「从尾部」的合法档（等同省略）⇒ 夹成 `hLen`；其余 < 1（0、-2…）⇒ 5。
2. 发射侧：`raiseInTree85` 白名单再加 `instr/instrb/instrrev`。**raise 与中止是两半**：
   不登记的话 RTL 报了 5，`n = …` 照样把目标写成 0/2。

#### 三、登记（发码清单）

触发面同为「右值里有内建调用」，仍走本机全量重发（`.temp/k85/full_rehash85.ps1`，750 行、跳过 5 行机器侧差）：
净变 **13 行**（`ucSegmentedControl`、`ucSlider_2022` 等含 `InStr` 赋值的模块）。

#### 四、门

- `-Category syntax`：**227/227**。
- 静态门 9 枚全 PASS（`manifest_coverage` 750/750、`rtl_proto_arity` 1781 对、
  `rtl_embedded` 138 份字节一致）。

#### 五、本刀没做的

- `InStrRev` 的 `start > hLen`：真机答 0（oracle I 格），C3 仍夹成 `hLen` 后照样查找。
  这是**另一条**口径（不是本刀的 < 1），单列 backlog。
- `InStrB` 的「起点是 Null」以及 Byte 数组的越界档没有真机对照（VBScript 无 InStrB），只按 InStr 同口径推。
- 定长串目标（`udtFixedField`/`fixedLen`）、`Len` 两档仍在 backlog。

### 11.66 第八十八刀：`InStr` 的空 needle / `InStrRev` 的 `start > Len` 答错

#### 一、病（`.temp/k88/edge.bas`，x64，改前 = 第八十七刀那枚二进制）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| E1 | `InStr(4, "abc", "")` | `4` | `0` | 0 |
| E2 | `InStr(3, "abc", "")` | `3` | `3` | 3 |
| E4 | `InStrRev("abc", "", 2)` | `3` | `2` | 2 |
| E5 | `InStrRev("abc", "", 4)` | `3` | `0` | 0 |
| E6 | `InStrRev("abc", "b", 4)` | `2` | `0` | 0 |
| E7 | `InStrRev("abc", "b", 99)` | `2` | `0` | 0 |
| E9 | `InStrRev("", "", 1)` | `0` | `0` | 0 |
| E3/E8/E10 | `InStr(1,"abc","")` / `InStrRev("abc","b",3)` / `InStrRev("abc","")` | 1 / 2 / 3 | 不变 | 1 / 2 / 3 |

真机读数：`.temp/k88/oracle_instr_edge.vbs`（E1/E2）、`oracle_rev2.vbs`（rev 4/99/空 needle）、
`oracle_rev3.vbs`（`InStrRev("abc","",2)`⇒2、`("abc","",1)`⇒1 —— 空 needle 交回的是**起点**，不是 Len）。

#### 二、改法（纯 RTL，不动发射侧）

`vb6rtl_string.c`：

1. `vb6_InStrC`：空 needle 交回 `start`，但 `start > Len` 时交 0
   （`return (start > vb6_BSTR_Len(haystack)) ? 0 : start;`）。旧写法无条件 `return start`。
2. `vb6_InStrB`：Byte 数组形与 BSTR 形同口径（无真机对照，按 InStr 同族推）。
3. `vb6_InStrRev`：
   - 空 needle ⇒ `return start`（旧写法交 `hLen`）；
   - `start > hLen` ⇒ **先于空 needle** 判 `return 0`（旧写法夹成 `hLen` 继续查）。

#### 三、读数（`.temp/k88/edge.bas` 改后）

E1–E10 逐格与真机一致（见上表「改后」列）。第八十七刀的 13 格探针
（`.temp/k87/instr.bas`，`start < 1` 那一族）在改后二进制上逐字不变。

#### 四、门

- `-Category syntax`：**227/227**。
- 静态门 9 枚全 PASS。**发码清单不变**（本刀只改 RTL 行为，不改任何语句的发射形状）。
- `check_owned_temp_keeps`（CI 的 `[STATIC]` 那一格）：首轮红一枚 `~vb6_InStrC|1`。
  归因：空 needle 分支多读一次 `vb6_BSTR_Len(haystack)`，表里那一行的「代表读目标」
  于是从 `vb6_TextFind` 换成 `vb6_BSTR_Len`；指针轴仍是 `keeps_pointer=0`（形参不外逃），
  与快照逐列对过后 `-Bless` 登记（净变 1 行）。

#### 五、本刀没做的

- `InStrB` 的空 needle / 越界两格没有真机对照（VBScript 无 InStrB）。
- 定长串目标（`udtFixedField`/`fixedLen`）、`Len` 两档仍在 backlog。

### 11.67 第八十九刀：定长串目标（UDT `String * N` 字段 / 模块级 `String * N`）没有语句级中止

#### 一、病（`.temp/k89/fixed.bas`，x64，改前 = 第八十八刀那枚二进制）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| T1 | `u.f = Left("abc", -1)`（UDT 定长字段） | `f=[     ] err=5` | `f=[keep ] err=5` | keep, 5 |
| T2 | `g = Left("abc", -1)`（模块级 `String * 5`） | `g=[     ] err=5` | `g=[keep ] err=5` | keep, 5 |
| T5 | `u.f = Chr(-1)` | `f=[     ] err=5` | `f=[keep ] err=5` | keep, 5 |
| T6 | `g = Chr(-1)` | `g=[     ] err=5` | `g=[keep ] err=5` | keep, 5 |
| T3/T4 | 对照 `"hi"` | `hi   ` | 不变 | `hi   ` |
| T9/T10 | 对照 截断/补空格 | `abcde` / `ab   ` len=5 | 不变 | 同 |

成因：守卫**已经发了**（`const int32_t _e0 = vb6_ErrSerial();`），但定长串那两条发码支根本不看它 ——
`vb6_FixedStrAssignFree(u.f, 5, vb6_Left(...))` / `vb6_BSTR_AssignMove(&g, vb6_LSetFree(..., 5))`
照写，于是「报 5 + 目标被抹成空格」。普通 BSTR 槽（第七十九/八十三刀）早就两半齐全，定长串这一档漏了。

#### 二、改法（`emitBstrSlotStore`，`cgen_util_ctrl.cpp`）

两条定长串分支各补一道与普通槽**同构**的门：右值**只算一次**进临时量，再拿算完之后的序号决定写不写。

- UDT 定长字段：`BSTR _avN = (guard == vb6_ErrSerial()) ? (rhs) : NULL;` +
  `if (…) vb6_FixedStrAssignFree/Assign(u.f, N, _avN); [else vb6_BSTR_Free(_avN);]`
  （自有串那一支要 else 收回，借来的没有可回收的东西）
- 模块级定长串：`BSTR _avN = (…) ? (vb6_LSet(Free)(rhs, N)) : NULL;` +
  `if (…) vb6_BSTR_AssignMove(&g, _avN); else vb6_BSTR_Free(_avN);`

#### 三、登记（发码清单）

全量重算 750 行，净变 **1 行**（`tests/Charts 2020/ucTreeMaps/Proyecto1.vbp`）——
语料里「定长串目标 + 会 raise 的右值」本来就少。

#### 四、门

- `-Category syntax`：**227/227**。
- 静态门 9 枚全 PASS；`check_owned_temp_keeps` PASS（本刀不动 RTL，K5 无漂移）。

#### 五、本刀没做的

- T7/T8：`u.f = CStr(Null)` / `g = CStr(Null)` 改前改后都是 `[     ] err=0` —— 定长串目标上
  Variant → String 仍走**宽松**的 `vb6_CStr`（Null 折成空串），raise 压根没发生，所以守卫也无从生效。
  真机是 `94 + 目标不动`。这是**转换口径**那一刀，另列 backlog。
- `Len` 两档仍在 backlog。

### 11.68 第九十刀：定长串目标上的 `94`（Variant → String 走的是宽松口）

#### 一、病（`.temp/k90/strict.bas` + `.temp/k89/fixed.bas`，x64，改前 = 第八十九刀那枚二进制）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| T1 | `u.f = CStr(Null)`（UDT `String * 5` 字段） | `f=[     ] err=0` | `f=[keep ] err=94` | keep, 94 |
| T2 | `g = CStr(Null)`（模块级 `String * 5`） | `g=[     ] err=0` | `g=[keep ] err=94` | keep, 94 |
| T3 | `u.f = v`（v = Null） | `f=[     ] err=0` | `f=[keep ] err=94` | keep, 94 |
| T4 | `g = v`（v = Null） | `g=[     ] err=0` | `g=[keep ] err=94` | keep, 94 |
| T5 | `v = 42; u.f = v / g = v` | `42   `（err=0） | `42   ` | `42   ` |
| T7 | `v = "abc"; u.f = v` | `abc  ` | `abc  ` | `abc  ` |
| T8 | `v = 7; u.f = CStr(v) / g = CStr(v)` | `7    ` | `7    ` | `7    ` |
| T9 | `v = "abcdefg"; u.f = v` | `abcde` | `abcde` | `abcde` |

成因：严格赋值站（第七十三刀 `valueIsVariant` / 第七十四刀 `s = CStr(<Variant>)`）两处都带
`fixedLen.empty()` 的前提，定长串目标因此整档绕开 —— 走的是宽松 `vb6_CStr` / `vb6_VariantToString`
（`Null` 折成空串），所以 **raise 压根没发生**，第八十九刀那道中止门也就无从生效。

#### 二、改法

1. RTL（`vb6rtl_string.c` + `vb6rtl_builtin.h`）新增两枚**原子**出口，口径与
   `vb6_BSTR_AssignFromVariant` 一致（Null ⇒ 94、目标一字不动）：
   - `vb6_FixedStrAssignFromVariant(char* buf, int32_t cap, vb6_VARIANT src)` —— UDT 内联 `char[N]`；
   - `vb6_LSetFromVariant(BSTR* target, int32_t cap, vb6_VARIANT src)` —— 模块级 `String * N`（BSTR 模型）。
   两枚内部都是「Null 先拦 → 否则 `vb6_CStr` 折串 → 既有定长写侧（右补空格 / 右截断）→ 收回折串临时」。
2. 发射侧（`emitBstrSlotStore`）：① 第七十四刀那条 `CStr(<Variant>)` 识别去掉 `fixedLen.empty()`
   前提；② 严格赋值站之后补一条定长串分支，按 UDT 字段 / 裸槽分别发上面两枚出口。
   守卫那行（`_eN`）照旧发，但这两枚出口自己是原子的（raise 即返回），所以中止不依赖它。

#### 三、读数

`.temp/k90/strict.bas` 九格全对（见上表）；第八十九刀的 `.temp/k89/fixed.bas` 十三格复跑：
T7/T8 由 `err=0` 变 `err=94`（目标仍 `keep`），T1/T2/T5/T6 仍是 `keep/5`，对照格不变。

#### 四、登记

- 发码清单全量重算 750 行，净变 **1 行**（`tests/GifcontrolSource/prjNewAniGif.vbp`）。
- `check_owned_temp_keeps`：新增两枚 RTL 出口 ⇒ 快照 +2 行
  （`+vb6_FixedStrAssignFromVariant|2`、`+vb6_LSetFromVariant|2`，两行都是 `keeps_pointer=0`、
  `handoff-reads:vb6_CStr`）—— 归因后 `-Bless`；`rtl_proto_arity` 由 1781 变 1783 对。

#### 五、门

- `-Category syntax`：**227/227**；静态门 9 枚全 PASS。

#### 六、本刀没做的

- `Len` 两档仍在 backlog。

### 11.69 第九十一刀：`Len` / `LenB` 收 `Null` 的位置（94 落到整型目标那一瞬）与 `LenB` 缺变体位

#### 一、真机口径（`.temp/k91/oracle_len.vbs`，VBScript／同一套 VBA 串运行期）

```
A Len(v=Null)      val=(Null) tn=Null  err=0     ← Len 本身把 Null 原样交回
B Len(Null)        val=(Null) tn=Null  err=0
C Len(Empty)=0     D Len("abc")=3    E Len(42)=2    F Len(v="hello")=5    G Len(True)=4
```

⇒ `Len(<Null>)` **不 raise**，答 **Null**；落到整型目标那一瞬才报 **94**、目标不动
（与第七十九刀那批 `W4b/W5b/W6b` 同一条口径）。`LenB` 那半没有 VBScript 入口，
按手册同一句（「Variant 一律按 String 看」）与 `Len` 同族推：`n = LenB(v)`（v=Null）也是 **94 + 目标不动**。

#### 二、病（`.temp/k91/len.bas`，x64，改前 = 第九十刀那枚二进制）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| T1 | `n = Len(v)`（v = Null） | `n=0 err=94` | **`n=-1 err=94`** | 目标不动, 94 |
| T12 | `n = LenB(v)`（v = Null） | `n=0 err=0` | **`n=-1 err=94`** | 同 T1 |
| T2 | `vv = Len(v)`（v = Null，Variant 目标） | `isnull=True vt=1` | 同（本来就对） | Null |
| T3–T7 | `Len(v="hello")` / `Len("abc")` / `Len(模块级 String * 5)` / `Len(u.f)` / `LenB(模块级 String * 5)` | 5 / 3 / 5 / 5 / 10 | 同 | 同 |
| T11 | `Len(v=Empty)` | 0 | 0 | 0 |
| T13 | `LenB("")` | 0 | 0 | 0 |

两处成因分得很开：

1. **T1 是「守卫发了、临时量没发」**。第七十九刀那条判据（右值是一枚调用、某实参静态类型是
   Variant）本来就看得到这一形，于是发了 `const int32_t _e0 = vb6_ErrSerial();`
   —— 但守卫形是 `if (_e0 == vb6_ErrSerial()) n = <RHS>;`，序号在**条件求值那一步**还没变，
   RHS 里的 `vb6_VariantToLong(vb6_LenV(v))` 就在条件之后把 94 报出来、语句照样把 0 写进目标。
   只有 Asc 族（第八十一/八十三刀）和除零（第八十五刀）打开了 `stmtAbortTmp_`，Len 这一族一直没打开。
2. **T12 是 `LenB` 压根没有变体位**。`vb6_LenV` 是第七十一刀给 `Len` 造的（Null 进 Null 出），
   `LenB` 没有对应出口 ⇒ 实参落到运行时签名表那一站，按 `BSTR` 折成
   `vb6_LenB(vb6_VariantToString(v))`（宽松折串，Null → 空串）⇒ 静默 0 / Err 0。

#### 三、改法（三半，缺一不可）

1. **发码侧打开临时量**（`stmt/cgen_assign.cpp`，在原来那个 `lit79 || 变体档` 的循环里）：
   整型标量目标 + 本刀的两枚（`len` / `lenb`）+ 实参是**真的整只 Variant 槽** ⇒ `stmtAbortTmp_`。
   ⚠ 两个**不许放宽**的方向，都用 750 行清单量过：
   - 不许按 `inferExprType` 认变体档：它对 UDT / 数组 / 属性读一律回退成 Variant ⇒
     `LenB(<UDT>)`（发的是 `sizeof`）、`UBound(arr, 1)` 这些**不可能 raise** 的右值全被套上 `_av`
     （第一版就是这么放的，当场 **54 行**变动）。改成与发码站同一问 `isDefinitelyVariantExpr` + 不用 `isArr` 那档。
   - 也不许一宽到「所有内建 + 整型目标」：那样仍有 **25 行**，其中 `n = VarType(v)`（根本不会 raise）与
     `n = UBound(vb6_VariantToSafeArray1DStrict(v), 1)` 这类**口径没量过**的右值也被套上 `_av`。
     本刀的判据只做到这两枚有真机读数的地方 —— 那一族见第六节的待办。
2. **RTL 补 `vb6_LenBV`**（`vb6rtl_string.c` + `vb6rtl_builtin.h`）：Null ⇒ Null；否则
   `vb6_LenB_BSTR(vb6_VariantToString(v))` 并把折出来的那枚串就地放掉。**不许**走 `vb6_LenB` 那个
   `_Generic` 宏 —— 它对非 BSTR 实参走 `default: sizeof(x)`，拿它答 Variant 槽会给 24/20。
3. **两处改道口都要认它**：`nullFamilyVExitCName` 表 +`{"lenb","vb6_LenBV"}`；
   并且 `cgen_expr_call_arg_variant.inc` 里 `rtParamType == "BSTR"` 那一站也要问一遍
   （`LenB` 在**运行时 C 签名表**里那一格写的是宏的第一臂，不是内置函数的形参档 ⇒ 它不走
   `paramBase == String` 那条改道口；只改表不改这一站，实测发出来还是 `vb6_LenB(VariantToString(v))`）。
   `scripts/gen_variant_slot_prefixes.ps1` 重跑：60 → **61 名**（`+ vb6_LenBV`）。

#### 四、读数（改后，本机 x64）

```
T1 n=-1 err=94    T2 isnull=True vt=1 err=0    T3 n=5    T4 n=3    T5 n=5
T6 n=5            T7 n=10                       T8 n=10   T11 n=0   T12 n=-1 err=94   T13 n=0
```

#### 五、登记

- 发码清单全量重算 750 行（跳过机器侧 5 行），净变 **3 行**：`tests/vb6probe/NullOpsProbe.bas`、
  `tests/test_joinfilter.bas`、`tests/VBFlexGridDemo/VBFlexGridDemo.vbp` —— 三行都是同一形状
  `if (_eN == serial) i = …(LenV(槽));` → `int32_t _avN = (_eN == serial) ? (…) : 0; if (_eN == serial) i = _avN;`。
- `check_owned_temp_keeps`：新增一枚 RTL 出口 ⇒ ① 手写名单 `src/common/owned_temp_keeps.hpp`
  两处 +`vb6_LenBV`（K13/K13c 双向对账的那一侧，**脚本不许改写**）；② 推导快照 `-Bless` **+1 行**
  （`+vb6_LenBV|0`）。`rtl_proto_arity` 1783 → 1784 对。

#### 六、本刀没做的（都已量过、登在这里）

1. **`LenB(<UDT 定长字段>)` 答 10，真机口径未量**：UDT 字段在本仓是内联 `char[N]`（ANSI），
   发的是 `vb6_LenB(vb6_FixedStrToBSTR(u.f, 5))` = 宽字符 ×2 = 10；真机是 5 还是 10 **没有读数**
   （VBScript 不支持定长串，采不到），本刀按「不猜」留原样，等真机采一次再定。
2. **`n = CLng(v)` / `CInt(v)` / `UBound(v)` 这一族**（Variant 实参 + 整型目标）同样有「94/13 报出来、
   目标被写 0」的形状，本机 750 行清单的宽度已量到 **25 行**；本刀刻意只做到 Len/LenB，
   泛化留给下一刀（判据要能分开「会 raise 的转换」与 `VarType` 这种不会 raise 的）。
