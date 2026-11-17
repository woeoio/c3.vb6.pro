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

