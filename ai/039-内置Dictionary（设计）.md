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


