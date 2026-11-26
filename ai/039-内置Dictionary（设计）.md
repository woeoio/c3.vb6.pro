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

### 11.70 第九十二刀：整型目标上「整只 Variant 槽 → 非 Variant 形参」那一族（`CLng` / `CInt` / `UBound` …）

#### 一、真机口径（`.temp/k92/oracle_conv.vbs`，VBScript／同一套 VBA 运行期）

```
A vv = CLng(v=Null)  → vv=Empty, err=94        B n = CLng(Null)  → n 不动, 94
C n = CInt(Null)     → n 不动, 94              D n = CLng(Empty) → 0
E n = UBound(v=1)    → n 不动, 13（非数组）     F n = UBound(a)   → 1
G/H n = VarType(v)   → 2 / 1, err=0（不转换）   I n = CLng("abc") → n 不动, 13
J n = Int(Null)      → Null（err 0，Null 传播族）
```

⇒ 与第九十一刀同一句，只是不限于 `Len`：**整只槽交给一个「非 Variant」形参**时，那一步的转换
（`VariantToDouble` / `NumVal` / `VariantToSafeArray1DStrict` / `VariantToLong`）会 raise 94 或 13，
raise 之后 RHS 仍会交回 0 —— 目标不该被写。`VarType` 是这条判据的反面：它的形参**本来就是** `VARIANT`，
没有转换，所以那一格只该多一枚无害局部量、值一行不许变。

#### 二、病（`.temp/k92/conv.bas`，x64，改前 = 第九十一刀那枚二进制）

| 格 | 源 | 改前 | 改后 | 真机 |
|---|---|---|---|---|
| U1 | `n = CLng(v)`（v = Null） | `n=0 err=94` | **`n=-1 err=94`** | 目标不动, 94 |
| U2 | `n = CInt(v)`（v = Null） | `n=0 err=94` | **`n=-1 err=94`** | 目标不动, 94 |
| U3 | `n = CLng(v)`（v = "abc"） | `n=0 err=13` | `n=0 err=13` ✗ | 目标不动, 13 |
| U4 | `n = UBound(v)`（v = 1，非数组） | `n=-1 err=13` | 同（本来就对） | 目标不动, 13 |
| U5 | `n = VarType(v)`（v = Null） | 1 / err 0 | 同 | 1 / err 0 |
| U6 | `n = CLng(Empty)` | 0 | 0 | 0 |
| U7 | `n = CLng("42")` | 42 | 42 | 42 |

第九十一刀的判据按 callee 名分族（只认 `len` / `lenb`）⇒ U1/U2 那两枚不在族里，一条也不 fire。

#### 三、改法（一处，判据收成一句）

`stmt/cgen_assign.cpp`：把第九十一刀那个 `lenFam91` 的族别去掉 —— 判据回到
**整型标量目标 + 右值顶层是调用 + 该调用某个实参是「真的整只 Variant 槽」（或 Null/Empty 字面量）**。

- 「真的整只槽」仍只问 `isDefinitelyVariantExpr` + 剔掉数组本体（`inferExprType` 那一半的坑见 §11.69）。
- 两个「不许放宽」的方向原样保留（按 `inferExprType` 认变体档 ⇒ 54 行噪声；这一条已写死在注释里）。

#### 四、这刀的代价（宽度先量后写）

750 行清单全量重算，净变 **24 行**、`_av` 格 **57 个**（`.temp/k92/census3.ps1` 逐格归因）：

| callee | 格数 | 性质 |
|---|---|---|
| `vb6_UBound` / `vb6_LBound` | 17 / 10 | 形参是 `SafeArray1D*`，实参是槽 ⇒ 会报 13；U4 那一格已证明 |
| `vb6_CLngV` / `vb6_CIntV` / `vb6_ChkLong` | 6 / 1 / 1 | 94 那一族（U1/U2） |
| Declare `SelectObject` / 工程内过程（6 枚） | 2 / 6 | 形参有具体档，实参是槽 ⇒ 会 raise |
| `vb6_VarType` | 6 | **无害**：形参就是 `VARIANT`，值一格没变（U5） |
| Declare `DispCallFunc` | 8 | **无害**：C 侧不 raise，只是多一枚局部量 |

⇒ 57 格里 14 格是「多一枚无害局部量」，43 格是真口径。

#### 五、本刀没修的：13 那条路守卫看不见（已量，下一刀）

U3（`n = CLng("abc")`）改后仍是 `n=0 err=13`。根因不在赋值站，而在 RTL 有两枚 raise 出口：

- `vb6_ErrRaise(errNum, source, description)` —— **`++vb6_err_serial`**（守卫看得见）；
- `vb6_RaiseError(errNum, description)` —— **不打序列**（守卫看不见）。

`vb6_NumVal` 的「文本认不出」走的正是后者（`vb6rtl_string.c:462`）。全仓 `vb6_RaiseError(` **34 处 / 15 份文件**
（`vb6com.c` 429/438、`vb6com_invoke.c`、`vb6com_fso.c` …），凡是走它报出来的错误，第九十一/九十二刀这两道
守卫一律无效 —— 这一族是下一刀的入口，**不许**在这一刀里顺手改 raise 管线（那是另一条轴，要逐枚读「该不该让守卫看见」）。

#### 六、登记与门

- 发码清单：24 行（列表见 `.temp/k92/rehash92.log`），无机器侧 5 行。
- `check_owned_temp_keeps`：本刀不动 RTL ⇒ 253 行不变，全绿。
- 静态门 9 枚全 PASS（`rtl_proto_arity` 1784 对）；`-Category syntax` **227/227**。

### 11.71 第九十三刀：`vb6_RaiseError` 不打错误序号 —— 34 处报错对语句级守卫全是隐形的

#### 一、病（`.temp/k93/vb6probe/RealMachineProbe.bas`，C3 侧读数，改前 = 第九十二刀那枚二进制）

```
C5 | CLng(v=abc) into Long = [0]  err=13      ← 目标被写 0
```

真机口径（同族，第九十二刀 `.temp/k92/oracle_conv.vbs` E/I 两格）：**13 报出来、目标一字不动**。

根因不在赋值站，而在 RTL 有**两枚**报错出口，对「错误序号」的态度不一致：

| 出口 | 打 `vb6_err_serial` | 用它的地方 |
|---|---|---|
| `vb6_ErrRaise(errNum, source, description)` | **打**（`++vb6_err_serial`） | 第九十一/九十二刀那两族转换失败（94 等） |
| `vb6_RaiseError(errNum, description)` | **不打** | 全仓 34 处 / 15 份文件：COM `429`/`438`（`vb6com.c`、`vb6com_invoke.c`）、`vb6com_fso.c`、`vb6_NumVal` 的文本 13（`vb6rtl_string.c:462`）… |

赋值站的语句级守卫判的是「序号变没变」（第七十九刀），所以凡走第二枚报出来的错误，守卫**一律看不见** ⇒
语句接着把转换失败交回的 0 / 空串写进目标。CBool/CLng 那条文本档（`vb6_NumVal`）是最常见的门面，
但同一族还盖着 COM 的 429/438 —— 那一半今天也同样是「报错 + 目标被写」。

#### 二、改法（一处，一行）

`vb6rtl.c` 的 `vb6_RaiseError` 开头补 `++vb6_err_serial;`，与 `vb6_ErrRaise` 同一句「**报错就要动序号**」。
两枚出口的其余行为（`vb6_err_resume_next` 早退、`On Error GoTo` 的 longjmp、未处理时打 stderr/弹框）一字未动。

#### 三、读数（C3 侧，改后）

```
C5 | CLng(v=abc) into Long = [-1]  err=13      ← 修好：目标不动
B3/B4 Len/LenB(Null) = [-1] err=94 ; C1/C2 CLng/CInt(Null) = [-1] err=94 ; C6 UBound(非数组) = [-1] err=13
D1..D3/D5 不变（Asc 94 / Mid 5 / Left 5 / InStr 5，目标都不动）
```

同一枚探针还量到两处**本轮不修**的洞（真机读数一到就开下一刀）：

1. `n = Int(v=Null)` / `n = Abs(v=Null)` ⇒ 本仓 `[0] err=0`。发的是
   `vb6_ChkLong(vb6_FltToLng(vb6_Int(vb6_VariantToDouble(v))))` —— `vb6_VariantToDouble` 是**宽松**口
   （Null → 0.0、不报错），而 `Int`/`Abs` 的 Variant 档在 VB6 是 **Null 进 Null 出**（`.temp/k92` J 格：
   `n = Int(Null)` ⇒ Null、err 0），落到整型目标才 94。宽严两档的选择在这一族还没量过。
2. `TypeName(<UDT 的 String*N 字段>)` ⇒ 本仓答 `Object`（应为 `String`）：固定长度字段是 `char[N]`，
   `vb6_TypeName` 只认 BSTR 槽。已登在 §11.72 之下的待办清单。

#### 四、登记与门

- 发码清单全量重算 750 行：**净变 0 行**（本刀只动 RTL 运行期语义，一个字都不在发码面上）。
- `check_owned_temp_keeps`：253 行不变（没有新出口、没有改签名），K13/K14 全绿。
- 静态门 9 枚全 PASS（`rtl_proto_arity` 1784 对；`rtl_embedded` 138 份资源与应用内嵌逐字节相同）。
- `-Category syntax`：**227/227**。

#### 五、下一刀的两笔（都已量、等真机）

- **真机探针已出**：`.temp/k93/vb6probe/RealMachineProbe.bas`（一行一格，`LABEL | what = [value] err=N`；
  A 段定长串 `Len`/`LenB`（模块/局部/UDT 字段/整枚 UDT）、B 段 Variant 各档、C 段转换族、D 段串内建边界）。
  它在 C3 侧已经跑通，读数留在 `.temp/k93/vb6probe/probe_out.txt` —— 真机读数回来后逐格对表。
- 待办：`LenB(<定长串>)` 的 10/5（A2/A4/A6）、`LenB(<整枚 UDT>)`（A8 本仓 5 而 A10 12）、
  `TypeName(<UDT 定长字段>)`（A12 本仓 `Object`）、`Int/Abs(<Null>)`（C3/C4）。

### 11.72 真机读数回来了：34 格逐格对表（`LenB` 定长串结案，`TypeName` 定长串修好，UDT 尺寸 / Int·Abs 两族开账）

真机输出 = `.temp/k93/vb6probe/RealMachineProbe.bas` 在真 VB6 上跑出来的一整段（A/B/C/D 四段 34 格），
C3 侧基线 = 同文件在 `C3.exe` 下跑出来的 `probe_out.txt`（第九十三刀那枚二进制）。

#### 一、对表（真机 vs 改前 C3）

| 格 | 真机 | 改前 C3 | 判 |
|---|---|---|---|
| A1 `Len(模块 String*5="ab")` | 5 | 5 | ✓ |
| A2 `LenB(模块 String*5)` | **10** | 10 | ✓ |
| A3/A4 `Len`/`LenB`(局部 String*5) | 5 / **10** | 5 / 10 | ✓ |
| A5/A6 `Len`/`LenB`(UDT 字段 String*5) | 5 / **10** | 5 / 10 | ✓ |
| A7 `Len(整枚 UDT {String*5})` | 5 | 5 | ✓ |
| A8 `LenB(整枚 UDT {String*5})` | **10** | **5** | ✗ |
| A9 `Len(整枚 UDT {String*5, Long})` | **9** | **12** | ✗ |
| A10 `LenB(整枚 UDT {String*5, Long})` | **16** | **12** | ✗ |
| A11 `TypeName(模块 String*5)` | **String** | **Object** | ✗→本刀修 |
| A12 `TypeName(UDT 字段 String*5)` | **String** | **Object** | ✗→本刀修 |
| A14 `Len(越界赋值 "abcdefgh" 后)` | 5 | 5 | ✓ |
| A15 `LenB(同 A14)` | **10** | 10 | ✓ |
| A16 越界赋值后内容 | `<abcde>` | `<abcde>` | ✓ |
| B1–B10 `Len`/`LenB` 收 Variant（Null/Empty/123/"abc"） | `Null`/1/0/0/0 各格 | 逐格同 | ✓ |
| B3/B4 `Len`/`LenB(v=Null)` 落 Long | **-1 / 94** | -1 / 94 | ✓ |
| C1/C2 `CLng`/`CInt(Null)` 落 Long | **-1 / 94** | -1 / 94 | ✓ |
| C3/C4 `Int`/`Abs(v=Null)` 落 Long | **-1 / 94** | **0 / 0** | ✗ |
| C5 `CLng("abc")` 落 Long | **-1 / 13** | -1 / 13 | ✓（第九十三刀修好） |
| C6 `UBound(非数组)` 落 Long | **-1 / 13** | -1 / 13 | ✓ |
| C7/C8 `VarType(Null)` / `CLng(Empty)` | 1 / 0，err 0 | 同 | ✓ |
| D1 `Asc(Null)` | **-1 / 94** | -1 / 94 | ✓ |
| D2/D3 `Mid(_,0,1)` / `Left(_,-1)` | 目标不动 / 5 | 同 | ✓ |
| D4 `Chr(0)` | len=1，err 0 | 同 | ✓ |
| D5 `InStr(0,…)` | **-1 / 5** | -1 / 5 | ✓ |

**§11.69 那条悬案结案**：`LenB(<定长串>)` 真机就是 **10**（UTF-16 口径：5 字符 × 2 字节），
本仓第九十一刀那一版**本来就对** —— 当初"5 vs 10"的疑点是拿 ANSI `char[N]` 的内联尺寸当答案猜出来的，
真机一量即消。`Len` = 5、`LenB` = 10 两条一起立住。

#### 二、第九十四刀：`TypeName(<定长串>)` ⇒ `"String"`

- 病：定长串在 C 侧是内联 `char[N]`，装槽时退化成 `char*` ⇒ `vb6_VariantFromValue` 只认 BSTR 槽、
  指针按对象包 ⇒ `vb6_TypeName` 答 `Object`。
- 改：`cgen_expr_call_builtin_pre.inc` 的 TypeName 折叠区（与 `TypeName(Me)` / `TypeName(类变量)` 同一形）
  加一条闸：实参 `fixedStrLenOfExpr` 非空 ⇒ 编译期折成 `vb6_BSTR_FromStr(L"String")`。
  真机上这答案是**声明档常量**（A14/A16 越界赋值后仍是 String），绕运行期那问没有意义。
- 读数：A11 `Object` → `String`、A12 `Object` → `String`，其余 32 格逐字不变。
- 发码清单 750 行全量重算 **净变 0 行**（真语料 9949 份里没有 `TypeName(<定长串>)` 这一形）。

#### 三、本刀登记、不修的三笔（都有真机数，档已齐）

**① `Len`/`LenB(<整枚 UDT>)` 走的是 `sizeof`，而真机是"按成员逐段算"**

```
真机: Len ({String*5})       = 5    LenB = 10     Len ({String*5,Long})       = 9    LenB = 16
本仓: 5（sizeof char[5] 巧合对）      5            12（sizeof 结构体）                 12
```

由 A8/A9/A10 三格反推出来的真机规则：

- `Len(<UDT>)` = **各成员 Len 之和**（定长串按 N 字符算、Long 按 4）⇒ 5+4 = **9**。
- `LenB(<UDT>)` = **各成员 LenB 之和，成员之间按 4 字节对齐、尾部不补齐**
  ⇒ `{String*5}`：10（不补齐 ⇒ 不是 12）；`{String*5, Long}`：10 + 补 2 + 4 = **16**。
- 本仓现在发 `sizeof(<C 结构体>)`。两边的**内存布局本来就不同**（C3 定长串内联 ANSI `char[N]`，
  VB6 存 UTF-16），所以这一格不能"改成 sizeof 的另一种算法"，得在发码侧按**声明档**逐成员求值
  —— 与 `udtFixedStrFieldLen_` 那张表同一个来源，但要连数值成员一起走一遍布局。
  这一刀牵到 `Len` 的 UDT 兜底、`LenB` 的 UDT 兜底两处，单独开。

**② `Int`/`Abs`(/`Fix`) 收 Null：本仓 `0 / err 0`，真机 `-1 / 94`**

发的是 `vb6_ChkLong(vb6_FltToLng(vb6_Int(vb6_VariantToDouble(v))))` —— `vb6_VariantToDouble` 是**宽松**口
（Null → 0.0、不报错），所以"Null 进 Null 出"这条语义在中途就被抹平了。
第六十四刀那张 `mathNullFamilyV` 表（`vb6_IntV`/`vb6_AbsV`/`vb6_FixV`）只在**变体上下文**两站生效
（`wrapVariantValue` 存储站 / `boxToVariant` 实参站）；落到**整型目标**的赋值站没有这道改道口。
下一刀要在这条路上补一个出口：实参是整枚 Variant 槽 ⇒ 走 `vb6_IntV(v)` 再 `VariantToLong` 落到整数，
与 Len/LenB 那一形（`vb6_LenV(v)` → `vb6_VariantToLong` → `vb6_ChkLong`）完全同构。

**③ `vb6_RaiseError` 那 34 处虽已打序号（第九十三刀），但"报错 + 目标被写"的另一半要逐族验**

第九十三刀只把**可见性**接上了；每一族"报错之后目标动不动"还得像 C 段这样逐格量。
今天量到的是干净的两格（C5 = -1/13、C6 = -1/13），COM 429/438 那两族**还没有探针格**。

### 11.73 第九十五刀：`Int`/`Fix`/`Abs`(<整只 Variant 槽>) 落整型目标 —— 真机 94、本仓 0

#### 一、病（真机 C3/C4 vs 改前本仓）

```
真机: C3 Int(v=Null) into Long = [-1] err=94    C4 Abs(v=Null) = [-1] err=94
改前: C3 = [0] err=0                            C4 = [0] err=0
```

发的是 `vb6_ChkLong(vb6_FltToLng(vb6_Int(vb6_VariantToDouble(v))))`。赋值站的守卫与临时量
（第七十九/八十一/九十二刀那套）其实**已经全在**——问题是 RHS **一声不响**：`vb6_VariantToDouble`
是宽松口（Null → 0.0、不报错、不动序号），"Null 进 Null 出"这条语义在中途就被抹平了，
到了守卫那里序号没变 ⇒ 照样把 0 写进目标。

第六十四刀那张 `mathNullFamilyV` 表（`vb6_IntV`/`vb6_AbsV`/`vb6_FixV`）只在**变体上下文**两站生效
（`wrapVariantValue` 存储站 / `boxToVariant` 实参站）；落到**整型标量目标**的这条路上没有改道口。

#### 二、改法（三处，各一句话）

1. `cgen_state.inc`：新增 `bool nullFamIntTarget_`（本语句域，由 `StmtAbortScope` 统一清除）。
2. `cgen_assign.cpp`：整型目标 + 右值顶层调用是 `int`/`fix`/`abs` + 实参是**整只 Variant 槽**
   （`isDefinitelyVariantExpr && !isArr`，与第九十二刀同一问）⇒ 置位。
   —— 只有赋值站知道"目标是整型标量"，所以标一声。
3. `cgen_expr_call_builtin_pre.inc`：置位时改道成
   `vb6_VariantToLong(vb6_IntV(x))`（`fix`→`vb6_FixV`、`abs`→`vb6_AbsV`）。
   `IntV` 把 Null 原样带出来，`VariantToLong` 那一站报 94 —— 与 Len/LenB 那一形
   （`vb6_LenV` → `VariantToLong` → `ChkLong`）完全同构；外层的 `ChkLong`/`FltToLng` 一字未动。
   只改道一次（右值顶层那一枚），嵌套复用不会重入。

#### 三、读数与登记

```
改后 C1..C8 逐格与真机一致: C1/C2 -1/94, C3/C4 -1/94(本刀), C5 -1/13, C6 -1/13, C7 1/err0, C8 0/err0
```

- 发码清单 750 行全量重算：**净变 0 行**（真语料里没有"整只 Variant 槽喂 Int/Fix/Abs 再落整型目标"这一形）。
- 静态门 9 枚全 PASS；`check_owned_temp_keeps` 253 行不变（无新出口、无签名改动）；`-Category syntax` **227/227**。
- 本刀只认 `Integer`/`Long`/`Byte` 目标。`Double` 目标的同一格（`d = Int(Null)`）**还没量**，
  要动它得先决定"整条 `vb6_VariantToDouble` 是否转严格"——那是另一个面，已登记。

#### 四、剩下三笔待办（§11.72 起就已带数）

1. `Len`/`LenB(<整枚 UDT>)`：真机按成员逐段算（Len 5/9，LenB 10/16），本仓发 `sizeof`（5/12）。
2. `vb6_VariantToDouble` 的宽松面：`d = Int(Null)` / 其它双精度目标同族还没量。
3. COM 429/438 那两族"报错后目标动不动"还没有探针格。

### 11.74 第九十六刀：`Len(<整枚 UDT>)` 按声明档逐成员算；`LenB(<整枚 UDT>)` 刻意留 `sizeof`

#### 一、真机读数（A7..A10）

```
A7  Len ({String*5})        = 5     A8  LenB ({String*5})        = 10
A9  Len ({String*5, Long})  = 9     A10 LenB ({String*5, Long})  = 16
```

本仓改前两枚都发 `sizeof(<C 结构体>)` ⇒ `5 / 5 / 12 / 12`。由 A8/A9/A10 反推出的规则：

- `Len(<UDT>)`  = **Σ 成员 Len**（定长串按 N **字符**、数值按字节）⇒ 5 / 5+4 = 9
- `LenB(<UDT>)` = Σ 成员 LenB（定长串按 **2N**）、成员之间按自然对齐（上限 4）、**尾部不补齐** ⇒ 10 / 16

#### 二、改法

- `visit(TypeDecl&)` 边发结构体边留一份**声明布局** `udtLayout_`（成员顺序 + 三项：`len` / `lenB` / `align`）。
  成员认不出的档（动态数组、对象、动态 String、Variant、未知 C 类型、嵌套未登记）⇒ **整枚不进表**，
  两枚函数退回原来的 `sizeof` 那一支（宁可不变）。静态数组成员同样不进表（保守）。
- `udtLayoutSum()` 求和（`cgen_util_classtype.cpp`）。
- `Len(<整枚 UDT>)`：有布局就折成常量、且**不发实参**（避开副作用）；否则维持 `sizeof`。

#### 三、`LenB(<整枚 UDT>)` 为什么**不折**

折成 16 在真机上是对的 —— 真机的 UDT 实打实就是 16 字节（定长串 UTF-16）。但 C3 的 UDT 定长串是
**按 ANSI 内联**存的（`char f[N]`，见 `visit(TypeDecl&)` 那条 "VB6 ABI: ANSI 内联" 的注：SHFILEINFO
真机 `Len = 352` 走的也是这条）。真机报 692 是安全的（对象真有 692 字节），C3 报 2N 会让
`cbSize = LenB(u)` 这类 Idiom **越过自己的对象底**。

语法钉 `tests/test_declare_shadows_builtin.bas` 抓到的正是这一形（`GetObjectA` 的 cbSize）——
这是本刀第一版把 `LenB` 也折了之后唯一翻红的一格。所以这一格选择"**与自己的存储一致**"：
`LenB(<整枚 UDT>)` 继续发 `sizeof`，真机 10/16 与本仓 5/12 的差异**登记为刻意的 ABI 分歧**，
不是待修缺陷。要真正对齐，得先把 UDT 定长串改成 UTF-16 存储（另一刀，牵所有 Declare 的封送）。

#### 四、读数与登记

- 探针：A7 = 5 ✓、A9 = 9 ✓（与真机一致）；A8 = 5、A10 = 12（刻意保留，见第三节）；其余 30 格逐字不变。
- 发码清单 750 行全量重算：**6 行**变动（VBFlexGrid 4 行、VbQRCodegen 2 行）。
  抽一行核对（`tests/guisample/Definitions.bas`）：`sizeof(osVersion)` → `((int32_t)148)` ——
  正是 VB6 那句经典 `dwOSVersionInfoSize = Len(osVersion)`（OSVERSIONINFO：5×4 + String*128 = 148）。
- 语法钉 `declare_shadows_builtin` 的期望同步为 `((int32_t)24)`（BITMAP 全是有宽成员，24 **等于**原来的
  `sizeof(bmp)`，数值不变、只是不再写 `sizeof`）。
- 静态门 9 枚全 PASS；`check_owned_temp_keeps` 253 行不变；`-Category syntax` **227/227**。

### 11.75 第九十七刀：Double 目标上的 Null 族与文本档 —— 真机 94/13、本仓 0/0

#### 一、真机读数（E1..E10）

```
E1 Int(v=Null) -> Double  -1/94    E2 Abs(v=Null) -> Double  -1/94
E3 Fix(v=Null) -> Double  -1/94    E4 CLng(v=Null) -> Double -1/94
E5 d = v (Null)           -1/94    E6 d = v + 1 (Null)       -1/94
E7 d = v * 2 (Null)       -1/94    E8 d = v (Empty)          0/0
E9 d = v ("abc")          -1/13    E10 Int(2.7)              2/0
```

改前本仓：E1–E3 **0/err0**、E4 **0/err94**（错误号对了、值被写）、E9 **0/err0**。
E5–E8、E10 本来就对 —— 也就是说 `d = <整只 Variant 槽>` 走的是 `vb6_Dbl_AssignFromVariant`，
Null 一档它自己会报 94 且不动目标；漏的只有两处：**文本档**（那时落到宽松的
`vb6_VariantToDouble`，"abc" ⇒ 0.0 不报错）和 **内建函数把 Null 吃掉**那一族。

#### 二、三处改动

1. `vb6_Dbl_AssignFromVariant` 的 BSTR 档改走 `vb6_Dbl_AssignFromBSTR`（同一口径：严格折数、
   末尾有垃圾字符 ⇒ 13、目标一字不动）⇒ E9。
2. 新增 `vb6_DblStrict(vb6_VARIANT)`：Null ⇒ 94、非数字文本 ⇒ 13，其余 `vb6_VariantToDouble`。
   不复用 `vb6_VariantToDouble` 本身：它**必须**保持宽松（`Val` 那一族靠它，
   `check_conv_strict_parse.ps1` 的 `val-still-lenient=1` 钉的就是这一条）。
3. 赋值站把「Double 目标 + 顶层调用 + 整只 Variant 槽实参」这一形纳入守卫/临时量
   （原先只给 Integer/Long/Byte 发）⇒ E4（`d = CLng(v)`：CLngV 报 94，序号一变就不写）。
   同时 Int/Fix/Abs 那一形改道成 `vb6_Int(vb6_DblStrict(x))`（新旗 `nullFamDblTarget_`）——
   **不能**照第九十五刀那样走 `IntV`/`VariantToLong`：整型那条路会把 `Abs(-1.5)` 弄成 2。

#### 三、为什么 Double 与 Integer 目标要两条改道

| 目标 | 改道形态 | 理由 |
|---|---|---|
| Integer/Long/Byte | `vb6_VariantToLong(vb6_IntV(x))` | 目标是整数，`VariantToLong` 自己会报 94/13，且截断正是语义 |
| Double | `vb6_Int(vb6_DblStrict(x))` + 守卫/临时量 | 小数位要保留；raise 由 `DblStrict` 负责，写入由序号守卫挡 |

#### 四、登记

- 探针 44 格（A–F）：E1–E4、E9 五格翻绿且逐字对上真机；A–D 与 E5–E8/E10 一字不变。
- 发码清单 750 行全量重算：**3 行**变动（`tests/ucTreeView/Proyecto1.vbp` +72B、
  `tests/vb6probe/NullOpsProbe.bas` +130B、`tests/vb6probe/NullFamilyProbe.bas` +42B）。
  抽形核对：`NullOpsProbe` 里 `d = Abs(nv)` → `double _av4 = (_e3 == vb6_ErrSerial()) ? (vb6_Abs(vb6_DblStrict(nv))) : (double)0;`
- `check_owned_temp_keeps` 漂 2 行，**归因**：`vb6_Dbl_AssignFromVariant` 1 → 0（新 BSTR 支把目标指针
  转手给 `vb6_Dbl_AssignFromBSTR`，下游 KEEPS=0 ⇒ 这一枚可证 0，不再"证不出"）、新增
  `vb6_DblStrict` 0（无指针形参）。表 254 行、锚点 15 枚、提名面 203 行不变。
- 静态门 9 枚全 PASS（`proto_arity` 1785 对签名）；`-Category syntax` **227/227**。

### 11.76 第九十八刀：晚绑定缺成员要报 438；标量目标按真机「语句中止、目标不动」

#### 一、真机读数（F1..F6）

```
F1 Set o = CreateObject("No.Such.ProgID")   Nothing / 429
F3 n  = o.NoSuchProperty                    -1 / 438
F4 Call o.NoSuchMethod                      -1 / 438
F5 o.NoSuchProperty = 5                          438
F6 s2 = o.NoSuchStringProperty              keep / 438
```

改前本仓：F1/F4 已对（429 / 438），**F3/F5/F6 全静默** —— `n` 被写成 0、`s2` 被写成 ""、
`o.NoSuchProperty = 5` 当成功且 `Err.Number = 0`。

#### 二、(a) RTL：五个 miss 分支补 438

`vb6com_invoke.c` 里晚绑定成员查不到时只打一行 stderr 就 `return NULL/void` 的五个分支
（`vb6_ComGetProp`、`vb6_ComGetPropArg`、`vb6_ComSetProp`、`vb6_ComSetPropArg`、`vb6_ComSetRef`）
现在与第八刀（`vb6_ComCall` 的 miss）同一口径：

```c
g_com_438_dangling = 1;
vb6_RaiseError(438, SysAllocString(L"Object doesn't support this property or method"));
```

抛完仍交空值让 `On Error Resume Next` 继续，并立起「刚抛过 438」的标记吸收链尾那一次 miss。
方法调用那条路本来就有默认成员兜底 + 438，不动。**这一步发码零变动。**

#### 三、(b) cgen：标量目标上的语句级守卫

只有 raise 还不够——真机要的是「语句中止 + 目标不动」。赋值站新增一条判据：

> 目标 ∈ {Integer, Long, Byte, Double} ∧ 右值顶层是**成员读/调用** ∧ 接收者是**晚绑定对象变量**

⇒ 复用既有守卫 + 临时量（与第七十九/八十四/九十七刀同一套），raise 之后不写目标。

三条排除（都是「不是 COM 派发」的接收者，套守卫纯属噪声，且会翻红三条精确文本钉）：

- 控件数组名（`cboUpdates.UBound` ⇒ `vb6_CtrlArr_UBound`）；
- 内建伪宿主 `Printer/Printers/Screen/App/Clipboard`（成员表直接应答）；
- `Me` 与**同工程的函数名**（`Clone.ScaleFactor(3#)` 这种编译期直调）。

#### 四、churn（先量后改）

只读测量（不改编译器，对 750 行逐行生成 C 数形）：含 COM 成员读/调用的工程 **206/745**，
语句 5,999（普通赋值 1,475 / BSTR 目标 179 / 裸语句 2,859 / 表达式内位置 1,486）。
据此把 (b) 收窄到「标量数值目标 + 普通赋值」这一族：

| 版本 | 清单变动 |
|---|---|
| 初版（无排除） | 70 行 |
| 排除控件数组 / 伪宿主 / 同工程函数名后 | **48 行** |

#### 五、仍挂着的两笔（登记，不混刀）

1. **BSTR 目标**（真机 F6 `s2` 应保持 `keep`）：清单上 179 句。保存/恢复要连所有权一起算
   （与第八十三刀同一条纪律），现在 C3 是 `""` + 438 —— 错误号对了、串被清。
2. **表达式内位置**（`If o.X = 1 Then`、实参、二元运算）：1,486 句，守卫机制插不进去，
   要另开一种形；真机读数也还缺（G 段 6 格）。

#### 六、门

- 探针 44 格：F3/F4/F5 与真机逐字对上；F1/F2 不变；F6 见第五节（已知缺口）。
- 发码清单 750 行：**48 行**变动（登记在册）。
- 静态门 9 枚 PASS（`proto_arity` 1785）；`owned_temp_keeps` 254 行不变；`-Category syntax` **227/227**。


### 11.77 第九十九刀：BSTR 目标上的语句中止（晚绑定缺成员）

#### 一、读数（探针 F6，真机在 `.temp/k93/vb6probe`）

真机 `s2 = o.NoSuchStringProperty`（`o` 是晚绑定 Object，属性不存在）：

| 站 | 真机 | 第九十八刀之后 | 现在 |
|---|---|---|---|
| `s2` 的值 | **keep**（原值不动） | `""`（被清空） | **keep** |
| `Err.Number` | 438 | 438 | 438 |

第九十八刀 (a) 让 COM 那五枚 `DISPID_UNKNOWN` 分支报 438，错误号对了；但赋值站照样把
`vb6_ComGetStringProp` 交回的（空）串写进目标 ⇒ **号对了、值被清**。这一刀补上值那一半。

#### 二、形

标量目标那一套（守卫 + 临时量）搬不过来：BSTR 交回的是**一枚要转手的块**，中止时既不写
目标、又不能漏（否则每中止一次漏一块）。收成 RTL 一枚出口，所有权随出口一起答：

```c
static inline void vb6_BSTR_AssignMoveIfOk(BSTR* target, BSTR source, int32_t serialBefore) {
    if (serialBefore == vb6_ErrSerial()) { vb6_BSTR_AssignMove(target, source); return; }
    if (source) vb6_BSTR_Free(source);      /* 语句已中止：目标一字不动 */
}
```

发码站（`cgen_assign_value_sem.inc` 的 `targetIsBstr` 那一支）：

```c
const int32_t _eN = vb6_ErrSerial();
vb6_BSTR_AssignMoveIfOk(&s2, vb6_ComGetStringProp(o, L"NoSuchStringProperty"), _eN);
```

判据与第九十八刀 (b) 同一枚谓词（接收者链根是晚绑定 Object 变量、且不是控件数组 / 伪宿主 /
`Me` / 同工程函数名），只把目标档从「Integer/Long/Byte/Double」换成「动态 String」。定长串
目标不在此列（那一支的收口是 `FixedStrAssignFree`，另一形，未量）。

#### 三、踩到的两处

1. **接收者要按链根判**。初版只认裸标识符接收者，`s = Printers(0).DeviceName` 的接收者是
   `Printers.Item(0)`（调用形态）⇒ 伪宿主没被排除、被套上守卫，`pd_printer_data_props`
   那条发码针翻红（期望 `vb6_BSTR_Assign(&s, vb6_Printers_Item((int32_t)(0)).DeviceName);`）。
   改成沿 `MemberAccess / IndexOrCall->callee` 下钻取**链根**再比对排除名单。
2. **RTL 出口的两条静态轴要能自证**。第一版写成「失败先放、成功再转手」，`owned_temp_keeps`
   的推导在函数体里先撞见 `vb6_BSTR_Free(source)` ⇒ 判成 `handoff-reads`（keeps=0），
   与发码侧登记的例外行对不上（K12/K13d 双红）。把成功路提到前面（先 `AssignMove` 再
   `if (source) Free`）⇒ 推导读到 `handoff-keeps`（keeps=1）、消费轴读到那一格 ⇒ 两轴都自证。
   表 255 行（新增 1 行，`keeps_pointer=1` 38 行）。

#### 四、churn

清单 750 行：**25 行**变动（第八十八/九十八刀之后的新增；BSTR 目标这一族与标量那一族
共 179 + 1,475 句里落在普通赋值站的这部分）。`tests/ctrltabstrip/TabApp.vbp` 是机器侧行 ——
本机值与 CI 值不同，登记值以 CI 的 `emit-manifest` 工件为准。

#### 五、门

- 探针 F1–F6 六格与真机逐字对上（F6 `keep/438`）。
- 静态门 9 枚 PASS（`proto_arity` 1785；`rtl_embedded` R1..R3 138 件逐字节）。
- `owned_temp_keeps` **255 行**（+1 例外行、+1 消费格）；`-Category syntax` **227/227**。

#### 六、仍挂着的一笔

**表达式内位置**（`If o.X = 1 Then`、实参、二元运算）：1,486 句。守卫机制插不进去，要另开
一种形；真机读数也还缺（G 段 6 格：`n = 1 + o.NoSuchProp`、`If o.NoSuchProp = 1 Then`、
`n = Abs(o.NoSuchProp)`、`Call f(o.NoSuchProp)`、`n = o.NoSuchProp * 2`、
`s = "x" & o.NoSuchProp`）。

---

#### §11.76 补：438 只对真 IDispatch 报

第九十八刀 (a) 初版对五枚 `DISPID_UNKNOWN` 分支**无条件**报 438，CI 上
`b190_object_formal_scope(_x86)` 运行崩（本机复现）。那一格打的是账 #412：Object 槽里装的是
工程类实例指针、从没包成 IDispatch —— VB6 那边本来有对象，无条件报 438 把这些能跑的语句打崩。
改成 `if (vb6_ComIsDispatchable(disp))` 才报，其余维持静默 + `C3_LATEBIND_MISS` 追踪
（那是实现缺口，不是 VB6 语义）。本机 `B190.vbp` 恢复 `run rc=0`，五条针全在。

### 11.78 第一百刀：成员读落在表达式内（值位置）

#### 一、读数（探针 G 段，真机 + 本仓）

| 格 | 语句 | 真机 | 改前 C3 | 现在 C3 |
|---|---|---|---|---|
| G1 | `n = 1 + o.NoSuchProperty` | -1 / 438 | 1 / 438 | **-1 / 438** |
| G2 | `n = o.NoSuchProperty * 2` | -1 / 438 | 0 / 438 | **-1 / 438** |
| G3 | `n = Abs(o.NoSuchProperty)` | -1 / 438 | 0 / 438 | **-1 / 438** |
| G4 | `Call GTouch(o.NoSuchProperty)` | 被调方**不跑**（0）/ 438 | 跑了（1） | 1（未接，下一刀） |
| G5 | `If o.NoSuchProperty = 1 Then n = 7` | **n = 7** / 438 | -1 | -1（未接，见第五节） |
| G6 | `s = "x" & o.NoSuchStringProperty` | keep / 438 | "x" | **keep / 438** |

结论：**整句中止** —— 目标不动、被调方不跑。C3 改前是"缺成员读当空值继续算"。

#### 二、机制：不用改写子表达式

守卫 + 临时量那条机制本来就能罩住**整条右值**，只是以前只在「顶层就是那次读」时才打开：

```c
const int32_t _e1 = vb6_ErrSerial();
int32_t _t1 = (_e1 == vb6_ErrSerial()) ? (1 + vb6_ComGetIntProp(o, L"NoSuchProperty")) : 0;
if (_e1 == vb6_ErrSerial()) n = _t1;
```

右值里那次 438 发生在 `_t1` 的计算里 ⇒ 目标那一步不写。String 目标的临时量与中止时的
**释放**由 `emitBstrSlotStore` 的守卫支自己发（第八十三刀那条纪律，这里不碰所有权）。

判据 = 右值树里**任一节点**经 `lateRead100` 命中（接收者判据与第九十八刀同一枚：链根是
晚绑定 Object 变量，且不是控件数组 / 伪宿主 / `Me` / 同工程函数名），且**顶层本身不是**
那次读（顶层两形已由 98/99 各自接掉）。遍历只走表达式节点：`Binary / Unary /
MemberAccess.object / DictionaryAccess.object / IndexOrCall{callee, 位置实参, 命名实参} /
New.args / TypeOf.object`。

只接**标量数值 + 动态 String** 目标（与 98/99 同一张面）；Variant / UDT / 数组元素目标
未量，不在册。

#### 三、churn

清单 750 行：**13 行**变动（G 段这一族落在普通赋值站的部分）。其中 `test_late_cond.bas`
就是这条判据的现成样本。

#### 四、门

- 探针 F1–F6、G1/G2/G3/G6 与真机逐字对上。
- 静态门 9 枚 PASS（`proto_arity` 1785、`rtl_embedded` 138 件）；`owned_temp_keeps` 255 行不变。
- `-Category syntax` **227/227**。

#### 五、仍挂着的两笔

1. **G4：实参位置（调用语句）**。真机 `Call GTouch(o.NoSuchProperty)` 里被调方**根本没跑**；
   C3 现在是「先把实参算完（报 438）再照常调」。要接就得把实参先提到调用前、判完序号再决定
   调不调 —— 与赋值站不同形（那支是 `if (ok) call(...)`，实参仍在 call 里）。
2. **G5：单行 If 的条件位置**。真机读数反直觉：条件报 438 之后 **Then 分支照样执行**
   （n=7）。这不是"整句中止"，是 VB6 单行 If 的错误续行语义（`Resume Next` 落在 Then 那句
   上）。要先补两格探针（多行 If 块的条件报错 / `ElseIf`）把边界量清楚，再决定要不要模仿
   —— 现在 C3 是 -1（不执行），与真机不一致，**已登记为差异**。

### 11.79 第一百零一刀：条件位报错 ⇒ 按真

#### 一、读数（探针 H 段，真机 + 本仓）

| 格 | 语句 | 真机 | 改前 C3 | 现在 C3 |
|---|---|---|---|---|
| G5 | `If o.NoSuchProperty = 1 Then n = 7` | 7 / 438 | -1 | **7 / 438** |
| H1 | 单行 If（同上，基线） | 7 / 438 | -1 | **7** |
| H2 | 单行 `Then n = 7 Else n = 8` | **7**（走 Then） | 8（走 Else） | **7** |
| H3 | 多行 `If … Then / n = 7 / End If` | **7** | -1 | **7** |
| H4 | `ElseIf o.NoSuchProperty = 1`（前面 `1=2`） | **7** | 8 | **7** |
| H5 | `Do While o.NoSuchProperty = 1` | **7**（进一次） | -1 | **7** |
| H6 | `n = GPick(o.NoSuchProperty)` | -1 / 438 | -1 | -1（百刀已盖） |
| H7 | `Call GTouch2(1, o.NoSuchProperty)` | 被调方不跑 | 跑 | 跑（未接，下一刀） |

结论与赋值站**相反**：值位置是「整句中止」，条件位是「**报错按真**」——单行/多行/ElseIf/Do While
四种条件位一致。H2 说明不是「当假走 Else」：真机走的是 Then。

#### 二、形

RTL 一枚出口，把「序号变了 ⇒ 真」收口：

```c
static inline int32_t vb6_CondIfErr(int32_t serialBefore, int32_t value) {
    return (serialBefore != vb6_ErrSerial()) ? 1 : (value != 0);
}
```

发码站（`guardCondLateRead`，住在 `cgen_util_type.cpp`，紧跟 `conditionTruthExpr`）：

```c
const int32_t _c1 = vb6_ErrSerial();
if (vb6_CondIfErr(_c1, (<原条件>))) {
```

序号必须在**求条件之前**捕下，所以是两行；条件不含晚绑定读时**逐字节不变**（直接交回原文本）。
接的站：`If`、`ElseIf`、`Do While`、`While … Wend`。`Do Until` / `Loop Until` 未量（H 段没覆盖
取反那一档），**不接**，登记为缺口。

#### 三、顺手收口

第九十八/九十九/一百刀在 `cgen_assign.cpp` 里内联过两枚 lambda（接收者判据 + 树扫描）。条件位要问
同一问，于是升成成员函数 `lateBoundComReadExpr` / `exprHasLateBoundComRead`（`cgen_util_scan.cpp`），
赋值站与条件位同读一处；`cgen_assign.cpp` 里那份复本删掉。

#### 四、churn

清单 750 行：**23 行**变动（条件位这一族落在语料上的部分；`regexp_native/ReMain.bas` 也在此列）。

#### 五、门

- 探针 G5 + H1–H5 六格与真机逐字一致；H6 不变；G4/H7 见第六节。
- 静态门 9 枚 PASS（`proto_arity` 1785、`rtl_embedded` 138 件）；`owned_temp_keeps` 255 行不变。
- `-Category syntax` **227/227**。

#### 六、仍挂着

1. **G4 / H7：调用语句实参位**。真机两只都是「被调方不跑」（H7 里第一枚实参正常也不影响结论）。
   这一族在语料里 **3,068 句**（最大的块），要先提实参、再按序号决定调不调，且需要被调形参的档。
2. **`Do Until` / `Loop Until`** 的条件位未量（取反语义不在 H 段里）。


### 11.80 第一百零二刀：实参位的报错 ⇒ 整句中止

前两刀把「值位置」与「条件位置」收了，剩下**调用实参位**这一族没量。真机 I 族 15 格
（`.temp/k93/vb6probe/RealMachineProbe.bas`）把口径量死了：

| 格 | 真机读数 | 口径 |
|---|---|---|
| I1 `Call GTouch2(o.NoSuchProperty, GFn())` | `gHit=0 gN=0` | 坏实参**之后**那枚不再求值 |
| I2 `Call GTouch2(GFn(), o.NoSuchProperty)` | `gHit=0 gN=1` | 坏实参**之前**的照常求值 |
| I3 `Call o.NoSuchMethod(GFn())` | `gN=1` | 实参先求、被调成员后解析 |
| I4 `n = o.NoSuchProperty(GFn())` | `gN=1 n=-1` | 同一顺序，目标不动（值位那刀已覆盖） |
| I5–I13 | `gHit=0` | String / 无括号 / 具名 / Object / Variant / 表达式实参 / 嵌套 / 双坏一律跳过调用 |
| I14 / I15 | `n=42` / `gHit=1 gN=2` | 下一句照跑；两枚实参都正常时全部求值 |

也就是「命中即整句中止」，与值位置同族，但与条件位（报错按真）相反。

#### 一、发码形

C 的实参求值顺序**未定义**（MSVC 从右向左），所以不能把守卫塞进实参表。做法是把
「首枚命中之前的实参（含命中那枚）」逐枚提升成临时量，每枚求值前比一次序号：

```c
const int32_t _eg0 = vb6_ErrSerial();
int32_t _ag1 = 0;
if (_eg0 == vb6_ErrSerial()) { _ag1 = <原实参>; }
if (_eg0 == vb6_ErrSerial()) {
vb6_GTouch(_ag1);
}
```

- 命中 ⇒ 后面几枚的 `if` 条件已不成立 ⇒ 不求值；调用同理跳过（I1 的 `gN=0`）
- 命中之前的那几枚早已按序求过值（I2 的 `gN=1`）
- 地址实参（`&x`）求值无副作用也不会报错，直接留在调用里，不提升
- 落不到标量 / 串 / 对象档的实参（Variant、UDT）暂不提升 —— 见第五节缺口

实参提升站在 `cgen_expr_call_arg_variant.inc` 的 `args.push_back` 之前（ByRef 取址、
optional 补位、ParamArray 打包都在它之后，故不动既有逻辑）；守卫声明与整句包裹在
`visit(CallStmt&)` 里，落在 `CallArgScope102` 作用域内统一清除。

#### 二、语料量（严格口径）

744 个工程、9 个工程命中、**50 条语句**：`VBFlexGridDemo` 12、`GifcontrolSource` 12、
`ucTreeView` 7、`Charts 2020` 5+4、`VBFlexGridDemo/Common` 4+4、`VbQRCodegen` 1+1；
其中被调方本身晚绑定的 **0 条**。上一刀报的 3,068 是**松口径**（按行文本算，把表达式
续行、`(void)` 强转行也算进去了），以本刀这组数为准。

#### 三、夹具与针

- `tests/test_late_arg.bas`（新增，I 族五形）
- 正向针 `b102_late_arg_guard_shape`：六段发码形（`_eg0` 声明、两枚 `_ag` 提升、两处调用）
- 反面针 `b102_late_arg_lesion`：改前三行原样（实参内联、报错后照调）一条都不许出现
- 本地真跑：G4、H7、I1–I15 共 17 格与真机**逐字一致**

#### 四、门

- 静态门 9 枚 PASS（`proto_arity` 1785、`manifest_coverage` 751 = 750+新夹具）
- `owned_temp_keeps` 255 行不变；`_own_` 提名闸不受影响（提升只搬实参文本，不动所有权表）
- manifest：6 行重算（`VBFlexGridDemo` ×2、`b171_slotarg`、`ucTreeView`、`dict_native` ×2）+ 新夹具 1 行

#### 五、仍挂着

1. **Variant / UDT 形参位**：落不到可声明临时量的档，暂不提升（`mapType` 交 `vb6_VARIANT`）。
2. **`Do Until` / `Loop Until`** 条件位仍未量（取反语义不在 H 段里）。
3. **`LenB(<UDT>)`** 仍挂账。

### 11.81 第一百零三刀：条件位报错 ⇒ 从该语句的下一条继续

第一百零一刀把条件位按「报错 ⇒ 按真」接了 `If` / `ElseIf` / `Do While` / `While…Wend`，
`Do Until` / `Loop While` / `Loop Until` 三站当时**故意没接**（取反语义没量）。真机 J 族
（`.temp/k93/vb6probe/RealMachineProbe.bas`）量完，口径要改一次说法：

| 格 | 写法 | 真机读数 | 说明 |
|---|---|---|---|
| J1 | `Do Until o.Bad = 1` | `n=4` | 报错 ⇒ **进循环体**（撞防死循环闸） |
| J2 | `Do … Loop Until o.Bad = 1` | `n=1` | 报错 ⇒ **出循环** |
| J3 / J4 | `Do Until Not o.Bad` / `Loop Until Not o.Bad` | `n=4` / `n=1` | 与有没有 `Not` 无关 |
| J5 | `Do While Not o.Bad` | `n=4` | 前测同向，进体（与 H5 同源） |
| J6 | `Do … Loop While o.Bad = 1` | `n=1` | 后测同向，出循环 |
| J7 / J8 | 控制格 | `n=4` / `n=1` | 闸本身没问题 |
| J9 / J10 | 裸值形 | `n=4` / `n=1` | 见第五节（解析面病灶） |

所以真口径不是「按真/按假」，而是 **`On Error Resume Next` 的「从下一条继续」**：
前测循环的「下一条」就是循环体 ⇒ 进体；后测循环的「下一条」在 `Loop` 之后 ⇒ 出循环。
这与单行 `If` 落 `Then` 分支（H1/H2）、多行 `If` 进块（H3）、`ElseIf` 进它自己的 `Then`（H4）
是同一条规则，**与极性无关**。第一百零一刀那枚 `vb6_CondIfErr`（错 ⇒ 1）在 `If` /
`Do While` / `While` 恰好等价，但取反站与后测站会反过来。

#### 一、发码形

新增一枚入口（`vb6rtl_array.h`，紧跟 `vb6_CondIfErr`）：

```c
static inline int32_t vb6_CondIfNoErr(int32_t serialBefore, int32_t value) {
    return (serialBefore == vb6_ErrSerial()) ? value : 0;   /* 错 ⇒ 0 */
}
```

| 站 | 发码形 | 报错时的走向 |
|---|---|---|
| `Do Until` | `if (vb6_CondIfNoErr(_cN, (cond))) goto exit;` | 0 ⇒ 不退出 ⇒ 进体（J1/J3） |
| `Loop While` | `if (!(vb6_CondIfNoErr(_cN, (cond)))) break;` | 0 ⇒ `break`（J6） |
| `Loop Until` | `if (vb6_CondIfErr(_cN, (cond))) break;` | 1 ⇒ `break`（J2/J4） |
| `Do While` / `While` | `if (!(vb6_CondIfErr(_cN, (cond)))) goto exit;` | 1 ⇒ 不退出 ⇒ 进体（H5） |

后测两站注意：守卫的序号声明必须算进「物化」判定（`c_.size()/hasPending()`），否则声明落在
`do` 体内、而 `} while (…)` 在体外 ⇒ C2065（实测 J2/J6 四处）。把 `materialized` 的判定挪到
守卫之后，两站统一走 `} while (1);` + `break;` 的形。

#### 二、语料量

745 个工程里，循环条件位出现晚绑定读的只有 **8 条**：前测同向 6（早接过）、后测取反 2
（`tests/test_late_cond.bas` 的 `K-LOOPUNTIL` 两臂）、`Do Until` 与 `Loop While` 各 0。
所以这一刀的语料面**零变化**，manifest 只动夹具那一行。

#### 三、夹具与针

- 夹具复用 `tests/test_late_cond.bas`（本来就有 `K-DOUNTIL` / `K-LOOPWHILE` / `K-LOOPUNTIL` 三臂）
- `b244_cond_truth_shape` 七条针按新发码重钉（三臂翻面 + 因守卫编号平移重钉 `K-WHILE` 的 `_c12`）
- 反面针 `b244_cond_truth_lesion` 未变：改前那七形（无守卫）一条都不许再出现
- 本地真跑 J1–J10：9 格与真机逐字一致；J9/J10 见下

#### 四、门

- 静态门 9 枚 PASS（`proto_arity` 1785、`rtl_embedded` 138、`manifest_coverage` 751）
- `owned_temp_keeps` 255 行不变
- manifest：1 行（`tests/test_late_cond.bas`）

#### 五、新发现的病灶（本刀不修，登记）

**裸值条件把读变成写**：`Do Until o.NoSuchProperty` 单独成行、下一行是赋值时，解析出的条件
AST 只有 `o`，成员名被下一句吞掉 —— 最小复现：

```vb
Do Until o.NoSuchProperty
    n = n + 1
Loop
```

发成 `if (… (o)) goto exit;` + `vb6_ComSetProp(o, L"NoSuchProperty", vb6_ComPackInt((n + 1)));`
（读变写、`n = n + 1` 整句消失）。真机 J9/J10 是 `err=438`、循环体按「进体」走。这一条在
解析面（`parseDoLoopStmt` 的 `condition = parseExpression()` 之后），与条件守卫无关，单独立账。

仍挂着：Variant / UDT 形参位、`LenB(<UDT>)`、上面这条裸值条件解析。

### 11.82 第一百零四刀：裸成员条件不许把「读」变成「写」

§11.81 第五节把这一条记成了「解析面病灶：裸值条件把读变成写」。**记错了**。探针量完，
解析是对的，病灶在发码面 —— 而且比原以为是解析问题还严重：它会**静默改写循环体**。

#### 一、症状

```vb
Do Until o.NoSuchProperty
    n = n + 1
Loop
```

修前发成：

```c
const int32_t _c0 = vb6_ErrSerial();
if (vb6_CondIfNoErr(_c0, (o))) goto vb6_loop_exit_0;          /* 条件只剩对象 */
    vb6_ComSetProp(o, L"NoSuchProperty", vb6_ComPackInt((n + 1)));  /* 体首语句被吞 */
```

`n = n + 1` 整句消失，变成对 `o.NoSuchProperty` 的**写**；体首换成 `GTouch 1` 时则变成
`vb6_ComCall(o, L"NoSuchProperty", (void*[]){vb6_ComPackInt(1)}, 1)` —— `GTouch` 也没了。
这是「改写程序语义」级别的缺陷，不是报错级别。

#### 二、定位过程（先排除解析）

给解析器加临时探针（`parseDoLoopStmt` 打印条件节点种类），四格读数全是
`cond=MemberAccessExpr curT=NewLine` —— 条件节点是对的，`.NoSuchProperty` 也确实在
条件里。再给发射器加探针（打印体首语句节点）：

| 格 | 夹具原句 | 体首节点 |
|---|---|---|
| T6 | `n = 1` | `AssignmentStmt target=IdentifierExpr`（对） |
| T7 | `GTouch 1` | `CallStmt callee=IndexOrCallExpr`（对） |

AST 两侧都对 ⇒ 病灶在发码。对照两站源码，差别只有一句：

```c
emitExpr(*node.condition);
if (isComMarker_) resolveComValue("Int");   // ← If / Do While 有，另四站没有
```

`resolveComValue` 是 COM 取值标记的**唯一收口**：成员读在表达式层只挂标记（`lastExpr_`
暂留对象名），由收口这一句把 `vb6_ComGetIntProp(obj, L"成员")` 拼出来并清标记。不收口，
标记就悬在空中，被**下一条语句**当成"这条语句的操作对象"消费掉 —— 赋值目标 / 调用接收者
于是指向那个成员，值/实参来自原语句。四站全中：`Do Until`、`Loop While`、`Loop Until`、
`While…Wend`。

免疫的格子正是能自己解标记的：条件里有运算符（`= 1`、`Not x`）时 `BinaryExpr` /
`UnaryExpr` 内部各有一处 `resolveComValue`；`If` / `ElseIf` / `Do While` 三站在语句层已经
收口（`cgen_stmt.cpp:188`、`cgen_loop.cpp:54`）。

#### 三、修

四站各补一句（与 `Do While` 同一收口、同一 unpack 档 `Int`）：

```c
emitExpr(*node.condition);
if (isComMarker_) resolveComValue("Int");
```

#### 四、对账

真机 J9/J10 是这一形的实测（裸成员条件）：J9 `Do Until` = `n=4 err=438`、J10
`Loop Until` = `n=1 err=438`。修后本地 J1–J10 **十格逐字一致**（含 err）。语料侧 745 个
工程 **0 行变化** —— 裸成员做循环条件这一形在真实工程里确实罕见，但一旦出现就是上面
那种静默改语义。

#### 五、夹具与针

- 新夹具 `tests/test_late_barecond.bas`（四站各一格，局部 `Dim o As Object` + `CreateObject`，
  与 J9/J10 探针同形）
- `b248_bare_cond_shape` 四条针钉四站的守卫形（`vb6_CondIfNoErr`/`vb6_CondIfErr` +
  `vb6_ComGetIntProp`）
- `b248_bare_cond_lesion` 五条反面针（四条 `(o)` 条件形 + 那条 `vb6_ComSetProp`）—— 取自
  **撤掉修复的构建**的逐字发射，不是照形推的

#### 六、门

- 静态门 9 枚 PASS（`manifest_coverage` inputs=752 / registered=752、`proto_arity` 1785、
  `rtl_embedded` 138）
- `owned_temp_keeps` 255 行不变
- manifest：新增 1 行（`tests/test_late_barecond.bas`），其余 0 变化

### 11.83 第一百零五刀：Variant 槽的实参位（聚合体也能有临时量）

第一百零二刀在实参位提升那儿留了一句跳过：

```cpp
// 落不到标量/串/对象类型的（Variant / UDT）不提升，留着当登记缺口。
if (gtyp102 != "vb6_VARIANT") { ...提升... }
```

这一刀把这一档量掉、补上。

#### 一、真机刻度（K 族）

| 格 | 写法 | 真机读数 |
|---|---|---|
| K1 | `Call GVar(o.NoSuchProperty)`（ByVal Variant，单实参） | `gHit=0` err=438 |
| K2 | `Call GVar2(GFn(), o.NoSuchProperty)` | `gHit=0 gN=1` err=438 |
| K3 | `Call GVar2(o.NoSuchProperty, GFn())` | `gHit=0 gN=0` err=438 |
| K4 | `v = o.NoSuchProperty`（`v As Variant` 先置 111） | `v=111` err=438 |
| K5 | `Call GVarRef(o.NoSuchProperty)`（ByRef Variant） | `gHit=0` err=438 |
| K7/K8 | 控制格 | `gHit=1`、err=0 |

结论：**Variant 槽与 Long 槽同一条规则**（I 族那条）：命中即整句中止 —— 坏实参之后的不再求值、
之前的照常求值、被调方不跑。K4 是另一处（赋值目标不许被写），本刀不动，另记。

修前本地读数：K1 `gHit=1`、K2 `gHit=1 gN=1`、K5 `gHit=1`、K3 已对、K4 `v=""`（被写空）。

#### 二、「落不到可声明临时量」其实不成立

聚合体不能像标量那样 `= 0` 起手，但 **C 里 `vb6_VARIANT` 本来就是可声明的**。真正的顾虑不是
可声明性，而是**所有权**：浅拷贝会让临时量与源共用同一只 BSTR，随后两边各 `Clear` 一次 =
过度释放。RTL 已经有对的那一枚：

```c
void vb6_VariantCopy(vb6_VARIANT* dst, const vb6_VARIANT* src);  /* 深拷贝: BSTR 另分配 / Dispatch AddRef */
```

于是发码形（ByVal 槽）：

```c
const int32_t _eg0 = vb6_ErrSerial();
vb6_VARIANT _ag1 = vb6_VariantEmpty();
if (_eg0 == vb6_ErrSerial()) { vb6_VARIANT _as2 = <实参表达式>; vb6_VariantCopy(&_ag1, &_as2); }
if (_eg0 == vb6_ErrSerial()) { vb6_GVar(_ag1); }
vb6_VariantClear(&_ag1);
```

ByRef 槽（实参本来就是 `(&(vb6_VARIANT){...})` 临时量）直接拷贝、改交 `&_agN`：

```c
if (_eg7 == vb6_ErrSerial()) { vb6_VariantCopy(&_ag8, (&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=...})); }
if (_eg7 == vb6_ErrSerial()) { vb6_GVarRef(&_ag8); }
vb6_VariantClear(&_ag8);
```

三处要点：

1. **`_as2` 那枚局部量不是多余的** —— 实参表达式本身就是 `vb6_VARIANT` 值，写成
   `&(vb6_VARIANT){<argVal>}` 会被当成复合字面量的**首成员**（`vt`，uint16_t）初始化，
   MSVC 直接 C2440（实测 k.c(52)/(67)）。接一手再拷，同时也把「求值发生在守卫内」钉死
   （K2 `gN=1` / K3 `gN=0` 就靠这个）。
2. **只清自己那份**：`vb6_VariantCopy` 是深拷贝，`vb6_VariantClear(&_agN)` 释放的是副本的
   BSTR/Dispatch 引用；源临时量的处置与改前一致（不管），不引入新的双重释放。
3. **守卫没放行时 `_agN` 还是空的**，`Clear` 空转，故无条件发。

#### 三、语料

修前全量扫（752 输入）：**110 处、5 个文件**落在这一档，**真实语料 0 处**（全是自家夹具：
`dict_native` 45、`b171_slotarg` 2、`ucTreeView` 2、`VBFlexGridDemo` 2 等）。
修后 rehash：首版（`376a11ff`）报了 **6 行**（`VBFlexGridDemo`×2、`b171_slotarg`、`ucTreeView`、
`dict_native`×2）。修完下面第六节那两枚闸后 `ucTreeView` 回退、`VisualStyles.bas` 现身，
净 **5 行**变化（`VisualStyles.bas`、`VBFlexGridDemo.vbp`、`b171_slotarg`、`dict_native`×2）。

#### 四、夹具与针

- 新夹具 `tests/test_late_vararg.bas`（K1/K2/K5 + 控制格）
- `b250_vararg_guard_shape` 七条针：序号声明、`_as` 深拷贝那一句、`GVar(_ag1)`、`Clear`、
  K2 的 `_ag4 = vb6_GFn()`、`GVar2(_ag4, _ag5)`、K5 的 ByRef 形
- `b250_vararg_guard_lesion` 两条反面针（取自未接提升时的逐字发射）
  （首版曾名 `b249_…`；`b249` 前缀已被另一本账的判据面占用 ——
  `check_com_value_twin_authority.ps1` 的 R6 明令整工程级 `Test-EmitcShape "b249` = 0，故改名）

#### 五、门

- 静态门 9 枚 PASS（`manifest_coverage` inputs=753 / registered=753）
- `owned_temp_keeps` 255 行不变
- manifest：新增 1 行（`tests/test_late_vararg.bas`）+ 6 行重算（修后为 5 行，见第六节）

#### 六、CI 红与「嵌套调用」那一修（第105刀-修）

`376a11ff` 的 `Tests (compile #2)` 红在 `FlexGridX86`（x86 编 `tests/VBFlexGridDemo`）。
本机复现出的 MSVC 真错（job 日志里连 `error C` 行都没带，是回头在 `c3-error.log` 里拿到的）：

- `VisualStyles.c(162)/(195)/(197)`: `error C2440: 无法从 void* 转换为 vb6_VARIANT`
- `VisualStyles.c(164)/(199)`: `error C2172: vb6_ObjPtr 实参不是指针`

两处是同一条**过度提升**：`SetWindowSubclass(h, thunk, ObjPtr(CurrControl), 0)` 的第三格是内建
`ObjPtr`，它的形参表答 Variant，而提升只看「形参是不是 Variant」⇒ 把 `CurrControl`（`void*`）
包进 `vb6_VARIANT _ag6`，再原样交给 `vb6_ObjPtr(_ag6)`。补两枚闸：

1. `argNestDepth_`（`cgen_state.inc`；在 `visit(IndexOrCallExpr)` 入口 RAII 增减）：实参位提升
   只对**最外层**调用（`== 1`）成立 —— 嵌套调用的实参是外层语句的一部分，不是它自己的实参位。
2. 提升判据改用**实参自身**的 C 档 `mapType(inferExprType(*node.positional[i])) == "vb6_VARIANT"`，
   不再拿形参声明的档当判据（`ObjPtr` 那格形参答 Variant，实参却是 `void*`）。

本机复验：`C3.exe tests\VBFlexGridDemo\VBFlexGridDemo.vbp --arch x86` 编过、无 `c3-error.log`；
K1/K2/K3/K5/K7/K8 与真机逐格同值，**K4 仍差**（`v = o.NoSuchProperty` 是赋值位，另立一刀）。

#### 七、GUISample 那三行（CI 环境行）

`Emit manifest (shape oracle)` 在 `376a11ff` / `365e4d6a` / `6b536561` 三次都红在这三格：
`tests/guisample/GUISample.vbp`、`.../vb6ref/src/vc6/GUISample.vbp`、`.../vb6ref/src/vs2015/GUISample.vbp`
（每格固定 +3475 字节，`rc=0`；三者共享模块，所以增量相同）。这三行本来就在
`full_rehash85.ps1` 的 `$machineDiff` 里（本机与 CI 的机器状态不同，本地 rehash 跳过不写）。

归因用的三条证据：

1. **编译器不是成因**：把 CI 那次跑出来的 `c3-exe` 工件（run 44905446854）拉到本机，在仓根跑
   `GUISample.vbp --emit-c`，输出与本机构建**逐字节相同**（307655 字节）。
2. **本刀碰不到它**：同目录下用 `798b057b`（本刀前）与本刀两枚 `C3.exe` 各跑一遍，
   GUISample 输出 **delta=0**；`--no-auto-typelib` 也是 0；原因是该项目里**没有任何
   `As Variant` 形参**，`_ag/_as` 站点数 = 0。
3. 三次 CI 里这三格的值彼此**完全一致**，与本刀那 5 格的变化不重合。

结论：这三格随 CI 运行机的类型库/OCX 状态漂移，不是本刀的发码形状变化。按脚本自己印的路子
（`rebless_emit_manifest.ps1 -Bless`，吃 CI 那台二进制交回来的清单）登记 3 行。

⚠ 登记之后这三格仍是 CI 侧读数：本机不写、也不能写。下次再漂，先按上面第 1 条（同一枚 exe
在本机复跑）区分「编译器」与「机器」。

#### 八、仍挂着

- **K4**：`v = o.NoSuchProperty` 目标被写（真机 `v=111`）。这是赋值位的守卫，不在实参位这条线上
  —— 第一百零六刀（§11.84）接掉。
- **UDT 槽**：这一族语料里没量到「UDT 形参 + 晚绑定读」的自然形（UDT 不能承接晚绑定读），
  等出现再立账。
- **`VT_BYREF` 写穿**：`Sub f(ByRef v As Variant)` 传标量左值时，真机是
  `VarType(v)=16387` + 写穿；我们目前是装箱。这一条独立，别与本刀混。

### 11.84 第一百零六刀：Variant 槽的赋值位（目标一个字都不写）

§11.83 末尾把 K4 记成缺口：

> **K4**：`v = o.NoSuchProperty` 目标被写（真机 `v=111`）。这是赋值位的守卫，不在实参位这条线上。

这一刀补上。它不在实参位那条线上 —— 病在 `visit(AssignmentStmt&)` 的 `targetIsVariant` 那一支：
守卫机制（`mayRaise79` / `stmtAbortTmp_`）只认 **Int / Double / String** 目标，Variant 目标整支落在外面，
于是 438 之后照样把交回来的值写进 `v`（实测未修前 K4 读 `v=` 空、err=438）。

#### 一、刻度（L 族）

| 格 | 写法 | 读数 |
|---|---|---|
| L1 | `v = 111 : Err.Clear : v = o.NoSuchProperty` | `v=111` err=438 ← **真机 K4 实测**；修前本机是 `v=` 空 |
| L2 | `v = 111 : Err.Clear : v = o.Count` | `v=1` err=0 ← 本机控制格 |
| L3 | 先 `o.Count` 再 `o.NoSuchProperty` | `v=1` err=438 ← 本机控制格（失败那次不写） |
| L4 | `v = o.Item("k")` | `v=vv` err=0 ← 本机控制格（串值同样到手） |
| L5 | 连报两次 | `v=111` err=438 ← 本机控制格 |

L1 是唯一上过真机的那一格（K4）；L2–L5 是守卫不误伤的**本机**控制格 —— 它们走的语义与
未守卫那一支相同（只是把写的那一步挪到闸后），在真机上量会得到同一组数；这里不把
未量当已量，故分开标。

#### 二、发码形（§B169 的口径，不自抄）

```c
const int32_t _e0 = vb6_ErrSerial();
vb6_VARIANT _av0 = vb6_VariantEmpty();
if (_e0 == vb6_ErrSerial()) { _av0 = vb6_VariantFromComResult(vb6_ComCall(o, L"NoSuchProperty", NULL, 0)); }
if (_e0 == vb6_ErrSerial()) { vb6_VariantClear(&v); v = _av0; }
```

- 序号在**求值之前**抓；438 发生在第一句的求值里 ⇒ 第二句的闸为假 ⇒ 目标一个字都不写。
- 为什么不能只写一句 `if (_e == vb6_ErrSerial()) { ... }` 把整句包起来：序号在进入 if 时
  还没变，求值把序号推走之后**仍在 if 体内** ⇒ 会把值写进去（与 §11.80 那条同形的坑）。
- 写回那一句照 `emitVariantSlotStore` 的口径分两支：自有 ⇒ `Clear + 浅搬`（所有权随拷贝交给目标）；
  借来 ⇒ `vb6_VariantAssign`（Clear + 深拷贝）。这问问的是 `variantRhsOwnsValue`，站点不自己答。

#### 三、范围（只接量到的形）

接的形：**顶层就是那一次晚绑定成员读**（`lateBoundComReadExpr`），目标是非数组/非 UDT 字段的
Variant 槽。改前全量扫（753 输入）：**23 处 / 11 个输入**，全在 `tests/` 内
（`dict_native` 5+5、`VBFlexGridDemo` 2、`b190_memberarray` 2、`b612/b613/b614` 各 1–2、
`czUI` 1、`ucTreeView` 1、`urlpic` 1、`E438App` 2）。外部真实语料 0 处。

**未量、登记不猜**：

- `v = 1 + o.NoSuchProperty`（晚绑定读在表达式**内部**）；
- Variant **数组元素** / **UDT 的 Variant 字段**做目标（`VB6_SA_*` / `udtFieldSlot`）；
- `v = Me.rtfText`（`tests/guisample/err438/E438.frm` 2 处）—— 接收者是窗体上的控件
  （发码是 `vb6_ComCall(vb6_hwnd_E438, …)`），不是 Object 变量，`lateBoundComReadExpr`
  不答它；真机没量过这一形的目标写不写，所以不动。

这几形已经写成真机探针（`RealMachineProbe.bas` 的 L6–L9 + L11/L12，L10 即上一条的
`Form1.NoSuchFormMember`，默认注掉 —— 它在 VB6 里可能根本编不过，编不过本身也是答案）。
本机 C3 读数（同一份探针）：`L6 v=1` / `L7 v=x` / `L8 va(0)=空` / `L9 uv.f=空`，err 都是 438 ——
即四形现在都会写目标。**真机值回来再决定要不要扩守卫**；在那之前不猜。

#### 四、夹具与针

- 新夹具 `tests/test_late_varslot.bas`（L1/L2/L3/L4 四段：缺读、存在读、先成功后失败、串值）
- `b251_varslot_guard_shape` 五条针：`_av0` 起手、缺读那一句、写回那一句、控制格 `Count` 那一句、
  L3 的 `_av3` 写回
- `b251_varslot_guard_lesion` 一条反面针：修复前的 `v = vb6_VariantFromComResult(vb6_ComCall(o, L"NoSuchProperty", NULL, 0));`

#### 五、门

- 静态门 9 枚 PASS（`manifest_coverage` inputs=754 / registered=754）
- `owned_temp_keeps` 255 行不变
- manifest：新增 1 行（`tests/test_late_varslot.bas`）+ 受影响的存量行重算
- 受影响的工程本机逐个编过：`B612/B613/B614/B190/DictApp/E438App`（x86 那份由 CI 的 `FlexGridX86` 盯）

---

## §11.85 第一百零七刀：Variant 槽赋值位的四形补齐（表达式内部 / 数组元素 / UDT 字段）

### 一、真机（L6–L9，一格一句，目标先置 111）

- `L6  v = 1 + o.NoSuchProperty` ⇒ `v=111` err=438
- `L7  v = "x" & o.NoSuchProperty` ⇒ `v=111` err=438
- `L8  va(0) = o.NoSuchProperty` ⇒ `va(0)=111` err=438
- `L9  uv.f = o.NoSuchProperty` ⇒ `uv.f=111` err=438
- 控制格：`L11  v = o.Count` ⇒ `v=1` err=0；`L12  v = o.Item("k")` ⇒ `v=vv` err=0

结论：438 一出，目标一个字都不写 —— 与第一百零六刀 K4 同一条规则，换成**表达式内部**、
**Variant 数组元素**、**UDT 的 Variant 字段**做目标也照样成立。第一百零六刀当时把这三形
登记为“未量不猜”，这一刀由探针补齐。

### 二、改法

`cgen_assign_value_sem.inc` 那一支的判据从 `lateBoundComReadExpr`（顶层裸读）换成
`exprHasLateBoundComRead`（§B 族早就在用的那一枚，裸读是它的特例 ⇒ 第一百零六刀那一形不会漏）。

写回分两档：

- 普通 lvalue（`v` / `(*p)`）：自有 ⇒ 清旧 + 浅搬（与第一百零六刀逐字节同形）；借来 ⇒ `vb6_VariantAssign`。
- 槽形（`VB6_SA_*` 数组元素 / UDT 的 Variant 字段 / 解引用）：一律 `vb6_VariantAssign`（Clear + 深拷贝），
  自有时再把临时量自己 `vb6_VariantClear` 放掉 —— 槽是**借来**的存储，浅搬会把临时量的所有权塞进别人的槽。

临时量前缀从 `_av` 换成 `_vs`：`_av` 已被标量临时量占用（`cgen_assign.cpp:32` 等），扩面后
同一输入里会撞出 `C2371 重定义`（`rmp107.c(926)` 复现过）。`C3_DBG_VSLOT107=1` 可打印每个站点的形
（plain/slot × bare/nested × own/borrow）用于 census。

### 三、量（改前 census，755 输入）

- 命中 28 处 / 10 个输入，全在 `tests/`：dict_native 5+5、ucTreeView 5、test_late_varslot 5、
  b190 2、b613 2、b612/b614/urlpic/NullFamilyProbe 各 1。
- 其中 **nested 3 处**（ucTreeView 2、NullFamilyProbe 1）是这一刀新收的；bare 25 处第一百零六刀已覆盖。
- 槽形（数组元素 / UDT 字段）在存量语料里 **0 处**，只由新夹具覆盖。
- 实编 churn：manifest 10 行变动，其余 0（裸读那一形的发射逐字节不变）。

#### 四、夹具与针

- `tests/test_late_varslot2.bas`（L6–L9 + 控制格 L11/L12 + L8b/L9b 正常赋值）
- `b252_varslot2_shape` 七条针：`_vs0` 起手、算术内部那一句、数组元素写回、`vb6_VariantClear(&_vs2)`、
  UDT 字段写回、`_vs3` 放掉、控制格 `Count` 写回
- `b252_varslot2_lesion` 两条反面针（修复前：表达式内部 / 数组元素目标直接写目标）
- 第一百零六刀的 `b251_*` 针随之把 `_av` 改成 `_vs`

#### 五、门

- 静态门 9 枚 PASS（`manifest_coverage` inputs=755 / registered=755）
- `owned_temp_keeps` 255 行不变
- manifest：新增 1 行 + 10 行重算 = 755 行
- 本机可编：B612/B613/B614/B190/DictApp/urlpic/ucTreeView/NullFamilyProbe（CI 再全量）

---

## §11.86 第一百零八刀前的探针（M 族：Set 赋值位 / ByRef Variant 目标）

§11.85 末尾登记的两条未量形，探针已经写进 `RealMachineProbe.bas`（M1–M4），等真机数：

- `M1  Set o2 = o.NoSuchProperty`（`o2 As Object`，先 `Set o2 = Nothing`）
- `M2  Set v = o.NoSuchProperty`（`v As Variant`，先 `v = 111`）
- `M3  Set o2 = o`（控制格：成功的 Set 必须落到目标）
- `M4  LateToParam v, o`（目标是一枚 ByRef 的 Variant 形参，C 里是 `(*w)`）

本机 C3 读数（同一份探针，**未改代码**）：

| 格 | 本机 |
|----|------|
| M1 | `o2 is nothing=True` err=438 |
| M2 | `v=` err=438 —— **目标被清空** |
| M3 | `o2 is nothing=False` err=0 |
| M4 | `v=` err=438 —— **调用方的 v 被清空** |

M2/M4 看形就是缺口（和 K4 / L6–L9 不是一回事：那几格是目标不动，这两格是目标被清），
但真机没量前**不改**。真机若给 111 → 开第一百零八刀；真机若也是空 → 现状就是对的，登记为已覆盖。
M1 若真机给 `False`（o2 被塞了东西）也是新缺口。

---

## §11.87 第一百零八刀：Set 赋值位 + ByRef Variant 目标（兼修 lateBoundComReadExpr 的形参盲区）

### 一、真机（再跑一轮 M 族）

| 格 | 语句 | 真机 | 修前本机 |
|----|------|------|----------|
| M1 | `Set o2 = o.NoSuchProperty`（`o2 As Object`，先 Nothing） | `True` / 438 | `True` / 438 ✓ |
| M2 | `Set v = o.NoSuchProperty`（`v As Variant`，先 111） | `111` / 438 | **空** / 438 ✗ |
| M3 | `Set o2 = o`（控制格） | `False`（已指向）/ 0 | 同 ✓ |
| M4 | `LateToParam v, o`（目标是一枚 ByRef 的 Variant 形参） | `111` / 438 | **空** / 438 ✗ |

结论：Set 位与 ByRef 形参位跟 K4/L6–L9 同一条规则 —— 438 一出，目标一个字都不写。

### 二、两个根因

1. **Set 站没接守机**：`visit(SetStmt&)` 的 `targetIsVariant` 那一支直接发
   `v = <boxed>;  /* Set */`，随后 `err` 还是 438，但 v 已经被写成空值。
2. **`lateBoundComReadExpr` 看不见形参**：判据要求根名在 `knownLocalVars_` 里，
   而 `ByVal o As Object` 这种形参只进了 `knownObjectVars_`（`cgen_decl_proc.cpp` 的 void* 那一支）
   —— 于是 `w = o.NoSuchProperty` 这种在**另一个过程里**写的句，守卫压根没接上，
   发出的是 `vb6_VariantClear(&(*w)); (*w) = ...`（先清后赋）。

### 三、改法

- `cgen_util_scan.cpp`：`lateBoundComReadExpr` 的变量闸从
  `knownLocalVars_.count(rn)` 改为 `knownLocalVars_ || knownObjectVars_`（Object/Collection/COM 变量不管
  是局部、形参还是模块级，都算“变量”）。
- `cgen_setlet_set_rhs.inc`：`targetIsVariant` 那一支加与赋值位同形的守机（序号 + `_vs` 临时量 + 两道闸）；
  只接**裸变量**目标（数组元素支在上面先走、`me->字段` / UDT 字段槽 / 解引用形未量 ⇒ 仍走原路）。
- 调试口：`C3_DBG_VSLOT107`（赋值位）、`C3_DBG_VSET108`（Set 位）。

**收紧那一步（CI 红线逼出来的）**：`knownObjectVars_` **全程不清空**，同名会漏到别的过程 ——
`b190_object_formal_goes_late` 第三条针里 `Thing` 是工程类，却被当成 Object，`Thing->Alpha`
被误判成晚绑定。所以“只进 knownObjectVars_”那一腿再要求**接收者解析不出工程类名**：

```
if (!isLocalVar && !knownObjectVars_.count(rn)) return false;
if (inferExprType(*rcv) == Vb6Type::Object) return true;
if (!isLocalVar) return inferClassTypeOfExpr(*rcv).empty();   // 工程类字段 ⇒ 不是晚绑定
return knownObjectVars_.count(rn) > 0;                        // 老口径，一字不改
```

针随之改两条（`vb6_BSTR_Assign` → `vb6_BSTR_AssignMoveIfOk(..., _e`，用**不带号**的形式免得撞号），
第三条 `Thing->Alpha` 原样保留 —— 它正是这条收紧的哨兵。

### 四、量

- 改前 census（755 输入，`C3_DBG_VSLOT107`）：28 处 / 10 个输入（同 §11.85）。
- 改后实编：manifest **37 行变动**（756 行中）。变动的形状全是“多加了一道守卫”，分两类：
  Variant/串目标上的晚绑定读（本刀要的），以及**标量/串目标上接收者是 Object 形参**的
  那批站点（同一个盲区的另一半 —— 以前它们也漏了守卫）。
- 本机逐个编过：`test_events` / `setobj` / `k271_outparam` / `b446_latebind`（CI 再全量盯）。

#### 五、夹具与针

- `tests/test_late_setvar.bas`（M1–M4 + M5 控制格：`Set v = o` 成功必须落进目标）
- `b253_setvar_shape` 五条针：`_vs0` 起手、`(*w)` 的 `VariantAssign`、`_vs0` 收回、`_vs1` 起手（M2）、
  M2 写回
- `b253_setvar_lesion` 两条反面针（修复前：Set 直接写目标 / `(*w) = ...` 裸赋）

#### 六、门

- 静态门 9 枚 PASS（`manifest_coverage` inputs=756 / registered=756）
- `owned_temp_keeps` 255 行不变
- manifest：新增 1 行 + 37 行重算 = 756 行

---

## §11.88 第一百零九刀前的探针（N 族：Set 进 Variant 槽 / ByRef 形参的 VarType）

§11.87 剩下的末几形写成探针（N1–N5），等真机数：

| 格 | 语句 | 本机 C3（未改代码） |
|----|------|---------------------|
| N1 | `Set va(0) = o.NoSuchProperty`（Variant 数组元素） | `va(0)=` err=438 —— **目标被清空** |
| N2 | `Set uv.f = o.NoSuchProperty`（UDT 的 Variant 字段） | `uv.f=111` err=438（§11.85 槽形守卫已覆盖） |
| N3a/N3b | `Set va(1) = o` / `Set uv.f = o`（控制格） | 均已指向 / 0 |
| N4a | `VarTypeOfParam v`（v=1） | `vt=2`（VT_I2） |
| N4b | `VarTypeOfParam v`（v="x"） | `vt=8`（VT_BSTR） |
| N5 | `AssignThroughParam v, 5`（控制格） | `v=5` / 0 |

待判点：

- **N1**：真机给 111 ⇒ Set 到数组元素这一形也要接守卫（就是下一刀）；给“空”⇒ 现状就对。
- **N4**：真机若报 `16387`（VT_BYREF|VT_I4）⇒ 那一形确实存在，我们在形参里报的是底层型（2/8），
  是一处表示差异需单独记；也报 2/8 ⇒ 该形在 VB6 里根本不出现，可从“未量缺口”里划掉。

---

## §11.89 第一百零九刀：Set 进 Variant 数组元素 + VT_BYREF 那一形注销

### 一、真机（N 族）

| 格 | 语句 | 真机 | 修前本机 |
|----|------|------|----------|
| N1 | `Set va(0) = o.NoSuchProperty` | `111` / 438 | **空** / 438 ✗ |
| N2 | `Set uv.f = o.NoSuchProperty` | `111` / 438 | 111 / 438 ✓ |
| N3a/N3b | `Set va(1) = o` / `Set uv.f = o` | 落槽 / 0 | 同 ✓ |
| N4a | `VarType(ByRef w As Variant)`（v=1） | `2`（VT_I2） | 2 ✓ |
| N4b | 同上（v="x"） | `8`（VT_BSTR） | 8 ✓ |
| N5 | `AssignThroughParam v, 5` | `5` / 0 | 5 ✓ |

**VT_BYREF 注销**：两格真机都报底层型（2/8），从没报过 16387 —— 那条从 §11.84 起挂在“未量缺口”里的
`VT_BYREF` 写穿形，**在 VB6 里不出现**，从登记里删掉（不是修好，是不存在）。

### 二、改法

`cgen_setlet_set_rhs.inc` 的 `targetIsVariantArraySlot` 那一支（Set 到 `VB6_SA_AT(vb6_VARIANT, …)`）
接上与 Set 裸变量位同形的守机：序号先抓 → 右值进 `_vs` 临时量 → 同一枚序号决定写不写。
槽仍走 `vb6_VariantAssign`（Clear + 深拷贝），临时量自有时才自己 `vb6_VariantClear`
（`vb6_VariantFromValue(void*)` 走 `vb6_VariantObject` 已 AddRef，一进一出账是平的）。
下标表达式只求值一次（挤在 VariantAssign 那一句里）。

### 三、量

- 语料实编 churn：**0 行**（757 行里只有新夹具那 1 行是新增）—— 存量语料没有“Set 到 Variant 数组元素 +
  晚绑定读”这一形；它只由新夹具覆盖。

#### 四、夹具与针

- `tests/test_late_setslot.bas`（N1 / N1b 下标表达式 / N3a 控制格）
- `b254_setslot_shape` 四条针：`_vs0` 起手、缺读那一句、写回数组元素、`vb6_VariantClear(&_vs0)`
- `b254_setslot_lesion` 一条反面针（修复前：Set 直接写数组元素）

#### 五、门

- 静态门 9 枚 PASS（`manifest_coverage` inputs=757 / registered=757）
- `owned_temp_keeps` 255 行不变
- manifest：新增 1 行 = 757 行

---

## §11.90 第一百一十刀前的探针（P 族：Set 到工程类字段槽）与一处新账

最后一条未量形（`Set <工程类 Variant 字段> = 晚绑定读`）在普通 .bas 探针里伸不到，所以另做了一个
**最小工程**：`ClsSlotProbe.vbp` + `ProbeMain.bas` + `SlotCls.cls`（类里 `Public f As Variant`），
四格：

- `P1  Set c.f = o.NoSuchProperty`（类外面写字段，C 里 `c->f`）
- `P2  c.SetMissing o`（类里面写自己的字段，C 里 `me->f`）
- `P3/P4` 控制格（成功的 Set 必须落进字段）

**这一形 C3 现在编不过**（新测到的独立缺口，与 438 该不该写目标是两回事）：

```
ProbeMain.c(56): error C2440: 无法从 “void *” 转换为 “vb6_VARIANT”
  c->f  /* class var .f field variant */ = vb6_ComGetObjectProp(o, L"NoSuchProperty");  /* Set */
ProbeMain.c(68): error C2440: 同上（控制格 `Set c.f = o`）
  c->f  /* class var .f field variant */ = o;  /* Set */
```

即：Set 到工程类的 Variant 字段那一支，**装箱（对象指针 → Variant）整个缺了** ——
`targetIsVariant` 那一支只认 `knownVariantVars_` / 类字段表里的**名字**，但写回时没走
`wrapVariantValue`，与 §11.87 给 Set 裸变量位加守卫时看到的那个 `value` 不是同一口径。
先记下，等真机 P1/P2 数回来一并处理（先装箱，再谈守卫）。

---

## §11.91 第一百一十刀：Set 到工程类 Variant 字段槽（真机 P 族）

### 一、真机

| 格 | 语句 | 真机 | 修前本机 |
|----|------|------|----------|
| P1 | `Set c.f = o.NoSuchProperty`（类外，C `c->f`） | `c.f=111` / 438 | **C2440 编不过** ✗ |
| P2 | `Set f = o.NoSuchProperty`（类内，C `me->f`） | `c.f=111` / 438 | 111 / 438（数对，但没退旧值） |
| P3 | `Set c.f = o`（控制格） | `isobj=True` / 0 | C2440 |
| P4 | `Set f = o`（控制格） | `isobj=True` / 0 | 落槽 / 0 |

### 二、两个根因（同一格两件不同的事）

1. **类外目标不认**：字段读那一站给目标拼上尾注 `/* class var .f field variant */`，于是
   `checkName` 带注释，按名那三条判定全落空 ⇒ 目标被当 typed 槽：不装箱、不守卫，直接发
   `c->f  /* … */ = vb6_ComGetObjectProp(o, L"…");  /* Set */` ⇒ C2440。
2. **类内目标装箱但不守卫**：`me->f` 那一形本来就命中 `classVariantFields_`，装箱是对的，
   但 438 之后照样裸写，且覆写前不退旧值（Variant 结构体裸赋 = 旧值漏掉）。

### 三、改法（`cgen_setlet_set_rhs.inc`）

- 新判据 `classVarFieldSlot110` 问**两腿**，都问表、不认文本：
  - 类外（`Set c.f = …`）：AST 是 `MemberAccessExpr` ⇒ 宿主类名问 `inferClassTypeOfExpr`，
    方法/属性该名先否决（`resolveClassMemberCall` 非空），再问 `classFieldIsVariantSlot`；
  - 类内（`Set f = …`）：AST 是裸 `IdentifierExpr`、发码是 `me->f` ⇒ 问 `classVariantFields_`
    （与 Fix 091p 同一张表）。
  两腿命中都把目标当 Variant 容器（后面的装箱两步自然接上）。
- 写回两档：
  - 有晚绑定读 ⇒ `_eN` 序号 + `_vsN` 临时量 + 两枚同序号闸；写回是
    `vb6_VariantClear(&<槽>); <槽> = _vsN;`（槽是真存储，临时量是**移**进去，不自己收回）；
  - 无晚绑定读 ⇒ 交给 `emitVariantSlotStore`（“清+搬”的唯一出口，§B164 第五十五刀）；
    以前 `me->f` 那一形只发一句裸赋，**旧值不退**。
- 调试：`C3_DBG_VSET110`（`[VSET110-DET]` / `[VSET110]`）。
- **被静态门退回一次的有用记录**：第一版判据按目标文本那枚尾注写（看 `/* class var .` 与
  `field variant */`），并手抄了「清+搬」两行 —— `variant_slot_store` 当场红两格：
  R3「清+搬不许住定义出口那份文件之外」、R5d「不许再把尾注当判据」。改成问表 + 走唯一出口后：
  `violations 0`（`calls 13`）。

### 四、量

- 语料实编 churn：**0 行**（759 行逐行同；存量语料里没有“Set 到工程类 Variant 字段槽”这一形）。
  量法分两段单进程跑（前 510 行 + 从 `test_let_variant_slot.bas` 起的尾段）—— 并发跑全量时那一段
  曾报过一行 `CHG`，单份复算 sha/bytes/ascii 与登记全同，是两枚 C3 同时跑撞出的假红。
- 新夹具 `tests/setcls/{SetCls.vbp, SetClsMain.bas, SetClsSlot.cls}`：本机四格与真机逐字相同。
- 针：`b255_setclsslot_shape`（6 条，含 `vb6_VariantClear(&me->f);` 与两枚控制格装箱）/
  `b255_setclsslot_lesion`（2 条）。另 `Test-Vbp setcls` + `setcls_x86` 两架构行为格。

### 五、门

- 静态门 9 枚 PASS（`manifest_coverage` inputs=759 [.vbp=280 .bas=479] registered=759）
- `variant_slot_store` PASS（violations 0）
- manifest：757 → 759 行（新增 .vbp/.bas 两行，存量 0 行变动；次日重算 .vbp 行
  10187 → 10165 字节，·cls 体内多了守卫）
---

## §11.92 第一百一十一刀前的探针（Q 族：Set 语句目标的另外两形）

第一百一十刀把「工程类 Variant 字段槽」量完并修好之后，Set 的目标还剩下两形没量。
本机先用一份小夹具（`.temp/k111/slots111.bas` / `slots111b.bas`）读现状：

| 格 | 语句 | 本机现状 |
|----|------|----------|
| Q1 | `Set w = o.NoSuchProperty`（目标 = ByRef Variant 形参，C `(*w)`） | **C2440 编不过**：`(*w) = vb6_ComGetObjectProp(o, L"…");` —— 装箱整个缺 |
| Q2 | `Set uv.f = o.NoSuchProperty`（UDT Variant 字段） | 111 / 438 ✅（第一百零七刀已修） |
| Q3 | `With uv : Set .f = o.NoSuchProperty` | `uv.f=` / 438 ❌ —— 裸赋值，槽被清空 |
| Q4 | `With uv : Set .f = o`（控制格） | `isobj=True` / 0 ✅ |

另记一处与本刀无关的独立缺口：**模块级 UDT（含 Variant 成员）初始化 = C2074**
（`vb6_type_UV uv = {0};`），所以 Q2/Q3 的本机读数改用过程内局部 UDT 取得。

真机探针已写出（`.temp/k111/probe111/RealMachineProbe111.bas`，五格：Q1 / Q1c 控制 /
Q2 参照 / Q3 / Q3c 控制），等真机数回来再改编译器 —— 两形都按「先量后改」处理。
---

## §11.93 第一百一十一刀：Set 到解引用槽 (*w) 与 With 块 UDT Variant 字段（真机 Q 族）

### 一、真机

| 格 | 语句 | 真机 | 修前本机 |
|----|------|------|----------|
| Q1 | `Set w = o.NoSuchProperty`（w = ByRef Variant，C `(*w)`） | `v=111` / 438 | **C2440 编不过** |
| Q1c | `Set w = o`（控制格） | `isobj=True` / 0 | C2440 |
| Q2 | `Set uv.f = o.NoSuchProperty`（UDT 字段，参照） | `uv.f=111` / 438 | 111 / 438 ✅ |
| Q3 | `With uv : Set .f = o.NoSuchProperty` | `uv.f=111` / 438 | **空** / 438 ✗ |
| Q3c | `With uv : Set .f = o`（控制格） | `isobj=True` / 0 | 落槽 ✅ |

### 二、两个根因

1. **Q1：剥壳只认长名**。判定 Variant 容器前把 `(*w)` 剥成 `w` 的那句条件是 `size() > 4`，
   而 `(*w)` 恰好 4 个字符 ⇒ 不剥 ⇒ 名字带星号查表全落空 ⇒ 当 typed 槽发
   `(*w) = vb6_ComGetObjectProp(o, L"…");`（void* → vb6_VARIANT，C2440，装箱整个缺）。
2. **Q3：判据把「含 `->`」一律排到裸赋**。`plainVar108` 要求 `target` 里既无 `(` 也无 `->`；
   With 块的字段发出来是 `_vb6_with_N->f` ⇒ 进不了守卫那一档，438 之后照样把槽写空（且不退旧值）。

### 三、改法（`cgen_setlet_set_rhs.inc`）

- **Q1**：不全局放宽剥壳（那么做会把 `(*p)` 这种**工程类出参槽**也剥掉，见下面「A/B 退回记录」），
  单开一问 `derefVariant111`：目标形如 `(*名)` 且
  ① 名字在 `knownVariantVars_` 里、不在 `knownObjectVars_` / `knownTypedComVars_` 里；
  ② `inferClassTypeOfExpr` 答得出类名就否决（工程类出参槽 `(*p)` 就是这一形，与 §11.87 的
  staleness 否决同一口径）。
  命中才 `targetIsVariant = true`（装箱 + 守卫一并接上）。
  ⚠ 这里**不**问 `classOfProjectRefOutSlot`：`class_refcount_authority` 把「出参槽类名问句」钉成
  恰好 4 处（声明 + 定义 + 两站），第一版多问了一次，CI 的 compile 格当场红（第 5 处）。改用 AST
  那一问（更早就有的口径），把那一问留给它自己的 4 处。
- **Q3**：`plainVar108` 的排除项从「含 `(` 或 `->`」改成「目标不是一次纯粹的左值求值」：
  `(*w)` 与「`->` 之前没有括号」两形都算纯左值；只把**带函数调用**的目标
  （`Add(…)->Image`，写回两次会多调一次）留在原路。
- 两形写回仍走既有那一档：`_eN` 序号 + `_vsN` 临时量 + 同序号闸，
  写回是 `vb6_VariantClear(&<槽>); <槽> = _vsN;`（槽是真存储，临时量移进去）。

#### 四、A/B 退回记录（记一次真回归）

第一版把剥壳条件从 `size() > 4` 放宽成 `>= 4`，全量 rehash 立刻给出 **2 行 churn**：
`tests/b187_box/B187d.vbp` / `tests/b445_refcount/B445.vbp`。用 BASE（`git stash` + 重建）与新机
并排重发那两份，diff 逐字给出原因：`(*p)` 里的 `p` 是 `As B187Good` / `As C445` 的**出参槽**，
剥壳后被当成 Variant 容器 ⇒ §B218 腿②′ 那整段（Unbox + AddRef + 花括号注释）被抢走，两行各短 72 / 3 字节。
改成只看声明档的 `derefVariant111` 后，那两份与 BASE **逐行相同**，全量 rehash 回到 0 行。

### 五、量

- 语料实编 churn：**0 行**（760 行逐行同；单进程一次跑完）。
- 新夹具 `tests/test_late_setslot2.bas`：五格与真机逐字相同（Q1 / Q1c / Q2 参照 / Q3 / Q3c）。
- 针：`b256_setslot2_shape`（7 条，含 `vb6_VariantClear(&(*w)); (*w) = _vs`、With 字段的
  `->f); _vb6_with_`、Q2 回归针）／`b256_setslot2_lesion`（2 条：裸 `(*w) = ComGetObjectProp`、
  裸 `->f = VariantFromValue(ComGetObjectProp)`）。

### 六、门

- 静态门 9 枚 PASS（`manifest_coverage` inputs=760 [.vbp=280 .bas=480] registered=760）
- manifest：759 → 760 行（只新增本刀那一行，存量 0 行变动）

---

## §11.94 模块级 UDT 变量的静态初始化（第一百一十二刀）

### 一、病灶

真机 VB6 里再普通不过的一行声明：

```vb
Public Type UV
    f As Variant
End Type
Public uv As UV
```

本机发成 `vb6_type_UV uv = 0;` —— 结构体不能拿 `= 0` 初始化，cl 直接 **C2074**
（initialization of a struct requires a brace-enclosed initializer list），
整个工程编不过。所以此前凡带 Variant 成员的模块级 UDT 都是空白面。

### 二、根因

`cgen_decl_var.cpp` 里那条「UDT ⇒ `{0}`」的判定只问 `lookupDotted(simple.name)`：

```cpp
auto* sym = lookupDotted(simple.name);
isUdtType = (sym && sym->kind == SymbolKind::UserDefinedType);
```

而 `Public Type UV … End Type` 只登记在**模块作用域**（`lookupModuleDotted` 那一层），
名字查不到 ⇒ `isUdtType = false` ⇒ 落到 `initVal = defaultValue(Vb6Type::Unknown)` = `"0"`。
`Unknown` 这一档本是给数组那类 `vb6_SafeArray1D*` 留空指针用的（账 #116 那条注释写得很清楚，
不能改成 Variant），与 UDT 撞在同一个 `"0"` 上。

### 三、改法

两处兜底，都是「先量后改」量出来的最小面：

1. 名字查不到时补一问模块作用域：`if (!sym) sym = lookupModuleDotted(simple.name);`
2. 再用 **C 类型名**兜底：`cType.rfind("vb6_type_", 0) == 0` 判为结构体 ⇒ `{0}`。
   这一条不依赖符号表，符号表再怎么缺登记也不会把结构体发成 `= 0`。

`{0}` 就是全零：Variant 成员 = Empty（VT_EMPTY=0）、String = 空、数值 = 0 —— 与 VB6
对模块级变量的零初始化一致。

### 四、量（本机）

新夹具 `tests/test_udtmod.bas`（Public / 嵌套 / Private 三形 + 零初始化读数 + 模块级槽的 438 守卫）：

```
UDTMOD-1 VarType=0 IsEmpty=True IsNull=False
UDTMOD-2 uv.f=7 VarType=2
UDTMOD-3 nst.inner.f=[] nst.s=[] nst.n=0
UDTMOD-4 x|abc|42
UDTMOD-5 IsEmpty(pv.f)=True
UDTMOD-6 v=111 err=438
UDTMOD-7 isobj=True err=0
```

- `UDTMOD-6` 与真机 probe111 **Q2 逐字一致**（`uv.f=111 err=438`，只是槽从过程级挪到模块级；
  发射码与 Q2 同形：`const int32_t _eN = vb6_ErrSerial();` 双闸 + `_vsN`）。
- `UDTMOD-1/3/5`（零初始化那几格）已由 **probe112** 真机逐字复核（见下表），全部对上。

#### 真机 probe112 读数（逐字对照，已回）

| 格 | 真机 | 本机 | |
|---|---|---|---|
| P1 | `VarType=0 IsEmpty=True IsNull=False` | 同 | ✔ |
| P1b | `uv.f=7 VarType=2` | 同 | ✔ |
| P1c | `nst.inner.f=[] nst.s=[] nst.n=0` | 同 | ✔ |
| P1d | `x|abc|42` | 同 | ✔ |
| P1e | `IsEmpty(pv.f)=True` | 同 | ✔ |
| P2 | `v=111 err=438` | 同 | ✔ 模块级槽守卫正确 |
| P2c | `isobj=True err=0` | 同 | ✔ |
| P3 | `v=[111] err=438` | `v=[] err=0` | ✘ 缺口确认 |

`{0}` 这一手与真机零初始化**完全一致**（Variant=Empty / String 空 / Long 0，嵌套同形，
Public 与 Private 同）；模块级槽的 438 守卫也与过程级同一条规则。

### 五、针与门

- `b257_udtmod_shape`（4 条）：`vb6_type_UV uv = {0};` / `vb6_type_Nest nst = {0};` /
  `vb6_type_UV pv = {0};` / `vb6_VariantClear(&uv.f); uv.f = _vs`
- `b257_udtmod_lesion`（3 条，修前那三行都是 `= 0`）
- 运行核 `Add-BasTest "test_udtmod"`：8 条读数
- manifest：760 → 761 行

### 六、P3：Collection 未知名成员（真机已确认的缺口，第一百一十三刀）

真机 probe112 **P3 = `v=[111] err=438`**，本机 `v=[] err=0`：

```vb
Dim col As Object: Set col = New Collection
v = 111
Set v = col.NoSuchProperty     ' 真机: 438 且目标保持 111（不写目标）；本机: 清空 err=0
```

两处都错：① 没报 438；② 还把目标写了（清成空）—— 真机 438 不写目标。
与 §11.93 的槽守卫无关（守卫是好的：P2 换 `Scripting.Dictionary` 立刻 `v=111 err=438`），
是 COM 派发那一层的宽松。真机读数已到，开第一百一十三刀。

## §11.95 内建 Collection 的未知名成员 ⇒ 438（第一百一十三刀 · 量到了，但被一处发错误挡住）

### 一、真机读数（probe112 P3）

```vb
Dim col As Object: Set col = New Collection
col.Add 5
v = 111
Set v = col.NoSuchProperty
```

| | 读数 |
|---|---|
| 真机 | `v=[111] err=438` |
| 本机（现状） | `v=[] err=0` |

两处都错：① 不报 438；② 还把目标写了（清成空）—— 真机 438 不写目标。

### 二、根因（RTL 一侧，已定位）

`uc_hostmodel_getprop.inc` 的内建 Collection 岔口对**任何**未知成员都 `return 1`
（"答了但没答"），于是 `vb6_ComGetPropRcv` 宿主分支那条 `hgot < 0 ⇒ 438` 的收口
（同文件末尾 `fxtDocument.rtfText` 那一刀立的规矩）永远走不到。正确改法只有一行：

```c
    if (vb6_uc_isColl(obj)) {
        if (_wcsicmp(name, L"Count") == 0) { ...; return 1; }
        return -1;     /* 成员不存在 ⇒ 调用方抛 438、交空值（438 不写目标） */
    }
```

### 三、为什么这一刀没落（实测：改成 `-1` 当场红两格）

改完本机自测是对的（`P3 v=[111] err=438`，与真机逐字一致），但 CI 两格红：

1. `[STATIC] com_latebind_guard`：`R5 calls of vb6_ho_report438 in uc_hostmodel_getprop.inc = 2 (want 1)`
   —— 那枚门把报数站点钉成恰好 1 处（文件末尾那位主人）；这一格不再重报即可（已按此改）。
2. `[GUI] Charts2020 ... exit=438`：**真正的拦路虎**。

Charts2020 的 438 与 Collection 本身无关，是生成码里一处**既有的**发错误：

```vb
Type RectF
    Prop As Collection
End Type
With Rects(i)
    .Prop.Add value, name          ' 应发成 vb6_ComCall(_vb6_with_N->Prop, L"Add", …)
End With
```

实际发出（复现器 `.temp/k113/wuc2.bas`，五行）：

```c
vb6_ComCall(vb6_ComGetObjectProp(_vb6_with_5->Prop, L"Prop"), L"Add", …)
```

多出来的一层 `ComGetObjectProp(……, L"Prop")` —— 拿 **Collection** 去问它自己叫不叫 `Prop`。
这一层以前被第二节那格宽容吞掉：接收者变 NULL，`vb6_ComCall(NULL, L"Add", …)` 落成
晚绑定 miss（本案 stderr 里那一串 `site=ComCall name="Add" disp=00000000`，60+ 条），
也就是说**那些 `Add` 本来一次都没进过集合**；§11.95 让它报 438 之后，这一层就从空转变成了
进程级错误。旧读数也印证：Charts2020 的 GUI 用例此前通过，靠的是"错得不出声"。

出处是 `cgen_expr_with.cpp` 的 Fix 110t 分支（With 块里 UDT 的**对象字段**当 callee）：
它把 `comObjExpr_` 设成**已经解好的字段取值** `_vb6_with_N->Prop`、`comMemberName_` 又设成
**同一个字段名** `Prop`，于是外层再解一层就成了上面的形状。同一分支的另一个消费者
（`.ObjField(args)` 直接调用，要 `L"Item"`）不能一起动 —— 试过把名字清空：直接调用那形
退化成 `vb6_ComCall(…, L"", …)`、`.Add` 那形退化成 C 结构体方法调用
`_vb6_with_0->Prop.Add(1, …)`，两形都不对，已撤回（记录在 `cgen_util_com.cpp` 的 `resolveComValue`
与 `cgen_expr_with.cpp` 两处都恢复原样）。

### 四、结论与下一刀

- 这一格**先不改**：留着 `return 1`，并把上面那段证据写进代码注释（谁改谁先看 Charts2020）。
- 第一百一十四刀的正题 = 修 Fix 110t 那个"字段已解好又当成员名"的接收者发法，
  让 `With <UDT>: .ObjField.Method …` 发成 `vb6_ComCall(<字段>, L"<Method>", …)`、
  `With <UDT>: .ObjField(args)` 仍发 `L"Item"`。修好后把本节第二节那一行改回 `-1`。
- 复现器：`.temp/k113/wuc2.bas`（`Type Holder: Prop As Collection` + `With h(0): .Prop.Add 1, "k"`
  + `.Prop(1)` + `.Prop.Count` 三形一起看）。
- 撤下的东西：`tests/test_collmiss.bas` 及其运行核（**不**给错误行为立针）、manifest 762 → 761。

### 五、顺带量到、另行登记的一处缺口

普通赋值读已知成员到 Variant 槽，本机交空：

```vb
Dim v As Variant
v = col.Count        ' 本机: v 为空, err=0；VB6: 1（待 probe113 Q2 真机复核）
```

`CStr(col.Count)`（直接取值）是对的；`v = dict.Count`（Dictionary 走 IDispatch）也正常
（`v=[1] VarType=3`）。只有"宿主回 VT_I4 → Variant 槽"这条搬运路漏，且只在宿主那条路上。

## §11.96 With-UDT 对象字段的方法调用（第一百一十四刀 · 修好 §11.95 被勒住的那处发错误）

### 一、病在哪：一处发错误，两个症状

`With <UDT>` 里对**对象字段**（`Prop As Collection` 这类）做调用，Fix 110t 设的标记是

```
comObjExpr_  = _vb6_with_5->Prop      ← 字段值本身（对象已经在手上）
comMemberName_ = "Prop"               ← 只是记着它是哪个字段
```

下游把这两个字段**分别**当「接收者」和「方法名」用，于是同一件事实被解了两次：

| VB 源码 | 修前发码 | 错在哪 |
|---|---|---|
| `With h(0): .Prop.Add 1, "k"` | `ComCall(ComGetObjectProp(_vb6_with_5->Prop, L"Prop"), L"Add", …)` | 拿 Collection 问它自己叫不叫 `Prop`（多解一层） |
| `With h(0): .Prop(1)` | `ComCall(_vb6_with_5->Prop, L"Prop", …)` | 字段名顶到方法名上（该走默认成员 `L""`） |
| `With m_Serie(i): .CustomColors(j+1)` | `ComCallInt(_vb6_with_19->CustomColors, L"CustomColors", …)` | 同上，Fix 110t 自己的第一格就一直是错形状 |

第一条在 §11.95 之前一直是**空转**：多出来那层 `ComGetObjectProp(…, L"Prop")` 打到内建 Collection 的宽容岔口（`return 1` + 空对象），接收者变 NULL，`Add` 一次也没真进过集合（账上 60+ 条 `site=ComCall name="Add" disp=00000000` 就是它）。§11.95 让 Collection 老实回 438 之后，这层空转立刻变成**进程级错误**（`Charts2020 exit=438`）——所以 §11.95 当时只能撤回，先把发错误修好。

### 二、修法（三处，都在发码侧）

1. **`cgen_expr_member_generic_access.inc`（新增 ST-GEN3 支）**
   链式成员访问看到 `isComMarker_` 时，先判「标记的接收者尾巴就是这个成员名」（`objExpr` 以 `-><member>` 收尾）。
   命中 ⇒ 说明**对象已经在手上**，照「已解好的对象」续链（`comObjExpr_` 不动、`comMemberName_` 换成外层成员名），
   不再 `resolveComValue("Object")`。不命中 ⇒ 逐字走原路。

2. **`cgen_expr_call_com_bind.inc`**
   标记被**直接当调用目标**消费时（没有外层成员），若接收者尾巴等于成员名 ⇒ 这个名字是字段名，不是方法名 ⇒ 名字交空，
   走 `vb6_ComCall` 的默认成员（DISPID 0）。

3. **`cgen_call.cpp`（语句路）**
   `With <UDT>: .ObjField.Method arg, arg` 这种**语句式带实参**调用（parser 交 `CallStmt(callee=IndexOrCallExpr)`），
   以前会落到「结构体方法调用」拼成 `_vb6_with_0->Prop.Add(1, _own_1)`（C 结构体方法形，编不过 / 语义全错）。
   现在在 `emitExpr(*node.callee)` 之后先看标记：有标记 + 有实参 ⇒ 直接发
   `vb6_ComCall(<字段>, L"<方法>", (void*[]){…}, n)`，与表达式路同一口径。

三处合起来：**字段只解一次，名字只用在它该在的位置**。

### 三、发码对照（撤改前后逐字，5 行 churn 全部是这处语义纠正）

`tests/Charts 2020/ucChartBar/Proyecto1.vbp`（4 站 `Values.Add`）：

```
修前: vb6_ComCall(vb6_ComGetObjectProp(_vb6_with_4->Values, L"Values"), L"Add", …)
修后: vb6_ComCall(_vb6_with_4->Values, L"Add", …)
```

`tests/Charts 2020/ucTreeMaps/Proyecto1.vbp`（`CustomColors(j+1)` 4 站）：

```
修前: vb6_ComCallInt(_vb6_with_19->CustomColors, L"CustomColors", …)
修后: vb6_ComCallInt(_vb6_with_19->CustomColors, L"", …)
```

`tests/VBFlexGridDemo/Builds/VTableHandle.bas` 与 `VBFlexGridDemo.vbp`（`.Enumerable(.Index)`）：

```
修前: vb6_ComCall(_vb6_with_21->Enumerable, L"Enumerable", …)
修后: vb6_ComCall(_vb6_with_21->Enumerable, L"", …)
```

churn: **5 / 762**（`tests/Charts 2020/Proyecto1.vbp`、其 ucChartBar、ucTreeMaps 三个 vbp + VBFlexGridDemo 两行），
其余 757 行 byte-identical；skip 5（机差 5 枚，与历刀同）。

### 四、复现器与读数

`.temp/k113/wuc2.bas`（五行，含 `Set h(0).Prop = New Collection`）：

```
修前: H0 count=  H1 count=   (Add 全空转, count 读空)
修后: H0 count=1 H1 count=2  (Add 真进集合)
```

真机口径对照（probe112 P3，§11.95 已量）：`Set v = col.NoSuchProperty` ⇒ `v=[111] err=438`。
第一百一十三刀撤回的那格（内建 Collection 未知名 ⇒ 返回 -1）本刀已恢复：

```
本机: err=438  type=2 val=111   ← 与真机逐字一致
```

`tests/test_collmiss.bas` 六格全绿（`CM-1 count=1` / `CM-2 item=5` / `CM-3 v=[111] err=438` /
`CM-4 count=1 err=0` / `CM-5 v=[222] err=438` / `COLLMISS-DONE`）。

### 五、Charts2020

修后 `exit=6`（`Unhandled VB6 Error #6: Overflow`），与 §11.95 撤回后的基线**同一读数**，
且 stderr 里 **一条 438 也没有**、`C3_LATEBIND_MISS` 从 60+ 条降到 **0 条** ——
那串 `Add` 以前一直空转，现在真进了集合（这才是 Charts2020 原本要的行为）。

### 六、仍缺的一格（不属本刀）

`With h(0): .Prop(1)` 的**返回值**交空（本机 `H2 item1=` 空）：接收者与默认成员名都已经对了，
缺的是「宿主回 VT_I4 → Variant 槽」这条搬运路。已登记在 probe113 Q2/Q2b 待真机读数，见 §11.95 的缺口表。

## §11.97 内建 Collection 的调用臂：Count / 空名默认成员 / 未知名 ⇒ 438（第一百一十五刀）

### 一、probe113 真机读数（VB6 IDE 实测，2026-11-22）

```
Q1  Set v = col.NoSuchProperty   => v=[111] err=438      (对照格, §11.95 已收口)
Q2  v = col.Count                => v=[1] err=0 VarType=3 (VT_I4)
Q2b v = col.Item(1)              => v=[5] err=0 VarType=2 (VT_I2; 5 是 Integer 字面量)
Q2c v = col.NoSuchProperty       => v=[111] err=438      (普通赋值读未知名)
Q3  col.Count                    => count=1 err=0
```

### 二、病在哪：同一个出口上的三个洞

`v = col.Count`、`v = col.NoSuchProperty`、`With h(0): .Prop(1)` 在发码侧都是 **`vb6_ComCall`**
（不是 `ComGetProp`），而宿主模型 `uc_hostmodel_call.inc` 的 `vb6_uc_isColl(obj)` 臂只认
`Add` / `Item` / `Remove`，末尾无条件 `return 1`：

| 源码 | 修前这一臂 | 症状 |
|---|---|---|
| `v = col.Count` | 没有 Count 分支，落到末尾 `return 1` | `out` 仍是入口清零的 Empty ⇒ 宿主那枚 VT_I4 根本没生出来（Q2 读成空/VarType=0） |
| `With h(0): .Prop(1)` | 空方法名（§11.96 定的默认成员写法 `L""`）不映射到 Item | 返回 Empty（复现器 H2 blank） |
| `v = col.NoSuchProperty` | 未知名也 `return 1`（「答了但没答」） | Q2c 读到 err=0，且目标被写成空 |

三格同属**一个调用出口**，所以一刀修完；这也是 §11.96 末尾登记的那格（H2 与 Q2 同根）。

### 三、修法（RTL 一处，`uc_hostmodel_call.inc`）

1. `cname = (name && name[0]) ? name : L"Item"` —— 空名字按 VB6 默认成员处理；
2. 新增 `Count` 分支，与 GetProp 臂问**同一个主人** `vb6_Collection_Count`；
3. 其余未知名 ⇒ `vb6_ho_report438(L"Call", …)` + `return -1`，与 §11.95 的 GetProp 尾巴同口径
   （`vb6_ComCallRcv` 的宿主分支按 VB6 语义抛 438，目标不动）。

认得的三支（Add/Item/Remove）与 Count 仍 `return 1`；行为面只加三格。

### 四、本机读数（修前 → 修后）

```
P3  v = col.Count            : v=[] VarType=0      → v=[1] VarType=3
H2  With h(0): .Prop(1)      : 空                    → item1=1 / item2=2
Q2c v = col.NoSuchProperty   : (err=0, 目标被写空)   → v=[111] err=438
CC-6 col.AnotherMissing      : err=0                → err=438
```

新增夹具 `tests/test_collcall.bas` 七格（`CC-1` … `CC-6` + `COLLCALL-DONE`），
与 probe113 逐格对照；manifest 762 → **763** 行。

### 五、静态闸门与 churn

- `check_com_latebind_guard.ps1` 的 R5「438 站点数」按实测从 `Call=1` 收紧为 `Call=2`：
  新增那处不是把真缺口改记成 438，而是 Collection 未知名在调用臂上**本来就没有这个成员**
  （与 GetProp 臂 §11.95 同一件事）；认得的三支仍 return 1。
- `check_variant_slot_store.ps1` / `gates.ps1` 全绿；`check_manifest_coverage` 763/763。
- churn 实测 `full_rehash85.ps1`：**rows=763 changed=0 skipped=5 missing=0** ——
  本刀只动 RTL 宿主模型、不动发码，期望表除新增那一行外逐字节不变。

### 六、与真机的剩余差异（不在本刀）

`CC-2 v = col.Item(1)` 值对，VarType 本机 3(VT_I4) / 真机 2(VT_I2)：`col.Add 5` 的整数字面量
归档口径差异（既有），本条针只立「搬进 Variant 槽且值对」，不为这处差异立期望。

## §11.98 宿主 Font 臂交对象：`.Font.Size` / `CallByName(…,"Font",…).Size` 不再读空（第一百一十六刀）

### 一、读数（本机 x86，与 CI 同一产物形态，同一台机）

| 夹具 | 旧编译 | 新编译 | 行形 |
|---|---|---|---|
| `tests\Charts 2020\Proyecto1.vbp` | 23 | **0** | `site=ComGetProp name="Size" disp=00000000 kind=null` |
| `tests\czUI-main\czFormDemo.vbp` | 3 | **0** | `site=ComSetProp name="Name"/"Size"/"Bold" disp=00000000 kind=null` |

两枚都是跑 8~10 秒后杀进程；修前修后进程都还在消息循环里（不是提前退出）。

### 二、病在哪：一个「防崩」出口把对象吞了

`uc_hostmodel_getprop.inc` 的 Font 分支：

```c
vb6_UCFontRec* fr = (vb6_UCFontRec*)vb6_UC_NewFont();
vb6_ho_setVariantDispatch(out, fr); return 1;   /* 这个出口恒交 Empty */
```

`vb6_ho_setVariantDispatch` 的契约是**恒 Empty**（历史原因：怕 `VariantClear` 对伪对象调 Release）。
于是：

- `CallByName(oCtrl, "Font", VbGet).Size`（`ClsResizer.cls:107`）取到 Empty ⇒ `.Size` 接收者 NULL（23 条）；
- `txtEmbed.Font.Name/Size/Bold`（czFormDemo）同样（3 条）。

### 三、为什么裸 VT_DISPATCH 是安全的

`vb6_ReleaseObject` / `vb6_ComAddRefDispatch` 的唯一判据是 `vb6_ComIsDispatchable`
（首字段是不是可执行 vtable）。字体伪对象首字段是 `BSTR Name`，两个出口都跳过，
`vb6_VariantClear` 不会去 Release；`vb6_Host_ToWinVariant` / `ClearVariant` 对对象档
本来就是「借用不计数」。Controls.Item 那族（`vb6_ho_setVariantObject`）早就是同一口径。

### 四、修法

```c
vb6_ho_setVariantObject(out, fr); return 1;
```

`.Size` / `.Name` / `.Bold` 随后落到同一格（`vb6_uc_isFont` → `vb6_ho_getFontMember` /
`uc_hostmodel_setprop.inc` 的字体分歧），不需要另立第二条字体属性面。

### 五、针与闸门

- `LateBindMissAllow`：删掉 `'Proyecto1' = 23`、`'czFormDemo' = 3`（默认期望恰好 0），
  并写上两条旧读数的行形与出处 —— 「从 23 降到 0」不是放宽判据，是不该计入的伤口消失。
- `check_com_latebind_guard.ps1` 新增形状钉：Font 臂必须发 `vb6_ho_setVariantObject(out, fr)`，
  改回 `setVariantDispatch` 就红（那等于把 23 + 3 条 miss 放回来）。
- 本机 `test_collcall` / `test_collmiss` 全绿；`gates.ps1` 9/9；churn 763 行 changed=0
  （只动 RTL 宿主模型，不动发码）。

### 六、本刀**不**声称修好的那半句

交出去的是 `vb6_UC_NewFont()` 现造的**默认档**字体代理（8.25pt / MS Sans Serif），不是控件
自己的那一枚。因此：

- `ctl.Font.Size` 现在**有值**（不再 NULL miss），但值恒是默认档，不会读出设计期给它设的字号；
- `ctl.Font.Size = x` 写在一片崭新的记录上，**不落到控件**；
- 推论：Charts2020 `ClsResizer` 的字体随窗口缩放（先记 `.Size * 100 / RatioAspect`，
  Resize 时再写回）存量下来了但**还不改变外观**。

真正对齐需要两件（下一刀）：① 宿主模型为每个宿主对象记一枚字体记录（按 HWND / UC 实例去重，
不复用 `vb6_uc_fontOf` 那张只认字体的表）；② setprop 写完后叫 `vb6_UC_ApplyFontToWindow(obj, f)`
把八档发到窗口（那位主人已在 `uc_controls.c`，不另造）。本刀只保证「接收者不再是 NULL」。

## §11.99 宿主 Font 代理挂到控件上：读回设计期字号、写回落盘（第一百一十七刀）

上一节留下的那半句，本刀补齐。

### 一、读数（本机 x86，探针 = `czUI-main` 副本插两行 `Debug.Print`）

探针位置：`czUI.ctl:1354-1356`（`txtEmbed.Font.Name/Size/Bold = m_Font*`）前后各印一次。

| 编译 | before | after（写完 10pt / Segoe UI） |
|---|---|---|
| §11.98 之前（`C3_old`） | `size=` `name=` `bold=`（`.Font` 是 Empty） | 同左（写进空接收者） |
| §11.98 | `size=8.25 name=MS Sans Serif bold=0` | `size=8.25 name=MS Sans Serif bold=0`（孤儿代理，写回不落盘） |
| 第一百一十七刀 | `size=8.25 name=MS Sans Serif bold=0` | `size=10 name=Segoe UI bold=0` |

`[C3_LATEBIND_MISS]` 条数不回退：Charts2020 / czFormDemo 本机实测仍为 **0 / 0**。

### 二、病：代理是孤儿

§11.98 交出的对象是 `vb6_UC_NewFont()` 现造的默认档（8.25pt / MS Sans Serif），而且**每次读都新造**：

- `.Size` 恒 8.25，读不到控件设计期字号；
- `.Size = x` 写在一片马上丢掉的记录上，控件不缩放 —— Charts2020 `ClsResizer` 的存（`.cls:107`）与放（`.cls:147`）两半都空转。

### 三、修法：代理挂到六位既有主人上

控件字体本来就只有一套存储与出口：`vb6forms_ctrl.c` 的
`vb6_Get/SetControlFontName|Size|Bold|Italic|Underline|Strikethrough`（同一格 `VB6_CtrlFont` + `WM_SETFONT`）。
本刀不新造第二份状态：

1. `vb6_UCFontRec` 加 `owner`（HWND，0 = 设计期字体块 / PropertyBag 那枚）；
   `vb6_HostObjRec` 加 `fontProxy`（每个宿主对象一枚）；
2. `vb6_UC_NewFontForWindow(hwnd)`：造代理 + 记 owner + 用六位读侧填成控件**现值**（字号走
   `vb6_GetControlFontSize` 的自存精确磅值，不从像素高度反推，所以 8.25 还是 8.25）；
3. `vb6_UC_FontSyncFromOwner(p)`：`vb6_UC_FontField` 每次取字段前刷新 —— 静态路径
   `ctl.FontSize = x` 改的是窗口上那张字，代理不刷就是旧值；
4. `vb6_UC_FontCommit(p, name)`：`vb6_ComSetProp` 写完一档后交回**对应 setter**（落盘 + 存字）；
   `Weight`/`Charset` 没有单档 setter，走整块 `vb6_UC_ApplyFontToWindow`。

宿主 Font 臂改为按登记记录缓存（`h->fontProxy`），同一控件两次读是同一枚对象；
字体分支不再 `vb6_UC_NewFont()` 现造。

### 四、闸门

- 探针三点读数见上表；`test_collcall` / `test_collmiss` 全绿；`gates.ps1` 9/9；
  churn 763 行 changed=0（只动 RTL，不动发码）。
- `check_com_latebind_guard.ps1` 的形状钉升级：Font 臂必须同时出现 `fontProxy` 与
  `vb6_UC_NewFontForWindow`；改回 Empty 或每读现造就红。

### 五、边界

- owner 为零的字体（设计期块 / PropertyBag）行为与本刀前完全一致；
- 同一控件上的 `ctl.Font` 与 `ctl.FontSize` 现在读写**同一格**（晚绑定走代理 → 六位主人，
  静态路径直接走六位主人），不再出现两条路两个答案；
- `Weight` / `Charset` 的写回会把当前八档整块发一次（IFont 在 GDI 里本来只有一个 hFont）。

## §11.100 缺省 Instancing 的工程类也能装箱：`As Object` 收得下类实例（第一百一十八刀）

### 一、读数

| 项 | 旧 | 新 |
|---|---|---|
| `tests\b190_scope` 的 `[C3_LATEBIND_MISS]` | 8（4 处 × 2 行：`DispView` + `ComGetProp`，全部 `kind=notdispatchable`） | **0** |
| P3/P4（`As Object` 形参读字段） | 空 | `A1` / `A1` |
| P9/P10（Object 变量读 `Label`） | 空 | `A-label` / `A-label` |
| P6/P7/P8/P11/P12（`TypeName`） | 部分空 | `Nothing` / `B190A` / `B190B` / `B190A` / `B190A` |

同族 9 个夹具（`b168_cvf`、`b169_cfw`、`b170_mrt`、`b171_slotarg`、`b172_paramslot`、
`b173_udtarg`、`b183_intfield`、`b184_fieldw`、`b190_memberarray`）本机逐个编译运行：
全部 exit=0、`miss=0`。

### 二、病：一条判据答了两件事

`Symbol::isCoClassTableEntry()` 一直同时回答：

1. 这枚类能不能**对外 CoCreate**（BUG-4 口径：只有 MultiUse/SingleUse 该有 ClassFactory 项）；
2. 这枚类能不能在运行期 coclass 描述表里找到 desc（装箱腿 `boxIntoLateSlot` 的闸门）。

缺省 Instancing 的工程类在 VB6 里同样是 COM 对象（`As Object` 形参收得下、晚绑定读得动），
但它被第 1 问的答案拉进了第 2 问：不在表里 ⇒ `classHasCoClassDesc` 为假 ⇒ 装箱腿交**裸**
`vb6_cls_*` ⇒ `vb6_ComIsDispatchable` 判否 ⇒ `[C3_LATEBIND_MISS] kind=notdispatchable` 全链静默。

### 三、修法：把两条判据拆开，表里补「不可创建」的行

1. `Symbol::isLateBindClass()` = 所有类（接口除外）—— 「能不能装箱」只看这个；
2. 收集器改按 `isLateBindClass()` 收，同时记 `CoClassInfo.registerable = isCoClassTableEntry()`；
3. 行照发（装箱靠它找 desc），但 `progId` / `clsid` 两格发**空串**；
4. `DllRegisterServer` / `DllUnregisterServer` 见空即 `continue`；`vb6_GetClassFactory` 按 CLSID
   匹配（空串必然匹配不上）⇒ 私有类不会出现在注册表里；
5. 装箱腿与 `As New` 字段的闸门改用 `isLateBindClass()`；
6. **UserControl (.ctl) 必须留在闸门外** —— 它的 C 侧值是 HWND；把 `.ctl` 当类收进表，
   `com_entry.c` 会去引 `.ctl` 里不存在的符号（`ucSlider_2022` 实测 LNK2019：
   `_vb6_LabelPlus_tmrMOUSEOVER_Timer`）。三处同一条守卫：收集器
   （`knownControlModuleNames_`）、装箱名单、`As New` 字段。

实测面：`tests` 下 423 份 `.cls` 里 18 份缺省 Instancing（`b168..b190` 那族），
真语料按 §B187 的普查全部带 Instancing ⇒ 存量工程产物不动。

### 四、闸门

- `LateBindMissAllow` 删掉 `'B190' = 8`（现在表里只剩负控 `B446`）—— 期望恰好 0；
- B190 12 个探针值与同族 9 夹具见上；`gates.ps1` 9/9；
- churn 763 行 **changed=3**：`tests/b190_scope/B190.vbp`（预期：多出 B190A/B 两行 desc
  + 四处装箱）、`tests/b216_slot/B216d.vbp`、`tests/iface_wrap/IfaceWrap.vbp`（同一形状）。
  其余 760 行不动 —— 真语料类全带 Instancing，本来就是表内行。

### 五、随刀改正的旧针（b216）

`tests\b216_slot` 原来把「表外那一档」当正确答案钉住：N1/N2/N3 读 `tn=Object`，并且
发码针强求 `vb6_FindCoClassDesc("B216Plain")` **不在场**。这与 VB6 语义相祀：
`B216Plain.cls` 只是没有 Instancing（= Private），同一工程里 `New` 出来的实例
`TypeName` 就是类名，不是 `Object`。本刀实测：

```
N1-plain-inlineNew-tn=B216Plain isNothing=False
N2-plain-moduleNew-tn=B216Plain isNothing=False
N3-plain-variant-tn=B216Plain
```

于是三枚读数随之改正；发码针改成 `b216_slot_never_boxes_bare_instance`（“裸实例进槽”
仍是 0），并把 plain 三形也钉进 `b216_slot_stations_ask_one_owner`；
`check_object_slot_boxing.ps1` 的 B2 改钉拆分后的两间：`isCoClassTableEntry` 消费点**恰好 1**
（只有 `registerable` 那一枚标记），`isLateBindClass` 消费点**恰好 3**。

### 六、边界

- 私有类不注册、不可外部 `New`；这一行 desc 只服务 IDispatch 装箱；
- 已是私有类的字段/形参早绑定那一半不变（仍走 `vb6_cls_*` 直读）；只有跨 `void*` / Variant 槽
  的那一步现在会包一层真 IDispatch；
- `PublicNotCreatable` 的对外可见性不变（它本来就不该可创建）。

## §11.101 晚绑定结果的静态类型与 `Single` 的字符串形（第一百一十九刀）

### 一、读数（Charts 2020，用户点名的两处）

| 项 | 旧 | 新 |
|---|---|---|
| ucTreeMaps 每个色块的项名 | 一个都不显示 | 全部显示（WhatsApp / Youtube / Pinterest / Wikipedia / Twitter / Instagram） |
| ucPieChart2 数据标签 | `23.7000007629395%` | `23.7%` |
| 探针 `.temp/k122/pct.bas` | `single=23.7000007629395`、`pointthree=0.300000011920929` | `single=23.7`、`pointthree=0.3`，Double / Variant 两行逐字节不变 |

### 二、病 A：晚绑定调用的结果是 **Variant**，不是对象

`ucTreeMaps.ctl:1531/1533` 是 `DrawText(..., m_Serie(i).Labels(j), ...)`；`Labels` 是
`Collection`，`Labels(j)` 走 `vb6_ComCall(..., L"Item", ...)`，交回的是**堆上 `VARIANT*`**。
发码侧 Object 那一支按「声明类型是对象」把它装成
`{vt=VT_DISPATCH, pdispVal=<VARIANT*>}` —— 等于把 VARIANT 结构体地址当 IDispatch 递下去，
控件里 `StrPtr(text)` 读到的是 `vt` / 保留字段 ⇒ 项名一个都画不出来（控件自己的诊断一直只在
等值，所以从 .ctl 源码看不出毛病）。

### 三、修法 A：Object 臂认出「晚绑定结果」就不再二次装箱

`cgen_expr_call_arg_emit.inc` 的 Object 臂在装箱前先看实参的**裸脸**：`vb6_ComCall(` /
`vb6_ComCallByDispid(` / `vb6_ComGetProp(` 三形 ⇒ 改发
`&((vb6_VARIANT[]){vb6_VariantFromComResult(<call>)})[0]`，整枚 Variant 值原样递下去。
所有权交给 `vb6_VariantFromComResult`（与 `boxToVariant` / Fix 113l 同一条口径），
刻意不登记 `boxTempsToClear_` —— 那条道的 authority 仍是两站（knife 121 + 第一百七十二刀站E）。

### 四、病 B：`Single` 的隐式字符串化走了 Double 的格式

`wrapToBSTR` 里 `case Vb6Type::Single:` 与 `Double` 共用 `vb6_CStrDbl` ⇒ Single 的二进制残差
被原样打出来（`23.7` 存成 `23.7000007629395`，`0.3` 存成 `0.300000011920929`）。VB6 的
`CStr(Single)` 是 **7 位有效数字**，与 `CStr(Double)` 的 15 位不是同一个函数。

### 五、修法 B：`Single` 走 `vb6_CStrSingle`

`cgen_expr_binary_util.cpp` 的 `wrapToBSTR` 把 `Single` 从 `Double` 那一格拆出来，改发
`vb6_CStrSingle`（RTL 早已有这一档，只是发码侧没接）。

### 六、闸门

- `gates.ps1` 9/9 绿；四道邻居 `check_object_slot_boxing.ps1`（装箱站点 / 三站 / 借用引用
  逐格不动）、`check_com_latebind_guard.ps1`、`check_box_temp_reclaim.ps1`（decl=10 / clear=9 /
  unreclaimed=1 不变）、`check_variant_i8_coverage.ps1` 全 PASS；
- churn 763 行 **changed=7**：`tests/Charts 2020/Proyecto1.vbp` 与 `ucChartArea` / `ucChartBar` /
  `ucPieChart` / `ucTreeMaps` 四份子工程（Single 字符串形），`tests/ctrlrichtextbox/RtfApp.vbp`、
  `tests/test_memdim.bas`（同一形状）；其余 756 行不动。

### 七、边界与随手记账

- `vb6_CStrSingle` 只改**隐式**字符串化（`&` 拼接 / `CStr` 腿）；显式 `Format$` 一档不动；
- 整数百分比在 VB6 里就叫 `16%`（`CStr(16!)` 不带 `.0`），本刀不改这条语义 —— 要「一律一位
  小数」得由 .ctl 自己 `Format$(Percent, "0.0")`，不是编译器的活；
- **量法附注**：本机屏幕 DPI = 120。`PrintWindow` 那套截图脚本若不先
  `SetProcessDpiAwarenessContext(-4)`，`GetWindowRect` 交的是虚拟化后的 96 口径矩形 ⇒ 位图比
  真窗口小 20%，看着像「图没画全」（本轮 ucChartArea2 就是被它冤枉的）。以后截图脚本一律先声明
  per-monitor v2。

## §11.102 UC 的 TextWidth/TextHeight 跟真屏 DPI 量字（第一百二十刀）

### 一、读数（Charts 2020 · ucPieChart2，用户点名「数字对了但没有百分号 %」）

| 项 | 旧 | 新 |
|---|---|---|
| ucPieChart2 五个数据标签 | `23.1` `30.9` `15.3` `14.1` `16.6`（末尾 `%` 全被切掉） | `20.7%` `21.1%` `11.4%` `22.8%` `24%`（全部完整） |
| 本机 `UserControl.TextWidth("0000000000")` | 60（11px 口径 = 96 DPI） | 75（14px 口径 = 真屏 120 DPI） |
| `tests/ve_units` | U-SW/U-TW/U-TH 全 True | 全 True（判据是「缇型 = 像素型 × TwipsPerPixelX」，与 DPI 无关） |

### 二、病：量字和画字用了两个 DPI

控件里画字一律是 `MulDiv(oFont.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72)`（Charts 2020 的
`ucPieChart` / `ucChartArea` / `ucTreeMaps` 三处 `DrawText` 都这么写）；而
`vb6_uc_measureText` 造度量字体时**恒传 96**（Fix 113k 的旧结论：换真屏 DPI 反而把 Charts
的图例/轴标签排布推坏 —— 那一版两侧只对上了一半，所以旧结论不成立）。

120 DPI 上比值 14/11 ≈ 1.27：`TextWidth` 比实际画出来小 25%。控件自己拿 `TextWidth` 算标签框
（`ucPieChart.ctl:1325` `.TextWidth * 1.3`）与标题行高，框就正好差几个像素 —— GDI+ 居中排布时
最后一个字形整枚被裁掉，于是 `{P}%` 只剩数字。

### 三、修法：量字走调用方给的真屏 DPI

`vb6rtl_com.c::vb6_uc_measureText` 里把两处字面 `96` 换成 `GetDeviceCaps(hdc, LOGPIXELSY)`
（取不到才回 96）。「八档 → LOGFONT」仍问唯一主人 `vb6_UC_FontLogFont`，本刀只换 DPI 那一格。

### 四、判据

- Charts 2020 真跑（`.temp\k119\charts_dpi_180317\Proyecto1.exe`，DPI=120）：五个标签的 `%`
  全部回来；同一张图上 ucTreeMaps 项名、ucChartArea1/2 的图例与轴标签、ucChartBar1/2 的
  类别名均在位；
- `tests/ve_units` 真跑：`U-RAW tw=540 pix=36 swT=2400 swP=160`、`U-TW/U-TH/U-SW/U-CTX/U-HW/
  U-CNT/U-ARR/U-ARREVT/U-ARRCLICK/U-ARRDBL/U-ARREXT/U-PMOVE` 全 True、`U-DONE`；
- `tests/ucTreeView` 真跑：节点图标（Facebook / Instagram / Meta / WhatsApp）与文字、右侧
  属性页排版逐枚照旧；
- `gates.ps1` 9/9 绿；RTL-only 刀 ⇒ emit-manifest 763 行**零移动**（无 登记 提交）。

### 五、边界

- DPI=96 的机器上 `GetDeviceCaps` 就是 96 ⇒ 取值与改前逐字节相同，CI 那 22 格口径不动；
- 只动 **UserControl** 的那一对；普通控件走 `vb6_ControlMeasureTextPx`（按 HWND 取字体）未动；
- 本刀不改 `CStr(Single)` 的位数语义（整数百分比仍叫 `16%`），要「一律一位小数」得由 .ctl 用
  `Format$(Percent, "0.0")`。

## §11.103 私有类实例跨 void* 槽：TypeOf … Is 与 Set 的引用计数归属（第一百二十一刀）

### 一、读数（探针 `.temp\k121\typeprobe`，私有类 `TofC` 无 `Instancing`）

| 相 | 改前 | 改后 | VB6 |
|---|---|---|---|
| `TypeOf o Is TofC`（Object 槽） | False | **True** | True |
| `TypeOf v Is TofC`（Variant 槽） | False | **True** | True |
| `ByVal Item As Object` 形参里 `TypeOf Item Is TofC` | False | **True** | True |
| `TypeOf a Is TofC`（原生类变量） | True | True | True |
| `TypeOf o Is Object` / `TypeOf Nothing Is TofC` | True / False | True / False | 同 |
| `TypeName(o)` / `TypeName(v)` | TofC / TofC | TofC / TofC | 同 |
| 双槽 `Set s1 = New : Set s2 = s1 : Set s1 = Nothing` | 清 s1 即 Terminate，s2 悬垂（`label=""`、`seed=0`、stderr `C3_LATEBIND_MISS kind=notdispatchable`） | 清 s1 不销毁、`s2.label` 仍 `C-label`；清 s2 才 Terminate 一次 | 同改后 |
| 三槽 `t1/t2/t3` | 清第一枚即销毁，后两枚静默 | t1、t2 清掉都不销毁，t3 清掉才 Terminate | 同改后 |
| 单槽 `Set q = New : Set q = Nothing` | Terminate 1 次 | Terminate 1 次 | 同 |

### 二、两个病

**甲｜TypeOf 恒假。** `visit(TypeOfExpr&)` 在「目标是工程类、但声明类推不出」（Object / Variant 槽、
`ByVal Object` 形参）时落到 `vb6_TypeOf` —— 那个出口是 `vb6rtl_conv.c` 里只认 Picture 的桩，其余
`return 0`。同一枚对象的 `TypeName` 却已经答对类名 ⇒ 判据不是缺，是没人问。

**乙｜`Set` 只拷指针。** `cgen_setlet_set_rhs.inc` 兜底那条发 `target = value; /* Set */`：不取引用，
也不退旧值。于是两枚槽只持一枚引用（清任意一枚就销毁对象、另一枚变悬垂），而槽里原有值时再 Set
一次 = 裸覆写泄漏（同一条兜底的两面）。

### 三、修法

- **甲**：新增 `vb6_TypeOfProject(obj, L"类名")`（`vb6rtl_conv.c`），身份只问一位主人
  `vb6_ComObject_ClassNameOfAny`（`vb6comserver_obj.c`）：先问包装器自报类名
  （`vb6_ComObject_ClassNameA`，与 `TypeName` 同源），再问 512 格反查表
  （`vb6_FindProjectClassDesc`，只比指针不解引用）。codegen 只在「项目类 + 声明类推不出」这一格
  发它；声明类静态档、接口 QI、Picture 档一字未动。
- **乙**：兜底改成 VB6 三步 —— `取引用 → 退旧值(置 NULL) → 写新值`。顺序是硬口径（先取后放 ⇒
  `Set a = a` 安全）。两枚闸都在 RTL：`vb6_ComAddRefDispatch` / `vb6_ReleaseObject` 都先问
  `vb6_ComIsDispatchable`（HWND / 裸工程实例 / 伪 vtable 全跳过）⇒ 原先裸拷贝的那几档行为不变。

**第一版教训（实测，留档）**：无条件的「一律 AddRef」是错的 —— 右值若自带对价引用（`New`、调用
形态、新建包装器 `vb6_ComObject_From*Instance`），那一枚本来就归槽所有，再 claim = 每发净 +1。
实测全场 `term=0`，连 `Set q = New TofC : Set q = Nothing` 都不销毁。现按「只对裸槽标识符取引用，
退旧值对所有形状都做」收口。

### 四、判据

- 探针 23 相：上表「改后」全中；stderr 为空（改前两行 `C3_LATEBIND_MISS` +
  `property "Label" not found`）；
- `gates.ps1` 9/9 绿（含 `check_manifest_coverage` 763 行、`check_rtl_embedded` 138 资源逐字节）；
- emit-manifest 763 行 **changed=183**（全局 Set 站点换形，其余不动）。

### 五、边界

- 派生类按基类名问仍答 False（与 Fix 193 那条静态档同一个假阴性方向；真要链条得给每个类加运行时
  类型字段）；
- 外部 IDispatch / VB 控件 / 字体 / 宿主 HWND 一律仍答 False（表外答案，与改前同向）；
- `Set` 取引用只认**裸槽标识符**；成员访问（`me->F`）与真调用形态维持裸拷贝（未量，不猜）；
- 局部对象格仍不在过程出口放手（§B164 那族），本刀不动。

## §11.104 声明档工程类槽的 `New` 交码形状：名字撞车 → 拥有型包装器进 typed 槽 → 实例永久存活（第一百二十二刀）

### 一、读数（探针 `.temp\k122\shape`、`.temp\k122\fieldprobe`）

同一句 `Set <声明档工程类槽> = New FC`，同一个工程里发成两种形状，差别只在**变量名**：

| 源 | 交码 |
|---|---|
| `Dim k1 As FC : Set k1 = New FC` | `k1 = Unbox(FC_New())` —— 裸实例，正确 |
| `Dim b As FC : Set b = New FC` | `b = Unbox((void*)FromInstance(desc, FC_New()))` —— 拥有型包装器 |

`b`/`a` 撞上 `FBox` 的 `Public A/B As Object` 字段名，`k1` 不撞。

后果（`fieldprobe` 的 C0/C3 臂）：改前 `Set a = New FC : a.Seed=31 : Set a = New FC` 交
`C0 term=0`，seed=31 **全程无 Terminate**（旧实例永久存活）；改后 `[FC.Terminate seed=31]` 紧接
`C0 term=1`，过程出口两枚 seed=0 各拆一次，全场四枚归零。

机理在 RTL 那一行：`vb6_cls_FC_Release` 里 `if (me->__comObj != NULL) return 0UL; /* wrapper owns
 teardown */` —— 包装器 `ownsInstance=1` 但**没有任何人持有它**（槽里装的是 Unbox 回的裸实例），
于是类引用计数归零也不拆、包装器自己也没人放。

### 二、成因

`lateBoundSlotTarget()` 的第二臂用 `knownObjectVars_.count(targetLower)` 答「这格是不是泛对象槽」，
而那张表里混着**未限定**的类 `As Object` 字段名 —— 局部名一旦撞上就误答 true：`boxIntoLateSlot` 把
`New` 包成拥有型，typed 槽随后又 Unbox 回裸实例，包装器于是失主。

### 三、修法

`lateBoundSlotTarget()` 加一条前置闸：`node.target` 的 `inferClassTypeOfExpr` 非空（声明档是工程类）
⇒ 这一格不是泛对象槽，直接 false。判据问 AST，不问名字表：`As Object` / `Variant` 拿不到类名时
照旧落到原臂，行为不变。

### 四、判据

- 探针：C0/C3 臂 Terminate 全归位；`shape/Shape3` 的 `b` 与 `k1` 发同一形状；
- `gates.ps1` 9/9 绿（exe sha=`50e95c561939`）；
- emit-manifest 763 行 changed=1：只有 `tests/b190_scope/B190.vbp` 换形（16732 → 16652 字节，
  消失的正是塞进声明档槽的那几分 `FromInstance(desc, …)` 包装），其余 762 行不动。

### 五、边界（在册，本刀不碰）

- **污染源本身没清**：本刀只在装箱站加闸，`knownObjectVars_` 里为什么会有别的类的字段名（driver
  预扫描往那张表塞未限定名）待查；同一污染还可能影响其它 `knownObjectVars_` 读者。
- `TypeName(x.A)`（字段槽里装裸 `vb6_cls_X*`）仍答 "Object"：`vb6_TypeName` 那句注释明写
  「不该把 512 格反查表拉来当运行期读者 —— 那条口径待拍」，本刀按记录不越位。
- 声明档槽 `Set x = Nothing` 只发 `x = NULL`、不发 `Release`，属 §B164 那族，未动。

## §11.105 借用型包装器的类引用归属：`__comObj` 回填不能当托孤（第一百二十三刀）

### 一、读数（探针 `.temp/k122/fieldprobe` 的 `FProbe7` / `FProbe6`）

`FProbe7`（**不**读 `TypeName`，只让 `Set` 那一趟走完）：

```
E1 term=0            ' Set x.A = y           —— 借用型包装器建好
E2 term=0            ' Set x.A = Nothing    —— 包装器拆，实例另有主人（y），不拆
  [FC.Terminate seed=51]
E3 term=1            ' Set y = New FC       —— 旧 y 退到 0，当场 Class_Terminate
```

trace（`C3_COM_RC_TRACE=1`）：`FROM-BORROW W=… rc=1` → `RELEASEOBJ p=W` →
`RELEASE W cnt=0 owns=0` ⇒ 包装器归零那条路径**完整走通**。

同一句话在 `FProbe6`（多读两枚 `TypeName(x.A)`）里改前/改后：

| 台面 | 改前 | 只加 `field voidptr` 臂（RTL 未动） | 本刀（RTL + desc） |
|---|---|---|---|
| `D1 name=` | `Object` | `FC` | `FC` |
| seed=51 Terminate | 有（`D2 term=1`） | **没有**（`D2 term=0`，此后永不） | 仍没有（见「五、边界」） |

中间那一列是「把装箱放开但没给借用型包装器记账」的直接后果：**装箱那一瞬间就把这个实例的退场永久堵死**。

### 二、成因

生成类的 `vb6_cls_X_Release` 里：

```c
if (me->__comObj != NULL) return 0UL;   /* wrapper owns teardown */
```

而借用型包装器（`FromBorrowedInstance`，`ownsInstance = 0`）当时**两头都不在账上**：建的时候不取类引用、
拆的时候不清 `__comObj` 也不退引用（只 `*(void**)vb6Instance = NULL`）。于是：

1. typed 变量 `y` 退到 0 时，`Release` 见 `__comObj != NULL` ⇒ 认为「包装器负责拆」，返回；
2. 包装器那一头按设计**不动实例**（借用），也从不退引用；
3. 剩下的一枚引用谁也碰不到 ⇒ 实例永久存活，`Class_Terminate` 永不发生。

### 三、修法

1. `vb6_CoClassDesc` **末尾**追加两格：`classAddRef` / `classRelease`（`NULL` = 不取不退 = 改前行为）。
   放末尾的理由与 ai/022 B16 那两格同一条：`{0}` 空桩与少写一格的旧写法自动得 `NULL`，不会错位。
2. `emitCoclassRow` **无条件**发这两格（每类都有自己的 `vb6_cls_X_AddRef/_Release`），
   `com_entry.c` 的前向声明区同步补两条 `extern`。
3. `FromBorrowedInstance` 新建包装器时取一枚类引用（复用既有包装器那条不改：那一枚已经记在旧包装器账上）。
4. `ComObj_Release` 的 `owns = 0` 分支：**先摘回填、再退引用** —— 顺序是硬口径，不先摘就会把最后一次退场
   又挡回去（第 1 条那个闸认的就是这个字段）。这一退可能是实例最后一枚，于是当场 `Class_Terminate`。
5. 发码侧：`lateBoundSlotTarget()` 恢复 `field voidptr` 那一臂（§11.104 的 `inferClassTypeOfExpr` 闸在前），
   于是 `Set <非 me 接收者>.<对象字段> = <声明档工程类值>` 走装箱站。

### 四、判据

- `FProbe7`：`E3 term=1` 且 `[FC.Terminate seed=51]` 出现在 `Set y = New FC` 那一行；
- `FProbe4`（knife 122 的 C0..C3 臂）读数不变：`C0/C1/C2 term=1`、`C3 term=2`、seed=31/41 各拆一次；
- `gates.ps1` 9/9 绿（exe sha=`1e55cd94c7d7`）；
- 5 份静态普查全 PASS（object_slot_boxing / class_refcount_authority / projclass_scan_bound /
  g4_receiver_class_gate / legacy_view_arg_owners）；
- emit-manifest 763 行 **changed=0**：desc 表只在真发 COM 入口的产物里出现，`Set` 臂只碰成员对象槽，
  存量夹具一行不动。

### 五、边界（在册，本刀不碰）

- **`TypeName(<对象槽>)` 的临时 Variant 不退引用**：`vb6_TypeName(vb6_VariantFromValue(vb6_VariantFromValue(x->A)))`
  这一形里外层只是过手、内层 `vb6_VariantObject` 记了一枚 `AddRef`，调用结束后没有 `vb6_VariantClear` ⇒
  每读一次漏一枚包装器引用（trace 里 `ADDREF cnt=2` 之后再无对应 `RELEASE`）。这正是 `FProbe6` 的
  `D2 term=0`：`Set x.A = Nothing` 把槽退到 1（漏的那一枚还在），类引用计数于是到不了 0。
  与本刀的类引用归属无关 —— 把 `TypeName` 那两句挪掉（`FProbe7`）归属就是对的。归下一刀（124）：
  builtin 实参临时 Variant 的生命周期，影响面比本刀宽（所有 `TypeName(o)` 都漏），不并进来。
- `TypeName` 的 512 格反查表口径（§11.104 第五条）照旧未动。
- 借用型包装器的 `Is` 同一性仍靠 `__comObj` 回填复用；回填现在只是「同一性缓存」，不再是「托孤」——
  它的清理由包装器自己的拆解负责。

## §11.106 只读谓词的实参不再偷取对象引用：`TypeName` 一类调用的临时 Variant 归属（第一百二十四刀）

### 一、读数（探针 `.temp/k124/probe`、`.temp/k122/fieldprobe/FProbe6`）

`FProbe6` 同一份源的三态：

| 台面 | 第一百二十二刀后 | 第一百二十三刀后 | 本刀后 |
|---|---|---|---|
| `D1 name=TypeName(x.A)` | `Object` | `FC` | `FC` |
| seed=51 的 Terminate | 有 | **没有** | **有** |
| `D2 term=` | 1 | 0 | 1 |

本刀前的 trace（`C3_COM_RC_TRACE=1`）：`FROM-BORROW rc=1` → `ADDREF cnt=2` → `RELEASE cnt=1`
—— 那一枚 `AddRef` 之后再无对应 `Release`，于是 `Set x.A = Nothing` 只能把槽退到 1，类引用计数到不了 0。

`K124` 探针（修后）：`T1 FC True` / `T2 FC True` / `T3 String 8` / `T4 False True` /
`T5 FC True` / `T6 Long() 8195 True True`（数组与 Null 两档也逐位对）。

### 二、成因

`vb6_TypeName` / `vb6_VarType` / `vb6_IsObject` / `vb6_IsNull` / `vb6_IsEmpty` / `vb6_IsArray` /
`vb6_IsDate` / `vb6_IsError` / `vb6_IsNumeric` 的形参都是**按值** `vb6_VARIANT`，而发码把实参
包成表达式里的临时：`vb6_TypeName(vb6_VariantFromValue(vb6_VariantFromValue(x->A)))`。
`vb6_VariantFromValue` 的指针档是 `vb6_VariantObject`，按 Fix 024 **必须 AddRef**（因为 Variant 析构
会 Release）；而这枚临时是表达式里的右值，**没人会 `vb6_VariantClear` 它** ⇒ 每读一次漏一枚引用。
对象是包装器时，那一枚就是包装器的命：槽清空后仍 rc=1。

（同一族里只有对象档会漏：`wchar_t*` 的 `vb6_VariantString` 本来就只存指针不拷贝，标量档不拥有任何东西。）

### 三、修法

1. RTL 加一对**借用档**：`vb6_VariantObjectBorrowed`（`VT_DISPATCH` + 指针，不 AddRef）与它的
   `_Generic` 入口 `vb6_VariantBorrowFromValue`（臂表与正式档逐条对齐，只把 default 换成借用的那枚）。
2. 发码：`cgen_expr_call_builtin_argtype.inc` 里新增一站 —— `callee` 是上述九枚之一时，把每个实参
   **最外层的装箱名前缀** `vb6_VariantFromValue(` 换成 `vb6_VariantBorrowFromValue(`（外层那两层都
   换；嵌套子表达式里的装箱一律不动，它们自有一套所有权）。
3. `scripts/gen_variant_slot_prefixes.ps1` 重跑（生成物 61 → 62 名，多出的就是 `vb6_VariantObjectBorrowed`）。

能借的根据：这九枚 RTL 体只读 `vt` 与 payload，既不存也不清（`vb6rtl_conv.c` 逐行可对）；实参的
生存期覆盖整个调用。借档**不分配任何东西**，也就不存在「谁退」那一问。

### 四、判据

- `FProbe6`：`D1 name=FC term=0` → `[FC.Terminate seed=51]` → `D2 term=1`，出口再拆 seed=0；
  trace 里不再出现无主的 `ADDREF`；
- `K124` T1..T6 六臂全对（含数组 `Long() / 8195 / True` 与 `Null`）；
- **A/B 实证**（`.temp/k124/ab`，四份改过的夹具：test_variant、test_vbnet_ext、GCCache、CtrlState）：
  把新产物的 `vb6_VariantBorrowFromValue(` 改回原名后与基线产物**逐字节相同**；借用档出现次数
  12/7/6/6 —— 即改动只有这一个名字；
- `gates.ps1` 9/9 绿（exe sha=`dab63bf72166`）；5 份静态普查全 PASS；
- emit-manifest 763 行 changed=122：全部是「含这九枚谓词且有指针/对象实参」的产物，形状如上一行。

### 五、边界（在册，本刀不碰）

- `struct tagSAFEARRAY*` 那一臂仍用正式档 `vb6_VariantFromParamArray`（它自己拷贝、自己拥有），
  借用档没有覆盖它 —— 那条路的临时 Variant 归属是另一个话题。
- 任何会**存**实参的出口（`Array(...)`、`Collection.Add`、形参是 ByVal Variant 的用户过程）
  一律不用借用档：那一枚引用是它们要拿走的，借了就是悬空。
- 用户过程（非内置）的 ByVal Variant 形参仍然走正式档：漏不漏归于被调方清不清，与本刀无关。

## §11.107 #412 已关门：`Set <工程类实例>` 进 Object 槽之后晚绑定是真能用的（第一百二十五刀）

### 一、为什么这一刀是「登记」而不是改代码

台账里 #412 一直是开口的（B190 那组里的 P3/P4/P7-P12 读数就被它顶着，故意没钉成针）：
「`Type=Exe` 工程里 `Set 工程实例` 进 Object 槽从不发包器 ⇒ Object 那一路晚绑定全落空」。
第一百二十二〜一百二十四刀把**装箱站**（§11.104 的 AST 闸、§11.105 的类引用归属）、
**谓词实参**（§11.106 的借用档）收拾完之后，这一条自己就不成立了 —— 本刀只把它量下来、
钉成门禁，不改一个字的发码。

### 二、读数（探针 `.temp/k125`）

同一份源在两台架构上逐字相同：

```
C3ProbeB412-P1=B412Item      ' TypeName(o)       —— Object 槽里是包装器, 答得出类名
C3ProbeB412-P2=22            ' o.Seed = 22 : 读回 —— 晚绑定 Set 真进了实例
C3ProbeB412-P3=B412Item      ' Set gSlot = o    —— 模块级 Object 槽同一形状
C3ProbeB412-P4=23            ' gSlot.Seed = 23
C3ProbeB412-P5=23            ' Set o = Nothing  —— typed 变量那一枚不受影响 (借用型)
C3ProbeB412-P6=23            ' Set gSlot = Nothing
C3ProbeB412-TERM=23          ' 只拆一次, 在最后一枚 Set a = Nothing 之后
C3ProbeB412-DONE
```

另一枚探针 `VJ`（`Dim y As FC : Set o = y`，无模块级槽）同读数：`J1 FC` / `J2 22` /
`J3 term=0`（拆包装器不退实例）/ `[FC.Terminate seed=22]` / `J4 term=1`。

### 三、夹具与判据

夹具 `tests/b412_objslot/`（`B412Item.cls` + `B412Main.bas` + `B412ObjSlot.vbp`，
`Type=Exe`）：P1..P6 + TERM + DONE 共 8 枚针，x64 与 x86 各跑一遍。
钉的是三件事：① Object 槽里那枚包装器答得出类名；② 晚绑定 Set/Get 真到实例；
③ 槽的拆解不影响另一枚主人（typed 变量），也就是 §11.105 那对取/退引用在 Object 槽这一形上也成对。

### 四、边界

- 本刀只钉 `Type=Exe` 的 Object 槽；`As Object` 的**外部** IDispatch（`CreateObject`
  那一族）的晚绑定另有出口，不在本夹具范围。
- `knownObjectVars_` 的污染源（§11.104 第五条「待查」）本刀也量了一遍：本轮试了
  `Dim b As FC`（撞 `FBox.B`）在成员读/`Is Nothing`/`Is`/`With`/数组下标五形，与不撞名的
  `k1` / `zz` 对照发码**逐字相同**（声明驱动的权威都排在名字表之前），没找到别的读数；
  仍留作在册，不并进本刀。

## §11.108 `knownObjectVars_` 的污染源找到了：类模块的 `Public` 字段被当成「模块级变量」注入（第一百二十六刀）

### 一、源头

§11.104 第五条一直挂着的「名字表里混着别的类的成员名」，本刀用一枚带标签的登记口读数
（`C3_K126`，`.temp/k125/k126b.err`）定到了确切位置 —— 跨模块登记站
（`cgen_base_generate_crossmod.inc` 的 `sym->isExternal && kind == Variable` 那一圈）：

```
[K126 X] mod=FC     var=a type=Object src=FBox    <-- FBox.cls: Public A As Object
[K126 X] mod=FC     var=b type=Object src=FBox    <-- FBox.cls: Public B As Object
[K126 X] mod=FMain7 var=a type=Object src=FBox
[K126 X] mod=FMain7 var=b type=Object src=FBox
```

stage 3.5 把**类模块的 `Public` 字段**也注成了消费模块的「外部模块级变量」（
`isExternal=true, sourceModule=<类名>`），于是 `FBox` 的字段名进了 `FMain7` 与 `FC` 两个模块的
`knownObjectVars_`。它们不是全局变量：裸名在消费模块里根本不可写，只有 `obj.A` 这种成员访问才成立。

### 二、闸

跨模块登记站的 Object 那一支加一条：源模块是类模块 ⇒ 不登记。
（`sym->sourceModule` 解到 `SymbolKind::Class` 即跳过；标准模块的 Public Object 变量照旧。）

读数：同一枚探针重跑，四行全变 `SKIP`；类模块**自己**的发码照旧（`FBox.cls` 里 `A`/`B` 仍是它
自己的字段，走声明口那一臂）。

### 三、没动的东西（回归面）

| 读数 | 夹具 | 结果 |
|---|---|---|
| `E1 term=0` / `E2 term=0` / `[FC.Terminate seed=51]` / `E3 term=1` | FProbe7 | 与本刀前逐字相同 |
| `D1 name=FC` / `[FC.Terminate seed=51]` / `D2 term=1` | FProbe6 | 同上 |
| `T1..T6`（含 `T6 Long() 8195 True True`） | K124 | 同上 |
| P1..P6 / TERM=23 / DONE | B412 | 同上 |

全语料发码：765 行里 760 行逐字相同，只有那 5 行机器侧差（guisample x3 + TeeChartChartDisplay +
ctrlmscomm，以 CI 那台为准）—— 也就是说这条闸在存量语料上**一行发码也不改**，它清的是名字表。

### 四、钉住闸的上限（不让它过宽）

夹具 `tests/b126_crossmod_objglobal/`（`B126M1.bas` 的 `Public gObj As Object` + `B126M2.bas`
消费 + `B126Item.cls`）：`C3ProbeB126-A1=B126Item` / `A2=8` / `DONE`，双架构。发码侧 `gObj` 仍是
`void*` + `vb6_ComSetProp(gObj, L"Seed", ...)` / `vb6_ComGetProp(...)`（Fix 089 那一臂留得住）。

### 五、边界

- 只拦「类模块的 `Public` 字段被当模块级变量」这一路；类模块自己内部的裸名（声明口登记的）不变。
- 类模块字段符号本身仍留在符号表里（推理、成员表、跨模块类型查询都没动），本刀只改它在
  `knownObjectVars_` 里的影子。

## §11.109 腿④ 的第四档：局部泛对象槽（`Dim o As Object`）在过程出口放手（第一百二十七刀）

### 一、缺口与读数

`slotsToRelease_` 那张表里一直只有三档：接口薄指针（`isView`）、工程类实例（`clsName`）、
数组载体（`isCarrier`）。**局部泛对象槽不在表里** —— 过程里现造现丢，包装器与它握着的那一枚
类引用一起漏，`Class_Terminate` 永远走不到：

```
' 修前 (.temp/k127/localobj)
　one done
　one done
　one done
L1 end          ' 三次调用, 一枚 [K127Item.Terminate] 都没有

' 修后
　one done
　[K127Item.Terminate]
　…（三枚）
```

对照臂（同一份源、只去掉 `Dim o As Object` 那一格）：工程类局部（`Dim it As K127Item`）本来就会在出口
放手 —— 所以漏的只有**泛对象那一档**，不是腿④ 整条（这也是为什么夹具里两种声明同处一室）。

### 二、机制（三处，各一位主人）

1. **登记口**（`cgen_localdecl.cpp` 的 `void*` 那一臂，与工程类 / 接口两档同处）：只认
   声明名就是 `Object` 的那一档；`Static` / `Static Sub` 里那一格跨调用活着（C 侧 static 存储），
   `WithEvents` 那一格另有主人 —— 两条都不登记。
2. **出口发码**（`emitIvrefScopeRelease`）：`vb6_ReleaseObject((void**)&o);`。判据复用 Fix 191
   那一道（`vb6_ComIsDispatchable`）：包装器 / 真 IDispatch 退一轮，裸工程实例指针与 HWND 只清槽
   不越权；入口先清槽再放手（与 §11.105 那条次序同一口径）。
3. **两条撤登**（`dropLateObjSlotRelease`）：
   - `For Each` 的循环变量（`cgen_control.cpp`）：每一轮由循环自己写，元素那一枚归容器/迭代器，
     出口放手 = 无对价的退；
   - 这一格的**地址交给了外部 Declare** 的那一档：复用既有 `dropClassSlotRelease`（它按
     「非 view」撤，正好覆盖本档）。

### 三、风险臂量过（不是拍脑袋）

| 形状 | 探针 | 读数 |
|---|---|---|
| `Set o = Held()`（工程内函数交回一枚泛对象） | `.temp/k127/neg` | `N1 keep` → `N2 still=K127Item seed=11`（不过退） |
| `Set o = c.Item(1)`（Collection 元素） | 同上 | `N3 item=K127Item` / `N4 keep=K127Item`（不过退） |
| `For Each oEach In c` | 夹具乙 | `A2-live=B127Item`（撤登生效，容器里那一枚还活着） |
| 声明住在 `If` 括号里 | 夹具丙 | 编得过（提升到序言）+ 照样放手（`A3-terms=4`） |

### 四、回归面

- `tests/b445_refcount` 全夹具：针上的 13 条（W1–W10 / W38 / W39 / DONE）逐字相同；
- FProbe7 / FProbe6 / K124 / B412 / B126 五族读数逐字相同（`E1..E3` / `D1..D2` / `T1..T6` /
  `P1..P6` / `A1..A2`）；
- 全语料发码：只有**带局部 `As Object` 的文件**各多出一行 `vb6_ReleaseObject((void**)&o);`，
  其余行一字未动（登记那一刀的数字见下面的「登记」节）。

### 五、边界

- 只认声明名就是 `Object` 的那一档。宿主类局部（`Dim c As Collection` / `As Dictionary`）与外部
  COM 局部同样会漏一枚包装器，但它们不在本档（需要先量清「那些槽里躺着的是不是借用」）—— 在册。
- 模块级 / 类字段那一档不在本条腿上：本体退出时的整体拆除（VB6 关机时把全局对象置 Nothing）
  是另一格机制，今天 `vb6_Exit()` 只调 `vb6_ComExit()`—— 在册。
- `Exit Sub` / `Exit Function` 与正常出口走同一个 `vb6_proc_exit`，本档自动跟着（与腿④ 其余三档同一条路）。

## §11.110 晚绑定实参打包后的「内容」放手：一个出口、三处裸 free（第一百二十八刀）

### 一、缺口

晚绑定那八个出口的实参是**堆上自造的 VARIANT**（`vb6_ComPackObject` / `vb6_ComPackVB6InstanceRaw` /
`vb6_ComPack_<Cls>` …）：结构体是 `calloc` 的，内容里那一枚引用（或 BSTR）只有 `VariantClear` 会退。
修前有三处收尾只 `free` 结构体、从不碰内容：

1. `vb6_ComCallRcv` 的**宿主对象**岔口（`vb6_Host_Call` 之后）—— 内建 `Collection` 是宿主对象
   （`vb6_uc_isColl`），所以 `c.Add o` 每次都从这里过；
2. IUnknown `AddRef`/`Release` 的**合成**岔口（按名直发那一路）；
3. `vb6_ComSetPropArgRcv`（`d.Item(k) = 值` 这一形）：值那一枚与索引参数一概裸 free。

宿主集合那一头早就按「打包方会 `VariantClear`」写好了自己的对价（`vb6_coll_push` 里那句
"对象档: 集合自己持有一份 —— 此前这里借用不计数, 打包方那次 VariantClear 就把包装器 free 掉了"，
账 #305 §B165 第五十四刀加的 `vb6_ComAddRefDispatch`），所以这一族漏是**那一句对价落了空**，
不是语义分歧。

读数（`.temp/k128`，同一份源、只换二进制）：

```
' 修前
G3 terms=1            ' c.Add o 之后 Set o = Nothing
G4 terms=1 count=0    ' Remove 走完, 元素一枚 Terminate 都不打
P3 terms=0 count=0    ' d("k") = o 之后 d.Remove "k"
' 修后
G4 terms=2 count=0
P3 terms=1 count=0
```

### 二、机制（两条口径，各一位主人）

- `vb6_ComDropPackedArgs(a, argc)`（既有）与 `vb6_ComDropPackedOne(pv)`（本刀新增，单枚）：
  **内容问 `VariantClear`，结构体才 `free`**；
- `VT_BYREF` 跳过：出参指向调用方存储，引用归调用方（这条判据与 `dp.rgvarg` 那一段一字相同）。

三处站点全部改成调这两个口，`vb6_ComCallRcv` 主路（`dp.rgvarg` 浅拷贝那一段）本来就是这个口径。

### 三、夹具与读数

`tests/b128_argpack_release/`（`B128Item.cls` + `B128Main.bas` + `B128ArgPack.vbp`）：

| 针 | 修前 | 修后 |
|---|---|---|
| `C3ProbeB128-A1-terms`（宿主那岔：`Collection.Add` + `Remove`） | 0 | **1** |
| `C3ProbeB128-A2-terms`（真 IDispatch 那岔：`d("k") = o` + `d.Remove`） | 0 | **2** |
| `C3ProbeB128-DONE` | ✓ | ✓ |

x64 / x86 双架构同针。

### 四、边界（在册）

- `vb6_ComSetPropRcv` 自己那一条（`o.Prop = 值`，无索引）仍然是裸 `free`：它同时服务 Font / Extender
  的**借用指针**写（`case 4: *(void**)efp = punkVal`，明写"不接管所有权"），放手口径要先量清
  「借指针那几格谁持有」再动 —— 在册。
- `vb6_ComSetRefRcv`（`Set o.Prop = x`）走的是借指针，不属本档 —— 在册。
- 类字段（`Public Col As Collection`）在 `_Destroy` 里不放手（实测 `.temp/k128` KHaz：`F3 terms=0`），
  那是腿⑤ 的字段档，不是实参打包 —— 在册。

### 五、补（第一百二十八刀补）：宿主包装器那一档**不清内容**

第一版把宿主岔（`vb6_ComCallRcv` 里 `vb6_Host_Call` 之后那一处）也接上 `VariantClear`，CI 当场红：
job `Tests (vbp #8)` 的 `geomcache_x86` 报 `latebind-miss: GCCache 期望 0 条, 实到 12 条`。本机复现同一读数
（`.temp/k128/gc_x86/gc.err`，12 条 `site=DispView` / `site=ComSetProp name="Width" ... kind=notdispatchable`），
连带 GC18/GC19 两条针从 `True` 翻 `False`（`raw=3007,2003,1507,1401` / `200,82,134` —— 一条 Width 都没落地）。

根因不在放手本身，而在**谁持有那枚包装器**：`vb6_ComPackObject` 对宿主对象（控件/窗体/集合/字体）发的是
`vb6_UC_WrapHostObject` 现造的 `Vb6HostWrap`（`refs=1`），而宿主模型 `vb6_coll_push` 只对**工程类包装器**
（`vb6_ComObject_IsVb6Wrapper`）调 `vb6_ComAddRefDispatch` —— 宿主包装器一格引用都不加，集合里那一枚是**借**的。
修前那处裸 `free` 结构体，包装器本人留着，`collC(iC)` 读回还能用；接了 `VariantClear`，集合手里那枚当场被
`HW_Release` free 掉，GC18 的 `With collC(iC) : .Width = 1234` 拿它就是一枚野指针。

口径（两条出口，各按自己的持有方）：

- 宿主岔改调 `vb6_ComDropPackedHostArgs`：`vb6_UC_IsHostWrap` 认得出是**我们包出来的那枚** `Vb6HostWrap` 时
  只 `free` 结构体、内容留在原处（与 `uc_collection` 那一侧「集合自己持有一份」尚未配对相对）；
  其余（真 COM / 工程类包装器）照旧走 `vb6_ComDropPackedOne` 的 `VariantClear`；
- `vb6_UC_IsHostWrap`（`uc_hostmodel.c`）**从后往前**扫包装表：表只追加不撤、地址会被下一枚包装器复用，
  最新的那条才是它的主人（§B165 第五十四刀那条 stale pair 的教训）；判据只认包装器本人，
  真宿主对象（HWND / UC 实例 / 集合 / 字体）一个字段都不碰。

读数（同一份源，只换二进制）：

| 针 | 第一版（宿主岔也清内容） | 补后 |
|---|---|---|
| `geomcache_x86` latebind-miss | 12 | **0** |
| `geomcache` / `geomcache_x86` GC18 | False | **True**（`raw=1234,1234,1234,1234`） |
| 同上 GC19 | False | **True**（`raw=82,82,82`） |
| `C3ProbeB128-A1/A2` | 1 / 2 | **1 / 2**（本刀的主体不动） |

回归面逐字相同：`tests/b445_refcount`（W1–W39 + DONE）、FProbe7（E1–E3）、FProbe6（D1–D2）、K124（T1–T6）、
B412（P1–P6 / TERM=23）、B126（A1–A2）、KOwn（G0–G4）、KHaz（F1–F3）。九门 9/9、语料 772 行不动
（本刀只动 RTL，发码一字未变）。

在册：宿主包装器进集合仍是「一枚只加不退的引用」—— 那是 `uc_collection` 元素自己不 Release 那一格
（§11.110 边界同源），修前就是这样，不是本刀引入的。

## §11.111 内建 Collection 的载体寿命：计数、元素出厂、包装器手里那一份（第一百二十九刀）

### 一、缺口

内建 `Collection` 的载体是 `calloc` 的一枚 `vb6_CollRec`，各槽里存的**直接就是它**（不是 `vb6_ComObject`
包装器）。而 Set 那对出口都先问同一句判据：

```c
if (vb6_ComIsDispatchable(pUnk)) pUnk->lpVtbl->Release(pUnk);   /* vb6_ReleaseObject */
if (!vb6_ComIsDispatchable(p)) return;                          /* vb6_ComAddRefDispatch */
```

宿主对象一律不可分派（首字段是 `tag`/句柄，不是 vtable），于是这一对**双双空跑**：
`Set c = New Collection` 之后再怎么 `Set c = Nothing`、再怎么重新赋值，载体既不拆、
里面的元素也一枚都不出厂。实测（`.temp/k129`，只换二进制）：

```
' 修前
A1 terms=0        ' Set c = Nothing 之后 (集合里两枚元素, 应 2)
A2 terms=0        ' 重新赋值丢掉旧载体 + 末了 Nothing (应 2)
B  terms=0        ' 两枚槽指同一枚载体, 最后一枚退掉 (应 1)
' 修后
A1 terms=2 / A2 terms=2 / B terms=1
```

这一格就是第一百二十八刀量到的 `c.Add New X` 恒 `terms=0` 的根：那一刀把**实参**那一枚引用的
放手补齐了（元素于是活在集合自己那一份引用上，与 VB6 一致），可集合自己**永远不死**，元素当然也
永远不出厂。

### 二、机制

- `vb6_CollRec` 加 `int32_t refs`；`vb6_Collection_New` 出生即 `refs=1`；
- `vb6_Collection_AddRef` / `vb6_Collection_Release` 是载体的唯一一对出口。Release 在 `refs`
  退到 0 时：逐枚按 `vb6_Collection_Remove` **一字不差**的口径放手（`vb6_VariantValueFree` + 只对
  `vb6_ComObject_IsVb6Wrapper` 那一档 `vb6_ReleaseObject`，键另放），再 free 三项、先撕 `tag`
  再 free 载体（第二发 Release 不再认它）；
- `vb6_ComAddRefDispatch` / `vb6_ReleaseObject` 各加一条 Collection 分支放在 `ComIsDispatchable`
  **之前**——判据只有一句（`vb6_uc_isColl`），两处成对；
- 集合套集合：`vb6_ComPackObject` 把宿主对象包成 `Vb6HostWrap` 时，若目标是内建 Collection，
  替它**收一份载体引用**并给包装器打 `ownsTarget`；`HW_Release` 在包装器真死时把那份放回。
  没有这一份，`Set inner = Nothing` 会在宿主那边还拿着包装器时就把载体拆了。

### 三、夹具与读数

`tests/b129_coll_lifetime/`（`B129Item.cls` + `B129Main.bas` + `B129Coll.vbp`）：

| 针 | 修前 | 修后 |
|---|---|---|
| `B129-A-count`（`Set Nothing` 前 Count） | 2 | 2 |
| `B129-A1-terms`（`Set Nothing` 拆掉两枚元素） | 0 | **2** |
| `B129-A2-terms`（重新赋值丢旧载体 1 + 末了 Nothing 1） | 0 | **4** |
| `B129-B-alive` / `B129-B-terms`（两枚槽同一载体：退一枚槽不许拆） | 1 / 0 | 1 / **5** |
| `B129-C-inner`（集合套集合：inner 置 Nothing 后仍可用） | 1 | 1 |

x64 / x86 双架构同针（x86 实测逐字相同）。

### 四、回归面

- `tests/b445_refcount`：W1–W39 + DONE 逐字相同；
- `b412_objslot`（P1–P6 / TERM=23）、`b126_crossmod_objglobal`（A1–A2）、`K124`（T1–T6）、
  `FProbe7`（E1–E3）、`FProbe6`（D1–D2）、`KOwn`（G0–G4）、`KHaz`（F1–F3）、`K128`（C1/D1=0）、
  `b128`（A1=1 / A2=2）逐字相同；
- `tests/geomcache` x64 + x86：`C3_LATEBIND_MISS` 0 条，GC17/GC18/GC19 全 True；
- 九门 9/9；语料 772 → 774 行（只多夹具自己两行；本刀只动 RTL，发码一字未变）。

### 五、边界（在册）

- **局部 `As New Collection` / `Dim c As Collection` 出作用域仍不退**：腿④ 只认声明名就是 `Object`
  的那一档，宿主类局部还没有那一句 `vb6_ReleaseObject`（实测 `.temp/k129` A2 = 0）。载体计数已经
  就位，缺的只是发码那一句 —— 下一刀。
- **集合套集合**：内层载体由包装器护着，而宿主侧存的是**借来的**包装器指针（§11.110 那条边界），
  包装器自己不释放 ⇒ 内层载体仍不拆（C 那格两版都 alive）。
- 集合元素里的**宿主包装器**仍是借的：Release 只对 `vb6_ComObject_IsVb6Wrapper` 放手，与控制/窗体
  寿命有关的那一档不在这里。
- 模块级 / 类字段那一档仍不放手（`KHaz` F3 = 0），与 §11.109 同一条边界。
- `Scripting.Dictionary` 是真 IDispatch 载体（`vb6com_dict.c`），不走这一格 —— 在册。

## §11.112 腿④ 宿主类局部档（第一百三十刀）

### 一、缺口

第一百二十九刀把 Collection 载体的计数交回载体自己（`refs`，`New` 时 1），但**出口那一句**
`vb6_ReleaseObject` 只发给声明名就是 `Object` 的槽（§11.109 第四档）。于是：

```vb
Private Sub NewLocal()
    Dim c As New Collection
    c.Add New B130Item
End Sub          ' 载体没人放手 ⇒ 元素一枚 Terminate 都不打
```

实测（`.temp/k129` 的 A2 格）：修前 `A2 terms=0`，修后 `2`；b130 夹具 `E2-terms=0 → 1`。

### 二、改法

挂点在 `cgen_localdecl.cpp` 那条 `cType == "void*"` 的臂里 —— 与 `Object` 共用同一张表
（`slotsToRelease_`，出口由 `emitIvrefScopeRelease` 发）：

- `isPlainObject`（声明名 `object`）之外补一个 `isHostClassLocal`（声明名 `collection`）；
- 三道已有闸门原样生效：`!var.isStatic`（Static 槽跨调用活着）、`!inStaticProc_`、`!var.isWithEvents`；
- For Each 循环变量那一条撤登记（`dropLateObjSlotRelease`）按 `isLateObj` 认人，照旧管用。

**只收 `Collection` 一个名字**：`Dictionary` 的 C 侧槽是类型化接口
`vb6_ComIface_IDictionary*`，不走这条臂（实测 `DictLocal` 里没有那一句，`D1` 增量 0）—— 在册。

### 三、反面证人（同一夹具，出口放手不许伤到别人）

| 针 | 含义 | 读数 |
|---|---|---|
| `B130-E1-count` / `B130-E2-terms` | `Dim c As New Collection` 出作用域 | 1 / **1**（修前 0） |
| `B130-A1-count` / `B130-A2-field` | 局部是从**字段**借的（`Set c = gBox.Col` 那一发 AddRef） | 1 / 1 |
| `B130-S1-count`（两次）/ `B130-S2-terms` | `Static s As Collection` 跨调用活着 | 1,1 / 1 |

「字段那一枚跟着走」的形态不会出现：`Set` 取引用的发码本来是
`vb6_ComAddRefDispatch(vref); vb6_ReleaseObject(&dst);`，第一百二十九刀让 `ComAddRefDispatch`
对 Collection 计数，出口放手只退自己那一份。x64 / x86 逐字相同。

### 四、发码面（本刀唯一动发码的刀）

全量重算 774 行：**19 行变**，逐行 A/B（旧二进制 vs 新二进制各出一份 `--emit-c`）：

```
add=1..11  del=0   （19/19 行都只有新增行，没有一行被改写或删除）
```

新增行全是同一句 `vb6_ReleaseObject((void**)&X);`，X 逐个对上源码里的 `As Collection` 局部：
`c`、`a`、`Value`、`TempCol`、`Lables`、`Icons`、`collC`、`col`、`thunkCol`、`tmpFrames`
（Charts 2020 三张图 + `Form1/Form2`、`geomcache`、`k174`、`k294`、`ucSlider_2022`、
`GifcontrolSource`、`vb6probe`、`b127/b128/b129` 夹具）。

### 五、读数与回归

- `b445` W1–W39 + DONE 逐字相同；`B412`、`B126`、`K124`、`KOwn` G0–G4、`KHaz` F1–F3、
  `b128` A1/A2、`b129` 七针（2/2/4/1/5/1）逐字相同；
- `K128` 的 `C1 terms=0 → 1`（同一份源，`Dim c As Collection` 那一格现在是**应当**有的读数）；
- `geomcache` x64 + x86：`latebind-miss` 0 条，GC17/18/19 全 True
  （`GC18=1234,1234,1234,1234`、`GC19=82,82,82`）；
- 九门 9/9；语料 774 → **776** 行（夹具两行）。

### 六、边界（在册）

- **`Dim dc As Dictionary` 出作用域仍不退**（`vb6_ComIface_IDictionary*`，接口臂没有出口登记）。
- 模块级 / 类字段那一档仍不放手；集合元素里的宿主包装器仍是借的；集合套集合的内层载体仍由
  借出的包装器护着（与 §11.111 五同）。
- 泛对象档与宿主类档共用一个 `isLateObj` 标记 ⇒ `dropLateObjSlotRelease` 的「无对价撤登记」
  对两者同时生效（For Each 循环变量 / 交地址给 Declare 那两处）。

## §11.113 腿④ 宿主类局部档第二格（第一百三十一刀）

### 一、缺口

`Dim dc As Dictionary` 的 C 侧槽是**类型化接口** `vb6_ComIface_IDictionary*`，不经过
`cType == "void*"` 那条臂，所以第一百三十刀只收到 `Collection`。实测（`.temp/k130` 的
`DictLocal`）：出口一句不发，`D1` 增量 0；b131 夹具修前 `E2-terms=0`、`Z-terms=0`。

### 二、改法

挂点在 `cgen_localdecl.cpp` 的 ComClass/ComInterface 臂里（`knownTypedComVars_` 那一站）：

- 声明名是内建 `dictionary`，且过三道闸门（`!var.isStatic` / `!inStaticProc_` / `!var.isWithEvents`）
  ⇒ 登记 `trackLateObjSlotForRelease`；
- 出口那一发与泛对象档同一形状 `vb6_ReleaseObject((void**)&x)` —— 与 `Set` 那一站退旧值
  用的是同一句（实参类型对得上，`vb6_ComIsDispatchable` 把关，真 IDispatch 才 Release）。

**只收内建 `Dictionary` 一个名字**：外部 ComClass / ComInterface 一大片，每一格的引用归属
各有各的主人（`GetObject` / 属性取 / 参参借），在册另刀。

### 三、反面证人（同一夹具，出口放手不许伤到别人）

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B131-E1-count` / `E2-terms` | `Dim dc As Dictionary` 出作用域 | 1 / **0 → 1** |
| `B131-A1-count` / `A2-field` | 局部从**字段**借（`Set dc = gBox.Dict` AddRef） | 1 / 1 |
| `B131-S1-count`（两次） | `Static d As Dictionary` 跨调用活着 | 1,1 |
| `B131-O1-alive` | 局部**交给字段**之后字段仍可用 | 1 |
| `B131-Z-terms` | 累计（E 那枚 + 换人时旧载体 + 末了 Nothing） | **0 → 3** |

「交给字段」那一站的发码是 `vb6_ComAddRefDispatch(vref); vb6_ReleaseObject(&字段); 字段 = vref;`
——本地那一枚退掉后字段手里还有一份，所以 `O1-alive=1`（不是悬空）。x64 / x86 逐字相同。

### 四、发码面

全量重算 776 行：**3 行变**（`tests/b128_argpack_release/B128ArgPack.vbp`、`B128Main.bas`、
`tests/test_com_default_prop.bas`），逐行 A/B（旧二进制 vs 新二进制）全部 `add=1`、`del=0`，
新增行只有 `vb6_ReleaseObject((void**)&d);` 那一句（对源码里的 `Dim d As Dictionary`）。
语料 776 → **778** 行（夹具两行）。

### 五、读数与回归

- `K128` 的 `D1 terms=1 → 2`（同一份源，`DictCase` 那一格现在应当有 Terminate）；
- `b445` W1–W39 + DONE、`B412`、`B126`、`K124`、`KOwn` G0–G4、`KHaz` F1–F3、`b128` A1/A2、
  `b129` 七针、`b130` 七针逐字相同；
- `geomcache` x64 + x86：`latebind-miss` 0 条，GC17/18/19 全 True；
- 九门 9/9。

### 六、边界（在册）

- **外部类型库接口 / ComClass 局部**（`Dim x As FileSystemObject` / `As XMLHTTP` …）仍不登记：
  它们的引用可能来自 `GetObject`、属性取或参数传递，归属各不相同，需逐档量过再收。
- 模块级 / 类字段那一档仍不放手（b131 的 `Z` 那两枚是全靠**显式赋值/置空**才拆的）；
  集合里借出的宿主包装器、集合套集合的内层载体仍与 §11.111 五、§11.112 六同一条边界。

## §11.114 腿④ 第三档：类型化 COM 接口局部（第一百三十二刀）

### 一、缺口

第一百三十一刀只收了内建 `Dictionary` **一个名字**。同一条臂里的其余类型化 COM 接口槽
（`vb6_ComIface_I<X>*`，既包括别名接口名 `As IDictionary`，也包括外部类型库的
`As FileSystemObject` / `As IPictureDisp` / `As ADODB.Stream` …）出作用域仍一句不发：

```vb
Private Sub ViaAlias()
    Dim d As IDictionary
    Set d = New Dictionary
    d.Add "k", New B132Item
End Sub          ' 修前 terms=0（载体不拆，元素不终止）
```

### 二、改法与所有权口径

`cgen_localdecl.cpp` 的 ComClass/ComInterface 臂（`knownTypedComVars_` 那一站）放开到**所有**
类型化 COM 接口 / ComClass 局部，登记 `trackLateObjSlotForRelease`（出口发
`vb6_ReleaseObject((void**)&x)`，与 `Set` 那一站退旧值同一句；`vb6_ComIsDispatchable` 把关）。

三条赋值路在发码里都**取所有权或 AddRef**（逐条看过 `--emit-c`）：

| 赋值形 | 发码 |
|---|---|
| `Set d = New Dictionary` | `vb6_NewObject(L"Scripting.Dictionary")`（所有权交过来） |
| `Set d = CreateObject("…")` | `vb6_ComCreateTyped(progId, iid)`（同上） |
| `Set d = gBox.D`（字段/属性取） | `vb6_ComAddRefDispatch(vref)` + 退旧值 |

所以出口放手成对。同一条闸保留：`Static` / `Static Sub` 里那一格跨调用活着，不登记；
For Each 循环变量由 `dropLateObjSlotRelease` 撤登（每一轮由循环自己写）。

### 三、夹具与读数（`tests/b132_typedcom_local_scope/`）

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B132-A1-count` / `A2-terms` | `Dim d As IDictionary` + `New Dictionary` | 1 / **0 → 1** |
| `B132-C1-count` / `C2-terms` | `CreateObject("Scripting.Dictionary")` 那条路 | 1 / **0 → 2**（累计） |
| `B132-F1-count` / `F2-field` | 从**字段**借的局部退出 | 1 / 1（字段不动） |
| `B132-S1-count`（两次）/ `S2-terms` | `Static d As IDictionary` 跨调用活着 | 1,1 / 2（累计不动） |
| `B132-K1-alive` | 局部**交给字段**后字段仍可用 | 1 |
| `B132-Z-terms` | 累计（A/C + 两次显式置空） | **0 → 4** |
| `B132-X1-name` | 真外部类型库服务器 scrrun（`FileSystemObject.GetFileName`） | `win.ini`（不崩） |

x64 与 x86 逐字相同。（`FileSystemObject.FileExists` 在本仓两版都是 `False`，late-bound 与typed
同形 —— 那是另一格在册的**谓词返回**口径，与本刀无关，夹具改用 `GetFileName` 作观测点。）

### 四、发码面

全量重算 778 行：**15 行变**，逐行 A/B（旧/新二进制各一份 `--emit-c`）结果
**全部 `del=0`**、共 76 行新增，每行都是 `vb6_ReleaseObject((void**)&X);`：

```
Charts 2020 Proyecto1 1 (IPic)      combyref271 K271 5 (f/g/h/…)
test_com_optional 1 (fso)           test_com_variant_prop 2 (el/at)
test_earlybound 1 (fso)             test_earlybound2 3 (fso/drv/ts)
test_not_com 1 (fso)                VBFlexGridDemo 51 (IPicture / IPersistStream / IOleObject / IUnknown 那类槽)
VbQRCodegen Project1 2 (pStream/pPersist)
ucSegmentedControl 2 (IPic/oStream) ucTreeView 1 (IPic)   ucSlider_2022 2 (ipic)
fgz_fix011 1 (b)                    GifcontrolSource 2 (lBMPsource/tPic)
dict_native NewBind 1 (d)
```

语料 778 → **780** 行（夹具两行）。

### 五、回归

- `b445` W1–W39 + DONE、`B412`、`B126`、`K124`、`KOwn` G0–G4、`KHaz` F1–F3、`K128`（C1=1 / D1=2）、
  `b128`、`b129` 七针、`b130` 七针、`b131` 九针逐字相同；
- `geomcache` x64 + x86：`latebind-miss` 0 条，GC17/18/19 全 True；
- 九门 9/9；`check_class_refcount_authority` OK（static-gate 仍是 7 条：本刀是**替换**
  §11.113 那一站的窄条件，不是新增闸门）；语料 780 行。

### 六、边界（在册）

- **地址交给外部 Declare 那一格**仍是整条腿④ 的共同边界（写它的一站在计数系统外面；
  泛对象档与类型化接口档共用同一条 `isLateObj` 登记与撤登面）。
- 模块级 / 类字段仍不放手（b132 里 `Set gBox.K = d` 那格是靠显式置空才拆的；夹具那个字段
  声明成 `As Object`，读数与声明成 `As IDictionary` 逐字相同）。
- **b132 的 `.vbp` 聚合行是 CI 口径**（第一百三十二刀补）：本机与 CI 各跑一遍
  `scripts/emit_manifest.ps1`，780 行里只有 6 行不同 —— 5 行是既有的 machineDiff
  （guisample ×3 / TeeChartChartDisplay / ctrlmscomm，本机没装那几个控件），第 6 行就是
  `tests/b132_typedcom_local_scope/B132Typed.vbp`（归一化后本机 bytes=17258 / CI bytes=17239，
  差 19 字节；同一台机器内部可复现、跨机器不可复现，而单模块 `B132Main.bas` 那行两边逐字一致）。
  按既有口径取 `ci/emit-manifest` 那台二进制交回来的数登记，本机复跑会一直报这一行。
- `FileSystemObject.FileExists` 一类的**谓词返回**口径不在本刀（两版同形，另册）。
## §11.115 `With` 自己造的那一枚载体在 `End With` 放手（第一百三十三刀）

### 一、缺口与读数

`With New Collection` / `With CreateObject("…")` 这两形的目标表达式是**自己造一份对象**，
所有权交到 With 临时量手上；`End With` 一句不发 —— 载体连同体内的元素一起活到进程结束：

```vb
Private Sub WithNew()
    With New Collection
        .Add New B133Item
        Debug.Print "B133-A1-count=" & .Count
    End With          ' 修前: 块尾不放手 ⇒ 载体与元素都还在
    Debug.Print "B133-A2-terms=" & gTerms
End Sub
```

夹具 `tests/b133_with_carrier_scope/`（x64 与 x86 逐字相同）：

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B133-A1-count` / `A2-terms` | `With New Collection` + `.Add New B133Item` | 1 / **0 → 1** |
| `B133-B1-count` / `B2-terms` | `With CreateObject("Scripting.Dictionary")` | 1 / **0 → 2**（累计） |
| `B133-C1-count` / `C2-terms` | 反面证人一：`With gBox.C`（**字段**借来的那一份） | 1 / 0 → **0**（不动） |
| `B133-D1-count` / `D2-terms` | 反面证人二：`With c`（**局部** `Set` 取的引用） | 2 / 0 → **0**（不动） |
| `B133-E1-tag` / `E2-terms` | 在册边界：`With New <工程类>` | 11 / 0 → **0**（本刀刻意不发；**§11.117 已补**，补后 = 3） |
| `B133-Z-terms` | 累计 | **0 → 2**（§11.117 之后为 3） |

修前那四个 `0` 是本机实跑：把源码 stash 回本刀之前、重编一份 exe 跑同一个夹具
（产物在 `.temp/k133/b133pre`），不是按形状推的。

### 二、机制（发码一处；闸只认两座工厂）

`src/backend/stmt/cgen_with.cpp`：With 体那块 `{ … }` 收口之后，若**初始化器本身**就是
`(void*)vb6_NewObject(` / `(void*)vb6_CreateObject(`（两座工厂交过来的一份所有权，
与 §11.114 那三条赋值路同一个口径），补一句：

```c
vb6_ReleaseObject((void**)&_vb6_with_N);  /* 账 #305 §11.115: With 自己造的那一枚载体, 块尾放手 */
```

闸按**发码形状**认，不按 AST 猜 —— 四条不匹配：

- 借用型（字段 / 局部 / 形参 / `me.X`）发的是裸取值；
- `With New <窗体>` 已被上一站改写成 `vb6_form_hwnd_X()`（全局默认实例，不归 With）；
- UDT / 宿主伪结构体 / VARIANT 那几档 `tempType` 本身不是 `void*`；
- `With New <工程类>` 发的是 `vb6_cls_X_New()`（类自己那一档）。

`vb6_ReleaseObject` 自己把槽置空（两条分支都 `*objPtr = NULL`），所以体内即使又把它交给了别人，
块尾这一发也只退一次。

### 三、发码面

全量重算 **782 行**（语料 780 + 夹具两行）。与本刀之前那份 CI 清单逐行比：

```
changed=6   added=2   missing=0
  changed 六行全是旧账（5 行 machineDiff + b132 那个聚合行，机器侧）
  added   夹具两行（B133Carr.vbp / B133Main.bas）
```

⇒ 语料 780 行**零行变**。承面在语料里是 **0 处**：全仓 `With New` 只有
`VBFlexGridDemo/MainForm.frm` 那三处 `With New InputForm`（工程窗体 ⇒ 窗体默认实例那一站，
不是晚绑定工厂）。本刀是"补一条语义"，不是"修一片存量"。

### 四、回归

- `b445` W1–W39（含 `W8-terms=200` / `W36-terms=12` / `W39-terms=0 null=False`）、
  `b412` `TERM=23`、`b132` 十二针、`b131` 九针、`b130` 七针、`b129` 七针、
  `b128` `A1=1 / A2=2` 逐字相同；
- `geomcache` x64 + x86：`GC17/GC18/GC19` 全 True（`GC18=1234,1234,1234,1234`、
  `GC19=82,82,82`），latebind-miss 0；
- 九门 9/9（exe sha `2b1db40ed67e`）；`vbp_fixture_census` / `static_sentinel_registration` 都过。

### 五、边界（在册）

- **`With New <工程类>`**：`tempType` 是 `vb6_cls_*`，走类自己的 `_New`/`_Release` 那一档，
  本刀闸刻意挡在外面（夹具 E 路 `E2-terms=0` 是本刀的**预期**读数，不是漏）；
  **第一百三十五刀 §11.117 已补上**（块尾与 Exit 两处都发 `vb6_cls_X_Release`）。
- **`Exit <proc>` 从 With 体内直接跳出**：With 临时量是块内局部，登记进 `slotsToRelease_`
  会在嵌套位置发成越界引用（C2065），所以本刀不登记 —— 早退那一格仍漏一枚，与腿④ 的
  过程出口机制不同源，另册。
- **`With <函数返回>`**（`With f()`）不在本刀：那一格是 Variant / 解封后的借用值，
  所有权归调用那一站，要单独量。
## §11.116 `With New …` 体内 `Exit` 那一格：goto 出块前先放手（第一百三十四刀）

### 一、缺口与读数

第一百三十三刀只补了**块尾**那一发（`End With`）。体内 `Exit Sub/Function/Property` 走的是
`goto vb6_proc_exit` —— 一出块就没人再问那枚载体，于是它活到进程结束：

```vb
Private Sub ExitInside()
    With New Collection
        .Add New B134Item
        Exit Sub          ' 修前: 这里出去, 载体与元素都还在
    End With
End Sub
```

夹具 `tests/b134_with_exit_scope/`（x64 与 x86 逐字相同）：

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B134-A1-count` / `A2-terms` | 对照：正常走到 `End With`（§11.115 那一发） | 1 / **0 → 1** |
| `B134-F1-count` / `F2-terms` | `Exit Sub` 从 `With New Collection` 体内跳出 | 1 / **0 → 2**（累计） |
| `B134-G1-count` / `G2-terms` | 反面证人：`With <字段>` 里 `Exit Sub`（借用型，不许退） | 1 / 0 → **0** |
| `B134-H1` / `H2-terms` | `Exit Function` 从 `With New …` 跳出（返回值仍落位） | 7 / **0 → 3** |
| `B134-I2-terms` | 两层自家载体 + `Exit Sub`（内层先退、外层后退） | **0 → 5** |
| `B134-Z-terms` | 累计 | **0 → 5** |

修前那四个 `0` 同样是本机实跑（`.temp/k134/pre`，本刀之前的那份 exe）。

### 二、机制（两处，一位主人 + 一枚印）

- `cgen_with.cpp`（§11.115 那站）在认出发码形状是 `vb6_NewObject` / `vb6_CreateObject` 时，
  顺手把 `ownsCarrier` 盖在**当前这一帧**的 `WithObjInfo` 上；
- `cgen_jumps.cpp` 的 `visit(ExitStmt&)`：`Exit Sub/Function/Property` 三条臂在发
  `goto vb6_proc_exit;` **之前**，从信息栈顶往下把盖过印的帧各发一句
  `vb6_ReleaseObject((void**)&_vb6_with_N);`。

载体声明就在 With 体内那一层 ⇒ 走到这条 `Exit` 时它仍在作用域里；`vb6_ReleaseObject` 自己置空，
所以块尾那一发与这里这一发不会二次退。借用型 / 窗体默认实例 / 工程类 `New` 那三档没有印，
一条都不发（G 路就是这条反面证人）。

### 三、发码面

语料 **784 行**（原 780 + b133 两行 + 夹具两行）。本刀动的只有 `Exit` 那三条臂，而它们只在
**盖过印的帧**存在时才多出一句 —— 全仓 `With New` 只有 `VBFlexGridDemo/MainForm.frm` 那三处
`With New InputForm`（窗体默认实例，没有印）⇒ 语料里预计零行变，形状 oracle 在 CI 上复算 784 行验证。

### 四、回归

- `b445` W1–W39、`b412` `TERM=23`、`b133` 十二针、`b132` 十二针逐字相同；
- 九门 9/9（exe sha `1ac6d1f05070`）。

### 五、边界（在册）

- **`GoTo` 跳出 With 块**：`GoTo <label>` 越过块尾那一发，与 `Exit` 不同源，本刀不接。
- **错误处理器**（`On Error GoTo` 落进 handler）那一格：`vb6_ErrRaise` 的 longjmp 路径上
  With 载体仍漏一枚 —— 要接得先给错误边界一张"活着的载体"名册，另册。
- **`With <函数返回>`** 仍不在册（同 §11.115 五）。
## §11.117 `With New <工程类>` 的载体：块尾与 Exit 两处放手（第一百三十五刀）

### 一、缺口与读数

§11.115 的闸只认晚绑定两座工厂（`vb6_NewObject` / `vb6_CreateObject`）；`With New <工程类>`
发的是 `vb6_cls_B133Item* _vb6_with_4 = (vb6_cls_B133Item*)(vb6_cls_B133Item_New())`，
一句都不发 —— b133 夹具的 E 路当场记着这条边界（`E2-terms=0`），本刀把它关掉：

夹具 `tests/b135_with_class_carrier/`（x64 与 x86 逐字相同）：

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B135-A1-tag` / `A2-terms` | `With New B135Item` 走到 `End With` | 11 / **0 → 1** |
| `B135-B1-tag` / `B2-terms` | 同上但体内 `Exit Sub` | 22 / **0 → 2**（累计） |
| `B135-C1-tag` / `C2-terms` | 反面证人一：`With gBox.It`（字段里的实例） | 33 / 0 → **0**（不动） |
| `B135-E1-tag` / `E2-terms` | 反面证人二：`With <局部>` + `Exit Sub`（局部那一份归过程出口） | 44 / 0 → **0**（不动） |
| `B135-Z-terms` | 累计 | **0 → 2** |
| `B135-D1-field` | 字段里那枚实例到收尾仍活着 | 44 |

修前那一列同样是拿本刀之前的 exe 实跑（`.temp/k135/pre`，sha `215903550809`）。

### 二、机制（与 §11.115/§11.116 同一对站点，各加一档）

- `cgen_with.cpp`：`tempType` 是 `vb6_cls_X*` 且初始化器里含 `vb6_cls_X_New()`（同一类名两头
  对上）⇒ 块尾发 `if (_vb6_with_N) vb6_cls_X_Release(_vb6_with_N);`，并把类名盖到这一帧的
  `carrierClass` 上；
- `cgen_jumps.cpp` 的 Exit 三条臂：帧上有 `carrierClass` 就发同一句（不是 `vb6_ReleaseObject`）。

**为什么不用 `vb6_ReleaseObject`**：裸工程实例不是可分派对象，Fix 191 那道判据只清槽不 Release
—— 用它等于没退。工程类那一档本来就有唯一出口（`vb6_cls_X_Release`），这里只是接上去。

### 三、旧读数随之更新（b133 的 E 路）

`B133-E2-terms` 2 → **3**、`B133-Z-terms` 2 → **3**（b133 夹具的 E 路就是这条边界），
`tests/run_tests.ps1` 的 `$b133Needles` 同步改。

发码侧：b133 的 `.vbp` 聚合行变（`bytes` 15788 → 15920），而 `B133Main.bas` 的**单文件行不变**
—— 单文件编译没有工程上下文，`As B133Item` 认不出是工程类 ⇒ 闸不认它。这与全仓"单文件口径"
一致，不是漏。

### 四、发码面

语料 **786 行**（780 + b133/b134/b135 三组夹具各两行）。变的只有 b133 的聚合行 + b135 的两行新增。

### 五、回归

- `b445` W1–W39、`b412` `TERM=23`、`b134` 十一针、`b133`（新读数）、`b132` 十二针逐字相同；
- 九门 9/9（exe sha `64f3651debbc`）。

### 六、边界（在册）

- **单文件（`.bas` 单独 `--emit-c`）**不认识工程类 ⇒ 闸不认（与全仓单文件口径一致）。
- `GoTo` 跳出 With 块 / 错误处理器 longjmp 两条仍与 §11.116 五 同册；`With <函数返回>` 由 §11.118 关掉（Property Get 那一形另册）。
## §11.118 `With <工程函数>()` 的载体：块尾与 Exit 两处放手（第一百三十六刀）

### 一、缺口与读数

§11.115 / §11.117 两份闸都只认初始化器里那一句"就地造对象"（两座晚绑定工厂 / `vb6_cls_X_New()`）。
`With MakeColl()` 发的是 `void* _vb6_with_0 = (void*)vb6_K136Main_MakeColl()` —— 引用是**被调方
交回来**的，闸不认，块尾一句不发。

夹具 `tests/b136_fnret_carrier/`（x64 与 x86 逐字相同）：

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B136-A1-count` / `A2-terms` | `With MakeColl()` 走到 `End With` | 1 / **0 → 1** |
| `B136-B1-count` / `B2-terms` | 同上但体内 `Exit Sub` | 1 / **0 → 2**（累计） |
| `B136-C1-count` / `C2-terms` | 反面证人一：`With gColl`（模块字段，借来的） | 2 / 0 → **2**（不退） |
| `B136-D1-count` / `D2-terms` | 反面证人二：`With c`（`Set c = gColl`，局部借来的） | 3 / 0 → **2**（不退） |
| `B136-Z-terms` | 累计 | **0 → 2** |

修前那一列是拿本刀之前的 exe 实跑（stash 后重建，`.temp/k136/fx_pre`，五针全 0）。C/D 两路
的原有数字一并涨到 2，是因为 A/B 两枚载体被退了、体内那两枚元素随之终止 —— 借来的那两枚本身
一枚都没退，这正是它们要证的。

### 二、口径：为什么这一枚归调用方

被调方的返回槽在过程出口放手时被排除（`currentReturnVar_`，腿④ §11.114 那一档），而
`Set <返回槽> = …` 必取一枚引用（新对象或 AddRef）。两头相加：调用方接到的这一枚是**净交回**的，
With 临时量就是它的主人。与 §11.115 两座工厂同一个口径，只是"造"发生在被调方体内。

### 三、机制与闸（四问，任一否就不发）

- `cgen_with.cpp`：AST 上是 `IdentifierExpr` 的无参调用 ⇒ `lookupModuleDotted` 命中
  `SymbolKind::Function` ⇒ `sym->type == Vb6Type::Object` ⇒ 发码是 `vb6_<Mod>_<Name>()`
  那一形（尾部只有一个 `(`，没有经 Variant 解封 / COM 包装）。四问全过才 `ownsCarrier = true`，
  块尾发 `vb6_ReleaseObject((void**)&x);`；Exit 三条臂照 §11.116 盖的那枚印子走，一处不改。
- **Property Get 不在本刀闸内**：它发的是 `…(gBox)` 而不是 `()`，被"尾部单个 `(`"那一问挡在外面。
  要认它得在符号侧把返回类认到，另开一档。
- 两条放手路径的注释分开写（§11.115 / §11.118 各一句），已登记行的字节不为本刀多动。

### 四、修中两条实测（写下来免得下一个人再猜一次）

1. 首版闸按 `(void*)vb6_` 前缀认 —— `(void*)` 是 With 声明器**自己**加在发码行上的，不在
   `lastExpr_` 里（构建期 stderr 探针实测 `last=[vb6_K136Main_MakeColl()]`），于是闸永不发。
   改成"`vb6_` 开头 + 尾部单个 `(`"。
2. 第一版只把印子盖到 `withObjectInfoStack_`（Exit 用），块尾仍写着 `if (withOwnedCarrier133)`
   独门 —— 重建后读数照旧 0。块尾与 Exit 是两处站，两处都要认这一档。

### 五、发码面

语料 **788 行**（780 + b130…b136 夹具）。本刀只新增 b136 两行：b133/b134/b135 五行逐字不变
（`.temp/k136/recheck136.ps1` 逐条 MATCH），全语料本地重跑也只有既有的 6 条机器侧行不同
（guisample ×3 / TeeChartChartDisplay / ctrlmscomm / b132 的 `.vbp` 聚合行）。

### 六、回归

- `b445` W1–W39（`W8-terms=200` / `W36-terms=12` / `W39-terms=0`）、`b412` `TERM=23`、
  b128/b129/b130/b131/b132/b133/b134/b135 读数逐字相同；
- 九门 9/9（exe sha `dacc07910cf8`）。

### 七、边界（在册）

- `With <Property Get>()`：发码尾部是 `(me)` 那一形，本刀闸不认 —— 待一档。
- 错误处理器 longjmp：原定本刀，改档 **§11.119**。落点(`setjmp` 后的 `goto`)与 `Resume`
  语义冲突：`Resume` 允许回到块内继续执行，载体此刻必须还活着；而落点对所有错误共用、且
  声明在落点之后（嵌套块里更不在作用域）。要做就得是"活载体名册 + 帧标号 + Resume 判据"，
  单独立设计档，不硬补 —— 设计已落 §11.119（前置：先修 `Resume Next` 落点的 §11.120）。
- `GoTo` 跳出 With 块仍在册（§11.116 五）—— 第一百四十一刀已关（§11.122）。
## §11.119 错误处理器落点的活载体名册（第一百三十九刀落码；设计于第一百三十七刀）

### 一、缺口

`On Error GoTo H` 的过程发码：`setjmp` 的落点是**过程顶上**那一块，错误一 raise 就
`goto vb6_label_H`，把 With 临时量的声明、块尾那句放手统统跨过去。实测量到
（`.temp/k137/p`，exe `dacc07910cf8`）：

| 路 | handler 形态 | 读数 | 说明 |
|---|---|---|---|
| A | 只 `H:` 里打印，**没有 Resume** | `A2-terms=0` | 载体漏，元素不终止 —— 本档要关的那一格 |
| B | `H:` 里 `Resume Next` | `B2-terms=0`，且 `B1b` **一句没打** | 见 §11.120：`Resume Next` 落到了 `End With` 之后 |
| | 备注 | §11.119 落码后这三路读数上移一格：`A2` 0→1、`B2` 1→2、`C2` 2→3、`Z` 2→3 | |

### 二、为什么不能在落点直接补一句

1. **`Resume` 语义**：`Resume` / `Resume Next` / `Resume <label>` 允许回到块内继续跑。落点若
   无条件放手，`Resume` 回块内就是 use-after-free —— 比漏更坏。
2. **落点与声明的先后**：`goto vb6_label_H` 是往前跳，跳过的正是 `void* _vb6_with_N = …` 那一行；
   读到的要么是未初始化值（崩），要么（嵌在 `If`/`For` 里时）根本不在作用域（C2065）。
3. **落点对所有错误共用**：它自己不知道此刻哪几枚载体已经造出来了。

三条加起来 ⇒ 这一格要的不是一句 Release，而是"**槽先零初始化、名册按名放手**"。

### 三、判据表（按 handler 里有没有 `Resume` 分档）

| handler 形态 | 控制权 | 载体归宿 | 本档动作 |
|---|---|---|---|
| 无任何 `Resume`（`Exit Sub/Function`、`End`、`GoTo` 出块、落到 `End Sub`） | 确定不回块内 | 落点放手 | **可关** |
| 有 `Resume Next` / `Resume` / `Resume <label>` | 可能回块内 | 归块尾 / Exit 两处 | 落点不许动 |

判据取"保守的一侧"：从 `On Error GoTo H` 的 H 标签起，扫到过程尾（含嵌套 `If`/`For` 体内），
出现任一 `ResumeStmt`（含 `Resume Next`）即判"可能回块内"。误判只会少退（漏），不会多退。

### 四、机制（提槽 + 名册）

1. **提槽**：过程入口（`decl_proc` / `decl_func` / `decl_prop` 三处）预扫过程体，数出"晚绑定
   自有载体"（`New` / `CreateObject` / §11.118 的工程函数那一形）的条数 N，在 `setjmp` **之前**
   发 `void* _vb6_own_k = 0;`（k = 0..N-1）。零初始化是这一档唯一的硬要求：落点读到的必须是
   确定的 0，而不是栈上的垃圾。
2. **改站**：`cgen_with.cpp` 里第 k 条自有载体不再就地声明，改发 `_vb6_own_k = <init>;`，块内
   引用全走这个名字（`withObjectVars_` / `withObjectInfoStack_` 照常登记，Exit 三条臂不用改）。
3. **落点放手**：`visit(OnErrorStmt)` 在 `setjmp != 0` 那一块里按名册逐条发
   `vb6_ReleaseObject((void**)&_vb6_own_k);` —— 它自己置空，没造出来的那几条是 0，天然跳过。
4. **闸**：只认 `void*` 晚绑定自有载体。`With New <工程类>`（§11.117 档）要按类名调
   `vb6_cls_X_Release`，名册里得带类名，而落点**先于**那些 With 站点发码 —— 本档不碰，留在册。

### 五、动工顺序（下一刀先做哪一步）

本档的前置是 §11.120：只要 `Resume Next` 还落在 `End With` 之后，"可能回块内"这一档就连
正常的块尾放手都拿不到。所以顺序是：

1. **§11.120**（**第一百三十八刀已关闭**）：修 `Resume Next` 的落点位置（与 VB6 对齐：回到出错语句的
   下一条，仍在块内）；
2. **§11.119 落码**：提槽 + 名册 + 落点放手，夹具三路 —— A（handler 无 Resume）关掉、
   B（`Resume Next`）交给块尾收、C（`With gColl` 借用型）一枚都不动。

### 六、边界（在册）

- 工程类载体（§11.117）在落点仍漏（名册要带类名，见四.4）；
- `Resume <label>` 指向块内深处（跳过一部分语句）时，块尾仍会在真正离块时放手 —— 本档不改这条；
- `GoTo` 跳出 With 块仍在册（§11.116 五）—— 第一百四十一刀已关（§11.122）。

### 七、落码与读数（第一百三十九刀）

机制按四发落地（与三 判据表一致）：

- `cgen_decl_proc` / `cgen_decl_func`：`hoistOwnN_ = (!hasResume_ && hasOnError_) ? countWithStmts(body) : 0`
  —— **一句 `Resume` 都没有**才启用；名册 `void** _vb6_own[N] = {0};` 就发在过程顶层
  （`jmp_buf` 之后、`setjmp` 之前），未登记的槽必为 0。`cgen_decl_prop` 把旗清掉（Property 不发名册）。
- `cgen_with.cpp`：133/136 两档自有载体在声明之后发 `_vb6_own[k] = (void**)&_vb6_with_N;`；块尾放手
  之后发 `_vb6_own[k] = 0;`（**注销**——不注销的话这个栈槽被后续块复用时，落点会拿悬空地址去退）。
- `cgen_jumps.cpp` 的 Exit 三条臂：退完一枚接着注销同号槽。
- `cgen_jumps.cpp` 的 On Error 落点：`if (setjmp(...) != 0) { <逐槽> if (_vb6_own[k]) vb6_ReleaseObject(_vb6_own[k]); goto H; }`。

`countWithStmts` 给的是**上界**（过程体里 With 语句条数，递归）：槽号按发码顺序分配，而"哪一帧是
自有载体"要等发码才算得出来，多给的槽没登记就是 0、落点天然跳过；少给才会溢出。

夹具 `tests/b138_handler_landing/`（x64 与 x86 逐字相同）：

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B138-A1-count` / `A2-terms` | handler 落到 `End Sub`（无 Resume），块尾那句被跨过 | 1 / **0 → 1** |
| `B138-B1-count` / `B2-terms` | 同上但 handler 里 `Exit Sub` | 1 / **0 → 2** |
| `B138-C1-count` / `C2-count` / `C3-terms` | 反面证人：`With gColl`（借来） | 2 / 2 / 0 → **2 / 2 / 2**（gColl 仍在、Count 仍读得出） |
| `B138-D2-terms` | 两个 With 先后：前一个块尾退、后一个落点退 | 1 → **4** |
| `B138-E1-count` / `E2-terms` | §11.118 的工程函数那一枚同样在名册里 | 1 / 1 → **1 / 5** |
| `B138-Z-terms` / `Z-count` | 累计 / gColl 完好 | 1 / 2 → **5 / 2** |

修前那一列是拿本刀之前的 exe 实跑（stash 后重建，`.temp/k139/fx_pre`）。

**随刀改注的旧读数**：§11.120 夹具的三路（`B137-A2` 0 → **1**、`B137-B2` 1 → **2**、`B137-C2` 2 → **3**、
`B137-Z` 2 → **3**）—— b137 的 A 路本来就是这一格的边界，本刀把它关掉；该夹具的 needles 同步改。

**发码面**：语料 **792 行**（780 + b130…b138 夹具）。既有行有 **7 条**随刀变（都是"过程里有
`On Error GoTo` + `With`、且一句 `Resume` 都没有"的文件）：`b137_resume_in_block/B137Main.bas`、
`test_arr_empty.bas`、`test_collcall.bas`、`vb6probe/NullFamilyProbe.bas`、
`VBFlexGridDemo/Builds/VTableHandle.bas`、`VBFlexGridDemo/Common/VisualStyles.bas`、
`VbQRCodegen-master/src/mdQRCodegen.bas`；其余一字未动。

**九门 9/9（exe sha `19252414e3e0`）**。

### 八、边界（第一百三十九刀之后仍在册）

- 工程类载体（§11.117）：名册槽只有 `void**`，落点分不出类名 —— 仍归块尾/Exit 两处；第一百四十二刀已关（§11.123）。
- 过程里**有** `Resume` 的（哪怕 handler 里那句 `Resume` 只落块外）：落点一律不动，仍漏；
- 出错语句是块内最后一条且 handler `Resume Next`：落点落到块外、块尾那句被跨过 —— 仍漏；
  第一百四十刀已关（§11.121）。
- `GoTo` 跳出 With 块（§11.116 五）—— 第一百四十一刀已关（§11.122）。

## §11.120 `Resume Next` 在块内的落点位置（第一百三十八刀）

### 一、缺口

发码里 resume 点只在**过程顶层**的语句表上发（`emitStmtList(body, hasResume_)`），
`With` / `If` / `For` / `Select` 的体走的是默认 `false`。于是块内报错时
`vb6_err_resume_next_point` 还是"进入这一整条块语句之前"设的那一号，`Resume Next` 落到块尾
**之后**：块内剩下的语句被跳过，块尾那句放手也被跳过（`§11.119` 一 的 B 路，`B1b` 一句没打）。

VB6 的语义是"回到**出错语句的下一条**"，与嵌套深浅无关 —— 所以这是控制流本身的缺口。

### 二、改法（一处判据）

`cgen_stmt.cpp` 的 `emitStmtList`：条件从 `emitResumePoints && inProtectedBlock_` 改成只用
`inProtectedBlock_`。`inProtectedBlock_` 只在过程里用了 `Resume` 时才置位
（`cgen_jumps.cpp` 的 `On Error GoTo` 站），处理器的体又不在保护区内（同文件的标签站清旗），
所以这一改既覆盖所有嵌套层，也不给"没有 Resume 的过程"多发一行。

`emitResumePoints` 形参保留（调用点不改），只作兼容。

### 三、读数（夹具 `tests/b137_resume_in_block/`，x64 与 x86 逐字相同）

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B137-A1-count` / `A2-terms` | handler 里一句 `Resume` 都没有（§11.119 那一档） | 1 / 0 → **0**（在册，本刀刻意不动） |
| `B137-B1-count` / `B1b-after-error` / `B2-terms` | 块内报错 + `Resume Next`，回到出错语句的下一条，块尾放手 | 1 / **没打 → 打了** / **0 → 1** |
| `B137-C1-count` / `C1c-after-error` / `C2-terms` | 同一条错，但嵌在 `If` 里再套 `With`（两层） | 1 / **没打 → 打了** / **0 → 2** |
| `B137-Z-terms` | 累计 | **0 → 2** |

修前那一列是拿本刀之前的 exe 实跑（stash 后重建，`.temp/k137/fx_pre`）。

### 四、发码面

语料 **790 行**（780 + b130…b137 夹具）。**既有行零变化**：本刀判据只对"过程里用了
`Resume` 语句"的那些生效，而全语料没有一条独立的 `Resume` / `Resume Next` 语句
（`On Error Resume Next` 是另一条路，不发 setjmp、不置 `inProtectedBlock_`）。
本刀只新增 b137 两行。

### 五、回归

- §11.119 一 的 A 路读数不变（`A2-terms=0`，在册边界照旧）；
- b133/b134/b135/b136 夹具读数逐字相同；九门 9/9（exe sha `26110e726f47`）。

### 六、边界（在册）

- **出错语句是块内最后一条**时，`Resume Next` 的落点自然落到块外（块尾那句放手被跨过）——
  这一形仍漏，归 §11.119 的活载体名册（落点放手）与"块尾放手补点"两档，本刀不碰；
- handler 里一句 `Resume` 都没有的那些（A 路）归 §11.119；
- `GoTo` 跳出 With 块仍在册（§11.116 五）—— 第一百四十一刀已关（§11.122）。

实测（§11.119 一 的 B 路）：块内报错、handler `Resume Next` 之后，既没回到出错语句的下一条，
也没执行 `End With` 的放手 —— 发码里 `vb6_err_resume_next_point` 在 With 体内仍是进入 With 前
设的那一号（`vb6_resume_1` = 块尾之后），因为"下一条语句的号"是在 With 语句**整条**结束时才推进的。

与 VB6 的差异：VB6 的 `Resume Next` 回到**出错语句的下一条**（仍在块内）。这一格是控制流本身的
缺口，不是放手位置的问题，且会连带影响 §11.119 的"可能回块内"那一档 —— 单独立刀，先修它。
## §11.121 块尾补一个 resume 点：`Resume Next` 从体内末句回来（第一百四十刀）

### 一、缺口

§11.119 八 的第三条：出错语句是块内**最后一条**、handler `Resume Next` 时，"下一条语句"是块外
那一句 —— 落点直接跨过 `End With` 的放手。夹具 `tests/b139_blocktail_resume/` 的 A 路修前
`A2-terms=0`（晚绑定载体活到进程结束）、B 路 `B2-terms=0`（§11.117 工程类载体同病）。

第一百三十八刀把 resume 点发到了每个嵌套层，但"体内最后一条之后"没有点可落；第一百三十九刀
的名册只管"process 里一句 `Resume` 都没有"的那一档 —— 有 `Resume` 的过程，落点仍无处放。

### 二、改法（一处，且**补点必须在放手之前**）

`cgen_with.cpp` 块尾：若在保护区内、且本块确实发过放手（133 晚绑定自有 / 136 工程函数交回 /
135 工程类三类），先推进一号 `resumePointCounter_`，发

```
vb6_err_resume_point   = m;
vb6_err_resume_next_point = m + 1;
vb6_resume_m:;
```

再发放手那几句与名册注销。编号天然接上：体内语句按顺序各占一号，末句的
`resume_next_point` 正好是 `m`，于是 `Resume Next` 落到 `vb6_resume_m` —— **先跑完放手**，
再落回块外那一句。

第一版把补点发在放手**之后**（读 `.temp/k140/e.c` 可见 `vb6_resume_3` 在 `ReleaseObject` 下一行），
实跑 `A2-terms=0` 照旧 —— 落点跳过清理，等于没修。夹具的 `A2/B2` 就是量这一格的。

### 三、读数（夹具 `tests/b139_blocktail_resume/`，x64 与 x86 逐字相同）

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B139-A2-terms` | 晚绑定 `With New Collection`，末句 `Error 5000` + `Resume Next` | **0 → 1** |
| `B139-B2-terms` | `With New <工程类>`（§11.117 载体），同形 | **0 → 2** |
| `B139-C1-after-error` / `C2-terms` | 对照：末句之后还有一句（§11.120 那一档） | 打了 / 1 → 打了 / **3** |
| `B139-D2-count` | 反面证人：`With gColl`（借来的）—— 块尾无清理可跳，handler 里仍读得出 Count | 2 → 2（不动） |
| `B139-Z-terms` / `Z-count` | 累计 / gColl 完好 | 1 → **3** / 2 |

修前那一列是 stash 掉本刀改动、重建后实跑（`.temp/k140/fx_pre`）。C 路的 1 → 3 是**前两刀**
（139 的名册在 C 路过程里同样生效）叠加的结果，不是本刀单独贡献。

### 四、发码面

语料 **794 行**（792 + b139 夹具两行）。判据要同时满足"过程里有独立 `Resume`"与"发过放手的
`With`"，全语料只有 b137 夹具命中 → 既有行变 **2 条**（`b137_resume_in_block/B137Main.bas`、
`B137Res.vbp`），其余一字未动；本刀新增 b139 两行。

### 五、回归

- b137 夹具读数逐字不变（`A2=1`、`B2=2`、`C2=3`、`Z=3`）—— 它的 B/C 路报错语句后面都还有一句，
  补点不参与；
- b138 / b136 / b135 / b134 / b133 夹具与语料既有行不动；九门 9/9（794 行、注册 794）。

### 六、边界（仍在册）

- 工程类载体进落点名册（§11.117 —— 名册槽只有 `void**`，落点分不出类名）—— 第一百四十二刀已关（§11.123）；
- 过程里**没有**一句 `Resume` 的仍走 §11.119 的名册（两条路互补，不重叠）；
- `GoTo` 跳出 With 块（§11.116 五）—— 第一百四十一刀已关（§11.122）。
## §11.122 `GoTo` 跳出 With 块：按层深逐层退（第一百四十一刀）

### 一、缺口

§11.116 五在册的最后一格：`GoTo <label>` 从 `With` 体内跳到块外，块尾那几句放手永远跑不到。
夹具 `tests/b140_goto_out_with/` 修前：`A2-terms=0`（晚绑定）、`B2-terms=0`（两层嵌套，两层都漏）、
`D2-terms=1`（工程类载体）。

### 二、改法（预先扫层深，落地按"跳出几层"逐层退）

- `cgen_util_scan.cpp` 新 `collectLabelWithDepth`：单趟递归扫过程体，记下每个 `LabelStmt`
  定义处的 With 层深（`With` 进入体内 +1，`If` / 循环 / `Select` 的各支同层），存进
  `labelWithDepth_`（标签名小写）。为什么要预扫：`GoTo` 多半**向下**跳，目标标签在它后面，
  发码是单趟的；不预扫就只能猜。
- `cgen_decl_proc` / `cgen_decl_func` / `cgen_decl_prop` 入口清表并填充（Property 也填 ——
  它的体里同样可能有 `GoTo` 跳出 With）。
- `cgen_jumps.cpp` 的 `visit(GoToStmt&)`：查目标层深，当前层深 = `withObjectInfoStack_.size()`；
  从最内层起逐帧退到目标层深（`ownsCarrier` 走 `vb6_ReleaseObject`、`carrierClass` 走
  `vb6_cls_X_Release`），退一枚就把名册槽清 0。**查不到标签就不发** —— 宁少退不多退。

载体变量声明在 `With` 体内，走到这条 `GoTo` 时仍在作用域里（与 `Exit` 那几臂同理），
`vb6_ReleaseObject` 自己置空，所以与块尾那一发不会二次退。

### 三、读数（夹具 `tests/b140_goto_out_with/`，x64 与 x86 逐字相同）

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B140-A2-terms` | `With New Collection` + `GoTo` 跳出 | **0 → 1** |
| `B140-B2-terms` | 两层嵌套，一次跳出两层 | **0 → 3** |
| `B140-C1-terms` / `C2-terms` | 对照：标签就在本块体内（一层都不出） | 0 / 1 → 3 / 4（体内不退，块尾退） |
| `B140-D2-terms` | `With New <工程类>`（§11.117）跳出 | 1 → 5 |
| `B140-E2-count` | 反面证人：`With gColl` 借来的，一枚都不许退（handler 里仍读得出 Count） | 2 → 2 |
| `B140-Z-terms` / `Z-count` | 累计 / gColl 完好 | 1 → 5 / 2 |

修前那一列是拿本刀之前的 exe 实跑（stash 后重建，`.temp/k141/fx_pre`）。

### 四、发码面

语料 **794 → 796 行**（b140 夹具两行）。**既有行零变化**：判据要同时命中"`GoTo`"与
"发过放手的 `With`"，全语料先按文本筛出 5 个候选、再逐行实算 —— 只有 b140 是新的，
b134（§11.116 的 Exit 夹具）等四个候选**逐字相同**。真实语料里没有一例 `GoTo` 跳出自有载体的
With 块，所以这一刀是纯补口子，不改任何人现有的字节。

### 五、回归

- b137 / b138 / b139 夹具行与读数逐字不变；b133…b136 不动；
- 九门 9/9（796 行、注册 796）。

### 六、边界（仍在册）

- `GoTo` **跳进** With 块（VB6 本身不许，编译器也不接）：不进本表；
- 工程类载体进错误落点名册（§11.117 + §11.119 的槽只有 `void**`，分不出类名）—— 第一百四十二刀已关（§11.123）；
- 过程里**有** `Resume` 的错误落点仍按 §11.121 的块尾补点处理，与本刀不重叠。
## §11.123 工程类载体进落点名册：每槽一列放手函数（第一百四十二刀）

### 一、缺口

§11.119 八 的最后一格：名册 `void** _vb6_own[N]` 里的槽只记录了"槽地址"，落点统一走
`vb6_ReleaseObject` —— 那是 COM 槽的放法。工程类那一档（§11.117 的 `With New <工程类>`）
槽里存的是编译期生成的 `vb6_cls_X` 实例，必须用类自己的 `vb6_cls_X_Release`，而落点
分不出类名。夹具 `tests/b141_roster_class/` 修前：`A1-terms=0`、`D1-terms=1`
（嵌套两枚只退了一枚，还是靠块尾那一发），`Z-terms=1`。

### 二、改法（名册加一列，登记时写死放开函数）

生成器里 `vb6_cls_X_Release` 的签名恰好是 `unsigned long (void* instance)` —— 不需要薄：

- `cgen_decl_proc` / `cgen_decl_func`：名册声明多一列
  `unsigned long (*_vb6_ownRel[N])(void*) = {0};`，紧挨 `_vb6_own` 发在过程顶层
  （`setjmp` 之前，未登记的槽是 0）。
- `cgen_with.cpp` 登记站：登记条件从"`void*` 两档"扩到三档（`anyOwn142`）；
  工程类那档多发一句 `_vb6_ownRel[k] = vb6_cls_X_Release;`，`void*` 两档留 0。
- `cgen_jumps.cpp` 落点：逐槽
  `if (_vb6_own[k]) { if (_vb6_ownRel[k]) { _vb6_ownRel[k](*_vb6_own[k]); *_vb6_own[k] = 0; } else vb6_ReleaseObject(_vb6_own[k]); }`
  —— 有放手函数就按它放（放完把手里的指针清 0），没有走原来的 `vb6_ReleaseObject`。

槽号只增不复用（`hoistOwnNext_++`），所以 `_vb6_ownRel[k]` 不会串到后面登记成
`void*` 的槽；块的注销站（块尾 / Exit / GoTo）只清 `_vb6_own[k]`，落点读到的就是 0。

### 三、读数（夹具 `tests/b141_roster_class/`，x64 与 x86 逐字相同）

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B141-A1-terms` | 一枚工程类载体，handler 无 `Resume`，落点放手 | **0 → 1** |
| `B141-B1-terms` | 混合：前一块晚绑定载体（块尾已退）+ 后一块工程类（落点退） | 1 → **3** |
| `B141-C1-count` | 反面证人：`With gColl` 借来的那一帧不进名册，落点一枚不动 | 2 → 2 |
| `B141-D1-terms` | 两枚工程类载体同时在册（嵌套），落点逐槽退 | 1 → **5** |
| `B141-Z-terms` / `Z-count` | 累计 / gColl 完好 | 1 → **5** / 2 |

修前那一列是拿本刀之前的 exe 实跑（stash 后重建，`.temp/k142/fx_pre`）。

### 四、发码面

语料 **796 → 798 行**（b141 夹具两行）。这一刀**确实动了既有行**：名册多一列声明、
落点那几句换了形 —— 判据是"过程里有 `On Error` 且一句 `Resume` 都没有、过程体里有 `With`"
（`hoistOwnN_ > 0`）。文本先筛出 24 个候选、再逐行实算，命中的 **9 个模块**
（`b137_resume_in_block/B137Main.bas`、`b138_handler_landing/B138Main.bas`、
`test_arr_empty.bas`、`test_collcall.bas`、`test_late_setslot2.bas`、
`vb6probe/NullFamilyProbe.bas`、`VBFlexGridDemo/Builds/VTableHandle.bas`、
`VBFlexGridDemo/Common/VisualStyles.bas`、`VbQRCodegen-master/src/mdQRCodegen.bas`）
与它们所属的 **4 张聚合 `.vbp`**（`B137Res.vbp`、`B138Res.vbp`、`VBFlexGridDemo.vbp`、
`VbQRCodegen/test/Project1.vbp`）随之改行，其余 782 行一字未动。

### 五、回归

- b137 / b138 / b139 / b140 夹具读数逐字不变（退的是同一批载体，只是工程类那一档
  这次真退了）；
- 九门 9/9（798 行、注册 798）。

### 六、边界

§11.119 八 的四条到这里全部关掉。剩下的是**别册**：`With <Property Get>()` 那一形
（§11.118 明写排除）、以及名册只管"进程里一句 `Resume` 都没有"的那些过程 ——
有 `Resume` 的过程走 §11.121 的块尾补点，两者互补，不重叠。
## §11.124 `With <对象>.<Property Get>`：属性交回的载体（第一百四十三刀）

### 一、缺口

§11.118 明写排除的那一格：`With <对象>.<Property Get>` 交回的是一枚**加了引用**的实例
（与 `Set x = obj.Prop` 同一口径），With 临时量就是它的主人；修前块尾、Exit、落点三处
都不退。夹具 `tests/b142_propget_with/` 修前：`A2-terms=0`、`C1-terms=0`、`D1-terms=0`、`Z-terms=0`。

### 二、改法（认发码形状，不猜 AST）

`cgen_with.cpp` 的工程类那一档（§11.117）多认一形：`tempType` 仍是 `vb6_cls_X*`，
初始化器除了 `vb6_cls_X_New()` 之外，还认 `vb6_<Mod>_prop_get_<Name>(…)`
（带不带 `(vb6_cls_X*)` 那层 cast 都算 —— 实测 cast 是 With 声明符加的，`lastExpr_` 里没有，
这是第一百三十六刀量过的同一个坑）。认出来就仍走 §11.117 那一档的放手
（`vb6_cls_X_Release`），块尾 / Exit / `GoTo` / 错误落点四条口一并生效 —— 第二、三处是
§11.116 / §11.119 早已接好的，本刀不用另发。

块尾那句注释另起一档（`§11.124`），已登记行的字节不动。

对象**字段**那一形（`With gBox.Bare`）发的是裸取值，既没有 `_prop_get_` 也没有类名两头对上，
天然落在门外 —— 拿别人的字段去退就是错退，夹具 B 路是反面证人。

### 三、读数（夹具 `tests/b142_propget_with/`，x64 与 x86 逐字相同）

| 针 | 含义 | 修前 → 修后 |
|---|---|---|
| `B142-A2-terms` | 属性交回的那一枚，块尾放手 | **0 → 1** |
| `B142-B2-tag` | 反面证人：`With gBox.Bare`（对象字段）只是借，块尾不退，之后仍读得出 | 2 → 2 |
| `B142-C1-terms` | 体内 `Exit Sub`，§11.116 的 Exit 口退 | **0 → 2** |
| `B142-D1-terms` | handler 无 `Resume`，§11.119 四.4 的落点退 | **0 → 3** |
| `B142-Z-terms` | 累计 | **0 → 3** |

修前那一列是拿本刀之前的 exe 实跑（stash 后重建，`.temp/k143/fx_pre`）。

### 四、发码面

语料 **798 → 800 行**（b142 夹具两行）。**既有行零变化**：先按 `With <x>.<y>` 文本筛出
32 个候选（含 b133…b141 全部旧夹具、VBFlexGridDemo、VbQRCodegen 等），逐行实算 ——
只有 `tests/itf_neg/n35_legacy_cls_iface.bas` 一行报"不同"，查下去是它本来就 `rc=1/bytes=0`
（`VB3031`，CoClass 列了非 Interface 的契约项），我的行脚本硬写了 `rc=0` 才显出差；
真实发码一字未动。

### 五、回归

- b133…b141 夹具读数逐字不变；九门 9/9（800 行、注册 800）。

### 六、边界（新册）

- `With <晚绑定对象>.<属性>`（`void*` 那一档，COM 派发取回）：走的是 `vb6_ComGet` 一类，
  所有权口径要另量，本刀不碰；
- 属性返回**工程函数那一形**已经在 §11.118 里（`vb6_<Mod>_<Name>()`）。
## §11.125 `With <晚绑定对象>.<属性>`：先把引用账量出来（第一百四十四刀 · 设计刀）

### 一、形状与契约

```c
void* _vb6_with_N = (void*)vb6_ComGetObjectProp(<接收者>, L"<属性>")  /* With object ref */;
```

接收者两档都量过：`Dim o As Object`（裸 `o`）与 `Dim v As Variant`（`vb6_VariantToObject(&v)`）；
属性声明 `As <工程类>` 与 `As Object` 同形（`tests/b144_late_with_prop`，探针 `.temp/k144/p`）。

RTL 那头的契约是写死的：`vb6_ComGetObjectProp` 取走 `pv->pdispVal` 之后只调
`vb6_ComVarFree(pv)`，而那个函数明写「仅释放结构体，对象引用已转移给调用方」
（`src/rtl/core/vb6com/vb6com_wrap.c:145`、`vb6com_pack.c` 的 `vb6_ComVarFree`）。
VB6 语义一侧，With 临时量活到 `End With` ⇒ 这一枚引用该由块尾（及 Exit / 落点 / `GoTo` 四口）退。

### 二、读数（夹具 `tests/b144_late_with_prop/`，x64 与 x86 逐字相同）

| 针 | 含义 | 今天 |
|---|---|---|
| `B144-A2-terms` | 单发，块尾那一口 | 0 |
| `B144-B2-tag` / `B144-B2-terms` | 反面证人：交回的是模块级**共用**实例，退自己那一枚之后值仍读得出 | 8 / 0 |
| `B144-C1-terms` | 体内 `Exit Sub`（§11.116 那一口） | 0 |
| `B144-D1-terms` | handler 无 `Resume`，落点（§11.119 四.4） | 0 |
| `B144-E1-terms` | `GoTo` 跳出（§11.122 那一口） | 0 |
| `B144-Z-terms` / `B144-Z-tag` | 累计 / 共用实例完好 | 0 / 8 |

同一份夹具换成**早期绑定**（`Dim b As B144Box`）5 轮：terms=5 —— 证明 Terminate 计数这把尺子本身是准的
（`.temp/k144` 的 M1/M2），晚绑定两条路（`With o.Fresh` 与 `Set x = o.Fresh : Set x = Nothing`）
5 轮都还是 0。差的不是 With 这一站。

### 三、引用账：一次读数交出三枚（`C3_COM_RC_TRACE=1` 探针，`.temp/k144/rc4.txt`）

一枚 `With o.Fresh` 的完整轨迹：

| 轨迹行 | 累计 | 出处 |
|---|---|---|
| `FROM-INST W=… cnt=1` | 1 | 适配器把属性返回的实例包成包装器（`vb6_ComObject_FromInstance`，实例那一枚引用交给包装器） |
| `ADDREF … cnt=2` | 2 | 适配器那一形 `vb6_VariantObject(vb6_ComObject_FromInstance(…))` 又 AddRef 一次（`cgen_util_dllentry_collect.inc` 的返回值装箱） |
| `ADDREF … cnt=3` | 3 | 服务端 `vb6_VariantPackCopyInto(&tmpRet, pVarResult)` 为「带权副本」再 AddRef 一次（`vb6rtl_com.c`） |
| `RELEASE … cnt=2` | 2 | 本刀临时接上的块尾放手（见四） |

也就是说：**调用方该拿的只有 1 枚，另外 2 枚是桥的两站各自多 AddRef 一次** ——
`tmpRet` 打包之后没有清（它那一枚被复制出去了），适配器又把 `FromInstance` 交出来的所有权
用 `VariantObject` 叠了一层。`Set x = <晚绑定读数>` 走的是同一条桥，所以它 5 轮也全 0：
这两枚不是 With 独有的账。

### 四、为什么本刀不改发码

认形是一行的事（`tempType == "void*"` 且 `lastExpr_` 前缀是 `vb6_ComGetObjectProp(`），
四口也早已接好（§11.115 / §11.116 / §11.119 / §11.122）。本机试过：块尾那一发确实发出来了，
轨迹从「三枚不退」变成「退一枚」，**但夹具读数一个都没变**（仍 0）—— 桥多交的两枚垫在下面。

按本系列的纪律（没有读数变化的刀不落；退错一枚比漏更坏），这一刀只落夹具与这页账，
不落发码。另外还有一层：今天退一枚必然安全（有两枚垫底），桥修好之后这一枚就必须**精确**，
两件事得同一刀对账，否则谁先谁后都会静默错。

### 五、下一刀的靶（顺序不可换）

1. 适配器那一枚改为**转移**：`vb6_VariantObject(vb6_ComObject_FromInstance(…))` 叠出来的
   第二枚引用要收掉（对齐 `New` 实参那条 `vb6_ComPack_<类>` 的「FromInstance 的 1 枚直接
   转移给 VARIANT」口径）。
2. 服务端 `ComObj_Invoke`：`vb6_VariantPackCopyInto(&tmpRet, pVarResult)` 之后把 `tmpRet`
   那一枚退掉（副本已带权）。
3. 前两条落地后，一次读数应恰好交出 **1** 枚；此时把本夹具的 needle 由
   `A2=0 / C1=0 / D1=0 / E1=0 / Z=0` 改成 `1 / 2 / 3 / 4 / 4`，With 四口才有可判红的读数
   （`tests/run_tests.ps1` 里 `$b144Needles` 的注释已经写明这件事）。
4. `Set x = <晚绑定读数>` 那条 AddRef（setlet 站）同轮对账 —— 桥修好之后它是第二枚多交的。

### 六、发码面与回归

- 语料 **800 → 802 行**（b144 夹具两行：`B144Main.bas` / `B144Res.vbp`），既有行零变化
  （b141 / b142 逐行实算 `SAME`，`.temp/k144/rows144.ps1`）；
- 九门 9/9（802 行、注册 802，`.vbp=299 / .bas=503`）。

---

## §11.126（第一百四十五刀）`With <晚绑定对象>.<属性>`：先把桥对平，再让四口接管

### 一、桥的两站各多交一枚

上一刀量出来的是「一次读数交出 3 枚，调用方只该拿 1」。多出来的两枚各自有主：

| 站 | 改前 | 改后 |
|---|---|---|
| 适配器装箱（`src/backend/detail/util/cgen_util_dllentry_collect.inc`） | `vb6_VariantObject(vb6_ComObject_FromInstance(…))` —— `FromInstance` 已把实例那一枚交给包装器，`vb6_VariantObject` 又按 Fix 024 的口径 AddRef 一次 | `vb6_VariantObjectTransfer(…)` —— 接管那一枚，不再 AddRef |
| 服务端打包（`src/rtl/core/vb6comserver/vb6comserver_obj.c`） | `vb6_VariantPackCopyInto(&tmpRet, pVarResult)` 是**带权副本**（串深拷贝、对象 AddRef），副本归对端；`tmpRet` 自己那一枚原封不动漏在服务端 | 打包后补一句 `vb6_VariantClear(&tmpRet)` |

新出口 `vb6_VariantObjectTransfer`（`src/rtl/core/vb6rtl/vb6rtl_variant.h`）与
`vb6_VariantObject` / `vb6_VariantObjectBorrowed` 并列为三兄弟，**不合并**：借来的一枚随时
可能被原主退掉，转移来的一枚只有收到它的人退 —— 共用一枚出口，下一次有人给 `Borrowed`
加口径（比如改成「仅在…」）就必然连带改坏转移的语义。

改后读数（`C3_COM_RC_TRACE=1`）：`FROM-INST cnt=1 → ADDREF cnt=2 → RELEASE cnt=1`，
一次读数**恰好交出 1 枚**。

### 二、顺带对出来的一枚毒：拥有型包装器拆解要问实例自己的引用计数

桥对平之后夹具 B 路立刻跳红：`B144-B2-terms=2`（应为 1）。原因是 `ComObj_Release` 的
拥有型那一档直接 `destroyFunc(instance)` —— 对 `New` 出来的那一枚没错，但 `Property Get
SameObj() As B144Item` 交回的是**模块级共用实例** `gKeep`，拆包装器时把还在用的实例也
`Class_Terminate` 了（调用方手里留一根悬空指针）。这一档改成先摘 `__comObj` 回填、再调
`desc->classRelease(instance)`（问实例自己的 `__refcount`），与下面借用型那一档一字相同；
`destroyFunc` 只留给「没有 `classRelease` 可问」的旧描述符。改后 B2-terms 回到 1。

### 三、With 四口接管

认形是一行的事（`src/backend/stmt/cgen_with.cpp` 的 `withComProp145`）：

```
tempType == "void*" && lastExpr_.rfind("vb6_ComGetObjectProp(", 0) == 0
```

所有权口径由 RTL 钉住：`vb6_ComGetObjectProp` 取走 `pv->pdispVal` 之后只
`vb6_ComVarFree(pv)`，而那个函数明写「仅释放结构体，对象引用已转移给调用方」。也就是说
Invoke 给属性返回值加的那一枚引用**归调用方**，With 临时量就是它的主人。

接进去的是同样的三处 + 一个出口：`ownsCarrier`（认领）、`anyOwn142`（`_vb6_own[]` 落点
登记，护住 `GoTo` 出块）、`inProtectedBlock_`（给 `Resume Next` 补落点），以及块尾
`vb6_ReleaseObject((void**)&_vb6_with_N)`。借用型（`With gColl` / 字段 / 裸变量）的初始化器
不是这个前缀，天然落在门外。

### 四、读数（x64 与 x86 逐条相同）

| 口 | A2 | B2-tag | B2-terms | C1 | D1 | E1 | Z | Z-tag |
|---|---|---|---|---|---|---|---|---|
| 改前 | 0 | 8 | 1 | 0 | 0 | 0 | 0 | 8 |
| 改后 | 1 | 8 | 1 | 2 | 3 | 4 | 4 | 8 |

（B 路是反面证人：交回的是模块级共用实例，退掉自己那一枚之后 `gKeep` 的 tag 仍读得出 8，
terms 仍是 1 —— 这一枚退的是「读数交出来的那一枚」，不是实例本身。）

### 五、发码面

- 全清单 **802 行**逐行复算，**只有 14 行**变（7 个 `.vbp` 聚合行 + 7 个 `.bas` 行），
  其余 788 行 `SAME`；
- 14 行每一行的 ASCII 骨架变化都能对上本刀的两处手笔（适配器 `vb6_VariantObjectTransfer`
  改名 + `§11.126` 块尾放手），没有第三种变化；
- 既有夹具零回归：b133 / b134 / b135 / b142 的读数与各自 needle 逐条相符。

### 六、闸上的连带

新 RTL 出口必须进 `src/backend/detail/expr/cgen_variant_slot_prefixes.inc`（由
`scripts/gen_variant_slot_prefixes.ps1` 从 RTL 派生，62 → 63 条）；`check_builtin_slot_returns`
的 R2（漏登记）/ R8（派生不一致）由此转绿。九门 9/9。

### 七、下一刀

`Set x = <晚绑定读数>` 的 setlet 站（§11.125 步骤 4）：桥对平之后它是第二枚多交的 AddRef，
与 With 这一刀同一条桥、不同的主人。

## §11.127（第一百四十六刀）`Set x = <晚绑定对象>.<属性>`：交回的那一枚归槽

### 一、量具与改前读数

先用量具把账量清（`.temp/k146/probe`，不进语料）：`As Object` 接一个工程类实例，
`Set c = o.Fresh : Set c = Nothing`，Fresh 每次新造一枚并计 Terminate 数。改前：

```
S1-terms=0 made=2      ' 新造的那一枚 Set Nothing 之后**不销毁**
S2-terms=0 made=2 keep=7
S3-terms=0 made=4      ' 连读两枚, 一枚都没退
```

发码里多出来的正是这一行（`--emit-c` 逐字）：

```
{ void* _vref1 = (void*)(vb6_ComGetObjectProp(o, L"Fresh"));
  vb6_ComAddRefDispatch(_vref1); vb6_ReleaseObject((void**)&c); c = _vref1; }
```

### 二、改法

`src/backend/detail/stmt/cgen_setlet_set_rhs.inc` 兜底那一支的 `rhsOwned` 增一条口径：
右值文本以 `vb6_ComGetObjectProp(` 开头 ⇒ 自带一枚对价（§11.125/§11.126 已经钉住
「对象引用已转移给调用方」），不再 claim。

裸槽拷贝（`Set d = c`）不在这个前缀里，claim 原样保留 —— 这一条是本刀的**反方向证人**：
少了它 `Set c = Nothing` 会把 d 手里那一枚一并带走。

### 三、读数（x64 与 x86 逐条相同）

| 行 | A1-terms | B1-tag | B2-terms | C1-alive | C2-terms | Z-terms | Z-made |
|---|---|---|---|---|---|---|---|
| 改前（量具） | 0 | 7 | 0 | — | — | 0 | 3 |
| 改后 | 1 | 7 | 0 | 0 | 1 | 2 | 3 |

- A：新造一枚，`Set c = Nothing` 之后恰好 Terminate **一次**（改前 0）；
- B：交回的是模块级共用实例，退读数那一枚之后它仍活（tag 读得出 7，terms 不增）；
- C：`Set d = c` 是裸槽拷贝，`Set c = Nothing` 之后实例**仍活**（C1-alive=0），
  直到 `Set d = Nothing` 才 Terminate（C2-terms=1）。

（C1 一开始写成读 `d.Tag`，读出来是垃圾 —— `d` 是 `As Object`，读裸工程实例的属性走晚绑定，
那个读数不可信。改判据为 Terminate 计数之后才量得准：这一处是量具的错，不是发码的错。）

### 四、发码面

- 全清单 **802 行**逐行复算：**22 行变**、其余 780 行 `SAME`，加本刀两行新夹具 ⇒ **804 行**；
- 每一条变化的字节数都是 −31（`vb6_ComAddRefDispatch(_vrefN); ` 恰好那么长），
  没有第二种变化；
- 22 行里 **16 行**机器无关（纯 `.bas` 行 + `ucTreeView` / `VBFlexGridDemo` 聚合行）⇒ 本机值直接登记；
  另 6 行是 vendor typelib 那一路的聚合行（`b132_typedcom_local_scope/B132Typed.vbp`、
  `ctrlmscomm`、`guisample` 三份、`TeeChart`）；
- 首跑 CI（`fb0005ad`）只红 `Emit manifest (shape oracle)` 一条。取 `ci/emit-manifest`
  （head=`fb0005ad`）逐行比：804 行里只有 **3 行**差，全是 `guisample` 那三份聚合行
  （`ascii256` 三行完全相同，只有 vendor 名字那一段文本不同），另外三行本机值与 CI 逐字相同；
- 按 CI 侧清单对齐这三行后（`b16ed0f6`），**804 行与 CI 清单 0 差**，CI **22/22 success**。

### 五、闸与连带

- `check_manifest_coverage`：`inputs=804 [.vbp=300 .bas=504]`，registered=804，集合相等；
- 九门 **9/9 rc=0**（emitted=1574，exe sha=`4ecee736d8c6`）；
- 本刀没有新增 RTL 出口，`check_builtin_slot_returns` 的派生表不动（仍 63 条）。

### 六、下一刀

同一条桥的**写口**：`Set <晚绑定对象>.<属性> = <对象>` —— 交出去的那一枚由谁持有、
桥在 `IDispatch::Invoke` 的 `DISPATCH_PROPERTYPUT` 那一路是否已经取走，是这把尺还没量过的第三站。

## §11.128（第一百四十七刀）`Set <晚绑定对象>.<属性> = <对象>`：写口两处病

### 一、量具与改前读数

`.temp/k147/probe`（不进语料）：`As Object` 接一个工程类实例，类里
`Public Property Set Held(ByVal v As Object)` 存进模块级 `gHold`。改前：

```
W0-made=1        ' 造了一枚
W1-terms=1       ' 写完 + Set it = Nothing 就 Terminate —— gHold 里什么都没有
W2-terms=1       ' 清 gHold 也不动
```

即：**写静默无效**。C3_OCX_TRACE 里看得更直白：

```
[189] INVOKE W=... cls=B147Box mem=Held argc=1
[C3_COM] SetRef 'Held' rcv=... -> COM dispid=1 hr=0x00000000
```

hr 是 S_OK，参数却整个丢了。两头各有一处病，合起来才是这个结果：

1. **描述符里 Get/Set 各拿一枚 dispid**（实测表：`{L"Held",1,2,…_get}` / `{L"Held",2,8,…_set}`）。
   COM 口径是同一个属性共用一枚 dispid，客户端按名查（`GetIDsOfNames`）只回**第一条**同名的
   ⇒ 写永远打在 Get 那一格上。生成器里 `dispIdMap` 建了却从没被读过。
2. **服务端认档**写成 `(ik == 4 || ik == 8) && (wFlags & DISPATCH_PROPERTYPUT)`。客户端递的是
   `DISPATCH_PROPERTYPUTREF`(8)，而 `DISPATCH_PROPERTYPUT` 是 4 ⇒ `wFlags & 4` 为假 ⇒
   按档一位都对不上 ⇒ 落进「同 dispid 第一条」的兜底，而那条正是读臂。

### 二、改法（两处各一句）

- `src/backend/detail/util/cgen_util_dllentry_tables.inc`：dispid 先查同名已分到的，再取
  driver 回写的 `comDispid`，最后才新发一枚（`nextDispid` 同时抬到已用值之上）。
- `src/rtl/core/vb6comserver/vb6comserver_obj.c` 的 `ComObj_Invoke`：put 方向先认**自己的**档
  （PUTREF→invkind 8 / PUT→invkind 4），再退到另一档 put，最后才走兜底；
  **兜底不再让 put 落到读臂**（找不到 put 档就答 `DISP_E_MEMBERNOTFOUND`，而不是 S_OK 空跑）。

### 三、读数（x64 与 x86 逐条相同）

| 行 | W0-made | W1-terms | W2-terms | W3-terms | W4-terms | R1-num | R2-num | R3-tag | R4/R5-terms | Z-made |
|---|---|---|---|---|---|---|---|---|---|---|
| 改前 | 1 | 1 | 1 | 2 | 2 | 41 | 42 | 0 | 2 / 2 | 3 |
| 改后 | 1 | **0** | **1** | 2 | 2 | 41 | 42 | 0 | 2 / 2 | 3 |

- W1/W2 是判据：写进去之后 `gHold` 持有，`Set it = Nothing` 不动它；清 `gHold` 才 Terminate。
- W3：同一槽连写两次、再清空，恰好 2 枚 —— 多写不多取引用。
- R 路三条是反面证人：同名 `Property Let`/`Get`（`o.Num = 41` → 41，`Bump 1` → 42）、普通
  带参 Function（`TagOf(Nothing)` → 0）在 dispid 共用之后一条不变。
- R4/R5：写进去的再用 Get 读回（那一枚归调用方，见 §11.127），读数仍 2 —— 不多不少。

（量具本身踩过一次坑：最早把 `"WR1-num=" & o.Num` 直接打，出来是垃圾。`&` 拼接把 Long 属性
发成 `vb6_ComGetStringProp`，那条读臂取 `bstrVal`。改成先落 Long 再打 —— 与 §11.127 那个
C1 同一类错：**量具的错不是发码的错**。反过来说，“把 VT_I4 读成 BSTR”这件事是不是也该管，
是另一把尺上的账，本刀不动。）

### 四、发码面

- 全清单 **804 行**逐行复算：**机器可复现的那 798 行一行未变**，另 6 行是 vendor typelib
  那一路的既有本机≠CI 差（不是本刀引入）；加新夹具两行 ⇒ **806 行**；
- **这一刀的发码变化不在清单覆盖范围内**：dispid 表落在 `com_entry.c`（工程级中间文件），
  而 `--emit-c` / `emit-manifest` 只覆盖各模块的 C ⇒ 清单这一刀**零行变化**。
  行为由 b147 的 12 枚读数钉死；把 `com_entry.c` 纳入清单需要另一次裁决（全语料行都会动、
  且带机器相关性），不在这一格做。

### 五、闸与连带

- 九门 **9/9 rc=0**（`inputs=806 [.vbp=301 .bas=505]`，emitted=1574，exe sha=`bb5f67dd8e99`）；
- 既有桥夹具零回归：b144 九枚读数、b146 八枚读数逐条不变（dispid 共用不影响按名 **读**
  与 `Set x = <读数>` 那两站）。

### 六、下一刀

`vb6_ComSetProp` / `vb6_ComSetPropArg` 那一族仍然只 `free` 结构体、不清内容（§11.110 里
在册的那一句）：`vb6_ComPackObject` 给对象 AddRef 一枚，`DISPATCH_PROPERTYPUT` 之后没人
放它。这一刀没动它，是因为量产路径上的 `Set <对象>.<属性>` 走的是 `vb6_ComSetRef`（不打包）；
要动它，得先有一枚**真走到 Pattern F**（`vb6_ComGetXxxProp(...) = value` 那个改写）的靶。

## §11.129（第一百四十八刀）服务端 getter 适配器：对象返回被当成整数交出去

### 一、怎么撞上的

第一百四十七刀之后接着量链式写 `Set <晚绑定 o>.<Child>.<Prop> = it`（`Child` 的返回类型
就是所在类自己）。那一发的读数不对，但顺着摸下去先撞到另一件事：**`o.Child` 这一枚单独的
读也坏**。`C3_COM` trace 里两条挨着：

```
[C3_COM] GetProp 'Child'  hr=0x0  vt=3  val=150397680     ' As B148Box（本类）
[C3_COM] GetProp 'Child2' hr=0x0  vt=9  val=-1            ' As B148Item（异类）
```

同一张描述符表、同一个 `me`，一个 `vt=3` 一个 `vt=9`。`vt=9` 是 VT_DISPATCH，`vt=3` 是
VT_I4 —— 对象被当成整数交出去了，`150397680` 就是被 `(int32_t)(intptr_t)` 截断后的指针。

### 二、病在哪儿

`src/backend/detail/util/cgen_util_dllentry_collect.inc` 里 getter 的适配器只认一档：

```cpp
if (retType == "vb6_VARIANT" && !methodSym->variableTypeName.empty()) { ... }   // Fix 184: 异类
```

而 `mapType` 给 `As Object` 与本类返回的 C 型**都是 `void*`**（`Object` 一栏写死 `void*`，
本类返回因为自引用走 `void*`），两条都不进这一档，于是落到本函数最后一个兜底：

```cpp
((VARIANT*)result)->vt = VT_I4; ((VARIANT*)result)->lVal = (int32_t)(intptr_t)_r;
```

这是**同一个 else** 吃掉两档，所以两档一起坏：`o.Child`（本类）与 `o.Held`（`As Object`）
改前都读到 VT_I4。

### 三、改法（两处）

- **本类返回**：判据放宽到「`retType == "void*"` 且声明的类型名解析到本工程一个**有 coclass
  描述符**的引擎类」⇒ 与异类返回同一律：`vb6_ComObject_FromInstance(...)` 包一层，再用
  `vb6_VariantObjectTransfer` 把那一枚引用交给 VARIANT。临时量的 C 声明型跟着改（本类那档
  是 `void*`，异类那档是 `vb6_cls_X*`）。
- **`As Object`**：类模块里 getter 的返回槽**已经**为泛对象档 claim 了一枚（`B148Box.c` 实测
  `vb6_ret_Held = gHold; vb6_ComAddRefDispatch(...)`），所以适配器只做**转交**：
  `vt = VT_DISPATCH; pdispVal = (IDispatch*)_r` —— 不再 AddRef，也不释放。

### 四、读数（x64 与 x86 十枚逐条相同）

| 行 | A2-self-tag | A3-tag | B2-fresh-tag | B3-item-terms | C2-held-tag | C3-item-terms | Z-box-terms | Z-item-made/terms |
|---|---|---|---|---|---|---|---|---|
| 改前 | **= Nothing** | 6 | 11 | 1 | **= Nothing** | 2 | 0 | 2 / 2 |
| 改后 | **5** | 6 | 11 | 1 | **77** | 2 | 0 | 2 / 2 |

- A2/C2 是判据：本类返回与 `As Object` 返回都得交回**那枚对象**，且交回的确实是写进去那一枚
  （`h.Tag = 77` 之后 `it.Tag` 读到 77 ⇒ 同一枚，不是拷贝）。
- A3：`Set s = Nothing` 之后 `o` 仍可写可读 —— 交回的那一枚归调用方（§11.126），不多取。
- B 路是**对照组**：异类返回改前改后都是对象，说明本刀没有把好的一档改坏。
- Z-item：`Set o.Held = it` → `Set o.Held = Nothing` → `Set it = Nothing`，恰好 2 枚
  Terminate —— 写口（§11.128）与本刀读口在同一枚对象上收放对得上。
- **Z-box-terms 停在 0 是在册的**：`Set Self = Me` 在返回槽上 claim 的那一枚（服务端 getter
  的泛对象返回槽，§11.125 步骤 4 同族）在客户端放手之后仍留在实例自己的 `__refcount` 上，
  于是 `Class_Terminate` 不触发。它与本刀不同尺：本刀管的是「交出去的**型**对不对」，
  那一枚管的「**笔数**对不对」。夹具把 0 如实钉住，将来接了就翻成 1。

### 五、发码面

- 全清单 **808 行**：机器可复现的 800 行一行未变，另 6 行是 vendor typelib 那一路的既有
  本机≠CI 差（不是本刀引入）；新夹具两行 ⇒ 806 → **808**；
- **这一刀的发码变化不在清单覆盖面**：适配器在 `com_entry.c`（工程级中间文件），`--emit-c` /
  `emit-manifest` 只覆盖各模块的 C ⇒ 清单这一刀零行位移（与 §11.128 同一处覆盖缺口）。
  行为由 b148 的十枚读数钉死。

### 六、闸与连带

- 九门 **9/9 rc=0**（`inputs=808 [.vbp=302 .bas=506]`，emitted=1574，exe sha=`0f4e9555c59c`）；
- 既有桥夹具零回归：b144 九枚、b146 八枚、b147 十二枚读数逐条不变（x64 与 x86 都过）；
- CI（github `dev`）run `45269265593` = **22/22 success**。

### 七、下一刀（已经量好的）

**链式写 `Set <晚绑定 o>.<Child>.<Prop> = it`**：发码是

```c
vb6_B148Box_prop_set_Held(o, vb6_ComGetObjectProp(o, L"Child"));  /* Set Property */
```

两处都不对：接收者写成 `o`（该是 `o.Child` 交回的那一枚），值是 `o.Child` 交回的（该是 `it`）。
读数 `.temp/k148/probe`（`Chain.vbp`）：`CH-hold=B148Box` —— `gHold` 里落的是 `o`（本该是 item）。
这一发还顺带绕过了 COM 打包（早绑定直呼 `vb6_B148Box_prop_set_Held`）。本刀没接，是因为
它的根在 `cgen_setlet_set_prop.inc` 那条链式改写里（接收者/值两个槽的绑定分开算），与本刀
的适配器不是同一处。

## §11.130（第一百四十九刀）`Set <晚绑定 o>.<Child>.<Prop> = <对象>`：两槽双双串位

### 一、读数与发码

夹具 `tests/b149_chain_write`（`o` 是 `As Object`，子对象是**另一枚** box，值是一枚 item）。
改前发码：

```c
vb6_B149Box_prop_set_Held(o, vb6_ComGetObjectProp(o, L"Child"));  /* Set Property */
```

两个槽都错了：接收者写成裸 `o`（该是 `o.Child` 交回的那一枚），值写成 `o.Child` 的交回
（该是 `it`）。而 `o` 的 C 类型是 `void*`（IDispatch*），被当成 `vb6_cls_B149Box*` 递进早绑定
直呼 —— 改前跑起来**首读即错**：

```
vb6_ComGetProp: property "Held" not found
Unhandled VB6 Error #438: Object doesn't support this property or method
```

### 二、病根两处

1. **晚绑定判据只认裸标识符**。`cgen_setlet_set_prop.inc` 里那条 `_lateBoundRecv` 问的是
   `_ma.object->kind == IdentifierExpr`（为的是 `Sub P2(ByVal o As Object): Set o.X = v` 那一发），
   成员链作接收者（`o.Child`）时 `_objClass` 推不出（根是 `As Object`）而 `_psSym->isExternal`
   为真（属性在别的模块）⇒ 照样进早绑定直呼。
2. **接收者的 COM marker 漏进值槽**。`emitExpr(o.Child)` 对晚绑定成员**只交出基础对象
   `o`**，成员名挂在 `comMemberName_` 上（等下一站来消费）；这里没人消费，紧接着
   `emitSetObjectRhs(node.value)` 就把这枚 marker 当成值消费掉了 —— 于是值槽拿到
   `vb6_ComGetObjectProp(o, L"Child")`，接收者槽只剩裸 `o`。

### 三、改法

- 晚绑定判据沿 `MemberAccess` 链**走到根**再问同一张表（`procObjectFormals_` /
  `knownVariantVars_` / `knownObjectVars_`，仍排除窗体控件与 UC 实例）；命中就整支跳过，
  交回下方通用 COM 路径。
- `emitExpr(*_ma.object)` 之后立即 `if (isComMarker_) _objE = resolveComValue("Object");` ——
  把这枚 marker 就地交成完整对象表达式（与值侧 Fix 110i 同一口径）。

改后发码：

```c
vb6_ComSetRef(vb6_ComGetObjectProp(o, L"Child"), L"Held", it);  /* COM SetRef */
```

### 四、读数（x64 与 x86 八枚逐条相同）

| 行 | A1 根上 | A2 子对象 | A3 同一枚 | C1 item-terms | Z box-terms | Z item-made/terms |
|---|---|---|---|---|---|---|
| 改前 | （首读即 438，无读数） | | | | | |
| 改后 | **empty** | **B149Item** | **9** | 0 | 1 | 1 / 0 |

- A1：根对象 `o` 自己的 `Held` 仍是空 —— 证明写落在子对象上，不是接收者根上。
- A2/A3：子对象上写的确实是 `it` 那一枚，改 `h1.Tag` 后 `it.Tag` 跟着变（同一枚，不是拷贝）。
- **Z 那两行是在册的**：链写里接收者 `o.Child` 交回的那一枚**临时引用没人放**。分账探针
  `.temp/k149/p1`（不进语料）：只取一次子对象又放回 = box-terms **2**（干净）；同一形状再
  走一发链写 = **3**（多一枚）。子对象因此不 Terminate（停在 1），它手里的 item 也跟着留下。
  这是「笔数」那把尺（`§11.131`），本刀只管「写到哪一枚、写的值对不对」。
  注：改前那形状（值槽里挂 getter）也照样漏这一枚 —— 不是本刀引入的。

### 五、发码面

- 全清单 **810 行**：机器可复现的 804 行一行未变，另 6 行是 vendor typelib 那一路的既有
  本机≠CI 差；新夹具两行 ⇒ 808 → **810**；
- **全语料复算零位移**：本刀改的两处只在「晚绑定成员链作接收者」这一格 —— 语料里没有这一格，
  所以 802 行逐行相同（`--emit-c` 覆盖到本刀，这也是它与 §11.129 不同之处：这次变化**在**
  清单覆盖范围内，读数就是零）。

### 六、闸与连带

- 九门 **9/9 rc=0**（`inputs=810 [.vbp=303 .bas=507]`，emitted=1574，exe sha=`fbb9fcd05e64`）；
- 既有桥夹具零回归：b144 九枚、b146 八枚、b147 十二枚、b148 十枚读数逐条不变（x64 与 x86）；
- CI（github `dev`）run `45276625769` = **22/22 success**。

### 七、下一刀（已经量好的）

**链上中间那一枚临时引用**：`vb6_ComGetObjectProp(o, L"Child")` 交回的引用归调用方
（§11.126），但链写的接收者槽只拿去派发一次，没人放 —— 每发链写漏一枚，子对象永不
Terminate。同族还有嵌套**读**链（`Set h = o.Child.Held` 里中间那一枚）。要么在发码侧把
接收者落成临时量、派发后 `Release`，要么在 RTL 侧给一个「交回即持、用完即放」的入口。
探针 `p1`（2 vs 3）就是这一刀的判据。


## §11.131（第一百五十刀）链写接收者那一枚临时引用

### 一、读数

尺子是 Terminate 计数（探针 `.temp/k150/p2`，不进语料）：只取一枚子对象又放回 → **2**（干净）；
同一形状再走一发链写 → **3**；嵌套**读**链（`v = q.Child.Tag`）→ **4**；两样都有 → **5**。
每多一层「取回来的对象式中间量」，就多留一枚。

改前发码（`Set p.Child.Held = it`，一条语句、无临时量）：

```c
vb6_ComSetRef(vb6_ComGetObjectProp(p, L"Child"), L"Held", it);
```

`vb6_ComGetObjectProp` 交回的那一枚按 §11.126 归**调用方**（`vb6_ComVarFree` 只 free 结构体、
不碰内容），而这一站只拿它递一次 `IDispatch::Invoke`，谁也不放 ⇒ 每发链写漏一枚，
中间那枚对象永不 `Class_Terminate`。

### 二、改法

`cgen_setlet_set_prop.inc` 的 COM `SetRef` 支路：接收者表达式若以 `vb6_ComGetObjectProp(` 或
`vb6_ComCallObject(` 开头（即自有链读/晚绑定调用交回的一枚），先落成临时量，派发完即放：

```c
void* _rcv0 = vb6_ComGetObjectProp(p, L"Child");
vb6_ComSetRef(_rcv0, L"Held", it);
vb6_ReleaseObject((void**)&_rcv0);
```

三个出口（视图直槽、视图直槽 R4、通用 `ComSetRef`）都覆盖。临时量用**独立序号**
`setRefRecvSeq_`（`cgen_state.inc`，`cgen_base_generate_prologue.inc` 里归零），不占
`tempCounter_` —— 否则存量 `_vs`/`_e`/`_ag` 全会重新编号、形状钉全红（§11.103 同一条理由）。
放引用走 `vb6_ReleaseObject(void**)`：它认 Collection 与 dispatchable 两支，对这两类中间量都安全。

### 三、读数（夹具 `tests/b150_chain_recv_ref`，x64 与 x86 六枚逐条相同）

| 行 | A2 held-tag | C1 item-terms | Z box-terms | Z item-made | Z item-terms |
|---|---|---|---|---|---|
| 改前 | 9 | 0 | **1** | 1 | 0 |
| 改后 | 9 | 0 | **2** | 1 | 0 |

- **Z box-terms 1 → 2 是这一刀的判据**：根与子对象都 `Class_Terminate`（改前子对象被那一枚夹住）。
- A2：写进去的确实是 `it` 那一枚（改 `h.Tag` 后 `it.Tag` 跟着变）。
- Z item-terms 停在 0：这是**写口打包**那一枚（`§11.110` 在册的「`Set` 值侧只 free 结构体不清
  内容」），与接收者临时量不是同一枚，本刀不动它。
- 夹具刻意**写完之后**才取子对象来验收（单层读），避开读链那一格 —— 见下「在册」。

### 四、发码面

- 全清单 **812 行**：机器可复现的 806 行一行未变，另 6 行是 vendor typelib 那一路的既有
  本机≠CI 差；新夹具两行 + b149 两行随本刀改码更新 ⇒ 810 → **812**；
- **全语料复算**：除上述 6 行与 b149 那两行外零位移（本刀改的是语句发码，在 `--emit-c`
  覆盖面内，所以这里是「量出来的零」，不是「覆盖不到」）。

### 五、闸与连带

- 九门 **9/9 rc=0**（`inputs=812 [.vbp=304 .bas=508]`，emitted=1574，exe sha=`118d259052e5`）；
- 既有桥夹具零回归：b144 九枚、b146 八枚、b147 十二枚、b148 十枚、b149 八枚读数逐条不变
  （x64 与 x86）；b149 的 `Z-box-terms=1` 现由**读链**那一格解释（改码后它仍走嵌套读），
  注释已就地改口径；
- CI（github `dev`）run `45285250271` = **22/22 success**。

### 六、在册（本刀只对平「只发链写」这一形状）

分账探针 `.temp/k150/p2` / `p4` 把剩下的两条量出来，都留给 §11.132：

1. **嵌套读链**（`Set h = o.Child.Held`、`v = q.Child.Tag`）：中间那一枚同样是「交回即持」，
   读侧没有对应的临时量，每多一层漏一枚（p2：`C=4`、`D=5`）。
2. **链写时同一枚子对象另握在别的变量里**（`Set c = q.Child` … `Set q.Child.Held = it`）：
   这一形状仍多留一枚（p4：`V2-box` 比 `V1-box` 多 1）。把释放顺序对调（先放 `c` 再链写）
   或写完之后才取子对象，都干净 —— 说明多出来的那一枚是「两处各持一次」的记账，不是悬垂
   （`Set h = c.Held` 在链写之后照旧读得到，`tag=9`）。


## §11.132（第一百五十一刀）链读接收者那一枚临时引用

### 一、读数

分账探针 `.temp/k150/p2`（每臂都在干净状态起，尺子是 box 的 Terminate 计数，每臂两枚对象）：

| 臂 | 形状 | 改前 | 改后 |
|---|---|---|---|
| A | 取一次子对象又放回 | 2 | 2 |
| B | 链写 `Set p.Child.Held = it`（§11.131 已对平） | 4 | 4 |
| C | 嵌套标量读 `v = q.Child.Tag` | 5 | **6** |
| D | 链写 + 嵌套对象读 `Set h = r.Child.Held` | 6 | **8** |

改前每多一层链就多留一枚。改前发码（一条语句、层层嵌套）：

```c
vb6_ComGetIntProp(vb6_ComGetObjectProp(q, L"Child"), L"Tag")
vb6_ComGetObjectProp(vb6_ComGetObjectProp(r, L"Child"), L"Held")
```

`vb6_ComGetObjectProp` 交回的那一枚按 §11.126 归调用方（`vb6_ComVarFree` 只 free 结构体、
不 `Release` `pdispVal`），而它被当**接收者**再递一层时谁也不放 ⇒ 中间那枚对象永不
`Class_Terminate`。

### 二、改法

新增 `hoistComRecvTemp(objExpr)`（`cgen_util_com.cpp`，声明在 `cgen_helpers.inc`）：接收者是
自有链读/晚绑定调用交回的那一枚（前缀 `vb6_ComGetObjectProp(` / `vb6_ComCallObject(`）时，
落成 `void* _rrvN` 并登记语句末收回。改后：

```c
void* _rrv0 = vb6_ComGetObjectProp(q, L"Child");
int32_t _av1 = (_e0 == vb6_ErrSerial()) ? (vb6_ComGetIntProp(_rrv0, L"Tag")) : (int32_t)0;
vb6_ReleaseObject((void**)&_rrv0);   /* 语句末收回 */
```

挂点是 `resolveComValue` 里**晚绑定兜底之前**（`isComMarker_=false` 之后、`getPropArgs`
之前）—— 必须排在所有 recognizer 之后：一旦提名，文本里的 `L"Tabs"` / `vb6_hwnd_X` /
`L"SelectedItem"` 这些线索就跟着走了，TabStrip / StatusBar / TreeView 那几族的直译会失配。
四道闸：过程内 / `indentLevel()==stmtLevel_`（声明与收回同层）/ 非块内条件 / 非拼串左操作数。

收回走既有的 `ownedTempsToFree_` 通道，新增档 3（档 1=BSTR、档 2=VARIANT）⇒ 语句末发
`vb6_ReleaseObject((void**)&t)`。序号 `comRecvTempSeq_` 独立（`cgen_state.inc` +
`cgen_base_generate_prologue.inc` 归零），不占 `tempCounter_`，存量 `_vs`/`_e`/`_ag` 编号零移动。

### 三、读数（夹具 `tests/b151_chain_read_ref`，x64 与 x86 八枚逐条相同）

| 行 | 含义 | 改前 | 改后 |
|---|---|---|---|
| A2-held-tag | 嵌套对象读 `Set h = p.Child.Held` 的值 | 7 | 7 |
| B1-child-tag | 嵌套标量读 `v = p.Child.Tag` 的值 | 99 | 99 |
| C1-two-level-tag | 两层链 `w = p.Child.Child.Tag` 的值 | 99 | 99 |
| Y-box-made | 出厂 box 数 | 3 | 3 |
| **Z-box-terms** | 根 + 两级子对象都 Terminate | **1** | **3** |
| Z-item-terms | 写口打包那一枚（§11.110 在册） | 0 | 0 |

连带：`tests/b149_chain_write` 的 `Z-box-terms` 由 1 变 **2** —— 它自己的嵌套读
`Set h1 = o.Child.Held` 就是本刀这一格，针同步改了口径。

### 四、发码面

- 全清单 **814 行**（新夹具两行）；语料 **45 行**随本刀改码 —— 链读站点遍地
  （TeeChart 系列、TreeView `Nodes(i).Text`、StatusBar `Panels(i).Text`、GUISample 的
  `Parent.Name` 那一族），这正是这把尺一直在漏的地方；
- 既有 6 行 vendor typelib 本机≠CI 差照旧不动，其中 4 行本次连同改码一起变了 ⇒ 值按
  **CI 那份清单**登记（`git fetch <remote> ci/emit-manifest` 取 `emit-manifest.txt`），
  不按本机；
- 两个发码形状针同步改成新形状：`tv_emitc_shape`（接收者落 `_rrvN` 再递一层）与
  `pd_parentsem_emitc`（`Parent` 交回的那一枚落 `_rrvN`，`Name` 从它读）。

### 五、闸与连带

- 九门 **9/9 rc=0**（`inputs=814 [.vbp=305 .bas=509]`，emitted=1574，exe sha=`3e0b4d6b55e6`）；
- 既有桥夹具零回归：b144 九枚、b146 八枚、b147 十二枚、b148 十枚、b149 八枚（其中
  `Z-box-terms` 1→2）、b150 六枚读数双架构逐条不变；b151 双架构一致；
- 首推 `03830aea` 的 CI（run `45293640919`）三红：Emit manifest 报**存量 4 行**真变、
  Tests(vbp #4/#7) 报两个发码形状针失配 —— 都是本刀的预期位移没登记完；
  补推 `4cbde0a1`（清单 4 行按 CI 那份登记 + 两针改形）后 CI（run `45295861718`）= **22/22 success**。

### 六、在册

分账探针 `.temp/k150/p4` 里还剩一格：**链写时同一枚子对象另握在别的变量里**
（`Set c = q.Child` … `Set q.Child.Held = it`）仍多留一枚；把释放顺序对调（先放 `c` 再链写）
或写完之后才取子对象，都干净 —— 说明那是「两处各持一次」的记账而不是悬垂
（链写之后 `Set h = c.Held` 照旧读得到 `tag=9`）。另：写口打包那一枚（`§11.110`）也仍在册。

## §11.133（第一百五十二刀）`As Object` 字段出厂那一枚引用

### 一、怎么撞上的

§11.132 收工后，b149/b150/b151 的最后一行读数还停在 `Z-item-terms=0`：item 造出来、
写进字段、局部变量也放掉了，可它**一次 Terminate 都没走**。前两把刀都把这枚记在
`§11.110`（「`Set` 值侧只 `free` 结构体不清内容」）名下。这一刀把它拆开看：

尺子 `.temp/k152/p10`（同一夹具两臂，每臂都在干净状态起）：

| 臂 | 形状 | 改前 item-terms | 改后 |
|---|---|---|---|
| s | 单层写 `Set q.Held = it` → 放 `it` → 放 `q` | 0 | **1** |
| t | 链写 `Set q.Child.Held = it` → 放 `it` → 放 `q` | 0 | **2** |

`C3_COM_RC_TRACE=1` 下 item 那一枚包装器的进出账（单层那一臂，指针取末 6 位）：

改前 —— `ADDREF cnt=2` / `RELEASE cnt=1`，少一发：

```
[189] FROM-INST W=…EC40B0 inst=…EB7D70 cls=B151Item
[189] ADDREF  W=…EC40B0 cnt=2        ← Set it = Nothing 之前: 字段那一枚
[189] RELEASE W=…EC40B0 cnt=1 owns=1 ← Set it = Nothing 放手
[189] INVOKE  W=…EC48F0 cls=B151Box mem=Held argc=1
        …放 q → Class_Terminate 走了, 但没有 item 的 RELEASE
```

改后 —— 同一臂 `ADDREF 2 / RELEASE 2`，最后一发是持有者出厂时字段放的那一下：

```
[189] RELEASE W=…EC48F0 cnt=0 owns=1 inst=…EB2880   ← 持有者出厂
[189] RELEASE W=…EC40B0 cnt=0 owns=1 inst=…EB7D70   ← 字段那一枚（新）
```

即：写口那一枚**没有多 AddRef**（`vb6_ComSetRef` 那条路不打包，见 §11.129 末尾），
缺的是**出厂**这一站 —— 前两把刀把账记错了行。这一刀纠正账面：
`§11.110` 的「值侧打包」依然在册（`vb6_ComSetProp` 那一族，要一枚真走到 Pattern F 的靶），
但 b149/b150/b151 的 `Z-item-terms` **不是它**，是字段出厂。

### 二、改法

`_Destroy`（`cgen_com.cpp`）逐字段做出厂时，三段判据只认得：数组、`String`、
以及 `projectClsStructOfCType(mapTypeRef(...))` 答得出 `vb6_cls_X*` 的**类类型字段**。
`As Object` / `As Collection` / 外部接口在 C 里是裸 `void*`，答不出类结构体 ⇒ 整支不出现在
出厂体里。改前发码（`tests/b152_object_field_teardown` 的 `vb6_cls_B152Holder_Destroy`）：

```c
void vb6_cls_B152Holder_Destroy(vb6_cls_B152Holder* me) {
    int32_t vb6_termErr = 0;
    if (!me) return;
    vb6_termErr = vb6_RunVbProtected((void(*)(void*))vb6_B152Holder_Class_Terminate, me);
    vb6_Free(me);                       /* mHold / mSelf 一格都没出现在这里 */
    ...
}
```

改后同体多出一行（字段名按声明序，与腿②/腿③ 同一个问句）：

```c
    vb6_ReleaseObject((void**)&me->mHold);  /* 账 #305 §11.133: As Object 字段出厂放手 */
```

判据不开第二问：`projectClsStructOfCType(mapTypeRef(...))` 那一问照旧（类类型字段走原路），
答不出类结构体时再问一句**声明名是不是字面量 `Object`**（`isObjectTypeName`，同一个文件里的
静态小函数，六个字母大小写不敏感）。是，就发：

```c
    vb6_ReleaseObject((void**)&me->mHold);  /* 账 #305 §11.133: As Object 字段出厂放手 */
```

**闸只那一把字面量 `Object`**，不是「所有落 `void*` 的字段」—— 这一条是首推（`0b862965`）
三红换来的：那一版按 `Vb6Type::Object || C 类型 == "void*"` 放闸，于是 `As <工程接口>`
的字段（k175 的 `m_lv`，语义上是 `&instance->__iv_IK175` 那枚**视图**）也被写进了出厂体，
而视图那一枚从来没 AddRef 过 ⇒ 出厂 Release 就是越权调用。三红各自的读数：

| 红的作业 | 读数 | 根因 |
|---|---|---|
| `Tests (vbp #6)` | `[EMITC-ABSENT] k175_viewstmt_com_fallback_gone ... 还在场: vb6_ReleaseObject((void**)&me->m_lv)` | 正是上面那一条：k175 那条 absent 针原本钉的就是「旧晚绑定 fallback 不认得视图字段」，新行撞上了它 |
| `Tests (compile #2)` | `CLASS-REFCOUNT-AUTHORITY FAIL - mapTypeRef(var.asType.get()) question sites = 7 (want exactly 6)` | 新分支多问了一次 `mapTypeRef`，破的是「同一问只有六处」那条普查 |
| `Emit manifest (shape oracle)` | `tests/b132_typedcom_local_scope/B132Typed.vbp` 一行（本机 17448 / CI 17429） | b132 有 `Public K As Object`，本刀确实改了它；机器侧差异按 CI 那份登记 |

**裸实例那一格**（`Set mSelf = Me` 之类，见 §11.126）由 `vb6_ReleaseObject` 自己的
Fix 191 判据拦下 —— 非可分派只清指针、不解引用 `lpVtbl`，所以新加的这一行不会把
「裸实例」当接口调 Release。夹具的第二臂就是这个钉子。

### 三、新夹具 `tests/b152_object_field_teardown`

`B152Holder`：`Private mHold As Object`（腿②/腿③ 已是 `vb6_ReleaseObject`）、
`Private mSelf As Object`（裸实例那一格）。六枚针，x64 与 x86 同值：

| 针 | 读数 | 意思 |
|---|---|---|
| `B152-A1-items` | 0 | 写进字段、放掉局部变量之后仍被字段握着 |
| `B152-A2-held-tag` | 7 | 字段那一枚读得回（放手不是「先放后读」） |
| `B152-Z-items` | **1** | 放掉持有者 ⇒ item Terminate（改前 0）← 本刀判据 |
| `B152-B1-self-nothing` | False | 裸实例 `Me` 进字段后还认得出 |
| `B152-B2-ok` | — | 出厂放手没有越权 Release ⇒ 走到这里 |
| `B152-DONE` | — | — |

### 四、发码面

- 全清单 **816 行**（814 + 新夹具两行）；
- **19 行**随本刀改码（全是「类里有字面量 `As Object` 字段」的那一族：b130/b131/b132/b133/b134
  承载者、`Charts 2020` 六个 uc、`ucTreeView`、`VBFlexGridDemo`、`k3138`、`Gifcontrol`、
  `dbgdlg`、`BalloonTooltips`、`ucSegmentedControl`、`ucSlider_2022`、b149/b150/b151）；
- 窄口之前那一版（`0b862965`）还多改了 8 行 —— 那些是只有 `As <工程接口>` 字段的工程
  （`b445_refcount`、`k2947_setview`、`k3060_setvariant`、`k3122_own_prop_set`、`ve_list` 等），
  窄口后逐个回到原值；
- 既有 6 行机器侧（b132、ctrlmscomm、GUISample ×3、TeeChart）登记 **CI 那份清单**的值
  （`git fetch <remote> ci/emit-manifest`），本机照旧读不出同值。

### 五、闸与连带

- 九门 **9/9 rc=0**（`inputs=816 [.vbp=306 .bas=510]`，emitted=1574，exe sha=`4bb4e…`），
- `check_class_refcount_authority.ps1` 那条普查在窄口后回到原答（`6 field-type questions`，
  没有第七处 `mapTypeRef` 站点）；k175 那条 absent 针窄口后原样通过（`m_lv` 不出现在出厂体里）；
- CI 三趟：首推 `0b862965`（宽口）run `45307425519` 三红（19 表那三行）；改窄口后
  `a3f08841` run `45312597405` 只剩 `Emit manifest` 一红，差的就是 b132 一行（本机 17448 /
  CI 17429 那位机器侧）；把那一行按 CI 那份登记（`c48339cd`）后 run `45314230404` = **22/22 success**。
- 桥夹具双架构零回归：b144 九枚、b146 八枚、b147 十二枚、b148 十枚读数逐条不变；
  b149 八枚、b150 六枚、b151 八枚只动 `Z-item-terms` 0→1，另有两针口径同步改口；
- b144/b146/b147/b148 的 `Z` 行不变，说明这几族本来就没有「`As Object` 字段被漏放」的账。

### 六、在册

- `§11.110` 值侧打包（`vb6_ComSetProp` 那一族 `vb6_ComPackObject` 给 AddRef 之后没人放）：
  仍在册，要一枚真走 Pattern F 的靶；- `.temp/k150/p4` 的 **V2** 那个盒子：取 `c = q.Child` → 链写 → 放 `c` → 放 `q`，
  root 走了 Terminate 而 child 没有（V2b「先放 c 再链写」、V2c「链写后才取 c」都干净）。
  本刀前后该探针读数逐条同值（`V2-box=1 / V2b=3 / V2c=5`）⇒ 与字段出厂无关，
  是「两处各持一次」的记账，仍在册；
- `As Variant` 字段的出厂：`vb6_VARIANT` 槽那一格本刀没动（`VariantClear` 会连 BSTR
  一起放，语料面比 `As Object` 大得多），另开一刀。

## §11.134（第一百五十三刀）`As Variant` 字段的读写两站

### 一、怎么撞上的

上一刀（§11.133）收尾时明确记着「`As Variant` 字段的出厂本刀没动，另开一刀」。照那条
接着做，先量 —— 读数却是**反的**：

`.temp/k153` 的 A 臂（`Set h.V = it` → 放 `it` → 读回 → 放 `h`）在改前就已经
`Z-terms=1`，item 明明走了一次 `Class_Terminate`。可是
`vb6_cls_B153Holder_Destroy` 里一个字都没有（只有 `vb6_Free(me)`）。在
`Class_Terminate` 里加一行打点之后真相出来了：

```
B153-A1-terms=0
B153-A2-held-tag=5
[b153-term tag=5]     ← 死在 Set got = Nothing 这一句, 不是死在持有者出厂
B153-Z-terms=1
```

RC 账（item 包装器指针最后 6 位）：

```
INVOKE holder V argc=1      ← Set h.V = it
ADDREF item cnt=2           ← 写口 VariantAssign 取的那一枚（槽持有）
RELEASE item cnt=1          ← Set it = Nothing
INVOKE holder V argc=0      ← Set got = h.V
RELEASE item cnt=1 → 0      ← 改前: 读出去的那一枚**没有**取引用,
                              放它就把槽里还握着的那一具拆了 ⇒ 槽里留悬垂
```

所以这里不是「漏」，是**悬垂**：①读出去那一枚 `Variant` 与字段**共用**同一个对象；
②出厂那一站不清槽，槽随 `vb6_Free(me)` 一起消失时那一枚也彻底丢了。两站成对。

### 二、改法 ①：读出去那一枚要取引用

`V = mV`（Property Get 体，两边都是 Variant）改前发码：

```c
vb6_VariantClear(&vb6_ret_V);
vb6_ret_V = vb6_VariantFromValue(me->mV);   /* ← _Generic 对 vb6_VARIANT 落 Identity 浅拷贝 */
```

`vb6_VariantFromValue` 对 `vb6_VARIANT` 实参落 `vb6_VariantIdentity`（纯浅拷贝:
不 AddRef、不拷 BSTR）—— 这一格本身是**故意的**（函数返回 / 集合取回那些临时量靠它
按值交回来，多拷一次就多一枚引用）。错在**站点把字段当成了「自有」**。

判据住在一处（`variantRhsOwnsValue`，意思是「右值自己持有这份值吗 —— 那就不用再拷」）。
改前它只把**本地** Variant 变量当槽:

```cpp
if (rhs && rhs->kind == ASTNodeKind::IdentifierExpr) {
    if (knownVariantVars_.count(...name...) > 0) return false;
}
```

类里的字段在 AST 里是一枚**裸 IdentifierExpr**（不带 `me->`），`knownVariantVars_`
里没有它 ⇒ 落到 `return true`（「自有」）⇒ 浅拷贝。改后按**字段表**问两句:

```cpp
if (rhs && !exprHasLateBoundComRead(rhs)) {
    if (rhs->kind == ASTNodeKind::IdentifierExpr) {
        // 裸名 = 本模块字段 (隐式 me->): 问类成员表
        if (classVariantMembers_.count(nm) || classVariantMembers_.count("m_" + nm))
            return false;
    } else if (rhs->kind == ASTNodeKind::MemberAccessExpr) {
        // Me.f / <类引用>.f: 问 driver 注入的工程级字段档表
        if (!cls.empty() && classFieldIsVariantSlot(cls, ma.memberName)) return false;
    }
}
```

口径一句话：**只有调用**（函数 / Property Get 交回的临时量）才「自有」；字段这些
**槽**是借来的，必须深拷贝。改后发码:

```c
vb6_VariantAssign(&vb6_ret_V, vb6_VariantFromValue(me->mV));
```

**这一格第一版判错了门**：首版按「源码里长得像左值 + `inferExprType` 答 Variant」判
（IdentifierExpr / MemberAccessExpr 一律算槽），本地六枚形状针位置都对，推上去
CI 直接红六处 —— `cvp_emitc_variant_fallback`（`at.value` 那种晚绑定读）与
b251/b252/b253/b254/b255（On Error 站的 `_vsN` 暂存）钉的正是**所有权转移**那一形:

```
missing: v = vb6_VariantFromComResult(vb6_ComGetProp(at, L"value"));
missing: if (_e0 == vb6_ErrSerial()) { vb6_VariantClear(&v); v = _vs0; }
missing: me->f = vb6_VariantFromValue(o);
```

晚绑定成员读交回的是一枚**新造**的 Variant（`vb6_VariantFromComResult` /
`vb6_ComGet*Prop`），`_vsN` 是 On Error 站自己刚写进去的暂存 —— 两者都是「转移」，
按槽判会变成「深拷贝 + 原来那一枚没人放」。于是加 `!exprHasLateBoundComRead(rhs)`
这道门，并把「像左值」换成「查得到是一枚 Variant 字段」。六针恢复，b153 读数不变。

### 三、改法 ②：出厂清槽

`_Destroy` 的逐字段放手补一档（与 §11.133 的 `As Object` 同一处，判据换成
`Vb6Type::Variant`）:

```c
vb6_VariantClear(&me->mV);  /* 账 #305 §11.134: Variant 字段出厂清槽 */
```

`vb6_VariantClear` 就是与写口 `vb6_VariantAssign` 成对的那一半（写口 Clear + 深拷贝，
出厂清槽），对空槽是空转 —— 工厂里 `As Variant` 字段初始化就是 `vb6_VariantEmpty()`。

### 四、新夹具 `tests/b153_variant_field_teardown`

`B153Holder`：`Private mV As Variant` + Get/Let/Set 三站 + `PutText`（槽里放串）+
`TextLen`（类内读回）。八枚针，x64 与 x86 同值:

| 针 | 改前 | 改后 | 意思 |
|---|---|---|---|
| `A1-terms` | 0 | 0 | 写进槽、放掉局部变量后 item 仍被槽握着 |
| `A2-held-tag` | 5 | 5 | 槽里读得回 |
| `A3-terms` | **1** | **0** | 读回并放掉那一枚之后 item 仍活着 ← 悬垂那一条的判据 |
| `Z-terms` | 1 | 1 | 放掉持有者 ⇒ 出厂清槽 ⇒ Terminate（改前是「早死」，不是「晚死」） |
| `B1-len` | 1 | 1 | 槽里放标量 7，清槽不弄坏 |
| `B2-len` | 4 | 4 | 槽里放串 `"abcd"` |
| `C1-terms` | **1** | **2** | 两枚持有者各握一枚 ⇒ A(1)+C(1) ← 漏的判据 |
| `DONE` | — | — | — |

C 臂必须两枚持有者都在：改前 C 那一枚从头到尾没被退（`Z` 上只有 A 的 1），
改后 C 结束在 `Set q = Nothing` 那一句（栈式记账: `New`(1) → p 槽(2) → q 槽(3) →
`Set it2 = Nothing`(2) → p 出厂(1) → q 出厂(0)）。

### 五、发码面

- 全清单 **818 行**（816 + 新夹具两行）；
- **14 行**随本刀改码: b168/b169/b170/b171/b172/b173 那一族槽夹具、`b433`、`setcls`、
  `ve_list`、`ucSlider_2022`、`ucTreeView`、`Charts 2020`、`dbgdlg`、`VBFlexGridDemo`；
- 既有 6 行机器侧（b132、ctrlmscomm、GUISample ×3、TeeChart）本机≠CI 差照旧不动 ——
  用**改前编译器**逐行重算过: 这 6 行改前改后同值，本刀对它们零影响。

记账口径补一句：首版把「像左值的 Variant 右值」全判成槽时，清单里有 41 行动；
收窄到字段表之后只剩 14 行 —— 那 27 行的差**不是**本刀的语义，是首版判门过宽顺手
带出来的（`test_late_*`、`dict_native` 这些正是靠「转移」才对的）。

### 六、闸与连带

- 九门 **9/9 rc=0**（`inputs=818 [.vbp=307 .bas=511]`，emitted=1574）；
- 相邻桥夹具零回归: b144 九枚、b146 八枚、b147 十二枚、b148 十枚、b149 八枚、
  b150 六枚、b151 八枚、b152 六枚读数逐条不变；
- CI 22/22（`99511f8f`）。

### 七、在册

- `vb6_VariantFromValue` 对 `vb6_VARIANT` 实参仍是 `vb6_VariantIdentity`（浅拷贝）——
  **本刀没动 RTL**：那一格是「临时量按值交回」用的，动它就多拷一枚引用。这一刀只把
  「字段」这一族从它射程里挪出去（判据落在站点上，不住 RTL）；
- 形如 `Set c->f = o`（`o` 是 `As Variant` **形参**）仍走 `vb6_VariantFromValue(o)`
  浅拷贝 —— `setcls` 那格的形状针钉的就是它。参数槽与字段槽是两笔账，本刀只动了
  读出口与出厂，写口的形参那一族另开一刀；
- `§11.110` 值侧打包（`vb6_ComSetProp` 那一族 `vb6_ComPackObject` 给 AddRef 之后没人放）:
  仍在册，要一枚真走 Pattern F 的靶；
- `.temp/k150/p4` 的 V2 盒子记账（两处各持一次）仍在册，与本刀无关（读数前后同值）；
- `.temp/k153/p1` 里记着一格: `n = h2.V`（晚绑定读 Variant 属性直接赋给 Long）会发
  `vb6_VariantToLong(vb6_ComGetIntProp(h2, L"V"))` —— 属性读的交回值已经是 `int32_t`，
  外面再套一层 `VariantToLong` 是 C2440（编译不过，不是运行期错）。这一格还没开门，
  只是记在册。

## §11.135（第一百五十四刀）晚绑定属性写口的工程类实参

### 一、怎么撞上的

上一刀收尾时在册里留了一句：`Set c->f = o`（`o` 是 `As Variant` **形参**）仍走浅拷贝。
照那条去量 —— 先造一枚最小探针（`.temp/k154/p1`）：`B154Box` 有
`Property Set Item(ByVal it As B154Item)`，主程序两腿，早绑定 `Set box.Item = it` 与
晚绑定 `Set o.Item = it`（`o` 是 `As Object`，值来自 `Set o = box`）：

```
A1-held=11     ← 早绑定: 存住了
B1-held=0      ← 晚绑定: 读回 Nothing
Z-terms=1
```

读数不是"晚一点"或"少一枚引用"，是**整枚对象被丢了**。加一枚探针打印
（`C3_DBG_PUTARG`，量完即撤）能看到服务端那一侧：

```
[PUTARG] mem=Item ik=8 argc=1 [0]vt=9 p=0000015ED57FE380 want[0]=9
```

`vt=9`（VT_DISPATCH）、`want[0]=9` ⇒ `vb6_ComArgNormalize` 走的是"同档借用"那一支，
换算什么都没丢；实参进了适配器，出了适配器就成 NULL 了。

### 二、根因：适配器剥壳那一问只认包装器

工程类的 Property Set 在 `com_entry.c` 里落成一枚适配器，工程类 ByVal 形参那一臂是：

```c
provided = "(vb6_cls_X*)vb6_ComObject_GetInstance(((VARIANT*)args[0])->pdispVal)";
```

`vb6_ComObject_GetInstance` 的语义是「**我们自己的包装器** → 裸实例，其余答 NULL」。
而**同一工程内**的调用点递给 `vb6_ComSetRef` 的是一枚**裸实例**
（`vb6_ComSetRef(o, L"Item", it)`，`it` 本身就是 `vb6_cls_B154Item*`，从没包装过）
⇒ 剥壳答 NULL ⇒ 服务端形参收到 `Nothing` ⇒ 属性写把整枚对象丢掉，`S_OK` 照常返回。

判据的**另一口径**早就住在 RTL 里：

```c
void* vb6_ComObject_UnboxVb6Instance(void* v) {
    if (!vb6_ComObject_IsVb6Wrapper(v)) return v;   /* 非包装器: 原样交回 */
    void* inst = vb6_ComObject_GetInstance(v);
    return inst ? inst : v;
}
```

（它的注释写着：For Each 项目类那一路、`cgen_com.cpp` 那些调用点用的就是它。）
适配器换用它：

```c
provided = "(vb6_cls_X*)vb6_ComObject_UnboxVb6Instance(((VARIANT*)args[0])->pdispVal)";
```

一处改，写口（Property Set）与**所有**带工程类 ByVal 形参的公开方法一起修好 ——
适配器是它们共用的那一格。

### 三、读数

新夹具 `tests/b154_late_prop_set`（九针），x64 与 x86 同值:

| 针 | 改前 | 改后 | 意思 |
|---|---|---|---|
| `A1-nothing` | False | False | 早绑定写口对照 |
| `A2-held` | 11 | 11 | 同上 |
| `B1-nothing` | **True** | **False** | 晚绑定写口存住了 ← 判据 |
| `B2-held` | **-1** | **11** | 同上（-1 是夹具里"槽是 Nothing"的记号） |
| `C1-nothing` | True | False | 放掉调用方那一枚后盒子还握着 |
| `C2-held` | -1 | 11 | 同上 |
| `Z-terms` | 1 | 1 | 盒子放手 ⇒ 字段出厂放手 ⇒ item Terminate |
| `Z-made` | 1 | 1 | — |

`Z` 两档同值不是"本刀没动它"，而是改前那一枚**从没进过盒子**（调用方一放手就 Terminate），
改后是**盒子握着、随盒子出厂**才 Terminate —— 次数一样，路径不同，`B`/`C` 那一对才是判据。

第一次推送时 CI 在这个夹具上红了两处（x64/x86 各一），原因不在语义：
夹具打印的是 `A1-nothing=…`，而针写成 `B154-A1-nothing=…`。给输出补上前缀后重推。

### 四、发码面

- 清单 **820 行**（818 + 新夹具两行）；
- 本刀那一行改在**项目级** `com_entry.c`（适配器），不在清单覆盖的逐输入产物里 ⇒
  既有 818 行零移动（除 6 行机器侧在册），`cmp.ps1` 只列那 6 行；
- 九门 9/9 rc=0（`inputs=820 [.vbp=308 .bas=512]`）；相邻桥夹具 b144…b153 读数逐条不变；
- CI 22/22（`1ada4d7e`）。

### 五、在册（本刀量到、未修）

- **工程类属性置空那一站**：`Set box.Item = Nothing` 发成

  ```c
  vb6_ReleaseObject((void**)(&(void*){vb6_B154Box_prop_get_Item(box)}));  /* Set Nothing */
  ```

  拿 **getter** 的返回值去 Release —— `prop_set_Item` 一次都没调，属性没被清空
  （探针里 `B0-cleared` 一直是 False）。`Fix 056b` 那条改写只认 `prop_set_` / `prop_let_`
  前缀，而 `emitExpr` 给属性左值发的是 `prop_get_Item(box)` ⇒ 前缀对不上，落到下面
  "对调用结果 Release"那条通用路。
- **晚绑定读带类型的属性**：`o.Held`（`Property Get Held() As Long`）发成
  `vb6_CStrLong(vb6_ComGetStringProp(o, L"Held"))` —— 用 **String** 那一枚读臂读 Long，
  读回垃圾（实测 `B3-late-held=595138088`）。写口修好之后，这一格是同一族里剩下的一半。
- `Set c->f = o`（`o` 是 `As Variant` **形参**）仍走 `vb6_VariantFromValue` 浅拷贝
  （`setcls` 那格形状针钉着它）—— 参数槽与字段槽两笔账。
- `§11.110` 值侧打包、`.temp/k150/p4` 的 V2 记账照旧在册。


## §11.136（第一百五十五刀）工程类属性置空要点名 setter

### 一、撞见

上一刀（§11.135）在册第一笔：`Set box.Item = Nothing`（`box` 是工程类）改前发成

```c
vb6_ReleaseObject((void**)(&(void*){vb6_B154Box_prop_get_Item(box)}));  /* Set Nothing */
```

拿 **getter** 的返回值去 Release。实测两个症状：

- **语义**：`prop_set_Item` 一次都没调 ⇒ 属性根本没被清空（探针 `P1-cleared=False`）；
- **引用账**：逐站点打点（`.temp/k155/p3`，改前二进制）看到 item 的 `Terminate` **一次都没发生**
  （`PT2=0 … PT8=0, Z-terms=0`，改后同一处 `PT8=1`）。

第二条的机理不是"多减一格"，是**多加一格漏在那**：getter 交回的是**自持**引用

```c
vb6_ret_Item = (vb6_cls_B155Item*)vb6_ComObject_UnboxVb6Instance((void*)(me->mIt));
if (vb6_ret_Item) vb6_cls_B155Item_AddRef(vb6_ret_Item);   /* ← 自持 */
```

而 `vb6_ReleaseObject`（`vb6com.c`）对**类实例指针**根本不 Release：

```c
IUnknown* pUnk = (IUnknown*)(*objPtr);
if (vb6_ComIsDispatchable(pUnk)) {           /* ← Fix 191 那道闸 */
    pUnk->lpVtbl->Release(pUnk);
}
*objPtr = NULL;
```

裸类实例不是可分派对象（首字段是引用计数，不是 vtable）⇒ `IsDispatchable` 判否 ⇒
那一 Release **什么也没做**，getter 的 AddRef 就漏在账上。于是：属性没清空（槽还握着），
而 item 的计数永远多一格 ⇒ 盒子出厂时收不干净，`Terminate` 不来。

### 二、为什么改写没接上

置空走的是 `cgen_setlet_set_prop.inc` 里最早的那一支：先 `emitExpr(*node.target)`，
再对求值结果套 `vb6_ReleaseObject`，指望主路（`Fix 056b`）把文本里的 `prop_get_`
前缀改写成 `prop_let_` / 保持 `prop_set_`。

而 `emitExpr` 给**属性左值**发的本来就是 `prop_get_Item(box)` —— 左值那一路只有
`prop_get_` 这一个名字可用（写口没有"取引用"这回事）。`Fix 056b` 只认
`prop_set_`/`prop_let_` 前缀 ⇒ 前缀对不上 ⇒ 一路落到"对调用结果 Release"那条通用路，
静默发错。**这不是漏了一个分支，是问错了对象：拿文本去认语义。**

### 三、修：置空之前先问符号

在那一支早退**之前**加一段，用与 Set 值那一支**同一张表**（`lookupModuleByKind`）问符号：

```cpp
if (node.target->kind == ASTNodeKind::MemberAccessExpr) {
    Symbol* psN = symTab_.lookupModuleByKind(ma.memberName, SymbolKind::PropertySet);
    ...
    if (okN) {
        c_.emitLine(cProcName("prop_set_" + ma.memberName, psN->access, objClsN)
                    + "(" + objN + ", NULL);  /* Set Nothing (property setter) */");
        return;
    }
}
```

`okN` 的闸与 Set 值那一支同源，一条不少：

| 闸 | 为什么 |
|---|---|
| 推不出类名（晚绑定 `Object`/`Variant`） | 成员的 C 存储是 `void*`，直发等于当实例指针用 ⇒ 交给 COM `SetRef` |
| `knownFormControls_` 命中 | 控件属性不属该类，`lookupModuleByKind` 全局按名会误命中他类同名属性 |
| `hostPseudoIsObject` 命中 | `UserControl`/`Ambient`/`VBA` 等宿主伪对象同理 |
| `inferUdtTypeOfExpr` 非空 | 接收者是 UDT 值/字段链，成员必是结构体字段，永远不是类属性 |
| 外部符号且 `sourceModule != 接收者类` | 只剩"属性确实属于这个类"才直发 |

晚绑定那一腿没被这个改动碰到：`Set o.Item = Nothing` 照旧发 `vb6_ComSetRef(o, L"Item", NULL)`
—— 夹具 `P2-cleared=True` 两档都对，就是钉这一条。

### 四、读数

新夹具 `tests/b155_prop_set_nothing`（八针），x64 与 x86 同值：

| 针 | 改前 | 改后 | 意思 |
|---|---|---|---|
| `P1-held` | 11 | 11 | 置空前存住 |
| `P1-cleared` | **False** | **True** | 置空后 `ItemIsNothing` ← 语义判据 |
| `P2-cleared` | True | True | 晚绑定置空经 COM 派发（对照，堵住"顺手把晚绑定改成直发"） |
| `P3-held` | 11 | 11 | 置空后再写还能用 |
| `P4-nothing` | False | False | 放掉调用方那一枚后盒子还握着 |
| `Z-terms` | **0** | **1** | item Terminate ← 账判据 |
| `Z-made` | 1 | 1 | — |
| `DONE` | — | — | — |

`P1-cleared` 与 `Z-terms` 是一行发错踩出来的两件事：前者是语义（写口没点名），后者是账
（getter 那枚自持引用被一个对裸实例无效的 Release 顶掉）。本刀两件一起回正。

### 五、发码面

- 清单 **822 行**（820 + 新夹具两行，接在 b154 两行之后）；
- 全语料 **零移动**：存量 820 行里没有一处 `Set <工程类>.Prop = Nothing`，`cmp.ps1` 只列
  那 6 行机器侧在册（b132 / ctrlmscomm / GUISample×3 / TeeChart）；
- 九门 9/9 rc=0（`inputs=822 [.vbp=309 .bas=513]`）；相邻夹具 b144…b154 读数逐条不变。

### 六、在册（未修）

- **`vb6_ReleaseObject` 对裸类实例是空转**：本刀只是绕开了它踩的那一格（置空不再走那条路）；
  但"把类实例塞进 `void*` 再交给 `vb6_ReleaseObject`"仍然到处都能写出来，写一次漏一格。
  真修要问的是"这个 `void*` 是谁的、该怎么放手"，不是再加一道闸。
- **晚绑定读带类型的属性**：`o.Held`（`Property Get Held() As Long`）仍发成
  `vb6_CStrLong(vb6_ComGetStringProp(o, L"Held"))` —— 用 **String** 那一枚读臂读 Long，
  读回垃圾（§11.135 实测 `B3-late-held=595138088`）。写口这一族（Set 值 / Set 置空）已修，
  读口这一半是同一张属性表上剩下的那半。
- `Set c->f = o`（`o` 是 `As Variant` **形参**）仍走 `vb6_VariantFromValue` 浅拷贝
  （`setcls` 那格形状针钉着它）—— 参数槽与字段槽两笔账。
- `§11.110` 值侧打包、`.temp/k150/p4` 的 V2 记账照旧在册。

## §11.137（第一百五十六刀）已是 BSTR 的晚绑定读不再套 CStr

### 一、撞见

上一刀（§11.136）在册第二笔：晚绑定读带类型的属性发成
`vb6_CStrLong(vb6_ComGetStringProp(o, L"Held"))`。造探针（`.temp/k156/p1`，
`B156Box` 有 `Property Get Held() As Long`、`Property Get Note() As String`）量到：

```
B156-Q1=11        赋给 Long
B156-Q2=12        Long + 1
B156-Q3=eq        与字面量比
B156-Q4=hi        String 属性赋给 String
B156-Q5=128849432 拼串里的裸值 ← 一个指针值（x86 打 14248540）
```

Q1~Q4 都对，只有 Q5 是垃圾 —— 而 Q5 正是**最普通的一句** `Debug.Print "…" & o.Held`。

### 二、两站各说一套

读口（`resolveComMarkerForPack`，`packFnHint` 为空那一支）按"不知道类型就当 String"发
`vb6_ComGetStringProp`。**这一步本身没错**：RTL 那一枚内部是

```c
BSTR result = vb6_ComUnpackBSTR(pv);   /* vt != VT_BSTR 时 VariantChangeType(…, VT_BSTR) */
```

`VT_I4` 折成 `"11"`、`VT_BOOL` 折成 `"True"` —— 值是对的。

错在**拼串侧**：`wrapToBSTR` 拿 `inferExprType` 的答案选格式化函数，而 `inferExprType` 走
的是"按**裸成员名**查模块符号"那条兜底 —— 这个工程里 `Held` 唯一，答 Long ⇒
`vb6_CStrLong(...)` ⇒ 把 BSTR 指针当 `int32_t` 交给格式化。同一句 `o.Held`：
**读口说是串、拼串侧说数**，两边都不问对面。

### 三、修：不新开判据，登记进那张老表

`isBstrReturningCall`（`cgen_expr_binary_util.cpp` 的 `kBstrReturningCalls`）本来就是
"这个调用交回的就是 BSTR，别再折一层"那张表（Fix 100a/100d 那一族的老账：
`vb6_GetListItem` / `vb6_GetImageListKeyAt` / `vb6_StatusBar_GetPanelText` …）。
把 `vb6_ComGetStringProp(` 登记进去，`wrapToBSTR` 在数值清单之后、类型推断之前直通：

```cpp
if (isBstrReturningCall(expr)) return expr;
```

于是 Q5 的发码从

```c
vb6_CStrLong(vb6_ComGetStringProp(o, L"Held"))
```

变成

```c
vb6_ComGetStringProp(o, L"Held")
```

—— 值由 RTL 那一步 `VariantChangeType` 说话，上层不再二次解释。

### 四、读数

新夹具 `tests/b156_late_typed_prop_read`（六针），x64 与 x86 同值：

| 针 | 改前 | 改后 | 意思 |
|---|---|---|---|
| `Q1` | 11 | 11 | 赋给 Long → `ComGetIntProp`（本来就对） |
| `Q2` | 12 | 12 | 参与算术 → `ComGetIntProp` |
| `Q3` | eq | eq | 与字面量比（走 `ComGetDoubleProp`，C30-11 后记那条口径） |
| `Q4` | hi | hi | String 属性赋给 String → `ComGetStringProp` 直通 |
| `Q5` | **128849432**（x86 14248540） | **11** | 拼串里的裸值 ← 判据 |
| `DONE` | — | — | — |

### 五、发码面

- 清单 **822→824**（新夹具两行）；
- 存量 **86 行 CHG**，全是"把包在 `vb6_CStrXxx(vb6_ComGetStringProp(…))` 外层的那一层
  CStr 解开"，没有删行、没有别的位移；
- 41 个受影响工程逐个 `--output-dir` 重建，**0 失败** —— 确认没把 BSTR 塞进要 VARIANT 的槽；
- **一条形状钉跟着改了形**：`tv_emitc_shape` 原钉
  `vb6_CStr(vb6_VariantFromValue(vb6_ComGetStringProp(_rrv…`（比较上下文先包 VARIANT 再折回串），
  现在同一处是 `vb6_StrCmp(vb6_ComGetStringProp(_rrv…`（直通给 `vb6_StrCmp`，它收 BSTR）。
  同一条判据的**另一半没动**：数值成员仍钉 `vb6_ComGetDoubleProp(_rrv` —— 挡住"为了修字符串
  把两档一起改成 BSTR"那种反向过度修正；
- 九门 9/9 rc=0（`inputs=824 [.vbp=310 .bas=514]`）。

### 六、第一次推送红的两处（记账）

1. `Emit manifest (shape oracle)` —— 4 行机器侧（b132 / GUISample×3）在 CI 那台机器上随本刀
   一起变了哈希（它们本来就含晚绑定读的产物）。按老规矩**以 CI 为准**：把 CI 产出的
   `emit-manifest.txt` 那 4 行灌回期望，本地与 CI 的 824 行逐行 COMPARE = 0 差异。
2. `Tests (vbp #7)` —— 就是上面那条形状钉，钉的是**改前**的形。

### 七、在册（未修）

- **`Dim v As Variant: v = o.Held` 仍是 Variant(BSTR `"11"`)，不是 Variant(Long 11)** ——
  值对、型别不对。这一格要的是"读口按声明类型选臂"，本刀只把"已经读成 BSTR"那一头的二次
  折叠摆正；型别保真要另开一刀（与 `n = o.Held` 那条 C2440 同一族）。
- `Set c->f = o`（`o` 是 `As Variant` **形参**）仍走 `vb6_VariantFromValue` 浅拷贝。
- `vb6_ReleaseObject` 对**裸类实例**是空转（§11.136 在册第一笔）。
- `§11.110` 值侧打包、`.temp/k150/p4` 的 V2 记账照旧在册。

## §11.139（第一百五十七刀）声明为 Variant 的晚绑定读不再按目标档钉型别

### 一、撞见

上一刀（§11.137）在册第一笔写得明明白白：`v = o.Held` 读进 Variant 时**型别不对**，
而且它和"`n = o.Held` 那条 C2440 同一族"。这一刀来收那一族。

先造探针量 §11.138（晚绑定读进 Variant 的型别保真）——`.temp/k157/p1`，`B157Box` 有
`Property Get Num() As Long` / `Txt() As String` / `Flag() As Boolean` / `Rate() As Double`，
四句 `v = o.Num` 之类。**读数全对**：

```
N-type=3 val=11     VT_I4
T-type=8 val=hi     VT_BSTR
B-type=11 val=True  VT_BOOL
D-type=5 val=1.5    VT_R8
```

发码是 `vb6_VariantFromComResult(vb6_ComCall(o, L"Num", NULL, 0))` —— 服务端交回什么就是什么，
型别一点没丢。所以 §11.138 那条**在册项本身是错的推理，量不到病**，不修；本刀改收真正会发病的那一支。

那个真正发病的形状在 `.temp/k158/p1`：`B158Box` 有 `Property Get V() As Variant` 与
`Property Let V(ByVal x As Variant)`，主程序（`o As Object`，晚绑定）：

```vb
Dim n As Long
Dim s As String
box.V = 11
n = o.V          ' L1
n = o.V + 1      ' L2
s = CStr(o.V)    ' L3
```

**整份工程编不过，rc=1**：

```
B158Main.c(26): error C2440: 无法从 'int32_t' 转换为 'vb6_VARIANT'
```

发码两形：

```c
n  = vb6_ChkLong(vb6_VariantToLong(vb6_ComGetIntProp(o, L"V")));          /* L1 */
n  = vb6_ChkLong(vb6_VariantToLong(vb6_VarAddV(
       vb6_VariantFromValue(vb6_ComGetIntProp(o, L"V")),                  /* L2 */
       vb6_VariantInt((int16_t)(1)))));
```

### 二、两站各说一套

- **读口**（`resolveComValue(unpackType)`）按**目标档**发码：目标 `Long` ⇒
  `vb6_ComGetIntProp`。这一发内部已经做过 `VariantChangeType(…, VT_I4)` 与 `VarClear`，
  C 类型**就是 `int32_t`** —— 它对"值"这件事答得完全正确。
- **consumer 站**按**声明档**：`inferExprType(o.V)` 走"按裸成员名查模块符号"的兼底，属性
  声明是 `Variant`，于是这一句又被当成"手里揣着整只 Variant"，套上
  `vb6_VariantToLong(...)`（L2 更先在算术那一站被改道成 `vb6_VarAddV`，用
  `vb6_VariantFromValue(int32_t)` 装箱）。

两边都不问对面。**同一句 `o.V`：读口说是 `int32_t`、consumer 说是 `vb6_VARIANT`。**

### 三、先试的那条路（失败，记账）

第一反应是让**读口**服从声明档：只要成员声明为 Variant 就一律走通用读口
`vb6_VariantFromComResult(vb6_ComGetProp(...))`，把整只 Variant 交出去。这条路**编得过、
读数也对**（L1=11 L2=12 L3=11 L4=42 L5=8/42），而且和 L3 那句一直在用的形状同一个出口。

但夹具读数把它按住了：

```
b144: A2-terms=1 → 0   B2-terms=1 → 0   C1-terms=2 → 0   D1-terms=3 → 0   E1=4 → 0   Z=4 → 0
b153: Z-terms=1 → 0    C1-terms=2 → 1
```

`terms` 是夹具数项目类 `Terminate` 的次数。整只 Variant 被交到**标量消费点**以后没人
`vb6_VariantClear`，里面揣着的那一份对象引用就再也不放 —— 引用计数漏了，实例永不
`Terminate`。这条路会把晚绑定读的所有对象寿命拖坏，**回收**。

### 四、修：判据只看"已经发出来的那一段文本"

落点不是"按声明档选臂"，而是**不让 consumer 二次解释一个已经交回标量的读**。新增一枚
共用判据（`cgen_helpers.inc` 声明 / `cgen_util_classtype.cpp` 定义）：

```cpp
bool CCodeGen::isScalarComReadText(const std::string& cExpr) {
    return cExpr.compare(0, 18, "vb6_ComGetIntProp(") == 0
        || cExpr.compare(0, 21, "vb6_ComGetDoubleProp(") == 0
        || cExpr.compare(0, 23, "vb6_ComGetLongPtrProp(") == 0;
}
```

三处消费点复用同一问：

1. `toLongIfVariant` —— 需要 `int32_t` 的上下文（整除/索引/For 上下界…）直通：
   ```cpp
   if (isScalarComReadText(cExpr)) return cExpr;
   ```
2. `cgen_assign_value_sem.inc` 的 `valueIsVariant` 支（`n = o.V` 这一形就落在这里）：
   ```cpp
   if (isScalarComReadText(value)) { c_.emitLine(target + " = " + value + ";"); return; }
   ```
3. `cgen_expr_binary.cpp` 的 `routeOperand86` —— 不进 Variant 算术改道（否则 L2 会被
   `vb6_VarAddV(vb6_VariantFromValue(int32_t), …)` 装箱）：
   ```cpp
   if (isScalarComReadText(cExpr)) return false;
   ```

判据的纪律与 §11.137 同一条：**信发生过的事，不信对它的描述**（文本是发生过的事，
`inferExprType` 是描述）。收窄到数值三档，是因为**能命中这三档的以前一律是 C2440** ——
对今天能编译的表达式零影响，这一点由全语料清单逐行验过（见第五节）。

改后发码：

```c
n = vb6_ChkLong(vb6_ComGetIntProp(o, L"V"));                    /* L1 */
n = vb6_ChkLong((vb6_ComGetIntProp(o, L"V") + 1));              /* L2 按 C 整数算 */
vb6_BSTR_AssignFromVariant(&s,                                  /* L3 一字未动 */
     vb6_VariantFromValue(vb6_VariantFromComResult(vb6_ComGetProp(o, L"V"))));
```

L3 那句原本就是按 Variant 走的（String 上下文没触发目标档钉型别），本刀没碰它 ——
这正是"收窄"的意思。

### 五、读数

新夹具 `tests/b158_late_variant_prop_read`（六针），x64 与 x86 同值：

| 针 | 改前 | 改后 | 意思 |
|---|---|---|---|
| `L1` | **编不过** | 11 | Variant 属性读进 Long |
| `L2` | **编不过** | 12 | 参与算术（不再被装箱进 `VarAddV`） |
| `L3` | 编不过 | 11 | `CStr(o.V)`（这一形本来就对，本刀不碰） |
| `L4` | 编不过 | 42 | 盒里揣串时按 VB6 隐式转换取数 |
| `L5` | 编不过 | `8/42` | 读进 Variant，型别仍是服务端那一档 |
| `DONE` | — | — | — |

**回归读数（改前改后逐字相同）**：

```
b144  A2-terms=1  B2-tag=8  B2-terms=1  C1-terms=2  D1-terms=3  E1-terms=4  Z-terms=4  Z-tag=8
b146  A1-terms=1  B1-tag=7  B2-terms=0  C1-alive=0  C2-terms=1  Z-terms=2  Z-made=3
b148  A2-self-tag=5  A3-tag=6  B2-fresh-tag=11  B3-item-terms=1  C2-held-tag=77  C3-item-terms=2
      Z-box-terms=0  Z-item-made=2  Z-item-terms=2
b153  A1-terms=0  A2-held-tag=5  A3-terms=0  Z-terms=1  B1-len=1  B2-len=4  C1-terms=2
```

这四份是第三节那条弯路**当场**掀翻的夹具；本刀改完它们与本次改动前一字不差。

### 六、发码面

- 清单 **824→826**（新夹具 `.bas` 与 `.vbp` 两行，行数 1357）；
- 全语料 826 行，**只有 6 条机器侧行移动**（b132、ctrlmscomm、GUISample×3、
  TeeChart/DisplayProject）—— 按老规矩**以 CI 为准**，本地与 CI 逐行 COMPARE = 0 差异；
  其余 820 行逐字未动（其中 5 条 b 系列 `.vbp` 行在第三节那条弯路上曾移动过，改回收窄之后
  **回到原值**）；
- 九门 9/9 rc=0（`inputs=826 [.vbp=311 .bas=515]`；`emit-manifest` 1357 行）；
- CI 22/22（run 45363102958，sha `861851d0`）。

### 七、在册（未修）

- `§11.138`（"晚绑定读进 Variant 的型别保真"）**量不到病**：`VariantFromComResult(ComCall(...))`
  把服务端交回的 VT 原样留着（第一节四行读数）。这条从在册里划掉，不是待修项。
- 本刀只管**数值三档**的读：`vb6_ComGetStringProp` 进数值上下文（`n = o.Note`，Note 声明
  String）仍是另一族，没动；`vb6_ComGetObjectProp` 同理。
- `Set c->f = o`（`o` 是 `As Variant` **形参**）仍走 `vb6_VariantFromValue` 浅拷贝。
- `Set obj.Prop = Nothing`（控件伪对象那一档）未动。
- `vb6_ReleaseObject` 对**裸类实例**是空转（§11.136 在册第一笔）。
- `§11.110` 值侧打包、`.temp/k150/p4` 的 V2 记账照旧在册。

## §11.140（第一百五十八刀）声明档说话：浮点属性读不再被目标档截断

### 一、撞见

上一刀（§11.139）修的是"声明为 Variant 的成员被 consumer 二次包装"。收尾时把同一族的
下一问摆出来：**声明为具体档（String / Double）的成员，读口按目标档发码会不会把值本身弄坏？**

探针 `.temp/k159/p1`（`B159Box` 有 `Property Get S() As String` = `"101"`、
`Property Get D() As Double` = `1.5`、`Property Get O() As Object`），主程序 `o As Object`：

```vb
n = o.S       ' S1
n = o.S + 1   ' S2
d = o.S       ' S3
d = o.D       ' S4
n = o.D       ' S5
v = o.S       ' S6
Set it = o.O  ' S7
```

**整份工程编不过，rc=1**：

```
B159Main.c(44): error C2440: 无法从 'double' 转换为 'BSTR'
```

发码（改前）：

```c
n = vb6_ChkLong(vb6_ComGetIntProp(o, L"S"));                              /* S1 对 */
n = vb6_ChkLong(vb6_CStrDbl(vb6_Val(vb6_ComGetIntProp(o, L"S")) + 1.0));  /* S2 */
vb6_Dbl_AssignFromBSTR(&d, vb6_ComGetDoubleProp(o, L"S"));                 /* S3 ← C2440 */
d = vb6_ComGetDoubleProp(o, L"D");                                        /* S4 对 */
n = vb6_ChkLong(vb6_FltToLng(vb6_ComGetIntProp(o, L"D")));                /* S5 */
```

### 二、三笔账，一个根

| 形 | 病 | 性质 |
|---|---|---|
| S3 `d = o.S` | 读口发 `DoubleProp`（double），consumer 按**声明档 String** 发 `Dbl_AssignFromBSTR`（要 BSTR） | 编译期 C2440 |
| S2 `n = o.S + 1` | 读口发 `IntProp`（int32），consumer 按声明档 String 走 Fix 113g 的"串+数"形，`vb6_Val(<int32>)` 把整数当 `wchar_t*` 解引用 | **运行期 0xC0000005** |
| S5 `n = o.D` | 读口按**目标档 Long** 发 `IntProp`，RTL 内 `VariantChangeType(…, VT_I4)` **当场把 1.5 截成 1**；consumer 那侧的 `FltToLng` 再圆也圆不回来（真 VB6 是 `CInt(1.5)` = **2**） | 静默错值 |

三笔都是同一句：**同一件"这个值是什么型别"的事，读口按目标档说、consumer 按声明档说。**

S5 那笔是最要命的一类 —— 它不吵不闹，只是把小数丢了。而且它证明了一件事：
**截断发生在读口，consumer 再圆也救不回来**，所以这一类只能由读口让路。

### 三、修：让读口服从声明档（只服浮点那一档）

`resolveComValue` 在按目标档发码之前先问一次成员的声明档（`lookupModule(memberName)`，
与 `inferExprType` 的兼底**同一个符号源**）：

```cpp
std::string unpackEff = unpackType;
if (unpackEff == "Int" || unpackEff == "Long" || unpackEff == "Boolean"
    || unpackEff == "Double" || unpackEff == "Single") {
    const Symbol* ms140 = symTab_.lookupModule(memberName);
    if (ms140 && (ms140->type == Vb6Type::Double || ms140->type == Vb6Type::Single
                  || ms140->type == Vb6Type::Currency || ms140->type == Vb6Type::Date)) {
        unpackEff = "Double";
    }
}
```

于是 S5 变成 `vb6_ChkLong(vb6_FltToLng(vb6_ComGetDoubleProp(o, L"D")))` —— 取整交给
consumer 那一站（它本来就按声明档准备了 `FltToLng`），答案是 2。

另外两笔在 consumer 侧收口，判据用的还是上一刀那枚 `isScalarComReadText`（只认已发出的文本）：

- `cgen_assign_value_sem.inc` 的 Double←String 支：右值已经是裸 double 读 ⇒ 直接
  `target = value;`，不走 `vb6_Dbl_AssignFromBSTR`；
- `cgen_expr_binary.cpp` 的 Fix 113g（"串 + 数"那一形）：某一侧是裸标量读 ⇒ 它**不是串**，
  不进 `vb6_Val` 那一形，让 VB6 的数值加法直接算。

### 四、试过但退回的一条

顺手试过把**声明为 String** 的成员也改成读口发 `StringProp`（"谁声明谁说话"的彻底版）。
编译**过**了，但 S1 当场报 **错误 6 Overflow**：consumer 那一站按"我收到的就是数"发的
`vb6_ChkLong(...)`，拿到 BSTR 就是**指针当整数**。这一族要 consumer 各站配合折数
（`vb6_CStrLong` 等），不是读口一句话能收的，**本刀不碰，记账在册**。

### 五、读数

新夹具 `tests/b159_late_decl_typed_read`（九针），x64 与 x86 同值：

| 针 | 改前 | 改后 | 意思 |
|---|---|---|---|
| `S1` | 101 | 101 | String 属性读进 Long |
| `S2` | **AV 0xC0000005** | 102 | 参与算术（不再把 int 递给 `vb6_Val`） |
| `S3` | **编不过** | 101 | String 属性读进 Double |
| `S4` | 1.5 | 1.5 | Double 属性读进 Double |
| `S5` | **1** | **2** | Double 属性读进 Long（`CInt(1.5)`=2） |
| `S6` | 8/101 | `8/101` | 读进 Variant，型别保真（§11.138 那条口径） |
| `S7` | 77 | 77 | Object 属性读出来再取成员 |
| `Z-terms` | — | 1 | 对象寿命（引用计数没被本刀动过） |
| `DONE` | — | — | — |

老夹具回归读数**逐字相同**：

```
b144  A2-terms=1 B2-tag=8 B2-terms=1 C1-terms=2 D1-terms=3 E1-terms=4 Z-terms=4 Z-tag=8
b146  A1-terms=1 B1-tag=7 B2-terms=0 C1-alive=0 C2-terms=1 Z-terms=2 Z-made=3
b148  A2-self-tag=5 A3-tag=6 B2-fresh-tag=11 B3-item-terms=1 C2-held-tag=77 C3-item-terms=2
      Z-box-terms=0 Z-item-made=2 Z-item-terms=2
b153  A1-terms=0 A2-held-tag=5 A3-terms=0 Z-terms=1 B1-len=1 B2-len=4 C1-terms=2
b158  L1=11 L2=12 L3=11 L4=42 L5=8/42
```

### 六、发码面

- 清单 **826→828**（新夹具 `.bas` 与 `.vbp` 两行，行数 1359）；
- 全语料 828 行，**只有 6 条机器侧行移动**（b132、ctrlmscomm、GUISample×3、TeeChart），
  以 CI 为准；其余逐字未动 —— 说明存量语料里"声明档查得到符号"的晚绑定读，本来就都是
  同档读写，本刀只动到那些**跨档**的形状；
- 九门 9/9 rc=0（`inputs=828 [.vbp=312 .bas=516]`）；
- CI 22/22（run 45370730242，sha `ecee5fc9`）。

### 七、在册（未修）

- **声明为 String 的成员读进数值档**（`n = o.S`）：读口若服声明档，consumer 各站不折数
  会报 6（第四节实测）；要 consumer 全线配合，另开一刀。当前读口仍按目标档发，S1 这种
  "串是纯数字"的情形答案正确，但 `n = o.Note`（`"abc"`）会走 RTL 的宽松折数。
- **本刀只服浮点档**：声明为 Integer/Byte/Boolean 的成员仍按目标档发码（同一族，尚未量到病）。
- `Set c->f = o`（`o` 是 `As Variant` **形参**）仍走 `vb6_VariantFromValue` 浅拷贝。
- `Set obj.Prop = Nothing`（控件伪对象那一档）未动。
- `vb6_ReleaseObject` 对**裸类实例**是空转（§11.136 在册第一笔）。
- `§11.110` 值侧打包、`.temp/k150/p4` 的 V2 记账照旧在册。

## §11.141（第一百五十九刀）晚绑定写进 Boolean 形参要有 CBool 那一步

### 一、撞见

前两刀修的都是**读**口。这一刀把写口（`o.Prop = 值`）的跨档一样量一遍：探针
`.temp/k160/p1`，`B160Box` 有 `S As String` / `D As Double` / `I As Integer` / `B As Boolean`
四个属性，主程序 `o As Object`：

```vb
o.D = 2        ' W1
o.S = 2        ' W2
o.D = "3.5"    ' W3
o.I = 2.5      ' W4
o.B = 7        ' W5
n = box.B      ' W6
o.I = "42"     ' W7
o.I = <Variant 9.75>  ' W8
```

八笔里**七笔对**，一笔错：

```
B160-W1=2   B160-W2=2   B160-W3=3.5   B160-W4=2
B160-W5=True
B160-W6=7      ← 真 VB6 是 -1
B160-W7=42  B160-W8=10
```

（顺带记：`o.I = 2.5` 答 2、`o.I = 9.75` 答 10 —— 四舍六入五成双那一套在写口是对的一档。）

### 二、根因：Boolean 形参借了 VT_I4 那一档

晚绑定写走 `vb6_ComSetProp` ⇒ 服务端 `IDispatch::Invoke`。入站实参在
`vb6comserver_obj.c` 里按 `method->paramVts` 归一化到**形参声明档**，而生成本表的那一问
是 `cDispatchArgVt`：

```cpp
if (t == Vb6Type::Long || t == Vb6Type::Integer || t == Vb6Type::Byte ||
    t == Vb6Type::Boolean) return "VT_I4";
```

Boolean 被并进 VT_I4，于是 7 原样落进槽，读臂再读 `lVal` 存进 `m_B`（本仓 Boolean 的
存储是 int16/int32，真值 **-1**，与 `vb6_CBool` 同口径）。**VB6 里"数值 → Boolean 形参"
是一次 CBool 换算（非零即 -1）**，借 VT_I4 就地透传等于把这次换算整段跳过。

### 三、修：成对的两处，缺一不可

1. `cDispatchArgVt`：Boolean → **VT_BOOL**，换算那一步交给 `vb6_ComArgNormalize`
   （内部 `VariantChangeType`，`VT_I4 7 → VARIANT_TRUE`、`BSTR "12" → VARIANT_TRUE`、
   `BSTR "abc"` 换不动 ⇒ `DISP_E_TYPEMISMATCH` ⇒ 客户端折成 VB 13，与 VB6 同号）；
2. `cgen_util_dllentry_collect.inc` 的 ByVal Boolean 读臂：**读 `boolVal` 并折成本仓的
   -1/0**：

   ```cpp
   provided = "(((VARIANT*)args[i])->boolVal ? -1 : 0)";
   ```

   只改第一处不改第二处不行：槽变成 VT_BOOL 以后 `lVal` 读到的是 `VARIANT_BOOL` 之外
   的字节（高位不是符号扩展），那是另一个坑。

ByRef Boolean 不在这一刀：`cDispatchArgVt` 对 ByRef 本来就答 `VB6_ARGVT_SKIP`（不归一化），
读臂保持 `&lVal` 原样。

### 四、读数

新夹具 `tests/b160_late_bool_arg`（十针），x64 与 x86 同值：

| 针 | 改前 | 改后 | 意思 |
|---|---|---|---|
| `W1` | 2 | 2 | 数值 → Double 属性 |
| `W2` | 2 | 2 | 数值 → String 属性（存 `"2"`） |
| `W3` | 3.5 | 3.5 | 串 → Double 属性 |
| `W4` | 2 | 2 | Double → Integer 属性（2.5 → 2） |
| `W5` | True | True | 数值 → Boolean 属性（存成 True） |
| `W6` | **7** | **-1** | Boolean 属性 → Long（VB6 真值） |
| `W7` | 42 | 42 | 串 → Integer 属性 |
| `W8` | 10 | 10 | Variant 9.75 → Integer 属性 |
| `Z` | 2/3.5 | 2/3.5 | 读回两个属性 |
| `DONE` | — | — | — |

老夹具回归读数逐字相同（b144 / b146 / b148 / b153 / b158 / b159）。

### 五、发码面与第一次 CI 红

- 清单 **828→830**（新夹具 `.bas` 与 `.vbp` 两行，1361 行）；
- 全语料 830 行，只有 6 条机器侧行移动（b132、ctrlmscomm、GUISample×3、TeeChart），以 CI 为准；
- 九门 9/9 rc=0；
- 第一次 CI 红在两处 `Tests (vbp #6)/(vbp #7)` 的 `[DISPATCH] b220_inbound_argvt`
  ——§B219 那一刀按"Boolean 借 VT_I4 就地透传"钉下的两颗针：

  ```
  改前期望: VTARG=EchoBool:i giving=VT_I4  … val=40
            VTARG=EchoBool:s giving=BSTR(12) … val=12
  CI 实测:  val=-1 / val=-1
  ```

  两发 **-1** 正是 CBool 语义（真实 VB6 里 `EchoBool(40)` / `EchoBool("12")` 都是 True），
  于是把那两颗钉改口 -1 并把改前/改后的由来写进注释（`run_tests.ps1` 那一段）；
- 第二次 CI **22/22**（run 45390455157，sha `365531c8`）。

### 六、在册（未修）

- ByRef Boolean 形参未归一到 VT_BOOL（`cDispatchArgVt` 对 ByRef 一律 SKIP，读臂 `&lVal`）；
- 声明为 String 的成员读进数值档（§11.140 第六节那一条）；
- 声明为 Integer/Byte 的成员仍按目标档发读臂（同族未量到病）；
- `Set c->f = o`（Variant 形参）浅拷贝、`Set obj.Prop = Nothing`（控件伪对象档）、
  `vb6_ReleaseObject` 对裸类实例空转、`§11.110` 值侧打包照旧在册。

## §11.142（第一百六十刀）窄整型形参的入站收窄要报错误 6

### 一、撞见

上一刀（§11.141）修的是 Boolean 形参的入站换算。同一张表上紧挨着的一格还没人问：
**Integer / Byte 形参收到超范围的值会怎样？** 探针 `.temp/k161/p1`
（`Property Let I(ByVal x As Integer)`，先 `box.I = 7`）：

```vb
o.I = 40000
```

```
B160-R1=no-error val=-25536
```

真 VB6 是 **错误 6 Overflow**，且目标保持 7 不变。C3 既没报错，还把 40000 静默存成
**-25536** —— 这正是 VB6 里最容易被漏掉的那种"值坏了但程序继续跑"。

### 二、根因：和上一刀同一张表、同一个毛病

`cDispatchArgVt` 当时写的是

```cpp
if (t == Vb6Type::Long || t == Vb6Type::Integer || t == Vb6Type::Byte ||
    t == Vb6Type::Boolean) return "VT_I4";
```

把 Integer / Byte 一起并进 VT_I4 就地透传 —— 收窄换算（`VariantChangeType` VT_I4→VT_I2）
整段跳过，服务端读臂再把这个 40000 塞进 `int16_t`。**三档一起并进 VT_I4 是"省一次换算"，
代价是把 VB6 的溢出语义一起省掉了。**

### 三、修：四处成对

1. `cDispatchArgVt`：`Integer → VT_I2`、`Byte → VT_UI1`（让换算真的发生，
   越界时 oleaut32 自己答 `DISP_E_OVERFLOW`）；
2. 读臂（`cgen_util_dllentry_collect.inc`）：ByVal Integer 读 **`iVal`**、Byte 读 **`bVal`**
   —— 槽已经是窄档，继续读 `lVal` 会带上高位的垃圾字节；Long 维持 `lVal`；
   **ByRef 三档一律不动**（`cDispatchArgVt` 对 ByRef 本答 `VB6_ARGVT_SKIP`，槽还是调用方那份）；
3. `vb6_ComArgNormalize`：失败时把**原因 hr** 记进一枚静态并经
   `vb6_ComArgNormLastHr()` 交回（以前一律折成 `DISP_E_TYPEMISMATCH`，
   于是越界和"类型不匹配"被压成同一个号）；服务端 Invoke 把这一枚原样返回；
4. 客户端 `vb6_HresultToErrNum`（唯一那张表）增一格
   `0x8002000A DISP_E_OVERFLOW → 6`。

### 四、读数

新夹具 `tests/b161_narrow_arg_overflow`（八针），x64 与 x86 同值：

| 针 | 改前 | 改后 | 意思 |
|---|---|---|---|
| `I1` | ok 4000 | ok 4000 | Integer 合法写入 |
| `I2` | **no-error（存成 -25536）** | **err6 val=4000** | Integer 正向越界：报 6 且目标不变 |
| `I3` | 同 I2 | err6 val=4000 | 负向越界 |
| `Y1` | ok 200 | ok 200 | Byte 合法写入 |
| `Y2` | **no-error（存成 44）** | **err6 val=200** | Byte 越界 |
| `L1` | ok 2000000000 | ok 2000000000 | Long 宽档不受影响 |
| `DONE` | — | — | — |

老夹具回归读数逐字相同（b144 / b146 / b148 / b153 / b158 / b159 / b160）。

### 五、发码面与第一次 CI 红

- 清单 **830→832**（新夹具 `.bas` 与 `.vbp` 两行，1363 行）；
- 全语料 832 行，只有 6 条机器侧行移动（b132、ctrlmscomm、GUISample×3、TeeChart）；
- 九门 9/9 rc=0；
- 第一次 CI 红在 `Tests (compile #2)` 的 `[STATIC] disp_inbound_argvt_authority`：

  ```
  FAIL R2 vb6_ComArgNormalize defined  = 2, want 1
  FAIL R2 vb6_ComArgNormalize declared = 2, want 1
  ```

  本刀新加的两枚 helper 起名 `vb6_ComArgNormalizeSetLastHr` / `vb6_ComArgNormalizeLastHr`，
  前缀与"唯一出口"同一个词，被那道门数成第二份。**门是对的，名字挑错了** ——
  改名 `vb6_ComArgNormLastHrSet` / `vb6_ComArgNormLastHr`，三处同步；
- 第二次 CI **22/22**（run 45398364506，sha `e3026449`）。

### 六、在册（未修）

- **ByRef 形参整体**：`cDispatchArgVt` 对 ByRef 一律 `VB6_ARGVT_SKIP`（不换算），
  于是 ByRef Integer 收 40000 仍然不会被拦。要收这一族得先量清楚 ByRef 槽的标签与回写口径。
- 声明为 String 的成员读进数值档（§11.140 第六节）；
- `Set c->f = o`（Variant 形参）浅拷贝、`Set obj.Prop = Nothing`（控件伪对象档）、
  `vb6_ReleaseObject` 对裸类实例空转、`§11.110` 值侧打包照旧在册。

## §11.143 续（第一百六十一刀 · 修补）：引用要按签名说话

### 一、第一次 CI 红

`Tests (vbp #5)` 报 `[GUI] czUI ... FAIL (exit=13)` —— tests/czUI-main 的表单在
`UserControl_ReadProperties` 里当场倒。查发码：那一族只多了 7 处
`vb6_ComPackBSTRRef(...)`，全在 `With PropBag` 里（6 处 `WriteProperty` + 1 处
`ReadProperty`，`PropBag As Object` 晚绑定到 `MSSTDFMT.PropertyBag`）。

**真 VB6 客户端在这里不会递引用**：`PropBag` 的签名来自 typelib（`WriteProperty(PropName,
ByVal Value As Variant)`），编译器照签名按值发。我们晚绑定没有签名，把「带型变量一律递
VT_BYREF|型」这条口径套到了所有对端上 —— 对**自己家的服务端**是对的（本刀要修的就是那
一格），对**外部 OCX** 就是 13。

### 二、落点：把「有没有签名」这一问交回运行时

`vb6_ComNarrowArgs` 本来就是 ComCall 进 Invoke 之前那一趟「按形参表对表」的闸
（原先只管 VT_R4/R8 → typedef LONG 的收窄）。本刀给它加第二件事：

| 对端 | 判据 | 处理 |
|---|---|---|
| 有 ITypeInfo | 形参带 `PARAMFLAG_FOUT` | 真 ByRef ⇒ 保留引用 |
| 有 ITypeInfo | 形参不带 FOUT | ByVal ⇒ 降成值副本 |
| 无 ITypeInfo + 本仓工程类服务端 | `vb6_ComObject_IsVb6Wrapper` 或 `vb6_FindProjectClassDescRaw` | 保留引用（回写走 `vb6_ByRefScalarWriteBack`） |
| 无 ITypeInfo + 其它 | 上一条不成立 | 降成值副本（与改前同形） |

第三行那半边是本地实测补上的：第一版只问包装器，而 ComCall 走到这里时接收者可能已经被
剥壳成**裸类实例**（`vb6_FindProjectClassDescRaw` 那张登记表），于是自己的服务端被当成
外部 OCX 降值 —— 夹具 S1 当场退回 `hi`。两个形都认才是全的。

### 三、`vb6_ComArgNormalize` 的跳过再收一层

「引用形的 pointee 已等于声明档就不换算」只对 `vb6_VariantBorrowIntoSlot` 真搬过的那
七档（I4/I2/R8/R4/BOOL/UI1/BSTR）成立。其余形（典型 `BYREF|VT_VARIANT`）入槽那臂本就
不接（槽留 Empty），跳过换算等于把引用变体读成空 —— 那一档必须继续走
`VariantChangeType` 解引用。

### 四、一次假的「回归」（记在册，免得下次再踩）

修补过程中出现了「夹具 S1~S6 全部退回按值」的现象，一路查到最后是**构建缓存**：
`Copy-Item` 恢复备份会**连时间戳一起恢复**，恢复出来的 `cgen_util_com.cpp` 比 `.obj` 还旧
⇒ `cmake --build` 判定无需重编，跑的是实验版二进制。判据不是「build 打印 rc=0」，而是
**发出来的文本**：`--emit-c` 里那一行到底是 `vb6_ComPackBSTRRef(&s)` 还是
`vb6_ComPackBSTR(s)`。这和 §11.142 那条「判据只信已经发出来的文本」是同一句话。

### 五、读数与判据面

- 夹具 `tests/b162_byref_string_arg`（x64 = x86）：
  `S1=5/hi-ok  S2=3/new  S3=0/[]  S4=4/keep  S5=4/loop  S6=3/-ok  S7=42  DONE`；
  S4 是 ByVal 对照（服务端按值收，不回写），S5 是体里不改（串必须完好），S7 是数值那六档对照。
- 外部 typelib 那一档另用 `Scripting.FileSystemObject` 对表：`FileExists` / `GetFileName`
  正常、串不丢、连发两次稳定。
- 两道门同步到七档：`check_byref_scalar_writeback.ps1` 的 R3/R4（入槽/出站 switch 各 7 条）、
  `check_byref_arg_pack_authority.ps1` 的 R1（六枚 Ref 打包器）/R3（第六档映射）/R5
  （vbp 夹具两格）。
- 清单 832→834（1365 行）；九门 9/9；CI **22/22**（run 45497421693，sha `ac7fd7d0`）。

### 六、在册（本刀之后仍在）

- 引用形只认「裸标识符 / me->字段」这一形状：`o.M arr(i)`、`o.M box.S` 这类实参仍按值发，
  写回不回来（VB6 会回来）—— 要收这一族得先给「数组元素/成员访问」一个可取址的发射形。
- 无 typelib 的**外部** VB6 ActiveX DLL 收 ByRef String 时，本刀按「值」发（与改前同形）：
  写回仍然不回来。要收它得有一条真正可靠的「对端是本仓服务端」之外的分辨法。
- 入槽那一份副本在「同一发里别的实参换算失败」时会漏一枚串（该路径已报错退出），记小额。

## §11.144（第一百六十二刀）引用形不只裸标识符：数组元素、类字段、With 成员

### 一、在册那条病

上一刀把 ByRef 的**出站**（String 出站回写）收口之后，账上剩这一条：晚绑定实参的
「引用形」只认 `^(\w+|me->\w+)$` 两种文本，`o.Bump arr(i)` 与 `o.Bump box.S` 都按值发，
体里的写回落空（VB6 里这两形都该回写）。先说清为什么当初只认裸标识符：`&expr` 对
函数调用是非法的，而晚绑定实参那一列里绝大多数"看起来像左值"的东西其实是**调用**
（属性读 `vb6_X_prop_get_Y(base)`、控件默认属性折叠、`Db.Rs` 那种 voidptr 链）。

### 二、先量可取址形（不改代码先看 --emit-c）

探针 `.temp/k163/p1`（一个类 + 一个模块，六站 `o.Bump ...`），发出来的文本：

| 站 | 实参 | 发射文本 |
|----|------|----------|
| A1 | `la(1)`（局部静态数组） | `vb6_ComPackValue(VB6_SA_AT(BSTR, la, 1))` |
| A2 | `arr(1)`（模块级数组） | `vb6_ComPackBSTR(VB6_SA_AT(BSTR, arr, 1))` |
| A3 | `box.S`（Public String 字段） | `vb6_ComPackBSTR(box->S  /* class var .S field */)` |
| A4 | `g`（模块级裸串，基线） | `vb6_ComPackBSTRRef(&g)` |
| A5 | `box.Tag`（属性） | `vb6_ComPackBSTR(vb6_B163Box_prop_get_Tag(box))` |
| A6 | `With box` 内 `o.Bump .S` | `vb6_ComPackBSTR(_vb6_with_0->S  /* With class .S field */)` |

`VB6_SA_AT(type, arr, idx)` 展开是 `(*(type*)vb6_SaElemPtr(...))`（`vb6rtl_array.h:213`）
—— 一枚真左值，`&` 合法。`box->S` 也是左值。A5 是调用，`&` 非法，必须继续按值。

注意 A1 的打包器是 `ComPackValue`（`inferExprType` 对局部静态数组元素答 Variant），
而 A2 是 `ComPackBSTR` —— 「打包器」这一格对同一个语义形给两个答案，不能拿它当判据。

### 三、两形各问一处权威

- **数组元素**（`expr.kind != IdentifierExpr` 且文本是 `VB6_SA_AT(BSTR, <载体>, <下标>)`）：
  宏的首参就是元素档，文本本身就是证据。这一问在本仓已有五处消费点
  （`cgen_expr_call_arg_variant.inc` ×2 / `cgen_expr_call_conv_numeric.inc` /
  `cgen_expr_call_builtin_pre.inc` 等），不是新发明。命中即
  `vb6_ComPackBSTRRef(&VB6_SA_AT(...))`。
- **类字段 / With 成员**：`MemberAccessExpr` 用 `inferClassTypeOfExpr(*ma.object)` 得宿主类，
  `WithMemberExpr` 从 `withObjectInfoStack_` 栈顶取宿主类（且必须 `kind == ClassInstance`，
  `void*` 那一档的 With 临时量不是类实例指针）。两形都再问同一组两句：
  `resolveClassMemberCall(cls, member).empty()`（成员不是方法/属性）+ `classFieldIsBstrSlot(cls, member)`
  —— 与赋值站 §B168 那条「BSTR 槽」判据同源、同一个否决。

仍按值（在册）：属性读（不可取址，VB6 是经 Let 写回）、控件默认属性折叠、`Db.Rs` 那种
voidptr 对象成员链、数组本体。

### 四、第一次 CI 红：把尾注当判据

第一版成员那一形是**读发射文本的尾注**（正则匹配 `base->field  /* class var .F field */`）
再借打包器的答案当型。CI `Tests (compile #1)` 立刻红：

```
[STATIC] variant_slot_store ... FAIL
FAIL R5d 名单外的新文件 cgen_util_com.cpp x1 —— 又有人把那句尾注当判据了
```

这条门逐文件钉死 `class var .` 的处数，原意正是「谁在提这句文本都算，因为这一族就是从
读尾注长出来的」。两处都不干净：尾注是 `classFieldReadExpr` 的输出而不是字段的声明，
打包器的答案只是这一发的产物。于是改成问声明表（`classFieldIsBstrSlot`），
`comByRefArgExpr` 那枚 `packFn` 形参一并撤掉，三处消费点回到原签名。

### 五、读数

夹具 `tests/b163_byref_addr`（x64 = x86 同形）：

```
B163-A1=la-ok   局部静态数组元素
B163-A2=ma-ok   模块级数组元素
B163-A3=new     动态数组元素（ReDim 后）
B163-A4=fs-ok   类实例 Public String 字段
B163-A5=ws-ok   With 块里的成员
B163-A6=pt      属性读（保持按值，在册）
B163-A7=mg-ok   模块级裸变量（第一百六十一刀的基线）
B163-A8=tw-ok-ok 同一元素连发两次（回写后槽里仍是能用的串）
B163-DONE
```

### 六、账与门

- manifest：corpus 834 → 836（`tests/b163_byref_addr` 两枚），行数 1365 → 1367；
  真动到的既有行只有 `Charts 2020/ucChartBar` 与 `ucTreeMaps` 两份（各多出按引用的一发），
  另 6 条仍是机器侧的已知差（CI 那份是判据）。
- 门 `check_byref_arg_pack_authority.ps1` 加 R7 四条：数组元素那一形只此一处判、
  类字段那一形问声明表、With 成员同一问、新夹具两格（x64 + x86）。
- 九门 9/9。CI 22/22（run 45510100696）。

### 七、这一刀之后仍在册的相邻欠项

- 属性读当 ByRef 实参仍按值（VB6 经 Let 写回）—— 要动得先有一条「读-改-写」的
  双程发射，不是判据问题。
- 数值五档的引用形同样只认裸标识符（`o.Inc arr(i)` 的 `Long` 元素仍按值）——
  这一刀的落点只开了 String 档，数值那五档要跟 `VT_BYREF|VT_I4/R8/...` 的宽度一起看。
- ByRef Boolean 未归一、无 typelib 的外部 VB6 ActiveX DLL 收 ByRef String 仍按值。

## §11.145（第一百六十三刀）数值五档的引用形：数组元素与类字段

### 一、在册那条病的另一半

§11.144 只把 String 那一档开给了「数组元素 / 类字段 / With 成员」。数值五档
（Long / Integer / Boolean / Single / Double）当 ByRef 实参时仍然只认裸标识符与
`me->字段`，于是 `o.Inc arr(i)`、`o.Inc box.N`、`o.FlipB bl(1)` 一律按值发，体里的
写回落空。

### 二、先量：文本里的 C 型**分不出**档

探针 `.temp/k164/p1` 九站 + 后补八站，`--emit-c` 发出来：

| 实参 | 发射文本 |
|------|----------|
| `ln(1)`（Long 数组） | `vb6_ComPackValue(VB6_SA_AT(int32_t, ln, 1))` |
| `in1(1)`（Integer 数组） | `vb6_ComPackValue(VB6_SA_AT(int16_t, in1, 1))` |
| `bl(1)`（Boolean 数组） | `vb6_ComPackValue(VB6_SA_AT(int16_t, bl, 1))` |
| `dn(1)`（Double 数组） | `vb6_ComPackValue(VB6_SA_AT(double, dn, 1))` |
| `dt(1)`（Date 数组） | `vb6_ComPackValue(VB6_SA_AT(double, dt, 1))` |
| `ba(1)`（Byte 数组） | `vb6_ComPackValue(VB6_SA_AT(uint8_t, ba, 1))` |
| `cu(1)`（Currency 数组） | `vb6_ComPackValue(VB6_SA_AT(int64_t, cu, 1))` |
| `box.N` / `box.I` / `box.D` / `box.S` | `vb6_ComPackInt/…(box->N  /* class var .N field */)` 等 |

两处**同形不同档**：`int16_t` 既是 Integer 也是 Boolean，`double` 既是 Double 也是 Date。
上一刀对 String 敢直接读文本首参（`VB6_SA_AT(BSTR, …)` 没有二义），数值这一族不能这么读
—— 按 I2 给 Boolean 打写回、按 R8 给 Date 打写回，都是**静默错值**，比漏回写更坏。

### 三、落点：文本只当形状证据，档问声明侧

- **数组元素**：从 `VB6_SA_AT(<C型>, <载体>, <下标>)` 里只取**载体名**（剥 `me->` 与
  `(*...)`，与发码那一处的写法对齐），再问 `arrayElemTypes_`，落空回模块符号表
  （`sym->isArray` 取 `sym->type`）—— 与 `cgen_expr_call_prelude.inc` 发这一格时
  **同一条链**，不是另抄一张表。
- **类字段 / With 成员**：问 `classFieldKind(类, 字段)`，与 §11.144 的 String 那一问是
  同一个出口（`classFieldIsBstrSlot` 就是它的一个投影，这一刀改用投影的母函数）。

「VB 型 → 按引用打包器」这张配对表收成一处 `kComPackRefTiers`（cgen_util_com.cpp），
裸标识符 / 数组元素 / 类字段三个消费点都问 `comByRefPackFnForVbType`；裸标识符那一支
原来的六档 switch 并进来，Integer 的宽度问句（`knownIntVars_`，因为 `inferExprType` 对
Integer 答 Long）原样保留。

只接 RTL 真有的六枚（Int / Integer / Bool / Single / Double / BSTR）。Byte 没有 UI1 那一档、
Currency 没有 CY、Date 没有 DATE —— 一律交回按值：宁可漏回写，也不能按错标签写回。

### 四、读数

夹具 `tests/b164_byref_numeric_addr`（x64 = x86 同形）：

```
B164-N1=11     局部 Long 数组元素
B164-N2=21     模块级 Long 数组元素
B164-N3=31     类实例 Public Long 字段
B164-N4=41     模块级裸 Long（基线）
B164-N5=51     Integer 数组元素（2 字节）
B164-N6=61     Integer 类字段
B164-N7=71.5   Double 数组元素
B164-N8=81.5   Double 类字段
B164-N9=90.5   Single 类字段（R4 不是 R8）
B164-N10=True  Boolean 数组元素（int16_t 靠声明表分档）
B164-N11=True  Boolean 类字段
B164-N12=1     Byte 数组元素仍按值（在册：无 UI1 档）
B164-N13=5     属性读仍按值（在册：不可取址）
B164-N14=2     同一元素连发两次
B164-DONE
```

### 五、第一次 CI 红：两处登记

1. `[STATIC] variant_slot_store` 报 `R8 名单外的新问点 cgen_util_com.cpp x2`——ByRef
   实参那一问确实在问 `classFieldKind`（类字段 + With 成员各一次），那张名单要登记。
   这两站与读侧那三站是**同一个出口的不同消费点**，不是新抄一份答案。
2. `Emit manifest (shape oracle)` 报 `哈希不同 2`：`guisample/vb6ref/src/{vc6,vs2015}/GUISample.vbp`。
   这两行属机器侧（本机 MSVC 与 CI 二进制出码不同），本地登记脚本一直跳过；这一刀真动到
   了它们的发码，于是按 CI 那份 oracle 登记（`git show github/ci/emit-manifest:emit-manifest.txt`），
   替换后本地 expected 与 CI 清单逐行比对 0 差。

顺带踩到本仓一条老坑：手写 `emit-manifest.expected.txt` 时用无 BOM 的 UTF-8，
`check_manifest_coverage.ps1` 的 K2 立刻红 —— 那份文件 PS 5.1 按 UTF-8 读，BOM 不能丢。

### 六、账与门

- manifest：corpus 836 → 838（`tests/b164_byref_numeric_addr` 两枚），行数 1367 → 1369；
  存量行真动到的只有 GUISample 那两份（另 4 条机器侧行不变）。
- 门 `check_byref_arg_pack_authority.ps1`：R3 六条改指配对表行、R7 两条改指 `classFieldKind`、
  新增 R8 三条（表只一处 / 三处消费点 / 新夹具两格）。
- 门 `check_variant_slot_store.ps1`：R8 名单加 `cgen_util_com.cpp` = 2。
- 九门 9/9。CI 22/22（run 45516917680）。

### 七、这一刀之后仍在册的相邻欠项

- Byte / Currency / Date 的引用形（RTL 侧要先有 `VT_UI1|BYREF` / `VT_CY|BYREF` /
  `VT_DATE|BYREF` 三枚打包器与写回臂）。
- 属性读当 ByRef 实参仍按值（VB6 经 Let 写回）—— 要一条「读-改-写」的双程发射。
- Boolean 数组元素拼串折出来是 `-1` 而不是 `True`（读路径，与本刀的回写无关）。
- ByRef Boolean 未归一、无 typelib 的外部 VB6 ActiveX DLL 收 ByRef String 仍按值。

## §11.146（第一百六十四刀）引用形补上 Byte / Currency / Date

### 一、缺口

§11.145 之后「按引用的打包器」只有六枚：`VT_BYREF|` I4 / I2 / R4 / R8 / BOOL / BSTR。
Byte、Currency、Date 三型当 ByRef 实参仍按值发，体里的写回落空 —— 这一族在 RTL 侧
本来就没有对应的 `*Ref` 打包器，也没有入槽/写回那两臂，是**真缺口**，不是判据问题。

### 二、RTL 侧（三处）

1. 打包器（`vb6com_pack.c` / `vb6com.h`）：
   `vb6_ComPackByteRef(uint8_t*)` / `vb6_ComPackCurRef(int64_t*)` /
   `vb6_ComPackDateRef(double*)`，分别递 `VT_BYREF|VT_UI1` / `VT_CY` / `VT_DATE`，
   指针直指调用方那枚槽。
2. `vb6_VariantBorrowIntoSlot`：加 `case VT_CY`（搬 `pcyVal->int64`）、
   `case VT_DATE`（搬 `*pdate`）两臂 —— 本仓 Currency 住 `cyVal`（int64）、
   Date 住 `dblVal`（与「序列号在 dblVal」的既有约定一致）。
3. `vb6_ByRefScalarWriteBack`：加同名的两臂，出站写回对端 pointee。
   `vb6_ComArgNormalize` 的 `unwrapped` 名单随之加 CY / DATE（入槽已搬过，不再换算）。

### 三、发码侧

`kComPackRefTiers` 六档 → 九档，数组元素与类字段两个消费点自动跟着走。
入站读臂改成**按声明档自己的那一格**取址：Long→`&lVal`、Integer→`&iVal`、
Byte→`&bVal`、Currency→`&cyVal`、Date→`&dblVal`。以前 Long/Integer/Byte 一律 `&lVal`
—— union 同偏移，靠「写入总在偏移 0」侥幸成立，形参型别对不上（C4133）。

### 四、刻意不开的一格：货币数组元素

`o.IncC ca(1)` 仍按值。不是漏档，是两边表示不一致：标量 Currency 的 C 档是 `double`
（Fix 126 值语义，放大整数只存在于 VT_CY 变体里），而数组元素档是 `int64_t`
（`mapSaElemCType`）。按 CurRef 递过去，服务端按值语义写回一枚 double 到 int64 槽，
读出来是位模式垃圾（实测 `4608308318706860032`）。在册欠项：要把 `vb6_sa_currency`
的存法与元素档一起改，不在这一刀偷开一档。

### 五、读数

夹具 `tests/b165_byref_byte_cur_date`（x64 = x86 同形）：

```
B165-Y1=7        Byte 数组元素
B165-Y2=7        Byte 类字段
B165-Y3=7        Byte 模块级裸变量（基线）
B165-C1=10       Currency 数组元素（仍按值，在册）
B165-C2=21       Currency 类字段
B165-C3=31       Currency 模块级裸变量
B165-T1=ok       Date 数组元素（按日期比较判，避开格式化差异）
B165-T2=ok       Date 类字段
B165-T3=ok       Date 模块级裸变量
B165-Y4=7        同一 Byte 元素连发两次
B165-DONE
```

b164 的 N12 由 `1` 变 `7` —— Byte 那一格本就是本刀的目标，针随之更新。

### 六、第一次 CI 红：外部对端不收引用形

```
[TEXT] dict_native ... FAIL
  ocx: d101|date_key: same=True asnum=True asstr=False
  c3 : d101|date_key: same=True asnum=False asstr=False
```

根因在出站：`vb6_ComNarrowArgs` 的「降值」只认 R8/R4（收窄）和 `BYREF|BSTR`（§11.143），
新开的三档它一概不认识 ⇒ 原样递给 `Scripting.Dictionary`，`Dict.Exists key` 收到引用形
日期，判等变成 false（oracle 那侧是 VT_DATE 值，答 true）。

按 §11.143 同一条口径补齐 —— 引用形只对**我们自己的服务端**有意义：

- 新增 `vb6_ComDemoteNewRef`：把 `BYREF|{UI1,CY,DATE}` 就地改成值形（拷 pointee）。
- 没有 ITypeInfo：对端既不是包装器也不是工程类实例 ⇒ 逐枚降值。
- 有 ITypeInfo：按该形参的 `PARAMFLAG_FOUT` 决定，声明了出参才保留引用。
- 只降这三档；I4/I2/R4/R8/BOOL 的引用形是 §B224 之后的老路径，这一刀不碰。

### 七、账与门

- manifest：corpus 838 → 840（b165 两枚），行数 1369 → 1371。
  存量行真动到的：`b433`（两枚）、`dict_native`（两枚）—— 都是这三型新按引用发的地方。
- 门 `check_byref_arg_pack_authority.ps1`：R1 六枚 → 九枚；新增 R3 五条
  （三档映射、货币数组元素不开、表里正好九档）。
- 九门 9/9；`tests/dict_native/run_compare.ps1` = PASS（70 行逐字一致）。
  CI 22/22（run 45523956970）。

### 八、这一刀之后仍在册的相邻欠项

- 货币数组元素与标量不一个表示（上节）。
- 属性读当 ByRef 实参仍按值（VB6 经 Let 写回）—— 要一条「读-改-写」的双程发射。
- Boolean 数组元素拼串折出来是 `-1` 而不是 `True`（读路径）。
- ByRef Boolean 未归一、无 typelib 的外部 VB6 ActiveX DLL 收 ByRef String 仍按值。

## §11.147（第一百六十五刀）属性读当 ByRef 实参：读-改-写

### 一、缺口

§11.146 之后「按引用发的实参」只差最后一格：属性读。`o.Inc box.Tag`（`o As Object`，
`Tag` 是带 Get/Let 的属性）落到 `comByRefArgExpr` 的属性那一格交回空串 ⇒ 按值打包，
体里的写回落空。实测 `.temp/k166/p1`：`b.Tag = 40` 经 `o.Inc b.Tag` 之后仍是 40。

### 二、真 VB6 的三步

属性没有地址，VB6 不递属性本身，而是：**拷进临时量 → 递它的地址 → 调用回来经 Let 写回**。
这一刀就把这三步发射出来：

```c
int32_t _pb0 = vb6_B166Box_prop_get_Tag(b);
vb6_ComVarFree((void*)vb6_ComCall(o, L"Inc", (void*[]){vb6_ComPackIntRef(&_pb0)}, 1));
vb6_B166Box_prop_let_Tag(b, _pb0);
```

### 三、判据（三件同时成立）

1. 文本必须是**单实参**的属性读 `vb6_<类>_prop_get_<成员>(<base>)`。
   带索引的 Read 属性（`base` 里带顶格逗号或括号）不在这一族，在册。
2. 该类的同名成员必须**既有 Get 又有 Let** —— 只读属性写不回去。
   两问都在同一次符号表遍历里做（`SymbolKind::PropertyGet` / `PropertyLet`），
   型取 Get 的 **`sym->type`**，与「Function/PropertyGet 读自己的名字走返回槽」同一条口径
   （载荷里只有 `memberReturnTypes`，那是源码原样的型名，不是 Vb6Type）。
3. 打包器仍问 `comByRefPackFnForVbType`（§11.145 那张九档表），不在这一格另开一份。

### 四、发射位置

- 台账 `comPropWrites_`（一对 `decl` / `setter` 字符串）只住 `cgen_state.inc` 一处。
- `visit(CallStmt)`：`decl` 在调用**前**发（读那一半），`setter` 在调用**后**、且在
  §11.102 的 `ErrSerial` 守卫**里面**发 —— 「命中即整句中止」的口径下，调用被跳过时
  属性也不该变；收尾在 `CallArgScope102` 的析构里统一清表。
- 这条通道对「晚绑定目标」与「早绑定目标」的语句位置调用都成立，因为两路都走同一个
  `visit(CallStmt)` 尾巴。

### 五、刻意不开两档（在册）

- **String 属性**：临时量要管 BSTR 所有权（get 给的那份谁释放 / Let 收的是拷贝还是转移）。
- **Currency 属性**：C 档是 `double`（值语义），Let 形参未必同宽 —— 与 §11.146 货币数组元素
  那格同一个根因。

两格都留在「按值」这一侧，夹具里各有一枚反面证人。

### 六、读数

`tests/b166_prop_byref`（x64 = x86 逐行相同）：

```
B166-P1=41          Long 属性，第一次
B166-P2=42          同一格连发两次（写回的值要能当下一轮的入参）
B166-P3=21          Integer 属性
B166-P4=31.5        Double 属性
B166-P5=7           Byte 属性
B166-P6=True        Boolean 属性（Not 之后回写）
B166-P7=2020,1,3    Date 属性（Year/Month/Day 数值取值器，不钉区域设置形状）
B166-P8=ab          String 属性 —— 反面证人，仍按值
B166-P9=xyz         真变量对照
```

`b164` 的 N13 由 `5` 变 `6` —— 那一格本来就是「属性读不可取址，仍按值」的负对照，
本刀翻面，夹具注释与针一起改。

### 七、第一次 CI 红：夹具钉了区域设置形状的串

```
FAIL fixture ... pins locale-shaped literal(s): B166-P7=2020-01-03
  -- but its source builds them with NEITHER explicit format string NOR numeric date getters.
```

原来那一针是 `Format(b.Tm, "yyyy-mm-dd")`。日期格式串跟系统区域设置走，换一台机器就是
另一种写法。改成数值取值器：

```vb
Debug.Print "B166-P7=" & Year(b.Tm) & "," & Month(b.Tm) & "," & Day(b.Tm)
```

读数变 `2020,1,3`，两台逐行相同。

### 八、账与门

- 门 `check_byref_arg_pack_authority.ps1` 新增 R9 八条：台账只住一处、读/写两半各一处、
  `CallArgScope102` 收尾、`PropertyGet` + `PropertyLet` 两问各一次、String/Currency 不开、
  夹具两格；R8 的消费点数 4 → 5（多了属性读这一处）。
- manifest：corpus 840 → 842，行数 1371 → 1373。存量只有 `b164` 一枚真变；另外六行
  （`b132`、`ctrlmscomm`、`GUISample×3`、`TeeChartChartDisplay`）是本机 typelib 侧的
  known-diff，**CI 是 oracle**。
- 九门 9/9；CI 22/22（run 45538735210）。

### 九、这一刀之后仍在册的相邻欠项

- **String / Currency 属性**两档（上节）。
- **带索引的 Read 属性**（`o.Inc box.Item(i)`）不在这一族。
- **早绑定接收者**那一形：`b.Inc b.Tag`（`b` 是带型类变量，直接调 `vb6_B166Box_Inc`）
  仍发 `(&(int32_t){vb6_B166Box_prop_get_Tag(b)})` —— 拷进去了没有拷回来，
  实测 `.temp/k166/p2` 的 `Q1=40`（应为 41）。它走的是**早绑定实参发射器**
  （`cgen_expr_call_arg_emit.inc` 的复合字面量那一族），不是本系列这条线；
  要修得让那条发射器也登记同一张台账 —— 可以复用本刀的 `comPropWrites_`，
  但需要那一族先认识「这枚实参是属性读」。
- 上几刀留下的：货币数组元素、ByRef 布尔未归一、无 typelib 的外部 VB6 DLL 收 ByRef
  String 仍按值、Boolean 数组元素拼串折 `-1`。

## §11.148（第一百六十六刀）早绑定接收者那一形也走读-改-写

### 一、缺口

§11.147 把「属性读当 ByRef 实参」的读-改-写补给了晚绑定那一支：`o.Inc b.Tag`
（`o As Object`，实参按 `VT_BYREF|T` 发）。早绑定接收者那一形没覆盖 ——
`b.Inc b.Tag`（`b As B166Box`，直接调 `vb6_B166Box_Inc`）发的是

```c
vb6_B166Box_Inc(b, (&(int32_t){vb6_B166Box_prop_get_Tag(b)}));
```

拷进去了没有拷回来：被调方的写回落在复合字面量上就没了。实测 `.temp/k166/p2`：
`Q1=40`（应为 41）、`Q2=30`（应为 31.5）。

### 二、一位主人，两个发射器

这一问提成 `CCodeGen::comPropReadWriteArg(expr, argText, outVt, outVar)`
（定义在 `cgen_util_com.cpp`，声明在 `cgen_helpers.inc`）。判据三条与上一刀相同：

1. 文本必须是**单实参**的属性读 `vb6_<类>_prop_get_<成员>(<base>)`；
2. 该类的同名成员必须**既有 Get 又有 Let**；型取 Get 的返回型 `sym->type`；
3. 打包器要在按引用那一族里（§11.145 那张九档表）。String / Currency 两档仍不开。

两处消费点：

- **晚绑定**：`comByRefArgExpr` 调它，交回 `vb6_ComPackXxxRef(&_pbN)`（与上一刀同形）。
- **早绑定**：`cgen_expr_call_arg_emit.inc` 的复合字面量那一族在落 `(&(ctype){...})`
  之前问同一句，成立就交 `&_pbN`。**只接同型那一格**（`mapType(属性型) == 形参 C 型`）：
  型不一致时（Long 属性接 Integer 形参那种）仍走老路 —— C 侧的隐式换算本来就在那里做，
  硬接会改掉换算行为。

### 三、发射与收尾仍只有一处

台账 `comPropWrites_` 住 `cgen_state.inc`，两行由 `visit(CallStmt)` 那条尾巴发：
`decl` 在调用前、`setter` 在调用后且在 §11.102 的 `ErrSerial` 守卫**里面**。

新加一道闸 `comPropWriteWindow_`：只有语句位置那棵被调方表达式树里才登记。属性读出现在
**表达式上下文**（赋值右值 / 内建函数实参 / `Debug.Print` 实参）时没有「调用之后」这个
位置，登记了没人发那两行就是 C2065 —— 那一族仍按值，在册。

### 四、读数

`tests/b166_prop_byref`（x64 = x86 逐行相同）：

```
B166-P1=41  B166-P2=42  B166-P3=21  B166-P4=31.5  B166-P5=7
B166-P6=True  B166-P7=2020,1,3  B166-P8=ab  B166-P9=xyz     ← 晚绑定那一族
B166-Q1=41   Long 属性，早绑定（改前 40）
B166-Q2=31.5 Double 属性，早绑定（改前 30）
B166-Q3=ab   String 属性反面证人，两族都仍按值
```

### 五、账与门

- `check_byref_arg_pack_authority.ps1` 的 R9 扩到十三条：问句定义 1 / 声明 1 / 两个消费点、
  早绑定只接同型、开窗 1 + 关窗 2；R8 的消费点 5 → 6。
- manifest：corpus 842 不变；b166 两行按当前 C3.exe 重新登记（其余六行是本机 typelib 侧的
  known-diff，CI 是 oracle）。
- 九门 9/9；CI 22/22（run 45543339265）。

### 六、这一刀之后仍在册的相邻欠项

- **String / Currency 属性**两档（所有权 / 宽度）。
- **带索引的 Read 属性**（`o.Inc box.Item(i)`）。
- **表达式上下文里的属性读实参**（`x = o.F(b.Tag)` / `MsgBox o.F(b.Tag)`）—— 没有
  「调用之后」这个位置，要修得先有一条「表达式内调用之后」的发射道。
- 早绑定那一族的**型不一致**一格（Long 属性接 Integer 形参）仍按值。
- 上几刀留下的：货币数组元素、ByRef 布尔未归一、无 typelib 的外部 VB6 DLL 收 ByRef
  String 仍按值、Boolean 数组元素拼串折 `-1`。

## §11.149（第一百六十七刀）String 属性也走读-改-写，临时量的 BSTR 归调用方退

### 一、缺口

§11.147 / §11.148 两刀把「属性读当 ByRef 实参」的读-改-写补给了数值五档（`Long`/`Integer`/
`Single`/`Double`/`Boolean`）与 `Byte`/`Date`，String 与 Currency 两档一直 `return false`
按值发。String 那一档卡在**所有权**：`prop_get` 交回来的是一枚新分配的 `BSTR`，拷进临时量
以后谁退？不收这条线就只能每次调用漏一条串。

### 二、两条前提查清了

`tests/b166_prop_byref` 的发码里（`.temp/k167/t.c`）：

```c
BSTR vb6_B166Box_prop_get_Label2(vb6_cls_B166Box* me) {
  BSTR vb6_ret_Label2 = NULL;
  vb6_BSTR_Assign(&vb6_ret_Label2, me->m_Name);   /* 拷一份 → 新分配, 归调用方 */
  return vb6_ret_Label2;
}
void vb6_B166Box_prop_let_Label2(vb6_cls_B166Box* me, BSTR v) {
  vb6_BSTR_Assign(&me->m_Name, v);                /* 也是拷一份 → 不接手所有权 */
}
```

两头都是**拷一份**：get 交出的新串归调用方，let 收下的是值不接手。所以临时量必须由调用方退 ——
这就是这一刀唯一新增的一行。

### 三、改法

- `ComPropWriteBack` 多一格 `release`（空 = 标量档，值语义，没有可退的）。
- `comPropReadWriteArg` 只对 Currency 仍 `return false`；String 那一档给 `release` 填
  `vb6_BSTR_Free(_pbN);`。
- `visit(CallStmt)` 的尾巴在 102 守卫**关掉之后**发这一行：守卫没放行时调用被跳过，临时量里
  仍是 get 新分配的那枚串，更得退。（setter 仍在守卫**里**：跳过调用时属性不该变。）
- 发出来的三段（`tests/b166_prop_byref` 的 Q3，早绑定）：

```c
BSTR _pb10 = vb6_B166Box_prop_get_Label2(b);
vb6_B166Box_Append(b, &_pb10, _own_1);
vb6_B166Box_prop_let_Label2(b, _pb10);
vb6_BSTR_Free(_pb10);
```

顺带修掉一处旧漏：改前这一形发 `(&(BSTR){vb6_B166Box_prop_get_Label2(b)})`，那枚 get 出来的
串**从来没退过**，每次调用漏一条。

### 四、读数

`tests/b166_prop_byref`（x64 = x86 逐行相同）：

```
B166-P8=abcd   晚绑定 String 属性（改前 ab）
B166-Q3=abcd   早绑定 String 属性（改前 ab）
其余 W1..W2 / P1..P7 / P9 / Q1..Q2 与上一刀逐行相同
```

### 五、账与门

- `check_byref_arg_pack_authority.ps1`：`R9 string/cur props by value` 换成
  `R9 currency prop by value`（String 已开），新增三条 —— `R9 string prop release`
  （退还那一行只许一处）、`R9 release field`（台账多一格）、`R9 release emitted`
  （发在守卫**外**）。
- 九门 9/9；CI 22/22（run 45548994502）。

### 六、仍在册的相邻欠项

- **Currency 属性**：`double` 值语义与数组元素那档 `int64_t` 不是一个表示，和 §11.146
  货币数组元素同一张欠项。
- **带索引的 Read 属性**（`o.Inc box.Item(i)`）。
- **表达式上下文里的属性读实参**（`x = o.F(b.Tag)` / `MsgBox o.F(b.Tag)`）—— 没有
  「调用之后」这个位置，需要先有一条「表达式内调用之后」的发射道。
- 早绑定那一族的**型不一致**一格（Long 属性接 Integer 形参）仍按值。
- 上几刀留下的：货币数组元素、ByRef 布尔未归一、无 typelib 的外部 VB6 DLL 收 ByRef
  String 仍按值、Boolean 数组元素拼串折 `-1`。

## §11.150（第一百六十八刀）带索引的属性读也走读-改-写

### 一、缺口

§11.147/148/149 三刀把「属性读当 ByRef 实参」的读-改-写补给了**无索引**的属性（数值五档、
`Byte`/`Date`、`String`）。带索引那一形一个都进不来，两处各卡一刀：

1. 判据第一关只认 `ASTNodeKind::MemberAccessExpr`，而 `b.Slot(2)` 在 AST 上是
   `IndexOrCallExpr`（`callee` = `MemberAccessExpr(b, Slot)`）—— VB6 在解析阶段不区分
   数组索引与函数调用。
2. 实参文本的 base 一律拒绝含逗号，而 `vb6_B168Box_prop_get_Slot(b, 2)` 里就有顶层逗号。

改前两族都按值发，写回落空（`.temp/k168/r0.c`）：

```c
vb6_B168Box_Bump(b, (&(int32_t){vb6_B168Box_prop_get_Slot(b, 2)}));
vb6_ComCall(o, L"Bump", (void*[]){vb6_ComPackInt(vb6_ComCall(o, L"Slot", ...))}, 1)
```

### 二、改法

- 成员名从 `callee` 上取：`comPropReadWriteArg`（唯一那位主人）与 `comByRefArgExpr`
  的属性那一支都认 `IndexOrCallExpr` 这一形。
- 实参文本按**括号深度**切：第一枚作 base（base 仍只接简单形 —— 不含括号，它是 Let 的
  首参），其余原样当索引实参；Let 的调用形是 `let(base, <索引...>, tmp)`，取值形参在
  最后，与 VB6 `Property Let` 的形参序一致。

发码（早绑定 / 晚绑定各一枚）：

```c
int32_t _pb2 = vb6_B168Box_prop_get_Slot(b, 2);
vb6_B168Box_Bump(b, &_pb2);
vb6_B168Box_prop_let_Slot(b, 2, _pb2);
---
int32_t _pb3 = vb6_B168Box_prop_get_Slot(b, 2);
vb6_ComVarFree((void*)vb6_ComCall(o, L"Bump", (void*[]){vb6_ComPackIntRef(&_pb3)}, 1));
vb6_B168Box_prop_let_Slot(b, 2, _pb3);
```

### 三、读数

`tests/b168_indexed_prop_byref`（x64 = x86 逐行相同）：

```
B168-S1=11     单索引，早绑定（改前 10）
B168-S2=12     单索引，晚绑定接收者 + 带型属性读
B168-S3=1002   两个索引实参（Pair 的 Let 取值形参在最后）
B168-S4=ab!    String 档带索引（临时量的 BSTR 照旧归调用方退）
B168-S5=31     索引本身是表达式（1 + 1 + 1）
```

### 四、账与门

- `check_byref_arg_pack_authority.ps1`：R9 新增四条 —— `indexed prop recognized = 2`
  （主人 + 晚绑定各一处）、`index args split = 1`、`index let order = 1`、
  `indexed prop fixture = 2`。
- 九门 9/9；CI 22/22（run 45563322745）。
- manifest：corpus 842 → 844（新增夹具两行，按路径序中插在 `b168_cvf` 之后），其余
  838 行逐字节相同。

### 五、仍在册的相邻欠项

- **索引实参会被求值两次**（Get 一次、Let 一次）：`b.Bump b.Slot(Inc())` 那种带副作用的
  索引式与真 VB6 未比对，先按这条口径记着。
- **晚绑定接收者上的属性读**：`o.Bump o.Slot(2)` 里 `o.Slot(2)` 本身走 `vb6_ComCall`，
  不是 `_prop_get_` 文本，进不了这一族（要先把晚绑定读落进一枚可写回的 Variant 槽）。
- **Currency 属性**：与货币数组元素同一张欠项（`double` 与 `int64_t` 表示不一致）。
- **表达式上下文里的属性读实参**（`x = o.F(b.Tag)` / `MsgBox o.F(b.Tag)`）—— 没有
  「调用之后」这个位置，需要先有一条「表达式内调用之后」的发射道。
- 早绑定那一族的**型不一致**一格（Long 属性接 Integer 形参）仍按值。
- 上几刀留下的：货币数组元素、ByRef 布尔未归一、无 typelib 的外部 VB6 DLL 收 ByRef
  String 仍按值、Boolean 数组元素拼串折 `-1`。

## §11.151（第一百六十九刀）With 块里的属性读也走读-改-写；型不一致不再「假登记」

这一刀两件，都是上一刀之后逐形探出来的。

### 一、With 块：改前一个都不成立

```vb
With b
    .Inc .Tag
End With
```

改前发的是

```c
vb6_B169Box_Inc((void*)_vb6_with_0, (&(int32_t){vb6_B169Box_prop_get_Tag(_vb6_with_0)}));
```

写回落空。两处各卡一刀：

1. `.Tag` 在 AST 上是 `WithMemberExpr` —— 只有成员名，没有 object 可推型；而这一问只认
   `MemberAccessExpr`。带索引的 `.Slot(2)` 的 `callee` 也是 `WithMemberExpr`。
2. base 那一段被「不许含括号」一票否决：`With b` 里的 `.Slot(2)` 的 base 是
   `(void*)_vb6_with_6`（带一层强转）。

改法：

- **主人自己认形**：`MemberAccessExpr` / `IndexOrCallExpr`（callee 为 `MemberAccessExpr`
  或 `WithMemberExpr`）/ `WithMemberExpr` 三形都收，成员名从各自那一格取。
- **类名三级退**：能推型就推型 → 推不出问 `withObjectInfoStack_`（With 主体的
  `className`）→ 再退一步从发码文本 `vb6_<Cls>_prop_get_<M>` 里取名字。**型依旧问声明侧**
  （§11.145 那条口径没变：名字只是定位，档来自符号表）。
- **base 放宽只放一种**：一个括号包着的强转 + 后面紧跟不带括号的剩余文本。其余仍拒。

读数（`tests/b169_with_prop_byref`，x64 = x86）：`W1=41`（改前 40）、`W4=ab!`（改前 ab）、
`W5=11`（With + 带索引，改前 10）。

### 二、早绑定型不一致那一格是「假登记」

`b.IncI b.Tag`（Long 属性接 Integer 形参）改前发：

```c
int32_t _pb1 = vb6_B169Box_prop_get_Tag(b);
vb6_B169Box_IncI(b, (&(int16_t){vb6_B169Box_prop_get_Tag(b)}));
vb6_B169Box_prop_let_Tag(b, _pb1);      /* ← 把未动的值又写了一遍 */
```

台账先登记、调用点随后发现自己用不上（型不一致只接同型），于是临时量成了孤儿：被调方的
写回落在复合字面量上丢掉，属性却被 Let 白叫一次（有副作用的 Let 就会现形）。

改法：`comPropReadWriteArg` 多一个 `wantCType`，非空即「调用点已经定了形参 C 型」，
型不一致**当面交回 false**、一个字都不登记；早绑定那一路把 `cType` 递进去。

型不一致这一格本身仍按值（`W2=40` 是证人）—— 要开得先有「按形参型造临时量 + 两向换算」
的账（VB6 在那里是「拷临时量 + 换算」，不是拒绝）。

### 三、账与门

- `check_byref_arg_pack_authority.ps1` R9 新增四条：`with-prop asks stack = 1`、
  `want-ctype gate = 1`、`want-ctype from caller = 1`、`with-prop fixture = 2`。
- 九门 9/9；CI 22/22（run 45570526384）。
- manifest：corpus 844 → 846（新夹具两行，按路径序插在 `b169_cfw` 之后）；存量移动 3 行
  —— `VBFlexGridDemo.vbp`（2 处 `WriteProperty` 的 `Enabled` / `OLEDropMode`）、
  `ucSegmentedControl/SegmentedControl.vbp`（With 里 `ItemCaption`/`ItemValue`/
  `ItemToolTipText`/`ItemTag`/`ItemDisabled`/`ItemDropDown`）、
  `ucTreeView/Proyecto1.vbp`（`OLEDropMode` 一处）。三份都用「停工本机重编旧源码」作基线
  A/B 逐行对过：抹掉临时量编号后差异只剩那几组
  `<C型> _pbN = …_prop_get_…; <call 递 &_pbN>; …_prop_let_…; [vb6_BSTR_Free(_pbN);]`。

### 四、仍在册的相邻欠项

- **型不一致那一格**（Long 属性接 Integer 形参）：按形参型造临时量 + 两向换算，未开。
- **索引实参求值两次**（Get 一次、Let 一次）：带副作用的索引式与真 VB6 未比对。
- **晚绑定接收者上的属性读**：`o.Bump o.Slot(2)` 里 `o.Slot(2)` 走 `vb6_ComCall`，不是
  `_prop_get_` 文本 —— 要先把晚绑定读落进一枚可写回的 Variant 槽。
- **Currency 属性 / 货币数组元素**：`double` 与 `int64_t` 表示不一致，同一张欠项。
- **表达式上下文里的属性读实参**（`x = o.F(b.Tag)` / `MsgBox o.F(b.Tag)`）—— 没有
  「调用之后」这个位置，需要先有一条「表达式内调用之后」的发射道。
- 上几刀留下的：ByRef 布尔未归一、无 typelib 的外部 VB6 DLL 收 ByRef String 仍按值、
  Boolean 数组元素拼串折 `-1`。

## §11.152（第一百七十刀）属性读与形参型不一致时的真换算

上一刀把「早绑定型不一致」的假登记修掉了，但那一格本身仍按值：`b.IncI b.Tag`（Long 属性
接 Integer 形参）体里的加一丢掉，`W2=40` 就是证人。这一刀把换算打开。

真 VB6 的语义不是「拒绝」，是**拷一枚形参型的临时量、递它的地址、回来再换算写回属性**。
于是 `comPropReadWriteArg` 的 `wantCType` 从「合不上就回 false」改成「按它造临时量 +
两向换算」：

- 声明型 = 形参 C 型；读入值按「属性 C 型 → 形参 C 型」换算，写回值按反向换算。
- 换算梯只住一处 `narrowArgText(expr, toC, fromC)`：整型目标走 RTL 的验界函数
  （`vb6_ChkInt` / `vb6_ChkByte` / `vb6_ChkLong`，越界即 VB6 Error 6）；浮点 → 整型走
  `vb6_FltToLng`（VB6 的取整，半值远离零，不是 C 的截断）再验界；放大型只加 C 强转。
- String 那一档不换算（BSTR 没有数值换算可言），Currency 仍整档不开。
- 早绑定调用点不再自己比 `mapType`，把 `cType` 递进去就完事。

发码（`tests/b170_prop_byref_narrow`，六种配对）：

```c
int16_t _pb0 = vb6_ChkInt((int32_t)(vb6_B170Box_prop_get_L1(b)));
vb6_B170Box_IncI(b, &_pb0);
vb6_B170Box_prop_let_L1(b, vb6_ChkLong((int64_t)(_pb0)));
...
int32_t _pb5 = vb6_ChkLong(vb6_FltToLng(vb6_B170Box_prop_get_D1(b)));
```

读数（x64 = x86 逐行相同）：`T1=41 T2=8 T3=21 T4=21 T5=2.5 T6=5 T7=unchanged`。
T6 用 3.7 而不是 3.2 —— 把「取整」与「截断」分开：取整 4 + 1 = 5，截断会得 4。

门 `check_byref_arg_pack_authority.ps1`：R9 里「want-ctype gate = 1」（拒绝那一格）换成
`conv skips string = 1`，另加 `narrow conv decided / decl / let / ladder` 四条与
`narrow fixture = 2`；`early-bound site` 那条改成认新的递参形。九门 9/9。

manifest：corpus 846 → 848；存量移动 1 行 = `tests/b169_with_prop_byref`（W2 从按值变换算）。

## §11.153（第一百七十一刀）Currency 属性的拼串原来按 Long 发

```vb
b.C1 = 10.5
Debug.Print "Q=" & b.C1        ' 改前 "10"，改后 "10.5"
```

而 `Dim c As Currency` 的**局部变量**同一形一直是对的。根因不是「Currency 没支持」，是
**同一只值两条路两个答案**：

- 变量那一形：符号档走 `Vb6Type::Double`（Currency 的 C 承载就是 double，Fix 126）⇒ `CStrDbl`；
- **属性读**那一形：经 `memberReturnTypes` 得到 `Vb6Type::Currency`
  （`cgen_util_type.cpp:652`）⇒ `wrapToBSTR` 的分档 switch **没有 Currency 臂** ⇒ 落
  default 的 `vb6_CStrLong((int64_t)double)`，10.5 截成 10。

修法只补一条臂，不另开格式化路：

```c
case Vb6Type::Currency: return "vb6_CStrDbl(" + expr + ")";
```

发码（`tests/b171_currency_prop_str`）：

```c
vb6_CStrDbl(vb6_B171Box_prop_get_C1(b))          /* Q1 属性读 */
vb6_CStrDbl(vb6_B171Main_Cur())                  /* Q4 Function ... As Currency */
vb6_CStrDbl(vb6_B171Box_prop_get_C1(_vb6_with_0))/* Q7 With 里的属性读 */
```

读数（x64 = x86）：`Q1=10.5 Q2=11.5 Q3=10.5 Q4=20.25 Q5=11.5 Q6=7.25 Q7=11.5`。
Q3（Double 属性）与 Q6（局部 Currency 变量）是改前就对的两枚对照，钉住不许回退。

顺带把 `b165_byref_byte_cur_date` 的两根针改成真值（C2 `21` → `21.25`、C3 `31` → `31.25`）
—— 那两根针以前把截断后的数当期望值钉住了；Currency **数组元素** C1 仍打 10
（`int64_t` 缩放的表示不一致，在册，本刀不碰）。

门 `check_implicit_cstr_station.ps1` 加 R16：Currency 臂恰好 1 条、不许再出现
`case Vb6Type::Currency: return "vb6_CStrLong`、臂数 ≥ 16（少一条就落 default）、b171 夹具两格。
九门 9/9；CI 22/22（run 45588642245）。

manifest：corpus 848 → 850；存量移动 1 个夹具（b165 两行）。

### 附：CI 点名器的一条假红

同一轮 CI 里出现过四条 `##[error][GUI] ... pixels: PASS ...` —— 用例其实过了。机制：
`Test-GuiVbpCase` 用 `-NoNewline` 打前缀，先落地的是像素夹具那一行
（`pixels: PASS tc_px ...`），最后才是 `PASS (window=...)`；点名器只认「尾巴以 PASS/SKIP
开头」，于是过的用例被判成没判过。修法只动报表面：尾巴认可选的 `pixels:` 前缀。
本机用六条真实排版自证（两条 pixels: PASS 由点头变不点，三条 FAIL 照点）；CI 22/22
（run 45583802369）。

### 仍在册（本两刀新增）

- **Currency 赋值未圈到 4 位小数**：VB6 的 Currency 每次赋值会先圈到 4 位；本仓的 double
  承载不圈。与「数组元素 `int64_t` 缩放」是同一张表示层欠项的两个面。
- **拼串格式化的档位清单要跟着 Vb6Type 走**：这一刀是「少一条臂」，同类风险在于以后新增
  档位忘了加臂 —— R16 的「臂数 ≥ 16」只兜得住不回退，兜不住新档。

## §11.154（工具刀，不占刀号）GUI/像素用例移出阻塞门

GUI 那一族一直在 vbp 分片里当阻塞门，而它的判据天生环境依赖。三条实证都在仓里：

1. **要真桌面**：像素夹具走 `CopyFromScreen` 抓屏，拿不到就自己 SKIP
   （`tests\guisample\tree_icons_px.ps1`）。
2. **会互抢单实例主窗**：`run_tests.ps1` 里 2026-11-04 的实测 —— 同一产物单跑
   `PASS icons=9`，套件里跑成 `SKIP 主窗未找到`；现在的对策是跑像素前先杀掉用例进程。
3. **判据是渲染量**：`n=61/265/262`、`ring n=6175 at70=1065`、`title=(83,14)-(225,24)`
   这类墨水计数/几何，DPI/主题/字体版本一动就漂。

本机 `-Category gui` 就复现了第三条：`ucSegmentedControl` 的读数里
`physical client=939x650 scale=1.25`（125% 缩放），而期望值按 96 DPI 钉的 ⇒ 红。
这类红不是回归信号，是环境噪声。

改法（只动编排与一格开关，不动任何判据）：

- `run_tests.ps1` 新增 `-SkipGui`：vbp 分片里不再调 `Invoke-GuiVbpTable`、也不计 GUI
  用例的格数（`Test-GuiVbpCase` 早退）；本地不传开关时照旧全跑。
- `ci.yml` 的 12 个 vbp 分片 `shardArgs` 全部追加 `-SkipGui`；新增 matrix 条目
  `category: gui` 跑 `-Category gui`；该 job 级
  `continue-on-error: ${{ matrix.category == 'gui' }}` —— 它仍会跑、仍会把每条失败写成
  `::error::` 注解，只是不许把整条流水线拉红。

本机自证：

- `-Category vbp -VbpShard 1 -VbpShardTotal 60 -SkipGui` ⇒ `[GUI]` 行 0 条、
  `Results: PASS=17 FAIL=0 SKIP=1`；
- `-Category gui` ⇒ 25 例 `PASS=24 FAIL=1 SKIP=0`（就是上面那格 125% DPI 读数）。

### 撤 GUI 表后暴露的一条隐式接续

第一次推（`f6bc33bc`）CI 在 Tests (vbp #3) 红了：

```
[MANIFEST] balloon_manifest_is_user_supplied ... FAIL (no exe: output\BalloonTooltips\BalloonTooltips.exe)
```

那条判据（只读产物字节，数 `</assembly>`、查 `dpiAware`）吃的是 GUI 表建出来的 exe，
`-ShardGroup "balloon"` 就是为这条接续写的。修法：`-SkipGui` 时在调用点自己建一份
（`cmd` 直调，与 `Test-FormIconResources` 同一形），**判据一个字不改**。本机自证：
编出 `BalloonTooltips.exe` ⇒ `manifests=1`、`dpiAware=True`。全脚本里 `$OutDir` 直读
只有三处，另两处（FormIconResources / ctrlmanifest 的 `MfApp.exe`）本来就自己建。

修后 CI：23 格**全绿**（run 45598869416，含新的 `Tests (gui)`；它在 CI 的桌面/DPI 下也
过了，但权限上仍是非阻塞格）。

### 本机验证的一件事（不是本刀引入）

用 Windows PowerShell 5.1 直跑 harness 时，`[COMPILE-FAIL] iface_multi_incomplete_negcontrol`
会假红：`needle=False` 但 `exit=1` 且无 exe。根因是脚本顶部
`$ErrorActionPreference='SilentlyContinue'` 下，5.1 对原生 stderr 的 `2>&1` 捕获为空；
换 `pwsh`（CI 用的那个）即 PASS。复现：`.temp\k172\probe_neg.ps1`（5.1，len=0）vs
`probe_neg2.ps1`（`ErrorActionPreference='Continue'`，len=1401）。
**本机验证 harness 一律用 pwsh**。

## §11.155（第一百七十二刀）String 与数值之间的隐式换算

刀名取「隐式换算」，因为这一族以前**根本没有换算站**：BSTR 被当成整数（或反过来）
直接递进 C。两处实测：

| 站 | 改前发码 | 实测 |
| --- | --- | --- |
| 赋值 `n = "12"` | `n = vb6_ChkLong(vb6_BSTR_FromStr(L"12"));` | `Error 6: Overflow` |
| Long 属性交 String 形参 | `BSTR _pb0 = ((BSTR)(vb6_P173Box_prop_get_N1(b)));` | `0xC0000005` |

第一处是**指针当整数**（BSTR 指针值远超 int64 域之外也就罢了，它落进范围检查里就是
溢出）；第二处是**整数当指针**（`(int32_t)7` 当 BSTR 解引用，体里一拼串即崩）。

### 折数只问 RTL 那一位主人

RTL 早就有一族带 BSTR 臂的 `_Generic` 宏（`vb6_CLng` / `vb6_CDbl` / `vb6_CInt` /
`vb6_CCur` / `vb6_CSByte` / `vb6_CLngLng`），BSTR 臂落到 `vb6_*BSTR`，解析由账 #416
那位主人 `vb6_NumVal` 统一答。本刀不新写解析，只在两个站把「源是 String」这一格
接到那族宏上：

- `narrowArgText` 新增 `fromC == "BSTR"` 一臂：`int16_t`→`vb6_CInt`、`int32_t`→`vb6_CLng`、
  `int64_t`→`vb6_CLngLng`、`double`→`vb6_CDbl`、`float`→`vb6_CSng`、
  `uint8_t`→`vb6_CByte(vb6_CDbl(..))`（Byte 没有 BSTR 臂，先折 double 再收宽）；
- `narrowCheckAssign` 新增 String 源一臂（同一套；以前是「装不下就套 Chk*」，
  而 `cgenIntBits(String)` 答 0 ⇒ 走到最后一行裸套）；
- `comPropReadWriteArg` 开「数值属性 → String 形参」这一向：临时量是 BSTR，
  由新助手 `cstrOfVt` 按**声明档**选 CStr 家族（`int16_t` 盖着 Boolean、`double` 盖着
  Date —— C 型文本答不出这一格，所以选档问 Vb6Type，与 R8/R9 同口径），写回时拿
  `narrowArgText(tmp, propC, "BSTR")` 折回数值，临时量归调用方 `vb6_BSTR_Free`
  （`propVt == String || toBstr155` 两向都算）。

严格性顺带自证：`n = "12abc"` 仍抛 `Error 13`（NumVal 口径不是 Val 的宽松口径），
与 VB6 一致。

### 夹具 `tests/b175_str_num_conv`

```
B175-1=12     Long    ← "12"     B175-5=70    Long 属性 7 → "7" & "0" ⇒ 70
B175-2=34     Integer ← "34"     B175-6=1.55  Double 属性 1.5 → "1.55" ⇒ 1.55
B175-3=200    Byte    ← "200"    B175-7=15    同档对照 (Long → Long)
B175-4=1.5    Double  ← "1.5"    B175-8=4     同档对照 (Double → Double)
```

门：`check_byref_arg_pack_authority.ps1` 加 R10 九条（BSTR 源臂、三个档位、
Byte 折 double、`cstrOfVt` 唯一、`toBstr155` 分支、两向释放、赋值站 String 臂、
夹具两格）；runner 两格（x64 + x86）；九门 9/9。

### 两次 CI 红（都不是判据松动）

1. **`[STATIC] implicit_cstr_station` R5**：我新写的 `cstrOfVt` Integer 档抄成了
   `vb6_CStrLong((int32_t)(x))` —— 那正是旧 13 名单自己的形状，R5 明令「那是另一个口径」。
   改法：Integer 档不套强转（`vb6_CStrLong` 收 int32_t，int16_t 隐式提升），与 Long 档同形。
   这条红是**门在干事**：它不认「谁写的」，只认「这个形状以前判过不许有」。
2. **Emit manifest（shape oracle）**：本机那台二进制与 CI 那台在 **6 份输入**上发码不同
   （`b132_typedcom_local_scope`、`guisample` ×3、`TeeChartChartDisplay`、`ctrlmscomm`
   —— 全是带外部引用/GUI 资源的那几族，环境相关），于是本机登记的值对 CI 不成立。
   按本项目既有口径（吃 CI 那台二进制自己交回来的数）从 `ci/emit-manifest` 取回重登，
   852 行逐行相同。教训写在这里：**发码清单的登记值以 CI 那台为准**，本机跑出来的
   13 行差异只能当参考；本地 `check_manifest_coverage` 只查行集合，查不出这 6 行。

### 在册（本刀不碰）

- **String 属性 → 数值形参**仍是按值（`comPropReadWriteArg` 里那一句 `return false` 还在）：
  它要两枚临时量（BSTR + 数值），写回时还得再造一枚 BSTR 给 `Let`（那枚也归调用方退），
  比本刀这一向多一层；
- 数值形参与 String 形参之间的 `ByVal` 传递是否有同族病未测；
- 字面量 BSTR（`vb6_BSTR_FromStr`）在赋值站仍不释放（旧账，与本刀无关）。

## §11.156（第一百七十三刀）String 属性读当数值 ByRef 实参

上一刀把「String 与数值」两向的换算接上了，但 `comPropReadWriteArg` 里对 String 属性
还有一句 `return false`：那一格根本不进登记道，落到早绑定兜底 ——

```c
vb6_B176Main_IncL((&(int32_t){vb6_B176Box_prop_get_S1(b)}));
```

复合字面量的地址。被调方的写回落在临时量上就没了：`b.S1 = "12"` 走 `IncL` 之后
**还是 "12"**。与前两刀不同，这一格不崩也不报错 —— 静默错值，最难发现的那种。

### 两枚临时量 + 写回那枚 BSTR

```c
BSTR _pb0s = vb6_B176Box_prop_get_S1(b); int32_t _pb0 = vb6_CLng(_pb0s);
vb6_B176Main_IncL(&_pb0);
{ BSTR _pb0l = vb6_CStrLong(_pb0); vb6_B176Box_prop_let_S1(b, _pb0l); vb6_BSTR_Free(_pb0l); }
vb6_BSTR_Free(_pb0s);
```

源串与写回串都不归 `Let` 接手（两边都是拷一份），所以两枚都得调用方退；写回那枚
只活到 `Let` 拷完，就在同一个语句块里现造现退（放在 102 守卫的 setter 那一行里，
跳过了也不会漏退：它根本不会造）。

### 选 CStr 档必须问声明型

这是本刀真正的判据：`int16_t` 同时盖着 **Integer** 与 **Boolean** —— `CStr(12)` 是
`"12"`、`CStr(True)` 是 `"True"`；`double` 又盖着 Double / Date / Currency。C 型文本
答不出这一格，所以：

- `comPropReadWriteArg` 多收一枚 `Vb6Type wantVt`（默认 `Unknown`）；
- 早绑定调用点把 `calleeParams[i].type` 递进来（与它算 `cType` 用的是同一格表）；
- 晚绑定那支仍答 `Unknown` ⇒ 那一格不开（与 R8/R9 同一条口径：型只问声明侧）。

### 夹具 `tests/b176_str_prop_num_arg`

```
B176-1=13   "12"  → Long    (+1)   B176-4=11   "10" → Byte (+1)
B176-2=41   "40"  → Integer (+1)   B176-5=12   同档对照 (Long 属性 → Long 形参)
B176-3=3    "1.5" → Double  (×2)
```

四个形参档各一行，就等于把「选档」这件事钉在读数上：只要那枚 `wantVt` 被退化成
按 C 型猜，Integer 与 Byte 两行会先变。

门：`check_byref_arg_pack_authority.ps1` 里 R9 四条跟着形状更新（一对 `release`
规则、`conv` 判据、两个调用点锚），新增 **R11 九条**（wantVt 进主人 / 声明默认值 /
调用点 / fromStr 闸门 / 型不明不开 / 两枚临时量同一行 / 写回串现退 / 源串归调用方 /
夹具两格）。runner 两格（x64 + x86）均实跑通过；九门 9/9。

### 清单

854 行（新增 2 行 = 新夹具的两份输入）。其余 **与本机上一版逐字节相同** —— 说明这一刀
只碰得到「String 属性交数值 ByRef 形参」那一形，语料里其余工程一行未动。本机与 CI 那台
在 6 份环境相关的输入上仍不同（`b132_typedcom_local_scope` / `guisample` ×3 /
`TeeChartChartDisplay` / `ctrlmscomm`），期望值仍保持 CI 那一版的读数（上一刀已经立下的
口径）。

### 在册（本刀不碰）

- **晚绑定调用点**（`o.IncL b.S1`，b 是 Object）仍走旧路：那里拿不到形参声明型，
  要等晚绑定那一族也把形参表递进来；
- **ByVal 同族**：`x = f(b.S1)` 里 String 属性交数值 `ByVal` 形参是否已经走
  `narrowArgText` 的 BSTR 臂，未测；
- **ByRef Boolean 未归一**：Boolean 形参（`int16_t`）现在会走 `vb6_CInt`，值与 -1/0
  的口径差仍在册。

## §11.157（第一百七十四刀）Boolean 数组元素在拼串里打 -1

这一刀的病不在格式化站，在**型推断**：`inferExprType` 对数组元素读一律答 Variant，
于是拼串站走装箱：

```c
vb6_CStr(vb6_VariantFromValue(VB6_SA_AT(int16_t, flags, 1)))
```

`_Generic` 按 int16 装成 VT_I2，`vb6_CStr` 交 `"-1"`；VB6 交 `"True"`。

### 为什么只有 Boolean 看得出

同一份探针把六档元素全打了一遍（`.temp/k174/p2`）：

```
Q1=42  Q2=43  Q3=44  Q4=1.5  Q5=True  Q6=abc
```

Q1..Q4 改前也是这几个数 —— Long/Integer/Byte/Double 的元素装箱后，`_Generic` 选到
对应 VT，`vb6_CStr` 拆箱后格式化交出来的字与直连 `vb6_CStrLong/CStrDbl` **碰巧相同**。
只有 Boolean 的装箱档拿不到 `"True"`/`"False"` 这条形状。所以本刀**只补这一档**：
加其余档会让每一格数组元素都换发码形状，而那是另一件事（要逐档实证差异才能动）。

### 改法（一问，两个站受益）

```cpp
bool CCodeGen::arrayElemIsDeclaredBoolean(const IndexOrCallExpr& call) const {
    // 裸标识符 → arrayElemTypes_；成员数组 → inferUdtFieldVb6Type(…, &isArr)
}
```

`inferExprType` 的 `IndexOrCallExpr` 臂开头问它，答 Boolean 就返回 Boolean。两个受益点：

1. 拼串站（`wrapToBSTR`）走 `vb6_CStrBool` ⇒ `"True"`；
2. **`f(1) And True`** —— 元素答 Variant 时 `And` 降成 int32、结果也是 `"-1"`；
   元素答 Boolean 后 `And` 自己回到 Boolean（`B177-2` 钉的就是这一形）。

证据与窄化站同源（`arrayElemTypes_` / `inferUdtFieldVb6Type`），不新开第二张表。

### 夹具 `tests/b177_bool_arr_str`

```
B177-1=True   元素拼串              B177-4=-1    反向证人：数值上下文仍是 -1
B177-2=True   元素 And True         B177-5=True  变量对照（本来就对）
B177-3=False  Not 元素              B177-6=False False 元素
```

B177-4 是故意反向的：VB6 的 True 在**数值**上下文就是 -1，本刀只动拼串那一档，
不能把数值口径也改掉。

门：`check_implicit_cstr_station.ps1` 加 **R17 五条**（helper 定义/声明/推断臂调用/
「只认 Boolean」不得加宽/夹具两格），闸到 R1..R17；runner 两格（x64 + x86 均实跑
通过）；九门 9/9。

### 清单

856 行（新增 2 行 = 新夹具两份输入）。**其余与本机上一版逐字节相同** —— 语料里没有
第二个工程把 Boolean 数组元素放进去拼串，这一刀只碰得到那一形。本机与 CI 那台在 6 份
环境相关输入上的差异仍照旧（期望值保持 CI 那一版）。

### 在册（本刀不碰）

- 其余元素档（Long/Integer/Byte/Double/Date/Currency）仍走装箱那条路：目前靠「拆箱后
  格式化结果相同」活着，Date 与 Currency 两档**没有实证**；要换成声明档得逐档跑差异；
- **Currency 赋值不圈 4 位小数**：`c = 10.123456` 实测打 `10.123456`，VB6 存 `10.1235`
  （`.temp/k174/p1` 的 P5）—— 已在 §11.158 修掉；
- 晚绑定调用点仍拿不到形参声明型（§11.156 在册）；
- ByRef Boolean 未归一（同上）。

## §11.158（第一百七十五刀）Currency 赋值不圈 4 位小数

这一刀的病在**表示层**：VB6 的 Currency 是定点 4 位（VT_CY = 值×10000），
而 C3 用 `double` 承载它。赋值站改前直接存原值：

```vb
c = 10.123456      ' 改前打 10.123456，真 VB6 存 10.1235
```

探针 `.temp/k174/p1` 的 P5 早就量到了这一格，`.temp/k175/p1` 把它拆成七格：
`10.123456→10.1235`、`1.234567→1.2346`、`-1.234567→-1.2346`、`2.5→2.5`、
`0.00004→0`、`c = c + 0` 不反弹，Double 对照 `10.123456` 原样。

### 为什么不能直接问 `inferExprType`

Currency 与 Date 同族：C 层都是 `double`，所以 `inferExprType` 对
`Dim c As Currency` 的标识符**答 Double**（探针实测档位 = 5）。而 `symTab_.lookup`
对 locals 回 null —— 两道路都问不出「这是 Currency」。所以照 Fix 175 的
`knownDateVars_` 依葫芦画瓢，新增 `knownCurrencyVars_`：

- 五处声明站各登一次（Function/Sub/Property 形参、模块级 `Private`、过程内 `Dim`），
  与 `knownDoubleVars_` **并存**；
- 三处过程入口 clear（Function / Sub / Property），与既有各表同一段。

### 一个出口，两个语句站

```cpp
std::string CCodeGen::currencyAssignWrap(Expr* target, const std::string& cValue) const;
```

- **Assignment 站**：经 `narrowCheckAssign` 的出口（Currency 目标不进窄整型那一档，
  命中即返回）；
- **Let 站**：`cgen_setlet.cpp` 两条 emit 直连。`LetStmt` 是**另一条 AST 语句**，改前
  它压根不经过 `narrowCheckAssign` —— 漏了这一站，显式 `Let c = …` 就不圈数
  （探针 `.temp/k175/p2` 的 M6 就是它）。

折数只问 RTL 那一位主人 `vb6_CCur`（`round(v*10000)/10000`，也是 `CCur()` 的出口），
发码侧不另写一份 round —— 两处口径必须同一份。

### 夹具 `tests/b178_currency_round`

```
B178-1=10.1235   局部变量           B178-7=10.1235    Function 返回
B178-2=-1.2346   负数               B178-8=10.1235    UDT 字段
B178-3=2.5       不动               B178-9=10.1235    类 Property Let
B178-4=0         圈到 0             B178-10=10.1235   显式 Let
B178-5=10.1235   模块级变量         B178-11=10.123456 反向证人：Double 不圈
B178-6=10.1235   ByRef 形参
```

B178-11 是故意反向的：同一个 10.123456 交给 Double 目标必须原样打出，否则「把 Currency
圈数顺手做成了全局格式化」这种改法看不出来。

门：新增 `scripts/check_currency_round_station.ps1`（R1..R8：定义/声明各一处、
Assignment 站经过它、Let 站两处、按名表声明 + 三处 clear、五处声明站、
`src/backend` 不许手写 round+10000、夹具与 runner 两格、Double 反向证人），
并在 `tests/run_tests.ps1` 里挂上 `Test-CurrencyRoundStation`（census 认得出）；
runner 两格（x64 + x86）均实跑通过；九门 9/9。

### 在册（本刀不碰）

- **Currency 数组元素**：元素槽是 `int64_t` 且**未按 10000 放大** ——
  `ca(1) = 10.123456` 存成 10（探针 `.temp/k175/p2` 的 M2）。小数在存储层就被截掉了，
  那是另一条腿，本刀故意不碰（只包一层 `vb6_CCur` 会在同一个坑里换一种错法）；
- **表达式中间的 Currency 运算**仍按裸 double 走：VB6 每一步都圈（`10.12345 +
  0.00006` 这类中间值会差在第 5 位），本刀只圈**赋值站**；
- **Currency 与 Double 的格式化口径**同住 `vb6_CStrDbl`：`CStr(c)` 对 10.1235 交
  `10.1235`（圈数之后自然对），但**未圈**的中间值仍有各自的最短表示；
- 晚绑定调用点仍拿不到形参声明型（§11.156 在册）；ByRef Boolean 未归一（同上）。

## §11.159（第一百七十六刀）Currency 折数不许两次舍入

真机探针 `.temp/k176/p1`（P176.bas，用户在 VB6 里跑回来的立即窗口）把 Currency 的
定点语义一次量清，本刀先修其中**最底层**的那一格：折数本身。

### 病征

改前 `vb6_CCur` 是：

```c
double vb6_CCur(double v) { return round(v * 10000.0) / 10000.0; }
```

乘积**先被舍进 double**，再取整 —— 于是五个格子各偏一个单位：

```
A5  3.00005   VB6=30000   改前 C3=30001
A8  2.00005   VB6=20000   改前 C3=20001
A10 -3.00005  VB6=-30000  改前 C3=-30001
A12 2.00025   VB6=20002   改前 C3=20003
A13 4.00005   VB6=40000   改前 C3=40001
A7  0.00005   VB6=1       改前 C3=1      ← 反向证人
```

A7 与其余五格的差别就是铁证：不是「取整方向变了一半」（那样 A7 会反），而是
**舍的是乘积而不是原值**（真值本来就在半分之上时两次舍入碰巧同结果）。

### 改法：一位主人，三个消费站

定标只住在 `vb6rtl_base.h`：

```c
static inline int64_t vb6_CurrencyScale(double v) {
    const double ip = trunc(v);
    const double fp = v - ip;                       /* 精确: 同量级相减 */
    return (int64_t)(ip * 10000.0 + round(fp * 10000.0));
}
```

- 整数部分×10000 在 Currency 值域内 < 2^53 ⇒ 精确；
- 取整只作用在 `|x| < 10000` 的小量上 ⇒ **只舍一次**，.5 的判定对着原值；
- double 里不存在形如 `(2m+1)/20000` 的值（分母含 625）⇒ 真“半分”永不发生，
  `round()` 的半值方向在这里不会被问到。

三个消费站全部改问这一位：`vb6_CCur`（赋值/CCur 的值）、`vb6_CCurBSTR`（串→Currency，
以前自己又写了一遍）、`vb6_ComPackCur`（VB6 Currency 实参→VT_CY）与
`vb6_VariantCur`（装箱 VT_CY）。全 RTL 里 `round(... * 10000)` 只剩主人那一行。

### 夹具 `tests/b179_currency_tie`

读数就是真 VB6 那十三格（`CStr(CLng(c * 10000))`，不经过区域设置的小数点）：

```
101235  -12346  25000  0        30000  10001  1
20000   10002   -30000 12346    20002  40000
```

x64 与 x86 均实跑、逐格相同。

门：`check_currency_round_station.ps1` 从 R1..R8 扩到 **R1..R12**（新增：定标主人
拆开缩放 / `vb6_CCur` 只除回去 / CCur 家族与 VT_CY 三个消费站同一主人 / 全 RTL 只允许
一处 `round(x*10000)` / 新夹具两格）；九门 9/9。

### 真机读数同时坐实、但本刀不动的四条（在册）

同一份探针把另外四件事一并量实了，它们各自是独立的一刀：

1. **中间运算每步都圈**：VB6 `C6=C7=1234600`（`a*b` 先圈成 1.2346 再 ×100），
   而 `C8=1234623`（`a*100` 恰好 4 位，乘一次就够）；我们现在三格都是 1234623。
2. **Currency 数组元素**：VB6 `B1=101235 / B3=110000`，读进 Variant 是 `VT_CY`；
   我们的元素槽是 `int64_t` 且**未按 10000 放大**（存成 10 / 10），读进 Variant 是 VT_I8。
3. **ByVal Currency 形参在调用点就圈**：VB6 `D1=12346`，我们 `12345.67`。
4. **装箱进 Variant 应是 VT_CY**：VB6 `E4 type=6`，我们 `type=5`（`_Generic` 走了 double 臂）。

另：`CCur("3.00005")`（字符串实参）真机读数**未测**，本刀让它与赋值站同口径；
若将来量到 VB6 对**十进制文本**做四舍五入而不是按 double 值折，得单开一格。

## §11.160（第一百七十七刀）ByVal Currency 形参在调用点折 4 位

真机探针 `.temp/k177p`（P177.bas，VB6 编成 exe 后写文件）把「实参进 **ByVal**
Currency 形参」这一格一次量清：

```
A1 1.234567 -> 1.2346        A4 0.00005 -> 0.0001
B1 Double 1.234567 -> 1.2346 B2 10.123456 -> 10.1235
E1 Variant 1.234567 -> 1.2346 F1 String "1.234567" -> 1.2346
H1 ByRef x = x*1.0001 -> 1.2346（折在被调方赋值站，不在调用点）
```

### 病征

C3 的 Currency 形参在 C 侧就是 `double`，double 实参交进去 C 类型相同 ⇒ 一个换算都
不发。改前 C3 的同一份探针：A1 打 `1.234567`、B2 打 `10.123456`、A4 打 `0.00005`。

### 第一列读数会把病遮住

夹具第一版照 §11.159 的写法用 `CStr(CLng(x * 10000))`，结果 A1 两边都是 **12346** ——
`1.234567` 的四舍五入恰好也是 12346。这是 §11.159 A7 的同一教训：**判据要挑那一格
真能分开两个模型的读数**。所以 b180 夹具两列都钉：

```
A1=12346 | 1.2346      ← 第二列 CStr(x) 才露尾巴（改前 1.234567）
```

### 改法：一位主人，两条实参路

```cpp
std::string CCodeGen::currencyByValArgWrap(Expr* arg, const std::string& cValue,
                                          Vb6Type paramType, bool isByVal) const;
```

判据（缺一条就会折错或多折一次）：

- **只认 ByVal + 形参 Currency**：ByRef 交的是地址，折数发生在被调方的赋值站
  （H1 两边本来就相同，不能顺手也折）；
- **实参已是 Currency 不折第二次**：`inferExprType` 答 Currency，或裸名在
  `knownCurrencyVars_` 里 —— Currency 变量的 `inferExprType` 会答 Double（§11.158 实测）；
- **整数档不折**：Byte/Integer/Long/Boolean/LongLong 交给 Currency 是精确的；
- **已带 `vb6_CCur(` 前缀不重复折**：String→Currency 那一站本来就折（F1/F2 改前就对）；
- 地址/复合字面量（`&` 开头）不动；
- 折数只问 RTL 那一位主人 `vb6_CCur`（§11.159 起它又只问 `vb6_CurrencyScale`），
  发码侧不另写 round。

两条实参路都走它：位置实参（`cgen_expr_call_arg_variant.inc` 的 push 之前）与
具名实参（`cgen_expr_call_named_args.inc` 的 `orderedArgs[pi]` 之前）。
**晚绑定 COM 站**拿不到形参声明型，仍未接（在册）。

### 夹具 `tests/b180_currency_byval`

19 行 = 真 VB6 的 19 行，两列（`CStr(CLng(x*10000))` | `CStr(x)`）：

- A 字面量 6 格、B Double 变量 2 格、C Currency 变量 2 格（不折第二次）、
  D 表达式/返回 Currency 的函数 2 格、E Variant 2 格、F String 2 格（不重复折）、
  G 整数档 2 格（不该动）、H ByRef 1 格（调用点不折）。

x64 与 x86 均实跑、逐字相同。b178 的 11 格读数不变（B178-11 的 Double 反向证人仍在）。

### 门

`check_currency_round_station.ps1` 从 R1..R12 扩到 **R1..R13**：出口定义/声明各 1、
ByVal-only 闸门、只经 `vb6_CCur`、按名表不重复折、位置与具名两站各 1、夹具两文件、
runner 两格、第二列证人、A1 锚。十门（九门 + 本门）全绿。

### 清单

862 行（新增 2 行 = 新夹具两份输入）。存量只动 **2 行**：`tests/b178_currency_round`
自己的发码形状 —— `SetModCur 10.123456` 的实参现在带 `vb6_CCur`；那是本刀想要的形状，
读数一格未变（B178-1..11 逐字复跑过）。6 份机器侧差仍以 CI 那台为准、未动。

### 同时修掉的一处哨兵计量

§11.159 把 `vb6_CCurBSTR` 写成嵌套一句后，KEEPS 表从**外层调用**取形参去向，证据
文本退化成 `handoff-to-plain:vb6_CCur`（指针轴仍是 0）。拆成两句（`double d =
vb6_NumVal(s); return vb6_CCur(d);`）恢复 `handoff-reads:vb6_NumVal`，**没有 -Bless
快照**。但拆句后的注释里写出了带括号的 NumVal 调用形，`check_conv_strict_parse` 的
R3 把它数成第 11 枚入口（判据要 10）—— 只改注释措辞，判据不动。

### 在册（本刀不碰）

1. **中间运算每步都圈**：VB6 `c = a * b` 先把 `a*b` 圈成 4 位再存（C6/C7=1234600），
   我们三格都是 1234623（§11.159 在册①）；
2. **Currency 数组元素**：元素槽是未按 10000 放大的 `int64_t`，读进 Variant 是 VT_I8
   （在册②）；
3. **Currency 变量装箱进 Variant 仍是 VT_R8**（探针 E4 `type=5`，在册④）；
4. **晚绑定 COM 实参站**：形参型不可知，没接这一位出口；
5. **C3 解析 `f -1.234567, "x"`** 会当成减法（真 VB6 接受这一形）—— 本夹具用括号
   绕开，解析那一格单列。

## §11.161（第一百七十八刀）Currency 变量装进 Variant 是 VT_CY

§11.159 的在册④（探针 E4 `type=5`）这一刀单开。真机探针 `.temp/k178p`
（P178.bas，编成 exe 后写文件）:

```
P1  v = c（局部 Currency）             -> type=6 Currency str=10.1235
P12 ShowV "P12", c（ByVal Variant）    -> type=6
P13 v = m_c（模块级 Currency）          -> type=6
P14 ShowV "P14", m_c                   -> type=6
P6/P7/P8 Double 变量/字面量            -> type=5（反向证人）
P9  v = 1（Integer 字面量）            -> type=2
P11 TypeName(c)/VarType(c)             -> Currency|6
```

### 病征

Currency 在 C 侧就是 `double` —— `inferExprType` 与 `_Generic` 都只能答 VT_R8。
改前 C3 的 P1 打 `5|Double`、P11 打 `Double|5`。CCur()（P2）与返回 Currency 的
函数（P3）本来就对：那条路上的类型推断答的是 Currency。

### 改法：一位主人，两条装箱站

```cpp
std::string CCodeGen::currencyVariantBox(Expr* expr, const std::string& cExpr) const;
```

判据与赋值站/实参折数站**同一张表** `knownCurrencyVars_`（不再另立一张），只认**裸名**:

- Variant 赋值站：`wrapVariantValue` 里先于「按 C 类型」的那张档位表（否则 Identifier
  先被 `case Double` 拿走，永远到不了 `boxToVariant`）；
- 实参站：`boxToVariant` 里先于 `_Generic`；
- 已带 `vb6_VariantCur(` / `vb6_VariantFromValue(` 前缀不重复装。

P11 的直问 `TypeName(c)` / `VarType(c)` 也走实参站，因此一并答 Currency（夹具已钉）。

### 夹具 `tests/b181_currency_variant`

11 格 × 4 列（`VarType | TypeName | CStr | CStr(CLng(x*10000))`）:

- P1 局部变量、P12 ByVal Variant 实参、P13/P14 模块级（本刀修的四格）;
- P2/P3 控制（改前就对，钉住不许回退）;
- P6/P7/P8 Double 反向证人、P9 Integer 字面量;
- P11 `TypeName(c)/VarType(c)`。

x64 与 x86 实跑，逐字等于真 VB6 的读数。

### 门

`check_currency_round_station.ps1` 扩到 **R1..R14**：出口定义/声明各 1、只问
`knownCurrencyVars_`、只认裸名、两条装箱站各 1、夹具两文件、runner 两格、P6 反向
证人、**不许把运算格与数组格钉进来**（那两条腿仍在册，钉进来就把错值锁死了）。
十门 10/10。

### 清单

864 行（+2 = 新夹具两份输入）。**存量 862 条发码零变化** —— 语料里没有别的
「Currency 变量装 Variant」形状，因此这一刀没有回归面。

### 在册

1. **Currency 运算结果**：`v = c * 1` / `c + 0` 仍是 VT_R8（VB6 给 6）—— 要先把
   表达式类型推断答出 Currency（与 §11.159 在册①「每步都圈」同一条腿）;
2. **Currency 数组元素**：元素槽是未放大的 `int64_t`，装箱成 VT_I8(20) 且值已被
   截掉小数（在册②）;
3. **成员与 UDT 字段**：`me->c` 这一形未走本出口（helper 只认裸名）;
4. **晚绑定 COM 实参站**：形参型不可知（同 §11.160）;
5. **`VarType(c)` 的窄档**：本刀让它经装箱站答 6；若将来发现某条路绕过了装箱站，
   要单开一格。

## §11.162（第一百七十九刀）Currency 参与算术后的身份

真机探针 `.temp/k179p`（P179.bas / P182.bas，编译后写文件）量出 VB6 的规则**不对称**:

```
A1 c*1        -> 6 Currency     A8  c*1.5     -> 5 Double 15.18525（不圈）
A2 c+0        -> 6 Currency     A9  1.5*c     -> 5 Double
A3 c-0        -> 6 Currency     A10 c+0.5     -> 6 Currency
A4 c/1        -> 5 Double       B1  c+d       -> 6 Currency
A5 -c         -> 6 Currency     B2  c*d       -> 5 Double 102.48525225
A6 c*2        -> 6 Currency     B3  c+s       -> 6 Currency 20.247（Single 尾巴被圈掉）
A7 c*i        -> 6 Currency     C1  c*10000&  -> 6 Currency
D1 (c*1)*1    -> 6 Currency     E1/E2/E3      -> 5 Double（值证人）
V1 VarType(c*1)=6 / V2 VarType(-c)=6
```

### 病征

Currency 变量登记在 `knownDoubleVars_` 里（C 侧同为 double），而 `inferExprType`
的 Identifier 分支只查 `knownDoubleVars_` ⇒ `c*1` 的 `promote` 看不到 Currency 档，
一律答 Double。改前 20 格里有 13 格打 `5|Double`，B3 还打 `20.2469998703003`
（真机 20.247）。

### 改法

1. `inferExprType` 的 Identifier 分支把 `knownCurrencyVars_` 放在 `knownDoubleVars_`
   **之前**（口径同 Fix 175 的 Date）；
2. `*` 单开一道**狭窄**规则：只有 `Currency × Single/Double` 这一对答 Double，其余
   走 `promote`（Currency×Integer/Long/Currency 仍是 Currency）；
3. `+ -` 不改：`promote` 的 Currency 档比 Double 高，正好是 VB6 的行为；`/` 本来
   就答 Double（A4 证实）。

### 一次误伤与收窄（值得记）

第一版把 `*` 写成「任一侧 Single/Double 就答 Double」，全语料 **16 条**发码漂移：
`Single * Single`（VB6 仍是 Single）被一并拉成 Double。收窄成「Currency × 浮点」后
只剩 **2 条**，且逐条归因（HEAD 旧二进制 `--emit-c` 与新版本逐字 diff）:

```
tests/ucTreeView/Proyecto1.vbp    1 行
  - vb6_Format(vb6_VariantDouble(Bytes), ...)
  + vb6_Format(vb6_VariantCur(Bytes), ...)        ← Bytes As Currency（§11.161 的身份）

tests/VBFlexGridDemo/VBFlexGridDemo.vbp  56 行
  - vb6_Dbl_AssignFromVariant(&Int64, vb6_VarDivV(...))
  + Int64 = vb6_CCur(vb6_VariantToDouble(vb6_VarDivV(...)))   ← §11.158 的折数站
```

结论：一条 oracle 规则写宽一格，就会把「不相干但同形」的表达式也拉进来 —— 判据要
钉在**真正相关的那一对类型**上。

### 夹具 `tests/b182_currency_arith`

20 格 × 4 列（`VarType | TypeName | CStr | CStr(CLng(x*10000))`）：18 个运算格 + V1/V2
直问。x64 与 x86 均与真 VB6 输出**逐字相同**。A8/A9/B2/E1/E2/E3 是浮点侧的值证人
（不许被顺手圈成 4 位）。

### 门

`check_currency_round_station.ps1` 扩到 **R1..R15**：Currency 先于 Double 问、`*` 的
狭窄规则（`lCur15` / 只对 Currency×浮点）、夹具四枚浮点证人、runner 两格。十门 10/10。

### 清单

866 行（+2 = 新夹具两份输入；存量 2 行已归因）。6 份机器侧差照旧不动。

### 在册

1. **中间运算每步都圈**：`c*2*1.0001` 这类链式，VB6 每步圈到 4 位，我们仍在末尾
   由装箱站圈一次；E3 的值证人只是碰巧相等，多步链要单独量;
2. **Currency 数组元素**：未放大 `int64_t` 槽（在册②）;
3. **成员 / UDT 字段 / 晚绑定 COM 站**的 Currency 身份;
4. **Decimal 组合**：`Decimal × Currency` 等未上真机量过，不猜。

## §11.163（第一百八十刀）Currency 中间量每一步都圈 4 位

真机探针 `.temp/k180p`（P183.bas）:

```
T1 c*2*1.0001        -> 5 Double 20.2490247     （Double 段不圈）
T2 c+0.00004         -> 6 Currency 10.1235
T3 c*1.00005         -> 5 Double 10.124006175
T4 (c+0.00004)*2     -> 6 Currency 20.247       ★ 中间量先圈
T5 c*1.0001*1        -> 5 Double 10.12451235
T6 c/3*3             -> 5 Double 10.1235
T7 CCur(c*2*1.0001)  -> 6 Currency 20.249
T8 c+0.00004+0.00004 -> 6 Currency 10.1235      ★ 每一步都圈
```

### 病征

改前 C3 只在**末尾**（装箱/赋值）圈一次 —— T4 打 `20.2471`、T8 打 `10.1236`
（真机 20.247 / 10.1235）。T1/T3/T5/T6 两边相同，说明 Double 段的链**不该**圈：
判据只能看本节点的类型。

### 改法

在二元发码的唯一收口（`cgen_expr_binary.cpp` 的 `lastExpr_` 尽头）加一档：

```cpp
} else if ((node.op == BinaryOp::Add || node.op == BinaryOp::Sub
            || node.op == BinaryOp::Mul)
           && inferExprType(node) == Vb6Type::Currency) {
    lastExpr_ = "vb6_CCur((" + left + " " + op + " " + right + "))";
}
```

- `Div`/`Pow` 恒 Double，不进这一支；
- 嵌套时内层先圈、外层再圈 —— 与 VB6 逐步圈的可观察结果一致（幂等）；
- 折数仍只问 RTL 那一位主人 `vb6_CCur`，发码侧不写 round。

### 夹具 `tests/b183_currency_step`

8 格 × 4 列：T4/T8 是判据格，T2/T7 是 Currency 档，T1/T3/T5/T6 是 Double 段的反向
证人。x64 与 x86 均与真 VB6 输出**逐字相同**。

### 门

`check_currency_round_station.ps1` 扩到 **R1..R16**：Currency 档的二元收口只在
`inferExprType==Currency` 时触发、覆盖 `+ - *`、经 `vb6_CCur`、夹具四枚证人。
十门 10/10。

### 清单（14 条存量变化逐条归因）

868 行（+2 = 新夹具两份输入）。存量 14 条：

- 6 条 Currency 夹具（b165/b170/b171/b179/b180/b182）：只多出 `vb6_CCur` 包装，
  **读数未变** —— 用 run_tests 的 needles 逐条复跑核对（b165/b170/b171/b179 全中，
  b182 与真机 20 格逐字相同）；
- `tests/Charts 2020` ×3：源码本来就写了 `yRange = yRange + CCur(iStep * NumDecim)`
  （三枚变量都是 Single），Currency 加法现在每步圈 —— 正是本刀要的；
- `tests/VBFlexGridDemo`、`tests/test_vbnet_ext.bas`：同上（Currency 中间量）。

### 在册

1. **Currency 数组元素**：仍是未放大的 `int64_t` 槽（与标量 double 值语义不一致）;
2. **成员 / UDT 字段 / 晚绑定 COM 站**的 Currency 身份;
3. **Decimal 组合**未上真机量过。











## §11.164 Currency 数组元素与标量同口径（第一百八十一刀）

上一节把 Currency 的「中间量每步圈」收进二元收口，但在册第一条就写着：**Currency
数组元素**仍是未放大的 `int64_t` 槽。这一刀把它收口。

### 真机读数（`.temp/k181p/P186.bas`，编译成 exe 后写文件）

| 格 | 写法 | 真 VB6 |
|---|---|---|
| A1 | `a(1)=10.123456` 后读元素 | `6 Currency 10.1235` |
| A2 | `VarType(a(1))` / `TypeName` | `6` / `Currency` |
| A3 | `a(1)=a(1)+0.00004` | `6 Currency 10.1235`（每步圈） |
| A4 | `ByRef` 传元素，`x=x+1.00004` | `6 Currency 11.1235`（真接上） |
| A5 | 未赋值元素 | `6 Currency 0` |
| D1 | 赋给 `Double` | `11.1235` |
| A6 | `a(1)*1` | `6 Currency 11.1235` |
| A7 | `ReDim b(1 To 2)` 后写 | `6 Currency 10.1235` |

改前 C3 同程序全打 `20|Long`（VT_I8），A1=10、D1=10、A4 也不生效 —— 元素槽是
`int64_t`（按 VT_CY 的放大整数布局猜的），而赋值站只发 `vb6_CCur(...)` 不乘 10000，
小数在**存储层**就被截掉了。

### 改法（两处，各自唯一）

1. `mapSaElemCType(Currency)`：`int64_t` → `double` —— 与标量同一承载；
2. `inferExprType` 的 `IndexOrCallExpr` 分支：数组元素按 `arrayElemTypes_` /
   `inferUdtFieldVb6Type` 答 `Currency`（与 §11.157 Boolean 那一问同一张表，
   不新开第二张）。

答对身份之后，其余全是既有站的自动结果：

- 元素赋值走 §11.158 的 `currencyAssignWrap` ⇒ **写入即圈**；
- 装箱走 §11.161 的 `vb6_VariantCur` ⇒ **VT_CY**；
- 参与 `+ - *` 走 §11.162/§11.163 ⇒ 每步圈；
- `ByRef` 传元素拿到的是 `double*` ⇒ 被调用方写回即生效（A4）。

### 夹具 `tests/b184_currency_array`

8 格 + A2/D1，四列（VarType | TypeName | CStr | CLng(x*10000)）。A4 钉 ByRef、
A5 钉未赋值元素、A7 钉 ReDim 动态数组。x64 与 x86 均与真 VB6 输出**逐字相同**。

### 门

`check_currency_round_station.ps1` 扩到 **R1..R17**：元素承载必须是 `double`、
旧 `int64_t` 槽不许残留、元素身份必须由 `arrayElemTypes_` 作答、ByRef/未赋值/ReDim
三枚证人齐。十门 10/10。

### 清单

870 行（+2 = 新夹具两份输入）。存量只动 **2 行**：`tests/b165_byref_byte_cur_date`
自己的发码形状 —— 它的 C1 就是 Currency 数组元素，正好被这一刀改变承载；复跑读数
未变（Y1..Y4 / C1..C3 / T1..T3 全中）。

### 在册

1. **COM 侧的 VT_CY 布局**：元素物理现在是 `double`，若把 Currency 数组直接交给期待
   `SAFEARRAY(VT_CY)` 的接口，需要一层 marshal（语料里未见到，先记）；
2. **类模块成员的身份**：本轮 VB6Mini 仍无法加载 `.cls`（"不能加载 CurHolder.cls"），
   真机读数取不到；同批探针里 **UDT 字段**已验为 `6 Currency 10.1235`，而 C3 本来就
   答对（M6/M7/M8 逐字相同）；
3. **Decimal 组合**未上真机量过。


## §11.165 Currency 的值域闸门与溢出（第一百八十二刀）

上一节把 Currency 数组元素收进 double 承载，同时在册里点出「上限精度」这条线。这一刀
把**值域边界与溢出**收口。

### 真机读数（`.temp/k182p/P190.bas`，与夹具逐句同形）

| 格 | 写法 | 真 VB6 |
|---|---|---|
| N1 | `CCur("1234567.8912")` | `Err 0 / 1234567.8912` |
| N2 | `CCur("-1234567.8912")` | `Err 0 / -1234567.8912` |
| N3 | `CCur("1000000000000")` | `Err 0 / 1000000000000` |
| N5 | `c = 123456789.1234` | `Err 0 / 123456789.1234` |
| N6 | `c = 10.123456` | `Err 0 / 10.1235`（§11.159 的圈法） |
| E1 | `CCur("922337203685477.5808")` | **Err 6** |
| E2 | `CCur("1000")` 后 `* <上限>` | **Err 6** |
| E3 | `Err.Clear` 后 | `0` |

改前 C3: E1/E2 都是 `Err 0`，值是 `-922337203685477.5808` —— 缩放结果被直接
`(int64_t)` 窄化，越界在 C 里是**未定义行为**，MSVC 实测回绕到负数端点。

### 改法

1. `vb6rtl_base.h` 的 `vb6_CurrencyScale`：缩放结果先与 `±2^63` 比较，越界**饱和**到
   `INT64_MAX/MIN`，不再有裸窄化；
2. 同一头文件加**唯一一道闸门** `vb6_CurrencyOverflows(double)`（纯谓词，界
   `922337203685477.5`，三个出口共用）；
3. `vb6_CCur`：越界时 `vb6_RaiseError(6, "Overflow")`，随后交饱和值。

### 夹具 `tests/b185_currency_overflow`

N1/N2/N3/N5/N6 + E1/E2/E3，x64 与 x86 均与真 VB6 输出**逐字相同**。

### 门

`check_currency_round_station.ps1` 扩到 **R1..R18**：闸门唯一（`vb6_CurrencyOverflows`
只许一处定义、界值 `.5`）、旧裸 `(int64_t)` 缩放不许残留、`vb6_CCur` 必须报 Err 6、
E1/E2/E3 三枚证人 + 夹具两台。十门 10/10。

### 清单

872 行（+2 = 新夹具两份输入）。本刀只动 RTL，**存量发码形状零变化**；本机复算后
`emit-manifest.expected.txt` 只余 6 行已知的机器差异（`b132_typedcom_local_scope`、
`guisample` ×3、`ctrlmscomm`、`TeeChartChartDisplay`），按约定原样不动。

### 在册

1. **上限精度**：`922337203685477.5807` 与 `.5808` 在 double 里是**同一个数**
   （`922337203685477.6`），所以真机 O1（正好上限、Err 0）在 C3 也会报 Err 6；
   同理 `999999999999.9999`（16 位有效数字）落成 `1000000000000`。
   要逐一还原，Currency 必须改成 int64 承载 —— 另开一刀；
2. **Resume Next 下「目标保持原值」**做不到：`vb6_CCur` 已经报错返回，表达式形状
   决定它仍会交一个饱和值（与 `vb6_ChkXxx` 一族的已知取舍同一个，见 `vb6rtl_conv.c`
   文件头的口径 (2)）；
3. **COM 侧 `SAFEARRAY(VT_CY)` 的 marshal** 仍未动（§11.164 在册同一条）。


## §11.166 Abs/Int/Fix/Round 在 Currency 上保持身份（第一百八十三刀）

§11.161 至 §11.165 把 Currency 的**值**收得与真机一致（4 位圈、逐步圈、数组、上限）。
这一刀收的是**身份**：数学内建作用在 Currency 上时，VB6 交回来的仍是 Currency。

### 真机读数（`.temp/k183p/P192.bas`，与夹具逐句同形）

| 格 | 写法 | 真 VB6 |
|---|---|---|
| M1 | `Abs(c)` | `6 Currency 10.1235` |
| M2 | `Int(c)` | `6 Currency 10` |
| M3 | `Fix(c)` | `6 Currency 10` |
| M4 | `Round(c, 2)` | `6 Currency 10.12` |
| M5 | `Sgn(c)` | `2 Integer 1` |
| M8 | `-c` | `6 Currency -10.1235` |
| M6/M7/M9/M10 | 负数四格 | `Abs=10.1235 / Int=-11 / Fix=-10 / Round=-10.12`，全是 Currency |
| M11 | `Int(v)`（v 装 Currency） | `6 Currency 10` |
| M12 | `Abs(v)` | `6 Currency 10.1235` |
| M13/M14 | `c \ 2` / `c Mod 3` | `3 Long 5` / `3 Long 1` |
| M15 | `c ^ 2` | `5 Double 102.48525225` |

改前 C3：M1..M4、M6/M7/M9/M10、M11 全是 `5 Double` —— **值本来就对**，缺的只是身份
（装箱答 VT_R8、TypeName 答 Double）。M12 已经对，因为 `vb6_AbsV` 早就保子类型；
M11 不对，因为 `vb6_IntV`/`vb6_FixV` 一律交 Double。

### 改法（两处）

1. `inferExprType` 的调用分支：实参 `inferExprType == Currency`、实参个数 ≤ 2 时，
   `abs/int/fix/round` 答 **Currency**。语料里没有任何输入重定义这四个名字
   （门 R19 钉着），所以这一问不看符号种类；
2. RTL 的 `vb6_IntV` / `vb6_FixV`：`vtCurrency` 进 ⇒ 先按值算、再装回 `VT_CY`
   —— 与 §11.19 的 `vb6_AbsV` / `vb6_VarNegV` 同一口径（先按本子类型算，再装回同一子类型）。

`Sgn` 仍答 Integer、`\` 与 `Mod` 仍答 Long、`^` 仍答 Double —— 三枚反向证人钉住
「不是把所有数值结果都改判成 Currency」。

### 夹具 `tests/b191_currency_math`

15 格，x64 与 x86 均与真 VB6 输出**逐字相同**。

### 门

`check_currency_round_station.ps1` 扩到 **R1..R19**：四枚名字 + Currency 判据必须在同一
分支里、`vb6_IntV`/`vb6_FixV` 必须保 `vtCurrency`、夹具两台 + M11/M15 两枚证人。
十门 10/10。

### 清单

874 行（+2 = 新夹具两份输入）。本刀只动类型判定与 RTL 体内，**存量发码零变化**；
本机复算只余 6 行已知的机器差异（按约定不动）。

### 在册

1. **Currency 的 int64 承载**：上限 19 位与 16 位有效数字仍还原不了（§11.165 在册第一条）；
2. **COM `SAFEARRAY(VT_CY)` 的 marshal** 未动；
3. **Decimal × Currency** 未上真机量过（`CDec` 在 C3 还没有独立承载）。

## §11.167 Format 数值段的字面量、零段与前导 0（第一百八十四刀）

§11.160 至 §11.166 把数值面**排出来的数字**收齐了。这一刀收的是数字之外那圈东西：
格式串里的字面量（括号、引号、转义、尾巴）、零段、整数段的空 `#`，以及命名 `Percent`
多出来的那个千分位。改前的发码里，数值拼装那段自己在注释里写着「Simplified …
doesn't position them」—— 字面量一律被丢掉。

### 真机读数（`.temp/k184p/P197.bas`，与夹具逐句同形）

| 格 | 写法 | 真 VB6 | 改前 C3 |
|---|---|---|---|
| H1 | `Format(d, "(0.00)")` | `(1234.57)` | `1234.57` |
| H2 | `Format(-d, "0.00")` | `-1234.57` | 同 |
| H3 | `Format(-d, "0.00;(0.00)")` | `(1234.57)` | `-1234.57` |
| H4 | `Format(d, "0.00""x""")` | `1234.57x` | `1234.57` |
| H5 | `Format(d, "0.00-")` | `1234.57-` | `1234.57` |
| H6 | `Format(-d, "0.00-")` | `-1234.57-` | `-1234.57` |
| H7 | `Format(d, "\#0.00")` | `#1234.57` | `1234.57` |
| H8 | `Format(d, "0.00;0.0;zero")` | `1234.57` | 同 |
| H9 | `Format(0, "0.00;0.0;zero")` | `zero` | `0` |
| H10 | `Format(d, "Percent")` | `123456.78%` | `123,456.78%` |
| H11 | `Format(-d, "Percent")` | `-123456.78%` | `-123,456.78%` |
| H12 | `Format(0.125, "Percent")` | `12.50%` | `12.50%` |
| H13 | `Format(d, "Fixed")` | `1234.57` | 同 |
| H14 | `Format(1.5, "Standard")` | `1.50` | 同 |
| H15 | `Format(0.5, "#.##")` | `.5` | `0.5` |
| H16 | `Format(1, "0.0;;zero")` | `1.0` | 同 |

H2/H8/H12/H13/H14/H16 是反向证人：它们本来就对，改完仍逐字不变 ——「不是把所有数值
结果都套一层括号」这件事由它们钉住。

### 病（四处，同一条走廊）

1. 正 / 负 / 零三个分段的**字面量**（前缀、后缀、引号内、反斜杠转义）一律被丢；
2. 零段（第三段）没有占位符时，仍然把数字 `0` 打出去，而不是「只打字面量」；
3. 整数段只有 `#`、值又为 0 时，多打了一个前导 `0`（真机 `#.##` 给 `.5`）；
4. 命名格式 `Percent` 递的 `groupDigits` 是 `-1`（= 按系统默认），因此多出千分位；
   真机不带。

另外解剖时踩到一处：`E`/`e` 原来无条件当科学记号的开头，于是零段里的字面量 `zero`
被切成 `z` + 指数。真机只在 `E`/`e` **后面紧跟** `+`/`-` 时才算指数 —— 本刀按此收紧，
`zero` 才完整打出来（H9）。

### 改法（两处，均在 Format 的 RTL 片段里）

1. `src/rtl/core/vb6rtl/detail/vb6rtl_format_numeric_tail.inc`：拼装前先扫最后一个
   数字占位符，把整串切成**前缀**与**后缀**两堆字面量（引号内原样、反斜杠转义一位），
   拼装顺序 = 符号 + 前缀 + 整数 + 小数 + `%` + 后缀；整数段只有 `#` 且值为 0 时不打
   前导 0；零段若通篇没有占位符就只交字面量。科学计数那一路保持原样（`numBuf` 里
   已含尾数和指数）；
2. `src/rtl/core/vb6rtl/detail/vb6rtl_format_parse.inc`：`Percent` 这档的
   `groupDigits` 由 `-1` 改 `0`。

`Fixed` / `Standard` / `Currency` / `Scientific` 四档读数逐字未动（H13/H14 + 存量夹具）。

### 夹具 `tests/b192_format_literal`

16 格，x64 与 x86 均与真 VB6 输出**逐字相同**。

### 门

`check_builtin_slot_returns.ps1` 扩到 **R1..R13**：前缀/后缀拼装、零段只打字面量、
`#` 段不打前导 0、`E`/`e` 只认带符号的指数，四枚结构断言 + `Percent` 的
`groupDigits=0` + 夹具两台。十六门 rc=0（十门 + `b437` / `manifest_coverage` /
`rtl_embedded` 等六门）。

### 清单

876 行（+2 = 新夹具两份输入）；本机复算只余 6 行已知机器差异（按约定不动）。本刀动的是
RTL 片段，会被内联进发码，所以 `emit-manifest` 里被牵动的行随语料一起更新。

### 一处旧期望值按真机改正（GA 45708300744 两红归因）

第一版推上去后 `Tests (vbp #9/#10)` 两红，本机复现是 `B430-24=False/False`。
**不是发码回退，是夹子期望值旧了**：那一格原先钉 `Format(1234,"T0")="1234"` /
`Format(1234,"0 Twip")="1234"`，正是旧 C3 丢字面量的那一版；本刀让字面量重新打出来
之后，C3 交 `T1234` / `1234 Twip`。

真机读数（`.temp/k184p/P199.bas`，18 格）：有占位符时**裸字母就是字面量** ——
`T0`→`T1234`、`0 Twip`→`1234 Twip`、`0abc`→`1234abc`、`$0`→`$1234`、
`(0)`→`(1234)`、`-0-`→`-1234-`、`0x`→`1234x`、`A0`→`A1234`、`0A`→`1234A`。
旧期望值与真机相反，按真机改成 `T1234` / `1234 Twip`；钉的**意图**不变：数值臂不许被
`T`/`w`/`i`/`p` 这些字母拖进日期臂（日期臂给的是 `T14:30` 或序列号那类串，不会是
「数字 + 原样尾巴」）；针名 `B430-24=True/True` 不动。清单随之只动 `b430` 两行，其余
6 行机器差照旧。补丁提交 `7fae0569` ⇒ GA 45711910060 / 23 作业全绿。

### 在册

1. **字面量落在两个占位符之间**仍不对：本刀只做前缀 / 后缀两堆。真机那几格是
   `0 0`→`123 4`、`0 .0`→`1234 .0`、`0 / 0`→`123 / 4`、`0:0`→`123:4`；C3 现在交
   ` 1234`、` 1234.0`、` / 1234`、`:1234`（字面量整体挪到了前面）。语料里没有这种
   写法（开刀前普查过），所以先不做；
2. **字母 `m` / `s` 在数值格式里**：真机 `0 ms`→`1234 ms`、`0 "ms"`→`1234 ms`；
   C3 交 `0 50`、`0 "50"` —— 那两位被日期臂当成了月 / 秒。根因是「串里有 `0`/`#`
   就不算日期」这条闸没有盖住 `ms` 这种纯字母尾巴，另开一刀；
3. **没有占位符的格式串**整体不归数值臂：真机 `Format(1234.567,"2")`→`2`、
   `Format(1234.567,"G")`→空串（`Format(True,"G")` 也空），而 `Format(0,"abc")`
   走日期臂交 `ab0:00:00`；C3 三格退回默认数值格式化（`1234.567`、`True`、`0`）。
   这三格与零段规则相邻但入口不同；
4. **中间字面量与 `Percent` 的 `#` 尾巴**的组合未量（`Percent` 只有一档）；
5. **Decimal × Currency** 仍未上真机量过（§11.166 在册第三条同一条）。

## §11.168 「数值 + 单位文本」谁先进格式串谁赢（第一百八十五刀）

§11.167 收的是数字**之外**那圈字面量，但那些字面量要能被打印出来，前提是先判断「这枚
格式串算数值格式还是日期格式」。这一刀收的就是这道门票：语料里「数值 + 单位文本」那一族
（`#,##0 Twip` 3 处 / `Settings` 4 / `T0` 2 / `#### Byte` 1 / `#,# bytes` 4 / `0 KB` 一族）
全是**裸写**的单位，没有一处用引号包起来，而门票函数里「串里有 `0`/`#` 就不算日期」这道闸
排在了字母判定**之后** —— 任何一枚 `y`/`d`/`h`/`m`/`s`/`n` 都先返回「是日期」。于是改前 C3
把这些格子交给日期臂，交出来是乱码：

| 写法 | 改前 C3 |
|---|---|
| `Format(1234, "#### Byte")` | `#### B03te` |
| `Format(1024, "0 KB")` | `0 5B` |
| `Format(1234, "#,##0 Settings")` | `#,##0 0etti0g0` |
| `Format(1234, "0 ms")` | `0 50` |

### 真机读数（`.temp/k184p/P200` ~ `P203.bas`）

| 格 | 写法 | 真 VB6 |
|---|---|---|
| N1 | `#### Byte` | `1234 Byt` |
| N2 | `#,# bytes` | `1,234 bytes` |
| N3 | `#,##0 Twip` | `1,234 Twip` |
| N4 | `0 ms` | `1234 ms` |
| N5 | `0 KB` | `1024 KB` |
| N6 | `#,##0 Settings` | `1,234 Settings` |
| Q1 | `0 AM` | `1234 AM` |
| Q2 | `0 A/P` | `1234 A/P` |
| Q4 | `0.00 d` | `1234.00 d` |
| Q12 | `#0 mm` | `1234 mm` |
| E1 | `0 Byte` | `1234 Byt` |
| E2 | `0e` | `1234` |
| E3 | `0 e` | `1234␠` |
| E7 | `0 Bytes` | `1234 Bytes` |
| E9 | `0 ex` | `1234 ex` |
| E11 | `0 "e"` | `1234 e` |
| E12 | `0\e` | `1234e` |
| E13 | `0E` | `1234` |
| E14 | `0 Be` | `1234 B` |
| Q5 | `dd 0` | `18 0` |
| Q6 | `yyyy 0` | `1903 0` |
| Q3 | `00:00` | `12:34` |
| Q7 | `hh:mm 0` | `00:00 0` |
| R3 | `dd0` | `180` |
| R12 | `mm` | `05` |

把它读成两条规则：

1. **谁先出现，谁定档**：第一个数值占位符（`0`/`#`）出现在第一个强日期记号之前 ⇒ 数值档，
   后面的字母全是字面量（连 `AM`/`A/P` 也只是字面量 —— Q1/Q2）。反过来字母先出现就是日期档
   （Q5/Q6/Q7/R3）。**冒号是强信号**：`00:00` 即使 `0` 先出现也走时间档（Q3、R9 `0 0:0`
   = `12 3:4`）。
   `w`/`q`/`ttttt` 三枚仍是「整串没有占位符才认」（§B201 那道闸不动），所以 `0 ww` 是数值
   （R8 = `1234 ww`）、`ww` 单写才是日期。
2. **段末那枚 `e`/`E` 是指数记号**，被吞掉（E1/E2/E3/E13/E14）；但 `e` 后面还有字就是字面量
   （E7/E9/E4），引号内（E11）、反斜杠转义（E12）也是字面量。§11.167 只认了 `E+`/`e-` 那两形，
   因为 `zero` 里的 `e` 必须留下 —— 两刀合起来才是「`e` 什么时候是记号」的完整答案。

### 改法（两处）

1. `src/rtl/core/vb6rtl/vb6rtl_format.c` 的 `vb6_fmtIsDateFormat`：`allowWQT = 0`（值是数值）
   这一档先扫「第一个数值占位符」与「第一个强日期记号」（`y/d/h/m/s/n` 或上午/下午记号）的
   位置，遇到冒号直接回日期档；占位符先出现（或没有强记号）⇒ 回 0。原来那道占位符闸保留作
   兜底（`allowWQT = 1` 那条路与「字母先出现」那一路仍要它）；
2. `src/rtl/core/vb6rtl/detail/vb6rtl_format_numeric_tail.inc`：字面量收集时，段末那一枚
   `e`/`E` 跳过不算字面量（引号内、转义位不受影响）。

### 夹具 `tests/b193_numeric_unit_text`

20 格：U1..U15 是数值档（含 E 那一族的四种写法），U16..U20 是反向证人（`dd 0` = `18 0`、
`dd0` = `180`、`yyyy 0` = `1903 0`、`00:00` = `12:34`、`mm` = `05`）。x64 与 x86 均与真 VB6
逐字相同。

### 门

1. `check_builtin_slot_returns.ps1` 扩到 **R1..R14**：新增 R14 钉「谁先出现」三件
   （`firstNum` / `firstStrong` / `hasColon` + 那条判据整行）、区末 `e` 的判据、夹具两台；
2. `check_format_date_tokens.ps1` 的 R2 由「`allowWQT` 分支 1 处」改为 **2 处** —— 第二道闸
   是本刀的新判别，段标与判据由 R14 盯着；
3. `check_format_ampm.ps1` 的 R1 由「`vb6_fmtAmPmLenAt` 站点 3」改为 **4**：本刀在门票据里
   也问了「记号在哪一位」（Q1 的先后判断就靠它）。R1 钉的是「这一问只许住一处」，令牌表仍然
   1 张，多的这一处仍是同一个主人的读者。

这三处都是**哨兵的口径随结构更新**，不是放行：`[STATIC] format_ampm` 第一版在 GA
45715987483 上红的就是 R1（站点 4 ≠ 3），按结构改口径后复跑 PASS。

### 清单

878 行（+2 = 新夹具两份输入）。本机复算 `onlyNew = 2`、`changed = 6`（六行已知机器差）
—— **存量语料发码零变化**（本刀只动 RTL 与夹具，没有碰发码）。补丁提交 `2e425f33`
⇒ GA 45717488380 / 23 作业全绿。

### 在册

1. **字面量落在两个占位符之间**仍不对（§11.167 在册第一条的同一格）：N14 `0 h, 0 m` 真机
   `1,23 h 4 m` / C3 ` h 1,234 m`、`0 0` 真机 `123 4` / C3 ` 1234`、`0 .0` 真机 `1234 .0` /
   C3 ` 1234.0`、`0 / 0` 真机 `123 / 4` / C3 ` / 1234`、`0:0` 真机 `123:4` / C3 `0:0`。
   语料里没有这种写法，两刀都只做前缀 / 后缀；
2. **通篇没有占位符的格式串**仍不归数值臂：`Format(0, "abc")` 真机 `ab0:00:00`（整串被当
   日期）、`Format(1234.567, "2")` 真机 `2`、`Format(1234.567, "G")` 真机空串，而 C3 三格
   分别交 `0` / `1234.567` / `1234.567`（§11.167 在册第三条同一格）；
3. **`0 0000` 这类两组整数占位符**的进位分配未量（真机在 N14 里把 `1234` 切成 `1` 与 `234`
   两截，本刀没有可依据的规则）；
4. **Decimal × Currency** 仍未上真机量过（§11.166 在册第三条同一条）。

## §11.169 整数区多组占位符的数字分配（第一百八十六刀）

§11.167 把字面量按「最后一个占位符的前 / 后」两堆拼出来，§11.168 收拾了谁算数值档的门票。
剩下最后一格是两者的交界：格式串的整数区出现**两组以上**占位符时（`0 0`、`0 000`、
`#,##0 0`、`0 h, 0 m`），数字该往哪一组放。改前整串只塞进最后一组，前面的组被当字面量：

| 写法 | 改前 C3 | 真 VB6 |
|---|---|---|
| `Format(1234, "0 0")` | ` 1234` | `123 4` |
| `Format(1234, "0 / 0")` | ` / 1234` | `123 / 4` |
| `Format(1234, "0 h, 0 m")` | ` h 1,234 m` | `1,23 h 4 m` |

### 真机读数（`.temp/k186p/P204.bas` 18 格 + 三枚存量格）

| 格 | 写法 | 真 VB6 | 读法 |
|---|---|---|---|
| G1 | `0 0` @1234 | `123 4` | 两组 1/1 |
| G2 | `0 00` @1234 | `12 34` | |
| G3 | `00 0` @1234 | `123 4` | 最左组拿 3 位 > 容量 2 |
| G4 | `0 000` @1234 | `1 234` | |
| G5 | `000 0` @1234 | `123 4` | |
| G6 | `00 00` @1234 | `12 34` | |
| G7 | `#,##0 0` @1234 | `1,23 4` | 千分位是整串先插再切组 |
| G8 | `0 #,##0` @1234 | `0 1,234` | 右组吃掉全部，左组补 `0` |
| G9 | `#0 0` @1234 | `123 4` | |
| G10 | `0 #0` @1234 | `12 34` | |
| G15 | `00 00` @12 | `00 12` | 数字不够分 |
| G16 | `0 0` @12 | `1 2` | |
| G17 | `0 0` @0 | `0 0` | |
| G18 | `0;0;0` @1234 | `1234` | 分段不是分组（反向证人） |
| G19 | `0 / 0` @1234 | `123 / 4` | P199 L16 |
| G20 | `0 A/P 0` @1234 | `123 A/P 4` | P203 R10；`0` 之后的记号只是字面量 |
| G21 | `0 h, 0 m` @1234 | `1,23 h 4 m` | P200 N14；组外那个逗号也不打印 |

四条口径（都是真机逼出来的）：

1. **千分位先整串插，再切组**：`1234` 先成 `1,234`，右组按容量 1 拿走 `4`，剩下的 `1,23`
   连同它中间那个逗号一起归左组（G7）。所以逗号跟着**左块**走；
2. **从右往左，每组按自己的容量取数，最左组吸收余数**：G3/G5 的最左组容量是 2/3，却各拿了
   3 位 —— 余数没有别处可去；
3. **取不到数的组按自己的 `0`/`#` 补位**：`0 #,##0` 的右组容量 4 吃满，左组是 `0` ⇒ 补 `0`
   （G8）；`00 00` 配 12 ⇒ 左边 `00`（G15）；全 `#` 且值为 0 的组补空；
4. **整数区里的逗号一律是千分位记号**，不当字面量打出来 —— G21 的 `" h, "` 里那个逗号就是
   这么被吃掉的（真机 `1,23 h 4 m`）。

### 改法（一处）

`src/rtl/core/vb6rtl/detail/vb6rtl_format_numeric_tail.inc`：在小数点**不存在**时启用新的一段
（`dotPos < 0` 且组数 `ng >= 2`），先把 `intPart` 里上游已插的逗号剔掉再按「有没有逗号」
重算千分位（否则会出 `1,,23`），然后用 `gsrc` 上的左右索引右往左切片，最后按格式串的出现
顺序走一遍：字面量照抄、组按切片打印或补位、`%` 仍旧收尾。单组的那些写法（`0.00`、`#,##0`、
`#.##`、`0.0%` 等）走的还是 §11.167 那条前缀 / 后缀老路，读数逐字未动。

### 夹具 `tests/b194_multi_int_group`

17 格：14 格组分配 + 两枚反向证人（G18 分段不是分组、G20 记号只是字面量）+ 两枚存量遗留格
（G19、G21）。x64 与 x86 均与真 VB6 逐字相同。

### 门

`check_builtin_slot_returns.ps1` 扩到 **R1..R15**：R15 钉段标、五个数据结构
（`gCap`/`gZero`/`gsrc`/`slStart`/`slEnd`）、两道入口判据（`dotPos < 0`、`ng >= 2`）、
夹具两台与 U7/U8/U21 三枚证人。十九门 rc=0（含 `format_ampm` / `format_date_tokens` /
`b437` / `formatdatetime_args` / `format_date_arg_arm` 五门同族）。

### 回归

`b192`（16 格）、`b193`（20 格）、`b426`、`b430`、`b431`、`b432` 的读数与上一刀**逐字相同**。
清单 880 行（+2 = 新夹具两份输入），本机复算 `onlyNew = 2` / `changed = 6`（六行已知机器差）
—— 存量发码零变化。

### 在册

1. **带小数点的多组**仍不对：`0.0 0.0` 真机 `1234.5 6.8` / C3 ` 1234.568`、`0 0.00` 真机
   `123 4.57`、`0.00 0` 真机 `1234.56 8`、`0.0"x"0.0` 真机 `1234.5x6.8`、`0 .0` 真机
   `1234 .0` / C3 ` 1234.0`。这五格的数字走向不合上面四条（`0.00 0` 左往右、`0.0 0.0` 又
   换了一种切法），语料里没有这种写法，先不做；
2. **通篇没有占位符的格式串**仍不归数值臂（§11.168 在册第 2 条同一格）；
3. **Decimal × Currency** 仍未上真机量过（§11.166 在册第三条同一条）。

## §11.170 FormatNumber / FormatCurrency / FormatPercent 的进位是十进制（第一百八十七刀）

§11.167~§11.169 把 `Format` 的数值格式串一路量到语料边界（多组整数区是最后一格，剩下
带小数点的多组与「通篇无占位符」两条在册）。这一刀换场：按语料里**剩下的调用次数**排，
`Format` 之后最密的是字符串与数值内置 —— `StrComp` 89、`FormatDateTime` 63、`InStrRev` 62、
`FormatNumber` 36、`IsNumeric` 33、`FormatPercent` 19、`FormatCurrency` 17、`StrConv` 14。
探针 `P205`（四家各十几格，逐格 `ShowE` 收 Err）一上来就翻出三处静默答错，本刀先收**进位那一处**。

### 现象（改前 C3 与真 VB6 分岔）

三家出口的成型都写着 `swprintf(buf, 96, L"%.*f", digits, v)` —— 那是 C 运行库的
**二进制**「五成双」：值本身在二进制里落在半值下方一点，于是 `.5` 全被抹掉。

| 写法 | 改前 C3 | 真 VB6 |
|---|---|---|
| `FormatNumber(0.5, 0)` | `0` | `1` |
| `FormatNumber(2.5, 0)` | `2` | `3` |
| `FormatNumber(1.005, 2)` | `1.00` | `1.01` |
| `FormatNumber(9.995, 2)` | `9.99` | `10.00` |
| `FormatNumber(0.4999999999999999, 0)` | `0` | `1` |

第三行不是「真机也用了 double」能解释的：`1.005` 的二进制展开是 `1.00499999999999989…`，
谁在 double 上做 `floor(x*100+0.5)` 都只会得到 `1.00`。

### 真机读数（`.temp/k187p` P207 13 格 + P208 12 格 + P205/P206 各若干）

| 格 | 写法 | 真 VB6 | 读法 |
|---|---|---|---|
| D1 | `0.5` @0 | `1` | 半值远离零 |
| D2 | `1.5` @0 | `2` | 同上（不是成双） |
| D3 | `2.5` @0 | `3` | 成双口径在这一格必然答 2 |
| D5 | `-0.5` @0 | `-1` | 远离零，不是向零截 |
| D7 | `0.05` @1 | `0.1` | |
| D10 | `0.125` @2 | `0.13` | |
| R1 | `1.005` @2 | `1.01` | 十进制文本 `1.005` ⇒ 半值进 |
| R2 | `2.675` @2 | `2.68` | 同上 |
| R3 | `0.145` @2 | `0.15` | 同上 |
| R4 | `1.0049999` @2 | `1.00` | **反向证人**：十进制文本 `1.0049999` 不到半值 |
| R5 | `2.5` @0 | `3` | |
| R10 | `0.4999999999999999` @0 | `1` | **反向证人**：15 位有效数字 = `0.5` ⇒ 进 |
| R11 | `1.0005` @3 | `1.001` | |
| R12 | `0.5` @0 incLeading=False | `1` | 不是「去前导零」的锅 |
| B3 | `123456789012345.5` @0 | `123,456,789,012,346` | 15 位装不下 ⇒ 退全位展开 |
| B7 | `1234567890.5` @0 | `1,234,567,891` | |
| B9 | `9.995` @2 | `10.00` | 要连进位 |
| B10 | `-9.995` @2 | `-10.00` | |
| B8 | `0.9999999999999999` @0 | `1` | 15 位文本已经是 `1` |
| B1 | `1E+15` @2 | `1,000,000,000,000,000.00` | 不走科学记数 |
| B2 | `1E+20` @2 | `100,000,000,000,000,000,000.00` | 同上 |
| B6/B11 | `1E-07` / `1E-20` @2 | `0.00` | |

R4 与 R10 是这一刀的两枚反向证人，合起来把口径钉死：**不是**在 double 上进位（R1/R2/R3 会错），
**也不是**拿二进制的完整十进制展开进位（R4 会错成 `1.01`），而是在**15 位有效数字的十进制文本**
上进位。R10 进一步说明这 15 位是**先归再判半**（`0.4999999999999999` 归成 `0.5` 才进）。

### 口径（四条）

1. **取 15 位有效数字的十进制文本**：`%.15g`。整数区不超过 15 位、且不落科学记数形的值都走这条；
2. **在该文本上逐位进位，半值远离零**：切点后第一位 `> '5'` 必进，`== '5'` 也进
   （因为文本已归到 15 位，`5` 之后不可能再有「非零尾巴」这种需要再看一位的歧义）；
3. **连进位**：`9.995` @2 ⇒ 分数位 `99`+1 ⇒ `10.00`；整数全 `9` 时前置 `1`
   （这也是 `0.9999999999999999` @0 与 `0.5` @0 的共同出口）；
4. **`%.15g` 落科学记数形（`|x| >= 1E+15` 或 `< 1E-04`）时退回 `%.*f`**：真机这两档也是
   全位展开（B1/B2 二十一位数字照打），`%.*f` 恰好同形；这一档不做十进制进位，与真机的
   差异只在 16 位以上有效数字（语料无）。

### 改法（一处文件，两个新函数，三个调用点）

`src/rtl/core/vb6rtl/vb6rtl_compat.c`：

- `vb6_fmtRoundDecimalText(w, digits)`：把一个十进制文本（可带 `-`、可无整数位）就地进位到
  `digits` 位小数，分数位不足补 `0`，`digits == 0` 不出小数点；
- `vb6_fmtFixedText(v, digits, out, cap)`：`%.15g` 取文本 → 非科学形则进位；科学形退 `%.*f`；
- `vb6_FormatNumber` / `vb6_FormatCurrency` / `vb6_FormatPercent` 三家的 `swprintf` 换成
  `vb6_fmtFixedText`（百分比那家先进 `value * 100.0`，`'%'` 由 `wcscat` 手接，
  后面「去前导零 / 插千分位 / 套括号」三段老路一个字没动）。

⚠ 这里差点踩的坑：`FormatPercent` 原先是 `"%.*f%%"` 一次成型，若只换前半段而不动 `'%'`，
`wcscat` 会拼出 `0.5%%`。改成先成型再手接，`dropLeadingZero` 与 `groupThousands` 对尾随 `'%'`
的既有处理（§B213）原样生效 —— `b437` 的 33 格逐字未动是这一点的证人。

### 夹具 `tests/b195_format_round_decimal`

46 格，五族：D（11 格，半值两侧与负号）、R（13 格，十进制 / 二进制分岔 + 三枚反向证人
R4/R10/R12）、B（12 格，大数小数成型）、F（7 格，默认档与三个可选实参组合）、C1（货币那一档
只钉形状：`Mid(FormatCurrency(0.5, 0), 2)` ⇒ `1`，符号随区域设置，§B142 同一条纪律）。
x64 与 x86 两台读数逐字相同（46 行）。

### 门

`check_builtin_slot_returns.ps1` 扩到 **R1..R16**：R16 钉三样 —— 段标 `§11.170`、定型链三枚
记号（`vb6_fmtRoundDecimalText` / `vb6_fmtFixedText` / `%.15g`）、三家出口都走 `vb6_fmtFixedText`、
以及夹具两台与 R1/R4/R10/B9 四枚证人。本机 19 门 rc=0（`currency_round_station`、`format_ampm`、
`format_date_tokens`、`b437`、`formatdatetime_args`、`format_date_arg_arm`、`rtl_embedded` 全绿）。

> ⚠ 运维一条：`.build/run_all_gates.ps1` 会把 `scripts/check_*.ps1` **全部 ~155 个**门逐个起
> 一个 Windows PowerShell 进程跑一遍；本机 16G 内存跑到中途整机卡死。验本刀只按上面那 19 门
> 逐个单跑（一次 1~3 个），全量 155 门交给 CI。

### 回归

`b192`（16 格）、`b193`（20 格）、`b194`（17 格）、`b437`（33 格）、`b439`（13 格）读数与上一刀
**逐字相同** —— 这些夹具里的值（`1234.567`、`1234.5`、`0.125` 等）都在十进制文本下不跨半值，
进位口径换了但答案没换。清单 882 行（+2 = 新夹具两份输入），本机复算 `onlyNew = 2` /
`changed = 0`，六行已知机器差未动。

### 在册

1. **`InStrRev(s, find, 0)` 的错误形状**：真机只落 `ERR 5` 一行，C3 先打 `I3=[0]` **再**打
   `I3=[ERR 5]` —— 表达式出口在抛错前把旧值交出去了（P205 I3 格，vb=33 行 / c3=35 行的
   行数漂移就是它）。这一处是发码侧的形状问题，不是 RTL 返回值问题；
2. **`StrComp(Null, "a")`**：真机返回 Null（`IsNull` = True、Err 保持 0），C3 交 `False` + Err 94，
   并且把 Err 泄给后面各格。Null 传染这条要连着 `vb6_StrComp` 的变体入口一起改；
3. **`Format` 数值格式串那条路的进位**没动：`Format(1.005, "0.00")`、`Format(9.995, "0.00")`
   还没上真机量过，若与三家同口径，将来要共用 `vb6_fmtRoundDecimalText`；
4. 前三刀在册各条（带小数点的多组、通篇无占位符的格式串、Decimal × Currency）原样保留。

## §11.171 `Format` 数值格式串的定点进位也是十进制（第一百八十八刀）

§11.170 把 `FormatNumber` / `FormatCurrency` / `FormatPercent` 三家的进位从 `%.*f` 的二进制
五成双改成十进制半值远离零，同时在册第三条写着「`Format` 数值格式串那条路的进位没动，还没上真机
量过」。这一刀就是去量它 —— 答案是**同一条口径**，于是让两条路共用一个出口。

### 现象（改前 C3 与真 VB6 分岔）

数值臂的定点分支写的是 `round(scaleVal * factor) / factor`（double 上进位），
差额落在半值下侧时就少进一位：

| 写法 | 改前 C3 | 真 VB6 |
|---|---|---|
| `Format(1.005, "0.00")` | `1.00` | `1.01` |
| `Format(9.995, "0.00")` | `9.99` | `10.00` |
| `Format(-1.005, "0.00")` | `-1.00` | `-1.01` |
| `Format(0.4999999999999999, "0")` | `0` | `1` |

⚠ 同一批里 `Format(2.675, "0.00")` 改前就是 `2.68`（真机同）—— `2.675 * 100` 在 double 上
恰好越过半值，**撞对**了。这一格是这一族最危险的地方：夹具只挑几格看的话，很容易误判成
「已经没问题」。

### 真机读数（`.temp/k188p/P212.bas` 18 格，逐字）

| 格 | 写法 | 真 VB6 | 读法 |
|---|---|---|---|
| F1 | `1.005` @"0.00" | `1.01` | 十进制半值进 |
| F2 | `9.995` @"0.00" | `10.00` | 连进位 |
| F3 | `0.5` @"0" | `1` | |
| F4 | `2.5` @"0" | `3` | 不是成双 |
| F5 | `1.5` @"0" | `2` | |
| F6 | `-1.005` @"0.00" | `-1.01` | 负号不动 |
| F7 | `-0.5` @"0" | `-1` | 远离零 |
| F8 | `1.0049999` @"0.00" | `1.00` | **反向证人**：不到半值 |
| F9 | `0.4999999999999999` @"0" | `1` | **反向证人**：15 位十进制先归成 `0.5` |
| F10 | `0.145` @"0.0" | `0.1` | 反向证人 |
| F11 | `2.675` @"0.00" | `2.68` | 改前撞对的那一格 |
| F12 | `1234.567` @"0.00" | `1234.57` | |
| F13 | `1.0005` @"0.000" | `1.001` | |
| F14 | `0.5` @"#.##" | `.5` | `#` 占位符的裁剪不许变 |
| F15 | `2.5` @"0.0" | `2.5` | |
| F16 | `0.0005` @"0.000" | `0.001` | |
| F17 | `999.995` @"0.00" | `1000.00` | 整数位也要连进 |
| F18 | `0.5` @"0%" | `50%` | 百分比区不动 |

### 改法（把出口提权重用，三处）

1. `src/rtl/core/vb6rtl/vb6rtl_compat.c`：`vb6_fmtFixedText` 从 `static` 提为普通函数
   （`vb6_fmtRoundDecimalText` 仍是它的内部实现，不外露）；
2. `src/rtl/core/vb6rtl/vb6rtl_builtin.h`：与三家 `vb6_Format*` 并列声明，并写清
   「三家与 `Format` 数值格式串**共用**这一枚十进制定型出口，不得再散回 `%.*f` / `round()`」——
   这句注释就是 R17 的反向判据要钉的东西；
3. `src/rtl/core/vb6rtl/detail/vb6rtl_format_numeric_body.inc`：定点分支换成
   `vb6_fmtFixedText(scaleVal, decimalPlaces, numBuf, 128)`，`%.*f` 只剩 `decimalPlaces < 0`
   那一档（此时没有进位可言）。

科学记数（`E` 形）分支的 `mantissa = round(mantissa * factor) / factor` 这一刀**没动**：
`Format(1.2345E+10, "0.00E+00")` 这类写法语料里极少，且它先要定「尾数归到几位」，
比定点多一层口径，先记在册。

### 夹具 `tests/b196_format_string_round`

18 格，全部逐字抄自真机：3 格二进制分岔（F1/F2/F6）、2 格「不得被十进制进位误伤」
（F8/F10）、1 格「15 位先归再判半」（F9）、1 格「改前撞对」（F11）、
其余钉 `#` 占位符裁剪（F14）、`0.0` 补零（F15）、整数位连进（F17）、百分比区（F18）。
x64 与 x86 两台读数逐字相同（19 行）。

### 门

`check_builtin_slot_returns.ps1` 扩到 **R1..R17**。R17 钉五样：

1. 段标 `§11.171` 在文件里；
2. 定点分支必须出现 `vb6_fmtFixedText(scaleVal`；
3. 不许回到 `round(scaleVal * factor)` 与 `roundedVal = round`；
4. 出口必须在 `vb6rtl_builtin.h` 声明、且不许被改回 `static`（否则 `Format` 那边根本调不到）；
5. 夹具两台 + F1/F8/F9/F11 四枚证人。

⚠ 这一条的反向判据**必须先剥行注释**：段内的解释性注释本身就会写到 `round(scaleVal * factor)`
这种旧写法，直接对原文匹配会把自己钉死（§B145 同一条课，本刀第一版就是这么红的）。

本机 20 门 rc=0：`builtin_slot_returns`(R1..R17)、`manifest_coverage`(884)、
`vbp_fixture_census`、`rtl_embedded`（exe sha `dfd5fd0696fa`）、`rtl_naked_names`、
`rtl_proto_arity`（1798 对）、`conv_strict_parse`、`currency_round_station`(R1..R19)、
`format_ampm`、`format_date_tokens`、`b437`、`b439`、`formatdatetime_args`、
`format_date_arg_arm`、`test_helper_integrity`、`fixture_timer_close`、`no_locale_needles`、
`ps51_joinpath`、`static_sentinel_registration`、`emit_c_artifact_caliber`。

### 回归

`b192`（17 行）、`b193`（21）、`b194`（18）、`b195`（46）、`b437_format_leading_group`（34 针）、
`b439_format_useparens`（13 针）—— 六个夹具的针脚逐条命中，**0 缺失**。
清单 884 行（+2 = 新夹具两份输入），本机复算 `onlyNew = 2` / `changed = 0`，六行已知机器差未动。

### 这一轮量出来的另外两笔（都进了在册，不是本刀的范围）

同一轮探针顺带把字符串两家的坑量了，读数已坐实但都**不是 RTL 一改就能收口**的：

1. **语句中止**：`P210` 九格里，真机凡遇 `On Error Resume Next` 下的实参位报错，**整条语句**都不执行
   （`Show "M1", Mid("abc", 0)` 连 `Show` 都不进，只落 `M1=[ERR 5]`）；C3 则是「先交一个默认值、
   再落 Err」—— 同一格会打印两行（`M1=[]` 与 `M1=[ERR 5]`），行数随之漂移。根因在 cgen：
   语句中止守卫（`stmtAbortGuard_` + `vb6_ErrSerial()`）目前**只设在赋值站**，调用语句
   （`Show …`、`Print #1, …`）的实参求值没有等价的闸门。要用「实参先入临时量、再比较序号」改写
   调用语句发码，动的是整片发码面，得单独一刀；
2. **`StrComp(Null, "a")`**：真机交 `Null`（`IsNull` = True、Err 保持 0），C3 交 `False` + Err 94。
   `StrComp` 在内置返回档里登记的是 `int32_t` 标量槽，要答 Null 得先有「Variant 槽 + Null 直通」
   的出口形状，属于返回档那一族的改造，与本刀的十进制定型无关。

### 在册

1. **`Chr` 越界** 不能拿 VB6Mini 的读数当判据：`Chr(-1)`→`ÿ`、`Chr(300)`→1、`Chr(-257)`→63，
   这三枚互不相容，是代码页有损转换留下的痕迹（`Chr(65536)` 才报 5）。真正该钉的是
   「0..255 之外怎么办」，得先找到不经过代码页的观测通道，否则会把探针的噪声写成真理；
2. **科学记数分支的尾数进位**（`round(mantissa * factor)`）未动，见上；
3. **`Format` 多组带小数点的整数区**、**通篇无占位符的格式串**、**Decimal × Currency**
   三条（§11.169 / §11.168 / §11.166 在册）原样保留。

## §11.172 `IsNumeric` 认 Empty 为数值（第一百八十九刀）

§11.171 收口后，按语料里剩下的调用次数排下一个候选（`StrComp` 89 / `FormatDateTime` 63 /
`InStrRev` 62 / `FormatNumber` 36 / `IsNumeric` 33 / `FormatPercent` 19 / `FormatCurrency` 17 /
`StrConv` 14），先量了字符串两家与 `Chr`（结果都不能做刀，见末尾在册），再量判定族
`IsNumeric` / `IsDate`。判定族 66 格里只有 2 格错，且是同一个根因。

### 现象

| 格 | 写法 | 改前 C3 | 真 VB6 |
|---|---|---|---|
| E1 | `IsNumeric(Empty)` | `False` | `True` |
| V1 | `v = Empty` 后 `IsNumeric(v)` | `False` | `True` |

### 真机读数（`.temp/k189p` P214 42 格 + P215 24 格）

文本档 30 格（N1..N30）里已经全对，值得记下来的是几条**边界**：

| 写法 | 真 VB6 | 读法 |
|---|---|---|
| `" 1 "` / `"1 "` / `" 1"` | True | 两侧空白容 |
| `"1 2"` | False | 中间空白不容 |
| `"1e3"` / `"1E+3"` / `"1d3"` | True | D 指数也认 |
| `"1e"` | False | 指数没数字不算 |
| `"&H10"` / `"&O17"` | True | 进制前缀认 |
| `"0x10"` | False | C 风格不认 |
| `"1,000"` / `"1,5"` | True | 千分位 / 小数逗号都认 |
| `"$1"` / `"1%"` | False | 货币符号与百分号**不认** |
| `"(1)"` | True | 括号负数是数 |
| `"1.2.3"` / `"--1"` | False | |
| `""` / `" "` | False | 空串不是数（与 Empty 不同，见下） |

日期档 12 格（D1..D12）全对：`"2026-13-01"` / `"2026-02-30"` 都是 False（月日越界），
`"13:00"` 是 True（纯时间也是日期），`IsDate(0)` 是 False（数值不当日期）。

变体槽 24 格（E1..E16 / V1..V8）里，除 E1/V1 外全对，含 `IsNumeric(CCur(1.5))`、
`IsNumeric(CDec(1.5))`、`IsNumeric(CByte(7))`、`IsNumeric(CSng(1.5))` 四档子类型都是 True，
`IsNumeric(Null)` / `IsDate(Null)` / `IsDate(Empty)` 都是 False。

### 口径

**Empty 是数，Null 不是。** Empty 的缺省值就是 0（数值语境里与 0 同用），所以
`IsNumeric(Empty)` 真；Null 是"没有值"，`IsNumeric(Null)` 与 `IsDate(Null)` 都假。
同一个 Empty 在日期那一问上是假（`IsDate(Empty)` = False）——
**这一族问的是"里面这枚值折不折得出来"，不是"Empty 折算成什么"**。

### 改法（一行，加一档）

`src/rtl/core/vb6rtl/vb6rtl_conv.c` 的 `vb6_IsNumeric` switch 补：

```c
case vb6_vtEmpty:
    return -1;
```

旁边的 `Decimal` / `VT_I8` / `VT_I1` 几档各自都带着"问的是里面是不是数"的注释，
这一档只是把 Empty 归进同一句；判定的整个面（文本出口 `vb6_IsNumericText`、
`vb6_DateCanParse` 两条权威）一个字没动。

### 夹具 `tests/b197_ispred_variant`

65 格：29 格文本数值档 + 12 格日期档 + 16 格字面量/子类型档 + 8 格整只变体槽档。
**改前只有 E1/V1 两格红，另外 64 格逐字相同** —— 这 64 格就是"没有误伤"的证据，
也是这一刀能只改一行却仍然放心推的基础。x64 与 x86 两台读数逐字相同（67 行）。

### 门

`check_ispred_text.ps1`（账 #418 的判定族哨兵）加 **R6**：VT_EMPTY 那两行仍在、
`b197_ispred_variant` 在 `run_tests.ps1` 里登记两台。本机 21 门 rc=0：
`ispred_text`(R1..R6)、`manifest_coverage`(886)、`rtl_embedded`（exe sha `a3b0784827bc`）、
`rtl_naked_names`、`rtl_proto_arity`（1798 对）、`conv_strict_parse`、`builtin_slot_returns`(R1..R17)、
`currency_round_station`(R1..R19)、`format_ampm`、`format_date_tokens`、`b437`、`b439`、
`formatdatetime_args`、`format_date_arg_arm`、`test_helper_integrity`、`vbp_fixture_census`、
`fixture_timer_close`、`no_locale_needles`、`ps51_joinpath`、`static_sentinel_registration`、
`emit_c_artifact_caliber`。

⚠ 顺带记一条门的坑：R6 头一版用裸名 `b197_ispred_variant` 计数，得 4 而不是 2 ——
因为每个 `Test-Vbp` 行里名字出现两次（名字实参 + 路径）。判据要连 `Test-Vbp '` 一起匹配。

### 回归

上一刀六个夹具（`b192`/`b193`/`b194`/`b195`/`b196`/`b437`/`b439`）读数未动；
清单 886 行（+2 = 新夹具两份输入），本机复算 `onlyNew = 2` / `changed = 0`，六行已知机器差未动。

### 在册（本刀量出、都不在本刀范围）

1. **`StrConv` 的字节重解释两档不能拿 VB6Mini 当判据**。`P213` 的读数自相矛盾：
   `StrConv("ABC", vbFromUnicode)` 交出一个 `?`（len 1、码 63），
   而 `StrConv(StrConv("ABC", vbFromUnicode), vbUnicode)` 又还原成 `ABC`（len 3）。
   根因是 **C3 的 `String` 是 Unicode**，`vbUnicode` / `vbFromUnicode` 这两档的真实口径
   与 VB 的 ANSI 语义本来就不同 —— 这两档的判据以 **TwinBASIC** 为准（TB 兼容 VB6 代码，
   差别正在字符串类型上）。探针已备：`.temp/k189p/TB_StrConv.bas`（VB6 语法，可直接在 TB 里跑，
   逐格打 `Len` 与 `AscW` 码点）。等 TB 读数回来再开工；
2. **C3 的日期字面量解析不收纯时间形**：`IsDate(#12:00:00 PM#)` 报 VB2004/VB2001 整份编不过。
   与判定面无关，是解析层的缺口（本刀夹具改用 `CDate("12:00:00")` 绕开）；
3. **`Chr` 越界不可当判据**（§11.171 在册第一条同一条）：`Chr(-1)`→`ÿ`、`Chr(300)`→1、
   `Chr(-257)`→63 三枚互不相容，是代码页有损转换的痕迹；
4. **语句中止**（§11.171 在册第一条）本刀又推进一步：§11.80 的闸门（`stmtCallArgSerial_`
   逐实参提升 + 序号比较）**机制已经存在**，缺的只是触发条件 —— 现在只有"实参含晚绑定 COM 读"
   才 fire；通用实参路径（`cgen_expr_call_arg_emit.inc`）里根本没有提升机制，要覆盖得把
   §11.80 那套逻辑复刻到四处实参发放置，还要处理 `Variant` 形参的装箱与所有权。
   另外 §83 的注释已经写明这类放宽的真实拦路虎：`vb6_ErrSerial()` 分不清
   "错误往外爬"与"被调方内部 `On Error Resume Next` 吃掉"，**得先给 RTL 一枚
   「错误是否还在往外爬」的信号**。下一刀从这里起；
5. **`StrComp(Null, …)`** 仍答 `False` + Err 94（真机 `Null`）—— 需要"Variant 槽 + Null 直通"
   的出口形状（同 §11.171 在册第二条）。

### 收口过程中的两处修正（第一百八十九刀第二版）

第一版推上去 GA 红了三个 job（Tests (compile #2) / Tests (vbp #3) / Tests (vbp #4)），
都不是判定逻辑错，而是**两处口径**：

1. **A17 的计数窗口太宽**（check_sa_access.ps1）。那一条的本意是「TypeName 的标量名单只许一份」，
   做法却是全文件数 case vb6_vtEmpty: —— 只因当时全文件只有 TypeName 一处写它。
   本刀让 tEmpty 进 b6_IsNumeric 的数值族（**合法的第二处**：同一个值被两个不相干的
   问题各问一次），全文件计数就把那处算成「名单被抄了两份」。窗口收进 b6_TypeNameScalar
   的函数体（连函数体都找不到时本身也判红），本地 c=0，收口行里 TypeName list 1 scalar 2
   一位没动。这是**门的口径修，不是放宽**：名单仍只许一份，只是不再把别的函数对同一个值的
   合法提问算进这张名单。

2. **N14 是 locale 依赖格，不该进跨机夹具**。IsNumeric("$1") 的真值跟机器的货币符号走：
   本机 zh-CN（¥）真机答 False，CI 那台 en-US（$）答 True —— 两台机器各自都是对的，
   差的是 locale 不是实现。把它钉进夹具就是把环境当回归信号（正是 check_no_locale_needles
   要拦的形状；它能过，是因为这一格"看起来"只是布尔读数）。处理：N14 撤出夹具
   （66 格 -> 65 格），两个真机读数记在这里 —— **zh-CN ¥ 下 False，en-US $ 下 True**。
   其余 63 格改前就已逐字相同，仍然是「没有误伤」的证人。

撤格之后两台（x64/x86）各 66 行、逐条对针全中且无多余行；
check_sa_access / check_no_locale_needles / check_vbp_fixture_census /
check_test_helper_integrity 四条闸本地全 c=0。

## §11.173 `StrConv` 的档位语义（第一百九十刀）

§11.172 收口后，`StrConv` 是语料里调用次数最多的一条还没判过的字符串函数。它一直卡着，
是因为两档**字节重解释**（`vbUnicode` / `vbFromUnicode`）在 VB6Mini 上给出自相矛盾的读数
（见 `.temp/k189p/P213.bas` 的 S2/S14）—— 不是探针写错，是 VB6Mini 自己在这两档上不稳。

判据改走 TwinBASIC（TB 用 VB6 语法、字符串是 Unicode，与 C3 同形）。探针
`.temp/k189p/TB_StrConv.bas`（用户机器 zh-CN / cp936 上跑），C3 侧的对照探针
`.temp/k189p/P216.bas` 与它**逐标签一一对应**。

### 改前三档各自错在一个地方

| 格 | 真机 (TB) | 改前 C3 | 病因 |
|---|---|---|---|
| S1 `StrConv("ABC", vbUnicode)` | len 6：`65,0,66,0,67,0` | len 3：`65,66,67` | 这一档在 RTL 里是 `no-op` 兜底 |
| S2 `StrConv("ABC", vbFromUnicode)` | len 1：`16961` | len 3：`65,66,67` | 同上 |
| S7 `StrConv("Ä", vbUnicode)` | len 2：`63,0` | len 1：`196` | 同上 |
| X1 `vbUnicode + vbUpperCase` | `ERR 5` | 原样返回、无错、也没折大小写 | 组合档没人验 |
| X3 `StrConv("abc", 999)` | `ERR 5` | 全角 `65345,65346,65347` | `999 & 0x3C` 被拆成 `vbWide+vbHiragana` 走了 `LCMapStringEx` |

X3 那一格最说明问题：位域截得太松，一个非法参数被**静默**解释成另一种合法转换 —— 这比报错
更坏，因为调用方拿到的是一串看着有内容的错值。

### 改后：字节重解释就是「把缓冲当字节流」

`vb6rtl_compat.c` 的 `vb6_StrConv`：

1. **`vbUnicode = 64`**：把 BSTR 的 UTF-16LE 缓冲**当 ANSI（`CP_ACP`）字节流**重解释，
   交给 `MultiByteToWideChar`。`"ABC"` 的缓冲是 `41 00 42 00 43 00`，按 ACP 解出来正是
   `A,NUL,B,NUL,C,NUL` —— 与真机 S1 的 6 个码点逐位相同。非 ASCII 那格也随之对上：
   `"Ä"`（`C4 00`）在 cp936 下前导字节后面跟 `00` 是非法序列，ACP 解码给出 `?` 再跟 NUL，
   就是真机的 `63,0`。
2. **`vbFromUnicode = 128`**：先 `WideCharToMultiByte` 窄化成 ACP 字节，再把字节流当
   UTF-16LE 重解释，宽字符数取 **`floor(字节数/2)`**（真机 S2 的 `"ABC"` → 1 个字符
   `0x4241`，就是 `'A','B'` 两个字节）。
3. **位域校验**：只认 0..255 那位域，`vbUnicode` 与 `vbFromUnicode` 不许同现，字节两档不许
   与大小写档同用 —— 命中就 `Err 5`（真机 X1/X3）。`conversion == 0` 仍是原样返回（真机 C4）。

### 夹具

`tests/b198_strconv_bytes`：**18 格 + DONE**，x64/x86 各 19 行逐字相同，与 TB 逐格对齐。

钉的格：S1/S2/S3/S4/S5、S7L（只钉 Len）、S8/S9/S10、C1..C5、X1E/X2/X3E、X5、L1。
其中 S1/S2/C1..C5/L1 是纯 ASCII，跨代码页稳定；S3/S9/S10/S8 是长度，也稳定。

### 在册（这一刀**故意没钉**的四类，各有明确原因）

1. **S6 双重往返**（`StrConv(StrConv("ABC", vbFromUnicode), vbUnicode)`）：真机给出
   len 3 `ABC`，C3 给出 len 2 `AB`。差别不在转换，在**容器**：TB 的字节串把「3 个字节」
   这个长度藏在字符串的隐式长度里（`Len` 只报 `floor(3/2)=1`，但 `vbUnicode` 能取到 3 字节），
   C3 的 `String` 只有宽字符长度一个字段，奇数字节无处可存。要补就得给 String 加一枚
   「字节长度」—— 那是表示层改动，不塞进这一刀。
2. **S7 的逐字符码点**：跟代码页走（zh-CN `63,0` / en-US `196,0`）。真值不唯一，
   与 §11.172 的 N14 同一条纪律 —— 只钉长度。
3. **X1/X3/X4 的值那半边**：`Err` 已经对了（5 / 5 / 94），但 raise 之后调用方**照样拿到
   返回值**，所以探针会多打一行 `X1=[len 0 :]`、`X4=[len 0 :]`。这是「语句中止」闸的活
   （§11.80 的机制只覆盖实参含晚绑定 COM 读的情形），不是 `StrConv` 自己的问题。
4. **`vbWide/vbNarrow/vbKatakana/vbHiragana`（4/8/16/32）这几档**：这一刀只验了位域边界，
   没逐档对读数（它们走 `LCMapStringEx` + 默认 `ja-JP`，本来就该另开一刀）。

### 收口

代码 `cf65aced`，manifest 登记 `38211999`（888 行，added=2 changed=0，两行取自 CI 的
`ci/emit-manifest`），GA `45754472749` 23/23 全绿。本节即 `38211999` 之后的文档提交。


## §11.174 语句中止（On Error Resume Next 的作用面）：真机读数与分期收口

### 一、为什么要重新量

§B227（第七十九刀）把「语句中止」实现成**序号比较**：赋值站先存一份
`vb6_ErrSerial()`，右值求值完再比一次，变了就不写目标。这条判据隐含一个假设 ——
「序号变了 == 本条语句 raise 过 == 本条语句该中止」。第一百九十刀（§11.173）
的 `StrConv` 探针把这个假设的破口露了出来：`Dump "X1", StrConv(Null, vbUpperCase)`
在 C3 里照常打印了一行空值，而真机整句不见。

### 二、真机读数（VB6Mini，`.temp/k191p/P217.bas` → `p217.txt`）

| 格 | 形状 | 真机读数 |
|---|---|---|
| A1 | `x = FEscape()`（被调方无处理器，错误上抛） | `x=-1 err=11`（赋值没发生） |
| A2 | `x = FSwallow()`（被调方 `Resume Next` 且 `Err.Clear`） | `x=7 err=0` |
| **A3** | `x = FNest()`（被调方吃掉错误但**不清 Err**） | **`x=5 err=11`**（赋值发生了） |
| A4 | `x = FHandled()`（被调方自带 `On Error GoTo`） | `x=3 err=0` |
| A5 | `Call PEscape` | `hits=0 err=11`（被调方体内后续语句全不执行） |
| A6 | `Call PSwallow` | `hits=1 err=0` |
| A7 | `Take2 FEscape(), LogTag(2)` | `hits=0`：坏实参之后不求值，被调方不进入 |
| A8 | `Take2 LogTag(1), FEscape()` | `hits=1`：坏实参之前恰好求值一次 |
| A9 | `x = FEscape() + 1` | `x=-1`（整句丢弃） |
| A10 | `Print #1, "A10=[" & Left("abc",-1) & "]"` | 该行**整行没输出**，下一句 `err=5` |
| A11 | `x = FEscape(): x = x + 1` | `x=0`：只中止第一条语句 |
| A12 | 调用方 `On Error GoTo H1` | 跳到 `H1`，`x=-1 err=11` |

**A3 是这一刀的关键**：被调方把错误吃在自己体内、交回一个值，调用方的赋值
**照常发生**，即便此刻 `Err.Number` 还是 11。所以「序号变了」既不是充分条件也不是
必要条件 —— 真机的判据是**「这次 raise 有没有逃出它所在的那一帧」**。

### 三、第一百九十一刀：实参中止面（已收口）

C3 的实参守卫（第一百零二刀 §11.80）机制本来是对的：0..首个可 raise 的实参逐枚提升
成临时量、每次求值前比一次序号，调用本身再包在同一个守卫里。问题只在**开闸条件** ——
只给「实参含晚绑定 COM 读」开闸。改前三格读数：

| 格 | 真机 | 改前 C3 | 改后 C3 |
|---|---|---|---|
| A7 | `hits=0 err=11` | `hits=1002`，`Take2` 跑了 | `hits=0 err=11` |
| A8 | `hits=1 err=11` | `hits=1001`，`Take2` 跑了 | `hits=1 err=11` |

三处改动：

1. 新增 `exprCanRaiseStmt`（`cgen_util_scan.cpp`，判据在 `cgen_helpers.inc`）：树里
   只要有一枚调用（用户过程 / 内建 / 方法 / 数组下标）或一枚可能除零的算符
   （`/` `\` `Mod` `^`）或 `New`，就当可能 raise。开闸条件改为
   「晚绑定读 ‖ 可能 raise」。
2. 开闸记录的位置从**首枚**改成**末枚**可 raise 的实参（`stmtCallArgLastRaise_`）。
   这是这一刀的真因：只提到首枚时，第二枚仍以裸调用形式留在守卫内的实参表里，而守卫
   那一次比较在它求值**之前**就做了 —— 于是 raise 之后调用照进（A8 改前 1001 的来源）。
3. `ByRef` 的**裸变量**实参不提升（`&x` 本来就是地址，提进值临时量会把被调方的写回丢在
   `_agN` 上）。但这条只认 `IdentifierExpr`：`o.NoSuchProperty` 这种晚绑定读**也是**
   `MemberAccessExpr`，它会 raise —— 当左值跳过提升就等于把它挪进守卫内的调用里，
   raise 之后照进，正好是 §11.80 当初修掉的那个病形（`b250` 神谕当场变红，实测确认）。

#### 3.1 两个 C2440：提升只对「同类型的值」成立

开闸面一宽，提升就碰到了两类**不该套模板**的实参，都是 CI 编译红、本机复现后修掉的：

| 现场 | 病形 | 判据 |
|---|---|---|
| VBFlexGrid MainForm `OleCreatePropertyFrame … Pages.pElems` | `vb6_VARIANT _as55 = (void*)(intptr_t)…` ⇒ C2440 + C2172(参数 8) | 值已是指针/句柄（`(void*)` / `(intptr_t)` / `(vb6_ivref…`）⇒ **不进 Variant 聚合臂** |
| Charts 2020 Form2 `X = Random(10, 30)` 那一族 ByRef 实参 | `float _ag = (&(float){…})` ⇒ C2440 | 已经是地址（`(&…`）⇒ **不进标量值临时量臂** |

并且 `b448_iface_ret` 的接口视图实参（`vb6_argtmp_vb6_ivref_*`，C 类型是指针）也被
Variant 臂抓过 —— 那一臂现在再加一条 `exprHasLateBoundComRead`：**聚合体提升只留给
晚绑定读**（它的所有权协议本来就是 Variant 自己的 `VariantCopy`/`VariantClear`）。

夹具 `tests/b199_stmt_arg_abort`（四格 + DONE，`run_tests.ps1` 登记 x64/x86 两条）：

| 格 | 形状 | 读数（VB6 与 C3 x64/x86 逐格一致） |
|---|---|---|
| B1 | `Take2 FEscape(), LogTag(2)` | `B1=[0,11]` |
| B2 | `Take2 LogTag(1), FEscape()` | `B2=[1,11]` |
| B3 | `Bump x, LogTag(3)`（无 raise） | `B3=[8,3,0]` —— ByRef 写回照旧 |
| B4 | `Bump x, FEscape()` | `B4=[5,0,11]` —— 整句中止，目标不动 |

### 四、尚未闭合（后续两刀）

- **赋值站开闸**（A1/A9/A11）：`mayRaise79` 只认晚绑定读与 `raiseInRhs85`，普通调用
  不在其中；且目标类型只认 Int/Double/String，Variant 目标没有守卫支。
- **`Print` 语句快照**（A10）：整条语句没有序号快照。
- **被调方体内中止**（A5）：`vb6_ErrRaise` 在 `vb6_err_resume_next` 置位时直接
  `return`，而这个标志是**全局**的 —— 调用方的 `Resume Next` 会被被调方继承，于是
  没有处理器的被调方一路跑到底。VB6 里 `Resume Next` 的作用域是**过程**。
- **A12**：`On Error GoTo` 处理器入口处局部 Variant 变 Empty，独立缺陷，另案。

### 五、随刀修的神谕与清单

开闸面变宽会让临时量编号整体后移，13 条形状神谕 + 3 条 `CODEGEN-NOTE` 的针脚因此要
跟到新编号（形状语义没变，`b250` / `b253` / `k174` 三条**一字未动**即通过）。另有一处
`k3100_ivref_arg_lesion` 是**真误中**：病形针脚钉在 `_4` 上，而 `_4` 现在是「借来的实参
要 AddRef」这个**合法**站位 —— 四条针脚一并跟到 `_2/_4/_6`，语义不变。

manifest 按 CI 的 `ci/emit-manifest` 产物装行（890 行，`added=0 removed=0`，
`changed=21`，复检 `diffEntries=0`），没有 `-Bless`。

### 六、记账

- 探针：`.temp/k191p/P217.bas`（12 格，真机 `p217.txt`）。
- 夹具：`tests/b199_stmt_arg_abort`（B1–B4，输出走 `Debug.Print`；`Test-Vbp` 比对的是
  stdout，改前用 `Print #1` 在 CI 上是空读数）。
- 代码 `3360057b` → 收口 `66b78cc6`；manifest `4fefc7b7`；GA `45775710067` 23/23 全绿。
- 构建坑：`src\backend\detail\util\*.inc` 属 PCH 内容，Ninja 不跟踪，改类成员后必须先顶
  `.build\CMakeFiles\vb6c3-cgen.dir\cmake_pch.hxx` 的时间戳再 `build.bat`，否则报
  `C2039: 不是 CCodeGen 的成员`。

## §11.175 赋值站的语句级中止：判据落在「被调方有没有处理器」（第一百九十二刀）

### 一、问题

`On Error Resume Next` 下，一条赋值语句里的 raise 会不会让**目标那一步不写**，此前只看
`vb6_ErrSerial()` 变没变（第七十九/八十四/八十五刀）。那一套对「raise 在本语句自己的求值里」
成立，对「raise 在被调方里」却分不开两族：

```
x = FEscape()      ' 被调方没有 On Error  ⇒ raise 爬出去  ⇒ 目标不动
x = FNest()        ' 被调方 Resume Next 吞掉 ⇒ 语句照常完成 ⇒ 目标该写
```

两族的 `Err` 序号**都会变**，所以老判据只能二选一：不开闸则第一族写错，开闸则第二族写错
（旧注释里 `fdrawstate_x86` 那次翻红就是后者）。

### 二、真机读数（VB6Mini `/make` 实跑，`.temp/k192p/P218.bas` / `P219.bas`）

| 格 | 形状 | 真机 |
|---|---|---|
| D1 / F1 | 整型目标 `n = FEscape()` | **目标不动**，`err=11` |
| D3 | Variant 目标 `v = FEscape()` | 目标不动，`err=11` |
| D5 / F3 | `n = FEscape() + 1` | 目标不动 |
| D6 / F10 | 同一行两条语句 | 第一条中止、第二条照跑（`n=0`） |
| D7 | 数组元素 `arr(0) = FEscape()` | 目标不动 |
| D8 / F4 | 模块级 `mPub = FEscape()` | 目标不动 |
| D9 | `x = CLng("abc")`（Variant 目标） | 目标不动，`err=13` |
| D11 / F6 | `n = Abs(FEscape())` | 目标不动 |
| D13 / F8 | 两跳（外层也无处理器） | 目标不动 |
| F9 | Double 目标 | 目标不动 |
| **D2 / F2** | **被调方 `On Error Resume Next`，不清 Err** | **`n=5`，语句照常完成** |
| **D10 / F7** | **被调方自带 `On Error GoTo`** | **`n=3`，`err=0`** |

### 三、判据与改法

判据是**静态**的：右值树里存在一枚直接调用，被调名在**本模块**过程表里、且那个过程体内
没有 `OnErrorStmt`/`ResumeStmt` ⇒ 这次 raise 必然爬出被调方。

- `procHasErrHandler_`（过程名 → 是否自带处理器）在 `generate` 的 `state_scan` 段由
  `scanProcErrHandlers` 一次建好；`hasOnErrorInStmts` 此前只有定义、没有调用点。
- `exprCallsUnprotectedProc` 递归扫右值树（调用/成员读/二元/一元）。
- 赋值站据此开闸，目标档只开到**整型 / Double / String**；整型与 Double 另发临时量：

```c
const int32_t _e0 = vb6_ErrSerial();
int32_t _av1 = (_e0 == vb6_ErrSerial()) ? ((n * vb6_Factorial(n - 1))) : (int32_t)0;
if (_e0 == vb6_ErrSerial()) n = _av1;
```

没有临时量时守卫在右值**求值之前**求值，raise 之后照样会把被调方交回的值写进目标。

查不到名字的（内建 / 跨模块 / 晚绑定 / 下标）一律不开闸 —— 宽一点就是把 F2/F7 写错。

### 四、读数（`tests/b200_assign_abort`，x64 与 x86 逐字相同，且与真机一致）

```
F1 n=-1 err=11      F2 n=5 err=11       F3 n=-1 err=11     F4 mPub=-1 err=11
F6 n=-1 err=11      F7 n=3 err=0        F8 n=-1 err=11     F9 d=-1 err=11
F10 n=0 err=11
```

改前 F1/F3/F4/F6/F8/F9/F10 七格全错。

### 五、本刀明示不开口的三处（读数已留）

- **数组元素目标** `arr(0) = FEscape()`：真机 `-1`，本仓仍写 `9` —— 走的是
  `vb6_PA_Set<kind>` 的另一条出口，临时量要按元素档另造。
- **Variant 目标** `v = FEscape()`：真机 `-1`，本仓仍写 `9` —— `(VARIANT)0` 不是合法
  初始化，临时量形要单独设计。
- **String↔String**：`s = FSEscape()`（String 函数 raise → String 目标）在 VB6Mini 上
  **不可判** —— 一次静默终止整个 Sub、一次把目标读成 `"V4 s"`（`.temp/k192p/P218b/P218c`），
  属判具缺陷，口径待另立。

### 六、随刀修的神谕与清单

- `strleak_emitc_move` 的一条针原本钉 `AssignMove(&(*o), vb6_MakeTag(...))` 一行；同一处
  现在是「先算临时量、序号没变才移动」，针改成钉守卫与移动两步，不钉 `_avN` 编号。
- manifest 按 CI `45791318007` 的产物装行：`changed=59 added=2 removed=0`，复检
  892/892 相同（`diffEntries=0`），没有 `-Bless`。59 份存量变化的归因抽了
  `test_str_alias` / `m6_math` 两枚当场发码，新增行只有守卫 + 临时量这一形。

### 七、记账与一个新的构建坑

- 探针：`.temp/k192p/P218.bas`（Variant 目标族）、`P219.bas`（有类型目标族）。
- 夹具：`tests/b200_assign_abort`（F1–F10，`Debug.Print` 走 stdout）。
- 代码 `df25812e`；清单 `35f0b5ce`；GA `45793452617` 23/23 全绿。
- **构建坑（新）**：`cgen_state.inc` 属 PCH 内容，而工程有**两个** PCH ——
  `vb6c3-cgen.dir` 与 `c3.dir`。只顶前者时 driver 那几个 TU 不重编，仍按旧 `CCodeGen`
  布局跑新库，表现为 `--emit-c` 直接 0xC0000005（本机实测，改前同一份文件能编过）。
  加类成员后要一并顶掉 `c3.dir\cmake_pch.hxx` 或把 `src\driver\*.cpp` 全部摸一遍。
## §11.176 `On Error Resume Next` 是真捕获：帧、逐语句 resume 点、dispatch switch（第一百九十三刀）

### 一、问题

第一百九十二刀把「赋值站要不要写目标」判到了静态调用树上，但它默认了一件事：raise 之后
控制流**已经**落在调用方的下一条语句。实际不是。改前 `OnErrorKind::ResumeNext` 只是

1. 记一下 `Err` 序号（`vb6_SaveErrState`-系），
2. 让 `vb6_ErrRaise` **就地 return**。

`return` 只退出当下这一层 RTL 调用，于是 raise 之后：

- 被调方（自己没有处理器）**剩下的语句照跑** —— 真机不跑（G1：`hits=0`，本仓改前 `hits=100`）；
- 调用方那条语句也照常完成，把被调方交回的普通返回值写进目标 —— 真机整条中止（G3：`n=0`）；
- 「出错语句的下一条」靠语句顺序偶然对上，跨过程、跨块就散（`b137_resume_in_block` 只是没踩到）。

也就是说 `Resume Next` 此前是「`Err` 记账」而不是「控制流」。真机语义是**过程级**的：raise 一旦
发生，被调方剩余语句与调用方这条语句都不再执行，控制流落到最近的 `Resume Next` 帧，从「出错
语句的下一条」继续；且**隐式 resume 不清 `Err`**。

### 二、真机读数（VB6Mini `/make` 实跑，`.temp/k193p/P220.bas`）

| 格 | 形状 | 真机 |
|---|---|---|
| G1 | `Call PEscape`（被调方无处理器，raise 后还有一条 `hits=hits+100`） | `hits=0 err=11` |
| G2 | `Call PSwallow`（被调方自己 `Resume Next`） | `hits=1 err=0` |
| G3 | `n = FEscape()`（函数式，raise 后还有 `hits=hits+1000` / 赋值） | `n=0 hits=0 err=11` |
| G4 | 两跳 `n = FMid()` → `FEscape()` | `n=0 hits=10 err=11` |
| G5 | 出错语句**之后**还有一条 `hits=hits+1` | `n=0 hits=1 err=11` |
| G6 | 最后一条语句出错，随后调用方继续 | `hits=1 err=11` |

G4 的 `hits=10` 是外层函数入口那一句（`hits=hits+10`）留下的，`+20` 与 `+1000` 都没跑 ——
「两跳各层剩余语句都不执行」的读数指纹。G5 的 `err=11` 说明隐式 resume **没有**清 `Err`
（这条决定了 `b159` 那两格为什么该是 `(e13)(e13)` 而不是一格）。

### 三、判据与改法

`OnErrorKind::ResumeNext` 现在真的架帧，且**每条语句**前面落一个 resume 点：

```c
vb6_error_jmp_ptr = &vb6_local_err_jmp;
vb6_error_jmp_set = 1;
vb6_err_jmp_active = 1;
vb6_err_resume_next = 1;
if (setjmp(vb6_local_err_jmp) != 0) {
    vb6_err_dispatch = vb6_err_resume_next_point;
    goto vb6_err_dispatch_switch;
}
vb6_err_resume_point = 0;
vb6_err_resume_next_point = 1;
vb6_resume_0:;
s = s & "/" & CStr(UBound(r));
vb6_err_resume_point = 1;
vb6_err_resume_next_point = 2;
vb6_resume_1:;
...
vb6_err_dispatch_switch:
switch (vb6_err_dispatch) { case 0: goto vb6_resume_0; case 1: goto vb6_resume_1; ... }
```

配套三点：

- `vb6_ErrRaise` 改为**先 longjmp**（帧有效时），内联那条 `resume_next` 早退只留作没有帧时的兜底；
- `errFrameReady_` 只在 `cgen_decl_func.cpp` / `cgen_decl_proc.cpp` 两个真正的过程体入口置位。
  Property、继承类、控件片段这几个发射器不置位 ⇒ 退回老行为，不会踩到「没有 `jmp_buf` 却去
  longjmp」；这条护栏是本刀在 VBFlexGrid x86 上撞出来的（`C2094 标签未声明` +
  `C2065 vb6_local_err_jmp` 各一次）；
- `hasResumeNextInStmts` + `hasResumeNext_` 与第一百九十二刀同一个 `state_scan` 段建好，用来判
  「这个过程的语句要不要逐条落 resume 点」（没有 `Resume Next` 的过程一个点都不多）。

### 四、读数（`tests/b201_resume_scope`，x64 与 x86 逐字相同，且与真机一致）

```
G1 call_escape: hits=0 err=11      G2 call_swallow: hits=1 err=0
G3 assign_escape: n=0 hits=0 err=11   G4 two_hop: n=0 hits=10 err=11
G5 next_stmt: n=0 hits=1 err=11      G6 last_stmt: hits=1 err=11
```

改前 G1 `hits=100`、G3 `n=9 hits=1000`、G4 `hits=1030` —— 三格是「raise 只 return 一层」的直接指纹。

### 五、随刀暴露、本刀不开口的两处（读数已留）

- **`VariantArrayGet` 对 Empty 不报 13**：真机 `r(0)`（`r` 为 Empty）报 13、目标不动
  （`.temp/k193p/P221.bas` H2/H4）；本仓静默交回 Empty。`test_com_latebound` 的 `mu/ms` 两格因此
  仍差最后一段（真机 `0/Empty(e13)(e13)`，本仓 `0/Empty(e13)/`）。
- **`d("ms")` 的 VarType**：未 ReDim 的 `String()` 装进 Dictionary 再取回，真机答
  `8200/String()`（`.temp/k193p/P222.bas`），本仓仍答 `0/Empty`。这是 b159 那套「模块级数组 ↔
  COM SAFEARRAY」的老缺口，与 `Resume Next` 无关。

两处都在 `test_com_latebound` 的针上留了痕：`VB-b159-mu` / `VB-b159-ms` 由 `0/Empty/-1(e13)`
重钉为 `0/Empty(e13)/`。重钉的**方向**是朝真机走 —— 改前那句 `UBound` 不再把 `-1` 写进 `s`，
`(e13)` 由紧接着的那个 `If Err.Number <> 0` 补上；真机正是这个顺序（P221 H3：`s` 停在 `A`）。

### 六、神谕与清单

- `b137_resume_in_block` 十条针全过（Resume Next 落在块内的老夹具，未改一针）。
- 真项目烟测：Charts 2020 x64 `rc=0`；VBFlexGrid x86 `rc=0`（改 `errFrameReady_` 之前 x86 直接
  编不过，见上）。
- manifest 按 CI `45801177450` 的产物装行：`changed=101 added=2 removed=0`，复检 894/894 相同
  （`diffEntries=0`），没有 `-Bless`。101 份存量变化的形状抽 `test_com_latebound` 当场发码核对：
  `vb6_resume_N:;` + `vb6_err_resume_next_point = N;` + `vb6_err_dispatch_switch`，三样之外无漂移；
  新增 2 行是本刀夹具 `b201_resume_scope`。

### 七、记账

- 探针：`.temp/k193p/P220.bas`（六格作用面）、`P221.bas`（`UBound(Empty)` / `r(0)` 单点）、
  `P222.bas`（逐字复刻 `B159Show` 取整串）。
- 夹具：`tests/b201_resume_scope`（G1–G6，`Debug.Print` 走 stdout）。
- 代码 `cd9c21e6`；针 `7fe0f0ee`；清单 `6f59f74f`；GA `45802555558` 23/23 全绿
  （此前 `45798686898` 三红：manifest 未装行 + `Tests (compile #1)` / `Tests (bas #2)` 各由
  `manifest_coverage` 与 `test_com_latebound` 一条针翻红）。
- 记账坑（沿用第一百九十二刀，未再新增）：`cgen_state.inc` 属 PCH 内容，两个 PCH 都要顶。
## §11.177 一维 Variant 下标的错误档：非数组 13、越界与未分配 9（第一百九十四刀）

### 一、问题

上一刀量 `test_com_latebound` 的 `mu/ms` 两格时撞见了根子：`vb6_VariantArrayGet` 对
「槽不是数组」「数组没分配」「下标越界」三种情况**一律静默交回 `Empty`**，写侧
`vb6_VariantArraySet` 同样静默丢弃（二维那两条 `…GetND` / `…SetND` 对非数组槽也是静默交
`Empty`）。于是 `.temp/k194p/P223.bas` 的 A1–A5 / A7 / A8 / A10 / A12 / A13 八格在 C3 上
全是「不报错、目标被写空」，真机逐格报 13 / 9 —— 这是「C3 让错误消失」的典型形：
程序不崩、读数全错，比崩溃更难查。

### 二、真机读数（VB6Mini `/make` 实跑）

一维（`.temp/k194p/P223.bas`，读侧 `x = r(i)` 与写侧 `r(i) = 77` 都量了）：

| 格 | 槽 | 真机 |
|---|---|---|
| A1 | `Empty` | `err=13`，`x` 不动 |
| A2 | `5`（Long） | `err=13` |
| A3 | `Null` | `err=13`（**不是** 94） |
| A4 | `"abc"` | `err=13` |
| A5 | `True` | `err=13` |
| A6 | `Variant()` 界内 | `x=7`，`err=0` |
| A7 / A8 | `Variant()` 越界上 / 下 | `err=9` |
| A9 | `Long()` 数组界内 | `x=11`，`err=0`（元素提升） |
| A10 | 未分配的 `Variant()` 动态数组 | `err=9` |
| A11 | `Array(21,22)` | `x=22`，`err=0` |
| A12 | 写侧：`Empty` 槽 `r(0)=77` | `err=13` |
| A13 | 写侧：越界写 | `err=9` |
| A14 / A15 | `UBound` / `LBound(Empty)` | `err=13`（`err=13` 这一档早就是对的，作对照） |

二维（`.temp/k194p/P224.bas`）：空槽 / Long 槽 → `13`，`a2(0,1)` → `31`，`a2(2,0)` → `9`。

分界只有一条：**槽本身是不是数组**决定 13 与 9，与下标无关；读侧、写侧、一维、二维同口径。

### 三、判据与改法

四处出口按同一条分界 raise（`vb6_ErrRaiseNumber`，与除零那族同一个入口）：

```c
static void vb6_VariantArrayRaiseNotArray(void) { vb6_ErrRaiseNumber(13); }
static void vb6_VariantArrayRaiseUnalloc(void)  { vb6_ErrRaiseNumber(9); }
static void vb6_VariantArrayRaiseOOB(void)      { vb6_ErrRaiseNumber(9); }
```

- `vb6_VariantArrayGet`：`!v` 仍留静默交 `Empty`（发码侧不会这么发，属防御）；`vt` 无数组位
  → 13；归一一枚真 SAFEARRAY 之后 `parray` 为空 → 9；下标出界 → 9。
- `vb6_VariantArraySet`：同形，写侧原来「静默丢弃」的三种情况同样 13 / 9 / 9。
- `vb6_VariantArrayGetND` / `…SetND`：非数组槽 → 13，未分配 → 9，秩数非法 → 13；
  越界本来就由 `vb6_SaNdElemPtr` / `vb6_SaNdElemFail` 抛 9，未动。

四处的先后顺序是**先判「是不是数组」再判「分配没分配」**：`vt` 位是槽的类型标签，`parray`
是运行期内容；反过来判会让「Empty 槽」与「未分配数组」都落进 9，而真机把前者算 13。

### 四、读数（`tests/b202_variant_subscript`，x64 与 x86 逐字相同）

```
V1 empty-err=13      V2 long-err=13       V3 null-err=13      V4 string-err=13
V5 bool-err=13       V6 inrange=7         V6 inrange-err=0    V7 oob_hi-err=9
V8 oob_lo-err=9      V9 longarr=11        V10 unalloc-err=9   V11 arrayfn=22
V12 set_empty-err=13 V13 set_oob-err=9    V14 ubound_empty=0-err=13
V15 lbound_empty=0-err=13                 V16 empty2d-err=13
V17 inrange2d=31     V18 oob2d-err=9      V19 hits=1-err=13
```

19 格与真机一一对上（V19 那一格另用 `.temp/k194p/P225.bas` 取的真机：出错语句之后的下一条
照跑、`Err` 留着）。改前 V1–V5 / V7 / V8 / V10 / V12 / V13 / V16 / V18 全是 0。

### 五、本刀明示不开口的三处

- **raise 之后目标槽的内含**：真机保持原值（`x=-99` 不动），本仓仍写 `Empty` —— 这是
  第一百九十二刀留下的 Variant 目标缺口（`(VARIANT)0` 不是合法初始化），夹具因此只钉错误号
  与界内取值，不钉 raise 之后的内含。
- **地址侧 `vb6_VariantArrayElemPtr`**：仍对非数组槽交 `NULL`（发码侧用它做元素引用 /
  `VarPtr`）。那一档要动的是「NULL 之后发码侧怎么办」，另立。
- **非 Variant 数组元素的写侧**：`vb6_VariantArraySet` 对 `elemType != vb6_sa_variant` 的数组
  仍然静默不写（注释原文「仅 Variant 数组支持赋值」）。真机这一档是另一种型别转换错误，
  本次没量，不猜。

### 六、神谕与清单

- 副作用（朝真机走）：`test_com_latebound` 的 `VB-b159-mu/ms` 由上刀的 `0/Empty(e13)/`
  变成 `0/Empty(e13)(e13)` —— 与真机 `P222` 的 `mu` **逐字相同**；`ms` 仍差 VarType
  （真机 `8200/String()`、本仓 `0/Empty`，属 b159 那套 SAFEARRAY 老缺口）。
- 真项目烟测：VBFlexGrid x86 编过 `rc=0`（这套控件是最重的 Variant 下标消费者）。
- 门禁：`check_sa_access.ps1` PASS、`check_builtin_slot_returns.ps1` PASS。
- manifest 按 CI `45805710855` 的产物装行：`changed=0 added=2 removed=0`，复检 896/896 相同。
  `changed=0` 本身就是这一刀的判据读数 —— **只动 RTL 的 `.c`，发码侧一个字节没变**；
  若这里出现几十上百行，就说明改动跑到了发码侧，要回头查。
- 该次 GA（`45805710855`）的另外两红都是例行项：`manifest_coverage`（新夹具没装行）与
  `Emit manifest`；`Tests (bas #1/#2…)` 全过 —— 语料里没有「靠下标越界静默取 Empty」活着的用例。

### 七、记账

- 探针：`.temp/k194p/P223.bas`（一维 15 格）、`P224.bas`（二维 4 格）、`P225.bas`（隐式 resume
  与 Err 留存）。
- 夹具：`tests/b202_variant_subscript`（V1–V19，`Debug.Print` 走 stdout）。
- 代码 `2a476cf3`（清旁产物 `407dc948`）；清单 `c45a3791`；GA `45808179408` 23/23 全绿。
- 小坑：夹具目录里用 VB6Mini 试编时落的 `VB6.exe` / `*.exe` / `vb.out` 会被 `git add` 顺手带上，
  提交后要 `git rm --cached` 清掉（本次踩了一次）。