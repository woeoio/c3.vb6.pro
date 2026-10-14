# 026-CoClass 与 COM 暴露计划书

> 日期：2026-09-24
> 状态：**设计已拍板（用户按默认套餐批准），未开工**
> 规范输入：`ai/018`（外部讨论存档）+ `ai/022` 的 D1/D7/D8（P0 设计记录）。本文**取代** 022 里关于 P5 的那几段作为实施依据；022 的 B11/B12 行待下一轮由该任务自己加指向本文的链接（本轮不编辑 022，它有活跃写者）。
> 一句话结论：**新写只认 `CoClass` 块（路 1）；VB6 存量 attribute 作为只读折算入口（路 2 降级），两者汇入同一个身份求解函数**。P5 只交付"组内那半"，"对外可用"整体押到 P6 的 B15/B17 之后，并在状态里分栏标明。

---

## 一、目的与两半（先钉状态口径）

| | (a) 组内契约聚合 | (b) 对外 COM 服务器 |
|---|---|---|
| 内容 | `CoClass` 块、`New <CoClass>`、`As <CoClass>`、契约完整性校验、ProgID 编译期映射 | 外部进程 `CreateObject`/`CoCreateInstance`、类工厂、注册、类型库 |
| 归属批次 | **P5 = B11/B12**（本文） | **P6 = B13–B17**（022） |
| 新增代码 | 纯前端 + 编译期改写 | 主要是**改造既有底子**（见三节） |

**状态位口径（重要，防误读）**：总表里 CoClass 一律用两栏标注 —— `组内可用` / `对外可用`。**B12 收口只点亮"组内可用"**；在 B15（类型库 dual/`TKIND_INTERFACE`）之前，对外部语言而言这些类等于不存在。谁把 B12 当成"CoClass 做完了"，就会在验收时踩空。

## 二、语法模型：主块 + 存量只读折算

### 实测数据（本轮 `git grep`，用于纠正一个曾被我误判的风险）

| attribute 键 | 语料出现次数（`tests/*` + `archive/*`） | 含义 |
|---|---|---|
| `VB_Createable` | **0** | 可创建性 |
| `SingleUse` | **0** | 单用/多用 |
| `Instancing` | **1** | VB6 的 Instancing 枚举 |
| `VB_Exposed` | 199 | 是否暴露给外部 |
| `VB_GlobalNameSpace` | 198 | 全局命名空间 |
| `MultiUse` | 209 | 线程模型 |

结论：**"开始消费身份类 attribute 会大面积改变存量产物"这个担忧不成立**（身份键在语料里近乎为零）。我（该会话）在讨论中一度把它当作反对路 2 的理由，此处撤回并留档，避免下一个人重走。真正决定选路 1 的是下面三条，与语料量无关。

### 选路 1（显式块）的三条理由

1. **tB 兼容**：本项目先例一律按 tB 语义适配（Delegate `a15c40b`、Overload `ea0591b`、Generics `7e2f9ed`、Interface/Inherits 022 B01–B08）。tB 的 CoClass 是显式块；走路 2 = 在这个特性上单独分叉。
2. **契约聚合需要"一个位置"**：默认接口 + 接口集合 + ProgID + 可创建性 + 自定义构造器 = 5~8 个扁平字符串键，拼错无诊断、读不出哪条是默认。而块形式的**解析成本已接近零**——属性行解析器现成（`src/parser/parser_interface.cpp:41` 的注释即"CoClass 系列先接受语法，由 P5/P6 消费"），AST/语义层两种方案都得写。
3. **实现与契约解耦**：路 2 下"类的 Public 成员就是对外契约"= 退回 VB6 的"全都暴露"。路 1 允许"对外只有 `IShape`，实现类另有私有成员"，这是 CoClass 的存在意义之一。

### 折算规则（路 2 的残余价值，纯读、不发码）

- 迁移工程里读到 `Attribute VB_Createable` / `Instancing` / `VB_Exposed` 时，**折算成同一条内部 CoClass 记录**，并给一条 info 级提示建议改写成块。
- **生效条件严格限定：该 `.cls` 没有任何显式 `CoClass` 块。** 有块时 attribute 一律忽略 + 警告 —— 于是"同一个类两个身份来源"在结构上不可能发生。
- 折算可以单独成批（纯诊断、零发码），甚至并进 B11。
- 现状：C3 今天**只消费 `Attribute VB_Name`**（`src/driver/driver_frontend.cpp:219-222`、`src/project/frm_parser.cpp:300-306`），其余 attribute 全是"存着没人读"，所以折算是一条全新的读取路径，不影响任何现有行为。

## 三、身份求解：唯一一处函数（不可让步）

CLSID / IID / ProgID 三者都必须经过**一个**函数求解，任何消费者（代码生成、dll_entry 表、类型库、注册器）只许调它。禁止三处各读一遍。

优先级（沿用 022 D8，扩展 ProgID 一栏）：

```
CLSID:  显式 [CoClassId("{…}")]  >  vbp 三段式 Class=Name; x.cls; {CLSID}  >  FNV-1a 确定性("coc:<Proj>.<CoCls>")
IID:    显式 [InterfaceId("{…}")] >  FNV-1a 确定性("itf:<Proj>.<Iface>")
ProgID: 显式 [ProgId("…")]        >  默认 "<vbp Name>.<CoClass名>"
```

- **禁止随机 GUID**：类型库与 `.tlb` 资源必须每次构建可复现（022 D8 原话）。
- 现有底子：vbp 三段式解析在 `src/project/vbp_parser.cpp:74-105`、消费在 `src/driver/driver_compile.cpp:76-79`（`classClsidMap_`）；确定性 GUID 生成在 `cgen_util_dllentry_prelude.inc:39-70`。**这些行号引自 022 D7/D8，开工前须复核是否漂移。**

## 四、块的语法草案

```vb
CoClass Circle
    [CoClassId("{2E1B5C40-8A5F-4C2A-9E2D-6D6A0F2B9C11}")]
    [ProgId("Shapes.Circle")]
    [ComCreatable(True)]
    [Implementation("CircleImpl")]

    Interface IShape                       ' 契约集合
    [Default] Interface ICircle            ' 默认接口，必须 ∈ 集合
End CoClass
```

- 宿主：`.bas` 与 `.cls`（022 D1 已定：`Interface`/`CoClass` 作模块级声明块，名字**工程级唯一**；`.frm/.ctl` 不接受）。
- 关键字 `CoClass` 走 022 D1 的四词登记 + **软关键字双登记**；`End CoClass` 定界同 `End Interface`。
- `[Implementation("…")]` 是**唯一**的实现类绑定形式（不做同名隐式、不做 `Implements` 反查 —— 隐式规则在跨模块必然歧义）。v1 一个 CoClass 绑**一个**实现类。
- `[Default, Source] Interface …` 只接受并存档，**不实现连接点**（用户已确认的红线）。

## 五、语义与校验（B11 的全部职责）

1. **契约聚合**：块内列出的每个接口，绑定的实现类必须满足其契约（复用 B02 的 `checkNewStyleInterface` 与 `IfaceRegistry`，不新造比对器）。
2. `[Default] Interface` 必须 ∈ 接口集合；集合内接口名去重。
3. `As Circle`（CoClass 名）= **默认接口视图**；取其他接口必须 `As IShape` 或经 QI。
4. `New Circle` = 编译期直调实现类工厂（不经注册表）。
5. `CreateObject("Shapes.Circle")` 且 ProgID **命中本工程 CoClass** → **编译期改写**成 `New <实现类>`；不做运行期查表。理由：今天 `CreateObject` 走真实注册表链路（`src/rtl/core/vb6com/vb6com.c:509-560`），组内工程根本没注册，运行期查表只会失败或链到机器上另一个同名组件。ProgID 与外部已注册组件同名时给**警告**。
6. **拒绝清单**（诊断 ASCII）：`Inherits` 一个 CoClass（v1 拒，照 3022 那批同族处理）；EXE 工程里 `[ComCreatable(True)]`（VB6 的 EXE 不能注册为本地服务器；EXE 只享受组内那半）；`[Implementation]` 指向不存在/不是类的符号；接口集合里有 legacy（VB6 式 `.cls` 当接口）成员。
7. **零新语法护栏**：没有 `CoClass` 块、也没有被折算 attribute 的工程，产物必须逐字节不变（8 文件 `--emit-c` 清单沿用 022 的仪式）。

## 六、批次

| 批 | 内容 | 对应 022 | 验收 |
|---|---|---|---|
| C01 | `CoClass…End CoClass` 块语法 + 属性行落到 AST（**不校验、不发码**） | B11 前半 | 单文件 `--syntax-only` 通路（D15-5 已三次验证），正例 + 畸形子句负例 |
| C02 | **身份求解唯一函数**（CLSID/IID/ProgID 三优先级）+ 可复现性测试 | B11 | 同一输入两次构建 GUID 全等；显式/vbp/确定性三档各一条断言 |
| C03 | 契约聚合校验 + 拒绝清单 + `As <CoClass>` 语义 | B11 | 缺接口/默认接口不在集合/`Inherits` CoClass/EXE+ComCreatable 四类负例 |
| C04 | **存量 attribute 折算**（纯读 + info 提示；有块则忽略+警告） | 本文新增 | 一份带 `Attribute VB_Createable` 的 `.cls` 折算成与手写块**同一条**内部记录 |
| C05 | 组内激活：`New <CoClass>` + `CreateObject(ProgID)` 编译期改写 + 默认接口派发 | B12 | 端到端 vbp 工程，断言 `CC1..CCn`；x86+x64 双跑 |
| → | 对外那半（工厂/IUnknown/IDispatch/类型库 dual/注册/外部激活） | **B13–B17，留在 022** | 不在本文范围 |

门与纪律全部沿用 022：`scripts/dev.ps1 -SkipTest` + `tests/run_tests.ps1 -Category all` 零新增失败；用例文件 GBK+CRLF；诊断 ASCII；改发码路径的批跑逐字节护栏。

## 七、与既有债务的接口点（开工前必读）

1. **D15-8**：新式接口名今天仍会进 `classSym->implementsNames`，ActiveX DLL 的 `cgen_util_dllentry_tables.inc:88-130` 会按 legacy 口径 mint IID → **B13/B16 必须在此分叉**，而分叉的判据就是三节那个唯一求解函数。
2. **D22-7④**：QI 认 `IUnknown` 目前返回本接口薄指针，P6 统一 —— CoClass 对外之前必须收掉这条。
3. **字段定序（D33，B08d 刚改）**：`__comObj → __iv_<I>… → __cvtbl → 祖先字段 → 自有字段`。022 D7 那句"背后绑定一个私有 `.cls` 实现类"写于 B04 之前，**须按当前布局重钉**（尤其 CoClass 视图若涉及 `container_of` 减法，偏移口径变了）。
4. **B15 是对外可见的前提**：`src/typelib/typelib_builder.cpp` 今天只会 `addDispInterface`（TKIND_DISPATCH）+ `addCoClass`，导不出 `TKIND_INTERFACE`/dual，且 `CreateTypeLib2(SYS_WIN64)` 写死（`typelib_builder.cpp:117`）。CoClass"对外可用"最迟卡在这一条。
5. 现成可复用的对外底子（优先扩展不新造，022 D8）：`src/rtl/core/vb6comserver/`（`_obj` 完整 IDispatch+QI+原子计数、`_factory` IClassFactory、`_cp` IConnectionPoint、`_pci` IProvideClassInfo2、注册表写入）、`cgen_util_dllentry_exports.inc` 的 DllGetClassObject 一族、`activex_dll.def`（`driver_link.cpp:207-225` 生成）。

## 八、硬边界（不做）

dispinterface 定义；`[Default, Source]` 连接点实现；一个 CoClass 绑多实现类；泛型类做实现类；把 CoClass 编成二进制对外分发以外的形态（与 023 的二进制边界一致）；`Inherits CoClass`。

## 九、待定时点（不挡 C01–C03）

- `[CoClassCustomConstructor]`（ByRef 实例入参 + 返回 HRESULT 的工厂过程）落在 C05 还是 B13 —— 取决于带参构造要不要先解决（见 [[constructor-feature-gap]]：**带参构造目前无排期**，所以自定义构造器极可能被迫押后，需在 C05 前确认）。
- CoClass 与非 ActiveX 的 `Sub Main` 启动共存规则（EXE 工程既然不能 ComCreatable，是否允许块存在但只用于组内——C03 的默认答案是允许）。
