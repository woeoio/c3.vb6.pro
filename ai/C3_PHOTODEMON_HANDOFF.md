# PhotoDemon 编译缺口交接（C3_PHOTODEMON_HANDOFF）

独立于其它工程的交接文档（不要并入 ucTreeView / ucSlider / 分段控件那几份）。

- 目标：完善 C3 编译器，最终能用 `--arch x86` 成功编译 `D:\vb_yqt4qPac\PhotoDemon-main\PhotoDemon.vbp`。
- VB6 参考 exe：`D:\vb_yqt4qPac\PhotoDemon-main\PhotoDemon.exe`（图标/界面基准，仅作对照）。
- 节奏约束：单进程串行构建、低并行度；不急于编过，重点是补齐语法 / 控件方法属性 / 类型与语义缺口；每类缺口单独提交。

## 1. 工程规模与依赖（只读摸底）

- 源码文件 592 份，合计约 20.4 MB：`.frm` 232、`.cls` 165、`.bas` 139、`.ctl` 56。
- `PhotoDemon.vbp` 的外部引用只有两条：`shell32.dll`（Microsoft Shell Controls And Automation）与 `stdole2.tlb`（OLE Automation）。
- 无第三方 OCX：界面全部是工程内自绘 UserControl（`Controls\pd*.ctl`，56 枚）。
- 主要模块：`Modules\`（139 份 .bas）、`Classes\`（165 份 .cls）、`Forms\`（232 份 .frm，含 `Forms\Options\`）。

## 2. 当前状态 / 未提交改动（2026-11-09 记录）

记录时的工作区（**这是当时 `git status` / `git diff --stat` 的原样产出**）：

```
$ git status --short
 M src/lexer/lexer_literal.cpp
 M src/parser/parser.hpp
 M src/parser/parser_decl_var.cpp
 M src/parser/parser_module.cpp
 M src/parser/stmt/parser_stmt.cpp
 M src/parser/stmt/parser_stmt_assign.cpp
?? .zcodeignore

$ git diff --stat
 src/lexer/lexer_literal.cpp            | 20 ++++++++++++++------
 src/parser/parser.hpp                  |  4 ++++
 src/parser/parser_decl_var.cpp         |  6 +++++-
 src/parser/parser_module.cpp           |  4 +++-
 src/parser/stmt/parser_stmt.cpp        |  2 +-
 src/parser/stmt/parser_stmt_assign.cpp | 26 +++++++++++++++++++++++++-
 6 files changed, 52 insertions(+), 10 deletions(-)
```

分支：`feat/photodemon-x86`（与 `dev` 同点，0 ahead / 0 behind）。

### 2.1 六个改动的归类与完整度

这一批是**同一类**：面向 PhotoDemon 语料的「解析器语法补口」，彼此不冲突，构成一个可编译的完整批次（不是半截）。

| 文件 | 补的语法 | 类别 | 完整度 |
| --- | --- | --- | --- |
| `src/lexer/lexer_literal.cpp` | 注释行以「空白 + `_`」结尾时延续到下一行（注释吞掉后续物理行） | 词法/注释 | 完整 |
| `src/parser/parser_module.cpp` | `Option Compare Binary` 同时接受 `TokenKind::Binary2` 与 `TokenKind::Binary` | 声明语句 | 完整 |
| `src/parser/parser_decl_var.cpp` | 定长字符串 `String * N` 的长度表达式改用 `parseExpression(10)`，避免把 `Const X As String * 3 = "pdm"` 的 `=` 当比较吃掉 | 类型/常量声明 | 完整 |
| `src/parser/parser.hpp` | 新增 `parseReDimStmtList()` / `parseReDimEntry(loc, preserve)` 声明 | 声明（配套） | 完整（与下两项同批） |
| `src/parser/stmt/parser_stmt.cpp` | 语句入口 `ReDim` 改走 `parseReDimStmtList()` | 语句分派 | 完整 |
| `src/parser/stmt/parser_stmt_assign.cpp` | 支持 `ReDim a(...), b(...)` 多数组，及 `ReDim a(i).b.c(n)` 成员链继续 | 语句解析 | 完整 |

- 编译状态：本地 `.build` 的增量构建（`build.bat`）返回 `ninja: no work to do`，说明这 6 份改动已全部进入当前 `.build/C3.exe`（源文件最后修改 11:54–11:56，C3.exe 生成 11:56:21），即**这批改动自身能编译通过**。
- 未完成的验证：这批改动**尚未**在 PhotoDemon 上跑过 `--syntax-only`，也**尚未**跑回归基线（现有测试 + GA）。因此「完整」只指其自身可编译，不表示已验证无回归。
- `.zcodeignore` 是未跟踪的工具忽略文件，不属于编译器改动。

### 2.2 中断说明

- 上一轮计划的 `--arch x86 --syntax-only` 全工程检查**未执行**（进程未启动即被取消），因此**目前没有任何 PhotoDemon 错误分类清单**，本文件不含错误统计。
- 用户机器因负载过重黑屏后，已按要求停止全部构建/测试；当前无残留 `cl.exe` / `link.exe` / `C3.exe` / `ninja` / `cmake` / `msbuild` / `PhotoDemon.exe` / python 进程。

## 3. 低负载分批检查（2026-11-09 起的新节奏）

硬约束：同一时刻只允许一个构建/测试进程；禁止整工程一把梭；每次只跑 1 个或少量文件，跑完消化再下一批；每批之间确认上一进程已退出；C3 只做增量构建，不 clean/rebuild。

### 批次 1：`Modules\GDIPlus.bas`（243 KB，工程最大 .bas）

- 前置：空闲内存 9.8 GB / 共 16 GB（占用 39%），无 `C3.exe` / `cl` / `link` / `ninja` / `cmake` / `msbuild` 残留。
- 命令：`C3.exe <该文件> --syntax-only --arch x86 --output-dir .build\pd_out`
- 读数：`EXIT=0`，耗时 **1.6 s**，峰值工作集 **50.5 MB**，stdout/stderr 均为 0 行。
- 结论：0 错误。该文件单独可过语法检查。
- 负向对照（探针，`.build\pd_probe\ProbeBad.bas`，3 行故意坏语句）：`EXIT=1`、0.2 s、6.7 MB，stderr 报 3 条 `VB2004/VB2003` → 说明 `--syntax-only` 的报错走 stderr，批次 1 的 0 错误不是「只解析文件头」的假绿。
- 成本推算：单文件约 1.6 s / 50 MB ⇒ 592 份串行约 16 分钟、内存峰值约 50 MB 量级；**实际仍按小批（建议 5–10 份/批）执行，不整工程连跑**。

### 批次 2：5 份大 .bas（Palettes / Plugin_FreeImage / Processor / Plugin_OpenJPEG / DibSupport）

- 读数：5 文件，总耗时 **4.4 s**，峰值 **49.5 MB**，退出码全 0，错误 **0**。
- stderr 各 1 行 `C3: AddressOf callback procs: N (...)`，是**提示**不是错误（Plugin_FreeImage 1 枚、Plugin_OpenJPEG 7 枚 AddressOf 回调）。

### 批次 3：主窗体 + 大 .bas + .cls + .ctl

- 文件：`Forms\MainWindow.frm`、`Modules\ImageExporter.bas`、`Modules\Main.bas`、`Classes\pdImage.cls`、`Controls\pdButton.ctl`
- 读数：5 文件，总耗时 **3.4 s**，峰值 **43.0 MB**，退出码全 0，错误 **0**。
- stderr 仅 1 行 CoClass identity 提示（pdImage），不是错误。

### 错误分类清单（累积）

| 类别 | 数量 | 代表错误（文件:行 + 原文） | 同根因说明 |
| --- | --- | --- | --- |
| 语法 | 0（批次 1–3，共 11 份） | — | — |
| 控件属性/方法缺失 | 0 | — | 见下方「口径发现」：`--syntax-only` 不产生这一类 |
| 类型不匹配 | 0 | — | 同上 |
| 缺失内置函数/常量 | 0 | — | 同上 |
| 其它 | 0 | — | — |

### 口径发现：`--syntax-only` 只做词法/语法，不做语义（2026-11-09）

- 探针 `.build\pd_probe\ProbeSem.bas`（`NoSuchBuiltin(1)` + `foo.Bar(2)` + `Call NoSuchSub`）单跑 `--syntax-only`：**EXIT=0、0 条报错** ⇒ 未定义函数/标识符/子程序都不会被报出。
- 同一探针跑 `--emit-c`：EXIT=0，报 2 条 `warning VB3001: 未声明的标识符`（`foo`、`NoSuchSub`）⇒ **语义检查只在 `--emit-c` / 真编译路径上发生**。
- 单文件 `--emit-c` 代价实测（`Modules\GDIPlus.bas`，243 KB）：**4.2 s / 63.4 MB**，stdout 0.3 MB，stderr 134 行**全部是 VB3001**（跨模块符号缺上下文导致的噪声，无真实缺口）。
- 结论：按文件跑 `--syntax-only` 已无法产出用户要的五类分布（预计全工程都是 0 语法错）；单文件 `--emit-c` 又被跨模块 VB3001 噪声淹没。**要拿到控件属性/方法、类型、内置缺失这三类真实缺口，必须带工程上下文（`PhotoDemon.vbp`）跑一次 `--emit-c` 或真编译**，这是下一步需要拍板的口径。

## 4. CI 探针：PhotoDemon x86 emit（分支 `probe/photodemon-emit`）

- 用户拍板：**整工程 emit 不在本地跑，改由 GitHub Actions 跑**，本地只取结果。`dev` 与 22 job 主链路未改动。
- 源码入仓：`tests/photodemon/`，**643 文件 / 18.5 MB**（排除 `Support\` 支持工具与重复的 `Resources\Assets\PD_icons.RES`；保留 vbp 引用的 `Resources\PD_icons.RES`，用 `git add -f` 入库）。vbp 外部引用原本是 `..\..\Windows\SysWOW64\...` 相对路径，在 CI 无法解析（报 2 条 VB4001），已把夹具里的两条 `Reference=` 改成 `C:\Windows\SysWOW64\...` 绝对路径（仅夹具适配，不动编译器）。
- workflow：`.github/workflows/pd_emit_probe.yml`，单 job（不拆矩阵），复用主链路的 VS 环境导出 + cmake/ninja 构建，然后 `C3.exe tests/photodemon/PhotoDemon.vbp --arch x86 --emit-c`，stdout/stderr 全量落盘并上传 artifact `photodemon-emit-logs`，job summary 输出按诊断码的归总计数。
- 触发：`workflow_dispatch` 可用，但 GitHub 要求 workflow 已存在于默认分支；本文件只在探针分支，故另加**仅限本分支**的 push 触发（不挂进 ci.yml）。
- **run 42581367618**：全部步骤绿，总耗时 **26m10s**；emit 本身 **1457.6 s（24m18s）**，stdout **49,057,097 字节**（生成的 C），stderr **4405 行**，**EXIT=0**。artifact：`photodemon-emit-logs`。

### 4.1 诊断归总（run 42581367618）

| 诊断码 | 级别 | 条数 | 原文样例 |
| --- | --- | --- | --- |
| VB3001 | warning | 3693 | `未声明的标识符: 'g_Language' (可能来自其他模块)` |
| VB3003 | warning | 548 | `assignment: 无法将 Integer 隐式转换为 ?` |
| VB3044 | warning | 19 | `Implements: interface 'ISubclass' not found` |
| VB4001 | warning | 2 | `TypeLib reference path not found: ...\..\..\Windows\SysWOW64\shell32.dll` |
| 非诊断行 | info | 143 | `C3: CoClass '<name>' identity: CLSID=...`（身份提示，不是错误） |

- **error 0 条**：整工程 emit 已能跑完并产出 49 MB C 源码；缺口的形态是「warning + 类型退化为 `?`」，不是硬报错。

### 4.2 按根因聚类（含最小复现）

探针固定夹具：`.build\pd_probe\repro\`（`Probe.vbp` + `Entry.bas` + `IThing.cls` + `Impl.cls`，4 文件，单进程 <1 s），复现输出：

```
Impl.cls(16,1): warning VB3044: Implements: interface 'IThing' not found
Entry.bas(12,5): warning VB3003: assignment: 无法将 Byte 隐式转换为 ?
Entry.bas(17,5): warning VB3003: assignment: 无法将 Integer 隐式转换为 ?
Entry.bas(26,5): warning VB3003: assignment: 无法将 ? 隐式转换为 Byte
```

| 编号 | 根因 | 影响面 | 证据 | 候选修法难点 |
| --- | --- | --- | --- | --- |
| **R1** | 跨模块 `Public` 全局变量不可见（模块级公共变量进不了项目符号表/或解析顺序不成立） | **最大**：3693 条（占全部诊断 87%）；top 标识符 `g_Language` 1266、`workingDIB` 691、`curDIBValues` 508、`g_WindowManager` 319、`g_Themer` 226；涉及 90+ 文件（Toolbar_ToolSelect.frm 94、Interface.bas 90、Processor.bas 87） | vbp 已含 `Module=Public_Variables; Modules\PublicVars.bas`，其中 `Public g_Language As pdTranslate`；`workingDIB` 在 `EffectPrep.bas:29` | 需在项目级把各 `.bas` 的 Public 变量登记进全局作用域并让后续模块解析时可见；难点是解析顺序与「同名遮蔽」，还要避免把真未声明标识符一并静音 |
| **R2** | 函数返回数组类型 `As Byte()` 丢失，调用表达式类型退化为 `?` | 中：`cZipArchive.cls` 中 `pvToOemString(...) As Byte()` 的调用处成批告警（Byte→? / ?→Byte / Integer→?） | 探针 `MakeArr() As Byte()` 赋值给函数名报 Byte→?；`b = MakeArr()` 报 ?→Byte | 类型系统需为数组返回类型保留元素类型（VB6 数组赋值为整体拷贝）；难点是数组表达式在赋值/参数传递/ReDim 三条路径上都要保持元素类型，且不能影响现有数组 ABI |
| **R3** | `ByRef` 数组参数的元素类型丢失（`dst(i)` 类型为 `?`），`ReDim` 后亦然 | 中：`Palettes.bas` 154 条（`dstDitherTable() As Byte`、`srcPalette() As RGBQuad`）、`DibSupport.bas` 等 | 探针 `FillArr(ByRef dst() As Byte): ReDim dst(0 To 3): dst(0) = 7` 报 Integer→? | 需让数组形参符号携带「元素类型 + 维度」，下标表达式按元素类型定类型；难点是与 R2 共用数组类型表示，改动会波及数组下标/UBound/指针取址等既有站点 |
| **R4** | 空方法体的「接口类」无法被 `Implements` 解析 | 小：19 条（clsMenuImage、pdClipboard、pdEditBoxW、pdFlyout、pdFocusDetector、pdInputKeyboard 等） | `ISubclass.cls` 存在且 vbp 已登记 `Class=ISubclass; Classes\ISubclass.cls`；探针 `IThing.cls`（`Public Function ... End Function` 空体）+ `Impl.cls`（`Implements IThing`）稳定复现 | 需把「只有签名、无实现体」的类识别为接口类型并允许 `Implements` 绑定；难点是接口成员表与实现类方法名前缀（`IThing_DoWork`）的绑定规则，以及不影响普通空体函数（VB6 里空体函数在普通类里也是合法的） |
| **R5** | vbp 用相对路径引用 shell32/stdole2，CI 无法解析 | 小（夹具问题，非编译器）：2 条 | `TypeLib reference path not found: D:\a\c3test\c3test\tests\photodemon\..\..\Windows\SysWOW64\shell32.dll` | 已把夹具两条 `Reference=` 改为 `C:\Windows\SysWOW64\...` 绝对路径，待下一次 CI 复跑确认 VB4001 归零 |

### 4.3 优先级建议（供拍板）

1. **R3/R2 合并成一类做**（数组类型表示）：二者同源，修好数组元素/返回类型后能一次清掉 VB3003 的大头，且不影响 22 job 判据的概率较低。
2. **R4**（Implements 空体接口）：面小、边界清晰，适合先做一个独立闭环。
3. **R1**（跨模块 Public 变量）：条数最大，但改项目符号表构建顺序，回归面最广，建议放在数组类型之后、并有 22 job 全绿兜底再做。
4. R5 已修夹具，复跑确认即可。

## 5. R5 复跑（run 42584803572）

- 夹具 vbp 两条 `Reference=` 改绝对路径后自动触发：全部步骤绿，stderr 4405 → **4403 行**，`VB4001` **0 条**（R5 关闭）。
- 该 run 同时作为 R2/R3 修复前的基线：`VB3001` 3693、`VB3003` **548**、`VB3044` 19。

## 6. R2 + R3 收口：数组类型（`--arch x86`）

### 6.1 根因（改前钉死）

`Vb6Type::Array` 是**位标记**（`元素类型 | Array`），但"数组性"历史上记在两处，互相不认：

| 声明形态 | 符号类型 | `Symbol::isArray` |
| --- | --- | --- |
| `Dim a() As Byte`（局部/模块动态数组） | `Byte`（只有元素类型） | true |
| `Function F() As Byte()`（返回数组） | `Byte|Array` | 未设置 |
| `Sub S(ByRef dst() As Byte)`（数组形参） | `Byte|Array` | 未设置 |

两处消费点各只认一半，于是同一个数组在两个表达式里显示出不同的类型：

1. `semantic_analyzer_expr.cpp` 标识符取值直接给 `sym->type` ⇒ 局部数组读出来是**裸元素类型**（R2：`MakeArr = a` 报 `Byte → ?`）；
2. 同一文件数组下标分支只判 `sym->isArray` ⇒ 形参数组下标不脱 Array 位，返回**数组类型本身**（R3：`dst(0) = 7` 报 `Integer → ?`）。

`typeToString` 不认 Array 位，复合值一律打印成 `?` —— 这就是 548 条 VB3003 里 499 条带 `?` 的来源。

### 6.2 修法（一处归一，两行落地）

- `src/semantics/type_system.hpp`：新增三个纯位运算小工具 `isArrayType` / `arrayElemType` / `makeArrayType`，只服务语义层的表达式类型。
- `src/semantics/semantic_analyzer_expr.cpp`：
  - 标识符取值：`sym->isArray` 时把 Array 位补进**表达式类型**（符号类型不动，发码层照旧按符号读）；
  - 数组下标：判据从 `sym->isArray` 放宽为 `sym->isArray || isArrayType(sym->type)`，元素类型统一走 `arrayElemType`。
- **未改** `canImplicitConvert`：两侧都归一成 `元素|Array` 后，既有的"相同类型即相容"已足够；额外放宽会掩盖真错。

### 6.3 探针与判据

- 夹具 `tests/pd_arrtype/p01_array_elem.bas`（1279 B）：函数返回数组、整体数组赋值、ByRef 数组形参下标、UDT 数组元素四种形态。
- 已接进 `tests/run_tests.ps1` 的 `pd_arrtype_element_type`（`all`/`syntax` 段）：
  - 反向：stderr 不许出现 `VB3003`；
  - 正向：发码签名必须仍是 `vb6_SafeArray1D* vb6_MakeArr(void);` 与 `void vb6_FillArr(vb6_SafeArray1D** dst);`（防"靠退化发码消告警"）。
- 双向读数：修前 **EXIT=0、3 条 VB3003**（`Byte → ?`、`Integer → ?`、`? → Byte`）→ 修后 **EXIT=0、0 条**，签名不变。
- 回归：`tests/run_tests.ps1 -Category syntax` **PASS=183 FAIL=0 SKIP=0**（含本用例）。

### 6.4 构建安全参数（用户硬规则）

- 一律 `CMAKE_BUILD_PARALLEL_LEVEL=2`，不依赖 ninja 默认并行。
- 本次改动含 `type_system.hpp`（被广泛包含）⇒ 全量重编 77 目标，峰值并发按 2 封顶；随后复验为 `ninja: no work to do`（空闲内存 12.3 GB / 16.1 GB）。
- 以后改这类头文件要**攒批一次验证**，不反复重编。

## 6.5 R2/R3 的 CI 实测影响面（run 42590998643，提交 `4f10e519`）

| 诊断 | 修复前（run 42584803572） | 修复后 | 变化 |
| --- | --- | --- | --- |
| `VB3003` | 548 | **57** | **−491** |
| `VB3001` | 3693 | 3693 | 0（R1 未动） |
| `VB3044` | 19 | 19 | 0（R4 未动） |
| stderr 行数 | 4403 | 3912 | −491 |

归因：掉掉的 491 条全部是类型打印成 `?` 的那一族；`typeToString` 不认 Array 位，归一后不再出现。剩余 57 条属**另一类根因**（`Integer/Double/String → Object`、`String → Date` 的隐式转换判据），不在 R2/R3 范围内。

## 7. R4 收口：可创建的空体接口类（`VB_Creatable = True`）

### 7.1 根因（探针钉死）

`Implements` 的接口识别在 `src/driver/driver_interface.cpp` 的 **Pass A2**，判据四条，其中第三条是"显式写了 `Attribute VB_Creatable` 就必须为 False"。PhotoDemon 的 `Classes\ISubclass.cls` 是 `VB_Creatable = True` + 全空体签名 + 被 35 个类 `Implements`，于是被挡在接口登记表外：

- 语义层退回 legacy 路径（`semantic_analyzer.cpp` 的 `lookupModule` + `kind==Class` 那一支），只发 `VB3044` warning；
- 发码层 `ivLookupIface("ISubclass")` 恒 nullptr ⇒ 每个实现类各发一份同名 vtable（C2011 撞名风险）。

`VB_Creatable` 在 VB6 里只管 `New X` 能不能建，**与 `Implements X` 无关**：可创建的类照样能被别人当接口实现。

三组探针对照（`.build\pd_probe\repro`，vbp = IThing.cls + Impl.cls + Entry.bas）：

| 探针 | 接口宿主形态 | 旧产物 | 新产物 |
| --- | --- | --- | --- |
| repro | `VB_Creatable = True` + 全空体 | `VB3044` ×1 | **0**，且发出 canonical `vb6_ivtbl_IThing` + `vb6_ivtbl_IThing_for_Impl` |
| r4a | `VB_Creatable = False` + 全空体 | 0（对照组：旧判据只认这一种） | 0 |
| r4b | `VB_Creatable = True` + **带实现体** | `VB3044` ×1 | **1**（负控：`allMembersBodyless` 这条仍拦得住，不会把普通类当接口） |

### 7.2 修法

`src/driver/driver_interface.cpp` Pass A2：删掉 `VB_Creatable` 那条闸（连同只服务于它的 `attrLiteral` lambda），判据收敛为三条：被 `Implements` 指到 + 是 .cls 且无 tB Interface 块 + 全成员无体签名。安全性由第一条兜底——只有真被 `Implements` 指到的名字才进表。

### 7.3 探针用例

`tests/pd_iface/`（`IThing.cls` 可创建空体接口 + `Impl.cls` 实现 + `Entry.bas` 调用）已接进 `tests/run_tests.ps1` 的 `pd_iface_creatable_host`：

- 反向：不许出现 `VB3044`；
- 正向：必须发出 `typedef struct vb6_ivtbl_IThing {` 与 `static const vb6_ivtbl_IThing vb6_ivtbl_IThing_for_Impl = {`（证明走的是 canonical 面，不是 legacy 各发一份）。

回归：`tests/run_tests.ps1 -Category syntax` → **PASS=184 FAIL=0 SKIP=0**（184 = 上一轮 183 + 本用例）。

构建：只动 `driver_interface.cpp`（非头文件），`CMAKE_BUILD_PARALLEL_LEVEL=2` 下 91 目标，峰值 `cl.exe` ≤ 2，构建前空闲内存 10.3 GB。

## 8. R1 收口：标准模块 `Public` 变量的工程级可见性（VB3001 误报）

### 8.1 结论先行：这是**诊断误报**，不是发码缺口

PhotoDemon 全工程 3693 条 `VB3001` 对应的跨模块读写**一直在正确发码**：探针产物是 `extern int32_t gX;` + 一份定义 + `gX = 41;`，EXIT=0。缺的只是诊断层"这个裸名在工程里有"这一问。

结论对后续判断很关键：**别把 R1 当成"符号表 / 注入器没接上"**。本次改动**没有**动 `SymbolTable` 结构、**没有**动 stage 3.5 的注入时序、**没有**动一行发码。判断"还剩哪些真缺口"时，3693 这一族应从"发码缺口"里划掉，归入诊断噪声。

### 8.2 根因

`SemanticAnalyzer::namesProjectLevel`（`semantic_analyzer_util.cpp`）是"裸名位是否合法"的唯一判定点，它只认四档工程级名字：`projModuleNames_` / `projPublicProcs_` / `projPublicConsts_` / `projFormNames_`。**没有"工程级 Public 变量"这一档**，而跨模块符号注入发生在其后（stage 3.5），分析期 `symTab_.lookup` 必然查不到。

### 8.3 修法（三处，均在诊断层）

1. `semantic_analyzer.hpp`：新增第五档 `projPublicVars_` 与 `setProjectPublicVarNames`。
2. `semantic_analyzer_util.cpp`：`namesProjectLevel` 加一行精确匹配（仅大小写折叠，不做前缀/驼峰放宽）。
3. `driver_semantics.cpp`：在逐模块分析之前、与 `projPubProcs` / `projPubConsts` 同一装配点收集；只收标准模块的 `AccessLevel::Public`（`Private` 不跨模块，标准模块里的 `Friend` VB6 不认），并照走 `packageBlockedNames_` 包边界。

### 8.4 断言用例与负控读数（本地探针）

夹具 `tests/pd_globals/`（ModA/ModB/ModC），用例 `pd_globals_public_var_visibility`：

| 类别 | 检查点 | 修前 | 修后 |
| --- | --- | --- | --- |
| 正面 | 跨模块读 `gX` / 写 `gX = 41` | 假 VB3001 | 0 |
| 正面 | `Public` 动态数组 `gArr(0)` | 假 VB3001 | 0 |
| 正面 | `Public` UDT 成员 `gColor.r` | 假 VB3001 ×2 | 0 |
| 正面 | `ModC` 自己的 `Private gX` 遮蔽 | — | 0 |
| 负控 1 | 拼错的全局名 `gXX` | — | **仍报 VB3001** |
| 负控 2 | `ModA` 的 `Private pvLocal` 在 `ModB` 引用 | — | **仍报 VB3001** |
| 负控 3 | 标准模块里的 `Friend fFr` | — | **仍报 VB3001** |

- 修前红：旧二进制同形状探针 `.build\pd_probe\r1\o.err` 给出 4 条假 VB3001（`gX`×2、`gColor`×2）。
- 负控 2 最初取名 `pv`，撞上 VB6 内建财务函数（单模块 `pv = 1` 也不报），改名 `pvLocal` 后成立——以后负控别用 `pv` / `fv` / `rate` 这类内建名。
- 回归：`-Category syntax` → **PASS=185 FAIL=0**；`check_manifest_coverage` → `inputs=610 [.vbp=224 .bas=386], registered=610, sets equal`。

### 8.5 CI 实测影响面

待本轮 `probe/photodemon-emit` run 回填（VB3001 应从 3693 实际下降；VB3044 应随 R4 从 19 降到 0）。

## 9. 暂缓项清单（记账型红灯，用户已授权可挂）

口径（用户 2026-11-09 定）：**emit manifest / 覆盖率 / census 计数这类非源码性质的记账型红可以暂缓，产品正确性类红必须修**。这里逐条留账，不静默忽略。

| # | run / job | 现象 | 为什么是记账型 | 处置 |
| --- | --- | --- | --- | --- |
| A1 | run 42590924685 · `Emit manifest (shape oracle)` | 报红：语料新增输入未登记哈希 | 比较器只比"哈希表行集合"，与程序行为无关；差异全部是"新增行" | **已消除**：`tests/photodemon/**` 从语料枚举里排除（见 9.1），新夹具行按本机 `emit_manifest.ps1` 生成（方法已用既有行逐字对账，见 9.2） |
| A2 | run 42590924685 · `Tests (compile #1)`（`[STATIC] manifest_coverage`） | K1 报 **115** 份语料输入未登记 | 同上，是 A1 的本地哨兵版 | **已消除**：排除后 K1 只剩新夹具，已登记；哨兵现为 `inputs=610 [.vbp=224 .bas=386], registered=610, sets equal` |

### 9.1 排除口径（用户 2026-11-09 拍板）

`tests/photodemon/` 是**第三方工程的 vendored 副本**（PhotoDemon 源集），只用于 `--arch x86` emit 探针取证，不是编译器夹具。它的 115 份 `.bas/.vbp` 每轮让形状门多跑 115 次 emit（约 +8 分钟）而无产品价值，因此从形状门语料里排除。

- 位置：`scripts/emit_manifest.ps1` 的语料枚举（**唯一权威**），一行 `-notmatch '\\tests\\photodemon\\'` 加注释。
- 可逆：删掉那一行即恢复原口径，别处不用改；恢复后需要按 CI 的 `ci/emit-manifest` 一次性补登 115 行。
- 不受影响：探针 workflow（`.github/workflows/pd_emit_probe.yml`）自己遍历 `tests/photodemon/**`，与形状门语料无关。

### 9.2 新夹具行的生成方法（避免"造哈希"）

用 `scripts/emit_manifest.ps1` 本身生成（不另写哈希逻辑）：把新夹具按原相对路径摊进一个临时根 `.build\msub\`（连同当前构建的 `.build\C3.exe`），`-Root` 指向它，产出即得行。**方法先对账**：用同一手法对既有语料 `tests/acc/acc_main.bas` 生成一行，与 `emit-manifest.expected.txt` 里已登记的那行**逐字节相同**（R1 这轮同时复核了 `pd_arrtype` / `pd_iface` 两行），证明临时根不影响哈希；再按排序位插入。
若 CI 侧出现 `sha256` 假差（本机 MSVC 14.29 与 CI 工具集注释编码不同），以 `ci/emit-manifest` 的读数为准覆盖，不猜。

## 10. 下一步

1. VB3003 剩余 57 条的隐式转换判据。
2. `--arch x86` 真编译（emit 后跑 cl/link），按编译错误反推下一批缺口。
3. 控件属性 / 方法缺失类缺口（需先有一份可编译的产物才能暴露）。
