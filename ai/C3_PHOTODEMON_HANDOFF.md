# PhotoDemon 编译缺口交接（C3_PHOTODEMON_HANDOFF）

独立于其它工程的交接文档（不要并入 ucTreeView / ucSlider / 分段控件那几份）。

- 目标：完善 C3 编译器，最终能用 `--arch x86` 成功编译 `D:\vb_yqt4qPac\PhotoDemon-main\PhotoDemon.vbp`。
- VB6 参考 exe：`D:\vb_yqt4qPac\PhotoDemon-main\PhotoDemon.exe`（图标/界面基准，仅作对照）。
- 节奏约束：单进程串行构建、低并行度；不急于编过，重点是补齐语法 / 控件方法属性 / 类型与语义缺口；每类缺口单独提交。

## 0. 工作区纪律（两条线共用同一工作区，2026-11-10 起）

- **禁用 `git add -A` / `git add .` / `git commit -a`**，一律按显式路径 `git add <file>...` 提交。
- 提交前先看 `git status`：出现不属于本刀的文件就**停下报告**，不要顺手带进去。
- 非 `src/**` 源码变动（文档、注释、handoff、`.temp` 探针）不跑 CI —— 空跑 GA 无意义，直接 skip。
- 起因：`tests/photodemon/*`（643 份）曾被另一条线的 `git add` 连带扫入并推送（`7230557e`），内容未改、不改写历史。连带后果是 `vbp_fixture_census` 报 `photodemon\PhotoDemon.vbp` 未登记 —— 本线记账，等内存主线清完再处理。

## 0.1 本机资源硬规则（2026-11-10 事故后，**任何一轮都适用**）

事故：本机跑 `PhotoDemon.vbp --emit-c --arch x86`（`.temp\pd_run.ps1`），峰值工作集 **12,518 MB**、
墙钟约 52 s 后 `bad allocation`，16 GB 机器进入交换、整机被拖死。stdout 0 B，未产出任何 C。

因此：

- **禁止**在本机整工程跑 `PhotoDemon.vbp`（`--emit-c` / `--syntax-only` / 任何模式、任何内存上限）。
  16 GB 撑不住；一进交换就是把整机拖死，不是"慢一点"。
- 整工程验证**只走 CI**（16 GB runner），本地只取 artifact / 日志。
- 本机只许：单文件探针、合成小语料（**≤ 60 模块**）、`-Category syntax` 的定向用例。
- 任何构建必须 `CMAKE_BUILD_PARALLEL_LEVEL=2`；起构建前先报空闲物理内存，**低于 8 GB 不许起**。
- 启动任何长命令前必须报「预计峰值内存 + 依据（历史读数/语料规模）」，不得凭感觉启动。
- 本轮实测的 per-module 斜率（Release，合成语料）：共享层后 **4.10 MB/模块**，改造前 7.64 MB/模块；
  PhotoDemon 539 模块 ⇒ 按此推算仍在 2 GB 以上，**本机永远不要试整工程**。


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

## 4. R2 + R3 收口：数组类型（2026-11-09）

口径已拍板：整工程 emit 不在本机跑，由 GitHub Actions 探针分支（`probe/photodemon-emit`）执行，本地只取 artifact。探针结论与本节的完整版记录在该分支的同名文件里；`dev` 只保留代码修复与本摘要。

### 4.1 根因

`Vb6Type::Array` 是**位标记**（`元素类型 | Array`），但"数组性"历史上记在两处，互相不认：

| 声明形态 | 符号类型 | `Symbol::isArray` |
| --- | --- | --- |
| `Dim a() As Byte`（局部/模块动态数组） | `Byte`（只有元素类型） | true |
| `Function F() As Byte()`（返回数组） | `Byte|Array` | 未设置 |
| `Sub S(ByRef dst() As Byte)`（数组形参） | `Byte|Array` | 未设置 |

两处消费点各只认一半：

1. `semantic_analyzer_expr.cpp` 标识符取值直接给 `sym->type` ⇒ 局部数组读成**裸元素类型**（R2：`MakeArr = a` 报 `Byte → ?`）；
2. 同文件数组下标分支只判 `sym->isArray` ⇒ 形参数组下标不脱 Array 位，返回**数组类型本身**（R3：`dst(0) = 7` 报 `Integer → ?`）。

`typeToString` 不认 Array 位，复合值一律打印成 `?` —— PhotoDemon 全工程 548 条 VB3003 里 499 条带 `?` 即由此而来。

### 4.2 修法

- `src/semantics/type_system.hpp`：新增 `isArrayType` / `arrayElemType` / `makeArrayType` 三个纯位运算工具（只服务语义层表达式类型）。
- `src/semantics/semantic_analyzer_expr.cpp`：标识符取值在 `sym->isArray` 时把 Array 位补进**表达式类型**（符号类型不动，发码层照旧按符号读）；数组下标判据放宽为 `sym->isArray || isArrayType(sym->type)`，元素类型统一走 `arrayElemType`。
- **未改** `canImplicitConvert`：两侧归一后既有的"相同类型即相容"已足够，放宽会掩盖真错。

### 4.3 探针与回归

- 夹具 `tests/pd_arrtype/p01_array_elem.bas`（1279 B，函数返回数组 / 整体赋值 / ByRef 数组形参下标 / UDT 数组元素）。
- 已接进 `tests/run_tests.ps1` 的 `pd_arrtype_element_type`（`all`/`syntax` 段）：反向钉 `VB3003` 不出现；正向钉发码签名仍为 `vb6_SafeArray1D* vb6_MakeArr(void);` 与 `void vb6_FillArr(vb6_SafeArray1D** dst);`（防"靠退化发码消告警"）。
- 双向读数：修前 EXIT=0 / **3 条 VB3003**（`Byte → ?`、`Integer → ?`、`? → Byte`）→ 修后 EXIT=0 / **0 条**，签名不变。
- 回归：`tests/run_tests.ps1 -Category syntax` → **PASS=183 FAIL=0 SKIP=0**（含本用例）。
- 构建：`CMAKE_BUILD_PARALLEL_LEVEL=2`；本次含 `type_system.hpp` 导致 77 目标重编，并发峰值按 2 封顶，随后复验 `ninja: no work to do`（空闲内存 12.3 GB / 16.1 GB）。

## 5. R4 收口：可创建的空体接口类（`VB_Creatable = True`）

### 5.1 根因

`Implements` 的接口识别在 [driver_interface.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_interface.cpp) 的 Pass A2，旧判据四条，其中一条是"显式写了 `Attribute VB_Creatable` 就必须为 False"。PhotoDemon 的 `Classes\ISubclass.cls` 是 `VB_Creatable = True` + 全空体签名 + 被 35 个类 `Implements`，被这道闸挡在接口登记表外 ⇒ 语义层退回 legacy 只发 `VB3044` warning，发码层 `ivLookupIface` 恒 nullptr，每个实现类各发一份同名 vtable。

`VB_Creatable` 只管 `New X` 能不能建，与 `Implements X` 无关——可创建的类照样能当接口被实现。

### 5.2 修法

删掉 `VB_Creatable` 那条闸（连同只服务于它的 `attrLiteral` lambda），判据收敛为三条：被 `Implements` 指到 + 是 .cls 且无 tB Interface 块 + 全成员无体签名。安全性由第一条兜底：只有真被 `Implements` 指到的名字才进表。

### 5.3 探针（三组，同一 vbp 形状 IThing/Impl/Entry）

| 探针 | 接口宿主形态 | 旧产物 | 新产物 |
| --- | --- | --- | --- |
| repro | `VB_Creatable = True` + 全空体 | `VB3044` ×1 | **0**，发出 canonical `vb6_ivtbl_IThing` + `vb6_ivtbl_IThing_for_Impl` |
| r4a | `VB_Creatable = False` + 全空体 | 0（旧判据只认这一种） | 0 |
| r4b | `VB_Creatable = True` + **带实现体** | `VB3044` ×1 | **1**（负控：空体判据没被放宽） |

用例 `pd_iface_creatable_host`（`tests/pd_iface/`，接进 `run_tests.ps1`）：反向钉 `VB3044` 不出现，正向钉 `typedef struct vb6_ivtbl_IThing {` 与 `static const vb6_ivtbl_IThing vb6_ivtbl_IThing_for_Impl = {`。
回归：`-Category syntax` → **PASS=184 FAIL=0 SKIP=0**。构建：`CMAKE_BUILD_PARALLEL_LEVEL=2`，91 目标，峰值 `cl.exe` ≤ 2。

### 5.4 R2/R3 的 CI 实测影响面（run 42590998643）

| 诊断 | 修前 | 修后 | 变化 |
| --- | --- | --- | --- |
| `VB3003` | 548 | **57** | **−491**（全部是类型打印成 `?` 的那一族） |
| `VB3001` | 3693 | 3693 | 0（R1 未动） |
| `VB3044` | 19 | 19 | 0（该 run 早于 R4） |

## 6. R1 收口：标准模块 `Public` 变量的工程级可见性（VB3001 误报）

### 6.1 结论先行：这是**诊断误报**，不是发码缺口

PhotoDemon 全工程 3693 条 `VB3001` 对应的跨模块读写**一直在正确发码**：探针产物是 `extern int32_t gX;` + 一份定义 + `gX = 41;`，EXIT=0。缺的只是诊断层"这个裸名在工程里有"这一问。

结论对后续判断很关键：**别把 R1 当成"符号表 / 注入器没接上"**。本次改动**没有**动 `SymbolTable` 结构、**没有**动 stage 3.5 的注入时序、**没有**动一行发码。判断"还剩哪些真缺口"时，3693 这一族应从"发码缺口"里划掉，归入诊断噪声。

### 6.2 根因

`SemanticAnalyzer::namesProjectLevel` 是"裸名位是否合法"的唯一判定点，它只认四档工程级名字：`projModuleNames_`（模块限定符）/ `projPublicProcs_`（`Public` 过程）/ `projPublicConsts_`（`Public` 常量与 Enum 成员）/ `projFormNames_`（窗体默认实例）。**没有"工程级 Public 变量"这一档**。

而跨模块符号注入（`driver_crossmod.cpp` 的 `globalPublicSyms` → `injectExternSymbol`）发生在 stage 3.5，即逐模块分析**之后**；分析当刻 `symTab_.lookup` 必然查不到，于是每个跨模块引用各配一条假 `VB3001`。`getPublicSymbols()` 本身是导出了 `SymbolKind::Variable` 的——所以问题不在登记，也不在注入。

### 6.3 修法（三处，均在诊断层）

1. `src/semantics/semantic_analyzer.hpp`：新增第五档 `projPublicVars_` 与 `setProjectPublicVarNames`，与 `projPublicProcs_` 同款。
2. `src/semantics/semantic_analyzer_util.cpp`：`namesProjectLevel` 末尾加一行 `projPublicVars_.count(lk)` 判定。**只做精确名字匹配（仅大小写折叠）**，不做前缀 / 驼峰放宽——那等于把误报换成漏报。
3. `src/driver/driver_semantics.cpp`：在与 `projPubProcs` / `projPubConsts` **同一装配点、同一时刻**（逐模块分析之前）收集：
   - 只收标准模块（类/窗体在 `isClassModule || isFormModule` 处已 `continue`，它们的 `Public` 变量是实例成员，得经对象名访问）；
   - 只收 `AccessLevel::Public`：`Private` 不跨模块；标准模块里的 `Friend` VB6 本来就不认（`Friend` 只对类模块有意义），不收；
   - 照样走 `packageBlockedNames_` 的包边界（被包 manifest 屏蔽的名字不进工程级名单），规则只有一份。

### 6.4 断言用例与负控读数

夹具 `tests/pd_globals/`（`ModA.bas` / `ModB.bas` / `ModC.bas`），用例 `pd_globals_public_var_visibility` 接进 `tests/run_tests.ps1` 的 syntax 段。

| 类别 | 检查点 | 修前 | 修后 |
| --- | --- | --- | --- |
| 正面 | 跨模块**读** `gX`（`gArr(0) = gX + 1`） | 假 `VB3001` | 0 |
| 正面 | 跨模块**写** `gX = 41` | 假 `VB3001` | 0 |
| 正面 | `Public` 动态数组 `gArr(0)` 读写 | 假 `VB3001` | 0 |
| 正面 | `Public` UDT 成员 `gColor.r` | 假 `VB3001` ×2 | 0 |
| 正面 | 同名遮蔽：`ModC` 自己的 `Private gX` | — | 0 |
| **负控 1** | 拼错全局名 `gXX` | — | **仍报 `VB3001`** |
| **负控 2** | `ModA` 的 `Private pvLocal` 在 `ModB` 引用 | — | **仍报 `VB3001`** |
| **负控 3** | 标准模块里的 `Friend fFr` | — | **仍报 `VB3001`** |

- 修前红：同一形状探针在旧二进制上 `.build\pd_probe\r1\o.err` 给出 4 条假 `VB3001`（`gX`×2、`gColor`×2）。
- 修后：上述 3 条负控全部照旧报出，正面的 5 类全部归零 ⇒ 不是"把告警关掉"。
- 踩到的一个坑：负控 2 最初用变量名 `pv`，而 `pv` 是 VB6 内建财务函数名（单模块里 `pv = 1` 也不报），属于**测试取名撞内建**，与实现无关；改用 `pvLocal` 后负控成立。以后写负控别用 `pv` / `fv` / `rate` 这类内建名。
- 回归：`tests/run_tests.ps1 -Category syntax` → **PASS=185 FAIL=0 SKIP=0**。
- 形状门：`scripts/check_manifest_coverage.ps1` → `inputs=610 [.vbp=224 .bas=386], registered=610, sets equal`（3 份新夹具行按 7.2 的方法生成，先与既有行逐字节对账）。

### 6.5 CI 实测影响面（run 42599592249，提交 `ba455c7f`）

`EMIT_EXIT=0`、`EMIT_SECONDS=792.7`、stdout 48,341,497 B、stderr 238 行。

| 诊断 | 修前 | 修后 | 变化 |
| --- | --- | --- | --- |
| `VB3001` | 3693 | **38** | **−3655** |
| `VB3003` | 57 | 57 | 0（R1 未动） |
| `VB3044` | 19 | **0** | **−19**（R4 生效） |

剩余 38 条按名字归总（**与 R1 无关，是下一轮的新根因**）：

| 名字 | 条数 | 判定 |
| --- | --- | --- |
| `Math` | 18 | `Modules\ComplexNumbers.bas` 与 `Forms\Effects_Distort_Droste.frm` 的 `Math.Log/Abs/Cos/Sin`。**旧判定「疑似 tB-only」已作废**：真 VB6 探针确认内建 `VBA.Math` 模块，18 处本来就是合法 VB6；已在 `84a74b69` 收口，见 §10.4 |
| `Parent` | 7 | `Controls\pdCommandBar.ctl` 里裸写 `Parent.Controls`（等价 `UserControl.Parent`）⇒ **UserControl 内建成员在裸名位未登记** |
| `vbCFFiles` | 7 | `pdClipboardMain.cls` 的 `Data.GetFormat(vbCFFiles)` ⇒ 内建常量表缺项（`vbCFText` / `vbCFBitmap` 都在，只漏这一个） |
| `ldResult` / `ERR_BUFF_TOO_SMALL` | 2 | `pdXCF.cls` 在 `Do ... Loop` 体内 `Dim` / `Const`，`Loop While` 条件里再用 ⇒ **块内声明未按 VB6 的过程级作用域处理** |
| `positionFailed` / `failsafePosition` | 2 | `UserControl_Support.bas` 同类：循环 / `If` 块内声明的量在块外引用 |
| `Randomize` | 1 | `pdPaletteChild.cls` 的 `Randomize -1`（带参数的语句形态）⇒ 内建语句未登记 |
| `vbAsyncStatusCodeEndDownloadData` | 1 | `pdDownload.ctl` ⇒ 内建常量表缺项 |

即：R1 之后 `VB3001` 已从"刷屏"降到 38 条，且**每一条都能指到具体的一类缺口**，不再有"跨模块全局变量"这一族。

### 6.6 回归与既有红的分界

- 本地：`-Category syntax` PASS=185；`pd_globals` 探针正面 0 条、三条负控各 1 条；`check_project_const_visibility` PASS；`check_manifest_coverage` → `inputs=610 … sets equal`。
- CI `42599592249`（R1 本体）：探针绿，计数见 6.5。
- CI `42602700528`（`de54cd09`）：`Tests (compile #1)` **success**。这是 R1 **唯一**引入的回归，已修：新分支在 driver 里又抄了一遍 `packageBlockedNames_` 查询，把 S3「全仓只许一处问」钉红；改为每个标准模块进声明循环前查一次、变量/常量/过程三支共用同一个只读指针，语义不变。
- 其余红与 R1 无关（可证据化）：`Tests (bas #2)` 的失败用例集合在 `91de0a68`（R1 前）与 `de54cd09` 逐条相同，**54 vs 54、零差**；`Tests (vbp #1..#12)`、`bas #1`、`Emit manifest` 在 R1 前的 run `42598038994` 里已红。
  - 记账型：`Emit manifest` 报 6 条 sha256 差（`k165_iface_slot` / `k172_viewarg` / `k174_vieweach` / `k175_viewstmt` / `k177_viewwith` / `k178_viewprop`），按用户口径暂挂。
  - 产品型：上述 6 个 k 用例与大量 `EMITC-SHAPE` / `EMITC-ABSENT` 是同一批既有债务（数十条），**是否由 R2/R3/R4 引入尚未归因**——下一轮第一件事是拿 R2 之前的提交跑一次同名 job 做二分，别把它记到 R1 账上。

## 7. 暂缓项清单（记账型红灯，用户已授权可挂）

口径（2026-11-09）：emit manifest / 覆盖率 / census 计数这类非源码性质的记账型红可以暂缓；产品正确性类红必须修。逐条留账，不静默忽略。

| # | run / job | 现象 | 为什么是记账型 | 处置 |
| --- | --- | --- | --- | --- |
| A1 | run 42590924685 · `Emit manifest (shape oracle)` | 语料新增输入未登记哈希 | 只比"哈希表行集合"，与程序行为无关，差异全是"新增行" | **已消除**：`tests/photodemon/**` 从语料枚举排除（见 7.1），新夹具行按 `emit_manifest.ps1` 自身生成（方法先与既有行逐字节对账，见 7.2） |
| A2 | run 42590924685 · `Tests (compile #1)`（`[STATIC] manifest_coverage`） | K1 报 **115** 份语料输入未登记 | 同上，是 A1 的本地哨兵版 | **已消除**：排除后 K1 只剩新夹具，已登记；现为 `inputs=610 [.vbp=224 .bas=386], registered=610, sets equal` |

### 7.1 排除口径（用户 2026-11-09 拍板）

`tests/photodemon/` 是**第三方工程的 vendored 副本**，只用于 `--arch x86` emit 探针取证，不是编译器夹具。它的 115 份 `.bas/.vbp` 每轮让形状门多跑 115 次 emit（约 +8 分钟）而无产品价值，因此从形状门语料里排除。

- 位置：`scripts/emit_manifest.ps1` 的语料枚举（**唯一权威**），一行 `-notmatch '\\tests\\photodemon\\'` 加注释。
- 可逆：删掉那一行即恢复原口径，别处不用改；恢复后按 CI 的 `ci/emit-manifest` 一次性补登 115 行。
- 不受影响：探针 workflow 自己遍历 `tests/photodemon/**`，与形状门语料无关。

### 7.2 新夹具行的生成方法

用 `scripts/emit_manifest.ps1` 本身生成：把新夹具按原相对路径摊进临时根 `.build\msub\`（连同 `.build\C3.exe`，必须是当前构建），`-Root` 指向它。**方法先对账**——同一手法对既有 `tests/acc/acc_main.bas` 生成一行，与已登记行逐字节相同（R1 这轮同时复核了 `pd_arrtype` / `pd_iface` 两行），再按排序位插入新行。若 CI 出现 `sha256` 假差（工具集注释编码不同），以 `ci/emit-manifest` 读数为准覆盖。

## 8. 下一步

1. `VB3001` 剩余 38 条按 6.5 的分类逐类收口（内建常量表缺项 / UserControl 内建成员 / 块内声明的过程级作用域 / `Randomize` 语句形态）；`Math` 18 条已定案为合法 VB6 并在 `84a74b69` 收口（§10.4）。
2. `VB3003` 剩余 57 条的隐式转换判据。
3. `--arch x86` 真编译（emit 后跑 cl/link），按编译错误反推下一批缺口。
4. 控件属性 / 方法缺失类缺口（需先有一份可编译的产物才能暴露）。

## 9. R4 回归修复与 sha256 二分归因（2026-11-09）

### 9.1 二分表：唯一引入点是 R4

对 6 条 `EMITC-SHAPE` 差异（`k165_iface_slot` / `k172_viewarg` / `k174_vieweach` / `k175_viewstmt` / `k177_viewwith` / `k178_viewprop`），在 R1 前的 `91de0a68` 与当前 head 之间逐个提交重跑形状 oracle：

| 提交 | 内容 | 6 条形状 |
| --- | --- | --- |
| `fa56b9e2` | R2/R3 之前 | 全部 **SAME** |
| `a7fc9137` | R2+R3（数组类型） | 全部 **SAME** |
| `fdbd29e5` | R4（可创建空体接口） | 全部 **DIFF** |
| `91de0a68` | R1 之前（语料排除） | 全部 **DIFF** |
| `a3541e9a` | R1 之后 HEAD | 全部 **DIFF** |

结论：**R4 是唯一引入点，R2/R3 与 R1 对产物零影响**。这也把 §6.6 里「尚未归因」的那笔账结清：不是 R1，也不是数组归一。

### 9.2 逐条判定：没有一条是「预期内形状变化」

| 夹具 | 判定 | 现象 | 根因 | 修法 |
| --- | --- | --- | --- | --- |
| `k172_viewarg` | **真回归** | cl C2039 + `boxtype=46832` | 形参归一后是薄指针，但没登新式名册 ⇒ 按 legacy 形状发 `.vtbl`/`.obj`；ByRef Variant 实参把 `vb6_ivref_IK172**` 当 `vb6_VARIANT*` | `cgen_decl_func.cpp` / `cgen_decl_prop.cpp` 归一后同时登 `knownIvrefVars_`；新增 `ivrefViewBoxExpr`，ByRef Variant 站发临时 `vb6_VARIANT`（`vb6_VariantFromProjectInstance`）并在 `boxTempsToClear_` 里清 |
| `k174_vieweach` | **真回归** | C2440（Variant 直接赋薄指针） | For Each 落最后一臂 `vb6_VariantFromStackVARIANT` | `cgen_control.cpp` 加 canonical 分支：解包对象 → 取类名 → `vb6_ivref_<I>_from_class` 派发 |
| `k175_viewstmt` | **真回归（静默）** | 语句位成员调用落 COM 兜底，读数错 | 语句位零实参调用只认 legacy 视图 | `cgen_call.cpp` 站D 补 `slotIvrefViewName` → `ivrefVtableCall(...)` |
| `k177_viewwith` | **真回归（静默）** | With 内调用落 COM 兜底，读数错 | P6.4 只认 legacy；With 接收者未登新式名册 | `cgen_expr_call_com_bind.inc` P6.4 二选一；`cgen_with.cpp` 接收者判定 + 名册登记/注销 |
| `k178_viewprop` | **真回归** | `rc=2/rn=3`（期望 4/2） | `Debug.Print x.Label` 走 COM getter；With 接收者同上 | Debug.Print 实参站补 `ivrefViewGetterReadExpr`；With 同 k177 |
| `k165_iface_slot` | **预期内形状变化** | 产物语义等价（`mi=16`） | R4 让该接口走 canonical | 只更新形状断言（`__iv_K165Iface` 系列） |

### 9.3 新增公共站（For Each / 视图赋值共用的唯一入口）

- `ivrefIfaceViewFromFn(iface)` → `vb6_ivref_<I>_from_class`；`emitIvrefIfaceViewFrom(module)` 在接口模块里发：
  `vb6_ivref_<I>* vb6_ivref_<I>_from_class(void* instance, const char* className)`，按类名 `_stricmp` 派发到 `vb6_iv_thin_of_<C>(instance)`，认不出返回 `NULL`（= Nothing）。
- 由 `cgen_base_generate_body_pass.inc` 在 `emitLegacyIfaceViewFrom` 之后调用，仅限接口类模块。

### 9.4 修后读数（本机，`CMAKE_BUILD_PARALLEL_LEVEL=2`）

- k172 `pushA=101/pushB=202/view2view=101/is-diff=0/is-same=1/nothing=1/after=0/boxtype=0/DONE`
- k174 `sum=303/skipped=2/DONE`；k175 `alpha=2/beta=10/t=15/f=1100/fb=1200/fc=1100/DONE`
- k177 `alpha=2/beta=10/t1=6/t2=15/DONE`；k178 `s=B1/n=501/rc=4/rn=2/beta=1/betan=1/w=8/let=0/DONE`；k165 `mi=16/DONE=1`
- 6 夹具 need/absent 形状自检：`need_missing=0 bad_present=0`
- `tests/run_tests.ps1 -Category syntax` → **PASS=185 FAIL=0 SKIP=0**
- `scripts/check_project_const_visibility.ps1`（S3 单点哨兵）→ **PASS**

### 9.5 形状门记账（按用户口径挂账，不逐行改）

本地跑 `scripts/emit_manifest.ps1` → `.build\r4_manifest.txt`（610 行，`rc!=0` 88 行属存量）；与 `emit-manifest.expected.txt` 比对 **changed=24 / only-in-exp=0**：

- 6 行 = 上述 R4 夹具；
- 18 行 = 含 canonical 接口的工程，因新增 `from_class` 派发器而多一段函数。

处置：延后统一 re-bless（`scripts/rebless_emit_manifest.ps1 -Bless` 或取 CI `ci/emit-manifest` 的 artifact），不在本轮逐行手改哈希。

## 10. VB3001 剩余 38 条：逐类收口

### 10.1 内建常量表补项（8 条）

两枚缺表常量，值取自 VB6 参考（Clipboard 格式 / `StatusCode Property` 页）：

| 常量 | 值 | 用途 | PhotoDemon 条数 |
| --- | --- | --- | --- |
| `vbCFFiles` | 15（`CF_HDROP`） | `Data.GetFormat(vbCFFiles)` | 7 |
| `vbAsyncStatusCodeEndDownloadData` | 6 | `AsyncProp.StatusCode <> ...` | 1 |

- 改动：[builtin_consts_ext.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/builtin/builtin_consts_ext.inc#L194-L204) 两行 `addConst`（同一处表，未新增第二答案源）。
- 修前红：`tests\test_builtin_const_ext.bas`（带 `Option Explicit`）在旧二进制上：
  `warning VB3001: 未声明的标识符: 'vbCFFiles'` 与 `... 'vbAsyncStatusCodeEndDownloadData'` 各 1 条，符号落成隐式 Variant。
- 修后：编译 stderr **0 行**，运行读数 `BCE-CF=15` / `BCE-ASYNC=6` / `BCE PASS`。
- 断言：`run_tests.ps1` 新增 `test_builtin_const_ext`（Compile+Run）与 `bce_emitc_reads_are_literals`（发码必须是字面量 `f = 15;` / `s = 6;`）；`test_intrinsic_consts` 追加 `IC-CF-FILES=15` / `IC-ASYNC-END=6`（无 `Option Explicit` 那一形同时被钉住）。
- 哨兵：`scripts/check_builtin_const_authority.ps1` → `OK ... table=555 numeric-folds=56 expr-folds=11 rtl-vb-defines=0`（表 +2，折叠计数未变 ⇒ 没有长出第二答案源）。
- 记账：新夹具 `tests/test_builtin_const_ext.bas` 尚未登记 emit-manifest（形状门/coverage 按用户口径暂缓，后期统一 re-bless/补登）。

### 10.2 `Randomize` 语句形态（1 条）

- 现象：`pdPaletteChild.cls` 的 `Randomize -1` 报假 `VB3001`，产物是 `(vb6_Randomize - 1);`（把 `-` 绑成中缀减法，函数没被调用）。
- 根因：`Randomize` 是软关键字，语句入口没有专属分派，落进通用「表达式 / 无括号调用」路径；**裸调用与正种**碰巧能成（`Randomize`→`vb6_Randomize(0.0)`、`Randomize 5`→`vb6_Randomize(5)`），**负种**被表达式层先吃成减法。
- 修法：[parser_stmt.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/parser/stmt/parser_stmt.cpp#L284-L309) 增加 `case TokenKind::Randomize`，与 `MsgBox` 同款语句形态：紧跟 `=`/`(`/`.`/`:` 时仍判为标识符走通用路径（软关键字可作变量名），否则按 `Randomize [seed]` 构造 `CallStmt`。
- 读数：修前探针 `Randomize`/`Randomize 5`/`Randomize -1` → `(vb6_Randomize - 1)` + 2 条 VB3001；修后 `vb6_Randomize(0.0);` / `vb6_Randomize(5);` / `vb6_Randomize((-1));`，stderr 0 行。`Randomize Timer` 仍发 `vb6_Randomize(vb6_Timer());`（Charts 2020 语料在用的形态）。
- 断言：`tests/test_randomize_stmt.bas`（Compile+Run，`RZ-OK=True`：同种子两次投放首值一致，负种同样生效）+ `rz_emitc_forms` 形状钉 + `rz_emitc_no_binary_sub` 反向钉。
- 回归：`-Category syntax` **PASS=185 FAIL=0**。

### 10.3 `Parent` 内建成员（7 条）

- **语义定案（后被 10.3.1 的真 VB6 探针推翻，保留原文以便对照）**：当时按文档推断 `VB6 UserControl.Parent` 返回**宿主容器对象**（窗体 / 容器控件），**不是容器窗口句柄**；`Parent.Controls` 就是父容器的控件集合，与 `UserControl.ParentControls` 是同一件事。RTL 一侧的答案因此是 `vb6_UC_ParentObject()`（值位）与 `vb6_UC_ParentControls`（集合），不是 `hWnd`。
- 语料形态：PhotoDemon `pdCommandBar.ctl` 7 处全是 `For Each eControl In Parent.Controls`，循环变量 `As Object`。
- 根因：`Parent` 是 UserControl 的内建名，但不在 host-pseudo 表里（RTL 没有 `vb6_UserControl_Parent` 全局，表里刻意不收）；语义层 `isDocumentHostObject` 不认它 ⇒ 假 VB3001；发码层裸 `Parent` 被拼成 `vb6_Parent_Controls`（未声明符号）。
- 修法（三处，均在既有单一答案源上扩展，未新增第二权威）：
  1. [semantic_analyzer_util.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/semantic_analyzer_util.cpp)：`isDocumentHostObject` 认 `parent`，仅 `DocumentKind::UserControl`，值位/限定位都合法。
  2. [cgen_expr_ident_builtin.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_ident_builtin.inc)：裸 `Parent` 落 `vb6_UC_ParentObject()`（与 Fix 133u 对 `UserControl.Parent` 的答复同一枚）。
  3. [cgen_expr_member_precheck.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_member_precheck.inc)：`Parent.Controls` 交出已同步的 `vb6_UserControl_ParentControls`（与限定写法同形；**不**立 COM 标记，否则会多解一层成 `ComGetObjectProp(..., L"Controls")`）。
- 修前 / 修后读数（`tests/pd_parentctl/PdParentCtl.vbp`）：
  - 修前：`warning VB3001: 未声明的标识符: 'Parent'`，产物 `vb6_ForEach_Init(vb6_Parent_Controls)`。
  - 修后：stderr 0 行；`hostObj = vb6_UC_ParentObject();` + `void* _fe_enum_0 = vb6_ForEach_Init(vb6_UserControl_ParentControls);`（与 `For Each c In UserControl.ParentControls` 的既有产出一字不差）。
  - 链式成员 sanity：`Parent.Controls.Count` → `vb6_ComGetStringProp((void*)vb6_UC_ParentControls(), L"Count")`。
- 断言：`pd_parent_controls_ctl`（正向两条 + 反向 `VB3001`/`vb6_Parent_Controls`）；负控 `pd_parent_typo_still_reported`（`Paren` 必须照旧报 VB3001，证明只精确放行这一枚名字）。
- 记账：两个新夹具（`PdParentCtl.vbp` / `PdParentTypo.vbp`）同样未登记 emit-manifest，按同一口径延后。

#### 10.3.1 真 VB6 语义复核（2026-11-09 晚）—— 上面「Parent = 宿主容器」被探针推翻

oracle 工程 `.build/oracle/parent/`（`PdParentSem.vbp` + `Form1.frm` + `UCProbe.ctl` + `Module1.bas`，不入库）：Form1 直挂 `Command2`/`Label2`/`ucP1`，`Frame1` 内 `Command1`，`Frame2` 内 `ucP2`，`Pic1` 内 `Command3`；`UCProbe` 内一枚 `Label1`。VB6 6.00.9782 `/make` 后运行，读数写 `%TEMP%\pdparent_build\parent_sem.txt`。

| 问法 | 真 VB6 答案 | oracle 键 |
| --- | --- | --- |
| `Command2.Parent.Name`（Form 直挂） | `Form1` | F-CMD2-PARENT-NAME |
| `Command1.Parent.Name`（Frame 里） | **`Form1`**，不是 `Frame1` | F-CMD1-PARENT-NAME / F-CMD1-PARENT-IS-FRAME=False |
| `Command3.Parent.Name`（PictureBox 里） | `Form1` | F-CMD3-PARENT-NAME |
| `ucP2.Parent.Name`（UC 摆在 Frame 里） | `Form1` | F-UCP2-PARENT-NAME |
| 直接父容器（对照项） | 叫 **`Container`**：`Command1.Container.Name=Frame1` / `.Caption=FrameOne`、`Command3.Container.Name=Pic1`、`ucP2.Container.Name=Frame2` / `.Caption=FrameTwo` | F-*-CONTAINER-* |
| UC 内部裸 `Parent`（`ucP1`、`ucP2` 各一份） | `Form1`；`.Caption=HostForm`；`TypeName=Form1` | P1/P2-PARENT-* |
| `Label1.Parent Is Me`（UC 内层控件） | `True` | P1-INNER-PARENT-IS-ME |
| `Label1.Parent.Name` | **`UCProbe`**（设计器类名，不是实例名 `ucP1`） | P1-INNER-PARENT-NAME |
| `TypeName(Label1.Parent)` | `UCProbe` | P1-INNER-PARENT-TYPENAME |
| `Parent` 的静态类型 | `Object`（`Set f As Form = Parent` 过编译；Form 宿主 err 0、Frame 宿主 err 13） | P1-PARENT-AS-* |
| 窗体自己 | `Form1.Parent` = Nothing（err 91） | 探针单独验过，未入表 |

三层语义（定案）：

1. **普通控件 / 容器控件 / UC Extender 的 `Parent` = 顶层窗体** —— 与挂在 Frame/PictureBox 里无关；"直接父容器"那一问是 `Container`。
2. **UC 内部裸 `Parent` = 顶层窗体**（UC 摆在 Frame 里也答 Form）。既有 `vb6_UC_ParentObject()` 的 GA_ROOT 口径**本来就对，不动**。
3. **UC 内层控件的 `Parent` = 该 UserControl 对象本身**（`Is Me` True）；对它问 `.Name` 得设计器类名，问 `.BackColor` 等属性照常可读。

实现（五处，全部落在既有单一答案源上）：

| 文件 | 改动 |
| --- | --- |
| [uc_hostmodel_getprop.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/rtl/core/vb6forms/uc/detail/uc_hostmodel_getprop.inc) | ① 原 `Container \|\| Parent` 合并臂拆开：`Container` 保持原逻辑（直接父容器；父是 UC 则交 Extender），新增独立 `Parent` 臂（父窗口是 UC 宿主 → 交 `r->me`；否则 `GetAncestor(GA_ROOT)`；窗体自己 → Empty）。② 新增 UC **实例**的 `Name` 臂，必须排在「实例指针归一成宿主 HWND」**之前**（归一后就分不出 `ucP1.Name` 与 `Label1.Parent.Name`），只认 `rn->hwnd` 非空的有窗实例。③ 新增 `BackColor`/`ForeColor` 臂（原答 438）。 |
| [uc_hostmodel.c](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/rtl/core/vb6forms/uc/uc_hostmodel.c) | `vb6_Host_TypeNameOf` 在 `findByHwnd` 之后补 `findByInstance`（`!IsWindow` 防撞车）：`vb6_ho_findInstance` 只认无窗宿主，有窗 UC 的实例指针原先答 `Object`。 |
| [cgen_expr_member_precheck.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_member_precheck.inc) | `.ctl` 里裸 `Parent.<其它成员>`（`Name`/`Caption`/`BackColor`/`hWnd`…）立 COM 标记，作用于 `vb6_UC_ParentObject()`（晚绑定，静态类型 Object）；排在 `Parent.Controls` 那一支之后。修前这些会落到 m22 兜底拼成 `vb6_Parent_Name` 等未声明符号。 |
| [Form1.frm](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/pd_parentsem/Form1.frm) / [UCProbe.ctl](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/pd_parentsem/UCProbe.ctl) / [PdParentSem.vbp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/pd_parentsem/PdParentSem.vbp) | 新夹具，与 oracle 同形，27 行读数（12 宿主侧 + 7+7 UC 侧 + PS-DONE）。 |

修前 / 中间态 / 修后：

- **修前红**（旧 `C3.exe` 编 `tests/pd_parentsem`）：3 条 `error C2065`——`vb6_Parent_Name` / `vb6_Parent_Caption` / `vb6_Parent_BackColor` 未声明，整个 exe 编不出。
- **中间态**（只修发码、RTL 回退到 HEAD）：
  - `PS-CMD1-PARENT-NAME=Frame1`、`PS-CMD1-PARENT-TYPENAME=Frame`（应为 Form1）
  - `PS-CMD3-PARENT-NAME=Pic1`、`PS-UCP2-PARENT-NAME=Frame2`（应为 Form1）
  - `P1-INNER-PARENT-IS-ME=False`、`P1-INNER-PARENT-NAME=`（空串）、`P1-INNER-PARENT-TYPENAME=Object`（应为 True/UCProbe/UCProbe）
  - 且进程在 `P1-PARENT-BACKCOLOR` 处提前结束（只出 18 行，缺 P2 与 `PS-DONE`）——`Parent.BackColor` 在旧宿主模型里答 438。
- **修后绿**：27 行逐字命中，与 oracle 可比子集**一一相同**（`PS-*` ↔ `F-*`、`P1/P2-*` ↔ `P1/P2-*`）。颜色只断言「窗体 BackColor = UC 宿主 BackColor」（VB6 两者都是 `-2147483633`；C3 同一语义用 `15790320` 表示，故不钉字面值）。

断言：`run_tests.ps1` 新增 `pd_parentsem`（`Test-Vbp`，27 条 needle）+ `pd_parentsem_emitc`（4 条发码形状：`ComGetStringProp/ComGetIntProp((void*)vb6_UC_ParentObject(), …)` 与内层控件 `ComGetObjectProp(vb6_hwnd_Label1, L"Parent")`）+ `pd_parentsem_emitc_absent`（`vb6_Parent_Name/Caption/BackColor` 不得出现）。`Parent.Controls` 的 7 条仍由 10.3 的 `pd_parent_controls_ctl` 与负控 `pd_parent_typo_still_reported` 覆盖。

回归：`check_rtl_embedded.ps1` / `check_rtl_proto_arity.ps1` OK；ucSlider 像素门 **PASS**（A1=0.899 ≥0.80、A2=1135、B=0、D stable）；ucTreeView 五窗像素门 **PASS**（2026-11-09 深夜整跑 `EXIT=0`，读数与本刀之前记录逐项一致：white=110072 / bg240=803 / dark=586 / hl=6849 / vivid=8497 / chk240=82.9% / label240=90.6% / sb240=93.8% / sbHover=6/6 / Contact `bg65=118852`；五窗 open+sync 与三轮重复点击 `sync=True` 全过）。

> 本节初稿把「`Corporate Hierarchy` clickSync 失败」记为「既有问题、非本刀回归」，**判法与结论都已作废**（HEAD 上复现只排除 `c9f35cda`，排除不了 R1/R2/R3/R4/Randomize）。重做的归因见 §10.3.2。

已知缺口（本刀不碰）：`vb6_TypeOf` 仍是恒 0 的桩 ⇒ 夹具没断言 `TypeOf Parent Is Form/Frame`；标准控件自身的 `.Name` 仍答空串（宿主表未登记普通控件），故 Container 一侧只用 `.Caption` 钉。

#### 10.3.2 `Corporate Hierarchy` clickSync 失败的二分归因（2026-11-09 深夜）

口径：每点 `git checkout <sha> -- src` → `ninja -C .build c3`（`CMAKE_BUILD_PARALLEL_LEVEL=2`、串行）→ **用这枚新 C3** 重编 `tests\ucTreeView\Proyecto1.vbp`（`--arch x86`）→ 只跑 `Corporate Hierarchy` 一条（`.build\bisect\clickprobe.ps1`，与门同一条 `SendMessageTimeout(BM_CLICK, 3000, SMTO_ABORTIFHUNG)`）。

| 提交 | 夹具 exe | 读数（同一判据 `BM_CLICK` 3s 同步） |
| --- | --- | --- |
| `ba4ed971`（PhotoDemon 开工前，区间下界） | 849408 | **红** `clickSync=false 3013ms`（**冷网络**；同一点暖缓存驱动的复跑为绿） |
| `7495c8ff` | 849408 | 绿 `278ms` |
| `a7fc9137`（R2+R3 数组类型） | 849408 | 绿 `298ms` |
| `fdbd29e5`（R4 可创建接口） | 849408 | 绿 `308ms` |
| `46e22ed1`（R1 工程级 Public 变量） | 849408 | 绿 `396ms` |
| `de54cd09`（R1 包边界） | 849408 | 绿 `306ms` |
| `43528bb8`（R4 回归修复） | 849408 | 绿 `327ms` |
| `2e19c556`（内建常量 8 条） | 849408 | 绿 `373ms` |
| `2600664d`（Randomize 语句形态） | 849408 | 绿 `314ms` |
| `e21b5b5c`（UC 内建 Parent 首刀） | 849408 | 绿 `294ms` |
| `84a74b69`（VBA.Math） | 849408 | 绿 `319ms` |
| `d680b4d2`（Do 体 Dim/Const） | 849408 | 绿 `308ms` |
| `fbc26f64`（Randomize/Rnd） | 849408 | 绿 `324ms` |
| `c9f35cda` / `3b47cc96`（HEAD） | 850432 | 绿 `291–327ms`；冷网络下 `3016ms` 红；**五窗像素门整跑 PASS** |

同一 HEAD 二进制、同一夹具，**只动 `JerarquiaEmpresa.xml` 一处**：

| 条件 | 读数 |
| --- | --- |
| 原 15 条 `src`（WinINET 命中缓存） | `true` 327ms |
| 同一 URL 加 `?cb=<随机>` 破缓存（强制回源） | **`false` 3015ms** |
| `src` 换成产物内本地 `img\customer.png`（不走网络） | `true` 258ms |
| 删掉整个 `<images>` 段 | `true` 320ms |
| 其余四枚按钮（均无远程图） | `true` 51–253ms |

成因（插桩实测；`.build\bisect\dbg` 是带 `TraceLog` 的夹具副本，不入库）：`Corporate Hierarchy` → `Form3.Show` → `Form3_Load` → `LoadFromXML(App.Path & "\JerarquiaEmpresa.xml")` → `CargarImagenes` → `ImageListAddImage` → `LoadPictureEx`（[ucTreeViewPlus.ctl](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/ucTreeView/ucTreeViewPlus.ctl#L2365-L2383)）→ **15 条 `https://randomuser.me/...` 头像**：`Microsoft.XMLHTTP` 异步 GET + `While readyState <> 4: DoEvents: Wend`。trace 逐图 `X1 → X2(rs=1) → X5(rs=4, st=200) → X6(ub=…) → X7` 共 15 组，末尾 `F4 after loadxml` / `F5 done`：暖缓存整段 **0.35s**、冷缓存 **3.55s**、网络打嗝时 **>30s**（有一次 30s 连 Form3 窗口都没出来）。BM_CLICK 的窗口过程就卡在这个 `DoEvents` 自旋里，3s 上限一到便报「事件没到按钮过程」——**是网络抖动，不是编译器语义**。

结论：

1. **区间内没有「首个变红的提交」**。反证提交 `ba4ed971`（早于 `7495c8ff`）用同一夹具、同一判据复现同一条报错（`contentA2 ... BM_CLICK 3s 未返回` + 同样的 MISS 行），且与 HEAD 在同等网络条件下逐字相同。
2. 该判据**天生依赖外网**（15 张头像串行下载），3s 预算贴着实测 0.35–3.55s 的分布 ⇒ 是**抖动的门**，不是稳定的绿/红。
3. `C3_HOST_MEMBER_MISS site=SetProp name="Font"` 是**另一件事**（boot 期就有、五窗各 2 行、与点击无关，`ba4ed971` 同样有）：[ucTreeViewPlus.ctl:2744](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/ucTreeView/ucTreeViewPlus.ctl#L2744) 的 `Set Text1.Font = Me.Font`（`UserControl_ReadProperties`）走 `vb6_ComSetProp(hEdit, L"Font", …)`；[uc_hostmodel_setprop.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/rtl/core/vb6forms/uc/detail/uc_hostmodel_setprop.inc) 只对 `vb6_uc_isFont(obj)` 那一支按「字体对象改 Size/Bold/…」处理，对**控件**的 `Font` 没有臂，于是落到 `vb6_ho_reportUnknown(L"SetProp", …)` 打一行 MISS 并**忽略写入**。`Text1` 是 `Visible=0` 的量字盒，不影响像素读数；属存量缺口。

处置计划（门侧已落地，见 §10.3.3；产品侧单列）：

- 门侧（主）：15 张头像落到 `tests\ucTreeView\img\`，`JerarquiaEmpresa.xml` 的 `src` 改本地相对路径 —— 改动 + 重拍参考，见 §10.3.3。
- 产品侧（次，单独一刀）：`uc_hostmodel_setprop.inc` 补 `Font` 臂（`Set ctl.Font = f` 按 `f` 的 LOGFONT 落 `WM_SETFONT`；共享 StdFont 别名/回写另记）。

#### 10.3.3 夹具脱网 + 参考重拍（2026-11-09 深夜）

改动（非 src，`[skip ci]`）：

- 15 张头像固定到 [tests/ucTreeView/img](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/ucTreeView/img)：`avatar_<key>.jpg`（与 `JerarquiaEmpresa.xml` 的 15 个 `key` 一一对应，文件名 ASCII）。
- [JerarquiaEmpresa.xml](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/ucTreeView/JerarquiaEmpresa.xml) 的 `src` 由 `https://randomuser.me/api/portraits/...` 改为 `img\avatar_<key>.jpg`（相对路径；`LoadPictureEx` 的本地分支走 `GdipLoadImageFromFile`，相对 CWD 解析，三个入口都以 exe 目录为 CWD）。`tests/` 下其余联网夹具见 §10.3.4。

重拍流程（旧参考保留，不覆盖）：

1. 真 VB6 重编改后夹具：`uctv_vb6_baseline.ps1 -OutDir .temp\vb6ref_fixed\ucTreeView`（exit=0，`Proyecto1.exe` 新产物）。
2. C3（HEAD 产物 `.build\C3.exe`，src 未变）重编同一夹具 → `.temp\uc_fixed\out`，并把 `img/`、`ejemplo.xml`、`JerarquiaEmpresa.xml`、`menu.xml` 拷进产物目录（与 `Test-VbpBuild` 同口径）。
3. 重拍参考：`UCTV_SAVE_FORMS=1 uctreeview_px.ps1 -Exe <C3产物> -OutDir .temp\vb6ref_fixed\px -RefExe <新VB6> -RefJson .temp\vb6ref_fixed\ref_metrics.json`（EXIT=0）。

旧参考值 → 新参考值（`ref_metrics.json` 全字段逐项比对，**0 处差异**）：

| 窗（cap） | textRows | vivid | photo | avatar | green | greenRim | firstTextW |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Form3（Corporate Hierarchy） | 10 → 10 | 21768 → 21768 | 9292 → 9292 | 6498 → 6498 | 13790 → 13790 | 87 → 87 | 187 → 187 |
| Form5（Menu / Tools） | 15 → 15 | 41664 → 41664 | 56127 → 56127 | 108113 → 108113 | — | — | 77 → 77 |
| Agenda tema oscuro（Contact） | 1 → 1 | 7120 → 7120 | 10955 → 10955 | 9793 → 9793 | — | — | 331 → 331 |
| Drag & Drop | 5 → 5 | 2161 → 2161 | 2484 → 2484 | 927 → 927 | — | — | 98 → 98 |
| Tree Explorer | 5 → 5 | 2354 → 2354 | 1608 → 1608 | 1225 → 1225 | — | — | 78 → 78 |

参考 PNG 也做了字节级对照（旧 `.temp\vb6ref\shots` vs 新 `.temp\vb6ref_fixed\px`）：`Corporate Hierarchy`、`Contact`、`Menu/Tools`、`Drag & Drop` 四张 **SHA256 相同**；`Tree Explorer` 差 7 字节（该窗无头像，属截图时序，不是图片内容）。

结论（与「randomuser 每次返回不同人像」的前提相反，有据可查）：`https://randomuser.me/api/portraits/<gender>/<n>.jpg` 是**每个 URL 固定的一张图**，不是每次随机。证据：① 重拍读数与旧参考一字不差（若内容变，photo/avatar/vivid 必漂）；② 四张参考 PNG 字节相同。此前 `repeatcontent[Contact]` 三轮 9950/8579/8739 的波动换成本地图后**仍在**（`9950 → 8579 → 8739`），说明那是 C3 重复开窗的绘制状态（既有、非本刀），与夹具联网无关；该判据阈值 3000，三轮都过。

判据改动：**0 条**。语义照旧（textRows/vivid/bg65/avatar/photo/green/greenRim），阈值不放松；参考读数未变，故无基线可调。像素级不变量原样不变（`chk240=82.9%`、`label240=90.6%`、`sb240=93.8%`、`sbHover=6/6`、`sbMid=6/6`、`sbLeft=6/6`、`white=110072`、`bg240=803`、`dark=586`、`hl=6849`、`vivid=8497`）。

稳定性：CI 同款调用（不带 `-RefJson`）连跑 3 次，79/79/80s，三次读数逐字相同、`rc=0`、`PASS` 全过。

#### 10.3.4 `tests/` 联网夹具清单（只查不改）

| 夹具 | 何时联网 | 进 CI | 失败形态 |
| --- | --- | --- | --- |
| `tests\ucTreeView\JerarquiaEmpresa.xml` + `LoadPictureEx`（XMLHTTP） | Form3 `Load` 启动即拉 15 张 | **是**（vbp 分片 → `Test-VbpBuild ucTreeView -PixelAssert`） | 卡 3s → 像素门红（本次抖动的源；已按 §10.3.3 脱网） |
| `tests\urlpic\UrlPicMain.bas` + `run_tests.ps1` 的 `urlpic_net`（randomuser `women/3.jpg`） | 真跑真 GET | **是**（`bas` 分片） | 3s TCP 探不通 → SKIP；探通但 429/内容变 → FAIL |
| `tests\urlpic` 的 `urlpic_local`（`127.0.0.1:8765` + 本地 `serve.ps1`） | 仅 loopback | 是 | 端口占用 → SKIP |
| `tests\test_com_events_winhttp\Form1.frm:37`（WinHTTP 访问 `www.example.com`） | 仅运行 exe | 否（只 `Test-VbpBuild`，不跑） | 不触发 |
| `tests\ucSegmentedControl\ucSegmentedControl.ctl` 的 XMLHTTP 分支 | 代码在，工程内无远程 `src` | 是（GUI 表），但不触发 | 不触发 |
| `tests\ucTreeView\exportado.xml`（4×icons8） | 无 Form 加载、不在 `-AssetFile` 名单 | 否 | 死资产 |

待处理：`urlpic_net` 是下一处同类隐患（DEFAULT 情况下网络不通会 SKIP 而不是红，但内容/限流变化会红）；按同一口径改成固定本地图或固定字节回放后再处理。

#### 10.3.5 `C3_HOST_MEMBER_MISS site=SetProp name="Font"` 单刀

**oracle（真 VB6 参考 exe，Form3 `Corporate Hierarchy`，120dpi，`WM_GETFONT`+`GetObject` 读 LOGFONT）**：

```
Form3 窗体     -13|w400|cs1|q0|MS Sans Serif      (窗体 Ambient 默认)
UC 宿主窗口    -16|w700|cs0|q0|Segoe UI           (m_Font，见 Form3.frx 的 font 属性)
Text1(隐藏)    -16|w700|cs0|q0|Segoe UI           (与 UC 宿主窗口逐字相同)
```

⇒ `Set Text1.Font = Me.Font` 的效果：**字面/字号/字重/字符集全部跟随**，质量档 DEFAULT（q0）；`Me.Font` 是 UC 自己的 `Font` 属性（`UserControl_ReadProperties` 里 `ReadProperty("Font", Ambient.Font)`，本夹具由 `Form3.frx`:0000 那枚设计期 StdFont 覆盖），不是窗体 Ambient。

**frx 证据**：`Form3.frx` 偏移 0 的 36 字节 StdFont 块里 Weight=400；真 VB6 稳态读回 700，是因为 VB6 的 `Set` 共享**同一枚 StdFont 对象**，之后 `UserControl.FontBold = N.FontBold`（`ucTreeViewPlus.ctl:1665-1669`，每个粗体节点一次）改的就是这枚对象的 Bold，Text1 跟着变。

**实现**（[uc_hostmodel_setprop.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/rtl/core/vb6forms/uc/detail/uc_hostmodel_setprop.inc) 第三百五十二刀）：宿主模型 `SetProp` 补 `Font` 臂 —— 值按 `vt=9/13` 解包取字体对象，`vb6_UC_ApplyFontToHostWindow`（`DEFAULT_QUALITY`，对齐 oracle 的 q0）落到目标 HWND 的 `WM_SETFONT`；取不到字体对象时仍走原来的 `reportUnknown`。

**取证**：

- MISS 消失：同一夹具 clickprobe 的 stderr 里已无 `C3_HOST_MEMBER_MISS site=SetProp name="Font"`（改前每个 UC 实例 boot 期各 2 行）。
- 效果：C3 侧 Text1 `WM_GETFONT = -16|w400|cs0|q0|Segoe UI`（改前是设计期默认 `MS Sans Serif`）；字面/字号/字符集/质量档与 oracle 一致。
- 无回归：五窗像素门 `PASS`，读数与改前逐项相同（`chk240=82.9% label240=90.6% sb240=93.8% sbHover=6/6`，Form3 `avatar=6356 photo=8997 green=13150 greenRim=91`，Contact `avatar=9950 bg65=118862`）。

**已知残差（单列，不在本刀；对象模型级缺口）**：VB6 的 StdFont 是**共享对象**（`FontBold`/`FontUnderline` 直接改那枚对象的 Bold/Underline），C3 把 `FontBold`/`FontUnderline` 建模为独立标志位（`vb6_UserControl_FontBold`，绘制时与字体对象合成），所以赋值那一刻的拷贝不会跟着后续 Bold 翻转走 —— 稳态下 C3 的 Text1 是 w400，真 VB6 是 w700。只影响就地编辑框的字重（`Text1` 为 `Visible=0` 的量字/编辑盒，当前判据不覆盖像素），不影响树区绘制。要完全对齐得把 `FontBold` 回写到字体对象（或给字体对象加绑定窗口表），会动到所有 UC 的绘制路径，另开一刀并重跑五窗。

### 10.4 `Math` 18 条：不是 tB 专有，是 VB6 自带 VBA 类型库的 `VBA.Math`

- **定案（推翻旧记的「疑似 tB-only」）**：`Math` 是 VBA 类型库里真实存在的**标准模块**（不是类/对象），VB6 隐式引用 VBA 库即拥有它；PhotoDemon 的 `INSTALL.md` 明确是「把 `.vbp` 载入 VB6 IDE 编译」，18 处因此本来就是合法 VB6。twinBASIC 的 `Math` 模块只是同一族，不是来源。
- 语料形态（复核计数）：`Modules\ComplexNumbers.bas` + `Forms\Effects_Distort_Droste.frm`，**代码位 18 处**（`Math.` 合计 22，含 3 处注释）：`Log` 9 / `Abs` 4 / `Cos` 4 / `Sin` 1。全部是静态限定调用，没有 `Set m = New Math` / `Dim m As Math`；`PhotoDemon.vbp` 里只有 shell32 与 stdole2 两条 `Reference=`，工程内无 `Math.bas` / `Math.cls`。旧记里写的 `Sqr` / `Round` 来自 C3 的测试夹具，不是 PhotoDemon 语料。
- **真 VB6 探针（2026-11-09，最小 BAS 工程逐个 `/make`）**：
  - `Math.Sin` / `Math.Cos` / `Math.Log` / `Math.Abs` / `Math.Round` / `Math.Sqr` **都能编过**（`VBA.Math.Sin` 也能）。
  - 成员矩阵：`Abs/Atn/Cos/Exp/Log/Rnd/Round/Sgn/Sin/Sqr/Tan` 存在；`Math.Int` / `Math.Fix` / `Math.Randomize(1)` **不存在**（「未找到方法或数据成员」）；`Math.NoSuchMember` 同样报未找到 ⇒ 说明编译器是在**真切解析**这个模块，不是忽略限定前缀。
  - `Dim m As VBA.Math` → 「Visual Basic 中不支持自动化类型」⇒ 它是模块，不是可实例化的类，也不能 `New`。
- **C3 现状复核（`.build\C3.exe`，最小探针）**：`Math.Sin/Log/Abs/Cos/Round/Sqr` 六种调用 `rc=0`、无 finding（`84a74b69` 生效）。两处小残差（未改代码）：① 白名单里的 `fix`/`int` 会走到 `vb6_vb6_Fix` / `vb6_vb6_Int`，`LNK2019` 失败而不是像 VB6 那样给「未找到方法或数据成员」；② `Dim m As VBA.Math` 在 C3 下能过，真 VB6 报「不支持自动化类型」。两条都不影响 PhotoDemon 的 18 处。
- 修法（与已有 `VBA.<成员>` 折叠同一条通路，零新语法）：
  1. [semantic_analyzer_util.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/semantic_analyzer_util.cpp#L789-L794)：`isDocumentHostObject` 认 `math`，**仅限定位**（与 `vba` 同款；值位的裸 `Math` 仍报 VB3001）。
  2. [cgen_expr_member_precheck.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_member_precheck.inc#L186-L194)：`Math.<成员>` 并入 `VBA.<成员>` 那一支，成员白名单 = `abs/atn/cos/exp/fix/int/log/rnd/round/sgn/sin/sqr/tan`（拼错的成员不会因前缀是 `Math` 就变合法；其中 `fix`/`int` 已在 §10.4.1 收紧删除）。
- 读数：修前 6 条 VB3001（`Math`）+ `vb6_Math_Log/Abs/Cos/Sin/Sqr/Round`（RTL 无此符号，真编译必 C2065）；修后 stderr 0 行，产物 `vb6_Log(1.0)` / `vb6_Abs((-4.5))` / `vb6_Cos(0.0)` / `vb6_Sin(0.0)` / `vb6_Sqr(9.0)` / `vb6_Round(2.6000000000000001, 0)`。
- 运行读数（真编译）：`VM-LOG=0` / `VM-ABS=4.5` / `VM-COS=1` / `VM-SIN=0` / `VM-SQR=3` / `VM-ROUND=3`。
- 断言：`test_vbamath`（Compile+Run 7 行读数）+ `vbamath_emitc_folds_to_builtins` 形状钉 + `vbamath_no_module_symbol` 反向钉；负控 `vbamath_bare_is_undeclared`（值位 `Set m = Math` 必须仍报 VB3001）。
- 记账：两个新夹具按同一口径延后登记 emit-manifest。

#### 10.4.1 收紧两处「过宽」（2026-11-09 夜）

- 变更（都是**收紧**，方向与上面的「补能力」相反）：
  1. 新增 [vba_math.hpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/common/vba_math.hpp)：`VBA.Math` 真实成员 11 个 + VBA 标准模块 8 个，两张名单只此一处；发码层 [cgen_expr_member_precheck.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_member_precheck.inc#L202-L209) 改问它，并删掉 `fix` / `int`。
  2. [semantic_analyzer_expr.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/semantic_analyzer_expr.cpp#L203-L216)：`Math.<成员>` 非真成员 → `VB3045`「未找到方法或数据成员」（语义层，错误级，不发 C）。接收者同名符号/模块在场时不判。
  3. [semantic_analyzer_typeref.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/semantic_analyzer_typeref.cpp#L36-L51)：`As` 位置出现 VBA 标准模块 → `VB3046`「Visual Basic 中不支持自动化类型」。限定写法（`VBA.Math`）一律判；裸名（`Math`）仅当工程无同名类型时判。**没有**禁整个 `VBA.*` 前缀：`VBA.ErrObject` / `VBA.Collection` 等类名照旧。
- 真 VB6 oracle（本轮补测）：`Dim m As Math` 与 `Dim m As VBA.Math` 同报「Visual Basic 中不支持自动化类型」；`VBA.Fix(1.7)` / `VBA.Int(1.7)` **合法**（`/make` 生成 exe）；工程里建 `Math.cls` 被 VB6 拒绝（「不能加载」，与库名冲突）⇒ 不存在用户类同名 Math 的合法场景。
- A/B 读数（同一夹具，仅换 C3 产物；`--emit-c`）：
  - `Math.Fix(1.7)` + `Math.Int(1.7)`：修前 `rc=0`、无 VB3045、C 里出现 `vb6_vb6_Int`；修后 `rc=1`、VB3045 ×2、C 里无 `vb6_vb6_Fix` / `vb6_vb6_Int`。
  - `Dim m As VBA.Math` + `Dim m As Math`：修前 `rc=0`、照常出 C；修后 `rc=1`、VB3046 ×2、不出 C。
  - `--syntax-only` 同报 VB3045 / VB3046 ⇒ 锁定在语义阶段，不是发码期报错。
  - 正向对照：`test_vbamath.bas` `rc=0`，`vb6_Log(1.0)` 仍在、`vb6_Math_` 仍无，7 行运行读数不变；同夹具新增 `Dim xForm As Form` / `Dim xErr As VBA.ErrObject` 也通过。
- 引用点核查：全仓 `vb6_vb6_` 只出现在注释里（无真实符号）；删名单后 `Math.Fix` / `Math.Int` 不再进入发码。
- 断言（`tests/run_tests.ps1`）：`vbamath_fixint_semantic_stage`（阶段+错误码）、`vbamath_fixint_not_found`（错误码 + 两处文案 + 反向钉 `vb6_vb6_Fix`/`vb6_vb6_Int`）、`vbamath_typeuse_semantic_stage`、`vbamath_typeuse_module_as_type`。`Test-CompileFail` 扩了可选 `-Absent`、`-Needles` 收数组（既有调用点兼容）。
- **新残差（未修，方向相反 = 欠）**：`VBA.Int(1.7)` 仍发 `vb6_vb6_Int` → `LNK2019`；真 VB6 接受该写法（exe 生成），`VBA.Fix` 则正常发 `vb6_Fix`。根因未定位（`Int` 的符号解析与 `Fix` 不对称），单独记账，不在本刀范围。
- 记录：新的 `Test-SyntaxFail` 未新增（复用既有 [run_tests.ps1](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/run_tests.ps1#L3803-L3818) 那枚）；`check_test_helper_integrity`（H1..H7）与 `check_diag_id_exclusivity` 均 OK。


### 10.5 过程级 `Dim/Const`（4 条）：`Do` 体声明在 `Loop While` 条件里

- 语料形态（4 条，两个文件同一形）：
  - `Classes\pdXCF.cls`：`Do` 体内 `Const ERR_BUFF_TOO_SMALL` / `Dim ldResult`，`Loop While (ldResult = ERR_BUFF_TOO_SMALL) ...`。
  - `Modules\UserControl_Support.bas`：`Dim positionFailed` / `Dim failsafePosition` 在 `Do` 体内，`Loop While positionFailed And (Not failsafePosition)`。
- 根因：语义层 [visit(DoLoopStmt)](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/semantic_analyzer_stmt.cpp#L126-L136) **先分析条件、后分析体**；而 `Dim`/`Const` 是 visit 到声明语句时才 `define` 进当前（过程级）作用域 ⇒ 回边条件里两个名字查不到。VB6 的 `Dim`/`Const` 写进任何块都算整个过程的声明，条件里引用它们是合法的。
- 修法：两行换序（先 `analyzeStmtList(node.body)`，后 `analyzeExpr(*node.condition)`）。**只改分析顺序**，不动 AST、不动发码、不动作用域结构。
- 读数（`tests\test_procscope_loopdim.bas`，`Dim inner` + `Const TOO_SMALL` 都在 Do 体、条件里各用一次）：修前 2 条 VB3001（`TOO_SMALL` / `inner`）→ 修后 0 条；真编译运行 `PS-LOOP=3,6,6`（3 轮后退出，读数符合手算）。
- 断言：`test_procscope_loopdim`（Compile+Run）+ `procscope_loopdim_no_vb3001`（反向钉 `VB3001` / `未声明的标识符`）。
- 回归：`-Category syntax` **PASS=187 FAIL=0**。
- 剩下未动的邻近形态（**已知、未报**）：`While <cond>`（`Wend`）与 `If <cond>` 仍是先条件后体；VB6 同样把后置声明当过程级，若后续语料出现同形再按本条处理，不在本轮扩大。

### 10.6 本轮 38 条的收口读数（等 CI 复核）

| 类 | 条数 | 提交 |
| --- | --- | --- |
| 内建常量表（`vbCFFiles` 7 + `vbAsyncStatusCodeEndDownloadData` 1） | 8 | `2e19c556` |
| `Randomize -1` | 1 | `2600664d` |
| UserControl 内建 `Parent`（`Parent.Controls` ×7） | 7 | `e21b5b5c` |
| `VBA.Math`（`Math.*` ×18） | 18 | `84a74b69` |
| 过程级 `Dim/Const`（Do 体 → `Loop While`） | 4 | 本提交 |
| **合计** | **38** | |

预期下一次探针 `VB3001 = 0`（若仍有残留，只能是本轮之外的第三形态）。`VB3003` / `VB3044` 仍按各自账走。

### 10.7 过程级 `Dim/Const` 4 条：调查定案（2026-11-09 夜，**只调查、未改代码**）

旧读数出处：`.build\pd_ci\r1\pd_emit.err`（2026-11-09 18:31，`EMIT_EXIT=0`、`STDERR_LINES=238`、其中 `error` 0 条 / `warning` 95 条）。

#### 10.7.1 逐处清单（4 条，一条不落）

| # | 文件 | 过程 | 声明位置（原文） | 引用位置 | C3 旧报（VB 号 + 文本） | 真 VB6 |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | `Classes\pdXCF.cls` | `Friend Function IsFileXCF`（276 行；`Do` 在 331，嵌于 `If useGZMode`→`If Files.FileLoadAsByteArray`） | 337 行 `Dim ldResult As Long`（**Do 体内**） | 342 行 `Loop While (ldResult = ERR_BUFF_TOO_SMALL) ...`；循环后 345 行再用一次 | `pdXCF.cls(342,29): warning VB3001: 未声明的标识符: 'ldResult' (可能来自其他模块)` | 接受 |
| 2 | 同上 | 同上 | 336 行 `Const ERR_BUFF_TOO_SMALL As Long = 3`（**Do 体内**） | 342 行同上 | `pdXCF.cls(342,40): warning VB3001: 未声明的标识符: 'ERR_BUFF_TOO_SMALL' (可能来自其他模块)` | 接受 |
| 3 | `Modules\UserControl_Support.bas` | `Public Sub ShowUCTooltip`（1035 行；`Do` 在 1231，直接位于过程体） | 1272 行 `Dim positionFailed As Boolean: positionFailed = False`（**Do 体内**，冒号连写） | 1287 行 `Loop While positionFailed And (Not failsafePosition)` | `UserControl_Support.bas(1287,16): warning VB3001: 未声明的标识符: 'positionFailed' (可能来自其他模块)` | 接受 |
| 4 | 同上 | 同上 | 1237 行 `Dim failsafePosition As Boolean`（**Do 体内**） | 1287 行同上 | `UserControl_Support.bas(1287,40): warning VB3001: 未声明的标识符: 'failsafePosition' (可能来自其他模块)` | 接受 |

#### 10.7.2 定性：四条同一根因 = 语句分析**顺序**，不是作用域建模缺失

四个候选逐一排除（每一条都有探针读数背书，见 10.7.3）：

| 假设 | 判据 | 结论 |
| --- | --- | --- |
| 过程内声明没进局部符号表（作用域建模缺失） | S2（两个块各 `Dim dup`）C3 报 `VB3002 重复声明: 'dup'` ⇒ `Dim` 确实 define 进了**过程级**作用域，且**没有**块作用域 | **否** |
| 块内声明被语法层拒 | S1/S3/S5b/S6 在 `--syntax-only` 与 `--emit-c` 下均 `rc=0`、0 诊断 | **否** |
| `Const` 与 `Dim` 路径不同 | 四条里 2 条 `Dim`、2 条 `Const`，是同一处 `visit(DoLoopStmt)` 顺序问题一次性收掉的（`d680b4d2`） | **否** |
| `Static` 相关 | S4 两处 `Static` 均通过，产物为 `static int32_t n = 0;` / `static int32_t m = 0;`（块内 `Static` 抬到函数级，与 VB6 一致） | **否** |

唯一成因：`visit(DoLoopStmt)` 历史上「先 `analyzeExpr(condition)`、后 `analyzeStmtList(body)`」，而 `Dim`/`Const` 是 visit 到声明语句时才 `define` ⇒ 回边条件查不到名字。**旧日志 18:31 早于修复 `d680b4d2`（20:18），所以那 4 条是修复前读数，不是当前缺口。**

#### 10.7.3 真 VB6 oracle（12 个最小探针，`.temp\procdim\vb6\`，未入库）

| 探针 | 形态 | 真 VB6（6.00.9782） | C3（HEAD `c0d57af4`） |
| --- | --- | --- | --- |
| S1 | `If` / `Do`（`Loop While False`）/ `For` 体内各 `Dim` 一只，块外读 | 编译过；`S1 inIf=11 inDo=22 inFor=33` | 0 诊断 |
| S2 | 两个块各 `Dim dup` 同名 | **编译错**：`行 11 : 当前范围内的声明重复` | `VB3002 重复声明: 'dup'`（rc=1） |
| S3 | `If` 体 `Const K_IF` + `Do` 体 `Const K_DO`，`Loop While i < K_DO`，块外读两者 | 编译过；`S3 K_IF=5 K_DO=6 i=6` | 0 诊断 |
| S4 | 过程级 `Static n` + `If` 块内 `Static m`，跨调用 | 编译过；`bump=1/2/3`、`blockstatic=1/2` | 0 诊断；产物 `static int32_t n = 0;` / `static int32_t m = 0;` |
| S5 | `Dim a()` + `ReDim` + `Preserve`，块外读 | `S5 a(1)=7 ub=5` | 0 诊断 |
| S5b | `Do` 体内 `Dim b()`（动态数组），块外 `ReDim` | `S5b b(2)=9 bub=4` | 0 诊断 |
| S6 | `With` 块内 `Dim`，块外读 | `S6 insideWith=44 count=1` | 0 诊断 |
| S7 | `Dim a(1 To 3)` 之后 `ReDim a(1 To 5)` | **编译错**：`行 5 : 数组维数已定义` | **rc=0、0 诊断 ⇒ 过宽**（见 10.7.6） |
| S8 | `While i < LIMIT` 条件引用**体内** `Const LIMIT` | **编译错**：`行 5 : 变量未定义` | `VB3001: 'LIMIT'`（方向同、等级不同） |
| S9 | `If flag = ON_VAL` 条件引用**体内** `Const ON_VAL` | **编译错**：`行 6 : 变量未定义` | `VB3001: 'ON_VAL'` |
| S10 | `While Not done` 条件引用**体内** `Dim done` | **编译错**：`行 6 : 变量未定义` | `VB3001: 'done'` |
| S11 | **pdXCF 同形**：`Do` 体内 `Dim ldResult`，`Loop While (ldResult = 3) And ...`，循环后再用 | 编译过；`S11 n=5 last=3` | 0 诊断；同形探针 P1b 真编真跑 `P1b=0,1600,False` |
| D1 | **风险面**：`Do While Not done` 前置条件引用体内 `Dim done` | **编译错**：`行 6 : 变量未定义` | **rc=0、0 诊断 ⇒ 过宽**（见 10.7.5） |

三条钉死的语义：

1. 块内 `Dim`/`Const`/`Static` 都是**过程级**：块外可读（S1/S3/S6）、跨调用保持（S4）、同名重复声明直接报错（S2）——**没有块作用域**。
2. VB6 是**单遍文本序**：名字必须在**使用点之前文本出现**过。`While`/`If` 条件在体之前（S8/S9/S10）⇒ 引用体内声明报「变量未定义」；`Loop While` 条件在体之后（S3/S11）⇒ 合法。**这条边界正好把 4 条与「不能放宽」分开**：pdXCF / UserControl_Support 的引用都落在 `Loop While`。
3. `ReDim` 与过程内 `Dim`：动态数组（`Dim a()` / `Dim b()`）任意块内声明后块外 `ReDim` 均合法（S5/S5b）；**定长数组不可 `ReDim`**（S7）。

#### 10.7.4 C3 现状对照

- 4 条对应形态全绿：P1/P2/P3/P4、P1b（pdXCF 形，真编真跑）、P9（ShowUCTooltip 形）、S3/S11 均 0 条 `VB3001`。
- `Modules\UserControl_Support.bas` **单文件 emit 直接复核**：`rc=0`、1287 行 **0 条**（全文件 56 条 `VB3001` 全是 `g_WindowManager` / `PDDebug` / `Fonts` 这类跨模块名，单文件口径的既有噪声，与第 3/4 条无关）⇒ 第 3、4 条在真实文件上确认收口。
- `Classes\pdXCF.cls` 在 dev 上**整体不可解析**（停在 1028/1055，见 10.7.6 第 1 项），所以 342 行只能用同形探针 P1b/S11 背书，无法在 dev 上直读该文件。

#### 10.7.5 改法候选与风险面（**仅列，不动手**）

| 候选 | 内容 | 影响面 | 风险 |
| --- | --- | --- | --- |
| A：维持现状（`d680b4d2` 的**无条件**换序） | 0 改动 | `visit(DoLoopStmt)` 一处 | **已实测过宽**：`DoWhileLoop` / `DoUntilLoop`（前置条件）也被换序，D1 真 VB6 报错而 C3 通过。属「多收 VB6 不收的写法」，不阻塞 |
| B：按 `loopKind` 分流 | 只有 `DoLoopWhile` / `DoLoopUntil` 先体后条件；`DoWhileLoop` / `DoUntilLoop` 保持条件先 | 同 A，约 6 行 if；AST/发码不动 | 若语料里有 `Do While` 引用体内声明，收紧后会**新增**报错。本轮未扫全工程（用户口径：不整工程 emit），动手前应 grep 语料核一遍；现有用例 `test_procscope_loopdim` 与本报告 P1–P4/P9/S11 全在后测形，不受影响 |
| C：过程级声明预扫描（首次进过程先收集所有 `LocalDeclStmt`） | 一次性对齐「过程级声明」直觉 | 覆盖**全部**语句类型，不是一处 if | 会把 S8/S9/S10/D1 那类 VB6 真错变成通过（漏报面扩大），与 C3 现有文本序行为相反。**不建议** |

风险面总述：动的是 [semantic_analyzer_stmt.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/semantic_analyzer_stmt.cpp#L126-L141) 里 `Do`/`While` 两个 visit 的**分析顺序**，影响所有 `Do` 块的条件类型推断与未声明名诊断；**不触碰** `SymbolTable` 结构、**不触碰**发码。现有用例中与本条相关的只有 `tests\test_procscope_loopdim.bas`，它在 `Loop While` 形，候选 B 不会改它的读数。

#### 10.7.6 调查中发现的、**不属于**这 4 条的三件事（如实记录，均未动手）

1. **dev 缺 `ReDim` 无括号成员链**（真缺口，非本轮范围）：`ReDim m_Layers(i).lMask.cProperties(0 To n) As T`（pdXCF 1028 / 1055 行）在 dev 上报 `VB2003`（`(1028,35)` 起 9 条 + `(1055,145)` 起 9 条）。`probe/photodemon-emit` 已有修复（`if (cur_.kind != TokenKind::LeftParen) { if (cur_.kind == TokenKind::Dot) continue; break; }` + `parseReDimStmtList/parseReDimEntry`），dev **没有**。⇒ 18:31 那份日志是用带此修复的二进制跑的；dev 上 pdXCF 目前整体不可解析。
2. **定长数组 `ReDim` 过宽**：S7 真 VB6 报「数组维数已定义」，C3 `rc=0` 无诊断。
3. **未声明名严重度差异**：S8/S9/S10/D1 真 VB6 是**编译错误**，C3 只发 `VB3001` warning（`Option Explicit` 下也未升级）。方向一致、等级不同。

**这 4 条不是工程源码问题**：真 VB6 接受同形写法（S1/S3/S11 均 `/make` 出 exe 并运行出预期读数）。局限说明：未对 `PhotoDemon.vbp` 整工程跑真 VB6 `/make`，「接受」由同形最小探针背书。

### 10.8 分支收敛与 Do 循环分流（2026-11-09 夜）

#### 10.8.1 候选 B 落地：`Do` 循环按 `loopKind` 分流

- 改动：[semantic_analyzer_stmt.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/semantic_analyzer_stmt.cpp) 的 `visit(DoLoopStmt&)` 一处 —— 只有 `DoLoopWhile` / `DoLoopUntil`（回边形）先体后条件；`DoWhileLoop` / `DoUntilLoop`（前置形）保持条件先。不碰 `SymbolTable`、不碰发码。
- 提交 `84ab0e00`。
- 读数：正向 P1/P2/P3/P4/P1b/P9 与 `test_procscope_loopdim` 全 0 条 `VB3001`（P1b 真跑 `P1b=0,1600,False`）；负控 `test_procscope_dowhile_bad`（D1）与 `test_procscope_condref_bad`（While/Wend、If、Do While 三处）均报 `VB3001`，不再静默通过。

#### 10.8.2 dev ↔ probe 解析器状态收敛

probe 分支上那批解析器补口原本压在**一个** WIP 提交 `5b01143a`（6 文件、无测试）。在 dev 上按主题拆成 4 个提交，每个带自己的夹具与断言：

| # | 补口 | dev 提交 | 夹具 | 修前读数 | 修后读数 |
| --- | --- | --- | --- | --- | --- |
| 1 | 注释行以「空白 + `_`」结尾时延续到下一物理行 | `720132fe` | `tests\test_comment_continuation.bas` | `--syntax-only` rc=1：`VB2002 unexpected token at module level: ByVal` | rc=0，真编真跑 `CC=1` |
| 2 | `Option Compare Binary` | `3fc84ba8` | `tests\test_opt_compare_binary.bas` | rc=1：`VB2001 expected 'Text' or 'Binary'` + `VB2003` | rc=0，真跑 `OCB=False,3`（真 VB6 S12 同读数 `eq=False len=3`） |
| 3 | 定长字符串 `String * N`（长度表达式降到算术级） | `2d738304` | `tests\test_fixed_string_const.bas` | rc=1：`VB2001 expected '=' in Const declaration` + `VB2004` | rc=0，真跑 `FSC=3,pdm`（真 VB6 S13 同读数 `len=3 val=pdm`） |
| 4 | `ReDim` 多数组 + 无括号成员链 | `99d94762` | `tests\test_redim_shapes.bas` | 多数组第二项不认；成员链在 `.lMask` 处断 ⇒ `VB2003` | rc=0，真跑 `RS=7,9,42` |

`pdXCF.cls` 在 dev 上现在的状态（本地真实副本，134853 B）：

- `--syntax-only`：**rc=0**、0 条诊断；
- `--emit-c`：**rc=0**、0 条 `error`、342 行 **0 条**诊断（单文件口径仍有 154 条跨模块 `VB3001` 噪声，属既有的无工程上下文噪声）；
- ⇒ **可解析**，§10.7.6 第 1 项的分支漂移已消除。

`dev..probe` 的 `src/parser`、`src/lexer` 差异现在只剩「probe 基于旧 dev」造成的历史噪声（probe 缺 dev 近期的 Math / Parent / R2R3 / R4 等改动），**4 个补口的内容已全部落在 dev**。

真 VB6 本轮新增两条 oracle：S12（`Option Compare Binary` 合法，Binary 下 `"abc" = "ABC"` 为 False）、S13（`Public Const X As String * 3 = "pdm"` 合法，`Len=3`、值 `pdm`，即 PhotoDemon `PublicConstants.bas:28` 的真实写法）。

回归：`-Category syntax` → **PASS=187 FAIL=0 SKIP=0**。

未动（排队）：`VBA.Int` 残差（下一轮第一件）、S7 定长数组 `ReDim` 过宽、未声明名的 warning/error 等级差异。

## 11. 步数截断与固化（2026-11-09）

### 11.1 截断点

上一轮在修 `scripts/check_iface_view_authority.ps1` 的产品/架构类红时撞步数上限，重构只落到源码、**没跑构建、也没改哨兵脚本**。按用户当轮的固化口径（「改了但没跑验证 / 改了一半缺配套 ⇒ `git checkout --` 回退，不要留」）整刀回退。

### 11.2 回退了什么（7 个文件，全部 `git checkout --`）

| 文件 | 截断时的半成品内容 |
| --- | --- |
| `src/backend/module/cgen_iface_vtbl.cpp` | 新增 `ivrefViewDeclType` / `ivrefViewIfaceOfReceiver`；`ivrefCType` 改为委托；`ivrefViewGetterReadExpr` 改问新入口 |
| `src/backend/detail/util/cgen_helpers.inc` | 上述三枚的声明 |
| `src/backend/cgen_util_com.cpp` | 新增 `viewGetterReadExpr`（legacy → 新式二选一）；站S1 改调它 |
| `src/backend/stmt/cgen_call.cpp` | Debug.Print 站改调 `viewGetterReadExpr`；站D 改问 `ivrefViewIfaceOfReceiver` |
| `src/backend/stmt/cgen_control.cpp` | For Each 站改问 `ivrefViewIfaceOfReceiver` |
| `src/backend/stmt/cgen_with.cpp` | 字面量 `"vb6_ivref_"` 换成 `ivrefViewDeclType`；接收者问句换 owner |
| `src/backend/detail/expr/cgen_expr_call_com_bind.inc` | P6.4 改问 `ivrefViewIfaceOfReceiver` |

回退后 `git status` 只剩未跟踪的 `.zcodeignore`（非本任务产物，未动）。**没有代码进入提交**；HEAD 仍是 `d680b4d2`。

为什么判为半截：源码侧的四处站点已接上，但**哨兵脚本的钉没改**（⑤ 的 `emitIvrefRefSet(target` 会从 4 变 5、⑨ 需要补两枚新入口的钉），且**一次编译都没跑** ⇒ 既不能证明编译通过，也不能证明 6 个 R4 夹具产物逐字节不变。留着它就是「红着且不知红在哪」。

### 11.3 重做时照这个方案（省得重新推演）

1. `cgen_with.cpp` 的 `"vb6_ivref_" + cIdent(...) + "*"` 不再自己拼：走 `ivrefViewDeclType(<名册答出的接口名>)`（该接口可能空体、`ivreg_` 里 slots 为空，`ivrefCType` 会交空串，所以不能直接用它）。
2. 新增 `ivrefViewIfaceOfReceiver(receiverText)`（归一问 `slotIvrefViewName`）作为**唯一**入口；`cgen_call.cpp` / `cgen_control.cpp` / `cgen_with.cpp` / `cgen_expr_call_com_bind.inc` 四处改调它，`slotIvrefViewName(` 回到 3 处 / 2 文件。
3. 新增 `viewGetterReadExpr`（先 legacy、空则新式）住在 `cgen_util_com.cpp`，站S1 与 Debug.Print 站各只调一次；`ivrefViewGetterReadExpr` 保持 1 定义 / 1 声明 / 1 调用。
4. 哨兵：⑤ 口径扩到 **5 处 / 3 文件**（加 `cgen_control.cpp` 的 For Each 站，理由 = 同一发「退旧 + AddRef」不许站点内联）；⑨ 补 `ivrefViewIfaceOfReceiver`（1/1/4）与 `viewGetterReadExpr`（1/1/2）的钉。
5. 验收顺序：`check_iface_view_authority.ps1` OK → 6 个 R4 夹具 emit 与截断前基线逐字节一致 → `-Category syntax` 全绿 → 提交。

### 11.4 已确认的中间结论（未随回退丢失）

sha256 二分归因**已全部完成，6 条全部确认（未确认 0 条）**，结论在 §9.1 / §9.2：`fa56b9e2`（R2/R3 前）与 `a7fc9137`（R2+R3）两站 6 条形状全 SAME；`fdbd29e5`（R4）/ `91de0a68` / HEAD 全 DIFF ⇒ **唯一引入点 = R4**；其中 5 条为真回归（已修并验证），`k165` 仅为 canonical 形状随动。归因数据与修复读数已在 §9.4 固化，无需重跑。

### 11.5 仍开着的红（口径不变）

- `check_iface_view_authority`：**产品/架构类**，必须修（即 §11.3）；上一轮已回退，今天它是干净的（无未提交改动）——CI 上的红来自 `e21b5b5c` 那版已提交的 R4 站点，重做时按 §11.3 收口。
- `manifest_coverage` / emit manifest 等记账型：按用户授权继续挂账，不逐个 re-bless。

## 12. 对账（2026-11-09 晚，HEAD `cb1a68f1`）

### 12.1 五条真回归 → 全部落在 `43528bb8`

`fix(cgen): 修 R4 canonical 视图的真回归 —— 形参/For Each/语句位/With 接收者`（13 文件，+258/−53）。逐条：

| 夹具 | 修在哪 | 形状 |
| --- | --- | --- |
| `k172_viewarg` | `cgen_decl_func.cpp` / `cgen_decl_prop.cpp`（形参归一后同登 `knownIvrefVars_`）；`cgen_util_classtype.cpp` 新增 `ivrefViewBoxExpr`；`cgen_expr_call_arg_emit.inc` | 薄指针按 legacy 发 `.vtbl` ⇒ cl C2039；ByRef Variant 实参把 `vb6_ivref_I**` 当 `vb6_VARIANT*` ⇒ `boxtype=46832`。修后形参走 canonical，ByRef 站发临时 `vb6_VARIANT` |
| `k174_vieweach` | `cgen_control.cpp` For Each canonical 分支；`cgen_iface_vtbl.cpp` 新增 `ivrefIfaceViewFromFn` / `emitIvrefIfaceViewFrom`；`cgen_base_generate_body_pass.inc` | For Each 落 `vb6_VariantFromStackVARIANT` ⇒ C2440。修后解包对象→取类名→`vb6_ivref_<I>_from_class` 派发 |
| `k175_viewstmt` | `cgen_call.cpp` 站D | 语句位零实参调用只认 legacy 视图 ⇒ 落 COM 兜底、读数错。修后 `slotIvrefViewName` → `ivrefVtableCall` |
| `k177_viewwith` | `cgen_expr_call_com_bind.inc` P6.4；`cgen_with.cpp` | P6.4 只认 legacy；With 接收者未登新式名册 ⇒ 静默走 COM 兜底。修后二选一 + 接收者判定/登记 |
| `k178_viewprop` | `cgen_call.cpp` Debug.Print 实参站；`cgen_with.cpp` | `Debug.Print x.Label` 走 COM getter ⇒ `rc=2/rn=3`。修后补 `ivrefViewGetterReadExpr` |

`k165_iface_slot` 不动发码，只随 canonical 更新 `tests/run_tests.ps1` 的形状断言（`mi=16` 不变）。HEAD 工作树已 grep 确认各修点在位（`ivrefViewBoxExpr` / `ivrefIfaceViewFromFn` / `emitIvrefIfaceViewFrom` / For Each 分支 / 站D / P6.4），代码与文档一致，不存在「只写在文档里」。

修后读数见 §9.4（k172 `boxtype=0`、k174 `sum=303`、k175 `alpha=2/beta=10`、k177 `t1=6/t2=15`、k178 `rc=4/rn=2`、k165 `mi=16`）。

### 12.2 形状门实测（本机，HEAD 二进制）

`scripts/emit_manifest.ps1`（默认 x64 口径，与 CI 一致）：

```
期望 610 行 / 实得 617 行 / 相同 585 / 哈希不同 25 / 缺席 0 / 多出 7
```

**不是「只剩 k165」**，25 的构成可逐条解释：

- 6 行 = 六个 R4 夹具（5 条已修真回归 + k165 canonical 随动）；
- 18 行 = 含 canonical 接口的工程，因新增 `from_class` 派发器多一段函数；
- 1 行 = `tests/test_intrinsic_consts.bas`（本轮常量表收口时追加了 `IC-CF-FILES` / `IC-ASYNC-END` 两条断言）；
- 7 行多出 = 本轮新增夹具未登记（`pd_parentctl` ×2、`test_builtin_const_ext`、`test_procscope_loopdim`、`test_randomize_stmt`、`test_vbamath`、`test_vbamath_bad`）。

CI（run `42617564259`，HEAD `cb1a68f1`）的 `Emit manifest (shape oracle)` **红**，红点即上表：`存量 21 份发码真的变了 + 新增 7 份未登记`。CI 少 4 行的原因是**环境**，不是版本差：

- 差集恰为 `guisample/GUISample.vbp`、`guisample/vb6ref/src/vc6|vs2015/GUISample.vbp`、`TeeChartChartDisplay/DisplayProject.vbp`；
- 四者的 `expected == CI`，本机值更大（如 GUISample 306612 vs 277667 字节）；
- 它们都 `Object=` 了本机注册的 OCX（Codejock / TeeChart5 / RICHTX32），CI runner 没有 ⇒ COM 绑定深度不同；
- 本机 19:41 的 `r4_manifest.txt` 与 20:38 的 `recon_manifest.txt` 这四行逐字节相同 ⇒ **早于本分支全部改动，非本轮引入**（属 `emit_manifest.ps1` 头注释里已记录的「语料读取机器状态」旧缺陷类）。

### 12.3 结论

未发现未修的真回归；25 行 DIFF 全部可归因、7 行多出均为新夹具。形状门仍按用户口径挂账，待统一 re-bless（`scripts/rebless_emit_manifest.ps1 -Bless`，或取 CI `emit-manifest` artifact）。

## 13. CI 触发规则 + 内建常量 oracle 复核（2026-11-09）

### 13.1 非 src 改动不触发 GA（`.github/workflows/ci.yml` 的 `on.paths` 白名单）

22 个 job 只证明「编译器没坏 / 判据还能判别」，与文档排版无关。自本日起 `ci.yml` 的 `push` / `pull_request`（均限 `dev`）加 `paths` 白名单，未命中的提交 **不触发**、不产生 run：

| 白名单条目 | 为什么必须跑 |
| --- | --- |
| `src/**` | 编译器本体；实质改动一律要跑 |
| `tests/**` | 改判据 / 夹具 / 像素门会改变**验证能力本身**，必须证明改完仍能红能绿 |
| `scripts/**` | 形状门 / 哨兵 / 测试驱动，同样属于验证能力 |
| `cmake/**`、`CMakeLists.txt`、`build.bat`、`build.sh` | 构建口径变了，产物可能跟着变 |
| `.github/workflows/**` | 工作流自身改动；跑一次才知道新配置是否真能起 |
| `emit-manifest.expected.txt` | 形状门的期望基线 |

配套约定（写在这里就是为了防止「没跑 CI = 绿」的误读）：

1. **「没跑 CI」≠「绿」**。skip 只表示该提交按规则不进入验证，**不**表示已验证通过。release / 对账 / 合并前必须有一次真实 run 的绿，或显式补跑。
2. `src/**` 内纯注释 / 纯格式 / 零语义改动：提交信息带 **`[skip ci]`**（GitHub Actions 原生识别，`[ci skip]` / `[no ci]` 同义）。带了这个标记就**没有**验证，别拿它当已验证。
3. `tests/**` 内的纯注释/排版同理可带 `[skip ci]`；但只要动了判据、夹具、像素门，就**不许**带。
4. 需要强制跑一次：`gh workflow run ci.yml --ref dev`（`workflow_dispatch` 不受 `paths` 限制）。
5. `ai/**`、`docs/**`、`www/**`、`pic/**` 等非源码目录不在白名单 ⇒ 纯文档/交接提交天然 skip。

另外两个 workflow 不受影响：`ci_t0.yml`（tag `github-test-*` / 夜间 cron / dispatch）与 `ci_vbman.yml`（`workflow_run` 联动 ci_t0 / tag / dispatch）都不监听 push，无需改。

**本条规则自身的落地提交带 `[skip ci]`**：改动只有 `ci.yml` 的 `on:` 过滤与本文档，22 个 job 的行为未变，按新规则不跑。

### 13.2 内建常量表（38 条中的 8 条）：用真 VB6 取 oracle 复核

上一轮已在 `2e19c556` 补表并收口（`vbCFFiles` 7 处 + `vbAsyncStatusCodeEndDownloadData` 1 处 = 8 条 VB3001）。本轮按要求补做**值本身的 oracle 证明**——不是从文档抄数，而是拿真 VB6 编译探针 exe，把常量的运行时值打印出来当基准：

- 编译器：`C:\Program Files (x86)\VB6Mini\bin\VB6.EXE`，版本 **6.00.9782**，`VB6.OLB` 296466 字节。
- 探针工程（中间产物，不入库）：`.build\oracle\cf\Probe.vbp`（Clipboard 全族）、`.build\oracle\async\Probe2.vbp`（`StatusCode` 单枚）。两者分开是因为一枚常量若不被识别会让整编失败，混在一起会连累另一枚的证据。
- 编译/运行：`VB6.EXE /make Probe.vbp /outdir <temp>` → `Probe.exe`（16384 字节）→ 运行写 `probe_out.txt`。
- **注意**：`/out <单文件路径>` 会报「权限被拒绝」，必须用 `/outdir <目录>`；且工程要复制到 `%TEMP%` 下再编，在仓库 `.build` 下直接编也可能失败。

`Probe.exe` 的实际输出（本轮重跑，逐字）：

```
vbCFFiles=15
vbCFText=1
vbCFBitmap=2
vbCFMetafile=3
vbCFDIB=8
vbCFPalette=9
vbCFEMetafile=14
vbCFRTF=-16639
```

```
vbAsyncStatusCodeEndDownloadData=6
```

与 [builtin_consts_ext.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/builtin/builtin_consts_ext.inc#L194-L204) 表内值逐项比对：

| 常量 | 表内值 | VB6 oracle | 结论 |
| --- | --- | --- | --- |
| `vbCFFiles` | 15 | 15 | 一致（**8 条缺口的 7 条**） |
| `vbCFText` | 1 | 1 | 一致 |
| `vbCFBitmap` | 2 | 2 | 一致 |
| `vbCFMetafile` | 3 | 3 | 一致（本轮补测，此前未取值） |
| `vbCFDIB` | 8 | 8 | 一致 |
| `vbCFPalette` | 9 | 9 | 一致（本轮补测，此前未取值） |
| `vbCFEMetafile` | 14 | 14 | 一致 |
| `vbCFRTF` | −16639 | −16639 | 一致 |
| `vbAsyncStatusCodeEndDownloadData` | 6 | 6 | 一致（**8 条缺口的 1 条**） |

- **无法验证项：无**。9 枚全部有真 VB6 输出背书，本轮未改动任何数值（表中值本来就对）。
- 修前/修后计数（已在 `2e19c556` 记账）：夹具 `tests\test_builtin_const_ext.bas` 在旧二进制上 2 条 `VB3001`（两枚常量各 1）→ 新二进制 0 条；PhotoDemon 探针口径 VB3001 8 处 → 0 处（该提交后探针 VB3001 总账 3693→0 的最后一截）。
- 断言侧：`run_tests.ps1` 的 `test_builtin_const_ext`（Compile+Run，读数 `BCE-CF=15` / `BCE-ASYNC=6`）+ `bce_emitc_reads_are_literals`（发码必须是字面量）+ `test_intrinsic_consts` 的 `IC-CF-FILES=15` / `IC-ASYNC-END=6`。这些断言证明的是**值本身**，不是「能编译」。
- 本轮**无新提交**：常量项已在 `2e19c556` 完成，本节的产出只是 oracle 复核记录。

### 14. Randomize + Rnd：与真 VB6 逐值对齐

`Randomize` 语句形态这一项，动到 `Rnd` 算法才发现**只修语句形态等于假修**（用户预告的第 3 条）。本轮按「`Randomize` + `Rnd` 序列对齐」收口。

#### 14.1 真 VB6 的算法（反汇编 + 探针双证）

oracle 不是文档抄的，是两条独立证据：

1. **反汇编** `C:\Windows\SysWOW64\MSVBVM60.DLL`（`dumpbin /disasm:nobytes`，image base `0x66000000`）：
   - `rtcRandomize` @ RVA `0x486B1`；有参分支 → `0x6608C074`（`__vbaR8Var` 转 double，`fstp qword`）→ `0x660EDF4A`；缺参分支 → `0x660486E0`（`call Timer` → **`fstp dword`**，即先落 Single）→ 同一个 `0x660EDF4A`。
   - `0x660EDF4A` 折叠：`u = 参数高 32 位`；`fold = ((u & 0xFFFF) << 8) ^ ((u >> 8) & 0xFFFF00)`；`state = fold | (old_state & 0xFF0000FF)`。
   - `rtcRandomNext` @ RVA `0x4871A`：`Rnd` 形参是 **Single**（`[ebp+8]` dword）；`==0` 直接返回当前 state；`<0` → `0x6608C089`：`state = ((u >> 24) + u) & 0xFFFFFF`（`u` 是 Single 位型）再走 LCG；`>0`/缺参直接走 LCG。LCG 写作 `eax = 0xFFC39EC3 - state*0x2BC03`，等价 `state*1140671485 + 12820163`（mod 2^24）。
2. **真 VB6 探针**：`C:\Program Files (x86)\VB6Mini\bin\VB6.EXE`（6.00.9782）。中间产物 `.build\oracle\rnd\RndSeq.vbp` / `RndSeq.bas`（不入库）。用 `Hex$(CLng(Rnd() * 16777216#))` **精确还原 24-bit state**（`Rnd` 是 Single，24 位尾数在 double 里精确）。

由此确定的语义（三条最容易踩的）：

- 有参 `Randomize n`：取 `(double)n` 的**高 32 位**折半 XOR，且 **`state` 的低字节沿用旧状态**（`old & 0xFF0000FF`）。所以**连投同一整数种子不复现序列**，旧测试的 `a = b` 断言本身就是错的。
- `Randomize 0` 是**确定性**的 `fold = 0`，不是 Timer。
- 裸 `Randomize` 用 `Timer()` 的 **Single 位型**（与有参不同，有参走 double 高 32 位）。
- `Rnd(<0)` 用 Single 位型重播种（不是 `seed & 0xFFFFFF`）；`Rnd(0)` 不推进。

#### 14.2 改了什么

| 文件 | 改动 |
| --- | --- |
| [vb6rtl_conv.c](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/rtl/core/vb6rtl/vb6rtl_conv.c#L100-L175) | `vb6_Rnd(int32_t)` → `vb6_Rnd(double)`；新增 `vb6_rndFoldState` / `vb6_rndStep` / `vb6_RandomizeTimer`；`vb6_Randomize` 改折半 XOR 播种（旧实现是 `(int32_t)seed & 0xFFFFFF` 截断 + `seed==0` 当 Timer + 裸调用用 `GetTickCount`，三处都错）；`vb6_Rnd_Full` 转发 |
| [vb6rtl_builtin.h](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/rtl/core/vb6rtl/vb6rtl_builtin.h#L99-L101) | 同步签名；新增 `void vb6_RandomizeTimer(void);` |
| [cgen_expr_call_builtin_fixup.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_call_builtin_fixup.inc#L407-L415) | 裸 `Randomize` 不再补 `"0.0"`，改发 `vb6_RandomizeTimer()` |
| [cgen_expr_call_com_bind.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_call_com_bind.inc#L1019-L1032) | `Rnd` 的 ident 展开是 `vb6_Rnd(1)`，括号里的 `1` **不是接收者**；`Rnd(x)` 必须用 x 替代它（旧发码 `vb6_Rnd(1, x)`） |
| [cgen_expr_call_arg_emit.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_call_arg_emit.inc#L116-L124) | `Rnd`/`Randomize` 的 Variant 实参显式 `vb6_VariantToDouble(...)`（旧发码把 `vb6_VARIANT` 结构体直接塞进 double 形参） |
| [cgen_call.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/stmt/cgen_call.cpp#L650-L656) | 语句路径同一规则（防御性：`vb6_RandomizeTimer`） |

#### 14.3 oracle 取值过程

- 探针工程 `RndSeq.vbp` 复制到 `%TEMP%` 下 → `VB6.EXE /make RndSeq.vbp /out <log>`（`/outdir` 也行）。两个坑：模块名不能与工程名（VBP 文件名）相同，否则报「名称与已存在的模块、工程或对象冲突」；`Debug.Print` 在 EXE 里不进 stdout，探针必须 `Print #f` 写文件。
- 探针覆盖 29 段：BASELINE（未播种）、`Randomize 0/1/2/5/42/100`、`42.5/42.25/0.5`、`-1/-2/-1.5`、Variant 42、`16777216/16777217`、变量种子（Single/Double/Integer/Long/Variant）、`Rnd(-1)` ×2、`Rnd(0)`、`Rnd(1)`、裸 `Rnd`、`Randomize 5` ×2（A/B 不复现）、裸 `Randomize`（只验值域）。
- 每段打 6 个值，共 29 行读数，作为 `run_tests.ps1` 的期望串逐行嵌入。

#### 14.4 修前 / 修后

- **修前红**（旧 C3.exe 编新夹具）：
  - 2 条 **error C2440**：`vb6_Randomize(vv)` 把 `vb6_VARIANT` 塞进 double（两处 Variant 种子）⇒ 链接不出 exe。
  - 4 条 **warning C4020** `vb6_Rnd: 实参太多`：`vb6_Rnd(1, (-1))` / `vb6_Rnd(1, 0)` / `vb6_Rnd(1, 1)`。C 里多传实参只是警告，**第一个实参 `1` 被当真种子** ⇒ `Rnd(0)` 变成取下一个、`Rnd(-1)` 根本不重播种 —— 能编、数是错的，正是「假修」的形状。
  - 裸 `Randomize` 发 `vb6_Randomize(0.0)`；`Randomize 0` 在 RTL 里被当 Timer。
- **修后绿**：同一夹具真编真跑，输出与 VB6 oracle **29/29 行逐字相同**（`Compare-Object` 无差异）。`run_tests.ps1` 的 30 条期望串全部命中；发码形状门 5 条在场、3 条不在场（`vb6_Rnd(1, ` / `vb6_Randomize(0.0)` / `(vb6_Randomize - 1)`）。
- 裸 `Randomize` 只能验 `RZ-BARE-RANGE=True`（Timer 播种天然不可复现），已在夹具里注明。

## 16. `VBA.Int` 根因 + 修（账 #478）

### 16.1 根因：限定内建名被「cIdent 转义 + 再拼前缀」拼两次

`VBA.Int(1.7)` → `vb6_vb6_Int` → `LNK2019`；同一处的 `VBA.Fix(1.7)` → `vb6_Fix` 正常。差别不在白名单缺项，而在 [cgen_expr_member_precheck.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_member_precheck.inc) 的 `_objLower == "vba"` 折叠分支：

```cpp
lastExpr_ = "vb6_" + cIdent(fnCanon);   // 修前
```

`Int` 是 **C 关键字**（[cgen_base_naming.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/cgen_base_naming.cpp) 的 cKeywords 含 `int`），`cIdent("Int")` 先返回 `vb6_Int`，这里再拼一次 `vb6_` ⇒ `vb6_vb6_Int`。`Fix` 不是关键字，`cIdent("Fix")` 原样返回 `"Fix"`，拼出来恰好等于 RTL 的 `vb6_Fix` —— **纯属巧合**。所以不是「白名单少了 Int」，而是这一格用错命名函数（该问内置表，不该现拼）。

### 16.2 真 VB6 四形态（6.00.9782 探针）

探针 `.temp/vbaint/vb6/{ProbeMain.bas,VbaIntProbe.vbp}`（CRLF + ASCII），`/make` 后跑 exe 写读数：

| 写法 | 真 VB6 |
| --- | --- |
| `VBA.Int(1.7)` / `VBA.Int(-1.7)` | 编过；`1` / `-2` |
| `VBA.Fix(1.7)` / `VBA.Fix(-1.7)` | 编过；`1` / `-1` |
| 裸 `Int(-1.7)` / `Fix(-1.7)` | 编过；`-2` / `-1` |
| `Math.Int` / `Math.Fix` | **不存在**（「未找到方法或数据成员」，C3 报 VB3045，负控已在 §10.4） |

**重要更正（用户口径）**：`Int(1.7)` 的真值是 **1，不是 2**。VB6 的 `Int` 是「向下取整（向 -∞）」，对正数 1.7 取整结果与 `Fix` 相同；`Int` 与 `Fix` 只在**负数**上分界（`-2` vs `-1`），`Int` 不是向上取整。断言据此钉：正例 `1`、负例 `-2`，分界必须落负例。

### 16.3 修法

`else if` 链里把「RTL 名」交给内置表，不再现拼；`rnd` 特殊形与零参 `()` 原样保留：

```cpp
} else if (auto itB86 = builtinFuncs.find(fnLower86); itB86 != builtinFuncs.end()) {
    lastExpr_ = itB86->second;
    if (vbaZeroArgFns.count(fnLower86)) lastExpr_ += "()";
} else if (vbaZeroArgFns.count(fnLower86)) { ... }
```

表里其余名字两种拼法逐字节相同（只有 `int` 撞 C 关键字），所以这一改只动 `Int` 一格。

### 16.4 修前红 / 修后绿

- **修前**（把该分支临时换回旧写法重建）：`--emit-c` 出现 `vb6_vb6_Int`；真编 `error LNK2019: 无法解析的外部符号 vb6_vb6_Int，函数 vb6_Main 中引用了该符号`，exit=1。
- **修后**：`--emit-c` 出 `vb6_Int(1.7)` / `vb6_Int((-1.7))` / `vb6_Fix((-1.7))`，无 `vb6_vb6_`；真编 exit=0，运行读数 `VI-POS=1 VI-NEG=-2 VF-POS=1 VF-NEG=-1 BI-NEG=-2 BF-NEG=-1 VA-ABS=3 VA-PASS`，与 16.2 的真 VB6 逐行相同。
- 夹具 [test_vbaint_qual.bas](file:///c:/Users/Administrator/Documents/c3.vb6.pro/tests/test_vbaint_qual.bas) + 三条 cell（`test_vbaint_qual` 真跑 / `vbaint_emitc_rtl_names` 形状 / `vbaint_no_double_prefix` 缺席）。

## 17. 冻结与固化（2026-11-10）：k489/k2947 的 canonical 四形收口

### 17.1 固化前工作区（原样）

`git status --short`：

```
 M src/backend/cgen_util_classtype.cpp
 M src/backend/decl/cgen_decl_prop.cpp
 M src/backend/detail/expr/cgen_expr_member_class_fallback.inc
 M src/backend/detail/expr/cgen_expr_member_class_module.inc
 M src/backend/detail/expr/cgen_expr_member_voidptr_com.inc
 M src/backend/detail/stmt/cgen_assign_prop_write.inc
 M src/backend/detail/stmt/cgen_assign_stmt_special.inc
 M src/backend/detail/stmt/cgen_setlet_set_prop.inc
 M src/backend/detail/util/cgen_helpers.inc
 M src/backend/expr/cgen_expr_member.cpp
 M src/backend/expr/cgen_expr_with.cpp
 M src/backend/module/cgen_iface_vtbl.cpp
 M src/backend/stmt/cgen_with.cpp
 M tests/run_tests.ps1
?? .zcodeignore
?? Form1.frm
```

`git diff --stat`：14 files changed, 377 insertions(+), 55 deletions(-)。

- `Form1.frm` 是 **0 字节**的编译残留（2026-11-09 21:28），已删除。
- `.zcodeignore` 是 2026-10-31 的本地工具文件，早于本刀、与本刀无关，保留未跟踪、不入库。

### 17.2 逐文件判定

| 文件 | 判定 | 依据 |
| --- | --- | --- |
| cgen_expr_member.cpp | **完整** | 入口拦「视图接收者的零参读」是 j/k 的根治刀；k2947 真编真跑 16/16 |
| cgen_util_classtype.cpp | **完整** | `ifaceViewOfExpr` 补 IndexOrCall(callee=MA) + WithMemberExpr 两形；k 行依赖后者 |
| cgen_decl_prop.cpp | **完整** | 带参桩转发 `ivForwardArgs`；k2947-i `get_item(x, 3)` 依赖 |
| cgen_with.cpp | **完整** | `With MakeI()` / `With h.OwnView` 的薄指针临时量；k2947-m/n |
| cgen_expr_with.cpp | **完整** | With 内 `.Owner` / `.Has` canonical 读写；k2947-m/n/w0 |
| cgen_assign_prop_write.inc | **完整** | Let 与 With Let 的新式二臂；k489 |
| cgen_assign_stmt_special.inc | **完整** | 带下标属性写 `x.Item(3)=v` + `knownIvrefVars_` 认作变量；k2947-i / k489 |
| cgen_setlet_set_prop.inc | **完整** | Set / Set Nothing / With Set 三站新式臂；k2947-o/b/p |
| cgen_helpers.inc / cgen_iface_vtbl.cpp | **完整** | `ivrefViewSetterCallExpr` 声明与实现（写档专用，不复用读档键序） |
| tests/run_tests.ps1 | **完整** | k489 + k2947 两套 oracle 换 canonical；16+9 正针、15+6 负针机检 0 缺 0 多 |

**整刀回退（半截/死分支）**：

| 文件 | 为什么退回 |
| --- | --- |
| cgen_expr_member_class_fallback.inc | 只有一条 `DBG2947-FB` 调试打印，该分支实测从未命中 |
| cgen_expr_member_class_module.inc | ivRead210 扩读档是**试改**：`DBG2947` 打印证明它被更早的 `memSym` 闸挡住、从未执行；根治改在入口拦，此扩属死码 |
| cgen_expr_member_voidptr_com.inc | 同上：`DBG2947-VP` 一条未出现，承载 j 的路径不在该分支；保留只会误导下一位 |

调试插桩（`DBG2947` / `-S` / `b` / `c` / `-FB` / `-VP`）已从保留的两个文件里删净。

### 17.3 根因与修法（k2947-j / k2947-k）

- **现象**：`h.OwnView.Has`（无括号）与 `With Me: .OwnView.Has` 运行值恒 0；带括号的 `h.OwnView.Has()` 正常。
- **根因**：接收者是薄指针视图时，这两形落进「类实例成员」那站，按 `resolveClassMemberCall(接口名, 成员)` 发 `vb6_IK2947_Has((void*)<视图>)`。接口模块函数的首形参契约是**实例**，`vb6_iv_thin_of_*` 每枚实现者都试不中 ⇒ 恒 0。
- **修法**：在 [cgen_expr_member.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/expr/cgen_expr_member.cpp) 的 visit 入口先拦：`!asCallCallee_ && !suppressVirtDispatch_` + `ifaceViewOfExpr(接收者)` 非空 + 槽是**读**档（裸名或 `get_`）且零参非 void ⇒ 接收者求值一次（裸标识符直接复用文本；复合表达式落 `__ivrecvN`，因为视图属性自带 AddRef）⇒ 形状仍交唯一作者 `ivrefVtableCall`。`ifaceViewOfExpr` 补 `WithMemberExpr` 一臂（`With Me` 的基类在 With 栈上）。
- **修前**：`K2947-j=0 K2947-k=0`（其余 14 行正确）；**修后**：`K2947-j=1 K2947-k=1`，16/16 与 oracle 逐字一致。
- **回归冒烟**：k489 / k165 / k172 / k174 / k175 / k177 / k178 全部真编真跑通过；k489 的 `vb6_BSTR_Assign(&s, x->vt->get_label(x));` 与 `n = x->vt->get_num(x);` 两针特意保住（裸标识符接收者不落临时量）。

### 17.4 CI 主链路 22 job（run `42651946234`，head `99d94762`）

| # | job | 状态 | 分类 |
| --- | --- | --- | --- |
| 1 | Build C3.exe | success | — |
| 2 | Baseline (M1) | success | — |
| 3 | Emit manifest (shape oracle) | failure | **记账**（10 份新输入未登记 + 21 份存量哈希变化） |
| 4 | Tests (bas #1) | failure | 产品/夹具（含 k2947 旧 oracle 与 j/k 值） |
| 5 | Tests (bas #2) | failure | 产品/夹具 |
| 6 | Tests (smoke) | success | — |
| 7 | Tests (syntax) | success | — |
| 8 | Tests (asm) | success | — |
| 9 | Tests (compile #1) | failure | **记账**（`[STATIC]`：builtin_opt_defaults / manifest_coverage / iface_view_authority） |
| 10 | Tests (compile #2) | failure | **记账**（`[STATIC]`：legacy_view_arg_owners / box_temp_reclaim / class_refcount_authority） |
| 11-20 | Tests (vbp #1 #2 #3 #4 #5 #6 #7 #8 #11 #12) | failure | 产品/夹具（k489/k166/iface_wrap/BalloonTooltips/iface_multi 一族） |
| 21-22 | Tests (vbp #9 #10) | success | — |

- **7 success / 15 failure**，失败集与 `c0d57af4` 逐字一致 ⇒ 上轮 5 个解析器/loopKind 提交**零新增红点**；`smoke` / `syntax` / `asm` / `Baseline` 全绿 ⇒ 最底层路径未带崩。
- **记账清单（不修、留账）**：`Emit manifest` 1 枚；`compile #1/#2` 共 6 条 `[STATIC]`。
- **尚未被任何 CI 验证的提交**：`d4a25a1c`（`[skip ci]`，仅文档）、`574f0029`、`30a099e3`、`3ebb656c`、以及本刀提交 —— run `42651946234` 的 head 停在 `99d94762`。本刀未手工触发 `workflow_dispatch`（按用户口径：文档/注释类不空跑；本刀是 src 改动，应由下一次 dev push 正常触发）。

## 18. `VB_Creatable` 交代：R4 保持不动（结论 b）

### 18.1 为什么会动它（如实）

不是 CI 先红，是我把 R4 之后的 vbp 红**误当成 R4 拆过宽**：

1. `d4a25a1c` 收 R4（`Fix 214`：去掉「显式 `VB_Creatable` 必须为 False」这条）；
2. 随后 vbp #1/#2/#3/#7/#11/#12 红，我没有先定位就回滚整刀（`574f0029`）；
3. 回滚后同一批 job **红得更狠**（缺 `vb6_ivref_IK174*` / `__iv_K165Iface` / `vb6_iv_thin_of_`），
   才查到 `43528bb8` 已经补齐「形参 / For Each / 语句位 / With 接收者」四形；
4. `3ebb656c` 撤销回滚，dev 回到 R4 + `43528bb8` 的正确组合。

截断前那次「恢复 `VB_Creatable` 校验」属于上面第 2 步的延续（选项 a），**没有留下任何 diff**
（工作区已确认 `src/driver/driver_interface.cpp` 无改动）。

### 18.2 拿真 VB6 说话（探针 + 读数）

探针：`.temp/vbcreatable/vb6/{Probe.vbp, IThing.cls, ThingImpl.cls, BasMain.bas}`，
VB6 `/make` 编译后运行，读数在 `readings.txt`：

```
IFACE_DISPATCH=42
IFACE_TAG=impl
CREATABLE_DEFAULT=0
CREATABLE_TAG=[]
```

- `IThing`：`MultiUse = -1`、`Attribute VB_Creatable = True`、`VB_Exposed = False`，
  `Ping() As Long` / `Tag() As String` **全空体**。
- `ThingImpl`：`Implements IThing`，`IThing_Ping = 42`、`IThing_Tag = "impl"`。
- `Dim i As IThing: Set i = t` ⇒ 派发到实现者（42 / "impl"）。
- `Set i = New IThing` ⇒ 合法，得到默认空体实例（0 / 空串）。

交叉证据（上一轮留下的 `.temp/procdim/vb6/S14_partialiface`）：同一枚 `VB_Creatable = True`
的空体 `IMul` 被 `Beta` 只实现 2/3 时，VB6 直接报
`类模块需要实现 'Three' 为接口 'IMul'` ⇒ VB6 确实把它当**接口契约**在逐成员校验。

**结论**：`VB_Creatable` 只管 `New X` 能不能建，与 `Implements X` 无关；VB6 允许一个
**可创建 + 全空体** 的类同时充当接口宿主。C3 的 R4 与该语义一致。

### 18.3 三条路的取舍

- **(a) 回退 R4：不采用。** 与 18.2 的 VB6 读数相反；且实测回滚会把 vbp #1/#2/#3/#7/#11/#12
  打红（`3ebb656c` 已撤销）。
- **(b) 保持 R4 不变、改用例：采用。** R4 之后剩下的红不是登记判据错：
  - k489 / k2947 的 oracle 钉的是 `.vtbl` / `.obj` legacy 形状（R4 之前的表示）⇒ **断言写错**，
    已在本刀换 canonical；
  - k2947 的 j/k 恒 0 是**真产品错**（视图接收者裸读被喂给实例形参）⇒ 已修（§17.3）。
- **(c) 收窄：不采用。** 探针表明 `VB_Creatable` 对 `Implements` 无任何影响，加前置条件没有
  VB6 依据；安全性已由第 1 条判据兜底（只有被 `Implements` 指到的名字才进表，
  `driver_interface.cpp` 的 `wanted` 集合），不需要再加闸。

### 18.4 本刀之后仍未完成（只留账，不在本轮开活）

- k166：两条发码 oracle 仍钉 legacy 形状。
- iface_wrap / BalloonTooltips：`Set <接口变量> = <Variant/Object>` 的拆箱与类名传递。
- iface_multi：夹具本身不合法（真 VB6 也拒绝，见 S14），应改成「部分实现 ⇒ VB3012」负控。
- 记账类：`Emit manifest` 1 枚 + `compile #1/#2` 的 6 条 `[STATIC]`。

## 19. CI 对账更正（以 `gh run list` 实际数据为准）

§17.4 那张表是 run `42651946234`（head `99d94762`）的，**不是最新**。补两条并更正：

| run | head | 状态 | 结果 |
| --- | --- | --- | --- |
| 42651946234 | 99d94762 | completed | 7 success / 15 failure（旧表） |
| 42660526229 | 30a099e3 | completed | **13 success / 9 failure** |
| 42669065145 | 650b7bac | **in flight** | Build / smoke 已 success，其余进行中 |

run `42660526229`（`30a099e3`）的 22 job：

- success（13）：Build C3.exe、Baseline (M1)、Emit? 否、Tests (smoke/syntax/asm)、bas #1、bas #2、
  vbp #4、#5、#6、#8、#9、#10。
- failure（9）：Emit manifest、compile #1、compile #2、vbp #1、#2、#3、#7、#11、#12。

**关键**：`30a099e3` 处在 `574f0029`（R4 回滚）**仍然生效**的窗口里，所以那一轮的 bas（含 k2947）
是绿的（legacy 路径下 j/k 本来就对），而 vbp #1/#2/#3/#7/#11/#12 红正是**回滚本身**造成的
缺符号（`vb6_ivref_IK174*` / `__iv_K165Iface` / `vb6_iv_thin_of_`）。

因此：

1. 「已被 CI 验证」的提交：`99d94762`、`30a099e3`；
2. **尚未被任何 CI 验证**：`d4a25a1c`、`574f0029`（回滚）、`3ebb656c`（撤销回滚）、
   `650b7bac`（本刀 src）、`e4f75a7e`（本刀文档）——其中 `650b7bac` 的 run `42669065145` 正在跑；
3. §17.4 里「未能验证 30a099e3」一句作废，以本节为准。

## 20. `Set <接口变量> = <Variant/Object>` 拆箱（§18.4 第一条欠项已清）

**站点**：`cgen_setlet_set_prop.inc` 的 ivref 目标分派 —— 原 `else` 末尾发
`VB4001: Interface binding: Set <x> = <expr> ...`。实测双击中：`iface_wrap`（`Set f = v`，
v Variant 往返 `Impl`）与 `BalloonTooltips`（`Set Subclass = vSubclass`，ByVal Variant 形参）。

**修法**（复用既有唯一主人，不另造视图）：
`vb6_VariantToObjectValChecked(RHS)` 取包装器 → `vb6_ComObject_ClassNameA` 认类名
（内部 `IsBoxed` 三闸，裸 HWND / 外部 IDispatch 安全）→ 名字非空才
`vb6_ComObject_GetInstance` → `vb6_ivref_<I>_from_class(inst, name)`
（`emitIvrefIfaceViewFrom`，与 For Each 视图站同一发）。裸标识符目标走
`emitIvrefRefSet`（新 AddRef / 旧 Release），`me->` 字段沿用 B04 纯赋值口径。
`As Object` 的裸标识符同档（`knownObjectVars_`），包装器指针直接用。
认不出 / 这枚实例没实现该接口 ⇒ **运行期 13 Type mismatch**（`vb6_RaiseError(13, ...)`），
整条 Set 被放弃、目标槽保持原值 —— 真值见 §21 的 VB6 探针（不是静默 Nothing）。

**生成码形状**（iface_wrap）：

```c
void* __ivbox1 = vb6_VariantToObjectValChecked(v);
const char* __ivname1 = vb6_ComObject_ClassNameA(__ivbox1);
{ /* vb6_ivref_IFoo*: 旧值 Release / 新值 AddRef */
  vb6_ivref_IFoo* __ivold1 = f;
  f = vb6_ivref_IFoo_from_class(__ivname1 ? vb6_ComObject_GetInstance(__ivbox1) : (void*)0, __ivname1);
  if (f) f->vt->AddRef(f);
  if (__ivold1) __ivold1->vt->Release(__ivold1);
}
```

**读数（修前 → 修后）**：

| 项 | 修前 | 修后 |
| --- | --- | --- |
| iface_wrap emit-c | exit 1，`VB4001 ... (tB Interface IFoo)` | exit 0，无 error |
| iface_wrap x64 真编 + 运行 | 无 exe | exit 0；`IFACE-CALL-OK` / `IFACE-DONE` |
| iface_wrap x86 真编 + 运行 | 无 exe | exit 0；同两行 |
| BalloonTooltips x86 真编 | exit 1，2 发 `VB4001 ... (tB Interface ISubclass)` | exit 0，`BalloonTooltips.exe` 555008 B |
| `-Category syntax` | PASS=187 FAIL=0 | PASS=187 FAIL=0 |

**BalloonTooltips 两条既存 warning（非本刀引入，记在册）**：
- `(1,1): warning VB4001: TypeLib reference path not found: ...\Windows\SysWOW64\stdole2.tlb`
- `cTT.cls(154,31): warning VB3003: If条件应为Boolean, 实际为 Object`

**范围**：新分支只落在原 `VB4001` 的 `else` 末尾、错误之前 ⇒ 此前能编过的形状一字未动。

## 21. 负控真值：`Set <接口变量> = <Variant/Object>` 认不出时是「运行期 13」

**探针**：`.temp/vb6probe_neg`（VB6Mini `C:\Program Files (x86)\VB6Mini\bin\VB6.EXE` `/make` 真编，
Sub Main 把 `Err.Number` 写 `neg.txt`；`IFoo` 可创建、`Impl` 实现、`Other` 不实现）：

| 用例 | VB6 真值 |
| --- | --- |
| A `As Object` ← `New Impl` | 0；`f.Ping()=42` |
| B `Variant` ← `New Impl` | 0 |
| C `Variant` ← `New Other`（未实现） | **13**，且 `f Is Nothing=False`（原值保留） |
| D `Variant` ← `CreateObject("Scripting.Dictionary")` | **13**，`f Is Nothing=False` |
| E `As Object` ← 外部 Dictionary | **13** |
| F `Set f = Nothing` | 0 |

结论：VB6 **从不静默给 Nothing**，一律 13 且放弃整条 Set、目标保持原值。C3 按此落地：
视图算完为空 → `vb6_RaiseError(13, vb6_BSTR_FromStr(L"Type mismatch"))`，不走赋值。
（`vb6_RaiseError` 在 `On Error Resume Next` 下会返回，所以赋值必须显式放进 `else`。）

**断言**（`tests/iface_wrap`，双架构，`run_tests.ps1` 9503-9512）：
`IFACE-CALL-OK`、`IFACE-NEG-KIND=13`、`IFACE-NEG-EXT=13`、`IFACE-NEG-KEEP=7`、`IFACE-DONE`。
`IFACE-NEG-KEEP=7` 钉住「13 之后旧值还在」（旧值 = 前面 `Set f = v` 装进去的那枚 `Impl`，
`Impl.IFoo_Bar()=7`；若被清成 Nothing 这里会先 91 再打印空/0）。

## 22. `iface_multi` 判据纠正：非法当合法 → 改负控 + 补 3/3 正向

真 VB6 对「`Implements IMul` 只写 2/3 成员」**编译期直接拒**（`类模块需要实现 'Three' 为接口 'IMul'`）。
C3 语义层早已按真值报 `VB3012`，但夹具仍按「能编过 + `B-THREE-ERR=438`」注册 ⇒ 4 个 vbp 分片常红。
本轮把判据改对，不是删用例：

- `tests/iface_multi`（Beta 少写 Three）→ **负控** `iface_multi_incomplete_negcontrol`：
  新增 `Test-CompileFails`，三条同时成立才绿 —— `exit != 0` / 输出含 `VB3012` / **无 exe**。
- `tests/iface_multi_full`（Beta 补 `IMul_Three=23`）→ **正向** `iface_multi_full` + `_x86`：
  `A-ONE/TWO/THREE=11/12/13`、`B-…=21/22/23`、`MULTI-DONE`；新增读数 `A-TNAME=Object` /
  `B-TNAME=Object`（TypeName 对接口视图仍答 `Object` 而非类名 —— 在册欠项，不是本刀引入）。
- k166 两条 oracle 同步改到 `iface_multi_full` 并按**当前**发码面改写（IMul 已走 tB 薄视图）：
  正面针 = `VB6_IVTBL_IMUL` / `__stdcall *one` / `*three` / 每类一张 `vb6_ivtbl_IMul_for_*` /
  工程级 `vb6_ivref_IMul_from_class` / 两条 `_stricmp` 分支；反面针不变（legacy 三发绝迹）。
  旧的 `vb6_iface_IMul_unimpl_Three_Beta` + `RaiseError(438)` 两针随非法夹具退场（该代码路径已不可达）。

## 23. `check_rtl_proto_arity` 1438 → 1454 的账（记录作废，不是"口径更正"）

- 脚本自 `1e168609`（2026-10-10）**未再改动** ⇒ 统计口径没变。
- 计数只是 `src/rtl` 的纯函数；逐树实测：
  - dev `d7410cc3` = **1454**
  - 刀前 `.temp/preknife` (`3ebb656c`) = **1454**
  - `fbc26f64`（11-09 21:20）= **1454**
  - `c69cbbba`（11-09 更早）= **1453**
  - `9be0a411`（11-09 UC 批前）= **1453**
- 结论：**1438 在任何可复现的树上都测不出来**；它早于 11-09 的 RTL 批次，属旧树读数。
  按用户口径：**原记录作废，原因未明**（不是我们之前记错口径）。当前真值 = 1454 对，exit 0。

## 24. 2026-11-10 重跑普查：**emit 在 2 GB 红线处被打断**（装配阶段失控，先于语义/发码）

> **【勘误 / 本节结论作废，见 §25】** 本节把"2 GB 处被打断"判成"回归"是错的。A/B 实测：
> `ba455c7f`（本机 Release）与 HEAD 的逐模块内存斜率完全相同（22.4 vs 22.5 MB/模块），
> 两版都在 5–6 s 撞 2 GB。这不是 ba455c7f..HEAD 的回归，是自 `ba50be1a`（2026-06-26）
> 就在的 O(模块数 × 类型库符号数) 结构问题。§25 有读数与根因。


口径：HEAD `2a2c3d29`，`.build\C3.exe`（03:27 构建），单进程串行，`PhotoDemon.vbp --emit-c --arch x86`，
看门狗 2 GB / 600 s。空闲内存 7.4–7.5 GB / 16.1 GB。

| 项 | 读数 |
| --- | --- |
| EXIT | **KILLED（PEAK_OVER_2048MB）** |
| 耗时 | 36.9 s（三次复跑 36.8 / 36.9 / 36.8） |
| 峰值工作集 | **2069.8 MB**（2 GB 上限触发即杀） |
| stdout | 0 B（一个模块的 C 都还没 dump） |
| stderr | 142 行，全部是 `C3: CoClass/Interface ... identity`（Pass D/E 的信息行），0 条 `error` |
| 中间目录 | 未创建任何新 `%TEMP%\C3C\<id>` ⇒ 根本没进 codegen |

内存曲线（`C3_TIMING=1` + 250 ms 采样，`.temp\pd_census2\memtrace.txt`）：

- `0…26 s`：线性 +22 MB/s（36 → 558 MB），对应 `parse 25955 ms`（596 份源文件）。
- `26 s`：142 行 CoClass/Interface 身份行打完（Pass E 结束），随后 **+144 MB/s**：
  `558 MB(t=26.4) → 772(28.5) → 1088(30.6) → 1455(32.7) → 1773(34.8) → 2070(36.9)`。
- `C3_TIMING` 只打出 `vbp+expand / lex / preprocess / parse / typelib+prepasses`，
  **`semantic` 那条 mark 从未出现** ⇒ 失控点在 Pass F（CoClass 形状校验）到 stage 3.4 成员合并之间。

旁证：`--syntax-only` 同样在 36.3 s / 2051.8 MB 被打断，默认 x64 架构同样 36.7 s / 2064.7 MB
⇒ 与 `--arch`、与 codegen/emit 无关，是**共用的工程装配阶段**。

判断：§6.5 的上一份普查（`ba455c7f`，CI run 42599592249）**跑完了**（`EMIT_EXIT=0`、792.7 s、
stdout 48,341,497 B）；按 144 MB/s 的斜率它撑不到 792 s ⇒ 这是一次**回归**，落在
`ba455c7f..2a2c3d29` 这 45 个提交里（iface 视图那批 43528bb8 / 650b7bac / 1075d990 / d7410cc3 是首要嫌疑）。
本轮**不猜**：下一轮第一步是逐刀二分构建 + 在该阶段加一行逐模块 stderr，定到具体提交与具体循环。

因此本轮拿不到「各诊断码条数 / 与上次普查的对比 / 距离 exe 还差什么」——因为**连 C 都没发出来**，
`cl`/`link` 无从谈起。当前第一个缺口不是语法，是装配阶段的内存/超线性行为。

## 25. 2026-11-10 内存"A/B 定案"：**不是回归；根因是每模块深拷贝整份类型库符号**（`ba50be1a` 起）

工具：`.temp\pd_census2\probe_v.ps1`（250 ms 采样 `PeakWorkingSet64`）；语料 `tests/photodemon/PhotoDemon.vbp`
（539 个模块，仓库内路径，cwd = 仓库根）；构建：`.temp\build-rel`（Release/Ninja，VS2019 14.29，`VB6C3_ENABLE_LLVM=OFF`）；
锚点：`.temp\wt-ba` worktree 独立 Release 构建。临时插桩为逐模块 stderr
`C3: SEMTRACE <i>/<N> <ms>ms ws=<MB>MB nsym=<模块级符号数> <模块名>`（两棵树同款；**已删除**）。

**A/B（同语料、同参数 `--syntax-only`、4 GB 上限）**

| 树 | 峰值 | 耗时 | 到被杀时已完成 | 该点工作集 | nsym | 斜率 |
| --- | --- | --- | --- | --- | --- | --- |
| HEAD（含 SEMTRACE） | 4078.9 MB | 8.8 s | 180/539（6189 ms） | 4069 MB | 3042 | 22.4 MB/模块 |
| `ba455c7f`（含 SEMTRACE） | 4176.7 MB | 8.9 s | 184/539（6242 ms） | 4173 MB | 3042 | 22.5 MB/模块 |

两版**不可区分** ⇒ `ba455c7f..2a2c3d29` 45 个提交**无罪**，二分没有可定位的提交。§24 的"回归"结论作废。

**形态（逐模块读数，无拐点）**

`1/539 22ms ws=272MB` / `2 53ms ws=295` / `3 85ms ws=317` / … / `88 3207ms ws=2123` / … / `180 6189ms ws=4069`。
模块 1 之前是 parse 1553 ms + typelib/prepasses 1006 ms（基线 272 MB），此后**每模块匀速 +21~23 MB**，
没有任何一个模块变陡 ⇒ 不是"某模块内容重"，是**每模块固定付一份全工程量级的钱**。

**根因**

`driver_semantics.cpp` 的逐模块循环（`for (auto& module : modules_)`）内部，第 257 行起把
`typelibParser_->cachedResults()` 的**全部** coclass / 接口成员 / 枚举 / 模块函数重新构造成 `Symbol`
塞进**该模块自己的** `SymbolTable`。实测每个分析器的模块级符号数 `nsym=3042`（全部来自类型库）。
539 个 `SemanticAnalyzer` 又全部被 `analyzers_.push_back()` 保留（crossmod / generics / codegen 都要用，不能提前释放）
⇒ 1.64 M 个 Symbol 对象常驻 ≈ 12 GB。`git log -S` 表明这段自 `ba50be1a`（2026-06-26，接口 vtable/早期绑定）就在，
不是最近 45 个提交。

**为什么 CI 能跑完**：CI 是 16 GB runner，792.7 s 成功说明它把 ~12 GB 吃了下来；本机 15.7 GB 总内存，
2 GB 红线先杀 — 差别是红线，不是代码。

**修法（已落地，2026-11-10）**：把类型库符号集**只造一份**（`Driver::buildSharedTypelibSymbols()` →
`typelibTable_`），`SymbolTable` 增加一层非拥有只读回退（`Scope::sharedScope_`，本地未命中才查共享层；
本地同名声明用 `shadowShared` 复现"遮蔽内建符号"语义）。写点审计：payload 字段只有该函数一处写点，
`scripts/check_typelib_payload_writers.ps1` 把"payload 只读"钉成断言。

**规模哨兵（已加）**：`scripts/check_typelib_scale_sentinel.ps1`，接进 `run_tests.ps1` 的
`[STATIC] typelib_scale_sentinel`。语料 = N 份同构小类 + shell32/stdole2 两条 Reference，
`--syntax-only` 跑 N1=8 / N2=64，判据是 **per-module 斜率 ≤ 5.5 MB/模块**（+ 峰值 ≤ 700 MB、墙钟 ≤ 180 s）。

**实测（Release，本机，2026-11-10）**

| 版本 | N=8 | N=64 | 斜率 |
| --- | --- | --- | --- |
| 改造前 `.temp\wt-ba\.build-rel`（`ba455c7f`） | 61.9 MB | 490.0 MB | **7.64 MB/模块** |
| 共享层后 `.build\C3.exe` | 45.3 MB | 274.9 MB | **4.10 MB/模块** |

`--emit-c` 同口径：基线 73.6 → 497.5（7.57 MB/模块），共享后 49.4 → 282.4（4.16 MB/模块）。
带 Reference 与不带 Reference 的斜率相同（274.8 vs 274.9 @ N=64）⇒ 剩下的 4.1 MB/模块**不是**类型库符号，
是每模块自身的 AST / 语义表 / 发码缓冲。

**⚠ 未解决**：整工程 `PhotoDemon.vbp --emit-c --arch x86` 在共享层后**仍然爆内存** ——
峰值 12,518 MB、约 52 s 后 `bad allocation`（`.temp\pd_run\post_ucprop.log.err`，stdout 0 B）。
共享层是**必要但不充分**：它拿掉了 3.5 MB/模块 的类型库克隆，但 539 模块 × ~20 MB 的模块自身开销
仍远超 16 GB。整工程验证只走 CI；本机按 §0.1 禁止再跑。

## 26. UC 属性桥 Let/Set 双入口（2026-11-10）

### 26.1 VB6 oracle（4 条，真 VB6 6.00.9782，x86，`.temp/oracle/letset`）

| 调用点 | 落到的入口 | 说明 |
| --- | --- | --- |
| `uc.X = 5` | Property Let | |
| `Set uc.X = obj` | Property Set | |
| `Set uc.X = Nothing` | Property Set | |
| `uc.X = Nothing`（**不带** Set） | Property Set | 关键：VB6 对 `Nothing` 字面量恒走 PUTREF |
| `uc.X = Empty` | Property Let | |
| `uc.X = obj`（不带 Set） | Property Let | 推翻"值是不是对象"分流 |
| `GET uc.X` | 返回 Set 写进去的那个值 | 两个入口写**同一份**属性存储 |

分流依据是**调用点形态**（有没有 Set / 右边是不是 Nothing），不是值的类型。C3 原先按"哪个符号先到"
（哈希顺序）挑一个 thunk，导致 7 份发码在 `prop_let_` / `prop_set_` 之间随机翻转（账 #418 那 13 行）。

### 26.2 实现

- `cgen_form_uc_props.inc`：`UcPropPick` 分开记 `letSeg`/`setSeg`（不再 merge）；**两侧都声明且同族**
  才发两个 thunk（`ucPropLet_X` → `prop_let_`，`ucPropSet_X` → `prop_set_`），否则保持今天的单 thunk、
  两个入口都退到它。`vb6_UcPropDesc` 加第 4 槽 `setref`（NULL = 退化到 `set`）。
- RTL：`vb6_UC_OwnPropSetRef` → `vb6_Host_SetPropRef` → `HW_Invoke` 按 `DISPATCH_PROPERTYPUTREF` 分流；
  `vb6com_invoke.c` 的 SetRef 接收者走 `vb6_Host_SetPropRef`。
- `cgen_assign_com_prop.inc`：`o.Prop = Nothing`（无 Set）改发 `vb6_ComSetRef(o, L"Prop", NULL)`。
- `cgen_setlet_set_prop.inc`：`Set <晚绑定接收者>.Prop = v` 原先按成员名全局命中工程类 Property Set
  **直发** `vb6_Cls_prop_set_X(o, …)`（`o` 实际是 HWND / IDispatch*）⇒ 0xC0000005。现加降级门：
  接收者推不出类名 且 是 `As Object`/`As Variant` 形参或变量、且不是窗体控件/UC 实例 ⇒ 走 `vb6_ComSetRef`。

### 26.3 探针读数（`.temp/ab13/ucprobe`，7 行一条不少）

`Module1.bas` 里 7 个 `Public Sub Pn(ByVal o As Object)` 强制晚绑定，UC 的 `Property Get/Let/Set X As Object`
各写一段 trace。改前（Set 优先）C1 = `SET[Object]`（错），且 C2 直接 AV；改后：

```
C1 o.X = 5           -> LET[Object]
C2 Set o.X = obj     -> SET[Object]
C3 Set o.X = Nothing -> SET[Object]
C4 o.X = Nothing     -> SET[Object]
C5 o.X = Empty       -> LET[Object]
C6 o.X = obj (no Set)-> LET[Object]
C7 GET               -> late-bound: Empty / early-bound(ucP1.X): Object(非 Nothing)
```

C7 的 `early-bound = Object` 证明 Let/Set 写的是同一份存储（最后写入方是 C6 的 Let，GET 读得回来）。
`late-bound = Empty` 是另一个既有缺口：`Property Get X As Object` 不生成 getter 桥
（`objCTypeOf` 只认 `vb6_cls_*`，`void*` 一律 `continue`），不是本轮引入，本轮不动。

### 26.4 已知缺口（明确不修）

1. `Set uc.X = obj` 但属性**只有** `Property Let`：仍退化到单 thunk（今天的形状），不新造错误腿。
2. `As Variant` 形参的属性完全不发桥（`Property Let X(v As Variant)` 晚绑定访问不到），与 `As Object`
   的 getter 缺口同上 —— 记缺口，别顺手加。
3. `Is Nothing` 对 `As Object` 形参恒为 False（探针 trace 里 `:null/:obj` 标记失效，`ObjPtr` 返回栈地址），
   与本刀无关，未修。

### 26.5 运行等价性（13 行到底影不影响运行）

| 工程 | 口径 | 结果 |
| --- | --- | --- |
| ucTreeView | `tests/guisample/uctreeview_px.ps1`（光标钉 (4,4)，2 轮/侧） | 除 `at=`（光标坐标）与 `try=`（前台重试计数）外**逐项相同**；base 自身两轮也只差 `at=` |
| VBFlexGridDemo | UC dump 4 份 + FORM dump | **逐字节相同**（sha 3a6851… / 82de16…） |
| Charts 2020 | 标题带 y=8..26（`band.ps1`，4 轮） | **0 差异**；图表主体随机数据驱动，同二进制两轮也不同，不可作读数 |

629 份全量诊断对账（`scripts/emit_manifest.ps1 -ListInputs` 枚举）：`inputs=629 / diff_inputs=0 / rc_mismatch=0`。

### 26.6 那 13 行变成了什么（形状门 7 份 → 15 份）

`emit_manifest.ps1` 全量复跑（Release `.build\C3.exe`，8,001,024 B，sha256 `1ef0b652…`）：
`# arch=x64 inputs=629 / 期望 629 行 / 实得 629 行 / 相同 614 / 哈希不同 15 / 缺席 0 / 多出 0`。

- **原来的 7 份（账 #418 那 13 行）全部落在新的 15 份里，交集 = 7**：`2020/Proyecto1.vbp`、
  `2020/ucChartArea`、`2020/ucChartBar`、`2020/ucPieChart`、`2020/ucTreeMaps`、`tests/ucTreeView`、
  `tests/VBFlexGridDemo`。
- 新增 8 份（`ucProgressCircular`、`ucSlider_2022`、`ucSegmentedControl`、`TeeChartChartDisplay`、
  `czUI-main/czFormDemo.vbp`、`guisample/GUISample.vbp` ×3）不是新语义，是同一形状变化的必然结果：
  凡**同时声明 Let+Set** 的属性都多一个 Let thunk，`vb6_UcPropDesc` 表项从 3 槽变 4 槽。
- ucTreeView 差分实证：emitted C `965,088 B → 968,550 B`，新增
  `ucPropLet_DragIcon` / `ucPropLet_Font` / `ucPropLet_SubTextFont`；表项形如
  `{ L"Font", NULL, ucPropLet_Font, ucPropSet_Font },`，单侧属性仍是 set 与 setref 同函数。

**性质判断**：这 13 行原先的翻转是"哈希顺序偶然"（`UcPropPick` 按符号先到者挑一个 thunk，7 份里挑中
Let 还是 Set 取决于目录枚举顺序），现在是**按 oracle 定的调用点规则**的确定性产物 —— 同一份输入两次
跑逐字节相同，与目录遍历顺序无关。

确定性取证（ucTreeView `tests\ucTreeView\Proyecto1.vbp`，连跑 3 次 `--emit-c`）：

```
run1 rc=0 bytes=968550 sha=CFCEA466B178F194F921533521C22387DC717A212B7328149F81382ED748A61F
run2 rc=0 bytes=968550 sha=CFCEA466...(同)
run3 rc=0 bytes=968550 sha=CFCEA466...(同)      distinct=1
```

双入口只出现在确实同时声明 Let+Set 的属性上（本工程 3 个：`DragIcon` / `Font` / `SubTextFont`），
其余 20+ 个仍是单个 `ucPropSet_*`（setref 同函数），符合"只声明一个的保持现状"。

**待批**：`emit-manifest.expected.txt` 这 15 行**尚未 bless**（保持红）。按此前裁决，bless 需单独批准；
批准后再跑一次 gate 确认 15→0。

## 27. 2026-11-10 整工程 emit 基线（CI，重构前对照）

分支 `probe/photodemon-emit` 已合并 `dev`（含 `aa715a0d` 共享类型库层、R1–R4 与 UC 桥），
CI run **42725438757**（`fe5cb1a3`，16 GB runner）跑 `PhotoDemon.vbp --arch x86 --emit-c`。
探针 workflow 本轮加了最小采样：emit 子进程每 2 s 记一次工作集 / 系统空闲物理内存 / stderr 字节数，
逐码计数与 stderr 分类落 artifact（`photodemon-emit-logs`）。

### 27.1 读数（改动前基准）

| 项 | 读数 |
| --- | --- |
| `EMIT_EXIT` | **0** |
| 墙钟 | **659.2 s** |
| stdout | **48,358,712 B** |
| stderr 总行 | **200** |
| 诊断行 | **57**，逐码：`VB3003` 57，其余 0（`VB3001`/`VB3044`/`VB4001` 均为 0） |
| 非诊断行 | **143** = `CoClass '…' identity` 142 + `AddressOf callback procs` 1 |

### 27.2 内存（CI **有**读数，但是"工作集"口径）

| 项 | 读数 |
| --- | --- |
| 采样峰值工作集 | **13,544.9 MB**（t≈66 s） |
| 采样最小空闲物理内存 | **~4 MB**（多次落在 80–500 MB） |
| ≥12 GB 的采样时长 | ~38 s（2 s 网格） |
| ≥10 GB 的采样时长 | ~108 s |

曲线（秒 → 工作集）：10 s 6.8 GB / 16 s 11.5 / 29 s 13.2 / 50 s 13.4 / 91 s 12.9 /
179 s 8.4 / 300 s 7.3 / 421 s 8.6 / 539 s 5.8 / 600 s 2.0 / 651 s 0.5。

口径说明（不许拿这张表当 commit 峰值）：空闲物理内存掉到 4 MB ⇒ runner 已在换页，
工作集被系统修剪，**13.5 GB 是下限不是上限**，真实提交量未采到。且采样用 `Get-CimInstance`
（每 2 s），墙钟含采样开销；`ba455c7f` 的 792.7 s 是另一份代码，两者不可直接比。

### 27.3 与旧基线的差异

- `VB3001` 3693 → **0**、`VB3044` 19 → **0**：R1（Public 变量工程级可见）与 R4（可创建空体接口宿主）在整工程口径成立。
- `VB3003` 57（R2/R3 后）保持，未再降。
- 耗时 792.7 s → 659.2 s，但代码与采样都变了，**不当作性能结论**。

### 27.4 相位归因（已补：run 42727942019，`C3_TIMING=1`）

同一探针第二次跑（`d53ecf84`）带 `C3_TIMING=1`，用"stderr 字节数"把相位行对回 2 s 采样时间轴：

| 相位结束 | 时刻 | 工作集 | 该相位墙钟 |
| --- | --- | --- | --- |
| parse | ~2 s | 704 MB | 1014 ms |
| typelib+prepasses | ~2 s | 704 MB | 505 ms |
| semantic | ~12 s | **10,528 MB** | 9243 ms |
| inherit-merge | ~12 s | 10,528 MB | 0 ms |
| crossmod+generics+marks | ~390 s | 8,902 MB（峰值 **13,204 MB @94 s**） | 378,117 ms |
| codegen（余下 539 模块） | ~1410 s | 平台期 8–10 GB，末段 <1 GB | ~1020 s |

**斜率陡增在 semantic**：2 s→12 s，704 MB→10.5 GB（≈1 GB/s，≈18 MB/模块）。
**AST 不是大头**：parse+typelib+prepasses 结束时总共 704 MB（592 份源、20 MB 文本）。
**第二个大头是 3.5 跨模块注入**：峰值 13.2 GB 出现在注入区间（+2.7 GB），墙钟 378 s；
注入是"每个消费模块 × 每个 Public 符号"的 35 字段深拷贝（[driver_crossmod.cpp#L365-L433](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_crossmod.cpp#L365-L433)）。
codegen 阶段没有新的分配峰值（8–10 GB 平台期是语义表仍常驻 + 换页修剪）。

口径警告（必须一起读）：
- 本轮墙钟 1410.3 s，不带 TIMING 的基线是 659.2 s（2.1×）。**这一轮只能用来做相位归因，不能当时间基线。**
- 空闲物理内存最低 **16 MB**，工作集被系统修剪，表里的 GB 是**下限**；未采到 commit 峰值。
- 相位边界用 stderr 落盘字节对齐，粒度约 2 s；semantic 的真实结束时刻在 10–12 s 之间。

### 28. `analyzers_` 生命周期审计（只读，2026-11-10）

#### 28.1 消费者清单

| # | 消费者 | 阶段 | 要什么 | 类别 | 可行替代 | 风险 |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | `dumpSymbols` [driver_semantics.cpp#L672](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_semantics.cpp#L672) | 3.1 循环内 | 自己一张表 | 仅自身 | 用完即弃 | 无 |
| 2 | `mergeInheritedMembers` [driver_classchain.cpp#L589-L598](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_classchain.cpp#L589-L598) | 3.4 | 自己 + 每个祖先的 Class 符号（成员表），**写回**自身 | 完整符号载荷 + 祖先 AST 指针（经 `classes_`） | 必须先算好"类 → 祖先成员"摘要；写入点仅 Class 符号 | 中：合并结果被 cl 端当结构体字段用，顺序/去重一变就出 C2039 |
| 3 | `markAddressOfCallbacks` [driver_crossmod.cpp#L35-L96](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_crossmod.cpp#L35-L96) | 3.5a | 遍历**所有**模块 AST + 查/写**所有**模块符号表 | 完整 AST + 完整表 | 改成"分析期就地标记"（AddressOf 在单模块内可判定） | 低-中：判定规则若依赖跨模块可见性则不能就地 |
| 4 | 3.5 注入 [driver_crossmod.cpp#L140-L515](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_crossmod.cpp#L140-L515) | 3.5 | 每模块 `getPublicSymbols()` 全载荷；对每个消费模块深拷贝 35 字段；回填 UDT 成员、`comProjectImplClass`；窗体限定键；祖先 Private UDT/Enum；再逐 analyzer 跑 `resolveDeferredCrossModuleOverloads` | 完整符号载荷（全部模块） | ①载荷改共享只读（不拷贝）；②或先做"工程级摘要表"再按需注入 | 高：`Symbol` 字段与写点遍布；`scripts/check_typelib_payload_writers.ps1` 钉了载荷写点 |
| 5 | 泛型 fixpoint [driver_generics.cpp#L408-L444](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_generics.cpp#L408-L444) | 3.5b | `takeGenericRequests()` + `analyzeExtraDecls()`（要 analyzer 活着，含 AST/类型系统） | analyzer 内部状态 | 保留"请求队列"摘要，物化期重建 analyzer | 中：PhotoDemon 无泛型，但回归语料有 |
| 6 | iface 标记 [driver_compile.cpp#L745-L756](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_compile.cpp#L745-L756) | 3.6 | 扫**所有**表，把匹配的 Class 符号 `isInterface=true` | 名称 + 写标志 | 改成"声明模块自己标记 + 名称索引查询" | 低：只写一个 bool |
| 7 | codegen voidfield 扫描 [driver_codegen_voidfield_scan.inc#L49-L68](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/detail/driver_codegen_voidfield_scan.inc#L49-L68) | codegen | 所有模块 Class 的**名字 + implementsNames** | 摘要 | 轻量索引 | 低 |
| 8 | codegen typedfield 扫描 [driver_codegen_typedfield_scan.inc#L24-L83](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/detail/driver_codegen_typedfield_scan.inc#L24-L83) | codegen | 所有模块 Public Constant 的**值/类型** + 本模块字段类型 | 摘要 + 自身表 | 轻量索引（常量表） | 低 |
| 9 | dup-module-vars [driver_codegen_dup_module_vars.inc#L29](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/detail/driver_codegen_dup_module_vars.inc#L29) | codegen | 本模块 Class 符号 | 仅自身 | 直接 | 无 |
| 10 | 逐模块 cgen [driver_codegen_module_loop.inc#L56-L82](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/detail/driver_codegen_module_loop.inc#L56-L82) | codegen | 本模块 SymbolTable + TypeSystem + AST（发码主体） | 完整 | 必须常驻到该模块发完码 | 低（天然可流式） |
| 11 | DLL 入口/typelib [driver_codegen_dll_entry.inc#L48-L54](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/detail/driver_codegen_dll_entry.inc#L48-L54)、[driver_codegen_dll_typelib.inc#L33-L87](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/detail/driver_codegen_dll_typelib.inc#L33-L87) | codegen（仅 `--dll`） | **所有** SymbolTable 指针 | 完整表 | 只在 DLL 时聚合 | 中：EXE 路径也走 `generateDllEntry(..., allSymTabs)`（L54） |

#### 28.2 跨模块顺序依赖

- **分析期（3.1）无跨 analyzer 依赖**：`analyzers_` 只在 driver 里被访问，`analyze()` 不读别的 analyzer。
  但依赖"**所有模块 AST 先就位**"：3.1 每个模块前会扫其它模块 AST 预登记 Enum/UDT/Public 名
  （[driver_semantics.cpp#L528-L663](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_semantics.cpp#L528-L663)）。
- **3.5 全局合并对顺序敏感**：`storageKey` 先到先得，外加显式裁决（.bas 胜 .cls、getter 胜 setter、
  UDT 成员空→非空换人），见 [driver_crossmod.cpp#L161-L250](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/driver/driver_crossmod.cpp#L161-L250)。
  这条顺序必须逐位保留（`emit-manifest` 是门）。
- **隐藏陷阱**：`getPublicSymbols()` 不排除 `isExternal`（[symbol_table.cpp#L528-L550](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/symbol_table.cpp#L528-L550)），
  泛型 fixpoint 第二轮会把上一轮注入的副本再导出一次。任何"摘要/共享"重构都要么复刻这个行为，要么用形状门证明等价。
- **阶段链**：3.1 → 3.4 → 3.5a → 3.5 → 3.5b → 3.6 → codegen；3.6 必须在注入之后（消费模块里的 iface 副本也要标记）。
- **AST 指针外泄**：`classes_` / `ifaces_` / `genericTemplates_` / `legacy` 形状都直接持有 `Module*`/`Decl*`，
  所以"放掉 AST"不是局部改动，得先把这几张注册表改成不持指针（或自有摘要）。

#### 28.3 候选方案

| 方案 | 改动面 | 预期收益 | 风险 |
| --- | --- | --- | --- |
| **A. 注入载荷共享（不动生命周期）** | `Symbol` 载荷拆"不可变共享部分 + 可变字段"；3.5 注入只建轻量壳；3.6 标记改旁路集合 | 砍掉 +2.7 GB 峰值与 378 s 注入时间；不改变阶段顺序 | 中-高：`Symbol` 字段/写点遍布，`check_typelib_payload_writers.ps1` 要同步；`comProjectImplClass`/`udtMembers` 回填必须留写权 |
| **B. 两遍流式（真"一次只活一个模块"）** | ①第一遍 parse+analyze 后导出工程摘要（Public 符号 + 常量 + coclass/implements 名 + 类祖先成员）；②第二遍重解析/重分析，注入摘要，发码后立刻释放；③注册表去 AST 指针 | 常驻降到 O(1 模块 + 摘要)，目标 <1.5 GB | 高：等价性全靠 `emit-manifest`；解析+分析跑两遍（这里约 +11 s，可接受）；`--dump-symbols`/DLL 全表聚合要特判 |
| **C. 只压 per-module 语义表** | Symbol 池化、字符串 intern、成员表按需/共享、局部变量符号分析后释放 | 线性降 30–50%（10.5 GB → 5–7 GB），最省事 | 低-中；可能仍要 6 GB+，单独不够，但可与 A/B 叠加 |

**判断**：13.5 GB 里 AST 只占 0.7 GB，**先放 AST 没有意义**。要达到"一次一个模块"，B 是唯一彻底的路，
但 B 依赖 A（否则摘要里 Public 载荷仍是全量拷贝）；A 可以独立落地、独立用 `emit-manifest` 验。

## 29. 方案 A 落地：跨模块注入载荷共享（2026-11-10）

### 29.1 第 1 刀：类成员表 11 张收进 `SymbolPayload`（`951574f8`）

- `Symbol` 新增 `std::shared_ptr<SymbolPayload> payload_`，读走 `P()`（const 引用，空载荷返回静态默认实例），
  写走 `W()`（`use_count()>1` 时 COW 克隆）；3.5 注入改用 `sharePayloadFrom()`，只花一个 `shared_ptr`。
- 11 张类成员表（`memberNames` … `memberSetParams`）的读写点全部改门（19 个文件、221 处）。
- 形状门：HEAD 现编旧二进制 vs 新二进制，**629/629 逐字节零差异**。
- 斜率（`check_typelib_scale_sentinel.ps1` 语料，本机）：N=8 86.7 → **56.9 MB**，N=60 585.2 → **324.2 MB**，
  **9.59 → 5.14 MB/模块**。
- `-Category syntax` 187/187 PASS。`check_typelib_payload_writers.ps1` 绿（`driver_crossmod.cpp` 载荷写点
  3 → 2，少的正是被共享替换掉的 `extSym->memberNames = srcSym->memberNames`；S6「读门不许当写门」新增）。

### 29.2 第 1.5 刀：`ComMethodSig` / `DelegateTarget` 上提到命名空间作用域（`7b991465`）

第二刀要把 `comMethods`/`comSourceMethods`/`comModuleFunctions`/`delegateTargets` 收进 `SymbolPayload`，
而载荷定义在 `Symbol` 之前，嵌套类型在那里不可见。故先做纯机械上提：字段与方法签名一字不改，
`Symbol` 内保留同名 `using` 别名，既有 `Symbol::ComMethodSig` 写法继续有效。

- 形状门：629/629 逐字节零差异（`emit-manifest.old.txt` vs `emit-manifest.cut2.txt`，数据行 629 = 629，diff 0）。
- `-Category syntax` 187/187 PASS；两道哨兵读数不变（`driver_crossmod` 载荷写点仍为 0）。

### 29.3 形状门期望文件挂账（**读之前先看这条**）

`emit-manifest.expected.txt` 当前落后 dev **15 行**，其中 **4 行与本系列改动无关**（dev 上早已存在的漂移）。
判读口径：

- 本系列每一刀都用**同一个 HEAD 现编的旧二进制**跑全量 629 输入，与改动后的新二进制逐字节比较；
  得到 0 差异 = 本刀没有改变发码形状。这个比较比「对 expected 文件」更强，因为 expected 本身陈旧。
- `emit_manifest.ps1` 末尾与 expected 比会报 15 行 FAIL（脚本 rc=1/4），**这是期望文件陈旧，不是代码回归**。
  不要据此回滚，也不要逐行手改 expected。
- 登记方式：等 CI 那台二进制交回清单后，用 `scripts/rebless_emit_manifest.ps1 -Bless` 统一重登记一次。
  在此之前，本系列所有「形状门通过」均指上述「旧 vs 新 629 零差异」口径。

### 29.4 第 2 刀：COM / Delegate 表迁入载荷（`fa70cfe2`）

- 迁入 `SymbolPayload`：`comMethods` / `comSourceMethods` / `comModuleFunctions` 三张签名表，
  以及 `delegateProcKind` / `delegateCallConv` / `delegateReturnType` / `delegateTargets`。
- 读写点改门（8 个文件）：读 `P()`、写 `W()`。`insertComMethod` 的 `find` 走 `P()`，
  末尾写入走 `W()` —— 迭代器在写入前已用完（`comSigReadShapeDiffers` 与 COMSIG-AMBIG 都在这之前），
  无失效窗口；COW 保证共享载荷不会被原地改。
- 形状门：上一刀二进制 vs 本刀，**629/629 逐字节零差异**（`emit-manifest.cut2.txt` vs `cut3.txt`）。
- 斜率（同一语料）：N=8 57.0 → **50.6 MB**，N=60 324.2 → **230.1 MB**，
  **5.14 → 3.45 MB/模块**。
- `-Category syntax` 187/187 PASS。
- 哨兵：`check_typelib_payload_writers.ps1` 绿（核心字段 22 个不变 —— 新迁入的 delegate 字段不在注入块写，
  不产生新权威点）；`check_typelib_scale_sentinel.ps1` 的 S0 载荷字段表同步扩充到 18 个，
  跨模块注入侧仍要求 `sharePayloadFrom` 且零逐字段拷贝。
  该脚本本身仍有两条**改前就存在**的 FAIL（`ExitCode` 读成空 + 斜率阈值 2.0），与本系列改动无关，未粉饰。

### 29.5 第 3 刀：过程形参表 `params` 迁入载荷（`c5b20db7`）

- 枚举方式：**先删 `Symbol::params` 字段，让编译器逐个报错定位**，再按接收者类型分别改 `P()` / `W()`。
  AST 侧的 `node.params`（Sub/Function/Property 声明节点）、`MemberInfo::params`、`ComMethodSig::params`
  同名但类型不同，全部未被误改；没有做任何批量替换。
- 3.5 注入：删除 `extSym->params = srcSym->params;`（由 `sharePayloadFrom` 一并共享）。
  窗体限定二级键是例外：`qExt->W().params = qSym->P().params;`（这条是**新建**限定副本，不是共享源本体）。
- 事件 ABI 回写（`applyEventHandlerAbi` 把 `isByVal` 改 true）改走 `evtSym->W().params`，COW 保证不污染共享源。
- `--dump-symbols`：输出代码从 `sym->params` 改成 `sym->P().params`，**输出文本不变**（形状门覆盖不到 dump，
  这条靠代码审读；dump 顺序/格式未动）。
- 形状门：上一刀二进制 vs 本刀，**629/629 逐字节零差异**（`emit-manifest.cut3.txt` vs `cut4.txt`）。
- `-Category syntax`：**187/187 PASS**。
- 斜率（同一语料）：N=8 **51.1 MB**、N=60 **268.3 MB**，**4.18 MB/模块**。
  **注意方向相反**：上一刀 3.45 → 本刀 4.18（+0.73）。重复 3 次稳定；用 pre-cut2 二进制复测同语料为
  5.12 MB/模块，排除随机噪声与后台负载。原因未定位（怀疑 procedure 载荷首建/COW 克隆路径），
  **整工程是否受益以 §30 的 CI 读数为准；若峰值/注入段不改善，回退本刀，保留前三刀。**
- 哨兵：`check_typelib_payload_writers.ps1` 绿（核心字段 22 个不变，注入权威区仍只 `driver_semantics.cpp`）；
  `check_typelib_scale_sentinel.ps1` 的 S0 载荷字段表扩到 **19 个**（加入 `params`）。
  该脚本自身两条改前既有 FAIL 仍在（`ExitCode` 读空、斜率阈值 2.0），未粉饰。

## 30. CI 整工程复跑（2026-11-10，载荷共享三刀合并后）

把 `dev`（`c5b20db7`）合到 `probe/photodemon-emit`（merge `f3979e10`），探针 workflow 加 `timing` 开关
（`13774b90`：push 默认**不带** `C3_TIMING`，相位归因用 `workflow_dispatch -f timing=1`），跑两次：

| run | 触发 | `C3_TIMING` | 用途 |
| --- | --- | --- | --- |
| 42753537997 | push | 关 | 与 §27 基线同口径比墙钟/峰值/stdout/诊断 |
| 42753601684 | workflow_dispatch | 开 | 相位耗时与注入段归因 |

### 30.1 同口径对照（都不带 `C3_TIMING`）

| 项 | 基线 42725438757 | 本次 42753537997 | 变化 |
| --- | --- | --- | --- |
| `EMIT_EXIT` | 0 | **0** | — |
| 墙钟 | 659.2 s | **71.4 s** | **−587.8 s（−89.2%）** |
| stdout | 48,358,712 B | **48,358,712 B** | sha256 `5CB9D6DF…96EE0` **逐字节相同** |
| stderr 行 | 200 | **200** | 相同 |
| 诊断 | VB3003=57，其余 0 | **VB3003=57，其余 0** | 相同 |
| 峰值工作集 | 13,544.9 MB | **8,456 MB** | **−5,088.9 MB（−37.6%）** |
| 最小空闲物理内存 | ~2 MB（`meta`）/ 4 MB（csv） | **5,211 MB** | **+5.2 GB 余量** |

`VB3001=0`、`VB3044=0`、`VB4001=0` 均未变；`VB3003=57` 未变。非诊断行仍是
142 条 `CoClass '…' identity` + 1 条 `AddressOf callback procs`。
**结论：发码逐字节不变，峰值与墙钟都实质下降，`params` 刀不需要回退。**

### 30.2 相位（都带 `C3_TIMING`，两次之间才能比）

| 相位 | 基线 42727942019 | 本次 42753601684 | 变化 |
| --- | --- | --- | --- |
| 总墙钟（带 TIMING） | 1410.3 s | **73.4 s** | −94.8% |
| parse | 1,014 ms | 1,211 ms | +197 ms |
| typelib+prepasses | 505 ms | 1,102 ms | +597 ms |
| semantic | 9,243 ms | **5,937 ms** | −3,306 ms |
| inherit-merge | 0 ms | 1 ms | 噪声 |
| **crossmod+generics+marks** | **378,117 ms** | **11,727 ms** | **−366,390 ms（−96.9%）** |
| unpack-rtl | 291 ms | 92 ms | −199 ms |
| codegen 余量（总墙钟减上面） | ≈1,021 s | **≈53 s** | −94.8% |
| 峰值工作集（TIMING 口径） | 13,204 MB | **8,454 MB** | −4,750 MB |
| 最小空闲物理内存 | 16 MB | **5,181 MB** | +5.2 GB |

**时间口径必须分开说**：基线那次带 TIMING 是 1410 s、不带是 659 s（2.1×），§27.4 已判为
"只能做相位归因、不能当时间基线"（当时空闲物理内存只剩 16 MB，整机在换页）。
本次带 TIMING 73.4 s、不带 71.4 s（+2.0 s，2.8%），**TIMING 开销本身可以忽略**；
墙钟的下降不是 TIMING 口径差造成的。

### 30.3 跨模块注入段（3.5）

用 stderr 字节偏移把相位边界对回 2 s 采样（本次 `crossmod` 行落在 18–20 s 采样之间，
与累计相位 8.25 s + 11.73 s ≈ 20.0 s 一致）：

| 项 | 基线（TIMING） | 本次（TIMING） |
| --- | --- | --- |
| 注入段起点工作集 | 10,528 MB @12 s | 3,274 MB @8 s |
| 注入段终点/窗口峰值 | 13,204 MB | 8,431 MB @20 s |
| 窗口内增量 | +2,676 MB | **+5,157 MB** |
| 注入段墙钟 | 378.1 s | **11.7 s** |

**两个"增量"不是同一个量，不许直接比**：基线起点 10.5 GB 是在空闲内存只剩 16 MB、
工作集已被系统修剪的状态下测的（真实提交量更高、增量被低估）；本次起点只有 3.3 GB，
机器全程有余量，读数可信。可比的是**绝对峰值**：13.2 GB → 8.45 GB（TIMING 口径）、
13.5 GB → 8.46 GB（不带 TIMING 口径）。注入耗时 378 s → 11.7 s 是干净的同比。

### 30.4 口径与保留项

- 本次两次 run 的 stdout sha256 与基线**完全一致**（48,358,712 B），因此"性能改善没有换来发码变化"是硬证据，
  比 §29.3 的 629 输入形状门更强。
- 本读数只证明**三刀合计**相对 `aa715a0d` 的收益；**没有**单独测 `params` 一刀（`fa70cfe2` vs `c5b20db7`）。
  §29.5 的合成语料斜率反向（3.45→4.18 MB/模块）在整工程口径没有兑现为回归，但也没有被解释掉；
  若以后要继续追，就在探针分支上跑一次 `fa70cfe2` 做对照。
- `--dump-symbols` 的可观测面：`params` 刀的 dump 代码从 `sym->params` 改读 `sym->P().params`，输出文本不变；
  本次 CI 不跑 dump，这条靠 §29.5 的代码审读 + 形状门兜底。

## 31. 三刀之后的 dev CI 归因（2026-11-10）

`dev` 推到 `github/dev`（`f12cc7ea`）触发三次 `VB6 C3 Regression`，结果逐轮收口：

| run | 结论 | 说明 |
| --- | --- | --- |
| 42754470554 | 22 job 中 3 红 | emit-manifest + compile ×2，见 31.1 |
| 42756305967 | 2 红 | BOM/W2 修好，emit-manifest 绿 |
| 42757586471 | 2 红 | 规模哨兵 rc 修好；**只剩 8 条既有 STATIC** |

### 31.1 形状门（emit manifest）

首轮红了 11 份：`2020/ucChart*`、`ucPieChart`、`ucProgressCircular`、`ucTreeMaps`、
`tests/czUI-main`、`VBFlexGridDemo`、`ucSegmentedControl`、`ucTreeView`。
**不是三刀引入**，证据链：
- 用 cut1 之前的二进制（`aa715a0d`，即 `951574f8` 的父提交）跑出的清单 `.temp/emit-manifest.old.txt`
  对这 11 份给出的 sha256 与 CI 本次"实得"**逐条相同**；
- 再拿 CI 清单（`ci/emit-manifest` 分支，bless 前）与 cut1 前清单做全量 629 行比对：
  **只有 4 行不同**（`GUISample` ×3 + `TeeChartChartDisplay`），而这 4 行正是 §29.3 挂账的
  "与本系列无关"那 4 行（本机 VS 16.11 vs CI VS 18.11 的工具链差异，期望文件里的就是 CI 值）。
- 因此：**三刀合计对 629 份语料的发码形状改动 = 0**；11 行属于 `aa715a0d` 的 UC 属性桥
  Let/Set 双入口（其提交信息自述的发码变化），已用 `rebless_emit_manifest.ps1 -Bless` 登记，
  复算 629/629 绿、`check_manifest_coverage` K1..K3 绿。

### 31.2 本系列引入并已修的 4 条

| 用例 | 原因 | 修法 |
| --- | --- | --- |
| `ps51_joinpath` / `test_helper_integrity`（H7）/ `static_sentinel_registration`（R4） | §29.5 编辑 `check_typelib_scale_sentinel.ps1` 时丢了 UTF-8 BOM，CI 的 PS 5.1 按 ANSI 读 → 解析失败 | 恢复 BOM（内容不动），三条本地 PS 5.1 复跑绿 |
| `com_sig_collision_policy`（W2） | 判据写死 `sym.comMethods.end()`；载荷共享后权威写法是 `sym.P().comMethods.end()` | 判据改认读门（闸本体与位置未动，不算放宽） |
| `typelib_scale_sentinel` | 哨兵自身缺陷：PS 5.1 下 `Start-Process -PassThru` 的 `ExitCode` 恒空 → 误报"C3 退出码非 0"（本机与 CI 都复现） | `Measure-Run` 改用 `System.Diagnostics.Process` + 异步抽干管道；PS 5.1 本地 PASS（N=8 51.1 MB / N=64 285.7 MB / 4.19 MB/模块，rc=0） |

### 31.3 剩下的 8 条：全部既有，与三刀无关

- `builtin_opt_defaults`：`vb6_Randomize` 缺"按实参个数"档的理由（更早的账）。
- `vbp_fixture_census`：`photodemon\PhotoDemon.vbp` 未登记（刻意只在探针分支跑）。
- `legacy_slot_key_owners` P3：9 vs 8（`cgen_iface_vtbl.cpp` 2 vs 1）。
- `iface_view_authority`：`ifaceViewOfExpr` 6 文件 vs 5、`vb6_ivref_` 7 vs 6、`emitIvrefRefSet` 5 vs 4。
- `legacy_view_arg_owners` V8：2 vs 1。
- `box_temp_reclaim` R1：3 vs 2。
- `class_refcount_authority`：8 vs 7。
- `object_slot_boxing` B11：`procObjectFormals_.count` 2 vs 1 —— 用 `git log -S` 定位到
  **`aa715a0d`**（`cgen_setlet_set_prop.inc` 的 Set 晚绑定接收者判断新增第二处直接问表）。

前 7 条在 `7230557e`（三刀之前的 dev）就红；第 8 条是 `aa715a0d` 引入、但**早于三刀**。
三刀没有新增任何 STATIC 红。**这些既有红本次不修**（不在本系列范围，且 `object_slot_boxing`
的正解是让 Set 站点改走 `inferClassTypeOfExpr` 那条唯一权威，属发码改动，需单独一刀 + 形状门）。

### 31.4 run 42757586471 的逐 job 结论（12 分片全绿）

20 个 job 里 18 绿 2 红，红的只有 `Tests (compile #1)` / `Tests (compile #2)`（上面 8 条 STATIC）：

| 绿 | job |
| --- | --- |
| 构建/形状 | `Build C3.exe`、`Emit manifest (shape oracle)`、`Baseline (M1)` |
| **vbp 分片 12/12** | `Tests (vbp #1)` … `Tests (vbp #12)` |
| 其它夹具 | `Tests (bas #1)`、`Tests (bas #2)`、`Tests (syntax)`、`Tests (asm)`、`Tests (smoke)` |

⇒ **三刀没有伤到既有夹具**：12 个 vbp 分片、bas/syntax/asm/smoke 与 629 形状门全部通过；
唯一的红是 compile #1/#2 里的 8 条既有 STATIC 哨兵。

## 32. 整编探针：emit → cl → link 的第一个真实读数（2026-11-10）

`dev` 三刀之后，`--emit-c` 已能整工程出 539 份 C（§30）。本节验证的是**下一段**：
把产出的 C 交给 `cl.exe` 编、再 link 成 exe。新 workflow `.github/workflows/pd_build_probe.yml`
（探针分支，`workflow_dispatch` + 本文件自己的 push），run **42754639429**（job 结论 `success`
只因整编 step 有 `continue-on-error: true`，**实际 `BUILD_EXIT=1`，没有 exe**）。

**这一节没有为了让它过而放宽阈值、桩掉函数或跳过一个文件。** 读数原样记录。

### 32.1 总量

| 项 | 读数 |
| --- | --- |
| 命令 | `.build\C3.exe tests/photodemon/PhotoDemon.vbp --arch x86 --output-dir .build/pd_full` |
| `C3_TIMING` | 1（时间口径只能跟同样带 TIMING 的读数比） |
| `BUILD_EXIT` | **1**（cl 返回 2，停在 cl 阶段，**未进 link**） |
| exe | **不存在** |
| 整编墙钟（含 emit + cl） | **6,596.2 s** |
| 峰值整棵进程树工作集 | **8,861 MB** |
| 最小空闲物理内存 | **4,858 MB**（runner 总 16,379 MB） |
| 工程 TU | **539 份，539 份都有 error** |
| 24 份 TU 撞 `C1003`（每 TU 100 错误上限） | 真实错误数**比下表更多**，被截断 |
| `c3-error.log` | 18,784,181 B / 85,020 行（cl 全量输出） |
| error 行合计 | **≈20,474** |
| warning 行合计 | **≈46,354** |

### 32.2 cl 错误码（按次数）

| 码 | 级别 | 次数 | 受影响 TU |
| --- | --- | --- | --- |
| C2079 | error | **16,152** | **539 / 539** |
| C2224 | error | 1,125 | 90 |
| C2039 | error | 971 | 84 |
| C2197 | error | 774 | 62 |
| C2198 | error | 549 | 112 |
| C2440 | error | 329 | 76 |
| C2065 | error | 278 | 99 |
| C2064 | error | 60 | 10 |
| C2045 | error | 50 | 17 |
| C2296 | error | 39 | 19 |
| C2036 | error | 34 | 4 |
| C1003 | **fatal** | 24 | 24（错误上限截断） |
| C2231 | error | 21 | 6 |
| C2063 | error | 19 | 19 |
| C2106 / C2069 / C2297 / C2095 / C2186 | error | 10 / 8 / 5 / 5 / 5 | ≤4 |
| C2102 / C2172 / C2223 / C2113 / C2112 | error | 4 / 4 / 3 / 3 / 2 | ≤3 |

warning 侧（不是失败原因，但量很大、且都是生成器形状问题）：C4005 宏重定义 **16,516**（≈30 个宏 × 539 TU）、
C4244 收窄 **8,440**、C4047 **7,628**、C4024 **6,940**、C4133 **2,932**、C4146 **2,185**、
C4090 587、C4028 560、C4022 352、C4305 143。

### 32.3 缺口分类（按"哪一类、影响多少文件"）

| # | 族 | 码 | 条数 | 影响 TU | 判据/位置 |
| --- | --- | --- | --- | --- | --- |
| G1 | **UDT 值类型在头里只有前向声明** | C2079 | **16,152** | **539（全）** | 6 个类型 `vb6_type_RectL` 3,771、`RGBQuad` 3,769、`PointAPI` 3,759、`PointLong` 2,690、`RectF` 1,617、`winRect` 539；落点 57 处，全在**头文件**里（`Fonts.h(319..327)`、`GDI_Plus.h(597..)`、`pdDisplays.h(113)`、`pdEditBoxW.h(201)`…），形如 `vb6_type_RectL rcMonitor;` 的值成员 |
| G2 | 接收者类型推不出（当非结构用） | C2224/C2296/C2231/C2036 | 1,225 | ~92 | `.MainSelection` 201、`.Width` 123、`.Height` 121、`.Left` 82、`.x/.y/.Top/.Right` … 接收者是 `void*` / 不完整类型 |
| G3 | 类实例的"方法名当字段发" | C2039 | 971 | 84 | **907 条集中在 `vb6_cls_pdImage`**：`GetActiveLayer` 237、`GetLayerByIndex` 112、`GetNumOfLayers` 66、`GetActiveLayerIndex` 54、`GetActiveDIB` 54、`IsSelectionActive` 53、`NotifyImageChanged` 51…（63 个不同成员名） |
| G4 | 实参个数与签名不符 | C2197 + C2198 | 1,323 | ~115 | 多传：`pdDropDown_AddItem` 144、`pdButtonStrip_AddItem` 126、`pdButtonToolbox_AssignTooltip` 72；少传：`ChkLong` 56、`Processor_FlagInitialNDFXState_Text` 36、`Int` 25、`IntDiv` 13。两侧都指向**控件/类方法的接收者与默认参数展开**，共 68 + 141 个不同被调方 |
| G5 | 值/指针/整数互转 | C2440 | 329 | 76 | `void* → float` 36、`vb6_type_RGBQuad → int` 31、`PointLong → int` 23、`ComplexNumber → intptr_t` 16、`RectF → int` 14、`int → VARIANT` 13、`double → pdSpinner**` 12 |
| G6 | 控件名/控件数组元素未声明 | C2065 | 278 | 99 | `pdFxPreview` 39、`sltQuality` 19、`vb6_hwnd_sltBrushSetting` 14、`MergedEntry` 12、`PD_MenuEntry` 10、`cmdFlyoutLock` 10（.frm 控件变量没进 TU） |
| G7 | 其它零散 | C2064 60（`takes 0`）、C2045 50（label 重复）、C2063 19、C2106 10、C2069 8（`void` 返回值被使用）、C2095 5、C2186 5 | ~157 | ≤20 | 单点形状问题，批量修不值当，先记账 |

**G1 的机理（最小复现，非推测）**：生成器在**每个模块头**里为用到但没定义的 UDT 写的是
`typedef struct vb6_type_X vb6_type_X;`（[cgen_base_generate_epilogue.inc:78-85](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/base/cgen_base_generate_epilogue.inc#L78-L85)），
而 `struct` 体只写进**定义方**的头。本机 2 模块探针（`Holder.cls` 里 `Private mR As RectL`，
`RectL` 由排在其后的 `ModA.bas` 定义）复现为：
`Holder.h` 里有 `vb6_type_RectL mR;` + 只有前向 typedef，`ModA.h` 里才有 `typedef struct vb6_type_RectL { … }`，
而 `Holder.c` 先 `#include "Holder.h"` ⇒ C2079。整工程里这 6 个 UDT 定义在
`Drawing2D.bas` / `PublicEnumsAndTypes.bas`，用到它们的头分散在 `Fonts.h`、`GDI_Plus.h` 等，
**头顺序由 vbp 顺序决定**，所以 539 个 TU 全中。这是**一个发码层的单点缺口**（"头里按值用到的 UDT 必须把定义
提到前向声明处，或按依赖排序/统一放到类型枢纽头"），不是 539 个独立问题。

### 32.4 RTL 缺符号（LNK2019/LNK2001）

**本次 0 条，因为根本没到 link**（cl `exit 2`）。cli 阶段 RTL 侧自身是干净的：RTL 的 TU 里只有
`vb6_di_stubs.c` / `vb6_di_win32_stubs.c` 出了诊断，共 **2 条 C4996 + 1 条 C4113**，**0 error**。
`LNK2019/LNK2001` 的函数族、导出表、入口点（`Sub Main`/`WinMain`）、manifest、依赖库
**均为未观测项**，要等 G1–G6 降到 0 之后才有读数。不许把"这次没红"当成"RTL 齐了"。

### 32.5 本机不可行，只能走 CI

- 一个 cl 调用（`/MP` 并行）在 16 GB runner 上墙钟 **6,596 s**、进程树峰值 **8,861 MB**、
  空闲物理内存最低只剩 **4,858 MB**。也就是说：**光是 cl 这一段就吃掉了半个机器**。
- 按 TU 摊：6,596 / 539 = **12.2 s/TU（还是并行之后）**。
- 本机硬规则要求 `CMAKE_BUILD_PARALLEL_LEVEL=2`、起构建前空闲 ≥ 8 GB。把 `/MP` 压到 2 换内存，
  墙钟至少 3 小时量级，且 emit 段自己就要 8.4 GB 峰值（§30）⇒ **本机串行整编在这台 16 GB 机器上不可行**，
  与 §0.1 的事故同一量级。整编**继续只走 CI**。
- `C_FILES=0 / OBJ_FILES=0` 不是失败：`--output-dir` 只落 exe 与 `c3-error.log`，`.c`/`.obj`
  在 `%TEMP%\C3C\<pid>_<n>` 且结束即清理。要看 C 就另开一次 `--emit-c`，别看整编产物。

### 32.6 探针自身的两个缺陷（记账，别误读）

1. workflow 里 `Classify` 步骤读的是 `pd_build.err`（C3 的 stderr，只有 7,760 行摘要），
   而 cl 的全量诊断在 `pd_full/c3-error.log` ⇒ 首轮 artifact 的 `pd_build.codes.txt` 是
   `CL_DIAG=0 / LINK_DIAG=0`（空表）。本节所有表格是**本地对 `c3-error.log` 重算**的结果。
   下次复跑前先把归类改成读 `c3-error.log`，否则还会拿到空表。
2. `continue-on-error: true` 让 job 显示 `success`。**看结论要看 `BUILD_EXIT` / `EXE_EXISTS`**，
   不要看 job 颜色。

### 32.7 下一步的顺序（按收益排序）

1. **G1 单点修**（预计一次能把 error 从 ~20k 降到 ~4k、TU 覆盖率从 539 降到几十）：
   让"头里按值出现的 UDT"拿到完整 `struct` 体（类型枢纽头或按依赖拓扑排序），修完重跑整编，
   同时盯 §30 的 629 形状门（这一刀**会**改发码形状，属预期改动，需重新 bless + 说明）。
2. G3（`pdImage` 方法当字段）与 G4（实参个数）是第二批：两者都指向"接收者/签名解析"，
   与 §28 的 `inferClassTypeOfExpr` 权威线同源，应合成一刀做形状门。
3. G2/G5/G6 再各自成刀。
4. 只有到了 link，才有 `LNK2019/LNK2001` 与入口/manifest 的第一份真实读数。

## 33. G1（头文件自足）落地 + 第二份整编读数（2026-11-10）

### 33.1 根因 / 修法 / 风险面

* **根因**：每个模块头里对"用到但不在本模块定义"的 UDT 只写 `typedef struct vb6_type_X vb6_type_X;`
  （[cgen_base_generate_epilogue.inc:78-85](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/base/cgen_base_generate_epilogue.inc#L78-L85)），
  而 `struct` 体只在**定义方**的头里；头顺序由 vbp 顺序决定 ⇒ 消费方一律 C2079（16,152 条 / 539 TU 全中，§32.3）。
* **修法**（不是"调 vbp 顺序"）：发码改成**两阶段**。阶段 1 把每个模块的 `.h/.c` 只暂存不落盘，
  同时收本模块落进 `.h` 的每个 UDT 定义块（`udtDefBlocks_`，原样文本，不重渲染 —— 重渲染会丢定长串长度与
  多维数组维度）；阶段 2 汇总成工程级表（同名取先到者，与 `#ifndef` 守卫同口径），
  再**按记号依赖序**把完整定义块（原样）插到每个 `.h` 的 `#include "vb6rtl.h"` 之后。
  拼接后的头在**自己那一份**里就完整，与谁先被包含无关。
* **风险面**：(a) 发码形状改变（629 门必须重登记）；(b) 头体积增加（本轮实测 +151,029 B / 629 输入）；
  (c) 块用 `#ifndef VB6_TYPE_<Name>_DEFINED` 守卫，重复展开安全，但**同名不同体**时先到者胜
  （与 C 的守卫语义一致，已记在代码注释里）；(d) 注入点固定在 `vb6rtl.h` 之后、跨模块 `#include` 之前，
  不参与 `#define` 修复块的顺序。
* **提交**：`a30c87ec`（6 文件，+107/−11）。

### 33.2 629 形状门：75 行差异逐条归因

| 口径 | 读数 |
| --- | --- |
| 本机清单 vs 期望 | 553 同 / **76 变**（其中 1 行 = `TeeChartChartDisplay/DisplayProject.vbp`，属 §29.3 既有工具链编码行） |
| **CI 清单** vs 期望（判据用） | 554 同 / **75 变**（那 1 行工具链差异在 CI 上不存在 ⇒ 差异收窄到 75） |
| 新增 / 缺席 | 0 / 0；`rc` 变化 0 |

**逐条分类的方法**（不是抽样）：对 75 行**每一行**取 `tests/photodemon` 以外的那份输入，
用两台编译器（`.temp/C3.g0.exe` = `a30c87ec` 之前；`.temp/C3.g1.exe` = 本刀）各 `--emit-c` 一次，
逐份 `git diff --no-index -U0`，并把新增行映射回新文件的**行号**，看它是否落在
`#ifndef VB6_TYPE_<X>_DEFINED` … 匹配 `#endif` 的区间里。结果：

| 项 | 合计 |
| --- | --- |
| 删除行 | **0**（75 份全部） |
| 新增行 | **4,528** |
| 其中 `#ifndef VB6_TYPE_*_DEFINED` 块起始行 | **418** 个块 |
| 其中落在守卫块**外**的非空行 | **2** 行 = `typedef struct vb6_cls_InhBase vb6_cls_InhBase;`（`tests/cls_inh/Inh.vbp` 与 `InhMain.bas` 各 1） |
| 其余新增行 | 块后空行（分隔） |

那 2 行的来路：这两个模块的 UDT 成员是**类指针**（`SibWrap.h` / `TWrap.h` 里的 `vb6_cls_InhBase* h;`），
捕获到的块文本带了一句类前向声明前导；它在模块头里本来也有一次，重复 typedef 在 C 里合法。
**没有任何一行是"守卫块以外的其它代码变了"** —— 这就是"逐条可解释"。

### 33.3 bless 与 CI 复算

* 登记吃的是 **CI 那台自己交的清单**（`ci/emit-manifest` 提交 `072089af`，注记 `head=a30c87ec`），
  不是本机清单 —— 否则那 1 行工具链差异会被写成本机值，回到 CI 又红。
* `rebless_emit_manifest.ps1 -Manifest <CI 清单> -Bless`：改写 **75** 行 / 新增 0 / 退场 0；
  末尾追加一条登记来源注记（注记 384 → 385 行，头 37 + 尾 348）；复算 **629/629 逐行相同**，
  `check_manifest_coverage` K1..K3 绿（inputs=629、registered=629、集合相等）。
* 提交 `d9040222`，已推 `github/dev` 与 `origin/dev`。CI run **42776279675**：
  `Emit manifest (shape oracle): success`（绿），红的 4 个 job（`compile #1/#2`、`vbp #5/#9/#11`）
  与上一轮 `42771669750` **完全相同**（那轮是这 5 个红，减去形状门）⇒ 本刀**零新增**红。

### 33.4 整编探针第二份读数（G1 之后）

同一命令、同一 `C3_TIMING=1` 口径（run **42771735403**，head = `ffe610e4`，`BUILD_EXIT=1`，**仍无 exe**）：

| 项 | 首轮 §32 | G1 后 | 变化 |
| --- | --- | --- | --- |
| `C2079` | 16,152 条 / **539 TU** | **0 条 / 0 TU** | **−16,152** |
| 有 error 的 TU | 539 / 539 | **287 / 539** | −252 |
| error 行合计 | 20,474 | **4,363** | −16,111（−78.7%） |
| warning 行合计 | 46,358 | 49,351 | +2,993 |
| `C1003`（每 TU 100 错误上限） | 24 TU | **15 TU** | **新撞上限的 TU = 0**；反而 9 个不再撞（`FormAnimSpeed.c`…） |
| 整编墙钟 | 6,596.2 s | **6,873.1 s** | +276.9 s（+4.2%，头变大后 cl 多读一点） |
| 峰值整树 WS | 8,861 MB | 8,891 MB | +30 MB |
| 最小空闲物理 | 4,858 MB | 4,805 MB | −53 MB |
| `c3-error.log` | 18,784,181 B / 85,020 行 | 15,307,520 B / 72,264 行 | 体积 −18.5% |
| link | 未进入（`LINK_DIAG=0`） | 未进入 | 无 `LNK2019/LNK2001` 读数 |

**其它错误码有没有被这一刀带偏**：24 个 TU 至少在一轮里撞过上限（首轮 24、本轮 15，本轮 15 个**全部**是首轮也撞过的），
直接比总数会被截断污染。所以另算一份**只含"两个 run 都没撞 C1003"的 515 个 TU** 的可比读数：

| 码 | 首轮 | G1 后 | Δ | 归因 |
| --- | --- | --- | --- | --- |
| C2079 | 15,426 | **0** | −15,426 | 本刀目标 |
| C2224 | 547 | 355 | **−192** | 原先有些是"不完整类型上的 `.x`"（G1 的下游），类型补全后消失 |
| C2440 | 234 | 184 | **−50** | 同上（不完整类型参与转换） |
| C2039 | 592 | 789 | **+197** | 原先被"类型不完整"挡住的成员查找现在真的走到了 ⇒ **暴露**，不是新增缺陷（G3 的正主） |
| C2198 | 316 | 383 | **+67** | 同上（调用形参在错误类型上被跳过） |
| 其余 18 个码 | — | — | **0** | C2036/C2045/C2063/C2064/C2065/C2069/C2095/C2102/C2106/C2112/C2113/C2172/C2186/C2197/C2223/C2231/C2296/C2297 |

全工程总数（含被截断的 TU）：`C2039` 971→1,333、`C2197` 774→887、`C2224` 1,125→668、
`C2198` 549→643、`C2440` 329→265、`C2065` 278→279。
**C1003 的方向是对的**：截断变少 ⇒ 可观测面更完整，而总数仍降 78.7%。

### 33.5 探针自身的两处修正 + Verdict 真的会红

1. `Classify` 改读全量日志 `.build/pd_full/c3-error.log`（72,264 行）并附 `pd_build.err`：
   本轮 artifact 的 `CL_DIAG=**53,714**`（首轮是 0）、`UNRESOLVED=0`、链路各码齐全。
2. **Verdict 步骤的负向验证（真跑，不是打一行字）**：临时在探针分支放
   `.github/workflows/pd_verdict_selfcheck.yml`，伪造 `.build/pd_build.meta` 里 `EXE_EXISTS=False`，
   把 Verdict 那段原样跑一遍 ⇒ run **42776484156** 结论 **failure**，红在
   `Verdict (exe must exist)` 这一步（`Fabricate` 绿）。证据取到后该文件已删除。
   同一形状在真实整编 run 上也复现：`42771735403` 的 job 结论 = **failure @ Verdict**（无 exe）。
3. 新发现的探针缺陷（下一轮生效）：`Classify` 的**受影响 TU 列恒 0** —— 首版正则要求"行首即文件名"，
   而 cl 给的是绝对路径（`D:\...\C3C\xxx/clsMenuImage.c(634)`）。已改成
   `^(?<f>.+?\.c)\(\d+\)\s*:` 再取 `GetFileName`（探针分支 `c2ac5701`，**未推**，
   随下一刀的探针 run 一起生效；本轮 §33.4 的 TU 数是本机对同一份 `c3-error.log` 重算的）。

### 33.6 这一刀之后的位置

* G1 归零：`C2079` **0**。整工程从"539 TU 全中、20,474 条 error"降到"287 TU 有 error、4,363 条"。
* 剩余按量：`C2039` 1,333（95 TU，其中 907 条集中在 `vb6_cls_pdImage`）= **G3**；
  `C2197+C2198` 1,530（**G4**）；`C2224+C2296+C2231+C2036` 762（**G2**）；
  `C2065` 279（100 TU，**G6**）；`C2440` 265（G5）。
* 仍**没到 link**：`LINK_DIAG=0`，`EXE_EXISTS=False`。`LNK2019/LNK2001` 仍是未观测项。
* 资源：峰值 8,891 MB、最小空闲 4,805 MB（runner 16,379 MB），比 §0.1 的 8 GB 红线低一半 ⇒ **整编继续只走 CI**。
* 下一刀 = **G3**（先给根因 + 影响面 + 修法 + 风险面，再动手；独立提交、独立报前后读数）。

## 34. G1b（UDT 同名不同体）落地 + 整编读数（2026-11-10）

### 34.1 根因 / 修法 / 风险面

* **根因**（G1 自己引入的回归）：G1 的工程级 UDT 表只按**类型名**汇总、同名**先到者**胜；阶段 2 对每个模块头里
  出现的**所有** `vb6_type_X` 记号注入先到块。模块私有 Type（对象模块（类/窗体）里的 `Type`、标准模块里显式
  `Private Type`）因此被当成全局类型 ⇒ 别人的字段体被注到 `vb6rtl.h` 之后、抢在本模块本地定义之前，
  本地定义又被 `#ifndef VB6_TYPE_X_DEFINED` 守卫跳过 ⇒ 本模块用了错的结构体。PhotoDemon 实测 **215 条
  C2039**，8 个结构体（分项见 §34.4），与 G3 的 pdImage 那批是两回事。
* **修法**（按 (模块, 可见性) 归档，不按名字先到者）：
  - `CCodeGen` 记录 `UdtDefRecord{moduleLower, typeName, isPrivate, block}`；`isPrivate = 显式 Private
    || isClassModule_ || isFormModule_`（`cgen_api.inc` / `cgen_state.inc` / `cgen_decl.cpp`）。
  - driver 侧两张表：`projectUdtDefs`（键 `模块名\x1f类型名`，**含**私有）、`projectPublicUdtDefs`
    （键类型名，**只** Public）。
  - 阶段 2 注入：**本模块有同名定义 → 注入自己的块**（保持 G1 的"头自足"，内容不再串类）；否则找 Public
    定义；只有别模块的私有定义 → **留前向 typedef、不静默换体**。UDT 依赖解析同样先同模块、再 Public。
  - 提交 `488f36fb`（4 文件）。
* **风险面**：(a) 发码形状改变 —— 629 门 **4 行**（§34.3）；(b) 实现继承布局里"基类私有 UDT 被派生类按值使用"
  不再重复注入派生头，改由派生头本来就有的 `#include "InhBase.h"` 供体（`tests/cls_inh` 编译用例绿，见 §34.3）；
  (c) 跨模块同名 **Public** Type 多份不同体仍取先到者 —— 那是 VB6 工程级唯一名约束之外的输入，不在本刀范围。

### 34.2 小语料探针（本机，7 模块：3 份 MergedEntry / 2 族 BITMAPINFOHEADER / 1 组 Public+私有依赖 / 1 消费模块）

语料 `.temp/g1b_probe/`。同一份输入分别喂 G1 二进制（`.temp/C3.g1.exe`）与 CI 交回的 G1b 二进制
（`c3-exe` artifact，sha256 `350f3c56…`）：

| 口径 | G1 二进制 | G1b 二进制 |
| --- | --- | --- |
| cl 诊断 | **C2039 ×11**（MergedEntry 7 = `hstColor/hstCount/v/n`；BITMAPINFOHEADER 4 = `Size/Width`）+ 级联 C2198 1 + C2440 1 | **0** |
| 产物 | 无（cl 失败） | `G1bProbe.exe` 554,496 B |
| `G1bHash1.h` | 被注 `hstColor/hstCount` 体，自己 `sKey/sItem` 被守卫跳过 | 只有 `sKey/sItem` |
| `G1bHash2.h` | 同上（错体） | 只有 `hstColor/hstCount` |
| `G1bHash3.h` | 同上（错体） | 只有 `v/n` |
| `G1bDib1.h` / `G1bDib2.h` | 两族互相顶替 | `biSize/biWidth` 与 `Size/Width` 各归各 |
| `G1bMain.h` | 注入 `G1bRec` | 注入 `G1bRec` **+ 它按值依赖的私有 `G1bInner`**（依赖解析先同模块） |

判读：回归在最小语料上 100% 复现，修法是**按模块取自己的块**，不是"少注入"。

### 34.3 629 形状门：4 行差异逐条归因 + bless

| 口径 | 读数 |
| --- | --- |
| CI 清单 vs 期望（`ci/emit-manifest`，commit 注记 `head=488f36fb`） | 629 行 / 相同 **625** / 哈希不同 **4** / 缺席 0 / 多出 0 |
| 逐行 diff 方式 | 同一输入分别喂 `.temp/C3.g1.exe` 与新二进制 `--emit-c`，字节级 + 大小写敏感行 diff；本机复算哈希与 CI 期望/实得**逐条相同** |

| # | 输入 | 差异 | 归因（为什么必须变） |
| --- | --- | --- | --- |
| 1 | `tests/Charts 2020/Proyecto1.vbp` | 10 行（2 个 `tSerie` 块） | 三个 `.ctl` 各有一份字段不同的 `Private Type tSerie`；本刀让每个控件头取**自己的**（`Labels/IconsFonts/keys/Values/LegendRect As RectF/Prop`），不再是先到者那份（含 `PT`、`LegendRect As RectL`） |
| 2 | `tests/Charts 2020/ucTreeMaps/Proyecto1.vbp` | 1 行 | `GDIPlusStartupInput.DebugEventCallback` 按本模块 `As Long` 发 `int32_t`（原为外来定义的 `intptr_t`） |
| 3 | `tests/cls_inh/Inh.vbp` | 24 行（删 3 个 `TPoint` 块） | `InhBase` 的 `Private Type TPoint` 不再强行注入 `InhMid/InhDerived/InhSib` 三个派生头；三份头本来就有 `#include "InhBase.h"`，类型仍完整可见（该用例 CI 编译/运行绿） |
| 4 | `tests/ucTreeView/Proyecto1.vbp` | 1 字节 | `POINTAPI` 取本模块字段 `X`（原为外来定义的 `x`）；字节数不变，所以只能用大小写敏感比较看出来 |

* `rebless_emit_manifest.ps1 -Bless`：改写 **4** 行 / 新增 0 / 退场 0；复算 **629/629 逐行相同**，
  `check_manifest_coverage` K1..K3 绿。提交 `95d2524d`。
* dev CI run `42792610693`：唯一新增红是 `Emit manifest`（登记前，预期）；`Tests (compile #1/#2)`、
  `vbp #11` 的红与上一轮 `42776279675` **同一组**（§31.3 那 8 条既有 STATIC + 既有 vbp 红），
  `cls_inh` / `ucTreeView` / `Charts 2020` 用例均通过。
* **登记后** dev CI run `42809450669`：`Emit manifest (Shape oracle)` **success**（629/629 逐行相同）；
  失败 job 只有 `Tests (compile #1/#2)` + `Tests (vbp #11)`，与 §31.3 的既有红**同一组**，本刀没添新红。

### 34.4 215 条按结构体分项（G1 后基线，可比子集口径见下表）

| 结构体 | C2039 | 结构体 | C2039 |
| --- | --- | --- | --- |
| `MergedEntry` | 62 | `MinMaxInfo` | 16 |
| `BITMAPINFOHEADER` | 61 | `PDP_NODE` | 6 |
| `PDP_HEADER` | 39 | `BITMAPINFO` | 4 |
| `ButtonEntry` | 26 | `WIN32_FIND_DATA` | 1 |
| **合计** | **215** | | |

G1 后全工程基线（`.temp/pd_g1/pd_full/c3-error.log`，run 42771735403）：`C2039=1,333 / 95 TU`、
`C2079=0`、`C1003` 截断 15 TU；逐码 `C2197 887`、`C2224 668`、`C2198 643`、`C2065 279`、`C2440 265`。

发码层同口径核对（`pd_emit_probe` artifact 的 `pd_emit.out`，逐模块头统计"同一头里同名的不同定义体"）：

| 结构体 | G1（run 42787817866）冲突头 | G1b（run 42792630463）冲突头 | 头文件数 / 不同体数（G1b） |
| --- | --- | --- | --- |
| `MergedEntry` | **2**（pdVariantHash、pdStringHash） | **0** | 3 / 3 |
| `BITMAPINFOHEADER` | **3**（pdICO、pdDIB、GDIPlus） | **0** | 6 / 2 |
| `PDP_HEADER` | **1**（pdPackagerLegacy） | **0** | 2 / 1 |
| `ButtonEntry` | **2**（pdButtonStripVertical、pdStrip） | **0** | 3 / 2 |
| `MinMaxInfo` | **2**（pdWindowManager、pdWindowSize） | **0** | 3 / 2 |
| `PDP_NODE` | **1**（pdPackagerLegacy） | **0** | 2 / 1 |
| `BITMAPINFO` | **2**（GDIPlus、Plugin_FreeImage） | **0** | 3 / 3 |
| `WIN32_FIND_DATA` | **1**（pdFSO） | **0** | 2 / 2 |
| **合计** | **14** | **0** | — |

"不同体数 > 1"是**跨模块**的差异（每个模块保留自己那份），"冲突头 = 0"才是本刀要的结论。

### 34.5 整编读数（G1b，CI run `42792630868`；基线 = G1 run `42771735403`）

环境口径：`BUILD_EXIT=1`（仍未到 link，**预期**）、`EXE_EXISTS=False`、`LINK_DIAG=0`、
`C3_TIMING=True`（`BUILD_SECONDS=6895.9`，与关计时的轮次不可比墙钟）、`PEAK_TREE_WS_MB=8895`、
`MIN_FREE_PHYS_MB=4398`（runner 16,379 MB）。

**215 条按结构体分项（前 → 后）**：

| 结构体 | G1 | G1b | 结构体 | G1 | G1b |
| --- | --- | --- | --- | --- | --- |
| `MergedEntry` | 62 | **0** | `MinMaxInfo` | 16 | **0** |
| `BITMAPINFOHEADER` | 61 | **0** | `PDP_NODE` | 6 | **0** |
| `PDP_HEADER` | 39 | **0** | `BITMAPINFO` | 4 | **0** |
| `ButtonEntry` | 26 | **0** | `WIN32_FIND_DATA` | 1 | **0** |
| **合计** | **215** | **0** | | | |

**全码表（未剔截断 TU，仅作总量参考）**：`C2039 1,333/95 TU → 1,118/84 TU`；`C2198 643/114 → 555/105`；
`C2440 265/72 → 257/70`；其余 20 个码逐条**不变**（`C2197 887/63`、`C2224 668/76`、`C2065 279/100`、
`C2064 61/10`、`C2045 50/17`、`C2296 39/19`、`C2036 34/4`、`C2231 21/6`、`C2063 19/19`、`C2106 10/4`、
`C2069 8/4`、`C2297 5/1`、`C2095 5/2`、`C2186 5/3`、`C2102 4/3`、`C2172 4/1`、`C2223 3/1`、`C2113 3/2`、
`C2112 2/1`）。**`C2079` 两轮都是 0**（G1 那刀没被破）。

**可比子集**（两轮都撞 `C1003` 的 TU 完全相同，共 15 个：`Filters_Transform / FormAnimBackground /
FormMain / FormMetadata / Interface / Layers / pdCanvas / pdLayerListInner / Processor / SelectionFilters /
SelectionUI / Snap / toolbar_Toolbox / toolpanel_MoveSize / toolpanel_TextAdvanced`，剔除后 **532 TU**）：

| 码 | G1 | G1b | Δ | TU |
| --- | --- | --- | --- | --- |
| `C2039` | 868 | 653 | **-215** | 82 → 71 |
| `C2198` | 408 | 320 | **-88** | 101 → 92 |
| `C2440` | 196 | 188 | **-8** | 64 → 62 |
| `C2197` / `C2224` / `C2065` / `C2064` / `C2045` / `C2296` / `C2036` / `C2231` / `C2063` / 其余 | — | — | **0** | 不变 |

`C1003`：两轮都是 **15 TU**、同一组 ⇒ **没有 TU 因这刀新撞上限**，可比子集口径成立。

**Δ 归因**（逐 TU 逐码比对）：非零 Δ 只落在 **11 个 TU**，且**全部**属于两轮分类表里那 8 个结构体所在的
模块（`pdPackageLegacyV1 45+19`、`pdICO 33+12`、`pdStringHash 32+20`、`pdVariantHash 30+19+5`、
`pdButtonStripVertical 18+9`、`pdDIB 18+3`、`GDI_Plus 14`、`pdWindowManager 12`、`pdStrip 8+4`、
`pdWindowSize 4+1`、`pdFSO 1+1+3`）。`C2198`/`C2440` 是 `C2039` 的**级联**：成员不存在 → cl 把表达式
当错误类型 → 跟着报"实参太少 / 不能转换"；成员一回来，级联同步消失。没有第 12 个 TU、没有第 4 个码变化。

结论：215 条 **降到 0**（无可消除剩余）；发码层同口径的 14 个"同头冲突"也降到 0（§34.4）；
`C2079` 保持 0；`C1003` 截断面不变；其余码零变化。

## 35. G3（E1+E2）落地 + 整编读数（2026-11-10）

范围：只补"模块限定无参函数返回值当接收者"这一条链的接收者推断（E1）与调用补全（E2）。
不碰成员方法表可见性、不碰 `getPublicSymbols` 的 Friend 导出、不碰包边界、不洗 `udtDefBlocks_`。

### 35.1 两次提交（本刀内的分叉与修法）

| 提交 | 内容 |
| --- | --- |
| dev `853a36fa` | E1：`inferClassTypeOfExpr` 的 `MemberAccessExpr` 分支在 base 类名为空时调 `findModuleQualifiedNoArgFunction`，返回类的工程类名；E2：值上下文 `funcName += "()"`；成员调用分支补 `()` |
| dev `be12a9fe` | **E2b（回归修复）**：`() ` 的补齐上移到 `cgen_expr_member_class_fallback.inc` 的 `node.object` 分支入口，成员调用 / 数据字段读两条腿共用 |

`853a36fa` 整编实测 `C2223 3 → 286`（+283）：E2 的 `()` 只补在成员调用分支，**字段读那一支漏补**，
于是 `vb6_PDImages_GetActiveImage->MainSelection` 原样发出（`->` 左边是函数指针）。
E2b 修后同一形态发成 `vb6_PDImages_GetActiveImage()->MainSelection`。

### 35.2 小语料探针（本机，3 类 + 1 字段形）

三种链形（E2b 后）：

- `vb6_G3Image_GetNumOfLayers((void*)vb6_G3Images_GetActiveImage())`
- `vb6_G3Layer_GetLayerOffsetX((void*)vb6_G3Image_GetActiveLayer((void*)vb6_G3Images_GetActiveImage()))`
- `vb6_G3DIB_GetDIBColorDepth((void*)vb6_G3Image_GetActiveDIB((void*)vb6_G3Images_GetActiveImage()))`

字段形（E2b 新增覆盖）：`vb6_G3Layer_GetLayerOffsetX((void*)vb6_G3Images_GetActiveImage()->MainSel)`。
小语料修前 3×`C2039` + 3×`C2198` 无 exe；修后 0 error，出 `G3Probe.exe`（556032 B）。

### 35.3 629 形状门

本机 625 相同 / 4 哈希不同（GUISample×3、TeeChart）。用旧二进制 `.temp/g1b_probe/bin/C3.exe` 与新二进制
逐份 emit，四份**字节完全相同**（306843 / 289698 / 288429 / 83528）⇒ 环境噪声，**无需 bless**。
CI 侧 `Emit manifest (shape oracle)` 在 run `42828836909` 上 **success（629/629）**。

### 35.4 整编读数（G1b run → G3 最终版 run `42828903579`，headSha `581206ea`）

`BUILD_SECONDS=4947.6`（基线 6895.9，`C3_TIMING=True` 下墙钟不可比）、`PEAK_TREE_WS_MB=8910`、
`MIN_FREE_PHYS_MB=4876`、`BUILD_EXIT 1 → 1`、`EXE_EXISTS False → False`。

| 码 | G1b | G3 | Δ | 可比子集 Δ |
| --- | --- | --- | --- | --- |
| `C2039` | 1118 / 84 TU | **88 / 12 TU** | **-1030** | -570 |
| `C2224` | 668 / 76 | **59 / 16** | **-609** | -341 |
| `C2198` | 555 / 105 | **261 / 65** | **-294** | -172 |
| `C2197` | 887 / 63 | **762 / 49** | **-125** | -81 |
| `C2440` | 257 / 70 | **214 / 55** | **-43** | -20 |
| `C2065` | 279 / 100 | 283 / 102 | +4 | **0** |
| `C2223` | 3 / 1 | **3 / 1** | **0** | **0** |
| `C2143` | 0 | 2 / 1 | +2 | 0 |
| `C2231` | 21 / 6 | 59 / 8 | +38 | **0** |
| `C2064` | 61 / 10 | 66 / 12 | +5 | **0** |
| `C2296` | 39 / 19 | 14 / 14 | -25 | -26 |
| 全码（excl `C1003`） | 4037 | **1960** | **-2077** | — |

`C2039` 分项：pdImage `1049 → 50`，非 pdImage `69 → 38`。残余 88 条 = 24 个 (类型,成员) 对：
pdImage 50 条 / 16 个成员（`GetLayerByID 8`、`RequestThumbnail 8`、`GetSaveState 5`、`EstimateRAMUsage 4`、
`GetLayerByIndex 4`、`GetOriginalFileFormat 4`、`IsSuspended 4` …）；非 pdImage 38 条 = `pdResize 32`、
`pdSelection 2`、`UserControl_Extender 2`、`pdCanvas 1`、`pdStatusBar 1`。

**`C2223` 已完全回落**：286 → 3，且这 3 条与基线**同一 TU、同一行**（`PDImages.c` 364 / 442 / 450，
`'->imageID'`），是 G1b 就有的旧账，不是本刀引入。

**总量上涨的三项全是截断面变化，不是新错误**：`C2231 +38` 里 +32 全在 `SelectionUI.c`（基线该 TU 被
`C1003` 截断，G3 后不再截断 ⇒ 原本被吞的诊断现形）；`C2065 +4`、`C2064 +5` 同理，可比子集 Δ 均为 0。

### 35.5 `C1003` 截断 TU 组

`15 TU → 3 TU`（`FormMain.c`、`FormMetadata.c`、`toolbar_Toolbox.c`），**新增 0**，消失 12。
中间版 `853a36fa` 曾新撞 1 个 `toolpanel_Selections.c`（18 条 `C2223` 把原有 ~70 条推过 100 条上限），
E2b 后该 TU 退出截断组 ⇒ **本刀没有 TU 因接收者写法变长而新撞上限**。

### 35.6 50 条 pdImage `C2039` 的逐成员核对

291 个 (成员, TU) 对：G3 前 27 对同 TU 已有正确 `vb6_pdImage_<member>(` 发码；G3 后 272 对。
剩 19 对**全部**是 `PDImages.GetImageByID(<实参>).Member` —— 带参模块函数当接收者，属 G3b，本刀不碰。
可见性缺口 = 0（50 条对应的成员在修前发码里都有定义）。

### 35.7 单列残余（不属本刀机理）

`Layers.bas:467-468` 的 `X.MainSelection.GetCompositeBoundaryRect.Left`：无参方法返回 UDT 再读字段，
现在发出 `vb6_pdLayer_SetLayerOffsetX(..., vb6_pdSelection_GetCompositeBoundaryRect((void*)vb6_PDImages_GetActiveImage()->MainSelection->Left);`
⇒ 2×`C2143`（缺 2 个右括号，`Layers.c` 672/673）。基线同一行发成 `...GetActiveImage.GetLayerByID(...).SetLayerOffsetX(...)`
⇒ `C2039`，即这两行本来就坏，只是错误码换了。根因在"无括号调用语句实参 + UDT 字段标记"的 `pendingChainObj_` 协议。

`dev` CI run `42828836909`：形状门 success；`Tests (compile #1/#2)` 与 `Tests (vbp #11)` 失败，**与
`c2b1e52c` / `853a36fa` 两轮完全同一组**，属既有红，不由本刀引入。


## 36. G4 / M1 收口：控件数组接收者的类感知形参表（2026-11-11）

- 提交：`b67f45ab`（`dev`，6 files changed / 262+ / 22-）；哨兵 `scripts/check_g4_receiver_class_gate.ps1`。
- 探针分支：`b67f45ab` cherry-pick 为 `a9715867`，推 `github` + `origin` 的 `probe/photodemon-emit`。
- 范围：本刀**只做 M1**。M2（菜单/控件事件处理器漏传 Index）、M3（内置/RTL 不补缺省）、
  以及「UDT 内嵌数组字段被当函数调用」（本刀新归出）**全部不进本刀**，见 36.7。

### 36.1 根因（M1）

形参表决策的**接收者类解析**只有 `inferClassTypeOfExpr` 一条路，它不认**控件数组元素**接收者
（AST 为 `IndexOrCallExpr`，如 `ddWhichFrame(1).AddItem ...`、`ctlMany(0).AddItem(...)`）。
解析答不出 ⇒ 落回「全表按成员名兜底」`symTab_.lookupModule(成员名)`，命中**跨类同名成员**的形参表：

| 成员 | 真实类（`pdDropDown`） | 兜底命中的类（`pdListSupport`） | 后果 |
| --- | --- | --- | --- |
| `AddItem` | 4 参（self + 4 = 5） | 5 参 | 多发 1 枚 Optional 缺省 + 1 枚 `_has_` 尾标记 |
| `AssignTooltip` | 4 参 | 6 参 | 多发 2 枚 |

即发码形状从「self + 4 参 + 4 `_has_`」变成「self + 5 参 + 5 `_has_`」，11 枚实参打 9 参原型 ⇒ `C2197`。
**注意：这不是"少传 / 多传用户实参"，而是补缺省那一层按错的形参表补齐。** 所以 M1 属"接收者推断带出来的级联"，
不是"真缺参数"。

### 36.2 两道闸 + 可检查形态（用户硬要求）

采纳点收敛为**唯一**成员函数 `CCodeGen::tryAdoptReceiverClassParams`（`src/backend/cgen_util_classtype.cpp`）：

1. `inferCallReceiverClassForParams` 先试 `inferClassTypeOfExpr`；答不出再看接收者是否是
   `IdentifierExpr` / `IndexOrCallExpr` 的 callee，查 `knownUserControlCtrlVars_`（控件数组变量 → UserControl 类）；
2. 类解析仍为空 ⇒ `return false`；
3. `findClassMemberCallParams` 未命中 ⇒ `return false`；只有 `true` 才写 `calleeParams`，
   否则**逐字回落**原 `lookupModule` 兜底（三处兜底代码一行未动）。

可检查形态 = `scripts/check_g4_receiver_class_gate.ps1`，5 条规则：

| 规则 | 断言 |
| --- | --- |
| A1 | `tryAdoptReceiverClassParams` 定义在权威文件恰 1 处、全 `src` 恰 1 处 |
| A2 | 函数体内 `inferCallReceiverClassForParams` 恰 1 次；`if (<x>.empty()) return false;` 出现在 `findClassMemberCallParams` **之前**；调用形如 `if (!findClassMemberCallParams(`；有 `return true;` |
| A3 | 调用点恰 2 处，均为 `if (tryAdoptReceiverClassParams(` 守卫形态，且 400 字符内出现 `calleeParams = std::move(` |
| A4 | 两处 class-unaware 兜底仍在（`if (!byClassCS254) { … lookupModule(maExpr.memberName) … }` 与 `callee_params.inc` 的 `lookupModule`） |
| A5 | `C3-SENTINEL(g4-m1-gate)` 在权威文件恰 1 处 |

已接入 `tests/run_tests.ps1` 静态分片（`Test-G4ReceiverClassGate`）；`check_static_sentinel_registration.ps1`
复验 `127 sentinels on disk, all registered`。**负测试**：删掉 `if (recvCls.empty()) return false;` ⇒ 哨兵判红；
还原 ⇒ PASS。即这道闸不靠口径说明，规则被改坏会立刻变红。

### 36.3 小语料探针（4 模块，`.temp/g4_probe/M1/`）

`G4Probe.vbp`：`G4Uc.ctl`（真接收者，`AddItem` 4 参）+ `G4Decoy.cls`（诱饵，`AddItem` 5 参）+ `G4Form.frm`
（`ctlOne` 单枚、`ctlMany` 控件数组 Index 0/1），四种形态覆盖两条 gate 调用点。

| 形态 | 修前实参数 | 修后 | 原型 |
| --- | --- | --- | --- |
| `ctlMany(1).AddItem "beta", 1`（无括号） | 11 | **9** | 9 |
| `ctlMany(0).AddItem("delta", 3)`（带括号） | 11 | **9** | 9 |
| `ctlMany(1).TagIt`（裸 MemberAccess、无括号无参） | 7 | **3** | 3 |
| `ctlOne.AddItem "alpha", 0`（IdentifierExpr 接收者） | 9 | 9 | 9（未受影响） |

修前 `cl /we4020` 3 处 `C4020`（本机 14.29 对同构造的码，CI 版是 `C2197`）；修后 **0 error / `CL_EXIT=0`**。

### 36.4 629 形状门（逐条归因，该 bless 才 bless）

`scripts/emit_manifest.ps1` 跑 629 个输入，修前（`C3_prefix.exe` sha256 `b89320be…`）与修后
（`C3_after.exe` sha256 `90c7b1d1…`）各一份：

- `629` 条数据行**逐行相同**，哈希不同 **0**，`rc=` 不同 **0** ⇒ **本刀不改任何既有语料的发码形状**。
- 两份清单**只有第 1 行 `# c3-exe sha256=` 不同**（二进制本身），`toolset` / `inputs=629` 行一致。
- ⇒ **不 bless，`emit-manifest.expected.txt` 未动。**

与 checked-in 期望相比有 4 份 `.vbp` 不一致（`tests/guisample/GUISample.vbp`、两个
`vb6ref/.../GUISample.vbp`、`tests/TeeChartChartDisplay/DisplayProject.vbp`）：这 4 份在**修前二进制上同样不一致**，
修前修后是同一组 ⇒ 本机 vs CI 环境噪声，不是本刀引入，不静默、不 bless，只在此记账。

### 36.5 整编读数（CI run `42856903368`，post-M1；基线 = `run 42812594786` / G3 收口版）

| 诊断 | 修前 总量 / TU | 修后 总量 / TU | Δ总量 | 可比子集 Δ |
| --- | --- | --- | --- | --- |
| `C2197` | 762 / 49 | **219 / 21** | **-543** | **-411** |
| `C2198` | 261 / 65 | 261 / 65 | 0 | 0 |
| `C2065` | 283 / 102 | 283 / 102 | **0** | 0 |
| `C2440` | 214 / 55 | 214 / 55 | **0** | 0 |
| `C2039` | 88 / 12 | 88 / 12 | 0 | 0 |
| `C2223` | 3 / 1 | 3 / 1 | 0 | 0 |
| `C2064` | 66 / 12 | 66 / 12 | 0 | 0 |
| `C2231` | 59 / 8 | 59 / 8 | 0 | 0 |
| `C2296` | 14 / 14 | 14 / 14 | 0 | 0 |
| 全码（excl `C1003`） | 1960 | **1417** | **-543** | — |

- **`C2065` / `C2440` 没有被顺带推高**：总量、TU 数、可比子集三项全为 0 变化。
- **`C1003` 截断组：3 TU → 1 TU**（`FormMain.c`、`FormMetadata.c`、`toolbar_Toolbox.c` → `FormMain.c`），
  **新增 0**，消失 2 ⇒ 本刀没有让任何 TU 新撞上限，反而把两个 TU 拉回可测区间。
- 可比子集（两轮都未撞 `C1003` 的 242 个 TU）：`C2197` `566 → 155`，其余 22 个码**全部 Δ=0**。
- `BUILD_EXIT=1`、`EXE_EXISTS=False`（两轮一致，工程仍未整体编过；本刀只收诊断，不宣称可执行）。

### 36.6 逐 TU / 逐成员归因

`C2197` 按 TU：**上升的 TU 数 = 0**；下降 30 个 TU，合计 -543。前三名
`toolbar_Toolbox.c 100→0`、`FormRechannel.c 70→0`、`toolpanel_Selections.c 50→0`。

按被调方：**12 个族归零**，全部是 M1 命中的控件接收者族 ——
`pdDropDown_AddItem 140→0`、`pdButtonToolbox_AssignTooltip 127→0`、`pdButtonStrip_AddItem 126→0`、
`pdButtonToolbox_AssignImage 66→0`、`pdDropDown_AssignTooltip 21→0`、`pdCheckBox_AssignTooltip 18→0`、
`pdButton_AssignTooltip 18→0`、`pdTitle_AssignTooltip 18→0`、`pdRadioButton_AssignTooltip 9→0`、
`pdContainer_SetPosition 8→0`、`pdButtonStrip_AssignTooltip 6→0`、`pdSlider_AssignTooltip 6→0`。

**唯一按函数看是涨的项要单独说清**：`vb6_pdStringStack_GetString 48→68`（+20）。它**不构成回归**：
20 条全部落在 `FormMetadata.c`，而该 TU 在基线被 `C1003` 截断（96 条封顶）、本刀后截断解除（64 条），
原本被吞的诊断现形；同一 TU 总量仍是 `96 → 64（-32）`。这正是"按 TU 看无上升项"的口径依据。

剩余 219 条 `C2197` 的归属：`pdStringStack_GetString 68`（截断解除现形）、`Num_Div 24`、`RtlMoveMemory 22`、
其余 105 条为单点/低频族（`UpdateAgainstCurrentTheme` 一族、`Tools_*_NotifyMouse*`、`Plugin_*_GetVersion` …），
本刀未触及。

### 36.7 单列项（不进本刀读数）

**A) 46 条待确认桶（已在最终 emit 工件上钉死，单独汇报过，此处只记账）**

- 工件：CI run `42856817850` 的 `pd_emit.out`（1078 段，每模块 `.h`+`.c`，`.c` 段行号与整编诊断逐条对齐）。
- `vb6_Num_Div` 24 条 = **真·实参溢出**，4 个调用点 × 6 枚多余实参
  （`pdSliderStandalone.c` 1394/1408、`Adjustments_Color_HSL.c` 925/927）。触发形态：
  **无括号语句式调用 + 首实参是 `/` 除法 + 被调方带 ByRef 出参**，真 callee 被塞成 `Num_Div` 的第 1 实参、
  其余实参漏到 `Num_Div` 实参表尾。→ 归 **M3 之后的「无括号调用 + 二元表达式实参」刀**。
- `vb6_lwa_RtlMoveMemory` 22 条 = **C2064 的级联，不是 arity 问题**。`CopyMemoryStrict` 全工程 504 个点
  全 3 实参、零失配；这 22 行**每行都同时带一条 `C2064`**（22/22 落在全工程 66 条 `C2064` 的 (file,line) 集合内），
  现场是 **UDT 内嵌数组字段被当函数调用发码**（`.pixelData(0, 0)`、`(*srcChannel).cBytes(...)`）。
  → 归 **「UDT 内嵌数组字段被当函数调用」刀**（本刀新归出）。
- 工件级前后一致性：`pdGIF.c`(1929 行) / `pdXCF.c`(3540) / `pdSliderStandalone.c`(1984) /
  `Adjustments_Color_HSL.c`(960) 四个模块的 `.c` 段**逐字节 SAME**；对照组 `toolpanel_Gradient.c` **DIFF**
  （M1 命中），`vb6_pdDropDown_AddItem` 的 11 实参点 **70 → 0**。

**B) M2 / M3**（范围批准时即约定单列，本刀不碰）：M2 = 菜单/控件事件处理器漏传 `Index`；
M3 = 内置/RTL 不补缺省。二者构成 `C2198 261 / 65 TU` 的主体：
`FormMain_Mnu*_Click` 一族 ~86 条（15+14+10+9+8+6+5+5+4+3+2+2+1+1+1）、
`pdResize_ConvertUnitToPixels 12`、`pdOpenSaveDialog_Get*FileName 10`、`IntDiv 7`、`Randomize 7`、`MonthName 7`。

### 36.8 下一刀排序（按本刀读数重排收益）

1. **M2 菜单/控件事件处理器漏传 `Index`** —— `C2198` 单族最大（`FormMain_Mnu*_Click` ~86 条 / 14 TU），
   机理单一（事件 raiser 补 `Index`），改动面小。
2. **「UDT 内嵌数组字段被当函数调用」** —— 一次收 `C2064 66 + C2197 22`，且是**错误码混用**（本该是取址/下标），
   不修则会持续污染 `C2197` 的读数口径。
3. **M3 内置/RTL 补缺省** —— `C2198` 剩下的散点族（`IntDiv/Randomize/MonthName/CSng/Str/…`）+ `Num_Div` 24 条。
4. 之后再看 `C2039` 残余 88 条（G3b：带参模块函数当接收者）与 `C2065 283`。

---

## §37 G4/M2：事件转发站按元素 Index 传参（C2198 −128；菜单 19 族 + Slider 11 族归零）

**提交**：`0e6d7719`（src + 哨兵 + 测试登记，dev / `github/dev`）、`dc507a19`（探针分支 cherry-pick）、
`595cc729`（形状门登记 1 行）。**整编验收**：CI run `42884417424`（`PhotoDemon x86 full build`，探针分支）。

### 37.1 机理（两处转发站，一个病根）

事件转发站把调用写死成 `{ extern void fn(); fn(); }` —— `void fn()` 是**无原型声明**，
而控件数组 / 菜单数组的处理器首形参是元素下标 ⇒ 0 实参打 1 形参 = **C2198**。
本刀把这两站收进唯一采纳点 `CCodeGen::eventIndexCallText`（`cgen_form_menu.cpp`）：

| 条件 | 发码 |
|---|---|
| handler 为空 / 无第一形参 / `elemIndex < 0`（非数组成员） | **逐字**原 0 参文本 `{ extern void fn(); fn(); }` |
| 第一形参 ByVal | `fn(N)` |
| 第一形参 ByRef | `T v = N; fn(&v)`，`T = mapType(p0.type)`（不写死 int16_t） |

- **菜单**（`emitMenuClickDispatch` 两处叶结点）：下标必须是 **`child.index` / `menuCtrl.index`**。
  `menuId` 是"第几个叶菜单项"的**顺序计数器**，与 VB6 的 `Index` 无关 —— 哨兵 B4 明令禁止把它传进来。
- **Slider 的 Change/Scroll 臂**（`cgen_form_wndproc_dispatch.inc`）：除传下标外同时补两件——
  ① 控件数组元素**没有** `vb6_hwnd_<名>`，句柄按 `vb6_CtrlArr_GetAt(&vb6_arr_<名>, idx)` 取（原为 C2065）；
  ② 工程 UserControl 名字里带 "slider" 会被 `frm_parser_util.cpp:206` 判成原生 Slider，而它的事件
  由 UC sink thunk 派发（thunk 早已实参化）⇒ 按 `findUserControlSpec(ctrl.controlTypeName)` 跳过。
  不跳过的后果是：修好 arity 之后**同一条 Change 会被派发两次**。

### 37.2 小语料（`.temp/g4m2_probe`，2 模块，五种形态全带正/负向）

| 形态 | 修前（`M2Form.before.c`） | 修后（`M2Form.after.c`） |
|---|---|---|
| 菜单数组 `mnuArr` idx 0/1，`_Click(Index)` | L61/64 0 参 ⇒ **C2198 ×2** | `int16_t vb6_evidx_0/1 = 0/1; …(&…)` |
| 非数组菜单 `mnuSolo`，`_Click()` | L67 0 参 | 逐字不变 |
| 菜单数组 `mnuZero` idx 0，`_Click()`（数组但无参） | L70 0 参 | 逐字不变 |
| 原生 Slider 数组 `sldArr` idx 0/1，`_Change(Index)` | L91-96 **两块重复** + `vb6_hwnd_sldArr` 未声明 ⇒ **C2198 ×2 + C2065 ×2** | 两块按元素各一，句柄 `CtrlArr_GetAt(…,0/1)`，实参带下标 |
| 非数组 Slider `sldOne`，`_Change()` | L98-100 0 参 | 逐字不变 |

修前 `EXIT=1`（cl 失败），修后 `EXIT=0`。负向三态的生成文本**逐字节相同**。

### 37.3 哨兵 `scripts/check_g4_event_index_forward.ps1`

B1 定义全域唯一且在权威文件；B2 标记 `C3-SENTINEL(g4-m2-event-index)` 全域恰好 1 处且在函数体内；
B3 函数体形状（回落文本逐字、`params.empty()` + `elemIndex < 0` 双闸、`isByVal` 两臂、类型取 `mapType(p0.type)`、禁 `int16_t`）；
B4 调用面计数（menu 3 次 = 1 定义 + 2 叶、wndproc 2 次、helpers 1 次声明、其余 0）+ 菜单必须 `.index` + 调用里禁 `menuId`；
B5 菜单不得再有写死 0 参；B6 Slider 臂（禁写死 0 参、必须保留 UC 跳过门、数组句柄必须走 `CtrlArr_GetAt`）。

- 已进 `tests/run_tests.ps1` 的 static 分片（`Test-G4EventIndexForward`），注册普查 `check_static_sentinel_registration.ps1` 通过。
- **负例 4/4 判红**（副本树 `-SrcRoot` 定点破坏）：`menuCtrl.index→menuId`（B4）、`elemIndex < 0 → < -1`（B3b）、
  删 UC 跳过门（B6）、数组句柄退回 `vb6_hwnd_`（B6）。

### 37.4 629 形状门

本机修前/修后各跑一遍，逐行比对：**1 行真变化**（`tests/ucSlider_2022/Proyecto1.vbp`），
另有 4 行是 §36.4 已记账的本机/CI 工具链噪声（两份清单里同样 4 条，本刀未变）。

该语料的变化**逐字节归因**（用 post-M1/pre-M2 二进制 `C3_after.exe` 与修后二进制各 `--emit-c` 对拍）：
**只少 6 行、0 增 0 改** —— 恰好 `ucSlider6` / `ucSlider4` 两块 Slider 臂（2 条 `scrollHwnd ==` +
2 条 `FireChange` 调用 + 2 个右括号）。这两个 UC 的 Change 由 sink thunk 派发，那两块是重复派发，必须消失。
数组 `ucSlider1` 的 6 个元素根本没被 Slider 臂发过（无 `_Change`/`_Scroll` 处理器，早已 `continue`）。

登记方式：**不吃本机清单**，`rebless_emit_manifest.ps1` 取 CI 自己那台的 `ci/emit-manifest`
（head `0e6d7719`）⇒ 仅 1 行改写，且 CI 哈希与本机产物哈希**相同**（`7a81869c…`）；
登记后复算 629/629 相同，dev 的形状门任务转 `success`。

### 37.5 整编读数（基线 run `42856903368` → 本刀 `42884417424`）

| 诊断 | 修前 总量/TU | 修后 总量/TU | Δ | 可比子集 Δ |
|---|---|---|---|---|
| **C2198** | 261/65 | **133/54** | **−128 / −11** | **−11**（64→53） |
| C2197 | 219/21 | 219/21 | 0 | 0 |
| C2065 | 283/102 | **228/96** | **−55 / −6** | −7（102→95） |
| C2440 | 214/55 | 214/55 | 0 | 0 |
| C2039 | 88/12 | 88/12 | 0 | 0 |
| C2064 | 66/12 | 66/12 | 0 | 0 |
| C2224 | 59/16 | 59/16 | 0 | 0 |
| 全码（excl C1003） | 1417 | **1276** | **−141** | — |

- **C1003 截断组：`FormMain.c` 1 TU → 0 TU**（新增 0）。基线的 `FormMain.c(803): fatal error C1003` 消失，
  即该 TU 从"100 条封顶"变成全部报完。
- `BUILD_EXIT=1` / `EXE_EXISTS=False` 两轮一致（工程仍有 1200+ 条错误，本刀不动这条结论）。
  `PEAK_TREE_WS_MB 8898→8892`、`MIN_FREE_PHYS_MB 4919→4851`、`BUILD_SECONDS 3558.8→6173.7`（后者是本次 runner 偏慢，非代码量）。
- **C2198 归零族：31 个族里的 30 个** —— 菜单 19 族 / 101 条（`MnuEdit 15`、`MnuImage 14`、`MnuFile 10`、
  `MnuLayerOrder 9`、`MnuLayerNew 8`、`MnuFileExport 6`、`MnuRotate/MnuFileImport/MnuLayerVisibility 5`、
  `MnuMetadata/MnuEditSpecial 4`、`MnuLayerReplace 3`、其余 8 族各 1–2）+ Slider 11 族 / 28 条
  （`FormGamma.sltGamma 3`、`FormColorize.sldHSL 3`、`FormColorHalftone.sltAngle 3`、
  `dialog.ColorSelector.sldRGB/sldHSV 3+3`、`toolpanel_Clone/Eraser/Paintbrush/Pencil`、`FormRangeFilter`、
  `FormModernArt`、`FormHarmonicMean`）。
- **唯一按函数看是涨的项**：`vb6_Menus_UpdateAgainstCurrentTheme 1→2`，新增那条落在 `FormMain.c(2909)` ——
  在该 TU 基线的截断点（803 行）**之后**，属"截断解除现形"，不是本刀引入（同一 TU 的 `hPopup_Mnu*` 声明
  在修前/修后 emit 里**逐字相同**）。
- 同理，本刀后新出现的 `C2374 42 条`（全部 `FormMain.c` 1659 行以后，`hPopup_MnuFile/MnuImage/MnuLayer…` 重定义）
  **全部是既有隐藏项**：同一个 `hPopup_*` 初始化语句在修前 emit 里一模一样，只是过去被 100 条上限截断吞掉。
- C2065 那 −55 条：基线里 56 条是 `vb6_hwnd_sld*` / `vb6_hwnd_slt*`（Slider 控件数组没有这枚句柄），
  本刀改成 `CtrlArr_GetAt` 后整族消失；`FormMain.c` 截断解除又现形 1 条，净 −55。

### 37.6 emit 工件读数（`pd_emit.out`，run `42856817850` → `42884401927`）

- `vb6_FormMain_Mnu*_Click` 的 **0 参调用点 352 → 2**；剩余 2 条 `MnuClearRecentMacros_Click` / `MnuTest_Click`
  的定义侧就是 `(void)` ⇒ 负向口径逐字保留。
- `vb6_Slider_FireChange` 块 **323 → 0**（PhotoDemon 的 slider 全是工程 UC `pdSlider`；`sldDitheringAmount`
  一类控件在 emit 里是 `vb6_UC_CreateEx + evtThunk + _sink_`，事件走 sink，不走 form 的 WM_*SCROLL 臂）。
- emit 总行数 `1,010,915 → 1,009,946`（−969）：323 块 ×3 行 + 菜单 350 处的改写。

### 37.7 单列项（本刀未碰，按批准范围）

- **UDT 内嵌数组字段被当函数调用**：`C2064 66 / 12 TU` + 连带 `C2197 22`（`RtlMoveMemory` 那一族每行都带 C2064）。
- **M3 内置/RTL 补缺省**：`pdResize_ConvertUnitToPixels 12`、`Plugin_libdeflate_* 15`、`Randomize 7`、
  `MonthName 7`、`IntDiv 7`、`pdOpenSaveDialog_GetSaveFileName 6`、`GetOpenFileName 4` …
- `Num_Div 24`（无括号调用 + 二元表达式实参）仍在原地。

### 37.8 下一刀排序（按本刀读数重排）

1. **UDT 内嵌数组字段被当函数调用** —— 一次收 `C2064 66 + C2197 22`，且是错误码混用（本该是取址/下标），
   继续留着会污染 C2197 的读数口径。
2. **M3 内置/RTL 补缺省** —— C2198 剩余 133 里最大的一块（前 7 族合计 55 条）。
3. `Num_Div` 24（无括号调用形态）+ `C2039 88`。
4. `C2065 228`：先按"截断解除现形 vs 真缺声明"两桶分开，再定刀。

### 37.9 本刀遗留的观测量

dev CI（`0e6d7719` / `595cc729`）除 3 个既有失败分片（`compile #1/#2`、`vbp #11`）外，
`Tests (vbp #5)` 掉了一条 `[PIXELS] uctreeview_px.ps1`：`'Contact  List' 内容段 BM_CLICK 3s 未返回 / 看门狗超时`。
同一份二进制在 `0e6d7719` 那一轮该分片 `PASS=48 FAIL=0`，且 `ucTreeView` 语料无菜单、无 Slider ⇒ 主机时序抖动，
与本刀无关，未单独处理。

## §38 G4 A–E：UDT 数组成员秩 / 语句首参负号 / 值成员空括号（C2064 66 → 0；C2197 −32、C2198 −10）

**提交**：`941a2afd`（A）、`d7b18ab7`（B）、`58fd516e`（C/D/E）、`206b4c7f`（三条 static 哨兵 + 三份夹具登记）、
`0337282a`（形状门登记 3 行）。**探针分支**：merge dev `0337282a` → `ac25f3aa`。
**整编验收**：CI run `42915022004`；基线 = §37 的 `42884417424`（同一探针作业、同一 x86 全量口径）。

### 38.1 机理（三处改点，都是"把不是调用的东西当调用发"）

| 刀 | 触发形态（PhotoDemon 实证） | 病根 | 改点 |
|---|---|---|---|
| **A** | UDT 内嵌动态数组字段 `x.arr(i, j)` / `(i, j, k)` | 成员访问叶子落进 `IndexOrCallExpr` 的 callee 分支，按"函数调用"发 `(vb6_<x>_arr)(i, j)` ⇒ C2064，连带 C2197 | 按实参秩发 `VB6_SA_ND_AT2 / AT3 / ATN`（新秩判定在 `cgen_util_classtype.cpp`），`cgen_expr_call_callee_member.inc` + `callee_withm.inc` 两处接收者接入 |
| **B** | 语句级 `obj.M -1 * halfWidth, ...`（`baseArrow.AddTriangle -1 * triangleHalfWidth, …`） | 旧判据只认 `LiteralExpr` 作右操作数；`-1 * w` 被解析成 `Sub(1, w)` ⇒ 整句落空 ⇒ `(callee - (1*w))(args)` ⇒ C2064（实测 9 处 / 7 TU：`Dialog_GradientEditor`、`FormLevels`、`Interface`、`PD2D`、`pdEdgeDetector`、`Tools_Clone`、`Tools_Gradient`） | `parser_stmt_assign.cpp` 的折叠判据放宽到"任意非 String/Date 数值表达式"（语句级不存在减法语句，`M -` 只能是调用首参负号） |
| **C/D/E** | 值成员空括号 `obj.Prop()` / `arr()` / `classField()` 读 | 值成员被按"函数调用"发码 ⇒ C2064（`term does not evaluate to a function taking 0 arguments`） | `cgen_expr_call_builtin_argtype.inc` 按"控件属性 / 数组 / 类字段"三类收口，值成员空括号退回值访问 |

### 38.2 夹具与哨兵

- 新夹具三份：`tests/b606_udtmem_nd`（A）、`tests/b607_neg_first_arg`（B）、`tests/b608_value_member_paren`（C/D/E）。
- 三条 static 哨兵：`check_g4_udt_member_nd.ps1` / `check_g4_neg_first_arg.ps1` / `check_g4_value_member_paren.ps1`；
  负例（定点破坏判据）**全部判红**，正例全绿；已注册进 `tests/run_tests.ps1` static 分片并通过
  `scripts/check_static_sentinel_registration.ps1` 普查。
- 本机 `tests\run_tests.ps1 -Category syntax -Incremental`：**PASS=190 FAIL=0 SKIP=0**。

### 38.3 形状门（629 → 632）与 bless 记录

- **本刀对存量 629 行零变化**，两条独立证据：① 本机 stash 前后用同源码各跑一遍受影响输入，发码**逐字节相同**；
  ② CI 自己那台的清单：629 行相同 / 0 行哈希不同 / 0 行缺席 / 3 行多出（多出的正是 b606/b607/b608）。
- bless 只登记新增 3 行（`b606/b607/b608`），存量 629 行逐行未动；`git diff --numstat` = 3 增 0 删；
  登记后复算 **632/632 相同**，coverage K1..K3 绿。提交 `0337282a`。
- **本机那 4 条 OCX 侧哈希差异一律不登记**：`tests/guisample/GUISample.vbp`、`tests/guisample/vb6ref/src/vc6/GUISample.vbp`、
  `tests/guisample/vb6ref/src/vs2015/GUISample.vbp`、`tests/TeeChartChartDisplay/DisplayProject.vbp`。
  归因链：本机哈希与 CI 官方 `C3.exe`（run `42886770374` 的 `c3-exe` 工件）复算结果**相同** ⇒ 不是 Debug/Release，
  也不是本刀；主工作树 `tests/guisample` 下有一批 **gitignore 的 OCX 侧文件**（`Codejock.*.ocx/.lic`、`RICHTX32.OCX`、
  `GUISample.exe.manifest`），C3 据此整块发"`Object=` 免注册 OCX 控件表"。
- **硬口径（写死，后续不许改判）**：**本机形状门必须从干净克隆跑** —— `core.autocrlf=false`（LF 工作树）、
  且不带任何 gitignore 的 OCX 侧文件。该口径下 gs 的哈希从 306989 字节档落回 LF 档（约 −25K 字节发码），
  但上述 4 行**仍然**因本机注册的 typelib/OCX 与 CI 不同而不同：干净克隆的读数是
  **632 行 / 628 相同 / 4 条既知环境差异 / 0 缺席 / 0 多出**。这 4 条永远不进基线，也不作为"发码变了"的证据。

### 38.4 整编读数（基线 `42884417424` → 本刀 `42915022004`）

| 诊断 | 修前 总量/TU | 修后 总量/TU | Δ |
|---|---|---|---|
| **C2064** | 66/12 | **0/0** | **−66 / −12** |
| **C2197** | 219/21 | **187/19** | **−32 / −2** |
| **C2198** | 133/54 | **123/49** | **−10 / −5** |
| C2296 | 14/12 | 13/12 | −1 |
| C2065 | 228/96 | 228/96 | 0 |
| C2440 | 214/55 | 214/55 | 0 |
| C2039 | 88/12 | 88/12 | 0 |
| C2224 | 59/16 | 59/16 | 0 |
| 全码合计（excl C1003） | 1276 | **1167** | **−109** |

- **没有新增任何错误码**（逐码对表：变化项只有上表 4 行，其余逐一相同）。
- `C1003` 两轮均 0 TU（截断在本轮基线的上一刀已解除）。
- 族级对账（run log 抽样 + 机理）：`vb6_lwa_RtlMoveMemory` 的 C2064/C2197 成对消失（A）；`ApplyTranslation` 那条
  `too few arguments` 消失（B）；`vb6_MonthName` 5/5 条抽样**原样保留**（M3 范围，见 §39）。
- `BUILD_EXIT=1` / `EXE_EXISTS=False` 两轮一致（工程仍有 1167 条错误，本刀不改这条结论）。
  `PEAK_TREE_WS_MB 8892→8885`、`MIN_FREE_PHYS_MB 4851→4832`、`BUILD_SECONDS 6173.7→4847.1`。

### 38.5 M3 取证更正（与 §36 M1 口径对齐）

- `pdStringStack_GetString` 的 17 处 6 实参**不是 Optional 缺省**：`pdStringStack.GetString(strIndex)` 只有 1 形参，
  实参串成了 `pdSerialize.GetString` 的 3 形参 + 2 枚 `_has_` 尾标记 ⇒ 与 §36 M1 **同一族**
  （接收者类解析不出 ⇒ 按成员名命中跨类同名形参表），**不是**"少写实参"。归 M1 已修口径，不再单列。
- `vb6_Randomize()` 那 7 处也**不是缺省**：源码是 `Randomize -1` / `Randomize 1234567` / `Randomize Timer`
  （`pdPaletteChild.cls:291-292` 对应 emit 431/432）—— 是**实参整体丢失**，随 2026-11-09 的 `Randomize` 语句收口
  （`2600664d`）已消失；当前 head 四形态探针：`vb6_Randomize(-1)` / `vb6_Randomize(1234567)` / `vb6_Randomize(<Timer>)` /
  裸形 `vb6_RandomizeTimer()`，全部带实参。
- 结论口径只有一句：**M3 的"缺省"只剩真正的 Optional 省略档**（`MonthName` / `WeekdayName` 一类）；
  同名成员形参表串台属 M1，实参丢失属解析器，两者都不计入 M3。

### 38.6 `Tests (vbp #5)` 留证（判真抖动，判据不动）

- 现象：run `42886770374` 的 `Tests (vbp #5)` 掉 `[PIXELS] uctreeview_px.ps1`：`'Contact  List' 内容段 BM_CLICK 3s 未返回`。
- 判据与处置：**同一 job `130077321416` 原地复跑 ⇒ 6m54s 全绿**；同一二进制上一轮该分片 `PASS=48 FAIL=0`；
  夹具 `ucTreeView` 无菜单、无 Slider（与本刀 M2 改动面无关）⇒ **真抖动（迟到），不是消息丢失**。
- 采集侧只加迟到判定（`uctreeview_px.ps1` 的 `Invoke-ClickSyncLate`：3s 超时后再等 7s 找窗体，
  命中记 `jitter[$btn]=BM_CLICK(late=…ms resp=…)`），**不改任何既有判据/超时阈值**。提交 `730e461d`。
- 交叉证据：本刀 CI `42913600247`（`0337282a`）与 `42918309147`（`76bbb0e9`）的 `Tests (vbp #5)` 均绿；
  两个 run 的失败分片都只有存量三项 `compile #1/#2`、`vbp #11`。

### 38.7 下一刀

M3（§39）已随本文件同轮落地提交 `76bbb0e9`：`MonthName/WeekdayName` 的 Optional 省略档补齐
（`kOD_MonthName{nullptr,"0"}` / `kOD_WeekdayName{nullptr,"0","1"}`，补齐点仍唯一），
配 `scripts/check_m3_builtin_opt_defaults.ps1`（标记唯一、表行唯一、档数 2/3、首档必填不补、与 RTL 声明逐档对账，
负例全红）。CI 形状门 `Emit manifest (shape oracle)` 在 `76bbb0e9` 上 **success**（632/632 相同），
即 M3 对存量语料零形状变化；本机干净克隆读数同为 632/628/4 既知/0/0。
下一轮整编（同探针流程、head `76bbb0e9`）用于取 `vb6_MonthName` 7 → 0 的验收读数并经 §39 收口。

## §39 M3：MonthName / WeekdayName 的 Optional 省略档（收 C2198）

**提交** `76bbb0e9`。**机理**：VB6 签名 `MonthName(month[, abbreviate])` /
`WeekdayName(weekday[, abbreviate[, firstdayofweek]])`，RTL 出口是 **2/3 参固定 C 签名**
（`vb6rtl_builtin.h:528-529`）；省略档不补 ⇒ `vb6_MonthName(Month(Now))` 1 实参打 2 形参 = C2198
（PhotoDemon 7 处：`ScreenCapture.bas` / `Plugin_EZTwain.bas` / `pdClipboardMain.cls`）。

- 补齐点仍**唯一**（`kBuiltinOptDefaults`）：`kOD_MonthName{nullptr,"0"}`、
  `kOD_WeekdayName{nullptr,"0","1"}`（abbreviate=False；firstDayOfWeek=vbSunday=1，RTL 亦把 <1 折成 1）。
  **首档（month/weekday）是必填档，写 nullptr 不补** —— 少了它就该停在 C2198，不许变成静默错值。
- 哨兵 `scripts/check_m3_builtin_opt_defaults.ps1`：S1 标记唯一；S2 表行各一；S3 档数 2/3 且首档
  nullptr；**S4 与 RTL 声明逐档双向对账**（任一侧改形即红）。负例（删表行 / 把首档补成 "0"）判红。
- 本机探针（`--emit-c`）：`vb6_MonthName(x, 0)`、`vb6_WeekdayName(2, 0, 1)`，显式两参形原样不动；
  裸 `Randomize` 仍走 `vb6_RandomizeTimer()`。
- 形状门：CI `Emit manifest (shape oracle)` 在 `76bbb0e9` 上 **success**（632/632 相同）；
  本机干净克隆 632 行 / 628 相同 / 4 条既知环境差异 / 0 缺席 / 0 多出。**无需 bless**。
- 整编读数：见 §39.1（本轮探针 head 与读数落地后补）。

## §40 语句级无括号调用的首参括号分组（收 Num_Div 族 C2197）

**提交** `f5a9dc1f`。**机理**：VB6 语句级没有"表达式语句" ⇒ `f (...)` 只能是调用，括号是**首参的分组**：

```
fHSVtoRGB (x - trackRadius) / hueSpread, 1, 1, tmpR, tmpG, tmpB
' = fHSVtoRGB((x - trackRadius) / hueSpread, 1, 1, tmpR, tmpG, tmpB)
```

（PhotoDemon `Controls/pdSliderStandalone.ctl:1135`，实测 4 个站点 / 24 条 C2197。）
表达式层把括号当实参表（f 只收到 1 个实参），语句层再把整个二元表达式当 callee、
把 `, 1, 1, …` 挂到调用结果上 ⇒ 实参被塞进 `/` 的 helper：
`vb6_Num_Div((double)(vb6_Colors_fHSVtoRGB(...)), (double)(hueSpread), vb6_VariantEmpty(), 1, 1, tmpR, tmpG, tmpB)`。

- 复用 Fix 110y 折叠法，扩到**裸标识符 callee**；`obj.Member` 那格仍由 110y 原判据管；
- 同族"无二元运算符"一格 `f (expr), a, b`（原判 VB3005 1/N、发码成 `vb6_f(...)(a,b)`）一并收；
  仅当首参**不是可寻址左值**时折叠 —— `f (v), x` 的括号在 VB6 是按值传临时，折叠会变成按引用
  传原变量（语义不同），该形**保持现状**，作为已知缺口留在原地；
- 哨兵 `scripts/check_np_call_firstarg_binop.ps1`：标记唯一、`cur_ == Comma` 判据、
  裸标识符 callee、恰 1 位置实参、可寻址左值豁免；负例（去掉 `!argAddressable`）判红。
- 本机探针（`--emit-c`）五形态：
  ① `(x-r)/h` → `vb6_fHSVtoRGB(vb6_Num_Div((double)((x - trackRadius)), (double)(hueSpread)), 1, 1, &tmpR, &tmpG, &tmpB)`
  ② `(x-r)*h` → 乘法直发 ③ `f (x-r), 1, …` → 6 实参直发 ④ `TakesThree x / h, 1, 2` 不变
  ⑤ `y = Dbl(x) / h`（赋值路）不变。
- 形状门：本机干净克隆 **632 行 / 与 §39 清单逐行比对 changed=0**（628 相同 + 4 条既知环境差异），
  即对存量语料**零形状变化**，**无需 bless**。
- 整编读数：见 §40.1（待下一轮探针落地后补）。

## §41 C2039 88：取证（本轮未改码）

本轮只做取证，不改判据、不动源码。CI `42933219765`（dev `f5a9dc1f`）：形状门 success、`vbp #5` success，
失败分片仅存量 `compile #1/#2`、`vbp #11` —— 与 §38.6 一致。

### 41.1 分族（基线 run `42915022004` = `ac25f3aa` 的 `c3-error.log`，共 88 条 / 12 TU）

| 成员 | 接收者 | 条数 | 族 |
|---|---|---|---|
| `tudWidth` / `tudHeight` / `tudResolution` | `vb6_cls_pdResize` | 12+12+8 = **32** | 设计期子控件当值用 |
| `GetLayerByID` / `RequestThumbnail` / `GetSaveState` / `GetLayerByIndex` / `GetOriginalFileFormat` / `EstimateRAMUsage` / `IsSuspended` / `GetActiveDIB` / `GetColorProfile_Original` / `GetNumOfLayers` / `SetDPI` / `SuspendImage` / `HasOriginalPalette` / `IsActive` / `GetCompositedImage` / `SetColorProfile_Original` | `vb6_cls_pdImage` | **47** | 接收者 = `PDImages.GetActiveImage` |
| `Top` / `Left` | `vb6_cls_pdSelection` | 2 | 待定 |
| `Index` | `vb6_UserControl_Extender_Type` | 2 | 待定 |
| `cmbZoom` | `vb6_cls_pdStatusBar` | 1 | 待定 |
| `CanvasView` | `vb6_cls_pdCanvas` | 1 | 待定 |

### 41.2 族①「设计期子控件当值用」（32 条，已在小语料复现）

`Controls/pdResize.ctl` 的 Begin 块里有三枚工程内 UserControl 子控件（`PhotoDemon.pdSpinner`）：
`tudWidth` / `tudHeight` / `tudResolution`。源码里既有**值上下文**也有**对象上下文**：

```
Private Function ConvertUnitToPixels(ByVal curUnit As PD_MeasurementUnit, ByRef tudSource As pdSpinner, ...)
pxWidthFloat = ConvertUnitToPixels(GetCurrentWidthUnit, tudWidth, GetResolutionAsPPI(), m_initWidth)  ' 对象（ByRef pdSpinner）
tudWidth = newWidth                    ' 赋值 → 默认属性 Let（pdSpinner 的 Value，VB_UserMemId=0）
vb6_ret_ResizeWidth = tudWidth         ' 读默认属性
```

小语料探针 `.temp/c2039probe/`（`Probe.vbp` + `Spinner.ctl` + `Holder.ctl`，`tudWidth` 为
`Begin Probe.Spinner tudWidth`）在本机 head 上**逐字复现**（`--emit-c`）：

```
d = me->tudWidth;                             /* 读 → 结构体成员（不存在） */
vb6_Holder_ReadIt((void*)me, me->tudWidth);   /* 传参 */
me->tudWidth = d;                             /* 写 → 结构体成员 */
```

同一份 emit 里成员调用是**对的**（`vb6_ComGetObjectProp(vb6_hwnd_tudWidth, L"IsValid")`），
即设计期子控件的句柄宏 `#define vb6_hwnd_tudWidth (*vb6_UC_DesignSlotOf(me,"tudWidth"))` 在，
只有「裸名当值/当对象」这条路没有折。**根因**：`ctrlDefaultPropOf`（账 #256）只查
`knownFormControls_` 里能映射到 `FrmControlType` 的表（`getDefaultPropertyName` 的 13 个内置型），
工程内 UserControl 子控件映射不到默认属性 ⇒ `folds=false` ⇒ 落到 `me-><名>`。

**修法（下一刀）**：两条腿都得有，且只许一个出口回答——
- 值/标量上下文：从该子控件的 `.ctl` 里读默认属性（`Attribute <Prop>.VB_UserMemId = 0`，
  实测 `pdSpinner.ctl:320 Attribute Value.VB_UserMemId = 0`），发 `vb6_ComGetObjectProp(vb6_hwnd_<名>, L"Value")` 再按型收；
- 对象上下文（`ByRef <UC>` 形参、`Set`、`With`）：发 `(vb6_cls_<Type>*)vb6_UC_InstanceOf(vb6_hwnd_<名>)`
  （与 `cgen_with.cpp:803` 的工程内 UC 实例同构；设计槽存的是 HWND，不是实例指针）。

### 41.3 族② 接收者 = `PDImages.GetActiveImage`（47 条）

emit 里同一形态在**大部分 TU 编得过**，说明不是「漏 `()`」那种纯拼写错（漏 `()` 应报 C2228，
而全量诊断表里 C2228 = 0）。需要按 TU 定位到真实行（`Tools_Clone.c(743)` 等）后再定论：
`pd_emit.out`（48 MB，`42927780444` 的工件）只按文件顺序拼接、无行号，无法直接对行。

### 41.4 下一刀入口

先做族①（32 条，已复现、机理清楚、判据唯一）：小语料正负 → 哨兵 → 本机重建 → 632 形状门 → 双推 + 整编读数。
族② 需先拿到**当前 head** 的 `c3-error.log` 与单 TU 行号（PD 整编 #3 的工件可直接给），再定。
