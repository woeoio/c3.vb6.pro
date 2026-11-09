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
认不出（外部对象 / 宿主句柄 / 非本族包装器）= NULL 视图 = Nothing —— 这是把原
VB4001（编译期）降成运行期「给 Nothing」，**方向是少报错**，真值待 VB6 探针钉（见 §21）。

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
