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
| `Math` | 18 | `Modules\ComplexNumbers.bas` 的 `Math.Log/Abs/Cos`。工程内没有 `Math` 模块/类，VB6 原生也没有这个对象 ⇒ **疑似 tB-only 写法**；下一轮先确认 PhotoDemon.exe 的 VB6 构建是否真编这份文件，再决定是补内置表还是按条件编译处理 |
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

1. `VB3001` 剩余 38 条按 6.5 的分类逐类收口（内建常量表缺项 / UserControl 内建成员 / 块内声明的过程级作用域 / `Randomize` 语句形态）；`Math` 18 条先定口径再动。
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

- **语义定案**：VB6 `UserControl.Parent` 返回**宿主容器对象**（窗体 / 容器控件），**不是容器窗口句柄**；`Parent.Controls` 就是父容器的控件集合，与 `UserControl.ParentControls` 是同一件事。RTL 一侧的答案因此是 `vb6_UC_ParentObject()`（值位）与 `vb6_UC_ParentControls`（集合），不是 `hWnd`。
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

### 10.4 `Math` 18 条：不是 tB 专有，是 VB6 自带 VBA 类型库的 `VBA.Math`

- **定案（推翻旧记的「疑似 tB-only」）**：`Math` 是 VBA 类型库里真实存在的模块（`Element of VBA.Math`：`Log` / `Rnd` / `Round` …），VB6 引用 VBA6 库即拥有它；PhotoDemon 的 `INSTALL.md` 明确是「把 `.vbp` 载入 VB6 IDE 编译」，`Math.Log` 等 18 处因此本来就是合法 VB6。twinBASIC 的 `Math` 模块只是同一族（`docs.twinbasic.com/tB/Modules/Math`），不是来源。
- 语料形态：`Modules\ComplexNumbers.bas` 5 行 / 18 个限定成员（`Math.Log` / `Math.Abs` / `Math.Cos` / `Math.Sin` / `Math.Sqr` / `Math.Round`）。
- 修法（与已有 `VBA.<成员>` 折叠同一条通路，零新语法）：
  1. [semantic_analyzer_util.cpp](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/semantics/semantic_analyzer_util.cpp#L789-L794)：`isDocumentHostObject` 认 `math`，**仅限定位**（与 `vba` 同款；值位的裸 `Math` 仍报 VB3001）。
  2. [cgen_expr_member_precheck.inc](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/backend/detail/expr/cgen_expr_member_precheck.inc#L186-L194)：`Math.<成员>` 并入 `VBA.<成员>` 那一支，成员白名单 = `abs/atn/cos/exp/fix/int/log/rnd/round/sgn/sin/sqr/tan`（拼错的成员不会因前缀是 `Math` 就变合法）。
- 读数：修前 6 条 VB3001（`Math`）+ `vb6_Math_Log/Abs/Cos/Sin/Sqr/Round`（RTL 无此符号，真编译必 C2065）；修后 stderr 0 行，产物 `vb6_Log(1.0)` / `vb6_Abs((-4.5))` / `vb6_Cos(0.0)` / `vb6_Sin(0.0)` / `vb6_Sqr(9.0)` / `vb6_Round(2.6000000000000001, 0)`。
- 运行读数（真编译）：`VM-LOG=0` / `VM-ABS=4.5` / `VM-COS=1` / `VM-SIN=0` / `VM-SQR=3` / `VM-ROUND=3`。
- 断言：`test_vbamath`（Compile+Run 7 行读数）+ `vbamath_emitc_folds_to_builtins` 形状钉 + `vbamath_no_module_symbol` 反向钉；负控 `vbamath_bare_is_undeclared`（值位 `Set m = Math` 必须仍报 VB3001）。
- 记账：两个新夹具按同一口径延后登记 emit-manifest。

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

