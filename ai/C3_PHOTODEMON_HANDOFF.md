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

### 6.5 CI 实测影响面

待 `probe/photodemon-emit` 复跑后回填（VB3001 应从 3693 实际下降；VB3044 应随 R4 从 19 降到 0）。

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

1. VB3003 剩余 57 条的隐式转换判据。
2. `--arch x86` 真编译（emit 后跑 cl/link），按编译错误反推下一批缺口。
3. 控件属性 / 方法缺失类缺口（需先有一份可编译的产物才能暴露）。

