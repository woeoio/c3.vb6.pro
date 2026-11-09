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

## 5. 下一步

1. 复跑 CI 验证 R5 归零（改夹具 vbp 后自动触发）；顺手把 artifact 里 `pd_emit.err` 分类结果回填本节。
2. 记录回归基线：现有测试套件结果 + GA 22 job 现状（全绿才算基线确认）。
3. 按 4.3 顺序逐类补，每类一提交（含最小复现用例），并写回本文件。
4. 之后的目标是 `--arch x86` 真编译（emit 之后跑 cl/link），届时再看控件属性/方法缺失类缺口。
