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

## 3. 下一步（待用户放行后再做）

1. 单进程、低并行跑 `--arch x86 --syntax-only`，把错误按「语法 / 控件属性方法缺失 / 类型不匹配 / 缺失内置函数或常量 / 其它」分类统计并回填本文件（含文件:行 + 原始报错文本 + 同根因合并说明）。
2. 记录回归基线：现有测试套件结果 + GA 现状（全绿才算基线确认）。
3. 之后按「影响面小、复用度高」逐类补，每类一提交、写回本文件。
