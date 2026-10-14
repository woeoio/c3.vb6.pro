# 031 - Try / Catch / Finally 计划书（错误处理现代化）

> 状态：**待拍板，未动手**。目标语法的形状对照 twinBASIC（本地 `D:\tools\twinBASIC_IDE_BETA_983` 有证据），
> 语义的落地对照本仓库**已有的** setjmp/longjmp 那一套，不另起运行时骨架。
> 需求出处 = `ai/009-产品定位与最终成果展望.md` §5.8 + 缺口表里"P0 必须支持：On Error"那一条的现代化延伸。

---

## 一、现状读数（全部实读，不是推测）

### 1.1 运行时：已经是"全局 Err + setjmp/longjmp + 帧栈"

`src/rtl/core/vb6rtl/vb6rtl.c`：

| 事实 | 位置 |
| --- | --- |
| 全局 Err 对象只有 `number/description/source` 三个字段 | `vb6rtl.c:342-348` |
| 传播机制 = `jmp_buf* vb6_error_jmp_ptr`（指向**某个仍在栈上的帧**的局部 buffer）+ `longjmp` | `vb6rtl.c:356-358` |
| 帧栈 `vb6_ErrFrame vb6_err_stack[]` + `vb6_SaveErrState/RestoreErrState` | `vb6rtl.c:366-404` |
| 抛出三级：`ResumeNext` → 直接返回；有活动 handler → `longjmp(ErrNum)`；**都没有 → `ExitProcess(ErrNum)`** | `vb6rtl.c:418-438` |
| 除零/取模已经走 `vb6_Num_Div/vb6_Num_Mod` 抛 11（字面量除数；变量除数仍是裸 C 除法） | `vb6rtl.c:444-459` |
| `Erl` 恒 0（cgen 不发 VB 行号） | `vb6rtl.c:410-414` |

**两条对本功能致命的读数**：
1. `vb6_ErrFrame` **只存 jmp/resume 状态，不存 Err 对象本身**（`SaveErrState` 抄的七个字段里没有 `vb6_err`）⇒
   Catch 块里再出一次错就会把 `ex.Number` 冲掉。所以 `ex` **必须是进入 Catch 时刻的快照**，不能是"读全局"的活视图。
2. 抛出时无 handler 就 `ExitProcess` ⇒ 今天"错误穿过本过程往外传"这件事**没有中间钩子**，
   Finally 想在外层帧被通知到，必须新加一条"重抛"出口（见 T31-D）。

### 1.2 前端：`On Error` 一族已经全通，且已有一块"保护区"状态机

- AST 侧已有 `OnErrorKind{GoToLabel,ResumeNext,GoToZero}`、`ResumeKind{ResumeHere,ResumeNext,ResumeLabel}`、`ErrorStmt`
  —— `src/parser/stmt/parser_stmt_jump.cpp:39-78`。
- 发码在 `src/backend/stmt/cgen_jumps.cpp`：`On Error GoTo L` ⇒ `if (setjmp(vb6_local_err_jmp) != 0) goto L;`（`:30-74`）、
  `Resume` ⇒ 走过程尾的 `vb6_err_dispatch_switch`（`:79-96`）、`Error n` ⇒ `vb6_RaiseError`（`:99-103`）。
- `Err.Number/Description/Source/Raise/Clear` 已有 lowers：`src/backend/expr/cgen_expr_with.cpp:208-216`
  + `src/backend/detail/expr/cgen_expr_member_precheck.inc:80-90`。
- 过程级的保护状态字段已经存在：`inProtectedBlock_`、`currentErrorHandlerLabel_`、`dispatchPoints_`、`resumePointCounter_`
  （`src/backend/detail/util/cgen_state.inc:376-379`），扫描函数 `hasOnErrorInStmts()` / `hasResumeInStmts()`
  （`src/backend/cgen_util_scan.cpp:164`）。

### 1.3 帧的粒度现在是"过程"，而且只在过程自己写了 `On Error` 时才有帧

`src/backend/decl/cgen_decl_proc.cpp:236-292`：`hasOnError_` 才声明 `jmp_buf vb6_local_err_jmp;` 并 `vb6_SaveErrState()`；
尾部 `vb6_RestoreErrState()` → 释放 ANSI 临时 → `emitIvrefScopeRelease()` → `return;` → 不可达的 dispatch switch。

⇒ 本功能要把**帧的粒度从"过程"降到"块"**（一个 `Try` 一块），这既是语义要求（"最近的活动块接住"），
也是把 `Erl`/嵌套语义做对的唯一路子。

### 1.4 一条既存缺陷 = T31 的先决条件

`Exit Sub/Function/Property` 发的是**裸 `return`**（`src/backend/stmt/cgen_jumps.cpp:104-141`），
绕过了 `vb6_RestoreErrState()` 与 ivref/ANSI 释放尾（`cgen_decl_proc.cpp:263` 的注释自己承认"Exit Sub 例外"）。

今天它的后果是 ErrFrame 栈在 `Exit Sub` 时**漏帧**（`vb6_err_stack_top` 只增不减，溢出后 `SaveErrState` 静默不存）；
到了 Try 这一族，同一处会让 `Exit Sub` **直接跳过 Finally** —— 那是 VB/C# 用户最不能接受的行为。
所以第一格必须先把它收口。

### 1.5 语料扫描：`Catch` 已经被人当标识符用了（软关键字是硬要求）

`tests/ archive/ publish/` 里 `try|catch|finally|throw`（不分大小写）23 处命中，除下面这一处全在注释或字符串里：

```vb
' archive/vbman/src/Shadow/clsSubClass.cls:484
On Error GoTo Catch
' :490
Catch:
```

⇒ `Catch` 一旦做成硬关键字，这行旧码立刻编不过。**四件套软关键字**（`Via`/`CoClass` 那套，见 §三 B1）是底线，
而且必须带一条 `Dim Try As Long` / `Sub Catch()` 的正例。

### 1.6 twinBASIC 对照（本地实测，不是记忆）

tB 自带的 monaco VB 语法表（`ide/monaco-editor-0.35.1-pre-min/vs/basic-languages/vb/vb.js`）里：
块配对写的是 `["try","end try"]`（与 `sub/end sub`、`synclock/end synclock` 同列），关键字面出现
`Try` / `Catch` / `Finally` / `Throw` / `TryCast`。⇒ 语法形状照它：`Try` 开块、`End Try` 收口。

---

## 二、语法（v1 提案）

```vb
Try
    Dim v As Long: v = UnsafeDiv(a, b)
    DoWork v
Catch ex As Exception          ' 变量部分可省：`Catch` 单独成行也要能过
    Debug.Print "兜住了: " & ex.Number & " " & ex.Description
Finally                        ' 也可省；Catch / Finally 至少要有一个
    ReleaseEverything
End Try
```

支持嵌套、支持与 `On Error` 同过程**但互斥**（D5）。`Try` 与 `End Try` 之间必须是语句序列，
不允许跨过程、不允许 `GoTo` 进/出块（块边界 = 语句边界）。

---

## 三、待拍板的口径（我的推荐值已经写好，逐条确认或直接说"照推荐"）

- **D1 子句数量**：v1 = **至多一条 `Catch` + 至多一条 `Finally`**，顺序固定 `Try → Catch → Finally → End Try`。
  多路 `Catch ... Catch` 与 `When <条件>` 后置成独立一格。
  理由：与 longjmp 模型最匹配（一次抛出只落一个接收点），且旧 VB 码 99% 是"一个兜底 + 一份清理"。
- **D2 `Exception` 是什么**：**不**做真 COM 类、不进 tlib。它是 cgen 认得的**合成只读对象名**，
  成员只有 `Number` / `Description` / `Source`（映射到 catch 时刻的快照）。
  理由：走真类要吃 CoClass/接口那一整族（成本 = 一整条线），而这里要的只是"把三个字段起个名字"。
- **D3 快照语义**：进 `Catch` 前把 `vb6_err` 三字段抄进块内局部（并 `Err.Clear` 的全局效果**留在 Catch 出口之后**）。
  理由 = §1.1 缺陷 1：帧栈不存 Err 对象，不抄就会被 Catch 体内的二次错误冲掉。
  顺带把 `ex.Number` 在 Catch 体内**钉成常量**（重新 raise 也不变）。
- **D4 Finally 必须跑的四个出口**：① 正常走到 `End Try`；② `Try` 体内 `Exit Sub/Function/Property`；
  ③ `Catch` 体内 `Exit Sub`；④ 块内抛出且本块不接（无 Catch，或 Catch 体内又抛）⇒ **先跑 Finally 再往外传**。
  口径：Finally 恰好一次，不多跑不漏跑；Finally 里的 `Exit Sub` 允许（直接离开，不再重抛 —— 这条要在手册里写明是"吞"）。
- **D5 与 `On Error` 混用**：同一过程内既有 `On Error GoTo/Resume Next` 又有 `Try` ⇒ **编译期报错**（新增诊断，
  走 `src/common/diagnostics.hpp`：解析级 2xxx 下一个空号 2015，语义级 3xxx 当前最大 3042）。
  理由：resume-point 的 dispatch switch 是**过程级** goto 标签，与块级 setjmp 撞车；两套并存一定有一条说谎。
- **D6 `Catch` 里不许 `Resume` / `Resume Next`**（同 D5 的理由，v1 判死并给明确错误文本；`Resume` 在 `On Error GoTo` 那侧照旧全通）。
- **D7 `Throw`**：v1 **不做**。`Err.Raise n, , "文本"` 已经是通的（§1.2）。
  `Throw New Exception(...)` 要的是"异常对象 + 类型化传播"，与 D2 的"快照对象"不是一回事，另立一格再议。
- **D8 `ex.Erl` / 行号**：不提供（`vb6_Erl` 今天恒 0，§1.1）。谁要就先做行号标签机制，别在本题里夹带。

---

## 四、批次表（一批 = 一个改动 = 一扇门，串行）

| 批 | 内容 | 主要落点 |
| --- | --- | --- |
| **T31-A** | 先决条件：`Exit Sub/Function/Property` 从裸 `return` 改成**走统一出口尾**（把 `RestoreErrState` + ANSI/ivref 释放接上）。零语法改动 | `cgen_jumps.cpp:104-141`、`cgen_decl_proc.cpp:258-273` |
| **T31-B** | 语法面：词法**软关键字四件套**（`token.hpp:156` / `lexer_keywords.cpp:59` / `token.cpp:34` / `parser_helpers.cpp:61` `isSoftKeyword`）**+ 第五处 `isEndBlock()`**（`parser_helpers.cpp:185-202`）；parser `Try..End Try`（`parseBlockUntil` + `expect`）；AST 六处登记（`ast_enums.hpp:115`、`ast_fwd.hpp`、`ast_visitor.hpp`、`ast_stmt.hpp` 类、printer `ast_printer_visitor_stmt.inc` + `..._support.inc:59` switch、`ast_clone.cpp:245`） | 新文件 `src/parser/stmt/parser_stmt_try.cpp` |
| **T31-C** | 语义面：`ex` 进过程作用域（只读、类型 = 合成的 Exception）、`Try` 空体判死、Catch/Finally 都缺判死、与 `On Error` 混用判死（D5）、`Resume` 在 Catch 里判死（D6） | `src/semantics/semantic_analyzer_stmt.cpp`（照 `visit(IfStmt):71`）、`semantic_analyzer_dispatch.cpp:38` |
| **T31-D** | 发码面：块级 `jmp_buf` + 帧保存/恢复（粒度从过程降到块）、Catch 快照（D3）、Finally 四出口（D4）、新 RTL helper `vb6_Reraise()`；`hasOnErrorInStmts/hasResumeInStmts` 学会递归进新块 | `cgen_stmt.cpp:126`（照 If 的块模型）、`cgen_jumps.cpp`、`src/rtl/core/vb6rtl/vb6rtl.c`（**改 RTL 要重建 C3.exe 才生效**，两处登记已齐） |
| **T31-E** | 文档与收口：手册新页 `docs/vb6-manual/…/Try 语句.md`（含"与 On Error 的关系"这一节）、031 台账回填、批次表 | — |
| 之后 | 另立：多 `Catch` + `When`、`Throw`/类型化异常、`ex.Erl`（要先做行号机制） | — |

---

## 五、判据手法（复用仓库现成通路，不发明新体系）

1. **纯语法**（syntax 组）：`Test-Syntax` / `Test-SyntaxFail`（`--syntax-only`）。
   正例：`Try/End Try` 最小形、`Catch` 不带变量、`Finally` 无 Catch、嵌套两层、**`Dim Try As Long` + `Sub Catch()`**（软关键字）、`clsSubClass` 那种 `On Error GoTo Catch` 标签写法照旧能过。
   负例（每条一根针，且都要真红）：缺 `End Try`（期望串 `expected 'End Try'`）、`Finally` 在 `Catch` 前、两条 `Catch`、空 `Try` 且无 Catch/Finally、Catch 里 `Resume`、同过程混用 `On Error`。
2. **发码形状**（compile 组）：`Test-EmitcShape` / `Test-EmitcAbsent` 钉块级 `setjmp`、快照三行、`goto vb6_try_epilogue` 各**只出现一次**；`Test-EmitcAbsent` 钉"没写 Try 的过程里一个 `setjmp` 都不许多出来"。
3. **真编真跑**（bas 组 `Add-BasTest`→`Test-Run`）：判据全部走"打印 Y/N + 退出码"，
   至少覆盖：兜住被调过程抛的错（**跨帧**）、`ex.Number=11`（除零）、Catch 体内再抛不脏 `ex`（D3）、
   Finally 计数 = 1 的四个出口（D4）、`Try` 里 `Exit Sub` 后 Finally 照跑（这条同时钉住 T31-A）、
   无 Catch 时错传给**外层** Try、外层也没有 ⇒ 仍旧按 VB6 走 unhandled（不能变成静默 0）。
4. **x64 与 x86 双验**：门只跑默认架构 ⇒ 布局/调用约定类改动必须本地补 x86 真跑（既有纪律）。
5. **逐字节护栏**：41 件工程 `--emit-c`，BASE = 从干净 HEAD 重建的编译器，"只允许本批夹具变"。
   T31-A 那一格尤其要过它 —— 它动的是**所有过程**的出口形状。
6. **负控**：每条新助手先拿一枚假 needle 证明它能红，再证明修复前的二进制红在哪几条。

---

## 六、风险与已知边界

- **longjmp 跳过 C 作用域**：RTL 是纯 C（无析构），且帧栈由 `Save/RestoreErrState` 手工配平 ⇒ 只要 T31-A 把 `Exit` 那条补齐，块级方案是自洽的。真正的风险点是 MSVC 对 `jmp_buf` 局部 + 优化的组合，x86/x64 都要真跑。
- **每块一次 `setjmp`** 的开销：可接受（一次调用量级），且只在**含 Try 的过程**里生成；`hasOnError_` 那条"没写就不生成"的按需原则照搬。
- **帧栈深度** `VB6_ERR_STACK_SIZE`：块级帧会让用量上升，先量一遍现有上限和最深用例，必要时随包加一条溢出诊断（现在溢出是**静默不存**）。
- **不改的东西**：`On Error` 一族的行为一个字节都不动（存量工程靠它）；`Err` 对象不动；除零只覆盖字面量除数这一既有边界照抄进手册。
- **明确不做**（v1）：`Throw`、类型化 `Catch <Specific>`、`When` 过滤、多 Catch、`ex.Erl`、异步/跨线程错误传播、错误堆栈。
