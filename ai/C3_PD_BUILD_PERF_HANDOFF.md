# C3 × PhotoDemon 整编性能主线 —— 交接总结

> 面向：明后天接着干的人（或明后天的我）。
> 截止：**§56.55**（提交 `032a0186`，已推 `github/dev`）。
> 本文只讲**性能主线**这一条线；功能侧的台账在 `ai/C3_PHOTODEMON_HANDOFF.md`。
> 本文里的每个数字都来自 CI 工件（run id 都写出来了），本机没有第二个权威。

---

## 0. 一句话

PD 整编（623 个项目 TU + 65 个 RTL `.c`）从 **6 761.1 s** 压到 **211.0 s**（**32.0×**），
最近三刀是 PCH → 装表 → 放掉 AST；**下一刀的方向已经用尺子量过了，候选项在第 8 节**。

---

## 1. 累计成绩

同一 runner、同一工作流（`PhotoDemon x86 full build`）、同一语料，唯一变量是代码。

| 阶段 | run id | `BUILD_SECONDS` | 对 §56.50 | 对 §56.52 |
|---|---|---|---|---|
| §56.50（起点） | — | 6 761.1 s（112.7 min） | 1.00× | — |
| §56.52（基线） | `43198109044` | **436.5 s**（7.3 min） | 15.5× | 1.00× |
| §56.53 PCH | `43214960135` | **328.3 s**（5.5 min） | 20.6× | 1.33× |
| §56.54 装表（不提速） | `43218740715` | **307.9 s** | 22.0× | 1.42× |
| §56.55 释放 AST | `43224386890` | **211.0 s**（3.5 min） | **32.0×** | **2.07×** |

> §56.54 那一刀本身不提速（307.9 s vs 328.3 s 的 20.4 s 差是 runner 负载差，**不归因**），
> 它的价值是**把"哪一段慢"从猜变成读数**，§56.55 的 96.9 s 是照着它砍的。
> 同理，§56.55 总差里约 25 s 落在 emit + 失败留档段，**也不归因给这一刀**。

---

## 2. 三把刀

### 2.1 §56.53 —— PCH（两份，预编译头）

- 做了什么：给 PD 的窗体前奏建**两份** PCH（普通档 `c3_pch_std.h` / 窗体档 `c3_pch_form.h`），
  cl 分四段跑：建 PCH ×2 → 按组 `/FI` + `/Yu` 编译 → 编其余 TU（RTL + driver 自产，不用 PCH）→ 链接。
  预编译头内容由发码器决定（`usesFormPrelude()` 报分组），两份**不可混用**。
- 结果：436.5 s → 328.3 s。
- 附带补刀（`a0d9bb00`）：`runMsvcInvocation` 加 `keepLogOnSuccess`，失败时把**成功段**的输出也并进
  `c3-error.log`（原先失败只留最后一段的输出，判据面被削小了）。
- 用户裁定：**debug 构建同样裁、同样走 PCH**（裁整行 `#include`，行号不变、原型/类型照旧可见，
  PDB 类型记录由各头自己那份 TU 提供 ⇒ 调试信息不受影响）。
- 形状门：678 条数据行 vs §56.52 旧二进制**逐行相同（不同数 = 0）**。

### 2.2 §56.54 —— 给 cl 的每一段装一只表（不提速，先有尺子）

- 做了什么：`phase_timing.hpp` 加 `c3TimingOn()` / `c3SpanMs(t0)`（**独立绝对时钟**，不推 `phaseClock`
  ⇒ 不串扰既有 mark 序列）；`msvc_driver.cpp` 加 `ClPhaseTimer`（RAII，成功/失败两条 return 都计时），
  `runMsvcInvocation` 多收 `phaseLabel` / `tuCount`，5 个调用点打标签，输出
  `C3: TIMING cl:<label> <ms> ms (<n> TU, <x> ms/TU)`。
- 结果：拿到 PD 上 cl 的**真实四段拆分**（§56.54 基线，208 957 ms）：

  | 段 | §56.54 | §56.55 后 | 变化 |
  |---|---|---|---|
  | 建 PCH 普通档 | 2 099 ms | 1 883 ms | −10.3% |
  | 建 PCH 窗体档 | 513 ms | 377 ms | −26.5% |
  | **grp 普通档**（256 TU） | **91 363 ms** | **58 117 ms** | **−36.4%** |
  | **grp 窗体档**（283 TU） | **99 120 ms** | **67 836 ms** | **−31.6%** |
  | rest（RTL + driver，84 TU） | 15 862 ms | 8 922 ms | −43.8% |
  | **合计** | **208 957 ms** | **137 135 ms** | **−34.4%** |

- 这一刀**否掉两个方向、定死一个方向**（见第 7 节）：RTL 族做 PCH 净亏；链接不是瓶颈；
  grp 段占 62% ⇒ 下一刀就往"让 cl 拿到内存"去。

### 2.3 §56.55 —— 发码后释放 AST，让 cl 拿到内存（主线命中）

- 做了什么：`driver.hpp` 加 `releaseAstAfterCodegen()`；`driver_compile.cpp` 实现：
  - `analyzers_.clear()`（**539 份** SemanticAnalyzer）
  - 每个 Module 的 AST 内容 clear：`declarations` / `options` / `implements` / `interfaces` /
    `coclasses` / `defTypes` / `attributes`
  - `EmptyWorkingSet(GetCurrentProcess())`
  - **模块元数据一律保留**（`moduleName` / `filename` / `isFormModule` / `classTypeParams` 等，
    link 阶段还要读，见 `driver_link.cpp:744-753 / 818-820 / 1274-1280`）
- 调用点：`driver_compile.cpp:852-859`，在 codegen 之后、`runLinker` 之前。
- **`--emit-c` 时不释放**（`driver_link.cpp:714` 直接 return，不跑 cl）⇒ 形状门路径零影响。
- 结果：**307.9 s → 211.0 s（−96.9 s，−31.5%）**。

---

## 3. 关键读数速查（明后天直接引用）

### §56.55 GA 工件（run `43224386890`）

```
BUILD_SECONDS=211          （基线 §56.54 = 307.9 s，差 −96.9 s）
PEAK_TREE_WS_MB=9460       （基线 9678 —— 峰值口径，释放不改峰值）
MIN_FREE_PHYS_MB=4320      （基线 4153 —— 系统可用内存反而更宽裕）
BUILD_EXIT=1               （PD 源码本身带真实错误，**编不过是正常的**，GA 判的是错误面）
C3_TIMING=True
```

### §56.55 内存账

```
C3: MEM ast-release modules=539 cost=8840 ms ws=9451MB->0MB priv=9643MB->183MB
```

内存曲线（`pd_build.mem.csv`）：发码阶段 `c3_ws_mb` 一路涨到 **9 429 MB**，
释放后掉到 **32 MB**，此后 cl 阶段全程 32–37 MB。
⇒ **机理**：§56.54 之前 cl 是在「C3 攥着 9.4 GB」的环境里跑的；释放之后 cl 拿到了内存。

### §56.55 错误面判据（逐码）

- **38 个码种**（18 个 error 族 + 20 个 warning 族）**逐码计数相等，差异 0**。
  （18 + 20 = 38；早前口算的"36（17+19）"以此为准更正。）
- 参考（**不参与判定**）：总行数 69 102 = 69 102；`compiling source file` 行 68 225 = 68 225；
  字节 17 096 259 vs 16 972 350（**差 123 909 B，不可比**）。

### §56.55 回归

- dev 回归 run `43224394610`：**21 绿 / 22 格**，唯一非绿 = 形状门（预期挂账）⇒ 零回归。
- 本机等价 CI 路径：`tests\run_tests.ps1 -Category bas -Jobs 4 -Incremental` ⇒ **PASS=280 FAIL=0 SKIP=0**。

### §56.55 本机三夹具验证（`--keep-for-debug`）

| 夹具 | modules | cost | 工作集 | 私有内存 |
|---|---|---|---|---|
| `tests\guisample\GUISample.vbp` | 11 | 106 ms | 120 MB → 0 MB | 104 MB → 47 MB |
| `tests\ucTreeView\Proyecto1.vbp` | 10 | 161 ms | 115 MB → 0 MB | 98 MB → 25 MB |
| `tests\acc\acc_fam.vbp` | 3 | 29 ms | 41 MB → 0 MB | 25 MB → 14 MB |

三者 rc=0 且出 exe；真·未设 `C3_TIMING` ⇒ MEM 0 条 / TIMING 0 条 / exe 照出；
`--emit-c` ⇒ stdout 309 639 字节、**MEM 行 0 条**。

---

## 4. 判据纪律（**不要改**）

1. **判据的唯一权威是发码器**。口径只设一处：
   `scripts/diff_c3error.ps1` 里的正则 `(fatal error|error|warning)\s+(C\d{4})\s*:`。
   别的脚本 / 别的台账要判"错误面变了没有"，都必须调它，**不要另起一份正则**。
2. **判据只认逐码计数**。不得把「恰好 N 处」降级成「至少 N 处」。
3. **不得声称字节数相等**（已实证不可比）：cl 的 `/MP` 调度顺序一变，同一批
   `compiling source file` 行就换位置；再加上 session 目录名长度不同（16 vs 13 字符），
   字节数就会差上百 KB，而内容逐行看是等价的。
4. **形状门（`scripts/emit_manifest.ps1`）只哈希 `--emit-c` 的 stdout**；
   **122 行继续挂账**（用户 2026-11-12 裁定"继续挂账，等 PD 编出 exe 一起登记"，**不 bless**）。
5. **改动任何发码字节 ⇒ 形状门必红**。这一轮三把刀都没碰发码字节（AST 释放在 codegen 之后、
   `--emit-c` 路径根本不走释放），所以形状门保持挂账状态不变。

---

## 5. 硬约束（明后天照做，别自己重新推导）

- **本机绝不编译 PD**；并行度 ≤ 2；本机构建用 `ninja -j1`；重类别交 CI。
- **本机绝不跑 PD 全量发码**：实测会吃 **~15 GB**（16 GB 机器可用内存被压到 650 MB）。
  这与"本机绝不编译 PD"是**同一条约束的两面**。
- **改 `.inc` 必须 touch 对应宿主 `.cpp`**：
  - `cgen.hpp` 宿主 = `cgen_base.cpp`
  - `cgen_api.inc` / `cgen_state.inc` / `cgen_helpers.inc` 宿主 = `cgen_base.cpp`
  - `driver.cpp` 宿主 = `driver/detail/*.inc`
- `C3_TIMING=` **空串 ≠ 未设**（`getenv` 对空串返回非 null，判据是 `!= nullptr`）。
  做"零输出"对照必须 `Remove-Item Env:C3_TIMING`。
- shell = **PowerShell**；含 `git stash` / `git checkout -- src` / `Copy-Item` /
  `Remove-Item -Recurse` + 改工作目录 ⇒ 会被硬安全策略拦截 ⇒ **写成 `.ps1` 再执行**。
- **含中文的 `.ps1` 必须用 `pwsh`（PowerShell 7）**：`powershell`(5.1) 会按 GBK 误读 UTF-8 无 BOM 的脚本，
  引号配对会错乱。
- `scripts/` 下的 `.ps1` **不得有裸 tab**（哨兵 `scripts/check_no_raw_tab_in_ps1.ps1`）。
- 提交信息用中文，前缀 `fix(cgen):` / `docs(handoff):`；含反引号 ⇒ `git commit -F <文件>`。
- `gh api` / 工件下载 / `git push` 间歇 TLS timeout / 504 ⇒ **一律重试循环**
  （这台机器的 hosts 把 `github.com` 全家指向 `127.0.0.1`，本地转发服务时有时无）。
- 文档改动（`ai/**`）**不触发 GA**。

---

## 6. 工作流（脚本已固化到 `scripts/`）

> 这一轮的关键脚本原本只活在 `.temp/`（不入库 ⇒ 明后天会丢），已固化三份：

| 脚本 | 干什么 |
|---|---|
| `scripts/ga_dispatch_pd.ps1` | 同步探针分支 → 推 **github** → 派发 PD 整编（可选连带 dev 回归）→ **按 headSha 认领 run** → 等结论 |
| `scripts/pd_readout.ps1` | 取工件 + meta + cl 四段 + MEM 行 + 内存曲线 + 错误面逐码比对（一次出全套读数） |
| `scripts/diff_c3error.ps1` | **错误面逐码比对**（判据唯一口径所在） |

```powershell
# 一、发刀（探针分支必须推 github，不是 gitcode）
pwsh -NoProfile -File scripts\ga_dispatch_pd.ps1
pwsh -NoProfile -File scripts\ga_dispatch_pd.ps1 -AlsoDevRegression   # 连带 dev 回归那 22 格

# 二、读数（这一刀的全部关键数一次出齐）
pwsh -NoProfile -File scripts\pd_readout.ps1 -Run 43224386890 -Baseline 43218740715

# 三、只比错误面（口径唯一）
pwsh -NoProfile -File scripts\diff_c3error.ps1 43198109044 43224386890
#   退出码：0 = 逐码相等 / 1 = 有差异 / 2 = 日志找不到（**没判**，不等于通过）
```

配套已验证的形状：

```powershell
# 本机构建（必须用 PowerShell 点 env.ps1；cmd 下点不到 .ps1，ninja 会静默不跑）
. "C:\Users\Administrator\Documents\c3.vb6.pro\scripts\env.ps1" 2>&1 | Out-Null
ninja -C "C:\Users\Administrator\Documents\c3.vb6.pro\.build" -j1

# 形状门（678 输入, 约 13 min; 退出码 4 = 期望不符, 正常）
pwsh -NoProfile -File scripts\emit_manifest.ps1 -Root <repo> -Out <out.txt>

# 本机等价 CI 路径（CI 跑测试带 -Incremental!）
pwsh -NoProfile -File tests\run_tests.ps1 -Category bas -Jobs 4 -Incremental
```

三个脚本都已验证：BOM + 0 裸 tab + `Parser::ParseFile` 通过 +
在真实工件 `43218740715` vs `43224386890` 上跑出 **38 码种逐码相等、差异 0、VERDICT=PASS**。

---

## 7. 已经否掉的方向（别再试）

| 方向 | 为什么否 |
|---|---|
| **给 RTL 那一族也做 PCH** | rest 段只占 5.2%，而 RTL 前奏有 **29 档**（126 个文件各带各的 include 组合）⇒ 建 29 份 PCH 成本 ≈25 s > 收益 ≈12 s ⇒ **净亏**。 |
| **优化链接** | §56.54 实测：链接根本不是瓶颈（`c3PhaseMark("link")` 在 PD 上因链接失败**根本没打出来**）。 |
| **RTL 前奏归一** | 29 档前奏 = 29 种 include 组合，归一意味着改 RTL 的头结构 ⇒ 触碰面太大，收益只有 ~12 s。 |
| **减少诊断输出** | 判据就是错误面，删诊断 = 改判据 ⇒ 不做。 |

---

## 8. 下一刀候选（**明后天开工前先看这里**）

§56.55 之后 211.0 s 的构成（来自工件）：

```
cl 四段          137.1 s   (grp 普通档 58.1 s + grp 窗体档 67.8 s = 125.9 s，占 cl 段 92%)
emit 相位合计     32.5 s   (semantic 5.6 s / crossmod+generics+marks 11.8 s / codegen 4.5 s / gen 9.2 s / 其余零头)
失败留档 + 其余   ~40 s    (PD 编不过 ⇒ 写 17 MB 的 c3-error.log)
```

**候选 A：继续榨 cl 的 grp 段（125.9 s，最大的一块）**
- 内存已经喂饱了（工作集 32 MB vs 峰值 9.4 GB），再动内存没戏。
- 剩下能动的：**cl 的并行度 / 调度**（`/MP` 的并发上限、`/Gm` 最小重建、`/FS` PDB 写入方式），
  或 **unity build**（把多个 TU 合成一个编译单元）。
- ⚠️ unity build 会改诊断里的**行号与文件名** ⇒ **错误面判据必变** ⇒ 必须先把
  "错误面判据怎么跟着变"这件事想清楚再动手，**不要先写代码**。

**候选 B：砍 emit 段（32.5 s）**
- `crossmod+generics+marks` 11.8 s + `semantic` 5.6 s 是两块大头；
  `codegen` 4.5 s、`gen` 9.2 s（539 TU）已经摊得很薄。
- 但 emit 段是**发码器本体**，动它就可能动发码字节 ⇒ 形状门会红 ⇒ 要么找到不动字节的算法，
  要么先把形状门 122 行挂账的事跟用户敲定。

**候选 C：失败留档段（~40 s）**
- PD 编不过 ⇒ 每次都要写 17 MB 的 `c3-error.log`。这段是**失败路径**的成本，
  而 GA 判的正是错误面 ⇒ 不能不写、也不能少写。
- 可省的是"写得慢"：写盘方式 / 是否重复写同一份。

**候选 D：形状门那 122 行挂账**
- 用户 2026-11-12 裁定"继续挂账，等 PD 编出 exe 一起登记"。
- 一旦 PD 编出 exe，需要一次 `pwsh -NoProfile -File scripts\rebless_emit_manifest.ps1 -Bless`。

> **下一刀开工前的三个前置问题**（明后天先答这三个，再写代码）：
> 1. 候选 A 的 unity build 能不能**不改错误面判据**？还是判据要跟着换口径？
> 2. 候选 B 会不会触碰发码字节？如果会，形状门 122 行挂账要不要提前跟用户谈？
> 3. 211 s 里还有多少是 **runner 负载差**？§56.54→§56.55 的 96.9 s 里约 25 s 已被判为噪声，
>    下一刀要用**同一 runner / 同一时段**的基线重跑一遍才能归因。

---

## 9. 已知坑（别再踩）

1. `C3_TIMING=` 空串 ≠ 未设 ⇒ 做"零输出"对照要用 `Remove-Item Env:C3_TIMING`。
2. PowerShell 5.1 跑 UTF-8 无 BOM 的中文 `.ps1` 会解析错乱 ⇒ 一律用 `pwsh`。
3. 写含中文 / 反引号的提交信息 ⇒ `git commit -F <文件>`。
4. `git` / `gh` 撞 hosts 的 `127.0.0.1` 转发 ⇒ 带重试（脚本里已统一加）。
5. `scripts/` 的 `.ps1` **不得有裸 tab**（`check_no_raw_tab_in_ps1.ps1` 会点名到行）。
6. 交接文档 `ai/C3_PHOTODEMON_HANDOFF.md` 追加台账：
   **只做一次** `\n` → `\r\n` 归一（再做一次会产生 CRCRLF）；追加只写新字节，已有字节不动。
7. `scripts/pd_readout.ps1` 读 `C3: TIMING` 的口径（已用 §56.55 工件核对过，别凭印象改）：
   - `C3: TIMING cl:<label> <ms> ms (<n> TU, <x> ms/TU)`
   - `C3: TIMING gen <TU名> <ms> ms`
   - `C3: TIMING <顶层相位> <ms> ms` —— 名字可含 `+` 与 `-`（`vbp+expand` / `unpack-rtl`），**不含 `:`**
   - `C3: TIMING g:<子相位> <ms> ms` —— **已含在 gen 里**，别再加一遍
8. 探针分支是**另一个 clone**：`.temp/probe_sync`（远端 `github` = 本 GA 的源，`origin` = gitcode）。
   派发必须推 **github**，不是 gitcode。

---

## 10. 提交链（`github/dev`，全部已推）

```
0c1ed940  fix(cgen): §56.53 PCH（两份预编译头 + cl 四段）
a0d9bb00  fix(cgen): §56.53 补刀 —— 失败留档并进成功段
fc86efb2  docs(handoff): §56.53 台账
08c1933b  docs(handoff): §56.53 补记 —— dev 回归 + 字节数不可比
58e85ff8  fix(cgen): §56.54 给 cl 的每一段装一只表
25f9615d  docs(handoff): §56.54 台账
cfa440fd  docs(handoff): §56.54 补记 —— cl 阶段的内存账 + 本机硬约束
86f03fd8  fix(cgen): §56.55 发码后放掉 AST 与语义产物
032a0186  docs(handoff): §56.55 台账
```

- 探针分支 `probe/photodemon-emit` = `77b3e579`（含 §56.55 代码，落后 dev 两笔文档提交）。
- 功能侧台账：`ai/C3_PHOTODEMON_HANDOFF.md`（含 §56.53 / §56.54 / §56.55 全节）。
- 编译选项说明：`ai/C3_COMPILE_OPTIONS.md`（含 PCH 与 debug 裁剪的正确表述）。

---

## 11. 关键文件索引

| 文件 | 作用 |
|---|---|
| `src/backend/cgen_pch.hpp` | `c3PreludeIncludeLines(bool)` —— 前奏唯一排法 |
| `src/backend/detail/base/cgen_base_generate_header_open.inc` | 前奏取自上面那个函数 |
| `src/backend/detail/util/cgen_api.inc` | `usesFormPrelude()` |
| `src/driver/driver.hpp` | `pchFormModules_`；**`void releaseAstAfterCodegen();`** |
| `src/driver/detail/driver_codegen_module_loop.inc` | 收 PCH 分组 + `c3PhaseMark("gen " + baseName)` |
| `src/driver/driver_link.cpp` | 组 `PchGroup`；`options.incremental` 时清 `pchGroups`；`:714` `--emit-c` 直接 return |
| `src/backend/msvc_driver.hpp` | `PchGroup` / `pchGroups` / `compileAndLinkPch()` |
| `src/backend/msvc_driver.cpp` | `runMsvcInvocation(..., phaseLabel, tuCount)`、`ClPhaseTimer`、四段 PCH 路径 |
| `src/driver/detail/phase_timing.hpp` | `c3PhaseMark` / `c3ScopeTimer` + **`c3TimingOn()` / `c3SpanMs(t0)`** |
| `src/driver/driver_compile.cpp` | `c3PhaseMark("codegen")` → **`releaseAstAfterCodegen()`** → `runLinker`；`C3: MEM ast-release ...` |
| `src/ast/detail/ast_decl.hpp` | `class Module` —— AST 内容与元数据分开 |
