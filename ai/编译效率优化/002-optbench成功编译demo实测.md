# 002-optbench 成功编译 demo 实测报告

> 日期: 2026-09-03
> 对象: `publish/demos/optbench`（120 个 VB 模块 + Main 的**可成功编译**合成工程）
> 说明: 001 报告受 vbman 552 错误失败态限制，只能评估失败路径。本 demo 补足**成功路径**下的端到端验证与收益测量，重点检验 opt1 `/MP` 与 opt3 增量缓存。

## demo 结构

`publish/demos/optbench/`（全部保留入库，构建产物 `out/` 不入库）：

| 文件 | 作用 |
|------|------|
| `gen_modules.ps1` | 生成器：`-Count N` 生成 N 个保守 VB6 语法模块（每模块 7 函数约 60 行）+ `Main.bas`（逐一调用全部模块防死代码消除）+ `OptBench.vbp`，UTF-8 无 BOM |
| `src/M001..M120.bas`, `Main.bas` | 生成结果（121 个源文件，已入库） |
| `OptBench.vbp` | 工程文件（Exe, Startup="Sub Main"） |
| `build_optbench.bat` | 完整优化参数构建：`--incremental --no-warn 3001,3003`（include 瘦身已改成固定行为，不再是参数） |
| `bench.bat` + `bench_run.ps1` | A/B 计时包装：`-Tag <名> -C3 <exe> -Clean 0|1`，Stopwatch 计时写 `bench_<Tag>.log` |

预热验证：先生成 8 模块跑通（产出 exe、`.c3obj` 9 个 obj），再放大到 120 模块正式实测。

## 实测环境

- CPU: **2 核**；编译器: VS2022 BuildTools x64；C3: `.build\C3.exe`(/MP) 与 `.build\C3_nomp.exe`（同批构建，A/B 可信）
- 工程: 120 模块 + Main，全量生成 121 个 `.c` 编译单位
- 基准命令：`C3 OptBench.vbp --output-dir out --incremental --no-warn 3001,3003`（bench.bat 刻意不带额外参数，opt4 已在 001 单独评估；**注**：`--trim-includes` 已于 2026-11-12 撤掉，见 §56.52）

## 实测数据

| 场景 | C3 版本 | Clean | 耗时 | 说明 |
|------|---------|-------|------|------|
| `mp_full` | `/MP` | 1 | **40.5s** | 首次运行：冷态（vcvars/cl 冷启动、Defender 扫描新文件） |
| `mp_full2` | `/MP` | 1 | **6.2s** | 同参数热态复测 |
| `nomp_full` | 无 `/MP` | 1 | **7.0s** | 热态对照 |
| `mp_hit` | `/MP` | 0 | **6.2s** | 缓存全命中（二次构建） |
| `inc_1file` | `/MP` | 0 | **7.3s** | 改动 `M001.bas` 后增量 |
| `inc_chk` | `/MP` | 0 | **7.0s** | 复测 + obj 时间戳验证 |

产物核对：全量成功后 `out\OptBench.exe`（约 430KB）可生成；`out\.c3obj\` 含 cache.txt（7.4KB）+ **121 个 obj**（M001..M120 + Main）——与编译单位一一对应。

## 关键发现

### 1. 成功路径端到端成立，opt3 前提被实证

001 报告中 opt3 增量"未解锁"的前提是成功路径可建立缓存，当时无实证。本 demo 证明：**120 文件全量编译一次即建立 121 个 obj 缓存，且可复现**（两次 clean 全量均成功 RC=0）。001 预估的"成功路径改 1 文件 ≈ 前端 + 1 文件重编"逻辑成立。

### 2. 增量是 obj 级精确复用（时间戳实证）

`inc_chk`（改动 `M001.bas` 一行、语义不变）后检查 `.c3obj` 时间戳：**仅 `M001.obj` 更新**（20:46:18），其余 120 个 obj 保持上一轮全量时间（20:44:19）。缓存记录（源哈希+deps+opts）命中判定精确，改 1 文件只重编该文件并重链接 exe，其余 120 个 cl 编译被跳过。

### 3. optbench 上 /MP 与增量"看不出收益"——因模块太小、固定开销主导

- 热态全量 6.2s（/MP）≈ 7.0s（无 /MP）：cl 单文件仅 ~0.02s 级，进程/调度开销占主导，2 核并行无收益。
- 增量 6.2~7.3s ≈ 热态全量 6.2s：每次运行 C3 的固定开销（解析 120 模块 + 哈希 + vc 环境初始化）约 5~6s，被省掉的 cl 时间（~1s）在里面不可见。
- **结论：optbench 模块太小，不适合量化 /MP 与增量的时间收益**；这两项收益的锚点仍在 vbman 级工程（204 个大文件，cl 占 88%）：/MP -40%（1.67x）、增量成功后改 1 文件 -88%。optbench 的价值在**正确性/机制验证与可复现性**。

### 4. 首次构建冷启动开销不可忽略

`mp_full` 40.5s vs 热态 6.2s，相差约 34s 属一次性环境开销（vcvars 初始化、cl 首次加载、新文件防病毒扫描等）。日常迭代场景应保留已热身进程/缓存，冷态计时需复测确认（本报告对 mp_full 即以 `mp_full2` 复核，避免冷态误判为 /MP 负优化）。

## 复现步骤

```bat
:: 1. 重新生成任意规模工程（可选，默认已含 120 模块源码）
powershell -ExecutionPolicy Bypass -File gen_modules.ps1 -Count 120

:: 2. 成功全量构建（产出 out\OptBench.exe 与 out\.c3obj\ 121 obj）
build_optbench.bat

:: 3. A/B 计时（C3.exe / C3_nomp.exe 切换；-Clean 1 清缓存全量，0 走增量）
powershell -ExecutionPolicy Bypass -File bench_run.ps1 -Tag run -C3 ..\..\..\.build\C3.exe -Clean 1
```

注意：`bench.bat` 的 extra 参数位以 `%~2` 解引号后拼接，空参时不会向命令行残留空串（早期版本在 extra 为空时把 clean 标志错位拼进 C3 命令行，导致 vbp 被当作普通源码逐行解析——已修复）。
