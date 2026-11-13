# scripts 目录说明

构建 / 测试 / 编译辅助脚本的集合。脚本清单总览见项目根 [README.md](../README.md) 第五节「脚本工具（`scripts/`）」。

## 可移植环境变量与自动探测

脚本历史上硬编码了特定开发机的路径（vcvarsall 位置、项目根目录），现已改为**环境变量优先、缺省自动探测**：

| 环境变量 | 作用 | 未设置时的默认行为 |
|----------|------|------------------|
| `C3_VCVARSALL` | `vcvarsall.bat` 的完整路径（想锁定特定 VS 版本时设置） | 通过 `vswhere` 自动探测本机 VS 实例（兼容 Community / Professional / Enterprise / BuildTools 及多实例共存），探测不到则报错退出 |
| `C3_PROJECT_DIR` | 项目根目录（含 `CMakeLists.txt` 的目录） | 脚本所在目录的上一级（`%~dp0..` / `Split-Path -Parent $PSScriptRoot`），一般无需设置 |

> 一般情况下**无需设置任何环境变量**：`build.bat` / `run_tests.ps1` / `env.ps1` 会用 `vswhere` 自动定位本机 Visual Studio，在任意开发机 / CI 环境零配置可用。环境变量仅用于显式锁定非默认工具链。

### 使用示例（锁定特定 VS 实例时才需要）

**cmd（当前会话临时生效）：**

```bat
set "C3_VCVARSALL=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat"
scripts\build
```

**PowerShell（当前会话临时生效）：**

```powershell
$env:C3_VCVARSALL = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat"
scripts\dev.ps1 -SkipTest
```

**持久化（用户级，写入注册表，新开终端生效）：**

```bat
setx C3_VCVARSALL "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat"
```

> 上例路径为 **Build Tools 版** VS2022 的 vcvarsall 位置；完整 IDE（Community / Professional / Enterprise）对应 `C:\Program Files\Microsoft Visual Studio\2022\<Edition>\VC\Auxiliary\Build\vcvarsall.bat`。

### 附注：CMake / Ninja 不在 PATH 时

`build.bat` 已内置兜底：`cmake` / `ninja` 不在 `PATH` 时，自动遍历本机所有 VS 实例，把完整 IDE 捆绑的 CMake / Ninja（「C++ CMake tools for Windows」组件）追加到 `PATH`。**无需手动处理。**

若本机只装了 VS Build Tools（不含捆绑 CMake）且 CMake 不在 PATH，可独立安装 CMake ≥ 3.20 与 Ninja ≥ 1.12 并加入 PATH，或设置环境变量后手动追加：

```bat
set "PATH=%PATH%;C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin;C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\Common7\IDE\CommonExtensions\Microsoft\CMake\Ninja"
```

## pr.bat —— 一键创建 Pull Request（只建不合）

`main` 是保护分支，本地无法直接 `git push origin main`，必须走合并请求。
`pr.bat` 把「推送 + 汇总提交 + 建 PR」做成一条命令，**只负责创建，绝不自动合并/审批**，
必须由人在 GitCode 网页审核后手动点合并。

| 参数 | 说明 |
|------|------|
| `-b 分支` | base 分支，默认 `main` |
| `-s 分支` | head 分支，默认当前分支 |
| `-t 标题` | PR 标题，默认取唯一提交的标题；多提交时为 `merge <head> into <base>` |
| `-p, --push` | 创建前先 `git push -u origin <head>`（origin 与本地不同步时**必须**加） |
| `-o, --open` | 创建后在浏览器打开 PR 页面 |
| `-n, --dry-run` | 预演：只打印将生成的标题/正文，不调 API |
| `-h, --help` | 帮助 |

```bat
scripts\pr.bat              REM 当前分支 -> main
scripts\pr.bat -p -o        REM 先 push，建完自动打开浏览器
scripts\pr.bat -n           REM 先看一眼会生成什么
```

行为约定：

- 前置检查：head 不能等于 base；head 必须在 origin 上存在且与本地一致（否则提示加 `-p`）
- 无新增提交（`base..head` 为空）直接报错，不建空 PR
- 已存在同源同目标的 open PR 时，直接打印其链接并退出，不重复创建
- 工作区有未提交改动只告警，不阻断（未提交内容不会进入 PR）

令牌（`scripts\pr.ps1`）取值顺序：`GITCODE_TOKEN` 环境变量 → `scripts\.gitcode_token` 文件首行 →
git 凭据管理器（`gitcode.com`，`oauth2` 用户）。

> 坑记录：GitCode PR 接口不接受含非 ASCII 字节的请求体（报 `Parameter description field contains
> invalid byte sequence`），`pr.ps1` 会把 JSON 中 > U+007F 的字符统一转义成 `\uXXXX` 后再发送。

## PD 整编性能线（§56.53–§56.55）

这一组三个脚本是 **PhotoDemon 整编性能主线**的工作流固化版。
PD 整编（623 个项目 TU + 65 个 RTL `.c`）在 CI 上**必然编不过**（PD 源码本身带真实错误），
GA 判的不是"编过"，而是**错误面逐码相等** + **总时间下降**。
当前成绩与完整台账见 [`ai/C3_PD_BUILD_PERF_HANDOFF.md`](../ai/C3_PD_BUILD_PERF_HANDOFF.md)。

> **这三个脚本含中文，必须用 `pwsh`（PowerShell 7）跑**：`powershell`(5.1) 会按 GBK 误读
> UTF-8 无 BOM 的源码，引号配对会错乱。

### `ga_dispatch_pd.ps1` —— 派发 PD 整编 GA

同步探针分支 → 推 **github**（不是 gitcode）→ 派发 `PhotoDemon x86 full build` →
**按 headSha 认领 run** → 等结论。全程带 TLS 重试（hosts 把 github.com 指向 127.0.0.1，
本地转发时有时无）。

```powershell
pwsh -NoProfile -File scripts\ga_dispatch_pd.ps1                        # 只派 PD 整编
pwsh -NoProfile -File scripts\ga_dispatch_pd.ps1 -AlsoDevRegression     # 连带 dev 回归那 22 格
pwsh -NoProfile -File scripts\ga_dispatch_pd.ps1 -NoWait                # 派完就走
```

| 参数 | 说明 |
|------|------|
| `-Ref` | 探针分支，默认 `probe/photodemon-emit` |
| `-ProbeWorktree` | 探针 worktree，默认 `.temp\probe_sync`（另一个 clone） |
| `-DevRef` / `-AlsoDevRegression` | dev 回归的分支与开关 |
| `-SkipProbeSync` | 不做同步，按当前 HEAD 直接派 |
| `-NoWait` | 派完不等待 |

run id 会写进 `.temp\pd_run.txt`（`pd_readout.ps1` 默认读它）。

### `pd_readout.ps1` —— 一次读齐这一刀的关键数

取工件 `photodemon-full-build-logs` ⇒ 打印 `pd_build.meta`（`BUILD_SECONDS` /
`MIN_FREE_PHYS_MB` / `PEAK_TREE_WS_MB`）、`C3: MEM` 行、**cl 四段耗时**、内存曲线，
最后按 `-Baseline` 跑一次错误面逐码比对。

```powershell
pwsh -NoProfile -File scripts\pd_readout.ps1 -Run 43224386890 -Baseline 43218740715
pwsh -NoProfile -File scripts\pd_readout.ps1            # 不给 -Run 就读 .temp\pd_run.txt
```

### `diff_c3error.ps1` —— 错误面**逐码**比对（判据唯一口径所在）

口径只设这一处：正则 `(fatal error|error|warning)\s+(C\d{4})\s*:`。
别的脚本 / 别的台账要判"错误面变了没有"，都必须调它，**不要另起一份正则**。

```powershell
pwsh -NoProfile -File scripts\diff_c3error.ps1 <oldRun> <newRun>
pwsh -NoProfile -File scripts\diff_c3error.ps1 -OldLog a.log -NewLog b.log
```

退出码：`0` = 逐码相等（通过）/ `1` = 有差异 / `2` = 日志找不到（**没判**，不等于通过）。

> **判据只认逐码计数**，不得把「恰好 N 处」降级成「至少 N 处」；
> 也**不得声称字节数相等** —— 已实证不可比：cl 的 `/MP` 调度顺序一变，
> 同一批 `compiling source file` 行就换位置，加上 session 目录名长度不同（16 vs 13 字符），
> 字节数就会差上百 KB，而内容逐行看是等价的。
