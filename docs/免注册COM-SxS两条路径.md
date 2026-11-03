# 免注册 COM：SxS manifest 与 LoadLibrary 直加载（待办 / 设计说明）

> 状态：**已记录，未开工**（2026-11-02）。当前实现仍是 `LoadLibraryW + DllGetClassObject` 直加载，
> 不引入 `CoCreateInstance`。本文只定义后续要补齐的两条路径与验收口径。

## 目标

manifest 路径与直加载路径**都要能正确运行**，不是二选一：

1. **manifest 优先**：进程级嵌入清单或运行时 `ActivateActCtx` 激活上下文，让 SxS 解析 OCX/DLL。
2. **直加载回退**：manifest 不可用（identity 取不到 / 系统不支持）时，仍走现有
   `LoadLibraryW + DllGetClassObject`。
3. 选择策略显式化：`-ComMode manifest|direct|auto`，默认 `auto`（manifest 优先、直加载回退）。

## manifest 路径硬约束

- 必须使用真正的 `<dependentAssembly>` + `<assemblyIdentity>`，**identity 只能来自可核实来源**：
  OCX 的 VS_VERSIONINFO（`FILEVERSION`/`PRODUCTVERSION`/`OriginalFilename`）、
  内嵌 typelib 的 `LIBID`/版本、或发行方公开的装配清单。**禁止硬编猜测值**——猜错只会得到
  `ERROR_SXS_*`，比不做更糟。
- 需要的证据（缺一不可，记进每条 OCX 的待办）：
  - `name`（装配名，通常为 OCX 文件名去扩展名或发行方清单名）
  - `version`（四段版本号，来自版本资源或 typelib）
  - `processorArchitecture`（x86/x64/amd64，必须与目标进程一致）
  - `publicKeyToken`（有签名时；无签名则留空，不得臆造）
  - `type`（通常 `win32`）
  - 对应 `.manifest` 文件与 OCX 的落盘位置关系
- 激活时机：需要时在 `OleInitialize` **之前** `CreateActCtx` + `ActivateActCtx`；
  进程级嵌入清单（linker `/MANIFEST`）与运行时 ActCtx 至少覆盖一种可用。

## 诊断（失败必须可区分）

| 现象 | 判据 |
|---|---|
| SideBySide | `GetLastError` 落在 `ERROR_SXS_*`（如 14001 / 14002 / 14003） |
| 文件缺失 | 目标 OCX/DLL 不在 exe 同目录或搜索路径 |
| 位数不匹配 | PE `Machine` 与进程位宽不一致（0x014C vs 0x8664） |
| typelib 缺失 | `LoadTypeLibEx(REGKIND_NONE)` 失败或 `GetLibAttr` 无 LIBID |

禁止统一报"加载失败"。

## 测试与 CI

- 测试入口加 `-ComMode manifest|direct|auto`，**同一套断言在两种模式下都跑**，不得各写一套期望。
- CI 至少覆盖 `direct`；`manifest` 在拿到可靠 identity 证据后纳入。
  取不到 identity 就报 **SKIP**，并写明缺哪个 identity（哪个 OCX、缺 name/version/arch 中的哪一项）。

## 预计涉及文件（开工时再动）

- `src/rtl/core/vb6com/vb6com.c`：现有 `LoadLibrary + DllGetClassObject` 路径保留，新增 manifest 分支
- `src/driver/*`：进程启动早期的 ActCtx 激活点（`OleInitialize` 之前）
- `tests/run_tests.ps1`：`-ComMode` 参数与同一套断言
- `.github/workflows/ci.yml`：direct 模式必跑，manifest 模式按 identity 证据决定纳入或 SKIP
- `tests/**/*.manifest`：按 OCX 逐一取证的清单文件
