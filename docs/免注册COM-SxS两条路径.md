# 免注册 COM：外部/合成清单（SxS）与 LoadLibrary 直加载

> 状态：**已实现，本机验收通过**（2026-11-04）。本文记录本轮实测结论、C3 的清单合成规则、
> 诊断口径与夹具判据。此前的"必须有 `<dependentAssembly>` + 厂商 identity"口径已被实测推翻。

## 一、结论（本机 Windows 10.0.26200，2026-11-04 实测）

SxS 认的是**进程级清单**（EXE 内嵌的 RT_MANIFEST，或运行期 `CreateActCtxW` 传入的那份文件）里的
`<file name="...">` + `<comClass .../>`，**不需要 `<assemblyIdentity>`、不需要 `<dependentAssembly>`**；
而**单放在组件旁边的 `<组件文件名>.manifest` 不会被 SxS 自动认作程序集清单**——它只是素材，
必须并进进程清单（或交给 `CreateActCtxW`）才生效。

参考物（工具包）：`D:\vb_yqt4qPac\免注册提升管理员工具包Dll_OCX\DLL_OCX\清单\MSDatGrd.ocx.manifest`
（2531 B）、`TabCtl32.Ocx.manifest`（1887 B）。它们的结构是：

```xml
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <file name="DLL_OCX\MSDatGrd.ocx" hashalg="SHA1">
    <comClass clsid="{...}" tlbid="{...}" description="..."></comClass>
    <typelib tlbid="{...}" resourceid="1" version="1.0" helpdir=""></typelib>
  </file>
  <comInterfaceExternalProxyStub name="..." iid="{...}" tlbid="{...}" proxyStubClsid32="{00020424-...}"></comInterfaceExternalProxyStub>
</assembly>
```

要点：**没有 `<assemblyIdentity>`**、**没有 `requestedExecutionLevel`**、`<file name>` 是
**相对清单所在目录**的路径（工具包清单放在 `...\清单\`，但路径按 exe/清单目录写成 `DLL_OCX\...`，
即实际生效的是并入 EXE 后的那份），`<comClass>` 可只带 `clsid`/`tlbid`/`description`；
要按 ProgID 建对象时再补 `progid`（工具包没写 progid，因为 VB6 的 `Object=` 控件按 CLSID 建）。

### 实测读数

| 场景 | 读数 |
|---|---|
| 外部 `<组件>.manifest` 单放、不激活 ActCtx | `CLSIDFromProgID` = `0x800401F3` (CO_E_CLASSSTRING)，`CoCreateInstance` = `0x80040154` (REGDB_E_CLASSNOTREG)，模块未加载 |
| 同一份内容作为进程清单（无 identity，`<file>`+comClass） | `CreateActCtx` = OK，`CoCreateInstance` = OK，loaded-module 指向清单同目录的副本 |
| 同一份内容作为进程清单（带合成 identity） | `CreateActCtx` = OK，`CoCreateInstance` = OK，**无 14001** |
| `<file name>` 写绝对路径 | `CreateActCtx` = 14001（ERROR_SXS_CANT_GEN_ACTCTX） |
| Codejock.CommandBars.v20.3.0.ocx（x86，VS_VERSIONINFO 20.3.0.0）+ 代写 per-file 清单 | 32 位进程：`CreateActCtx` = OK、`CoCreateInstance` = OK、loaded-module = 我们目录里的副本（SxS 映射压过注册表） |
| 客户端请求 `2.0.0.0`、私有程序集 identity 是 `1.0.0.0` | `CreateActCtx` 失败，错误码落在 `ERROR_SXS_*`（14001 一族），manifest 模式点名后退出 |

补充：`CODEJOCK 全家` 的 OCX 有 RT_MANIFEST #2，但只有 `<trustInfo>`（UAC），**没有
`<assemblyIdentity>`**；`RICHTX32.OCX`、`NewTab01.ocx` 连 RT_MANIFEST 都没有。这不影响外部清单口径：
C3 不依赖厂商 identity。

## 二、C3 的两条路径

- `direct`（默认行为的既有路径，**一行未改**）：`LoadLibraryW + DllGetClassObject + IClassFactory::CreateInstance`。
- `manifest`：进程启动早期 `CreateActCtxW` + `ActivateActCtx`（在 `CoInitializeEx`/`OleInitialize` 之前），
  之后 `CLSIDFromProgID + CoCreateInstance` 由 SxS 解析到清单里声明的本地组件。

模式开关：`-ComMode direct|manifest|auto`（cgen 只在显式给旗时注入一行；另有 `C3_COM_MODE` 等价注入）。
`-ComManifest <路径>` / `C3_COM_MANIFEST` 可指定现成清单。

### manifest 模式的清单选择顺序

1. 显式清单（`-ComManifest` / `C3_COM_MANIFEST`）——优先，校验后直接 `CreateActCtx`；
2. exe 旁的 `<exe 名>.manifest`；
3. 都没有、且工程声明了 `ComLib=` / `Object=` 组件 ⇒ **运行期合成一份聚合清单**，写到 exe 目录旁
   （文件名 `c3com.<pid>.<tick>.manifest`，每次唯一，避开 SxS 按清单路径缓存解析结果的坑）。

### 合成规则（`vb6com.c` 的 `comComposeManifest`）

1. 组件来源：`ComLib=` 表 + `Object=` 的 OCX 表；按**组件文件路径**去重——
   同一个 DLL 声明多个 coclass 时只能有**一个** `<file>`，`<comClass>` 都挂它下面
   （重复 `<file name>` 块会导致整份清单解析不出 ProgID 映射）。
2. `<file name>`：取组件路径相对 **exe/清单目录** 的形式；绝对路径只有落在 exe 目录下才压得成相对，
   压不成就**点名失败**（`ERROR_SXS*` 14001 的根因就是绝对路径）。
3. 优先复用现成的 `<组件文件名>.manifest`：抽出它的 `<file>...</file>` 块，仅把 `name` 属性改写成
   目标相对路径，其余（`typelib`、`comInterfaceExternalProxyStub`、`description`）原样保留。
   抽不出/改写失败才退回合成。
4. 合成块：`<file name="..." hashalg="SHA1"><comClass clsid="..." threadingModel="Apartment" progid="..."></comClass></file>`；
   缺 CLSID 的条目点名失败，不写半条。
5. **不合成 identity**：实测与参考物都表明进程清单不需要 `assemblyIdentity`（有/无都成功）。
   若确需一份带 identity 的外部清单（人工用途），取值来源如下，并注明：
   - `name`：组件文件名（含扩展名）；
   - `version`：组件 `VS_VERSIONINFO` 的 `FILEVERSION` 四段（如 20.3.0.0）；
   - `processorArchitecture`：PE 头 `Machine`（0x014C → `x86`，0x8664 → `amd64`）；
   - `publicKeyToken`：`null`（无签名时）；
   - **这组 identity 是工装合成，不是厂商签名声明**，不得当作发行方身份使用。

### 诊断（失败必须点名，不许静默回退）

| 现象 | C3 输出 |
|---|---|
| 清单文件不存在/不可读/无 `<file>+<comClass>` 与 `<dependentAssembly>` | `manifest mode FAILED -- missing: ...` + 退出码 1 |
| 合成失败（组件不在 exe 目录、文件不存在、缺 CLSID、空清单） | `manifest mode FAILED -- missing: <组件名: 具体缺项>` |
| SxS 拒绝（版本/架构不匹配等） | `CreateActCtx failed: error 14001 (ERROR_SXS_CANT_GEN_ACTCTX ...)` |
| ProgID 查不到 | `CreateObject` 抛 VB6 错误 429，`Err.Description` 带 ProgID 与 HRESULT（`0x800401F3/0x80040154`） |
| `auto` 回退 | `[C3_COM] auto: fallback to direct -- <原因>`（写明是清单不可用还是 ActCtx 建失败） |

`C3_COM_KEEP_MANIFEST=1` 可保留合成清单便于人工核对；`C3_COM_TRACE=1` 打印
`[C3_COM] using manifest (ActCtx activated): <路径>`。

### 释放

唯一释放点 `vb6_ComExit()`：`DeactivateActCtx(0, cookie)` → `ReleaseActCtx(h)`，随后删除本次合成的清单。
非正常退出（`ExitProcess`/崩溃）由操作系统回收。

## 三、夹具判据（`tests\com_manifest\run_manifest_tests.ps1`）

1. **ProgID 免注册**：完全未注册的 `test_activex_dll.dll` + 进程清单（`<file>`+comClass）
   ⇒ `CreateObject("TestAXDLL.Calc")` 成功；同工程 `direct` 模式必须失败（对照）。
2. **外部清单代写**：组件旁放 `<组件>.manifest`（声明另一个 ProgID `ExtManifest.Calc`），
   **不给 `-ComManifest`** ⇒ 运行期合成后 `CreateObject("ExtManifest.Calc")` 成功（证明外部清单被消费）。
3. **合成清单**：无外部清单、只有 `ComLib=` ⇒ 合成后 `C3_COM_PREFER_REG=1` 强行走注册表/ SxS 路径仍成功。
4. **反向**：客户端 `dependentAssembly` 请求 `2.0.0.0`、私有程序集是 `1.0.0.0`
   ⇒ manifest 模式点名失败（`manifest mode FAILED` + `error 14xx`），不许假绿。
5. **直加载不回归**：`ComLib= + -ComMode direct` 成功；`test_activex_dll` 的 `dll_entry.c` 形状断言
   与 b595/b597 输出逐行不变。
6. 每次运行用**唯一目录 + 唯一清单文件名**（SxS 按客户端清单路径缓存 ActCtx，换内容不换路径会读到旧映射）。
