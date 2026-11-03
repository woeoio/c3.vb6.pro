# crashprobe

最小启动型调试器（`CreateProcessW` + `DEBUG_ONLY_THIS_PROCESS`），用来在"先复现再修"的取证环节抓：

- 异常码 / first-chance 标记 / 异常地址
- `accessType` 与**故障访问地址**（AV 时是读还是写、读/写了哪个地址）
- faulting module + 模块内 offset
- 调用栈（模块内地址）
- 一份与目标同位宽的 minidump

它**不是**通用调试器，不接 CI，只服务于本仓的崩溃取证。

## 1. 两个二进制 + 一个包装

| 文件 | 位宽 | 用途 |
|------|------|------|
| `crashprobe.exe` | x64 | 调试 64 位目标 |
| `crashprobe32.exe` | x86 | 调试 32 位目标（**C3 默认产物就是 x86**） |
| `crashprobe.ps1` | — | 薄包装：读目标 PE 头的 `machine` 字段，自动选上面两个之一 |

两个 exe 由**同一份** [crashprobe.cpp](crashprobe.cpp) 编译（见 [build.bat](build.bat)），差异只有：

```cpp
#if defined(_M_X64)      // CONTEXT 用 Rip/Rbp/Rsp, StackWalk 用 AMD64
#elif defined(_M_IX86)   // CONTEXT 用 Eip/Ebp/Esp, StackWalk 用 I386
```

**不要拿 x64 版去调 WOW64（32 位）目标**：那需要 `WOW64_CONTEXT`、`Wow64StackWalk` 和混合
dump，本工具刻意不做那一档 —— 位宽不符时栈和 dump 会"半对半错"，比抓不到更误导。
调试器与目标必须同位宽，dump 位宽随之与目标一致。

## 2. 用法（只调包装）

```powershell
# 最常用：第二个参数省略时，dump 落在目标同目录、同名 .dmp
powershell -NoProfile -ExecutionPolicy Bypass -File tools\crashprobe\crashprobe.ps1 `
    -Target .\B5b.exe

# 指定 dump 路径
powershell -NoProfile -ExecutionPolicy Bypass -File tools\crashprobe\crashprobe.ps1 `
    -Target .\B5b.exe -DumpPath .\B5b.dmp

# pwsh 下位置参数也认
pwsh -File tools\crashprobe\crashprobe.ps1 .\B5b.exe .\B5b.dmp
```

直接调 exe（已知目标位宽时）：

```powershell
tools\crashprobe\crashprobe32.exe .\B5b.exe .\B5b.dmp
tools\crashprobe\crashprobe.exe   .\x64app.exe .\x64app.dmp
```

### 位宽自动选择规则

包装只读 PE 头，不猜、不探测：

1. `MZ` @0x00；`e_lfanew` @0x3C；`PE\0\0` @`e_lfanew`；`machine` @`e_lfanew+4`。
2. `machine=0x014C`（`IMAGE_FILE_MACHINE_I386`）→ `crashprobe32.exe`
3. `machine=0x8664`（`IMAGE_FILE_MACHINE_AMD64`）→ `crashprobe.exe`
4. 其它 machine（ARM64 等）→ **直接拒绝**，报"不认识的 PE"。

### 包装退出码

| 退出码 | 含义 |
|--------|------|
| `0` | 目标正常退出（`EXIT_CODE=0`） |
| `2` | 目标文件不存在 |
| `3` | 不是能识别的 PE / PE 头损坏 |
| `4` | 缺对应位宽的工具二进制（先跑 `build.bat`） |
| 其它 | = crashprobe 退出码 = **目标进程退出码**（异常终止时即异常码，如 `0xC0000005`） |

## 3. 输出字段含义

```
ARCH=x64|x86                    本次使用的工具位宽（应与目标 PE machine 一致）
PID=...                         被调试进程 ID
EXCEPTION code=0x.........       异常码 (NTSTATUS)；0xC0000005=AV, 0x80000003=断点, ...
          firstChance=1|0       1=首次机会（目标自己的 SEH 还能处理）, 0=二次机会
          addr=...              异常指令地址 (ExceptionRecord.ExceptionAddress)
  accessType=0|1|8              AV 的访问类型: 0=读, 1=写, 8=执行
  faultAddr=...                 触发 AV 的那个**数据地址** (读/写/执行的目标)
FAULT_MODULE=<dll 全路径> base=... offset=0x...
                                addr 落在哪个模块 + 模块内偏移（无符号, 可比对 map/PDB）
STACK:                          16 帧，模块内绝对地址；本仓 EXE 无 PDB 时对照同名 .map
  #00 ...
DUMP=<路径>                     minidump 落盘路径（位宽 = 目标位宽）
EXIT_CODE=0x........ (...)      目标进程退出码；异常终止时 = 异常码（括号内是 signed 十进制）
```

说明：

- `EXCEPTION` 可能打印多次（首次机会 + 二次机会）。**判断用第一条 `0xC0000005` 的
  `firstChance=1`** —— 二次机会那条的 `addr/faultAddr` 与它相同，只是重复。
- `accessType=0` 且 `faultAddr=0x8` 是典型的"空指针 + 偏移"取证形态（例如把裸实例当
  `IDispatch` 交给 OLEAUT32）。
- 目标自己的 `printf`/控制台输出会**穿插**在工具输出里（子进程继承同一控制台），
  例如上面的 `S1 assigned` —— 不要把它当成工具的字段。

## 4. 退出码 = 异常码

`crashprobe` 的 `main` 返回 `EXIT_CODE`，因此：

- 目标 AV 退出 → 工具也以 `0xC0000005` 退出（PowerShell 里 `$LASTEXITCODE = -1073741819`，
  Windows 进程退出码是 `3221225477`）。**这个非零值本身就是取证结果**，脚本应判"抓到
  AV"而不是"工具失败"。
- 目标正常退出 → 工具以目标退出码退出；`EXIT_CODE=0` 才是"没抓到异常"。

抓到首个 AV 后，工具在 **second-chance** 阶段主动 `TerminateProcess(异常码)`，避免无头
环境里 WER/系统对话框接管导致挂死。

## 5. 可靠性前提：本机 WER 已禁用或不可用

工具只保证在**本机 Windows Error Reporting 已被禁用或不可用**时，抓到的是目标那一发：

```
HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting\Disabled = 1
WerSvc 服务 = Disabled / Stopped
```

WER 可用时，系统会接管第二机会异常（甚至抢先做转储/重启），拿到的可能不是目标异常。
在 CI/共享机上不要把本工具当作 WER 的替代品。

## 6. 真实输出样例

### x86 目标（32 位，C3 默认产物）

```
[crashprobe] target=...\.temp\m4\b5b\x86_pre\B5b.exe bits=x86 tool=crashprobe32.exe
ARCH=x86
PID=10012
S1 assigned
S2 before Add
EXCEPTION code=0x80000003 firstChance=1 addr=77CA8498
EXCEPTION code=0xC0000005 firstChance=1 addr=7754CF77
  accessType=0 faultAddr=00000004
FAULT_MODULE=C:\WINDOWS\System32\OLEAUT32.dll base=77530000 offset=0x1CF77
STACK:
  #00 7754CF77
  #01 77556CB7
  #02 77554479
  #03 77554D01
  #04 7150E6E0
  #05 003E915D
  ...
DUMP=...\b5b_x86_pre.dmp
EXCEPTION code=0xC0000005 firstChance=0 addr=7754CF77
  accessType=0 faultAddr=00000004
EXIT_CODE=0xC0000005 (-1073741819)
```

### x64 目标

```
[crashprobe] target=...\.temp\m4\b5b\prefix_out\B5b.exe bits=x64 tool=crashprobe.exe
ARCH=x64
PID=6336
S1 assigned
S2 before Add
EXCEPTION code=0x80000003 firstChance=1 addr=00007FFC7C17F58D
EXCEPTION code=0xC0000005 firstChance=1 addr=00007FFC7B3BFF30
  accessType=0 faultAddr=0000000000000008
FAULT_MODULE=C:\WINDOWS\System32\OLEAUT32.dll base=00007FFC7B320000 offset=0x9FF30
STACK:
  #00 00007FFC7B3BFF30
  #01 00000196A53A0000
  ...
DUMP=...\b5b_x64_ref.dmp
EXCEPTION code=0xC0000005 firstChance=0 addr=00007FFC7B3BFF30
  accessType=0 faultAddr=0000000000000008
EXIT_CODE=0xC0000005 (-1073741819)
```

## 7. 构建

```bat
tools\crashprobe\build.bat
```

- 自动找 `vcvarsall.bat`：环境变量 `VCVARSALL` → `vswhere` → 已知 VS2019 路径。
- 先 `vcvarsall x64` 编 `crashprobe.exe`，再 `vcvarsall x86` 编 `crashprobe32.exe`。
- 需要 `dbghelp.lib` + `psapi.lib`（两位宽都有）。

## 8. 已知限制

1. **不接 CI**：只在本地取证时手工调用。
2. 只支持 x86 / x64 目标；ARM64 等直接拒绝。
3. 栈只有模块内地址；要符号请对照 `.map`（本仓不产 PDB）。
4. 只抓第一个 `0xC0000005`；其它异常（`0xC0000094` 除零、`0xC00000FD` 栈溢出等）打到
   stdout 但**不写 dump**（按需再扩）。
5. 依赖本机 WER 状态（见 §5）。
