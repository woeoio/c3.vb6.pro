# C3 编译器完整编译参数参考

> 本文档由 `C3 --help` 及 `src/driver/driver_args.cpp` 参数解析逻辑整理而成，
> 涵盖 C3 支持的全部命令行参数（开源后不再有任何隐藏参数）。

## 基本用法

```
C3 [选项] <源文件...>
```

- 源文件可为 `.bas` / `.vbp` / `.frm` / `.cls` 等
- 编译 `vbman` 等 ActiveX DLL 工程时**必须在 MSVC 环境中运行**
  （先执行 `vcvarsall.bat x64`，见 scripts/build.bat）

## 基本选项

| 参数 | 说明 |
|---|---|
| `-o <文件>` | 输出文件路径（DLL/EXE） |
| `--output-dir <目录>` | 输出目录（默认: 源文件所在目录） |
| `--target <平台>` | 目标平台: `win-x86` / `win-x64` / `linux-x64` / `macos-arm64` |
| `--arch <架构>` | 目标架构: `x64`（默认）或 `x86`（用于 32 位 COM 组件） |
| `--gui <模式>` | GUI 模式: `native` / `webview` / `none` |
| `--syntax-only` | 只做语法检查 |
| `-d, --define <N=V>` | 定义条件编译常量（如 `-d:DEBUG=-1`） |
| `-O <级别>` | 优化级别 0-3 |
| `-g, --debug` | 生成调试信息 |
| `--dll` | 编译为 ActiveX DLL（自动生成 TypeLib .tlb） |
| `-v, --verbose` | 详细输出 |
| `-h, --help` | 显示帮助 |
| `-V, --version` | 显示版本 |

## 调试/转储选项

| 参数 | 说明 |
|---|---|
| `--dump-tokens` | 输出词法分析后的 Token 流 |
| `--dump-ast` | 输出抽象语法树 |
| `--dump-symbols` | 输出符号表 |
| `--dump-preprocess` | 输出预处理后的源码 |
| `--dump-ir` | 输出中间表示 |
| `--dump-frm` | 输出 `.frm` 窗体描述 |
| `--emit-c` | 仅生成 C 代码，不编译链接 |
| `--emit-llvm` | 生成 LLVM IR |
| `--keep-for-debug` | **保留中间文件便于调试（不清理）** |
| `--compat-check` | 兼容性检查模式 |

### --keep-for-debug 说明（调试中间 C 代码的关键参数）

- 默认情况下 C3 编译结束后会**清理中间文件**
- 加 `--keep-for-debug` 后，生成的中间 C 文件（如 `xxx.c`、`dll_entry.c`、`test_prop.c`）
  及 RTL 头文件会保留在临时目录：`%TEMP%\C3C\<会话ID>\`
- 链接失败（LNK2019 等）时用该参数可检查生成代码与实际符号名是否一致

### --emit-c 说明

- 仅做 VB6 → C 的代码生成，不做 cl/link，速度最快，适合快速定位前端问题
- 生成的 `<模块名>.h` / `<模块名>.c` 写入中间目录 `%TEMP%\C3C\<会话ID>\`（模块名取
  `Attribute VB_Name`，**不是**源文件名）
- **同时把这两份内容原样 dump 到 stdout**。注意：`--output-dir` **不接收** C 文件，
  用 `--emit-c --output-dir X` 后 `X` 目录仍为空
- 因此程序化比对生成代码时，应哈希 **stdout**（或读中间目录），不要只看 `--output-dir`；
  否则会得到"两边都为空"的自洽假绿

## TypeLib/COM 选项

| 参数 | 说明 |
|---|---|
| `--typelib <文件>` | 显式引用 TypeLib |
| `--no-auto-typelib` | 禁用自动 TypeLib 加载 |
| `--progid <前缀>` | ActiveX DLL 的 ProgID 前缀 |
| `--libid <字符串>` | 显式指定 TypeLib 的 LibID |

## 性能/杂项选项

| 参数 | 说明 |
|---|---|
| `--incremental` | 增量编译（obj 级缓存，跳过未变化的 .c） |
| `--no-warn <ID列表>` | 抑制指定 ID 的警告，逗号分隔（如 `3001,3003`） |

> 跨模块 include 瘦身（只 include 本模块真正引用到的外部模块）**不是选项**，是编译器固有
> 行为（账 §56.52，用户 2026-11-12 裁定「副作用选项没啥好处」）。**debug 构建同样裁**
> （`-g` / `--debug` 不例外）：裁掉的是整行 `#include`，行号不变，被引用符号的原型与
> 类型照旧可见，PDB 里那些头的类型记录由它们自己那份 TU 提供 ⇒ 调试信息不受影响。
>
> 预编译头（PCH，账 §56.53）同样**不是选项**：编译时按前奏（普通档 / 窗体档）建两份
> 预编译头，各组 TU 用 `/FI` + `/Yu` 复用 `windows.h` / `vb6rtl.h` 的公共前奏。
> 它的产物（`c3_pch_std.h` / `c3_pch_form.h` 与对应 `.pch`）落在**中间目录**里，
> 随会话目录一起清理。

## 静态库选项

C3 通过 `Declare` 语句里的 `Lib` 子句自动区分动态/静态链接（详见 `ai/024-静态库链接计划书.md`）：

| `Lib` 取值 | 分类 | 行为 |
|---|---|---|
| 裸名 或 `xxx.dll` | 动态 | 沿用旧行为（`LoadLibrary` / 导入库） |
| `xxx.lib` | MSVC 静态库 | 参与静态链接 |
| `xxx.obj` | 目标文件 | 参与静态链接 |
| `xxx.a` / `xxx.o` | MinGW 静态库 | **本版未支持**（只有 MSVC 后端；会在编译期被拒收，见下） |

> **本版（v1）只支持 MSVC 后端的 `.lib`/`.obj`。** MinGW `.a`/`.o` 是另一套归档与符号表格式，
> 排在批次 T06（详见 `ai/024-静态库链接计划书.md`）。现在引用 `.a`/`.o` 会得到一条明确的
> 编译期拒收，而不是链接期一堆 `LNK1104`。

命令行选项：

| 参数 | 说明 |
|---|---|
| `--libdir <目录>` | 静态库**搜索根**，可重复；顺序即查找顺序 |
| `--extra-lib <库>` | **附加**静态库（`.lib`/`.obj`），可重复；用于没有对应 `Declare` 的链接依赖，例如静态库之间互相引用 |

`--libdir` 只负责"裸文件名去哪找"。完整寻址顺序（`Lib` 值 → 实际路径）：

1. **绝对路径** → 直接使用
2. **含路径分隔符**（如 `deps\x.lib`）→ 相对**工程目录**解析
3. **裸文件名** → 依次在这些根中查找：
   1. vbp 的 `LibDir=` 条目（按出现顺序）
   2. 命令行 `--libdir`（按给定顺序）
   3. `<工程目录>/Lib`（默认根，最后兜底）
4. 全都找不到 → 报错并**列出已搜索的全部根**（便于定位）

另外做**后端 × 格式**一致性检查（避免把错误丢给链接器）：MSVC 后端引用 `.a`/`.o`
会被拒收，MinGW 后端引用 `.lib`/`.obj` 同样被拒收，并提示应改用哪种扩展名。

### 怎么用：最小可运行例子

静态库成员是**按需拉取**的 —— 链接器只为"尚未解析的符号"去归档里取成员。由此三条语义：

- **引用了但从不调用的库不会报错。** 成员根本没被拉出来，连它的内部依赖都不必满足。
- **同一个 `.lib` 被多条 `Declare`（或同时被 `Declare` 与 `ExtraLib`）引用会自动去重**成一条链接输入。
- **只有"名字写了、文件哪儿都找不到"才报错。** `Declare` 的 `Lib` → 编译期 `VB5004`，并列出搜过的每个根；
  `ExtraLib=` 的裸名 → 链接期 `LNK1181`（系统导入库如 `ws2_32.lib` 本来就不在工程目录里，裸名必须放行给
  链接器按 `LIB` 环境变量与 `/LIBPATH` 去找）。

假设 `deps/` 下有 `mylib.lib`（导出 `my_add`）与它依赖的 `helper.lib`（**没有**对应 `Declare`）：

```vb
Declare Function MyAdd Lib "mylib.lib" Alias "my_add" (ByVal a As Long, ByVal b As Long) As Long
```

```bat
C3 工程.vbp --libdir deps --extra-lib helper.lib
```

`Lib "mylib.lib"` 由 `--libdir deps` 找到；`helper.lib` 没有 `Declare`，只能靠 `--extra-lib` 进链接
（少给这一条就会得到 `LNK2019` 未解析符号）。等价的 `.vbp` 写法见下。

**C 侧签名怎么对齐**（照下表写，x64/x86 都对；实测夹具即按此编写）：

| VB6 声明 | C3 生成的形参 | 你的 C 函数应写 |
|---|---|---|
| `ByVal x As Long` | `intptr_t` | `intptr_t`（x64 下 64 位；x86 下 4 字节，恰等于 VB6 Long） |
| `ByRef x As Long` | `int32_t*` | `int32_t*` —— **别写 `intptr_t*`**，x64 下会写坏调用者 |
| `ByVal s As String` | `char*` | `const char*` / `char*`，**ANSI**（不是 UTF-16） |
| `Sub` / `Function ... As Long` | `void` / `intptr_t` | 同左 |

⚠️ **返回 `As String` 仍是 VB6 语义（`BSTR`）**，与 `ByVal ... As String` 的 ANSI 口径不同。
C 函数返回 `char*` 时不要声明为 `As String`；请声明为 `As Long`（指针）再自行转换。

**调用约定**（x86 下必须一致，x64 下无关）：

- C3 的 `Declare` **默认 `__stdcall`**，与 VB6 一致（要 cdecl 得显式写 `Declare ... CDecl`）。
- 但 **MSVC 编译 C 源码的默认是 `__cdecl`**，不是 stdcall。所以给 VB6 用的静态库请**显式**写
  `__stdcall`（或整库加 `/Gz`），否则归档里的符号是 `_name`，而 C3 声明要的是 `_name@N` →
  **x86 下链接期报 `LNK2019: 无法解析的外部符号 _name@16`**。
- 库确实是 `__cdecl` 时，声明写成 `Declare Function F Lib "x.lib" Alias "name" CDecl (...) As ...` 即可。
- **x64 下 `__cdecl` 与 `__stdcall` 是同一套 ABI**（只有一个约定），所以上面这条只在 `--arch x86` 时产生后果。

> ⚠️ **签名写错，编译器多半不会替你发现 —— 请逐参精确对齐上表。** 实测完整矩阵：
>
> | 错误形态 | x86（`--arch x86`） | x64（默认） |
> |---|---|---|
> | 参数**个数**不一致 | ✅ `LNK2019: 无法解析的外部符号 _name@4` | ❌ 静默 → 每次运行值都不同 |
> | 参数**类型/宽度**不一致 | ❌ 静默（**有时甚至恰好算对**，更具欺骗性） | ❌ 静默 → 每次运行值都不同 |
> | **调用约定**不一致 | ✅ `LNK2019: 无法解析的外部符号 _name@16` | ➖ 无意义（x64 只有一套 ABI） |
>
> 原因是 x86 的 `_name@N` 把**参数字节数**编进符号名，而它只反映"占几个栈槽"，
> 不反映槽里装的是什么：`Integer` 与 `Long` 在 x86 上都只占 4 字节栈槽（`int16_t` 会被提升），
> 所以 `@N` 相同、链接器分不出。**x64 则完全不带这个后缀 —— 默认构建下一条护栏都没有。**
>
> 实测同一条错误声明（`ByVal a As Long` 误写成 `ByVal a As Integer`，C 侧是 `long`）：
> x86 得 203（巧合），x64 连跑 5 次得到 5 个互不相同的垃圾值。
>
> 想要一个**会报错**的安全网，只能靠 x86 的两条（元数错、调用约定错）。
> 结论：静态库这条路，**VB6 声明的每个参数类型都必须与库的 C 原型严格一致，这靠人，不靠工具。**

### 工程文件中的对应键

`.vbp` 里支持与环境无关的写法，两者都支持一行多项、以 `;` 分隔：

```ini
LibDir=<工程相对目录1>;<工程相对目录2>
ExtraLib=deps\third_party.lib;deps\helper.obj
```

`LibDir=` 的相对路径以**工程文件所在目录**为基准。

### DeclareWide：禁用 ByVal String 的 ANSI 转换

VB6 的 `Declare` 对 `ByVal x As String` 默认生成 ANSI 字符串（`vb6_BSTR_ToANSI` +
临时缓冲 + `vb6_FreeANSI`），适配绝大多数 Win32 `A` 版 API。但调用 `W` 版 API 时
需要传入**原始宽字符串**，此时用 `DeclareWide` 关闭该转换：

```vb
' 默认 Declare：x 被转成 char*（ANSI）
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal s As String) As Long

' DeclareWide：s 直接以 BSTR 传入，不做 ANSI/Unicode 转换
Private DeclareWide Function lstrlenW Lib "kernel32" (ByVal s As String) As Long
```

效果对比（`--emit-c`）：

```c
/* 默认 Declare */
char* _ansi_0 = vb6_BSTR_ToANSI(s);
n = lstrlenW(_ansi_0);
vb6_FreeANSI(_ansi_0);

/* DeclareWide */
n = lstrlenW(s);
```

`DeclareWide` 只影响 `ByVal ... As String` 参数的转换，其余解析规则与 `Declare` 完全一致。

## 常用示例

```bat
REM 编译 EXE 工程
C3 工程.vbp --output-dir <dir>

REM 编译 ActiveX DLL 工程（vbman 等，须在 vcvarsall x64 环境）
C3 --dll --arch x64 -o <输出目录>\xxx.dll 工程.vbp

REM 保留中间 C 文件以便调试链接错误
C3 --dll --arch x64 --keep-for-debug -o xxx.dll 工程.vbp

REM 仅生成 C 代码（快速验证，不链接）
C3 --emit-c --output-dir <dir> 工程.vbp

REM 转储中间产物
C3 --dump-tokens --dump-ast 工程.vbp

REM 引用静态库: 追加搜索根 + 无 Declare 的附加静态库
C3 工程.vbp --libdir third_party --libdir deps/win64 --extra-lib deps/helper.lib
```

## 中间文件位置

- 中间 C 文件目录：`%TEMP%\C3C\<会话ID>\`
- `--output-dir` 仅用于最终 DLL/EXE 及错误日志（`c3-error.log`）
- 例外：`--emit-c` 的 `<模块名>.h/.c` 只落中间目录 + stdout，**不落** `--output-dir`（见上文）
