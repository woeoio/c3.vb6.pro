# C3 — 现代化 VB6 编译器

> 官网：<https://c3.vb6.pro> ｜ 仓库：<https://gitcode.com/woeoio/c3.vb6.pro>（本地路径 `D:\code\vi\c3.vb6.pro`）
>
> C3 把 VB6 工程（`.vbp` / `.bas` / `.cls` / `.frm` / `.frx`）直接编译为**原生 Windows x64 / x86 可执行文件或 ActiveX DLL**，
> 源码无需任何修改，无运行时 DLL 依赖。

---

## 一、这是什么

C3 是一个用 C++17 从零手写的 VB6 编译器前端 + C 代码生成后端：

```
VB6 源码 (.bas/.cls/.frm/.frx + .vbp)
   ↓ 编码检测 (GBK / UTF-8 / UTF-16)
   ↓ 预处理   #Const / #If / #ElseIf / #Else
   ↓ 词法     LL(2) + Unicode
   ↓ 语法     递归下降 + Pratt（14 级优先级）
   ↓ 语义     两遍扫描 + 符号表 + 类型系统 + TypeLib 导入
   ↓ 代码生成 VB6 → C
   ↓ 编译链接 shell-out 到 MSVC 的 cl.exe / link.exe
原生 EXE / DLL（静态链接 VB6 运行时）
```

**核心信条：不支持的语法必然报错，绝不静默误编。** 错编译比不编译更危险。

### 产品定位

| 维度 | 说明 |
|------|------|
| 定位 | 旧 VB6 项目迁移工具 + 用 VB6 语法写新项目的工具 |
| 形态 | 编译器 + 工具链内核（CLI），**不做 IDE** |
| 开源 | 是 |
| 与 twinBASIC 的关系 | 不竞争。twinBASIC 是闭源一体化 IDE，C3 做开源底层，对接任意前端 |
| 目标用户 | 需要迁移旧项目、偏好开源的开发者 |

---

## 二、当前状态（2026-09-14）

| 项目 | 状态 |
|------|------|
| 版本 | 权威值见根目录 `VERSION` 文件（`C3 --version` 与 CMake 工程版本均由它自动派生） |
| 里程碑 | M1 ~ M33 全部达成 |
| 最新基线 tag | `v0.1.0-m33-vbman-dll`（提交 `5f4bed1`） |
| 编译期 | 大型真实工程（VBMAN，125 模块含 4 个窗体）**error C = 0**，链接期 LNK2019 / LNK2005 / LNK1104 均已清零 |
| 产出验证 | VBMAN.dll 2,566,144 字节，可 `regsvr32` 注册（运行期行为仍在对齐中） |
| 内置函数覆盖 | 143 个 VB6 内置函数中已实现 126 个（88%），详见 `ai/021-VB6语法成员覆盖率统计.md` |
| 回归测试 | `tests/` 下 98 个 `.bas` + 13 个 `.vbp`，由 `scripts\test.bat` 驱动 |
| 代码规模 | 编译器 C++ 约 4.0 万行，RTL（C）约 1.8 万行 |

> 权威进度以 `ai/004-进度表.md` 为准；每次会话恢复修复上下文请读 `ai/C3_FIX_HANDOFF.md`。

---

## 三、快速开始

### 3.1 环境要求

| 工具 | 版本 | 说明 |
|------|------|------|
| Visual Studio 2022 Community | MSVC 14.4x | 提供 `cl.exe` / `link.exe` / `rc.exe` |
| CMake | ≥ 3.20（实测 3.31.6） | 需先加载 vcvarsall 才在 PATH |
| Ninja | 1.12+ | 生成器 |
| Windows SDK | 10.0.22621+ | 窗体/COM 相关头文件 |

`vcvarsall.bat` 默认路径：
`C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat`

### 3.2 构建编译器

```bat
scripts\build              REM 增量构建
scripts\build clean        REM 清理 .build 后完整构建
```

产物在 **`.build\C3.exe`**（注意：不在 `.build\Release\` 下）；`build.bat` 会顺带把它复制到 `publish\C3.exe`。

构建后建议先冒烟验证产物可用：`scripts\test smoke`（用例 `tests\smoke.bas`）。

PowerShell 一键构建 + 测试（Agent 会话推荐）：

```powershell
scripts\dev.ps1                  # 构建 + 测试
scripts\dev.ps1 -SkipTest        # 只构建
scripts\dev.ps1 -SkipBuild       # 只测试
```

### 3.3 编译一个 VB6 程序

```bat
.build\C3.exe tests\hello.bas -o output\hello.exe
.build\C3.exe 工程.vbp --output-dir output
.build\C3.exe --dll --arch x64 -o output\VBMAN.dll 工程.vbp
```

### 3.4 运行测试

```bat
scripts\test               REM 全部
scripts\test run           REM 仅运行测试
scripts\test compile       REM 仅编译测试
scripts\test syntax        REM 仅语法检查
scripts\test smoke         REM 仅冒烟测试（构建后快速验证产物可用，见 tests\smoke.bas）
```

```bat
scripts\compile tests\hello.bas        REM 单文件编译（自动带 MSVC + RTL 环境）
scripts\run hello                      REM 运行 output\hello.exe
```

### 3.5 全量对照回归（C3 vs 真 VB6）

`tests\regress_all.ps1` 把 `tests\` 下除 `vbman` 外的全部用例（vbp / bas / frm / aux，当前 118 个）与**微软真 VB6**（本机 `VB6Mini`，命令行 `/make` 编译）逐个对照：双方各自编译、运行，按矩阵判定，结果增量写入 `tests\_vb6ref\report-*.csv`（含每个用例的双方编译结果、退出码与失败备注）。

```powershell
# 前置：先构建 C3（scripts\build），MSVC 环境同 run_tests.ps1（设 C3_VCVARSALL 或让脚本经 vswhere 探测）

.\tests\regress_all.ps1                 REM 全量（约 25 分钟）
.\tests\regress_all.ps1 -Filter test_iif*   REM 只跑匹配用例
.\tests\regress_all.ps1 -List           REM 枚举用例，不执行
.\tests\regress_all.ps1 -CompileOnly    REM 只比编译，不运行
```

可选覆盖：参数 `-C3 <路径>` / `-VB6 <路径>`，或环境变量 `C3_EXE` / `VB6_EXE` / `C3_TESTS_WORKROOT`。

判定矩阵：

| 判定 | 含义 |
| ---- | ---- |
| PASS | 双方编译 OK，运行退出码一致（双方都有输出时还比内容） |
| REG-C3 | 真 VB6 可编译而 C3 失败 —— **可疑回归，需排查** |
| C3-EXT | C3 可编译、真 VB6 不支持（C3 扩展语法/能力） |
| BOTH-FAIL | 双方都失败（注意：tests 里有故意的负例用例，属预期） |
| RUN-DIFF / RUN-TIMEOUT | 编译都过但运行行为不一致 / 一方超时 |
| SKIP-AUX | 辅助模块（非独立用例），跳过 |

备注：真 VB6 编译需 `VB6.EXE` 可无界面运行且其链接器对工作目录可写（本仓库约定工作目录 `tests\_vb6ref\`，该目录已 gitignore，运行产物不进库）。

### 3.6 vbman 专项对照回归（只测 vbman）

`tests\regress_vbman.ps1` 是 3.5 的补集：只对 `tests\vbman` 的权威工程（`src\VBMAN.vbp`，`Type=OleDll`，129 模块）做 C3 vs 真 VB6 对照。OleDll 无运行阶段，改用**导出表比对**替代运行比对（需 MSVC 环境的 `dumpbin`）。

```powershell
.\tests\regress_vbman.ps1                  REM 完整对照（C3 全量编译 + VB6 /make）
.\tests\regress_vbman.ps1 -SyntaxOnly      REM 只跑 C3 前端（秒级，不调 cl、不用 VB6）
.\tests\regress_vbman.ps1 -List            REM 枚举目标，不执行
.\tests\regress_vbman.ps1 -Target VBMAN_EXE.vbp   REM 换目标（src 下的 vbp）
```

判定矩阵：PASS / REG-C3 / C3-EXT / BOTH-FAIL（含义同 3.5）、`EXPORT-DIFF`（C3 缺 VB6 的导出符号）、`DLL-MISSING`（报成功但无产物）。报告落 `tests\_vb6ref_vbman\report-*.csv`。

与 3.5 的两点差异：

- VB6 侧在 `tests\vbman` 的**副本**里编译，因此不改写 `src\VBMAN.vbp` 的 Reference 路径、也不覆盖 `dist\` 下产物 —— 无需事后 `git checkout` 恢复子模块
- 导出表按「VB6 的导出被 C3 全覆盖」判定，C3 多导的符号（如 `DllMain`）只记入备注，不算差异

可选覆盖：参数 `-C3` / `-VB6` / `-Arch`，或环境变量 `C3_EXE` / `VB6_EXE` / `C3_VBMAN_WORKROOT`。

---

## 四、命令行参数

完整参数见 `ai/C3_COMPILE_OPTIONS.md`，常用项：

| 参数 | 说明 |
|------|------|
| `-o <文件>` | 输出文件路径 |
| `--output-dir <目录>` | 输出目录（默认源文件所在目录） |
| `--arch <x64\|x86>` | 目标架构（默认 x64；32 位 COM 组件用 x86） |
| `--target <平台>` | win-x86 / win-x64 / linux-x64 / macos-arm64 |
| `--gui <模式>` | native / webview / none |
| `--dll` | 编译为 ActiveX DLL（自动生成并嵌入 TypeLib） |
| `--syntax-only` | 只做语法检查 |
| `-d, --define <N=V>` | 定义条件编译常量，如 `-d:DEBUG=-1` |
| `-O <0-3>` / `-g` | 优化级别 / 调试信息 |
| `--typelib <文件>` / `--no-auto-typelib` | TypeLib 显式引用 / 禁用自动加载 |
| `--progid` / `--libid` | ActiveX DLL 的 ProgID 前缀 / TypeLib LibID |
| `--dump-tokens / --dump-ast / --dump-symbols / --dump-preprocess / --dump-ir / --dump-frm` | 各阶段中间产物转储 |
| `--emit-c` | 只生成 C 代码，不调用 cl/link |
| `--keep-for-debug` | 保留中间文件（默认清理） |
| `--compat-check` | 兼容性检查模式 |
| `-v` / `-h` / `-V` | 详细输出 / 帮助 / 版本 |

中间文件目录：`%TEMP%\C3C\<会话ID>\`。调试链接错误（LNK2019 等）时加 `--keep-for-debug` 检查生成的 C 与符号名是否一致。

编译 ActiveX DLL 等场景**必须在 MSVC 环境中运行** C3（先 `vcvarsall.bat x64`）。

---

## 五、目录结构

```
c3.vb6.pro/
├── src/                   编译器源码（C++17）与 VB6 运行时（C）
│   ├── common/            诊断、源码管理、编码转换（GBK↔UTF-8）
│   ├── lexer/             词法分析（LL(2) + Unicode）
│   ├── preprocessor/      条件编译 #Const / #If
│   ├── parser/            递归下降 + Pratt 表达式（分 expr/stmt/decl 三个文件）
│   ├── ast/               AST 节点、Visitor、打印器
│   ├── semantics/         符号表、类型系统、两遍语义分析
│   ├── project/           .vbp / .frm / .frx 解析
│   ├── com/               TypeLib 解析（前期绑定）
│   ├── typelib/           CreateTypeLib2 内建生成 .tlb
│   ├── backend/           C 代码生成（cgen_expr/stmt/decl/com/form/util）+ MSVC 驱动
│   ├── driver/            CLI 入口、编译流水线编排、RTL 资源解包
│   ├── rtl/core/          VB6 运行时：4 个家族子目录（vb6rtl / vb6com / vb6comserver / vb6forms）
│   └── ir/                历史遗留空壳（已停用，不参与构建产物）
├── tests/                 回归测试：98 个 .bas + 13 个 .vbp + 多模块/COM/窗体用例
├── scripts/               构建、测试、编译、运行脚本（见下）
├── ai/                    全部设计文档、进度表、开发历程（99 篇）
├── publish/               发布包：C3.exe + 内置迷你 MSVC 工具链 + 右键菜单 + demos
├── www/                   官网 c3.vb6.pro 静态站点源码 + scp 部署脚本
├── archive/vbman/         VBMAN 大型 VB6 库（真实工程回归靶子）+ 编译问题日志
├── output/                测试编译产物（exe/out/err），git 忽略
├── .build/                CMake + Ninja 构建目录，可随时删除重建
├── _res/ logo/ pic/       品牌素材、favicon、演示动图
└── reference/             预留给外部参考项目 clone（当前为空）
```

### 脚本工具（`scripts/`）

| 脚本 | 用法 | 说明 |
|------|------|------|
| `build.bat` | `build [clean]` | 构建 C3.exe |
| `test.bat` | `test [all\|run\|compile\|syntax\|smoke] [verbose]` | 回归测试（smoke = 冒烟） |
| `compile.bat` | `compile <source> [outdir]` | 编译单个 .bas/.frm/.vbp |
| `run.bat` | `run <exename> [timeout]` | 运行 output/ 下的 EXE |
| `dev.ps1` | `dev [-SkipBuild] [-SkipTest]` | 一键构建 + 测试 |
| `env.ps1` | `. .\scripts\env.ps1` | dot-source 加载 MSVC 环境 |
| `ga_dispatch_pd.ps1` | `ga_dispatch_pd.ps1 [-AlsoDevRegression]` | 派发 PD 整编 GA（按 headSha 认领 run），可选连带 dev 回归 |
| `pd_readout.ps1` | `pd_readout.ps1 -Run <run> [-Baseline <run>]` | 取 PD 整编工件，一次读齐 meta / cl 四段 / MEM 行 / 内存曲线 / 错误面 |
| `diff_c3error.ps1` | `diff_c3error.ps1 <oldRun> <newRun>` | 错误面**逐码**比对（判据唯一口径所在）；退出码 0 = 逐码相等 |

> **PD 整编性能线**的工作流、读数口径与当前成绩见 [`ai/C3_PD_BUILD_PERF_HANDOFF.md`](../ai/C3_PD_BUILD_PERF_HANDOFF.md)。

> **跨机器使用前必读**：`build.bat` / `env.ps1` 的路径配置（环境变量 `C3_VCVARSALL` / `C3_PROJECT_DIR`）说明见 [`scripts/README.md`](scripts/README.md)。

---

## 六、架构要点

**决策 006：后端走 C 代码生成 + MSVC，不走 LLVM。**
`src/backend/llvm_backend.*`、`pe_writer.*`、`src/ir/*` 均为 2 行空壳，CMake 中 LLVM 后端默认 `OFF`。
CLI 仍保留 `--dump-ir` / `--emit-llvm` 开关，但没有消费方。

**决策 008：库优先（library-first）。**

| 产物 | 内容 |
|------|------|
| `vb6c3-core.lib` | 词法 / 语法 / 语义 / 符号 / 类型 / 诊断 |
| `vb6c3-cgen.lib` | C 代码生成后端 |
| `C3.exe` | 薄 CLI 驱动（链接上述两个库） |

**RTL 内嵌（P10，源码级）。** 14 个 RTL 头文件 + 25 个 RTL `.c` **源码**（共 39 项）以 `RCDATA` 资源嵌入 `C3.exe`
（见 `src/driver/c3rtl.rc`），运行时由 `rtl_embedded.cpp` 解包到 `%TEMP%\C3C\<会话ID>\rtl\`，
与生成的 C 代码一起交给 `cl.exe` 编译（`/Gy` + 链接器 `/OPT:REF` 自动剔除未引用的 RTL 代码）。
因此**用户机器上不需要安装任何 VB6 运行时**，产物静态链接、零依赖。
改动 `src/rtl/` 后直接重编 C3.exe 即生效（`scripts\build.bat`），无需任何预编译步骤。

**RTL 组成**（`src/rtl/core/`，按家族分目录）：

| 目录 | 文件数 | 职责 |
|------|--------|------|
| `vb6rtl/` | 9 | `vb6rtl.h`（伞头）+ 7 个子头 + `vb6rtl.c`：BSTR 字符串、VARIANT、SAFEARRAY、数学、转换、日期、文件 I/O、Err 对象、Format |
| `vb6forms/` | 13 | `vb6forms.h` / `.c` + 10 个实现 + 内部头：Win32 窗体与控件（缇↔像素 1/15）、消息循环、菜单 |
| `vb6com/` | 8 | `vb6com.h` / `.c` + 5 个实现 + 内部头：COM 客户端 CreateObject / GetObject / IDispatch 后期绑定 / VARIANT 封送 |
| `vb6comserver/` | 7 | `vb6comserver.h` / `.c` + 4 个实现 + 内部头：COM 服务端类工厂、DllGetClassObject / DllRegisterServer |
| `di/` | 8 | `vb6_di_stubs.c`（手写：序号 / msvbvm60 运行时 / 动态加载）+ 7 个按 Lib 家族生成的转发桩（win32 / user32 / gdiplus / crypto / net / com / shell）：Win32 API `Declare` 转发（解决 x64 下 msvbvm60 等无导入库的问题）。生成物由 `scripts/gen_di_stubs.ps1` 产出，家族来自 C3 生成头里的 `/* vb6_di_lib: <lib> */` 标记 |

子目录只影响**源码树可读性**：47 个文件解包到同一个平铺临时目录，RTL 内部 `#include` 一律按 basename，
且各家族之间**无交叉 include**，因此编译路径与拆分前完全一致。
新增 / 重命名 / 删除 RTL 文件需同步 **5 处**（`c3rtl.rc` / `rtl_embedded.hpp` / `rtl_embedded.cpp` /
`driver.cpp` / `CMakeLists.txt`）。

---

## 七、已支持与未支持

### 已支持

- **语言**：模块 / 类模块 / UDT / Enum / 常量、Pratt 14 级表达式、软关键字、条件编译（含 `Win64` / `VBA7` 等内置常量）、
  GoSub/Return、With、For Each、On Error / Resume / Err 对象、默认属性解析、ParamArray / Optional / 命名参数
- **COM**：CreateObject / GetObject、前期绑定（TypeLib 导入）与后期绑定（IDispatch）、Implements、WithEvents 事件
  （内部类 / 窗体控件 / 外部 COM 三条路径）、IEnumVARIANT 集合枚举、ActiveX DLL 产出（内建生成 .tlb）
- **窗体**：`.frm` + `.frx` 二进制资源（图片 / 图标 / ImageList / 文本）、控件数组、MDI、菜单、Timer、缇↔像素换算
- **控件**：VB6 工具箱 21 类内置控件中 **14 类全链路可用**（窗口创建 + 属性读写 + 事件分发）≈ 67%
  —— Form / MDIForm、CommandButton、TextBox、Label、CheckBox、OptionButton、ListBox、ComboBox、
  Frame、PictureBox、HScrollBar / VScrollBar、Image、Timer、Menu、WebBrowser（WebView2）
- **窗体实现形态**：RTL `vb6forms.c` 用 Win32 原生窗口类重实现（BUTTON / EDIT / STATIC / LISTBOX … + VB6_SHAPE / VB6_LINE 自绘），
  不是复刻 VB6 运行时，因此产物零依赖、静态链接，且支持 VB6 自身不具备的 x64
- **控件能力三档**：属性层最厚（50+ 双向读写含 Font 六件套 / Back&ForeColor / Alignment / TabIndex / ToolTipText / MousePointer…），
  事件层次之（约 25 个含 KeyPress / Validate(Cancel) / QueryUnload(Cancel) / MouseEnter / Mouse Leave），**方法层近乎空白**（仅 AddItem / RemoveItem / Clear + Timer）
- **运行时**：126 个 VB6 内置函数（字符串 / 数学 / 日期 / 转换 / 文件 I/O / 数组 / 财务 等），静态链接

### 尚未支持

- **半接线（RTL 与属性表已实现，缺 `controlTypeToWin32Class` 窗口类映射 → 编过能跑但运行时不可见）**：
  Shape / Line、DriveListBox / DirListBox / FileListBox —— 差一步接入，推进计划见第十章路线图与仓库 Issues
- **未实现**：Data / OLE（x64 无 DAO / MDAC 支撑，列为豁免但会明确报错）、SSTab、Toolbar / StatusBar、CommonDialog
- **控件方法**：Move / SetFocus / ZOrder / Refresh / Drag 未实现；绘图语句 PSet / Line / Circle / Print 为 VB6 **关键字语法**（非函数调用，需专用语句产生式），Cls / PaintPicture 同样未实现
- **窗体相关**：Form_Click / Paint / DragDrop 事件、PictureBox 作容器、任意深度容器嵌套（当前仅 Frame 单层子控件）、per-monitor DPI（缇换算按 96 DPI 硬编码）
- 多接口 `Implements IFoo, IBar`（已决策跳过）
- 部分内置函数：注册表 4 函数（GetSetting / SaveSetting / GetAllSettings / DeleteSetting）、
  FormatDateTime、CVErr、GetAttr / SetAttr、Erl / Tab / Spc 等，清单见 `ai/021`
- LSP / DAP / 格式化器 / Linter / 窗体设计器 —— 见 `ai/013-工具链与生态建设规划.md`，均**尚未实现代码**

> 遇到不支持的语法，C3 会明确报错，不会静默生成错误代码。

---

## 八、文档导航（`ai/`）

| 文档 | 用途 |
|------|------|
| `003-开发计划.md` | 分阶段计划与里程碑定义 |
| `004-进度表.md` | **权威进度**：P0~P33 全部阶段与任务状态 |
| `009-产品定位与最终成果展望.md` | 定位、可行性分级、VB6 语言边界情况清单 |
| `011-参考项目综合分析与技术借鉴.md` | 对 FreeBASIC / vb6parse 等的参考分析 |
| `012-第二期VBA-VBS-ASP规划.md` | 第二期方言扩展（VBA / VBS / ASP） |
| `013-工具链与生态建设规划.md` | LSP / DAP / 格式化器 / 包管理器 等长期规划 |
| `020-P20遗留项任务清单.md` | 遗留缺陷清单 |
| `021-VB6语法成员覆盖率统计.md` | 内置函数覆盖率与缺口优先级 |
| `C3_COMPILE_OPTIONS.md` | **完整命令行参数参考** |
| `C3_FIX_HANDOFF.md` | **修复任务交接**：新会话直接读它继续修 bug |
| `开发历程/` | 99 篇按阶段记录的修复与决策过程 |
| `ENV.md`（根目录） | 开发环境速查：路径、脚本、编码约定 |
| `CONTRIBUTING.md`（根目录） | **参与贡献指南**：协作流程 / PR 要求 / 测试规范 / 编码约定 |

VBMAN 编译过程的问题日志见 `archive/vbman/c3log/001.md ~ 054.md`。

---

## 九、开发约定

| 规则 | 说明 |
|------|------|
| 换行符 | 全项目统一 **CRLF** |
| VB6 工程文件 | `.frm` / `.bas` / `.cls` / `.vbp` 必须 **GBK**，否则 VB6 IDE 打开乱码 |
| C/C++ 源码 | **UTF-8**（与 CMake / MSVC 默认一致），CMake 已加 `/utf-8` |
| 部分 `.md` 带 UTF-8 BOM | 编辑时保持原样 |
| 构建目录 | `.build\` 可随时删除重建；遇 release-only 崩溃先 `--clean-first` 全量重建再复现 |
| 崩溃追踪 | 设环境变量 `C3_CRASH_TRACE=1` 启用 dbghelp 栈追踪 |

### 版本号约定（唯一权威源）

**权威源只有一处：根目录 `VERSION` 文件**，内容形如 ` 0.10.4`（允许前后带空格，读取方会自动剔除）。

自 2026-09-17 起，原有各处副本已改为**从该文件自动派生**，不再需要手工同步：

| 使用方 | 取值方式 | 用途 |
|--------|----------|------|
| `scripts\release.bat` | 直接读 `VERSION` 文件并剔除前后空格 | 发布目录名与发布包名 |
| `CMakeLists.txt` 的 `project(vb6c3 VERSION ...)` | `file(READ "${CMAKE_CURRENT_LIST_DIR}/VERSION")` + `string(STRIP)` | CMake 工程版本元数据 |
| `src/driver/driver.cpp` 的 `Driver::printVersion()` | CMake 经 `target_compile_definitions` 注入编译期宏 `VB6C3_VERSION_STRING` | `C3 --version` / `-V` 的输出 |
| 本文件第二章「当前状态」表 | 指向本文件，不写死具体值 | 文档展示 |

**升版本只需改 `VERSION` 一处**，CMake 工程版本与 `--version` 输出会自动跟随。

自检（在子模块根目录执行）：

```bash
cat VERSION                                       # 权威值，升版本只改这里
grep -n "file(READ" CMakeLists.txt                # CMake 从文件读取
grep -n "VB6C3_VERSION_STRING" CMakeLists.txt     # 宏注入给 c3 目标
grep -n "CMAKE_CONFIGURE_DEPENDS" CMakeLists.txt  # 改 VERSION 会自动触发重配置
grep -c '0\.10\.' src/driver/driver.cpp           # 应为 0：driver 内无硬编码版本
./publish/C3.exe --version                        # 应与 VERSION 一致
```

**历史事故（同一问题发生过三次，现已从机制上消除）**：

| 时间 | 提交 | 情况 |
|------|------|------|
| 2026-09-17 | `afbea7d` | 首次统一为 0.10.2 —— 此前三处各不相同（`VERSION` 0.10.1 / `driver.cpp` 0.10.0 / `CMakeLists.txt` 0.1.0） |
| 2026-09-17 | `0569942` | 升到 0.10.3 时**只改了 `VERSION`**，另两处仍停在 0.10.2 |
| 2026-09-17 | `368dc69` | 升到 0.10.4 时**仍只改了 `VERSION`**，`README` 甚至停留在 0.10.0 |
| 2026-09-17 | 本次提交 | 根治：`CMakeLists.txt` 改为从 `VERSION` 读取、`driver.cpp` 改用编译期宏，三处副本降为一处 |

---

## 十、路线图（摘要）

1. **VBMAN.dll 运行期行为对齐** —— 当前编译 / 链接 / 注册全通，下一步是运行期功能差异
2. **内置函数补齐** —— DoEvents、FormatDateTime、CVErr、注册表 4 函数、GetAttr / SetAttr
3. **控件补齐（67% → 100%）—— 当前首要方向**：定义「100%」= 21 类内置控件能创建 + 属性可读写 + 事件能分发 + 常用方法可调用
   - P0：接通 Shape / Line / Drive-Dir-FileListBox 五个半残控件的窗口类映射（≈ 67% → 90%）
   - P1：控件方法 Move / SetFocus / ZOrder / Refresh / Drag，以及绘图语句（PSet / Line / Circle / Print 关键字语法 → 翻译为同名 RTL 函数）+ DrawWidth / ScaleLeft 等画布属性
   - P2：Data / OLE 明确报错边界、容器任意深度嵌套、PictureBox 作容器
   - P3：Form_Click / Paint / 拖放事件、per-monitor DPI
   - 任务拆分与认领状态见仓库 Issues
4. **c3-lsp 语言服务器 + VS Code 插件** —— P0 级生态组件（`ai/013`）
5. **c3-dap 调试适配器** —— 依赖代码生成阶段输出 VB6 行号 ↔ C 行号映射
6. **第二期方言** —— VBA / VBS / ASP（`ai/012`）

详细规划与工期估算见 `ai/013-工具链与生态建设规划.md`。

---

## 十一、参考来源

C3 在语法规则核对、测试语料与实现思路上参考了以下项目（完整索引见 `ai/003-开发计划.md` 第七节）。源码副本放在 `reference/`（该目录已被 `.gitignore` 忽略，不入库）：

| 项目 | 版本 | 许可 | 应用位置 | 参考内容 |
|------|------|------|----------|----------|
| [vb6parse](https://github.com/scriptandcompile/vb6parse) | 1.0.1 | MIT | P1 Lexer/Parser · P4 RTL 清单 | Rust 实现的 VB6 解析器，覆盖工程 / 窗体 / 模块 / 控件，附 30+ 个真实 VB6 工程语料；120+ Token、迭代式 Pratt、160+ 内置函数清单 |
| VisualBasic6.g4（随 [proleap-vb6](https://github.com/uwol/proleap-vb6-parser) 发布） | — | AGPL-3.0 | P1 语法规则 | 权威语法参考：50+ 语句、~130 个歧义关键字、14 级表达式优先级 |
| [proleap-vb6](https://github.com/uwol/proleap-vb6-parser) | — | AGPL-3.0 | P1 AST 设计 | 基于 ANTLR4 的 VB6 分析器与转换器，AST / ASG 元模型类清单 |
| vb6grammarfuzz | 0.1.0 | — | P2+ 质量保障 | 基于 `VisualBasic6.g4` 的 VB6 语法模糊测试器（仓库内本地工具，依赖 vb6parse，无公开地址） |
| [RustASP](https://github.com/ferocknew/rustasp) | — | — | P1.3 Pratt 解析器 · P2 作用域链 | Rust 实现的经典 ASP 服务器原型，Pratt 核心实现与作用域链设计参考 |
| aspgo | — | — | P4 运行时设计 | Go 实现的经典 ASP 服务器（含 VBScript），Variant / COM / Error 运行时设计思路参考（本地参考项目，无公开地址） |
| [FreeBASIC (fbc)](https://github.com/freebasic/fbc) | 1.20.0 | 编译器 GPL-2.0+ / RTL LGPL-2.1+ | 全局架构参考 | 成熟 BASIC 编译器源码，RTL 组织（IR_VTBL）、x64 调用约定与代码生成策略 |
| [AvaloniaVisualBasic6](https://github.com/BAndysc/AvaloniaVisualBasic6) | — | MIT | P1 语法规则 | C# + Avalonia 复刻的 VB6 IDE 与语言，含可视化设计器与 VB6 兼容工程格式（VB6.g4 来源） |
| [twinBASIC](https://twinbasic.com) | — | 闭源 | P5 兼容性验证 | VB6 兼容一体化 IDE，作为 C3 语义兼容性的对照基准 |

> aspgo / vb6grammarfuzz 为本地参考项目，无公开仓库地址；各参考项目的本地路径索引见 `ai/003-开发计划.md` 第七节。

---

## 十二、参与贡献

欢迎参与 C3 开发！协作流程（分支模型 / PR 要求 / CI 门槛）、测试规范、代码风格与文件编码约定，请先阅读 **[CONTRIBUTING.md](CONTRIBUTING.md)**。

- **核心信条**：不支持的语法必然报错，绝不静默误编——错编译比不编译更危险
- **动手前**：先在仓库 Issues 中认领任务，避免与他人撞车；个人想法草稿见根目录 [todo.md](todo.md) 索引
- **快速上手**：`scripts\build` 构建 → `scripts\test` 全量回归（PR 提交前必须全绿）
- **外部贡献者**：fork 本仓库 → 在自己 fork 里开分支 → 向 `main` 发 PR
