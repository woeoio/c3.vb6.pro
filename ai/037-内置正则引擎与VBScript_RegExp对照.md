# 037 内置正则引擎（替代 VBScript.RegExp）

> 状态：**首版可用，未达 100% 对齐**。对照 51 条用例中 42 条与真组件逐字一致；余下 9 条是本文件第 6 节列出的已知差距。
> 关联：`ai/036-VBScript_RegExp_完全参考手册.md`（真组件语义手册，本引擎的判据出处）。

## 1. 为什么要自己写

`CreateObject("VBScript.RegExp")` 在原实现里会走注册表激活 `vbscript.dll`，把整个 VBScript 引擎拖进进程：

- 依赖 `vbscript.dll` 存在且可注册（精简系统、Wine、容器、Server Core 上都没有）；
- 首次激活 6~15 ms，且每个 `RegExp` 实例都是一个外部 COM 对象；
- 32/64 位互操作与 DPI 无关，但**部署形态**被绑死在一个外部 DLL 上。

所以：**凡是程序里写 `CreateObject("VBScript.RegExp")`，一律由 RTL 内置的等价对象抢先接管**，进程内不再加载 `vbscript.dll`。

## 2. 怎么接进去

拦截点在 `src/rtl/core/vb6com/vb6com.c` 的 `vb6_CreateObject` 最前面：

```c
if (vb6_RegExp_IsProgramId(progId)) { ... return vb6_RegExp_New(); }
```

- 认的 ProgID：`VBScript.RegExp` / `VBScript.RegExp.1`（大小写不敏感）。
- **逃生开关**：环境变量 `C3_REGEXP_PREFER_MS=1` 时不拦截，仍旧走注册表激活真组件。用于「怀疑引擎有偏差，想拿真组件复现一遍」的现场排查。
- 引擎实现：`src/rtl/core/vb6com/vb6com_regexp.c`（单文件，无外部依赖，纯 C，UTF-16 内部表示）。

RTL 是 **RCDATA 嵌进 `C3.exe`** 的，新增一个 `.c` 必须五处同时登记，漏一处就是「改了没生效」：

| 位置 | 内容 |
| --- | --- |
| `CMakeLists.txt` | 源文件加入 RTL 静态库 |
| `src/driver/c3rtl.rc` | `227 RCDATA "../rtl/core/vb6com/vb6com_regexp.c"` |
| `src/driver/rtl_embedded.hpp` | `#define RTL_VB6COM_REGEXP_C 227` |
| `src/driver/rtl_embedded.cpp` | 解包表加一行 `{ RTL_VB6COM_REGEXP_C, "vb6com_regexp.c" }` |
| `src/driver/driver_link.cpp` | `sourceFiles.push_back(rtlDir + "/vb6com_regexp.c")` |

**验证嵌入是否为最新**：`powershell -File scripts/check_rtl_embedded.ps1`（逐字节比对磁盘与 exe 里的 136 份 RTL 资源）。改完 `.c` 必须重跑它，不能只看编译成功。

## 3. 对象模型（与真组件同形）

| 对象 | 成员 |
| --- | --- |
| `RegExp` | `Pattern` / `IgnoreCase` / `Global` / `Multiline`、`Test(s)`、`Execute(s)`、`Replace(s, rep)` |
| `MatchCollection` | `Count`、默认成员 `Item(i)`（0 基） |
| `Match` | `Value`、`FirstIndex`、`Length`、`SubMatches` |
| `SubMatches` | `Count`、默认成员 `Item(i)`（0 基，i=0 是第 1 个捕获组） |

全部是手写 `IDispatch`（`GetIDsOfNames` 自解析名字，不走类型库），因此晚绑定 `Object` 变量、`CallByName`、`For Each`（`_NewEnum`）都可直接用。

## 4. 支持的语法（VB6/VBScript 子集）

| 类别 | 项 |
| --- | --- |
| 字面量 | 普通字符、`\n` `\r` `\t` `\f` `\v`、`\xHH`、`\uHHHH`、`\cX` |
| 字符类 | `[...]`、区间、取反 `[^...]`、类内 `\d \w \s \D \W \S`、类内转义 |
| 简写类 | `\d \D \w \W \s \S` |
| 点号 | `.`（不吃 `\n`，吃 `\r` —— 与真组件一致，见 036） |
| 锚点 | `^` `$`（`Multiline` 下按 `\r\n`/`\n`/`\r` 断行）、`\b` `\B` |
| 分组 | `(...)`、`(?:...)`、`(?=...)`、`(?!...)` |
| 量词 | `*` `+` `?` `{n}` `{n,}` `{n,m}`，均支持惰性 `?` 后缀 |
| 后向引用 | `\1` … `\9` |
| 选项 | `IgnoreCase`、`Global`、`Multiline` 三个属性 |

匹配器是**回溯 + 续延（continuation）**结构，不是 NFA 模拟：`SEQ` 里每个子节点带一个「后续怎么接着走」的续延，`ALT` 逐支试、失败回退。所以惰性量词、正向预查、后向引用的语义直接来自控制流，不需要额外机制。

`Replace` 的替换串支持 `$1`..`$9`、`$&`、`$$`、`` $` ``、`$'`。`Global=False` 时只替换第一处。空匹配（如 `a*` 命中空串）按真组件行为**前进 1 个字符**再找，不会死循环。

## 5. 对照方法（判据从哪来）

**不信文档，只信跑出来的数**。基准是 `cscript` 跑的**真** `VBScript.RegExp`：

- `tests/regexp_native/ReCases.vbs` —— 纯 VBScript，51 条用例，每行输出 `pattern | 动作 | 结果`；
- `tests/regexp_native/ReMain.bas` + `ReApp.vbp` —— C3 侧同源同用例，写 `re_c3.txt`；
- `tests/regexp_native/run_compare.ps1` —— 跑两边、逐行 diff、只打印不一致行。

用例覆盖：字面量/`.`/类/量词（含惰性、`{n,m}`）/分组与嵌套/`(?:)`/`(?=)` `(?!)`/`^$`/`\b`/后向引用/`\xHH` `\uHHHH`/IgnoreCase/Multiline/Global/空匹配推进/`Replace` 各替换符/`SubMatches`/三类非法模式的错误号。

当前读数：**51 行中 42 行一致**。

## 6. 已知差距（未对齐的 9 条）

### 6.1 组捕获值（3 条）

| 模式 | 目标串 | 真组件 | 本引擎 |
| --- | --- | --- | --- |
| `(a)(b)?` | `a` | `s0=[a] s1=[]` | `s0=[a] s1=[a]` |
| `(a)\|(b)` | `b` | `s0=[] s1=[b]` | `s0=[] s1=[]` |
| `((a)(b))?` | `ab` | `s0=[ab] s1=[a]` | `s0=[a] s1=[a]` |

症状统一：**`SubMatches(i)` 取到的永远是第 1 组的值**，而 `Count` 正确。两个候选原因都已在代码里改过一轮（`(i)` 参数经 `reArgIndex()` 拆 `VT_VARIANT`/`VT_BYREF` 两层；组号改成「左括号一出现就先分配」，让外层组号小于内层），**改动后读数一字未变** —— 说明还没打中真正的病灶，下一刀要从这里继续：

1. 打一条诊断用例，把 `m.SubMatches(0)`、`m.SubMatches(1)`、`m.SubMatches.Item(1)` 三种写法分别输出，确认 VB 侧到底是走「属性带参」还是「取对象再调默认成员」；
2. 若确认走的是 `m_Invoke` 的 `DISPATCH_PROPERTYGET` 带参分支，则在分支里把 `reArgIndex` 的返回值也一并打印出来，确认参数解码；
3. 顺带核 `K_GROUP` 节点在「可选/择一」分支里失败回退时 `caps[]` 是否被正确还原 —— `((a)(b))?` 的 `s0` 至今仍是 `a` 而不是 `ab`，很像外层组的 capture-end 没落。

### 6.2 模式错误号（4 条）

| 模式 | 真组件 Err.Number | 本引擎 |
| --- | --- | --- |
| `[` | 5019 | 9 |
| `a{3,2}` | 5017 | 9 |
| `(?<x>a)` | 5017 | 9 |
| `(abc` | 5020 | 9 |

`9` 是 `DISP_E_EXCEPTION`（`0x80020009`）的低 16 位 —— 也就是**错误号在 HRESULT 里被当成 `DISP_E_*` 系解掉了**，没走到 `EXCEPINFO.scode`。代码里已从 `vb6_ComFailExcep(...)` 改成直接返回 `vb6_ComVbHresult(err)`（`0x800A0000|err`），读数仍是 9，说明这条腿的转换发生在更外层（`vb6_ComCheckError` 的取值顺序或调用方另有一套映射），下一刀要顺着 `vb6_ComCheckError → Err.Number` 全链走一遍。

错误号分档已在解析器里落好：`RE_ERR_SYNTAX=5017`（通用语法错）、`RE_ERR_MISSING_RB=5019`（缺 `]`）、`RE_ERR_MISSING_RP=5020`（缺 `)`），只差把号送出去。

## 7. 扩展路线（借 PCRE2 / Python `re` 补 VB 没有的语法）

现状是「严格 VB 兼容子集」：真组件不支持的写法，本引擎也报 `5017`，这样对照表才有意义。要往超集走，**必须是由属性/开关显式打开的第二档语义**，不能悄悄改默认行为（否则同一段 VB6 代码行为变了，违反兼容性红线）。建议：

```vb
' 提案形态（尚未实现）
re.Extended = True     ' 打开扩展语法档
```

扩展分三批，按「实现成本 / 实际用量」排：

**第一批（解析器小改，1~2 刀）**
- 命名组 `(?<name>...)` / `(?'name'...)` + `\k<name>`（真组件报 5017，见 6.2）；
- 局部内联选项 `(?i:...)` `(?-i:...)`；
- `\A` `\z` `\Z`、`\Q...\E`；
- `{n,m}` 惰性已支持，补「固有名次组」`\g{n}`。

**第二批（要加匹配器指令，3~5 刀）**
- 后顾 `(?<=...)` `(?<!...)`：当前续延结构在**已知定长**时可实现（先回退 k 个位置再正查），变长后顾需要反向扫描，成本高；
- 原子组 `(?>...)` 与占有量词 `*+` `++`：靠「失败后不回退」的标记位，可显著压住灾难性回溯；
- 条件组 `(?(1)yes|no)`。

**第三批（接外部库，另开一刀）**
- `\p{L}` 之类的 Unicode 属性、`\X` 字素簇：自己维护表不现实，直接链 PCRE2；
- 接法：`#ifdef C3_WITH_PCRE2` 编译期开关，`Pattern` 里出现本引擎不认的构造且 `Extended=True` 时，把整条模式交给 PCRE2，用同一套 `Match/MatchCollection` 外壳包回去 —— **对象模型和 VB 侧代码一行不改**，只换匹配内核。
- Python `re` 只作为**语义参考**（尤其 `(?P<name>)`、`\g<name>`、`re.split` 的零宽分割行为），不引入运行期 Python 依赖。

无论哪批，**每加一项都要先在 `ReCases.vbs` 里补一条对照用例**，并且注明「真组件此处报 5017，扩展档返回 X」——扩展档的判据由 PCRE2/Python 提供，不再由 VBScript 提供。

## 8. 遗留与注意

- **回溯无预算**：`(a+)+$` 这类灾难性模式目前会把线程跑满，与真组件一致（真组件也不设限）。第 7 节第二批的原子组/占有量词是缓解手段；若要在默认档加步数上限，属于**语义变更**，必须在提交信息里点名，且默认关闭。
- `Global=True` 时空匹配前进 1 字符的边界（末尾空匹配、`\b` 处空匹配）目前只用 4 条用例覆盖，语料还需要加长串/多空匹配的用例。
- 未接 `run_tests.ps1`：对照是文本 diff，不进 GUI 像素判据体系；跑法是直接执行 `tests/regexp_native/run_compare.ps1`。
- 性能未测：未做长串（>10 万字符）与高频 `Execute` 的基准。首版目标是「对得上」，不是「跑得快」。
