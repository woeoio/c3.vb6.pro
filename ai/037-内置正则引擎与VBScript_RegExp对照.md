# 037 内置正则引擎（替代 VBScript.RegExp）

> 状态：**首版完成，51/51 用例与真组件逐字一致**（oracle = cscript 跑的 VBScript.RegExp，夹具见 `tests/regexp_native`）。
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

当前读数：**51 行全部一致**（`PASS regexp_native 用例=51 行 两侧逐字一致`）。

## 6. 两个把首版钉在 42/51 的坑（已修，留档）

### 6.1 `SubMatches(i)` 恒取第 1 组 —— 参数是 `VT_VARIANT|VT_BYREF`

症状：`(a)(b)?` 在 `a` 上，`SubMatches(1)` 返回 `a` 而不是空；`(a)|(b)` 在 `b` 上两个下标都空；`((a)(b))?` 的第 0 组返回 `a` 而不是 `ab`。

根因：C3 生成的晚绑定调用把下标参数包成 `VT_VARIANT | VT_BYREF`，最初的取值代码只认裸 `VT_I2/VT_I4`，于是所有下标都落回默认值 0 —— 两个下标都读到第 1 组，表现出来就像「组号错了」。`MatchCollection.Item(i)` 没暴露这个问题，因为它的默认值 0 恰好是第一个元素。

修法：`reArgIndex()` 逐层拆 `VT_BYREF`（`V_VARIANTREF` / `V_I2REF` / `V_I4REF` / `V_R8REF` / `V_BSTRREF`），带环路防护，最多拆 4 层；同时接住 `SubMatches` 的两种调用形态（「属性带参」与「取对象再调默认成员」），免得只支持一种时又静默读错。

### 6.2 `V_VARIANT` 不是 oleauto.h 的宏 —— 本地能编、CI 全红

第一版 `reArgIndex()` 用了 `V_VARIANT(v)`，本机 VS2019 的 cl 放过（当成未声明函数，只在 `/W3` 下给 C4013 警告），CI 上 VS2022 的 cl 19.51 直接 `error C2102: '&' requires l-value`。RTL 是一个整体编译单元集合，**这一个文件编不过，整个 RTL 编不过，22 格全线红**（构建和 smoke 都过，因为 C3.exe 本身没受影响）。

修法：只用 oleauto.h 里真有的宏；需要字符串形态的下标时走 `V_BSTRREF` + `_wtoi`（补 `#include <stdlib.h>`）。

教训：**RTL 改动不能只看本机构建**。本机 cl 与 CI cl 版本不同（19.29 vs 19.51），`/W3` 的「未声明函数当 extern int」老行为会掩盖真错误；每刀推之前要么把 CI 的 cl 版本对齐，要么至少用 CI 产物编一个最小工程验证。

## 7. 尚未覆盖（不影响 50/50，但要做）

1. **未接 `run_tests.ps1`**：两份对照都是文本 diff，不进 GUI 像素判据体系；跑法就是
   `tests/regexp_native/run_compare.ps1`（VB 兼容面）与 `run_ext_compare.ps1`（扩展面）。
   要进 CI 得先给 harness 加一种 "text-compare" 模式。
2. **性能未测**：>10 万字符长串、高频 `Execute`、`Replace` 大文本都没有基准数。
3. **回溯有硬上限（先前写错，此处为准）**：`RE_STEP_LIMIT = 2000000`，超过按「此次尝试不匹配」收手，
   搜索继续下一个起点。真组件不设限，所以这是**静默差异**：灾难模式在长串上本引擎会认输。明细见 ai/038 第 8 节。
4. **语料仍偏薄**：`Global=True` 的末尾空匹配、`\b` 处空匹配、多行 `\r\n` 混排、`$0` 写法，覆盖条数还不够。
5. **`Err.Description` 比对未做**：目前只钉 `Err.Number`（真组件给的是中文描述，跨 locale 比对字节没意义）。

## 8. 扩展语法：已默认开启（详见 ai/038）

原计划的「`Extended` 属性开关」**不做了**（用户决定）：引擎默认就是超集，不设开关。
已落地的扩展项（命名组、`\k<name>`、`\A \z \Z`、`\Q...\E`、定长后顾、原子组、占有量词、
内联选项 `(?i:…)`/`(?s:…)`/`(?-i:…)`/`(?i)`、条件组 `(?(n)yes|no)`）全部默认生效，
判据由 `run_ext_compare.ps1` + `ext_oracle.py`（Python re，条件组那 7 条手写期望）提供：
**扩展面 45 行逐字一致**；VB 兼容面 **50 行逐字一致**。两条对照都已接进 `run_tests.ps1`
的静态段（`[TEXT]` 行，真跑两侧再 diff，不是「只编就算过」）。性能与灾难回溯的实测读数见
ai/038 第 9 节；与真组件的分歧清单、剩下未做的项（`\p`/`\X`，已定暂缓）见 ai/038。

下面这段早先的三批路线图保留，但执行方式改为「直接默认支持」，不再有开关一说：

现状是「VB 兼容子集 + 默认开启的扩展超集」。真组件不支持的写法，本引擎按 PCRE2/Python 语义工作，
分岧已逐条记在 ai/038 第 4 节。扩展项每加一条，都要在 `ext_oracle.py` 里补一条 Python 对照用例。

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

## 9. 遗留与注意

- 模式错误号分档：`RE_ERR_SYNTAX=5017`（通用语法错）、`RE_ERR_MISSING_RB=5019`（缺 `]`）、`RE_ERR_MISSING_RP=5020`（缺 `)`）。错误在**第一次 `Test/Execute/Replace`** 时抛（赋值不抛），与真组件一致；HRESULT 走 `vb6_ComVbHresult(err)`（`0x800A0000|err`），不走 `DISP_E_EXCEPTION` —— 否则 VB 侧只读到 9。
- `IgnoreCase` 用 `towupper/towlower` 逐字符比较，没做完整大小写折叠（`ß` → `SS` 之类）；真组件也是这样。
- 引擎只在 `CreateObject` 这一条腿上拦截；`New VBScript.RegExp`（VB6 里做不到，没有类型库）与 `GetObject` 不受影响。
- 判据是**文本 diff**，不是像素；每改一次引擎都要重跑 `tests/regexp_native/run_compare.ps1`，并看一眼 `PASS ... 51 行` 这行字。
