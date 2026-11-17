# 内置 Scripting.FileSystemObject（设计）

> 目标读者：接这一刀的我和下一个接手的人。写法沿用 `ai/039`（内置 Dictionary）那一套：
> **先写读数口径，再写怎么落地，最后写验收**。凡是本机没有真机读数的格子，标 `[待读数]`，
> 不拿手册当读数糊过去。

---

## 0. 一句话结论

把 `Scripting.FileSystemObject` 做成 **C3 自带对象**（RTL 里一枚 IDispatch 实现，`vb6com_fso.c`），
与内置 `Scripting.Dictionary`、`VBScript.RegExp` 同一条腿：目标机不装 Scripting Runtime
（`scrrun.dll`）也能跑，语义由我们这边钉死，而不是由机器上装的是哪个版本的 scrrun 说话。

**不做的**（写清楚，免得下一刀顺手扩）：`IFileSystem3`/`IFileSystem2` 的 `GetObject`、
`ShortPath`/`ShortName` 的 8.3 折叠在 NTFS 关掉 8.3 时的边角、`CopyFile` 的进度回调、
`TextStream` 的 Unicode 写（`CreateTextFile(..., True)` 那一格先按 BOM 落地，见 §5.5）。

---

## 1. 现状与动机

- 语料里 **12 份** 直接 `CreateObject("Scripting.FileSystemObject")`：
  `tests/m7_test/M7Main.bas`、`tests/test_com{,2,3,_optional}.bas`、
  `tests/test_earlybound{,2}.bas`、`tests/test_not_com.bas`、`tests/test_p1324.bas`、
  `tests/test_p24.bas`、`tests/declare_out/declare_cwex_lit_com.bas`。
- 用到的成员面（本轮普查，去重后 14 个）：`FileExists` / `FolderExists` / `BuildPath` /
  `GetAbsolutePathName` / `GetFileName` / `GetExtensionName` / `GetSpecialFolder` / `GetFolder` /
  `Drives`（含 `For Each`）/ `Folder.Name` / `Folder.Path` / `Drive.*` / `OpenTextFile` / `TextStream.*`。
- 今天这条腿落到 `scrrun.dll`：开发机能跑，目标机缺 Scripting Runtime 时 **错误 429**
  （`ActiveX component can't create object`）。Dictionary 在第二十六刀之前也是这个下场。
- 编译期还有一条隐性依赖：`driver_frontend.cpp` 的 `commonProgIds` 里就列着
  `Scripting.FileSystemObject`，靠机器上注册的 scrrun 类型库给 `Dim fso As FileSystemObject`
  供符号。**目标机缺 scrrun 时这一条也要有兜底**（§7.1，本刀先记录，不改）。

---

## 2. 接线点（照 Dictionary 的 5+2 处抄，别自己发明）

| # | 位置 | 做什么 |
|---|---|---|
| 1 | `CMakeLists.txt`（`vb6com_dict.c` 那一段旁边） | 编译器自身构建里加入 `vb6com_fso.c` |
| 2 | `src/driver/c3rtl.rc` | `229 RCDATA "../rtl/core/vb6com/vb6com_fso.c"` |
| 3 | `src/driver/driver_link.cpp` | `sourceFiles.push_back(rtlDir + "/vb6com_fso.c")` |
| 4 | `src/driver/rtl_embedded.cpp` | `{ RTL_VB6COM_FSO_C, "vb6com_fso.c" }` |
| 5 | `src/driver/rtl_embedded.hpp` | `RTL_VB6COM_FSO_C = 229`（228 是 Dictionary，229 空着） |
| 6 | `src/rtl/core/vb6com/vb6com.c` 的 `vb6_CreateObject` | 内建拦截（`C3_FSO_PREFER_MS=1` 时让给 scrrun，做对照组） |
| 7 | `src/rtl/core/vb6com/vb6com_collection.c` 的 `vb6_NewBuiltinObject` | `New Scripting.FileSystemObject` 那条腿 + `vb6_ComRememberProgID`（否则 `TypeName` 答 "Object"） |

外加 `vb6com.h` 三枚声明（`vb6_Fso_New` / `vb6_Fso_IsProgramId` / `vb6_Fso_IsClassName` /
`vb6_Fso_IsInstance`），以及 `src/common/owned_temp_keeps.hpp` 若新出口带 BSTR 形参要过 K13d
（本刀只出对象，理论上不新增 owned 出口；一旦新增，**先跑 `check_owned_temp_keeps.ps1`**）。

---

## 3. 对象模型

scrrun 的 FSO 是 8 个 coclass，彼此靠 `IDispatch` 组合，**没有**可变状态需要共享
（唯一例外：`TextStream` 的游标）。所以我们按 8 个独立对象实现，各自一枚 `Dispatch` 表：

```
FileSystemObject
├─ Drives → Drive   (IEnumVARIANT)
├─ GetDrive/GetFolder/GetFile
├─ Folder → Folders/Files/SubFolders (各自 IEnumVARIANT)
├─ File
└─ TextStream（游标：位置 / 行号 / 列号 / AtEndOfStream / AtEndOfLine）
```

每个对象用 Dictionary 那套骨架（`vb6com_dict.c` 的 1:1 抄法）：
`QueryInterface/AddRef/Release` + `GetTypeInfoCount/GetTypeInfo`（答 0/E_NOTIMPL）+ `GetIDsOfNames`
（名字→DISPID 的大小写不敏感表）+ `Invoke`（DISPID switch，`DISP_E_MEMBERNOTFOUND` 兜底）。

**DISPID 编号**：不追 scrrun 的真实编号（那要靠类型库），本刀用**自有编号表**，
`GetIDsOfNames` 是唯一入口 —— 与 Dictionary 当前做法一致。风险：走 `--early-binding`
类型库直接发 DISPID 的路径会打不中（§7.3）。

---

## 4. 成员清单与分刀

按"语料里用得上的先做"排序。每格后面标注口径来源：`读` = 真机读数、`册` = 手册、
`待` = 待读数（先按手册/常识落地，读数回来再改）。

### 刀 1（本刀，落地骨架 + 纯路径面）

| 成员 | 返回 | 口径 |
|---|---|---|
| `BuildPath(path, name)` | String | 册：`\` 结尾去重、无结尾补 `\`；`name` 以 `\` 开头则直接拼 |
| `GetFileName(path)` | String | 册：最后一个 `\`（或 `:`）之后；路径以分隔符结尾返回空串 |
| `GetBaseName(path)` | String | 册：`GetFileName` 去掉最后一个 `.` 及其后 |
| `GetExtensionName(path)` | String | 册：最后一个 `.` 之后（无点 → 空串） |
| `GetParentFolderName(path)` | String | 册：去掉最后一段；根返回空串 |
| `GetAbsolutePathName(path)` | String | 册：相对路径按**当前目录**展开，`.`/`..` 折叠 |
| `FileExists(path)` | Boolean | 待：目录名传进去要答 False（不是 True） |
| `FolderExists(path)` | Boolean | 待：文件名传进去答 False |
| `DriveExists(spec)` | Boolean | 册：`"C"` / `"C:"` / `"C:\"` 三形等价 |
| `GetTempName()` | String | 待：形如 `radXXXXX.tmp`（长度/字符集要读数） |
| `GetSpecialFolder(n)` | Folder | 待：0=Windows、1=System、2=Temp（本机 `GetSpecialFolder(0).Path` 是 `C:\Windows`，见语料 `declare_cwex_lit_com.bas:35`） |

### 刀 2（对象面）

`GetDrive` / `GetFolder` / `GetFile` / `Drives` / `Folder.{Name,Path,ShortPath,Type,Size,DateCreated,DateLastModified,DateLastAccessed,Attributes,Files,SubFolders,IsRootFolder,ParentFolder,Drive,CreateTextFile,Delete,Copy,Move}` /
`File.{Name,Path,ShortPath,Size,Type,DateCreated,DateLastModified,DateLastAccessed,Attributes,ParentFolder,Drive,Copy,Move,Delete,OpenAsTextStream}`。

### 刀 3（TextStream）

`CreateTextFile` / `OpenTextFile` / `TextStream.{Read,ReadLine,ReadAll,Skip,SkipLine,Write,WriteLine,WriteBlankLines,Close,AtEndOfStream,AtEndOfLine,Line,Column}` + `Files` 集合的 `IEnumVARIANT`。

### 刀 4（写面 + 集合增删）

`CreateFolder` / `DeleteFile` / `DeleteFolder` / `CopyFile` / `CopyFolder` / `MoveFile` / `MoveFolder`。

### 刀 5（默认翻面）

成员面覆盖语料后把**默认**从 scrrun 翻到内建（`C3_FSO_PREFER_MS=1` 保留为对照组），
并补一条 `fso_native` 夹具（真机 vs C3 双跑对账）。

---

## 5. 语义口径（逐条写清"按什么写"）

### 5.1 路径代数

- 分隔符只认 `\`；`/` 在 VB6 的 FSO 里**不是**分隔符（`GetFileName("C:/a/b.txt")` 待读数确认）。
- `BuildPath("C:\Windows", "System32")` = `C:\Windows\System32`；`BuildPath("C:\Windows\", "System32")`
  不得出现双反斜杠；`BuildPath("C:\Windows", "\System32")` 待读数（手册说"name 带前导反斜杠则忽略 path 的尾部分隔"）。
- `GetAbsolutePathName(".")` 的基准是**进程当前目录**（VB6 的 `CurDir`），不是 `App.Path`。
  语料 `test_p24.bas:17/26` 与 `test_earlybound2.bas:49` 都在读这一格 —— 真机读数必须连"启动时 CWD"一起记。
- 不发 `Drive` 前缀的 `..` 折叠越界（`C:\..` → `C:\`）。

### 5.2 存在性与错误

| 场景 | 期望 | 来源 |
|---|---|---|
| `FileExists("C:\__nonexistent_file__")` | `False`，**不报错** | 语料 `test_not_com.bas:9` 断言 |
| `FolderExists("C:\__nonexistent_folder__")` | `False` | 同上 `:16` |
| `GetFolder("X:\nope")` | 错误 **76** `Path not found` | 待 |
| `GetFile("C:\nope.txt")` | 错误 **53** `File not found` | 待 |
| `OpenTextFile("C:\nope.txt")` | 错误 **53** | 待 |
| `CreateFolder` 已存在 | 错误 **58** `File already exists` | 待 |
| 无权限 | 错误 **70** `Permission denied` | 待 |
| 路径含非法字符 | 错误 **52** `Bad file name or number` | 待 |

错误走 `vb6_RaiseError(n)`（RTL 既有出口），`Err.Description` 用 VB6 的原文（要读数逐条钉）。

### 5.3 日期与大小

- `DateLastModified` 等是 `Date`（Double 序列号）：FILETIME → 本地时间 → OLE 序列号。
  `[待读数]` 取舍：秒以下是否截断（VB6 的 FSO 文档说"精度到秒"，要真机验）。
- `File.Size` = 字节数（`Long`，>2GB 时 VB6 报错/溢出，待读数）。
- `Folder.Size` = **0**（VB6 文档：文件夹的 Size 无意义，返回 0），待读数。

### 5.4 `Type` 字符串

`File.Type` 取注册表 `HKEY_CLASSES_ROOT\.<ext>` 的默认值描述（如 `"Text Document"`），
取不到就回 `"<EXT> File"`；`Folder.Type` 恒 `"File folder"`；`Drive.Type` 见 §5.6。
`[待读数]` 这一格对 CJK 系统（本机中文）会返回中文描述 —— 夹具断言要按真机字符串写，不能写英文。

### 5.5 TextStream 编码与换行

- `CreateTextFile(path[, overwrite[, unicode]])`：默认 `unicode = False` ⇒ **ANSI（进程代码页）**，
  与 VB6 的 `Open ... For Output` 不是一回事（后者写 ANSI 也写 CRLF）。
- 换行恒 `CRLF`；`ReadAll` 把 `CRLF`、`LF`、`CR` 都当行尾（`ReadLine` 会吃掉行尾）。
- `unicode = True` ⇒ 写 UTF-16LE **带 BOM**（`FF FE`），`ReadAll` 要能识别 BOM。
- 读的时候 `Tristate` 参数（`OpenTextFile(path, iomode, create, format)` 的 `format`）：
  `-2` 系统默认 / `-1` Unicode / `0` ASCII，`[待读数]` 缺省 `-2` 的实际行为。

### 5.6 枚举顺序与 `Drive`

- `Drives` 按**盘符字母序**（`A:` → `Z:`），只列存在的。
- `Folder.Files` / `SubFolders` 的顺序：`[待读数]`（NTFS 上是索引序，不是字母序；夹具不能假设字母序）。
- `Drive.{DriveLetter,Path,RootFolder,TotalSize,FreeSpace,AvailableSpace,VolumeName,FileSystem,SerialNumber,IsReady,DriveType}`：
  `Path` = `"C:"`（**不带**反斜杠，要读数确认），`RootFolder.Path` = `"C:\"`。

---

## 6. 读数探针（ProbeFso.bas，交真机跑）

采集口径与 Double 那两轮一致：`App.Path` 落日志，一行一结论，C3 侧同一份源码双跑对账。

```vb
' ProbeFso.bas —— 内置 FSO 的口径读数（真 VB6 F5 或 VB6Mini /make 跑）
Option Explicit
Private fh As Integer
Private Sub L(s As String)
    Print #fh, s
End Sub
Sub Main()
    fh = FreeFile
    Open App.Path & "\c3_vb6_probe_fso.log" For Output As #fh
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    L "# fso probe " & Now
    ' A. 路径代数
    L "A01 " & fso.BuildPath("C:\Windows", "System32")
    L "A02 " & fso.BuildPath("C:\Windows\", "System32")
    L "A03 " & fso.BuildPath("C:\Windows", "\System32")
    L "A04 " & fso.GetFileName("C:\a\b.txt") & "|" & fso.GetFileName("C:\a\")
    L "A05 " & fso.GetBaseName("C:\a\b.tar.gz")
    L "A06 " & fso.GetExtensionName("C:\a\b.tar.gz") & "|" & fso.GetExtensionName("C:\a\b")
    L "A07 " & fso.GetParentFolderName("C:\a\b.txt") & "|" & fso.GetParentFolderName("C:\")
    L "A08 " & fso.GetAbsolutePathName(".") & "|" & fso.GetAbsolutePathName("..\x")
    L "A09 " & fso.GetFileName("C:/a/b.txt")
    L "A10 " & fso.GetTempName()
    ' B. 存在性（不报错的那几格）
    L "B01 " & fso.FileExists("C:\Windows\System32\cmd.exe")
    L "B02 " & fso.FileExists("C:\Windows")
    L "B03 " & fso.FolderExists("C:\Windows")
    L "B04 " & fso.FolderExists("C:\Windows\System32\cmd.exe")
    L "B05 " & fso.DriveExists("C") & fso.DriveExists("C:") & fso.DriveExists("Q:")
    ' C. 错误号（每格 On Error 单独接）
    On Error Resume Next
    Dim o As Object
    Set o = fso.GetFolder("Q:\nope"): L "C01 " & Err.Number & " " & Err.Description: Err.Clear
    Set o = fso.GetFile("C:\nope.txt"): L "C02 " & Err.Number & " " & Err.Description: Err.Clear
    Set o = fso.OpenTextFile("C:\nope.txt"): L "C03 " & Err.Number & " " & Err.Description: Err.Clear
    fso.CreateFolder "C:\Windows": L "C04 " & Err.Number & " " & Err.Description: Err.Clear
    Set o = fso.GetSpecialFolder(0): L "D01 " & o.Path & "|" & o.Name
    Set o = fso.GetSpecialFolder(1): L "D02 " & o.Path
    Set o = fso.GetSpecialFolder(2): L "D03 " & o.Path
    ' E. 对象面
    Set o = fso.GetFolder("C:\Windows")
    L "E01 " & o.Name & "|" & o.Path & "|" & o.Type & "|" & o.Size
    L "E02 " & o.DateLastModified & "|" & o.Attributes & "|" & o.IsRootFolder
    Set o = fso.GetFile("C:\Windows\System32\cmd.exe")
    L "E03 " & o.Name & "|" & o.Size & "|" & o.Type & "|" & o.ShortPath
    L "E04 " & o.DateCreated & "|" & o.DateLastModified
    Dim d As Object
    Set d = fso.GetDrive("C")
    L "F01 " & d.Path & "|" & d.RootFolder.Path & "|" & d.DriveType & "|" & d.FileSystem
    L "F02 " & d.TotalSize & "|" & d.FreeSpace & "|" & d.VolumeName
    ' G. TextStream
    Dim ts As Object, p As String
    p = App.Path & "\fso_probe_ansi.txt"
    Set ts = fso.CreateTextFile(p, True, False)
    ts.Write "a" & vbCrLf & "b": ts.Close
    Set ts = fso.OpenTextFile(p, 1)
    L "G01 " & ts.AtEndOfStream & "|" & ts.ReadLine & "|" & ts.Line
    L "G02 " & ts.ReadAll
    ts.Close
    p = App.Path & "\fso_probe_uni.txt"
    Set ts = fso.CreateTextFile(p, True, True)
    ts.Write "中文": ts.Close
    Set ts = fso.OpenTextFile(p, 1)
    L "G03 " & ts.ReadAll
    ts.Close
    L "# done"
    Close #fh
End Sub
```

`[待读数]` 不是"永远不测"，而是**本刀没有它就按手册落地，读数回来后单独一刀改口径并登记**——
与 Double 那两轮同一个纪律。

---

## 7. 风险与留白

1. **编译期类型名**：`Dim fso As FileSystemObject` 现在靠 scrrun 的类型库给符号。
   目标机/CI 无 scrrun 时，语义分析会报未定义类型。兜底方案（下一刀）：
   在内置类名表里登记 `FileSystemObject` 及其 7 个伴生类名，与 `vb6_Dict_IsClassName` 同款。
2. **早期绑定 DISPID**：走类型库直接发 DISPID 的路径不认我们自有的编号表 ⇒ `[待查]`
   是否需要在 typelib 导入阶段把 scrrun 的 coclass 换成"内建"标记。
3. **`C3_FSO_PREFER_MS` 对照开关**必须留着（Dictionary 的 `C3_DICT_PREFER_MS` 同款），
   否则真机对照跑不了。
4. **长路径 / UNC**：`\\?\` 前缀与 UNC `\\server\share` 的路径代数要单独读数，本刀不碰。
5. **默认翻面时机**：刀 5。在成员面没覆盖语料之前，**默认仍走 scrrun**，
   内建实现只在 `C3_FSO_NATIVE=1` 下启用 —— 语料 12 份的形状与行为一格不变。

---

## 7.5 刀1 实测（2026-11-18，本机）

`C3_FSO_NATIVE` 两条腿都验过（同一份 C3 编译的 exe，只差环境变量）：

- **默认（不设开关）**：`Set fso = CreateObject("Scripting.FileSystemObject")` 仍拿真 scrrun，
  `GetFolder("C:\Windows")` 正常、`Err.Number = 0`；
- **`C3_FSO_NATIVE=1`**：走 `vb6com_fso.c`，`GetFolder` 未实现 → 运行期打一行
  `vb6_ComCall: method "GetFolder" not found`，槽位为 `Nothing`。

实测发现两条要改的话：

1. **未实现成员不是 438**。`GetIDsOfNames` 答 `DISP_E_UNKNOWNNAME` 后，`vb6_ComCall` 的做法是
   **打一行到 stderr + 交 Nothing**，`Err.Number` 仍为 0；而 VB6 对不存在的成员报 438。
   §4 里写的「立刻报 438」**不准确**；要把 438 接到 `vb6_ComCall` 的 dispid 解析失败那一站，
   是**全局行为改动**（影响所有 COM 对象），不在本刀，先记账。
2. **`GetTempName` 的形状猜对了**：真 scrrun 答 `rad0ZMPV.tmp`，内建答 `radOS4MN.tmp` ——
   `rad` + 5 位大写字母/数字 + `.tmp`，且**不创建文件**。这一格可以从 `[待读数]` 提到「已对齐」。

14 行路径/存在性读数（`A01`~`A10`、`B01`~`B05`）两边**逐字相同**，唯一差异是 `GetTempName` 的
随机字母本身。过程中改掉一个真 bug：`GetAbsolutePathName` 折叠段时在盘符后没补反斜杠，
`C:\Users\x` 会展开成 `C:Users\x`（已修，修后与 scrrun 一致）。

## 8. 验收

- 刀 1/2/3 各自：`fso_native` 夹具（同一份 .bas 真机与 C3 双跑，逐行 diff）；
- `C3_FSO_NATIVE=1` 下：`tests/test_com{,2,3,_optional}.bas`、`test_earlybound{,2}.bas`、
  `test_not_com.bas`、`test_p1324.bas`、`test_p24.bas`、`m7_test` 全绿（这 12 份就是回归面）；
- 默认（不设开关）：全量 274 例零变化 + 发码清单零变化（内置对象只在运行期取代 scrrun，
  编译产物只多一条 `#include`，如果连这条也不加，形状完全不动）；
- 静态普查：`owned_temp_keeps`（BSTR 出口）、`builtin_opt_defaults`（若改 callee）、
  `manifest_coverage`。
