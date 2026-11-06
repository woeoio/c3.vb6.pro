# ActiveX DLL 里 `App.Title` 答的是谁（账 #605 的一句口径，给真 VB6 环境用）

第320刀（账 #592）把 `App.Title` 收成三面一值：读、写、`MsgBox` 默认题头都问 `vb6rtl_system.c` 里那一格，
真值 = `.vbp` 的 `Title=`，**入口模板播种**。播种那一步发在 **Exe 入口模板**的 5 个 `vb6_Init();` 之后 ——
于是 ActiveX DLL 工程（`Type=OleDll`）压根没人播种，`App.Title` 落回 `App.EXEName` 兜底那一档。
台账里这一格当时就写明是**在册缺陷（任务 #605）**，因为**修法取决于 VB6 那一侧答什么**，而本机没有 VB6。

## 要回答的一问

**在一枚 VB6 编出来的 ActiveX DLL 里读 `App.Title`，交回来的是下面哪一种？**

| 候选 | 探针会打成的样子（本工程的 `Title="AppTitleProbe-Title"`） |
|---|---|
| 甲 = 本 DLL 工程自己的 Title | `P1-title=[AppTitleProbe-Title]` |
| 乙 = 空串 | `P1-title=[]` |
| 丙 = 宿主进程（这里是 cscript）的标题 | `P1-title=[…宿主相关…]` |

P3/P4 是**同一问的对照**：`App.EXEName` 与 `App.Path` 答 DLL 还是答宿主，决定 P1 那一句该怎么读
（若 P3=`[cscript]` 而 P1=`[AppTitleProbe-Title]`，就说明 DLL 里的 `App` 对象是**两套口径混着**的，那更要写清）。
P5 问的是**运行期能不能写** `App.Title`（C3 那一侧第320刀实现了写；VB6 若报「只读」，那条实现就是超发）。

**Why it matters**：真语料 `D:\vb_yqt4qPac` 里 **48 份 `Type=OleDll` 工程，37 份带非空 `Title=`**，
其中**读 `App.Title` 的有 82 行**（最集中的一份是 `PrintRpt\frmPrint.frm:485` 往后连续十来行，
拿它拼窗口/报表题头）。也就是说这一档不是纸面缺口：按现在的 C3 编下去，那 82 行拿到的是 DLL 的文件名而不是 Title。

## 步骤（VB6 IDE + 一次 cscript）

1. 用 VB6 打开 `ai/probes/app_title_dll/AppTitleProbe.vbp`，**Make AppTitleProbe.dll**（编译到该目录）。
2. 注册它：`regsvr32 /s "…\app_title_dll\AppTitleProbe.dll"`。
   ⚠ `regsvr32 /s` 成功也**完全静默**，别拿「没弹框」当没注册；验法 =
   `reg query "HKCR\AppTitleProbe.Probe\CLSID"` 看有没有吐出 `{…}`。
   （没有管理员权限时的替代：先在 VB6 IDE 里 **F5** 跑这个 DLL 工程 —— IDE 在跑时 `CreateObject` 就能建起来，
   但那条路只能拿 P1/P2/P5，P3/P4 会被 IDE 自己污染，读数请标注「IDE 路」。）
3. 用 **32 位** cscript 跑客户端，并且**重定向到文件**再读（别在控制台里看）：
   `C:\Windows\SysWOW64\cscript.exe //nologo client.vbs > client_out.txt 2>&1`
4. 把 `client_out.txt` 的**五行编号行原样贴回来**即可。

打包三条坑（本目录的文件已按它们写好，改的时候别再踩）：
① `.vbp` 的 `Type=` 是**枚举值**，ActiveX DLL 只写 `Type=OleDll`，写成描述性文字 VB6 直接报错；
② 别抄 VB5 的 `Flavor=` 键（VB6 报「包含无效关键字」），`Reference=` 那几行的绝对路径可以整行删；
③ **位数**：VB6 产物是 32 位 COM，64 位 `cscript` 建不出来（报 `0x800A01AD`「ActiveX 部件不能创建对象」，
而注册其实是成功的）⇒ 必须 `SysWOW64`。

任何一条**编不过 / 运行期报错**，把 VB6 或 cscript 原话贴在那个编号上 —— 报错本身就是答案
（那就说明那一面在 DLL 里根本不存在，C3 也就不该播种它）。

## 回收状态（2026-11-02 读数到手 → 第330刀已发货）

用户代跑的原样五行存在 `client_out.txt`（宿主 = 32 位 cscript，非 IDE）：

```
P1-title=[AppTitleProbe-Title]
P2-len=19 eqEmpty=False
P3-exename=[AppTitleProbe]
P4-path=[D:\c3.vb6.pro\ai\probes\app_title_dll]
P5-after-write=[written-by-probe] err=0
```

| 编号 | 答的是哪一种 | 落成的口径 |
|---|---|---|
| P1/P2 | **甲** = 本 DLL 工程自己的 `Title=`（19 字符，非空） | `dll_entry.c` 的 `DllMain` 在 `vb6_Init()` 之后发 `vb6_SetAppTitle(L"<.vbp Title>")` —— 与 Exe 那五个入口模板逐字节同形，一位主人、两个入口，不另设第二枚 setter |
| P3 | DLL 自己的模块名（**不是** cscript） | `App.EXEName` 改问「本 RTL 静态链在哪枚模块」 |
| P4 | DLL 自己的目录（**不是** cscript 的目录） | `App.Path` 同上 —— 与 P3 共用一位主人 `vb6rtl_system.c::vb6_SelfModulePath()` |
| P5 | 运行期可写，`err=0` | 第320刀那条写实现不是超发，本刀沿用 |

P3/P4 这一对照是关键：它把「`App` 在 DLL 里到底指谁」从一问变成三问同答 —— VB6 的 `App` 那一族问的是
**执行这段代码的那枚模块**（msvbvm60 读的是模块自己那份工程头），不是进程。所以 C3 这一侧不能只在
`Title` 上播种，`Path`/`EXEName` 也必须按自身模块答，否则同一枚 `App` 会一半答 DLL、一半答宿主。

### 本刀没动的两档（在册，缺读数）

- `App.hInstance`：VB6 在 DLL 里答 DLL 的实例句柄还是宿主的，**没有读数** ⇒ 本刀刻意不动（仍 `GetModuleHandleW(NULL)`）。
- `App.PrevInstance`：DLL 里「已有另一实例」按谁算（同枚 DLL 被两个宿主加载？还是宿主自己起过两次？），**没有读数** ⇒ 不动。
  要清这两格，在 `ProbeCls.cls` 里各加一行把 `App.hInstance` 与 `App.PrevInstance` 打进编号读数，再发一次即可（同一枚 cscript 客户端）。

### C3 这一侧落地后的读数（`tests/b605_dllapp`，同一把尺）

```
BASE（改前）  READ_ANS=605-title=[disp_probe]      exe=[disp_probe]      tailOk=False
第330刀       READ_ANS=605-title=[B605 Dll Title]  exe=[B605k]           tailOk=True
```

`tailOk` 是 `App.Path` 的判别位：产物 DLL 单独编进带 `b605_self` 标记的子目录、探针住在外面 ——
两边同住一个目录的话，`App.Path` 改前改后同值，那一发根本不成立（判据设计上的坑，别再踩）。
## 追问三发（P6/P7/P8，2026-11-02 补）

第330刀已把 P1/P3/P4 落成代码（Title 由 `DllMain` 播种、`App.Path`/`App.EXEName` 问自身模块），
但同一次读数里剩下的两档**没有证据**，不敢猜；再加上一档是本刀翻 README 时新发现的缺口。

| 编号 | 问的是 | 为什么 C3 需要它 |
|---|---|---|
| P6 | `App.hInstance` 答的是**本 DLL 的 HMODULE** 还是**宿主进程**的？ | C3 今天答 `GetModuleHandleW(NULL)` = 宿主。真语料 48 份 OleDll 里读 11 行（`SmartMail` 5 行、`皮肤程序` 拿它 `LoadIcon`）。VB 代码常把它交给 `LoadIcon`/`CreateDialog` 找**自己**的资源，答错模块就是找不到图标/对话框 |
| P7 | `App.PrevInstance` 在一枚 DLL 里答什么、报不报错？ | C3 今天是拿 `GetModuleFileNameW(NULL)` 的文件名建互斥体（=按宿主算）。读 3 行。「本应用程序已有另一实例」在 DLL 语境里到底指谁，手册没说 |
| P8 | `App.Major` / `Minor` / `Revision` / `Build` 答什么？ | **C3 今天这三档一律折成字面 `0`**（`cgen_expr_member_form_builtin.inc`），而驱动里 `verMajor_/verMinor_/verRevision_` 已经从 `.vbp` 读进来了、只喂给了 VERSIONINFO 资源。读 8+8+6 行 —— 版本判定 `If App.Major >= 2` 那一族今天恒假 |

P6 里 `own=` 那一发是本刀自带的独立证人：`GetModuleHandleExW(6&, AddressOf ProbeHost.Anchor, h)`
—— `6 = GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS(4) | UNCHANGED_REFCOUNT(2)`，值取自本机
`Windows Kits\10\Include\10.0.19041.0\um\libloaderapi.h:281-283`（不背常数）。它问的是「这段代码
所在的那枚模块」，所以 `same=` 直接回答 P6 那一问；`own=[0]` 就说明 Declare 没调通，那一发不算数。

### ⚠ 这三发要重新 Make + 重新注册才能拿

`.vbp` 里把版本号改成了 **`MajorVer=7 / MinorVer=3 / RevisionVer=11`**（故意不是 0 —— 否则 P8 的
读数与「编译器压根没折」那一档分不开），并新增了标准模块 `ProbeHost.bas`（P6 的 `AddressOf` 锚点）。
所以请按 README 上面第 1~3 步**重 Make 一次 `AppTitleProbe.dll`、重跑一次 `regsvr32 /s`**，
再用 32 位 cscript 跑 `client.vbs`（现在会打八行）。P1~P5 顺带再量一次，正好当重复实验的证人。