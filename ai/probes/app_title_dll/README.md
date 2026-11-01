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
