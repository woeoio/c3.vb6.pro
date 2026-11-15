# tabstrip_manual — TabStrip 真鼠标手动验证样例

C30-2 的**手动**验证工程。它不进 `tests/`，所以不参与语料/清单/夹具普查扫描；它存在的唯一理由
是：TabStrip 的 `Click` 走的是“通用控件子类收 `WM_LBUTTONUP`”那条路，**只有真鼠标**才能把它验出来。

## 怎么跑

```powershell
# 1) 先有 C3.exe（仓库根）
cmd /c build.bat

# 2) 编这一份（输出目录要先存在）
New-Item -ItemType Directory -Force .build\tabman | Out-Null
.build\C3.exe tools\tabstrip_manual\TabManual.vbp -o .build\tabman\TabManual.exe

# 3) 手点
start .build\tabman\TabManual.exe
```

自动复现（真起进程、真读控件文本、真发鼠标消息，不需要人坐在机器前）：

```powershell
& .\.build\probe\verify_tab.ps1
```

脚本按 PID + 窗体类前缀 `VB6_Form_` 认窗口（不靠标题字串，编码变了也不误判），三项标签的
x 中心由 `.build/probe/scan_items.ps1` 扫出来（0–44 / 45–89 / 90–134 ⇒ 中点 20 / 65 / 110）。

## 预期读数（实跑结果，2026-11-15）

```
T0 载入后   : tabs=3  sel=1  cap=常规  clicks=0   [原生 count=3 curSel=0]
T1 点第2页  : tabs=3  sel=2  cap=视图  clicks=1
T2 点第3页  : tabs=3  sel=3  cap=高级  clicks=2
T2b 点第1页 : tabs=3  sel=1  cap=常规  clicks=3
T3 加一页   : tabs=4  sel=1  cap=常规  clicks=3
T4 删当前页 : tabs=3  sel=1  cap=视图  clicks=3
RESULT: 正常关闭 exit=0
```

人眼该看到的三件事：

1. 标签条 3 页（常规 / 视图 / 高级），第 1 页底色**红**；切页时色块跟着翻（红 → 绿 → 蓝）。
2. 三个 `PictureBox` 严丝合缝贴在**页区**上 —— `ClientLeft/ClientTop/ClientWidth/ClientHeight`
   是相对 TagStrip 的页区矩形，不是整窗；`TabStrip` **不是容器**，页内容是自己摆上去的。
3. `clicks=` 每点一次 +1（`Click` 从通用子类来）。

## 写法上的硬约束（别改回去）

凡是链式读要**当内建函数的实参**用，一律先落到同型局部变量：

```vb
Dim n As Long
n = TabStrip1.Tabs.Count          ' ✅ 赋值形 → vb6_TabStrip_TabsCount(...)
Debug.Print CStr(TabStrip1.Tabs.Count)   ' ❌ 仍走 COM 晚绑定，真跑 438
```

这一条与本目录的 `frmTab.frm` 头注释、`ai/内置控件/TabStrip 控件（选项卡条）.md §6.5` 是同一件事，
缺口修好之后三处一起改。

## 文件

| 文件 | 说明 |
|------|------|
| `TabManual.vbp` | 工程（`Startup="frmTab"`） |
| `frmTab.frm` | 窗体：TabStrip1 + 三个 PictureBox + 两个按钮 + 一个 Label |
