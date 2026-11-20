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
T0b 悬停第1页(有 tip-a) : 气泡可见=True 文本='tip-a'
T1 点第2页  : tabs=3  sel=2  cap=视图  clicks=1
T2 点第3页  : tabs=3  sel=3  cap=高级  clicks=2
T2b 点第1页 : tabs=3  sel=1  cap=常规  clicks=3
T3 加一页   : tabs=4  sel=1  cap=常规  clicks=3
T4 删当前页 : tabs=3  sel=1  cap=视图  clicks=3
T5 悬停第1页(空 ToolTipText)   : 气泡可见=False
T6 悬停第2页(带 tip-c)         : 气泡可见=True 文本='tip-c'
T6b 悬停第3页(空 ToolTipText)  : 气泡可见=False
T7 鼠标离开              : 气泡可见=False
RESULT: 正常关闭 exit=0
```

两个坑（都在脚本里留了留痕注释）：

- 气泡只能用**真光标**（`SetCursorPos`）验：用 `PostMessage(WM_MOUSEMOVE)` 假悬停会被
  `TrackMouseEvent(TME_LEAVE)` 当场判成“已离开”（真光标不在控件上），气泡一显就收。
- 不能拿 `TTM_GETTEXT` 读气泡文本：那条消息收 `TOOLINFO*`，跨进程时 USER32 不封送指针，
  目标进程直接挂（同理 `TCM_GETITEMRECT`）。本例自绘的气泡把文本放在窗口标题上，
  `WM_GETTEXT`（< WM_USER）会被封送，安全。

人眼该看到的三件事：

1. 标签条 3 页（常规 / 视图 / 高级），第 1 页底色**红**；切页时色块跟着翻（红 → 绿 → 蓝）。
2. 三个 `PictureBox` 严丝合缝贴在**页区**上 —— `ClientLeft/ClientTop/ClientWidth/ClientHeight`
   是相对 TagStrip 的页区矩形，不是整窗；`TabStrip` **不是容器**，页内容是自己摆上去的。
3. `clicks=` 每点一次 +1（`Click` 从通用子类来）。
4. 鼠标停在第 1 页 / 第 2 页上时，标签下方弹出一枚浅黄提示条（文字分别是 `tip-a` / 删页后的
   `tip-c`）；停在**没写 ToolTipText** 的页上没有气泡。

## 写法（不再是限制，是判据）

这份窗体现在故意写成 **VB6 教科书的原样语法**，一个变量中转都不用：

```vb
lblInfo.Caption = "sel=" & CStr(TabStrip1.SelectedItem.Index) _
                & "  cap=" & TabStrip1.Tabs(n).Caption
ShowPage TabStrip1.SelectedItem.Index
```

C30-2 收尾之前这两形会折成 COM 晚绑定（真跑 `Run-time error 438`），当时只能先落局部变量；
现在形状认在 AST 上（`tabStripChainViewOf`），读 / 写 / 类型三面同一条，所以这里写回原样。
README 旧版那条“先落变量”的约束已作废。

## 文件

| 文件 | 说明 |
|------|------|
| `TabManual.vbp` | 工程（`Startup="frmTab"`） |
| `frmTab.frm` | 窗体：TabStrip1 + 三个 PictureBox + 两个按钮 + 一个 Label |
