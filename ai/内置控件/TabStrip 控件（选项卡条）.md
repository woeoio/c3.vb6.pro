# TabStrip 控件（选项卡条）

> 部件：Microsoft Windows Common Controls 6.0 (SP6)，MSCOMCTL.OCX

## 1. 概述
只画"标签条"本身的选项卡控件，**不是容器**（与 SSTab 的最大区别）：
页面内容要自己用 PictureBox/Frame 摆放，并在 Click 事件里切换显示。
代价换来了：标签可带图标、可多行、布局更自由。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Tabs` | 标签集合（Tab 对象：`Caption`、`Key`、`Index`） |
| `SelectedItem` | 当前选中的 Tab 对象（用 .Index 或 .Key 判断） |
| `ClientLeft / ClientTop / ClientWidth / ClientHeight` | **关键**：标签条下方可供放页面内容的矩形区域（缇） |
| `ImageList` | 关联 ImageList 后，Tab.Image 可给标签加图标 |
| `MultiRow` | True 时标签过多自动排成多行 |
| `HotTracking` | True 时鼠标悬停标签高亮 |
| `Enabled / Visible / ToolTipText` | 常规属性 |

## 3. 事件
`Click` —— **核心**：点击标签后触发，随后读 SelectedItem 切换页面；
另有 DblClick、KeyDown/KeyPress/KeyUp、OLEDragDrop 系列等常规事件。

## 4. 与 SSTab 的取舍

| | SSTab | TabStrip |
|---|---|---|
| 容器 | 是，子控件自动随页显隐 | 否，需自己切换页面容器 |
| 标签图标 | 不支持 | 支持（ImageList） |
| 多行标签 | TabsPerRow | MultiRow |
| 适用 | 快速做设置对话框 | 主界面多页签、需图标 |

## 5. 示例（PictureBox 数组做页面）

```vb
' 窗体上放 TabStrip1 和 picPage(0..2)（PictureBox 数组，BorderStyle=0）
Private Sub Form_Load()
    TabStrip1.Tabs.Add , "t1", "常规", 1      ' 第4参为 ImageList 图标索引
    TabStrip1.Tabs.Add , "t2", "视图", 2
    TabStrip1.Tabs.Add , "t3", "高级", 3
    Dim i As Integer
    For i = 0 To 2
        picPage(i).Move TabStrip1.Left + TabStrip1.ClientLeft, _
                        TabStrip1.Top + TabStrip1.ClientTop, _
                        TabStrip1.ClientWidth, TabStrip1.ClientHeight
        picPage(i).Visible = (i = 0)
    Next
End Sub

Private Sub TabStrip1_Click()
    Dim i As Integer
    For i = 0 To 2
        picPage(i).Visible = (i = TabStrip1.SelectedItem.Index - 1)
    Next
End Sub
```

> 坑：Client* 坐标是相对 TabStrip 的，Move 页面时要加上 TabStrip.Left/Top；页面容器建议用 PictureBox（可再嵌套控件）并设 BorderStyle=0 视觉无缝。

## 6. 实现状态（C30-2，已落地）

### 6.1 落在哪

| 层 | 文件 | 干什么 |
|----|------|--------|
| RTL | `src/rtl/core/vb6forms/vb6forms_tabstrip.c` | 原生 `SysTabControl32` 的整套复刻（建窗、灌项、读写回、事件判据） |
| 声明 | `src/rtl/core/vb6forms/vb6forms_prop_ctrl.h` | 全部出口原型（`check_rtl_proto_arity.ps1` 按这张表校实参个数） |
| 登记 | `rtl_embedded.cpp` / `rtl_embedded.hpp` / `driver_link.cpp` / `CMakeLists.txt` / `c3rtl.rc`（`226 RCDATA`） | 编进 c3rtl 并随 exe 取出来 |
| 解析 | `frm_parser*.cpp`、`form_pseudo.hpp` | `Begin MSComctlLib.TabStrip` → `FrmControlType::TabStrip` |
| 发码 | `cgen_util_ctrl.cpp`（读写表 + `controlPropType`）、`cgen_state.inc`（名字/句柄/下标解析）、`cgen_form_ctrl_style_apply.inc`（`vb6_TabStrip_Init` + 设计期 `BeginProperty Tabs`）、`cgen_expr_call_com_bind.inc`（`Tabs.Add/Remove`）、`cgen_call.cpp`（无实参 `Tabs.Clear` 的语句码头） | 把 VB6 语句折成上面的出口 |

### 6.2 支持面（1 基，与 VB6 一致）

- 方法：`Tabs.Add [index], [key], [caption], [image]`（四位全可选，省略位按位置对齐）、`Tabs.Remove n|key`、`Tabs.Clear`
- 读：`Tabs.Count`、`Tabs(i).Caption` / `.Key` / `.Index` / `.Image` / `.Selected`、`SelectedItem.Index` / `.Caption` / `.Key`
- 写：`Tabs(i).Caption` / `.Key` / `.Image`、`Tabs(i).Selected = True`、`SelectedItem.Index = n`、`SelectedItem.Caption` / `.Key`
- 属性：`MultiRow` / `HotTracking` / `ShowTips`（创建时进样式位）、`ClientLeft/Top/Width/Height`（页区矩形）、通用几何与 `Visible/Enabled`
- 事件：`Click`（走通用控件子类的 `WM_LBUTTONUP`，与“点了一下”的 VB6 语义同）

这些链式成员**在哪里都同样可用**：赋值右值、`CStr(...)` 之类内建函数实参、子过程/函数实参、`Debug.Print` 实参、`If` 比较项。形状认在 AST 上（`tabStripChainViewOf`，只认登记表里类型为 TabStrip 的控件），读 / 写 / 类型推断三面共用这一条判形 —— 不再存在“某种写法能跑、另一种不行”。

### 6.3 两处必须知道的实现口径

1. **`Click` 不接 `TCN_SELCHANGE`**。`TCN_SELCHANGE`（`-551`）只在选中项**真的换了**时来，拿它当 `Click` 会漏报（点已选中的标签也该响）。两条都接是**一次点击两声**。RTL 的 `vb6_TabStrip_OnNotify` / `SimSelect` 留着给无头判据与以后真要区分“变页”的场合，不在派发链上。
2. **`SelectedItem.Index` 回问原生**（`TCM_GETCURSEL`），不读缓存。早期版本答的是“我们上次设的值”，于是用户真点完标签读数恒 1（`verify_tab.ps1` 的 T1/T2 两行就是那条红的证人）。注意与 UpDown 那次的时序相反：原生 tab 是“先切、后通知”，通知里读到的已是新页。

### 6.4 已实测（实证读数，非推断）

无头夹具 `tests/ctrltabstrip`（`Test-Vbp` x64 + x86，`-ExpectExitCode 0`）：

```
TS1-tabs=3        TS2-sel=1         TS3-cap=True/True/True   TS3-len=2/2/2
TS4-key=k1        TS5-add=4/True    TS6-rm=3/True/True       TS7-type=Long/String
TS8-client=300x176                  TS9-flags=False/True/True
TS11-arg=3/1      TS12-sub=1        TS13-wcap=True/True/0
TS14-wsel=2/2     TS15-selset=1     TS16-rmkeep=2/2/True     TS17-rmcur=1/1
TS10-clear=0/0    TABSTRIP-DONE
```

TS11 / TS12 就是“VB 原样写法”那两形（`CStr(TabStrip1.Tabs.Count)`、`ShowIdx TabStrip1.SelectedItem.Index`），
TS13–TS17 是写侧与 `Remove` 语义；夹具本身不再用“先落局部变量”那种绕行写法。

手动样例 `tools/tabstrip_manual`（真鼠标消息，`.build/probe/verify_tab.ps1`）：

```
T0 载入后   : tabs=3  sel=1  cap=常规  clicks=0   [原生 count=3 curSel=0]
T1 点第2页  : tabs=3  sel=2  cap=视图  clicks=1
T2 点第3页  : tabs=3  sel=3  cap=高级  clicks=2
T2b 点第1页 : tabs=3  sel=1  cap=常规  clicks=3
T3 加一页   : tabs=4  sel=1  cap=常规  clicks=3
T4 删当前页 : tabs=3  sel=1  cap=视图  clicks=3
RESULT: 正常关闭 exit=0
```

`frmTab.frm` 现在是**纯 VB6 教科书写法**（`ShowPage TabStrip1.SelectedItem.Index`、
`CStr(TabStrip1.Tabs.Count)`、`CStr(TabStrip1.SelectedItem.Index)`、`TabStrip1.Tabs(n).Caption`），
不再有任何“先落局部变量”的绕行。

### 6.5 已知缺口（写在这里，不装看不见）

1. **控件数组元素上的链式访问**（`TabsArr(0).Tabs(1).Caption`）仍走旧路：句柄不是标量槽，`tabStripChainViewOf` 显式回 false。单枚控件实例（正常用法）全部支持。
2. **`Tabs(i).Selected = False` 不发码**：VB6 的 TabStrip 没有“全不选”状态，照 `TCM_SETCURSEL(-1)` 写会让原生控件清掉选中，反而比 VB6 多出可观测差异。`= True` 正常选中。
3. **`Tabs.Remove` 之后的选中项**：删前面的页时保持“同一页仍选中”（序号自动前移），删当前页时顶上来那一页接管选中（等同 VB6 “选其下一位”），末页被删则钳到新的末项。已由 TS16/TS17 钉住。
4. **`Tabs(i).ToolTipText` 尚未开**（RTL 侧有 `tip[]` 槽位，但没有读写出口与发码表项）。