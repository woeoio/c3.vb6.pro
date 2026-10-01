# Slider 控件

# Slider 控件

               

**Slider** 控件是包含滑块和可选择性刻度标记的窗口。可以通过拖动滑块，用鼠标单击滑块的任意一侧或者使用键盘移动滑块。

**语法**

**Slider**

**说明**

在选择离散数值或某个范围内的一组连续数值时，**Slider** 控件十分有用。例如，无需键入数字，通过将滑块移动到刻度标记处，可以用 **Slider** 对被显示的图象设置大小。要选择某个范围内的数值，需将 **SelectRange** 属性设置为 **True** 并对控件编程，这样当按下 SHIFT 键时就可选择范围。

可以水平或者垂直地放置 **Slider** 控件。

**发行注意** 为了在应用程序中使用 **Slider** 控件，必须将 MSCOMCTL.OCX 文件添加到工程中。在分布应用程序时，应把 MSCOMCTL.OCX 文件安装到用户的 Microsoft Windows System 或 System32 目录中。关于如何把 ActiveX 控件添加到工程的详细信息，请参阅《程序员指南》。

## 本项目的实现口径（ai/029 C29-SL-a：窗口 + 创建样式 + 标量属性面）

`Slider` 在原生的 **`msctls_trackbar32`**（comctl32 自带窗口类，`ICC_BAR_CLASSES` 已在运行时初始化里请求过）上实现，语法面与本页上文一致。本页上文那条"必须把 MSCOMCTL.OCX 加入工程并拷到 System32"的发行注意 **在本项目里不适用**：该 OCX 只有 32 位，64 位进程里 `CoCreateInstance` 直接失败，所以本项目不引它、也不需要带它。

| 写法 | 读数与实现 |
| --- | --- |
| `Slider1.Orientation` | `0` = 水平（默认）、`1` = 垂直。就是原生样式位 `TBS_VERT`；**运行期改是有效的**（写样式位 + 一次 `SWP_FRAMECHANGED` 换帧，实测滑块形状与位移轴跟着换） |
| `Slider1.TickFrequency` | 下发原生 `TBM_SETTICFREQ`（先 `TBM_CLEARTICS` 再设，免得新旧刻度叠画）。**原生没有回读这条消息**，所以读回来的是本项目自存的那一档 |
| `Slider1.TravelIsVert` | **本项目的扩展读数，不是 VB6 属性**：问控件自己"这杆横着走还是竖着走"（读 `TBM_GETTHUMBRECT` 的滑块长短边）。判据用它把"样式位写进去了"升级成"控件真按那一档在走" |
| `Slider1.TickPresent` | **本项目的扩展读数，不是 VB6 属性**：`TBM_GETTICPOS(0)` 问"当前到底画没画刻度"。实测光挂 `TBS_AUTOTICKS` 而不给 `TBM_SETTICFREQ` 是**不画**刻度的（默认频率 0） |
| `Enabled` / `Visible` | 通用窗口状态（`EnableWindow` / `ShowWindow`），读回是 VB6 的 `True`/`False` |

**本批还没有的一面**（`Min` / `Max` / `Value` / `SmallChange` / `LargeChange` / `SelStart` / `SelEnd` / `SelectRange` 与 `Change` / `Scroll` 两条事件）在 ai/029 的 C29-SL-b / SL-c 两格里排；`TickStyle` 四档当时写着"本机拿不到 VB6 枚举真值（OCX 未注册、类型库读不到）"—— **那句是错的**，见下面 C29-SL-h 那一格：类型库不必注册就能从文件里读，四档与它的数值都已按真值实现。

**两个坑（实测踩过的）**：① 判方向**别问 `TBM_GETCHANNELRECT`** —— 它返回的矩形永远把行程长度放在 x 分量上，水平杆与垂直杆答同一组数，量它等于什么都没量；② `TBM_GETTHUMBRECT` / `GETCHANNELRECT` 的**返回值不是成功标志**（实测返回 0 而矩形填得好好的），只看矩形内容。

## 本项目的实现口径（ai/029 C29-SL-b：`Min` / `Max` / `Value` / `Small·LargeChange` / `Sel*`）

值面**全部直问直发控件**，没有一格自存。原生对应：`TBM_SETRANGE`/`GETRANGEMIN`/`GETRANGEMAX`、`TBM_SETPOS`/`GETPOS`、`TBM_SETLINESIZE`/`GETLINESIZE`、`TBM_SETPAGESIZE`/`GETPAGESIZE`、`TBM_SETSEL`/`GETSELSTART`/`GETSELEND`；`SelectRange` 是样式位 `TBS_ENABLESELRANGE`。

| 写法 | 读数与实现 |
| --- | --- |
| `Slider1.Min` / `.Max` | 原生这条消息只吃 **16 位**（`lParam` 是两个半字；实测 40000 会截成 -25536），所以下发前钳到 ±32767，**读回也是那一个钳过的值** —— 答出去的与控件真走得动的始终是同一个数。VB6 的 `Min`/`Max` 是 Long（类型库里这一对写的是 `VT_I4`，见下面 C29-SL-h 那一格），超界那一档 OCX 怎么办本机拿不到真值（跑不起来），押后 |
| `Slider1.Value` | `TBM_SETPOS`/`GETPOS`。**越界交给控件钳**（量程 10..100 时 `Value = 500` 读回 100、`= 5` 读回 10），我们不自己钳第二遍 |
| `Slider1.SmallChange` / `.LargeChange` | 原生 line / page 尺寸，真往返。实测默认档是 **1 / 20**（VB6 文档写 1 / 5）：小的一条对得上，大的那条本项目**照原生答 20**，不拿文档去改控件的读数 —— 等拿到 VB6 真值再拍 |
| `Slider1.SelectRange` | 样式位 `TBS_ENABLESELRANGE`，**运行期可改**（实测关掉之后 `CStr` 就答 `False`）。类型是 Boolean，所以 `CStr(sld.SelectRange)` 打 `True`/`False`、装箱是 `VT_BOOL` |
| `Slider1.SelStart` / `.SelEnd` | 只有 `SelectRange = True` 时设进去才生效（实测没挂那位时 `TBM_SETSEL` 答不回来）。原生"没设过"答 `-1` ⇒ getter 折成 **0**（无区段）。两端互相顶：起点越过终点时终点跟上来。**注意 `SelEnd` 不是 VB6 的名字** —— 类型库里那一对话是 `SelStart` + `SelLength`（见下面 C29-SL-i 那一格），`SelEnd` 是本项目按原生 `TBM_SETSELEND` 加的口，因为存量代码在用就留着了 |

**改设计期值面要注意顺序**：动 range 会让控件连带重算 `pos`、`page`、`selstart`（实测把 range 从 0..100 收到 10..100，`page` 从 20 变 18、`selstart` 跟到 10）。`vb6_Slider_Init` 的参数序因此是定死的：**range → line/page → pos → 刻度 → Sel**。

## 本项目的实现口径（ai/029 C29-SL-c：`Change` / `Scroll` 两条事件）

两条事件**不走 `WM_NOTIFY`**，走的是与 ScrollBar 同一条通道：控件给**父窗**发 `WM_HSCROLL`（横杆）或 `WM_VSCROLL`（竖杆），`LOWORD(wParam)` 是原生的 `TB_*` 码、`HIWORD(wParam)` 带当前值、`lParam` 就是控件句柄（派发按它认来源）。

一次真拖拽实测是这一串：`TB_THUMBTRACK(5)`×N → `TB_THUMBPOSITION(4)` → `TB_ENDTRACK(8)`；方向键是 `TB_LINEUP(0)` → `TB_ENDTRACK(8)`。

| 写法 | 触发与实现 |
| --- | --- |
| `Slider1_Change()` | 按"**值真的变了**"发（控件当前值 vs 上次派发时的基准）。所以拖拽过程中那一串 5 会连续触发，而落点那条 4 与收尾那条 8 不会再多发一次 |
| `Slider1_Scroll()` | 只认滑块那两档（码 4 / 5），与本项目 HScrollBar / VScrollBar 已发货的口径同一档（同一条原生通道不允许两套分法）。"点轨道 / 按方向键算不算 Scroll" 本机拿不到 VB6 真值 ⇒ 押后 |
| `Slider1.SimNotify(code, pos)` | **判据专用助手，不是 VB6 方法**（与 `DTPicker.SimChange` / `MonthView.SimDateClick` / `RichTextBox.SimNotify` 同先例）：先把值推到 `pos`（真手势都是控件先动、再发通知），再按原生那一档发一条**真**消息进父窗 |

**程序化写 `Value` 不会触发 `Change`**：实测 `TBM_SETPOS` / `TBM_SETRANGE` 一条通知都不发，本项目刻意照原生、不伪造（VB6 那颗 OCX 里这一走会 raise）。判据里 `SC10`/`SC11` 是这条口径的哨兵，哪天要齐平就得翻红逼人来拍。

## 本项目的实现口径（ai/029 C29-SL-d：常规事件面 `Click` / `DblClick` / `KeyDown` / `KeyUp`）

这一档**不走父窗那两条滚动通知**，走的是控件自己的子类过程（子类化换的是那一枚 HWND 的 `WNDPROC`，所以发给这个窗口的每一条消息都先过我们这一段，再转给原生过程）。

| 写法 | 原生落点与实现 |
| --- | --- |
| `Slider1_Click()` | `WM_LBUTTONUP`。实测真手势的下/抬两条都会到子类过程；裸的一条抬起**不会**惊动父窗那条通道（值不动 ⇒ `Change`、`Scroll` 都不跟） |
| `Slider1_DblClick()` | `WM_LBUTTONDBLCLK`，与 `Click` 是两档：合成一条双击消息只点 `DblClick`，`Click` 的计数不涨 |
| `Slider1_KeyDown(KeyCode, Shift)` | `WM_KEYDOWN`，`KeyCode` 就是 `wParam`（`VK_LEFT` 与 `vbKeyLeft` 同为 37）。**按键会让控件真动**（走 `SmallChange` 那一档）并发父窗 `TB_LINEUP(0)` ⇒ `Change` 会顺带跟一次；紧跟的 `WM_KEYUP` 发 `TB_ENDTRACK(8)`、值不再动 |
| `Slider1_KeyUp(KeyCode, Shift)` | `WM_KEYUP`，同上 |
| `Slider1_GotFocus()` | 控件自己的 `WM_SETFOCUS`（还是那条子类化的路）。这一档以前只对 PictureBox / Frame / Label / Image 四类发，Slider 挂上是**编得过、永不触发**的死代码；现在改成按"有没有另一条原生来源"排除 —— TextBox / ComboBox / ListBox / 命令按钮 / 复选 / 单选这六类的焦点通知走父窗的 `WM_COMMAND` 码（`EN_SETFOCUS=256`、`CBN_SETFOCUS=1024`、`LBN_SETFOCUS=4`、`BN_SETFOCUS=6`），其余窗口态控件（Slider / ListView / TreeView / DTPicker / MonthView / RichTextBox / 两个滚动条 / 文件系统三件套）一律补齐 ⇒ 同一次焦点变化只会有一处发 |
| `Slider1_LostFocus()` | 控件自己的 `WM_KILLFOCUS`，同上 |

四条实测口径值得记：

1. **不需要 `WS_TABSTOP` 也能收到按键**：原生轨道条在 `WM_LBUTTONDOWN` 里自己就把焦点抢过去了（实测：按下之后跟着一条 `WM_SETFOCUS`，"focus after click = 控件"）。所以"按一下方向键调音量"这种写法在本项目里是真的会跑。**Tab 键导航**这一半后来补齐了一半：创建样式现在立 `WS_TABSTOP`（账 #83(a)）、窗体显示时焦点自己落在第一枚 tabstop（账 #157），**模态窗体里按 Tab 实测已经真跳格**；普通（非模态）窗体那条泵仍没调 `IsDialogMessage`，还差这一刀（账 #83(b)）。
2. **`Click` 这一档是通用的、不是 Slider 专属**：以前"控件级 `_Click` 处理器"只有那种会往父窗发原生通知的控件（命令按钮 / 复选 / 单选 / 列表 / 组合框 / 文件系统三件套 / SSTab / 工具栏）才会被接上，其余控件（Label / Image / PictureBox / Frame / TextBox / 滚动条 / Slider）的 `xxx_Click` 是**编得过、永远不被调用**的死代码。本格把这些补上了，同时按上表把那批"已有原生 Click 来源"的控件排除掉 —— 否则一次点击会从两条路各发一次。
3. **`Slider1.SimStdEvent(kind, wParam)` 是判据专用助手，不是 VB6 方法**（与 `SimNotify` / `DTPicker.SimChange` / `RichTextBox.SimNotify` 同先例）：`kind` 0=Click 1=DblClick 2=KeyDown 3=KeyUp，把对应的那条原生消息**同步**送进控件自己的过程。无头环境点不了鼠标，而直接调处理器会绕开整条派发链。
4. **焦点不靠 `WS_TABSTOP` 也到得了，但 Tab 导航仍然不通**：原生轨道条在 `WM_LBUTTONDOWN` 里自己就把焦点抢过去（实测按下之后紧跟一条 `WM_SETFOCUS`），所以"点一下再用方向键调"这种写法真会跑；而 Tab 键切换要消息泵里有 `IsDialogMessage`：模态循环本来就带（现在实测真跳格），普通窗体那条主泵没有 ⇒ 只剩这一刀，记在账 #83(b)。

## 本项目的实现口径（ai/029 C29-SL-e：`ToolTipText` 与 `Tag`）

这两条是**所有可见控件通用**的属性（不只 Slider），本项目走同一对 RTL 入口：
`vb6_SetToolTipText` 把文本存进窗口属性 `VB6_ToolTipText`，并顺手把它注册到共享的 tooltip 控件
（`TOOLTIPS_CLASS`，`TTS_ALWAYSTIP`）；`vb6_SetControlTag` 同理存 `VB6_Tag`。

| 写法 | 读数与实现 |
| --- | --- |
| `Slider1.ToolTipText = "音量"` | 存一份拷贝（`SysAllocString`）并注册工具项；`CStr(Slider1.ToolTipText)` 打回 `音量`、`TypeName` 答 `String`、`VarType` 答 `8`。**没写过就读答案是空串**（VB6 同） |
| `Slider1.Tag = "sld-a"` | 同上，存在窗口属性里；`If Slider1.Tag = "sld-a"` 走字符串比较 |

两条以前都坏在**档位**上：读写面早就登记了，但类型表里没有它们，而 getter 的 C 返回型写 `void*`
⇒ 装箱那张 `_Generic` 表把"其他指针"送去 `VariantObject`，于是 `CStr` 打空、`TypeName` 答 `Object`、
`Tag` 的比较答假，而同一枚属性的 `Len` / `InStr` / 赋值这几条反而是对的。现在归到 `String` 档、
返回型改 `wchar_t*`，消费面统一。

**登记那一步也有人验了（ai/029 C29-SL-j，账 #148）**：上面那句"顺手注册到共享的 tooltip 控件"以前
只验过"串存得住"，没验过"工具真进宿主了没有"。现在加了一枚扩展读数 `Slider1.ToolTipRegistered`
（**不是 VB6 属性**，与 `TickPresent` 同族）—— 拿 `TTM_GETTEXT` 向宿主把工具文本读回来，
读得到才算登记成立。产物真跑：设过的是 `True`、从没设过的是 `False`。
**一条探针教训值得记**：裸编的 C 探针（不嵌 Common-Controls 6.0 的 manifest ⇒ 走 comctl v5）里
`TTM_ADDTOOLW` 直接返回失败，照那个读数会判"气泡从来没注册上"—— 那是**版本差异不是缺陷**，
凡是"登记/渲染"这类行为，判据要问在产物里（同 C29-5a / C29-V6 那两轮的旧账）。

**设计期那一条也接上了（ai/029 C29-SL-f）**：`.frm` 里写的 `ToolTipText = "..."` / `Tag = "..."` 会在建好窗口之后发一次对应的 setter（两条创建路共用同一趟），所以读回来就是 `.frm` 里那个值；写了空串或干脆没写都**不发**（与默认读数同为 `""`，多发一条只是改产物）。资源引用形态（`"frmTest.frx":0000`）不是文本，跳过不发。

## 本项目的实现口径（ai/029 C29-SL-h：`TickStyle` 四档 + `GetNumTicks`）

**先把"拿不到 VB6 真值"那一句订正掉**：OCX 里嵌的那张类型库**不需要注册**就能读 —— 用
`LoadTypeLibEx` 传 `REGKIND_NONE` 加上文件路径，132 张类型表全出来了（探针 `.build/slprobe/sltlb.cpp`，
x86 编、按路径读文件）。`TickStyleConstants` 的四个成员名与
数值就是这么读出来的，不是照文档猜的。

| 写法 | 读数与实现 |
| --- | --- |
| `Slider1.TickStyle` | VB6 枚举 `TickStyleConstants`：`0 = sldBottomRight`（默认，那三位样式位全清）、`1 = sldTopLeft`（`TBS_TOP`；同一位在竖杆上念作 `TBS_LEFT`）、`2 = sldBoth`（`TBS_BOTH`）、`3 = sldNoTicks`（`TBS_NOTICKS`）。读写都是**窗口当前的样式位**，不是自存：`CStr` 打 0..3、`TypeName` 答 `Long`、`VarType` 答 `3`。越界值（`4`、`-1`…）落 `0` 那一档 ⇒ 答出去的就是控件真在走的那一档（与 `Orientation` 同一口径） |
| 设计期 `TickStyle = 2` | 直接进 `CreateWindow` 的样式参数。折算只有一处（`sliderStyleBits`），**两条创建路共用**：顶层控件与容器（Frame / PictureBox）里的子控件都算。以前第二条创建路对 Slider **一格样式都不挂**（创建参数实测是 `1342177280` = 只有 `WS_CHILD|WS_VISIBLE`，连默认的 `TBS_AUTOTICKS` 都没有）—— 账 #83 那条缺口的一个具体落点，本批顺手补上 |
| `Slider1.GetNumTicks` | VB6 那一面（dispid `0x000f`、只有 propget ⇒ 只读），就是原生 `TBM_GETNUMTICS`。它同时是 `TickStyle = 3` 唯一的控件侧证人：实测那一位一挂，读数从 11 变 0，而 `TickPresent`（`TBM_GETTICPOS(0)`）**照旧答"有刻度"**（刻度只是不画、那张表还在） |
| `Slider1.ChannelTop` | **本项目的扩展读数，不是 VB6 属性**：`TBM_GETCHANNELRECT` 矩形的上边。刻度画在哪一侧，通道就被顶下去几像素，判据用它证"这一档真到了控件、而且控件真按它重排了" |

三条实测口径值得记：

1. **`TickStyle = 1` 与 `= 2` 之间没有稳的几何维度**：两档的通道位置只差一两个像素，而且**谁高谁低
   随控件高度翻面**（探针在 40px 高答 20 / 19，夹具那枚 400 缇答 18 / 19）。判据因此只写"与 `0` 那档
   不同"，1/2 两档的区分全靠样式位读回 —— 第一版把 "1 > 2 > 0" 写进判据，当场就翻红了。
2. **运行期写这一档是真有效的**：写样式位 + 一次 `SWP_FRAMECHANGED` 之后，每一条形与"创建时就带那
   一位"**逐字相同**（来回切四档也一样）；只写样式位不换帧则停在旧布局，看着像没生效。
3. **竖杆上那一位念 `TBS_LEFT`**：commctrl 里 `TBS_TOP == TBS_LEFT == 0x0004`、
   `TBS_BOTTOM == TBS_RIGHT == 0x0000`，VB6 那两条合名（BottomRight / TopLeft）照着这一点起 ——
   所以这张对应表是 1:1 的，不是凑的。

**类型库顺带读出来、本项目还没做的几面**：`SelLength`（dispid `0x0008` —— VB6 的选区其实是
`SelStart` + `SelLength`，`SelEnd` 这个名字是本项目早先自己加的）、`ClearSel`（`0x000e`，方法）、
`Text`（`0x0010`，BSTR：拖动时那颗气泡里显示的字符串）与 `TextPosition`（`0x0011`，枚举
`sldAboveLeft = 0` / `sldBelowRight = 1`），再加事件面的 `KeyPress` / `MouseDown` / `MouseMove` /
`MouseUp`。分别记在账 #146、#147 与 #141。

## 本项目的实现口径（ai/029 C29-SL-i：选区的 VB6 那一面 —— `SelLength` 与 `ClearSel`）

类型库（`ISlider` 的 dispid 表）给的是 **`SelStart`(0x0007) + `SelLength`(0x0008)** 这一对，两条都 `VT_I4`，
另有一条方法 **`ClearSel`(0x000e)**，文档原话 "Sets the SelLength to 0"。**VB6 那一面没有 `SelEnd` 这个名字** ——
本页上文表格里那条 `SelEnd` 是本项目早先（C29-SL-b）按原生 `TBM_SETSELEND` 自己加的口，因为存量夹具在用，
留着不撤，这一格把 VB6 真的那一对照着补上。

| 写法 | 读数与实现 |
| --- | --- |
| `Slider1.SelLength` | **起点不动、终点 = 起点 + 长度**，走原生 `TBM_SETSEL`；读回是两端之差，**空区段一律 0**。远端超出量程时由**控件夹住**（实测量程 10..100 写 900 落 100），我们不钳第二遍。类型档是 `Long`（`CStr` 打数字、`VarType` 答 `3`） |
| `Slider1.ClearSel()` | 直发 `TBM_CLEARSEL`。清完之后 `SelStart` / `SelEnd` / `SelLength` 三条都读回 `0`（原生那两端答 `-1`，本项目折成 0，与 SL-b 同一条） |
| 设计期 `SelStart = 25` + `SelLength = 15` | `.frm` 写的是 VB6 那一对时，创建那一趟折成 `(起, 止) = (25, 40)` 下发。**改之前 `SelLength` 整条被丢**（创建参数里那一位恒 `-999`），真 VB6 工程的选区建起来就是空的 —— 运行期写口一直是好的，缺口只在设计期那一趟，与 `ToolTipText`/`Tag`（账 #142）同一形状 |

三条实测口径（探针 `.build/slprobe/slmeasure13.c`，量程 10..100）：

1. **"没碰过"与"已清空"是两种空态，起点还不一样**：从没写过的控件 `GETSELSTART` 答的是**量程下限**（实测 10）
   而 `GETSELEND` 答 0（终点比起点小 ⇒ 拿减法会得负数，所以长度折成 0）；`ClearSel` 之后两端都答 `-1`。
   所以 `SetSelLength` 在"没有选区"那一态**锚量程下限**（与控件自己那一态一致），而不是锚折叠出来的 0 ——
   否则 `ClearSel()` 之后紧跟一句 `SelLength = 5` 会被量程夹成 `(10,10)`，读回来还是 0：
   一条"编得过、跑了、什么都没发生"的静默 no-op，本项目最忌这一形。夹具里 `SI4b` 专钉这条。
2. **没挂 `SelectRange` 就是写不进去**（实测整条 `TBM_SETSEL` 不生效），本项目**不伪造**：`SI5`/`SI6` 钉的是
   "写 7 读回 0"，而不是假装存了一份。与 `SelStart`/`SelEnd` 在 SL-b 里的口径同一档。
3. **`ClearSel` 是方法不是属性**，发码要走控件方法那一支改道；落回兜底就编成
   `vb6_ComCall(裸 HWND, L"ClearSel", …)` —— BASE 实测就是这样：那句"清空"跑了、区段一点没动、零诊断。
   目前接的是**调用形** `sld1.ClearSel()`；不带括号那一形（VB6 也允许写 `Slider1.ClearSel`）仍然会被
   当成属性读掉，那一整片记在账 #143。

## 本项目的实现口径（ai/029 C29-SL-k：`Text` 气泡与 `TextPosition`）

类型库里那两条是 declaration-side 的真值：**`Text` = dispid `0x0010`、`VT_BSTR`**，文档原话
"the string displayed in the ToolTip as the slider's position changes"；**`TextPosition` = `0x0011`**，
枚举 `TextPositionConstants { sldAboveLeft = 0, sldBelowRight = 1 }`。所以 `Slider1.Text` **不是窗口标题**，
是拖动时那颗气泡里显示的串 —— 这一格把它从通用的 `vb6_GetControlText`（`GetWindowText`）那支抢过来自己实现。
行为面（气泡怎么摆、什么时候摆）探针量，见下表下面那四条。

| 写法 | 读数与实现 |
| --- | --- |
| `Slider1.Text` | 存一份 BSTR 拷贝在窗口属性里，**同时把 tooltip 宿主里那条工具的文本一起换掉** ⇒ 读回来问的是宿主而不是我们的口袋。没写过就读回空串（VB6 同）；但气泡在 `Text` 空着时打的是**当前的值**（原生那颗本来就画数字） |
| `Slider1.TextPosition` | `0 = sldAboveLeft`（默认，摆在滑块上侧）、`1 = sldBelowRight`（下侧）。**原生没有这条消息**（气泡摆哪儿是控件自己定的）⇒ 自存 0/1、摆的时候用它定偏移；越界读落 `0` 那一档，与 `TickStyle` 同口径 |
| `Slider1.BubbleVisible` | **本项目的扩展读数，不是 VB6 属性**：`IsWindowVisible` 问气泡宿主此刻摆没摆出来（实测无头也答 `True`） |
| `Slider1.BubbleTop` | **扩展读数**：宿主窗口的上边。**只用来比两档的高低**，不钉绝对像素 |
| `Slider1.BubbleText` | **扩展读数**：`TTM_GETTEXTW` 问宿主里此刻挂着的那句文本 |

摆的时机与 `Change`/`Scroll` 是同一条通道（父窗那两条 `WM_HSCROLL`/`WM_VSCROLL`）：`TB_THUMBPOSITION(4)`
与 `TB_THUMBTRACK(5)` 摆、`TB_ENDTRACK(8)` 收，其余档（点轨道、方向键）不动气泡 —— 与 `Scroll` 事件同一分法。
派发那一行**只发一次、不按控件展开**，RTL 里按窗口类名把 ScrollBar 筛掉；容器（Frame / PictureBox）里的
子滑杆这一路还没接（连它的 `Change`/`Scroll` 派发一起记在账 #83 那一族）。

1. **原生轨道条自己那枚 `TBS_TOOLTIPS` 服务不了自定义串**：它自带一条工具、文本是它自己画的数字，
   `TTM_POP` 在无头里什么都不触发（可见性 0、一条 notify 都不发）。而且 `TBS_TOOLTIPS` **只有创建时给
   才建得出那枚气泡**（事后写样式位留着，但 `TBM_GETTOOLTIPS` 恒 0 —— 与 DTPicker 的 `DTS_SHOWNONE` 同族），
   而运行期 RTL 只拿得到 HWND，补不回来 ⇒ 气泡改由 RTL 自持一枚 `TTS_ALWAYSTIP` 宿主 +
   `TTF_TRACK|TTF_ABSOLUTE` 工具来摆（v6 下 `TTM_ADDTOOLW` 回 1、`TTM_GETTOOLCOUNT` 跟着涨、
   `TTM_TRACKACTIVATE(TRUE)` 之后 `IsWindowVisible` 就答 1）。
2. **`TTM_GETTEXT` 的返回值不是成功标志**：实测 `rc=0` 而文本照样复制进缓冲 ⇒ 证人只看缓冲。
   与上一格 `ToolTipRegistered`（账 #148）是同一处坑、同一个修法。
3. **这一格的探针是反着骗人的**：同一份源码编两份，不挂 Common-Controls 6.0 清单那份（v5）
   `TTM_ADDTOOLW` **直接返回 0**，挂上才回 1。上一格差点据此报出一条假缺陷，所以本格的判据全部
   写在产物里（`SK1`…`SK7`），探针只用来定机制。
4. **`Text` 空着时气泡打数字、写了就用写的串**这一条是**本项目的口径**：VB6 文档只说"显示这个串"，
   没说两者怎么共存，那颗 32 位 OCX 在本机跑不起来拿不到真值。哪天拿到 VB6 真值要翻，红的是 `SK6`。

## 本项目的实现口径（ai/029 C29-SL-l：控件的**零实参方法**两形同归）

VB6 里 `Slider1.ClearSel` 与 `Text1.SetFocus` 都允许**不带括号**写。本项目把控件的零实参方法收成
一张表（`controlZeroArgMethod`），两条发码码头共用：带括号那一形在表达式路收尾，不带括号那一形在
语句路收尾 —— 只接一头就会留下"编得过、跑了、什么都没发生"。

| 写法 | 发码 |
| --- | --- |
| `Slider1.ClearSel` / `.ClearSel()` | 两形都是 `vb6_Slider_ClearSel((void*)vb6_hwnd_Slider1);` |
| `Text1.SetFocus` / `Text1.SetFocus()` | 两形都是 `vb6_SetControlFocus((void*)vb6_hwnd_Text1);` |

`vb6_SetControlFocus` 在 RTL 里就一句 `SetFocus(hwnd)`（带空守卫）。焦点属于**线程输入队列**，与窗口
可见/激活无关，所以无头跑里也真能拿到焦点（`SimStdEvent kind=4` 走同一条路，实测会发 `WM_SETFOCUS`）。
拿不到焦点的那几种（控件被禁用、窗口不属于本线程）原生就是回 `NULL` 什么都不做 —— 本项目**不伪造、
不重试**：真 VB6 在那里 raise 一个错误号，而我们还没有运行期错误面。

**哪些控件接、哪些刻意不接**：`Label / Image / Shape / Line / Timer / Menu / Data / OLE /
CommonDialog / Winsock / ImageList / StatusBar / ProgressBar / Form` 都不接（`Form` 有自己的
`SetFocus` 那条路）。控件数组那一形（`txtSearch(Index).SetFocus`）也接 —— 认得
`vb6_CtrlArr_GetAt(&vb6_hwnd_txtSearch, Index)` 这种宿主表达式。

**一条订正**（本格实测翻出来的，别再照旧账写）：`SetFocus` 这一形**以前在运行期也是响的**，
响在 RTL 的宿主对象应答表上（Fix 112，`uc_hostmodel_call.inc` 里有一条按名字应答的
`SetFocus(obj)`）—— 发码是 `vb6_ComCall(裸 HWND, L"SetFocus", NULL, 0)`，而那条调用在进
`IDispatch` 之前会先问"这是不是我们的真窗口"，是就按 Win32 语义应答。所以改道换来的是**不再拿
HWND 当假 IDispatch 去问那张表**（也省一次 `VARIANT` 分配/释放），读数逐字没变。
真正一声不响的是**表里没有的那些名字**，`ClearSel` 就是一个：改之前那句"清空选区"跑了、区段一点没动、
零诊断（夹具 `SN3` 前后读数 `10/20` → `0/0`）。
⇒ 判据里 `SN1`/`SN2` 因此只当**行为钉**（红不了，但会把"焦点面被改坏"拦下来），能红的是 `SN3`
与发码那一正一反（存量工程 `VBFlexGridDemo` 里 31 处 `vb6_ComCall(…, L"SetFocus")` 在新编译器上
只剩 1 处 —— 剩下的那处是 `Me.SetFocus`，走的是窗体自己那条路，不归这张表）。

**`With` 块里的控件方法（ai/029 C29-SL-l / SL-m 两格补齐）**

`With Slider1 … End With` 里既能写属性也能写方法，两形（带括号与不带括号）都接好了：

| 写法 | 结果 |
| --- | --- |
| `With Slider1: .Min = 8: .Max = 72: .Value = 12: .TickFrequency = 8` | 全部落到 `vb6_Slider_Set*(_vb6_with_N, …)`，读回就是那些值（说明 §4 那个例子实测通：`W1=12/8/72`、`W2=1/8/8`、`W3=2/9`） |
| `With Slider1: .SetFocus` / `.ClearSel()` | `vb6_SetControlFocus((void*)_vb6_with_N);` / `vb6_Slider_ClearSel((void*)_vb6_with_N);` —— **这一形以前是编译不过**（发成 `_vb6_with_0.SetFocus()`，`HWND` 是 struct 指针），不是运行期没反应 |

每一层 `With` 有自己的临时量（`_vb6_with_0`、`_vb6_with_1`…），所以嵌套是栈式的：
内层里的 `.X` 绑的是**最内层**那一帧（与 VB6 同），要动外层那个控件就写它的**全名**。
`With` 块里的**带实参**控件方法（例如判据助手 `.SimNotify(5, 40)`）还没接，写全名 `Slider1.SimNotify(5, 40)` 即可。

## 本项目的实现口径（ai/029 C29-SL-n：控件级 `_DblClick` 这一形整条修好）

本页上文 §3 写的"常规：Click、DblClick、KeyDown/KeyUp"里，`DblClick` 这一形以前是**两种坏**：

1. **只挂 `Slider1_DblClick` 的滑杆从来没被子类化** ——"要不要子类化"这一判断被抄成了三份
   （发消息臂的一份、装的一份、拆的两份都漏了 `_DblClick`），于是子类过程与
   `WM_LBUTTONDBLCLK` 那条臂在产物里生成了，却一次也没被 install ⇒ 处理器编得过、永不触发。
   现在三份共用一处判据（`controlNeedsSubclass`），装/拆/arm 不可能再各说一遍。
2. **按 VB6 标准签名写就编译不过** —— 消息臂以前写死成无参调用：
   `extern void vb6_Slider1_DblClick(); vb6_Slider1_DblClick();`
   而 VB6 的签名是 `Sub Slider1_DblClick(Cancel As Integer)` ⇒ `error C2198: 用于调用的参数太多`。
   现在**按处理器自己声明的形参数**发：声明了 `Cancel` 就发
   `int16_t vb6_dblcancel = 0; …DblClick(&vb6_dblcancel);`，写成无参的存量代码仍发 `…DblClick()`。

| 写法 | 触发 |
| --- | --- |
| `Slider1_DblClick(Cancel As Integer)` | 控件自己的 `WM_LBUTTONDBLCLK`（经子类过程，与 `Click` 两档、互不牵连）。`Cancel` 进来是 `0`；原生轨道条没有"默认动作"可取消，所以我们只把这枚局部量交给你写、**不回读** |
| `Slider1_MouseDown/MouseUp/MouseMove(Button, Shift, X, Y)`、`Slider1_KeyPress(KeyAscii)` | 一直是通的（这三条形以前就在"装不装"那两份表里），本批把它们与 `_DblClick` 归到同一处判据 |

`Slider1_Paint` **仍然不接**：本项目的 `_Paint` 只给有绘制表面的控件（PictureBox）——
原生轨道条自己画自己，我们不去伪造一次绘制回调。（真 VB6 里这颗 OCX 到底发不发 Paint，本机拿不到真值：
那颗 OCX 只有 32 位、在 64 位进程里跑不起来。哪天要接，红的是"发了两次绘制"这种形状，先想清楚再动。）

## 本项目的实现口径（ai/029 C29-SL-o：挂在 Frame / PictureBox 里的滑杆，设计期值面也下发）

本项目有**两条创建路**：窗体上的顶层控件走一条，容器（`Frame` / `PictureBox` / `SSTab`）里的子控件走另一条。
以前设计期的那批值面属性只有第一条会下发到窗口，于是同一个 `.frm` 写法放在容器里就是另一种行为
（账 #83 记下过同样的形状：`Text`/`Caption`、`Enabled`/`Visible`、`ToolTipText`/`Tag` 都各修过一头）：

```vb
Begin VB.Frame Frame1            ' .frm 里的写法（容器子控件那一形）
   Begin MSComctlLib.Slider Slider1
      Min             =   10
      Max             =   100
      Value           =   42
      SmallChange     =   2
      LargeChange     =   8
      TickFrequency   =   5
```

| 位置 | 修之前读回 `Min/Max/Value/Small/Large/TickFreq` | 现在读回 |
| --- | --- | --- |
| 窗体上（顶层那条形） | `10/100/42/2/8/5` | 不变 |
| `Frame` 里（容器那条形） | `0/100/0/1/20/0`（一条属性都没下发，读回的是控件自己那一档） | `10/100/42/2/8/5`（与上一行逐字相同） |

两条路现在共用同一处出口（`emitSliderDesignTimeInit`）：参数序、`-999` 哨兵（`.frm` 没写过的那一条不发，
保持控件默认）、以及 `SelStart` + `SelLength` 折成终点那一条，都只有一份。**只修一头是本线踩过三次的形状**
—— 补完仍然只在顶层通，读回来却看不出错。

"没写过"那一档是什么，得在**产物**里问，不能拿裸探针的数当口径：没收到量程消息的轨道条在本项目产物里
答 `0..100`（夹具里那枚什么都没写的 `sld5` 读回 `Min=0 / Max=100 / SmallChange=1 / LargeChange=20`），
而同一份裸探针代码（不带 Common-Controls 6.0 清单）答上限 10 —— 这一族差异已经骗过本线两次（C29-5a/V6、SL-j）。

`LargeChange`（页长）另有一条控件自己的口径：**`.frm` 写了就以写的为准**；
只写 `Min`/`Max` 而没写 `LargeChange` 时，轨道条会按新量程自己重算页长
（探针实测：量程从 `0..10` 改成 `10..100`，默认的 `20` 被重算成 `18` —— 这也是 Init 里
必须先立量程、再立页长、最后才放滑块位置的原因，顺序反了读回来的就是 `18`）。
我们没有另存一份值，所以读回来的就是窗口真在用的那一个。

## 本项目的实现口径（ai/029 C29-SL-p：`FontSize` 存什么读什么 —— 本页 §4 那条例子量出来的）

§4 的例子（拖滑杆改字号）整条是通的，但它撞出一条**通用控件属性**的坏（不止滑杆，所有带字体的控件都是同一支代码）：
`Text1.FontSize = 8` 之后读回来是 `8.25`，写 `10` 读回 `9.75`，写 `14` 读回 `14.25`。
原因是字体在窗口上只能按**整数像素**存在（96 DPI 下 1pt = 1.3333px），而读侧是从那个像素高度
**反算**点号 —— 往返必然落回格点，请求的那个小数从来没被存下来。

| 写法 | 修之前读回 | 现在读回 |
| --- | --- | --- |
| `Text1.FontSize = 8` | `8.25` | `8` |
| `Text1.FontSize = 10` | `9.75` | `10` |
| `Text1.FontSize = 14` | `14.25` | `14` |
| 从没设过字体的控件 | 按窗口字体反算（本机 `8.25` = 系统默认 11px） | **一字不动**（没有自存就读窗口） |

现在请求的点号按窗口自存，读的时候有自存就答自存。这是"窗口表示不了的属性要自存"那一族的既有口径
（`TickFrequency`、`TextPosition` 同一族）。写 `0` 也照答 `0` —— 这一档要靠一枚 Set 旗标撑住：
`SetPropW(hwnd, 名, 0)` 在 Win32 里等于删属性，而 `0.0f` 的位正好是 0。

**渲染仍然是整数像素**：自存只改"答出去的数"，不改窗口真用的字体 ——
`Text1.FontSize = 8` 之后画出来还是 11px 那枚字体（与真 VB6 一致：它答你设进去的那个数，画的是取整后的）。
判据为此带了一枚扩展读数 `FontPixelHeight`（**不是 VB6 属性**，问的是窗口现在真在用的字体像素高度）：
只读自存的那个数问不出窗口有没有真换字体，所以两条一起钉。

## 本项目的实现口径（ai/029 C29-SL-q：设计期字体下发到窗口，`FontName` 读得回来了）

还是 §4 那条例子往下追出来的两件事：

1. **`.frm` 里写的 `FontName` / `FontSize` 现在会打到窗口**（两条创建路都发：窗体上的控件与
   `Frame` / `PictureBox` 里的子控件各一份）。以前这条整条丢掉 —— 设计期在 IDE 里调好的字号与字体，
   换到本项目就是系统默认那一档。与 `Enabled`/`Visible`/`Value`、`ToolTipText`/`Tag` 同一族缺陷：
   **缺口只在创建那一趟，运行期赋值一直是好的**。只在 `.frm` 真写过时发，没写过的控件一字不动。
   `CommonDialog` 除外 —— 它的 `FontName`/`FontSize` 是"要弹出来的对话框用什么字体"，不是控件外观。
2. **`控件.FontName` 现在读得出字体名**。以前它是个"两种答案"的属性：`Len(Text1.FontName)` 答得对
   （8，就是 `Consolas` 的长度），而 `Print ... & Text1.FontName` 打出来是**空串** ——
   读侧没登记成字符串档，打印那一面先装箱再转换就丢了。看着像"字体没设上"，其实窗口早就换好了。

| 写法 | 结果 |
| --- | --- |
| `.frm` 里 `FontSize = 20`（窗体上的文本框） | 运行时 `Text1.FontSize` 答 `20`，窗口真用 20pt（像素高度是它的 ~1.33 倍） |
| `.frm` 里 `FontName = "Arial"` | 运行时 `Text1.FontName` 答 `Arial` |
| `.frm` 里字体写在 `Frame` 内的控件上 | 同上（第二条创建路同一处出口，实测各钉一条） |
| `.frm` 没写字体 | 维持系统默认，读回的是窗口真在用的那一档（不伪造） |

