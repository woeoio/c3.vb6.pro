# TeeChart 控件（图表）——内置实现设计与验收

> 目标：`TeeChart5.ocx` **不再需要**。`tests/TeeChartChartDisplay` 这类工程在无 OCX 的 CI 上
> 真跑，并且能出 **x64** 产物（OCX 只有 32 位，宿主路线下 x64 永远链不上——见
> `tests/run_tests.ps1` 里那条用例被写成 `Arch = "x86"` 的注释）。
>
> 状态：设计/台账（尚未落代码）。下文所有读数都是本仓实测，不是估计。

## 0. 为什么不能照 MSHFlexGrid 那一刀做

MSHFlexGrid / TabStrip 那种内置控件的形状是「一层属性面 + 几个带参方法」：
`FrmControlType` 加一格、`controlTypeToWin32Class` 给个窗口类、`vb6forms_prop_ctrl.h`
列一张 名字→RTL 函数 的表，发码侧逐个点名叫函数。

TeeChart 不是这个形状。实测厂商示例（`...\TeeChart Pro v5 ActiveX Control\Examples\Visual Basic\Visual Basic 5 & 6`，
40 个 `.frm`）的调用面：

| 顶层成员 | 次数 | 顶层成员 | 次数 |
|---|---|---|---|
| `Series` | 143 | `Legend` | 11 |
| `Panel` | 56 | `Page` | 11 |
| `Axis` | 53 | `TimerEnabled` | 7 |
| `Aspect` | 36 | `AddSeries` | 5 |
| `Canvas` | 18 | `Zoom` | 5 |
| `Walls` | 14 | `Printer`/`Export`/`Import` | 4/4/3 |
| `ShowEditor` | 13 | `SeriesCount` | 6 |
| `Repaint` | 12 | `ClearChart`/`RemoveAllSeries` | 各 1 |

二层路径已经是指针链，例如（括号里是出现次数）：

```
.Series(i).FillSampleValues (26)      .Series(i).AsPie.ExplodedSlice.Value(i) (9)
.Panel.Gradient (9)                   .Series(i).AsSurface.UsePalette (6)
.Axis.Bottom.CalcPosPoint(i) (5)      .Series(i).AsBar.BarStyle (4)
.Axis.Left.SetMinMax (3)              .Page.MaxPointsPerPage (3)
.Legend.Font.Bold (2)                 .Walls.Back.Brush.Color (2)
```

⇒ 需要的是**一套对象模型 + 一个绘图引擎**，不是一张属性表。

## 1. 硬约束：38/40 的示例靠 `.frx` 里的 Delphi `TPF0` 流

`grep OleObjectBlob` 实测：40 个 `.frm` 里 **38 个**带 `OleObjectBlob`，设计期的 Series /
标题 / 轴 / 面板样式全在 `TPF0` 流里（`src/project/frx_reader.cpp:450` 已经把负载切出来了，
只是交给了 `IPersistStreamInit::Load` 给真 OCX）。

不解析 TPF0 的后果要说清：`Series(0)` 在运行期并不存在，`Form_Load` 里第一句
`Series(0).FillSampleValues 10` 就落空。P1 用「`Series(i)` 访问即建号」绕过（见 §4），
P2 再把 TPF0 接上；两者不冲突，P2 只是给同一条链补一个更早的数据来源。

## 2. 集成点（照 C30-6 那 23 处登记的清单）

| 层 | 文件 | 要加的东西 |
|---|---|---|
| 类型 | `src/project/frm_parser.hpp` | `enum class FrmControlType` 加 `TeeChart` |
| 解析 | `src/project/frm_parser_util.cpp` | `parseControlType`：`teechart` / `tchart` / `teeChart5.ocx` → `TeeChart`；`controlTypeToWin32Class` → `"VB6_TCHART"` |
| 建窗 | `src/backend/detail/module/cgen_form_create_controls.inc` | `hasTChart` 标志 + `vb6_RegisterTChartClass(hInstance)`（与 `MshFlexGrid` 同一条：自注册类必须先注册，否则 `CreateWindowExW` 拿 NULL） |
| 运行时 | `src/rtl/core/vb6forms/vb6forms_teechart.c`（新） | 窗口类 + 绘制 + 对象模型 |
| 属性面 | `src/rtl/core/vb6forms/vb6forms_prop_ctrl.h` | 只登记**标量**面（`Visible/Enabled/TimerEnabled/...` 那几档），深成员走 §3 的通用口 |
| 六处登记 | `CMakeLists.txt`、`src/driver/c3rtl.rc`（RCDATA 编号接 245 之后）、`rtl_embedded.hpp`、`rtl_embedded.cpp`、`driver_link.cpp` | 少一处 = 解包/内嵌都对、cl 不编它 |
| 判据 | `tests/run_tests.ps1` | 该用例撤 `Arch="x86"`、撤 `RequiresFiles`、撤 `FailIfAssetPresentAndLoadFails` |
| 门 | `scripts/check_rtl_proto_arity.ps1` | controlExit 行数随新增出口改（C30-2 那次 24→25 同型） |

## 3. 发码侧的唯一新口子：通用成员入口

逐成员发函数在 TeeChart 上不可行（成员太多且是链）。所以给 `FrmControlType::TeeChart`
一档**通用**发法，一处实现、后面所有成员都走它：

```c
// 取成员（可能是对象，也可能是标量）
VARIANT* vb6_TChart_GetMember(void* hwnd, const wchar_t* name);
// 设成员
void     vb6_TChart_SetMember(void* hwnd, const wchar_t* name, VARIANT* v);
// 调方法（无/带实参统一走 VARIANT 数组，与 vb6_ComCall 的实参打包同一口径）
void*    vb6_TChart_Call(void* hwnd, const wchar_t* name, void* args, int32_t argc);
```

`Series(i)` / `Panel` / `Axis.Bottom` 这些返回的都是**真 IDispatch**（复用
`vb6forms_memberobj.c` 那套 vtbl），后续整条链（`.AsPie.ExplodedSlice.Value(i) = 0`）
走已有的晚绑定通道，**不再需要任何 cgen 特例**——这正是 `Nodes`/`Panels`/`Recordset`
当年选的那条 cheapest route。

### 3.1 实测：cgen 侧其实已经通了（2026-11-16）

上面那三个函数的“通用发法”不用新写。在夹具里临时写 `TChart1.AddSeries 5` /
`TChart1.Series(0).FillSampleValues 4` / `TChart1.SeriesCount`，用 `--emit-c` 看产物：

```c
vb6_ComCallObject(vb6_ComGetObjectProp(vb6_hwnd_TChart1, L"AddSeries"), NULL,
                  (void*[]){vb6_ComPackInt(5)}, 1);
vb6_ComVarFree((void*)vb6_ComCall(
    vb6_ComCallObject(vb6_ComGetObjectProp(vb6_hwnd_TChart1, L"Series"), NULL,
                      (void*[]){vb6_ComPackInt(0)}, 1),
    L"FillSampleValues", (void*[]){vb6_ComPackInt(4)}, 1));
vb6_CStr(vb6_VariantFromValue(vb6_ComGetStringProp(vb6_hwnd_TChart1, L"SeriesCount")))
```

即：取对象、默认成员 `Item(i)`、方法调用**全部已走通用晚绑定**，链多深都不需要特例。
缺的只有一处——这些调用的接收者是**控件 HWND**（`vb6_hwnd_TChart1`），而
`vb6_ComGetProp` / `vb6_ComCall` 只认 IDispatch。所以第二刀不是加发码，而是在 COM 层
加一枚解析钩子：

```c
// vb6com_wrap.c / vb6com_invoke.c 各入口开头
if (!isComObject(p)) { void* d = vb6_CtlDispatchOf(p); if (d) p = d; }
```

`vb6_CtlDispatchOf` 由 `vb6forms_teechart.c` 实现（窗口属性上挂 chart 根 IDispatch），
其余控件返回 NULL ⇒ 既有行为逐字不变。口径与 `vb6_DI_*` 的弱符号习惯一致。

## 4. 分刀

### P1 引擎骨架 + 无 blob 饼图（第一刀）

**状态：骨架已落并双档跑通（commit `67b89e1f`），成员路由未接。**

已实测（本机，夹具 `tests/ctrlteechart/`，无 `OleObjectBlob`）：

| 判据 | 读数 |
|---|---|
| x64 编译 `TcApp.vbp` | rc=0，`TcApp.exe` 621,056 B（OCX 路线下这一档链不上） |
| x64 运行 | exit=0，stdout `TC-BUILTIN-START` / `TC-WIDTH=3600`，stderr `[TCHART] paint=1` `paint=2` |
| x86 编译 + 运行 | 同上 |
| `check_rtl_proto_arity` | PASS，1732 对 |
| `check_manifest_coverage` | PASS，`inputs=728 [.vbp=262 .bas=466]` |

`TC-WIDTH=3600` 证明设计期几何从原生属性面读得回；`paint=1/2` 是 `WM_PAINT` 里
`C3_TCHART_TRACE=1` 打的点（窗口没建出来或没画，这两行就没有）。

- 新窗口类 `VB6_TCHART`：`Panel` 底色/边距、`Header`（居中）、`Legend` 边框、`Axis`
  上下左右与刻度、`Series` 绘制（Line / Bar / Pie / Point / Area）。
- 对象模型：`TChart` 根 → `Panel` / `Header` / `Footer` / `Legend` / `Axis` / `Series`
  集合；`Series` 有 `Add` / `AddXY` / `FillSampleValues` / `Clear` / `Color` / `Title` /
  `SeriesType`。`Series(i)` **访问即建号**（P2 之前顶住 TPF0 缺位）。
- 验收：`Exploding Pie`、`Multiple Pies` 在 **x64 与 x86** 各跑一次，饼图起得来，
  `AsPie.RotationAngle` / `AsPie.ExplodedSlice.Value(i)` / `OtherSlice.Style` 能改能读。

### P2 TPF0 反序列化（38 个示例的设计期外观）

**状态：已落（`1cc913da` 解析器 + `abc512f6` 发码接线 + `4b263220` 夹具），数据点集待补。**

- `src/project/tpf0_reader.{hpp,cpp}`：组件（ShortString 类名 + 实例名 + 属性表 +
  子组件表）+ 全档位值（Int8/16/32/64、Extended/Single/Double/Currency/Date、
  String/Ident/WString、Set/List/Collection/Binary）
- 发码接线：`emitControlFrxProps` 里按 `.frx` 的 `OleObjectBlob` 发
  AddSeries / SetSeriesColor / SetSeriesTitle / SetHeaderText
- 实测（`tests/ctrlteechart` 带真实 706 B TPF0 负载，x64/x86 逐行相同）：
  `TC-HEADER0LEN=21`（“TeeChart display demo”）/ `TC-DESIGNSER=2` /
  `TC-SER0TYPE=1`（Bar）/ `TC-SER0COLOR=255`（clRed）/ `TC-SER1COLOR=32768`（clGreen）；
  首次绘制 `[TCHART] paint=1 series=3`
- 验收：`tests/TeeChartChartDisplay` 的**设计期** Series/标题是 TPF0 给的，
  而不是运行期补的（已满足）；点集（YValues 的 vaBinary）尚未解。

**覆盖面实测**（`D:\...\TeeV5Demo\*.frx`，186 个）：162 个是 TPF0 负载且**全部读通**，
另 24 个不以 LB 开头（其它资源块）。修前是 154/161，卡在两处写法差异（`c99f63e2`）：
1. 集合两种写法：Delphi 6+ 先写 Int32 count；TeeChart v5 这份流里 `CustomAxes`
   直接是一串 vaList 项（无 count），以 vaNull 收尾
2. `0x0E` 在样本里当集合用（枚举本应是 vaCurrency）—— 先试集合，试不通再回退 Currency
另加尾部截断容错（`DrawlineByCode.frx` 末尾只剩 3 字节残尾）。

### P3 对齐已录制的真 VB6 读数 + 收口

- `tests/guisample/teechart_display_px.ps1` 头部那组数（x86、96dpi 虚拟化、
  `BitBlt` 抓 TChart 子窗自己的 DC）：

  | 判据 | 真 VB6 读数 |
  |---|---|
  | 标题文字 bbox 中心 | `centerX=155.5`（`(91,14)-(220,24)`） |
  | 图例上边 | `y=26, x=310..371, n=62` |
  | 图例左右竖边 | `x=310 y=26..55 n=30` / `x=371 y=26..58 n=33` |
  | 顶轴线 / 底轴线 | `y=53..54` / `y=220..221` |

  修前读数（`Panel.MarginRight/Bottom` 丢事件时）：`centerX=161.5`、图例 `x=322..383`、
  底轴 `y=225`。这三个差异就是 P3 的判红面。
- `ShowEditor` / `Printer` / `Export` / `Zoom` 先落**显式空实现并登记缺口**（返回 0 / 不动作，
  不静默 438），文档 §缺口 逐条列。

## 5. 缺口（P1 交付时必须写进本文档，不许口口相传）

- **成员路由（P1 第二刀，已落 `a831ca57`）**：`vb6_TChart_HostGetProp/HostCall/HostSetProp`
  三条宿主尾叉（认不出返 0 ⇒ 落回原 438 分支）+ 自带最小 IDispatch（集合/系列/根方法）。
  判据已过：`[TCHART] paint=… series=1`、`TC-SERIES=1`。
- **对象模型三层（P1 第三刀，已落 `579a7851`）**：Panel（Color/四边距/Visible/BorderVisible）、
  Header（Text/Caption/Visible）、Axes（Item(i)/Count/Bottom/Left/Right/Top）、
  Axis（Automatic/Minimum/Maximum/Visible/Title）、Title（Caption/Text/Visible）。
  读数：`TC-PANEL=15790335` / `TC-HEADERLEN=4` / `TC-AXTITLELEN=2` / `TC-MARGINR=12`。
- **仍未接**：`Walls` / 3D（`Aspect.View3D` / `Perspective`）—— P1 只记录不绘制。
- **Marks / Pen / Brush 只到值面**：`Series(i).Marks.Pen.Brush/AsPie` 的标量都能
  存能读能触发重绘（`ff655f7a`），但渲染侧还没把这四层画出来（逐点标签、独立笔色）。
- `For Each s In TChart1.Series` 已通（`df4e781f`，夹具 `TC-FOREACH=3`）。
- **TPF0 点集未解**：系列的数据（`YValues` 的 vaBinary 负载）还没接，
  设计期只到“系列在、类型/颜色/标题对”——要真点就用 FillSampleValues/Add。
  优先级实测出来很低：186 个厂商示例里 **135 个在 Form_Load 里调 FillSampleValues**，
  设计期带数据的只有 3 个，而那 3 个的 Binary 分别是 `BackImage.Data` /
  `Brush.Image.Data`（图片，不是点集）——所以样本里根本没有设计期点集要解。
- `AsSurface` / `AsCandle` / `AsHorizBar` 等按类型切面的属性：P3 之后逐族补，
  未补的走「认得出名字、值进 bag、不绘制」。
- 3D（`Aspect.View3D` / `Walls` / `Perspective`）：P1 只记录不绘制。
- 剪贴板 / JPEG 导出 / 打印：P3 空实现。
- `Canvas`（`IsScreenHighColor` / `TextOut` 那族）：P3 起给只读标量。
