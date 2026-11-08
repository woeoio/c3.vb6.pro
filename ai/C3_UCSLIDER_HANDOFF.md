# C3 ucSlider_2022 对齐交接

> 本文件只记 `D:\vb_yqt4qPac\ucSlider_2022\Proyecto1.vbp` 这一个工程；ucTreeView 的结论仍归
> `ai/C3_UCTREEVIEW_HANDOFF.md`，两边不要混写。
>
> 裁判：真 VB6 产物 `D:\vb_yqt4qPac\ucSlider_2022\ucSliderDemo.exe` + 参考图
> `ref-ucSliderDemo.png`（参考图只有主界面一张，hover/按下/滚轮等状态一律现场跑真 exe 取证）。

## 1. 工程结构（已盘点）

| 文件 | 角色 | 备注 |
|---|---|---|
| `Proyecto1.vbp` | 工程 | 单窗体，无第三方 OCX |
| `Form1.frm` | 窗体 | 9 个控件：Picture1(0x393939) 内嵌 ucSlider1(5)；ucSlider1(0..4)、ucSlider6、ucSlider4、LabelPlus LP1(Visible=False)、Timer |
| `ucSlider.ctl` | UserControl | 自绘滑块；`BackStyle=0 Transparent`、`ScaleMode=3 Pixel`、`ScaleWidth=318`、`Windowless=-1`；GDI/GDI+ 全手写 |
| `LabelPlus.ctl` | UserControl | 悬浮读数气泡，默认不可见，仅 MouseEnter 时显示 |
| `clsSubClass.cls` | 类 | LaVolpe 式子类化（VirtualAlloc + 机器码 + vtable 探针） |
| `*.frx` | 设计期二进制 | 字体 + `arrimgknob` PNG（3 个控件有真图；`0x0D20/0x0DDA/0x1A72/0x1D5E` 是 cbSize=20 的空记录，无 payload） |

依赖只有 `VB.UserControl / VB.PictureBox / VB.Timer`，没有 32 位 OCX；但 VB6 产物本身是 PE32，
GDI+ `Declare`/`.frx` 取证也统一在 x86 口径下做，**编译一律 `--arch x86`**。

## 2. 首编

- 夹具：`tests/ucSlider_2022/`（厂商目录只读，全部编译/运行都在副本里）
- 命令：`.build\C3.exe tests\ucSlider_2022\Proyecto1.vbp --output-dir <out> --arch x86`
- 生成器注意：会话临时目录成功后自动删除，要看生成码加 `--keep-for-debug`
- 已修前置（编译器/运行期）：
  - `TextWidth/TextHeight` 入参 BSTR（`cgen_util_type.cpp`）
  - `UserControl.DrawWidth`（`vb6rtl_userctl.h` / `vb6rtl_com.c` / `vb6forms_uc_internal.h` / `uc_host.c`）
  - `With ... .pag` 读 UC 的 `ClassInstance`（`cgen_with.cpp` / `cgen_state.inc` / `cgen_expr_with.cpp`）
  - UC 事件调用按 ByVal 语义（`cgen_form_wndproc_subclass.inc`）
  - 控件数组元素调用（`cgen_expr_call.cpp`）
  - GDI+ 两个缺失 stub（`vb6_di_gdiplus_brush_stubs.c`）

## 3. 启动崩溃链（4 条根因，已修）

产物从「能编译」到「能启动」依次踩到：

1. `Property Set Font` 里 `With UserControl.Font: .Name = New_Font.Name` 的右值只发对象本体，
   `vb6_ComPackValue((*New_Font))` 把对象指针写进 `Name` ⇒ `vb6_uc_fontLogFontImpl` 的 `wcslen` 崩
   （0xC000041D）。修：未知成员 COM SetProp 分支补 `comMarkerValueForWrite`（`cgen_assign_prop_write.inc`）。
2. `UserControl.Font = .ReadProperty("Font", Ambient.Font)` 是裸赋值（无 `Set`），`vb6_ComCall` 的堆
   `VARIANT*` 被硬转成字体指针 ⇒ `vb6_UC_SetFont` 后同一处崩。修：`comObjectRefFromCallExpr`
   （`cgen_assign_host_pseudo.inc`）。
3. `LP1.Caption = ucSlider6.Value`：VB6 形参缺省 ByRef ⇒ C 形参 `BSTR*`，调用点却传 BSTR 值，
   被调方 `(*New_Caption)` 把首字符当指针 ⇒ OLEAUT32 `SysStringLen` 崩。修：写方向 String 形参传
   `&(BSTR){...}` 复合字面量（`cgen_assign_prop_write.inc`）。
4. `Form_Load` → `clsSubClass.ssc_Subclass` → `zAddressOf` 拿 `ObjPtr(oCallback)` 当 vtable 指针，
   在 `+0x1C/+0x6F8/+0x710/+0x7A4` 扫 8 个 DWORD 当代码指针读首字节；C3 的 `g_ComObjectVtable` 只有
   7 项（0x1C 字节），越界读到 `.rdata` 相邻值 `0x65725f5f` ⇒ memmove AV。修：vtable 后垫整片 0 槽
   （`vb6comserver_obj.c` 的 `vb6_PaddedComVtable`），让探针读 0 提前返回。

副作用见 §7 已知缺口。

## 4. 运行期取证（差异清单草案）

取证环境：本机 1920x1200 @120dpi（125%），VB6/C3 都按 DPI 虚拟化跑；主窗体 830x469 DIP，
物理截图 1038x586。两侧同一窗口区域逐像素比（通道差和 >24 记差异）。

全窗体 45,957 px 差异，**全部集中在两处**：

- **A｜Picture1 内 ucSlider1(5)**（x≈795-975, y≈105-315）
  - A1 背景：C3 在 UC 范围内刷 `f0f0f0`（窗体底色），VB6 是 Picture1 的 `0x393939`
  - A2 旋钮：C3 是白色圆角占位块，VB6 是 `.frx` 的 PNG 图旋钮（深色带白条纹）
- **B｜ucSlider1(4) 内部** 出现 221x21 物理像素纯黑实心条（x=769..989, y=418..438）；VB6 同区无。

一致项（无需动）：窗体底色、4 个竖向滑块、绿色/橙色横向滑块、刻度线与刻度文字、窗口尺寸。
截图底部 y≈576+ 的差异是桌面壁纸字样，属采集污染，判据里必须排除。

根因判断（待逐条收口后回填证据）：

- A1：UC 宿主创建时 ambient BackColor 只按父**类名**查窗体类色，Picture1 这种容器控件查不到，
  落回窗体底色；应优先取父窗口自身的 BackColor（`vb6_GetControlBackColor`）。
- A2：容器（Picture1）内控件的属性包发射走 `cgen_form_frame_menu.inc` 那一路，缺 `.frx` 二进制
  blob 分支，`arrimgknob` 被当 DocString 发成乱码 ⇒ 图片旋钮加载失败。
- B：透明 UC 的 `WM_PRINTCLIENT`（composeBelow=0）分支既没铺底、又照常新建离屏 DIB，DIB 零初始化
  为纯黑，随后整块 BitBlt 到目标 ⇒ 兄弟控件合成时把黑块盖到下层。

## 5. Guard（已落地）

`tests/guisample/ucslider_px.ps1`（已挂进 `tests/run_tests.ps1` 的 `Test-VbpBuild -PixelAssert`）。
采集侧只抓**客户区** 1020x539（`GetClientRect + ClientToScreen`），并硬性校验 `ch == 539 && ch < 576`，
桌面壁纸那条污染进不来（没有靠放宽阈值绕过）。

判据与双向实测（同一脚本、同一相机）：

| 断言 | 判据 | 真 VB6 参考 exe | 旧 C3 产物（HEAD） |
|---|---|---|---|
| A1 Picture1 内 `#393939` 占比 | `>= 0.80` | 0.881（绿） | 0.349（红） |
| A2 PNG 旋钮蓝灰体像素数 | `>= 400` | 1135（绿） | 0（红） |
| B ucSlider1(4) 纯黑像素数 | `== 0` | 0（绿） | 4641（红） |

A2 抓的是 `(11,12,15)/(31,35,46)/(54,65,85)` 这一族蓝灰（PNG 旋钮实体），
只判「非白」会把旧产物的白色占位圆角块判绿，因此不采用。
基线重测命令（换机器/换 DPI 时先跑这条）：

```powershell
powershell -NoProfile -File tests\guisample\ucslider_px.ps1 `
  -Exe "D:\vb_yqt4qPac\ucSlider_2022\ucSliderDemo.exe" -Baseline
# BASELINE ucslider cw=1020x539 scale=1.25 A1.bg=0.881 A2.knobPngBlue=1135 B.slider4PureBlack=0
```

## 6. 已知缺口

- `clsSubClass` 的 vtable 探针在 C3 产物上恒返回 0（§3.4 护栏副作用）⇒ 子类化不生效、`WM_MOUSEWHEEL`
  不转发。滚轮行为暂不可对齐；不得用伪造/补洞方式让它「看起来能跑」，如需支持应另立账做真实 thunk。
- `LabelPlus` 默认 `Visible=False`，主界面不出现；hover 状态需现场跑真 VB6 取证后再加用例。
- objcache 键未覆盖内嵌 RTL 的 `.inc` 改动（§7 那条坑），改了 `.inc` 的夹具构建必须
  `C3_NO_OBJCACHE=1`，否则读数是旧运行时的。

## 7. 差异收口进度

### A1｜Picture1 内 UC 背景被刷成窗体底色（已收口 2026-11-08）

- 修复点：`uc_host_create.inc` 的 ambient BackColor 取值。原来只 `vb6_Forms_QueryClassBg(父类名)`；
  `VB.PictureBox` 的类名不在类背景表里（那张表只登记窗体）⇒ `bg = -1`，ambient 保留上一次窗体底色。
  现改为：类名查不到时，若父窗口自己 Set 过 BackColor（`VB6_BackColorSet` 哨兵）就取
  `vb6_GetControlBackColor(父窗口)` —— 与 VB6「Ambient.BackColor = 容器底色」语义一致。
  实测 Picture1 生成码确实 `vb6_SetControlBackColor(Picture1, 3750201 /* 0x393939 */)`，哨兵在位。
- 读数：`A1.bg` 0.349 → **0.926**（真 VB6 0.881，判据 `>= 0.80`，绿）。
- 夹具回归命令（**必须带 `C3_NO_OBJCACHE=1`**，见 §6 末条）：
  `$env:C3_NO_OBJCACHE="1"; powershell -File scripts\c3_build.ps1 -Vbp tests\ucSlider_2022\Proyecto1.vbp -OutDir <out> -Extra "--arch x86"`

**坑（本轮踩到）**：改 `src/rtl/**` 的 `.inc` 后，`build.bat` 会正确重跑 `rc.exe`
（`check_rtl_embedded.ps1` R1..R3 全绿，126 条逐字节一致），但夹具构建的 objcache 全命中
（`OBJCACHE rtl=83/83`），exe 里还是旧 `uc_host.o`，读数纹丝不动（0.349）。绕开办法：
`C3_NO_OBJCACHE=1`（或 `--clean-first`）。长期应把 objcache 键扩到内嵌 RTL 全量哈希，另立账。

### A1 收尾：Ambient.BackColor 改成每实例（2026-11-08，随 A1 一起收口）

- 起因：A1 改完后 guard 三条里 A1 绿了，但逐像素复核发现 `ucSlider1(4)`（窗体上，不是 Picture1 里）
  整片 73 DIP 被刷成 `#393939`（参考 VB6 是 f0f0f0）。根因是 `vb6_Ambient_BackColor` 是**裸全局**：
  创建时写一次，后创建的 Picture1 内那颗 UC 把它改成 `#393939` ⇒ 先创建的控件**后续重绘**读到的是
  别人的容器底色。`uc_host_window.c` 的 4 个铺底站点（composeBelow / paint / AutoRedraw ×2 /
  WM_ERASEBKGND）全吃这一条。
- 修复：`vb6_UCRec` 加 `ambientBack`，HostCreate 时定格（类名路径取不到就沿用当时的全局值）；
  `vb6_uc_fillAmbient` 加 `oleAmbient` 形参，4 个站点全部改读 `r->ambientBack`。全局仍照写
  （`Ambient.BackColor` 的生成码读它，且只在创建期读）。
- 复核：新旧截图逐像素比，**只有** rows 71..269 / x 791..931（Picture1 内那枚 UC 的矩形）发生变化，
  其余 0 px 变动 ⇒ 没有副作用面。读数 `A1=0.899 A2=1135 B=4641(黑条仍在，见下)`。

### A2｜旋钮没画 `.frx` PNG（已收口 2026-11-08）

- 定位：`Form1.c` 里 Picture1 内的 `ucSlider1(5)` 发的是
  `vb6_UC_BagPutStr(_bag, L"arrimgknob", L"\uE874\uF9B1…")` —— 把 `.frx` 二进制当 DocString 发成乱码；
  顶层那几颗走的是正确的 `vb6_UC_BagPutBin(_bag, L"arrimgknob", <PNG 字节>)`。
- 根因：容器（Picture1/Frame）内子控件的发射路径 `emitUserControlChild`（`cgen_form_frame_menu.inc`）
  的 bag 块只有 PropertyBag 块 + DocString 两支，**缺 DocProperty 二进制支**；顶层
  `cgen_form_create_controls.inc` 有（`readDocBlobBytes` + `detectFormat`）。
- 修复：在该路径补同一判据的 `FrxBlobBytes` 支 —— `readDocBlobBytes` 命中且 `detectFormat`
  认出 PNG/JPEG/BMP 才按原始字节 `vb6_UC_BagPutBin` 进 bag，其余仍退回 DocString。生成码已核对：
  `vb6_UC_BagPutBin(_bag, L"arrimgknob", "\211PNG…", 604)`。
- 读数：`A2.knobPngBlue` 0 → **1135**（真 VB6 也是 1135，判据 `>= 400`，绿）；`A1.bg` 0.899 仍绿。

### B｜ucSlider1(4) 顶部多一条 221x21 纯黑条（已收口 2026-11-09）

定位（上一轮已报，本轮落地）：

1. 形状：黑条在客户区 `(760..981) x (380..400)`；宽度=控件宽，高度 21 px。上层兄弟宿主
   （`VB6_UserControlHost`，客户区 `750..991 x 320..401`）与 slider4 矩形 `750..981 x 380..471`
   的交集正好是 `221 x 21`。
2. 机制：`vb6_uc_paintInto` 里 `vb6_uc_dibCreate` **无条件**建一张零初始化 DIB（纯黑）；当本控件是
   **透明 UC 且 composeBelow==0**（正被别人 `WM_PRINTCLIENT` 合成）时，两个铺底分支都不进
   （`transparent && composeBelow` 假、`!transparent` 假）⇒ DIB 一个像素都不铺，控件也只画自己
   那点内容；末尾 `BitBlt(target, 0,0, dib.w,dib.h, dib.memDC, ...)` 把整张 DIB **连没画过的黑底**
   盖回 target ⇒ 兄弟合成时这 221x21 黑块被叠进 slider4 的 DIB。

修法（[uc_host_window.c](file:///c:/Users/Administrator/Documents/c3.vb6.pro/src/rtl/core/vb6forms/uc/uc_host_window.c)）：

```c
const int directTarget = (transparent && !composeBelow);
if (!adc && !directTarget) vb6_uc_dibCreate(&dib, target, cw, ch);
```

即这条路径**不建 DIB**，`dc` 直接落到 `target`（target 里已是合成好的下层，调用方也已按本窗矩形
设好 viewport/clip —— 与 VB6 windowless 语义一致）。没有保留「建 DIB 但保证铺底」的分支：
要保证铺底就得把下层内容也复制进来，代价比直接画高，且现有 `useDib==0` 的兜底路径本来就支持
直画 target，行为面最小。AutoRedraw（`adc`）与非透明 UC 两条路径完全未动。

验证：

- guard 四门全绿：`A1.bg=0.899 A2.knobPngBlue=1135 B.slider4PureBlack=0`（rc 0）。
  旧产物（`.temp\ucslider_gen`，HEAD 前产物）仍红：`0.349 / 0 / 4641`（rc 1），双向能力保留。
- **逐像素证据**（`b_out` 截图 vs 上一轮 `a1b_out` 截图，同相机 1020x539，阈值通道差和 >24）：
  变化 4651 px，其中 **4641 px = rows 380..400 / x 760..980，正是那条黑条**（纯黑→f0f0f0）；
  余 10 px 全部落在 `(1017..1019) x (534..538)` 右下角尺寸抓手，**同一 exe 连拍两张也会在这 10 px
  上变**（semi-transparent 抓手透出桌面），属采集噪声，不是画面差异。
- 其它透明 UC 合成未被画坏：ucTreeView 五窗像素门整套重跑 **PASS**（rc 0）——
  `Tree Explorer / Corporate Hierarchy / Drag and Drop / Menu·Tools / Contact  List` 五窗的
  `scr`/`content` 读数、hover、数字微调、XML 重载、`chk240=82.9% / label240=90.6% / sb240=93.8%`
  等判据全绿，与 `ai/C3_UCTREEVIEW_HANDOFF.md` 的基线一致。

## 8. GA 22 job 收口（2026-11-09）

首跑（head `bf0874f8` = A2 收口后）红 4 格：`Emit manifest` / `Tests (compile #1)` /
`Tests (compile #2)` / `Tests (vbp #3)`。逐格归因与修法：

1. **vbp #3｜guard 在 CI 上 CAPTURE-CONTRACT ch=431**。CI runner 是 96 DPI（客户区 `816x431`），
   上一轮把采集闸写死成 `ch == 539`（本机 125% 的读数）—— **判据对、契约错**。改成：
   `ch < 576` 保留（壁纸那条照样进不来），另加 `ch == round(431 * cw/816) ± 1` 认定「这是
   816x431 DIP 那张窗体的客户区」，不再写死任何机器读数。同时把 pixel FAIL 的输出从
   `    pixels FAIL [x]: …` 改成用例行 `[PIXELS] <script> ... FAIL (…)`：`ci.yml` 的命名 matcher
   只认用例行，原来 tally 到 1 个 FAIL、注解却报 “naming found only 0”，自相矛盾。
2. **compile #1/#2｜两份静态 census 的计数没跟上上一轮已收口改动**（判据仍 exact，只改数字+注释）：
   - `check_com_marker_write_sites.ps1`：7 → **8**（第 6 个写侧 = `With class .unknown COM SetProp`
     的写，正是 §3.1 那刀补的 `comMarkerValueForWrite`）。
   - `check_formatdatetime_args.ps1`：22 → **24**（新增 `vb6_UserControl_TextWidth` /
     `TextHeight` 两行，正是 §2「首编前置」那刀把 UC 的 `TextWidth/TextHeight` 形参改成 BSTR）。
3. **Emit manifest｜605 行，改写 6 / 新增 2**。新增 2 = 新夹具
   `tests/ucSlider_2022/{Proyecto1.vbp, modIOleInPlaceActiveObject.bas}`（上一轮加了夹具没登记）。
   改写 6 用「**旧编译器**」逐行核过（取 CI 为最后一个绿 head `b689a88d` 交的 `c3-exe` 工件当 A/B，
   `--emit-c` 对 6 份工程逐行 diff），只有上一轮那三种形状，没有第四种变化：
   `With class .unknown COM SetProp` 由 `vb6_ComPackValue((*New_Font))` 改成读回值
   `vb6_ComGetStringProp/ComGetDoubleProp(...)`；String ByRef 写实参多出 `&(BSTR){}`；
   控件数组默认属性从 `vb6_GetControlText(vb6_CtrlArr_GetAt(...))` 改成直接取元素值。
   登记后本机复算：`期望 605 / 实得 605 / 相同 605 / 哈希不同 0 / 缺席 0 / 多出 0`，
   `check_manifest_coverage` K1..K3 绿。
