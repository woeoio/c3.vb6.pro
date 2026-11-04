# C3 / ucSegmentedControl 断点续跑交接（2026-11-05）

## 0. 结论一句话

**编译已通（x86，PE32）→ 能启动、窗口与 15 个子窗都建出来、GDI+ 自绘已跑 → 效果与真 VB6 参考仍有明显差**：
图标全缺（只剩文字）、文字基线/位置偏移。当前正在补“设计期 `.frx` 二进制属性（`image_N`）进属性包”这一站点，补丁已写、**尚未编译验证**。

## 1. 目标与基准

- 源码：`D:\vb_yqt4qPac\ucSegmentedControl-source`（VB6 用户控件工程，核心 `ucSegmentedControl.ctl` 90KB）
- 参考产物：同目录 `SegmentedControl.exe`（339,968 B，2026/11/5 2:40:27，32 位 PE）
- 参考图：同目录 `SegmentedControl.png` = **949x689**；实测参考 exe 运行窗口 = **766x558**，**不是 1:1** ⇒ 按用户要求，以“真 VB6 exe 运行采集”为基准：
  `.temp\ref_run.png`（766x558，由 `.temp\probe_ref.ps1` 采集）

## 2. 实测读数（已确认）

参考 exe（VB6 真产物）：

```
ALIVE pid=13528 mainwnd=14746614 title=[LabelPlus]
TOPCOUNT=8
TOP 766x558@205,205 client=751x520 org=212,235 class=ThunderRT6FormDC title=[Form1]
TOP 0x0@768,480 class=ThunderRT6Main title=[LabelPlus]          (隐藏的 app 对象窗)
TOP 1x1@0,0 class=GDI+ Hook Window Class / VBMsoStdCompMgr / MSCTFIME UI / IME
MAIN 766x558@205,205 client=751x520 org=212,235 class=ThunderRT6FormDC title=[Form1]
KIDS=2
  VBFocusRT6              0x0@212,235 (vis=False)
  ThunderRT6CommandButton 73x25@812,459 client=72x24 text=[<-- ItemVisible]
```

C3 产物（上一轮采集，`.temp\c3_run.png`，766x558）：

```
MAIN 766x558@128,128 client=751x520 class=VB6_Form_Form1 title=[Form1]
KIDS=15：1×Button(73x25@735,382 text=<-- ItemVisible) / 12×VB6_UserControlHost / 1×Static(177x17@159,438 "CtrlType = 1 - CheckBar")
```

⇒ 窗口尺寸/客户区与参考 **完全一致**（766x558 / 751x520）。
⇒ 子窗口数量不同：VB6 里 UserControl 是 **windowless（无独立 HWND）**，C3 用 `VB6_UserControlHost` 承载（12 个），这属实现差异，**不是像素残差**；像素对照以位图为准。

## 3. 编译真实结果与失败分类（按发生顺序，全部已按站点修）

| # | 现象 | 归类 | 根因 | 修复文件 |
|---|------|------|------|----------|
| 1 | `C2065 'IsEqualGUID': undeclared identifier`（生成的 DI 桩里引用了 SDK 内联名） | cl 编译失败 | ole32 无该导出，guiddef.h 里是内联/宏 | `src/rtl/core/di/vb6_di_stubs.c` 手写 `vb6_di_IsEqualGUID`（memcmp 16B） |
| 2 | 启动即 `Unhandled error 9: Subscript out of range`，VEH `rva=0x21d14` → `vb6_UBound` | 运行期错误 | `mIOleInPlaceActivate.c:332`：ParamArray 转发时 `vb6_PA_SetLong` 把 `SAFEARRAY*` 截成 VT_I4；`varParam = arrParam` 的 `_Generic` 落 default 成假 VT_DISPATCH | `vb6rtl_paramarray.c`（`vb6_PA_SetSafeArray`）、`vb6rtl_com.c`（`vb6_VariantFromParamArray`）、`vb6rtl_variant.h`（`_Generic` 两支）、`vb6rtl_runtime.h`、`cgen_expr_call_named_args.inc` |
| 3 | `0xc0000005 av read 0` @ oleaut32 `DispCallFunc` ← `OleInvoke` ← `InitIPAO` | 运行期崩溃 | `.ctl` 把 `Me` 当 `As Object` 实参传给 `InitIPAO/SetIPAO`；C3 进程内 UC 宿主模型里 `me` 是裸实例结构体、`__comObj` 恒 NULL，`ObjPtr(me)` 交回实例地址 ⇒ `DispCallFunc` 把首字段当 vtable | `cgen_expr_call_arg_variant.inc`：设计器模块整枚 `Me` 作实参时 ByRef 交 `&(void*){0}`、ByVal 交 `(void*)0` |
| 4 | 交裸 `0` 后新崩溃 `InitIPAO+0x82 av read 0` | 运行期崩溃 | 被调方 `ObjPtr(Ctrl)` 生成 `vb6_ObjPtr((*Ctrl))`，`Ctrl=0` 解引用空指针 | 同上（改交 `&(void*){0}`） |

**当前编译状态：cl + link 全通过（exit 0）**，产物 `Project1.exe` 662,016 B，PE 实测 `machine=0x014C`（x86）、`optmagic=0x010B`（PE32）、`subsystem=2`（GUI）。

编译告警（不判失败，如实记录）：`VB4001` `stdole2.tlb` 路径缺失；`VB3003` `.ctl(2063)` Byte 隐式转换；`VB3001` `Form1.frm(83/101)` 裸 `Controls`（codegen 已发 `vb6_ComGetObjectProp(vb6_hwnd_Form1, L"Controls")`，语义层仍报未声明）。

## 4. 效果差异（参考 vs C3，肉眼+图像）

| 站点 | 参考 exe | C3 | 状态 |
|------|----------|----|------|
| 窗口/客户区 | 766x558 / 751x520 | 766x558 / 751x520 | ✅ 一致 |
| 圆角/边框/选中态底色 | 有 | 有（GDI+ 已跑） | ✅ 大体一致 |
| 分隔线 | 有 | 有 | ✅ |
| **分段项图标** | 每项左侧有 PNG 图标 | **全缺，只剩文字** | ❌ 正在修 |
| **文字位置/基线** | 图标右侧、垂直居中 | 偏左/偏上（图标位空缺导致） | ❌ 随图标修复再对齐 |
| 顶部条 4 项（CheckBar） | 图标 + New Item | 仅 New Item | ❌ 同上 |
| `CtrlType=1` 小条 | 4 个图标 | 显示 “昵/d/=” 等文字（Segoe UI Symbol 字形） | ⚠️ 字体字形差异，需单列 |

## 5. 根因（图标缺失）与已写补丁

`.frm` 里 `image_N = "Form1.frx":xxxx` 是 **DocProperty 二进制块**：

```
[DWORD cbSize][1B 00][GUID 16B][WORD 00][WORD 01][DWORD nBytes][4B 0][payload]
payload 起点 = offset+32，cbSize == 28 + nBytes
例：0x00BE → cbSize 0x254，nBytes 0x238，payload = PNG（89 50 4E 47 …）
```

原 codegen 的 `FrxReference` 分支只试 `readPropBagBlob`（识别 UserControl 属性包块）→ 失败后退 `readDocString`；PNG 二进制过不了“UTF-16 全部可打印”校验 ⇒ **`image_N` 根本没进属性包** ⇒ `ReadProperty("image_N", 0)` 回默认 0 ⇒ `ItemImage(i)` 空 ⇒ 无图标。

已写（**未编译验证**）：

- `src/project/frx_reader.hpp`：新增 `struct FrxBlobBytes` + `readDocBlobBytes(size_t)`
- `src/project/frx_reader.cpp`：实现（固定偏移 24 取 nBytes，校验 `cbSize == 28+nBytes`、边界）
- `src/backend/detail/module/cgen_form_create_controls.inc`：`FrxReference` 分支在 DocString 之前插一级 `readDocBlobBytes` + `detectFormat != Unknown` ⇒ `emitFrxBagBlob`（走 `vb6_UC_BagPutBin`，RTL 侧变成 `VT_ARRAY|VT_UI1`）

⚠️ **同类站点还有一处未改**：`src/backend/detail/module/cgen_form_frame_menu.inc`（约 L298-310）有同样的 `FrxReference` 三段逻辑，需同步加同一级（否则从另一条创建路径进来的 UC 仍丢图）。

## 6. 已改动文件清单（相对 `github`/`origin` 的 `dev`）

上一轮（已通过编译验证）：

```
src/semantics/symbol_table.cpp
src/backend/decl/cgen_decl_api.cpp
src/backend/detail/util/cgen_state.inc
src/backend/detail/expr/cgen_expr_call_builtin_pre.inc
src/backend/cgen_util_classtype.cpp
src/backend/expr/cgen_expr_binary.cpp
src/backend/detail/expr/cgen_expr_ident_builtin.inc
src/backend/module/cgen_com.cpp
src/backend/detail/stmt/cgen_assign_host_pseudo.inc
src/backend/detail/expr/cgen_expr_call_arg_emit.inc
src/backend/detail/module/cgen_form_prelude.inc + 新增 src/common/frm_code_scan.hpp
src/backend/module/cgen_form.cpp
src/rtl/core/vb6rtl/vb6rtl_userctl.h  /  vb6rtl_com.c（Extender ToolTipText / vb6_Extender_Move）
src/rtl/core/di/vb6_di_stubs.c（IsEqualGUID）
src/rtl/core/vb6rtl/{vb6rtl_paramarray.c,vb6rtl_com.c,vb6rtl_variant.h,vb6rtl_runtime.h}
src/backend/detail/expr/cgen_expr_call_named_args.inc
src/backend/detail/expr/cgen_expr_call_arg_variant.inc
src/rtl/core/di/**（批量重新生成后的桩文件改动，随构建脚本产生）
```

本轮新改（未验证）：

```
src/project/frx_reader.hpp
src/project/frx_reader.cpp
src/backend/detail/module/cgen_form_create_controls.inc
```

## 7. 下一步动作清单（按顺序，单步验证）

1. `build.bat` 重建 C3（x86），确认新增 `readDocBlobBytes` 编过（本机 MSVC `/std:c11`）。
2. 重编夹具工程到 `.temp\ref\` 平行目录，`--arch x86`，PE 断言 `machine=0x014C`。
3. `.temp\probe_ref.ps1` 同参数跑 C3 产物 → 截 `.temp\c3_run2.png`，确认图标出现。
4. 若图标出现但位置/尺寸不对 ⇒ 查 `DrawImage` / `ImageListDraw` / `m_IconsSize*nScale` / DPI（`GetWindowsDPI`）。
5. 同步 `cgen_form_frame_menu.inc` 的同类分支。
6. 逐像素/rect 对照：窗口与客户区、12 个 UC rect、分段按钮边界、文字基线、选中态/悬停配色、分隔线、圆角、边缘像素 ⇒ 记录改前/改后/差值。
7. 建夹具 `tests\ucSegmentedControl\`（复制 `.vbp/.frm/.frx/.ctl/.ctx/.bas/.pag/image.jpg`，**不依赖 D: 盘**），新增断言：PE `machine=0x014C` + 32 位进程 + 窗口 766x558/751x520 + 关键像素（图标区非空、选中态色、分隔线）。
8. 登记 `tests\run_tests.ps1` 用例表、`check_vbp_fixture_census.ps1`、`emit-manifest.expected.txt`（按 CI 实际清单，**禁 `-Bless`**）。
9. 回归 `scripts\baseline.ps1`（3/0/0）、`scripts\test.bat gui|syntax|compile`（8/0/0）；免注册 COM 两路径（`LoadLibrary+DllGetClassObject`、SxS manifest）不得回归。
10. 分功能提交 → 快进推 `github`+`origin` 的 `dev`（不 force）→ 跑 GA 至全绿。

## 7.5 本轮续跑进展（夹具化 + 像素/rect 对照）

已落：夹具 `tests\ucSegmentedControl\`（`.vbp/.vbw` 除外 `.vbw` 被 ignore、`Form1.frm/.frx`、
`Form2.frm/.frx`、`ucSegmentedControl.ctl/.ctx`、`PropPagSC.pag`、`mIOleInPlaceActivate.bas`、
`image.jpg`、`Manifest.RES`），**不依赖 D: 盘**。CI 口径下 `*.res` 默认被忽略，按仓库已有先例
（`tests\BalloonTooltips\BalloonTooltips.res` 等）用 `git add -f` 强制登记厂商清单；它是进程
清单（`#1`），不进 SxS 装配路径。

新增判据 `tests\guisample\segmented_px.ps1`（已单跑 PASS，rc=0）：

```
main class=VB6_Form_Form1 title=[Form1]; window=766x558 client=751x520
hosts visible=12 hidden=1
  0,0 751x49 | 0,48 217x209 | 16,304 185x33 | 216,320 104x41 | 224,368 193x41 | 224,56 261x41
  400,104 272x33 | 464,432 113x33 | 512,144 73x265 | 592,432 145x33 | 72,416 362x41 | 8,464 137x41
  hidden: 216,104 217x209 (ucSegmentedControl5, Visible=False)
buttons: 600,224 73x25 [<-- ItemVisible]
uc2 vivid=145 blue=138; uc3 vivid=1185 blue=625; uc4 vivid=1734 blue=1734
uc11 vivid=970 blue=209; uc1 vivid=802 blue=161
accent 0078D7=10715 CE00EF=400
```

⇒ **12 个可见 UC 宿主的 rect 全部等于 `Form1.frm` 的设计期 twips/15**（逐条核对），也就是
"设计期布局被原样复现"；真 VB6 参考的窗口/客户区读数与之完全相同。

像素对照（客户区 751x520，同处 96dpi 虚拟化空间）：

```
真 VB6 参考 (SegmentedControl.exe) vs C3 (Project1.exe)
MEAN_ABS_DIFF_PER_PIXEL = 3.967 / 765   (改前含图标缺失/白底方块时明显更大)
强调色: 0078D7 参考 10670 / C3 10715   CE00EF 参考 554 / C3 400
```

仍有残差（已肉眼+裁剪逐区核对，**尚未修**，故未写进断言）：

1. 顶部 CheckBar 第 3 项：真 VB6 是 **Disabled 灰显**，C3 仍可点（Enabled 语义/属性包差异）。
2. "Reporto" 项的图标映射不同（参考为柱状图图标，C3 为文档图标）。
3. `CtrlType=1` 小条与底部翻页条上的 PUA 符号字形（Segoe UI Symbol）在两产物上取字不同
   （参考 4 个图标 vs C3 的 "K/d/=" 三个）。
4. 少量文字 pen 位置相差 1–2px（字体度量）。

## 7.6 GA 收口修复（同一轮）

- `[STATIC] builtin_slot_returns` 红在 GA：RTL 新增 `vb6_VariantFromParamArray`，checked-in 的
  `src\backend\detail\expr\cgen_variant_slot_prefixes.inc` 少一名（derived=41 / current=40）。
  修法是重跑派生脚本而不是加白名单：`pwsh scripts\gen_variant_slot_prefixes.ps1` ⇒ 41 名，
  复跑 `check_builtin_slot_returns.ps1` = PASS。
- `Emit manifest` 红的 7 份：d71875a4 的 codegen 改动（`Me`→NULL 收窄、UDT 返回值落临时）确实
  改了发码形状。按仓库口径取 CI 权威清单 `git fetch github ci/emit-manifest`（head=d71875a4）
  的对应行逐行登记，未用 `-Bless`。
- 注意：本机 emit 结果与 CI 在 4 份输入上不同（`tests/TeeChartChartDisplay/DisplayProject.vbp`、
  `tests/guisample/GUISample.vbp`、`tests/guisample/vb6ref/src/{vs2015,vc6}/GUISample.vbp`）——
  这 4 份依赖外部 OCX/typelib 的解析，**登记必须以 CI 那台为准**，不要拿本机哈希覆盖。

## 7.7 GA 仍红：夹具在 CI 上 `FAIL (compile)` —— 两个 C3 能力缺口（未修，接手从这里开始）

证据：GA run **41916340378**（head=fc8b5b9d）产物 `test-logs-vbp-2` 里的
`ucSegmentedControl\c3-error.log`（81927 B）与 `compile.log`：

```
ucSegmentedControl.c(3326): error C2197: 'vb6_SafeArray1D *vb6_Split(BSTR,BSTR,int32_t,int32_t)':
                            too many arguments for call        (×4)
ucSegmentedControl.c(3353): error C2197: 'void *vb6_GetObject(const wchar_t*, const wchar_t*)':
                            too many arguments for call
ucSegmentedControl.c(3365): error C2197: 同上
```

本机 `--emit-c`（`.build\C3.exe`）对同两处发出的形态：

```c
/* .ctl:2166  Base64Decode Split(SrcImg, ",")(1), DataArr */
vb6_ucSegmentedControl_Base64Decode((void*)me,
    vb6_Split(vb6_VariantToString((*SrcImg)), vb6_BSTR_FromStr(L","), -1, 0, 1, NULL, -1, 0), &DataArr);
/* ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 8 个实参, 而 RTL 声明只有 4 个:
   vb6rtl_array.h:116  vb6_Split(BSTR expr, BSTR delimiter, int32_t limit, int32_t compare) */

/* .ctl:2195  If GetObject(SrcImg, Len(tBmp), tBmp) Then  /  2211  Call GetObject(hBmp, ...) */
if (vb6_GetObject(vb6_VariantToLong((*SrcImg)), ((int32_t)sizeof(tBmp)), (void*)&(tBmp))) { ... }
/*   ^^ 3 个实参, 而内置 vb6_GetObject 只有 2 个; 源里是 mIOleInPlaceActivate.bas:104
     Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA"(hObject, nCount, lpObject) */
```

归因（两条，均需改 C3 而不是改夹具）：

- **A. 省略实参补档补错了括号层**：`Split(SrcImg, ",")(1)` 的 AST 里内层调用与外层
  `Base64Decode` 的缺档补丁（`cgen_expr_call_builtin_fixup.inc:326 kOD_Split`）叠在一起，
  结果外侧 `Base64Decode` 的缺省尾巴 `1, NULL, -1, 0` 被拼进了**内层 `vb6_Split(...)` 的
  实参表**。判据是那 4 枚恰好是外侧调用的缺省档而 `kOD_Split` 只有 4 档。
- **B. 模块级 `Declare` 被同名的内置出口遮蔽**：`GetObject` 在 VB6 里是 gdi32 的 API
  （本模块自己 Declare 了），C3 却按内置 `vb6_GetObject(pathName, progId)` 发。VB6 语义：
  **本模块的 Declare 优先于同名内置**（`Split` 那份内置表在
  `cgen_expr_ident_builtin_table.inc:131`，遮蔽闸门要加在同一处）。

⚠ **本机 `.build\C3.exe` 是旧产物（04:52 构建，其内嵌 RTL 的 `vb6_Split` 是 8 形参），
所以本机编译能过、CI 新编出来的过不了。接手第一步必须 `scripts\build.bat` 重建，否则一切
本机复现都是假象。**

下一步（按顺序）：
1. `scripts\build.bat` 重建 → 本机复跑 `C3 tests\ucSegmentedControl\SegmentedControl.vbp --arch x86`
   应出现同样 6 条 C2197（否则说明还有别的差异，先查 RTL 头）。
2. 修 A：把缺档补丁改成在**实参列表层**拼接（不要对渲染后的字符串做括号定位）。
3. 修 B：在 `cgen_expr_ident_builtin_table.inc` 命中前先查"本模块/本工程同名 Declare 或过程"。
4. 回归：`test.bat gui` 10/0/0 + `check_builtin_opt_defaults.ps1` / `check_builtin_slot_returns.ps1`
   / `check_manifest_coverage.ps1` 三个哨兵绿；再提交、双远端快进、GA 全绿。
5. 修完 codegen 一定又会让若干份输入的发码形状变化 ⇒ 按 §7.6 的口径用 CI 的
   `ci/emit-manifest` 逐行登记（**不要**拿本机哈希覆盖那 4 份依赖外部 OCX/typelib 的输入）。

## 7.8 本轮收口状态（停点）

GA run **41917073251**（head=75c61f4a）：`Tests (compile #1)` 已绿（`builtin_slot_returns`）；
`Tests (compile #2)` 红在 `[STATIC] owned_temp_keeps`（K13c：RTL 新 mint 了
`vb6_VariantFromParamArray`，发码侧 `src\common\owned_temp_keeps.hpp` 的 `kOwned` /
`kOwnedVariant` 两张名单没有它）；`Tests (vbp #2)` 红在 `[GUI] ucSegmentedControl ... FAIL (compile)`
（就是 §7.7 的 Split/GetObject 两条）。K13 本轮已修并本机 PASS：

- `src\common\owned_temp_keeps.hpp`：两张名单各补一行 `vb6_VariantFromParamArray`（VARIANT 档）；
- `scripts\owned_temp_keeps_164.txt`：`[own]` 节补同名一行（双向对账）。
- 本机复跑 `check_owned_temp_keeps.ps1` = PASS（owned 115 / alias 17 / U 5，nomin 184）。

**未完成（下一轮从这里继续）：** §7.7 的 A（`Split(s, d)(1)` 的省略实参补档多出 4 枚，
`vb6_Split` 实际只剩 4 形参）与 B（模块内 `Declare Function GetObject ... Alias "GetObjectA"`
被内置 `vb6_GetObject` 遮蔽）。修完还要重跑 `test.bat gui`（10/0/0）、重登记发码变化
（用 CI 的 `ci/emit-manifest`）、再推 GA。

> 进展更新：GA run **41919273594**（head=bf288a31，即上面 K13 修复之后）**只剩
> `Tests (vbp #2)` 一个红**——就是本节这两条；`compile #1/#2`、`Emit manifest`、其余
> vbp 分片、bas/syntax/asm/smoke 全绿。⇒ 收敛点已经很清楚：只差这两个发码缺陷。

**两个本机陷阱（下一轮开头先处理，否则复现全是假象）：**

1. `scripts\build.bat` 在本机 checkout 里是 **LF-only**（CI 那边是 CRLF），cmd 会把
   `if errorlevel` / `echo` 这类行啃掉首字符 ⇒ 必须复制成 CRLF 副本再用
   `Start-Process -FilePath $env:ComSpec -ArgumentList '/c','<crlf 副本>'` 跑（并设
   `C3_PROJECT_DIR` + `C3_VCVARSALL`）。
2. `build.bat` **不重建** `.build\pre\src\driver\c3rtl.rc`（它停在 11-04 13:51，而
   `src\driver\c3rtl.rc` 是 11-04 19:54）⇒ 本机 `C3.exe` 内嵌的仍是**旧 RTL**
   （那个旧 RTL 的 `vb6_Split` 是 8 形参），所以本机能编过、CI 新编的编不过。
   判断"本机产物是否可信"的唯一口径：`--emit-c` 出来的 `vb6_Split(` 形态与
   `src\rtl\core\vb6rtl\vb6rtl_array.h:116` 的形参数是否一致。

## 7.9 §7.8 的两个缺陷已修（本机闭环），待推送 + GA

### 缺陷 A：调用结果再下标被并进内层实参表

- 根因：P6.5 那段"callee 已经是 `func(obj)`"的**通用拆分**（
  `src/backend/detail/expr/cgen_expr_call_com_bind.inc`，`callee.back() == ')'`
  且外层有 positional）对**任何**以 `)` 结尾的 callee 都生效。`Split(SrcImg, ",")(1)`
  的 callee 是内层 Split 的完整调用，于是外层的 `1` 被接到内层实参尾部；随后
  内置补档按 `vb6_Split` 名再补 `NULL, -1, 0` ⇒ 8 实参（C2197）。
  这是**括号层级/实参归属**问题：外层括号是下标，不属于内层调用。
- 修法：在同一处加一条**先于** P6.5 的括号层级闸 —— callee 已是 SafeArray1D* 载体
  调用（`isSafeArrayCarrier` 名单内的 `vb6_Split(` / `vb6_Filter(` /
  `vb6_ArrayCreate(` / 匿名临时量 `_arr_`）时，外层实参交给 `VB6_SA_AT(<elem>, <内层调用>, <下标>)`，
  **绝不**并入内层实参表。元素档随载体（Split/Filter→BSTR，Array→vb6_VARIANT）。
  没有在内层按形参个数特判。
- 发码对照（`tests/ucSegmentedControl/ucSegmentedControl.ctl:2166`）：
  修前 `vb6_Split(src, delim, -1, 0, 1, NULL, -1, 0)`（8 实参）
  → 修后 `VB6_SA_AT(BSTR, vb6_Split(src, delim, -1, 0), 1)`；
  RTL `src/rtl/core/vb6rtl/vb6rtl_array.h:116` 声明
  `vb6_Split(BSTR, BSTR, int32_t, int32_t)`（4 形参）⇒ 两侧一致。
  `--emit-c` 全量 diff 只有这 3 个站点的 6 行（其它"调用结果再下标/再成员"形状未受影响；
  `Array(...)(i)` 同族也确认按新规则发码）。

### 缺陷 B：显式 Declare 被内置同名遮蔽

- 根因：`src/backend/detail/expr/cgen_expr_ident_builtin.inc` 的**静态内置名表**
  在符号表查询**之前**命中，`{"getobject","vb6_GetObject"}` 把模块级
  `Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA"(3 参)`
  顶掉 ⇒ 3 个实参落进 2 形参的内置出口（C2197）。`symbol_table.cpp` 的 `define()`
  已经让内置符号整体让位（Fix 048 那条），但这里没问符号表。
- 修法：命中内置名时先问本模块作用域 —— 有 `DeclareSub/DeclareFunc` 且带形参表
  （存在性 + 参数形状；Lib/Alias 由 Declare 符号本身携带，发射侧登记在
  `declareAliasMap_`）⇒ **Declare 赢**，落回原有别名转发桩。闸只在"真有 Declare"，
  没有 Declare 的同名站点（如 `GetObject(, "Excel.Application")`）照旧走内置。
- 发码对照（`ucSegmentedControl.ctl:2195/2211`）：修前
  `vb6_GetObject(vb6_VariantToLong((*SrcImg)), sizeof(tBmp), &tBmp)`（3 实参）
  → 修后 `vb6_lwa_GetObjectA(0L, ((int32_t)sizeof(bmp)), (void*)&(bmp))`。

### 两个独立回归夹具（不只靠 segmented 兜底）

- `tests/test_split_result_subscript.bas`：`Split("a,b,c", ",")(1)` + `Array(10,20,30)(1)`。
  探针 `Test-CodegenNote "split_result_subscript"` 两头钉：
  needle `VB6_SA_AT(BSTR, vb6_Split(..., -1, 0), 1)`（绑定层）、
  absent `..., -1, 0, 1` / `-1, 0, 1, NULL` / `_arr_0(1)`（实参个数与旧合并形状）。
- `tests/test_declare_shadows_builtin.bas`：3 参 Declare GetObject。
  探针 `Test-CodegenNote "declare_shadows_builtin"` needle `vb6_lwa_GetObjectA(...)`、
  absent `vb6_GetObject(`；对照针 `"getobject_builtin_kept"`
  （`test_p24_04b/test_getobject.bas`）证明真内置站点没有被顺带改掉。

### 重建 RTL 嵌入 / 真跑

- `.build\Release\c3rtl.res` 06:08 重新生成（`src/driver/c3rtl.rc` 吃当前 RTL 头）；
  本机 segmented 编译 **rc=0**，`output/ucSegmentedControl/compile.log` 无
  `c3-error.log` ⇒ CI 上那 6 条 `C2197` 消失。
- `tests/run_tests.ps1 -Category gui`：**PASS=10 FAIL=0 SKIP=0**。
  ucSegmentedControl `window=766x558 client=751x520 hosts=12+1`，
  `uc2 145/138, uc3 1185/625, uc4 1734/1734, uc11 970/209, uc1 802/161`，
  `#0078D7=10715`、`#CE00EF=400` —— 与本轮修前**逐项相同**；
  `output/ucSegmentedControl/pixels/segmented_c3.png` 与修前
  `.temp/segpx/segmented_c3.png` **字节相同**（sha256 `4484B813…`）⇒ 对参考的
  `MEAN_ABS_DIFF 3.967/765` 不变（不是重算，是图没动）。

### 发码清单重登记（未用 `-Bless`）

- `emit-manifest.expected.txt` 本机现算：
  新增 `tests/test_split_result_subscript.bas`、`tests/test_declare_shadows_builtin.bas`；
  更新 `tests/ucSegmentedControl/SegmentedControl.vbp`（`sha256=e1f248b1…`）。
- 其余 4 份仍是**本机与 CI 不同**的既有条目（TeeChart 与 GUISample×3，依赖外部
  OCX/typelib 解析），保持 CI 值不动；最终以 CI 的 `ci/emit-manifest` 复核。
- `check_manifest_coverage.ps1`、`check_builtin_slot_returns.ps1`、
  `check_builtin_opt_defaults.ps1`、`check_owned_temp_keeps.ps1` 本机 PASS。

## 7.10 把 `Tests (vbp #2)` 的像素判据改成宿主无关（本机已验，待 GA）

上一跑 41929711664（head `c3867290`）只剩 `Tests (vbp #2)` 红，且性质已从"编不过"变成"判据在 CI 上不成立"：
6 条 `C2197` 已消失（`compile.log` 无 error C），失败只有三条：

- `window=767x559 != 766x558`（CI 客户区仍是 751x520，只差非客户区 1px）；
- `#0078D7` 数到 **0**（<8000）——CI 宿主的系统高亮色不是本机这一档；
- `#CE00EF` **237**（<250）。

三处都改在 `tests/guisample/segmented_px.ps1`，判据语义不变、不再依赖本机：

1. **外框**：客户区 `751x520` 仍精确钉死（= Form1.frm 11265x7800 twips/15）；
   外框改为运行期 `GetSystemMetrics` 现算
   `client + 2*(SM_CXSIZEFRAME+SM_CXPADDEDBORDER) + SM_CYCAPTION`，容差 1px
   （120dpi 虚拟化在非客户区取整）。本机实测 `766x558` vs 现算 `767x559`（差 1，PASS），
   CI 实测 `767x559` 与现算一致（差 0，PASS）。旧断言把本机的 1px 取整当产品判据。
2. **选中段底色**：控件画的是 `vbHighlight`（`ucSegmentedControl.ctl:1864`/`:1991`），
   判据改为数**运行期** `GetSysColor(COLOR_HIGHLIGHT=13)` 的颜色像素，仍要求 >= 8000。
   本机该值就是 `#0078D7`，读数 `highlight=#0078D7(10715)` 与硬编码时**一致**（不是放水）；
   CI 换一档系统高亮色时按它自己那档计数。
3. **TAB2 品红指示条**：色彩来源是 `TabIndicatorColor=15663310`（= `#CE00EF`），
   指示条宽度 = 选中项宽 - Radius，而项宽 = `TextWidth + Padding*2`
   （`ucSegmentedControl.ctl:362-368` / `:739-742`）⇒ **像素总量随宿主字体度量变**
   （本机 101x4=400，CI 字体回退后 237）。所以不再钉绝对像素数，改钉结构：
   品红族像素 >= 150 且 bbox 高度 2..6（`TabIndicatorSize=4` 名义厚度）、水平跨度 >= 40。
   "没画/画成竖条色块"仍红；读数行同时保留精确 `CE00EF=400` 与族计数/bbox。

本机复跑 `tests\run_tests.ps1 -Category gui`：**PASS=10 FAIL=0 SKIP=0**，
读数与本轮修前相同（`uc2 145/138, uc3 1185/625, uc4 1734/1734, uc11 970/209, uc1 802/161`；
`highlight=#0078D7(10715)`、`CE00EF=400 magenta=403 bbox=392,115-492,118 101x4`；
`expect_outer=767x559` vs 实测 `766x558`）。

## 8. 风险 / 注意

- `Manifest.RES` 是**进程清单**（非运行期 SxS manifest）；夹具化时不要把它塞进 SxS 装配路径，否则与两条免注册 COM 路径互相干扰。
- `PropPagSC.pag` 只服务设计器属性页，**不影响运行期**，夹具登记时可只保源码不参与构建。
- `Form1.frx` 157KB 里既有 DocString 也有 PNG 图片块；新增 `readDocBlobBytes` 用 `detectFormat` 做闸门，**不得**把普通字符串属性误判成二进制。
- C3 成功编译会删除 `%TEMP%\C3C\<hash>_<pid>` 中间目录；要看生成 C 必须用 `--emit-c`，别当成缓存。
- 参考 exe 主窗 `Title=[LabelPlus]` 来自 `.vbp` 的 `Title`；C3 产物 `Title=[Form1]` 来自 `Form1.frm` 的 `Caption`。**以窗体 Caption 为准**（参考的可视窗也是 `[Form1]`）。
- 本机 exe 安全警告靠仓库根 Medium 完整性标签解决，**不要再加低完整性标签**；脚本用 `Start-Process -RedirectStandardOutput/-RedirectStandardError -NoNewWindow`。
- 采集脚本坑：委托（`EnumProc`）里的 `Write-Output` 不进管道 → 必须先收集到 `ArrayList` 再统一输出（`.temp\probe_ref.ps1` 已按此写法）。
