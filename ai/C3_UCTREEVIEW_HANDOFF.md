# ucTreeView 工程（C3 编译还原）交接文档

## 0. 任务与口径（不可变）

- 工程：`tests\ucTreeView\Proyecto1.vbp`（源自 `D:\vb_yqt4qPac\ucTreeView\`，夹具已入库，不依赖 D:）。
- 真 VB6 参考产物：`D:\vb_yqt4qPac\ucTreeView\Proyecto1.exe`（32 位）；参考图 `Proyecto1.png`。
- 架构：只能 `--arch x86`；产物必须 PE `machine=0x014C` + 运行后 `IsWow64Process=TRUE`；不符要明确 FAIL。
- 不许猜/不许静默误编：先编、分类（生成 C / cl / 链接 / 崩溃 / 绘制或事件静默无效）、给证据（日志、`C3_CRASH_TRACE`、`C3_OCX_TRACE`、`C3_HOST_438`、像素）再修。
- 不许放宽断言 / 删用例 / SKIP 当通过；缺资产才许 SKIP，且"资产在却失败判红"。
- 1:1 基准 = 真 VB6 产物运行截图（`VB6.EXE /make`，96dpi 虚拟化，截图用 `CopyFromScreen`，禁用 `PrintWindow`）。
- RTL 改动必须重建内嵌 RTL：重新生成 `c3rtl.res` 并嵌入 `C3.exe`；核对新符号能在 `C3.exe` 字节内搜到；核对 `.build` 的 `c3rtl.rc`/`c3rtl.res` 时间戳。
- 免注册 COM：`LoadLibrary + DllGetClassObject` 与 SxS manifest 两条路径都要能编能跑。
- 仓库根完整性必须 Medium（否则启动弹"打开文件 - 安全警告"）；测试启动用 redirect/`-NoNewWindow`。
- `SmoothScrollFactor` 无效值只告警不致命（与 VB6 一致），不得为消警告改夹具。
- 流程：小步提交，推 `github`/`origin` 的 `dev`（快进、不 force），GA 22/22 全绿再进下一步。

## 1. 已落盘修复（RTL 侧，已重建并验证内嵌）

`.build\C3.exe` 于 2025-11-05 21:06 重建成功（`c3rtl.rc.res` 21:06:17，`C3.exe` 21:06:44），
字节内已搜到 `vb6_UC_SetFontFlag` / `vb6_UC_ExtenderZOrder` / `vb6_UserControl_FontBold`。

1. `src/rtl/core/vb6rtl/vb6rtl_userctl.h`
   - `struct vb6_UserControl_Extender_Type` 末尾新增 `Parent/Top/Left/Name/ZOrder` 五字段
     （对得上 ucScrollbar.ctl 的 `UserControl.Extender.Parent/Top/Left/Name/ZOrder`，C2039 ×8）。
   - 新增 `extern int16_t vb6_UserControl_FontBold/FontUnderline;` 与
     `vb6_UC_SetFont(void*)` / `vb6_UC_SetFontFlag(int,int)` / `vb6_UC_ExtenderZOrder(int32_t*,int32_t)`。
2. `src/rtl/core/vb6rtl/vb6rtl_com.c`
   - Extender 全局初始化补 5 个新字段 + 两个字体标志定义。
   - `vb6_UserControl_TextWidth` 造 LOGFONT 时叠加 `FontBold/FontUnderline`（否则测宽不含粗体）。
3. `src/rtl/core/vb6forms/vb6forms_uc_internal.h`
   - `vb6_UCRec` 增 `fontBold/fontUnderline`；`vb6_UCSaved` 增同名字段（push/pop 存档）。
4. `src/rtl/core/vb6forms/uc/uc_host.c`
   - push：`ext` 存档 + 新字段填 `Parent=r->parent, Top=r->extTop, Left=r->extLeft, Name=r->displayNameBstr, ZOrder=vb6_UC_ExtenderZOrder`；字体标志存档/装载。
   - pop：恢复 ext 与两个字体标志。
5. `src/rtl/core/vb6forms/uc/uc_host_window.c`（新增约 90 行，插在 `vb6_uc_selectOwnFont` 之后）
   - `vb6_uc_fontHFontFor(font,bold,underline)`：按 (Font 对象, Size, bold, underline) 缓存 HFONT（16 槽，不销毁）。
   - `vb6_uc_applyFontNow()`：把当前 `vb6_UserControl_Font` + 两个标志**当场 SelectObject 进活动 paint DC**；
     同时把 `vb6_ucCurrent->font` 同步为新字体（`Set UserControl.Font` 持久）。
   - `vb6_UC_SetFont` / `vb6_UC_SetFontFlag` / `vb6_UC_ExtenderZOrder`。
6. `src/backend/cgen_base.cpp`（文本 fixup）
   - `fixupUserControlParentChain` 增 `.BackColor` → `vb6_Ambient_BackColor`（ucScrollbar.ctl:2512）。
   - 新增 `fixupExtenderParentRunMode`：`<...>_prop_get_Extender(...).Parent.RunMode` → `vb6_Ambient_UserMode`（ucScrollbar.c:2976；该 .ctl 自己的 Extender 是 Property Get，不是 UserControl.Extender）。
   - `emitLine` 增该分支。

## 2. 生成码侧修复（本轮新增，**尚未重编夹具验证**）

1. `src/backend/detail/stmt/cgen_assign_host_pseudo.inc`
   - 新增 `vb6_UserControl_FontBold/FontUnderline` → `vb6_UC_SetFontFlag(0/1, …)`；
   - 新增 `vb6_UserControl_Font` → `vb6_UC_SetFont((void*)…)`（裸赋值形态兜底）。
2. `src/backend/detail/stmt/cgen_setlet_set_rhs.inc`
   - `Set UserControl.Font = x` 走 Set 语句那条路：目标是 `vb6_UserControl_Font` 时改发 `vb6_UC_SetFont((void*)x)`。
   - 定位依据：生成码 `vb6_UserControl_Font = me->m_SubTextFont;  /* Set */` 就是本文件末尾的 `target + " = " + value + ";  /* Set */"`。
3. `src/backend/detail/expr/cgen_expr_member_class_fallback.inc`
   - Fix 088c 回退臂 `classFieldReadExpr(..., /*isVoidPtr=*/false, ...)` → `classFieldIsComPtr(clsFromAst, node.memberName)`。
   - 根因：`oParent->Parent->Childs.Count`（Node.cls `Childs As Collection`）中段 `Childs` 不带 `voidptr` 尾注，
     外层 `.Count/.Add/.Remove` 直发结构体成员（ucTreeViewPlus.c:2588 C2224；Form4.c:382/410）。
     单级 `DestNode->Childs` 已带尾注并走 COM 分派（Form4.c:390 是正例）。
4. `src/backend/cgen_base_naming.cpp` + `detail/util/cgen_helpers.inc`
   - 新增 `classFieldIsComPtr(cls, member)`：查 `classVoidFieldMap_`（含 `m_` 前缀回退）与
     `classTypedFieldMap_` 的 `COM:` 前缀；类名键大小写不敏感兜底。
   - 新增 `dataObjectArgExpr(name)`：按 `currentProc_` 里同名形参的 `isByVal` 决定
     `(void*)name`（ByVal ⇒ C `void*`）或 `(void*)(*name)`（ByRef ⇒ C `void**`）。
5. `src/backend/stmt/cgen_call.cpp`（两处：约 595 / 705 行）
   - 省略实参补齐由 `defaultValue(param.type)` 改 `defaultValueForParam(param)`（类形参 → `NULL`），
     ByRef 分支的 C 类型改用 `cParamClassPtrType`（空则退回 `mapType`）。
   - 根因：`ExpandAllNodes(inst, vb6_VariantEmpty(), 0)` vs 形参 `vb6_cls_Node*`（Form3.c:149/156、Form6.c:432 C2440）。
6. `src/backend/detail/stmt/cgen_assign_stmt_special.inc`
   - `isUC292` 收窄：只有当 UC 类**自己**用 `findClassMemberWriteParams` 声明了同名可写成员时才把 `writeFn` 置空；
     否则仍用 `getControlPropWriteFn`。
   - 根因：`ucTreeViewPlus1.Height = …` 落到晚绑定，LHS 被发成 `vb6_GetControlHeight(hwnd) = v`（Form1.c:987 / Form2.c:249 C2106）。
7. `src/backend/cgen_util_ctrl.cpp`
   - PictureBox 读表增 `image` → `vb6_FormGetImagePicture`（Form6.c:463 `_vb6_with_2.Image` C2039）。
   - RTL 侧该函数早已存在（`vb6forms_picture_prop.c`，按 `VB6_AutoRedrawBmp` 窗口属性取图），与窗体 Image 同实现。
8. `src/backend/detail/expr/cgen_expr_member_precheck.inc` / `cgen_expr_call_callee_member.inc` / `cgen_util_com.cpp`（2 处）
   - 三处硬写 `(void*)(*<name>)` 改为 `dataObjectArgExpr(name)`。
   - 根因：`ByVal Data As DataObject` 的 C 形参是 `void*`，`(*Data)` 非法（Form4.c:422 C2100）。
9. `src/backend/detail/expr/cgen_expr_call_arg_emit.inc`（ByRef 实参 else 分支）
   - 实参是**用户控件实例名**（`knownUserControlCtrlVars_` 命中）且被调方形参是类指针（`cParamClassPtrType` 非空）时，
     不再发 `&TVP`，改发复合字面量临时槽：
     `&((vb6_cls_X*){ (vb6_cls_X*)vb6_UC_InstanceOf(vb6_hwnd_TVP) })`。
   - 根因：控件在生成 C 里只有 `vb6_hwnd_TVP`，没有同名变量（Form2.c:240 C2065）。

## 3. 剩余错误逐类清单（旧日志 `.temp\uctv_err3.txt`，40 条；行号为旧中间目录）

| # | 站点 | 错误 | 根因 | 处置 |
|---|------|------|------|------|
| 1-8,10 | Node.c:479-513 | C2001/C2146/C2065 `Path` | `Optional PathSeparator As String = "\"` 默认值把 VB 源码形态直接抄进 C 宽字面量 | 已修 `semantic_analyzer_util.cpp` `evalOptionalDefault`（反斜杠/引号/换行转义），待重编验证 |
| 9 | ClsTreeExplorer.c:249 | C2102 `&`要求左值 | `&(TOOLTIPS_CLASS)`：Const 的 C 侧是 `#define` 文本 | 已修 `cgen_expr_call_arg_emit.inc` `isAssignableOutTarget161b`（按 SymbolKind 拒 Constant/EnumMember），待验证 |
| 11-21 | ucScrollbar.c:863-988 | C2039 `Parent/Top/Left/ZOrder/Name` + C2198 `vb6_CStr` | `vb6_UserControl_Extender_Type` 缺字段 | RTL 第 1 节已加；C2198 是字段解析失败级联 |
| 22-23 | ucTreeViewPlus.c:2588 | C2224 `.Count` + C2198 | 链式 `Parent->Childs` 缺 voidptr 尾注 | 第 2.3 节 |
| 24-27,30 | ucTreeViewPlus.c:2667-2674, 2919 | C2065 `vb6_UserControl_FontBold/FontUnderline` | RTL 无声明 + 无落 DC 入口 | RTL 第 1 节 + 第 2.1 节 |
| 28 | ucScrollbar.c:2512 | C2224 `.BackColor` | `vb6_UC_ParentObject().BackColor` | `cgen_base.cpp` `.BackColor` fixup |
| 29 | ucScrollbar.c:2976 | C2224 `.Parent` | `_prop_get_Extender(...).Parent.RunMode` | `fixupExtenderParentRunMode` |
| 31-32,39 | Form3.c:149/156、Form6.c:432 | C2440 VARIANT→`vb6_cls_Node*` | 省略类形参补 `vb6_VariantEmpty()` | 第 2.5 节 |
| 33 | Form2.c:240 | C2065 `TVP` | UC 控件无同名 C 变量 | 第 2.9 节 |
| 34-35 | Form1.c:987、Form2.c:249 | C2106 左值 | `isUC292` 一刀切置空 writeFn | 第 2.6 节 |
| 36-37 | Form4.c:382/410 | C2224 `.Remove/.Add` | 同上链式字段缺尾注 | 第 2.3 节 |
| 38 | Form4.c:422 | C2100 | `(void*)(*Data)` 对 ByVal `void*` | 第 2.8 节 |
| 40 | Form6.c:463 | C2039 `Image` | PictureBox 读表缺 `image` | 第 2.7 节 |

## 4. `SmoothScrollFactor`（用户补充前提）

- 源码：`ucScrollbar.ctl:812-819` 是**用户自定义** Property Get/Let（`m_SinSmoothScrollFactor`），不是内建属性；
  `ucTreeViewPlus.ctl:31/50` 的 `SmoothScrollFactor=   0,15` 是设计期属性袋里的**非法数值写法**（逗号小数）。
- 真 VB6：报 `行 31/50: 属性 SmoothScrollFactor在 ucScrollbar1/2 中有无效值`，**只告警**，忽略该值、用默认（`ReadProperties` 的 0.3），照常出 exe。
- C3 现状：**不报错、不告警、静默丢弃**该设计期非法值（生成码里没有来自 `0,15` 的 SetProp），
  与 VB6 的"产物行为一致、诊断少一条 warning"；这**不是** C3 缺陷，不改夹具、不放宽断言。
- 是否被消费：**是**。`ucTreeViewPlus.ctl:2705/2710` 的代码 `ucScrollbar1/2.SmoothScrollFactor = 1` 生成
  `vb6_ComSetProp(hwnd, L"SmoothScrollFactor", vb6_ComPackInt(1))`，落到 `prop_let_SmoothScrollFactor` →
  `m_SinSmoothScrollFactor`，再被 `SmoothStep = Abs(vstep) * m_SinSmoothScrollFactor`（ucScrollbar.c:1919）消费。
  所以"被忽略但看起来一样"的只是那个非法设计值，属性本身已实现。

## 5. 复用的既有结论

- `ai/C3_SEGMENTED_HANDOFF.md`：UC 宿主模型、`ForeColor`/`MousePointer` 落 DC、`vb6_UC_RefreshCurrent` 空实现、gdiplus 桩位宽。
- `tests\guisample\segmented_px.ps1`：`GrabRegion` 被动抓取、光标断言、点击耗时断言写法。
- `tests\guisample\tree_icons_px.ps1`：TreeView ImageList 透明/mask 断言。

## 6. 本机陷阱（实测踩过）

- `PowerShell` 工具禁 `cmd /c`：构建用 `Start-Process <bat>`；`.temp\build_c3.bat` 里的 `cd /d` 必须带引号（路径含空格），
  旧文件曾写成裸 `cd /d` + 换行，静默失败但构建在继承 cwd 里仍能跑。
- 构建：`vcvarsall.bat x64`（`D:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat`）
  + `cmake --build .build --target c3`。日志写 `.temp\build_c3h.log/.err`。
- RTL 重建核对：`Get-Item .build\CMakeFiles\c3.dir\src\driver\c3rtl.rc.res` 时间戳；用字节搜索确认符号进入 `C3.exe`。
- 编译夹具命令：
  `.\.build\C3.exe tests\ucTreeView\Proyecto1.vbp --arch x86 --output-dir output\ucTreeView`
  错误日志在 `output\ucTreeView\c3-error.log`，中间 C 保留在 `C:\Users\ADMINI~1\AppData\Local\Temp\C3C\<id>`。
- `C3` 会打 `warning VB4001: Unknown control property '.'Image' in With block` —— 第 2.7 节修后应消失。
- 完整性标签 Medium、截图 `CopyFromScreen`、`IsWow64Process=TRUE`、PE `machine=0x014C` 判据均在后续步骤。
- 步数/上下文压力：每清一类即提交；本文件随时补写"下一轮第一条"。

## 7. 阶段状态（续跑轮次更新，2026-11-05 23:2x）

- 编译：**cl 40 → 0**；链接：**LNK2019 22 → 0**；产物 `output\ucTreeView\Proyecto1.exe`（818,688 B，23:09:50）已出。
- PE `machine=0x014C` ✓、`IsWow64Process=TRUE` ✓、仓库根完整性 Medium ✓。
- 运行：`[C3_HOST_438]` 消失；**启动成功**（2026-11-05 23:35，`title=Form1`，见第 10 节）。
- 仅余 2 条 warning（均非致命，见第 11 节）。

## 8. 链接期 22 条 LNK2019 的逐类根因与修改站点

1. **19 条 `vb6_di_*` 未定义**（GDI+/theme/shell/version 导入桩缺失）
   - 处置：按**失败 session 目录**生成桩，未改生成器：
     `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\gen_di_stubs.ps1 -SessionDir "$env:LOCALAPPDATA\Temp\C3C\<失败session>"`
   - 实测补齐：`vb6_di_GdipAddPathPolygon / SetRectEmpty / InitCommonControls / SHGetFileInfoA /
     GetCurrentThemeName / PathIsDirectoryA / GdipGetPathWorldBoundsI / GdipCreateMatrix /
     GdipTranslateMatrix / GdipRotateMatrix / GdipTransformPath / GdipSetPenDashStyle /
     GdipGetImagePixelFormat / GdipGetImageDimension / GetThemeDocumentationProperty /
     FindFirstFileA / FindNextFileA / SHGetSpecialFolderLocation / GetLogicalDriveStringsA`。
   - 注意：脚本是 **add-only 增量合并**（保留旧符号、删过期文件），可重复跑。
2. **`_vb6_VBA_Left_`**（`Node.GetFullPath` 的 `VBA.Left$(Path,1)`）
   - 站点：`src/backend/detail/expr/cgen_expr_member_precheck.inc`（VBA 分支 `vbaFn` 命中之后追加
     `builtinFuncs` 兜底，并在 fnName 上 strip `$`）；`src/backend/expr/cgen_expr_member.cpp`
     补 `#include "backend/detail/expr/cgen_expr_ident_builtin_table.inc"` 提供 `builtinFuncs`。
   - 生成码实测：`vb6_Left(...)`（旧 session 行 3699）。
3. **`_vb6_ucTreeViewPlus_prop_set_SelNode`**（.ctl 里是 `Public Property Let SelNode(...)`）
   - 站点：`src/backend/detail/module/cgen_form_uc_props.inc`
     - `UcPropPick` 增 `std::string setSeg`；PropertyLet/PropertySet 臂记 `prop_let_` / `prop_set_`；
       merge 传播 `setSeg`；两处发射点改用真实动词。
4. **`_Cls`**（`ucTreeViewPlus.ctl:1408` 裸写 `Cls`，被发成裸标识符 `Cls()`）
   - 站点：`src/backend/cgen_util_com.cpp` 新增 `hostPseudoMethodRtlName()`（`HPF_METHOD` 且非 `HPF_BARE`
     即映射到 `vb6_UserControl_Cls`）；声明在 `src/backend/detail/util/cgen_helpers.inc`；
     调用点 `src/backend/detail/expr/cgen_expr_ident_builtin.inc`（`hostPseudoBareName` 分支之后）。

## 9. UC 自有方法桥（运行期 438 根因，三处叠加）

现象：`[C3_HOST_438] site=Call name="TrackMouseWheelOnHwndStop"`（`With ucScrollbar1: .TrackMouseWheelOnHwndStop`，
`ucTreeViewPlus.ctl:832`；目标 `ucScrollbar.ctl:2228 Public Function TrackMouseWheelOnHwndStop()`）。

站点：`src/backend/detail/module/cgen_form_uc_methods.inc`
1. 只收 `SymbolKind::Sub` 且排除空形参表 ⇒ 该 **无显式返回类型的 Function + 0 形参** 根本没生成桥。
   改为：`Sub` 与 `Function` 都收；Function 仅收 `type == Variant`（无显式返回类型）；允许 0 形参
   （`params.size() <= 4`）。
2. 首次放开后 cl 报 **C2440 int→vb6_VARIANT**（`vb6_ucTreeViewPlus_Array` 等）：
   内置符号（`isBuiltin`，`VBA.Array`/`CallByName`）被误当自有方法 ⇒ 加 `if (ps->isBuiltin) continue;`。
3. 编译链接通过但仍 438：新 thunk 的 `if (!me || !argv) return 0;` 对零参调用（调用点传 `argv=NULL`）
   一律判未命中 ⇒ 改
   ```c
   if (!me) return 0;
   if (argc > 0 && !argv) return 0;
   ```
   并给 `Function` 桥写回 `*outRet`（`vb6_VARIANT retMeth = <call>; if (outRet) *(vb6_VARIANT*)outRet = retMeth;`）。
- 实测生成：`vb6_ucScrollbar_ucHostDesc` → `vb6_ucScrollbar_ucMethods, 17`；`..._ucMeth_TrackMouseWheelOnHwndStop` 存在。
- 实测运行：启动后不再出现 `[C3_HOST_438]`。

## 10. 启动期崩溃（已收口，2026-11-05 23:35 实测启动成功）

**现象**：启动后 2–3 秒退出，exit `0xC000041D`（回调内异常的外层包装码）。

真实异常（`C3_CRASH_TRACE=1` 原始输出 `.temp\uct_err.txt`）：`code=0xc0000005`，
第一次 `av read target=0x650060`（oleaut32!SysStringLen 读 `bstr[-4]`），
栈（map 按 `Rva+Base − 0x400000` 对齐）：

```
_vb6_BSTR_Len +0x17 ← _vb6_Len +0xC
← _vb6_ucTreeViewPlus_RecurCalcTreeWidth +0x123
← _vb6_ucTreeViewPlus_Update +0x8D ← UserControl_Resize ← ucHostResize ← uc_wndproc
```

即生成码 `RecurCalcTreeWidth` 的 `Len(N->SubText)` 拿到坏 BSTR。第一次 Resize 时树还没有节点，
循环体不执行；第二次 Resize（Show 之后）才进循环 ⇒ 第一次崩溃点。

**根因（两层，逐层实测）**：
1. `N = (vb6_cls_Node*)vb6_ComUnpackObject(&v)`（`vb6com_pack.c:247` 附近）此前只做
   `vb6_UC_UnwrapHost`，对工程类包装器 `vb6_ComObject`（vtable/refCount/desc/vb6Instance，
   x86 偏移 12 才是实例）**不还原实例** ⇒ 直接把包装器当 `vb6_cls_Node*` 按字段读。
   修：`vb6com_pack.c` 在 UC 解包之后加
   `if (vb6_ComObject_IsBoxed(obj)) { inst = vb6_ComObject_GetInstance(obj); if (inst) obj = inst; }`
   （`vb6comserver_obj.c` 已有这两枚函数；IsBoxed 带 IsWindow/可读闸，外部 IDispatch 原样返回）。
   → 第一次 av（0x650060）消失，改为第二层的 NULL。
2. `vb6_ComUnpackObject`→`vb6_ComObject_GetInstance` 返回 **`vb6Instance=NULL`**（`C3_IV_TRACE=1` 实测：
   `[IV] OK pdisp=00BBDEB8 inst=00000000`，而同一地址在早前一次是 `inst=00BC4A00`）。
   根因是 **Collection 借用了对象引用却不计数**：`vb6_coll_push` 只做 `vb6_VariantValueCopy`
   （浅拷贝 dispatch 指针、不 AddRef），打包方那次 `VariantClear` 把包装器 free 掉，
   集合里的 `pdispVal` 悬空（地址被复用后 `vb6Instance` 为 0）⇒ For Each 拿到 NULL。
   修（`uc_collection_api.inc`，三处配平）：
   - `vb6_coll_push`：dispatch 档 `vb6_ComAddRefDispatch`（集合自己持有一份）；
   - `vb6_coll_itemCopy`：取回的格子自己持有一份（`Item`/`EnumNext` 共用，调用方随后放掉）；
   - `vb6_Collection_Remove`：`vb6_ReleaseObject(&v->pdispVal)` 与 push 的 AddRef 配对。
   第二层崩溃签名：`at rva=0x117CE`（`RecursiveDraw +0xDEE`），`av read target=0x84`（`N=NULL` 后读 `N->Height`）。

**收口实测（23:35）**：重编 `c3rtl.rc.res`（2,884,212 B，23:33:53；C3.exe 23:33:54）→
`C3 --arch x86` 退出 0（无 error，仍只有 2 条 warning）→ 产物 818,688 B（23:35:32）→
启动 **存活**：`pid=4160 title=Form1 hwnd=9569580`，stderr 无 `[C3_CRASH]`、无 `[C3_HOST_438]`。
stdout 的 `THE IMAGE: n DON'T EXIST` 是夹具自己的 `Debug.Print`（`ucTreeViewPlus.ctl:985`，
缺 tooltip 图片，VB6 同样会打），非缺陷。

**残差（不影响启动）**：stderr 仍有 1 条
`[C3_HOST_MEMBER_MISS #1] site=SetProp name="Font" obj=… isWin=1 cls=Edit parent=VB6_UserControlHost`
—— 宿主成员分派对原生 Edit 的 `Font` 赋值还没有应答臂；像素对照前要确认它是否影响文字度量
（下一轮先查 `vb6_uc_hostmodel_setprop` 的 Font 臂）。

## 11. 剩余 2 条 warning 的口径

```
ucTreeViewPlus.ctl(1234,17): warning VB3003: If条件应为Boolean, 实际为 String
ucTreeViewPlus.ctl(1347,5): warning VB3001: 未声明的标识符: 'Cls' (可能来自其他模块)
```

- 均**非致命**，VB6 也接受（`Cls` 的代码生成已正确走 `vb6_UserControl_Cls`，只是诊断没消）；
  与 `SmoothScrollFactor` 的 `无效值` 告警同档：**只告警、不失败**。不要在判红里算它们。
- 待做（不阻塞启动）：与真 VB6 的同一处告警对照，确认 `VB3003` 在 VB6 里是否同样存在。

## 12. 下一轮从哪继续（按顺序）

1. **`[C3_HOST_MEMBER_MISS] SetProp Font (cls=Edit)`**：查 `vb6_uc_hostmodel_setprop` 的 `Font` 臂，
   补上对原生 Edit 的字体应答（没有就如实记录为残差，先别猜着改）。
2. 小步提交（每类一次，推 `github`/`origin` 的 `dev`，快进不 force，跑 GA）：
   - 提交 A：`tests\ucTreeView\` 夹具 + 本 handoff；
   - 提交 B：生成器（UC 方法桥三处、`prop_let_/prop_set_` 动词、`VBA.Left$` 兜底、裸 `Cls`）；
   - 提交 C：`vb6_di_*` 桩批量补齐（纯桩文件）；
   - 提交 D：本第 10 节 RTL 修复（`vb6com_pack.c` + `uc_collection_api.inc`，**已**重建 `c3rtl.res`）。
3. 真 VB6 1:1 截图（`VB6.EXE /make` + `CopyFromScreen`）→ 对照窗口/客户区、控件 rect（twips/15）、关键像素。
4. 写 `tests\guisample\uctreeview_px.ps1` + 登记 `tests\run_tests.ps1` + `tests\emit-manifest.expected.txt`（现算，禁 `-Bless`）。
5. 交付 `output\ucTreeView\` 与对照表，收口回调主控。

## 13. 本机陷阱（本轮新增实测）

- `powershell -Command "$p=..."` 外层会吞 `$p` ⇒ 直接 `.\.temp\build_c3.bat`。
- `Start-Process -RedirectStandardError` 的文件在进程启动瞬间就创建，**"文件存在 ≠ 跑完"**；必须 `-Wait` 或同命令串行。
- 成功编译后 session 目录被 `cleanupOldSessions` 删掉，只能读**失败** session；`c3-error.log` 成功时仍是旧内容，
  以退出码 + exe 时间戳为准。
- 本机只有 `157615014017600_7680` 这一个失败 session 目录（22:59:35），生成码行号可能比当前生成器晚一版，
  但 `RecurCalcTreeWidth` 等函数体未受本轮生成器修复影响。
- map 符号化：`Rva+Base − 0x400000` 才是与 VEH 打印 `rva` 同口径的偏移；`0001:` 段偏移会差 0x1000。
