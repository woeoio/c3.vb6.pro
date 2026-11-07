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

**回归与定点化（GA 首轮 22→21 后修）**：全局在 `vb6_ComUnpackObject` 里解包实例打断了
`k174_vieweach_station_a`（`K174-sum=303/skipped=2` → `0/4`：接口视图分支必须拿到
**IDispatch 包装器**才能问 `ClassNameA` 选实现类）。最终形态：
- `vb6_ComUnpackObject` **保持原语义**（只做宿主包装器解包，交 IDispatch）；
- 新增 `vb6_ComObject_UnboxVb6Instance()`（`vb6comserver_obj.c`，与 `IsBoxed` 同闸；
  非包装器/未绑实例原样交回）；
- 只在 `cgen_control.cpp` 的 **For Each 项目类**分支（`knownClassVars_`）发射
  `void* _vb6_fe_box_N = vb6_ComUnpackObject(&v); x = (vb6_cls_X*)vb6_ComObject_UnboxVb6Instance(_vb6_fe_box_N);`。
- 本地实测（23:54）：`K174-sum=303 / K174-skipped=2 / K174-DONE` ✓；
  ucTreeView 重编后仍 `ALIVE title=Form1`、无 `[C3_CRASH]`/`[C3_HOST_438]` ✓。

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

## 14. GA 记录与下一轮第一条（2026-11-06 00:05）

**run 42059617000 (dev=daa595cd) = FAIL**，逐格：
- `Emit manifest (shape oracle)`：新夹具 `tests/ucTreeView/Proyecto1.vbp` 未登记哈希 +
  存量 **15 份**发码哈希变（本轮生成器改动所致：UC 方法桥 / Property 动词 / VBA.Left$ / 裸 Cls /
  For Each 项目类）。处置：先不带 `-Bless` 跑 `scripts/rebless_emit_manifest.ps1` 看逐行 diff，
  确认形状再登记（**不许拿 -Bless 当通过**）。
- `Tests (compile #1)`：`emitted_declared` 红 = `vb6_FormGetImagePicture` 有 RTL 定义、伞头无原型
  → **已在 03c89a73 补进 `vb6rtl_class_com.h`**。另 `host_pseudo_table_census` / `uc_instance_exit` /
  `object_slot_boxing` 红，待逐条看差异（多为预期形状变更，必须核对后再同步断言）。
- `Tests (compile #2)`：`vbp_fixture_census` / `manifest_coverage` 红 = 新夹具未登记
  `tests\run_tests.ps1` 的 vbp/GUI 表 + 清单；`com_latebind_guard` / `variant_slot_store` 红，待查。
- `Tests (vbp #8)`：`czUI` FAIL（`exit=-1073740771`，`latebind-miss 期望 3 实到 0`）。
  根因：Collection 的 AddRef/Release 对**宿主包装器**也计数 ⇒ 双份寿命。
  **已在 03c89a73 收窄为 `vb6_ComObject_IsBoxed` 为真才计数**（待 GA 复验）。
- `Tests (bas #2)`：`test_declare_byref_string_out` / `test_declare_byval_string_out`
  `output mismatch` —— **下一轮第一条**要查（先本地跑该两条 bas，逐字对比期望输出，
  再决定是本轮哪一处生成/RTL 改动引起）。

本地复验（03c89a73 之前一次构建，00:02）：`K174-sum=303 / skipped=2 / DONE` ✓；
`output\ucTreeView\Proyecto1.exe` `ALIVE title=Form1` ✓。

**下一轮执行顺序**：
1. 本地跑 bas 两条 string-out 用例，定位 output mismatch；
2. 逐格读 compile #1/#2 的静态门禁 diff（`host_pseudo_table_census`、`uc_instance_exit`、
   `object_slot_boxing`、`com_latebind_guard`、`variant_slot_store`、`vbp_fixture_census`、
   `manifest_coverage`），区分"预期形状变更"与"真回归"；
3. 登记 `tests\run_tests.ps1` 的 ucTreeView 条目 + `emit-manifest.expected.txt`（逐行现算）；
4. 写 `tests\guisample\uctreeview_px.ps1`；真 VB6 截图对照；交付。

## 15. 续跑第二轮：di 桩回退收口 + bas 出参修复（2026-11-06 00:30）

**A. `src/rtl/core/di` 回退 + 19 条桩增量补回**
- 回退：`git checkout 841c1d3c^ -- src/rtl/core/di`（841c1d3c 的生成器整篇改写把既有桩改坏，
  实测 czUI `av EXECUTE(DEP) target=0xffff0559`，帧 `vb6_di_SetWindowRgn`）。回退后 czUI
  `ALIVE title=czForm Demo`。
- 回退的副作用（本轮实测复现）：ucTreeView 链接 19 条 `LNK2019`。原型**不是猜的** —— 从失败
  session `C:\Users\ADMINI~1\AppData\Local\Temp\C3C\162810628553000_12952\*.h` 逐条取 C3 生成头
  的 `extern ... vb6_di_<name>(...)`，按 `841c1d3c` 的原域放置、**只追加不改既有**：
  - `gdiplus/vb6_di_gdiplus_path_stubs.c`：`GdipAddPathPolygon` / `GdipGetPathWorldBoundsI` / `GdipTransformPath`
  - `gdiplus/vb6_di_gdiplus_image_stubs.c`：`GdipGetImagePixelFormat` / `GdipGetImageDimension`
  - `gdiplus/vb6_di_gdiplus_draw_stubs.c`：`GdipCreateMatrix` / `GdipTranslateMatrix` / `GdipRotateMatrix`
  - `gdiplus/vb6_di_gdiplus_brush_stubs.c`：`GdipSetPenDashStyle`
  - `vb6_di_user32_stubs.c`：`SetRectEmpty` / `GetCurrentThemeName` / `GetThemeDocumentationProperty` / `InitCommonControls`
  - `vb6_di_shell_stubs.c`：`SHGetFileInfoA` / `SHGetSpecialFolderLocation` / `PathIsDirectoryA`
  - `vb6_di_win32_stubs.c`：`FindFirstFileA` / `FindNextFileA` / `GetLogicalDriveStringsA`
- 重编实测（00:28）：`c3rtl.rc.res` 2889120 B @ 0:28:56、`C3.exe` 15388160 B @ 0:28:57；
  内嵌符号逐条核实为 True（`vb6_di_GdipAddPathPolygon` / `GdipSetPenDashStyle` / `SetRectEmpty` /
  `GetCurrentThemeName` / `SHGetFileInfoA` / `FindFirstFileA` / `InitCommonControls` /
  `vb6_ComObject_UnboxVb6Instance`）。
- ucTreeView：`C3 --arch x86` 退出码 0（无 LNK2019），启动 **`ALIVE title=Form1 hwnd=2883614`** ✓。

**B. `Tests (bas #2)` 两条 Declare String 出参**
- 根因已定位（见第 4 节第 4 条）：`cgen_expr_call_arg_emit.inc` 的 `isAssignableOutTarget161b`
  裸标识符臂 `if (!is) return false;` 误伤局部变量（不在全局 `symTab_`）。
- 修：`if (!is) return true;`。重编后本地实测（00:24/00:25）：
  `len=14 / name=[Administrator] / byref-name-ok=Y`、`byval-name-ok=Y` ✓。

**下一轮第一条**：逐格读 GA 静态门禁 diff（第 14 节第 2 条那 7 个），再登记
`tests\run_tests.ps1` 与 `emit-manifest.expected.txt`；随后写 `uctreeview_px.ps1` 与真 VB6 截图对照。

## 16. 第三轮：目标类（Extender 成员 + FontBold/FontUnderline）现状核实（2026-11-06 00:35）

本轮指令要求清 `Extender.Parent/Top/Left/ZOrder/Name` 结构体缺成员 + RTL 缺
`FontBold/FontUnderline` 共 14 条。**实测该类的 cl 错误数已经是 0** —— 这类在前几轮
（5e428b2b 生成器收口 + 03c89a73/ffc3fe59 RTL）已经清掉，不是剩余项。

实测（本轮 00:34，`C3 --arch x86` 全量重编 ucTreeView）：
- cl 错误条数 **0 → 0**，退出码 0；产物 `.temp\o_uctv_check\Proyecto1.exe` 生成。
- 仅剩 2 条非致命 warning（VB3003 / VB3001，见第 11 节），按口径不算失败。
- 因此本轮无该类代码改动、未重建 RTL（`c3rtl.rc.res` 仍是 00:28:56 那版，新符号已核对过）。

当前真实状态（供下一轮直接接手）：
- 已修并已提交：编译器 cl 类全部（含 Extender 成员 / FontBold / FontUnderline / 字符串
  字面量反斜杠），LNK2019 类全部，启动期 AV 两层，czUI DEP，bas Declare 出参。
  dev 顶端 = `8cedb604`（+ 本轮 docs 提交）。
- 未提交的本地改动：`tests\run_tests.ps1` 已加 `Test-VbpBuild "ucTreeView" ... -Arch "x86"`
  （vbp_fixture_census 实测已因此转绿），但 `emit-manifest.expected.txt` 仍缺
  `tests/ucTreeView/Proyecto1.vbp` 那一行 ⇒ manifest_coverage 仍红。
- 剩余待办（按类）：① 静态门禁 7 格 diff 分类（第 14 节第 2 条）；② emit-manifest 逐行
  现算登记（本地全量已起算，输出在 `.temp\emit-manifest.new.txt`，禁 `-Bless`）；
  ③ `[C3_HOST_MEMBER_MISS] SetProp Font (cls=Edit)` 影响面核实；④ `uctreeview_px.ps1` +
  真 VB6 截图对照；⑤ 交付与回调。

**下一轮第一条**：`tests\run_tests.ps1` 的 ucTreeView 登记 + `emit-manifest.expected.txt` 新行
（同一类：登记），提交推 GA；之后再按类读静态门禁 diff。

## 17. 第四轮：夹具登记 + 首份测量读数（2026-11-06 01:05）

**A. 登记（已收口）**
- `tests\run_tests.ps1:9305` 新增 `Test-VbpBuild "ucTreeView" ... -Arch "x86"`；语法 ParseFile 通过。
- `emit-manifest.expected.txt`：新增 `tests/ucTreeView/Proyecto1.vbp`
  （sha256=172f9c3a… ascii256=e4470177… rc=0 bytes=923835，本机 `.build\C3.exe` 现算，
  与 CI artifact 同值）。
- 本机全量现算**不可用**（单输入 >1 分钟，两个并发作业互相拖慢，已终止）；改用既定先例：
  CI run `42065548065`（head 8cedb604）的 `emit-manifest` artifact 登记 11 行存量。
  其中 5 行（czFormDemo / VBFlexGridDemo / Units / B445 / SegmentedControl）ascii256 变，
  6 行 `Charts 2020` 家族（带空格路径，正则 `(\S+)$` 会截断，必须用 `(.+)$`）。
- 本地门禁实测：`compare_emit_manifest` **595/595 逐行相同**；`check_manifest_coverage`
  `inputs=595 registered=595 集合相等` OK；`check_vbp_fixture_census` `220/214+6基线` OK。

**B. 首份 1:1 测量（只测不改）**
- C3 产物：`output\ucTreeView\Proyecto1.exe`，820224 B，PE `machine=0x014C`，
  进程 Wow64=True（32 位），主窗 `Form1`。
- 窗口/客户区：C3 `767x498` / 客户端 `752x460`；VB6 `767x498` / `752x460` —— **完全一致**。
- 子窗相对窗口原点（两边逐项相同）：UC 宿主 `(15,38) 177x445`；`AllowEdit (215,206) 161x17`；
  `Contact List (215,232) 217x177`；按钮/复选框等坐标两边一致。
- 像素：整窗差异 61.6%；树区域（rel 15,38 177x445）差异显著 —— C3 主色 `240,240,240`
  与 `38,38,38` 文字，缺 VB6 的 `255,255,255` 底与 `255,180,180` 选中高亮（1323 采样点）。
- 肉眼对照：VB6 左树有节点（Facebook/Meta/Instagram/WhatsApp）与图标；**C3 左树区域空白**。
- 运行日志实测（C3）：`THE IMAGE: 1 DON'T EXIST`（另有数条）—— 图像列表加载失败；
  另有已知 `[C3_HOST_MEMBER_MISS] site=SetProp name="Font" obj=… cls=Edit`。

**下一轮第一条**：查 `THE IMAGE: N DON'T EXIST` 的加载路径（App.Path/img 与 LoadPicture
调用点），确认节点不绘制是否与之同因；再查树区背景色/选中高亮的画刷来源。

## 18. 第五轮：先清 GA 红（compile #1/#2 五条静态哨兵），2026-11-06 04:3x

主控调整优先级：树区功能修复之前，先把 GA 红清零。三个 run 的结论：

| run | head | 结论 | 红在哪 |
|---|---|---|---|
| `42065548065` | 8cedb604 | failure | compile #1/#2 + Emit manifest（另有当时的夹具/manifest 未登记） |
| `42066368425` | 92892724 | failure | compile #1/#2 + Emit manifest |
| `42068532400` | 82907ac3 | failure | compile #1/#2（Emit manifest 已绿） |

五条红的归属 = **本轮前两刀（5e428b2b / 841c1d3c）自己引入的真实回退**，不是判据过紧：

1. `host_pseudo_table_census` — `cgen_expr_ident_builtin.inc` 问装配出口 2 次（要 1）。
   修：`hpRtl` 提前算一次，裸名/裸写方法两档共用。
2. `uc_instance_exit` U1 — `cgen_expr_call_arg_emit.inc` 手拼 `vb6_UC_InstanceOf`。
   修：新出口 `CCodeGen::ucInstanceThisPtrExpr`（住 `cgen_util_classcall.cpp`），
   expr/ 里 0 处手拼。
3. `object_slot_boxing` B7 — `src/rtl` 的 `vb6_ComObject_IsBoxed` 9 处（要 4）。
   修：低层身份闸拆名 `vb6_ComObject_IsVb6Wrapper`；类名 BSTR/ANSI 两档仍只叫
   `IsBoxed`，集合计数改叫新名字。计数回 4。
4. `com_latebind_guard` R9 — `vb6rtl_com.c` 又出现八档→LOGFONT 就地抄写（`lfWeight`=1）。
   修：叠在 `vb6_ComIface_Font` 副本上，再交唯一主人 `vb6_UC_FontLogFont`；`lfWeight`=0。
5. `variant_slot_store` R5d/R9c —
   - `src/backend/module/cgen_com.cpp:10` 注释里引用了 `class var .` 原句（R5d 按先例
     改写成不点名说法，白名单条目保持 0）；
   - `cgen_expr_member_voidptr_com.inc` 新增一档成员读发文站（调 `classFieldReadExpr`，
     不是手拼尾注）—— R9c 白名单由三处登记为四处（**登记同步**，非放宽）。

**重编与内嵌**：`.temp\build_c3.bat` 通过；`c3rtl.rc.res` 2890196 B @ 04:02:50；
`check_rtl_embedded`：126 个 RTL 资源与磁盘逐字节一致、无孤儿 id；
`vb6_ComObject_IsVb6Wrapper` 已出现在 exe 字符串。
**定向编译**：`C3 --arch x86 tests\ucTreeView\Proyecto1.vbp` 成功，`Proyecto1.exe` 820224 B。
**提交**：`fc7e38f6`（双远端 dev）。
**GA**：`42100365008`（head fc7e38f6）**22/22 success，0 failure，0 skipped**。

**下一轮第一条（回到主线）**：编译并运行 `.temp\probe_coll.bas`（Collection 的 Variant
索引取值），再用 `SetForegroundWindow` 重拍 `output\ucTreeView` 截图，分离
「图像失败 → 节点不绘制」与第二条原因。

## 19. 树区空白 = 两条独立因果（2026-11-06，已修完并 GA 22/22）

### 19.1 分离验证的两个结果（都不是猜的）

- **结果 A「有节点无图」**：没有出现。把 `ImageListDraw` 在函数头直接
  `ImageListDraw = True : Exit Function` 短路（其余不动），SetForegroundWindow +
  SetWindowPos(TOPMOST) + CopyFromScreen 重拍：树区（UC 宿主）**仍全 240 灰**，
  `white=0 bg240=77703/77703 dark=0 darkRows=0 hl=0` ⇒ 节点也没画。
- **结果 B「仍然无节点」**：成立。⇒ 修图标不可能让树出图；必须另找一条。

### 19.2 因果链一：`cImageList(index)` 返 0（图标 / 「THE IMAGE: n DON'T EXIST」）

插桩副本（`.temp\ucview_probe`，只动副本）输出（`.temp\probe4b.out`）：

```
IMG_ADD path=...\out_probe4\img\facebook.png h=49357968 size=20x20 vt=8
IMG_ADDED uc=16356176 coll=16328896 count=1 err=0
IMG_CHK now=0 addrh=15724776          <-- 同一帧、hImage 还活着，取回已经是 0
...
ILD1 uc=16356176 coll=16328896 err=0  <-- 同一个 UC 实例
ILD2 count=3 err=0                    <-- 集合里确有 3 项
ILD3 idx=1 vt=3 err=0                 <-- 索引是 VT_I4 = 1
ILD4 img=0 err=0 desc=
THE IMAGE: 1 DON'T EXIST
```

⇒ 不是实例不一致、不是 key/索引语义、不是 Add 没成功、也没有被 `On Error Resume Next`
吞掉任何错误。最小复现（`.temp\probe_coll2.bas`，实测
`.temp\probe_coll2c.out`）：

```
BYREF_PARAM=111     ' c.Add p   (形参, cgen 发 vb6_ComPackInt 按值)  -> 取回 111
LOCAL_VAR=0         ' c.Add t   (局部变量, cgen 发 vb6_ComPackIntRef 按引用) -> 取回 0
```

根因：`ImageListAddImage` 里的 `hImage` 是**局部变量**，cgen 按 VB6 默认 ByRef 发
`vb6_ComPackIntRef(&hImage)`（`VT_BYREF|VT_I4`，形状本身没错），而
`uc_hostmodel.c` 的 `vb6_Host_FromWinVariant` 把 `VT_BYREF` **刻意折成 Empty**
（原注释「归 §B144 的另一问」）⇒ 集合里存的是空值，取回恒 0。
修法：`uc_collection_api.inc` 在入库前先 `VariantCopyInd` 解引用（VB6 的 Collection
本来就是存值拷贝），`vb6_Host_FromWinVariant` 自身口径不动。

### 19.3 因果链二：`hdc = 0`（真正的「树区空白」）

在 `Update`/`Draw`/`RecursiveDraw` 插桩后（`.temp\probe5.out`）：

```
UPD font=NOTHING (exit)      ' Initialize 第一帧, m_Font 还没 ReadProps
DRAW enter hdc=0 nodes=0 ... ' Resize
UC_SHOW
DRAW enter hdc=0 nodes=0 ...
IMG_ADD ...                  ' Form_Load 加图
UPD enter
DRAW enter hdc=0 nodes=3 ... ' 节点确实有 3 个, Draw 真的跑了, 但 hdc=0
```

⇒ 节点构建没问题（`Nodes=3`、`RD childs=3`），是 **`UserControl.hdc` 为 NULL**：
GDI+ `GdipCreateFromHDC 0` 画进空 DC，`DrawText` 同样落空，屏幕上只剩宿主铺的 240。

根因：`.ctl` 头部 `AutoRedraw = -1  'True` 是**设计期**属性，VB6 在创建 UserControl 时
就置 `AutoRedraw=True`（hDC = 常驻离屏位图）；C3 只认 VB 代码里的
`UserControl.AutoRedraw = True`，`.ctl` 头部这一行**整个被丢弃**（句柄链
`vb6_uc_push` 里 `r->hdc ? … : (adSurf.memDC ? autoRedrawDC : NULL)` ⇒ NULL）。

修改站点（4 处 + 判据）：
1. `src/rtl/core/vb6forms/vb6forms_controls.h`：`vb6_UserControlDesc` **末尾追加**
   `int32_t autoRedraw`（只许追加，保持前缀字段序一致）。
2. `src/backend/module/cgen_form.cpp`：从 `formControl.properties` 读 `AutoRedraw`
   并把新槽作为 desc 初值**最后一行**输出（注意前一行要补逗号）。
3. `src/rtl/core/vb6forms/uc/detail/uc_host_create.inc`：建窗后
   `if (desc->autoRedraw) vb6_uc_autoRedrawSet(r, 1);`（必须**当场**建表面：
   `vb6_uc_push` 只在 `adSurf.memDC` 已存在时才交 adSurf）。
4. `src/rtl/core/vb6forms/uc/detail/uc_collection_api.inc`：`vb6_coll_fromWin` 先解
   `VT_BYREF`（`VariantCopyInd`），`vb6_Collection_Add/AddKeyed` 走它。
5. `tests/guisample/uctreeview_px.ps1`（新增）+ `tests/run_tests.ps1` 的
   `Test-VbpBuild -AssetDir img -PixelAssert …`：真桌面 `CopyFromScreen`；
   判 5 条（白底占比≥15%、240 灰底≤40%、深色行≥10、`255,180,180`≥50、彩色像素≥300）。

### 19.4 改前 / 改后读数（同一台机、同一抓法、同一区域）

| 读数 | 改前（hdc=0） | 改后 |
|---|---|---|
| 树区白底 `255,255,255` | 0 | 48052（≈61% 面积） |
| 宿主灰底 `240,240,240` | 77703（100%） | 5446（7%） |
| 深色像素 / 文本行 | 0 / 0 | 15607 / 445 |
| 选中高亮 `255,180,180` | 0 | 2521 |
| 图标彩色像素（max-min≥60） | 0 | 3559 |
| `THE IMAGE: n DON'T EXIST` | 3 条 | 0 条 |

`tests/guisample/uctreeview_px.ps1` 已经两向验过：旧产物
（`.temp\out_probe4\Proyecto1.exe`，hdc=0）**FAIL rc=1**；新产物 **PASS rc=0**。

### 19.5 产物 / 提交 / GA

- `C3 --arch x86 tests\ucTreeView\Proyecto1.vbp --output-dir output\ucTreeView`
  ⇒ `output\ucTreeView\Proyecto1.exe` **820224 B**；`img\` 已随产物。
- `c3rtl.rc.res` 重建（2292xxx B 级），`check_rtl_embedded` 126 项逐字节一致。
- 提交：`a70c8b75`（cgen+RTL+判据）、`40a091a2`（emit-manifest 登记 11 行），双远端 dev。
- GA：`42113452221`（head 40a091a2）**22/22 success，0 failure，0 skipped**；
  前一枚 `42111911172` 只红「Emit manifest (shape oracle)」= 形状登记未同步（11 行）。

### 19.6 本轮未做、已登记的真实缺陷（别当成已修）

1. **`Debug.Print <默认成员调用>` 打印的是返回 VARIANT 的指针**：
   `.temp\probe_coll.bas` 里 `c.Add 111` 后 `Debug.Print c(1)` 印出 `7953936`
   （`(int32_t)(vb6_ComCallByDispid(...))` 直接强转指针），而 `n = c(v)` 正常。
   生成码侧 `Debug.Print` 实参转换需要按 Variant 取值，不是本轮树区因果链。
2. `ucTreeViewPlus.ctl(1347,5): warning VB3001: 未声明的标识符: 'Cls'`：
   `UserControl.Cls` 未解析。当前 Draw 自己会填满背景，单帧看不出问题，
   但连续重绘可能残留旧像素 —— 与 1:1 目标仍有差，待单独查。
3. `ucTreeViewPlus.ctl` 里 `youtube/mensajero/tik-tok` 三张图两套 `img` 目录都真缺，
   真 VB6 同样取不到 ⇒ 不是缺陷，别去补假图。

## 20. 设计期背景色这一族（2026-11-06 本轮）

### 20.1 口径更正（重要，别再按旧结论查）

- 之前记的「**Label 白底 255**」实测取样点其实是 **CheckBox** 的控件矩形
  (顶排 `AlternateNodeColors/ShowTreeLines/ShowPlusMinus/CheckBoxes/FullRowSelect/
  PlusMinusStyle/PlusMinusAlignRight/AllowEdit`，x≈269 起、y≈48 起、行距 30)。
  真 Label (`BorderRadius…MarginBottom`/`SelectionColor`) 在这份夹具里**本来就是 240**。
- 所以残差 #2 的真身是「**CheckBox 白底**」：VB6 里 CheckBox 默认 BackColor =
  `&H8000000F&`(按钮面 240) 且不透明；C3 落成真 BUTTON 窗口后向父窗要底色时走的
  是 **WM_CTLCOLORSTATIC**（不是 `WM_CTLCOLORBTN`，后者只对命令按钮/owner-draw 生效）。

### 20.2 三条根因与修改站点

1. **9 枚色块（Picture2.BackColor）根本没发**：`emitDesignerStyleProps` 的设计期
   ForeColor/BackColor 只在**顶层控件创建路**发（`cgen_form_ctrl_style_apply.inc`），
   容器（Frame/PictureBox/SSTab）里的子控件那条路（`cgen_form_frame_menu.inc`）一条不发。
   Picture2 全在 Picture1 里 ⇒ 9 个 BackColor(`&H00FFE8CC&` 等) 全丢。
   修：在 frame_menu 那条路补 `vb6_SetControlBackColor/ForeColor`，口径与顶层逐条相同
   （只发显式写过且类型为 Integer 的；`0x8000000n` 系统色原样传负值，RTL 再 GetSysColor）。
2. **Label 默认底色**：VB6 轻量 Label 默认 `&H8000000F&`(240)，C3 落成 STATIC 窗口后
   类刷是 `COLOR_WINDOW`(255)。修：`vb6_CreateControl` 里对 `STATIC` 且
   `(style & 0x1F) != SS_BITMAP(0x0E)` 的控件在创建时写 `VB6_BackColor = COLOR_BTNFACE`
   （PictureBox/Image 走 `SS_BITMAP` 自绘，排除）。
3. **PictureBox 容器不接子控件的 CTLCOLOR**：Label/CheckBox 在 Picture1/Picture3 里时
   `WM_CTLCOLORSTATIC` 发给 PictureBox（同样是 STATIC 类），`vb6_ImageSubclassProc`
   此前不接 ⇒ 转给原 STATIC WndProc → 白。修：加 `WM_CTLCOLORSTATIC/EDIT` 一臂，
   统一交 `vb6_ApplyCtlColorStatic`。
4. 放行判据从"生成侧看 `VB6_BackColorSet` 哨兵"改成"**RTL 判定**"：
   `vb6_ApplyCtlColorStatic` 在无哨兵时也接管 `Button` 类的 6 个复选 style
   (2/3/4/5/6/9)，默认取 `vb6_GetControlBackColor`(=BTNFACE 240)。
   生成侧三处（主窗体 WndProc、PictureBox 子类、sstab 子类）都改成无条件调，
   由 RTL 返回 0 表示"不接管"。

### 20.3 实测读数（125% DPI，959×623 整窗，VB6 实拍为基准）

- 9 格中心色逐格一致（C3 = VB6 = `255,255,255 / 180,180,180 / 99,99,255 / 0,0,0 /
  99,99,255 / 204,204,204 / 0,0,0 / 204,204,204 / 204,232,255`）。
- CheckBox 区 240 占比 **83.4%**、白 0%（改前 0% / 83.4% 白）。
- Label 区 240 占比 **90.9%**、白 0%（改前同值；本来就没坏）。
- 整窗 diff：**12.37% → 7.74%**（46258/597457）。分区：左控件列 26.08→12.4%、
  右色块/面板 19.18→10.59%、左边框 11.25→6.29%、右条 5.31→2.22%；
  树区 6.23%、底部按钮 12.79% 未动；底条 16.06→27.84% 属 DWM 圆角/阴影随屏幕背景变化。

### 20.4 判据 / 构建 / 提交

- `tests\guisample\uctreeview_px.ps1` 新增：9 格中心 RGB（±3）；8 枚 CheckBox 与
  7 条 Label 的 240/纯白占比；取样点全部来自运行时 `EnumChildWindows` 真实控件矩形
  （不写死像素坐标，96/125 DPI 都成立）。旧产物必 FAIL（已用 `.temp\old_uctv\Proyecto1.exe` 验过）。
- C3 重建后 `c3rtl.rc.res` 2894996 B、`C3.exe` 15449088 B；产物 820224 B、`machine=0x014C`。
- 陷阱：`.inc` 改动**不会**触发 `cgen_form.cpp` 重编（依赖跟踪漏），改了
  `detail/module/*.inc` 必须 touch `src\backend\module\cgen_form.cpp` 再 build，
  否则生成码还是旧的（本轮踩过一次）。

## 21. 残差 #3/#4：选中行文字色/白底 + 行高偏小（2026-11-06 本轮）

### 21.1 根因 1：`Set UserControl.ForeColor` 落不到绘制的 DC + AutoRedraw 表面默认 OPAQUE

- `.ctl:1646-1693` 每段 DrawText 前都 `UserControl.ForeColor = N.TextColor`（选中行 16777215 白、
  SubText 13421772 灰、普通行 0）。但这些 Draw/度量全发生在**事件**里（不是 WM_PAINT），
  目标 DC 是 `UserControl.hDC` = AutoRedraw 的持久 DIB 表面 (`r->adSurf.memDC`)。
- 改前 `vb6_UC_SetForeColor` 只写 `vb6_uc_activePaintDC`，而事件里该变量恒 NULL
  （取证：6 次 `[T1879 fg-nodc] ole=16777215/13421772/0`，SetForeColor 全丢）；
  `vb6_uc_applyFontNow` 同样 `if (!hdc) return`，`Set UserControl.Font` 也全丢。
- 新建 DIB DC 的 GDI 默认是 `BkMode=OPAQUE(2)` + 白 BKCOLOR + 黑 TEXTCOLOR ⇒ Tahoma 文字
  被画成**黑字+白底块**：选中粉底内实测 pureblack=224、纯白 968..1410（VB6 参考：白字、
  纯黑 ~10、白只有笔画 ~258）。
- 修：`vb6_uc_adEnsure` 建表面时补 `SetBkMode(memDC, TRANSPARENT)`、
  `SetTextColor(memDC, resolve(ForeColor))`、选入 UC 自己的字体；
  `vb6_UC_SetForeColor` 回落到 `vb6_UserControl_hDC`；`vb6_uc_applyFontNow` 两档都落
  （paint 用 `r->hdc`，事件/度量用 `adSurf.memDC`）；`vb6_uc_selectOwnFont` 改为 UC 字体优先。

### 21.2 根因 2：字体块没进 UC 的 PropertyBag ⇒ `m_SubTextFont` 被默认值覆盖成 8.25

- 生成侧（`cgen_form_create_controls.inc`）对 `BeginProperty Font/SubTextFont` 是
  **prop_let 直接回放**，而 `UserControl_ReadProperties` 收到的 bag 只装了标量属性
  （CheckBoxes/MinNodeHeight/... 7 项），**没有 Font/SubTextFont**。
- `.ctl:6992/7006`（生成码）`Set m_Font = .ReadProperty("Font", Ambient.Font)` /
  `Set m_SubTextFont = .ReadProperty("SubTextFont", Ambient.Font)` 缺项 ⇒ 都返回**默认值**
  `vb6_Ambient_Font` ⇒ 前面回放的 9pt 被覆盖成 8.25，且两枚字段指向**同一个对象**。
- 取证（改前 `[T1879 calc]` 6 次全为）：`lfH=-13 face=Arial Narrow tmH=16 extL=0`
  （= 8.25pt@120，子文本也是它）⇒ DT_CALCRECT 高 16，N.Height=16+16+6=38。
- 修：字体块也进 bag —— 新增 `vb6_UC_BagPutFont`（`vb6_ComPackObject` 包成 VT_DISPATCH，
  `ReadProperty` 再解包回八档指针）；emit 侧对 `propertyBlocks` 里名字以 `font` 结尾、
  且 `ReadProperties` 真读过 `.ReadProperty("<name>"` 的块，补 `vb6_UC_BagPutFont`。
- 取证（改后）：主文本 `lfH=-13 tmH=16`、子文本 `lfH=-15 tmH=20`（均 Arial Narrow @120，
  extL=0）⇒ N.Height=16+20+6=42。

### 21.3 改前/改后读数（125% DPI，959×623，同一夹具）

| 项 | 改前 C3 | 改后 C3 | VB6 参考 |
|---|---|---|---|
| 选中高亮 bbox | 206×35 (y56..90) | **206×39 (y56..94)** | 205×41 (y56..96) |
| 选中行纯白像素 | 968（白底块） | **355（只剩白字笔画+复选框）** | ~258（白字） |
| 选中行纯黑像素 | 224 | **12** | ~10 |
| 行距（行 2→3） | 38 | **42** | 44 |
| 子文本 LOGFONT | 无（与主文本同 8.25） | **9pt → lfH -15 / tmH 20** | - |

整窗 diff（对上一轮产物）：**3.80%**（22702/597457）。

### 21.4 判据（`tests\guisample\uctreeview_px.ps1` 新增）

- 粉底上纯白 ≤ 600（改前 968 红 / 改后 355 绿）；选中行纯黑 ≤ 30（224 红 / 12 绿）；
  选中行白字笔画 ≥ 100。
  （"深色≈0" 只用**纯黑**判：选中行里的 Facebook 图标本身是深蓝 lum≈101，用 lum 阈值会假红。）
- 高亮 bbox 高度 = 41*(dpi/120) ±2；行距 = 44*(dpi/120) ±2（行距按 `(r+g+b)<450` 的
  墨迹带起点算，避开图标/复选框切带）。
- 旧产物 `.temp\old_uctv\Proyecto1.exe` 已验 **FAIL**（4 条新判据全红）。

### 21.5 其它

- `.temp` 里的 `C3_T1879` 取证代码已全部移除（`uc_host_window.c` dump/打印、
  `vb6_di_user32_stubs.c` 的 include 与 DT_CALCRECT 分支）；`C3.exe` 内嵌 RTL 已核对
  含 `vb6_UC_BagPutFont`、无 `C3_T1879`，`scripts\check_rtl_embedded.ps1` 126/126 一致。
- 提醒仍然有效：改 `detail/module/*.inc` 必须 touch `src\backend\module\cgen_form.cpp`
  再 build（本轮又踩了一次）。

### 21.6 收口：`vb6_uc_selectOwnFont` 只能用窗口那一格 + 取证清理（2026-11-06）

- **教训（写死）**：`vb6_uc_selectOwnFont` **不能**优先 `vb6_ucCurrent->font`。
  非 AutoRedraw UC 的临时 DIB 是在 `vb6_uc_push` **之前**选字体，此刻
  `vb6_ucCurrent` 还停在上一个 UC（或更早实例的陈旧字体）⇒ ucSegmentedControl 的
  CheckBar 拿到 MS Sans Serif，PUA 字形（U+E16C..）整片不画
  （`segmented_px.ps1` 实测 `uc6 darkPx=54 < 400`）。回退为单一来源
  `vb6_ControlFontStored(hwnd)` 后恢复到 581。逐段 `Set UserControl.Font` 仍由
  `vb6_uc_applyFontNow` 落到本实例的绘制 DC（那里 hdc 已经指向正确的 DC）。
- 同时删掉 `vb6_uc_adEnsure` 里多出来的第三次 `vb6_uc_selectOwnFont` 调用：
  静态哨兵 `scripts\check_control_dc.ps1` 钉 1 定义 + 2 调用（对外发 DC 的
  `vb6_uc_autoRedrawDC` + 非 AutoRedraw 的 `vb6_uc_paintInto` 那档）。
- 取证清理：`C3_TMP_SF`（`vb6_UC_SetFont` 里的字体指针打印）与 `C3_TMP_DT`
  （`vb6_di_DrawTextW` 里的 GetTextColor/GetBkMode/LOGFONT/TEXTMETRIC 打印）全部删除；
  重编后 `findstr` 核对 `.build\C3.exe`：`vb6_UC_BagPutFont` 在、`C3_TMP_` 无。
- **本机残红（非本轮回归，已 A/B）**：`悬停 uc1 禁用项的光标 IDC_NO (65539 vs 65559)`。
  用上一轮 `publish\C3.exe` 编同一夹具（`.temp\seg_pre\`）复现**同一读数**，
  且该条在上一轮 GA（96dpi runner）上是绿的 ⇒ 与本轮改动无关，属本机
  125% DPI 光标主题句柄差异；CI 面按 96dpi 判定。
- 提交：`a4d4ae6d`（selectOwnFont 回退 + 哨兵 + 取证清理）、`1876366c`
  （登记 7 份 UC 工程发码形状：字体块进 bag）。
- **GA**：run `42140484128`（head `1876366c`）= **22/22 全绿**；此前 run
  `42136304681` 的三红（Emit manifest / `[STATIC] control_dc` 调用 3 处 /
  `ucSegmentedControl` uc6 darkPx=54）三条都已消除。


## 22. 残差 #5/#6：TextBox 3D 内凹 + VScrollBar 自绘（2026-11-06）

### 22.1 残差 #5 根因：`WS_BORDER` 换成了主题的 `EDGE_ETCHED` 深框

- VB6 `ThunderRT6TextBox` 实测（`GetWindowLong`）：
  `style=0x540108C1`（**无** `WS_BORDER`/`WS_EX_CLIENTEDGE` 之外的边位）、`ex=0x204`
  （`WS_EX_CLIENTEDGE`(0x200) | `WS_EX_NOPARENTNOTIFY`(0x4)）；
  `GetClientRect`=87x29，即**每边 2px** 内凹。
  取样像素：左/上 = 236(1px) + 254(1px)，下/右 = 254(1px) + **131**(1px)，内部 255。
- C3 改前：`style=0x50810080`（含 `WS_BORDER` 0x00800000）、`ex=0`；主题 EDIT 把
  `WS_BORDER` 画成 1px `COLOR_WINDOWFRAME = 0x646464`(100,100,100) ⇒ 91x33 周长 244px
  深框、顶行整行深色（旧产物实测 `edFrameDark=1464` / `edTopDark=546`，6 枚）。
- 修法：**两半**。① `cgen_form_ctrl_style_apply.inc` 的 TextBox 去 `kWsBorder`；
  ② `cgen_util_ctrl.cpp::controlContainerExStyleBit` 给 TextBox 返回
  `WS_EX_CLIENTEDGE=0x00000200`（顶层路与容器子控件路共用这一个出口，
  `cgen_form_frame_menu.inc:445` 也吃它）。VB6 的 `NOPARENTNOTIFY` 只影响父窗通知，不立。
- 改后实测（同一坐标、逐点比对）：左 x399=`236` / x400=`254` / x401=`255`，
  上 y298=`236` / y299=`254`，下 y329=`254` / y330=`131` —— **与 VB6 逐像素相等**。
  `edFrameDark=0` / `edTopDark=0` / `edEdge236=726`（6 枚合计，单枚 ≈121~244）。

### 22.2 残差 #6 根因（三层，缺一层都不出图）

1. **缺 `SBS_VERT`**：C3 发码给 VScrollBar 的 style 是 `0x50010000`（无方向位）；
   VB6 `ThunderRT6VScrollBar` = `0x54010001`。已在两条创建路补 `0x1`，改后 `0x50010001`。
2. **33x31 根本放不下原生箭头**：@120dpi `SM_CYVSCROLL=21`、`SM_CXVSCROLL=21`，
   竖排两根箭头要 2*21=42 > 31 ⇒ 原生 SCROLLBAR 只走 `WM_ERASEBKGND`，用类背景刷
   （`COLOR_SCROLLBAR`=白）铺满 ⇒ 1023/1023 纯白。对照实验：同一控件拉到 33x200 /
   100x31 就正常画出箭头与槽 ⇒ 不是"没实现"，是"这个尺寸下系统就不画"。
   VB6 这个尺寸也不画箭头（6 枚实测全是 `240,240,240 = 960` + `255,255,255 = 63`，
   `whiteRows=[15]`），所以 C3 也照 VB6 **自绘观感**而不是画箭头：
   `src/rtl/core/vb6forms/vb6forms_scroll.c` 新增
   `vb6_ScrollBarPaintTo`（`FillRect` 240 全底 + 左边 1px 白 + `mid=top+(bottom-top)/2` 处 1px 白；
   横向无 VB6 样本，镜像成上边 1px + 宽/2 处 1px）+ `vb6_ScrollBarSubclassProc`
  （接管 `WM_ERASEBKGND`/`WM_PAINT`/`WM_PRINTCLIENT`/`WM_PRINT`，原过程存 `VB6_SBar_OrigProc` 属性）
   + `vb6_InstallScrollBarSubclass`（`GWLP_WNDPROC` + `SetPropW` + `InvalidateRect`）。
3. **z-order 反转 + 缺 `WS_CLIPSIBLINGS`**：VB6 的 .frm 里"先列出的控件在兄弟最上"
   （`Form1.frm` Picture3 下先 `VScroll1`(85 行) 再 `Text1`(95 行)）。C3 逐枚
   `CreateWindowExW`，新窗默认插到兄弟顶部 ⇒ 顺序整体反过来，`Text1` 的白底盖住 `VScroll1`
   （决定性证据：把 Edit 隐藏后同一区立刻变成 960/63）。修法：设计期创建后
   `SetWindowPos(hwnd, HWND_BOTTOM, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE)`；
   并给每枚子控件补 `style |= WS_CLIPSIBLINGS`（VB6 每枚都有；缺位时**下层**控件
   可越过 z-order 画进上层兄弟的窗口区）。修后兄弟序
   `z[0]=ScrollBar z[1]=Edit z[2]=Static`。

### 22.3 断言（`tests/guisample/uctreeview_px.ps1`）

- TextBox：`edFrameDark(100±2) == 0`；顶行无深色（`edTopDark == 0`）；
  **每枚都有 236±2 亮边**（`edEdgeOk == 枚数`）—— 只判"深框没了"会漏掉
  "连 3D 内凹一起没有"这种假绿。
- VScrollBar：`240±1` 占比 ≥ 85%；纯白占比 3%~10%；6 枚各自"存在一整行 ≥90% 白"
  与"存在一整列 ≥90% 白"（不写死第 15 行：`GetWindowRect` 与桌面像素常差 1px）。
- 新产物：`edFrameDark=0 edTopDark=0 edEdge236=726/6 sb240=93.8% sbWhite=6.2% sbMid=6/6 sbLeft=6/6` → PASS。
- 旧产物 `.temp\old_uctv\Proyecto1.exe`：`深框 1464px > 0; 顶行深色 546px > 0;
  236 亮边缺失 (0/6, 共 0px); VScrollBar 240 占比 0% < 85%; 纯白占比 97.9% 不在 3%~10%` → FAIL。

### 22.4 当前产物 vs 真 VB6 实拍（959x623，Δ=`|ΔR|+|ΔG|+|ΔB|>=6`）

口径 `.temp\diffreg.ps1`（与上轮 `diff2png.ps1` 同一判据，首轮 ±3px 平移扫描确认
最佳偏移 =(0,0)，无整体错位）。

| 分区 | rect (x,y,w,h) | 改前 | 改后 |
|---|---|---|---|
| title      | 0,0,959,32     | 6.19% | 5.09% |
| leftEdge   | 0,32,20,591    | 46.07% | 46.03% |
| treePanel  | 20,32,232,591  | 5.41% | 5.41% |
| optChecks  | 252,32,278,258 | 8.32% | 8.32% |
| labelEdits | 252,290,278,250| 16.30% | **6.25%** |
| swatchCol  | 530,32,250,383 | 8.44% | 8.44% |
| dotsCol    | 780,32,179,383 | 5.03% | 5.03% |
| bottomBar  | 252,540,458,83 | 19.70% | 19.70% |
| rightMid   | 530,415,180,125| 0.00% | 0.00% |
| rightBlank | 710,415,249,208| 7.84% | 7.83% |
| **整窗**   | 959x623        | **9.23%** (55133px) | **8.00%** (47807px) |
| **客户区** | 10,32,940,582  | **6.24%** | **4.96%** (27114px) |

- 整窗读数受**非客户区状态**影响：同一二进制 4 次抓屏 6.13%~8.06%，差全部落在
  `title`+`leftEdge`（外框激活态 38 vs 59 灰、底部桌面渗入）；`leftEdge` 那 46%
  本质是窗口最外 9px 边框条，不是窗体内容。
- **客户区四次抓屏 4.96% / 4.96% / 4.96% / 4.96%**（27114~27138px，σ≈0.02pp）——
  这是可复现的那一半；对 VB6 的净改善 = 6.24% → 4.96%（-1.28pp / -6986px）。
- 上轮记录的 7.74% 与本轮 8.00% 同脚本同参考图，+0.26pp 即上述外框态差异，
  客户区口径不受它影响。

### 22.5 收口

- 提交：`371e677b`（两条残差的 RTL/发码/断言）；随后 `ae11dcc8` 把
  `emit-manifest.expected.txt` 重新登记 22 行（吃 CI 那台二进制回传的清单：
  `git fetch github ci/emit-manifest`，本机 3 份 guisample 的差异是**本机 OCX/typelib 态**，
  CI 清单里那 3 份与期望一致，不能拿本机输出登记）。
- GA：run `42157587845`（head `371e677b`）= 21/22（唯一红是 Emit manifest，预期）；
  run `42158448470`（head `ae11dcc8`）= **22/22 全绿**。
- 提醒仍然有效：改 `detail/module/*.inc` 必须 touch `src\backend\module\cgen_form.cpp` 再 build。

## 23. 残差 #7：底部按钮长标题不换行（BS_MULTILINE）

### 23.1 根因（实测，不是猜）

- **VB6 侧**：`ThunderRT6CommandButton` 实测 `style=0x54012000`（WS_CHILD|WS_VISIBLE|
  WS_CLIPSIBLINGS|WS_TABSTOP|**BS_MULTILINE=0x2000**）、`ex=0x4`；窗口 86x51 @120dpi，
  client 与 rect 同尺寸（无边框内缩）。caption：`Tree Explorer` / `Corporate Hierarchy` /
  `Drag and Drop` / `Menu / Tools` / `Contact  List`（双空格）。
- **C3 侧**：原生 `BUTTON`，改前 `style=0x54014000`（只有 `BS_NOTIFY=0x4000`，**无 0x2000**）、
  `ex=0`。系统在没有 BS_MULTILINE 时把 caption 强制单行绘制；有了这一位才按**空格**
  word-wrap（不是按 `\n`、也不是按控件宽硬断）。
- **断行阈值证据**（同 DPI 内存 DC，`GetTextExtentPoint32`，字体档与绘制一致）：

  | caption | VB6 字体 Arial Narrow -16 | C3 字体 MS Sans Serif -13 | VB6 行数 |
  |---|---|---|---|
  | Tree Explorer        | 82  | 82  | 2（Tree / Explorer） |
  | Corporate Hierarchy  | 119 | 121 | 2（Corporate / Hierarchy） |
  | Drag and Drop        | 87  | 89  | 2（Drag and / Drop） |
  | Menu / Tools         | 78  | 78  | 1 |
  | Contact  List        | 72  | 71  | 1 |

  按钮宽 86，系统文本区 = 80（每边 3px）⇒ 阈值落在 [78,82)：前三枚过线、后两枚不过线。
  两套字体档在该阈值下 5/5 一致 —— 所以**换行决策与字体档无关**，只差行距（见 23.2 末）。
- **系统能力确认**：不重编、只在已运行的原生产物上 `SetWindowLong(GWL_STYLE, style|0x2000)`
  并重绘，三枚立刻变两行且断行位置与 VB6 相同 ⇒ **不需要自绘 caption**，也不需要用
  owner-draw 伪造；竖直居中由系统按钮自己按多行块算。
- 改前墨迹（参考图同坐标）：C3 `80x13 / 78x13 / 80x13` 单行 @ y568-580；
  VB6 `50x34 / 57x34 / 52x34` @ y558-591（bands `[558..568][578..591]`，行距 20）；
  `Menu / Tools` 76x11、`Contact  List` 71x11 两边本来就一致。

### 23.2 修改站点

- `src/rtl/core/vb6forms/vb6forms.c::vb6_CreateControl`（唯一创建站点），
  `CreateWindowExW` 之前：`win32Class == "BUTTON"` 且 `style & 0xF` ∈
  `{BS_PUSHBUTTON(0), BS_DEFPUSHBUTTON(1)}` ⇒ `style |= 0x2000L`。
  只补 push 型：check/radio/group 的 0x2000 是另一层语义（文字/图标换行），不动。
- 改为后 C3 按钮实测 `style=0x54016000`（= 旧 `0x54014000` | 0x2000）。
- 改后墨迹：`49x29 / 58x29 / 54x29`（bands `[560..569][576..588]`），
  后两枚仍 `76x12 / 70x10` 单行 —— **行数逐枚与参考一致**。
- 行距差：VB6 = 20（`TextHeight` = Arial Narrow 9.75pt 的 tmHeight=20，来自窗体级
  `BeginProperty Font` 的 ambient 继承），C3 = 16（默认 MS Sans Serif 8.25 的 tmHeight=16）。
  即整块上沿 +2px、下沿 -3px。这是**字体继承**残差（C3 的窗体级 Font 块没落到窗体窗口，
  子控件也无从继承），与换行规则是两件事，单独记一笔。

### 23.3 断言（`tests/guisample/uctreeview_px.ps1`）

- 5 枚 push 按钮按 caption 定位，各判：`style & 0x2000` 必须非 0；
  墨迹行带数 == 参考（`2/2/2/1/1`）；两行的三枚墨迹 bbox 高 >= `26 * dpi/120`
  （参考 VB6 34px@120dpi、C3 字体档实测 29px、单行 ≤13px，26 是"至少两行"的下限，
  同时不把字体档差异误判）。
- `bottomBar` 对真 VB6 实拍的 diff 上限 20%（只在 `$env:UCTV_VB6_REF` 给了同尺寸
  参考图时计算，CI 无参考图 ⇒ 该子项 SKIP）。
- 新产物（3 次连跑一致）：`btns=[Tree Explorer=2L/29px | Corporate Hierarchy=2L/29px |
  Drag and Drop=2L/29px | Menu / Tools=1L/12px | Contact  List=1L/10px] twoLine=26
  bottomBarRefDiff=18.92%` → PASS。
- 旧产物 `.temp\out_probe4\Proyecto1.exe`：`'Tree Explorer' 缺 BS_MULTILINE
  (style=0x54014000); 行数 1 != 2; 墨迹高 13px < 26px`（三枚同） → FAIL。

### 23.4 分区 diff（口径 `.temp\diffreg.ps1`，959x623，`|ΔR|+|ΔG|+|ΔB|>=6`）

| 分区 | rect (x,y,w,h) | 改前(cap6) | 改后(cap7) |
|---|---|---|---|
| leftEdge   | 0,32,20,591    | 46.07% | 46.07% |
| bottomBar  | 252,540,458,83 | 19.70% | **18.92%** |
| bottomBar 客户区行 540..613 | 252,540,458,74 | 9.93% | **9.06%** |
| 5 枚按钮含边框整块 | 259/349/439/529/619,548,86,51 | 15.34% | **14.01%** |
| **整窗**   | 959x623        | 8.00% (47825px) | **7.96%** (47532px) |
| **客户区** | 10,32,940,582  | 4.96% (27138px) | **4.91%** (26845px) |

- 逐枚按钮：`Tree Explorer` 15.89%→13.43%，`Corporate Hierarchy` 22.50%→20.75%，
  `Drag and Drop` 17.44%→14.98%；两枚单行按钮 `Menu / Tools` 10.74%、`Contact  List`
  10.15% **一点没变** ⇒ 本刀的收益确实在 caption 侧，剩余 10~20% 是按钮面（主题 vs
  VB6 经典 bezel）的另一族残差。
- **leftEdge 46.07% 的定性（本刀顺带钉住）**：窗口最外 9px 条 `x0..8` 逐像素 100%
  不同（C3 `59,59,59` vs VB6 `38,38,38`，激活态边框色）；把这条非客户区排掉，客户区
  等价读数 = `x10..19` 的 diff **1.83%**（108/5910px）。所以 leftEdge 不需要再处理。

### 23.5 断言的 96dpi 修正（第二次提交）

首版把 VB6 120dpi 参考的行数写死成 `2/2/2/1/1`，CI（96dpi）判红：
`'Tree Explorer' 行数 1 != 2` / `'Menu / Tools' 行数 2 != 1`（GA run 42163469487）。

- 根因：期望值必须由**该系统/该 DPI 下按钮的真实字体**推出来，而不是抄 120dpi 参考。
  - 120dpi（本机 120dpi 桌面）：`MS Sans Serif -13` extent = 82/121/89/78/71，
    客户宽 86 ⇒ 阈值 80 ⇒ 2/2/2/1/1，与 VB6 逐枚一致。
  - 96dpi（CI）：同一逻辑的 font 档 = `MS Sans Serif -11`（光栅档，`lfHeight` 不是
    120dpi 的等比 0.8），extent = 63/94/70/64/59，客户宽 69 ⇒ 阈值 63 ⇒
    **1/2/2/2/1** —— Tree 恰好 63 不过线、Menu 恰好 64 过线。
- 改法：`UP::TextWidthOf(btn, cap)` = `WM_GETFONT → GetObjectW(LOGFONTW) →
  CreateFontIndirectW → GetTextExtentPoint32W`（测试进程屏幕 DC，与按钮同 DPI）；
  `wantLines = (capW > clientW - 6) ? 2 : 1`。两行墨迹高下限仍按 DPI 折算
  （`26 * dpi/120`）。顺带修 `GetDC/ReleaseDC` 的 DLL 归属（`gdi32 → user32`，
  写错时 `EntryPointNotFoundException` 让 `ext` 恒 -1、断言退化成全单行）。
- 96dpi 下 C3 与 VB6 的换行结果**确实不同**（VB6 96dpi 的 `Arial Narrow -13` 是
  67/62 ⇒ 2/1），那是 23.2 末的**字体继承残差**（窗体级 `BeginProperty Font` 没落到
  窗体窗口、子控件不继承），本刀不动它；断言因此在 96dpi 用"自身字体/自身 DPI"的自洽
  期望，120dpi 下仍 5/5 与参考一致。
- 本机复跑（120dpi，带 `UCTV_VB6_REF=.temp\vb6_pw.png`）：
  `btns=[Tree Explorer=2L/29px/ext82 | Corporate Hierarchy=2L/29px/ext121 |
  Drag and Drop=2L/29px/ext89 | Menu / Tools=1L/12px/ext78 |
  Contact  List=1L/10px/ext71] twoLine=26 bottomBarRefDiff=18.92%` → PASS。
- 旧产物仍红：`style=0x54014000` 缺 BS_MULTILINE、三条 `1L/13px < 26px`。
- 提交：`d6957a8c`（`tests/guisample/uctreeview_px.ps1`）；GA run **42164731930**
  = **22/22 全绿**（head `d6957a8c`）。

## §24 残差 #8：窗体级 `BeginProperty Font` 没落到窗口 + 子控件不走 ambient 继承

§23.2 末尾已经把这族留作入口（"窗体级 `BeginProperty Font` 未落到窗体窗口"）。本节
把它拆成两条独立缺陷，分开取证、分开修，并给出改后与真 VB6 的 LOGFONT 逐项对照。

### 24.1 取证（改前，120dpi 本机）

- `.frm` 数据（`tests/ucTreeView/Form1.frm`）：窗体级块 `Name="Arial Narrow" / Size=9.75 /
  Charset=0 / Weight=400`；底部 5 枚 `CommandButton`、`Check1`(AllowEdit)、`Label1`
  **没有**自己的 Font 块（按 VB6 规则应继承 ambient）；`Picture1/Picture2` 部分元素与
  `Command2` 数组自带 `MS Sans Serif 8.25`；`Text1(0)` 自带 `Arial Narrow 11.25`。
- 生成码（`--emit-c`）：Form1 的窗体级 Font 块**一处都没发射**；整个产物的
  `vb6_UC_ApplyFontToWindow` 调用 0 处。核到 `FrmControl::propertyBlocks` 里的 Font
  块只有两条路径会用（UC 实例 / 带 setter 的控件），**窗体自己的块没人读**。
- 实测（`.temp\fontprobe2_before.txt`）：
  - 窗体 `WM_GETFONT` = 一枚真实 HFONT，但读回 `MS Sans Serif / lfHeight=-13 /
    weight=400 / charset=1` —— 这是 `vb6_Vb6DefaultGuiFont()` 的默认档，不是 0。
  - 全部子控件各自一枚 `MS Sans Serif -13` 的 HFONT，与父窗口字体无关。
- 真 VB6（`D:\vb_yqt4qPac\ucTreeView\Proyecto1.exe` @120dpi，`.temp\fontprobe2_vb6.txt`）：
  - 窗体 = `Arial Narrow / lfHeight=-16 / weight=400 / charset=0 / quality=0`。
  - 所有**没有自带 Font 块**的子控件（5 枚按钮 + `Check1`）**共用同一枚** HFONT
    （handle 1560942434）—— 这就是 ambient 继承的直接证据。
  - 自带 `MS Sans Serif 8.25` 的 `Command2` 系列 = `MS Sans Serif / -14 / 400 / 0`
    （VB6 对 8.25pt@120dpi 取整到 -14，C3 截断到 -13，属既有 1px 残差，不在本节范围）。

### 24.2 两条独立根因

| # | 层 | 根因 | 判据 |
|---|---|---|---|
| A | 发射层 | 窗体级 `BeginProperty Font` 从未生成码 | `--emit-c` 里 0 处 `vb6_UC_ApplyFontToWindow`；窗体 `WM_GETFONT` = 默认档 |
| B | RTL 层 | `vb6_CreateControl` 恒发 `vb6_Vb6DefaultGuiFont()`，不读父窗口字体 | 每个子控件各自一枚 `MS Sans Serif -13`，与窗体不同的字体 |

**时序坑（改后第一版仍读默认档的原因）**：`vb6_form_create_Form1` 是在
`CreateWindowExW` **同步**执行的 `WM_CREATE` 里跑（那一刻创建全部子控件），而
`vb6_CreateFormWindowB` 末尾才执行 `vb6_UC_ApplyFontToWindow(hwnd, NULL)` —— 无条件发
NULL 会把刚落下的设计期字体当场覆盖回 8.25pt。实测症状很能说明问题：**子控件已按
ambient 继承到 `Arial Narrow -16`，窗体自己却退回 `MS Sans Serif -13`**。

### 24.3 修改站点

| 文件 | 位置 | 做什么 |
|---|---|---|
| `cgen_form_create_controls.inc` | `emitFrmFontBlock`（136）+ 窗体级发射（152） | 新增 Font 块发射器；在**任何子控件创建之前**把窗体块落到窗体窗（`vb6_UC_ApplyFontToWindow(hwnd, _vb6_fblk)`） |
| `cgen_form_ctrl_style_apply.inc` | 顶层控件路径 | 控件自己带 `BeginProperty Font` 的块也发射（此前静默丢弃） |
| `cgen_form_frame_menu.inc` | 容器子控件路径 | 同上，Frame / PictureBox 内的子控件 |
| `vb6forms.c` | `vb6_CreateControl`（701–733） | 子控件字体改为**按父窗口已落的 LOGFONT 复制一份**（ambient 继承），父无字体才退默认档；必须是复制而非共享句柄（换字体时会 `DeleteObject` 旧句柄） |
| `vb6forms.c` | `vb6_CreateFormWindowB`（469–474） | 默认档兜底加门控 `if (!vb6_ControlFontStored(hwnd))`，不再覆盖 WM_CREATE 里落的设计期字体 |
| `uc/uc_controls.c` | `vb6_UC_FontLogFont` | `Charset` 是设计期数据：0 就是 ANSI_CHARSET，`f ? f->Charset : DEFAULT_CHARSET`（旧的把 0 当"没填"换成 1） |

改后 `--emit-c` 里 Form1 的窗体块出现在**全部子控件创建之前**（第 15358 行），全产物
32 处 `vb6_UC_ApplyFontToWindow`（此前 0 处）。

### 24.4 改后实测（120dpi，`.temp\fontmetrics_c3.txt` vs `.temp\fontmetrics_vb6.txt`）

| 对象 | C3 LOGFONT | VB6 LOGFONT | TextHeight | TextWidth |
|---|---|---|---|---|
| 窗体 | `Arial Narrow -16 400 charset=0` | 同 | 20 | 82 ("Tree Explorer") |
| Button `Tree Explorer` | `Arial Narrow -16 400 0` | 同 | 20 | 82 |
| CheckBox `AllowEdit` | `Arial Narrow -16 400 0` | 同 | 20 | 55 |
| Label `MarginBottom`（父 Picture3 无块） | `Arial Narrow -16 400 0` | 同（VB6 Label 无窗口，按同 LOGFONT 等价） | 20 | — |
| Label `SelectionColor`（父 Picture1 自带 8.25 块） | `MS Sans Serif -13 400 0` | 同 | — | — |

- 最后一行是**反证**：它必须继承 Picture1 的 8.25 档，而不是窗体的 Arial Narrow。
  只测前四行的话，"子控件都在继承窗体字体"与"子控件都无视容器字体"这两种实现都会判绿。
- VB6 的 `Label` 是 windowless 控件（`EnumChildWindows` 里看不到），所以 Label 的对照
  用"同一枚 LOGFONT ⇒ 同一套度量"作等价证明；TextHeight/TextWidth 由测试进程按
  LOGFONT 本地重建字体后量（跨进程 `SelectObject` 到自家 DC 不可靠，会读到测试进程
  的默认字体）。
- 墨迹 bbox（`.temp\inkbtn2.ps1`，`.temp\vb6_pw.png` vs `.temp\px11\uctreeview_c3.png`）
  5 枚逐像素同：`50x34@(277,558)`、`57x34@(363,558)`、`52x34@(456,558)`、
  `76x11@(534,568)`、`71x11@(627,568)`；两行按钮的 lineBands 也是 `558..571 / 578..591`。
  墨迹点数 VB6 218 vs C3 197 是 AA quality（VB6=0 DEFAULT，C3=3 NONANTIALIASED）的差别。

### 24.5 分区 diff（口径 `.temp\diffreg.ps1` / `.temp\r1.ps1`，959x623，`|Δ|三方和>=6`）

| 分区 | rect (x,y,w,h) | 改前(post-#7) | 改后 |
|---|---|---|---|
| bottomBar 客户区行 540..613 | 252,540,458,74 | 9.06% | **7.69%** |
| 5 枚按钮含边框整块 | 5×(86,51) | 14.01% | **10.77%** |
| **整窗** | 959x623 | 7.96% (47532px) | **7.67%** (45813px) |
| **客户区** | 10,32,940,582 | 4.91% (26845px) | **4.59%** (25126px) |

逐枚按钮：`Tree Explorer` 13.43%→12.54%，`Corporate Hierarchy` 20.75%→**13.47%**，
`Drag and Drop` 14.98%→13.11%，`Menu / Tools` 10.74%→**8.66%**，`Contact  List`
10.15%→**6.09%**。两枚单行按钮这次也动了 —— 上一轮它们没变是因为字体还是默认档，
现在 caption 的墨迹位置/宽度终于与 VB6 同档；`bottomBarRefDiff` 18.92%→**17.70%**。

### 24.6 断言（`tests/guisample/uctreeview_px.ps1`）

1. 窗体：`WM_GETFONT` 非 0，且面名 = `Arial Narrow`、`lfHeight = -(9.75*dpi/72)`（120dpi
   -16 / 96dpi -13）、`charset = 0`、`weight = 400`。
2. 子控件继承：Button / CheckBox / Label(MarginBottom) 必须与窗体同面名同高同 charset；
   Label(SelectionColor) 必须继承容器的 `MS Sans Serif`（反证）。STATIC 不答
   `WM_GETFONT`，改读 RTL 存在窗口属性 `VB6_CtrlFont`（`GetPropW`）。
3. 按钮行数：`$wantLines`（用按钮自己的字体实测）+ VB6 参考 `2/2/2/1/1`。参考那条
   **门控在 Arial Narrow 是否已安装**：CI（windows-latest）没有这枚 Office 附带字体，
   GDI 把它替换成 Arial（-13 实测 77/115/83/75/74，两枚短标题因此翻成两行）；VB6 在同机
   也会做同样替换，所以那条参考判据在 CI 上不成立 —— 记 `faceInstalled=no(...SKIP)`
   并写进读数，字体自洽判据照常执行。
4. 旧产物仍红（`.temp\out_probe4\Proyecto1.exe` 复跑 FAIL）。

### 24.7 回归与夹具归因

| 夹具 | 结果 | 归因 |
|---|---|---|
| `uctreeview_px` | PASS（本机 120dpi：`ext82/119/87/78/72`、`2L/2L/2L/1L/1L`；CI 96dpi：`ext77/115/83/75/74`、全 2 行 = Arial 替换） | 本族正主 |
| `segmented_px` | 悬停光标 `IDC_NO` 那条红（本机 120dpi） | **与本族无关**：用 `git checkout HEAD~1 -- src/` 造出改前二进制复跑，读数除计时外逐字相同、同样红 ⇒ 120dpi 本机现象；CI（96dpi）绿 |
| `teechart_display_px` | PASS | 该夹具 96dpi（不感知 DPI，被虚拟化），字体块发射后读数不变 |
| `tree_icons_px` / `editor_scroll_px` | PASS | 同上，无像素移动 |
| 发码形状门（595 语料） | 9 份哈希变化 | 逐份 diff 只有字体块脚手架（每块 +11 行 + 1 行 `vb6_UC_ApplyFontToWindow`）与新增调用，**无其它形状变化**；据此 `-Bless` 登记 |

### 24.8 提交 / GA

- `ca2796e3`：RTL 两站 + 三个 cgen `.inc` + px 测试断言。
- `a3e55adb`：96dpi 参考行数按字体可用性门控 + `emit-manifest.expected.txt` 登记 9 行。
- GA run **42174516241**（head `ca2796e3`）：**失败 2 格** —— 发码形状门（9 份哈希，
  见 24.7）+ `tests (vbp #11)`（`'Menu / Tools' 行数 2 != 参考 1`，CI 字体替换所致，见 24.6）。
- GA run **42176788091**（head `a3e55adb`）：**22/22 全绿**。

## §25 VB 系统色 `&H80000000& Or n` 被当 COLORREF（2026-11-06）

### 25.1 现象与工程

`tests\VBFlexGridDemo\MainForm.frm`（仓库内唯一一份 VBFlexGridDemo，无第二候选）：

```
30:      Begin VB.PictureBox Picture2
31:         BackColor       =   &H8000000D&      ' COLOR_HIGHLIGHT (系统色)
32:         ForeColor       =   &H8000000E&
...
658:Private Sub Picture2_Paint()
659:Picture2.Cls
660:Picture2.Print "Drag/drop me"
```

设计器里该值是 `&H8000000D&`，运行后控件背景是黑（实测 `RGB(13,0,0)`，见 25.2）。

### 25.2 四档链路实测（120dpi / 当前主题）

| 档 | 读数 | 位置 |
|---|---|---|
| ① 解析/存储 | `frm_parser_util.cpp:70` `stoull` 得 **2147483661**（无符号 32 位），以 `long long` 存进属性表；**没有**在此折符号 | `frm_parser_util.cpp:69-75` |
| ② 发射 | `.temp\sc\emit.c`（`--emit-c` 夹具原文）两条创建路都发 `-2147483635`：<br>`694: vb6_SetControlBackColor((void*)vb6_hwnd_T13, -2147483635);`（顶层）<br>`938: vb6_SetControlBackColor((void*)vb6_hwnd_C13, -2147483635);`（容器子控件）<br>VBFlexGridDemo: `vb6_SetControlBackColor((void*)vb6_hwnd_Picture2, -2147483635);` | `cgen_form_ctrl_style_apply.inc:410/415`（顶层）、`cgen_form_frame_menu.inc:519/524`（容器子）、`cgen_form_wndproc_create.inc:167`（窗体自身） |
| ③ 属性/RTL 落地 | 旧读法：`vb6_GetControlBackColor() = 0x8000000D (-2147483635)`（保号正确）。旧绘制层裸 `(COLORREF)` 造刷 = `CreateSolidBrush(0x8000000D)`；新唯一出口 `vb6_ColorToColorRef(-2147483635)` → `GetSysColor(13) = 0x00D77800` | `vb6forms_ctrl.c:885` |
| ④ 像素 | 旧产物 `pixel=RGB(0,0,13)`（探针按 BGR 打印，实际 `#0D0000`）；新产物 `pixel=RGB(215,120,0)`，与同一行 `GetSysColor(13)=0xD77800 = RGB(215,120,0)` **逐值相等**（实际 `#0078D7`） | `.temp\pic2probe.before.txt` / `after.txt`，采样点控件客户区 (6,6)（避开中心文字） |

**根因**：`Picture2_Paint` 的 `Cls` 落到 `vb6_Form_Cls`（`vb6forms_draw.c`），该处及全仓另有
3 份 inline 判据各写各的，但**绘制层多条站点点压根不判**：直接把 `vb6_GetControlBackColor()`
的原始 VB 负值裸转 `(COLORREF)` 造刷。`CreateSolidBrush(0x8000000D)` 取低 24 位（`#0D0000`）
—— 系统色索引被当成红色分量。

### 25.3 修法：唯一出口

新增 `vb6_ColorToColorRef(int vbColor)`（`vb6forms_ctrl.c:885`，声明 `vb6forms_prop.h`）：

* 判据**逐位吻合**：`(v & 0xFFFFFF00u) == 0x80000000u && (v & 0xFFu) <= 30u` → `GetSysColor(n)`；
* 其余（含非系统负值）→ `(COLORREF)(v & 0x00FFFFFFu)`。

特意不采用 `v < 0`（会把 `&HFF00000D&` 误当系统色）或"只判低字节 n<=30"（会把
`&HFFFFFF00&` 误翻 `GetSysColor(0)`）。夹具里的 `N=&HFFFFFF00&` 就是这条反例。

所有要 COLORREF 的站点改为转调它（**不再有第二份符号判据**）：

* `vb6forms_ctrl.c`：`vb6_NativeTreeColorMsg`、`WM_CTLCOLORSTATIC` 子控件文字色/背景色、
  GroupBox 标题条背景；
* `vb6forms_draw.c`：`vb6_DrawForeColor`、PSet/Line/Circle 的 `color` 实参、
  AutoRedraw 位图填充、`Form_Cls` 填充、Printer 三条（`g_prnDrawFg` 的"未设"哨兵从
  `>= 0` 改为 `!= -1`，否则系统色负值会被当成未设）；
* `vb6forms.c`：窗体类背景刷、GroupBox 标题背景；
* `vb6forms_widget_prop.c`：`SetAutoRedraw` 初始位图；
* `vb6forms_picture_prop.c`：两处 Picture 背景刷；
* `vb6forms_shape.c`：Shape/Line 的 BorderColor/FillColor；
* `uc_host_window.c`：`vb6_uc_resolve_color`（直接转调）、ambient 背景刷两处、
  UserControl 前景/子控件文字色。

getter/setter 仍存/答 **VB6 Long 原值**（`&H8000000D&` 读回仍是 `-2147483635`）。

### 25.4 代码生成侧唯一符号折回出口

`frm_parser.hpp:69` 新增 `frmValueToVbLong(long long)`：纯算术（`unsigned long long` 掩码
`0xFFFFFFFF` + 减 `0x100000000`）把无符号 32 位位型折回负 Long；三个 `.inc` 的 5 个设计期
颜色发射站全部改调它，删掉 `(int)(unsigned int)` 双强转。`frm_parser_util.cpp` 的注释同步
改成"解析保留无符号位型，符号由发射侧折回"。

### 25.5 哨兵与运行时断言

* 静态：`scripts/check_color_ref_authority.ps1`（挂进 `tests\run_tests.ps1` 的静态 census 段，
  `Test-ColorRefAuthority`），5 条规则：
  C1 唯一出口恰好一处且三档判据齐全；C2 禁止裸 `(COLORREF)vb6_GetControl*Color`；
  C3 除唯一出口函数体那一行外禁止任何 `GetSysColor(<含 & 的表达式>)`；
  C4 `src/backend|project|driver` 禁止 `(int)(unsigned int)`；
  C5 5 个设计期颜色发射站必须含 `frmValueToVbLong`。
  白名单只两条，均带理由注释：W1 = `vb6forms_ctrl.c` 的
  `return GetSysColor((int)(v & 0xFFu));`（唯一出口本体）；W2 = `frm_parser.hpp` 的
  `frmValueToVbLong`（代码生成侧唯一折回出口，纯算术）。
* 运行时：`tests\syscolor\SysColor.vbp` / `SysColorForm.frm`
  * 31 枚系统色（0..30）走**顶层控件** `T0..T30`，另 31 枚走**容器子控件** `C0..C30`
    （父 = PictureBox `Panel`）；
  * 反例：`W=&H00FFFFFF&`（白）、`K=&H00000000&`（黑）、`N=&HFFFFFF00&`（负值但非系统色）；
  * Timer 第一拍逐枚 `Cls`（走 `vb6_Form_Cls` 的 BackColor 填充），第三拍 `GetDC/GetPixel(4,4)`
    与 `GetSysColor(idx)` 逐条比对，打印 `SC-TOP ok=31/31` / `SC-CHILD ok=31/31` /
    `SC-NONSYS ok=3/3` / `SC-DONE`；
  * 注册两例（x64/x86）：`Test-Vbp "syscolor"` / `"syscolor_x86"`。

### 25.6 回归

* 旧产物（`.temp\VBFlexGridDemo.old.exe`，改前编译器产出）跑探针：`pixel=RGB(0,0,13)`（红）。
* 新产物（`output\VBFlexGridDemo\VBFlexGridDemo.exe`）：`pixel=RGB(215,120,0)` =
  `GetSysColor(13)`（绿）。
* 新夹具 x64 / x86 实跑：`SC-TOP 31/31`、`SC-CHILD 31/31`、`SC-NONSYS 3/3`、`SC-DONE`。
* GUISample（9 格色块 + CheckBox/Label 同族回归面）：`-Category guisample` PASS
  （`samples=55`、`checkpoints=2/2`、`tree_icons_px` / `mdi` 全 PASS）。
* RTL 改动已重包 `c3rtl.res`（`c3rtl.rc` 追加账注释触发重编）；
  `scripts\check_rtl_embedded.ps1` PASS（内嵌资源与磁盘逐一相同），
  `C3.exe` 内已含新符号 `vb6_ColorToColorRef`。

### 25.7 提交 / GA

- 修复提交：**`1aa6ebd8`** `fix(rtl): VB 系统色 (&H80000000& Or n) 收成唯一 COLORREF 出口`
  （14 处消费点 + 5 处设计期发射站 + 新出口/声明 + `c3rtl.rc` 账注 + 哨兵脚本 + `syscolor` 夹具/注册），
  已推 `origin`(gitcode) 与 `github`(fxl447098457/c3test) 的 `dev`。
- 登记提交：**`2d693a69`** `登记（发码形状）：tests/syscolor/SysColor.vbp 一行`，同样双远端。
- GA run **42184975114**（head `1aa6ebd8`）：**20/22**，两红均为「新夹具未登记发码哈希」，
  不是发码/行为回归：
  - `Emit manifest (shape oracle)`：期望 595 行 / 实得 596 行 / 相同 595 / 哈希不同 **0** /
    缺席 0 / 多出 1（`tests/syscolor/SysColor.vbp`）；
  - `Tests (compile #1)`：`[STATIC] manifest_coverage` K1 同一条（新夹具不在期望清单里）。
- 处置（不许放宽阈值、不许删用例）：从 CI 自己交回来的 `github/ci/emit-manifest`
  （提交注记 `ci: emit manifest head=1aa6ebd8`）取 596 行清单，
  `scripts/rebless_emit_manifest.ps1 -Manifest .temp\emit-manifest-ci.txt -Bless`
  → **改写 0 / 新增 1 / 退场 0**；`compare` 复核 596/596 逐行相同、`check_manifest_coverage` K1..K3 绿。
  未用本机全量清单登记：本机可解析 vendor OCX/typelib，复算会剩 guisample×3 + `TeeChartChartDisplay`
  共 4 行机器侧差（与期望清单 B5 注同一条口径），登记会把 CI 的判据读数写歪。
- GA run **42189620978**（head `2d693a69`）：**22/22 全绿**（`success`，失败 job 0）。
- 本机 `-Category vbp` 全量长跑（`.temp\vbp_run.log`）：`syscolor` / `syscolor_x86` **PASS**，
  另有两族红，均归因**本机采集上下文**，非产品问题（同一份代码在 GA 上绿）：
  - `OBJCACHE` / `OBJCACHE-NEG`「没读到 OBJCACHE 读数」：本机 PS 5.1 下
    `$ErrorActionPreference='SilentlyContinue'` + `2>&1 | Out-String` 把 C3 的 stderr 读成空串
    （`.temp\oc_probe2.ps1`：`match=False`、log 空），换 `pwsh` 7 同一脚本 `match=True`；
    CI 的 `shell: pwsh` 走 7。`C3: OBJCACHE rtl=0/83 user=0/2 total=0/85` 本身正常输出。
  - `nonascii_missing_frx`「缺 .frx 却未报 VB4004」：同一条 stderr 吞字（诊断走 stderr）。
    `.temp\frx_probe3.ps1` 同一命令：PS 5.1 `VB4004=False` / `pwsh` 7 `VB4004=True`（ASCII 目录与
    「中文叉」目录两档都如此）⇒ 告警本身在，丢的是本机 5.1 的 stderr 采集。
  - `c29listview` / `c29listview_x86`「output mismatch」：同一枚 exe 的中文行在本机被写成
    `???`（`output\c29listview.out`），ASCII 行全对、needle 的中文对不上。
    `.temp\lv_probe.ps1` 用 `StandardOutputEncoding=936` 跑同一 exe 读到 `LV2-COL-TEXT=姓名`
    ⇒ 发码/打印正确，是本机宿主控制台代码页把中文降级；CI 那台按 ACP 选编码（见 run_tests
    `Get-C3ChildEncoding`），GA 该两格绿。




## §26 BalloonTooltips 家族 — 工具提示的"主题皮"：StrPtr 空串口径（2026-11-06）

### §26.1 夹具与"原始行"

夹具落在 `tests\BalloonTooltips`（自包含 10 件；源工程 `D:\vb_yqt4qPac\BalloonTooltips`，
参考产物 `BalloonTooltips.exe` 原样在该目录运行）。

**`.frm` / `.vbp` 里没有任何 `BeginProperty ToolTip*` 块** —— 每个控件的 tooltip 颜色/标题
是 `Form_Load` 里 `cTT.CreateToolTip` 的实参（末两位 = 前景色, 背景色；空 = 默认）：

```
frmBalloonTooltips.frm
149: cTT.CreateToolTip cmbBalloonTooltip, "...", cmbBalloonTooltip.List(0), TTI_INFO
150: cTT.CreateToolTip cmdBalloonTooltip, "...", cmdBalloonTooltip.Caption, TTI_WARNING_LARGE, vbRed, vbCyan, True
151: cTT.CreateToolTip cmdButtonInPic, "...", cmdButtonInPic.Caption, TTI_WARNING_LARGE, vbMagenta, vbGreen, True
152: cTT.CreateToolTip txtBalloonTooltip, "Help, I'm being repressed!", txtBalloonTooltip, TTI_ERROR_LARGE, vbBlue, vbYellow, True
153: cTT.CreateToolTip lblBalloonTooltip, "...", lblBalloonTooltip, TTI_WARNING
154: cTT.CreateToolTip lblInFrame, "...", lblInFrame, TTI_INFO
139: (按钮改色后) cTT.CreateToolTip txtBalloonTooltip, "...", "Changed TextBox Tooltip", TTI_INFO_LARGE, vbBlack, &HE1FFFF, True, False
140: (按钮改色后) cTT.CreateToolTip lblBalloonTooltip, "...", "Changed Label Tooltip", TTI_INFO_LARGE, , , True, False
```

窗体 `Font` 块是 `Microsoft Sans Serif / Size=12`，只喂窗体与控件，**不喂 tooltip**。

根因两行在 `cTT.cls`：

```
107: lToolTipStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX: If bBalloon Then lToolTipStyle = lToolTipStyle Or TTS_BALLOON
108: hWndTT = CreateWindowExW(WS_EX_TOPMOST, StrPtr(TOOLTIPS_CLASS), 0, lToolTipStyle, ...)
109: SetWindowTheme hWndTT, StrPtr(sEmpty), StrPtr(sEmpty)      ' sEmpty As String = ""
139: If lForeColor > -1 Then SendMessageW hWndTT, TTM_SETTIPTEXTCOLOR, lForeColor, 0
140: If lBackColor > -1 Then SendMessageW hWndTT, TTM_SETTIPBKCOLOR, lBackColor, 0
```

### §26.2 VB6 vs C3 逐组读数（`.temp\ttprobe.ps1`，120dpi，1920x1200）

探针口径：`EnumWindows` 抓本进程可见 `tooltips_class32` → `GetWindowLongW(GWL_STYLE/EXSTYLE)`
→ `TTM_GETTIPBKCOLOR/TTM_GETTIPTEXTCOLOR/TTM_GETMAXTIPWIDTH` → `WM_GETFONT` + `GetObjectW(LOGFONTW)`
→ 字体上 `GetTextExtentPoint32W("Ag")` 当 TextHeight；像素主色用 **tooltip 自身 DC + BitBlt**
（不受遮挡/主题外框影响，采样内部 1/3..边界避开标题与边框）。

| # | 控件 (VB6 / C3 class) | 期望背色 | VB6 像素 / bk | C3 改前 像素 | C3 改后 像素 / bk | VB6 尺寸 | 改前 | 改后 |
|---|---|---|---|---|---|---|---|---|
| 1 | PictureBox / Static（含 cmdButtonInPic） | #00FF00 | #00FF00 / 65280 | #FFFFFF | #00FF00 / 65280 | 291x99 | 294x107 | 291x99 |
| 2 | CommandButton "Enable/Disable Tooltip" | #00FF00 | #00FF00 / 65280 | #FFFFFF | #00FF00 / 65280 | 291x99 | 294x107 | 291x99 |
| 3 | Frame / Button（空标题） | #FFFFE1 | #FFFFE1 / 14811135 | #FFFFFF | #FFFFE1 / 14811135 | 298x99 | 298x107 | 298x99 |
| 4 | Label In Frame / Static（VB6 无窗口） | #FFFFE1 | — | #FFFFFF | #FFFFE1 / 14811135 | — | 298x107 | 298x99 |
| 5 | ComboBox | #FFFFE1 | #FFFFE1 / 14811135 | #FFFFFF | #FFFFE1 / 14811135 | 274x99 | 274x107 | 274x99 |
| 6 | TextBox / Edit | #FFFF00 | #FFFF00 / 65535 | #FFFFFF | #FFFF00 / 65535 | 273x78 | 273x87 | 273x78 |
| 7 | CommandButton "Change TextBox Tooltip" | #00FFFF | #00FFFF / 16776960 | #FFFFFF | #00FFFF / 16776960 | 327x99 | 327x107 | 327x99 |
| 8 | Label "Label Tooltip" / Static（VB6 无窗口） | #FFFFE1 | — | #FFFFFF | #FFFFE1 / 14811135 | — | 262x107 | 262x99 |

共同读数（VB6 / C3 改前 / C3 改后 三者**完全相同**的项）：

- `style=0x94000043`（含 `TTS_ALWAYSTIP|TTS_NOPREFIX|TTS_BALLOON = 0x43`），`exstyle=0x8`，`TTM_GETMAXTIPWIDTH=200`
- 字体 = 运行期 `NONCLIENTMETRICS.lfStatusFont`：**Microsoft YaHei UI / lfHeight -15 / weight 400 / charset 1**；`TextHeight("Ag")=20`
- `TTM_GETTIPBKCOLOR` / `TTM_GETTIPTEXTCOLOR` 三者逐组相同（VB6 与 C3 的 UDT 参数一直是对的）

唯一不同：`lfQuality` 改前 = **5**(ClearType, 主题皮) → 改后 = **0**；像素改前全白、客户区高 +8/宽 +3。

### §26.3 根因（实测闭环，不猜）

C3 旧 `vb6_StrPtr` 按 `SysStringLen` 判空 ⇒ `StrPtr("")` 答 0 ⇒
`SetWindowTheme(hWndTT, NULL, NULL)`。对 uxtheme 来说 `NULL, NULL` 是"**清空主题名**"而不是
"关视觉样式" ⇒ comctl32 v6 继续用主题皮 ⇒ 白底、客户区 +8/+9px、ClearType。
VB6 的 `StrPtr("")` 是非 0（唯一答 0 的是 `vbNullString`）⇒ uxtheme 真的关样式 ⇒ 经典气泡，
底色就是 `TTM_SETTIPBKCOLOR`。

决定性控制实验（`.temp\tt_space.txt`）：只把夹具副本的 `sEmpty` 改成非空 1 字符串
（`Const sEmpty As String = "1"`），C3 立刻与 VB6 **逐组全等**（颜色/尺寸/quality/字体），
不碰 RTL。⇒ 差异机制就是这一枚指针，不是 UDT、不是 TTM 没调、不是字体、不是 DPI 折算。

### §26.4 修改站点

| 文件 | 改动 |
|---|---|
| `src/rtl/core/vb6rtl/vb6rtl_compat.c` | `vb6_StrPtr` 改直通 `return (void*)s;`；删掉"按长度判空"，注释写清 VB6 口径 |
| `src/backend/cgen_base_naming.cpp` | `defaultValue(Vb6Type::String)` 由 `vb6_BSTR_Empty()` 改 `NULL`（未赋值 String / 返回槽 / Optional 省略位 = 空指针） |
| `src/backend/detail/expr/cgen_expr_ident_dispatch.inc` | `vbnullstring` 字面量 → `NULL` |
| `src/rtl/core/vb6rtl/vb6rtl_builtin.h` | `vb6_StrPtr` 声明注释同步 |
| `src/backend/module/cgen_com_events.cpp` | `evtArgIsLvalue` 显式排除 `NULL/true/false/nullptr`（新语义在 `RaiseEvent` 省略位暴露 `&(NULL)` → C2101，见 §26.9） |
| `scripts/owned_temp_keeps_164.txt` | `vb6_StrPtr` 的"为什么"栏随函数体变空（值仍 0） |
| `emit-manifest.expected.txt` | 形状登记（253 行 + 4 行，见 §26.9） |

### §26.5 guard（先落，先红后绿）

`scripts/check_tooltip_authority.ps1`（挂在 `run_tests.ps1` 静态 census 段，`Test-ToolTipAuthority`）：

| 规则 | 判据 |
|---|---|
| T1 | `vb6_StrPtr` **恰好 1 处定义**、函数体必须 `return (void*)s;`、不得出现长度判据（`SysStringLen`/`vb6_BSTR_Len`/`L` 前缀） |
| T2 | `vb6_StrPtrVar` 必须转调 T1 的出口（不许自己解引用/判空） |
| T3 | tooltip 底色/字色只许 `vb6_ToolTipApplyVisual` 一家发 `TTM_SETTIPBKCOLOR/TTM_SETTIPTEXTCOLOR`，且该函数内必须经 `vb6_ColorToColorRef`（系统色常量同一解析） |
| T4 | tooltip 的 `WM_SETFONT` 同口径，只许 `vb6_ToolTipApplyVisual` |
| T5 | `SetWindowTheme` 只许 DI 桩（测试注入）且**原样透传**两个指针，不许在 RTL 侧改判空 |
| T6 | tooltip 宿主恰好 2 站：`vb6forms_style.c` + `vb6forms_slider.c`（新增第 3 站即红） |

- 现状：T3/T4 散点 = **0**，白名单为空（无需豁免）；T5 桩 1 站；T6 宿主 2 站
- **改前状态**（6e31fbdf 那棵树）：`FAIL T1 vb6_StrPtr reintroduced a length test…` ⇒ **红**（RC=1）
- **改后**：`PASS tooltip authority: StrPtr 唯一出口 1 / tooltip 颜色散点 0 / tooltip 字体散点 0 / SetWindowTheme 站 1 / 宿主 2`（RC=0）

### §26.6 断言

`tests\guisample\balloontips_px.ps1`（UTF-8 BOM；`run_tests.ps1` 的 `$GuiCases` 里 BalloonTooltips
用例已 `Arch="x86"` + `PixelAssert=@("guisample\balloontips_px.ps1")`）：

- 8 组悬停（§26.2 的 8 行）+ 点 "Change TextBox Tooltip" 后重测 2 组（TextBox/Label → #FFFFE1）
- 每组：`mode=#xxxxxx` 逐通道 **±3**；`TTM_GETTIPBKCOLOR == COLORREF(期望)`（精确）；
  `style & 0x43 == 0x43`；字体面名 == 运行期 `lfStatusFont.face`；`|lfHeight - 参考| <= 1`
- **旧产物红**：`output\BalloonTooltips_x86\BalloonTooltips.exe`（18:53，改前 x86）→ **rc=1**，
  6/6 可悬停组都是 `mode=#FFFFFF` + 尺寸 107 高 + `quality=5`
- **新产物绿**：`output\BalloonTooltips\BalloonTooltips.exe` → **rc=0**，8/8 全中
- 首轮偶发不弹：3 轮"移开 → 回位抖 1px → WM_MOUSEMOVE"重试；3 轮都不弹仍 **FAIL**（不放宽）
- 无交互桌面（屏幕 BitBlt 全黑 / 无可见窗口）→ **rc=2 SKIP**，harness 记 `PixelCheck=SKIP`

### §26.7 构建 / 内嵌 RTL / 产物

- `build.bat`（Ninja/Debug，VS2019 14.29 + SDK 10.0.19041）重建；`c3rtl.rc.res` 与 `C3.exe` 同批时间戳
- `scripts\check_rtl_embedded.ps1`：内嵌 126 个资源与磁盘对象**逐字节一致** ⇒ 新 `vb6_StrPtr` 已进 `C3.exe`
- `C3 --arch x86 tests\BalloonTooltips\prjBalloonTooltips.vbp --output-dir output\BalloonTooltips`；
  PE machine = **0x014C**（与参考 VB6 产物同档）

### §26.8 提交 / GA

| 提交 | 内容 | GA |
|---|---|---|
| `6e31fbdf` | test(tooltip): T1..T6 哨兵 + 像素判据 + b180 新口径（改前 T1 红） | — |
| `47093781` | fix(rtl): StrPtr 直通 + `defaultValue(String)=NULL` + `vbnullstring=NULL` | 42218842311（3 红） |
| `d5410248` | fix(cgen): `evtArgIsLvalue` 排除 NULL 等常量（C2101） | 42225830556（1 红） |
| `780316a6` | 登记：emit-manifest 4 行吃 CI 清单 | 42227322947 **22/22** |

### §26.9 红项逐条归因（未放宽阈值、未删用例）

**Round 1 — run 42218842311（head 47093781）3 红：**

1. `Tests (compile #2)` `[STATIC] owned_temp_keeps`：K5 快照漂移 1 行 `~vb6_StrPtr|0`。
   `keeps_pointer` 仍是 **0**（没变），变的只是"为什么"栏：旧函数体调 `vb6_BSTR_Len`
   ⇒ 记 `handoff-reads:vb6_BSTR_Len`；新函数体无任何调用 ⇒ 空。`-Bless` 只改这一栏。
2. `Tests (vbp #11)` `FlexGridX86 FAIL (compile)`：`VBFlexGrid.c` 的 `onBeforeClipboardAction`
   省略位发出 `&(NULL)` ⇒ cl `error C2101: '&' on constant`。这是 `defaultValue(String)=NULL`
   暴露的既有左值判定漏洞（`evtArgIsLvalue` 把 `NULL` 当标识符）。修 cgen（d5410248）后本地
   x86 重编绿。
3. `Emit manifest`：252/596 行形状变化 —— `defaultValue(String)` 由 `vb6_BSTR_Empty()` 改 `NULL`，
   正是本次想要的形状；逐行登记。

**Round 2 — run 42225830556（head d5410248）1 红：**

- `Emit manifest`：596 行里 **4 行**与本机不同（`tests/guisample/GUISample.vbp` ×3 +
  `TeeChartChartDisplay/DisplayProject.vbp`）。这 4 个夹具的发码含外部引用/路径形，
  本机解析 ≠ runner。按脚本既定口径（`git fetch ci/emit-manifest` 取 CI 自己交回的数）
  重登记（780316a6）。复算 596/596 相同，`check_manifest_coverage` K1..K3 绿。

**Round 3 — run 42227322947（head 780316a6）：22/22。**

### §26.10 与用户前提的两处实测校正（须记）

1. `.frm` / `.vbp` 里**没有** tooltip 属性块 —— 每个控件的颜色来自 `cTT.CreateToolTip` 实参
   （§26.1 行号）。而且用户记的 "change TextBox #00FFFF" 实测是 **"Change TextBox Tooltip"
   按钮**的背色；TextBox 自己是 **#FFFF00**。
2. 字体/字号**没有差**：VB6 与 C3（改前/改后）的 tooltip 字体都是
   `Microsoft YaHei UI / lfHeight -15 / weight 400 / charset 1`，`TextHeight("Ag")=20`。
   用户看到的"字号差异"实际是主题皮把气泡高度从 99 撑到 107（字号未变）。
   窗体 Font 的 `Microsoft Sans Serif / 12` 不上 tooltip。

## §27 ucTreeView 四条缺陷 — guard 落地与续跑基线（2026-11-06）

### 27.1 guard 哨兵（已落盘、已跑）
- 文件：`scripts/check_uctreeview_authority.ps1`（**UTF-8 BOM**，PowerShell 5 无 BOM 会把中文行断成语法错误）。
- 注册：`tests/run_tests.ps1` → `Test-UcTreeViewAuthority`，STATIC shard，紧跟 `Test-ToolTipAuthority`。
- 判据：T1 事件分派 miss 可观测 / T2 模态与对话框唯一出口 / T3 TreeView 槽位三件套 / T4 多窗体统一入口 + 控件数组容器转发。
- 白名单：W1 夹具 `Form1.frm` 的用户 Declare（不扫 tests，天然不命中）；W2 `vb6forms_ctrl.c` 的 `vb6_CdShowColor`（ColorDialog/CommonShowColor 的唯一出口，放行但断言仍在）。
- **实跑（当前工作树）：PASS**；**基线与 HEAD 956ce26c 的 worktree：FAIL 11 红**（T1×2、T3×7、T4×1），T2 全绿。
- 改前/改后散点：RaiseEvent 静默分发 1→0；`vb6_VariantMissing` 0→1；`Is` 身份拆箱 0→2；槽位目标提存 0→1；数组容器转发 0→1；后端 `ShowWindow` 恒 0；`ChooseColor` 7 站全在 DI 桩 + 1 站唯一出口。

### 27.2 四条缺陷各自状态
1. **展开丢 icon（已确认根因/已改/已重编/已视觉验证绿，px 断言未加）**：省略 `Optional As Variant` 实参被填 `vb6_VariantEmpty()`，VB6 语义应为 **Missing**（`vt=10`,`0x80020004`）。改 `vb6rtl_variant.h` + `cgen_base_naming.cpp`。
   - **子项不能切换（同属第 1 条，独立成因，已改/已验证）**：`Is` 生成裸 `==`，类实例 vs `vb6_ComCall` 的 VARIANT* 恒不等 ⇒ `RadioStyle` 单选恒清空。改 `cgen_expr_binary.cpp` + `vb6com_pack.c`（`vb6_ComResultIdentity`/`vb6_ComObjectUnbox`）。实测 `RADIO N=... eq=0` 全假 → 修后 Mostrar→Ocultar→(Disabled 不响应)→Mostrar 三态可达。
   - **附带第 3 成因（已改/已验证）**：`With ... .Add(...).Image = 1` 的 `vb6_Node_Add` 被求值两次 ⇒ 多插一行（4 行 vs VB6 3 行）。改 `cgen_util_ctrl.cpp` 的 `emitVariantSlotStore`（`_vtg_` 提存）。
2. **三点按钮不弹色选（已确认根因/已改/已重编/新绿旧红）**：`Picture1` 是**控件数组容器**，转发安装块 `if (!knownControlArrays_.count())` 整类跳过 ⇒ `Command2` 的 WM_COMMAND 留在 PictureBox。改 `cgen_form_frame_menu.inc` 逐元素 `vb6_fw` 转发。实测：新产物 `#32770 '颜色'` 弹出并 `WM_CLOSE` 后 `IsWindow=false`、主窗 `WM_NULL` 有响应；旧产物 `DIALOG=` 空。
3. **LoadFromCustomXML 无效果（已确认根因/已改，px 断言未加）**：产物目录缺 `ejemplo.xml`，`xmlDoc.Load` 返回 False（`vt=11`,`parseError 0x800C0006`）。已在 `run_tests.ps1` 的 ucTreeView 注册加 `-AssetFile ejemplo.xml`；新产物点击后出现 `library` 根节点，旧产物无。
4. **底部五按钮（部分定位，未闭合）**：id 101..105 逐一 PostMessage 均**进了事件**，`vb6_form_show_FormX` 的 modal=0 一致；
   - Form3/Form5：`Form_Load` 内 `LoadFromXML App.Path\JerarquiaEmpresa.xml`、`menu.xml` 缺资产 ⇒ ctl 弹 `XML Load Error`（`#32770` 阻塞、窗体停在隐藏态）。资产在 `tests\ucTreeView` 里，只是没随 `-AssetFile` 复制。
   - Form6（`Contact  List`）：**进程崩溃 = 0xC0000005（AV，不是干净退出）**。取证口径：
     `$p.WaitForExit(8000)` → `ExitCode=-1073741819`；WER `Application` 日志
     `异常代码 0xc0000005 / 错误偏移 0x0007cea0 / 模块 Proyecto1.exe`（偏移稳定可复现）。
     早先记的"exit 0"是 PowerShell `$p.ExitCode` 为 `$null` 的读数假象。
     插桩点：`enter → imagelist init → 67 张图 → nodes done`，随后在 `ExpandAllNodes` 内崩，
     `update done` 未打印；`C3_NO_QUIT_IN_DOEVENTS=1` 仍崩；Form1 的 `WM_CLOSE/WM_DESTROY`
     均未触发（`[F1DBG]` 无输出）。真 VB6 参考产物（`D:\vb_yqt4qPac\ucTreeView\Proyecto1.exe`）
     同操作 Form6 **可见**。
   - 确定性代码错（已定位未改）：Form6 把 `StdPicture` 返回值塞进 `VT_BSTR / .bstrVal=<IPicture*>` 传给 `ImageListAddImage(SrcImg As Variant)`（C4133）；插桩改成 `VT_DISPATCH` 后图像循环可跑完。

### 27.3 已改但尚未验证/RTL 重建状态
- RTL：`vb6rtl_variant.h`、`vb6com_pack.c`、`vb6comserver.h`、`vb6comserver_cp.c`（本轮新增 `vb6_UCEventMiss` 计数）—— **`c3rtl.res` 未重建**（必须先重跑 build，再 `scripts\check_rtl_embedded.ps1` 核对内嵌符号）。
- cgen：`cgen_base_naming.cpp`、`cgen_expr_binary.cpp`、`cgen_util_ctrl.cpp`、`cgen_form_frame_menu.inc`、`cgen_com_events.cpp`（miss else 臂）。
- tests：`run_tests.ps1`（ucTreeView 注册 `-AssetFile ejemplo.xml` + sentinel）。
- 未验证项：本轮所有 cgen/RTL 改动的最终产物（`output\ucTreeView`）尚未重编；`uctreeview_px.ps1` 的四条新断言尚未扩；`c3rtl.res` 未重建。

### 27.4 工作树
- HEAD = `956ce26c`，**尚无提交**（本轮全部改动在工作树）。
- 未提交：上述 RTL/cgen/tests + `scripts/check_uctreeview_authority.ps1`。
- 诊断物（勿提交）：`.temp/uctv_dbg_sess`（插桩会话）、`.temp/relink_form1.ps1`（已扩 Form6.c + vb6rtl_system.c）、`.temp/uctv_f6_exit.ps1`、`.temp/uctv_clickcap.ps1`、`.temp/uctv_f4_ids.ps1`、`.temp/uctv_f4_all.ps1`、`.temp/uctv_forms.ps1`、`.temp/uctv_base`（基线 worktree）。
- 另：不要提交未跟踪的 `.zcodeignore`。

### 27.5 续跑顺序
1. 定位 Form6 的 **AV（0xC0000005 @ 0x0007cea0）**（优先用**正式构建链**插桩；调试会话里
   RTL 的 `[FDBG] Register/Unregister` 一条未出 ⇒ 该 obj 未必真正链入，结论不可靠）。
2. 修 `IPicture* → Variant` 装箱（cgen：对象值必须 `VT_DISPATCH`），并重跑 Form6。
3. 五按钮资产：`-AssetFile` 补 `JerarquiaEmpresa.xml`、`menu.xml`，并同步 `output\ucTreeView`。
4. 扩 `tests/guisample/uctreeview_px.ps1`：icon 计数、三态切换、色选模态（`#32770` + `WM_CLOSE` + 10s 看门狗）、XML 节点数、五窗体标题；旧红新绿、无桌面 SKIP。
5. `C3 --arch x86` 重编 → 重建 `c3rtl.res` → `check_rtl_embedded.ps1` → 产物到 `output\ucTreeView`。
6. 双远端提交 + GA 22/22（红项逐条归因），补本文件读数。
### 27.6 本次续跑落地读数（2026-11-06，接 §27 继续）

**guard 实跑**：`scripts\check_uctreeview_authority.ps1` → 当前树 **PASS（0 红）**；
基线 worktree `956ce26c` → **FAIL 11 红**（T1×2 / T3×7 / T4×1），T2 全绿；SKIP=0。
判据落点：T2 的唯一出口断言收在 `vb6_ShowForm(` + `SW_SHOWDEFAULT` + `vb6_CdShowColor`（W2 白名单），
T3 的"只求值一次"收在 `emitVariantSlotStore` 的 `_vtg_` 提存。

**四条各自定位结论（本轮实测）**：
1. 展开丢 icon + 子项不能切换：**两个成因均已改并视觉验证绿**；`uctreeview_px.ps1` 断言尚未扩。
2. 三点按钮色选：**已改、旧红新绿**（新产物 `#32770 '颜色'` → `WM_CLOSE` → `IsWindow=false`）。
3. LoadFromCustomXML：**已改、旧红新绿**（资产 `ejemplo.xml` 后出现 `library` 根节点）。
4. 底部五按钮：**未闭合**。id 101..105 全部进事件且 `modal=0`；
   Form3/Form5 缺 `JerarquiaEmpresa.xml` / `menu.xml`（`XML Load Error` 模态框阻塞 Form_Load）；
   Form6 为 **AV 0xC0000005 @ 0x0007cea0**（WER 实证），崩点在 `ExpandAllNodes` 内。

**Form6 的确定性代码错（生成码实证）**：
`.temp\uctv_dbg_sess\Form6.c:425`
`vb6_ucTreeViewPlus_ImageListAddImage(..., (&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=vb6_Form6_GetContactImgProfile(...)}), &(vb6_VARIANT){0}, 0);`
实参是 `As StdPicture` 的函数返回值（C 侧 `vb6_ComIface_Picture*`），被**当 BSTR 装箱**。
发射站点 = `src/backend/detail/expr/cgen_expr_call_arg_emit.inc` 的 ByRef-Variant 复合字面量开关
（`cType == "vb6_VARIANT"`，`switch (inferExprType(...))`，`default:` 末尾兜底发 `VT_BSTR`）。
RTL 侧正确出口已存在：`vb6_VariantFromValue` 对任意指针走 `default: vb6_VariantObject`
（`src/rtl/core/vb6rtl/vb6rtl_variant.h:263`，置 `vb6_vtDispatch` 且 `vb6_ComAddRefDispatch`）。

**下一步（按此顺序，每步落盘）**：
1. 修上述装箱站：对象/未知指针实参不得发 `VT_BSTR`（改发 `vb6_VariantFromValue(<expr>)` 或
   `VT_DISPATCH + .pdispVal`），只对**确证是 BSTR** 的实参保留 `VT_BSTR`。
2. 重编夹具 → 重跑 Form6：期望不再 AV；若 `ExpandAllNodes` 仍崩，再用正式构建链插桩定位。
3. `-AssetFile` 补 `JerarquiaEmpresa.xml`、`menu.xml`；重跑 id 101..105，五个窗体逐一对账。
4. 扩 `tests\guisample\uctreeview_px.ps1` 四条断言（含模态收口 + 10s 看门狗）。
5. `C3 --arch x86` 重编 → 重建 `c3rtl.res` → `check_rtl_embedded.ps1` → 双远端提交 + GA 22/22。
### 27.7 第 4 条闭合读数（2026-11-06 续跑）

**根因（两处，均为生成侧，均有实测对照）**

1. `As StdPicture` 返回值当 BSTR 装箱（AV 的直接前因）。
   - 发射站：`src/backend/detail/expr/cgen_expr_call_arg_emit.inc` 的 ByRef-Variant 复合字面量
     开关里 `default:` 之前（新增 `case Vb6Type::Object`）。
   - 改前生成码（`.temp\uctv_new2\Form6.c:411`）：
     `(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=vb6_Form6_GetContactImgProfile(...)})`
   - 改后：`vb6_VARIANT _vb6_vararg_N = vb6_VariantFromValue((void*)(...));` 传 `&_vb6_vararg_N`
     （`_Generic` default 臂 → `vb6_VariantObject`：`vt=VT_DISPATCH` + `vb6_ComAddRefDispatch`）。
     工程类实例仍先经 `vb6_ComObject_FromInstance`（防 `VariantClear` 把裸结构体首字段当 vtable）。
   - 读数：改前 WER `0xC0000005 / 偏移 0x0007cea0`；改后 AV 消失。

2. UC 自有 Public 方法的按名桥**丢掉了 ByRef 标量形参**（438 的直接前因）。
   - 发射站：`src/backend/detail/module/cgen_form_uc_methods.inc`。旧判据 `if (!pi.isByVal) ok=false`
     把 `ucScrollbar.ctl:2222 Public Function TrackMouseWheelOnHwnd(lHwnd&)` 整条跳过
     ⇒ 生成的 `vb6_ucScrollbar_ucMethods[]` 只有 7 项、无 `TrackMouseWheelOnHwnd`
     ⇒ `ucTreeViewPlus.Update` 的 `With ucScrollbar1: .TrackMouseWheelOnHwnd UserControl.hwnd`
     落进 `uc_hostmodel_call.inc` 末尾 ⇒ `[C3_HOST_438]` + `[C3_438] site=Call name="TrackMouseWheelOnHwnd"`
     ⇒ **Run-time error '438'** 打断 Form_Load ⇒ Form6 停在隐藏态。
   - 改法：ByRef **标量**形参加桥（局部临时量取址；引擎实参本来就是打包值，无可回写地址，
     与 VB6 对非左值实参给临时量同义）；ByRef 对象/Variant/UDT/数组仍不发。
     改后 ucMethods 表含 `{ L"TrackMouseWheelOnHwnd", 1, ... }`。

**五窗体实测（新产物 `.temp\uctv_new3\Proyecto1.exe`，资产 3 份 XML + img）**
`Corporate Hierarchy → VB6_Form_Form3 vis=True`；`Menu / Tools → Form5 vis=True`；
`Contact  List → Form6 'Agenda tema oscuro' vis=True`；`Tree Explorer → Form2 vis=True`；
`Drag and Drop → Form4 'Drag & Drop' vis=True`；全程无 `#32770`、进程存活。
（Form6 改前：`vis=False` + `#32770 'VB6 Runtime Error' 438`；更早：AV 0xC0000005。）

**嵌入 RTL**：`scripts\check_rtl_embedded.ps1` → PASS，126 个资源逐字节一致（`c3rtl.res` 已随
`scripts\build.bat` 重建；本轮之后再改 RTL 需重跑同一脚本）。

**残留（非第 4 条，另记）**：Form6 运行时 stderr 打 `THE IMAGE: 1..7 DON'T EXIST`
（`ucTreeViewPlus.ctl:985` 的 `cImageList(index)=0` 分支）⇒ 联系人头像未进 ImageList。
属于第 1 条 icon 家族的**另一实例**（key/索引口径），未在本轮闭合。
### 27.8 断言扩写与旧红新绿读数（2026-11-06 续跑，缺陷 1/2/3/4 逐条）

`tests\guisample\uctreeview_px.ps1` 追加**交互判据**段（缺陷 1/2/3/4 各自独立取样，红项并入
同一条 FAIL 通道，不另开"软失败"）：

| 判据 | 操作 | 阈值（不放宽） | 新产物实测 | 旧产物实测 |
|---|---|---|---|---|
| 1a 展开不丢图标 | 双击首行 Facebook 收起→展开，各量一次树区 | 展开态可见行 ≥ 收起态 + 3；**图标彩色像素**展开态 ≥ 收起态 + 20（粉底 255,180,180 已从 vivid 剔除） | t1=84 行/1715、t2=56 行/998 → 过 | 76/1552 vs 56/998 → 过（该旧产物已含缺陷 1 修复） |
| 2 三点按钮 | BM_CLICK 第一个 `...` | `#32770` 3s 内出现；`WM_CLOSE` 后 `IsWindow=false`（3s，再 IDCANCEL 3s 兜底）；主窗 `WM_NULL` 2s 内应 | `'颜色'` 出现且 `close=True` | `colorDialog=none` → **红** |
| 3 LoadFromCustomXML | 点击 `Load from custom XML` | ①无 `#32770` ②旧高亮 ≤100 ③勾选蓝(99,99,255) 从 >20 降到 ≤5 ④可见行数变化 ≥5 | `rows 56->12, chk 73->0, hl 0->0, dark 123` | 同口径过（该旧产物已含缺陷 3 修复） |
| 4 底部五按钮 | 逐个 BM_CLICK 五个 caption | 每枚 3s 内出现**可见** `VB6_Form_*` 且 `WM_CLOSE` 3s 内消失；主窗 2s 内应 | Tree Explorer / Form3 / Drag & Drop / Form5 / Agenda tema oscuro 全过，五枚 close=True | `Corporate Hierarchy` 3s 无窗体，后续 caption 找不到，主窗不响应 → **红** |

收口口径落实：所有"等窗口"上限 3s（`Wait-Until`/`Wait-Gone`），模态盒探测用
`EnumWindows` 找 `#32770` 且要求 `IsWindowVisible`，关闭 = `PostMessage(WM_CLOSE)` →
`IsWindow=false` 校验 → 失败再 `WM_COMMAND IDCANCEL` 3s；**无任何无界等待**。
看门狗兜底仍由套件 `-RunTimeoutSec`（默认 60s，超时杀进程 + 诊断）。

产物与哨兵：`output\ucTreeView\Proyecto1.exe`（833536 B，2026-11-07 00:58）→ px **PASS**；
`scripts\check_uctreeview_authority.ps1` → PASS（RaiseEvent 1 / ChooseColor 7 / ShowWindow 散点 0）。
`output\ucTreeView\` 同时补 `img\` + `ejemplo.xml` + `JerarquiaEmpresa.xml` + `menu.xml`。

未闭合（不在四条口径内，另记）：Form6 的 `THE IMAGE: 1..7 DON'T EXIST`（联系人头像未进
ImageList，`ucTreeViewPlus.ctl:985` 的 `cImageList(index)=0` 分支）。

### 27.9 续跑状态快照（2026-11-07 01:15，中断点；可直接照此续跑）

**A. guard 哨兵实际状态（三条，均本机实跑）**

| 脚本 | 判据 | 实跑 | 读数 |
|---|---|---|---|
| `scripts\check_uctreeview_authority.ps1` | T1 RaiseEvent 事件分派 miss；T2 ChooseColor DI 出口唯一；T3 TreeView 槽位出口；T4 多窗体 + 控件数组转发；W2 白名单 | **PASS** | `RaiseEvent 分发 1`、`ChooseColor DI 站 7`、`后端 ShowWindow 散点 0`、`Missing+Is+槽位出口齐`、`数组容器转发齐`；白名单仅 `vb6_CdShowColor`（理由：ChooseColor 的 DI 桩，非后端散点）。基线 `956ce26c` 同一脚本 **FAIL 11 红**（T1×2 / T3×7 / T4×1）。 |
| `scripts\check_box_temp_reclaim.ps1` | R1 登记站=2 / 声明=1 / 收回=1；R2 收回在同一语句尾；R3 缩进闸；R4 夹具 10 声明 / 9 收回 / 1 故意不收；R5 不重复释放、不用 VariantAssign | **PASS** | `decls=10 clears=9 unreclaimed=1 push=2 drain=1 guarded=1`（Object 分支改成不登记的复合字面量后回到 2 站）。 |
| `scripts\check_builtin_slot_returns.ps1` | R2 漏登记；R8 派生对账 | **先 FAIL → 重跑 gen 后 PASS** | FAIL 时：`R2 漏登记 vb6_VariantMissing`、`R8 derived=42 current=41 added=1 removed=0`；重跑 `scripts\gen_variant_slot_prefixes.ps1` 后 `table=42` → PASS。 |

`check_rtl_embedded.ps1` 上次为 **PASS 126 资源**（`c3rtl.res` 随 `scripts\build.bat` 重建）；
本轮 RTL 源码**未再改**（`git status` 里只有两个 cgen 文件），故内嵌核对无需重跑即成立，
重编后仍应复核一次。

**B. 四条缺陷各自处在哪一步**

| # | 缺陷 | 步骤状态 | 实测读数 |
|---|---|---|---|
| 1 | 展开丢图标 + 子项不能切换 | 已确认根因 / 已改代码 / 已重编 / 已跑断言（**绿**） | 三个独立成因：① `Optional As Variant` 省略实参填 `vb6_VariantEmpty()` 而非 `Missing` → ImageList key 口径失效；② `Is` 生成裸 `==`，工程类实例与 COM 包装指针不同层 ⇒ 单选恒 false；③ 槽位双求值多插一行。px：展开 `84 行 / 1715 vivid`，收起 `56 行 / 998 vivid`，断言 `≥+3 行 且 ≥+20 vivid`（粉底 255,180,180 已剔除）。 |
| 2 | 三点按钮不弹色选 | 已确认根因 / 已改代码 / 已重编 / 已跑断言（**绿**） | 根因：`Picture1` 控件数组容器未逐元素转发 → `vb6_fw`。px：新产物 `#32770 '颜色'` 出现、`WM_CLOSE` 后 `colorClosed=True`；旧产物 `colorDialog=none`（红）。 |
| 3 | LoadFromCustomXML 无效 | 已确认根因 / 已改代码 / 已重编 / 已跑断言（**绿**） | 根因：夹具缺 `ejemplo.xml` 资产 → `xmlDoc.Load` 失败。px：`rows 56->12`、`chk 73->0`、`hl 0->0`、`dark 123`、无 `#32770`。 |
| 4 | 底部五按钮不弹窗体 | 已确认根因 / 已改代码 / 已重编 / 已跑断言（**绿**） | 两处生成侧根因：① `As StdPicture` 返值按 BSTR 装箱 → `0xC0000005 @ 0x0007cea0`；② 按名桥丢 ByRef 标量形参（`TrackMouseWheelOnHwnd`）→ `438`。px：五窗体 `vis=True` 且逐个可 `WM_CLOSE`；旧产物 `Corporate Hierarchy` 无窗体 + 主窗不响应（红）。 |

**C. 本轮已改但尚未验证的文件（必须重编后才会生效）**

| 文件 | 改了什么 | 是否已重编 |
|---|---|---|
| `src\backend\detail\expr\cgen_expr_call_arg_emit.inc` | `case Vb6Type::Object`：命名临时量 + `boxTempsToClear_` 登记 **改成** 不登记的复合字面量 `{...vt=VT_DISPATCH, .pdispVal=...}`（只为让 R1 回到 2 站） | **否**。`C3.exe` 时间 00:33，该文件 01:10 ⇒ 产物 `Proyecto1.exe`(00:58) 仍跑的是"命名临时量"版本，**该分支的运行时行为尚未验证**；重编后必须重跑 px 与五窗体判据。 |
| `src\backend\detail\expr\cgen_variant_slot_prefixes.inc` | 重跑 gen 补 `vb6_VariantMissing`（41→42 名） | **否**（数据表，随 C3 重编生效）。 |
| RTL（`vb6comserver*` / `vb6rtl_variant.h` / `vb6com_pack.c` / `c3rtl.res`） | 本轮未改；均为已提交内容 | `c3rtl.res` 已随上一次 build 重建；本次无 RTL 变更。 |

**D. 工作树状态**

- 已提交并已推 `origin dev` + `github dev`：`70267038`（四条修复 + 哨兵）、`99053716`（px 四条交互判据）。
- 未提交：`src\backend\detail\expr\cgen_expr_call_arg_emit.inc`、`src\backend\detail\expr\cgen_variant_slot_prefixes.inc`（+10/-12）。
- 未跟踪：`.zcodeignore`（**不提交**）。
- GA 最近一次运行 `42257818768`：两格静态红（R8 派生不一致、R1 push=3），修法已落地待重编验证。

**E. 下一步顺序（每完成一条落一次盘）**

1. 重编 `C3.exe`（`.temp\build_crlf.bat` + `C3_VCVARSALL`）→ 确认 Object 分支能编译、`c3rtl.res` 重建；
2. 重生成 `output\ucTreeView\Proyecto1.exe` → 重跑 `tests\guisample\uctreeview_px.ps1`（须仍 PASS：`act=[expand 84/1715 | xml 56->12 | colorDialog='颜色' close=True | 五窗体]`），若 Object 复合字面量版本破坏第 4 条则回退该分支并另找不改 R1 的写法；
3. 重跑三条静态哨兵 + `check_rtl_embedded.ps1`；
4. 提交（cgen 两文件 + 本节读数）→ 推双远端 → `gh run watch` 追 GA 到 22/22，红项逐条归因；
5. 残留 `THE IMAGE: 1..7 DON'T EXIST` 不在四条口径内，待用户确认后再动。
### 27.10 重编 + 两格静态红转绿（2026-11-07 01:30 落地读数）

**重编**：`scripts/build.bat`（CRLF 副本 + `C3_VCVARSALL`）→ rc=0，`C3.exe` 时间 01:14:48。
只重编两只 TU（`cgen_util_type.cpp` 含槽位表 `.inc`、`cgen_expr_call.cpp` 含实参发射 `.inc`），
即两处改动都进二进制。重生成的 `output\ucTreeView\Proyecto1.exe` = 833536 B，01:26:09，
PE x86；产物 C 里对象档实参已是 `(&(vb6_VARIANT){.vt=VT_DISPATCH, .pdispVal=(IDispatch*)...})`。

**px 判据（新产物，实测 PASS）**：
`act=[expand pre=56/1005 t1=84/1715 t2=56/998 | xml rows=56->12 chk=73->0 hl=0->0 dark=123 |
colorDialog='颜色' colorClosed=True | form[Tree Explorer]='Tree Explorer' close=True |
form[Corporate Hierarchy]='Form3' close=True | form[Drag and Drop]='Drag & Drop' close=True |
form[Menu / Tools]='Form5' close=True | form[Contact  List]='Agenda tema oscuro' close=True]`
⇒ 第 4 条（`Contact  List` → Form6）在**不登记**的复合字面量版本下仍绿，说明对象档不再按
BSTR 装箱且不依赖 `boxTempsToClear_` 的第三站。

**四条守卫（同一二进制，全绿）**
- `check_uctreeview_authority.ps1` PASS
- `check_builtin_slot_returns.ps1` PASS（`table=42`）
- `check_box_temp_reclaim.ps1` PASS（`push=2 drain=1`，阈值未放宽）
- `check_rtl_embedded.ps1` PASS（126 资源逐字节一致；本轮无 RTL 改动，未重建 `c3rtl.res`）

### 27.11 GA 42262126054（head 35119006）红格归因与处置

**红格清单（5 格 / 22）与原文证据**

| job | 失败步骤 | 关键原文 | 归类 |
|---|---|---|---|
| Emit manifest (shape oracle) | 清单对表 | `发码清单比较: 期望 596 行 / 实得 596 行 / 相同 490 / 哈希不同 106 / 缺席 0 / 多出 0`；`处置… pwsh scripts/rebless_emit_manifest.ps1 -Bless` | **登记类**（含 1 条真形状修正：见下） |
| Tests (compile #1) | STATIC | `[STATIC] variant_slot_store ... FAIL`：`R1 出口本体有 9 枚 emitLine (要恰好 6)` | **真形状变化**（本轮 hoist 提存加了 3 行）→ 哨兵按新形状重钉（未放宽：新增 `_vtg_` 存在性判据） |
| Tests (compile #1) | STATIC | `[STATIC] owned_temp_keeps ... FAIL`：`K13c 推导出来的 owned 1 枚发码侧那份里没有: vb6_VariantMissing` | **登记类**（`src/common/owned_temp_keeps.hpp` 手写清单补 1 行） |
| Tests (bas #1) | EMITC-SHAPE | `[EMITC-SHAPE] b164_slot_store_shape ... missing: vb6_VariantClear(&e);` | **真回归**（`clearArg` 被写成 `(&e)`）→ 改回逐字节旧形 `&<target>` |
| Tests (bas #2) | EMITC-SHAPE | `[EMITC-SHAPE] b169_class_field_store_shape ... missing: vb6_VariantClear(&_vb6_with_0->PV)` | 同上（同一个 `clearArg` 根因） |
| Tests (vbp #1) | `[VBP-BUILD] ucTreeView` 像素断言 | `FAIL 交互: 展开/收起无差异 (两次双击后可见行数都是 33/33…)`；读数 `dpi=96 … faceInstalled=Arial Narrow:no` | **环境/采集侧**（合成双击未进窗口）→ 修脚本可靠性 |

**为什么 `&e` 那条是真回归**：本地复算 `C3.exe tests\test_fe_slot_leak.bas --emit-c` 曾见
`vb6_VariantClear((&e))`（多一层括号），是 hoist 提存改造里把非提存路也包进 `"(&" + target + ")"`；
改回 `"&" + target` 后逐字命中 `vb6_VariantClear(&e);`，b169 同批复算命中
`vb6_VariantClear(&_vb6_with_0->PV);`。

**双击那条为什么是采集侧而非业务**：`ucTreeViewPlus.ctl:1228` 的 `UserControl_DblClick`
用 `GetCursorPos` + `ScreenToClient(UserControl.hwnd)` **自己取点**，不读 lParam；旧的
`mouse_event` 在 CI 96dpi 桌面上没把两条消息送进宿主（两次读数 33/33），本地 120dpi 恰好能中。
处置：改成「SetCursorPos + 向 UC 宿主 PostMessage MOVE/DOWN/UP/DBLCLK/UP」为主路
（不依赖前台/焦点），台面鼠标为兜底；候选行最多 4 个、每步 800ms 有界；判据阈值未动。

**同一提交的本地复算（修后）**
- `check_variant_slot_store.ps1` PASS（`def 1 | calls 11 | violations 0`）
- `check_owned_temp_keeps.ps1` PASS（`K13/K13b/K13c two-way`）
- `uctreeview_px.ps1` PASS：`act=[expand pre=57/1005 a=84/1715 b=56/998 at=121,140 via=post | xml …
  | colorDialog='颜色' colorClosed=True | 五窗体 close=True]`
- `check_box_temp_reclaim.ps1` / `check_uctreeview_authority.ps1` 仍绿（收尾复跑见 §27.12）

**补记**：上一次运行 `42257818768`（head `99053716`）同样是这 6 格红（Emit manifest / vbp #1 /
bas #1 / bas #2 / compile #1 / compile #2），并非"只剩两格"——§27.9 的 D 段以此为准。

### 27.12 GA 42285817255（head 9a2d3648）红格清单、归因与登记（2026-11-07）

**run 与 head 核对**：`gh run list --branch dev --limit 5` → `42285817255` 的 `headSha=9a2d3648`，与本机
`git rev-parse HEAD` 逐字相同（就是被推上去的那一笔）。22 格里红 2 格，其余 20 格 success。

| job | 失败步骤 | 关键原文 | 归类 |
|---|---|---|---|
| Emit manifest (shape oracle) | Hash the emitted C of every corpus input | `expectation-check=FAIL rc=1`；`发码清单比较: 期望 596 行 / 实得 596 行 / 相同 575 / 哈希不同 21 / 缺席 0 / 多出 0` | **登记类**（见下逐行归因） |
| Tests (vbp #1) | Run regression (vbp #1) | `pixels FAIL [uctreeview_px.ps1]: FAIL 交互: XML 载入后仍有旧选中高亮 4446 (Nodes.Clear 未生效); XML 载入后勾选框蓝仍在 52 (CheckBoxes=False 未生效); XML 载入前后可见节点行 33->33 几乎没变 (树没被 XML 重填)` | **采集侧**（点击没进事件/没刷屏）→ 修脚本可靠性 |

同一行的 `act=` 读数（原文摘录）：`expand pre=33/660 a=58/1000 b=33/659 at=157,200 via=post
hit=[VB6_UserControlHost|;VB6_UserControlHost|] | xml rows=33->33 chk=52->52 hl=4446->4446 dark=496 |
colorDialog='Color' | colorClosed=True | 五窗体 close=True`。
⇒ 展开那一段这次是**通的**（33→58→33，`via=post`，合成双击已进窗口）；只有 XML 按钮那一次点击
在读数上完全没落地（整块树一模一样：行数/勾/高亮都不动），而它之后的 `...` 按钮 BM_CLICK 又是通的
（`colorDialog='Color'`）⇒ 不是 BM_CLICK 机制坏，是**紧跟在合成鼠标输入之后的第一次点击被吞**。

**21 份哈希的逐条归因（用 CI 自己交回的清单，不吃本机全量重跑）**

`git fetch github ci/emit-manifest` 的提交注记 `ci: emit manifest head=9a2d3648`（祖先链逐笔都在），
把三份清单逐行对齐后读数：

| 对照 | 不同的行 |
|---|---|
| expected vs `99053716` 那轮清单 | 106 |
| `99053716` → `35119006`（对象档实参改复合字面量） | 3（`tests/ucTreeView/Proyecto1.vbp`、`tests/Charts 2020/ucTreeMaps/Proyecto1.vbp`、`tests/Charts 2020/Proyecto1.vbp`） |
| `35119006` → `9a2d3648`（`clearArg` 还原 `&target`） | 98 |
| `9a2d3648` vs expected 剩余 | **21** |

即：21 = 本轮 ucTreeView 四条缺陷那笔 cgen 改动（UC 方法/事件/命名/二元表达式）留下的形状（106 里
85 份被 `clearArg` 还原救回）里，再叠加 3 份对象档复合字面量。**都是有意改的形状，不是行为回归**：
- 3 份对象档：`As StdPicture` 等实参现在发 `(&(vb6_VARIANT){.vt=VT_DISPATCH, .pdispVal=…})`，
  不再登记 `boxTempsToClear_`（authority 钉死的两站不许有第三站）；
- 其余 18 份：UC/COM/控件数组重工程（ucChart*/ucProgressCircular/ucSegmentedControl/guisample/
  VBFlexGrid/ve_list/ctrlstate/dbgdlg/test_events/cls_inh 等），形状随那笔 cgen 改动变。

**本机复算（同一二进制 `.build\C3.exe`，只取这 21 里 6 份做逐字节核对，不跑全量）**

```
rc= sha256=941074de… bytes=2196162 tests/Charts 2020/Proyecto1.vbp
rc= sha256=4ff592bd… bytes=280776  tests/Charts 2020/ucChartArea/Proyecto1.vbp
rc= sha256=51652f25… bytes=32918   tests/cls_inh/InhMain.bas
rc= sha256=986488957… bytes=239040 tests/czUI-main/czFormDemo.vbp
rc= sha256=0d641cdc… bytes=16648   tests/m7_test/M7Test.vbp
rc= sha256=42dd2780… bytes=24401   tests/test_activex_dll/test_event_dll.vbp
```
6/6 与 CI 清单的新哈希**逐字相同**（归一化口径照 `emit_manifest.ps1`：头部 `<PATH>`、根 `<ROOT>`、
Latin-1 往返）⇒ 这台二进制与 CI 那台对同一输入给同一份 C，登记的是"CI 交回来的数"而不是"我这台的数"。

**登记读数**（`scripts\rebless_emit_manifest.ps1 -Manifest .temp\ci_manifest.txt -Bless`）
```
已登记 emit-manifest.expected.txt：注脚 384 行（头部 37 + 末尾 347）+ 数据行 596 行（改写 21 / 新增 0 / 退场 0）
发码清单比较: 期望 596 行 / 实得 596 行 / 相同 596 / 哈希不同 0 / 缺席 0 / 多出 0
OK check_manifest_coverage: K1..K3 (inputs=596 [.vbp=221 .bas=375], registered=596, sets equal)
```

**XML 那条的处置（只动脚本可靠性，阈值一字未动）**

`tests\guisample\uctreeview_px.ps1` 的缺陷 3 段加了三层隔离：
1. 展开块结束后把光标挪到窗体左上（`SetCursorPos(wr.L+8, wr.T+5)`）——把"合成鼠标输入残留"最小化；
2. 点击最多 3 次（每次 1200ms、每次重新按 caption 定位按钮），两次之间 `RedrawWindow(uc, 0x0181)`，
   把「进了事件但没刷屏」与「压根没进事件」分开；
3. 第 2/3 次改走**同步** `SendMessageTimeoutW(BM_CLICK, SMTO_ABORTIFHUNG=2, 3000ms)`，超时记
   `sync-timeout;` —— 这一步能证明按钮窗口过程到底有没有处理这条消息；诊断串同时带
   `cls=/en=/vis=/cap=/try=`，以及 `gone;`（按钮中途被重建）。
断言仍按原口径：`hl<=100`、`chk<=5`、`|rowsΔ|>=5`、`dark>=30`，没有放宽、没有删用例。

本地复跑（改后，120dpi，新产物）：`PASS … act=[expand pre=56/1005 a=84/1715 b=56/998 at=121,140
via=post | xml rows=56->12 chk=73->0 hl=0->0 dark=122 try=1 cls=Button en=True vis=True cap=Load from
custom XML | colorDialog='颜色' | colorClosed=True | 五窗体 close=True] · px_rc=0`。

### 27.13 「合成鼠标输入之后第一次点击是否被吞」单发探针（2026-11-07）

**探针**：`.temp\probe_click_once.ps1`（`-Mode c3|vb6`、`-Sc A,B,C,D,E,F,G,H`）。
每个场景**只对目标按钮点一次**（无重试、无阈值放宽）；前序合成输入与这次点击分开记录。

口径：
- 画面读数 = `PrintWindow(PW_RENDERFULLCONTENT=2)` 离屏量树区 `hl`（严格 (255,180,180)）与 `dark`（lum<110），
  与桌面上有没有窗口遮挡无关；
- 捕获/焦点 = `GetGUIThreadInfo(线程id)` 与 `AttachThreadInput + GetCapture/GetFocus/GetActiveWindow` 双路；
- 消息是否进 WndProc = 单发 `SendMessageTimeoutW(BM_CLICK, SMTO_ABORTIFHUNG=2, 3000ms)`，
  返回非零即"按钮窗口过程已把这条消息处理完"，`sync=ok:<ms>`。

**探针自身两处错误（先记，避免把工具读数当产品结论）**

1. `TidOf` 曾返回 `GetWindowThreadProcessId` 的 **out 参数（进程 id）**，不是函数返回值（线程 id）
   ⇒ `GetGUIThreadInfo` 恒 `err=87`、`AttachThreadInput` 恒 False。改成返回函数返回值后两条均正常
   （本节所有 `cap/focus` 都是修后读数；§27.12 里"捕获/焦点采集不可用"的备注作废）。
2. `HlPoint` 首版把落点存成数组，打印成 `pt=,`、落点退化为盲点 ⇒ 合成双击一次都没展开（`exp == fresh`）。
   改为标量 `$px/$py`（粉底首点 +24,+6）后 `exp` 真正变化（`hl 6865→0, dark 671→1079`），
   前序"合成输入"状态才真正成立。

**C3 读数**（`.temp\uctv_new3\Proyecto1.exe`，本机 120dpi；每行 = 独立启动一次进程）

原始行（`…` 处是重复的 `cap/focus/act` 字段，只做省略；B 场景复跑一次为 `HIT t=247`，其余数值同上）：
```
SC=A HIT t=272 fresh=[hl=6865 dark=671] exp=[na] mid=[hl=6865 dark=671] after=[hl=0 dark=122] tid=8556 cap=| focus=Button|AlternateNodeColors act=VB6_Form_Form1|Form1 | attach=True cap=| focus=Button|AlternateNodeColors act=VB6_Form_Form1|Form1
SC=B HIT t=254 fresh=[hl=6865 dark=671] exp=[hl=0 dark=1079] mid=[hl=0 dark=636] after=[hl=0 dark=122] … cap=| focus=Button|AlternateNodeColors act=VB6_Form_Form1|Form1
SC=C HIT t=246 fresh=[hl=6865 dark=671] exp=[hl=0 dark=1079] mid=[hl=0 dark=636] after=[hl=0 dark=122] … cap=| focus=Button|AlternateNodeColors act=VB6_Form_Form1|Form1
SC=D HIT t=268 fresh=[hl=6865 dark=671] exp=[hl=0 dark=1107] mid=[hl=0 dark=1107] after=[hl=0 dark=122] … cap=| focus=Button|AlternateNodeColors act=VB6_Form_Form1|Form1
SC=E HIT t=274 fresh=[hl=6865 dark=671] exp=[hl=6865 dark=671] mid=[hl=6865 dark=671] after=[hl=0 dark=122] … cap=| focus=Button|AlternateNodeColors act=VB6_Form_Form1|Form1
SC=F HIT t=269 fresh=[hl=6865 dark=671] exp=[hl=0 dark=1079] mid=[hl=0 dark=670] after=[hl=0 dark=122] … cap=| focus=Button|AlternateNodeColors act=VB6_Form_Form1|Form1
SC=G HIT t=284 … mid=[hl=0 dark=636] after=[hl=0 dark=122] … dbg=… sync=ok:100
SC=H STABLE 8s fresh=[hl=6865 dark=671] exp=[hl=0 dark=1079] mid=[hl=0 dark=636] after=[hl=0 dark=636] … cap=| focus=Button|AlternateNodeColors act=VB6_Form_Form1|Form1
```

| 场景 | 前序合成输入 | 单次点击 | 读数 | cap/focus/act |
|---|---|---|---|---|
| A 基线 | 无 | PostMessage BM_CLICK ×1 | **HIT t=272**，`6865/671 → 0/122` | `cap=`空 / `focus=Button\|AlternateNodeColors` / `act=VB6_Form_Form1\|Form1` |
| B | Post 双击展开→双击收起→光标移出树区 | 同上 ×1 | **HIT t=254**，`mid 0/636 → after 0/122` | 同上 |
| C | 同 B，但**光标留在树区**（复现 CI 失败态） | 同上 ×1 | **HIT t=246**，`mid 0/636 → after 0/122` | 同上 |
| D | 只投一对 DOWN/UP（无 DBLCLK） | 同上 ×1 | **HIT t=268**，`mid 0/1107 → after 0/122` | 同上 |
| E | `mouse_event` 真实输入队列双击 | 同上 ×1 | **HIT t=274**；但 `exp==fresh`（真实双击没进树，见下） | 同上 |
| F | Post 双击后，单发**真实鼠标单击**按钮中心 | 真实单击 ×1 | **HIT t=269**，`mid 0/670 → after 0/122` | 同上 |
| G | Post 双击后，单发**同步** BM_CLICK | 同步 ×1 | **HIT t=284**，`sync=ok:100`（消息 100ms 内被窗口过程处理完） | 同上 |
| H 对照 | 同 B 的合成输入，**完全不点** | 无 | **STABLE 8s**，`after == mid == 0/636` | 同上 |

E 的口径限制：本机 `SetForegroundWindow` 没能把进程提到真前台时，`mouse_event` 的真实双击
不落在树上（`exp` 与 `fresh` 一模一样），所以 E 只能当"真实输入之后的再一次点击仍 HIT"，
不能当"真实双击确实展开了树"的证据；该证据由 B/C/D 的 Post 路径提供。

**真 VB6 对照读数**（`.temp\vb6probe\Probe.exe`：Form1 "Probe"、Go 按钮 = `Caption="CLICKED"`、
Dbl 按钮 = `Caption="DBLCLK"`，VB6 `/make` 产物；同一探针只点一次）

```
SC=A HIT t=169 cap='CLICKED' … cap=| focus=ThunderRT6CommandButton|Go  act=ThunderRT6FormDC|Probe
SC=B HIT t=155 cap='CLICKED' … cap=| focus=ThunderRT6CommandButton|Dbl act=ThunderRT6FormDC|Probe
SC=C HIT t=159 cap='CLICKED' … cap=| focus=ThunderRT6CommandButton|Dbl act=ThunderRT6FormDC|Probe
SC=D HIT t=161 cap='CLICKED' … cap=| focus=ThunderRT6CommandButton|Go  act=ThunderRT6FormDC|Probe
SC=E HIT t=171 cap='CLICKED' … cap=| focus=ThunderRT6CommandButton|Go  act=ThunderRT6FormDC|Probe
SC=F HIT t=515 cap='CLICKED' … cap=| focus=ThunderRT6CommandButton|Dbl act=ThunderRT6FormDC|Probe
SC=G HIT t=170 sync=ok:7 cap='CLICKED' … cap=| focus=ThunderRT6CommandButton|Dbl act=ThunderRT6FormDC|Probe
SC=H STABLE 8s cap='Probe' … cap=| focus=ThunderRT6CommandButton|Dbl act=ThunderRT6FormDC|Probe
```

**结论（明确口径）**

1. **C3 不吞**：A/B/C/D/F/G **六个**场景（前序输入确已生效的）+ E（真实输入后的一次点击）里，
   合成输入之后**唯一的那一次**点击都生效（HIT），包括把光标留在树区的 C 场景。H 对照组 8s 内
   画面零变化（`after == mid`），证明 B/C 的 `636→122` 不能由"后台自己刷屏"解释，只能归因给那次点击。
2. **捕获已释放**：所有 C3 读数 `cap=`（空），即点击发生前目标线程**没有鼠标捕获**；
   `focus` 停在另一枚按钮、`act` 是主窗体。`uc_host_window.c:518-519` 的
   `WM_LBUTTONDOWN→SetCapture` / `WM_LBUTTONUP→ReleaseCapture` 成对，双击序列的两次 UP
   都把捕获放掉了 —— "捕获未释放"这一假设被实测否定。
3. **消息确实进 WndProc**：G 场景单发同步 BM_CLICK，C3 `sync=ok:100`（VB6 `ok:7`），
   且随后画面变成 `0/122`。不是"消息没进队列/窗口过程"，也不是"进了但事件没跑"。
4. **真 VB6 同语义**：VB6 在 A..G 七个场景同样**全部 HIT**（`cap=` 空同样成立），H 稳定。
   即"合成鼠标输入之后的第一次点击"在 Windows 上本来就该生效；没有任何 VB6 吞点击的行为可对照。
5. 因此**结论 = 【重试是正当的】**：上一轮 CI 那次 XML 点击不落地，属**脚本时序/采集侧**
   （同一提交的下一轮 `try=1` 一次点中即通过，§27.12 已记），不是 C3 消息处理的缺陷；
   px 脚本的"最多 3 次 + 同步 BM_CLICK"是**对不可控桌面时序的容错**，不是掩盖产品缺陷。
   判据阈值仍按原口径（`hl<=100`、`chk<=5`、`|rowsΔ|>=5`、`dark>=30`），未放宽、未删用例。

**探针的局限（须记）**：本机 120dpi、有交互桌面；CI 是 96dpi，坐标/时序不可逐位复用，
故本探针只在本机跑（未进 CI shard）。它回答的是"语义层"（C3 与 VB6 是否吞点击），
不回答"CI 那台桌面在那一瞬的窗口是否被别的窗口抢焦点"——后者由 px 脚本的有界重试兜住。

### 27.14 编码取证 + 真 VB6 基准升级 + 构建口径（2026-11-07）

**全流程 pwsh 7**：本节所有采集/测试脚本一律 `pwsh -NoProfile -File …` 执行（不再有 5.1 分支）。

#### 27.14.1 字节级取证（不推断）

| 文件 | `D:\vb_yqt4qPac\ucTreeView\`（用户已转 UTF-8） | `tests\ucTreeView\`（仓库夹具） | 判定 |
|---|---|---|---|
| Form1.frm | 42417 B，无 BOM，**合法 UTF-8** | 42415 B，无 BOM，**非 UTF-8**；按 cp1252 合法（`ED 7A`=í、`F1`=ñ） | 仓库=原西语编码 |
| ucTreeViewPlus.ctl | 106567 B，无 BOM，合法 UTF-8 | 106553 B，非 UTF-8；cp1252 合法（`FA 6C`=úl、`F3 6E`=ón） | 仓库=原西语编码 |
| ucScrollbar.ctl | 145062 B，**纯 ASCII** | 145062 B，纯 ASCII | 两处**逐字节相同** |
| Form2..6 / Cls / Node | — | 无 BOM、cp1252（Form6 含 `F1`） | 同上 |

结论：**仓库夹具本身就是原编码（cp1252），不存在"夹具被转码"的保真缺陷**，不需要修夹具；
用户 `D:\` 那份才是被转成 UTF-8 的副本。

#### 27.14.2 `/make` 编码 A/B 对照（pwsh，VB6Mini `C:\Program Files (x86)\VB6Mini\bin\VB6.EXE`）

| 模式 | 做法 | 退出码 | 日志原文（关键行） | 产物 |
|---|---|---|---|---|
| ansi | 原字节复制，0 字未改 | **1** | `Form1.log` 行 416 `不能加载控件 ucTreeViewPlus1`；`ucTreeViewPlus.log` 行 31 `属性 SmoothScrollFactor在 ucScrollbar2 中有无效值。` / 行 50 `…ucScrollbar1…` | 无 |
| utf8 | Form1.frm + ucTreeViewPlus.ctl 解 cp1252 后写 UTF-8（无 BOM） | **1** | 与 ansi **逐字相同** | 无 |

结论：**编码不是变量**（两种编码的失败日志一模一样）；根因是 `SmoothScrollFactor = 0,15` 的
**逗号小数**（作者西语环境；本机 zh-CN 小数分隔符是 `.`，VB6 解析 `0,15` 为无效值）。
VB6 IDE 把它当警告、命令行 `/make` 当致命错误 —— 与用户现场一致。

#### 27.14.3 SmoothScrollFactor 合法性/等价性（先从 .ctl 查清，未猜）

- `ucScrollbar.ctl`：`Public Property Get/Let SmoothScrollFactor() As Single`；
  `PropBag.ReadProperty/WriteProperty("SmoothScrollFactor", 0.3)`（**默认 0.3**）；
  `Initialize` 里写 `0.15`；`FixDec(S) = Val(Replace(S, ",", "."))` ⇒ 作者本意就是 **0.15**；
  `If factor > 1 Then … / 100` ⇒ **1 = 关闭平滑**。
- `ucTreeViewPlus.ctl` `UserControl_ReadProperties`（2705/2710 行）运行期强制
  `ucScrollbar1/2.SmoothScrollFactor = 1` ⇒ 设计期 `0.15`→`0.15` **对行为中性**（被父 UC 覆盖）。
- 基准副本：`.temp\vb6ref\ucTreeView`（原夹具复制），**字节级**把 `=   0,15` 改成 `=   0.15`
  （ASCII 定位替换，非 ASCII 字节逐字节保留，不解码不重编码）。脚本：
  `tests\guisample\uctv_vb6_baseline.ps1`。
- 实测：`patched(bytes, 0,15->0.15): ucTreeViewPlus.ctl`；`vb6 exit=0 exe=True`；
  产物 `.temp\vb6ref\ucTreeView\Proyecto1.exe` **331776 字节**（真 VB6 基准，x86 无关）。

#### 27.14.4 C3 侧登记（同一份原编码夹具）

- frm 解析把 `0,15` 判成 ident（`FrmValue::fromIdent`）；**嵌套 UC 的设计期非标准属性根本不 emit**
  （只 emit `vb6_UC_CreateDesignUserControl(type,name,x,y,w,h)` 几何）⇒ **静默丢弃**：不警告、不报错。
- 运行期回落 `ReadProperty("SmoothScrollFactor", 0.3)`，随后被父 UC 强制 `= 1` ⇒ 与 VB6 运行时等价；
  **不会**把无效值当 0/负数用。
- C3 读夹具走 `SourceBuffer::readAndConvertToUtf8`（GBK/UTF-8 均可），本轮 C3 编的正是仓库**原编码**
  文件，`C3 exit=0`。⇒ 不存在"VB6 只吃原编码 / C3 只吃 UTF-8"的夹具层坑。

#### 27.14.5 构建口径（必须记）

- `.build\C3.exe` 是 **x64 宿主**，默认 `--arch x64` 产物启动即 `0xC0000409`（已知问题）；
  **本夹具必须 `--arch x86`**：产物 833536→834048 字节、machine=0x14C、可正常启动。
- 关闭 objcache 全量重编：`$env:C3_NO_OBJCACHE='1'` + `scripts\c3_build.ps1 -Extra "--arch x86"`。
- RTL 有改 ⇒ 必须先 `scripts\_tmp_c3_rebuild.bat`（vcvars+cmake）重编 C3，再跑
  `scripts\check_rtl_embedded.ps1`，最后编夹具。

#### 27.14.6 真 VB6 基准读数（`.temp\vb6ref\ref_metrics.json`，UC 子窗 PrintWindow）

| 窗 | w×h | textRows | firstTextW | vivid | photo | avatar | dark | 其它 | 字体 |
|---|---|---|---|---|---|---|---|---|---|
| Tree Explorer | 321×615 | 5 | 78 | 2354 | 1608 | 1225 | 1542 | white=191527 | Segoe UI -16 400 |
| Form3 (Corporate) | 541×601 | 10 | 187 | 21768 | 9292 | 6498 | 5155 | greenRim=87 / green=13790 | Segoe UI -16 700 |
| Drag & Drop | 301×441 | 5 | 98 | 2161 | 2484 | 927 | 529 | white=127827 | Arial Narrow -16 400 |
| Form5 (Menu/Tools) | 563×584 | 15 | 77 | 41664 | 56127 | 108113 | 4778 | lightGreen=11285 | Segoe UI -16 700 |
| Agenda tema oscuro | 331×491 | 1 | 331 | 7120 | 10955 | 9793 | 147049 | bg65=118627 | Segoe UI Semibold -14 700 |

#### 27.14.7 C3 对真基准的红项（本轮 `.temp\guard_ref4.log`，EXIT=1）

- **组 1（内容级，全部按真 VB6 读数判）**
  - Tree：textRows 6≠5；firstTextW 7 < 78-2；vivid/photo/avatar 全 0；dark 168 < 1388。
  - Form3：textRows 9≠10；firstTextW 127 < 185；vivid 13436 < 19591；photo 0 < 8363；avatar 54 < 5848；
    dark 1448 < 4640；字体 `MS Sans Serif|-13|400` ≠ `Segoe UI|-16|700`。
  - Drag & Drop：textRows/首行/四个颜色判据全 0（整窗空白）。
  - Form5：textRows 5≠15；firstTextW 22 < 75；vivid 36728 < 37498；dark 1814 < 4300；字体同上错。
  - Agenda：avatar 0 < 8814；字重 600 ≠ 700（face/字高已对）。
- **组 2**：第 2 轮后进程崩溃 `code=0xc0000005 rva=0x19385`（av read 0x8d8a2），
  同时 `THE IMAGE: 4/5/6 DON'T EXIST`（`outTail`）。WM_DESTROY 三行清理已加，尚未证明覆盖 UC 宿主/子窗销毁路径。
- **组 3**：六枚 `pos`/`Value` 已随点击 ±1，`Text1` 也**真联动**（`guard_ref5.log`：
  `num[BorderRadius] up: txt='5'->'6'` … 六枚全部 `txt` 与 `Max-Value+Min` 一致）。
  ⚠ 上一轮判"txt 不联动"是**采集缺陷**：跨进程 `GetWindowText` 对 Edit 只回缓存标题，
  `_Change` 早已把文本改掉也读不到；已改成 `WM_GETTEXT`（`[UP]::WTextReal`），与
  `OBJCACHE` 那组"PS 5.1 吞 stderr"同类 —— 属采集侧，不是产品。
  ~~真剩下的红只有 `NodeSpacing`/`MarginBottom` 两项树区像素未变 ⇒ setter 没触发重绘~~
  **（2026-11-07 实测已推翻）** 用 `.temp\probe_ns2.ps1`（`WM_SETTEXT` 走
  `Text1(i)._Change` 真路径 + `PrintWindow(2)` 打 UC 宿主，与遮挡无关）逐项 A/B：
  - `MarginBottom` 0→1：C3 `ucDiff=4269` / 真 VB6 `4718` ⇒ 两边都重绘，只是真 VB6 的
    `sig`（采集口径）没反映出来 —— 那半条是**采集侧**，判据改用逐字节 diff。
  - `NodeSpacing` 16→17：C3 `ucDiff=0` / 真 VB6 `8732` ⇒ **真产品缺陷**，已定位修复（见 27.14.9）。
    原因是 `Rect.Left = <Double 表达式>` 在 UDT 字段上被编成裸存（C 截断），真 VB6 是
    取整：16/2=8 vs 17/2=8.5→9。`NS->18` 起 C3 也能变（截断把 8.5 吞成 8）。

#### 27.14.8 pwsh 与 c29listview

- `OBJCACHE`/`OBJCACHE-NEG`/`nonascii_missing_frx` 三项**待用 pwsh 重验**；若转绿须在此记明
  "之前的红是 PS 5.1 采集缺陷，不是产品"。
- `c29listview` 中文 `???`：该夹具 `C29ListView.frm` 是**无 BOM 合法 UTF-8**，与 ucTreeView 的
  编码成因**不同源**，需单独归因（本轮未动）。

#### 27.14.9 账 #478：UDT 字段赋浮点表达式被 C 截断（真 VB6 取整）

**症状**：`NodeSpacing` 16→17 树区 `ucDiff=0`，真 VB6 `8732`（同一取样器、同一 `WM_SETTEXT`
路径、同一 `PrintWindow(2)` 口径）。其余五枚数值控件两边都变。

**最小复现**（`.temp\udtprobe\`，3 个文件，`--emit-c` 即可，不用 VB6）：
```
' TestCls.cls
Private Type Rect
    Left As Long
End Type
Private Rect As Rect          ' ← 变量名与类型名同名
Private lLocal As Long
Public Sub Test()
    Rect.Left = 7 / 2         ' 修前: me->Rect.Left = vb6_Num_Div(7,2)      → 3 (截断)
    lLocal = 7 / 2            ' 修前: me->lLocal = ChkLong(FltToLng(...))    → 4 (取整)
End Sub
```
修后两条都发成 `vb6_ChkLong(vb6_FltToLng(vb6_Num_Div(7,2)))`。

**根因**：`narrowTargetTypeOf → inferUdtFieldVb6Type` 用 `symTab_.lookupModule(udtName)`
查 UDT 符号。`Private Rect As Rect` 这种**变量名 == 类型名**时，`lookupLocal("Rect")` 先命中
变量（`SymbolKind::Variable`），`kind != UserDefinedType` ⇒ 整个 UDT 字段分支落空 ⇒
`narrowCheckAssign` 答 Unknown，赋值走裸存。注意 `cgen_base_generate_state_scan.inc:144`
早就为同一个同名坑在**登记侧**加了 `lookupTypeSymbol` 回退，**推断侧**漏了。

**修复**（`src/backend/cgen_util_classtype.cpp`，5 处全改 `lookupTypeSymbol`）：
- `inferUdtFieldVb6Type`（字段类型推断主路）
- `inferUdtTypeOfExpr` 的 `MemberAccessExpr` / `WithMemberExpr` 分支（含嵌套 UDT 的 `refSym`）
- `udtFieldTypeRefNameOfTarget`、`udtFieldObjCType`
（`lookupTypeSymbol` 内含 `<name>$ty` 回退 + 按 kind 的模块/本地查找，是同名遮蔽下唯一能拿到
`UserDefinedType` 的路。）

**验证**（C3 vs 真 VB6，`.temp\probe_ns2.ps1` 扫 `NodeSpacing`，`ucDiff` 逐字节）：
| 16→ | C3 修前 | C3 修后 | 真 VB6 |
|---|---|---|---|
| 17 | 0 | 7484 | 8732 |
| 18 | 7484 | 7667 | 8732 |
| 19 | 7667 | 7667 | 10005 |
| 20 | 7667 | 8586 | 10511 |
| 40 | 11884 | 11884 | — |
两边量化相位一致（17≡18、19≡20 的成对关系对上），绝对字节数差异来自渲染/抗锯齿，属预期。
`MarginBottom` 0→1：C3 4269 / VB6 4718，两边都重绘（原 `sig` 判据是采集侧）。

**回归口径**：C3 重建（`scripts\_tmp_c3_rebuild.bat`，EXIT=0）；RTL 未动，无需
`check_rtl_embedded.ps1`；夹具 `--arch x86` + `C3_NO_OBJCACHE=1` 重编 = 834560 字节。



### 27.15 设计期 StdFont（`.frx` 引用）与 UC 窗口字体落地（2026-11-07）

#### 27.15.1 红项（先出红）

`guard_ref10.log`（**修前**，C3 产物 834560B / x86）：
```
Form3  font=[MS Sans Serif|-13|0|400]   firstTextW=127   真 VB6: Segoe UI|-16|0|700  187
Form5  font=[MS Sans Serif|-13|0|400]   firstTextW=22    真 VB6: Segoe UI|-16|0|700   77
```
`.frm` 里 `font = "Form3.frx":0000` / `subtextfont = "Form3.frx":0028`，`.frx` 记录是
**StdFont 序列化块**（不是 DocString、也不是 PropertyBag）。

#### 27.15.2 根因（两条，独立）

1. **`.frx` 的 StdFont 块被当成 DocString**：`readDocString` 只能取到 ASCII 名字
   （"Segoe UI"），bag 里成了 `VT_BSTR`；`ucTreeViewPlus.ctl` 的 `ReadProperties`
   是 `Set m_Font = .ReadProperty("Font", Ambient.Font)` —— 只有 `VT_DISPATCH(IFont)`
   才认，字符串槽被忽略 ⇒ 整窗退回 ambient `MS Sans Serif 8.25pt 400`。
2. **`font` 在 `kSkip` 名单里被无条件跳过**：`kSkip` 本是给 `BeginProperty Font`
   复合块让路的，但 frx 引用形式没有 propertyBlock，跳过即整块丢失。

#### 27.15.3 修（3 处）

- `FrxReader::readFont(offset)`（`frx_reader.hpp/.cpp`）：逐字节核对
  `[DWORD cbSize][GUID 0BE35203-8F91-11CE-9DE3-00AA004BB851][DWORD 版本 1/0x02000001]`
  `[WORD 字重][DWORD 字号×10000][BYTE 标志][ASCII 名字+NUL]`；
  name 起点 = 记录偏移 +31，`nameMax = cbSize - 27`，字号限 4pt..1000pt。
- `cgen_form_create_controls.inc`：`emitFontObjFields(...)` 从 `emitFrmFontObj` 抽出
  （emit 输出**零差异**）；`kSkip` 对 `Font`/`*Font` + `FrxReference` + `readFont().found`
  精确豁免；frx 字体分支发 `vb6_UC_BagPutFont(_bag, <name>, IFont*)` 后 `continue`。
- `uc_host_window.c`：`vb6_uc_syncHostFont(HFONT)` —— VB6 的 `UserControl.Font` 改的
  是**控件窗口**字体（`WM_GETFONT` 读回的就是它）；先 `GetObjectW` 比对 LOGFONT 七字段，
  相同则 return（否则每节点一次 `InvalidateRect`），不同才走
  `vb6_SetControlFontFromLogFont`（句柄所有权仍归它的「只删自己造的」账）。

#### 27.15.4 绿（`guard_ref12.log`）

```
Form3  font=[Segoe UI|-16|0|400]  firstTextW 127→175
Form5  font=[Segoe UI|-16|0|400]  firstTextW  22→ 42
Form6  font=[Segoe UI Semibold|-13|0|700]  全部通过
```
- 字高/字面已落地；Form3/Form5 的**字重仍 400 vs 真 VB6 700**（未闭环，见 27.14.7）。
- `Form3` textRows 11 vs 真 10、`Form5` textRows 5 vs 真 15 —— 与字体无关，另有根因。

#### 27.15.5 构建/门禁口径

- RTL 有改 ⇒ `scripts\_tmp_c3_rebuild.bat` + `scripts\check_rtl_embedded.ps1`（本轮已过）。
- 夹具：`C3_NO_OBJCACHE=1` + `scripts\c3_build.ps1 -Extra "--arch x86"` ⇒ 834560B。
- **本地 GUI 判据的环境前提**：本机 Firefox（`MozillaWindowClass`）在**前台**时，
  `Charts2020` 的 `ucdump` 恒 0 bmp（无真实 `WM_PAINT`），`segmented_px.ps1` 的光标
  判据读到 `IDC_ARROW(65539)` 而不是 `IDC_NO(65559)`（`WindowFromPoint` 命中浏览器）。
  这不是产品回归：把 Firefox 最小化后，同一 exe 手工带 `C3_UC_DUMPDIR` 启动 =
  **27 张 bmp**（`.temp\chartprobe\ucdump`）。⇒ 本地跑 `test.bat gui` 前必须先最小化浏览器。

### 27.16 UserControl_Show 延迟到"控件变可见"（Form4 整窗空白闭合，2026-11-07）

#### 27.16.1 红项（修前，`guard_ref4.log` / `.temp\px_s1.log` 同口径）

```
Drag & Drop  textRows=0  firstTextW=0  vivid=0  photo=0  avatar=0  dark=0
```
即整窗空白（其余四窗已有内容，只有 Form4 全 0）。真 VB6 基准：`textRows=5 / firstTextW=98 /
vivid=2161 / photo=2484 / avatar=927 / dark=529`（`.temp\vb6ref\ref_metrics.json`）。

#### 27.16.2 根因（插桩实测，不猜）

`C3_UC_TRACE` 插桩（`.temp\fx_f4dbg`，Latin1 字节级改副本，保原编码）逐条读数：

- Form4 `Form_Load`：`F4LOAD → F4ADD×5 → F4DONE`（节点**确实**加进树里）；
- `UserControl_Show` 触发的 `UPD/FONTOK/DRAW` **全部出现在 `F4LOAD` 之前** ——
  此时 `m_Font` 尚未落地、树还空，`ucTreeViewPlus.Update` 开头 `If m_Font Is Nothing Then Exit Sub`
  直接把绘制砍掉 ⇒ 整窗只剩背景。

真 VB6 的事件顺序是 `Form_Load → 控件变可见 → UserControl_Show`；C3 原来在**建窗阶段**
（父窗体还没 `Show`）就同步调 `desc->show`，顺序反了。

#### 27.16.3 修（RTL 4 文件）

- `vb6forms_uc_internal.h`：`vb6_UCRec` 新增 `int32_t pendingShow;`
- `uc_host.c`：`vb6_uc_activate` 里 `desc->show` 段改为
  `if (!IsWindowVisible(hwnd)) { r->pendingShow = 1; trace("show.defer") } else { 原地立即发 }`；
  新增 `vb6_uc_fireShowCb`（`EnumChildWindows` 回调，逐控件 `pendingShow && ready && desc->show`
  就 `vb6_uc_push → show → vb6_uc_pop → heap_probe`）；
  `vb6_UC_FirePendingShows(parentHwnd)` 对外出口。
- `vb6forms_controls.h`：声明 `void vb6_UC_FirePendingShows(void* parentHwnd);`
- `vb6forms.c`：`vb6_ShowForm` 在**延迟 Form_Load 派发完**、`ShowWindow(SW_SHOWDEFAULT)` **之前**
  调 `vb6_UC_FirePendingShows(hwnd)` ⇒ 对齐 VB6 的"Form_Load 先跑完，控件才变可见"。

#### 27.16.4 绿（`.temp\px_s1.log`，本机 120dpi，真 VB6 读数作判据）

| 窗 | 项 | 修前 | 修后 | 真 VB6 |
|---|---|---|---|---|
| Drag & Drop | textRows | 0 | **5** | 5 |
| Drag & Drop | firstTextW | 0 | **99**（≥98−2） | 98 |
| Drag & Drop | dark | 0 | **678**（≥529×0.9） | 529 |
| Drag & Drop | avatar | 0 | **3089**（≥927×0.9） | 927 |

插桩顺序也转正：`F4DONE` 之后才出现 `UPD/FONTOK/DRAW`。

仍红（归 §27.17 图片/图标链路，与本修无关）：`vivid=1334 < 2161×0.9`、
`photo=1339 < 2484×0.9`（Form4 的 `folderClose.png` 等 PNG 未上屏）。

#### 27.16.5 构建口径

RTL 有改 ⇒ `scripts\_tmp_c3_rebuild.bat`（EXIT=0）+ `scripts\check_rtl_embedded.ps1`
（`R1..R3 OK, 126 RTL resources byte-identical, exe sha=b597e2af185d`）；
夹具 `C3_NO_OBJCACHE=1` + `scripts\c3_build.ps1 -Extra "--arch x86"` ⇒ **835584B**（含 show defer）。
本次 guard 全量读数：`.temp\px_s1.log`（EXIT=1，红项见 §27.14.7 + 本节）。

### 28 现场快照（截断恢复，2026-11-07）

#### 28.1 git 实况

- HEAD `6cbecc78`（`fix(uctreeview): UserControl_Show 延迟到控件真正可见`，已推）。
  其前：`97cd6c9a`（px 脚本 System.Drawing）、`807d5077`（设计期 StdFont）、`916b4a1a`（真 VB6 基准 + 账 #478 + 滚动条命中步进）。
- 本轮未提交改动（9 个源文件，`git diff --stat` 125+/12-）：
  `src/backend/decl/cgen_decl.cpp`、`src/backend/cgen_util_ctrl.cpp`、
  `src/backend/detail/util/cgen_helpers.inc`、`src/backend/detail/stmt/cgen_assign_value_sem.inc`、
  `src/backend/detail/expr/cgen_expr_call_arg_emit.inc`、`src/backend/stmt/cgen_stmt.cpp`、
  `src/backend/expr/cgen_expr_binary_util.cpp`、
  `src/rtl/core/vb6rtl/vb6rtl_bstr.h`、`src/rtl/core/vb6rtl/vb6rtl_string.c`。
- 未跟踪：`.zcodeignore`（工具生成，未纳入）；`把`（0 字节误重定向产物，删后又被重现，待再删）。
- stash：`stash@{0}` = WM_DESTROY 清槽（组 2，第 2 轮崩，未闭环）；`stash@{1}/{2}` 为历史无关项。

#### 28.2 本轮已实测（全部来自实跑，未推断）

1. **UDT 定长串 ABI（`As String * N` → `char[N]`）**
   - 修前（`.temp\shtest.bas`，同一台机器）：`Len(tSHFI)=20`、`SHGetFileInfo ret=0`、`hIcon=0`。
   - 修后：`Len=352`、`ret=1`、`hIcon≠0`、`name=[notepad.exe]`（flags `&H100 Or &H1 Or &H800 Or &H200`）。
   - 发码：结构体侧 `char name[N]`；读侧 `vb6_FixedStrToBSTR`（`wrapToBSTR` 只认成员访问，
     裸定长串仍走 BSTR 模型）；ByRef String 实参用 `&(BSTR){vb6_FixedStrToBSTR(...)}` 复合字面量；
     写侧 `vb6_FixedStrAssign` / `vb6_FixedStrAssignFree`（`emitBstrSlotStore` 新增
     `udtFixedField` 参数，仅 `value_sem` 传 true）。
2. **`Chr$(0)` 语义**
   - 修前（`.temp\chrtest.bas`）：`len(s)=4`、`lenchr=0`、`instr(s,Chr$(0))=1`。
   - 修后：`len=5`、`lenchr=1`、`instr=4`。
   - 动机：`ClsTreeExplorer.GetSpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)`；
     `Chr$(0)` 为长度 1 的 NUL 串（原实现 `vb6_BSTR_FromStr` 经 `wcslen` 退化成空串）。
3. **Declare 出参定长串整幅回写**（新 `vb6_BSTR_FromANSIBufFixed(buf, N)`，含尾部 NUL）
   - 修前插桩：`GSF csidl=0 z=1 p=[]`（InStr 命中空 needle 规则返回 1）→ `INIT 1 f=[]` →
     `SHGET ret=0 hIcon=0 len=20 path="\"`。
   - 修后插桩（`.temp\fx_fs4`）：`GSF csidl=0/38/5 ret=1 n=512 z=31/23/33
     p=[C:\Users\Administrator\Desktop]` 等；`SHGET ret=1 hIcon=5704179/5835251/73860413
     len=352`。
4. **guard 全量（`.temp\px_fs2.log`，EXIT=1，120dpi，`-RefJson .temp\vb6ref\ref_metrics.json`）**
   - Tree Explorer：`textRows=4`(ref 5)、`firstTextW=79`、`vivid=562`、`photo=643`、
     `avatar=314`、`dark=952`。修前 `firstTextW=7 / vivid=0 / photo=0 / avatar=0 / dark=168`。
   - Corporate：`textRows=11`(ref 10)、`firstTextW=175`(ref 187)、`photo=0`、`avatar=58`、
     `weight=400`(ref 700)。
   - Drag & Drop：`textRows=5`、`firstTextW=99`、`dark=678`、`avatar=3089`；
     `vivid=1334`(ref 2161)、`photo=1339`(ref 2484)。
   - Menu/Tools：`textRows=5`(ref 15)、`firstTextW=42`(ref 77)、`weight=400`(ref 700)。
   - Contract/Agenda：`vivid/photo/avatar=0`。
   - 组 3（数值微调）：六枚 `up/dn` 各一轮 `pos/Value/Text1` 均联动（读数在日志）。
   - 组 2（重复点击）：第 2 轮 5 个按钮全部 `OPEN-FAIL`；Menu/Tools 第 1 轮崩溃
     `code=0xc0000005 rva=0x45dfc`，outTail 含 `THE IMAGE ... DON'T EXIST`。
5. **构建口径**：`scripts\_tmp_c3_rebuild.bat` EXIT=0（含本轮 RTL/发码改动）；
   夹具 `C3_NO_OBJCACHE=1` + `c3_build.ps1 -Extra "--arch x86"` EXIT=0；
   `.temp\utv_fs2\Proyecto1.exe` 与 `.temp\fx_fs4` 均为本轮产线。

#### 28.3 未验证（不补推断）

- 未跑 `scripts\check_rtl_embedded.ps1`（RTL 已改）。
- 未跑 GA，未核 `ci/emit-manifest`；未跑 tests 全量（`Chr$(0)`/定长串语义可能影响 golden）。
- `GetDrive()` 实测 `GetDrivesCount drives=0`（UBound=0），盘符根未进树；未归因。
- `Debug.Print ... & cInitialFolders.Count` 读出 `9039188`/`13836620` 这类指针值（Loop 上界走
  `vb6_ComGetIntProp` 未受影响）；未归因。
- `f = f & Chr$(0)` 后再 `InStr(f, Chr$(0))` = 0（定长串赋值/拼接路径丢 NUL）；未归因，fixture 不用此形态。
- `c29listview` 中文 `???` 两条红未复核。
- `.temp\fx_f4dbg` 为插桩副本（Latin1 保原编码），临时件不入库。

#### 28.4 下一步

1. 先跑 RTL 门禁 + GUI 10 例 + tests 全量，红项逐条归因；
2. `GetDrive`/`Collection.Count`；Form3 头像与 `weight=700`、卡片右缘；Form5 `textRows=15` 与字体；
   Contract 头像链；Drag & Drop `vivid/photo`；
3. 组 2：先解 `stash@{0}` 的崩溃根因再重放 WM_DESTROY 清槽；
4. 组 3 已绿，按 GA 结果确认后可关账。
