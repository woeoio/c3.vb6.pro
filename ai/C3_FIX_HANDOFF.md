# C3 编译器错误修复 — 任务交接文档

> 用途：新会话恢复上下文用。新开会话后直接说「读取 C3_FIX_HANDOFF.md 并继续修复」。
> 更新日期：2026-09-22
>
> **本文档按 owner 要求只保留「未完成项」与「仍在生效的口径/工具事实」**，已完成修复的过程
> 叙述已删除（3310 行 → 约 200 行）。回溯通道有两条，都不必靠本文档：
> - 删除前的完整版：`git show 1465da1:ai/C3_FIX_HANDOFF.md`（含原 §1–§44 全部叙述与取证）
> - 每个修复在**代码落点处自带注释**：`grep -rn "Fix <N>" src tests`（注释里写的是根因与判据，
>   比本文档的转述更不易过期）。已提交项另见 `git log --oneline`。
>
> §D 是一行式索引，标明每个已完成项原来的 §区间，便于上面那条 `git show` 定点取阅。
>
> **下文里的行号是删除时的位置，重构后可能已漂移**（`src/backend/` 下大量逻辑已拆进
> `detail/{expr,stmt,util,base,module}/*.inc`）；文件路径是本轮复核过的真实路径，定位一律用
> `grep -rn <符号或字面量> src`，不要把行号当坐标。

## A. 当前状态与验收基线

- 长期目标：让 C3 编译**并运行** `D:\vb_yqt4qPac\VBFLXGRD-master\Standard EXE Version\VBFlexGridDemo.vbp`，
  画面与用户给的参考截图一致。
- 现状（2026-09-22 15:00）：demo 编译链接全绿、进消息循环、画面各验收点一致，**鼠标悬停与左键
  点击均不再崩**（用户实测）。全量门禁 **`PASS=93 FAIL=0 SKIP=1 TOTAL=94`**（`regress50`；SKIP =
  `test_vbman`，`VBMANLIB.cVBMAN` 未在 WOW6432Node 注册，属环境）。
- **画面验收于 2026-09-22 16:30 关闭**：最后一个已知差异（"Drag/drop me" 框空白）随 **Fix 185**
  落地，提交 `76fd19a`，门禁 `regress51` = `PASS=93 FAIL=0 SKIP=1 TOTAL=94`（与 `regress50` 同分）。
  即：**demo 的运行期画面已与参考图逐项一致**，B 组剩下的都是非画面项（潜在错值/崩溃、覆盖面、
  口径缺口）。注意参考图原件 `Desktop\2026-09-21_231559.png` 已不在盘上，比对基线只剩
  `.temp/demo_shot*.png` 这一串（`demo_shot39.png` = 改前基线，`now185.png` = 改后）。
- 验收口径（重要，别反向"修好"）：**"与参考图一致" ≠ "完美"**。参考图里 `Partial Scr…` 本身就是
  截断的，不要为了消除截断去改控件尺寸（原 §21/§23/§25）。
- 门禁号与 Fix 号是**两条独立序列**；编号若与代码注释冲突，以代码注释为准并在此说明（原 §28
  有整张重排表）。任务 #33 原挂 "Fix 191"，已让号为 **Fix 192**。
- 分支：工作在 `fan/dev`；`main` 受保护，只能走 MR。

## B. 未完成项（每条都可直接开工；出处 = 删除前的 §号）

### B1 Fix 185 的**边界**：控件级 `_Paint` 只接了 PictureBox（Label / Frame / Image 仍无派发）
Fix 185 已落地（见 D 表），当时按"爆炸半径一个点"的口径把 WM_PAINT 派发与 `Print/Cls` 两条
都门在 `FrmControlType::PictureBox` 上 —— 依据是全工程度量 `<控件>_Paint` 只有 `MainForm.frm:658`
一处（`UserControl_Paint` **不在此列**：它走 Fix 112 的宿主路径 `vb6_<ctl>_ucHostPaint`，
`src/backend/module/cgen_form.cpp:189/241`，本来就通）。所以 **Label / Frame / Image 的
`<控件>_Paint` 现在仍然是死代码**（成员侧不认 `print`/`cls`，调用侧照旧落
`vb6_ComCall(HWND, …)` 运行期 no-op）。要扩面注意两点：
① 先重量真实工程里 `<控件>_Paint` 的出现面，别照抄 PictureBox 分支；
② `Image` 的 WM_PAINT 已被 `vb6_InstallImageSubclass`（`vb6forms_picture_prop.c:305-312`）占走，
而它和代码生成的 `vb6_InstallControlSubclass` **共用同一个 `VB6_OrigProc` 属性名** ⇒ 后装的一方
静默 no-op（见 B13）；扩到 Image 必须先解掉那条，否则"加了派发"和"没加"表现完全一样。

另记一条**覆盖面**事实（本轮度量）：`tests/` 里没有任何用例写 `<控件>.Print`/`.Cls`
（`grep` 只命中 `Debug.Print`），所以门禁**不经过**这条新路径 —— 它既不构成回归风险，
也意味着 Fix 185 没有自动化断言，只能靠 VBFlexGridDemo 的 GUI 用例保证"编得过、跑得活"。

### B2 Fix 180 —— `ReDim` UDT 成员数组不带 `As` 时按 Variant 分配载体（步长不符）
出处 §22 L1539–1545、§30（根因已收窄到一行判定）。现象：`ReDim h1.Items(0 To 1)`（不带 `As`）发成
`vb6_SafeArrayReDim1D(vb6_sa_variant,0,1)` → 16 字节步长，而读写走 `VB6_SA_AT(vb6_type_TOwned,…)` → 越界。
**落点**（已复核仍在）：`src/backend/stmt/cgen_redim.cpp:12-18`（只有 `node.targetExpr` 才走复杂通道）
+ `:74-79` 的 `arrayUdtElemTypes_.find(lowerVar)` —— 键是**裸变量名**，`h1.items` 永远捡不到 →
`isUdtArray=false` → 发 Variant 版；多维分支 `:282` 共用同一个 `udtCType`。
**修法**（约 6–8 行）：按最后一个 `.` 切 owner/member → 用 `knownUdtVars_` 取 ownerCType
（登记处 `src/backend/decl/cgen_decl_var.cpp:180`）→ 调已有的
`udtFieldObjCType(ownerCType, member)`（`src/backend/cgen_util_classtype.cpp:338`）；建议抽成
`resolveUdtMemberArrayCType()`，与 `resolveReDimComplexElemType`(:200-231) 共用。**owner 解析不到时
保持回落 Variant，不要猜。** 边界：类字段 `me->Foo.Items` 与 With 里的 `.Items` 需要
`classMemberVars_`/`withObjectInfoStack_`，本轮不扩。
**验证判据**：生成行变成 `vb6_SafeArrayReDim1D_Udt((int32_t)sizeof(vb6_type_TOwned),0,3)`。
（`--dump-symbols` 不打印 UDT 成员字段，证不了 `typeRefName`，见 C 组第 5 条。）

### B3 Fix 192 —— `ByVal <x> As Currency` 传 Win32 `POINT` 的 x64 编组 + DI 桩按声明形状分流
出处 §43 第 2 条、§40 前的勘察。x64 下**只有打包成单个 64 位整数**（x 低 32、y 高 32）才对：
`.temp/abi_probe.c` 以真原型为基准、在可见窗口上实测 —— 我们的 `(double)` 桩与上游的
`(intptr_t,intptr_t)` 桩**都返回错误的 HWND**，`(int64_t)` 返回正确值。受影响面是
`WindowFromPoint`/`ChildWindowFromPoint` 的 hover 路径（VBFlexGrid + 整个 ComCtlsDemo 家族），
两条线上一直搞错。
**难点不是选边**：DI 桩只按 **API 名**索引，同一 API 的两种真实 VB6 声明形状无法共存（工作树已自
相矛盾：`WindowFromPoint` 用上游 2 参、`ChildWindowFromPoint` 用我们 1 参，而 VBFlexGrid 里两者都是
`As Currency`）。要修必须让**桩名按声明形状区分**（codegen + `scripts/gen_di_stubs.ps1` 同改）。
**但先量冲突面再设计**：`.temp/di_shape_survey.py` 走完 `D:\vb_yqt4qPac` 全部 `.frm/.bas/.cls/.ctl/.pag`
的结果是 **1116 个 Declare 组里只有 1 个 API 的元数不同**（`user32!WindowFromPoint`）；~351 条"同元数
不同 As 类"是 `PtrSafe`/非 `PtrSafe` 双声明惯用法（`As LongPtr` vs `As Long`），桩取 `intptr_t` 无害。
⇒ **不要为单一样本给 583 个桩全部改名**；窄修 = 一个备用符号 + 调用点按元数分流。
（该 survey 的正则需 `re.M`，否则 `^` 只匹配字节 0，会静默报 0 组。）

### B4 `Set <接口字段> = Me` 存成类实例而不是 COM 身份
出处 §44 第 3 条。`Set .OriginalIOleIPAO = This` 把 `me` 直接塞进接口字段，而 VB6 语义是存该接口的
**COM 身份**（`me->__comObj`）。后果：`Set … = Nothing` → `vb6_ReleaseObject` 对 `me` 解引用 `lpVtbl`
（`me->__comObj` 为 NULL 时读 `NULL+0x10`）。本轮**只在运行时兜住**（`vb6_ComIsDispatchable` 挡掉非
COM 接收者），代码生成侧未动 —— 该修的是赋值处的右数取 `__comObj`。
同族前置认知（§44）：`As LongPtr` 数组是**手写伪 vtable** 的载体；`QueryInterface/AddRef/Release` 在
**所有**接口 vtable 里固定在槽 0/1/2，按名后期绑定会去调 IDispatch 槽 5 的 `GetIDsOfNames`。

### B5 `Set 变量 = New <本工程类>` 之后再传该变量仍 AV（EXE 工程类 IDispatch 桥接的未覆盖面）
出处：原「合并自 origin/main」块里**唯一不重复**的一段（EXE 工程类 IDispatch 桥接，提交
`3dd3689/17f349a/df42abc/3b797f9`）的"已知限制（未覆盖）"。`comPackExpr` **只识别字面 `NewExpr`**
（已复核：`src/backend/cgen_util_com.cpp:228` 仍只有 `expr.kind == NewExpr` 分支），所以
`Set h = New bHello` 后把变量 `h` 传出去仍走裸结构体指针 → AV。同段另记一条风险：`__comObj` 成为
EXE 类结构体首字段属**全量布局变更**。

### B6 嵌入清单的**三方**不变式没有固化成守卫
出处 §43 第 1 条。`src/driver/c3rtl.rc`（`<id> RCDATA "../<path>"`）↔ `src/driver/rtl_embedded.hpp`
（`RTL_NAME = <id>`）↔ `src/driver/rtl_embedded.cpp`（`{ RTL_NAME, "basename" }`）必须逐个 id 对齐，
而 `.cpp` 只按 **basename** 索引 ⇒ id 错位的表现是"抽出另一个文件"而不是报错。
合并提交 `c9b7986` 就这样丢过 RCDATA 205（`vb6_di_unknown_stubs.c`，Fix 164 的文件）没嵌进 C3.exe，
同族 DI 桩 **489 → 421，162 个手写桩丢失**（桩体已恢复并提交为 `c0b3729`，**不变式检查本身仍只有
`.temp` 里的脚本**：`.temp/yqt_embed_xcheck.sh <rev|WORKTREE>`、`.temp/di_body_audit.py`（比**桩体**
而非只比名字——名字是全并集也可能逐符号取了对方桩体而静默回退手工修复）、`.temp/di_sig_divergence.py`）。
可开工的两步：把这三份脚本收进 `scripts/` 并挂进 CI；在 CMake/构建期加一条 rc↔hpp↔cpp 计数与 id 断言。

### B7 `On <expr> GoTo 100, 200`（数字标签入口分发）解析不通
出处 §12 L1057–1061，**已复核仍在**：`src/parser/stmt/parser_stmt_jump.cpp` 的 `parseOnStmt`(:17-35)
只在 `next_` 是 `Error`/`GoTo`/`GoSub`（或 Identifier=="local"）时分派；`On k GoTo …` 时 `next_` 是
标识符 `k` → 直接报 "expected 'Error GoTo'…"，`parseOnGoToStmt` 不可达。修法很小（`peek()/peek2()`
向前看一步，约 6 行）。相关能力（数字行号标签、`GoTo/GoSub/Resume/On Error GoTo` 数字）已由
Fix 171 落地，缺的是这一条分发。

### B8 160-D —— `With <返回 UDT 的函数>()` 类型推断缺失
出处 §10 L835–854（任务 #14）。复现：`Common.bas:698/708/718` 的 `With GetAppVersionInfo()`。
`src/backend/stmt/cgen_with.cpp:443` 起只有在 `tempType=="void*"` **且** `inferUdtTypeOfExpr`
（`src/backend/cgen_util_classtype.cpp:166`）命中时才切结构体，而 `inferUdtTypeOfExpr`
**没有"调用返回 UDT 的函数"分支**。
两个坑：① 同文件对函数右值发 `&(<expr>)` → 非法，需先落值临时；② 注册表不能按生成
顺序取（`AppMajor`:697 出现在 `GetAppVersionInfo`:726 **之前**）→ 需从符号表拿返回类型名，而
`src/semantics/symbol_table.hpp:224` 的过程符号只有 `Vb6Type returnType`（枚举，表达不出"哪个
UDT"），带名字的是 **Class 符号专属**的 `memberReturnTypes`（`unordered_map<string,string>`，
消费点 `src/backend/cgen_util_classcall.cpp:457`）⇒ 要么给过程符号补一个返回 UDT 名，要么在
`inferUdtTypeOfExpr` 里查 `SymbolTable` 的过程签名。

### B9 `With New <Form 模块>` 类型未推断 → 整块退化成 COM 后期绑定（任务 #13）
出处 §10 L856–867（原记为"160-W 跨过程发射泄漏"：`MainForm.c:617/628` 用了上一个过程的
`_vb6_with_14/15`），§37 L2472 已把根因改判为本条。现象是 `InputForm` 对话框静默失效。

### B10 Fix 170b —— `vb6_VariantArray()` 恒定打 `VT_ARRAY|VT_VARIANT`
出处 §13 L1100–1102。⇒ `VarType(b) = vbArray + vbByte` 这类判定永远不成立（定义在
`src/rtl/core/vb6rtl/vb6rtl_variant.h`）。建议按载体自身 `elemType` 推出元素 VT。
**注意**：`vb6_VariantToByteArray`/`vb6_VariantToSafeArray1D` 在 Variant 已持数组时是直接
`return v.parray`（**别名不拷贝**），改这里要连带回答所有权问题。

### B11 `Print #f, "x="; <返回 String 的用户 Function>` 把字符串打成 int32
出处 §16 L1254–1256（Fix 176 只修了 `Debug.Print`）。**已复核仍在**：
`src/backend/stmt/cgen_file_io.cpp:114` 非 BSTR 分支发 `vb6_Print(fnum, vb6_Str((int32_t)(val)))`，
`visit(WriteStmt&)` 同形（:153）。⇒ 写文件的探针读数会静默误判。

### B12 `uc_host.c:334` 的 UserControl 内嵌 edit 仍用 `DEFAULT_GUI_FONT`
出处 §23 L1601、§27 L1824（**已复核仍在**）。与 Fix 181 同类：VB6 口径是 MS Sans Serif 8.25pt，
现代 `DEFAULT_GUI_FONT` 是 Segoe UI 9pt。改的时候连带看 C 组第 3 条的 GDI 所有权陷阱。
### B14 两条待复验的签名/覆盖面嫌疑（出处清楚，但我这轮**没能**在代码里定位到原行号）
- §40 L2655–2657：`vb6forms_axsite.c` 的 3 条 **C4113**（签名与槽位不符）⇒
  `IOleInPlaceSiteWindowless` vtable 初始化顺序与接口顺序对不上（不是 NULL 槽，但会调错函数）。
  我按原记的行号（:855）没找到，需重新从当次构建的 MSVC 输出取证。
- §41 L2696–2700：极小复现 `tests/Charts 2020/ucChartArea/Proyecto1.vbp` **还没沉淀成门禁用例**；
  三条**2026-10-04 真编译重测**（`.build/b187out/`，`--arch x64` + `--arch x86` 各一遍，工程名一律是目录下的 `Proyecto1.vbp` 而不是 `<UC名>.vbp`）：**ucChartArea 与 ucPieChart 现在编得过、两份 exe 都出** ⇒ 可以直接升格进门禁；仍失败的只有 `ucChartBar` / `ucProgressCircular` / `ucTreeMaps` 三件，而且**三件的根因互不相同** ⇒ 拆成 B29 三条分别开工。§40 L2674–2675：`run_tests.ps1`
  只注册了主 vbp `Charts2020`，**5 个 UC 子工程完全不在门禁**。

### B15 非 `-g` 构建下关闭路径偶发 AV（约 16 次 1 次）
出处 §42 L2752–2759。`code=0xc0000005 / at rva=0x15a271 / av write target=0x0`，帧
`#4 0x15a271 #5 0xa9df0 #6 0x5ed33 #7 0x138643`（同形态在 Fix 188 之前是 `rva=0x15a201`）。
复现法记在 `.temp/teardown_av_note.md`（`demo_close_repeat.ps1 -Runs 20` + **同一次构建**的 PDB）。
· **账 #184 之后先重测再决定去留**：本条与 B23（已出）是同一族——被 OS/COM 按 `__stdcall`
  调的过程以前一律以 cdecl 发码。关闭路径正好走 `RemoveWindowSubclass` + 一批 subclass thunk，
  所以这条的 1/16 有可能已经跟着 #184 一起没了；先跑那 20 次，别先动手改代码。
· **2026-10-04 重测完了（照上面那条指示，只测量、没动代码）⇒ 三组各 20 次全 0，但这组读数判别不了**：
  ① 修后 x64（`.build/b23fix64/VBFlexGridDemo.exe`）plain start→WM_CLOSE 20 次：`codes: 0x00000000=20 clean=20 crashfiles=0 stuck=0`；
  ② 修后 x86（`.build/b23fix3`）按「先点 FlexGrid 再关闭」那形 20 次：同样 20/20 rc=0、0 crashfile；
  ③ **修前基线件**（`.temp/demo188/VBFlexGridDemo.exe`，md5 `0595d20e83e8518bde410fa9c70ed60f`，原样拷到 `.build/b15base/` 再跑，原工件没动）
     同 shape 20/20 0，**连当年那条 1/16 用的原夹具**（`.temp/demo_close_repeat.ps1 -Runs 20`，按 stderr 里 grep `C3_CRASH` 计数）也是 20/20 `traces=0`。
  ⇒ 口径：基线本人不复现 ⇒ 「0/20」**既不能记给 #184，也不能证明这条已经没了**（p=1/16 时 20 次全绿的概率是 (15/16)^20≈0.28，本来就不够判）。
  ⚠ 另有一条分析侧的订正，比读数更要紧：**#184 那一刀在 x64 没有字节后果**（MSVC 在 x64 忽略 `__stdcall`，桩只是 cdecl 转发），
     而 B15 的原始现场正是 x64（`base=00007FF6…`）⇒ 本来就**不该指望** #184 收掉它；「先重测」这一步的价值是把这条期望判死。
  下一轮的抓手（别再靠加大抽样次数）：`c3_crash.txt` 那份现场里 `#10 = 应用帧 ← USER32 ← COMCTL32+0x2CED8 ← COMCTL32+0x2CBD4` ⇒ 关闭期有一次
     消息派发进了应用码，应用码里对 **NULL 解引用写**。同族的 B13（两套子系统共用属性名 `VB6_OrigProc` ⇒ 能把 NULL 写进 `GWLP_WNDPROC`）
     机制不同但同一张桌子：B13 是「把 NULL 存成 wndproc」，本条是「拿着 NULL 写」。⇒ 先做 B13（有确定修法：独立属性名 + 取到 NULL 就不写的守卫），
     再做一次抽样，看这条有没有跟着少一个候选。工件都留着：`.build/b15base/`（基线件副本）、`.build/b23_close.ps1`（这次修好了退出码取法：
     `Start-Process -PassThru` 拿不到 `ExitCode`，改成 `[Diagnostics.Process]::Start($psi)`；顺带记一条——`ProcessStartInfo` 在 .NET Framework 里**没有**
     `RedirectStandardErrorFileName` 这个属性，赋值会抛 PropertyAssignmentException，要落文件只能自己读 `StandardError`）。
· **2026-10-04 再补一行（账 #185 已出 ⇒ 上一条指的「先做 B13」这一步做完了，但没收到本条头上）**：#185 收掉的是「事件层被自绘层整层挤掉」那一族，
  本条的现场是应用码**对 NULL 解引用写**（不是把 NULL 存成 wndproc，也不是没装），机制不同 ⇒ **#185 不构成对本条的解释**。
  下一轮别再抽行了（累计 0/60，基线本人不复现）：要么用**同一次构建**的 map/pdb 把 `rva=0x15a271` 那一格钉成函数名（nearest-symbol 不算，见记忆里那条），
  要么就按「抽样判不了」长期挂起。

### B16 Date 可见性的三条**故意不做**的缺口（Fix 175 的边界）
出处 §24 L1666–1672，已复核三条**全部仍未实现**：
① 类模块 `Private d As Date` 需要与 `knownDateVars_` 平行的 `classDateMembers_`（`grep classDateMembers_ src` = 0）；
② 局部 `Const X As Date` 未登记（`src/backend/decl/cgen_localdecl.cpp` 一带）；
③ **纯时间** Date（`x<1`，如 `CStr(0.5)`）仍输出 `1899-12-30 12:00:00`，VB6 应只给 `12:00:00`。

### B17 frx 侧的两条静默失败
出处 §26 L1771–1780。① `FrxReader::load` 失败静默（`lastError_` 无消费者）→ 应发 warning；
② **带 GUID 头**的 frx 条目（`0x0344` = VBFlexGrid 的 `WallPaper`；以及 `FormatString = "…frx":0000`
这类"字符串存 frx"）按 `headerSize=12` 读出垃圾 `imgSize` → 返回空。属另一类问题，与 Fix 182 无关。

### B18 库限定名末段不在白名单时仍命中 `Vb` 前缀→Long 兜底
出处 §10 L701。`VBA.Collection` 这类还会命中 `src/backend/cgen_base_type.cpp` 自己的 `Vb` 前缀启发式 →
判成 `Long`（本 demo 无此类取值，所以"未动"）。同类隐患的唯一记录，判据见 C 组第 6 条。

### B20 `Byte` 数组元素直接进 `If` 比较不成立
出处 §44 L2809–2811。`If buf(3) = 65` 走 Variant 通道（`vb6_VariantFromValue` + `vb6_VarCmpLongEq`）
且**不成立**，中间转一次 `Long` 就对。目前只在 `tests/test_asany_subscript.bas` 的注释里点名，未动。

### B21 两条仍在推进中的族（无 Fix 号，任务列表里是 #9 / #12 / #19 的尾巴）
- VARIANT ↔ typed 转换族（含 160-F 的重估结论）。
- Extender / Ambient 成员访问器的剩余小簇（Fix 161/162/183 之后仍有零星无赋值路径）。

### B24 OLE 拖放的 `hdrop` 旁路日志没人钉
出处 = 账 #175 的跨门工件对形 #300→#301：`oledd_test.txt`(s4) 从 `hdrop FAIL hr=0x1` 变成
`hdrop first=C:\a.txt` —— 两边都不红，因为 `run_tests.ps1` 里 grep 不到 `hdrop`/`oledd`，
它是夹具自己写的日志文件。首拖成败正是这类测试最容易漂的地方 ⇒ 值得挑一轮把它翻成针面。
### B25 账 #179 = 已出（提交 `30a3f4ff`，门 #303）；剩下半条并入 #159
· 已修：容器直调控件公共成员时的宿主上下文 —— 发码在每个 .ctl 实例方法体首 `vb6_UC_PushInstance((void*)me)`、统一出口尾 `vb6_UC_PopInstance()`（前提：`Exit Function/Sub/Property` 早已走 `vb6_proc_exit`，所以配对不漏弹 —— 这也判掉了"路 B 要等 T31-A"那条顾虑）。实测两侧读数由 `mode=3 / mode=1` 变成同值，针面 `U-CTX=True`。
· 剩下半条 = `UserControl.hWnd` 在控件代码里读不回真值，**归 #159**（hwnd 被当对象装箱）。**订正 `30a3f4ff` 提交信息里那句"比较器没问题、值真的是 0"**：`vb6_VarCmpLongNe` 的第一形参是 `vb6_VARIANT*`（vb6rtl.c:593），发码传的是 `&vb6_UserControl_hWnd`（`void*` 全局的地址）⇒ `<>0` 与 `=0` 走同一个被 reinterpret 的槽位，**这对探针判别不了**。
· 那条下一步已跑完（换了探针形态：直接在 `vb6_uc_push` 里打 `r->hwnd` 与全局，两边同值且非 0）⇒ **值是实的，坏的是读法**。**#159 已出**（提交 `ef260daf`，门 #304）：类型 oracle 与那三份成员名单收成一张表 `kHostPseudoRows`，`UserControl.hWnd` 现在发 `(-(vb6_UserControl_hWnd != 0))`而不是 `vb6_VarCmpLongNe(&vb6_UserControl_hWnd, 0)`。本条到此结，剩下的三件表外欠账另立 B26。

### B26 宿主伪成员那张表**外**欠的三件（账 #159 顺带量到）
· `PropertyPage` 一族全局（`vb6_PropertyPage_hwnd` / `_hWnd` / `_ScaleMode` / `_ScaleHeight` / `_Changed`，
定义在 `vb6rtl_com.c:666-670`）**运行期从没有人填值** —— 全是初值 NULL / 1 / 0 / 0，两个 `hwnd`/`hWnd`
拼写并存也只是为了迁就发码。⇒ .pag 代码里 `PropertyPage.hWnd` 恒 0、`Changed` 读写都不落到页上。
**但这一条实测下来不是缺陷, 是产品形态**：driver 把 `.pag` 只当「类模块 + 设计器模块」编进产物（`driver_frontend.cpp:155-170`），RTL 侧**没有任何运行期实例化属性页的人**（`grep PropertyPage src/rtl` 只剩那五行定义），而 VB6 里属性页本来就只由设计器/属性浏览器承载 —— Standard EXE 运行期那些值恒为初值**与 VB6 一致**。⇒ 那张表要保的是它**编得过 + 类型答对**（#174 与 #159 已各自结掉），**不要**去「填」这五个全局。真要做属性浏览器（`PropertyBag` / `IPerPropertyBrowsing`）时再建 .pag 上下文，那时才用得上 #179 那套 Push/Pop。
· 那张表刻意不收的成员 = RTL 根本没声明的人：`vb6_UserControl_Left/Top`（只有 `vb6_Extender_Left/Top`）、
`vb6_PropertyPage_ScaleWidth`（只有 `ScaleHeight`）、`Appearance` / `BorderStyle`（旧 `kNumericHostMembers`
里躺着，永远不可能命中，因为写它们就是 C2065）。要这些名字得先在 RTL 补声明与填值，别在表里挂空名
（哨兵 `ROW-SYMBOL-MISSING` 就是拦这个的）。
· 普查读数：表外手写的宿主符号点 **24 处** —— `cgen_form.cpp:231-245`（按实例 `#undef/#define`，rev20 那套）、
`cgen_base.cpp:61-148`（`Parent.<成员>` 链改写）、`cgen_with.cpp:650-652`（「`With <伪对象>` = 它的 HWND」
这条规则，里面还写死了 `vb6_PropertyPage_hwnd` 那一格拼写）、`cgen_util_type.cpp:824-828`（RTL 宿主**方法**
的形参类型表，与成员值面两回事）、`:864`（认 `vb6_Ambient_DisplayName` 是 BSTR —— 这一处现在表里已答
String 档，下一轮可以试着让 `rewriteByteArrayValue` 改问表）。前四类是**规则**不是成员清单，不动它；
最后一处是真正的第二份口径。

### B27 czUI 的 x64 产物启动期 AV —— **判掉：源码形状限制，不改编译器**
出处 = 账 #180 的真跑面（A/B 已证与那一刀无关：拿改前的编译器编同一份夹具，退出码一模一样
`0xC000041D`）。这一轮把它量到底了，结论是**这条不该由编译器修**：

· 现场（`C3_CRASH_TRACE=1` + `-g` 的 .map 符号化）：AV 落在 `gdiplus.dll+0xF2E1`，读
`0x14FE1F88` —— 一个 32 位量级的地址。栈：`vb6_form_create_frmDemo+0x1027` ←
`vb6_czControl_prop_let_BackColor` ← …，全在 GDI+ 调用上。
· 根因在**工程的 VB 源码**：`tests/czUI-main` 的 Declare 把指针一律写成 `As Long` ——
`GdipCreateFromHDC(ByVal hDC As Long, ByRef graphics As Long)`（被调方把 64 位指针写进
4 字节槽，高半截落到栈上别处）、`GdipDeleteGraphics(ByVal graphics As Long)`（把截断值
递回给 API）、`GdiplusStartup(ByRef token As Long, ByRef inputbuf As Any, …)`。这类声明
**62 行**，是 VB6 只在 32 位跑留下的形状。⇒ 要 x64 就得先把这些声明改成 `LongPtr`
（**改 VB 代码的活**，与 tB 的口径一致：它也是要求源码用 LongPtr，而不是把 Long 偷偷加宽），
不是往 RTL/发码里塞补偿。VB6 工程引用了 32 位 OCX 时同理，别去试 x64。
· 所以本线口径：**czUI 只在 x86 那格钉**（CI 现状 `Test-GuiVbp "czUI" -Arch "x86"` 就是对的），
x64 那档不补用例、不补产物；将来若要把 x64 立成目标，先改夹具源码，再谈别的。
· 顺带留一条**独立**的观察（不在本条结案范围）：崩溃轨迹器 `vb6_CrashTraceVEH`
（`vb6rtl.c:82`）在爆栈现场会**重入** —— 那份 c3_crash.txt 里同一递归栈打了 4 段，第一现场
（gdiplus 那段）被压在最后。诊断工具一重入就把现场盖掉，值得单独挑一轮加个"每进程只记一次"
的闸；但那是诊断面的质量，不影响任何产物行为。

### B22 隐式函数声明（C4013）现在没有**守卫**，只有 census
出处 = 账 #173 / #174。两格都已出（#173 补 `vb6com_internal.h` 的 `extern double vb6_VariantToDouble(VARIANT)`；#174 把裸名 `SelectedControls` 接进 `kPropertyPageHostMembers`），`tests/Charts 2020` 整个构建的 C4013 从 **14 → 0**。但**没有任何东西阻止它再长回来**：
· 为什么必须当缺陷：x86 cdecl 下被隐式声明的函数按 `int` 取返回值，而 `double` 返回值躺在 x87 栈 ST0 上、调用方永不 `fstp` ⇒ 每调一次漏一层栈，八层之后栈满、之后任何浮点取值得 QNaN `0x7FF8...`（#173 的炸法）。x64 走 XMM0，全静默。
· 现成的收口办法：对 RTL 源加 `/we4013`（`src/backend/msvc_driver.cpp` 那一条 `cmd << " /W3"` 旁边）。**前提**是生成码侧也零 C4013 —— 生成码的雷由 #174 那格清了，但只清了这一个名字，未解析裸名的**兜底仍然是发裸名**（`cgen_expr_ident_builtin.inc` 尾部的 Fix 110u 一族），别的工程换个名字就又会漏。
· 所以顺序建议：先量「语料里还有没有别的未解析裸名调用」（`--emit-c` 全语料跑一遍 cl 数 C4013，数法见记忆库「数 cl 的警告必须自己重跑 cl」），再决定是上 `/we4013` 还是在 cgen 侧把未解析裸名**判死**（后者才是单一权威，但要先确认不会把「隐式 Variant 局部」那条兜底一起打掉 —— 它就是 Fix 110u 立着的理由）。
### B28 `run_tests.ps1` 的 PASS/TOTAL 不同源（门的判据没受影响，账面会误导）
出处 = 门 #311 的 11 份 job 日志（副本在 `.build/gate311/`）：vbp 分片 1 / 2 / 4 各打 `PASS=TOTAL+1`（43/42、46/45、43/42），
asm 片打 `PASS=13 SKIP=0 TOTAL=14`（差 1），其余七片自洽。本轮 11 片 `FAIL=0` ⇒ 门是绿的，这条只动**账面**。
后果：任何写成「PASS == TOTAL」的自洽式检查在这里都会假红/假绿 ⇒ 门的判据只看 `FAIL=0` 与工件行。
下一轮动 `tests/` 时顺手把计数收成一处口径（别为它单开一轮门）。

### B29 三件 Charts UC 子工程编不过的根因（2026-10-04 一次真编译量到的，每条都带生成码原文）
出处 = `.build/b187out/x64_*/c3-error.log` 与那份临时生成码（`%TEMP%\C3C\...`）。三件互相独立，别并成一刀。
· **① `ucChartBar` + `ucProgressCircular`：设计期字体块的 Size 被打成非法浮点字面量** —— 生成码原文
  `_vb6_f119->Size = 12f;`（MSVC `error C2059: 语法错误:"数字上的错误后缀"`；Form2.c 里 4 处、Form1.c 里 5 处）。
  发码处 = `src/backend/detail/module/cgen_form_create_controls.inc:116-118`：`snprintf(szBuf, "%.4g", sz)` 之后直接
  `+ "f"` ⇒ **整数值**（12 / 9 / 10）出来就是 `12f` 这种非法 token。同文件 830 行那条路用 `std::to_string(fSize143)`
  所以没事 —— 两条路两套口径，正是这类的常态。正确修法不是就地补小数点：字面量的形状该由**一个**出口负责
  （`cgen_expr.cpp:15` 的 `floatingLiteral` 已经保证「没有 . 或 eE 就补 .0」，把 93 行那句 `+ "f"` 与这里的 `%.4g` 一起收进它，
  再加一条 census 哨兵：`src/backend` 里任何 `+ "f"` 前面必须是 `floatingLiteral` 的结果）。
  **→ ① 已出（提交 `33cb239a` + 收口红 `936124e8`，门 #313（run 37198483318，head `936124e8`，attempt 1）= 11 job 全绿、11 片 `FAIL=0`）**。
  **中间门 #312 红过一次，红因是我这一刀**：我把三处 `std::to_string(v)+"f"` 一起"收进权威"了，而那三处**本来就合法**（to_string 必带 6 位小数）—— 其中设计期 FontSize 那条的字面形状正被 `sl_emitc_native`（vbp#4 片，tests/ctrlslider）钉着 ⇒ 变成 `20.0f` 就照红。
  ⇒ 退回：权威头再加一个 `floatFixed6Literal`（= to_string + 后缀，形状规则仍在同一处），三处改调它 ⇒ 发码字节与 HEAD 一字不差；只有真正坏掉的 `%.4g` 那处走 `floatSingleLiteral(sz, 4)`。
  **工具事实（这次学到的，别再犯）：A/B 的工程清单不能手挑** —— 本仓门禁在发码形状层面的断言面实测有 30 个文件（`.build/b188_surface.txt`，由 `b188_surface.py` 从 run_tests.ps1 的 `Test-EmitcShape`/`Test-CodegenNote` 里推出来）。我第一版手挑 8 个工程 ⇒ 本地全绿、门上一片红。加宽版（37 × 两架构 = 74 次发射）现在只有 ucChartBar 两片被改、数值全等、增删 0 行。

  改法按上面那条口径做: 新增 `src/common/float_literal.hpp` 作为**唯一出口** (`floatingLiteralText` / `floatSingleLiteral`，经典 locale + 「没有 `.` 或 `eE` 就补 `.0」`)，把五处手拼点全接过去 (设计期字体块、设计期 FontSize、属性袋 VT_R4、IFont 袋、语义层 Fix 133z 那段手写 `%.9g`+补点+拼 f)。字体块位数照旧 4 位 ⇒ 8.25 的写法一个字都不变。
  配对读数 (同一份工程、两台编译器、x86): `ucChartBar` 的 C2059 **4 → 0** (只剩本条②的 1 条 C2065)，`ucProgressCircular` 的 C2059 **16 → 0** (剩③的 C2084/C2065/C2198)。
  判据两面: 发码形状针 `fontsize_literal` 钉在**真实工程**上 (Needles `->Size = 12.0f;` + 证人 `->Size = 8.25f;`，Absent `->Size = 12f;` 就是改前的形状)；哨兵 `scripts/check_float_literal_shape.ps1` (R1 手拼后缀=0 / R2 权威头带着补点与经典 locale / R3 调用点>=4) 在缺这刀的树上 R1~R3 全红，R1 当场点出那 6 个手拼点。护栏 `--emit-c` A/B (BASE = 已发布的 `a4e3b574` 冷编) 8 工程 × 两架构 = 14 份逐字节相同，只有 ucChartBar 两片被改，且归一化浮点 token 后逐行相等、数值多重集相等、增删 0 行。

· **② `ucChartBar`：宿主 UC 实例的句柄名发成了没声明的标识符** —— 生成码原文
  `int32_t i_end = (vb6_ComGetIntProp(vb6_hwnd_ucChartBar1, L"Count") - 1);`（`error C2065`）。
  `ucChartBar1` 是**控件数组里那一枚实例**，而实例句柄住在 `vb6_arr_ucChartBar` 里 ⇒ 与账 #157 那条同族
  （拼 `vb6_hwnd_<名>` 而不是走 `ctrlHwndExprForInit` / `vb6_arr_*` 那条既有出口）。修法 = 把这一处也改读那个唯一出口。
  **→ ② 已出（提交 `ea5155af`，门 #314（run 37200966607，head `0efb29b0`，attempt 1）= 11 job 全绿、10 片 `FAIL=0`）**。两条订正，开工前先看这两句：**①根因不是「实例句柄拼错名」** —— 数组实例的句柄表达式本来就走 `vb6_CtrlArr_GetAt(&vb6_arr_<名>, i)`、设计期下发那一路一直是好的；坏的是**整体成员** `Count` 少了一条分支，于是被当成「控件属性」去 COM 兜底，而兜底那一条自己拼了 `vb6_hwnd_<数组名>`（数组压根没有这个变量）。**②「同窗体就编得过」那一半更危险**：同一窗体内 `vb6_hwnd_<名>` 若因为别的原因存在，它编得过、`Count` 恒答 0 —— 只看编不编得过会把这当成修好了。
· **③ `ucProgressCircular`：同一个过程发了两遍体** —— `Form1.c(10)` 与 `Form1.c(13)` 报
  `error C2084: 函数"void vb6_Fo..."已有主体` + `C2198`（实参数不符）⇒ 某个 `Form_*` 事件在两份名单里各发一次。
  这条要先把两份名单找出来（`grep` 发事件 thunk 的那两处），再定哪一份是权威。
· **④ 顺带一条诊断面的事实**（不算缺陷，记着省时间）：`.pag`/`.frm` 报错的**行号不含设计期头块** ——
  `PropPagFMR.pag(594,17)` 实际指向文件第 720 行（720 - 126 = 594，前 126 行是 `VERSION` + Begin/End 控件块）。
  按报的行号去找源码会一无所获 ⇒ 要么按 `文件行 = 报的行 + 头块行数` 折回去，要么改成报真实行号。

· **⑤ 事件臂的处理器函数名按「控件设计期拼写」现拼，而过程定义按 Sub 自己的拼名**（账 #190）—— `ucChartBar` 的 C2065 清掉之后 x86 真编译走到链接期才红：`Form1.obj : error LNK2019 无法解析的外部符号 _vb6_Form1_cboLabelsPositions_Click`、`_vb6_Form1_ChkAxisY_Click`（`fatal LNK1120: 2 个`），物证 `.build/b189new/ucChartBar/c3-error.log`。产物里逐行对上：存在性查询 `symTab_.lookup(ctrl.controlName + "_Click")` **大小写无关**、命中了；发出去的名字却是 `cProcName(ctrl.controlName + "_Click")` —— 拿控件名现拼。而定义那一侧用符号自己的名字。Form1.frm 里恰好两枚不同（控件 `ChkAxisY` / 过程 `ChkAxisy_Click`；控件 `cboLabelsPositions` / 过程 `CboLabelsPositions_Click`）⇒ C 大小写敏感 = 两个符号。普查：`cProcName(ctrl.controlName + "_…")` 那一形 **27 处**（`cgen_form_wndproc_create.inc` 18 / `…_dispatch.inc` 6 / `cgen_form_ctrl_style_apply.inc` 1），而同一文件里 501/523/544 三处已经是**对的**写法（`cProcName(clickSym->name)`）⇒ 口径本来就有，只是没收成一处权威。修法 = 一枚 `eventHandlerFn(ctrlName, suffix)`：lookup 一次，命中就交 `cProcName(sym->name, sym->accessLevel)`，没命中交空串让调用方**不装这条臂**；配 census 哨兵禁掉「现拼」那一形。动手前先按门 #312 那条教训查一遍形状断言面 （`.build/b188_surface.txt`）里有没有针钉的是控件拼写那一版函数名。
  **→ ⑤ 已出（提交 `a27758c3`，门 #315（run 37202837330，head `a159b85d`，attempt 1）= 11 job 全绿、10 片 `FAIL=0`）**。两句留给后人：①普查必须按**形状**（`cProcName(<任意> + "_`）而不是按变量名 —— 我第一遍只扫了两个名字，漏了 `info.ctrlName` 那 16 处（焦点/鼠标/Validate 一整批）和菜单 2 处，"27 处"这个数是错的，真数 43；②这一族的红**只在链接期现形**（`LNK2019`），所以判据里必须有一次**真编译真链接**，`--emit-c` 的形状针 + 语义层的存在性检查都拦不住它。顺带一条好消息：这一刀之后 ucChartBar **第一次产出 exe** ⇒ #187 的升格清单多一件，剩下的独立缺陷是 B29⑥（数组元素的成员访问）。
· **⑥ 控件数组元素的成员访问是另一条通路，且现在是坏的**（账 #191）—— 探针夹具（`.build/b189ve_probe`，ve_units 的副本）x86 真跑 `.build/b189probe.out`：`U-ARR-RAW count=3 lb=0 ub=2`（②那三个数是对的）之后 `U-ARR-M //` —— `uArr(0).SW()`/`uArr(1).SW()`/`uArr(2).SW()` **三枚全空**，下一行读 `uArr(0).Left` **段错误**（rc=139）。发码形状（`vb6_CtrlArr_GetAt(&vb6_arr_uArr, aj)` 直接进 `vb6_ComCall(..., L"SW", …)`）说明运行期把数组元素当成了 IDispatch 兜底，而设计期那一路是 `vb6_UC_InstanceOf(vb6_CtrlArr_GetAt(...))` 再打生成的 `prop_let` —— **少包的就是 `vb6_UC_InstanceOf` 那一层**（待量，别当结论用）。真工程全靠这一形：`ucChartBar1(i).AddSerie …` / `.Move …`。账 #189 的判据**刻意没碰**这一形（只钉整体成员三个数 + 一圈计数），两条缺陷不要互相掩盖。
  **→ ⑥ 已出（提交 `f5b4a03b`，门 #317（run 37208673511，head ea3eeeec，attempt 1）= 11 job 全绿、10 片 FAIL=0）**。修法 = Fix 112 那段从发码点整体抽成 `CCodeGen::emitUcInstanceMemberExpr`（`src/backend/cgen_util_classcall.cpp`），单枚交 `vb6_hwnd_<名>`、数组元素交 `vb6_CtrlArr_GetAt(&vb6_arr_<名>, i)` —— 两条路只差这一个串；元素那条（`cgen_expr_member_precheck.inc`）补上这次调用，放在控件属性表与 COM 兜底**之前**。**A/B 又抓出第二条独立的坑（左值侧，与直发本身不是一件事）**：元素改道之后 `ucChartBar1(i).Font.Size = ...` 落进 `tryRewriteCOMLvalue` 的 Pattern C/D2，而那段是**字符串级**改写 —— 只取 `prop_get_` 那一段、尾巴 `.Size` 整段丢掉 ⇒ 发成 `prop_let_Font(elem, <double>)`，成员名没了、值实参的槽还是 `vb6_ComIface_Font*`，实测 Form2.c 471-477 四条 C2440，把刚编得过的 ucChartBar 打回编不过。修法 = 在折叠**之前**给"带尾巴的 prop_get_ 目标"一条独立落点：取回那枚对象再对它写一层（StdFont 本身是真 IDispatch）⇒ `vb6_ComSetProp(vb6_X_prop_get_Font(直发的实例), L"Size", pack(值))`；认不得的尾巴一律 `return false` 让通用路径发成非法左值、当场红 —— 静默丢成员名比编不过更坏。**判据三面**：`tests/ve_units` 真跑证人 `U-ARRM-methods=3`（x64 与 x86 都跑到 `U-DONE`，既有 `U-ARR=True` 不动）+ 发码形状针 `ucarr_member_call`（Absent=改前那条形）与 `ucobj_chain_write`（钉在**真实工程** ucChartBar 的 Form2.frm:223 上）+ 哨兵 `scripts/check_uc_instance_exit.ps1`（U1 expr/ 手拼 InstanceOf=0 / U2 权威三样齐 / U3 两条路都在调它 / U4 带尾巴的左值必须有独立落点且排在折叠之前）；假 needle 验红做过 —— 把 `(Pattern C-tail)` 改名 ⇒ U4 当场点名，还原回绿。**真编译配对读数（x86）**：ucChartBar **rc=0 且产出 exe**；ucProgressCircular 仍红，但错误类严格是基线的**子集**（16xC2059 + 2xC2065 + 1xC2084 + 25xC2198 -> 1xC2065 + 1xC2084 + 12xC2198；C2059 归 #188，剩下两条正是 B29③④，不是这一刀）。**护栏**：A/B = 40 工程 x 两架构（80 份快照，BASE=`4d1dbf45`）=> 只有 6 份被改（ucChartBar 两片各 8 对、ucProgressCircular 两片各 1 对，ve_units 两片是手改夹具那两份），每条差异行逐行打印归因、三形都属"数组元素的成员访问从 COM 兜底挪到 UC 实例直发"，其余逐字节相同、OFFENDERS 0。**这一刀没修好的那半**：extender 属性（`uArr(i).Left` 那形仍然段错误）另立**账 #193**；另外"单枚那条链"在整面 80 份快照里**没有一处**走带尾巴的写（基线 172 处 Pattern C/D2 全是空尾巴）=> 那条形状针欠着，别以为两边都钉过了。
· **⑦ `ucTreeMaps`：#186 把 parser 那一格清掉之后，红移到了 cl 阶段**（账 #192）—— `--emit-c` 现在零诊断（rc=0），真编译 x86 报 `PropPagFMR.c` 里 **10× C2039 + 2× C2059 + 1× C2198**（物证 `.build/b186tm/c3-error.log`）。两族根因不同，别并成一刀：**①** `hDC` / `CurrentX` / `CurrentY` / `ScaleMode` / `TextHeight` / `Clear` 全部报 `"X" 不是 "HWND__" 的成员` —— 伪成员/方法被发成 HWND 的结构体成员访问，与账 #159（宿主伪成员收成一张表）、账 #174（PropertyPage 的 SelectedControls 裸名调用）**同源**：那张表里 PropertyPage 这一型没登全。开工先读 `scripts/check_host_pseudo_table*.ps1` 两条既有哨兵，把缺的名一次性补进那张表，**别在发码侧再加一层 if**。**②** `C2059 语法错误 ";"`（两处，发空了的语句）+ `C2198 vb6_di_TextOutW 用于调用的参数太少`（Declare 桩参数数不对 —— 记忆库口径：手写一枚 `vb6_di_` 单桩，别重跑 `gen_di_stubs`）。另记一条工具事实：**ucTreeMaps 不在门禁里**（#187 的升格还没做），所以这格的读数只能靠本地真编译，别假设 CI 会跑到它。
  **→ ⑦ 的两条落点已经量到（2026-10-04，x86 真编译，物证 `.build/b192new/ucTreeMaps/c3-error.log` = 11xC2039 + 2xC2059 + 1xC2198），并且 ⑦ 原来那句"再补一张表"要订正**。①**C2039 那一族不是"PropertyPage 伪成员没登记"，是 With 块里那枚控件**：源 = `PropPagFMR.pag:256-266` 的 `With Picture1`（VB.PictureBox，用 `.hDC / .CurrentX / .ScaleMode / .CurrentY / .TextHeight`）与 `pag:618-619` 的 `With lstFonts`（VB.ListBox，`Call .Clear`）。With 前言把 FormControl 那档发成 `HWND _vb6_with_N = (HWND)vb6_hwnd_<控件>`（`cgen_with.cpp:134-143` + `:766`），而成员那一路（`cgen_expr_with.cpp:25-90`）**从不查宿主伪成员表** —— 表在 `cgen_util_com.cpp:469-540`（`kHostPseudoRows`，键 = 宿主种类 usercontrol/propertypage/extender/ambient + 成员名，`hostPseudoFind` :542），它只试 `getControlPropReadFn`（`cgen_util_ctrl.cpp:247`，里面有 scalewidth/scaleheight 但没有 hDC/CurrentX/CurrentY/ScaleMode/TextHeight），全落空就撞上 `cgen_expr_with.cpp:88` 那条兜底 `lastExpr_ = tempVar + "." + cIdent(member)` ⇒ 拿 HWND 当结构体拼成员。`Clear` 同形：零实参方法表 `controlZeroArgMethod`（`cgen_util_ctrl.cpp:1414-1445`）只登记了 setfocus/clearsel。⇒ 这一格真正的问题是「With 里那枚控件的宿主/复刻成员从哪张表读」，与 #159 那张表同源但**键的维度不同**（表按宿主种类，With 这条路按控件类型）=> 修法要把两处收成一处口径，不是在 With 里再加一条 if。②**C2059 那两条是独立的小事**：`Optional ... = 0&` 的默认值文本被原样拼进 C —— `evalOptionalDefault`（`src/semantics/semantic_analyzer_util.cpp:382`）是字面量发码权威 `visit(LiteralExpr)`（`src/backend/expr/cgen_expr.cpp:52-93`，Long 走 `0L`）之外**另写的一份**：Long/Integer 那支直接 return `lit->rawText`（:390-392）=> 词法留在 rawText 里的 VB 类型后缀 `&`/`%` 漏进 C（Double 那支倒是剥了 `!`/`#`，:425）。所有 `_has_*` 补参的消费者（`cgen_call.cpp:544`、`cgen_expr_call_opt_pad.inc:11`、`cgen_util_classcall.cpp:606` 等）共用这一个串 => **一处修，八处齐**。③C2198 `vb6_di_TextOutW` 参数太少 = Declare 单桩那一族（照既有口径手写一枚 `vb6_di_` 桩，**别**重跑 `gen_di_stubs`——它按当次 session 整文件重写会撤掉别人的桩）。
  **→ ⑦ 的开工家底（同一轮只读盘点，file:line 都核过；结论 = 这一格要拆两半，别当一刀做）**。**已有、只差接线**：①`vb6_GetCurrentX/vb6_SetCurrentX/vb6_GetCurrentY/vb6_SetCurrentY` 已在 `src/rtl/core/vb6forms/vb6forms_widget_prop.c:268/279/286/297`（声明 `vb6forms_prop_form.h:79-82`），**按 HWND 存窗口属性 `VB6_CurrentX/Y` ⇒ 任何控件都能用**，可 `src/backend/` 里**零引用** = 全没接；②ScaleMode 的 getter 同在 `vb6forms_widget_prop.c:256`，同样没接；③`vb6_ClearList` 在 `vb6forms_list.c:114`（LB_RESETCONTENT），而 `controlZeroArgMethod`（`cgen_util_ctrl.cpp:1415-1447`）只登记了 setfocus/clearsel 两条 ⇒ **加一行 `clear` 就把 With 那条与语句那条两个头一起接上**（两条头各自 key 在 `knownFormControls_` 上：`stmt/cgen_call.cpp:359-370` 与 `detail/expr/cgen_expr_call_callee_withm.inc:306-355`）。**还不存在、要新造**：④控件的 `.hDC` 没有任何导出口（`_Paint` 那套两层规则已经在 `cgen_form_wndproc_subclass.inc:376-395`：BeginPaint → SetPropW "VB6_PaintDC" → RemoveProp，消费侧 `vb6forms_ctrl.c:674-677` 带 GetDC 兜底 —— 但它现在是 static，`.hDC` 正好复用这一条，不用新发明）；⑤控件/窗体的 `TextHeight`/`TextWidth` 只有 UserControl 版（`vb6rtl_com.c:731/744`、`uc_host.c:882/886`），没有按 HWND 的；⑥带 HWND 参数的 `ScaleX`/`ScaleY` 也没有（只有 `vb6_UserControl_ScaleX/Y`，`vb6rtl_com.c:804/810`，无 HWND），但换算的零件是现成的：`vb6_ScalePxToUser`/`vb6_ScaleUnitsPerPx`/`vb6_WindowScaleModeSelf`（`vb6forms_window.h:58-67`）。**表的缺口**：`kHostPseudoRows` 里没有 picturebox/image/listbox/form 那一档，`CurrentX`/`CurrentY` 在任何一档都没有；`propertypage` 那档只有 changed/hwnd/scaleheight/scalemode/selectedcontrols 五条。⇒ 开工顺序建议：先把「With 那枚控件的接收者认回控件身份」这一条接通（它一处解三形：CurrentX/CurrentY/ScaleMode + Clear），再单独立一格做控件的 DC 与文本量度（那是要往 RTL 加功能的）。
  **→ ⑦ 的一半已出（提交 `a11a7da8`，门 #320（run 37228796120，head 42e4338f，attempt 1）= 11 job 全绿、9 片日志各自 FAIL=0，其中 `[VBP] piccur` / `piccur_x86` / `with_ctrl_cursor_clear` 三行在 CI 上真 PASS）= 账 #192 的接线那半**。读数是自己复算的：`_vb6_with_0 = (HWND)vb6_hwnd_Picture1`、`_vb6_with_11 = (HWND)vb6_hwnd_lstFonts` ⇒ 撞兜底的接收者就是**一枚普通控件被 With 包住**，不是 PropertyPage 宿主（⑦ 原来那条"宿主伪成员"的定性只对了一半：表是按宿主种类建的，而 With 这条路本来就先查 `getControlPropReadFn` 再查 `controlZeroArgMethod` —— 两张表都齐，**只是档位里没有那几行**）。改 = 补表：PictureBox 的 currentx/currenty（读表 + 写表成对补）、`controlZeroArgMethod` 加 clear（ListBox/ComboBox），三条都接在 RTL 早就有的出口上（`vb6_GetCurrentX/Y`、`vb6_SetCurrentX/Y` 按 HWND 存窗口属性；`vb6_ClearList` = LB_RESETCONTENT）。**刻意不给成通用行**：VB6 的 CurrentX/CurrentY 只在画得上去的那几枚上有，给成通用 ⇒ `List1.CurrentX` 也答 0 = 伪造成功。判据：新夹具 `tests/piccur` 两片三头钉（写了读得回 + **另一枚没被写坏**（RTL 按 HWND 存，写成全局一份就露馅）+ 清空前确实是 2，拦住"Clear 什么都不做、本来就 0"那种假绿）；**负控 = `40bbef67` 那台真编同一份夹具** ⇒ PicForm.c 150/151/160 三条 C2039、rc=1。形状针 `with_ctrl_cursor_clear`（Present 直发出口、Absent 是 `tempVar + "." + 成员名` 那一形）。真工程读数（x86, ucTreeMaps）：C2039 **10→4**（剩 hDC / ScaleMode×2 / TextHeight）+ C2198 1；TextHeight 那行在光标通了之后换了个红法（C2039 → C2440 "BSTR→float"），根因同一个。护栏 A/B = 80 份（BASE = `a7524ab7` 那台）⇒ 被改的产物 0 份。剩下的那一半见 §B31/§B32。


### B30 `3%`（Integer 类型后缀）在词法层就被拒（账 #195，2026-10-04 一次顺带测量量到的）

`Public Sub S(Optional ByVal a As Integer = 3%, Optional ByVal b As Long = 0&)` 报
`error VB1005: 十进制数字超出 64 位整数表示范围` 并带一整片级联（VB2004 / VB2001 / VB2002 / VB2003 共 13 条），
而**同一位置**换成长整型后缀就全过 —— 所以这不是"后缀都不支持"，是 `%` 那一支单独断了。读数 =
`.build/p194/m_*.bas`（`C3.exe <file> --syntax-only`，一形一条）：

| 声明里的默认值形状 | 诊断条数 |
|---|---|
| `Optional ByVal a As Long = 12&` | 0 |
| `Optional ByVal a As Long = &H10&` | 0 |
| `Optional ByVal a As Long = 12` | 0 |
| `Optional ByVal a As Integer = 3%`（就这一形）| 13 |

**上一轮的编译器（`ea3eeeec` = 门 #317 那台）读数完全相同** ⇒ 与账 #194 那一刀无关，是存量缺陷。
发现路径也是它带的：#194 的夹具本来要摆三形后缀，`3%` 那条在 parser 就死了，于是夹具改用小写无后缀的 `3`。

落点候选（**还没证**，开工第一步就是证它）：`src/lexer/lexer_number.cpp` 的类型后缀 switch 里
`case `%`: text += advance(); break;` —— 这一支不置任何标志（`&` 置 isLong、`!`/`#` 置 isFloat），
于是带着 `%` 的串落回"无后缀十进制按数值大小定档"那一段。开工前先把两条测过再定范围：
`Dim y As Integer: y = 3%`（普通赋值位置）与 `Print 3%`，
确认是"整条 `%` 都不支持"还是"只在参数默认值里断"。

**→ §B30 已出（提交 `963c5e71`，门 #319 全绿 —— 明细见 §D 那一行）**。落点候选那条**证完了，方向对但差一处**：不只是参数默认值那一段 —— `y = 3%` / `&HFF%` / `&O17%` / `&B101%` / `-4%` / `3% + 1` / `= 3%` 七形在修前**全红**（`.build/p195_*.bas`，`--syntax-only`），而 `12&` / `&H10^` / 无后缀三形一直通（证人：排除"整条后缀机制坏了"那种误读，也排除"`^`/`&` 一起坏"）。根因是 `case ` + pct + `: text += advance(); break;` 这一支**只吃字符不置标志**（`&` 置 isLong、`!`/`#` 置 isFloat、`^` 置 isLongPtr，四条里只有 ` + pct + ` 没置），加上 radixDigits 的剥离表里也没有 ` + pct + ` —— 于是带着 ` + pct + ` 的串落回"无后缀十进制按数值大小定档"那一段，parseIntLit 看见残留的 ` + pct + ` 就报"超出 64 位"。**一条规则抄了四份**（十进制 + &H + &O + &B）再加一张剥离表，漏的就是其中一份 ⇒ 修法四处一起齐，并配哨兵 `scripts/check_int_suffix_sites.ps1`（S1 置标志/无空支、S2 剥离表同时带 ^ 与 %、S3 四个消费点各认一次 %（并同时数 & —— 四份拷贝一分家就报）、S4 显式 % 的 16 位守卫还在）；假 needle 验红：把 % 支改回空支 + 剥离表去掉 % ⇒ S1/S1/S2/S3 四条一起点名。**边界**：显式 % 超出 -32768..32767 报词法错，不按 int32 收下再让 int16 去截（那是把值改错：2147483648 静默回绕成 -2147483648）—— 配一条负例钉住。顺带把 #194 那枚 optdef 夹具先前为绕开本缺陷改成的无后缀两处（`3%` / `5%`）**恢复原样**，x64 与 x86 都真跑到 OD-DONE。护栏 A/B = 80 份（BASE = `40bbef67` 那台）⇒ 被改的产物 0 份：整面存量里一处 % 都没有（有一条就压根编不过），所以它又只能靠自己的夹具响 —— 与 #188/#194 同一课，连着三条了。

### B31 控件的绘图面缺一半：`.hDC` / `TextHeight` / `TextWidth` / 带 HWND 的 `ScaleX`·`ScaleY`（账 #196）

⑦ 剩下的四条 C2039 全在这里。**不是接线问题** —— 后端两张表今天没有可指的名字，因为 RTL 里没有：

- 控件的 `.hDC`：没有任何导出口。可复用的规则已经有了 —— `_Paint` 那套两层（`cgen_form_wndproc_subclass.inc:376-395`：BeginPaint → SetPropW "VB6_PaintDC" → RemoveProp，消费侧 `vb6forms_ctrl.c:674-677` 带 GetDC 兜底），但它现在是 static。晚绑定那一路的 `hDC → GetDC` 在 `uc/detail/uc_hostmodel_getprop.inc:86`。
- 按 HWND 的 `TextHeight`/`TextWidth`：只有 UserControl 版（`vb6rtl_com.c:731/744` + 静态 `vb6_uc_measureText:693`，按实例走 `uc_host.c:882/886`）。签名可以直接照 `vb6_UC_TextHeightOf(void* inst, BSTR text)` 换 HWND。
- 带 HWND 参数的 `ScaleX`/`ScaleY`：只有 `vb6_UserControl_ScaleX/Y`（`vb6rtl_com.c:804/810`，无 HWND），而裸名那条改写还挂在 `isDesignerModule_` 上（`cgen_expr_ident_builtin.inc:215-232`）⇒ `.pag` 里的裸 `ScaleX(...)` 认不出来。换算的零件是现成的：`vb6_ScalePxToUser` / `vb6_ScaleUnitsPerPx` / `vb6_WindowScaleModeSelf`（`vb6forms_window.h:58-67`，见 §B32 那条 ScaleMode 的坑）。

开工顺序建议（2026-10-05 订正）：hDC 与 TextHeight/TextWidth **两条都已出**（门 #323 / #324），ucTreeMaps 的 UnicodePrint 编译面因此全清 —— 但"整条通了"这句要说得更准：它现在停在链接期缺三枚桩（§B36/#201），而页里的控件压根没被创建（§B34/#199）也还没解。剩下的：带 HWND 的 `ScaleX`/`ScaleY`（与 §B32 那条 ScaleMode 同源，那条已通，所以这一条现在做得对了）、#199 的甲/乙口径（页从来没被创建 ⇒ 页里的绘图面至今白画）、新记的 §B37/#202（Frame 标题带按默认字体量）。另：ucTreeMaps 的第一趟真跑已量到 x64 启动期 AV（新账 §B38/#203，x86 是好的）。再一条今天量出来的新账 §B39/#204（**没被写过字体**的控件整张 Font 面读空、文字量按系统默认字体算 —— 读数是 `bName= bfs=0 bpf=0 bth=16`，改前设计期那张 18pt 现在读得到）：它与 #202 同一处出口的覆盖面，开工顺序上 #204 排在 #202 前面（它有产品后果：默认字体那批控件的排版与文字量现在是错的）。**#202 与 #204 同日都出了**（提交 `8252b080`，门 #328 = completed/success，但逐 job 读数没拿到 —— 见 §D 那一行写清的边界）：出口递出文件、六处站点接上、默认字体那批控件的文字量从 16/240 回到 13/195。同族还欠一条 §B40/#205（状态条那两处被 D11 具名豁免着，动它之前要先读 `tests/ctrlstatusbar` 的判据口径 —— 那是另一位作者的族）。账 #203 归因已落地（工程里 `As Long` 截指针，一枚 .ctl 就 10 处），但它剩下的不是"怎么修"而是**"要不要在这些 vendored 工程上追 x64"** —— 已按既有口径就地判掉（不追，x86 是那一档的目标，见 §B38 那条 census）。

**→ 本轮（2026-10-05）过后订正两句**：① §B31 里那条 `.ScaleMode` 不属于本账，它是 §B32（账 #197）的读表那一半，已随 #197 出掉 —— **ucTreeMaps 的 C2039 实测 4→2**，剩下的两条就是 `hDC` 与 `TextHeight`（+ C2198/C2440 各 1，同根）。② 做 §B32 的时候顺带量出一条**更大的一格**（见 §B34，账 **#199**）：`.pag` 的设计块控件**从来没被创建** —— 同一份 emit 里 PropPagFMR 那一段 `vb6_CreateControl` **0 处**、同工程 Form1 那一段 **14 处**，页里的 `vb6_hwnd_Picture1` 是 `#define ... (*vb6_UC_DesignSlotOf(me, "Picture1"))` 而那个槽位按需新建、初值 NULL（`uc_host.c:641-652`）⇒ 属性页里 `With Picture1` 打的是一枚空句柄。**这条不修，本账剩下的两条做完 UnicodePrint 也还是白画** —— 接下去的开工顺序改成：#199 → #196。

**→ hDC 那一半已出（2026-10-05，提交 `50676a8b`，门 #323 = run 37238929955、head f26f9590、attempt 1，11 job 全绿：10 片各自 FAIL=0（含先前被本机代理拦成 HTML 的 smoke / compile 两片，重下到了）；本轮五条新用例在 CI 上逐行真 PASS —— `[VBP] dcsurf`(vbp#3) / `[VBP] dcsurf_x86`(vbp#4) / `[CODEGEN-NOTE] dcsurf_dc_shape`(syntax) / `[CODEGEN-NOTE] dc_read_real`(syntax) / `[STATIC] control_dc`(compile 片，那一片共 13 行 [STATIC]）**：口径**不新写**。RTL 里「这枚控件的绘图 DC 从哪儿来」早就有一处（`vb6forms_ctrl.c` 的 Print/Cls 在用那两条），这一刀把它抽成 `vb6_ControlDrawDC` 让 `.hDC` 也走它，再补上「句柄交得回 VB 代码」那一半 —— `intptr_t vb6_GetControlHDC(void* hwnd)`：派发期那张直接交出去（既不缓存也不 ReleaseDC），否则按 HWND 缓存进窗口属性 `VB6_ObjectDC`（**VB6 是一个对象一张**，反复读必须读回同一个值；不缓存就是每读一次漏一张），白拿的那张当场 `ReleaseDC`；归还点在 PictureBox/Image 那层自己的 `WM_DESTROY`（`vb6forms_picture_prop.c:307-308`，#185 那套分层槽位）。后端两张表各补一行 `hdc`（PictureBox 与 Form，**刻意不给通用行** —— 与 #192 那条 CurrentX 同一个道理：`List1.hDC` 答一个数就是伪造成功）。RTL 动过 ⇒ touch `c3rtl.rc` 重编 C3.exe（#156 那条旧资源坑）。
判据四面：新夹具 `tests/dcsurf` 两片四头钉（DS01 同一枚反复读 + With 那一形三个数彼此相等且非零 / DS02 两枚互不相等 / DS03 `GetDeviceCaps(hDC, LOGPIXELSX) > 0` 问的是 GDI，证明交回来的是一张**活的 DC** / DS04 `GetPixel` 各自等于**自己那枚**的设计期底色；DS05 只钉前缀 —— dpi 是函数的数，不钉绝对值）；**负控 = 改前那台真编同一份夹具** ⇒ DcForm.c(167) error C2039 "hDC" 不是 "HWND__" 的成员、BUILD rc=1、一条读数都不出；形状针两条（夹具 `dcsurf_dc_shape` + 真工程 `dc_read_real`，Absent 里连**晚绑定那一形** `vb6_ComGetLongPtrProp(X, L"hDC")` 一起钉住）；哨兵 `scripts/check_control_dc.ps1` D1~D5，五条假 needle 全按"人会怎么改坏"植（绕过权威自己 `GetDC` / 权威不再认派发期那张 / 白拿那张不归还 / 销毁时不撤名 / 表里给成通用行）逐条能红，还原后全绿。
**护栏**：A/B 86 份（BASE = 改前那台 `cp`，43 工程 × 两架构）⇒ 被改的产物 8 份、**OFFENDERS 0**，逐行归因 = 14 对 K2（`X.hDC` / `ComGetProp(X, L"hDC")` → `vb6_GetControlHDC(X)`）+ 5 条 WD（`warning VB4001: P17.1: Unknown control property '.'hDC' in With block` 消失 —— 这条没了正是这一刀的目的，属性认识了不该再报不认识）。**顺带一条实测订正（差点按类推写错）**：Charts2020 的 `Proyecto1`/`ucProgressCircular` 里 `Picture2.hDC` 以前走的是晚绑定那条 `vb6_ComGetProp(X, L"hDC")` —— 我原本按 #143/#191 那族"原生控件没 IDispatch ⇒ 恒答 0"去推它，**探针量出来不是**：`.build/b196probe/`（一次性夹具，不进 tests/）用**改前那台**真编真跑裸形 `picA.hDC`，得 `PB01-HDC=671159075`、`PB02-DPI=96` —— 那是一张**活的**窗口 DC（`uc/detail/uc_hostmodel_getprop.inc:86` 里晚绑定那条本来就认 hDC）。所以本刀在裸形上改的不是"答 0"，是**语义**：VB6 是一个对象一张，那条兜底每读一次新取一张、也从不归还；现在两形同归一处出口、反复读回同一个句柄（DS01 钉的就是这个），并在窗口销毁时归还。**真红的那一半是 With 形**（`With Picture1 : TextOut .hDC` = C2039，编不过）。
**真工程配对读数**：ucTreeMaps x86 的 C2039 **2→1**（只剩 `TextHeight`；那条 `TextOutW` 的 C2198 本是 hDC 的级联，跟着一起消失）。
**欠的两条 + 本刀留下的一条口径**：按 HWND 的 `TextHeight`/`TextWidth`（要一枚"一个实参的控件方法"新机制 + 两形码头）、带 HWND 的 `ScaleX`/`ScaleY`（裸名改写仍挂在 `isDesignerModule_` 上，`.pag` 里的裸名认不出来）；另 —— **Form 自己那张缓存 DC 目前没有归还点**（归还只在 PictureBox/Image 那层的 `WM_DESTROY`）。语料里 `Form.hDC` 现存 **0 处**（本轮 86 份产物里没有一份因它而变），窗口销毁漏一张、进程结束由系统收回 ⇒ 不当成本轮的红，但记在这里，等 #199 那条属性页的路一起收。**顺序照旧：#199 → #196 剩下的两条**。

**→ 同账第二条（按 HWND 的 TextHeight/TextWidth）也已出（提交 `060eddca` + 针面订正 `ac3df329`，门 #324 = failure **红在我自己的两条新针**、门 #325 = run 37243854108、head ac3df329、attempt 1 = 11 job 全绿：10 片各自 FAIL=0，本轮七条相关用例在 CI 上逐行真 PASS —— `[VBP] dcsurf` / `dcsurf_x86`、`[CODEGEN-NOTE] dcsurf_dc_shape` / `dc_read_real` / `dcsurf_text_measure` / `text_measure_real`、`[STATIC] control_dc`）**：不需要新机制的"发明" —— 现有 `controlZeroArgMethod`（`cgen_util_ctrl.cpp:1452`）那套**表交名字、码头拼实参**的协议本来就支持带实参：With 形交裸名 + `pendingChainObj_`，由 `cgen_expr_call_com_bind.inc:842` 收下、`cgen_expr_call_opt_pad.inc:74-78` 把句柄**前置**到实参表前面 ⇒ 新表 `controlOneArgMethod` 只登记 `textheight`/`textwidth` × Form/PictureBox 两档（**依旧不给通用行**），两条码头各查一次（With 形 `cgen_expr_with.cpp`、带括号裸形 `cgen_expr_call_com_bind.inc`）；实参签名表补两行 `{"void*","BSTR"}`（Fix 113 那一味：漏了就等于把 vb6_VARIANT 裸喂给 `GetTextExtentPoint32W`）。RTL 那头新增 `vb6_ControlMeasureTextPx`（DC 走**同一处** `vb6_ControlDrawDC`、字体走 `WM_GETFONT` = 与 Print 同口径）+ 两个出口 `float vb6_ControlTextWidth/Height`（float = VB6 的 Single，别让生成 C 去做 double→float 收窄），**交回的单位过 `vb6_ScalePxToUser` + `vb6_WindowScaleModeSelf`**（#175/#197 那一份单位表，缇型对象上交像素就是 #177 那一味）。
判据：`tests/dcsurf` 加长 —— TH01 两形逐数相等且非零（With 与裸形不同归一处就是一头 0）、TH02 **比值**判据（同字体的缇框 vs 像素框 >4 倍；实测 16px vs 240 = 15 倍，不钉绝对数所以 DPI 变了不假红）、TH04 宽度随文字变（拦住"恒答一个常数"）；TH05/TH06 是只钉前缀的读数行。负控 = 改前那台真编同一份夹具 ⇒ rc=1、无 exe。哨兵 `check_control_dc.ps1` 加了 D6~D9（五条假 needle 逐条能红：单位不折算 / With 那条码头断 / 表少一档 / 出口返回档改回 double / 别处手拼发码 ⇒ 全红，还原回绿）。**顺手抓到哨兵自己一条假绿**：D9 第一版把 `vb6_ControlTextWidth` 与 `vb6_ControlTextHeight` 当**一段**正则来抓，改坏宽度那半时高度里的 `vb6_ScalePxToUser` 还在 ⇒ 整段照样过；改成"一个一个函数各自取身体"后那条 needle 才真的红（同 #197 的"注释行不算"、"计数要打印实测值"，是同一类自欺）。
护栏：A/B 86 份（BASE = 这两刀之前那台）⇒ 被改 8 份、**OFFENDERS 0**；逐行归因 = 32 对读法替换（hDC 与 TextHeight/TextWidth 都算 K2，配对正确性用**整行**判 —— 公共前缀会把 `Co` 这种片段折掉，残段里搜全名搜不到，这是本轮第二课）+ 8 条 WD（VB4001 那三条诊断消失）。真工程配对：ucTreeMaps 两台（x64/x86）的 C2039/C2440 **全部消失**，`VB4001 Unknown control property` 只剩 1 条（`ScaleX/ScaleY` 那一族另计），红点从"编不过"推进到 **LNK2019 ×3**（`AddFontMemResourceEx` / `GdipNewPrivateFontCollection` / `GdipPrivateAddMemoryFont` 三枚 DI 桩没登记，见新账 §B36/#201）。
**本刀没接的两处（记下不装绿）**：① 不带括号那条语句形（`picA.TextHeight "x"`，`cgen_call.cpp` 那一支）刻意没接 —— 语料 0 处，接它要先证明那条路上实参表怎么交；② 窗口字体那条判据今天**当不了判据**（见新账 §B35/#200）。

**第三条（带 HWND 的 `ScaleX`/`ScaleY`）今天的形状量清了（2026-10-05）**：`--emit-c` 读 PropPagFMR 那一页，`ScaleX(.CurrentX, .ScaleMode, vbPixels)` 发出来是**裸的 `ScaleX(...)`**（一个既没声明也没定义的 C 函数名）：

    vb6_di_TextOutW(vb6_GetControlHDC(_vb6_with_0), ScaleX(vb6_GetCurrentX(_vb6_with_0), vb6_WindowScaleModeSelf(_vb6_with_0), 3), ScaleY(...), vb6_StrPtr((*Text)), vb6_Len((*Text)));

两边实参已经按那一枚窗口算了（`.ScaleMode` → `vb6_WindowScaleModeSelf(_vb6_with_0)`），缺的只是**换算那一站**。注意接收者是 UC 时不缺：`ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)` 早由 Fix 111 发成 `vb6_UserControl_ScaleX` —— 那一站是**纯单位换算**（`vb6_ScaleUnitsPerPx`，账 #177 收成一张表），不吃窗口句柄；窗体型接收者要的是同一条换算但**vbUser(0) 那一档得问这枚窗口自己的 ScaleWidth/ScaleHeight**，所以缺的出口是`vb6_WindowScaleX/ScaleY(void* hwnd, double x, int32_t from, int32_t to)`，接线照 #196 前两条那**三处码头**（With 形 / 带括号裸形 / 语句路），一头不接就是 #143 那一族。

**一条今天的意外读数（别据此判"没接也没事"）**：这台工程今天 rc=0、0 error 出了 exe —— 但那份产物的 map 里**既没有 `ScaleX` 这个符号、也没有 `vb6_PropPagFMR_UnicodePrint`**（`PropPagFMR` 一共只剩 11 条别的符号）⇒ 那个调用点所在的函数被链接器的 **/OPT:REF 当未引用代码删掉了**，所以那条隐式声明根本没走到解析那一步。换句话说：**这一条今天不挡这台工程出 exe，是因为那页的入口本身还没接上（#199）**；一旦页被真正调用，它就是 LNK2019。⇒ 判据要自己钉（`--emit-c` 断"产物里不许出现裸 `ScaleX(`/`ScaleY(`" + 真跑一头换算读数），**不能拿"这台工程现在编得过"当这条通了**。

**→ 第三条（`ScaleX`/`ScaleY` 的换算那一站）也已出（2026-10-05，提交 `479b202f`，门 #329 = run 37260903318、head 511d356e、attempt 1 = **11 job 全 completed/success**（含 Build C3.exe 那一片；这个 overall 读数今天独立落到两次 —— watcher 的 poll 8 与事后复核各一次）。**用例级的 PASS 行仍然没拿到**：这轮 jobs 数组里没给 `log_url`，改走 `/actions/jobs/{id}/logs` 拿到 10/11 份，但取回的正文里一条 `[VBP]/[CODEGEN-NOTE]/[STATIC]` 都没有（本机那台中转对日志端点给的不是真日志正文，重试到后面连 jobs 那一条也顶成非 JSON）⇒ 这一行不写任何"CI 逐行真 PASS"的句子）**：先记一句**预告被普查推翻** —— 上一段写的「缺的出口是 `vb6_WindowScaleX/ScaleY(void* hwnd, ...)`」（理由：vbUser(0) 那一档要问这枚窗口自己的 ScaleWidth/ScaleHeight）**在语料里没有承载点**。全仓 `ScaleX(`/`ScaleY(` 的调用点普查：Charts 2020 五份 .ctl 各 4 处 + `PropPagFMR.pag` 2 处 + `tests/VBFlexGridDemo`（`MainForm.frm` 2 + `Common.bas` 2 + `Builds/VBFlexGrid.ctl` 一批）+ `archive/ctxWinsock.ctl` 2 处，实参只有三种形状 —— 字面 `vbPixels(3)`、`vbContainerSize`/`vbContainerPosition`、`.ScaleMode`/`Me.ScaleMode`/`UserControl.ScaleMode`（#197 下发之后答 1 或 3），**0 处传 0=vbUser** ⇒ 这一刀发的是**不带 HWND** 的那一条：`double vb6_ScaleUnitX/Y(double x, int32_t fromScale, int32_t toScale)`。

**没接的那一头写明**：传 0=vbUser 时 VB6 要的是这枚窗口自己的用户坐标系，而 `vb6_ScaleUnitsPerPx` 的既有口径（#177 里就写明 User/ContainerPosition/ContainerSize/unknown 都按像素）会把它当成像素 ⇒ 那种调用现在静默给恒等值。语料 0 处，记在这里不装绿；要接就是「换算拿 HWND」那一条，三处码头得把句柄交进 `vb6_ScaleUnitX/Y`。

**做法 = 一处权威 + 三处码头 + 把 Fix 111 那两条改成薄壳**：
- RTL（`vb6rtl_com.c`）：`vb6_ScaleUnitX/Y` 就是原来 `vb6_UserControl_ScaleX/Y` 的身体，UC 那两条改成**直接 return 这一条** ⇒ 「UC 那一档」与「窗体型那一档」从此不可能分家（不是在 #177 那张表旁边再抄一遍）。
- 后端：新增 `controlScaleMethod(ctrlType, memberLower)`（`cgen_util_ctrl.cpp:1518`，紧挨 `controlZeroArgMethod`/`controlOneArgMethod`，**表只交名字**），登记 Form/PictureBox 两档，**不给通用行**（同 #192/#196 那条口径：`List1.ScaleX` 答一个数就是伪造成功）。三处码头各查一次 —— With 形 `cgen_expr_with.cpp`、带括号裸形 `cgen_expr_call_com_bind.inc:263-290`（拦在 C29-4 那块 StatusBar 之前）、**裸名** `cgen_expr_ident_builtin.inc:242-246`。裸名那一处就是本节开头说的那格：以前改写挂在 `isDesignerModule_` 上（只有 .ctl/.pag 那两种模块种进得去，且发的是 `vb6_<Host>_<Member>` 那种宿主伪成员形状），这一刀把它放宽到 `isFormModule_ || isPropertyPageDesigner_`；`.ctl` 里裸写那一形**没动**，仍由 `kHostPseudoRows` 那两行 UC 条目负责（`cgen_util_com.cpp:495-496`，HPF_BARE）⇒ 现在也落到同一个换算上，只是多一层薄壳。
- With 形**刻意不置 `pendingChainObj_`**：换算不吃句柄，置了调用点就会把 HWND 前置成第一个实参 ⇒ 实参表错位（与 #196 第二条那一条协议正好相反，差别就在「要不要对象」）。

**判据**：夹具 `tests/dcsurf` 加长 SX10/SX11 —— **四形逐数相等**（`picB.ScaleX(1440, 1, 3)` / 裸 `ScaleX(...)` / `Me.ScaleX(...)` / `With picB : .ScaleX(...)`）再与 `ScaleY` 那一形、一枚**问窗口**的证人 `GetDeviceCaps(LOGPIXELSX)` 对上：一头钉「四形同归一处」，另一头钉「数真的是按 DPI 换算的而不是烘出来的常数」。`SX10-FOURFORMS=True` 进 `$dcSurfExpected`（两台各一条），SX11 留原始读数。**两台真跑逐行相同** = `SX11-RAW pic=96 bare=96 me=96 with=96 y=96 dpi=96`（1440 缇 @96dpi = 96 像素；`.build/b196out/n64.out` / `n32.out`）。负控：#196 那台 BASE 编今天这份夹具 ⇒ exit 2、C2039×3 + C2440×7、诊断面三条 `VB4001 P17.1: Unknown control property`（hDC / TextHeight / **ScaleX**）三条全在、一条读数都不出（`.build/b196out/b64.log`）。

**形状针两条**进回归：`dcsurf_scale_units`（present = 四形各自那一行 `sx1 = vb6_ScaleUnitX(1440, 1, 3)` 等；absent 钉住改前的两形 —— 拿 HWND 当 IDispatch 问的 `vb6_ComCallDouble(vb6_hwnd_picB, L"ScaleX"` 与 With 那一形 `.ScaleX(1440, 1, 3)`）、`scale_units_real`（钉在**真工程** ucTreeMaps：present `vb6_ScaleUnitX(vb6_GetCurrentX(_vb6_with_0)`，absent `", ScaleX(vb6_GetCurrentX(_vb6_with_0)"` 与 `L"ScaleX"`）。两条 absent 面钉的就是改前产物里的样子 ⇒ 在改前那台上必红 —— **这一句是按产物形状推的，没再单独拿 BASE 那台跑一遍针面**（那一台跑的是整份夹具的真编，见上面那条负控）。

**哨兵** `check_control_dc.ps1` 新增 D12：`controlScaleMethod` 定义 1、三处码头各≥1、字面 `"vb6_ScaleUnitX` 恰好 2 行且分属 {`cgen_util_ctrl.cpp`, `cgen_expr_ident_builtin.inc`}、身体认两名且指向两个出口、不给通用行。假 needle 真红过：把 With 那条码头注掉 ⇒ `FAIL D12 少了一条码头: cgen_expr_with.cpp`，还原回绿（`.build/b206_sent*.txt`）。**哨兵自己第一枪又是规则写太紧**：硬编码字面量那条我先写「恰好 1 处」，而那张表自己那一行也是字面量 ⇒ 真数是 2；改成「恰好 2 行且分属这两个文件」才是设计本意（与 #196 第二条那条「计数要打印实测值」同一类自欺 —— 红的是我写的判据，不是产品）。

**护栏 A/B** 86 份（BASE = #202/#204 那台的快照 `b204_new_emit`）⇒ changed=6、**OFFENDERS=0**，逐份归因：ucTreeMaps 两台各 1 块（PropPagFMR 那行 `ScaleX(...)` → `vb6_ScaleUnitX(...)`）；**VBFlexGridDemo 两台各 1 块 2 行** —— `MainForm.frm` 那两行布局以前发成 `vb6_VariantToDouble(vb6_VariantFromComResult(vb6_ComCall...(L"ScaleX"...)))`（把 HWND 当 IDispatch 问属性 ⇒ 交回 Empty、再按数值打），现在答得出数；这两份一开始被我当 offender 抓到，查实是正当改道后补进 WD（与 #197 那轮同工程那条 `Me.ScaleMode` 同族）；dcsurf 两台各 4 块，全是夹具新增行、无一行删除。**结构性断言**：86 份新产物里以调用形式出现的裸 `ScaleX(`/`ScaleY(` **0 处**（`grep -rlE "(^|[^.a-zA-Z_0-9])Scale[XY]\(" b206_new_emit/` 返回空）。

**真工程**：ucTreeMaps 两台 rc=0，exe x64 660,480 → **660,992**、x86 562,688 不变（段对齐，见 §B35 那条「尺寸没变当不了没重编的判据」）；诊断面 109 VB3001 + 6 VB3003 + 1 VB4001，而**仅剩那一条 VB4001 现在是 `TypeLib reference path not found`** ⇒ `Unknown control property` 这一族在真工程里**归零**（`.build/b196out/tm64.log` / `tm32.log` 各 grep 0）。**本账三格（hDC / TextHeight·TextWidth / ScaleX·ScaleY）到此全部出完**；开工顺序里 #196 那一格收口，剩下的还是 #199 的甲/乙口径（页里的控件从没被创建 ⇒ 那一页的绘图面至今白画）与 §B40/#205。


### B32 `VB6_ScaleMode` 在整个 RTL 里没有任何人写它（账 #197）

`grep -rn "VB6_ScaleMode" src/rtl src/backend` 的全部命中只有 setter 自己（`vb6forms_widget_prop.c:265` 的 SetPropW）与两条读点，**0 个调用者**。⇒ 谁去读控件或窗体的 `ScaleMode` 都只会拿到缺省 1，跟 .frm/.pag 里写的设计值无关；所有按单位换算的路径（ScaleX/ScaleY、缇/像素互转）因此都建立在一个恒为 1 的数上。这与 #154/#160 同族（设计期属性从没下发到窗口），但影响面更大：它决定的是**量出来的数对不对**，不是某一枚控件的外观。
开工第一步：先量"设计期 ScaleMode 有没有进过 me->properties 表"（`.frm` 里 `ScaleMode = 1` 那行是 Form 级的，控件级只有 PictureBox 一类容器有），再决定下发点落在创建那两路的哪一处（同 #83/#151 那条"两条创建路都要打"）。

**→ 已出（提交 `9565bdbb`，门 #322（run 37235573381，head ada93533，attempt 1）= 11 job 全绿、11 片日志**全部下到**（这轮代理没拦）：10 片各自 FAIL=0（第 11 片是 Build C3.exe，不打这个计数），本轮五条新用例在 CI 上逐行真 PASS —— `[VBP] scalemode` / `scalemode_x86` / `[CODEGEN-NOTE] scalemode_design_write` / [CODEGEN-NOTE] scalemode_read_real` / `[STATIC] scalemode_writers`）**。读数与做法：
- **RTL 一个字节没动** —— `vb6_GetScaleMode` / `vb6_SetScaleMode` 早就在（`vb6forms_widget_prop.c:256/263`，声明 `vb6forms_prop_form.h:77/78`，由 `vb6forms.h` 传递可见），缺省 1=缇。这格从头到尾是**发码侧没人调用**，不是运行时缺出口 ⇒ 不 touch `c3rtl.rc`，也就绕开账 #156 那条「RTL 嵌在 C3.exe 资源里，探针测的是旧 RTL」的坑。
- **普查**（`.build/b197_census.py`，93 份 .frm/.pag/.ctl 设计块）：19 处 ScaleMode —— UserControl 10 处（9×3 + 1×1，**已由 .ctl 注册那条路接走**，`cgen_form.cpp:393-436` 把它交给 `vb6_UC_WindowScaleMode`）、PictureBox 6 处（全 3）、Form 3 处（1×1 + 1×3）。⇒ 缺的就是 PictureBox 与 Form 这两档，而 VB6 的设计器**只在这三类块里写这一行**，所以口径收成「设计块写了就发」：不按值筛（声明 1 与没声明在产物里要分得开），也不再另开一张控件型白名单（Frame 那一类根本不写这行，写了就是给人读的）。
- **一处权威 × 三条落点**：新增 `CCodeGen::emitDesignerScaleModeProp`（`cgen_util_ctrl.cpp`，紧挨 #154/#160 那两个同型出口），三条路各调一次 —— 顶层 `cgen_form_ctrl_style_apply.inc`、容器子控件 `cgen_form_frame_menu.inc`（账 #83/#151 那条「两条创建路都要打」）、窗体自己那档 `cgen_form_wndproc_create.inc` 的 P20-40 块。
- **读写成对**：两张表各补 `scalemode` 两行（PictureBox 与 Form）—— 读给 **`vb6_WindowScaleModeSelf`**，也就是 #175 那张单位表的同一处，程序读到的数与几何换算用的数**不可能分家**；写给 `vb6_SetScaleMode`。以前 `Picture1.ScaleMode` 走的是 COM 兜底（把 HWND 当 IDispatch 问属性 ⇒ 交回 Empty）。
- **判据四面**（新夹具 `tests/scalemode`，x64+x86 两片）：SM01 三处设计值各自读回（`SM01-MODE=3/3/1`）；SM02/SM03 **问窗口** —— 同尺寸的缇框与像素框，ScaleWidth 与框内按钮 Left 必须差一个单位比（本地实测 1170 vs 78、120 vs 8 = 15 倍 @96dpi；针面只取 `>3` 这个**比值**，不钉绝对数，DPI 变了不假红）；SM04 运行期切档后两枚**逐数相等**（证明这数是活的，不是创建时烘死的）；SM05 另一枚没被写坏（RTL 按 HWND 存，全局一份就露馅，同 #192 双向钉）；SM06/SM07 **第二条创建路**（Frame 里那枚 PictureBox）也发了这一句。**负控 = 改前那台真编同一份夹具**：`SmForm.c(166)/(182)` 两条 C2039 "ScaleMode" 不是 "HWND__" 的成员 + C2198，rc=1、一条读数都不出（物证 `.build/b197out/b64/c3-error.log`）。
- **形状针两条**：`scalemode_design_write`（夹具：Present 三句 SetScaleMode 与 `vb6_WindowScaleModeSelf(_vb6_with_`，Absent 三形 HWND 取成员）+ `scalemode_read_real`（钉在**真工程** ucTreeMaps 那一行上）。两条都在改前那台上当场红（缺 4/缺 1 + 漏 1），改后绿 —— 见 `.build/b197_needle.txt`。
- **哨兵** `scripts/check_scalemode_writers.ps1`（W1 手拼发码点=1 / W2 权威一次定义且仍读设计块那条属性 / **W3 三条落点各≥1 且注释行不算** / W4 读写成对≥2 each / W5 RTL 里写 `VB6_ScaleMode` 这个属性名=1、读≥1）；假 needle 验红：**把容器那条落点注释掉** ⇒ W3 点名、exit 1。**这轮自己踩的工具坑两条**：① 第一版 W3 用整篇正则，把调用点注释掉**照样绿** —— 植针要按「人会怎么改坏」植（注释掉/删掉），我先前植的是改个 `XX` 前缀，`Contains` 当子串照样命中，等于没验；② PASS 消息里写死 `W1=1 W3=3` 会掩盖真数，改成打印算出来的计数。
- **护栏 A/B** = 82 份（41 工程 × 两架构，BASE = 改前那台，`.build/b197_base/C3.exe` 用 cp 不重建）⇒ **12 份被改，OFFENDERS 0**：10 份各多 1 行 K1（ctrltabindex / czUI / piccur / tabwalk / ve_units，就是普查里那 5 个声明了 ScaleMode 的工程）；2 份各改 2 行 K2（VBFlexGridDemo 两架构）—— 那是 `Me.ScaleMode` 以前发成 `vb6_ComGetProp(vb6_hwnd_MainForm, L"ScaleMode")`（**把 HWND 当 IDispatch 问属性 ⇒ 交回 Empty**），现在答得出数。逐行归因见 `.build/b197_ab.txt`。
- **真工程读数**：ucTreeMaps x86 的 C2039 **4→2**（剩 hDC / TextHeight = #196）。**czUI-main**（语料里唯一在 **Form** 级声明 3=Pixel 的工程，几何读数真的会被这一档挪动）真编真起窗：`CZUI_BUILD=0` / `CZUI_ALIVE`（`.build/b197_gui.bat`）。
- **留给下一轮**（这一行写于 #197 收线时，2026-10-05 当天 #196 的 hDC 那一半已随门 #323 出掉 ⇒ 剩下改成）：#196 欠按 HWND 的 TextHeight/TextWidth（要一枚"一个实参的控件方法"机制：现有那张 `controlZeroArgMethod`（`cgen_util_ctrl.cpp:1452`）有三条码头 —— With 形 `cgen_expr_with.cpp:72`、语句形 `cgen_call.cpp:400`、带括号裸形 `cgen_expr_call_com_bind.inc:80`，新机制同样要三条都接）+ 裸 `ScaleX/ScaleY`（那条改写仍挂在 `isDesignerModule_` 上，`.pag` 里的裸名认不出来）；以及 §B34 那条**新账 #199**（甲/乙口径待拍）。

### B34 `.pag` 的设计块控件从来没被创建（账 #199）

做 §B32 的时候顺带量出来的，比 §B31 剩下的两条更底层：同一份 `--emit-c` 里 **PropPagFMR 那一段 `vb6_CreateControl` 0 处，同工程 Form1 那一段 14 处**。页里那枚 `Picture1` 的句柄是
`#define vb6_hwnd_Picture1 (*vb6_UC_DesignSlotOf(me, "Picture1"))`（`cgen_form.cpp:207-212` 那条 UserControl 内嵌控件的路数），而 `vb6_uc_designSlotIn`（`uc_host.c:641-652`）是**按需新建、初值 NULL** —— 没有任何人往里写。⇒ 属性页里 `With Picture1 : .hDC / .CurrentX / .ScaleMode / .TextHeight` 全打在 **NULL 句柄**上：编得过（#197 之后连 C2039 都不报了）、发码对、运行期一句不响。
**这条不修，#196 做完 ucTreeMaps 的 UnicodePrint 也还是白画** —— 所以控件线的顺序改成 #199 → #196。
**已量到（2026-10-05，读码，未动产品）**：那三条"先量"的问题一次答完 ——
① `.pag` **解析没问题**：`driver_frontend.cpp:172` 对 `.pag` 与 `.frm` 走同一个 `FrmParser::parse`，`frm_parser.cpp:485-487` 认根块 `Begin VB.PropertyPage`、`:371-372` 递归收子控件 ⇒ 页里的 `Picture1` **在 `frmDesc.formControl.children` 里躺着**（那条 `#define` 本身的存在就是证据）。
② 分岔在**发码的模块种类那一问**：`cgen_base_generate_decl_pass.inc:39-46` —— `if (module.isFormModule && frmDesc) emitFormFramework(...); else if (frmDesc) emitDesignerControlDecls(...);`。而 `isFormModule` 只对 `.frm` 置真（`driver_frontend.cpp:161-163`，`.pag` 只被折进 `isClassModule`），`isDesignerModule_ = frmDesc && !isFormModule`（`cgen_base_generate_prologue.inc:15`）。`vb6_CreateControl` **只在 `emitFormFramework` 的片段里出现**（`cgen_form_create_controls.inc` → `cgen_form_ctrl_style_apply.inc:339`，容器子控件在 `cgen_form_frame_menu.inc:132`）⇒ 页压根没进那条路，`emitDesignerControlDecls`（`cgen_form.cpp:170-213`，注释自己写着"不发射窗体窗口框架 … 子控件句柄保持 NULL"）只发槽位宏与名字登记。
③ `.vbp` 侧另有第三档：`VbpSourceType::PropertyPage`（`vbp_parser.cpp:67`）。
⇒ 定性：**这是 (b) 一条模块种类的发码分支，落成 (c) 一个没接完的架构决定** —— 那句"槽位由宿主窗体写"对 `.ctl`（被窗体的 CreateControls 填）成立，对**独立一页**不成立，而属性页的宿主/装页 runtime 本仓没有。两条候选路摆在这里：**甲** = 页也走 `emitFormFramework`（跟 .frm 一样发创建 + 句柄赋值，代价是页要有一枚真窗口/承载体）；**乙** = 保留槽位语义，补一个"页装好时把创建出的句柄写进槽"的宿主（改动小、语义正，但要先回答"谁创建页的窗口"）。语料侧的事实：`PropPagFMR` 在整个 VbEclipse 里**只被 .vbp 列了名，没有任何一处实例化** ⇒ 这一格不做完，页里的绘图面永远是白画，但它也不挡别的工程。开工前先拍甲/乙这条口径。判据照 #83/#151/#160 那条纪律：**两头钉**（新格有没有真的发创建 + 旧格有没有被抢），别只看编译过没过 —— 这一格"编得过"今天就已经是绿的样子了。

### B35 控件窗口字体换了，文字量跟着不动（账 #200，**已出**）

做 #196 第二条时撞出来的，**归因还没做**，所以那条判据今天当不了判据（夹具里只留读数 TH06）。读数：`tests/dcsurf` 里 `picB.FontSize = 20`（发码确实是 `vb6_SetControlFontSize(vb6_hwnd_picB, 20)`）之后再量 `picB.TextHeight("Xg")`，**量回来还是 16**；设计期写 18pt 的 picA 也量 16，而 `picA.FontSize` 读回 18（自存的那份读得回来，问窗口的没有）。两头都哑：`picA.FontPixelHeight`（RTL 里 `vb6_ControlFontPixelHeight` = WM_GETFONT + GetObject，#153 那套"问窗口的证人"）今天**整条 Debug.Print 一行都没打出来**——发码是齐的（`vb6_ControlFontPixelHeight(vb6_hwnd_picA)`），运行期那句却像被吃掉，这本身又是一条要查的。
三种可能分不开，按顺序问：① `vb6_SetControlFontSize` 的 WM_SETFONT 对 PictureBox（自绘 + 分层子类那层，见 #185）到底发没发出去；② 发了的话，设计期那条 pass 有没有被后面的创建/装子类 pass **再设回默认**（同 #185"同名槽位被后装那层挤掉"那一族的另一味）；③ `FontPixelHeight` 那句为什么不打（夹具的语句被谁吞了）。
**为什么值得单开一账**：#177/#129 两轮的结论都是"文字量必须按控件自己的字体与单位"，单位这一半这轮接通并钉住了（TH02 实测 15 倍），字体这一半现在只能靠代码口径（`WM_GETFONT`，与 Print 同源）自证，**运行期没人验证过**。真工程里 Charts 2020 的图例/标题全走这条路，字号一歪就是整张图的排版歪。
**→ 已出（2026-10-05）**：归因 = ①②③ 三条猜的**都不是**，真身在窗口类本身。一次性探针（`.build/b200probe/fontprobe.c`，裸 STATIC、谁也没子类化过）实测 `WM_SETFONT` 之后 `WM_GETFONT` 回 NULL，`STM_SETFONT`/`STM_GETFONT` 那一对同样回 NULL —— STATIC 这个类**压根不记字体**。于是原来那三条读法（`.FontName/.FontSize` 背后的 `vb6_GetControlLogFont`、Print 落笔前的选字体、#196 的 `vb6_ControlMeasureTextPx`）在这类窗口上永远拿不到用户设的那张，只能拿 DC 的默认字体画；而 `.FontSize` 读回来还是设计值（那是另一份自存的属性）⇒ 两头都不报错，只有量出来的数不动。③ 那条 `FontPixelHeight` 整行不打也归到同一条因：它自己就是 `WM_GETFONT + GetObject`，拿到 NULL 就直接 `return 0`。

**改**：一处出口 + 一份自存。新增 `static HFONT vb6_ControlFont(HWND)`（先问窗口，回 NULL 再读窗口属性 `VB6_CtrlFont`），setter `vb6_SetControlFontFromLogFont` 把自己 `CreateFontIndirectW` 出来的那张存进这个**新名字**（#185「一层一个属性名」那条纪律），三条读法全改走这一处。旧字体的找法刻意改成「先读自存、找不到才问窗口」—— 顺序反了会对真记字体的那几类（EDIT/BUTTON）**双删**（`WM_GETFONT` 回来的正是我们上一轮存进去的那张，两边都当成旧字体）。**顺带修掉一处 GDI 泄漏**：改前 STATIC 每写一次字体就漏一张，因为那句 `DeleteObject` 依赖的 `WM_GETFONT` 恒回 NULL ⇒ 旧字体永远找不到；现在找得到、也删得掉。已知边界（记在这里不装绿）：窗口销毁时最后一张字体不被 `DeleteObject`（Windows 回收属性表、但不认识 GDI 对象），一枚控件至多一张 —— 普通控件没有统一的 WM_DESTROY 挂钩（只有 picture/image 那层有，就是 #185/#196 归还 DC 的那一站），要补这一头得先造机制。

**运行期读数**（真跑，x64 与 x86 **逐行相同**）：`TH06-FONTRAW a=29 b=16 b2=27 fsA=18 pfA=24 pfB=27` —— 设计期 18pt 的 picA 从 16 → **29**；运行期改 20pt 的 picB 从「改前改后都是 16」→ **16 / 27**；证人 `FontPixelHeight` 两台都答 **24 / 27**（改前那一句整行都不打）。TH03 因此从「只留读数」升回**真判据**：`ok7 = (tA > tB) And (tB2 > tB) And (pfA >= 18)`，`TH03-FONT=True` 进 `$dcSurfExpected`（两台各一条）。真工程配对：ucTreeMaps 两台仍 rc=0、诊断面 0 error、warning 面 109 VB3001 + 6 VB3003 + 1 VB4001（与 #201 那轮同一条），exe 659,968 / 562,688 **与 #201 那轮两个数字一模一样** —— 同尺寸一开始被我读成「这台没真的重新链接、拿的是旧产物」的形状，所以把两台产物目录**删空重跑**（09:31 / 09:32 新写的文件）：照样同尺寸 ⇒ 那是 PE 段对齐把这点增长吸收了，不是旧产物。**这一枪打在自己身上是有价值的**：判据「产物逐字节相同」本来就当不了判据（记忆里那条冷编两次 md5 就不同），但**反过来**"尺寸没变"也不能当成"没重编"的判据 —— 要问就删空目录再问一次。

**护栏**：哨兵 D10 四条 —— 出口在 `vb6forms_ctrl.c` 恰好定义一次 / 文件里带 `WM_GETFONT` 的那一行恰好 1 且必须落在出口体内（第 dl+1..dl+8 行）/ `VB6_CtrlFont` 写者恰好 1 / 读者 >= 1。**假 needle 真红过**：把 measure 那一处换回裸 `SendMessageW(hw, WM_GETFONT, 0, 0)` ⇒ `FAIL D10 问窗口字体的那一行 = 2 处 -> vb6forms_ctrl.c:325 | :733`，换回来即绿。**这一刀自己的第一枪红得不该怪产品**：D10 落地那一次报的是 3 处，多出来那两条是我给两行**行尾注释**写了 `WM_GETFONT` —— 哨兵跳行首 `//`、不跳行尾注释，于是「注释把形状写出来」就自己造了红；措辞改成不嵌那个 token 才对上（同「计数要打印实测值」那一类自欺，只是这回是自欺的方式换了个方向）。**A/B** 86 份（BASE = #201 那份 emit 快照、NEW = 现在这台）⇒ changed=2、**OFFENDERS=0**；那 2 份就是 dcsurf 两台，逐行归因全是夹具自己新增的那几行（pfA/pfB 两条声明与两次 `vb6_ControlFontPixelHeight` 读、ok7、TH03 那条 Debug.Print、TH06 那串 concat 变长）—— 后端一字节未动，这正是「只动 RTL 的一刀在发码面应当一字不变」这条立论的钉法。

**剩下的同族**：探针这一轮又补了一枚裸 `BUTTON`(BS_GROUPBOX)，**同样回 NULL** ⇒ `vb6forms.c` 里 Frame 标题带那处读法一并定了罪，另立新账 §B37/#202（写法已经现成，缺的是把出口跨文件递过去）。

### B36 ucTreeMaps 现在只差三枚 DI 桩（账 #201，**已出**）

#196 两条接通之后，`tests/Charts 2020/ucTreeMaps/Proyecto1.vbp` 两台的诊断面从 C2039/C2440 **全清**，红点推到链接期：`LNK2019` 三条 + `LNK1120` —— 缺的符号是 `_vb6_di_AddFontMemResourceEx@16`、`_vb6_di_GdipNewPrivateFontCollection@4`、`_vb6_di_GdipPrivateAddMemoryFont@12`，调用者是工程自己的 `vb6_FontMemRes_AddMemFonts`（把字体从内存资源装进私有 FontCollection，GDI+）。
口径照记忆里那条：**手写一枚 `vb6_di_*` 单桩，别重跑 `gen_di_stubs`**（它按当次 session 整文件重写，会撤掉别人的桩），并同步 `scripts/di_stubs_manifest.txt` 与 `check_di_stubs.ps1` 的普查，不然哨兵会说谎。这三枚桩是**真实现**（AddFontMemResourceEx 要真的把字节交给 GDI、GdipPrivateAddMemoryFont 要真的建私有集合）而不是空桩，否则 FontMemRes 这条链只是从"编不过"变成"跑起来没字"——那是 #196 一开始就要躲开的"静默空转"。做完这一格，ucTreeMaps 就是第四枚能进 `Test-VbpBuild` 清单的 UC 子工程（#187 那三条之外的第一条新增）。
**→ 已出（2026-10-05）**：三枚桩手写进 `src/rtl/core/di/vb6_di_stubs.c`，清单基线 620→624（`check_di_stubs.ps1 -Update`），ucTreeMaps 两台第一次真编真链出 exe（x64 659,968 / x86 562,688，诊断面 0 error），并已升格进 `Test-VbpBuild` 清单。细节与那条"pdv 为什么交 NULL"的取舍在 §D 的账 #201 那一行。

### B37 Frame 的标题带是按默认字体量的（账 #202，**已出**）

做 #200 时把那颗一次性探针补了一枚裸 `BUTTON` / `BS_GROUPBOX`：`WM_SETFONT` 之后 `WM_GETFONT` **同样回 NULL**（读数 `C plain BUTTON : WM_GETFONT=0000000000000000 want=... match=0`）。于是 `src/rtl/core/vb6forms/vb6forms.c` 里 Frame 标题带那一处（接管绘制那趟的 `HFONT hf = (HFONT)SendMessageW(hwnd, WM_GETFONT, 0, 0);`）拿到的是 NULL ⇒ `SelectObject` 整步跳过，带高 `th` 与带宽 `sz.cx` 都按 DC 的默认字体算，而组框自己画标题用的是我们 `WM_SETFONT` 给过去的那张 —— 帧字体一改，那条白带要么盖不住标题、要么盖过头。今天不红是因为语料里没人改 Frame 的字体（设计期默认 8.25pt 与系统默认算出同一个数），这正是「编得过、跑了、什么都没发生」那一族的另一味。

**修法在 #200 里已经造好了**，缺的是把它跨文件递过去：那一行改读 `vb6_ControlFont` ⇒ 要先把出口从 `static` 提出来、进内部头；同时 `VB6_CtrlFont` 的写侧要覆盖到 Frame 这一档（现在唯一那条写者是 `vb6_SetControlFontFromLogFont`）。哨兵 D10 今天刻意只圈 `vb6forms_ctrl.c`（规则注释里写着原因），跨出来那一天应当把普查范围一起放开成整个 `src/rtl` —— 别留一份「两处口径、只守一处」的表。

**→ 已出（2026-10-05）**：修法就是本账预设的那一条 —— 出口从 `static` 提出来、进 `vb6forms_internal.h`，那一行改读 `vb6_ControlFont`。同一条根因今天**一并接上了六处**（这一处出口的覆盖面本来就不该按账号算）：`vb6forms.c` 的 groupbox 标题带读 + 控件创建路存、`vb6forms_ctrlarr.c` 问模板要字体 + 发给新窗口之后存、`vb6forms_shape.c` 三处图钮标题的文字量、`vb6forms_widget.c` 的 Label AutoSize 宽度。后两族今天没写判据 —— shape/widget 那两处改的是"画/量的时候用哪张字体"，观感类后果（截字/留白）没有一条现成的数值判据能钉，**别把"编得过、跑了"当成它们被验过**；钉住的是哨兵 D11 那条普查（见 §B39 收线段）。


### B38 ucTreeMaps 的 exe 第一次真跑：x86 起窗、x64 启动期 AV（账 #203，**已归因，按口径就地判掉**）

#201 让这台工程第一次真编真链出 exe，本条是那份产物的**第一趟真跑**读数（`.build/b200_tmrun.ps1` 收 stdout/stderr 与退出码，`.build/b182_winprobe.ps1` 数窗口）：

- **x86** = `alive=True mainhwnd=0x1606BC`，顶层窗口 6 条，其中 `cls=VB6_Form_Form1 txt=Form1` 是可见的 ⇒ **窗体真起来了**（另五条是 GDI+ Hook Window、ComboLBox、MSCTFIME 与两枚 Default IME，都是系统的）。
- **x64** = 8 秒内自己退出，`code=-1073740771`（0xC0000409 fail-fast），而 crash trace 打的是 `code=0xc0000005 at rva=0xce8a7463` + `av read target=0x48777bc3`（那是个野值），24 帧里只有 5 帧落在 exe 内（rva=0x2602c / 0x35317 / 0x2194e / 0x21a4c / 0x21f74），stdout 0 字节、Form1 那一枚窗口根本没出现。

⇒ 这一格从「编得过」推进到 VB6 那一侧的「能不能跑」；x64 起不来是**新缺陷**，不是 #201 那三枚桩的余波（那三枚在 x86 那条路上同样被调到，窗体照样起来）。**下一步（还没做）**：把 x64 那条 AV 归因 —— 顺序照 #182 那一味先拿带符号的产物把那几个 rva 落成函数名（**没符号化的 rva 名单不构成结论**），再分岔问「是产品发码把指针按 32 位存了」还是「工程自己的 Declare 把指针写成 `As Long`」，后者是改 VB 源码、不动编译器（#163/#175 那一族早已立过口径）。注意 x86 这一侧今天是**好的**，所以任何「回归坏了」的判据都不该把它算进去；反过来说，门禁今天只对这台工程断言「编得过」（Test-VbpBuild），跑得起来这件事还没进任何判据。

**→ 已归因（2026-10-05，靠符号化、不靠猜）**：把这台工程用 `-g --keep-for-debug` 重编一遍（产物带 PDB+MAP，`.temp/b204_tm_g.ps1`），再跑一次拿**新的** RVA，喂给 `.temp/b204_sym.ps1`（照 `.temp/sym3.ps1` 的写法，只是把 exe 路径做成参数 —— sym3 那份把 VBFlexGridDemo 的路径写死了）⇒ 逐帧落成**生成的 .c 的 file:line**。

读法上有一条会反复咬人：`#0` 落在 `rtl/vb6rtl.c:93` —— 那一行是崩溃轨迹**打印器自己**（`CaptureStackBackTrace`），而 `#1..#4` 全在 exe 之外（ntdll 的派发链）⇒ **真正的出错指令是 `#5`**：`rtl/vb6_di_win32_stubs.c:411`，调用链 `#6 FontMemRes.c:288` → `#7 :302` → `#8 :148` → `#9 :82`。而 `FontMemRes.c:288` 那一行是 `vb6_di_RtlMoveMemory((void*)&(lAddress), lpArray, 4);`（在 `IsArrayDim` 里）。

往 VB 源码看就见底了（`tests/Charts 2020/ucTreeMaps/FontMemRes/FontMemRes.ctl`）：`34: Declare Function VarPtrArray Alias "VarPtr" (Ptr() As Any) As Long` 与 `186: Function IsArrayDim(ByVal lpArray As Long)` —— **指针在 x64 上被 `As Long` 截成 32 位**，所以那个"读目标地址" `0xffffffff927ee7d0` 根本不是指针，是被截过又符号扩展的残值。⇒ 定性 = **工程自己的声明没匹配 64 位**，与 #163/#175/B27 同一条口径：**改 VB 源码（`As LongPtr`），不动编译器**；x86 那台 LongPtr 就是 4 字节、今天实测起窗正常。

**下一刀怎么做**（还没动手）：① 这两处 `As Long` → `As LongPtr`，并把同工程里其余"把指针当数交出去"的声明一起扫（census：`Declare ...` 里参数/返回是指针形状的那些逐条判定，不是猜）；② 判据要落在**运行面** —— 门禁今天对这台工程只断言「编得过」（`Test-VbpBuild`），"跑起来起不起窗"没人钉；现成的形态是 `Test-GuiVbp`（#175/#177 那批用过它），起窗判据 = 进程还活着 + 顶层出现 `VB6_Form_Form1`（本地读数已证 x86 这样认得出）。③ 两台都要真跑：x86 是回归护栏（今天好的不许坏），x64 才是这一刀的判据。

**census（今天数的，不是估的）**：那两处只是**第一个被撞到的**，同这一枚 `.ctl` 里指针形状的 `As Long` 还有 —— `34 VarPtrArray(...) As Long`（返回的就是指针）、`37 TlsSetValue` 的 `lpTlsValue`、`41 GdiplusStartup` 的 `token`（ULONG_PTR）、`43/44/45/46` 那四枚 GDI+ 的 `mFontCollection`（`GpPrivateFontCollection*`）、`50/51` 的 `lpszFilename`（字符串指针）与 `186 IsArrayDim(lpArray As Long)` ⇒ **一处文件就 10 处**；再算上 `.pag` 里 `CHOOSEFONT` / `OPENFILENAME` 那两枚 UDT 的成员（`lpstrFile As Long` 这种），整个工程的 x64 面是**一片**，不是一刀。⇒ 这一格真正的分岔不是"怎么修"（修法就是 `As LongPtr`，逐条改），而是"**要不要在这些 vendored 工程上追 x64**"——仓库既有口径是「32 位 OCX 工程不追 x64」（#163/#175/B27 都是把结论停在"改 VB 源码"、没真去改），而这条工程 x86 今天起窗正常。**这一问今天不再问一遍，按既有口径就地判掉**（#163/#175/B27 同一条，加上那句原话「如果 api 申明没匹配 64 位，首先你得修改 vb 代码才行的，而不是调整编译器源码」）：结论 = **源码形状限制，x64 这一档对这些 vendored 工程不追**，产物继续只在 x86 那档钉用例；#203 就此结在「x64 已知不支持（工程侧）、x86 正常」这条口径上。真要追 x64，那是一批**纯工程源码**的 `LongPtr` 现代化（逐条判定 + 两台各真跑），要用户另立目标才动，不是编译器的事。

### B39 创建期下发的那张默认字体没人存：没被写过字体的控件整张 Font 面读空（账 #204，**已出**）

#200 收了"改过字体读不到"那一半，这半是**没改过字体**的那一半。读数（`.build/b204probe` = dcsurf 的一份**临时拷贝**，`.build/b204_run.bat` 真跑，仓里那份夹具一字节未动）：一枚从头到尾没人写过字体的 PictureBox 答 `bName= bfs=0 bpf=0 bth=16` —— `.FontName` **空串**、`.FontSize` **0**、证人 `FontPixelHeight` **0**、`TextHeight("Xg")` **16**；同一趟里设计期给了 18pt 的那枚答 `aName=MS Sans Serif afs=18 apf=24`（这条是 #200 刚接通的那一路，读得到）。**对照 VB6**：picB 该答 MS Sans Serif / 8.25 / 约 13-14 像素，而 16 那个数是Segoe UI 9pt 的高度 —— 也就是**默认字体那批控件的文字量与 Print 现在全按系统默认字体算**，不只是读空。

根因是 #200 那一处出口的**覆盖面**，不是读法：`vb6forms.c` 的创建路（Fix 181 那一站，`vb6_Vb6DefaultGuiFont()` 每枚控件新建一张 MS Sans Serif 8.25 再 `WM_SETFONT` 过去）与控件数组那条创建路（`vb6forms_ctrlarr.c` 跟着把模板的字体 `WM_SETFONT` 给新窗口）都**只发不存** —— 而 STATIC 那一类窗口不答 `WM_GETFONT`（#200 已用探针钉死），于是 `vb6_ControlFont` 两问皆空。顺带这一处还留着一张没人认领的 GDI 对象：Fix 181 造的那张字体，改前改后都只有等某次字体赋值才可能被删，而删它靠的正是"读得到旧字体"。

修法照 #200 那一条口径走到底：**创建路把刚发给窗口的那张也存进同一个槽位**（`VB6_CtrlFont`），两条创建路都要给（#83/#151/#160 那条"两条创建路读同一个数"的纪律 ⇒ 判据两头钉）。**护栏的形状要先想清楚再动**：哨兵 D10 今天写的是「`VB6_CtrlFont` 的写者恰好 1」，这一刀会把写者变成 3（setter + 两条创建路），所以那条 census 要**按理由改数**、并把每个写者是谁写进规则注释 —— 不是把上限放宽就完事。判据建议照 #153 那套"两头钉"：`bName` 等于 `MS Sans Serif` 且 `bfs` 约 8.25（自存往返）+ `bpf` 在 12..16 之间（问窗口的证人）+ `bth` 从 16 掉到那个区间（**文字量真的跟着换了字体**，这一条才是本账的产品后果）。

为什么单开一账而不是并回 #200：#200 的症状是"改了没生效"，本账的症状是"从来没生效过"，两者的修法在不同文件、判据也不同一头；并在一起会把 D10 那条 census 的理由写得说不清。

**同一条根因的第三形（读码即得，尚未真跑取证）**：控件数组那条创建路（`vb6forms_ctrlarr.c:102`）问模板控件要字体用的是裸 `SendMessage(hTemplate, WM_GETFONT, 0, 0)`，而模板若是 STATIC 那一类这一问**恒回 NULL** ⇒ 外面那句 `if (hFont) SendMessage(hNew, WM_SETFONT, ...)` 整步跳过 —— 运行期新建的数组元素**连一次 WM_SETFONT 都没收到**，拿到的是系统默认字体。⇒ 这三处（#202 的 groupbox 标题带、本账的两条创建路）缺的是**同一件前置**：把 #200 那一处出口从 `static` 提出来、进内部头，谁要读字体都从它走。所以下一批的边界应当是「一处出口递出文件 + 三个站点各归其位」，而不是按账号各修一次；判据两头钉之外还要给数组元素那一形补一条（模板改了字号，运行期新建出来的元素读回来要跟着变）。

**→ 已出（2026-10-05）**：创建路那一站现在把刚发给窗口的那张**同时存进同一个槽位**，并且存的动作收在唯一写口 `vb6_ControlFontStore` 里（setter 也改走它，所以 `SetPropW(..., L"VB6_CtrlFont")` 全仓仍然只有一行）。读数（真跑，x64 与 x86 **逐行相同**）：从没被写过字体的 picB 从 `name= fs=0 pf=0 th=16 / twip=240` 变成 `name=MS Sans Serif fs=8.25 pf=11 th=13 / twip=195` —— **产品后果在最后两格**：默认字体那批控件的 `TextHeight`、`Print` 的行距、缇/像素换算此前一律按 Segoe UI 9pt 算（Fix 181 那一路的字体明明发给了窗口，量的人却问不到）。判据 FR01 三头钉：自存的往返（`fnB = "MS Sans Serif"`、`fsB` 落在 8..9）+ 问窗口的证人（`pfB0` 10..16）+ **文字量真的跟着换**（`tB` 10..16），进 `$dcSurfExpected` 两台各一条；FR02 留原始读数。

**假 needle 真红过**：把创建那一句 Store 注掉、重新内嵌 RTL 再编一台 ⇒ `FR01-DEFAULT=False`、`FR02-RAW name= fs=0 pf=0 th=16`，而 TH05/TH06 那几格跟着一起回到旧数（16/240）—— 这一条判据不是自洽假绿。

**顺手抓到的一条**：picB 运行期改成 20pt 之后 `tB2` 从 27 变成 **32**。理由是 setter 现在从**真的那份** MS Sans Serif 的 LOGFONT 起步、只换 `lfHeight`；改前它从一个问不到的 LOGFONT 起步（NULL ⇒ 那份 lf 是空的），于是"改字号"顺带把字体族也换掉了。⇒ 这一味与 #153 那条"存什么读什么"是同族：**setter 的起点必须是要改的那张**，否则一次赋值会改两件。

**护栏的形状照预告改了**：哨兵新增 D11 —— 出口定义 1 + 内部头声明 1（D10 里那条 `static` 的正则同步改掉，否则提出来就哑）；全 `src/rtl` 里带 `WM_GETFONT` 的代码行 = 1 + **状态条那两处的具名豁免**（见 §B40/#205，豁免设计成"自己会消失"：那两处一旦改走出口，计数从 3 掉到 1 当场红）；`VB6_CtrlFont` 的 `SetPropW` 写者仍恰好 1、存字体的调用站点恰好 3 且必须**分属三个文件**。数组那一路今天运行期不可达（`vb6_CtrlArr_Load` 全仓零调用者，发码侧没有 `Load <数组>(n)` 这条路 —— 与 #117 那片"要靠控件数组"的欠账同一片），接进来是**口径统一**、没写判据。

### B40 状态条那两处问字体被 D11 具名豁免着（账 #205，**已出**）

D11 放开普查范围到整个 `src/rtl` 的那天，只剩两处没接：`vb6forms_statusbar.c:153` 与 `:396`（都是裸 `SendMessageW(hw, WM_GETFONT, 0, 0)`，拿到的同样是那个 NULL）。没顺手改的理由有两条，都得写清楚而不是含混过去：① 状态条是**另一位作者的族**（记忆里那条分工还活着：StatusBar / ListView / ImageList / ProgressBar 接手前先看对方写到哪一步）；② 接上之后**面板宽度会跟着变**（那两处量的就是面板文字），而它的夹具判据今天钉的是什么还没读 ⇒ 改之前要先读 `tests/ctrlstatusbar` 那条口径，别拿一条"我以为钉的是宽度"的猜测去动别人绿着的针。

修法是现成的：两处改读 `vb6_ControlFont`，然后**把 D11 里那条具名豁免删掉**（豁免被设计成一旦删就红，不会烂在那儿）。

**→ 已出（2026-10-05，提交 `a4986c88`，门 #330 = run 37262853393、head 14caaf2a、attempt 1 = **11 job 全 completed/success**，逐片计数全为 FAIL=0：syntax 151/151、bas#1 44/44、asm 13/14、vbp 四片 48+1skip / 54 / 50 / 50（那条 SKIP 是 `test_vbman` 的 COM 32 位视图没注册，早就在那儿）。本轮新用例在 CI 上逐行真 PASS：`[VBP] sbfont`(vbp#3) 与 `[VBP] sbfont_x86`(vbp#4) —— 这就是"字体参与排版"那条判据在两台真跑过；邻居那条没动：`[VBP] ctrlstatusbar ... PASS`(vbp#2)。这一轮的日志**取全了 11/11**：jobs 数组里没有 `log_url`，改走 `/actions/jobs/{id}/logs`，并且**自己处理 302** —— 禁用自动跳转先拿 `Location`，再**不带 Authorization** 取正文（urllib 会把 Authorization 头带到跨主机的重定向上，存储端就答 401）。第一次取时有 4 片回的是「101 行、0 条用例行」那种**截断形状**，重取才见到真数 ⇒ "日志取到了"不等于"读到了"）**：动手前先读了那一位作者的期望表，**结果把判据的形状改了**。`tests/ctrlstatusbar` 那 42 条里与宽有关的只有 SB10-W2=120 与 SB36-SETW=123，两条钉的都是**显式给过的 Width**，而 `vb6_StatusBar_GetPanelWidth` 返回 `e->width`（请求值），排版结果在 VB 侧**根本没有现成的门** ⇒ ① 这一刀动不到那两条针（改完逐条复核：42 条 + 事件那 10 条，两台全在、一字未动）；② 判据得**问窗口本人** —— `SB_GETPARTS`（WM_USER+6=1030）交回各格右边界，Declare 用 `LongPtr` + `ByRef … As Any`（同 LabelPlus 那一形）。

**读数**（两台逐行相同）：改前 `ct8=140 ct20=140 after=140` —— 同一串 "WWWWWWWWWW" 在 8pt 与 20pt 两枚状态条上量出**同一个宽**，运行期把字号改成 20 也不动 ⇒ 字体压根没参与排版；改后 `ct8=110 ct20=260 after=260`。证人 `pf8=11 pf20=27` 改前改后都一样 ⇒ **设计期那两张字体本来就下发到位了**（#204 存的），只是没人去问。画的那一遍（`C3SbPaint`）没有独立判据 —— 观感不进数值面，钉住的是「量的与画的问同一处」这条普查。**两头钉 + 一条护栏**：SF02 两枚同串不同字号必须不等宽（拦住"字体不参与"），SF04 运行期改字号后重排要变宽且与 20pt 那枚对上（拦住"只认设计期、不认运行期"），SF05 **显式给过 Width 的那格右边界不许挪**（拦住"顺手把所有面板都重排"）。

**负控 = 同一份夹具在改前那台真编真跑** ⇒ `SF02-CONTENTS-TRACKS-FONT=False`、`SF04-RUNTIME-FONT=False`、旧数全回来（`.build/b207out/x64.out` / `x32.out`）；这一对 False 就是"新针能红"的物证，不是推论。新夹具 `tests/sbfont` 两台各一条进清单（RAW 那三行只钉前缀 —— 110/260 是 DPI 的函数，#147 那条口径）。

**护栏**：D11 那条**具名豁免按设计自己消失了**（`$rawAll` 从「want 3 = 出口 1 + 豁免 2」收成「want exactly 1」，PASS 行现在打印 `全仓裸问 1`）—— 这正是 #202/#204 那轮把豁免写成"删掉就红"的目的。A/B：这一刀只动 RTL ⇒ 86 份产物**逐字节相同**（`b207_new_emit` vs `b206_new_emit`，diff=0），另外把两份状态条工程也纳入普查面（4 份新快照，BASE 里没有对应份，只作留档不当判据）；相邻三枚哨兵（`check_di_stubs` 624=624、`check_uc_scale_units`、`check_host_pseudo_table` 55 行）各自复跑一遍全绿 —— 新夹具里那枚 `Declare … SendMessageW` 没要新桩。

**顺带量出来一条新账（§B41/#206）**：`Panels(i).Width` 读的是**请求值**而不是排版后的宽（`vb6_StatusBar_GetPanelWidth` 直接 return `e->width`），所以 sbrContents/sbrSpring 那两档在 VB 侧读不到几何。今天**没有**据此改它 —— VB6 那一读数的单位口径（缇还是像素）还没量准，拿猜去改就是给这一族埋第二根雷。
补一句为什么本地量不到基准：全仓没有一份**由 VB6 设计器写出来的**状态条设计块（`grep -rln --include=*.frm --include=*.pag "StatusBar"` 只命中我们自己那三份夹具），所以"设计值与实际宽的比"这条路在这儿取不到证据 ⇒ 这条账要动，得先拿到外部读数（原生 OCX 跑一遍，或 MSDN 原文）。

### B41 `Panels(i).Width` 交回的是请求值，不是排版后的宽（账 #206，开着）

#205 做判据时撞见的：状态条的排版结果只活在 `w->rights[]`（`SB_SETPARTS` 那一份），而 `vb6_StatusBar_GetPanelWidth` 交回 `e->width` —— 设计块没给 Width 的面板一律读回 **0**（实测 `SF01-RAW` 改前那版就是 `small=0 big=0`）。后果不止"读不到数"：`sbrSpring`/`sbrContents` 两档在 VB 代码里没有任何几何可查，凡是按面板宽定位的东西（比如提示气泡、覆盖层）只能自己再算一遍。
开工第一步不是改 getter，而是**把 VB6 那一读数的单位钉准**（`Panel.Width` 文档写的是缇，而 rights[] 是像素；拿 OCX 原生版对照最快，见 `tests/ctrlstatusbar` 那份 .frm 的设计值与实测宽的比）。口径钉错 ⇒ 把"读回 0"换成"读回错单位的数"，比现在更难查。
### B42 设计期 `.frx` 的 List/ItemData 只接了 ListBox 一档，ComboBox 那 17 处全落空（账 #207，**已出**）

Fix 195 那轮把 .frx 三种 blob 的**布局**钉准了（字符串 / 字符串表 / 整数表），但发码侧的接线只写了一档：`emitControlFrxProps` 里 `if (ctrl.controlType == FrmControlType::ListBox)` 才发 `LB_ADDSTRING` / `LB_SETITEMDATA`。语料普查：`List =` / `ItemData =` 指向 .frx 的共 **17 + 17 处，全在 ComboBox 上**（Charts 2020 的 ucTreeMaps / ucChartBar / ucPieChart / ucProgressCircular 四份 demo 的 "Number of Series"、"Chart Style"、"Legend Position" 那一类）⇒ 编出来的下拉框是空的。

**产品后果不是"难看"，是整块图不画**：`ucTreeMaps/Form1.frm:357` 写的是 `If Combo1.ListIndex = -1 Then Combo1.ListIndex = 4: Exit Sub` —— 空组合框 ⇒ ListIndex 恒 -1 ⇒ 每次 Form_Load 都在第一句退出 ⇒ 图体空白（真跑截图为证）。

**改**：不新开机制 —— 同一处出口、按控件型取**消息对**（`LB_ADDSTRING`/`LB_SETITEMDATA` 与 `CB_ADDSTRING`/`CB_SETITEMDATA`），ListBox 那一路发出的文本逐字节不变。两条创建路（顶层 `cgen_form_ctrl_style_apply.inc` / 容器子控件 `cgen_form_frame_menu.inc`）本来就共用这一个 lambda，所以只改一处。

**读数**（真跑，两台逐行相同）：夹具 `tests/frxdata` 加长 —— 一枚 ComboBox 复用**同一份** blob 的两个偏移（证明缺的是接线不是解码器），`FD8-COMBO=True`（ListCount=3 / List(0)="1234" / ItemData(0)=5 / ItemData(2)=-7 四头各自对上）、FD9 留原始读数；**负控 = 改前那台真编真跑同一份夹具** ⇒ `FD8-COMBO=False`、`count=0 item0= id0=-1 id2=-1`（`.build/b211out/b64.out` / `b32.out`，改后在 `g64.out` / `g32.out`）。真工程：ucChartBar 的 demo 四个组合框现在都答得出设计值（"Grouped Column" / "2 Series" / "TOP" / "Aling Left"），ucTreeMaps 的 demo 点 Random 之后**树图整片画出来**（分块、名字、图例 2000..2004）—— 截图 `.build/b211out/bar32.png` / `tm_after.png`。

**护栏 A/B**（BASE = `b207_new_emit`）⇒ inputs=98、same=86、**changed=4**（ucChartBar 两台各 +48 行、ucTreeMaps 两台各 +15 行）、new-only=8（这轮把 LabelPlus / ucChartArea / ucPieChart / frxdata 也纳入普查面，BASE 里没对应份）；逐行归因 = **全部是新增的 combo 发码，0 行删除、0 行无法归因**（分类器要认整段两行形状：`{ wchar_t* vb6_witem = vb6_Utf8ToWide(...)` 那一行不含 CB_*，第一版因此报了 16 条假"无法归因" —— 又是"计数按整行判"那一课）。真工程配对：ucChartBar 两台 rc=0、ucTreeMaps 两台 rc=0。

### B43 程序改 `ListIndex` 该不该发 `Click`（账 #208，**判掉：不许凭猜改产品**）

#207 的夹具里顺手量到的：`Combo1.ListIndex = 2` 之后 `li=2`、`text=你好` 都对，但 `clicks=0` —— 挂在该组合框上的 `Combo1_Click` 一次也没进。VB6 的口径是**程序改 ListIndex 会触发 Click**（`ucTreeMaps/Form1.frm:357` 那句 `Combo1.ListIndex = 4: Exit Sub` 整个就是靠这个惯例来启动首次绘制的：设值 ⇒ 发 Click ⇒ `Combo1_Click` 里 `Clear / Form_Load / Refresh`）。后果：那一页现在必须**人手点一下 Random** 才画得出图。
开工先量三件事，别直接改：① ListBox 那一档同不同形（VB6 两类都发）；② 发 Click 的时机 —— 是"赋值即发"还是"下一条消息才发"（决定重入：`Combo1_Click` 里又调 `Form_Load`，而 `Form_Load` 第一句就是那个赋值 ⇒ 会递归，VB6 靠"赋值时 ListIndex 已改好"让第二次进来不再走那一支，我们要不要同一顺序）；③ 用户点击与程序赋值**不能双发**（同 #171 那条"按来路筛"的纪律）。判据两头钉：程序赋值 ⇒ 恰好 1 次；用户点击 ⇒ 恰好 1 次；重复赋同一个值 ⇒ 0 次（VB6 是不是这样要先量，别照猜钉）。

**#208 的机制已经量到（2026-10-05，读码）**：用户那一路是通的 —— `cgen_form_wndproc_create.inc:396` 把 `CBN_SELCHANGE`(code=1) 映到 `_Click()`（ListBox 那一路在 :316，容器里的在 :523/:544），而 Windows 对**程序**发的 `CB_SETCURSEL` 不回通知 ⇒ `vb6_SetListIndex`（`vb6forms_list.c:52`）设完就没人再发那条消息 ⇒ clicks=0。现成的先例两条：`vb6forms_richtextbox.c:478` 就是"程序化补发一条 `WM_COMMAND(MAKEWPARAM(id, code), hwnd)` 走完整派发链"，StatusBar 的 `SimClick` 同形。⇒ 缺的是"设完值补发那条通知"这一小步，而且**必须只在该控件挂了 Click 处理器时发**（否则又是 #190 那族"发码引用不存在的处理器"）。

**→ 这一条照读码去做了，然后被自家夹具拦下（2026-10-05，同一天判掉）**：补发通知的改动写进 `vb6_SetListIndex` 之后，新夹具两头（combo/list 各恰好 1 次、同值 0 次）确实绿，ucTreeMaps 的 demo 也**不用点 Random 就自己画出图**了（`.build/b211out/tm_startup.png`）—— 但 `tests/ctrlfiles` 的 **CF14 / CF15 当场红**（两台都红）。看它们的写法就知道谁错了：CF14 断的是 `auxList.ListCount = 3` 且**第一条必须是 "dbl:"**，而那三条 item 是夹具自己 `fileList_DblClick` / `fileList_Click` **手工调用处理器**加进去的（`CfForm.frm:106/120`）—— 也就是说这套夹具一直按「程序改 ListIndex 不发 Click」写，而它钉的那份口径来自 VB6 本体。⇒ **本机没有 VB6，这条"VB6 会不会发"我量不了**；两难之间只有一种立场站得住：**不凭猜改产品**。已把 RTL 与夹具全部回退（`git checkout` 那三处），回退后复跑 ctrlfiles 17 条 needles 两台全在、frxdata 回到 #207 的读数。顺带把 ucTreeMaps 那一页"要点一下才画"重新定性：**那大概率就是 VB6 的原样行为**（作者写 `Combo1.ListIndex = 4: Exit Sub` 的意图是给组合框一个默认项，绘制留给用户点 Random），所以 #207 之后剩下的"启动不自动画"**不算缺陷**，本账到此结掉。如果哪天要重开，前置条件写死：先拿到能跑的 VB6（或原生 OCX 的对照实例）量出真口径，再动 `vb6_SetListIndex`。

**重开时要带着的三条旁证（本轮顺手量的，别重新找）**：① **两个 vendored 工程都把这个惯用法当启动路径** —— `ucTreeMaps/Form1.frm:357` 与 `ucChartBar/Form1.frm:524-530`（后者一次设四枚组合框的 ListIndex 然后 `Exit Sub`），作者的意图明显是"设默认项 ⇒ 触发 Click ⇒ 重跑 Form_Load 才画"；若 VB6 不发，这两页在 VB6 里也永远空白，那不太像 released demo 的样子。② 反方向：`tests/ctrlfiles` 的 CF14/CF15 按"不发"写（手工调处理器模拟点击）。③ `VBFlexGridDemo/MainForm.frm:439/448` 也设了 `ListIndex = 0`，而那台 demo 启动画面是完整的（`.build/b208out/f1.png`）—— 但它不依赖 Click 的副作用，所以这条**中性**。本机可查的仲裁者只剩一个：`D:\tools\twinBASIC_IDE_BETA_983`（tB 是 VB6 语义的再实现，它怎么处理"程序改 ListIndex 发不发 Click"至少是一份可比对的证据）。

### B44 动态数组的**元素**访问裸读描述符：未分配 / 越界在 VB6 是错误 9，这里却是原生 AV（账 #209，**已出，门 #333**）

现场是 `b213` 那台交互探针撞出来的：`.build/b211out/bar32/Proyecto1.exe`（Charts 2020 ucChartBar demo，x86）点 `Random` 之后窗口消失。三轮 Windows 应用日志同一条读数 —— `异常代码 0xc0000005 / 错误偏移 0x0001abf2 / 出错模块 = exe 自己`。把同一份工程用 `-g --keep-for-debug` 重编再点，`C3_CRASH_TRACE=1` 给出 `av read target=0xc`，栈里 `uc_hostmodel_call.inc:34` → `ucChartBar.c:591`，那一行是 `With m_Serie(Index)` 发的 `&(VB6_SA_AT(vb6_type_tSerie, me->m_Serie, Index))`。`0xc` 正好是 `vb6_SafeArray1D` 里 `lBound` 的偏移 ⇒ 描述符本身是 NULL，不是 `me` 是 NULL（`m_Serie` 在类结构里差一百多字节）。

为什么是空数组：demo 的 `Form_Load` 走的是「四枚组合框 `ListIndex = ...` 然后 `Exit Sub`」那一支（就是 #208 那条），四枚 `_Click` 处理器一个都没进 ⇒ `AddSerie` 里那句 `ReDim Preserve m_Serie(SerieCount)` 从没执行。**真 VB6 在这种状态点 Random 会弹 Run-time error 9，也不会画图** —— 所以「点一下没反应」不是又一处产品缺陷，反倒给 #208 那条判掉添了一条旁证（作者的启动路径本身就依赖程序改 ListIndex 发 Click）。产品缺口只有裸读这一条。

做法：`VB6_SA_AT`（`vb6rtl_array.h:100`，全仓**唯一**的一维描述符裸解引用点）的宏体改走 `vb6_SaElemPtr(arr, idx, sizeof(type))` —— header 里 `static inline`，热路径三条比较（NULL / 小于下界 / 大于上界），不满足就调 `vb6_SaElemFail`（vb6rtl_array.c，紧挨 UBound/LBound 那两条 rev2），抛 `vb6_ErrRaise(9, "VBA.Information", "Subscript out of range")`；那两条注释里 v1 时代就明写着「`VB6_SA_AT` 读 NULL+0xc → 0xC0000005」，本刀把元素这一半接回同一口径。**步长仍按调用方写明的 `sizeof(type)`**，没改成读描述符的 `elemSize` —— 那是 Fix 170/rev3 记下的独立历史坑，本刀只加检查、不动步长语义。RTL 里那几处内部调用点（`vb6rtl_compat.c` 的 Split/Filter、`vb6rtl_date.c` 的 ArraySet*）本来就自己写着「arr 为空或下标出范围就提前返回」，中心有了检查之后它们那圈守卫从「唯一的防线」变成「提前返回的语义」，没改。

判据：`tests/test_arr_empty.bas` 尾部加七枚私有过程 —— 未分配读 / 未分配写 / `ReDim m(2 To 5)` 的上越界 / 下越界 / UDT 数组的 `With s(2)`（demo 那个形状）/ `Erase q` 之后 / 以及两枚**负控**（`EA-ea-in=20/30/50` 与 `EA-ea-loop=100` 钉住「在范围内的读写照旧」，检查误报就红在它们身上）；顺带把以前根本没钉的 `EA-err=9` 钉上。本地 x86+x64 各真编真跑，21 条 needle 一行不缺。demo 那头是同源的 A/B：改前 `PROCESS-GONE exit=0xC000041D` + AV 轨迹，改后 `GONE code=0x00000009` + stderr `Unhandled error 9: Subscript out of range`。

护栏：发码面零变动 —— A/B 98 份 `--emit-c`（49 份工程 × 两个位数）逐字节相同（`inputs=98 same=98 changed=0`），因为改的只有嵌在 C3.exe 里的 RTL。新哨兵 `scripts/check_sa_access.ps1`（已进 `[STATIC] sa_access`）五条规则，两条负控真红：宏退回裸算术 ⇒ A1+A5 红；冷路径不抛 9 ⇒ A3 红（第一次负控差点骗过：抛 9 那个锚点在文件里有两处，替换命中了 UBound 那条，于是假绿 —— 锚点要从被改的那枚函数起找）。

同族剩下的那一半（**本刀刻意没动**）：多维那一支 `VB6_SA_ND_AT1/2/3`（`vb6rtl_array.h:176/180/185`）与 `cgen_expr_call_prelude.inc:349` 的 4+ 维 `_ndoff_` 兜底仍是裸寻址，存量 1468 处 / 4 份工程（大头 VBFlexGridDemo 每位数 632 处）。哨兵 A4 钉的是「后端只许那一条兜底行」，把它接进检查时 A4 的口径要一起改。 **（已由 §B46 接上：A4 的口径改成「一条都不许有」，另加 A6/A7。）**

CI 读数只到片级：门 #333 = run 37278620691、head d61d9065、attempt 1 = 11 job 全 completed/success，10 片各自 FAIL=0（vbp #1 那片 PASS=48/FAIL=0/SKIP=1，唯一 SKIP 仍是已知的 `test_vbman`）；`[STATIC] sa_access ... PASS` 这一行是逐行读到的。但 bas 两片（PASS=44 / PASS=43，各 FAIL=0）的日志里**没有** `test_arr_empty` 那条用例级行 —— 那片只落一行摘要，所以那 21 条 needle 的逐针读数只有本地那份（x86+x64 各真编真跑），CI 这头不假装有。

### B45 门 #332 那条红不是产品崩了，是夹具在通知到齐之前就读了计数（账 #211，**已出，门 #333**）

门 #332（run 37274000331，head 213f8ad5）attempt 1 与 attempt 2 **同一条红、同一形状**：
`[VBP] pbsub_x86 ... FAIL (output mismatch)`，`Got: PB00-LOAD / PB01-PAINT px=255 ×2 / PB-CNT mdown=0 mup=0 click=0 img=0 lb=0 paint_ok=1 / PB-DONE`。
注意两头：进程 **rc=0 且 PB-DONE 打出来了** —— 不是崩溃，是那五条鼠标通知一条没读到；同一片里 x64 的 `pbsub` 绿。

**先把 #209 摘出去**（三条旁证，不靠“看着不像”）：① `tests_pbsub_*.vbp` 两份 `--emit-c` 里 `VB6_SA_AT` / `VB6_SA_ND_AT` / `SafeArray` **全是 0 处** —— 这一份夹具压根不走那一刀；
② 本地复现出来的坏跑 stderr **空**（既没有 `[SA] elem access out of range` 也没有 `Unhandled error 9`）；③ 门禁用的 RTL 与产物 exe 是同一台 C3.exe 现编的，改的只有 `vb6rtl_array.{c,h}`。

**真因是夹具自己的驱动节拍**：五条通知是 `PostMessage` 发出去的（异步），而计数在**固定第三拍**读。本地把同一份产物在 6 个 CPU  hog 下跑 24 趟，2 趟复现出 CI 那个全 0 形状（22 趟正常）；
把“读到就退”改成“五条到齐才退、最多等 20 拍”再跑，24 趟全绿，其中**有一趟到第 8 拍才到齐**（150ms×8 ≈ 1.2s，远超原来的 3 拍预算）。
⇒ 饿机器上到齐时间可以超过判据的固定节拍，这就是门 #332 两次同形红的原因；#209 只是把产物的二进制布局挪了一点，把这条本来就存在的概率推过了阈值。

**改的是判据的形状，不是判据的强度**：把“等多久”和“断言什么”拆开 —— `allIn = (五条计数都 >= 1)` 每拍重算，`ElseIf allIn Or mStep >= 20`，
`PB-CNT` 那一行**一个字没动**（精确值继续当“翻倍探测器”用），新加一行 `PB-WAIT done=` 把“是不是靠到齐退出的”钉出来（needle `PB-WAIT done=True`）。
**负控真做过**：把五句 `PostMessage` 注掉再编再跑 ⇒ 等满 20 拍，输出 `PB-CNT mdown=0 ...` + `PB-WAIT done=False`，九条 needle 一条不剩 —— 真丢通知照旧红，等待没把它盖住。

验证：加固后的夹具 x86 与 x64 各 24 趟带载全绿（每条 9 针齐），不带载的 sanity 输出对形与加固前逐行相同；A/B 98 份里 96 份一字不动，变的两份就是 pbsub 两片，逐行差异只有 `allIn` 那一族 + 那句 `PB-WAIT`。
下一刀若还要动这一族，记着 §C 第 4 条那份“天生会抖”名单现在多了一份 `pbsub`（它的红是**通知到齐与否**，不是像素/坐标）。
同族普查（同一台、同一批 6 个 CPU hog）：`tests/tabwalk/WalkForm.frm` 相 3 也只用"下一拍读"等异步按键的结果（行 336-349，余量 1 拍），但它相 2 本来就是"到齐才走 + 40 拍保险丝"的写法 —— 真跑 21 趟（1 sanity + 20 带载）读数一字不变（`AK-pre=optA/down=optB/wrap=optA/up=optB/clicks=3`、`TW-ticks=26`），**复现不出就不动夹具**，只把这条留在这里。

### B46 多维数组的**元素**访问同样是裸寻址，而 4 秩那一条连“在范围内”都崩（账 #214，**已出，门 #334**）

#209 收的是一维那一支，多维那一支当时只记了一句“同族下一刀”。这一刀的三条读数全取本地真跑（x86 与 x64 逐行相同，`.build/b216out/ndfix32/run.out` 与 `ndfix64/run.out`）：

· `Dim a2(1 To 2, 1 To 3)` 读 `a2(3,1)` 拿回 **12** —— 就是隔壁那格 `a2(1,2)`。老宏的算式 `(i-lb0) + (j-lb1)*cnt0` 里 i 越界只是滑进同一块 buffer 的下一列，所以它既不崩也不响，是本族里最坏的一种（**静默给错数**）；真 VB6 在这一条是错误 9。
· `Dim d() As Long` 从没 ReDim 就取 `d(1,1)` ⇒ 读 `(NULL)->data` ⇒ 0xC0000005。
· `Dim a4(1 To 2, 1 To 2, 1 To 2, 1 To 2)` **每一格都在范围内也当场崩**（x86/x64 同形）。这一条不是 RTL 的错，是发码侧：`cgen_expr_call_prelude.inc` 那一支把实参拼成 `(int[]){indices[0], indices[1]}` 只塞两枚下标，却按 `actualDimCount` 交给 `vb6_SafeArrayND_Offset` ⇒ 秩 >=4 时读的是 `indices[2]/[3]` 那两块栈上垃圾，偏移成了垃圾再拿去寻址。1/2/3 秩各有专用分支所以一直是对的；这条洞只在下标 >=4 时存在，而这类形状在 98 份真工程产物里一处都没有（`_ndoff_` 计数 0）—— 没被走到才活到今天。

做法与一维那一刀同一形状：`vb6rtl_array.h` 里 `vb6_SaNdElemPtr`（`static inline`，热路径）先问**形状**再逐维问**上下界**，冷路径 `vb6_SaNdElemFail` 抛 9，口径与 `vb6_SaElemFail` / UBound·LBound rev2 一致（有 On Error 走处理器，没有就报错退出）。形状那一问不多读字段：`dimCount` 必须落在 1..16，而一维描述符的首字段是魔数 0x5A1D=23069，天然落不进去 ⇒ 顺带认出“声明 1D、ReDim 成 ND”那族双面形；阈值抄的是 `vb6_LBoundND` 已有的那条，不是新发明的数。步长仍按调用方写明的 `sizeof(type)`，没改成读描述符的 `elemSize`（Fix 170/rev3 那条独立历史坑，本刀只加检查、不动步长语义）。`VB6_SA_ND_AT1/2/3` 的宏体改走它，新加 `VB6_SA_ND_ATN(elemType, arr, rank, idx)` 给 4 秩以上。

发码那一支改成一次把**秩数与全部下标**交出去：`VB6_SA_ND_ATN(T, (vb6_SafeArrayND*)arr, <n>, ((const int32_t[]){...}))`。两处口径值得单独记：① 秩数取“源码写了几枚下标”而不是 `actualDimCount` —— 两者不等时以前是静默少传，现在交给运行期按 `dimCount` 问一句（交回 9，不再拿越界的下标去寻址）；② 那对**最外层括号不是装饰**：预处理器按顶层逗号切宏实参，`{a, b, c}` 里的花括号挡不住它，少了这对括号就是 cl C4002「参数过多」（实测 `P4.bas` 撞出来的）。

判据：新夹具 `tests/test_arr_nd.bas`（`Add-BasTest "test_arr_nd"`，进 bas 那两片）。八条 needle 全取真实输出 —— `NA-udt=5/6`（2 秩 UDT 元素，VBFlexGrid 那一族的形状）、`NA-in=0/23/211/2112`（1/2/3/4 秩各 16 格往返 `bad=0` + 四枚范围内读数，第四个数就是 4 秩那一格）、`NA-dyn-in=102/23`（动态二维 ReDim 后的范围内读写）、`NA-above=9 / NA-below=9 / NA-null=9 / NA-rank=9`、`NA-DONE`。三枚“范围内”读数是负控：检查误报就红在它们身上，只钉 9 的那些钉不住误报。秩数不符那一针钉的是“不再拿越界下标去寻址”，**不是**“这就是 VB6 的口径”——真 VB6 在编译期就拒（“Number of dimensions doesn't match”），编译期判死是另一件事，本刀没做。

护栏四层：
- **发码面** A/B 98 份 `--emit-c` 逐字节相同（`inputs=98 same=98 changed=0`），census `VB6_SA_ND_AT2` 1468→1468、`VB6_SA_ND_ATN` 0→0、`_ndoff_` 0→0。这 1468 处的**调用文本一个字没动**（宏体在嵌进 C3.exe 的 RTL 里），所以 A/B 证的是“没别的东西跟着变”，而不是“这 1468 处编得过”——后者要靠真编译。
- **真编译四片**（VBFlexGridDemo 与 Charts 2020 主工程 × x86/x64）BUILD rc=0 且各出得了 exe。
- **哨兵** `scripts/check_sa_access.ps1` 改口径 + 加规则：A4 从“后端只许那一条 `_ndoff_` 兜底行”改成 **0 条**（两支都不许再手算元素地址），新增 A6（四条宏各 1 处且宏体都经 `vb6_SaNdElemPtr`；inline 定义 1 处，体内形状问句 + 两条上下界比较 + 冷路径调用都在；声明 1 + 定义 1 且真的抛 9）与 A7（后端必须发 `VB6_SA_ND_ATN`、下标实参按 `((const int32_t[]){` 拼，且不再出现 `(int[]){`、不再直接发 `vb6_SafeArrayND_Offset(`）。**七条负控逐条真红**过一遍（后端手算一行 / 退回只塞两枚 / 少那对括号 / 宏体退回裸算术 / 热路径不问界 / 冷路径不抛 9 / 后端直接发 Offset），每轮跑完按字节还原三张源文件并复核 byte-identical；`powershell` 与 `pwsh` 两个壳都绿。
- **真工程运行面**留一条**读数纪律**给下一次：`tests/Charts 2020/Form2.frm:504` 是 `Randomize Timer`，图体数据每次启动都不同 ⇒ “哪几枚 UC 宿主对 hover 有反应”**不能当判据**：同一枚**老** exe 三趟就给出 `[5,6,7,8,9]` / `[4,5,6,7,8,9]` / `[5,6,7,8,9]`，基线像素 MD5 每趟都不同。该问的换成“新产物有没有走进这一刀”：新 charts 三趟 + 新 grid 两趟带 `C3_SA_TRACE=1` 的真跑，`saNd=0 / sa1d=0 / unhandled=0`，一条通知没少、一个进程没丢 ⇒ 1468 处调用点在真工程里**全在范围内**，加检查不改变任何可观察行为。“图真画出来了”那一面本来就由门禁的 `Charts2020`（`-DumpMinColors 40`）与 `FlexGridX86` 两片钉着。

CI 读数：门 #334 = run 37287453662、head 323ab077、attempt 1 = 11 job 全 completed/success，逐片 `FAIL=0`（smoke 1/1、compile 24/24、asm 13/14、bas 两片各 44/44、vbp 四片 48+1skip / 54 / 50 / 50、syntax 151/151）。`[STATIC] sa_access ... PASS` 这一行是在 Tests (compile) 那片**逐行**读到的。bas 两片仍旧只落摘要行（那片 `用例行=0`），所以这一轮的用例身份改按**两片之和**归因：#333 那轮是 44+43 = 87，这一轮是 44+44 = 88 —— 一片从 43 长到 44，正好是本轮新增那枚用例；这也是上面那条「插行要插在整条语句之后」被修好之后 CI 侧的旁证（插在多行实参中间的话，那两片会在启动阶段就 ParserError）。那八条 needle 的**逐针**读数只有本地这两份（x86+x64 各真编真跑），CI 这头不假装有。`Build C3.exe` 那片这次取回的正文是 37 KB 的“截形状”（一条用例行都没有）—— 按 §B40/#205 那条纪律，那不等于那片没跑，它的结论 success 与其余 10 片都读到了。

同族**下一格已经立账（#212）**，口径先钉在这里：这一刀与 #209 都把「错误 9」交回给了运行期，而 VbEclipse 那批越界点全在 UC 的**实例方法体内**（`m_Serie` / `m_valArray` 那类成员数组），所以真正的未知数不是「抛不抛得出 9」，是**抛的那一跳会不会把实例栈留在中间态** —— 已登记的那条只量过「没处理器 ⇒ Unhandled error 9 + 进程按 9 退出」这一面。要补的是两面：① 方法体内写了 `On Error GoTo` 的，9 有没有落回**它自己**那个处理器（而不是越过它落到调用方，或干脆落不回）；② 抛过之后那枚对象的成员照旧读得出数（`With` 那一族的现场就是取 `&(me->m_arr(k))` 当对象，跳走之后 `me` 的实例栈必须已经解链）。这两面随后当场就量了：探针（`.build/pe212/` → 升格成夹子 `tests/pberr/`）**两台全绿**，读数、三枚负控与登记见 **§B47** —— 结论是这一格**没欠产品改动，只欠判据**。

工具事实两条（与 §C 第 14 条同族）：① 往 `run_tests.ps1` 插新用例只许插在**整条语句之后** —— 本轮把 5 行插进了上一枚多行 `Add-BasTest` 的实参中间，`git diff --numstat` 是 `+5/-0`（一个字没删）却把那条语句劈成两半；`[PSParser]::Tokenize` 一量就是错，改完 0 错才算过。所以“+N/-0”不是插入点合法的证据，插完必过一道 parse check。② `git remote -v` 会把 remote URL 里内嵌的 token 原样打到屏幕上；取 remote 用 `git ls-remote` 或 `git config --get remote.<n>.url` 并且只回显 sha。


### B47 抛 9 这一跳在**实例方法**里两面都通（账 #212，**已出：判据夹子 `tests/pberr`，门 #335**）

起因接 #209/#214：把「裸读越界」换成「抛 9」之后，VbEclipse 那批调用点全在 UC 的**实例方法体内** —— `With m_Serie(Index)` 发的就是 `&(me->m_arr(k))`，取成员数组那一格当对象用。当时已登记的读数只有「没处理器 ⇒ stderr `Unhandled error 9` + 进程按 9 退出」这一面，另外两面一直没量过：① 方法自己写着 `On Error GoTo` 时，9 会不会**越过**它落到调用方（或干脆落不回）；② 跳走之后那枚实例的状态还读不读得出 —— longjmp 若没把实例栈解链，现场正是「错误被吃掉 + 对象已坏」那一族最难查的形状。

做法：**这一刀没有产品改动**（`src/` 一字未动，C3.exe 与门 #334 那台是同一台 ⇒ 没重发 A/B，改动面只有 `tests/`）。产出是把两面钉成判据：新增夹子 `tests/pberr/`（`PbErr.vbp` = `PbErrMain.bas` + `PbHold.cls`，`ExeName32` 与文件名同名，按 §C 那条「exe 名只有一个权威」的口径），四枚方法分开问 —— `OneDimAbove` / `TwoDimAbove` / `NullDyn` 各自写着处理器（分别走 #209 那一支、#214 那一支、以及从没 ReDim 的成员动态数组），`NoHandler` **故意不写**，`At(ix)` 是「抛过三次之后读成员」那枚证人；驱动里调用方再写一层处理器接 `NoHandler`。

读数（x86 与 x64 **逐字节相同**，`.build/b219out/{x64,x86}.out`；两台 `BUILD rc=0 / RUN rc=0`）：

    PE-1d=9 / PE-2d=9 / PE-null=9 / PE-caller=9 / PE-state=20-40 / PE-DONE

⇒ 两面都按真 VB6 那一面通：错误落回**它自己那枚方法**的处理器、没处理器时继续往上交给调用方、三次抛出之后 `m_items(2)` 与 `m_items(4)` 照旧读出 20 与 40（实例栈没留在中间态）。所以「#209/#214 把越界交回运行期错误」这件事在实例方法这条路上是完整可用的，不是「抛得出、接不住」。

**三枚负控都是改夹具、不改产品**，每枚都必须让一条登记过的 needle 变红 —— 不然那条针只是装饰：
· A：`OneDimAbove` 的下标从 9 换成 2（范围内）⇒ 不再发生抛出，输出 `PE-1d=20`，针 `PE-1d=9` 红；
· B：`Class_Initialize` 里把 `m_items(4)` 置 0 ⇒ `PE-state=20-0`，证人那条针红；
· C：驱动里删掉 `On Error GoTo Caller` ⇒ 传播上来的 9 没人接，stderr 打 `Unhandled error 9: Subscript out of range`、进程按 **RUN=9** 退出，`PE-caller` / `PE-state` / `PE-DONE` 三条整片消失。
C 那一枚顺带把 #209 的「未处理错误按 9 退出」口径在**实例方法链**上又验了一遍。另注：`Exit Function` 那种「拿掉处理器前的出口」的变异**不是**负控 —— 抛出发生时根本走不到那一行，输出照旧（实测过，所以换成 A 那种「让抛出根本不发生」的改法）。

登记与护栏：`$pbErrExpected` 六条针 + `Test-Vbp "pberr"` / `"pberr_x86"`，插在 `erase_sub_x86` 那条**整语句之后**（不是插进多行实参中间），按 §B46 那条新约束复跑 `[PSParser]::Tokenize` = **0 错**；`git diff --numstat tests/run_tests.ps1` = `+14/-0`。夹具三份源文件从工具写出的 LF 归一成仓里用的 CRLF（`crlf/ lone_lf=0`）之后**重编重跑过一遍**，六条读数一字不变 —— 归一化也算一次改动，改完要重测。

已知边界（不装绿）：这钉的是**工程类实例方法**那一路，不是 UC 宿主里带 `Extender` / `ScaleMode` 的那一路（#175/#179 那一族的调用依赖面）；只断言 `Err.Number`，没钉 `Err.Source` / `Err.Description`（那三条口径 #209 记过）；`On Error Resume Next` 那一形在实例方法里的行为这一格没问。用例级 PASS 行只有 CI 那份能给（本地没有跑单条 vbp 的入口，`run_tests.ps1` 只有 Category/分片），所以这一格的「针真的被断言」证据是上面那三枚负控，不是本地一条 PASS。

### B48 同一条口径在仓里写了两份，发码那份对位运算恒答 Boolean（账 #216，**已出，门 #337；中间门 #336 红在自己身上，见本节末段**）

起因是账 #214 那条路上顺手量到的形状：位运算数一进**字符串上下文**就打 True/False。探针（`.build/pcprobe/`，改前 `o_base.out` → 改后 `o_new.out`）—— `"x=" & (34 Or 51)` 发 `True`、`CStr(a Or b)` 发 `True`、`Left(a Or b, 2)` 发 `Tr`、`"x=" & (Not 5)` 发 `True`，而**同一条表达式先赋给 Long 变量再打印是 51**。发码那头 Fix 039 早就把两侧化成 int32 再做 `& | ^`，所以缺的不是算，是**问类型**那一步答错。答错会外溢成三件事：字符串上下文选 `vb6_CStrBool`、装箱走 `vb6_VariantBool`、COM 实参走 `vb6_ComPackBool` —— 第三条在真工程里是响的：`tests/VBFlexGridDemo/MainForm.frm` 的 `Render(hDC Or 0, X Or 0, Y Or 0, CX Or 0, ...)` 一直在把设备上下文句柄按 **VT_BOOL** 交给 COM。

结构上的关键：这条决定在仓里**写了两份**。语义层那份（`src/semantics/semantic_analyzer_expr.cpp` 的 `visit(BinaryExpr/UnaryExpr)`）从一开始就是 VB6 的口径 —— 两侧都 Boolean 才 Boolean，否则数值提升，否则 Variant；`Not` 也分开答。发码层那份（`src/backend/cgen_util_type.cpp` 的 `inferExprType`）却对 `And/Or/Xor` 与 `Not` **无条件** `return Vb6Type::Boolean`，而 `Eqv/Imp` 更漏进了算术支路（只有语义层把它们认成位运算）。两份各写一份 ⇒ 修一份必留另一份，所以这一刀不是"再补一条规则"，是把规则收进 `TypeSystem::bitwiseResult` / `TypeSystem::logicalNotResult` **一处**，语义层与发码层四个点全调它。

落地的口径（`src/semantics/type_system.cpp`）：`bitwiseResult(a,b)` = 两侧都 Boolean→Boolean，两侧都数值→`promote`，其余→Variant（非数值不再落进 promote 拿 String，`"a" And "b"` 在 VB6 是 Type Mismatch，交运行期）。`logicalNotResult(t)` = Boolean 与 Variant 跟着操作数走；Byte/Integer→Integer（答 Byte 会让 `b = Not b2` 绕过溢出检查、把 -1 静默 wrap 成 255）；Long/LongPtr/LongLong/ULong 原样；Single/Double/Currency/Decimal→Long（VB6 的 `Not` 先把操作数化成整数）；其余→Variant。语义层原先 `Not Byte`→Byte、`Not Double`→Variant 两条随收口一起改到同一口径 —— 实测带动的发码是 0（见下面第二次 A/B）。

判据：`tests/test_bitops.bas` 二十四针，x86 与 x64 各真编真跑、输出逐行相同（`.build/b220out/{x64,x86}.out`）—— 数值那一面 `BF-or-lit=51 / BF-or-var=51 / BF-xor=17 / BF-and=34 / BF-not=-6 / BF-int-or=3 / BF-mixed=-1 / BF-cstr=51 / BF-left=51 / BF-byte-and=80 / BF-byte-or=95 / BF-byte-not=-86 / BF-int-not=-2 / BF-dblnot=-3 / BF-eqv=-7 / BF-imp=-5 / BF-assigned=51 / BF-sum=85`，布尔那一面 `BF-bool-or=True / BF-bool-and=False / BF-boolbox=False / BF-cond=hit / BF-boolcond=miss` 一并钉住，拦"一路改回去全推成数值"那个反向错。负控 = 把两处权威毒成"恒 Boolean"（也就是改前那份答案）重编一台：16 条数值针全变 `True`/`Tr`（含七条新针），而那五条布尔面针与 `BF-assigned=51` 一条不动 ⇒ 夹具两头都咬得住。还原之后 24 条读数与 x86≡x64 复核通过（`grep -c POISON` = 0）。

发码面 A/B（100 份 `--emit-c`）：`inputs=100 same=90 changed=8 new-only=2`（新-only 那两份是本轮才进 A/B 名单的 `tests/pberr`）。census：`vb6_CStrBool(` 162→142、`vb6_VariantBool(` 564→556、`!= 0) ? 0 : -1` 864→864、`vb6_VariantToLong(` 1044→1044。八份差异里最大的是 VBFlexGridDemo 两片（95 块、+170/−165），归因法：先把 `_vcmp_<n>` / `_vb6_with_<n>` / `_ndoff_<n>` / `_tmp<n>` / `_vb6_select_<n>` 归一，再按**整份文件里每行的出现次数**比（不是按 diff 块 —— 块里有一行带关键字就当整块有解释，那是假归因）⇒ 次数有变的行 36 条，逐条读，全部落在三类，没有第四类：① 收窄目标上的位运算赋值补上范围助手（`BufferVT = …` 16+16 条改走 `vb6_ChkInt`，另有 `vb6_ret_CalcHash`→`vb6_ChkLong`、`KeyCode` 3+1 条、`VT`、`vb6_ret_LoWord/MakeWord/LoByte`→`vb6_ChkInt/ChkByte`）；② `(X And k) = k` 那一族从内联 `== k` 改走 `vb6_VarCmpLongEq` + 一枚装箱临时，每片净 +5 行，正好是新增的五条 `vb6_VARIANT _vcmp_N = vb6_VariantFromValue((int32_t)(X & k))` 声明行（83743→83748、83719→83724）；③ COM 实参 `vb6_ComPackBool((hDC Or 0))`→`vb6_ComPackInt(...)` 三条，就是开头那条真工程症状。

第二次 A/B（专给"收成一处"这件事本身）：权威版重扫 100 份，与手搓那版的产物**逐字节相同 100/100** ⇒ 把规则从两处搬到一处、并让语义层也改调它，带动的发码为 0；顺带证明语义层那两条改动（`Not Byte`、`Not Double`）在真工程面上没有消费者。

真编译：grid/charts × x86/x64 四片全过（`.build/b220out/demo/`，每片输出目录是本次新建 ⇒ 判据只认这次跑出来的 exe：mtime 20:20–20:21，四片各 `error C`=0、`LNK`=0，exe 1.67 / 2.03 / 0.94 / 1.15 MB）。记一条读数纪律：那台 cmd 上 `echo GRID32_BUILD=!B!` 把 `!B!` 原样打了出来（延迟展开没生效）⇒ **rc 在这一片不是判据**，别拿它当"编过了"。

哨兵：`scripts/check_bitwise_authority.ps1`（B1 定义各 1 份、B2 声明各 1 份、B3 两层四个调用点合计 ≥4、B4 两份消费点里不许再手写作答）。当前绿 `PASS bitwise authority: defs 1+1, callsites 4, inline answers 0`；负控是把那两份消费点退回 HEAD 的版本再跑同一枚哨兵（`.build/b220_headtree/`）⇒ B1/B2/B3/B4 **全红**，且 B4 直接点名改前那两条：`BinaryOp::And || bin.op == BinaryOp::Or || bin.op == BinaryOp::Xor) return Vb6Type::Boolean` 与 `UnaryOp::Not) return Vb6Type::Boolean`。已接进门禁 compile 那片：`[STATIC] bitwise_authority`。

门 #336 那一次红**红在我自己的登记行上，不是产品**：两条 `Add-BasTest "test_bitops" "$Tests\test_bitops.bas"` 落盘成了 `tests` + TAB + `est_bitops.bas` —— 我在生成脚本里写的是 `\t`，经工具层折一级反斜杠后 Python 拿到的是 `\t` 的转义形式，于是**制表符吃掉了一个字母 `t`**。两片 bas 在 11 秒时被仓里那枚注册表自检 `[FATAL] 路径里有控制字符` 打掉。三条一起记下：① `[PSParser]::Tokenize = 0 错` 拦不住这一条（真 TAB 是合法的字符串内容），拦它的是 `Assert-TestRegistry` 那枚自检 —— **门靠它在 11 秒时报错，比跑完整片便宜得多**；② 第一次修只把 TAB 换成反斜杠，得到 `\est_bitops.bas`（仍少一个 t，而且再也看不出曾经错）⇒ 修完必须**按字节断言目标串**（`chr(92)+'test_bitops.bas"'` 在里面、整份文件 `chr(9)` 计数 = 0），别只看 repr（`\` 与 `\<TAB>` 在 repr 里几乎一样）；③ 我本地只用编译器直 + 一枚 .bat 跑过夹具本身，**没经过注册表那条路** ⇒ 本地全绿而门红在登记面上：改登记面就要让登记面自己也跑一遍。同轮 vbp #4 的 `tabwalk ... FAIL (output mismatch)` 与本轮无关：它的 `--emit-c` 产物在改前基线（`b216_new_emit`）、手搓那版与权威版三份里**逐字节相同**（x86/x64 皆然，27303 字节）⇒ 归到已登记的 #213 夹具节拍余量；修完登记的 #337 上 `tabwalk` 与 `tabwalk_x86` 两片复绿。

边界（这一格没做的）：① `exprYieldsVbBoolean`（`src/backend/expr/cgen_expr.cpp`）问的是"这条表达式**产出**的是不是 VB 布尔"，服务于 `Not` 的发码形状，与类型口径重合的那部分（比较 / TypeOf / Not）实测一致，但它不是同一件事，没并进这处权威；② 一元负号 `Negate` 那一份仍是语义层写 Boolean→Integer、发码层透传，两份也不一致，这次没碰（没测到消费者）；③ 浮点参与 `Not` 时 VB6 的取整是银行家舍入，只钉了 `Not 2# = -3` 这一条整数值；④ String 参与位运算的运行期错误号（VB6 是 13）与 `Eqv/Imp` 的溢出行为都没问。

### B49 体级声明有四条路、两种形状，Dim 那份副本还落在后面（账 #215，**已出，门 #338**）

起点是账 #216 收尾时顺手量到的另一件事：过程体内 `Const A = 1, B = 2` 之后再用 A、B，**每一枚名字各报一条 VB3001**（不是"只登记第一个"，是一枚都不登记）。同形状的 `Dim x As Long, y As Long` 却一条不报 —— 说明体级声明这条路有**两种形状**在并存。`--dump-ast` 一眼看穿：`Dim a, b` 出的是**两条 LocalDecl**，`Const A = 1, B = 2` 出的是**一条 LocalDecl 里装着 MultiDecl**；而 `semantic_analyzer_stmt.cpp:251 visit(LocalDeclStmt)` 的 switch 只认 `VariableDecl` 与 `ConstDecl` 两个 kind（src/semantics 里 `MultiDecl` 一个引用都没有），default 支路直接什么都不做。发码那侧从 `cgen_localdecl.cpp:31` 有 MultiDecl 分支，所以**代码照发、值照对**（真编真跑 `V=3`），只是名字在符号表里不存在。

四条路、两种形状的来源：`parseDimStmt` 当年为了支持逗号列表，**在语句层又手写了一遍声明符解析**（自己吃名字、剥后缀、读维度、读 `As`、读初值），而 `parseConstStmtInBody` / `parseStaticStmtInBody` / `parseAccessDeclInBody` 走的是共享的 `parse*DeclList` → 返回 MultiDecl。两份实现从此各走各的：共享那份后来补了两步（`WithEvents`，以及 Task #40 的「VB6 类型后缀即类型声明」——`Dim dl&` 要落 `int32_t` 而不是 Variant），**手写那份没有**。于是实测出一条真值差：`Dim a&, b&` 里第一枚是 `int32_t a`、第二枚是 `vb6_VARIANT b`（探针 `.build/b220c/bd_probe.bas`，改前 `bd_base.emit`）。这个差别 `TypeName` 看不出来（装箱后的 Long 照打 "Long"），只有不给值时 `VarType` 才分得开 —— 改前 `3/0`、改后 `3/3`。

做法（收成一处，不留第二份）：新增 `Parser::wrapBodyDecls(loc, decl)` —— **体级声明只有一种形状：一条声明符一条 LocalDeclStmt**（多枚就地展开成 Block）；四条路全部改走它，`parseDimStmt` 里那 40 行手写展开**删掉**、改调共享的 `parseVariableDeclList`，WithEvents 与后缀即类型那两步因此自动追平。模块级本来就是另一种机制（`parser_module.cpp:202` 展平），这轮没去动它 —— 两处的消费者集合不同，先各自收口，不强行并一条。

判据两头：`tests/test_bodydecl.bas` 八针 x86+x64 各真编真跑逐行相同（`BD-dim=9 / BD-suffix=Long/Long/15 / BD-empty=3/3 / BD-const=345 / BD-static=33 / BD-arr=13/2/3 / BD-variant=Long/String/1 / BD-DONE`），语法片再加一枚 `[CODEGEN-NOTE] bodydecl_one_per_declarator`，正针钉发码里的四行声明（`int32_t u1 = 0;`、`int32_t u2 = 0;`、`const int32_t c2 = 4;`、`static int32_t st2 = 0;`），**Absent 钉「诊断里不许再有 VB3001」**。**负控是真跑出来的**：把那两份 parser 文件 `git checkout HEAD` 退回去重编一台，同一份夹具读出 `BD-empty=3/0` 与 7 条 VB3001（`[CODEGEN-NOTE]` 那枚会直接红），换回来之后 `3/3` 与 0 条；其余六条针改前改后一致，所以它们是**护栏不是靶子** —— 这一格的靶子只有 `BD-empty` 与 Absent 那两枚，写台账要说清，别让八条针看起来都像会红。

发码面 A/B（BASE = 账 #216 收口那台 `b220b_new_emit`，100 份 `--emit-c`）：`inputs=100 same=98 changed=2 new-only=0`，census 四项一动不动（`vb6_CStrBool(` 142、`vb6_VariantBool(` 556、`!= 0) ? 0 : -1` 864、`vb6_VariantToLong(` 1044）。changed 的两份都是 VBFlexGridDemo，按「归一计数器 + 整份文件每行出现次数」的口径归因：delta −276 行，**逐条都是诊断行**（`VB3001` 778 → 502）；产物内容没动 —— `#undef` 去掉缩进后 2082 = 2082、声明行 552 = 552。⇒ 这一格在真工程面上就是**去掉 276 条噪音**，加上"哪天有人写 `Dim a&, b&` 就不再静默落 Variant"。真编译四片（grid/charts × x86/x64）全部本次新建目录出 exe、`error C`=0、`LNK`=0。

哨兵 `scripts/check_bodydecl_shape.ps1`：P1 定义/声明各 1 份，P2 四条路各调一次（调用点合计 4），P3 手写展开不许回来（`parser_stmt_assign.cpp` 里 `expectName("expected variable name")` 必须为 0，而共享那份 `parser_decl_var.cpp` 必须 ≥1），P4 除 `wrapBodyDecls` 内那一处外别处不许把声明列表直接包成 LocalDeclStmt（parser 全范围计数 = 1），P5 语义层**不许**再补 `case MultiDecl`（那等于把两种形状再造一遍，计数必须为 0）。当前绿：`defs 1, callsites 4, hand expansion 0, raw wraps 1, semantics MultiDecl 0`；负控 = 把这四个文件退回 HEAD 跑同一枚 ⇒ P1/P2/P3/P4 共十条红（P5 在 HEAD 上也绿，因为那儿本来就没写分支 —— 它是"保持为 0"的哨兵，不是"抓到本次改动"的哨兵）。已接进门禁 compile 那片：`[STATIC] bodydecl_shape`。

边界与下一格：① 语义层 `visit(LocalDeclStmt)` 里 Dim 那支仍是自己内联造符号（`semantic_analyzer_stmt.cpp:255`），没并进 `registerVariable` —— 不在这一格的问题面上，没动；② `cgen_localdecl.cpp:31` 的 MultiDecl 分支从此 unreachable-by-construction，本轮没删（P4 已经把「只有一处能包」钉住，删它是另一格的清理）；③ `Static Sub` / `Static Function` 两形本来就不是变量列表，没走展开；④ VBFlexGridDemo 里**还剩 502 条 VB3001**，是另一族，**另立新账 #217**（本轮一条没动，只做了 census 与一条重要的读数警告）。502 条按名字分 12 组：UC/PB 宿主词汇 —— `UserControl` 278（全在 VBFlexGrid.ctl）、`PropertyPage` 127（三个 .pag：General 78 / Style 38 / Clip 11）、`Extender` 32、`Ambient` 7；库名成员访问 —— `VBA` 37（ctl 18 / Common.bas 15 / 两枚 .frm 各 2）；VB6 内在常量 —— `vbSrcCopy` 5、`vbPicTypeIcon` 5、`vbPicTypeBitmap` 3、`vbPicTypeEMetafile` 1；另有 `Is` 3（疑似 `TypeOf … Is` 的 Is 被当标识符）、`Interface` 3、工程内常量 `CTRLINFO_EATS_RETURN` 1。**一条必须先处理的读数**：这些告警自己报的 (行,列) 与源文件那行的文本对不上 —— 例如 `UserControl` 的首条指向 VBFlexGrid.ctl:2349 第 4 列，而那行是 `VBFlexGridComboButtonWidth = -1`；`PropertyPage` 指向 `.pag:18` 的 `End`。⇒ 按名字分家的数字可信，**按行定位不可信**（.ctl/.pag 走的是翻译后的虚拟源，行号映射没跟着回来），#217 开工前要么先把定位修对，要么别拿行号做判据。（订正 2026-10-05：这一段里两处猜测是错的。`Is` 的三条来自 `Case Is`，不是 `TypeOf … Is`；`Interface` 那一组压根不存在，三条真名是 `OLEGuids.IObjectSafety` / `OLEGuids.IOleInPlaceActiveObjectVB` / `OLEGuids.IOleControlVB` —— 那是我自己按 UTF-8 硬读 GBK 告警造成的假条目，见 §C16。按 GBK 重读后 502 条落在 **14** 组名字上，一条不差；加上 Charts 2020 的 532 条一起分家，记在 §B50 头部。）

### B50 `Case Is > 2` 里那枚 Is 是 parser 造的，36 条 VB3001 与一枚隐式局部都是它换来的（账 #217 第一刀，**已出：门 #341 唯一红是 frmevents 抖动、#342 全绿**）

先把 #217 的分家钉完（两份真工程各 x64/x86 各一次 --emit-c，GBK 解码后按名字数；502 + 532 = 1034 条，
按成因是**五种 + 4 条未归家**，五种里只有一种(第⑥族)是真缺陷）：① **文档类隐式对象** —— `UserControl` 278+366、`PropertyPage` 127+5、
`Extender` 32+20、`Ambient` 7+74，外加 .pag 里裸写的 `Changed` 8、`hDC` 6、`Controls` 2；发码走的是
`kHostPseudoRows` 那张唯一权威表（账 #159 收的），语义层不认识这些名字 ⇒ 纯诊断。实测发码：
`UserControl.hDC` → `vb6_UserControl_hDC`、`Ambient.UserMode` → `vb6_Ambient_UserMode`、
`Extender.Tag` → `vb6_Extender_Tag`、`PropertyPage.hWnd` → `vb6_PropertyPage_hWnd`。② **`VBA.` 限定** 37 条
（`VBA.Choose` / `VBA.DateAdd` / `VBA.Year`）—— 发码剥前缀走内在函数（`VBA.Choose(j,a,b)` 实测发成
`(j)==1 ? (a) : ((j)==2 ? (b) : NULL)`）⇒ 纯诊断。③ **内在常量缺档** 28 条（`vbSrcCopy` 7、`vbPicTypeIcon` 6、
`vbPicTypeBitmap` 4、`vbPicTypeEMetafile` 1、`vbHitResultHit` 8、`vbAsyncTypeByteArray` 1、`vbAsyncReadForceUpdate` 1）
—— `builtin_consts*.inc` 那张表没登记，但发码照旧出对值：`vbSrcCopy` 折成 `13369376`(SRCCOPY)，
`vbPicTypeIcon` 原样发名、由 `vb6rtl_userctl.h` 的 `#define … 3` 接住 ⇒ 表缺档，且常量现在有**两份权威**。
④ **外部类型库限定** 3 条 = 上面订正掉的那组 `OLEGuids.*`，来自 `Implements OLEGuids.IObjectSafety`，
没有那份注册的类型库可查 ⇒ 源码侧/引用侧的账，不是编译器欠的。⑤ **跨模块 Public Const 看不见** 1 条：
`CTRLINFO_EATS_RETURN` 写在 `Builds\VTableHandle.bas`（`Public Const … = 1`）而在 VBFlexGrid.ctl 里用 —— 发码那侧靠 `#define CTRLINFO_EATS_RETURN (1)` 活着，语义层那条跨模块常量解析待查。
⑦ **未归家 4 条**（全在 Charts）：`ppProgressCircular.pag` 的 `BF` 2 与 `B` 1、`ucProgressCircular.ctl` 的
`Count` 1。名字短得像被截了半截（`BF`/`B` 像在 `&H…` 那一类字面量上、`Count` 像伪对象/集合成员），
但**没量过就不归家** —— 下一格先按 §C16 那两步（显式 gbk + 一枚最小 .bas 探针把行号拿到手）定形。

**⑥ 这一刀出的那族：`Is` 36 条（grid 3 + Charts 33）**。#217 的告警行号本来不可信（.ctl/.pag 是翻译后的
虚拟源），所以拿一枚最小 .bas 探针（`Case Is > 2, 1` / `Case Is > 5` / `Case Is < 0` 三形 + `a Is b` +
`TypeOf o Is Collection`）走 --emit-c：告警落在**这三行本身**（6,17,19），后两形一条不出 ⇒ `Is` 一家当场定形
为 `Case Is`，与 `TypeOf … Is` 无关。读 parser：`parseCaseValue` 为了借一次优先级解析，造了一枚
`IdentifierExpr("Is")` 当左操作数**放进 AST**。语义层 visit(IdentifierExpr) 见到没声明的名字 ⇒ 写着
`Option Explicit` 时每形一条 VB3001；没写时走 VB6 隐式声明那支，把 `Is` 登记成 Variant，发码就在每枚
用到它的过程序言发一枚 `vb6_VARIANT Is = vb6_VariantEmpty();`（同一份夹具去掉 Option Explicit 实测 2 枚）。
发码从头只读比较符与右操作数，所以**值一直是对的** —— 这一族是"AST 里撒了谎、换来两条副作用"，不是算错。

修法（一处形状，不开第二条路）：VB6 的 `Is` 在 `Case Is` 里占的是**测试表达式**的位置，压根不是标识符 ⇒
`CaseClause::CaseValue` 加 `relOp`(BinaryOp) + `hasRelOp`(bool)，parser 里那枚占位标识符只用于借优先级、
出函数就丢，`cv.value` 从此只装右操作数；`cgen_select.cpp` 三档（字符串 / Single / 整数）改读 `cv.relOp`；
`ast_clone.cpp` 两个字段都抄。无比较符的 `Case Is` 那支本来只看 `isIsClause`，值留空即可（改前塞进去的
那枚标识符从来没人读）。

判据（四件，本地全绿）：① 探针两面 —— VB3001 3→0（整个 err 空），隐式局部 2→0，而 `_vb6_select_` 的
条件行逐字不变；② `tests/test_caseis.bas` 7 针真跑 x64+x86 全对（六个关系符各档 `CI-rel=lt/le/eq/gt/ge/ge`、
`CI-ne=is4/ne4`、Is 与值混写 `CI-mixed=big/big/two/rest`、字符串 Select `CI-word=beq/gtm/rest`、
Single Select 那一档 `CI-half=hi/lo/mid`、Variant 测试表达式 `CI-var=hi/lo/hi`）；③ 两条 [CODEGEN-NOTE] ——
`caseis_is_not_an_identifier` 用同一份夹具**两头钉**（`nopeHereIsNotAName` 照旧报 VB3001，`'Is'` 必须不再出现，
拦"把整条诊断关掉"那种修法），`caseis_no_phantom_local` 钉序言里没有 `vb6_VARIANT Is`；④ 哨兵
`scripts/check_caseis_shape.ps1`（C1 两个字段各 1 份 / C2 parser 只设一次 hasRelOp 且不得再把造出来的名字
存进 cv.value / C3 发码读 cv.relOp 三处且不得再把 cv.value 当 BinaryExpr 拆 / C4 clone 两字段各抄一次），
[STATIC] caseis_shape 在 compile 片 PASS，该片 26→27；两条负控各自红过（删 clone 那行 ⇒ C4 红；把假标识符
塞回 cv.value ⇒ C2 红），恢复后按 md5 验过逐字节相同。

A/B 护栏（发码面）：inputs=100 same=90 **changed=10**，每份都是 `+0` 行的纯删除，条数 33/21/9/3/3 各两档
（Charts 主工程 33 = ucChartArea 9 + ucChartBar 21 + ucTreeMaps 3，VBFlexGridDemo 3）—— 逐行看全部是删掉的
VB3001 `Is` 诊断行，**unattributable=0**；CENSUS `'Is'` 138→0，CENSUS `_vb6_select_` 7346→7346（每个 Select Case
站点照旧发码）。这一族的规模是全仓的：138 条里 .ctl/.pag 占大头，说明它在别的工程只会更多。

边界与下一格：① 剩下 96% 的噪声是 ①②③ 三种"发码认识、语义层不认识"，收法与 §B50 这一刀不同 —— 要么把
`kHostPseudoRows` 从 backend 提到 `src/common` 让语义层也问它（一处权威两个消费者，照账 #188 那个先例），
要么在 `namesProjectLevel` 那条豁免位上补文档类隐式对象；动手前要先答一句「语义层把这些名字认识之后，
`lastExprType_` 该不该跟着换档」—— 换档会动类型判定，必须重新做一遍 A/B。② `Controls`/`Count` 那 3 条是
表**刻意不收**的成员（RTL 无对应全局），属账 #159 末尾说的"RTL 侧缺口，另立账"。③ `CTRLINFO_EATS_RETURN`
一条要先量"跨模块 Public Const 在 .ctl 里到底解析不解析"，别顺手并进豁免位。④ §C16 那条读数教训是这一格
最重要的副产品：census 的**名字**是唯一可信刻度，行号与自己解码出来的拼写都要复核。
### B51 文档隐式对象只有"这份文档是哪一类"这一个前提，而那个前提在仓里猜过两处、语义层根本没有（账 #217 第二刀，**已出，门 #342**）

症状是一整片噪声：全仓发码语料里 **VB3001 共 2934 条**（两份真工程 499 + 499，其余在 czUI / ve_units /
Charts 各子工程）。按名字分家就是 §B50 说的 ① 那一族：`UserControl` 644、`PropertyPage` 132、`Ambient` 81、
`Extender` 52、外加 `VBA.` 37 —— 合起来 946 条。发码那侧从来是对的（成员与类型由 `kHostPseudoRows`
回答，账 #159），语义层的 `visit(IdentifierExpr)` 却只认「工程级名字」两个位（`namesProjectLevel`），
于是每条 `<对象>.<成员>` 都当成未声明标识符。更实在的一面在**没写 Option Explicit 的模块**：那条
隐式声明支路会把这些名字登记成 Variant 局部，发码真发出来 —— 探针 `b227out/vb_qual.bas` 读到的就是
`vb6_VARIANT VBA = vb6_VariantEmpty();` 外加一对 `#pragma push_macro/pop_macro("VBA")` 护栏，每枚用到的
过程一枚，从来没人读。

结构上的根因是**同一个前提写了两遍、第三处压根没有**：driver 按扩展名算 `isControlModule` /
`isPropertyPageModule`（`driver_frontend.cpp` 里那几个 bool），发码那边又拿 `controlTypeName` 找
"PropertyPage" 字符串猜第二遍（`cgen_form.cpp` 的 Fix 110f），而语义层两手空空 —— `Module` 上连一个
表示文档类别的字段都没有（`ast_decl.hpp` 只有 isClassModule / isFormModule / isInterfaceModule）。
VB6 的 `UserControl` / `PropertyPage` 只在对应类别的文档里存在，`Extender` / `Ambient` 只有 UserControl 有，
`VBA` 是全局库前缀哪都有 —— 这四条规矩必须落在**一格**里，不能散在字符串猜测上。

收成一处：`common/types.hpp` 加 `DocumentKind { Standard, Form, UserControl, PropertyPage }`，
`Module::docKind` 由 driver **一处**按扩展名写（就在原来写 isFormModule 那个位置），语义层与发码层两处读：
语义层新增 `SemanticAnalyzer::isDocumentHostObject(name)`（(类别, 名字) 一格一格对，四枚文档对象各自认类别、
`vba` 恒真），在 `visit(IdentifierExpr)` 的未找到支里插在 `namesProjectLevel` 之后当第三个豁免位，
**且要求 `memberObjCtx_`**（限定符位）；发码层把 `isPropertyPageDesigner_` 改成读 `module.docKind`
（`emitDesignerControlDecls` 多收一个参数），那句字符串猜测删掉。豁免只压诊断、**不改类型答案**：
`lastExprType_` 仍旧答 Variant，成员的类型继续由发码那张表回答 —— 这一格刻意不碰类型判定，
要碰是另一格（碰就得重做一遍发码 A/B，理由见 §B50 的"边界"）。

判据两头 + 一张哨兵：① 两份真工程逐档清零 —— VBFlexGridDemo 499→**19**、Charts 主工程 499→**34**，
x86 与 x64 读数逐字相同；② 编译面夹子 `tests/dochost/dhExp.ctl`（开着 Option Explicit，五枚符号照旧发出来
而 Absent 钉 VB3001 = 0）与 `tests/dochost/dhImp.ctl`（故意不写 Option Explicit，Absent 钉三枚死局部
`vb6_VARIANT UserControl/Ambient/VBA`）—— 顺带一条工具事实：**.ctl 可以单独喂 --emit-c**，
不必为编译面判据造一整份工程；③ 哨兵 `scripts/check_dochost_authority.ps1`：D1 枚举 1 份四档齐、
D2 **docKind 的写入点全仓恰好 1 处**（多一处就是第二个权威）、D3 谓词 定义/声明/调用 各 1 且调用点带
memberObjCtx_、D4 旧的 `controlTypeName.find("PropertyPage")` 必须为 0、D5 名字表五枚齐全。
[STATIC] dochost_authority 在 compile 片 PASS（该片 27→28）。

发码面 A/B（BASE = 上一刀那台 HEAD `ab1db9c0`，NEW = 这一刀）：inputs=100 same=82 **changed=18**
（9 份工程 × 两档），每份都是 `+0` 行的纯删除，`unattributable=0`；CENSUS VB3001 **2934→158**；
宿主符号一条没动 —— `vb6_UserControl_ScaleWidth` 436=436、`vb6_Ambient_UserMode` 94=94、
`vb6_PropertyPage_hWnd` 16=16、`vb6_UserControl_hWnd` 70=70、`vb6_Extender_Tag` 4=4，
`_vb6_select_` 7346=7346，`VB7006` 0=0（豁免位没把包屏蔽那条诊断一起吃掉，这条是专门钉的）。
隐式局部那一面在这批语料里读不出增量（这些工程全写 Option Explicit，走的是"只 warn 不登记"那一支），
所以那一头靠 dhImp.ctl 钉，不假装 A/B 能证明。

边界与下一格：① 余下 158 条里最大的一族是 ③ **内在常量** 28 条（`vbSrcCopy` 折成 13369376、
`vbPicTypeIcon` 由 RTL 的 `#define` 接住 —— 两份权威，见 §B50 的 ③），已登记为下一格；
② 裸写的伪成员（`Changed` 8 / `hDC` 6 / `Controls` 2 / `Count` 1 + grid 那 1 条裸 `UserControl`）
这一格**故意没管** —— 裸名要按 `HPF_BARE` 放行，风险是遮住真打错的变量名，得单独定判据；
③ `CTRLINFO_EATS_RETURN`（跨模块 `Public Const`）与 `BF` / `B` 那三条未归家的，都要先量再收；
④ 类型答案那一面（语义层认识这些名字之后 `lastExprType_` 该不该跟着换档）没动，动它之前先想清楚
为什么这一格的答案是"不换"：换档会改类型判定，而这一格的发码 A/B 判据只对"纯删诊断行"成立。

**门 #342 落定（run 37345079456、head `910b37b8`、attempt 1）= 11 job 全 completed/success。**
compile 片 27→28 里新那枚就是 `[STATIC] dochost_authority ... PASS`（`caseis_shape` 同片照旧绿）；
syntax 片 154→156 = `dochost_ctl_no_undeclared` 与 `dochost_no_phantom_locals` 两枚进门禁且绿；
bas 两片 47+47 与 #341 同（第一刀的两枚 test_caseis 已在里面）；vbp 四片 50/54/49(+1 SKIP)/51，
唯一的 SKIP 仍是 test_vbman（COM 未注册 32-bit 视图，与 #335/#337/#339 同形）。
**顺带把 #341 那条红结掉**：同一份 frmevents 夹具在 #342 `PASS`（vbp #2，54/59，红=0），而 #341 那次
是 `.out/.err` 双 0 字节、距上一条 PASS 只 3 秒、发码逐字节相同 ⇒ **两刀都不是它的因**，账 #79 那条
"拖放点按窗口位置现算"的抖动换了症状出现（从"少几条 EV"变成"整个进程没输出"），本账没动它，
继续留在 #79。


### B52 `Picture.Line` 尾部的 B / BF 是语法不是名字，而 RTL 用两枚**外部链接的裸名 C 全局**接了它三年（账 #220，**已出：门 #343 十一 job 全绿**）

症状硬得没有歧义：一份工程只要有个叫 `B` 的模块级变量就**编不过**。探针 `b228out/clash_b.bas`
（`Public B As Long` + `Public BF As String`，Main 里 `B = 7`）在改前实测 **BUILD-RC=1、5 条诊断
（C2373 重定义 ×4 + C2166 赋值给 const 对象）、exe=False**；同一条语句 `Picture1.Line (x,y)-(x2,y2), c, BF`
在 `ppProgressCircular.pag` 里其实只有 **3 条源语句**（297 / 299 / 474），但一份工程发码出 **8 条调用**（`DrawPalette` 4 + `PropertyPage_Initialize` 4），语料四份投影各 8 条 —— 这 8 条全靠那两枚全局才落得下地。

成因是三段接力的最后一环：Fix 102 把 `(x1,y1)-(x2,y2)` 吸收成 Line 的实参表之后，尾部 `, color` / `, BF`
是按普通实参 `parseExpression()` 发出去的 —— 名字进了 AST，发码就原样发裸名，RTL 那侧为了让它有个落脚处
写了 `const int32_t B = 1; const int32_t BF = 2;` 并 extern 出去。问题在 VB6 的名字空间与 C 的名字空间
**是共享的**：生成的模块 C 会 #include 那批 RTL 头，而用户模块级变量在 C 里同样是裸名 —— 于是头里 extern
的在编译期撞，.c 里非 static 定义的还会在链接期撞（LNK2005），只有 static 的不撞。

修法收成一处：parser 在 Line 的**尾部循环**里认 style 那一格（`trailingIdx == 1`，即 color 已经给出），
把 `B` / `BF` 折成 `LiteralExpr(Integer)`，值沿用删除前两枚全局的 1 / 2（本刀刻意不改语义）；RTL 的两枚定义
与两行 extern 一起删。**位置口径是关键**：`Line (a,b)-(c,d), B` 那一格按 VB6 是 **color**，用户的 `B` 必须
照旧成立 —— 折错格就是"修一处撞车、制造一处静默错值"，所以夹具两头都写。

判据三件：① 真跑 `tests/test_nameclash.bas`（x64 + x86 两片）—— 改前那份就是 no exe，改后
`NC-B=13 NC-BFLEN=2` / `NC-ACC=15` / `NC-BOX=3/8` / `NC-DONE`，数值、串长、累加、装箱四形都归了用户；
② 发码面 `tests/pcline/PcForm.frm` 的 [CODEGEN-NOTE] `pcline_flag_folded` —— 四条 needles 钉位置
（两条 style 格折成 `vb6_ComPackInt(2)` / `(1)`，两条 color 格留 `vb6_ComPackInt(B)` / `(BF)`），
两条 Absent 钉改前语料实测的那两形 `vb6_ComPackValue(BF)` / `(B)`；③ 哨兵
`scripts/check_rtl_naked_names.ps1`：N1 B/BF 在 RTL 里 0 份、N2 **非 static 的裸名文件作用域数据全局 =
一份钉死的名单**（现在 5 枚：`Changed`（账 #219 欠的）、`g_hoCount`、`g_uc_descCount`、`g_uc_recCount`、
`g_uc_dumpSeq`；多一枚就红，缩小要在同一次提交里改名单）、N3 折旗标只许一个地方（style 位守卫 1、
交出字面量 1、**旗标名字进 AST 必须为 0**）、N4 值口径 B=1 / BF=2 钉死。负控实测：往
`vb6rtl_com.c` 插一行 `int32_t b220probe = 0;` ⇒ N2 立刻红并点名，撤掉复绿。

发码面 A/B（BASE = HEAD `a7a929c9` 那台，NEW = 这一刀，inputs=100）：**changed=4**（ucProgressCircular
那两份 × 两档），其余 96 份一行没动；每份 +8 行 / −8 行，且每一行差异都只在最末一格 ——
`vb6_ComPackValue(B|BF)` → `vb6_ComPackInt(1|2)`，行内其余字符逐字相同；外加 3 行 VB3001 纯删
（ppProgressCircular.pag 63/65/240 那三处），Charts 主工程 VB3001 34→**31**。

边界与下一格：① **本刀只把名字归还用户，没让 Line 画出来** —— 那 8 处现在仍是
`vb6_ComCallObject(vb6_ComGetObjectProp(vb6_hwnd_Picture1, L"Line"), L"Item", {...}, 6)`，宿主应答表里
`Line` 从没登记，运行期是**静默空转**（登记为账 #221）；② `B=1 / BF=2` 这两个值沿用旧全局，没有对过类型库
（VB6 文档那一面是 B/C/F 三个位，`BF` 到底是 2 还是 1|8 待查），归 §B51 边界 ③ / 账 #218 那一族，
先读库再动；③ `Changed` 那枚裸名全局仍在名单上，由 #219 收（**同一轮已把它定罪成实测**：`Public Changed As Long` 现在编不过，C2371 重定义 —— 见 §C18 末段）。
**门 #343 落定（run 37350249122、head `91a004fe`、attempt 1）= 11 job 全 completed/success、逐片 FAIL=0。**
bas 两片 47→**48** 与 47→**48** = `test_nameclash` 与 `test_nameclash_x86` 进了门禁（那两片只落摘要行，
所以逐针读数算本地那份：x64 真跑 `NC-B=13 NC-BFLEN=2` / `NC-ACC=15` / `NC-BOX=3/8` / `NC-DONE`）；
compile 片 28→**29**，新那枚逐行读到 `[STATIC] rtl_naked_names ... PASS`；syntax 片 156→**157**，
逐行读到 `[CODEGEN-NOTE] pcline_flag_folded ... PASS`；vbp 四片 50/54/49(+1 SKIP)/51 与 #342 同形
（唯一的 SKIP 仍是 test_vbman，COM 未注册 32-bit 视图），asm 13/14、smoke 1/1。


### B53 `Picture.Line` 的八处调用两跳都返回成功、一笔都不画：兜底把方法当属性取，而 RTL 两处都登记成「认识但什么都不做」（账 #221 = C29-PL-a，**已出：门 #344 红在自己的名单针 → 门 #345 十一 job 全绿**）

这条缺陷的形状是**没有任何诊断的**：源侧 `ppProgressCircular.pag` 三句 `Picture1/2.Line (…)-(…), c, BF`
（297 / 299 / 474，一份工程发码出 8 条调用），发码交出去的是
`vb6_ComCallObject(vb6_ComGetObjectProp(vb6_hwnd_PictureN, L"Line"), L"Item", {…}, 6)` ——
先把**方法名当属性**取回一个对象，再对那个对象取默认成员 `Item`。而 RTL 两处都把 Line 登记成
"认识但不做事"：`vb6forms_axcontainer.c:304` 的属性位交回 `axSetEmpty` + `S_OK`（**Empty，不是 IDispatch**，
所以第二跳根本没有可调的对象），`uc_hostmodel_call.inc:152` 干脆 `return 1`，注释写着
「宿主对象: 未知方法一律空实现」。⇒ **两跳都成功、两跳都空**，编得过、跑得起、退出码照旧 0。
这已经是本线第三条「控件方法落进 COM 兜底 = 运行期静默空转」（账 #143 的 SetFocus、账 #196 的
hDC/TextHeight/ScaleX 各一条），差别只在 Line 是**六个实参**那一形。

需求面先量了一遍：`Circle` / `PSet` / `Point` / `PaintPicture` 在 Charts 2020 / VBFlexGridDemo /
czUI-main / VbQRCodegen-master 四份真源码里 **0 处**，`Cls` / `Print` 早已是原生出口
（语料里 `vb6_ControlCls` 14、`vb6_ControlPrint` 2）⇒ Line 是这条绘图面上最后一个裸着的。

顺带把账 #220 欠的那半收掉：那一刀为了不改语义，折出来的还是旧 RTL 全局的 `B=1 / BF=2`，
只认两个词。现在按**字母位**折 —— `B=1 矩形 / C=2 椭圆 / F=4 填充`，按串里每个字母置位，
所以 `BF=5`、`CF=6`，`C` / `F` / `CF` 三形从此有了落脚点。这张位口径是 parser 与 RTL **同一张表**
的两头，哨兵 N5 就是钉这个：RTL 里 `style & 1` / `& 2` / `& 4` 各必须出现、`vb6_ControlLine` 必须 1 份。

实现是三处 + 一处判定：RTL 新增 `vb6_ControlLine(hwnd, x1,y1,x2,y2, color, style)`
（DC 走 `vb6_ControlPrint`/`vb6_ControlCls` **同一处** `vb6_ControlDrawDC`，坐标按
`vb6_WindowScaleModeSelf` + `vb6_ScaleUserToPx` 折 —— 那是 #196/#197 的单位表，不另算一遍缇/像素；
`color<0` 交回控件自己的 ForeColor；不填充那档必须显式 `NULL_BRUSH`，留着上一枚画刷就会顺带填一块）；
后端加一张 `controlCanvasMethod` 表（照 `controlZeroArgMethod` / `controlOneArgMethod` /
`controlScaleMethod` 那三张的规矩：**表只交名字、实参由码头拼**）；码头在
`cgen_expr_call_callee_withm.inc` 的 Fix 185 那块之后接；**成员侧**
`cgen_expr_member_form_builtin.inc` 的 PictureBox 那块要让 Line 也打标记。

最后那一处是这一刀自己踩出来的：表与码头都写好后重编一台，产物**照旧**是
`ComGetObjectProp(…, L"Line") + Item`，夹具四形全 False —— 因为成员侧没给标记时，兜底已经把
`comObjExpr_` 换成 HWND 表达式（那段注释里 Fix 023e/089d 说的正是这件事），调用侧那张表根本收不到标记。
于是哨兵加了 N6 第四条（成员侧必须问表），把"只接一头"这种半成品形状钉死。

判据两面，都真跑：① 编译面 `tests/pcline/PcForm.frm` 的 [CODEGEN-NOTE] `pcline_flag_folded`
换了针 —— 四条正针钉原生形状与其位置（`…, (int32_t)255, (int32_t)5)` 是 BF、`(int32_t)1)` 是 B，
`(int32_t)B, (int32_t)0)` / `(int32_t)BF, (int32_t)0)` 是用户自己的名字**留在 color 格**、style 缺省 0），
三条 Absent 钉 `ComGetObjectProp(vb6_hwnd_picA, L"Line")` 与改前那两形 `vb6_ComPackValue(B|BF)`；
② 运行面新夹具 `tests/pcline/PcDraw.{frm,vbp}`（门禁 `pclinedraw` + `pclinedraw_x86`）—— 画完**问像素**：
`PL01-LINE` 对角线那枚必须红而旁边那枚必须不红、`PL02-BOX` 边框蓝而中心不蓝、
`PL03-FILL` 中心必须绿（这条同时证明 BF 与 B 折出来不是同一个数）、`PL04-CIRCLE` 沿 y 开小窗口找到切点
且中心不红。四形 x64 与 x86 **逐行相同**：`PL05-RAW diag=255`、`PL06-RAW boxedge=16711680 boxmid=16777215`、
`PL07-RAW fillmid=65280 circletop=255`。两条工具读数：画与问都要放在 **Timer 第一拍**（先写在
Form_Load 里，`GetPixel` 一律回 -1 = CLR_INVALID，挪出去就全对，同 dcsurf 那条"窗口活着的时候问"）；
DC 用 `GetDC(控件 hwnd)`（pbsub 那一族），`picC.hDC` 在这枚夹具上读出 0 —— 那是 #196 那一族的另一问，
本刀不动。

发码面 A/B（BASE = 账 #220 那台 `b228_new_emit`，NEW = 这一刀 `b230_new_emit`，inputs=100）：
**changed=4**（Charts 主工程 + ucProgressCircular 各 × 两档）、same=96；每份 +8/−8 行，
每一行都是同一条调用换了出口；CENSUS 兜底形状 `L"Line"` **32→0**、原生 `vb6_ControlLine(` **0→32**
（正好 8 处 × 4 份投影，别处一份都没多），而 `VB3001` 146=146、`VB6_SA_AT(` 20972=20972、
`_vb6_select_` 7346=7346 —— 一条噪声没新增。

真工程那一头：`ucProgressCircular/Proyecto1.vbp` 两档真编译仍 rc=1、27 条 error C/档，但**逐文件归因**
下来全在两处旧账 —— `Form1.c` 的 C2198 ×25 + C2084（控件数组事件臂两种形参表 + thunk 发两遍，
**另立新账 #222**）与 `ucProgressCircular.c` 的 C2065 `"Count"`（账 #219 的 HPF_BARE 那半），
而 Line 所在的 `ppProgressCircular.c` **零条诊断**、错误行里没有一条提到 `vb6_ControlLine`；
BASE 语料里那几行 thunk 与声明**逐字相同**（changed=4 的差分行里没有一条是它们）⇒ 那些红不是本刀的因。

边界与没验的一头：① `ScaleLeft/ScaleTop` 的**原点偏移**没进来（语料的 PictureBox 都是 0，
真给非零原点的工程会画偏）；② VB6 那条「Line 之后 CurrentX/CurrentY 移到终点」没进来
（调用点从没读回它，接进来要先定 CurrentX 的单位口径，与 #192 那一格同问）；
③ `With picA : .Line (…)` 那一形**没测** —— 账 #192/#150 记着带实参的 With 形至今没接；
④ 表里**刻意不给 Form 那一档**（只有 PictureBox 那处成员侧打了标记），给了就是"广告比应答复"，
接 Form 之前先找出 `Me.Line` 那条码头；⑤ `Circle/PSet/Point` 的需求面是 0，所以这一格没为它们留出口。
**两道门：#344 红在自己的名单针上，#345 全绿。** #344（run 37356616867、head `9513720f`、attempt 1）
= compile 片 `PASS=28 FAIL=1`，红的不是产品也不是本刀的判据，而是**别人那枚哨兵**
`[STATIC] control_dc ... FAIL: D2 call sites of the drawing-DC authority = 5 (want exactly 4)` ——
`check_control_dc.ps1` 的 D2 钉的是「拿绘图 DC 的调用点恰好 4 处」这份名单，Line 接进同一处权威
是正当扩容，名单没跟着改就是红。修法按本仓既有规矩（**名单扩大/缩小必须在同一次提交里改这里**）
把 4 改成 5 并把 `ControlLine` 写进名单，顺带订正该文件头 D2 那行注释（还写着 3 枚 —— 上一格加
文字量出口时也没同步，正是同一族"名单跟着代码走"会烂的地方）。
⚠ 一般式（同批进记忆）：**往一族共用的出口上加站点时，同一轮把 `scripts/check_*.ps1` 全部跑一遍**
（20 枚、几十秒），别只跑自己新写那枚 —— 这次红在别人那枚上，本地完全没读到，代价是一整轮门。
其余十片 #344 当场就绿了，两枚新夹具在 CI 上各自 PASS（`[VBP] pclinedraw ... PASS` /
`pclinedraw_x86 ... PASS`；vbp 四片 TOTAL 之和 226→228 正是这两枚）。

**门 #345 落定（run 37359094337、head `b0e4da16`、attempt 1）= 11 job 全 completed/success、逐片 FAIL=0。**
compile 片 29/29（`[STATIC] control_dc ... PASS` 与 `[STATIC] rtl_naked_names ... PASS` 同片都在）；
syntax 片 157/157（本刀换过针的 `pcline_flag_folded` 在里）；bas 两片 48 + 48（账 #220 那两枚仍在）；
vbp 四片 50(+1 SKIP)/55/50/51，唯一的 SKIP 仍是 test_vbman（COM 未注册 32-bit 视图，与 #342/#343/#344 同形）；
asm 13/14、smoke 1/1；`Build C3.exe` 那片日志正文不含用例行（历轮同形的 empty-shell）。



### B54 `.pag` 里裸写的 `Changed` 发码一直是对的、诊断却每条一响；RTL 那枚裸名全局是它的第二条权威，撞掉之后那张表才是唯一一处（账 #219 两刀，**已出：门 #346（run 37366164862、head `0026d87f`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0**）

两件症状，都是实测：

① **撞名就编不过**：任何工程里一枚 `Public Changed As Long` 与 RTL 文件作用域那枚 `int16_t Changed`（外部链接的裸名 C 全局）撞成 C2371 重定义。探针 `tests/test_rtl_naked_changed.bas` 改前 BUILD-RC=1 / no exe，改后 RC=0 / exe / `NC219-CHANGED=6 NC219-VT=3`（x64 本地真跑过，x86 交门）。同族第三枚 `g_hoCount` 一类仍留在哨兵名单里。
② **噪声**：`.pag` 里 `Changed = True` 每条配一句 VB3001，语料 24 条（PropPagLP.pag 4 条 × 两档 + ppProgressCircular.pag 4 条 × 两档）。

关键读数：**这一格的发码从来是对的**。`vb6_PropertyPage_Changed` 语料 186 处、裸名 `Changed` 在产物 C 里 0 处，`Changed = True` 早就发成 `vb6_PropertyPage_Changed = (-1);` —— 成员叫什么、能不能裸写由 `kHostPseudoRows` 那张表回答（账 #159 就收了）。所以 RTL 那枚裸名全局既没人引用、又占着 C 的全局名字空间，是纯负担 ⇒ 删（定义与 extern 各换成一条记录这组读数的注释）。语义层补的是 `visit(IdentifierExpr)` 未找到分支里那一格"裸写的文档成员"，与限定符位那一格（`memberObjCtx_`，账 #217）互斥；类型答案仍走 Variant，由发码层按表回答。

第二刀是架构那一刀，不这么写就还得再抄四遍：**语义层这一格不许自己认名字，只许问那张表**。表原先住在 `src/backend/cgen_util_com.cpp` 的匿名 namespace 里，语义层够不着 ⇒ 把它搬进 `src/common/host_pseudo.hpp`（`inline` 的表 + `hostPseudoFind` + 新增的 `hostPseudoBareEligible` 一句出口，与 `float_literal.hpp` / `int_literal.hpp` 同一家族），发码侧 `CCodeGen::hostPseudoBareName` 与语义层 `isDocumentBarePseudoMember` 都只问这一句。往后往表里加一行 `HPF_BARE`，诊断面顺带就放行了，不必两处同改。common 不得依赖 semantics，故表里的 `Symbol::toLower` 换成头文件自带的 `hostPseudoLower`（同一件事：ASCII 小写）。

读数：发码语料 A/B（BASE = `b230_new_emit` = 账 #221 收线那台，NEW = `b233_new_emit` = 这台）inputs=100/100、changed=12、same=88，12 份里**每一条差异都是删掉一行 VB3001**（逐份 `+0/−N`），产物 C 一行没动 ⇒ 这一刀只动诊断面。VB3001 146→98，名字档 `Changed` 24→0、`hDC` 24→0，其余 18 种名字一个没动（`new-names={}`）；`vb6_PropertyPage_Changed` 186=186、`int16_t Changed`(VBFlexGrid 的 UC 形参面) 10=10。`.ctl` 那 24 条是表把 `ScaleWidth`/`hDC`/`hWnd` 一并发放行顺带收掉的 —— 这就是"问表"与"抄名字"的差别。

判据两头（本地探针，`.build/b233out/`）：`pagBare.pag` 里 `Changed = True` / `If Changed Then` 零 VB3001 且发码命中 `vb6_PropertyPage_Changed` 三处，而同一份文件里打错的 `Changd` 照报；`basBare.bas`（标准模块）里裸写 `Changed` 仍报 VB3001 ⇒ 放行按文档类别，不是一片名字；`ctlBare.ctl` 裸 `hDC` 发成 `vb6_UserControl_hDC` 且零 VB3001。

夹具与哨兵：`tests/test_rtl_naked_changed.bas`（x64/x86 各一形，三针）；`tests/dochost/dhBare.ctl`、`dhBare.pag`、`dhTypo.pag` 三条 [CODEGEN-NOTE]（`dhTypo` 是负控：表里没这名字 ⇒ 必须照报，且不许凭空发 `vb6_PropertyPage_Changed`）；`check_host_pseudo_table.ps1` 的 `$tblPath`、`must`、`deny` 三处跟着表搬家，`must` 从此含"语义层必须问表"那一条；`check_rtl_naked_names.ps1` 的 N7 换成谓词 定义/声明/调用/带门 = 1/1/1/1 + 问表 1 + 成员名字面量 **0**。哨兵红过一次是当场演示的负控：把 N7 里读那三个文件的一行删掉 ⇒ 六条计数全 0、五条 FAIL。

剩下的同族（本账没做完，读数已钉住）：`Controls` 4 条全在 ppProgressCircular.**pag**（那张表 propertypage 档没有这一行；收不收要先问 VB6 里 .pag 裸写 `Controls` 是谁；源码那三行已读: `ppProgressCircular.pag:460` 是 `Set oPC = Controls.Add(App.Title & ".ucProgressCircular", "ProgCirc")`、`:484 Controls.Remove` —— 运行期往这页上动态加/删 UC，而 `:490` 紧接着用的 `SelectedControls(0)` 是表里登记过的那枚）；`Count` 2 条在 ucProgressCircular.**ctl**（表里也没这行，而 #159 的边界写明"RTL 没有对应全局的行刻意不收"⇒ 那是 RTL 侧缺口，另立账）；`ScaleWidth` 2 条在 frmDemo.**frm**:168（`If ScaleWidth > 0 Then` —— 窗体自有的量走的是另一条路，不在这张表里，与 #68/#120 那族同面）；另 26 条是内在常量一族 ⇒ 账 #218。
### B59 同一个 VB 类型写成两种拼法，`mapTypeRef` 给出两种 C 类型 —— `OLE_COLOR` 是 `int32_t`，`stdole.OLE_COLOR` 是 `void*`（账 #228，**已出：已过：门 #353（run 37392243541、head `23659871`、attempt 1）= 11 job 全 completed/success、非绿 0**）

**探针实测**（`.build/b255probe/probe.bas`，四枚 Sub 一次 `--emit-c`，3 秒）：`As OLE_COLOR` ⇒ `void vb6_BareColor(int32_t c)`；`As stdole.OLE_COLOR` ⇒ `void vb6_QualColor(void* c)`；`StdFont` / `stdole.StdFont` 同形。VB6 里这两种写法同义（类型库限定名），⇒ **限定名那一档在类型权威里掉到了兜底 `void*`**。

**为什么值得做**：这条正是 #222 第二格（门 #350 红 → #351 绿）剩下的最后一条不同源。把 emit dump 里三处签名对齐的检具（`.build/b244_agree.py`）跑五件工程 49 枚 thunk：thunk↔typedef 不符 0、thunk↔处理器不符 **1** = `VBFlexGridDemo/UserEditingForm.frm:350` `VBFlexGrid1_EditSetupWindow(BackColor As stdole.OLE_COLOR, ForeColor As stdole.OLE_COLOR)` vs `.ctl:1117` `Public Event EditSetupWindow(ByRef BackColor As OLE_COLOR, ByRef ForeColor As OLE_COLOR)` —— 发送侧交 `int32_t`、处理器收 `void*`，x64 上高 32 位是垃圾。**语料里就这一枚**（全语料扫"处理器形参类型名 vs `Public Event` 声明"：14 枚工程内 UC / 132 条事件 / 4 枚命中的处理器，其中 3 枚是控件数组合法的前置 `Index`，1 枚就是这条）。

**开工前要定的两件**：① 折的位置 = `mapTypeRef` 的 `dotPos` 那一支（现在只对**工程符号** `lookupModule(shortName)` 试裸名，内在/枚举名没试 ⇒ 掉兜底），改法 = 试裸名过 `typeSys_.resolveTypeName` 与 ivref/Class 符号，**只有查得到才折**（查不到照旧走原路，免得把 `Scripting.Dictionary` 这类真外部类型拉成原生）；② 全语料的限定名共 39 种 / 124 处，绝大多数是 `oleguids.*`、`msdatasrc.*` 这类真 COM 接口/结构（`void*` 就是对的），改完必须证明这一族**一条都不动** —— 拿 b244/b247 那两份 92 件 emit 捕获做 BASE 逐份对。

**旧说法订正（别照着做）**：本节初版猜的是"容器自己把形参写成 `LongPtr` 而事件声明是 `Long`"—— 上面那两条 grep 把它否了：两边都是 `OLE_COLOR`，只是拼法不同 ⇒ 这一格是编译器侧的类型权威问题，不是用户代码形状。

**落地读数（提交 `23659871`，判据两面钉）**：改的就一处 —— `mapTypeRef` 里从"按名字形状/名单"那几档（`Vb` 前缀 / `OLE_` 前缀 / `Enum` 结尾 / `comObjTypes` / `vb6EnumAliases` / `LongPtr` / `LongLong`）开始问 `aliasName` = 点号最后一段；上面 ivref / Class / UDT / 枚举 / ComClass 那几档**符号**分支一律不动（那里折裸名要符号真查得到，真外部类型就该留 `void*`）。
① **语料 A/B 单变量**（BASE = 上一轮那台捕的 90 份）：`inputs=90 same=88 changed=2`，4 条差异行**全部**是那一枚处理器的`void*`→`int32_t`（声明 + 定义 × 两架构），**`oleguids.*` / `msdatasrc.*` 那一族真 COM 限定名（全语料 39 种 / 124 处）一条都不动** —— 这正是"只折该折的"要的证据。② 三处同源检具（52 枚 thunk）：thunk↔typedef 不符 0、thunk↔处理器不符 **1 → 0**，#222 那一族到此收平。③ 夹具 `tests/test_alias_spellings.bas` **两面都钉**：该折的折（`OLE_COLOR` 与 `stdole.OLE_COLOR` 同型）+ 不该折的不折（真外部类型 `StdFont` 两种拼法都留 `void*`），Absent 两条专防"把所有点号都剥掉"那种过折；BASE 那台跑同一份夹具 = 少一枚 needle 且命中一条 Absent ⇒ 真红。④ 真编译矩阵两架构：9 件工程 rc=0 出 exe（含被改到的 VBFlexGridDemo），ucProgressCircular 仍是那 1 条 C2065 裸名 `Count`。
⑤ 一条行尾读数的用处：`git ls-files --eol` 显示新夹具是 `i/lf`，一开始以为是漏了规矩，读了 `.gitattributes` 才确认**这就是本仓 `core.autocrlf=true` 下 .bas 的正常形态**（只有字节敏感的 ai/028 那几份标了 `-text`）⇒ 先读仓库自带的那份说明再动手"修"，别把正常项当缺陷。

### B60 控件数组元素的 extender 属性读进 `&` 拼接 ⇒ 裸 int 进 BSTR 槽 = 两架构必崩（账 #229，**已出：已过：门 #355（run 37398206822、head `187dc3e7`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m16s**）

**症状按"属性名撞不撞内置函数名"分家，很误导**：同一枚数组元素，`& uArr(0).Left` 崩、`& uArr(0).Top` 只是绕远装箱、非数组的 `uPix.Left` 一切正常。三条读数放一起才看出来是**同一处**：类型推断里"`对象.成员` 的对象位是不是控件"这一问被抄成两份，一份只认 `控件名.属性`（P20-42 那条兜底前面），一份只认 `控件名(i).属性`（C29-1a 那份名单）。数组元素那一形从前一条前面掉下去，撞上"按成员**裸名**查模块符号" ⇒ `Left` 命中返回 String 的 VB 内置函数 ⇒ 判定为String ⇒ 拼接面不再套数值转换 ⇒ `vb6_BSTR_Concat(L"...", vb6_GetControlLeft(CtrlArr_GetAt(...)))`把 int32_t 当 BSTR 指针解引用。实测 ve_units 探针：x64 与 x86 都是 0xC0000005，且崩点就在那一行（`U-ARREXT-RAW` 整行不打印）。

**改法**：两条对象形态收成**一个出口** `CCodeGen::ctrlTypeOfMemberObject(obj, outType)`（单枚 / 元素同一条规则，登记表 `knownFormControls_` 只这一处读），P20-42 那一处改问它，属性类型仍出自那张表 `controlPropType`。（当时留在明处的那格"这张表之外还有一份自带名单"已在**下一格 B62（账 #231）**收掉。）

**判据（两头 + 负控）**：夹具 ve_units 加一头，`U-ARREXT-RAW l=3600 t=1320 w=1200 h=1140` 那一行**就是崩溃现场本身**（四枚读法全在 `&` 拼接里），另一行 `U-ARREXT=` 问四个变量等于设计期几何（3600/1320/1200/1140）—— 只钉 RAW 那一行会放过"值对不上"的修法，只钉判据行会放过崩溃。负控 = 用改前那台编译器跑**同一份**夹具 ⇒ 两架构 rc=0xC0000005 且 U-ARREXT 整行不出现；修后两架构 rc=0、两条读数逐字相同。

**护栏**：语料 A/B（BASE = 上一轮 #228 之后那台捕的 90 份）⇒ `inputs=90 same=88 changed=2`，改到的只有 ve_units（两架构各一份），差异行**全是本轮夹具自身新增的 15 行**（纯插入，产品发码零改动）—— 这条读数的含义是：**全语料没有别的数组元素 extender 读法**，所以这一刀只能靠新夹具守，指望存量用例发现它是做梦。 真编译矩阵两架构：9 件工程 rc=0 出 exe（Charts 2020 全家 / czUI / VBFlexGridDemo / ve_units），ucProgressCircular 仍 1 条 C2065 裸名 Count；LabelPlus 那一条 MISSING 是它压根没有 .vbp（无效输入，不是红）。合并 github/dev 的 rev39（LoadRes* 一族）之后重跑同一份夹具：两档 15 条读数全 True。

### B61 数组元素的 extender 属性**写后读回不是请求值**（缇→像素→缇 取整损失），VB6 存的是缇（账 #230，**已量到，未开工**）

同一轮探针量出来的：`uArr(0).Left = 5000` 之后读回 **4995**（5000 缇 = 333.33 像素，控件位置只能落整数像素，读回时再乘回去就丢 5 缇）。VB6 的 `Left` 是**属性值**而不是窗口位置的投影，写什么读什么。同族的既有账是 #206（`Panels(i).Width` 交回请求值还是排版后的宽，单位口径待量）—— 两格合起来是一个问题：**控件几何属性到底以哪一侧为准**（存 VB 侧的值 vs 问窗口）。动它之前要先定口径（VB6 语义 = 存 VB 侧），并且别忘了 `Move` 与容器排版会改窗口而不改 VB 侧的值。


### B62 控件属性的**类型**有两处权威 —— `controlPropType` 那张表 + `inferExprType` 里的三份自带名单（账 #231，**已出：已过：门 #356（run 37401361458、head `442da251`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m23s**）

**为什么这一格值得单开**：同一个问句（"Shape1.FillStyle 是什么类型"）以前有**两处**各自作答 —— 那张表（`cgen_util_ctrl.cpp`）与 C29-1a / C29-1b / C29-9 在 `cgen_util_type.cpp` 里手抄的三份名单（`kNumericFc` 14 条 / `kStringFc3` 5 条 / `kStrFcCd` 7 条 + `kNumFcCd` 7 条）。两份名单与表重合的只有 left/top/width/height 四条，**重合是靠"恰好一样"才没出事**；账 #229 崩的那条就死在这道缝上（同一问被抄成两份，一份只认 `控件名.属性`、一份只认 `控件名(i).属性`）。所以这一刀不改任何答案，只把两处并成一处。

**改法**：那 28 条名字逐条搬进 `controlPropType`（通用段补 10 条数值名 + 文件系统三控件一节 + CommonDialog 一节），`inferExprType` 里三份名单连它的循环一起删，只留**一次问话** —— 位置仍在 `MemberAccessExpr` 那条 case 的**最前面**（原 C29-1a 的位置），所以答案的**先后顺序**也没动；C29-1a 那道 `!= FrmControlType::Unknown` 的闸一并搬进表里（自定义 OCX/UC 的属性面归类型库，不让这张表按名字形状抢答）。

**这一刀的护栏比往常硬**：refactor 的失败模式不是崩，是"某条名字在整个语料里根本没人这样写" —— 那种漏在 emit A/B 上是**哑的**（上一格 #229 的读数就是"changed 全是夹具自身新增行"）。所以两头一起钉：
- 语料 A/B：BASE = 改前那台（`C3_base231.exe` 冷存；先用它复捕一份，证明与上一轮那 90 份**逐字节相同**才承认它是 BASE）⇒ 改后 `inputs=90 changed=0 same=90`（覆盖面不止这 45 份 .vbp：`tests/` 下 **250 份单文件 .bas 用例**也各用两台 emit 一遍、诊断文本一并入读 ⇒ 同样 `changed=0`，脚本 `.build/b286_basab.py`；合计 **340 份捕获一字不差**），**产品发码零改动**（这一刀的正确答案就是 0，不是"逐行归因后 0"）。
- 新哨兵 `scripts/check_ctrl_prop_type_authority.ps1`：A1 旧名单标识符与它的循环变量回潮 = 0；A2/A3 `controlPropType` 与 `ctrlTypeOfMemberObject` 各"定义 + 声明 + **恰好一个**调用者" = 3 次提及；A4 **28 条名字逐条**必须在表里答到（`p == "<名>"`）；A5 那道 Unknown 闸必须还在且只有 1 处。三条负控（假插一份 `kNumericFc` / 把 `fillstyle` 改名 / 多开一个调用点）各让对应规则红，跑完按 md5 还原源文件。已进回归 `[STATIC] ctrl_prop_type_authority`（第 22 道）。
- 全 22 份 `check_*.ps1` 逐份绿；真编译定点 4 件（ve_units 两架构 / ucTreeMaps x64 / VBFlexGridDemo x64 —— 最后一件正是 A5 那道闸的对象）全部 rc=0 出 exe、诊断 0 条。

**工具事实（踩过才记）**：python 的 **bytes** 字面量里写 "src" + 反斜杠 + "rtl" 时，那枚反斜杠-r 会被折成一枚真 CR 塞进文件 —— 于是那一行被劈成两半，PowerShell 报 ParserError 而**字节数看着完全正常**。此后校验 CRLF 文件必须同时数 `lone_lf` 与 `lone_cr`（只数 LF 会放过这一类）。


### B63 Form 的绘图状态属性：写侧从没接线（C2106，压根编不出来）+ 读回恒 +1 + 笔位在一个属性名上挂两种编码（账 #233，**已出：已过：门 #357（run 37405355138、head `05f04f31`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m05s。订正一句读数方法：那台 watcher 回读 jobs 时被本机代理顶了一次，只写出 `jobs=0 non-success=0` 就收线 —— `conclusion=success` 配 0 条 job 不是"全绿"，是**没拿到读数**；补一次按 run id 回 API 复核才数到 11 条 （`.build/b309_verify357.py`）**）

**这一族的缺陷形状是"只接了一半"**。rev38（`9a420157`）新做了 Form/Printer 的绘图方法家族，读侧四条属性（`Me.CurrentX / CurrentY / ForeColor / DrawWidth`）硬编码在 `cgen_expr_member_form_builtin.inc` 的一份**侧表**里，注释写着"写侧走赋值语句自己的发射路径，在这里加写侧是不可达的死代码"—— 实测**恰好相反**：赋值发 C 时先问 `getControlPropWriteFn(Form, 名)`，没登记就退化成"把读函数当左值" ⇒ `Me.DrawWidth = 3` 发成 `vb6_Form_DrawGetWidth(vb6_hwnd_X) = 3;` ⇒ **C2106，两架构一条产物都出不来**。RTL 那四个 setter 其实**早就写好也声明好了**（`vb6forms_draw.h:51-54`），只差 cgen 那张表没登记 —— "备齐了入口却没人接"。

**同一刀量到的另外两格**：① `vb6_DrawSetI` 存 `v+1` 而 `vb6_DrawGetI` 不减 ⇒ 写 3 读回 4，且 `PSet/Line` 的 Step 那一路每画一次多带 1（累积漂移）；② `VB6_CurrentX/Y` 这**一个窗口属性名**上挂着两套编码 —— 绘图家族 int32(+1)、`vb6_Form_Print`（`vb6forms.c:1541`）float 位图案 ⇒ `Print` 把笔位推到 46 之后 `PSet` 读到 779103232，反过来 `PSet` 之后 `Print` 永远打在 y=0；③ 顺带一格：`vb6_DrawScaleMode` 按"带 +1 的编码"读 `VB6_ScaleMode`，而它唯一的写者 `vb6_SetScaleMode` 存裸值（第三种编码撞同一个名）。

**改法（三处都朝"一处权威 + 读写成对"收）**：① cgen 把 currentx/currenty/drawwidth **读写成对**登记进 `getControlPropReadFn` / `getControlPropWriteFn` 的 Form 档（与 #197 那条 `scalemode` 同一格口径），侧表整块删掉；② `vb6_DrawSetI` 存裸值，与 `vb6_DrawGetI` 同一套编码（"0 与没设过不可分辨"这一问改由各属性自己的缺省档兜：`DrawWidth` setter 先钳 >=1、`VB6_BackColor` 的写者本来就存裸值）；③ 笔位归一 —— 绘图家族内部 6 处读 / 8 处写改问那份 **float 出口**（`vb6_GetCurrentX/SetCurrentX`），并撤掉本文件那四个 int32 导出访问器；`vb6_DrawScaleMode` 改问 `vb6_WindowScaleModeSelf`。`forecolor` **刻意没动**：它早就由通用行答给 `vb6_GetControlForeColor`（侧表那一条从没命中过），与绘图家族自带的 `VB6_DrawForeColor` 是两份存储 —— 那一问另开账（见下）。

**读数（改前 / 改后各一头）**：新夹具 `tests/fdraw/FDemo.vbp`（Form 绘图状态真跑，Timer 第一拍）—— 改前用上一台编译器跑**同一份夹具**：两架构 build-rc=1、4 条 C2106、产物出不来；改后两架构 rc=0、10 条读数全对（`FD04-RAW xy=300,130 dw=3`、`FD07-PIXEL=True` 画到就问得到、`FD08-neg=True raw=15790320` 旁边那点还是背景色、`FD09-AFTERPRINT=0,46` 即 Print 推进的数**绘图这一路读得到**）。FD09 那行的 y 与字号/DPI 有关 ⇒ 只打印不当判据，判据换成落在 30..3000 的那枚布尔（FD10）。

**护栏**：语料 A/B（BASE = `C3_base231.exe`）⇒ 90 份 vbp + 250 份单文件 .bas = **340 份捕获 `changed=0`** —— 这条读数同时说明"全语料没有一处这样写"，也就是这一族能带着 C2106  shipped 的原因：**零覆盖**。所以补了哨兵 `scripts/check_form_draw_state.ps1`（S1 笔位窗口属性只许 `vb6forms_widget_prop.c` 一处 / S2 cgen 侧表那四个名回潮 = 0 / S3 三个名在读写两张表的 Form 档**成对** / S4 `vb6_DrawSetI` 那条 SetPropW 不许再 +1），四条各用一处假改动证明会红（S3 那条假改动改的正是本次的伤：write=0），跑完按 md5 还原。23 份 `check_*.ps1` 逐份绿；邻域真跑四枚夹具（fdrawstate / pclinedraw / dcsurf / ve_units）期望串零缺失；真编译矩阵 4 件全出 exe。

**留给下一格（账 #224 剩下的）**：`vb6_DrawAcquire` 与 `vb6_ControlDrawDC` 是"拿 DC"这个决定的两份实现，而 `check_control_dc.ps1` 扫不到前者；`Print` 按像素推进、绘图按用户单位收 —— 单位口径那一问本刀刻意没碰；以及 `Me.ForeColor`（存 `VB6_ForeColor`）与画点用的 `VB6_DrawForeColor` 是两份色彩存储。


### B64 「这枚窗口的绘图 DC 从哪儿来」被实现了两遍 —— 绘图方法家族改问唯一权威（账 #234，**已出：已过：门 #358（run 37409833257、head `2b3baf82`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m44s**）

**这一格没有 bug 症状**，正因此才值得单开：rev38 在 `vb6forms_draw.c` 里把"派发期先问窗口属性 `VB6_PaintDC`、否则 `GetDC`、fromPaint 那张不许 Release"这整条口径**又写了一份**，与账 #185/#196 收口在 `vb6forms_ctrl.c` 的那份**逐条同口径**（它自己的提交说明就写着"完全同一口径"）。两份都活着的时候行为一模一样，谁都看不出问题；**一改其中一份，另一份就静默分家** —— 而 #185 那一族（同名窗口属性/同一张 DC 被两处各拿一遍）的教训正是这种分家的现场。census 面上 `check_control_dc.ps1` 的名单只盯 `vb6forms_ctrl.c` ⇒ 那个洞在这刀一起补，不留"以后再收"。

**改法（三处，零行为改动）**：① `vb6_ControlDrawDC` 去掉 `static`；② 在 `vb6forms_internal.h` 里声明它 —— 那个文件的既有职责就是"文件级 static 不跨编译单元可见 ⇒ 跨族共享符号集中声明"，声明旁边写清 `*pFromPaint=TRUE` 那张不许 `ReleaseDC`；③ `vb6_DrawAcquire` 只留"NULL 先挡（Printer 那一路另有出口）+ 把 `&out.fromPaint` 直接交给权威"，两分支与旧代码逐条等价（NULL→`{NULL,FALSE}`；非 NULL→派发期那张或 `GetDC`）。文件头那段"HDC 来源（关键）"的口径说明改成指向权威，不再复述实现。

**哨兵跟着长（这一步才是本格的主体）**：`check_control_dc.ps1` D1 的"定义"匹配**排除以 `;` 结尾的声明行** —— 不排除，本刀新增的那条头文件声明当场被数成"第二处定义"= 冤案红（这条是改 census 时最容易踩的：名单一扩，先把"声明/定义"的区分补上）；D2 钉死的调用点数 **5→6**，多的那条就是 draw.c；新增 **D13** = `vb6forms_draw.c` 里不许再出现 `GetDC(` 或 `GetPropW(...L"VB6_PaintDC")`（注释行跳过）。三条各用一处假改动证明会红、跑完按 md5 还原。

**读数**：23 份 `check_*.ps1` 全绿；邻域真跑四枚夹具（`fdrawstate` 含 FD07 像素证人 / `pclinedraw` / `dcsurf` / `ve_units`）28 条 needle **零缺失**；语料 emit A/B 90 份 `changed=0`（RTL 改动天然不进 `--emit-c`，这条只证前端没被碰到，不当行为护栏）；真编译矩阵 4 件全出 exe、诊断 0 条。踩到的一次真红：第一次行切片把 `vb6_DrawAcquire` 的收尾 `}` 一起替换掉了 ⇒ C2143/C2065 一片、四枚夹具 build-rc=1 —— 又是"哨兵只扫源码抓不到、必须真编译"那一族，插入/替换的边界必须是**整条语句**。


### B65 画笔色有**两份存储**从不汇合 —— `Me.ForeColor = vbRed` 之后不带颜色的 `PSet` 画出来是黑（账 #235，**已出：已过：门 #360（run 37415128806、head `ea8dc7c9`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 8m31s**）

**读数是这格的起点**（`.build/b308probe`，两架构一致）：`Me.ForeColor = vbRed` 之后，**不带颜色参数的** `Me.PSet` 落笔是 **0（黑）**，而 `Me.PSet (x,y), vbRed` 是 255。成因不是"写侧没接"（那是 #233），而是**同一件事有两份存储**：`Me.ForeColor` 走通用行 → `vb6_SetControlForeColor`（窗口属性 `VB6_ForeColor`），而绘图家族取色走私有的 `vb6_DrawForeColor`（另一枚属性 `VB6_DrawForeColor`，自带一套 +1/-1）。**控件那侧一直只有一份** —— `vb6forms_ctrl.c` 的 Line(:800) / Print(:854) / 子控件回显(:642) 全问 `vb6_GetControlForeColor` ⇒ 这一刀是把 Form 绘图**对齐到已经正确的那一份**，不是新立口径。

**改法只有一处**：`vb6_DrawForeColor` 改成问那份唯一出口；私有的 setter 与两个导出（`vb6_Form_DrawGetForeColor` / `SetForeColor`）撤掉 —— 这两个名字在 `src/`（cgen 侧）与 `tests/` 里**零引用**，因为 `forecolor` 早在 #233 那轮就确认由通用行答复、侧表那条从没命中过，所以删的是"备好了却没人接的入口"。Printer 那族**刻意不动**（没有 HWND，色值存静态量 `g_prnDrawFg`，与窗体那份不冲突）。

**为什么这里不需要 Fix 187 那道哨兵**：`VB6_ForeColor` 未设时读回 0，而 0 就是 vbBlack = VB6 的默认画笔色，"没设过"与"设成黑"**默认档重合**，所以合并不会把黑色误判成未设；Fix 187 那条坑属于 **BackColor**（它默认是 `COLOR_BTNFACE` 而不是黑，才必须另立 `VB6_BackColorSet` 哨兵）。这个区别写进了注释，免得下一位照 #187 再补一枚哨兵。

**判据（新加一头，两头钉）**：`tests/fdraw` 的 **FD11** = `Me.ForeColor = vbBlue` 之后不带颜色的 `PSet (100,120)` 那点必须 = 16711680，**同时**先前带颜色画的 (40,60) 那点仍是 255 —— 只钉前者会放过"整屏刷成蓝"那种假绿。**假针证红是真做的**：把 `vb6_DrawForeColor` 逐字改回改前那份私有存储、重编 C3.exe、跑同一份夹具 ⇒ `FD11-forecolor=False`、`FD12-RAW pen=0`（两架构），还原重编 ⇒ `True / pen=16711680`。（`b308` 那次读数与这条互相印证：同一件事，一台改前、一台改后。）

**护栏**：哨兵 `check_form_draw_state.ps1` 加 **S5** = 属性名 `VB6_DrawForeColor` 在 `src/rtl` 里不许再出现在任何 `GetPropW/SetPropW` 行（第二份存储回潮 = 红），假针负控已证会红；23 份 check 全绿；邻域四枚真跑夹具（fdrawstate / pclinedraw / dcsurf / ve_units）零缺失；emit A/B 90 份 `changed=0`（只动 RTL ⇒ 这条仅证前端没被碰，不当行为护栏）；真编译矩阵 4 件出 exe。**#224 现在只剩两格**：Print 推进笔位按像素而绘图按用户单位（②-单位），以及"拿 DC 的 census 已并但 Form 的 `Print` 仍自己 `GetDC`+无条件 `ReleaseDC`"（记在 #224③ 的订正里）。









### B66 门 #359 那条红不是产品坏了 —— 夹具的收线坐在自己那道闸后面（账 #236，**已出：已过：门 #360（run 37415128806、head `ea8dc7c9`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 8m31s**）

症状只有半条：`Tests (vbp #4)` 退 1，别的十片全绿。拿不到 job 日志（PAT 没有 actions:read），
但**工件（artifact）拿得到** —— `actions/artifacts/<id>/zip` 会先 302 到签名 URL，跟过去时**必须把
Authorization 头丢掉**，否则对象存储回 403（`.build/b335_dl.py` 那份是活模板）。工件里 
`ResAlpha.out` 27 字节、两条 needle 都在，**却没有配对的 `ResAlpha.err`** —— 这就是定位本身：
`Invoke-TestExe` 正常退出那条路 `WriteAllText` 两份都写（哪怕 stderr 是空的），只有超时那条路
「非空才写」，所以**整片唯一缺 .err 的那枚就是被 60s 杀掉的那枚**。本地照做：45s 不退出（rc=124）。

成因在夹具，不在发码：`tChk_Timer` 第一行 `If done Then Exit Sub`，而 `tick = tick + 1` 在它自己
那趟的末尾 —— 第一拍把 done 置上之后每一拍都提前返回，`tick` 永远停在 1，
`If tick >= 30 Then Unload Me` 再也够不着。两条 needle 早就打完了，所以红得很像"CI 抖动"。

修法：计数器先走，每条路径都有界（Picture 没挂上时也在 30 拍后关窗 —— 那时缺 needle 会正常报红，
而不是把整片拖到超时）。第 24 道哨兵 `scripts/check_fixture_timer_close.ps1` 把这条口径钉住：
凡「靠计数阈值收线」的 `*_Timer` 处理器，自增必须出现在**第一条提前返回之前**；
F1 覆盖面下限（实数 29，下限 20）、F3 适用面非空，两条都是防"哨兵自己没电"。
负控：把 HEAD 那份旧夹具放回原路跑 => F2 红并点名行号；换回修好的 => 绿。

一条工具事实（本轮踩过）：`git show HEAD:<path>` 交回的是**索引里的 blob（LF）**，而工作树因 
core.autocrlf 是 CRLF —— 按 `
` 切那份字节的脚本会把整个文件看成一行，于是"改好了"其实没改。
切行之前先看分隔符，别假定。

### B67 `Print` 住在窗体那套代码里，DC / 字体 / 色彩 / 单位四件全是自己答的 —— 缇档 Print 一行推进 16 而 TextHeight 答 240（账 #237，**已出：已过：门 #361（run 37421550422、head `d86b478c`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 8m49s**）

症状两半，都是实测（探针 `.build/b347probe`，x64 与 x86 一模一样）：
`Me.ScaleMode = vbTwips` 下 `Me.CurrentY = 0 : Print "AB"` 之后 `CurrentY = 16`，而同一枚窗体自己答
`Me.TextHeight("AB") = 240` —— 笔位那份存储从 #233 起就是**用户单位**，Print 却把一个**像素行高**存回去；
另一半更糟：把笔位放到 `ScaleHeight / 2`（缇档 = 客户区正中）再 Print，墨落在**第 3 行**（那是上一行
Print 的残留），因为落点坐标是用户单位被直接递给 `TextOutW` 当像素用了 —— 真实落点在客户区外，看不见。

`vb6_Form_Print` 以前住在 `vb6forms.c`，四件都自带一份答案：自己 `GetDC` + 无条件 `ReleaseDC`
（派发期 `BeginPaint` 挂在窗口上的那张 `VB6_PaintDC` 它从不问 ⇒ `_Paint` 里 Print 落不进那一轮）、
不选字体（拿 DC 默认字体而不是这枚窗体的 `Font`，与 #200 同一味）、不 `SetTextColor`、单位自己算。
本文件里的 `vb6_DrawUserToPx` 也是第二套口径：写死 `v * dpi / 1440`，**只认缇**，
`Point/Inch/Centimeter/Millimeter` 全按缇算（差 20 倍），纵向还用横向的 dpi —— 正是 Fix 184
那句"所有换算共用那一对真实 DPI 出口"没覆盖到的角落。

修法 = 把 Print 搬进 `vb6forms_draw.c` 并让它问四件已有的权威：DC `vb6_DrawAcquire`（#234）、
字体 `vb6_ControlFont`（#200）、色 `vb6_DrawForeColor`（#235）、单位改问 `vb6_ScaleUserToPx` /
`vb6_ScalePxToUser`（`vb6forms.c` 那一族唯一权威）。推进量刻意取**刚写那串字的
`GetTextExtentPoint32W().cy`** —— 那正是 `vb6_ControlTextHeight` 量的同一个量，于是"Print 之后
CurrentY 的增量 == Me.TextHeight(同一串)"是两条路真汇合而不是各写一遍。家族里其余 11 个换算调用点
一律显式写出纵/横那一档。

判据（`tests/fdraw`，两架构真跑，针面从 12 条扩到 18 条）：FD13/FD14 把推进与 `TextHeight` 比相等
（缇、点两档），FD16 是"72 点与 1 英寸落在同一行"的跨单位巧合 —— 蓝长行留在右边、红短行盖在左边，
两个 x 窗口各读一个颜色，所以"第二行根本没画"赢不了。全都不钉绝对数 ⇒ DPI 变了不假红。
行为负控：把落点与推进改回改前的形状 ⇒ FD13/FD14/FD16 全 False（`curTw=16` 对 `th=240`、
`rowB=75` 对 `rowR=-1`），还原后全 True。哨兵 `check_form_draw_state.ps1` 加 S6/S7：Print 恰定义一次
且必须住在家族文件、体内七件权威一个都不许少；换算码头必须交回权威、每个调用点必须带 0/1 那一档、
本文件不许再出现 `dpi` / `LOGPIXELSX` / `1440`（四条假针逐条证红）。
护栏：24 道哨兵 red=0、90 份 emit 捕获 changed=0（纯 RTL + 夹具）、邻域五枚真跑全绿、矩阵 4 行干净。

### B68 与 Variant 比较的那个标量操作数没装箱 —— `vb6_VarCmpEq(&variant, &double)`（账 #238，**已出：已过：门 #365（run 37436391539、head `9b9c1551`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m40s**）

写 FD13 时撞见的：`(Me.CurrentY = thTw)` 在两个数**实测相等**（`FD15-RAW` 打出来 `th=240 curTw=240`）
的情况下答 False，反过来写 `(thTw = Me.CurrentY)` 也 False，而算术那条路 `(Abs(Me.CurrentY - thTw) < 0.001)` True。
发码直接给出现场：`eqA = (vb6_VarCmpEq(&_vcmp_8, &thTw));` —— 第一个实参是装箱好的 `vb6_VARIANT`，
第二个是把**裸 `double` 变量的地址**当 `vb6_VARIANT*` 递了进去（RTL 原型 `int32_t vb6_VarCmpEq(vb6_VARIANT*, vb6_VARIANT*)`），
于是读到的 vt/值是那块内存的巧合内容。症状不是崩而是**两个相等的数答 False**（静默给错答案那一族）。
左边是 Variant（`Me.CurrentY` 的类型 oracle 交 Variant 档）、右边是本地 Double/Single 时都会走到这条路；
`VarType()` 读数 `4,5`（Single / Double）也对得上。开工前先量清到底有多少调用点把非 VARIANT 的地址递进
`vb6_VarCmp*` 这一族（cgen 的比较发码处），修法应是**两侧都装箱**而不是换 RTL 原型。
本刀的夹具暂时用算术形式表达判据（`tests/fdraw` 里那段注释点名了这一格，修好就换回直接相等）。

**暴露面（按函数作用域逐处回溯声明扫 90 份捕获，工具 `.build/b367_varcmp_scope.py`）**：全语料 158 处 `vb6_VarCmp*(&A, &B)` 里 126 处两侧都是 `_vcmp_N` 临时、32 处两侧都是声明为 `vb6_VARIANT` 的局部，**0 处把标量局部的地址当 VARIANT* 递进去** —— 即本刀那两处是目前唯一知道的形状，且它们在新写的夹具里。也就是说这是一格**潜伏缺陷**：要求 一侧是 Variant 类型的表达式、另一侧是 「`Dim … As Double/Single/Long`」这种标量局部，而它一旦出现就是静默错答案（不报 C 类型错、不崩）。修之前先把判据 钉进发码针（含"裸标量地址不许进 VarCmp"那条哨兵）。

**已收（这一族的形状 = "取址那一问按名字形状答，没按类型答"）**：四个取址点（`variantAddr158n`
的裸名字那一支、VarCmpLong 的左右两条腿、兜底那一对 `vb6_VarCmp*(&A,&B)`）现在全部问同一处
判据 `CCodeGen::cmpOperandMayTakeAddr(名字, AST)` —— 定义在 `cgen_util_type.cpp`、声明在
`cgen_helpers.inc`，回答只来自 `isDefinitelyVariantExpr`（它读声明那几张表 + 符号表，正是这几张表
把 `Dim d As Double` 钉成 `double` 的）。不许取址就走装箱，装箱沿用那条 `_Generic
vb6_VariantFromValue` —— 对**已经是** `vb6_VARIANT` 的表达式它命中 `vb6_VariantIdentity` 恒等直传，
所以"权威认不出的真 Variant"最坏只多一枚临时，不会改语义（这条是敢翻转默认档的依据）。
读数：新夹具 `tests/test_varcmp_scalar.bas` 16 条判据两架构全 True，BASE 那台同一份夹具 7 条 False
（VC01..VC05 / VC10 / VC15 —— 相等答 False、`Not(...)` 那一面也答错，正是本账的形状）；
`tests/fdraw` 的 FD13/FD14 由算术形式**换回直接相等**（本刀的用后归还，改前那台上它 False）。
发码面 `Test-CodegenNote "varcmp_scalar_boxed"` 两头钉：必须出现 `vb6_VariantFromValue(d)` /
`(&gV` 那类装箱，必须不出现 `vb6_VarCmpEq(&v, &d)` 等五条旧形状。
哨兵 `scripts/check_variant_cmp_boxing.ps1`（第 25 道 [STATIC]）三条规则各用一处假改动证红：
判据 bodies 不再问权威 → V1 红（同一次假改动同时让夹具回红 = 行为负控）；撤掉 VarCmpLong
那一腿的问话 → V2 + V3 双红；复制一枚声明 → V1 decl 红。
语料 A/B `inputs=90 changed=2`，逐行归因 = 装箱语句 + 被改写的比较 + `_vcmp_N` 编号平移，**未归因 0**；
唯一一处真形状变化是 `ucProgressCircular` 的裸名 `Count`（那枚名字本来就 C2065 ⇒ 该工程今天编不出 exe，
BASE/NEW 两台的 build 结果逐条相同），所以这一格在**能编过的语料里是零暴露**。

**订正上一轮那句"0 处"**：它只统计了"在同一个函数体里找得到声明"的名字，因此漏掉了**未声明的裸名**
那一形（`&Count`）。量暴露面时要把"没有声明的名字"单独列一类，否则会把 1 读成 0。

### B69 `PictureBox` 的 Print 还留着第二份笔位，而且那份是像素（账 #239，**已出：已过：门 #364（run 37429327451、head `e5c66e4d`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 12m40s**）

`vb6_ControlPrint`（`vb6forms_ctrl.c`）把光标存在窗口属性 `VB6_PrintX` / `VB6_PrintY` 里、按**像素**推进，
而控件的 `CurrentX` / `CurrentY` 读写的是 #233 那份 float 笔位（`cgen_util_ctrl.cpp` 两档都登记到
`vb6_GetCurrentX/Y` / `vb6_SetCurrentX/Y`）—— 两份存储从不汇合：`pic.Print "AB"` 之后 `pic.CurrentY` 一动不动，
而 `pic.TextHeight` 答的是**用户单位**（#177 那条换算）。这正是 #224 那一族在控件侧的镜像，
也是 #237 在 Form 侧刚拆掉的那个形状。语料里只有一处控件侧 Print（`VBFlexGridDemo/MainForm.frm`），
所以存量用例照旧全哑 —— 判据得新写，别指望 GA 红。

**已收（同一形状别再抄一份）**：`vb6_ControlPrint` / `vb6_ControlCls` 现在是两条**码头**，转调
`vb6_Form_Print` / `vb6_Form_Cls` —— 那一份经过 #233(笔位) / #234(DC) / #235(色) / #237(单位) 之后
五件都问的已是唯一权威，所以控件侧不必再写第二遍。实测三对读数（同一枚 `PictureBox`，两架构一致）：
Print 之后笔位推进 `0 → 13`（= 同一枚控件自己答的 `TextHeight`）、笔位放到 60 时墨的落点 `第 2 行 → 笔位那一行`、
`Cls` 之后 `CurrentY` `400 → 0`；缇档推进 `0 → 195`。**顺带第二格（它同时是本刀的防回归）**：
家族的 `vb6_Form_Cls` 以前自己按 `VB6_BackColor` 判空取背景色 —— 黑色存进窗口属性就是 NULL，
按值判空 = 读成"没设过" ⇒ 回落按钮面；改成问带 Fix 187 哨兵那份唯一出口 `vb6_GetControlBackColor`。
不这么改，控件那户 Cls 会因转调而从实测 `0` 变成 `15790320`（改前控件答 0、家族答按钮面，两份答案一旦合一就必须选对的那份）。
量到的另一格只读不响：`vb6forms_picture_prop.c` 里还有一处 `GetPropW(VB6_BackColor)`，但它在
`C3_FORMS_TRACE` 那枚 env 闸里的 fprintf 内（同时打印 set 旗标），是诊断文本不是第二份行为答案。

**顺带量到的新事实（属 #232 那一族）**：`Me.Cls` 在 VB→C 那一路压根没出口 —— 发成
`vb6_ComCall(vb6_hwnd_<Form>, L"Cls", NULL, 0)`（emit 物证 `.build/b373_emit_base.c:327`），运行期 no-op。
所以家族那份 `vb6_Form_Cls` 今天只被控件那户到达；窗体自己的 Cls 何时能跑，等 #232 那条前端出口。

### B70 窗体绘图语句的"裸形"与 "Me." 形各缺一条路（账 #232，**② 已出：门 #368（run 37454147931、head `9e5c9f91`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m46s；①③ 未开工**）

探针 `.build/b433_shapes.txt`（同一枚 `.frm` 每次只放一条语句，`--emit-c` 看发码；工具 `.build/b431_232probe.py` 那套形状表）。**七形七样**（窗体上）：

| 写的形状 | 今天发出来 | 症状 |
| --- | --- | --- |
| `PSet (100, 100)` | `PSet(100, 100);` | **未声明的 C 函数** ⇒ cl C2065，编不过 |
| `Circle (200, 200), 50` | `Circle(200, 200, vb6_VariantEmpty(), 50);` | 同上 |
| `Cls` | `Cls();` + 一条 VB3001 警告"未声明的标识符" | 同上（警告 + C2065 两层都不响） |
| `Me.Cls` | `vb6_ComCall(hwnd, L"Cls", NULL, 0)` | 编得过、跑得起、**什么都不做**（CLINE 那族静默空转） |
| `Me.Print "AB"` | `vb6_ComCallObject(vb6_ComGetObjectProp(hwnd, L"Print"), L"Item", ...)` | 同上，一笔不画 |
| `Line (0, 0)-(10, 10)` | —— | **VB2001 expected ')' (got ,)** —— parser 就不认 |
| `Line -(20, 20)` | —— | **VB2001 expected 'Input' after 'Line'** |

已经通的三形钉住当基准：`Print "AB"` → `vb6_Form_Print(...)`、`Me.PSet (x, y)` → `vb6_Form_PSet(...)`、`Me.Line (0,0)-(10,10)` → `vb6_Form_Line(...)`。
所以这一格不是"绘图面没做"，是**同一条方法的两条写法各缺一段路由**：
① 裸形（`PSet` / `Circle` / `Cls`）今天当"用户模块过程/隐式声明"处理 ⇒ 需要走 `Print` 那条"窗体上下文语句"的改写；② `Me.` 形（`Cls` / `Print`）没进控件方法改道那张表 ⇒ 落 COM 兜底；③ 裸 `Line` 的两形在 parser 就没出口（`-` 续画形连 `Line` 语句自己都不认）。
**开工顺序按"缺口小→大"**：先 ②（登记两行 + 一条发码针，和 #143/#149 同一套手法），再 ①（`Print` 那条改写已有先例），最后 ③（要动 parser 的语句形状，且 #220/#221 那套 B/BF 旗标折叠要一起接）。判据按 #221 那条口径：**画完再问像素**，发码面另钉一条"三形都不许出现 `vb6_ComCall`"。语料里 `Charts 2020` / `czUI` 有没有用裸形，开工前先用同一枚形状表扫一遍语料。

**开工前的两条读码结论（省一轮定位）**：② 那一格**不是没登记** —— `cgen_expr_call_callee_withm.inc:697` 里
`cls` 早已有路（`vb6_Form_Cls((void*)obj)`），但 `Me.Cls` 实测仍落 `vb6_ComCall(..., L"Cls", NULL, 0)`
⇒ 零实参的 `对象.方法` 在**语句位置**被当成"属性读并丢弃"，压根没进这条方法分派（先查 parser/语义那一步
的"是不是 CallStmt"，别在 cgen 的表里加行 —— 加了也不会到）。`Me.Print` 同理：裸 `Print` 走的是
`cgen_file_io.cpp:83-93` 那条"无文件号的 Print 语句 = 窗体 Print"的改写，而 `Me.Print` 是成员调用形 ⇒
表里没有 `print` 这一档 ⇒ 落 COM 兜底。所以 ② 的最小形状 = 让这两种写法进到**已有**的那条分派，
而不是新增第二份答案。

**② 已出（门 #368（run 37454147931、head `9e5c9f91`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m46s，提交 `1454f5b8` + 修正 `9e5c9f91`）—— 先订正上一轮那条读码结论**：窗体自己那枚接收者**本来就认得**，
`cgen_form_ctrl_registry.inc:10` 把窗体名也登记进 `knownFormControls_`（类型 Form），
所以缺的不是"入口"而是**名字表少两档**。证据 = 同一条 `Me.` 上 `Me.TextHeight("AB")` 一直是通的
（`vb6_ControlTextHeight((void*)vb6_hwnd_<窗体>, …)`，走的正是 `formCtrlSlot` + 表那两条码头），
只有 `print` / `cls` 因为不在表里、而三处画布码头又各自硬编码 PictureBox 才落进兜底。

**语料先扫（按上面定的规矩）**：`Me.Cls` 0 处、`Me.Print` 0 处、裸 `Cls` / `PSet` / `Circle` 0 处、
`PictureN.Cls` 5 处（本来就通）⇒ 这一格是**潜伏缺陷**（编得过、跑得起、一笔不画那一族，本线第四次），
不是当前的编译阻塞；"编不过"那两格是 ① 与 ③。

**十三形实测**（探针 `.build/b434probe/Me232Form.frm`：同一份 .frm 把两接收者 × 两成员 × 括号/实参/裸形全摆上，
`--emit-c` 逐行看，比上一轮"一次一条"更快也更硬）。改前只有两条坏：
`Me.Cls` → `vb6_ComCall(vb6_hwnd_Me232Form, L"Cls", NULL, 0)`、
`Me.Print "AB"` → `vb6_ComCallObject(vb6_ComGetObjectProp(同一 HWND, L"Print"), L"Item", …)`；
改后十三形全部落到真出口，`Me.Cls` 与 `Me.Cls()` 从此给**同一条** `vb6_ControlCls((void*)vb6_hwnd_X);  /* Form.Cls */`，
`Me.Print "AB"` / `Me.Print ("CD")` → `vb6_ControlPrint(…, BSTR)`、裸 `Me.Print` → `(…, 0)`（VB6 的空行），
而 `Pic1.Cls` / `Pic1.Print "GH"` 两形**连尾注释都逐字节没动**（`/* PictureBox.Cls */` 那一份原样保住）。

**收成一处（三条码头 + 一处打标记问同一张表）**：
① 表 `controlCanvasMethod` 加 `cls` / `print` 两档，接收者给 `Form` 与 `PictureBox`；`line` 那一档**仍只给 PictureBox**
—— Form 的 Line 是 12 参签名 `vb6_Form_Line`（带 Step 相对位），签名不同不能并表，那是账 #224 剩下的口径，别顺手并进去。
② 表达式码头（`cgen_expr_call_callee_withm.inc` 原 Fix 185 那块）改成 `formCtrlSlot` + 表驱动，实参形状按成员各自签名拼
（cls 一条不交、print 交第 0 条、缺实参交 0 = 空行）；同文件下面的 Line 码头，接收者折开也改问 `formCtrlSlot`
（它以前自己查 `knownFormControls_`，只认裸小写名那一形态）。
③ 语句码头（`stmt/cgen_call.cpp` 原 PictureBox 硬编码那块）同样改问表 + `formCtrlSlot`。
④ 打标记那一路（`cgen_expr_member_form_builtin.inc:401`）撤掉 `print` / `cls` 名单，只问表 —— 表加一档它自动跟着长
（这一处以前正是"表加一档、它漏一档"的洞）。
⑤ 撤掉重复的答案：Form/Printer 绘图段里 `cls` 的 **Form 那一支删掉**（画布两档已由表先接走），只留 Printer ——
`Printer.Cls` 在 VB6 是"结束文档"(`vb6_Printer_EndDoc`)，不是清画布。

**判据三面**：
① 真跑夹具 `tests/fdraw/FDForm.frm` 加 FD18 / FD19 / FD20，每条两头钉（墨 + 笔位）。
FD20 要写成**增量**（`Me.CurrentY - penBefore` 对 `TextHeight` 取等）：第一版写成 `Me.CurrentY = thA`，
在改前那台上因为继承上一行裸 `Print` 的推进而**假绿**（读数 16,16）—— 判据钉错方向的现场，留着当反面教材。
BASE 那台同一份夹具 = 两架构三条**全 False**（`FD18 raw=3,-1` / `FD19 raw=-1` / `FD20 raw=0,16`），
新台两架构三条全 True（`FD18 raw=3,-1` 的两头由笔位那一半负责红：改前 Cls 没跑，但墨被窗口自己重画抹掉过，
所以只看墨会读数一样 —— 这也是为什么每条都要两头）。
② 发码针 `Test-CodegenNote "form_canvas_family"`（顺带把账 #224 欠的 ⑤ 一次还掉）：五枚必须出现 + 三枚必须不出现
（`vb6_ComCall(vb6_hwnd_FDForm` / `vb6_ComGetObjectProp(vb6_hwnd_FDForm` / `vb6_ComCallObject(`）；
BASE 产物实测"缺 2 枚 + 命中 3 枚各 1"⇒ 两头都真能红。
发码针的一条格式规矩（本刀踩过，代价是一整轮门 #367 红在自己身上）：`Invoke-CodegenProj` 比对之前先把 emit 输出做 `-replace '\s+',' '` —— 所以 needle 里不许写连续两个空格。我第一版照抄产物的对齐写成 `);  /* Form.Cls */`（两个空格）⇒ 永远匹配不上：本地 `-Category syntax` 164/1、CI 同一枚红，改成单空格即对。要钉**逐字节形状**（含对齐与续行）得用 `Test-EmitcShape`，只有那一条不折叠空白。
③ 哨兵 `scripts/check_form_draw_state.ps1` 加 S9 四条：S9.1 三个 C 出口名只许住在表文件里（任何码头自己拼名字 = 又一份答案）；
S9.2 表里 `cls` / `print` / `line` 三档与 `Form` / `PictureBox` 两档接收者齐；S9.3 调用点恰好 7（定义 1 + 声明 1 + 码头 5）；
S9.4 打标记那一路的 PictureBox 判据必须问表、且不许把成员名抄成名单。**只看代码行**（S8 那条"注释里出现名字把判据读哑"
的教训已经交过学费）。两条假改动各证红：码头自己拼 `vb6_ControlCls` → S9.1 + S9.3 双红；名单抄回打标记处 → S9.4 红。

**护栏**：语料 A/B `inputs=90 changed=0 same=90` —— 这一刀只改"谁能答"、不改"答什么"，所以**逐字节相同**才是对的判据
（与账 #231 / #234 同一口径）；25 道 [STATIC] 全跑一遍 rc=0（门 #344 那条"往一族加站点要同时跑别人的名单哨兵"的教训）。

**边界与下一格**：`With pic : .Cls` / `.Print "x"` 仍没接 —— With 那条码头问的是 `controlZeroArgMethod` /
`controlOneArgMethod` 两张**按实参个数**分的表，画布家族不住在那里；要接就把 With 码头也改问这张表
（与账 #224 的 ③「With 里带实参那形没接」是同一个问题，一起定口径）。剩下两格照旧：① 裸 `Cls` / `PSet` / `Circle`
发成未声明的 C 调用（C2065；语料 0 处，但先于 ③ 做 —— `Print` 那条"窗体上下文语句"改写已有先例可以照）；
③ 裸 `Line` 两形在 parser 就拒。另记一条形状：`Print` 现在有两个入口名 —— 语句形 `Print "x"` 由 parser 打了
`isFormPrint` 直调 `vb6_Form_Print`，画布形走表给的 `vb6_ControlPrint`；账 #239 之后这两枚是**同一份身体**
（前者转调后者），等 ① 那一刀把语句形也收进同一张表时一起归一。

**① 的开工家底（同一轮只读盘点 + 实测，file:line 都核过）**。探针 `.build/b463probe/B232A.frm`（一份 .frm 里把裸形与 `Me.` 形各摆两条，`--emit-c` 逐行看）读数：

| 写的形状 | 今天发出来 | 判 |
| --- | --- | --- |
| `PSet (100, 100)` | `PSet(100, 100);` | C2065，编不过 |
| `PSet 120, 120`（无括号） | `PSet(120, 120);` | 同上 |
| `Me.PSet (110, 110)` | `vb6_Form_PSet((void*)vb6_hwnd_B232A, 0, 1, 110, 110, 0, 0)` | 通（当基准） |
| `Circle (200, 200), 50` | `Circle(200, 200, vb6_VariantEmpty(), 50);` | C2065；**还自己补了一枚 Empty 实参** |
| `Circle 220, 220, 60` | `Circle(220, 220, 60);` | C2065 |
| `Me.Circle (210, 210), 55` | `vb6_Form_Circle((void*)vb6_hwnd_B232A, 0, 1, 210, 210, 55, 0, 0, 0, 0, 0, 0, 0, 0, 0)` | 通 |
| `Cls` | `Cls();` + 一条 VB3001「未声明的标识符」 | C2065 |
| `Me.Cls` | `vb6_ControlCls((void*)vb6_hwnd_B232A);  /* Form.Cls */` | 通（② 那一刀刚接的） |

三条结论：

**①-a 两形是两条不同的路**。`Cls` 走 `stmt/cgen_call.cpp:420` 那条 "Fix 010m 裸调用"（`callExpr` 里没有 `(` ⇒ 自己拼实参表，:581 `callExpr += "(" + bareArgList + ")"`）；`PSet (100, 100)` / `Circle (200, 200), 50` 走表达式路（postfix 已经把括号里的东西做成实参表，所以 `Circle` 那枚 `vb6_VariantEmpty()` 是**可选形参补齐**补出来的）⇒ 裸形一旦进了正确的分派，那套补齐与 Step/hasXY 旗标都不必重写。AST 侧也印证（`--dump-ast`）：裸 `Cls` 的 callee 是 `IdentifierExpr`，而 `PSet (...)` 的 callee 是 `IndexOrCallExpr(IdentifierExpr, args)` ⇒ **折叠器两种都要认**。

**①-b 折叠的位置在语义层，不在 cgen**。`SemanticAnalyzer::visit(CallStmt&)`（`src/semantics/semantic_analyzer_stmt.cpp:219`，今天整函数只有 `analyzeExpr(*node.callee);` 一行）是这一族唯一"语句已成形、符号查得到、文档种类也知道"的位置（`currentModule_->docKind` 由 `driver_frontend.cpp:254-257` 一处写：UserControl / PropertyPage / Form / Standard）。在那里把"callee 是查不到符号的裸标识符 + 名字是画布动词 + 本模块是有画布的那三种 docKind"折成 `Me.<verb>`（`MemberAccessExpr(MeExpr, 名)`；`IndexOrCallExpr` 那一形就换它内层的 callee）—— 折完之后**下游一行都不用改**：`Me.Cls` 走 ② 那轮的画布表、`Me.PSet` / `Me.Circle` 走 withm 的 Form/Printer 段，`Option Explicit` 那条 VB3001 也自然停掉（名字不再"未声明"）。反过来若在 cgen 的两条码头各拦一次 = 第 5、6 份答案，正是 ② 那一轮刚清掉的形状。

**①-c "画布动词"这份名单必须有唯一的家，而且要两层都能问**。现状是三处各自硬编码：`cgen_util_ctrl.cpp::controlCanvasMethod`（cls / print / line）、`cgen_expr_call_callee_withm.inc:636` 的 `isPrinterDraw`（pset / line / circle / point / cls）、以及语义层也在问的那张宿主伪成员表 `src/common/host_pseudo.hpp:73`（`usercontrol` / `cls` 一行，`HPF_METHOD` ⇒ 限定名才生效，裸名要 `HPF_BARE` 那一位）。本仓对"两层都要问的单一出口"已有定死做法 —— `host_pseudo.hpp` / `float_literal.hpp` / `int_literal.hpp` 都住 `src/common/`（账 #159 / #188 / #194 那三轮的结论）。所以这一刀的**第一步是新增**`src/common/canvas_drawing.hpp`：一行一个动词，字段 =（小写名、哪些接收者有这一档、**实参形状那一档**），然后 `controlCanvasMethod`、`isPrinterDraw`、语义层的折叠判据三处全改成问它。"实参形状"那一档必须有：`Line` 在 Form 与 PictureBox 上签名不同（12 参 `vb6_Form_Line` 带 Step/hasXY，对 7 参 `vb6_ControlLine`），这正是 ② 那一轮**没有**把 line 给 Form 那一档的原因（上面记着）；形状收进表之后，"签名不同不能并表"这条边界就变成表里的一个取值，而不是两处代码。

**开工顺序**：(1) 立 `canvas_drawing.hpp`，先把 `cls` / `print` 两档搬过去 —— 预期 A/B `changed=0`；(2) 把 `pset` / `circle` / `point` / `line` 的 Form 档搬进同一张表，withm 的 Form/Printer 段改成按表里的形状 packing —— 仍应 `changed=0`；(3) 语义层的折叠器上线（这一步才是 ① 真正修好的时刻：`Cls` / `PSet` / `Circle` 三形从 C2065 变成通），判据两头钉 = `tests/fdraw` 加一形真跑像素证人 + 一条发码针（"三形都不许再出现未声明裸调用，也不许出现 `vb6_ComCall`"）+ 负控用改前那台数 C2065 的条数。**③ 那一格必须另开**（裸 `Line (0,0)-(10,10)` 在 token 层就报错：`TokenKind::Line` 在 `parser_stmt.cpp:136` 只认 `Line Input`，而坐标对续画的吸收住在 `parser_expr_postfix.cpp:183-189` 且**只认 callee 是 MemberAccessExpr 且成员名是 line** ⇒ 语义层折叠救不了它，得动 parser；动 parser 时按 ①-c 那同一张表放行，不要再列第四份动词名单）。

**两件别顺手做**：(a) 别把裸形折进 `cgen_file_io.cpp` 的 `isFormPrint` 那条 —— 那是 parser 认 `Print` 是**关键字**才有的路，`Cls` / `PSet` 不是关键字，照抄就要动词法 ⇒ 白多一份形状；(b) 折叠判据里"名字查不到符号"这一问必须留着 —— 用户自己写 `Sub PSet(x, y)` 时那枚过程**该**赢（VB6 的模块内作用域），无条件折就是"修一处静默、造一处调错函数"。
### B71 门 #369/370 那两条红只有 runner 上现形 —— 本机那台 cl 压根不诊断「实参过多」（账 #240，**已出：门待回填**）

**读数**：门 #369（run 37456420314、head `39d9b119`、branch dev）11 job 里两片红，各红一条且是同一枚夹具的两个架构 —— `Tests (vbp #2)` = `[VBP-BUILD] olecon ... FAIL rc=1 exe=False`（该片 PASS=54 FAIL=1 SKIP=1）、`Tests (vbp #3)` = `olecon_x86`（PASS=53 FAIL=1 SKIP=0）；其余九片全绿。**引入方式不是改了产品**：`39d9b119` 那轮新增 [STATIC] vbp_fixture_census 把五份"跟踪着却没登记"的 .vbp 逼出册登记成编译面用例，olecon 是其中一份（提交说明里写着本地 x64+x86 rc=0 且出 exe）。同批登记的 dbgdlg（就是那枚缺 `vb6_di_PageSetupDlgA` 桩、为它才补的夹具）在两片上都 PASS ⇒ 桩表与 RTL 内嵌在 CI 上是对上的，红只跟着 olecon 走。

**已排除的六条**（每条都有实物，不是推理）：(1) 夹具没进仓 —— `git ls-files tests/olecon` 有 .frm+.vbp 两份，且目录里根本没有 .frx（那条 `oc_src.bin` 只在运行期读）；(2) CI 那台 C3.exe 与我本地这台不同 —— 把 run #369 的 `c3-exe` 工件下载下来真跑，x64 与 x86 都 rc=0 出 exe；(3) `-Incremental` —— `Test-VbpBuild` 压根不传 `--incremental`；(4) 架构/命令行差异 —— 本地按登记时的两条命令行（默认 x64 与 `--arch x86`）逐字复跑；(5) RTL 内嵌资源 id 对调（账 #225 那一族）—— 那样会全线 LNK2005×1225，不会只有一条红；(6) 源码本身依赖注册表里的 VB6 类型库 —— olecon 走的是仓内原生那一条（`driver_link.cpp:45` 每次都带 `vb6forms_olecon.c`，产物里全是 `vb6_OleCon_*`/`vb6_RegisterOleConClass`，没有查注册表的路）。

**剩下的唯一差异是机器**：runner 用 vswhere -latest（VS2022 + 新 SDK），本机只有 VS2019 14.29.30133 + SDK 10.0.19041。红出现在 `runLinker` 那一段（只有那一支才打 `intermediates kept at`），而 cl/link 的整段输出只落在 `c3-error.log` —— 它既不被 `Test-VbpBuild` 打印，也不在 CI 的工件通配符（`output/**/*.out|*.err|*.txt|*.dat|scores.txt`）里 ⇒ **门上看不见病因**。

**本刀（一）**（`tests/run_tests.ps1`，+23/-0，纯测试面）：`Show-BuildErrorLog` 接进 `Test-VbpBuild` 的失败分支，挑 `error C####` / `: error ` / `fatal error` / `LNK####` / `unresolved external` / `=== C3 Diagnostics` 那几行（最多 25），一条都不匹配时退回尾巴 15 行。判据形状来自本地实物：一枚刻意失败的工程（b475）日志 264 行，262 行是 RTL 的 C4819/C5105/C4028 警告，直接摊 40 行尾巴会把唯一的 `error C2063` 挤出去。

**归因（门 #370 的 [diag] 读数，两片各 4 行，一模一样）**：`Form1.c(100): error C2197: 'void vb6_OleCon_Init(void *,const wchar_t *,int,int,int,int)': too many arguments for call` —— **4 行 = 10 个实参减 6 个形参**，一枚多余实参报一行。对着源码量：定义 `src/rtl/core/vb6forms/vb6forms_olecon.c:920` 是 10 参（`... autoActivate, autoVerbMenu, borderStyle, sourceDoc, sourceItem`），发码 `src/backend/detail/module/cgen_form_ctrl_style_apply.inc:970` 也发 10 个，只有 `vb6forms_prop_ctrl.h:334` 那份原型还停在 6 参 —— 体 grew 上去、头没跟。**为什么只有 runner 红**：VS2019 的 cl 在 C 模式下对「实参多于原型」根本不诊断（本机用 6 参原型 + 10 实参的最小夹具 `b481/t10.c` 实测 rc=0，只在类型对不上那一枚上给 warning C4024），新 cl 把它按 C 标准的约束报成 error ⇒ 本机真编两遍都编不出这个病，判据必须换形状。

**修法（本刀二）**：`vb6forms_prop_ctrl.h` 的原型补齐成 10 参（+3/-1，只动头；`src/rtl/**` 改了要 touch `src/driver/c3rtl.rc` 再重编 C3.exe，否则内嵌的还是旧字节 —— 账 #156 那条）。运行面零改动：10 个实参本来就一直发着，本机那台把多余 4 枚照 cdecl 传过去了，所以旧产物行为不变；这一刀只是让**下一台编译器**也认。本地验：新 C3.exe 真编 olecon `=== x64 rc=0 OleCon.exe 465920 字节 / === x86 rc=0 OleCon.exe 414720 字节`（这两个数与登记那轮记录逐字对上）。

**本刀（三）= 结构性哨兵**`scripts/check_rtl_proto_arity.ps1`（第 27 道 —— 本轮之前实测 26 份 check_*.ps1）+ `run_tests.ps1` 的 `[STATIC] rtl_proto_arity`：判据 = RTL 里**两头都有**的名字，声明侧参数个数集合必须等于定义侧（R1），外加 census 地板「文件数 >= 100 且比较对数 >= 900」（R2，实测 127 份 / 1237 对）—— 路径写错或正则被改坏时不许变成"绿着的空转"。两条设计约束记下：① **只有第 0 列开始的行算签名**，这一条同时把所有调用点排干净（调用都缩进在函数体里），② 只比**个数**不比类型拼写（头写 `const X*`、体写 `X*` 是合法的，硬比只造噪声），认不出的参数形态（数组/函数指针/默认值）跳过、不计入也不报红。**负控两头跑过**：把头削回 4 参 ⇒ rc=1 并点名 `vb6_OleCon_Init: 头 4 参 (…prop_ctrl.h:334) 对不上 体 10 参 (…vb6forms_olecon.c:920)`；改回 10 参 ⇒ rc=0（同一份 census 读数 1237 对）。

**这一轮为什么不跑语料 A/B**：本刀只动一份 RTL 头，而 `--emit-c` 的产物里压根没有 RTL（记忆里那条老读数），C++ 源一行没动 ⇒ 编译器的发码逻辑同一个程序，只有内嵌的 RCDATA 变了。90 份 emit 的 BASE 那台已被覆盖，拿它比只会量到这一个月的**夹具漂移**（账 #239 那条教训：改了夹具再跑 A/B = 假归因），所以换成正对靶子的三面：
① 发码实物 = 10 个实参（本地 keep-for-debug 的 `Form1.c:100` 逐字读过）；② 定义 = 10 参；③ 头补齐后 = 10 参 —— 三头同值，再加两架构真编 rc=0 出 exe。**这一族的边界（记下别越界）**：哨兵管"头追不上体"，管不到"发码递的实参个数 ≠ 体"。真要钉那一头得让**每个控件方法的发码形状**与 RTL 原型对账，那是把 `controlOneArgMethod` / `controlZeroArgMethod` 那几张表的签名也拖进对账面的一件大活（且只有新 cl 才看得见后果）—— 已另立 §B72 记着，本轮不顺手做。



## C. 仍在生效的口径与工具事实（与本文档等长的一半价值在这里；完整版见记忆库）

- **子类化分层的槽位口径（账 #185 起）**：RTL 里**每一层**窗口子类用**自己**的窗口属性名存它下面那层的 wndproc ——
  `VB6_OrigProc` = 发码的事件层、`VB6_ImageOrigProc` = PictureBox/Image 的自绘层、`VB6_GBox_OrigProc` / `VB6_GfxBtn_OrigProc` /
  `VB6_SSTab_OrigProc` 各自一层，而「这层装过没有」那一问**只看自己那层的名字**。两层同名 = 后装的那层静默不装，
  症状是「处理器编得过、消息臂发得对、一次也不响」—— 这类缺陷只有运行期看得见，所以判据必须带一枚**没人跟它抢的证人**
  （本线用的是同窗体上的 Label：同为 STATIC，只是样式不含 SS_BITMAP）。画的序也收成一条：最外层 BeginPaint/EndPaint 一次，
  DC 经 `VB6_PaintDC` 交给下面那层画表面，再抬用户的 `_Paint`。哨兵 `scripts/check_subclass_slots.ps1`（S1 一名一文件）拦的就是同名。

- **`AddressOf` 的调用约定口径（账 #184 起）**：VB6 的 `AddressOf` 交出去的是 **`__stdcall`
  调用桩的地址**，不是本体地址；C3 现在按它在**定义模块**里发桩（口径一处：
  `CCodeGen::addressOfTargetCName`），本体保持 `__cdecl`。判这类刀时记三条：
  ① 小夹具不算红判据——MSVC 写的调用方有 EBP 帧，`leave` 会把 ESP 拉回来，所以
  `EnumWindows(AddressOf cb)` 在改前也读数全对（实测 HEAD~1 与 HEAD 两份产物逐字相同）；
  真判据要用**不自我吸收的调用链**（comctl32 的 `SetWindowSubclass` thunk）或发码形状针。
  ② 页堆只能把「谁在读已释放堆块」钉成确定现场，钉不出「谁按 stdcall 调它」——
  把归因钉死的是**一次「关掉一条路」的对照实验**（拷一份工程出去给 `FlexSetSubclass`
  体首加 `Exit Sub`：崩溃 3/4 ⇒ 0/6）。
  ③ A/B 分类器要允许「纯改名」这一类差异（判据 = 两边都把 `aoThunk_` 去掉再比），
  并把 `C3:` 开头的诊断行滤掉——`--emit-c` 的 stdout 与 stderr 混流，census 行会伪装成
  「多了一行」。

1. **RTL 是嵌进 `C3.exe` 的 RCDATA**：改 `src/rtl/**` 必须重编 C3.exe 才生效，真凭据是构建日志里
   出现 `Building RC object CMakeFiles\c3.dir\src\driver\c3rtl.rc.res`。新增 RTL 文件还要同时进
   `C3RTL_EMBEDDED_FILES`（`CMakeLists.txt`），否则照编不误却永远没嵌入（见 B6 的三方不变式）。
2. **符号化崩溃必须 `-g`**：`--keep-for-debug` **不产 PDB**，那种 exe 的 RVA 一个名字都出不来。
   x64 用 `.temp/sym3.ps1`（基址 `0x140000000`），x86 用 `.temp/sym_x86.ps1 -Exe <exe> -Base <hSelf>`
   （`sym3.ps1` 硬编码 x64 基址，x86 上是错的）。`/Od` 下栈扫描有假阳性，RVAs 大于镜像即丢。
   给人手测崩溃用 `.temp/case190g.ps1`（`-g --keep-for-debug --output-dir .temp\case190g` +
   `C3_CRASH_TRACE`/`C3_COM_TRACE`，留进程不杀）。
3. **改 per-control GDI 对象成共享缓存前先 `grep` 谁删它**：`vb6_SetControlFontFromLogFont`
   （`vb6forms_ctrl.c`）无条件 `DeleteObject(hOld)`，只挡 `OBJ_FONT`，进程级缓存同样会被第一个写
   `Font.*` 的控件销毁（GDI 会复用句柄值 → 别处字形乱变）。
4. **"0 个 MSVC 错误 / 它能跑" ≠ 健康**：C3 只在**有错时**写 `_c3_msvc_out.txt`，警告全丢，所以
   `void*` 喂给 typed `X*` 这类静默误编会给出绿色构建 + 运行期崩。`/W3` 的诊断事后不可恢复。
   同理"截图截到了"也不代表没崩：`WM_CREATE` 里的 AV 常常照样留下窗口 —— 用重定向的 stderr 判活。
5. **`--dump-*` 打在 `.vbp` 上不是只读**：它 dump 完会继续 codegen + 调 MSVC（别人门禁在跑时会抢 CPU）。
   要 dump 就单 `.bas` + `--output-dir .temp/scratch`。另 `--dump-symbols` **从不打印 UDT 成员字段**，
   回答不了"`typeRefName` 填了没"，这类问题去看生成的 C。
6. **类型名解析对了 ≠ 载体/落点对了**：`TypeSystem::resolveTypeName` 是**首命中即返回**的启发式瀑布，
   语义层与代码生成共用它，所以错得很"自洽"（表现成"类型一致，只是符号缺失"而不是不匹配）。而
   `mapSaElemType`/`mapSaElemCType`、`inferExprType`（会拿**元素**类型回答整数组引用）、按 **C 类型串**
   登记的键（`Date`/`Double`/`Currency` 都发 `double`）都是**独立的第二道映射**，哪里漏了都会
   "绿色构建 + 运行期崩/静默错值"。问"改了没生效"先看 `--dump-symbols` 与生成的 C。
7. **门禁与验证**：`.temp/gate.ps1 -Tag regressNN` → `output/yqt_regressNN.log`（UTF-8；裸 `>` 重定向
   给 UTF-16LE，`grep` 读不出东西）。快验收只跑 GUI 面用 `.temp/gate_vbp.ps1`（`-Category vbp`，约
   255 秒 vs 全量 20–35 分钟）。判完成 = 有 `Results:` 行 **且** 无存活 C3/cl/link/ninja。
   两类**环境性**红要先排除再相信：`fatal error C1060`（散在无关 `rtl/*.c` = 内存压力，重跑）、
   大面积 `FAIL (compile)` + 某段异常快 = `INCLUDE` 被污染（把 `scripts/dev.ps1` 和门禁链进同一个
   PowerShell 进程即触发，**必须分两次启动**）。
8. **写用例的已知坑**：`Test-Syntax` 只看退出码（证明不了运行期语义，`Add-BasTest` 才会跑）；
   VB6/C3 的 `String` 是 Unicode，`StrPtr` 给 **UTF-16** 字节；`vbNullChar` 编成 `vb6_BSTR_Empty()`；
   控制台输出用 `Debug.Print` 而不是 `Print`。
9. **共享工作树纪律**：本文件与整棵树同时有另一个会话在写。**动手前后都 `git status --porcelain`**；
   `git diff --numstat | awk '$2 > $1+5'` 抓"一次 Edit 静默删掉大块"（曾把 `vb6rtl_builtin.h` 从
   315 行砍成 144 行，症状是**所有**测试红）；提交**按位置逐 hunk 筛**，不整文件 `git add`；
   不杀别人的 `C3.exe`/`cl.exe`，不 `checkout`/`restore` 别人的文件，不用裸 `git stash`。
10. **临时目录会被冲掉**：`%TEMP%\C3C\<pid>\` 里放的是 `_c3_msvc_out.txt` 与生成的 `.c/.h`，
    任何一次后续编译（含测试套件）都可能删掉 ⇒ 构建的**同一条命令**里把产物快照到 `.temp/gen/`。
    取会话目录用 `grep -a "kept at" | tr -d '\r' | sed 's#.*kept at: ##'`（`[0-9]{13}` 会截断 15 位 id）。
    生成 `.c` 的行号**不映射** `.bas` 行号。
11. **控件族"方法"的调用侧有两条互不相通的路径，只改一处会剩一半 no-op**（Fix 185 落地时实测）：
    带实参的 `Pic.Print "x"` 走 `src/backend/detail/expr/cgen_expr_call_callee_withm.inc` —— 该片段
    include 在 `cgen_expr_call_com_bind.inc` **之前**（`src/backend/expr/cgen_expr_call.cpp:42-43`），
    所以新的控件特判必须追加在 withm 末尾才抢得到通用 COM 绑定前面；**无括号**的 `Pic.Cls` /
    `List1.Clear` 根本不进 withm，而是落在 `src/backend/stmt/cgen_call.cpp` 的 `isComMarker_` 分支
    （Fix 086 的 `List1.Clear` 就是同一形态先例）。配套协议：语句自己 `c_.emitLine(...)` 之后把
    `lastExpr_` 置 `"0"`，`cgen_call.cpp` 的 `else if (callExpr == "0")` 分支保证不会再补发一条裸 `0;`。

12. **GUI 真跑的取像与取属性有两个工具假读数，先修工具再下结论**（2026-10-05，量 VbEclipse 四台时撞的）：
    ① `Graphics.CopyFromScreen(窗口矩形)` 截的是**屏幕上那块像素**，我们的表单常被别人的窗口盖住 —— 实测截回来的是用户那个聊天窗口，
    看着像"程序画成这样"。要取窗口**自己**的像素就用 `PrintWindow(h, hdc, PW_RENDERFULLCONTENT=2)`（被遮挡也取得到）。
    ② PowerShell 里 `[DllImport("user32.dll")]` 不带 `CharSet` 时按 **Ansi** marshal `StringBuilder`，而 `GetWindowTextW`/`GetClassNameW`
    写的是宽字符 ⇒ 取回来永远只有**第一个字符**（"Button"→"B"、"Round Corners"→"R"），看着像产品把控件文字截了。
    加 `CharSet=CharSet.Unicode` 才是真数。两条的共同点：**红的是探针，不是产品** —— 与 #161 那条"探针缺样式位冤案"同族。
    现成工具：`.build/b210_click.ps1`（列子窗口 + 按 caption 找按钮发 `BM_CLICK` + PrintWindow 取像 + 只杀自己起的 PID）。
13. **夹具的编码不统一，改别人那份要按字节改**：`tests/frxdata/FrxData.frm` 是 **GBK**（里面那句中文列表项是判据的一部分），
    而 `tests/dcsurf/DcForm.frm` 是 UTF-8 —— 用按文本读写的方式改 GBK 那份会把注释与中文字面量整段换掉。
    写这类补丁的规矩：二进制读、只插 ASCII、写完用 `decode('gbk')` 自证，并核 CRLF 数与 lone-LF=0。
14. **提交进仓的 .ps1：BOM 只认 EF BB BF，中文另有一道吞字节的坎**（本轮写 `scripts/check_sa_access.ps1` 撞的）：
    ① 把 BOM 手写成 EF BF BB 不是 BOM，那是个合法字符 (U+FFFB)，powershell 5.1 与 pwsh 7 都会把首行当命令名，报
    `?# 无法识别` —— 看起来像"BOM 会坏 PowerShell"，实际是字节序写反。仓库里 `tests/run_tests.ps1` 是 EF BB BF + CRLF，照它。
    ② 无 BOM 时 5.1 按 GBK 读 UTF-8 字节：**行尾中文字的最后一个字节会被当 GBK 首字节，吞掉紧随的那个 ASCII** ——
    吞掉双引号 ⇒ 字符串未闭合（ParserError 指向**下一行**，红因看起来毫不相关）；吞掉换行 ⇒ 下一行并进注释（参数/语句全空）。
    CI 那份 `check_scalemode_writers.ps1` 是"中文只在注释、且注释行后面还是注释或空行"才侥幸活着。
    ③ 稳妥写法：消息串一律 ASCII，中文只放注释，且每条中文注释行以 ASCII 字符收尾；写完全文核 `lone-LF=0`。
    本轮另有一条同族旧坑复发一次：python 里写 `b"\r\n"` 要想清楚 —— Bash 工具会先折一级反斜杠，
    落到文件里就是真换行，把补丁脚本自己写坏。反斜杠一律 chr(13)/chr(10) 拼。
15. **`C3: AddressOf callback procs: N` 数的不是过程个数，是符号副本**（2026-10-05 实测，读法见 `src/driver/driver_crossmod.cpp` 里那条注释）：
    `markAddressOfCallbacks` 内层那圈对**每个模块的符号表**各 +1（定义模块一份 + 每个引用它的外部副本一份），
    所以 VBFlexGridDemo 源码里去重后 37 个 `AddressOf` 目标名，这一行报 **49**。当「这条路走没走」的信号够用；
    **待办**：按过程名归并后再报（去重），下轮有别的源码刀时顺手改，不为它单开一轮门。
16. **读 MSVC/C3 的中文告警必须显式按 gbk 解码，否则 census 会造出不存在的条目**（2026-10-05，账 #217 撞的）：
    C3 的诊断正文是中文、按控制台码页写进日志。用 UTF-8 + errors='replace' 读时，名字前面的中文字节会连吃字符
    —— 我因此把 `OLEGuids.IObjectSafety` / `OLEGuids.IOleInPlaceActiveObjectVB` / `OLEGuids.IOleControlVB` 三条
    归成一簇叫 `nterface` 的假条目，还据此在 §B49 记了一条不存在的账。同一份日志 `decode('gbk','replace')` 重读，
    502 条落在 14 组名字上、一条不差。同族第二条坑：正则的 `^` 不加 `re.M`，在整篇文本里只匹配文件开头，
    会得到"这条日志里 0 条告警"这种假阴性 —— 而 `grep -c` 明明报 502。**数字与工具对不上时先怀疑读法**，
    别拿第一个读数分家。

17. **.ctl / .pag 可以单独喂 `--emit-c`**（2026-10-06，写账 #217 第二刀的判据夹子时确认）：
    不必为"编译面判据"造一整份工程（.vbp + 宿主窗体 + .frx），一枚带 `Begin VB.UserControl X` 头行和
    `Attribute VB_Name` 的 .ctl 就能直接跑 —— 走的是 driver 同一个按扩展名分派的入口。
    配套的两条读数口径：① `Test-CodegenNote` 的 Absent 钉 VB3001 时**要钉 ID 而不是钉中文正文**
    （正文按控制台码页写，跨码页不稳，见 §C16）；② 全仓 VB3001 总数是这类"跨工程同一族"改动最好的
    横截面判据（本格 2934→158），比逐工程数数更难被局部巧合骗过。

18. **RTL 导出的 C 名字与用户模块级变量共用同一个名字空间**（2026-10-06，账 #220）：生成的模块 C 会
    #include 那批 RTL 头，而 `Public B As Long` 在 C 里也是**裸名** —— 所以 RTL 里头文件 extern 的裸名全局
    在**编译期**撞（C2373 重定义 + C2166 给 const 赋值），.c 里**非 static** 定义的在**链接期**撞（LNK2005），
    只有 `static` 的不撞。口径：RTL 只用 `vb6_` / `VB6_` 前缀导出名字；语法旗标（`Line` 的 `B`/`BF` 这种）
    一律由 parser 折成字面量，**不许**为了让发码"有个名字落脚"而在 RTL 补一枚全局。
    哨兵 `scripts/check_rtl_naked_names.ps1` 的 N2 把现存名单钉死（5 枚，`Changed` 那枚由账 #219 收），
    负控 = 往 `vb6rtl_com.c` 插一行 `int32_t b220probe = 0;` 立刻红并点名。这条口径的两头各有实物（同一轮探针 `.build/b229out/`，两台都 no exe）：`Public Changed As Long` ⇒ **C2371 重定义；不同的基类型**（头里 `extern int16_t Changed;` 那一枚，编译期撞）；`Public g_hoCount As Long` ⇒ **LNK2005 + LNK1169**（只在 `uc_host.c` 里非 static 定义、它那个头没进生成的模块 C，链接期撞）。读法一条：cl 的诊断**不在 C3.exe 的控制台输出里**，只在 `<output-dir>/c3-error.log`（按 gbk 解），否则会出现「BUILD-RC=1 且控制台 grep error C 得 0 条」这种假象。

19. **一条控件方法要"两头都接"才算接上：成员侧打标记 + 调用侧查表**（2026-10-06，账 #221 踩的）：
    后端那张"表只交名字、实参由码头拼"的做法（`controlZeroArgMethod` / `controlOneArgMethod` /
    `controlScaleMethod` / 现在的 `controlCanvasMethod`）只在**成员侧把 `comObjExpr_` 留成小写控件名**
    时才拿得到控件类型；成员侧不打标记，兜底 Fix 023e/089d 已经把 `comObjExpr_` 换成 HWND 表达式，
    调用侧那张表**根本不会被问**。症状与账 #143 一模一样：产物照旧 `ComGetObjectProp(hwnd, L"方法名")`
    再取 `Item`、两跳都 `S_OK`、一笔不画、一条诊断都不打。所以新加一张这种表时，
    `scripts/check_rtl_naked_names.ps1` 的 N6 四条（声明 1 / 定义 1 / 调用侧 ≥1 / **成员侧 ≥1**）
    必须四条都绿才算这一格做完；只看到"表建好了、码写好了"就提交，等于交一半。
    同批两条夹具读数纪律：画完问像素要放在 **Timer 第一拍**（Form_Load 里 `GetPixel` 全 -1 = CLR_INVALID），
    DC 用 `GetDC(控件 hwnd)` 而不是 `控件.hDC`（后者在这枚夹具上读出 0，属 #196 那一族的另一问）。



- **跨层的判据表住在 `src/common`（账 #219 起）**：一张"哪个成员叫什么 / 能不能裸写"的表（`kHostPseudoRows`）原先在 `src/backend/cgen_util_com.cpp` 的匿名 namespace 里，四个消费点都在发码侧。语义层要问同一件事时**不许把名字抄进 semantics**（抄一份就是第二个权威，本账那 48 条噪声就是这么来的），而是把表搬进 `src/common/host_pseudo.hpp`，两头只问 `hostPseudoBareEligible(obj, member)` 这一句。common 不得向上依赖 semantics，所以表里的 `Symbol::toLower` 换成头文件自带的 `hostPseudoLower`。哨兵 `check_host_pseudo_table.ps1` 的表路径 / `must` / `deny` 三处跟着表搬家，含义是"这张表只许有一个家、语义层只许问它"

- **夹具（`.ctl/.frm/.pag/.bas`）里的注释一律写 ASCII**（账 #222 本轮实测）。C3 读源走 ANSI(GBK) 那一套解，UTF-8 中文注释**只有在字节两两配成合法 GBK 时才不出事**：`tests/ve_units/ucUnitPix.ctl:75` 那条带 `①②③` 的注释字节序配出了 `U+FFFD` ⇒ 当场 167 条 `VB1005 意外字符 / VB2002 / VB2003`、整工程 rc=1；同一行换成全 ASCII 注释就 rc=0。逐变量实测：只把 `①②③` 换成 `A)/B)/C)` 仍然红 ⇒ 踩雷的是这一行里某个字节对，不是某枚特定字符，别拿"哪枚字符不行"去记。老夹具里的中文注释能活下来纯属运气。所以新增判据行时注释写 ASCII，改完先 `--syntax-only` 或直接真编译一遍再下结论。


## D. 已完成项一行索引（叙述已删；原文在 `git show 1465da1:ai/C3_FIX_HANDOFF.md` 的对应 §区间。§B55..§B58 那四节 = 账 #225/#222/#226/#227，四格都过门（#349/#351/#351/#352），叙述在 `git show dd73c049:ai/C3_FIX_HANDOFF.md`）

| 原 § | 内容 | 状态 |
|---|---|---|
| §1–§3 | vbman 无窗体 bisect 基线 65 → 7、含窗体全量 129 条目 → **0 编译错**；阻塞项"陈旧对象"解除 | 已收敛（092r-z 时代） |
| §4, §5, §5b | Fix 084 系列、`m_uData` 模块级 UDT（010n 扩展）、088e/089 系列 552 → 216 | 已提交 |
| §6, §7, §8 | 剩余错误分布快照（已过期）、关键文件地图（已过期：`cgen_*.cpp` 已拆成 `src/backend/detail/{expr,stmt,util,base,module}/*.inc`）、会话纪律 | 快照过期，纪律仍生效（已上收 C 组第 9 条） |
| §9 | ToolsTlsThunks C2224 ×37（COM 集合簇）"建议方向未实施" | 实已实施：`src/backend/detail/expr/cgen_expr_member_generic_access.inc:123-124` |
| §10(链接期) | 443 未解析符号、LNK2005 重定义 | 已由 093a/093b 解决，规范叙述在 `ai/开发历程/88-*`、`89-*` |
| §10(158a–160-G) | VBFlexGridDemo 编译错 236 → **31**：Variant 比较 C2088、`_Generic` 转换族、`Mid$/Left$/Right$` 实参、Property Let 槽位、ByVal String 隐式 CStr、`Form.hWnd` 成员误判、LSet 三处、`MSDATASRC` 别名、`ERROR_NOENTRY` RTL 中断 | 已提交（含 1 次 Edit 连带删 171 行原型的事故记录） |
| §11 前后 | Fix 161（`VB.`/`VBA.` 库限定名被降成 Long）30 → 16、162（Extender Width/Height 无赋值）、163（跨单元无原型 `void*` 调用）、164（DI `unknown` 族缺导入库 → 64 例链接失败）、165（GUI 入口点按启动对象决定 + 找回被生成器删掉的 39 个桩）、166（`&(void*){…}` 多一层间接） | 已提交 |
| §12, §13 | Fix 167（`Sub Main` 驻留语义）、168/169（Extender 结构体守卫、COM `VARIANT*` 裸拼 → Err 380）、170（`arrName() = expr` 整数组赋值当成 0 号元素写）、172/173 —— **demo 第一次真正进消息循环** | 已提交；尾巴见 B7/B10 |
| §14–§20 | 参考图差异逐项定位：174/176（`Debug.Print` 打包）、177（单元格存**悬垂 BSTR**，`477c90f`）、178 前两轮猜测 | 已提交；178 见下一行 |
| §20–§22 | Fix 178：UDT 赋值/`LSet` 对含所有权成员的结构体做**深拷贝**（C 的浅拷贝导致"行 ≥2 别名到行 0"，整格显示第 0 行） | 已落地并验证 |
| §23–§28, §18 | Fix 181：控件默认字体改 VB6 口径 MS Sans Serif 8.25pt + `NONANTIALIASED_QUALITY`；悬垂 HFONT 隐患按"每控件一份字体"收口 | 已落地（尾巴见 B12） |
| §25–§26, §33 | Fix 182：容器子控件**第二条发射路径**缺设计期属性（frx/Text/List…），`b953147`；CellPicture 预览空白由此关闭 | 已提交（尾巴见 B17） |
| §29–§32 | Fix 184：网格只画 13 行的根因是 RTL 里**缇/像素两套 DPI 口径混用**（`*15` 硬编码 vs 真实 DPI），统一走 `vb6_DpiX/Y`；Fix 183 验证失败**已回退并作废** | 已落地（dpi=96 下逐位相同，所以控制台用例不动） |
| §34–§35 | Fix 185：控件级 `_Paint` 从不派发（`Picture2_Paint` 是死代码）+ `Print/Cls` 编成 `vb6_ComCall(HWND,…)` 运行期 no-op ⇒ "Drag/drop me" 框空白。四处协同：`SubclassInfo.hasPaint` + WM_PAINT 派发（`BeginPaint`→挂 `VB6_PaintDC`→调用户过程→`RemoveProp`+`EndPaint`，**`return 0`**）、成员侧 PictureBox 分支（必须落在 023e/089d 兜底**之前**）、调用侧 withm（带实参）+ `cgen_call.cpp`（无括号）两条、RTL `vb6_ControlPrint/vb6_ControlCls`。落地后 demo 画面与参考图**逐项一致**（最后一个已知差异关闭）。另记：`.temp/fix183_block.bin` 永久作废 | 已落地；边界 → B1 |
| §36–§41 | Fix 187：`Declare As String` 只有名字以 A 结尾才做 ANSI 编组 → `GetProcAddress` 返回 0 → Charts2020 自造子类化 thunk `call NULL`（关闭崩溃）；含极小复现与一次无效取证（反向 A/B 未设 `C3_CRASH_TRACE`）的更正 | 已落地；尾巴见 B13/B14 |
| §42 | Fix 188：`Startup = Sub Main` 的工程关窗后进程不退出（没人投 `WM_QUIT`，按 Forms 计数收口） | 已落地；偶发 AV 见 B15 |
| §43 | Fix 189（**仅测试侧**）：门禁 `Test-GuiVbp -AutoExitSec` 无条件 `pass++`，窗口出现后自退崩溃照记 PASS；现改为区分"我们杀的"与"它自己退的"并读退出码 | 已落地（同节两条遗留 → B6/B3） |
| §44 | Fix 190：`As Any` ByRef 实参的 `[]` 下标链判为非左值 → 把元素**值**当 memcpy 目的地址（悬停即写 0x0）；Fix 191：`As LongPtr` 数组按 Variant 载体分配打断手写伪 vtable + `AddRef/Release` 直发槽 1 + `vb6_ComIsDispatchable` 守卫 | 已落地 `1465da1`（遗留 → B4/B20） |
| 合并块 2858–2867 | EXE 工程类 IDispatch 桥接（`VBMAN_DEMO` 启动即崩 → 通过），提交 `3dd3689/17f349a/df42abc/3b797f9` | 已提交（未覆盖面 → B5） |

| 账 #173（提交 `332f6363` = <vbeclipse> rev38，门 #299） | x86 上 `vb6_VariantToDouble` 无原型 → x87 栈泄漏 → Charts 2020 启动即 error 6「Overflow」；补真实原型 + 同族 census （C4013 14→2，剩 B22 那两处） | 已提交并过门 |
| 账 #174（提交 `3684b1d7` = <vbeclipse> rev39，门 #300） | 属性页裸名 `SelectedControls(i)` 改走宿主内建裸名表，发射成 RTL 既有出口 `vb6_PropertyPage_SelectedControls`；新用例 `pp_selctrl_hostmember`，语料 A/B 118/120 一字不动 | 已提交并过门 |

| 账 #175（提交 `e4bf2b17` = <vbeclipse> rev40，门 #301） | 控件坐标的**单位**收成「容器的 ScaleMode」一处权威（`vb6_ScalePxToUser` / `vb6_ScaleUserToPx` + `vb6_ContainerScaleMode` / `vb6_WindowScaleModeSelf`），替掉散在 12+ 处的写死缇；Charts 2020 的饼/柱/面积/矩形不再整幅画在画布外（图体空白），czUI 运行期 `Move` 的字面量恢复像素语义。判据 = `Test-GuiVbp -DumpMinColors`（数 `C3_UC_DUMPDIR` 每控件绘制缓冲的不同颜色；BASE 4..10 / NEW 63..512）+ 结构哨兵 `scripts/check_uc_scale_units.ps1`（BASE 树 20 处红）。同轮量到 #176（见 B23）、TextWidth 量纲（见 #177） |
| 账 #177 + #178（提交 `d5f3e190` = <vbeclipse> rev41，门 #302） | 文字量纲（`UserControl.TextWidth/.TextHeight`）跟着容器声明的 ScaleMode 走，单位表收成一份（`vb6_ScaleUnitsPerPx`，`ScaleX/ScaleY` 转调；顺手订正 6=毫米/7=厘米 抄反）；「控件代码运行在自己的上下文里」补了两处出口 —— 早绑定走发码（`.ctl` 里把 `UserControl.TextWidth(t)` 重定向为 `vb6_UC_TextWidthOf((void*)me,t)`；选宏而不做 enter/leave 配对，因为 `Exit Function` 漏帧是已知欠账），晚绑定走 `OwnPropGet/OwnPropSet/OwnMethodCall` 三处 push/pop。新夹具 `tests/ve_units`（判据写成「缇型 = 像素型 × TwipsPerPixelX」⇒ 与 DPI 无关，CI 两片读数与本地逐字节相同）+ 哨兵两条新规则（出现第二张单位表即红）。红侧实测：#175 产物上 `U-TW/U-TH=False`；像素型控件读数全落在同一产物跑两遍的 ±5% 自抖区间内 |
| 账 #179（提交 `30a3f4ff` = <vbeclipse> rev42，门 #303） | 控件代码不管被谁调都跑在自己的宿主上下文里：发码在 .ctl 实例方法体首 `vb6_UC_PushInstance((void*)me)`、统一出口尾 `vb6_UC_PopInstance()`（谓词 `ucCtxScoped()` 一处，三个发射点共用；能配对的前提是 `Exit Function` 早已走 `vb6_proc_exit`），替掉 #178 那种逐个成员补按实例出口的做法。不变式：charts/czui/flex/ve_units 逐函数 push=pop，无 .ctl 的工程 0/0。针面 `U-CTX`（只取 ScaleMode 一位 ⇒ 与 DPI 无关）。顺带把自己一条错结论订正回 #159（见 B25） |
| 账 #159（提交 `ef260daf` = <vbeclipse> rev43，门 #304） | 宿主伪成员（UserControl / PropertyPage / Extender / Ambient）的「叫什么 / 是什么类型 / 能不能裸写 / 赋值要不要拆」收成一张表 `kHostPseudoRows`（55 行、35 标量），四个消费点 + 拼写规范化都改读它（此前是五份互不相交的清单）。**根因在读法不在值**：类型 oracle 答 Variant ⇒ 发码把 `&vb6_UserControl_hWnd`（8 字节 `void*` 的地址）当 `vb6_VARIANT*`（16 字节）递给 `vb6_VarCmpLongNe` ⇒ `<>0` 恒假；`CStr` 落进 `_Generic` 的 `default: vb6_VariantObject` ⇒ 空串。红→绿：同一夹具 BASE 发 `vb6_VarCmpLongNe(&…,0)`、NEW 发 `(-(… != 0))`；真跑 `U-HW=True`（x64 + x86 两片同值）。--emit-c A/B 六工程只有 charts +12/-12、flex +2/-2、ve_units +8/-8 变，逐行归因全在这一家，czui/ve_list/iface_wrap 零差异、六工程行数未增删。哨兵 `scripts/check_host_pseudo_table.ps1`（逐行对 `vb6rtl_userctl.h` 的 extern 类型 + 禁五份旧清单回潮，两条假针都能红）进 `[STATIC] host_pseudo_table_census`。顺带：`.ctl` 里裸写 `hWnd` 现在能解析。表外欠账见 B26 |
| 账 #180（提交 `6533bae4` = <vbeclipse> rev44，门 #305） | §B 的 B19 落地：`vb6_UserControl_ContainerHwnd` 从 `int32_t` 换成 `void*`（声明 / 定义 / push-pop 快照字段 / 写入点四处一起；x64 上原来那句 `(int32_t)(intptr_t)r->parent` 把 HWND 高 32 位当场丢掉），#159 那张表里对应行跟着答 `LongPtr` ⇒ 哨兵从此钉住这个宽度（改回 int32_t 就 `ROW-TYPE-MISMATCH` 红）。顺带修掉类型答对后露出的一格：`CStr(句柄)` 落进 `_Generic` 的 `default: vb6_VariantObject` ⇒ 打空串，现在在 CStr 分诊处按 `LongPtr` 走 `vb6_CStrLongLong((int64_t)(intptr_t)(x))`（**不往 `_Generic` 补 `void* :`** —— 那会把真对象引用一律打成数字）。判据 `U-CNT` 三头（容器非零 / 容器≠自己 / CStr 非空串），红侧实测：只改宽度不加 CStr 支路 ⇒ `cnt=` 空、`U-CNT=False`。--emit-c A/B：五工程逐行相同（宽度住在 RTL 里），六工程行数未增删；真跑 ve_units x64+x86 两片绿、charts/flex/flex_x86/czui_x86 无崩。另立 B27（czUI x64 退出码 0xC000041D，A/B 证为预存）|
| 账 #161（提交 `a0b4c67a` = C29-RT-e，门 #306） | **读数推翻本账前提**：一次赋值（`Text` 或 `SelText`）聚焦/不聚焦都**恰好 1 条** `_Change`，连做三次 1/1/1；「只动选区」`chg=0` ⇒ 当年猜的「`_Change` 挂 EN_UPDATE、选区变也发」当场否掉。唯一出 2 的形状是**焦点刚落到这枚控件之后的第一格** —— RichEdit 补发的通知落进下一个泵窗口，被算进相邻那一格（与 #162 同族）⇒ 判据污染，不是产品缺陷。**只改判据**：`RT84` 从 `>= 1` 回到钉死（排空后 `= 1` 且再泵一轮不加发），新增 `RT91`（只动选区 ⇒ Change 0 条，只问 Change，因自发 SelChange 条数天生 2..7 抖）与 `RT92`（一次赋值恰好 1 条）。产品侧一行未动；本地 x64×4 / x86×3 稳定，门里两片三格全 Y |
| 账 #181（提交 `e16e42cd` = C29-CH-f，门 #307） | 崩溃轨迹每进程只记一次：新增零依赖头 `vb6rtl_crash.h` 声明 `vb6_CrashTraceClaim(int slot)`（定义在 vb6rtl.c，4 槽 InterlockedCompareExchange），三个出口（stderr / c3_crash.txt 的 AV 段 / 末段 EXCEPTION）共用这一处口径；**槽位各记一次**而非全局一次（全局一次会让 stderr 那份抢掉文件那份）。为什么必须：记轨迹本身会再触发异常（栈溢出时 fprintf 拿 CRT 锁、栈扫描再读同一批页）⇒ 处理器套处理器 —— 实测 czUI x64 那份 c3_crash.txt 把同一递归栈写了 4 遍、第一现场（gdiplus AV）被压到最后。改后 **6 段 → 2 段且第一段就是 FAULT gdiplus.dll+0xF2E1**。产品行为零改动（三个出口只在 C3_CRASH_TRACE 下装；不带 env 不写文件、退出码不变）。两条工具事实：诊断面的头不该依赖别的 RTL 头（塞进 vb6rtl_runtime.h 会让 vb6forms.c 炸在 vb6_SafeArray1D 未声明）；新增 RTL 头要三处登记（CMakeLists / c3rtl.rc 的 RCDATA 号 / rtl_embedded 的枚举与号→名表），少一处就是 C1083 |
| 账 #182（提交 `0b13dcf6` = C29-CH-g，门 #308） | x64 的崩溃轨迹以前**一条应用帧都没有**，三处各自把路堵死：① `vb6forms.c` 那段栈扫描整段包在「故障地址属于哪个模块」那一问里 ⇒ `call` 跳飞（rip=0x1）时这一问直接失败、一个候选都不打；② `vb6rtl.c` 的栈指针线性扫描只在 `#ifdef _M_IX86` 里编；③ 扫描的 x64 分支用 `wsprintfA("0x%016llX")` —— 用户态 wsprintf **不认 `ll`**，一直打成 `0xlX -> <一串字节>+0xlX`（这段以前从没被执行到，①改完才露出来）。改法：扫描目标一律换成**主 exe 镜像**（生成代码都在 exe 里，扫系统模块没用）、扩到 `_M_X64`（取 `Rsp`、按 `ULONG_PTR` 步长）、偏移改打 32 位 RVA；顺手把不在本模块镜像里的帧照实标 `(outside exe)`（以前把 ntdll 的地址减掉自己的基址打成「rva=0x4376...」，看着像自己的符号），文件侧 AV 头把 `ExceptionInformation[0]==8` 认成 EXECUTE(DEP)（以前一律写 WRITE，与 vb6rtl.c 里 Fix 187 同口径）。读数：同一夹具同一 env，改前只有 5 帧派发链 + 零条扫描行；改后 `st+0 rva=0x2581ae` 落进 `vb6_ComCall`，`st+37/st+40/st+48` 落在工程自己的 `VTableHandle`（`IOleIPAO_EnableModeless` / `GetVTableIPAO` / `ActivateIPAO`）上 —— **账 #183 就靠这几行归的因**；B23/B15 那两条以前归不了因，根因也在这条工具哑火上 |
| 账 #183（提交 `d9e34590` = C29-CH-h，门 #308） | Fix 191 那条「按名 `AddRef`/`Release` 直发槽位」**少绕了一层 vtable**：`((void**)disp)[1]` 读的是 `disp+8`（对象自己的第二个字段），槽位要先从对象首字读出 vtable 再取；同文件里 `vb6_ComIsDispatchable` 自己是两级读法 ⇒ 同一件事两套口径。在 `VTableHandle.bas` 手搭的伪 `IOleInPlaceActiveObject` 上，`VTableIPAODataStruct` 的第二字段恰好是 `RefCount As Long` ⇒ 读出来是 1 ⇒ `call 1`（AV EXECUTE(DEP) target=0x1）。触发条件量到是**确定的、不是偶发**：向 FlexGrid 子窗发**一条 WM_LBUTTONDOWN**（只这一条；`WM_LBUTTONUP` / 右键 / 滚轮都不发）再对主窗发 WM_CLOSE ⇒ x64 `-g` 与 x64 无 `-g` 各 16/16 复现。改成 `lpVtbl->AddRef/Release`（与 `vb6com.c` 里 `vb6_ReleaseObject`/`vb6_ComAddRefDispatch` 同形）后两台各 12 次关窗干净退出、零条崩溃现场；负控（改前那台）同条件 2/2 仍崩。census：`grep -E "\(\(void\s*\*\*\)" src/rtl` = **0** ⇒ 全 RTL 再无手写槽位读法，按 vtable 调一律 `->lpVtbl->`（218 处）。顺带把 `vb6_ComGetProp` 那条 `fallback to IDispatch` 从字体/Extender/宿主三个岔口**之前**挪到之后（它对根本不走 IDispatch 的调用也照打，本轮差点据此把嫌疑引向没执行过的路径）。四条判据面事实进 memory：崩溃后**退出码仍是 0**（判据只能看 `c3_crash.txt` / `[C3_CRASH]`）；`c3_crash.txt` 写在**被测进程当前目录**（相对路径）；`Start-Process -PassThru` 的 `.ExitCode` 在碰过 `.MainWindowHandle` 后拿到 `$null`；**产物架构读 PE 头别看目录名**（本轮一次漏传 `--arch x86`，目录名 `b182gx86n` 的产物其实是 x64，整条「x86 侧读数」当场作废）。x86 那台**起窗就 0xC0000374** 与本刀无关（把 #183 退回单验仍崩）⇒ B23 继续开着 |
| 账 #172（提交 `3e7231c7`，门 #309（run 37165123051，head `3e7231c7`，attempt 1）= 11 job 全绿、非绿 0） | **真相不是「Date 默认值偶尔发垃圾」，是日期字面量 `#...#` 从来没有值**：parser 建 `LiteralExpr` 时只挂原文（`case TokenKind::DateLiteral` 一句 return），发码侧 Date 档照 `node.doubleValue` 打 —— 而构造函数只写了 `intValue(0)`，清的是 4 个字节，8 字节槽的高半从没人写过 ⇒ Ninja/Debug 恰好读到 0.0、VS 生成器/Release 读到 -6.277e+66。**所以改前 Debug 那台也不是对的**（VB6 里 `#1/1/1900#` 是 2.0），只是错得稳定。复现不需要另一台机器：同源码同生成器、只加 `/RTCu` 冷编一台 C3.exe，`--emit-c` 一比就把 4 处垃圾点钉出来（flex 的 ComboCalendar Min/MaxDate 的 ret 赋值 + 各自 `Select Case x To y` 折出的区间边界）；改后两台**逐行零差异**。修法：① 一处出口 `foldDateLiteralToOADate()`（定义 parser_helpers.cpp、声明 parser.hpp）—— 斜杠 M/D/Y、连字符 D-M-Y、带 `H:N[:S]` 与 AM/PM、两位年份 <50→2000s / ≥50→1900s、闰年与真日历校验、OLE epoch 1899-12-30=0.0 用 daysFromCivil 无循环算；② 构造函数改整体清零，把「只清半个联合体」这一类堵住。认不出的形状**照旧留 0、不发新诊断**（宁可不许把现在编得过的工程编红）。判据：新夹具 `tests/test_datelit.bas` 进 bas 队列 —— 新编译器 14 条读数全对（门工件 `test-logs-bas-1/job8/test_datelit.out` 原样可查、`.err` 0 字节）；负控 = 同一份测试喂改前的编译器：10 条变 False 且 `L-serial=-6.27743597849989e+66`。A/B 护栏（BASE=临时回退四份源文件重编、NEW=修复后，同配置同生成器，六工程 --emit-c）：charts/czui/ve_list/iface_wrap/ve_units **0 行变化**，flex 4 行且分类器要求「只有数字变」，新值只有 {2.0, 2958465.0, 2958465.999988426} 三个 OLE 序列 |
| 账 #184（提交 `ab45e2d3`，门 #310（run 37172942961，head `ab45e2d3`，attempt 1）= 11 job 全绿、非绿 0） | **B23 的根因 = `AddressOf` 把本体的裸地址交给了 OS**：x86 上本体是 `__cdecl`（`ret` 不弹参），而 Win32/COM 回调是 `__stdcall` ⇒ 每回调一次把调用方的 ESP 少弹 N*4 字节。VBFlexGrid 每枚网格都经 comctl32 的 `SetWindowSubclass` 装了 6 形参的 SUBCLASSPROC ⇒ 每条消息少弹 24 字节 ⇒ 起窗期堆损坏 `0xC0000374`（改前本地 3/4~5/6 崩、窗口从来出不来；页堆之下现场 6/6 钉在创建 `tooltips_class32` 那一刀，因为 comctl32 正是那条消息链上第一个读到被踩坏的栈的人）。改法按 VB6 口径：被取址的标准模块过程在**定义模块**里另发一枚 `__stdcall` 转发桩（参数表逐字复制 `makeProcSignature` ⇒ 个数/宽度/顺序与本体一致，体内原样转调；Private→`static` 桩，Public→非 static 且原型进自家 `.h`），取址点交桩地址；**本体一个字不改** ⇒ 直接调用那条路与 RTL 那批 cdecl 登记面（Timer / Form_Resize / Winsock / OLE 拖放 / `vb6_di_qsort` 的 cmp）全不牵连。口径只在一处 `addressOfTargetCName`，与委托桩（a15c40b）共用同一套`splitProcSignature` / `thunkArgsFromParams`；标记趟 = stage 3.5c（必须晚于 3.5 跨模块链接才认得归属模块），绑定到 `Delegate` 的 `AddressOf` 一步都不碰。**FlexGridX86 从此挂进门禁**（B23 记的「这份 demo 完全在回归之外」就是它藏这么多轮的原因）。判据四面：`test_addrof_cb` 真跑 x64+x86（六条读数全 Y，CI 两片同值 136 枚窗口，判据写成自洽式 ⇒ 不依赖桌面有几枚窗口）+ `addrof_cb_shape` 发码形状针（两条 `Absent` 就是负控：HEAD~1 的发码里一个 `aoThunk_` 都没有、取址点写的是 `(void*)vb6_CountWin`）+ 静态哨兵 `scripts/check_addressof_thunk_sites.ps1`（同一份脚本在 HEAD~1 的树上 R1~R4 全红）+ A/B `--emit-c` 六工程两架构：**除 flex 外五份逐行相同**，flex `+81/−0` 且 40 处差异全是「取址点改交桩地址」，删除 0 行。另：本轮顺手把 `Test-GuiVbp` 的窗口标题读法从 ANSI 改回 Unicode（`CharSet.Ansi` 一直把类名/标题截成 `V|V`）。遗留：census 那句 `AddressOf callback procs: 49` 数的是**符号副本**（定义模块一份 + 每个引用模块一份），不是过程数（实际 36 枚桩），下轮有别的源码刀时顺手去重，不为它单开一轮门。 |
| 账 #185（提交 `a4e3b574` = C29-PB-a，门 #311（run 37177647595，head `a4e3b574`，attempt 1）= 11 job 全绿、11 片 `FAIL=0`） | **B13 那条「同名属性」的真形状是「装了但没装」**：RTL 建窗时给 STATIC+SS_BITMAP（PictureBox/Image）装自绘子类 `vb6_InstallImageSubclass`，发码为带事件的控件装事件子类 `vb6_InstallControlSubclass`，两层把「下面那层的 wndproc」存在**同一个属性名** `VB6_OrigProc` 上、又都写「属性已存在就不装」⇒ 后装的事件层被整层丢掉。夹具实测：PictureBox 的 Paint/MouseDown/MouseUp/Click 与 Image 的 Click **五条全 0**（x86 与 x64 同形），同一窗体上的 Label（也是 STATIC，样式不含 SS_BITMAP ⇒ 没人抢）一直响 —— 判别力就在这条对比里。**订正 §B 旧 B13 那句「可把 NULL 写进 GWLP_WNDPROC」**：四处还原点（widget.c、picture_prop.c、shape.c:474、forms.c:513）全有 `if (orig)` 守卫，NULL 写不进去；照旧措辞开工会被引向一条不存在的通路。改法两条口径：① 分层槽位（自绘层换 `VB6_ImageOrigProc`，与既有 GBox/GfxBtn/SSTab 同规矩）；② 画的序收成一条 —— 最外层那一臂 BeginPaint/EndPaint 一次，DC 经 `VB6_PaintDC` 交下去让自绘层画表面（BackColor+Picture），再抬用户的 `_Paint` ⇒ 用户笔画在表面之上（VB6 口径）；自绘层已有 DC 就只画、不再第二次 BeginPaint（RTL 那句 trace 新增 `shared=` 一位，实测 `shared=1` = 这条走通）。判据四面：`tests/pbsub`（自驱 Timer 往三枚子窗 PostMessage，两头钉 —— 五条事件针 + PB01 那颗像素必须等于设计期 BackColor 的红）进门禁 x64+x86 两片；**负控 = HEAD 的编译器编同一份夹具**，两边都只剩 PB00/PB06 且 `mdown=0 mup=0 click=0 img=0 lb=1 paint_ok=0`；哨兵 `scripts/check_subclass_slots.ps1` 四条规则（S1 一名一文件 / S2 自绘层不碰裸名 / S3 共享 DC 两头 / S4 向下转调必须在抬处理器之前）同一份脚本在 HEAD 的树上 6 行全红；`--emit-c` A/B 7 工程 × 两架构（BASE 冷编注意 VS 生成器把产物放在 `.build/Debug/C3.exe`）**12 份逐行相同、flex +2/-0 正是那两句新增、删除 0 行、OFFENDERS 0**。顺带三条读数：census 显示 RTL 里拿窗口属性当「已装」旗标的只有那两趟子类化（另三条是重排队列旗标）⇒ 同名碰撞到此为止；frmevents 全套 EV 针齐 + rc=0；FlexGrid x86 起窗 4/4 无崩。 |
| 账 #188（提交 `33cb239a` + 收口 `936124e8` = B29①，门 #313（run 37198483318，head `936124e8`，attempt 1）= 11 job 全绿、11 片 `FAIL=0`；上一轮门 #312 红一条，红因见 B29① 那段的订正） | **C 里的 `12f` 是非法 token**（MSVC C2059 "bad suffix on number"），整数值的 Single 必须写 `12.0f`。这条规则在编译器里散着三份，其中「.frm 设计期字体块」那处（`snprintf("%.4g")` 后直接拼 f）抄漏了补小数点那步 ⇒ 设计期字号写成整数的工程**直接编不过**（实测 Charts 2020 的 ucChartBar 4 条 C2059、ucProgressCircular 16 条）；VB6 默认字号 8.25 带小数点 ⇒ 门禁里所有夹具都恰好躲过，这条潜伏了不知多少轮。同一个形状在语义层 Fix 133z 早被修过一次（`Optional ... As Single = 1!`），当时只在那一处补的小数点 ⇒ 规则没收口。**改**：新增 `src/common/float_literal.hpp` 一处权威（`floatingLiteralText` / `floatSingleLiteral` / `floatFixed6Literal`，经典 locale + 「没有 . 或 eE 就补 .0」），五处发射全接过去；位数各按原样（字体块 4 位、to_string 那三处保持 6 位固定形状 ⇒ 发码字节不动）。**判据**：`fontsize_literal` 发码形状针钉在**真实工程** ucChartBar 上（Needles `->Size = 12.0f;` + 证人 `->Size = 8.25f;`，Absent 是改前的 `->Size = 12f;`）；哨兵 `scripts/check_float_literal_shape.ps1`（R1 手拼后缀=0 / R2 三个函数各定义一次且带着补点与经典 locale / R3 调用点>=4）在缺这刀的树上 R1~R3 全红、R1 当场点出那 6 个手拼点。**配对读数**（同一份工程 × 两台编译器 × x86）：ucChartBar C2059 4→0（只剩 B29②的一条 C2065）、ucProgressCircular 16→0（剩 B29③）。**护栏**：A/B 从门禁形状断言面取清单（37 文件 × 两架构）= 只有 ucChartBar 两片被改，归一化浮点 token 后逐行相等、数值多重集相等、增删 0 行、OFFENDERS 0。 |
| 账 #189（提交 `ea5155af` = B29②，门 #314（run 37200966607，head `0efb29b0`，attempt 1）= 11 job 全绿、10 片 `FAIL=0`） | **VB6 里 `arr.Count` 问的是数组本身，不是某一枚控件的属性** —— 发码侧以前只有 `LBound`/`UBound` 在成员读取那一路各写了一条 if，`Count` 漏了 ⇒ 掉进 COM 兜底，发成 `vb6_ComGetIntProp(vb6_hwnd_<数组名>, L"Count")`。两种红法各一条实测：**跨窗体**引用（Charts 2020/ucChartBar 的 Form2 引用 Form1 那枚 static）= `error C2065 未声明的标识符`；**同窗体** = 编得过、恒答 0（`For i = 2 To arr.Count - 1` 一格也不走，这类是静默错值那一族）。与账 #157 同族：句柄类表达式必须走 `vb6_arr_*`，不许凭空拼 `vb6_hwnd_`。**改**：新增唯一出口 `CCodeGen::ctrlArrayMetaMemberExpr()`（`src/backend/cgen_util_ctrl.cpp`，声明在 `detail/util/cgen_helpers.inc`），成员读取那一路把它放在**问控件属性之前**（`cgen_expr_member_form_builtin.inc`）—— 这三条永远不是控件属性，顺序本身就是口径；旧代码那两条 if 还额外挂着 `readFn.empty()` 当前提，等于把「属性表里恰好没有 LBound」当成了条件，一并撤掉。**判据三面**：①`tests/ve_units` 塞一枚 UC 控件数组 `uArr(0..2)`，两头钉 —— `U-ARR-RAW count=3 lb=0 ub=2`（三个数各对上）+ `U-ARR=True`（内含「拿 Count 当上界**真的**圈了三圈」，只钉前头那条的话 循环不走也绿）；**负控 = `a4e3b574` 的编译器编同一份夹具** ⇒ `frmUnits.c(214)/(220): error C2065 "vb6_hwnd_uArr"`、BUILD rc=1。②发码形状针 `ctrlarr_member`（Needles 三枚 `vb6_CtrlArr_*(&vb6_arr_uArr)`，Absent = 改前那条形；两台编译器实测 2/0 ⇒ 这枚针真能红）。③哨兵 `scripts/check_ctrl_array_members.ps1`（A1 三个 RTL 出口只许出现在权威那一个文件里 / A2 权威三条分支齐且留着空串出口 / A3 `memLower == "lbound"` 那一形归零 + 权威至少一个调用点），假 needle 验红做过：临时文件里多写一处 `vb6_CtrlArr_GetCount(` ⇒ A1 当场点名，删掉回绿。**护栏**：A/B 仍从门禁形状断言面取清单（38 工程 × 两架构 = 76 次发射；BASE = `a4e3b574`，所以 #188 那 4 对 `12f`→`12.0f` 也在账上、已逐行归因）= 只有 ucChartBar 与 ve_units 四片被改、行数一字不差、增=删、其余逐行相同、OFFENDERS 0。**顺带两条新账**（都在 ucChartBar 真编译通下去之后才现形，见 §B29 ⑤⑥）：**⑤** 那两枚控件的事件臂函数名按控件设计期拼写现拼、过程定义按 Sub 自己的拼名 ⇒ 大小写不一致就是两个符号，链接期 `LNK2019` ×2；**⑥** 控件数组**元素的成员访问**（`uArr(i).SW()` 返空、`uArr(i).Left` 直接崩）另是一条独立通路。 |
| 账 #190（提交 `a27758c3` = B29⑤，门 #315（run 37202837330，head `a159b85d`，attempt 1）= 11 job 全绿、10 片 `FAIL=0`） | **事件臂调用的函数名以前是按「控件的设计期拼写」现拼的，而过程定义发的是 Sub 自己的拼写** —— VB6 标识符大小写不敏感、C 敏感，于是"改了控件名没改过程名"（VB6 完全合法，两枚照样配一对）在产物里就是**引用一个没人定义的函数**：链接期 `LNK2019` + `fatal LNK1120`。这条红**只在链接期现形**：存在性那一步 `symTab_.lookup(控件名 + "_Click")` 本来大小写无关、命中了，语义层与 `--emit-c` 的形状都看不出坏了 ⇒ 缺的只是"命中之后按谁的名字发"。实测 ucChartBar 的 Form1 恰好两枚（控件 `ChkAxisY` / 过程 `ChkAxisy_Click`:435，控件 `cboLabelsPositions` / 过程 `CboLabelsPositions_Click`:439）⇒ 正好 2 个外部符号。**改**：新增唯一出口 `CCodeGen::eventHandlerFn(ctrlName, suffix)`（存在性与名字出自**同一次** lookup，命中交 `cProcName(sym->name)`，没命中交空串让调用方**不装这条臂** —— "没处理器就不装"从此是结构性事实，不是每条 if 各自记得查）。43 处现拼全改读它：create.inc 18 / subclass.inc 16 / dispatch.inc 6 / cgen_form_menu.cpp 2 / ctrl_style_apply.inc 1。AccessLevel 不用跟着改：`extern void f();` 之后再声明同名 static 是合法 C（早先那条 static 胜出）⇒ 只有**拼写**要紧。**普查教训（比修本身值钱）**：第一遍只按 `ctrl.controlName` / `ctrlName` 两个变量名扫 ⇒ **少了 18 处**（`info.ctrlName` 那一族 = 焦点/鼠标/Validate 一整批 + 菜单两枚），是第二、三趟按剩余量逼出来的。⇒ census 要按**形状**写（`cProcName(<任意> + "_`），别枚举变量名。**判据三面**：①新夹具 `tests/evtcase`（x64+x86 两片）摆三枚控件 —— `cmdRun`/`CmdRun_Click`、`chkOpt`/`ChkOpt_Click` 只差大小写，再加拼写一致的 `cmdSame`/`cmdSame_Click` 当**证人**（排除"BM_CLICK 自己没驱动起来"那种假红，口径同账 #185 那枚 Label）；驱动 = Timer 第一拍对三枚发 `BM_CLICK`、第二拍打 `EC-CNT run=1 chk=1 same=1` 自退。**负控 = `a27758c3` 那台编译器真编同一份夹具**：`CaseForm.obj : error LNK2019 无法解析的外部符号 _vb6_cmdRun_Click / _vb6_chkOpt_Click` ⇒ **恰好两枚、证人那枚没有第三条**（物证 `.build/b190pre/c3-error.log`）；修复后 x64 与 x86 都是 `run=1 chk=1 same=1`。②真工程配对读数：ucChartBar x86 的 LNK2019 **2 → 0**，并且**这个工程第一次产出 exe**（590,336 字节，`.build/b190chart/Proyecto1.exe`）⇒ B29 那张表里它从此归入"编得过"，#187 的升格候选多一件。③哨兵 `scripts/check_event_handler_names.ps1`（E1 现拼=0 / E2 权威必须"从符号取名"且留着空串出口 / E3 调用点>=40）；假 needle 验红做过 —— 临时塞一枚现拼 ⇒ E1 当场点出文件与行号，删掉回绿。**护栏**：A/B 沿用 #188 那份从门禁形状断言面推出来的清单 + evtcase/pbsub（39 工程 × 两架构 = 78 次发射，两台编译器之间只差这一刀）⇒ 只有 4 份产物被改（ucChartBar 两片、evtcase 两片）各 2 行，每行都验过"整行只换一个标识符"（`old.replace(旧名,新名) == new`），其余 74 份逐字节相同、OFFENDERS 0。 |
| 账 #186（提交 `4d1dbf45`，门 #316（run 37205289964，head `4d1dbf45`，attempt 1）= 11 job 全绿、10 片 `FAIL=0`） | **`Erase m_tvFiles(lIndex).bvData` 这一形以前在 parser 就被拒**（VB2001 "Erase 不支持带下标的形式"），而它是 VB6 的合法写法、真工程在用（`tests/Charts 2020/ucTreeMaps/FontMemRes/PropPagFMR.pag:720` —— 报的行号 594 是折掉 126 行设计期头块之后的，与 B29④ 那条对得上）⇒ 整个工程卡在最前面。**改**：`EraseStmt` 加 `targets`（与 `ReDimStmt.targetExpr` 同一套机制，Fix 100），`varNames` 仍存**去下标的点链名** ⇒ Variant 成员那一问（`isVariantArrayTarget`，Fix 090p）继续只有一处判据；点链→表达式树收成 `Parser::buildDottedNameExpr`，ReDim 那趟改为转调它（**别再抄第二份**）；发码侧 `emitExpr` 出 `VB6_SA_AT(vb6_type_TElem, m_items, 1).bvData` ⇒ 销毁它并置 NULL。1D/ND 那一问不需要新表 —— UDT 成员的动态数组**恒为一维**（结构体发码单点 `cgen_decl.cpp:506`）。**边界刻意守住**：`Erase arr(i)`（无成员）继续报诊断 —— VB6 只在元素是 Variant 时允许，本仓没那条通路，静默降级成"销毁整个数组"比编不过更坏；配一条负例 `erase_neg_indexed` 钉住"不许过头"。**自己撞出来的一条（值钱）**：空括号的次序是**先吃完整条点链、再吃那组括号**（Fix 082 的另一半）。第一版我在基名后就直接判括号，于是 `Erase m_bag.MaxWidths()` / `Erase .MaxWidths()`（VBFlexGrid 6590 就是这形）残留一个裸 `(` ⇒ VB2003/VB2002。**是 A/B 逐字节比抓到的**：78 份快照里 flex 两片从 5.3 MB 掉成 1.3 KB ⇒ 快照**体量**差比"差异行数"更早报警 —— 行数那种判据对"整片塌掉"反而钝。**判据四面**（`tests/erase_sub` 两片 + 一负例 + 一形状针）：①`EA01-KEEP s0=30 s2=50`（只擦那一格，拦"销毁整个数组"）；②`EA02-RECYCLE s1=20 len=1`（那一格可再分配并用 —— ReDim 自己会先销毁旧数组 ⇒ 没置 NULL 就是**双释放**、现场直接崩；反过来没真销毁会读到残留的 40）；x64 与 x86 都 `keep_ok=True recycle_ok=True`。**负控 = `4d1dbf45` 那台编同一份夹具** ⇒ VB2001 + VB2003、BUILD rc=1、一条读数都不出现。③`erase_neg_indexed`（`Test-CompileFail`，针取消息里的**英文片段** —— GBK 控制台下中文会被折行/转码，既有那批负例也是这个口径）；④`erase_member_shape`（`Test-CodegenNote`）钉住"销毁 + 置 NULL"那一整行，Absent 是退回整数组那一形。**护栏**：A/B = 40 工程 × 两架构（80 次发射，`a159b85d` 那台 vs 现在这台）⇒ 生成的 C **逐字节相同 78/78**（新夹具两个旧快照里没有，不算）。**顺带把下一格量出来了**：ucTreeMaps 现在零诊断过 parser、真编译推进到 cl 才红（10×C2039 / 2×C2059 / 1×C2198）⇒ 见 B29⑦ 与新账 #192。 |
| 账 #191（提交 `f5b4a03b` = B29⑥，门 #317（run 37208673511，head ea3eeeec，attempt 1）= 11 job 全绿、10 片 FAIL=0） | **工程内 UserControl 的数组元素访问以前不走直发那条出口** —— Fix 112 只给"单枚控件"接了 `vb6_UC_InstanceOf` 那一层，数组元素另有一条值上下文通路（`cgen_expr_member_precheck.inc`）不认识工程内的 .ctl => 整条落进 `vb6_ComCall` 兜底，而宿主值不是 IDispatch ⇒ 交回来是**空值**（编得过、跑得起来、就是静默不干活；探针 `b189probe.out` 实测 `uArr(0..2).SW()` 三枚全空）。**改**：那段 40 行从发码点抽成唯一出口 `emitUcInstanceMemberExpr`，两条路各交自己的宿主串（单枚 `vb6_hwnd_<名>` / 元素 `vb6_CtrlArr_GetAt(&vb6_arr_<名>, i)`），顺序放在控件属性表与 COM 兜底之前。**连带第二条（左值侧，独立缺陷）**：改道之后 `ucChartBar1(i).Font.Size = ...` 落进 `tryRewriteCOMLvalue` 的字符串级折叠，它只取 `prop_get_` 那一段、把尾巴 `.Size` **整段丢掉** => `prop_let_Font(elem, <double>)`（值槽是 `vb6_ComIface_Font*`）= Form2.c 四条 C2440，工程从"编得过"退回编不过；修 = 折叠之前加一条"带尾巴的 prop_get_ 目标"落点 —— 取回那枚对象再写一层（StdFont 本身是真 IDispatch），认不得的尾巴 `return false` 让它当场红。**判据三面**：真跑证人 `U-ARRM-methods=3`（x64+x86）+ 发码针 `ucarr_member_call`与 `ucobj_chain_write`（后者钉真实工程）+ 哨兵 U1~U4（假 needle 验红：改 `(Pattern C-tail)` 一名 => U4 点名）。**真编译配对读数（x86）**：ucChartBar rc=0 且出 exe；ucProgressCircular 错误类严格是基线子集（C2059 归 #188，剩 B29③④ 两条）。**护栏**：A/B 80 份 => 只有 6 份被改、逐行归因、OFFENDERS 0。**留给下一轮**：①extender 属性（`uArr(i).Left` 仍 AV）= 新账 **#193**；②"单枚那条链"没有夹具（整面 80 份快照里带尾巴的写只有元素这一形）；③任何把访问从兜底挪到直发的刀，都要连带问一句"左值侧谁在按字符串改写它" —— 基线上这条链根本到不了 Pattern C/D2。 |
| 账 #194（提交 `40bbef67` = B29⑦ 的第二条，门 #318（run 37211897524，head 9bb2d08b，attempt 1）= 11 job 全绿；能下到的 3 片日志各自 FAIL=0，剩下 8 片与工件zip被本机网络代理拦成 HTML 页（预签名 URL 全拿不到），所以 optdef 两片的 CI 逐条 PASS 行**没有物证** —— 由 job 结论 + 与 evtcase 同款接线 + 本地两台真跑推定，下一轮补读） | **VB 的整数类型后缀是词法记号，不该出现在生成 C 里** —— 发码侧 `visit(LiteralExpr)` 本来就按**数值**重打（Integer 无后缀 / Long 带 L / 超 32 位带 LL），但语义层那份 `evalOptionalDefault` 是**另写的一份**：Integer·Long 两支直接 `return lit->rawText` ⇒ `Optional ByVal FontIndex As Long = 0&` 发成 `if (!_has_FontIndex) (*FontIndex) = 0&;` = C2059 "bad suffix on number"（真工程物证：ucTreeMaps 的 PropPagFMR.pag:740/886 两条，`.build/b192new/ucTreeMaps/c3-error.log`）。**改**：新增 `src/common/int_literal.hpp` 一处权威 `intLiteralText(数值, 是否 Long 档)`（LL 那一步的 32 位界限判定照搬：MSVC 的 long 是 32 位，`2147483648L` 会退成 unsigned long）；发码侧两支、语义侧两支、`cgen_util_ctrl.cpp` 里 Slider 设计期那两处手拼 `+ "L"` 全改读它（Slider 那两处换完**字节不变**）。**判据三面**：①新夹具 `tests/optdef` 两片两头钉 —— `OD-VAL=True`（四枚默认值 12/3/16/0 各按声明落地，十六进制那枚顺带证明"按数值重打"没改错数）+ `OD-EXPL=True`（显式实参照样赢；只钉前头那条的话"恒取默认值"也绿）；**负控 = `ea3eeeec` 那台真编同一份夹具** ⇒ OptForm.c 116/118/119 三条 C2059 + C2065 "H10"、rc=1、一条读数都不出。②发码形状针 `optdef_default_shape`（Absent 是抄源码那三形）与 `optdef_default_shape_real`（钉在**真工程** ucTreeMaps 那一行上 —— 该工程今天还因 #192 红着，但发码面已过，所以这枚针现在就能红能绿）。③哨兵 `scripts/check_int_literal_shape.ps1`（I1 手拼整数后缀=0 / I2 权威一次定义且留着 LL 界限与无后缀支路 / I3 调用点>=3 且两侧各有 / I4 语义侧不许再 `return lit->rawText`）；假 needle 验红：把语义侧那一支改回 rawText ⇒ I4 当场点名，还原回绿。**真编译配对读数（x86, ucTreeMaps）**：C2059 2→0，C2039 10 与 C2198 1 原样（属 #192）。**护栏**：A/B = 40 工程 x 两架构 = 80 份（BASE = `ea3eeeec`）⇒ **被改的产物 0 份**、OFFENDERS 0 —— 这条读数本身是信息：整面存量里没有一处带后缀的 Optional 默认值 ⇒ 它只在真工程里响，**夹具必须自己带一份**，否则规则又会被"全绿"掩盖（与账 #188 的 `12f` 同一课）。**顺带量出的独立缺陷另立 §B30（账 #195）**：`3%` 在词法层就被拒。 |
| 账 #195（提交 `963c5e71` = §B30，门 #319（run 37215210897，head a7524ab7，attempt 1）= 11 job 全绿；可读到的 4 片各自 FAIL=0，里面 `[SYNTAX] int_suffix_forms` / `[SYNTAX-FAIL] int_suffix_neg_range` / `int_suffix_shape` 三条都 PASS，而且顺带把 #318 欠的那条物证补上了 —— `[VBP] optdef` 与 `optdef_x86` 两片在 CI 上真跑真 PASS。静态哨兵所在那一片的日志仍被本机网络代理拦成 HTML 页 ⇒ `int_suffix_sites` / `int_literal_shape` 的 CI 行还是没有物证（本地两台全绿 + 假 needle 验红做过）） | **VB 的 Integer 后缀 % 在词法层就被拒** —— 类型后缀的消费在词法器里**抄了四份**（十进制 + &H + &O + &B）外加一张 radixDigits 剥离表，而 `case %:` 那一支只吃字符不置标志（`&`/`!`/`#`/`^` 都置），于是 `3%` 落回"无后缀十进制按数值大小定档"那一段，parseIntLit 看见残留的 % 就报"十进制数字超出 64 位整数表示范围"，一条合法语句级联 9-13 条诊断、整个工程卡在最前面。**读数**：七形（十进制 / 三种进制 / 负数 / 表达式 / 参数默认值）修前全红，`12&` 与 `&H10^` 与无后缀三形一直通 = 证人；上一轮那台（ea3eeeec）读数完全相同 ⇒ 存量，与 #194 那条"把 0& 抄进生成 C"不是一件事（那条过了 parser、这条压根过不去）。**改**：四处一起认下 % 并置 isInteger，十进制那支发 IntegerLiteral（按数值），剥离表补 %；显式 % 超出 -32768..32767 **报词法错**而不是按 int32 收下再让 int16 截（那是静默把 2147483648 变成 -2147483648）。**判据四面**：正例夹具 `tests/intsuffix`（Test-SyntaxMulti）+ 边界负例 `tests/intsuffix_neg`（针取 ASCII 片段 "(-32768..32767)")+ 形状针 `int_suffix_shape`（三种进制各按数值落地 `vb6_ChkInt(255)/(15)/(5)`，参数默认值那形顺带钉住 #194 的出口）+ 哨兵 `check_int_suffix_sites.ps1`（假 needle 验红：改回空支 + 剥离表去 % ⇒ 四条一起点名）。另把 #194 的 optdef 夹具先前为绕开本缺陷改成的无后缀两处恢复原样，x64+x86 真跑到 OD-DONE。**护栏**：A/B 80 份 ⇒ 被改的产物 0 份（存量里一处 % 都没有 ⇒ 又只能靠夹具响，与 #188/#194 同一课）。 |
| 账 #197（提交 `9565bdbb` = §B32，门 #322（run 37235573381，head ada93533，attempt 1）= 11 job 全绿、11 片日志**全部下到**（这轮代理没拦）：10 片各自 FAIL=0（第 11 片是 Build C3.exe，不打这个计数），本轮五条新用例在 CI 上逐行真 PASS —— `[VBP] scalemode` / `scalemode_x86` / `[CODEGEN-NOTE] scalemode_design_write` / [CODEGEN-NOTE] scalemode_read_real` / `[STATIC] scalemode_writers`） | **`VB6_ScaleMode` 这个窗口属性全仓 0 个写者** —— 读的一侧早就齐了（`vb6_GetScaleMode` / `vb6_WindowScaleModeSelf` / `vb6_ContainerScaleMode`，缺省 1=缇），写的一侧 `vb6_SetScaleMode` 一个调用点都没有 ⇒ 谁问 ScaleMode 都答缺省。这一档决定的是**量出来的数对不对**（ScaleWidth/ScaleHeight、控件几何 #175、文字量纲 #177 全按它折算），不是某一枚控件的外观，所以症状是"处处差 15 倍"而不是"某处坏"。**改**：一处权威 `emitDesignerScaleModeProp` × 三条落点（顶层创建路 / 容器子控件路 / 窗体 WM_CREATE 的 P20-40 块）+ 读写两张表成对补 `scalemode`（读给**同一个** `vb6_WindowScaleModeSelf`，程序读到的数与换算用的数不可能再分家）。RTL 一字节未动 ⇒ 不碰 `c3rtl.rc`，绕开 #156 那条旧资源坑。普查 93 份设计块 19 处（UC 10 已通 / PictureBox 6 / Form 3）⇒ 口径「设计块写了就发」，不按值筛也不开控件型白名单。**判据四面**：夹具 `tests/scalemode` 两片（存设计值 + 问窗口的**比值**判据不钉绝对数 + 运行期切档逐数相等 + 另一枚没被写坏 + Frame 里那枚钉第二条创建路）；负控=改前那台真编 ⇒ SmForm.c 166/182 两条 C2039、rc=1；形状针两条（夹具 + 真工程 ucTreeMaps）在改前那台当场红；哨兵 W1~W5（假 needle：注释掉容器那条落点 ⇒ W3 点名 exit 1；**工具教训**：植针要按人会怎么改坏植——先前改成 `XX` 前缀，`Contains` 当子串照样命中，等于没验）。**护栏**：A/B 82 份 ⇒ 12 份被改、OFFENDERS 0（10 份各 1 行新增 SetScaleMode；2 份 = VBFlexGridDemo 各 2 行，`Me.ScaleMode` 从 `vb6_ComGetProp` 兜底换成真读数）。**真工程配对**：ucTreeMaps x86 C2039 **4→2**；czUI-main（唯一 Form 级声明 3=Pixel 的工程）真编真起窗。**顺带新账 #199**：`.pag` 的设计块控件从来没被创建（PropPagFMR 段 CreateControl 0 处 vs 同工程 Form1 段 14 处，页里的 `vb6_hwnd_Picture1` 是按需新建、初值 NULL 的槽位）⇒ 属性页里 `With Picture1` 全打在空句柄上，#196 做完 UnicodePrint 也还是白画，控件线顺序改成 #199 → #196。 |
| 账 #196（提交 `50676a8b` = §B31 的 hDC 那一半，门 #323（run 37238929955，head f26f9590，attempt 1）= 11 job 全绿、10 片各自 FAIL=0（smoke / compile 两片这轮重下到了），本轮五条新用例逐行真 PASS：`[VBP] dcsurf` 与 `dcsurf_x86`、`[CODEGEN-NOTE] dcsurf_dc_shape` 与 `dc_read_real`、`[STATIC] control_dc`） | **「这枚控件的绘图 DC 从哪儿来」在 RTL 里早就有一处口径，但句柄交不回 VB 代码** —— Print/Cls 走的那条（派发期用外层 BeginPaint 挂上的 `VB6_PaintDC`，否则回落 `GetDC`）一直是 static 且只在内圈用，`.hDC` 没有出口 ⇒ 真工程那一形 `With Picture1 : TextOut .hDC, ...`（Charts 2020/ucTreeMaps 的 PropPagFMR.pag:258）只能撞 `cgen_expr_with.cpp` 那条 "hwnd.成员" 兜底 = C2039。**改**：把那条口径抽成 `vb6_ControlDrawDC`（三头共用：Cls / Print / 新出口），新出口 `vb6_GetControlHDC(void* hwnd)` 按 **VB6 的"一个对象一张"** 语义把派发期那张直接交出（不缓存不释放）、否则把回落那张按 HWND 缓存进 `VB6_ObjectDC` 并归还白拿的那张，归还点在 PictureBox/Image 那层自己的 `WM_DESTROY`（#185 那套分层槽位）；后端读写两张表各补一行 `hdc`，**只登记 PictureBox 与 Form**（通用行 = `List1.hDC` 也答一个数 = 伪造成功，同 #192）。RTL 动过 ⇒ touch `c3rtl.rc` 重编 C3.exe。**判据四面**：新夹具 `tests/dcsurf` 两片四头钉（DS01 反复读+With 三个数相等非零 / DS02 两枚互不相等 / DS03 `GetDeviceCaps>0` 证明是一张**活的 DC** / DS04 `GetPixel` 各自等于**自己那枚**的设计期底色；DS05 只钉前缀，dpi 不钉绝对值）；**负控 = 改前那台真编同一份夹具** ⇒ DcForm.c(167) C2039 "hDC" 不是 "HWND__" 的成员、rc=1；形状针两条（夹具 + 真工程，Absent 连晚绑定那一形 `vb6_ComGetLongPtrProp(X, L"hDC")` 一起钉）；哨兵 `check_control_dc.ps1` D1~D5（五条假 needle 按"人会怎么改坏"植：自己抢 DC / 权威不认 PaintDC / 不归还 / 销毁不撤名 / 表给成通用行 ⇒ 逐条能红，还原回绿）。**护栏**：A/B 86 份 ⇒ 被改 8 份、**OFFENDERS 0**，14 对 K2 + 5 条 WD（`VB4001 Unknown control property '.'hDC'` 那条诊断消失是这一刀的目的）逐行归因。**真工程配对**：ucTreeMaps x86 C2039 **2→1**（只剩 TextHeight；TextOutW 的 C2198 是它的级联，一起没了）。**顺带一条实测订正**：Charts2020 `Proyecto1`/`ucProgressCircular` 的 `Picture2.hDC` 以前走晚绑定 `vb6_ComGetProp`，探针（`.build/b196probe`，改前那台真编真跑裸形）量到 `HDC=671159075 / DPI=96` ⇒ 那是一张**活的** DC，不是"恒答 0"那一族（`uc_hostmodel_getprop.inc:86` 本来就认 hDC）；本刀在裸形上改的是**语义**（每读一次新取一张、从不归还 → VB6 的一个对象一张 + 销毁归还），真红的那一半是 With 形（C2039）。**留下的口径**：Form 自己那张缓存 DC 暂无归还点（语料 `Form.hDC` 现存 0 处、进程结束由系统收回，记下不装红）；本账欠按 HWND 的 `TextHeight`/`TextWidth` 与带 HWND 的 `ScaleX`/`ScaleY` 两条，顺序 #199 → #196。 |
| 账 #196 第二条（提交 `060eddca` + 针面订正 `ac3df329` = §B31 的 TextHeight/TextWidth，门 #324（run 37242648474，head 6db5203b）= **failure，逐 job 归因：产品侧一片绿**（Build success、其余 9 片各自 FAIL=0、dcsurf / dcsurf_x86 / dcsurf_dc_shape / dc_read_real / control_dc 五条真 PASS），红的只有 Tests (syntax) 里我新加的两条 [CODEGEN-NOTE]：`missing: vb6_ControlTextHeight((void*)_vb6_with_0` —— 我本地拿**前缀** `_vb6_with_` 验过就登记成 `_vb6_with_0`（那个 Sub 里两个 With picA：hDC 那条是 _vb6_with_0、文字量那条是 _vb6_with_1），另一条又在结尾多写一个 `)`（产物那位置是逗号）⇒ 登记的根本不是产物里的串；按 emit 逐字重验后改针（两台 present 1/2/1 与 1/1、absent 全 0），产品一字节未动。门 #325（run 37243854108，head ac3df329，attempt 1）= 11 job 全绿、10 片各自 FAIL=0，七条相关用例逐行真 PASS） | **按 HWND 的文字量缺半个出口** —— `With Picture1 : .CurrentY + .TextHeight(Text)`（ucTreeMaps 的 PropPagFMR.pag:265）是 #196 收完 hDC 之后该工程仅剩的那条 C2039。**改**：不发明新机制 —— 现有 `controlZeroArgMethod` 那套「表交名字、码头拼实参」本来就支持带实参（With 形交裸名 + `pendingChainObj_`，调用点把句柄**前置**进实参表），新表 `controlOneArgMethod` 只登记 textheight/textwidth × Form/PictureBox 两档（**不给通用行**，同 #192/#196 那条口径），两条码头各查一次；实参签名表补两行 `{"void*","BSTR"}`（漏了就是把 vb6_VARIANT 裸喂给 GetTextExtentPoint32W = Fix 113）。RTL 新增 `vb6_ControlMeasureTextPx`（DC 走**同一处** `vb6_ControlDrawDC`、字体走 `WM_GETFONT` = 与 Print 同源）+ `float vb6_ControlTextWidth/Height`，**单位过 `vb6_ScalePxToUser` + `vb6_WindowScaleModeSelf`**（#175/#197 那一份表）。判据：夹具加长三头 TH01 两形逐数相等非零 / TH02 缇框 vs 像素框 **>4 倍**（实测 16 vs 240=15 倍，比值判据不钉绝对数）/ TH04 宽度随文字变；负控 = 改前那台 rc=1 无 exe；哨兵 D6~D9 五条假 needle 逐条能红。**顺手抓到哨兵一条假绿**：D9 第一版把两个函数当一段正则抓，改坏宽度那半时高度里的换算还在 ⇒ 整段照过；改成"一个一个函数各自取身体"才红（与"注释行不算""计数要打印实测值"同一类自欺）。**A/B** 86 份 ⇒ 被改 8 份、OFFENDERS 0（32 对读法替换 + 8 条 WD；配对正确性要用**整行**判 —— 公共前缀会把 `Co` 折掉，残段里搜全名搜不到）。真工程配对：ucTreeMaps 两台 C2039/C2440 **全清**，红点推进到 **LNK2019 ×3**（三枚 GDI+/字体内存桩没登记 ⇒ 新账 #201）；`VB4001 Unknown control property` 只剩 1 条。**没接的**：不带括号的语句形（语料 0 处）；窗口字体那条判据今天当不了判据（改了字号量回来不动、`FontPixelHeight` 那句整行不打 ⇒ 新账 #200，夹具里只留读数 TH06）。 |
| 账 #201（提交 `4ced55cc` + 记账 `871a627f` = §B36 的三枚 DI 桩，门 #326（run 37247978838，head 871a627f，attempt 1）= 11 job 全绿、10 片各自 FAIL=0；CI 自己量的两条新用例真 PASS —— `[VBP-BUILD] charts_ucTreeMaps ... PASS (679,424 bytes)` 与 `charts_ucTreeMaps_x86 ... PASS (585,728 bytes)`，顺带 `[STATIC] di_stubs_census`、`[VBP] dcsurf`、`[STATIC] control_dc` 三行还在原样绿） | ucTreeMaps 编译面全清之后卡在链接期三条 LNK2019。**改**：三枚桩手写进 `src/rtl/core/di/vb6_di_stubs.c`（就是"手写那一档"，gen_di_stubs 不动它），`check_di_stubs.ps1 -Update` 把基线从 620 冻到 624（少一个就红，正是这道哨兵存在的理由）。读数说清楚为什么生成器没发它们：`AddFontMemResourceEx` 那一形第三参在真源码里写的是 `ByRef DESIGNVECTOR`（无 `As` ⇒ Variant），落在生成器"形状不定就跳过"那一档；GDI+ 那两枚名字里带 Font 会被路由到 text 族，但当时那次会话没引用到 ⇒ 定义集里就没有。**一个刻意的取舍**：`AddFontMemResourceEx` 的 pdv 交 **NULL** 而不是把那个 `vb6_VARIANT*` 转手给 GDI —— VB6 那句 `AddFontMemResourceEx(.bvData(0), n, 0&, cnt)` 里 `0&` 的意思是"没有 design vector"，而 Win32 对这一意的拼法就是 NULL；把 variant 的地址当 `PDWORD` 交进去，GDI 读到的是 VARIANT 头两个字，那既不是 NULL 也不是 design vector，是伪造。gdiplus 那两枚走 `LoadLibrary`+`GetProcAddress`（`gdiplus.h` 是 C++ only，且本仓的 Declare 一律不要求导入库），解析不到回 `GpStatus` 的 2，与 gdiplus 族生成桩同形；GDI+ 的 `GdiplusStartup` 桩早就有（工程自己调）。**结果**：ucTreeMaps 两台（x64/x86）**第一次真编真链出 exe** = 659,968 / 562,688 字节，诊断面 **0 error**、116 条 VB 警告（`VB3001 未声明的标识符: Ambient / Extender / Is / PropertyPage` 那一族 = 属性页与 extender 面还欠着，见 §B34/#199 与 #193）。**护栏**：这一刀只动 RTL 与清单 ⇒ 发码面该一字不变，A/B 86 份（BASE = 改前那台的 emit 快照）**changed=0、offenders=0**。**升格**：`Test-VbpBuild charts_ucTreeMaps` + `_x86` 两条进清单（#187 那条口径的正解 —— 列进来这一次就是"从此不许退回编不过"），run_tests.ps1 里那条"刻意不列"的注释同步改成只剩 ucProgressCircular。**下一步**（这一格解开的）：把工程真跑起来的观测面 —— 属性页那条路（#199 的甲/乙）与 extender/`VB3001` 那一族；以及 #200 那条窗口字体的归因（文字量唯一还没被运行期验证的一头）。 |
| 账 #187（提交 `a51e84aa` = B14 后半，门 #321（run 37230353501，head 59811f14，attempt 1）= 11 job 全绿、每片 FAIL=0） | **门禁里以前没有"真工程必须编得过"这一格正面积** —— 三个 helper 各管两头：`Test-Vbp` 要跑起来校验输出（第三方真工程没有自退出口 ⇒ 进不去），`Test-VbpBuildFail` 只钉"必须红"，中间那格"必须绿"没人管。后果这两轮反复撞到：#188 的 `12f`、#190 的 LNK2019、#191 的数组元素、#192 的 With 光标全落在 Charts 2020 的 UC 子工程上，而它们在清单里**压根不存在**，"编不过"在门禁里连红都算不上。**改**：新增 `Test-VbpBuild`（真编译真链接 + 断本轮新出的 exe 在，不跑），三枚 UC 子工程 × 两架构六条进清单（ucChartBar / ucChartArea / ucPieChart；实测 x64 696,320 / 576,000 / 576,000，x86 590,336 / 496,128 / 498,176 字节）。两处刻意的写法：①**输出目录按用例隔离** —— 这四份 .vbp 的 ExeName32 全写 `Proyecto1.exe`，共用 $OutDir 会互相盖掉（#166 那条的姊妹坑：名字修对了还会串味）；②**先删干净再编、rc 才是判据** —— 只 `Test-Path` 会命中上一轮的旧 exe。ucProgressCircular / ucTreeMaps 刻意不列（今天还红着：B29③④ 与 #192/#196），列进来就是把已知的红当基线，等编过之后再加那一次才算"不许退回编不过"。**能红性**：同一份 helper 指 ucTreeMaps ⇒ `[VBP-BUILD] FAIL rc=1 exe=False`、计数器 total=3 pass=2 fail=1；指三枚绿的 ⇒ 全 PASS。分片走 `Enter-VbpShard` 轮转。 |
| 账 #200（提交 `77ce2b65` = §B35 的控件窗口字体，门 #327（run 37251283970，head 39394bcd，attempt 1）= 11 job 全绿、10 片各自 FAIL=0；CI 自己量的九条相关用例逐行真 PASS —— [VBP] dcsurf ... PASS 与 dcsurf_x86 ... PASS（两台各自的真跑判据面）、[STATIC] control_dc ... PASS（含这一刀新加的 D10）、[VBP-BUILD] charts_ucTreeMaps ... PASS (679,424 bytes) 与 charts_ucTreeMaps_x86 ... PASS (585,728 bytes)、四条形状针 dcsurf_dc_shape / dc_read_real / dcsurf_text_measure / text_measure_real 全 PASS） | **控件窗口字体换了、文字量跟着不动** —— 做 #196 第二条时撞见的，那时只留了读数 TH06。**归因靠探针、不靠猜**：`.build/b200probe/fontprobe.c` 拿一枚**裸 STATIC**（谁也没子类化过）实测 `WM_SETFONT` 之后 `WM_GETFONT` 回 NULL、`STM_SETFONT`/`STM_GETFONT` 那一对同样回 NULL ⇒ 这个窗口类**本身不记字体**。于是原来那三条读法（`vb6_GetControlLogFont` = .FontName/.FontSize 背后那一条、Print 落笔前的选字体、#196 的 `vb6_ControlMeasureTextPx`）在 PictureBox/Label 那一类窗口上永远拿不到用户设的那张，只能按 DC 的默认字体画；而 `picA.FontSize` 读回来还是 18（那是另一份自存的属性）⇒ 两头都不报错，只有量出来的数不动（设计期 18pt 与运行期改 20pt 都答 16）。证人 `FontPixelHeight` 那句整行不打也归到同一条因：它自己就是 `WM_GETFONT + GetObject`，拿到 NULL 直接 `return 0`。**改**：一处出口 + 一份自存 —— 新增 `static HFONT vb6_ControlFont(HWND)`（先问窗口、回 NULL 再读窗口属性 `VB6_CtrlFont`），setter 把自己 `CreateFontIndirectW` 出来的那张存进这个**新名字**（#185「一层一个属性名」同纪律），三条读法全改走这一处；旧字体的找法刻意是「先读自存、找不到才问窗口」，顺序反了会对真记字体的 EDIT/BUTTON **双删**（`WM_GETFONT` 回来的正是我们上一轮存进去的那张）。**顺带修掉一处 GDI 泄漏**：改前 STATIC 每写一次字体漏一张，因为那句 `DeleteObject` 依赖的 `WM_GETFONT` 恒回 NULL、旧字体永远找不到。已知边界（写在 §B35，不装绿）：窗口销毁时最后一张字体不被 `DeleteObject`（Windows 回收属性表、不认识 GDI 对象），普通控件没有统一的 WM_DESTROY 挂钩。**判据**：TH03 从「只留读数」升回真判据 `ok7 = (tA > tB) And (tB2 > tB) And (pfA >= 18)`，`TH03-FONT=True` 进 `$dcSurfExpected`（两台各一条）；真跑读数两台**逐行相同** = `TH06-FONTRAW a=29 b=16 b2=27 fsA=18 pfA=24 pfB=27`。**护栏**：哨兵 D10 四条 + 假 needle 真红过（把 measure 那一处换回裸 `SendMessageW(hw, WM_GETFONT, 0, 0)` ⇒ `FAIL D10 问窗口字体的那一行 = 2 处 -> vb6forms_ctrl.c:325 | :733`，换回即绿）；**D10 自己第一枪红得不该怪产品** —— 它报的是 3 处，多出来那两条是我给两行的**行尾注释**写了 `WM_GETFONT`：哨兵跳行首 `//`、不跳行尾注释，于是「注释把形状写出来」就自造了红 ⇒ 措辞改成不嵌那个 token。**A/B** 86 份（BASE = #201 那份 emit 快照、NEW = 现在这台）⇒ changed=2、**OFFENDERS=0**，那两份就是 dcsurf 两台、逐行归因全是夹具自己新增的那几行；注释定稿后又单独重发两台 emit 与快照**逐字节相同**（VB 注释不进发码）。真工程配对：ucTreeMaps 两台仍 rc=0、诊断面 0 error、exe 659,968 / 562,688 与 #201 那轮同尺寸。**新账 #202（§B37）**：探针这一轮补了一枚裸 `BUTTON`(BS_GROUPBOX)，**同样回 NULL** ⇒ `vb6forms.c` 里 Frame 标题带那处读法一并定了罪，修法就是把 #200 这一处出口跨文件递过去。 |
| 账 #202+#204（提交 `8252b080` = §B37/§B39 的字体覆盖面，门 #328（run 37255522248，head ed66b03f，attempt 1）= **completed / success**，这条结论今天独立读到两次（watcher 的 poll 17 与事后复核各一次）；**逐 job 的 FAIL=0 与用例级 PASS 行没拿到** —— 本机那台中转在取 jobs / 日志这一段反复把 JSON 顶成 HTML（11 片日志只落下一片），所以下面不写任何"CI 逐行真 PASS"的句子。这一批的风险点恰好是「改了 shape/widget 两族共用的字体问法」会不会动到别人的判据，那一格只有逐 job 读数能答；中转恢复后按 §D 这一行补记（补不上就按 #328 的 overall success 结，边界写清楚）） | **#200 立了那一处出口，但出口只在自己的文件里用** —— 本刀把覆盖面补齐：`vb6_ControlFont` 从 `static` 提出来进 `vb6forms_internal.h`，新增**唯一写口** `vb6_ControlFontStore`（setter 也改走它，所以那一槽位的 `SetPropW` 全仓仍只一行），六处站点接上：`vb6forms.c` 的 groupbox 标题带读（#202）与创建路存（#204）、`vb6forms_ctrlarr.c` 问模板 + 存给新窗口、`vb6forms_shape.c` 三处图钮标题的文字量、`vb6forms_widget.c` 的 Label AutoSize 宽度。根因还是探针那一条：STATIC **与** BUTTON 两类窗口收到 `WM_SETFONT` 之后都不答 `WM_GETFONT`（`.build/b200probe/fontprobe.c` 两种都实测），所以"只发不存"= 事后谁都问不到。**读数**（真跑，两台逐行相同）：从没被写过字体的 picB `name= fs=0 pf=0 th=16 / twip=240` ⇒ `name=MS Sans Serif fs=8.25 pf=11 th=13 / twip=195`；产品后果是**默认字体那批控件**的 `TextHeight` / `Print` 行距 / 缇换算不再按 Segoe UI 9pt 算（Fix 181 那一路发的字体量的人问不到）。**判据** FR01 三头钉（自存往返 + 证人 + 文字量跟着换）进 `$dcSurfExpected` 两台各一条，FR02 留原始读数；**假 needle 真红过**（把创建那一句 Store 注掉重编一台 ⇒ `FR01-DEFAULT=False`、TH05/TH06 跟着回到 16/240）。**顺手抓到**：picB 改 20pt 后 `tB2` 27→32 —— setter 现在从真的那份 LOGFONT 起步、只换 `lfHeight`；改前从一个问不到的 LOGFONT 起步，等于"改字号"顺带把字体族也换了（与 #153"存什么读什么"同族）。**护栏**：哨兵新增 D11（出口定义 1 + 声明 1，D10 那条带 `static` 的正则同步改掉否则一提出来就哑；全 src/rtl 的 `WM_GETFONT` = 1 + 状态条两处的**具名豁免**，豁免设计成一旦删就红；存字体站点 = 3 且分属三文件）。**没验过的两头写明**：shape/widget 那四处改的是"用哪张字体画/量"，今天没有数值判据钉住，别当成已验证；数组那一路 `vb6_CtrlArr_Load` 全仓零调用者（发码侧没有 `Load <数组>(n)`），接进来是口径、不是产品修复。**A/B** 86 份（BASE = #200 那份 emit 快照）⇒ changed=2、**OFFENDERS=0**，那两份就是 dcsurf 两台、4 个差异块全是夹具新增行（无一行删除）。真工程：ucTreeMaps 两台 rc=0、0 error、warning 面 109 VB3001 + 6 VB3003 + 1 VB4001 一字未动，x64 exe 659,968 → **660,480**（RTL 真的重新内嵌了，x86 那份尺寸没动 = 段对齐吸收，见 §B35 那条"尺寸没变当不了没重编的判据"）。**新开 §B40/账 #205**：状态条那两处没接，理由是别人的族 + 面板宽度会跟着变、要先读 `tests/ctrlstatusbar` 的判据口径。 |
| 账 #196 第三条（提交 `479b202f` = §B31 的 `ScaleX`/`ScaleY` 换算那一站，门 #329 = run 37260903318、head 511d356e、attempt 1 = **11 job 全 completed/success**（含 Build C3.exe 那一片；这个 overall 读数今天独立落到两次 —— watcher 的 poll 8 与事后复核各一次）。**用例级的 PASS 行仍然没拿到**：这轮 jobs 数组里没给 `log_url`，改走 `/actions/jobs/{id}/logs` 拿到 10/11 份，但取回的正文里一条 `[VBP]/[CODEGEN-NOTE]/[STATIC]` 都没有（本机那台中转对日志端点给的不是真日志正文，重试到后面连 jobs 那一条也顶成非 JSON）⇒ 这一行不写任何"CI 逐行真 PASS"的句子） | **窗体型接收者缺的是换算那一站，而上一轮预告的修法被普查推翻** —— 预告写的是「缺的出口 = `vb6_WindowScaleX/ScaleY(void* hwnd, ...)`」，理由是 vbUser(0) 那一档要问这枚窗口自己的 ScaleWidth/ScaleHeight；这一轮普查调用点 = Charts 2020 五份 .ctl 各 4 处 + `PropPagFMR.pag` 2 处（=22）+ `tests/VBFlexGridDemo/Builds/VBFlexGrid/VBFlexGrid.ctl` **94 处（全是 `UserControl.` 前缀，Fix 111 早就通了）** + `MainForm.frm` 2 + `Common.bas` 2 + `archive/ctxWinsock.ctl` 2，实参只有三种形状 —— 字面 `vbPixels(3)` / `vbContainerSize`·`vbContainerPosition` / `.ScaleMode`·`Me.ScaleMode`·`UserControl.ScaleMode`（#197 下发之后答 1 或 3），**0 处传 0=vbUser** ⇒ 发的是**不带 HWND** 的那一条。**改**：RTL 新增 `vb6_ScaleUnitX/Y(double x, int32_t from, int32_t to)`（就是原 `vb6_UserControl_ScaleX/Y` 的身体），UC 那两条改成**薄壳直接 return 这一条** ⇒ Fix 111 那一档与窗体型这一档不可能分家（而不是在 #177 那张单位表旁边再抄一遍）；后端新增一处权威表 `controlScaleMethod`（`cgen_util_ctrl.cpp:1518`，只登记 Form/PictureBox 两档、**不给通用行**）+ 三处码头各查一次（With 形 `cgen_expr_with.cpp`、带括号裸形 `cgen_expr_call_com_bind.inc:263-290`、裸名 `cgen_expr_ident_builtin.inc:242-246`）；裸名那一处把改写门从 `isDesignerModule_` 放宽到 `isFormModule_ || isPropertyPageDesigner_`（§B31 开头那格「`.pag` 里裸 `ScaleX(...)` 认不出来」就是它），`.ctl` 里裸写那一形**没动** —— 仍由 `kHostPseudoRows` 那两行 UC 条目负责（`cgen_util_com.cpp:495-496`，HPF_BARE），现在只是多一层薄壳；With 形**刻意不置 `pendingChainObj_`**（换算不吃句柄，置了调用点会把 HWND 前置成第一个实参 ⇒ 表错位）。**没接的一头写明**：传 0=vbUser 时 VB6 要的是这枚窗口自己的用户坐标系，而 `vb6_ScaleUnitsPerPx` 的既有口径（#177 就写明）把 User/Container/unknown 都按像素 ⇒ 那种调用现在静默给恒等值；语料 0 处 ⇒ 记在这里不装绿，要接就是「换算拿 HWND」那一条。**判据**：夹具 `tests/dcsurf` 加长 SX10/SX11 —— SX10 是**四形逐数相等**（`picB.ScaleX(1440,1,3)` / 裸 `ScaleX` / `Me.ScaleX` / `With picB : .ScaleX`）再加 `ScaleY` 那一形与一枚**问窗口**的证人 `GetDeviceCaps(LOGPIXELSX)`，一头钉「四形同归一处」、一头钉「数真是按 DPI 换算的而不是烘出来的常数」；`SX10-FOURFORMS=True` 进 `$dcSurfExpected`（两台各一条），SX11 留原始读数。**两台真跑逐行相同** = `SX11-RAW pic=96 bare=96 me=96 with=96 y=96 dpi=96`（1440 缇 @96dpi = 96 像素，`.build/b196out/n64.out` / `n32.out`）。**负控** = #196 那台 BASE 真编今天这份夹具 ⇒ exit 2、C2039×3 + C2440×7、诊断面三条 `VB4001 P17.1: Unknown control property`（hDC / TextHeight / **ScaleX**）全在、一条读数都不出（`.build/b196out/b64.log`）。**形状针两条**：`dcsurf_scale_units`（present = 四形各自那一行，absent 钉住改前两形 —— `vb6_ComCallDouble(vb6_hwnd_picB, L"ScaleX"` 与 `.ScaleX(1440, 1, 3)`）+ `scale_units_real`（钉在**真工程** ucTreeMaps 那一行）。**哨兵** `check_control_dc.ps1` 新增 D12（权威定义 1 / 三处码头各≥1 / 字面量恰好 2 行且分属 {`cgen_util_ctrl.cpp`, `cgen_expr_ident_builtin.inc`} / 不给通用行）：假 needle 真红过（注掉 With 那条码头 ⇒ `FAIL D12 少了一条码头: cgen_expr_with.cpp`，还原回绿），**而它第一枪红的是我自己的规则**——写「恰好 1 处」时忘了表自己那一行也是字面量。**护栏 A/B** 86 份（BASE = `b204_new_emit`，即 #202/#204 那台快照）⇒ changed=6、**OFFENDERS=0**，逐份归因 = ucTreeMaps 两台各 1 块（`ScaleX(...)` → `vb6_ScaleUnitX(...)`）+ **VBFlexGridDemo 两台各 1 块 2 行**（`MainForm.frm` 那两行布局以前发成 `vb6_VariantToDouble(vb6_VariantFromComResult(vb6_ComCall…(L"ScaleX"…)))` = 拿 HWND 当 IDispatch 问属性 ⇒ 交回 Empty、再按数值打，现在答得出数；这两份一开始被我当 offender 抓到，查实是正当改道后补进 WD，与 #197 那轮同工程那条 `Me.ScaleMode` 同族）+ dcsurf 两台各 4 块（全是夹具新增行、无一行删除）。**结构性断言**：86 份新产物里以调用形式出现的裸 `ScaleX(`/`ScaleY(` **0 处**（`grep -rlE "(^\|[^.a-zA-Z_0-9])Scale[XY]\("` 在 `b206_new_emit/` 上返回空）。**真工程**：ucTreeMaps 两台 rc=0，exe x64 660,480 → **660,992**、x86 562,688 不变（段对齐）；诊断面 109 VB3001 + 6 VB3003 + 1 VB4001 而**那条 VB4001 现在是 `TypeLib reference path not found`** ⇒ `Unknown control property` 这一族在真工程里**归零**（`.build/b196out/tm64.log` / `tm32.log`）。**本账三格（hDC / TextHeight·TextWidth / ScaleX·ScaleY）到此全部出完**；剩下的红点不是编译面 —— #199 那条（`.pag` 里的控件从没被创建，甲/乙口径待拍）与 §B40/#205（状态条那两处 D11 具名豁免）。**本刀的用例级读数后来在门 #330 上补齐**（head 14caaf2a 含本刀的代码）：`[VBP] dcsurf ... PASS`(vbp#1) 与 `dcsurf_x86 ... PASS`(vbp#2)、`[CODEGEN-NOTE] dcsurf_scale_units ... PASS` 与 `scale_units_real ... PASS`(syntax)、`[STATIC] control_dc ... PASS`(compile)，另 `[VBP-BUILD] charts_ucTreeMaps ... PASS (679,424 bytes)` 与 `charts_ucTreeMaps_x86 ... PASS (586,240 bytes)` —— 那两字节数是 CI 那台的读数，与本地那两台（660,992 / 562,688）不同档，别混用。取日志的方法与那 4 片"101 行 0 用例行"的截形状见 §B40 那一行。 |
| 账 #205（提交 `a4986c88` = §B40 的状态条那两处问字体，门 #330 = run 37262853393、head 14caaf2a、attempt 1 = **11 job 全 completed/success**，逐片计数全为 FAIL=0：syntax 151/151、bas#1 44/44、asm 13/14、vbp 四片 48+1skip / 54 / 50 / 50（那条 SKIP 是 `test_vbman` 的 COM 32 位视图没注册，早就在那儿）。本轮新用例在 CI 上逐行真 PASS：`[VBP] sbfont`(vbp#3) 与 `[VBP] sbfont_x86`(vbp#4) —— 这就是"字体参与排版"那条判据在两台真跑过；邻居那条没动：`[VBP] ctrlstatusbar ... PASS`(vbp#2)。这一轮的日志**取全了 11/11**：jobs 数组里没有 `log_url`，改走 `/actions/jobs/{id}/logs`，并且**自己处理 302** —— 禁用自动跳转先拿 `Location`，再**不带 Authorization** 取正文（urllib 会把 Authorization 头带到跨主机的重定向上，存储端就答 401）。第一次取时有 4 片回的是「101 行、0 条用例行」那种**截断形状**，重取才见到真数 ⇒ "日志取到了"不等于"读到了"） | **D11 里那条具名豁免按设计自己消失了** —— 状态条是 RTL 自己注册的窗口类，`vb6forms_statusbar.c` 那两处（sbrContents 排版量宽 + 画文字）以前都是裸问窗口，恒回那个 NULL ⇒ 量的与画的都按 DC 的默认字体。**动手前先读了另一位作者的期望表，结果判据的形状变了**：`tests/ctrlstatusbar` 那 42 条里与宽有关的只有 SB10-W2=120 / SB36-SETW=123，两条钉的都是**显式给过的 Width**，而 `vb6_StatusBar_GetPanelWidth` 返回 `e->width`（请求值）⇒ 排版结果在 VB 侧没有现成的门，这一刀既动不到那两条针、也借不到那条门 ⇒ 判据改去**问窗口本人**：`SB_GETPARTS`(WM_USER+6=1030) 交回各格右边界（Declare 用 `LongPtr` + `ByRef … As Any`，同 LabelPlus 那一形）。**读数**（两台逐行相同）：改前 `ct8=140 ct20=140 after=140` = 同一串文字在 8pt 与 20pt 两枚上量出同一个宽、运行期改字号也不动；改后 `ct8=110 ct20=260 after=260`；证人 `pf8=11 pf20=27` **改前改后都一样** ⇒ 设计期那两张字体本来就下发到位（#204 存的），只是没人去问。**判据三头**（新夹具 `tests/sbfont`，两台各一条进清单）：SF02 两枚同串不同字号必须不等宽（拦"字体不参与"）、SF04 运行期改字号后重排要变宽并与 20pt 那枚对上（拦"只认设计期"）、SF05 显式给过 Width 的那格右边界不许挪（拦"顺手把所有面板都重排"）；RAW 三行只钉前缀（110/260 是 DPI 的函数，#147 那条口径）。**负控 = 同一份夹具在改前那台真编真跑** ⇒ `SF02=False`、`SF04=False`、旧数全回来（`.build/b207out/x64.out` / `x32.out`）。**护栏**：哨兵 D11 从「want 3 = 出口 1 + 状态条具名豁免 2」收成「want exactly 1」，PASS 行打印 `全仓裸问 1` —— #202/#204 那轮把豁免写成"删掉就红"就是为了这一刻；这一刀只动 RTL ⇒ A/B 86 份产物**逐字节相同**（`b207_new_emit` vs `b206_new_emit`，diff=0），另把两份状态条工程纳入普查面（4 份新快照，BASE 里没对应份 ⇒ 只留档不当判据）；相邻三枚哨兵各自复跑全绿（`check_di_stubs` 624=624、`check_uc_scale_units`、`check_host_pseudo_table` 55 行）⇒ 新夹具那枚 `Declare … SendMessageW` 没要新桩；**别人那一族的 42 + 10 条判据逐条复核两台全在、一字未动**（`.build/b207c_check.py` 就地从 run_tests.ps1 抽期望表，不手抄）。**没验的一头写明**：画的那一遍（`C3SbPaint`）没有独立数值判据 —— 钉住的是"量的与画的问同一处"这条普查与 measure 那两头。**新账 §B41/#206**：`Panels(i).Width` 交回的是请求值而不是排版后的宽（sbrContents/sbrSpring 在 VB 侧读不到几何，实测改前那版读回 0），今天刻意没改 —— VB6 那一读数的单位口径（缇 vs 像素）还没量准，猜着改等于把"读回 0"换成"读回错单位的数"。 |
| 账 #207（提交 `9e2c439a`，门 #331 = run 37266997643、head 59416448、attempt 1 = 11 job 全 completed/success，逐片 FAIL=0（vbp #1 那片 PASS=48 / FAIL=0 / SKIP=1，唯一 SKIP 是已知的 `test_vbman`）；本轮判据行在 CI 上逐行真 PASS —— `[VBP] frxdata ... PASS`，邻居 `[VBP] sbfont` / `sbfont_x86`、`[STATIC] control_dc`、`[CODEGEN-NOTE] dcsurf_scale_units` / `dcsurf_text_measure` / `dcsurf_dc_shape` 一条没动。= §B42 的 .frx ComboBox 那一档） | **Fix 195 把 .frx 的三种 blob 布局钉准了，接线却只写了一档** —— `emitControlFrxProps` 里只有 `controlType == ListBox` 才发 `LB_ADDSTRING`/`LB_SETITEMDATA`，而普查 `List =`/`ItemData =` 指向 .frx 的 **17 + 17 处全在 ComboBox 上**（Charts 2020 四份 demo 的 "Number of Series" / "Chart Style" / "Legend Position"）⇒ 下拉框是空的。**产品后果不是难看，是整块图不画**：`ucTreeMaps/Form1.frm:357` 那句 `If Combo1.ListIndex = -1 Then Combo1.ListIndex = 4: Exit Sub` 靠"空 ⇒ 第一句就退"把 Form_Load 掐死。**改**：不新开机制 —— 同一处出口按控件型取**消息对**（`CB_ADDSTRING`/`CB_SETITEMDATA`），ListBox 那一路发出的文本逐字节不变；两条创建路（顶层 / 容器子控件）本来就共用这一个 lambda ⇒ 只改一处。**判据**：夹具 `tests/frxdata` 加长一枚 ComboBox，**复用同一份 blob 的两个偏移**（这样缺的就只能是接线，不可能是解码器）；FD8 四头各自对上（ListCount=3 / List(0)="1234" / ItemData(0)=5 / ItemData(2)=-7），FD9 留原始读数；**负控 = 改前那台真编真跑同一份夹具** ⇒ `FD8-COMBO=False`、`count=0 item0= id0=-1 id2=-1`（`.build/b211out/b64.out` / `b32.out`；改后 `g64.out` / `g32.out`，两台逐行相同）。**真工程读数（截图）**：ucChartBar 的 demo 四个组合框都答得出设计值（"Grouped Column" / "2 Series" / "TOP" / "Aling Left"），ucTreeMaps 的 demo 点 Random 之后**树图整片画出来**（分块 + 名字 + 图例 2000..2004）—— `.build/b211out/bar32.png` / `tm_after.png`。**护栏 A/B**（BASE = `b207_new_emit`）⇒ inputs=98、same=86、changed=4（ucChartBar 两台 +48、ucTreeMaps 两台 +15）、new-only=8，逐行归因 **0 删除、0 无法归因**（第一版分类器只认含 `CB_*` 的那一行，把两行一组的 `{ wchar_t* vb6_witem = ...` 报成 16 条假 offender —— 又是"按整行判"那一课）；ucChartBar / ucTreeMaps 两台真编 rc=0。**顺带把四台 VbEclipse 的当前视觉基线立了**（读数在 022 那条 LAST_RUN）：Charts 主窗体十枚控件全画、czUI demo 全画、VBFlexGrid demo 全画、ucTreeMaps demo 要人手点 Random（= 新账 §B43/#208：程序改 `ListIndex` 不发 `Click`，而 VB6 发，那句 `Combo1.ListIndex = 4: Exit Sub` 整个就是靠这个惯例启动首绘的）。**工具课两条写进 §C 第 12/13 条**：`CopyFromScreen` 会截到别人盖在上面的窗口（要用 `PrintWindow(PW_RENDERFULLCONTENT)`）、PowerShell 的 `DllImport` 不带 `CharSet` 时按 Ansi marshal ⇒ `GetClassNameW` 永远只回一个字符（"Button"→"B"，看着像产品截文字）；另外 `tests/frxdata/FrxData.frm` 是 **GBK**，改它必须按字节插 ASCII。 |
| 账 #209（提交 `aaee1bac` = §B44 的动态数组元素检查 + 新哨兵 `scripts/check_sa_access.ps1`，门 #333 = run 37278620691、head d61d9065、attempt 1 = 11 job 全 completed/success、10 片 FAIL=0，`[STATIC] sa_access ... PASS` 逐行读到） | **一维动态数组的元素访问以前不问「描述符在不在、下标在不在范围」** —— `VB6_SA_AT` 是裸指针算术，`With m_Serie(Index)` 撞上从未 ReDim 的数组就读 NULL+0xc ⇒ 原生 0xC0000005（ucChartBar demo 点 Random，三轮同一偏移 0x1abf2，`-g` + `C3_CRASH_TRACE` 符号化到 ucChartBar.c:591）；收成 `vb6_SaElemPtr` 一处 inline + `vb6_SaElemFail` 冷路径抛 9，与 UBound/LBound 那两条 rev2 同口径。判据七枚私有过程（含两枚负控钉住范围内读写照旧）本地 x86+x64 全绿；A/B 98 份 `--emit-c` 逐字节相同（只动 RTL）；新哨兵 `[STATIC] sa_access` 五条规则、两条负控真红。多维那一支（ND_AT1/2/3 与 `_ndoff_` 兜底，存量 1468 处）刻意没动，见 §B44 末段 | 已提交，待门 |
| 账 #211（提交 `d61d9065` = §B45 的 pbsub 驱动加固；门 #333 同批，CI 上 `[VBP] pbsub ... PASS` 与 `[VBP] pbsub_x86 ... PASS` 两条逐行读到，两片各自 FAIL=0） | **门 #332 两次同一条红 `[VBP] pbsub_x86 ... FAIL (output mismatch)`，但进程 rc=0、PB-DONE 照打 —— 红的是判据读法，不是产品**：五条鼠标通知是 PostMessage 异步发的，夹具在**固定第三拍**读计数；本地带载 24 趟复现 2 趟全 0，改成“到齐才读、最多 20 拍”后 48 趟（x86+x64）全绿，其中一趟到第 8 拍才到齐。先把 #209 摘干净：pbsub 的两份 `--emit-c` 里 `VB6_SA_AT`/`SafeArray` 0 处，坏跑 stderr 空（无 `[SA]` 无 `Unhandled error 9`）。判据强度没降：`PB-CNT` 那行一字未动，新增 `PB-WAIT done=` 钉住“是不是靠到齐退出”；负控 = 注掉五句 PostMessage ⇒ 等满 20 拍读出全 0 + done=False，九针全缺。A/B 98 份：96 份一字不动，变的两份正是 pbsub 两片，逐行只多 `allIn` 那一族 | 已提交，待门 |
| 账 #214（提交 `323ab077` = §B46 的多维元素访问出口 + 4 秩发码补全 + 哨兵 A4 改口径与新 A6/A7；门 #334 = run 37287453662、head 323ab077、attempt 1 = 11 job 全 completed/success、逐片 FAIL=0，`[STATIC] sa_access ... PASS` 逐行读到，bas 两片之和 87→88 正是本轮新用例那枚进了门禁且绿） | **多维动态数组的元素访问同样不问「描述符在不在、每一维在不在范围」**：2 秩越界是**静默拿到隔壁那格**（`a2(3,1)` 回 12，VB6 是错误 9），从没 ReDim 取 `d(1,1)` 读 `(NULL)->data` 当场 0xC0000005，而 4 秩以上另有一条更糟 —— 发码把实参拼成 `(int[]){i0, i1}` 只塞两枚下标却按实际秩数交出去，于是 `Dim a4(1 To 2,1 To 2,1 To 2,1 To 2)` **在范围内也**崩（x86/x64 同形，1/2/3 秩各有专用分支所以一直是对的）。收成 `vb6_SaNdElemPtr` 一处 inline（先问形状 `dimCount`∈1..16 —— 一维描述符首字段 0x5A1D 天然落不进去，阈值抄的是 `vb6_LBoundND` 那条已有的 —— 再逐维问上下界）+ `vb6_SaNdElemFail` 抛 9；AT1/AT2/AT3 宏体改走它，新加 ATN 一次交全秩数与全部下标，复合字面量必须裹最外层括号（少了就是 cl C4002，实测）。新夹具 `test_arr_nd` 八针 x86+x64 各真编真跑逐行相同；A/B 98 份 `--emit-c` 逐字节相同（census ND_AT2 1468→1468、ATN 与 `_ndoff_` 各 0→0）；真编译 grid/charts × 两位数四片 rc=0；哨兵 A4 改成 0 条 + 新 A6/A7，七条负控逐条真红且逐字节还原；真工程带 `C3_SA_TRACE` 五趟 `saNd=0` ⇒ 那 1468 处全在范围内，加检查无可观察行为变化。另留一条读数纪律：`Form2.frm:504` 的 `Randomize Timer` 让「哪几枚 UC 宿主对 hover 有反应」在同一枚老 exe 上三趟就不同，这条不能当判据 | 已提交，门 #334 绿 |
| 账 #212（夹子提交 `0364dfbb` = §B47 的 `tests/pberr` + `run_tests.ps1` 登记；门 #335 = run 37293132648、head 0364dfbb、attempt 1 = 11 job 全 completed/success，四片 vbp 的 TOTAL 之和 224→226 正是 pberr 与 pberr_x86 进了门禁且绿（vbp #1 由 PASS=48 SKIP=1 TOTAL=55 变 PASS=50 SKIP=0 TOTAL=56，vbp #4 TOTAL 55→56，vbman 那枚 SKIP 随分片挪到 vbp #3），`[STATIC] sa_access ... PASS` 逐行读到） | **#209/#214 交回运行期错误 9 之后，「抛在实例方法体内」这一路的另外两面一直没量过** —— 方法自己写了 `On Error GoTo` 时 9 会不会越过它 / 跳走之后那枚对象的成员还读不读得出（longjmp 没解链实例栈的话，现场就是「错误被吃掉 + 对象已坏」）。实测**没有产品改动**：x86 与 x64 真编真跑逐字节相同 `PE-1d=9 / PE-2d=9 / PE-null=9 / PE-caller=9 / PE-state=20-40 / PE-DONE` ⇒ 落回自己那枚方法的处理器、没处理器时交给调用方、三次抛出后成员照旧读出 20-40，这一路完整可用。三枚负控全走夹具（下标换成范围内 / 证人成员置 0 / 删掉驱动的 `On Error GoTo`），分别让 `PE-1d` `PE-state` 与「整片读数消失 + RUN=9」变红 —— 顺带记一条：**拿掉处理器前的 `Exit Function` 当不了负控**，抛出发生时根本走不到那一行（实测输出照旧）。登记按 §B46 那条新约束插在整条语句之后、PSParser=0 错、numstat +14/-0；夹具三文件从 LF 归一成 CRLF 之后重编重跑，六条读数不变。边界写明：钉的是工程类实例方法那一路，不是 UC 宿主带 `Extender`/`ScaleMode` 那一路；只断言 `Err.Number`；`On Error Resume Next` 那一形没问 | 已提交，待门 |
| 账 #216（提交 `1df66f09` + 登记修复 `332bfd14` = §B48 的"位运算与 Not 的结果类型收成 TypeSystem 一处权威" + `tests/test_bitops.bas` 24 针（x86+x64 两形）+ 新哨兵 `scripts/check_bitwise_authority.ps1`；门 #337 = run 37313706942、head 332bfd14、attempt 1 = 11 job 全 completed/success，bas 两片之和 88→90 = test_bitops 两形进了门禁且绿，compile 片 24→25 里新那枚就是 `[STATIC] bitwise_authority ... PASS`，pberr/pberr_x86 在本 head 复绿；#336 那次红是登记行的 TAB（见 §B48 末段），同轮 tabwalk 一条红按其产物逐字节相同归到 #213） | **`And/Or/Xor/Eqv/Imp` 与 `Not` 的结果类型这条口径在仓里写了两份，发码那份对位运算恒答 Boolean** ⇒ 位运算数一进字符串上下文就打 True/False（`CStr(a Or b)`、`"x=" & (Not 5)`），装箱走 `vb6_VariantBool`，COM 实参更把 `hDC Or 0` 按 VT_BOOL 交出去（真工程 VBFlexGridDemo 的 `Render(...)` 就是这一条）；而同一表达式先赋给 Long 变量再打印一直是对的 ⇒ 差的是"问类型"不是算数。收成 `TypeSystem::bitwiseResult` / `logicalNotResult` 一处，语义层与发码层四个点全调它。24 针 x86+x64 各真编真跑逐行相同；负控=把权威毒成恒 Boolean ⇒ 16 条数值针全红、五条布尔面针不动；发码面 A/B 100 份 changed=8 且 VBFlexGridDemo 那 36 条有变行逐条归到三类；收成一处之后与手搓那版**逐字节相同 100/100**；真编译四片 0 error C / 0 LNK；哨兵在 HEAD 那两份消费点上 B1–B4 全红 | 已发货，门 #337 绿 |
| 账 #217 第一刀（提交 `eef2c199`+`ab1db9c0` = §B50 那一族 `Case Is` 的假标识符不再进 AST + `tests/test_caseis.bas` 7 针 x86/x64 + 两条 [CODEGEN-NOTE] + `[STATIC] caseis_shape`） | 门 #341（唯一红 = frmevents 抖动，与本刀无关）→ 门 #342 全绿；发码语料 CENSUS `'Is'` 138→0，A/B 100 份 same=90 changed=10 全 +0 行 unattributable=0 | 已发货，门 #342 绿 |
| 账 #217 第二刀（提交 `d2942bc8`+`910b37b8` = §B51 那一族文档隐式对象收成 `Module::docKind` 一处写两处读 + `tests/dochost/dhExp.ctl`/`dhImp.ctl` + `[STATIC] dochost_authority`） | 门 #342 = run 37345079456、attempt 1、11 job 全绿；compile 片 27→28、syntax 片 154→156；全仓语料 VB3001 2934→158，两份真工程各 499→19 / 499→34，宿主符号与 `_vb6_select_` 一动不动 | 已发货，门 #342 绿 |
| 账 #220（本节 §B52 = `Picture.Line` 尾部的 `B`/`BF` 由 parser 在 style 格折成字面量 1/2、RTL 那两枚裸名 C 全局连 extern 一起删 + `tests/test_nameclash.bas`（x64/x86 真跑）+ `tests/pcline/PcForm.frm` 的 [CODEGEN-NOTE] 四针两 Absent + `[STATIC] rtl_naked_names`） | 发码语料 A/B inputs=100 changed=**4**（ucProgressCircular 两份 × 两档），每份 +8/−8 行且每行只差最末一格 `vb6_ComPackValue(B\|BF)` → `vb6_ComPackInt(1\|2)`，另 3 行 VB3001 纯删（Charts 主工程 34→31），其余 96 份一行没动；撞名探针改前 RC=1/5 诊断/no exe → 改后 RC=0/exe/`NC-B=13` | 已发货，门 #343 绿（11 job 全 completed/success；bas 两片 47→48、compile 28→29、syntax 156→157） |
| 账 #221 = C29-PL-a（提交 `27767255` = `Picture.Line` 从 COM 兜底改道到原生 `vb6_ControlLine`：RTL 新出口 + `controlCanvasMethod` 表 + withm 码头 + 成员侧打标记；旗标改按字母位折 B=1/C=2/F=4 ⇒ BF=5、`C`/`F`/`CF` 从此有落脚点；`tests/pcline/PcDraw.{frm,vbp}` 画完问像素四形各钉两头 + `pcline_flag_folded` 换针 + 哨兵 N5/N6） | 发码 A/B inputs=100 changed=**4** same=96，每份 +8/−8 全是同一条调用换出口；CENSUS `L"Line"` **32→0** / `vb6_ControlLine(` **0→32**，VB3001 146=146、`VB6_SA_AT(` 20972=20972、`_vb6_select_` 7346=7346；x64 与 x86 真跑逐行相同 `PL01..PL04=True`（`diag=255 boxedge=16711680 boxmid=16777215 fillmid=65280 circletop=255`）；哨兵 PASS `fold bits 1+1+1, handoff 1+1, AST 0, RTL 1/bits, canvas 1+1+2+1`；只写表+码头不写成员侧时夹具四形**全 False**（N6 第四条由此起）；真工程 ucProgressCircular 两档仍 rc=1，27 条诊断逐文件归因到 #222（Form1.c 25×C2198+1×C2084）与 #219（`Count`），`ppProgressCircular.c` 零条 | 已发货，门 #344 红在 control_dc 名单（已改 4→5）→ **门 #345 全绿** |
| 账 #215（提交 `48feae8e` = §B49 的"体级声明收成一条声明符一条 LocalDeclStmt" + `tests/test_bodydecl.bas` 8 针（x86+x64 两形）+ `[CODEGEN-NOTE] bodydecl_one_per_declarator` （Absent 钉 VB3001）+ | 门 #338 = run 37321722861、head bab5b0ef、attempt 1 = 11 job 全 completed/success；compile 片 25→26 里新那枚就是 `[STATIC] bodydecl_shape ... PASS`（逐行读到），syntax 片 151→152 是 `[CODEGEN-NOTE] bodydecl_one_per_declarator ... PASS`，bas 两片之和 90→92 = test_bodydecl 与 test_bodydecl_x86 进了门禁且绿；vbp #3 那条 SKIP 仍是 test_vbman（COM 未注册，与 #335/#337 同形）） | **体级声明有四条路、两种形状**：`Dim a, b` 由 parseDimStmt 在语句层手写一遍声明符解析并出两条语句，`Const/Static/体级 Public` 的多声明符行把 MultiDecl 原样塞进 LocalDeclStmt，而语义层 visit(LocalDeclStmt) 的 switch 不认这个 kind ⇒ **一枚名字都不登记、每条使用一条 VB3001**（VBFlexGridDemo 一片 778 → 502 条，全部是诊断行）；手写那份副本还落在共享实现后面，漏了 WithEvents 与「后缀即类型」两步 ⇒ `Dim a&, b&` 第二枚静默落回 Variant（TypeName 看不出，VarType 3/0 才看得出）。收成 `Parser::wrapBodyDecls` 一处，四条路全调它，手写展开删掉。A/B 100 份 same=98 changed=2 且两条差异逐条归到诊断行；真编译四片 0 error C / 0 LNK；哨兵在 HEAD 树上 P1..P4 十条红 | 已发货，门 #338 绿 |
| 账 #219（§B54 = RTL 裸名 `Changed` 的定义与 extern 撤掉 + `visit(IdentifierExpr)` 补"裸写的文档成员"那一格 + `kHostPseudoRows` 搬进 `src/common/host_pseudo.hpp` 并新增 `hostPseudoBareEligible`，语义层与发码层同问一句 + `tests/test_rtl_naked_changed.bas`（x64/x86）+ `tests/dochost/dhBare.ctl`/`dhBare.pag`/`dhTypo.pag` 三条 [CODEGEN-NOTE] + 两份哨兵跟着改） | 发码语料 A/B inputs=100 changed=**12** same=88，每份差异**全是删一行 VB3001**、产物 C 一行没动；VB3001 146→98（`Changed` 24→0、`hDC` 24→0，`new-names={}`）、`vb6_PropertyPage_Changed` 186=186；撞名探针改前 RC=1/no exe ⇒ 改后 RC=0/exe/`NC219-CHANGED=6`；哨兵 `tbl 55 rows` 与 `bareDoc 1/1/1/1/tbl1/names0` 全 PASS，红过一次是删掉 N7 读文件的负控演示 | 已发货，门 #346（run 37366164862、head `0026d87f`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0 |
| 账 #225（§B55 = `src/driver/rtl_embedded.hpp` 两枚 id 与 `c3rtl.rc` 对上 + 新哨兵 `scripts/check_rtl_resource_ids.ps1` 三处逐条对账 + `tests/run_tests.ps1` 注册 `[STATIC] rtl_resource_ids`） | 上游 `9a420157`（rev38 绘图家族，经 `c5aab787` 合进 dev）把 `.rc` 的 `223=头/224=体` 与 hpp 的 `223=C/224=H` 写对调 ⇒ 解包把 463 行的体写成 `vb6forms_draw.h`，而 `vb6forms.h` include 它、被 39 个 RTL 文件 + 每份生成模块 .c 带到 ⇒ 28 枚符号 × 约 49 份定义 = LNK2005×1225 + LNK1169；门 #347 十一片九红（只有不链接的两片绿） | 单变量真编译：HEAD 冷编 `RC=1 diag=1226` → 只改这两行 `RC=0 exe=True diag=0`；哨兵 census 125/125/125、孤儿 0，两条 R3 负控真红、还原逐字节相同 | 已发货，门 #349（run 37380276520、head `ca9f3719`、attempt 1）= 11 job 全 completed/success、非绿 0 |
| 账 #222（§B56 = 元素键唯一出口 `ctrlElemKey` + 事件 ABI 在登记那一刻翻（`prepareEventHandlerProc`/`applyEventHandlerAbi`）+ prelude 那份逐元素前向声明删掉 + `ucSinkEvents_` 让 `WM_LBUTTONUP` 两条 hasClick 判据都问表；夹具 ve_units 两头 + 发码针 `ucarr_evt_thunk_per_element`（Absent 三条 = 改前真实形状）+ 哨兵 `scripts/check_uc_array_event_sites.ps1`） | BASE 那台真编译 ucProgressCircular `RC=1 exe=None diag=27 kinds={'C2198': 25, 'C2084': 1, 'C2065': 1}` ⇒ 修后 `RC=1 diag=1`（只剩 `Count` C2065，属 #219）；真跑 ve_units 两档 `U-ARREVT=True`（`U-ARREVT-RAW i1=1 h1=1 i2=2 h2=2 ret=7`）；三条假形状逐条真红、还原 MD5 相同 | A/B inputs=92 same=74 changed=18，差异行 direct=752 blockfall=24 **无法归因=0**，x64/x86 逐份对称。**第二格**（撤 prelude 那张第二表、类型收成 `mapTypeRef` 一处）：单变量 A/B `inputs=92 same=86 changed=6`、28 条差异全在 `evtThunk` 签名、未归因=0；三处同源检具 cb 不符 5→0 / 原型不符 6→1；ucTreeMaps 两档出 exe | 已过门：#351（run 37386871868、head `482c3273`、attempt 1）= 11 job 全 completed/success、非绿 0（#350 红在两档 ucTreeMaps → #351 绿）|
| 账 #226（§B57 = `vb6_UserControlDesc` **末尾**加 `click` 槽 + `uc_host_window.c` 在 mouseUp 转调之后补一次 + `cgen_form.cpp` 发 `ucHostClick` 封装与末槽） | 语料里六枚 Charts UC 的 `RaiseEvent Click` 全写在 `UserControl_Click`，而那条子过程此前没人调（desc 压根没有 click 槽；`dblClick` 槽也 0 调用者）；#222 把兜底 arm gate 掉之后不补落点就是"编得过但一条 Click 都收不到" | 真跑真手势 `U-ARRCLICK=True`（`U-ARRCLICK-RAW hw=True idx=2 hits=1 ret=0`，两档相同）；负控 = 注释掉那条转调重编 ⇒ `U-ARRCLICK=False` 现场 `idx=-1 hits=0` 而 Fire 那头照旧 True；哨兵 C1/C2/C3 钉末槽与转调顺序 | 已过门：#351（run 37386871868、head `482c3273`、attempt 1）= 11 job 全 completed/success、非绿 0 |
| 账 #227（§B58 = `uc_host_window.c` 补 `case WM_LBUTTONDBLCLK` → `desc->dblClick` 一档 + 夹具第三头问 dbl/hits 两个计数 + 哨兵 C4 逐槽钉非零） | 与 #226 同形：desc 按位置填满 ⇒ 发码面永远看不出，`desc->` 读数里 dblClick 0 个调用者；六枚 UC 的 `RaiseEvent DblClick` 全静默（`CS_DBLCLKS` 早就立着） | 真跑两档 `U-ARRDBL=True`（`hw=True idx=2 dbl=1 hits=0 ret=0`）；负控 = 注释那条转调 ⇒ False 且现场 `idx=-1 dbl=0 hits=0`，另两头照旧 True；C4 假形状真红 2 条、还原 MD5 相同；92 份 emit 与上一轮逐份相同（纯 RTL 那一刀） | 门 #352（run 37388004707、head `50fb6d5f`、attempt 1）= 11 job 全绿、非绿 0，wall 8m41s |
| 账 #228（§B59 = `mapTypeRef` 的别名档改问类型本名 `aliasName`（点号最后一段），符号那几档不动 + 夹具 `tests/test_alias_spellings.bas` 两面钉） | 同一 VB 类型两种拼法给出两种 C 类型：`OLE_COLOR`→`int32_t` 而同义的 `stdole.OLE_COLOR` 掉兜底 `void*`；实物 = `.ctl` 声明 `EditSetupWindow(... As OLE_COLOR)` 而容器写 `As stdole.OLE_COLOR` ⇒ 发送侧交 4 字节、处理器收 8 字节指针（x64 高 32 位是垃圾） | 单变量 A/B `inputs=90 same=88 changed=2`、4 条差异全是那一枚处理器的 `void*`→`int32_t`，真 COM 限定名一族（39 种 / 124 处）零改动；三处同源检具 52 枚 thunk 的两类不符 → 0/0；BASE 那台跑同一夹具真红（少一枚 needle + 命中一条 Absent）；9 件工程两架构 rc=0 | 门 #353 |
| 账 #229（§B60 = 两条对象形态收成一处出口 `ctrlTypeOfMemberObject`，P20-42 的兜底改问它；类型仍出自 `controlPropType` 那张表） | `& uArr(0).Left` 判成 String（撞内置函数 Left）⇒ 拼接不套数值转换 ⇒ 裸 int 进 BSTR 槽 = 两架构 0xC0000005；同元素 `.Top` 只绕远装箱、非数组 `.Left` 正常 ⇒ 症状按名字分家很误导 | 夹具两头（RAW = 崩溃现场本身 + 四枚变量对上设计期几何）；BASE 那台跑同一份夹具两架构真崩、修后两架构 rc=0 读数相同；语料 A/B 90 份里只有 ve_units 变（纯夹具新增行，产品发码零改动） | 已过：门 #355（run 37398206822、head `187dc3e7`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m16s |
| 账 #231（§B62 = C29-1a/1b/C29-9 手抄在 `inferExprType` 里的三份名单（`kNumericFc`/`kStringFc3`/`kStrFcCd`+`kNumFcCd`）逐条搬进 `controlPropType` 那张表，问话只留一次且仍在 case 最前；`!= Unknown` 那道闸跟进表里） | 控件属性的**类型**两处各答，重合的四条靠"恰好一样"才没出事（#229 就是这道缝）；表外那 28 条名字散在推断函数里，改一处就把另一处的旧答案留在原地 | 这一刀**刻意零发码改动**：BASE 先冷存复捕证明与上一轮 90 份逐字节相同，改后 `inputs=90 changed=0 same=90`；新哨兵 `check_ctrl_prop_type_authority.ps1`（A1 旧名单回潮 0 / A2+A3 一处定义+恰好一个调用者 / A4 28 条名字逐条在表里 / A5 Unknown 闸 1 处）+ 三条负控各让一条红；22 份 check 全绿、真编译 4 件全出 exe | 已过：门 #356（run 37401361458、head `442da251`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m23s |
| 账 #233（§B63 = Form 的绘图状态属性收成"两张表成对登记 + 一份笔位存储 + 一套编码"：cgen 侧表删掉、`vb6_DrawSetI` 存裸值、笔位归 float 那一户、ScaleMode 改问 #197 那道权威） | `Me.DrawWidth = 3` 发成"把读函数当左值" ⇒ **C2106，两架构零产物**（写侧从没登记）；读回恒 +1（Step 累积漂）；`VB6_CurrentX` 一个属性名两套编码（绘图 int32 vs Print float 位图案）互读必错 | 新夹具 tests/fdraw 两头钉（写后读回 + 像素证人 + Print 之后读得到同一个数），BASE 那台跑同一份夹具真红；语料 A/B 340 份 changed=0（= 这一族零覆盖）；新哨兵 check_form_draw_state.ps1（S1 存储唯一 / S2 侧表不回潮 / S3 读写成对 / S4 编码对称）四条各证能红 | 已过：门 #357（run 37405355138、head `05f04f31`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m05s。订正一句读数方法：那台 watcher 回读 jobs 时被本机代理顶了一次，只写出 `jobs=0 non-success=0` 就收线 —— `conclusion=success` 配 0 条 job 不是"全绿"，是**没拿到读数**；补一次按 run id 回 API 复核才数到 11 条 （`.build/b309_verify357.py`） |
| 账 #234（§B64 = 绘图方法家族改问唯一权威 `vb6_ControlDrawDC`；`vb6forms_internal.h` 声明、`vb6_DrawAcquire` 只挡 NULL；census 跟着长：D1 排除声明行 / D2 5→6 / 新 D13 禁 draw.c 自己开 DC） | 无 bug 症状的重复实现：两份同口径 ⇒ 一改就静默分家，而 `check_control_dc.ps1` 的名单原本扫不到第二份所在文件 | 零行为改动（两分支逐条等价）+ 三条负控各证哨兵会红 + 邻域四枚真跑夹具 28 条 needle 零缺失 + emit A/B 90 份 changed=0 + 矩阵 4 件出 exe | 已过：门 #358（run 37409833257、head `2b3baf82`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m44s |
| 账 #235（§B65 = 画笔色两份存储合一：`vb6_DrawForeColor` 改问唯一出口 `vb6_GetControlForeColor`，撤掉私有 setter 与两个零引用导出；Printer 那族不动） | `Me.ForeColor = vbRed` 之后不带颜色的 `Me.PSet` 画出来是 **0（黑）**（改前两架构实测），带颜色的那条才是 255 —— 控件那侧本来就只有一份，Form 绘图自己另存了一枚 | 新夹具一头 FD11/FD12（**两面**：新点要蓝、旧点仍红）+ 逐字回退重编那台跑同一夹具真红（False / pen=0）+ 哨兵新 S5（属性名回潮=0，假针证红）；23 份 check 全绿、邻域四枚零缺失、矩阵 4 件出 exe | 已过：门 #360（run 37415128806、head `ea8dc7c9`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 8m31s |
| 账 #236（§B66 = 门 #359 的 vbp#4 红归因到夹具：阈值收线的 `*_Timer` 里自增排在提前返回之后 => exe 永不关窗，被 60s 超时杀；修法=计数器先走，哨兵 `check_fixture_timer_close.ps1` 第 24 道钉住这条口径） | 已过：门 #360（run 37415128806、head `ea8dc7c9`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 8m31s |
| 账 #237（§B67 = Print 从 vb6forms.c 搬进绘图家族：DC / 字体 / 色彩 / 单位四件都改问已有权威，推进量取刚写那串字的 extent（与 TextHeight 同一个量）；顺带撤掉本文件自带那份只认缇的 v * dpi / 1440，换算全部交回 vb6_ScaleUserToPx / vb6_ScalePxToUser 并显式写纵/横） | FD13/FD14 推进 == TextHeight（缇、点）、FD16 = 72 点与 1 英寸同一行（蓝/红分居两个 x 窗口）；行为负控改回旧形三条全 False + 哨兵 S6/S7 四条假针逐条能红 | 已过：门 #361（run 37421550422、head `d86b478c`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 8m49s |
| 账 #239（§B69 = 控件那户 `Print` / `Cls` 撤掉自存的像素笔位 `VB6_PrintX` / `VB6_PrintY`，两条函数改成只转调家族的 `vb6_Form_Print` / `vb6_Form_Cls`；家族那份 `Cls` 的背景色改问带 Fix 187 哨兵那份唯一出口 `vb6_GetControlBackColor`） | `pic.Print` 既不读也不动 `pic.CurrentX/Y`（实测推进 0 而同一枚控件答 `TextHeight` = 13、笔位放 60 而墨落在第 2 行、`Cls` 之后 `CurrentY` 仍是 400）= #237 在 Form 侧刚拆掉的「一份存储两种单位」在控件侧重演；而合并之后若不换色彩出口，控件那户 `Cls` 会从实测 0 变成按钮面 | `tests/pcline` 加一枚 picP + 五条 needle（PL08-PEN / PL09-CLSPEN / PL10-TWIPADV / PL11-BLACKCLS / PL12-STACK + 两条 RAW）两架构真跑；BASE 那台跑同一份夹具真红（Q01 推进 0 / Q03 Cls 后仍 400 / Q05 笔位不动）；三处**行为**负控各让自己那一条红（推进归零 → 08/10/12、像素当用户单位存回 → 只 10、色彩自己答 → 只 11）；哨兵 `check_form_draw_state.ps1` 新 S8 四条假针各证红（属性名回潮 0 / 两条码头必须转调 / 定义恰好一次且住在家族里 / `Cls` 只许问那条色彩出口，且**只看代码行** —— 先前只查体内文本时注释把判据顶成假绿）；`check_control_dc.ps1` 的 D2 因这一刀 6→4（哨兵先响，归因写进规则注释）；24 份 check 全绿 + 邻域 fdraw / pbsub / dcsurf 两架构零缺失；零后端改动 ⇒ 发码逐字节不可能变（RTL 是 exe 的资源），判据只能落在真跑那一头 | 已过：门 #364（run 37429327451、head `e5c66e4d`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 12m40s |
| 账 #238（§B68 = `vb6_VarCmp*(&A,&B)` 那四个取址点以前按"名字形状像左值"决定能不能 &，现在全问一处判据 `CCodeGen::cmpOperandMayTakeAddr`（回答只来自 `isDefinitelyVariantExpr`：声明那几张表 + 符号表），不许取址就走 `_Generic vb6_VariantFromValue` 装箱） | `Dim d As Double` 的地址被当 `vb6_VARIANT*` 递进 RTL ⇒ 按 VARIANT 布局读一个 8 字节标量：**相等的两个数答 False**（VC01..05 / VC10 / VC15 七条实测，两架构一致），且读过头；编译、运行、诊断都不响 = 静默给错答案那一族（#123 的 Byte 是同一形状的另一档）。FD13/FD14 当初因此只能写成 `Abs(a-b)<0.001` | 新夹具 `tests/test_varcmp_scalar.bas` 16 条两架构真跑（每条相等配一条近似不等，Long/String 两条本来通的腿也钉住）+ BASE 同一份夹具 7 条 False + 发码针 `varcmp_scalar_boxed`（必须装箱 / 五条旧形状必须不出现）+ FD13/FD14 换回直接相等 + 新哨兵 `check_variant_cmp_boxing.ps1`（第 25 道 [STATIC]，V1 一处判据且真问权威 / V2 五处问话 / V3 每条 & 拼接都被问过 + 旧形状禁回潮）三条假改动各证红；语料 A/B inputs=90 changed=2 逐行归因完（未归因 0），唯一真形状变化在编不过的 ucProgressCircular 裸名 `Count` ⇒ 能编过的语料零暴露；邻域 fdraw / pcline / bool_display / datelit 两架构零回归（后两条输出与 BASE 逐行相同）。附带订正：上一轮那句"暴露面 0 处"只统计了找得到声明的名字，漏了未声明裸名那一形 | 已过：门 #365（run 37436391539、head `9b9c1551`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m40s |
| 账 #232 ②（§B70 = `Me.Cls` 落 `vb6_ComCall(hwnd, L"Cls")`、`Me.Print "AB"` 落"取 Print 属性 + Item 下标"，两形都编得过、跑得起、一笔不画） | 画布家族的名字被答了四遍（成员侧名单 + 表达式码头 + 语句码头 + Form/Printer 段），而每一处都只认 PictureBox。**订正**：窗体自己那枚接收者其实认得（`cgen_form_ctrl_registry.inc:10` 把窗体名也登记进 knownFormControls_，类型 Form），缺的是 `controlCanvasMethod` 里 `cls` / `print` 那两档 ⇒ 上一轮"零实参的对象.方法在语句位置压根没进分派"那条结论只对了一半。语料 `Me.Cls` / `Me.Print` 各 0 处 ⇒ 潜伏缺陷；与账 #143 / #149 / #221 同一味（落 COM 兜底 = 静默空转，本线第四次） | 表加两档（Form + PictureBox；`line` 刻意仍只给 PictureBox —— Form 那一形是 12 参签名 `vb6_Form_Line`，签名不同不并表）+ 三条码头与打标记处全改问这张表与 `formCtrlSlot` + 撤掉 Form/Printer 段里 `cls` 的 Form 那一支（只留 Printer=EndDoc）。判据三面：①FDForm 加 FD18/19/20（墨 + 笔位两头钉，FD20 必须量**增量**，写成 `Me.CurrentY = thA` 会在改前那台因为继承上一行 Print 的推进而假绿）= BASE 两架构三条全 False、新台两架构全 True；②发码针 `form_canvas_family` 五必须出现 + 三必须不出现（BASE 缺 2 命中 3 ⇒ 两头能红），顺带还掉账 #224 的 ⑤；③`check_form_draw_state.ps1` 新 S9 四条（名字只在表里 / 三档两接收者齐 / 调用点恰好 7 / 打标记处必须问表），只看代码行，两条假改动各证红。护栏 = 语料 A/B inputs=90 **changed=0 / same=90**（只改谁能答）+ 25 道 [STATIC] 全绿 | 已过：门 #368（run 37454147931、head `9e5c9f91`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m46s；**第一轮门 #367 红在本刀自己的发码针上**（needle 里写了两个空格，而 `Invoke-CodegenProj` 比对前折空白 ⇒ 永远匹配不上），`9e5c9f91` 一处改掉 |
