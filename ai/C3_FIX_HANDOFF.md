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

### B41 `Panels(i).Width` 交回的是请求值，不是排版后的宽（账 #206，**两格都已出 —— 第十八刀 + 第二十二刀（门 #475/#476）**）

#205 做判据时撞见的：状态条的排版结果只活在 `w->rights[]`（`SB_SETPARTS` 那一份），而 `vb6_StatusBar_GetPanelWidth` 交回 `e->width` —— 设计块没给 Width 的面板一律读回 **0**（实测 `SF01-RAW` 改前那版就是 `small=0 big=0`）。后果不止"读不到数"：`sbrSpring`/`sbrContents` 两档在 VB 代码里没有任何几何可查，凡是按面板宽定位的东西（比如提示气泡、覆盖层）只能自己再算一遍。
**口径已经量准（2026-10-10，四发外部读数）—— 这一族的开工前置就此清掉，两格都动得了**。本机另有别人用 VB6 存下来的整批工程（`D:\vb_yqt4qPac`，5623 份 .frm），不是本仓那三份夹具，于是"设计值与实际宽的比"这条路取得到证据：
- ① **类型库那一手**（探针 `.build/b314_panel_tlb.cpp` → `b318_tlb_out.txt`，读 `C:\Windows\System32\comctl32.ocx` 内嵌的 ComctlLib 5.0 SP2，133 型）：`IPanel.Width` / `Left` / `MinWidth` 的 propget 一律 **VT_R4(Single)**；`PanelAutoSizeConstants` = sbrNoAutoSize 0 / sbrSpring 1 / sbrContents 2；`PanelStyleConstants` = sbrText 0 / Caps 1 / Num 2 / Ins 3 / Scrl 4 / Time 5 / Date 6 / Kana 7 ⇒ 与 RTL 那批 `VB6_SBR_*` 对上。**库里的 docstring 只有属性名、不带单位** ⇒ 单位不可能从 TLB 取，只能从②③取。
- ② **设计块那一手**（`.build/b321_panel_unit.py`：338 枚 `Begin MSComctlLib.StatusBar` 里 92 枚带面板宽）：VB6 设计器写的那一行键名是 **`Object.Width`**（不是 `Width`）。`sum(面板宽)/控件的 _ExtentX` 中位数 **0.986**、63/92 落在 [0.9,1.1]；`sum/控件自己的 Width(缇)` 中位数 **1.738**、落在 [0.9,1.1] 的 **0 枚** ⇒ 面板铺满的是 `_ExtentX` 那一档 = **himetric(0.01mm)**，不是缇。系数 1.7639 = 100/56.6929 与同一块里 `Width`:`_ExtentX` 逐位对上（3780:6668、7365:12991）。
- ③ **代码那一手**（`.build/b323_panel_code_unit.py` + `b324_panel_context.py`）：真源码里 `Panels(i).Width = …` 共 27 处，其中 3 处右式直接就是 `Me.ScaleWidth`（`报表/ListView报表打印模块/200612128325487/Form/PagrVisual.frm:504`、`报表/表格打印模块/20073722249267/Form/PagrVisual.frm:785`、`数据库/IC卡考勤系统源代码/200582218428876/…/BatchSetup.frm:483`），而**这三张窗体都没写 ScaleMode ⇒ 缇档** ⇒ 运行期读写的单位是**缇**。旁证：同族 `ColumnHeaders(i).Width` 赋的 literal 幅值 med=1100 / p75=2000 / max=3200（缇那一档，不是像素那档的几十..几百）。
- ④ **②③合起来才叫钉准**（`.build/b322_listview_unit.py`）：ListView 的 `ColumnHeader` 设计块里 `Object.Width` = 2646 / 2117 / 1411，反算缇正好 **1500 / 1200 / 800**（反算像素是 100 / 80 / 53.3，最后那枚不整）⇒ **设计器存的是「缇 × 72/127」（himetric），属性本身是缇**。这条换算与 DPI 无关（两个都是绝对长度单位）⇒ 发码期折一次就够：`tw = MulDiv(hm, 72, 127)`。
⇒ 于是这一族要补的是**两格**，且不必再等外部读数：
- **A 设计块那一格**（`cgen_form_ctrl_style_apply.inc:847` 那一趟读 `p.properties["Width"]`）：真 VB6 那一行键名带 `Object.` 前缀，而 `frm_parser.cpp:81` 保留点号全名 ⇒ **今天的真实工程里面板宽压根没进 RTL**（这才是"读回 0"的第一半；"读请求值"那半是第二格）。补法 = 认 `Object.Width` / `Object.MinWidth`，发码期按 72/127 折成缇再交给出口。语料面：全 5623 份 .frm 里 `Object.Width` **1634 处 / 356 份**，点号集合项形（`.Width =`）在状态条上 **0 处** ⇒ 我们夹具那形是自己造的，不是 VB6 的。
- **B 单位那一格**（RTL `vb6forms_statusbar.c`）：`Vb6PanelEntry.width/minWidth` 注释写**像素**、`SbLayout` 也按像素参与排版，而 VB 侧口径是**缇** ⇒ 出口两侧改用已有的那对唯一换算（`vb6_TwipToX` / `vb6_XToTwipX`，#184/#175 收成一处的那对，**别新开一份**），getter 改交**排版后的宽**（`SbApplyParts` 那份 offsets 差值）折回缇。
**刻意不跟着改的一格（记下别当遗漏）**：TLB 说 propget 是 Single，而 `vb6forms_memberobj.c:505` 走 `memSetI4` —— 折出来的缇取整后 I4 与 R4 在 `=` / `CStr` 两头同值，改它要动成员表与类型权威(#231)，收益只有 `VarType()` 一项。
**第二格的口径也有主了（2026-10-10 同一批语料，`.build/b363_panel_read_width.py`）**：扫 9546 份源文件找到 **10 处读面板宽/左**的写法，其中四行是"把一枚浮动的进度条按面板定位"—— `Form1.frm` 的 `.Move (sb.Panels(n).Left + pading), (sb.Top + pading), (sb.Panels(n).Width - pading*2), …`、`modAddProgBar.bas` 的 `pb.Width = sb.Panels(lPan).Width - 45`、`frmDataSource/frmEdit` 的 `ProgressBar1.Move 3060, …, Me.sbStatusBar.Panels(1).Width - 3060`。这些覆盖层只在** getter 交回排版后的几何 **时才对得准（弹簧档的面板"请求值"跟实际宽根本不是一回事），而设计块那一头量到的是 `sum(面板宽) ≈ _ExtentX`（63/92 铺满）—— 两条互证 ⇒ **§B41 的靶子成立：VB6 交回的是排版结果**。**但这仍是"作者期望"级证据，不是 OCX 行为级**：真要改，动手前有两件必须一并做 —— ① 邻居夹具 `tests/ctrlstatusbar` 那两条针（SB10-W2=120 / SB36-SETW=123）钉的就是"显式给过的 Width"，改了 getter 它们必翻 False ⇒ 判据要按 #157 那条重述（就地取基线问增量，别拿"总数=各步之和"当式子）；② VB 侧存储要从"像素"换成"缇"（本刀刻意留在像素那一档，见 §B129），两件事一起做才不会又留一份第二答案。
**判据必须自己造**（A/B 对这族永远沉默：本仓三份状态条夹具都没有 `Object.Width`，全语料 0 处 ⇒ 形状门 `changed=0` 是"零覆盖"而不是"没改到"，#233 那一课在这里重演）：夹具加一枚**按 VB6 真实形状**写的 `BeginProperty Panels {…}` / `BeginProperty PanelN {…}` / `Object.Width = …` 状态条，两头钉 —— 设计值折成缇读得到（拦 A）、弹簧/内容档读回排版后的宽（拦 B），证人用 `SB_GETTEXT`/`SB_GETPARTS` 问窗口本人（与 #205 那三条同源）。
**第一格已出（第十八刀，2026-10-10，门 #463 全绿 = run 38011638007、head `177fca98`、attempt 1、12 job 全 completed/success、非绿 0、wall 10m22s；`Emit manifest (shape oracle)` 那一跑同绿 ⇒ 新登记的那行被 CI 那台独立复算证实，而 vbp 四片全绿 ⇒ 新夹具 `sbhm` 两台真跑过了 —— 缺 .vbp 会让 Test-Vbp 直接报失败，绿就是跑了）**：发码侧认了 `Object.Width` 那一档 —— 新建一枚 RTL 出口 `vb6_StatusBar_SetPanelWidthHm`，它只做一件事：转调本族唯一那枚按真实 DPI 的权威 `vb6_HimetricToPxX`（`vb6forms.c`，与 #184/#175 那对缇换算住在一起）。**读数**（真跑，两台逐行相同）：夹具 `tests/sbhm`（**按 VB6 设计器的真实形状写**：`BeginProperty Panels {GUID}` + `BeginProperty PanelN {GUID}` + `Object.Width`）—— 改后 `w1=57 w2=100 e3=467 cw=467`（1500 hm 与 2646 hm 折成 57/100 px，正是 MulDiv(hm,96,2540)），改前那台**跑同一份夹具** `w1=7 w2=14`（两枚固定档面板退化成文字宽 = 设计值整格没进来）⇒ `HM01-HM-DESIGN-WIDTH` False→True。**同一刀里顺带修掉一格实测缺陷**：弹簧档的 `MinWidth` 被当成"加在剩余空间上的加项"而不是下限（`SbLayout` 第一遍占位 + 第二遍兜底 = 同一个答案两处），读数 `e3=493` 而 `cw=467`，多出来的正好一枚 MinWidth=26 ⇒ 判据 `HM03-SPRING-TILES` False→True。**护栏**：新哨兵 `check_statusbar_panel_hm.ps1`（第 51 道 [STATIC]）四头，负控两头各证能红 —— BASE 树（`wt_k17neg3`，改前那台编译器 + 同一份夹具）一次报出 R1/R2/R3 九条，名单被清空的副本报出 R4；夹具那四条判据里只有 HM01/HM03 是真翻 False 的存在性证人，**HM02 在 BASE 上也 True**（7:14 恰好也落在容差内）⇒ 它只是"只许一处换算"的形状护栏，不当罪证。**刻意留下的一格**：`vb6_HimetricToPxX` 之外，himetric 与像素/缇的折算在 RTL 里已有 13 个非注释行、住 5 份文件（§B129）—— 本刀只把**新增**那一处放进权威，没顺手并表（各有一份自己的判据面）。

**第二格动手前的三件前置全部量完（2026-10-10 同日续，读数留档 `.build/b436_panel_unit_readings.json`，脚本 `.build/b436_panel_unit.py`）**：
- **单位 = 缇，这条从「作者期望」升到「行为级」**：扫 9185 份源文件拿到 `Panels(…).Width/Left/MinWidth` 共 **36 行**（Width 35 / Left 2 / MinWidth 4），其中 **1 行把面板宽显式除以 `Screen.TwipsPerPixelX`** （`Global.bas`：`.Create 100, 2, o.Panels(1).width / Screen.TwipsPerPixelX` —— 拿 Win32 API 摆位子，只有缇才需要这一步）；**5 行右式直接含 `ScaleWidth/ScaleHeight`**（比上一轮 ③ 的 3 处多 2 处）；**3 行与另一枚 `.Width` 同式**（`pnl.Width = StatusBar1.Width - StatusBar1.Panels(1).Width` —— 控件的 Width 定义就是缇）；**2 行 literal 大过 3000**（`Master.frm` 的 3500 与 7000 —— 像素那一档不可能，7000 px 比屏幕还宽）。⇒ 缇这一档现在有 9 处独立站点、跨 7 个工程，加上 ④ 那条与 DPI 无关的单位换算，口径不必再等读数。
- **邻居那两条针的 exact 形状（判据重述要照着这个改，别照数字改）**：`tests/ctrlstatusbar/CtrlStatusBar.frm:74` 的 `SB10-W2=120` 读的是**设计块里那行无前缀的 `.Width = 120`**（第十八刀认的是带前缀的 `Object.Width`，语料真形；这两形在 RTL 里进的是同一格 `e->width`），`CtrlStatusBar.frm:120-121` 的 `SB36-SETW=123` 是**运行期赋值 `Panels(1).Width = 123` 的往返**；同族还有 `SB6-MINW1=40`（设计块 MinWidth）与 `SB35-SETMINW=77`（运行期往返）。⇒ 存储换成缇 + getter 改交排版结果之后这四条都会跟着动，判据按 #157 那条重述：**就地取基线问增量**，别写「读数 = 各步之和」那种手算式。
- **RTL 那一头现在的三个事实**（读码对过，不是推断）：① `Vb6PanelEntry.width/minWidth` 注的就是**像素**，`SbLayout` 把 `want[i]` **原样**递给 `SB_SETPARTS`；② `vb6_StatusBar_GetPanelWidth` 交回 `e->width`（请求值），排版结果只活在 `SB_SETPARTS` 那一份里；③ 设计块那一档经 `vb6_StatusBar_SetPanelWidthHm` → `vb6_HimetricToPxX`（第十八刀）。⇒ 出口两侧要用的那对权威**已经有了** —— `vb6_TwipToX` / `vb6_XToTwipX`（`vb6forms.c:179/217`，#184/#175 收成一处的那对），**别新开一份**；而 hm↔缇 那条与 DPI 无关的比例现在在 RTL 里有**三份写法**（`axsite/ax_load.c:142` 的 `vb6_twipsToHimetric`、`vb6forms_olecon.c:580-581` 的 `*1440L/2540L`、`vb6forms_olecon.c:609/860` 的 `*15L*2540L/1440L`）—— 正是 §B129 那份名单，这一刀要往下减名单，不能加。
- **顺手量到新的一格（同批一起交，别单开一刀）**：`Panels(i).Left` **全仓 0 个出口** —— RTL 的名字表 `kPanelNames`（`vb6forms_memberobj.c:113`）八枚里没有它，发码侧那 14 条 recognizer 也没有它，`grep -rn "PanelLeft" src` 空。而语料里正是那两处按面板定位覆盖层的写法在问它（`modAddProgBar.bas`：`pb.Left = sb.Panels(lPan).Left + 25`、`Form1.frm` 的 `.Move (sb.Panels(nPanel).Left + pading), …`）⇒ 今天读回空。它与 Width 同源（都该由 `SB_GETPARTS` 那份 offsets 折回缇），**所以它就是第二格的同一刀的另一面**，不另立账。

**第二格已出（第二十二刀，2026-10-10，门 #475 = run 38034455442、head `64b4d43e`、attempt 1、11 绿 + 唯一红是这道形状门（红因正是那三行还没登记）；登记之后由门 #476 = run 38035393406、head `de181042`、attempt 1、12 条 check-run 全 completed/success、非绿 0 复验 —— 其中 `Emit manifest (shape oracle)` 绿 ⇒ 那三行的新哈希被 CI 那台独立复算证实，而 `Tests (vbp #1..#4)` 四片全绿 ⇒ sbhm / ctrlstatusbar / sbfont 两架构真跑过）**：单位这一档从头到尾换成**缇**，读侧改问窗口本人。
- **动了什么**：① `Vb6PanelEntry.width/minWidth` 存的换成缇，`SbLayout` 的两道折（固定档 + 弹簧下限）都走 `vb6_TwipToX`；② 新立一枚本文件私有的 `SbPanelEdgesPx`（`SendMessageW(SB_GETPARTS)`），`GetPanelWidth` 与**新增的 `GetPanelLeft`** 都由它折回缇（`vb6_XToTwipX`）—— 唯一真相是窗口那份 offsets，不是 `e->width`；③ 新立一对与 DPI 无关的权威 `vb6_TwipsToHimetric` / `vb6_HimetricToTwips`（`vb6forms.c`，缇↔hm = 1440/2540），两枚 `*Hm` 设计出口改调缇档那一枚，第十八刀那版"把像素存进缇档"的形状（状态条里调 `vb6_HimetricToPxX`）归零；④ `axsite/ax_load.c` 那份旧写法改成**只转调**新权威 ⇒ §B129 名单从 5 份/13 行减到 **4 份/14 行**（`vb6forms.c` 由 1 行长成 3 行，净 -1 份文件）；⑤ 设计块那一行 `MinWidth` 从此走新的 `vb6_StatusBar_SetPanelMinWidthHm`（它同样是 himetric，读数在下面）；⑥ `Panels(i).Left` 那一面（改前全仓 0 个出口）一次登记齐：`kPanelNames` 补 `L"Left"`（DISPID 9）+ `memInvokePanel` 的 case + 发码侧 recognizer 一条。
- **读数**（真跑，x64 与 x86 两份 run.out 逐行相同）：夹具 `tests/sbhm` 的证人从像素换成 `像素 * Screen.TwipsPerPixelX`（**用 Screen 那一条独立路，不用本仓那对权威，否则就是自己证自己**）⇒ HM09..HM12 四条同批 True；三条新判据在 BASE（`b437_base_k21.exe` 跑同一份夹具）上全 False、在改后全 True：`HM13-SPRING-WIDTH-FOLLOWS-WINDOW`（弹簧档从前请求宽恒 0）、`HM14-PANEL-LEFT`（BASE 上还多三条 `property "Left" not found`）、`HM15-DESIGN-MINWIDTH-NOT-RAW`（BASE `minw=26 frm=26` = 设计档原样吐回来，改后 `minw=14`）。邻居夹具按 #157 那条重述：`SB6-MINW1` 从 40 变 **22**（设计块那 40 是 himetric），`SB10-W2=120` 不动（固定档往返，缇→8px→缇），`SB36-SETW=123` 那一形作废 ⇒ 换成 `SB36-RAW before=5880 after=5880 sum=6630 bar=6000` + `SB36-SPRING-FOLLOWS-WINDOW=True`（BASE：`before=0 after=123 sum=243`，判据 False）。
- **邻居第三份夹具的一条读数（没改判据，记下就行）**：`tests/sbfont` 的 `SF01-RAW fix=` 从 **900 变 60** —— 那枚固定档面板在设计块里写的是**无前缀**的 `.Width = 900`（语料里状态条上这一形 **0 处**，是夹具自己造的写法），按缇档就是 900 缇 = 60 像素；三条判据 SF02/SF04/SF05 两台上仍是 True，而 RAW 行钉的是前缀不是数 ⇒ 门不受影响，只有形状门那一行的哈希跟着变（已按 CI 门 #475 那台自己交的清单登记，并由门 #476 的形状门复验（12 格全绿），移动的三行 = ctrlstatusbar / sbfont / sbhm，其余 396 行逐字节没动）。**这一条也正是"无前缀 Width 走缇、无前缀 MinWidth 走 himetric"那条不对称的活样本**：两份邻居夹具各撞上一半。
- **顺手量到、刻意不带的一格（另立 §B132 / 账 #303）**：`sum=6630` 比 `bar=6000` 多出 630 缇 = **42 像素**，正好是那枚 `sbrTime`+`sbrContents` 面板的测量文本宽 —— `SbLayout` 第一遍把内容档的实测宽放进 `want[i]` 却没累加进 `fixedTotal`，于是弹簧把同一格宽吃两次，最右边界越过客户区。**这是改前就在的缺陷**，只是从前 getter 交请求值所以看不见。所以 SB36 那趟只钉"往弹簧档写宽不该改变它报出来的数"这一半（增量不变），"三格铺满"那一半留给 §B132，RAW 那行把 `sum` 与 `bar` 两个数一起打出来，等它落地就看得见。**（已落地 = 第二十三刀，读数与护栏见 §B132 末；`SB36-TILES-BAR` 那条判据就是它补上的）**
- **护栏**：哨兵 `check_statusbar_panel_hm.ps1`（第 51 道 [STATIC]）改形 + 加长：R2 名单钉 4 份/14 行、三枚权威定义各 1 处、`ax_load.c` 只许转调 1 处（它就是它退名单的条件）、状态条三道折各 2 处、本地 `MulDiv` **0** 处、第十八刀那形状 `vb6_HimetricToPxX(` **0** 处（墓碑）；R3 兜底那行改钉 `take < floorPx` 且必须问 `vb6_TwipToX`；R1 加两枚产物计数（`SetPanelMinWidthHm`=3、`GetPanelLeft`=3 —— 少一个落点就是"设计值整格不见"/"读回空"）；**新增 R7**：一枚成员从"名字认得"到"真做得动"要过的六格登记（memberobj 名字表 / DISPID / 派发 case / memberobj 调用 / RTL 定义 / header 声明 / 发码 recognizer）各恰好 1 处 + 设计 `MinWidth` 那三处同步 + 排版那一趟必须折一次（单位用代码形状钉，不靠注释）；R4 加两条 census 防空转。**负控 = 真 BASE 树**：`git archive HEAD` 出一份副本（`.build/wt_k22neg`）+ `b437_base_k21.exe`，`-Root` 指过去一次报出 **25 条**，而第二十/二十一刀的 R1/R5/R6 那几条在同一台上一条没红 ⇒ 新钉的每一条都是本刀的存在性证人，不是重报上一刀的账。51 道 [STATIC] 全跑一遍 = 51/51 绿。
- **本刀自己漏过又被哨兵抓住的一格（记下形状，不是遗漏）**：`vb6_StatusBar_SetPanelMinWidthHm` 的 **header 声明当时没写**，而产物照跑、构建日志一条不报（改前/改后两份 `run.out` 逐行相同 = 这格静默的直接证据：调用返回 void、实参按 ABI 到位，隐式声明在这条链上不出声）⇒ 补声明 + R7-MINW-DECL 钉住。这与账 #174（`SelectedControls` 发成隐式声明）同形，而本仓的构建日志里没有 cl 警告这一族（§B…，已记录），所以**这一类只能靠数落点**。
- **第二十四刀收掉的那两格（原记「刻意留下」，归 §B129 / 账 #298）**：`vb6forms_olecon.c` 那 7 行、`vb6forms_picture_prop.c` 那 3 行、`ax_site_ext.c` 的 `k = 96.0/2540.0` 已并到 `vb6forms.c` 那一处权威上（名单现 1 份 / 5 行，值面与护栏见 §B129 末）；`ax_site_ext` 那条高 DPI 改动本机仍取不到读数（这台就是 96）⇒ 未定罪那一格改记 §B133。另记一条口径不对称（是有意的）：设计块里**带前缀**的 `Object.Width` 与**无前缀**的 `MinWidth` 都按 himetric 折，而**无前缀**的 `Width` 按缇直交 —— 因为语料里状态条上 `Width =` 那形 0 处（是我们自己造的），而 `MinWidth =` 那形是真设计器写的（457 枚里 398 枚与 `Object.Width` 逐字相同）。

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

### B61 数组元素的 extender 属性**写后读回不是请求值**（缇→像素→缇 取整损失），VB6 存的是缇（账 #230，**已出 —— 跟着 §B74 那次并表一起答对；覆盖面由第二十七刀补上，见 §B135**）
> **2026-10-10 复测（把这节从「未开工」改成「已出」的依据）**：拿 CI 那台 `C3.exe`（与 head 同一枚二进制）编一枚控件数组探针真跑，两架构同形 —— `arrT(0)` 的设计值 1007/449/3001/247 原样读回、`arrT(1).Left = 5007` 读回 5007、`Move` 之后四档原样读回 ⇒ 元素那条来路交的就是元素自己的窗口，缓存按 (窗口, 属性名) 存，**跟着账 #230 的并表一起答对了**。本节原来那句「未开工」是 #230 出完之后没回头改的标题；真正欠的是**覆盖面**（全门禁没有一条针向数组元素要过几何），由第二十七刀补上 ⇒ 见 §B135。探针 `.build/b590_b61/PB61.frm`。


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

### B70 窗体绘图语句的"裸形"与 "Me." 形各缺一条路（账 #232，**② 已出：门 #368（run 37454147931、head `9e5c9f91`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m46s；**① 已出：门 #374（run 37533797244、head `2ef2ee71`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，含新的 [CODEGEN-NOTE] form_canvas_bare 所在的 Tests (syntax) 片与跑 FD21/FD22/FD23 的 vbp 四片**；**③ 已出：门 #375（run 37540458054、head `129358c4`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，含新针 form_canvas_tail 所在的 Tests (syntax) 片与跑 FD21..FD25 的 vbp 四片**）

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
**① 已出（门 #374（run 37533797244、head `2ef2ee71`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，含新的 [CODEGEN-NOTE] form_canvas_bare 所在的 Tests (syntax) 片与跑 FD21/FD22/FD23 的 vbp 四片）—— 三条改动、三面判据、外加一道补刀（旧哨兵那条断言是假绿）**

**读数（同一份夹具、两台编译器的 `--emit-c`，前后只差两行）**：改前 `Cls` 发 `Cls();`、`PSet (170, 130)` 发 `PSet(170, 130);`，而 stderr 还多一条 VB3001「未声明的标识符 'Cls'」；改后这两行是 `vb6_ControlCls((void*)vb6_hwnd_FDForm); /* Form.Cls */` 与 `vb6_Form_PSet((void*)vb6_hwnd_FDForm, 0, 1, 170, 130, 0, 0); /* Form.PSet */`。真编译那一面更硬：基线台两架构都死在 `LNK2019 无法解析的外部函数 Cls（在 vb6_tmrF_Timer 中被引用）` + 同族那条 `PSet` + `LNK1120` ⇒ **压根没有 exe**。所以这一格不是"跑起来一笔不画"那一族，是"编不过"那一族。

**六形实测（探针 `.build/b494_probe2/ShForm.frm`：一份 .frm 把裸形全摆上，`--emit-c` 逐行读）**：`Cls` / `PSet (5, 6)` / `PSet 12, 13`（无括号那形）/ `Circle 30, 31, 32` / `Circle 70, 71, 20, 255`（带颜色那形，hasColor 折成 1）/ `Point (40, 41)` 六形都折成真出口；**只剩「括号坐标对 + 逗号尾巴」那一形没折** —— `Circle (20, 21), 22` 仍发 `Circle(20, 21, vb6_VariantEmpty(), 22);`，`PSet (60, 61), 255` 同样漏（发 `PSet(60, 61, vb6_VariantEmpty(), 255);`）。

**下一格的根因（本轮量清；上面先前那句「callee 树多套一层」是假说，已被读数推翻 —— 扁平逗号形六形全通，
就证明折叠器认得这一形；漏的不是折叠，是 parser 没把尾巴收进来）**：
`src/parser/parser_expr_postfix.cpp:162-181` 那趟 Circle/PSet 的逗号尾巴吸收，判据是 `call->callee->kind == MemberAccessExpr` 再去取 `memberName`；裸写的 callee 是 `IdentifierExpr` ⇒ `isCircleCall`/`isPSetCall` 恒假 ⇒ `, r` 漏到外层表达式，剩下的 `(...)` 桩由发码的可选形参补成 `vb6_VariantEmpty()`。
同一文件 183 行往下的 `-(x, y)` 那趟（`Line` 的两点形）**同样是 MemberAccessExpr-only** ⇒ ③ 的裸 `Line (0,0)-(10,10)` 是同一刀的第二头；再加 `src/parser/stmt/parser_stmt.cpp:136` 那一格 `TokenKind::Line`（今天只认 `Line Input`，裸 `Line (0,0)` 走 `parseLabelOrAssignmentOrCall` 就报 `expected ')'`）。

**这一格要的形状（照 ①-c 的口径，别单开第五份名单）**：两处吸收都改成「callee 是 MemberAccessExpr **或** IdentifierExpr，
名字问 `canvas_drawing.hpp`」—— 这两处现在硬编码的 `circle` / `pset` / `line` 是画布动词名单的**第四份副本**，
改完由哨兵跟着 census；裸 `Line` 的语句形要把 `TokenKind::Line` 折成「一枚叫 Line 的裸调用」再交给语义层那把折叠，
而不是再拼一条专用发码头。判据两头：六形 + `Line` 两形各一条发码针（present 真出口 / absent 裸名），
负控仍用「同一对编译器 + 带裸形的夹具」那枚能红的对照。覆盖面本轮再扫一遍全树：`tests/` 里 `Line (` / `Circle (` / `PSet (` 的裸形只有本刀自己的夹具，其余全是 `Line Input` ⇒ 这一格是**潜伏缺陷 + 收口完整性**，
优先级排在任何「真语料编不过 / 跑坏」的账之后。

**收成一处（新增 common + 三处码头改问它）**：`src/common/canvas_drawing.hpp` = 15 行（动词, 接收者）矩阵 + owner 两档（METHOD / DRAW）+ 四个 accessor，与 `host_pseudo.hpp` / `float_literal.hpp` / `int_literal.hpp` 同一层（common 不向上依赖 semantics）。`controlCanvasMethod` 退成"把 ctrlType 换成接收者位再问表"；withm 绘图码头那段 `isPrinterDraw` 的五条名字硬名单改成 `canvasDrawEntry(名, 接收者)`；打标记那一路的 enddoc 也改问表。**折叠本身住在语义层一条语句的位置**：`visit(CallStmt&)` 里 `analyzeExpr` 之前调 `foldBareCanvasVerb`，五道门 = pass 2 / docKind==Form / callee 是裸 IdentifierExpr 或 IndexOrCallExpr 的内层那枚 / 名字在那张表里有 Form 档 / 三问查不到符号（`symTab_`、工程级、祖先）。折完下游一行没动 —— `Me.Cls` 走 ② 那张表，`Me.PSet` 走 Form/Printer 段。

**判据三面**：① 真跑 `tests/fdraw` 加 FD21（墨问两次：画完 `>0`、裸 `Cls` 之后 `<0`，读数 `raw=3,-1`）与 FD22（`Me.ForeColor = vbBlue` 之后裸 `PSet (170, 130)`，问得到那一条蓝行 = `raw=130`），两架构 24 行输出全对、stderr 0 字节。② **边界证人** FD23：模块里自己写了 `Private Sub Circle(x, y)` 时那枚过程**该赢** —— `Circle 7, 9` 发成 `vb6_Circle((&(int32_t){7}), (&(int32_t){9}))`、读数 `FD23-usercircle=1`；上一段那条"三问查不到符号"的门不是装饰，这一枚就是它。③ 发码针 `Test-CodegenNote "form_canvas_bare"`：present 三枚（两条真出口 + 那枚 `vb6_Circle`）、absent 四枚（`Cls();` / `PSet(170, 130);` / `VB3001` / `vb6_Form_Circle((void*)vb6_hwnd_FDForm`）；基线台实测「缺 1 枚 + 命中 3 枚」⇒ 真会红。present/absent 里那对边界针在两台上都绿 —— 它钉的是"不许过折"，不是这一刀的形状。

**补刀 = 上一版那条断言是假绿**：S9.2 只查"每个动词名在不在表里"，于是把 `{"point",  CR_CANVAS_FORM, …}` 整行删掉它 **rc=0 一声不响**（负控实测）。而"只接了一半接收者"恰恰是 ② 那一轮的定义。现在换成（动词 × 接收者）15 对矩阵 + 行数恰好 15：删同一行 ⇒ rc=1 两条（`has 14 rows, want exactly 15` 与 `no longer pairs point with CR_CANVAS_FORM`），恢复 ⇒ rc=0。另两条假改动（码头自己拼名字 / backend 不再问表）上一轮已各证红。

**护栏**：语料 A/B 用**这一对编译器**重捕 90 份 = `inputs=90 changed=0`；而同一对编译器在带裸形的 `tests/fdraw` 上差 2 行 ⇒ 这对编译器确实有区分力，那个 0 是语料真没有裸形（与本账早先"裸 `Cls` / `PSet` / `Circle` 语料 0 处"的扫描对上）。`scripts/check_form_draw_state.ps1` rc=0；`run_tests.ps1` ParserError=0。

**边界**：Printer 那一族的名字这轮一起搬进表（`pset` / `circle` / `point` / `line` / `enddoc` 各有一行 PRINTER），但 `Printer.Cls` = 结束文档那层语义只在表里挂一行 DRAW，没有新造第二种"清画布"。`With pic : .Cls` 照上一轮的记录押后；③ 裸 `Line` 两形照旧（parser 就不认）。

**③ 已出（门 #375（run 37540458054、head `129358c4`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0，含新针 form_canvas_tail 所在的 Tests (syntax) 片与跑 FD21..FD25 的 vbp 四片）—— 裸写动词的尾巴改由 parser 收进来；名单只剩表那一处**

**读数（两台都在本机编的编译器、同一份工作树文件）**：改前那台（HEAD `2ef2ee71` 在本机冷编，`.build/wt_base374`）对 `tests/fdraw/FDemo.vbp` 做 `--emit-c` 直接 **exit 1**，三条错误全落在那一枚裸写的 `Line (300, 40)-(300, 120)` 上（`FDForm.frm(253,24): expected ')' (got ,)` + VB2003 + VB2002 —— 一行崩掉整段过程）；同一台在探针 `.build/b494_probe2/ShApp.vbp` 上进得去，发出来的是 `Circle(20, 21, vb6_VariantEmpty(), 22);` 与 `PSet(60, 61, vb6_VariantEmpty(), 255);` —— 逗号之后的实参**看着在**，其实是尾巴漏到外层之后由可选形参补齐拼出来的，整条调用仍是未声明的裸名（C2065/C2064 那一族）。改后这一族六形 + Line 两点形全落真出口，夹具两架构 26 行输出逐字相同、stderr 0 字节：`FD24-barecircle=True raw=-1,59`、`FD25-bareline=True raw=-1,38`。

**因由与修法（一句话）**：`src/parser/parser_expr_postfix.cpp:162-181`（Circle/PSet 的逗号尾巴）与 183 行往下的 `-(x, y)`（Line 的第二点）两趟吸收都把判据写死成 `call->callee->kind == MemberAccessExpr` ⇒ 只认带接收者那一形（`Me.Circle` / `Pic1.Line`），裸写的 callee 是 `IdentifierExpr` 就收不到尾巴；而且这两处还各自把 `circle` / `pset` / `line` 三个名字**自己拼成名单** —— 那是画布动词名单的**第四份副本**。修法 = 尾巴形状那一档收进 `src/common/canvas_drawing.hpp`（`canvasVerbTailKind`：NONE / COORD / TWO_POINT，一问一处答），parser 两处改成「callee 是 MemberAccessExpr **或** IdentifierExpr，名字问表」。语义层那把折叠（账 #232①）一行没动 —— 折成 `Me.<动词>` 之后就走 ①② 已经收好的两条码头。

**判据三面 + 一次被迫的改名**：① FD24 / FD25 各两头钉（同一支颜色先扫到空、画完扫得到；颜色用这枚窗体别处不出现的 vbGreen / vbMagenta，免得拿别人的墨当证人）。② 发码针新增 `form_canvas_tail`（present 三枚真出口；absent 两枚 = `vb6_VariantEmpty()` 与 `Circle(240, 90`，钉住"尾巴没漏成补齐桩"）。③ 边界证人 **从 Sub Circle 改成 Sub Point** —— Circle 现在要留给 FD24 真画一个圆，模块级同名过程会把它吃掉；这条改名不是将就，它正是那条边界自己在说话（折叠只在名字查不到符号时才动）。`form_canvas_bare` 的那对边界针跟着换成 `vb6_Point(...)` / `vb6_Form_Point(...)`。红侧两头都有实物：HEAD 那台对这份夹具 exit 1（`Test-CodegenNote` 要求 rc=0 ⇒ 必红），而探针里那两条裸名是它逐字存在的产物。

**哨兵 S10 三条**（`scripts/check_form_draw_state.ps1`）：parser 必须 include 那张表、必须问 `canvasVerbTailKind(`（**恰好一次** —— 一问两用，两处 gate 读同一个答案）、两处都必须 gate 在 `tailKind == CANVAS_TAIL_*`、并且不许再出现 `== "circle"` / `== "pset"` / `== "line"` 这种自己拼的名单；表那一头三档尾巴齐。两条假改动各证红：parser 里补一句 `|| tailName == "circle"` ⇒ rc=1 点名"fourth copy"；表里删掉 `line` 那一行 ⇒ rc=1 两条（`no longer maps "line"` + `no longer maps CANVAS_TAIL_TWO_POINT`）；恢复 ⇒ rc=0。

**护栏**：语料 A/B `inputs=90 changed=0` —— BASE 是 **HEAD 那笔在本机冷编的那台**（同一套 VS2019 工具链，`.build/wt_base374`），不是我第一版误用的 CI 工件（那一版读出 `changed=14`、1068 行差异，全是晚绑定 COM 读面的 `vb6_ComGetIntProp` ↔ `vb6_VariantFromComResult(vb6_ComGetProp)` 一族、画布动词一行都没有 ⇒ 记进 §B73）；pcline（唯一用带接收者 `Line` 两点形的夹具）emit 前后逐行相同、四枚 line 出口不动；27 道 [STATIC] 全 rc=0；`-Category syntax` 167 例 0 红。

**边界与下一格**：`Line -(x, y)` 与 `Line Step (x, y)-(x2, y2)` 两形仍不通 —— 读数是 `LnForm.frm` 上 `Line Step (5, 5)-(15, 15)` 报 `expected 'Input' after 'Line'`。它们的堵点在**语句层**：`src/parser/stmt/parser_stmt.cpp:136` 那一格 `TokenKind::Line` 今天只把「后面跟 `=` / `(` / `.`」交回调用路，跟 `-` / `Step` 就去做 `Line Input`。要接得先定两件事：(a) 把裸 `Line` 交下来成一枚 `IdentifierExpr` 调用（折叠与尾巴吸收都已就位）；(b) VB6 那句"从 CurrentX/CurrentY 起笔"的起点口径怎么在 `vb6_Form_Line` 的 12 参里表达（现在那两格是 step 旗标 + 坐标，没有"取当前点"这一档）⇒ 口径要用户拍，别自己定。

### B71 门 #369/370 那两条红只有 runner 上现形 —— 本机那台 cl 压根不诊断「实参过多」（账 #240，**已出：门 #371（run 37471108271、head `bb252705`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0；红过的 olecon / olecon_x86 两片转绿，新的 [STATIC] rtl_proto_arity 跟着 Tests (compile) 一起跑绿；dev 已镜像到 gitcode（origin/dev 由 `39d9b119` 快进到 `bb252705`）**）

**读数**：门 #369（run 37456420314、head `39d9b119`、branch dev）11 job 里两片红，各红一条且是同一枚夹具的两个架构 —— `Tests (vbp #2)` = `[VBP-BUILD] olecon ... FAIL rc=1 exe=False`（该片 PASS=54 FAIL=1 SKIP=1）、`Tests (vbp #3)` = `olecon_x86`（PASS=53 FAIL=1 SKIP=0）；其余九片全绿。**引入方式不是改了产品**：`39d9b119` 那轮新增 [STATIC] vbp_fixture_census 把五份"跟踪着却没登记"的 .vbp 逼出册登记成编译面用例，olecon 是其中一份（提交说明里写着本地 x64+x86 rc=0 且出 exe）。同批登记的 dbgdlg（就是那枚缺 `vb6_di_PageSetupDlgA` 桩、为它才补的夹具）在两片上都 PASS ⇒ 桩表与 RTL 内嵌在 CI 上是对上的，红只跟着 olecon 走。

**已排除的六条**（每条都有实物，不是推理）：(1) 夹具没进仓 —— `git ls-files tests/olecon` 有 .frm+.vbp 两份，且目录里根本没有 .frx（那条 `oc_src.bin` 只在运行期读）；(2) CI 那台 C3.exe 与我本地这台不同 —— 把 run #369 的 `c3-exe` 工件下载下来真跑，x64 与 x86 都 rc=0 出 exe；(3) `-Incremental` —— `Test-VbpBuild` 压根不传 `--incremental`；(4) 架构/命令行差异 —— 本地按登记时的两条命令行（默认 x64 与 `--arch x86`）逐字复跑；(5) RTL 内嵌资源 id 对调（账 #225 那一族）—— 那样会全线 LNK2005×1225，不会只有一条红；(6) 源码本身依赖注册表里的 VB6 类型库 —— olecon 走的是仓内原生那一条（`driver_link.cpp:45` 每次都带 `vb6forms_olecon.c`，产物里全是 `vb6_OleCon_*`/`vb6_RegisterOleConClass`，没有查注册表的路）。

**剩下的唯一差异是机器**：runner 用 vswhere -latest（VS2022 + 新 SDK），本机只有 VS2019 14.29.30133 + SDK 10.0.19041。红出现在 `runLinker` 那一段（只有那一支才打 `intermediates kept at`），而 cl/link 的整段输出只落在 `c3-error.log` —— 它既不被 `Test-VbpBuild` 打印，也不在 CI 的工件通配符（`output/**/*.out|*.err|*.txt|*.dat|scores.txt`）里 ⇒ **门上看不见病因**。

**本刀（一）**（`tests/run_tests.ps1`，+23/-0，纯测试面）：`Show-BuildErrorLog` 接进 `Test-VbpBuild` 的失败分支，挑 `error C####` / `: error ` / `fatal error` / `LNK####` / `unresolved external` / `=== C3 Diagnostics` 那几行（最多 25），一条都不匹配时退回尾巴 15 行。判据形状来自本地实物：一枚刻意失败的工程（b475）日志 264 行，262 行是 RTL 的 C4819/C5105/C4028 警告，直接摊 40 行尾巴会把唯一的 `error C2063` 挤出去。

**归因（门 #370 的 [diag] 读数，两片各 4 行，一模一样）**：`Form1.c(100): error C2197: 'void vb6_OleCon_Init(void *,const wchar_t *,int,int,int,int)': too many arguments for call` —— **4 行 = 10 个实参减 6 个形参**，一枚多余实参报一行。对着源码量：定义 `src/rtl/core/vb6forms/vb6forms_olecon.c:920` 是 10 参（`... autoActivate, autoVerbMenu, borderStyle, sourceDoc, sourceItem`），发码 `src/backend/detail/module/cgen_form_ctrl_style_apply.inc:970` 也发 10 个，只有 `vb6forms_prop_ctrl.h:334` 那份原型还停在 6 参 —— 体 grew 上去、头没跟。**为什么只有 runner 红**：VS2019 的 cl 在 C 模式下对「实参多于原型」根本不诊断（本机用 6 参原型 + 10 实参的最小夹具 `b481/t10.c` 实测 rc=0，只给一句 `warning C4020:太多实际参数`，连错误都不升），新 cl 把它按 C 标准的约束报成 error ⇒ 本机真编两遍都编不出这个病，判据必须换形状。

**这条"本机看不见"后来量准了，而且两个方向不对称**（b494/arity.c，同一台 14.29、同一套 `/std:c11 /W3`）：6 参原型 + 10 实参 ⇒ 只有 `warning C4020`，编译照过（订正：本段先前写的 C4024 是凭印象，C4024 是"形参与实参类型不一致"，不是这一形）；反过来 10 参原型 + 6 实参 ⇒ `error C2198:太少实际参数`，本机就红。不对称有两层用：① 它解释了为什么"多递实参"只能靠对源码比来防（真编防不住）；② 它也划出了 §B72 的范围 —— 只需盯"多递"那一侧，"少递"已经被今天任何一次真编兜住。

**修法（本刀二）**：`vb6forms_prop_ctrl.h` 的原型补齐成 10 参（+3/-1，只动头；`src/rtl/**` 改了要 touch `src/driver/c3rtl.rc` 再重编 C3.exe，否则内嵌的还是旧字节 —— 账 #156 那条）。运行面零改动：10 个实参本来就一直发着，本机那台把多余 4 枚照 cdecl 传过去了，所以旧产物行为不变；这一刀只是让**下一台编译器**也认。本地验：新 C3.exe 真编 olecon `=== x64 rc=0 OleCon.exe 465920 字节 / === x86 rc=0 OleCon.exe 414720 字节`（这两个数与登记那轮记录逐字对上）。

**本刀（三）= 结构性哨兵**`scripts/check_rtl_proto_arity.ps1`（第 27 道 —— 本轮之前实测 26 份 check_*.ps1）+ `run_tests.ps1` 的 `[STATIC] rtl_proto_arity`：判据 = RTL 里**两头都有**的名字，声明侧参数个数集合必须等于定义侧（R1），外加 census 地板「文件数 >= 100 且比较对数 >= 900」（R2，实测 127 份 / 1237 对）—— 路径写错或正则被改坏时不许变成"绿着的空转"。两条设计约束记下：① **只有第 0 列开始的行算签名**，这一条同时把所有调用点排干净（调用都缩进在函数体里），② 只比**个数**不比类型拼写（头写 `const X*`、体写 `X*` 是合法的，硬比只造噪声），认不出的参数形态（数组/函数指针/默认值）跳过、不计入也不报红。**负控两头跑过**：把头削回 4 参 ⇒ rc=1 并点名 `vb6_OleCon_Init: 头 4 参 (…prop_ctrl.h:334) 对不上 体 10 参 (…vb6forms_olecon.c:920)`；改回 10 参 ⇒ rc=0（同一份 census 读数 1237 对）。

**这一轮为什么不跑语料 A/B**：本刀只动一份 RTL 头，而 `--emit-c` 的产物里压根没有 RTL（记忆里那条老读数），C++ 源一行没动 ⇒ 编译器的发码逻辑同一个程序，只有内嵌的 RCDATA 变了。90 份 emit 的 BASE 那台已被覆盖，拿它比只会量到这一个月的**夹具漂移**（账 #239 那条教训：改了夹具再跑 A/B = 假归因），所以换成正对靶子的三面：
① 发码实物 = 10 个实参（本地 keep-for-debug 的 `Form1.c:100` 逐字读过）；② 定义 = 10 参；③ 头补齐后 = 10 参 —— 三头同值，再加两架构真编 rc=0 出 exe。**这一族的边界（记下别越界）**：哨兵管"头追不上体"，管不到"发码递的实参个数 ≠ 体"。真要钉那一头得让**每个控件方法的发码形状**与 RTL 原型对账，那是把 `controlOneArgMethod` / `controlZeroArgMethod` 那几张表的签名也拖进对账面的一件大活（且只有新 cl 才看得见后果）—— 已另立 §B72 记着，本轮不顺手做。




### B72 控件方法的"发码实参个数"与 RTL 原型之间没有对账 —— 头/体那一半已钉住，这半只欠"多递"那一侧（账 #240 记下；**三张控件方法表 = 第十二刀 §B122，Winsock 那一族 = 第十三刀 §B123，都已出**；集合 Clear 两枚与 CommonDialog Show* 六枚 = 第十四刀 §B125；剩下的 Data 一族经查整户是死码：发码侧 = 第十五刀 §B126，RTL 那两枚 0 调用者的出口 = 第十六刀 §B127，§B124 记着当时的探针与那份闭合清单）

账 #240 那一刀把"头追不上体"钉死了（`check_rtl_proto_arity.ps1`：RTL 里两头都有的名字，声明侧参数个数集合必须等于定义侧）。**没钉住的是第三头**：cgen 递出去的实参个数。本轮三者恰好同源（体 10 = 发码 10 = 补完的头 10），所以新 cl 的诊断只落在头那一份上。反过来的形状照样要命：若有人给某枚 `vb6_Ctrl_*` 加形参、只改头与体，而**发码仍递旧的个数**，那么多递这一侧在本机只是 `warning C4020`（见 §B71 那条不对称测量），到新 cl 才升成 `error C2197` —— 也就是说它会以**"门红、本地全绿"的形状再来一次，而这次红在别人刚登记的用例上**。

要收的形状（照 §B70 的 ①-c 那张表的做法，别再单开第四份名单）：控件方法的"名字 → RTL 出口 + 实参形状"本来就该只有一处答案。现在 `controlZeroArgMethod` / `controlOneArgMethod` / `controlCanvasMethod` 三张表里**没有"这一档递几枚实参"这一格** —— 那个信息住在调用点（各码头自己拼参数串）。所以第一步是把实参个数写成表里的一个取值，第二步才是同一条针两面都问：拿表里的出口名去 RTL 头里查参数个数，与表里的形状对。第一步与 §B70 工单的第 (1) 步是同一件活（画布动词那张 `src/common/canvas_drawing.hpp`），所以这一格**排在 ① 之后做**，不要为它先立一张只有旗标没有形状的表。

> **2026-10-10 开工前先看这条（第十一刀之后顺手量的）**：这条账的「第三头」之外又数出一格同一族的现形 —— **零实参那三族的名字被抄了两遍**：表 `controlZeroArgMethod`（`cgen_util_ctrl.cpp:1494`，答 `vb6_SetControlFocus` / `vb6_Slider_ClearSel` / `vb6_ClearList`）之外，语句码头 `cgen_call.cpp:373-384` 把 `ListBox/ComboBox + clear` 那一行**又硬编码了一遍** `vb6_ClearList(...)`（同一段往上 356-358 还有 `vb6_Ws_Close/Listen/Connect` 三枚名字也是就地写）。⇒ §B72 的第一步应该是**先让所有码头都问那三张表**（名字的唯一住所），第二步才是把实参个数写进表里、让哨兵两面都问（表 ↔ RTL 原型 ↔ 发码点）。顺序别倒： arity 进了表而发码点还各抄各的名字，等于给两份名单各自配一个数。

### B73 同一笔提交，CI 编出来的 C3.exe 与本机编的那枚，发码不一样（账 #232③ 撞见，**已结案（2026-10-09）：真凶 = `ComMemberInfo::returnType` 没有默认初值，账 #245 已修；本轮两台工具链 × 两架构 × 全语料 796 份捕获逐字节相同；第 44 道哨兵把它钉住 = **门 #439**（run 数 439、head `faa37573`、branch dev、attempt 1）= 12 条 check-run 全 completed/success、非绿 0，含新的 `[STATIC] com_sig_field_defaults` 所在的 `Tests (compile)` 片**）

**读数（三方对跑 `--emit-c`，同一份工作树文件、同一个 cwd）**：① 本机冷编的 HEAD（`2ef2ee71`）与本机当前树（多 ③ 那一刀）⇒ `inputs=90 changed=0`。② 拿门 #374 的工件当 BASE（`gh run download 37533797244 -n c3-exe`，CI 的 Build job 用 vswhere -latest = VS2022）对同一台本机 HEAD 树 ⇒ **`changed=14`，1068 行差异里 1060 行是同一族**：`vb6_ComGetIntProp(oFont, L"Name")` ↔ `vb6_VariantToString(vb6_VariantFromComResult(vb6_ComGetProp(...)))`、`vb6_ComCallInt(...)` ↔ `vb6_ComVarFree((void*)vb6_ComCall(...))`，另有 `ComGetDouble` / `ComGetObject` / `ComGetBool` 同形；落在 Charts 2020 六份子工程（`IAFPService` / `ITilterAccess` / `IMyCompany` 那几枚晚绑定对象）与 VBFlexGridDemo 的 `PropFont` 读面上。**画布动词一行都没有** ⇒ 与 ③ 那刀无关（逐条 grep 过）。③ 再钉一颗反向钉子：本机 `847ee9f4` 那台（更早两笔提交）与本机 `2ef2ee71` 那台在这一族上**逐字相同**（`vb6_VariantFromComResult` 各 70 处），CI 那台是 66 处 —— 所以这不是"少一笔提交"，是**同一份源码在两台构建机上做出不同决定**。

**为什么最可疑的是"顺序"而不是 `#if`**：`vb6_ComGetIntProp` 这套带类型 getter 在本树里到处都在（`cgen_util_com.cpp` 六处、十个文件提到），两台都编得出来，只是**在哪些调用点上选它**不同 —— 这正是"一排识别器轮询、第一条命中就答"的形状。若那条队列走的是无序容器（`unordered_map` / 指针序 / 静态初始化序），MSVC 两版换哈希实现就会换首命中 ⇒ 决定随构建机变，而**门不会报**：Build 与 Tests 用的是同一枚工件，自洽。两条便宜的分辨办法：(a) 本机同一笔提交冷编两次跑同一份语料，两边逐字相同就排除"真随机"，只剩工具链；(b) 把发码里所有「遍历容器取第一个命中」的识别器列出来（`grep -n "unordered_" src/backend src/semantics`），凡答案会随遍历序变的，换成有序容器或换成显式优先级 —— 与本仓既有的定死做法一致（单一权威 + census + 结构性哨兵）。两条同族的旧账可以拿来对照：门 #162（CI 上关掉的 Timer 仍报 21 拍，本地 1 拍，归因未定）与 MV-d（本地红、CI 绿，最后是挂钟）——**"CI 与本地不同"这一族已经有三种因由，这是第四种，而且它差在发码面上**，比运行时那几种更硬。

**这一格今天就能用的口径**：**语料 A/B 的 BASE 必须与 NEW 同一工具链、同一构建方式**（本机冷编 worktree：`git worktree add .build/wt_baseX <sha>` + `.build/wt_pre_build.ps1 -Src ...`，实测约 6 分钟）；**CI 工件只能用来跑测试，不能当发码基线**。反过来也成立：凡是「CI 红、本地绿」且差异落在发码形状针上的门，先怀疑这一格，再怀疑自己的刀。

**2026-10-07 复现并收紧了变量（工件 = 门 #380 的 `c3-exe`，run 37556525156、head `09621c38`，与本机那枚同一笔提交；探针 `.build/b556*`、`.build/b557probe/`、`.build/b560_*`、`.build/b559_*`）**

- **同一台机器、同一份源码、同一笔提交**：把 CI 那枚 C3.exe 下载到本机跑同一份语料（90 份 `--emit-c` 捕获）对本机冷编的那枚 ⇒ `same=76 / changed=14 / 两侧各 534 行`。两边读同一份注册表、同一批类型库、同一份源文件 ⇒ **唯一的变量就是那枚 exe**。
- **每台自己再跑一遍逐字节相同**（本机 `b560_local1/2`、CI `b560_ci1/2`，diff 0 行），而对 `--dump-symbols` 的输出**两份二进制完全一致**（同一工程 166,063 字节，diff 0 行）⇒ 语义层的可见状态一样，分家在**发码期那一步**；也不是每次运行会抖，是**决定随构建而变、构建内自洽**。
- 形状（14 份捕获全是这两族）：CI = `vb6_ComGetIntProp(oFont, "Name")` / `vb6_ComCallInt((*New_Font), "Charset", …, 1)`；本机 = `vb6_VariantToString(vb6_VariantFromComResult(vb6_ComGetProp(oFont, "Name")))` / `vb6_ComVarFree((void*)vb6_ComCall(…))`。⇒ 差的不是行号而是**同一个成员由谁答**：一头进 `cgen_util_com.cpp:296/731` 那条前期绑定支路（拿 `comMethods` 的 `returnType` 选带类型的 getter），另一头根本没进那条支路。**后果不是两种写法等价**：`Name` 是 BSTR，被 `vb6_ComGetIntProp` 交出来再喂 `vb6_StrPtr(…)`（GDI+ 那三处）就是崩/错值那一形。
- **本机这一枚的答案根本不出自那张签名表**（这条推翻了上面"未初始化 `returnType`"那一条嫌疑的位置）：本机交的是 `vb6_VariantToString(vb6_VariantFromComResult(vb6_ComGetProp(oFont, L"Name")))` —— 这是**晚绑定**那一路的发码，而 `cgen_util_com.cpp:296/731` 的前期绑定支路里**没有**这一形（它的兜底是 `ComGetStringProp` / `VariantFromComResult(vb6_ComCall(...))`）⇒ 本机压根没进那条支路（`isEarlyBoundCom_ && earlyBoundSym_` 这一对没成立，或 `comMethods.find()` 没命中）。CI 那枚进了。⇒ **分家点在"这枚变量有没有被登记成前期绑定的 COM 变量"那一层**（`cgen_decl_var.cpp:261-283` / `cgen_decl_func.cpp:169` / `cgen_decl_prop.cpp:141` 填 `knownTypedComVars_` 那三处，问的是 `lookupModuleDotted(类型名)` 有没有拿到一枚 ComClass/ComInterface 符号），不在签名表的内容里。三枚最小探针（含照抄语料形状的 `P245c.cls`）两台上都交通用那一路 ⇒ **裸 .cls 里 `StdFont` 连 `comMethods` 都没有**，语料那一层多出来的东西是**工程级**的（`.vbp` 的 `Reference=` 与类型库自动加载注进每个模块那一路）。
- **下一轮的两步（按顺序，别跳）**：① 把 `P245c.cls` 装进一枚带 `Reference=` 的 .vbp（照抄语料 .vbp 里那几行引用），在两台上各 emit 一次 —— 复现了就说明触发条件是引用/自动加载；没复现就回去删语料工程的文件表（`ucChartArea/Proyecto1.vbp` 一条条摘 `.ctl`）直到剩最小集合。② 复现之后在两台上打一条 stderr 诊断（`C3:` 开头的信息面，note 级在成功编译里不打印），报`lookupModuleDotted("stdfont")` 拿到的是什么 kind、`comMethods` 里有没有 `name` 这一档 —— 看在哪一步分家，再定修法。判据用已经建好的对跑机器：`.build/b556_diff.py`（CI 工件 vs 本机），改完 **changed=14 → 0** 才算证死。
- **两条支撑这条嫌疑的旁证**：① CI 那枚对 `Name`（BSTR 属性）交的是 `vb6_ComGetIntProp` —— 这**不是任何正确导入会给的答案**（`IFont.get_Name` 是 `[out, retval] BSTR*` ⇒ 该走 `GetStringProp`），所以那一格的答案不是从签名来的；② 本机这一格压根没进前期绑定支路（`find` 没命中或标记没置）⇒ **同一张 `comMethods` 表在两台里内容不同**。两者都不像"从签名算出来的答案"，而这条嫌疑的落点也因此从「签名表的内容」挪到了「这枚变量有没有被登记」（见上一条）。
- **已排掉的三条**：① 与画布动词那一刀无关（`vb6_Form_(Cls|Print|PSet|Line|Circle)` 零差），本轮新加的 `geom-cache` / `flt-to-int` 两族也零差；② `src/com/typelib_parser*.cpp` 的循环全按 COM 数组下标或 vector 走 ⇒ 上一轮猜的「导入侧遍历无序容器取首命中」不成立；③ **三枚最小探针都不复现**（`.build/b557probe/`：`P245.cls` 类字段 `As StdFont` 读 `.Name`/`.Size` + 一枚 `ByRef … As StdFont` 形参；`P245b.cls` 把同样的成员读放进 Declare 实参里 `GdipCreateFontFamilyFromName(StrPtr(m_Font.Name), …)` 与 `MulDiv(m_Font.Size, …)`；`P245c.cls` 干脆照抄语料那一形 —— `Dim m_TitleFont As StdFont` + `Public Property Set TitleFont(New_Font As StdFont)` 里 `With m_TitleFont / .Name = New_Font.Name / .Charset = New_Font.Charset` + `ByVal oFont As StdFont` 形参）：三枚在两台上**逐字节相同**，而且都交通用那一路（`vb6_ComGetProp(oFont, L"Name")`，连前期绑定支路都没进）。⇒ 单文件探针里 `StdFont` 那枚符号根本没带 `comMethods`；语料那两处住在带 `Reference=` 的 `.ctl` 工程里 ⇒ **下一轮从「把 P245c 装进一个带引用的 .vbp」起手**，而不是继续删语料的文件。
- **顺手记下的一条独立缺陷（与 #245 脱钩，但照样该修）**：`typelib_parser.hpp:44-56` 那三枚枚举字段（`ComMemberInfo::returnType`、`ComMemberInfo::kind`、`ComParamInfo::type` / `::direction`）**没有默认初值**，而 `typelib_parser_desc.cpp:134-163` 的 `parseVarDesc` 只在 `varkind == VAR_PERINSTANCE` 那一支赋 `returnType`，其余支交回的是**没赋值的栈对象字段**；这些成员再经 `insertComMethod` 落进 `comMethods`，且 `parseVarDesc` 把 `kind` 一律写成 `PropertyGet` ⇒ 账 #219 那条「setter 不许盖 getter」的保护挡不住它。这一格按它自己的规矩收（补默认初值 + 每支都赋值 + 哨兵钉「结构体数据成员必须有默认初值」），**判据 = 本机改前改后 90 份捕获逐字节相同**（本轮已证本机这三枚字段对语料的发码没有影响）；同族前例 #172、#238。

**重测（2026-10-08，#260 与 #258 都已进之后）—— 那 14 份没塌，只把嫌疑清单削短了一格**。口径比上面那两轮宽（`.vbp` **加** `.bas` × 两架构，排除比该头新的 `tests/setobj`）：门 #401 的 CI 工件 `c3-exe`（md5 `51d4f14c`）对本机**同头**冷编那台（`aeb883a9`，md5 `441577c6`）= **captures=786 / same=772 / changed 仍是 14**（名单与上面那两轮逐份相同），差异行 1516 条按形归类 **OTHER=0**（typed-IntProp 746 / generic-ComCall 526 / generic-String 184 / generic-GetProp 40 / 其余 22）。

**新增的那一条**：两侧现在**都**带着 #260 那层「解开右值」的形状（`vb6_ComPackValue(<解开的读法>)`），所以「With 那一支没消费标记」这一族嫌疑在此出局，剩下的只差在**定档那一格**：CI 命中 `cgen_util_com.cpp:800` 的前期绑定支（`isEarlyBoundCom_ && earlyBoundSym_` 且 `comMethods` 命中），本机同一格没命中、落到 364/369 行按 unpackType 定档。这把上面「下一步两步」里的第②步具体化了 —— 要查的是**发码期那张「这枚变量是哪档 COM 类型」的表**：`knownTypedComVars_` 六个登记点（`cgen_decl_{func,proc,prop,var}.cpp`、`cgen_localdecl.cpp`、`cgen_base_generate_crossmod.inc:31`）每处都是「按类型名查符号、再问它是不是 ComClass」，而 `StdFont` 在库里同时有 `ComInterface IFont` 与 `ComClass StdFont` 两档 ⇒ 首命中若取决于无序容器的桶序或跨模块注入次序，两台工具链换哈希就换答案，**而门抓不到**（Build 与 Tests 同一枚工件，自洽）。

**2026-10-08 三台对跑把变量收窄到只剩一个 = C3.exe 自己的工具链**（工件 = 门 #408 的 `c3-exe`，run 37709604233、head `f2d9c400`；本机同一笔提交冷编**两台** = `.build/wt407`、`.build/wt407b`）：
· **复现**：把 CI 那枚与本机那枚放在**同一台机器、同一个 cwd、同一份源**上跑 `--emit-c`，语料 394 份 × 两架构 = 788 份捕获里 **14 份不同**，全部落在 Charts 2020 的六枚 UC 工程上（`Proyecto1` 102 行/片、`ucChartArea`/`ucChartBar`/`ucPieChart` 13、`ucProgressCircular` 18、`ucTreeMaps` 22），其余 620 份可比捕获逐字相同；
· **差异族只有一形**：早期绑定 COM 的**出口选择** —— CI 发带类型的出口（`vb6_ComGetIntProp(oFont, L"Size")` / `vb6_ComCallInt((*New_Font), L"Charset", …)`），本机发泛型出口 （`vb6_VariantToLong(vb6_VariantFromComResult(vb6_ComGetProp(oFont, L"Size")))` / `vb6_ComCall(…)+vb6_ComVarFree(…)`）；
· **排除环境**：两枚 exe 在同一份工程上打出的是**同一条** `VB4001: TypeLib reference path not found: …\STDOLE2.TLB` ⇒ 类型库解析结果两边一致，注册表/机器不是变量；
· **排除运行期随机**：同一枚 exe 连跑 5 次（本机）/3 次（CI 那枚）`--emit-c` 的 md5 **一字不差** ⇒ 每个二进制内部是确定的；
· **排除构建实例**：本机同头两台冷编（build1 6,228,992 B / build2 6,228,992 B，md5 不同因链接时间戳，但 emit 在 12 份对照捕获上 **diff=0**）⇒ 换个构建目录不改变答案；
· **剩下的唯一变量 = 编译器版本**：读 PE 可选头 `MajorRuntimeVersion/MinorRuntimeVersion` ⇒ 本机 = **14.29**（VS2019 16.x，`c3_build.ps1` 里写死那套），CI = **14.51**（VS2022 一线）；
· **机制候选**：`Scope::symbols_` 是 `std::unordered_map<std::string, std::unique_ptr<Symbol>>`，而全仓有 **35 处** `for (… : ….symbols())` 的range-for 直接吃它的迭代顺序（`symbol_table.cpp:332/352/436/465` 这四条是**查找本身**，其中 332 那条 `lookupModuleOverloadByLoc` 的兜底就是「扫到哪个算哪个」；`driver_codegen_dll_typelib.inc` / `driver_codegen_typedfield_scan.inc` 那几条是**注入与收集**，同键后写覆盖 ⇒ 赢家随迭代顺序）⇒ 哈希迭代顺序是 STL **实现细节**，跨大版本会变，这正是「同一份源码、两个二进制、各自稳定、答案不同」的形状；
· **spike（把 `Scope::symbols_` 与 `symbols()` 换成 `std::map` 后重编那台再对跑）**：**答的是否，且比原先更值钱**：把 `Scope::symbols_` 与 `symbols()` 换成 `std::map` 重编那台（`.build/wt407b`，spike 已撤）与本机无序版对跑那 12 份对照捕获 = **diff 0**，而 CI-vs-两台仍是同样的 102/13/18 ⇒ 符号表的哈希迭代顺序**不进这一路的发码**，这一条嫌疑排除。更新的读数：CI 那一侧的答案本身是错的 —— 它连 `L"Name"`（BSTR 属性）都发成 `vb6_ComGetIntProp`（int 出口喂 `vb6_StrPtr` 的 BSTR 槽），本机发的是 `vb6_VariantToString(vb6_VariantFromComResult(vb6_ComGetProp(oFont, L"Name")))`。⇒ 这一格从「发码不一致」升级成「**上线的那一份在这一形上是错的**」，而两台的差唯一变量仍是 C3.exe 自己的工具链（本机 14.29 / CI 14.51）。下一步要么把两侧工具链钉成一套（`ci.yml` 里钉 VS 版本 —— 与 ai/030 那条 `toolsetTag` 空转的欠账同一刀），要么让 **CI 自己吐出 `--emit-c` 语料**当发码神谕（本机的 emit 只能证形状存在，不能证上线的那份长什么样）。
· **今天就能用的口径**（把上一段那条收紧）：**发码形状的证据必须来自与产物同一工具链的那台**；本机的 `--emit-c` 只能证「形状存在」，不能证「CI 上也一样」。跨工具链的对跑已经可复用：`python .build/s73_ab.py`（全语料，两台）与 `python .build/s73_three.py`（本机两台 + CI 一台，12 份捕获）。另一条口径不变：**门只测默认架构 ⇒ 布局类改动要 x86 真编译真跑**。

**2026-10-08 口径拍板 + 神谕换边**（用户：不钉 VS 版本，只要求 VS2019 及以上）：既然两侧工具链都合法，这一格就只剩两条腿 —— ① 把发码里任何**依赖构建期顺序的输入**消掉（产品活）；② 让**证据来自产物那一台**（门禁活）。② 已落地：`ci.yml` 新增与 Tests 并行的 `Emit manifest (shape oracle)` 作业（needs build ⇒ 墙钟不变），`scripts/emit_manifest.ps1` 对全语料 x64 跑 `--emit-c`，**归一化后按 sha256 逐行留档**：只取 stdout（stderr 带机器特定警告）+ 把 `// Generated by vb6c3 (C3.exe) from <绝对路径>` 换成 `<PATH>` + 把本次 checkout 的根换成 `<ROOT>`（不归一化两边每一行都不同）；清单头三行写 `c3-exe sha256` / `toolset vs=… msvc=…` / `inputs=…`，工件 `emit-manifest` 存 14 天。清单可比、不搬全文（几十 KB）⇒「CI 上那份发码长什么样」从此是一行 diff 就能问出来的事。先只留档不判红（两侧现在已知差 14 行），等 ① 修到逐行相同再把「与 checked-in 清单比对」升成硬判据（现在就升 = 天天红）。
· **① 的候选已按性价比重排**：全仓只有 4 处 `std::sort`，其中 3 处的比较器**允许平局** —— `cgen_form_uc_methods.inc:188` 与 `cgen_form_uc_props.inc` 按小写名排（同名多档时谁在前随实现走），`cgen_iface_vtbl.cpp` 的 views 排序同理。`std::sort` 不稳定、其结果随实现/版本变，正是「同一二进制内部自洽、两台之间不同」的形状 ⇒ **下一轮先把比较器补成全序（平局再按稳定 id/名），别再换容器**。已排除：`Scope::symbols_` 换 `std::map` 的 spike（与无序版 0 差）；全仓无「键是指针」的哈希容器（键全是字符串 ⇒ 顺序只随 STL 版本，不随 ASLR）；`vb6BuiltinObjTypes` / `projectClassNames` / `knownTypedComVars_` 都只做 `count/find`，不存在首命中。
· 这格顺手量出一条**与工具链无关的真缺陷**：`cgen_assign_com_prop.inc:96` 的 `resolveComValue(wcProp == "list" ? "BSTR" : "Long")` 是**按成员名硬编码一档** —— 走进这一支时 `Name` / `Caption` 这类 BSTR 属性也被按 Long 解包（`vb6_ComGetIntProp` 喂 `vb6_StrPtr` 的 BSTR 槽）。该单独修：这一问交给「这枚 COM 属性是什么型」的唯一出口（#260 那条解封出口同族），不该按名字猜。

**2026-10-08 门 #410 之后的两台读数**（清单机器跑通了，但读数推翻了两条口径）：
· **门 #410 = run 37719504496、head `7e192d2e`、attempt 1、12 job 全 completed/success**（新增的 `Emit manifest (shape oracle)` 与 Tests 并行，墙钟没变）。两侧清单都 `inputs=395`、行数相同，但 **changed=319（不是 14）**，而且 CI 那份**一致更大**（单文件 .bas 恰好 +7 字节，多文件工程更多）。⇒ **「拿 CI 清单对本机清单」里混着机器变量**，#245 的判据仍然是**同一台机器上两台 exe 对跑**（`python .build/s73_ab.py`，那 14 份）。本机三连跑逐字节相同（`tests/acc/acc_main.bas` = `c6fa267a9252…` bytes=1778）⇒ 工具与运行都不抖；另外把本机 checkout 用 `subst X:` 换个根再跑同一份输入，归一化后哈希不变 ⇒ 归一化本身不漏根。
· **`# toolset` 那行我读错了，已改**：CI 报 `msvc=14.29.30133` 看着与本机同 —— 但脚本取的是 `Get-ChildItem VC\Tools\MSVC | Select -First 1`，那是**按字母序最小的一档**，runner 上多版本共存时未必是 Build 作业真正选中的那档 ⇒ 这一版读数**不能**用来判「两侧工具链相同」。现在读**二进制自己的 PE 头**（`pe-lnk=<MajorLinkerVersion>.<Minor>`，optional header 偏移 +2/+3）并把已装的 MSVC 目录全列出来。
· **撞见一条环境侧线索**：/RTC 那台跑 `ucChartArea` 时诊断里有 `TypeLib reference path not found: <工程目录>\..\..\..\..\WINDOWS\SysWow64\STDOLE2.TLB` —— 引用的 TLB 路径按工程目录的四级 `..` 拼出来，落在这台机器上不存在的位置。**这正是 #245 那一族的输入侧**（`StdFont` 的成员签名从哪来）。⇒ 下一步先问「这台机器上 stdole2.tlb 经哪条路能找到 / 找不到时 `comMethods` 是空还是半空」，别再猜排序。
· **/RTCu 那条嫌疑既没证实也没排除**：Debug(/RTC1) 那台交 rc=0、无 Run-Time Check Failure —— 但 /RTCu 只盯**函数内的局部标量**，而这里的读法是 `ComMemberInfo`（vector 里的堆对象）的成员被四处复制 ⇒ 这个探针对它天生是哑的，**别把「没报」读成「没有」**。静态已量到的形状：`ComMemberInfo::returnType/kind`、`ComParamInfo::type/direction` 无默认初值，而 `parseVarDesc` 只在 `varkind == VAR_PERINSTANCE` 那支写 `returnType` ⇒ 非 PERINSTANCE 那档交给 `driver_semantics.cpp:141/175/206/345` 复制，读的是没赋过值的字段（照 #172/#238 同族修：补默认初值 + 每支都赋值 + 哨兵钉住）。
· **工具链这条嫌疑已经用二进制自己的 PE 头证实**（本机直接读 #410 的 `c3-exe` 工件 vs 本机那枚）：
  `CI pe-lnk=14.51` / `LOCAL pe-lnk=14.29`（optional header = e_lfanew+24，链接器版本在其 +2/+3 ⇒ 绝对 +26/+27；
  我第一版脚本写的是 +6/+7，读出来是 `6.0` —— 已修，两份二进制现在 PS 与 Python 两种读法同值）。
  ⇒ **#245 的「唯一的变量是那枚 C3.exe 怎么编出来的」这条口径站住了**，而且 runner 上多档 MSVC 共存时
  `vswhere -latest` 选中的是 14.51 —— 这正是旧那行 `msvc=14.29.30133` 把我骗了一轮的原因。
· **同一台机器上的文本差第一次拿到了（13 行、一种形状），「未初始化读」这条嫌疑回到首位**。
  把 #410 的 `c3-exe` 工件下载到本机，与本机那枚跑**同一份** `tests/Charts 2020/ucChartArea/Proyecto1.vbp`
  ⇒ `5267 行 / 差 13 行`，逐条都是同一问「COM 成员读走哪条出口」：
  本机 = `vb6_ComCall((*New_Font), L"Name", NULL, 0)` / `vb6_VariantFromComResult(vb6_ComGetProp(oFont, L"Name"))`（**泛型 VARIANT 那一路**），
  CI = `vb6_ComGetIntProp((*New_Font), L"Name")` / `vb6_ComCallInt(...)`（**带类型那一路**）。
  **关键点不是"两台各挑一条路"，而是 CI 那侧 `Name`(BSTR) / `Size` / `Bold` / `Charset` 八条成员全被判成同一个 C 类型**
  （`mapType(sig.returnType)` 一律给 `int32_t` ⇒ 全走 `ComGetIntProp`）—— 成员各不相同的类型读出来却**整齐划一**，
  只有两种解释：某处读的是**没赋过值的字段**（同一段代码、同一个栈槽 ⇒ 同一个"垃圾"，因此该二进制内自洽、跨二进制不同），
  或某处**所有分支都掉进同一个兜底**。本机那枚的答案是"泛型"（等于 `mapType` 没落进任何具名档 ⇒ 也是兜底，只是兜到另一档）。
  ⇒ 下一条 measurable：临时给 `TypeLibParser::parseFuncDesc` 打一行 stderr（成员名 / `invkind` / `elemdescFunc.tdesc.vt` /
  算出来的 `returnType`），用**本机 14.29** 那枚编 `_IFont`/`IFontDisp`，看它到底读到什么 —— 若 `vt` 是 `VT_BSTR` 而
  `returnType` 不是 `String`，缺陷就在 `mapTypeDesc`；若 `returnType` 是个枚举外的数，就是未初始化读。**别再换容器、也别先动 `std::sort`**。
· **临时探针（本机 14.29 那枚，`C3_TLBTRACE=1` 把 `parseFuncDesc/parseVarDesc` 读到的签名打到 stderr，跑 ucChartArea，用完已撤 —— 文件回到 md5 `b10ced97…`、exe 重编）**：
  本机读到的是**每成员各异、且大体正确**的档：`Name` get ⇒ `returnType=8`(String)、`Bold/Italic/Underline/Strikethrough` ⇒ 11(Boolean)、
  `Charset/Weight` ⇒ 2(Integer)、`Size` 三档 ⇒ 3/12/6，四条 **put(invkind=4) ⇒ 0(Empty)**。
  ⇒ **分家不在类型库读取那一层**（这台机器上读对了），在发码期那一步「问不问前期绑定那张表、答案从哪一份 sig 取」。
  顺着这条把两处**口径不齐**钉出来（都是能独立成立的真缺陷，与工具链无关）：
  ① `insertComMethod`（`src/driver/driver_semantics.cpp:23`）**只挡"setter 覆盖 getter"这一向**，同名两份 **getter** 之间是后写覆盖 ⇒
     赢者随上游容器的迭代序，而迭代序随 STL 版本变 —— 这正是「同一二进制内自洽、两台之间不同」的形状；
  ② `resolveComValue`（`src/backend/cgen_util_com.cpp:327`）判据只看 `comMethods.find` 命中，**不看 `isPropertyGet`**，
     而同族的另一处（`src/backend/detail/expr/cgen_expr_call_arg_emit.inc:90`）**是看的** ⇒ 同一张表两种问法。
     这条能单独解释 CI 那侧的错：若赢进来的是 **put 那份**（`returnType=Empty`），`mapType(Empty)="int16_t"`（`cgen_base_type.cpp:28`）
     ⇒ 落进 `int16_t` 档 ⇒ **`Name`(BSTR) 被发成 `vb6_ComGetIntProp`**，与 #410 工件在本机跑出的 13 行差异逐条对上。
  ⇒ 下一刀的形状：COM 签名表的**填充按名字定序** + **两处消费点问同一条出口**（getter 优先、非 getter 一律退回泛型），
     哨兵钉「读 `comMethods` 的每处先问 `isPropertyGet`」；判据 = 同一台机器两台 exe 的文本差 13→0，加上 CI 清单与本机清单里这 14 份收敛。
· **跨机器那 319 行的谜底已对上，而且要订正我上一条的说法 —— 不是"机器变量"，是编译产品那台工具链的源码字符集**
  （拿 #411 的 `emit-samples` 与本机那份原文逐行 diff 才看得见；清单只有哈希，看不见这一层）：
  `tests/acc/acc_main.bas` 两份 raw 只差 **4 行** —— 两行是 `// Generated by … from <路径>`（清单里已归一化掉），
  另两行是**发码里的中文分节注释**：CI 那枚交 **UTF-8** 字节（`过程实现` 12 字节、`入口点` 9），本机 14.29 那枚交 **GBK**（8、6）
  ⇒ 净差 (12−8)+(9−6) = **+7**，与清单 `bytes` 那一列的差**逐字节对上**。
  ⇒ 机制：C++ 源里的 `"…"` 字面量按**编译器对源文件字符集的判定**存进 .exe —— MSVC 14.29 按 ANSI(GBK)、14.51 按 UTF-8，
  于是**发码里的中文注释是不确定的**（这本身是一格该修的缺陷：产品发出的注释应当是 ASCII，或按固定编码显式落字节；
  与既有口径「夹具注释一律 ASCII」是同一族，只是这次在**产品自己**身上）。
  ⇒ 读清单时的操作口径：一行差里可能混着两类 —— 编码类（覆盖面大、语义无关）与 #245 的 COM 出口类（实测 14 份）；
  先按"差异行是否只含非 ASCII 字节"把前者剔掉，再看剩下的。
· **清单加第二列 `ascii256`（只哈希纯 ASCII 行）+ 本刀的语料爆炸半径**：全语料 395 份输入里，账 #245 这一刀只改动 **7 份**
  （`Charts 2020` 的 6 枚 UC 工程 + `tests/VBFlexGridDemo/VBFlexGridDemo.vbp`），其余 **388 份逐字节不变**、`rc` 一个都没变
  ⇒ 零附带影响；而 §B73 之前那串"14 份捕获不同"就是同一批的另一面：7 份 × 两架构 = 14。
  从此清单每行两列并排（`sha256` = 归一化全文、`ascii256` = 只取纯 ASCII 行）：**两侧 `ascii256` 相同而 `sha256` 不同 ⇒ 只差在编码族（#267）；
  两列都不同才是语义差** ⇒ 跨工具链比对不会再被 319 行编码差淹没。CI 那份已经用 git 回读（`ci/emit-manifest` 分支，门 #413 起），
  读数从此不需要 `api.github.com`。
· **新工件 `emit-samples`**：清单只有哈希，跨机器对不出「差在哪一行」⇒ `-Samples` 点名的三份输入（`tests/acc/acc_main.bas`、`tests/asm/AsmTest.bas`、`tests/Charts 2020/ucChartArea/Proyecto1.vbp`）把**未归一化原文**一起留档 ⇒ 下一轮把 CI 那份与本机 `.build/emit-samples/` 直接 diff，那 7 个字节是什么一眼可见。
**2026-10-08 账 #245 已修 = 根因是一处未初始化字段，跟构建期顺序无关**：`parseVarDesc` 只在
`varkind == VAR_PERINSTANCE` 那一支给 `member.returnType` 赋值，而 **dual 接口的属性是 VAR_PROPERTY(3) 的
VARDESC**（stdole 的 `StdFont` 八条属性全在这一档）⇒ 交出去的是没赋过值的 16 位枚举字段。临时探针实测：
本机那一枚**连跑三次**读到 `29620 / 6971 / 50156`，三个垃圾值都落进 `mapType` 的 `default` ⇒ 本机发码"看着稳";
CI 那一枚（`pe-lnk=14.51`）稳定落进 `int16_t` 档 ⇒ 发成 `vb6_ComGetIntProp`。**订正上一条那句
"决定随构建变、构建内自洽"—— 真相更糟：它随进程变，只是本机这一档一直兜在同一支，A/B 才显得自洽。**
修法三件：① `parseVarDesc` 无条件 `member.returnType = mapTypeDesc(&pVD->elemdescVar.tdesc, pTI)`
（三种 varkind 的真类型都在 `elemdescVar.tdesc` 上）；② `ComMemberInfo`/`ComParamInfo` 四个标量字段补默认初值
（`Vb6Type::Variant` = "不知道"，`direction = In`，`kind = Method`）；③ `resolveComValue` 的早期绑定支与同族
另一处对齐口径 —— **先看 `it->second.isPropertyGet`**（put/method 那份签名不许当属性读来解包），
**类型未知一律退 `vb6_VariantFromComResult(vb6_ComGetProp(...))`**，不再按 `unpackType` 猜一档
（猜档就是 `L"Name"` 被发成 IntProp 的那条通道；同文件 800 行那条本来就查 `isPropertyGet` ⇒ 两处从此同口径）。
判据：`ucChartArea/Proyecto1.vbp` 同机三台（BASE 未修 / NEW 已修 / CI 那枚 14.51）按**行多重集**对 ⇒
`CI-vs-BASE = 12` 行、**`CI-vs-NEW = 5` 行，且这 5 行全是"CI 那枚还没修"的形状**
（`ComGetIntProp(oFont, L"Name")` ⇄ 修好的 `ComGetStringProp(...)`）；NEW 交出的档位与类型库真值一致
（`Name`→String、`Size`→double、`Charset/Weight/Bold/Italic/Underline/Strikethrough`→int）；
真编译 `Proyecto1.exe` rc=0、零 error、零 C4244；本地 `-Category compile` 组 40/40 全绿。
新哨兵 `scripts/check_com_prop_type_authority.ps1`（第 34 道 [STATIC]）C1/C2/C3a/C3b/C3c/C4 六条各用一处
假改动证红、跑完逐份还原核 md5。两条 reusable：
- **C2 的第一版是假绿**：判据只扫"以 `;` 结尾"的行，而这些声明行都带行尾中文注释 ⇒ 一条都看不见；
  先脱 `//…` 再判才生效。**"注释让判据变哑"这一族已经第三次踩到**（脱注释、别按 EndsWith 过滤）。
- **C4 census 顺手量到 `comMethods.find(` 的后端消费者是四处**（`cgen_util_com.cpp` /
  `cgen_expr_call_arg_emit.inc` / `cgen_expr_call_com_bind.inc` / `cgen_expr_call_prelude.inc`）
  ⇒ "把这个问题合成一处出口"这一格还欠着；本轮把名单钉死、把其中两处口径并齐，第五种问法要进来必须先过这条。

**2026-10-08 回读通道自己坏了两轮（门 #414/#415 全绿而清单从没上分支）**：本机到 `api.github.com` 那条被本地反代整段拒（403 + `Via: 2.0 Caddy`，带不带 token 都一样），而 `github.com` 的 git 通路是通的 ⇒ 清单改成 CI 用 git 推回 `ci/emit-manifest` 分支、本机 `git fetch` 读。**但这条通道上线后一直没送东西回来**：分支 tip 的提交信息始终停在 `head=1c84b437`（#413 那一趟），连 #414/#415 都没动过。两个**互相独立**的成因，都只在这一趟才凑齐：
- ① **浅仓库拒绝 push**。`actions/checkout@v4` 默认 `fetch-depth: 1`，工作区带 `.git/shallow`；基于已有 tip 叠提交 ⇒ 祖先链上有浅标记，git 直接拒推。#413 侥幸成功是因为那轮分支还不存在、走的是 `--orphan`（历史被切断，压根没有缺失的父提交）⇒ **「上一趟成了」把这一条掩盖成「不是浅的问题」**。修法：`emit-manifest` 作业的 checkout 补 `fetch-depth: 0`（只这一档，其余作业不动：全量 fetch 的成本只落在一个 job 上）。
- ② **`git checkout` 把手上那份清单换掉了**。`emit-manifest.txt` 在仓库根、被 .gitignore 忽略，而 checkout 对「被忽略的未跟踪文件」是直接覆盖的（不像普通未跟踪文件会拦）⇒ `checkout -B ci/emit-manifest FETCH_HEAD` 之后根上是**分支里那份旧的**，`add`/`commit` 看不见任何变化，`push` 回 `Everything up-to-date` 也是退出 0。修法：切分支前先 `Copy-Item` 到 `$env:RUNNER_TEMP`，切完再挪回来。
- **这一族真正的教训是判据形状**：按设计「不判红」的那一步（`exit 0` + 只打一行）**默认是沉默的，沉默就会被读成成功**。#413 之后每一步都打 `publish-ok`，而它其实是 `Everything up-to-date` 那条 0 ⇒ 两轮白等。现在这一步收尾自己把远端 tip 读回来打一行 `publish-ok: base=<基> remote-tip=<sha>`，本机轮询按 `head=<本次 sha8>` 归属，**推没推上去不再需要读日志**；顺带去掉 `git commit ... | Out-Null`（吞掉的正是「nothing to commit」这条唯一线索）。
- 顺手把 `if (-not (Test-Path $hold))` 补上：清单没生成时（前一步 throw）这一步打 `publish-skipped` 直接退，别拿分支里那份旧的去叠一个假提交。

**2026-10-08 编码族那一格改口 —— 既不是工具链、也不是随机机器变量：`--emit-c` 的 stdout 走了「按控制台代码页转字节」那一层（账 #267，本刀已修）**。上一条那句「MSVC 14.29 按 ANSI(GBK) 存字面量、14.51 按 UTF-8」是错的，三条读数各自把它否掉：① `CMakeLists.txt:20-23` 早就给 MSVC 挂了 `/utf-8`；② 本机冷编的那枚与 CI 工件那枚 `C3.exe` 里，`过程实现`/`入口点`/`前向声明` 的 **UTF-8 字节都有命中、GBK 命中 0** ⇒ 两枚二进制的字面量同为 UTF-8；③ 决定性的是**同一次运行里同时存在两个答案**：中间目录 `%TEMP%\C3C\<会话>\AccMain.c` 是 UTF-8，`--emit-c` 重定向到文件的 stdout 是 GBK（首处不同在第 1141 字节，stdout=185=GBK 的「过」首字节、文件=232=UTF-8 的「过」首字节）。
· 现场在 `src/driver/main.cpp:109-220` 的 `ConsoleUtf8Buf`：句柄是控制台就 `WriteConsoleW`（渲染与 chcp 无关），句柄是**管道/文件**就 `WideCharToMultiByte(cp_, …)`，`cp_ = GetConsoleOutputCP()`（取不到退 `GetACP()`）。对**诊断文本**这个口径是对的（脚本自己那条注释就是实测来的：PS 5.1 与 pwsh 7 都按控制台代码页解，写 UTF-8 给它们反而乱码）；错在**产物**也跟着走 —— 发码的权威是写进中间目录、cl 拿 `/utf-8` 编的那两份（`msvc_driver.cpp:259`），stdout 只是它的另一个载体。这一格的形状正是本项目一直在抓的那种：同一个问题两份答案，其中一份还随机器变。
· 修法 = 产物那一路收成唯一出口：`encoding.hpp` 的 `writeStdoutRaw`（直接 `fwrite` 内部 UTF-8；stdout 是文本流，`\n` 照旧翻成 CRLF，与写那两份文件的 `ofstream` 一致），发码分支先 `std::cout.flush()` 再交字节（顺序不乱），诊断留在老一层不动。
· 判据两层。**运行期** `Test-EmitcByteCaliber`：把控制台代码页钉成两档各跑一趟，要求 stdout 的前缀**逐字节**等于 `.h + CRLF + .c + CRLF`，且非 ASCII 字节在场、整体能按 UTF-8 **严格**解码（少了后两条，纯 ASCII 夹具上会空转）。改前那一枚在本机当场红：`cp=936: 第 1141 字节与中间目录那两份不同 (stdout=185 文件=232)` + `cp=936: 不是合法 UTF-8 (被控制台代码页折过)`；改后绿。**结构性** `scripts/check_emit_c_artifact_caliber.ps1`（第 35 道 [STATIC]）：emitC 分支必须走 `writeStdoutRaw` 且先 `cout.flush`、出口调用点恰好 1 处、`options.emitC` 的消费者恰好 2 处且名单钉死（发码那一路 + `runLinker` 早退）、全仓不许再出现 `std::cout << cgen.`。六条负控逐条能红，且都跑在**源码树副本**里（`/b470_tree`），不碰共享树。
· 爆炸半径：本机改前/改后两份全语料清单逐行比 —— `sha256` 差 319 行、**`ascii256` 差 0 行** ⇒ 这一刀只换编码载体，一处语义没动。
· **门 #416 的读数**（回推通道修好后 CI 交回的第一份清单，head `fd5e8d83`、`inputs=395`、`pe-lnk=14.51`）：与本机改后那份对跑 —— **`sha256` 相同 394 / 395**，跨机器只差一份输入：`tests/VBFlexGridDemo/VBFlexGridDemo.vbp`，而且它**不是编码差**（本机改前改后两趟的 `ascii256` 相同，CI 那份与两者都不同）。⇒ 账 #245 那一格到此才真收口：`VAR_PROPERTY` 未初始化 `returnType` 修掉之后，两台只剩这一处**机器状态**待归因 —— 该工程 `Reference=` 三行指向 `C:\Windows\SysWOW64\*.tlb`，产物里还嵌着 28 处仓库内绝对路径 ⇒ 疑点从「哪台工具链编的」挪到「这台机器注册了哪些类型库」。证据已备好：`emit_manifest.ps1` 的 `-Samples` 默认名单加了这一份，下一趟门会把它**未归一化的原文**交回来，逐行一比就清楚差的是成员面还是环境串。
· 两列的用法随之改口径：**`sha256` 是跨机器可比的主列**（这一刀之后才成立），`ascii256` 退成「差在编码还是差在内容」的分诊列 —— 它会**漏**：一条带中文注释的语句里 `vb6_ComGetIntProp` 换成 `vb6_ComGetStringProp`，整行被它扔掉就看不见（#414 那 13 行差、以及 #416 这一处，都是靠 `sha256`／原文才露出来的）。

**2026-10-08 门 #417 的读数 + 最后那一行归因到手（账 #245/§B73 与 #267 一起收口）**
- **门 #417（head `0f678021`，发码出口那一刀之后）CI 清单 vs 本机清单**：395 行里 **`sha256` 只差 1 行** —— 还是 `tests/VBFlexGridDemo/VBFlexGridDemo.vbp`（CI `16d614bd…` / 本机 `793dbb81…`），而**其余 394 行连全文哈希都对上了** ⇒ #267 那一刀把跨机器只剩这一处，且它不再是编码问题。另两个对照组：**CI #416 vs CI #417 = 0 行差**（同一台机器、换了发码出口 ⇒ 字节不变 —— 正是这一刀该有的形状：它只把「随机器变」那一份答案改掉，不动 UTF-8 环境里的产物）；本机改前 vs 改后 = 319 行 `sha256` 差、**`ascii256` 零差**。
- **那一行的成因 = 语料里有一处输入随机器注册的类型库变**，不是编译器。三条读数钉死：① 这台机器 `C:\Windows\SysWOW64\OLEGuids.tlb` **文件不存在**，但 `HKCR\Typelib\{5A2B9220-…}\1.0\win32` 注册着，指向**另一份该工程的副本** `D:\vb_yqt4qPac\VBFLXGRD-master\…\OLEGuids.tlb` ⇒ 本机把那份类型库导入了，CI 上既没文件也没注册 ⇒ 没导入。② 在**本机**把 .vbp 里那一行 `Reference=` 删掉再发码，得到 `sha=16d614bd… ascii=62dcde7f…` —— **与 CI 那一行逐字节相同**（复现，不是猜）。③ 把同一行改指**该工程自己带进仓库的那份** `OLEGuids\OLEGuids.tlb`（`git ls-files` 里有，19860 字节，没被 ignore），得到 `sha=793dbb81…` = **与本机原来那行相同** ⇒ 换路径不改本机行为，只把 CI 拉回来。
- 差出来的形状（一句话就读得懂）：类型库在的时候那些字段是**带接口的指针** `vb6_ComIface_IOleObject* PropOleObject = NULL;`（`IOleControl` / `IPerPropertyBrowsing` / `IOleInPlaceObjectWindowless` / `IDispatchUnrestricted` 一族，连 `typedef struct vb6_ComIface_*` 一起），不在的时候**退化成** `void* PropOleObject = 0;`。按**归一化后的纯 ASCII 行多重集**数：只在「原样」里 110 种 / 138 行，只在「缺库」那版 92 种 / 118 行（早先记的 124/106 把每模块一行的 `// Generated by … <绝对路径>` 也算进去了 —— 探针各在自己的根下跑，路径天然不同；归一化之后剩下的全是这一族字段与 typedef）。⇒ 这不是"两种写法等价"：CI 一直在编一份**弱一档**的类型库成员面。
- **修法 = 改语料不改产品**：`.vbp` 里那一行从 `\\?\C:\Windows\SysWOW64\OLEGuids.tlb` 改成相对路径 `OLEGuids\OLEGuids.tlb`（工程目录下那份）。这一族的口径写下来：**`Reference=` 不许指向"这台机器恰好有什么"** —— 要类型库就把那枚类型库带进工程目录，用相对路径引它。（另一头不用动：路径找不到时 C3 是**响的** —— `(1,1): warning VB4001: TypeLib reference path not found: …` 走 stderr，实测在；所以这类分家一开始就能在日志里看见，只是过去没人去翻。）顺带记下 `msdatsrc.tlb` 那一行 `..\..\..\..\Windows\SysWOW64\…` 也是同一族的隐患（按 checkout 深度解析），但它的成员在产物里 0 命中 ⇒ 眼下不参与这次分家，别顺手改出一片新红。
- 待门 #418 证实：CI 那一行应当变成 `793dbb81…`（= 本机）⇒ **395/395 逐行相同**。到那一步 §B73 这一格才算真的合上，而「与 checked-in 清单比对」也可以从留档升成硬判据（清单头一行 `c3-exe sha256` 继续留着，它现在只是身份注脚，不再是差因）。
**2026-10-08 门 #418 证实 395/395 —— 「与 checked-in 清单比对」从留档升成硬判据（§B73 这一格的落地形状）**
- 门 #418（head `425083b3`，= 语料那一行 `Reference=` 改成相对路径之后）交回的清单里 `tests/VBFlexGridDemo/VBFlexGridDemo.vbp` 那行是 `793dbb81…` —— 与本机一致 ⇒ **395 行 `sha256` 零差**。本机独立复算一遍（`.build/b498_manifest.txt`，`pe-lnk=14.29` 对 CI 的 `14.51`）：395 份输入全共享、哈希差 0 行、缺席/多出各 0。注脚那两行（`c3-exe sha256`、`toolset pe-lnk=… runner=…`）两边天生不同 —— 这正是它们只作身份注脚、不参与比较的理由。
- 升成硬判据的三件：① checked-in 的 `emit-manifest.expected.txt`（408 行 = 13 行注脚 + 395 行数据；注脚里写着「什么时候才该重新登记」与「不要顺手更新」）；② `scripts/compare_emit_manifest.ps1` = **唯一一份比较逻辑**，CI 判红、本机复算、`-Bless` 重新登记三头都走它（绝不在 `emit_manifest.ps1` 里再抄一份 —— 那正是本项目一直在防的形状）；③ `emit_manifest.ps1` 出完清单就地调它，不一致 `exit 4`，由 ci.yml 已有的那句 `if ($LASTEXITCODE -ne 0) { throw }` 把作业刷红。
- 两处自证（这一步按设计会「判绿」，所以它得自己报自己）：**判定结果写进清单本身**（`# expectation-check=PASS` / `FAIL rc=N`；`#` 开头不参与比较）⇒ 结论随工件与 `ci/emit-manifest` 分支一起回来，本机不必能读 REST；**找不到比较脚本 = `exit 3` 并把 `NO-CMP-SCRIPT` 写进清单**（上一版这里写的是 `SKIPPED` + `exit 0` = 静默绿，改成红）。
- 覆盖率也自证：清单注脚自己写着 `inputs=<N>`（= 递给 C3.exe 的输入条数），而比较只吃 `#` 以外的行 ⇒ 万一有人把一整行数据折成注释，两边会一起少掉同一行、看着「全对」而覆盖悄悄缩水。`compare_emit_manifest.ps1` 现在要求 `inputs=` 与数据行数相等，不等就红。
- 负控四条都对着 CI #418 那份**真清单**跑（不拿假文件糊）：绿控 `395/395 相同 rc=0`；**改一个哈希字符** ⇒ `rc=1` 且点名 `tests/acc/acc_fam.vbp`、把期望/实得两行都列出；**少一行输入** ⇒ `rc=1` 且点名 `tests/acc/acc_ok.vbp`；**期望文件不在** ⇒ `rc=2 MISSING`。`-Bless` 的形状核对：注脚留着、数据行整批换成新的。
- 这一格踩到的工具坑（台账里「续行符是反引号」那条的又一件实物）：`Add-Content -LiteralPath $Out -Encoding UTF8 -Value` 把值写在**下一行**而没加反引号 ⇒ PS 5.1 报 `MissingArgument`，**可脚本退出码仍然是 0**，那句结论只 echo 在控制台、**从没进文件**（原图：`b498_manifest.txt` 里压根没有 `expectation-check` 那行）。只看 rc 会以为这一步通了 —— 判据要读工件本身。
- §B73 到此合上的是「跨机器同不同」这一问；仍欠的那格没变：**COM 签名表定序 + `comMethods` 四处消费者合一**（账 #245 的尾巴）。从这一轮起，改发码形状若没先归因就登记，门会直接指出是哪一份输入变了。
- **2026-10-08 第二次咬合，这回咬的不是本线**：门 **#423**（head `576b694d` = 用户的「窗体设计期 BackColor 含系统色现正确发射」那一刀）唯一红 = `Emit manifest (shape oracle)`，11 个 Tests 格全绿。清单照旧经 git 通道回来（`ci/emit-manifest` 尖端 `deace63b`，注明 `head=576b694d`），本机用**唯一那份比较逻辑**复算 ⇒ **395 行里 7 行差**，其余 388 行零动。两头读数把这一格钉死：① 语料里在 **form-level**（`Begin VB.Form` 那一层）写了 `BackColor` 的窗体恰好 **8 枚**，落在 7 个工程目录里（`Charts 2020/Form2`、ucChartArea/Form1、ucChartBar/Form1+Form2、ucPieChart/Form1、ucProgressCircular/Form2、czUI-main/frmDemo、dcsurf/DcForm）—— ucChartBar 两枚 ⇒ 它那行字节 **+180**，其余六份 **+89/+90** = 正好一行 `vb6_SetControlBackColor(…)`；② 另外 15 份带 `BackColor` 的 .frm 全在 depth>=2（控件层）⇒ 一行都没多发。**方向、条数、字节数三头一致 ⇒ 属「发码确实该变」那一档**，按注脚那条规则重登记，登记来源 = CI 那台 14.51 自己交的清单（不是本机复算）。
- **登记之后两侧都过，且登记是活的**：`compare_emit_manifest.ps1` 拿 CI 那份对上新表 = **rc=0 / 395 全同**；往新表里改一个哈希字符 ⇒ **rc=1 并点名那一行**（`tests/cls_neg/ci_n22_outsider.bas`）。**门 #424 已回 = run 37767816870、head `d54f1347`、attempt 1：12 条 check-run 全 completed/success、非绿 0，含 `Emit manifest (shape oracle)`** ⇒ 新表与 CI 那台的发码对上；这道闸第一次走完「咬住改动 → 归因 → 重登记 → 复验」整条路。
**2026-10-08 第三次咬合，咬的还是本线之外的刀（门 #425，head `9158e7f0` = 用户那三笔 3DMenu 修复）**
- 逐 job 读数：**唯一红仍是 `Emit manifest (shape oracle)`，11 个 Tests 格全 completed/success**。取数踩到一条形状：`commits/<sha>/check-runs` 只认**完整 sha** —— 给 8 位短 sha 它答 `total=0`，看着像「门没跑」，其实是没有那条记录。以后一律先 `git rev-parse`。
- 本机比较 = **395 行里 5 行差**（其余 390 行零动，缺席/多出各 0），五份字节全在**涨**：`tests/BalloonTooltips/prjBalloonTooltips.vbp` +219、`tests/VBFlexGridDemo/VBFlexGridDemo.vbp` +197、`tests/ctrlimagelist/CtrlImageList.vbp` +326、`tests/ctrlsstab/CtrlSSTab.vbp` +203、`tests/frmevents/FrmEvents.vbp` +223。
- **归因这次不靠推测，靠 A/B 冷编**：在 `.build/wt_b6` 冷编 `9158e7f0` 那枚 C3.exe，与旧头那枚对同一批输入逐行 diff（`.build/b618_ab.txt` / 明细 `.build/b618_diffs/`）。三笔刀各命中自己的那一族，且**方向和条数都反证得回去**：① `92258bab`（frx 图片数组从 .h 里的 `static` 定义改成 .h extern + owning .c 里一份定义）命中 4 份 —— 这 4 份正是语料里**唯一**发得出 frx 数组的工程（`vb6_frx_icon_frmBalloonTooltips` / `vb6_frx_pic_Picture1` / `vb6_frx_imglist_ImageList1_1..2` / `vb6_frx_tabpic_SSTab1_0`）；反向对照组 `tests/frxdata/FrxData.vbp`（.frx 里只有 List 记录，0 枚数组）**逐字节相同**，Charts 那几份 UC 工程也 0 枚数组（UC 侧根本不发这条）。② `3e9c1baf`（卸载重入的 `VB6_Unloading` 标记）命中 1 份 = `tests/frmevents/FrmEvents.vbp`，差异恰好是 4 加 1 删那几条 `SetPropW/RemovePropW(hwnd, L"VB6_Unloading", …)`；语料里写了 `Sub Form_Unload` 的输入只有它一份（另两份 .frm 在 `tests/test_form/`，没有任何 .vbp/.bas 引用它们 ⇒ 进不了清单）。③ `9158e7f0`（LoadPicture 缺文件抛 53）只动 RTL ⇒ 395 行零命中，与「RTL 活在 C3.exe 资源里、不进发码文本」对上。**本机每份的字节增量与 CI 那 5 行逐一相等**（上面那五个数就是两边的共同读数）。
- 光对上形状不够，还要问 extern 有没有落单的定义（那才会把链接弄断）：结构化读数（`.build/b619_struct.txt`）= 4 份工程 5 枚数组各 **1 条定义 + 1 条 .h extern**，`decl-only-no-def` / `def-not-declared` / `used-but-undefined` 三个集合**全空**；再把 `tests/ctrlsstab` 用新 exe 真编真跑一遍（x64、独立输出目录 `.build/b620_out`，不碰共享 `output/`）：build.log 里 `error C` 0 条、`unresolved external` 0 条，exe 出得来，跑到 `CTRLSSTAB-VISDONE` / `CTRLSSTAB-CLICKDONE` 且 rc=0。⇒ 判「发码确实该变」，不动产品。
- **重登记**（来源 = CI 那台 14.51 自己交的清单 `0ca2eedc`，`head=9158e7f0`，不是本机复算）：`compare_emit_manifest.ps1 -Bless` ⇒ numstat **5/5**、13 行注脚零动、BOM+CRLF 保持。两头自证：正控 rc=0 / **395 全同**；负控往新表改一个哈希字符 ⇒ rc=1 并点名 `tests/BalloonTooltips/prjBalloonTooltips.vbp`（期望/实得两行列出）。**门 #426 已回 = head `5f8834f2`、attempt 1：12 条 check-run 全 completed/success、非绿 0，含 `Emit manifest (shape oracle)`** ⇒ 新表由 CI 自己复验通过（这是这道闸第二次走完「咬住别人的改动 → A/B 归因 → 重登记 → 复验」）。

**2026-10-09 结案：那 14 份差异不再复现，而且真凶找到了 —— 不是「工具链会换迭代序」，是一枚没写初值的字段**

- **读数（这次是本机同一台机器、同一个 cwd 上的真双臂）**：把门 #437（head `15e119f3`）的 `c3-exe` **工件**下载下来
  （`gh run download 37881230099 -D .build/b872_ciexe -n c3-exe`，6,120,448 B / sha `6ff5674d2899` / PE 工具链 = **14.51**），
  对本地那枚（6,278,656 B / sha `cd32f425947b` / 工具链 = **14.29**）；两枚 exe 的 `src/` 只差 `c3rtl.rc` 的一行注释
  （`git diff 15e119f3..HEAD --stat -- src` = 1 file, 1 insertion）。跑法 = `scripts/emit_manifest.ps1 -ListInputs` 那份**唯一枚举口径**，
  逐条 `--emit-c --arch <x64|x86>`，字节走 `Start-Process -RedirectStandardOutput`（OS 级重定向，不经控制台代码页），逐文件比 SHA-256：
  **x86 `same=398 changed=0 rcDiff=0` / x64 `same=398 changed=0 rcDiff=0`**（探针 `.build/b873_ab_pilot.ps1`，读数 `.build/b875_full_x86.log`、`.build/b875_full_x64.log`）。
  另有一层独立对照：本机全语料清单与 CI 自己那台交的清单（`ci/emit-manifest`，`head=15e119f3`）**398/398 逐行相同**。
- **机制**：`ComMemberInfo::returnType`（`src/com/typelib_parser.hpp`）曾经没有默认初值，而 `parseVarDesc` 只在 `VAR_PERINSTANCE`
  那一支给它赋值 ⇒ dual 接口的属性（`VARDESC.kind = VAR_PROPERTY`，如 `stdole.StdFont.Name`）交出去的是**没写过的 16 位字段**。
  同一枚二进制连跑三次读到 29620/6971/50156（这是 #245 记的读数），而**换一台工具链编出来的二进制会稳定落进另一个档**
  ⇒ 早期绑定的 COM **出口选择**跟着换（泛型 `vb6_ComGetProp` ⇄ 带类型的 `vb6_ComGetIntProp`，而 `Name` 是 BSTR，后者压根是错的答案）。
  这条把当年那句「剩下的唯一变量 = C3.exe 自己的工具链」订正到位：**工具链只是把残值固定成了另一个值，变量从来是那个没写的字段**。
  当年最可疑的「35 处 `symbols()` 迭代吃 unordered 序」不是这一格的因（spike 换 `std::map` 已答否，那句现在也只对它自己那一格成立）。
- **护栏（第 44 道哨兵）**：`scripts/check_com_sig_field_defaults.ps1`，登记为 `[STATIC] com_sig_field_defaults`。
  F1 名单里 8 枚签名载体 struct 必须在指定文件里**按 `struct <名字> {` 找得到**；
  F2 每枚**值类型或裸指针**字段必须带默认初值（`std::string` / `std::vector` / `std::unique_ptr` 这些自己会初始化，跳过）；
  F3 检查到的字段数不许退化（地板 15，实测 **21**），且 `returnType` 必须在 `ComMemberInfo` 与 `ComMethodSig` **两头都出现**。
  四条负控各自红并点到自己的工作：A 摘掉 `ComMemberInfo.returnType` 的初值 ⇒ F2；B 把 `struct ComMethodSig` 改名 ⇒ F1；
  C 往 `ComParamInfo` 插一枚无初值的 `bool` ⇒ F2；D 把 `returnType` 改名（靶子消失）⇒ F3。植完按 md5 逐份还原，还原后复跑绿。
  ⚠ **F1 第一版只 `IndexOf('struct ' + $nm)`，B 那档假绿** —— `'struct ComMethodSig'` 是 `'struct ComMethodSigRenamed'` 的**前缀**，
  改名照样命中。改成 `[regex]('struct\s+' + 名 + '\s*\{')`（要求紧跟 `{`）才咬得住。一般式：**"这枚符号还在不在"这类判据，
  锚点必须带边界**，子串匹配对"改名"这种坏法是瞎的 —— 与本线「哨兵能红才算护栏」同调。
- **口径落定**：① 发码形状的证据必须来自**与产物同一工具链**的那台，这条从今天起有了一条便宜的执行方式 ——
  CI 的 `c3-exe` 工件可以直接 `gh run download`（不必 `api.github.com`），拿到本机与本地那枚并排跑；
  ② 登记发码清单时**吃 CI 那台的数**（`scripts/rebless_emit_manifest.ps1`，§B100），本机复算只当交叉验；
  ③ 跨工具链的发码分家从此有了哨兵 + 门禁两道网，剩下的同类风险只在「新加签名载体字段没写初值」这一形，已由 F2 挡住。

### B74 控件几何写进去的数与读出来的数天生差一格 —— VB 侧读数从没被存过（账 #230，**已出：门 #376（run 37547498187、head `a9c47c82`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0**）

读数（探针 `.build/b506geo`，两架构逐字相同）：`txtA.Left = 5000` 读回 **4995**、`.Top = 444` 读回 **441**、`.Width = 7777` 读回 **7770**；设计期写的 `1007,449,3001,247` 读回 `1005,450,3000,240`。原因不在值里而在**读法**：四个 getter 现场 `GetWindowRect` + `vb6_ScalePxToUser` ⇒ 每一次读写都被像素网格重新量化一遍。
语料扫出来这不是边角：105 份设计文件、770 枚控件里 **238 处**几何值不是 15 的倍数（NewTab-test 86、ctrlslider 22、btnfocus 15、ctrltabindex 9、ctrlprop 8…，多半是在别的 DPI 上排出来的），**没有一条存量针在量它们**（夹具读的都是自己写进去的 15 的倍数）⇒ 这一族编得过、链接过、A/B 发码零改动，三头全哑。

改：一处存储 `vb6_GeomCacheRead/Write`（值 + 写它时容器的 ScaleMode；窗口属性名一档一个：`VB6_GeomL/T/W/H` + `VB6_GeomMode`；值存 2v+1 ⇒ 恒为奇数、永不为 0，因为 `SetPropW(…,0)` 等于删属性）。
读侧两道闸，都是**问一句**而不是**列一张清单**：容器的档位换了 ⇒ 那个数不再代表同一件事 ⇒ 投影；按存着的数换算出的像素 ≠ 窗口现在的像素 ⇒ **别人**挪过这枚窗口 ⇒ 投影（ComboBox 建窗时自己补下拉高度、PictureBox AutoSize、SSTab 排版都属这一格，点位散在 RTL 十几处，逐处补失效句就是下一份"抄两遍"的账）。
写侧三个来路都接：四个属性 setter、`vb6_ControlMove`（按 mask 逐档 —— #193 当时只验了 Left/Top 那半边）、`vb6_CreateControl`（设计值）。

**建窗那一档记的是缇，不是容器单位** —— 这一条推翻了我开工前的计划（原打算把 `vb6_CreateControl` 里写死的 `vb6_TwipToX` 改成按容器 ScaleMode 折算，好与 #175 的口径对齐）。实证 `tests/czUI-main/frmDemo.frm`：那枚窗体声明 `ScaleMode = 3  'Pixel`，同时写 `ClientWidth = 6600` 与 `ScaleWidth = 440`（同一块客户区的**两种单位**各写一遍），而子控件 `Width = 6240` —— 按缇是 416 像素（占满 440 宽的客户区，对），按像素就是 6240 像素（放不下）⇒ .frm 的几何恒按缇写；像素档容器里读回像素数（= 投影）本来就是 VB6 的答复，夹具 GC06 把它当**边界**钉住而**不当罪证**。全仓语料里"子控件住在声明了 ScaleMode 的容器里"是 **0 处**，这条口径只能靠那一份真工程的数字自证。

判据：新夹具 `tests/geomcache`（两架构真跑，GC01..GC08 全 True）—— GC01/02 设计值（顶层路 + 容器子控件路）、GC04 赋值、GC05 Move，各带一枚**像素证人**（user32 的 rect + kernel32 的 MulDiv 现算，故意不走产品自己的换算）钉住"存下那个数"没把窗口挪走；GC06 像素档容器边界；GC07 切 ScaleMode ⇒ 回投影、切回来 ⇒ 还是那一个数；GC08 ComboBox 被 RTL 自己加高 ⇒ 缓存被像素闸判废、跟着窗口答。
负控 = 改前那台编译器跑**同一份夹具**：GC01/02/04/05/07 五条 False（raw 就是那批量化后的数），GC03/06/08 两边同数 ⇒ 夹具抓得住这一刀，也没把没动的东西算进账里。
新哨兵 `scripts/check_ctrl_geom_cache.ps1`（第 26 道 [STATIC]）：S1 存储唯一 / S2 四档读写成对 / S3 来路计数 8+4 / S4 建窗记缇（并禁建窗路问容器 ScaleMode）/ S5 四个 getter 各过缓存一次且 `return (int)vb6_ScalePxToUser` 那一形 0 次 / S6 像素闸还在。**六条各用一处假改动证过能红**（植完立刻还原，还原后 PASS 且文件逐字节相同）。
护栏：语料 A/B（BASE = 改前那台**本机冷编** `.build/wt_base230`，见 §B73 那条口径）⇒ inputs=90 **changed=0 / same=90**；会读几何的存量夹具**同一份产物两台编译器对输出**（x64，`.build/b511_geocheck.py`）共 15 枚跑出读数 —— ctrlslider / btnfocus / ctrltabindex / ctrlprop / combofocus / c29listview / ctrlmanifest / ctrlshape / ctrlsstab / ctrlstatusbar / scalemode / sbfont / dcsurf **13 枚逐行相同**，只有两条差异且都是天生抖的数：`modal` 的 `M2 …/busy=0→8` = **在途 tick 条数**（夹具自己的注释就写着"那个界是看负载的"，账 #162 同族；套件钉的是前缀 `M2-returned=Y` ⇒ 不在断言面里），`ve_units` 的 `U-CNT-RAW cnt=3015690→3146762` = 那一行打印的是**容器 HWND 本身**（两边的 `ok=True/True` 都在）。`NewTab-test` / `tabwalk` 本地没跑出读数（exe 名没猜中），`VBFlexGridDemo` 两台都超时不自退 ⇒ 这三枚由门覆盖，而本轮门 #376 绿。

### B75 控件几何的第二份实现撤掉了 —— UC 宿主模型那对 getprop/setprop 现在只问 vb6forms_ctrl.c 的出口（账 #247 = §B75 那一格，**已出：门 #385（run 37570038967、head `a22cb110`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0**）

- **读数**（探针 `.build/probe247/`，同一枚窗体两条路对跑，x64）：`Set f = Me` 之后 `f.Width = 7222` ⇒ 晚绑定那路读回 **7215**、发码那路也读回 **7215**（两条都退化成投影，写进去的数压根没存）；反过来 `Me.Width = 6011` 之后发码读回 6011、晚绑定读回 **6015** ⇒ 同一枚窗口同一个属性两个答案。负控 = 本笔父提交 `571ca6aa` 本机冷编那台跑**同一份夹具**：GC09/GC10/GC12 三条 False（7215 / 3330 / 6015），GC01..GC08 与 GC11 逐行同数（x64 与 x86 两片都是这三条翻红，没有第四行）。
- **病灶**：读侧自己量窗口矩形 + `vb6_XToTwipX`（恒按缇），写侧把四档一起读成缇再整体 `MoveWindow` 推回 —— 所以写 `.Left` 会顺带把 Top/Width/Height 重量化一遍，而两边都不碰账 #230 那张 VB 侧读数缓存。这是 #234（拿 DC）/ #235（画笔色）/ #239（笔位）那一族的第四个样本，住在 RTL 的宿主层而不是 cgen。
- **改法**：四档读写都转调 `vb6_Get/SetControl(Left|Top|Width|Height)`（`vb6forms_prop.h` 进 uc_hostmodel.c 的 include），单位跟着**所在容器的 ScaleMode**（#175 的口径），写进去的数进同一张缓存；写侧从此各档写各档。被这一刀断了调用点的两份自带实现（`vb6_ho_ctrlRect` / `vb6_ho_isForm`）一起删。
- **刻意没并的那一格**：`ScaleWidth` / `ScaleHeight` 仍走宿主自带的 `vb6_ho_clientTwips` + `scaleMode == 3` 折算。它们的另一份权威是 `vb6_GetScaleWidth`（问窗口**自身** ScaleMode），两者是不是同一件事没量过 —— 并错了会把 #175 的「容器档位」搬进「自身档位」，所以留在原样，开工前先按 §B79 那条一起量。
- **语料暴露面**（90 份 emit 捕获里数「晚绑定读/写几何」的调用点，每架构）：**60 处** —— VBFlexGridDemo 31 / Charts 2020 主工程 21 / czUI-main 4 / c29listview 3 / ctrltoolbar 1。Charts 那一处的源形就是 `With CtrlNames(i) : .Left = FW * Rects(i-1).Left / 100`（`ClsResizer.cls:138-142`，唯一调用点 `Form2.frm:581`）—— 而它今天**因为 §B79 一条也不执行**，所以这一刀在 Charts 上是「接线接上但没有流量」；真流量是 czUI 的 `With Parent : .Left/.Top/.Width/.Height`（读侧）。窗体那一档两条路等价是有根据的：表单创建传 `hWndParent = NULL`（`vb6forms.c:396`）⇒ `GetParent` 给 0 ⇒ `ScreenToClient(0, …)` 不改点、`vb6_GetScaleMode(NULL)` 答 1=缇 ⇒ 与改动前逐字节同一。
- **判据** = `tests/geomcache` 升 GC09..GC12 四条（晚写晚读 / 晚写直读 / 像素证人 / 直写晚读；两架构真跑，NEW 12/12 True）+ 哨兵 `check_ctrl_geom_cache.ps1` 加 **S7**（四档两头各问出口恰好一次、`MoveWindow(` 与自带投影助手回潮 = 0）。S7 五处各用一处假改动证过能红（`.build/b583_redproof.py`，植完按 md5 还原，还原后 PASS）。全套 30 道 [STATIC] 零红。护栏：这一刀只动 RTL，`c3rtl.rc` 未增删文件（`check_rtl_resource_ids` 绿），发码面零改动 ⇒ 语料 A/B 天然是 0，所以行为护栏交给夹具两架构 + 门上的 Charts2020 / czUI 两片。
- **本轮自己撞的一条工具口径（值钱）**：**别拿 `git checkout -- <file>` 当「还原我的改动」** —— 它还原到 HEAD，会把这一格尚未提交的刀一起抹掉（本轮真抹了一次，靠重新应用才回来；哨兵的假改动还原一律按字节写回自己存的那份）。
- **边界**：窗体**自身** Left/Top 写后再读那条 AV（§B78）与 Controls 集合答空（§B79）都是改前就有的独立缺陷，本刀一条没碰；GC09..GC12 刻意只走 Width/Height，就是为了不把判据押在会崩的那一形上。

### B76 浮点交给整数目标时**截断**，而 VB6 四舍五入 —— 两条路各错一半（账 #248，**已出：门 #380（run 37556525156、head `09621c38`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0**）

探针 `.build/b515_round.bas`（`Public Sub Main` + `Debug.Print`，本机实测）：`l = 6.73` ⇒ **6**（VB6 = 7）；`l = -6.73` ⇒ **-6**（VB6 = -7）；`l = 7 / 2` ⇒ **3**（VB6 = 4）；`i = 6.73` ⇒ **6**（Integer 同病）。
另一头：`CLng(6.5)` ⇒ **7**、`CLng(7.5)` ⇒ **8** ⇒ 显式转换是 half-up，而 VB6 的取整是 **banker's**（6.5→6）。所以隐式赋值与显式 `CLng` **不是同一个口径**；更要紧的是 `l = 5 / 2` ⇒ 2、`s = 6.5 : l = s` ⇒ 6 这两条看着"对"（与 banker's 恰好一致），其实是截断把 .5 那一半**伪装**成对的 —— 别拿它们当"实现正确"的证据。
发现途径值得复用：给夹具写像素证人时先用了 `Me.ScaleX(v, 1, 3)` 赋给一枚 Long，得 6 而 MulDiv 得 7，两边对不上 —— 换算那两处本身都没错，错在**赋值那一截**（夹具最终改用 kernel32 MulDiv 当证人，见 §B74 那句"不走产品自己的换算"）。
范围没量：哪些目标类型 / 哪些发码分支吃这个隐式转换，要先在语料里数出来再定口径（`l = a / b` 在 VB6 代码里很常见，几何、缩放、分页大小全是这一形）。
**已出的读数与收法（提交 `09621c38`，门 #380 = run 37556525156、head `09621c38`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0）**：

范围量完了：发码侧只有一处吃这个隐式转换 —— `CCodeGen::narrowCheckAssign`，而它按**源类型的位宽**分两种错法。`cgenIntBits(Single)` 答 32、Long 目标也 32 ⇒ 那句 `装得下，不套` 把 `l = 某Single` **整个放过**（C 在赋值处截断）；`Double`（含 `a / b` 的结果）答 64 > 32 ⇒ 套上了 `vb6_ChkLong(...)`，但 double 直接进它的 `int64_t` 形参 ⇒ **C 在调用边界上截断**。同一件事的另一半答案一直躺在 `vb6_CLng` 里（它内部走 `round()`），所以 `l = 7 / 2` 交 3 而 `l = CLng(7 / 2)` 交 4 —— 一个决定抄了两遍，就是 #234（拿 DC）/#235（画笔色）/#239（笔位）/#247（宿主模型几何）那一族。
**语料暴露面**：90 份 emit 捕获里 **36 处**把除法/浮点结果存进窄整型目标（其中 **21 处在 Charts 2020 的主工程**）⇒ 真实排版算式一直差一格。
**收法 = 一条出口、两头同源**：RTL 新增 `vb6_FltToLng(double) = (int64_t)round(x)`（`src/rtl/core/vb6rtl/vb6rtl_conv.c`），`vb6_CInt` / `vb6_CLng` 改成问它；发码侧在那条短路**之前**加一道闸 —— 来源是 Single / Double / Currency ⇒ 先套 `vb6_FltToLng` 再套溢出检查。于是隐式收窄与显式 `CLng`/`CInt` 从此同一个答案。
**判据** = `tests/test_f2lng.bas` 13 条两架构真跑（NEW 13/13 True；**BASE = 本笔的父提交冷编那台，同一份夹具 4/13 ⇒ 九条 False**）+ 发码针 `f2lng_round`（四条必须出现、四条必须不出现，方向两头都实测过）+ 新哨兵 `scripts/check_float_to_int_round.ps1`（第 27 道 [STATIC]，六条各用一处假改动证红）。
**取数口径的教训（已写进夹具头注释）**：判据要选**截断与舍入永远不同**的那一形 —— 第一版循环写成 `k/(2k+1)`（恒小于 .5），在坏编译器上照样全绿；改成 `(4k+3)/4`（恒 .75）才真咬人。`.5` 那一族（`(2k+1)/2`）只钉「隐式==显式」这条不变式、**不钉具体答案**，因为平局口径还没定（见 §B77 ①），钉了就把未定的口径焊死在针面上。
**护栏** = 语料 A/B：BASE 那台**重新冷编到本笔的父提交 sha**（`5fac77b1`）⇒ captures=90 / same=72 / changed=18 / **未归因 0 行**，改动行的形状只有 K1（原本套了检查、现在多一层 helper）224 行 + K2（原本压根没套检查）16 行。上一轮拿更旧的基线（`d8ae8301`）做同一件事时报了 28 行 `vb6_LSet*` 噪声，那全是别人定长串那一刀的 ⇒ **A/B 的 BASE 必须是本笔的父提交**，这是规矩不是偏好。整套 29 道 [STATIC] 零红。

### B77 取整那一刀剩下的两半 —— `.5` 平局的口径（①**仍开，等拍板**）与非裸标识符的目标（②**已出 = 账 #261 / 门 #406；它刻意留下的那条边界已由账 #262 收掉，见 §B90 与 §B77② 末段**）

① **平局的口径**（未开工）：今天 `CLng(6.5)`=7、`CLng(7.5)`=8（half-away-from-zero，`vb6rtl_conv.c` 里 rev37 定过），而 VB6 本体是 banker's（6.5→6、7.5→8）。改就是改**公共答案**（`vb6_FltToLng` 一处），`test_f2lng` 的 F2L09/F2L10/F2L23 钉的是不变式、不用跟着改，但要**补**一条具体平局答案的针。开工前先问用户要不要把 banker's 立成目标口径 —— 这格刻意留在 §B76 之外，就是为了不让一个未拍的口径焊进判据。

② **非裸标识符目标**（账 #261，已出 = 门 #406；它留下的那条边界已收，见本节末段）：`narrowCheckAssign` 第一道门原本就是 `target->kind != IdentifierExpr ⇒ 原样返回`，所以 `arr(0) = a / b`、`u.field = 6.73` 这些形一直走截断。

- **改（收成一处）**：新增 `CCodeGen::narrowTargetTypeOf(Expr*)` —— 「这枚左值是哪档窄整型」只在这里答，答案只许是 Byte / Integer / Long 三档之一，认不出（Array / UDT 整体 / String / Object / Variant / Boolean / 推不出）一律 Unknown ⇒ 调用方原样发，与改前同形。成员那一档问既有唯一出口 `inferUdtFieldVb6Type`，数组元素那一档问 `arrayElemTypes_`（六条声明路填的那张表）；裸标识符那串 `knownByteVars_/knownIntVars_/knownLongVars_ + inferExprType` 兜底**逐字搬过去**，一条没改。
- **两条闸是量出来的，不是想出来的**（这一刀自己带出来的两枚雷，形状 = rev36 那枚「指针送进标量闸」换入口重来）：
  - **空下标 = 整体数组赋值**（Fix 170 的 `dst() = src()` / `bb() = s`）发的是**数组描述符指针**而不是元素值 ⇒ 实测 test_array.bas 改后 `EXIT=0x00000006`、`wa-clone/wa-ub/wa-str/wa-rt` 四行整片不打印。闸 = 先问有没有下标再问元素档。
  - **数组成员的整体赋值**（`With x : .Data = baData`）同一枚雷，而且**档位看不出来**：`inferUdtFieldVb6Type` 的头注释写着「含 Array 标志」，实测语义层从来不带 —— 数组性记在 `mi.isArrayDynamic` / `mi.arraySize`，`mi.type` 存的是**元素**档 ⇒ `Data() As Byte` 答 Byte ⇒ 发出 `vb6_ChkByte(指针)` ⇒ VbQRCodegen 的 Project1 **BUILD-RC=0 而启动期 Unhandled VB6 Error #6**（BASE 同一份源在消息循环里好好待着）。闸 = 那条走查顺手用 `outIsArray` 带回数组性（同一趟，不开第二张表），并**订正那句假注释**。
- **暴露面（A/B 对 HEAD 那台冷编 `382f94a3…`）**：inputs=394（`.vbp` + `.bas`）× 两架构 = **captures=788 / 逐字节相同=752 / 只多套检查=36 / 未归因=0 / run-failures=0**。判据形状 = 把所有 `vb6_Chk*( … )` 与 `vb6_FltToLng( … )` 括好配对拆掉后**两侧必须逐字节相同**（比"数差异行"强：它同时证明别的什么都没动）。新增包裹数（两架构合计）ChkLong 1894 / ChkInt 528 / ChkByte 440，其中 **673/架构 走那一份取整出口** —— 比 #248 量到的 36 处大一圈，因为那一轮只数了裸标识符。**两架构计数逐条同号** ⇒ 这一刀不按架构分家。改动半径 18 份输入：Charts 2020 六份工程、VBFlexGridDemo 五份、VbQRCodegen 两份、czUI、dbgdlg、BalloonTooltips、test_array/test_m5/test_f2lng。
- **判据** = `tests/test_f2lng.bas` 从 13 头扩到 **30 头**：F2L14..22 六形新覆盖（成员 Long/Integer/Byte、成员被 Double 表达式喂、模块级 UDT 成员、静态 Long/Byte 数组元素、动态数组元素、UDT 数组元素的字段）；F2L23 = 成员侧的「隐式 == 显式 CLng」不变式（平局方向照①的口径**不钉答案**）；F2L24..26 = 不许动的证人（整除源 / 整数源 / Double 成员）；**F2L27/29/30 = 上面那两条闸**（整体数组赋值两种拼法都必须保持裸形，改坏就 error 6）；**F2L28 = 这一刀刻意留的边界**（成员数组的**元素**，callee 不是裸标识符 ⇒ Unknown），钉成 6 并在夹具注释里写明「§B90 落地那天这条必须换成 7」。负控 = HEAD 那台跑同一份夹具 ⇒ **F2L14..23 十条两架构全 False**（`raw=6,7 / 1,2 / 6,7 / 6,8 / 2,3 / 6,7 / 6,7 / 3,4 / 6,7 / 0,-1`），改后 30/30 全 True（x64 与 x86 各一遍，`EXIT=0x0`）。
- **发码针两头** = `f2lng_round` 扩到 10 必须出现（含 `r.Pb = vb6_ChkByte(vb6_FltToLng(...))`、`VB6_SA_AT(uint8_t, bArr, 1) = ...`、`VB6_SA_AT(vb6_type_RectF2L, ur, 0).Px = ...`）+ 5 条 BASE 裸形必须不在 + **3 条"闸被拆了才会出现"的形状两头都不许在**（`dst2 = vb6_ChkLong(vb6_ArrayAssign1D(...))` 等）。
- **哨兵** = `check_float_to_int_round.ps1` 加 S7（定义/声明/被问各恰好一次；体内必须问 `inferUdtFieldVb6Type(target, &…` 与 `arrayElemTypes_.find(`；必须拒 `fldIsArray`；必须拒 `positional.empty()`；不许答 Boolean）与 S8（旧那道 `kind != IdentifierExpr → return cValue` 闸门计数必须为 0；`narrowTargetTypeOf` 除定义/声明外无消费者）。**八处假改动各处证过能红**（N1..N8），每次按 md5 原样还原（`b2e3439a…` 逐字节回原）。

- **② 的最后一格已由账 #262 收掉（同批）**：#261 当时把「成员数组的元素」（`p.Pixels(0)` 这种 callee 不是裸标识符的左值）刻意留在闸门外，理由记在原注释里 = **那一形的左值本身还是错的（多维成员被折成一维）**。#262 把左值修对之后，`narrowTargetTypeOf` 的数组档现在也接 callee 是成员访问 / With 成员，并且**必须问出数组性**（`if (!fldIsArray) return Unknown`）才答档位 —— 于是 `p.Pixels(0) = 6.73` 从截断 6 变成 round 7（夹具那一头从 `28-...-open-defect`=6 翻成 `28-member-array-elem-rounded`=7，另加 F2L31/32 钉多维的两个相邻格）。**整体赋值两种拼法仍保持裸形**（F2L29/30 一字未动 —— 那是 VbQRCodegen 那枚启动期 error 6 的闸）。① 那条 `.5` 平局口径**仍然开着等拍板**，本批没有动它的方向。

### B78 `& Me.Left` 把裸 int 交给 BSTR 槽 —— 对象位那一条权威不认 MeExpr（账 #249 = C29-GE-d，**已出：门 #390（run 37583390188、head `84d55a7e`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0**）

- **前提订正**：本节原记「窗体**自身** Left/Top 赋值之后再读那个属性就 AV」—— 实测**跟写过没有关系**。探针 `.build/probe249/PForm.frm`（十二相，一步一句，每相单独起进程）：`Debug.Print "x=" & Me.Left` 作为定时器里**第一条**语句就崩（rc=0xC0000005，`av read target=` 一个野地址），而 `v = Me.Left` 赋给 Long 两架构都正常；同一枚窗体在 `&` 里读 `.Top` / `.Width` / `.Height` 全打得出数。所以分家的轴不是「写后读」，是**「属性名撞不撞返回 String 的 VB 内置函数」** —— 与 #229 / #231 同一族，误导方式也一模一样。
- **根因一行**：`inferExprType` 的 MemberAccess 那一档先问 `CCodeGen::ctrlTypeOfMemberObject`（#229 收成一处的那条「对象位是不是一枚窗体控件」），可它只认 `控件名.属性` 与 `控件名(i).属性` 两种形态 —— `Me` 是 **MeExpr**，它答「不是」 ⇒ 成员名掉进兜底那条「按裸名查模块符号」 ⇒ 命中内置函数 `Left`（返回 String）⇒ 类型答 String ⇒ `wrapToBSTR` 直通不套转换 ⇒ `vb6_BSTR_Concat(vb6_BSTR_FromStr(L"L"), vb6_GetControlLeft(vb6_hwnd_X))`，把 int32_t 当 BSTR 解引用。窗体**名**那一形本来是通的（`cgen_form_ctrl_registry.inc:10` 把窗体名也登记进 knownFormControls_，类型 Form），缺的只有 MeExpr 这一档。
- **改**：只教那一条权威认 MeExpr（窗体模块内、knownFormName_ 非空 ⇒ `FrmControlType::Form`）。**这个函数不答任何具体类型** —— 属性类型仍只出自 `controlPropType` 那一张表，而那张表的第一档（left/top/width/height ⇒ Long）当年就是为躲这对内置函数名设的。`.ctl` 里 `Me.` 那一形不在这次闸口内（`isFormModule_` 为假 ⇒ 照旧回落），**没实测过**，别按这条推断它。
- **判据三头**：夹具 `tests/geomcache` 升 GC16 —— **那一行就是崩溃现场**（父提交的产物里整行不出现，`GC-DONE` 跟着没）+ 值一起钉（`L<数字>`）；再加两枚发码针 `gc_emitc_meleft_wrap`（套了 `vb6_CStrLong(...)` 的整条语句必须在）与 `gc_emitc_meleft_raw`（父提交那条裸形必须不在）。只钉读数会放过「两边都不套」那一族，只钉发码会放过「值错了但没崩」。
- **负控** = **本笔父提交** `9936cf3f` 本机冷编那台跑同一份夹具 ⇒ 15 行、`RUN-RC=0xC0000005`、GC01..GC15 全 True（含 GC15 raw=5/5，即 #254 在效）而 GC16/GC-DONE 双双缺席；改后 x64 `True/L390`、x86 `True/L1170`，其余十五行两边逐字相同。探针五样形状（只读 / 直写后读 / 宿主写完直读 / 宿主写一档读一档 / 四档读写）改前全崩、改后全打得出数。
- **护栏** = 语料 A/B 对本笔父提交那台，emit 捕获 inputs=90 **changed=0 / same=90** —— 这一刀改的是发码闸口，可全语料没有任何一处 `& Me.<几何>` ⇒ 存量零暴露，**这一格只能靠新夹具守**（geomcache 不在那 90 的名单里，它的发码形状由上面两枚针直接钉，父提交与本笔各验过一遍）。
- 工具坑（栽过一次才记的）：PSParser 在实参位置的 `@(` 续行里按**字面**数括号 —— 单引号串里不配平的括号会把后面那行独立的 `)` 变成野 token ⇒ **整份 run_tests.ps1 ParserError，而退出码照旧 0**。所以两条发码针都用整条语句（括号自然配平）作串；改完 run_tests.ps1 必跑一次 Tokenize 数错误，别只看它跑绿。

### B79 窗体的 `Controls` 集合答空（窗体自身的 Count 另有硬填 0）—— Charts 的 ClsResizer 在唯一调用点上静默不做任何事（账 #250 = C29-GE-b，**已出：门 #389（run 37581465287、head `9936cf3f`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0**）

- **改前读数**（探针 `.build/probe247/` v8 + 同一份夹具两架构对跑）：一枚带 TextBox + PictureBox 的窗体上 `Me.Controls.Count` = **0**、`For Each o In Me.Controls` **零条**、`Me.Controls.Item(1).Name` = **空**、`.Width` = **0**（一律空值，不崩）。
- **根因一句话**：`uc_hostmodel_getprop.inc` 的 Controls 档问的是 `vb6_ho_setVariantDispatch` —— 而那一枚**刻意**只交 Empty（注释写着：让字体代理那一路继续走 Nothing 分支，别对裸指针调 Release）。rev14 为了「真交对象」另立了 `vb6_ho_setVariantObject`，**但只把 `Item` / `Add` 两个调用点改过去** ⇒ Controls 交回 NULL。而下游三条其实**全都写好了**：`vb6com_foreach.c:35/123/175` 认这个集合，`uc_hostmodel_call.inc:35/53` 的 Item 用的正是新出口 —— 断的只有最上面那一档。
- **改**：Controls 那一档换 `vb6_ho_setVariantObject`；**Font 代理那一档保持 Empty 原样**（它依赖 Nothing 分支，别顺手并表 —— 那一格与 #247 的「并到唯一出口」方向相反，是有意的）。
- **改后读数**：`Controls.Count` = **5**、`For Each` 走完 **5** 条；其余十四行与改前逐字节相同（x64 与 x86 两片都是这样）。负控 = 本笔父提交 `a22cb110` 本机冷编那台跑**同一份夹具** ⇒ 只有 GC13/GC14 两条 False raw=0 → True raw=5，别的行一条不差。
- **与 #247 的先后**：上一刀撤掉的那份第二实现，是**在这一刀之后才第一次被走到** —— `For Each oCtrl In oForm.Controls` ⇒ 成员 ⇒ `.Left = …` 落宿主模型 setprop，而那条路今天已经是转调四个出口。所以 Charts 的 resizer（`ClsResizer.cls:92` 的 `For Each`，唯一调用点 `Form2.frm:581`）从今天起才有流量；#247 的台账里那句「接上但没有流量」到这里才闭合。
- **判据** = `tests/geomcache` 的 GC13/GC14 两条（钉「集合交得出对象 + 枚举走得完」，负控两面验过）。**不新立哨兵**：这一格的失效形状被那两条读数直接咬住，翻回去就红。
- **没修的那一半另立 §B81（账 #252）**：枚举出来的成员**答不出自己的 VB 名**（`o.Name` 空、`TypeName` 答 Control）。数出来的 5 条是真的，名字那一档还是空的。

### B80 `vb6_UC_ParentMove` 把**缇**直接交给 `MoveWindow`（账 #251 = C29-GE-h，**已出：门 #400 = run 37652210198、head `d199adc0`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0**）

- **一句话的病**：`src/rtl/core/vb6forms/uc/uc_host.c` 里那一支 `MoveWindow(fw, left, top, width, height, TRUE)` 四个实参一个换算都没有，而这四个数是 **VB 侧的量纲**（容器是窗体时 = 缇）。发码那侧 `UserControl.Parent.Move l,t,w,h` 由 `cgen_base.cpp:123` 的文本重写折成 `vb6_UC_ParentMove(l,t,w,h)`；czUI 全屏/恢复那一支（`czUI.ctl:1163`、`:1185` 那句带 `Screen.Width` 的）存进 `m_Saved*` 的数来自 `With UserControl.Parent : .Left = …`，而 #247 之后那一读走的是 `vb6_GetControlLeft` 那对进出口（缇）⇒ 整条链的量纲从头到尾是缇，只有落地那一步按像素摆 ⇒ 607 缇落在 607 像素上（#175/#247 那一族的老形状，只是这次住在 call 那一路）。
- **同一格还欠着另一头**：`MoveWindow` 不填 #230 那张 VB 侧几何缓存 ⇒ `Parent.Move` 之后读 `Me.Width` 拿到的是窗口位置的投影，不是写进去的那个数（属性赋值与 Move 是同一件事的两个来路，#230 钉的就是这一条）。
- **改（照 §B75 的收法，不在这一层再拼一份摆位）**：转调 `vb6_ControlMove((void*)fw, (double)left, (double)top, (double)width, (double)height, 15)` —— 单位（`vb6_ContainerScaleMode` + `vb6_ScaleUserToPx`）、坐标空间（子控件走父客户区、顶层窗体走屏幕）、#230 的缓存写入、以及尺寸真变时的 Resize 一律只在那一处答。本单元**不** include `vb6forms_prop.h`（三头混一个 TU 会撞，沿用 `vb6_ControlMove` 自己 `extern vb6_InvokeFormResize` 那条纪律）；头文件 `vb6rtl_userctl.h` 那条映射注释补了量纲一句。
- **判据** = `tests/ve_units` 给 UC 补 `MoveParent`（一行 `UserControl.Parent.Move l, t, w, h`）+ 窗体侧两行：`U-PMOVE-RAW` 既钉四档 VB 侧读回 = 写进去的那四个数，也钉窗口对 `MulDiv(缇, DPI, 1440)` 的**差值为 0**（证人 = user32 `GetWindowRect` + kernel32 `MulDiv` + `GetDeviceCaps` 的真实 DPI，不经产品换算 —— #230 那轮的课：换算证人不能走 `Me.ScaleX`；**钉差值而不钉像素绝对数** ⇒ 换 DPI 不漂）；`U-PMOVE` 收成一条布尔。四个数 607/451/2407/1811 都不是 15 的倍数 ⇒ 两头互相冒充不了。两架构真跑 `U-PMOVE-RAW l=607 t=451 w=2407 h=1811 dx=0 dy=0 dw=0 dh=0` + `U-PMOVE=True`，其余 26 行逐字不变。
- **负控** = BASE = 本笔父提交 `ad16bc9c` 本机冷编那台（`.build/C3_base251.exe`，md5 `2832d2a6…`）跑**同一份**夹具 ⇒ 两架构都是 `l=9105 t=6765 w=23340 h=14700 dx=567 dy=421 dw=1396 dh=859`、`U-PMOVE=False`（9105 = 607 像素折回缇），其余 26 行两台同数 ⇒ 归因不会串台。
- **护栏** = 语料 emit A/B（`.vbp` + `.bas` × 两架构，inputs=392 / captures=784）**same=784 / changed=0 / 未归因=0 / run-failures=0** —— 这一格是纯 RTL（发码文本里没有它），这个 0 是**预期**且只证明「cgen 一字未动」，真伤由上面那两行钉（同型读法见 §B64/#234：只改 RTL 时 A/B 一个字节都不差）。哨兵 `check_ctrl_geom_cache.ps1` 长 **S8**（函数体内那条转调语句恰好一次 + mask 15 + 体内 `MoveWindow(`/`SetWindowPos(` 零次），三处假改动各证过一次能红（改回裸 `MoveWindow` ⇒ 双条红；mask 15→7 ⇒ 转发条红；整条撤成 `(void)0` ⇒ 转发条红），每次按 md5 原样还原；30 道 [STATIC] 全 RC=0；`run_tests.ps1` PSParser 零错。
- **语料暴露面** = 真流量 2 处（`czUI.ctl:1163` 恢复原位、`:1185` 全屏），都在 `ToggleFullScreen` 里 ⇒ **要人手点一下才跑**，而门上 czUI 那格的读数只有「窗口起没起」（`Test-GuiVbp` 没给 `DumpMinColors`）⇒ 这一格在门上永远不会红，判据只能自己造；除此之外全仓 0 处。
- **刻意没接的两格（都带实测）**：① `Parent.Move` 只认四档齐全那一形 —— `vb6_UC_ParentMove` 的签名是四个 `int32_t`，少给实参（VB6 允许 `Parent.Move l, t`）实测发成参数不足的调用，`ucUnitTwip.c(143): error C2198: too few arguments for call 'vb6_UC_ParentMove'`、编不过；语料 0 处 ⇒ 另立账（要接得让发码那一跳带 mask，而不是在 RTL 再开一个重载）。② `With UserControl.Parent : .Move …` 那形本刀未量。
- **风险记账**：四档全交（mask 15）是唯一形态；转调之后 `Parent.Move` 会像属性赋值一样填缓存，并在尺寸真变时触发 `Form_Resize` / `<Ctrl>_Resize` 排队 —— 这是**新行为**（VB6 本人就是这么答的，此前走裸 `MoveWindow` 时它们一条也不发）。发码面零改动 ⇒ 语料里唯一会跑到它的工程（czUI）只在人手触发时才有行为变化。

### B81 控件的 VB 身份从没登记过 —— `.Name` / `TypeName` / `Controls(名字)` 三条读法一起落空（账 #252 = C29-GE-f，**已出：门 #392 = run 37605281964、head `9ee3d0ae`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0**）

读数（#250 那一刀的同一轮探针，改后）：`For Each` 走完 **5** 条，但第一条的 `CStr(o.Name)` = **空串**，`TypeName(o)` = **Control**。集合本身是通的，缺的是成员的身份。
两处叠在一起：① 发码那条创建路把 `vb6_CreateControl` 的**第二形参（controlName）恒传空串**（emit 实测：`"EDIT", ""`）；② 标准控件**压根不进宿主登记表** —— 全仓只有一个调用点 `vb6rtl_system.c:549` 给**窗体**注册过（`vb6_HostObj_Register(hwnd_, NULL, "Form", 1, -1)`）。于是 `vb6_Host_GetProp` 的 Name 档（`h && h->name[0] ? h->name : (r && r->ctrlName[0] ? … : L"")`）对窗体子控件一律交空。
后果面：`Controls("txtDoc(0)")` 那档按名查找（`uc_controls.c:91` 的桥已经写好）、`ClsResizer` 的名字比较、以及任何 `ctrl.Name` 写法今天都拿不到数 —— 所以这一格是 #250 的**下半场**，做完 Controls 集合才算真通。
开工先定两格，别一上来就发码：① `Index` —— 非数组控件今天答 **-1**，注册时若图省事传 0，就是把存量答案改了（必须保持 -1）；② `typeName` —— 登记表里一空，`TypeName` 就从 "Control" 变成真类型名（VB6 语义更对，但那是**公共答案的改变**，要先数有多少条存量针在钉 `TypeName`，与 #159 那张宿主伪成员表对一遍再动）。修法本身在 cgen：把 VB 名（和后面的类型名）从 `cgen_form_create_controls.inc` 递进创建/注册那一步，而不是在 RTL 里猜。

- 2026-10-07 账 #255 那轮的补读（同一份窗体、枚枚成员各答各的之后）：**身份**那一半已经不是本账的事了 —— 三枚子控件的 `.hWnd` 是三个数、`.Left` 跟着各自的数、经 `For Each` 写 `.Width` 落在自己窗口上（GC17..GC19 已进门禁）。剩下的仍是名字面：`.Name` 空、`TypeName` 答 Control、成员那条读 `.Tag` 也空。
- **2026-10-07 第二轮读数（#255 落地之后，探针 `.build/probe255/` 模式 N）—— 身份面缺的到底是哪几块**：成员 hWnd 已各答各的（本账原来那条「成员答不出身份」的前提要收窄），剩下的缺口是：① `.Name` **两条路都空**（不只晚绑定：直读 `txtA.Name` 也交空串）⇒ 不是「成员没有身份」而是「创建时压根没记名字」；② `TypeName` 直读答 **String**、晚绑定答 **Control**，VB6 两边都该答 `TextBox` —— 这一条有**真流量**：Charts 的 ClsResizer 按 `If TypeName(CtrlNames(i)) = FBuf(j).CtrlTypeName` 挑字体档，两边恒不等 ⇒ 排版算式即便接通也一条也配不上（#255 只把「写给谁」修对，「是不是这型」还没对）；③ `.Tag` 只有晚绑定路空（直读 `txtA.Tag` 交 TAG-TXTA 正常）⇒ 晚绑定的字符串属性读法另有一问（与 #88 同族）；④ `Me.Controls("picP")` 交回 NULL。
- **可开工的形态（本轮找到的单一存放点）**：`vb6_ucHo[]` 每格**已有** `name` / `index` 字段，Fix 148 的按名查找 `vb6_UC_ControlsItemByName` 就是拿它比对 `_wcsicmp(r->name, base)` —— 只是标准控件创建时**从没填过**（两条创建路把 `vb6_CreateControl` 的第二形参当 caption 传，而 RTL 那个形参名叫 controlName，两边读反了）。所以这一格 = 「**两条创建路各填一次那份现成的记录**（含类型名），四个读面（直/晚 × `.Name`/`TypeName`、`Controls(name)`、`.Index`）都改问它」，不需要新开表。填表时守住实测：非数组控件 `.Index` 今天两条路一致答 **-1**，别顺手传成 0。


- **改（实际落地的样子）**：读侧一字未动 —— `vb6_ucHo[]` 本来就带 `name`/`typeName`/`index` 三格，四条读法（晚绑定 `.Name`、`.Index`、`vb6_Host_TypeNameOf`、`vb6_UC_ControlsItemByName` 的 `_wcsicmp(r->name, base)`）早就只问这张表；缺的是**创建时没人往里填**（全仓唯一一处 `vb6_HostObj_Register` 在 `vb6rtl_system.c`，登记的是窗体，传的 name 还是 NULL）。于是两条创建路各发一次既有出口 `vb6_HostObj_Register((void*)句柄, VB名, VB类型名, 0, 下标)`：顶层用该文件现成的 `dsHwnd`，容器子控件走既有出口 `ctrlHwndExprForInit(child)`。类型名交原样串（`VB.TextBox`），去掉 `VB.` 那一档仍由 RTL 那一处统一做。
- **顺带把 `VB6_UC_MAX_OBJ` 128 → 512**：这张表是**进程级**的，实测语料 VBFlexGridDemo 182 枚控件、Charts 2020 164 枚（单文件最多 73 枚），128 会**静默装不下**（满了直接 return ⇒ 那几枚退回答不出名字的旧行为，正是本格要治的病）。
- **判据** GC20..GC23（夹具 `tests/geomcache`，两架构真跑 23/23 True）：GC20 两头钉（成员自报 `.Name` 指到的 hWnd 必须就是发码那条路自己那枚 `txtA.hwnd`）；GC21 `TypeName(成员)` = TextBox（Charts 就靠这个数配字体档）；GC22 钉「非数组答 **-1**」——专门拦以后有人填表时顺手传 0；GC23 `Controls(picP)` 找得回同一枚窗口。负控 = 父提交 `84d55a7e` 冷编那台 ⇒ GC20 raw=0/24118394、GC21 raw=[]、GC22 raw=-999、GC23 raw=0/8848806，其余十六行逐字相同。
- **护栏**：语料 emit A/B（BASE = 本笔父提交）inputs=90 / changed_files=56 / added=490 / removed=4，**逐行归因 = 身份登记 486 行 + 上一笔 #255 的 unpack 包裹（4 加 4 删），OTHER=0**；发码针 `gc_emitc_identity_both_create_routes` 钉两条（顶层 `txtA` 与容器子控件 `lblP` 各一条 —— 少一条就是「容器里的控件答不出名字」那种半通）；24 道 [STATIC] 零红。真工程读数（不作判据）：Charts x86 窗口 145ms 出现、进程存活、stderr 仍 0 字节；VBFlexGridDemo x86（控件最多那台）编得过、跑满 60s 观察窗不退。- **收线后的下一格读数（2026-10-07，枚 Charts 2020 主窗体的 14 枚子窗口）**：用 `.temp/yqt_children.ps1`  对跑改前 / 改后两枚产物 —— **每一枚的尺寸逐格相同，位置整体平移 (+182,+182)**（那是两扇窗体自身摆放不同，  `GetWindowRect` 是屏幕坐标）。⇒ 读数只能说明「排版器至今仍未动过任何一枚 UC 的相对布局」，**不能**当作 #252 无效的证词（它把身份面修好了，CI 判据在那儿）。- **读数回来了，而且把顺序定死了（探针 `.build/probe255/` 模式 P）**：窗体**自身**的 `.Name` 两条路都空（`Me.Name` 与 `Set g = Me : g.Name` 都交空串 —— 登记窗体那一处传的是 name=NULL），而 `Container` 这一档在宿主模型里**压根不存在**（`TypeName(o.Container)` 答 **Empty**，`o.Container.Name` 空）。于是排版器那句 `If oCtrl.Container.Name = oForm.Name Then` 今天是靠**「空 = 空」真空通过**（实测 gate=True）。
- **账 #257 已落地（2026-10-07，本节上面那两条「开工前」的读数与顺序就此闭合）**：`Container`/`Parent` 一档 + 窗体自身的名字**同一批**补齐，落地形状、判据 GC24..GC28、负控与 A/B 读数见 §B86。本节仍开着的只有一条：`vb6_HostObj_Register` 那条「同 hwnd 已登记就补齐、只准第一次说话」的病是**登记时序**的坑（UC 宿主那条路同享），日后若又出现「后到的登记被吞」，先问 §B86 的③。

- **风险记账**：登记之后标准控件进 `vb6_ucHo[]`，于是 `vb6_uc_collectChildren` 的第一趟（按登记表次序）不再为空 —— **枚举次序可能从「窗口 z-order」变成「创建次序」**。本夹具的判据全是次序无关的（计数 / 求和 / 遍历全部），次序若真的动了，只有 tabwalk / frmevents 那一族会在门上现形。

### B82 窗体**自身**的 `Count` 那一档硬填 0 —— Controls 集合接通之后 Charts 2020 启动期错误 9（账 #254 = C29-GE-c，**已出：门 #389（run 37581465287、head `9936cf3f`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0**）

- **症状与因果**：门 #387 / #388（head `2883fcbc` / `ad33e452`）两片红的都是 `Tests (vbp #3)` 里唯一一条 `[GUI] Charts2020 ... FAIL (Main window not available within 5s)`。本地一跑就戳穿了这句报法：同一份产物 **123ms 自己退了**（`Unhandled error 9: Subscript out of range`，`C3_SA_TRACE=1` 给 `[SA] elem access out of range: arr=010E8768 idx=0 lb=0 ub=-1` = 一个空数组的第一格），而 `Test-GuiVbp` 的 5s 观察圈对 `HasExited` 也走同一条 throw —— **Fix 189 只补了 AutoExitSec 那一路**，早退在这条针上长得跟超时一模一样。
- **根因**：#250 把窗体的 Controls 集合接通之后，`ClsResizer.cls:78 SaveControlsPositions`（唯一调用点在 Form2）第一次有流量，而它给格子数组写的是 `ReDim Rects(oForm.Count - 1)`。那个 `Count` 走宿主模型 `uc_hostmodel_getprop.inc:76`，那里一直**硬填 0**（旧探针也打过：`Me.Count`=0）⇒ 数组是空的 ⇒ `For Each` 第一格写 `Rects(i)` 越界。§B75 那句「接上但没有流量」到这里才闭合：接上之后走的第一件事就是崩。
- **改**：那一档转调新增的 `vb6_UC_ControlsCountOf(formHwnd)`（`uc_controls.c`，只包一层 `vb6_uc_collectChildren`）⇒ VB6 里 `Form.Count` 与 `Form.Controls.Count` 是同一件事，数法只此一处；宿主模型三档 Count（集合 / RTL Collection / 窗体自身）各转调自己的唯一出口，没有一档自己数。
- **判据 GC15 两头钉**（`fc = n` 且 `4 <= fc <= 8`）：第一版只写等式，拿 #250 之前那台跑（两个数都是 0）GC15 反而 **True** —— 恒等式放过「两个都坏」，必须补范围。负控 = 本笔父提交 `a22cb110` 冷编那台 ⇒ GC13/GC14 False raw=0、**GC15 也 False**（raw=0/0），其余十二行两边同数；真工程证人 = 同一份夹具两台对跑，只有编译器不同：改前 WINDOW-MS=-1 / 123ms 错误 9 退，改后 WINDOW-MS=131、活着、无 crash 文件。
- **护栏** = 语料 A/B（BASE = 本笔父提交）inputs=90 **changed=0 / same=90**（纯 RTL 刀，发码零改动）+ 30 道 [STATIC] 零红。
- **本刀没碰、顺手量到的下一格**：① 枚举成员经 `Collection`（`CtrlNames(i)`）取回之后再写几何，stderr 打 `vb6_ComSetProp: property "Top" not found` 一族 —— 接收者不再被宿主认得，属 §B81（账 #252）的下半场；② resizer 走通之后，缇/像素那一档在 `vb6_UC_ParentMove` 上还欠着（§B80）。

### B83 modal 的跳格判据吃在途 tick —— 相位推进前加一道真实间隔闸（账 #253，**已出：第一半由门 #389 在 CI 上核好；第二半随账 #249 那笔，清掉过期负控文案**）

- **读数**：门 #387 的 CI 工件与 #386（绿）逐档比对，实质差的是 `MW2=cmdX → txtSecond`、`MW-new=4/repeat=txtMain/hops=3 → repeat=txtSecond/hops=1`，同一片里 `busy=6 → 17`（模态返回时排队的 WM_TIMER 条数），而 `MW-seq` 两片一致 ⇒ 四枚站都走到了。
- **机理** = #162 / #213 那一族：定时器在途的拍会**连着排空**，而 VK_TAB 是 post 给泵**异步**消化的 ⇒ 两拍挤在同一瞬间时读到「还没动」，第一次回头被提前判定。本地两台编译器（父提交与本笔）跑同一份产物都稳定给 4/3 ⇒ 这是判据与拍序耦合，不是产品行为回退。
- **改（第一半）**：`tMain_Timer` 里相位推进前加一道真实间隔闸 —— `Timer` 差 < 0.03s 的拍不计相位，仍计入 `gStray`。本地 x64 与 x86 都是 `MW-new=4/repeat=txtMain/hops=3` + `MW-seq` 一字不差；门 #389 的 modal 工件确认 CI 上回到这条（`busy` 那条数本来就不是判据）。
- **第二半**：夹具与 runner 里「`C3_OCX_NO_DLGMSG=1` 关掉泵里的 `IsDialogMessage` ⇒ 红 `MW-new`（回到 1）」那条负控文案**已过期**。实测（本笔父提交与本笔两台一致）设与不设 `MW-new`/`MW-seq` 一字不差，而 `C3_OCX_NO_TABNAV=1` 仍会红（`MW-seq` 翻成 z-order `txtMain,cmdY,cmdX,txtSecond`）。原因在 `vb6forms.c:884`：账 #163 之后主泵里 `vb6_TabNavKey` 排在 `IsDialogMessageW` **之前**，VK_TAB 早被自研导航器吃掉，那个开关只剩非 Tab 的几条对话框键。两处注释改成实测口径。

### B84 集合里的 `With` 接收者与 `For Each` 成员的身份别名 —— Charts 排版算式落不地的两刀（账 #255 = C29-GE-e，**已出：门 #391 = run 37594343602、head `c6c6de2d`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0**）

- **读数**：门 #390 之后 Charts 2020 x86 启动期 stderr 168 条 `vb6_ComSetProp/GetProp: property "…" not found`（六档各 28 条：Left/Top 只写、Width/Height 又写又读）。探针 `.build/probe255/`（八相，一条语句一个 marker，最后一个 marker 就是案发现场）分家出两刀：① `With 集合(i)` 的接收者是那次默认 Item 调用返回的 calloc VARIANT 的**地址**；② 枚枚成员解回**同一枚**对象（三枚的 `.hWnd` 同为 4326406、`.Left` 一律 120 = Timer 的设计值、`.Width` 一律 0，写 `.Width` 只落进 Timer 的存属性）。
- **机理②**：宿主成员交给调用方时包着一枚真 vtable 的包装器（为的是栈上 VARIANT 收尾的 `VariantClear`→Release 不野调用），登记在 `g_uc_wraps[]` 这张「包装器→原对象」表里；这张表**只追加、从不受让** —— `HW_Release` 归零 `free` 之后块地址被下一枚包装器原样复用，而 `vb6_UC_UnwrapHost` 从第 0 格扫表，先撞上那条早已释放的旧登记。这一族的症状是「全都答同一枚」，不是崩。
- **改**：① 「COM 调用的结果当**对象引用**用」收成唯一出口 `CCodeGen::comObjectRefFromCallExpr`（`cgen_util_com.cpp`）：With 三条接收者分支与 Set 那一段就地手改字符串前缀（同一决定的第二份抄本）一起转调它 —— `vb6_ComCall(` ⇒ 改名 `vb6_ComCallObject`（一体化 Unpack+Free，一字不变），`vb6_ComCallByDispid(` / `vb6_ComGetProp(` ⇒ 就地包一层 `vb6_ComUnpackObject(...)`；交回的已经是对象的那两支（`vb6_ComCallObject` / `vb6_ComGetObjectProp`）**不进清单**。② `HW_Release` 引用归零时先 `vb6_UC_DropWrapPair`（末位填补，查表按指针、次序无意义）再 `free`。
- **判据**：`tests/geomcache` 升三条 —— GC17 钉「成员互不相同」（每枚成员的 `Width` 加总，跟四枚控件自己之和比：只有各答各的才对得上；picX 不进断言，它是像素制容器，同一个数在两种单位里不是同一个值）；GC18/GC19 钉「经集合的 With 写落地」，**两头**（控件自报的数 + 窗口自己的像素）。负控 = 本笔父提交 `84d55a7e` 本机冷编那台跑同一份夹具 ⇒ `GC17 False raw=4006/7918/5`、`GC18 False raw=3007,2003,1507`、`GC19 False raw=200,82,134`，其余十四行同数、照旧 `GC-DONE`。
- **CI 复核**：CI 侧两枚发码针按名各自 PASS（`[EMITC-SHAPE] gc_emitc_with_member_unpack` 在 vbp #3、`[EMITC-ABSENT] gc_emitc_with_member_raw` 在 vbp #4），夹具两架构也 PASS，工件里 GC17..GC19 的读数与本地逐字相同（7918/7918/5、1234×4、82,82,82）。
- **护栏**：emit A/B（BASE = 本笔父提交）inputs=90 **changed=2 / same=88**，两份各只 2 行、变的正是那一层 `vb6_ComUnpackObject(...)` 包裹（`tests_Charts 2020_Proyecto1` 的 x64 与 x86 捕获）；两枚发码针 `gc_emitc_with_member_unpack`（套了的那条整语句必须在，实测 2 处）/ `gc_emitc_with_member_raw`（父提交那条裸形必须不在），x64/x86 两份 emit 都核过；24 道 [STATIC] 零红。Set 那一侧的转调由「语料 0 处命中 + 这份 A/B」共同担保它是惰性的。
- **刻意不当判据的那条**：Charts 的 `not found` 由 168 → 0（stderr 文件 0 字节、窗口 141ms 出现、进程存活）—— 「没有诊断」可以是另一种沉默，所以它只记作现场读数，判据住在 GC17..GC19。这一条在本轮差点被当成结论：第一版夹具 `coll.Add txtA` 直接 AV，顺着它才量到下面的 §B85，而「168→0」那时看着像已经修好了。

### B85 晚绑定实参的**装箱档**不跟着上一步真正交出的值走 —— BSTR 进 `vb6_ComPackObject` ⇒ VT_DISPATCH 野 Release 崩（账 #256 = 提交 `799f0072` + 台账 `3ebbc641`，**已出：门 #396 = run 37632990355、head `3ebbc641`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0（wall ≈10m45s）**）

- **一句话的病**：裸控件名出现在**晚绑定调用的实参位**时，发码会把它折成默认属性读数（一枚 BSTR 或一个 int），而**装箱那一层问的是另一张表** —— `comPackExpr` 的 IdentifierExpr 分支里 `knownObjectVars_` 对任何控件名都答 Object，于是发出 `vb6_ComPackObject(vb6_GetControlText(vb6_hwnd_txtA))`：里面是串，外面按 dispatch 存进 VT_DISPATCH，语句收尾那个栈上 VARIANT 走 `VariantClear` ⇒ 对 VT_DISPATCH **无条件** `Release` ⇒ 按 BSTR 头几字节解 vtable。探针 `.build/p256`（`coll.Add txtA, ` 一条带键的 Add）两架构都 EXIT=0xC0000005，崩点就在那条 Add 上（x86 `av read target=0x4`）。这是 #122 / #123 / #88 那一族（**装箱档与值类型对不上**），不是「对象位 vs 值位」那一问 —— VB6 对 Variant 形参取默认属性恰恰是对的。
- **改（收成一条出口，不在打包处再抄一份判定）**：新增 `CCodeGen::ctrlDefaultPropOf(lower)`（实现 `cgen_util_ctrl.cpp`，声明 `cgen_helpers.inc`）一次回答「折不折 / 折成哪一枚 / 读函数是谁 / 折出来什么型」，**两个消费点**同时改问它：① 标识符发码 `cgen_expr_ident_symbol.inc`（原先那两段各写一份的折叠判定合一，发出来的 C 逐字未变）；② `comPackExpr` 打包档 —— 折叠检查排在 `knownVariantVars_` / `knownObjectVars_` 那串**之前**（实测把它放在那串之后，这一格一条都不变：控件名同时登记在 `knownObjectVars_` 里）。型那一问另开一张与 `getDefaultPropertyName` 逐行对着写的 `defaultPropType(ctrlType)`，**不动** `controlPropType` 那张表 —— 它答的是「显式 `对象.属性`」，而 Text / Caption / Value 挂在默认属性上时从没在里面登记过，硬塞会把 #231 那一格的暴露面一起改动；各行型取自对应 RTL getter 的 C 返回型（`vb6_GetControlText` 回 wchar_t* ⇒ String，`vb6_GetCheckValue` / `vb6_GetOptionValue` / `vb6_GetScrollValue` 回 int ⇒ Integer / Boolean）。打包的类型分支同时收进 `comPackFnForVbType`，`comPackExpr` 末尾那段兜底一字未改。
- **刻意没接管的两档**（记读数，不当判据）：① `Picture`（PictureBox / Image 的默认属性）在新表里答 Unknown ⇒ 仍走今天的 `ComPackObject`（那本来就是个对象）；② `Set <泛对象槽> = <控件名>` 那一形本刀未动，另立 §B87 —— 它才是「该不该折」那一问，与本刀那条轴不是一件事。
- **判据 GC29..GC31**（夹具 `tests/geomcache`，为这一格补了一枚 `chkA As CheckBox`，两架构真跑 31/31 True）：GC29 折出来那枚 BSTR 按 BSTR 装 ⇒ 语句活下来**且**集合里那一枚回读的串就是 `txtA.Text` 自己那句（先 `txtA.Text = ` 塞一句非空串，不让「空 == 空」冒充通过）、`TypeName` 答 String；GC30 同一格的第二个属性名（`lblP.Caption`，读函数同名而控件不同），并钉 `k30 <> k29`（两枚各算各的）；GC31 钉**新表另一行**真的驱动打包（`chkA.Value` ⇒ `vb6_ComPackInt`，读数 `1/[Long]` —— Integer 落 VT_I4 是既有打包口径、不是本账，判据按实读数钉死，将来谁改成 VT_I2 这一格会红）。**证人**：`collC.Add oC`（`For Each` 出来的 Object 变量，不是控件名）必须**照旧**按对象打包 —— 少了这一格，「把实参一律改成标量档」那种修法也能过前三条。
- **负控**（BASE = 本笔父提交 `262d7eab` 冷编那台，同一份夹具）：GC01..GC28 两架构全 True（⇒ 那台带着 #257，归因不会串台），**GC29 那一行根本不出现**、无 `GC-DONE`、EXIT=0xC0000005 —— 崩点即判据；BASE 的发码那三行正是 `vb6_ComPackObject(vb6_GetControlText(...))` / `(vb6_GetCheckValue(...))` 三个形。
- **护栏**：语料 emit A/B（口径同上一轮的宽版：`.vbp` **加** `.bas`，两架构各一次）inputs=391 / captures=782 / **same=780 / changed=2** —— 两条改动都是本夹具（x64、x86 各一次），每次 -=3 / +=3，**逐行归因 unattributed=0**、run-failures=0。这一条同时兜住那两处「零行为改动」的重构：发码面合一与 `comPackFnForVbType` 抽取若动过任何一个字，语料里就会出现归不上号的差行。「语料 0 处」这条老读数本轮再核一次成立（除夹具外没有任何一份工程把控件名的默认属性塞进晚绑定实参）。发码针两头：`gc_emitc_pack_follows_default_prop`（三条新形 + 那条 Object 证人，实测各命中 1，其中证人那条在 BASE 上也命中 1 —— 它是「不该动的东西」）/ `gc_emitc_pack_object_over_value`（三条旧形残留 0，BASE 上各 1）。哨兵跟着长：`check_ctrl_prop_type_authority.ps1` 加 A6（`defaultPropType` 只准 definition + declaration + **恰好一个**调用点，且那个调用点必须在 `ctrlDefaultPropOf` 体内；那条里再出现 `controlPropType` 就红）+ A7（`ctrlDefaultPropOf` 恰好 4 处 = 定义 + 声明 + 两个消费点，少一个或多一个都红）—— 两条都用假改动证过能红（往 `cgen_util_com.cpp` 插一行提及 ⇒ A7 双条红；把旧形状塞回那条 ⇒ A6 双条件红；插完按 md5 原样还原）。29 道 [STATIC] 全 RC=0；run_tests.ps1 PSParser 零错。**门 #396 复核（CI 侧）**：两片 vbp 工件里的 GCCache.out 各 31 行、无一条 False，GC29/GC30/GC31 三行与本地逐字相同（`[GC29TXT]/[String]`、`[GC30CAP]`、`1/[Long]/1`）；两条发码针（形在 / 旧形不在）是 vbp 那两片的用例之一，11 job 全绿即含它们。
- **风险记账**：所有晚绑定调用实参位置上的**裸控件名**从此按标量装箱。答案没变的那两批是 `As Object` 类型化形参（P20-31 抑制默认属性那条）与 With 块（`suppressDefaultProp_` 同一道闸）。新表与旧表从此分工：`controlPropType` 管显式属性读法，`defaultPropType` 只管默认属性那一格 —— A6 钉住这条分界。

### B86 宿主模型没有 `Container`/`Parent` 一档，而窗体自身的 `.Name` 从没登记 —— 排版器那句判据靠「空 = 空」真空通过（账 #257 = C29-GE-g，**已出：门 #394 = run 37615366430、head `18ac1c65`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0（wall 10m59s）**）
- **症状与因果**：`ClsResizer.SaveControlsPositions` 循环第一句 `If oCtrl.Container.Name = oForm.Name Then` 是两个缺陷叠在一起才「通过」的：宿主模型的 getprop 长串里压根没有 `Container`/`Parent` 这一档（`TypeName(o.Container)` 答 **Empty**），而窗体登记走 `vb6_Forms_Register(void* hwnd)`（签名里没有名字，恒传 NULL）⇒ 两边都交空串，「空 = 空」恒真。真流量不止 Charts：`tests/BalloonTooltips/cTT.cls` 的 `objControl.Container.hWnd`（无窗口控件那支）问的也是这一档。
- **改（一批三处，顺序是定死的）**：① `uc_hostmodel_getprop.inc` 新增一档，交回 `(void*)GetParent(接收者)`，走 #250 那条唯一出口 `vb6_ho_setVariantObject`（无父 = 顶层窗体 ⇒ 交回空值）；② `vb6_Forms_Register(hwnd, name)` 加一个 name 形参，发码那一处（唯一调用点 `cgen_form_wndproc_create.inc:45`）交 `formName`；③ `vb6_HostObj_Register` 的「同 hwnd 已登记就整条 return」换成**只补还空着的 name / typeName**（治「后到的带名登记被静默丢掉」那条病，UC 宿主那条路同享；`isForm`/`index` 刻意不动 —— 非数组控件 `.Index` 恒 **-1** 是存量答案，GC22 钉着它）。
- **口径不是新决定**：容器成员资格在 `vb6_uc_collectChildren`（uc_controls.c）里本来就是「GetParent(控件) == 容器」那条关系，这一档问同一条 ⇒ 「oCtrl.Container.Name = X.Name」与「oCtrl 属于 X.Controls」是同一件事，不会两套答案互相矛盾。VB6 在这一层与 Window 一致（PictureBox 有自己的 Controls 集合，实测其内的 Label 不在窗体集合里 = n 少两枚）。另一层的同名答案 `vb6forms_axcontainer.c:195`（OCX 扩展器对象，记录里只有 hwndForm）不是这条链，本刀刻意没动。
- **判据 GC24..GC28**（夹具 `tests/geomcache`，两架构真跑 28/28 True）：GC24「没有第三种答案」+ 范围（`cntOther=0` 且 `cntForm+cntNone=n` 且 `cntForm>=3`，沿用 #254 那条教训 —— 光写恒等式会放过「两边都坏」）；GC25「相等是有内容的相等」（容器名 == 窗体名 == 模块名，且 `Len(nmForm)>0`）；GC26 证人只能来自窗口本身（容器的 `Caption` = 夹具标题、容器的 `hWnd` = `Me.hwnd`）；GC27 **反面**（`picP` 自己的成员属于 `picP`，同一句判据必须 False —— 否则「全都相等」和「空 = 空」在这一格里长得一模一样）；GC28 窗体自己没有容器（顶层无父 ⇒ Empty），拦的是「把接收者原样退回来」那种修法（那种修法 GC25 照样 True）。
- **负控**（BASE = 本笔父提交 `fcae9af6` 冷编那台，同一份夹具）：GC24 raw=0+5+0/5（五枚成员全交回空值 ⇒ 那一档确实不存在）、GC25 raw=[]/[]（就是那个真空通过本身）、GC26 raw=[]/0、GC27 raw=[lblP]/[]/[]，GC01..GC23 与 GC28 两边同 True。**订正自己开工前的一个假设**：原以为窗体上的 Timer（无窗口控件）交回空容器，实测 `Container` 也答窗体（五枚成员 5+0+0）⇒ 空值那一路在夹具里只有窗体自己走过，所以 GC28 在 BASE 上也是 True —— 它是反面证人，不是存在性证人，存在性由 GC24/25/27 三头钉。
- **护栏**：语料 emit A/B（BASE = 父提交那台，全量 136 份 vbp，比历史的「90 份」口径更宽，**口径记下以免两次读数被当成互相矛盾**）INPUTS=136 / CHANGED=61 / added=71 / removed=71，**逐行归因 OTHER=0**（每一条都是 `vb6_Forms_Register` 那一行的换形，一份窗体模块一条，61 份文件共 71 枚窗体）。同一对编译器再把口径放宽一轮（`.vbp` **加** `.bas`，两架构各一次 ⇒ 782 份捕获）：same=660 / changed=122 / added=142 / removed=142 / **unattributed=0** / run-failures=0 —— 142 = 71 枚窗体 × 两架构，两架构的差行同号同形 ⇒ 这一刀没有按架构分家的形状；发码针两头 `gc_emitc_form_name_registered`（带名的整条语句必须在，实测命中 1）/ `gc_emitc_form_name_missing`（旧的那条无名形必须不在，实测残留 0）；29 道 [STATIC] 全 RC=0；run_tests.ps1 PSParser 零错。现场读数（不作判据）：Charts 2020 x86 两台对跑 —— 改后 94ms 出窗、改前 130ms，两边 stderr 都是 0 字节、进程都存活。
- **风险记账**：窗体登记补了名字 ⇒ `oForm.Name` 从空串变模块名，「名字 == 名字」这类比较的两侧同时变真；排版器那句从此真判（同容器为真、异容器为假）。残余缺口另立：**`Form.Parent` 在 MDI 工程里应当答 MDIForm**，本刀按「顶层窗口无父 = 无容器」给的是空值，实测面只有 BalloonTooltips 与 Charts 两处 `.Container`，MDI 语料 0 处；打包那一格已出（§B85 = 账 #256），`Set <泛对象槽> = <控件名>` 那一面换到 §B87（账 #258）—— 两问的轴不同：#256 问「装箱档跟不跟上一步交出的值」，#258 问「该不该折」。

### B87 `Set <泛对象槽> = <控件名>` 交出的是默认属性读数，不是控件本身 —— 「该不该折」这一问当时有两份答案（账 #258 = C29-GE-j，**已出：门 #402 = run 37666864389、head `98347f53`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0（wall ≈10m19s）**）

- **一句话的病**：Set 的右值是**对象引用上下文**（VB6 在这里交出控件本身，默认属性只在值上下文展开：`s = Text1` 取 `.Text` 才是对的），而这一问当时有两份答案 —— 环境闸 `suppressDefaultProp_`（With 块与 `As Object` 形参设它）与 P16 的事后手术（在**已经发好**的 value 文本里 `find("vb6_hwnd_")` 再截一段，只认 `knownWithEventsCtrlVars_` 那一种目标）。两份各盖一半形状，于是 `Dim o As Object: Set o = Text1` 发成 `o = vb6_GetControlText(vb6_hwnd_Text1)  /* default prop: .Text */;  /* Set */`、Variant 槽发成 `vb6_VariantFromValue(vb6_GetCheckValue(...))`。
- **只有半个症状会说话**（这是它躲过前几轮的原因）：那条写落到 `vb6_ComSetProp(<BSTR 指针>, L"Text", …)`，RTL 在 `vb6_getDispid` 那一步认不出接收者，只在 **stderr** 打一条 `property "Text" not found` 就返回 —— `Err.Number` 仍是 0，读回那一路也一样静默交空串。BASE 实测（`aeb883a9` 那台、同一份夹具、两架构逐字相同）：`SO1-RAW err=0 tn=Object own=alpha read=`（写没落地、读回空）、`SO3-RAW err=0 vt=3 tn=Long`（VB6 应 `vt=9`）。⇒ 「不崩」不等于「对」；这一问与 §B85（账 #256）不同轴 —— 那格问「装箱档跟不跟上一步交出的值」，本格问「该不该折」。
- **改（收成一处）**：新增唯一出口 `CCodeGen::emitSetObjectRhs(value)`（定义 `cgen_setlet.cpp`、声明 `cgen_helpers.inc`）—— `visit(SetStmt&)` 的**五个**右值发码点（MyBase 属性写 / 控件属性原生 setter / 跨类 `prop_set_` / COM SetRef / 主路）全部转调它；闸只在右值**整枚是一枚标识符**时开，所以子树（`Set o = f(Text1)`）照旧按值上下文，不会顺着漏进嵌套实参位。同时新增第二件唯一出口 `ctrlObjectRefExpr(lower, cName)`：「把这枚控件当**对象**交出去发什么 C」与折叠那支的接收者问同一条 `makeCtrlHwndArg`（WithEvents 控件变量那一档交回变量自己），撤掉 `cgen_expr_ident_symbol.inc` 里 suppress 那一支的 `"vb6_hwnd_" + cIdent(node.name)` 手拼 —— 那一拼对 ListView 槽变量、ImageList 的 `vb6_com_<名>`、WithEvents 变量三档都会发出 C 里根本不存在的名字。P16 那截字符串手术随之撤掉（typed 控件变量那一档的发码**逐字不变**，见护栏）。
- **判据 = 新夹具 `tests/setobj`**（六头，两架构真跑 6/6 True）：SO1 存在性 = 经槽写进窗口再经槽读回（`own=beta read=beta`，起点是 `alpha` ⇒ 「空=空」蒙不过去）；SO2 存在性 = 槽是**活引用**不是快照（改窗口后再读槽 `got=gamma`）；SO3 存在性第二档 = Variant 槽 `vt=9 tn=CheckBox`；SO4/SO5/SO6 是证人（SO4 = 同一名在**值**上下文照旧折 ⇒ 拦「一律不折」那种修法；SO5 = typed 控件变量那一档在 P16 撤后仍交回 `vb6_hwnd_Command1`；SO6 = 裸 HWND 交给 `As Object` 形参后晚绑定属性写**真落窗** ⇒ 担保修法所用的那个表示是现成机制认得的，判据不是自证的）。负控 = 本笔父提交 `aeb883a9` 冷编那台跑同一份夹具：SO1/SO2/SO3 两架构全 False（raw 就是上面那两条读数），SO4/SO5/SO6 照旧 True、`SO-DONE` 齐。
- **发码针两头**：`so_emitc_set_rhs_is_object`（三条新形各命中 1；其中 `cmdW = vb6_hwnd_Command1;  /* Set */` 在 BASE 上也命中 1 —— 它是「不该动的东西」那一类证人，不当存在性证据）/ `so_emitc_default_prop_in_set`（BASE 原地发出的那两条旧形残留 0）。
- **护栏** = 语料 emit A/B（`.vbp` **加** `.bas` × 两架构，口径同 #256/#257 那两轮）inputs=394 / captures=788 / **same=786 / changed=2 / 未归因 0 / run-failures=0** —— 两份改动正是本夹具的 x64 与 x86 各一次。这一条同时兜住本刀那两处「零行为改动」（`ctrlObjectRefExpr` 顶掉手拼、P16 撤掉）：任一改动漏了字，语料里就会出现归不上号的差行 —— 而 typed 控件变量那一档在语料里**有实物**（`tests/test_form/form_test_withevent.frm:44/45`，全仓 544 条 `Set` 里仅有的两条裸标识符右值），正是这条护栏的落点。哨兵 = 新第 32 道 [STATIC] `scripts/check_set_rhs_object_context.ps1`：S1 站点恰好 7（声明 + 定义 + 五个写侧出口）、S2 `cgen_setlet_set_prop.inc` 里 `emitExpr(*node.value)` 必须 0 处、S3 P16 那截 `value.find("vb6_hwnd_")` 不许回来、S4 `ctrlObjectRefExpr(` 恰好 3 处、S5 那一支不许手拼且必须问权威恰好 1 次、S6 闸的「整枚标识符」那一行必须在（少了它闸就会顺着子树开）。四条假改动各证过一次能红（撤一个站点 ⇒ S1+S2 双条；把手术插回 ⇒ S3；把权威换回手拼 ⇒ S4+S5 双条；把闸写成恒开 ⇒ S6），每次按 md5 原样还原；32 道全 RC=0；`run_tests.ps1` PSParser 零错。
- **刻意没接管的两档**（记读数，不当判据）：① `Set o = <控件数组名>` 那一形两架构都发 `vb6_hwnd_<数组基名>` ⇒ C2065 未声明，**改前改后同形**（元素句柄住在 `vb6_arr_<名>` 里，那是 #189/#191 那一族；VB6 自己也不接受这一形）；② `cgen_with.cpp:130/141` 那两处手拼问的是 With 登记表里的 `ctrlOrigName`，不是这条标识符出口，本刀没动。
- **暴露面**：语料 0 处坏形状（除上面那两条已通的证人形）⇒ 纸面缺口，与 #256 同族但错的是值不是命。
- **风险记账**：泛对象槽与 Variant 槽里的裸控件名从此存 HWND（以前存默认属性读数）⇒ `TypeName` 从「Object」那句谎变成真控件名、`VarType` 从 3/8 变 9，而**经槽的属性写从此真的落到窗口上**（以前静默不响）。语料 0 处 ⇒ 门上的暴露面只有本夹具；若门在别处红，先按这条轴归因（是谁开始依赖那句谎）。另记：§B73（CI 与本机发码差）与这刀不同轴（那条是发码**路**选错，这条是发码**值**选错），下一轮按 #401/#402 的工件重数那 14 条。
- **门 #402 复核（CI 侧）**：`SetObj.out` 两片各 13 行、无一条 False，六行判据与本地逐字相同（`SO1-RAW err=0 tn=TextBox own=beta read=beta` / `SO2-RAW got=gamma own=gamma` / `SO3-RAW err=0 vt=9 tn=CheckBox` / `SO4-RAW s=gamma` / `SO5-RAW err=0 cap=cmd` / `SO6-RAW err=0 got=delta own=delta` + `SO-DONE`）；两片 `SetObj.err` 都是 **0 字节**（BASE 上那两条 `vb6_ComSetProp: property "Text" not found` 从此不再出现 —— 记作现场读数，判据住在 SO1..SO3 三头存在性上）。两条发码针（新形在 / BASE 旧形不在）是 vbp 那两片的用例之一，11 job 全绿即含它们。
- **与门 #401（head `aeb883a9`）的夹具输出对跑（逐份归因）**：123 份 `.out` = **105 份逐字节相同 + 16 份有差 + 2 份新增**（`SetObj` 两架构）。16 份差的**每一条**都落在句柄/DC/窗位或拍号/端口/事件计数那两族形状上（`GCCache` 的 GC16 `L<窗位>` 与 GC20/GC23/GC26 的 `raw=` 尾数、`VeUnits` 的 `U-CNT-RAW cnt=`、`PcDraw` 的 `PL00-DC`、`FrmEvents` 的 `EV24 … X=`、`ModalApp`/`TabWalkApp` 的句柄与 `busy=`、`TmApp` 的拍数、`WsApp` 的端口与 `W1`/`W4`/`W9`、`RtfApp` 的字符位），而**所有 `=True` / `=Y` / `eq=Y` 那些判据 token 两侧同号** ⇒ 这一刀在门上没留下任何判据面的动静。**顺带订正那份「天生会抖」名单**：`GCCache` / `PcDraw` / `VeUnits` 从这轮起也在册 —— 它们的读数行里印的就是窗口句柄与 DC，本来就该抖（下次别把这四份当回归信号）。

### B88 内层泵把那条 WM_QUIT 抄走了 —— 最后一个窗体卸载之后进程不退出（账 #259 = 提交 `990cb883` + 台账 `473c6dfc`，**已出：门 #397 = run 37642453080、head `473c6dfc`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0（wall ≈14m14s）**）

- **门上的样子**：#395（head `262d7eab` = 纯文档笔，代码与全绿的 #394 逐字相同）唯一红 `[VBP] combofocus_x86 … run timeout: 60s (cpu=62ms, alive) last='COMBOFOCUS-DONE' lines=7` —— 七行判据全打完了、窗口一个不剩（`MainWindowTitle` 空）、进程不退。**订正 #256 那一轮写下的话**：当时记的是「本地两台 EXIT=0 ⇒ 不是产品确定性缺陷」，这句错了 —— 本地加负载就把同一条挂复现出来。
- **机制（三处临时标记量出来的，探针已撤、RTL 回到提交前逐字）**：挂住那几轮读到 `loop-enter` → `unreg count=0 depth=1`（PostQuitMessage 确实发了）→ `doevents-ate-quit` → **没有 loop-exit**。一个线程只有一份 quit：VB6 的「最后一个窗体卸载 ⇒ 进程结束」在这仓里就是 `vb6_Forms_Unregister` 投的那条 WM_QUIT，而 `vb6_DoEvents` 的 `PeekMessage(PM_REMOVE)` 把它从队列里摘走、当普通消息 Dispatch 掉，外层 `GetMessage` 就永远等不到。饿机器上这件事的形状是：tick 6 的 DoEvents 把已经排队的 tick 7 抽出来跑，而 tick 7 正是卸载那一个 ⇒ quit 投在同一个 DoEvents 里、被同一个 DoEvents 吃掉。本地读数：5 枚并发烧 CPU 时 combofocus_x86 **13/14 挂**，不饿时 1/20；这就是「CI 才红、本地绿」的全部原因。
- **改（一处助手 + 两个内层泵各接一句）**：`vb6_RePostQuitIfTaken(&msg)` —— 抽到 quit 就原样投回去并停泵。装的两处 = `vb6_DoEvents` 循环体第一道闸，与 `vb6_ShowForm` 那条模态循环的收尾（模态那条同形：它的 `GetMessage` 也会抄走 quit，且循环体里原本那句 "收到 WM_QUIT, 模态循环结束" 的 trace 是**死码** —— GetMessage 取到 quit 返回 0，循环直接结束，从来不会把它送进循环体）。外层 `vb6_MessageLoop` 刻意**不**投：它是消费者，给它加投回等于把退出信号发两遍。`vb6forms_ctrlarr.c` 里那条 `vb6_MDIMessageLoop` 是 MDI 工程的外层泵，同属消费者，没动。
- **判据 = 新夹具 `tests/appqquit`**（把「卸载发生在内层泵里」做硬，不再靠饿机器碰：每拍 `Sleep 80` 而 Interval=40 ⇒ `DoEvents` 起手时那条 WM_TIMER 必已在队列里 —— **⚠ 这句只到门 #402 为止**：它靠的是同一枚 Timer 在自己泵里重进，而那条可重入本身是缺陷；形状已在门 #403 之后挪到第二枚计时器，见本节最后一条）。两行读数各钉一件事：`Q-DONE` 说判据都写完了，`Q-ORDER=unload-inside-doevents` 说**形状真到了那一格**（少了这一格，一个从没进过洞的编译也照样绿）；本体判据是「进程自己退」—— run_tests.ps1 对 run timeout 直接判 FAIL，所以不需要在夹具里加「再等一拍」那种断言（#213 那条教训）。读数：改后两架构 0/10 挂、EXIT=0x0 两行齐；负控 = 父提交 `262d7eab` 那台编同一份夹具 ⇒ **5/5 挂**（两行齐、无窗口、不退）；原夹具 combofocus_x86 从 13/14 挂变 0/15 挂。
- **护栏**：新哨兵 `scripts/check_quit_pump_invariant.ps1` —— P1 泵数（vb6forms.c 里 GetMessage 三条 + PeekMessage 一条，数目变了必须说清哪条是外层）、P2 `vb6_RePostQuitIfTaken` 恰好三处（定义 + 两个内层泵）、P3 DoEvents 那道闸在不在、P4 模态循环收尾补投在不在、P5 全仓 `PostQuitMessage` 只剩两处（登记处投 + 助手补投）且**外层泵里没有** hand-back。规则能用假改动证红：删掉 DoEvents 那一闸 ⇒ P2 + P3 双条红（插回按 md5 原样还原）。30 道 [STATIC] 全 RC=0；run_tests.ps1 与本哨兵 PSParser 零错。本刀纯 RTL，`src/backend` 一字未动，所以发码面按定义不变（那条「只改 RTL 时 --emit-c 零差异」的老口径，判据落在真跑与哨兵上，不落在 A/B 上）。 **门 #397 复核（CI 侧，按超时形状查）**：四片 vbp 工件共 115 份夹具输出，**每一枚 `.out` 都有配对的 `.err`**（缺 `.err` = 那枚是被 60s 超时杀掉的，见 §B44/#359 那条分辨法）⇒ 一片都没有被杀；`CbApp.out` 两架构各 7 行齐、`QApp.out` 两架构各 2 行（`Q-DONE` + `Q-ORDER=unload-inside-doevents`）都正常退出。
- **风险记账**：DoEvents 从此在抽到 quit 时立刻停泵并返回 —— 与 VB6 的「DoEvents 交出控制权」不冲突（那条 quit 只在最后一个窗体卸载时才存在，而那本来就是进程要结束的信号）。模态 Show 若在 quit 之后返回，调用方代码会继续跑一小段再落到外层循环 —— 与 VB6 同形（VB6 也是把退出交给消息循环，不由 Show 直接终结进程）。另记一条**没查**的同类面：`vb6_DoEvents` 的 `count > 1000` 那道安全闸与 quit 停泵共用一个返回口径（返回的处理条数），VB6 的 DoEvents 返回值在本仓不是判据，未钉。
- **夹具形状已随 Timer 不可重入改掉（门 #405 的唯一真红，2026-10-08）**：#405（run 37687393362、head `c62fcf61`）两片红，逐份对工件比下来只有**一条判据线换了答案** —— `QApp.out` 的 `Q-ORDER=unload-inside-doevents` 变 `unload-outside`（其余 GCCache/PcDraw/VeUnits/RtfApp/TmApp/ModalApp/TabWalkApp/WsApp 全落在 §B44 那份在册抖动名单的两族形状上，且第三片带着同样三条差却全绿）。归因**不在 #261**：合进来的 `90fb8069` 把 Timer 事件过程改成不可重入（`vb6_DispatchTimer` 给每格加 `inCallback`，同一枚的下一拍丢弃），而这份夹具的"确定性"恰恰是靠**同一枚 t1 在 tick 6 的泵里重进**做出来的 —— tick 7 被那道守卫丢掉了，卸载于是落到泵外。**守卫是对的（VB6 本人就不可重入），过时的是夹具形状**；本机在同一台编译器上按旧夹具读出 `unload-outside`、EXIT=0x0（负控方向已实测，不是推测）。形状挪到**第二枚计时器**：t1（40）每拍 Sleep 80 再 DoEvents，那 80ms 里 t2（20）**必定已到期**（Sleep 保证钟差 ≥ 4 个周期，WM_TIMER 在 DoEvents 起手那次 PeekMessage 就生成），卸载由 t2 干 ⇒ 洞一点没少走：卸载仍在内层泵里、那条唯一的 WM_QUIT 仍被同一个内层泵抽走、仍靠 `vb6_RePostQuitIfTaken` 投回，而「进程自己退」这条本体判据照旧。`gLetGo` 只在 tick 6 递出，所以 t2 不可能提前卸载、也不可能把形状记到没参与的那条路上。读数：x64 三连 + x86 三连 ⇒ **6/6 `unload-inside-doevents` + `Q-DONE` + EXIT=0x0**，两枚针面一字未改（改的是夹具，不是判据）。顺带把 `run_tests.ps1` 里那段描述旧形状的话与 §B88 上面那条括号一起订正了 —— 台账与注释里的机制句都是待测主张。

### B89 `With <对象> : .Prop = <别的对象>.Prop` —— 右值那一整步成员读取被丢掉，交出去的是「那个对象本身」；漏下的标记还会劫持下一条语句（账 #260 = C29-GE-i，**已出：门 #401 = run 37660544554、head `aeb883a9`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0（wall ≈10m57s）**）

- **一句话的病**：`cgen_assign_prop_write.inc` 的 `WithObjKind::COMObject` 那一支是唯一**没消费 COM 标记**的写侧出口 —— 右值整枚就是一枚 `对象.成员` 读取时，成员名还挂在 `comObjExpr_`/`comMemberName_`/`isComMarker_` 上没人取，于是发出 `vb6_ComSetProp(_vb6_with_N, L"Prop", vb6_ComPackValue(<别的对象>))`：把那个对象塞进属性槽。不响、不崩、值永远错。
- **为什么只有 With 这一形掉**：那张「packer → 解封类型」的表被抄成了**三份**（`cgen_assign_com_prop.inc` 两处 = Fix 110i 与 SetPropArg、`cgen_util_comwrite.cpp` 一处 = P25b、`cgen_assign_host_pseudo.inc` 一处），三份都有，唯独 With 那一支一份也没问 ⇒ 同一句写在 With 外面一直是对的。五形实测（`D:/c3.vb6.pro/.build/p245ref/P245e.cls`）：字面量 ✓、数值 ✓、`X.M + 1` ✓、`X.M & "x"` ✓、**`X.M` ✗**；三种接收者（类字段 / 局部 / 形参）都掉，工程内**类实例**的 With 不掉（走 prop_let 直调，不经这条路）。
- **第二个症状（A/B 量到的，读码时没料到）**：漏下的标记会被**下一条语句**的 COM 调用路消费 —— Charts 六份 UC 工程里 `PropertyChanged "TitleFont"` 这类调用被改写成 `vb6_ComVarFree((void*)vb6_ComCall((*New_Font), L"Charset", (void*[]){…}, 1))`：32 条「往字体对象乱写」＋ 32 条本该发出的 PropertyChanged 整条不见。**"只接了一半"这一族的第二格症状是"漏下的状态还会咬邻居"**，所以这一族的护栏必须是站点计数（哨兵 C1），不能只钉被修的那一行。
- **改（收成一处）**：新增 `CCodeGen::comMarkerValueForWrite(packFn, valExpr)`（`cgen_util_com.cpp`，紧挨 `resolveComMarkerForPack`）—— 无标记原样返回，有标记就按 packer 反推解封类型、问唯一出口 `resolveComValue`、交回解开的值表达式；那张 hint 表**只许住在这个函数里**。三处旧抄本撤掉、With 那一支补上 ⇒ 写侧共五个消费点。顺带两处口径对齐：host_pseudo 那份的默认档是 Variant（Boolean 那一档会发出 `vb6_ComPackBool(vb6_VARIANT)`），现在跟表走 Long，与 `resolveComMarkerForPack` 对 Boolean 的答案一致；`resolveComMarkerForPack` 里 Bool/Int 两条合并到同一答案。
- **判据** = 新夹具 `tests/comwith`（一枚 .frm，`Scripting.Dictionary` 晚绑定，**不依赖类型库** ⇒ 与本机注册表 / `Reference=` 相对路径无关，见 §B73 那一坑）：`CW-WITH`/`CW-LOCAL`/`CW-PARAM` 三种接收者各钉一条（存在性，每格问 `err` 与读回的数）、`CW-EXPR` 钉「算术形不得二次解封」、`CW-DIRECT`（With 外同句）与 `CW-LIT`（With 内字面量）是反面证人。两架构真跑 6/6 True。**BASE（父提交 `2b86a45f` 冷编那台）跑同一份夹具** ⇒ `CW-WITH-RAW err=0 cm=0`、`CW-LOCAL` 同、`CW-PARAM-RAW err=13`、`CW-EXPR`/`CW-DIRECT` 都 `cm=13`（第二个症状：BASE 把 `e1 = Err.Number` 发成「把 Err.Number 再写进挂着的 `d1.CompareMode`」⇒ d1 变 13）⇒ **CW-DIRECT 在两台不同数，它是被这一刀连带修好的，不能当存在性证人**；两台都 True 的只有 `CW-LIT`。
- **发码针两头**：`cw_emitc_with_member_read_resolved`（三条新形各命中 1，含 With 外那条证人形）/ `cw_emitc_with_object_over_value`（BASE 的三条旧形残留 0）。
- **护栏** = 语料 emit A/B（`.vbp` + `.bas` × 两架构，captures=786）**same=772 / changed=14 / 未归因 0 / run-failures=0**；14 份改动 = 本夹具 2 + Charts 六份工程 × 两架构 12，逐行按形对上：With-COM 读解开 266 / 假写消失 266、劫持形消失 32 / PropertyChanged 回来 32。这一格与纯 RTL 刀相反 —— 发码**必须**变，A/B 在这里既是归因面也是"改动半径 = 预期半径"的证据。
- **哨兵** = `check_com_marker_write_sites.ps1`（第 31 道 [STATIC]）：C1 站点恰好 7（声明 + 定义 + 五个写侧出口）、C2 那张 hint 表六档在 src/backend 里各恰好一次、C3 With 那一支的消费必须排在发 `/* With COM SetProp */` 那行之前。三处假改动各证过一次能红（撤消费者 ⇒ C1+C3 双条红；把表再抄一份 ⇒ C2 红；把消费挪到发行之后 ⇒ C3 红），每次按 md5 原样还原；31 道 [STATIC] 全 RC=0；`run_tests.ps1` PSParser 零错。
- **风险记账**：Charts 那六份工程是门上跑的 GUI 夹具（`Test-GuiVbp`，部分带 `DumpMinColors`）。这一刀让**以前静默失效的属性写真的落到 IDispatch 上**，还会开始发 `UserControl_PropertyChanged` ⇒ 字体档从此生效、UC 的属性包从此被标脏。若门红在 Charts 的像素/计数读数上，先按这个方向归因（是"以前没画上"变"画上"），别顺手把判据调宽。
- **门 #401 复核（CI 侧，逐份归因）**：四片 vbp 工件共 121 份 `.out`，**每一枚都有配对的 `.err`**（缺 `.err` = 那枚被 60s 超时杀掉，§B44 那条分辨法）⇒ 一片没被杀；两片各 13 行的 `WithCopy.out` 与本地逐字相同（六格 `err=0 cm=1` + 六个 True + `CW-DONE`，x64 与 x86 同号）。**与门 #400（head `d199adc0`）的夹具输出对跑**：104 份逐字节相同、15 份有差、2 份是本轮新增（`WithCopy` 两架构）。15 份差的**每一条**都落在在册抖动名单那两族形状上（句柄/DC 编号：VeUnits `U-CNT-RAW cnt=`、GC20/GC23/GC26 的 `raw=` 尾数、PcDraw `PL00-DC`、ModalApp/TabWalkApp 的句柄；拍号与事件计数：TmApp `T5`/`T11`、ModalApp `busy=`、WsApp 端口与 `W4`、RtfApp 的字符位与事件计数）⇒ **没有一条判据线换过答案**（所有 `=True` / `=Y` 行两侧同号）。这一格把上面那条风险记账直接结掉：#260 在门上没有留下任何像素/计数面的动静。
- **§B73（CI 与本机发码差）下一轮再量**：本轮 A/B 那 14 份改动全部归得上号 ⇒ 那批 CI-vs-本机差异里凡是本账形状的都该随 #260 一起消失；剩下的按新的门工件重数，口径不变（先取工件、按行多重集比）。

### B90 UDT 成员数组的第 2..N 维在 parser 就被丢掉 ⇒ `M(3, 3)` 发成 `int32_t M[4]`，三个不同下标读写同一格（账 #262 = C29-DM-a，**已出：门 #409 = run 37712639451、head `b334858c`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0**；同批带走 §B91 那格 #263，并把 #261 留的边界一起改了口径）

- **一句话的病**：`src/parser/parser_decl.cpp` 的 UDT 成员那一支遇到 `M(a, b, …)` 时把后续维**解析并丢弃**（原注释就写着「解析并丢弃后续维度」），`TypeMember` 从头到尾只有一个 `arraySize`。于是 `M(3, 3) As Long` 在结构体里只有第一维（`int32_t M[4]`，16 字节而不是 64），而两条下标路（`obj.M(i, j)` 与 `With x : .M(i, j)`）都只折一层括号 ⇒ `m.M(0, 1)` / `.M(0, 2)` / `.M(0, 3)` **三句全发成 `m.M[0] = …`**。不响、不崩、静默错值。
- **读数（探针 `.build/p248b/mem2d.bas`、`.build/p262/m262b.bas`，两架构）**：同三句写在过程内的 `Dim one(3, 3) As Long` 是通的（`AL-PLAIN a=11 b=22 c=33`），写在 UDT 成员上是 `a=33 b=33 c=33` ⇒ **分家的轴是「UDT 成员 × 多维」**，不是多维本身（#211 那条多维 ND 出口只管过程内数组），也不是成员本身（一维成员数组一直是通的，#261 的 F2L20/29/30 就钉在它上面）。三维探针同一形状：`T(1, 2, 3)` 发成 `int32_t T[2]`，`T(0,0,0)` 与 `T(0,0,3)` 同一格。
- **真流量**：`tests/Charts 2020/LabelPlus/LabelPlus.ctl:154-156` 的 `Public Type COLORMATRIX : M(0 To 4, 0 To 4) As Single` —— GDI+ 的 ColorMatrix 要 5×5=100 字节连片，我们给出 20 字节；`:1085/:1086/:1145/:1159` 那几条 `.M(3, 3)` / `.M(4, 4)`（阴影透明度与图片透明度）互相覆盖。语料普查（脚本 `.build/b264_dims.py`）：**多维 UDT 成员声明全仓 1 处，而它的 38 处 `.M(` 用法全部给两个实参（arity=2 = 38/38）** ⇒ 修完不需要照顾"少给一维"的兼容路径，那路径今天也没人走。
- **改（一档一源 + 一条出口）**：① AST `TypeMember` 加 `moreDims`（第 2..N 维的 lower/upper），`arraySize` 仍恒是第一维上界 ⇒ 现有消费者一律照旧，多出来的维只在这一个字段里；`arrayRank()` 就地由 `arraySize + moreDims.size()` 算，不留第二份真相。② 克隆那趟（`ast_clone.cpp`）一起带上 —— 泛型实例化会 `cloneTypeDecl`，漏了就把这一格在 `Of T` 里悄悄还原回一维。③ 语义层（`semantic_analyzer_decl_type.cpp`）**只传秩**：`mi.arrayRank = 1 + moreDims.size()`（纯语法计数，不折常量 —— 各维格数住在发码侧，抄一份到这里就会和 `cgen_decl` 分家）。④ `cgen_decl.cpp` 每一维过**同一条** `tryEvalConstInt` 折成 `T name[n0][n1][…]`。⑤ 下标新增唯一出口 `CCodeGen::udtFixedMemberIndex(fieldExpr, args, rank)`：层数只从秩来、每层一个方括号、**步长归 C 算**（所以没有"每行几格"那张表）；缺的那层补 `[0]`，多给的实参照发 ⇒ C 当场报错，不再静默少折一维。两条下标路都改成问它。
- **顺手收的两条 rank-blind（同一族，A/B 里逐字节不变）**：`cgen_expr_call_arg_emit.inc` 那处把定长成员数组包成 `vb6_SafeArray1D` 交给 ByRef 形参时，格数与元素尺寸都写死 `sizeof((arg)[0])` —— 秩 1 正好，秩 ≥2 数的是**一行**。改成按秩把 `[0]` 剥到底（`lay44` 问 `mi.arrayRank` 一次），秩 1 的产物逐字不变。⚠ 载体仍是一维描述符：「多维成员整体交给 ByRef 数组形参」这一形语料 0 处、**没实测**，这里只保证数出来的格数与字节对，别读成"多维 ByRef 已支持"。
- **同一批把 #261 留的边界收了**（那是 §B77② 记下的"§B90 落地那天一起改口径"三头之一）：`narrowTargetTypeOf` 的数组那一档以前只接 `callee 是裸标识符`，现在 callee 是成员访问 / With 成员时也问那条唯一的字段出口，**并且要求问出数组性**（认不出是数组就答 Unknown —— 整体赋值那条路早在空下标那道闸就退了）。于是 `p.Pixels(0) = 6.73` 从截断 6 变成 round 7 ⇒ `test_f2lng` 的 **F2L28 翻成 `28-member-array-elem-rounded`**，新增 **F2L31/32**（多维成员的两个相邻格各自取整、互不覆盖），发码针两头对调（裸形 `p.Pixels[0] = 6.73…` / `p.Grid[0] = …` / `uint8_t Grid[2];` 从"必须在"改成"必须不在"）。F2L29/30（成员数组的**整体**赋值保持裸形）一字未动 —— 那道闸守的还是 VbQRCodegen 那枚启动期 error 6。
- **判据** = 新夹具 `tests/test_memdim.bas`：每条读数写的都是**第一维相同、第二维不同**的格子（正是旧折法会并成一格的那一族）—— `MD-LONG=11/22/33`、`MD-SINGLE=111/222`、`MD-R3=1/3/40`、`MD-WITH=11/77`、`MD-MODULE=44/55`；`MD-LEN=64/64/16` 钉**布局**（结构体与字段的字节数，不是读数）；`MD-ONE=9` 是"一维成员不许被这刀带坏"的反面证人；`MD-STR` / `MD-STR-COPY` 钉 #263；`MD-ALIAS=N` 是不变式（三个格子两两不同 —— 只看具体数字会放过"全并成一格但恰好写同一个值"那种假绿）。两架构真跑 **11/11 行齐、EXIT=0x0、两边逐字相同**。
- **负控（两台、两份源）**：BASE = 本笔父提交 `c4175e64` 冷编那台（`.build/C3_base262.exe`）。① 同一份夹具 ⇒ **BUILD-RC=1，一条 `error C2065: "_i"`**（#263 挡住，量不到 #262 的读数 —— 这本身就是"两格缺陷叠在同一条走查上"的实物）；② 把 String 成员摘掉的同一份源（`.build/p262/m262b.bas`）⇒ BASE 编得过，答 `33/33/33`、`222/222`、`3/3`、`LEN=16/16`，改后答 `11/22/33`、`111/222`、`1/3/40`、`64/64` ⇒ **两格各自的读数在负控里分得开，没有互相掩盖**。
- **护栏** = 发码针 `memdim_layout`（12 条必须在：四枚结构体的多维声明 + 三条下标路 + 拷贝助手的扁平指针行；8 条必须不在：`int32_t M[4];` 那种少一维的声明、`m.M[0] = 11;` 那种少一层括号的下标、`sizeof(d->S[_i])` / `d->S[_i]` 那种 #263 形状）+ 新哨兵 `scripts/check_udt_member_dims.ps1`（第 33 道 [STATIC]，U1..U8：**下标折算一个出口两条路**、层数只从秩来、维数只在 parser 数、语义层只传秩、每维过同一条折叠出口、拷贝助手走扁平指针、元素档那道闸必须问数组性、ByRef 那处按秩剥到底），**八处假改动各处证过能红**（`.build/b262_negctl.py`，每次按字节原样还原并复跑确认绿）。33 道 [STATIC] 全 RC=0；`run_tests.ps1` 与本哨兵 PSParser 零错。
- **语料 emit A/B**（对本笔父提交那台 `.build/C3_base262.exe`，judge = 先把所有 `vb6_Chk*` / `vb6_FltToLng` 配对拆干净，再按**行多重集**比，剩下的差异必须逐行归到形上）⇒ inputs=395 / captures=790 / 逐字节相同=774 / 改动=16 / **未归因 0** / run-failures=0。16 份分两族：**维度族 6 份** = `tests/Charts 2020/Proyecto1.vbp` × 两架构（112 条下标各多出一层括号、12 条成员声明多出一维）+ 两份新夹具；**取整族 10 份** = #261 那条边界收掉之后多套的检查（VBFlexGridDemo 两份、dbgdlg 两份、erase_sub 四份、Common.bas 两份），拆掉包裹后与 BASE 逐字节相同。**两族的行数自己对上**：`subscript missing a dimension (BASE) 112` ⇄ `subscript now carries every dimension 112`、`member declaration with ONE dimension 12` ⇄ `gained a dimension 12`、拷贝助手 `6 ⇄ 8`。**A/B 还抓到我自己写进去的一格非法 C**：ByRef 那处第一版写成 `(int32_t)sizeof(tLogFont.lfFaceName)[0]`（= 给 `sizeof` 的结果取下标），秩 1 本该逐字不变 ⇒ dbgdlg 那 4 行差异就是它；`--emit-c` 不调 cl，这类东西只能靠逐行多重集比抓（真编译要等门，代价是一整轮）。
- **风险记账**：Charts 那枚色彩矩阵从此**真的按 5×5=100 字节交给 GDI+**，`.M(3, 3)`（阴影透明度）与 `.M(4, 4)`（图片透明度）不再互相覆盖 ⇒ 门若红在 Charts 的颜色/不透明度读数上，先按"以前静默错值、现在按 VB6 的意思走"这个方向归因（#260 那格同形），别调宽判据；dbgdlg 的 `tLogFont.lfFaceName[i] = vb6_Asc(...)` 从此套上 `vb6_ChkByte`，`Asc` 的值域本在 0..255 内 ⇒ 不会因此报 6，但这一族从此**多一道越界闸**（与 #261 同一口径）。

### B91 含所有权元素的定长数组成员：自动深拷贝助手在**声明 `_i` 的那一句里**就用 `_i` ⇒ C2065，工程编不过（账 #263，**已随 #262 同批出：门 #409**）

- **这条不是找出来的，是量 #262 时撞出来的**：给 #262 写探针时加了一枚 `S(1, 2) As String` 的 UDT，`--emit-c` 出来能编 —— 真编译那台**直接 BUILD-RC=1**：`m265.c(12): error C2065: "_i": 未声明的标识符`，第二枚同类 UDT 再来一条（`m265.c(22)`）。所以本格的第一个读数不是"值错"而是"编不过"，而它一直藏在**语料里没有这种形状**这件事后面。
- **触发条件（三条同时满足，缺一条就不响）**：① UDT 成员是**定长数组**（`X(N) As …` / `X(a To b) As …`）；② 元素档**带所有权**（`String` / `Variant` / 含所有权成员的子 UDT —— 不然走 `memcpy` 快路，那条一直是对的）；③ 这枚 UDT 会被**整体赋值 / 按值传参 / 返回**（ ⇒ Fix 178 那套自动拷贝助手被生成）。助手发的循环头写的是 `{ int32_t _n = (int32_t)(sizeof(d->S) / sizeof(d->S[_i])); int32_t _i;` —— `_i` 出现在它自己那条声明**之前**，C 当场拒。语料普查：0 处（这也是为什么它比 #262 更早存在却从没红过）。
- **改（与 #262 同一处代码、同一个函数体）**：逐格走**扁平元素指针** —— `{ ET* _dp = (ET*)&d->X; const ET* _sp = (const ET*)&s->X; int32_t _i; int32_t _n = (int32_t)(sizeof(d->X) / sizeof(*_dp)); for (…)`，`ET` 由那一档本来就分好的三路（String ⇒ `BSTR` / Variant ⇒ `vb6_VARIANT` / 子 UDT ⇒ `vb6_type_X`）给出，认不出元素档就**不动这一格**（上面那条 memcpy 快路已经走过了位拷）。这一改同时治了第二条更晚的病：**秩**——旧的 `sizeof(d->S)/sizeof(d->S[_i])` 就算能编，在 `BSTR S[2][3]` 上数的是**行数**（2 不是 6），会少拷四格。
- **判据**：`tests/test_memdim.bas` 的 `MD-STR=ab/cd` 与 `MD-STR-COPY=ab/cd`（同一枚 2-D String 成员，赋值前/后各读一次 —— 只有赋值后那条才证明助手真的逐格 `SysAllocString` 过）；发码针 `memdim_layout` 两头：`int32_t _i; int32_t _n = (int32_t)(sizeof(d->S) / sizeof(*_dp));` 必须在，`sizeof(d->S[_i])` 与 `d->S[_i]` 必须不在；哨兵 `check_udt_member_dims.ps1` 的 U6（把 `sizeof(*_dp)` 换回 `sizeof(d->" + idx + ")` ⇒ U6 红，N6 实测过，按 md5 逐字节还原）。
- **负控**：本笔父提交那台（`.build/C3_base262.exe`）编同一份夹具 ⇒ **BUILD-RC=1、1 条 C2065(`_i`)**；把 String 成员摘掉的同一份源（探针 `.build/p262/m262b.bas`）两台都编得过，BASE 答的是 #262 那一族（`33/33/33`、`222/222`、`3/3`、`LEN=16/16`），改后 `11/22/33`、`111/222`、`1/3/40`、`64/64` ⇒ **两条账各自的读数在负控里分得开**，没有互相掩盖。
- **风险记账**：语料 0 处 ⇒ A/B 里这一族只该出现在新夹具上；门上若出现"以前编不过的 UC 现在编得过"，先按这条归因。旁边一格**没量**：元素档是 `As <项目类>` / `As Object` 的定长成员数组走的是**上面那条 memcpy 快路**（`scalarOwned` 只认 String / Variant），也就是"两个 UDT 共享同一批对象引用、其中一边 Clear 就悬"——这一刀**没动**它（BASE 也那样），它到底是第几格缺陷要先数语料里有没有这种成员，别把这条读成"已修"。

### B92 门红的时候拿不到「是哪一枚夹具红」—— 失败注解里只有 exit code 1（2026-10-08，**已出：门 #421 = run 37758983155、head `5fae2205`、branch dev、attempt 1 = 12 条 check-run 全 completed/success、非绿 0**；落地的形状与下面预计的不同，见收线那两条）

- 现象：门 #418 的 `Tests (vbp #3)` 红，本机拿不到**用例名** ⇒ 只能用排除法（`f8ab0ab6` 吃的是同一份语料、那一片在 #419 全绿）判成「已知那族按负载抖」，红因没落到字面上。这类"看得见红、看不见是谁"的格子，下一轮还会再来。
- 端点读数（三条都实测过）：`/actions/jobs/{id}/logs` 与 `/actions/artifacts/{id}/zip` 用本机 remote 里那把凭据一律 **401/403**（它没有 `actions:read`）；`/check-runs/{id}/annotations` **通**，但里面只有 `Process completed with exit code 1.` 加一条 Node 20 弃用警告。另：`/actions/runs`、`/actions/runs/{id}/jobs` 反而**匿名可读**（403 只在带那把凭据时出现 —— 反代也会随机把 JSON 顶成 HTML，判据是"正文不以 `{` 或 `[` 开头就重试"）。
- 已经可以换掉的通道：**`/repos/<o>/<r>/commits/<sha8>/check-runs?per_page=50` 匿名可读**，逐 job 的 `name / conclusion / output.annotations_count` 都在，`id` 就等于 job id。以后"门看不见"那一路先走它，别再守 watcher 与 run-id 那条（历史坑：watcher 自己提前 exit 0、jobs=0 一律当失败）。
- 欠的那一刀（小、常设）：`tests/run_tests.ps1` 每条用例失败时发一行 `::error::<用例名> FAILED <一句读数>` ⇒ GitHub 把它收成 failure annotation，下次门红就直接给用例名，不用再走排除法。做的时候两件事钉住：① **别动**现有的人读那套汇总行（两边读数都要在）；② 发完**不改退出码**（判红仍由汇总那一句决定）—— 这一格按设计不该自己变红，它只负责把名字带回来。

- **收线（门 #421 = run 37758983155、head `5fae2205`、branch dev、attempt 1 = 12 条 check-run 全 completed/success、非绿 0；发码清单仍 395 行零差 ⇒ 这一刀零发码变化）**：落地的形状与上一条预计的**不是同一个地方** —— 没去动 `tests/run_tests.ps1`，改在 ci.yml 那一层认它**已经打出来**的行。理由（本轮实数）：harness 66 个前缀点 / 80 个计票点 / 78 处 FAIL 字样，逐点插 `::error::` 等于重写；而点名要的只是那 66 行的排版，一层就够。上面钉住的两条一字未违：① 人读的输出**一行不落原样打到控制台**（判排版用的是内存里剥掉颜色码的副本，CI 那份日志与做之前一致）；② **退出码照旧透传**（判红仍然只看 run_tests 自己那一句）。
- 匹配式收成「用了例行的排版、且收尾不是 PASS/SKIP」，**不是**「行尾写 FAIL」：78 处 FAIL 字样里 26 处带原因（78 是含 "FAIL 字样的行数，其中 77 行是 Write-Host 直接打 token、另 1 行先把串拼进变量）（`FAIL (compile)` / `FAIL rc=…` / `FAIL: a | b`），另有 3 处 codegen 那族的计票点压根不打 FAIL 字样（`tests/run_tests.ps1:1509/1555/1580`）、只补一行 DarkGray 读数（missing / unexpected 那一族）—— 只认行尾 FAIL 会把这 26+ 处静默漏掉，而「没名字」在 CI 上看着跟「没失败」一模一样。读数原样进注解：它比 FAIL 那五个字值钱。
- 三档不许静默：`rc=0` 直接退（否则前缀行被别人的换行截断会在**绿轮**里凭空点出名字）；`rc≠0` 却一条名没抓到 → 自证；点到的名比 harness 那一行 `Results: … FAIL=n` 少 → 自证。注解正文一律 **ASCII**：`Write-Host` 的中文会被按系统代码页重编码（本机复算时两条中文自证行整条变 U+FFFD），与账 #267 的 `--emit-c` 是同一个根；用例名与类别本来就是 ASCII，所以点名不靠中文。
- 本机复算四档（测试副本由脚本从 ci.yml **机械抽出**、只把 run_tests 那一行换成桩 ⇒ 测的就是发货那一份）：`ok` 0 条注解 / `fail` 点名 3 条（三种排版各一：光 FAIL、FAIL 带原因、不打 FAIL 只给 DarkGray 读数）/ `short` 点名 1 条 + 「计票 3 只抓到 1」自证 1 条 / `weird` 两条自证；退出码 0/1/1/1 原样透传。外加 66 个前缀点 × 四种收尾的静态复算 **0 漏**。两条本机坑记档：`$_.Matches` 在 System.String 上不存在 ⇒ 命中了行也在取组那步抛 NullArray、注解照样 0 条（第一版就是这么红的，靠自证那条才没被误判成「没有失败用例」）；`Tee-Object` 落盘的编码 5.1 是 UTF-16、7 是 UTF-8 ⇒ 这条通道不能建在中间文件上。
- **订正上面那条端点读数**：`commits/<sha>/check-runs` 里的 `output.annotations_count` 本机今天**读不到**（12 条 check-run 一律 `ann=null`），能用的只有 `/check-runs/<id>/annotations` 那个数组；而且**每一格本来就有 1 条 notice**（Node.js 20 弃用那条）⇒ 回读点名要按 `level` 筛，别把基数当成结果。
- **CI 侧的红还没实跑过**：#421 是绿的，而这条通道按设计在绿轮里一声不出 ⇒ 目前的证据只有本机四档 + 66 点静态复算。**下一次门真红，第一件事是回读 `check-runs/<job-id>/annotations` 验点名**；若那条「少点」自证也响，就按它给的差额补匹配式，别再走排除法。

### B93 别人分支里「对我们有用」的两件已捞进 dev —— Join-Path 的 5.1 兼容真修 + vi 的三节待办（2026-10-08，**已出：门 #420（run 提交 head `dd463407`、branch dev、attempt 1）= 12 条 check-run 全 success、非绿 0**）

- **口径**（用户 2026-10-08 定）：只维护 `dev` 一条线；别人名下的分支（`origin/ferock/*`、`origin/vi/dev`、`github/main` 那条 orphan）默认不对齐，**但里面的内容对我们有用就捞过来**。所以判据不是「哪条分支多几笔提交」，而是「这段文字/这一刀在 dev 上有没有等价存在」。
- **怎么核分叉**（本轮实测的办法）：`git merge-base --is-ancestor <ref> github/dev` 找谁不在 dev 里，再对多出的那笔**按新增行**逐行去 dev 的文件版本里查在不在。用这条路核 gitcode 那份同名 `origin/dev`（与 github/dev 提交图分叉）独多的 `847ee9f4 docs: 账 #240 收线` —— 它新增 6 行，dev 现有版本 **6/6 都在** ⇒ 内容不欠，别去「合」它。
- **第一件 = `tests/run_tests.ps1` 与 `tests/regress_all.ps1` 里三参数 `Join-Path`**（出处 origin/ferock/0.10.7 的 `5735aff6`）：Windows PowerShell 5.1 的 `Join-Path` 只有 `-Path` 与 `-ChildPath` 两个位置参数，第三段得走 `-AdditionalChildPath`；写成位置参数会抛 `ParameterBindingException`，**而它待在双引号子表达式里 ⇒ 只剩一行噪声、那一段静默变空**。本机同一台实测（`Get-MsvcToolset`）：改前 5.1 下 **Include 段数 6→1、Lib 段数 3→0、空段 8**（VS 那一段之后全是空分号），而 pwsh 7 一直是 6/3/0 ⇒ 后果就是 ferock 写的 `cl.exe C1083 打不开 stddef.h/windows.h`，**谁用 5.1 跑回归，编译组整片假红**。改法：不套子表达式，把路径写成双引号内插（变量名后直接跟反斜杠再接下一段），两壳同解；**别改成 `-AdditionalChildPath`** —— 那是 7 才有的名字，5.1 下换写法只是换一种坏法。**同一个决定在两份 harness 里各抄了一遍**（run_tests 与 regress_all 那三行逐字符相同）⇒ 没有哨兵的话下一次还是「修一处漏一处」。
- **哨兵 = `scripts/check_ps51_joinpath.ps1`（第 36 道 [STATIC]）**：走真 AST（`Language.Parser::ParseFile` + `CommandAst`），只数**位置参数**（具名参数后面的值单独扣掉，所以 `-Path a -ChildPath b` 不误报），大于 2 就红；扫 `tests` 与 `scripts` 下全部 `.ps1` = 58 份。**它第一跑就钓出三条没人知道的**：① `regress_all.ps1` 那三行是同一缺陷的第二份；② `tests/run_small.ps1` 是「**.bat 的内容挂 .ps1 后缀**」（首行 `@echo off`，里面全是别人桌面的绝对路径，全仓 0 引用）⇒ 改名成 `run_small.bat`（`git mv`，字节不动）；③ `scripts/vswhere.ps1` **无 BOM 且 4 行中文在行尾** ⇒ PS 5.1 按 ANSI 读时吃掉换行、解析报「缺少右花括号」（记忆里那条老坑的又一件实物）⇒ 补 BOM。哨兵自带两条边界：文件数 < 40 直接 `exit 2`（「扫到 0 份」= 瞎了的判据），解析失败也算违规（否则 ②③ 这种「根本没被解析过」的文件永远隐身）。负控：临时放一枚三参数探针文件 ⇒ 违规 1 处、`rc=1`；删掉后 `rc=0`。
- **第二件 = `todo/vi.md` 多出的三节**（MCP/CLI 对外接口、`.lib/.obj` 产物、WinDevLib 32→64 API 自动翻译）：先按**节**对过集合 —— dev 那份 6 节与分支版逐字相同、dev 无独有节 ⇒ 可安全并表，numstat **52/0 纯插入**。WinDevLib 那条与我们「x64 的 Declare 形状是源码活」的口径是**同一问题的两种答案**（它主张按 API 签名自动把 Long 提级成 LongPtr，我们主张改 VB 源），留在 todo 里作对照，不是马上做。
- **`FUNDING.md`（vi 那 5 笔里唯一没捞的）**：dev 有 `CONTRIBUTING.md` 没有 `FUNDING.md`。它讲捐赠渠道与分配规则，属仓库治理与法务口径，**不由我拍**，等用户点头再并。
- **门 #420 的收线读数**：12 条 check-run 全 success（Build / `Emit manifest (shape oracle)` / syntax / compile / smoke / asm / bas#1#2 / vbp#1..#4），而发码清单那边仍是 `# expectation-check=PASS`、395 行零差 ⇒ 这一刀**零发码变化**（只动 harness 与文档，符合预期）。另记两件常设事实：① 本仓 [STATIC] 哨兵从 35 道长成 **36 道**；② 以后谁要在 5.1 下跑本地回归，现在跑得动了（实测 `-Category compile` 在 5.1 下 PASS=42 FAIL=0，改前那台同一条命令整片 C1083）。


### B94 量：`comMethods` 的消费者不止「四处」—— 四份文件五处 find，外加一处从不问这张表的写侧（账 #245 尾巴的开工测量，2026-10-08，**三刀都已出（第二刀 = 行为针 #270，第三刀 = 写侧那格 #245/#271，见本节末）**）

- **数据面**：`Symbol::comMethods` = `unordered_map<小写方法名, ComMethodSig>`（`src/semantics/symbol_table.hpp:282`，`ComMethodSig` 在 `:272-281`）。全仓**唯一**改写出口 = `insertComMethod`（`src/driver/driver_semantics.cpp:23-32`，落在 `:31` 那句 `sym.comMethods[key] = std::move(sig)`，`:28` 那条挡的是「setter 不许盖掉 getter」），三个调用点 `:153`（coclass 默认接口）/ `:218`（ComInterface）/ `:357`（ComGlobalNs 提升）；`clear` / `emplace` / `insert` / `erase` / 裸赋值 一处都没有。写侧是干净的，欠的一直在读侧。
- **读侧 = `src/backend` 里五处 `comMethods.find(`、落在四份文件**（哨兵 C4 钉的正是这四份**文件名**，`scripts/check_com_prop_type_authority.ps1:98-108`，扫描范围只到 `src/backend`）：`cgen_util_com.cpp:326`（`resolveComValue`，`:330` 先问 `isPropertyGet`）/ `cgen_util_com.cpp:800`（`resolveComMarkerForPack`，`:801` 同样先问）/ `cgen_expr_call_arg_emit.inc:89`（`:90` 问）/ `cgen_expr_call_com_bind.inc:615`（`:620` 零实参那一读、`:637` 带实参的形状）/ `cgen_expr_call_prelude.inc:209`（`:230` 起）。⇒ 「四处」是**文件**口径，动手时要按**五处**数。
- **五处干的是同一件事**：把 `mapType(returnType)` 翻成「发 `vb6_ComCallBSTR` / `Int` / `Double` / `Object` / 裸 `vb6_ComCall` 哪一个」。今天没有那个"一处出口"，事实上的基准是 `resolveComValue` 的早绑定分支（哨兵 C3 把它钉住：先看 `isPropertyGet`、类型未知就退 `vb6_VariantFromComResult(vb6_ComGetProp(` 而不许按 `unpackType` 猜档）。**分歧在 prelude 那一处**：`cgen_expr_call_prelude.inc:230-244` 拿到签名后直接按返回类型选，**没有 `isPropertyGet` 那道闸**（另四处都先问）⇒ 这就是「合一」要收的第一颗，也正是 C3 那条口径在第五处根本没人执行。
- **还有第五种问法压根不问这张表**：属性**写入**那一路 `comMarkerValueForWrite`（`cgen_util_com.cpp:696-708`）纯按 `packFn` 名字反推类型档，`driver_semantics.cpp:22` 的注释就是这么写的（「属性写入走名字化的晚绑定, 不查 comMethods」）；同一文件的晚绑定 `unpackType` 阶梯（`:358-368`）也是按上下文猜。⇒ 收口那天这两处要么接进同一枚出口，要么把「为什么不能问」写在出口旁边；C4 那句「名单外再开第五处读法必须先回来把口径并进来再登记」现在正好轮到它自己兑现。
- **判据面已有的**：`[STATIC] com_prop_type_authority`（`tests/run_tests.ps1:895` 的 `Test-ComPropTypeAuthority`，注册在 `:5081`）跑 C1..C4；行为面 = `test_earlybound`（`:2277`）/ `test_earlybound2`（`:2284`）/ `test_com_default_prop`（`:2288`）。**没有一处断言签名表的定序**（本轮把 tests/ 与 scripts/ 都 grep 过，零命中）⇒ #245 那条「定序」要么落成一枚真哨兵，要么在收线时写明它是本轮根因修（`VAR_PROPERTY` 的 `returnType` 未初始化那一族，见 §B73）的一部分，别再只挂名。
- 开工顺序（还没动）：先把五处 `find` 收成一枚出口（一次问完「该发哪个 `vb6_Com*` + 是不是 getter」，prelude 那颗按 C3 的口径补齐），C4 那份「四个文件名」随之换成「一处出口 + 名单外不许再读」；写侧那两处单独立一条「为什么它不问」。布局与形状类改动 ⇒ 按既有口径 x86 与 x64 都要真编译，护栏仍是 90 份 emit 逐字节相同那一条。
- **2026-10-08 晚 · 第一刀已落**：五处 `comMethods.find` 收成**一枚出口** —— `CCodeGen::comSigViewOf` （唯一读表处，交出 `found` / `isGet` / `retCType` / `sig`）+ `comGetterExpr`（属性 getter 的档位选择： 非确认的 getter ⇒ 返回空串，由调用方保留自己的晚绑定路；类型未知 ⇒ 一律 `vb6_VariantFromComResult(vb6_ComGetProp(…))`） + `comTypedCallExpr`（带实参调用的档位选择：未知 / 表里没有 ⇒ 裸 `vb6_ComCall`）。原来那五处 （`cgen_util_com.cpp:326` 与 `:800`、`cgen_expr_call_arg_emit.inc:89`、`cgen_expr_call_com_bind.inc:615` 的两支、 `cgen_expr_call_prelude.inc:209`）现在只问这枚出口；`src\backend` 里 `comMethods.find(` 只剩 1 处（就在 comSigViewOf 里）。
- **护栏 = 全语料发码逐字节**：新编的 C3.exe 跑 `scripts/emit_manifest.ps1` 对刚登记的那张表 = **395/395 全同、rc=0** （本机 `pe-lnk=14.29` 对 CI 的 `14.51` 照样逐行对上 —— 这条硬判据从此就是「零行为改动」的度量，不用再造 90 份 emit 的临时脚本）。 **反过来它也说明两件别的事**：① 两处按 C3 口径的**定向改道**（`resolveComMarkerForPack` 早绑定那一支原先「类型未知就落到 packer 反推那一档」、 `cgen_expr_call_com_bind.inc` 零实参 getter 那一支原先「类型未知直接发 StringProp」）**在 395 份输入里一次都没被走到** ⇒ 眼下只有哨兵守着， 行为面**欠一枚针**（要造一枚 `returnType` 落在未知档的 dual 属性夹具，才谈得上钉死这两条）；② 合一本身没改动任何一处发码， 所以「五份抄本互不一致」这件事以前在语料里没暴露过 —— 它只是随时会炸的形状。
- **哨兵跟着换形状**（`scripts/check_com_prop_type_authority.ps1`，C4 那句「名单外再开第五处读法必须先回来把口径并进来」这回兑现）： C3 改盯唯一出口 —— 出口里必须仍有 VARIANT 兜底、必须仍看 `sv.isGet`、且**整个函数体剥掉注释之后**不许出现 `unpackType` / `packFnHint` （第一条就踩了 `check_control_dc` 那条老坑：说明文字里本来就写着 `unpackType`，不剥注释 = 假红，方向照旧是**变严不是变松**）； C4 钉三件 —— `src\backend` 里读这张表的文件恰好 1 份、`comMethods.find(` 在其中恰好出现 1 次、以前那四处消费者必须**仍逐处** `comSigViewOf(`。 七档对照（`.build/b639_neg.txt`，全部在一份拷贝里做，不碰工作树）：绿控 rc=0；N1 摘掉 getter 闸 / N2 摘掉 VARIANT 兜底 / N3 让上下文猜档复活 / N4 重开第五处读法 / N5 让某个消费者不再问出口 / N6 在权威里放第二次 find —— 六档**各自点名相应规则**；N7（无关改动）仍绿。
- ~~仍欠的那格：COM 签名表的定序 + 上面那枚行为针~~ **2026-10-09 订正**：那枚行为针已落（#270，见本节末）；**定序那一格还开着** —— 政策没写死、`C3: COMSIG-DUP` 没响、也没哨兵，只剩这一件。
- **2026-10-08 晚 · 「定序」那一问先量清形状（读码，未动产品）**：写侧的政策其实只有两条 ——  `insertComMethod`（`src/driver/driver_semantics.cpp:23-32`）里 ① incoming 是 PropertyPut/PutRef **且**表里已有 `isPropertyGet` ⇒ 直接 return； ② 否则 `sym.comMethods[key] = std::move(sig)`（**后写覆盖，没有任何比较**）。而三个调用点各写**自己的符号**： `:153` 写 coclass 的 `sym`（成员来自 `cc->defaultIface->members`）、`:218` 写 ComInterface 的 `sym`、`:357` 写 ComGlobalNs 提升出来的 `sym` ⇒ 「三处来源互相覆盖」这条**不成立**（§B94 说的「四处」一直是**读侧**，写侧本来就一处）。
- 于是这一问收窄成一句话：**同一份 members 列表里出现两枚同名 getter 时谁赢** —— 现在是「后出现的那份赢」， 顺序来自 `parseTypeLib` 那串按 index 走的 vector（**同一枚类型库上是确定的**；跨机器那处分家是 §B73 的另一条线，已修）。 ⇒ 该做的**不是**改覆盖政策，而是两件：① 同名两份 getter（`returnType` 或 `params` 还不一样）出现时**响一声** —— 走 stderr 的 `C3:` 行， 因为 note 级诊断在成功的编译里根本不打印（§C 那条）；② 把政策**写死**（保第一份还是保最后一份、理由落在哪一方）并钉进哨兵 ——  眼下只有 setter-vs-getter 那一条有政策，**getter-vs-getter 一条都没有**。
- **那一半已经量完（2026-10-08 深夜，全 395 份输入）**：在**临时工作树** `.build/wt_b7` 里给 `insertComMethod` 插一条 stderr 探针 （`C3: COMSIG-DUP` = 真的发生了覆盖；`C3: COMSIG-KEEP` = setter 被闸挡住），冷编一台探针版 C3.exe 跑全语料 `--emit-c` 数命中。 主树与 dev 上那台 exe 都没动，探针也不进仓。读数（`.build/b678_dup.txt` / `.build/b679_dup.txt`）：
  · **探针是活的**（自证）：`KEEP` 去重后 **458** 条 —— 也就是「get/put 成对」这件事在语料里到处发生，只是被那条闸挡住了；
  · **真的发生覆盖 = 46 处**（按 符号+名字 去重）：`Scripting.Dictionary` / MSXML 6 一族 / Shell Folder View / NewTab 那些接口。 **订正**：我先前把这批命中记到 `tests/ve_units/Units.vbp` 头上了 —— 那是**读数工具自身的坑**： 字典按 (符号,名字) 存，留下的键是**最后**一次命中的输入而不是第一次；而且 Units.vbp 的 `Reference=` 只有 stdole2 一条。 这些符号根本不是哪个 .vbp 引来的，是**自动加载**（见下面第 ③ 条）。
  · 于是这一问的真相比 §B94 原来写的还窄一层：**put→get 这个方向本来就是对的**（getter 覆盖 setter 正是想要的）， 而 get→put 那个方向由那条闸兜住 ⇒ **对 get/put 成对来说，结果与枚举顺序无关**，这也是为什么它从来没出过症状。 **真正没政策、也没有实物的那一格 = getter-vs-getter / method-vs-method 同名**（例如 `Item` 两枚 dispid、形参数不同）：语料 **0 命中**。
  · **该做的两件因此不变，但理由变了**：① 一声 `C3: COMSIG-DUP`（走 stderr，note 级在成功编译里根本不打印 —— §C 那条） 至少要在 **old 与 new 的 `isPropertyGet` 相同** 时才响（get/put 成对不响，否则一编译就 458 条噪声）； ② 政策写死的那一句现在可以写明：「**getter 永远赢 over setter**（已由那条闸实现，两向都对）」+「**同类相撞（getter-vs-getter / method-vs-method）目前后写覆盖，且语料 0 例**」—— 后者是纸面缺口，判据得自己造（一枚同名两 dispid 的类型库夹具，或至少一枚哨兵钉住「同类相撞必须按 memid/vtableIndex 定序，不许按容器遍历序」）。
- **行为针（#270）的配方已量好，下一轮照这个做，别再试错**（三条实测）：
  ① **「类型未知」那一档到底是哪些** —— `CCodeGen::mapType`（`src/backend/cgen_base_type.cpp:19-63`）只有   `BSTR / int32_t / int16_t / double / float / void*` 六档会被 `comGetterExpr` 认，其余全落 VARIANT 兜底；  按 `src/common/types.hpp` 的枚举序对着普查里那批 `rt=` 读数：  **`rt=12` = `Variant` → `"vb6_VARIANT"` = 未知档**（`rt=3`=Long、`8`=String、`9`=Object、`11`=Boolean 都是已知档）。  所以针面要挑的是**返回 `Variant` 的属性**，普查里现成的样本 = `IXSLProcessor.input/output`、`IMXWriter.output`、  `IXSLTemplate.stylesheet`(`rt=9` 不算)、`Dictionary.item`(`rt=12`)。
  ② **枚举返回的属性不算**（我原本拿它当靶子，实测推翻）：`StdFont.Charset`（`FontCharset` 枚举）两台编译器都交   `vb6_ComGetIntProp(f, L"Charset")` ⇒ `mapTypeDesc` 把 `TKIND_ENUM` 折成 `Long`，落在已知档，永远走不到兜底。  （探针工程 `.build/b682probe/`，用完可弃。）
  ③ **不用写 `Reference=` 也有 COM 符号可绑**：一个**只有 .bas、没有 .vbp、零 Reference** 的输入，  `--dump-symbols` 里就有 **41 个 ComClass + 125 个 ComInterface**（`Folder` / `Dictionary` / `FileSystemObject` / MSXML 一族…），  而 stderr **一条 `Loading TypeLib ref` 都不打** —— 这就是 `cgen_base_type.cpp:374-378` 那段注释说的「类型库自动加载」。  两点用处：针面可以完全不碰机器路径；两点风险也要写进账：**(a)** 它与 §B73 那条「`Reference=` 不许指向这台机器恰好有什么」是同一族，  只是更隐蔽（连 .vbp 里的一行都没有）；**(b)** 它不响 ⇒ 哪天 CI 的 runner 少注册一枚库，发码会**静默**换一档（发码清单会红，这层兜得住，  但归因会很难 —— 建议顺手记一条：以后见 `comMethods` 相关的清单红，先问「这台机器有没有那枚库」而不是先怀疑 cgen）。

- **2026-10-09 · 第二刀 = 那枚行为针已落（#270），同时把「两条改道大概走不到」这条猜测量死**：在**临时工作树** `.build/wt_b8` 给五处消费者各插一条 `C3: COMPATH` 探针（站点 + `found` / `isGet` / 成员名 / `retCType` / 命中与否），冷编探针版跑全 395 份输入。读数：`.build/b689_paths.txt`（计数）、`.build/b691_paths2.txt`（带成员名）。
  · **五处里语料走到的是四处**：`resolveComValue` **689** 次（`Bold`/`Italic`/`Underline`/`Strikethrough` 各 ~100×int16、`Name` 87×BSTR、`Size` 90×double、`Charset` 77×int16…）、实参打包那处 **50** 次、`cgen_expr_call_com_bind.inc` 的早绑定块 **62** 次、`resolveComMarkerForPack` **2** 次（`tests/VbQRCodegen-master/test/Project1.vbp` 的 `Width` 与 `Height`，都 `ret=int32_t` 且命中）。⇒ **「pack 那一支大概是死路」被否掉：它是活的，只是只有 2 份命中。**
  · **com_bind 那块里「零实参 getter」那一支语料 0 次，但它不是死代码**：同一块有一条 `GetWindow`（VBFlexGridDemo，`f=1 g=0 args=0`）走进块内、只被 `isGet` 那一格口径拦住 —— 为分清这两件事专门加了纯观察的 `S4b`（看了不改发码）。⇒ 缺的不是形状，是「表里那条成员登记成 getter」这一格实物。
  · **于是照这个形状造出真夹具，两支都点亮**：`IXMLDOMAttribute.value` = MSXML 6 里返回 `Variant` 的 `Property Get` ⇒ `mapType` 交回 `vb6_VARIANT`（未知档）。`at.value` 走 `resolveComValue`；`at.value()` 走 com_bind 零实参那一支（探针记 `COMPATH S4a mem=value`）。**顺带第二条结论：`d.Count()` 这种「调用形 + 零实参 + 登记为 getter」在 C3 的语法面下真能写、也确实落到那一支。**
- **#270 的针面已进回归**：夹具 `tests/test_com_variant_prop.bas`（注释全 ASCII），登记 `tests/run_tests.ps1:2359-2379` 三条判据 —— 全部复用现成的 `Test-EmitcShape` / `Test-EmitcAbsent`，没造新机器。
  · 四行各写进各自的变量 ⇒ 每行发码唯一，**四枚消费点分开钉**：`v = at.value`（Variant 目标）、**`s1 = at.value`（目标是 String —— 这一行就是「按右值上下文猜档」那一刀唯一会红的）**、`s2 = at.value()`（零实参那一支）、`n = Len(at.value)`（实参打包那一支）。四行都必须交回 `vb6_VariantFromComResult(vb6_ComGetProp(at, L"value"))`；`s1` 与 `s2` 只差变量名 ⇒ 钉的是**「拼法不选档」**。
  · 两面钉：证人 `s3 = at.nodeName`（同一份库里真 BSTR 档）必须仍发 `vb6_ComGetStringProp`，挡掉「一律兜底」蒙过判据；反面 `Test-EmitcAbsent` 禁止 `GetStringProp` / `GetIntProp` 落在 `L"value"` 上。
  · **负控 = 真会红，而且只它红**：在 `.build/wt_b8` 把未知档那一行改成「猜 String」重编一台（`.build/b695_mut.py`）⇒ `cvp_emitc_variant_fallback` **四条 needle 全缺**、`cvp_emitc_no_context_guess` **点名那条形状在场**，而证人那条仍 PASS（`nodeName` 本就在已知档，变异没把一切都改掉）。再把这台变异版跑全语料清单 = **396 行里只有新夹具那一行哈希动**，其余 395 行逐字不动（`.build/b698_mut.txt`）⇒ 针既咬得住、又只咬这一格。
  · **A/B**：同一份夹具在 §B94 之前的 BASE（`.build/b630_base_C3.exe`）与之后的 NEW 上发码**逐字节相同** ⇒ 这是**回归针，不是修复针**，钉的是「合一之后这条口径别再漂回去」。
  · **清单跟着进一格，但别用 `-Bless`**：`emit-manifest.expected.txt` 395 → **396** 行（`+1/-0`，只手插新夹具那一行）。理由：本机 `Sort-Object FullName` 与 CI 那份对同前缀的 `*.bas` / `*.vbp` 次序不同（`.` 0x2E 排在 `_` 0x5F 前），`-Bless` 会把 **23 行一起重排**成噪声（实测过，回退了）。插完 `compare_emit_manifest.ps1` 本地读 **396/396 全同**；次序不参与判定，所以 CI 那边照样对得上。
  · **37 道 `[STATIC]` 哨兵本地全 rc=0**。新增 `.bas` 不碰任何「恰好 N 处」的普查：`check_vbp_fixture_census` 只数 `.vbp`、`check_fixture_timer_close` 只数 `.frm`（下限 20）；`tests/` 那侧也没有「每枚 .bas 必须登记」的闸（只有 `run_tests.ps1:2484` 那条注册表自检，它走的是**已登记队列**，不是目录）。

- **2026-10-09 · 第三刀落之前，先把上一轮那句「同类相撞语料 0 例」订正掉**。这一轮的探针不再只问「有没有覆盖」，而是问「**这次覆盖会不会改掉读取方真用的东西**」（`isPropertyGet` / `returnType` / 形参表的 名·型·方向 —— ByRef 出参探测读的就是后两样）。全 396 份输入的读数（`.build/b705_ambig.txt`、`.build/b708_quiet.txt`）：
  · **blessed 那一向 put→get = 21930 次**。上一轮写的「46 处」是**按 (符号,名字) 去重之后的键数**，不是出现次数 —— 那句「46/46 全落在 put→get」方向没错，但把总量说小了一个数量级（同一条坑：要按**每条的重数**比，不能拿「这个词在这个块里出现过」当归因）。
  · **同类 put→put = 994 次 ⇒ 「同类相撞语料 0 例」是错的**：实物就是 `Dictionary.item` / `IDictionary.item` —— `put_Item` 与 `putref_Item` 折成同一个小写键。它三项全同、读形没变，所以从来没有症状；**真正 0 例的是「读形变了又不是 blessed」那一种**。这条分界比上一轮的口径可判定得多，产品与判据都按它落。
- **产品侧落点（还是那一处唯一写侧 `insertComMethod`，`src/driver/driver_semantics.cpp`）**：① 政策两条写在出口旁边；② 新助手 `comSigReadShapeDiffers` = 「这次覆盖会不会改掉发码读到的东西」这一问的唯一住所（比 getter 标志 / 返回档 / 形参 名·型·方向）；③ 一声 `C3: COMSIG-AMBIG` 走 stderr（§C 那条：note 级诊断在成功的编译里根本不打印），只在「读形变了 **且** 不是 blessed」时响，**响完仍然后写覆盖** —— 它是等实物来定政策的哨子，不是判死。
- **两半边都有实测**：静的那半边 = 全 396 份输入 **0 条 AMBIG**，且发码清单 **396/396 逐字节不动**（stderr 不进哈希；顺带证了 `memid` / `vtableIndex` 不被发码读）。响的那半边 = 在 `.build/wt_b8` 把默认接口第一枚 getter 换个 `returnType` 再插一次 ⇒ 同一份 `tests/test_com_default_prop.bas` 当场响 **45 条** —— 而那枚输入正是新判据 `csa_emitc_no_ambig_noise` 盯的那一份，所以谁把条件放宽（去掉 `!blessed`，或让助手无条件 `return true`），CI 上就点名它。**判据两头：一条断它不许响，一条（第 38 道哨兵）断它凭什么响。**
- **第 38 道哨兵 `scripts/check_com_sig_collision_policy.ps1`**（`[STATIC] com_sig_collision_policy`，登记数 37 → 38，`PSParser` 0 错、BOM+CRLF 原样）：W1 两张表的写侧各**恰好一条赋值语句**（按**出现次数**数，不按文件名 —— 第一版按文件名数，同一个文件里多开一处照样绿，M6 档把它抓出来了，已改）；W1b 事件表 `comSourceMethods` 也登记在册（它今天不吃这条政策，哪天要吃必须先回来并口径）；W2 getter-beats-setter 那道闸必须仍真的 `return`；W3 响声必须问 `comSigReadShapeDiffers(it->second, sig)`、条件里必须带 `!blessed`、标签必须就是 `C3: COMSIG-AMBIG`；W4 助手必须真比那五项；W5 必须以 `return false;` 收尾（不许退化成无条件响）；W6 写侧调用点恰好 3。**九档对照全 MATCH**（`.build/b714_neg.txt`）：绿控 + 无关改动绿，七档假改动各自点名相应规则。
- **本轮自己踩到的实现坑（就记在这节里）**：`oldSetter = it->second.isPropertyPut || ...` 写在「先判 `it != end()`」**之前**，等于对 `end()` 解引用 —— 第一版就这么把编译器自己打成 `0xC0000409`，每份输入一进来就崩，`--emit-c` 连发码都到不了。教训两句：往哈希表里加东西的函数**先判存在再解引用**；改完先跑**一份最小真实输入**（本机当场三条输入全崩，而那道新哨兵那时是绿的 —— 静态判据看不见运行期的自己）。
- §B94 到此三刀落完。这一格剩下的只有一句「等实物」：真出现两枚同类 getter 相撞且返回档不同时按什么定序 —— 哨子会先响，再由政策决定（今天的默认仍是后写覆盖）。

### B95 四道 [STATIC] 哨兵从来没被门跑过 —— 其中一道早已过期成假红（2026-10-08，**已修：登记 + P3 换形状 = 门 #428（head `1564827f`）12 格全绿；那句 census 也已钉成第 37 道 = 门 #429（head `c2fa956d`）12 格全绿**）

起因是本轮把 36 道 `scripts/check_*.ps1` 全跑了一遍（§B94 改了其中一道，想确认没撞别人）：**35 绿 1 红**，红的正是 `check_quit_pump_invariant.ps1` 的 P3 —— 而门 #420/#421/#424/#426/#427 全是绿的。两条读数一起看才知道不是产品：
- 逐份数登记（`tests/run_tests.ps1` 里 grep `check_.*\.ps1`）：`scripts/` 有 **36** 道，登记进 harness 的只有 **32** 道。 从没跑过的四道 = `check_quit_pump_invariant` / `check_com_marker_write_sites` / `check_set_rhs_object_context` / `check_udt_member_dims`。 ⇒ **哨兵写在 scripts/ 里 ≠ 有门禁**：这道闸的覆盖面由「登记」那一行决定，而登记本身没人 census。这一格与 #166「两个权威」同族 —— 名单在两个地方各存一份。
- P3 那条红的成因**不是不变量丢了，是判据的形状跟不上**：它把 DoEvents 的函数体用 `\{[\s\S]*?\}` 懒配到**第一个 `}`** ——  那是内层 `if` 的收尾（3e9c1baf 在臂里加了「第二抽就 `vb6_End()`」那一层），于是窗越收越小、真实现场落在窗外； 再加上它要匹配的是一条**单行**写法 `… WM_QUIT) { vb6_RePostQuitIfTaken(&msg); break; }`，而那句话后来被拆成多行还夹了注释。 ⇒ 同一条坑第二次出现（第一次是 `check_control_dc` 被别人的 1350 字符函数体打红，口径写在 §C）：**扫窗不许定长，判据不许吃行尾排版**。
- 修法是把它换成结构问法：扫窗懒配到**第一个顶格 `}`**；判据前先剥 `//…` 与 `/*…*/`（那条臂里现在本来就写着中文说明，  不剥就会把散文读成一次调用 —— 同族另一半那条坑）；然后按**顺序**问三件事：`WM_QUIT` 那一步之后，  `vb6_RePostQuitIfTaken(&msg)` 必须排在 `break;` 之前，且两者都必须排在 `TranslateMessage` 之前（先到的那件若是 Translate，quit 就被内层泵吃掉了）。 顺带新写 **P6** 钉住 3e9c1baf 定下的那条口径：同一条 quit 被抽到**第二次**（`>= 2`）⇒ `vb6_End()` 恰好一次，且 `C3_NO_QUIT_IN_DOEVENTS` 那道开关不许消失。
- 七档对照（全在一份拷贝里跑，`.build/b649_neg.txt`）：绿控（未动 / 还原后）rc=0； N1 摘掉 hand-back、N2 臂里不 break、N3 让 TranslateMessage 抢到前面、N4 把 `>= 2` 改成 `>= 1`、N5 撤掉 `vb6_End()`、N6 改名那道开关 ——  **六档各自红并点名相应规则**。
- 登记动作：四道一起补进 `run_tests.ps1`（照现有 `Test-*` 那套逐字抄形：`$script:total++` / `[STATIC] <label> ...` / rc=0 才 PASS /   失败打前六行输出），登记数 **32 → 36**，`PSParser` tokenize **0 错**，BOM+CRLF 原样。登记之后门才会真的替这道闸守着；  **这一句下一段就钉上了**（第 37 道哨兵，见下面三条读数）。
- **第 37 道 = `scripts/check_static_sentinel_registration.ps1`**（登记后本机跑绿：`37 sentinels on disk, all registered`）。 三条规则：**R1** 盘上每一道 `check_*.ps1` 必须被 `tests/run_tests.ps1` 按**文件名**引用一次以上（新做哨兵忘记登记 = 当场红并点名）； **R2** 引用到的每一个 `check_*.ps1` 必须真在盘上（登记一个不存在的名字 = 那一格永远看不见失败，与 `Test-Compile` 重名遮蔽同族）； **DETAIL** 红的时候报两边的条数（`onDisk=` / `referenced=`），好分辨是漏登记还是名字打错。
- **R1b 是这一道自己的负控教出来的**：我拿「把调用行删掉、只留 `function` 定义」当假改动，R1 居然照样绿 ——  因为函数体里那句 `-File …\check_x.ps1` 本身就是一次引用。⇒ 补一条按结构问的 **R1b**：引用某道哨兵的那个 `Test-*` 函数必须**还有一处调用** （数 `^\s+Test-X\s*$`），没有就点名 dead cell。**覆盖面是「跑不跑」，不是「提没提」** —— 与这一格的起因（存在面 != 登记面）是同一件事的两层。
- 五档对照（`.build/b663_neg.txt`，在一份 `scripts/` + `tests/run_tests.ps1` 的拷贝里跑，不碰工作树）：绿控两档 rc=0（发货态 / 还原后）； N1 只删调用行 ⇒ **R1b** 点名 `Test-StaticSentinelRegistration`；N2 盘上多一道没登记的 ⇒ **R1**（`onDisk=38 referenced=37`）； N3 引用一个盘上没有的名字 ⇒ **R2**（`onDisk=37 referenced=38`）；N4 整段登记（函数 + 调用）删掉 ⇒ **R1**。 顺带一条自证：**这一道第一次跑就是因为自己没登记而红** —— R1 在那一刻被真实地验活了一次，之后才把它自己补进 harness。
- 全道复跑（含这一道）：**37 道绿、0 红**（`scripts/check_*.ps1` 逐道 powershell 跑，读数 `.build/b666_sweep.txt`）。  **CI 侧也数到同一件事**：门 #429 的 `Tests (syntax)` 那格里这 37 道全 PASS（含新登记的四加一道）， 而这一道自己被登记这件事由它自己守着 —— 它第一次跑就是因为没登记而红。

### B96 内在常量其实有三处答案 —— RTL 的 `#define`、发码的逐名折叠、语义层那张表（账 #218 第三刀，**已出：门 **#432**（run 37864959948、head `18821555`、branch dev、attempt 1）= 12 条 check-run 全 completed/success、非绿 0，其中 `Emit manifest (shape oracle)` 也在内**）

**先订正自己的量法**：那张常量表住的是**两份**片段 —— `src/semantics/builtin/builtin_consts.inc` + `builtin_consts_ext.inc`（后者在 `registerBuiltins` 里排在后面，两段共用同一个 SymbolTable）。上一轮只读了第一份，于是报出「表 292 / 只在折叠里 26」，实际是 **表 485 名 / 折叠 56 处（54 个不同名字，`vbObject` 与 `vbUseSystemDayOfWeek` 各被折了两次）/ RTL 宏 14**。这一格不是小数点问题：「谁不知道谁」的判断一旦建立在漏了一份的名单上，改法就整条歪（下面第二条是它的账单）。

**第一轮的真实红法**：把「只在折叠里」的 26 枚连同「只在 RTL 里」的 14 枚一起插进 consts.inc ⇒ 其中 6 枚（`vbGeneralDate/vbLongDate/vbShortDate/vbLongTime/vbShortTime/vbUseSystem`）ext 那份早就有了 ⇒ **每一份工程**编译都响 6 行 `error VB3002: 重复声明`（`.build/b738_run.ps1` 的读数，rc=1 而日志里没有一条 `error C` —— 语义层的错在构建日志里就是这么安静）。撤掉那 6 枚、只搬其余 34 枚，并把这条规矩钉成哨兵的 C4。

**缺陷读数（BASE = HEAD `56d12f51` 本机冷编那台，NEW = 本刀之后）**：
- 没有 `Option Explicit` 的模块里 `p = vbPicTypeBitmap`，发码是 `p = vb6_ChkLong(vbPicTypeBitmap)`，而 `vbPicTypeBitmap` 是**语义层按未声明标识符建起来的隐式 Variant 局部**（同一过程里还发 `#pragma push_macro("vbPicTypeBitmap")` + `#undef vbPicTypeBitmap` + `vb6_VARIANT vbPicTypeBitmap = vb6_VariantEmpty();`）⇒ 真跑交出 **0**（VB6 是 1）；拿它拼串交出的是**空串**（Empty 拼出来什么都没有）。夹具里那三条赋值在 BASE 上是 `0,0,0`。
- **一枚名字两个答案**：`Const c = vbUseSystem` 折出 **-1**（读的是 ext 表），同一份文件里直接读 `vbUseSystem` 折出 **0**（读的是发码折叠）。谁对？仓库自带的 VB6 手册 `docs/vb6-manual/09-常数/Date 常数.md` 两张表都写 **0** ⇒ ext 那份错了（就在 `vbUseCompareOption(-1)` 隔壁，八成是让两枚混了一次），已按手册改成 0 并把值钉进哨兵 C5。
- 反过来说，**带 `Option Explicit` 时那 14 枚读得是对的**（裸 C 名撞上 RTL 的 `#define`）。⇒ 判据必须挑那一形**不带** Option Explicit 的模块，拼串那一形当覆盖就是自欺（夹具头部因此写明「故意不带」并留了原因）。

**改**：34 枚进表（14 枚 RTL + 20 枚原先只在折叠里；6 枚 ext 已有的不搬）；删 26 行逐名折叠、删 14 行 RTL `#define`；RTL 变了就在 `src/driver/c3rtl.rc` 末尾记一行 re-embed（那一格的血泪见 memory `rtl-reembed-c3rtl-rc-landmine`）；ext 的 `vbUseSystem` 改成 0。**刻意没动**那 54 枚「表里已有、发码仍逐名折叠」的存量 —— 那是另一刀的活，这一轮先把数量钉死不许长（C3 棘轮：出现次数 56 / 不同名字 54）。

**逐行归因**（`.build/b759_att3.py`，原始字节 + `emit_manifest.ps1` 同一套归一化，不是 PowerShell 重编码那份）：396 行清单里 **15 行**变了，15/15 全部由三类指纹解释完 —— A 名字换成值（含 `(13369376)` → `13369376` 这种括号形状）、B 隐式变量那套自救行整段消失、C 装箱比较塌成直接比较（`vb6_VARIANT _vcmp_N = 值;` + `vb6_VarCmpLongEq(&_vcmp_N, E)` → `-(E == 值)`；Common.bas 3 处 / VisualStyles 2 处 / FlexGridDemo 7 处，**行数差正好等于这些装箱声明被删掉的条数**：4324→4321、1568→1566、86130→86123）。另有 4 行（`cc_id/Id.vbp`、`cc_id/Id2.vbp`、`pkg_cls/samepkg_ok.vbp`、`ve_list/VeList.vbp`）**逐行多重集完全相同**、只差 `#include` 那一段的模块顺序 —— 那不是本刀的行为差，是**另一格缺陷**，立成 §B97。

**判据**：夹具 `tests/test_intrinsic_consts.bas`（40 枚名字 + 三条赋值那形 + `Const`/变量/直读三条 `vbUseSystem`），`Add-BasTest` 钉 16 条期望（BASE 上 `IC-ASSIGN=0,0,0`、`IC-RTL-PT=,,,,`、`IC-RTL-CONST=0`、`IC-VS-CONST=-1` 四条当场红，NEW 全绿；而 `IC-FOLD-*` 那 6 条**两台同数** —— 那是「只改数从哪来、不改数」的半边证据）；`Test-EmitcShape ic_emitc_reads_are_literals` 钉 5 行发码；`Test-EmitcAbsent ic_emitc_no_implicit_var` 钉 4 条自救形状不许在场；**第 39 道哨兵** `scripts/check_builtin_const_authority.ps1`（C1 RTL 里 `#define vb*` 必须为 0、C2 折叠里不许有表不认识的名字、C3 棘轮、C4 两份表片段不许重名/不同值、C5 三枚手册值、C6 表达式档棘轮）—— 六档假改动**各自红并点名自己的规则**（`.build/b761_neg.txt`），植完按 md5 还原、还原后复跑绿。

**下一格**：① 那 54 枚「表里已有、发码仍折叠」可以按棘轮一批批撤（每撤一批就改 C3 的数）；② `isConstIdent`（`src/backend/cgen_base_naming.cpp:219`）里那 11 枚 RTL 名的硬编码名单，在本刀之后只对「typelib 先注册的 EnumMember」那一档还有用 —— 真正的修法是让 `lookupConstSym` 也认 EnumMember，那会牵动 `wrapConstArgForByRef` 一串形状，单独量过再动；③ §B97（include 与模块 init 调用序的无序容器）—— 已出，见下一节。

### B97 多模块工程的 `#include` 段**与入口点里模块 init 的调用序**都来自一枚 `unordered_*`（**已出：门 #435**（run 435、head `d1f73e25`、branch dev、attempt 1）= 12 条 check-run 全 completed/success、非绿 0 —— 含 `Emit manifest (shape oracle)` ⇒ 那 29 行纯排列的登记与 `ccid_emitc_init_order` 那枚顺序针在 **CI 那台二进制**上也对齐）

**读数**（2026-10-09 查到源头并落地）：
- 因果链：`src/backend/cgen_base.cpp:207` 收的参数是 `const std::unordered_set<std::string>&`，
  它由 `SymbolTable::getExternalModuleNames()`（`src/semantics/symbol_table.cpp:462-471`）造出来，
  那里是**遍历 `moduleScope_->symbols()` = `std::unordered_map`** 捡 `isExternal` 的符号
  ⇒ 桶布局由「整个模块作用域的名字集合」决定；这枚 set 喂着三处会出顺序的码头：聚合 .c 的
  `#include` 段（`cgen_base_generate_c_open.inc:53`）、.h 的 crossmod `#include` 段（`…_crossmod.inc:56`）、
  **入口点里 `vb6_mod_<X>_init()` 的调用序**（`cgen_base_generate_entry.inc` 五处同形循环）。
  另五处迭代它的是纯 membership 查询（`cgen_base_type.cpp:258`、`cgen_expr_ident_builtin.inc:181`、
  `cgen_expr_member_class_module.inc:254`、`cgen_expr_member_m22_module.inc:98`、
  `cgen_assign_stmt_special.inc:532`），换序无后果，但类型一起改了，理由见哨兵 P2。
- 触发的证据来自账 #218 那刀的 A/B：**4 份输入**清单 sha 变了而**逐行多重集一字不差**，
  而这 4 份压根没用到那 40 枚名字，身上唯一的变量是「表里多了 34 个名字」。
- ⚠ **订正上一轮自己写的一句**：当时说「`modules_` 是按工程顺序排的 vector（= `Id.vbp` 的声明序
  `IdMain, IdIfaces, IdBlocks, CircleImpl, VbpImpl`）」—— 那是读码时的想当然，实测不对。
  `runParse` 末尾有一步**刻意**的 `std::stable_sort`（`driver_frontend.cpp:327-339`）把**类模块整体前移**
  （标准模块里 `Dim WithEvents x As 某类` 要那枚类符号先进表，否则晚绑定退化 + 运行期解引用野 vtable），
  实测 `modules_` = `CircleImpl, VbpImpl, IdMain, IdIfaces, IdBlocks`。所以这一刀的口径是
  「**要保的是确定性**」，权威 = `modules_` 的下标序（那枚有序 vector 本来就在调用点手里），
  **不是**「照抄 .vbp 的字面行序」；VB6 的模块级变量惰性到首次引用才建，这条 eager 的 init 块
  本来就是 C3 的实现选择 ⇒ 这句话别说过头。助手名也跟着改过一次（`extModulesInProjectOrder`
  → `extModulesInModuleVectorOrder`），名字比注释更不会被后人读错。

**改**：`getExternalModuleNames()` 改返回**去重升序**的 vector（成员集合一处没变，变的是顺序的来源）；
driver 用 `extModulesInModuleVectorOrder(i)` 按 `modules_` 下标造名单（泛型模板体照旧排除；符号表若报了
**工程表里没有**的名字，按字典序追加在末尾 —— 实测语料内是空集，留着只为「成员集合与改动前一致」）；
`CCodeGen::generate` 的参数与 `externalModules_` 成员类型跟着换成 vector；入口点那 5 份同形的 init 循环
收成**一处** `emitExtInits()`（调用点仍是 5 处 —— 那是入口点模板的形状数，不是重复实现）。
**刻意没动**：五份里三份发在「本模块自己的 init」之后、两份发在其**之前**（现 157 / 205 两格）——
own-vs-others 的相对位置是改动前就有的形状，动它是第二件事（真改运行期初始化序），要另立的口径。

**护栏**：
- 逐行归因（`.build/b795_perm.py`，原始字节 + 与 `emit_manifest.ps1` 同一套归一化）：397 行清单变了
  **29** 行（= 语料里全部多模块输入的数目），29/29 都满足「HEAD 与新的两份产物**逐行多重集完全相同**」
  且挪了位置的行**全部**落在 `#include "<X>.h"` 与 `*_init(...)` 两类的形状里（`moved=` 计数：
  VBFlexGridDemo 181 行、Charts 2020 77 行、cc_id 13 行…），**未解释 0**。登记后 compare 复算 397/397 相同。
- 顺序针 `Test-EmitcShape ccid_emitc_init_order`：把五行 `vb6_mod_*_init();` **当一枚 needle** 断
  （行序本身就是读数，那枚助手故意不折叠空白）。两台编译器 A/B 实测：HEAD 那台**不命中**、本刀**命中**。
- 真编译真跑（这是布局类改动，门只测默认架构）：`Id.vbp` 在 **x64 与 x86** 都 `rc=0` / 零条 `error C` /
  出 exe / 运行输出同前（`cc_id`）—— 本机 `.build/b794_abrun.ps1`。
- **第 41 道哨兵** `scripts/check_module_order_authority.ps1`：P1 名单出口必须 `std::vector<std::string>`
  且不许退回 unordered_set、P2 `externalModules_` 成员同上、P3 driver 必须经那枚排法（定义 1 次 + 调用 1 次）
  且本地变量不许退回 unordered_set、P4 入口点里发射 init 的循环**只许一枚**而 `emitExtInits();` **必须仍是 5 处**、
  P5 两处 include 发射循环各恰好 1 次。五种坏法（出口退无序 / 成员退无序 / 不排了直接吃符号表 /
  把 init 循环内联回第二处 / 复制一枚 include 循环）**各自红并点名自己的规则**，植完按 md5 还原后复跑绿
  （`.build/b799_neg97.py`）。全 41 道哨兵 `total=41 red=0`。
- 一条自我打脸：中途我用临时脚本自己算 sha、自己验 needle，报了「与登记不符 + 针没命中」，
  差点把一次纯注释/改名重构判成行为变化；官方工具（`compare_emit_manifest.ps1`）说 397/397 相同，
  两份 stdout 逐字节相同。**临时校验脚本自己也是被测对象** —— 与官方读数冲突时先怀疑脚本。

**下一格**：① own-vs-others 那个相对位置的分歧（两份模板把别人的 init 发在自己的之前）；
② §B99 那两笔「登记」缺口；③ 若要把口径升成「`.vbp` 的字面先后」，得先决定那步类模块前移还算不算数 ——
那是语义层的依赖需求，不能只改发码侧。
④ 门 #435 之后新加的第 42 道哨兵（见 §B99 末）不在那一轮里，它跟着下一趟门跑。
### B98 判据助手重名遮蔽 = `[COMPILE]` 那一格从 2026-09-24 起一条 cl 也没编过（**已出：门 #433**（run 数 433、head `de4c340d`、attempt 1）= 12 条 check-run 全 completed/success、非绿 0）

**读数**（先量再动，`tests/run_tests.ps1`）：
- 80 枚顶格 `function` 里**恰好一枚重名**：`Test-Compile` 在 303（`param([string]$Name, [string]$Source)`，体里
  `& $C3 $Source --output-dir $OutDir @IncArg` ⇒ **真构建**）与在 1612（`param([string]$Name, [array]$Sources)`，体里
  `Invoke-CodegenProj` ⇒ **只发码**）。PowerShell 里后定义的赢 ⇒ 遮档生效，构建那枚变成死代码。
- 被遮的那一格打印的仍是 `[COMPILE] $Name`，组注释写的是「综合测试 (编译+运行, 以 Main 为程序入口)」，
  于是 10 份夹具（`test_comprehensive.bas` / `test_comprehensive2.bas` + `$formTests` 里 8 枚 `.frm`）**不再产出 exe**，
  而**没有任何地方报红** —— 发码退出码 0 就是绿。这一族比 #166/#231/#245 那几格更阴：被遮掉的半边不产生噪声，它只是安静地不跑。
- 来源可查：构建那枚自 M5（`9924f3cb`）就在；遮它的那枚是 **`40eea3f1`（2026-09-24，那一笔的本意是给 `Test-CompileFail` 铺助手区）** 种下的，
  `git log -S'function Test-Compile {'` 只这两笔，HEAD 仍是 2 枚 ⇒ 已经遮了 15 天，其间每一轮门的 `[COMPILE]` 都是空的。
- 调用点清点：303 那枚的三个调用点（5247 / 5248 / 5269）全在被遮的那一格；syntax 组那一处
  `Test-Compile "ci_pos3_retval_receiver" @(两枚 .cls)` 传的是**数组形**，它落在遮档里其实是**对的** ⇒ 不能跟着回到构建那枚。
- 接回构建之前先量这 10 份夹具今天编不编得过（`.build/b777_compile_probe.ps1`，x64、进程内灌 vcvars）：
  **10/10 `rc=0 clErrLines=0 exe=1`** ⇒ 恢复构建不会带进假红。

**改**：只发码那枚改名成它真正做的事 —— `Test-CodegenOk`，标签跟着换成 `[CODEGEN-OK]`；
`ci_pos3_retval_receiver` 显式改调 `Test-CodegenOk`（行为逐字不变）；构建那枚留回 `Test-Compile`；
组注释订正成「综合测试与 P7 窗体夹具: **真构建**（cl + link 出 exe, 不跑）」—— 原那句「编译+运行」本来也不准，那一格从不跑 exe。

**第 40 道哨兵** `scripts/check_test_helper_integrity.ps1`（只扫 `tests/run_tests.ps1`，不起 cl）：
- H1 `function` 名字**不许重复**（按出现次数数；这一条对**所有**函数名，不限 `Test-*`）；
- H2 每个 `function Test-*` 必须在自己的定义之外被**代码行**提到至少一次。只数代码行 —— 整行以 `#` 开头的注释不算引用。
  这条不能松：「把调用点注释掉」是覆盖面安静消失的**第二种写法**，与重名遮蔽同后果，按全文数它永远绿；
- H3 两枚的身份钉死：唯一那枚 `Test-Compile` 体里必须有 `--output-dir`、且不许出现 `Invoke-CodegenProj`；
  `Test-CodegenOk` 必须存在且体里有 `Invoke-CodegenProj` ⇒ 谁再把只发码的那枚改回同名，当场红；
- H4 `Test-CodegenOk` / `Test-EmitcShape` / `Test-EmitcAbsent` 三者标签互不相同且都不是 `[COMPILE]`（日志要读得出这格只发了码）。
- 绿读数 `functions=81 Test-* helpers=68 no shadowing, all reachable`；五种坏法（植重名 / 注释掉一个调用点 /
  构建那枚不再构建 / 只发码那枚改回同名 / 标签换成 `[COMPILE]`）**各自红并点名自己的规则**，植完按 md5 还原、还原后复跑绿
  （`.build/b783_neg78.py`）。全 40 道哨兵重扫 `total=40 red=0`。

**下一格**：① 门 **#433** 是这一格**第一次真编译**那 10 份夹具 —— 本机只验过 x64，CI 的 toolset 上报红就是真红不是抖；
compile 那个 job 的耗时从「发码 ~0.4s/份」涨到「编译 23–41s/份」（10 份 ⇒ 约 +5 分钟，job 上限 45 分钟）。
② H2 的可达性只管 `Test-*` 前缀；`Invoke-* / Get-*` 那批内部助手重名会被 H1 抓到，但**不可达**抓不到 ——
要扩就得把前缀名单一起钉成 census，别默默扩出假红。③ §B97 已出（见上一节的读数与订正）。

**门读数**：#433 里 `Tests (compile)` 这一格是**第一次真编译**那 10 份夹具（本机 b777 先量过 10/10 出 exe），整格 completed/success ⇒ 恢复构建没带进假红，也没有把 job 顶到 45 分钟上限。

### B99 门 #434 唯一红 = 别人那一笔改了语料/RTL 而没做两次登记（归因完，登记已补；①②两道哨兵都已出 = 门 #436 / 门 #438）

**门读数回填**：门 **#436**（head `52df9e99` = §B99① 那一刀）= 12 条 check-run 全 completed/success、非绿 0；
门 **#437**（head `15e119f3` = 用户的 ai/032 rev2 整型提升）= 11 绿、**唯一红正是 `Emit manifest (shape oracle)`**，
逐 job 结论走 `commits/<sha>/check-runs`（本机新探针 `.build/b871_checkruns.py`，只认完整 sha）；
门 **#438**（head `19ba4241`，含 §B99② 与 §B100 两格 + 把那 7 行登记落账）= 12 条全绿 ⇒ **#437 那格红已由这一笔洗白**。
⚠ 一条值得留下的形状：**门红会顺着 dev 往下传染** —— `15e119f3` 那 7 行不登记，下一个推 dev 的人（我）门里同一格照样红，
红得与他自己的改动无关。这正是「自救面必须便宜」的理由：登记现在是一条命令，不登记才是给下游添堵。

57b1f64e（ai/032：VB.NET 风格运算符 + 四档窄/无符号整型，用户点名的那批）之后，门 #434 的 12 条 check-run
里 11 条绿、唯一红是 `Emit manifest (shape oracle)`。归因走的是 §B92 那条回读通道（CI 每次把
`emit-manifest.txt` 提交到 `ci/emit-manifest` 分支，本机 `git fetch github ci/emit-manifest` 就拿得到
**它自己那台二进制**交的清单）：

- 读数为：CI 那份 **398** 行 vs 仓里 `emit-manifest.expected.txt` **397** 行；**共有的 397 行逐行完全相同**，
  差集只有一条 —— `tests/test_vbnet_ext.bas`（那一笔新加的夹具）。⇒ **红的原因不是发码变了**，
  是「往语料里加了一份输入而没把它的哈希登记进期望清单」，比较器把「新增 1」判红。
  顺带钉住一句好读数：那批新语法（lexer/parser/TypeSystem/cgen/RTL 五路都动了）**没有改动任何一份存量发码**。
- 第二格独立的漏：那一笔改了 `src/rtl/core/vb6rtl/vb6rtl.c` / `vb6rtl_builtin.h` / `vb6rtl_conv.c`
  三处 RTL，却**没 touch `src/driver/c3rtl.rc`** —— 而 RTL 是嵌在 C3.exe 资源里的（见 memory
  `rtl-reembed-c3rtl-rc-landmine`）：CI 每次从零重编所以看不出来，本机增量重编时 `rc.exe` 不重跑就会
  拿到**旧 RTL**，测出来的一切都是假的。本轮补了一行 marker（并把 69 条 `VE-*` 读数在本机真编真跑复核过：
  x64 `rc=0` / 出 exe / 采样 9 条全命中）。
- **§B99① 已出（第 42 道哨兵）**：`scripts/emit_manifest.ps1` 加了 `-ListInputs`（**枚举口径只此一份**，哨兵绝不自己再 glob 一遍 tests/ —— 那是给「谁算语料」开第二份权威），`scripts/check_manifest_coverage.ps1` 拿它问三件事：K1 枚举集合 == 期望清单的路径集合（两个方向的差集都要空，红时直接点名是哪几份）、K2 期望清单仍是 BOM+CRLF 且每行都长成 `sha256=<64hex> ascii256=<64hex> rc= bytes= <relpath>`（手改最容易弄坏的就是这两样）、K3 枚举本身不许退化（>=200 份、`.vbp` 与 `.bas` 两类都在）—— 没有 K3，「两边都空」会在 K1 上假绿。绿读数 `inputs=398 [.vbp=139 .bas=259], registered=398, sets equal`；三条坏法各自红：往 tests/ 丢一份未登记的 .bas ⇒ K1 点名它、往期望清单塞一条不存在的路径 ⇒ K1「no longer exist」、把某行哈希首字符换成 `z` ⇒ K2、把 `-ListInputs` 指向只吐两行的桩 ⇒ K3（`.build/b817_cov.py`，植完按 md5 还原后复跑绿）。**这一道把「少登记」从 CI 的二十分钟挪到了本机的一秒。**
- 一次自撞（值得记，因为它是 #78 那道的第一个真实猎物）：注册第 42 道的那支脚本**跑了两遍** —— 第一遍在「打印读数」时因 GBK 控制台崩在 UnicodeEncodeError（中文进 print 要走 `sys.stdout.buffer`），**而插入已经落了盘**；第二遍的锚点 `function Test-ComSigCollisionPolicy {` 仍然唯一 ⇒ 又插了一枚同名助手 = 正是 #78 那个形状。**第 40 道哨兵当场报红并点名 989/1003 两行**，否则 compile 那一格会安静地跑两遍、`TOTAL` 只多 1 而没人知道为什么。留下的一般式：**跑崩在中途的补丁脚本 = 半边状态，重跑前必须先核现状**；而且插入型锚点最好在插入后**不再匹配**（把新内容本身写进锚点，或先查同名函数在不在）。
- **§B99② 已出（第 43 道哨兵 `scripts/check_rtl_embedded.ps1`）**：这一格的读数**不问 git 而问工件本身** —— 解析 `src/driver/c3rtl.rc` 的 `^\s*(\d+)\s+RCDATA\s+"([^"]+)"` 名单，再逐字节走 `C3.exe` 的 PE 资源目录（`e_lfanew` → 可选头的资源数据目录项（PE32+ 在 opt+112 / PE32 在 opt+96，按 magic 分）→ `RT_RCDATA`=10 的三级目录 → 叶子 `IMAGE_RESOURCE_DATA_ENTRY{RVA,Size}` → 按节表 RVA→文件偏移），把每枚 id 的**内容 SHA-256**与磁盘那份文件对照。三问：R1 = rc 里每个 id 都必须在 exe 里存在且**逐字节相同**（红时点名 id、路径、两份字节数、两份哈希）；R2 = exe 里不许有 rc 没登记的孤儿 id；R3 = 两边各 `>=100` 份（没有 R3，「两边都空」会在 R1 上假绿）。绿读数 `OK check_rtl_embedded: R1..R3 (125 RTL resources byte-identical between disk and C3.exe, no orphan ids; exe sha=a029563f5878)`；登记为 `[STATIC] rtl_embedded`（`Test-RtlEmbedded`），随趟 43/43 全绿、`-Category compile` 53 PASS / 0 FAIL 且新那行 `[STATIC] rtl_embedded ... PASS` 确实在日志里出现。
  - 四条负控（`.build/b836_rtl_neg.py`，每份植入文件按 md5 `try/finally` 还原，还原后复跑绿）：A 往磁盘 `vb6rtl.h` 追加一行注释 ⇒ **R1 id 100 红**（`disk=829B` vs `embedded=778B`，就是「改了 RTL 没重 embed」的形状）；B 把 100/101 两行 `RCDATA` **引号内的路径**对调 ⇒ **R1 id 101 红**（`disk=778B` vs `embedded=35111B`，**门 #225 那格**）；C exe 带一条 rc 没登记的 id ⇒ **R2 红**；D rc 清单退化成 0 条 ⇒ **R3 红**。
  - ⚠ 负控设计的一课（第一版真没红）：**把 `100 RCDATA "…"` 与 `101 RCDATA "…"` 两整行对调**什么也没改 —— id↔路径 的对应关系还是对调前那一对，哨兵正确地绿。要造出 #225 那种红只许**交换引号里的路径**（或改 id 数字）。一般式：**造负控前先问「我动的那个字段，是不是判据真正读的那个字段」**。
  - 为什么弃用「diff 口径」（`src/rtl/**` 出现在本次改动而 `c3rtl.rc` 没出现就红）：那是**改动面**的读数 —— 要拿 base、只在提交那一点有效，而且往 .rc 里补一行 marker 就骗得过去。问工件是**状态**的读数 —— 任何时候跑都成立，并顺带抓住 diff 口径抓不到的两种病（.rc 改了没重编、id 与文件对调）。机器无关性：两侧都取自同一个 checkout（磁盘 RTL + 本机 `.build/C3.exe`），不比常量哈希 ⇒ 换机不换期望值；代价是它验的是**本机** exe，CI 那台由「每次从零全量重编，`rc.exe` 必跑」保证（门 #434 那次红就不是 RTL 不同步，而是缺登记 —— 两格独立）。
  - ⚠ 工具坑（PowerShell 里手写 PE 解析会连撞四样，全部由实测翻出，不是猜的）：①命令调用会把逗号分隔的实参**贪吃**成数组（`@(U32 $e, U32 ($e+4))` 被当一枚三参调用 ⇒ Object[]→Int32 转换报错）；②返回「数组的数组」会被**摊平**（一枚两字段条目读成两枚条目 ⇒ 语言级目录算错偏移，`rva 0x0 not mapped`），要用 `[pscustomobject]@{…}`；③切片要 `[Array]::Copy`，别拿字节数组的 `-join`；④资源目录里所有存的都是**相对资源节基址**的偏移，而叶子给的是 RVA，两套换算混用一次就读出 0 字节。
- **§B99 那条「留着的结构活」两格都已落地**：① = 第 42 道（登记面，见上），② = 第 43 道（内嵌面，见本条）。本账到这里没有未开工项。

### B100 形状门红了，别人不会处理 —— 现在红话里带着三种成因与一条命令（**已出：门 #438**（run 37886205704、head `19ba4241`、branch dev、attempt 1）= 12 条 check-run 全 completed/success、非绿 0，其中 `Emit manifest (shape oracle)` 也在内 ⇒ 新登记的 398 行在 CI 那台上复算对齐）

**现象与读数**（用户报「`Emit manifest (shape oracle)` 在其他分支推送时太容易挂了，人家不太懂怎么处理」，2026-10-09 实测）：
`15e119f3`（ai/032 rev2 = 整型之间的提升改成 VB.NET 的二进制数值提升）推上 dev 之后，那一格红。走 §B92 那条
git 回读通道（`git fetch github ci/emit-manifest` 拿**它自己那台**交的清单）+ 仓里唯一的比较逻辑复算，读数是
**期望 398 / 实得 398 / 相同 391 / 哈希不同 7 / 缺席 0 / 多出 0**。⇒ 门没坏、也不是漏登记（那是门 #434 那一格，
已由第 42 道哨兵管住）：**发码真的变了**，正确处置是「看过这 7 行、然后重新登记」。问题是「重新登记」过去的做法
要本机编一台 C3.exe、再跑二十分钟全语料，钥匙只在少数人手上 —— 于是这道保护对协作者变成了卡住别人的东西。

**改**（判据与登记逻辑仍然只有 `compare_emit_manifest.ps1` 那一份，新工具只是操作它的壳）：
- 红的时候把**三种成因分开**打出处置，并把下一步那条命令直接印出来：① `新增 N 份输入没有登记哈希` = 不是回归
  （新夹具本来就没有旧期望可比）；② `存量 N 份发码真的变了` = 先看是不是自己要的形状；③ `登记里有、这次清单里没有`
  = 夹具被删/改名或枚举口径变了，**这种不许用登记抹平**。CI 上同一批文字进 `::error::` 注解（§B92 那条通道），
  别人不必翻整篇日志。
- `-Bless` 改成**保持现有行序**（新输入按清单自己的顺序追加在末尾）。实测语料枚举次序与这张清单的历史次序差
  **68 个位置**（同一批路径、两种排法）：照清单整写一遍 = `git diff` 里 398 行全动而真变化只有 7 行，review 的人
  看不出哪里是真变化（§B97 那 29 行当初手 splice，躲的就是这个）。
- `-Bless` 的编码写死 **BOM+CRLF**（`[IO.File]::WriteAllText` + `UTF8Encoding($true)`）：`Set-Content -Encoding UTF8`
  在 Windows PowerShell 5.1 带 BOM、pwsh 7 **不带**，而 CI 那一步跑的正是 pwsh 7 ⇒ 谁在 CI 上登记一次就把第 42 道
  哨兵的 K2 弄红。实测两台（`powershell` 5.1 与 `pwsh` 7.6.5）输出**逐字节相同**（sha `3ec5ed3f7269`）。
- 新工具 `scripts/rebless_emit_manifest.ps1`：一条命令吃完上面这套 —— 默认吃 CI 交回的那份数（**不必本机装 MSVC**，
  blob 用 `cmd` 重定向取，字节不经控制台代码页），核对它出自哪笔提交、是不是当前 HEAD 的祖先（不是就拒绝，防把
  别人没验过的形状钉进期望；`-ForceHead` 才放行），只跑一次是「只看」，`-Bless` 才写，写完**自己复算 + 跑 K1..K3**
  才算成，最后打印那两行 git 命令。`-Manifest` 走本机清单那条路，`-AllowVanish` 是给「确实删了夹具」的出口。
- `CONTRIBUTING.md` 新增「门里那一格 `Emit manifest (shape oracle)` 红了怎么办」（三种成因一张表 + 命令），
  顺带订正编码表里那句「bat/ps1 为 GBK」—— 实测 43 份 `scripts/check_*.ps1` 全部是 **BOM+CRLF 的 UTF-8**，
  按 GBK 存会把哨兵自己读坏。

**一次自撞（工具的第一条猎物是它自己，值得记）**：`$rc = RunCompare …` 里那个函数**没转发子进程的输出** ⇒
PowerShell 把「一整串分类读数 + 末尾的退出码」整体当成返回值塞进 `$rc`，于是 `if ($rc -ne 0)` 对数组**恒为真** ——
登记明明成功（文件已经写对了）却被报成「登记被挡住」，而那些本来给人看的处置文字反而全被吞掉。
修法：函数里 `& powershell @argv | ForEach-Object { Write-Host $_ }` 再取 `$LASTEXITCODE`。
一般式：**用变量接住一个「既打印又返回码」的函数 = 接住的是数组**；`Write-Host` 走宿主流、不会被捕获，正好用它转发。

**本轮真登记了一把（落地读数）**：rebase 到 `15e119f3` 之后重编 `C3.exe`（sha `cd32f425947b`），本机全语料清单
与 CI 那台 `head=15e119f3` 的清单 **398/398 逐行完全相同** ⇒ 两台工具链（本机 pe-lnk=14.29 / CI pe-lnk=14.51）
交叉对上；登记 **7 行改写、行序零移动**，登记后复算 398/398 绿、K1..K3 绿（`inputs=398 [.vbp=139 .bas=259]`）、
43 道哨兵全绿、`-Category compile` = 53 PASS / 0 FAIL。
负控（都实测过）：删一条数据行而注脚仍写 `inputs=398` ⇒ 「发码清单自相矛盾」先红；注脚改成 397 ⇒ 「拒绝登记：
有 1 份输入…（`tests/acc/acc_fam_main.bas`）」且**目标文件一个字节没动**；补 `-AllowVanish` ⇒ 放行。

**顺带第二笔同一形状的漏（第 43 道哨兵的第一次实战）**：`15e119f3` 又改了 `src/rtl/core/vb6rtl/vb6rtl.c`(id 101)
与 `vb6rtl_builtin.h`(id 135) 而**没 touch `src/driver/c3rtl.rc`** —— 与门 #434 那格同形（这是第二次）。本轮补了
marker、重编，之后 `check_rtl_embedded` 报 **125/125 逐字节相同**。它的价值当场就兑现了：不重编的话，本机测出来的
那 69 条 `VE-*` 读数用的会是旧 RTL。

**欠着**：门跑起来才算收线。另一档没动、要拍口径 —— 要不要让这道门**对 PR 只报不挡**（`continue-on-error`），
只在推 dev 时才硬判红。现在它是硬的：新加的自救面已经把「处理它」压成一条命令，本线倾向保持硬判。

**登记面自己的一格漏（2026-10-11，第三十二刀撞见并已修，提交 `adc4b61f`）**：`compare_emit_manifest.ps1` 的 `-Bless` 只保**头部**注释块（`HeaderRows` 在第一条数据行就停），而每一刀的「登记来源」是写在清单**末尾**的 ⇒ 这一档会静默吃掉末尾那些注脚，还照样报「已登记」。门 #501 之后那次登记实测：末尾 3 行 → 0 行、注释 23 → 20、84975 → 81654 字节，而报告里只说「注脚 20 行」。修法两条一起：末尾那段一并收进写出（`@($hdr) + $ordered + $tailCmt`），落盘前**两头各数一遍注释总数**、对不上就 `exit 2` 拒绝登记（宁可挡下，不许吃掉），报告改成「注脚 N 行（头部 X + 末尾 Y）」。两头对跑（同一份带 3 行末尾注脚的清单，分别交给改前那版（从 HEAD 的 blob 取）与改后这版各登记一遍，`.build/b968_negctl.out`）：改前 tail 3→0 且报成功；改后 tail 3→3、字节数不动、复算仍 400/400 绿。⇒ 口径补一句并交给工具自己钉：**登记完必数注释行数**。

**逐 head 二分：定位「哪一笔发码改动让这一行变的」不必本机冷编基线**（同一轮用的手法）：CI 每一步都把自己那台的清单提交到 `ci/emit-manifest`（§B92 那条 git 回读通道），那一串提交本身就是「每一笔 head 上这一行长什么样」的历史 ⇒ `git fetch github refs/heads/ci/emit-manifest` 后按 head 逐格读哈希、看它在哪一笔第一次变，就够了。这一发用它把 8 行移动归到两笔（7 行 = `e589b1a3`，1 行 = 本刀自己那枚夹具），零编译；再配合**每行字节增量都是 13 的整数倍**（`vb6_ChkLong(` 加右括号 = 13）把成因钉到「x64 上多套了几层窄化检查」那一格 —— 这两条读数合起来才把「本刀只改 RTL，RTL 不进发码，那为什么形状门红了」这句问话答死。

### B104 那 5 条「窗体伪成员」不是缺项 —— 发码面早就答对了，只有诊断在报噪声（2026-10-09 量完，**未开工**）
> **2026-10-10 复测（把这节的标题从"未开工"改成"已出"的依据）**：探针 `.build/b556_probe/P104.frm`（`Option Explicit` + 裸写 `w = ScaleWidth` 与 `Debug.Print ScaleWidth, ScaleHeight, WindowState`）在现在这台 C3.exe 上：三条读数 `vb6_GetScaleWidth/…ScaleHeight/…WindowState` 各就位，**诊断一条都不报**（rc=0，stderr 空）。 这一格的堵点由 §B109（账 #278 第四刀，门 #445）解掉 —— 语义层现在include `common/form_pseudo.hpp` 并问 `kFormPseudoRows`（`semantic_analyzer_expr.cpp:7` 与 `:158`）。**本节原来的"未开工"是这刀之后没回头改的标题**， 不是还欠着一格。

探针 `.build/b902_probe`（一枚最小 `.frm` + `Option Explicit`，三种写法并排），实测产物与诊断**各是一半**：

| VB 源码那一行 | 发码实际发出 | 配的语义诊断 |
|---|---|---|
| `w = ScaleWidth` | `w = vb6_ChkLong(vb6_GetScaleWidth(vb6_hwnd_Form1));` | **VB3001** |
| `Debug.Print ScaleWidth, ScaleHeight, WindowState` | 三条都是 `vb6_Get…(vb6_hwnd_Form1)` | **VB3001 ×3** |
| `Debug.Print Me.ScaleWidth` | `vb6_GetScaleWidth(vb6_hwnd_Form1)  /* Form.ScaleWidth via Me */` | 无 |

⇒ 发码面早就有一条完整通路：`cgen_expr_ident_symbol.inc:7-34`（Fix 056 —— `isFormModule_` 时问 `getControlPropReadFn(Form, 名)`，命中就发
`读函数(窗体句柄)`），**缺的只是语义层跟着放行**。这正是 `host_pseudo.hpp` 页头那句症状的**反方向版本**：那边是「表里加一行而发码没跟上」，
这边是「发码跟上了而语义没跟上」⇒ 产物是对的、配一条噪声。所以 §B102 ④ 记的「真缺项 5 处」要改记「**诊断面**缺项」，
下一刀动的是语义层一处放行，不是 RTL、不是发码。

**为什么这一刀不是「加三行放行」**：那条通路的权威 `getControlPropReadFn` 是 `CCodeGen` 的成员函数（`src/backend/cgen_util_ctrl.cpp`，实测函数体 **496 行**），
而语义层只许 include `common/` 与 `semantics/`（证据 = `semantic_analyzer_util.cpp:1-6` 那六行家规）。要「跟着问同一张表」，先得让那张表**到得了 common**：

- 函数体本身干净 —— 实测零 `this` 依赖（`cIdent` / `moduleName_` / `symTab_` / `isFormModule_` 各 0 次命中），只吃 `(FrmControlType, propLower)` ⇒ **可整段搬成自由函数**；
- `FrmControlType` 现在住在 `src/project/frm_parser.hpp:78`，common 不许向上依赖 project ⇒ 要么把那个枚举一起下到 `common/`，要么给 common 那份出口换一个不带枚举的小签名；
- 19 处调用点全在 backend，留 `CCodeGen::getControlPropReadFn` 一行转发即可零改动。
- **形状上的真决定（这格的关键）**：`getControlPropReadFn` 的答案 =「`switch` 之前那批通用行」∪「`case Form:` 那批专有行」，而裸写通路今天拿的是**并集**
  ⇒ 放行也必须拿并集。**只在 common 立一份 Form 专有名单 = 又造一个权威**（与 §B103 分家那条同一个病）。所以这一刀 = 把那张表**整个**挪到两层都问得到的地方，不是补名单。

**下一条要先量的**（别照抄本节的结论）：剩下 7 处（`TmForm2` 1 / `Printers` 1 / 裸 `Controls` 2 / 裸 `UserControl` 1 / 常量 2）里，
裸 `Controls` 与裸 `UserControl` 很可能是**同一种形状**（发码有路、语义没放行 —— `cgen_expr_ident_builtin.inc:214-233` 那条 designer 分支就在那儿）。
每一处先跑一次「产物 vs 诊断」并排读数再定性，别把「census 里有这条」当成「产物是坏的」。

**⚠ 自纠（探针把自己骗了一次）**：第一版探针用 Python 列表拼 VB 源码，相邻两条字符串**漏了逗号** ⇒ Python 做相邻字面量拼接，
`w = ScaleWidth` 与 `Debug.Print …` 合成一行，于是行号与诊断全对不上（报在 `(6,35)` 那种文件里根本不存在的列上）。
把生成的夹具**原样 dump 出来看一眼**才发现。⇒ 夹具是脚本拼的时候，「先核夹具本身」是第 0 步，不是最后一步。

### B128 账 #278 第十七刀已出 = arity 哨兵认「这是一条声明」时被**行尾块注释**挡住 —— RTL 一族 8 对成对隐身，剥注释收成一处之后一起数进来（判据自己的盲区，2026-10-10，门 #461 attempt 1 全绿（run 38005452748、head `628f07b8`、12/12 全 completed/success、非绿 0、created→updated 10m53s；改后那版 `compared=1267` 在 `Tests (compile)` 那一格里跑绿 ⇒ 新覆盖面在 CI 上没报出任何一条真头尾不一致）)

- **怎么撞上的**：不是另起的一格，是上一刀的读数自己露出来的。撤掉 `vb6forms_prop_ctrl.h` 里那两行带尾注的声明（`…);  /* Recordset 链透传 */`）之后，`decl_names 1394` 与 `compared 1259` **一根毛没动**，只有 `def_names 2136→2134` 掉了两 —— 说明这两枚"声明"在 R1 里从来不存在。（第十六刀当时把这格记成"待拍"，本轮把它办成账。）
- **根因（读码，就一格）**：R1 认声明的判据是「整行去注释后以 `;` 结尾」，可它**去注释只剥 `//`** —— 成对的 `/* … */` 只在**续行**那一支才剥 ⇒ 单行声明带行尾块注释时尾巴是 `*/` ⇒ 声明侧不计数；而 R1 只比"两头都有"的名字，声明既然不存在，头尾**一起隐身**。RTL 头里这种写法是主流不是例外。
- **修法**：把注释剥离收成一个出口 `Strip-Cmt`（`//` 到行尾 → 成对 `/* … */` → 从残留的 `/*` 处整段切，那表示正处在跨行块里），首行 / 续行 / 参数表**三处都调它**。从前是同一件事抄三份、每份剥的法子还不一样 —— 本线那一族「一个决定抄两遍」第一次抄在**判据自己**身上。
- **实数（诚实记账：这是一张网，不是一条已存在的缺陷）**：`decl_names 1394→1402`、`compared 1259→1267`（多 8 对）、`def_names 2134` 不变，**R1 全绿** ⇒ 这一族今天没有真的头/体不一致；它防的是"下一次头追不上体"正好落在一枚带尾注的声明上（账 #240 那一格的形状）。
- **负控 3/3**（改前那版与改后这版**并排跑同一棵副本树**，靠 `-Root`）：N1 基线两台都绿；**N2 只给一枚"带行尾块注释"的声明多加一枚形参 ⇒ 改后这版红并点名 `vb6_Data_RecordsetObj`，改前那版答绿 rc=0** —— 这一条就是盲区存在的**实证**（不是"我觉得会漏"）；N3 还原后两台同绿且新版 `compared` 只增不减（1259→1267）⇒ 覆盖面变宽没变窄。
- **自己的负控先骗了一次（值得记）**：第一版 N2 的补丁串多写了一个 `)` ⇒ 参数表里留下括号 ⇒ `Get-Arity` 按头注里写明的边界"认不出就跳过"（返回 −1，不计入也不报红）⇒ 两台都绿，看着像"盲区根本不存在"。⇒ **负控不红时第一嫌疑是自己的假改动没造出真形状**，不是判据坏了；查法是去副本里把那行 mutated 文本打出来看一眼。
- **顺手**：R2 的实测注释从「112 份 / 1108 对」订正成「§B128 前 112/1108、后 127/1267」；`tests/run_tests.ps1` 里 `Test-RtlProtoArity` 的说明补一句「认声明前先剥 `/* … */`」。**零产品代码改动**（只动那道哨兵与 harness 的注释）⇒ 形状门 398 行不必重钉。

### B127 账 #278 第十六刀已出 = RTL 那头两枚 0 调用者的 recordset 出口一起撤掉：§B124 那份清单到此清完（哨兵长出 RTL 那一头，2026-10-10，门 #460 attempt 1 全绿（run 38003417727、head `e7ff91a9`、12/12 全 completed/success、非绿 0、created→updated 10m30s；含 `Tests (compile)` 那一格里新钉的 RTL 面，与 `Test-VbpBuild "c29data"` / `c29data_x86` 两架构真跑 ⇒ 撤掉两枚 RTL 导出没把 recordset 那条链碰坏））

- **为什么这一格能单独走、不必等口径**：第十五刀撤的是**读者**（发码侧那五格 `vb6_Data_Self(` 前缀判据），留下 RTL 的两枚**被读者** —— `vb6_Data_Self`（把 Recordset 求值透传成控件自己）与 `vb6_Data_FieldValueStr`（把 `Fields("x").Value` 折成静态缓冲的宽串）。当场逐枚数过：全仓只剩「声明 + 定义 + 别人的注释」，调用者 **0**。留着一枚没人叫的导出，下一轮就有人把它当成"这里有个能用的入口"再补一格直译 —— §B124 那份清单的成因正是这一对读者/被读者。recordset 成员面的答案已经唯一，撤它不再需要拍什么。
- **动了什么**：`src/rtl/core/vb6forms/vb6forms_data.c` 撤两枚定义（−13 / +4 —— 留下的四行写清"为什么没了、唯一的路在哪"），`vb6forms_prop_ctrl.h` 撤两行声明（−2 / +2）。仍活的：`vb6_Data_FieldValue`（它是 `vb6_Data_Bind` 与 memberobj 字段那档的后端，不是那两枚的别名）、`vb6_Data_RecordsetObj`、`vb6_Data_Init`。
- **RTL 一刀的老规矩（本轮实测两条，其中一条是自己撞的坑）**：改 `src/rtl` 必须 touch `src/driver/c3rtl.rc` 再重编 C3.exe —— 账 #156 那条：只改 RTL 不 touch，本机增量重编会拿**旧 RTL** 测出一堆看着正常的读数，而 CI 从零重编什么都看不见。读数一：`check_rtl_embedded.ps1` 报 **125 份 RCDATA 逐字节相同、无孤儿 id**（新 exe sha `64345f3f4479`）⇒ 新 RTL 确实进了资源。读数二（**教训**）：那台 touch 一开始**静默没执行** —— 临时 `.ps1` 里留了一行中文注释而那份文件没 BOM ⇒ PS 5.1 按 ANSI 读，中文行尾吃掉换行、把下一行 `$rc = …` 一起并进注释 ⇒ `Get-Item` 拿到 null，打印出来的是一行光秃秃的 `TOUCHED `。⇒ 三条都记着：临时脚本**一律 ASCII**（写完数一遍非 ASCII 字节）、`Write-Output` 的变量要在**打印里也验一次**（空值就是没跑）、事后用 `touch` 补上再让内容级哨兵复核。
- **判据**：`check_data_recordset_shape.ps1` 长出一整头新的（原防空转那格改号 R4）—— 那两枚名字在 `src/rtl` 的**去注释后**代码里各恰好 **0** 次；`//` 与 `/* … */` 都切，且专门钉住"同一行尾部带块注释"那种**从前的真实写法**（负控 N11）。
- **负控 6/6**（全在 `-Root` 副本树上，还原后两份 RTL 文件 md5 逐字节对上）：假定义回来 ⇒ 点名 `vb6forms_data.c:201`；带回尾注的声明回来 ⇒ 仍红；同样的名字只写在 `//` 行里 ⇒ **不许误红**；整户 `src/rtl` 挪走 ⇒ 路径不存在 + `RTL-FILES=0` 两头红；还原 ⇒ 绿。**另把第十五刀那八档在新哨兵上重跑一遍 ⇒ 仍 8/8**：改了判据不许悄悄废掉旧面。
- **证人（真编 = RTL 这一刀唯一会坏的那一环）**：`tests/c29data/DataApp.vbp` 用新编的 C3.exe 真编（先灌 vcvars 进进程环境）⇒ rc=0、`DataApp.exe` 469504 B、日志里 `error C` **0** 条 / `unresolved external` **0** 条 / `fatal error` **0** 条。真跑两架构由门上的 `Test-VbpBuild "c29data"` / `c29data_x86` 负责。
- **护栏（各管各的工件，别拿一个替另一个）**：发码那一头 = 对**现有** `emit-manifest.expected.txt` 复算全语料 ⇒ **期望 398 行 / 实得 398 行 / 相同 398 / 哈希不同 0 / 缺席 0 / 多出 0**，未重钉；这一条只证「cgen 一字未动」（RTL 不写进 emit 文本）。RTL 那一头由两件工件管：`check_rtl_embedded.ps1` （125 份 RCDATA 逐字节相同 ⇒ 新 RTL 真进了 exe）+ 上面那枚真编证人（链接阶段才看得见未定义符号）。
- **§B124 到此收口**：那条「下一刀的形状」四条全清 —— ① 撤 cgen 侧死码 = 第十五刀；② RTL 导出去留 = 本刀（选「去」）；③ 判据 = 一整道哨兵而不是一条结构针（针只钉得住产物那头，而这一族的两份答案住在源码里）；④ 证人 = 门上那套装在 `tests/c29data` 的断言 + 两头 A/B。

### B126 账 #278 第十五刀已出 = C29-Data 的发码侧死码整户撤干净，`Data1.Recordset.<成员>` 从此只有一条路（第 50 道哨兵 `check_data_recordset_shape.ps1`，2026-10-10，门 #459 attempt 1 全绿（run 37999228928、head `fe5f2689`、12/12 全 completed/success、非绿 0、created→updated 10m29s；**新哨兵所在的 Tests (compile) 片绿** ⇒ 那三面判据在 CI 上真跑过，而 **Emit manifest (shape gate) 那一跑同绿** ⇒ 「398 份逐字节相同」被 CI 那台独立复算证实））

- **动了什么**（六份 `src/backend` 文件，−163 / +19 行）：照 §B124 那份**闭合**清单划 —— 五处 `find("vb6_Data_Self(") == 0` 判据（`cgen_expr_call_callee_withm.inc`、`cgen_expr_call_com_bind.inc` 的 `axSlotObj` 并项与那一格直译、`cgen_util_com.cpp`、`stmt/cgen_call.cpp`）、抠句柄的助手 `dataSelfHwndExpr`（`detail/util/cgen_state.inc`）、三处 `vb6_Data_FieldValueStr(` 发射、两处反过来认那枚发射当前缀的识别器（`cgen_util_com.cpp` 与 `cgen_expr_member_generic_access.inc` 两处）、标量四枚直译分支（`BOF`/`EOF`/`RecordCount`/`FieldCount`）。顺带带走那台把日志路径写死在**某个本机用户目录**上的 `getenv("C3_CG_TRACE")` 探针 —— 它只住在被撤的那一格里，留着就是下一台往仓库里写绝对路径的机器。**留着的**：`dataVars_` 与 `dataNameOfExpr` —— `Data1.DatabaseName` / `Data1.BOF` 那一档（控件自身的属性表）还在用，撤了就断活路。
- **口径按 §B124 推荐的 (a) 落地**：`Data1.Recordset.<成员>` 一律由 memberobj 那枚真 IDispatch 应答（`vb6_Data_RecordsetObj` + `vb6_Com*`），发码侧不再有第二种形状。(b)「恢复直译」不作 —— 它要的正是**同一个成员两个调用者**（cgen 与 memberobj 各一条），与本线「一个事实一处答案」相反；而 §B124 排掉的那条嫌疑（走 COM 会把数值成员按字符串读）实测不成立，直译的收益面也就没了依据。
- **判据 = 第 50 道哨兵**：`scripts/check_data_recordset_shape.ps1`（登记为 `tests/run_tests.ps1` 的 `Test-RsRecordsetShape`，`check_static_sentinel_registration.ps1` 复核 census 49→50 全登记）。三面 ——
  - **R1 产物**：新夹具 `tests/ctrlzero/RsForm.frm`，十形 = 方法面五枚（`Refresh` 裸语句 / `Call … MoveFirst` / `MoveLast` / `MoveNext` / `MovePrevious`）+ 读取面四枚（`Fields("id").Value` / `Fields(1)` / `If … BOF` / `RecordCount` 赋给 String）+ `With DataZ.Recordset : .Refresh` 那一形。宿主表达式 `vb6_Data_RecordsetObj(vb6_hwnd_DataZ` 交出**恰好 10 次**；九枚成员名的宽字面量逐条钉数（`Refresh` 2 / 四枚 `Move*` 各 1 / `Fields` 2 / `BOF` 1 / `RecordCount` 1 / `Value` 1）；撤掉的那七枚直译出口在产物里**恰好 0 处**。
  - **R2 源码**：那七枚名字在 `src/backend` 的**非注释行**里再出现成字符串字面量就红（并点名 `文件:行`），`dataSelfHwndExpr` 同名标识符同罪。注释行放行是刻意的 —— §B124 那些「从前这里有一格直译」的历史注释必须活得下去，一条把注释也算作第二份答案的哨兵，下一轮就没人肯把来路写进台账。
  - **R3 防空转**：`src/backend` 扫到的 `.cpp/.inc/.hpp` ≥ 100（实测 122）、产物字节 ≥ 4000、两张名单自己的长度也钉死（7 / 9）—— 名单被清空就是把哨兵改成绿着的空转。
- **负控 8/8 按预期**（全在 `-Root` 副本树上，真树未动）：NC0 副本基线绿；NC1 往撤除名单塞一枚假 needle（`vb6_ComCall`）⇒ R1 DEAD-RETURNED 红并报出「6 枚」；NC2 塞一枚源码里真有的名字（`vb6_Data_RecordsetObj`）⇒ R2 红并点名 `cgen_util_ctrl.cpp:517`；NC3 真改动 —— 把 `"vb6_Data_Self("` 写成一非注释行的字面量 ⇒ 红；**NC3b 同一条串放进注释行 ⇒ 依旧绿**（R2 那条放行面本身也要有判据）；NC4 删夹具的 `With` 那一形 ⇒ SHAPE（10→9）与 MEMBER（`L"Refresh"` 2→1）两头同红；NC5 把 `src/backend` 整户挪走 ⇒ SRC + `SCANNED-FILES=0` 红；NC6 全部还原 ⇒ 绿。两份改过的 `.ps1` 与登记后的 `tests/run_tests.ps1` 各跑 PSParser ⇒ 0 错；本地 50 道静态哨兵逐道跑一遍全绿（`TOTAL=50 RED=0`）。
- **A/B（零行为改动，取两头，因为语料只吃 `.vbp`/`.bas`）**：① 全语料形状门 —— `emit_manifest.ps1` 对 `emit-manifest.expected.txt` **398 份输入逐字节相同**，未重钉；② `.frm` 家族不在语料里，另拿**同一份夹具**两台并排：BASE = 本机冷编的 `wt_base_k15 @ d2caf7c0`（那一版还带着死码），对 `tests/c29data/DataApp.frm`（装在门上真编真跑那枚证人，11965 B）、`tests/ctrlzero/ZeroForm.frm`（12859 B）、新夹具 `tests/ctrlzero/RsForm.frm`（8056 B）三份产物**逐字节相同**。⇒ 撤的确实是一条没人走的路；而 `RsForm.frm` 在 BASE 上发的也是同一套 memberobj —— 这正是 §B124 那句「恒假」的产物侧证据。
- **顺手订正**：`check_rtl_proto_arity.ps1` 头注释里那格「行数恰好 15」订正成 23 —— 代码在第十四刀已改到 23，注释没跟上。这类「注释与判据两个数」迟早骗下一次推理（本轮就是靠它才没把 15 当成新的实数）。
- **§B124 剩下的那一半**（仍未拍板，本刀没碰）：RTL 那枚 `vb6_Data_Self` 导出（`src/rtl/vb6forms_data.c` + `vb6forms_prop_ctrl.h`）现在全仓 0 个调用者 —— 连 memberobj 都不叫它。撤它要动 `src/rtl` ⇒ 按控件线规矩必须 touch `c3rtl.rc` 再重编 C3.exe（账 #156 那条坑：只改 RTL 不 touch rc，探针测的是旧 RTL 且一行 trace 都不打印）。留着它的代价是 RTL 里一枚永不应答的导出，撤它的代价是一次 rc + 全量重编 —— 与本线其它「先拍口径」那几格一起排。

### B132 `SbLayout` 没把内容档面板的实测宽从 `fixedTotal` 里扣 ⇒ 弹簧把同一格宽吃两次，末格右边界越过客户区（账 #303，**已出 —— 第二十三刀，2026-10-10，门 #478/#479**）

第二十二刀把 `Panels(i).Width` 换成"问窗口本人"之后才露面的存量缺陷 —— 从前 getter 交请求值，这一格在任何读数里都不存在（与 §B41 第一格那句"缺的是接线不是解码器"同形）。

- **读数**（`tests/ctrlstatusbar` 真跑，x64 与 x86 逐行相同）：`SB36-RAW before=5880 after=5880 sum=6630 bar=6000` —— 三枚面板宽度之和比控件自身的宽多 **630 缇 = 42 像素**，而 42 正好是那枚 `.Style = 5`（sbrTime）+ `.AutoSize = 2`（sbrContents）面板的测量文本宽。另一份夹具 `tests/sbhm`（两枚固定档 + 一枚弹簧档，没有内容档）里 `HM03-SPRING-TILES` 是 True（`e3 = cw`）⇒ 多出来的那一格只跟着**内容档**出现。
- **成因**（读码对过，不是推断）：`SbLayout` 第一遍只在 `else`（固定档）那一支累加 `fixedTotal`，`VB6_SBR_CONTENTS` 那一支把实测宽放进 `want[i]` 却不累加 ⇒ 第二遍 `remaining = clientW - fixedTotal` 把同一格宽又发给了弹簧 ⇒ `acc` 的末值 > 客户区宽。
- **为什么不是"看着不挤"**：VB6 里状态条各格铺满客户区（`SB_GETPARTS` 的末值就是宽），而 §B41 第二格之后 `Panels(i).Width` / `Left` 交的正是这份 offsets ⇒ 读数**跟着一起偏大**，语料里那四处"按面板宽摆覆盖层"的写法（`pb.Width = sb.Panels(lPan).Width - 45` 那一族）会多摆出一格的量，把时钟那一格压住。
- **判据现成**：`ctrlstatusbar` 的 `SB36-RAW` 已经把 `sum` 与 `bar` 打在同一个窗口里（第二十二刀刻意只钉"增量不变"那一半）⇒ 修好那趟加一条 `SB36-TILES-BAR=True` 即可，证人同源，不用新建夹具。

**第二十三刀的落地与读数（2026-10-10，门 #478/#479）**：
- **动了什么（就一行代码）**：`SbLayout` 的 `VB6_SBR_CONTENTS` 那一支补上 `fixedTotal += want[i];` —— 占位的两档（固定档 + 内容档）各从剩余空间里扣一次。顺带订正一处注释：第一遍上方那句"弹簧面板先按 MinWidth 占位"是**第十八刀之前**的旧形状（那一档早就改成第二遍兜底了，注释一直没跟上）。
- **读数**（真跑，x64 与 x86 两份 run.out 逐行相同）：`SB36-RAW before=5250 after=5250 sum=6000 bar=6000` ⇒ `SB36-TILES-BAR` True；BASE（`.build/b476_base_k22.exe` = 第二十二刀那台，跑**同一份夹具**）`before=5880 after=5880 sum=6630 bar=6000` ⇒ False。弹簧档从 5880 降到 5250 缇，差的正好是那一格时钟的 630 缇 = 42 像素。**另两份夹具一台没动**：`tests/sbhm`（两固定 + 一弹簧，没有内容档）与 `tests/sbfont`（固定 + 内容，没有弹簧）改后输出与第二十二刀那台**逐行相同**，且与 BASE 也相同 ⇒ 这一刀只动"弹簧与内容档同时在场"那一形 —— 而那两份夹具各缺一条腿，正是本账归因时用的两条互证。
- **护栏**：哨兵 `check_statusbar_panel_hm.ps1` 长出 **R8** —— `fixedTotal += want[i];` 在 `vb6forms_statusbar.c` 里必须**恰好 2 次**：0 次 = 本账的形状（内容档不占位），3 次 = 第十八刀撤掉的那个"MinWidth 当加项"复发。两个方向都是真形状，所以钉死数而不是钉下限。**负控 = 第二十二刀那棵树**（`git archive HEAD` 出 `.build/wt_k23neg` + `b476_base_k22.exe`）：一条命令**只报 R8 这一条红**，第二十二刀的 R1/R2/R3/R5/R6/R7 全部仍绿 ⇒ 判据是特定的。51 道 [STATIC] 全跑 = 51/51 绿；改 RTL 已 touch `c3rtl.rc` 重编，`check_rtl_embedded` OK（exe sha `9d77f8558161`）。
- **判据里的容差是留给换 DPI 的，不是留给这台的**：`SB36-TILES-BAR` 写的是 `Abs(sum - bar) <= 45` 缇（3 像素）—— 本机实测 `sum` 与 `bar` **逐字相等**（6000 = 6000），但每格是 px→缇 各折一次回来的，96 DPI 之外会留取整余量，钉死相等会把判据变成这台机器的读数（#171 那条"判据不许钉绝对数"同形）。
- **形状门那一格**：门 #478（run 38036598915、head `83f61e33`、attempt 1）= 11 绿 + 唯一红是形状门，而移动的**只有一行** = `tests/ctrlstatusbar/CtrlStatusBar.vbp`（新增那两条 `Debug.Print` 进了产物）。
  顺带量到一条此前没记过的事实：**`.frm` 里过程体外的注释不进 `--emit-c` 产物** —— 同一趟 `tests/sbfont` 只把那段说"VB 里没有现成的门"的注释换成 ASCII，清单上它那一行**逐字节没动**（第二十二刀那次 sbfont 之所以动，是因为 `*Hm` 那行**发码**真变了）。⇒ 以后"只改过程体外注释"的刀不必预留重登记，也别把注释改动算进门红里。
  登记 = 按 CI 那台自己交的清单（`scripts/rebless_emit_manifest.ps1 -Bless`，numstat 1/1、行序零移动、复算 399/399、`check_manifest_coverage` 绿）。**复验 = 门 #479**（run 38037616826、head `29582214`、attempt 1、12 条 check-run 全 completed/success、非绿 0，其中 `Emit manifest (shape oracle)` 绿 ⇒ 那一行的新哈希被 CI 那台独立复算证实，而 `Tests (vbp #1..#4)` 四片全绿 ⇒ `SB36-TILES-BAR` 在两架构真跑过）。
- **这一趟没验的一格（不立案，只记形状）**：当弹簧档的 `MinWidth` **大于**它分到的空间时，第二遍的兜底照样会把末格顶出客户区（下限赢）—— 本机夹具够不着这个条件（下限折出来 22 缇，实际分到 3600+ 缇），要一枚故意收窄的夹具才测得到，而**那种情形下 VB6 本人交回什么没有读数** ⇒ 别照猜改。

### B135 控件数组元素的几何第一次被问 —— 六格读数两架构齐，BASE 那台逐条翻回量化值（账 #230 第二十七刀，2026-10-10，门 #485（源码那一跑）+ 门 #486（登记那一跑，12/12 全绿））

- **这格为什么开**：`tests/geomcache` 那 31 条钉的是账 #230 的 VB 侧几何缓存，可每一条读的都是**单枚具名控件**；元素那条来路在发码里交的是另一枚窗口（`vb6_CtrlArr_GetAt(&vb6_arr_<名>, i)`），而存储按 (窗口, 属性名) 走 ⇒ 「同一处权威」在原理上成立，但**原理上不是判据**：全门禁从没向数组元素要过几何（#193 那枚 `U-ARREXT` 读的是自己写进去的 15 的倍数，量化差天生看不见）。顺带把 §B61 那格旧标题办掉（订正依据见那一节的复测注）。
- **落地**：新夹具 `tests/geomarr`（一枚 TextBox 数组两枚设计期元素 `arrA(0)/arrA(1)` + 一枚同表单控件 `soloT`，几何值刻意都不是 15 的倍数）。GA01/GA02 钉两枚元素的设计值原样读回，GA03 钉运行期赋值，GA04 钉 `Move`，GA05 钉「写元素 1 不动元素 0」，GA05b 钉两枚是两枚窗口（并把 `.Index` 0/1 打出来），GA06 钉单枚那条路没被带跑；每格配一枚**不经产品换算**的像素证人（user32 的 rect + kernel32 的 MulDiv）。发码两头：`ga_emitc_arr_geometry_authority` 钉元素的四档读写都走那四个具名出口（含单枚那条同表的一起钉），`ga_emitc_arr_geometry_com_prop` 钉元素属性落回 COM 属性访问那一形**不在**。元素的 `.Move` 今天走泛派发 `vb6_ComCall(那一枚, L"Move", …)` —— 属「带实参方法」那一族（§B72 那四张表外的一格），行为面由 GA04 钉，本刀不把它的形状当罪证。
- **为什么单独一枚夹具，而不是往 geomcache 加控件**（本轮最贵的一条读数，别的刀也用得上）：那 31 条里有**五条**的形状是「把所有成员一起读 / 一起写」—— GC17 把每一枚的宽求和、GC18 用 `With collC(i)` 给**每一枚**写 1234、GC20..GC23 按成员问身份。第一版我把数组加进 GCForm，当场读出 `GC32 raw=1007,2501,1234,247` —— 那个 1234 就是 GC18 写进去的，另外 GC13/14/15 的计数范围也被推着走。⇒ 往这类夹具加控件 = 同时重述别人的判据；本刀不办，另起一枚小夹具。
- **负控（这六格能红，且红的就是被并表那一处）**：拿**上一代**（9 月 30 日那台、账 #230 之前）编同一份夹具真跑 ⇒ `GA01=False raw=1005,450,3000,240`、`GA03=False raw=4995,450,7770,3000`、`GA06=False raw=2220` ⇒ 症状正是 §B61 记的那个量化差；同一台 `GA05b=True 但 idx=-1/-1`（`.Index` 那档是账 #252 之后才接上的，两枚夹具的读数顺手把那一格也钉住了）。
- **CI 侧读数**：门 #485 = run 38057764933、head `5cd7c3b4`、attempt 1 = 12 条里 10 绿 + 两条红，而两条红是**同一件事**（`Emit manifest (shape oracle)` 与 `Tests (compile)` 里的 `[STATIC] manifest_coverage` ⇒ 新语料那一行还没登记哈希，其余 399 行逐行相同）。同一跑的两片 vbp 工件里 `GeomArr.out` 各 8 行、**无一条 False**、`GeomArr.err` 都在（= 正常跑完，不是被超时杀 —— 分辨法见 §B133 那一格的 `.err` 对形），读数与本机两台逐字相同。登记那一跑 = **门 #486**（run 38059103972、head `59fc3421`、attempt 1）= 12 条 check-run 全 completed/success、非绿 0，其中 `Emit manifest (shape oracle)` 与 `Tests (compile)` 两格同绿 ⇒ GeomArr 那一行的新哈希被 CI 那台独立复算证实。
- **顺带撞出的独立缺陷**（另开 §B136，本刀一条没碰，且**本格第一版把因由归错了**，订正见那一节的读数）：同一枚表达式 `arrA(0).hwnd <> arrA(1).hwnd` 单独打也读 False，**不是**「并进 And 链才坏」；先 `CLng` 落进声明过的 Long 再比才读 True。

### B136 两枚窗口的句柄直接比 `=` / `<>` 恒答「相等」—— 指针被 `_Generic` 一律投成对象档，而 VarCmp 那族没有数值档（账 #230 第二十八刀，2026-10-11，**已出 —— 门 #487（源码那一跑，11 绿 + 唯一红是形状门）+ 门 #488（登记那一跑，12/12 全绿）**；剩下「句柄装进 Variant 槽」那一格另开 §B137）

- **读数（三形并排，两架构同形；本账第一版把因由归到「并进 And 链」，那是错的，这三条把它钉正）**：① `TF(arrA(0).hwnd <> arrA(1).hwnd)`（**没有** And 链、直接当参数打）⇒ **False**；② 同一格里的 `ok5 = (…) And (arrA(0).hwnd <> arrA(1).hwnd)` ⇒ 整条 **False**（与①同因，链不链无关）；③ 同一对句柄先 `h5a = CLng(arrA(0).hwnd)` 存进**声明过的 Long** 再 `TF(h5a <> h5b)` ⇒ **True**，而且同行的 `.Index` 打出 0/1、`PxLeft/PxW` 两枚窗口证人各报各的数 ⇒ 两枚窗口确实是两枚。⇒ 分岔不在「嵌不嵌套」，在**发码那一步问不问得出操作数的类型**：控件属性读来的 `.hwnd` 没有声明类型可问 ⇒ 掉进装箱那条路，而声明过的变量（Long/LongPtr）走原生数值比较 —— 探针 `.build/b605_lp_probe/LPcmp.bas` 的 `--emit-c` 里，两枚 `Dim As LongPtr` 变量之间同一个 `<>` 发的就是 `a1 = (-(h1 != h2));`（那条探针**只取发码形状，没跑成**，构建红在另一格上，别当运行期读数用）。
- **机理（两格咬在一起，读码定位）**：`vb6_VariantFromValue` 的 `_Generic` 把**任何指针**都投给 `vb6_VariantObject`（那张表自己最后一行写着「其余指针 (void*/class*/type*) → VariantObject」，`src/rtl/core/vb6rtl/vb6rtl_variant.h:176-200`）；而 `vb6_VarToDouble_internal` 只认 VT_I2/I4/R4/R8/BOOL/BSTR，`default: return 0.0`（`src/rtl/core/vb6rtl/vb6rtl.c:604-619`）⇒ 两枚对象档都折成 0.0 ⇒ Eq/Ne/Lt/Gt/Le/Ge 六条一起塌。发码现场就是 `vb6_VARIANT _vcmp_0 = vb6_VariantFromValue(vb6_GetControlHwnd(…))` 配 `vb6_VarCmpNe(&_vcmp_0, &_vcmp_1)`。
- **语料暴露 ≈ 0**（先量了再定性：机制坏 ≠ 产品在坏）：CI 工件里两枚真工程的 emit —— Charts ucChartArea 4 条 VarCmp 行、VBFlexGridDemo 159 条，含 `hwnd` 的只有 2 条，而那两条走的是 `vb6_VarCmpLong*` 数值档（不经过坏的那一步）；自家 modal / tabwalk 夹具比句柄**一律**写成 `HexEq(f, x.hwnd)` 字符串比（37 处）⇒ 今天没有一条真流量踩在上面。
- **修法要先定一档**（这格的关键）：VB6 里 `LongPtr` 的比较是**数值**比较，对象引用的 `=` 走**默认属性**、`Is` 走**身份** ⇒ 缺口在「装箱档按什么算」——`.hwnd` 的值类型是 LongPtr，不该进 VariantObject。**别在 `vb6_VarToDouble_internal` 里给 VT_DISPATCH 补一档就当修完**：那等于顺手把「两枚对象引用相等」也改成按指针比，是另一个口径。这一问与账 #228（同一 VB 类型两种拼法给出两种 C 类型）、账 #256（装箱档不跟值类型走）同族，开工前先对齐那两格的答案。
- **下一刀的形状（2026-10-10 量完，两条都不必再猜）**：
  ① **哪些名字中招** —— 伪成员表（`src/common/form_pseudo.hpp` + `host_pseudo.hpp`，共 38 名）里，
     `controlPropType` 答不上类型的有 **24 名**；按 RTL 声明的 C 返回型过一遍（`.build/b623_pseudo_types.py`
     把 24 行逐条打出来），真正是**指针/句柄**的只有 `hwnd`(`void*`) 与 `hdc`(`intptr_t`) 两枚
     （`mouseicon` 交的是 `void*` 但那本来就是 IPicture 对象档、`caption` 是 `wchar_t*` = BSTR，字符串那族
     `vb6_VarCmp*` 自己会认 VT_BSTR ⇒ 不中招）。⇒ 本刀的范围就是这两枚，不是补一张 24 行的名单。
  ② **该问哪张表** —— 这不是「两格机制」要并表的那种重复：`form_pseudo.hpp` 那张表只管**发码调哪个 RTL 符号**，
     `controlPropType` 只管**这枚属性是什么 VB 类型**（账 #231 写死的分工），而且**先例已经在原地** ——
     `left`/`top`/`width`/`height` 四行同时住在两张表里（伪成员表给 getter、类型表给 Long），
     没人觉得那是一格该并的重复。⇒ 落地就是给 `controlPropType` 补 `hwnd`/`hdc` 两行答 `LongPtr`
     （与 RTL 的 `void*`/`intptr_t` 同宽，账 #180 早就定了「句柄按指针宽度」这一档），
     哨兵 `check_ctrl_prop_type_authority.ps1` 的 A4 名单从 28 条跟着长，别在 `vb6_VarToDouble_internal` 里动手（见上格那条「别把对象引用也改了」）。
  ③ **风险面已点名**：改的是**类型答案**，所以会牵动打印/装箱/拼接，不只是比较 —— 自家 `tests/modal` 与
     `tests/tabwalk` 判句柄用的正是 `HexEq(f, x.hwnd)`（把句柄转成十六进制串再比），这两枚夹具就是这刀的
     现成反面证人：若 `hwnd` 打成 LongPtr 之后那套串的形状变了，它们当场红。护栏照旧 = 全语料 A/B
     （BASE 用 CI 那台 `c3-exe` 工件即可，不必冷编）+ 改动行逐份归因 + 两架构真跑。

  ④ **这条形状与源码里一条既有的刻意决定相撞，动手之前必须先翻它** —— `src/backend/cgen_util_ctrl.cpp:79-81`（账 C29-SL-e 的收尾）写着「同一段里 `MouseIcon` / `Hwnd` 的 getter 也是 `void*`，但那两条**本来就是对象**，装箱成对象是对的 … 刻意不动」。VB6 那边 `MouseIcon` 确实是 `StdPicture` 对象，而 `hWnd` 是**句柄数值**（`TypeName(Form1.hWnd)` 在 VB6 里答 "Long"，不是 "Object"）⇒ 那条决定把两件事当了一件。落地要么明确翻掉它、要么在类型表里把「句柄」与「对象」分档，别绕过注释默默改。
  ⑤ **动之前只有一件事要先量**：今天有没有什么东西真吃着 `hwnd` 的**对象装箱** —— 第一个要问的就是`tests/modal` / `tests/tabwalk` 里那 37 处 `HexEq(f, x.hwnd)`（实参位的打包档归谁说话 = §B85/账 #256那条「装箱档跟不跟上一步真正交出的值」在这儿是活的）。若那两枚夹具换了档还能绿，就说明对象装箱没人吃；若红，就先按「两形两答」那族记下答案，再定这一刀的形状。
- **判据面**：`tests/geomarr` 的 GA05 / GA05b / GA05c 三行已经是这格的现成证人（②/①/③ 三形齐；GA05c 在 `.build/b606_inline/` 那份**带这一形的副本**里量到，正式夹具今天只有 ②/③ 两行 ⇒ 开工时把 ① 也升进门禁，并把两枚发码形状各钉一针：属性读那一形**不许**再发 `vb6_VarCmp*(vb6_VariantFromValue(vb6_GetControlHwnd(...)))`，而 ③ 那形的原生数值形状不许漂走）。

**第二十八刀的落地（把上面 ①②③④⑤ 五问一次办完；产品改两处、夹具改两枚）**：

- **改的两处**：`controlPropType`（`src/backend/cgen_util_ctrl.cpp`）补两行 —— `p == "hwnd" || p == "hdc"` 答 `Vb6Type::LongPtr`（与 RTL 的 `void*` / `intptr_t` 同宽，账 #180 早定了「句柄按指针宽度」那一档）；`cgen_expr_call_conv_cstr.inc` 里账 #180 那条 LongPtr 分支加一道闸：先问现成的 `cExprIsVariantCarrier(实参)`，已是 VARIANT 就**什么都不转**、交给默认那条 `vb6_CStr` 自己按档案拆。哨兵 `check_ctrl_prop_type_authority.ps1` 的 A4 名单 28→30 条（**是名单跟着长，不是又开一张表**）。
- **④ 那条刻意决定已明确翻掉**：原注释把 `MouseIcon` 与 `Hwnd` 并成「本来就是对象、装箱成对象是对的」。现在 `MouseIcon` 留在对象档（VB6 那边它确实是 `StdPicture`），`hWnd` / `hDC` 改答句柄数值 —— VB6 里 `TypeName(Form1.hWnd)` 答 "Long" 不是 "Object"。注释改的是**结论**，留着「当初为什么并成一件」，不是抹掉。
- **本刀撞出来的一条，比修本身贵**：「类型答 LongPtr」**不等于**「实参在 C 侧是标量」。属性读没登记具名出口时会落 COM 兜底（`soloT.hDC` 就是这一形），交出的表达式是 `vb6_VariantFromComResult(vb6_ComGetProp(…))` = 一个 `vb6_VARIANT` 结构体 ⇒ 账 #180 那条 `(int64_t)(intptr_t)(…)` 裸转的产物是 **C2440 + C2198**（压根编不过，不是运行期才错）。**闸摘掉真编译过一次**：`NEW-BUILD-RC=1`（两处报错都在 `GAForm.c:299`），装上 `rc=0`。同一族在比较面早有解（Fix 158m「C 表达式确实是 VARIANT 就升级成 Variant 档」）、在 CStr 面也有（同一文件下面那条 Fix 225）⇒ 这一刀只是把**那一句现成的问话搬到 LongPtr 那一档**，不新开名单。⇒ 以后动「类型答案」这种单一权威，凡下游有「按类型裸转 C」的分支，都要问一遍实参的 C 形状；这一族的红藏在**编译期**，不在判据里。
- **读数（两架构真跑 + 同一份夹具在 BASE 那台对跑，`.build/b638_k28/report.txt`）**：GA05c 三形 —— NEW 读 `inline=True chain=True eqself=True`，BASE 读 `inline=False chain=False eqself=True`（⑤ 那一问的答案就在 `eqself`：BASE 上它也是 True，因为两边都折成 0.0 ⇒ 反面证人吃的是「不许恒 False」，而「不许恒 True」靠 BASE 那侧的 False 与 `eqself` 两头一起夹住）。GA05d 两条 stringify 面 —— `"H" & arrA(0).hwnd`：NEW 交回 `H5442058`、BASE 只交 `H`（`catlen` 1→8，**改前打的是空串**，与 §B136 上面那条「CStr(句柄) 打成空」同源）；`CStr(soloT.hDC)` 两台都非空（那一格正是靠上面那道闸保住的，不是靠换档）。x86 与 x64 判据行同号；x64 那台还顺手把 BASE 一个**负数**读数（把 64 位 DC 折成 int32 打出的 `-419357591`）改回完整句柄数 ⇒ 这一格在 x64 才现形。
- **判据**：`tests/geomarr` 长出 GA05c / GA05d 两行（`$gaExpected` 两条针，两架构都吃）+ 两枚发码形状用例 —— `ga_emitc_handle_value_compare`（原生 `-(vb6_GetControlHwnd(…) != vb6_GetControlHwnd(…))` 在，`vb6_CStr(vb6_VariantFromValue(vb6_VariantFromComResult(…L"hDC"…)))` 在）与 `ga_emitc_handle_not_object_box`（`vb6_VariantFromValue(vb6_GetControlHwnd(…))` 那一形必须绝迹）。三根针都是**先从实际产物里取**再钉：出现次数 NEW 2/1/0、BASE 0/1/7，括号配平与「不含 `[ ] * ?`」（`-like` 会把它们当字符类）逐条核过；`run_tests.ps1` 改完 PSParser 0 错。
- **顺带办的第二枚夹具**：`tests/ctrlstatusbar` 的 SB0 原来钉的是「`Me.hwnd` 在 Debug.Print 里是空串」那格**已知局限**（runner 的注释就那么写着，期望串 `SB0-HWND= CAP=CtrlStatusBar CL=` 里那个「等号后直接一个空格」就是在钉空）。这一刀之后它不再是局限 ⇒ 期望重述成 `SB0-HWND-NUM=True CAP=CtrlStatusBar CL=`（吃 `SBTF` 那个布尔，**不吃会抖的句柄值**），注释跟着订正。BASE 那台读 `SB0-HWND-NUM=False` ⇒ 换了读法这一格仍然能红，红的就是被修那一处；`CL=` 后面依然什么都不钉（Controls 那一头归 #250 那一族）。
- **护栏（语料 A/B，本机冷编两台、同一工具链）**：BASE = 把 HEAD（`5ccca69c`）`git archive` 到 `.build/b632_base_k28` 冷编那台（`Release/C3.exe`，`md5 4401bc5ead3e06b850faae50b82da99d`），NEW = `.build/wt_vb/.build/C3.exe`（`md5 e450afd55d21fab5c45b2880170117e1`）⇒ **inputs=400 / same=399 / changed=1 / rc_mismatch=0**（`.build/b636_ab_corpus.txt`）。唯一那份是 `tests/ctrlstatusbar` 的 SB0 那一行，而变的正是被修的那一面（BASE 发 `vb6_CStr(vb6_VariantFromValue(…))`、NEW 发 `vb6_CStrLongLong((int64_t)…)`）⇒ 这刀在**产品**面的暴露就这一处。邻居三枚真跑逐行对形（`geomcache` 32 行 / `modal` 20 行 / `tabwalk` 18 行）：差异只在**已在册的抖动族**（`GC20/GC23/GC26` 的 `raw=` 句柄尾数、`GC16` 的窗位、`M1` 那对句柄、`EX-fr1/EX-f2` 的句柄），判据行逐条同号；`tabwalk` 那两枚 `HexEq` 仍 `eq=Y` / `eqDeep=Y` ⇒ ⑤ 要的那个读数到手：**对象装箱那档没人吃**，换档不会把它们带走。51 道 [STATIC] 全绿（含账 #238 那枚 `check_variant_cmp_boxing` 与账 #231 的 A1..A7），零红。
- **登记（形状门，第七次走这条自救路）**：门 #487（run 38067894261、head `48e58641`、attempt 1）= 12 条里 11 绿 + 唯一红 `Emit manifest (shape oracle)`，红因是两行还没按 CI 的读数重登记（`ctrlstatusbar` = 发码 + 夹具文本、`geomarr` = 夹具文本），其余 398 行逐字节没动。按 CI 那台自己交的清单登记（`github/ci/emit-manifest`，提交 `1496cac7`、注脚 `head=48e58641`），登记后复算 400/400、`K1..K3` 绿；**另外用本机新编那台把整份清单跑了一遍与 CI 对形：400/400 逐行相同** ⇒ 这一刀的产物没有第二个答案。
- **门**：源码那一跑 = **#487**（run 38067894261、head `48e58641`、attempt 1）= 11 绿 + 唯一红形状门（不是产品红，是没登记）；登记那一跑 = **#488**（run 38069065264、head `9b76b2e3`、attempt 1）= 12 条 job 全 completed / success、非绿 0（`Emit manifest (shape oracle)` 同绿 = 那两行的新哈希被 CI 那台独立复算证实）。

### B137 控件句柄**装进 Variant 槽**那一路吃类型答案 —— 装箱改按指针宽度走，比较面补上 `VarCmp` 缺的那一档 `VT_I8`（账 #230 第二十九刀，2026-10-11，**已出 —— 门 #489（源码那一跑，12 条里 11 绿 + 唯一红是形状门）+ 门 #490（登记那一跑）**；顺带量出来的那一格另开 §B138）

- **读数（两台同形 ⇒ 这不是第二十八刀带来的，是它没 cover 到的那一格）**：探针 `.build/b633_probe/GAForm.frm` 里 `Dim v7 As Variant: v7 = arrA(0).hwnd`，随后 `VarType(v7)` = **9**（VT_DISPATCH）、`CStr(v7)` = **空串**、`CLng(v7)` = **0** —— 改前（BASE `5ccca69c` 冷编）与改后（第二十八刀）两台**一模一样**。读数在 `.build/b635_new64_run.out` 与 `.build/b635_base64_run.out` 的 `P-v7type` 那一行。VB6 的答复是 `VarType(hWnd)` = 3（vbLong）、`CStr` 打数字。
- **机理与 §B136 同一条，只是换了消费面**：装箱进 Variant 槽那条路（`boxToVariant`，`src/backend/cgen_util_ctrl.cpp:2443` 起）按 VB 类型只显式改写 Boolean/Byte/Single/Date 四档，其余一律交 `vb6_VariantFromValue` 由 `_Generic` 选 ctor ⇒ 指针（`void*` / `intptr_t`）落那张表的最后一行 `VariantObject`。**它没吃 `controlPropType` 的答案**（那把答案在比较面与 stringify 面已经吃上了 —— 就是第二十八刀）。
- **为什么这一格不只是「打出来不好看」—— 以及本账原先那句风险陈述怎么被读数换掉**（2026-10-11 第二十九刀订正）：那一档确实是个**假对象**（把 HWND 当 `IDispatch*` 装进 `VT_DISPATCH` 槽），但它**不是野调用**：`vb6_ComAddRefDispatch` 里面的 `vb6_ComIsDispatchable` 有一条 `if (IsWindow((HWND)disp)) return 0;`（`src/rtl/core/vb6com/vb6com_invoke.c:48`，那段注释就是为宿主 `Controls` 枚举塞裸 HWND 这一形写的），AddRef 与 Release 两头都被它挡住 ⇒ 本账上一版写的「任何真去 AddRef/Release 它的路都是发虚调用」**不成立**，代价全在**值面**：`VarType` 答 9、`TypeName` 借 `vb6_Host_TypeNameOf` 答出**控件名**「TextBox」（不是 "Object"，所以比 §B136 那条更像对象、更容易被当成对象用）、`CStr` 交空串、`CLng` 答 0、与任何数比较恒不等。同族的关系仍然成立，只是罪名换了：§B85（账 #256）那一格交出的是**真** VARIANT 的野 Release，这一格交出的是一档**看着像对象、值面全错**的假档案。
- **暴露已经量过：这一格是潜伏项，不是现场失血**（与 §B134 同型 —— 机制坏，而语料不喂）。全语料 400 份输入的 `--emit-c` 里，`vb6_VariantFromValue(vb6_GetControlHwnd(` / `…Hdc(` 这一形命中 **0 处**（脚本 `.build/b648_box_census.py`、报告 `.build/b648_box_census.txt`；同一条正则在带 `v7 = arrA(0).hwnd` 那枚探针的产物 `.build/b635_emit_new.c` 上命中 2 处 ⇒ 不是正则空跑）。而**全语料 A/B 也说不到它**：两台在这一格发码相同，A/B 只看得见「改前改后有没有差」⇒ 这类格子的优先级只能按口径拍，不能按流量拍。剩下的那一件开工前量：`boxToVariant` 补 LongPtr 一档会不会把**真对象**（C 层也是 `void*`）一起改掉 —— 答案仍在「类型是否已知」这一问上（账 #180 早就因此不肯给 `_Generic` 补 `void*` 档）。
- **口径要先拍的那一档**：句柄进 Variant 槽该答 VT_I8 / VT_INT（对齐 VB6 那个 Long）还是保持对象档。**别在 `vb6_VariantFromValue` 的 `_Generic` 里给 `void*` 加数值档** —— 真对象引用在 C 层也是 `void*`，那会把所有对象一律打成数字（比 §B136 那一格宽得多，装箱那条路是全场最常走的一条）。
  **这一档已由第二十九刀拍掉**：按**指针宽度** —— `(intptr_t)` 交 `_Generic` 自己选，x86 落 VT_I4、x64 落 VT_I8（与账 #180「句柄按指针宽度」以及用户对「句柄的本质也是结构体指针，vb6 用 long(x86) 表示指针没问题」那条口径一致）；「保持对象档」那一支被读数否掉（值面全错）。「别在 `_Generic` 里给 `void*` 加数值档」这条约束照办 —— 新档只加在 `boxToVariant` 那一处、按 **VB 类型** 分诊，`_Generic` 那张表一个字没动。


**第二十九刀的落地（两条消费面一起接上那张类型表）**：

- **改的两处**：`boxToVariant`（`src/backend/cgen_util_ctrl.cpp:2443` 起）在 Boolean/Byte/Single/Date 之后补第五档 —— `case Vb6Type::LongPtr:` 交 `vb6_VariantFromValue((intptr_t)(…))`，让 `_Generic` 按**指针宽度**自己选档（x86 = `long` → VT_I4，正是 VB6 对 hWnd 的答案；x64 = `long long` → VT_I8，不截断，与 Fix 082 的 VarPtr 档同族）；RTL 的 `vb6_VarToDouble_internal`（`src/rtl/core/vb6rtl/vb6rtl.c:604`，`vb6_VarCmp*` 全族的取数口）补 `case VT_I8`。**函数头那句「需要显式档案的四种」跟着改成五种**，不是新开一张表。
- **RTL 那一刀不是为本刀新档打的补丁，它自己就是一格现场失血**：`VarCmp` 取数口没有 VT_I8 档 ⇒ 任何 64 位整数档案在比较面被抹成 0.0。读数不靠本刀的新档：**声明过的 LongPtr 今天已经会装成 VT_I8**（`_Generic` 的 `long long` 一支），所以 x64 上 `VarPtr(x) <> 0` 实测答 **False** —— 改前改后两台同形（BASE `5ccca69c` 也答 False），只是取数那一侧（`VariantToLong/ToBoolean/ToDouble/ToIntPtr`）早就有 VT_I8，偏偏比较这一侧漏了。⇒ 这一刀把 §B136 那条「恒相等」的家族在 Variant 槽这一面也端掉了。
- **第二十八刀那条护栏在这里第二次用上，而且更省事**：`boxToVariant` 头部本来就有两道前返 —— 裸 COM 结果走 `vb6_VariantFromComResult`，以及 `cExprIsVariantCarrier(cExpr)` 为真就恒等直传。所以「类型答 LongPtr 但 C 侧交回来是个 `vb6_VARIANT` 结构体」那一形（属性读没登记具名出口时落 COM 兜底）压根进不了新档 ⇒ 不会重演第二十八刀那次 C2440。这一条是读码 + 装上真编过的，不是推的。
- **读数（两架构真跑 + 同一份夹具在 BASE 那台对跑，`.build/b663_ga05e/`）**：GA05e 两行 —— NEW x86/x64 都读 `GA05e-notobj=True str=True num=True eq=True ne-other=True` + `GA05e-vp-nz=True obj=True`；BASE 那台两架构都读 `notobj=False str=False num=False eq=False`（装箱面），而 `vp-nz` **只在 x64 红**（BASE x64=False、BASE x86=True）—— 因为 x86 那条地址本来就是 VT_I4。这一记不对称正是 RTL 那一刀的靶子：同一条期望串两架构都吃，两台都必须能红。反面证人两头都在：`Set o8 = soloT` 在四台（含 BASE 两台）都仍答 `9`/「TextBox」⇒ 换档没把真对象带走。探针那台另读到 `P-typename` 在 x86 从「TextBox/9」变成 VB6 的答案 **Long/3**，x64 变 **Variant/20**（名字面与 VB6 分叉，另记 §B138 之后）。
- **判据**：`tests/geomarr` 长出 GA05e 两行（`$gaExpected` 两条针，两架构都吃）+ 一枚发码形状针 `ga_emitc_handle_variant_box`（钉的是实测产物：`vb6_VariantFromValue((intptr_t)(vb6_GetControlHwnd(vb6_CtrlArr_GetAt(&vb6_arr_arrA, 0))));` 与 `vb6_VariantFromValue((intptr_t)(pv12));`，出现次数 NEW 2 / BASE 0）；反面证人复用第二十八刀那枚 `ga_emitc_handle_not_object_box`（裸的 `VariantFromValue(句柄读)` 必须绝迹），它从这一刀起**同时管比较面与装箱面**。`run_tests.ps1` PSParser 0 错（TOKENS 31210 / ERRORS 0），文件仍是 BOM+CRLF、0 条裸 LF。
- **哨兵跟着长一条 A8（不是新开一道）**：`check_ctrl_prop_type_authority.ps1` 的 thesis 就是「那张类型表的答案必须被消费掉」，A8 钉三件 —— a) `boxToVariant` 里必须有 `case Vb6Type::LongPtr:` 且经 `(intptr_t)`；b) 这个数值装箱在全 `src/` 里**只许一处发射**（实测 census = `cgen_util_ctrl.cpp=1`，第二处就是又开一条平行装箱路）；c) RTL 的 `vb6_VarToDouble_internal` 体里必须有 `case VT_I8`。三条都过了**能红这一关**：把档名换掉 / 把 `VT_I8` 换掉，同一条正则立刻不命中（`.build/b655_a8_check.py`）。
- **护栏（语料 A/B，本机冷编两台、同一工具链；报告 `.build/b662_ab_corpus.txt`）**：BASE = `5ccca69c` 冷编那台，NEW = 本刀新编那台（md5 `4d5ef027…`）。**inputs=400 / same=397 / changed=3 / rc_mismatch=0**，三份逐条归因：两份是**本刀自己动的夹具**（`ctrlstatusbar` 的两行 = 第二十八刀那次 SB0 重述在 BASE 上仍读 False；`geomarr` 的 18 行 = GA05e 新增 + 那两行的产物），一份是**产品面**：`tests/VBFlexGridDemo` 的 2 行 —— `VBFlexGrid.ctl:10383 Public Property Get ColImageList(...) As Variant`（手册原文「returns the image list control **or handle**」）里 `ColImageList = .Handle` 那一处装箱换成了数值档；那枚工程在门里是真跑的（`Test-GuiVbp` 的「FlexGridX86」那一枚）。另：语料里 `vb6_VariantFromValue(vb6_GetControl` 那一形 BASE 有 9 份命中、NEW 归零的那部分只在夹具内 —— 其余 8 份是别的指针档（类型表不认识它们），本刀刻意不动。
- **邻居真跑逐行对形（`.build/b659_k29/report.txt`）**：`geomcache` 32 行 / `modal` 20 行 / `tabwalk` 18 行 / `ctrlstatusbar` 44 行，差异 8 / 2 / 4 / 2 行，全部落在**已在册的抖动族**（`GC20/GC23/GC26` 的 `raw=` 句柄尾数、`GC16` 的窗位、`M1` 那对句柄、`EX-fr1/EX-f2` 的句柄、SB0 那一行本身），判据行逐条同号；`tabwalk` 那两枚 `HexEq` 仍 `eq=Y` / `eqDeep=Y`。51 道 [STATIC] 全绿（含新的 A8），零红；RTL 动过 ⇒ `src/driver/c3rtl.rc` 已 touch 重编，`check_rtl_embedded.ps1` 报 125 份资源与磁盘逐字节相同。
- **形状门（第八次走这条自救路）**：门 #489（run 38073869124、head `3920dd83`、attempt 1）= 12 条里 11 绿 + 唯一红 `Emit manifest (shape oracle)`，红因是两行还没登记（`geomarr` = 发码 + 夹具文本、`VBFlexGridDemo` = 发码），其余 398 行逐字节没动。按 CI 那台自己交的清单登记（`github/ci/emit-manifest`，提交 `db9c1e01`，注脚 `head=3920dd83`），登记后本机复算 400/400、`K1..K3` 绿；**另用本机新编那台跑了整份清单对形：与 CI 那份 400/400 逐行相同** ⇒ 这一刀的产物没有第二个答案。

### B138 `VarPtr` / `StrPtr` / `ObjPtr` 的**类型答案**是 Long，而它们的 C 侧早在 Fix 082 就成了 `intptr_t` —— x64 一步式进槽把地址折成 32 位（2026-10-11 量；**这一格已由别的分支出了 —— `bebd2083`「VarPtr/StrPtr/ObjPtr 的类型答案改 LongPtr」随 PR #9 合进 dev（合并提交 `226f95c6`），本节的 ① 那条口径就此拍成「答 LongPtr**」；语料两条消费面各数一遍 = 0 处暴露，与 §B134 同型的潜伏项）

- **读数（两台同形，与第二十九刀无关）**：探针 `.build/b633_probe/GAForm.frm` 的 `P-v13-step` 那一行，x64 NEW 读 `types=3/20 strs=-1285887440/1016621361712` —— 左边是**一步式** `v13 = VarPtr(h5a)`（发码 `v13 = vb6_VariantLong((intptr_t)&(h5a));` ⇒ VT_I4，高 32 位没了，还打成负数），右边是**两步式** `pv13 = VarPtr(h5a)` 再进槽（发码 `vb6_VariantFromValue((intptr_t)(pv13))` ⇒ VT_I8，全址）。BASE x64 同形（`-1783634224/298864076496`）。同一条问句（「这枚指针宽度值进 Variant 槽该答什么」）两个答案 —— 而第二十九刀之后这两个答案分属**两张表**：语义层的内置表 vs 类型那一问。
- **机理那一处就在内置表**：`src/semantics/builtin/builtin_funcs.inc:292-294` 三行 `addBuiltinFunc("VarPtr"|"StrPtr"|"ObjPtr", Vb6Type::Long)`，而 `cgen_expr_call_builtin_pre.inc:109` 起那条 Fix 082 把发的 C 表达式做成 `intptr_t`。**别把这一格当成第二十九刀没修干净**：`boxToVariant` 吃的是 `inferExprType` 的答案，答案给 Long ⇒ 走 `vb6_VariantLong` 一支，与本刀的 LongPtr 档无关。
- **⚠ 那一格里 `TF(v13 = pv13)` 读到的 True 不是覆盖**：发码是 `vb6_VarCmpLongEq(&v13, pv13)`，而这族 helper 的签名是 `int32_t vb6_VarCmpLongEq(vb6_VARIANT*, int32_t)`（`vb6rtl.c:678`）⇒ 右操作数**也**被折成 32 位，两边一起截断才「相等」。⇒ 第三件已量完（读码）：同一个问句「Variant 与标量怎么比」有 **6 个发码点**（`cgen_expr_binary.cpp:445、448` 显式 `(int32_t)(…)`；`571、582、604、614` 交 `scalarArg158m(…)` 裸传、由形参隐式折），全都汇进同一个 `int32_t` 槽。
- **暴露两面都数过（`.build/b678_b138_exposure.py`，两架构各一遍）**：书写面 9 份语料文件 / 233 个调用点（VarPtr 114、StrPtr 94、ObjPtr 25；大头 `VBFlexGridDemo` 的 VTableHandle.bas 80 / Common.bas 74 / VisualStyles.bas 37）；**产物面两步都是 0** —— 一步式截断装箱在 12 份相关输入的 emit 里 0 处；`vb6_VarCmpLong*` 那一族虽然多（VBFlexGridDemo 一份 154 处 / 131 行），标量侧全是 `1`、`Col`、`Size.CX`、`LBoundCols` 这类 Long，指针宽度那侧 0 处（最接近的一条是 `GetWindowThreadProcessId(...)`，返回 DWORD，折了不丢）。⇒ **今天只会红在纸面上**。
- **开工前要拍的两件**：① 口径 —— 这三枚内置在 x64 该答 LongPtr（对齐账 #180「句柄按指针宽度」与第二十九刀那一档）还是守住 VB6 的 Long；拍之前先想清楚改的是**内置表那一行**，它会同时改装箱面、比较面、实参打包面（第二十九刀的 A8 只管到控件属性那一头）。② 护栏 —— 改类型答案要把每个「按类型裸转 C」的分支问一遍实参的 C 形状（`cExprIsVariantCarrier`，§B136/§B137 两次都是这一条）。顺带还有两格同日量到、都归这类：**`TypeName` 对 VT_I8 答「Variant」而 VB6 答「Long」**（x86 这一刀之后就对，x64 不对，`vb6rtl_conv.c:377` 那张 switch 没有 VT_I8 一档）；以及 `vb6_VarCmpLong*` 那 6 处发射该收成一处问话。
- **【2026-10-11 复核订正】这一格不在本线上做了：核实靠的是 dev 那一份源码树，不是提交信息** —— `git show github/dev:src/semantics/builtin/builtin_funcs.inc` 的 **299-301** 行现在写的是 `addBuiltinFunc("VarPtr"|"StrPtr"|"ObjPtr", Vb6Type::LongPtr)`（293-294 就是这条改动的说明：这三枚交的是地址、宽度随架构，而 `Vb6Type::Long` 在 C 侧是 32 位）；本机 vbe-line 那一份同一处仍是 `Vb6Type::Long)`（292-294）⇒ ① 那条口径由他们拍成「答 LongPtr」（与账 #180「句柄按指针宽度」、第二十九刀那一档同向），② 的护栏跟着他们那一跑的门走。本节尾巴上那两格**没被这一改带走**，仍归本线：`TypeName` 对 VT_I8 答「Variant」而 VB6 答「Long」（`vb6rtl_conv.c:377` 那张 switch 缺 VT_I8 一档，x64 不对）+ `vb6_VarCmpLong*` 那 6 处发射收成一处问话。
- **【2026-10-11 第三十二刀，账 #305】上面那两格的第一格已出：`TypeName` 的 VT_I8 档补上了，顺手把同一条决定的另外三处漏答一起收**（读数是这一刀的探针打出来的，不是推的）：同一枚 `v = VarPtr(hx)` 在 x64 装 VT_I8(20)、在 x86 装 VT_I4(3)，而 RTL 把「这枚 Variant 算不算数值 / 该报什么类型名 / 打包成什么档」以 `switch (v.vt)` 的形状实现了 **11 遍** —— §B138 那一改（内置答案换 LongPtr）之后 x64 第一次真往这些表里喂 I8，当场四处答错：`vb6_TypeName` 落 default ⇒ "Variant"、`vb6_IsNumeric` ⇒ 假、`vb6_ComPackVariant` 落 default ⇒ **地址静默打包成 VT_EMPTY**（这一处比错名字严重，它是丢数据不报错）、`vb6_ho_variantToLong/ToDouble` ⇒ 读回 0。修法 = 四处各补 `case VT_I8`（档名沿用 SDK 的 `VT_I8`，与那五处早就答了的写法一字同），**判据两头**：`tests/test_variant_boxing.bas`（就是 §B137 那枚装箱哨兵）加五条针 `VB-ptr-tn=Long / -isnum=True / -cbool=True / -selfeq=True / VB-str-tn=Long`，并**新register 一枚 x86 双胞胎** `test_variant_boxing_x86` —— 同一个 `$boxNeedles` 两架构逐字相同，28/28 命中；改之前那一跑（`.build/b944_i8run`，同一支探针）x64 读的是 `typename=Variant / isnumeric=False`。第 53 道哨兵 `check_variant_i8_coverage.ps1` 钉的是**闭合 census**（V1 站点数下限 / V2 每张 vt 表必须在 must 或 exempt 里 / V3 must 里每个函数体必须真有那一条 case / V4 名单里每个名字必须还存在 / V5 名单不许缩），六发假改动各红各的（`.build/b954_negctl_i8_report.txt`：V1..V5 全红过、baseline rc=0）。`vb6_VarCmpLong*` 那 6 处并表**没做**，仍是本账的第二格（见 §B141 与任务）。【2026-10-11 收线】源码那一跑 = **门 #501**（run 38088742545、head `e51fe008`、attempt 1）= 12 条里 **11 绿 + 唯一红是 `Emit manifest (shape oracle)`**；那枚新加的 x86 双胞胎 `test_variant_boxing_x86` 第一次进门禁就在绿的里 —— `Tests (bas #1)`/`Tests (bas #2)` 两片全 green，与 x64 同针同过（28/28 那条针面在两台都命中）。红的 8 行**不是本刀带来的**：按 `ci/emit-manifest` 逐 head 二分，7 行（`tests/Charts 2020` 的 6 份工程 + `tests/test_varptr.bas`）在 `e589b1a3`（§B138b：LongPtr 的位宽按目标架构取档）那一跑第一次变，而 `bebd2083`（§B138 只换类型答案）那一跑这 8 行一字未动 ⇒ **换答案不改发码，加窄化检查才改**；第 8 行 = `tests/test_variant_boxing.bas`，是本刀新增那 5 行读数的源文本进了产物（唯一一处出自本刀）。每行增量都是 13 的整数倍（+442=34×13、+65=5×13 那一档四份、+26=2×13、+13=1×13）= `vb6_ChkLong(` 那一层壳的长度，与本刀的 RTL 四张表无关（RTL 不进发码）。其余 392 行逐字节未动、缺席 0 / 多出 0、rc 两侧全 0。登记按 §B100 那条自救路径取 CI 自己那份（`github/ci/emit-manifest` 的 `ee9ccb9b`），复算 400/400、`check_manifest_coverage` K1..K3 绿，登记跑 = **门 #503**（run 38090841379、head `adc4b61f`、attempt 1）= 12 条全 completed/success、非绿 0 —— 其中 `Emit manifest (shape oracle)` 这一跑就是绿的 ⇒ 登记的那 8 行与 CI 那台自己的复算对齐。【第二格的形状顺手量了一遍，供下一刀】那 6 处不是一个形状：**445/448 两形硬写 `(int32_t)(…)`**，而 571/582/604/614 四形走 `scalarArg158m(…)` —— 同一个「标量操作数怎么窄化进 `vb6_VarCmpLong*` 的参数槽」被答了两遍，且 §B138b 那层 `vb6_ChkLong` 只可能从后一条形状出来（`cgen_util_type.cpp` 的 `narrowCheckAssign` 那条路），前一条在 x64 上仍是**裸截**。可达性按老规矩先问（数的是「比较形」：`If/While/ElseIf … =|<> VarPtr|ObjPtr|StrPtr(…)`）—— 语料 **0 处**：那 46 处指针消费全是赋值形（`.uId = ObjPtr(…)`、`uIdSubclass = vSubclass`），走的是 `narrowCheckAssign` 那条已经被 §B138b 接住的路。⇒ 第二格今天是**纸面缺口**（与 §B134 同型），要合的是「同一个决定答了两遍」这个形状本身，不是救火；判据得自带一枚新针，否则合完仍是零暴露。


### B139 tabwalk 相 3 每步只留**恰好一拍**余量去读 post 出去的按键结果 —— 「判据与拍号耦合」这一族的第三处实物，收法是一张登记表（账 #213，2026-10-11 第三十刀，**已出 —— 门 #491（源码那一跑，12 条里 10 绿 + 两红都归完）+ 门 #493（登记那一跑，12/12 全绿、wall 11m04s）**）

- **这一格早就在门上红过一次，当时被归成「夹具节拍余量」**：门 #336 那轮的 `[VBP] tabwalk ... FAIL (output mismatch)`（§B48 末段记着：它的 `--emit-c` 产物在改前基线／手搓版／权威版三份里**逐字节相同**（x86/x64 皆然，27303 字节）⇒ 判成与产品无关，归到已登记的 #213，#337 两片复绿）。**归因当时是对的，只是没人回来把夹具改掉** —— 本刀就是把那格余量还掉，并把这一族第一次写成一张表。
- **机理 = 同一件事的第四个读数**：VK_DOWN / VK_UP 是 `PostMessage` 交给泵**异步**消化的，而旧代码 `gAkStep` 每拍 +1 ⇒ 每步之间正好只留一拍，读到的是「上一拍发的那一声」的结果；在途 `WM_TIMER` 连着排空时两拍挤在同一瞬间 ⇒ 那一拍看见「还没动」。同族：#162（CI 上已关掉的 Timer 仍报 21 拍）、门 #332（pbsb 五条通知读成全 0，饿机器上到第 8 拍才到齐）、#253（modal `hops=3→1`，§B83）。
- **本机量不出症状，所以这一刀的立论不是「复现」**（与 #253 同一面）：**改前**那台 C3.exe 编**今天这份夹具**，两架构 ×（空载 + 8 枚忙等压 4 vCPU）各 8 趟 = 32/32 命中 `AK-pre=optA/down=optB/wrap=optA/up=optB/clicks=3`，墙钟 0.9s → 2.4s（`.build/b688_tabwalk/report.txt`）。⇒ 收益全在 CI 那一头；本地能给的只有机理、负控与一张表。
- **修法（`tests/tabwalk/WalkForm.frm` 相 3）**：三步改成**四步** —— 最后一枚键也要等它自己那一拍（`up=` 由相 4 读，以前同样是固定一拍之后才读，同一格缺陷换个消费面就还在）；每步先问 `WhereIs(GetFocus()) = gAkLast`，没动就 `gAkWait = gAkWait + 1` 并 `Exit Sub`（**不重发** —— 每拍重发会跳站），等满 `AK_WAIT_MAX = 40` 拍就照常读并把 `gAkLate + 1` ⇒ 真丢一条键仍然读成错的、仍然红，只是不再由「队列这一刻排空了几条」决定。四步 × 40 拍 × 50ms ≈ 最坏 8s，仍在 `Invoke-TestExe` 那 60s 之内。
- **判据一个字没改**：`$twNeedles` 里那条 `AK-…/clicks=3` 原样保留，只**新加** `TW-SETTLE late=0`；`maxwait` 只打不钉（钉死它就是把「本地一拍就到」焊进判据）。读数（`.build/b691_k30_check/report.txt`）：改后两架构 × 空载/2×超订 各 8 趟 = **32/32 全针命中**，`late=0`、`maxwait=0`、`TW-ticks=12`（旧的是 11，多出来那一拍正是新加的「只等不发」那一步）。
- **负控（这一格不是把判据改成永真）**：`.build/b694_negctl` 那份**拷贝**里把那三步的 post 全关掉（守卫改成 `If gAkStep < -1 Then`），同一份夹具重编真跑 ⇒ x64 `down=optA/wrap=optA/up=other(0)/clicks=0`、`TW-SETTLE late=2/maxwait=40`、125 拍、8.0s；x86 `late=3/maxwait=40`、129 拍、8.5s —— `AK-*` 与 `TW-SETTLE` 两条针**都缺**，而进程**自己在 8 秒退**（`rc=0`）。⇒ 「保险丝有限」的意思就在这儿：没有它，这一趟就是账 #236 那种被 60s 超时杀掉、只剩半截 `.out` 的现场。
- **收法 = 登记表，不是逐处补 sleep**：第 24 道哨兵 `scripts/check_fixture_timer_close.ps1` 长出第二条问句（原来只管「阈值收线的自增顺序」）。W1 = 凡处理器里出现 `PostMessage(` 的 `*_Timer` **逐枚登记**（现 4 枚：ModalDlg/`t2_Timer`、ModalMain/`tMain_Timer`、PbForm/`tmrStep_Timer`、WalkForm/`tWalk_Timer`），新发 post 却没登记就红；W2 = 每枚带着**自己的** settle 口径，且必须在**被它保护的那次读取之前**（modal 的 `(Timer - gLastBeat) < PHASE_MIN_SEC`、pbsb 的 `ElseIf allIn Or mStep >= 20 Then`、tabwalk 的 `If gAkWait < AK_WAIT_MAX Then Exit Sub`）；唯一刻意不等的那枚（ModalDlg 只 post `WM_CLOSE`）登记成 last-act：post 必须仍是处理器**最后一条语句**，一旦有人在它后面补读法就红；W3 = 每步只自增一次 + 保险丝取 5..100（太小则泵还没消化就放弃，太大直接吃 60s 预算）；W4 = 夹具打的 `TW-SETTLE late=` 必须被 runner 钉住，两头缺一即红（账 #226 那一课：只打不判的读数等于没打）。**七发假改动各红各的**（`.build/b697_neg_report.txt`：baseline rc=0；未登记的新 post / 撤掉 settle 行 / 把 settle 挪到读取之后 / 把 post 挪出每步守卫 / 把 `AK_WAIT_MAX` 改成 400000 / 改掉打印标签 / 给豁免那枚补一句读法 ⇒ 七条不同的 FAIL）。
- **门上的读数（这一刀的收益确实只在 CI 那一头）**：`Tests (vbp #2)` 与 `Tests (vbp #3)` 两片里的 `TabWalkApp.out` 各打 `TW-SETTLE late=0/maxwait=0`、`TW-ticks=12`、`AK-pre=optA/down=optB/wrap=optA/up=optB/optA=N/optB=Y/clicks=3` ⇒ 两架构的期望串**一字未变**，改的只是「什么时候读」；本机那 32 趟也是 `maxwait=0`（一拍就到）。取读数走的还是工件那条路：`gh run download <run> -n test-logs-vbp-<k>`（本机拿不到 job 日志，工件拿得到）。
- **同一跑里那条与本刀无关的红（归因照 §B133 那套「缺 `.err` 的就是被超时杀掉的那枚」）**：`Tests (vbp #1)` 红在 `ctrldlg_probe` —— 那片工件里 `DlApp.out` 停在 `DL10` 且**没有 `DlApp.err`**，而上一片全绿的 #490 同一枚是 15 行齐 + `.err` 在（`.build/b706_vbp1` / `b709_prev_vbp1`）。这已是**第二次**（第一次 = 门 #484），而 §B133 那条待办写得很清楚：再红一次就给它一条**具名的逐例墙钟预算**，不放宽判据、也不把全局 60s 抬大 ⇒ 立新账 §B140 去办，别混进本刀那笔登记。

- **护栏**：51 道 `[STATIC]` 本地全 rc=0；语料形状清单本机重算一遍 = **inputs=400 / 未变 399 / 移动 1**，移动的正是 `tests/tabwalk/TabWalkApp.vbp`（两侧 rc=0，无新增/缺失输入，`.build/b699_manifest_k30.txt`）。那一行**按 CI 读数登记** —— 本刀没动 `src/`，除夹具自身之外一行都不该动，动了就是归因没做完。
- **本刀没做的**：`clicks=3` 与 `AK-up` 读的是**落点**不是通知条数的时刻，那两条口径没动；modal 与 pbsb 的既有等法只是被**登记进同一张表**（它们本来就在答同一问），没有改写；`appqquit` / `combofocus` 那一族（#259 / 门 #403）住的是另一问（进程肯不肯自己退），仍由账 #236 那三道闸管。**下一枚写「post 完再读」的夹具必须先进 W1 那张表**，这条口径才算有权威。
### B147 第三十六刀的 oracle 补课**引入了一族静默回归**：浮点->窄整型目标的取整与范围检查整块丢了（2026-10-11 形状门 #522 的 10 行逐行归因量到；**已出 = 第三十七刀两格**（第一格 `210bba68`：判据收成一条；第二格 `b66869d4`：打包器那一侧改问文本）；门 **#527**（head `b66869d4`）唯一红 = 形状门，按 CI 读数登记在 `41bce27a` ⇒ **门 #528**（head `41bce27a`）十二 job attempt 1 全绿）

- **读数**（`.build/b1121_attr/report.txt`；两份二进制都取自 **CI 自己的工件**：BASE = run 38110792863 (#516, head `d8ccff59` = 本批登记过的那台)、NEW = run 38119054002 (#522, head `67f03217`)；同一份工作区的 VB 源，两台各跑一遍 `--emit-c` 再逐行比）。移动的行分两族：
  - **族一 = 本刀要的**：`(vb6_VariantToLong(x) + 1)` -> `(vb6_VarAddV(x, vb6_VariantFromValue(1)))`（VBFlexGridDemo 12 行、test_paramarray 2、test_varcmp_scalar 1、test_variant_boxing 11、Charts/Proyecto1 的 `lPoints` 那两行）。
  - **族二 = 回归**：`VB6_SA_AT(uint8_t, clrFore, 0) = vb6_ChkByte(vb6_FltToLng(vb6_Num_Div(...)))` -> `= vb6_ChkByte(vb6_Num_Div(...))`；`LineOut = vb6_ChkInt(vb6_FltToLng(...))` -> `vb6_ChkInt(...)`；dbgdlg 的 `tLogFont.lfHeight = vb6_ChkLong(vb6_FltToLng((Num_Div(...) * (-1))))` -> `vb6_ChkLong((Num_Div(...) * (-1)))`。**vb6_FltToLng 整块消失** = 浮点进 Byte/Integer/Long 目标不再四舍六入五取偶、也不再报错误 6，改成 C 在调用边界上向零截断。落在 4 份真工程（Charts 2020 全家）+ tests/dbgdlg。
- **根因（一句话）**：`narrowCheckAssign`（`src/backend/cgen_util_type.cpp:1360`，取整那一档在 `:1383`）按 `inferExprType(value)` 选包装 —— Double/Single/Currency 套 `FltToLng`，**Variant 不套**（它假设发码已经交出整只槽，槽的提取由赋值那一路先做完）。而 `67f03217` 只改了 oracle 的一侧：Div 见到任一侧 Variant 就答 Variant，可**改道**那一步（`cgen_expr_binary.cpp` 的 `isVariantOperand86`）问的是 **C 文本前缀表**。两条口径不一致 ⇒ 「oracle 说是槽、发码交的是裸 double」这一族就掉进 `Variant` 那一档又不提取 = 族二。实测触发者都是**文本表读不出、AST 读得出**的那一批：`VB6_SA_AT(uint8_t, clrFore, i)`（元素访问）、`vb6_UserControl_TextHeight(...)`、`vb6_Screen_TwipsPerPixelY()`。
- **与门 #517 是同一条不一致的两个极**：那一红（C2440）是「发码已是槽、oracle 还答 Double」= **响**的那极；这一族是反方向 = **静默**的那极。⇒ §B145 那条通用读数要说得更准：**不是「改了 C 类型就顺手改 oracle」，而是「谁答『这枚 C 表达式是什么型』只能有一处」** —— 现在同一问句有两名答案（`inferExprType` 走 AST、`cExprIsVariant` 走文本），任何一侧单独动都会造出一个极。
- **为什么判据面拦不住**：这些行的差别只在小数取整，编译 rc=0、启动照跑、两台同形；本机三枚夹具（53/53）、VBFlexGridDemo 与 Charts 2020 的 rc 与 exe 全都在。**只有清单的行哈希动了** —— 形状门这一格的价值就在这。⇒ 「登记之前逐行归因」不是仪式：这次它抓到的是产品缺陷，不是清单抖动。
- **修法形状（第三十七刀的题面，别先斩后奏）**：把「这枚 C 表达式是不是整只槽」收成**一条判据** = 文本权威 ∪ AST 权威（`cExprIsVariant(text)` ∪ `isDefinitelyVariantExpr(ast)` ∪ `knownVariantVars_`），**发码改道与 Div 的 oracle 同吃这一条**。方向性是关键：oracle 只能说「AST 那半」，而改道能吃「AST ∪ 文本」⇒ **oracle 答 Variant 的集合 ⊆ 改道真的发火的集合**，族二按构造不可能出现（改道一发火就交出真槽，`narrowCheckAssign` 的 Variant 档就名副其实）。顺带：§B146 那三条（CDec/CVErr/GetAllSettings）由 AST 那半自动接住，不必再手抄名字。 **⚠ 这句是开工前的预测，实测不成立**（同一刀之后把探针 `.build/b1119_b146/probe.bas` 再跑一遍 `--emit-c`：65/67/71 三行**仍是裸 C 的 `+`**） —— 原因是 `isDefinitelyVariantExpr` 对**内置函数名一律返回 false**（它的 IndexOrCallExpr 分支里 `kBuiltinFuncNames.count(bname)` 那道闸，本意是把 IIf/Choose/Switch 这族「按上下文发类型化的」内置函数挡在外面）。所以 §B146 不是被本刀顺手结掉的，它要的是另一刀：**让那道严格判据能读出「哪些内置函数的 C 返回本来就是槽」**，名单从 RTL 头的返回类型派生（见 §B146 的 census 那格）。
- **开工前要先量的两件**：① 装箱面 —— 改道一宽，`vb6_VariantFromValue` 会吃到更多 C 形状（它有 `default: vb6_VariantObject` 兜底，但 UDT 结构体实参那一族会 C2440，先在语料里数出来）；② String 那侧的口径不能松动：静态 String 的一侧仍走拼接（`VarAdd` 是按数值强制，实测 `"5"+1` 交 6），而**盒里是串的 Variant** 参与 `+` 现在会交回错误 13 —— 那正是 VB6 本人的答案（`Dim v: v="5": v+1` = Type mismatch），要在台账里写死成口径而不是巧合。
- **落地（两格）**：① 「这枚 C 表达式是不是整只槽」新增一条共用的 AST 判据 `variantOperandAST`（`cgen_util_type.cpp`，= `isDefinitelyVariantExpr` 且排除「Variant 数组本体」那一档），Div 的类型 oracle 改吃它，发码改道吃 `isVariantOperand86 ∪ variantOperandAST`（`cgen_expr_binary.cpp` 新增 `routeOperand86`；`isVariantOperand86` 本人刻意**不加** AST 那半 —— 它还供着 Fix 086 那条 `vb6_VariantToLong(...)` 提取，那一档只能对真的是槽的文本出手）。方向性由此构造：改道 ⊇ oracle。② dev 刚合进来的 `packArgValue`（Fix 266，一条收口八处装配站点）只认 `vb6_VariantFromComResult(` / `vb6_VariantFromValue(` 两条前缀，读不到改道交回的 `(vb6_Var…V(...))` ⇒ 实测 `Charts 2020 ClsResizer.c(144)/(150)` 报 **C2440「无法从 vb6_VARIANT 转换为 double」**。补的那一条**不是**跟着拆标量：这一族改交 `vb6_ComPackValue`（整只 VARIANT 原样递给对端），因为拆成标量就是再一次在编译期替运行期选数值档 —— 正是 §B145 要消灭的那件事。它原有的两条前缀分毫未动。
- **判据（本机，全语料 400 份逐行）**：BASE = CI 在 dev head `c8adb814` 的那枚 exe（含第三十六刀 + oracle 补课 + dev 的 Fix 254-268），NEW = 本刀两格之后的本机 exe（`.build/b1147_k37/report.txt`）。**移动的行 = 54，且 54 行全是族二回到 BASE 形状**（`vb6_ChkByte(vb6_FltToLng(vb6_Num_Div(...)))` 这类回来了）；`lost-FltToLng` = **0**；routed 那一族相对 BASE **一行没动**（VBFlexGridDemo 与 test_paramarray / test_varcmp_scalar 都是 -0/+0）。真编译两台：Charts 2020 两台 COMPILE-RC=0 且出 `Proyecto1.exe`、VBFlexGridDemo 两台出 `VBFlexGridDemo.exe`；`test_variant_boxing` 两台 **56/56 零 miss**（本刀新加三枚行为针：`arv / 2` 进 Long 与进 Byte、定长 Byte 数组元素 `cad(1) / 2` 进 Byte —— 都钉 4，回归那台交 3；两台都不碰 §B77① 的平局口径）；`check_rtl_embedded` 125/125（exe sha `db90d808965e`）、`check_variant_i8_coverage` violations=0。
- **处置**：门 #522 的 10 行**不登记**。等第三十七刀之后按 CI 那一跑自己的读数一次性归因（那时族二应当回到 BASE 形状或变成「交出槽 + 提取 + 检查」三件套，族一按运算符逐行对上）。 ⇒ **订正（同一刀之内）**：族二已由上面两格消掉，本机全语料逐行验过（54 行全是回到 BASE 的那一族），所以登记改到「这一跑的 CI 读数」上做 —— 那时清单里属于 dev 的 Fix 254-268 那一批行也会一起进来（他们的形状门红了没登记），逐行归因要按两个 head 分开对：dev 的账归他们，本线只认自己那 54 行 + routed 族。 **收线数字**：登记后那跑（head `41bce27a` = **门 #528**）十二个 job 全绿；本刀在全语料上的净效果 = 清单从「第三十六刀移动 10 行」收到「只剩 5 行」，而那 5 行逐行都带运行期出口；dev 的 Fix 254-268 在同一口径下移动 **0 行**。

### B146 内联 `CDec(...)` 参与的算术两架构都编不过：`error C2088`「`+` 对 struct 非法」（2026-10-11 第三十六刀的 H 族探针顺手量到；**已出 = 第三十九刀**（提交 `87107920` 本体 + `217b2f91` 按 CI 读数登记形状门那 4 行；门 **#536**（head `87107920`）= 12 job 里 11 绿、唯一红是形状门 ⇒ 登记后 **门 #537**（head `217b2f91`，attempt 1）十二 job 全 completed/success；那张手抄前缀表从此与 RTL 逐名对账 = 第 54 道哨兵 `check_builtin_slot_returns.ps1`））

- **读数**（`.build/b1095_arith/probe6.bas`，两台同一份错误形状：三条 `error C2088` 落在 `probe6.c(18)/(21)/(24)`，另两条 `C2198` 是它们的下游）：`CStr(CDec("1.25") + 1)`、`TypeName(CDec("1.25") + 1)`、`v = CDec("1.25") + 1` 三形全红，而同一份探针里的 H 族（Variant 盒参与的标量目标 / 比较 / 数组下标 / 数值形参）两台都编得过 ⇒ 坏的只有**内联那一形**。
- **发码**：`vb6_DebugWriteBSTR(vb6_CStr(vb6_VariantFromValue((vb6_CDec(vb6_VariantFromValue(vb6_VariantFromValue(vb6_BSTR_FromStr(L"1.25")))) + 1))));` —— `vb6_CDec` 的返回类型是 `vb6_VARIANT`（`vb6rtl_builtin.h:403`），而它**不在** `cExprIsVariant` 那张前缀表里（HEAD 与本刀之后都不在，`git show HEAD:src/backend/cgen_util_type.cpp | grep -c` = 0），于是 Fix 086 的提取不 fire，C 的 `+` 直接落在 struct 上。
- **不是第三十六刀引入的**：这三行的发码文本里没有本刀那四枚出口名 ⇒ 走的仍是老路径，两台 BASE/NEW 同形。顺带说明这条**存量 0 处**：编不过的形状不可能活在语料里（除非有人正等着它）。
- **修法形状（开工前先量）**：把返回 `vb6_VARIANT` 的**内置出口整族**登记进那张前缀表（不是只补 `CDec` —— 那张表是「这枚 C 表达式是不是 Variant」的唯一权威，登记之后四则运算的改道 / 比较 / 解包 / 装箱全部自动读它，而一族一起数才不会留下第二份名单）。名单从 `vb6rtl_builtin.h` 的返回类型里抽，抽完要逐枚量可达（§B143 那条「名单上有格 ≠ 有人在读」）。
- **census 实到**（2026-10-11，`.build/b1118_b146_census.py`，零编译；站点匹配前先剥注释 —— §B145 那条课在这里同样管用）：RTL 头里返回 `vb6_VARIANT` 的函数 **33 枚**。两本名单双向都不齐 —— 发码真的发出的那一批里 **10 枚答不出**（`vb6_CDec` / `vb6_CVErr` / `vb6_GetAllSettings` / `vb6_IIfVariant` 四条连 `cExprIsVariantCarrier` 都不接；`vb6_VariantByte` / `VariantDate` / `VariantInt` / `VariantSingle` / `VariantArrayGetVal` / `VariantArrayElemPtr` 六条只被 carrier 那侧接住，而**算术改道问的是 `cExprIsVariant`**）；反过来那张手抄前缀表里 **12 条是死名**（`vb6_VariantFromI2/I4/R4/R8/BSTR/Bool/Date/UI1/SafeArray/SafeArray1D` + `vb6_VariantNothing` + `vb6_DispCallByVtbl`：`src/rtl` 与 `src/backend` 各 0 处，只活在表里）。⇒ 「名单上有格 ≠ 有人在读」与「读的人没登记」在同一条表上同时发生。
- **正确形状不是补名字**：语义层早就答对了 —— `src/semantics/builtin/builtin_funcs.inc:133/137/259` 把 CDec / CVErr / GetAllSettings 都登记成 `Vb6Type::Variant`，而 `isDefinitelyVariantExpr`（`src/backend/cgen_util_type.cpp:575`）读的就是这张表。实测 `--emit-c`（`.build/b1119_b146/emit.txt`）：65 行 `CDec("1.25") + 1`、67 行 `CVErr(5) + 1`、71 行 `GetAllSettings(...) + 1` 三形都发成**裸 C 的 `+`**（= C2088），而同文件 77 行 `v + 1`（`v` 是 Variant 变量）正确走 `vb6_VarAddV` ⇒ 差的就是**改道那一问没有去问 AST 那一侧**。这条与 **§B147** 是同一个缺口的两极，修法同一条（一条判据 = 文本 ∪ AST）。 **订正（第三十七刀之后，就在同一份探针上重测）**：改道现在已经吃了 `variantOperandAST`（= `isDefinitelyVariantExpr` 且排除 Variant 数组本体），三行**依旧不接** → 先条件是那道严格判据先能读出「CDec / CVErr / GetAllSettings / CallByName / LoadRes* 这些内置函数的 C 返回本来就是 `vb6_VARIANT`」，否则 AST 那半永远答 false。派生名单的口径：`builtinFuncs` 的右值（VB 名 → RTL 函数名）与「RTL 头里返回 `vb6_VARIANT` 的那 33 枚」取交集，再把 IIf/Choose/Switch 那族**按上下文发类型化**的排除掉。
- **另读到的两格**（同一份探针）：`vb6_VariantArrayGetVal(` 与 `vb6_VariantArrayElemPtr(` 只在**晚绑定下标**那一路发出来，实测 `o.Result(0) + 1` 走的是 `vb6_VariantFromComResult(` 那条（表里有）⇒ 这两枚在算术上是**纸面缺口**；而 `IIf(True, CDec(1), CDec(2)) + 1` 发成 `vb6_IIfLong(...)`（`cgen_expr_call_builtin_pre.inc:79-96` 的档位选择只认「实参是 knownVariantVars_ 里的标识符」，两支都不是 ⇒ 落到 Long 档）⇒ IIf 那一族的档位选择问的是**第三种**口径，等 §B147 的那条判据落地之后再回头看它要不要一起吃。
- **暴露面**：VB6 真语料 `D:\vb_yqt4qPac` 里 `CDec(` 21 处 / 14 份（§B142 数过），其中 VBFlexGrid 那两处是**赋值**（`RowID = CDec(0)`），不是算术 ⇒ 这一格今天也是潜伏项，但严重度是**硬编译错误**（不是静默答错数），开工顺序不必排在 §B144 之前。
- **第三十九刀落地（2026-10-11，按上面「修法形状」那一族一起数，没有只补 `CDec`）**：
  - 表里进 10 名（`vb6_CDec` / `vb6_CVErr` / `vb6_GetAllSettings` / `vb6_IIfVariant` / `vb6_VariantArrayGetVal` / `vb6_VariantByte` / `vb6_VariantInt` / `vb6_VariantSingle` / `vb6_VariantDate` / `vb6_VariantIdentity`）、出 11 名（`vb6_VariantFromI2/I4/R4/R8/BSTR/Bool/Date/UI1/SafeArray/SafeArray1D` + `vb6_VariantNothing`）。表 32 条 = RTL 槽出口 30 + `vb6_VariantFromValue`（_Generic，逐档核过每档都交槽）+ `vb6_DispCallByVtbl`（下面那条）。
  - **`vb6_DispCallByVtbl` 不是死名，是本刀唯一一把 census 报错的**：RTL 里 0 处、后端字面量里也 0 处，可它由**语料工程自己**发出来 —— `tests/VbQRCodegen-master/src/mdQRCodegen.bas:1407` 写着 `Private Function DispCallByVtbl(...) As Variant`，而单文件（非多模块）发码时工程过程的 C 名是裸的 `vb6_<proc>`（`cProcName`：只有多模块才拼模块名），于是 `variantReturnFuncs_` 也够不着（那张集合里只有 `vb6_<模块>_<proc>` 与 `vb6_pv_<proc>`）。撤掉它的实测后果：mdQRCodegen 五处 `hResult = DispCallByVtbl(...)` 少掉 `vb6_VariantToLong` 提取，整只槽进 `vb6_ChkLong(int64_t)` —— 这是 400 份 A/B 里唯一一条**归因不到**的移动，所以「RTL 里没有这枚符号」从来不等于「后端不会 emit 这个名字」。⇒ **同形缺口另立 §B148**（这枚名字只是它唯一有人踩过的实例）。
  - **A/B（`.build/b1231_ab.py`，BASE = 第三十八刀那台 exe，400 份 `--emit-c`）**：`inputs=400 changed=4 unexplained=0 rc_diff=0`。四份全归因到本刀进的那十名：`tests/cc_dll/CoDll.vbp` ±4（第三十八刀那三枚夹具成员的 `CDec`）、`tests/test_variant_boxing.bas` ±14（新针自身）、`tests/VBFlexGridDemo/Common/Common.bas` ±2 与 `VBFlexGridDemo.vbp` ±36（语料里那两处 `RowID = CDec(0)` + `CVErr` 那一族）。护栏口径照旧：**没有一行不许无法归因**。
  - **判据（`tests/test_variant_boxing.bas` 末尾 +9 枚针，两台逐行相同，x86 同针同过）**：`dec-add=1234568`（内联 `CDec` 直接进 `+` = 第三十六刀那族运行期出口）、`dec-vt=14`、`dec-tn=Decimal`、`dec-dbl=2.5`（Double 目标）、`iif=2`；证人三枚 `err-vt=10`、`err-iserr=True`、`reg-*`。负控 = BASE 那台**整条夹具编不过**（6 条错，两台同形：`.c(331)` C2088、`(337)` C2088、`(339)` C2440、`(344)` C2440 + 两条下游 C2198）⇒ 前四枚针的红面在编译而不是读数，正是这一族的形状。`GetAllSettings` 那两枚只钉「内置出口交出的槽能走到 Variant 目标」这个形状，不钉 VB6 口径（另立 §B149）。
  - **第 54 道哨兵 `scripts/check_builtin_slot_returns.ps1`**（登记在 `run_tests.ps1` 的 `[STATIC] builtin_slot_returns`）：R1 表→RTL（死名/错档）、R2 (RTL 槽 ∩ 后端字面量)→表（漏登记）、R3 指针返回（`vb6_VARIANT *f(`）一律不许进表、R4 census 不许读空。实到 `table=32 rtlSlot=33 rtlPtr=1 rtlMacro=25 emittedNames=1264 (RTL槽∩emit)=30` ⇒ 两个方向同时零漂移。两条负控真红（表里塞 `vb6_VariantBogus(` ⇒ R1 红；撤 `vb6_CDec(` ⇒ R2 红），按 md5 还原后复绿。踩到的坑记两条：① 后端字面量那把针原先要求**紧跟 `(`**，于是 `{"cdec","vb6_CDec"}` 那张内置名表整个读空（R4 当场报「扫描读空」—— 没有 R4 就是一道绿着的坏哨兵）；② 「发射面」必须扫**字符串字面量**而不是全文，否则注释里提一句就把名字算成活的。


### B148 单模块（非多模块）发码时，工程自己的 Variant 过程直接赋给非 Variant 目标 = 整只槽进 `vb6_ChkLong`（2026-10-11 第三十九刀归因 `vb6_DispCallByVtbl` 那一行时顺手量到；**覆盖面已量 = 存量 0 处**，所以这一格按纸面缺口押后，不在 §B147 那条口径之前）

- **读数**（`.build/b1241_single/SPro.bas`，两台 BASE/NEW **同形**，所以不是本刀引入的）：
  ```
  Private Function MakeDec() As Variant : MakeDec = 1234567 : End Function
  Public Function Twice(ByVal n As Long) As Long
      Dim h As Long : h = MakeDec() : Twice = h
  End Function
  ```
  ⇒ `h = vb6_ChkLong(vb6_MakeDec());`（`vb6_ChkLong` 的形参是 `int64_t`，`vb6rtl_builtin.h:392`）—— 真编就是 C2440。
- **成因**：「这枚 C 表达式是不是槽」的 AST 那一半（`inferExprType` / `variantOperandAST`）认得出工程过程的返回类型，所以算术那一族两台都对（同文件 `MakeDec() + 1` 正确走 `vb6_VarAddV`）；**赋值那一族的消费点问的是文本前缀表**，而单模块发码把工程过程拼成裸 `vb6_<proc>`（`cProcName`：只有 `isMultiModule_` 才拼模块名），`variantReturnFuncs_` 里登记的却是 `vb6_<模块>_<proc>` 与 `vb6_pv_<proc>` ⇒ 名字对不上，表里就只有靠人手工钉的那一枚（`vb6_DispCallByVtbl`）。
- **修法形状（开工前先量覆盖面）**：让预扫描在单模块模式下也登记**裸名**那一形，或者让赋值那一族改问 AST 那半（§B147 的「改道 = AST ∪ 文本」正是这条口径，只是赋值消费点还没吃它）。⚠ 前者不能无条件做：`cProcName` 的注释本身就写着「单模块的裸名会撞 RTL 函数名（用户 `Private Sub ErrRaise` ⇒ `vb6_ErrRaise`）」⇒ 把裸名塞进「返回槽」那张集合，撞到 RTL 同名符号时是**按名字**答 Variant，会把 RTL 那枚真调用一起吞了。先数语料里这个形状有多少处（存量只从 `--emit-c` 就能数，不必真编），再定哪一侧。
- **暴露面已量（2026-10-11，两把 census 相互独立，都零编译）**：`.build/b1263_b148_census.py` 扫 400 份 `--emit-c` 产物，找 `= vb6_Chk*(vb6_X(` 且 `X` 在**同一份产物里**声明返回 `vb6_VARIANT` ⇒ **0 处**；`.build/b1264_chknames.py` 把定义假设也撤掉，只数 `= vb6_Chk*(vb6_X(` 的全部内层名（1754 处，全是 `FltToLng` / `IntDiv` / `IIfLong` / `Screen_Width` 那一族标量出口）⇒ 里面**没有一枚**是工程自己的过程。结论：这一族今天在语料里 0 处，实物只有 §B146 里那枚幸存的钉子（`vb6_DispCallByVtbl`）+ 本账的复现探针 `.build/b1241_single/SPro.bas`。所以它是 §B147 那条口径（一条判据两头吃）还没铺到的**第三个消费点**，等下一份真语料撞上它再动，别为一形去改一片。

### B149 `GetAllSettings` 交出的形状**谁都读不出**：VarType 答 12、UBound 答 0、元素读回空，而缺键那一档该答 Null（2026-10-11 第三十九刀写针面时读 RTL 读到，第四十刀开格前量完；**2026-10-12 第四十刀：格①② 已出**（提交 `cc65f2cf` / 门 **#539** = run 38165976944、head `2723ede2`、attempt 1 = 12 job 全 completed/success），**格③ 已出**（第四十一刀，见 §B153））

- **真跑读数**（探针 `.build/b1265_nd/t149.bas`，x64 新台，`SaveSetting` 两枚键之后）：
  ```
  vt=12 isnull=False        <- VB6: VarType=8204 (vbArray+vbVariant), IsNull=False
  ub1=0                     <- VB6: UBound(a,1)=1 (两枚键, 0..1)
  ub2=0                     <- VB6: UBound(a,2)=2 (第二维 1..2)
  cell02=                   <- VB6: a(0,2) = "V1"; 这里交**空串**
  miss-vt=12 miss-isnull=False   <- VB6: 应用没有条目时交 Null (VarType=1, IsNull=True)
  ```
  五读全部**静默**（rc=0、没有一声诊断），也就是这条 API 现在等于交不出东西。
- **发码现场把三格分开**（`.build/b1265_nd/e.txt`）：
  - ① **标签**：`vb6rtl_registry.c:113` 造了一枚 1×2N 的 `vb6_SafeArray1D` 交替表，却把 `result.vt` 留在 `vb6_vtVariant` 上、只写 `result.parray` ⇒ 既不是数组档也不是 Null（`vb6rtl_variant.h:172` 那行现成的正确拼法是 `vb6_vtArray | vb6_vtVariant` = 8204，正是 VB6 的 VarType 答案）。这一格**只改 RTL**，且缺键那一档一起补（交 `VT_NULL`）。
  - ② **二维本体 + 晚绑定的第二下标**：`v(0, 2)` 发成 `vb6_VariantArrayGet(&v, 0)` ⇒ **第二个下标整块被丢**，静默拿第 0 格；而同一份产物里 `UBound(v, 2)` 走的是 `vb6_UBoundND((vb6_SafeArrayND*)(vb6_VariantToSafeArray1D(v)), 2)` —— 两条路对同一枚槽的**维数**判断不一致（一边按 ND 问，一边按 1D 取）。补它要造真 2-D（`vb6_SafeArrayCreateND(variant, 2, bounds)`，dim1 = 0..n-1 / dim2 = **1..2**）并给晚绑定元素读补 ND 那一臂；第三十九刀 §B144 那条口径在这里同样管用：**先问谁拥有这枚描述符**，别在发码侧再拼一份。
  - ③ **`DeleteSetting` 的参数个数**：RTL 是 `vb6_DeleteSetting(BSTR,BSTR,BSTR)`（`vb6rtl_class_com.h:57`）三枚全必填，而 VB6 是 1/2/3 参三形（删键 / 删节 / 删应用）⇒ 本探针写两参时 `error C2198 用于调用的参数太少`（第一次跑就是这么红的）。这一格与 `check_rtl_proto_arity`（1282 对签名）同族但不同问题：那条钉「头与体一致」，这里是「**可选形参**的默认档没人补」。
- **外部口径读数**（VB6 真语料 `D:\vb_yqt4qPac`，按 API 逐个数）：`GetAllSettings` **9 行 / 6 份**，其中 4 行是**同名成员**（`eControl.GetAllSettings()`、`clrDepth.GetAllSettings`）与 1 行关键字清单字符串，真正的内置函数消费者只有 1 份：`文件/多窗口的文件编辑器.../mdoceditor/E-Type/Main.bas:269`，它写的形状就是那张二维表的证人 —— `varFiles = GetAllSettings(ThisApp, ThisKey)` 之后 `For i = 0 To UBound(varFiles, 1)` 且取 `varFiles(i, 1)`（第二维下标从 **1** 起，列 1 = 键名 / 列 2 = 值 ⇒ 与 MSDN 的「first column 键名、second column 值」对得上）。同一族的 `SaveSetting` 1151 行 / 293 份、`GetSetting` 1178 行 / 307 份 —— 写侧与单键读侧被大量真代码用着，而**我们的写侧走的是 VB6 标准路径** `HKCU\Software\VB and VBA Program Settings\<app>\<section>`（`vb6_RegBuildKey`），所以判据可以做**真往返**（SaveSetting → GetAllSettings → DeleteSetting 收尾），不必靠假注册表。
- **格①② 落地（2026-10-12 第四十刀）**：① 标签与缺键那一档**一起**补 —— 没有条目一律 `result.vt = vb6_vtNull`（读回 VarType=1 / IsNull=True），有条目挂 `vb6_vtArray | vb6_vtVariant`（= 8204，正是 VB6 的 VarType 答案）。② 真造二维：`vb6_SafeArrayCreateND(vb6_sa_variant, 2, bounds)`，`bounds[0]={0,numValues}` / `bounds[1]={1,2}`（VB6 的第二维下标从 **1** 起，列 1 = 键名、列 2 = 值 —— 与上面那行 mdoceditor 证人对得上），填格走唯一一枚 `vb6_SaNdElemPtr` + `vb6_VariantAssign`。发码侧「第二枚下标整块丢」是同一个缺口的另一极，收在 §B150。
- **判据（已落地）**：`tests/test_registry_settings.bas` 十八枚针，x64 与 x86 **逐行相同**（读数 `.build/b1280_rs/r6.txt` / `r3.txt`）：缺键两档 `R1/R4 = vt=1 isnull=True`；两枚键那档 `R2 = vt=8204 / lb1=0 ub1=1 lb2=1 ub2=2 / cell01=K1 cell02=V1 cell12=V2`；`R3` 钉「交回的是快照」（写槽不写注册表，VB6 同形）；`R5` 钉同一枚元素的**地址**那一形（§B150 的下半场）。负控 = BASE 那台跑**同一份**新夹具（`.build/b1282_neg/nr.txt`）：18 枚**全错**（`R2-cell12=` 空、`R5-elemptr-nonzero=False`）而 rc=0 —— 静默答错那一族，不是编不过，所以针面能红这一条是**读数**给的而不是编译错误给的。
- **格③ 仍开着**（本刀刻意不碰）：VB6 的 `DeleteSetting` 有 1/2/3 参三形，我们那枚 RTL 出口只认三参 ⇒ 夹具里四条删除一律写全。这一格与 `check_rtl_proto_arity` 同族但不同问题（那条钉「头与体一致」，这里是「**可选形参**的默认档没人补」）。
- ⚠ 动 RTL 必 touch `c3rtl.rc` + 重编 C3.exe + `check_rtl_embedded`（本刀实测：125 份 RCDATA 逐字节对上，exe sha 从 `fed613ce` 变 `d6ebfe95`，arity 哨兵的成对数从 1269 变 1270 —— 那枚新出口 `vb6_VariantArrayElemPtrND` 的头/体两条都被数到了）。

### B150 晚绑定槽的元素读写被抄了六遍，而多维那枚下标整块丢掉（2026-10-12 第四十刀 **已出** = 提交 `cc65f2cf` / 门 **#539**；发射侧一处 + 认文本侧一张表 + RTL 一枚档位表）

- **题面（读数在 §B149 那枚真二维槽上）**：`v(0, 2)` 发成 `vb6_VariantArrayGet(&v, 0)` —— 第二个下标在**发射点**就被丢，静默读回第 0 格；可同一份产物里 `UBound(v, 2)` 走的是 ND 那一问 ⇒ 同一枚槽「是几维」这个决定说了两遍、漏了一遍（发码侧按一维取，运行期按二维问）。四个发射点（`cgen_expr_call_prelude.inc` / `cgen_expr_call_callee_ident.inc` / `cgen_expr_call_callee_member.inc` 两处）各抄一遍，两个写点（`cgen_assign_value_sem.inc` / `cgen_setlet.cpp`）各抄一遍。
- **发射侧收成一处**：`variantSlotElemRead(slotArg, positional)` —— 秩数 = 实参个数，一枚走 `vb6_VariantArrayGet`（既有语义，存量产物逐字节不变），两枚及以上走 `vb6_VariantArrayGetND(slot, N, …)`；写侧 `variantSlotAssignLine` 把已发出的元素读换成那**一对**写出口（Get→Set / GetND→SetND，装箱后的值追加在参数表末尾 —— 两枚 ND 出口的形参顺序就是为这个改写定的）。
- **下半场是哨兵报出来的**（这一格值得单独记）：把新名 `vb6_VariantArrayGetND(` 登记进 `variantPrefixes` 时，第 54 道哨兵 **R2 报红**（「RTL 声明返回 vb6_VARIANT、后端会 emit、表里没有」）。顺着那一问把「认文本」的站点全扫了一遍，量到**六处手写字符串比较、抄的是两种前缀** —— 带尾括号的 `"vb6_VariantArrayGet("`（4 处）与不带括号的 `"vb6_VariantArrayGet"`（2 处）。新名进表之后带括号那四处对它一律答 `false`，而两种退化都不是我要的：
  - `VarPtr(v(1,2))` 那一路（`cgen_expr_call_builtin_pre.inc`）落进复合字面量分支 ⇒ `(intptr_t)&(void*){vb6_VARIANT}` ⇒ **C2440**（响亮，编不过）；
  - `Set v(1,2) = 对象` 那一路（`cgen_setlet_set_rhs.inc` 的「左值是不是调用结果」那一长串 &&）因为认不出而答 true ⇒ 整条赋值被当「求值后丢弃」处理（**静默**）。
  ⇒ 「发射时只改四处」并不等于收完：**认文本那一问也是同一个决定的第二份实现**。修法是一张表管三种名（读 / 取址 / 写，参数形态三者同形，所以改写只是换函数名）：`kSlotElemForms` + 两条出口 `isVariantSlotElemRead` / `variantSlotElemAddrRead`，六处站点全部改问它；RTL 侧为多维补一枚取址出口 `vb6_VariantArrayElemPtrND(v, rank, …)`（越界交错误 9 = 唯一那枚 `vb6_SaNdElemPtr` 的口径，与一维那枚返 NULL 的既有语义刻意不同，一维那条不在本格）。
- **兜底那一格要说清**（不是「先让它绿」）：`nd->dimCount` 落不进 1..16 时那其实是一枚**自家 1D 描述符** —— 1D 的首字段是魔数 `0x5A1D`（= 23069，绝不可能是维数）。这一问认得出来，于是三条 ND 出口一律退回**同一枚一维出口**。为什么必须退而不是抛：VBFlexGridDemo 的 `ArrRows(iCol, iRow)` 今天就是这个形（自家把多维折成一维那一族 = §B151），改前读第一下标读得好好的，把它换成错误 9 等于拿一个新缺陷换掉一个存量缺陷。折叠本身另立 §B151。
- **RTL 侧同一批**：档位表 `vb6_VariantElemFromSlotData` 从 `vb6_VariantArrayGet` 的函数体里提出来只留一份（一维与多维都从它出来）；`vb6_sa_elem_size` public 化成 `vb6_SaElemSize`（「步长由元素类型决定」这一件事不许有第二份答案）。
- **判据与护栏**：夹起见 §B149（`R2` 钉读、`R3` 钉写、`R5` 钉地址）。发码三形的实物在 `.build/b1280_rs/e.txt` 第 99 / 107 / 117 行（`vb6_VariantArrayGetND(&v, 2, 0, 1)`、`vb6_VariantArraySetND(&v, 2, 1, 2, vb6_VariantFromValue(…))`、`(intptr_t)vb6_VariantArrayElemPtrND(&v, 2, 0, 1)`）。400 份语料 `--emit-c` 只动 **1 行**（VBFlexGridDemo，BASE 发 `vb6_VariantArrayGet(&ArrRows, iCol)` → 新发 `vb6_VariantArrayGetND(&ArrRows, 2, iCol, iRow)`，运行时行为由上面那条兜底保住），而**下半场（认文本合一）实测零额外移动** —— 同一份语料两发逐行相同（`.build/b1285_emit_new.txt` vs `b1287_emit_new.txt`，只有新夹具自己那一行变），也就是说把六处抄法收成一张表这一步**不改产物**。
- **结构性哨兵**：第 54 道 `check_builtin_slot_returns.ps1` 加了一条 **R5** —— src/backend 除 `cgen_util_type.cpp` 之外任何文件的**字符串字面量**里出现 `vb6_VariantArray(Get|GetND|Set|SetND|ElemPtr|ElemPtrND)(` 就红（剥注释，与 §B145 那条课同口径；`GetVal` 那一枚不在名单里，它是按值再取一层，不是同一种改写）。负控实测：临时插一枚 `src/backend/_b1289_negative.inc` 写 `"vb6_VariantArraySet("` ⇒ `FAIL R5 第二份抄法`；拔掉 ⇒ `PASS (R1..R5)`，census 尾数 `slotReadLit=10`。

### B152 门红的自救面自己也会红：并行那条路打的不是「<名字> ... <判定>」排版，§B92 的点名器一次都没接住（2026-10-12 门 **#538** 撞见，**已出** = 提交 `8020360c`（点名器第二拼法）+ `2723ede2`（§B145 自伤修复）/ 门 **#539**）

- **读数**（判据是拿门自己的工件回放，不是本机模拟）：`Tests (bas #1)` 打 `Results: PASS=61 FAIL=2`，两条真红是 `[RUN] test_variant_boxing: output mismatch` 与 `[RUN] test_variant_boxing_x86: output mismatch`；而 §B92 那条注解通道打的是 `run_tests exit code 1 but no case line matched` + `harness tallied 2 failures but naming found only 0` —— 数得着、点不出名，自救面在最需要它的那一轮哑了（它自己那句「fix this cell」就是给这一天写的）。
- **根因还是同一个决定有两份写法**：`tests/run_tests.ps1` 的**并行汇总分支**（jobs>1）打的排版是 `  [RUN] <用例名>: <原因>`（六条原因同一形状：compile FAIL / compile FAIL 无 error C 行 / no exe / run timeout / run error / output mismatch），而 §B92 的匹配器只写了**串行**那条 `  [类别] <名字> ... <判定>`。⇒ 一条判据两种拼法，读的一侧只认一种；这与 §B150 那六处手写「认文本」的比较是同一族毛病，只是这次在 CI 侧。
- **修法**：给点名器补**同一判据的第二拼法**（正则 `^\s*\[(.+?)\]\s+([^:\[\]]+?):\s+(.+)$`，先试老排版、不中再试这条），**不往名单里手抄用例名**（那会变成第三份）。
- **判据（两半都用门 #538 的原件回放）**：把 `gh run view --log --job 114541302116` 取回的 bas #1 全文喂进同一份匹配代码 —— 改前 tally=2 / named=0，改后 tally=2 / **named=2**，两条名字正是 `test_variant_boxing` 与 `test_variant_boxing_x86`；负控 = 同一份脚本吃**绿的** bas #2 工件 ⇒ tally=0 / named=0（绿轮不许凭空点名，这条限制是 §B92 原有的，仍然守着）。
- **顺带把 #538 那两条红归完因**（它们是本刀的**成功**信号，不是缺陷）：第三十九刀在 `test_variant_boxing` 末尾钉的 `VB-b146-reg-vt=12` / `reg-isnull=False` 钉的正是当时那个错答；第四十刀把 §B149 格①② 修成 VB6 口径之后，门立刻把这两枚针打红 ⇒ 那两枚「形状证人」是**会红**的判据。按真跑读数改口成 `vt=1` / `isnull=True`（x64 与 x86 逐行相同，`.build/b1306_box/r6.txt` / `r3.txt`，与 CI 上 `test-logs-bas-1` 工件里 `test_registry_settings` 的 18 枚读数同形）。
- **⚠ 本刀另有一格自伤要在这一页记死**（同一次修复里已改回）：第四十刀的台账补丁把「插入锚点」写成了 `anchor[0]` —— 锚点已经是**字符串**，`[0]` 取到的是第一个字符 `#`，于是 §B145 那枚条目标题被换成一个 `#`（标题前半句被吃，正文与其余 141 条毫发无伤，`git diff --numstat` 只多一行改动，完全看不出)。修法是照 `git show 50084bce:` 的原件把丢掉的半句接回去。**规则**：台账这类一行一条目的文件，插完之后必须数 `^### ` 的条数（改前/改后各一次，只许多出插进去的那几条），别只数 numstat。

### B153 内置过程「省略的实参补哪一档」这一问有 38 位主人，而语义层那本 `isOptional` 名单是**刻意不被使用的第二份**（2026-10-12 第四十刀之后量完；**第四十一刀 已出** = 提交 `ef978c0d` / 门 **#541**（11 job 绿，唯一红是形状门的本刀那一行，已按 CI 读数登记）→ 登记后的验证轮 **门 #542**（head `6220eb40`、run 38171582528、attempt 1）**12 job 全 completed/success**，一张 29 行的表 + 第 55 道哨兵，§B149 格③ 顺带收掉）

- ** census 实到（零编译，`D:\c3.vb6.pro\.build\b1319_pad_census2.py`）**：`src/backend/detail/expr/cgen_expr_call_builtin_fixup.inc` 一个文件里，带 `args.size()` 判断的内置名 **38 枚**，代码形状有 **4 种**：① `argList += ", <字面量>"`（多数）② `while (args.size() < N) args.push_back(...)`（`Mid` / `MidB` / `StrConv` / `FormatNumber` / `ErrRaise`）③ 一张默认值数组按下标取（`InputBox` 的 `kInputBoxDefaults[7]`）④ **反向截断**（`ErrRaise` 的 `args.size() > 3` 把 VB6 的五参砍成三参）。另有 5 枚的默认值不是常量而是**当模块的 Option Compare**（`Replace` / `Filter` / `InStrRev` / `Split` / `StrComp` 的 `cmpDefault`），所以任何"表 + 一个函数"的收法都必须把这条上下文带进去。
- **第二份名单确实存在且没人读**：`src/semantics/builtin/builtin_funcs.inc:59-60` 自己的注释写着 —— 「VB6 Optional 参数标记 `isOptional=true`（语义保留），**但 RTL C 函数不接受 Optional padding 和 IsMissing `_has_` 尾叜 —— cgen 用硬编码补默认值**」。⇒ 语义层那本名单为的是 `IsMissing`/包装档，发码层的 38 处为的是补参数，**同一个「第 N 格可不可省、省了交什么」的问题两个方向各答一次、互不认识**（§B146 那本前缀表的形状）。
- **语料可达性（`D:\vb_yqt4qPac`，14385 份，按 GBK 解码、只数括号调用形）**：`Mid` 911 处省略 / `StrConv` 426 / `GetSetting`（三参，缺 default）258 / `InStrRev` 323 / `Shell` 217 / `Round` 161 / `StrComp` 147 / `FormatNumber` 51 —— 也就是说这 38 处**大多有人在踩**，收表不是整理死码。同一把 census 的 `Left` 74 处"一参"是 `Public Property Let Left(...)` 那种**同名声明**，属于 §B149 那把 census 撞过的同名假阳性，别当缺口。
- **本刀要顺手收掉的那一格 = §B149 格③**：`DeleteSetting` 是**语句形**（不带括号），逐行数是 11 处调用，其中 **3 处只给两参**（删整节）。而 `vb6rtl_registry.c:96` 的 else 分支**已经**实现了「key 为空 ⇒ `RegDeleteKeyW(节)`」—— 缺的只是发码侧没给第三枚实参，于是两参写法在 runner 上是 `error C2198 用于调用的参数太少`（第四十刀的探针就是这么红的）。⇒ 这一格是「RTL 已备好、发码不接」那种最便宜的真缺陷。
- **刻意不做的那半格**：`DeleteSetting app`（一参，删整个应用）在语料里 **0 处**，且 RTL 开头 `if (!appName || !section) return;` 直接挡死 —— 按「先证明可达再动」的口径只登记不实现，要做得连带 `RegDeleteTreeW` 那一档的语义与 VB6 对齐。
- **修法形状（开工时照这个）**：一张 `builtinArgDefaults` 表：内置 C 名 → 每一档省略实参要补的 C 表达式（含 `cmpDefault` 这类上下文占位）+ 一个"最多几枚、多了截断"的上限。发射点仍在一处（今天就在 `cgen_expr_call_builtin_fixup.inc` 里，位置合理），但那 38 处**手抄的判断**改问表；`isOptional` 那本名单要么被表吃到、要么撤掉，不许继续两份并存。判据 = **零行为改动**：401 份语料 `--emit-c` 逐行不变（这一刀是纯收口），加一枚 `DeleteSetting` 两参的真跑夹具（`tests/test_registry_settings.bas` 里加 R6：两参删除后整节不见 → `GetAllSettings` 答 Null），BASE 负控 = 同一份夹具在 `cc65f2cf^` 上编不过（C2198）。哨兵：把「表里的名 ↔ 语义表 ↔ RTL 声明的参数个数」三面接进 `check_rtl_proto_arity` 那一家（它已经有 1270 对签名可查），并钉「`cgen_expr_call_builtin_fixup.inc` 里除表那一条路径外不许再出现 `args.size()` 与 `callee == "vb6_` 的合写」。
- **落地（第四十一刀）**：表 = `kBuiltinOptDefaults`（**29 行**，就在 `cgen_expr_call_builtin_fixup.inc` 里 `cmpDefault` 之后），一趗循环替了 28 段手抄；每一行按 **RTL 声明的绝对槽位**给缺省实参，`nullptr` 那一档让缺省链**停住**（少写的实参原样交给 cl 的 C2198，不再把 `nullptr` 拼进 `std::string`），`"%cmp%"` 是唯一的非字面档（发码时换成 `cmpDefault`，即本模块的 Option Compare），`intoArgs` 那一列记下**下游会按 args 重建 argList**的六枚（Mid / MidB / StrConv / FormatNumber / Clipboard_SetData / ErrRaise）——它们必须同时写进 `args`，只补 argList 会被 conv_numeric / conv_output 抹掉（这条口径就是 Format 那格早先写下的那一课）。
- **零行为改动已量**：一份 40 形探针（`.build/b1332_k41/b141.bas`，覆盖表里 29 行中的 26 行）在 BASE（`.build/C3_base_k41.exe`，第四十刀那台）与 NEW 上 **141 行只有 1 行不同**，而那 1 行正是要修的 `DeleteSetting` 两参形（读数 `.build/b1332_k41/{base.txt,new.txt}`）；401 份语料清单逐行对照，**只动 1 行** = `tests/test_registry_settings.bas`（本刀给它加了 R6，bytes 5793 → 6742），其余 400 行的 sha256 与**顺序**全同（`.build/b1336_manifest_new.txt` vs `emit-manifest.expected.txt`，本机那台 14.29 产出与 CI 那台 14.51 登记的那些行逐字相同）。
- **门 #541 的三个读数（2026-10-12，回填）**：12 job 里 **11 绿**，唯一红是 `Emit manifest (shape oracle)`；点名的就是本刀自己的夹具。CI 自己的那份（分支 `ci/emit-manifest`，标题 `ci: emit manifest head=ef978c0d`）与登记清单对账 = **相同 400 / 移动 1 / 新增 0 / 消失 0**，移动的那行 = `tests/test_registry_settings.bas`（sha256 `8fc73c73…` → `44140bbb…`，bytes 5793 → 6742），正是本刀给它加了 R6 三枚读数的结果；而**本机那台 14.29 独立复算出来的正是同一串 sha256**（`.build/b1336_manifest_new.txt`），所以登记这一行不是抄本机的答案、也不是猜的（两台互验，与第四十刀同一口径）。登记后 `emit-manifest.expected.txt` 逐行只动 **1/1**、27 条 `#` 注释行不变、字节数不变（旧 sha 与新 sha 同长），`check_manifest_coverage` 复算 inputs=401 / registered=401。
- **登记之后那一轮（门 #542，2026-10-12）**：head `6220eb40`（run 38171582528，attempt 1）—— 12 job **全 completed/success**，其中 `Emit manifest (shape oracle)` 由 #541 的红转绿，也就是那一行按 CI 自己的读数登记之后，**CI 再算一遍还是同一串**。这一轮是本刀的收线轮，没有新缺陷读数。
- **§B149 格③ 落地**：`DeleteSetting` 两参 = 删整节——RTL 早就实现了（`vb6rtl_registry.c:99-110`：key 为空走 `RegDeleteKeyW`），缺的只是发码那一档。夹具 R6 三枚钉真跑（`R6-before-vt=8204` / `R6-two-arg-delete-vt=1` / `R6-two-arg-delete-isnull=True`，两台逐行相同，`.build/b1337_r6/{nr6.txt,nr3.txt}`）；负控 = BASE 那台跑**同一份**新夹具根本编不过（`.build/b1337_r6/b6/c3-error.log` 第 87 行 `error C2198: “vb6_DeleteSetting”: 用于调用的参数太少`，x64/x86 同样，两个 RC 都是 1）。一参那档按上一格的口径**没接**：`kOD_DelSetting` 前两档是 `nullptr`，写一参会响亮编不过而不是静默无事。
- **第 55 道哨兵** `scripts/check_builtin_opt_defaults.ps1`（R0..R8，census `rows=29 arrays=29 rtlArityNames=2359 semParamTables=92 leadingDefault=3 notPadding=12`）：表 ↔ RTL 逐名档数（R0 死名 / R1 档数不齐）、表 ↔ 语义那本 `isOptional` 的可选档集合（R5）、表内名字不许在同一个文件里再往 argList 追加（R4，这就是「新写一形硬编码补档就红」那条）、留在发码侧那 12 枚「按实参个数做别的决定」的段必须逐名带理由且恰好等于名单（R7，多一名少一名都红）、四项下限不许读空（R6）。**四把负控各点各的红**：插一行 `vb6_NotARealExit` ⇒ R0 死名；把 Round 指到 6 档数组 ⇒ R1 档数不齐 + R5 形参个数不齐；表外加一段 `if (callee == "vb6_Space") { argList += ", 0"; }` ⇒ R4 表外补档；插一行**重复的** Join ⇒ R8。
- **R8 是自证挣来的一条**（不是凑数）：第一版负控插了两行重复的 DatePart/Filter，本期望的 R1 一声不出——因为扫描用的是按名字取数的哈希表，**同名两行只有后写那行生效**，前一档的缺省答案静默失效（与 022 记的「PSParser 0 错而绑定已变」同一族，这次发生在哨兵自己身上）。所以把「一枚出口只许一行」写成判据，而不是只把插件改成不重复的。
- **顺带量到的一条不齐（不在本刀）**：语义表把 `Dir` 的 pathname、`CurDir` 的 drive 标成必填，而 VB6 手册是 `Dir[(pathname[, attributes])]` / `CurDir[(drive)]` —— **第一档本身可省**。发码侧的表答对了，语义那本答错；今天不影响产物（通用 padding 对 builtin 干脆不吃 `isOptional`），所以哨兵里用 `$leadingDefault`（3 条，逐条带手册依据）登记而不改语义表 —— 改它会动 `IsMissing`/包装档，那是另一格的事。
- **剩下的口径（不是代码）**：语义那本名单与本表的关系现在是「哨兵强制两边齐」，不是「一家吃掉另一家」。真要合成一份，得先让 `addBuiltinFunc` 那 10 枚（无形参表）也把形参补齐 —— 而给它们补 calleeParams 会把 Fix 024 P2 / Fix 029 那两条按形参类型决定的**包装改道**一起打开，那是一次独立的刀，不是本刀的零行为改动口径能养得起的。

### B154 `ReDim` 重建动态数组时**元素类型在创建那一步就丢了**：造出来的是 variant 档、写进去的是声明那档的步长，两者不同一格（2026-10-12 第四十一刀量完；**2026-10-12 第四十二刀 已出** = 提交 `b86b61af` / 门 **#545**（12 job 里 10 绿，两处红都是登记本身：形状门的 20 行 + 清单覆盖门的 K1，均已按 CI 读数落账）→ 登记后的验证轮 **门 #547**（head `a2e38c30`、run 38183854441、attempt 1）**12 job 全 completed/success**，一张表 + 两处强转方向 + 夹具十七枚针 + 哨兵加四条规则）

- **真跑读数**（探针 `D:\c3.vb6.pro\.build\b1366_nd\b151.bas`，x64 新台，`.build/b1366_nd/r6.txt`）：`Dim a() As String` + `ReDim a(1 To 3, 1 To 4)` + `a(2, 3) = "X"` 之后
  ```
  P3-cell=1299894904        <- 该答 X；这里把一枚 BSTR 当整数打出来了
  P4-vt=8204 tn=Variant     <- VB6: VarType=8200 (vbString=8 + vbArray=8192), TypeName="String()"
     （本行最初写的是 8208，那是没有根据的归纳；第四十二刀按它自己 `vb6rtl_variant.h:20-32` 的枚举重算：8200 / Byte()=8209 / Long()=8195）
  P10-box-cell=             <- v = a 之后 v(2,3): 空串 —— 写的那一格与读的那一格不是同一格
  ```
  同一段里 `Dim b() As Long` + `ReDim b(2,2)` + `b(1, 2) = 77` 读回 **77**（P7）而两维的 LBound/UBound 两档全对（P1/P2/P5/P6）⇒ **不是下标那一路坏，是元素类型那一路坏**。
- **发码现场**（`.build/b1366_nd/e.txt` 第 72 / 73 / 85 / 123 行）：同一件事被说了两遍且互相矛盾 ——
  - 创建：`a = (vb6_SafeArray1D*)vb6_SafeArrayReDimND(vb6_sa_variant, 2, bounds)` —— **`vb6_sa_variant` 是写死的**，声明里的 `As String` 到这一步没了；
  - 写：`vb6_BSTR_AssignMove(&VB6_SA_ND_AT2(BSTR, (vb6_SafeArrayND*)a, 2, 3), …)` —— 按 **BSTR 的步长**算地址；
  - 于是 x64 上写入落在 stride 8 的位置而描述符按 stride 24 数格子 ⇒ 写入与读出不是同一格（P10 交空串就是这件事的读数），而 P4 那声 8204 是它的**症状**；
  - 另有一格独立的：`Debug.Print a(2, 3)` 发成 `vb6_DebugWriteLong((int32_t)(VB6_SA_ND_AT2(BSTR, …)))` ⇒ 把 BSTR 当整数打（账 #180 那一族「指针被当数字打」的另一处，与类型档无关；**本刀没动它**，已挪进 §B155 格⑤）。
- **可达性（真语料 `D:\vb_yqt4qPac`，10138 份，GBK 解码，`.build/b1356_nd.txt` / `b1357_nd2.txt`）**：二维 `ReDim` **414 处**、二维 `Dim` **172 处**（含 `Dim a(3, 3) As Long` 那一族）；同一轮量到 `Variant = 二维数组` 那一形是 **0 处**（那是 §B151，与本格不同形）。⇒ 本格是**有人踩的**，且踩的是静默写错格。
- **修法形状（开工时照这个）**：先问 §B144/§B150 那条口径 —— **谁拥有这枚描述符**。元素类型只有一个答案来源：声明里那个 `As <类型>`（`vb6_sa_*`）。【开工后的订正】这里写的「只被一维那条路吃到」是错的，而且错得保守：**一维的裸写 ReDim 同样坏**（`Dim s1() As String` + `ReDim s1(5)` 也发 `vb6_SafeArrayReDim1D(vb6_sa_variant, …)`，台账原文只拿到了多维那一个读数，因为写侧读侧自洽时它们不会自己报错）；相反，**固定数组两维那条一直是对的**（`Dim f(5) As String` → `vb6_sa_bstr`、`Dim g(2, 3) As Long` → `vb6_sa_long`，实测）⇒ 缺口不在「多维」而在「ReDim 这一句话没问声明」。收成一处 = 创建与访问都问同一张类型表（第四十刀已经把元素步长收成 `vb6_SaElemSize` 一枚出口，**类型档**是它的另一半）。判据 = 两台真跑的 P3/P4/P7/P10 四枚 + 存量 401 份清单逐行护栏（这一刀会**改产物**：`vb6_sa_variant` 那枚字面量要换成按声明算出来的档，所以登记清单会动，动的那几行必须逐行归因）。
- **顺带**：`Erase a` 之后 `UBound(a, 1)` 交 **0**（P11），VB6 是数组未分配 ⇒ 这一档要不要跟着改，等本格定完类型档再说（未分配数组的 UBound 口径是另一条，先记在这里）。

- **格① 落地（元素档改问声明）**：一枚出口 `resolveReDimElemType(lowerVar, typeRef)`（`src/backend/expr/cgen_expr_array.cpp`，与 `resolveArrayElemType` / `resolveArrayUdtElemCType` / `resolveArrayClassElemType` 同一族）—— 写了 `As` 子句照旧优先，没写就问 `arrayElemTypes_`（七处写入点填的那张表：三条形参路（`cgen_decl_func` / `_proc` / `_prop`）+ 两条模块级（`cgen_decl_var:81` 固定 / `:125` 动态）+ 两条局部（`cgen_localdecl:102` / `:140`），访问侧 `VB6_SA_AT` / `VB6_SA_ND_AT2` 吃的本来就是它），表里没有才回落 Variant（隐式 Variant 数组 / 成员链 / With 成员 —— 沿用 Fix 180 的「不去猜」纪律）。`src/backend/stmt/cgen_redim.cpp:92` 那一行改问它，一维与多维、Preserve 与非 Preserve 四条发射路一起接上（它们本来共用同一个 `saElemType`）。**为什么 As 子句那一边不反过来写表**：真语料 2323 行带 As 的 ReDim 里与声明相异的只有 8 行（7 行还是正则抓错声明），没有实物 ⇒ 只补「读」这一头（`.build/b1370_k42/census2.txt`）。固定数组那条路一直是对的（`Dim f(5) As String` 发 `vb6_SafeArrayCreate1D(vb6_sa_bstr, 0, 5)`，实测），所以本刀不是新口径、是把动态那条接回同一个既有答案来源。
- **格② 落地（同一支路上撞见的响缺陷）**：多维的 Variant 槽那一支在 BASE 上**根本编不过** —— 赋值那行把 `(vb6_SafeArray1D*)` 套在 `vb6_VariantFromValue(...)` **外面**（一维那条没有这层强转，所以只有多维红）= cl C2440（`.build/b1373_fix/b6.log`）。去掉强转又暴露下一层：`vb6_VariantFromValue` 是 _Generic，名单里只有 `vb6_SafeArray1D*` 这一档，裸递一枚 `vb6_SafeArrayND*` 掉进 `default: vb6_VariantObject` ⇒ 槽的 vt 成 DISPATCH，而 `vb6_VariantArraySet/GetND` 开头只认 `vt & vb6_vtArray` ⇒ **一声不响地既不写也不读**（实测 R8 读回空串）。修法是把投影写进装箱那一句**里面**（`vb6_VariantFromValue((vb6_SafeArray1D*)…)`），沿用仓里「1D 指针当载体、用时投回 ND」的既有约定，不给 _Generic 开新档。
- **判据（两台真跑）**：新夹具 `tests/test_redim_elem.bas` 十七枚针，x64 与 x86 **逐行相同**（`.build/b1373_fix/{rf6,rf3}.txt`），注册在 `test_redim_elem` + `test_redim_elem_x86`。负控两形：① 同一份夹具摘掉 R8 那一对（BASE 编不过的那条）在 BASE 上真跑 ⇒ `R3-box-cell=` 与 `R4-box-cell=` 两枚交**空串**、其余 15 行同（`.build/b1373_fix/rnb6.txt`）；② 整份夹具在 BASE 上压根编不过（C2440）。三枚 `*-label-*` 钉的是**今天**的 VarType/TypeName（8204 / "Variant"）—— 那是 §B155 格①② 欠的答案，谁修对标签这三枚就红，是成功信号不是回归。
- **发码 A/B（13 形探针）**：`.build/b1370_k42/p42.bas` 覆盖「裸写 × 带 As × 一维 × 多维 × Preserve × UDT × 声明 Variant × Variant 槽」，BASE/NEW 只差 **7 行**，全是要修的创建档；带 As 的两形、UDT 两形、声明 Variant 那形、Variant 槽那形**逐字节不变**。
- **护栏（清单逐行，含本刀自己算错的一次）**：语料 401 → **402** 份（新夹具进清单），**移动 20 行 / 新增 1 行 / 消失 0**，rc=1 那 77 行（刻意负例）一处没动；20 行共 **78 处**发码差异，逐行分型 **全部**是创建档那一处（byte / string / long / integer / longptr / double → `vb6_sa_*`），装箱强转那一格**零存量移动**（语料里没有 `ReDim <Variant 槽>(i, j)` 的实物）。移动的 20 份里 6 份是 `tests/Charts 2020` 家族（Proyecto1 28 处、ucProgressCircular 13 处、ucChartBar 5 处…），其余 14 份是老的 array 族夹具与 VBFlexGridDemo / M6Test / arr_cls / dbgdlg / ve_list。**归因工具自己错过一次**：第一版按 `\S+` 取 relpath 当键，把带空格的路径截成最后一段 ⇒ 六份 `…/Proyecto1.vbp` 同名互相覆盖，于是数出来「移动 14 / 差异 26」而真值是 20 / 78；改判据 = 键必须取 `bytes=<n> ` 之后的**整行余量**并断言无重复键（`.build/b1385_attr2/attr2.txt`；第一份 `.build/b1376_attr/attr.txt` 留着当反面样本）。
- **登记（门 #545 的两处红都是登记本身）**：形状门红 = 那 20 行；`Tests (compile)` 红 = `check_manifest_coverage` 的 K1（新夹具进了输入集而清单没登记）—— 两处都由同一笔登记解决。CI 自己那份（分支 `ci/emit-manifest`，标题 `ci: emit manifest head=b86b61af`）与登记后的清单逐行对 = **402 / 402 全同、按哈希不同 0、缺席 0、多余 0**；而本机那台 14.29 独立复算的 402 行与 CI 那台 14.51 **逐行相同**（0 行不同），所以登记的数不是抄一台的答案。登记后注释仍是 27 行（头 20 + 尾 7），数据行 402，`check_manifest_coverage` 复算 inputs=402 / registered=402。 **验证轮（门 #547，2026-10-12）**：head `a2e38c30` 的 12 job 全 completed/success，上一轮红的两枚门双双转绿（`Emit manifest (shape oracle)` 与 `Tests (compile)`）—— 「按 CI 的读数登记」这一步被第二机重算过一遍。其余十个 job 里含 bas / vbp 两架构的真跑：本刀改了 20 份产物的创建档，而**一枚针面都没被改红**（包括 Charts 2020 那六份真工程与 VBFlexGridDemo）。
- **结构性哨兵**：`scripts/check_sa_access.ps1` 加 **A8..A11**（这一族的答案来源只许一处）：A8 唯一出口的定义 1 / 声明 1 / 被叫 ≥1 且函数体两头都问，并钉住 cgen_redim.cpp 里「声明处只问 As」= 0；A9 `vb6_sa_variant` 字面量在 src/backend 非注释里恰好 3 处（分档表 Variant 档、default 档、Variant 槽那枚刻意写死的 Preserve）；A10 装箱里那层投影恰好 1、反向的「强转盖装箱」恰好 0；A11 声明表的下标站点 ≥8、`.find` ≥3、`.erase` ==1（清表只许一处，撤掉它上一个过程的档会漏给下一个过程）。**六把负控各点各的红**（`.build/b1378_neg/neg.txt`）：改回只问 As ⇒ A8 两声；第 4 枚字面量 ⇒ A9；强转写在装箱外 ⇒ A10；撤掉装箱里的投影 ⇒ A10；第二处 erase ⇒ A11；助手改问 UDT 那张表 ⇒ A8 三声；末轮全部还原后 PASS 且 md5 对得上。没新增脚本，所以第 55 道哨兵的注册表与 `check_static_sentinel_registration` 都不动。
- **语料暴露面（真语料 `D:\vb_yqt4qPac`，10138 份，GBK 解码，零编译）**：裸写 `ReDim`（无 As 子句）4189 行，其中目标在同份文件里有声明的 2324 行；内置标量档 1596 行（byte 596 / string 343 / long 315 / integer 214 / boolean 66 / double 30 / single 24 / longptr 8）（`.build/b1370_k42/census.txt`）⇒ 本刀不是纸面缺口。

### B158 两枚多维 Preserve 搬旧元素的下标算术与访问侧那两套**相反** ⇒ 「长一格」把元素搬到别的格（2026-10-12 第四十四刀的探针撞见；**2026-10-12 第四十五刀 已出** = 提交 `c6b8bfc5`（与第四十四刀同批）。门 **#551**（head `c6b8bfc5`、run 38192712871、attempt 1）12 job = 11 绿 + 唯一红 `Emit manifest (shape oracle)` = 那 6 行登记本身；按 CI 自己那份（`ci: emit manifest head=c6b8bfc5`）落账 = 提交 `66ee2562`，登记后本机复算 402/402 逐字段相同；验证轮 **#552**（head `66ee2562`、run 38193693659、attempt 1）**12 job 全 completed/success**）

- **形状**：`vb6_SafeArrayReDimPreserveND` 与 `_Udt` 各写过一份 `oldStrides[minDims-1]=1; oldStrides[d]=oldStrides[d+1]*bounds[d+1].cElements`（**最后一维最快**），而访问侧那两套权威 —— 头里的 `vb6_SaNdElemPtr` 与体里的 `vb6_SafeArrayND_Offset`（`vb6rtl_array.c:526`）—— 都是 `stride[0]=1; stride[d]=stride[d-1]*cnt[d-1]`（**第 0 维最快**）。两套只在第一格（ ordinal 全 0 之外唯一一致的那格）agree，于是 Preserve「长一格」= 把其余格搬走。
- **读数**（2×3 长成 2×6，六格全在增长**之前**写；单发探针 `.build/b1403_k44probe/pr4.bas`）：修前那台 `POST-1-3=[] POST-1-4=[r2c2] POST-1-5=[r2c3] POST-2-2=[] POST-2-3=[] POST-2-4=[r1c3]`（`rpn4_pre45.txt`，六格里三格空、三格跑到新列上），修后 x64/x86 逐行相同、且与 BASE 那台的 `rpb4.txt` 一字不差（三格归位、新列全空）。⇒ **这一格是第四十二刀发货之后才现形的**：修前创建档错（variant 24 字节）时，两套相反的算术错得互相吸收，谁都看不见；把创建档改对，搬家就露了。
- **暴露面**：与 §B155 格④ 那份重数同一批（真语料 `ReDim Preserve` + 秩数 ≥2 = 18 行 / 11 份；凡新旧界让两套算术给出不同格位的都在搬家，增长与缩减都会）。门禁语料里今天只有夹具自己（R14/R15/R16）与 `test_arr_nd.bas` 走多维 Preserve。
- **修法（已落）** = 收成一枚 static 出口 `vb6_SaNdSlot(bounds, rank, ordinal)`（紧挨两枚 Preserve 放），两枚各 2 次调用（旧格位与新格位），`oldStrides/newStrides` 两份表撤掉，线性枚举一起改成正序（`idx[d]=tmp%copyCounts[d]; tmp/=copyCounts[d]`），秩数变化时高位维取 0（与改前同语义）。**丢弃旧元素那一段的枚举同批改成正序**：`linear` 本身就是访问侧的格位，`indices[d]` 必须由同一套算术解出来，否则 in/out 判错、释放与漏释放的正是错的那几格（BSTR / VARIANT 两支）。只动 RTL ⇒ 发码一字不挪。
- **判据**：夹具 R16 三枚（`R16-keep=a11/a12/a13` / `R16-keep2=b21/b22/b23` / `R16-new=| ub2=6`）。**R14/R15 都不是分辨针** —— 一枚只差这一刀 RTL 的编译器（HEAD 的 RTL + 第四十四刀的后端）上整份夹具只有 R16 那三行改色、其余 26 行逐字不动（`.build/b1405_k45build/rng6.txt`），因为那两枚读的不是 (1,1) 就是增长之后才写的格。这是 [[redim-udt-member-stride-test-trap]] 的第三种形：以前两次是「写与读同一枚错步长所以自洽」，这一次是**搬运算术与访问算术相反**，所以分辨针必须钉在两套算术不同意的格上。
- **护栏**：402 份清单本机复算 = **移动 0 行**（这一刀只动 RTL；上面那 6 行全是第四十四刀的）。哨兵加 **A16**：`vb6_SaNdSlot` 定义恰好 1 / 定义+调用恰好 5 / `oldStrides`+`newStrides` 恰好 0 / 每枚 Preserve 体内调用恰好 2 / 降序枚举 `for (d = arr->dimCount - 1; d >= 0; d--)` 恰好 0。**四把负控各点各的红**（`.build/b1405_k45build/neg_probes_a16.py`）：第二份步长表复活 ⇒ strideTab；枚举改回降序 ⇒ desc；一枚 Preserve 不用权威 ⇒ all=4 且枚内=1；出口改名 ⇒ def=0 且 all=4；末轮还原后 `vb6rtl_array.c` 的 md5 与植前一致（922e1e5e…）。RTL 改动 ⇒ touch `c3rtl.rc` + 重编 C3.exe，`check_rtl_embedded`（125 份 RCDATA 逐字节同磁盘）与 `check_rtl_proto_arity`（1286 签对）双验，55 道 `check_*.ps1` 全绿。
- **过门与两枚独立复算**：登记按 CI 自己那份走（本机不参与权威）。**CI-vs-CI 对照** = 把这一轮 CI 交的 402 行与上一轮 CI 交的那份直接比（本机零参与）：恰好同样 6 行不同 / 消失 0 / 多出 0 / rc 变化 0，且每行字节差 = 该枚界临时量出现的次数 × 2 个后缀字符（Charts 2020 +8、test_arr_nd +8、VBFlexGridDemo +32、mdQRCodegen +4、VbQRCodegen 的 Project1 +4，夹具自己 +3837 是新增源码行），读数 `.build/b1409_bless/ci_vs_ci_attribution.txt`。另补一枚更硬的对形：`pr4.bas` 由 BASE 那台与修后那台**各编一次**、用同一种捕获方式跑，两份 stdout **逐字节相同**（209 字节、md5 `f6eb0797cbe96b05f4647fa7a42f924e`，`.build/b1411_byteeq/{out_base,out_new}.txt`）—— 之前那句「与 BASE 一字不差」原先只有编码归一后的逐行对形做证，这一枚是字节级的做证。

### B157 `DlApp` 那枚真弹框夹具的具名预算给到 150s 仍没兜住（2026-10-12 门 #548 / 门 #569 两轮各红一次；**2026-10-13 第五十二刀 已出** = 提交 `8e252950` + 补半刀 `c2106ec0`；门 **#571**（run 38274199940、head `c2106ec0`、attempt 1，逐 job 复核 12/12 completed/success；前一笔 `8e252950` 的门是 **#570**，也 12/12 绿，但它验的是**只接了一半**的那台）—— 因不在预算那一个数，**在探针命中之后没退**）

- **读数**：门 #548 的 `Tests (vbp #1)` 红 = `[VBP] ctrldlg_probe_x86 ... FAIL (run error)`；那片工件 `DlApp.out` 停在 `DL10=Y`（针面要 `DL1..DL14` + `CTRLDLG-DONE`）而 `DlApp.err` **不存在** —— 这就是 `run_tests.ps1` §B140 注释里写死的「被杀现场」形状（超时那支只在 stderr 非空才写 .err，崩溃那支两枚都写）。同一形状在门 #484 / 门 #491 出现过两次，那两次的结论是余量不是判据，于是 §B140 给 `DlApp` 开了**具名 150s**（`$RunBudgetSec['DlApp']`，两片 × 两架构共用一个数）。
- **这一轮的新信息**：150s 仍然没兜住 `ctrldlg_probe_x86`（同轮另三片 vbp 与两片 bas 全绿）。本机空载 2.8s、8 枚忙等压 4 vCPU 读到 13.8 / 31.4 / 34.5s（§B133 量的）⇒ 那个 150s 是从「×11 放大」**本地外推**出来的，而 CI 上的放大倍数不止 11 —— 放大倍数不该由本地外推。
- **三条候选，都不许动判据**（`DL11..DL14` 一条不能少，少了等于把真模态那一格判没了）：① 把 `DlApp` 那个数再抬（最便宜，但只是把外推换成更大的外推）；② `DL10→DL11` 之间那一步是等真模态框的响应，判据面改成**每步轮询 + 保险丝**（门 #253 / #213 那一族已经这么修过两次：期望串不动，把「一步只留一点余量」换成「每步自己等到它那条来」）；③ 把真模态那几枚单独排到并发度低的 job 里（动 workflow，代价最大）。先等验证轮的复跑读数：绿 ⇒ 是负载余量，按 ①/② 挑一条；仍红 ⇒ 就不是余量，先本机下复载复现（门 #484 那次的做法）。 **验证轮 #549 里 `ctrldlg_probe_x86` 是绿的**（同一份 150s 预算）⇒ 那一轮的红按负载余量落定，本格仍开着：余量随并发漂，150s 不是安全数。
- **顺带一格观测面的缺口**：四枚用例（`ctrldlg` / `ctrldlg_x86` / `ctrldlg_probe` / `ctrldlg_probe_x86`）共用同一个**产物名** `DlApp`，日志按产物名落盘 ⇒ 同一 job 里后跑的把先跑的盖掉，红了也看不出是哪一枚的现场。这一格与 §B156 那个「临时量按数组名拼」是同一个毛病的两个面（**按名字当键必须带全身份**，见 memory `gbk-diagnostic-census-hazard` 那条第七坑）。
- **根因是读数给的，不是猜的**：给探针插一枚只有 `C3_CDTRACE=<文件>` 才打印的计时针（用完按 md5 复损，没上线），把「逐条 needle 的到达时刻」与 RTL 侧的 `show-enter / poll / fired / show-return` 一起读（本机空载 2 轮 + 8 枚忙等压 4 vCPU 3 轮，`.build/k52_tr_idle.txt` / `k52_tr_load.txt` / `k52_tr1..2_{idle,load}.log`）：**找到框不慢** —— 第一发取消的延迟空载 16..47 ms、压满 16..156 ms；慢的是**命中之后没人退** —— 一枚框被投了 **13..259 发** `WM_COMMAND/IDCANCEL`（线程每拍重数一遍窗口，`EnumThreadWindows` 每次都又认出同一枚 #32770），于是同一枚 `Show*` 的墙钟从空载 0.22..0.38 s 被推到 **0.64..4.64 s**。机制上的洞：写侧那枚 `vb6_CdProbeFired` **一个读者都没有**，而循环的退出条件读的是一枚**从不被赋值**的局部量 `found`（回调开头就 `(void)lp;`）⇒ `if (found || …)` 那半永远为假。门 #548 / #569 那两轮「具名 150 s 也没兜住」的放大，是这道机制自己造的，不是 OS 弹框贵。
- **落地三处，都在同一格上**：① `vb6_CdProbeProc` 的循环改成**两条真事件出口** —— Fired（那一发已经递进框里）或 Stop（调用方已经从 `pFn` 回来）立刻退，`found` 那枚死局部量删掉、`EnumThreadWindows` 不再收 out-param；② 上限从「400 拍」换成**按毫秒记的保险丝**（60 s）—— 实测 `Sleep(10)` 落在 15.6 ms 的系统 tick 上、压满时读到过 109 ms 一拍，所以拍数不是时间单位，注释里那句「最多等约 4 s」从头就把拍当成了秒；③ `vb6_CdProbeArm` 把 Fired **归零** —— 这条责任是新读者带来的：不归零，第二发 `Show*` 会看见第一发的命中而**一发都不投**（DL13/DL14 立刻红）。顺带把「找不到就不动，那种情况由套件的 -RunTimeoutSec 兜底（最多 60 s）」订正成指向 §B140 那张具名预算表。
- **观测面那一格一起补**（就是上面原来那条）：`Invoke-TestExe` 多一条 `-LogName`（= **用例名**），现场文件按它落盘；预算仍按**产物名**问（§B140 规矩③ 逐字未动，那张表的键还是 `DlApp`）。两个调用点（`Test-Run` / `Test-Vbp`）都补 ⇒ 与并行 .bas 那条（本来就按 `$it.Name`）对齐 ⇒ 四枚 ctrldlg 用例从此各留各的 `ctrldlg*.out`，红了不必猜是哪一片的现场。这一格与 §B156 那个「临时量按数组名拼」仍是同一个毛病：**按名字当键必须带全身份**。
- **负控两条，各点各的红**：① 行为 —— 保险丝改成「已到点」(`fuse = GetTickCount()`) 重编一台 ⇒ 线程一上膛就退、没人取消 ⇒ 夹具卡在真模态里**不返回**（本机挂到被杀，15 枚针只剩前 10 枚）⇒ 那四条针吃的是这条机制而不是巧合；② 结构 —— 把循环里那句读改回一个无人读的比较 ⇒ `check_run_budget.ps1` rc=1，红在 B7 那条文案。还原按 md5 核：`vb6forms_ctrl.c` = `2e7ccbdde99698a1ad71346a27cbcbbc`，重编后 x64/x86 两片的 15 枚 needle 全 Y 才继续。
- **判据没动**（§B140 规矩②）：`DL11..DL14` 一条不少，`DlApp = 150` 那个数也**没抬** —— 这一刀削的是机制自己造成的放大，不是把外推换成更大的外推（规矩①）。配对 A/B（同一台机器、8 枚忙等、各 6 轮、两台都**不插针**、只换 C3.exe）：改前 1643 / 1888 / 2393 / 2799 / 6338 / **9122** ms，改后 1789 / 1794 / 1821 / 1985 / 2919 / **4102** ms（`.build/k52_ab_*.txt`）；不带 `C3_CDPROBE` 时两片仍是 11 行的老形状（`DL1..DL10` + `CTRLDLG-DONE`，逐字未变）。
- **护栏**：55 道 `check_*.ps1` 全绿（B7 是 §B140 那道哨兵新加的第 7 问，**没新增脚本** ⇒ 注册名单不动；B7 那五条用字面匹配而不是正则 —— 把 `for (int i = 0; i < 400; i++)` 当正则喂进去会因为 `))+` 是嵌套量词而抛 ArgumentException，报出来的是脚本自己的错、不是这一格的判定）；402 份清单本机复算 = **移动 0 行 / 消失 0 / 多出 0**（这一刀只动 RTL 与 runner，发码面零改动，`.build/k52_cmp.txt`；门上 `Emit manifest (shape oracle)` 那一 job 也直接绿 = 独立复算对上）；`check_rtl_embedded` 125 份 RCDATA 逐字节同磁盘；改 RTL ⇒ touch `c3rtl.rc` + 重编 C3.exe 那条规矩照走。
- **本刀自己只接了一半，是门 #570 的工件把它揭穿的**（这一条比刀本身更值得抄进教训）：补丁脚本里那条『两个调用点都补 -LogName』用的锚点是「这条形状全文出现 **2** 次」，而落盘那句写的是 `replace(old, new, 1)` ⇒ **断言验多处、替换只做一处**，`Test-Run`(.bas 串行) 接上了、`Test-Vbp`(.vbp，四枚 ctrldlg 正走这条) 没接上。揭穿它的是 CI 自己的产物：`test-logs-vbp-1` 与 `test-logs-vbp-4` 两片里仍然只有 `DlApp.out`(各 117 字节 = 上膛那枚的 15 行输出)，不带 `C3_CDPROBE` 那两枚的现场照旧被盖掉 ⇒ **数 numstat 不够，要按「每个身份各留一份」去产物里点数**。补的那半刀 = `c2106ec0`，并用摘出来的真函数两头各跑一遍自证：给 `-LogName` ⇒ 现场是 `ctrldlg_probe_x86.out`；不给 ⇒ 退回 `DlApp.out`（其余调用点一字未动）。以后写「同一形状 N 处」的补丁，要么 `replace(old, new)`（不限次）配 count==N 的断言，要么逐处带上下文锚点 —— 半个都算红。
- **还剩的口径**：CI 上的放大倍数仍**不由本地外推**（本机这组数只量得出「自家那部分放大了多少」）；「框压根没出现」那一路现在由 60 s 保险丝与逐例预算两头兜，两头都不是判据。

### B156 同一过程里第二条多维 ReDim 把同名的界临时量发两遍 ⇒ cl C2374（2026-10-12 第四十三刀撞见；**2026-10-12 第四十四刀 已出** = 提交 `c6b8bfc5`（与第四十五刀同批）。门 **#551**（head `c6b8bfc5`、run 38192712871、attempt 1）12 job = 11 绿 + 唯一红 `Emit manifest (shape oracle)` = 那 6 行登记本身；按 CI 自己那份（`ci: emit manifest head=c6b8bfc5`）落账 = 提交 `66ee2562`，登记后本机复算 402/402 逐字段相同；验证轮 **#552**（head `66ee2562`、run 38193693659、attempt 1）**12 job 全 completed/success**）

- **形状**：一个过程里有两条**多维** ReDim 语句指向同一枚数组（`ReDim Preserve p(1 To 2, 1 To 3)` 之后再 `ReDim Preserve p(1 To 2, 1 To 5)`，或先 `ReDim` 再 `ReDim Preserve`）。`cgen_redim.cpp:191` 把界数组的临时量按**数组名**拼（`_redim_bounds_<名>`），每条语句各发一遍定义 ⇒ 同一个 C 函数里两次定义同名数组：`error C2374: '_redim_bounds_p0': 重定义；多次初始化`。**一维那条不吃这一刀**（它把 lBound/uBound 直接递进调用，不发临时量），所以「Preserve 增长」这一族在一维里一直是通的。
- **读数**（两台同形，与元素档无关）：`.build/b1395_k43/n6/c3-error.log` = 第四十三刀的编译器，`test_redim_elem.c(161)` 与 `(198)` 两声；`.build/b1395_k43/b6/c3-error.log` = 第四十二刀之前的编译器，同一两位同样两声（那一轮还多一声 C2440，那是 §B154 的负控）。⇒ 这一格比第四十二、四十三两刀都老。
- **暴露面**（零编译，GBK 解码）：真语料 `D:\vb_yqt4qPac` 里「同一份文件内同一枚数组被多维 ReDim 两次以上」的**文件 43 份**（按文件计，同一**过程**才是必要条件，所以这是上界；多维 ReDim 行总数 379）；门禁语料 `tests/` 里 2 份 —— `test_arr_nd.bas`（`d` 两条，分在两个过程里 ⇒ 一直没红）与 `test_redim_elem.bas`（第四十三刀新加的 `mp` 两条，**故意**分在 Main 与 GrowMp）。⇒ 门上不红，真语料一进来就撞。
- **修法（已落）**：界临时量按**语句**唯一化 —— 新增一枚出口 `nextNDBoundsTemp(base)`（声明 `cgen_helpers.inc` / 定义 `cgen_redim.cpp`），两条发码路（`visit(ReDimStmt)` 与 `emitReDimComplexTarget`）都问它，序号 `ndBoundsSeq_` 跟着 `tempCounter_` 在过程头一起复位；不是把两条语句合并（两条的界本来就不同），也不开兼容入口（0 调用者的兜底名就是死码），固定数组那条 `_bounds_` 刻意**不**卷进来（它按声明名，而 VB 不许同过程重名声明）。**判据**：独立探针 `pr3.bas` 在修前那台压根编不过（`.build/b1403_k44probe/pb3/c3-error.log` = C2374），修后 exit 0 且 `W1-kept=aa ub2=6 / W2-cell=cc c23=bb`；夹具 R15（同一过程先 `ReDim g1(1 To 2, 1 To 3)` 再 `ReDim Preserve g1(1 To 2, 1 To 5)`）就是这一形，两台逐行相同（`.build/b1405_k45build/{rk5n6,rk5n3}.txt`）。**护栏**：402 份清单本机复算 = 移动 6 行 / 消失 0 / 多出 0 / rc 变化 0，每行的字节差都恰好等于「那枚临时量出现的次数 × 2 个后缀字符」（Charts 2020 4 次 ⇒ +8、test_arr_nd 4 次 ⇒ +8、VBFlexGridDemo 16 次 ⇒ +32、mdQRCodegen 2 次 ⇒ +4、VbQRCodegen 的 Project1 2 次 ⇒ +4，夹具自己那 +3837 是新增的源码行）。哨兵加 **A14/A15**（拼法只许一处：出口声明 1 / 定义 1 / 被叫恰好 2；`"_redim_bounds` 字面量恰好 1；旧的裸拼形状恰好 0；序号复位点恰好 1；`_bounds_` 仍恰好 2）。

### B155 描述符自己的那一档：创建侧吃到了（第四十二刀）、首次分配就是多维 Preserve 也接上了（第四十三刀）、**装箱标签与 TypeName 现在都问它**（第四十六刀）；剩下**晚绑定写侧**那一处还不问（2026-10-12 第四十二刀量完；格⑤「多维元素被打成整数」另记着）

- **格① 装箱标签从不问描述符 —— 已出（第四十六刀，落地与读数见本节末那三条）**：`vb6rtl_variant.h:170 vb6_VariantArray(void*)` 把 vt 写死成 `vb6_vtArray | vb6_vtVariant`，而 `vb6_VarType(v)` 就是 `return v.vt`（`vb6rtl_conv.c:399`）⇒ typed 动态数组一律答 8204，声明那档没进标签。按仓里现成的枚举算（`vb6_vtBSTR=8 / vb6_vtByte=17 / vb6_vtLong=3 / vb6_vtArray=0x2000`，`vb6rtl_variant.h:20-32`）VB6 该答 String()=**8200**、Byte()=**8209**、Long()=**8195**、Variant()=8204（注册表那枚夹具钉的 8204 是对的，因为 GetAllSettings 交回的就是 Variant 数组）。而且「档 → vt」这件事已经有**三份写法**：这一枚写死的、`vb6rtl_array.h:189` 写死 BYTE 的、`vb6rtl_com.c:408/530` 那两枚只认 BSTR/VARIANT/I4/I2/R8/R4 的三元表达式 ⇒ 修法与 §B150 同形：提一枚 `vb6_SaVtFromElemType` 出来，四个消费点都问它（`vb6_sa_udt` 那一档交什么要单独定口径，VB6 是 8192+36=8228）。动 RTL ⇒ touch `c3rtl.rc` + 重编 C3.exe + `check_rtl_embedded`。
- **格② TypeName 没有数组分支 —— 已出（第四十六刀）**：`vb6rtl_conv.c:402` 的 switch 只认标量 vt，8204 落 `default` ⇒ 答 "Variant"；VB6 答 "String()"。格① 修完这一格才有意义（标签里有元素档才谈得上拼 `()`），同一家表就是 §B141 那两张（TypeName / IsNumeric）的第三张。
- **格①② 落地 = 第四十六刀（2026-10-12，纯 RTL 那一族）**：**「这枚数组是哪一档」第一次只有一处答** —— 新出口 `vb6_SaVtFromElemType`（`vb6rtl_array.c`，就住在尺寸表 `vb6_SaElemSize` 隔壁：两张表吃同一个枚举，也必须同一时刻一起改）+ 认载体那一问 `vb6_SaVtOfArrayPtr`（手里只有 `void*` 时先认描述符：一维首字段是魔数 `0x5A1D`、多维首字段是秩数 1..16，魔数落不进 1..16 —— 与 `vb6rtl_string.c` / `vb6rtl_com.c` 里靠魔数认载体的那几处同形；NULL 与认不出的一律回落改前那答案，不猜）。**四个消费点一起改问它**：`vb6_VariantArray`（`vb6rtl_variant.h`，装箱那一句一直写死着 `vb6_vtArray|vb6_vtVariant`）、`vb6_VariantFromUdtBytes`（`vb6rtl_array.h`，写死 BYTE）、`vb6rtl_com.c:411/536` 那两条 `VT_ARRAY` 进口（各自折过一枚三元表达式，BSTR/VARIANT 之外一律答 `vtLong` ⇒ `VT_ARRAY|VT_I2` 进来标签答 Long、载体却是 int16 —— 标签与描述符两个答案，拿着错标签再交回 COM 就是错档）。`vb6_GetAllSettings` 那句写死的 8204 也改成问描述符（它本来就是 variant 档 ⇒ 读数一字不动，这一枚是**归队**不是改动）。`vb6_TypeName` 的标量名单抽成 static `vb6_TypeNameScalar`，数组那一支按 VB6 拼 `<元素名>()`。
- **读数**（两台 43 行逐行相同，`.build/k46run2_both.txt`；针面 29 枚 → 42 枚）：`String()`=8200、`Long()`=8195、`Boolean()`=8203、`Byte()`=8209、`Integer()`=8194、`Single()`=8196、`Double()`=8197、`Currency()`=8198、`Object()`=8201、`Variant()`=8204 —— 改前**每一档**都是 8204/"Variant"。第四十二刀留下的那三枚 `*-label-*` 当时钉的就是 8204，这一刀把它们翻红，是预期内的成功信号（夹具注释原文：「谁修对标签这三枚就红」）。证人两枚：`R17-boxed` 把同一枚 Long 数组两头一起钉（标签 8195/"Long()" + **信描述符**的那一枚元素读回 4242 ⇒ 两张表必须同值）；`R17-boxw` 往盒里写 99、读回还是 4242 ⇒ 格③ 那支「对 typed 数组静默 no-op」第一次有针面，谁修对它就红。
- **暴露面（真语料 `D:\vb_yqt4qPac`，10138 份 .bas/.cls/.frm/.ctl/.pag，GBK 解码，零编译；读数 `.build/k47_b155_cell3_census.txt`）**：这一问在真代码里就是**拿 `VarType` 问「这是不是一枚数组、哪一档」** —— `vbArray` 这枚符号出现在 **173 行 / 36 份**文件里，其中 `VarType(...) = … vbArray` 这种比较形 **27 行 / 21 份**（ComCtlsDemo 的三枚 .ctl + Common.bas、GDI+ 控件那一族…）⇒ 本刀改的就是这批工程今天读得到的那两个答案。
- **格③ 的暴露面顺手量了（第四十六刀，同一份 census）**：把声明过的数组装进 Variant 槽的写法（`Dim v As Variant` + `Dim a() As Byte` + `v = a`）在 10138 份里只有 **7 行 / 去重 4 份工程**，且**没有一处**往那个盒里写下标（`v(i) = …`）—— 4 份全是 `cRegistry.cls` / `frmCollect.frm` 那一族：`ReDim abData(cData)` → `RegQueryValueExByte(…, abData(0), cData)`（Declare 的 `As Any` 落在元素上）→ `vValue = abData` → `GetValue = vValue` 交回调用方。⇒ ① 格③ 按「先证明可达」继续保持登记、不实施；② 这个形状恰好是上面那条 `VarType = vbArray + vbByte` 的消费方 —— 它以前一律读回 8204，本刀之后读回 8209，是这批工程在 VB6 里本来就有、我们一直答错的那一句。
- **打到了谁 / 没打到谁**（这一条决定本刀的爆炸半径）：按**基底档**分支去读数组标签的消费点，读码只有 VarType 与 TypeName 两枚；`vb6_VariantPackBorrow`（唯一的「槽 → 真 VARIANT」表）**压根没有数组那一档**，8204 与 8200 都落 `default` ⇒ 数组交给 COM 两架构都是静默 `VT_EMPTY`，改前改后一样（那一格另记 §B159，不记在本刀账上）；`vb6_VariantClear` 不碰数组（盒是借来的视图，载体归 VB 变量自己）；`vb6_IsArray` / `vb6_VariantToSafeArray1D` / `vb6_VariantToByteArray` / 那几条 ND 出口只问 0x2000 那一位；`vb6forms_winsock.c` 里那枚 `VB6_WS_T_BYTEARRAY=8209` 比的是**用户传进来的形参**，不是我们的标签。⇒ 一把纯标签刀，能改坏的只有「用户自己打印的那两个答案」。顺带把 §B151 那一形重新量了一遍（探针 `.build/k46probe151_both.txt`，两台同形）：`Dim a(1 To 2, 1 To 3) As Variant` + `v = a` 之后 `VarType(v)` 仍交 **9 / "Object"**、元素读回空 ⇒ 固定多维数组装箱走的是**另一条路**（压根没问这枚新出口），这一刀没碰到它，§B151 照旧开着；同一份探针里**一维**固定数组 `Dim g(1 To 3) As Long` 则答 8195/"Long()" 且元素读回 300 ⇒ 新出口吃的正是这一族。
- **护栏、哨兵与一把连锁**：`c3rtl.rc` 补 touch 注脚（§B99② 那一格）+ C3.exe 重编（`[1/2] Building RC object` 真的重跑了）。**402 份清单本机复算 = 移动 1 / 消失 0 / 多出 0 / rc 变化 0**，那一行就是夹具自己（源码改了 ⇒ 发码跟着变，其余 401 行逐字节相同 = 「RTL 体不进发码」的又一次实证；登记仍只取 CI 自己那份读数）。哨兵加 **A17**（`check_sa_access.ps1`：档表声明 1 / 定义 1 / 被叫恰好 5；认载体出口声明 1 / 定义 1 / 被叫恰好 2；写死的 `vb6_vtArray|vb6_vtVariant` 恰好 2 且只许在那两条兜底里；旧三元形状 `vb6_vtArray | (baseVt` 恰好 0；TypeName 的 `case vb6_vtEmpty:` 恰好 1、`vb6_TypeNameScalar` 被叫恰好 2），匹配前先**剥注释**（§B145 那条 phantom 站点的教训）。**三把负控各点各的红**（`.build/k46_probe_report.txt`）：一条进口退回三元 ⇒ 「被叫 4」+「三元 1」两声；数组那一支自己抄名单 ⇒ 「名单被叫 1」一声；装箱那句写死 ⇒ 「认载体被叫 1」+「写死 3」两声；末轮全还原后 PASS 且三份文件 md5 对得上。**全哨兵重扫（55 道）撞出一格连锁**：`check_variant_i8_coverage.ps1` 三条红 —— 那张名单搬进 `vb6_TypeNameScalar` 之后，`switch (vt)`（裸形参）不在它的站点形状里 ⇒ V1 的站点数 13→12，而 V3/V4 还指着 `vb6_TypeName` 要那一档。修法按 §B145 那条口径「**名单随表走**」：正则补认裸形参那一形（全仓只有这一枚，站点数回到 13 = 新表顶替旧表的位置，不是凭空多一枚），两份 must 名单里的 `vb6_TypeName` 换成 `vb6_TypeNameScalar`。判据是**站点数不许掉**：只换名字而不宽正则，census 就静默少一张表，而站点计数正是这道哨兵唯一看得见名单断电的途径。
- **门号**：门 **#554**（head `2554b7fc`、run 38200946379、attempt 1）12 job = **11 绿 + 唯一红 = `Emit manifest (shape oracle)`**，红的就是那 1 行登记本身（形状门报「一行文本变了」，预期内）⇒ 按 CI 自己那份（分支 `ci/emit-manifest` 的 `ci: emit manifest head=2554b7fc`，提交 `89c5b473`）落账 = 提交 `d98907af`，登记后本机复算 **402/402 逐字段相同**、注脚 27 行（头 20 + 尾 7）一字未动；验证轮 **#555**（head `d98907af`、run 38202084084、attempt 1）**12 job 全 completed/success**（`Tests (bas #1/#2)` 两片把 42 枚针读全 = CI 那台 14.51 与本机 14.29 在这一族读数上同值）。文档两笔 `c95eb15f` / `c3d358ee` 各带 `[skip ci]`，推完 `/commits/c3d358ee/check-runs` 的 `total_count = 0` 自证门没被文档重跑。
- **格③ 晚绑定写侧只认 Variant 档**：`vb6_VariantArraySet` / `SetND`（`vb6rtl_com.c:832/856`）开头 `if (arr->elemType != vb6_sa_variant) return;`。修前所有裸 ReDim 的数组都是 variant 档，所以这一支**看着**能动，但它按 24 字节步长写进一枚按 8 字节步长被直接访问的缓冲区（写进去那格谁也读不到）；本刀把创建档改对之后，同一支路对 typed 数组变成**静默 no-op**（比原来的内存错乱轻，但仍然不是 VB6 的行为 —— VB6 会把值转换后写进那一格）。修法 = 与 `vb6_VariantElemFromSlotData`（档位表，读方向）反着的那一枚**写方向**出口：按 elemType 落点 + 按目标类型转换。本格是本刀撞出来的**新暴露**而不是新引入：那一形（`Dim x As Variant` + 多维 ReDim）在 BASE 上压根编不过，修后才第一次走到这里。
- **格④ 首次分配就是多维 Preserve ⇒ 堆被写坏（响的）—— 2026-10-12 第四十三刀 已出 = 提交 `b139b357`（+ 行号订正 `328f5d61`、清单登记 `c72bcd51`）/ 门 **#548**（2 红都归因）→ 验证轮 **门 #549**（head `c72bcd51`、run 38188663070、attempt 1）**12 job 全 completed/success**：`vb6_SafeArrayReDimPreserveND(arr, dimCount, bounds)` 的形参表里**压根没有 elemType**，`if (!arr) return vb6_SafeArrayCreateND(vb6_sa_empty, …)` —— `vb6_sa_empty` 在步长表里落 `default: return 4`，于是分配按 4 字节一步、访问按 `sizeof(BSTR)=8` 或 `sizeof(vb6_VARIANT)=24` 一步 ⇒ 写出界。**修法** = 与一维那枚 `_T` 同一个决定：形参收 `elemType`，体内三条出口一起同形（`udt` 这里答不了 → 回落 variant；`arr` 为空 → 按调用方给的档建；档与描述符自己身上的不符 → 数据不可保留、按正确的重建，这条自愈与一维那条一字不差），发码侧 `cgen_redim.cpp:231` 把 `saElemType` 递下去 —— 而不是像修前那样只在不带 Preserve 的那一支递（:248），**同一件事两处写法不同形的第三例**。不开兼容入口：这一枚全仓只有一处发码，多一枚没有调用者的兜底名就是死码（一维那枚 `ReDimPreserve1D` 就是前车之鉴，今天 0 调用者）。**真跑读数**：`.build/b1382_preserve/` 修前 BASE 四次 `0xC0000374 / 0xC0000374 / 0xC0000374 / 0`（堆损坏，落点随布局漂，一次侥幸）；第四十三刀之后同一份 z.bas 两次都 `0` 且逐行出全（`.build/b1382_preserve/{rkz6,rkz6b}.txt`）。另有一把**只含这一形、两台都编得过**的探针 `pr2.bas`（`.build/b1395_k43/`）：修前那台 `0xC0000005` 且**一个字都没打印**（`rpb2.txt` = 0 字节），修后 `0` 且 `Y11-box-cell=six` 出。**判据**：夹具 `tests/test_redim_elem.bas` 尾部接 R10..R14 七枚新针（两台逐行相同，`.build/b1395_k43/{rn6,rn3}.txt` 的 diff 为空），针面从 17 枚长到 24 枚。**其中只有 `R13-box-c23` 一枚能分辨修前** —— 四枚直接读写（R10/R11/R12/R14）在修前的编译器上照样打出同样的值，因为写与读用的是同一枚错步长（`.build/b1395_k43/drive3.log` 的 BASE 段实测）；这条正是 [[redim-udt-member-stride-test-trap]] 那一课的又一形：**分辨针必须钉在「信描述符自己那一档的消费方」上**。护栏：402 份清单本机复算 = **移动 1 行 / 新增 0 / 消失 0 / rc 变化 0**，而那一行就是夹具自己（修前门禁语料里多维 Preserve **0 行** ⇒ 这一刀对存量产物逐字节无影响，55 道哨兵全绿）。哨兵加 **A12/A13**：A12 头与体都收档（各恰好 1，头追不上体在 runner 上是 cl C2197 = 账 #240 那一族）+ 体内 `vb6_sa_empty` 兜底不许复活 + udt 那一挡还在；A13 后端只发一处且那一处把 `saElemType` 递进去。**四把负控各点各的红**（`.build/b1396_sentinel/`）：头撤形参 ⇒ A12；体内改回 empty ⇒ A12；udt 的 `==` 改 `!=` ⇒ A12；后端不递档 ⇒ A13；末轮全还原后 PASS 且三份文件 md5 对得上。RTL 改动 ⇒ touch `c3rtl.rc` + 重编 C3.exe，`check_rtl_embedded`（125 份 RCDATA 逐字节同磁盘）与 `check_rtl_proto_arity`（1270 签对）双验。**门 #548**（head `b139b357`、run 38186952774、attempt 1）12 job = **10 绿 + 2 红**，两处红都归因到了读数：① `Emit manifest (shape oracle)` 红 = 那 1 行登记本身（报的就是一行文本变了，预期内）；② `Tests (vbp #1)` 红 = `[VBP] ctrldlg_probe_x86 ... FAIL (run error)` —— `tests/ctrldlg` 真弹模态框那一族，**与本刀无关，两条证据**：CI 自己那份清单里 `tests/ctrldlg/DlApp.vbp` 那一行 `970b3de4… / rc=0 / bytes=14157` 与登记**一字不差**（⇒ 那枚产物根本没被本刀改过，形状门也没报它）；那片工件 `DlApp.out` 停在 `DL10=Y` 而 `DlApp.err` **不存在**，正是 `run_tests.ps1` §B140 注释写死的「被杀现场」形状（超时那支只在 stderr 非空才写 .err，崩溃那支两枚都写）。同一轮另三片 vbp 与两片 bas 全绿；那片工件按产物名落盘、四枚用例共用一个 `DlApp.out`，所以「哪一枚留下的」不署名 —— 这条只当形状看，不当指认。**新的信息只有一条**：具名预算已经给到 150s 还是没兜住这一枚 ⇒ 另记 §B157。**登记按 CI 自己那份**（分支 `ci/emit-manifest`，标题 `ci: emit manifest head=b139b357`）落账：CI 402 行 vs 登记 = **移动 1 / 新增 0 / 缺席 0 / 多余 0 / rc 变化 0**，那一行就是夹具自己（`tests/test_redim_elem.bas` `84e0ece1…` → `27aeb337…`）；而本机 14.29 那台的独立复算与 CI 那台 14.51 **402 行四个字段（sha256 / ascii256 / rc / bytes）全同、0 行不同** ⇒ 登记的数不是抄一台的答案。那一行移动的归因是结构性的：修前门禁语料里多维 `ReDim Preserve` **0 行**（第四十三刀重量），所以本刀不可能改到任何存量产物的形状 —— 401 行逐字节没动就是这条的实证。**暴露面订正**：真语料按「`ReDim Preserve` + 秩数 ≥2」重数 = **18 行 / 11 份**（本节上一轮写的 17 少一行，未复核到具体是哪一行；两份读数都零编译、都不含 `tests/`），其中首次分配 9 行、带 `As` 子句的 3 行（带 As 也一样吃不到 —— 那一档根本没进 RTL 的形参表）；按档分：Variant 3 / String 3 / UDT 一族 4（走 `_Udt`，本来就对）/ Long 2 / Byte 2 / Integer 2 / 其余按名找不到声明。**撞出来的一格新问题**记在 §B156（界临时量按数组名拼 ⇒ 同一过程两条多维 ReDim 编不过）—— 那一格已由第四十四刀出（界临时量按**语句**唯一化），夹具因此多一枚 R15 把两句收回同一个过程，R14 仍留着当跨过程的证人；顺带订正：R14 那句「读回来就证明 Preserve 保住了」是**说满了**，它读的那两格恰好是 RTL 里两套下标算术唯一一致的地方，真正测「保住」的是第四十五刀那枚 R16（见 §B158）。
- **格⑤ 多维元素被打成整数（与类型档无关，但同属「多维那一路少问一层」）**：`Debug.Print a(2, 3)`（`a` 是 `String` 多维数组）发成 `vb6_DebugWriteLong((int32_t)(VB6_SA_ND_AT2(BSTR, (vb6_SafeArrayND*)a, 2, 3)))` —— 把一枚 BSTR 指针截成 int32 打印，实测交 `1299894904` 而不是 `X`；同一份产物里**一维**那条正确地发 `vb6_DebugWriteBSTR(VB6_SA_AT(BSTR, s1, 2))`（`.build/b1370_k42` 的探针第 157 / 159 行，本刀之后仍然如此）⇒ 「这个元素是不是串」这一问在打印发射点上只接了一维。与账 #180（句柄被当数字打）同族，也与 §B150 「多维那一路各自挑档」同一个毛病。修法形状 = 打印那一问改问元素的 C 类型（一处，两条路都吃），不要再按维数分叉。
- **顺带（本格没动的两条口径）**：`Erase a` 之后 `UBound(a, 1)` 交 0（VB6 是数组未分配，见 §B154 上面那条 P11）；以及 §B151 那条「自家把多维折成一维」的折叠形状（真语料 `Variant = 二维数组` 0 处，继续只登记不实施）。

### B159 数组那一档里 VB6 有答案、我们故意答不上的格子（2026-10-12 第四十六刀量到并记下；**未开工**）

- **Date 数组**：`mapSaElemType`（`src/backend/expr/cgen_expr_array.cpp:20`）把 Date 折成 `vb6_sa_double` —— 数组里存的确实是 double，这本身按 VB6 没错；但档表对它只能答 8197/"Double()"，而 VB6 答 8199/"Date()"。修法形状 = 枚举里加一枚**真正的 date 档**（步长仍是 8），而不是给档表补一条「声明是 Date 就答 Date」—— 那要把声明表递进 RTL，等于把 §B154 那一刀重抄一遍。
- **UDT 数组**：档表对 `vb6_sa_udt` 回落 variant（8204/"Variant()"，实测就是 `R17-udt` 那一条）。VB6 那一侧「把 UDT 数组装进 Variant」没有可抄的口径：`TypeName` 打的是 UDT 自己的名字，而 UDT 名字从没进运行期 —— 与 §B81/#252「控件创建时没记 VB 名/型」同族，**先要有名字，才谈得上打名字**。
- **Object / LongPtr 数组**：两者都落 `vb6_sa_ptr` ⇒ 8201/"Object()"。VB6 对工程类数组是按对象自己的类型名答（`Dim c() As CommandButton` 打 `CommandButton()`），要的还是元素的名字、不是数组的名字。LongPtr 那一档是 Fix 191 为伪 vtable 特意加进 `vb6_sa_ptr` 的（步长必须是指针宽度），它的 VarType 该答什么本身要先拍口径。
- **从没 ReDim 的数组**：描述符是 NULL ⇒ 回落 8204/"Variant()"（`R17-undim` 钉的就是这一形）。VB6 在这里是运行时错误「Array not even dimensioned」，与 §B154 末尾那条「`Erase a` 之后 `UBound(a,1)` 交 0」是同一族没拍的口径。
- **相邻但不同的一格（响度更高）：数组当 COM 实参交出去压根没有那一档 —— 归因已完成（2026-10-12 第四十七刀）；**已出 = 第五十刀**。
  - **真跑读数**（探针 `.build/k47probe/p47com.bas`，两台逐行相同，`.build/k47run_both.txt`）：拿 `Scripting.Dictionary` 当对端，五形数组实参**全部交回 0 / "Empty"** —— 装箱进 Variant 槽的那一枚（`v = a`）、裸递的 `Long()`、`String()`、多维 `ReDim nd(1 To 2, 1 To 3)`、以及 ND 折成一维指针那一形，一律静默交空。两条证人钉住「读回那一路本身是通的」：`N5 scalar=3 val=42`、`N6 str=8 val=hello`、`N7 count=6`；同一片里 `N8 boxedvt=8209 directvt=8195` ⇒ **槽里那一档第四十六刀已经答对，丢的是出站的表**（`vb6_VariantPackBorrow` 无数组臂 ⇒ `default: VT_EMPTY`）。
  - **语料归因**（`.build/k47_attr7.txt`；脚本 v3→v7 收紧了四轮，每一轮的判据都先用正控验过）：**实参位上递数组**的形状 = 1887（粗筛 按行）→ 1815（要求数组落在那条调用的实参跨度内）→ 1806（撤掉 `HDTF.pszText = StrPtr(Buffer)` 那形属性写左值）→ **1772**（遮掉字符串字面量：`PropBag.ReadProperty("InitBands", 0)`、`Debug.Print "…pal"`、`DDSBack.DrawText … "ARMA "` 匹配到的都是**字面量里的名字**）。
    - ① **in-project 调用 = 1598 行（90%，接收者是本工程类/模块/`Me`）**。再看被调方那一枚形参声明成什么：**1516 行本来就写成 `As T()`**（VB6 递数组的正格），实测发码 = 直调 `vb6_<类>_<方法>(me, &arr)` 对 `T**` 形参 —— 形状本来就对，**不打进本格的靶子**（这一族是 §B160 的对照组，不是它的病人）。
    - ② **真 COM 实参 = 15 行 / 11 份，而且确认都吃 PackBorrow**：`PropBag.WriteProperty("ImageData", m_ImageData)` 那族 **8 行 / 8 份**（`GDI+透明控件` 三枚 .ctl + `vbalImageList` + `ctlNewTab` + `Charts 2020` 的 LabelPlus/FontMemRes），以及 `oPC.AddPaletteColors m_PF_Colors` 那族 **7 行 / 3 份**（`oPC` = `Dim oPC As Object` + `Controls.Add(App.Title & ".ucProgressCircular", …)`）。PropertyBag 那一族不是纸面推测 —— C3 的 PropertyBag 是 `src/rtl/core/vb6forms/uc/uc_propbag.c` 里一枚**原生 IDispatch**，调用点照发 `vb6_ComCall(bag, L"WriteProperty", …)`，臂里直接读 `pd->rgvarg` ⇒ 「设计期存图」那一格今天静默丢负载（写进去是 Empty，读回来是默认值）。
    - ③ **74 行 / 42 份** 接收者是外部类型库对象（DirectX7 / DAO / ADO / DirectInput）—— 形状正是本格，但那批工程吃的类型库 C3 今天不支持 ⇒ **不可达**，按「先证明可达」不入靶子。
    - ④ **假阳性累计约 115 行**（1887→1772），全部由上面那四轮收紧逐轮交出，没有一轮是凭形状猜的。
  - **结论与下一步**：可达性已证、暴露面从「1887 行待归因」收成「**15 行真 COM + 33 行 §B160**（那 33 行后来在第四十九刀分成两档：过程内 Dim 的落 §B160、模块级与形参数组落新格 §B161）」。修本格的出口要新写「`vb6_SafeArray1D/ND` → 真 SAFEARRAY」的换算（RTL 里今天只有 `vb6rtl_paramarray.c` 一处建真数组，而且用完就 `SafeArrayDestroy`），所有权按 §B144 那条口径（交出去的那份归对端 `VariantClear`）—— 面不小；**同族里更响也更便宜的是 §B160**，下一刀先做它。

- **修法前置读数（2026-10-12 第五十刀 = 量，不动代码）**：这一格的两半各有主人，**别一刀顺手猜完**：
  - **出站（我们的槽 → 真 VARIANT）= §B159 本人**：唯一的表是 `vb6_VariantPackBorrow`（`vb6rtl_com.c:575`，`switch (v.vt)` 那一条），   没有数组臂 ⇒ 落 `default: VT_EMPTY`。它只有两条进口，都在同文件：`vb6_ComPackVariant`(:603，客户端实参) 与 `vb6_VariantPackCopyInto`(:623，   服务端返回值)。**回收方已经写死**：`vb6_ComCall` 的发码 epilogue 对每一枚实参副本 `VariantClear`（`vb6com_invoke.c:277-286`，`VT_BYREF` 那一档跳过），   打包用的 `VARIANT*` 结构体随后单独 `free`(:288-297) ⇒ 新建的真数组交出去会被这一趟清掉，对端要留就得自己 `VariantCopy`（加引用）。   串那一档今天就是这么办的（`:612` 与 `:628` 各 `SysAllocString` 一份新副本）⇒ 数组臂照同一口径建，所有权不用新拍。
  - **回程不用新代码**：`uc_propbag.c` 那枚原生 PropertyBag 存的就是**真 Windows VARIANT**（`WriteProperty` → `VariantInit` + `VariantCopy(&it->val, vValue)` :104-105；   `ReadProperty` → `VariantCopy(result, &it->val)` :80），压根不经过我们那枚入站表；而读回来的那一步走 `vb6_VariantFromComResult`，   它**本来就有** `VT_ARRAY` 那一支（`vb6rtl_com.c:330-413`：按 `psa->pvData` 逐元素搬、BSTR 转移指针、VARIANT 递归装箱、末尾 `pv->parray = NULL`    让 `VariantClear` 不动它，档名问第四十六刀那张表）。⇒ **只补出站那一臂，`PropBag.WriteProperty → ReadProperty` 这 8 行就是一条完整的往返**   （写进去真数组、读回来描述符），这一半自成一刀、可验。
  - **入站（真 VARIANT → 我们的槽）另开一格 §B162**：`m_PF_Colors` 那 7 行的接收者 `oPC` 是 `Controls.Add(App.Title & ".ucProgressCircular", …)` 造出来的   **自家 UC** ⇒ 那一趟的对端是我们自己的 IDispatch 服务端，走 `vb6_VariantBorrowIntoSlot`(:640) —— 体里既无 `VT_ARRAY` 也无 `vb6_vtArray`。   那一半还带一条未决口径（借 vs 拷：借来的描述符指向对端的 `pvData`，存活期只有这一趟 `Invoke`），所以分开做、别混在这一刀里。
  - **可达元素的宽度**（15 行的元素档一次数完，判据不用猜）：`m_ImageData()` 六行**全是 `As Byte`**（`GDI+透明控件` 的 QQButton/QQImage/UserControl1 各一份 + 一份旧拷贝），   `m_PF_Colors()` 七行**全是 `As Long`**（`ucProgressCircular.ctl:248`）⇒ 两档都是 **blittable、一维、动态 `ReDim`、非 NULL**。   v1 只接 Byte→VT_UI1 / Long→VT_I4（同宽度的 Integer/Single/Double 顺手，反正表里现成）；**String() / Variant() / UDT 元素的深拷贝刻意不做**，   那一档继续答 VT_EMPTY、另记一格（照 §B144「不知道就别说知道」的口径）。多维真语料 0 处，但 `SAFEARRAYBOUND[]` 本来就吃秩数，两支一起发不额外冒险。
  - **档表仍只问一处**：新建 SAFEARRAY 的 `vt` 只能问第四十六刀那两枚 `vb6_SaVtOfArrayPtr` / `vb6_SaVtFromElemType`（§B155 格①②），   哨兵 A17 已经把「又自己折一档」钉红。动 RTL ⇒ 必须 touch `c3rtl.rc` 再重编 C3.exe（§B156 那条：只改 `src\rtl` 不 touch rc，探针测的是旧 RTL 而且看着完全正常）。

- **已出（2026-10-12 第五十刀 = 出站那一臂 + 发码那一半 + 回程元素表；门 **#566**（run 38237845954，head `82189090`，attempt 1）= 12 job 里 11 绿，唯一红 = 形状门；按 CI 自己那台逐行归因（基线 = CI 对 head `5be7abcb` 那一笔）：本刀只动 **4 行**（Charts 三份工程 +14/+10/+2 字节，与本机「关改道 vs 开改道」`--emit-c` 逐行结果一字对上 = 7/5/1 处打包器改名；另 1 行是本刀改的夹具），另外 13 行是别的分支那 24 格欠的账 ⇒ 一并按 CI 读数登记（`761ce8e7`）；**验证轮 #567（head `761ce8e7`）12 job 全绿**）**：上面那条「修法前置读数」里两条口径都被实测推翻，先订正再记账。
  - **订正①「只补出站那一臂…这一半自成一刀、可验」是错的：还缺发码那一半**。`comPackExpr`（`src/backend/cgen_util_com.cpp` 的 IdentifierExpr 臂）替数组实参挑打包器时问的是 `inferExprType` ⇒ 模块级 `mL()` 发成 `vb6_ComPackInt(mL)`、`mSt()` 发成 `vb6_ComPackBSTR(mSt)`（描述符指针进 BSTR 槽 = 越界读），与 §B161 那形住过的是同一张 switch。这一刀把「这枚实参是不是数组本体」改问**单一权威** `CCodeGen::isArrayCarrierArgName`（`cgen_util_type.cpp` 新函数 = `knownArrays_` 减 `knownObjectVars_`/`knownVariantVars_`/`arrayClassElemTypes_` 减 `cExprIsVariant`；第四十九刀那一站同时改读它），**秩数护栏 `arrayDimCounts_ <= 1` 只加在 COM 这一站** —— `vb6_VariantFromValue` 的 `_Generic` 压根没有 `vb6_SafeArrayND*` 那一档，ND 会落 `vb6_VariantObject` = 对描述符发一次 VT_DISPATCH 的野 Release（§B85 那一族），多维继续归 §B151。
  - **订正②「回程不用新代码」是错的：那张元素表只认六档**。两条进口各自手抄过一份元素 switch（`case VT_BSTR: et =` 那形），认不出的档回落 `vb6_sa_variant`、**步长跟着按 `sizeof(VARIANT)`** ⇒ 一枚真 `VT_ARRAY|VT_UI1` 进来被按 24 字节步长 memcpy = **堆越界读**，标签还答 8204/"Variant()"（Byte 那一档实测）。修法 = 在 `vb6rtl_array.c` 里与 `vb6_SaElemSize`/`vb6_SaVtFromElemType` 同处补**反向表** `vb6_SaElemTypeFromVt`（VARENUM→元素档，认不出答 `vb6_sa_variant` = 改前行为），两份手抄撤掉，`VT_ARRAY|VT_UI1`/`VT_ARRAY|VT_BOOL` 补进两条 case 表 ⇒ 哨兵 A17 现在钉「问反向表 3 处 / 手抄 0 处」。
  - **出站那一臂**：`vb6_ComPackArrayOut`（认 1D 魔数 0x5A1D 或 ND 秩数、要求 `elemSize == vb6_SaElemSize(et)`，`SafeArrayCreate` + `SafeArrayAccessData` + 整块 memcpy）+ `vb6_VariantPackBorrow` 里**数组那一问排在 switch 之前**（改前那枚 `vb6_vtArray|元素码` 的标签对不上任何一条 case ⇒ 一路落 `default: VT_EMPTY`，数组静默变空值）+ 唯一那张可搬白名单 `vb6_ComSaVtBlittable`（bool/int/byte/long/single/double 六档，**出入站同一张表**；串 / Variant / ptr / 货币 / UDT 一律答 0 = 不猜，§B161 记着深拷贝那条没拍的口径）。所有权照上面那条 §B144 的口径，没新拍。
  - **顺带关掉的漏口**：`vb6_VarArith` 为两枚实参新建的那份真数组，在借用契约里禁 `VariantClear` ⇒ 补显式 `SafeArrayDestroy`（否则 `arr + 1` 每跑一次漏一整条 SAFEARRAY）。
  - **判据（两台逐行相同）**：夹具 `tests/test_com_latebound.bas` 九条 `VB-b159-*` + 探针 `tests/tools/disp_probe.c` 那条 `CALL=ECHOARR`（两架构各钉一次，见 §B162）。BASE 读数（`.build/k50probe/run_base_*.txt`）：`l`/`by`/`fx`/`nd`/`va` 交 **3/"Long"**、`st` 交 **8/"String"**、`lo` 交 **0/"Empty"**；修后（`run_fix2_*.txt`）`l=8195/Long()`、`by=8209/Byte()`、`bl=8203/Boolean()`、`sg=8196/Single()`、`fx=8195/Long()`、`lo=8195/Long()` 且元素读回 7100、`st=0/Empty`（白名单外那一档改交空值，不再交指针）。
  - **两条行为负控都真红**（`.build/k50_neg2.py` 撤→重编→复测→按 md5 复原）：白名单里抽掉 `case vb6_sa_byte:` ⇒ `VB-b159-mb` 从 `8209/Byte()/2/66` 变 `0/Empty/-1(e9)` 而其余八条一字不动；抽掉 §B162 那条入站臂 ⇒ 探针整个进程死在 `ECHOARR` 那一趟 `Invoke` 里（0xC000041D，`CALL=ECHOARR` 那行根本没打印，同一片 AMT/QUARTER/ECHO 全正常）。静态那头四种假改动也逐条报红（`.build/k50_negsent_out.txt`：白名单与反向表不对称 / 数组那一问删掉 / 那一问排到已知变量型之后 / 手抄 switch 回来）。
  - **语料 A/B 逐行归因**（关改道 vs 开改道两台 `--emit-c` 对行，`.build/k50_row_attrib_u8.txt`）：三份能编的真工程里动的行 = root Charts **14** / ucProgressCircular **10** / ucTreeMaps **2**，每一行都是同一批调用点上打包器换了名字（`WriteProperty` 的 `m_Caption`/`m_PictureArr`/`c_bvData` + `oPC.AddPaletteColors` 的 `m_PF_Colors`）⇒ 打的正是本格那 15 行里的真 COM 那一族，没有旁落。
  - **刻意留下的一格**：`Dim v As Variant: v = arr`（赋值那一站没装箱）两台仍 **3/"Long"**，读数已钉成 `VB-b159-va`，另开 **§B163**。

### B160 把数组递给一枚 `As Variant` 形参：发码交的是**描述符指针的地址**，被调方按槽去读 ⇒ 静默答 40320 / "Variant"（2026-10-12 第四十七刀的探针撞见；**档① 已出** = 第四十八刀 `cc7a35a9`，门 #558 十二 job 里十一绿、唯一红是形状门那一行登记（`da2868e0`）；**但本格只覆盖了 33 行暴露面里「过程内 Dim 的数组」那一档，语料主体那一档是新格 §B161**）

- **读数**（探针 `.build/k47probe/p47proj/` = `Main.bas` + `Sink.cls`，x64 真跑 `.build/k47run_proj/P47Proj.exe`）：工程类 `Public Sub TakeArr(v As Variant)` 被调用点 `s.TakeArr b`（`b` 是 `Long()`）发成 `vb6_Sink_TakeArr(s, &b)`，而形参的 C 型是 `vb6_VARIANT*` ⇒ 被调方 `VarType(v)` 交 **40320**、`TypeName(v)` 交 **"Variant"**；VB6 那一侧交 8195 / "Long()"。**同一枚探针里的对照组就是正解**：把已经装好箱的 `v = a` 递进同一个形参，被调方答 **8209 / "Byte()"** —— 装箱那一路第四十六刀已经通了，缺的只有「数组实参 → Variant 形参」这一步的装箱。cl 侧只给一条不兼容指针的**警告**：BUILD-RC=0、exe 照跑、一个错误都不打印。
- **暴露面**（`.build/k47_attr7.txt` 的 `callee` 那一栏，判据 = 按实参下标去查被调方那一枚形参的声明）：**33 行 / 12 份**。其中 **26 行是同一枚帮助函数** `VBHacks.IsArrayInitialized(ByRef arr As Variant)`（PhotoDemon 整族在用；它按 `VarPtr(arr) + 8` 去摸 SAFEARRAY 指针，拿到我们描述符指针的地址之后**答「数组没初始化」** —— 不是崩，是整条分支静默走另一头）。**另有 5 行是 `File:=pngBytes` 那形具名实参 ⇒ 我的下标算术吃不下它，这 5 行还没逐行复核**，动针面之前先重数一遍（§B147 那条「认槽」的判据现成）。
- **同族三格别混**：§B155 格③ = 往盒里**写下标**；§B159 = **出 COM** 那张表没有数组臂；本格 = **进过程**那一步没装箱。三格合起来才是「数组当值用」的整条路。
- **档① 的覆盖面订正（第四十九刀量到，别拿这一格当整个暴露面）**：探针矩阵 `.build/k48probe/P49Main.bas` 十六枚调用点逐条归因（发码侧 stderr 探针，三处候选站点全打上标记再跑）⇒ **过程内 `Dim` 的数组**（动态、固定一维、多维 ND 三形）确实落到 `cgen_expr_call_arg_emit.inc` 那条 switch 的 `default`，第四十八刀那条判据打到了；而**模块级数组与数组形参**被 `inferExprType` 答成**元素**档，走的是同一张 switch 的标量那几臂 —— 那几臂把描述符指针直接塞进标量字段，本刀的装箱压根没被问过。⇒ 上面那 26 行 `IsArrayInitialized` 在第四十八刀之后**仍是坏的**，归 §B161。
- **修法（口径先写死再动手）**：数组实参进 `As Variant` / 不带类型那形时，发码改交**一枚临时槽**（`vb6_VARIANT _t = vb6_VariantFromValue(arr); f(me, &_t);`），装箱仍问第四十六刀那枚 `vb6_VariantArray` 载体出口 ⇒ `VarType` / `TypeName` / 元素读回三处一起对上；所有权 = 临时槽与载体都归调用方（借来的视图，与 §B155 那条一致）。**ByRef 回写那一档刻意先不做**：暴露面里 26/33 是只读判断，回写的口径（改完数组要不要搬回调用方描述符）单独定，别在这一刀里顺手猜。
- **门与判据（档①）**：门 #558（head `cc7a35a9`）—— 十二 job 十一绿，唯一红 = 形状门，差异只有 `tests/test_variant_boxing.bas` 一行（就是这一刀新加 b160 四条针的那份夹具），其余 401 行逐行相同；CI 那台交回的 sha256（`b4082bd1…` → `f625bbfa…`）与本机 `.build/C3.exe` 的 `--emit-c` 重算值一致 ⇒ 已按 CI 读数登记（`da2868e0`），登记后复算绿、覆盖哨兵 K1..K3 绿。判据 = `tests/test_variant_boxing.bas` 里 b160 四条（本地动态 Byte / 本地动态 Long / 本地固定 Long / 本地多维 ND）两台逐行相同；哨兵 = `check_sa_access.ps1` 的 A18（第四十九刀已把那条判据的位置也钉进去）。 验证轮与 #561 同批复算（那一笔登记带 `[skip ci]`，没单独跑门）。

### B161 同一形第二档：**模块级数组 / 数组形参**递进 `ByRef As Variant` 时，被按**元素档**分箱 —— 描述符指针塞进标量字段，两台静默答 3/"Long"、17/"Byte"、8/"String"（2026-10-12 第四十九刀；**已出**，门 #560（登记 `f2ebf8bf`）→ 验证轮 #561 十二 job 全绿）

- **怎么撞见的**：第四十八刀收线之前拿真语料那一形做 A/B（`.build/k48probe/P49.vbp` = 一枚 `IsArrayInitialized(ByRef arr As Variant)` + 八枚数组形状），发码里**一处装箱都没有** ⇒ 那一刀没打到病人。于是把三处候选站点（`cgen_expr_call_arg_emit.inc` 的两条 switch、`cgen_expr_call_arg_variant.inc` 的 Fix 113j）都插上 stderr 探针重编一台复算 ⇒ **只有站点 A 那一条 switch 被走到**，113j 与第二条 switch 零命中；臂的选择由 `inferExprType(实参)` 决定。
- **读数**（同一枚探针矩阵，x64/x86 两台逐行相同，`.build/k49base_both.txt`）：
  - 模块级 `Private mLg() As Long` ⇒ 发 `.vt=VT_I4, .lVal=(int32_t)(mLg)`，被调方 `VarType` 答 **3**、`TypeName` 答 **"Long"**（VB6 交 8195 / "Long()"）；`mBy() As Byte` 答 **17/"Byte"**（该 8209/"Byte()"）；`mSt() As String` 答 **8/"String"**（该 8200/"String()"，且那一臂**连 cast 都没有** —— 裸描述符指针进 BSTR 槽）；模块级固定 `mFx(1 To 3) As Long` 同 3/"Long"。
  - **元素读静默答空**：被调方 `v(1)` 交 `VarType=0` / `CLng=0`（不是错误 9、也不 AV）⇒ `IsArrayInitialized` 那一族按 `VarPtr(arr)+8` 摸 SAFEARRAY 指针时拿到的是截断过的整数，一律答「数组没初始化」，整条分支静默走另一头。这一形 cl 也不报错（BUILD-RC=0）。
  - 对照：过程内 `Dim` 的动态/固定/多维三形走 `default` ⇒ 第四十八刀的装箱生效、两台答 8209/8195。
- **修法（收成一处，两处都归它答）**：把「这枚实参是不是数组本体」那一问**提到分档 switch 之前**（判据仍 = 数组声明表 `knownArrays_`，与元素访问 `VB6_SA_AT` 同一张；排除晚绑定对象 / 工程类实例 / 类数组 / 已是槽的表达式），装箱仍走外面那层 `vb6_VariantFromValue`（第四十六刀的载体出口），发在分档之后、吃的是**进 switch 之前那枚原样名**（`arrRaw92y`）⇒ 两条臂一起落到同一枚具名临时槽 `_vb6_arrarg_N`。第四十八刀住在 `default` 里的那条判据同时**撤掉**（同一问不许住两处）。ByRef 回写仍刻意不做（§B160 口径）。
- **修后读数**（`.build/k49fix_both.txt`，两台逐行相同）：`m-long=8195/Long()`、`m-byte=8209/Byte()`、`m-str=8200/String()`、`m-fix=8195/Long()`、`m-elem=8195/Long()` 且 `m-elem-evt=3`、`m-elem-val=210` —— 标签与元素读一起对上；b160 原有四条针不动。
- **判据**：`tests/test_variant_boxing.bas` 加模块级四枚数组（动态 Long / 动态 Byte / 动态 String / 固定 Long）+ 一枚带元素读的针，`tests/run_tests.ps1` 的 b160 名单尾附七条；负控 = 把那条判据的 `> 0` 改成 `> 99` 重编一台 ⇒ 两台退回 BASE 形状（本地那四条也塌回 `16368/Variant()` 一族的垃圾、x86 第二行就断）⇒ 针面真咬着。哨兵 `check_sa_access.ps1` A18 改成钉「一处定义 / 一处使用 / 问数组表一处 / `_vb6_arrarg_` 一处 / `arrRaw92y` 两处 / `argIsArr92y` 零处 / **定义在 switch 之前、装箱在 switch 之后**」，四条假改动逐条报红、复原后绿（`.build/k49_negsent_out.txt`）。
- **顺带记一笔工具坑**（同族第三次的同一个坑）：哨兵里那条「只剥行注释」的正则写成**真换行**时，PS 的 `[^<NL>]` 会把 CRLF 行尾的 `\r` 当作仍在类内，于是每条 `//` 注释连**下一整行**一起吞掉 ⇒ 站点数被压成 0（这次是 `$beAll` 重建时才发现）。写成字面 `\r\n` 才是对的。
- **门**：门 **#560**（head `2e5c07b3`）十二 job 十一绿，唯一红 = 形状门，差异只有 `tests/test_variant_boxing.bas` 一行（正是这一刀加 b161 七条针的那份夹具），其余 401 行逐行相同 ⇒ 按 CI 那台的读数登记（`f2ebf8bf`：`f625bbfa…` → `f7bd0d50…`，与本机 `--emit-c` 重算值一致）；验证轮 **#561**（head `f2ebf8bf`）**十二 job 全绿** ⇒ 登记后的清单与 CI 那台逐行相同。同一轮也顺带复算证实了第四十八刀那次登记（`da2868e0`，那一笔带 `[skip ci]` 没单独跑门）。

### B162 真 SAFEARRAY **递进我们自己的 IDispatch 服务端**时折不回描述符：`vb6_VariantBorrowIntoSlot` 没有数组臂（2026-10-12 第五十刀的量撞见，**已出 = 同第五十刀**，门 #566 / 登记 `761ce8e7` / 验证轮 #567 全绿）

- **读数**：`vb6rtl_com.c:640` 整条体扫过 —— `VT_ARRAY` / `vb6_vtArray` 各 0 处 ⇒ 入站只搬标量档 + `VT_DECIMAL` + 串/对象，数组递进来落它自己的兜底档。
- **暴露面**：`oPC.AddPaletteColors m_PF_Colors` 那 **7 行 / 3 份**（`Charts 2020\ucProgressCircular`，`m_PF_Colors()` = `Dim … As Long` 在 `ucProgressCircular.ctl:248`）  —— 接收者是 `Controls.Add` 造出来的自家 UC ⇒ 必走服务端那一半。与 §B159 是同一条往返的两位主人，**别混刀**（§B160/§B161 就是这么分开的）。
- **未决口径（开工前先拍）**：借 还是 拷 —— 借的那一形描述符直接指对端的 `psa->pvData`，存活期 = 这一趟 `Invoke`（与 `BorrowIntoSlot` 「谁也不接管」的契约一致，  也与 §B155 的载体口径最近），但服务端过程里一编辑元素就是写回对端的数组，且 `VT_ARRAY|VT_BSTR` 借过来编辑串会越界；拷的那一形要有人拆（建一份自己的描述符 + 逐元素搬，  回程 `VariantClear` 归对端）。VB6 语义是数组按引用传（服务端改元素，调用方看得见）⇒ 借才是对的一档，但要钉死「blittable 才借」。

- **已出（2026-10-12 第五十刀，与 §B159 的出站同批；门 **#566**（run 38237845954，head `82189090`，attempt 1）= 12 job 里 11 绿，唯一红 = 形状门；按 CI 自己那台逐行归因（基线 = CI 对 head `5be7abcb` 那一笔）：本刀只动 **4 行**（Charts 三份工程 +14/+10/+2 字节，与本机「关改道 vs 开改道」`--emit-c` 逐行结果一字对上 = 7/5/1 处打包器改名；另 1 行是本刀改的夹具），另外 13 行是别的分支那 24 格欠的账 ⇒ 一并按 CI 读数登记（`761ce8e7`）；**验证轮 #567（head `761ce8e7`）12 job 全绿**）**：口径拍成**拷**，理由是实读出来的，不是挑的。
  - **改前不是「静默空值」而是崩**：撤掉那一臂再跑探针，进程在 `ECHOARR` 那一趟 `Invoke` 里返回 0xC000041D，`CALL=ECHOARR` 那行根本没打印出来（同一片里 AMT/QUARTER/ECHO 全正常 ⇒ 崩在这一格自己）。于是本格从「可押后」变成**本批的承重墙**：§B159 的发码改道若单独上线，语料那 7 行 `oPC.AddPaletteColors` 就从「静默空值」升级成「崩」。
  - **为什么不能借**：服务端这一趟结束就 `CoTaskMemFree(coercedArgs)`（`vb6comserver_obj.c`，槽阵列是 calloc 的、逐槽清空后释放），而 VB 代码会把这枚 Variant 存进变量、带出这次调用 ⇒ 借来的 `psa->pvData` 在那之后就是悬垂的。
  - **落的形状**：`vb6_VariantArrayIntoSlot`（白名单与出站**同一张表**；元素档认不出、认出来的档与真数组的 vt 不是同一档、或秩数 >1 ⇒ 一律交回 `vb6_vtEmpty`）+ 导出的 `vb6_VariantSlotRelease(void*)`，在服务端那句 `CoTaskMemFree` 之前逐槽调用。`vb6_VariantClear` 不碰数组标签（第四十六刀定下的非拥有先例）⇒ 不会双释放。
  - **VB6 的「数组按引用传」那一半刻意不做**：回写与 §B160/§B161 的 ByRef 回写是同一条口径，补的时候两头一起补。
  - **真跑读数（两台同形，已进门）**：`CALL=ECHOARR hr=0x00000000 resultvt=8195 arrvt=8195 dims=1 lo=0 hi=2 arrsum=60` —— 一枚真 `VT_ARRAY|VT_I4`（元素 10/20/30，界 0..2）进服务端、被 `EchoVal` 原样交回、客户端按真 SAFEARRAY 读出界与元素和。`arrsum` 只能来自入站那份**拷贝**的元素（出站那一步是新建 SAFEARRAY + 搬自己的 `data`）⇒ 这一条同时是「描述符真的可用」的证人，不只是「指针活着回来了」。

### B163 `Dim v As Variant: v = <数组名>`：赋值那一站没装箱，槽读回来是 3 / "Long"（VB6 交 8195 / "Long()"）（2026-10-12 第五十刀的量撞见；**第五十一刀已出** = 提交 `0a7f508c`，门 **#568**（run 38250433179，head `0a7f508c`，attempt 1）= 12 job 里 11 绿、唯一红 = 形状门（就是本刀那一行登记）⇒ 按 CI 那台读数登记 = 提交 `411aa223` ⇒ 验证轮 **#569**（head `411aa223`）attempt 1 唯一红 = `Tests (vbp #4)` 的 `ctrldlg_probe … FAIL (run error)` —— 同一片在 #568 刚绿过、这一笔只动了期望清单没动任何测试输入 ⇒ 墙钟余量那一族（§B140 给的那枚 `DlApp = 150` 具名预算，与门 #484 同一形状）；**attempt 2 复跑后 12 job 全 completed/success**）

- **读数**（`.build/k50probe/run_base_*.txt` 与 `run_fix2_*.txt`，两台逐行相同）：模块级 `Private b159L() As Long`（已 `ReDim`、元素有值）赋给 `b159Va As Variant` 之后把这枚 Variant 递过 COM，对端读回 **3/"Long"**、`UBound` 交错误 9 —— 第五十刀把出站的表补好了，可这一形在**发码**那一步就把数组折成了元素档，压根没带着数组标签走到表跟前去。同一份探针里的 `nd`（多维）也是 3/"Long"，但那是 §B151 的形状，别混进本格。
- **与同族的关系 + 站点已量到（2026-10-12 第五十刀收线那一趟顺手量的，探针 `.build/k51probe/`）**：§B160/§B161 修的是**实参位**（发码在调用那一站装箱，判据 = `isArrayCarrierArgName`）；本格是**赋值**那一站，站点 = `CCodeGen::wrapVariantValue`（`src/backend/cgen_util_ctrl.cpp:2336`，里面那张只留四档的小表问 `inferExprType`，而 `inferExprType` 对数组名答的是**元素**档 ⇒ 与 §B161 同一个病、不同的一站）。**上面那句「赋值走的是 AssignMove 那一路」是猜的，实测不作数**：这一路今天就是走 `wrapVariantValue`。
- **发码读数**（`.build/k51probe/p51_emit.txt`，x64、RC=0，四枚对照）：模块级 `Private mL() As Long` ⇒ `v1 = vb6_VariantLong(mL);`；模块级 `mB() As Byte` ⇒ `v2 = vb6_VariantByte((uint8_t)(mB));`（**整枚描述符指针被截成一字节**）；过程内 `ReDim loc(0 To 1) As Long` ⇒ `v3 = vb6_VariantLong(loc);`；对照组 `v4 = 5` 那条一字未动 ⇒ 这一站对非数组是对的，缺的**只**是数组那一问。
- **修法形状（按读数收窄，不用再造一站）**：全仓 13 处引用（声明 + 定义 + 约十处调用）⇒ 赋值/形参/属性返回那几路**已经只走这一枚出口**，不像第四十九刀那条内联 switch；所以修法 = 在 `wrapVariantValue` 开头问同一枚 `isArrayCarrierArgName`、装箱仍交 `vb6_VariantFromValue`（第四十六刀的载体出口）、秩数护栏照旧 `<= 1`（ND 仍归 §B151，同一理由：`_Generic` 没有 ND 那一档）。哨兵那头 A18 现在钉的是「这条权威被问两次」，落地之后要改成**三次**并逐形归因（别把站点数当免费判据白送）。
- **判据已经先钉上了**：`VB-b159-va=3/Long/-1(e9)` 钉的是**还没修**的读数 —— 谁修了它这一行会红，那时按新读数重钉并回来本格收线。同一形状旁边另有 `mu`（从没 ReDim = NULL 载体）与 `ms`（String() = 白名单外）两条，那两条读数是对的，不许顺手改。
- **覆盖面已数（2026-10-12 第五十刀收线那一趟顺手做的，屏具 = `.build/k51_census.py`，读数 = `.build/k51_b163_census.txt`）**：真语料 `D:\vb_yqt4qPac` 的 10110 份 .bas/.cls/.frm/.ctl/.pag 里 **命中 36 份 / 41 行**（判据 = 一枚 Variant 档的名字被赋以同文件声明的**裸数组名**）。逐行看：绝大多数是**函数返回变量**那一形（VbEclipse 那族 20 份 `IIterator_items = m_ArgumentList` 之类 + SQLiteCursor/SQLiteDataSet 的 `GetRows = ArrRows` 4 行 + 三份 cRegistry 的 `vValue = abData` + 三份 MListViewEx 的 `LVGetColOrder = cols`），少量是普通赋值 ⇒ 本格**不是纸面缺口**（与 §B151 那形「命中 0」不同）；但**语料主体走的是返回那一站，不是 P51 探针那种局部赋值**。
- **两站同一问，覆盖面不一样（`loc` 那一形已被订正，见下）**（探针 `.build/k51self/P51B.vbp`～`P51I.vbp`，全部 `--emit-c` 实测、RC=0）：**返回站** `RetLocal = a`（过程内 `Dim a() As Long`）已经发成 `vb6_VariantFromValue(a)` —— 第四十八刀那批装箱吃的正是它；而 `RetModule = mArr`（模块级数组）仍发 `vb6_VariantLong(mArr)`。**赋值站 = `wrapVariantValue`（`cgen_util_ctrl.cpp:2336`，全仓唯一真实调用点在 `cgen_assign_value_sem.inc:305`）**：普通名字的过程内数组在这里也是对的（P51F/P51G：`zv = vb6_VariantFromValue(z)`、`bv = vb6_VariantFromValue(b)`，Sub/Function、启动与否都无关），**模块级数组在这里折成元素档**（P51E/P51H：`v1 = vb6_VariantLong(mArr)`）。
- **`v3 = vb6_VariantLong(loc)` 的真正原因是撞名，不是「过程内数组在赋值站也折」**（上一条最初那么写是**错的**，已按 P51E→P51H 的改名复算订正）：同一份语句只把数组名从 `loc` 改成 `loc2`，发码立刻从 `vb6_VariantLong(loc)` 变成 `vb6_VariantFromValue(loc2)` ⇒ `Loc` 是 VB 的内在函数（返回 Long），`inferExprType` 先答了内在函数的档。⇒ 本格有两个成因、同一站：**① 模块级 / 形参数组**（§B161 那一族）与 **② 名字撞上内在函数的数组**；修法都是「把 `isArrayCarrierArgName` 那一问提到 `inferExprType` 之前」，一问治两个成因。
- **两枚已上线的站点不受这个撞名影响（P51I 实测，同一份 `Dim loc() As Long`）**：第四十九刀那一站发 `vb6_VARIANT _vb6_arrarg_0 = vb6_VariantFromValue(loc);` 再 `vb6_TakesV(&_vb6_arrarg_0)`，第五十刀那一站发 `vb6_ComCall(o, L"Add", {vb6_ComPackBSTR(…L"k"), vb6_ComPackValue(loc)})` —— 因为 `isArrayCarrierArgName` 问的是 `knownArrays_` 那张**名字表**，不是 `inferExprType`。⇒ 这条也是「单一权威」那次收成的红利：新站照抄同一问就一起免疫。
- **屏具那条坑（同族第四次，值得记）**：这条屏第一次跑打印的是「命中文件 0 / 命中行 0」—— 而**它是假的**：`RE_ASSIGN` 只认 `:` 注释，VB 源里行尾注释绝大多数是 `'` 起头 ⇒ 每一条带注释的赋值都静默不匹配。第二版加了内建正控（一份合成文件必须恰好命中那两行、元素读与标量那两行不许命中），第一次跑就把这个假零顶红了 ⇒ 修完才是上面那 36/41。**census 之前先给屏喂一份已知答案的样本**，别拿「0 命中」当结论 （与 §B151 那条 0 命中不同：那一条当时也这么读，但正控证明过屏能命中别的形状）。
- **与 §B149/§B150 的关系**：那两格讲的是**晚绑定槽元素**读写，与本格的整枚数组赋值不同形，别并做一格。

- **落地**（一处发码 + 一条顺序判据 + 一枚新哨兵）：`wrapVariantValue` 在 Fix 170 那个「整枚数组引用」块之后、`inferExprType` 那张表之前问同一枚 `isArrayCarrierArgName`（`src/backend/cgen_util_ctrl.cpp:2371`），命中且 `arrayDimCounts_` 给的秩数 `<= 1` 才交 `vb6_VariantFromValue(cExpr)`；ND 仍归 §B151（`_Generic` 没有 ND 那一档 ⇒ 落 `default` = `vb6_VariantObject` = 对一枚描述符发 Release，§B85 那一族）。装箱没有另造一站 —— 这一站全仓 13 处引用本来就是赋值/返回几条路的共同出口。
- **上面那句「两站」要订正成「一站两形」**：函数/属性**返回变量**那一路（发码里的 `vb6_ret_<名> = ...`）吃的就是同一枚 `wrapVariantValue` —— VB6 的 `F = arr` 本来就是给返回变量赋值，`cgen_assign_value_sem.inc:305` 那唯一一处真实调用点两条路共用。真语料 A/B 动的那 5 行全是 `vb6_ret_` 那一形，就是这个原因。
- **判据**：`VB-b159-va` 从 `3/Long/-1(e9)` 重钉成 `8195/Long()/3/7002`（与 `ml` 同形），x64+x86 两台各 9/9 命中、其余八条针逐字节未动；`.build/k51_check.py` 那九份探针 RC 全 0 ⇒ `v1/v2/v3/vb6_ret_RetModule/mv` 全改走 `vb6_VariantFromValue`，而三条标量对照组（`v4 = vb6_VariantLong(5)`、`vb6_ret_FLocal = vb6_VariantLong(1)`、`fv = vb6_VariantLong(2)`）一字未动。
- **两条负控都真跑了**（不是预告）：① 秩数护栏 `<= 1` 改 `<= 0`（等于关掉这一问）⇒ 冷编一台真跑，`va` 退回 `3/Long/-1(e9)`，其余八条不变（`.build/k51_negbeh.py`，还原按 md5 核过、重编后两台复验）；② 把那一问挪到 `inferExprType` 表之后 ⇒ 哨兵 A18 新加的「赋值站顺序」那条当场红（exit 1，`.build/k51_negctrl.py`，undo 后 md5 对上）。
- **哨兵跟着改**：`check_sa_access.ps1` 的 A18「问那条出口的地方」2→3，并新增一条**顺序**判据（赋值站那一问的锚点必须排在 `Vb6Type vtype = inferExprType(...)` 之前，锚点找不到也报红）；55 道哨兵在最终树上全绿。
- **覆盖面 = 同一台机器、只换 C3.exe 的 A/B**（第四十九刀那条教训的正面执行，不拿自己的夹具当语料）：
  - 清单那 402 行登记语料（`.build/k51_base_C3.exe` vs `.build/k51_new_C3.exe`，两份清单逐行按末段 relpath 对）⇒ **只动 1 行** = `tests/test_com_latebound.bas`，`bytes 10519→10524`（+5 正好是 `vb6_VariantLong(`→`vb6_VariantFromValue(` 的字数差），其余 401 行逐字节相同。
  - 真语料 `D:\vb_yqt4qPac` 那 36 份命中文件全部 `--emit-c` 两台对打（`.build/k51_corpus_ab.py`）⇒ **动 4 份 / 5 行**：`GetWindows = WindowName`、`IIterator_items = m_EnumList`、`Lines = m_Lines`、`Values = m_Values` + `Keys = m_Keys`。五条 BASE 全是 `vb6_Variant{Long,String}(载体)`、NEW 全是 `vb6_VariantFromValue(载体)`，载体类型逐条核过都是 `vb6_SafeArray1D*`（`_Generic` 认那一档 ⇒ 不是 §B85 那种野 Release 形状），**零条旁溢**。
- **H6 那一格是本刀自己撞出来的坑，单独记**：改针面时把新写的注释拼到最后一行，结果整条 `Add-BasTest "test_com_latebound" ...` 被拼在 `#` 那行的行尾 ⇒ 那枚夹具**从注册表里整个消失**，而 `[Parser]::ParseFile` 报 0 错、当时 55 道哨兵**无一红**（只有 `Add-BasTest` 的调用数从 114 掉到 113，没人盯它）。补了 `check_test_helper_integrity.ps1` 的 **H6**：① 剥掉注释后的 token 文本里「名字 + 空白 + 引号或括号」的调用形次数必须等于 AST 找到的调用点数（少了 = 有一条被注释吃掉），② 每个调用点必须是它那一行第一个非空 token。负控跑过：把那一行重新拼回注释行尾 ⇒ `FAIL H6 Add-BasTest: 剥注释后的调用形 114 处, AST 调用点 113 处`，还原后绿。
- **刻意不动的那两条仍在门上**：`mu`（从没 `ReDim` = NULL 载体）与 `ms`（`String()`）读数一字未改 —— `String()` 那一形现在在赋值站也发 `vb6_VariantFromValue`（真语料那三行 `m_Lines/m_Values/WindowName` 就是它），但过 COM 出站仍按 §B161 的白名单拒发 ⇒ 针面上的 `0/Empty/-1(e9)` 是 §B161 那条还没拍的深拷口径，不是本刀欠的。
- **【2026-10-13 第五十三刀：「往里读」那一族的两遍实现合一 + census 从 13 张补到 19 张】**（本刀是**结构刀**：两台行为读数一字不动，判据在护栏那一头；门 **#572** = run 38287649438、head `5016037e`、attempt 1、逐 job 复核 12/12 completed/success —— 含形状门 `Emit manifest (shape oracle)` 直接绿 = 登记清单零移动，与本机 402/402 的复算对上）
  - **落地三处**：① `vb6rtl_com.c` 里 `vb6_VariantFromComResult`(134 行) 与 `vb6_VariantFromStackVARIANT`(121 行) 那两张逐条相同的 23 档表并成一枚 `static` 核 `vb6_ComVariantIntoSlot(VARIANT* pv, int steal)`，两枚公开名留成薄壳(发码仍按名字发射 ⇒ `--emit-c` 的形状一字不动)；所有权那四处差异(BSTR 偷/拷、DISPATCH 置 NULL/守卫 AddRef、SAFEARRAY 的 BSTR 元素偷/拷、`pv->parray` 动不动)全部改吃 `steal` 一枚参数。② **数组性改成问旗标**：`pv->vt & VT_ARRAY` 一面旗，不再是八行 `case VT_ARRAY|...` 名单 —— 名单外那一档(`VT_ARRAY|VT_DATE` 一类)以前落进 `default` 的数组块，块尾还 `pv->parray = NULL` 把原值抹掉，而那块底下换 BSTR 的兜底**两枚双胞胎都从来没走到过**；现在未列的标量档统一走那一格。③ `vb6_Host_ToWinVariant` 撤掉自己那张十档表，改问 §B145 的 `vb6_VariantPackBorrow` + 宿主管道要的那一步(`SysAllocString` 串副本)；**对象那一档刻意不 AddRef**，与改前一字不差 (反向的 `Host_ClearVariant` 只放串，这里给一次 AddRef 就成双份)。顺带 `vb6_CallByName` 的结果表(八档 + `default: result.lVal = retVal.lVal`)撤掉，改问薄壳 `vb6_VariantFromStackVARIANT` —— 那一格的 Date/Currency/I8/Decimal 以前是「标签照抄、数值只搬 32 位」。
  - **census 补认**：`check_variant_i8_coverage.ps1` 的站点正则从「只认接收者叫 `v`」放宽到**任何拼法**(`p->vt` / `in->vt` / `pv->vt` / `varElem.vt` / `retVal.vt` / `expr.vt` / 裸 `comVt`)，并把 `detail/*.inc` 一起扫、函数体里的 `#include "….inc"` **就地展开**(Format 那张表就住在 .inc 里，不展开就会红在「这一档没写」上)。站点 **13 → 19**，九张隐身的逐张归因：新增 must `vb6_ComVariantIntoSlot` + `vb6_Format`(两档都答)，新增 exempt `vb6_EnumNext`(全仓 0 调用者的死码)、`vb6_SaElemTypeFromVt`(元素档表不是槽表)、`vb6_Host_FromWinVariant`(见 §B165)、`pb_Read`(设计期属性包，I8/Decimal 无写者)；V1 下限 13 → 19，must 名单下限 11 → 13、decMust 8 → 10。新增 **C1**：两枚薄壳的体里再出现 `switch` 就红，核里那枚 `switch` 恰好一处 —— 钉的就是「同一问不许再抄第二遍」。
  - **两条负控各点各的红**(`.build/k53_nc.py`)：把 `VT_DECIMAL` 那一档从 Format 的 .inc 表里撤掉 ⇒ `FAIL D1/D2 vb6_Format ...`（**改前这一撤无人报错，那正是本刀要堵的洞**）；往薄壳 `vb6_VariantFromStackVARIANT` 里塞回一枚 `switch` ⇒ 同时红 `V2/V3`、`D1/D2` 与 `C1`。还原按 md5 核过(`48002cd2…` / `e96df0ea…`)，还原后 PASS。
  - **护栏**：55 道哨兵全绿(`SENTINEL-RED=0`，含改后的 A17：COM 那两条进口并成一枚核之后，问档表的地方 6 → 5、问反向表 3 → 2，两条都在规则里写明了归因)；定点行为两台真跑 `test_com_latebound` 15 枚针全中(`miss=0`，含 §B159 那八条数组往返针) + `test_variant_boxing` 82 枚针全中；探针 `.build/k53_p/probe.bas` 改前改后两台 12 行读数逐字节相同 = 结构刀的行为零变化。
  - **本刀试错的一条(记下来，别再来一遍)**：数组臂上试过「问 `SafeArrayGetElemsize` 拿对端宽度、与尺寸表不符就逐元素装箱」—— 结果 `VB-b159-ml` 的元素从 7002 变 **3**、`mb` 从 66 变 **0**，且随后 `0xC0000374`(堆写坏)：那张宽度答案与我们的尺寸表打架时，逐元素那一支写的是 24 字节的**槽**，而描述符按 4 字节步长分配。已**整支撤掉**、恢复成原来那条 memcpy，越界读那一格照旧开着并写明(§B159 家，宽度这一问要先量准再动手)。今天的实测：认不出档的元素进来仍是越界读。
  - **§B142 剩下什么**：`vb6_Host_FromWinVariant`(往里读那一族的宿主存值腿)与它连着的四站 → **另立 §B165（第五十四刀已出）**；`vb6_EnumNext` / `vb6_EnumRelease` 两枚 0 调用者的死码撤不撤 → 等用户拍。本刀的针面因此**没有新增行为针**(结构刀的判据在哨兵与零变化上)，§B141/§B142 那批 Decimal 针原样不动。


### B164 一条语句里没被赋给任何变量的带权结果，从来没进过回收的名单—— 表达式临时串、Collection 取回的临时槽、For Each 的循环变量各漏一只（2026-10-13 第五十三刀的量撞见；**形④ 已出 = 第五十五刀（两半同批：For Each 那一站 + 普通赋值那一站的「槽到槽」）** —— 「把一枚值收进 Variant 槽」收成一枚出口、两半句一起答；①②③⑤ 那几形缺的仍是口径不是手法）

- **量法**：判据问**增量**不问绝对值（沿用 `tests/test_str_leak.bas` 那台引擎 = `kernel32!K32GetProcessMemoryInfo` 的 `PagefileUsage`），并且**同一形连跑三轮看有没有平台**—— 有平台 = 堆预热，不是漏。探针与读数都在 `.build/k53_p/`（`leak5.bas`～`leak8.bas`，x64 本机，除注明外均为每次一趟 200k 或 400k 次访问）。
- **五条读数**：① `n = Len("ab" & "cd")` 三轮 = 24,182,784 / 24,018,944 / 23,949,312（每轮 200k 次 ⇒ **120 B/次，三轮不减** = 真漏，拼接那枚临时 BSTR 没人 `SysFreeString`）；② `n = Len(gArr(1) & "!")` 200k = 30,003,200（同形，换一条读数来源）；③ `n = VarType(gColl(2))` 10k = 1,257,472（**126 B/次**，Collection 取回那枚带权副本 = `vb6_coll_itemCopy` 的 `SysAllocString`，当它只作表达式用时没人回收）；④ `For Each e In gColl` 500×40 = 2,301,952（**这条 115 B/元素 是被形⑤污染的读数** —— 那一版的计时窗里循环体还调了一句 `Len(e)`，量进来的是「本站的漏 + 换算临时串」两笔钱。第五十五刀把窗体改成一句都不分配(只问 VarType)之后重量 = **81/71 B/元素**，见下面那条落地)；⑤ `n = Len(v)`（`v` 是数值 Variant）400k = 13,877,248（**35 B/次**，换算出口的临时串）。
- **对照组同样量了，而且都是零**（这条决定修法不是「补一个 free」）：`v = "abcdefgh"`、`v = RetStr()`、`v = RetVar()`、`v = gColl(1)`（各 400k / 10k 次）全部 0 增量 ⇒ **赋给 Variant 槽那一路是通的**（发码在赋值前发 `vb6_VariantClear(&v)`，见 `.build/k53_p/leak6.c` 的 189/209/229/249 行）；`For Each e In gArr`（VB 数组，不经 VARIANT）也 0。⇒ 漏的三形有一个共同点：**结果没落到任何变量上，或者落到的那一站自己抄了裸赋值**。
- **归因（发码逐条读出来的，不是推的）**：机制在产物里已经有两枚现成的出口 —— `vb6_ComVarFree((void*)vb6_ComCall(...))`（语句丢弃 COM 结果那一形，`leak6.c:173`）与 `vb6_VariantClear(&v)`（Variant 赋值那一形）。所以本格不是「没有回收机制」，而是「**一条语句里没被赋给任何变量的带权结果」这一形从来没进过发码的回收名单**。For Each 那一站是第二个成因、同一条决定：`src/backend/stmt/cgen_control.cpp:434` 发的是一句`e = vb6_VariantFromStackVARIANT(&fe_var);`（裸结构赋值，前面没有清旧值那一步），而普通赋值那条路有 —— **同一个决定在两处写法不同形**（§B163 那条「一站两成因」、§B157 那条「按名字当键必须带全身份」是本族第三次撞）。
- **为什么这一格先不修（缺的是口径，不是手法）**：两条路都能走，代价不同，要拍一条 —— ① **语句级回收**：发码把本条语句里造的带权临时量登记成一份名单，语句尾按名单发 free；改动全在发码，覆盖面 = 一切表达式形状，产物行数会动（要走一轮清单登记）。② **运行期 arena**：`vb6_BSTR_Concat` 那一族出口改从「每条语句复位」的堆上分配，语句尾整块 free；改动在 RTL + 一份出口名单，一次治全部，但前提是先把「哪些出口交的是带权串」数清楚（`scripts/check_builtin_slot_returns.ps1` 那张表就是底子，§B146 收过一半）。两条都不许动坏既有那条「赋值目标的所有权」（账 #119 的 `AssignMove`）。
- **开工前的站点账（2026-10-13 第五十三刀收线那一趟补的量，屏具 = `.build/k54_census.py`，读数 = `.build/k54_census.txt`）**：`src/backend` 那 122 份 .cpp/.hpp/.inc 里，**交出带权值**的出口共 14 枚、发码站点 **715 处**（前五 = `vb6_ComCall` 202/46 份、`vb6_VariantFromComResult` 112/30、`vb6_VariantToString` 68/21、`vb6_ComGetProp` 70/28、`vb6_BSTR_FromStr` 58/24）；**回收那侧**只有 **40 处**（`vb6_VariantClear` 15/7、`VariantClear` 13/7、`vb6_ComVarFree` 4、`vb6_ComVarClear` 2、`vb6_FreeANSI` 4、`SysFreeString` 2）。⇒ 差一个量级；而且这只是**站点数**，一处发码在循环里对应的是无限多次运行。**匹配口径要写明**：这一笔按**逐词 `\b`** 数，先前顺手那枚「293」是**子串**匹配（把 `vb6_ComCallByDispid` 一起算进去了）⇒ 作废。
- **① 那条路不是新机制，仓里已经有整形的先例**：M22 的 `ansiTempsToFree_` —— 名单住 `detail/util/cgen_state.inc`（一枚 `std::vector<std::string>`），登记两处 `push_back` 在 `detail/expr/cgen_expr_call_arg_emit.inc`，**每条语句尾**统一发 `vb6_FreeANSI` 再 `clear()` 在 `stmt/cgen_stmt.cpp:126-129`（那行注释写的是「每条语句后统一清理, 无内存泄露」），过程入口/出口另各 clear 一次 = `decl/cgen_decl_{func,proc,prop}.cpp`。照这一形扩一枚 `ownedTempsToFree_` 就是 ①，**但名单必须带档**：实测五形里 ①②⑤（拼接与换算的临时串）交 `SysFreeString`，③④（Collection 取回的槽、For Each 的循环变量）交 `vb6_VariantClear` ⇒ 一份名单两种 free，这正是 §B142/§B144 「所有权要跟值类型走」那一问的又一形，别只发一种。
- **语料暴露面（屏具 = `.build/k54_corpus.py`，读数 = `.build/k54_corpus.txt`，屏自带内建正控）**：真语料 `D:\vb_yqt4qPac` 那 10138 份 .bas/.cls/.frm/.ctl/.pag 里 **`For Each` 2107 行 / 769 份**，**Collection 声明 550 枚 / 330 份** ⇒ 形④ 不是纸面缺口。「这 2107 行里有多少遍历的正是 Collection」要类型解析才知道，本格没数，别把它读成 2107 次泄漏。**屏具那条坑（同族第五次，值这个钱）**：第一次跑打印的是「Collection 声明 183」—— 那是**假的**，正则写成了 `(?:As\s+New\s+)?Collection`，只认带 `New` 的那一形，`As Collection` 整族隐身；内建正控（一份合成样本必须恰好命中 3 行 `For Each` 与 2 行声明）第一次跑就把这个假低顶红了 ⇒ 改成 `As\s+(?:New\s+)?` 才是 550。**census 之前先给屏喂一份已知答案的样本**，别拿低命中当结论。
- **判据已经想好**（开工时照这个钉，别再新造引擎）：上面 ①②③④⑤ 五形各一针，问**每形每次的增量上限**（改前实测 ~120/126/115/35 B 每次，改后要求 < 8 B），外加一条负控：把新增那行回收注释掉 ⇒ 针必须红。对照五形（`v = gColl(1)` 等）读数不许变，它们是「别把赋值那一路改坏」的证人。
- **第五十五刀落地（形④ 已出：#577→登记 `e75692dc`→#578 一条真红→`b708ee8e`→#579→登记 `f3859efb`→验证轮 #580 十二绿）**：「把一枚值收进 Variant 槽」= **清旧 + 交接新**，两半句是一个决定，收成 `CCodeGen::emitVariantSlotStore(target, expr, rhsOwnsValue, viaTemp)` 一枚(声明 `detail/util/cgen_helpers.inc`、定义 `cgen_util_ctrl.cpp`)。
  - **两半句都在本轮现过形**：只搬不清 = 每遍历一个元素把上一轮那只 BSTR 忘了(BASE x64 两次 1,642,496 / 1,622,016、x86 1,425,408，÷ 20,000 次 ⇒ **81/71 B/元素**；补上清之后两台 = 0 / 4,096)。只清不拷 = 把**别人还活着**的串 free 掉：数组那一站发的是 `vb6_VariantFromValue(VB6_SA_AT(BSTR, ...))`，`VB6_SA_AT` 是裸读、`_Generic` 对 `wchar_t*` 只把指针装进槽 ⇒ 中途那一台(只补清、数组站照旧)实测 `arr-sum` 从 **820 变 813/815**(数组自己的元素被回收后又被同尺寸分配复用)。负控 NCB(数组站按「自有」发)两台**当场 0xC0000374、一行都不打印**(`.build/k55_ncb_arr.txt`)。⇒ 与 §B165 那条「所有权必须成对」是同一条纪律的另一面，而且这次是**本刀自己把洞打活的**，不是现成的红。
  - **出口答两路 + 一路顺序**：`rhsOwnsValue=true`(右值是这一句刚 mint 的带权值: `vb6_VariantFromStackVARIANT` / `wrapVariantValue` 的内联临时值) → 清 + 直赋；`false`(右值是借来的) → `vb6_VariantAssign`(清 + 深拷贝，值档住第五十四刀那对出口)；`viaTemp=true`(右值里出现目标自己 `v = CDate(v)`: 先清会把右值要读的自己清成 0) → 先收栈上临时量再走同一对清+搬，**发码逐字节与第五十四刀那台相同**。
  - **六处站点接线**：`cgen_control.cpp` For Each 四条槽位路(站① ParamArray / 站② VB 数组(借) / 站③ COM 集合两臂) + `cgen_assign_value_sem.inc` 普通赋值两形(站④ 直发 / 站⑤ 自指右值) + `cgen_assign_stmt_special.inc` FnRetUdt 的 UDT Variant 成员(站⑥ —— **这一站是哨兵把「手抄那一对」的匹配放宽之后自己撞出来的**：它在出口外自己抄那两行，接线时发码沿用清+搬 = flag true，零行为改动)。
  - **判据**：新夹具 `tests/test_fe_slot_leak.bas` 两台各 **11 针同过**(`$b164Needles`，针面从真读数里取、不手数)；`api=1/1/1/1` 是「四次读数都真取到」的正控(缺它 delta 可能是自己减自己的假零)；两个计时窗刻意让循环体一句都不分配，否则量进来的是形⑤、两笔账互相顶包；`held` 那枚证人刻意用 `s = CStr(e)` 而不是 `held = e` —— 后者是 Variant→Variant 的浅拷贝，会和槽共用缓冲区(另记一条开放项)。发码形状针三行 = 集合站的清+搬 + 数组站的 `vb6_VariantAssign`。
  - **哨兵** `scripts/check_variant_slot_store.ps1`(第 56 道)：R1 出口三形(6 枚 emitLine + 必含 `vb6_VariantAssign` 那一路 + 不许 switch) / R2 声明与定义各一处、调用点 ≥7 / R3 「清+搬」那一对的手抄形状只许住出口那份文件、恰好两形(先后顺序也由这条钉) / R4 **还有哪些站在裸发槽位赋值 = 钉死名单**(新增即红，修掉一处必须同步把名单改短)。读数 `def 1 | calls 7 | shapes 2 | open stations 3 | violations 0`；五条负控各红对应一条(全部在 `src/backend` 的**副本**上做，真实树一个字节不动)：站点手抄→R3、撤一处调用→R2、名单外新站点→R4、出口丢借用那一路→R1、出口截半→R1+R3。
  - **第二半句在站④ 也欠着(门 #578 抓出来的，见下面那条落地②)**：`last = e2` 两边都是 Variant 槽时，`wrapVariantValue` 发的是 `vb6_VariantFromValue(e2)` —— 对 vb6_VARIANT 是 _Generic 的 identity **浅拷贝** ⇒ 两枚槽共用同一只 BSTR / 同一个 x86 DECIMAL 缓冲(Rev16 那格数组元素的同一个成因)。本刀给 For Each 补上「先清旧值」之后，这份共用从「安静地别名」变成**二次 free**：实测 `test_host_decimal` 的 x86 台 12 次跑里 **7 次 0xC0000374**，而 BASE(第五十四刀那台) **0/12** —— BASE 反而不炸，是因为它从不放旧值，同一只块每轮只被`Clear(&last)` 放过一次，代价就是形④ 那条漏。修法仍在这枚出口里: 站④ 先问「右值本身是不是一枚活着的槽」(裸标识符且在 `knownVariantVars_` 里) ⇒ 是就走 `vb6_VariantAssign`(清 + 深拷贝)；函数返回/集合取回那些内联临时量照旧清+搬。判据补一针 `B164-alias=40/40`(夹具里 300×40 次 `keep = e` 往返) + 发码形状针一行，读数存档 `.build/k55_nc_alias.txt`(7/12 → 0/12)。
  - **新登记三条(本刀刻意不动，各有前置问)**：**形⑤** 名单里那三站仍在裸发 —— `cgen_base_generate_evt_impl.inc:103`(事件实参 `pName = vb6_VariantFromStackVARIANT(&args[i])`)、`cgen_control.cpp:83`(`For` 循环起点 `varAcc = vb6_VariantFromValue(start)`)、`cgen_setlet.cpp:246`(Property Let 进槽)，每一站都得先问「右值是不是借来的」才能定 flag，不是改个写法；**形⑥** 站⑥ 交给出口的 flag 是按「发码不许变」定的，`Fix 232` 那一问(UDT 的 Variant 字段 + 借来的值 → 深拷贝)还没在这一拍板；**形⑦** ParamArray 那一站的中间槽 `_pa_vN` 每轮被裸覆盖、从不清 —— 当时记的「`vb6_PA_GetVariant` 交回来的是浅拷贝」**第五十六刀实测是错的**（四窗相减量到每元素漏 87/61 字节，带权副本才配这个数；RTL 那句「转交所有权、由调用方 Clear」才对），已连同修法一起收进 **§B166**。另有一条纸面观察(没量到后果)：`held = e` 这种 Variant→Variant 赋值是浅拷贝，槽与源共用那只 BSTR，谁先 Clear 谁就悬垂 —— 要处理得先定「槽之间该不该各持一份」。
  - **护栏**：全仓 **56 枚哨兵**除清单那条 K1(新夹具待登记)外全绿；`check_rtl_embedded` 125/125(RTL 本刀没动，所以不必 touch `c3rtl.rc`)；BASE/NEW 的 `sum` / `last-len` / `arr-sum` / `dec` / `held` 五行证人读数逐行对齐(证人不许动)。**门号链**：**门 #577**(head `a2f47ca3`、attempt 1)= 12 job 里 10 绿 + 两条预期红(`Emit manifest` 报 4 移动 + 1 新增；`Tests (compile)` 那条 `[STATIC] manifest_coverage` 同源)。四行移动**逐行看过**：`test_arr_empty.bas`(数组站 → 清+深拷)、`test_host_decimal.bas`(集合两站补清，-2/+4)、`SbHm.vbp`(同一站、循环变量 `pv`)、`Charts 2020/Proyecto1.vbp`(只多一行 `vb6_VariantClear(&Value)`，那一处是站①) ⇒ 都是本刀要的形状；站⑥ 在语料里没命中，`viaTemp` 那一路逐字节没动 ⇒ 没有别的行因此移动 → 登记 `e75692dc`(吃 CI 自己那份 head=a2f47ca3 的 404 行清单；本机复算 404/404、K1..K3 sets equal、注释 29 行没被吃掉) → **验证轮 #578**(head `e75692dc`)= 12 job 里 11 绿 + **一条真红**:`Tests (bas #1)` 里 `[RUN] test_host_decimal_x86 ... output mismatch` —— 那是第五十四刀自己的夹具把本刀的另一半句照出来了(见上面「站④ 那一格」)，本机复现 = 同一台 x86 跑 12 次 7 次 `0xC0000374`、BASE 那台 0/12。→ 第二半刀 `b708ee8e` → **门 #579**(head `b708ee8e`)= 11 绿 + 唯一红 = 形状门报 11 行移动(逐行都是「两行并成 `vb6_VariantAssign` 一行」这一形)→ 登记 `f3859efb`(吃 CI 那份 head=b708ee8e 的 404 行清单，本机复算与它 404/404 逐行相同)→ **验证轮 #580**(run 38326102269、head `f3859efb`、attempt 1)= 12 job 逐格复核全 completed/success(其中 `Emit manifest (shape oracle)` 与 `Tests (compile)` 两格同绿 = CI 那台独立复算证实了 `f3859efb` 那份登记) ⇒ 第五十五刀(两半句同批)就此收线
- **形⑪（借来的 BSTR 进槽）另立 §B167，第五十七刀已出**：本格问的是「没落到任何变量上的带权值谁回收」，那一形问的是「落到槽里的值是谁的缓冲」—— 同一枚出口的**第二问**，所以两格共用 `emitVariantSlotStore` 与它的 `rhsOwnsValue`，但判据与站点各算各的。
- **顺带量到、刻意不归本格**：`CDec(5)` 装进 Collection 之后读回 **0/Empty**、x64 的 `VarPtr` 那枚 I8 同样读回 **0/Empty**（两台逐行相同，`.build/k53_p/x64b.out` 与 `x86b.out`）；直接装箱那一路是对的（`VarType=14/Decimal`、`20/Long`）。那是「往外读」那张表缺 Decimal/I8 两档 = **§B142 的最后一族**，第五十三刀打的就是它。

### B165 宿主存值腿(`vb6_Host_FromWinVariant`)改问借用权威就是悬空指针 —— 这一问要连着四站一起改（2026-10-13 第五十三刀撞出来、**第五十四刀已出** = 值档的「谁分配谁释放」收成一对出口 `vb6_VariantValueFree`/`vb6_VariantValueCopy` + 宿主环五站同批 + 哨兵 C2 + 夹具 `test_host_decimal` 两台 12/12；**门 #575（head `06873ec1`）= 12 job 里 10 绿 + 两条预期红（形状门只报那一行新夹具、compile 那道静态 `manifest_coverage` 同源）→ 登记 `fe49c260`（吃 CI 自己那份 head=06873ec1 的清单，本机复算 403/403）→ 验证轮 **#576**（run 38307415571、head `fe49c260`、attempt 1、逐 job 复核 12/12 completed/success，形状门与 compile 那道静态从此绿）**）

- **怎么撞出来的**：第五十三刀把往外读那一腿(`vb6_Host_ToWinVariant`)接上唯一权威之后，顺手想把反向那枚(`vb6_Host_FromWinVariant`，九档手抄表，Date/Currency/I8/Decimal 落 default ⇒ Empty)接上 §B144/§B162 的 `vb6_VariantBorrowIntoSlot`。改完 x64 读数对了(`c.Add CDec(5)` 之后 `VarType(c(1))=14`、`CStr=5`)，**x86 读回 `14|Decimal|2066035336255469780997`** —— 标签对、数值是垃圾。
- **为什么**：x86 那一台的 DECIMAL 槽里装的是一枚**指针**，而 `BorrowIntoSlot` 的语义是借用(§B144 注释写死的：存活期 = 这一趟 Invoke)。`vb6_Collection_Add` / `AddKeyed` / `Host_SetProp` 都是「转一手再**存下来**」的消费者：`FromWinVariant → vb6_coll_push → Host_ClearVariant`，存进集合的那一格指向调用方那枚临时 VARIANT，一返回就是悬空。改前它落 default 交 Empty，**错得保守**，所以这条洞一直没现形。
- **要一起改的四站**：① `vb6_Host_FromWinVariant`(取回时给一份私有副本) ② `vb6_Host_ClearVariant`(回收那枚副本，现在只放 BSTR) ③ `vb6_coll_push`(存进集合时深拷，现在只有 BSTR 那一档拷) ④ `vb6_coll_itemCopy`(取回时深拷，同上)。四站同批才收支对称 —— 与 §B160/§B161 那条「所有权要跟值类型走」、§B185「一层一个名字」是同一个纪律的另一面。
- **判据(已经量好的)**：`.build/k53_p/probe.bas` 两台并排 —— `c.Add CDec(5)` 之后 `VarType/TypeName/CStr` 三读数与 `For Each` 那一行，x64 与 x86 **都要**答 `14/Decimal/5`；x86 直接装箱那一路本来就对(`box-dec=14|Decimal|5`)，所以红的一定是存值腿。指针那一腿(x64 交 20、x86 交 3)按架构分针，别塞进两台同针的名单。
- **顺带一条没量的**：`Date` / `Currency` 直接装箱就是 5/Double(两台)，`VarType(c.Add CCur(12.34))` 读 5 —— 那是**装箱档**的形状不是这一族，与 §B142 无关；要不要按 VB6 的 6/7 去钉，另开一格先量口径。
- **落地（一处权威 × 五站，同批）**：新增一对**值档**出口 —— `vb6_VariantValueFree`（BSTR + x86 的 DECIMAL `CoTaskMemFree`）与 `vb6_VariantValueCopy`（BSTR 另分配一只 + x86 的 DECIMAL `CoTaskMemAlloc` 后抄 16 字节；分配失败就把 dst 改成 Empty，不留悬指）。这一问原来住两遍（Clear 里的 rev11 三行 + Assign 里的 rev16 一段），现在合成一遍，`vb6_VariantClear`/`Copy`/`Assign` 三枚一律转问这一对。**对象那一档刻意不进这对出口** —— 加减引用仍只住 Copy/Clear 两枚、判据仍是 `vb6_ComIsDispatchable`，于是宿主存的借来指针照旧不计数，与改前一字不差。五站 = `vb6_Host_FromWinVariant`（九档手抄表整张撤掉，改问 `vb6_VariantBorrowIntoSlot` 拿布局、再问 `ValueCopy` 拿所有权；`VT_ARRAY|VT_BYREF` 与 `VT_UNKNOWN` 照旧交 Empty）/ `vb6_Host_ClearVariant` / `vb6_coll_push` / `vb6_coll_itemCopy` / `vb6_Collection_Remove`。
- **判据（两台同针）**：新夹具 `tests/test_host_decimal.bas` 12 针，针面从 `$b165Needles` 解析不手数，注册两条（x64 与 `-Arch "x86"` 同针）—— 正是「两台都要答」那条口径。新台两台均 **12/12 零 miss**（`.build/k54_fix_x64.txt` / `k54_fix_x86.txt`）：`coll=14|Decimal|5`、`two=5/7`、`fe-1=14|Decimal|5`、`fe-2=14|Decimal|7`、`rem=3/7/42`。基线 = 第五十二刀收线那台 exe（不含五十三/五十四），两台**逐行相同**、恰好五针红：`coll/fe-1/fe-2` 全答 `0|Empty|`、`two=/`、`rem=3//42`（`.build/k54_ab_base52_*.txt`，新台同针在 `k54_ab_new54_*.txt`）。证人不允动：`box=14|Decimal|5`（直接装箱本来就对）、`str=8|String|abc`、`long=3|42`；`arr=0|Empty` 钉住「数组 / ByRef / Unknown 三档本刀刻意不收」。屏自带正控 = 故意加一条假针 `B165-coll=0|Empty` ⇒ 两台各报一次 MISS（`.build/k54_neg_*.txt`，本轮重跑过一遍）。
- **归因订正（这条比刀本身值钱）**：「只 Copy 不 Assign 的两处调点（`vb6rtl_array.c` 的 Variant 数组整体赋值、`vb6rtl_date.c` 的 `vb6_ArraySetVariant`）也坏」是**按形状推出来的，A/B 量不出来** —— 那两形的针 `B165-elem=5/7` 与 `B165-whole=5/7/5` 在 BASE 与 NEW × 两台**逐字节相同**（独立探针 `.build/k54_p/arr.bas` 四份亦同，`.build/k54_arr_base52_*` vs `k54_arr_new54_*`）。⇒ 它们是**纸面隐患**，本刀按同一对出口顺带收掉，但台账与源码注释里都不许写成「实测修好了两处」；量到真坏的只有宿主环那五站。同故把哨兵文案里那句「实测调点两处」改成了「按形状推出、A/B 量不出」。
- **负控两条，各红各的**：A（结构性，静态哨兵，`.build/k54_nc_a2.py`，跑完按 md5 还原）① 把 `vb6_coll_push` 退回手抄浅拷贝 ⇒ `FAIL C2 vb6_coll_push no longer asks vb6_VariantValueCopy`；② 在同一站里插回一张 `switch` ⇒ `FAIL C2 ... owns a switch again`。两条各自**只红一行**，各对应 C2 的两个面。B（行为性）= 摘掉 `ValueCopy` 里那枚 DECIMAL 分配重编两台 ⇒ **x86 当场 `0xC0000374`（堆损坏）、只打印到第一针；x64 十二针全过**（`.build/k54_ncb_{x86,x64}.txt`）。这条**不对称本身**就是「x86 的槽里装的是指针、x64 是内联」的正面证据。
- **护栏**：`check_variant_i8_coverage.ps1` 读数 `switches 18 I8 must 13 exempt 5 | DEC must 10 exempt 8 | inbound core 1 | value pair 2 ring 5 | violations 0`；站点 19→18 = 撤掉宿主存值腿那张表，它自己的两条豁免随表一起撤，V1 下限同步 19→18（归因写在哨兵注释里）。邻居夹具原样绿：`test_variant_boxing` 两台 82/82 针全中、`test_com_latebound` 两台 15/15 针全中。发码清单 402 行**逐行不动**、只多新夹具那一行（`.build/k54_manifest_new.txt` 403 行 vs 登记 402：相同 402 / 哈希不同 0 / 缺席 0 / 多余 1）⇒ 形状门预期只红这一行，登记照纪律只吃 CI 自己那份读数。RTL 动了就 touch `c3rtl.rc` 重编（本轮 `check_rtl_embedded` 全对）。
- **下一格先手**：§B142 到此只剩「两枚 0 调用者的死码 `vb6_EnumNext`/`vb6_EnumRelease` 撤不撤」一条等拍；§B164 剩下的 ①②③⑤ 四形缺的是口径不是手法（① 语句级带档名单，先例 = M22 的 `ansiTempsToFree_`；② 运行期 arena），站点账与语料暴露面已量清，取决于用户拍哪条 —— **形④ 那条已经照「增量上限」的口径钉死了**(改前 81/71 B/元素、改后 < 8 B，两台同针)，它留下的三条新开放项(形⑤/⑥/⑦)也各有一条要先问的话，见上面那两条。


### B166 一枚 ParamArray 往返里现造的三样东西没人回收 —— 数组改归被调方统一出口销毁、中间槽改问 §B164 那枚出口（2026-10-13 第五十六刀量完并**已出** = 提交 `cd1ba9b6`；门链 **#581**（run 38333461651、head `cd1ba9b6`、attempt 1）= 12 job 里 10 绿 + 两条**预期**红（形状门「6 份发码真的变了」+「新增 1 份输入没登记哈希」；`Tests (compile)` 那条 `[STATIC] manifest_coverage ... FAIL` 同源 —— 都靠 annotations 定性，十个真跑 job 全绿 ⇒ 新夹具在 CI 那台上也过）→ 登记 `ec5e4a5a`（吃 CI 自己那份 head=cd1ba9b6 的 405 行清单；本机复算与它 405/405 逐行相同、29 行注脚未动）→ **验证轮 #582**（head `ec5e4a5a`、attempt 1）= 12 job 逐格复核全 completed/success，`Emit manifest` 与 `Tests (compile)` 两格同绿）

- **题面（一形三问，是探针自己撞出来的，不是读码猜的）**：`For Each v In <ParamArray>` 这一次调用里编译器现造了三样带权的东西 —— 调用点的 `SAFEARRAY* _pa_N = vb6_PA_Create(k)`、每个字符串实参的 `vb6_BSTR_FromStr(L"...")`、被调方每轮读元素的中间槽 `_pa_vN`。三样都没人回收，而 RTL 里那枚 `vb6_PA_Destroy`（`vb6rtl_paramarray.c:55`）**全仓 0 个调用者** —— 出口早就写好了，发码那侧从没见过它。
- **量（两台 × 两架构，判据问增量不问绝对值 = `K32GetProcessMemoryInfo` 的 PagefileUsage）**：
  - 数组那一半：`long-window` 那一窗刻意用 **Long 档实参**（槽里不带权 ⇒ 它非 0 只可能是数组本身漏了）：BASE **158/109 字节/次调用**，本刀之后**精确 0**。
  - 中间槽那一半：四窗相减 `((A8-P8)-(A1-P1))` 把打包代价与每次调用的尾巴一起抵掉，只剩 `_pa_vN` 自己 ⇒ BASE **87/61 字节/元素**，本刀之后 49,152 / 188,416 总字节 = **页粒度噪声**（0.7 / 2.7 字节/元素）。
  - **一条口径订正（比刀本身值钱）**：§B164 形⑦ 原记「`vb6_PA_GetVariant` 交回来的是浅拷贝」是**错的** —— 若真是借用，裸覆盖不配每元素漏 87/61 字节。RTL 那句注释「转交所有权：栈上 v 已无栈外引用，交出其内部缓冲由调用方 Clear」才是对的（`SafeArrayGetElement` 对 VT_VARIANT 数组交的是带权副本）。已在本格把那句改掉。
- **落地（两半各一枚出口，不新造机制）**：① 新增 `emitParamArrayReclaim(node.params)`（`cgen_util_ctrl.cpp` 定义 + `detail/util/cgen_helpers.inc` 声明），Sub/Function/Property 三个过程出口各调一次，位置就在 M22 的 ANSI 临时量之后 —— 口径 = 这枚数组的寿命就是这一次调用。安全性实测过：每个 VB 过程的 PA 实参都由**调用点**的 `vb6_PA_Create` 新造，连 `Inner b`（把一枚 ParamArray 再递出去）那形也是重新装箱（`.build/k56_p/fwd.bas` 的产物），所以销毁碰不到别人的指针。② For Each 的 `_pa_vN` 那一站改问第五十五刀那枚 `emitVariantSlotStore(…, true)`，并在块尾补一声 `vb6_VariantClear(&_pa_vN)`。
- **判据**：新夹具 `tests/test_pa_reclaim.bas` 两台同针 8 枚（精确针 = `long-window=0`；`loop-ok` 是**带宽**针，因为页粒度会抖：同一台两次读 49,152 与 188,416，而 BASE 同针 6,135,808 / 4,276,224 必红）+ 发码形状针 4 行（`Test-EmitcShape "b166_pa_reclaim_shape"`）+ 三枚读数证人（`sum8=40` / `count=3` / `last=XXXX` —— 后两枚钉的是「清了前面那份，循环变量手上那份还是对的」，即借用与所有权没被搞反）。
- **哨兵**：第 57 道 `scripts/check_paramarray_reclaim.ps1`：P1 出口 1 定义 / 1 声明 / 调用点 ≥3，P1b 出口本体恰好 1 枚 `emitLine` 且发的就是那句销毁（空壳出口不算通过），P2 `vb6_PA_Destroy` 的发射只许住出口那份文件，P3 `vb6_PA_Create` 的发射点唯一（出生地唯一，配对才谈得上），P4 那一站必须走 §B164 出口且 `_pa_v = …` 裸形为 0，P5 块尾收尾存在。五条负控各红各的一条，全部在 `src/backend` 的**副本**上做（真实树一个字节未动）：站点手抄→P1+P2、出口本体摘掉→P1b、中间槽改回裸覆盖→P4、收尾摘掉→P5、数组出生地开第二处→P3。
- **清单逐行归因（本机复算，`.build/k56_attrib.py` → `.build/k56_attrib.txt`）**：移动 6 行 + 新增 1 行；16 处差异**全部**归到本刀那两形 = 14 × `vb6_PA_Destroy(…)` + 2 × `vb6_VariantClear(&_pa_v…)`，**OTHER=0**，且 `-0` = 没有任何一行是被搬走的。
- **本刀刻意没动，新登记三条（各有前置问）**：
  - **形⑧ 字符串实参被现造两次**：打包站点发 `vb6_PA_SetBSTR(_pa_N, i, vb6_BSTR_FromStr(L"…"))`，而 `vb6_PA_SetBSTR` 内部**又** `SysAllocString` 一份（数组最终拿到第三份）⇒ 交出去那两份没人 free。四窗抵干净之后的残差实测 **123/100 字节/字符串元素**（Long 档窗口 = 0 把它单独钉住，探针 `.build/k56_p/pa_probe3.bas`）。修法要先定一句口径：`vb6_PA_SetBSTR` 的参数是「拷贝的原料 `const wchar_t*`」还是「一份所有权」—— 今天的实现是前者，所以站点该直接给 `L"…"`、别先 mint。
  - **形⑨ `ParamArray x()` 不带 `As Variant` 时 For Each 走错臂 ⇒ cl C2440**（本刀的探针第一次跑就撞上）：`cgen_control.cpp` 里「这是不是一枚 ParamArray」有**两位主人** —— 选臂的 `isCollArray` 问符号表的 `isArray`（带 `As Variant` 才登记），同函数下面的 `isPA092q` 问形参的 `isParamArray`（parser 两种拼法都置位，`parser_decl_var.cpp:336`）。裸拼法掉进 COM 集合臂 ⇒ `vb6_VariantToObjectVal(SAFEARRAY*)` C2440。语料 95 处 PA 签名里 **16 处裸拼法**（带 As 的 79 处全是 Variant），与 For Each 的组合今天 0 处 ⇒ 编不过但暂时不响。
  - **形⑩ `Exit For` 跳过块尾那一声收尾**：产物里 `vb6_loop_exit_N:` 落在那对花括号**之外**（实测 `.build/k56_p/exitfor.bas`），走这条路时 `_pa_vN` 手上最后一份带权值没人收（每次调用一份，有界）。要覆盖它得给循环栈带一枚「块尾回收清单」，与 §B164 的 ①/② 那条口径同族 —— 等拍。

### B167 一枚**借来的 BSTR** 收进 Variant 槽时，槽与源头共用同一只缓冲 —— 「这枚右值自带所有权吗」这一问有几位主人（2026-10-13 第五十七刀量完并**已出** = 提交 `9d520d4a`；门链 **#586**（head `9d520d4a`、attempt 1）= 12 job 里 10 绿 + 两条预期红（形状门 6 移动 + 1 新增、`Tests (compile)` 那条 `[STATIC] manifest_coverage` 同源，十个真跑 job 全绿 ⇒ 新夹具在 CI 那台上也过）→ 登记 `4076cf99` → 验证轮 **#587**（run 38350859716、head `4076cf99`、attempt 1）= 12 job 逐格复核全 completed/success，`Emit manifest` 与 `Tests (compile)` 两格同绿 = CI 那台独立复算证实了这份登记）；**形⑭（同一枚权威的第二腿）第五十八刀已出** = 提交 `c3efcc30`，门链 **#589**（run 38360606587、head `c3efcc30`、attempt 1）= 12 job 里 11 绿 + 一条预期红（形状门「存量 7 份发码真的变了」；本刀没加新输入 ⇒ `manifest_coverage` 那格**直接绿**，与登记口径无关）→ 登记 `d9248927`（7 行 = 本刀 1 行 + 别人那两笔欠的 6 行，逐行点名）→ 验证轮 **#590**（run 38362819214、head `d9248927`、attempt 1）= 12 job 逐格复核全 completed/success，`Emit manifest (shape oracle)` 同绿 = CI 那台独立复算证实了这份登记

- **题面**：装箱出口 `vb6_VariantString(BSTR)`（`src/rtl/core/vb6rtl/vb6rtl_variant.h:159`）只做 `v.vt = vb6_vtBSTR; v.bstrVal = val;` —— 既不 `SysAllocString` 也不加引用；`_Generic` 的 `wchar_t*` 那一档（同文件 241/262 两张表）指的也是它。所以「把一枚 VB String 装进 Variant 槽」只要右值是**一处活存储**，槽拿到的就是别人那只缓冲：源头一次重赋值 ⇒ 槽里读的是已释放的块；槽一次 Clear ⇒ 把源头还活着的那只块放掉。
- **这一问的主人有几位**（发码实测，探针 `.build/k57_p/stat.bas` / `stat2.bas` / `own.bas`，产物同一台 `--emit-c`）：**三站两形只有一位答了**。UDT 的 Variant 字段（`r.Fv = s`）与 Variant 数组元素（`arrV(0) = s`）发 `vb6_VariantAssign`（Fix 232 / Rev16 = 清 + 深拷）⇒ 对；**普通 Variant 槽**那一站（局部、形参、模块级、以及 `vb6_ret_<返回 Variant 的函数>`）发的是 `vb6_VariantClear(&v); v = vb6_VariantString(s);` ⇒ 根本没问。两条文本路到同一处：`wrapVariantValue` 的 String 档发甲形（`vb6_VariantString(x)`），`boxToVariant` 落 `_Generic` 发乙形（`vb6_VariantFromValue(VB6_SA_AT(BSTR, …))`）—— **同一站、两形、同一个偷**（与 [[two-variant-predicates]] 那条「一条问有两条答案」同族，判据要两头钉）。
- **后果实测**（两台 BASE，读数 `.build/k57_base_x64.txt` / `k57_base_x86.txt`，两台**逐行相同**）：`s = "sharedbuf": v = s: v = 1` 之后再读 `s` ⇒ `B167-W2=|9`（源串读成了它自己那一行的前缀 = 别人用过的堆块，Len 还是旧的 9）；200 轮的 `v = s: w = s: v = 1: w = 2`（同一块被放过两次）之后读 `s` ⇒ `||8`。**另三窗（局部/模块级/形参）在 BASE 上读着"对"** —— 那只块还没被人复用而已，是悬空读不是好读数。
- **别把"没崩"读成"没有"**：`.build/k57_p/dbl.bas` 那枚纯二次 free 的形状两台都 RC=0（本机 CRT 不立刻校验堆链），只把源串读成 `2|1`。⇒ 判据不能靠崩溃，要靠**值面**与**发码形状**。
- **落地**（一问一处）：新增 `CCodeGen::variantRhsOwnsValue(Expr* rhs, const std::string& rawC)`（声明 `detail/util/cgen_helpers.inc:88`，定义 `cgen_util_ctrl.cpp:2396`）—— 第五十五刀住在站点里的那截内联判据（`rhsIsLiveSlot164`：右值是不是活槽）**并进来**，新答的一形是 `exprBorrowsBstrStorage`（同文件 2382）：串数组元素认宏前缀（`VB6_SA_AT(BSTR,` / `VB6_SA_ND_AT2(BSTR,`，与左值侧 Fix 177/rev11 同一口径）、字面量与**顶层是调用**（内在函数、`vb6_BSTR_Concat`、工程函数与 Property Get 的返回、`Public Const` 折出来的 `vb6_BSTR_FromStr`）算现造自有，剩下的存储（裸名 / UDT 与类的串字段）问 `inferExprType` 答 String 才算借用。§B164 站④ 改成问这一枚，答案 false 时由 `emitVariantSlotStore` 走 `vb6_VariantAssign`（值档的深拷住第五十四刀那对出口）⇒ **没有新机制**。
- **判据**：新夹具 `tests/test_bstr_borrow_slot.bas` 两台各 **15 枚针**（`$b167Needles`）：9 枚借用来路（局部 / 形参 / 模块级 / UDT 串字段 / 两种串数组元素=甲乙两路 / Variant 返回槽 / 二次 free 的 200 轮 churn）+ 4 枚证人（UDT Variant 成员与 Variant 数组元素 = **本刀不动**的两位旧主人，读数一字不许变）+ 三枚现造对照（字面量 / `Left()` / 槽到槽）。BASE 上红 W2 与 W7 两枚、NEW 两台全中。发码形状针 7 行（`b167_bstr_borrow_shape`：甲乙两路都归 `vb6_VariantAssign`）+ **反面** `Test-EmitcAbsent` 2 行（现造的串不许被多拷一次 —— 那正是 §B164 形① 欠的那笔漏）。
- **哨兵** `scripts/check_bstr_borrow_slot.ps1`（第 58 道）：Q1 那一问的定义/声明/调用点（定义成死码也算红）/ Q2 借用那一形各一处（单独露出来是给后来的站点当口径）/ Q3 站④ 必须问它、且旧内联判据不许回来（跨行调用先把空白压平再认）/ Q4 借用那一问的四条腿都在（两种宏前缀 + 「顶层是调用」+ 类型表）/ Q5 **硬填 flag 的站点计数钉死 6** —— 新站点不问这一问就写死 `true` 当场红。五条负控各点各的红（NC2 的还原同时红 Q1+Q3 两条，是同一处回退的自证），原地插桩、跑完按 md5 还原并复核。
- **清单（门 #586 的读数 + 独立复算）**：形状门报 **6 移动 + 1 新增**。**关键读数**：别人那一棵 rebase 进来的树 `9ba20883`（**不含本刀**）在门 #584 / #585 上就已经报「存量 6 份发码真的变了」—— 同样这 6 行 ⇒ 这 6 行整行是 `wip/comctls-variant-fixes` 那条分支欠的账（随本刀的 rebase 第一次进 CI），两形都点得名：13 处 `(!_has_X)` → `(-(!_has_X))`（IsMissing 交 0/1，发码点 `cgen_expr_call_builtin_pre.inc`）+ 1 处 1D 描述符首字段 `signature 0x5A1D` → OLE 兼容的 `cDims 1, 0xA1D1`（发码点 `cgen_expr_call_arg_emit.inc`）。本刀自己的形状落在 `tests/dbgdlg/dbgdlg.vbp` **那一行里另加 3 处**（`.build/k57_attrib.txt` 逐行：`vb6_ret_sLeftOfNull` / `vb6_ret_sLeftOfLastNull` / `vb6_ret_sAPIFilter` = **Variant 返回函数收下自己的 String 形参**这一形，BASE 发 `Clear + 直赋`、NEW 发 `VariantAssign`），新增那一行是本刀的夹具。⇒ 登记 `4076cf99`（吃 CI 自己那份 head=9d520d4a 的 406 行清单；**本机独立复算 406/406 逐字节相同**、29 行注脚未动、K1..K3 sets equal）。**登记那一笔的题面把门号写成了 #583 —— 错了**：#583 是别人 head=`c5bc5ac4` 的那轮，本刀的门是 **#586**（run 38348063962、head `9d520d4a`），订正写在本格里。真语料暴露面（`D:/vb_yqt4qPac` 10138 份 `.bas/.frm/.cls/.ctl/.pag`，拿**改前那台** `--emit-c` 数「紧挨着的 Clear + 直赋 `vb6_VariantString(存储)`」两行式，甲乙两条文本路都数，按存储形态分档；正控 = 四份手核探针各自命中已知数 6/3/5/12）：**476 处、命中 252 份**（`.build/k57_exposure_base.txt`；分档 = 甲路 227（裸名 146 / 成员 42 / 元素 39）+ 乙路 249（成员 138 / 裸名 76 / 元素 35）—— 两条文本路同一个偷 ⇒ 本刀接管掉的正是这 476 处。口径要说全：10138 份输入里 **1031 份 `--emit-c` 就失败**（语料本身编不过 / 缺控件）、**48 份超时**（含 `GDI+透明控件/QQGrid.ctl` 那枚挂 45 分钟没出声的，另格另账）⇒ 真覆盖 9059 份，这个数是**下界**）
- **本刀另撞见 / 登记，刻意不动**：**形⑫** `Fn(...).V = <String>`（返回 UDT 的 Variant 成员在**函数体外**被赋值）发的是一句不装箱的裸赋值 `vb6_Make().V = s;` ⇒ cl **C2440 编不过**（`.build/k57_p/run_fnu/c3-error.log`，探针 `fnu.bas`）—— 响的，不是静默的，另格另量；【**2026-10-13 第五十九刀订正：那一格 = §B168，已合上** —— 目标侧一旦问工程级那张「类 → vb6_VARIANT 字段」表，`Fn(...)` 这形本来就被 `inferClassTypeOfExpr` 的 `IndexOrCallExpr` 那一支答得出类名，实测同一形从 `vb6_F12Main_MakeH()->PV = s;` 变成 `vb6_VariantAssign(&vb6_F12Main_MakeH()->PV, vb6_VariantString(s));`（探针 `.build/k59_p/f12`），**没有为它另开一刀**】；**形⑬ 那条推测已作废**（打本刀这一格时顺手量的：`With r` 里 `v = .Name` 与 `.FV = .Name` **两形都发** `vb6_VariantAssign`，
    探针 `.build/k58_p/w13.bas` / `w13b.bas` 的产物）—— With 成员作为**右值**时 `inferExprType` 答得出 String，本刀的借用那一腿收得住；**三枚发借来指针的 RTL 出口**（`vb6_ErrDescription` / `vb6_ErrSource` / `vb6_ho_variantToBstr`，#113 那条注记在册）在本刀的「顶层是调用 = 现造」这一腿里被读成自有 ⇒ 仍悬空，修法是给这一腿一份从 RTL 派生的名单（§B146 那一族的手法）；**形⑥ 站⑥** 与 **形⑤ 那三站**的前置问（「右值是不是借来的」）现在这枚权威已经能答，接线时按各自目标是不是长期槽逐站定。
- **形⑭（第五十八刀）：那一腿的例外名单**：`b167TopIsCall` 答「现造的」对 RTL 的大多数出口成立，但有 **3 枚**交回的是
  **别人那一只还活着的块** —— `vb6_ErrDescription`(`return vb6_err.description;`)、`vb6_ErrSource`、
  `vb6_ho_variantToBstr`(`return v->bstrVal;`)。名单只住一处（`kBorrowedBstrExports167`，`cgen_util_ctrl.cpp`），
  命中即答「借来」⇒ 站④ 交 `vb6_VariantAssign`。
  **名单怎么来的**：不是按文本猜 —— 返回 BSTR 的出口全仓 **90 枚**，先按数据流派生（`.build/k58_audit.py`：68 枚自有 /
  2 枚直接派生成借来 / 7 枚进手核队列，其余落复合档所以这三数不加成 90 —— 这条不完备正是它必须配一条独立复算的原因），
  手核队列逐枚读体定档 ⇒ 借来的最终就这 3 枚，`vb6_IIfBSTR` 是**刻意的转手参数**（它 free 未被选中的那一支），
  `vb6_DecimalToBSTR` 的注释本身就写着「把权抄给调用方」；而哨兵 Q6 用的**形状规则**独立扫这 90 枚，
  派生出的正好也是这 3 枚（`.build/k58_q6_proto.py` 的正控）⇒ 两条互不依赖的路同一个答案。
- **形⑭ 的读数**（两台 BASE = 第五十七刀那一台，两台**逐行相同**）：`Err.Raise 5, "errsrcw9", "errdescw8"` 之后
  `sA = Err.Description : sA = 1` 再读 `Err.Description` ⇒ **`B167-W8b=9|9`**（源串读成一位数、Len 还是旧的 9 = 那只块已被复用），
  `Err.Source` 那一窗 ⇒ **`B167-W9b=8|8`**；新两台 ⇒ `errdescw8|9` / `errsrcw9|8`，19 枚针同过、两台同针
  （`.build/k58_ab.txt` / 合并后复跑 `.build/k58_ab2.txt`）。判据仍是**两头钉**：`W8a/W9a` 问「槽里存的是请求值」，
  `W8b/W9b` 问「Err 自己那一面还在不在」。发码形状针 7→9 行（新增这两枚借来调用归 `vb6_VariantAssign`），
  反面证人 2 行原样（`Left()` / 字面量一次都不许多拷）。
- **形⑭ 的哨兵 = 第 58 道加 Q6**（双向对钉）：名单只许定义一处、恰好 3 名、不许重名；反向从 RTL 现推 —— 凡「体内无 mint 标记
  且某条 `return` 交回**成员访问**（非调用起手）」的 BSTR 出口都必须在名单里，名单里的名反过来也必须还呈那一形。
  五条负控各点各的红且**只红 Q6**（少一名 / 多一名自有 / RTL 那枚改成交自有 / RTL 新添一枚借来的 / 名单换名开第二家），
  全部在 `src/backend` + `src/rtl` 的**副本**上做（真实树一个字节未动，`.build/k58_nc.py`）；正控 = 派生规则在 90 枚里推出的正好是那 3 枚。
- **形⑭ 的真语料暴露面 = 0 处**（这条本身是「下界口径」的反面教材）：`D:/vb_yqt4qPac` 里 **1067 份**输入含 `Err.Description`/`Err.Source`，
  拿改前那台逐份 `--emit-c` 数「Clear + 直赋这两行式」（甲乙两条文本路都数，正控 = 夹具必须报 2）⇒ **处数 0**
  （其中 220 份 emit 就失败、11 份超时 ⇒ 真跑通 836 份，`.build/k58_exposure_b2.txt`）。⇒ 这一形在语料里今天**编得过、响不着**
  （真实写法是把 `Err.Description` 交给 String 目标 / `MsgBox` / `Debug.Print`，那一路由 #113 的深拷管着），
  本刀的价值在**那一问的答案本身**：站④ 的答案已经错过一次，形⑤⑥ 接线时不必再继承这个错。`vb6_ho_variantToBstr`
  那枚今天只能以调用实参的形态出现（`unboxArg` 那一处），进槽那一路还没现形 ⇒ 名单收住它、行为判据只覆盖 Err 两枚。



### B168 「对象.类的公共 Variant 字段」当赋值目标时**根本不装箱** —— 目标侧「这是不是一枚槽」有三张名单各答各的（2026-10-13 第五十九刀 **已出** = 实现 `f909ef7f` → 门 **#591**（run 38379525625、attempt 1、10 绿 + **两条同源红** = 形状门那 2 行新增未登记 + `Tests (compile)` 里那道 `[STATIC] manifest_coverage` 同一件事）→ 登记 `d29a387e` → 验证轮 **#592**（run 38381571448、head `d29a387e`、attempt 1、12 job 逐格 completed/success）；§B167 挂着的 **形⑫** 随本刀一并合上；同族还剩三形未开工，见末尾三条）

- **题面（真编译撞出来的，不是读码猜的）**：`Dim h As Holder` + 类里 `Public PV As Variant`，写 `h.PV = <右值>` 发的是
  `h->PV  /* class var .PV field */ = <右值>;` —— **一次装箱都没有**。除 Variant→Variant 那一形（结构体抄 = 合法但浅拷 =
  §B167 记的 `held = e` 同族）以外全是 **cl C2440 ⇒ 整个工程编不过**。真读数：
  `uselet.c(18): error C2440: “=”: 无法从“BSTR”转换为“vb6_VARIANT"`（探针 `.build/k59_p/uselet.bas` + `holder.cls`，
  四形 RHS 各发一句裸赋值：`= 5` / `= v2` / `= vb6_BSTR_FromStr(L"lit")` / `= vb6_Left(…)`，`.build/k59_p/use2.bas`）。
- **同族的第二形（同一处病）**：§B167 登记的 **形⑫** `Fn(...).V = <String>` 发 `vb6_Make().FV = s;` ⇒ 同样 C2440。
  两形的共同点 = **目标**是一条成员访问文本，而发码侧认槽的那几张名单都没覆盖它。
  【**第五十九刀实测：这一形随本刀一并合上** —— 接收者是 `IndexOrCallExpr` 时 `inferClassTypeOfExpr` 本来
  就答得出类名，所以那张新表也认它：同一形从 `vb6_F12Main_MakeH()->PV = s;` 变成
  `vb6_VariantAssign(&vb6_F12Main_MakeH()->PV, vb6_VariantString(s));`（探针 `.build/k59_p/f12`，两台对照）
  ⇒ **没有为它另开一刀** —— 这正是「收成一处」收得住的证据。】
- **为什么**：目标文本出自 `cgen_expr_member_class_fallback.inc:263`（类公共字段那一支，名字里带 `/* class var .X field */`），
  而 §B164/§B167 那枚站④ 认槽用的是 `knownVariantVars_`（局部/模块/形参的**变量名**）、
  `classVariantMembers_`（读侧 Fix 089j 用的那张类字段名单）、以及**文本** `VB6_SA_AT(vb6_VARIANT, …)`（数组元素）——
  **三张名单三种口径，没有一张认这条目标** ⇒ 与 [[two-variant-predicates]] 同族：一条决定有两种答案时，
  造出来的正是另一个极端。修法方向 = 把「这枚目标是不是长期槽」也收成一处权威（lvalue 侧一枚，与 `variantRhsOwnsValue` 配对），
  而不是在站④ 再补一个 `find("/* class var .")`。
- **暴露面（真语料，纯文本两趟，口径 = 写点不算读点）**：**94 处赋值 / 28 份文件 / 9 枚带公共 Variant 字段的类**；
  点名的病人 = `D:/vb_yqt4qPac/NewTab-main/control-source/cls/cPropertyData.cls`（`Public Value As Variant`）的
  `iProp.Value = …` 一族（同项目的 `cNewTabTheme.cls` / `pagNewTabThemes.pag` 等 28 份）。⇒ 这一格**不是纸面缺口**：
  那棵工程今天整棵编不过（读侧 `.build/k59_exposure3.txt`，脚本自带两趟扫描：先收「类→公共 Variant 字段」，再收「变量→类型」）。
- **开工前要先答的两问 —— 都已答（第五十九刀）**：① 装到哪一档 = **不另开一档**，接站④ 现成的
  `wrapVariantValue`（标量 → `vb6_VariantLong/Double/Bool/String…`，`_Generic` 那串一个字没改）；
  ② 旧值回收 = 已登记的 `variantRhsOwnsValue` 那一枚权威，lvalue 侧一旦被认成槽，答案自己就进去了。
- **落地（一处问 + 一张表，不新造机制）**：
  ① driver 新增一趟预扫描 `src/driver/detail/driver_codegen_variantfield_scan.inc`，建工程级
  「类(小写) → vb6_VARIANT 标量字段」表，与 `classVoidFieldMap`/`classTypedFieldMap` 同格、setter 注入。
  判据 = 「类模块发 struct 时这条字段写的是不是 `vb6_VARIANT`」：无 `As` 子句 = 默认 Variant（Fix 089j 那一路）、
  `resolveTypeName == Variant`；**数组字段刻意不进**（C 字段是 `vb6_SafeArray1D*`）、String/定长串（BSTR）、
  项目类/COM/未知对象名（void* 或类指针，正是 Fix 090bx 那批漏网名字）、UDT（`vb6_type_x`）都不算槽。
  ② 后端一处查询 `classFieldIsVariantSlot(className, member)`（住 `cgen_api.inc`，与 `projectClassNameOf` 同格），
  类名问既有出口 `inferClassTypeOfExpr`（`IdentifierExpr`/`MeExpr`/链式成员/`IndexOrCallExpr` 返回类都在它覆盖里）；
  **不**认发出来的文本尾注 `/* class var .X field */`。③ 赋值站（`cgen_assign_value_sem.inc`）在 Fix 232 那一支之后
  补一条 MemberAccessExpr 判据；**否决**：member 若是该类的 method/属性（`resolveClassMemberCall` 非空）仍归属性写那一路
  —— 与读侧 `cgen_expr_call_callee_member.inc` 的 Pattern C/D 同一条判据。
- **实测读数**（BASE = `C3_base_k59.exe`）：`h.PV = 7` → `h->PV = 7;`、`h.PV = "s"` → `h->PV = vb6_BSTR_FromStr(L"s");`、
  `Me.PV = x` → `me->PV = (*x);`；改后三形分别是 清+`vb6_VariantLong(7)` / `vb6_VariantAssign(&…, vb6_VariantString(s))` /
  `vb6_VariantAssign(&…, vb6_VariantFromValue((*x)))` —— 借来的值一律深拷贝，现造的仍清+搬，与 §B164/§B166/§B167 同一口径。
  **浅存那一半是 BASE 编得过的**：证人工程 `.build/k59_p/b6`（全部写走类内 `Me.PV = v`）两台 BASE 都答 `WIT-1=bbb`
  （字段跟着调用方那枚槽一起变 = 别名），NEW 两台答 `aaa`（`.build/k59_wit1.txt`）。
- **护栏**：新夹具 `tests/b168_cvf`（一份 .cls + 一份 .bas 的 .vbp）两台各 9 行读数 + 发码形状针 7 枚 + 反面证人 4 枚
  （BASE 两台 = 各 7 枚 C2440、**根本没有 exe**；NEW 两台 9/9 全对且逐行相同 —— `.build/k59_fix1.txt`）；
  哨兵 = `check_variant_slot_store.ps1` 加 **R5** 五子规则（查询出口 1 枚且住 `cgen_api.inc` / 调用点不许归零 /
  那张表只在一趟预扫描里建且注入 1 处 / `class var .` 那句尾注**逐文件钉死 11 处** = 认文本那条歧路不许复活 /
  老表 `classVariantMembers_` 的写者仍只住一份文件 x4），7 枚假针在副本上各自只红自己那条、还原后重新绿
  （`.build/k59h_nc.py`）；语料 A/B = 408 输入 × 两架构 = 816 捕获，**只有新夹具那 2 份变、OTHER=0、run-failures=0**
  （`.build/k59j_ab_report.txt`），CI 那台自己交的清单复核 = 406 行逐行相同 + 2 行新增、**0 行移动** ⇒ 这张表在存量工程上没改过一个字。
- **⚠ 同族还剩三形（本刀刻意没碰，各有实测读数）**：
  ① **With 块内** `.PV = <值>`（探针 `.build/k59_p/holder4.cls` + `use5.bas`，存档 `.build/k59_p/run/h4.txt:215`）
  发的是 `vb6_holder4_prop_let_PV(_vb6_with_0, s);` —— 这个名字全仓**既无声明也无定义**（类里只有一枚公共字段，
  没有 Property Let），走的是 `cgen_assign_prop_write.inc:729` 那条「With 类属性直调」的路，它把公共字段也当属性拼 ⇒
  cl 只给 **C4013**（隐式外部 int，`.build/k59_p/k59_h4out/c3-error.log`），链期必 LNK2019。**修法要先拍口径**：
  那一站该改问本刀这张表（字段 ⇒ 走站④）还是该补一枚 field_let 出口 —— 与 §B164 形⑤ 那两条 `/* Let */` 裸发是同一族。
     【**2026-10-13 第六十刀订正：这一形已随 §B169 收掉** —— 那一站没补 field_let 出口，而是先问工程级那张「类 → 标量字段 → 声明档」表：是数据字段就按档发字段赋值 (Variant/BSTR 交 §B164/§B167 那两枚槽位出口，其余标量落到既有的字段直发那一支)。实测 DllPrint.vbp 那 281 句凭空的 vb6_<Cls>_prop_let_<F>(...) 全清零，本刀夹子在 BASE 上的 4 枚 LNK2019 也一并没了。】

  ② **Variant 数组字段**的元素写 `h.PA(1) = 9` / With 内 `.PA(0) = s` 发成 `h->PA(1) = 9;`（把 `vb6_SafeArray1D*` 当可调用）
  ⇒ 实测 **cl C2064 ×2**（同一份 c3-error.log）。这张表刻意不认数组字段，所以那一形今天仍然响。
  ③ 跨对象**读**回来那一腿：`CStr(h.PV)` 发 `vb6_CStrDbl(h->PV)`（`inferExprType` 对「别类的 Variant 字段」答 Double ⇒
  选错 CStr 的重载族），串值实测读回 `0|1` 而不是 `sharedbuf|9`（NEW 第一版夹具的 W2..W6 全 0 —— 改判据为走类内
  `ReadBack()` 之后才对上，存档 `.build/k59_fix1.txt` 的第一轮）。同一枚表能答这一问，但那是**表达式类型 oracle**，
  动它会扫到全语料的选档 ⇒ 另立一格，别顺手接。


### B169 「`<对象>.<字段> = v` 里这枚字段是哪一档」有三份答案，With 块那一站因此**凭空拼一个函数名**（2026-10-13 第六十刀 **已出** = 实现 `ca3a1e67` → 门 **#593**（run 38399757587、attempt 1、12 条里 10 绿 + 两条**同源**红 = 形状门「2 行新增未登记 + 1 行移动」与 `Tests (compile)` 里那道 `[STATIC] manifest_coverage` 是同一件事）→ 登记 `ee414607`（吃 CI 那台自己交的 410 行清单：407 行逐行相同 / 1 行移动 / 2 行新增、缺席 0 多出 0）→ 验证轮 **#594**（run 38402409894、head `ee414607`、attempt 1、12 job 逐格 completed/success）；§B168 在册的 ① 随本刀合上）

- **两形同一病**（都是「这枚左值属于哪一档」没人统一答）：
  ① **With 块里写类的标量数据字段** —— 属性写那一站在查不到 Property Let/Set 之后按命名约定**造一个名字** `vb6_<Cls>_prop_let_<F>(...)`（它的 `isDataField090c` 只查 `classTypedFieldMap_`，那里只有**对象**字段，于是所有标量字段都被判成「不是字段」）。那函数全仓既无声明也无定义 ⇒ cl C4013 + 链期 LNK2019。真读数：语料 `DllPrint.vbp` 的产物里 **281 句**这种调用（emit census `.build/k60_reach.txt`）；本刀夹子在 BASE 两台各 **4 枚 LNK2019 + LNK1120、无 exe**（`.build/k60_run/base_x64/c3-error.log`）；文本口径（带接收者纪律）另有 47 个写点 / 6 份文件（`.build/k60_census2.txt`）。
  ② **直接路写类的 String 字段**（`h.PS = s` / 类内 `Me.PS = v`）—— BSTR 档那五道判据问的全是「本模块认不认识这个名字」，于是发一句裸指针赋值：与源头共用一只 BSTR（源头一重赋值字段就跟着变）、旧值没人回收、而类析构 `vb6_BSTR_Free(那枚字段)` 会把**借来**的那只块 free 第二遍。真读数：MSNMacro 工程 `CDLG.cls` 两句 `Me.SelectedPrinterName = vb6_ComGetStringProp(myP, L"DeviceName")` 与 `= vb6_BSTR_FromStr(L""")`（`.build/k60_reach.txt`）；文本口径 13 个写点 / 6 份文件。
- **落地（一处答案 × 三站接线，不新造机制）**：
  · §B168 那趟预扫描的表从「只答 Variant」扩成「答声明档」—— `driver_codegen_variantfield_scan.inc` 改名 `driver_codegen_fieldkind_scan.inc`，值 = `Vb6Type`，键仍是小写类名 + 字段名两种拼法；**只收标量数据字段**：数组 / 定长串（`FixedStringTypeRef`）/ 项目类·COM·未知对象名 / UDT 一律不入表；工程枚举（`As typeField`，struct 发 `int32_t`）按 **Long** 入表，判据 = 符号表指着 `SymbolKind::EnumType`（与 Fix 050/098 给 COM 枚举、XxxConstants 定的底层档同一个答案）。
  · 一处查询 `classFieldKind(className, memberName)`（`cgen_api.inc`；交回 `Unknown` = 表里没有 ⇒ 调用点**不许**读成「那就是字段」）+ 三个投影 `classFieldIsVariantSlot` / `classFieldIsBstrSlot` / `classFieldIsScalarData`，投影里只许转问它，不许再抄一遍 find。
  · 站③ 补 String 档判据（与 §B168 的 Variant 那一支同形、同一条 `resolveClassMemberCall` 否决）；站③/站④ 那两串决定**整体搬进** `emitBstrSlotStore` / `emitVariantSlotAssign`（`cgen_util_ctrl.cpp`，紧挨第五十五刀那枚 `emitVariantSlotStore`）—— 搬出来不是为了整洁：With 那一站需要的正是同一串决定（右值折算 / 自有还是借来 / 定长收口 / 自指先收临时量），在那儿再抄一遍就是第四份答案。With 那一站按档发：Variant/BSTR 交这两枚出口（COM 标记仍由同一条腿消费），其余标量档落到既有的字段直发那一支（与对象字段同一条路）。
- **判据**：新夹具 `tests/b169_cfw`（一份 .cls + 一份 .bas 的 .vbp）两台各 6 行读数 —— `B169-1=aaa|3` 是别名那一半的行为面（BASE 上 `h->PS = s;` 之后再改 `s` 会读到 `bbb`）；发码形状针 **8 枚**（直接路 借来 / 自有字面量 / 类内形参 + With 站 BSTR / Variant 清+搬 / Long / 枚举）+ 反针 **7 枚**（BASE 的两条裸赋值 + 四个凭空的 prop_let 名字 + With 站不许发裸赋值）：NEW 两台 8/8 且反针 0 中，BASE 那台 0/8 且 6 形病灶全在。哨兵 `check_variant_slot_store.ps1` 的 R5 改问同一张表 + 新 **R6/R7**（两枚槽位出口各 1 定义 / 1 声明 / 调用点 >= 2；手抄 `vb6_BSTR_Assign(&…)` 的发码点**逐文件钉死** 3 / 4 / 1 = 还没并进那枚出口的三站）—— **8 枚假针**在副本上各自只红自己那条、还原后重新绿（`.build/k60_nc.py`）；`check_bstr_borrow_slot.ps1` 的 Q3 那一针跟着搬到出口本体的文件里（第五十五刀那枚站④的形状搬了家，判据不搬就是假绿）。
- **护栏**：语料 emit A/B = 410 输入 × 两架构 = 820 捕获，**same=816 / changed=4 / unattributed=0 / run-failures=0**（`.build/k60_ab_report.txt`）—— 那 4 份 = 新夹具两架构 + 第五十九刀那枚 `b168` 夹子；58 道哨兵 57 绿（唯一红 = 新夹具两份输入没登记哈希，登记后即 #594 全绿）。
- **一条读数订正（第五十九刀写下的那条「负控」）**：`b168` 夹子里 `h.PS = "plainstring"` 今天不再是裸存 —— 那张表认得 String 档，发 `vb6_BSTR_AssignMove(&h->PS  /* class var .PS field */, vb6_BSTR_FromStr(L"plainstring"));`（CI 那份清单 +21 字节移动的那一行就是它，与本刀 A/B 的逐行归因一条不差）。针与反针的措辞都跟着改了；**Long 字段那条负控仍然在** ⇒ 这张表答的是「三档各归各位」，不是「把所有字段都拖进槽语义」。
- **⚠ 同族还剩三形（各有实测读数，本刀刻意不碰）**：
  ① **类内** `With Me` 写自己的数据字段 —— 那一站把接收者解析成 `COMObject`，发 `vb6_ComSetProp(me, L"PS", ...)`，而字段没有 dispid ⇒ 运行期一声 `property "PS" not found`，这一句**静默丢了**（本刀夹具的第五形实测，已从期望里撤掉；语料口径 = 全语料 78 处 `With Me` 写的**全是属性**，写数据字段 **0 处** ⇒ 纸面缺口，先证明可达再动）。
  ② **定长串**字段（`As String * N`）刻意不入那张表（收口宽度 = 账 #118 那一格）：语料里跨对象/类内写定长串字段 **0 处**（带接收者纪律实测），且 #118 的补空格/截断本来就欠着 ⇒ 同一格里一起量。
  ③ §B168 在册的两条未动：Variant **数组字段**的元素写（`h.PA(1) = 9` ⇒ cl C2064，语料 0 处）与跨对象**读**回来走 `vb6_CStrDbl` 数值档（同一枚表能答，但那是**表达式类型 oracle**，动它会扫到全语料的选档 ⇒ 另立一格）。

### B170 「这枚成员读是哪一档」按**裸成员名**查全局，字段名撞上内置函数名就读成内置的返回档（2026-10-13 第六十一刀 **已出** = 实现 `790ececb` → 门 **#597**（run 38417888070、attempt 1、12 job 里 6 绿 + 6 红：`Tests (compile)` 那道 `[STATIC] manifest_coverage` 是本刀新夹具未登记（流程性），另外 5 道（4 个 vbp 分片 + 形状门）是 `656b2086` 那一笔带进来的**继承红**，见本节末尾）→ 登记 `43fe7761`（吃 CI 那台 head=790ececb 自己交的 412 行清单：本刀只动 **4 行** = 移动 2 行 b168/b169 + 新增 2 行 b170；**刻意不登记**别人那 7 行，理由在末尾那段）→ 验证轮 **#598**（head `43fe7761`：本刀自己那道红应清，那 5 道继承红不归本刀））

- **归因订正（这条最重要）**：§B168 在册 ③ 写的是「别类的 Variant 字段读一律答 Double」—— 错。对照组 `Public PZ As Variant`（名字不撞内置表）两台都读对 `withform|8`。真机制 = `inferExprType` 的 MemberAccess 兜底 `symTab_.lookupModule(成员名)` **按裸名查全局**，`PV` 命中 `builtin_funcs.inc:317 addBuiltinFunc("PV", Double)`（VB6 金融函数）⇙ `CStr(h.PV)` 发 `vb6_CStrDbl(那枚槽)`（RTL `_Generic` 有 `vb6_VARIANT:` 臂，按 VarR8 解）⇙ 串值折成 `0|1`；同名的 `Public PV As String` 也答 Double（声明档整个被吞）。
- **最响的一形是崩，不是打印差异**：`Public Command As Long` 撞内置 `Command()`（答 String）⇙ 发**裸 BSTR 读**，把 7 当指针解 ⇒ BASE 两台 `RUN-RC=0xC0000005`（与账 #229/#231 那枚 `SSTab1.Tab` 撞内置 `Tab` 同一条病：那一次修的是控件档，这一次补的是类字段档）。
- **类内两条腿本来就对**：`Me.PV` 与裸名 `PV` 在类自己的过程里读回 `withform`（探针 K61-8/9）⇒ 只欠点号那一路 + With 体内那一路两站。
- **落地**：两站各补一问 `classFieldKind`（§B168/§B169 那张工程级表），带 §B169 同一枚 `resolveClassMemberCall` 否决；表答不出（数组/定长串/对象字段不入表）维持既有答案 ⇒ **不是**旧账担心的那一动（改表达式类型 oracle 扫全语料选档）。A/B = 410 输入 × 两架构 = 820 捕获，**same=816 / changed=4 / 未归因=0 / 跑挂=0**，移动的 6 行全是字段读换档（b168 的 W1/W9 两台 + b169 的 ReadBack 两台），且 b168/b169 夹具期望读数一字未动。
- **判据**：新夹具 `tests/b170_mrt`（三形撞名 + 不撞名对照 + With + 类内两腿 + VarType/TypeName + 槽里换数字）两台真跑 9 行；发码针 5 + 反针 5（BASE 上五枚反针全在场、NEW 上 0 在场）；哨兵 `check_variant_slot_store.ps1` 长出 **R8**（读侧问点数恰好 2 + 名单外新问点即红），两条负控各红自己那条。⚠ 负控第一版把那一问**注释掉**没红 —— 哨兵数的是文本出现次数，注释也算 ⇒ **负控必须删文本**。
- **语料暴露面（粗算，别当结论用）**：内置名池 172 枚（`addBuiltinFunc` 66 + 发码 ident 表 165，去重 172）；真语料 `D:\vb_yqt4qPac` 33 份工程 / 9389 份源文件里 **24 份**声明过与池子同名的类成员，点号读 18230 处 —— 但 `.Left` 那 10689 处多数是控件属性（已被 case 头部 `controlPropType` 接管），按宿主归因还没做。
- **继承红（不是本刀的，但挡本刀的收线）**：`656b2086`（别人那笔 `e84a9515` = ComCtlsDemo 点 Command2 退出）合进 dev 之后 CI run 38404170829 已红 4 个 vbp 分片 + 形状门（ve_units / arr_cls_elem / cls_inh / cc_demo_exe / cc_demo_dll_external / ctrlslider / FlexGridX86，另 `sl_emitc_native` 缺针）。本机用**不含本刀**的 HEAD 源码编出的编译器复现同一红 = `frmUnits.c(288) error C2064`，而 A/B 证明 ve_units 产物带不带本刀逐字节相同 ⇒ 机制在他改的 `cgen_expr_call_com_bind.inc` / `cgen_setlet_set_prop.inc` 那一族；#594（含六十刀、不含他那笔）12/12 绿。⇒ 那 7 行形状**故意不登记**（夹具现在它根编不过，形状不是结论），等那一格修完由它自己登记；本刀的「验证轮全绿」在那之前拿不到。

### B171 「把这枚成员读交给要字符串 / 要槽的消费者」时，「这是不是一枚槽」还有第二位主人 —— 而且它按**文本**认（2026-10-14 第六十二刀 **已出** = 实现 `6bb5aea7`（其间把 dev 上别人的两笔合进来 = 合并头 `a096b2b9`），门 **#603**（run 38449801858、attempt 1、12 job 里 10 绿 + 两红，两红都是登记那一件：`Tests (compile)` 的 `[STATIC] manifest_coverage` = 新夹具没哈希，形状门 = 本刀 5 行 + 别人 53 行）→ 登记 `60f1f0db`（吃 CI 那台 head=a096b2b9 自己交的 414 行清单 = `f2dfd910`，本刀只动 5 行）→ 验证轮 **#604**（run 38452950277、head 60f1f0db、11 绿 + 唯一红 = 形状门那 53 行，按 CI 那台同一份清单复算 414/414：相同 361 / 哈希不同 53 / 缺席 0 / 多出 0，且那 53 行里没有一行是本刀的），探针 `.build/k62_p/{K62.vbp,use62.bas,Holder62.cls}`）

- **两台都没有 exe**：含 §B170 那台 6 枚 C2440，不含的那台 4 枚 C2440（两枚 `vb6_VARIANT→BSTR` + 两枚 `vb6_VARIANT→double`）。
- **病①（与撞名无关，两态都坏）**：`Public PZ As Variant` 整枚递给 `As Variant` 形参 ⇒ 装箱站发 `(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=h->PZ  /* class var .PZ field */})` —— 结构体进 BSTR 槽。站点 = `cgen_util_comwrite.cpp:746-752`（default 臂）与 `cgen_expr_call_arg_emit.inc:883-886`（按 `inferExprType` 分档、Variant 落 default）；那一格「是不是槽」问的是 `cExprIsVariant(文本)`，**认不出**带 `/* class var .X field */` 尾注的那一枚 ⇒ 该走的两条（交槽的地址 / `vb6_VariantToString` 解包）今天都没走。
- **病②（只在撞名时）**：`Len(h.PV)` / `Left(h.PV, 4)` 这类内在函数消费者，§B170 之后发**裸 `h->PV`** ⇒ C2440；同一站对不撞名的 `PZ` 发 `vb6_VariantToString(h->PZ)` ⇒ 对。⚠ **对 §B170 的一句话订正**：这两条腿在 §B170 之前是「静默答错」（`(int32_t)sizeof(h->PV)` 恒 8 / `vb6_CStrDbl` 把串值折成 0），类型答对之后变成**编不过** —— 严重度更高、可归因，但它是一处**新的编译失败面**：`tests/**` 的 A/B（816 同 / 4 变）里没有这种工程，**真语料里有没有要单独数**。
- **与 §B150 的关系**：第四十刀补格已把「晚绑定槽元素读的『认文本』那一问」收成一处；本刀这条是同一问的**第二份**，且这份直接决定发码臂 ⇒ 修法 = 两站改问那一处权威（并把「带 `class var .X field` 尾注的成员读」教进认槽那张出口），不许再各自 `find`。
- **开工前先量两件事**：① **已量**：`cExprIsVariant(` 在 src/ 里有 **80 处调用点、散在约 20 份文件**（最密的三个 = `cgen_expr_call_arg_emit.inc` 16 / `cgen_expr_binary.cpp` 9 / `cgen_expr_call_arg_variant.inc` 5），另有第二枚同名权威 `cExprIsVariantCarrier`（`cgen_util_ctrl.cpp:2773` = §B150 收的那一处）⇒ **修法不该在 80 处各补一次**，而是教 `cExprIsVariant`（`cgen_util_type.cpp:792`，Fix 038b-1 那枚「按 C 文本认槽」的出口）认得「带 `/* class var .X field */` / `/* With class .X field */` 尾注的成员读」—— 那两条尾注本身就是 §B168/§B169 那张表发出去的，答案已经有了，只是这一问没去问它；② 真语料里「把一枚类 Variant 字段整个递给 Variant 形参」有几处（决定这是硬错误族还是纸面族）。
- **开工前第二件已量**（2026-10-14 第六十二刀；census = `.build/k62_census2.py`，读数 = `.build/k62_census2.txt`。口径比第一版收紧了三处：只算**模块级**字段（过程内的 `Dim v As Variant` 不是成员读）、只算**工程类**（.cls/.frm 的块名，`uFile.BufferArray` 那种 UDT 成员整个剔掉）、成员名若同时是该类的 Sub/Function/Property 就按 §B170 那条否决剔掉）：33 份工程 / 5221 枚类 / 180 枚模块级 Variant 字段 / 2598 枚绑定到工程类的接收者名 —— 后两条是这把尺子的正面控制，它们非零才说明 48 不是数不出来。
  - 「整枚类 Variant 字段落在调用实参位」= **48 处**，按接收者分两族：**38 处是内在函数**（`Val(iProp.Value)` 一族 32 枚、`CStr` 2、`Int` 1），撞的是 §B170 病② 那条腿，不是本刀那两站；**10 处是工程过程或 API 声明**（`Childs.Add TempNode, TempNode.key, Index` 五枚 + `GetPathFromIDList(Node.key)` 四枚，全部出自 ucTreeView 那一族 Node 的 `Public key As Variant`，声明在 Node.cls:20，被调方形参在 Node.cls:236 `Optional ByVal key As Variant`）。⇒ **本刀那两站在真语料里的暴露面 = 10 处，且全压在一枚类的同一个字段上** —— 是硬错误族，但不是 widespread 族。
- **§B170 那条「静默答错 → 编不过」的新编译失败面，真语料存量 0 处**（同一把尺子顺带量的）：模块级类字段的名字撞进那 172 名内在函数池的共 **77 枚，其中 Variant 档 0 枚**（77 枚全是 String/Long/Integer/Object 那些档，如 `fRender.Shell` / `frmdata.Str` / `node.Left`）⇒ 「撞名的 Variant 字段被内在函数消费者拿走」这一形在存量语料里根本不存在；那 77 枚的类型答案由「内在函数的返回档」改成「声明档」，方向是改对，不是改坏。语料工程并非台台能编，这一条只到「形状计数」为止，不冒充真编译读数。
- **修法**（一句话：那一问只许有一处按文本认的出口，两条发码路只盖词、不判断）：`cExprIsVariant`（`cgen_util_type.cpp:806` 那枚 Fix 038b-1 的「按 C 文本认槽」出口）此前只认得「内置函数交槽」那一族（前缀表）与「数组元素是槽」那一族，成员读一律答否 ⇒ 装箱站把整只 `vb6_VARIANT` 塞进标量槽。本刀把这一问收成 `cExprIsVariantSlotFieldRead` 一处，判据钉在**串尾**那个 ` field variant */` —— 不许用「包含」判：`vb6_CStr(h->PZ /*…*/)` 那种**已经消费过槽**的表达式会被包含判误认成槽（反向 C2440）。词那一头仍只有两条成员读的发码路盖（点号 `cgen_expr_member_class_fallback.inc` + With `cgen_expr_with.cpp`），而且只盖 Variant 那一档 ⇒ String/Long/Object 那些档的尾注一字不动，「认得这个词」不会把它们误认成槽。
- **为什么是交地址而不是解包**（这是本刀真正的口径决定，两台实测）：可寻址的那一档走 `&(h->PV)`，另一档 `(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=vb6_VariantToString(h->PV)})` **也编得过**，但实测读数差三样 —— 串值都对（`[8|withform]`）；数值槽被静默压成串（交 `[8|7]`，VB6 与地址那一档交 `[3|7]`）；ByRef 回写整个丢（`Mutate` 之后再读还是旧值）。地址那一档与「一枚普通 Variant 变量作实参」同形（`&w`），也与 UDT 里那枚 Variant 字段那一格同一条（`cgen_expr_call_arg_emit.inc` 早就在发 `&(...)`）⇒ §B147 那条「静默答错比编不过更坏」在这一格是**反向成立**的：只把尾注教进 `cExprIsVariant` 而不给地址那一档，等于把硬错误换成静默错。
- **站号订正**（对上面「病①」那一段）：台账先前点名的 `cgen_expr_call_arg_emit.inc:883-886` **不是**这两站。真站点 = ① `cgen_expr_call_arg_emit.inc:1142`（Fix 092t 那条左值链的「实参是不是 Variant 字段」那一问）、② `cgen_expr_call_arg_emit.inc:1286-1308`（M22 那条 `cType == "vb6_VARIANT"` 默认臂的装箱）、③ `cgen_util_comwrite.cpp:746-752`（`packLetValueArg` 的 default 臂）—— 三格问的是同一句话。另有第四处不是装箱站而是**第二个主人**：`isDefinitelyVariantExpr`（§B147 的 AST 那一半）的成员读分支，此前靠 `symTab_.lookupModule(裸成员名)` 兜底，撞内置名时答「不是槽」⇒ 病② 那两条腿（`Len(h.PV)` / `Left(h.PV,4)`）。本刀让它与 §B170 给 `inferExprType` 补的那一格同口径同否决（宿主是工程类 + `resolveClassMemberCall` 为空 ⇒ 问 `classFieldKind`）。
- **判据**：新夹具 `tests/b171_slotarg`（`B171-1..12` + 一条内层 `inner vt=… s=[…]` 证人 + `B171-DONE`）**两台逐行相同**；撞名（`PV` 撞内置 `PV()`）与不撞名（`PZ`）必须同答 `[8|withform]`，数值槽 `[3|7]`，String 字段那枚证人 `[8|strfield]` 照旧走 BSTR 装箱，`Call Mutate(h.PV)` 之后 `B171-12=[8|changed]` 钉回写。发码针 6 枚（`Test-EmitcShape b171_class_slot_arg_shape`）+ 反针 5 枚（`Test-EmitcAbsent b171_class_slot_arg_lesion`）。⚠ 前序 12 枚尾注针（b168 的 7 + b170 的 5）必须按新串重钉 —— 逐行核对过：那 12 行删掉的只有 ` variant` 那个词。
- **哨兵**：`check_variant_slot_store.ps1` 长 **R9**（写 ` variant` 尾注恰好 2 处 = 点号 + With；问 `cExprIsVariantSlotFieldRead` 恰好 6 处 = 定义 + 委托 2 / 声明 1 / 装箱站 3，名单外新写词与新问点都即红），PASS 行多出 `slotword171 2 | slotask171 6`；**R8** 那份名单从 2 钉到 3（第三站是同一问的第三个读点，不是新抄一份）。两条负控各红各条（删 With 那站的词 ⇒ R9 写侧红；删 `cExprIsVariant` 里那行委托 ⇒ R9 读侧红），还原逐字节相同后复跑 rc=0。全仓 58 道哨兵跑一遍：57 绿 + 唯一红 = `check_manifest_coverage` 的 K1（新夹具未登记，流程性）。
- **护栏 A/B**（BASE = knife 61 那台 / NEW = 本刀那台，414 份输入 x 两台 = 828 次捕获）：same=820 / changed=8 / 跑挂 0。变的 8 次 = 4 份输入 x 两台 = b168 / b169 / b170 三份前序夹具 + 新夹具 b171，每一行按三类归因：尾注多那个词、实参改成取址、以及一处**顺带的收益** —— `vb6_VariantFromValue(vb6_VariantFromValue(me->PV …))` 那双层装箱塌成一层（Fix 092t 那一问认出槽之后不再补装）。b168/b169/b170 的 runtime 读数两台一字未动。
- **合并与重测**（这一格是本轮真正多出来的教训）：推门之前 dev 上连来两笔别人的（`e117919c` = UC 属性包 / Is 同一性 / frx 设计期那一族，`0aca30d4` = Err 清空点），两次都不是快进 ⇒ `9e637a99` + `a096b2b9` 两个合并头。合完**重编**了一台（`cmake --build` 只跟踪 `src/driver/c3rtl.rc` 本身，RTL 的 .c 变了必须 `touch c3rtl.rc` 才重嵌；这台的 `check_rtl_embedded` 一开始报 R1 三份不一致就是这个），再重跑 b171 两台 14/14、发码针 6/6 + 反针 0/5、以及 58 道哨兵，才推的门 —— 否则 #603 那 10 绿说的就不是合并后那台。
- **继承状态**（不归本刀、但挡形状门那一格）：CI 那台的清单在 `a096b2b9` 上相对仓里移动 56 / 新增 2，本刀只落自己那 5 行；**另外 53 行刻意不登记** —— 那是 `e117919c`/`0aca30d4` 两笔的发码（属性包与 frx 那族进了产物）。与 §B170 那 7 行的差别是：这批夹具在 CI 上现在**是绿的**（四个 vbp 分片 + 两个 bas 分片全绿 = 那一族的编译失败已被它们自己修掉），形状已定型，只欠它们自己的作者落账；登记权在别人手里，本刀不替它判。⇒ 本刀的「十二格全绿」在那之前拿不到，但 #604 那一红与本刀无关这一点是**按 CI 自己那份清单**证的，不是按本机那台。

### B174 「同一条臂的另一半方向：数值实参进 BSTR 形参」—— Fix 108c 那张「数值→CStr」名单里没有 InStr 一族，`InStr(n, "1")`（n As Long）编得过、跑到那一行直接 AV（2026-10-14 第六十五刀 **已出** = 实现 `d6b90496` → 合并 `e78e3bee` → 门 **#613**（run 38496676186、head e78e3bee、12 job 里 10 绿 + 两红都归登记）→ 登记 `5bc74e88`（移动 4 行 + 新夹具 2 行，420→422，哈希取自 CI 那台为 head e78e3bee 自己交的 `d4b2ca8e`）→ 再并别人两笔（`8886c238` = 类成员解析 memo 性能刀、`682c194d` = c3_build.ps1 默认 --incremental，后者只动脚本 ⇒ 合并后那台重编重跑一字未改）→ 验证轮 **#616**（run 38498746101、head 682c194d、12 job **全绿**，含 `Emit manifest` 那一格）；探针 `.build/k65_p/{K65.vbp,K65Main.bas}`、夹具 `tests/b174_numstrarg/{B174.vbp,B174Main.bas}`）

- **形已证（一台干净源码编的合并后编译器，`.build/C3.exe` = `0f8af1a8`）**：`Dim n As Long: n = 12: Debug.Print InStr(n, "1")` 发的是 `vb6_InStr(1, n, vb6_BSTR_FromStr(L"1"))` —— 整型 `int32_t` 直接交给 `wchar_t*` 形参；cl 只给 **C4047**（在 -W3 之外）⇒ **BUILD-RC=0、exe 有**，x64 真跑到那一行 **RC=139 原生 AV，一行都不打印**。VB6 那一形把 string1 隐式 CStr ⇒ 应答 1。⇒ 这是「静默把数值当指针」那一族（账 #229/#231/#122 同族）在内在函数实参位的新入口，比打印差异更响。
- **机制就在六十四刀刚改的那条臂的另一半**：`cgen_expr_call_builtin_argtype.inc` 的 Fix 108c 名单 `strArgPos108c` （Replace{0,1,2} / StrComp / StrReverse / Trim / LTrim / RTrim / UCase / LCase / StrConv / Mid{0} / Left{0} / Right{0}）是「哪个内在函数的哪一位实参该把数值折成 CStr」的**逐函数白名单**，InStr/InStrB 一条都没有（探针读数 `PC callee=vb6_InStr hit108c=0`）。⇒ 与 §B173 是同一句话：**同一句决定（「实参档位与形参档位不符就折算」）按方向被抄成两半，一半有名单、一半没有**。
- **拍的那一头 = 不补名单，把这一问改成问形参档 + 问类型，并且先并掉第二个主人**：动手前扫出这一问其实有**两位**主人 —— Fix 175（`cgen_expr_call_arg_emit.inc`，认 `calleeParams[i].isByVal && type == String`，形状走 `wrapToBSTR`）与 Fix 108c（`cgen_expr_call_builtin_argtype.inc`，13 枚白名单 + 自己一份 if/else 形状表，且在 `argList` **已拼成一整串之后**用 `replaceInArgList108c` 按「从上次命中往后找第 N 次出现」做文本替换）。所以本刀不是给 InStr 加第三条名单项，而是：**站点收成 Fix 175 那一站**（实参逐条还握在手上、且已经用 AST 问过形参档的那一站）、**形状收成 `wrapToBSTR` 那一位**（全仓数值→BSTR 的形状权威）、**「哪一位是字符串形参」收成两处答案**（有形参表 ⇒ 问表；没有 ⇒ 问新出口 `builtinNoParamTableStringSlot`），13 枚白名单与那截文本手术整段撤掉。
- **两处新出口（都在 `cgen_util_type.cpp`，`classFieldReadExpr` 之后）**：`typeTakesImplicitCStr(Vb6Type)` = 哪些档 VB6 会被隐式 CStr（Date/Double/Single/Currency/Decimal/Long/Integer/Byte/Boolean/LongPtr/SByte/UInteger/ULong/ULongLong ⇒ 两位旧主人两份清单的**并集**；串/对象/数组/Variant 一律不介入）；`builtinNoParamTableStringSlot(callee, argc, pos, firstArgType)` = 没有形参表的内在函数里哪几位是 VB 语义上的字符串形参，目前**只有 `vb6_InStr` 一枚**（`InStr(s1,s2)` / `InStr(s1,s2,compare)` ⇒ 0,1；`InStr(start,s1,s2[,compare])` ⇒ 1,2 —— 串位随「有没有写 start」整体移位，与 `builtin_fixup` 选 `vb6_InStr`/`vb6_InStrC` 时同问，只是那里问 AST 这里也问 AST）。
- **草稿被两份实测读数各打回一次（这两条是本格真正的教训）**：① v1 把「哪一位是字符串形参」去问**运行时 C 签名表**（`getRuntimeParamCType`）—— A/B 当场在 VBFlexGridDemo 上把 `LenB(Version)` 的 `sizeof(vb6_type_DLLVERSIONINFO)` 打成 `sizeof(int32_t)`：那张表里 `vb6_LenB` 那格的 `"BSTR"` 只是 `_Generic` 宏的第一个分支（`vb6rtl_builtin.h:19`），而 VB6 的 Len/LenB 对数值交的是**字节宽度**、根本不做 CStr（发码侧另有按类型的 sizeof 分派，Fix 065/084h/125/160w）。⇒ **Len / LenB / Asc / StrPtr 是刻意不进表的**，不是来不及。② v2 在 `arg_variant` 末尾新开了第三站、又自带一份形状表（Boolean 折成 `-1` 而不是 `True`）—— 那是把本刀要消灭的东西**再抄第三遍**，整段回退重做。站点与形状两位权威都必须只有一个，这条要在动下一个同类刀之前再读一遍。
- **判据**：新夹具 `tests/b174_numstrarg` 两台真跑 **18 针 + `B74-SN` + `B74-DONE` 逐行相同**（`InStr(n,"1")`=1 / `Len(n)`=4 / `InStr(d,"5")`=3 / `InStr(b,"-1")`=0 / `Left(n,1)`=1 / `Mid(n,2,1)`=2 / `Replace(n,"1","9")`=92 / `StrComp(n,"12")`=0 / `Trim(n)`=12 / 串与 Variant 与 UDT 槽那三形都=3 / `Len(s)`=8 / `Len(v)`=8 / `Left(q,10)`=4000000000（q As ULong）/ `InStr(2,s,"t")`=3 / `LenB(n)`=4 / `LenB(u) > LenB(n)`=True / 用户过程 `ShowNum(n)` 打 12）。发码针 8 条（含 `vb6_InStr(1, vb6_CStrLong(n), vb6_BSTR_FromStr(L"1"))`、`vb6_InStr(1, vb6_CStrBool(b), …)`、`vb6_Left(vb6_CStrULong(q), 10)`、`vb6_InStr(2, s, vb6_BSTR_FromStr(L"t"))`、`(int32_t)sizeof(n)`、`(int32_t)sizeof(vb6_type_U74)`）+ 反针 5 条（钉的正是病形 `vb6_InStr(1, n,`、以及「start 位不许折」`vb6_InStr(vb6_CStrLong(2), s,`、「Len/LenB 不许被这条规则吃掉」`vb6_LenB(vb6_CStr` / `vb6_Len(vb6_CStrLong`）。行为负控 = BASE 那台（`C3_base_k65.exe`）拿同一份源：`B174Main.c(42) error C2440`（double → BSTR）、**无 exe**。
- **哨兵（第 59 道 `scripts/check_implicit_cstr_station.ps1`）**：R1/R2 = 两位新问句的**调用点唯一**（各恰好 1 处，且必须在 `cgen_expr_call_arg_emit.inc`；声明/定义行按前缀排除）；R3 = `strArgPos108c` 全仓 0 处（名单不许复活）；R4 = argtype 那站的代码里不许再出现 `vb6_CStr`（文本手术不许回来）；R5 = `vb6_CStrLong((int32_t)(` 在 `src/backend` 0 处（第二份形状表不许回来）；R6 = census 下限（`gateCases>=14` / 形参表名单 >=1 / `wrapToBSTR`@emit >=1 / 站点形状数不许多）。PASS 行 `census: backendFiles=116 gateCases=14 slotTableNames=1 wrapToBSTR@emit=1 castShapes=0`。六条负控（`.build/k65_nc.py` + `-RootOverride` 在副本上做，真实树一个字节未动）各红各条，`rc=1 viol=1 hit=True` x6。
- **护栏 A/B（逐行归因，两份）**：① 仓内宽面（BASE = `C3_base_k65.exe` / NEW = 本刀那台，**844 次捕获**）= same 832 / changed 12 / OTHER 0 / 跑挂 0；那 12 行**逐行归完**，每一行都是「数值实参多了那一折」或其后继影响，分类器 = 每行 `vb6_CStr*` 记号数不许减少。② 外部真语料（`.build/k65_ab_corpus.py`，`D:\vb_yqt4qPac`）**240 次捕获** = same 230 / changed 10 / OTHER 0；那 10 次是**同一行 `Dir` 的 5 份拷贝**（`FD.lpszFileName() As Byte` 那族，属 §118 定长串旧账，本刀顺带给出的折法），已登记成下一格 §B175 的第一形。③ 独立复算 = CI 那台为 head e78e3bee 自己交的清单，移动的 4 行 + 新增的 2 行与本机 `k65_manifest_diff.py` 的结论一字不差。
- **口径记一句**：`.vbp`（带 .cls）那一份产物在本刀前后**形状未变** —— 本刀动的两问都只在内在函数调用点上，工程类成员调用走的是另一条臂；这条用来证明「覆盖面没外溢」，别当成"改动惰性"（§B123 那条教训）。
- **真语料暴露面 = 0 处**（`.build/k65_census2.py`，10140 份 .bas/.cls/.frm/.ctl/.pag 按 GBK 解码，带括号深度 + 串字面量跳过的分参扫描）：正控 = 语料里 `InStr` **恰好两个实参**的调用 **2642 处**（⇒ 不是没人写这一形）；首参是裸标识符且同文件里该名字声明成数值档的粗尺命中 **3 处**，**逐条绑到声明后三处全是假阳性**（两处真宿主是 `Public Function StripPath(t)` 的无类型形参 = Variant，`Dim t As Integer` 在别的过程里；一处是 `Dim i As String, v As String`）⇒ 与 §B172/§B173 同格：探针可达的硬错误，不是 widespread 族。⚠ 第一把尺子（`.build/k65_census.py`）把「三实参形 `InStr(iStart, s, ":")` 的首参 = start」也计了进来，报出 346 处 —— 那是**合法的 start 位**，不是病形；数完必须打样例眼看（第五十六刀那条 census 正控纪律），已按形状改细。

### B175 「没有形参表的那 66 枚内在函数，字符串形参位没人答」—— GetSetting/SaveSetting/MsgBox 的标题位/InputBox/SendKeys/SetAttr 全落在 Fix 175 看不见的地方：数值实参直达 `wchar_t*`，而 MsgBox 那一形**另有一位只吃文本的主人**（2026-10-14 登记新格，**量完，未开工**；探针 `.build/k66_probe.bas` → `.build/k66_emit.txt`、语料尺 `.build/k66_census.py` → 输出在案）

- **形已证（一台 = 合并后 `682c194d` 那台，`md5=a930b95b`，`--emit-c` 实测）**：`GetSetting` 三形都裸送 —— `vb6_GetSetting(s1, s1, s1, n1)`、`vb6_GetSetting(n2, s1, s1, s1)`、`vb6_GetSetting(s1, s1, s1, d1)`；`vb6_SaveSetting(s1, s1, s1, n3)`、`vb6_DeleteSetting(n4, s1, s1)`、`vb6_InputBox(n1, vb6_BSTR_Empty(), …)`（省略位已经补 `vb6_BSTR_Empty()`，唯独交上来的那枚不折）、`vb6_SendKeys(n2)`、`vb6_SetAttr(n3, 0)`、`vb6_MsgBox(s1, 0, d2)`（**标题位**是 BSTR，buttons 位是 int32 —— 这一形正好证明「按 C 签名反推位号」又一次会答错档）。整型/浮点直达 `wchar_t*` = 与 §B174 同一个入口，cl 只给 C4047、x64 跑到那行 AV。
- **根因就是六十五刀那句的**另一半**：这些 callee 在 `builtin_funcs.inc` 里是 `addBuiltinFunc(`（**只登记返回档、没有类型化形参表**，全仓 66 枚，§B173 数过），⇒ Fix 175 的 `i < calleeParams.size()` 那一问根本问不出「这一位是字符串形参」，而本刀新立的第二处答案 `builtinNoParamTableStringSlot` 目前**只有 `vb6_InStr` 一枚**。RTL 那头的签名是硬的：`vb6_GetSetting(BSTR,BSTR,BSTR,BSTR)`、`vb6_SaveSetting(BSTR×4)`、`vb6_DeleteSetting(BSTR×3)`、`vb6_MsgBox(BSTR,int32_t,BSTR)` / `vb6_MsgBox1(BSTR)`、`vb6_InputBox(BSTR×3,int32×2,BSTR,int32)`、`vb6_SendKeys(BSTR,int32)`、`vb6_SetAttr(BSTR,int32)`、`vb6_SavePicture(void*,BSTR)`、`vb6_AppActivate(BSTR,int32)`（数值 PID **另有出口** `vb6_AppActivateByPid` ⇒ 那一枚不算病形）。
- **第二个主人（这一条是本格最想要的靶子）**：`MsgBox n5` 实测发的是 `vb6_CStrLong` —— 折它的不是 Fix 175，而是 `cgen_expr_call_conv_output.inc` 里那段**只在 MsgBox 上、按实参文本前缀连判**的链（`vb6_ComGetIntProp` / `vb6_ComVtableGetInt` / `vb6_ComGetDoubleProp` / `vb6_VariantFromComResult` / `vb6_ComCall(` / 已折过的 `vb6_VariantToString` / `inferExprType==Variant`，最后 Fix 161e 那条兜底才走 `wrapToBSTR`）。⇒ 同一句决定（「数值/Variant 实参进 BSTR 形参 ⇒ 折算」）**又有一位主人**，且它吃的是文本、只吃 prompt 那一位（所以标题位漏了 `d2` 那形）。它站在 arg_emit **之后**（阶段序：arg_emit → arg_variant → named_args → builtin_fixup → opt_pad → conv_cstr → conv_numeric → **conv_output** → conv_varargs_lenb → builtin_argtype），拿的是已经拼好的 `args[]` 文本，撤它之前必须先量那六条文本前缀判据有没有承载 §B146 那一族「这串是不是裸 COM 结果」的**另一种问句**（六十五刀刻意把 argtype 那站的三条同类问句留着，理由相同）。
- **顺带量到的邻居（同一站的另一半方向，先记账不动）**：`vb6_InputBox(s1, s1, v1, …)` 里 `v1 As Variant` 也是裸送进 BSTR 形参 —— 那一位既没被 conv_output 的链吃到（它只管 prompt），也不在 Fix 175 的范围（`typeTakesImplicitCStr` 刻意不含 Variant，`cExprIsVariant` 先短路）。属 §B142/§B146「装箱档」那一族，**不并进本刀**：修法与判据都在那边，先记下这一形在这条路上也漏。
- **真语料暴露面 ≠ 0（`.build/k66_census.py`：`D:\vb_yqt4qPac` + 本仓，6 411 431 行，GBK 解码，只数**数字字面量**坐在 BSTR 形参位上 —— 字面量不需要类型推断，是决定性下限）**：合计 **461 处**。分位：**`GetSetting` 第 4 位（default）= 369 处**、`SaveSetting` 第 4 位（setting）= 78 处、`InputBox` 第 3 位 = 9 处 / 第 2 位 = 1 处、`Format` style 位 = 2 处、`MsgBox` prompt 位 = 2 处（其中一处正是本仓自己的 `tests/msgbox_num.bas:24` ⇒ Fix 161e 那一刀就是为了它）。调用总量做分母：`MsgBox` 16 070、`Format` 9 415、`GetSetting` 1 145、`SaveSetting` 1 080、`SendKeys` 857、`InputBox` 327、`SavePicture` 117、`SetAttr` 48、`AppActivate` 12、`DeleteSetting` 15。⇒ **与 §B172/§B173/§B174 三格不同，这一格不是纸面缺口**：`GetSetting("A","S","K", 0)` 是 VB6 注册表读写最标准的写法，语料里 369 处，而它今天发的是一枚裸 `int32_t` 当 `wchar_t*`。⚠ 这把尺子只数**字面量**；变量实参（`GetSetting …, nDef`）还要绑声明才能定档，真实暴露面只会更大 —— 报数时按"下限"口径写，别当总数。
- **修法取向（开工前先拍的三个问）**：① 把「哪一位是字符串形参」的第二处答案**补全**（`builtinNoParamTableStringSlot` 从一枚扩成一张表：`vb6_GetSetting{0,1,2,3}` / `vb6_SaveSetting{0,1,2,3}` / `vb6_DeleteSetting{0,1,2}` / `vb6_MsgBox{0,2}` + `vb6_MsgBox1{0}` / `vb6_InputBox{0,1,2}` / `vb6_SendKeys{0}` / `vb6_SetAttr{0}` / `vb6_SavePicture{1}` / `vb6_Format{1}`，每行都要有一句「VB6 收了数值会隐式 CStr」的读数依据，不许从 C 签名表反推 —— 六十五刀那条教训①就在这张表的隔壁位上）；② `conv_output` 那段 MsgBox 链**撤不撤**：先 A/B 出那六条文本判据各自的承载面，能撤就并到 ①，不能撤就至少在 R 侧把「同一个问句有两个主人」钉成结构性报红；③ 判据面要两头 —— 病形反针（`vb6_GetSetting(s1, s1, s1, n` 这类）+ 「不许多吃」的反针（`vb6_MsgBox(s1, 0, …)` 的 buttons 位、`vb6_AppActivate` 的 PID 出口、`Mid(s1, n3, 1)` 的数值位、以及 `Len/LenB` 那一族字节宽度答案）。
- **A/B 预告**：`GetSetting` 那 369 处意味着 `--emit-c` 宽面一定会动；上一格同类刀的教训是「changed 全要逐行归因」，这一格的归因预算按**每处 4 位**算，别让 OTHER 名单没人管。

### B173 「同一句决定抄了两遍，第二遍只吃文本」—— 内置/运行时那条实参转换臂没问 AST，UDT 的 Variant 成员进 `InStr` 就是整枚槽进 BSTR 形参（2026-10-14 第六十四刀 **已出** = 实现 `2cbc8156` → 门 **#607**（run 38475899573、head 2cbc8156、attempt 1：12 job 里 10 绿，两红都归登记 —— `Tests (compile)` 的 [STATIC] manifest_coverage 报新夹具两行没哈希、`Emit manifest` 报形状差）→ 登记 `3091eee9`（只补本刀那 **2 行**，416→418，哈希取自 CI 那台为 head 2cbc8156 自己交的 `22882407`）→ 合并别人那笔后重编重测（merge `475ea5e6`）→ 验证轮 **#609**（run 38478861097、head 475ea5e6、attempt 1）= **10 绿 + 两红，两红的注解只点别人的名**（存量 53 行移动 + 他们那两行新夹具 `tests/test_uintfields.vbp`/`tests/tu_main.bas`），本刀的 b173 两行已从红名单消失、vbp 分片全绿；探针 `.build/k64_p/{K64.vbp,K64Main.bas,K64Holder.cls}`、夹具 `tests/b173_udtarg/{B173.vbp,B173Main.bas,B173Holder.cls}`）

- **站点归因（三枚 `fprintf` 探针重编一台，一次跑完）**：`cgen_expr_call_arg_variant.inc` 里「这枚实参是不是整只 `vb6_VARIANT` 槽」这句决定**抄了两遍** —— 上面那条臂（Fix 029，形参类型已知，体在 `:79-226`）第一问是 `isDefinitelyVariantExpr`（AST）再补 `cExprIsVariant`（文本）；下面那条臂（Fix 038b-2，内置/运行时：`calleeParams` 为空或档记成 Variant/Empty，体在 `:279-445`）第一问**只有文本**。谁落哪一遍由「这个内在函数在 `builtin_funcs.inc` 里有没有登记**类型化形参表**」决定，与实参形状无关：`Len`/`Left`/`Right`/`Replace`/`Val` 走 `addBuiltinFuncWithParams` ⇒ 上臂 ⇒ AST 认得 UDT 成员（Fix 090l 那一格，`cgen_util_type.cpp:752-754`）⇒ 解包；`InStr` 只有 `addBuiltinFunc("InStr", Vb6Type::Long)`（start 可选 ⇒ 没登记形参表）⇒ 下臂 ⇒ 文本认不出 `u.UV` ⇒ 裸 `vb6_InStr(1, u.UV, …)` = **cl C2440**，两台各三枚、都没有 exe。⇒ §B147「改道吃 AST ∪ 文本」在这一路是破的，而破的那一半恰是**AST 那一位能答、文本那一位答不出**的那一档。这条臂的管辖面有多少：`builtin_funcs.inc` 里 **66 枚**只登记返回档（`addBuiltinFunc(`，无类型化形参表）而 **92 枚**带形参表（`addBuiltinFuncWithParams`）——前者全部落在这一臂上，含 InStr / Format / MsgBox / InputBox / Array / Command / Now / Timer / VarPtr一族。
- **探针读数（决定性那六行，BASE 那台）**：`PA callee=vb6_InStr i=0 rt=BSTR arg=u.UV`（下臂收到裸槽）+ `PC callee=vb6_InStr hit108c=0 arg=u.UV`（Fix 108c 那张「数值→BSTR」名单也不管 InStr）；同一枚 `u.UV` 在 `PC callee=vb6_Len / vb6_Left / vb6_Replace / vb6_Val` 那里**到达时已经**是 `vb6_VariantToString(u.UV)`，且这四枚**一条 `PA` 都没发**（⇒ 它们在上臂就被办完了，`rtParamTypeUsable` 那一位为假）；`PA callee=vb6_CStr i=0 rt=vb6_VARIANT` 与 `PA callee=vb6_Format i=0 rt=`（形参就要整枚槽 / 运行时表里没登记）都直送 = 对。⇒ 坏的不是「InStr 那一族特殊」，是**只有它落在这条只吃文本的臂上**。
- **修法（拍的那一头 = 让下臂也问 AST）**：下臂第一问改成 `isDefinitelyVariantExpr(*node.positional[i]) || cExprIsVariant(argVal)`（`:294`，与上臂 `:88` 同一条口径）。下面那三枚手抄（裸 COM 前缀 / `knownVariantVars_` / Fix 089i 查被调符号返回档）**刻意留着**当并集的另一半 —— 089i 对「模块 Function 但返回档还没推断出来（`Unknown`）」那一格比 AST 那位宽，撤它要另花一份读数，不在本刀。
- **判据**：新夹具 `tests/b173_udtarg` 两台真跑 **19 针 + DONE 逐行相同**（`InStr` 九形全答 3；证人 = 数值槽找 "7" 答 1 / 找不着答 0 / `Len`=8 / `Replace`=withf0rm / `CStr`=withform / `Left`=wi / `Val`=0 / String 变量 `s` 继续裸送）。发码针 7 条 + 反针 4 条，反针钉的正是病形 `vb6_InStr(1, u.UV,`（BASE 产物命中 3 次、NEW 0 次），另三条钉「不许二次包装」「String 变量不许被误认成槽」「上臂那一路一字未动」。BASE 那台拿同一份源两台都 **cl C2440 ×3、无 exe** = 行为负控。
- **哨兵（`check_variant_slot_store.ps1` 长出 R10）**：「实参是不是整枚槽」那一问的 AST 位全仓钉成 `cgen_expr_call_arg_variant.inc`=2（两条臂）+ `cgen_expr_call_arg_emit.inc`=2，名单外新增即红 ⇒ 下一次有人**再抄第三遍**会自己撞出来。三条负控各红自己那一条（下臂撤 AST / 上臂撤一处 / 别处新加一处），都在 `src/backend` 的副本上做，真实树一个字节未动（`.build/k64_nc.py`）。
- **真语料暴露面 = 0 处**（`.build/k64_corpus.py`，14340 份 .bas/.cls/.frm/.ctl/.pag 按 GBK 解码）：正控 = `InStr` 一族在全语料出现 **6525** 次（⇒ 不是没人写）；带 `As Variant` 成员的 UDT 只有 **17** 份源；粗尺命中 4 处逐条看全是控件的 `.Tag`（`Val(tmrHighlightIcon.Tag)`，同名不同物）⇒ 与 §B172 同格：**探针可达的硬错误，不是 widespread 族**，粗尺只用来证伪「没人这么写」。
- **顺带量到的下一格 = 已登记 §B174**（本刀没动）：`InStr(n, "1")` 里第一参是 Long ⇒ 发 `vb6_InStr(1, n, …)`，整型进 `wchar_t*` 槽，cl 只给 C4047、x64 真跑**直接 AV**（本机实测 RC=139）；`vb6_InStr` 不在 Fix 108c 那张「数值→BSTR」名单里。**订正**：本格初稿写过一句「真 VB6 那一形报 Wrong number of arguments」—— 那句没量过，按 VB6 的 InStr 语义应是 string1 隐式 CStr（n=12 ⇒ 答 1）；口径 = 「补那一档」还是「按规则取代名单」，见 §B174。


### B172 「成员读还有第三种拼法」—— 类形参那一档发的是 `(*h)->PV`，两条认槽判据都看不见它（2026-10-14 第六十三刀 **已出** = 实现 `1a83dbea` → 门 **#605**（run 38465837733、attempt 1、12 job 里 10 绿 + 两红都归登记）→ 登记 `dd37e987`（补 2 行，414→416）→ 验证轮 **#606**（run 38467986037）= 11 绿 + 唯一红 `Emit manifest`，那 53 行是别人的；探针 `.build/k63_p/{K63.vbp,K63Main.bas,K63Holder.cls}`）

- **同一枚夹具、同一枚字段，两种接收者拼法给出两种命运**（两台实测）：
  - 局部对象变量（`Dim h As New K63Holder`）那一形 ⇒ `vb6_InStr(1, vb6_VariantToString(h->PV  /* class var .PV field variant */), …)` **对**（第六十二刀那条判据吃得到它）；
  - 类形参（`Private Sub UseAll(h As K63Holder)`）那一形 ⇒ `vb6_InStr(1, (*h)->PV, …)` —— **cl C2440（`vb6_VARIANT` → `BSTR`）x64 与 x86 各一枚，两台都没有 exe**。
- **不是第六十二刀带进来的**：BASE（knife 61 那台）与 NEW（knife 62 这台）在 `K63Main.c` 那一行**逐字节相同**（同一份源、只换编译器）。⇒ 这是一条一直在的病，本刀的尾注只是让**另外两形**能被认出来，`(*h)->X` 这一形既不写尾注（文本判据答否），也没落进 §B170/§B171 补的那两格 AST 判据 —— 同一枚 `h.PV` 在不同接收者拼法下由**不同的发码站**发出去，而 §B169 那张表只被其中两站问过。
- **覆盖面 = 19 枚内在函数消费者实测，只有 InStr 那一族坏**：`Val / CStr / Int / Fix / Len / Left / Right / Mid / UCase / Trim / Asc / InStr / CBool / CDbl / CSng / CInt / Abs / Sgn / Format` 都拿同一枚 `(*h)->PV` 走一遍 —— 除 `InStr` 外全部正确（`vb6_Val(vb6_VariantToString(…))`、`vb6_Len(vb6_VariantToString(…))` 那两形在参数脸上也对；`vb6_Format((*h)->PV, …)` 与 `vb6_CDblV/CSngV/CIntV/CBoolV((*h)->PV)` 本来就要整只槽 ⇒ 直送是对的）。⇒ 机制收窄成一句：**InStr 那一族的实参转换只问「按文本认槽」那一位主人，没问 AST 那一位** —— §B147 立的规矩（oracle ⊆ 改道，改道可以吃两半）在这一格是破的。
- **真语料暴露面 = 0 处**（两把尺子）：① 第六十二刀那把（只算模块级字段 + 工程类 + §B170 的裸成员名否决）扫出的 48 处「整枚类 Variant 字段落在调用实参位」里，内在函数那一族 38 枚全是 `Val`/`CStr`/`Int`，**`InStr` 一条都没有**；② 新加的粗尺 `.build/k63_census.py`（正面控制：10138 份源、`InStr` 出现 6372 次、首参是带点成员读的 621 行）逐行看，那 621 行是 UDT 的定长 String 成员（`OFN.lpstrFile`）与属性调用（`Me.FListItemText(i, j)`）两类，都不是这一形 ⇒ 粗尺只用来证明「不是没人写 InStr，是没写在这一形上」，不冒充绑定分析。
- **修法取向（口径要先拍，未定稿）**：不该再给 InStr 单独补一句特判。两条候选其实是同一句话的两头 —— ① 让 `(*h)->X` 那条成员读也问 §B169 那张表并**盖同一个词**（写词那一头从 2 处变 3 处 ⇒ 哨兵 R9 的名单要跟着长，而「成员读有几条发码路」这一问第一次被证明数错了）；② 让 InStr 那一族的实参转换改问 AST 那一位主人（`isDefinitelyVariantExpr`）。拍哪头之前先把**「成员读的发码站到底有几条」**量准 —— 这正是 R9 想钉的那件事，本格的读数说明它钉的名单本身就是缺的。
- **落地（本刀选的口径 = ①，且把「拼尾注」这一步本身也收了）**：发码站实测**三条** —— 点号那一站（局部对象变量 `h.PV`）、Fix 088c 那一站（类形参 AST 推出的 `(*h)->PV`）、With 体内那一站。修法不是在 InStr 上补特判，也不是只给第三站补一个词，而是把「拼这条成员读尾注 + 它是哪一档」这**一步**搬进唯一出口 `CCodeGen::classFieldReadExpr(objC, clsName, memberName, isVoidPtr, markerLabel)`（定义在 `cgen_util_type.cpp`，声明在 `cgen_helpers.inc`），三站一律只调它、自己不再手拼也不再查表；档位那一问仍走 §B168/§B169 那张表（`classFieldIsVariantSlot`）。⇒ 哨兵 R9 的名单从「恰好 2 处」订正成「写词 1 处（唯一出口）+ 发文站 3 处（新增 R9c，按 `= classFieldReadExpr(` 数）」，R5d 那处手抄计数随两处生产者搬家从 3 → 1（11 → 9）。
- **判据与读数**（新夹具 `tests/b172_paramslot`，两台真跑 x64/x86）：十条读数 `3,3,withf0rm,1,3,0,8,3,3,withf0rm` + `B172-DONE`，两台逐字节相同。发码针 7 条 + 反针 6 条（反针钉的正是病形 `vb6_InStr(1, (*h)->PV`）。**证人那条是 String 档字段** `PS`：它现在也带尾注、但**不带** ` variant` 那个词 ⇒ 仍按裸 BSTR 直递、不解包（钉住「认得尾注」不会把别档误认成槽）。
- **护栏（窄面 A/B，两台 = 256 次发码）**：BASE = 同一棵树只缺本刀那四份 src 编出的编译器（`8c5c3e07…`）、NEW = 本刀这台（`f6c7d898…`）；拿「过期 BASE 那一轮曾移动的 128 份捕获」做窄面 ⇒ **same=126 / changed=2 / OTHER=0 / 跑失败=0**，动的只有新夹具那两台，前五刀的全部存量产物逐字节相同（本刀那两站的重写是文本中性的，这条被证明）。上一轮宽面（416 输入 × 两台 = 832 份，`.build/k63_ab_report.txt`）里其余 118 份移动全部逐行归因到别人已并进 dev 的三刀（`0aca30d4` ErrClear / `333fec1b` 双括号 / `65f246c0` UC 属性包），与本线无关。
- **门与验证轮的读数**：#605（head `1a83dbea`）两道红都在登记那一族 —— `Tests (compile)` 的 [STATIC] manifest_coverage 报新夹具两行没哈希（本机 58 道哨兵当时也是同一道红），`Emit manifest` 报形状差；登记 `dd37e987` 只补本刀那 **2 行**（414→416，注脚写明来源 = CI 那台为 head 1a83dbea 自己交的 `1b240bd1`）。验证轮 #606（head `dd37e987`、run 38467986037、attempt 1）= **11 绿 + 唯一红 `Emit manifest`**；拿 CI 为 dd37e987 交的清单（`9f8ccca3`）独立复算：**ci=416 / expected=416 / same=363 / differ=53 / added=0 / vanished=0**，而那 53 行里**没有一行**属于 b15x/b16x/b17x 任何一份本线夹具 ⇒ 剩下的红与本刀无关，登记权在别人手里（与 §B171 收线时同一格局）。
- **本机 58 道哨兵**：`check_manifest_coverage` K1 随登记转绿（inputs=416 / registered=416 / sets equal），其余 57 道全绿；R9 的读数是 `slotword171 1 | slotst172 3`（写词一处、发文站三处）。
- **继承状态**：dev 自 `656b2086` 起带别人那几道红，本线门号只看自己 head 那一轮。
- **本刀剩下的那一头（不在本格，已另立 §B173）**：InStr 那一族的实参转换至今**只问文本那一半**（`cgen_expr_call_arg_variant.inc` 只吃 `cExprIsVariant`）。本刀把第三站的尾注补上之后两半在这一形一致了，但「AST 说是槽、文本认不出」这一极在别的表达式形状上仍然出得来 —— 实测 **UDT 的 Variant 成员**（`u.UV`，C 里就是整枚 `vb6_VARIANT`）发成 `vb6_InStr(1, u.UV, …)` ⇒ 同一枚 cl C2440（探针 `.build/k64_p`，`K64Main.c(39)`，两台同形）。

### B151 把真二维数组赋给 Variant 槽，槽读回来是 VarType=9 / "Object" 而不是 8204 —— 自家多维描述符被折成一维（2026-10-12 第四十刀量到；**未开工**）

- **读数**（探针 `.build/b1268_ndread/{tnd,tag}.bas`，`.build/b1268_ndread/rt.txt`）：`Dim a(1 To 2, 1 To 3) As Variant: v = a` 之后 `VarType(v)` 交 **9**、`TypeName(v)` 交 **"Object"**；VB6 交 8204 / "Variant()"。而同一枚 `v` 的 `UBound(v, 2)` 那一路是按 ND 问的 ⇒ 装箱那一档少认多维。
- **与 §B150 的关系**：本刀那三条 ND 出口在真语料里唯一**真跑到**的二维形状是 §B149 新造的 `GetAllSettings`（我按 VB6 口径造了真 ND 描述符）；VBFlexGridDemo 那枚 `ArrRows` 走的是折叠 1D ⇒ 靠 §B150 那条兜底维持既有读数。**没有这一格，兜底那条就是永久项**。
- **覆盖面已量（2026-10-12 第四十一刀之后，两把 census 都在 `D:b_yqt4qPac` 上跑，10138 份 .bas/.cls/.frm/.ctl/.pag，按 GBK 解码）**：① `.build/b1358_nd3.txt` 只吃**本格那一形**（左值声明成 Variant、右值是同一份文件里声明的固定二维数组、且没被同名标量遮蔽）⇒ **命中文件 0 / 命中行 0**；② `.build/b1357_nd2.txt` 放宽到「二维数组名不带下标出现在任何一行」⇒ 268 行，逐行看是**别的两类**：整枚数组按引用传进过程（`m3MatMultiply M12, M1, M2` = 形参 `M() As Single`，不是装进 Variant 槽）与同名标量的遮蔽（`Dim R2 As Single` 与 `Dim R2(3,3)` 撞名）。⇒ 按「先证明可达再动」的口径，本格**继续登记、不实施**；§B150 那条兜底（魔数 0x5A1D 退回一维）撑得住。真正有量的是**动态**那一族（二维 `ReDim` 414 处、二维 `Dim` 172 处），它撞的是元素读写与 `UBound(a, 2)`，与本格的装箱档不同形——那才是下一刀该问的那一格。

### B145 Decimal 的精确算术：Variant 参与 `+ - * /` 时**不该在编译期猜数值档**（2026-10-11 第三十六刀 **已出** = 提交 `b142d338` + 补课 `67f03217`；口径是用户当场拍的「要做精确的 decimal 算数」；门 #517 一台真红已归因修掉（补课 `67f03217`）；下一跑 = **门 #522**（head `67f03217`，12 job 里 11 绿、唯一红是形状门）。**注意**：#522 那次的形状门报的 10 行里有 5 行是静默回归（**§B147**），已在第三十七刀修掉，清单最终只登记五行（门 #527 / 登记 `41bce27a` / 收线 #528））

- **题面比 Decimal 宽**（探针 `.build/b1095_arith/probe5.bas`，两台逐行相同的 BASE 读数）：发码层把 Variant 操作数在 `+ - *` 上发成 `vb6_VariantToLong(...)`（`cgen_expr_binary.cpp:767`，Fix 086 那一格），只有 `/` 发 `vb6_VariantToDouble`。于是盒里是 **Double** 时也一起坏 —— `1.5+1` 答 **3**、`1.5*2` 答 **4**、`1.5-1` 答 **1**、`0.1+0.2` 答 **0.2**；是 **Decimal** 时 `1.25+1` 答 **2**、`1.25*2` 答 **2**、`1.25-1` 答 **0**、`0.1+0.2` 答 **0**、`TypeName(v+1)` 答 **Long**。而静态类型 `As Double` 的对照组一直对 ⇒ **坏的不是 RTL 的取数口，是发码层替运行期选一档**。
- **口径**（这一格回答 §B142 那个没拍的「各表手换还是统一交 oleaut32」）：交 oleaut32 的 **Variant 级**算术引擎 `VarAdd/VarSub/VarMul/VarDiv` —— VB6 自己的 Variant `+ - * /` 就是这台。Decimal 的 96 位精确、Double 保小数、Null 传播、Date 加天数、舍入与越界政策一起继承，运行期不再有任何一档答案住在我们这里。两台逐行相同（纯 C 探针 `.build/b1097_vararith/report.txt`）：`R8(1.5)+I4(1)` → `2.5`；`DEC(1.25)+I4(1)` → **vt=14 DECIMAL scale=2 lo=225**；`DEC(0.1)+DEC(0.2)` → **DECIMAL 0.3**；`DEC(1)/DEC(3)` → **scale=28** 的 Decimal；`I4+I4` → 仍 `VT_I4`（整型那一族不动）；`Null+1` → Null；`Date+1` → Date；`"abc"+1` → `DISP_E_TYPEMISMATCH`；`DECmax+1` → `DISP_E_OVERFLOW`。
- **订正 §B142 第三轮那句「`VarDec*` 一族在本机够不着」**：它只对 **To** 那一族成立。算术那一族**够得着** —— `oleauto.h:860-873` 声明 `VarDecAdd/Div/Mul/Sub/Neg/Round/Cmp/CmpR8`，x64 的 `oleaut32.lib` 存裸名、x86 那份存 `_VarDecAdd@12` 这一族，两台实测能链能跑（`.build/b1093_decapi/report.txt`）。⇒ 「各表手换 96 位算术」那条路本来就没出局；本刀仍选 Variant 级那一族，理由是**答案更少、继承的语义更多**。
- **标签纪律按那一族分家**（同两台实测）：`VarDecAdd` 那一族**不写** `DECIMAL.wReserved`（拿 `memset(0xAA)` 铺底去接，结果里那两字节仍是 `0xAA`）⇒ 用它必须自己补 `vb6_DecSetTag`；而 `VarAdd` 那一族交回的 VT_DECIMAL **自带 14**。⇒ §B143 第五轮「标签必须在造出来那一步末尾统一写」仍成立，只是「谁造的那一步」现在有两族答案。顺带一条 `VARCMP_*` 的读码坑：SDK 是 `LT=0 EQ=1 GT=2`（`oleauto.h:910-913`），**不是** 1/0/2 —— 探针里等值两枚 DECIMAL 答 1 是对的，别据此判红。
- **错误号映射按运算符分家**（两台实测，本片唯一一处本地差异）：`1.0/0` 在 **x64 交 `DISP_E_DIVBYZERO`、x86 交 `DISP_E_OVERFLOW`** ⇒ 两条码都要认、都映射 VB **11**；而 `DISP_E_OVERFLOW` 在 `+ - *` 是 VB **6**（Overflow）。`DISP_E_TYPEMISMATCH` → VB 13。把加减乘的 Overflow 报成 11 就是第二个答案，所以这一格不许收成一条。
- **落地**（一份提交：RTL 三份 + 发码两份 + `c3rtl.rc` touch + 哨兵 + 夹具与针面）：
  - RTL：四枚出口 `vb6_VarAddV/SubV/MulV/DivV` 转调一枚 static `vb6_VarArith(op, a, b)`（`vb6rtl_com.c`）。实参用**借用**语义装箱（绝不 `VariantClear` 那两份 —— 那会把槽自己的 BSTR / 对象 Release 掉），结果走既有的 `vb6_VariantFromStackVARIANT`（它复制 BSTR / AddRef 对象，所有权留在栈上的结果里，随后 `VariantClear`），末尾再补一次幂等的 `vb6_DecSetTag`。
  - RTL：「槽 → 真 VARIANT」那张 vt 表收成唯一一份 `vb6_VariantPackBorrow`，并补上 §B142 census 点名的 **Decimal / Currency / Date 三格**（以前都落 `default` ⇒ 静默交 `VT_EMPTY`）；`vb6_ComPackVariant` 改成在它上面补一步「借用的串 → 带权的副本」，COM 那一侧每一格逐字未动。两条「往里读」的 x86 Decimal 臂改从结构体的**头**铺 16 字节（并撤掉那句把 `pdecVal` 说成指针的注释与那个判空）+ 补 `VT_I8` 档（缺它时 `vb6_VariantFromComResult` 的 `default` 走的是**数组**那段，而 `VarAdd(VT_I8, VT_I4)` 本来就交回 vt=20）。⇒ **§B143 在册的两格就此结**；§B144 的「往里读」那一半也结，只剩服务端那一处 `memcpy`。
  - 发码：`cExprIsVariant` 前缀表登记四枚出口；`Add/Sub/Mul/Div` 一处改道，判据从 `extractVar86` 里提出来共用（`isVariantOperand86`，不再抄第二遍）；**任一侧静态是 String 的那一族刻意不接管**（VB6 有一侧是串就走拼接，而 `VarAdd` 按数值强制 —— 实测 `"5"+1` 交 6）。
- **哨兵跟着改**：`check_variant_i8_coverage.ps1` 的名单随表走（`vb6_ComPackVariant` → `vb6_VariantPackBorrow`，Decimal 的 must 7→8、exempt 6→5），读数 `switches 13 I8 must 11 exempt 2 | DEC must 8 exempt 5 | violations 0`。**并且站点匹配前先剥注释** —— 本刀写进 `vb6rtl_com.c` 的一句注释（它引用了 `switch (v.vt)` 这个形状）自己数出来一个站点、还被算给上一个函数 `vb6_VariantToObject`，两条 V2/D2 当场报红。phantom 站点比漏站点更坏：它凭空造出一格「要豁免」⇒ 判据是**注释里的形状不算形状**（这条是那道「门只认一种写法」的老坑的反面）。
- **判据**：`$boxNeedles` 43→53（十枚 `VB-ar-*`），两台各 **53/53 命中、零 miss**；负控 = 上面那批 BASE 读数（同一批 VB 形状，改前答 3/4/1/2/2/0/Long）。平局一条不钉（2.5/3.5 属 §B77① 未拍口径），Long 目标那条用 `CDec("5.75")+1 = 6.75 → 7`。消费侧四种形状都编译过（标量目标 / 比较 / 数组下标 / 数值形参）；真工程 `VBFlexGridDemo` 两台各 0 条 error、exe 出得来（`.build/b1107_vfg/report.txt`）；`check_rtl_embedded` 125/125（exe sha `2d1e5b061697`）、RTL 四道哨兵 + `check_manifest_coverage`(inputs=400 registered=400) 全过。
- **门 #517（head `b142d338`）= 12 job 里 10 绿，两条红**：形状门（照预期 —— 这一刀动了发码，登记留给 CI 自己那份读数）+ **`Tests (vbp #2)` 一片真红 = Charts 2020 `ClsResizer ... FAIL (compile)`**。归因（本机复现 `.build/b1112_charts`，两架构同一份错）：`ClsResizer.c(144)/(150)` 报 `error C2440「无法从 vb6_VARIANT 转换为 double」`，落点是 `Prop.Add CallByName(...) * 100 / 225` 那一族的 **COM 实参打包** —— `comPackExpr` 按 `inferExprType` 选打包器，而 `inferExprType` 的 Div 那行是**无条件答 Double** 的特例，本刀把 Variant 的 `/` 发成交运行期出口之后，那个位置的 C 表达式已经是整只槽了 ⇒ 选了标量形参的 `vb6_ComPackDouble(槽)`。修法一处：Div 那行先看两侧，任一侧 Variant 就答 Variant（`+ - *` 一直走 `promote`，那里本来就写着「Variant 参与的结果总是 Variant」，所以只有 `/` 这一条是自己抄出来的特例）。= 提交 `67f03217`。**本机复测**：Charts 2020 两台 rc=0 且 exe 出得来、`test_variant_boxing` 两台 53/53 零 miss、VBFlexGridDemo 两台 rc=0。
- **通用的一条**（这条比本格更值钱）：**改发码的 C 类型，就必须同时改那条表达式的类型 oracle** —— 发码里有半打消费点问的是 `inferExprType` 而不是 C 文本（COM 打包是其中一个），两边不一致时不是"少优化一点"，而是**编不过**。判据面看得见它，是因为语料里真有一枚 `Variant / 数` 当 COM 实参用；A/B 400 份就是为这种形状准备的（本机三枚夹具全绿放行不了它）。
- **但 oracle 那补课只修好了一极**：它把「发码已是槽、oracle 还答 Double」的 **响** 那一极（门 #517）修掉了，却造出了反方向的 **静默** 那一极（浮点->窄整型目标丢取整）—— 读数与根因在 **§B147**，清单因此暂不登记。
- **试过并撤掉的另一种形状**（记下来免得再走一遍）：把五枚 typed 打包器（`vb6_ComPackBSTR/Int/Bool/Double/Object`）改成 _Generic 路由器 —— 真函数挂 `...Raw`，旧名吃 `vb6_VARIANT` 那一档交 `vb6_ComPackVariant`（发码文本一字不动、覆盖面比 oracle 那条宽，形状先例是 `vb6_UserControl_ScaleX` 那一族）。本机两架构都改红成 `LNK2001 无法解析 _vb6_ComPackBSTRRaw`，而拿 dumpbin 对着**那次链接的输入对象**查过：符号确实定义在其中，且只有 BSTR 这一枚报、另外四枚 Raw 不报 —— 归因没做完就没发货，整笔还原（还原后 oracle 那条是绿的）。⇒ 「解释不了的链接错误」不许靠换写法绕，先归因。

- **形状门这次一定会动**（这是**发码**改动，与 §B142/§B143 那几刀「改 RTL 不动发码」不同）：语料里 `(vb6_VariantToLong(x) + 1)` 那一族换成 `(vb6_VarAddV(x, vb6_VariantFromValue(1)))`。登记照老规矩**只取 CI 自己那份读数**，并且要按 §B145 的口径逐行归因：动了的行只许是「Variant 参与 + - * /」那一族，出现别的变化 = 改道吃到了不该吃的形状。
- **本刀之后剩下的**：**§B146**（内联 `CDec` 算术，本刀 H 族量到）；`\` `Mod` `^` 三条仍走编译期那一档（`\` 背后是 ai/032 rev2 的 `intDivResultType` 四档判据，动它要先过那一格；`Mod`/`^` 的 Variant 形**还没读数**，开工先量）；静态类型 `As Date` 的变量参与算术仍装箱成 `VT_R8` 而不是 `VT_DATE`（只有**盒里**是 Date 才吃到新的 Date 语义）；`CStr(Null)` 那一问的口径（VB6 是错误 94）没量过、没接。


### B144 我们的 COM 服务端交回一枚 Decimal 时，两架构都是静默错（x64 把它放在 offset 8、oleaut32 从 offset 0 读；x86 直接把堆地址当数值）（2026-10-11 第三十五刀第二格收线之后顺手量出；**已出 = 第三十八刀**（提交 `bd009197` 本体 + `8d8bfb3d` 针面补刀 + `be388a82` 按 CI 读数登记形状门；门 **#532** 两条红 → **#533** 只剩形状那一行 → **#535**（head `be388a82`）十二 job attempt 1 全绿；同一格把 §B143 在册的最后一处也带走了））

- **量法**（不进仓）：把 `tests/cc_dll` 整目录复制到 `.build/b1086_dec/CoDll`，只往 `CImpl.cls` 追加三枚公有成员 —— `Plain() As Long`（对照组）、`Amt() As Variant = CDec(1234567)`、`Half() As Variant = CDec("1.25")`；客户端是现成的**免注册**探针形状（`tests/tools/disp_probe.c` 那套 `LoadLibrary` + `DllGetClassObject` + `IDispatch_Invoke`），扩展成 `.build/b1086_dec/probe_dec.c`：把返回的那枚 VARIANT **整个结构体按字节打出来**，再交给 oleaut32 自己做 `→BSTR/→R8/→I4` 三问（独立证人 —— 我们的槽自洽不代表 OLE 读到的也是它）。两台各编各跑，日志 `.build/b1086_dec/report.txt`。
- **对照组是好的**：`Plain` 交回 `vt=3`、`REPLY_BYTES=0300…92100000…`，三问全答 4242 ⇒ 这条管线本身通、指针宽度也对，坏的只有 Decimal 那一档。
- **x64 读数**（`sizeof(VARIANT)=24`）：
  - `Amt` ⇒ `vt=14`、`REPLY_BYTES=0E00 0000 00000000 | 00000000 00000000 | 87D61200 00000000` —— 十进制 1234567 就在最后 8 字节里，**但它躺在 offset 16**；oleaut32 三问全答 **0**（hr 都是 S_OK）。
  - `Half` ⇒ `REPLY_BYTES=0E00 0000 00000000 | 0E00 0200 00000000 | 7D000000 …` —— 这 16 字节从 offset 8 起正好是一枚合法的 `wReserved=14, scale=2, lo=125`（= 1.25），而 oleaut32 三问答 **131086** = `0x0002000E` —— 那正是**从 offset 0 数起的低 64 位**（`0E 00 00 00 00 00 00 00` + `0E 00 02 00`）。⇒ ** oleaut32 读 `VT_DECIMAL` 的位置是结构体的头（offset 0），不是 offset 8**：x64 也一样，与 §B142 ② 在 x86 上量到的「整枚 DECIMAL 盖住 VARIANT」是同一条规则，只是这里的产物是「数值变成头字节的整数解释」而不是「地址」。
- **x86 读数**（`sizeof(VARIANT)=16`）：`Amt` ⇒ `REPLY_BYTES=0E00 0000 00000000 | B0011901 00000000` —— 后 8 字节是 **`0x011901B0`，一枚堆地址**（我们槽里的 `pdecVal` 被当数值铺进了数据区），oleaut32 三问答 **18416048**；`Half` 同形（`0x01190198` → 18416024）。⇒ 与 §B142 那句「只填指针时 oleaut32 把那枚地址当数值答出来」是同一个形状，只不过这次发生在**出站**那一侧。
- **落点在读码里是现成的**：`src/rtl/core/vb6comserver/vb6comserver_obj.c:283` —— `memcpy(pVarResult, &tmpRet, sizeof(VARIANT))`，注释写着「invokeFunc 的返回槽按 `vb6_VARIANT` 布局，**前 16 字节与 OLE VARIANT 内容一致**」。这句主张现在被两台读数**同时推翻**：对 `VT_I4`/`VT_R8` 那几档确实一致，而 `VT_DECIMAL` 在我们槽里（x64 offset 8 内联 / x86 指针）**与 OLE 要求的「从 offset 0 铺整枚 16 字节」根本不是一回事**。同文件 `:258-267` 的实参方向也是同一处主张的下游（把 16 字节的 OLE VARIANT 整个 memcpy 进 24 字节的槽，然后按槽布局读 —— Decimal 形参交进来时同样错位）。
- **修法形状（等开工时照 §B142 的口径，别新造第二份）**：出站/入站都改问**唯一出口** —— 现在 `vb6_DecimalAsVariant`（`vb6rtl.c:319`，两份布局各一臂，含补标签）干的正是「槽 → 合法 OLE VARIANT」这件事，只是它是 `static`。所以这一格不是「给 memcpy 打个 Decimal 补丁」，而是：**服务端两侧的 VARIANT↔槽 双向换算收成一对出口**（出：`vb6_DecimalAsVariant` 提公开名；入：与 `vb6_VariantFromComResult` 的 x86/x64 两臂合一），谁都不许再直接 `memcpy` 结构体。顺带把 §B143 记在册的那两格（`vb6_ComPackVariant` 的 Decimal/Currency/Date 出站档 + `vb6rtl_com.c:336/467` 两条 x86 导入臂）一起收 —— 它们问的是**同一个问句**（「这枚 DECIMAL 在哪、谁有权拥有它」），分开修就是第三次把同一件事抄两遍。
- **【2026-10-11 第三十六刀之后本格的进度】**「往里读」那一半已经落地：`vb6_VariantPackBorrow`（唯一一张「槽 → 真 VARIANT」的表，Decimal 走提了公开名的 `vb6_DecimalAsVariant`）+ `vb6_VariantFromStackVARIANT`/`vb6_VariantFromComResult` 两条 x86 Decimal 臂改从结构体的头铺 16 字节 + `VT_I8` 档 —— 见 §B145。**本格剩下的只有** `vb6comserver_obj.c:258-267/283` 那两处裸 `memcpy`（服务端自己那一侧），而它现在可以直接问那两枚出口，不必再造第三份。
- **判据形状已备好**：这枚探针本身就是夹具 —— 两台各钉「对照组 4242」+「`Amt` 三问都答 1234567」+「`Half` 三问都答 1.25（BSTR 那条随区域设置，照打不钉）」，并且**负控必须是 BASE 那台**（今天它答 0 / 18416048 / 131086 三种数）。真实夹具要不要走注册表那条 `CreateObject` 路（`tests/cc_dll_client`）另说：免注册这条路已经把两端都吃到了，而注册表那条在共享机器上要负责反注册（`tests/com_activate.ps1` 的 try/finally 就是为它写的）。
- **开工前量到的所有权坑（2026-10-11 第三十七刀收线之后读码，决定下一刀的形状）**：入站那半边**不能**直接改问现成的「往里读」出口 —— `vb6_VariantFromComResult`（`vb6rtl_com.c:302`）的语义是**消耗所有权**（`VT_BSTR` 与 `VT_DISPATCH` 两档把 `pv->bstrVal` / `pv->pdispVal` 置 NULL 拿走指针）。服务端适配器手上那枚 `VARIANT` 是**调用方的**（`IDispatch::Invoke` 的契约：被调方不许改动 `rgvarg`），拿它去做「转移」会把对端的实参抽成空；换成「先 `memcpy` 一份私有副本再交给它」也不行 —— 转移走的是**副本里的裸指针**，也就是我们最后 clear 掉的是**调用方拥有的那只 BSTR**。⇒ 入站必须保持**借用**语义（今天的 `memcpy(slot, src, 16)` 在标量那几档本来就是借用，坏的只有 Decimal 那一档）。
- **因此这一格的形状 = 把「一枚 DECIMAL 住在 VARIANT 的哪一字节」再收一枚出口**（与 §B142 那枚 `vb6_DecimalAsVariant` 成对，两份布局各一臂）：`DECIMAL* vb6_DecimalPtrFromVariant(const VARIANT *v)` —— x64 交 `&v->decVal`（内联，就是结构体头），x86 交 `(DECIMAL*)v`（整枚 VARIANT 被 DECIMAL 盖住）。它同时替换掉 `vb6_VariantFromComResult` 与 `vb6_VariantFromStackVARIANT` 里那两处「从结构体的头铺 16 字节」的内联写法（第三十六刀刚修的那两臂），服务端入站那一臂则 `slot.pdecVal = vb6_DecimalPtrFromVariant(src)`（借用，不分配 ⇒ 也就不需要事后 clear）；出站仍走 `vb6_DecimalAsVariant`。**其余各档分毫不动**（今天它们按 16 字节 memcpy 恰好对齐，改它们不在本格）。
- **夹具落点已确定**（不用另起炉灶）：`tests/cc_dll/CImpl.cls` 今天只有一枚 `Twice`，而 `tests/tools/disp_probe.c` 只认 `VT_I4`（其余打 `resultvt=` ⇒ Decimal 今天**看得到标签看不出数值**，这正是没被抓住的原因）。加三枚公有成员（`Amt = CDec(1234567)`、`Quarter = CDec("1.25")`、`EchoVal(ByVal v As Variant) = v` 覆盖入站）+ 探针补两枚**与区域设置无关**的读数（`VariantChangeType` 到 `VT_CY` 打 `*(int64_t*)&cyVal` ⇒ 1234567→123456700、1.25→125；到 `VT_I8` 打整数），针面 `Test-DispatchInvoke` 两台各钉一遍；负控 = BASE 那台（x64 答 0、x86 答那枚堆地址，见上面的读数）。
- **第三十八刀落地（2026-10-11，本格的形状就是上面那两条「所有权坑 + 因此这一格的形状」照办的）**：
  - 新出口三枚 + 服务端两处裸 `memcpy` 一起换掉：`vb6_DecimalPtrFromVariant`（`vb6rtl.c`，唯一回答「手里这枚 OLE VARIANT 的那 16 字节 DECIMAL 住在哪」，两架构都交**结构体的头**，分歧只在 SDK 的称呼）；`vb6_VariantBorrowIntoSlot`（入站，**借用** —— x86 把 `pdecVal` 指向对端内联的那 16 字节，存活期 = 这一趟 `Invoke`，不分配也就不 clear）；`vb6_VariantPackCopyInto`（出站，表仍问 §B145 那份唯一的 `vb6_VariantPackBorrow`，这里只补「带权」那一步：BSTR 新分配、DISPATCH 走 `vb6_ComAddRefDispatch`）。另外 `vb6_VariantFromComResult` / `vb6_VariantFromStackVARIANT` 里那两处「从结构体的头铺 16 字节」的内联写法也改问第一枚 —— 于是「数偏移」这件事在仓里只剩一处。
  - **读数**（新探针给返回值做 `VariantChangeType`，两台**逐行相同**）：`AMT` 交 `dec_cy=12345670000 dec_i8=1234567`、`QUARTER` 交 `12500 / 1`、`ECHO`（入站 5 走一圈）交 `50000 / 5`。
  - **负控 = BASE 那台跑同一枚新探针**（把 `vb6comserver_obj.c` 那一格单独还原后重建一发 exe，验完按 md5 还原回来）：`AMT` 交 `0 / 0`（x64）与 `196693200000 / 19669320`（x86 那枚堆地址）、`QUARTER` 交 `1310860000 / 131086`（x64）与 `196695840000 / 19669584`（x86）⇒ 四条针在 BASE 上全红；而 `ECHO` 在 BASE 上**也 HIT** —— 旧的两处 memcpy 一进一出方向对称、把同一批错位字节又搬了回来，所以它是「改成借用之后没把 Decimal 读丢」的反面证人，不是能红的判据（上面「判据形状已备好」那条预期的就是这件事，现在有了实测的形状）。
  - **订正**：本格先前预测 CY 那一档是 ×100（`1234567→123456700`），实测 oleaut32 按 **×10000** 标度（`12345670000`；`1.25→12500`）。针面按**实测**钉，不按预测钉。
  - 两个坑：① **RTL 是按 C 编的**（`/std:c11`），`out->pdispVal->AddRef()` 直接 C2039「不是 IDispatch 的成员」；而这枚编译错误顺带挡住了一条真雷 —— `VT_DISPATCH` 标签下挂着的可能是一枚**裸子控件 HWND**（宿主分派层的对象表示），裸 AddRef 会把 HWND 首字段当 vtable 解引用，受守卫的 `vb6_ComAddRefDispatch` 内含 `vb6_ComIsDispatchable`。② `src/driver/c3rtl.rc` 给 rc.exe 的 touch 注释**必须整块 `/* */`** —— 以 `#` 开头的行被当预处理器指令，构建当场 `fatal C1083/C1021`（本刀第一次重建就栽在这一句）。
  - **口径记一笔（没顺手补）**：出站现在按那张唯一的表答，数组那一档（槽里 `vb6_vtArray|…` + `vb6_SafeArray1D*`，带 `0x5A1D` 魔数、**不是**真 SAFEARRAY）交回 `VT_EMPTY`。改前那两处裸 memcpy 是把自家描述符指针贴个 `VT_ARRAY` 标签递出去 = 对端按 OLE 布局读同一批字节 = 悬垂；两种都不对，而今天没有任何一条路真在跨服务端边界返回数组（语料 0 处）⇒ 补它要另造一份 SAFEARRAY 转换，属独立一格，不在本格。
  - **夹具与针面**：`tests/cc_dll/CImpl.cls` 加 `Amt` / `Quarter` / `EchoVal` 三枚公有成员 ⇒ `dll_entry.c` 的 `methodCount` 1→4（`cc_dll_identity_single_source` 那条针按读数搬，不是产物走样）；`tests/tools/disp_probe.c` 学会读 Decimal（CY + I8 两档，转换失败打 FAIL 而不静默）；`ax_dll_dispatch_invoke` 的 absent 名单加了 `CALL=AMT` —— 新成员泄进 legacy 那张脸就该红。
  - **门的红因（#532 两条、#533 剩一条；都不是产品缺陷，但 vbp 那一条是本刀自己带进去的）**（门号先按 API 复核再写：`bd009197`→#532、`8d8bfb3d`→#533、`be388a82`→#535，原来这里记的 #529 在仓库的作业列表里查不到对应跑步，按 API 换成真号）：
    - `Emit manifest (shape oracle)` = 存量 **1 份**发码真的变了（`tests/cc_dll/CoDll.vbp`，就是夹具源变了那一行）⇒ #532、#533 连红两趟都是它，按 CI 自己的读数登记在 `be388a82`（1 行，逐行归因见下面那条登记笔），#535 因此转绿。
    - `Tests (vbp #3)`（只在 #532 红，#533 起转绿）= `[DISPATCH] cc_dll_dispatch_iface_only ... FAIL`，缺的正是 `EXTRA_IID / QI_EXTRA / CREATE_IFACE / VTBL_GET_*` 那一整段 —— **根因在针面那一格**：往 Needles 数组里插三条针时，把收尾的 `") @(` 一起吃掉了。PowerShell 里这一笔**语法仍然对**（`[Parser]::ParseFile` 报 0 错，本刀推之前就是这么自证的），变的是**绑定**：`"CALL=ADD"` 从 Absent 挪进了 Needles，`ExtraIid` 往后挪一格 ⇒ 探针只拿到两个参数，`if (argc >= 4)` 那一段整体没跑。x86 那一格侥幸没红，只因为它的数组收尾本来就在下一行（`"DONE") @(`）。
    - **自救面（第 40 道哨兵 `scripts/check_test_helper_integrity.ps1` 长出 H5）**：用 AST 钉住 `Test-DispatchInvoke` 三个调用点的绑定形状（位置实参个数 / 数组实参个数 / Needles 与 Absent 各几条针 = 6/6/1、7/10/1、8/10/1）。负控 = 把那一处收尾再吃掉一次 ⇒ H5 三条同时报红；按 md5 还原后 `235e7d0c` 复绿。这一族与 #78（同名遮蔽）、#166（两个权威）同形：**红不在语法，在谁吃到哪个槽位**。
    - 顺带一条**取数口**（此前记在 memory 里是「job-logs 本机拿不到」）：`gh run view <run> --log` 把整份作业日志拉到本机（本刀就是靠它 40 秒定位到那三条 missing），`gh run download <run> -n test-logs-vbp-3` 还能按分片取夹具的 .out/.err。
  - **本地已验**：cc_dll 与 test_activex_dll 两架构真编译真跑；上面那组 BASE 负控；`run_tests.ps1` PSParser 0 错；`check_rtl_embedded`（125/125 逐字节，证明新 RTL 真嵌进了 exe）、`check_variant_i8_coverage`（13 张表 / DEC must 8 / 0 违规）、`check_rtl_resource_ids`、`check_rtl_naked_names`、`check_rtl_proto_arity`、`check_manifest_coverage`（inputs=400 / registered=400）全绿。形状门那一行（`tests/cc_dll/CoDll.vbp`）**必然移动**（夹具源变了），按口径等 CI 自己的读数再登记。
- **暴露面**：全语料 `CDec` 只有 21 处 / 14 份文件（见 §B142 那条），且都在**同一进程内**用，从不 crossing COM ⇒ 这一格今天没症状，是潜伏项；严重度定在「静默答错数」（与 §B142 同级），不是崩溃。


### B143 Decimal 拿去做算术 / 比较时静默按 0 算（`+` `*` `=` 三条错而 `/` 对；2026-10-11 第三十四刀收线那一趟顺手量出；**前两条已出** = 第三十五刀第一格(标签/换算出口) + 第二格(四条数值口 + 覆盖面哨兵换成「每档 × 每表」，门 #515/516)；第三格 = COM 出站那一档与 x86 两条导入臂 **已由第三十六刀 §B145 出**（三档补全 + 两臂改从头铺），服务端那一处 `memcpy` 已由第三十八刀并入 §B144 出 ⇒ **本账四格全出**）
- **读数**（一次性探针 `.build/b1050_dec/probe.bas`，x64 与 x86 **逐行相同**）：`v = CDec(5)`、`w = CDec("1.25")` ⇒ `v + 1` 交 **1**、`w * 2` 交 **0**、`(v = 5)` 交 **False**、`(v > 6)` 交 **False**，而 `v / 2` 交 **2.5**（对）、`CLng`/`CInt`/`CDbl`/`CSng`/`CCur`/`CDate` 六条全对、`CBool(CDec(5))` 也交 True。⇒ **同一枚值按运算符分家**。
- **发码给的落点**（`--emit-c` 同一次跑的产物）：`P-add` 发成 `(int32_t)(vb6_VariantToLong(v) + 1)`、`P-mul` 发成 `(int32_t)(vb6_VariantToLong(w) * 2)`、比较发成 `vb6_VarCmpLongEq(&v, 5)` / `...Gt(&v, 6)`，而 `/` 发成 `vb6_Num_Div((double)vb6_VariantToDouble(v), 2)`。第三十四刀补的正是 `ToDouble` 那一档 ⇒ `/` 当场对、`+ * =` 还错在 **`vb6_VariantToLong` 没有 Decimal 档**（`default: 0`，于是 5 变 0、0+1=1）。比较那一族走的是 §B138 那批 `vb6_VarCmpLong*` 数值口，同一个漏答。
- **可达性要先量再补**（本格的开工前置）：`CBool` 交对说明 **`vb6_VariantToBool` 不在 CBool 的落点上** —— 名单上有格 ≠ 有人在读。下一刀第一件事是把这 7 处**逐个找证人**（谁把值交到它手里），补完每一处都要有对应的真跑读数；`scripts/check_variant_i8_coverage.ps1` 的口径同时改成「每档 × 每表」（Decimal 已是第二次撞在同一张表上，§B141 第一次）。
- **顺带量到的一条布局推论（未验证，别当已定罪）**：`VBFlexGrid.ctl:9611-9614` 在 `RowID` 那一档写 `RowID = CDec(0)` 之后紧接 `VarDecFromI8 .ID, RowID` —— 把**一枚 Variant 按引用**交给一枚期望 `DECIMAL*` 的 oleaut32 出口。按第三十四刀量到的 x86 布局（VT_DECIMAL 的 16 字节内联盖住整个 VARIANT，而 `vb6_VARIANT` 的 x86 union 只有 8 字节 + 堆上那只 `pdecVal`），这一步要么写进结构体之外、要么写进结构体而没有更新指针，**两种都不对**；但它今天没有症状（FlexGridX86 那格在门上绿），所以先登记成**要证的案**：判据 = 让一枚大 ID 真走到那一支，两台各读回 `RowID` 的值与 `TypeName`。`VarDecFromI8` 在本机是可链接的（RTL 的 DI 桩 `vb6_di_com_stubs.c:209-211` 就在调它）⇒ 与 §B142 那句「`VarDecTo*` 一族够不着」不冲突：够得着的方向是 **From**，够不着的是 **To**。

- **【第三十五刀第一格顺带把「可达性」量完了（探针 `.build/b1062_reach/probe3.bas`，x64 与 x86 逐行相同）】**上面那句「`CBool` 答对 ⇒ `vb6_VariantToBool` 不在落点上」只对了一半 —— 它不在 **CBool** 的落点上，但**在这张表自己的落点上**。逐条读数（`v = CDec(5)`）：`If v Then` 答 **False**（该 True）、`Dim i As Long: i = v` 答 **0**、`Dim lp As LongPtr: lp = v` 答 **0**、`Dim b As Boolean: b = v` 答 **False**、`TakesLong(ByVal x As Long)` 实参答 **0**、`TakesBool` 答 **False**；只有 `s = v`（String 目标）答对（`[5]`，它走第三十四/五刀补的那两个出口）。
- **下一刀（第二格）的题面因此变窄**：要补的是 **赋值 / 形参 / 条件 / 算术 / 比较** 这五条路径共用的那几枚数值口（`vb6_VariantToLong` / `ToLongPtr` / `ToBool`），每补一处就用上面那批旁证人钉住；而 `If v Then`这一条证明「针面不一定要经过内置函数」。另两格（`vb6_ComPackVariant` 的出站档、x86 两条导入臂）仍要自己的证人，不能顺手改。
- **【2026-10-11 第六轮 = 第三十五刀第二格落地：四条数值口 + 覆盖面哨兵换成「每档 × 每表」】**（可达性先量、后补格，这是上一轮定的口径，本刀照做了。）
  - **量法**：发码探针 `.build/b1069_reach/probe4.bas`，七条路一次全量（赋给 Long / LongPtr / Boolean、带 Boolean 判值的 `If`、算术、比较、交数值形参），`--emit-c` 直接读落点。**落点分布是不均匀的**：`vi = vd2`、`vb = vd2`、`vsum = vd2 + 1`、`vn(vd2)` **四条路全折进 `vb6_VariantToLong` 一张表**，`vl = vd2` 折进 `vb6_VariantToLongPtr`，`If vd2` 折进 `vb6_VariantToBool`，而 `(vd2 = 5)` 折进 `vb6_VarToDouble_internal` —— 后者同时是 Eq/Ne/Lt/Gt/Le/Ge **六条比较与它们 Long 形的唯一取数口**，补一格通六条。
  - **BASE 读数**（第三十五刀第一格那台，两台逐行相同）：`I4=0 LP=0 BB=False IF=False ADD=1 EQ=False ARG=0`；修后两台各自 `5 / 5 / True / True / 6 / True / 5`。
  - **政策一处不新造**：整型两档用本表对 double 的 `round()`、布尔档用本表的「非零即 True」、比较档原样交 double；四格取数一律只问 §B142 那枚唯一出口 `vb6_DecimalToR8`，**不留第二条十进制拼装**。判据只喂整数 Decimal —— 平局取整(2.5/3.5)属 §B77① 那个未拍的口径，针面里一条平局答案都没钉。
  - **哨兵那半才是结构性那一半**：`check_variant_i8_coverage.ps1` 的站点正则**以前不认带强转的写法**（`switch ((vb6_vartype)v->vt)`）⇒ 「比较走哪个数值」那一整口从来不在 census 里。**这就是 §B137(VT_I8) 与 §B143(Decimal) 两次漏档都能全绿的原因** —— 不是"名单上有格没人读"，而是**那张表压根不在名单上**。正则认了之后站点 11→13（下限跟着提到 13），名单从一份变两份（VT_I8 must 11 / exempt 2 ＋ VT_DECIMAL must 7 / exempt 6），闭合性(V2/D2)、挂名不挂档(V3/D1)、名字存在(V4)、针面下限(V5)四条**都按档各问一遍**；新增 D3 钉「豁免理由不许短于 12 字」—— 空理由等于在 census 上开一个静默的洞。
  - **盲区也量了并写进哨兵注释**：接收者不叫 `v` 的 vt 表还有 **7 处**（`ax_propbag.c` p->vt / `uc_hostmodel.c` in->vt ×2 / `vb6rtl_com.c` varElem.vt + **pv->vt ×2（正是 §B143 第三格那两条 COM 导入臂）** / `vb6rtl_paramarray.c` retVal.vt）。**没把下限调到 20**，因为那 7 处一条都没归因 —— 收紧正则之前先红在一堆没归因的站点上，比不收紧更坏。
  - **判据**：`$boxNeedles` 37→43（新增 `VB-dec-i4=5` / `lp=5` / `bb=True` / `if=True` / `add=6` / `eq=True`），x64 与 x86 两台各 **43/43 命中、零 miss**（名单从 `run_tests.ps1` 里读出来逐条对，不手数）。
  - **负控两头**：① 行为面 = **剥掉两格档**（`ToLong` 与 `VarToDouble_internal`）重建那台跑**同一份夹具** ⇒ 恰好 4 条 MISS（`i4 / bb / add / eq` —— 正是被剥那两张表的落点，一条不多一条不少），装回后回到 43/43；② 哨兵面 = 四条新判据各用一处假改动证能红（剥两格 ⇒ D1 红两条；从 decMust 摘一枚表名 ⇒ D2 闭合红 + V5 红；把一枚豁免理由改成 `"x"` ⇒ D3 红；把一枚 must 名字改错 ⇒ V4 红），每处跑完按字节还原并核 byte-identical。
  - **护栏**：`check_rtl_embedded` 125/125 逐字节相同；naked_names / proto_arity(1276 对) / resource_ids(125) / static_sentinel(53 道) / manifest_coverage(K1..K3, inputs=400) 全过。
  - **落地**：源码 `f2ab0d38`（RTL 一份 .c ＋ `c3rtl.rc` 来源注脚 ＋ 夹具与针面 ＋ 哨兵）；源码那一跑 = **门 #515**（run 38109731485、head `f2ab0d38`、attempt 1）= 12 条里 11 绿 + 唯一红是这道形状门；登记取自 CI 那台的 `64a06733`（注脚 head=f2ab0d38、`c3-exe sha256=cd79b2308024…`），登记那一跑 = **门 516**。移动的只有 **1 行** = `tests/test_variant_boxing.bas`（bytes 12310→13633；新增那 18 条语句进产物、8 行 ASCII 注释不进），其余 **399 行逐字节未动** = 「只改 RTL 不动发码」第四次由 CI 自己那台复算证实。
  - **剩下的（第三格题面，读数已在上面）**：`vb6_ComPackVariant` 的出站 Decimal/Currency/Date 档 + x86 两条 COM 导入臂（`vb6rtl_com.c:336-338` 在 `vb6_VariantFromComResult`、`:467-469` 在 `vb6_VariantFromStackVARIANT`，都写成 `memcpy(result.pdecVal, pv->pdecVal, 16)` = **拿内联那 16 字节的前 4 字节当指针**）；判据 = 现成的进程内 IDispatch 那套夹具（`tests/cc_dll` + `cc_dll_client`）真让 COM 交回一枚 Decimal，两台各对。

### B142 Decimal 这一档在「往外读」的那张表上（2026-10-11 第三十三刀 census 出来、第三十四刀**第一格已出** = 换算出口与写侧那枚 vt 标签；**「往里读」那一族已由第五十三刀收成一枚核 + census 补认到 19 张（门 **#572** 全绿，读数见上格）** —— 剩下的两条都不在这一问里：宿主存值腿已由第五十四刀收线（§B165），两枚 0 调用者的死码等用户拍）

- **范围（数出来的，不是推的）**：`switch (v.vt)` 那一族共 10 处站点（站名单 = `scripts/check_variant_i8_coverage.ps1` 的 must 名单）里只有 §B141 刚补的 `vb6_TypeName` / `vb6_IsNumeric` 两处答 Decimal，另外 8 处一律落 `default`；再加上 `vb6_Format` 的类型分诊表（`src/rtl/core/vb6rtl/detail/vb6rtl_format_extract.inc:20`，12 档，`VT_DECIMAL` 出现 0 次）= **9 处**。`CLng` / `CDbl` / `CInt` 的 Variant 形（`vb6_CLngV` 等三枚薄壳，`vb6rtl_builtin.h:116-118`）全都转调 `vb6_VariantToDouble` ⇒ 一处漏答同时打掉三条内置函数，实测两架构两台读数全 0。**另一处同表的重罪形状**：`vb6_ComPackVariant`（`vb6rtl_com.c:562-575`）的 `default: pv->vt = VT_EMPTY`，而这张表里不只缺 Decimal —— **Currency 与 Date 也没档**（案例表只到 I2/I4/R4/R8/BSTR/DISPATCH/BOOL/UI1/I8）⇒ 这两档交给 COM 调用时**静默变空值**，与账 #305 里那条「VT_I8 落 default ⇒ 地址变 VT_EMPTY」同一个严重度分级。§B142 开工时这一族要一起数（不是只数 Decimal）。
- **要先拍的一件事（口径，不是手法）**：各表是各自手换（`VarDecToR8` / `VarDecToInt4` / `VarDecFromR8` 那一族一条条写），还是统一交 oleaut32 的 `VariantChangeType` —— 后者把 Decimal 的语义与它自己的舍入、越界政策一起继承过来（VB6 的 `CLng(2.5)` 与 `CLng(3.5)` 都答 2 那一族政策本来就在 oleaut32 手里），前者要自己把这套政策写第二遍。§B141 里那个候选臂走的是 `VariantChangeType`（无格式串那侧），形状可以复用；**注意它的 x86 臂读出来是堆地址形状，这一格要先分清是臂写错还是槽里的 DECIMAL 真没落地**（写侧至今无实测反证）。
- **暴露面**：语料里 `CDec` 4 处，全在 VBFlexGrid（`VBFlexGrid.ctl:9613` 的 `RowID = CDec(0)`、`:7740/7757` 的 `VntArr(...) = CDec(Buffer)`、`Common.bas:613`）；**VB6 真语料**（`D:\vb_yqt4qPac`，5623 份工程）里 `CDec(` 有 **21 处 / 14 份文件**：除 VBFlexGrid 两版之外还有 VBSQLite12 的 SQLiteDataSet/SQLiteCursor/SQLiteConnection、ComCtlsDemo 与几份散件 ⇒ 这一档集中在「数据库 / 单元格取数」这一族场景，不是纸面缺口；⇒ 证人可以钉在那枚真工程上，不用只靠 `.bas` 夹具。结构性那一半：`check_variant_i8_coverage.ps1` 那张闭合名单要不要从「VT_I8 一档 × 每表」扩成「每档 × 每张表」的覆盖面（§B141 已经撞过一次，Decimal 这一撞是第二次）。
- **【2026-10-11 第三十三刀之后又补了三轮临时探针：这一格的口径已经量出来了，不用猜】**（探针是往 `detail/vb6rtl_format_extract.inc` 那张表里插一条带自述的 `VT_DECIMAL` 臂，跑完**已整文件还原**到 HEAD（md5 `5a8335b8b70b`，`git diff` 0 行）并重建了 C3.exe，`check_rtl_embedded` 回到 125/125；日志在 `.build/b1011_probe_run/report.txt`）。**① DECIMAL 本体是好的**：两架构读出的字段逐条对 —— `CDec(5)` 给 `scale=0 sign=0x00 hi=0 lo=5 mid=0`，`CDec("1.25")` 给 `scale=2 lo=125`；x86 那台 `pdecVal` 也不是野指针（`00C89C78` 那类堆地址，提根出来就是上面那五个字段）⇒ 「写侧坏了」这条**正式结案**。**② `VariantChangeType(VT_DECIMAL → VT_BSTR)` 在 scale=0 这一格交空串**：hr=S_OK、vt=8、**len=0**；另一枚 `CDec("1.25")`（它的 scale=2）则交 len=4 的 "1.25"。原地（src==dst）与异地（src≠dst）**两条都红、形状一致** ⇒ 不是别名的问题。所以 `CStr(Decimal)` / 无格式串的 `Format` **不能只靠这一条**；整数型 Decimal（scale=0，也就是 VBFlexGrid 那句 `RowID = CDec(0)`）必须自己拼数字（96 位量 + scale 插小数点，`vb6_CStrLongLong` 那一族就是同形的先例）。**③ `VarDecToR8` / `VarDecToString` / `VarDecToI4` 在这套工具链里够不着**：SDK x64 的 `oleaut32.lib` 按字面 grep —— `VarDecFromR8` 2 处、`VariantChangeType` 有，而这三枚 **0 处**；再用 `ctypes.windll.oleaut32` 按名字取也是 False（`VarDecFromR8` True）⇒ 不是 import lib 漏了，是本机 dll 不按名字导出这一族（多半只按序数）。⇒ 上面那个口径问题的一半直接被消除：**「各表手换 `VarDec*`」这一路出局**（⚠ §B145 实测订正：出局的只有 **To** 那一族 —— `VarDecAdd/Div/Mul/Sub/Neg/Round/Cmp/CmpR8` 两架构都能链能跑，本刀最终交的是 Variant 级的 `VarAdd` 一族），能用的 OLE 出口只有 `VariantChangeType`；而它在**数值档**（→R8/→I4/→I8/→BOOL）面对 scale=0 可靠不可靠**还没量**，这一问留给开工第一发。**④ 顺带一条工具事实**：SDK 的 `tagDEC` 低 64 位联合成员叫 **`lo64`**（不是 .NET 那套的 `loVal`）；写成 `loVal` 在 x86 是 C2039「不是 "tagDEC" 的成员」，而 x64 那台先编过了（那一行住在 `#else` 臂里，本机又是先跑 x64）⇒ 又一条「门只测默认架构，布局/ABI 类刀子必须本机 x86 真编译」的实物
- **【2026-10-11 同日第四轮探针：上面那条的 ①② 两句作废，根因是一个字节】**（探针是在同一枚 `DECIMAL(5)` 上做 A/B：一枚照 `vb6_CDec` 整数臂的写法**手填**字段，另一枚交给 `VarDecFromStr(L"5")`，然后把两者的 16 自节逐个打印出来对）。**唯一的差异在第一个字节**：`VarDecFromStr` 的那枚是 `0E 00 00 00 00 00 00 00 05 00 ...`，手填的那枚是 `00 00 00 00 00 00 00 00 05 00 ...` —— 也就是 **`DECIMAL.wReserved` 这两字节其实是 oleaut32 自己的 vt 标签（它写 14 = VT_DECIMAL）**，而 `vb6_CDec` 的手填路径把它留成 0。**一旦把这一枚补成 14，两枚就逐字节相同，而且 `VariantChangeType` 的答案从「S_OK 但给 0 / 给空串」变成正确的 `d=5` 与 `len=1`（x64 实测，日志 `.build/b1023_run/report.txt`）**。⇒ **上面 ② 那句「VariantChangeType 在 scale=0 这格交空串」是错的**：不是 oleaut32 的毛病，是我们交出去的 DECIMAL 根本没被它认；`CDec("1.25")` 之所以对，是因为那一档走的是 `VarDecFromStr`，14 是 OLE 自己写的。⇒ **上面 ① 那句「写侧无罪」也要改写**：写侧确实有罪，但罪名不是「字段填错 / 所有权」，而是**少写了 `wReserved` 这一枚 vt 标签**（它的字段值一直是对的：scale/sign/Hi32/Lo32/Mid32 逐条对得上）。**修法因此变小了一半**：① `vb6_CDec` 的整数档与 `VarDecFromR8` 失败时的手填兜底要补这一枚（x64 是内联匿名 struct 的 `wReserved1`、x86 是 `DECIMAL.wReserved`，两份布局各写一处，形状与现在那几行字段赋值同形）；②读表那 9 处照 census 补，而且**不必再自己拼十进制数字**（①修好之后 **x64 实测** →BSTR 与 →R8 都对，可以统一走它；x86 那台还欠一发，见下面那句）。**本轮剩下的一个未量**：x86 那台这一轮打印出 7.6e6 那一类不成形的数（手填与 `VarDecFromStr` 两枚给同一个错答案）—— 这是**探针自己在 x86 的读法问题**（不是产品的，因为它们的字节打印已逐字相同），下一刀不再用打印定性，直接用行为判据（`CStr`/`CLng`/`CDbl` 三条针两台各自对）。探针全部还原：该 `.inc` 回到 HEAD（bytes 3919），C3.exe 重建后 `check_rtl_embedded` 125/125，两台的产品读数回到第三十三刀落地版（tn=Decimal / isnum=True / CStr 与 Format 两台皆空）。
- **【2026-10-11 第五轮 = 第三十四刀落地：两处根因都在「格式」上，两台行为判据各六针全中】**（按上一轮定的口径不再用打印定性，直接上行为判据：`CStr`/`CLng`/`CDbl` 两台各自对）。**① 写侧那枚标签**：`vb6_CDec` 手填字段的两路补 `vb6_DEC_VT(...) = vb6_vtDecimal`（宏在 `vb6rtl_variant.h:75`，两份布局各一条，调用点不抄第二遍 `#ifdef`）。**② 读侧 x86 的布局重叠**（这才是前一轮那个「堆地址」读数的来历）：x86 的 `VARIANT` 装 `VT_DECIMAL` 时**不走 `pdecVal` 指针** —— 它那 16 字节整个被一枚 `DECIMAL` 盖住（`wReserved`↔`vt`、`scale`/`sign`↔`wReserved1`、`Hi32`↔`wReserved2`+`wReserved3`、`lo64`↔数据区）。所以只往 `out->pdecVal` 里填地址时，oleaut32 按内联那 16 字节去读，读到的「数值」就是塞进数据区的**那枚地址本身**（`0x015E6B98` = 22965144，与当时打出来的数逐位对得上）。同一个重叠也解释了 ①：`wReserved` 压根不是「保留」，它就是 `vt`。
  - **接线**：唯一出口 `vb6_DecimalAsVariant`（两份布局只在这里分岔，交出去的是副本、不借槽里那块 CoTaskMem）+ 两枚换算出口 `vb6_DecimalToBSTR` / `vb6_DecimalToR8`（`vb6rtl.c:311-363`，都走 `VariantChangeType`）。读侧接三处：Format 的类型分诊表补 `VT_DECIMAL` 那一格（`detail/vb6rtl_format_extract.inc:46`；无格式串走精确出口、带格式串才转 double，形状照 `:40` 那条 `VT_I8` 先例）、`vb6_VariantToDouble` 的 Decimal 档（`vb6rtl.c:380`；一漏同时打掀 `CLng`/`CDbl`/`CInt` 三条内置函数，它们全是转调它的薄壳）、`CStr`（它本来就是 `vb6_Format(x, NULL)` ⇒ 串那一侧只有这一个出口，没有第二份实现可留）。
  - **判据**：`$boxNeedles` 从两条钉到六条（`VB-dec-tn` / `VB-dec-isnum` / `VB-dec-cstr=[5]` / `VB-dec2-cstr=[1.25]` / `VB-dec-clng=5` / `VB-dec-cdbl=1.25`），x64/x86 两台各 **34/34 命中、零 miss**（名单直接从 `run_tests.ps1` 里读出来逐条对，不手数）；`VB-dec-fmt=[5.00]` 照打不钉 —— 带小数点的格式串随区域设置变，与上面 `AR-fmt2` 同口径。
  - **本刀之后剩下的（已用一发一次性探针量过可达性，别再照名单补 case）**（`.build/b1050_dec/probe.bas`，两台读数逐行相同）：`CLng`/`CInt`/`CDbl`/`CSng`/`CCur`/`CDate` 六条内置函数**现在全对**（都转调 `vb6_VariantToDouble` = 本刀补的那一档）；`CBool(CDec(5))` **也答对** ⇒ `vb6_VariantToBool` 那一处不在 CBool 的落点上，补它之前要先证明可达（名单上有格 ≠ 有人在读）。**真正坏的是算术与比较**：`v + 1` 交 **1**、`w * 2` 交 **0**、`v = 5` 交 **False**，而 `v / 2` 交 **2.5**（对）。发码给了落点：前两条发成 `vb6_VariantToLong(v)`（那张表没有 Decimal 档 ⇒ `default: 0`，于是 5 变 0、0+1=1），比较发成 `vb6_VarCmpLongEq(&v, 5)`（同一族数值口），而 `/` 走的是已补的 `vb6_VariantToDouble` —— **同一枚值按运算符分家**，这就是「一张表补了、另一张没补」的现场。剩下 7 处 vt 读表 + x86 那两条 COM 导入臂一起收在下一刀。`scripts/check_variant_i8_coverage.ps1` 那张闭合名单今天只闭「VT_I8 一档 × 每表」，Decimal 这一撞是第二次（§B141 是第一次）⇒ 下一刀把口径改成「每档 × 每表」，并且**先量可达再补**。
  - **落地**：源码 `f2a21b77`（RTL 四份 + `c3rtl.rc` 按约定 touch + 夹具与针面）；`check_rtl_embedded` 125/125 逐字节相同（exe sha `cdd3d3a35264`）；RTL 四道哨兵（naked_names / proto_arity / resource_ids / static_sentinel_registration）+ `check_manifest_coverage` 全过。源码那一跑 = 门 #510（run 38101654250、head `f2a21b77`、attempt 1）= 12 条里 11 绿 + 唯一红是这道形状门（行为面零回归）；移动的 **6 行**分两段：1 行 = 本刀这枚夹具自己（另加 2 条语句，源文本进产物 11333→11618，重写的 10 行注释不进产物），另 **5 行 = `2020` 那一族五份 .vbp 不是本刀动的** —— 拿 CI 自己那份按 head 二分复算（`61d14874` = CI@8fc09f8c 对 `expected@8fc09f8c`）证明那五格在本刀之前那一跑（PR #11 合进来那个 head）就已与期望不一致，只是那条线当时没跑到这道门；登记取自 CI 那台 `github/ci/emit-manifest` 的 `31011213`（注脚 head=f2a21b77），登记那一跑 = 门 #511（run 38102478203、head `45d31ebb`）。台账与 `c3rtl.rc` 里那句「罪在写侧/所有权」已就地订正（写侧的罪是少写标签，字段值一直是对的）。
- **【2026-10-11 第三十五刀第一格：标签收成一个写者，两台不对称因此消除]**。纯 C 探针（不进仓、不碰 C3.exe）`.build/b1053_dec/probe.c` 量到 **`VarDecFromR8` 造的那枚 `wReserved` 也是 0**，只有 `VarDecFromStr` 写 14 ⇒ 上一刀那句「它自己造的那些也写 14」只对字符串臂成立，而 `vb6_CDec` 的**浮点臂正走 FromR8**。
  - **现场症状（VB 级探针 `.build/b1055_dec/probe2.bas`，两台并排）**：x64 上 `CStr(CDec(2.5))` 交**空串**、`CDbl`/`CLng` 交 **0**，而同一枚值写成 `CDec("2.5")`（字符串臂）三条全对；x86 那一台**两臂都对** —— 因为读侧 x86 是内联铺（`memcpy(out, &d, 16)` 时重铺了标签），x64 则照抄槽里的字节。⇒ 这是上一刀留下的**台不对称**，不是新引入的；也就是说「只补了手填那两臂」看着像补全了，其实没。
  - **收法按「一处权威」而不是「逐臂补」**：新增唯一出口 `vb6_DecSetTag`（`vb6rtl_variant.h`，两份布局各一条 + x86 判空），`vb6_CDec` **末尾调一次**，两个手填臂里散写的那两行撤掉；读侧 x64 也补这一枚，与 x86 同形。**代价的一条工具钱：把槽侧的宏 `vb6_DEC_VT`（x64 那份是匿名 struct，成员叫 `wReserved1`）套到 SDK 的真 `DECIMAL`（成员叫 `wReserved`）上 ⇒ x64 C2039 而 x86 照编过**，本机两台都编才抓住的（没浪费门的一跑）——与本线老矩「布局/成员名类改动必须 x86 真编真跑」同一条，第二次命中。
  - **判据**：夹具另加 `Dim vd4 As Variant` + `vd4 = CDec(2.5)` + 三条读数，`$boxNeedles` 从六条到 **九条**（新增 `VB-decf-tn=Decimal` / `VB-decf-cstr=[2.5]` / `VB-decf-cdbl=2.5`），本机两台各 **37/37 命中、零 miss**；改前的负控就是那两份探针读数（x64 `Q-float-cstr=[]`，修后 `[2.5]`）。**刻意不钉 `CLng(CDec(2.5))`**：探针实测 oleaut32 自己按**四舍六入向偶数**（2.5→2、3.5→4、-0.75→-1，与 VB6 的 CLng 同形），而我们现在答 3（`CLng` 走的是 double 那一族）⇒ 这一格属 §B77① 那个还没拍的平局口径，不把未定的答案焨进针面。
  - **落地**：源码那笔（`ec6f0c3d`）（RTL 三份 + `c3rtl.rc` 溯源 + 夹具与针面，夹具注释一律 ASCII）；源码那一跑 = **门 #512**（run 38105635521、head `ec6f0c3d`、attempt 1）= 12 条里 11 绿 + 唯一红是这道形状门（移动的只有自己那枚夹具的 1 行，其余 399 行逐字节未动 = 「改 RTL 不动发码」第三次被 CI 复算证实）；登记取自 CI 那台的 `6342c2fc`（注脚 head=ec6f0c3d），登记那一跑 = **门 #513**（run 38106446965、head `b1430c9a`）。本地：`check_rtl_embedded` 125/125（exe sha `dfb44b3af080`）、RTL 四道哨兵 + `check_variant_i8_coverage`（switches 11 / must 10 / violations 0）+ `check_manifest_coverage` 全过。
- **【本刀读数带出的新格（静态推论，未行为复现）：x86 的「往里读」那一族把内联 `lo64` 当成 `pdecVal` 指针读】** `vb6rtl_com.c:336-338` 与 `:467-469` 两条 x86 臂都写成 `memcpy(result.pdecVal, pv->pdecVal, sizeof(DECIMAL))`，注释还留着「x86 的 DECIMAL 在外部由 pdecVal 指向（vb6_VARIANT 现同布局）」这句 —— 按上面 ② 那条布局事实，`VT_DECIMAL` 的 VARIANT 里那 8 个字节是 `lo64` 本体而不是指针，于是这一步是**从一枚数值当地址去读 16 字节**（小数值如 5 就是读地址 0x5 ⇒ 原生 AV，不是读错数这么轻）。修法与 ② 同形：`memcpy(result.pdecVal, pv, sizeof(DECIMAL))`（从 VARIANT 的头铺）。暴露面未量：要 x86 上真让 COM 交回一枚 Decimal 才吃得到；outbound 那侧 `vb6_ComPackVariant` 本来就没 `VT_DECIMAL` 档（落 `default` 变 `VT_EMPTY`，§B141 那条已记）。⇒ 下一刀把「往外读」那 7 处与这一格一起收，而这一格的判据得是**真跑一条 COM 返回 Decimal**（进程内那套现成夹具就够），不能再用打印。
- **【2026-10-13 第五十三刀开工前的量：「往里读」那一族现在有几张表、差在哪几格，都是数出来的】**（屏具 `.build/k53_sites.py` + `.build/k53_fn.py`，读数 `.build/k53_p/`）① **census 是瞎的**：`src/rtl` 里按 vt 分诊的表共 **22 张**，`check_variant_i8_coverage.ps1` 的站点正则只认 13 张 —— 隐身的 9 张是 `ax_propbag.c:126(p->vt)` / `uc_hostmodel.c:267,291(in->vt)` / `vb6rtl_array.c:61(comVt)` / `vb6rtl_com.c:87(varElem.vt)` / `vb6rtl_com.c:309,441(pv->vt)`（正是 §B143 那两条 x86 导入臂）/ `vb6rtl_paramarray.c:277(retVal.vt)` / `detail/vb6rtl_format_extract.inc:20(expr.vt)`**= §B142 第三十四刀刚补过 Decimal 档的那张 Format 表本身不在哨兵眼皮底下**。② **「往里读」有两张逐档相同的双胞胎**：`vb6_VariantFromComResult`（`vb6rtl_com.c:302-435`，134 行 / 23 档）与 `vb6_VariantFromStackVARIANT`（`:436-556`，121 行 / **同一串 23 档**），两枚都由发码按名字发射（`cgen_expr_binary.cpp` 一族 + `cgen_control.cpp:434`）；剥掉注释逐行 diff 之后语义差**只有四处** —— BSTR（偷所有权 vs `SysAllocString` 复制）、DISPATCH（置 NULL 转移 vs `vb6_ComAddRefDispatch` 守卫 AddRef）、SAFEARRAY 的 BSTR 元素（偷指针 vs 逐个复制，两支先后顺序也相反）、以及 **`default:` 的落点**：前者把未列档接进**数组那一块**（`pv->parray` 把标量字段当 SAFEARRAY 解 ⇒ `SafeArrayGetDim(垃圾)`，块尾还 `pv->parray = NULL` 抹掉原值），后者让未列档落到**换 BSTR 那一块**。⇒ 收成一枚带所有权参数的核之前，这一格得先拍「未列档到底交什么」，因为两枚双胞胎现在交的是**两个不同答案**。③ **稀疏的三张**：`vb6_CallByName` 的 `retVal.vt` 表只有 8 档且 `default: result.lVal = retVal.lVal`（Date/Currency/I8/Decimal 进来一律读成垃圾标签 + 半截数值）；`vb6_EnumNext` 10 档；`vb6_Host_ToWinVariant` 10 档、`vb6_Host_FromWinVariant` 9 档（两枚都缺 Date/CY/I8/Decimal ⇒ 落 default）。④ **暴露面这次有实测了**（不再只是语料计数）：`c.Add CDec(5)` 之后 `VarType(c(1))` 与 `For Each` 读回 **0/Empty**，x64 与 x86 两台逐行相同（`.build/k53_p/x64b.out` / `x86b.out`），而同一枚值直接装箱读是 **14/Decimal** ⇒ 罪在 `Host_ToWinVariant` 那一格，与 §B141 当初「写侧无罪」的结论同一形状；x64 另有一条 `VarPtr` 装的 I8 同样读回 0/Empty。⑤ **死码一格**：`vb6_EnumNext` / `vb6_EnumRelease` 全仓（含发码与哨兵）**0 个调用者**，真 For Each 走的是 `vb6_ForEach_Next`（`vb6com_foreach.c:151`）⇒ 本格不作答，撤不撤由用户拍（§B124 后半那两枚 0 调用者的 recordset 出口是先例）。
- **本刀范围按这些读数收窄**：把「往里读」收成一枚带所有权口径的核（双胞胎合一）+ 把 `Host_ToWinVariant` / `Host_FromWinVariant` / `vb6_CallByName` 三张稀疏表改问既有权威（`vb6_VariantPackCopyInto` / `vb6_VariantBorrowIntoSlot`），判据 = ④ 那四条针两台 + census 从 13 张补到 22 张（补认之后 must/exempt 两边都要重新逐张归因，V1 那条站点下限跟着动，别提前抬）。

### B141 `CDec` 打在整数上两架构都是空串，而这一刀的探针顺手量出 `CLng(VarPtr(x))` 在 x64 报 6 **是正解**（2026-10-11 量；**第一格已出一半**（`vb6_TypeName` / `vb6_IsNumeric` 两档 = 第三十三刀 `9a6b7a56`，Format 那臂回退、剩余 9 处读表另立 §B142）；第一格是缺陷、第二格是防误修的读数）

- **`CDec` 那一格**：探针同一支里 `Len(CStr(CDec(vp)))` 在 **x64 与 x86 都读 0** ⇒ 不是 §B138 那一档的漏答（那条已在账 #305 收口时把 `vb6_CDec` **写进 exempt**，理由就落在哨兵里），是 `CDec(整数)` 本身交出来的东西`CStr` 不动。开工前先量清是哪一段：`vb6_CDec` 造的那枚 DECIMAL 的 scale/sign/Lo32 三个字段，还是 `vb6_Format`/`CStr` 对 VT_DECIMAL 的读法（`vb6rtl_misc.c:276-281` 那段注释写着 x86 的 DECIMAL 走指针形式 —— 两架构两份布局，正是 §B99② 那族的地形）。**【2026-10-11 量完，读的是源码不是推的】断点在「往外读」那一侧**：`vb6_CStr` 只是 `vb6_Format` 的薄壳（`vb6rtl_conv.c:233-235`），而 `grep -rn DECIMAL src/rtl`（连 `vb6_vtDecimal` 一起数）的全部落点是 `vb6rtl.c:366/415`（两枚按指针读的取数）+ `vb6rtl_com.c:329/460`（往 COM 打包）+ `vb6rtl_misc.c:272/275/316`（`vb6_CDec` 自己）+ `vb6rtl_variant.h:29`（枚举那一格）—— **`vb6rtl_format.c` 一个都没有** ⇒ Format/CStr 那张表上压根没有 Decimal 这一档，两架构都落 default ⇒ 空串。同一枚值 `TypeName` 也答 "Variant"（账 #305 那张 13 档的表里没有 Decimal —— 见 `scripts/check_variant_i8_coverage.ps1` 的 must 名单）。⇒ 要补的是**往外读的那两档**（`vb6_Format` 的 Decimal + `vb6_TypeName` 的 Decimal），不是 `vb6_CDec` 的字段填充；两份布局那条警告照旧有效（x86 走 `pdecVal` 指针、x64 内联），补的时候两条臂都要问。【2026-10-11 落点量准】上面那句「`vb6rtl_format.c` 一个都没有」要精确一层：`vb6_Format` 的函数体在 `vb6rtl_format.c:231`，而那张按类型分诊的表住在 **`src/rtl/core/vb6rtl/detail/vb6rtl_format_extract.inc:20`**（六个 `.inc` 在函数体内被 `#include`，非独立编译单元 —— 所以只 grep `.c` 会看不见这张表）。它现有 **12 档**（Empty/Null/Integer/Long/Double/Single/**VT_I8**/Byte/Currency/Date/BSTR/Boolean），`default:` 一律交空串 ⇒ Decimal 就落在这一格。**同一张表上已有这一族的先例**：`detail/vb6rtl_format_extract.inc:40-42` 那条 `case (vb6_vartype)VT_I8:` 正是 Task #44 补的，形状是「无格式串 ⇒ 走精确十进制出口（`vb6_CStrLongLong`），带格式串才转 double 进通用数值分支」⇒ Decimal 那一档照这个形状写（精确那侧可用 `VarDecToString` 或 `VarDecToR8`，取哪一条要按 VB6 的 `CStr(Decimal)` 口径先量）。两架构两份布局由既有的 `vb6_VARIANT_DECVAL(...)` 宏吃（`vb6rtl_misc.c` 里 `vb6_CDec` 自己就用它写、x86 那半是 `CoTaskMemAlloc` 的指针形），读的时候同一个宏。⇒ 下一刀的三个落点定了：① 那张 `.inc` 表补 Decimal 档（两条臂）② `vb6_TypeName`（`vb6rtl_conv.c:387` 那张表）补 "Decimal" ③ 判据两头：`CStr(CDec(n))` 非空 + 值逐字对，`TypeName(CDec(n))="Decimal"`；并顺手问一句`check_variant_i8_coverage.ps1` 那张闭合名单要不要从「VT_I8 一档」扩成「每档 × 每张表」的覆盖面（§B141 之后每一档都会再撞一次）。【同一天补两句读数】**写侧是好的**：`vb6rtl_misc.c:271-331` 的 `vb6_CDec` 四档都答（Integer/Long/Byte 手填字段、Single/Double/Currency 走 `VarDecFromR8`、**BSTR 走 `VarDecFromStr`**、Decimal 直通），所以「空串」这个罪确实只在读侧。**真暴露面也量完了**：语料里 `CDec` 有 4 处，全在 VBFlexGrid —— `VBFlexGrid.ctl:9613` 的 `RowID = CDec(0)`（属性把 Decimal 交给调用方）、`:7740/7757` 的 `VntArr(...) = CDec(Buffer)`（ExportToArray 那族把单元格文本打成 Decimal 塞进 Variant 数组）、`Common.bas:613`；⇒ 下一刀的证人可以钉在那枚真工程上（VBFlexGridDemo 本来就在门禁里），不用只靠 `.bas` 夹具。
- **`CLng(VarPtr(x))` 报 6 不是缺陷，别顺手修**：x64 上那枚地址是 12 位十进制、超出 Long 范围，VB6 对越界转换就是错误 6 Overflow ⇒ 探针里这一行**刻意放在最后**（它会把进程终止，前面五条针先打出来），夹具里也刻意**没有**这一行。记一笔是因为它长得太像回归了：`rc=6` + 半截输出，与 §B140 那两枚「被看门狗杀掉」的工件同一形状 —— 分辨法仍是看 `.err` 与最后一条针。
- **【2026-10-11 第三十三刀，账 #305】上面那句「要补的是往外读的那两档」只出了一档半，范围按读数收窄**：`vb6_TypeName` 补 `case vb6_vtDecimal:` ⇒ "Decimal"（`vb6rtl_conv.c:405`），`vb6_IsNumeric` 把 Decimal 折进那一族返回真值的组（`:356`，与 §B138 尾巴那批 `VT_I8` 同一条口径、同一张表）。**Format 那一臂刻意没发**：把它补上再跑，读数仍然不对（x64 对 `CDec(5)` 仍空、对 `CDec("1.25")` 答 "1.25"；x86 那台读出的是 15643776 那一类堆地址形状）⇒ 单补那张表不成立，臂已回退（`detail/vb6rtl_format_extract.inc` 本刀零改动，候选块与两台读数留在 `.build/b990_format_arm_candidate.inc`）。**判据两头今天只钉得住两条**：`$boxNeedles` 加 `VB-dec-tn=Decimal` / `VB-dec-isnum=True`，x64/x86 两台各自 30/30 全中；CI 上补档前那台（head `e51fe008`）同一夹具这两条都缺 ⇒ 负控成立。`CStr`/`Format` 两条读数照打但**不进针面**（落地版两台都打空 —— 那张表里 Decimal 这一格今天还是没有）。**【本刀的新读数：把「罪在写侧/所有权」这个推论收回】**临时探针（`.build/b997_dec/probe.bas`，不进仓）里 `CDec(5)` / `CDec("1.25")` / `CDec(2.5)` 之后 `VarType` 恒答 14、`TypeName` / `IsNumeric` 答对，而 **`CLng` / `CDbl` / `CInt` 在两台（本机新编 + CI 补档前）×两架构（x64/x86）全读 0** —— 发码一看就明白：这三条发成 `vb6_CLngV` / `vb6_CDblV` / `vb6_CIntV`（`vb6rtl_builtin.h:116-118`），三枚都是薄壳，一律转调 `vb6_VariantToDouble`（`vb6rtl.c:309`），而那张 `switch (v.vt)` 没有 Decimal 档 ⇒ `default: return 0.0`。所以那个 0 是**又一张读表没答**，不是 DECIMAL 本体坏了 —— 写侧至今没有任何针对它的实测证据（`vb6_CDec` 四档的形状照旧成立）。按 `scripts/check_variant_i8_coverage.ps1` 那张站名单复算 Decimal 的覆盖面：**10 处 vt 表里只有 2 处答**（本刀这两处），其余 8 处 —— `vb6_VariantToLong`(`vb6rtl.c:246`) / `ToLongPtr`(`:264`) / `ToBool`(`:287`) / `ToDouble`(`:308`)、`vb6_ComPackVariant`(`vb6rtl_com.c:554`)、`memVariantToI4`(`vb6forms_memberobj.c:257`)、`vb6_ho_variantToLong`(`uc_hostmodel.c:183`) / `ToDouble`(`:194`) —— 加上那张 Format 分诊表（`detail/vb6rtl_format_extract.inc`，`VT_DECIMAL` 出现 **0 次**、12 档里根本没这一格）一起 = **§B142** 的范围。源码那一跑：门 #505（run 38094605541、head 9a6b7a56、attempt 1） = 12 条里 11 绿 + 唯一红是这道形状门；红的那 **1 行** = `tests/test_variant_boxing.bas` 自己（本刀给这枚夹具加了 9 行语句，源文本进了产物 bytes 10000→11333；源注释不进产物，这一点在 out.c 里 grep 过两种编码都是 0 命中），其余 **399 行逐字节未动**、缺席 0 / 多出 0、rc 两侧全 0 ⇒ RTL 的这两档确实不动发码。登记取自 CI 自己那份（`github/ci/emit-manifest` 的 `e0758f7d`，注脚 head=9a6b7a56）= 那一笔 `59e72197`，登记那一跑 = 门 #506（run 38095497701、head 59e72197、attempt 1）= **12 条全 completed/success、非绿 0**，其中 `Emit manifest (shape oracle)` 这一跑就是绿的 ⇒ 登记那一行与 CI 那台自己的复算对齐。改 RTL ⇒ `src/driver/c3rtl.rc` 一并 touch，`check_rtl_embedded` 125/125 逐字节相同（exe sha `7633844cda3d`）。

### B140 `ctrldlg_probe` 撞满 `Invoke-TestExe` 的 60s 全局跑步预算 —— 具名的逐例预算已收成一处权威（账 #304 / 第三十一刀，提交 `3f0b1823` + 入账 `b306ce9e`，2026-10-11 量并落地，**门 #495 全绿**；口径仍是 §B133 那条待办写死的两条规矩：不放宽判据、不把全局 60s 抬大）

- **两次现场，同一形状**：门 #484（attempt 1，`[VBP] ctrldlg_probe ... FAIL (run error)`）与门 #491（`Tests (vbp #1)`，同一格）。
  两次都是那片工件里 `DlApp.out` 停在 `DL10`、**`DlApp.err` 不存在**（`Invoke-TestExe` 正常退是两枚都写，超时那一支只在 stderr 非空时才写）；
  而 #484 attempt 2、#490 的同一枚都是 15 行齐（`DL11..DL14` + `CTRLDLG-DONE`）⇒ 判据内容一条没少，缺的是墙钟余量。
- **为什么它天生吃余量**：`DL11` 那一格是**真弹模态框**（`CommonDialog` / `MsgBox` 那一路），本机空载 2.8s、
  8 枚忙等压 4 vCPU 下 13.8 / 31.4 / 34.5s（§B133 量过）⇒ 同一枚用例的墙钟能被负载放大 ×11，而预算是全局那一个 `-RunTimeoutSec = 60`。
- **该办的事（口径不用再拍，§B133 那条待办已经写死）**：给这一枚（以及以后任何「真弹框」的）一条**具名逐例**预算 ——
  今天 `Invoke-TestExe` 只有全局那一个出口，所以缺的是「按用例给一个数」这条路，不是把 60 抬大；
  **也不许去放宽判据**（`DL11..DL14` 四条读数一条不能少 —— 少了就等于把真模态那一格判没了）。
  开工时顺带问一句：这条逐例预算登记在**哪**（runner 里的具名表 = 又一处「每个用例都有的东西」，
  要一起进 census/哨兵，别在夹具里各写各的 —— §B139 那张表就是这么长出来的）。
- **落地（一处权威 × 一张具名表 × 两个消费点）**：`tests/run_tests.ps1` 里 `$RunBudgetSec` 是唯一的表（今天一枚：`'DlApp' = 150`，覆盖 `ctrldlg` / `ctrldlg_probe` 两片 × x64/x86），`Get-RunBudgetSec` 是唯一的出口（表里没有就回落全局那一个数）。两条等待路都问它：串行路 `Invoke-TestExe` 用 `$budgetSec`，并行路 `Invoke-BasSetParallel` 在**进 runspace 之前**逐枚把数盖在 `$it.BudgetMs` 上（runspace 够不到脚本函数 —— 同一条规则在上面的读数逻辑里已经踩过一次）。**全局 `[int]$RunTimeoutSec = 60` 一字未动，`DL11..DL14` 四条读数一条未减**。键 = **产物的名**（串行路 `-Name $baseName` 与并行路的源文件基名同一个口径，两边都来自 §B166 收成那一处的 `Resolve-VbpExeBase`）⇒ 同一枚夹具在两条路上必然问到同一个数。
- **判据 = 配对读数，两头都看得见**（本机、**没有**合成负载 —— 当时另有写者的 C3.exe 在跑，不抢 CPU）：`DlApp.exe` 带 `C3_CDPROBE=1` 空载实测 1.2 / 2.1 / 1.85 / 1.5s；scratch 拷贝里只插一行 `[BUDGET]` 打印，其余不动 ⇒ ① 全局压到 **0s** 而表里 `'DlApp' = 150`：解析成 150，`ctrldlg_probe` **3/3 PASS**；② 同一个全局、只把键改成一个不存在的名字（就是本节最怕的「静默按全局走 = 只接了一半」）：解析成 **1**，**3/3 被看门狗杀掉**，`.out` 停在 `DL10`/`DL13` 且进程 `alive` —— 与门 #484 / #491 那片工件同形。⚠ 顺带钉一条坑：`-RunTimeoutSec 0` **不是**这条路上可用的负控（同一份拷贝两种读数：PASS 与被杀都出现过，纯墙钟余量），要拿负控就把全局压成非零、并让夹具本身跑得比它久（`.build/b916_pairing` / `b914_budget_trace`）。
- **第 52 道哨兵 `scripts/check_run_budget.ps1`（B1..B6，注册名 `run_budget`）**：B1 权威只一处（解析函数恰被提 3 次 = 定义 + 两个消费点、表名恰 3 次 ⇒ 绕过解析函数直接读表就变数）；B2 每个键必须在**表外**也出现过一次（拼错的键不会报错、只会静默按全局走）；B3 全局那一句必须逐字还在（抬全局 = 本账禁止的那种「修红」）；B4 两个等待点的新形状各恰 1 处、旧的「就地乘全局」两种形状必须 0 处；B5 预算落在 (60, 600]（不大于全局 = 静默收紧，超过 600 会把一整片的墙钟吃掉）；B6 表的适用面非空。八发假改动 + baseline 见 `.build/b911_negctl_report.txt`：**六条规则各被一发打红、baseline rc=0、零误报**。护栏：52 道 [STATIC] 全 rc=0（284s），`check_static_sentinel_registration` 同绿（本刀不改 `src/`，形状门该按 400 行全同过去）。
- **门 #495（run 38085195428、head `b306ce9e`、attempt 1）= 12 条全 completed/success、非绿 0**。两条与本刀直接相关的读数：`Emit manifest (shape oracle)` 绿 = **400 行一字未动**，正是「本刀不改 `src/`」那个预期；`Tests (vbp #1)`~`#4` 四条绿 = 承载 `ctrldlg` / `ctrldlg_probe` 的那几片本轮没红 —— **这不新增证据**，那两条红本来就是按负载抖的（门 #484 / #491），真判据仍是本机那一对配对读数（表里有键 3/3 过、键拼错 3/3 被杀）。

### B133 OLE 容器 / Picture 那两条单位换算路没有行为面（账 #301，**Picture 那一格已出 —— 第二十六刀，2026-10-10，门 #483（源码那一跑）+ 门 #484 attempt 2（登记那一跑，12/12）；olecon 与 ax_site_ext 两格仍开着**）

并表 §B129 时给搬掉的三处各找一枚判据，结果是三处都没有**行为**读数可拿 —— 形状钉得住，行为钉不住：

- `vb6forms_olecon.c`：`tests/olecon/OleCon.vbp` 在 `tests/run_tests.ps1:5138-5139` 是 **`Test-VbpBuild`**（两架构都只编不跑）
  ⇒ 它只证"这族 RTL 编得过、链得上"，证不了 `SetExtent`（px→hm 交出去那份尺寸）与 `olcSizeToObject`（hm→px 交给
  `SetWindowPos` 那份尺寸）这两个决定。本仓没有"运行期真被嵌进容器"的夹具 —— 要补就得有一枚在跑起来的 UI 里被当对象装载的
  控件，而且 `ScaleMode` 得非默认才看得出单位走错。
- `vb6forms_picture_prop.c`：那条折算住在**非 Stretch** 那一支（`VB6_Stretch` 为 0 时按图的天然尺寸画）。
  全 `tests/` 只有三枚 `VB.Image` 盒（`.build/b505_picture_face.py` 数出来的）：`resload/ResAlpha.frm` 与
  `VbQRCodegen-master/test/Form1.frm` 都写 `Stretch = -1`（走的是另一支），`pbsub/PbForm.frm` 那枚 `Img1`
  **从来没装过图** ⇒ `GetPropW(hwnd, L"VB6_IPicture")` 那一支压根进不去。**非 Stretch + 真装着图 = 0 覆盖**。
- `axsite/ax_site_ext.c`：DPI 从写死 96 换成真实 DPI ⇒ **本机就是 96 DPI，这条改动在本地永远不会改变读数**；
  它要的是一枚高 DPI 下的 ActiveX 容器读数（CI 那台的显示缩放没量过，别假设）。

⇒ 三条现在都只靠 R9 那五条调用点 + 三条墓碑托着。这格不是缺陷，是**覆盖缺口**。要补的话优先级：
olecon（唯一会把尺寸真交给外部对象的一枚）> 非 Stretch 的 Picture（补法很轻：一枚 `Stretch = 0` + 装图 + 读回画出来的宽）>
`ax_site_ext`（要换显示缩放，本机无手段，先挂着）。

**第二十六刀的落地（把上面优先级里的第二格办掉了）**：`tests/resload/ResAlpha.frm` 长出一枚 `imgN` —— **不写 `Stretch` 行**（= 0 ⇒ 走 `vb6forms_picture_prop.c` 那条按图的天然尺寸画的支路），盒子 96 x 64 px （1440 x 960 缇），装的是 `alpha.png`（48 x 48，左半红 / 右半全透明）。`tChk` 里 GetDC/GetPixel 沿 y=24 扫 x=0..119，问三件事：红像素（`255`）走到哪、非蓝像素走到哪、白底计数。判据 `NS01` 把红段收在 [18,30]（天然宽 ~23），而**盒子边是 95** —— 这一格钉的正是"按图的尺寸"与"按盒子的尺寸"两种决定的分界。

- 两架构真跑读数**同号同形**：x64 与 x86 都是 `NS01-NATURAL=True / NS01-RAW red=23 nonblue=23 white=0`，  原有的 `PICSET=True`、`PAINTED=True` 一并保住（`$resAlphaNeedles` 长出两条新针）。
- **负控（这枚针能红，且红的就是被并表那一处）**：把唯一权威 `vb6_HimetricToPxOnDc` 的除数从 `2540` 改成 `254`  （10 倍单位错，正是并表之后只有那一处能犯的错误），在 HEAD 归档出来的源码副本 `.build/b569neg_k26` 里冷编一台   C3.exe，用它重编**同一份夹具** ⇒ `NS01-NATURAL=False / NS01-RAW red=95 nonblue=95 white=0`。  同一份夹具、同一台机器、只有编译器不同 ⇒ 这枚针确实吃那处换算。
- **一条读法教训（第一版探针把自己骗了）**：先前扫的是"任何非背景色的最后一格"，两台都读回 95 = 盒子边，  看着像"非 Stretch 那一支根本没生效"。真因是画之前还有一步**白底填满整盒**（`Render` 前那句 `FillRect`），  于是"非蓝"数到的是填充而不是图 ⇒ 证人必须认**图自己的颜色**（红），并把白底计数单独打出来当旁证  （`white=0` ⇒ 这一支走的是 RGBA `AlphaBlend`，压根没进白底那条）。**"覆盖了"与"看对了"是两件事**，  判据里带一条 RAW 读数才看得见这种自我欺骗。

**剩下的两格照旧开着**：`vb6forms_olecon.c`（`tests/olecon/OleCon.vbp` 在 `tests/run_tests.ps1:5138-5139` 是 `Test-VbpBuild` ⇒ 只编不跑，要真读数得有一枚被当对象装进跑起来的容器的控件）与 `ax_site_ext.c`（DPI 从写死 96 换成真实 DPI ⇒ 本机就是 96，本地永远不会变读数）。两处现在仍由哨兵 R9 的调用点 + 墓碑托着。
- **门（第二十六刀）**：源码那一跑 = **#483**（run 38051241817、head `238563d9`、attempt 1）= 12 条 check-run 全 completed、11 绿 + 唯一红 `Emit manifest (shape oracle)` —— 夹具文本进了产物而那一行还没按 CI 的读数重登记，形状门按设计说话；登记那一跑 = **#484**（run 38052053772、head `7a5ca1c2`）attempt 1 = 11 绿 + 唯一红 `Tests (vbp #1)` 里的 `[VBP] ctrldlg_probe ... FAIL (run error)`，**attempt 2（同一枚 head、只重跑失败那片）= 12/12 全绿、非绿 0**。`Emit manifest (shape oracle)` 在 #484 的 attempt 1 就是绿的 ⇒ 第二十六刀那一行的新哈希被 CI 那台独立复算证实。
- **那一发红的归因（走的是「GUI / 计数类红」那套三步，不是「看着像抖」）**：① 两份 CI 工件对形 —— #483 的 `DlApp.out` 交满 15 行（`DL1..DL14` + `CTRLDLG-DONE`），#484 那份停在 `DL10` ⇒ 卡住的正是真弹框那一格（`DL11` 之前）；而 483→484 之间 `src/` 逐字相同（那一笔只动清单）。**还有一条更硬的对形（下次这类红先问它，不用读日志）**：#483 那片工件里有 `DlApp.err`、#484 里没有 —— `Invoke-TestExe` 的两条出口不一样：正常跑完是 `.out` 与 `.err` **两枚都写**（`tests/run_tests.ps1:412-413`），而超时那一支只在 stderr 非空时才写 `.err`（`:397-399`，它只把被杀前已经流出的那半截 stdout 落盘）⇒ 「`.out` 只剩前 10 行 + `.err` 不存在」这对着就是**超时现场**，而不是崩溃（崩的那支走到 412 两枚都写）。② 拿 **CI 那台自己编的 `C3.exe`**（工件 `c3-exe`）在本机重编同一份 `tests/ctrldlg/DlApp.vbp`、设 `C3_CDPROBE=1` 真跑：空载 2.8s 交 15 行全对；**8 枚忙等压 4 vCPU**（2 倍超订）连跑三趟 13.8 / 31.4 / 34.5s 仍 15 行全对 ⇒ 这条用例的墙钟能被负载放大 ×11，而预算是全局的 `-RunTimeoutSec = 60`（`tests/run_tests.ps1:23` 那条注释本来就在讲「5s 在并行编译下会误杀」这一族）。③ 只重跑失败那片一次。⇒ 定性 = **判据的墙钟余量**，产品那一头四条读数（DL11 取消报 32755 / DL12 取消不改进数 / DL13 模态循环 ≥30ms / DL14 静默返回）一条没少。待办记在这儿：这枚「会真弹框」的用例若**再**红一次，就给它一条具名的逐例墙钟预算（今天 `Invoke-TestExe` 只吃全局那一个数、没有逐例出口），**别去放宽判据、也别把全局 60s 整个抬大**。
### B129 himetric↔像素/缇 的折算：4 份文件并成 1 份权威（账 #298，**已出 —— 第二十四刀，2026-10-10，门 #480 全绿**）

第十八刀做 `Object.Width` 那一档时顺带数出来的：那时是 `src/rtl` 里含 `2540` 的**非注释行 = 13 行 / 5 份文件**。
第二十二刀并掉 `axsite/ax_load.c`（那一处现在只转调 `vb6_TwipsToHimetric`，哨兵另钉一条"只许转调 1 次"守住这个条件），
**第二十四刀并掉剩下的三处调用点** ⇒ 名单收成 **1 份文件 / 5 行**，全在权威本人 `vb6forms.c` 里：公式档
`vb6_HimetricToPxOnDc(hm, dpi)`、浮点对 `vb6_HimetricToPxF` / `vb6_PxToHimetricF`、缇档一对
`vb6_TwipsToHimetric` / `vb6_HimetricToTwips`（屏 DPI 那两枚 `vb6_HimetricToPxX/Y` 是转调，本身不含 2540 ⇒ 不进名单）。

搬掉的三处（从前各写各的，`ax_site_ext.c` 那处还把 DPI 写死成 96 —— 与 #184 修 `vb6_TwipToX` 之前同形，
当时 VBFlexGridDemo 的网格在 120 DPI 下被缩小 20%）：

- `vb6forms_olecon.c` 7 行 → 0：`olcSizeToObject` 交给 `SetWindowPos` 的宽高改问 `vb6_HimetricToPxX/Y`；
  两枚 `SetExtent` 那四行把 `* 15L * 2540L / 1440L` 换成 `vb6_TwipsToHimetric(vb6_XToTwipX(px))` 那样的三步转调。
- `vb6forms_picture_prop.c` 3 行 → 0：两枚 `MulDiv(hm, dpi, 2540)` 换成 `vb6_HimetricToPxOnDc(hm, dpi)`，
  **刻意继续喂 DC 的 DPI**（不是屏 DPI）—— 图进打印机/虚屏 DC 时要跟着那个 DC 缩，这条口径不属于"屏 DPI 一族"。
- `axsite/ax_site_ext.c` 1 行 → 0：`TransformCoords` 两向 × 两轴改问浮点对，DPI 从 `vb6_DpiX/Y` 拿。

**并表前先量了四条值面**（不是"看着同形就搬"；扫描脚本 = `.build/b500_olecon_rounding.py`）：
`picture_prop` 搬走的是同一条公式 ⇒ 构造上同值；`olecon` 的 px→hm 两处 96 DPI 下**逐值相同**（1440/96=15 整除），
顺带把 `px * 15L * 2540L` 那截 32 位溢出消掉（旧写法从 px 56365 起开始漂）；`olecon` 的 hm→px 一处由"两次截断"
变成"一次就近"，himetric 1..40000 全扫差值只有 **0 或 +1 px**、从没 ≥2；`ax_site_ext` 那条**本机就是 96 DPI ⇒ 只有形状能证**
（未定罪那一格改记 §B133，别当已验）。

**顺带修的是这条哨兵自己的盲区**：R2 数"非注释行"时从前只剥每行的 `//`，**块注释的续行**因此被当代码数。
按跨行剥法（账 #278 第十七刀定的那条）重算第二十二刀那台（`.build/wt_k23neg`）= **12 行 / 4 份**，比当时钉着的 14 行少 2 ——
`vb6forms_olecon.c` 与 `vb6forms_picture_prop.c` 各有一行 `/* … 2540 … */` 的续行被数了进去。
⇒ "名单只许往下减"这条纪律一直带着两份注释的水分，读数偏高而没人察觉。

**第二十四刀的落地与读数**：
- **判据**：R2 名单缩成 1 份/5 行，且那 5 行 + 两枚屏 DPI 包装逐行钉死；新增 **R9** 钉的是**调用点**
  （olecon 2+4+4、picture_prop 2、ax_site_ext 4）外加三条旧形状墓碑（`MulDiv` / `* 15L *` / `const double k` 各 0 处）——
  并表之后"这一族有几处换算"只剩一个答案，剩下的问题只有"有没有人把旧写法偷偷搬回来"。
- **负控** = 第二十二刀那棵树：只报 R2（名单/总数/权威本人）+ R9（五条调用点全 0、三条墓碑非 0）+ R8
  （那台还没第二十三刀），而 R1/R3/R5/R6/R7 照旧绿 ⇒ 新红是特定的。本树 PASS：
  `himetric sites 5 in 1 files … callers olc=2+4+4 pp=2 axs=4`。
- **真跑面 = 一台都不许动**（`vb6forms.c` 是所有夹具都链进去的那一枚，这一刀的护栏因此是"读数逐行相同"而不是"读数变好"）：
  `ctrlstatusbar` / `sbhm` / `sbfont` 三枚**两架构**共六份 run.out 与第二十三刀那台**逐行相同**（44/44、19/19、7/7，
  `.build/b506_k24_ab.py`），`resload` 七条针 + `resalpha` 两条针全中，`tests/olecon` **x64 与 x86 都编得过**。
- **RTL 已 touch `c3rtl.rc` 重编**：`check_rtl_embedded` OK（125 份逐字节相同，exe sha `1876e1efc5ed`）。
- **形状门那一格：一行都没动，所以本刀没有登记动作**：门 #480（run 38041505334、head `35bac8e2`、attempt 1、12 条 check-run 全 completed/success、非绿 0）里 `Emit manifest (shape oracle)` **同绿** ⇒ 398 行逐行相同，这是"纯 RTL + 注释"那一类刀**预期**的形状（发码文本里没有 RTL，清单自然不动），CI 那台因此独立证实了本机那句"这一刀零发码改动"。

### B131 `Panels(<数字>).Index` 把整数交给 `wchar_t*` 槽 ⇒ 启动期 AV（账 #300，**已出 —— 第十九刀，2026-10-10，门 #464 全绿 = run 38014099324、head `597bd6d1`、attempt 1、12 job 全 completed/success、非绿 0、wall 10m52s；`Emit manifest (shape oracle)` 同绿 => 那行改写被 CI 独立复算证实，`Tests (vbp #1..#4)` 四片全绿 => sbhm 的两头判据两架构真跑过）

`.build/b351_probe/P299.frm` 跑到第 6 行崩（bash 报 139；BASE 那台**同样崩在同一行** ⇒ 与第十八刀无关，是存量）。
产物形状把两件事叠在一处：

- 成员名表把 `Index` 这一档答成 **`vb6_StatusBar_GetPanelIndexByKey`**（`cgen_util_com.cpp:328-329` 两个分支同名），
  而那枚出口的签名是 `int32_t (void* hwnd, const wchar_t* key)`；
- 于是 `SB1.Panels(1).Index` 发成 `vb6_StatusBar_GetPanelIndexByKey((void*)vb6_hwnd_SB1, 1)` ——
  **整数 1 进 `wchar_t*` 槽**，RTL 里 `if (!key || !key[0])` 去读地址 0x1 ⇒ AV。
  `Panels("tp").Index`（字符串下标）反而是对的，所以这格在存量判据 SB14 里从没露过面：**症状按"下标写的是数字还是键"分家**。

**修法方向（别按名字补一格）**：这一枚成员本来就有早绑定的答案 —— `vb6forms_memberobj.c:114` 那张面板名单里
`Index` 在册，由 p->index+1 答；崩的原因和第十五刀撤掉的 C29-Data 直译是**同一个形状**：
COM 侧那张 recognizer（`cgen_util_com.cpp` 里按成员名硬拼出口的那一段）抢在 memberobj 之前答了同一枚成员。
所以第一步是问"这一族里还有哪些名字被两边同时答"（census），第二步才决定撤哪一边 ——
#229 那条"数组元素 extender 属性两形同归一处出口"是同一课。

**判据**：崩的那一行要变成一枚真跑夹具（`Panels(1).Index` 与 `Panels("k").Index` **两头**都要钉，
只钉一头正是这格能活着发货的原因）；负控 = 改前那台在同一份夹具上 AV / 无产物。
**已出（第十九刀，2026-10-10，门 #464）**：数字/键两条各答各的 —— 新出口 `vb6_StatusBar_GetPanelIndex(hwnd, index)`（面板存在就交回自己的 1 基下标，不存在交 0，与 `*ByKey` 那枚的"未找到=0"同一口径），键下标仍走 `*ByKey`；发码侧只把 `index` 那一条的两个参数改成不同名字，三处同步（header 声明 / RTL 定义 / 发码点，各恰好 1 处，哨兵 R5-TRIPLE 钉住）。
**读数**（真跑）：夹具 `tests/sbhm` 加 HM06/HM07 两头 —— 新台两架构 `HM06-INDEX-BOTH-FORMS=True`、`HM07-RAW idx9=0`、9 行跑到 `HM-DONE`、rc=0；**负控就是那枚 AV 本身**：BASE 那台跑**同一份**夹具 `rc=0xC0000005`、只打 6 行、停在 HM05（注意：我那个 .bat 里 `if exist (...)` 块内的 `%ERRORLEVEL%` 是解析期展开的，打出来一片假 `RUN_EXIT=0` —— 退出码要用 python 直接起进程才算数，这是 [[c3-build-test-hazards]] 那条 for 循环课的同一形状）。
**哨兵**：`check_statusbar_panel_hm.ps1` 长出 R5 两头（三处 `sbFinishByKey` 的两个出口名不凸相同 + 新出口三处同步），各自用一处假改动证红（多一处调用点 → R5-PAIRS=4；发码点改名 → R5-TRIPLE 报 `cgen_util_com.cpp got 0`），植完按 md5 还原（`c64f4f0bcf53`，字节级相同）。R1 也加两条产物面（by-index 恰好 2、by-key 恰好 1），负控直接用 BASE 树 + 同一份夹具就红。
**可复用的一条**：凡是"两个参数理应不同名字"的函数型式（一个答数字下标、一个答键名），**结构性地**查得出来：数调用点 + 要求两个字面量不相等。这一样的针不靠语义、也不靠记忆，改错就红。
**刻意没接的一格**：越界那一问（`Panels(9).Index`）在 VB6 是错误 9，本刀只交 0 且只钉前缀 —— 与 #209 那族（动态数组越界裸读→AV）的口径是同一个，要动就一起动。

### B130 `CStr(成员对象的数值成员)` 交出空串（账 #299，**已出 —— 第二十刀 + 第二十一刀（门 #474）+ 第二十五刀（门 #481 = 源码那一跑，门 #482 = 登记那一跑，收掉四条发射路里的第四条）**；第十八刀的夹具撞见）

`tests/sbhm` 的 HM05 一行里两种写法同时问同一枚属性，两台读数是**定论级**的：
`bare=100`（`"x=" & Panels(2).Width` 走 memberobj 的真出口 `vb6_StatusBar_GetPanelWidth`）而
`cs=`（`CStr(Panels(1).Width)`）**打空**。产物侧看得到原因的形状：CStr 那一支发成
`vb6_CStr(vb6_VariantFromComResult(vb6_ComGetProp(vb6_ComGetObjectProp(vb6_hwnd_StatusBar1, L"Panels")…`
—— **整条链掉进了 COM 晚绑定兜底**，而晚绑定那一头对状态条面板对象一无所知（RTL 里"认识但什么都不做"），
于是交回 Empty。同族已知样本：账 #143（控件方法两形都落 COM 兜底）、#221（Picture.Line 两跳都空 ⇒ 静默不画）、
#229（数组元素 extender 属性读进 `&` 拼接 = 裸 int 进 BSTR 槽）。
**机制已量到（2026-10-10 探针 `.build/b351_probe/P299.frm`，x64 真跑 + `--emit-c` 两头对看）**：同一枚 `SB1.Panels(1).Width` 四种写法，产物与读数是这样分的 —— `& 裸拼接`、`= 给 Long`、`= 给 Variant` **三形都走早绑定**（`vb6_StatusBar_GetPanelWidth`，读数 1200），只有 **`CStr(…)` 那一形换了路**：`vb6_CStr(vb6_VariantFromValue(vb6_VariantFromComResult(vb6_ComGetProp(` …晚绑定链…`)))` ⇒ 读数空。⇒ 分岔不在"成员是谁"，在 **CStr 的实参那一路把接收者重新按通用 COM 解了一遍**，解出来的对象在 RTL 里没有 Panels/Buttons 这一档 ⇒ "认识但什么都不做" ⇒ Empty。旁证（同一枚探针里另两族的**裸拼接**形）：`TB1.Buttons(1).Width` 与 `TV1.Nodes.Count` 的裸形**也**发成 `vb6_ComGetStringProp(…)` —— 数值成员按字符串取（#88/#229 那一族），只是这几族从没被钉过。
所以第一步应是枚**跨控件探针**（同一枚成员读法 × {直接拼接, CStr, CLng, 赋值给 Variant} × {memberobj 家族：Panels/Columns/Nodes/Buttons} 四形），
而不是给 `Panels` 单补一格 CStr 特判（那正是本线第 N 次给同一事实写第二份答案）。**第十九刀之后补的 census（`.build/b373_two_answer_census.py`，读数留档）**：**StatusBar 的面板成员是"一枚事实两个答案"的完整样本** —— 发码侧那张 recognizer（`cgen_util_com.cpp` 里按 `memLower == "…"` 硬拼出口的 14 条）答 `count/key/index/text/width/minwidth/autosize/style/tooltiptext`，而 RTL 成员表 `kPanelNames`（`vb6forms_memberobj.c:114`）答的是**同一批 8 个名字**（key/index/text/width/minwidth/autosize/style/tooltiptext）。两边都活：直接拼接与赋值那一形走发码侧那张表（实测产物里就是 `vb6_StatusBar_GetPanelWidth`），而 `CStr(…)` 那一形换到通用 COM 晚绑定那条路（空值），memberobj 那一头则在把面板当对象交出去之后才被问到（`Panels("k").Index` 的键形态同理会分岔）。#300 只是这 8 格里**第一个被发现答错**的（数字下标把整数递进 `wchar_t*` 槽）—— 所以下一步不是继续按名字补格子，而是**先定一枚权威**：面板成员的读法只由一处答（第十五刀撤 Data 一族时定的就是 memberobj 那枚真 IDispatch，本族可照同一口径），然后把"哪一格现在由谁答"写成 census 哨兵（重叠 = 红）再动 `CStr` 那一路 —— 它要修的是"实参那一步把接收者重解了一遍"，不是给 Panels 单开一档。

**六形读数（2026-10-10 探针 `.build/b351_probe/P300b.frm`，x64 真跑 + `--emit-c` 两头对看，读数留档 `.build/b376_emit.txt` / `.build/b377out/run.out`）**：同一枚 StatusBar 上六种写法 —— 直接链式 `SB1.Panels(1).Width` 交 **1200**、`Panels.Count` 交 **2**（两形都走早绑定，产物里就是 `vb6_StatusBar_GetPanelWidth` / `vb6_GetPanelsCount`）；把成员**当对象用**的四形**一条读数都交不出**：`Set po = SB1.Panels(1)`、模块级 `Dim mPo As Object` 赋同一链、`For Each pv In SB1.Panels`（**这一形连一行都不打印**：枚举器从假对象那儿拿不到 `IEnumVARIANT` ⇒ 一次循环都没进，产物里 `vb6_ForEach_Init(vb6_ComGetObjectProp(vb6_hwnd_SB1, L"Panels"))` 那一支是空转）。四形的产物头**一模一样**：都从 `vb6_ComGetObjectProp(vb6_hwnd_SB1, L"Panels")` 起，成员读则按上下文分档 —— 拼接/CStr 那三形发 `vb6_ComGetStringProp(po, L"Width")`（读数空），`With` 那形发 `vb6_ComGetIntProp(_vb6_with_1, L"Width")`（读数 0）。
**订正上一轮的归因**：`CStr` 那一路**不是**独立的分岔口。四条对象形里根本没有 `CStr` 参与（`With ... .Width` 赋给 Long 也一样是 0），它们的共同点是**接收者那一头**：`Panels` 被当成"HWND 上的一个 COM 属性"去读。⇒ 症状叫"CStr 打空"是因为 VB 代码里最常拿 CStr 去打印成员对象，病名要改成**"成员集合的对象档压根没立起来"**。
**断点已定位到具体那一格（读码 + 姊妹控件对照）**：memberobj 那一头的 `VB6_MEMCK_PANELS` 三样都齐 —— `memCollCount` 答 Count、`memCollClear` 答 Clear、`memObjKindOfColl` 映射到 `VB6_MEMK_PANEL`，而 `memInvokePanel` 八个成员**全部转调** `vb6_StatusBar_GetPanel*`（⇒ 值层从来只有一份真相，重复的只是"名字→出口"那张映射）。缺的是**对外入口**：`vb6forms_memberobj.c` 尾部只有事件参数用的 `vb6_StatusBar_PanelAt(hwnd, index)`，**没有 `vb6_StatusBar_Panels(hwnd)`**；发码侧那两处"集合被当对象用"的拦子（`cgen_util_com.cpp` 的 nodes / buttons / listimages / object 四条，value 语境与 object 语境各一处）也就没有 panels 这一条 ⇒ 头掉回 `vb6_ComGetObjectProp(HWND, L"Panels")`，正是 C29-8b 当年给 Nodes 写下的那句"HWND 不是 IDispatch，链接过、运行期读数全空"。
**顺带量到两条**（都影响刀形）：① `vb6_ComUnpackBSTR`（`vb6com_pack.c:107`）对非 BSTR 的 VARIANT 会走 `VariantChangeType` ⇒ 真 IDispatch 交回 VT_I4 也读得出 "1200"，所以"数值成员被按字符串取"那一族（#88/#229）**在对象头立起来之后不再是空值的原因**，别顺手去改那一段；② `memColl_Invoke` 的 `Item` 按 Key 那一支的名单里**没有 PANELS**（`else idx = 0`）⇒ 集合对象一旦立起来，`Panels("k")` 会答 Nothing；这是同一条刀要顺手补的第二格，不是新账。
**刀形（按读数收窄，一处权威 = memberobj 那枚真 IDispatch，第十五刀撤 Data 的同一口径）**：RTL 三格 —— 补 `vb6_StatusBar_Panels(hwnd)` 入口、`Item` 按 Key 补 PANELS 一档、`VB6_MEMD_INDEX` 从 inline `p->index` 改成问 `vb6_StatusBar_GetPanelIndex`（#300 那格留下的**第三份答案**，一起撤）；发码侧两处拦子各补 panels 一条，并且 `statusBarNameOfExpr` 必须认新头 `vb6_ComCallObject(vb6_StatusBar_Panels((void*)vb6_hwnd_X), L"Item", …)` —— 不认就会把链式那 8 条与 `Panels.Add/Remove` 的直译一起打回晚绑定，那是**行为改动**而不是收口，负控要能把它照出来。

**第二十刀已出（落的是一格，量出来的剩两格）**：负控与判据都取同一份夹具、两台编译器 ——
BASE（dev 头 `4e9d05ff`，不含本刀）在 `For Each` 那一形 **一次都没进循环**
（`HM08=False`、`fe=0 idx=0 txt=0`），NEW（合完 dev + 本刀）x64 与 x86 两台都是
`HM08=True`、`fe=3 idx=6 txt=5`（下标之和 6、文本长度之和 5 = "A"+"BB"+"SP"），
而 HM01..HM07 两台完全一致 ⇒ 这一格是真红的变绿，也不是顺手把邻居推动。
**没落的三形与它的题面**：产物里 `vb6_ComGetObjectProp(vb6_hwnd_SB1, L"Panels")` 还剩 **8 处**
（按产物行数的账：HM05 的 CStr 那一路 ×1、`CLng(…)` 三行 ×3、`Set po =` 两形 ×2、
模块级 `mPo =` ×1、`With` ×1 = 8），
它们的头是 `cgen_setlet_set_rhs.inc` 那条与 `cgen_with.cpp` 里三处
`comObjectRefFromCallExpr` 直接拼的，**不经过** `resolveComValue` / `resolveComMarkerForPack`
这两个拦子 ⇒ 下一刀不许再补第四个拦子，题面是把"集合本体怎么立"收成一处出口，让四条发射路都问它。
**病名第二次改口（这次是加一格，不是换一格）**：同一枚 `Panels(i).Width` 在
`For Each` 的对象那一形交回 **157**，而 `CLng(Panels(i).Width)`（实参那一路）交回 **0**
⇒ #299 至少有两格：(a) 集合对象没立起来（本刀落了），(b) **实参那一步把接收者重解了一遍**
（现在主要剩这一格，它才是最初"CStr 打空"看到的那张脸）。RAW 行把两个数一起打出来，
就是为了下一次改 (b) 时能当场看见 157 与 0 合上。
**这条刀踩过的坑（已写成契约）**：把 `panels` 拦成 `vb6_StatusBar_Panels(...)` 之后，
"什么算集合对象表达式"这张名单在仓里住**三处**（`cgen_expr_member_generic_access.inc` 两处 +
`cgen_state.inc` 的 `isControlCollectionExpr` 一处），外加 `statusBarNameOfExpr` 的新头 ——
只补拦子、名单漏一处的后果是实测的：`Panels.Count = 3` 发成
`vb6_StatusBar_Panels((void*)vb6_hwnd_SB1).Count` ⇒ cl C2224（void* 上点成员）。
哨兵 `check_statusbar_panel_hm.ps1` 因此长出 **R6** 六格（产物 obj=1 / ForEach=1 /
三处名单 2+1+1 / 拦子两处 / RTL 入口+声明+Item 按 Key 各 1 / inline `p->index` 6 处只许降），
负控不是假针：拿 dev 头那棵树整棵跑一遍，R6 十一条全红而 R1..R5 仍绿。
**发码清单（这一行现在又是旧的，且是**已知**的旧）**：夹具体改了 ⇒ `tests/sbhm/SbHm.vbp`
那一行必须按下一次过门时 CI 那台的读数重新登记。本轮顺带把那行**先前**的旧账更正过一笔
（`4e9d05ff`，原委：他们那批 33 行的登记是在**还没有这枚夹具**的树上算的 —— CI `d206761f`
那份清单压根没有 tests/sbhm 这一行 —— PR 合并时这一行从我们这边继承了第十九刀之前的哈希；
逐行对过 CI 两份清单与 dev 的登记，除这一行外 398 行全部相同）。**本刀带 `[skip ci]` 推，
没有过门**：谁下一次跑 GA，先按 §B73 的口径把这一行登记掉再判形状红。

**第二十一刀已出（落的正是上一轮写下的那道题面：三条发射路改问同一处出口）**：
- **动了什么**：`src/backend/detail/util/cgen_state.inc` 长出唯一出口 `memberCollectionObjectExpr(hostExpr, memberName)` —— 六族成员集合的创建式（Nodes / Buttons / Panels / ListImages / ListItems / ColumnHeaders）从此只在这一处写。宿主槽按各族自己的事实走：ImageList 无窗口 ⇒ `vb6_com_X` 是实例指针；TreeView / Toolbar / StatusBar 是真窗口 ⇒ `vb6_hwnd_X`；ListView 用 `listViewHwndExprOf`（它还认 With 槽里那枚裸 HWND 变量）；跨窗体 TreeView 查 `externalTreeViewAccs_`。出口只交**串**：三条路的返回语义不一样（值语境要改 `lastExpr_` 与 `isComMarker_`、对象语境直接 return、默认成员下标那一路拼完还得再套一层 Item），语义留在调用点。三份抄本一起撤：`resolveComValue` 里四段（buttons / listitems+columnheaders / nodes / listimages）再加 panels 那一段、`resolveComMarkerForPack` 里五段、`cgen_expr_call_com_bind.inc` 那段四族 if/else。行数 +14/-132、+7/-38、+44/0，`cgen_util_com.cpp` 51906 到 46587 字节。
- **读数**：夹具 HM09..HM12 从「只打 RAW」升成判据，NEW 的 x64 与 x86 两台四形全 True；A/B 拿第二十刀那枚 exe 编跑同一份夹具，四形全 False 而且 want=0（`HM08-RAW ... want=` 那一格当时留的就是这个数）。HM01..HM08 两台与第二十刀一字不差 ⇒ 不是把邻居推动的。
- **病名第三次改口（这次是撤掉一格）**：上一轮写的 (b)「实参那一步把接收者重解了一遍」**不是第二个机制**。HM12 钉的正是那张脸（`CStr(Panels(1).Width)` 与 `CLng(Panels(2).Width)`），出口收口之后它直接从空串与 0 变成 57 与 100 —— 与 (a) 同因：默认成员下标那一路的表也缺 panels 这一行。⇒ 本账到此只剩**单位**那一格，它归 §B41 第二格（getter 该交排版后的宽），与 #298/§B129 合批。
- **零行为改动的证据**：撤掉最后一份重复（com_bind 那四族）前后，7 份工程 `--emit-c` 逐字节相同（TbApp / TvfApp / CtrlImageList / C29ImgObj / C29ListView / SbEvent / SbHm，快照 `.build/b422before` 与 `.build/b423after`）；相对 dev 头，产物变的只有状态条那几行 —— 变的正是本刀要修的那张脸。
- **哨兵 R6 跟着改形**（`check_statusbar_panel_hm.ps1`）：从「三处名单 + 两处拦子」改成钉出口 —— 定义恰好 1 处且住在 `cgen_state.inc` / 问出口 2+1（值语境与对象语境、下标 binder）/ 认前缀那三张名单 2+1+1（这一张不改：它是第二十一刀的前提而不是收口对象，漏一处就是实测过的 C2224）/ **创建式字面量的代码行 census**：六族 2+1+2+2+2+2 = 11，且只许出现在 `cgen_state.inc`，别处出现即 `R6-ONE-MAP-SITE`（这一格防的就是「表又被抄走一遍」，#234 拿 DC 抄两遍、#235 画笔色两份存储同一族）/ 产物里 `vb6_ComGetObjectProp(vb6_hwnd_StatusBar1, L"Panels")` = 0（墓碑）/ inline `p->index` 小于等于 6 只许降。PASS 行同步长出 onemap=11。
- **负控**：拿第二十刀**之前**那棵树（`.build/wt_mf @ 4e9d05ff`）整棵跑新哨兵 —— R6 全族红（`R6-ONE-MAP-SITE` 同时点到 `cgen_util_com.cpp` 与 `cgen_expr_call_com_bind.inc` = 那张表从前真住两处；出口定义 0；`R6-THIRD-ANSWER` 读数 7 大于 6），两条 census 防空转的守卫同时红，而 R1..R5 全绿 ⇒ 红得对症，也不是把旧账一起算进这一刀。
- **清单那一行仍然欠**（`tests/sbhm/SbHm.vbp`）：夹具体与产物都改了，要按过门那台的读数登记。第二十刀带的是 `[skip ci]`、没门；本刀是源码级改动 ⇒ **不带 skip ci**，同一趟把第二十刀那半一起过门。门绿之后按 §B73 的口径登记，再判形状红。

- **门 #471（head `3dc0f6a3`、run 38025387585、attempt 1）= 12 job 里 11 绿、唯一红是 `Emit manifest (shape oracle)`**，而那一红是**已知**的：`tests/sbhm/SbHm.vbp` 那一行记的还是第二十刀之前的产物（第二十刀 + 第二十一刀都改了这枚夹具与它的产物）。同一趟 CI 已把这颗头的新清单发到 `ci/emit-manifest`（`26b7b81b`，`ci: emit manifest head=3dc0f6a3`），**登记取的就是这台 CI 的读数**（§B97：本机冷编那台与 CI 那台的 `bytes/sha256` 会因 include 定序漂，所以只认 CI 那一行）—— `bytes` 13142 到 18905，其余 **398 行逐行相同**（`compare_emit_manifest.ps1` 对本机复算：399/399 相同、哈希不同 0、缺席 0、多出 0；`check_manifest_coverage.ps1`：inputs=399 / registered=399，两套集合相等）。登记那一笔**不带 skip ci**：它要让形状门自己再复算一遍，才算这一批过门。

**两趟门一起补上（第二十刀原带 `[skip ci]`、第二十一刀那趟形状门按预期红）**：**门 #474 = run 38026139882、head `c81bb3d5`、attempt 1、12 job 全 completed/success、非绿 0**，其中 `Emit manifest (shape oracle)` 那一跑**绿** ⇒ 登记的那一行被 CI 那台独立复算证实，`Tests (vbp #1..#4)` 四片全绿 ⇒ `tests/sbhm` 的 HM01..HM12 两架构真跑过 —— 第二十刀与第二十一刀第一次同时过门。同批 `#471`（head `3dc0f6a3`）是**产清单的那一趟**：11 绿 + 形状门红，红因就是当时还没登记的那一行，不是回归。**这条账到此收口**：`vb6_ComGetObjectProp(<HWND>, L"Panels")` 在产物里归零、六形两头对上；剩下的只有一格，且它不属于本账 —— `Panels(i).Width` 交回的仍是**请求值**（单位还是像素那一档），那是 §B41 的第二格，与 #298/§B129（折算住 5 份文件）合批动，别在这里补。


**第二十五刀的落地与读数（2026-10-10，门 #481 = run 38045656192、head `12155d3b`、attempt 1 = 11 job 全绿 + 唯一红是形状门那一行；按 CI 读数登记之后门 #482 = run 38046722506、head `ed23f83b`、attempt 1 = 12 job 全 completed/success、非绿 0）= 四条发射路里剩下的那一条：`With` 的头**
- 上一轮那句"题面是把集合本体怎么立收成一处出口，不再补第四个拦子"照办了：这一刀**没有新增拦子**，
  只是让 `cgen_with.cpp` 的头去问第二十一刀那处唯一出口（`memberCollectionObjectExpr`，答不出就一字不变）。
- **病名又改口一次（这次是缩小）**：第二十一刀记的"三形仍读空/读 0"到本刀实测只剩**一形** ——
  `Set 局部 =` 与模块级 `Set mPo =` 早就交回真集合对象（`local=3 mod=3`，走的是账 #258 那处 `emitSetObjectRhs`
  之后落到对象语境的出口），`For Each` 也通（`foreach_direct=3`）。**唯一还在坏的是 `With`**。
- **一形两个症状**（`.build/b351_probe/P308.frm` 的产物 + 两架构真跑）：① 接收者退成宿主 HWND
  （`void* _vb6_with_1 = (void*)vb6_hwnd_SB1` ⇒ 体内 `.Count` 拿 HWND 当 IDispatch 问，交 **0**）；
  ② 更坏的一条 —— 那枚没被消费的标记**漏进体内第一条语句**，实测
  `withN = .Count` 发成 `vb6_ComSetProp(vb6_hwnd_SB1, L"Panels", …)`：**一次属性读变成往控件写属性**
  （不响、不崩、也不改变任何读数 —— 状态条照旧完好，所以它以前从没露头）。
- **读数**（同一份夹具两台并排，x64 与 x86 逐行相同）：BASE（第二十四刀那台 `b516_base_k25.exe`）
  `HM16-RAW with=0 item2=0 item1=0 early=3 foreach=3` ⇒ `HM16-WITH-COLLECTION=False`；
  NEW `with=3 item2=2 item1=1 early=3 foreach=3` ⇒ True。`HM16-BAR-INTACT` 两台都 True（那一写没写坏东西，
  所以判据是"体内那条语句必须是读"，不是"数值变好"）。
- **判据**：`tests/sbhm` 长出 HM16（三形并排钉：块内 `.Count`、`.Item(2).Text` 与 `.Item(1).Text` 的长度，
  另外两条 `early`/`foreach` 是**跨路对照**——它们与 BASE 同数，所以按 #257 那条口径它们只是反面证人，
  不当"这一档存在"的证据；存在性由 `with=3 item2=2 item1=1` 这三条在 BASE 上全 0 的钉）。
  真正独立的证人只有一头：**设计块里那两个 `A`/`BB`**（长度 1/2 由 .frm 决定，不由 RTL 决定）。
  夹具文本进了产物 ⇒ 形状门那**一行**要按 CI 的读数重登记（本线第四次走 `rebless_emit_manifest.ps1`）。
- **护栏**：邻居两份夹具一台没动（`ctrlstatusbar` 44/44、`sbfont` 7/7 与第二十四刀那台逐行相同）；
  **全语料清单本机重算 = 399 行里 398 行逐字节相同、只有 `tests/sbhm` 那一行动**（新旧两台 exe 对同一批输入并排，
  新增 0 / 缺席 0）⇒ 这一刀在语料里的暴露面就是那枚夹具自己写的这几行，别的工程一个字节都没变；
  哨兵 `check_statusbar_panel_hm.ps1` 的 R6 长出**第四条发射路**的三格 ——
  `R6-EXIT-CALLS`（`cgen_with.cpp` 问那处出口恰好 1 次）、`R6-CREATOR-EMIT` 11 → **13**（多出的两处就是两枚 With 头）、
  `R6-WITH-LEAK-TOMBSTONE`（产物里 `vb6_ComSetProp(vb6_hwnd_StatusBar1, L"Panels"` 必须 0 处）。
  **四档负控跑在源码副本 `.build/b523neg_k25` 上**（`-Root` 指过去，不碰工作树）：
  N0 副本基线绿；N1 只把 exe 换成第二十四刀那台 ⇒ 红**恰好两条**（11≠13 + 墓碑 2 处），且**没**碰源码那一格；
  N2 从源码里撤掉 With 那一问 ⇒ 红恰好一条（EXIT-CALLS got 0）；N3 让它问两次 ⇒ 红恰好一条（got 2）；
  N4 还原 ⇒ 绿。⇒ 三面各自能红，且红的是自己那一条。
- **刻意留在本刀之外的一格（另立 §B134 / 账 #302）**：这枚标记在**非集合**那一族里同样没被消费 ——
  实测 `With SB1.Font : fz = .Size : End With` 发成 `vb6_ComSetProp(vb6_hwnd_SB1, L"Font", …)` 而 `fz` 从头到尾没被赋值。
  把 With 的头改成"照 Set 那样无条件消费"是同一格的更宽修法，但那是**全语料形状**的改动（`With <对象>.<成员>`
  在语料里 70 处）⇒ 不能顺手带，先数暴露面再动，写进 §B134。

### B134 `With <对象>.<成员>` 的头从不消费那枚 COM 标记 —— 体内第一条语句会把它当成自己的写目标（账 #302，**开着 —— 第二十五刀量到，只收了六族成员集合那一档**）

第二十五刀把成员集合那一族接上唯一出口之后，剩下的形状实测如下（探针 `.build/b351_probe/P309.frm`，`--emit-c`）：

- `With SB1.Font` → 头 `void* _vb6_with_0 = (void*)vb6_hwnd_SB1  /* With object ref */`（接收者是宿主 HWND，
  拿它当 IDispatch 问任何一档成员都不会有答案）；
- 体内第一条 `fz = .Size` → `vb6_ComSetProp(vb6_hwnd_SB1, L"Font", vb6_ComPackValue(… L"Size" …))` ——
  **一次属性读被发成往控件写属性**，而 `fz` 在产物里只有初值 `fz = (-1)`，从此再没被赋过值。
  两件事同源：`emitExpr(*node.object)` 交出的是"宿主槽 + 一枚挂着的 `comObjExpr_/comMemberName_` 标记"，
  Set 那一路会调 `resolveComValue("Object")` 去消费它（账 #255 写明的口径），With 的头**从来没消费**。

**为什么本刀不顺手改掉**：把它改成无条件消费是**一行**（`if (isComMarker_) resolveComValue("Object");`），
但那会动到语料里 `With <对象>.<成员>` 的 **70 处**（`grep -rc "^\s*With\s+\w+\." tests/**/*.frm,*.bas,*.ctl`）
—— 里面有多少真的停在"未消费的标记"上、多少已经走宿主结构体/UDT/集合那几条正确分支，**一都没数**。
这一族的既往教训（#234「拿 DC 抄两遍」、#260「With 里整枚成员读取被丢成对象本体」）都是"先数覆盖面再动权威"，
所以题面写成两步：

1. **数暴露面**：对全语料跑 `--emit-c`，按两条机器可判的形状筛 —— ① With 头那行等于宿主槽本身
   （`(void*)vb6_hwnd_<控件名>`）且体内第一条语句刷出 `vb6_ComSetProp(… , L"<成员>"` 的；
   ② 只看产物里 `/* With object ref */` 与 `/* COM SetProp */` 同块相邻出现的（后者更宽，误报由前者收）。
   读数出来才有资格说"这一格值不值得动发码"。
2. **若要动，就一次到位**：With 的头与 Set 的头**共用同一个消费动作**（不许再写第三份前缀判据），
   并把 §B130 那处唯一出口保持在唯一 —— 消费是"要不要问"，答案仍然只有那一处。
   判据两头：六族集合那一形（本刀的 HM16）必须照旧 True，非集合那一形要么钉"体内是读不是写"的墓碑、
   要么钉它交回的**是**成员值 —— 别钉"数值没变"，那正是本账要推翻的形状。

**第一步的读数（2026-10-10，`.build/b543_with_hijack.py`，输入清单 = 形状门那 399 份的同一份）**：

- 判据就按上面 ① 落地：源码侧扫 `With <裸名>.<成员>`，只保留"那枚裸名在产物里确实是 `vb6_hwnd_<裸名>`"的头，
  再去产物里数 `vb6_ComSetProp(vb6_hwnd_<裸名>, L"<成员>"` 的处数。**全语料 = 73 处这种头，其中接收者真是控件槽的只有 3 处，
  而被顶掉的体内第一条语句 = 0 处** —— 那 3 处全在 `tests/sbhm` 本线自己那枚夹具里，且第二十五刀之后连那一格也不发写了。
- 另外 **70 处的接收者不是控件槽**（对象变量 / 窗体 / UC 实例），它们走宿主结构体或真 IDispatch 那两条分支，本刀一行都没动
  （两台 exe 对同一批输入并排，那一族逐字节相同）。
- **一条误报的教训（比读数值钱）**：先用的宽判据「ComSetProp 的目标不是 with 临时、而值里含 with 临时」数出 140→138 处，
  逐行归因后**全是合法写法** —— 例如 `vb6_ComSetProp(TempFont, L"Name", vb6_ComPackBSTR(_vb6_with_97->FontName))`
  是 `With Me` 体内把 `.FontName` 交给另一枚对象变量。⇒ **「目标不是 with 临时」不等于「目标就是头上漏出来的标记」**；
  这一族的判据要把**源码那枚头的成员名**对进产物的写入成员名，两头配起来才算数（只看产物分不开）。
- 破坏面比开账时写得更大：探针 `.build/p340`（把头换成 `With StatusBar1.Font`，体内第一条是**与这枚 With 毫无关系**的
  `pcZ = 1`）两台都发成 `vb6_ComSetProp(vb6_hwnd_StatusBar1, L"Font", vb6_ComPackInt(1))` —— **被顶掉的是那条语句自己的左值**，
  于是那条赋值整个消失，还往控件写进一个数。所以这一族要拍的口径不是"读还是写"，是**头的标记必须被消费掉**
  （消费是一个动作，答案仍然只从 §B130 那处唯一出口拿）。

⇒ 结论：**语料暴露 = 0 ⇒ 本账是潜伏项**，与 #244 同一处置，不当编译阻塞排产；真动手时按上面两步走，
  第二步要连"控件的成员对象要不要有真 IDispatch 身份"一起拍板 —— 只接发码那一半会把形状从「读丢成写」换成
  「拿 HWND 当 IDispatch」—— 那一半下面这段本来就写了没量。

**另一头没量的**：接收者就算消费了标记，非集合那一族仍会落到 `vb6_ComGetObjectProp(<HWND>, L"Font")`
= 拿 HWND 当 IDispatch —— 因为**控件本身不是 IDispatch**（#159/#179 写过的那道墙）。所以本账真正的最后一步
不在发码层，在"控件的成员对象要不要有真 IDispatch 身份"那格（与 §B130 的第二十/二十一刀同一族，
只是那一族有 memberobj 可接，`Font` 这一族没有）。**别按名字补**：`Font`/`Picture`/`Icon` 各补一格就是
第 N 份答案。

### B125 账 #278 第十四刀已出 = 两枚集合 Clear 与 CommonDialog 那六枚 Show* 进零实参那张表：「拼法只许来自表」在 Show* 这一族落地（另立一条新判据：哨兵的文件头读不成注释就会绿着空转，2026-10-10，门 #456 attempt 1 全绿（run 37995142510、head `3f972622`、12/12 全 completed/success、非绿 0、created→updated 10m22s；**Emit manifest (shape gate) 那一跑同绿** ⇒ 「398 份逐字节相同」被 CI 那台独立复算证实））

- **动了什么**：`controlZeroArgMethod` 长出八行 —— `(ImageList, clear)` → `vb6_ImageList_ClearImages`、`(StatusBar, clear)` → `vb6_StatusBar_ClearPanels`、`(CommonDialog, show*)` → `vb6_CdShowOpen/Save/Color/Font/Printer/About`，八条都是 `controlExit("vb6_…", 1, outArgc)`（一枚实参 `void* hwnd`）。六个调用点改问它：语句码头三处（`cgen_call.cpp` 的 ImageList / StatusBar / CommonDialog 三块）、COM 绑定码头两处（`cgen_expr_call_com_bind.inc:443/510`）、表达式码头一处（`cgen_expr_call_callee_withm.inc:399`）。
- **这一族特有的那一味**：CommonDialog 的出口名从前不是写死的，是**现拼**的 —— `"vb6_CdShow" + toupper(mCd[4]) + mCd.substr(5)`，两条码头各拼一遍。那不是"名字抄两遍"，是**命名规则住在调用点**（§B115 的口径），后果比抄名字更阴：名单加一枚就默认拼法对它成立。现在表逐条给出六个名字，拼法归零。
- **两枚 Clear 为什么不并成一枚**：接收者表达式不一样 —— ImageList 是宿主对象，槽位是 `vb6_com_<名>`；StatusBar 是窗口，槽位是 `vb6_hwnd_<名>`。表管「叫什么 / 递几枚」，「递给谁」留在码头（第十三刀的 Winsock 同样口径）。
- **判据三面 + R5 收严**：`check_rtl_proto_arity.ps1` 的 R3 行数 `15 → 23`、R5 名单加这八枚、R4 夹具补三枚控件与八条调用，覆盖面 `恰好 14 枚 → 恰好 22 枚`。**R5 的口径这一刀收严了**：从前只比「整条字面量等于名字」，而码头习惯把左括号拼进同一条串（`"vb6_Ws_Close(" + preWs`），于是**漏掉一半站点** —— 现在比 `"名字"` 与 `"名字("` 两种开头（第十三刀因此少报过；N5 那档"退回 HEAD 的老形状"要等这次收严才三面齐红）。
- **顺手立的新判据（本轮撞出来的一课）**：`check_static_sentinel_registration.ps1` 加 R4 —— 每一份 `check_*.ps1` 的字节头必须是 **UTF-8 BOM 恰好一次**、去掉 BOM 后**首行以 `#` 开头**。起因是这一批的补丁脚本把 BOM 写了两遍 ⇒ 首行变成 `"?# …"` ⇒ PS 5.1 报两行 CommandNotFound **却继续跑到底、exit 0** —— 一道坏掉的哨兵在门上是绿的，比红难发现得多（R1/R2/R3 只查"有没有被登记"，看不见"登记了却跑不起来"）。负控：把双 BOM 复现一遍 ⇒ R4 当场红，还原后 md5 逐字节对上。
- **A/B**：BASE = 本机冷编的 `wt_base_k15 @ d2caf7c0`；全语料 **398 份输入逐字节相同** ⇒ 零行为改动、零登记，`emit-manifest.expected.txt` 未动。本地 49 道静态哨兵全绿；6/6 负控按预期红（抹掉 ShowAbout 一行、把 ClearImages 的个数说成 2、在 `cgen_call.cpp` 里再拼一遍 `"vb6_CdShowOpen"`、拆掉夹具那六条 Show*、四份源码加夹具退回 HEAD、骨架原样=绿）。
- **§B72 到今天**：三刀收完 —— 第十二刀（三张表长出「个数」那一格 + 四面判据）、第十三刀（Winsock 八枚）、第十四刀（本节）；Data 那一族经查是死码，另立 §B124，撤干净 = 第十五刀 §B126（口径按 §B124 的推荐 (a) 落地）。`src/backend` 里剩下的手抄出口名从此都要过 R5 那道名单，加档不改名单就会红。

### B124 新账 = C29-Data 那一族的**发码侧直译**全是死码：三条码头认的前缀 `vb6_Data_Self(` 早已没人再发，而今天活的是 memberobj 那枚真 IDispatch（2026-10-10 量；两半都出完 —— cgen 侧死码 = 第十五刀 §B126，RTL 那两枚 0 调用者的导出 = 第十六刀 §B127（改 `src/rtl` 按控件线规矩 touch `c3rtl.rc` 重编，那一条也记在 §B127）)

- **怎么撞上的**：把 §B123 末条列的 Data 那一族（`refresh` + 四枚 `Move*`，5 枚 × 三条码头 = 15 处手抄出口名）照 Winsock 那一刀的样收成一张表，写完跑 R4 —— **EMIT-CENSUS 当场红：夹具里那五条 `DataZ.Recordset.Refresh/Move*` 一枚也没走到那三条码头**，产物是 `vb6_ComCall(vb6_Data_RecordsetObj(vb6_hwnd_DataZ), L"Refresh", NULL, 0)`。⇒ 那张表收的是**没人再读的答案**，不是「同一个事实的两份答案」，方向错 ⇒ 这一刀整个撤回（改动存在 stash `k14-superseded-by-B124`，工作树回到 `d2caf7c0`）。**这条 census 正是第十二刀立的那一头** —— 没有它，这次会留下一张钉死 20 行的表和一段永不执行的分支。
- **读数（两个方向都查了，不是推断）**：
  - 后端**没有任何地方再发出** `vb6_Data_Self(`：`grep -rn "vb6_Data_Self" src/backend` 只有 6 处，全是**读**它 —— `cgen_expr_call_callee_withm.inc:290`、`cgen_expr_call_com_bind.inc:143` 与 `:237`、`cgen_util_com.cpp:127`、`stmt/cgen_call.cpp:298`，加上 `detail/util/cgen_state.inc:594-597` 那枚 `dataSelfHwndExpr`（从前缀里抠 hwnd）。⇒ 五处 `find("vb6_Data_Self(") == 0` 恒假，其下的直译分支（方法 15 处 + 标量那四枚 `BOF/EOF/RecordCount/FieldCount`）一条都进不去。
  - 活的出口是 `vb6_Data_RecordsetObj`（`cgen_util_ctrl.cpp:517` 在 `recordset` 那一档交的名字）→ RTL `vb6forms_data.c:473` 转给 `vb6_MemberObj_NewRs` ⇒ **真 IDispatch**（成员名表 `vb6forms_memberobj.c:136-137` 里 `BOF/EOF/RecordCount/FieldCount/Fields/MoveFirst/MoveLast/MoveNext/MovePrevious/Refresh` 十一枚齐，应答侧 `:630` 用 `vb6_Data_RecordCount(p->owner)`、`:664` 用 `vb6_Data_MoveFirst(p->owner)`）。⇒ **RTL 那批 `vb6_Data_*` 出口本身是活的**，只是调用者从 cgen 变成了 memberobj。
  - `tests/c29data/DataApp.frm`（装在门上真编真跑那枚）的 `Data1.Recordset.RecordCount` / `.CurrentRow` / `.Refresh` / `.MoveFirst` 全走 memberobj 那一路，DT1/DT2 那些断言今天就是靠它过的 ⇒ 撤死码**不会**动到它，但必须拿它当证人。
  - `vb6_Data_Self` 那枚 RTL 导出（`vb6forms_data.c:452` + `vb6forms_prop_ctrl.h:328`）全仓 0 个调用者 ⇒ 一枚死符号。
- **排掉的一条嫌疑（量过才写）**：本以为走 COM 那一路会把数值成员按字符串读（#88 那一族的翻版）。实测 `n = DataZ.Recordset.RecordCount` 发的是 `vb6_ComGetIntProp(recordsetObj, L"RecordCount")`、`If … RecordCount > 3` 也是 IntProp ⇒ 成员类型走的是既有的 COM 型别表，这一格没有洞。探针在 `.build/probe_data.frm`（三形：赋值 / 比较 / 布尔）。
- **读数补全（同日第二轮，九形探针 `.build/probe_rs.frm`：语句 / `Call` / 赋值 / `If` / `With` 成员 / `Fields("name")` / `Fields(1)` / `.BOF` / `.RecordCount`）**：产物里 `vb6_Data_<成员>(` 的直译 **0 处**，九形全走 `vb6_Data_RecordsetObj` + `vb6_Com*Prop` / `vb6_ComCall`（With 那一形把 `_vb6_with_N` 交给 COM）。⇒ 死码清单因此是**闭合**的，不是抽样：
  - 判据（认那个不再被发出的前缀）5 处：`cgen_expr_call_callee_withm.inc:290`、`cgen_expr_call_com_bind.inc:143` 与 `:237`、`cgen_util_com.cpp:127`、`cgen_call.cpp:302`；加抠句柄的助手 `cgen_state.inc:596-597`（`dataSelfHwndExpr`）；
  - 直译发射 3 处（都在上述死判据之下）：`withm:311`、`com_bind:208`、`com_bind:254` —— 都是 `vb6_Data_FieldValueStr(`；
  - 反过来认那枚发射当前缀的识别器 3 处（于是也恒假）：`cgen_util_com.cpp:122`（`Fields(...).Value` 折回值本身）、`cgen_expr_member_generic_access.inc:42` 与 `:118`；
  - 标量四枚分支（`BOF`/`EOF`/`RecordCount`/`FieldCount` 直译）在 `cgen_util_com.cpp:127-133` 同一格死判据里；
  - RTL 导出 `vb6_Data_Self`（`vb6forms_data.c:452` + `vb6forms_prop_ctrl.h:328`）全仓**含 RTL 自己**都 0 个调用者。
- **顺手订正 §B123 末条 ① 的定性**：那 15 处（`refresh` + 四枚 `Move*` × 三条码头）不是"等着收进表的手抄"，而是**死码里的手抄** —— 收进表就是给没人读的答案立新权威（本轮真做过一次，R4 的 census 当场把它抓回来，改动已撤回未提交状态）。同一条末条里的 ② 两处集合 Clear 与 ③ Show* 六枚经探针是**活的**，已由第十四刀收掉（见 §B125）。
- **下一刀的形状**（当时写的四条，跑之前谁也没料到 census 会先把方向否掉；现状：① 与 ③ 已由第十五刀 §B126 出掉 —— ③ 没有落成「一条结构针」而是整道哨兵 `check_data_recordset_shape.ps1`，因为针只钉得住产物那一头，而这一族的两份答案是**源码里那五格判据**；④ 两头都取了（全语料 398 份 + `.frm` 家族同夹具两台并排）；② 由第十六刀 §B127 出掉，选的是「去」）：要动 RTL，与控件线同一批规矩：改 `src/rtl` 必 touch `c3rtl.rc` 再重编 C3.exe，账 #156 那条坑。① 撤五条前缀判据 + 方法侧 15 枚出口名 + 标量侧那四枚 + `dataSelfHwndExpr`；② `vb6_Data_Self` 导出去留（去 ⇒ touch rc）；③ 判据换成一条**结构针**：`Data1.Recordset.<成员>` 的产物必须只有 memberobj 那一形（`vb6_Com*Prop(vb6_Data_RecordsetObj(…))` / `vb6_ComCall(同一枚, L"<成员>"…)`），不许再出现任何 `vb6_Data_<成员>(` 直译 —— 这一条同时把「以后有人再把直译接回来」挡住；④ 证人 = `tests/c29data` 那套装在门上的断言不动，另加一条 A/B：撤完 398 份逐字节相同（死码的撤动本该零差异，若差一行就说明有一条我判成死的其实活着）。
- **口径要先拍的那一格**（拍完才动手，别默认）：发码侧要不要恢复**直译**？两案 ——(a) 只撤死码，口径定为「Recordset 一律走 memberobj 那枚真 IDispatch」（改动小、产物一字不变、RTL 出口的唯一调用者是 memberobj，一处答案）；(b) 恢复直译（#192 当年那条注释的意图：绕开 Invoke 装箱、少一层），代价是**同一个成员两个调用者**（cgen 与 memberobj 各一条），正与本账「一个事实一处答案」相反。⇒ 本线推荐 (a)，(b) 只有在量出 memberobj 那一路有实际代价（每次 Invoke 装箱 + ID 匹配的成本落在热路径上）时才回头。

### B123 账 #278 第十三刀已出 = Winsock 那一族的「有哪些成员 / 各自叫什么 / 收几枚实参」收成一张表：个数第一次变成**接不接这条形的条件**（arity 哨兵的 R3/R4/R5 三面跟着长，2026-10-10，门 #455 attempt 1 全绿（run 37988792329、head `d2caf7c0`、12/12 全 completed/success、非绿 0、wall 10m53s；含 arity 哨兵所在的 Tests (compile) 片，而 **Emit manifest (shape gate) 那一跑也绿** ⇒ 「398 份逐字节相同」被 CI 那台独立复算证实））

- **动了什么**：`cgen_util_ctrl.cpp` 长出 `controlWinsockMethod(memberLower, outArgc)` —— 八枚成员一行一条（`close`/`listen`/`connect` 1、`accept` 2、`bind` 3、`senddata` 2、`getdata`/`peekdata` 4，个数取自 RTL 原型 `vb6forms_prop_ctrl.h:821-828`）。两条码头改问它：表达式码头（`cgen_expr_call_callee_withm.inc`）里那张 `kWsKnown` 名单 + 八条分支的字面量一起撤掉（**表的键就是名单**，认不认与叫什么从此一处分家不了），语句码头（`cgen_call.cpp:356-358`）那三枚手写的也不留。实参形状（出参取址 `&变量`、`(int32_t)` 强制、空串兜底）照旧留在调用点 —— 与前几张表同一套规矩。
- **这一刀的增量不止是"少抄一遍"**：语句码头那一形手里一枚实参都没有，所以它现在拿的是**表里「个数 == 1」的那几档**（`if (argcWs != 1) fnWs.clear();`），其余成员自动不接、落回带实参那条码头。也就是说个数不只是被对账的数，它成了发码侧的判据 —— 将来谁给 `vb6_Ws_Close` 加一枚形参，这一头会自动停止接它，而不是编出一条递少一枚的调用。
- **判据**：`check_rtl_proto_arity.ps1` 的 R3 行数 `7 → 15`（第十二刀 7 + 这一族 8）、R5 名单加那八枚、R4 夹具 `tests/ctrlzero/ZeroForm.frm` 补一枚 `MSWinsockLib.Winsock wsZ` 与八条调用（无括号三形走语句码头、带实参五形走表达式码头），覆盖面 `恰好 6 枚 → 恰好 14 枚`。**7/7 负控按预期红**：抹一行（TABLE-ROWS + EMIT-CENSUS）、表里把 bind 说成 2 枚（TABLE-VS-RTL 与 EMIT-VS-TABLE 两头齐红）、把名字换成 RTL 查无此名（TABLE-VS-RTL + EMIT-CENSUS）、在 `cgen_call.cpp` 里再拼一遍 `"vb6_Ws_SendData"`（NAME-COPIED）、拆掉夹具那五条带实参的调用（EMIT-CENSUS）、四份文件加夹具退回 HEAD（TABLE-ROWS + NAME-COPIED + EMIT-CENSUS 三面红）、骨架原样（绿）。
- **一条读数订正（负控把自己教了一遍）**：本以为「表里的名字换成 RTL 没有的出口」会红在 `EMIT-VS-TABLE`，实测红的是 `EMIT-CENSUS` —— R4 是**按表里的名字去产物里找**，名字一换，旧名字在产物里整个消失，于是覆盖面先掉下来。⇒ 「改名」与「改个数」是两种红，前者只能被 census 抓到，这一格写进判据面（要盯住 census 掉下来，不能只盯住数目不符）。
- **A/B：全语料 398 份输入逐字节相同**（BASE = 本机冷编的 `wt_base_k13 @ 86ff9be1`，同一台工具链，§B73 口径）⇒ 这一刀**零行为改动、零登记**，与第十二刀那次「一行注释变了、看过再登记」不同；`emit-manifest.expected.txt` 未动。本地 50 道静态哨兵全绿。
- **§B72 剩下的（实数已核；顺带订正 §B122 末条 ② ③ 两格的数。这一条里的 ①② 两格已由第十四刀 §B125 收掉，Data 那一格订正见 §B124）**：① Data 那一族比先前记的多 —— `refresh` + 四枚 `Move*` = **5 枚 × 三条码头 = 15 处**（`cgen_expr_call_callee_withm.inc:297-301` / `cgen_expr_call_com_bind.inc:257-261` / `cgen_call.cpp:303-307`），另 `vb6_Data_FieldValueStr` 两处（那两处的实参整形逻辑也是抄的两遍：数字下标折成宽字面量）；② `vb6_ImageList_ClearImages`、`vb6_StatusBar_ClearPanels` 各 1 处（都是"无括号那一形"在语句码头收尾）；③ CommonDialog 那六枚 `Show*` 与这几族**不是同一味** —— 它的出口名是按拼写现拼的（`"vb6_CdShow" + toupper(mCd[4]) + mCd.substr(5)`，`cgen_call.cpp:404-406`），那是**命名规则住在调用点**，归 §B115 那句「拼法只许来自表」管，收法也不同（要一张「成员名 → 拼好的出口名」的表，而不是「成员名 → 名字 + 个数」）。
- **一处被实测推翻的猜想（本线第四次撞在同一课上）**：看见语句码头只写了 close/listen/connect 三枚，我推断「不带括号的 `wsX.Bind` 因此没有出口 ⇒ 落 COM 兜底、编得过而不做事」。跑一发 `--emit-c` 就否掉了：`tests/ctrlwinsock/WsForm.frm:189` 的 `wsB.Bind 0` 发的是 `vb6_Ws_Bind((void*)vb6_hwnd_wsB, (int32_t)0, L"")` —— **带实参那一形走的是另一条形**（表达式码头），两边都通。⇒ 结论：扣住修法不动的理由必须量过，「代码里只写了三枚」不等于「只有三条路」（§B115 那一课在这里再记一遍）。

### B122 账 #278 第十二刀已出 = 控件方法那张表现在同时答「名字 + 递几枚实参」，§B72 欠的那一头由同一道哨兵四面一起对账（`[STATIC] rtl_proto_arity` 扩三条 + 一枚新夹具，2026-10-10，门 #454 attempt 1 全绿（run 37984714811、head `b281c79a`、12/12 全 completed/success、非绿 0、wall 10m53s；含 [STATIC] rtl_proto_arity 所在的 Tests (compile) 片，而 **Emit manifest (shape gate) 那一跑也绿** ⇒ 本机冷编登记的那一行覆写被 CI 那台复算证实）

- **动了什么（三张表各长一格 + 三个调用点把手抄的那份撤掉）**：`controlZeroArgMethod` / `controlOneArgMethod` / `controlScaleMethod` 各多一个 `int* outArgc`，出口名与个数由同一条 `controlExit("vb6_…", N, outArgc)` 交出 —— 表恰好 **7 行**（`vb6_SetControlFocus` 1、`vb6_Slider_ClearSel` 1、`vb6_ClearList` 1、`vb6_ControlTextHeight` 2、`vb6_ControlTextWidth` 2、`vb6_ScaleUnitX` 3、`vb6_ScaleUnitY` 3）。删掉 `cgen_call.cpp` 里 Fix 086 那截把「类型是 ListBox/ComboBox」与 `vb6_ClearList` 又硬编码一遍的分支 —— 语句码头下方 C29-SL-l 那一格早已问同一张表，同一个事实住两处，改名字或改个数必有一份落后。另把两处手写出口名改问表：`cgen_expr_call_callee_withm.inc` 的 `Clear` 一行（注释原样保留），`cgen_expr_ident_builtin.inc` 里窗体/属性页裸写 `ScaleX/ScaleY` 那一格（问的是表里 Form 那一行）。
- **判据（原来两头的哨兵扩成四面）**：`check_rtl_proto_arity.ps1` 本来的判据面是「RTL 的头 == RTL 的体」（账 #240）。本轮加三条：**R3** 表里的出口名必须能在 RTL 头里查到原型且个数相同（**查不到原型也算红** —— 没有声明就没有担保；只有定义没声明同样算红），`cgen_util_type.cpp` 那张运行时参数表若有同名行也必须同数，行数恰好 7；**R4** 跑一次 `--emit-c`（只走前端，不起 cl）数新夹具 `tests/ctrlzero/ZeroForm.frm` 里**实际递出的实参数**与表对，夹具跑出恰好 6 枚；**R5** 那七枚出口名在 `src/backend` 别处再出现成字符串字面量 = 第二份答案 ⇒ 红（两处例外：表自己那份文件，与那张类型表 —— 后者由 R3 的 TABLE-VS-TYPEORACLE 对账）。`-Root` 形参一并补上，负控全在副本里跑。**9/9 负控按预期红**：抹掉一行（TABLE-ROWS）、把 SetFocus 的个数改成 2（TABLE-VS-RTL）、出口名换成 RTL 里没有的（TABLE-VS-RTL 的另一档）、只动类型表那一行删掉一项（TABLE-VS-TYPEORACLE 且**不许**连带动 TABLE-VS-RTL）、在别的 backend 文件里把名字再拼一遍（NAME-COPIED，注释行不算）、把表里的个数改成与产物不符（EMIT-VS-TABLE + TABLE-VS-RTL 两头齐红）、拆掉夹具三条调用（EMIT-CENSUS）、骨架原样（绿）、**把那五份文件退回 HEAD**（TABLE-ROWS 与 NAME-COPIED 同时红 —— 这一刀撤掉的正是那个形状）。
- **实测的代价（形状门红 1 行，看过再登记）**：A/B 的 BASE 用**同一台工具链本机冷编**（`git worktree add --detach .build/wt_base_k12 cc950f30` + cmake/Ninja，实测 6 分钟）—— §B73 那条口径：CI 那枚工件只能用来跑测试，不能当发码基线。全语料 **398 份输入 397 份逐字节相同**；唯一变的这一行在 `tests/Charts 2020/ucTreeMaps/Proyecto1.vbp`：`vb6_ClearList((void*)vb6_hwnd_ListSubFonts);  /* ListBox.Clear */` → `… /* clear */`。出处是 `PropPagFMR.pag:593` 的 `ListSubFonts.Clear` —— 从前由 Fix 086 那截答，现在由 C29-SL-l 答，而这一族的语句码头一律用**小写成员名**做随行标记（同一趟里 `SetFocus` 是 `/* setfocus */`，With 块那形仍是 `/* ListBox.Clear */`，因为它在表达式码头、注释归那格自己拼）。**C 代码一字未变，变的只是注释** ⇒ 判定为接受，用 `scripts/rebless_emit_manifest.ps1 -Manifest <本机清单> -Bless` 登记（覆写 1 / 新增 0 / 退出 0，复算 398/398 绿，`check_manifest_coverage` K1..K3 绿）。⇒ 这一格走的是 §B100 那条自救路径，只是 BASE 用的是本机冷编那台（§B73 的口径），登记前先给了**逐行归因**。
- **剩下没收的同一族（当时的实数；① 已由第十三刀 §B123 收掉，② ③ 两格的数在 §B123 末条订正）**：R5 今天只钉这七枚，另有三族仍是手写出口名 —— ① Winsock 零实参三枚（`vb6_Ws_Close` / `vb6_Ws_Listen` / `vb6_Ws_Connect`）**两条码头各抄一遍**（`cgen_call.cpp:356-358` 与 `cgen_expr_call_callee_withm.inc:455-459`，后者还把左括号拼进同一条字面量 ⇒ 按名字 grep 只数得出 1 处，别据此以为没重复）；② ListBox/ComboBox 的成员链（`AddItem` / `RemoveItem` / `List(idx)`）在 withm 里现拼，`RemoveItem` 的实参还在调用点手里；③ `vb6_Data_MoveFirst` / `vb6_Data_MoveLast` 各 1 处。这三族的「个数」一旦进表，R3/R4/R5 是现成的判据面，**不必再开新哨兵**。
- **§B72 的账怎么算**：那一节要的两步（实参个数写成表里的一个取值 ＋ 同一条针两面都问）对这三张表已做完；它当时要求排在 §B70 的画布表之后 —— §B70 的 ①②③ 都已出（门 #368/#374/#375），前置成立。画布那一族（`src/common/canvas_drawing.hpp`）今天有名字表但**没有「个数」那一格**，也就没进 R3/R5 的面，那是 §B72 剩下的那一半。

### B121 新账 = 同一个未声明名、同一枚 .ctl，两个工程给出**两种发码**：主 Charts 折成 Empty 编得过，独立 UC 交裸名 C2065（2026-10-10 量，**未开工**）

- **读数（同一枚 exe，各出一趟 --emit-c，再各真编一遍）**：`tests/Charts 2020/ucProgressCircular/ucProgressCircular.ctl` 里那句
  `If hBrush = 0 Or Count = 0`（文件有 `Option Explicit`，全工程 grep 无 `Public Count`，RTL 也没有裸名 `Count`）——
  - 吃 `tests/Charts 2020/Proyecto1.vbp`（harness 里 `Test-GuiVbp Charts2020` 真编真跑那枚）⇒
    `vb6_VARIANT _vcmp_0 = vb6_VariantFromValue(vb6_VariantEmpty());` ⇒ **BUILD-RC=0，出得了 Proyecto1.exe（1,195,008 字节）**；
  - 吃 `tests/Charts 2020/ucProgressCircular/Proyecto1.vbp`（harness 里**刻意不列**那枚，见 `run_tests.ps1` 的账 #187 注）⇒
    `vb6_VariantFromValue(Count)` 裸名 ⇒ `ucProgressCircular.c(1503): error C2065: “Count”: 未声明的标识符` ⇒ C3 exit 2。
  两边都挂着同一条 VB3001 ⇒ census 里那两行 `Count` 其实是**这一格的两个脸**。
- **为什么必须单开（它是 §B101 那刀的前置，不是尾巴）**：升级的判据正是「未声明名 + `Option Explicit` ⇒ error」。  这一格说明发码侧对同一个输入**今天就有两个答案** ⇒ 升级之前得先定位那条分岔（哪条分支把裸名折成 Empty、哪条把  裸名直接交出去、分岔的输入条件是什么），否则升级只是给「两个答案」再加一句诊断，而真编那一头照旧一半能过一半不能。
- **与 §B115/§B118/§B120 同族**：一个事实两个答复。只是这一族不在宿主表那一路，在「未声明标识符」的发码那一路。
- **排除面（第十五刀等门期间的合成探针，`.build/k16probe/`，两形都不复现 —— 这条负读数把搜索面收窄了）**：拿两枚**手写**模块做最小复现（一枚兄弟类 `ModA.cls` 交 `Public Property Get Count() As Long`，另一枚接收者裸用 `Count`），各跑「ModA 在工程里 / 不在」两台并排：
  - 接收者是标准模块（`ModB.bas`，`s = CStr(Count)`）：两趟都发 `vb6_VARIANT Count = vb6_VariantEmpty(); /* 隐式变量 */` ⇒ **同形**；
  - 接收者是类模块（`ModB.cls`，`If x = 0 Or Count = 0 Then …`，与 §B121 那句同形）：两趟同样都发隐式 Variant 局部 ⇒ **同形**。
  ⇒ 那条分岔**不是**「兄弟类成员污染裸名」这一件事本身，它还要 Charts 那两个 UC 独有的某样东西（候选：`.ctl` 的宿主/设计器那一层给本模块留下的名字、`Friend`/过程同名、或 driver 给 UC 单独走的那条跨模块登记路 —— 账 #219/#222 记过它）。下一刀别再用合成夹具试这条路，直接拿 Charts 的副本做**减法**：一次撤一样，看哪一样撤掉就翻成裸名。
  （顺带一条独立读数，与本账无关但值得记：同一份 `ModB.bas`，工程里有第二枚模块时发 `vb6_ModB_Probe`，只有它自己时发 `vb6_Probe` —— 过程名的模块前缀按「工程里模块数」决定，这是既有设计，不是本轮要动的。）
- **分岔定位到一格（2026-10-10，拿 Charts 的副本做减法，`.build/k17charts/`）**：把 `UserControl=ucPieChart\ucPieChart.ctl` 那一行从工程副本里撤掉，`ucProgressCircular.ctl:949` 那句 `If hBrush = 0 Or Count = 0 Then Exit Function` 的产物从 `vb6_VARIANT _vcmp_0 = vb6_VariantFromValue(vb6_VariantEmpty());` 翻成 `vb6_VARIANT _vcmp_0 = vb6_VariantFromValue(Count);`（裸名 ⇒ 真编就 C2065）。两趟**都没有** `vb6_VARIANT Count = …` 那条隐式局部声明（全工程数：隐式声明 0 / Empty 内联 1↔0 / 裸名 0↔1）⇒ 差别只落在**一个分支**上：`cgen_expr_ident_builtin.inc:253-256` 的那条 `if (leakedClassMember110u)` ⇒ 交 `vb6_VariantEmpty()`。
  ⇒ 「两个答案」的发源不是两条码头各写一遍，而是**一枚布尔量按「本工程别处有没有同名的公开成员」决定一枚未声明裸名的值** —— 这个输入条件在语言层面根本不该参与。合成探针（上一格那条排除面）走的是**第三条路**：给裸名发一条隐式 Variant 局部声明，两趟同形。⇒ 下一刀要问的是「为什么 `.ctl` 里那枚裸名没进合成夹具走的那条声明路」（两条路的入口条件各是什么、谁该赢），那才是收成一处的落点；**别去调 `leakedClassMember110u` 的取值**，那只是在两个错答案里挑一个。
- **根因（同日第三轮，读码 + 两份副本对照，已闭合）**：`ucProgressCircular.ctl` **写着 `Option Explicit`**（`grep -c` 答 1），而那枚裸名的两条出路分别住在**两个层**：
  - 兄弟 `ucPieChart` 在工程里 ⇒ 语义层的 `lookupModule` 命中那枚 **external** 成员（`Public Property Get Count`，`ucPieChart.ctl:342`），于是**压根没进**「未声明」那一格（`semantic_analyzer_expr.cpp:170-184`）⇒ 既不登记隐式局部、也不出那条 3001；轮到**发码层**的 Fix 110u 发现「这名字来自别的**类**模块，按 VB6 不许裸访问」，把符号清空并按 `leakedClassMember110u` 交 `vb6_VariantEmpty()`。
  - 兄弟不在 ⇒ 语义层走「未声明 + `Option Explicit`」那一格：**只 warn、不登记隐式局部**（这正是 §B101 要升级的那一格）⇒ 发码层没有符号、也没有 `leakedClassMember110u`，落到最后那条 `lastExpr_ = cName` 交**裸名** ⇒ C2065。
  ⇒ 合成夹具两形都同形，是因为它**没写 `Option Explicit`**：两趟都进「未声明 + 宽松」那一格、都登记隐式局部（第三条路），110u 那条分支根本没被触发。**一句话根因：110u 说的是一条**语义**规则（裸标识符不得解析成其它类模块的公开成员），执行却放在发码层 —— 于是「工程里有没有同名兄弟」决定了枚未声明裸名走哪一层，两个答案都只是这个错位的副产品。**
  ⇒ 收成一处 = **把 110u 从发码层挪进语义层的符号解析**（external 符号来自 `SymbolKind::Class` 模块时就不给命中），两条出路自动汇成一条：宽松模块 ⇒ 隐式 Variant 局部（合成夹具那条形，也是 VB6 的正解），`Option Explicit` 模块 ⇒ 那一格该出诊断 —— 而诊断的严重级正是 §B101 要拍的板。**顺序因此定死：先挪 110u（本轮），再谈升 error（§B101）**；挪完 `cgen_expr_ident_builtin.inc` 里那两处（`leakedClassMember110u` 的置位与 `:253-256` 那条 Empty 出口）应一起撤干净，别再留兜底。
- **分岔已定位到一格（2026-10-10，同一轮接着量的）**：把 `UserControl=ucPieChart\ucPieChart.ctl` 那一行从**工程副本**里删掉（`ucPieChart.ctl:342` 有 `Public Property Get Count() As Long`），同一枚 `ucProgressCircular.ctl` 的同一句就从 `vb6_VariantFromValue(vb6_VariantEmpty())` 翻成 `vb6_VariantFromValue(Count)`（裸名）⇒ **兄弟类的一个属性名，改变了「另一个类里一个未声明裸名」的发码兜底**，而两趟的诊断一字不差（都只有一条 VB3001）。⇒ 两个答案里没有一个是对 VB6 的（VB6 有 `Option Explicit` 时这一句根本编译不过）；这一格归 §B121，下一刀先查「兄弟类成员在哪一步被当作本工程已知名字接住」（账 #222/#219 记过 driver 的跨模块 CoClass 登记会在本模块作用域留下符号，同一族的另一张脸），再定未声明名的发码口径。
- **全部 10 行 VB3001 按「发码实际怎么走」分完类（2026-10-10，一枚两文件的最小复现 `Count`：只有 C1 ⇒ 裸名；加上有 `Public Property Get Count` 的 C2 ⇒ Empty）**：
  - **折成 Empty（`leakedClassMember110u` 那一条）**：只有 `tests/Charts 2020/Proyecto1.vbp` 一处 —— 而它恰好是**唯一「靠这条兜底才真编得过」的正例**（`Test-GuiVbp Charts2020` 真编真跑），代价是把夹具里 `Count`（应为 `lCount`）那个笔误 **静默变成 gradient 分支永不执行**（VB6 里这句根本编译不过，因为该 .ctl 第 24 行有 `Option Explicit`）。
  - **交裸名**（`lastExpr_ = cName;` 那一条）：`ucProgressCircular/Proyecto1.vbp`（真编 ⇒ C2065，就是它今天不绿的直接原因之一）、`interp_neg/in_n3`、`in_n4`、`test_caseis_neg.bas`（三枚刻意负例，只吃 `--emit-c`）、`pkg_xmod/friend_host.bas` 与 `friend_open_host.bas`（单文件输入，发成 `Hidden1();` 这种裸调用；它们的项目形态分别由 `friend_bad.vbp`(VB7006) 与 `friend_open_ok.vbp`(0 条) 判)。
  - **两条都没走**（名字被别的路径吃掉）：`ext_show_test/Module1.bas` 的 `Form2` 与 `VBFlexGridDemo/Common/Startup.bas` 的 `MainForm` / `InitVisualStylesFixes` —— 前两个是 `Load/Unload <窗体名>` 那一路（§B111/§B113 的地盘），产物里根本没有那个名字。
- **⇒ 结论（这一格不能单独修）**：把两条兜底统一成任何一条都有代价 —— 统一成 Empty 会把「裸名交给 C 判定」那条路（账 #6 那族：`VK_UP` 靠 windows.h 的同名宏被救活）**静默改成另一个值**；统一成交裸名会让今天**唯一那枚正例** （主 Charts）当场编不过。⇒ 正确的批次是**与 §B101 的升级同批**：升级把「`Option Explicit` + 未声明」在 VB 层判死之后，两条兜底都只服务**宽松模块**，那里 Empty 才是 VB6 的正解（隐式 Variant），而夹具笔误由升级当场逼出来（届时主 Charts 要么改那一个词 `lCount`，要么老实红）。

### B120 账 #278 第十一刀已出 = 宿主符号的**装配**与**四档名单**各收成一处：§B115 那句「拼法只许来自表」到这里才算落完（语料零暴露的一条真洞顺手堵上，2026-10-10，门 #453 attempt 1 全绿（run 37976675602、head `e211bdc5`、12/12 含形状门 ⇒ 本机 A/B 的「语料零暴露」被 CI 那台 398 行复算证实、wall 11m06s））

- **撞上的方式**：第十刀把 With 那一处改问 `canonicalHostPseudoMember` 之后，给自己留了一句「对象那一段仍按源码拼写抄」。开工先把这条契约的全仓读数取了一遍 —— `vb6_<对象>_<成员>` 的**装配**在发码侧有**四份**（With 块 / 赋值 `cgen_assign_host_pseudo.inc` / 裸名 `cgen_expr_ident_builtin.inc` / 限定符 `cgen_expr_member_m22_module.inc`），而「是不是宿主伪对象」这四档名单抄了**五份**（cgen_with 两处、assign、util_type、obj_dispatch）。⇒ §B115 只修了四份装配里的一份。
- **那条洞是真的，不是纸面**：表里 `obj` 列从前是**小写**，所以「对象那一段」压根没有权威，装配时只能抄源码拼写。实测（同一枚 .ctl，只改大小写）：`With UserControl` 通，`With usercontrol` 发 `vb6_usercontrol_hWnd`（RTL 无人声明）⇒ 真编 C2065；限定符那一形 `usercontrol.hDc` 同一条洞发 `vb6_usercontrol_hDC`。VB6 源码不区分大小写 ⇒ 这一格在真工程里的表现是「换个写法就编不过」，而写法和语义毫无关系。
- **落点**：表里 `obj` 列改存 **RTL 那一份拼法**（PascalCase），`hostPseudoFind` 两边折小写（一行）⇒ 对象那一段有了权威，且逐行对 RTL 那道检查从此查的是**权威自己**；新增 `CCodeGen::hostPseudoRtlSymbol` 一处装配（`vb6_` + 表 obj + `cIdent(表 rtl，查不到时源码成员名)`），四个发码点改问它；新增 `hostPseudoObjectKnown` / `hostPseudoIsObject` 把五份名单换成一处。**非宿主的那类限定符**（工程内模块名 `Mod.Foo`）那条路一字不动 —— 刻意不接这个出口，接了就是把「模块.成员」也拖进宿主表的答案里。
- **判据（还是那道哨兵，新加第 7) 段；没有新开道，仍 49 道）**：结构侧钉「出口定义 1 / 声明 1 / 四条发码路各问 1 / `src/backend` 里 `objLower* == "usercontrol"` 手抄名单 0 处」；行为侧加一枚**大小写发码针**（新夹具 `tests/dochost/dhWithCase.ctl`：小写 `With usercontrol` + 小写 `usercontrol.hDc`）钉 `(void*)vb6_UserControl_hWnd` 与 `n = vb6_UserControl_hDC` 两头在场、`vb6_usercontrol_` 0 次。表本身从前的「行数 < 40 才红」收成**恰好 54**（改这张表要同批改这道针 —— 少认一行 = 那一行的逐行 RTL 检查隐身）。
- **同一枚坑第二次踩（值得记两次）**：`Select-String` 默认**不区分大小写** ⇒ 新写的行正则 `(UserControl|PropertyPage|Extender|Ambient)` 配上退回小写的那一行**照样匹配**，于是负控 N4 红的是 `ROW-SYMBOL-MISSING`（红是红了，说的却不是这件事）。加 `-CaseSensitive` 之后才红成 `TABLE-ROWS: 53`。⇒ **凡拿正则去认权威自己的拼法，匹配与比较都必须区分大小写**（与第十刀那条 `-ne` 同一族，同一天第二次撞）。
- **负控 6/6（全在 `src`+`scripts`+`tests` 的副本上跑，走哨兵的 `-Root`；共享树未动）**：N0 = 本刀这套判据吃 **第十刀那台 exe** ⇒ `NEEDLE-CASE` 三条红 ⇒ 这道针抓的正是本刀；N1 名单回来一条 ⇒ `MEMBER-LIST`；N2 赋值退回现拼 ⇒ `ASSEMBLY-SITE`；N3 出口定义写两遍 ⇒ `OBJECT-KNOWN`；N4 表里一行退回小写 ⇒ `TABLE-ROWS`；N5 无关改写 ⇒ 仍绿。
- **护栏（34 份产物读出来的，不是推的）**：`--emit-c` 在**改前/改后逐字节相同、rc 一格未变** —— 12 份吃宿主 With/赋值的（`VBFlexGridDemo.vbp` 5,465,077 / Charts 五份 / `czFormDemo.vbp` 229,736 / `czUI.ctl` 199,426 / 三枚 dochost 夹具）+ 22 份吃「限定符.成员」那条路的（vbman_host / pkg_* / acc / asm / interp_neg / test_caseis / FormEvents / VBFlexGridDemo 的 Common *.bas）。⇒ **语料零暴露**（corpus 里宿主对象名全是规范大小写），本刀的价值是「换个写法」再也撞不开那一格；形状门那 398 行交 CI 复算，census 仍 10 行。

### B118 账 #278 第十刀已出 = With 块那枚宿主句柄的拼法改问那张表：§B115 落地 + 第八刀扣着的 `.pag` 值位一起放开，而**它当时那条理由被实测推翻**（零行为改动、语料 8 份输入逐字节相同，2026-10-10，门 #452 attempt 1 全绿（run 37971715373、head `6b94caf5`、12/12 含形状门 ⇒ 本机登记的 398 行哈希与 CI 那台复算对齐、wall 10m01s））

- **开工第一件事把上一账的前提量了一遍，结论是它错了**：拿门 #451 的 `c3-exe` 工件（`.build/b191_base`， 就是第九刀过门那台）在同一枚 `.pag` 探针上 `--emit-c` ⇒ **rc=0**，产物里同时有 `void* _vb6_with_0 = (void*)vb6_PropertyPage_hwnd` 和 `n = vb6_PropertyPage_hWnd;` 两行；再去 RTL 数， 两种拼写**都声明且都定义**（见 §B115 的订正）。⇒ 教训：**一条「所以要扣着某格放行」的理由， 如果来自推断而不是产物读数，就在它拦着放行那一刻去量**；这一格因此多活了 24 小时的 VB3001。
- **缺陷的真实形状 = 同一个事实两个答复**，而且只差大小写。今天两枚全局都是 NULL（§B119：`.pag` 的宿主 全局全仓 0 个写者）⇒ 现价 0；一旦有人给 `…_hWnd` 接上写者，With 块读到的还是没人写的那一枚 ⇒ **静默错宿主**。所以本刀是**零行为改动的收口**（与账 #234「拿 DC 这个决定实现了两遍」同族）， 不是修一个正在响的症状 —— 别把它记成修复。
- **落点一格，不开新权威**：`src/backend/stmt/cgen_with.cpp` 那句改成问 `canonicalHostPseudoMember` （读 `kHostPseudoRows` 的 `rtl` 列，与赋值 `cgen_assign_host_pseudo.inc`、裸名 `cgen_expr_ident_builtin.inc` 同一条出口）；匹配也从「`== "UserControl"` / `compare(0, 12, "PropertyPage")` 前缀」改成表里那两个对象名 的精确匹配（旧前缀那支会把 `PropertyPageFoo` 一起接走）。语义层跟着放开 `propertypage` 的值位 —— 扣着的理由没了，两档就同形了。 **第十一刀（§B120）把这条出口再上收了一层**：装配改问 `hostPseudoRtlSymbol`，`canonicalHostPseudoMember` 只剩成员名规范化那一问。
- **判据（两道已有哨兵各扩一条，没新开道）**：
  ① `check_host_pseudo_table.ps1` 加 4)/6)：结构侧钉「那一处问表**恰好 1 次**」(0=手抄回来、2=一行两个答复) +「`vb6_PropertyPage_hwnd` 在 `src/`（RTL 除外、注释除外）出现 **0 次**」；行为侧跑一次 `--emit-c`（只走前端， 不起 cl）钉「With 头与限定读落在**同一个符号**上（区分大小写）」「`vb6_PropertyPage_hWnd` 出现 ≥2」 「夹具上 VB3001 **恰好 1 条**，且是裸位的 `VBA` 那一头」—— 最后一句是反面证人：放开的是那张表与那两个位， 不是「凡是文档名都合法」。夹具 `tests/dochost/dhWithHost.pag` 是**新增文件而不是语料输入**（清单只枚举 `.vbp/.bas`）⇒ 形状门那 398 行不必重登记。
  ② `check_dochost_authority.ps1` 的 D3b **换读法**（不是放宽）：第八刀钉「体内查 `qualifierPos` ≥ 3」， 放开 `.pag` 那一格之后体内只剩 2 次 ⇒ 改成逐条点名守位的两档（`vba` 1 + `extender`+`ambient` 1，各恰好 1）， 并新增 **D3d** 钉文档自身那两行**不带**位置（放开之后不许退回去）。
- **自己造的一条坑，值得记住**：PowerShell 的 `-ne` **不区分大小写** ⇒ 第一版 `TWO-ANSWERS` 判据在 BASE 那台 上抓不到（差别正好只有 `h` 的大小写），N0 只报出两条。改成 `-cne` 之后才报出三条。**凡是「只差大小写」 的判据，比较符必须写死区分大小写的那一个。**
- **负控 7/7（全部跑在 `src`+`scripts`+`tests` 的副本上，走哨兵的 `-Root` 形参；共享树未动）**： N0 = 新哨兵吃 **BASE 编译器** ⇒ 三条红（TWO-ANSWERS / NEEDLE-SAME / NEEDLE-BARE）⇒ 这道判据抓的正是本刀； N1 手抄回去 ⇒ WITH-ASK(0) + OLD-SPELLING；N2 那一问写两遍 ⇒ WITH-ASK(2)；N3 再塞一枚成员字面量 ⇒ DENY 红； N4 把 `&& qualifierPos` 加回 `.pag` 那行 ⇒ D3d 红；N5 无关改写 ⇒ 两道都绿。
- **护栏（8 份产物读出来的，不是推的）**：全语料里含 `With UserControl` / `With PropertyPage` 的那 9 份文件所属的 **8 枚清单输入**逐一对比 BASE vs NEW 的 `--emit-c` ⇒ **8/8 逐字节相同、rc 全 0、VB3001 计数一格未变** （`VBFlexGridDemo.vbp` 5,465,077 / `Charts 2020/Proyecto1.vbp` 2,114,886 / `ucTreeMaps` 525,312 / `ucChartBar` 497,806 / `ucProgressCircular` 381,114 / `ucPieChart` 280,694 / `ucChartArea` 273,411 / `czFormDemo.vbp` 229,736，另 `VBFlexGridBase.bas` 56,622）。⇒ 本刀在语料上是**零暴露**，判据只能靠夹具； §B108 那条提醒反过来也成立：零差异既不证明修好了，也不证明没修。 **判据只住一处**：这条发码针留在 `check_host_pseudo_table.ps1` 里（它跑在门的 [STATIC] 那一趟，CI 同样 enforce），刻意**没有**再往 `tests/run_tests.ps1` 的 `Test-CodegenNote` 家族抄一份 —— 同一件事两个住所正是本刀要修的东西。
- **census 未动**：语料 VB3001 仍 10 行（`.pag` 里 `With PropertyPage` 在语料是 0 处）。本刀关掉的是 §B115 那一格（它本来不在 §B116 那份前置清单里，是清单上面那一格）—— §B101 之前缺的仍是 §B116 列的三条负例针 / 五份单文件输入（任务 #282 的口径题）/ 源码 bug `Count`，另加下面 §B119 那一格新账；那一格不在 §B101 的路上，别把它算进升级的前置。

### B119 新账 = `.pag` 的宿主全局全仓 **0 个写者**：属性页里文档自身那一族的读数一律是缺省值（2026-10-10 量，**未开工**）

- **读数**：`vb6_PropertyPage_hwnd` / `_hWnd` / `_ScaleMode` / `_ScaleHeight` / `_Changed` 只在 `src/rtl/core/vb6rtl/vb6rtl_com.c:1072-1077` 各定义一次（初值 NULL / NULL / 1 / 0 / 0），全 `src/` grep **没有任何一处赋值**；对照 `.ctl` 那一族是有写者的（`vb6_UserControl_hWnd` 由 `uc_host.c:192/223` 换入换出，ScaleMode 那一条由账 #197/#198 接上了两条创建路）。
- **后果**：属性页里 `With PropertyPage`、`PropertyPage.hWnd`、裸写 `ScaleMode` 拿到的是 NULL 和缺省， 而不是这一页自己的宿主窗口 —— 与账 #197 同一形状，只是那一族修在 `.ctl`，这一族整个没人接。
- **兼容别名要一并想清楚**：RTL 同时留着大小写两枚（注释自称 "both spellings denote the same concept"） ⇒ 一个事实两枚存储。接写者时只接一头：另一头要么做成真别名，要么删掉 —— 第十刀之后**发码侧只会交 带大写 H 的那一枚**，所以那头已经没人写了。
- **语料暴露**：VBFlexGridDemo 三枚 `.pag` 里 `With PropertyPage.SelectedControls(0)` 6 处，那一形走 HPF_METHOD 通道、不读这枚全局 ⇒ 症状今天不现形。排产前先量「有没有真工程读属性页的 `.hWnd`/`ScaleMode`」。
- **定性已量（2026-10-10 本刀之后顺手做的，结论：这不是「RTL 漏写一枚全局」）**：拿 `tests/VBFlexGridDemo/VBFlexGridDemo.vbp`（`PropertyPage=` 三行都在）出码 —— 三枚页的类**照常发**（`vb6_cls_PPVBFlexGridGeneral_New()` 在、`PropertyPage_Initialize/ApplyChanges/SelectionChanged` 与控件 `_Click` 都在、页内部自调也在），但 **`vb6_cls_PP*_New()` 的调用点 0 处**；且 `src/` 里 `IPropertyPage` / `IPropertyPageSite` / `ShowPropertyPages` / `vb6_PP_` **全仓 0 命中** ⇒ C3 压根没有属性页运行期，没人创建页实例 ⇒ 宿主句柄恒 NULL 是**自洽**的，接上写者也没有写它的人。
- **排产含义**：清这一格的前提是先拍「C3 要不要有属性页运行期」（VB6 那一套是 IDE 宿主的页容器，编译产物自己从不显示它）⇒ 这是**功能范围决策**，不是缺陷修复，别把它当账 #197/#198 那一族的续刀排。


### B117 账 #278 第九刀已出 = 工程级名单长出第五格：标准模块的 `Friend` 过程（census 11→10，整工程输入归零，哨兵 46 扩三条，2026-10-10，门 #451 attempt 1 全绿（run 37968379735、head `e1c92200`、12/12 含形状门 ⇒ 本机登记的 398 行哈希与 CI 那台复算对齐、wall 10m48s））

- **怎么撞上的**：给 §B101 数前置清单时逐行读 census，发现第 8 行 `friend_open_ok.vbp → OpenSecret`
  的**输入是整工程**，而 `pkg_s03_friend_open` 是一枚 `Test-Vbp` **正例**（真跑出 `PKG-FRIEND-OK`），
  产物里那句一直是 `vb6_FrOpenPkg_OpenSecret();` ⇒ 语义层那条 VB3001 是**假话**，而 §B116 把它记成了
  「刻意负例」。**教训：census 的行要按「这枚输入是谁的夹具、正例还是反例」分类，不能按名字像不像噪声分类。**
- **根因一格**：`driver_semantics.cpp` 那份工程级名单（`projPubProcs`）的条件是 `acc == Public`，
  而 VB6 里标准模块的 `Friend` 过程**就是工程级裸名**（`src/common/types.hpp` 那句
  `Friend = 2, // VB6无此关键字，保留` 是错的：parser 认、发码认、兄弟模块一直调得动）。
  §B116 从 ① 那格挪进 ②/新增的这条订正见上。
- **两种后果与 §B106 同一对，而这第二种不是噪声**（两台同一枚 exe 各跑一遍，探针在
  `.build/b176_fr`（`Option Explicit`）与 `.build/b182_loose`（宽松））：
  ① 严格模块：`Friend Sub SecretSub` 被兄弟模块裸调 ⇒ 产物对（`vb6_Provider_SecretSub();`）而多配一条
     VB3001 ⇒ 改后 **1→0**；
  ② 宽松模块：**那枚名字落成一枚隐式 Variant 局部**，调用发成
     `vb6_VARIANT SecretSub = vb6_VariantEmpty();  /* 隐式变量 */` + `SecretSub();`
     ⇒ 真编译 `error C2063「不是一个函数」`、**BUILD-RC=1 零产物** ⇒ 改后出 exe 且真跑出
     `FRIEND-SUB-OK / PUB-SUB-OK` 两行。⇒ 又一格「**门绿 + 有 VB3001** 必须逐条读」的实物
     （§B106 那一族的第一格编不过样本：前面几格都只是产物对而诊断错）。
- **落点一格，且不新开权威**：放行只认「工程里有没有这个名字」，而**包边界**那一份答案早就在
  `driver_compile.cpp` 按 manifest 算好了（`packageBlockedNames_`：`modExported && (public || friend && friendVisible)`）
  ⇒ 新分支**去问那张现成的表**，不在名单建造点抄一遍 manifest 规则。这一条是硬的：`XmodPkg` 写
  `Friend=False`，它的 `Hidden1` 若被放行，`friend_bad.vbp` 那声 **VB7006**（`Test-VbpBuildFail
  pkg_s03_friend_blocked` 钉的「is not exported by package」）会被吞掉 ⇒ 放行吃掉边界 = 把一个缺陷
  换成另一个缺陷。
- **判据（扩哨兵 46，不新开一道 —— 同一条决定的家在这里）**：`V3` 宽松档两头钉（隐式局部不许出现 +
  真出口 `vb6_PcvFriend_SecretSub();` 必须在）、`V4` 反面证人（现成夹具必须仍报 VB7006 且退非零）、
  `S3` 结构（建造点认两档 `acc == AccessLevel::Friend` 恰好 1 处；那一问 `packageBlockedNames_.find`
  全仓**恰好 1 处**）。**四条负控 + 一台旧 exe**：N1 摘掉 Friend 档 ⇒ S3 红；N2 摘掉那一问 ⇒ S3 红；
  N3 把那一问写两遍 ⇒ S3 红（`found 2`）；N4 把夹具 manifest 翻成 `Friend=True` ⇒ V4 红
  （`now exits 0`，边界真被吞）；N5 无关改写 ⇒ 仍绿。**行为负控用门 #450 的 `c3-exe` 工件**
  （`.build/b191_base`，第八刀那台）跑同一份哨兵 ⇒ **只有 V3 两条红**，其余全绿 ⇒ 这条判据抓的正是本刀。
  全部跑在 `src`+`tests` 的**副本**上（哨兵的 `-Root` 形参），共享工作树一格未动。
- **护栏与读数（三件都是工件读出来的，不是推的）**：
  ① `friend_open_ok.vbp` 的 `--emit-c` 产物**改前/改后逐字节相同**（3318 vs 3318）⇒ 严格档那一头只改「谁能答」；
  ② 语料 census **11 → 10**，且少的那一行正是整工程输入那一行 ⇒ **整工程输入首次 0 行**；
  ③ 全语料形状门重算（`scripts/emit_manifest.ps1 -Out` + `compare_emit_manifest.ps1`，不 -Bless）：
     **期望 398 行 / 实际 398 行 / 相同 398 / 哈希不同 0 / 缺席 0 / 多余 0，rc=0** ⇒ 这一刀在整份语料上
     **一行发码都没动、一个退出码都没翻**。语料里那五枚 `Friend` 过程只有两枚在标准模块（都在包里），
     而它们的两个消费者都写着 `Option Explicit` ⇒ 宽松档那一格在本刀是**零暴露**（§B108 的反向提醒：
     「零差异」不能反过来证明「以前只是噪声」—— 那一条是靠真编真跑定的，见上面 ②）。
- **§B101 的形状因此收紧**：升级要动的只剩 ① 三条负例针（必须与翻严重级同批挪进 `Test-CompileFail`）＋
  ② 五份单文件输入（口径题，任务 #282）＋ ③ `Count` 两行（源码 bug）。

### B116 §B101 那刀的前置清单（2026-10-09 量完，**未开工**）：升 error 会连带翻 rc，而形状门逐行钉的就是 rc

> **2026-10-10 读到这里先看这条**：下面三件（三条负例针 / 五份单文件输入 / 源码 bug `Count`）经 §B121 的普查之后，**不再是三件独立的小事** —— 它们是同一次发货的三面：升级判死会当场逼出 `Count`（主 Charts 是今天唯一靠「兄弟类泄漏 → Empty」那条兜底才编得过的正例），而单文件那五行的 rc 翻与不翻就是 §B101 的 a/b 两案。统一未声明名的发码兜底（§B121）必须与升级同批，理由与代价都写在那里。

第八刀之后语料 census 剩 **11 行 VB3001**。那 12 处「解析不出」当时**没有清零** —— 这一句写在第九刀
之前，是本节自己的错（下面 ① 那格把一整条**整工程输入**读成了刻意负例）：**第九刀（§B117）**
清掉的那一行 `friend_open_ok.vbp → OpenSecret` 出自一枚 `Test-Vbp` 的**正例**夹具
（`pkg_s03_friend_open`，跑起来真打印 `PKG-FRIEND-OK`），而它的调用在发码里一直是**对的**
（`vb6_FrOpenPkg_OpenSecret();`）⇒ 那是真缺项，不是噪声。清完之后 census 剩 **10 行**，
整工程输入那一档**首次归零**。所以 §B101（解析不出 + `Option Explicit` ⇒ error）现在缺的
不是修法，是**这三格后果各落在谁身上**：

- **① 刻意负例 3 行**（`nopeHere` / `alsoNope` / `nopeHereIsNotAName`）：
  判据在 `tests/run_tests.ps1` 里以 `Test-CodegenNote ... @("VB3001", "nopeHere")` 的形状**钉那条诊断在场**，
  实数是三处：**5782 / 5783 / 5793**。（`grep -c '"VB3001"'` 全文件有 8 处，另外五处 5788 / 5815 / 5824 /
  5826 / 6128 都在 **Absent 名单**里 —— 那种针要求「这条诊断不许出现」，升级动不到它。上一版把 6128 算进来
  写成「四处」就是把两类针混了。）
  升级后号不变（`error VB3001` 仍含 "VB3001" 子串）⇒ 子串判据照过，**但 `Test-CodegenNote` /
  `Invoke-CodegenProj` 那批助手先要求 exit code == 0** ⇒ 这三条会当场红在退出码上。正解不是放宽助手，
  而是把这**三枚**挪进 `Test-CompileFail` 那一族（B08e-6 就是为「Expect 诊断 + 非零退出」建的）——
  而那一步做在改动之前会当场红，见下面「动手顺序」那条订正。
- **② 单文件输入的自然结果 5 行**（`Form2` / `Hidden1` / `OpenSecret` / `InitVisualStylesFixes` /
  `MainForm`）：出自
  `emit_manifest.ps1` 把 `.bas` 当**独立输入**跑（工程里没有兄弟模块，"未声明的标识符（可能来自其他模块）"
  **是对的描述**）。升级后这五份独立输入**直接编译失败**（退出码 ≠ 0）⇒ 要拍：升级只对整工程输入生效，
  还是把这五份夹具补上兄弟模块（`Form2` 那份补一枚 .frm、`Startup.bas` 那份补它引的模块、
  `friend_*.bas` 那两份补上包 —— 注意后两份一补就变成「整工程输入」，与 §B117 那一格同形）。**这一格没定之前不能动 §B101。**
- **③ 源码 bug `Count` 2 行**（`ucProgressCircular.ctl:931`，两份 Charts 输入各报一次）：那工程本来就不出 exe，
  但它是**形状门的输入** ⇒ 见下面那条 rc。**2026-10-10 订正 + 补读数**：「那工程本来就不出 exe」只对`ucProgressCircular/Proyecto1.vbp` 那枚独立工程成立（它今天确实出不了：裸名 `Count` ⇒ C2065）；而同一枚 .ctl 在 `tests/Charts 2020/Proyecto1.vbp`（harness 里真编真跑的正例）那一趟发码把同一个裸名折成了 `vb6_VariantEmpty()` ⇒ **出得了 Proyecto1.exe（1,195,008 字节、BUILD-RC=0）**。⇒ 一个未声明名两个答案（§B121）。对升级的后果也因此要说准：rc 会从 0 翻掉的不是那枚本就不建的独立工程，而是**主 Charts 正例** ⇒ 「升级」与「改夹具源码」在同一批里耦合，先拍 §B121 那条分岔，再谈 §B101。
- **⚠ 谁都没写下来的一格（本轮量的真正收获）**：`emit-manifest.expected.txt` 每行的格式是
  `sha256=… ascii256=… rc=<退出码> bytes=… <relpath>` —— **rc 是判据的一部分**。
  今天交 warning 的输入 rc=0；升级后凡是"带 `Option Explicit` 且报了 3001"的输入 rc 变 ≠0 ⇒
  **清单里那一行的 rc 字段跟着变**，与发码一字未改无关。⇒ §B101 那一刀必须自带
  「全语料重算 + `-Bless`」，并且红话里要能分清"哪些行的 rc 翻了"（这就是它该有的判据：翻 rc 的集合
  == ①②③ 之外不许有一份）。登记走仓库自己的出口 `scripts/compare_emit_manifest.ps1 -Bless`（§B100 / §B112 那条）。
- **动手顺序**：① ~~先把四条负例针迁进 `Test-CompileFail` 族，可单独一刀~~ **订正（写下这行之后就自己推翻了两分钟）**：挪不动 —— 那族助手要的是“非零退出 + 期望诊断”，而那四条今天走的是 exit 0 的助手（`Test-CodegenNote` / `Invoke-CodegenProj` 都要求退出码为 0）。**在改动之前搬家 = 先要求改动已经发生**，门会当场红。所以 ① 只能与翻严重级**同批**落地（同一刀里：改助手归属 + 改严重级 + 重登记清单）。② 
  ② 再拍 ② 那格口径；③ 然后才翻严重级 + 重登记清单 + 一条"升级只对工程输入生效"的哨兵（钉 ①②③ 三个集合的条数，
  不许多也不许少）。④ 最后把 `--suppress-warning 3001` 那条抑制通道复核一遍：按号抑制在分家（§B102）之后才是干净的。


### B114 账 #278 第八刀已出 = 宿主表长出 HPF_CHANNEL：`Controls` 两个位放行 + `.ctl` 自身对象名放行（census 16→11，第 48 道哨兵扩三条，2026-10-09，门 #449 红在哨兵自己身上、改完由门 #450 收线）

- **§B105 那格 a/b 按 b 走**（台账当时就写着"倾向 b"，形状也是现成的：`canvas_drawing.hpp` 的
  `CANVAS_OWNER_METHOD/DRAW` 早就是"一行标由哪条码头回答"）。a 案的代价量过了：往 RTL 再放一枚
  永远不该被发码的恒 NULL 空桩 = 第二个陷阱，而表里那一行从此说谎。
- **表长出第二份契约**：`HostPseudoRow` 末尾加 `channel`（只在 `HPF_CHANNEL` 行填），
  `vb6_<对象>_<rtl>` 那条老契约对它不适用。三处消费同一件事实：语义层 `isDocumentChannelMember`
  （**两个位**都放行——`Controls.Add(…)` 里 `Controls` 站的正是限定符位，而老的两条分支一条只管
  限定符位的"文档对象名"、一条只管裸位的"文档成员名"，集合名当限定符用两头都不接）、
  发码层 `cgen_expr_ident_builtin.inc` 不再硬编码 `lower == "controls"` 与 `"vb6_UC_Controls()"`、
  哨兵换判据。
- **`isDocumentHostObject` 加了一个位置形参而不是新开一条**：`.ctl` 里自身对象名在值位也合法
  （`With UserControl` ⇐ VB6 等价于 `Me`），实测产物本来就有路（`_vb6_with_2 = (void*)vb6_UserControl_hWnd`），
  今天只是多配一条 VB3001。**同一句写在 `.pag` 里刻意不放** ⇒ 见 §B115。  **第十刀订正**：那一格现已放开，而当时扣着它的那条理由（小写 h 一发码就 C2065）实测是假的 —— RTL 两种拼写都声明且定义了 ⇒ 读数与后果见 §B118。
- **读数（改前那台 exe = `.build/b136_C3_new.exe` 与改后同一批输入）**：`.ctl` 探针 3→0、
  `.pag` 探针 2→**1**（留下的那条就是反面证人）、Charts 两份输入各 3→1（剩的是源码 bug `Count`）、
  `VBFlexGridDemo.vbp` 1→0；四份输入的 `--emit-c` 产物**逐字节相同** ⇒ 这一刀只改"谁能答"，
  不改"答什么"（与 §B108 那条相反：这次零产物差是真的零行为差）。语料 census 16→**11**。
- **哨兵（第 48 道，扩而非新增）**：`CHANNEL-ROW-NODOCK`（标了旗标却不点名码头）、
  `CHANNEL-ROW-TYPE`（码头行还声称有值类型）、`CHANNEL-DOCK-MISSING`（码头符号在 `src/rtl/*.h` 里
  没人声明）；`$must` 加 `hostPseudoChannel` 的声明 + 两个消费点，`$deny` 加"发码侧再硬编码一次
  `vb6_UC_Controls()`"。本机绿：`OK: one host-pseudo table (54 rows, 35 scalar)`。
  ⚠ 一条工具读数：`#include "common/host_pseudo.hpp"` 第一次加在 `cgen_util_ctrl.cpp` /
  `cgen_util_type.cpp` 上都没用 —— 那个 `.inc` 的真正宿主是 `src/backend/expr/cgen_expr_ident.cpp`；
  报错只有 C3861「找不到标识符」，不告诉你在哪个 TU。
- **12 处清到 9 ⇒ 真缺项只剩 §B115 那一格与它后面的东西**：`Controls` 2 + `UserControl` 1 已放行，
  剩下的 census 11 行 = 刻意负例 6 + 单文件自然结果 3（§B110 那条口径题）+ 源码 bug 2（`Count`）。
  ⇒ §B101（解析不出 + `Option Explicit` ⇒ error）前面**只剩那一格口径**了。
- **门 #449（run 37960257781、head 52ff3502、attempt 1，16:35:50Z→16:45:00Z）唯一红 = `Tests (compile)`，
  红的不是产品而是这一刀的哨兵**：`[STATIC] dochost_authority ... FAIL`。
  那条 D3 钉的是**改形之前**的调用点（`memberObjCtx_ && isDocumentHostObject(node.name)`），
  而这一刀把「哪个位合法」搬进了唯一出口（调用点变成 `isDocumentHostObject(node.name, memberObjCtx_)`）
  —— 两条正则同时失配（调用点 0 次、外面那道 `&&` 也 0 次）。
  本机复现同一对 FAIL，所以门那条红没有第二种解释。
- **修法是把不变量搬到它现在住的地方，不是把 D3 放宽**：D3 三处（定义 / 声明 / 调用点）都改成「带着那个位置形参」；
  新增 **D3b** 钉那个出口**真的**按位置分档（体内查 `qualifierPos` ≥ 3 —— `vba` / `propertypage` / `extender`+`ambient` 各拒一次裸位），  **第十刀换读法**：放开 `.pag` 那一格之后体内只剩 2 次，D3b 没有放宽成 ≥2，改成逐条点名守位的那两档（`vba` / `extender`+`ambient` 各恰好 1）并加 D3d 钉放开的那两行不许退回按位置扣。
  这条接手了老 D3 守的那件事（**裸位的真缺项必须还在响**）；
  新增 **D3c** 钉调用点不许在外面再 `&&` 一次（两份答案 = 老形状）。
  哨兵另外长了 `-Root` 形参，为的是能在**做过的副本**上跑负控而不碰共享树。
- **三条负控各红各的那一条**（`.build/b163_negctl.py`，副本已删）：
  A 调用点改成外面 `&&` ⇒ D3 调用点 0 + D3c 1；
  B 留着形参但体内不用 ⇒ D3b 0；
  C 声明里去掉那个形参 ⇒ D3 声明 0。
  还原原始副本 ⇒ `PASS … predicate 1+1+1 pos 3 outer 0`。
- **收线的那一轮 = 门 #450**（run 37962617927、head `7858c69e` = 4dd9ff27 哨兵改形 + 7858c69e 台账、
  attempt 1，16:55:48Z→17:05:14Z ≈ wall 9m26s）：
  **12 条 check-run 全 completed/success、非绿 0**，
  含 `Tests (compile)`（上一轮红的那一格）与 `Emit manifest (shape oracle)` ⇒ 第八刀连同它的哨兵改形一起过了。
  #449 那一轮因此有了第二种用途：它是**本线第一次把「门红 → 归因到哨兵自己 → 改判据形状 → 三条负控自证 → 复跑绿」走完整条**，以后 STATIC 那一格红了先按这条走，别默认是产品回归。
  顺带跑的邻居：`check_static_sentinel_registration`
  （49 道全登记）、`check_host_pseudo_table`（54 rows / 35 scalar）都绿，PSParser 0 错。

### B115 账 #278 第十刀已出（落地与判据见 §B118）= `With PropertyPage` 那句拼法改问那张表（2026-10-09 量、2026-10-10 **订正本账当时那条 C2065 的判断**）

- **本账留下的唯一一条订正**：当时写的是「`.pag` 里发码把 `hwnd` 的小写 h 直接拼出来 ⇒ 与 RTL 的
  `extern … vb6_PropertyPage_hWnd` 对不上 ⇒ 真编译是 C2065」。这个「真编译」从未真编过 —— RTL 把
  **两种拼写都声明且定义了**（`vb6rtl_com.c:1072/1073` 两行、`vb6rtl_userctl.h:181/182` 两条 extern，
  中间那句注释是 `// both spellings denote the same concept`），BASE 那台编这枚探针 rc=0。
  缺陷的真实形状不是「编不过」而是「**一个事实两个答复**」：同一份产物里 With 头取 `…_hwnd`、
  限定读取 `…_hWnd`，两枚各自独立的全局。⇒ 那条被用来**拦着一次放行**的推断，代价是一格合法语义
  多配一条 VB3001，一直配到第十刀。教训写在 §B118 第一条。
- **判据方向那句是对的**（`vb6_<对象>_<成员>` 的拼法只许来自表），第十刀照它落的。


### B113 账 #278 第七刀已出 = `Load/Unload <窗体名>` 的实参是**对象位**：§B111 那一格收掉（接已有的闸，不是新加判定，2026-10-09，门 #448 attempt 1 全绿（run 37947447538、head `028f7342`、12/12 含形状门 ⇒ 本机登记的哈希与 CI 那台复算对齐、wall 10m11s））

- **定性沿用 §B111 的实测**：VB6 里 `Load` / `Unload` 的实参站在对象位，**不取默认属性**。改前窗体模块里指着
  自己那枚窗体名发的是 `vb6_UnloadForm(vb6_GetControlText(vb6_hwnd_FDForm)  /* default prop: .Caption */)`
  = **BSTR 进 `void*` 槽**；而同一句写成 `Unload Me` 一直是 `vb6_UnloadForm(vb6_hwnd_FDForm);`
  ⇒ 同一条 VB 语义两条路两种答案（与 §B104 的"发码对、诊断错"反方向：**发码错、两层都不响**）。
- **落点 = 把新形接进已有那一条闸，不是再写一份判定**：`cgen_expr_call_arg_emit.inc` 里"被调方形参声明为
  `Object` ⇒ 抑制默认属性"那一支（Fix 142 时代就有）只对**有 `calleeParams` 的被调方**生效，内置的
  `load` / `unload` 没有形参表 ⇒ 走不到。新支把"实参是工程内窗体名"也送进同一个 `suppressDefaultProp_`，
  抑制之后用的是**唯一出口** `ctrlObjectRefExpr`（与 `Set` 的右值同闸，账 #258）⇒ 拼法仍然只有一份。
- **三形同归的实测**（同一枚 exe 改前 / 改后各出一趟码）：
  `Unload Me` → `vb6_UnloadForm(vb6_hwnd_FDForm);` **未变**；
  `Unload FDForm`（自己）→ 改前 fold=1 / 改后 fold=0，且对象位那一形从 1 处变 **2 处**；
  兄弟名 `Unload PfOther` → 两台都是 `vb6_form_hwnd_PfOther()` **未变**（那条本来就对）。
- **语料暴露 0 处 ⇒ 判据自己造，但要造在不动别人判据的地方**：`tests/fdraw/FDForm.frm` 末尾加一枚
  **没人调用**的 `Public Sub UsOwnNameProbe()`（cgen 会发模块里每一个过程 ⇒ 形状被钉住，而夹具真跑那 26 行判据
  一条不动）。本地真编真跑过：cl 出 exe、`FD-DONE` 在、`=False` **0 条**、耗时 30.9s（构建）+ 一次跑完自然退出。
- **新判据一条两头钉** `Test-CodegenNote "form_unload_ownname"`（`tests/run_tests.ps1`，紧跟 fdraw 那三条画布针）：
  needle `"vb6_UnloadForm(vb6_hwnd_FDForm);"` 必须在（拦"把调用整个删掉"这种坏修法）、
  absent `"vb6_UnloadForm(vb6_GetControlText(vb6_hwnd_FDForm)"` 不许在（拦本次真伤）。
  负控 = 改前那台 exe 出码：absent 那条当场红（fold=1）⇒ 不是单侧绿灯。
- **形状门**：本机全语料重算 + `compare_emit_manifest.ps1 -Bless` 登记，差异**只有 `tests/fdraw/FDemo.vbp` 一行**
  ⇒ 这一行同时是"改动面就这么多"的证据：own-name 的 `Load/Unload` 在全语料 0 处，别家产物本该一字不动。
- **刻意不收的一形**（写死在注释里，别下次又当本账的尾巴）：`Unload Picture1` 这种**控件名**在 VB6 是
  error 43（Object required），今天交什么继续交什么 —— 本账只并"窗体名的两条路"，不去替 VB6 的错误行为发码。


### B112 账 #278 第六刀已出 = `VK_UP`：**未声明的裸名被 C 的同名宏救活**，编译器与门都看不见（第 16 行 / 真缺项清到 9，2026-10-09，门 #447 attempt 1 全绿（run 37943664181、head `ca0ee097`、12/12 含形状门；两条新发码针跑在 vbp 四片里，四片全 completed/success））

- **这一格不是编译器的缺陷，是夹具的缺陷，而门一直是绿的**。`tests/tabwalk/WalkForm.frm` 声明了
  `WM_KEYDOWN` / `WM_KEYUP` / `VK_TAB` / `VK_DOWN`，**独漏 `VK_UP`**（`AK-pre/down/wrap/up` 那一路 step2 用它）。
  发码把裸名原样交出去，`<windows.h>` 里同名宏（`winuser.h`: `VK_UP 0x26` = 38）**恰好**接住 ⇒ 两架构真跑全绿、
  链接全过，只有诊断配了一条 VB3001；而 VB6 本人会直接拒源（`Option Explicit` + 未声明标识符）。
- **改前 / 改后同一枚 exe 的产物读数**（不是推理）：
  改前 `vk = vb6_ChkLong(vb6_IIfLong((((-(gAkStep == 2))) != 0), VK_UP, 40));`
  改后 `…, 38, 40));`，且前导多一行 `#define VK_UP (38)`。VB3001 那一份输入的 stderr 从 123 字节（一条）变 **0 字节**。
- **形状门：只有这一行动**，且登记走仓库自己的出口 `scripts/compare_emit_manifest.ps1 -Bless`（保持现有行序、
  按路径替换、注脚留着）⇒ `emit-manifest.expected.txt` 的 numstat 是 **+1/-1**。
  **没有手抄哈希**（手抄一份判定 = 又造一个权威，正是本项目一直在防的形状）。全语料本机重算一遍：
  398 行里 **397 行逐行相同**、唯独 `tests/tabwalk/TabWalkApp.vbp` 变（`b80ef267…`、bytes 32431→32448）
  ⇒ 其余输入真的没被牵连，这条同时是"改动面收窄"的证据。
- **值不是背出来的**：`D:\Windows Kits\10\Include\*\um\winuser.h` 实测 `VK_UP 0x26` / `VK_DOWN 0x28` / `VK_TAB 0x09`，
  与夹具已有的那两枚一致才动笔（§B165/#165 那轮三次栽在**抄错的常量**上，凡写数值先回头对 SDK 头）。
- **新判据两条**（`tests/run_tests.ps1`，就钉在 tabwalk 现有那组旁边）：
  `tw_emitc_vkup_folded` 钉 `, 38, 40));` **必须在**；`tw_emitc_vkup_bare` 钉 `, VK_UP, 40)` **不许在**。
  负控 = 拿**改动前那份夹具**跑同一枚 exe 出码：两条各自红且方向相反（folded 缺 / bare 在场）⇒ 不是单侧的绿灯。
- **12 处到现在（订正 §B110 末段那句"剩 4 处"）**：本刀清 `VK_UP` ⇒ 累计清 **9**，
  剩 **3 处** = 裸 `Controls` 2 + 裸 `UserControl` 1（§B105，撞在宿主表对 `rtl` 字段的契约上，口径 a/b 待拍）。
  census 由 17 行降到 **16 行**（同一枚 NEW exe、同一份 `.build/b125_census.py` 的口径）。
- **可复用的那一族形状（本刀真正的收获，比这一格值钱）**：
  **「未声明的裸名被 C 的同名宏救活」在三个信号里只有诊断在报警** —— 编得过、链得过、真跑绿，
  而今天的 VB3001 只是 warning。⇒ ①**"门绿 + 有 VB3001"这一组合本身就值得逐条读**，别当噪声攒着；
  ②§B101 那刀把它**升成 error 之后，这一族会当场变成硬错误**，那正是升级的价值 —— 但升之前 census 必须
  逐条归因完成，否则"把夹具的真缺陷升成编不过"与"把合法写法判成噪声"是同一个错的两个方向。
- 顺带同一次扫描的读数：语料里 `Unload <自己的窗体名>` = **0 处**（own-name 0 / 兄弟名 1 / 其余一律 `Unload Me`）
  ⇒ §B111 那条继续作为潜伏项挂着，不占本轮排产。


### B110 账 #278 第五刀已出 = 工程级**窗体名**进第四份名字表（`TmForm2` 那一处清掉，第 49 道哨兵，2026-10-09，门 #446 attempt 1 全绿（run 37936265039、head `1bda219c`、12/12 含形状门、wall 10m25s））

- **缺的是什么**：VB6 里**窗体名**站在裸名位就是它的默认实例（`Unload TmForm2` / `Set f = TmForm2`）。发码侧一直答对 ——
  `cgen_expr_ident_symbol.inc` 的 Fix 086 支路拿 `knownFormModuleNames_` 查，命中就发 `vb6_form_hwnd_TmForm2()`；
  语义层那份"工程级已有的名字"只有三格（Public 过程 / Public Const / 模块名），而模块名那格**只认限定符位**
  ⇒ 同一句里 `TmForm2.Visible` 不报、`Unload TmForm2` 报。census 里那条 `TmForm.frm:131` 就是这么来的。
- **不能并进模块名那一格**：模块名只许站限定符位（`Module1.ShowForm2` 里认 `Module1`），窗体名**两个位都合法**。
  合并就把裸写的模块名也一起放行了 —— 那是把一条 VB6 里根本不合法的写法判成合法。
- **这次 A/B 的"产物零差"该怎么读（与 §B108 那条规矩相反的一面，值得钉）**：自建探针 `.build/b123_probe`
  （一枚 Option Explicit 的 .frm + 一枚**宽松** .bas，都用裸名指同一枚兄弟窗体 `PfOther`）实测两档：
  BASE = `vb6_UnloadForm(vb6_VariantToObjectVal(PfOther));` 配一行 `vb6_VARIANT PfOther = vb6_VariantEmpty();  /* 隐式变量 */`
  ⇒ **卸掉的是那枚空 Variant，不是窗体**；NEW = `vb6_UnloadForm(vb6_form_hwnd_PfOther()  /* form default instance */)`。
  而全语料 398 输入里"宽松模块裸用兄弟窗体名"= **0 处** ⇒ A/B 的 `artifacts changed = 0` 是**零覆盖**，不是无害。
  ⇒ 同一条结论再确认一次：**产物差异不能当"这条诊断值不值得修"的判据**，红只能自己造（本刀 V3 就是为它写的）。
- **落地 = 第四份工程级名单，且与发码侧同一个建造点**：新增 `Driver::collectFormModuleNames()`（定义在
  `driver_semantics.cpp`，声明在 `driver.hpp`）；两处消费 —— 语义侧 `setProjectFormNames(...)` 下发，
  发码侧 `driver_codegen_typedfield_scan.inc` 里那份**就地扫描 `modules_` 撤掉**换成调用它。
  判据落在 `namesProjectLevel` 第四格。留两份扫描 = 同一事实两份权威，将来谁改判据谁漏改另一边（§B97/§B107 同族）。
- **那条自我排除不是装饰**：`projFormNames_.count(lk) && !(currentModule_ 与它同名)` —— 抄的是发码侧
  `lower != knownFormName_` 那半句，两层必须同一个答案。放行"窗体模块里指着自己那个名字"= 把一条发码侧走**另一条支路**的
  写法判成合法；而实测那条支路的答案本身是坏的 ⇒ 见 §B111（先把口径对齐，别把坏答案钉成规范）。
- **A/B（全语料 398 输入，BASE = `.build/b123_C3_base.exe` = 不含本刀那台 exe）**：`VB3001 18 → 17`、
  `artifacts that changed = 0`、`diagnostics ADDED = 0`、per-id delta 只有 `VB3001` ⇒ 形状门不该动一行。
- **第 49 道哨兵** `scripts/check_project_form_names.ps1`：**V1** 裸写的兄弟窗体名不许报 VB3001；
  **V2 反面证人**（同一枚窗体里另放一枚真不存在的名 `pfNoSuchNameAnywhere`）必须**仍然**报 —— 只钉放行不钉反面 = 把闸门整个关掉；
  **V3** 产物三头：`vb6_form_hwnd_PfSibling(` ≥ 2 处（严格 .frm + 宽松 .bas 各一次）、不许出现 `vb6_VARIANT PfSibling`、
  不许出现 `vb6_VariantToObjectVal(PfSibling`；**S1** 建造点唯一（定义 1 / 声明 1 / 语义侧同串 2 / 发码侧 1）+
  旧那份 `formModuleNames.insert` 不许回来 + 语义层问 `projFormNames_` 恰好 1 处 + setter 1 声明 1 调用。
  **两条负控各自红过**：A `-Exe …\b123_C3_base.exe` ⇒ V1 + V3 三条红；B 把 `namesProjectLevel` 里那一格注释掉 ⇒
  S1 单独报"semantics asks projFormNames_ 0 times"。植入都按 md5 原样还原。
- **12 处到现在**：本刀清 `TmForm2` ⇒ 累计 **8**（§B107 常量 1 + §B108 `Printers` 1 + §B109 窗体伪成员 5 + 本刀窗体名 1），
  剩 **4 处** = 裸 `Controls` 2 + 裸 `UserControl` 1（§B105，撞在宿主表对 `rtl` 字段的契约上，口径 a/b 待拍）、
  `VK_UP` 1（夹具补一枚 `Private Const`，值回 `winuser.h` 对，别抄记忆里的数 —— §B165/#165 那轮三次栽在抄错常量上）。
- **【2026-10-11 重算订正】上面那句「剩 4 处」过期了**：全语料 census（`.build/b922_census_now.py` → `b922_census_now.txt`，400 份输入、同一枚 NEW exe）报 VB3001 **10 行**，真缺项那一格 = **0**。**逐行归因**：刻意负例 5（`nopeHere` / `alsoNope` / `nopeHereIsNotAName` / `Hidden1` / `OpenSecret`）+ 拆包自然结果 3（`Form2` / `InitVisualStylesFixes` / `MainForm`）+ 源码 bug 那一条的两个输入（`ucProgressCircular.ctl:931` 的 `Count`，两份 Charts 工程各一次）= 10。少掉的 7 行 = 裸 `Controls` 4 + 裸 `UserControl` 1 + `VK_UP` 1（第六刀）+ `friend_open_ok.vbp` 里那次 `OpenSecret` 1（第九刀的 Friend 名单）。
  **并排读数（项目模式 `--emit-c`，`.build/b921_b105/`）= 产物本来就是对的**：`Proyecto1.vbp` 两处裸 `Controls` 报 0 条、产物 `vb6_UC_Controls()` ×2；`VBFlexGridDemo.vbp` 裸 `UserControl` 报 0 条、产物 `vb6_UserControl_Extender` + `vb6_UserControl_ParentControls` ×2；`TabWalkApp.vbp` 的 `VK_UP` 报 0 条、产物 `#define VK_UP (38)`。⇒ **§B105 那格等拍的 a/b 口径不再需要**（它被后面的刀顺带清掉了，只是一直没人重算 census）。账 #278 只剩两件事：拆包自然结果那 3 条的口径（§B101 前面那一格，仍在等拍）+ `Count` 这条夹具源码 bug 要不要动。
- **重算 17 行（同一枚 NEW exe，`.build/b125_census.py` → `b125_census_out.txt`）**：**13 个唯一源码位置 = 17 行**，
  按成因分格：刻意负例 5 位置 / 6 行（`nopeHere`、`alsoNope`、`nopeHereIsNotAName`、`Hidden1`、`OpenSecret` 两输入）、
  拆成单文件编译的自然结果 3（`Form2` / `InitVisualStylesFixes` / `MainForm`）、源码 bug 1 位置 / 2 行
  （`ucProgressCircular.ctl:931` 的 `Count`，两份 Charts 输入各报一次）、真缺项 4 位置 / 6 行（`Controls` 2 名 × 2 输入 + `UserControl` + `VK_UP`）。
- **⇒ §B101 前面又冒出一格口径题（本刀顺流量出来的，不是代码题）**：那 3 条"自然结果"出自 `emit_manifest.ps1`
  把 `.bas` 当**独立输入**跑（工程里没有兄弟模块，"未声明的标识符（可能来自其他模块）**是对的描述**）。
  "解析不出 + `Option Explicit` ⇒ error"一旦生效，**这些独立输入会直接编译失败**，而形状门按行钉死清单 ⇒
  要么升级只对整工程输入生效，要么把那三份夹具补上兄弟模块。两种都是口径，要先拍再动。

### B111 新账 = 窗体模块里 `Unload 自己的窗体名` 交出的是 Caption 字符串（2026-10-09 量；**已由第七刀收掉 ⇒ 落地与判据见 §B113**；当时语料暴露 0 处）

- 同一个探针顺带量到的（`.build/b123_probe/PfSelf.frm`，三行体的窗体，站在**自己**的模块里）：
  `Unload PfSelf` ⇒ `vb6_UnloadForm(vb6_GetControlText(vb6_hwnd_PfSelf)  /* default prop: .Caption */);`
  = **BSTR 进 HWND 槽**，账 #229 那条"裸 int 进 BSTR 槽"的反方向同族；而同一句换成兄弟名（`Unload PfOther`）
  发的是 `vb6_form_hwnd_PfOther()` ⇒ **同一句 VB 两种答案**，差别只在接收者是不是自己。
- **对象位那一格是对的**：同一模块里 `Set ff = PfSelf` ⇒ `ff = vb6_hwnd_PfSelf;  /* Set */`
  （`cgen_form_ctrl_registry.inc:10` 把窗体自己的名字也登记进 `knownFormControls_`，类型 Form）
  ⇒ 坏的只有"取默认属性交给按 HWND 定型的出口"这一步，不是整条 own-name 通路。
- **两层都沉默**：语义层不报 VB3001（BASE/NEW 都不报 —— 模块自己的名字本来就解析得出），发码照发，链接照过
  ⇒ 只有真按 VB6 语义该卸的那枚窗体没卸。与 §B104 的"发码对、诊断错"正好相反，这一格是**发码错、两层都不响**。
- **语料暴露 = 0 处**（扫全部 `.frm` 的 `^\s*Unload\s+<名>`：own-name **0** / 兄弟名 1 / 其余一律 `Unload Me`）
  ⇒ 潜伏项，与 #244 同形，别当编译阻塞排产。
- **修法方向（先记不动手）**：语句位/HWND 位的 own-name 应与对象位**同源** —— 折成 `vb6_form_hwnd_<自己>()`，
  不是退到控件默认属性那条支路。判据两头钉：`Unload <自己>` 发 hwnd 出口 + `Unload Me` 的发码不回潮；
  再补一枚"VB3001 不许因为这条改动而新增"的护栏（本刀那条自我排除正是靠它对齐口径，两处必须一起想）。

### B109 账 #278 第四刀已出 = 窗体裸写属性那张表搬到 common，两层都问它（§B104 的堵点接上，第 48 道哨兵，2026-10-09，门 #445 attempt 1 全绿（12/12，含形状门；第一趟 #444 红在本刀自己带出的三处旧哨兵，head 9194b997））

- **定性（沿用 §B104 的实测）**：语料里那 5 条窗体伪成员（czUI `ScaleWidth`×2、VbQRCodegen `WindowState`/`ScaleWidth`/`ScaleHeight`）
  是"发码对、诊断错"—— `cgen_expr_ident_symbol.inc` 在窗体模块里拿 `getControlPropReadFn(FrmControlType::Form, name)`
  查名字表，命中就发 `vb6_Get…(hwnd)`，产物一直是好的；缺的只是语义层问不到那张表。所以本刀是**放行**，不是新能力。
- **堵点就是 §B104 记下的那一件**：表住在 backend，语义层要问就得把 496 行的 `getControlPropReadFn` 搬到家，
  而它吃 `FrmControlType`（在 `src/project/frm_parser.hpp`），common 不许向上依赖 project。**本刀的解法不是搬家，是折位**：
  表里只写 `FPQ_FORM / FPQ_COMMON_DIALOG / FPQ_SHAPE / FPQ_LINE` 四个类别位，backend 一侧用一处
  `formPseudoKindBits(FrmControlType)` 把枚举折成位（其余类型一律 0 = 原来"通用段对所有类型都答"的行为）。
  分界写法与 `canvas_drawing.hpp`（账 #234）同族。
- **落地**：新家 `src/common/form_pseudo.hpp`（`kFormPseudoRows` 34 行 + `formPseudoReadFn` + `formPseudoIsBare`）。
  backend 的 `getControlPropReadFn` 把**通用段 25 行 + 窗体臂 9 行**整张搬进表（顺序未改：通用行在前、`formOnly` 在后），
  函数里改成一次表查询，窗体那一臂**刻意留空**；语义层在"未找到标识符"那一串里加一格
  `!memberObjCtx_ && currentModule_->isFormModule && formPseudoIsBare(name)`，停的是同一批名字上的 3001 与隐式局部。
- **A/B（全语料 398 输入，BASE = `.build/b114_C3_base.exe` = 未含本刀、另在一台 worktree 编的 exe）**：
  **`artifacts that changed = 0`**（backend 那一半是纯搬迁，逐字节对上）、`diagnostics ADDED = 0`、
  **VB3001 23 → 18**，消失的 5 行恰好是 czUI 2 + VbQRCodegen 3；per-id delta 只有 `VB3001` ⇒ 形状门不该动一行。
- **第 48 道哨兵** `scripts/check_form_pseudo_table.ps1`：V1 四个真名字（Option Explicit 的窗体）不许报；
  **V2 反面证人**（同一枚窗体里放一枚真不存在的名）必须继续报；**V3 宽松窗体**两头钉 —— 产物不许出现
  `vb6_VARIANT ScaleWidth` 那枚隐式局部，且必须有 `vb6_GetScaleWidth(` / `vb6_GetScaleHeight(` / `vb6_GetWindowState(`
  三条真出口（只钉诊断的哨兵会把"什么都不发"判成绿）；**S1** 消费点唯一 + 老硬编码行不许回来 +
  **窗体那一臂里 `propLower ==` 必须 0 处**；**S2** 表 34 行不重名、语义层消费点恰好 1 处。
- **三条负控各自红过**：A `-Exe …\b114_C3_base.exe` ⇒ V1 四条 + V3 一条红；B 把语义层那行的
  `formPseudoIsBare(node.name)` 注释掉 ⇒ S2 报"0 个消费点"；C 往窗体臂植一枚
  `if (propLower == "bogusrow")` ⇒ S1 单独红。三处植入都按 md5 原样还原。
- **两条工具读数（本轮当场踩到，都写进哨兵的注释里）**：
  ① PS 里用 `LastIndexOf("`n")` 从匹配点反推"行首"来判注释，会把**上一整行**当成前缀算进来 ⇒
  一条好行被误判成注释、表计数 34 变 33；改成"逐行先切掉 `//` 之后再匹配"。（§B108 那条"跳过注释行"的
  规矩本身没错，错在反推行首的手法。）
  ② 用脚本往 **BOM + CRLF** 的文件里插函数，写回必须带 BOM 且统一 CRLF —— 本刀一次插完
  `PARSE-ERRORS=2`、`lone_lf=14`，补回 BOM 后双双归零。哨兵计数那条也跟着被带偏过一次。
- **12 处到现在**：累计清 **7** 处（§B107 常量 1 + §B108 `Printers` 1 + 本刀窗体伪成员 5）。
  剩下 4 处：**裸 `Controls` 2 + 裸 `UserControl` 1**（§B105，撞在宿主表对 `rtl` 字段的契约上，口径 a/b 待拍）、
  **`TmForm2` 1**（窗体名自动实例化，已量成同形、放行点还没找）、**`VK_UP` 1**（夹具漏写 `Private Const`，§B106）。
  这 12 处清完（或明确记账不收）才轮到 §B101 那一刀：解析不出 + `Option Explicit` ⇒ error。

- **门 #444（head `5b238770`）的红是自己带出来的，不是抖**：`Tests (compile)` 三条 [STATIC] 红 ——
  `scalemode_writers` W4 / `control_dc` D5 / `form_draw_state` S3。三条数的是**旧位置**里的
  `return "vb6_WindowScaleModeSelf";` / `return "vb6_GetControlHDC";` / 读表 Form 那一臂内的名字，
  窗体那一行搬进表之后各自少一行 ⇒ 报红。**处置不是放宽阈值**，是把它们改成问新权威
  （读侧 = backend 剩那条 + common 表里那条，合起来数；写表没搬 ⇒ 不动），
  三条都拿"把表里那一行注释掉"证过会红（W4=1 / D5=1 / S3 currentx read=0）。
  读数规矩：**搬走一张表之前先 `grep -l` 谁在数它的位置** —— 这类"钉死位置"的判据是**跟着权威搬家**的。
- **同一轮第二次踩到两条 PS 坑**（都写进过 memory，这次是补丁自己中招）：
  ① 新加的计数**没跳注释行** ⇒ 那三条改完后第一版负控"注释掉一行"仍然绿（§B108 那条规矩的第三格）；
  ② 把 `Strip-RowComments(...)` **嵌进表达式里**（`$x = @([regex]::Matches(Strip-RowComments(...), 'p'))`、
     `$rd = $rd + Strip-RowComments(...)`）在 PS 里是 ParserError —— 三条哨兵 `PARSE-ERRORS=2/3`，
     而跑起来只表现为"判据好像变了"。规矩：**命令行调用要自成一条语句**（`$raw = …` 然后 `$txt = Strip-RowComments $raw`），
     改完任何 .ps1 先 `[System.Management.Automation.PSParser]::Tokenize` 数 `$e.Count` 再谈读数。

- **搬表的完整性用机械对表证，不靠眼看**：从父提交 `git show` 里抠出旧的通用段与窗体臂两份 `if (propLower == …) return …"`，与 `kFormPseudoRows` 按 (名字, 读函数, 三条类型例外) 逐条对 ⇒ **MISSING = 0**；多出来的两条 (`fontpixelheight` / `borderstyle`) 是旧代码那两条**跨行写**的条件被我的按行正则漏抓，不是表里多了东西；旧的两条重复 (`scalewidth` / `scaleheight` 通用段与窗体臂各写一次、答案相同) 在表里合成一条 —— 这三点都要读数对上才算搬完。
- 复跑读数：本地 `-Category compile -Jobs 20 -Incremental`（CI 同一口径）从 **PASS=55 FAIL=3** 回到 **PASS=58 FAIL=0**。

### B108 账 #278 第三刀已出 = 内置全局对象那张名单少一格（`Printers`，第 47 道哨兵，2026-10-09，门 #443 attempt 1 全绿（12/12，含形状门，head 5ebc8ee2））

- **定性（先量后动，`.build/b110_printers.py`）**：语料里唯一那条 `Printers` = `tests/dbgdlg/cDlg.cls:2456` 的
  `For Each iPrn In Printers`。产物是 `void* _fe_enum_0 = vb6_ForEach_Init(vb6_Printers_Collection());`、
  裸名残留 0 处、rc=0 ⇒ **与 §B104 同形**：发码面对、只有诊断报噪声。所以这一格要的是"语义层放行"，不是新能力。
- **根因比 §B106 更窄**：语义层早就有一张"内置全局对象"的名单 —— `src/semantics/builtin/builtin_funcs.inc`
  里的 `addObj`（Debug/Err/Screen/App/Printer/Forms/Clipboard/Console），发码侧
  `cgen_expr_ident_dispatch.inc` 的内置对象段各自硬编码拦名。**两份名单不一致的那一格就是 `Printers`**
  （发码侧有、`addObj` 里没有）。⇒ 修法 = 补 `addObj("Printers")`，四行（含注释），机制一行没动。
- **但这条不是"只是噪声"**（同一份夹具 `.build/b112_probe` 两台并排跑）：宽松模块里裸 `Printers` 落进
  "隐式 Variant 局部"那一格，产物实打实变了 ——
  BASE = `vb6_VARIANT Printers = vb6_VariantEmpty();  /* 隐式变量 */` 加
  `vb6_ForEach_Init(vb6_VariantToObjectVal(vb6_Printers_Collection()))`；NEW = 两行都没了，ForEach 直接拿集合。
  **语料里恰好没有宽松用法，所以 A/B 的产物零差异是运气的读数，不是"这条无害"的判据** —— 判据要自己造夹具。
- **A/B（全语料 398 输入，BASE = `.build/b112_C3_base.exe` = 未含本刀的那枚 exe）**：
  `artifacts that changed = 0`、`diagnostics ADDED = 0`、**VB3001 24 → 23**，唯一消失的一行是
  `tests/dbgdlg/dbgdlg.vbp` 的 `Printers`；per-id delta 只有 `VB3001` ⇒ 形状门不该动一行。
- **第 47 道哨兵** `scripts/check_builtin_global_objects.ps1`：V1 裸 `Printers`（一份 Option Explicit、一份宽松）
  都不许报；V2 产物两头钉 —— `vb6_Printers_Collection()` 恰好 2 处，且不许出现隐式局部、也不许出现那层
  `vb6_VariantToObjectVal` 包装；V3 反面证人 `bgoNoSuchNameAnywhere` 必须继续报；
  **S1 才是这一刀的结构性抓法**：把发码段拦的每个裸名与 `addObj` 名单对表，缺格就红 ——
  等于把今天这个 bug 的形状改成了一个会自己响的洞。段界用 `vbokonly`（内置常量段的第一个名字）划；
  `me` / `parent` 两条豁免写死了理由（各走自己的出口，不属于"内置全局对象"那一格）。
  S2 钉名单只在 `addObj` 一处、恰好 9 枚，且不许重名。
- **负控三条各自红过**：① `-Exe …\b112_C3_base.exe` ⇒ V1 + 两条 V2 红；② 把 `addObj("Printers")`
  **注释掉** ⇒ S1 报 `printers` 未登记 + S2 报 8 枚。**②第一次红不起来**：纯文本正则把 `// addObj("Printers")`
  也算进了名单 ⇒ 那条"护栏"对"有人注释掉一行"完全无感；两处采集都改成跳过注释行才真的会响（读数规矩：
  静态 census 的正则必须能区分"写着"与"写着但被注释"）。③ 往发码段植一枚假拦截 `if (lower == "bgofakeobj")`
  ⇒ S1 单独红。两处植入都按 md5 原样还原（`0ce682ba…` / `9ddf4ee2…`）。
- **12 处到现在**：本刀清掉 `Printers` 1 处 ⇒ 累计清 2 处（§B107 的 `CTRLINFO_EATS_RETURN` + 本刀）。
  **已量成"发码对、诊断错"的累计 10 处**（窗体伪成员 5 + 裸 `Controls` 2 + 裸 `UserControl` 1 + `TmForm2` 1 + `Printers` 1），
  它们的放行**各撞在不同处**：`Printers` 撞在 `addObj` 少一格（本刀已清）、窗体伪成员那 5 条撞在 backend 那张
  496 行的 `getControlPropReadFn`（§B104）、裸 `Controls` / 裸 `UserControl` 撞在宿主表的 `rtl` 契约（§B105）。
  剩下 `VK_UP` 是夹具漏写 `Private Const`（§B106）。
- **下一条**：§B104 那张表搬家（放行 5 条）与 §B105 的口径 a/b（放行 3 条）**都还压着**，两条都不是补一格名单；
  这两件清完才轮到"解析不出 + Option Explicit ⇒ error"那一刀（§B101）。
### B107 账 #278 第二刀已出 = 跨模块 `Public Const` / `Public Enum` 成员进工程级名字表（第 46 道哨兵，2026-10-09，门 #442 attempt 1 全绿（12/12，含形状门，head b5e94e2d））

- **落地**（三处，全在"名单"这一侧，发码一行没动）：`driver_semantics.cpp` 建表那一趟多认两类 AST
  （`ConstDecl` 取 `Public` 的名字、`EnumDecl` 取 `Public` 的每个 `EnumMember`），存进第三份名单
  `projPubConsts`；`semantic_analyzer.hpp` 一个 setter + 一个成员；`namesProjectLevel` 多问一句
  （**位置无关** —— 裸名位就是常量的合法位置，不像模块名要挑限定符位）。
- **A/B（全语料 398 输入，BASE = 未含本刀的那枚 exe）**：`artifacts that changed = 0`、
  `diagnostics ADDED = 0`、**VB3001 25 → 24**，唯一消失的一行是 `VBFlexGridDemo` 的 `CTRLINFO_EATS_RETURN`；
  per-id delta 只有 `VB3001`（没有别的号跟着动）。⇒ 形状门不该动一行。
- **探针读数**：同一份 `.build/b909_probe` 修前 3 条 VB3001、修后 **0 条**，产物
  `vb6_ret_SumIt = ((42 + 3) - 0);` **逐字节不变** —— 折叠本来就是发码在做的事，本刀只是让语义层别再报错。
- **第 46 道哨兵** `scripts/check_project_const_visibility.ps1`：自己造一份临时工程（不进语料、不动清单）。
  **V1** 正例不许报 VB3001 且产物必须出现"字面量相遇"的折叠；**V2 反面证人** = 同一工程里放一枚真不存在的名
  （`pcvNoSuchNameAnywhere`）**必须继续报 VB3001**（只钉 V1 的哨兵会给"干脆什么都不报"背书 —— 那是同一个洞的另一种坏法）；
  **S1/S2** 钉建造点与消费点各一处（三份名单的 insert 数 2/1/1、三个 setter 各 1 声明+1 调用、语义层消费点各 1）。
  `-Exe` 参数是专门为负控加的。
- **两条负控各自红**：V1 用未含本刀的 exe ⇒ 红在"PCV_ANSWER 仍报 VB3001"；S1 植一枚假 insert ⇒
  红在"3 sites (want exactly 2)"；按 md5 删回原位后双双绿。
- **顺带两条读数**：① `Public Enum` 成员的产物形状 = `typedef enum vb6_enum_X { vb6_enum_X_A = 0, … }` + 用点折成字面量；
  ② 那枚 bogus 名在产物里仍是**裸标识符**（`bogus = vb6_ChkLong(pcvNoSuchNameAnywhere);`）⇒ §B101 那一格没被本刀动过，
  它正是"解析不出 + Option Explicit 将来升 error"的靶子。
- **12 处的现状**：本刀清掉 1 处（`CTRLINFO_EATS_RETURN`）；**已量成"发码对、诊断错"的累计 9 处**
  （窗体伪成员 5 + 裸 `Controls` 2 + 裸 `UserControl` 1 + `TmForm2` 1 —— 最后这枚本轮补量：
  `Unload TmForm2` 发的是 `vb6_UnloadForm(vb6_form_hwnd_TmForm2())`，rc=0，与 §B104 同形）；
  剩下 `VK_UP` = 夹具漏写 `Private Const`（§B106），`Printers` = **还没量**，按 §B104/§B105 那把尺子先定性再动。

### B106 跨模块 `Public Const`：工程级名字表**只收过程、不收常量** —— 第 10 处"发码有路、语义没放行"，而且连标准模块之间都报（2026-10-09 量完，**下一刀**）

探针 `.build/b909_probe`（`KConstants.bas` 两枚 `Public Const` + `KThing.cls` 用两枚 + `KMain.bas` 用一枚）：

- 诊断 **3 条 VB3001**：`.cls` 里两枚，以及 **`.bas` 用另一枚 `.bas` 的 Public Const 也报**（`KMain.bas` 的 `K_SIMPLE`）。
  这不是"类模块才有的麻烦"，是最普通的工程级常量进不了表。
- 产物**对**：发码是 `v = (1 + 42);` —— 语义层早把两枚常量折成字面量了，`KMain` 那侧同样折好。
  ⇒ 与 §B104/§B105 同形：**缺的是放行，不是能力**。

**根因范围就一处**：`driver_semantics.cpp:96-117` 那张工程级表逐条 `switch (d->kind)` 只收 `SubDecl / FunctionDecl / PropertyDecl`
（注释写明"只有标准模块的 Public 过程能被裸名点到"），**模块级 `Public Const` 从来没进表**；
而 `namesProjectLevel`（`semantic_analyzer_util.cpp:722-727`）是语义层唯一的"这名字工程里有"闸门，闸门后面紧跟的就是
`optionExplicit_` 那条 3001（`semantic_analyzer_expr.cpp:159-161`）。
⇒ 落点：同一趟建表里把标准模块的模块级 `Public Const` 收进第三份名单（**要一起考虑 `Public Enum` 成员** —— 问表的时候别只问一半），
`namesProjectLevel` 多问一句。发码不碰：常量折叠那条路今天就在跑。

**于是账 #278 那 12 处里的"常量"那一格就此定位，两枚走两条完全不同的路**：

- `CTRLINFO_EATS_RETURN`（`VBFlexGrid.ctl` 用 `VTableHandle.bas:42` 的 `Public Const`）= **就是这一格**，
  不是"内在常量表缺项"，别去动 #217/#218 那三格机制。⚠ 语料里这一处的**产物看不见**：用它的那个
  `IOleControlVB_GetControlInfo` 属于 `Implements OLEGuids.IOleControlVB`，而那枚接口找不到（VB3044），
  整个子程序没进发码 —— 所以别拿这条 census 当"产物坏了"的证据，也别拿它当"产物好了"的证据。
- `VK_UP`（`tests/tabwalk/WalkForm.frm:345`）= **夹具自己的源码 bug**：同文件 202-205 行手写了
  `WM_KEYDOWN / WM_KEYUP / VK_TAB / VK_DOWN` 四枚 `Private Const`，**独独漏了 `VK_UP`**；
  今天能编过是因为发码里留下的是**裸 `VK_UP`**，撞上了 `winuser.h` 的真宏（emit 里只有 `#define VK_TAB (9)`、
  `#define VK_DOWN (40)` 两枚）。**VB6 本人在这里报 Variable not defined** ⇒ 修法 = 给夹具补
  `Private Const VK_UP As Long = &H26`，与 `Count` 同族：改测试源码，不改编译器。

**⚠ 一条本轮又用了两遍的读法**（§B105 末尾那条的续用）：诊断行号对 `.frm/.pag/.ctl` 是**代码段相对行号，列号才是物理的**。
本轮标定：`.pag` 偏移 234、`TmForm.frm` 偏移 23、`VBFlexGrid.ctl` 偏移 21。三处都是靠列号一比对就对上的；
按行号直接读文件会读到设计期属性块，看起来像"census 是假的"。

### B105 另外 3 处也是「产物对、诊断错」，但这一族的放行撞在表的 `rtl` 契约上（2026-10-09 量完，**要拍一个口径**）

接 §B104 的同一把尺子（先看产物再看诊断）量剩下的裸写文档成员，**三处全是噪声**：

- `.pag` 裸 `Controls` 两条 —— `ppProgressCircular.pag` 的 `Set oPC = Controls.Add(…)` 与 `Controls.Remove "ProgCirc"`，
  发码已经是 `vb6_ComCallObject((void*)vb6_UC_Controls(), L"Add", …)`（产物 b904_charts.c 第 4716 / 4780 行），rc=0。
  ⇒ 语义层那两条 VB3001 与 §B104 同形：**名字归文档对象模型管，发码有路，语义没放行**。

**放行为什么不是「加一行表 + 去掉一个 `!memberObjCtx_`」**（我按这个思路走到第三步就撞上）：

1. 语义层的链在 `semantic_analyzer_expr.cpp:142/146` 按**位置**分了两条：文档隐式**对象名**只在限定符位放行（`memberObjCtx_ &&`），
   文档**裸成员**只在裸位放行（`!memberObjCtx_ &&`）。而 `Controls.Add(…)` 里的 `Controls` 正是站在**限定符位**的裸成员 ⇒ 两条都不接，落到 159 报 VB3001。
2. 表里只有 `{"usercontrol","controls",…HPF_BARE}`，没有 propertypage 那一行 ⇒ 就算放开位置，`.pag` 还是问不到。
3. **给 `.pag` 补那一行 = 在表里说谎**：`HostPseudoRow::rtl` 的含义写死在页头 ——「发成 `vb6_<对象>_<rtl>`」，
   而 `check_host_pseudo_table.ps1` 的 R1 就是拿这条钉每一行（`vb6_PropertyPage_Controls` 在 `vb6rtl_userctl.h` 里没人）。
   真实的答复是一条**专用码头** `vb6_UC_Controls()`（`.ctl`/`.pag` 共用，RTL 内部按当前实例回落 formHwnd），
   不是 `vb6_PropertyPage_Controls` —— 那边 `vb6_UserControl_Controls` 虽然存在却是**恒 NULL 的空桩**，
   `cgen_expr_ident_builtin.inc:217-221` 注释里已经写明"用 NULL 会让枚举得到空集"，所以那条路刻意绕开它。
   ⇒ 再补一枚同名 NULL 全局 = 往 RTL 里放第二个陷阱。**这一族缺的不是行，是表里没有"这一档由专用码头回答"这个概念。**

**要拍的口径（两形，我倾向 b）**：
- **a**：给 RTL 补 `extern void* vb6_PropertyPage_Controls;`（照 `vb6_UserControl_Controls` 那枚空桩的样子）+ 表加一行 + 语义层放开位置。
  便宜，但表里从此有两个恒 NULL 的陷阱名，而它们**永远不该被发码**（发码走 `vb6_UC_Controls()`）。
- **b**：给表加一枚 `HPF_CHANNEL` 旗标 = 「这行的答案是专用码头，`rtl` 字段不代表 `vb6_<对象>_<rtl>`」，
  哨兵的 R1 对带这枚旗标的行**换成另一条判据**（码头的符号名在册、且表里不许出现第三种拼写），语义层按表放行。
  这正是 `canvas_drawing.hpp` 的 `CANVAS_OWNER_METHOD/DRAW` 已经在用的形状 —— 一行标"由哪条码头回答"，两层都只问表。

**⚠ 顺带一条会骗人的读数规矩（本轮自己差点被骗）**：诊断里的行号对 `.frm`/`.pag`（`.ctl` 待核）是**代码段相对行号，不是文件物理行号** ——
`ppProgressCircular.pag(226,15)` 与 `(250,5)` 的实体是物理第 **460** 与 **484** 行（偏移 234；列号是物理列，`Controls` 分别落在第 15 与第 5 列，一比对就对上了）。
我按 226 去读文件时读到的是设计期属性 `Top = 120`，差点据此判"这条 census 是假的"。
⇒ **拿 census 的行号回文件对现场之前，先用列号把偏移标定出来**（列号可信，行号不可信）。这一条与 [[gbk-diagnostic-census-hazard]] 同族。

**于是 12 处的分格又移走 3 处**：诊断面噪声累计 8 处（窗体伪成员 5 + 裸 `Controls` 2 + 裸 `UserControl` 1，后两处待按同一把尺子复量），
剩下**可能真要动产品**的是 `TmForm2`（窗体名自动实例化）1 / `Printers` 1 / 常量 2（`VK_UP`、`CTRLINFO_EATS_RETURN`）——
其中常量那两枚先要问一句 VB6 到底从哪里拿到它们（工程引用里的类型库？还是模块内 `Const`？），别默认"补内在常量表"。



### B103 账 #278 第一刀已出 = Implements 那一族从 VB3001 分家（新号 VB3044，缺槽那条并回既有的 VB3012）+ 第 45 道哨兵（2026-10-09，**已出：门 #441**）

- **落地**：`src/common/diagnostics.hpp` 加 `SemImplementsInterfaceNotFound = 3044`；`src/semantics/semantic_analyzer.cpp:270` 改 3044、`:295` 改 **3012**（`SemInterfaceNotImplemented` 早就在用，两条报的是同一件事）。
  **严重级一条都没动**（两条都还是 warning）—— 这一刀只买「一个号只表示一件事」。为什么必须买：严重级是在**调用点**选的（`diag_.warn` / `diag_.error`），
  而往后按**号**能做的动作有两处会连坐 —— ①「解析不出 + Option Explicit 升 error」若按号翻面，会把 `Implements` 的两条一起判死；
  ② `--suppress-warning <id>` 那条抑制通道（`driver_compile.cpp:46`）按号抑制，今天抑制 3001 就是把接口检查一起闭嘴。
- **A/B**（BASE = `.build/wt_base278` @ `be95c866` 冷编；41 个输入 = 报 VB3001 的 16 个 + 写了 `Implements` 的 25 个）：
  **11 行 `VB3001 → VB3044`、文案一字未动**；`GONE=0` / `ADDED=0`；**41 份 `--emit-c` 产物逐字节相同**；rc 一台不差。⇒ 发码零影响，形状门（398 行清单）不该动一行。
- **⚠ 诚实的覆盖缺口**：`:295` 那条（legacy 缺槽 → VB3012）**在语料里 0 条读数**（`VB3012_total=0`；语料的 `Implements` 全落在「接口找不到」那一档，
  `itf_neg/n08_missing_slot.cls` 走的是新式那条路）。那一格只有静态哨兵的 S4 钉着，**没有行为判据** —— 补行为要新写一枚 legacy 夹具
  （`.cls` 当宿主接口 + 少写一个 `Iface_Member`），本批没做，记在这里当欠账。
- **第 45 道哨兵** `scripts/check_diag_id_exclusivity.ps1`：S1 使用点恰好 **7** 处（逐处点文件名+行号）/ S2 `semantic_analyzer.cpp` 里 **0** 处 /
  S3 新号 **1 声明 + 1 使用** / S4 legacy 缺槽那处用 3012。登记在 `tests/run_tests.ps1` 的 `[STATIC]` 那一趟（`Test-DiagIdExclusivity`），
  `check_static_sentinel_registration` 报 **45 道盘上 / 45 道登记**。
- **负控**（先证它能红）：往 `semantic_analyzer.cpp` 追加一行假借用 ⇒ **S1 报 8 处、S2 点名 `semantic_analyzer.cpp:387`，两条一起红（rc=1）**；
  删回原位（md5 `108bcf13c7404cb4fc79919a5815a0c2` 对上）⇒ 绿。
- **门 #441**（run 37900602129、head `acd6e666`、branch dev、attempt 1）= 12 条 check-run 全 completed/success、非绿 0，
  其中 **`Emit manifest (shape oracle)` 也在内** ⇒ 本机那 41 份「产物逐字节相同」在 CI 那台上同样成立，分家这一刀没动发码一根线。
- **顺带一条工具坑**（写 census 时自己栽的）：A/B 第一版按 `r[2]` 比 ID，而那个五元组是 `(文件, 行, **级别**, 号, 文案)` ⇒ BASE/NEW 两边永远相等，
  41 行全报「没差异」，真差异一条都没显示。**计数器取错列报出来的"绿"比不跑更坏 —— 它把 A/B 变成背书。** 第二版按 `r[3]` 比号、
  再单独用 `(文件,行,文案)` 当键报「同一行换了号」，那 11 行才现形。

### B102 ①②两件的落地：门对 PR 只报不挡已发货；「判死」被 census 判死（不能整条升 error）（2026-10-09）

**② 已出**：`emit-manifest` 那一格加 `continue-on-error: ${{ github.event_name == 'pull_request' }}` ——
**PR 上报而不挡、推 dev 仍硬判红**（注释里写了为什么这格天生会红：任何一次合法的 codegen 改动都会让几十行哈希动，
而那正是要看的读数；红话里现在带着三种成因与一条命令 = §B100）。
**门 #440**（run 数 440、head `0ed4b51c`、branch dev、attempt 1）= 12 条 check-run 全 completed/success、非绿 0
⇒ 改完 workflow 后 dev 这条路照旧产出真判据（`Emit manifest (shape oracle)` 自己也是 success）。
⚠ 一条诚实的边界：#440 是**绿**的一轮，所以它只证明「push 路上这格仍是真判据的一部分」，**没有**证明
「PR 上红会被吞成 success」与「dev 上红会真红」这两半 —— 那两句要等第一次真红才拿得到读数（PR 红 ⇒ 该 job 显示 success
而注解里有 `::error::shape-oracle:`；dev 红 ⇒ 该 job 显示 failure）。若哪天发现 dev 上也把它吞了，就是把
`continue-on-error` 的表达式条件写错了，届时按 §B100 那套自证补一条「push 事件下这格失败必须让 run 红」的读数。

**① 影响面前置 census（`.build/b892_census3001.py`，单台 x64 全语料 398 份，读 stderr）**：
`inputs_with_VB3001=16 / total_VB3001=25 / 其中所在文件写了 Option Explicit = 25（一枚不差）/ 没写 = 0`。
25 处按成因分格（这才是关键，不是总数）：

| 族 | 实物 | 该不该判死 |
|---|---|---|
| **真源码 bug** | `ucProgressCircular.ctl:931 'Count'`（×2 份输入各一次） | 该（VB6 也编不过） |
| 宿主伪成员 / 文档隐式对象 | `ScaleWidth`×3、`ScaleHeight`、`WindowState`、`Printers`、`Controls`×2、`UserControl` | **不该** —— VB6 合法，是我们那张表还缺项（#159 / #174 / #217 第二刀 / #252 同族） |
| 应有而未登记的常量 | `VK_UP`、`CTRLINFO_EATS_RETURN` | **不该**（#218 内在常量同族） |
| 跨模块 Public 过程裸调 | `InitVisualStylesFixes`、`MainForm`、`TmForm2`、`Form2` | **不该** —— VB6 里标准模块 Public 成员本来就能裸调；而窗体名会自动实例化 |
| 故意的负例 / 包可见性夹具 | `nopeHere`、`alsoNope`、`nopeHereIsNotAName`、`Hidden1`、`OpenSecret`×2 | **不该** —— 这些夹具今天的期望就是「warning 一条、其余照跑」 |

⇒ **整条把 `VB3001` 升成 error 会红掉 25 位里至少 19 位合法或故意的用法**，这一刀按这个形状**不能做**。
反向那条也一样死：**整条补隐式声明**（把 no-OE 那支的落地搬到 OE 分支）会把 `InitVisualStylesFixes(...)` / `MainForm` 这种
跨模块**调用**声明成一枚同名 Variant 局部 ⇒ 遮蔽真符号 ⇒ 换一种编不过。
**所以 ① 剩下的唯一活路是先把「语句位」分开**：同一个裸标识符，站在**调用位**（`Foo` / `Foo 1,2` / `Foo(…)` 作语句）与
站在**变量位**（赋值左值、表达式操作数）是两个不同的问题，而今天这两支都从同一个 fallthrough 出去、只在 `optionExplicit_` 上分岔
（`semantic_analyzer_expr.cpp:159-175`）。下一步就是给这一支带上「它在语句里是什么位置」这个信息 —— 那是 #150/#143/#239
那一族（语句路 vs 表达式路）的第四次同问，**该在语义层带上下文，不该在发码侧猜**。
量到的现成家底：登记点 `symbol_table.hpp:546-564`（`implicitVars_`，键 `<mod>`+模块小写+`$`+过程小写）；
发射点三处同形（`cgen_decl_func.cpp:441` / `cgen_decl_proc.cpp:355` / `cgen_decl_prop.cpp:322`）。
另：census 顺带证实 `Count` 那两处只在 Charts 的两份输入里（`Proyecto1.vbp` 与 `ucProgressCircular/Proyecto1.vbp`），
不是全仓蔓延 —— 所以最后那 2 处若要"改源码笔误"也是可控的小改（要用户点头才动别人的 `.ctl`）。

**① 的开工家底（同一轮量到，下一轮不必重新找）**：「把位置信息手递进表达式分析」这手段**仓里已有实物** ——
`semantic_analyzer.hpp:166` 的 `bool memberObjCtx_ = false`，由 `visit(MemberAccessExpr)` 在 `:182-185` save/set/restore，
而 `visit(IdentifierExpr)` 里两处判据正在吃它（`:142` 限定符位 / `:146` 裸写位，两处给不同答案）—— #159 那张表能收成一处靠的就是它。
要加的只是第二枚同族旗标（暂名 `assignTargetCtx_`），置位点已数清：`semantic_analyzer_stmt.cpp:45 / 55 / 67`
（三条语句路各一次 `analyzeExpr(*node.target)`）与 `:192`（另一条带 target 的语句）；而 `:37-38` 已经有
`node.target->kind == ASTNodeKind::IdentifierExpr` 的特判 ⇒ 「左值是裸标识符」在那个位置可判。
⇒ 改动面 = 一枚旗标 + 三到四处 save/restore + `semantic_analyzer_expr.cpp:159` 那支 fallthrough 多问一句「站在赋值目标位吗」；
**只有答「是」的那批升 error**（census 里就是 `Count` 那 2 条），其余（调用位 / 读值位）留 warning。

**⚠ 上面这条图纸自己错了，就地订正（核过源码，不是推测）**：「只把**赋值目标位**升 error」**打不到那 2 条真 bug** ——
`ucProgressCircular.ctl` 全文里 `Count` 只出现一次，而且是 `If hBrush = 0 Or Count = 0 Then Exit Function`，
**站在比较的读值位**（`.build/b888_probe` 里那个 `Count = 7` 是我为了让最小复现走赋值路**自己写的**，不是实物形状）。
⇒ 按「赋值位」切，那枚工程照旧发出不可编译的 C，等于白做一刀。
**真正的分界不是位置，是「这个名字到底能不能解析」**：VB6 在 `Option Explicit` 下拒绝的是**任何**解析不出来的名字，
而我们今天解析不出来的 25 条里，23 条是**我们缺项**（宿主伪成员 / 该登记的常量 / 跨模块 Public 过程 / 包成员 /
自动实例化的窗体名），2 条是**源码自己没有**（`Count`）。所以能做的切法只有一种 ——
**先把那 23 条的解析补齐，再把剩下的「解析不出 + Option Explicit」升 error**（顺序不能反：先升 error 就是把我们的缺项算成别人的 bug）。
现成的顺路证据：`Startup.bas:36` 的 `InitVisualStylesFixes` 明明该被 `namesProjectLevel`（标准模块 Public 过程裸名位）挡掉，
却仍然出了 3001 ⇒ 那枚集合本身漏收，属 #217/#218 那一族的又一格，**它修好之前不要碰严重级**。
下一轮的动作因此改成：①先量「23 条各自缺在哪张表」（`namesProjectLevel` 的收录条件 / `kHostPseudoRows` 的覆盖面 /
窗体名自动实例化那一族）；②补一张是一张，每补一张就顺手钉一枚哨兵；③全部清零后再回来把 `warn` 换 `error`，
那时 census 应该恰好剩 2 条 —— 那才是这一刀的绿照。

**② 的第二个理由（同一轮量到，不是推测）**：`VB3001 未声明` **不是一个报点，是五个**，各自语义还不一样 ——
`semantic_analyzer_expr.cpp:160`（标识符位，warn）、`semantic_analyzer_stmt.cpp:99`（**For 循环变量**，warn，文案是"未声明的变量"）、
`semantic_analyzer.cpp:270 / 295`（隐式变量登记那一段的回声，warn）、`semantic_analyzer_decl.cpp:121 / 239 / 405`
（`Gosub` 到未声明标签，**已经是 error**）、`semantic_analyzer_expr.cpp:632`（error）、`semantic_analyzer_util.cpp:375`（warn）。
⇒ "把 VB3001 升 error" 这句话本身就含糊：真要做，得**先按报点分格**、逐点定严重级，而不是把一个 DiagnosticID 整个翻面
（翻了会把 For 循环变量、跨模块调用这些合法用法一起判死，还会与早已是 error 的那两族混在一起，读数没法归因）。
另记一条没查完的线索（下一轮从这里接）：`Startup.bas:36` 的 `Call InitVisualStylesFixes` 本该被
`namesProjectLevel` 挡掉（`VisualStyles.bas:204` 有 `Public Sub InitVisualStylesFixes`，而 `projPubProcs` 就是扫标准模块
顶层 `Sub/Function/Property` 的 Public 名字建的，`driver_semantics.cpp:87-117`）—— 它却仍出 3001，
**最可疑的是那枚 Sub 落在 `#If VBA7 / #Else` 条件编译块里**（`VisualStyles.bas` 是典型的双分支文件），
即工程级名字表在建表时有没有把条件编译的分支剪掉，决定了这条名字在不在表上。这一格查清之前，23 条里的"跨模块调用"那 4 条没法定案。

**③ 那条线索查清了，并且推翻 census 的一半读法（同一轮，接在上面那条"最可疑是条件编译"之后 —— 那个猜测是错的）**：
- `VisualStyles.bas:204` 的 `Public Sub InitVisualStylesFixes()` **不在任何 `#If` 块里**（该文件的 `#If` 全部在 168 行之前闭合），
  而 `namesProjectLevel`（`semantic_analyzer_util.cpp:722-727`）读的就是 `projPublicProcs_` 那张小写表 ⇒ 条件编译这条**排除**。
- 真相在**输入种类**：那两条报点出自 census 把 `tests/VBFlexGridDemo/Common/Startup.bas` 当**独立输入**跑了一遍
  （`emit_manifest.ps1` 的语料同时收 `.vbp` 与 `.bas`）。单文件编译时工程里没有 `VisualStyles.bas` ⇒
  "未声明的标识符（可能来自其他模块）"**是对的描述**。按整工程重跑 `VBFlexGridDemo.vbp` 实测：`Startup.bas:36/37` 那两条**不见了**，
  只剩 `CTRLINFO_EATS_RETURN`、裸 `UserControl`，以及两条 `Implements: interface ... not found`（那是**另一个语义复用同一个 DiagnosticID**，
  报点在 `semantic_analyzer.cpp:270`）。
- ⇒ **census 那 25 条要按输入种类重读**：standalone `.bas` 那 7 份贡献的 **9 条不该算进"解析缺项"**（把工程拆成单文件编译的自然结果 ——
  `InitVisualStylesFixes`/`MainForm`/`Form2`/`Hidden1`/`OpenSecret`/三条负例夹具都在里面）；**整工程级只剩 16 条**，
  其中源码 bug 仍只有 `Count`×2，其余 14 条才是真缺项，按表分：窗体名自动实例化 1（`TmForm2`）、
  窗体伪成员 5（`ScaleWidth`×3 / `ScaleHeight` / `WindowState`）、全局对象 1（`Printers`）、PropertyPage 裸 `Controls` 2、
  裸 `UserControl` 1、常量 2（`VK_UP` / `CTRLINFO_EATS_RETURN`）、`Implements` 接口找不到 2、包成员 1（`OpenSecret`，那是 pkg_xmod 的**故意**用例）。
- 下一轮因此收窄成两句：**先把 `Implements` 那一族从 `SemUndeclaredIdentifier` 里分家出独立 ID**（否则任何"翻面"都会把接口警告一起判死），
  再逐张补上面那 6 张表；两条都做完、census 只剩 `Count`×2 时，才轮到把"解析不出 + Option Explicit"升 error。

**④ 重测把 ③ 的两处读法一起订正（同一轮；新的消息无关计数器 `.build/b900_census.py`，输出 `.build/b900_census_out.txt`）**：
- **老 census 那条正则要求文案长成 `…: '名字' (` 这个形状** ⇒ 英文文案的 `Implements` 行**一条都没匹配上**。
  所以 ② 表里那 25 条从来**不含** Implements；而 ③ 说「整工程级只剩 16 条」并把 Implements 2 塞进分格 —— **两处都错：16 是「有 VB3001 的输入数」，不是行数**，
  Implements 那几条本来就在 25 之外。教训：按文案形状匹配的 census，换一种文案就静默少一批；**改成按 `级别 + VB号` 解析、文案只当数据**才算 census。
- 重算（NEW exe，398 个输入）：**VB3001 = 25 行 / VB3044 = 5 行 / VB3012 = 0 行**；25 行 = 整工程输入 17 行 + standalone `.bas` 8 行。
  去重（两份 Charts 工程报同一批 `.ctl`/`.pag`、`friend_open_host.bas` 同现两处）后 **21 个唯一源码位置** =
  刻意负例 5（`nopeHere` / `alsoNope` / `nopeHereIsNotAName` / `Hidden1` / `OpenSecret`）+ 拆成单文件编译的自然结果 3
  （`Form2` / `InitVisualStylesFixes` / `MainForm`）+ 源码 bug 1（`ucProgressCircular.ctl:931` 的 `Count`）+ **真缺项 12**。
- 那 12 处按表分（这才是下一刀的对象；③ 写的「6 张表 14 条」里那两个数字都要换掉）：
  窗体名自动实例化 **1**（`TmForm2`）/ 窗体伪成员 **5**（czUI `ScaleWidth`×2、VbQRCodegen `WindowState`+`ScaleWidth`+`ScaleHeight`）/
  全局对象 `Printers` **1** / 裸 `Controls` **2**（`ppProgressCircular.pag:226,250`）/ 裸 `UserControl` **1** / 常量 **2**（`VK_UP`、`CTRLINFO_EATS_RETURN`）。

### B101 `Option Explicit` 在场时只发 warning、发码却把裸名直接发出去 ⇒ 产物必然 C2065（账 #278，2026-10-09，**未开工；① 的严重级那一刀见上一节，已被 census 否掉**）

**⚠ 自纠（同一轮内两次改口，第二次是实测定的）**：本节最初写成「隐式未声明标识符在 `.bas` 落地、在 `.ctl` 不落地 = 两份答案」——
**那是探针自己的混淆**：我把 `Option Explicit` 只写进了 `.ctl` 那份。补跑干净的 2×2（同一枚 C3.exe，`.build/b888_probe/`）：

| 上下文 | 无 `Option Explicit` | 带 `Option Explicit` |
|---|---|---|
| `.bas`（`g_exitfn.bas` / `b_opt.bas`） | `vb6_VARIANT Count = vb6_VariantEmpty(); /* 隐式变量 */` ⇒ 编得过 | **`Count = 7;` 裸名发进 C** ⇒ C2065 |
| `.ctl` UserControl（`ctlimpl/ucBare.ctl` / `ucNoOpt.ctl`） | 同上，补声明 ⇒ 编得过 | 同上，裸名 ⇒ C2065 |

⇒ **只有一份答案，而且它是坏的**：落地逻辑在 `cgen_decl_func.cpp:441` / `cgen_decl_proc.cpp:355` / `cgen_decl_prop.cpp:322` 三处同形，
UC 与标准模块走的是同一批序言发射器（`PushInstance` 就在同一段里）；带 `Option Explicit` 时分析器**不登记**隐式变量（这按 VB6 是对的），
可诊断只到 `VB3001` **warning** 级（rc 仍是 0、`--emit-c` 照出产物），于是「警告一句、产物编不过」。

**真实代价（这就是它值得修的原因）**：Charts 的 `ucProgressCircular` 今天真编译 x86 **只剩这一条错** ——
`ucProgressCircular.c(1504): error C2065: 'Count'`，源码是原作者的笔误（`ucProgressCircular.ctl:949` 的裸 `Count`，全文件仅此一处，
上两行刚 `GdipGetPointCount mPath, lCount`），而那枚 `.ctl` 第 24 行**写着** `Option Explicit` ⇒ VB6 本人对这行也是编译错误。

**要拍的口径（比原来的 A/B 更清楚了）**：
- **判死（与 VB6 一致）**：`Option Explicit` 在场 ⇒ `VB3001 未声明的标识符` 升成 **error**，`--emit-c` 不出产物、`rc!=0`。
  好处：不再产出必然编不过的 C，且语义层就把源码 bug 报出来（报错位置是 `.ctl:949`，比 C2065 好读得多）。
  代价：`ucProgressCircular` 这类「原作者笔误 + Option Explicit」的存量工程**从此明确编不过**，要过那道门得先改语料那行（这不算「为绕开工具 bug 改夹具」—— 工具没坏，是语料坏了）。
- **宽容**：像无 `Option Explicit` 那样补一枚隐式 Variant 局部（今天的行为只在「不带 OE」时成立）⇒ 那工程立刻编得过、可升进门禁；
  代价是比 VB6 宽，且把源码 bug 变成运行期恒 `Empty` 的静默行为（`Count = 0` 为真，那函数直接 `Exit Function`）。
- **本线倾向**：**判死**。理由：今天的中间态（只 warning、却发不可编译的 C）是最坏的一种 —— 它把错误推到下游、还留下 482 行的 cl 日志当现场；
  而「宽容」会让一条真实的源码 bug 在产物里变成静默的错行为。**这条要用户点头才动**（它会给存量语料添红：`--emit-c` 面与 `[VBP]` 门禁都可能新增判死，需全语料逐行归因）。
- 顺带一条独立的小缺口（与口径无关，可以直接修）：无论判死还是宽容，**都不该出现「warning + 不可编译的 C」这个组合**；
  即便选宽容路线，也要先有一条哨兵钉「发码侧引用了任何未被声明的裸标识符 ⇒ 语义层必须已经登记它」——那才是这一族的单一权威问法。

**开工家底**：登记点 `semantic_analyzer_expr.cpp:163` + `symbol_table.hpp:546-564`（`implicitVars_`，键是 `<mod>` + 模块名小写 + `$` + 过程名小写）；
发射点三处同形（`cgen_decl_func.cpp:441` / `cgen_decl_proc.cpp:355` / `cgen_decl_prop.cpp:322`）—— **三处**正是本仓一贯要收成的那一处。

## C. 仍在生效的口径与工具事实（与本文档等长的一半价值在这里；完整版见记忆库）

- **VB.Timer 的节拍口径 = Win32 SetTimer 那一档（系统计时 tick，实测 ~15.6 ms；`Interval` 不足一个 tick 就往上取整，另有一条 `USER_TIMER_MINIMUM=10 ms` 钳位），判据一律不钉绝对拍号**（门 #407 之后定，2026-10-08）：winmm `timeSetEvent` 那条路（`Interval=20` 实得 ~50 拍/秒、`Interval=5` ~199）曾把精度提到 ms 级，代价是它自发出去的 WM_TIMER 是一条**真实待处理消息**、长期占住线程队列 ⇒ 硬件输入被饿死（3DMenu 实测点一下就不动、标题不再随点击变换，而 VB6 编译的同一份代码正常 —— VB6 内部就是 SetTimer）。现在 winmm 只作派发窗无效时的兜底，且带 `posted` 合并。**两头都要活的后果**：`tests/c29timer` 那六条判据从「秒级窗口里的绝对拍数」改成机制（开了要跑 / 改 Interval 两向都重排 / 关掉要停 / 各槽周期互不串 / 小 Interval 到地板为止），名义间隔取 100/200/500 ms 这一档 —— 系统 tick 是 15.6 ms 还是被别的进程 `timeBeginPeriod` 提到 1 ms，读数都落在同一条带里（取整误差 ≤7%）；带里那道上界（T6 `<=150`）是**退回 winmm 的哨兵**（旧口径的 199 会当场红）。**同族提醒**：凡是"在秒级窗口里数拍"的判据都吃这台机器的全局时间精度，写之前先问一句这条读数在 1 ms tick 的机器上是否还成立。
- **哨兵判"这枚函数体里有没有某个调用"之前必须先剥注释，且扫窗不许是定长的**（`check_control_dc.ps1` 在合并后被两样东西同时打红）：
  ① 它的 `[\s\S]{0,900}?\n\}` 是**隐含尺寸假设** —— 别人把 `vb6_GetControlHDC` 长到 ~1350 字符，窗口读不到体尾的 `}` ⇒「体没找到」整条假红。
  修法不是换个更大的数，是**懒配到第一个顶格 `}`**（那正是"这枚函数的体"本身，不含尺寸假设）。
  ② 四条内容判据按字面 `-match 'GetDC\('` —— 而那笔改动在**注释**里写了"旧版在这里 GetDC(过程)" ⇒ 把一句说明读成一次调用。
  修法是判据前先剥 `//…` 与 `/*…*/`。**方向是变严不是变松**：真调用照样红（假改动 C1 证），散文里提名字不再红（C2 证），
  去掉缓存写照样双条红（C3：D3 + D4）。凡是"扫源码文本判结构"的哨兵都吃这两条 —— 定长窗口与不剥注释都是**迟早会红的假红**。

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
| 账 #230（提交 `a9c47c82`，门 #376 = run 37547498187、head `a9c47c82`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | 控件几何的**读法**缺陷：四个 getter 现场把窗口 rect 折算回容器单位 ⇒ 写 5000 读回 4995、设计值 1007 读回 1005（像素网格把每一次读写都量化一遍）。语料 105 份设计文件、770 枚控件里 **238 处**几何值不是 15 的倍数，且**没有一条存量针量过**（夹具读的都是自己写进去的 15 的倍数）⇒ 编译 / 链接 / A/B 三头全哑，红只能自己造。 | 一处存储 `vb6_GeomCacheRead/Write`（值 + 写它时容器的档位；属性名一档一个；值存 2v+1 避开 `SetPropW(…,0)`=删属性）。两道闸都问一句、不列清单：档位换过 ⇒ 投影；按存的数换算的像素与窗口现在的像素不符 ⇒ 投影（**别人**挪过窗口，ComboBox 建窗自加高就是这一格）。写侧三来路齐：setter×4 / `vb6_ControlMove`（按 mask 逐档，#193 只验过 Left/Top）/ `vb6_CreateControl`（设计值，**恒按缇** —— czUI frmDemo 同时写 `ClientWidth=6600` 与 `ScaleWidth=440`、子控件 `Width=6240` 只有按缇才放得下，所以"把建窗换算改成容器 ScaleMode"这个原本的计划被实测推翻）。 | 判据 = 新夹具 `tests/geomcache` 两架构真跑 8 条全 True，GC03/04/05 各带一枚**像素证人**（user32 rect + kernel32 MulDiv，刻意不走产品自己的换算）钉住没挪窗口；负控 = 改前那台跑同一份夹具 ⇒ 五条 False、三条两边同数。新哨兵 `check_ctrl_geom_cache.ps1`（第 26 道）六条各用一处假改动证红。护栏 = A/B（BASE = 改前本机冷编那台）inputs=90 **changed=0 / same=90** + 10 枚会读几何的存量夹具两架构输出逐行相同（唯一差异 `modal` 的 `busy=0→8` 是在途 tick 条数，账 #162 同族且套件只钉前缀）。顺带新立 §B75（UC 宿主模型那份第二实现没接缓存）与 §B76（浮点→整数赋值截断、`CLng` 是 half-up 而 VB6 是 banker's） |
| 账 #248（提交 `09621c38`，门 #380 = run 37556525156、head `09621c38`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | 浮点交给整数目标时按 C **截断**，而同一句写成 `CLng(...)` 是 `round()` ⇒ 一个决定两个答案（`l = 6.73` 交 6、`l = 7 / 2` 交 3 对 `CLng(7 / 2)` 交 4）。暴露面 = 90 份 emit 里 **36 处**把浮点结果存进窄整型（21 处在 Charts 2020 主工程）⇒ 真实排版算式一直差一格。 | RTL 新增唯一出口 `vb6_FltToLng(double) = (int64_t)round(x)`，`vb6_CInt` / `vb6_CLng` 改问它；发码侧 `narrowCheckAssign` 在 `装得下，不套` 那道短路**之前**加一道浮点闸（Single / Double / Currency）。 | 判据 = `tests/test_f2lng.bas` 13 条两架构真跑（负控 = BASE 冷编到**本笔的父提交** `5fac77b1` 跑同一份夹具 ⇒ 4/13）+ 发码针 `f2lng_round`（四真四禁）+ 新哨兵 `check_float_to_int_round.ps1`（第 27 道）。护栏 = A/B（同一个 BASE）captures=90 / changed=18 / **未归因 0 行**（K1 224 + K2 16）+ 29 道 [STATIC] 零红。取数口径立了两条：判据要选**截断与舍入永远不同**的那一形（`.75`），`.5` 那一族只钉不变式不钉答案（平局口径未拍，见 §B77 ①）；A/B 的 BASE **必须是本笔父提交**，否则别人的刀会冒充我的行。剩余两半 = §B77 |
| 账 #247（提交 `a22cb110`，门 #385 = run 37570038967、head `a22cb110`、branch dev、attempt 1= 11 job 全 completed/success、非绿 0） | **控件几何的第二份实现在 UC 宿主模型里，撤掉了**：晚绑定的一句話（`With obj : .Width = …`、UserControl 里的 `Parent.Width`）走 IDispatch ⇒ `vb6_Host_GetProp` / `vb6_Host_SetProp`，而这两支自带一份几何 —— 读侧自己量窗口矩形 + 写死缇，写侧把四档一起读成缇再整体 `MoveWindow`（所以写 `.Left` 会顺带把 Top/Width/Height 重量化一遍），两头都不碰 #230 那张缓存 ⇒ 同一枚窗体的同一个属性两个答案（实测 `f.Width = 7222` 两条路都读回 **7215**；反过来 `Me.Width = 6011` 之后宿主那路读回 **6015**）。**改** = 四档读写都转调 `vb6_Get/SetControl(Left|Top|Width|Height)`（单位跟着**所在容器**的 ScaleMode = #175 的口径，写进去的数进同一张缓存），各档写各档；被断调用点的 `vb6_ho_ctrlRect` / `vb6_ho_isForm` 一起删。**刻意没并**：`ScaleWidth` / `ScaleHeight` 的另一份权威问的是窗口**自身**档位，并错了就把容器口径搬进自身口径，留给 §B79 一起量。**语料暴露面** = 每架构 60 处（flex 31 / Charts 21 / czUI 4 / listview 3 / toolbar 1）；Charts 那一处源形是 `With CtrlNames(i) : .Left = FW * Rects(i-1).Left / 100`（`ClsResizer.cls:138`，唯一调用点 `Form2.frm:581`），而它今天**因 §B79 一条也不执行** ⇒ 这一刀在 Charts 上是「接上但没有流量」，真流量是 czUI 的 `With Parent`；窗体那一档两条路等价是有根据的（表单 `CreateWindowExW` 传 `hWndParent = NULL` ⇒ `GetParent` 给 0 ⇒ `ScreenToClient` 不改点、`vb6_GetScaleMode(NULL)` 答 1=缇）。**判据** = `tests/geomcache` 升 GC09..GC12（晚写晚读 / 晚写直读 / 像素证人 / 直写晚读，两架构真跑 NEW 12/12 True；**BASE = 本笔父提交 `571ca6aa` 本机冷编那台跑同一份夹具** ⇒ 只有 GC09/GC10/GC12 三条 False，GC01..GC08 与 GC11 逐行同数，x64 与 x86 两片都是这三条）+ 哨兵 S7（四档两头各问出口恰好一次、`MoveWindow(` 与自带投影助手回潮 = 0），S7 五处各用一处假改动证过能红、植完按 md5 还原，全套 30 道 [STATIC] 零红。**顺手量到三格独立的存量缺陷**（改前改后两台同形，已分别立账 §B78 窗体自身 Left 写后读 AV / §B79 窗体 Controls 与 Count 答空 / §B80 `vb6_UC_ParentMove` 把缇当像素），本刀一条没碰。|
| 账 #250（提交 `2883fcbc`，门 #389 = run 37581465287、head `9936cf3f`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | 窗体的 Controls 集合交不出来：宿主模型那一档问的是 `vb6_ho_setVariantDispatch`（那一枚**刻意**只交 Empty，好让字体代理继续走 Nothing 分支、别让 oleaut32 去 Release 一个裸指针），rev14 为「真交对象」另立了 `vb6_ho_setVariantObject`，却只把 `Item` / `Add` 两个调用点改过去 ⇒ Controls 交回 NULL，而下游三条（Count / Item / For Each，`vb6com_foreach.c` 里早已备好集合分支）一条也进不去。实测 `Me.Controls.Count`=0、`For Each` 零条、`Item(i).Name` 空、`.Width` 0（一律空值、不崩）⇒ Charts 的 ClsResizer 在唯一调用点上静默不做任何事。 | Controls 那一档换 `vb6_ho_setVariantObject`；**Font 代理那一档保持 Empty 原样**（它依赖 Nothing 分支，与 #247 的并表方向相反，是有意的）。 | 判据 = GC13/GC14（集合交得出对象 + 枚举真走得完）；负控 = 本笔父提交 `a22cb110` 冷编那台跑同一份夹具 ⇒ 只有这两条 False（raw=0 → 5），其余十四行逐字节相同。成员名那一半另立 §B81（账 #252）：枚举出来 5 条是真的，`o.Name` 仍空、`TypeName` 答 Control。 | 已过：门 #389。**这一刀带出的下一格 = §B82（账 #254）** —— 接通之后 ClsResizer 走的第一件事就撞上 `Form.Count` 硬填 0 ⇒ Charts 启动期错误 9，门 #387/#388 两片因此红在 `Tests (vbp #3)`。 |
| 账 #254（提交 `9936cf3f`，门 #389 = run 37581465287、head `9936cf3f`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | 宿主模型里窗体**自身**的 `Count` 档一直硬填 0（`uc_hostmodel_getprop.inc:76`）。#250 把 Controls 集合接通之后 Charts 2020 的 ClsResizer 第一次有流量，而它 `ReDim Rects(oForm.Count - 1)` 铺出来的格子是空的，`For Each` 第一格写 `Rects(0)` ⇒ 运行期错误 9（下标越界）⇒ 启动期整进程退出。CI 把它报成 `[GUI] Charts2020 ... Main window not available within 5s`：runner 的 5s 观察圈对**进程早退**抛同一句话（Fix 189 只补了 AutoExitSec 那一路）⇒ 本地直接跑产物才看得见真读数。 | 那一档转调新增的 `vb6_UC_ControlsCountOf(obj)`（`uc_controls.c`，只包一层 `vb6_uc_collectChildren`）⇒ 与 `Controls.Count` 同一处收集，不开第二份。 | 判据 GC15 **两头钉**（`fc = n` 且范围 `4..8`）—— 只写等式的话在 #250 之前那台反而 True（两个数都 0）。负控 = 本笔父提交那台 ⇒ 只有 GC15 False（raw=0/5），GC01..GC14 逐行同数；真工程证人 = 同一份夹具两台对跑：改前 123ms 错误 9 退、改后 131ms 出窗且活着、无 crash 文件。护栏 = emit A/B inputs=90 changed=0 / same=90（纯 RTL 刀）+ 30 道 [STATIC] 零红。 | 已过：门 #389（attempt 1，11 job 全绿；#387/#388 那两片红到这一笔一起结掉）。 |
| 账 #249（提交 `84d55a7e`，门 #390 = run 37583390188、head `84d55a7e`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | 本节原记「窗体自身 Left/Top 写后再读就 AV」—— 实测**前提错**：`Debug.Print "x=" & Me.Left` 作为**第一条**语句就崩，而 `v = Me.Left` 赋给 Long 两架构正常。分家的轴是「属性名撞不撞返回 String 的 VB 内置函数」：`inferExprType` 那条「对象位是不是窗体控件」的权威（#229 收成处的 `ctrlTypeOfMemberObject`）只认 `控件名.属性` 与 `控件名(i).属性`，而 **Me 是 MeExpr ⇒ 答「不是」** ⇒ 成员名掉到按裸名查模块符号 ⇒ 命中内置 `Left`（String）⇒ `wrapToBSTR` 不套转换 ⇒ `vb6_BSTR_Concat(L"L", vb6_GetControlLeft(hwnd))` 把 int32_t 当 BSTR 解引用。窗体名那一形本来就通（`cgen_form_ctrl_registry.inc:10`），缺的只有 MeExpr。 | 只教那一条权威认 MeExpr（窗体模块内 ⇒ `FrmControlType::Form`），类型仍只出自 `controlPropType` 那张表（第一档 left/top/width/height=Long 本就是为这对撞名设的）。`.ctl` 里 `Me.` 那一形不在闸口内、未实测。 | 判据三头：夹具 GC16 那一行**就是崩溃现场**（BASE 上整行不出现、`GC-DONE` 跟着没）+ 值一起钉；发码针 `gc_emitc_meleft_wrap`（套 `vb6_CStrLong` 的整条语句必须在）/ `gc_emitc_meleft_raw`（父提交那条裸形必须不在）。负控 = 本笔父提交 `9936cf3f` 冷编那台 ⇒ 15 行、rc=0xC0000005、GC01..GC15 全 True；改后 x64 `True/L390`、x86 `True/L1170`。护栏 = emit A/B inputs=90 **changed=0 / same=90** ⇒ 存量零暴露，这一格只能靠新夹具守。工具坑：PSParser 在实参 `@(` 续行里按字面数括号，单引号串里不配平的括号 ⇒ 整份 run_tests.ps1 ParserError **而退出码照旧 0**，故两条针都用整条语句。 | 已过：门 #390（run 37583390188、head `84d55a7e`、branch dev、attempt 1）= 11 job 全 completed/success、非绿 0 |
| 账 #253（提交 `ad33e452` = 第一半；第二半随账 #249 那笔。第一半的门 = #389 = run 37581465287、head `9936cf3f`、attempt 1 = 11 job 全 completed/success、非绿 0） | modal 的跳格判据在 CI 上按负载抖：门 #387 的 CI 工件与 #386（绿）实质差在 `MW2=cmdX → txtSecond`、`MW-new=4/repeat=txtMain/hops=3 → repeat=txtSecond/hops=1`，同一片 `busy=6 → 17`。机理 = #162 / #213 那一族（在途 WM_TIMER 连着排空 + VK_TAB 由泵异步消化）⇒ 两拍挤在同一瞬间读到「还没动」，第一次回头被提前判定；本地两台编译器都稳定给 4/3 ⇒ 判据与拍序耦合，不是产品回退。 | 第一半：`tMain_Timer` 相位推进前加一道真实间隔闸（`Timer` 差 < 0.03s 的拍不计相位，仍计入 `gStray`）。第二半：夹具 + runner 里那条「`C3_OCX_NO_DLGMSG=1` ⇒ `MW-new` 回到 1」的负控文案实测**已失效**（`vb6forms.c:884`：#163 之后 `vb6_TabNavKey` 排在 `IsDialogMessageW` 之前，VK_TAB 已被自研导航器吃掉），两处注释改成实测口径 —— 本相的开关侧负控只剩 `C3_OCX_NO_TABNAV` 一条（它仍会红：`MW-seq` 翻成 z-order `txtMain,cmdY,cmdX,txtSecond`）。 | 判据 = 套件原有的 `MW-new=4/repeat=txtMain/hops=3` 与 `MW-seq`（本地 x64/x86 两条都对），门 #389 的 ModalApp 工件确认 CI 上回到同一条读数；开关侧 A/B 现场跑过（NO_TABNAV 红、NO_DLGMSG 无变化）。 | 已过：门 #389（第一半）；第二半随账 #249 那笔过门 #390 = run 37583390188、head `84d55a7e`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0 |
| 账 #255（提交 `863d3698`，门 #391 = run 37594343602、head `c6c6de2d`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | 门 #390 之后 Charts 2020 起得了窗但排版算式不落地：x86 启动期 168 条 `vb6_ComSetProp/GetProp: property "…" not found`。两刀：① `With 集合(i)` 的接收者是那次默认 Item 调用返回的 calloc VARIANT 的**地址**（With 那条分岔只管「表达式是 VARIANT 值」，漏了「VARIANT* 指针」这一档；Set 那一路早做了同一件事，做法是就地手改字符串前缀 = 第二份抄本）；② 宿主包装器的「包装器→原对象」表只追加不让，`free` 掉的块地址被下一枚复用、扫表从第 0 格起 ⇒ 枚举出来的每枚成员都解回**第一枚**（实测三枚 hWnd 同数、Left 一律 120 = Timer 设计值、写 `.Width` 全落在 Timer 上）。 | ① 收成唯一出口 `comObjectRefFromCallExpr`，With 三条接收者分支 + Set 那段抄本一起转调；② `HW_Release` 归零时先 `vb6_UC_DropWrapPair` 再 `free`。 | 判据 = `tests/geomcache` 升 GC17（成员互不相同的一份宽度总和）+ GC18/GC19（经集合的 With 写落地，两头：控件自报 + 窗口像素），两架构真跑 19/19 True；负控 = 父提交 `84d55a7e` 冷编那台 ⇒ 只有这三条 False（4006/7918、三档纹丝不动），其余十四行同数。护栏 = emit A/B inputs=90 **changed=2 / same=88**（只 Charts 那份的两行，变的就是那层 `vb6_ComUnpackObject(...)`）+ 两枚发码针（形在 / 裸形不在，x64 与 x86 都核）+ 24 道 [STATIC] 零红。 | 已过：门 #391 = run 37594343602、head `c6c6de2d`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0。CI 侧两枚发码针按名各自 PASS（`[EMITC-SHAPE] gc_emitc_with_member_unpack` 在 vbp #3、`[EMITC-ABSENT] gc_emitc_with_member_raw` 在 vbp #4），夹具两架构也 PASS，工件里 GC17..GC19 的读数与本地逐字相同（7918/7918/5、1234×4、82,82,82）。现场读数（不作判据）：Charts x86 stderr 由 168 条 → 0，窗口 141ms 出现、进程存活。顺带立的新账 = §B85（账 #256：控件名当对象用被折成 `.Text`，再进 VT_DISPATCH 就是野 Release）；名字面那一半仍在 §B81（账 #252）。 |
| 账 #252（提交 `9ee3d0ae`，门 #392 = run 37605281964、head `9ee3d0ae`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | 控件的 VB 身份从没登记：`vb6_ucHo[]` 带 `name`/`typeName`/`index` 三格、四条读法（晚绑定 `.Name`、`.Index`、`vb6_Host_TypeNameOf`、按名查找）本来就只问这张表，可全仓只有**窗体**登记过（`vb6rtl_system.c` 那一处，name 还传 NULL）⇒ 标准控件走 `IsWindow` 兜底：名字一律空串、类型一律 Control、按名查找一律 NULL，**直读 `txtA.Name` 也空**。Charts 2020 是真流量：`ClsResizer` 靠 `TypeName(CtrlNames(i)) = FBuf(j).CtrlTypeName` 挑字体档，恒不等 ⇒ #255 把接收者修对之后仍一条不落地。 | 两条创建路各发一次既有出口 `vb6_HostObj_Register`（顶层 `dsHwnd` / 容器子控件 `ctrlHwndExprForInit`），读侧一字未改；`VB6_UC_MAX_OBJ` 128→512（表是进程级的，实测 VBFlexGridDemo 182 枚、Charts 164 枚，128 会静默装不下）。 | 判据 GC20..GC23（两头钉 + 「非数组答 -1」专拦填 0 + 按名找回），两架构 23/23 True；负控 = 父提交那台四条全 False；护栏 = A/B added 490 逐行归因（486 登记 + #255 的 8/4，OTHER=0）+ 两路各一条发码针 + 24 道 [STATIC] 零红。 | 已过：门 #392 = run 37605281964、head `9ee3d0ae`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0。CI 复核：夹具两架构工件里 GC20..GC23 四条全 True 且与本地同数（hWnd 两头相等），发码针 `[EMITC-SHAPE] gc_emitc_identity_both_create_routes` 在 vbp #1 的日志里按名 PASS；本笔唯一的行为风险（登记表对标准控件非空 ⇒ `collectChildren` 第一趟不再为空，枚举次序可能由 z-order 变创建次序）由 tabwalk 现场排除：CI 工件 `TW-seq=cmdIn1,cmdIn2,cmdDeep,cmdTop2,cmdInPic,cmdTop1,` 与期望逐字相同。。刻意没碰：成员读 `.Tag` 仍空（晚绑定字符串读法，#88 同族）、直读 `TypeName(txtA)` 仍答 String（实参折成默认属性 = §B85/#256）。风险：登记表从此对标准控件非空 ⇒ `collectChildren` 第一趟不再为空，枚举次序可能由 z-order 变创建次序（本夹具判据次序无关）。 |
| 账 #257（提交 `3de665fe` + 台账 `18ac1c65`，门 #394 = run 37615366430、head `18ac1c65`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0（wall 10m59s）） | 宿主模型的 getprop 长串里 `Container` 与 `Parent` **一档都没有**，而窗体登记走 `vb6_Forms_Register(void* hwnd)`（签名里根本没有名字，恒传 NULL）⇒ `oCtrl.Container` 交回空值、`空.Name` 交回空串，Charts 排版器 `SaveControlsPositions` 第一句 `If oCtrl.Container.Name = oForm.Name Then` 是**「空 = 空」的真空通过**（探针模式 P 实测 gate=True）。真流量还包括 `tests/BalloonTooltips/cTT.cls` 的 `objControl.Container.hWnd`。 | ① 那一档新增，交回 `GetParent(接收者)` 走 #250 的唯一出口 `vb6_ho_setVariantObject`（无父 ⇒ 空值），口径与 `vb6_uc_collectChildren` 的成员判据同一条关系，不开第二份；② `vb6_Forms_Register` 加 name 形参，发码那一处交窗体模块名；③ `vb6_HostObj_Register` 的「同 hwnd 已登记就 return」改成**只补还空着的 name/typeName**（isForm/index 不动，`.Index` 恒 -1 是存量答案）。 | 判据 GC24..GC28（两架构 28/28 True）：GC24「无第三种答案」+ 范围，GC25 两边非空且等于模块名，GC26 证人只来自窗口（Caption + hWnd == Me.hwnd），GC27 **反面**（picP 的成员属于 picP ⇒ 同一句必须 False），GC28 窗体自己无容器（拦「原样退回接收者」那种修法）。负控 = 父提交 `fcae9af6` 冷编那台跑同一份夹具 ⇒ GC24 raw=0+5+0、GC25 raw=[]/[]、GC27 raw=[lblP]/[]/[]，GC01..GC23 同数。护栏 = A/B inputs=136 **changed=61 / added=71 / removed=71 / OTHER=0**（全是 Forms_Register 那一行），口径再放宽到 `.vbp` + `.bas` × 两架构 = 782 份捕获复跑：**changed=122 / added=142 / removed=142 / unattributed=0**+ 两头发码针（带名形在 / 无名形不在）+ 29 道 [STATIC] 零红。 | 已过：门 #394 = run 37615366430、head `18ac1c65`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0（wall 10m59s）。CI 复核：`Tests (vbp #3)` 与 `(vbp #4)` 两片工件里的 `GCCache.out` 各 28 行、标签逐字相同，GC24 raw=5+0+0/5、GC25 [GCForm]/[GCForm]、GC26 [GeomCache]/<同一枚 hWnd 两头相等>、GC27 [lblP]/[picP]/[GCForm]、GC28 [Empty]，**无一条 False**，两片都走到 GC-DONE。两条发码针的按名读数不在工件里（工件只带 exe 的 .out/.err，runner 控制台不在其中）⇒ 本机已把两头各自核过（带名形命中 1、无名旧形残留 0），CI 侧由那两片的 job 全绿背书。开工前那个「无窗口的 Timer 应当无容器」的假设被实测推翻（五枚成员 5+0+0，Timer 也答窗体），故 GC28 只作反面证人。残余：MDI 工程里 `Form.Parent` 应答 MDIForm，本刀按「顶层无父 = 无容器」交空值，语料 MDI 0 处 ⇒ 未实测，另记账。 | 
| 账 #256（提交 `799f0072` + 台账 `3ebbc641`，门 #396 = run 37632990355、head `3ebbc641`、attempt 1 = 11 job 全 completed/success、非绿 0） | 裸控件名出现在**晚绑定调用的实参位**上时被折成默认属性读数（`vb6_GetControlText` 那枚 BSTR / `vb6_GetCheckValue` 那个 int），而装箱那层只听 `comPackExpr` 的已知变量型 —— 它对任何控件名都答 Object ⇒ 发出 `vb6_ComPackObject(vb6_GetControlText(...))`：串进 VT_DISPATCH，语句收尾的栈上 VARIANT 走 `VariantClear` 对 VT_DISPATCH **无条件** `Release` ⇒ 按串头几字节解 vtable，两架构都 0xC0000005、崩点就在那条 Add 上。轴不是「对象位 vs 值位」（VB6 对 Variant 形参取默认属性恰恰是对的），是**装箱档与值类型对不上**（#122 / #123 / #88 同族）。 | 收成一条出口 `ctrlDefaultPropOf(lower)`：折不折 / 折成哪一枚 / 读函数是谁 / 折出来什么型只此一处回答，标识符发码那两段旧判定与打包档同时改问它；型另开 `defaultPropType(ctrlType)` 与 `getDefaultPropertyName` 逐行对着写（不动 `controlPropType` 那张管显式属性读法的表）；类型分支收进 `comPackFnForVbType`，兜底一字未改。Picture 与 Set 那一面刻意不接管。 | 判据 GC29..GC31（为该格补一枚 CheckBox，两架构真跑 31/31 True，每格都回读集合里那一枚而不是只问活没活）+ 一条**证人**：`For Each` 出来的 Object 变量照旧按对象打包；负控 = 父提交 `262d7eab` 冷编那台（GC01..GC28 全 True 而 GC29 那行不出现、EXIT=0xC0000005）。护栏 = emit A/B inputs=391 / captures=782 / same=780 / changed=2（只有本夹具，逐行归因 unattributed=0，同时兜住那两处零行为改动的重构）+ 两头发码针（三新形在 / 三旧形不在）+ 哨兵 A6/A7（用假改动证过能红）+ 29 道 [STATIC] 全绿。 | 已过（CI 两片 GCCache.out 各 31 行无 False，三行读数与本地逐字相同）。新开残余 = §B87（账 #258：`Set <泛对象槽> = <控件名>` 交出的是默认属性值而不是控件本身，错值不崩）。
| 账 #259（提交 `990cb883` + 台账 `473c6dfc`，门 #397 = run 37642453080、head `473c6dfc`、attempt 1 = 11 job 全 completed/success、非绿 0） | 门 **#395**（head `262d7eab` = 纯文档笔，代码与全绿的 #394 逐字相同）唯一红 `[VBP] combofocus_x86 … run timeout: 60s (cpu=62ms, alive) last='COMBOFOCUS-DONE' lines=7` —— 七行判据全打完、无窗口、进程不退。根因不在 #257：**一个线程只有一份 WM_QUIT**，而 `vb6_DoEvents` 用 `PeekMessage(PM_REMOVE)` 把它摘走当普通消息 Dispatch 掉；饿机器上 tick 6 的 DoEvents 抽出已排队的 tick 7，tick 7 卸载最后一个窗体并投出 quit，同一个 DoEvents 再把它吃掉 ⇒ 外层 `GetMessage` 永远等不到。本地 5 枚烧 CPU 的并发下 13/14 挂，不饿时 1/20 —— 这就是「CI 才红本地绿」的全部原因。 | `vb6_RePostQuitIfTaken`：抽到 quit 就原样投回并停泵。两处装机 = DoEvents 循环第一道闸 + `vb6_ShowForm` 模态循环收尾（那条循环体里原本 "收到 WM_QUIT" 的 trace 是死码：GetMessage 取到 quit 返回 0，循环直接结束）。外层 `vb6_MessageLoop` 与 MDI 那条 `vb6_MDIMessageLoop` 是消费者，刻意不投。 | 新夹具 `tests/appqquit` 把形状做硬（每拍 Sleep 80 / Interval 40 ⇒ DoEvents 起手时 WM_TIMER 已在队列 —— 该形状靠的是同一枚 Timer 在自己泵里重进，90fb8069 按 VB6 把它改成不可重入之后夹具已过时，门 #403 唯一真红就是它；形状已挪到第二枚计时器 t2，判据一字未改，见 §B88 末条）：`Q-DONE` + `Q-ORDER=unload-inside-doevents` 钉形状，「进程自己退」钉本体（run timeout 即 FAIL）；两架构 0/10 挂 EXIT=0x0，负控 = 父提交那台 5/5 挂，combofocus_x86 由 13/14 挂变 0/15 挂；新哨兵 P1..P5（假改动证过能红）+ 30 道 [STATIC] 全绿。 | 已过（CI 四片 115 份夹具输出每枚 .out 都配着 .err ⇒ 零超时；CbApp.out 两架构 7 行齐、QApp.out 两架构 2 行齐）。**订正**：#256 那一轮把这条红记成「本地两台 EXIT=0 ⇒ 不是产品确定性缺陷，按 CI 负载时序一族查」—— 那句错了，产品确实有一个确定性洞（负载只是把它放大到每次都中）。 |
| 账 #251（提交 `3d1dbac4` + 台账收线一笔，门 #400 = run 37652210198、head `d199adc0`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | `vb6_UC_ParentMove` 把四个实参原样交给 `MoveWindow`，而这四个数是 **VB 侧量纲**（容器是窗体 = 缇）：czUI 的全屏/恢复存的就是宿主模型读回的那四档（#247 之后全是缇），发码又把 `Parent.Move` 折成这一支 ⇒ 607 缇落在 607 像素上；而且裸 `MoveWindow` 不填 #230 那张缓存，Move 之后读回的是投影 | 转调唯一收口 `vb6_ControlMove(fw, L,T,W,H, 15)`（单位 / 坐标空间 / 缓存 / Resize 只在那一处答）；本单元不 include `vb6forms_prop.h`，沿用它自己的局部 extern 纪律；头文件映射注释补量纲一句 | 判据 = `tests/ve_units` 新增 `U-PMOVE-RAW` + `U-PMOVE`（UC 补 `MoveParent` 走真正的 `Parent.Move` 那一形；两头各钉：窗口对 `MulDiv(缇,DPI,1440)` 的差值为 0（证人 = user32 + kernel32，不经产品换算；钉差值不钉像素绝对数 ⇒ 换 DPI 不漂）+ 四档 VB 侧读回 = 写进去的数；607/451/2407/1811 都不是 15 的倍数 ⇒ 两头互相冒充不了；两架构真跑 True）；负控 = 父提交 `ad16bc9c` 冷编那台跑同一份夹具 ⇒ `l=9105 t=6765 w=23340 h=14700 dx=567 dy=421 dw=1396 dh=859` + False，其余 26 行两台同数 | 已过：门 #400 十一 job 全绿；CI 复核 = `test-logs-vbp-1` 与 `test-logs-vbp-2` 两片工件里的 `VeUnits.out` 各 29 行、**无一条 False**、`U-DONE` 齐，且 `U-PMOVE-RAW l=607 t=451 w=2407 h=1811 dx=0 dy=0 dw=0 dh=0` + `U-PMOVE=True` 两行与本地两台逐字相同（差值为 0 这条判据在 CI 的 DPI 上同样成立）。纯 RTL 刀：A/B same=784 / changed=0 / 未归因=0 只证 cgen 零改动，真伤由那两行钉。哨兵 S8 三处假改动各证过能红；30 道 [STATIC] 零红；语料真流量 2 处都在 czUI 的 `ToggleFullScreen` 里（要人手触发 ⇒ 门上看不见）。刻意没接：`Parent.Move l, t`（少实参）实测 C2198 编不过、语料 0 处，另立。 |
| 账 #260（提交 `d604121a` + 台账 `aeb883a9`，门 #401） | `With <对象> : .Prop = <别的对象>.Prop` 发成 `vb6_ComSetProp(w, L"Prop", vb6_ComPackValue(<别的对象>))` —— With 那一支是**唯一没消费 COM 标记**的写侧出口，成员读取整段不见；而那张「packer → 解封类型」的表本来就抄了三份，With 一支也没问。第二症状（A/B 量到）：漏下的标记被下一条语句消费 ⇒ Charts 六份 UC 里 32 条 `PropertyChanged` 被改写成「往字体对象乱写」 | 新增 `CCodeGen::comMarkerValueForWrite(packFn, valExpr)` 一处（无标记原样返回；有标记按 packer 反推解封类型 → 问 `resolveComValue`），三处旧抄本撤掉、With 那一支补上 ⇒ 写侧五个出口共用一张表；host_pseudo 的默认档 Variant 跟表走 Long（与 `resolveComMarkerForPack` 对 Boolean 的答案一致） | 判据 = 新夹具 `tests/comwith`（`Scripting.Dictionary` 晚绑定、不依赖类型库）：三种接收者各一条 + 算术形不二次解封 + 两条反面证人；两架构 6/6 True。负控 = 父提交 `2b86a45f` 冷编那台 ⇒ 4 条 False（其中 CW-DIRECT/CW-EXPR 是**被这一刀连带修好的**，两台不同数）。发码针两头（三新形在 / 三旧形不在）。护栏 = A/B captures=786 / changed=14 / 未归因 0 / run-failures=0（2 份是本夹具，12 份 = Charts 六工程 × 两架构；解开 266 / 假写消失 266 / PropertyChanged 回来 32）+ 新哨兵 C1..C3 三处假改动各证过能红 + 31 道 [STATIC] 零红 | 已出（门 #401 = run 37660544554、head `aeb883a9`、attempt 1、11 job 非绿 0；`WithCopy.out` 两架构各 13 行与本地逐字相同；与门 #400 的工件对跑：104 份相同 + 15 份只抖在册形状 + 2 份新增）|
| 账 #258（提交 `79d70956` + 台账收线一笔，门 #402） | `Set <泛对象槽> = <控件名>` 交出的是**默认属性读数**而不是控件：这一问当时有两份答案（环境闸 `suppressDefaultProp_` + P16 事后在发好的文本里截 `vb6_hwnd_`），两份各盖一半形状 ⇒ 泛对象槽存 BSTR、Variant 槽存 vt=3，而那条写只在 stderr 留一句 `property not found`、`Err.Number` 照旧 0。收成一处 `emitSetObjectRhs`（五个右值出口全转调，闸只开在整枚标识符上）+ `ctrlObjectRefExpr`（句柄拼法与折叠那支同一条出口，撤掉手拼与 P16） | 已出（门 #402 = run 37666864389、head `98347f53`、attempt 1、11 job 非绿 0、wall ≈10m19s；`SetObj.out` 两片各 13 行与本地逐字相同、`SetObj.err` 两片 0 字节；与门 #401 的工件对跑 = 105 份逐字节相同 + 16 份只抖在册形状 + 2 份新增，判据 token 一条没换）。**§B73（账 #245）没随这一刀塌**：同头重测 changed 仍 14、差异行全在「typed ↔ generic」那一族、OTHER=0，方向钉成「CI 走 typelib 签名那一档、本机走 unpackType 那一档」，而两台二进制 `--dump-symbols` 逐字相同 ⇒ 下一格查 `knownTypedComVars_` 六个登记点的首命中是不是随遍历序变（这条门抓不到，Build 与 Tests 同工件自洽）。 |
| 账 #261（提交 `73570cf9` + 台账收线一笔，门 #406 = run 37694019059、head `c4175e64`、attempt 1 = 11 job 全 completed/success、非绿 0） | 浮点交给整数目标这一件事，**目标侧当时也只答一半**：`narrowCheckAssign` 第一道门写着 `target->kind != IdentifierExpr ⇒ 原样返回`，于是成员写入与数组元素写入整片不经取整也不经越界闸 —— 实测六形（`r.Left = 6.73` / `arr(0) = (4*1+3)/4` / `r.Top = r.Left + 0.73` / 模块级 `m_r.Width = 6.73` / `dArr(1) = 6.73` / `ur(0).Left = 6.73`）两架构全按 C 截断（6/1/6/6/6/6），而同一句写成 `CLng(...)` 是 round() ⇒ 一个决定两个答案，#248 那一族的另一半（当时只数了裸标识符的 36 处）。 | 新增唯一出口 `CCodeGen::narrowTargetTypeOf(Expr*)`：「这枚左值是哪档窄整型」只在这里答，只许 Byte / Integer / Long，认不出（Array / UDT 整体 / String / Object / Variant / Boolean）一律 Unknown ⇒ 与改前同形。成员那一档问既有出口 `inferUdtFieldVb6Type`，数组那一档问 `arrayElemTypes_`，裸标识符那串登记表**逐字搬过去**。两条闸都是量出来的：① **空下标 = 整体数组赋值**（Fix 170 的 `dst() = src()`）右边是数组描述符指针 ⇒ 漏闸实测 `test_array.bas` EXIT=0x00000006、四条 `wa-*` 整片不打印（语料里同形还有 11 处/架构，含 VBFlexGrid `Common.bas` 的 `B() = Text` 那条 Fix 170 原始崩溃点）；② **成员数组的整体赋值**（`With x : .Data = baData`）—— `inferUdtFieldVb6Type` 头注释写着「含 Array 标志」而**实测从来不带**（数组性记在 `mi.isArrayDynamic` / `mi.arraySize`，`mi.type` 存的是元素档）⇒ 档位答 Byte ⇒ 发出 `vb6_ChkByte(指针)` ⇒ VbQRCodegen 的 Project1 **BUILD-RC=0 而启动期 Unhandled VB6 Error #6**（BASE 同一份源不报）。收法 = 同一条 udtMembers 走查里用 `outIsArray` 把数组性带回（不开第二张表）+ 订正那句假注释。 | 判据 = `tests/test_f2lng.bas` 从 13 头扩到 **30 头**（F2L14..22 六形新覆盖 + F2L23 成员侧「隐式==显式」不变式【平局方向照 §B77① 不钉答案】+ F2L24..26 不许动的证人 + F2L27/29/30 那两条闸 + **F2L28 刻意留的边界钉成 6，§B90 落地那天必须翻成 7**）；负控 = 本笔父提交冷编那台跑同一份夹具 ⇒ **F2L14..23 十条两架构全 False**（raw=6,7 / 1,2 / 6,7 / 6,8 / 2,3 / 6,7 / 6,7 / 3,4 / 6,7 / 0,-1），改后 30/30 全 True（x64+x86，EXIT=0x0）。发码针 `f2lng_round` 两头扩到 10 必在 / 5 条 BASE 裸形必不在 / 3 条「拆闸才会出现」的形状两侧都必不在（单引号串里括号一律配平，见 §B78 那条 PSParser 坑）。护栏 = 哨兵 `check_float_to_int_round.ps1` 加 S7（定义/声明/被问各恰好一次 + 必须问那两条既有出口 + 必须拒 `fldIsArray` 与 `positional.empty()` + 不许答 Boolean）与 S8（旧闸门计数必须 0 + 无别的消费者），**八处假改动各处证过能红**、每次按 md5 逐字节还原；语料 emit A/B 对本笔父提交那台（judge v2：把所有 `vb6_Chk*(…)` 与 `vb6_FltToLng(…)` 配对拆掉后两侧必须**逐字节相同**）⇒ inputs=394 / captures=788 / 逐字节相同=752 / 只多套检查=36 / **未归因 0** / run-failures=0；新增包裹合计 ChkLong 1894 / ChkInt 528 / ChkByte 440（两架构），其中穿取整出口 1346 ⇒ **每架构 ≈673 处浮点→窄整存储从此 round**，比 #248 那 36 处大一圈——闸门把整片成员/数组目标关在判据面外，A/B 与门都看不见。 | 已过 = 门 **#404**（run 37694019059、head `c4175e64`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0，wall ≈10m31s；这一扇门的 head 里除了本刀还有账 #262/#263 那笔夹具重形状的续笔 `c4175e64` ⇒ #261 的代码与它同过一次门）。**CI 侧复阅**：bas 两片里 `test_f2lng.out` (job14 x64) 与 `test_f2lng_x86.out` (job16 x86) 各 **30 行、False 零条**、`F2L-DONE` 齐；vbp 两片 `QApp.out` 各两行（`Q-ORDER=unload-inside-doevents` + `Q-DONE`）、配对的 `.err` 全 0 字节 ⇒ 一片都没被 60s 超时杀。另：本轮 #403（head `c62fcf61`）那两片红的归因见 §B88 末条 —— 红的是appqquit 那份夹具的形状（被 90fb8069 的 Timer 不可重入改掉），与本刀无关。**同批把这里记的那条边界收了**：账 #262 落地时 F2L28 从「故意钉 6」翻成 `28-member-array-elem-rounded`=7，夹具扩到 32 头（新增 F2L31/32 = 多维成员的两个相邻格）。**风险记账**：这 36 份捕获从此多一道越界闸（整型源也一样）—— 值真越界时 VB6 本人报 6、我们以前静默截断；若门红在 Charts / VBFlexGrid 的像素或计数读数上，先按这条方向归因（#260 那格「以前没画上→画上」同形），别调宽判据。剩下的边界 = §B90（成员数组的元素与二维成员数组，F2L28 已钉）与 §B77①（平局口径等拍板）。 |
| 账 #262（提交 `d84dab1f` + 合并 `b334858c`，门 #409 = run 37712639451、head `b334858c`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | UDT 成员的**多维**定长数组：parser 把第 2..N 维「解析并丢弃」，所以 `M(3, 3) As Long` 只发出 `int32_t M[4]`（16 字节而不是 64），而 `m.M(0, 1)` / `m.M(0, 2)` / `m.M(0, 3)` 三句**全发成 `m.M[0]`** ⇒ 三个格子一格内存，不响不崩静默错值。真流量 = Charts 2020 LabelPlus 的 GDI+ 色彩矩阵 `M(0 To 4, 0 To 4) As Single`（GDI+ 从 20 字节的结构体里读 100 字节，而阴影透明度与图片透明度写同一格）；语料全仓**多维成员声明 1 处**、它的 38 处用法全部给两个实参（arity=2 = 38/38）。与 #211（过程内多维 ND）分家的轴是「成员 × 多维」，与 #261（一维成员数组）分家的轴是维数。 | AST `TypeMember` 加 `moreDims`（`arraySize` 恒是第一维上界 ⇒ 一档一源，现有消费者照旧）+ `arrayRank()` 就地算；克隆那趟一起带上（泛型实例化 clone TypeDecl，漏了就悄悄还原）；语义层**只传秩** （`mi.arrayRank = 1 + moreDims.size()`，各维格数留在发码侧，抄第二份就会分家）；`cgen_decl` 每一维过**同一条** `tryEvalConstInt` 折成 `T name[n0][n1]`；下标新增唯一出口 `CCodeGen::udtFixedMemberIndex`（层数只从秩来、步长归 C 算 ⇒ 没有第二张「每行几格」的表；缺的那层补 `[0]`，多给的照发让 C 报错），`obj.M(i,j)` 与 `With : .M(i,j)` 两条路都改问它。同批把 ByRef 数组实参那处的 `sizeof((arg)[0])` 改成按秩剥到底（秩 1 逐字不变）。顺带收了 #261 那条边界：`narrowTargetTypeOf` 的数组档现在也接 callee 是成员访问/With 成员，但**必须问出数组性**才答档位。 | 判据 = 新夹具 `tests/test_memdim.bas`（每条读数写的都是**第一维相同、第二维不同**的格子 = 正是旧折法会并成一格的那一族；`MD-LEN=64/64/16` 钉布局、`MD-ONE=9` 是一维不被带坏的反面证人、`MD-ALIAS=N` 钉不变式）两架构 11 行齐 + EXIT=0x0 + 两边逐字相同；BASE 两台负控分两份源（含 String 成员那份**编不过** = #263 挡在前面；摘掉 String 成员那份答 `33/33/33`、`222/222`、`3/3`、`LEN=16/16` ⇒ 两格读数互不掩盖）。发码针 `memdim_layout` 12 必在 / 8 必不在；护栏 = 新哨兵 `check_udt_member_dims.ps1`（第 33 道 [STATIC]，U1..U8，**八处假改动各处证过能红**、每次按字节原样还原）；语料 emit A/B 对 `C3_base262.exe`。 | 门待回填。**风险记账**：Charts 那枚色彩矩阵从此**真的按 5×5 传给 GDI+**（以前 20 字节里读 100），而 ByRef 那处的逐字不变是靠 A/B 实测出来的（第一版我把它写成 `sizeof(x)[0]` —— 那是给 sizeof 的结果取下标，非法 C，A/B 那 4 行 dbgdlg 差异就是这么被抓出来的）；门若红在 Charts 的颜色/透明度读数上，先按「以前静默错值现在对了」这个方向归因，别调宽判据。多维成员整体交给 ByRef 数组形参那一形**语料 0 处、没实测**，别读成已支持。 |
| 账 #263（提交 `d84dab1f` + 合并 `b334858c`，门 #409 = run 37712639451、head `b334858c`、branch dev、attempt 1 = 11 job 全 completed/success、非绿 0） | 含所有权元素（String / Variant / 子 UDT）的**定长数组成员**，其自动深拷贝助手把循环头发成 `{ int32_t _n = (int32_t)(sizeof(d->S) / sizeof(d->S[_i])); int32_t _i;` —— `_i` 用在**声明它自己那句话之前** ⇒ C2065，工程编不过（实物：`.build/p262/m265.c` 两条 C2065，一维 `S(3) As String` 与二维 `V(1,1) As String` 都中招）。这是量 #262 时撞出来的，不是找出来的；语料普查 0 处 ⇒ 它比 #262 更早存在却从没响过。 | 逐格改走**扁平元素指针**：`{ ET* _dp = (ET*)&d->X; const ET* _sp = (const ET*)&s->X; int32_t _i; int32_t _n = sizeof(d->X)/sizeof(*_dp);`（ET 由那一档本就分好的三路给：BSTR / vb6_VARIANT / vb6_type_X；认不出就**不动这一格**）。顺手带掉第二条更晚的病：旧写法即便能编，在 `BSTR S[2][3]` 上 `sizeof/sizeof([0])` 数的是**行数**，会少拷四格。 | `tests/test_memdim.bas` 的 `MD-STR=ab/cd` + `MD-STR-COPY=ab/cd`（赋值后才见深拷贝真的逐格 `SysAllocString`）；发码针两头（扁平指针行必须在、`sizeof(d->S[_i])` 与 `d->S[_i]` 必须不在）；哨兵 U6 + 负控 N6。BASE 那台跑同一份夹具 ⇒ BUILD-RC=1、一条 C2065(`_i`)。 | 门待回填。**风险记账**：语料 0 处 ⇒ 门上只该看见新夹具；若看见「以前编不过的工程现在编得过」，那是这一条。旁边没量的一格：元素档是 `As <项目类>` / `As Object` 的定长成员数组走的是 memcpy 快路（两边共享引用 ⇒ 一边 Clear 就悬），这一刀没动它，别读成已修。 |
| 账 #213（提交 `e0b62a8f` + 入账 `43e18c8e` + 登记 `db824476`，门 #491（源码那一跑 = run 38078861543、head `43e18c8e`、attempt 1 = 12 条里 10 绿 + 两红都归完）+ 门 #493（登记那一跑 = run 38080158814、head `db824476`、attempt 1 = 12 条全 completed/success、非绿 0、wall 11m04s；其中 `Emit manifest (shape oracle)` 这一条绿就是「登记的那一行正是 CI 那台自己交的那一份」的独立复算） | tabwalk 相 3 那三步方向键**每步只留一拍**去读 `PostMessage` 交给泵异步消化的结果 ⇒ 在途 `WM_TIMER` 连着排空时那一拍看见「还没动」，`down=` 打成 optA。这一格在门 #336 真红过一次，当时按「`--emit-c` 产物三份逐字节相同」归成夹具节拍余量（§B48 末段）、#337 复绿，**却没人回来改夹具**；与 #162 / 门 #332（pbsb）/ #253（modal）是同一件事的四个读数。本机两台编译器都量不出症状 （改前那台跑今天这份夹具：两架构 × 空载/2×超订 各 8 趟 = 32/32 命中，墙钟 0.9s→2.4s）⇒ 立论是机理与负控，不是复现。 | 相 3 三步改**四步**（最后一枚键也等自己那一拍，`up=` 因此不再由固定拍号决定）；每步先问 `WhereIs(GetFocus()) = gAkLast`，没动就 `gAkWait + 1` 再 `Exit Sub`（**不重发**，每拍重发会跳站），等满 `AK_WAIT_MAX = 40` 拍照常读并 `gAkLate + 1` ⇒ 真丢键仍然读错仍然红，只是不再看队列这一刻排空几条；最坏 8s，仍在 60s 跑步预算内。**判据一个字没改**，只新加 `TW-SETTLE late=0`；`maxwait` 只打不钉。 | 两架构 × 空载/2×超订 各 8 趟 = 32/32 全针命中、`late=0`、`maxwait=0`、`TW-ticks=12`（旧 11）；负控 = `.build/b694_negctl` 那份拷贝里关掉三步的 post ⇒ x64 `late=2/maxwait=40`、x86 `late=3`，`AK-*` 与 `TW-SETTLE` 两条针都缺，而进程 8s 自己退（`rc=0`）。第 24 道哨兵 `check_fixture_timer_close.ps1` 长出 W1..W4：发过 post 的 `*_Timer` 逐枚登记（4 枚）、各自的 settle 口径必须在被保护的读取**之前**、豁免那枚必须仍是「post = 最后一条语句」、每步一自增 + 保险丝 5..100、证人打了必须有人读 —— 七发假改动各红各的（`.build/b697_neg_report.txt`），baseline rc=0。护栏：51 道 [STATIC] 全 rc=0；形状清单 inputs=400 / 未变 399 / 移动 1（正是 tabwalk 那一行）。 | 源码那一跑 = 门 #491：`Build`/`asm`/`bas ×2`/`compile`/`syntax`/`vbp #2`/`vbp #3`/`vbp #4`/`smoke` 十条绿，两条红都归完 —— 形状门红的正是 tabwalk 那一行（CI 那台自己交的清单 = `ci/emit-manifest` 的 `a4738567`，登记前本机复算 400 份 = 399 同 / 1 异、sha256 逐字相同），另一条红是 `ctrldlg_probe` 第二次撞 60s 跑步预算（那片工件 `DlApp.out` 停在 DL10、`.err` 不存在 = 超时现场；#490 同一枚是 15 行齐）⇒ 与本刀无关，另立 §B140 去办（本刀不改 `src/`，除 tabwalk 两片与那一行清单之外不该有别的读数）。CI 两片 tabwalk 都读 `late=0/maxwait=0`、`TW-ticks=12`、期望串一字未变。 |
| 账 #304（提交 `3f0b1823` + 入账 `b306ce9e`，门 #495 全绿，回填见 `19e80c12`） | 「这一枚用例该给几秒」以前只有全局那一个数，而 GUI 里真弹模态框的那枚天生吃墙钟余量：门 #484 与门 #491 的 `Tests (vbp #1)` 两次红同一形状 —— 那片工件里 `DlApp.out` 停在 `DL10`、`DlApp.err` **不存在**（超时那一支只在 stderr 非空才写 .err，崩溃那支两枚都写）；#490 / #484 attempt 2 同一枚是 15 行齐 ⇒ 判据一条没少，缺的是余量。本机空载 1.2~2.1s、8 枚忙等压 4 vCPU 时 §B133 量到 13.8 / 31.4 / 34.5s（×11 的放大），预算却是全局那一个 60。 | 收成一处：`$RunBudgetSec`（具名表，今天 `'DlApp' = 150` 一枚）+ `Get-RunBudgetSec`（唯一出口，表里没有回落全局）；串行路问 `$budgetSec`、并行路进 runspace 前逐枚盖 `$it.BudgetMs`。**全局 60s 未动、`DL11..DL14` 未减**；键取产物基名（与 §B166 那一处同源），两条路同数。 | 配对读数两头：全局 0s + 表有键 ⇒ 解析 150、`ctrldlg_probe` 3/3 PASS；同一个全局、键拼错 ⇒ 解析成回落值并被杀，3/3 现场停在 `DL10`/`DL13` 且 `alive`（正是门上那次的形状）。第 52 道哨兵 B1..B6 六条规则各被一发假改动打红、baseline rc=0（`.build/b911_negctl_report.txt`）；52 道 [STATIC] 全 rc=0。 | 门 #495（run 38085195428、head `b306ce9e`、attempt 1 = 12 条全 completed/success、非绿 0）；`Emit manifest (shape oracle)` 同绿 = 400 行一字未动（= 「本刀不改 `src/`」的预期读数），`Tests (vbp #1)`~`#4` 四条绿 —— 但那一格的真判据是本机的配对读数（表里有键 3/3 过、键拼错 3/3 被杀），不是门的一次绿（这两枚夹具的红本来就按负载抖）。 |
| 账 #305（提交 `55e075ae` 源码 + `adc4b61f` 登记，门 #501 源码 / 门 #503（run 38090841379、head adc4b61f、attempt 1）= 12/12 全绿、非绿 0（含这道形状门本身）） | §B138 那一格由别人的分支出了（内置答案改 LongPtr），但它尾巴上那两格没被那一改带走。第一格量完 = RTL 的 `switch (v.vt)` 十一遍实现里**四处漏答 VT_I8**：x64 上 `TypeName(VarPtr)` 报 "Variant"、`IsNumeric` 报假、`vb6_ComPackVariant` 落 default **把地址静默打包成 VT_EMPTY**、UC 宿主两枚取数读回 0（同一枚值 x86 全部答对） | 四处各补 `case VT_I8`（档名用 SDK 的 `VT_I8`，与已答的五处同一写法）+ `c3rtl.rc` 按约定 touch （`check_rtl_embedded` 拿产物字节核，125/125 相同）；**不新造权威**：把「哪些表必须答这一档」写成一张闭合名单 | 判据 = §B137 那枚装箱夹具加五条针 + **新增 x86 双胞胎**（同一份 `$boxNeedles` 两架构逐字相同，真跑 28/28 命中）；改前读数（同一支探针）x64 = `typename=Variant / isnumeric=False`；第 53 道哨兵 V1..V5 五发假改动各红各的、baseline rc=0；本刀不改发码 ⇒ 形状门预期 400 行全同 —— **这句不中**：实际红 8 行，但其中 7 行是本刀之前别人那笔 §B138b（`e589b1a3`）移动的（上一轮登记 `db824476` 早于那一笔，期望清单本来就没吃过它的形状），第 8 行才是本刀自己那枚夹具的源文本；按 `ci/emit-manifest` 逐 head 二分定的案，登记仍按 CI 读数走（§B100） | 源码跑 = **门 #501**（run 38088742545、head `e51fe008`、attempt 1）= 12 条里 11 绿 + 唯一红是这道形状门（红的就是那 8 行，11 片测试全绿，含新加的 `test_variant_boxing_x86` 第一次进门禁）；登记跑 = **门 #503**（run 38090841379、head `adc4b61f`、attempt 1）= 12 条全 completed/success、非绿 0 —— 其中 `Emit manifest (shape oracle)` 这一跑就是绿的 ⇒ 登记的那 8 行与 CI 那台自己的复算对齐 |
| 账 #305（第三十三刀 = §B141 两档，提交 `9a6b7a56` 源码 + `59e72197` 登记，门 #505（run 38094605541、head 9a6b7a56、attempt 1） 源码那一跑 / 门 #506（run 38095497701、head 59e72197、attempt 1） 登记那一跑 = 12/12 全绿、非绿 0） | `CDec(...)` 交出来的 `VT_DECIMAL` 在「往外读」这一侧只补了两张表：`vb6_TypeName` ⇒ "Decimal"、`vb6_IsNumeric` ⇒ 真；判据两头只钉这两条（两台 30/30，CI 补档前那台两条都缺 = 负控），`CStr`/`Format` 不进针面。**Format 那一臂实测补上仍不对 ⇒ 回退，连同其余 8 处 vt 表另立 §B142**；同一批读数把「写侧坏了」这个推论收回（`CLng`/`CDbl`/`CInt` 的 0 来自 `vb6_VariantToDouble` 的 `default`，不是 DECIMAL 本体） |
| 账 #305（第三十四刀 = §B142 第一格，提交 `f2a21b77` 源码 + `45d31ebb` 登记，门 #510（run 38101654250、head `f2a21b77`、attempt 1）源码那一跑 / 门 #511（run 38102478203、head `45d31ebb`）登记那一跑） | Decimal 这一档「往外读」第一次有了一枚换算出口，同时把写侧少写的那枚标签补上。两处根因都是**格式**不是数值：① `DECIMAL.wReserved` 是 oleaut32 的 vt 标签（手填字段那两路留 0 ⇒ `VariantChangeType` 返回 S_OK 却交 0 / 空串）；② x86 的 VARIANT 装 `VT_DECIMAL` 不走指针，整枚 16 字节内联盖住 VARIANT（只填 `pdecVal` 时 oleaut32 把那枚地址当数值答出来） | 标签走唯一宏 `vb6_DEC_VT`；换算走 `vb6_DecimalToBSTR` / `vb6_DecimalToR8` 两枚（本机只够得着 `VariantChangeType`）；接线三处 = Format 分诊表那一格 + `vb6_VariantToDouble` 的 Decimal 档 + `CStr`（本就转调 Format）|两台各 **34/34 针全中、零 miss**（针面名单从 `run_tests.ps1` 现读现对）；`VB-dec-fmt` 照打不钉（区域设置）。形状门那一跑红的 **6 行**逐段归因：1 行本刀夹具 + 5 行是 PR #11 合并欠下的登记（按 CI 那份逐 head 二分定的案）。下一刀 = 剩下 7 处 vt 读表 + x86 那两条 COM 导入臂（把内联 `lo64` 当指针读，本刀新案） |
| 账 #305（第三十五刀第一格 = §B142 标签收成一个写者，提交 `ec6f0c3d` 源码 + `b1430c9a` 登记，门 #512（run 38105635521、head `ec6f0c3d`）源码那一跑 / 门 #513（run 38106446965、head `b1430c9a`）登记那一跑） | 纯 C 探针量到 `VarDecFromR8` 也不写 `wReserved`（只有 `VarDecFromStr` 写 14）⇒ 第三十四刀只补了手填两臂，**x64 的浮点臂仍交空串 / 交 0**，而 x86 被读侧内联铺法救回 = 台不对称 | 标签改成唯一出口 `vb6_DecSetTag`，`vb6_CDec` 末尾调一次，两行散写撤掉；读侧 x64 也补这一枚与 x86 同形 | 针面从 6 到 9 条，两台各 37/37 零 miss；`CLng(CDec(2.5))` 刻意不钉（§B77① 未拍，探针实测 oleaut32 是 banker's）；形状门红得只有自己那 1 行、399 行零动；顺带把 §B143 的可达性量完（赋值/形参/条件/算术/比较五条路径全都答 0/False，只有 String 目标对）|
| 账 #305（第三十五刀第二格 = §B143 四条数值口 + 覆盖面哨兵换口径，提交 `f2ab0d38` 源码 + `d8ccff59` 登记，门 #515（run 38109731485、head `f2ab0d38`、attempt 1）源码那一跑 = 11 绿 + 唯一红是形状门 / 门 516（run 38110792863、head `d8ccff59`）登记那一跑） | Decimal 在 `vb6_VariantToLong`/`ToLongPtr`/`ToBool` + 比较取数口 `vb6_VarToDouble_internal` 四张表里都没档 ⇒ default 交 0；七条路两台各答 0/False（`CDec(5)+1`=1、`(CDec(5)=5)`=False）。**而那张比较取数口从来不在 census 里** —— 站点正则不认带强转的 `switch ((vb6_vartype)v->vt)` | 四格各补一档，政策照**本表对 double 的既有答案**（round / 非零即真 / 原样交 double），取数一律只问唯一出口 `vb6_DecimalToR8`；哨兵站点正则认强转(11→13)、名单按档分两份(I8 11+2 / DEC 7+6)、新增 D3「豁免理由不许短于 12 字」 | 针面 37→43，两台各 43/43 零 miss；负控两头：剥两格重建那台跑同一份夹具 ⇒ 恰好 4 条 MISS(落点正是被剥那两张表)，四条新判据各用一处假改动证能红；只改 RTL 不动发码第四次由 CI 复算证实(399/400 行逐字节未动)；盲区也量了：接收者不叫 `v` 的 vt 表还有 7 处，其中 2 处正是下一格那两条 COM 导入臂 |
