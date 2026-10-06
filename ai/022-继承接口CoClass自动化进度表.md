# 022-继承接口CoClass 自动化进度总表

> 本文件是每小时自动化任务（"实现继承、接口与CoClass"）的**唯一状态源**。
> 每次运行开始先读本文件，结束前必须更新本文件（状态头 + 批次清单 + 运行日志）。
> 规范输入: `ai/讨论记录/018-接口继承与CoClass设计思路.md`（含 tB 文档要点与分阶段设计思路全文）。

OPEN_LOCAL_DIVERGENCE: **已收口（2026-09-27 入夜 = C29-MV-d）** —— 本地 `ctrlmonthview` 的 MV24/MV25/MV30 三条红既不是"本地环境/增量缓存"、也不是"OS 版本相关"，而是**挂钟相关**：comctl v6 把当前时间混进 `MCM_GETCURSEL` 的负载里（日没错、零头 0.765），`CLng` 下午进位、上午舍掉 ⇒ 本机下午一直红、CI 跑在上午一直绿。修法与全部读数见 029 §九 C29-MV-d 那一格。留下的一般式：**日期面判据按 `CLng` / `=` 比之前，产品读数必须归成整数日**，否则同一棵树的判据随时段翻脸，而门只会读到"上午那一半"。
STATUS: IDLE             # NOT_STARTED | DESIGN | BUSY | IDLE | ALL_DONE
LAST_RUN: 2026-10-06T16:17:00+08:00   # 本轮追加：**门 #364 绿 = 账 #239 收线**（run 37429327451、head `e5c66e4d`、branch dev、attempt 1、11 job 全 completed/success、非绿 0、wall 12m40s），**紧接第二格 = 账 #238 已出（提交待回填，等门）= `vb6_VarCmp*(&A,&B)` 那四个取址点从此问一处判据** —— 以前"这个名字能不能当 `vb6_VARIANT*` 交出去"按**名字形状**答（裸标识符就 &），于是 `Dim d As Double` 的地址被按 VARIANT 的布局读：实测**相等的两个数答 False**（VC01..05 / VC10 / VC15 七条，两架构一致）、而且读过头 = 静默给错答案那一族（#123 的 Byte 是同一形状的另一档）。新判据 `CCodeGen::cmpOperandMayTakeAddr(名字, AST)` 只回答裸名字那一形，答案全部来自既有的 `isDefinitelyVariantExpr`（声明那几张表 + 符号表）；不许取址就走 `_Generic vb6_VariantFromValue` —— 对已是 VARIANT 的表达式它命中 `vb6_VariantIdentity` 恒等直传，**所以敢把默认档翻过来的依据是"最坏只多一枚临时"**。判据四件套：新夹具 `tests/test_varcmp_scalar.bas`（16 条两架构真跑，每条相等配一条近似不等，Long/String 两条本来通的腿一起钉）+ BASE 同一份夹具 7 条 False + 发码针 `varcmp_scalar_boxed`（必须出现装箱、五条旧形状必须不出现）+ 新哨兵 `check_variant_cmp_boxing.ps1`（第 25 道 [STATIC]，三条假改动各证红；其中"判据不再问权威"那一枚假改动**同时**让夹具回红 = 一条假改动喂两个面）。**本账的用后归还**：`tests/fdraw` 的 FD13/FD14 从 `Abs(a-b)<0.001` 换回**直接相等**（当初就是因为这一格才写成算术的）。语料 A/B `inputs=90 changed=2`，逐行归因 = 装箱语句 + 被改写的比较 + `_vcmp_N` 编号平移，**未归因 0**；唯一真形状变化落在 `ucProgressCircular` 的裸名 `Count`（那枚名字本来就 C2065 ⇒ 工程编不过 ⇒ **能编过的语料零暴露**）。邻域 fdraw / pcline 零 False，bool_display / datelit 两架构输出与 BASE 逐行相同。**方法上的订正（写进 §B68）**：上一轮那句"暴露面 0 处"只统计了在同函数找得到声明的名字 ⇒ 漏掉**未声明裸名**那一形，把 1 读成了 0；量暴露面必须把"没有声明的名字"单独列一类。**下一格 = #232**（窗体上裸写 `PSet (x,y)` / `Line -(x,y)` 两形在 parser 就没出口，#224 的发码针仍欠）。等用户口径的仍是 #219 裸名 `Count` 与 #230/#206 的单位口径。
# 上一轮（2026-10-06T15:10:00+08:00）= # 本轮追加：**账 #239 已出（提交待回填，等门）= 控件那户 Print/Cls 撤掉那份自存的像素笔位** —— `vb6_ControlPrint` / `vb6_ControlCls` 改成两条只转调家族 `vb6_Form_Print` / `vb6_Form_Cls` 的码头，`VB6_PrintX` / `VB6_PrintY` 在 src/rtl 里归零，全仓 Print/Cls 一份实现。同一份判据的前后实测（两架构一致）：Print 之后笔位推进 `0→13`（= 同一枚控件自己答的 TextHeight）、笔位放 60 时墨的落点 `第 2 行→笔位那一行`、Cls 之后 CurrentY `400→0`、缇档推进 `0→195`。**第二格同时是防回归**：家族那份 Cls 以前自己按 `VB6_BackColor` 判空 ⇒ 黑色（存进窗口属性就是 NULL）读成「没设过」回落按钮面，改问带 Fix 187 哨兵那份唯一出口；不换的话合并之后控件那户 Cls 会从实测 0 变成 15790320。判据 = `tests/pcline` 加一枚 picP + 五条 needle（PL08-PEN/PL09-CLSPEN/PL10-TWIPADV/PL11-BLACKCLS/PL12-STACK + 两条 RAW）+ BASE 那台跑同一份夹具真红；三处行为负控各让自己那条红（推进归零→08/10/12、像素当用户单位存回→只 10、色彩自己答→只 11），哨兵 `check_form_draw_state.ps1` 新 S8 四条假针各证红 —— **其中一条是「判据只看体内文本会被注释顶成假绿」**：Cls 那问的注释写着权威名而代码答的是属性名，改成只扫代码行才真的红。`check_control_dc.ps1` 的 D2 因这一刀 6→4（哨兵先响、归因写进规则注释）；24 份 check 全绿 + 邻域 fdraw/pbsub/dcsurf 两架构零 needle 缺失；**零后端改动 ⇒ 发码逐字节不可能变**（RTL 是 exe 的资源），所以这一族的红只能落在真跑那一头。另量到一格属 #232：`Me.Cls` 在 VB→C 那一路没出口，发成 `vb6_ComCall(hwnd, L"Cls", NULL, 0)` 的运行期 no-op ⇒ 家族那份 Cls 今天只有控件到达（emit 物证 `.build/b373_emit_base.c:327`）。**下一格 = #238**（与 Variant 比较的标量操作数没装箱；FD13/FD14 已改成算术形式并点名）。等用户口径的仍是 #219 裸名 `Count` 与 #230/#206 的单位口径。
# 上一轮（2026-10-06T14:20:00+08:00）= # 本轮追加：**门 #361 绿 = 账 #237 收线**（run 37421550422、head `d86b478c`、branch dev、attempt 1、11 job 全 completed/success、非绿 0、wall 8m49s）。Print 现在住在绘图家族里问四件权威（DC / 字体 / 色彩 / 单位），推进量取刚写那串字的 extent = TextHeight 同一个量；同刀撤掉 vb6forms_draw.c 自带那份只认缇的 `v * dpi / 1440`。针面 12→18 条，行为负控与四条假针逐条能红。**本轮撞出来的两格新账已进 §B：#238（§B68）= 与 Variant 比较的标量操作数没装箱** —— `(Me.CurrentY = thTw)` 在两边都是 240 时答 False，发码 `vb6_VarCmpEq(&_vcmp_8, &thTw)` 把裸 double 的地址当 `vb6_VARIANT*` 递进去，静默给错答案那一族的新成员（**暴露面量过：90 份捕获里 158 处 `vb6_VarCmp*(&A, &B)`，两侧全是装箱好的 `_vcmp_N` 或 `vb6_VARIANT` 局部，0 处漏装箱 ⇒ 这是一格**潜伏缺陷**，存量用例永远撞不到；另注意取得这个数字的方法本身→先前按"全文件名字→类型"扫出来的 "32 处 BSTR/int32_t" 是**跨函数污染**的假读数，改成逐处往回找同一个函数体内的声明才对**）；**#239（§B69）= PictureBox 的 Print 还自存一份像素笔位 VB6_PrintX/Y**，与 pic.CurrentX/Y 那份 float 从不汇合，语料只有 1 处控件 Print ⇒ 同样零覆盖。**下一格 = #238**：先把 cgen 比较发码那一处"哪一侧装箱"的判据读准（找 `vb6_VariantFromValue` 与 VarCmp 的生成点），修法应是两侧都装箱而不是改 RTL 原型；本刀已把 FD13/FD14 暂时写成算术形式并在注释里点名。等用户口径的仍是 #219 裸名 `Count` 与 #230/#206 的单位口径。
# 上一轮（2026-10-06T14:05:00+08:00）= # 本轮追加：**账 #237 已出（提交 `32b5eedf`，等门）= `Print` 从 vb6forms.c 搬进绘图家族，DC / 字体 / 色彩 / 单位四件都改问已有的权威**。症状实测（两架构一致）：缇档 `Me.CurrentY = 0 : Print "AB"` 之后 CurrentY = **16**，而同一枚窗体自己答 `TextHeight("AB") = 240`；笔位放到 `ScaleHeight/2` 时墨落在客户区外（Print 把用户单位当像素递给 TextOutW）。同刀顺手撤掉 `vb6forms_draw.c` 自带的那份 `v * dpi / 1440` —— 它**只认缇**（Point/Inch/cm/mm 全按缇算，差 20 倍）且纵向用横向 dpi，正是 Fix 184 说要消灭的第二套口径；换算改问 `vb6_ScaleUserToPx` / `vb6_ScalePxToUser`，11 个调用点显式写纵/横。推进量刻意取**刚写那串字的 `GetTextExtentPoint32W().cy`**（= `vb6_ControlTextHeight` 量的同一个量），所以判据是两条路真汇合：FD13/FD14（缇、点两档的"推进 == TextHeight"）+ FD16（72 点与 1 英寸必须落在同一行，蓝长行与红短行分居同一行两段 x 窗口 ⇒ "第二行没画"赢不了），针面 12→18 条、不钉绝对数。行为负控：落点与推进改回旧形 ⇒ FD13/FD14/FD16 全 False（`curTw=16` 对 `th=240`、`rowB=75` 对 `rowR=-1`）；哨兵 S6/S7 四条假针逐条能红（Print 定义位置、七件权威、码头交回权威、0/1 那一档、dpi 字样）。护栏：24 道哨兵 red=0、90 份 emit changed=0、邻域五枚真跑全绿、矩阵 4 行干净。**本轮最大的一条是撞出来的新缺陷 #238（§B68）**：写 FD13 时 `(Me.CurrentY = thTw)` 在两个数都是 240 的情况下答 False —— 发码 `vb6_VarCmpEq(&_vcmp_8, &thTw)` 把**裸 double 的地址**当 `vb6_VARIANT*` 递了进去，即"与 Variant 比较的标量操作数没装箱"= 静默给错答案那一族的新成员，本刀的夹具暂走算术形式并在注释里点名。**另记 #239（§B69）**：`vb6_ControlPrint` 还留着第二份像素笔位 `VB6_PrintX/Y`，与 `pic.CurrentX/Y`（读 float 那份）从不汇合 —— 就是 #237 在控件侧的镜像，语料只有 1 处控件 Print 所以存量用例全哑。**下一轮第一件事 = 盯本批这个头的门**，绿了回填 §B67 与 022；然后按 #238（比较装箱，量调用点面）→ #239（控件侧笔位）→ #224 剩下的发码针 → #232（parser 两形）的顺序走。等用户口径的仍是 #219 裸名 `Count` 与 #230/#206 的单位口径。
# 上一轮（2026-10-06T13:00:00+08:00）= # 本轮追加：**门 #360 绿 = 账 #235 + 账 #236 两格一起收线**（run 37415128806、head `ea8dc7c9`、branch dev、attempt 1、11 job 全 completed/success、非绿 0、wall 8m31s）。#235 = 画笔色只有一份存储（`Me.ForeColor` 现在走得到 PSet/Line/Circle，FD11/FD12 两针 + 哨兵 S5）；#236 = 门 #359 那条红的归因与修法 —— **红在夹具不在产品**：`Tests (vbp #4)` 里唯一缺 `.err` 的 `ResAlpha.out` 就是被 `-RunTimeoutSec` 60s 杀掉的那枚，成因是「If done Then Exit Sub」走在「tick = tick + 1」前面，计数器被自己那道闸冻住、`Unload Me` 永不达。**本轮新学的取证路**：PAT 读不到 job 日志，但 `actions/artifacts/<id>/zip` 读得到，跟 302 到签名 URL 时**必须甩掉 Authorization 头**（带着它去对象存储必 403）；工件里就是每枚夹具真跑落盘的 .out/.err，而「整片缺哪一份 .err」本身就是判据 —— 正常退出那条路两份都写、超时那条路非空才写。第 24 道哨兵 `check_fixture_timer_close.ps1` 已把这条口径钉住（负控 = 把 HEAD 那份旧夹具放回原路必红，第一次跑它不红是因为 Python 把正则里写的反斜杠-b 吃成了退格符 ⇒ 负控必须真跑一次才算数）。**合并面复核**（merge 别人的 `9037dfc3` rev40 之后重编 C3.exe，exe md5 `fd88a5ea416205818104596992be2516`）：24 道哨兵 red=0、90 份 emit 捕获 changed=0、邻域五枚真跑全绿（fdrawstate 12 针 / pclinedraw 5 / dcsurf 10 / resalpha 2 / ve_units 3）、定点真编译 4 行干净。**下一格候选**：#224②（`Print` 推进笔位按像素而绘图按用户单位，缇档 Print 一行后 CurrentY=16 而非 240）、#232（窗体裸写 `PSet (x,y)` / `Line -(x,y)` 两形 parser 就拒）、#218（内在常量三格机制）。等用户口径的仍是 #219 裸名 `Count` 与 #230/#206 的单位口径。
# 上一轮（2026-10-06T12:41:00+08:00）= # 本轮追加：**门 #359 那条红归因完成 = 账 #236（§B66，提交 `ffcb52f7`，等门）** —— 红的不是产品，是别人那格 rev40 带进来的夹具把自己的收线坐在了自己那道闸后面：ResAlpha 的 tChk_Timer 第一行「If done Then Exit Sub」，而「tick = tick + 1」在它自己那趟的末尾 ⇒ 第一拍把 done 置了之后每一拍都提前返回、tick 永远停在 1，「If tick >= 30 Then Unload Me」再也够不着。**定位手段（PAT 拿不到 job 日志时的替代品，下次直接用）**：actions/artifacts/<id>/zip 拿得到，跟 302 到签名 URL 时**要把 Authorization 头挥掉**，否则对象存储回 403；工件里**唯一缺 .err 的那枚就是被 60s 超时杀掉的那枚**（正常退出那条路两份都写，超时那条路「非空才写」）。本地照做：45s 不退出；修后 0.2s 自己关窗。修法 = 计数器先走、每条路径都有界，第 24 道哨兵 check_fixture_timer_close.ps1 把这条口径钉住（F2: 凡阈值收线的 *_Timer，自增必须在第一条提前返回之前；F1 覆盖面下限、F3 适用面非空）。负控：把 HEAD 那份旧夹具放回原路 ⇒ F2 红并点名行号。另一条工具课：`git show HEAD:<path>` 交回的是索引里的 **LF** blob，工作树因 core.autocrlf 是 CRLF —— 按 CRLF 切那份字节的脚本会把整个文件看成一行，「改好了」其实没改（本轮踩了两次）。**同敞推的还有账 #235（`2ab90321`，笔色只有一份存储，FD11/FD12 两针）**。本轮先 git merge github/dev（别人的 `9037dfc3` rev40）再重建 C3.exe（exe md5 `fd88a5ea416205818104596992be2516`）：24 道哨兵 red=0、90 份 emit 捕获 changed=0、邻域五枚真跑全绿（含新的 resalpha，FD 针面已扩到 12 条）、定点真编译 4 行干净。**下一轮第一件事 = 推 `ffcb52f7` 上 dev、盯那个头的门**，然后回填 #235/#236 两行。
# 上一轮（2026-10-06T12:20:00+08:00）= # 本轮追加：**门 #358（账 #234）绿 + 账 #235 已提交（`2ab90321`，等门）= 画笔色只有一份存储了** —— #234 那一刀是"两份同口径的实现合成一处"（无 bug 症状，但一改就静默分家；census 跟着长时**先补"声明 vs 定义"那道区分**，否则新增的头文件声明被数成第二处定义 = 冤案红），门 #358 = run 37409833257、head `2b3baf82`、11 job 全绿、wall 9m44s。#235 是紧接着的**有症状**那一格：`Me.ForeColor = vbRed` 之后不带颜色的 `PSet` 画出来是黑（改前两架构实测 0 对 255），因为 Form 绘图家族在私有属性 `VB6_DrawForeColor` 上另存了一枚画笔色，而控件那侧（Line / Print / 子控件回显）一直只问 `vb6_GetControlForeColor` ⇒ 这一刀把 Form **对齐到已经正确的那一份**，撤掉私有 setter 与两个零引用导出（Printer 那族无 HWND，不动）。**为什么不需要 Fix 187 那道"Set 过"哨兵**：ForeColor 的默认档本来就是黑 = 0，与"没设过"重合，合并不会误判；那条坑专属 BackColor（默认 BTNFACE）。夹具新一头 FD11/FD12 **两面钉**（新点要蓝 + 旧点仍红，否则"整屏刷蓝"也算赢），并把 `vb6_DrawForeColor` 逐字改回改前那份私有存储、重编 C3.exe 跑同一份夹具 ⇒ FD11=False / pen=0，还原重编 ⇒ True / 16711680；哨兵加 S5（`VB6_DrawForeColor` 再出现在任何 GetPropW/SetPropW 行 = 红）并假针证红。23 份 check 全绿、邻域四枚真跑零缺失、emit A/B 90 份 changed=0（只动 RTL，仅证前端没碰）。**#224 只剩两格**：Print 推进笔位按像素而绘图按用户单位（缇档实测 16 而不是 240），以及 Form 的 `Print` 仍自己 `GetDC` + 无条件 `ReleaseDC`（派发期那张的归属那一问）。**
# 上一轮（2026-10-06T11:20:00+08:00）= # 本轮追加：**账 #234（§B64，提交 `67352011`，已过：门 #358（run 37409833257、head `2b3baf82`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 9m44s）= 拿 DC 这个决定被实现了两遍，本刀撤第二份** —— rev38 在 `vb6forms_draw.c` 里把"派发期先问 `VB6_PaintDC`、否则 GetDC、fromPaint 那张不 Release"整条又写了一份，与 #185/#196 收口在 `vb6forms_ctrl.c` 的那份**逐条同口径**（它自己提交说明就写着"完全同一口径"）⇒ 没有 bug 症状，但一改其中一份另一份就静默分家。**census 同步才是主体**：`check_control_dc.ps1` 原本只盯 ctrl.c，那个洞一起补 —— D1 的"定义"匹配要**排除以 ; 结尾的声明行**（否则本刀新增的头文件声明当场被数成第二处定义 = 冤案红；名单一扩先补"声明/定义"区分），D2 钉死的调用点数 5→6，新 D13 禁 draw.c 再出现 `GetDC(` / `GetPropW(...VB6_PaintDC)`。三条各证能红 + 23 份 check 全绿 + 邻域四枚真跑夹具 28 条 needle 零缺失 + emit A/B 90 份 changed=0（RTL 不进 emit，这条只证前端没碰）。**又踩一次"哨兵绿、编译器红"**：行切片把 `vb6_DrawAcquire` 的收尾 `}` 一起换掉 ⇒ C2143 一片、四枚夹具 build-rc=1 —— 插入/替换的边界必须是**整条语句**，改完必先真编译再信哨兵。**
# 上一轮（2026-10-06T10:40:00+08:00）= # 本轮追加：**门 #356 绿之后紧接的第二格 = 账 #233（§B63，提交 `0c971018`，**已过：门 #357（run 37405355138、head `05f04f31`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m05s**。读数方法订正一句：那台 watcher 回读 jobs 时被本机代理顶了一次，只写出 `jobs=0 non-success=0` 就收线 —— `conclusion=success` 配 0 条 job 不是"全绿"，是**没拿到读数**；补一次按 run id 回 API 复核才数到 11 条）= Form 的绘图状态属性从"压根编不出来"到两头对上** —— 读侧四条属性硬编码在 cgen_expr_member_form_builtin.inc 的一份侧表里，注释还写着"写侧不需要"，实测**恰好相反**：赋值发 C 时先问 getControlPropWriteFn(Form, 名)，没登记就退化成"把读函数当左值" ⇒ `Me.DrawWidth = 3` 发成 `vb6_Form_DrawGetWidth(vb6_hwnd_X) = 3;` ⇒ **C2106，两架构零产物**（RTL 那四个 setter 早就写好也声明好了，只差这张表）。同一刀收掉同族另两格：`vb6_DrawSetI` 存 v+1 而读侧不减（写 3 读回 4、Step 每画一次多带 1）；`VB6_CurrentX/Y` 一个窗口属性名挂两套编码（绘图 int32(+1) vs Form Print 的 float 位图案，互读必错）⇒ 笔位归一到 float 那一户并撤掉四个 int32 导出；`vb6_DrawScaleMode` 改问 #197 那道权威。**判据形状值得复用**：新夹具 tests/fdraw（Timer 第一拍，两架构 Test-Vbp），BASE 那台跑同一份夹具= build-rc=1 + 4 条 C2106，改后 10 条读数全对（含 FD07 像素证人与 FD09-AFTERPRINT=0,46 那枚"Print 推进的数绘图读得到"）；FD09 的 y 与字号/DPI 有关 ⇒ 只打印不当判据。语料 A/B 340 份 changed=0 —— 这条同时解释了这一族为什么能带着 C2106 shipped：**零覆盖** ⇒ 新哨兵 check_form_draw_state.ps1 四条各证能红（S3 那条假改动改的正是本次的伤：write=0）。23 份 check 全绿、邻域真跑四枚夹具零缺失。留下的一格：`Me.ForeColor`（VB6_ForeColor）与画点用的 VB6_DrawForeColor 是两份色彩存储，本刀刻意没碰。**
# 上一轮（2026-10-06T09:50:00+08:00）= # 本轮追加：**账 #231（§B62）= 控件属性的"类型"两处权威合成一处**（提交 `bd11b997`，已过：门 #356（run 37401361458、head `442da251`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m23s）—— C29-1a/1b/C29-9 手抄在 `inferExprType` 里的三份名单（`kNumericFc` 14 / `kStringFc3` 5 / `kStrFcCd` 7 + `kNumFcCd` 7）逐条搬进 `controlPropType` 那张表，问话只留一次、位置仍在 `MemberAccessExpr` 那条 case 的最前面（答案顺序一字未动），`!= FrmControlType::Unknown` 那道闸跟进表里（自定义 OCX/UC 的属性面归类型库）。**这一刀的判据是"零改动"**：refactor 的失败模式不是崩而是"某条名字语料里没人这样写"，那种漏在 emit A/B 上是哑的（上一格 #229 的读数就是"changed 全是夹具自身新增行"）⇒ BASE 先冷存复捕（`C3_base231.exe` 与上一轮 90 份逐字节相同才认），改后 `inputs=90 changed=0 same=90`；补新哨兵 `scripts/check_ctrl_prop_type_authority.ps1`（A1 旧名单回潮 0 / A2+A3 一处定义 + 恰好一个调用者 / A4 那 28 条名字逐条必须在表里 / A5 Unknown 闸 1 处），三条负控各让一条红、跑完 md5 还原；22 份 `check_*.ps1` 全绿、真编译 4 件（ve_units 两架构 + ucTreeMaps x64 + VBFlexGridDemo x64）全出 exe。两条工具事实：① python 的 **bytes** 字面量里写反斜杠 + "rtl" 会被折成一枚真 CR 塞进 CRLF 文件 ⇒ 那一行被劈两半，PowerShell 报 ParserError 而字节数看着正常 ⇒ 校 CRLF 要**同时**数 lone_lf 与 lone_cr；② 插 [STATIC] 块时先按整节边界插、再回头核对标签串有没有替换干净（本轮第一条就把 `[STATIC] ctrl_array_members` 留下了，改完 `uniq` 数过 22 个标签才信）。**
# 上一轮（2026-10-06T09:30:00+08:00）= # 本轮追加：**账 #229 已出 = 门 #355（run 37398206822、head `187dc3e7`、attempt 1）= 11 job 全 completed/success、非绿 0，wall 10m16s**（这一轮的门跑在合并后的 dev 级 head 上：先 `git fetch` 到别人新推的 `bd5a4711`（rev39 LoadRes* 一族），合流后重编 C3.exe、21 枚 `check_*.ps1` 与 `check_rtl_resource_ids` 全绿、同一份 ve_units 夹具两架构 15 条读数全 True，才推 `HEAD:dev` ⇒ 祖先两枚 (`7162cd2e` / `bd5a4711`) 都用 `--is-ancestor` 核过）—— **症状按"属性名撞不撞内置函数名"分家的缺陷**：`& uArr(0).Left` 两架构 0xC0000005，同元素 `& .Top` 只是绕远装箱，非数组 `uPix.Left` 正常。根因不是 Left 特殊，而是类型推断里"`对象.成员` 的对象位是不是控件"被抄成两份（一份只认 `控件名.属性`、一份只认 `控件名(i).属性`），元素那一形从 P20-42 那条兜底前面掉下去 ⇒ 撞返回 String 的内置函数 `Left` ⇒ 拼接不套数值转换 ⇒ 裸 int 进 BSTR 槽。收成一处出口 `ctrlTypeOfMemberObject`。判据两头：RAW 那行就是崩溃现场（四枚读法全在 `&` 里），另一行问四个变量 = 设计期几何 3600/1320/1200/1140；负控 = 改前那台跑**同一份新夹具** ⇒ 两架构崩且整行不出现。A/B 90 份只有 ve_units 变且差异全是夹具自身新增行⇒ 产品发码零改动、全语料没有别的该形状 ⇒ 只能靠新夹具守。**本轮另两格**：#228（库限定名 `stdole.OLE_COLOR` 与裸名 `OLE_COLOR` 给出两种 C 类型，收成问类型本名一处，真 COM 限定名 39 种 / 124 处零改动，门 #353 全绿）、#227（`desc->dblClick` 落点，门 #352 全绿 wall 8m41s）；§B 按口径只留未完成项（删 §B55..§B58，叙述在 `dd73c049`）。**新开的 #230（§B61）已量到读数**：数组元素 `Left = 5000` 读回 4995、`Top = 2000` 读回 1995，而 `Width = 900` / `Height = 600` 精确往返（15 缇/像素整除）⇒ 丢的是"缇→整数像素→缇"那一档，与 #206 同族；写面本身不崩、邻枚不受影响（`uArr(2).Top` 仍是 2520）⇒ 这一格是**口径题不是崩溃题**，先定"以哪一侧为准"再动。四份 VbEclipse 语料工程现在仍只剩 `ucProgressCircular` 那 1 条 C2065 裸名 `Count`（#219 最后一格，等口径）。**
# 上一轮（2026-10-06T08:40:00+08:00）= # 本轮追加：**账 #228 已出（提交 `23659871`，门 待回填）= 同一个 VB 类型的两种拼法从此给出同一种 C 类型** —— `As OLE_COLOR` 答 `int32_t`，而 VB6 里同义的库限定名 `As stdole.OLE_COLOR` 一路掉到 `mapTypeRef` 末尾的兜底 `void*`；实物 = VBFlexGrid 的 `Public Event EditSetupWindow(... As OLE_COLOR)` 对上容器 `UserEditingForm.frm:350` 写的 `As stdole.OLE_COLOR` ⇒ 发送侧交 4 字节、处理器收 8 字节指针（x64 高 32 位是垃圾）。改一处：别名那几档（Vb 前缀 / OLE_ 前缀 / Enum 结尾 / 两张名单 / LongPtr / LongLong）问 `aliasName` = 点号最后一段，符号那几档（ivref/Class/UDT/枚举/ComClass）一律不动。读数：单变量 A/B `inputs=90 same=88 changed=2`、4 条差异全在那一枚处理器、真 COM 限定名一族（39 种 / 124 处）**零改动**；三处同源检具（52 枚 thunk）thunk↔typedef 0 不符、↔处理器 **1→0**（#222 那一族收平）；夹具 `tests/test_alias_spellings.bas` 两面钉（该折的折 + 不该折的不折），BASE 那台跑同一份夹具真红；真编译矩阵 9 件工程两架构 rc=0（ucProgressCircular 仍 1 条 C2065 裸名 `Count`，属 #219 那一族，等口径）。同轮 **#227 = `desc->dblClick` 的落点已过门 = 门 #352（run 37388004707、head `50fb6d5f`、attempt 1）11 job 全绿、wall 8m41s**，**#222 第二格 = 门 #351 全绿**（修 #350 上本线自己带进去的第二张类型表）。§B 按口径只留未完成项（删 §B55..§B58，叙述在 `dd73c049`）。两条自检：① 旋转 LAST_RUN 要**替换**那一条，用插入写法会留下两条（本轮真留下过一条，已删）；② 新夹具 `git ls-files --eol` 读成 `i/lf` 是 `core.autocrlf=true` 下 .bas 的**正常形态**（`.gitattributes` 只给字节敏感的 ai/028 夹具标 `-text`）⇒ 先读仓库自带说明再动手"修"。**
# 上一轮（2026-10-06T07:40:00+08:00）= # 本轮追加：**账 #222 的第二格已出 = 门 #351（run 37386871868、head `482c3273`、attempt 1）11 job 全 completed/success、非绿 0**（上一轮 #350 两档红 = `charts_ucTreeMaps` x64+x86，是本线自己带进去的回归：prelude 解析 `.ctl` 的 `Public Event` 时自带第二张类型表 `kTypeMapPre`，缺 `Variant` 一档 ⇒ thunk 声明 `int32_t a0` 而回调 typedef 与容器侧处理器原型都是 `vb6_VARIANT` ⇒ C2440；顺带量出该表把 `Integer` 记成 `int32_t` 而 `mapTypeRef` 给 `int16_t` ⇒ VBFlexGrid 三十多枚事件形参一直按错宽度接）。修法 = **撤表**，类型名现拼 `SimpleTypeRef` 喂 `mapTypeRef`，三处（typedef / 处理器原型 / thunk）收成同一权威。检具读数（5 件工程 49 枚 thunk）：与 typedef 不符 **5 → 0**、与处理器原型不符 **6 → 1**（剩那条 = 容器自己把形参写成 `LongPtr` 而事件声明是 `Long`，用户代码形状）。单变量护栏 A/B（BASE = 把 prelude 回退到 HEAD 那版冷编）：inputs=92 same=86 changed=6、28 条差异行**全部**落在 `evtThunk` 签名上、未归因=0、两档逐份对称。真编译矩阵两架构：ucTreeMaps **出 exe 且诊断 0**，Charts 2020 / ucChartBar / ucPieChart / ucChartArea / czUI / VBFlexGridDemo / ve_units 全 rc=0，ucProgressCircular 仍是那 1 条 C2065 裸名 `Count`（属 #219 那一族，未动，等用户口径）。判据：哨兵 `check_uc_array_event_sites.ps1` 加 E5（假 `kTypeMapPre` 真验红）+ 发码针 `ucevt_thunk_type_same_authority`（Absent = 改前真实形状）。**第二格：账 #227 已出（提交 `50fb6d5f`，**门 #352 = run 37388004707、attempt 1、11 job 全绿、wall 8m41s**）= `desc->dblClick` 那一槽 cgen 一直在填、宿主从不转调** —— 与 #226 同形（布局式填表 ⇒ 发码面永远绿），六枚 UC 的 `RaiseEvent DblClick` 全静默；`uc_host_window.c` 补独立一档 `case WM_LBUTTONDBLCLK`（MouseDown/Up 刻意不挂：物理双击 = DOWN/UP/DBLCLK/UP，Click 已由两条 UP 供过）。判据 = ve_units 第三头真手势，**问两个计数**（`dbl=1` 且 `hits=0`，只问一个就会放过"挂错到点击那一格"）：两档 `U-ARRDBL-RAW hw=True idx=2 dbl=1 hits=0 ret=0` ⇒ True；负控 = 注释那条转调 ⇒ False 且现场 `idx=-1 dbl=0 hits=0`（另两头照旧 True，还原后 MD5 逐字节回位）；哨兵加 C4（鼠标/点击五槽逐槽钉非零 + 恰 1 处 + 位置）；92 份 emit 与上一轮逐份相同（纯 RTL 那一刀）。21 枚 `check_*.ps1` 本地全绿。一条又踩了一次的工具雷：`io.open(p,"wb")` **先截断再算表达式** ⇒ 补丁脚本里 join 报错就把原文件清成 0 字节（这次靠哨兵把整族读数报成 0 才暴露，`git checkout --` + 重放补丁、MD5 回位）；规矩 = 先把新内容整个算成 `body` 再 write。**
# 上一轮（2026-10-06T06:17:51+08:00）= # 本轮追加：**门 #349（run 37380276520、head `ca9f3719`、attempt 1）= 11 job 全 completed/success、非绿 0 = GA 回到全绿** —— 账 #225 那一刀把上游 `9a420157`（rev38 绘图家族）对调的 RTL 资源 id 修回来了：症状不是"少一个文件"而是"头里装着体"（`vb6forms_draw.h` 里是 463 行的体），而 `vb6forms.h` include 它、被 39 个 RTL 文件 + 每份生成模块 .c 带到 ⇒ 28 枚符号在约 49 个 TU 各定义一份 = LNK2005×1225 + LNK1169 ⇒ 前一台门 #347（head `c5aab787`）十一片九红、只有不链接的两片绿。判据：新哨兵 `scripts/check_rtl_resource_ids.ps1` 三处 125 条逐条对账（两条 R3 负控真红、还原逐字节相同）+ 单变量真编译（HEAD 冷编 RC=1 diag=1226 → 只改这两行 RC=0 exe=True）。同时回填：账 #219 的门 = #346（head `0026d87f`）11 job 全绿。**第二批（账 #222 控件数组 UC 事件臂同源化 + 账 #226 UC `UserControl_Click` 的 click 落点）已提交待门**：真编译 `ucProgressCircular` 的 `Form1.c` 诊断 **26 → 0**（BASE 那台量到 C2198×25 + C2084×1；只剩 #219 那一族的 `Count` C2065）；真跑 `tests/ve_units` 两档两头 `U-ARREVT=True`（i1=1 h1=1 i2=2 h2=2 ret=7）+ `U-ARRCLICK=True`（真手势 WM_LBUTTONUP → desc click 槽 → UC_Click → 本元素 sink；负控 = 注释掉那条转调 ⇒ False 且现场 idx=-1 hits=0，而 Fire 那头照旧 True）；语料 A/B inputs=92 same=74 changed=18（改到的全是含 UC 的工程，两档逐份对称），差异行 776 条**逐条归因、无法归因 0**；22 枚静态哨兵本地全绿。四条 VbEclipse 工程里现在真编译不过的只剩 `ucProgressCircular`（1 条诊断，属 #219 的裸写未声明名那一族）。
# 上一轮（2026-10-06T03:47:00+08:00）=   # 本轮追加：**账 #219 已出（提交 `b5ae1240`，门 待回填）= `.pag` 里裸写的 Changed 那句 VB3001 停了、RTL 那枚裸名全局撤了、判据收到 kHostPseudoRows 那张表上** —— 起点是账 #220 收线时积下的两枚实物：`Public Changed As Long` 撞成 C2371（`.build/b229out/pjChanged.bas`、改前 BUILD-RC=1 / no exe）与 #174 那句「裸写伪成员的诊断面还没收」。  **关键读数：这一格的发码从来是对的**。`Changed = True` 在 `.pag` 里早就发成 `vb6_PropertyPage_Changed = (-1);`，语料 186 处、裸名在产物 C 里 0 处（账 #159 那张表回答的）⇒ 那枚裸名 C 全局既没人引用、又占着 C 的全局名字空间，是纯负担；而语义层从来没问过那张表，于是每条合法裸写配一句噪声（24 条）。做法三处：① RTL 的定义与 extern 各换成一条记录实测的注释（§C18 那条口径的第三枚实物，前两枚是 #220 的 B/BF）；② `visit(IdentifierExpr)` 的"未找到"分支补一格「裸写的文档成员」，与限定符位那一格（`memberObjCtx_`，#217）互斥，类型答案仍走 Variant；③ **这一格不许自己认名字**：把表从 `src/backend/cgen_util_com.cpp` 的匿名 namespace 搬进 `src/common/host_pseudo.hpp`（inline 表 + `hostPseudoFind` + 新增 `hostPseudoBareEligible`），发码与语义两头问同一句；common 不得向上依赖 semantics，故表里的 `Symbol::toLower` 换成头文件自带的 `hostPseudoLower`。**把 "changed" 抄进语义层，本账剩下的那几枚（hDC / Controls / ScaleWidth）就得抄四遍**；收成问表之后，表里加一行 HPF_BARE 顺带把诊断面放行了 —— `.ctl` 里那 24 条 hDC 就是这样一并收掉的。  读数三套：① 发码语料 A/B（BASE=`b230_new_emit` = #221 收线那台，NEW=`b233_new_emit`，inputs=100）changed=**12**、same=88，逐份只有 `-N` 而**每一条差异都是一行 VB3001、产物 C 一行没动** ⇒ 这一刀只动诊断面；VB3001 146→98，名字档 `Changed` 24→0、`hDC` 24→0、`new-names={}`，其余 18 种一个没动；`vb6_PropertyPage_Changed` 186=186、`int16_t Changed` 10=10（后者是 VBFlexGrid 的 UC 形参面，与本刀无关）。② 本地探针三头（`.build/b233out/`）：`pagBare.pag` 零报且三处命中 `vb6_PropertyPage_Changed`、同一份里打错的 `Changd` 照报；`basBare.bas`（标准模块）裸写 Changed **仍报** ⇒ 放行按文档类别、不是一片名字；`ctlBare.ctl` 裸 `hDC` 发成 `vb6_UserControl_hDC` 且零报。③ 撞名那头真编真跑：`tests/test_rtl_naked_changed.bas` 改前 RC=1 / no exe ⇒ 改后 RC=0 / exe / `NC219-CHANGED=6 NC219-VT=3`（x64 本地，x86 交门）。  哨兵两份跟着表搬家：`check_host_pseudo_table.ps1` 的 `$tblPath`、`must`（表与出口在 common、语义层必须问表）、`deny`（语义层不许出现成员名字面量）三处，census 从只扫 src\backend 扩成 backend+common+semantics；`check_rtl_naked_names.ps1` 的 N7 换成 谓词 定义/声明/调用/带门 = 1/1/1/1 + 问表 1 + 成员名字面量 **0**，N2 名单同批从 5 枚缩到 4 枚（`Changed` 撤下去了）。哨兵红过一次是当场演示的负控：N7 里读那三个文件的三行被我的整块替换吃掉了 ⇒ 六条计数全 0、五条 FAIL；补回去即绿。**一条工具口径同批记下**：旧的 `[System.Management.Automation.PSParser]::Tokenize` 报错时既没消息也没行号（只给 COUNT=1），换成 `[Language.Parser]::ParseFile` 才指出 L128「逗号后缺表达式」—— 那是我往数组字面量末尾多写的一个逗号；以后改 `run_tests.ps1` / 哨兵的常规检查一律用后者。  剩下的同族读数已钉（本账没做完）：`Controls` 4 条全在 ppProgressCircular.**pag**（表里 propertypage 档没这一行，要先问 VB6 里 .pag 裸写 Controls 是谁）；`Count` 2 条在 ucProgressCircular.**ctl**（表里也没，而 #159 的边界写明「RTL 没有对应全局的行刻意不收」⇒ 那是 RTL 侧缺口，另立账）；`ScaleWidth` 2 条在 frmDemo.**frm**（窗体成员走另一条路，不在这张表里）；另 26 条是内在常量一族 ⇒ 账 #218。台账 §B54 一节 + §C 一条（跨层的判据表住在 common）+ §D 一行。  下一格候选：#222（ucProgressCircular 编不过的第一半：控件数组事件臂三种形参表 + prelude 按元素发 24 遍，三件已量完）、#218（内在常量三格机制：表 408 / RTL 宏 14 / 发码特判 1，动之前先读类型库）、#192 剩下的 RTL 半边、#193、#206 / #208 / #213 / #79。
# 上一轮（2026-10-06T02:55:00+08:00）= 本轮追加：# 本轮追加：**账 #221 = C29-PL-a 已出（提交 `27767255` + 名单订正 `b0e4da16`，**门 #344 红在自己的名单针 → 门 #345 = run 37359094337、head b0e4da16、attempt 1 = 11 job 全 completed/success、逐片 FAIL=0**）= Picture.Line 从 COM 兜底改道到原生 GDI，旗标改按字母位折** —— 起点是账 #220 落地时看见的形状：源侧 `ppProgressCircular.pag` 三句 Line（297/299/474）一份工程发码出 8 条调用，全都发成 `vb6_ComCallObject(vb6_ComGetObjectProp(vb6_hwnd_PictureN, L"Line"), L"Item", {…}, 6)`，即**先把方法名当属性取、再对取回的东西取 Item**；而 RTL 两处把 Line 登记成"认识但什么都不做"（`vb6forms_axcontainer.c:304` 属性位交 `axSetEmpty`+S_OK —— 交回的是 Empty 不是 IDispatch，第二跳无处可调；`uc_hostmodel_call.inc:152` 直接 `return 1`，注释「未知方法一律空实现」）⇒ **两跳都成功、两跳都空、一条诊断都不打**，本线第三条控件方法静默空转（#143 SetFocus / #196 hDC·TextHeight·ScaleX / 这一条 Line），差别只在 Line 是六个实参那一形。需求面先量：`Circle/PSet/Point/PaintPicture` 在四份真源码里 0 处，`Cls/Print` 早已原生（语料 `vb6_ControlCls` 14、`vb6_ControlPrint` 2）⇒ Line 是绘图面最后一个裸着的。做法三处+一处判定：RTL 新出口 `vb6_ControlLine(hwnd,x1,y1,x2,y2,color,style)`（DC 走 Print/Cls **同一处** `vb6_ControlDrawDC`，坐标按 `vb6_WindowScaleModeSelf`+`vb6_ScaleUserToPx` 折 #196/#197 那张单位表，`color<0` 落 ForeColor，不填充那档显式 `NULL_BRUSH`）+ 后端一张 `controlCanvasMethod`（照零/一实参/ScaleX 三张表的规矩：表只交名字、实参由码头拼）+ withm 码头 + **成员侧**也要打标记。最后那一处是自己踩出来的：表与码头都写好、重编一台，产物照旧是 `ComGetObjectProp(L"Line")+Item`、夹具四形全 False —— 成员侧不打标记时兜底 Fix 023e/089d 已把 `comObjExpr_` 换成 HWND 表达式，调用侧那张表根本收不到标记；于是哨兵加 **N6 第四条（成员侧必须问表）**，把"只接一头"钉死，这条口径同时进 §C19。值口径订正：账 #220 那一版为不改语义沿用了旧全局的 B=1/BF=2 且只认两个词，现在按**字母位**折 —— B=1 矩形 / C=2 椭圆 / F=4 填充 ⇒ BF=5、CF=6，`C`/`F`/`CF` 从此有落脚点；这张位口径是 parser 与 RTL 同一张表的两头，N5 钉的就是这个（RTL 里 `style & 1/& 2/& 4` 各出现 + `vb6_ControlLine` 1 份）。判据两面都真跑：编译面 `pcline_flag_folded` 换针（四条正针钉 `vb6_ControlLine(...(int32_t)255, (int32_t)5)` 一形与 color 格留用户名字那两形，三条 Absent 钉兜底形状与改前 `vb6_ComPackValue(B|BF)`；假针 `vb6_ControlLine(...(double)987654)`=False 证匹配没失灵）+ 运行面新夹具 `tests/pcline/PcDraw.{frm,vbp}` 进门禁（`pclinedraw` / `pclinedraw_x86`），画完**问像素**、四形各钉两头：`PL01-LINE=True`（对角线红、旁边不红）`PL02-BOX=True`（边框蓝、中心不蓝）`PL03-FILL=True`（中心绿 —— 顺带证明 BF 与 B 不是同一个数）`PL04-CIRCLE=True`（沿 y 开小窗口找到切点、中心不红），x64 与 x86 **逐行相同** `diag=255 boxedge=16711680 boxmid=16777215 fillmid=65280 circletop=255`。两条工具读数：画与问都要放在 **Timer 第一拍**（先写在 Form_Load 里 GetPixel 一律 -1=CLR_INVALID），DC 用 `GetDC(控件 hwnd)` 而非 `控件.hDC`（后者在这枚夹具读出 0，属 #196 另一问）。发码 A/B（BASE=`b228_new_emit`，NEW=`b230_new_emit`，inputs=100）changed=**4**、same=96，每份 +8/−8 且每行都是同一条调用换出口；CENSUS 兜底 `L"Line"` **32→0**、原生 `vb6_ControlLine(` **0→32**（正好 8×4 投影），`VB3001` 146=146、`VB6_SA_AT(` 20972=20972、`_vb6_select_` 7346=7346 ⇒ 零新增噪声。真工程那一头 `ucProgressCircular/Proyecto1.vbp` 两档真编译仍 rc=1（27 条 error C/档），但**逐文件归因**全在旧账：`Form1.c` C2198×25 + C2084（控件数组事件臂出现两种形参表、`evtThunk_*_Click` 被发两遍 ⇒ **另立新账 #222**）与 `ucProgressCircular.c` C2065 `"Count"`（账 #219 的 HPF_BARE 那半），而 Line 所在的 `ppProgressCircular.c` **零条诊断**、错误行没有一条提到 `vb6_ControlLine`，BASE 语料里那几行 thunk 逐字相同 ⇒ 那些红不是本刀的因。哨兵 PASS `B/BF 0, census 5 pinned names, fold bits 1+1+1, handoff 1+1, flag names in AST 0, RTL 1/bits, canvas 1+1+2+1`；`run_tests.ps1` PSParser=0 错、全文件 eol 统一。**门 #344 那一红归因清楚并修在同一批**：compile 片 28/29，红的是一条钉死名单（`check_control_dc.ps1` 的 D2「拿绘图 DC 的调用点恰好 4 处」），Line 接进同一处权威是正当扩容 ⇒ 名单 4→5 + 文件头那条没同步的旧注释（还写着 3）一并订正；#344 其余十片当场就绿、两枚新夹具 `[VBP] pclinedraw` 与 `pclinedraw_x86` 各自 PASS（vbp 四片 TOTAL 之和 226→228 正是这两枚）。**门 #345 = run 37359094337、head `b0e4da16`、attempt 1 = 11 job 全 completed/success**：compile 29/29（control_dc 与 rtl_naked_names 同片都在）、syntax 157/157、bas 48+48、vbp 四片 50(+1 SKIP test_vbman)/55/50/51、asm 13/14、smoke 1/1。一条一般式同批进记忆：**往一族共用的出口加站点时，`scripts/check_*.ps1` 全部要跑一遍**（20 枚几十秒），别只跑自己新写那枚 —— 这次就红在别人那枚的名单上、本地没读到。边界写明五条：`ScaleLeft/ScaleTop` 原点偏移没进来（语料都是 0，非零原点会画偏）、Line 之后 `CurrentX/CurrentY` 移到终点那条没进来（调用点从没读回，且要先定单位口径，与 #192 同问）、`With picA : .Line (…)` 那一形**没测**（#192/#150 记着带实参的 With 形至今没接）、表里**刻意不给 Form 那一档**（只有 PictureBox 那处成员侧打了标记，给了就是广告比应答复）、`Circle/PSet/Point` 需求面为 0 所以没留出口。下一格候选：#222（ucProgressCircular 编不过的第一半）、#219（裸写伪成员 + RTL 名单里那枚 `Changed`，已有 C2371/LNK2005 两枚实物）、#218（内在常量实为三格机制：表 408 / RTL 宏 14 / 发码特判 1，先读类型库）、#192 剩下的 RTL 半边、#193、#206/#208/#213/#79。
# 上一轮（2026-10-06T01:37:00+08:00）= 本轮追加：本轮追加：**账 #220 已出（提交 `4b75691f`+`91a004fe`，**门 #343 = run 37350249122、head 91a004fe、attempt 1 = 11 job 全 completed/success、逐片 FAIL=0**）= Picture.Line 的 B/BF 由 parser 折成字面量，RTL 那两枚裸名 C 全局连 extern 一起删** —— 起点是账 #217 第二刀边界里那条「BF / B 未归家」：Fix 102 把 `(x1,y1)-(x2,y2)` 吸收成 Line 的实参表之后，尾部 `, color` / `, BF` 走的是普通实参路 ⇒ 名字进 AST、发码原样发裸名，而 RTL 为了让它有落脚处写了 `const int32_t B = 1; const int32_t BF = 2;` 并 extern 到头文件。探针 `.build/b228out/clash_b.bas`（`Public B As Long` + `Public BF As String`，Main 里 `B = 7`）改前实测 **BUILD-RC=1 / 5 条诊断（C2373 重定义 ×4 + C2166 赋值给 const 对象）/ exe=False**，改后 **RC=0 / exe=True / CLASH-B=8 BFLEN=1 / 诊断 0 条**。结构上的定罪理由不是「这两个名字不好」，是 **C 的名字空间与 VB 的名字空间共享**：生成的模块 C 会 #include 那批 RTL 头，而用户模块级变量在 C 里同样是裸名 —— 头里 extern 的在编译期撞，.c 里**非 static** 定义的在链接期撞（LNK2005），只有 static 的不撞（这条口径写进 §C18）。做法收成一处：parser 的 Line 尾部循环只认 **style 那一格**（`trailingIdx == 1`，即 color 已给出），把 `B` / `BF` 折成 `LiteralExpr(Integer)` 1 / 2（沿用删除前两枚全局的值，本刀刻意不改语义），RTL 的两枚定义与两行 extern 删掉。**位置口径是这一刀的全部内容**：`Line (a,b)-(c,d), B` 那一格按 VB6 是 color，用户的 `B` 必须照旧成立 —— 折错格就是「修一处撞车、制造一处静默错值」，所以夹具两头都写。判据三头：① `tests/test_nameclash.bas` x64+x86 真跑（`NC-B=13 NC-BFLEN=2` / `NC-ACC=15` / `NC-BOX=3/8` / `NC-DONE`，改前那份就是 no exe；CI 侧 bas 两片 47→**48**+**48** 进门禁且那片 FAIL=0，但那两片只落摘要行 ⇒ 逐针读数算本地这份）；② `tests/pcline/PcForm.frm` 的 [CODEGEN-NOTE] `pcline_flag_folded` —— 四条针钉位置（两格 style 折成 `vb6_ComPackInt(2)` / `(1)`，两格 color 留 `vb6_ComPackInt(B)` / `(BF)`），两条 Absent 钉改前语料实测的那两形 `vb6_ComPackValue(BF)` / `(B)`，假针负控 `vb6_ComPackInt(987654)`=False 证文本匹配没失灵；③ 哨兵 `scripts/check_rtl_naked_names.ps1` N1..N4（N2 = **非 static 裸名文件作用域数据全局的钉死名单**，现存 5 枚 `Changed` / `g_hoCount` / `g_uc_descCount` / `g_uc_recCount` / `g_uc_dumpSeq`，`Changed` 由账 #219 收；负控 = 往 `vb6rtl_com.c` 插一行 `int32_t b220probe = 0;` ⇒ N2 当场红并点名，撤掉复绿，插拔两头都按字节核过）。发码面 A/B（BASE = HEAD `a7a929c9` 那台，inputs=100）**changed=4**（ucProgressCircular 那两份 × 两档），其余 96 份一行没动；每份 +8/-8 且**每行差异只在最末一格** `vb6_ComPackValue(B|BF)` → `vb6_ComPackInt(1|2)`，行内其余字符逐字相同；另 3 行 VB3001 纯删（ppProgressCircular.pag 63/65/240 三处），Charts 主工程 VB3001 **34→31**。台账 §B52 一节 + §C18 一条 + §D 一行（`git diff --numstat` 记 +51/-0，三处锚点各命中一次）。**顺带量到一枚新账 #221**：那 8 处 Line 现在仍是 `vb6_ComCallObject(vb6_ComGetObjectProp(vb6_hwnd_PictureN, L"Line"), L"Item", {…}, 6)` —— 先把 Line 当**属性**取对象、再对它取 Item，而宿主应答表（Fix 112 那张）里 Line 从没登记、RTL 也没有 Line 的实现 ⇒ **运行期静默不画**；本刀只把名字归还用户，没让它画出来。**§C18 那条口径随后被两枚探针坐实**（`.build/b229out/`，两台都 no exe）：`Public Changed As Long` ⇒ C2371 重定义（头里 extern 那枚，编译期撞）；`Public g_hoCount As Long` ⇒ LNK2005+LNK1169（只在 .c 里非 static 定义、它的头没进生成的模块 C，链接期撞）⇒ 账 #219 现在带着实测症状。一条读法同批记下：cl 的诊断不在 C3.exe 的控制台输出里、只在 `<out>/c3-error.log`（按 gbk 解），否则会出现「BUILD-RC=1 且控制台 grep error C 得 0 条」这种假象。门 #343 逐片读数：compile 片 28→**29** 里逐行读到 `[STATIC] rtl_naked_names ... PASS`、syntax 片 156→**157** 里逐行读到 `[CODEGEN-NOTE] pcline_flag_folded ... PASS`、vbp 四片 50/54/49(+1 SKIP test_vbman)/51 与 #342 同形、asm 13/14、smoke 1/1；`Build C3.exe` 那片日志正文不含用例行（历轮同形的 empty-shell），结论按 overall success 与逐片 FAIL=0 记。下一条读数纪律：`--emit-c` 的产物里出现 `ComGetObjectProp(..., L"方法名")` 这种「把方法当属性取」的形状，就要问一句运行期谁应答它 —— 编得过与画得出是两件事。下一格候选：#221（Line 绘图面）、#218（内在常量两份权威，动之前先读类型库；本刀的 B=1/BF=2 沿用旧值同属这一族）、#219（裸写伪成员 + 名单里那枚 `Changed`）、#192 剩下的 RTL 半边、#193、#206 / #208 / #213 / #79。
# 上一轮（2026-10-06T01:20:00+08:00）= 本轮追加：本轮追加：**账 #217 连出两刀** —— 第一刀 提交 `eef2c199`（门 #341，head ab1db9c0）：`Case Is > 2` 里那枚 Is 是 parser 造的假标识符。探针三面（`Case Is > 2, 1` / `Case Is > 5` / `Case Is < 0`）在 .bas 上告警行号是真的（6/17/19），`a Is b` 与 `TypeOf o Is Collection` 一条不出 ⇒ 靶子当场定形；`parseCaseValue` 为借优先级造的 `IdentifierExpr("Is")` 一直留在 AST 里，语义层按未声明名字处理 —— 开着 Option Explicit 每形一条 VB3001（两份真工程 36 条、发码语料全仓 138 条），关着时每枚用到的过程真发一枚没人读的 `vb6_VARIANT Is = vb6_VariantEmpty();`（探针 2 枚→0）。收成 `CaseValue.relOp/hasRelOp`，占位符出函数即丢，发码三档改读 cv.relOp，clone 两字段都抄。判据：`tests/test_caseis.bas` 7 针 x86+x64 真跑 + 两条 [CODEGEN-NOTE]（neg 那份**两头钉**：真没声明的名字照旧报、`'Is'` 必须不再报）+ [STATIC] caseis_shape（compile 片 26→27，两条负控各自红过）。A/B：inputs=100 same=90 changed=10、每份 +0 行纯删诊断、unattributable=0、`_vb6_select_` 7346=7346。第二刀 提交 `d2942bc8`（**门 #342 = run 37345079456、head 910b37b8、attempt 1 = 11 job 全 completed/success**：compile 片 27→28 里 `[STATIC] dochost_authority ... PASS`、syntax 片 154→156 是两枚 dochost 针，bas 两片 47+47，vbp 四片 50/54/49(+1 SKIP test_vbman)/51，红=0；**上一轮 #341 那条 frmevents 红本轮 PASS** ⇒ 双 0 字节 + 3 秒内失败 + 发码逐字节相同那三条旁证成立，账 #79 的抖动换了症状继续开着，与这两刀无关）：**文档隐式对象收成一处前提** —— `UserControl`/`PropertyPage`/`Extender`/`Ambient`/`VBA.` 在语义层此前完全不存在，而"这是哪一类文档"这个前提仓里猜过两处（driver 按扩展名、cgen_form 再从 controlTypeName 找 "PropertyPage" 字符串）。加 `DocumentKind` + `Module::docKind`（driver **一处**写，哨兵钉 writers=1），语义层 `isDocumentHostObject` 只放行**限定符位**、**不改类型答案**，发码改读 docKind 删掉字符串猜测。读数：VBFlexGridDemo VB3001 **499→19**、Charts 主工程 **499→34**（x86/x64 逐字相同），全仓语料 CENSUS VB3001 **2934→158** 而宿主符号一动不动（ScaleWidth 436=436 / Ambient_UserMode 94=94 / PropertyPage_hWnd 16=16 / hWnd 70=70 / Extender_Tag 4=4），VB7006 0=0；A/B inputs=100 same=82 changed=18 全 +0 行、unattributable=0；夹子 `tests/dochost/dhExp.ctl`（Absent VB3001）+ `dhImp.ctl`（不写 Option Explicit，Absent 三枚死局部）+ [STATIC] dochost_authority D1..D5（compile 片 27→28，四条规则各自负控红过、恢复按 md5 核）；真编译 spot：grid x64/x86 各 0 error C / 0 LNK、出 exe。**别人那一刀（`1b0fb795` 除零抛 Err 11 + Load 控件数组 + B 族字节串）核过**：它的门 #340 是 attempt 3 全绿、attempt 1/2 取不到读数（GitHub 的 jobs?attempt=N 会退回最新一次，三次 job id 完全相同 —— 记进台账工具面）；本地探针 x64+x86 各一遍：`10 \ 0` / `10 Mod 0` / `10 / 0` 全报 Err 11、`LeftB/RightB/MidB/AscB` 读 AB/BC/B/65、0 error C、0 LNK ⇒ 三面的行为都对，没有需要我修的回归。我这一格的红只有一条：门 #341 `Tests (vbp #2)` 的 `[VBP] frmevents ... FAIL (output mismatch)`，工件里 `FrmEvents.out` 与 `.err` **都是 0 字节**，且它距上一条 PASS 只有 3 秒（不是 60s 超时闸），而 frmevents 的发码在 BASE 与 NEW 之间**逐字节相同**（592 行 identical）⇒ 与这两刀无关，是 #79 那一格（拖放点按窗口位置现算）换了个症状：从"少几条 EV"变成"整个进程没输出"。已登记，下一格收。
# 上一轮（2026-10-05T22:05:00+08:00）= 本轮追加：**账 #215 已出（提交 `48feae8e`，门 #338 全绿）= 体级声明收成「一条声明符一条 LocalDeclStmt」** —— 起点是账 #216 收尾时顺手量到的另一件事：过程体内 `Const A = 1, B = 2` 之后再用这两枚名字，**每一枚各报一条 VB3001**（原账写的是"只登记第一个"，读数推翻：一枚都不登记），而同形状的 `Dim x, y` 一条不报。`--dump-ast` 给出形状差：Dim 出两条 LocalDecl、Const 出一条装着 MultiDecl 的 LocalDecl，而 `semantic_analyzer_stmt.cpp:251` 的 switch 只认 VariableDecl / ConstDecl（src/semantics 里 MultiDecl 零引用），default 什么都不做；发码那侧有分支 ⇒ **代码照发、值照对、只有符号表没有**。四条路两种形状的根因是 parseDimStmt 当年**在语句层手写了一遍声明符解析**，而另三条走共享的 parse*DeclList —— 两份实现从此分叉：共享那份后来补了 WithEvents 与 Task #40 的「后缀即类型」，手写那份没有，于是实测出一条真值差：`Dim a&, b&` 第二枚落 `vb6_VARIANT`（探针 `.build/b220c/bd_probe.bas`，改前 6 条 VB3001 + Variant；改后 0 条 + int32_t）。做法：新增 `Parser::wrapBodyDecls` 一处，四条路全调它，Dim 那 40 行手写展开删掉改调共享的 parseVariableDeclList（两步落后因此一起追平）；模块级那套 parser_module.cpp:202 的展平**没并进来**（消费者集合不同，先各自收口）。**两头判据**：`tests/test_bodydecl.bas` 8 针 x86+x64 逐行相同 + 语法片 `[CODEGEN-NOTE] bodydecl_one_per_declarator`（正针钉四行声明，Absent 钉 VB3001 不许出现）。**负控是退码重跑出来的**：把两份 parser 文件 checkout 回 HEAD 重编一台 —— 同一份夹具读出 `BD-empty=3/0` 与 7 条 VB3001，换回来 `3/3` 与 0 条；其余六条针改前改后一致 ⇒ **靶子只有两枚，另外六枚是护栏**（这种区分要写进台账，不然八条针看着都像会红）。发码面 A/B 100 份 same=98 changed=2（只有 grid 两片），census 四项一动不动，那 −276 行**逐条是诊断行**（VB3001 778→502），产物 `#undef` 去缩进 2082=2082、声明行 552=552 ⇒ 这格在真工程面上就是去噪音 + 防住 `Dim a&, b&` 那类静默 Variant。真编译四片全部本次新建目录出 exe、0 error C / 0 LNK。哨兵 `scripts/check_bodydecl_shape.ps1` P1..P5（当前绿 defs 1 / callsites 4 / 手写展开 0 / raw wraps 1 / 语义层 MultiDecl 0；HEAD 树上 P1..P4 十条红，P5 在 HEAD 也绿 —— 它是"保持为 0"的哨兵不是抓本次的），已接进 compile 那片 `[STATIC] bodydecl_shape`。 **门 #338 落定**：门 #338 = run 37321722861、head bab5b0ef、attempt 1 = 11 job 全 completed/success；compile 片 25→26 里新那枚就是 `[STATIC] bodydecl_shape ... PASS`（逐行读到），syntax 片 151→152 是 `[CODEGEN-NOTE] bodydecl_one_per_declarator ... PASS`，bas 两片之和 90→92 = test_bodydecl 与 test_bodydecl_x86 进了门禁且绿；vbp #3 那条 SKIP 仍是 test_vbman（COM 未注册，与 #335/#337 同形）。登记面这轮**先按字节断言再跑**（上一格门就红在 Add-BasTest 的路径里混进制表符），109 条注册 ctrl=0 / missing=0、PSParser=0 错。**边界写明**：语义层 Dim 那支仍内联造符号没并进 registerVariable（不在本格问题面）；`cgen_localdecl.cpp:31` 的 MultiDecl 分支从此 unreachable-by-construction、本轮没删；VBFlexGridDemo 还剩 **502 条 VB3001 全在另一族**（按名字：UserControl 278 / PropertyPage 127 / VBA 37 / Extender 32 / Ambient 7 / vbSrcCopy 5 / vbPicTypeIcon 5 / Is 3 / vbPicTypeBitmap 3 / vbPicTypeEMetafile 1 / CTRLINFO_EATS_RETURN 1）⇒ **另立账 #217**（.ctl/.pag 宿主词汇与 VB6 内在常量，与 #159/#176/#179 同族），本轮一条没动。下一格候选：#217（本轮新立）、#192 剩下的 RTL 半边、#193、#206 与 #208（要 VB6/tB 本人仲裁）、#213。
# 上一轮（2026-10-05T20:40:00+08:00）= 本轮追加：**账 #216 已出（提交 `1df66f09` + 登记修复 `332bfd14`，门 #337 全绿）= 位运算（And/Or/Xor/Eqv/Imp）与一元 Not 的结果类型收成 TypeSystem 一处权威** —— 起因是账 #214 那条路上顺手量到的形状：位运算数一进字符串上下文就打 True/False（探针 `.build/pcprobe/`，改前 `o_base.out` → 改后 `o_new.out`：`"x=" & (34 Or 51)` 由 `True` 变 51、`Left(a Or b,2)` 由 `Tr` 变 51、`"x=" & (Not 5)` 由 `True` 变 -6），而**同一条表达式先赋给 Long 变量再打印一直是 51** ⇒ 缺的不是算（Fix 039 早把两侧化成 int32 再 `& | ^`），是"问类型"那一步答错；答错外溢成三件事：字符串上下文选 `vb6_CStrBool`、装箱走 `vb6_VariantBool`、COM 实参走 `vb6_ComPackBool` —— 第三条在真工程里是响的，`tests/VBFlexGridDemo/MainForm.frm` 的 `Render(hDC Or 0, X Or 0, ...)` 一直在把设备上下文句柄按 **VT_BOOL** 交给 COM。**结构上的关键**：这条决定在仓里写了两份 —— 语义层那份（`semantic_analyzer_expr.cpp`）从一开始就是 VB6 的口径（两侧都 Boolean 才 Boolean，否则数值提升，否则 Variant；Not 也分开答），发码层那份（`cgen_util_type.cpp` 的 `inferExprType`）对 `And/Or/Xor` 与 `Not` **无条件** return Boolean，而 `Eqv/Imp` 更漏进算术支路；两份各写 ⇒ 修一份必留另一份，所以这一刀不是"再补一条规则"而是**把第二份删掉、两层四个点全调一处**（`TypeSystem::bitwiseResult` / `TypeSystem::logicalNotResult`）。`Not` 那一档定的是 Boolean/Variant 跟着操作数走、Byte/Integer→Integer（答 Byte 会让 `b = Not b2` 绕过溢出检查、把 -1 静默 wrap 成 255）、Long/LongPtr/LongLong/ULong 原样、浮点与 Currency/Decimal→Long、其余 Variant；语义层原先 `Not Byte`→Byte、`Not Double`→Variant 两条随收口一起改（带动的发码实测为 0）。**判据** `tests/test_bitops.bas` 二十四针、x86 与 x64 各真编真跑逐行相同（`.build/b220out/{x64,x86}.out`）：数值面 51/51/17/34/-6/3/-1/51/51/80/95/-86/-2/-3/-7/-5/51/85，布尔面 `BF-bool-or=True / BF-bool-and=False / BF-boolbox=False / BF-cond=hit / BF-boolcond=miss` 一并钉住（拦"一路改回去全推成数值"那个反向错）。**负控**=把两处权威毒成"恒 Boolean"（就是改前那份答案）重编一台 ⇒ 16 条数值针全变 `True`/`Tr`（含七条新针），而五条布尔面针与 `BF-assigned=51` 一条不动 ⇒ 夹具两头咬得住；还原后 24 条读数与 x86≡x64 复核通过。**发码面 A/B** 100 份 `--emit-c`：changed=8、census `vb6_CStrBool(` 162→142、`vb6_VariantBool(` 564→556、`!= 0) ? 0 : -1` 与 `vb6_VariantToLong(` 不动。最大那份 VBFlexGridDemo（95 块、+170/−165）的归因换了方法：先把 `_vcmp_`/`_vb6_with_`/`_ndoff_`/`_tmp`/`_vb6_select_` 归一，再按**整份文件里每行的出现次数**比 —— **不能按 diff 块**（块里只要有一行带关键字就当整块有解释，那是假归因，第一次就这么漏了 157 行）。次数有变的行 36 条逐条读，全落在三类且没有第四类：① 收窄目标上的位运算赋值补范围助手（`BufferVT`→`vb6_ChkInt` 16+16、`vb6_ret_CalcHash`→`vb6_ChkLong`、`KeyCode`、`VT`、`LoWord/MakeWord/LoByte`）；② `(X And k) = k` 那一族从内联 `== k` 改走 `vb6_VarCmpLongEq` + 一枚装箱临时，每片净 +5 行正对上那五条新声明；③ COM 实参 `vb6_ComPackBool(hDC Or 0)`→`vb6_ComPackInt(...)` 三条。**第二次 A/B**（专验"收成一处"本身）：权威版与手搓那版 100 份**逐字节相同 100/100**。真编译 grid/charts × 两位数四片全过（`.build/b220out/demo/`，输出目录本次新建 ⇒ 只认这次跑出的 exe：四片各 `error C`=0、`LNK`=0、exe 1.67/2.03/0.94/1.15 MB）；记一条读数纪律：那台 cmd 把 `echo GRID32_BUILD=!B!` 的 `!B!` 原样打了出来（延迟展开没生效）⇒ **这一片的 rc 不是判据**，别拿它当"编过了"。**哨兵** `scripts/check_bitwise_authority.ps1`（B1 定义各 1、B2 声明各 1、B3 两层四个调用点合计 ≥4、B4 两份消费点里不许再手写作答），当前绿 `defs 1+1, callsites 4, inline answers 0`；负控=把那两份消费点退回 HEAD 的版本（`.build/b220_headtree/`）跑同一枚 ⇒ B1..B4 全红且 B4 直接点名改前那两条 （`BinaryOp::And || Or || Xor) return Vb6Type::Boolean`、`UnaryOp::Not) return Vb6Type::Boolean`）；已接进门禁 compile 那片 `[STATIC] bitwise_authority`，用例登记走 x86+x64 两形、PSParser=0 错。台账：§B48 一节 + §D 一行（`git diff --numstat` = +23/−2，两条定点回填各命中一次），并把 §B47 与 §D 里账 #212 的「门 待回填」补成 **门 #335**（run 37293132648、head 0364dfbb、attempt 1 = 11 job 全 completed/success；四片 vbp 的 TOTAL 之和 224→226 正是 pberr 与 pberr_x86 进了门禁且绿，vbman 那枚 SKIP 随分片从 vbp #1 挪到 vbp #3；`[STATIC] sa_access ... PASS` 逐行读到）。**门 #336 那次红红在自己的登记行**（两条 Add-BasTest 的路径落盘成了 `tests` + 制表符 + `est_bitops.bas`：生成脚本里的 `\t` 被工具层折一级后成了转义，制表符还吃掉一个字母 `t`），两片 bas 在 11 秒时被仓里那枚注册表自检 `[FATAL] 路径里有控制字符` 打掉；三条纪律一起记：`[PSParser]::Tokenize = 0 错` 拦不住真制表符、把制表符换成反斜杠的"第一次修法"仍少一个字母（要按字节断言目标串），而**本地只直接跑夹具、不经过注册表那条路就会漏这一类**。门 #337 = run 37313706942、head 332bfd14、11 job 全 completed/success：bas 两片之和 88→90（test_bitops 两形进门禁且绿）、compile 片 24→25 里新那枚是 `[STATIC] bitwise_authority ... PASS`、pberr/pberr_x86 在本 head 复绿；#336 同轮的 `tabwalk ... FAIL` 归到已登记的 #213 节拍余量 —— 它的 `--emit-c` 产物在改前基线/手搓版/权威版三份里逐字节相同（27303，x86+x64），#337 上 tabwalk 与 tabwalk_x86 两片复绿。**边界写明**：`exprYieldsVbBoolean`（`cgen_expr.cpp`，问的是"这条表达式产出的是不是 VB 布尔"、服务于 Not 的发码形状）没并进这处权威 —— 不是同一件事；一元负号 `Negate` 那两份仍不一致（语义层 Boolean→Integer、发码层透传），这次没碰（没测到消费者）；`Not 2.5` 那种银行家舍入只钉了整数值那一条；String 参与位运算的运行期错误号（VB6 是 13）与 `Eqv/Imp` 的溢出行为都没问。下一格候选（都没开工）：**#215**（过程级 Const 的多声明符行：本轮已把前提改正是**一枚都不登记** —— `Dim a, b` 在体级出两条 LocalDecl、`Const A = 1, B = 2` 出一条带 MultiDecl 的，而语义层 `visit(LocalDeclStmt)` 的 switch 没有 MultiDecl 分支（发码有），所以代码照发、值照对、只有每条使用报一条 VB3001）、#192 剩下的 RTL 半边、#193、#206 与 #208（要 VB6/tB 本人仲裁）、#213。
# 上一轮（2026-10-05T17:52:00+08:00）= 本轮追加：**账 #212 已出（判据夹子 `tests/pberr`，门 #335 全绿）= 抛 9 这一跳在实例方法里两面都通，产品一字未动**：接 #209/#214 —— 越界元素访问改成抛 9 之后，VbEclipse 那批调用点全在 UC 的**实例方法体内**（`With m_Serie(Index)` 发的就是 `&(me->m_arr(k))`，取成员数组那一格当对象），而当时只量过「没处理器 ⇒ `Unhandled error 9` + 进程按 9 退出」这一面。另两面一直没量过：① 方法自己写着 `On Error GoTo` 时 9 会不会**越过**它落到调用方（或干脆落不回）；② 跳走之后那枚实例的状态还读不读得出（longjmp 没解链实例栈的话，现场正是「错误被吃掉 + 对象已坏」那一族）。实测两面全通：x86 与 x64 真编真跑**逐字节相同** `PE-1d=9 / PE-2d=9 / PE-null=9 / PE-caller=9 / PE-state=20-40 / PE-DONE`（两台 BUILD rc=0 / RUN rc=0，`.build/b219out/{x64,x86}.out`）⇒ 落回自己那枚方法的处理器、没处理器时继续交给调用方、三次抛出后 `m_items(2)/m_items(4)` 照旧读出 20-40。**三枚负控都是改夹具、不改产品**，每枚都必须让一条登记过的针变红：A 把 `OneDimAbove` 的下标 9 换成范围内的 2 ⇒ `PE-1d=20`（针红）；B 把 `Class_Initialize` 里 `m_items(4)` 置 0 ⇒ `PE-state=20-0`；C 删掉驱动里的 `On Error GoTo Caller` ⇒ 传播上来的 9 没人接，stderr 打 `Unhandled error 9: Subscript out of range`、进程 **RUN=9** 退出、`PE-caller`/`PE-state`/`PE-DONE` 整片消失（顺带把 #209 的未处理口径在实例方法链上又验一遍）。**一条负控设计的坑**：把处理器前的 `Exit Function` 拿掉**不算**负控 —— 抛出发生时根本走不到那一行，实测输出照旧；要红就得让抛出根本不发生（A 那种改法）。登记：`$pbErrExpected` 六条针 + `Test-Vbp `pberr`/`pberr_x86`，插在 `erase_sub_x86` 那条**整语句之后**（按上一轮那条新约束），复跑 `[PSParser]::Tokenize` = 0 错、`git diff --numstat` = +14/-0；夹具三份源文件从工具写出的 LF 归一成 CRLF 之后**重编重跑一遍**，六条读数一字不变（归一化也算改动）。边界写明（§B47）：钉的是工程类实例方法那一路，不是 UC 宿主带 `Extender`/`ScaleMode` 那一路；只断言 `Err.Number`；`On Error Resume Next` 那一形没问；用例级 PASS 行只能由 CI 给（本地没有跑单条 vbp 的入口）。另记一条工具事实：`git remote -v` 会把 URL 内嵌的 token 原样打到屏幕上（本轮不小心走过一次，已在 §B46 末段写下「取 remote 只用 ls-remote / config --get 且只回显 sha」）。下一格候选（都没开工）：#213（tabwalk 相 3 只留 1 拍余量 —— 已量过、复现不出，按纪律不动夹具）、#192 剩下的 RTL 半边、#193、#206（单位口径要 VB6 本人仲裁）、#208（同样要 VB6/tB 口径）。
# 上一条（2026-10-05T17:35:00+08:00）= 本轮追加：**账 #214 已出（提交 `323ab077`，门 #334 全绿）= 多维元素访问收成一处带检查的出口，顺带补上 4 秩那一条只塞两枚下标的发码**：门 #334 = run 37287453662、head 323ab077、attempt 1 = 11 job 全 completed/success，逐片 FAIL=0（smoke 1/1、compile 24/24、asm 13/14、bas 44/44 两片、vbp 四片 48+1skip / 54 / 50 / 50、syntax 151/151），`[STATIC] sa_access ... PASS` 逐行读到；bas 两片之和 87→88 = 本轮新用例 `test_arr_nd` 进了门禁且绿（那片仍只落摘要行，逐针读数只有本地 x86+x64 那份）。三条本地读数（两台逐行相同，`.build/b216out/ndfix{32,64}/run.out`）：2 秩越界以前是**静默拿隔壁那格**（`a2(3,1)` 回 12，VB6 是错误 9 —— 本族最坏的一种，不崩也不响）、从没 ReDim 取 `d(1,1)` 读 (NULL)->data 当场 AV、而 4 秩以上 `Dim a4(1 To 2,×4)` **在范围内也**崩 —— 根因不在 RTL 在发码：`cgen_expr_call_prelude.inc` 那一支把实参拼成 `(int[]){i0, i1}` 只塞两枚下标却按 `actualDimCount` 交给 `vb6_SafeArrayND_Offset` ⇒ 读的是 `indices[2]/[3]` 那两块栈上垃圾（1/2/3 秩各有专用分支所以一直是对的；98 份产物里 `_ndoff_` 计数 0，所以这条洞一直没被走到）。做法照 #209 那一形状：`vb6_SaNdElemPtr` 一处 inline（先问形状 `dimCount`∈1..16 —— 一维描述符首字段是魔数 0x5A1D 天然落不进去，阈值抄的是 `vb6_LBoundND` 已有那条 —— 再逐维问上下界）+ `vb6_SaNdElemFail` 抛 9；AT1/AT2/AT3 宏体改走它、新加 ATN 一次交全秩数与全部下标；复合字面量**必须裹最外层括号**（预处理器按顶层逗号切宏实参，花括号挡不住，少了就是 cl C4002，实测）。护栏：A/B 98 份 `--emit-c` 逐字节相同 + census ND_AT2 1468→1468 / ATN 与 `_ndoff_` 各 0→0（发码文本没动 ⇒ A/B 只证“没别的跟着变”，那 1468 处“编得过”另由真编译 grid/charts × 两位数四片 rc=0 证）；哨兵 A4 口径从“恰好 1 条兜底行”改成 **0 条**，新加 A6（多维那一支的唯一入口）与 A7（发码形状），**七条负控逐条真红**且每轮逐字节还原、两个壳都绿。**一条读数纪律**（下轮别再上当）：`Charts 2020/Form2.frm:504` 是 `Randomize Timer` ⇒ “哪几枚 UC 宿主对 hover 有反应”**不能当判据** —— 同一枚**老** exe 三趟给 [5,6,7,8,9] / [4,5,6,7,8,9] / [5,6,7,8,9]，基线像素 MD5 每趟都不同；该问的换成“新产物有没有走进这一刀”：新 charts 三趟 + 新 grid 两趟带 `C3_SA_TRACE=1` 真跑`saNd=0 / sa1d=0 / unhandled=0`、进程一条没丢 ⇒ 那 1468 处全在范围内，加检查无可观察行为变化（“图真画出来了”那一面本来由门禁的 `Charts2020`(-DumpMinColors 40) 与 `FlexGridX86` 钉着）。工具事实两条：① 往 `run_tests.ps1` 插用例只许插在**整条语句之后** —— 本轮把 5 行插进上一枚多行 `Add-BasTest` 的实参中间，`+5/-0` 看着是纯新增却把语句劈两半，`[PSParser]::Tokenize` 一量就有错，改完 0 错才算过（CI 侧旁证：那两片会直接 ParserError）；② `git remote -v` 会把 URL 里内嵌的 token 打到屏幕上，取 remote 改用 `git ls-remote` / `git config --get remote.<n>.url` 且只回显 sha。**下一格 = 账 #212**（口径已写进 §C3 台账 §B46 末段）：#209 与 #214 都把错误 9 交回运行期，而 VbEclipse 那批越界点全在 UC 的**实例方法体内**，所以要补两面 —— 方法体内写了 `On Error GoTo` 的 9 有没有落回它自己的处理器、抛过之后那枚对象的成员还读不读得出数（`With` 那一族现场就是取 `&(me->m_arr(k))` 当对象）。探针已写在 `.build/pe212/`（`.cls` + `.bas` + `.vbp`），**还没真编真跑**，别当已验证。
# 上一条（2026-10-05T16:10:00+08:00）= # 本轮追加：**门 #333 落定（账 #209 那批全绿）+ 账 #211 收线 + 量出一枚新缺陷**：门 #333 = run 37278620691、head d61d9065、attempt 1 = 11 job 全 completed/success、10 片各自 FAIL=0，逐行读到的三条新判据面 —— `[VBP] pbsub ... PASS` 与 `pbsub_x86 ... PASS`（加固后的夹具在 CI 上两头都绿）、`[STATIC] sa_access ... PASS`（#209 的哨兵）；bas 两片（44/43，各 FAIL=0）的日志里没有 `test_arr_empty` 那条用例级行，所以那 21 针的逐针读数只算本地那份。“门 #332 那条红”的归因（账 #211）：pbsub 的五条通知是 PostMessage 异步发的，夹具在固定第三拍读计数 —— 本地带 6 个 CPU hog 复现 2/24 全 0，改成“到齐才读、最多 20 拍”后 48 趟全绿、其中一趟到第 8 拍才到齐；先把 #209 摘干净（pbsub 的两份 `--emit-c` 里 `VB6_SA_AT`/`SafeArray` 0 处、坏跑 stderr 空），负控=注掉五句 PostMessage ⇒ 等满 20 拍读出全 0 + `done=False`。同族普查：`tabwalk` 相 3 也只留 1 拍余量，但 21 趟带载复现不出 ⇒ 不动夹具，只记台账。**新量到的一枚（还没开工）**：4 秩及以上的数组下标寻址当场 AV —— `Dim a4(1 To 2, 1 To 2, 1 To 2, 1 To 2)` 写读一轮 RUN_RC=0xC0000005 且一条输出都没有（x86/x64 同形），而 1/2/3 秩各 0 错位；根因在发码那一行 `cgen_expr_call_prelude.inc:348`：4+ 维那条兜底把实参写成 `(int[]){i0, i1}` 只塞了**两个**下标却按实际秩数传给 `vb6_SafeArrayND_Offset` ⇒ 读越界的 indices[2]/[3]。存量语料里 0 处（4 份工程只用到 2 秩），所以不挡 VbEclipse，但这是编译器自己的洞。下一刀把 ND 那一支（`VB6_SA_ND_AT1/2/3` + 这一行）一起收成**一处带检查的出口**：秩数不定就交一个 indices 数组进去，越界/NULL 一律抛 9（与 #209 同口径）。
# 上一条（2026-10-05T14:45:00+08:00）= # 本轮追加：**账 #209 已出（提交 待回填，门 待回填）= 一维动态数组的元素访问收成一处带检查的出口**：交互探针（b213）量到 ucChartBar demo（x86，`b211out/bar32`）点 Random 之后窗口消失，三轮 Windows 应用日志同一条 `0xc0000005 / 偏移 0x1abf2`；同一份工程 `-g --keep-for-debug` 重编 + `C3_CRASH_TRACE=1` ⇒ `av read target=0xc`，符号化到生成码 `ucChartBar.c:591` 的 `With m_Serie(Index)` ⇒ `VB6_SA_AT`（`vb6rtl_array.h:100`，全仓唯一的一维描述符裸解引用点）对“未分配 / 越界”一个都不问。修法：宏体走 `vb6_SaElemPtr`（header 里 inline，三条比较）+ `vb6_SaElemFail`（冷路径抛错误 9），与 UBound/LBound 那两条 rev2 同口径；步长仍按 `sizeof(type)`，没碰 Fix 170/rev3 那条 elemSize 的坑。**顺带把 #208 那条判掉坐实**：demo 的 `Form_Load` 只设四枚组合框的 ListIndex 就 `Exit Sub`，四枚 `_Click` 从没进 ⇒ `ReDim Preserve m_Serie` 没跑；真 VB6 在这种状态点 Random 同样弹错误 9、同样不画图，所以“点一下没反应”不是又一处产品缺陷。判据：`test_arr_empty.bas` 加七枚私有过程（未分配读/写、上/下越界、UDT `With s(2)`、Erase 之后 + 两枚**负控** `EA-ea-in=20/30/50`/`EA-ea-loop=100` 钉住范围内读写照旧），顺带补钉以前没钉的 `EA-err=9`；本地 x86+x64 各真编真跑 21 条 needle 齐。护栏：A/B 98 份 `--emit-c` 逐字节相同（`inputs=98 same=98 changed=0`，只动 RTL）；新哨兵 `[STATIC] sa_access`（`scripts/check_sa_access.ps1` 五条规则，两条负控真红）。同源 A/B：改前 `PROCESS-GONE exit=0xC000041D`，改后 `GONE code=0x00000009` + stderr `Unhandled error 9: Subscript out of range`。多维那一支（`VB6_SA_ND_AT1/2/3` + `_ndoff_` 兜底，存量 1468 处 / 4 份工程，VBFlexGridDemo 每位数 632）刻意留给下一刀；工具坑一条进 §C 第 14 条（PS 的 BOM 只认 EF BB BF + 无 BOM 时中文行尾吞 ASCII）。
# 上一条（2026-10-05T12:55:00+08:00）= # 本轮追加：**VbEclipse 四台的当前视觉基线立起来了 + 账 #207 已出（提交 `9e2c439a`，门 待回填）**。基线（x86 真编真跑 + `PrintWindow` 取窗口自己的像素，`.build/b208out/*.png`）：**Charts 2020 主窗体十枚控件全画**（ucChartBar1 柱状 + 月份轴、ucPieChart1 饼、ucChartArea1 折线 + 图例、ucTreeMaps1 分块图带名字、"Venta diaria USD $532.00"、ucChartBar2 横向条、ucChartArea2 面积、ucPieChart2 环形带百分比、LabelPlus 三枚环）；**czUI demo 全画**（Data Traffic / Kill Switch 开关 / 65% 进度条 / Search / Connect / Location）；**VBFlexGrid demo 全画**（日期列无 0:00:00 尾巴、每行数值不同、CellPicture 预览在）；**ve_units / ve_list 自退 exit=0**（ve_units 的 U-* 判据全 True；ve_list 交的是数据面，没有 True 判据行）。**账 #207 = 设计期 .frx 的 List/ItemData 只接了 ListBox 一档**，ComboBox 那 17 处全落空 ⇒ 下拉框是空的，而 `ucTreeMaps/Form1.frm:357` 那句 `If Combo1.ListIndex = -1 Then Combo1.ListIndex = 4: Exit Sub` 让 Form_Load 第一句就退 ⇒ **图体空白**。改法是不新开机制：同一处出口按控件型取消息对（`CB_ADDSTRING`/`CB_SETITEMDATA`），ListBox 那一路文本逐字节不变，两条创建路本来就共用它。读数：夹具 `tests/frxdata` 加一枚 ComboBox **复用同一份 blob 的两个偏移** ⇒ 改后 `FD8-COMBO=True`（3 / 1234 / 5 / -7 四头对上），**改前那台真跑 = False + count=0**（`.build/b211out/b64.out`）；真工程：ucChartBar demo 四个组合框都答得出设计值，ucTreeMaps demo 点 Random 之后树图整片画出（分块 + 图例 2000..2004）。A/B inputs=98、same=86、changed=4、0 删除 0 无法归因。**新账 §B43/#208 已判掉（同一天）**：照"程序改 `ListIndex` 该发 Click"去补发通知，两头判据绿、ucTreeMaps 也自动画了，但 `tests/ctrlfiles` 的 CF14/CF15 当场红 —— 那两条是夹具**手工调用 `fileList_Click` 模拟点击**的，说明这套夹具一直按"不发"写。本机没有 VB6 可量 ⇒ **不凭猜改产品**，RTL 与夹具全部回退并复跑确认（ctrlfiles 17 条 needles 两台全在）。重新定性：那一页"要点 Random 才画"大概率就是 VB6 原样，不算缺陷。**两条工具课写进台账 §C 第 12/13 条**：`CopyFromScreen` 会截到别人盖在上面的窗口（改用 `PrintWindow(PW_RENDERFULLCONTENT)`）、PowerShell 的 `DllImport` 不带 `CharSet` 时按 Ansi marshal ⇒ `GetClassNameW` 只回一个字符（看着像产品截文字，其实是探针）；另外 `tests/frxdata/FrxData.frm` 是 GBK，改它必须按字节插 ASCII。开着的还有：账 #199 的甲/乙口径（仍等用户拍）、#206（状态条面板宽度口径）。
# 上一条（2026-10-05T12:15:00+08:00）= # 本轮追加：**门 #329 落定（账 #196 第三条那批全绿）+ 账 #205 已出（提交 `a4986c88`，门 #330 = run 37262853393、head 14caaf2a、attempt 1 = 11 job 全 completed/success、逐片 FAIL=0，`[VBP] sbfont` 与 `sbfont_x86` 两条在 CI 上真 PASS，邻居 `[VBP] ctrlstatusbar ... PASS` 没动）**：门 #329 = run 37260903318、head 511d356e、attempt 1 = **11 job 全 completed/success**（含 Build C3.exe），但**用例级 PASS 行这轮也没拿到** —— jobs 数组里没给 `log_url`，改走 `/actions/jobs/{id}/logs` 拿到 10/11 份，取回的正文里一条 `[VBP]/[CODEGEN-NOTE]/[STATIC]` 都没有（本机中转对日志端点给的不是真日志正文，重试到后面连 jobs 也顶成非 JSON）⇒ 台账里不写任何"CI 逐行真 PASS"的句子（§B31 与 §D 两处都按这个边界写清了）。**账 #205（状态条那两处问字体）**：动手前先读了另一位作者的期望表，**判据的形状因此改了** —— `tests/ctrlstatusbar` 那 42 条里与宽有关的 SB10-W2=120 / SB36-SETW=123 钉的都是**显式给过的 Width**，而 `vb6_StatusBar_GetPanelWidth` 返回 `e->width`（请求值）⇒ 排版结果在 VB 侧没有现成的门，于是判据去**问窗口本人**：`SB_GETPARTS`(WM_USER+6=1030) 交回各格右边界（Declare 用 `LongPtr` + `ByRef As Any`，同 LabelPlus 那一形）。读数（两台逐行相同）：改前 `ct8=140 ct20=140 after=140` = 同一串文字在 8pt 与 20pt 两枚上量出**同一个宽**、运行期改字号也不动；改后 `ct8=110 ct20=260 after=260`；证人 `pf8=11 pf20=27` **改前改后都一样** ⇒ 设计期那两张字体本来就下发到位（#204 存的），只是没人去问。判据三头（新夹具 `tests/sbfont` 两台各一条）：SF02 两枚同串不同字号不等宽、SF04 运行期改字号后重排变宽并与 20pt 那枚对上、SF05 显式给过 Width 的那格右边界不许挪；RAW 三行只钉前缀。**负控 = 同一份夹具在改前那台真编真跑** ⇒ SF02/SF04 都 False（`.build/b207out/x64.out` / `x32.out`）。护栏：D11 那条**具名豁免按设计自己消失**（`$rawAll` 从 want 3 收成 want exactly 1，PASS 打印 `全仓裸问 1`）；只动 RTL ⇒ A/B 86 份**逐字节相同**（diff=0）；相邻三枚哨兵（di_stubs 624=624 / uc_scale_units / host_pseudo_table 55 行）各自复跑全绿；**别人那一族 42 + 10 条判据逐条复核两台全在、一字未动**（`.build/b207c_check.py` 就地从 run_tests.ps1 抽期望表）。没验的一头写明：`C3SbPaint` 那一遍没有独立数值判据。**新账 §B41/#206 开着**：`Panels(i).Width` 交回的是请求值而不是排版后的宽（sbrContents/sbrSpring 在 VB 侧读不到几何，实测改前读回 0）—— 刻意没改，因为 VB6 那一读数的单位口径（缇 vs 像素）还没量准。开着的还有：账 #199 的甲/乙口径（仍等用户拍）、账 #196 已全部出完（hDC / TextHeight·TextWidth / ScaleX·ScaleY 三格）。
# 上一条（2026-10-05T11:46:00+08:00）= # 本轮追加：**账 #196 第三条已出（提交 `479b202f`，门 待回填）= `ScaleX`/`ScaleY` 换算那一站**：上一轮预告的「缺的出口 = 带 HWND 的 `vb6_WindowScaleX/Y`」**被普查推翻** —— 调用点 = Charts 2020 五份 .ctl 各 4 处 + `PropPagFMR.pag` 2 处（=22）+ `VBFlexGrid.ctl` **94 处全是 `UserControl.` 前缀**（Fix 111 早就通了）+ `MainForm.frm` 2 + `Common.bas` 2 + `archive/ctxWinsock.ctl` 2，实参只有三种形状 —— 字面 `vbPixels(3)` / `vbContainerSize`·`vbContainerPosition` / `.ScaleMode`·`Me.ScaleMode`·`UserControl.ScaleMode`（#197 下发后答 1 或 3），**0 处传 0=vbUser** ⇒ 发的是**不带 HWND** 的那一条：RTL 新增 `vb6_ScaleUnitX/Y(x, from, to)`，`vb6_UserControl_ScaleX/Y` 改成**薄壳直接 return 这一条**（Fix 111 那一档与窗体型这一档不可能分家，也不是在 #177 那张单位表旁边抄第二遍）；后端一处权威表 `controlScaleMethod`（只登记 Form/PictureBox，不给通用行）+ 三处码头各查一次（With 形 / 带括号裸形 / 裸名），裸名那一处把改写门从 `isDesignerModule_` 放宽到 `isFormModule_ || isPropertyPageDesigner_`（§B31 开头那格「`.pag` 里裸 `ScaleX(...)` 认不出来」就是它；`.ctl` 里裸写那一形没动，仍走 kHostPseudoRows 的 UC 两行）。**没接的一头写明**：传 0=vbUser 时 VB6 要的是这枚窗口自己的用户坐标系，而 `vb6_ScaleUnitsPerPx` 的既有口径把 User/Container/unknown 都按像素 ⇒ 那种调用静默给恒等值；语料 0 处 ⇒ 记着不装绿。**判据** SX10 = 四形逐数相等（`picB.ScaleX` / 裸 / `Me.` / `With .`）+ `ScaleY` + 一枚问窗口的证人 `GetDeviceCaps(LOGPIXELSX)`，进 `$dcSurfExpected` 两台各一条；**两台真跑逐行相同** = `SX11-RAW pic=96 bare=96 me=96 with=96 y=96 dpi=96`；负控 = #196 那台 BASE 真编今天这份夹具 ⇒ exit 2、C2039×3 + C2440×7、诊断面三条 `VB4001 Unknown control property`（hDC/TextHeight/ScaleX）全在。**针两条** `dcsurf_scale_units` + `scale_units_real`（后者钉真工程 ucTreeMaps 那一行）。**哨兵 D12**：假 needle 真红过（注掉 With 那条码头 ⇒ 点名 `cgen_expr_with.cpp`），而它第一枪红的是我自己写的规则（「恰好 1 处」漏了表自己那一行也是字面量 ⇒ 真数 2）。**A/B 86 份**（BASE = `b204_new_emit`）⇒ changed=6、**OFFENDERS=0**：ucTreeMaps 两台各 1 块 + **VBFlexGridDemo 两台各 1 块 2 行**（`MainForm.frm` 那两行布局以前发成「拿 HWND 当 IDispatch 问 ScaleX ⇒ Empty、再按数值打」，现在答得出数；这两份一开始被我当 offender 抓到，查实是正当改道后补进 WD）+ dcsurf 两台各 4 块（全新增行）。**结构性断言**：86 份新产物里裸 `ScaleX(`/`ScaleY(` 作调用出现 **0 处**。真工程：ucTreeMaps 两台 rc=0、x64 exe 660,480→**660,992**、诊断面 `Unknown control property` **归零**（仅剩那条 VB4001 是 TypeLib path）。**本账三格（hDC / TextHeight·TextWidth / ScaleX·ScaleY）到此全部出完。**开着的：账 #199 的甲/乙口径（仍等用户拍 —— 页里的控件从没被创建 ⇒ 那一页的绘图面至今白画）、§B40/账 #205（状态条那两处被 D11 具名豁免着，动之前先读 `tests/ctrlstatusbar` 的判据口径）、账 #203 已按既有口径判掉（工程自己的 `As Long` 截指针 ⇒ 不追 x64）。
# 上一条（2026-10-05T10:35:00+08:00）= # 本轮追加：**账 #202+#204 已出（提交 `8252b080`，门 #328（run 37255522248，head ed66b03f，attempt 1）= **completed / success**，这条结论今天独立读到两次（watcher 的 poll 17 与事后复核各一次）；**逐 job 的 FAIL=0 与用例级 PASS 行没拿到** —— 本机那台中转在取 jobs / 日志这一段反复把 JSON 顶成 HTML（11 片日志只落下一片），所以下面不写任何"CI 逐行真 PASS"的句子。这一批的风险点恰好是「改了 shape/widget 两族共用的字体问法」会不会动到别人的判据，那一格只有逐 job 读数能答；中转恢复后按 §D 这一行补记（补不上就按 #328 的 overall success 结，边界写清楚））= 字体那一处出口的覆盖面补齐**：`vb6_ControlFont` 从 static 提到 `vb6forms_internal.h`，新增唯一写口 `vb6_ControlFontStore`，六处站点接上（groupbox 标题带读 / 控件创建路存 / 控件数组问模板+存 / shape 三处文字量 / widget 的 Label AutoSize）。根因仍是探针那一条：STATIC **与** BUTTON 两类窗口收到 `WM_SETFONT` 后都不答 `WM_GETFONT` ⇒ "只发不存"就等于谁都问不到。**读数**（两台逐行相同）：从没被写过字体的 picB 从 `name= fs=0 pf=0 th=16 / twip=240` 变成 `name=MS Sans Serif fs=8.25 pf=11 th=13 / twip=195` ⇒ 默认字体那批控件的 TextHeight / Print 行距 / 缇换算不再按 Segoe UI 9pt 算。判据 FR01 三头钉 + **假 needle 真红过**（注掉那一句 Store 重编一台 ⇒ FR01=False、旧数全回来）。顺手抓到 setter 的起点问题：改 20pt 后 tB2 27→32，因为现在从真的那份 LOGFONT 起步（改前等于把字体族也换了）。护栏新增 D11（全 src/rtl 裸问 = 1 + 状态条两处**具名豁免**，豁免一旦删就红；存字体站点 3 且分属三文件）。**两头写明没验**：shape/widget 那四处无数值判据；数组那路 `vb6_CtrlArr_Load` 全仓零调用者（发码侧没有 `Load <数组>(n)`）⇒ 接进来是口径不是产品修复。A/B 86 份 changed=2、OFFENDERS=0；ucTreeMaps 两台 rc=0、0 error，x64 exe 659,968→660,480。**新开 §B40/账 #205**（状态条那两处，动之前先读 `tests/ctrlstatusbar` 的判据口径 —— 别人的族）。开着且下一步就等的：账 #203（ucTreeMaps x64 启动期 AV，要先符号化）、账 #199 的甲/乙口径（仍等用户拍）、#196 剩下的带 HWND 的 `ScaleX`/`ScaleY`。
# 上一条（2026-10-05T09:25:00+08:00）= # 本轮追加：**账 #200 已出（提交 `77ce2b65`，门 #327（run 37251283970，head 39394bcd，attempt 1）= 11 job 全绿、10 片各自 FAIL=0；CI 自己量的九条相关用例逐行真 PASS —— [VBP] dcsurf ... PASS 与 dcsurf_x86 ... PASS（两台各自的真跑判据面）、[STATIC] control_dc ... PASS（含这一刀新加的 D10）、[VBP-BUILD] charts_ucTreeMaps ... PASS (679,424 bytes) 与 charts_ucTreeMaps_x86 ... PASS (585,728 bytes)、四条形状针 dcsurf_dc_shape / dc_read_real / dcsurf_text_measure / text_measure_real 全 PASS；本批一次网络插曲：api.github.com 那台中转有半小时把 JSON 顶成 302→m.baidu / body="OK"，我一度把域名钉到 GitHub 真 IP —— 直连才是**一直**坏的，改回常规解析立刻通（已进记忆）；§B35 收线 + 新开 §B37/#202）**：控件窗口字体那一问归因到了窗口类本身 —— 一次性探针 `.build/b200probe/fontprobe.c`（裸 STATIC、谁也没子类化）实测 `WM_SETFONT` 之后 `WM_GETFONT` 回 NULL，`STM_SETFONT`/`STM_GETFONT` 同样回 NULL ⇒ STATIC 压根不记字体，所以 .FontName/.FontSize 的读、Print 落笔前的选字体、#196 的文字量这三条在 PictureBox/Label 那一类上永远按 DC 默认字体画（设计期 18pt 的 picA 与运行期改 20pt 的 picB 都量 16，`picA.FontSize` 却照读 18 —— 两头都不报错）。**改** = 一处出口 + 一份自存：新增 `vb6_ControlFont`（先问窗口、回 NULL 再读窗口属性 `VB6_CtrlFont`），setter 把自己创建那张存进这个新名字，三条读法全走这一处；旧字体刻意「先自存后问窗口」，反了会对真记字体的 EDIT/BUTTON 双删。顺带修掉一处 GDI 泄漏（改前 STATIC 每写一次字体漏一张，因为删旧字体那一句依赖的 WM_GETFONT 恒回 NULL）。**判据**：TH03 升回真判据 `ok7 = (tA > tB) And (tB2 > tB) And (pfA >= 18)`、`TH03-FONT=True` 进 `$dcSurfExpected`；真跑两台逐行相同 `TH06-FONTRAW a=29 b=16 b2=27 fsA=18 pfA=24 pfB=27`（证人从「整行不打」到 24/27）。**护栏**：哨兵 D10 四条，假 needle 真红过（换回裸 WM_GETFONT ⇒ 报 2 处并逐行点名）；**第一枪是自己造的** —— D10 落地那次报 3 处，多出的两条来自**行尾注释**里写的 `WM_GETFONT`（哨兵跳行首 //、不跳行尾注释）。**A/B** 86 份（BASE = #201 那份 emit 快照）⇒ changed=2、OFFENDERS=0，两份就是 dcsurf 两台且逐行都是夹具自己新增的行；注释定稿后重发两台 emit 与快照逐字节相同。真工程：ucTreeMaps 两台仍 rc=0、0 error、尺寸与 #201 同。**新开 §B37/账 #202**：探针补测裸 `BUTTON`(BS_GROUPBOX) **同样回 NULL** ⇒ `vb6forms.c` 里 Frame 标题带那处读法一并定罪（今天不红只是因为语料没人改 Frame 字体）；#199 的甲/乙口径仍等拍，#196 剩下的带 HWND 的 `ScaleX`/`ScaleY` 也还开着。
# 上一条（2026-10-05T08:55:00+08:00）= # 本轮追加：**账 #201 已出（提交 `4ced55cc` + 记账 `871a627f`，门 #326 = run 37247978838、head 871a627f、attempt 1，11 job 全绿、10 片各自 FAIL=0；CI 自己量的两条新用例真 PASS —— [VBP-BUILD] charts_ucTreeMaps ... PASS (679,424 bytes) 与 charts_ucTreeMaps_x86 ... PASS (585,728 bytes)）**：#196 两刀把 ucTreeMaps 的编译面清空之后，红点落在链接期三条 LNK2019 —— `AddFontMemResourceEx` / `GdipNewPrivateFontCollection` / `GdipPrivateAddMemoryFont` 三枚 DI 桩没人发。为什么生成器没发：那一形第三参在真源码里写的是 `ByRef DESIGNVECTOR`（无 As ⇒ Variant），正落在生成器「形状不定就跳过」那一档；GDI+ 那两枚当时那次会话没引用到。照口径**手写进 `src/rtl/core/di/vb6_di_stubs.c`**（手写那一档，gen_di_stubs 不动它 —— 重跑会按当次会话整文件重写、撤掉别人的桩），并把 check_di_stubs.ps1 的基线 620→624 冻住（以后再少任何一个当场红）。**一条刻意的取舍**：pdv 交 NULL，不把那个 vb6_VARIANT* 转手给 GDI —— VB6 那句 `AddFontMemResourceEx(.bvData(0), n, 0&, cnt)` 里 0& 的意思是「没有 design vector」，Win32 对这一意的拼法就是 NULL；把 variant 地址当 PDWORD 交进去，GDI 读到的是 VARIANT 的头两个字，那既不是 NULL 也不是 design vector，是伪造。gdiplus 那两枚走 LoadLibrary+GetProcAddress（gdiplus.h 是 C++ only，本仓 Declare 一律不要求导入库），解析不到回 GpStatus 2，与 gdiplus 族生成桩同形；`vb6_di_GdiplusStartup` 早已有桩、工程自己调 ⇒ 接上的是一条已经跑通的链。**结果**：ucTreeMaps 两台第一次真编真链出 exe（本地 659,968 / 562,688 字节），诊断面 0 error、116 条 VB3001 警告（Ambient / Extender / Is / PropertyPage 那一族 = 属性页与 extender 面还欠着，见 #199 与 #193）。**升格**：Test-VbpBuild charts_ucTreeMaps 与 _x86 进清单 （#187 那条口径的正解 —— 列进来这一次就是「从此不许退回编不过」），黑名单注释同步只剩 ucProgressCircular。护栏：这一刀只动 RTL 与清单 ⇒ 发码面该一字不变，A/B 86 份 changed=0、offenders=0。**新账 #200 还开着**（窗口字体那条：改了字号文字量照旧、`FontPixelHeight` 那句证人整行不打 —— 下一批第一步就是把证人折进 TH06 那一行做隔离实验）；#199 的甲/乙口径仍等拍。
# 上一条（2026-10-05T07:45:00+08:00）= # 本轮追加：**账 #196 第二条（按 HWND 的 TextHeight/TextWidth）已出（提交 `060eddca` + 针面订正 `ac3df329`，门 #324 = failure 但红在我自己的两条新针、门 #325 = run 37243854108、head ac3df329、attempt 1 = 11 job 全绿、10 片各自 FAIL=0；本轮七条相关用例逐行真 PASS —— [VBP] dcsurf 与 dcsurf_x86、[CODEGEN-NOTE] dcsurf_dc_shape / dc_read_real / dcsurf_text_measure / text_measure_real、[STATIC] control_dc）**。这一条**不需要发明机制**：现有 controlZeroArgMethod 那套「表只交名字、码头拼实参」本来就吃得下带实参的形 —— With 形交裸名 + pendingChainObj_，调用点（cgen_expr_call_com_bind.inc → cgen_expr_call_opt_pad.inc）把句柄**前置**进实参表。新表 controlOneArgMethod 只登记 textheight/textwidth × Form/PictureBox 两档（**不给通用行**），两条码头各查一次；实参签名表补两行 {void*, BSTR}（漏了就是把 vb6_VARIANT 裸喂进 GetTextExtentPoint32W = Fix 113 那一味）。RTL 新增 vb6_ControlMeasureTextPx：DC 走**同一处** vb6_ControlDrawDC、字体走 WM_GETFONT（与 Print 同口径，Fix 129），两个出口 float vb6_ControlTextWidth/Height 交回前过 vb6_ScalePxToUser + vb6_WindowScaleModeSelf（#175/#197 那份单位表）。判据三面：TH01 两形逐数相等且非零、TH02 **比值**判据（同字体的缇框 vs 像素框 >4 倍，实测 16px vs 240 = 15 倍，不钉绝对数所以 DPI 变了不假红）、TH04 宽度随文字变；TH05/TH06 只钉前缀当读数。负控 = 改前那台真编同一份夹具 ⇒ rc=1 无 exe。哨兵 check_control_dc.ps1 加 D6~D9，五条假 needle 逐条能红。**护栏** A/B 86 份（BASE = 这两刀之前那台）⇒ 被改 8 份、OFFENDERS 0，32 对读法替换 + 8 条诊断消失逐行归因。**真工程配对**：ucTreeMaps 两台（x64/x86）的 C2039/C2440 **全部消失**，红点从"编不过"推进到链接期 LNK2019 ×3 ⇒ 新账 **#201（§B36）**：缺 AddFontMemResourceEx / GdipNewPrivateFontCollection / GdipPrivateAddMemoryFont 三枚 GDI+ 桩 —— 要手写单桩（别重跑 gen_di_stubs），而且这三枚必须是**真实现**，否则只是从"编不过"换成"跑起来没字"。**新账 #200（§B35）**：改字号之后文字量不动 —— 运行期 picB.FontSize = 20（发码确实是 vb6_SetControlFontSize）之后再量还是 16，设计期 18pt 的 picA 也量 16，而 picA.FontSize 读回 18（自存的读得回、问窗口的没换）；更怪的是 FontPixelHeight 那句（问窗口的那枚证人）整条 Debug.Print 一行都没打出来。归因没做，所以**窗口字体那条判据今天当不了判据**，夹具里只留读数 TH06 —— 本刀的判据只钉得住"两形同归一处 + 单位真折算 + 宽度随文字变"这三头。**本轮两条工具课（已进记忆库）**：① 形状针"验证的那一串 == 登记的那一串"（门 #324 就是白红在这一条上）；② 哨兵的正则**一个函数抓一段**，别把相邻两个函数当一段 —— D9 第一版因此假绿了一次（改坏宽度那半时高度里的换算还在，整段照过）。**下一步**：#199 的甲/乙口径（页里控件从来没被创建）还等拍，#201 那三枚桩是 ucTreeMaps 出 exe 的最后一段，#200 是文字量唯一还没被运行期验证的那一头。
# 上一条（2026-10-05T06:35:00+08:00）= # 本轮追加：**账 #196 的 hDC 那一半已出（提交 `50676a8b`，门 #323 = run 37238929955、head f26f9590、attempt 1，11 job 全绿：10 片各自 FAIL=0，这轮连先前被本机代理换成 HTML 的 smoke / compile 两片都重下到了；本轮五条新用例逐行真 PASS —— [VBP] dcsurf 与 dcsurf_x86、[CODEGEN-NOTE] dcsurf_dc_shape 与 dc_read_real、[STATIC] control_dc）**：RTL 里「这枚控件的绘图 DC 从哪儿来」早就有一处口径（Print/Cls 在用），缺的是「句柄交得回 VB 代码」那一半 ⇒ 真工程那一形 `With Picture1 : TextOut .hDC`（Charts 2020/ucTreeMaps 的 PropPagFMR.pag:258）只能撞 "hwnd.成员" 兜底 = C2039。改 = 那条口径抽成 vb6_ControlDrawDC（Cls/Print/新出口三头共用），新出口 vb6_GetControlHDC 按 VB6 的「一个对象一张」把派发期那张直接交出、否则按 HWND 缓存进 VB6_ObjectDC 并当场归还白拿的那张，归还点在 PictureBox/Image 那层自己的 WM_DESTROY（#185 那套分层槽位）；后端读写两张表各补一行 hdc，只登记 PictureBox 与 Form（通用行 = List1.hDC 也答一个数 = 伪造成功，同 #192）。RTL 动过 ⇒ touch c3rtl.rc 重编 C3.exe。判据四面：夹具 tests/dcsurf 两片四头钉（反复读+With 三个数相等非零 / 两枚互不相等 / GetDeviceCaps>0 证明是一张活的 DC / GetPixel 各自等于自己那枚的设计期底色）；负控 = 改前那台真编 ⇒ DcForm.c(167) C2039、rc=1；形状针两条；哨兵 check_control_dc.ps1 D1~D5 （五条假 needle 按「人会怎么改坏」植，逐条能红、还原回绿）。护栏 A/B 86 份 ⇒ 被改 8 份、OFFENDERS 0，逐行归因 14 对 K2 + 5 条 WD（VB4001 那条「不认识 .hDC」的诊断消失就是这一刀的目的）。真工程配对：ucTreeMaps x86 的 C2039 2→1（只剩 TextHeight；TextOutW 的 C2198 本是它的级联，一起没了）。**一条按类推差点写错的读数（值钱）**：我原以为裸形 `picA.hDC` 以前走晚绑定那条 = 恒答 0（#143/#191 那一族），拿改前那台真编真跑一枚只发裸形的一次性探针（.build/b196probe，不进 tests/）量出来是 HDC=671159075 / DPI=96 —— 一张**活的** DC。所以这一刀在裸形上改的是语义（每读一次新取一张、从不归还 → 一个对象一张 + 销毁归还），真红的那一半是 With 形。⇒ 类推「兜底=静默空转」之前先编一枚最小的真跑。**新账 #199 已量到落点（§B34）**：`.pag` 的解析没缺（driver_frontend.cpp:172 对 .pag/.frm 同一条 FrmParser::parse，子控件进了 children），分岔在发码的模块种类那一问 —— cgen_base_generate_decl_pass.inc:39-46 那条分岔：isFormModule 且 frmDesc ⇒ emitFormFramework(...)，否则只按 frmDesc ⇒ emitDesignerControlDecls(...)，而 isFormModule 只对 .frm 置真 ⇒ 页永远走后者，只发槽位宏（`vb6_hwnd_X = (*vb6_UC_DesignSlotOf(me,"X"))`，槽位初值 NULL）、永不发 vb6_CreateControl。槽位语义对 .ctl（宿主窗体填）成立、对独立一页不成立，而属性页宿主 runtime 本仓没有。两条候选口径待拍：**甲** = 页也走 emitFormFramework（要一枚真窗口/承载体）；**乙** = 保留槽位、补一个「页装好时把创建出的句柄写进槽」的宿主。语料事实：PropPagFMR 在整个 VbEclipse 里只被 .vbp 列名、没有任何一处实例化。下一格照旧 #199（先拍口径）→ #196 剩下的两条（按 HWND 的 TextHeight/TextWidth 要一枚「一个实参的控件方法」机制，现有 controlZeroArgMethod 的三条码头都要各接：cgen_expr_with.cpp:72 / cgen_call.cpp:400 / cgen_expr_call_com_bind.inc:80）。
# 上一条（2026-10-05T05:20:00+08:00）= # 本轮追加：**账 #187 过门收线（门 #321 = run 37230353501、head 59811f14、attempt 1，11 job 全绿、每片 FAIL=0；六行 [VBP-BUILD] charts_* 在 CI 上真 PASS，CI 自己量的 exe 字节 x64 715,264/594,944/598,016、x86 613,376/519,680/520,704）** + **账 #197 已出（提交 `9565bdbb`，门 #322 = run 37235573381、head ada93533、attempt 1，11 job 全绿、11 片日志全部下到 (这轮代理没拦)、10 片各自 FAIL=0；本轮五条新用例在 CI 上逐行真 PASS：[VBP] scalemode 与 scalemode_x86、[CODEGEN-NOTE] scalemode_design_write 与 scalemode_read_real、[STATIC] scalemode_writers）= `VB6_ScaleMode` 这个窗口属性从此有写者**：读的一侧早就齐了（`vb6_GetScaleMode` / `vb6_WindowScaleModeSelf` / `vb6_ContainerScaleMode`，缺省 1=缇），写的一侧全仓 0 个调用点 ⇒ 谁问 ScaleMode 都答缺省，而 ScaleWidth/ScaleHeight、控件几何（#175 那一族）、文字量纲（#177）全按这一档折算 —— 它错不是"某一枚控件长歪"，是**量出来的数全是错的**。改 = 一处权威 `emitDesignerScaleModeProp` × 三条落点（顶层创建路 / 容器子控件路 / 窗体 WM_CREATE）+ 读写两张表成对（读给**同一个** `vb6_WindowScaleModeSelf`）。**RTL 一字节未动** ⇒ 不碰 `c3rtl.rc`，绕开 #156 那条旧资源坑。普查 93 份设计块 19 处（UC 10 已通 / PictureBox 6 / Form 3）。判据四面：夹具 `tests/scalemode` 两片（存设计值 + 问窗口的**比值**判据，不钉绝对数 + 运行期切档逐数相等 + 双向钉 + Frame 里那枚钉第二条创建路）；负控=改前那台真编 ⇒ 两条 C2039、rc=1；形状针两条在改前那台当场红；哨兵 W1~W5（假 needle 验红：注释掉容器那条落点 ⇒ W3 点名 exit 1）。护栏 A/B 82 份 ⇒ 12 份被改、OFFENDERS 0（其中 VBFlexGridDemo 两片各 2 行是 `Me.ScaleMode` 从 `vb6_ComGetProp` 兜底换成真读数）。真工程配对：ucTreeMaps x86 **C2039 4→2**；czUI-main（唯一 Form 级声明 3=Pixel 的工程）真编真起窗。**顺带量出新账 #199（§B34）：`.pag` 的设计块控件从来没被创建** —— 同一份 emit 里 PropPagFMR 段 `vb6_CreateControl` **0** 处、同工程 Form1 段 **14** 处，页里的 `vb6_hwnd_Picture1` 是"按需新建、初值 NULL"的 UC 槽位 ⇒ 属性页里 `With Picture1` 全打在空句柄上（编得过、发码对、一句不响）。**控件线顺序因此改成 #199 → #196**：#196 剩下 hDC + 按 HWND 的 TextHeight/TextWidth + 裸 `ScaleX/ScaleY`（那条改写还挂在 `isDesignerModule_` 上）+ `vb6_di_TextOutW` 的 C2198；#193（数组元素 extender 属性仍 AV）在后面。
# 上一条（2026-10-05T00:35:00+08:00）= 本轮追加：**账 #192 的一半已出（提交 `a11a7da8`，门 #320 = 11 job 全绿、9 片 FAIL=0，piccur 两片与形状针在 CI 上真 PASS）= ⑦ 的接线那半**。读数自己复算：`_vb6_with_0 = (HWND)vb6_hwnd_Picture1`、`_vb6_with_11 = (HWND)vb6_hwnd_lstFonts` ⇒ 撞兜底的接收者就是**一枚普通控件被 With 包住**，不是 PropertyPage 宿主（⑦ 那句定性只对一半：With 那一路本来就先查 `getControlPropReadFn` 再查 `controlZeroArgMethod`，两张表都在、**档位里没这几行**）。改 = 补表：PictureBox 的 currentx/currenty（读表+写表成对）、clear → `vb6_ClearList`（ListBox/ComboBox），全接在 RTL 早就有的出口上（`vb6_GetCurrentX/Y`/`vb6_SetCurrentX/Y` 按 HWND 存窗口属性）。刻意不给成通用行（VB6 只在画得上去的那几枚上有，给成通用 = `List1.CurrentX` 答 0 = 伪造成功）。判据：新夹具 `tests/piccur` 两片三头钉（写了读得回 + **另一枚没被写坏** + 清空前确实是 2，拦住"Clear 什么都不做"那种假绿）；负控 = `40bbef67` 那台真编同一份夹具 ⇒ 三条 C2039、rc=1；形状针 `with_ctrl_cursor_clear`。真工程：ucTreeMaps 的 C2039 **10→4**（剩 hDC / ScaleMode×2 / TextHeight），TextHeight 那行换了个红法（C2440 BSTR→float），根因同一个。护栏 A/B 80 份 ⇒ 被改的产物 0 份。**另立两格**：§B31 = 账 #196（控件的 .hDC / 按 HWND 的 TextHeight/TextWidth / 带 HWND 的 ScaleX/Y —— RTL 里真没有，其中 .hDC 可复用 _Paint 那套 BeginPaint+SetPropW+GetDC 兜底的两层规则）、§B32 = 账 #197（`VB6_ScaleMode` 全仓 0 个调用者 ⇒ 任何 ScaleMode 读数恒为缺省 1，单位换算建立在它之上 = #154/#160 同族但影响面更大）。**下一轮**：#196 → ucTreeMaps 就能整条过 UnicodePrint；#193；#187（Charts UC 升格进门禁，ucChartBar 已能出 exe）。
# 上一条（2026-10-04T23:55:00+08:00）= 本轮追加：**账 #194 的门 #318 落定（11 job 全绿，能下到的 3 片日志各自 FAIL=0）+ 账 #195 已出（提交 `963c5e71` = §B30，门 #319 = 11 job 全绿、可读到的 4 片各自 FAIL=0，int_suffix 三条用例在 CI 上真 PASS；optdef 两片的 CI 物证也在这批日志里补齐）**。#195 = `%`（Integer 类型后缀）在**词法层**就被拒：`case %:` 那一支只吃字符、不置标志（`&`/`!`/`#`/`^` 都置），radixDigits 的剥离表也没 `%`，于是 `3%` 落回"无后缀按数值大小定档"那一段，parseIntLit 看见残留的 % 就报"超出 64 位"并级联 9-13 条诊断 —— 一条规则**抄了四份**（十进制 + &H + &O + &B）加一张剥离表，漏的是 % 那一份。读数：七形（十进制/三种进制/负数/表达式/参数默认值）修前全红，`12&` 与 `&H10^` 与无后缀三形一直通（证人），上一台（ea3eeeec）读数相同 ⇒ 存量。**改**：四处一起认下 % 并置 isInteger，剥离表补 %；显式 % 超出 -32768..32767 报词法错而不是按 int32 收下再让 int16 截。**判据四面**：正例 `tests/intsuffix` + 边界负例 `tests/intsuffix_neg`（针取 ASCII "(-32768..32767)"）+ 形状针 `int_suffix_shape` + 哨兵 `check_int_suffix_sites.ps1`（假 needle 验红：改回空支 + 剥离表去 % ⇒ S1/S1/S2/S3 四条一起点名）。顺带把 #194 的 optdef 夹具先前为绕开本缺陷改成的无后缀两处恢复原样（`3%` / `5%`），x64+x86 都真跑到 OD-DONE。护栏 A/B 80 份 ⇒ **被改的产物 0 份**（存量里一处 % 都没有 ⇒ 又只能靠夹具响，与 #188/#194 连着同一课三条）。**门 #318 的读数要留一句给后人**：11 job 全绿（API attempt 1），能下到的 3 片日志各自 FAIL=0，但另外 8 片日志与全部工件 zip 被本机网络代理拦成 HTML 页（预签名 URL 全拿不到，重试 4 轮无效）⇒ `optdef` 两片在 CI 上的逐条 PASS 行**没有物证**，下一轮补读。**下一轮**：#192（With 块里那枚控件的成员 —— 家底已盘清并记在 B29⑦：CurrentX/CurrentY/ScaleMode/Clear 四件 RTL 早就有、只差接线；`.hDC` 与 TextHeight/TextWidth 与带 HWND 的 ScaleX/Y 是真没有）、#193（数组元素的 extender 属性仍 AV）、#187（Charts UC 升格进门禁）。
# 上一条（2026-10-04T23:12:00+08:00）= 本轮追加：**门 #317 落定 = 账 #191 收线**（11 job 全绿、各片 FAIL=0；`ucarr_member_call` / `ucobj_chain_write` / `ve_units` 两片 / 哨兵 `uc_instance_exit` 四件在 CI 上全 PASS）+ **账 #194 已出（提交 `40bbef67`，门 #318（run 37211897524，head 9bb2d08b，attempt 1）= 11 job 全绿；能下到的 3 片日志各自 FAIL=0，剩下 8 片与工件zip被本机网络代理拦成 HTML 页（预签名 URL 全拿不到），所以 optdef 两片的 CI 逐条 PASS 行**没有物证** —— 由 job 结论 + 与 evtcase 同款接线 + 本地两台真跑推定，下一轮补读）= B29⑦ 的第二条**：`Optional ... = 0&` 这类 VB 整数后缀被语义层 `evalOptionalDefault` 当文本抄进生成 C ⇒ C2059（真工程物证 ucTreeMaps PropPagFMR.pag:740/886）。改 = 新增 `src/common/int_literal.hpp` 一处权威（Integer 无后缀 / Long 带 L / 超 32 位带 LL），发码侧两支 + 语义侧两支 + Slider 设计期两处手拼全改读它（Slider 那两处换完字节不变）。判据三面：`tests/optdef` 两片两头钉（四枚默认值各按声明落地 + 显式实参照样赢；负控 = `ea3eeeec` 那台真编同一份夹具 ⇒ 三条 C2059 + C2065、rc=1）+ 发码针两枚（`optdef_default_shape_real` 钉在真工程那一行）+ 哨兵 `check_int_literal_shape.ps1` I1~I4（假 needle 验红做过）。真编译读数：ucTreeMaps 的 C2059 2→0（C2039 10 / C2198 1 原样，属 #192）。**A/B 被改的产物 0 份** —— 这本身就是信息：存量 80 份快照里没有一处带后缀的 Optional 默认值，规则只在真工程里响，夹具得自己带（与 #188 的 `12f` 同一课）。**顺带量出一条独立缺陷另立 §B30（账 #195）**：`3%`（Integer 后缀）在词法层就被拒，VB1005 + 13 条级联，上一台编译器读数相同 ⇒ 存量。**下一轮**：#192（With 块里那枚控件的成员从哪张表读 —— 读数与订正已记在 B29⑦）、#193（数组元素的 extender 属性仍 AV）、#187（Charts UC 升格进门禁）。
# 上一条（2026-10-04T22:20:00+08:00）= 本轮追加：**账 #191 = B29⑥ 已出（提交 `f5b4a03b`，门 #317（run 37208673511，head ea3eeeec，attempt 1）= 11 job 全绿、10 片 FAIL=0）** + 两条教训。①控件数组元素的成员访问以前不走 UC 直发那条出口（Fix 112 只给单枚接过），整条落 `vb6_ComCall` 兜底 => 空值；改 = 抽成唯一出口 `emitUcInstanceMemberExpr`，单枚交 `vb6_hwnd_<名>`、元素交 `vb6_CtrlArr_GetAt(&vb6_arr_<名>, i)`，两条路只差这一个串。②**A/B 抓出的第二条**：元素改道之后 `ucChartBar1(i).Font.Size = ...` 落进左值改写的字符串级折叠（Pattern C/D2），它只取 `prop_get_` 那一段、把尾巴 `.Size` 整段丢掉 => 发成 `prop_let_Font(elem, <double>)`，四条 C2440 把刚编得过的 ucChartBar 打回编不过；修 = 折叠之前加一条"带尾巴的 prop_get_ 目标"独立落点（取回那枚 StdFont 再写一层），认不得的尾巴 `return false` 当场红。**读数**：ucChartBar x86 **rc=0 且产出 exe**；ucProgressCircular 错误类严格是基线子集（16xC2059+2xC2065+1xC2084+25xC2198 -> 1xC2065+1xC2084+12xC2198，C2059 归 #188、剩 B29③④）。**判据**：`U-ARRM-methods=3` 真跑证人（x64+x86 都到 `U-DONE`）+ `ucarr_member_call`/`ucobj_chain_write` 两枚形状针（后者钉真实工程）+ 哨兵 `check_uc_instance_exit.ps1` U1~U4（假 needle 验红：改 `(Pattern C-tail)` 一名 => U4 点名）。**护栏**：A/B 40 工程 x 两架构 = 80 份 ⇒ 只有 6 份被改、每条差异行逐行归因、OFFENDERS 0。**留给下一轮**：#193（数组元素的 extender 属性 `Left/Top` 仍 AV，这一刀刻意没碰）、#192（ucTreeMaps 过 parser 后红移到 cl：10xC2039 那族 PropertyPage 伪成员）、#187（Charts UC 子工程升格进门禁 —— ucChartBar 现在真编得过、出 exe）。
# 上一条（2026-10-04T21:25:00+08:00）= 本轮追加：**账 #186 已出（提交 `4d1dbf45`，门 #316（run 37205289964，head `4d1dbf45`，attempt 1）= 11 job 全绿、10 片 `FAIL=0`）** + 账 #189/#190 的门号回填（#314 / #315 都 11 job 全绿、10 片 FAIL=0）。**#186 = Erase 支持带下标的成员数组目标**：`Erase m_tvFiles(lIndex).bvData` 以前在 parser 就 VB2001，整个 ucTreeMaps 卡在最前面。改 = `EraseStmt` 加 `targets`（与 ReDim 的 `targetExpr` 同一套），点链→表达式树收成 `Parser::buildDottedNameExpr` 让 ReDim 转调（不再抄第二份），发码出 `VB6_SA_AT(vb6_type_TElem, m_items, 1).bvData` 销毁并置 NULL；1D/ND 不用新表（UDT 成员的动态数组恒一维，结构体发码单点）；`Erase arr(i)`（无成员）**继续报诊断** + 一条负例钉住"不许过头"。**自己撞的那条值钱**：空括号必须"先吃完整条点链再吃括号"（Fix 082 的另一半）—— 第一版在基名后直接判括号，于是 `Erase .MaxWidths()`（VBFlexGrid 6590）残留裸 `(` ⇒ VB2003/2002；**是 A/B 逐字节比抓到的**：flex 两片从 5.3MB 掉成 1.3KB ⇒ 快照**体量**差比"差异行数"更早报警。判据四面：`tests/erase_sub` 两片两头钉（只擦那一格 + 那一格可再用，后者顺带是双释放探测器；负控 = `4d1dbf45` 编同一份夹具 ⇒ VB2001、rc=1）+ `erase_neg_indexed` 负例 + `erase_member_shape` 形状针。护栏 A/B = 40 工程 × 两架构 ⇒ 生成的 C 逐字节相同 78/78。**下一轮**：#192（ucTreeMaps 过了 parser 之后红移到 cl：10×C2039 全是 PropertyPage 伪成员被当 HWND 结构体成员发 —— 与账 #159/#174 同源，先读 `check_host_pseudo_table*.ps1` 那张表再补名；另有 2×C2059 + 1×C2198 是第二族）；B29⑥（数组元素的成员访问，根因与落点已量清）；#187（Charts UC 子工程升格进门禁 —— ucChartBar 现在能出 exe 了）。
# 上一条（2026-10-04T20:55:00+08:00）= # 本轮追加：**账 #190 = B29⑤ 已出（提交 `a27758c3`，门 #315（run 37202837330，head `a159b85d`，attempt 1）= 11 job 全绿、10 片 `FAIL=0`）**：事件臂调用的函数名以前按**控件的设计期拼写**现拼，过程定义发的却是 **Sub 自己的拼写** —— VB6 大小写不敏感、C 敏感，"改了控件名没改过程名"就变成引用一个没人定义的函数（链接期 LNK2019 + LNK1120）。这条红**只在链接期现形**：存在性那一步 symTab_.lookup 本来就大小写无关、命中了，`--emit-c` 的形状与语义层都看不出坏 —— 缺的只是"命中之后按谁的名字发"。改 = 新增唯一出口 `CCodeGen::eventHandlerFn()`（存在性与名字出自同一次 lookup，命中交 `cProcName(sym->name)`，没命中交空串让调用方不装臂），43 处现拼全接过去。**普查教训**：第一遍按 `ctrl.controlName`/`ctrlName` 两个变量名扫 ⇒ 少了 18 处（`info.ctrlName` 那 16 + 菜单 2），"27" 是错的 ⇒ census 一律按**形状**写、别枚举变量名。判据三面：新夹具 `tests/evtcase`（三枚控件：两枚只差大小写 + 一枚拼写一致的**证人**，Timer 发 BM_CLICK 自驱）x64/x86 两片，负控 = `a27758c3` 那台编同一份夹具 ⇒ LNK2019 恰好两枚、证人没第三条；真工程配对读数 ucChartBar x86 **LNK2019 2→0 且第一次产出 exe**；哨兵 `scripts/check_event_handler_names.ps1`（E1 现拼=0 / E2 从符号取名 + 空串出口 / E3 调用点>=40）假 needle 验红过。护栏 = 78 次发射 A/B，只有 4 份产物各动 2 行且每行验过"整行只换一个标识符"，其余逐字节相同、OFFENDERS 0。**下一轮**：B29⑥（数组**元素**的成员访问：`uArr(i).SW()` 返空、`uArr(i).Left` 崩 ⇒ 设计期那一路多包了一层 `vb6_UC_InstanceOf`，运行期这路没包，待量）；#187 的升格清单从此多一件（ucChartBar 已能产出 exe）；B29③（ucProgressCircular 的 C2084 过程重发体）与账 #186（Erase 带下标）仍在。
# 上一条（2026-10-04T20:20:00+08:00）= # 本轮追加：**账 #189 = B29② 已出（提交 `ea5155af`，门 #314（run 37200966607，head `0efb29b0`，attempt 1）= 11 job 全绿、10 片 `FAIL=0`）**：VB6 里 `arr.Count` 问的是数组本身 —— 发码侧只有 `LBound`/`UBound` 各一条 if、`Count` 漏了 ⇒ 掉进 COM 兜底，发成 `vb6_ComGetIntProp(vb6_hwnd_<数组名>, L"Count")`：跨窗体是 C2065（Charts 2020/ucChartBar 的 Form2 就这么红的），同窗体是**编得过、恒答 0**（`For i = 2 To arr.Count - 1` 一格不走）。改 = 新增唯一出口 `CCodeGen::ctrlArrayMetaMemberExpr()`，并把这三条放在**问控件属性之前**（旧那两条 if 还挂着 `readFn.empty()` 当前提，一并撤）。判据三面：`tests/ve_units` 塞一枚 UC 控件数组 `uArr(0..2)` 两头钉（三个数 + 拿 Count 真的圈三圈；负控 = `a4e3b574` 编同一份夹具 ⇒ `frmUnits.c(214)/(220)` C2065、rc=1）+ 发码形状针 `ctrlarr_member`（三枚出口 / Absent 是改前那条，两台编译器实测 2→0）+ 哨兵 `scripts/check_ctrl_array_members.ps1`（A1 出口只在权威一个文件 / A2 三条分支齐 / A3 手写分支归零），假 needle 验红做过。护栏 = 沿用 #188 从形状断言面推出的清单（38 工程 × 两架构 = 76 次发射），只有 ucChartBar 与 ve_units 四片被改、行数不动、增=删、OFFENDERS 0（BASE 那 4 对 `12f`→`12.0f` 已逐行归因成 #188 的账）。**下一轮**：②清掉之后 ucChartBar 走到链接期才红 ⇒ 新记 B29⑤（事件臂函数名按控件设计期拼写现拼 vs 过程按 Sub 自己的拼名，大小写不一致就是两个符号，`LNK2019` ×2；普查 27 处现拼、同文件已有 3 处是对的写法 ⇒ 收成一处 `eventHandlerFn` + census 哨兵）与 B29⑥（数组**元素**的成员访问：`uArr(i).SW()` 返空、`uArr(i).Left` 崩；设计期那一路多包了一层 `vb6_UC_InstanceOf`，运行期这路没包 —— 待量）。
# 上一条（2026-10-04T19:40:00+08:00）= # 本轮追加：**账 #188 = B29① 已出（提交 `33cb239a` + 收口 `936124e8`，门 #313（run 37198483318，head `936124e8`，attempt 1）= 11 job 全绿、11 片 `FAIL=0`；**中间门 #312 红一条，红因是我这一刀**）**：C 里 `12f` 是非法 token（C2059 bad suffix），整数值的 Single 必须写 `12.0f`。这条规则在编译器里散着三份，「.frm 设计期字体块」那处（`snprintf("%.4g")` 后直接拼 f）抄漏了补小数点 ⇒ 设计期字号是整数的工程直接编不过（Charts 2020 的 ucChartBar 4 条、ucProgressCircular 16 条 C2059）；VB6 默认字号 8.25 带小数点 ⇒ 门禁全部夹具恰好躲过。改：新增 `src/common/float_literal.hpp` 一处权威（`floatingLiteralText` / `floatSingleLiteral` / `floatFixed6Literal`，经典 locale + 没有 `.` 或指数就补 `.0`），五处发射接过去，位数各按原样 ⇒ 发码字节只在真正坏掉那一处变。**门 #312 那一红的教训（比修本身更值钱）**：我顺手把三处 `std::to_string(v)+"f"` 也"收进权威"，而那三处**本来就合法**，且设计期 FontSize 的字面形状正被 `sl_emitc_native`（tests/ctrlslider，vbp#4 片）钉着 ⇒ 变成 `20.0f` 就红。第一版 A/B 只跑手挑的 8 个工程 ⇒ 本地全绿、门上一片红。**A/B 的工程清单必须从门禁的形状断言面推出来**：本仓发码形状层面的断言实测覆盖 30 个文件（`.build/b188_surface.txt`，由 `.build/b188_surface.py` 从 run_tests.ps1 推），加宽版 A/B = 37 文件 × 两架构 74 次发射，结果只有 ucChartBar 两片被改、数值多重集相等、增删 0 行。判据：`fontsize_literal` 发码形状针钉在**真实工程** ucChartBar 上（Needles `12.0f` + 证人 `8.25f`，Absent 是改前的 `12f`）；哨兵 `scripts/check_float_literal_shape.ps1` R1~R3 在缺这刀的树上全红（R1 点出 6 个手拼点）。配对读数（同一工程 × 两台编译器 × x86）：ucChartBar C2059 4→0、ucProgressCircular 16→0，剩下的正是 B29②③。**下一轮：B29② `ucChartBar1.Count` 这类控件数组整体成员发成没声明的 `vb6_hwnd_<数组名>`（C2065），与账 #157 同族（句柄表达式必须走 `ctrlHwndExprForInit` / `vb6_arr_*` 那一处）；然后 B29③ 的过程重发体。**
# 上一条（2026-10-04T13:20:00+08:00）= # 本轮追加：**账 #185 = C29-PB-a 已出（提交 `a4e3b574`，门 #311（run 37177647595，head `a4e3b574`，attempt 1）= 11 job 全绿、11 片 `FAIL=0`）**：B13 那条「同名属性」的真形状是**装了但没装** —— RTL 建窗时给 STATIC+SS_BITMAP（PictureBox/Image）装的自绘子类，与发码为带事件的控件装的事件子类，把「下面那层的 wndproc」存在**同一个窗口属性名** `VB6_OrigProc` 上、两边都写「属性已存在就不装」⇒ 后装的事件层被整层丢掉：夹具实测 Paint/MouseDown/MouseUp/Click + Image 的 Click **五条全 0**（x86、x64 同形），而同窗体上没人抢的 Label 一直响。改：① 自绘层换自己的槽位 `VB6_ImageOrigProc`（口径 = 一层一名，与 GBox/GfxBtn/SSTab 同规矩）；② 画的序收成一条 —— 最外层 BeginPaint/EndPaint 一次，DC 经 `VB6_PaintDC` 交下去画表面，再抬用户 `_Paint`。判据：`tests/pbsub` 自驱进门禁 x64+x86（两头钉：五条事件针 + PB01 那颗像素必须是设计期 BackColor 的红），负控 = HEAD 的编译器编同一份夹具（两边只剩 PB00/PB06、`paint_ok=0`）；哨兵 `check_subclass_slots.ps1` 四条规则在 HEAD 树上 6 行全红；A/B `--emit-c` 7 工程×两架构 = 12 份逐行相同、flex +2/-0。另两件只是**测量**：B15 那条重测 0/60（连基线件也不复现 ⇒ 这组判不了，且 #184 在 x64 本来无字节后果）；B14 换成真编译读数（ucChartArea/ucPieChart 已编得过、可升格）+ 新立 B29 四条（设计期字体 Size 发成 `12f` 这种非法浮点字面量、宿主 UC 实例句柆名 C2065、同一个过程发两遍体、`.pag` 报错行号不含设计期头块）。**下一轮从 B29 ① 开刀（字面量形状收进 `floatingLiteral` 这一个出口）。**
# 上一条（2026-10-04T11:30:00+08:00）= # 本轮追加：**账 #184 = B23 出（提交 `ab45e2d3`，门 #310（run 37172942961，head `ab45e2d3`，attempt 1）= 11 job 全绿）**：`AddressOf` 一直把**本体的裸地址**交给 OS —— x86 上本体是 `__cdecl`，回调是 `__stdcall` ⇒ 每回调少弹 N*4 字节；VBFlexGrid 每枚网格的 SUBCLASSPROC（6 形参 = 24 字节）因此把起窗期栈踩掉 ⇒ `0xC0000374`（改前 3/4~5/6 崩，改后 14/14 起窗）。改法按 VB6 口径：被取址的标准模块过程在定义模块里另发一枚 `__stdcall` 转发桩（参数表逐字复制本体、体内原样转调），取址点交桩地址，**本体一字不改** ⇒ 直接调用路与 RTL 的 cdecl 登记面全不牵连；口径一处 `addressOfTargetCName`，与委托桩共用签名拆分；标记趟 stage 3.5c（晚于跨模块链接）。判据：`test_addrof_cb` 真跑 x64+x86 全 Y + `addrof_cb_shape` 发码形状针（HEAD~1 一个 `aoThunk_` 都没有 = 负控）+ 哨兵 `check_addressof_thunk_sites.ps1`（HEAD~1 树上 R1~R4 全红）+ A/B 六工程两架构（五份逐行相同、flex +81/−0 全是桩与改名）。**并把 FlexGridX86 挂进门禁**（B23 那句「这份 demo 完全在回归之外」从此不再成立）。
# 上一条（2026-10-04T08:55:00+08:00）= # 本轮追加：**账 #172 = 已出（提交 `3e7231c7`，门 #309（run 37165123051，head `3e7231c7`，attempt 1）= 11 个 job 全绿、非绿 0）**：本账的前提换掉了 —— 不是「Date 型返回默认值读到未初始化内存」，是**日期字面量 `#...#` 从来没有值**：parser 建 `LiteralExpr` 只挂原文，构造函数又只清了 union 的前 4 个字节，Date 档发码读的是 8 字节的 `doubleValue` ⇒ Debug 恰好 0.0、VS/Release 与本机 /RTCu 那台是 -6.277e+66；**VB6 里 `#1/1/1900#` 是 2.0，所以改前 Debug 那台也是错的，只是错得稳定**。工具手法记一笔：**复现这条不用去配 VS 生成器那台** —— 同源码同生成器、只加 `/RTCu` 冷编一台 C3.exe，`--emit-c` 一比就把 4 处垃圾点钉出来（flex 的 ComboCalendar Min/MaxDate 的 ret 赋值 + 两条 `Select Case x To y` 折出的边界），改后两台**逐行零差异**（这才是本账要的判据形态：常量不再随编译器自身配置漂）。修法收成一处：`foldDateLiteralToOADate()`（parser_helpers.cpp 定义、parser.hpp 声明）折 M/D/Y 与 D-M-Y（VB6 两种分隔符口径不同）、`H:N[:S]` + AM/PM、两位年份 <50→2000s / ≥50→1900s、闰年与真日历校验、OLE epoch 1899-12-30=0.0 用 daysFromCivil；构造函数改成整体清零；**认不出的形状照旧留 0、不发新诊断**（宁可算错，也不把现在编得过的工程编红）。判据两头：新夹具 `tests/test_datelit.bas`（14 条读数）进 bas 队列，门工件 `test-logs-bas-1/job8/test_datelit.out` 全对、`.err` 0 字节；负控 = 同一份测试喂**改前**的编译器 → 10 条 False 且 `L-serial=-6.27743597849989e+66`（那条垃圾常量第一次在真跑输出里直接看见）。A/B（BASE 用「临时回退四份源文件重编」拿到，同配置同生成器）：六工程里五工程 **0 行变化**，flex 4 行、且只允许数字变，新值只有 {2.0, 2958465.0, 2958465.999988426}。顺带：本账旧记录里「CI 正是 Release ⇒ 门里带着垃圾常量」那句按 /RTCu 的证据改写 —— 垃圾来自**编译期折叠**，与 RTL 无关、与运行时配置无关。STATUS 保持 IDLE。
# 上一条（2026-10-04T07:35:00+08:00）= # 本轮追加：**账 #182 + #183 = 已出（提交 `0b13dcf6` = C29-CH-g、`d9e34590` = C29-CH-h，门 #308（run 37161372755，head `d9e34590`，attempt 1）= 11 个 job 全绿、非绿 0）**：先修**归因工具**、再用它归因并修掉一个 100% 复现的关闭路径跳飞调用。#182（工具）：x64 的崩溃轨迹以前**一条应用帧都没有** —— 三处各自堵住：① `vb6forms.c` 那段栈扫描整段包在「故障地址属于哪个模块」那一问里，`call` 跳飞（rip=0x1）时这一问直接失败 ⇒ 一个候选都不打；② `vb6rtl.c` 那段 Esp/Rsp 线性扫描只在 `#ifdef _M_IX86` 里编；③ 扫描的 x64 分支用 `wsprintfA("0x%016llX")` —— 用户态 wsprintf **不认 `ll`**，一直打成 `0xlX -> <一串字节>+0xlX`（这段以前从没被执行到，①改完才露出来）。改法：扫描一律改扫**主 exe 镜像**（生成代码都在 exe 里，扫系统模块没用）、扩到 `_M_X64`（取 Rsp、按 ULONG_PTR 步长）、偏移改打 32 位 RVA；顺手把不在本模块镜像里的帧照实标 `(outside exe)`（以前把 ntdll 的地址减掉自己的基址打成「rva=0x4376...」，看着像自己的符号，会把人带到沟里），并把文件侧 AV 头的 `ExceptionInformation[0]==8` 认成 EXECUTE(DEP)（以前一律写 WRITE，与 vb6rtl.c 里 Fix 187 那条口径对齐）。读数：同一份夹具改前只有 5 帧派发链、栈扫描一段都没有；改后 `st+0 rva=0x2581ae` 落在 `vb6_ComCall` 里，`st+37/st+40/st+48` 落在工程自己的 `VTableHandle`（`IOleIPAO_EnableModeless` / `GetVTableIPAO` / `ActivateIPAO`）上 —— **#183 就是靠这几行定下来的**。#183（产品）：Fix 191 那条「按名 AddRef/Release 直发槽位」少绕了一层：`((void**)disp)[1]` 读的是 **disp+8（对象自己的第二个字段）**，而槽位要先从对象首字读出 vtable；`vb6_ComIsDispatchable` 自己是两级读法，两口径不一致。在 `VTableHandle.bas` 手搭的伪 `IOleInPlaceActiveObject` 上，`VTableIPAODataStruct` 第二字段恰好是 `RefCount As Long` ⇒ 读出来是 1 ⇒ `call 1`。触发条件量到是**确定的不是偶发**：向 FlexGrid 子窗发一条 WM_LBUTTONDOWN（只这一条；WM_LBUTTONUP / 右键 / 滚轮都不发）再对主窗发 WM_CLOSE ⇒ x64 `-g` 与 x64 无 `-g` 各 16/16 复现。改成 `lpVtbl->AddRef/Release`（与 `vb6com.c` 里那两处同形）后两台各 12 次关窗全干净、一条崩溃现场都不写；负控（改前那台）同条件 2/2 仍崩。census：全 RTL 再无 `((void**)x)[n]` 这种手写槽位读法（`grep` = 0），按 vtable 调一律走 `->lpVtbl->`（218 处）。顺带把 `vb6_ComGetProp` 那条 `fallback to IDispatch` 从字体/Extender/宿主三个岔口**之前**挪到之后 —— 它对根本不走 IDispatch 的调用也照打，本轮差点据此把嫌疑引向没执行过的路径。四条工具事实（都已进 memory）：崩溃后的**退出码不是判据**（预修复那台 rc=0x00000000 同时写了 c3_crash.txt）；`c3_crash.txt` 是相对路径 = 落在被测进程的**当前目录**（不给 `-WorkingDirectory` 就一路假绿）；`Start-Process -PassThru` 的 `.ExitCode` 在碰过 `.MainWindowHandle` 之后会拿到 `$null`（取 rc 前先用 EnumWindows 按 pid 自己找句柄）；**产物架构别看目录名** —— 一次漏传 `--arch x86`，`b182gx86n` 其实是 x64，整条「x86 侧读数」当场作废（一行读 PE 头就能验掉）。另记：x86 那台**起窗就 0xC0000374** 与本轮无关（把 #183 退回去单验仍崩）⇒ B23 继续开着。STATUS 保持 IDLE。
# 上一条（2026-10-04T10:55:00+08:00）= # 本轮追加：**账 #181 = 已出（提交 `e16e42cd` = C29-CH-f，门 #307（run 37156897477，head `e16e42cd`，attempt 1）= 11 个 job 全绿、非绿 0）**：崩溃轨迹**每进程只记一次**的闸 —— 诊断工具不再把自己的现场盖掉。症状是实测不是推测：czUI x64 启动期 AV 那份 c3_crash.txt 里**同一份递归栈被写了 4 遍**、真正的第一现场（gdiplus 那条 AV）反而排在最后一段；B23/#176 那种堆损坏归因读的正是这份文件 ⇒ 这个闸是所有后续崩溃调查的前置。根因：记轨迹这件事本身会再触发异常（栈溢出时 fprintf 要拿 CRT 锁、CaptureStackBackTrace 与栈扫描要再读同一批页）⇒ 处理器套处理器，而三个出口（`vb6_CrashTraceVEH` 打 stderr、`vb6_heapCorruptVEH` 写文件、`vb6_crashFilter` 写文件末段）**一个闸都没有**。修法 = 一处 claim 口径共用：新增零依赖头 `vb6rtl_crash.h` 声明 `int vb6_CrashTraceClaim(int slot)`（定义在 vb6rtl.c，InterlockedCompareExchange 一张 4 槽表）。**槽位各记一次而不是全局一次** —— 三个出口写的是不同落点、各有价值，全局一次会让 stderr 那份抢掉文件那份；FILTER 天然只跑一次，给它一个槽是为了让「只记一次」在同一处能读全。两条踩过的坑记进 §C：① 第一版把声明塞进 `vb6rtl_runtime.h`，`vb6forms.c` 提前吃它就炸（那个头自己要求排在 array/builtin 之后，报 `vb6_SafeArray1D` 未声明）⇒ 诊断面的东西不该挑宿主；② 新增 RTL 头要**三处登记**才生效（CMakeLists 源列表 / c3rtl.rc 的 RCDATA 号 222 / rtl_embedded.hpp 枚举 + .cpp 号→名），少一处就是 C1083 找不到包含文件。读数：同一夹具同一 env，改前 6 段 ⇒ 改后 **2 段且第一段就是 FAULT gdiplus.dll+0xF2E1**（原始现场回到第一行）；产品行为零改动（三个出口本来就只在 C3_CRASH_TRACE 下装，不带 env 不写文件、退出码不变）。回归面：ve_units / charts / flex / ctrlrichtextbox 四份产物 cl 0 error，ve_units 十三条针（含 U-HW / U-CNT）全 True、rt 103 行无一条 N、charts+flex 各活过 14s 无崩。STATUS 保持 IDLE。
# 上一条（2026-10-04T09:10:00+08:00）= # 本轮追加：**账 #161 = 已出（提交 `a0b4c67a` = C29-RT-e，门 #306（run 37155455037，head `a0b4c67a`，attempt 1）= 11 个 job 全绿、零新增失败**：**读数把本账的前提推翻了** —— 「RichTextBox 带焦点时一次赋值发两条 `_Change`」不复现。三个读数（纯 VB 夹具，没往 C 里加 trace）：① 一次 `Text` 赋值 聚焦/不聚焦都 chg=**1**；③ 只动 `SelStart`（文本不动）chg=**0**、sel=1 ⇒ 当年那条病灶猜测（`_Change` 挂在 EN_UPDATE ⇒ 选区变也发）**当场否掉**，arm 选的来源没发错；⑥ 一次 `SelText` 赋值也是 1；⑩ 聚焦下连做三次、每次各自取基线 = **1/1/1**。唯一出现 2 的形状是 ⑧：**焦点刚落到这枚控件之后的第一格** —— RichEdit 因焦点变化补发的那条通知还没落地，被下一个 `DoEvents` 算进相邻那一格，与账 #162（已关掉的 Timer 仍报 21 拍）同族 ⇒ 判据被在途通知污染，不是产品翻倍。所以这一刀**只改判据、产品侧一行未动**：`RT84` 从 `>= 1` 回到钉死（取基线前 `SetFocus + DoEvents + DoEvents` 排空，然后 `= 1` 且**再泵一轮不加发**，后半句把「还有一条在途」直接判红）；新增 `RT91` = 只动选区两次 ⇒ Change 必须 0 条（**只问 Change**：自发 SelChange 的条数天生会抖，同一份产物连跑 E3 的 sel 位在 2..7 之间跳，实测第 3 次抓到 6/2，拿它当判据会把环境抖动读成回归）；新增 `RT92` = 聚焦一次 `Text` 赋值恰好 1 条 + 再泵一轮不加发；`$rtNeedles` 加 `RT91=Y`/`RT92=Y`（RT1..RT89 那条循环没动，RT90 仍钉精确值）。判据不是空钉：⑧ 那格量到过 2，所以 `= 1` 有判别力。本地稳定性：x64 连跑 4 次、x86 连跑 3 次全 Y；门里两片 `RtfApp.out` 的 RT84/RT91/RT92 全 Y、`CTRLRICHTEXT-DONE` 齐。STATUS 保持 IDLE。
# 上一条（2026-10-04T07:35:00+08:00）= # 本轮追加：**账 #180 = 已出（提交 `6533bae4` = <vbeclipse> rev44，门 #305（run 37153706069，head `6533bae4`，attempt 1）= 11 个 job 全绿、零新增失败）**：§B 的 B19 落地 —— `vb6_UserControl_ContainerHwnd` 从 `int32_t` 换成指针宽度。x64 上写入点 `vb6_UserControl_ContainerHwnd = (int32_t)(intptr_t)r->parent` 把 HWND 的高 32 位当场丢掉，而语料里 VBFlexGrid.ctl 把它当 HWND 直接交给 `MapWindowPoints` / `GetWindowLongW`（产物 5 处），charts 另有 12 处存进自己的 Long 字段 ⇒ 声明 / 定义 / push-pop 快照字段 / 写入点四处一起改成 `void*`（与 `hWnd` `hDC` 同档），#159 那张表里对应那一行跟着答 `LongPtr` —— **于是哨兵从此钉住这个宽度**：谁改回 int32_t，`ROW-TYPE-MISMATCH` 就在门里红（判据分工要写清： HWND 是否 ≥2^31 不可控，运行期针抓不到截断本身，抓得到的是「槽没被填」「读错成 hWnd」）。**顺带露出并修掉一格**：类型答对之后 `CStr(UserControl.ContainerHwnd)` 打**空串** —— `CStr` 对 LongPtr 实参落到 `vb6_CStr(vb6_VariantFromValue(x))`，而 `_Generic`（vb6rtl_variant.h:198-216）没有指针那一档 ⇒ 命中 `default: vb6_VariantObject`。只在类型已知的 CStr 分诊处加一条 LongPtr 支路（`vb6_CStrLongLong((int64_t)(intptr_t)(x))`）；**不能**往 `_Generic` 补 `void* :` —— 真对象引用在 C 层也是 `void*`，那一条会把所有对象一律打成数字。夹具 ve_units 加 `U-CNT`（三头：容器非零 / 容器≠自己 / `CStr(容器)` 非空串）；**红侧实测**：只落 ① 不加 ② 那条支路 ⇒ `cnt=` 空、`U-CNT=False` ⇒ 这条针确实钉的是新代码路径。**--emit-c A/B（六工程 × 两个 exe）**：charts / czui / flex / ve_list / iface_wrap **逐行相同**（槽宽住在 RTL 里，不改发码文本），只有 ve_units 两枚 UC 的 CntStr 换形状，六工程行数未增删。真跑：ve_units x64 + x86 两片 `U-CNT=True`（cnt 打出真实句柄数）、其余六条针一字未动；charts / flex / flex_x86 / czui_x86 各活过 14s 无崩，四份产物 cl 0 error。门工件对形 vs 门 #304：115 份里 108 份逐份哈希相同，`VeUnits.out` 两片各 +2 行（本刀），另 6 份（FrmEvents / ModalApp / RtfApp / TabWalkApp / TmApp / WsApp）**只改值不改判据** —— 拖放 X 坐标、HWND 编号、字符位置、timer 拍数 ±1、临时端口与多收一个数据报，全在已开账的抖动族（#79 / #161 / #162）。**另量到一条与本刀无关的预存事实**：czUI 的 **x64** 产物退出时 rc=0xC000041D，拿改前的编译器跑同一份夹具是**同一个码** ⇒ 预存（CI 那格本来就是 `-Arch "x86"`，x86 侧无崩）⇒ 新立 B27。STATUS 保持 IDLE。
# 上一条（2026-10-04T06:05:00+08:00）= # 本轮追加：**账 #159 = 已出（提交 `ef260daf` = <vbeclipse> rev43，门 #304（run 37150507733，head `ef260daf`，attempt 1）= 11 个 job 全绿、零新增失败）**：宿主伪成员（UserControl / PropertyPage / Extender / Ambient）的「叫什么 / 是什么类型 / 能不能裸写 / 赋值要不要拆」收成**一张表** `kHostPseudoRows`（55 行、35 标量），四个消费点 + 拼写规范化都改读它 —— 此前是五份互不相交的清单（裸名两张表、类型 oracle 两条硬编码、赋值一份 15 枚名单、规范化只剩 hdc）。**根因在读法不在值**（B25 那条订正就此坐实）：类型 oracle 认不得这些 RTL 全局量 ⇒ 答 Variant ⇒ 比较发成 `vb6_VarCmpLongNe(&vb6_UserControl_hWnd, 0)`，即把 8 字节 `void*` 的**地址**当 `vb6_VARIANT*`（16 字节）递进 RTL ⇒ `<>0` 恒假；`CStr` 那一路落进 `_Generic` 的 `default: vb6_VariantObject` ⇒ 打空串。值本身是真的：临时探针（用完已撤）在 `vb6_uc_push` 里打 `r->hwnd` 与全局，两者同值非 0。**红→绿**：同一份夹具只换编译器 —— BASE 发 `vb6_VarCmpLongNe(&…, 0)` / `vb6_VarCmpNe(&hWnd, &hDC)`，NEW 发 `(-(vb6_UserControl_hWnd != 0))` / (-(hWnd != hDC))；真跑 `U-HW-RAW twip=True/True pix=True/True`、`U-HW=True`，CI 两片（x64 + x86）同值，其余五条针一字未动。**--emit-c A/B 六工程**：只有 charts +12/-12、flex +2/-2、ve_units +8/-8 变，逐行归因全在这一家（`ComPackValue(DisplayName)`→`ComPackBSTR`、`ComPackInt(Enabled)`→`ComPackBool`、`PropertyPage.hwnd`→`_hWnd` 拼写归一、两枚 `As Integer` 的 Extender 返回补 `vb6_ChkInt`、`CStr(ScaleMode)` 从装箱路改直取、hWnd/hDC 两条比较）；czui / ve_list / iface_wrap **零差异**，六工程行数一字未增删（35224/4545/82872/1057/321/690）。真跑：charts 截图与 #175 收线时同一幅，flex/charts 各活过 14s 无崩，三工程 cl 0 error。哨兵 `scripts/check_host_pseudo_table.ps1`（逐行把表与 `vb6rtl_userctl.h` 的 extern 类型对照 + 禁那五份旧清单回潮；两条假针 —— hWnd 改答 Long、收一个 RTL 没声明的成员 —— 都能红）进回归 `[STATIC] host_pseudo_table_census`，门里 PASS。**表外欠账另立 B26**：PropertyPage 一族全局运行期从没人填值、`vb6_UserControl_Left/Top` 等成员 RTL 根本没声明、表外手写宿主符号 24 处。STATUS 保持 IDLE。
# 上一条（2026-10-04T04:20:00+08:00）= # 本轮追加：**账 #179 = 已出（提交 `d5f3e190` 之后第一刀 `30a3f4ff` = <vbeclipse> rev42，门 #303（run 37144897822，head `30a3f4ff`，attempt 1）= 11 个 job 全绿、零新增失败）**：B25 里那条"控件宿主成员随调用路径变值"的**路 B** 落地 —— 一处口径替掉逐成员补出口。实测（同一枚夹具加 `Ctx()` 报自己的两个宿主成员）：改前 控件自己里 `mode=1`、容器里调 `mode=3`（残值来自先初始化的那枚像素型控件）；改后两侧同值。**为什么能做成配对**：`Exit Function/Sub/Property` **早就**走 `vb6_proc_exit` 统一出口尾（`cgen_decl_func.cpp:412` / `_proc.cpp:366` / `_prop.cpp:320` 三处同一条 Fix <vbeclipse>），所以体首 Push、出口尾 Pop 不会漏弹 —— 这条同时把 B25 里"路 B 要等 T31-A"那个顾虑判掉了：**前提本来就在，不用等**。RTL 侧新增 `vb6_UC_PushInstance/vb6_UC_PopInstance`（栈在 RTL 内部，`vb6_UCSaved` 不外泄给生成码；查不到实例压一份原样快照、Pop 等值还原；深度 64 溢出整对空转 ⇒ 任何情况配平不破）。与既有那个**故意不弹栈**的 `vb6_UC_Enter(hwnd)`（"最近进入者"语义，正是残值来源）分开命名。谓词 `ucCtxScoped() = isDesignerModule_ && !isPropertyPageDesigner_` 只写在 `cgen_state.inc` 一处，三个发射点共用。**不变式（--emit-c 全项目扫，逐函数配平）**：charts push=pop=659、czui 121、flex 775、ve_units 10；`ve_list` / `iface_wrap`（无 .ctl）**两个都是 0** ⇒ 谓词精确排除类模块与窗体。判据：夹具针面加 `U-CTX`（缇型那枚必须报 `1/…`、像素型必须报 `3/…`，只取 ScaleMode 一位数字 ⇒ 与 DPI 无关），红侧 = 上面那行改前读数（同一份探针、只差这一刀）；原四条针一字未动。回归读数：Charts 2020 每控件绘制缓冲颜色数与 #177 同区间（pie 455/244、min 63，同一产物跑两遍自身 ±5% 抖动内）；czUI 子窗口 rect 逐条相同；`ve_list` 三条针、`iface_wrap` 两条针原样命中。**⚠ 本格有一条自我订正**：提交信息 `30a3f4ff` 里我写"`UserControl.hWnd` 读回 0 不是 #159、比较器没问题、值真的是 0" —— **这个结论是错的**：`vb6_VarCmpLongNe` 的第一形参是 `vb6_VARIANT*`（`vb6rtl.c:593`），发码传的是 `&vb6_UserControl_hWnd`（一个 `void*` 全局的地址），于是 `<>0` 与 `=0` 两次读数走的是**同一个被 reinterpret 的槽位**，这对探针根本判别不了（教训：**并排打两个互补比较**看着像双向证据，其实共用同一个坏读法时两边都不可信）。⇒ 那条回到 #159（hwnd 被当对象装箱）名下，下一步换不经比较的读法（把 hWnd 交给收 LongPtr 的 Declare 桩、由桩回打真实数值）先分清"值 0"还是"读法错"。B25 已按这个改写。STATUS 保持 IDLE。
# 上一条（2026-10-04T02:55:00+08:00）= # 本轮追加：**账 #177 + #178 = 已出（提交 `d5f3e190` = <vbeclipse> rev41，门 #302（run 37142437975，head `d5f3e190`，attempt 1）= 11 个 job 全绿、零新增失败）**：接 #175 的量纲那一族，新夹具 `tests/ve_units`（两枚同尺寸 UC，只差 `.ctl` 声明的 `ScaleMode` 1/3）一次量出两件事。**#177 量纲**：`UserControl.TextWidth/.TextHeight` 交的是设备像素（`vb6_uc_measureText` 用 `GetDC(NULL)` + `GetTextExtentPoint32`，外层原样交出）⇒ 缇型控件里文字比同一枚控件的 ScaleWidth 小 15 倍。顺带把**第二张单位表**收掉：Fix 184 给容器侧接真实 DPI 时 `vb6rtl_com.c` 里另抄了一张 ScaleMode 表（纵向还用 `LOGPIXELSX`），现在只剩 `vb6forms.c` 的 `vb6_ScaleUnitsPerPx` 一份，`ScaleX/ScaleY` 转调它。**⚠ 顺手订正我自己 #175 里写错的两个枚举值**：新表最初把 6/7 当「厘米/毫米」还给了 `10.0/dpi` 这种没根据的系数 —— VB6 是 `vbMillimeters=6 / vbCentimeters=7`（值以 `src/semantics/builtin/builtin_consts_ext.inc:201-202` 为准）。语料没有 6/7 档控件 ⇒ 无读数变化，但**枚举值要查表、别按「数字大=单位大」猜**（与 #165 抄错 `WS_EX_CONTROLPARENT`、#163 抄错 `BM_SETCHECK` 同型）。**#178 上下文**：`vb6_uc_push` 的换入点只有四类（窗口消息 paint/mouse/show/size、HostCreate 的 init/props、Refresh、composite dump），而**容器→控件**有两条进入路都不换入 ⇒ 同一句 `UserControl.TextWidth` 在控件自己的 `Initialize` 里给 600（缇）、被容器调时给 40（像素，因为全局残值 0 被当成像素档）。早绑定那条**只能在发码侧治**（实测：给三个入口补了 push/pop 之后容器调用照旧 40 ⇒ 它根本不走桥）—— 照 rev20 给 ScaleWidth 补 `vb6_UC_ScaleWidthOf(me)` 的同一形状，在 `.ctl` 生成文件里把 `UserControl.TextWidth/.TextHeight` `#undef/#define` 成 `vb6_UC_TextWidthOf((void*)me, t)`；**选宏而不是 enter/leave 配对，是因为 `Exit Function/Property` 漏帧是已知欠账**（见 ai/031），配对会在提前返回时漏弹栈。晚绑定那条同时在 `OwnPropGet/OwnPropSet/OwnMethodCall` 三处补 push/pop。**判据**：夹具把读数写成「缇型 = 像素型 x `Screen.TwipsPerPixelX`」⇒ 与 DPI 无关；针面 `U-SW/U-TW/U-TH/U-DONE`，`Test-Vbp ve_units` + `ve_units_x86`。**红侧实测**：同一份夹具跑在 #175 的产物上是 `U-SW=True / U-TW=False / U-TH=False`；哨兵 `scripts/check_uc_scale_units.ps1` 加了两条（`vb6rtl_com.c` 里再出现 `1440` = 有人抄回第二张表；权威必须在场），对 BASE 树 `wt_base175` 报 22 处红。**CI 侧读数与本地逐字节相同**（`VeUnits.out` 两片：`I-MODE=1 I-TW=600 I-SW=2400` / `I-MODE=3 I-TW=40 I-SW=160` / `U-RAW tw=600 pix=40`）⇒ 这条判据不挑机器。**像素型控件必须一动不动，实测成立**：Charts 2020 每控件绘制缓冲颜色数落在与 #175 同一区间（pie 438/235 vs 437/238、treemap 286 vs 304），而**同一份产物连跑两遍自身就有 ±5% 抖动**（pie 438→449、treemap 286→269）⇒ 差异全在抖动内，最小值 63 稳定；czUI 子窗口 rect 逐条相同。跨门工件对形 #301→#302：**要按名字取多重集比**（62 个名字在两片里各有一份 x64/x86 拷贝，⚠ 只比 `shard/同名` 会把拷贝配错对 —— 第一版脚本就是这么把 `DlApp.out` 误报成「多了四行判据」，换成按名多重集比之后它不在差异里）：115 个名字里 **7 个不同** —— `VeUnits.err/.out`（新夹具，各 2 份）+ `FrmEvents.out`（拖放点 X，台账 #79）+ `ModalApp.out`/`TabWalkApp.out`（只有句柄十进制，`eq=Y`/`eqDeep=Y`/`clicks=3` 未动）+ `TmApp.out`/`WsApp.out`（拍数、端口、在途计数，#162 同族）⇒ **没有一条判据行变化**。分片文件数 59/40/33/45 → 58/38/38/47（新增 2 枚 `Test-Vbp` ⇒ 后续用例换片，与 #174 那条提醒同一件事）。**顺带把 #176 的读数补进 B23**：`VBFlexGridDemo` 起窗即 `0xC0000374` 是**未定序**的 —— 同一份产物连跑三次「崩/崩/活」⇒ 现在挂 `Test-GuiVbp` 只会给门加随机红，下一步是 `C3_PAGEHEAP=1` + `-g` 的 map/pdb 拿栈。STATUS 保持 IDLE。
# 上一条（2026-10-04T02:05:00+08:00）= # 本轮追加：**账 #175 = 已出（提交 `e4bf2b17` = <vbeclipse> rev40，门 #301（run 37140783159，head `e4bf2b17`，attempt 1）= 11 个 job 全绿、零新增失败）**：`tests/Charts 2020` 的图体（饼/柱/面积/矩形/环）**整幅空白**，框和标题照画。**根因不在绘制，在量纲**：`vb6_uc_push` 无条件把 `r->scaleWidth/Height`（设备像素）折算成缇交出，而语料 9 份 `.ctl` 里 8 份声明 `ScaleMode=3(Pixel)`（第 9 份 FontMemRes 没写 = VB6 缺省 1=缇，`cgen_form.cpp:381` 把缺省折成 1）—— rev18 那句「desc->scaleMode 声明的是 1」的前提是错的。实测链（`--target win-x86` + 生成码打标探针）：`GdipCreatePath/GdipAddPathPie/GdipCreateSolidFill/GdipFillPath` 四条**返回码全 0**、`Total=288`、角度 `a0=-90 sweep=32.5`、颜色 `FFE34447`、句柄全都有效，而 `mWidth=2503 / 饼直径 1973` 落在只有 **169x137 像素**的窗口里 ⇒ 排除「没画」，只剩「画到画布外」。**同一个「写死缇」的决定在 RTL 里是 12+ 处的重复决策**：控件 Left/Top/Width/Height 读写 8 处 + `vb6_ControlMove` 4 处 + `vb6_GetScaleWidth/Height` + UC 鼠标坐标（那一处还把 15 写死成 `*15.0f`）⇒ 收成一处权威 `vb6_ScalePxToUser / vb6_ScaleUserToPx`（单位取自容器声明的 ScaleMode；**缇那一档直接复用 `vb6_TwipToX/XToTwipX` ⇒ 窗体侧逐字节等价**，语料里 71 份 .frm 有 3 份声明 `ScaleMode=3` 但发码从不设 `VB6_ScaleMode` 窗口属性（`vb6_SetScaleMode` 零调用者）⇒ 窗体那一面今天恒读 1）+ 两个取法 `vb6_ContainerScaleMode` / `vb6_WindowScaleModeSelf`。**判据要能红**：`Test-GuiVbp` 加 `-DumpMinColors`，读 `C3_UC_DUMPDIR` 落下的每控件绘制缓冲数**不同颜色**（「窗口出来了 + 退出码 0」这两条对「画到画布外」全绿）—— BASE 实测 4..10（14 份里 11 份 <=10）、NEW 63..512，Charts2020 那例钉 40。结构哨兵新文件 `scripts/check_uc_scale_units.ps1`（几何文件里再出现裸 `vb6_Twip*/XToTwip*` 或 `* 15.0f` 就报红；对 BASE 树 `wt_base175` 实测 20 处红、对新树绿）。**czUI 也做了同一次 A/B**：宿主窗口（cls=V）rect 两边逐条相同（设计期那条创建路没动），而运行期 `txtEmbed.Move pad, tbTop, w-pad*2, tbH` 从 BASE 的 `561,531x415x33`（字面量被缇口径折成 0.5px）变成 `568,538x400x18`（= 源码字面量的像素语义），BASE 截图缺 Kill Switch 拨钮 / 65% 进度条 / Connect 按钮 / Location 标签四件、NEW 全在 ⇒ 这一刀把 rev18「用缇自洽」换来的 czUI 半坏布局一并修回 VB6 口径。跨门工件对形 #300→#301（同分片，四片文件数 59/40/33/45 两边一致 ⇒ 没换片）：177 份里 **167 份逐字节相同**，10 份变化逐条归因、**没有一条是判据行** —— `ModalApp.out`(s1,s4) 只有句柄十进制（`Y/N` 未动）、`TabWalkApp.out`(s1,s4) 只有句柄（`eq=Y`/`eqDeep=Y` 未动、#171 钉的 `clicks=3` 一字未改）、`WsApp.out`(s1,s2) 是端口与在途计数、`TmApp.out`(s2,s3) 是拍数 ±1（#162 同族，未钉）、`FrmEvents.out`(s4) `EV24 X=-92→-118`（台账 #79 那条按窗口位置现算的拖放点）、`oledd_test.txt`(s4) `hdrop FAIL hr=0x1 → hdrop first=C:\a.txt` —— 这一条两边都不红是因为**没人钉它**（`run_tests.ps1` 里 grep 不到 `hdrop`/`oledd`），它是某夹具自己写的旁路日志 ⇒ 记一条待办：这条值得钉。**同轮量到、刻意没混进本刀的三件**：#176 `VBFlexGridDemo` 起窗即 `0xC0000374`（堆损坏），BASE 与 NEW 同形 ⇒ 与本刀无关，且它没有 `Test-GuiVbp` 用例、CI 从来没见过；`UserControl.TextWidth/TextHeight` 仍交设备像素；窗体那一面的 ScaleMode 没接线。前两条已作为 **账 #177 + #178** 提交 `d5f3e190`（等门）。**⚠ 一条自己写错又当场订正的值**：新单位表里我最初把 `6/7` 当成「厘米/毫米」并给了 `10.0/dpi` 这种没根据的系数 —— VB6 是 `vbMillimeters=6 / vbCentimeters=7`（值以 `src/semantics/builtin/builtin_consts_ext.inc:201-202` 为准）。语料里没有 6/7 档控件 ⇒ 无读数变化，但**枚举值要查表不要按「数字大=单位大」猜**（与 #165 抄错 `WS_EX_CONTROLPARENT`、#163 抄错 `BM_SETCHECK` 同型）。STATUS 保持 IDLE。
# 上一条（2026-10-04T00:40:00+08:00）= # 本轮追加：**账 #174 = 已出（提交 `3684b1d7` = <vbeclipse> rev39，门 #300（run 37136425255，head `3684b1d7`，attempt 1）= 11 个 job 全绿、零新增失败）**. `tests/Charts 2020` 三处 `Set m_oX = SelectedControls(0)`（其中进编译清单的是 `PropPagLP.pag:445` 与 `ppProgressCircular.pag:490`）此前发成**裸 C 调用** `SelectedControls(0)` → cl 只给 C4013「假设外部返回 int」就编过。 **这不是噪声而是账 #173 那一族的下一个雷**：x86 cdecl 下被隐式声明的函数按 int 取返回值，而 `double` 返回值躺在 x87 栈 ST0 上、调用方永不 `fstp` ⇒ 每调一次漏一层栈，八层之后栈满、之后任何浮点取值得 QNaN `0x7FF8000000000000`；x64 走 XMM0 则完全静默。出口 RTL 早就位：`vb6rtl_userctl.h:182` 的 `static inline void* vb6_PropertyPage_SelectedControls(int32_t)`（属性页在 EXE 运行期永不加载 ⇒ 集合恒空 ⇒ 返 NULL 就是正确退化）与成员访问形态的识别（`cgen_expr_member_generic_access.inc:118`），**缺的只是名单里那一名**：裸名表 `kPropertyPageHostMembers`（`cgen_expr_ident_builtin.inc`，由 `isPropertyPageDesigner_` 按 `.pag` 的 controlTypeName 选表，两侧读同一位）此前只有 hwnd/changed/scalemode/scaleheight 四项。补这一名而不在调用发射处再开支路：**「哪些裸名属于宿主伪对象」这张表是单一权威**。 判据：新用例 `pp_selctrl_hostmember`（走现成的 `Test-EmitcShape`，不新建工程）钉发射面含 `vb6_PropertyPage_SelectedControls(0)`；**红侧实测** = 同一份 `.pag` 用修复前的冷构建编译器（`s174_BASE_C3.exe`，同 worktree、同配置、只差这一刀）发射 `= SelectedControls(0);` 且针 MISSING。语料 A/B（BASE vs NEW，`tests/**/*.vbp` 120 枚 `--emit-c` 逐字节对形）：**118 枚一字不动、2 枚变化、合起来 3 处发射点逐处都是这一件事**，failed=0。本构建的 C4013 计数：#173 之前 14 → #173 之后 2 → **本格之后 0**。 **⚠ 加一枚 `Test-EmitcShape` 用例会让 vbp 分片计数器 +1**（该帮手也调 `Enter-VbpShard`）⇒ 后续用例**换片**（搬迁不是删除），跨门对形时按「同名文件换分片」读。 **另一条订正（账 #172，本格没修）**：之前记的「CI 正是 Release ⇒ 门里带着那个垃圾常量」**说过头了** —— 用 `-G Ninja -DCMAKE_BUILD_TYPE=Release`（= CI 那套）冷编的 C3.exe 重发 `tests/VBFlexGridDemo`，`vb6_ret_ComboCalendarMinDate`/`MaxDate` 两处都发 `= 0.0`、全文件 grep 那个天文数字 0 命中 ⇒ 读未初始化内存这条缺陷仍成立（同一台构建里它是**确定值**），但暴露它的是**生成器**那一维（VS 生成器/Release），不是「Release」本身 ⇒ 优先级降一档，再追时别拿配置当线索。STATUS 保持 IDLE。
# 上一条（2026-10-04T00:05:00+08:00）= 本轮追加：**账 #173 = 已出（提交 `332f6363` = <vbeclipse> rev38，门 #299（run 37134633182，head `332f6363`，attempt 1）= 11 个 job 全绿、零新增失败）**：Charts 2020 在 **x86** 上启动即 `Unhandled VB6 Error #6: Overflow`（`rc=0x6`，窗口从未出现）。**rev37 那条结论「RGBtoARGB 的 Opacity 实参被喂了 OLE_COLOR」经实测推翻** —— 入口探针读 `RGBColor=-2147483635 / Opacity=100` 逐次都对；把 `CByte(` 的每一处调用点打标（`vb6_CByteT(__LINE__, …)`，TU 内联 helper，不依赖符号化）读出 `[P_CB] ucProgressCircular:1727 v=nan`，再就地取基线问增量：同一句旁存的 double 打 bits=`7FF8000000000000`，而**同帧再调 `vb6_Abs(-2.0)` 得 2** ⇒ Abs 没坏，坏在那一次 `fstp` 弹了空栈。根因＝`vb6com_invoke.c` 里 `vb6_VariantToDouble` **没有可见原型**（这一族刻意不 include `vb6rtl.h` 以免撞 VARIANT，见 `vb6com_internal.h:18`），cl 只给 `C4013「假设外部返回 int」` 就放过；x86 cdecl 把 double 返回值放在 ST0，调用方按 int 取值就**永不弹栈** ⇒ 每调一次漏一层 x87，八层后栈满，之后任何浮点取值得 QNaN。`cvtsi2sd 出不了 NaN` 是这条链的锁；x64 的 double 走 XMM0、没有 x87 栈 ⇒ **只在 x86 现形**（门跑默认架构，看不见）。修法＝在 `vb6com_internal.h` 按真实签名登记 `extern double vb6_VariantToDouble(VARIANT)`（形参用 Windows VARIANT 的依据＝rev11 立下的`vb6_VARIANT` 逐平台同布局）。同批 census 把本构建的 14 处 C4013 降到 2：另补 `vb6forms_uc_internal.h` 的内建 Collection 九个成员面（此前 `uc_hostmodel_*.inc` 全靠隐式声明）与 `vb6comserver{,_pci}.c` 的 `<wchar.h>`；剩下 2 处在**生成码**里 = `PropPagLP.c(293)` / `ppProgressCircular.c(379)` 的 `SelectedControls(0)` 裸调用 ⇒ 新账 #174（台账 B22：RTL 出口 `vb6_PropertyPage_SelectedControls` 早就位，缺的只是 `kPropertyPageHostMembers` 那一名）。**⚠ 门看不见这一格**：`[GUI] Charts2020 … PASS` 在 CI 上是真的（run 37116568387 / vbp #3 的 job 日志实证）—— 同一份代码在那台 runner 上漏的层数没撑满栈，本机必红 ⇒ 这一族的证据只能是**本地同参数两份产物的 A/B**（基线 `rc=0x6` vs 改后 x86 与 x64 都 `STATE=alive win='Form2'`、stderr 0 字节），门只回答「没把别的弄红」。另记一条**已被实测推翻的修法**：把 `vb6_Abs` 换成 `x<0?-x:x` 以避开 CRT 的 `fabs` —— 照红，症状与它无关，别再走。跨门工件对形 #298→#299（同分片）：105 份里 **100 份逐字节相同**，5 份变化逐条归因且**没有一条是判据行** —— `FrmEvents.out` 的 `EV24 X=-92→-118`（= 台账 #79 那条按窗口位置现算的拖放点）、`ModalApp.out`/`TabWalkApp.out` 只剩句柄十进制（`eq=Y`/`eqDeep=Y` 未动，`clicks=3` 一字未改）、`TmApp.out` 与 `WsApp.out` 是拍数/端口/在途计数，且**同门内跨分片对照同样在动**（`W1=534/6` vs `534/5`、`T11` 两片 51/61）⇒ 与 #162 同族，不是本格带的。产品侧零 cgen 改动（`git show --stat` = 4 个文件全在 `src/rtl/`、+31 行纯声明）。⚠ 工具事实重录一条：**数 cl 的警告必须自己重跑 cl** —— 驱动不透出 cl 的 stdout/stderr，只有 `-v` 里那行 `执行: cl.exe …` 是真命令；而 `--emit-c` 会换一个**新的 session 目录**，`_cc.cmd` 与 `version_info.res` 都得跟着搬（纯 emit 不产 .res ⇒ LNK1104）。STATUS 保持 IDLE。
# 上一条（2026-09-30T17:50:00+08:00）= 本轮追加：**账 #171 = C29-FS-h 已出（提交 `ec4fa91`，门 **#244**（run 36698127120，attempt 1））**：单选钮在**焦点进入**时也会替父窗发一条 `BN_CLICKED`，而那一刻它自己并没被勾上（在 `_Click` 到达的那一刻两头取样：`Value` 与 OS 的 `BM_GETCHECK` 都答「没勾」）⇒ 号本身就是 0，#158 那道 `code == 0` 筛不掉它；VB6 的 `Click` 只在用户真选了这枚时发生 ⇒ 多出来那条不该升成事件。口径收成一处 = RTL 新函数 `vb6_RadioClickCounts`（非 `BS_AUTORADIOBUTTON` 一律放行），发码侧只给 `OptionButton` 的 arm 补一道条件尾，`CommandButton`/`CheckBox` 一字不加（复选框取消勾选是一次正经 `Click`）。判据两头各钉一处、都在现成夹具里（不新增 Test-Vbp 调用 ⇒ 不动分片）：`btnfocus` 的 `BF-noclick` 把单选钮算进去（BASE 读 `N` ⇒ NEW 读 `Y`），`tabwalk` 的 `clicks` 从「只钉字段存在」翻成钉 `3`（BASE 读 `6`）；筛过头则读 `0` ⇒ 两个方向都能红。`AK-pre/down/wrap/up`、`optA=N/optB=Y`（勾选交接仍靠那声 `BM_CLICK`，本格没动它）、`TW-*`、`BF-each` 一字未动；x64 与 x86 各跑两枚夹具读数一致；门后跨门对形 #243→#244：`TabWalkApp.out` 两片（x64/x86）同时 `clicks=6 → 3`，其余八份是句柄号、端口、拍数与未钉计数。**别把 `BfApp.out` 一字未动当成这根针会红的证据**（旧夹具没数 optA、新夹具＋已修编译器也不发 ⇒ 两边都 `Y`），红侧证据只有「同一份新夹具 + 同配置的 pre-fix 基线编译器」那一对。**⚠ 一条判据面的教训**：`BF-noclick` 这条口径 #158 就立过，但当年只数了 cmd/chk/容器三枚、少了一型单选钮 ⇒ 这条 OS 自发的通知在 CI 里躲了整条 #158→#168 线 —— **同一条口径要覆盖每一型，少一型等于没有**（与「550 行全绿等于没护栏」同型）。语料 125 枚 `--emit-c` 对形（BASE = 同一提交 `9a62dab` 的**同配置**冷编基线）差异一共 **21 处，逐条都在单选钮的 Click arm 上**（`btnfocus` 1、`tabwalk` 2、`Charts 2020/ucChartBar` 2、`VBFlexGridDemo`+`bisect5/6/7` 各 4；看着「arm 比 .frm 的 Sub 多」的是控件数组 `Option2_Click(Index)` 一处理器两枚元素）；其余输入逐字节相同。样式针面 NEEDLE-FACE=OK（坏=0）、`ERRORS=0`。**顺带翻出一条新账 #172**（本格没修）：第一次拿 VS 生成器/Release 的基线对主树的 Debug/Ninja，凭空多出四份「Date 默认值 `0.0` 对 `-6.2774359784998866e+66`」的差异 —— 不是回归也不是噪声：**编译器自己在读未初始化内存**，Debug 恰好零填充掩盖了它，而 **CI 用的正是 Release/Ninja**（`ci.yml:59`）⇒ 门里产物带的就是这个天文数字，只是没有断言钉。已排除 `defaultValue()` 表漏 Date（那张表发字面量，不随配置变）；同一垃圾值在两处同时出现（返回局部初值 + `Select Case` 折出来的区间） ⇒ 来自一条被解析出来的默认值/区间记录。附带纪律：**字节护栏的 BASE 与 NEW 必须同配置**。STATUS 保持 IDLE。
               # 上一条（2026-09-30T12:30，门 #235 = 账 #163 已出（自研跳格导航器）+ 账 #168 量完方向键那一半）= **账 #163 = C29-FS-e 已出并过门（门 #235，head `bdb7f2a`，attempt 1 十一 job 全绿）= 跳格拿回自己手里**。同一轮里 #168 的量也问完了（探针 `.build/cp2` 加四条 flag、八份读数，`.build/grpprobe_out.txt`），三条结论都推翻了我开工前的假设：**`WS_GROUP` 是无辜的**（加与不加进站与方向键两条读数一字不差 ⇒ 不改发码）；**产品那一站是 OS 把 `WS_TABSTOP` 自己挪到单选组里勾选那枚上**（`TabStop=0` 那两份读数里 tick 2 打 `stA=50014009` 带、`stB=50004009` 不带，创建时两枚都不带）；**方向键不走是 z-order 与 TabIndex 反了**（`.frm` 里 optB 先声明 ⇒ optA 是链尾，OS 的 VK_DOWN 按 z-order 找下一枚 ⇒ 一跳就跳出容器；探针里勾选那枚是链首，同形状方向键本来就会走组）。⇒ 两张表都归一个来源，修法只剩一条。**落地（只在 RTL 一侧，`src/backend` 一字未动 ⇒ 发码面逐字节不变，`needle_face_check.py` 13 例 29 根样式针坏=0）**：`vb6_CreateControl` 把创建时那份 WS_TABSTOP 存成窗口属性 `VB6_TabStop`（**值写 1/2 不写 1/0** —— `SetPropW(…,0)` 等于删属性，#107 那一课；能这么做是因为 VB6 的 TabStop 是设计期属性、运行期不改，`TabIndex` 才是能改的那个）；新增 `vb6_Form_MoveTabFocus` + `vb6_tabCollect`（同父窗按 (`TabIndex`, z-order) 稳定排序、容器不进站但就地展开孩子、嵌容器同规则、深度上限 8、进站读那份属性 + visible + enabled）；主泵与模态泵各接一枚 `vb6_TabNavKey`，**只吞 VK_TAB**，`C3_OCX_NO_TABNAV=1` 整条退回旧行为。**判据换法是这一格真正的收获**：`TW-order` / `MW-seq` 钉**首次访问的次序**而不是逐拍读数 —— 本轮同一份产物 x64 用 22 拍、x86 用 13 拍才走齐六站，拍号只能当保险丝（与 WS17 / #162 那一族「条数是时序不是不变量」同一课，只是这次是次序版）。开/关两侧读数：`TW-order` = `cmdIn1,cmdIn2,cmdDeep,cmdTop2,cmdInPic,cmdTop1` 对 `cmdIn1,cmdTop2,optA,cmdTop1,cmdInPic,cmdDeep,cmdIn2`；`TW-orenter` N 对 Y；`MW-seq` = `txtMain,txtSecond,cmdX,cmdY` 对 `txtMain,cmdY,cmdX,txtSecond`；而 `MW-new/repeat/hops` 两侧**同值** ⇒ 改的只是次序、站点集合没动。两架构逐行相同，`ctrltabindex` / `btnfocus` 一字不动；两条负控各红各的针（NO_TABNAV 红次序、NO_DLGMSG 红 `MW-new`）—— 一个开关管一条判据，红的时候分得清是导航器坏还是泵坏。**账 #168 没结、只剩方向键那一半**：`AK-down=cmdTop1` 在导航器开关下都不变（这一格只吞 VK_TAB），下一手在同一个采集器里接 VK_UP/VK_DOWN（同父窗、同型 `BS_AUTORADIOBUTTON`、按 TabIndex 走并改勾选）。STATUS 保持 IDLE。
               # 上一条（2026-09-30T11:40，门 #234 = 账 #165 已出（发码那一格的手抄常量）+ 账 #168 探针量完）= **账 #165 = C29-FS-d 已出并过门（门 #234，head `b88432a`，attempt 1 即 11 job 全绿）= 根因是一个手抄的常量**：容器那一位 `WS_EX_CONTROLPARENT` 在单一权威 `controlContainerExStyleBit` 里被写成 `0x00040000`，而 SDK 头 `winuser.h:2855` 里它是 **`0x00010000`**（`0x00040000` = `WS_EX_APPWINDOW`）。改成对数之后 x64 与 x86 实测：跳格**真的走进容器**，一圈七站全到达（`cmdIn1 → cmdTop2 → optA → cmdTop1 → cmdInPic → cmdDeep → cmdIn2`），`TW-in2 / TW-deep / TW-inpic` 三条从缺陷读数翻成判据；`modal`（`MW1..MW4` 逐字不变）、`ctrltabindex`、`btnfocus` 三套一字不动。**这一格最贵的一课不是那个数，是两把仪器都撒了谎**：① 上一轮「焦点史问完 ⇒ OS 本人不认容器里的兄弟」的结论**作废** —— 产品压根没挂上过这一位，摘不摘都一样，那条对照是假的；② 「探针与产品逐条对形」时探针打 `exGrp=65536`（SDK **符号**）、产品打 `EX-fr=262144`（抄错的**字面量**），**字段名对上了、数值没对上**，而台账写的是「样式位对上了」⇒ 那十条排除表全是在错误前提下做的。通用式（已进 memory 两条）：**手抄常量一律去 SDK 头对值、证人行打印符号对应的值**；**两侧对形要把被比较的数值本身并排读出来**。**夹具自己的第二条错也一并订正**：那三条恒 N 的直接来源是相 2 只给 9 拍、而一圈实测要 10–13 拍 ⇒ 「走不到」其实是「没走到」；现在出口改成「七站齐了就收」，拍号退化成保险丝（40 拍），另把 `TW-new` 与 `gSeen` 条数差 1（tick 1 的 `DoEvents` 重入本 Timer 先记一站又被抹掉）一起归零。针面：3 条运行期判据 + 8 根样式针（`262144 → 65536`，反面四条换成当前真实样式串），哨兵 `.build/needle_face_check.py` 13 例 29 根全对（坏=0）、PS `ERRORS=0`，负控 = BASE 编译器（错常量）跑同一份新夹具 ⇒ 三条一起红。**跨门同分片工件对形（#230 → #234）**：四片里只有 5 份 `.out` 变了 —— `TabWalkApp.out`(v1/v4，本批) + 已知噪声四份（`ModalApp.out` 句柄十进制与 `busy=7→6`、`TmApp.out` 计时器拍数 49/50/0/1 = 账 #162 那族、`WsApp.out` 端口号、`FrmEvents.out` OLE 拖放 X 坐标 -92→-118 = 任务 #79），其余逐字节相同。**路线再改一次**：#165 不需要自研导航器；剩下 **#163**（次序 = z-order，这条才需要 `vb6_MoveTabFocus`）与 **#168**（新读到：容器里 `TabStop=0` 的单选组反被当一站、落在勾选那枚上，而组内方向键不走 —— 发码侧 `WS_GROUP` **一枚都没发**，census 按对数：`WS_GROUP`=`0x00020000`（winuser.h:2802）在 src 全域零命中 —— 第一版扫的 `0x00002000` 其实是 `ES_DISABLENOSCROLL`，同一天第二次撞上「常量的数没去头里对」），读数 `TW-orenter` **只打不钉**。下一格 = #168（先探针量 `WS_GROUP` 的两条后果，再决定发哪一位）。STATUS 保持 IDLE。
               # 上一条（2026-09-30T09:55，本轮不起门 = 账 #165 只测量、结论后来作废）= **账 #165（容器跳格）本轮只测量、零产品代码改动 —— 但把这条账从「根因未明」推到「分岔问完」**：结论是 **OS 本人不认这棵树上的容器兄弟**，不是我们收到通知后又把焦点搬走。仪器没补桩、没动共享 RTL —— 让控件自己报焦点史（#158 之后按钮族 `BN_SETFOCUS` / `BN_KILLFOCUS` 是通的）：新做最简拓扑排障夹具 `.build/twmin`（一枚 Frame + 框里两枚按钮 + 框外两枚 + Timer，与裸码探针 `.build/cp2` 一一对形），读到 `got-top1 / lost-top1+got-in2 / lost-in2+got-top2`，**全程没有一次 `got-in1`** ⇒ `IsDialogMessage` 从容器里的 `cmdIn2` 直接跳到窗体级 `cmdTop2`，跳过了同一个父窗里 `GW_HWNDNEXT` 上的 `cmdIn1`。外圈候选至此**十条全部实测无罪**：`WS_GROUP`、创建顺序、容器挂纯转发子类、子类同步重发 `WM_COMMAND`、窗体自己带 CONTROLPARENT、Common-Controls 6.0 清单、孩子带 `BS_NOTIFY`、容器/孩子/两边带 `WS_CLIPSIBLINGS`、整棵树在 `ShowWindow` 之后才建、首位不可见子窗口（产品那枚 `VB6_TIMER`）、容器在窗体子链的**链尾**（探针新加 flag 131072；同一次读数 `Z-formkids=T1,T2,GRP,none` 证明真挪过去了）；两边 USER32 视图逐字对形（顺带记正一对常量：`GW_CHILD=5`、`GW_OWNER=4`，拿 4 当 5 会一直打 `none`），`GetParent(窗体)=0` 也排掉「窗体上面还压一层外壳窗」。⚠ **本线第九次探针自伤**：第一版想在同一拍内 `DoEvents` 四次再立刻读 `GetFocus()`，结果现象整个消失 —— `DoEvents` 走 `PeekMessage + DispatchMessage`，**绕开了模态泵里的 `IsDialogMessage`**，post 出去的 VK_TAB 被直接派给按钮消化掉 ⇒ 通用式：**在别人的消息循环内部插一次自己的泵，会改变被测的那条路**；要看事件史就让被观察对象自己报告，别在外面代打泵。台账里原先写的下一手（补 `vb6_di_GetNextDlgTabItem` 单桩再到产品进程里问）**作废** —— 同一个问题用产品自带的事件面就问完，零共享改动。**路线因此改变**：剩下的「为什么」要往 USER32 内部刨（钩子 / Spy 一类外部仪器），性价比低；而 #163 早记着一条自研 `vb6_MoveTabFocus`（按 TabIndex 递归走控件树，自己处理 VK_TAB / Shift+VK_TAB），一旦落地**同时**结 #163（顺序不再按 z-order）、#165（容器里的控件进得了序）、#168（组内方向键自己实现 —— OS 的组语义本来就没给我们），且不再依赖对话框管理器黑盒。那是特性级改动（RTL + 发码两侧），**等用户点头再开工**；在那之前 `tests/tabwalk` 的 `TW-in2 / TW-deep / TW-inpic` 继续钉 **N**（缺陷读数，导航器落地那天三条一起翻 Y，不许提前改针）。门：本轮无代码改动 ⇒ 不起门，GATE_BASELINE 仍停 **#230**；本轮五个纯文档提交（`7a7fa94` `0e19cfa` `fcc86d6` `d2fec52` `1c19bb0`）按「没有源码级变动不跑 Actions」留本地，随下一次源码推送带走。
               # 上一条（2026-09-30T08:35，门 #230 = 账 #164 已出、同一格里另半条分出 #168）= **账 #164（C29-PT-a）= PictureBox 不再进 tab 序**：`controlTabStopStyleBit` 的排除表里有 Label / Image / Shape / Line / Frame / Timer / Menu，**就是漏了 PictureBox** ⇒ 两条创建路都给它立上 `WS_TABSTOP`，跳格在它上面停一站，而 VB6 的 PictureBox 拿不到焦点、连 `GotFocus` / `LostFocus` 这对事件都没有。先量机制再改（裸码探针 `.build/picstop/picstop.c`，同一棵 `Button / STATIC(SS_BITMAP) / Button` 的树只差这一位）：挂着 ⇒ `IsDialogMessage(VK_TAB)` 走 `B1 → Static'PIC' → B2 → B1`，摘掉 ⇒ `B1 → B2 → B1` ⇒ **修法就是排除表补一类**，派发与 RTL 一行不碰；`.frm` 真写了 `TabStop` 的那条分支照旧优先（#83(a) 的"写了照发"口径不动）。判据两头：运行期（顶层路）`tests/tabwalk` 的 `TW-picstop` 翻成 `Y`、`TW-seq` 里 `pic1` 那一站消失；容器里那条路**没法用跳格证**（#165 未结，对话框管理器不下钻）⇒ Frame2 补一枚 `picDeep`，按实测数钉发码针（顶层 `1417675278L, 262144L,` / 容器 `1350566414L, 262144L,` 正面，两串带 `WS_TABSTOP` 的旧数进 `tw_emitc_notplain` 反面）。**门 #230 的工件自带一次前后对照**（同一片 shard、同一台 runner，门 #227 vs #230）：`TabWalkApp.out` 只有四行动 = `TW-picstop` N→Y、`TW-seq` 少 `pic1` 一站、`TW-new`（非判据）与两行按**句柄数值**打印的证人行；其余 11 行逐字节相同。整片 vbp-1（49 个文件）只有 `TabWalkApp.out` + `ModalApp.out` 两份动，而 `ModalApp` 动的两行是 `M1-startup=Y/second=N` 括号里的句柄数与 `busy=14→7`（**结论位一位没动**）；vbp-2（39 个文件）只有 `TmApp.out` 的 `T11=Y/51→50`（拍数，账 #162 那一族）动，邻居 `BfApp.out` / `CbApp.out` / `TiApp.out` 与上一轮**逐字节相同**。本地另备一份架构对照：x64 与 x86 两份日志把两行句柄数抹掉后**逐行相同**，与 CI 那份比只差 `TW-new` 一格（同一个 exe 本地连跑三次读过 3、3、4 ⇒ 这一条已就地标注成非判据，别再拿它当不变量；CI 里 x86 那份根本取不到——两架构写同一个 `TabWalkApp.exe` 名，工件里只剩一份，同 #156 那条已知限制）。**连带修好的一格**：48 件语料逐行 `--emit-c` 归因（`.build/pic_guard.py`）改动 2 件 / 4 行 = 创建样式 `-65536` 三处（tabwalk 的 `pic1`、`picDeep`、ctrltabindex 的 `picBox`）+ **一处 `vb6_Form_SetInitialFocus` 换了目标**（`picBox → cmdIn0`）：#157 那档"把焦点交给 TabIndex 最小且拿得到焦点的控件"以前会选中一枚拿不到焦点的静态窗口，现在跳过它 ⇒ 同一处排除表带来的**修正**，不是副作用；`UNEXPLAINED=0`，负控（`WS_TABSTOP` 换成 65537 + 去掉焦点目标那条规则）报 `UNEXPLAINED=8` ⇒ 护栏能红。邻居 `ctrltabindex` 四条运行期读数 BASE 与 NEW 逐字相同（CI 那侧 `TiApp.out` 也逐字节相同）。**流程**：本批提交前先跑了 #158 那批补下的针面哨兵 `.build/needle_face_check.py` —— 就是门 #225 红一条时缺的那一步；它先在本地报出 `tw_emitc_cparent` 的 `1417740814L, 262144L,` 找不到，改完再跑才 `坏=0`。手册 `PictureBox 控件.md` 补"本项目的实现现状"一格（原生 `STATIC`+`SS_BITMAP|SS_CENTERIMAGE`、不进 tab 序、可当容器但子控件目前仍走不进跳格 = #165）。本线接下来：#165（对话框管理器不下钻容器，根因未明）→ #163（跳格按 z-order 而非 TabIndex，#160 已把 TabIndex 发下去了，可以做自己的 `vb6_MoveTabFocus`）→ #158 剩下的按钮族焦点通知已随 `BS_NOTIFY` 出完，#161 / #167 / #159 仍在账上；**#164 只结了一半**——同一格当初还照出「容器里 OptionButton 组内方向键不走」（门 #230 工件里那一行与 #227 逐字节相同=没碰），已另开 **#168**，与 #165 很可能同根。
               # 上一条（2026-09-30T07:05，门 #227 = 账 #158 按钮那一半已出）= **账 #158（按钮那一半）= C29-BN-a 已出并过门（门 #227，head `69845a9`；前一轮门 #225（head `db576da`）红一条，红的是针不是行为））**：按钮族的 `_GotFocus` / `_LostFocus` 一直不触发，但**不是码表的错** —— 发码那两支撑的 `BN_SETFOCUS=6` / `BN_KILLFOCUS=7` 本来就是对的，缺的是创建时没挂 **`BS_NOTIFY`**（`0x4000`=16384，实测自 SDK 头）。裸码探针 `.build/bnnotify/bnnotify.c`（同一父窗两枚按钮只差这一位）：挂了的那枚，**程序化 `SetFocus` 与 `IsDialogMessage(VK_TAB)` 两条路都把 6/7 送到父窗**；没挂的一条不送，而 `BN_CLICKED=0` 不需要那一位（⇒ 按钮的 `_Click` 一直是通的）。上一条收线时写的"Windows 对程序化 SetFocus 根本不发"就此作废：**是探针自己缺样式位造的冤案**，教训写成一条通用的——**定罪 OS 之前先问"我这枚探针的窗口创建样式位齐不齐"**；同批第二手：探针第一版把子类别链接到 `DefWindowProc`，连 `BN_CLICKED` 都收不到，看着像"父窗 `WM_COMMAND` 这条路本身断了"，子类化探针必须 `CallWindowProc` 转回原 BUTTON 窗口过程。三处一起动（缺一处就是本线反复踩的"只接一头"）：① 新出口 `controlButtonNotifyStyleBit`，两条创建路各调一次（顶层 + 容器子控件，与 #83(a) 的 `WS_TABSTOP`、#83(b2) 的 `WS_EX_CONTROLPARENT` 同族），Frame 是 `BS_GROUPBOX` 不在名单里；② 同一次发码读数量到**焦点码表那一趟只遍历顶层控件**（ID 100+），容器子控件（ID 200+）**一条 arm 都没发过** ⇒ 收成 lambda `emitFocusArms158` 两处共用，容器侧遍历口径逐字照抄已被验证的 `walkCmdChildren142`（跳 Menu 与无 Win32 类、同一 DFS、ID 从 200 起）；③ 挂了 `BS_NOTIFY` 之后同一个 id 会携 `code=6/7` 进来，而按钮的 Click arm 原本**只看 id** ⇒ 一次焦点移动会被算成一次点击（本批自己引入的 hazard），那条 arm 现在按 `code == 0` 过滤。夹具 `tests/btnfocus`：七拍各一次程序化 SetFocus（顶层三型按钮 + TextBox，容器里再来按钮 + TextBox），**最后一拍落到没挂任何处理器的 cmdGo** ⇒ `BF-each` 问"恰好一次"不是"至少一次"。A/B：BASE `BF-cmd/chk/opt/in/intx` 全 `0/0`、`BF-tx=1/1`、`each=N` → NEW 六条全 `1/1`、`BF-both=Y`（两条创建路读同一串数）、`BF-noclick=Y`、`BF-each=Y`；x64 与 x86 两份日志 cmp 逐字节相同；发码针正 6 / 反 3（反面钉的正是 BASE 的三个旧形状 ⇒ 针能红）；邻居 tabwalk / combofocus / ctrltabindex 用新编译器复跑，读数与台账逐字相同。**爆炸半径**（48 件语料 vs BASE 逐行 `--emit-c`，`.build/bn_guard.py`）：6 件 / 42 行全归因 = 创建行 `+16384` 35 处 + Click 加过滤 3 处（**全在本批新夹具，语料里一件按钮 `_Click` 都没有**）+ 容器新增焦点 arm 4 条，`UNEXPLAINED=0`；把 `BS` 换成 16385 时报 `UNEXPLAINED=35`（护栏能红）；另一条按类名归因的独立脚本证明 30 处样式位改动**全部落在 `"BUTTON"` 那一段**，`EDIT`/`STATIC`/通用控件一行没动。**新开一条纸面缺口 #167**（本批没动，也没让它变坏）：创建侧的容器清单是 `Frame/PictureBox/SSTab`，而派发侧两趟都只认前两类 ⇒ 谁把控件放进 SSTab 页里，事件 id 会整体错位；全仓没有任何 `.frm` 往 SSTab 里放控件（grep 零命中）⇒ 零覆盖的潜在分歧，开工前要先补夹具把"派发认到的 id"与"创建写的 id"逐个对上。【本批自己的一手】门 #225 红的那条是 `cs_emitc_tabstop`（账 #83(a) 那格）——它把创建样式数**逐字**钉着，而本批给按钮挂了 `BS_NOTIFY` ⇒ 三串数字全部挪了一格（顶层 CheckBox `1409351683L`→`1409368067L`、Frame 里的 `1342242819L`→`1342259203L`、显式 `TabStop = 0` 那枞 `1409286147L`→`1409302531L`；每串在两份 emit 里数过次数，旧 4/3/1→ 0、新 0→ 4/3/1，一对一）。**流程上真正缺的一步补成了哨兵**：`.build/needle_face_check.py` 扫 harness 里所有 `\d{6,11}L, \d+L,` 形状的针、按它所属工程的 `--emit-c` 逐条对正面/反面——改"每个控件都有的样式位"时，光对"哪些行变了"不够，变掉的那一行可能正是某根针钉着的那一行，而针面散在几十处 `Test-EmitcShape/Absent` 调用里。对现在的编译器报 `坏=0`、对 BASE 报 `坏=3`（正是上面三串）⇒ 哨兵能红；以后再动样式位，这一条与逐行归因一起跑。另加两条反面针 `1409302535L` / `1409302784L`（Frame / Label 挂上 `BS_NOTIFY` 就是排除表被改坏）。
               # 上一条（2026-09-30T05:55，门 #224 = 账 #158 的 Combo 那一半已出）= **账 #158（Combo 那一半）= C29-CB-a 已出并过门（门 #224，head `a54193a`）**：ComboBox 的 `_GotFocus` / `_LostFocus` 从项目能编事件的那天起就没触发过 —— 同一段派发里四张焦点码表，`EN_=256/512`、`LBN_=4/5`、`BN_=6/7` 三张是对的，**只有 Combo 这一格写成 `code == 1024 / 2048`**，那是 `CBEM_SETTBILLOS/CBEM_KILLTBILLOS`（**发给** ComboBoxEx 的消息号，永远不会作为 `WM_COMMAND` 的 notification code 出现）⇒ 真码 `CBN_SETFOCUS=3` / `CBN_KILLFOCUS=4`，一行常量，不是口径分歧。夹具 `tests/combofocus` 四条读数各管一件事不许省：`CF-cb1`/`CF-cb2` 是这一刀（BASE `0/0`、`0/0/1` → NEW `1/1`、`1/1/1`），`CF-lb`/`CF-tx` 是**邻座没被带坏**（同一种"焦点通知走父窗 WM_COMMAND"的机制），`CF-cross` 钉住 `LBN_SETFOCUS=4` 与 `CBN_KILLFOCUS=4` **同号**、全靠左边 `id ==` 过滤分开那条串台路，`CF-each` 问"每条通知恰好一次"（双发是本线已知缺陷族 #161）。`cb2` 刻意多挂一条 `_Validate`：带 Validate 时 LostFocus 走"先问 `VB6_ValidateCancel`"那份发码，不覆盖就等于那条分支从没跑过。发码正 3 / 反 2（反面钉 `code == 1024`、`code == 2048` 在这本工程里再出现不了；`1024` 只在 RTB 的 `EN_UPDATE` 那条上才是对的，那行没动）。x64 与 x86 两份日志 `cmp` 逐字节相同；**爆炸半径 = 存量 0 件**（44 件语料比出来的变化全是 #160 的 `vb6_SetTabIndex` 插入，UNEXPLAINED=0；全仓 grep 的 Combo 焦点通知只命中本批新夹具）。**#158 剩下那一半的定罪被本批推翻了一半**（顺手量完，形状改，未开工）：上一条写的"Windows 对**程序化** `SetFocus` 根本不发 `BN_SETFOCUS`"是**探针自己缺样式位**造的冤案 —— 裸码探针 `.build/bnnotify/bnnotify.c`（同一父窗两枚按钮只差 `BS_NOTIFY`）实测：挂了那一位的按钮，程序化 `SetFocus` 与 `IsDialogMessage(VK_TAB)` **两条路都发** `code=6/7` 到父窗，没挂的一条不发（`BN_CLICKED` 不需要那一位 ⇒ 缺位只影响焦点）。全仓 grep `BS_NOTIFY` **零命中**（`LBS_NOTIFY` 给 ListBox 挂了，注释还写着"列表框同款"）⇒ 按钮族那一半从"子类化自发"改成"**类型样式位补一位、两条创建路共用一个出口**"，与 #83(a) 的 `WS_TABSTOP`、#83(b2) 的 `WS_EX_CONTROLPARENT` 同族，样式位改动照 #83(a) 那条护栏走（针面重算脚本化 + 逐行严格对上 + x86 真编真验）。第二手教训同批记下：探针第一版把子类别链接到 `DefWindowProc`，连 `BN_CLICKED` 都收不到，一度把结论带歪 —— 子类化探针必须链回原类窗口过程。
               # 上一条（2026-09-30T05:32，门 #223 = 账 #160 已出）= **账 #160 = C29-TI 已出并过门（门 #223，head `eda3906`）**：设计期 `TabIndex` 现在真的下发到窗口了。RTL 那一对（`vb6_GetTabIndex`/`vb6_SetTabIndex`，存窗口属性 `VB6_TabIndex`，没存过时 getter 答 0）与 backend 的属性表映射**早就齐**，缺的只有"创建时没人发这一句" ⇒ `.frm` 写着 `TabIndex = 5`、运行期读回 0，全工程每一枚都一样（#157 那格的注释就是这么写的，本批一并改掉）。新出口 `emitDesignerTabIndexProp`，**两条创建路各调一次**（顶层那一趟 + 容器子控件那一趟）。口径三条：① 写了照发，且**每个父窗自己从 0 编号**（框架里的控件不占窗体那一串的号 ⇒ `TI-explicit=5/2/1` 与 `TI-in=0/9` 同时钉住"照写的发"和"不跨容器连续编号"）；② 没写用**同一父窗内的声明序号**兜底（VB6 存盘从不省这一行；兜成 0 会让好几枚同时声称 0 ⇒ 兜底档刻意钉在能分辨的位置：`TI-fallback=3/1/5`）；③ 无窗口的（Timer/Menu/ImageList/CommonDialog/Data/Unknown）不发 —— 反面针两条钉住 Timer 与窗体自己那枚 `hwnd`。一条自伤在实现里就避掉了：顶层那条循环有多处 `continue`，**自增必须放循环体第一句**，放末尾会让后面所有控件集体少一号。`TI-runtime=40/Y` 钉发码期那一句没盖死运行期 setter。x64 与 x86 两份 `tabindex.log` `cmp` 逐字节相同。**爆炸半径**（BASE = `.build/post_83b2_C3.exe` = HEAD 那一版码，44 件语料 `--emit-c` 逐行比，`.build/ti_guard.py`）：**changed-files=12、纯插入 86 行、UNEXPLAINED=0**；同一条脚本换成更早的 BASE（`pre_157_C3.exe`）时报 **UNEXPLAINED=24** ⇒ 分类器不是空转的（这条护栏能红）。存量里没有任何用例读 `.TabIndex`（全仓 grep 只命中本批这枚新夹具）⇒ 不动别人读数。**与 #163 接上**：运行期现在有号了，"按 VB6 的 TabIndex 跳格"不再是做不出来；`emitFormInitialFocus` 仍留在发码期算（两处注释已按这个口径改过 —— 挪到运行期要连带搬排除表与"每父窗自己编号"，不值当）。
               # 上一条（2026-09-30T04:41，门 #222 = 账 #166 已出）= **账 #166 = TEST-1 已出并过门（门 #222，head `77c9d8c`，attempt 1 十一 job 全绿，wall 562s = 9m22s；顶上叠着 #83(b2) 收线那笔纯文档提交 `0cbfd0e`，同一轮门里）**：产物名原本有**两个权威** —— 驱动落盘按 `.vbp` 的 `ExeName32`（`driver_compile.cpp:88`：优先它的 stem，缺键才用 vbp 文件名），而 `Test-GuiVbp` 默认按**文件名**、`Test-Vbp` 连 `-ExeName` 口子都没有、`Test-VbpDll` 同样。收成唯一出口 `Resolve-VbpExeBase`（显式覆盖 > `ExeName32` 的 stem > 文件名 stem；VB6 这个键可写 `Foo` 也可写 `Foo.exe`，两种都折成 stem），三个调用点一起改过去，`-ExeName` 降级成显式覆盖；`no exe` 处加哨兵把"按什么名字找的 / 文件名 / 声明的 ExeName32"一并打出来。**爆炸半径先做 census 再动手**（`.build/registered_vbp_census.py` 扫 harness 里注册的 99 个 vbp 路径）：解析值≠文件名的只有 `prjBalloonTooltips`（本来就传 `-ExeName "BalloonTooltips"`，覆盖值与新解析值相同 ⇒ 存量不变）和新增证人本身 ⇒ **对存量注册用例的产物名解析改动 = 0 件**。**能红的针自己造**：`tests/exename/NameProbe.vbp` 故意 `ExeName32="RenamedProbe.exe"`，本地实测目录里只产出 `RenamedProbe.exe`（没有 `NameProbe.exe`）⇒ 旧口径按文件名找必红 `FAIL (no exe)`，新口径 PASS。**CI 收线读数**：`test-logs-vbp-4` 里有 `RenamedProbe.out`（30 字节，`NP-ok=Y/tick=1` + `EXENAME-DONE`），四片工件里**没有任何** `NameProbe.*` ⇒ 权威确实只剩一个。resolver 的五件真 vbp 单测（tabwalk=TabWalkApp / probe=RenamedProbe / 显式覆盖=Whatever / 无键=M6Test / balloon=BalloonTooltips）与 `run_tests.ps1` 词法检查（TOKENS=19651 ERRORS=0）都在本地做过。**通用式两条**：① 改"harness 里到处都有的一份口径"之前先做**枚举面 census**，别靠"跑一遍看看"；② 门要看 **head_branch + head_sha** 归属，`gate/merge-dev` 那两轮红（#219 `43845ca`、#221 `807e98c`）是别人的合并门，而 watch-gh-actions 会自己提前 exit 0 ⇒ 结论一律回 Runs API 复核。
               # 上一条（2026-09-30T04:23，门 #220 = 账 #83(b2) 已出（发码那一半）、新账 #165 #164 #166）= **账 #83(b2) = C29-FS-c 已出并过门（两轮：`#218`（head `7be26c9`）红两条 = 夹具自己的命名 ⇒ 对齐（`5922033`）⇒ `#220` attempt 1 十一 job 全绿，wall 563s = 9m23s）**：容器（Frame / PictureBox / SSTab）的创建参数里补上 `WS_EX_CONTROLPARENT`，收成 `controlContainerExStyleBit` 一处判断、**两条创建路共用**（顶层那条以前恒写 `0L`，容器子控件那条硬写 `"L, 0L,"` —— 容器嵌在容器里走的正是第二条）。**这一格只结了发码那一半，行为那一半翻出新账**：运行期读数证明那一位真落到窗口上（`EX-fr=262144 / EX-f2=262144 / EX-pic=262144`）、容器里的子控件 visible+enabled+`WS_TABSTOP` 齐全、父链与 `GetActiveWindow()` 都对、RTL 打点确认 `IsDialogMessage` 认下那条 VK_TAB（`handled=1`）—— **可跳格仍然进不了容器**（站点序列照旧 `cmdIn1 → cmdTop2 → cmdTop1 → pic1`）。裸 Win32 探针（`.build/cp2`）同一棵树会走 `A → B → T1 → T2 → A`，五种候选逐个实测排除：容器带 `WS_GROUP`、容器孩子创建顺序、容器挂 comctl32 `SetWindowSubclass`（`subRet=1` 真装上）、窗体自己也带这一位（产物里改 RTL 试过 `form=0x40100`，读数一字不变）、Common-Controls 6.0 清单（`mt.exe` 挂上照旧会走）。没查完的是**产物窗体每枚一个注册类 + WndProc 由生成代码接管** ⇒ 根因记 **账 #165**。针面按"能不能红"分两层写死：能红的是发码五条（三型容器各一条 + 普通控件两条反面，摘掉出口或只接一头必红）；运行期 `in2/deep/inpic` 钉的是**当前的 N**（缺陷读数），注释写明 #165 落地那天必须翻成 Y，不许顺手改成 Y 当已通过。爆炸半径 BASE = HEAD~1 干净 worktree 冷编：`changed-files=5、container-bits=8 行、UNEXPLAINED=0`（脚本 `.build/fs_guard.py`，逐件 ctrlstate/ctrlshape/ctrlsstab/ctrlslider 各 1 行 + tabwalk 4 行）。白捡两条独立缺陷：`TW-picstop=N` = PictureBox 自己占一跳（`controlTabStopStyleBit` 排除表里没有它，VB6 里它拿不到焦点）、`AK-pre=optA/down=cmdTop1/optB=N` = 容器里的 OptionButton **组内方向压根不走** ⇒ **账 #164**。**通用式两条**：① **验到"必要位发出去了/落到窗口上了"绝不等于验到"行为修好了"** —— 判据要问用户看得见的那个动作，并且用探针去问"这一位够不够"，别把文档里的必要性当充分性；② `IsDialogMessage` 这条路给不了 VB6 口径（顺序是 z-order=#163、容器还进不去=#165），自研 `vb6_MoveTabFocus`（按 `TabIndex` 递归，前置 #160）现在看是正解不是备选。**门 #218 那两条红又是夹具侧**：`Test-Vbp` 默认按 vbp 文件名找产物，而产物名由 `.vbp` 的 `ExeName32` 决定 —— 两个权威管同一个名字，本地手动跑（跑的是那个不一致的名字）完全看不出，CI 两条架构一起 `FAIL (no exe)`；本批先把命名对齐，**单一权威 + census 自检记成账 #166**（census：120 件 vbp 里 56 件没写该键、59 件与文件名一致、5 件真不一致，注册用的只有 `prjBalloonTooltips` 靠 `-ExeName` 兜住）。
               # 上一条（2026-09-30T03:10，门 #217 = 账 #83(b1) 已出、账 #162 已结）= **账 #83(b1) = C29-FS-b 已出并过门（两轮：`#216`（head `b6d6fcd`）红一条 ⇒ 那条红是我自己的夹具 → 加固（`35e27be`）⇒ `#217` attempt 1 十一 job 全绿，wall 610s = 10m10s）**：主（非模态）那条泵现在每条消息先问一次 `IsDialogMessageW(GetActiveWindow(), …)`，与模态那条**共用同一个开关** `C3_OCX_NO_DLGMSG=1`（`IsDialogMessage` 会吞掉它处理掉的那条按键消息，必须留整块关掉的退路；语料里 `_KeyDown` 只有两处、没有一处按 `VK_TAB` 写判据）。判据接在 `tests/modal` 第二个相位（`MW1..MW4` + `MW-new=4/repeat=txtMain/hops=3`），负控 = 同一份产物关开关 ⇒ `MW-new` 回 1。**顺序故意不钉**：实测站点序列 `txtMain → cmdY → cmdX → txtSecond → 回 txtMain`，而 TabIndex 是 1→2→3→4 的另一头 —— 对话框管理器走 **z-order**，那一刀另开 **账 #163**。**门 #216 那条红 = 判据自伤（本线同一形已经第五次往上数）**：`tMain.Enabled = False` **不取消已经在途的 tick**，而那些拍是在 `ModalDlg.Show vbModal` 的循环里被排空的，旧 handler 一看 `gState` 就当成「下一相」，径直往模态窗体 post `VK_TAB`，把它的初始焦点挪走（`D1-first=N`、`D2-notB=N`）；x64 恰好没赶上那一拍就全绿 ⇒ **同一份产物两个位数读数不一致**，只有 x86 复现得出。修在夹具侧（`gBusy` 闸 + 把挡掉的拍数显式打进读数 `/busy=`）。**顺带结掉账 #162**：三条读数对上之后「在途积压」就是成立的那条 —— 门 #215（无闸）两片分别 `ticks=21`｜`ticks=1`，门 #217（有闸）同两个位置 `ticks=1/busy=7`｜`ticks=1/busy=0`，本地 x64 `busy=0`、x86 `busy=1` ⇒ 那 20 拍没消失、是被**数出来**了，条数看负载，判据一条没受害。**通用式：`Enabled = False` 只保证「从下一次起不再排」，不保证「这一拍不会再来」** —— 凡是靠计时器相位驱动的夹具，handler 头上一枚相位闸是必需件不是加固。半 (b2)（容器不带 `WS_EX_CONTROLPARENT`）本批只量了改前 baseline（新夹具 `tests/tabwalk` + 产物 `.build/twbase`：从 Frame 里的 `cmdIn1` 起步只走得到 `cmdIn1, cmdTop2, cmdTop1`，`cmdIn2` 与嵌在 Frame 里那枚 Frame 的 `cmdDeep` 一步不到），接线下一批出。
               # 上一条（2026-09-30T02:20，门 #215/#216 → #217 = 账 #83(b1) 已出、账 #162 已结）= **账 #157 = C29-FS-a 已出并过门（两轮）：`#214`（head `39c2e66`）attempt 1 红两条 → 补两刀（`c7b05f2`）→ `#215`（head `9d33dc9`）attempt 1 只剩一条红（`NewTab` 5 秒内没出窗），**同一 head `rerun --failed` 的 attempt 2 十一 job 全绿**（逐 job bas#2 375s / bas#1 332s / vbp#3 271s / vbp#4 251s / vbp#2 250s / vbp#1 214s / Build 207s / asm 75s / smoke 46s / syntax 33s / compile 21s）⇒ 那条判为 CI 侧偶发（本地 x64/x86 × BASE/NEW 四种组合都在 5 秒内出窗）**。**先把 s-0 的第 2 条推翻再动手**：模态窗体显示后 `act=fg=root`（本来就是活动/前台窗）、Timer 拍里 `SetFocus` 真落地（`eqA=Y`）⇒ "父窗没激活所以 SetFocus 失败"是推的；真实缺口只有"显示时没人把焦点交给第一枚 tabstop"。做法：发码期递归控件树选 **TabIndex 最小且拿得到焦点**那枚（排除口径复用 `controlTabStopStyleBit`，再剔设计期 `Enabled=False`/`Visible=False`），WM_CREATE 发一句 `vb6_Form_SetInitialFocus`，RTL 存成窗口属性、`vb6_ShowForm` 激活后 `RemovePropW` **即应用即销**。只能在编译期选：设计期 `TabIndex` 从没下发到窗口（生成代码里 `vb6_SetTabIndex` 一处都没有 ⇒ **#160**）。判据 `tests/modal/`（新夹具两张窗体）：`ModalDlg` 把四条排除一次摆出来（TabIndex 0=Label / 1=显式 `TabStop=False` / 2=禁用 / 3=藏起来），**赢家 cmdA 是 TabIndex=4 且 .frm 里最后声明的** ⇒ "照创建顺序挑"会给出 cmdB，所以 `D1-first=Y` 一根针钉住四条排除 + 顺序；`M1-startup` 另问非模态那一形。负控（BASE=`pre_157_C3.exe`）`M1=N`、`D1=N`，D2..D7 在 BASE 上**也是 Y**（那六枚是防伪证，不证明本批这一刀 —— 老实写明）。读数全走 `ByVal As LongPtr` 形参助手，绕开控件 `.hwnd` 的装箱（**#159**：`CStr(ctl.hwnd)` 空串、与 `GetFocus()` 比较恒假）。**门照出来的两条各自归因**：① `NewTab` `C2065 'vb6_hwnd_txtSearch_0'` —— 我给控件数组按别的控件的命名推了个全局名，仓里早有出口 `ctrlHwndExprForInit`（数组句柄在 `vb6_arr_<名>` 里）；② `ctrlrichtextbox` 的 RT84/RT85 —— 那枚窗体 TabIndex 最小的就是 RichTextBox，现在它在显示时拿到焦点，而**带着焦点**时一次 `SelText=` 从两条通道各发一条 Change（仪器副本 `X84=2/instr=1`，不聚焦是 1）⇒ 那是"两条通道"族的翻倍形态，另记 **#161**，夹具按此**重述**判据（RT84 `>=1`；RT85 原来那句手算累计 `3+e2` 换成就地取基线问增量）并把"控件已聚焦"显式写出来，之后 BASE/NEW 两份产物**逐字节相同**。邻居复验：`ctrlslider` 逐字相同、`frmevents` 十一条针计数一致（只 `EV05-PAINT`/`EV16-LOSTFOCUS` 先后变）。爆炸半径 `.build/blast_157.py`：119 件 / 33 个文件 / 新增 43 行**全部**含 `vb6_Form_SetInitialFocus`、`UNEXPLAINED=0`（合法型只认"NEW 多出行且每条都是这一句、BASE 一行不少"）。**另外两格顺带记上**：模态里 `VK_TAB` 实测已真跳格 ⇒ #83(b) 只剩"主（非模态）泵没调 IsDialogMessage"；容器那一半也测了 —— 站点序列 `A → B → A → B`，对话框管理器**不走进 Frame**（全仓 `WS_EX_CONTROLPARENT` 命中 0）⇒ 框架里的控件要进 tab 序得先补那一位（照两条创建路的老规矩，两头各一份）。CI 产物里还照出一条本地复现不上的差异（`M2-returned` 的 ticks 本地 1、CI 21，对话框只活了 4 拍 ⇒ 最像"在途积压被模态循环排空"，未证）⇒ **#162**，连带把手册 Timer 那页"最多一枚在途"改成"这个界是负载相关的"。
               # 上一条（2026-09-30T00:40，门 #213 = 账 #156 / C29-T2）= **账 #156 = C29-T2 已出并过门（门 #213，head `893cccf`，attempt 1 即 11 job 全绿，整轮 613s = 10m13s；逐 job bas#1 352s / bas#2 284s / vbp#3 261s / vbp#2 261s / Build 255s / vbp#4 229s / vbp#1 217s / asm 113s / smoke 43s / syntax 26s / compile 19s）⇒ #156 收线**：上一格 s-0 记的"模态窗体的 Timer 一次都不跳"**是症状、不是根因**。RTL 里临时加三行 trace（`C3_OCX_TRACE=1`，量完即回退，另把 C3.exe 重编 + touch 了 `c3rtl.rc` —— RTL 是**嵌在 C3.exe 资源里**的，只改 `src/rtl` 不重编就等于什么都没测，这是本轮第一个坑）直接抓到：`attach owner=MainForm id=100` 与 `attach owner=DlgForm id=100` **同号**，thunk 那行给 `user=100 slot=0 running=0` ⇒ 第二枚窗体的到期被当成第一枚那已停的格吞掉。根因是 `vb6_TimerAttach` 的 id 取自 `g_nextControlId`，而编译器每建一枚窗体都在 CreateControls 开头发一次 `vb6_ResetControlId()`（`cgen_form_create_controls.inc:11`）—— 控件 id 每窗复位，`g_timerTable` 却是进程内一张表、派发只按 id 找（先建那格先命中）。两种后果：前一格停了就一次不投（s-0 看到的），前一格还活着就整格打在**前一枚窗体**的事件过程上（翻倍，s-0 没看见）。修法：独立永不复位的 `g_nextTimerId`（1000 起），`vb6_TimerAttach` 与 legacy `vb6_SetTimer` 两处。判据**双向**：夹具 `tests/c29timer` 加第二枚窗体 `TmForm2`（自带 20ms Timer），主窗体那一枚这一趟设 100ms ⇒ T11 问第二枚跳没跳、T12 问第一枚速率有没有被抢；负控（BASE = 撤掉这一刀重编的 `pre_slt_C3.exe`）`T11=N/0`、`T12=N/61`，修复后 x64 `Y/50`+`Y/10`、x86 `Y/51`+`Y/10`，T1..T10 逐条不变。两枚故意差 5 倍周期 ⇒ 翻红时差的是量级不是抖动。**CI 侧一句实话**：`tmtimer` 与 `tmtimer_x86` 的 exe 同名、日志按 exe 名落盘，vbp #2 那片只有一份 `TmApp.out`（后写的盖掉先写的），这一族拿不到"两份逐字节相同"，x86 读数取本地真跑 —— 想拿就得让 x86 变体的日志带 label（另记一笔，不在本批）。**同批把 s-0 的第 2 条推翻**（探针 `.build/slfoc/FocProbe.vbp`）：模态窗体显示后 `act=fg=root`（确实是活动/前台窗），且在 Timer 拍里 `SetFocus` **真落地**（`eqA=Y`）⇒ "父窗没激活所以 SetFocus 到子窗会失败"是推的、不成立；#157 的范围收窄成"显示时没人把焦点交给第一枚 tabstop（焦点停在窗体那一层）+ `Activate` 里调的 SetFocus 会被随后的激活流程收回"。另开 **#158**：焦点真落到按钮上之后 `_GotFocus` 一次都不发 —— 纯 Win32 探针 `bnprobe.c` 证明 Windows 对 BUTTON 的**程序化** SetFocus 根本不发 `BN_SETFOCUS`（同一段里 `BM_CLICK` 的 `id=42 code=0` 正常到 ⇒ 路是通的；`bnvals.txt`：`BN_CLICKED=0 BN_DBLCLK=5 BN_SETFOCUS=6 BN_KILLFOCUS=7`，我们的码表本身没错），顺带在同一条码表里发现 **ComboBox 用的是 1024/2048 = `CBEM_*`**、真码 `CBN_SETFOCUS=3 / CBN_KILLFOCUS=4` ⇒ Combo 的 GotFocus/LostFocus 今天永远不来；**#159**：`CStr(ctl.hwnd)` 空串、`TypeName(ctl.hwnd)` 给 "Control" ⇒ 句柄被当对象装箱，`GetFocus() = ctl.hwnd` **恒假** —— v4 探针的 `MB0/MB3-focusA=N` 正是踩在这条坏判据上（本线第六次判据自伤；`A=0` 那半独立有效）。
               # 上一条（2026-09-29T22:50，门 #211 = C29-SL-r）= **C29-SL-r（账 #83 的半 (a)）已出并过门（门 #211，head `4debd38`，attempt 1 即 11 job 全绿，job 跨度 587s = 9m47s；逐 job Build 208s / vbp#1 213s / vbp#2 231s / vbp#3 251s / vbp#4 224s / bas#1 358s / bas#2 377s / asm 86s / smoke 41s / syntax 25s / compile 24s）**：用户点头"按建议做"后的第一格 —— 把 #83 那条 `WS_TABSTOP` 旧账按实测重写再修。探针 `.build/sltab/TabProbe.vbp` 读数（029 的 `C29-SL-r-0`）：#83 原话"容器子控件缺这一位"**只说对一半，实际两条创建路都不立** —— 顶层按钮（`.frm` 没写 TabStop）也读回 0、Frame 里的按钮与文本框同样 0，只有显式写了 `TabStop = 0` 那枚碰巧对（它要的就是 0）；`WS_TABSTOP` 这个常量在整个 backend 里以前**一次都没出现过**，而 RTL 的 `vb6_GetTabStop` 一直按"没窗口就是 True"写着 ⇒ 缺的只是创建那一步。做法：新增唯一出口 `controlTabStopStyleBit`（与 `controlTypeStyleBits` / `controlNeedsSubclass` 同处，`.frm` 写了就按写的、没写则除拿不到焦点的 14 类之外一律立），**两条创建路各调一次**；立位**进创建参数**（不是建好再 SetWindowLong）⇒ 两条路各钉一枚 style 数值针。顺带补上同一枚属性的**读侧**（#124 同族）：`TabStop` 归 Boolean 档，以前 `CStr` 答 `-1`/`0`、`TypeName` 答 `Long`、装箱 `VarType` 是 3，现在 True/False、Boolean、11；数据来源仍是窗口 `GWL_STYLE`，没自存 ⇒ 判据不是自洽假绿（#148 那条）。判据 `ST1-tab=True/True/False/False`（顶层 / Frame 里 / Frame 里的 Label / 显式 False 那枚）与 `ST2-bool=11/Boolean`，BASE 逐条 `0/0/0/0`、`3/Long` ⇒ 两条能红；正针三条 + 行为钉一条（显式 False 那枚的 style，BASE 上也在，防"默认立"被改成"一律立"）+ 反向针两条（Label/Frame 的立位版永远不该出现）。**这一格真正的成本在存量针面**：改的是每个控件的创建 style ⇒ 16 条钉着 style 数值的针要跟着走，用脚本 `.build/slr_needles.py` 重算（+65536），条件写死成**"新数出现且旧数消失"才改** ⇒ `ctrlwinsock` 那条自动没动（旧数还留在另一枚控件上、串仍命中，于是它已不再专门证明 Winsock 那枚的样式 —— 这条松在哪里写明，不装看不见）。爆炸半径 `.build/blast_slr.py`：`changed-files=32`、`changed-lines=550`、`UNEXPLAINED-LINES=0`，并且**拆开算清怎么对上的**：273 行是逐行严格按"BASE 行 style 数字 +65536 之后与 NEW 逐字相同（多重集相等）"、4 行是 `GetTabStop` 读侧形状且全在本批自有夹具里（非本夹具 0 行）、其余是夹具新增行 ⇒ "550 行全绿"这种含糊话没写。四份 `CtrlState.out`（CI x64/x86 + 本地两份）逐字节相同（250 字节）；`ctrlslider` 夹具输出与上一扇门 #209 的工件逐字节相同 ⇒ 没动任何已有读数。**门号跳格一笔**：`#210` 是别人 `workflow_dispatch` 的一轮（head `9f36ae2`，失败），不是本线推送 ⇒ **认 head，别拿门号连续当判据**。**留一句没兑现的话（老实记）**：立位之后"模态窗体里 Tab 该能用了"只是从 `vb6forms.c:1012` 那条循环**推**出来的，本批**没实测**；做半 (b) 时第一件事就是先测它（模态窗体 + 两枚按钮 + 送一个 `VK_TAB`，问焦点归谁），否则 (b) 的 A/B 无从对照 —— 半 (b)（主泵走 `IsDialogMessage`）在任务 #155 名下待做。｜上一轮: **C29-SL-q 已出并过门（门 #209，head `a87548f`，attempt 1 即 11 job 全绿，job 跨度 573s = 9m33s；逐 job Build 207s / vbp#1 253s / vbp#2 210s / vbp#3 282s / vbp#4 229s / bas#1 251s / bas#2 364s / asm 91s / smoke 88s / syntax 28s / compile 21s；推送重试到第三次才过 —— GitHub 连吃 502/504，与门结果无关）**：账 #154 结 —— `.frm` 里写的 `FontName` / `FontSize` **整条没打到窗口**（探针 `.build/slfont`：写 Consolas+14 的文本框读回 空串/8.25，写 20 的读回 8.25 且窗口像素高度还是默认 11）。解析侧一直留着这两条（`frm_parser.cpp` 原样存进 `ctrl.properties`），缺的只是发码 ⇒ 与 #125 / #142 同一形状：**缺口只在创建那一趟，运行期赋值一直是好的**。新增 `emitDesignerFontProps`，**两条创建路各调一次**（窗体上那枚 `txtH`、Frame 里那枚 `txtI` 各钉一条 —— 这一族栽过几次「只接一头」），只在 `.frm` 显式写过时发，排除 `Timer`/`Menu`/`ImageList`（没窗口）与 `CommonDialog`（它的 Font* 是 ChooseFont 的字段，D6 口径，再发就是两遍）；**发序先 Size 后 Name**（Name 那支在窗口没有字体时会先造 `-13` 的底子，反过来会盖掉字号）。RTL 加 `vb6_SetControlFontNameW`：设计期的字体名在生成码里是 **C 字面量不是 BSTR**，喂给上面那支会让 `SysStringLen` 去量串尾之后的内存；没改用 `vb6_BSTR_FromStr` 现造临时串 —— 本仓刚为「发码里多一枚没人放的临时串」开过 #119。**本批真正的收获是一条顺带查出来的旧坏**：`FontName` 从没登记成 **String 档**，于是同一枚属性两种答案 —— `Len(ctl.FontName)` 直接拿指针、答得对（Consolas 读出 8、MS Sans Serif 读出 13），而 `Print ... & ctl.FontName` 先装箱再 CStr、打出来是**空串**（正是 line 62 那段注释给 `Tag` 记过的同一证状）。它一直没被看见是因为判据常用打印面，而打印面看着像「字体没设上」——**第一版读数就是把这条误读成设计期没生效的**，靠同批的 `FontPixelHeight` 证人当场翻案才分成两头量（写侧确实没发、读侧发了也读不出），两头各自独立钉，才没把两件事当一件事收掉。判据 `SQ1-dt=Arial/20`、`SQ2-child-dt=20`、`SQ3-real=Y` 三条**全能红**（BASE 逐条 `/8.25`、`8.25`、`N`；这批复用红针要留意：本批 BASE 是 SL-p 之后的那一枚、证人已存在，所以 `SQ3` 的红是「窗口真没换字体」，不再是上一格那种「证人还不存在」—— **同一条针在不同 BASE 上红因可以不同**）。发码正针四条（含容器那一条 Size）、反向针一条（装箱形状 NEW 归零）。x64 与 x86 两份 run.out 逐字节相同（106 行 / CI 侧 1635 字节，四份彼此相同）、`.err` 各 0 字节、token 数 `ERRORS=0`；爆炸半径 118 件（BASE `pre_slq_C3.exe`）`changed-files=1`、5 行、`UNEXPLAINED-LINES=0` ⇒ 存量里既没有第二处设计期字体、也没有第二处读控件 `FontName`。**留下另立账（未做）**：usercontrol 实例（`Begin czFormDemo.czControl … FontSize = 11`，`czUI-main` 7 处）的设计期属性整条不落发码 —— 那种实例压根没有窗口，属「uc 实例设计期属性」另一族，与本格无关，刻意不碰（同时改两族红了分不清）。⇒ Slider 这页接下来 = #83 剩下的族（`vb6_DTP_Init` / `vb6_TreeView_Init` / MonthView 平铺搬进共用出口）或容器子控件缺 `WS_TABSTOP`。｜上一轮: **C29-SL-p 已出并过门（门 #208，head `1f7e634`，attempt 1 即 11 job 全绿，job 跨度 520s = 8m40s；逐 job Build 128s / vbp#1 259s / vbp#2 236s / vbp#3 234s / vbp#4 251s / bas#1 348s / bas#2 390s / asm 164s / smoke 36s / syntax 23s / compile 21s）**：这一格是**照用户点名的那页文档做事**的第一格 —— 把 `ai/内置控件/Slider 控件（滑杆）.md` §4 那个「实时缩放字号」的例子整条搬进探针真跑（`.build/slfont/FontProbe.vbp`），例子本身是通的（`With` + 冒号续行、`Change` 联动、`Caption` 串数值都对），但量出一条**通用控件属性**的坏：`Text1.FontSize = 8` 读回 `8.25`、写 `10` 读回 `9.75`、写 `14` 读回 `14.25`。因由不是换算写反，是**请求值从没存过**：字体在窗口上只能按整数像素存在（96 DPI 下 1pt = 1.3333px），而 getter 一直从那枚 LOGFONT 的像素高度**反算**点号 ⇒ 往返必然落格。修法按本仓既有口径「窗口表示不了的属性自存」（`TickFrequency` / `TextPosition` 同族）：写侧把点号存进窗口属性、读侧有自存才答自存，**没设过字体的控件仍走原来那支反算**（夹具里那枚未设过的 txtG 两边都读 `8.25`；邻居 `ctrldlg` 那处 CommonDialog 字号走独立支路，BASE/NEW 两份 run.out `fc /b` 逐字节相同）。**Set 旗标是这条的命门**：`SetPropW(hwnd, 名, 0)` 等于删属性（账 #107），而 `0.0f` 的位正好是 0 —— 少那一枚「写 0」会静默退回「没设过」；先例照 `VB6_BackColor` / `VB6_BackColorSet` 那一对。**判据两头**（#148 那条教训的第二次应用）：光读自存的数就是自洽假绿，所以新加一枚 C3 扩展证人 `vb6_ControlFontPixelHeight`（问窗口真在用的像素高度，只给读侧、登记在通用那一组并**排除 CommonDialog**）。四条读数 `SP1-round=10` / `SP1b-round14=14/8.25` / `SP2-real=Y/26` / `SP3-zero=0`，BASE 逐条 `9.75` / `14.25/8.25` / `N/26.25` / `0`；其中 `SP3` 老实标成**行为钉**（BASE 也读 0：0pt 折成 `lfHeight=0`、反算也是 0），`SP2` 前半在 BASE 的红因是**证人还不存在**（那一次落 `vb6_ComGetIntProp(裸 HWND, L"FontPixelHeight")`）而不是「字体没换上」—— 两条都写进台账不当捷径用。反向针一条能红（那条形 NEW 归零）。针面 `SP1b-round14=14/` 刻意写成**前缀**：后半那枚是 DPI 的函数（CI 与本机同为 96 DPI、都读 `8.25`，但钉死等于把本机写进判据）。x64 与 x86 两份 run.out 逐字节相同（103 行 / CI 侧 1589 字节，四份彼此相同）、`.err` 各 0 字节、token 数 `ERRORS=0`；爆炸半径 118 件（BASE `pre_slp_C3.exe`）`changed-files=1`、5 行、`UNEXPLAINED-LINES=0`。**顺带开的新账（#154，未修）**：普通控件的 `.frm` **设计期** `FontName` / `FontSize` 从来没打到窗口上 —— 发码面只有 CommonDialog 那一条分支应用这两个，与 #125 / #142 同一形状（缺口只在创建那一趟）。⇒ Slider 这页接下来 = 要么收 #154（设计期字体下发，两条创建路都要接），要么按 #83 剩下的族搬共用出口（`vb6_DTP_Init` / `vb6_TreeView_Init` / MonthView）。｜上一轮: **C29-SL-o 已出并过门（门 #207，head `81e845c`，attempt 1 即 11 job 全绿，job 跨度 595s = 9m55s；逐 job Build 213s / vbp#1 245s / vbp#2 233s / vbp#3 267s / vbp#4 231s / bas#1 351s / bas#2 379s / asm 89s / smoke 45s / syntax 26s / compile 17s）**：账 #83 的 **Slider 半边**结 —— 容器（Frame / PictureBox / SSTab）里那枚滑杆以前**压根不发** `vb6_Slider_Init`，`.frm` 写的 Min/Max/Value/Small·LargeChange/TickFrequency/Sel* 整条丢掉，控件停在未收到量程消息那一档。做法只有一刀：把顶层那趟里那段内联（原 `cgen_form_ctrl_style_apply.inc:574-611`，38 行）搬成共用出口 `CCodeGen::emitSliderDesignTimeInit`，两条创建路各调一次 —— 参数序、`-999` 哨兵、`SelStart`+`SelLength` ⇒ 终点那一折全只有一份（那段注释记的是「哪一端先动会把后面的吃掉」，写两遍迟早写歪）；`hwndExpr` 由调用方给 ⇒ 控件数组那条槽也自动跟着走。**#83 当初留的那句「Max 看着到了，别当覆盖」查实是巧合**：BASE 的发码面里 `sld7` 只有 `vb6_CreateControl` 一行、一条 range 都没发，而**同一件产物**里顶层那枚什么都没写的 `sld5` 读出 `SB9-defaults=0/100/1/20/0/0` —— 未收量程消息的轨道条在产物里就答 0..100（裸探针同一档是 10，「探针读数不给产品定罪」第三次，与 C29-5a/V6、SL-j 同族）；新读数里 `Min` 0→10 才是真证人，因为高低两端打包在**同一条** `TBM_SETRANGE` 的 lParam 里。判据形状 = **两条路读同一串数**（不是「看着非零就算过」）：夹具 `sld7` 补齐与顶层 `sld4` 同一批属性 + `SO1-child=10/100/42/2`、`SO2-child-page=8/5`、`SO3-child-sel=20/15/Y`（BASE 同夹具读出 `0/100/0/1`、`20/0`、`0/0/N`），发码正针一条钉住折出来的 `35L`；这格**没有可禁的反面形状**（BASE 压根不发），反向针空着、红点全在正向针与三条真跑上，另加一条假针实测 False 证明核对不空转。x64 与 x86 两份 `run.out` 逐字节相同（99 行 / CI 侧 1526 字节，四份彼此相同）、`.err` 各 0 字节；`run_tests.ps1` token 数 `ERRORS=0`。爆炸半径 118 件（BASE `pre_sln_C3.exe`）：`changed-files=1`、`changed-lines=1`、`UNEXPLAINED-LINES=0` ⇒ 存量里没有第二枚容器子滑杆。**#83 仍开着**：同一形状还欠 `vb6_DTP_Init`（CustomFormat）、`vb6_TreeView_Init`、MonthView 多月平铺那几段只有顶层发得出的，出口已就位、一次一族搬过去并各补一枚同形判据（搬完没判据就等于「编得过、看不见」）；另有一条独立的 —— 容器子控件缺 `WS_TABSTOP`（Tab 顺序那一档）。｜上一轮: **C29-SL-n 已出并过门（门 #206，head `36d4fba`，attempt 1 即 11 job 全绿，wall 539s = 8m59s；逐 job Build 208s / vbp#1 234s / vbp#2 254s / vbp#3 221s / vbp#4 227s / bas#1 324s / bas#2 322s / asm 80s / smoke 45s / syntax 23s / compile 27s）**：账 #141 结 —— 控件级 `_DblClick` 这一形整条修好。量法是把探针做成"**每枚控件只挂一条处理器**"（`.build/slwith/MouseProbe.vbp`，五枚滑杆各挂一条）—— 这样"装没装子类"完全由那一条形决定，一眼看得出哪一形没进判据。结果 #141 的前提只说对一半：`MouseDown/MouseMove/KeyPress` 一直是通的（三处表都有它们），**只挂 `_DblClick` 的那枚从没被 install**（子类过程与 `WM_LBUTTONDBLCLK` 那条 arm 都生成了 ⇒ 处理器编得过、永不触发）。同一形还有第二处、症状更硬：**arm 把 `_DblClick` 写死成无参调用**，而 VB6 签名是 `Sub X_DblClick(Cancel As Integer)` ⇒ 按标准签名写直接 `error C2198 用于调用的参数太多`（BASE 实测 BUILD-RC=1）。本仓夹具躲过它是因为 SL-d 当年写的是无参那一形 ⇒ **"编得过"不等于"签名对"**。做法两刀、都不新开表：① 唯一出口 `controlNeedsSubclass(ctrl)`（与 `controlZeroArgMethod` 同处），**四趟一起改**（发 arm 的顶层与容器子控件两处、装的一趟、拆的一趟），焦点与 Click 那两条排除表跟着走；② `_DblClick` 的 arm **按处理器自己声明的形参数**发（带 Cancel ⇒ `fn(&vb6_dblcancel)`；无参 ⇒ 仍 `fn()`），两形各钉一条形状针。判据 `SN7-dblone=1/0`（次数 / 传进来的 Cancel），x64 与 x86 逐字节相同（1460 字节）；四正针里三条 BASE 全没有。**这一批的负控同样是"编不过"**：BASE 编当前夹具 BUILD-RC=1（C2198 一处 + C2039 两处，后者是 With 那一形、门 #205 已修）。存量效果：`ctrlfiles` 的 `fileList`（只挂 `_DblClick`）现在也被装/拆了 —— 正是 SL-d 那条注释里"子类过程定义得好好的、一次都没被 install" 的同一枚控件、另一条形。爆炸半径（BASE 取 SL-l 之前那一枚，一趟兜 l/m/n 三批）：`changed-files=7`、266 行、`UNEXPLAINED-LINES=0`。`_Paint` 仍按 Fix 185 只给 PictureBox；With 里的带实参方法（`.SimNotify(5,40)`）留在 #150 名下。⇒ 下一格 = **C29-SL-o**：容器子滑杆不发 `vb6_Slider_Init`（探针实测顶层 `10/100/42/2/9/5` vs Frame 内 `0/100/0/1/20/0`），即账 #83 剩下的那一半，读数已记进 029 的 `C29-SL-o-0`。｜上一轮: **C29-SL-m 已出并过门（门 #205，head `69c79b7`，attempt 1 即 11 job 全绿，wall 542s = 9m02s；逐 job Build 177s / vbp#1 246s / vbp#2 180s / vbp#3 243s / vbp#4 228s / bas#1 352s / bas#2 350s / asm 69s / smoke 48s / syntax 22s / compile 26s）**：账 #150 结 —— **With 块里的控件方法**。这一格是被用户一句"嵌套 With 你修过了吗"问出来的：查历史查到 **Fix 092o**（`d6a4dcf`，2026-09-10，"With 目标表达式在外层上下文中生成 —— 延后入栈"，起因 `cAliyunCaptcha` 222 行嵌套 `With .ReturnJson()` 报 C2039）+ 继承那批的 **B08e-1/B08e-2**（With 内类成员走虚表派发），回归覆盖在 `form_test_p17_with.frm` 的 `P17.1: Nested With`；**本线此前一格都没碰过 With**。顺着问出去量了一遍，挖出真洞：**`With` 的成员访问只认属性，方法名落到"未知属性"兜底**（`cgen_expr_with.cpp:52`），发成 `_vb6_with_0.SetFocus()` —— HWND 是 struct 指针 ⇒ **整件工程编译不过**（`BUILD-RC=1` + 三条 VB4001），而修复前的编译器逐字相同 ⇒ 预存缺陷、非回归。 修法沿用 ClassInstance 那条 **Fix 090s 协议**：`asCallCallee_` 为真就交裸函数名 + `pendingChainObj_`（接收者 = With 入口那枚 HWND）由调用点补实参，为假才交完整调用；表还是 SL-l 那张 `controlZeroArgMethod`，不再开第二份名字表。**第一版没走协议 ⇒ 带括号那形被调用点又补一对括号，发成 `f(hwnd)()`（非法 C）** ⇒ 两形各钉一条判据才拦得住这一族。判据 SN4/SN5/SN6（With 内不带括号 / With 内带括号 / 方法与属性赋值混在同一个 With 块）；**这一批的负控不是"读数不对"而是"编不过"** —— 修复前的编译器编当前夹具 `BUILD-RC=1`。判据自伤同族第二次：`SN4` 第一版读 N，因为那一步 `sld6` **已经有焦点**，而 `SetFocus` 落在已有焦点的窗口上不重发 `WM_SETFOCUS` ⇒ 问增量之前要先把焦点挪开（与 SL-l 那条"落点被禁用"是同一课的两个方向：一个问能不能拿到、一个问会不会重发）。**刻意没接的两形**留在账上：`WithEventsCtrl` 那一支有同一形状的坏兜底、With 里的**带实参**控件方法（`.SimNotify(5, 40)`，判据助手，写全名即可）。爆炸半径（BASE 仍取 SL-l 之前那一枚，一趟兜两批）：`changed-files=6`、260 行、`UNEXPLAINED-LINES=0`（对照 SL-l 单批 252 行 ⇒ 净增量就是 With 那几条）。CI 工件四份逐字节相同（1444 字节）。⇒ Slider 这页只剩 #141（`_DblClick`/`_Paint` + `KeyPress`/`Mouse*` 四条事件）与 `Min`/`Max` 超 16 位那一档。｜上一轮: **C29-SL-l 已出并过门（门 #204，head `441b6f6`，attempt 1 即 11 job 全绿，wall 524s = 8m44s；逐 job Build 144s / vbp#1 278s / vbp#2 212s / vbp#3 239s / vbp#4 239s / bas#1 375s / bas#2 279s / asm 87s / smoke 44s / syntax 26s / compile 24s）**：账 #143 结 —— 控件的**零实参方法两形同归**。不带括号的 `Slider1.ClearSel` / `Text1.SetFocus` 以前掉 `vb6_ComCall(裸 HWND, L"…", NULL, 0)`；现在收成一张表 `controlZeroArgMethod`，**两条码头共用**（带括号在 `cgen_expr_call_com_bind.inc` 的 marker 支、不带括号在 `cgen_call.cpp` 的语句路），宿主认三种形态：`vb6_hwnd_X`、裸小写名、控件数组的 `vb6_CtrlArr_GetAt(&vb6_hwnd_X, i)`。RTL 新增 `vb6_SetControlFocus`（就一句 SetFocus(hwnd)）。**这条账的前提被翻了两次，最后一版以读数为准**：原来说"被吞成取属性"⇒ 发码级订正成"落到 COM 兜底"；本批真跑又订正一次 —— **兜底那条形在运行期是响的**，因为 RTL 有一张宿主对象应答表（Fix 112，`uc_hostmodel_call.inc:69` 按名字应答 `SetFocus(obj)`），`vb6_ComCall` 在解 vtbl 之前先 `IsWindow` 认宿主 ⇒ BASE 与 NEW 的焦点读数**逐字相同**。**通用式：发码落到 `vb6_ComCall` ≠ 运行期静默空转**，下这个结论前先查那张表里有没有这个名字。真正一声不响的是表里没有的名字 —— `ClearSel` 就是（夹具 BASE `SN3-clearsel=10/20` → NEW `0/0`）。所以本批**能红的**是 SN3 + 发码正反针（`VBFlexGridDemo` 里 31 处 ComCall 形状 → 1 处，剩那处是 `Me.SetFocus` 走窗体自己那条路），而 `SN1`/`SN2` 老实标成**行为钉**（红不了，防的是焦点面被人改坏）。**判据自伤一条（本线第四次同一形）**：第一版拿 `sld1` 当"把焦点挪开"的落点，而 `sld1` 在前一步 SL11 被 `Enabled = False` 了 —— 禁用窗口拿不到焦点（Win32 语义），焦点从没离开过 sld6，SN2 两条增量双双读 N；换成一枚启用又没挂处理器的（`sld5`）才对。**写焦点类判据前先问那枚落点现在能不能拿焦点**。接谁不接谁写死在表里：Label/Image/Shape/Line/Timer/Menu/Data/OLE/CommonDialog/Winsock/ImageList/StatusBar/ProgressBar/Form 不接（真 VB6 在那里 raise 错误号，本项目还没有运行期错误面，"什么都不做"比伪造成功诚实）。爆炸半径 118 件：`changed-files=6`、252 行、`UNEXPLAINED-LINES=0`。另外把上一格欠的那句"未真跑"收回了：`With` 块探针真编真跑（`W1=12/8/72`…`W5=VOL/String`）⇒ 说明 §4 那个例子运行期也通。CI 工件四份逐字节相同（1384 字节）。⇒ Slider 这页只剩 #141 的 `_DblClick`/`_Paint` 与 `KeyPress`/`Mouse*` 四条事件、`Min`/`Max` 超 16 位。｜上一轮: **C29-SL-k 已出并过门（门 #203，head `0c99d8f`，attempt 1 即 11 job 全绿，wall 525s = 8m45s；逐 job Build 166s / vbp#1 250s / vbp#2 218s / vbp#3 258s / vbp#4 252s / bas#1 351s / bas#2 354s / asm 80s / smoke 41s / syntax 22s / compile 17s）**：账 #147 结 —— Slider 的 `Text` 气泡 + `TextPosition`。**先订正一条名字带来的误判**：类型库（`LoadTypeLibEx` + `REGKIND_NONE`）里 `Text` = dispid `0x0010`、`VT_BSTR`，文档原话 "the string displayed in the ToolTip as the slider s position changes" ⇒ 它是**气泡里那句串，不是窗口标题**，这一格把它从通用 `vb6_GetControlText`（`GetWindowText`）那支抢过来自己实现。路线是探针定的（`.build/slprobe/slmeasure17.c`，同一份源码编两份、一份用 `mt.exe` 挂 Common-Controls 6.0 清单）：轨道条自带的 `TBS_TOOLTIPS` **服务不了自定义串**（它那条工具的文本是它自己画的数字、`TTM_POP` 无头里一条都不发），而且 `TBS_TOOLTIPS` **只有创建时给才建得出气泡**（事后写样式位留着但 `TBM_GETTOOLTIPS` 恒 0，与 DTPicker 的 `DTS_SHOWNONE` 同族），而运行期 RTL 只拿得到 HWND、补不回来 ⇒ 气泡改由 RTL 自持一枚 `TTS_ALWAYSTIP` 宿主 + `TTF_TRACK|TTF_ABSOLUTE` 工具来摆，派发那条 `vb6_Slider_BubbleNotify` **只发一次、按窗口类名筛掉 ScrollBar**（4/5 摆、8 收，与 `Scroll` 同一分法；容器子控件那一路仍在账 #83）。三条判据证人一律问**宿主**（`IsWindowVisible` / `GetWindowRect.top` / `TTM_GETTEXTW`），夹具 SK 七条 **x64 与 x86 产物逐字节相同**，BASE 上六条里五条红（`SK1` 老实记着只当证人 —— 未登记时 `Text` 走的是「拿 HWND 当 IDispatch 问它要属性」那支，串照旧回得来）。两条坑：**`TTM_GETTEXT` 回 0 而文本照样复制进缓冲** ⇒ 只看缓冲（与上一格 #148 同一处坑、同一个修法）；**v5 探针里 `TTM_ADDTOOLW` 直接返回 0**，差点据此把气泡判成做不出来 ⇒ 判据一律写在产物里。`TextPosition` 原生没有对应消息 ⇒ 自存 0/1、摆的时候用它定高低，判据只比相对不钉像素（`TickStyle` 那条教训）。爆炸半径 118 件一趟：`changed-files=1`、61 行、`UNEXPLAINED-LINES=0`。CI 工件对照：两份 `SlidApp.out` 各 1337 字节、彼此相同且与本地两份相同。⇒ Slider 这页剩下 #143（不带括号的控件方法形）、#141 剩下的 `_DblClick`/`_Paint` 与类型库读出来的 `KeyPress`/`Mouse*` 四条事件。｜上一轮: **C29-SL-j 已出并过门（门 #202，head `f7ec817`，attempt 1 即 11 job 全绿，wall 547s = 9m07s）**：这一格不添新面，是把**已发货东西里没人验过的那一步**补上证人 —— 裸编探针（无 v6 manifest ⇒ comctl v5）里 `TTM_ADDTOOLW` 返回失败，照那个读数会判 `ToolTipText` 的气泡从没注册上；**产物真跑 `SJ1-reg=Y/Y/N` 推翻了这条定罪**（版本差异不是缺陷），证人 `vb6_ToolTipRegistered` 留在判据里当防回归，第三格（从没设过的 sld1 = N）是刻意留的负控。顺带把 #147 要用的四条原生读数钉进账与任务里，其中**订正了一条我自己读错的**：`code=-12` 是 `NM_CUSTOMDRAW` 不是 `TTN_GETDISPINFO`（`TTN_FIRST=-520`），所以"气泡文本走 dispinfo"至今无实证。｜上一轮: **C29-SL-i 已出并过门（门 #201，head `133175c`，attempt 1 即 11 job 全绿，wall 494s = 8m14s；逐 job Build 211s / vbp#1 239s / vbp#2 216s / vbp#3 238s / vbp#4 251s / bas#1 269s / bas#2 278s / asm 87s / smoke 51s / syntax 33s / compile 17s）**：账 #146 结 —— 类型库读出来的 VB6 选区面补齐（`SelStart` + **`SelLength`** + 方法 **`ClearSel`**，VB6 那一面压根没有 `SelEnd` 这个名字，那条是 SL-b 按原生 `TBM_SETSELEND` 自己加的）；顺带把设计期那一对折进创建参数（以前 `.frm` 写了 `SelLength` 也整条丢掉，与账 #142 同形状）。两条口径值得记：空选区有两种、起点不一样（没碰过答量程下限、CLEARSEL 之后答 -1 折 0），所以 `SetSelLength` 在空态锚下限，免得 `ClearSel()` 之后那句写长度变成静默 no-op；没挂 `SelectRange` 就是写不进去，不伪造。CI 工件对照已同日补做：`test-logs-vbp-3`(x64)、`-4`(x86) 与本地真跑三份 `SlidApp.out` **逐字节相同**（81 行 / 1213 字节，SI 十条齐、两份 `.err` 各 0 字节）；第一次 `gh run download` 被 TLS 握手超时挡了一轮，重试才拿回来。｜上一轮: **C29-SL-h 已出并过门（门 #200，head `cf7ab6b`，attempt 1 即 11 job 全绿，wall 475s = 7m55s；逐 job Build 194s / vbp#1 200s / vbp#2 213s / vbp#3 240s / vbp#4 269s / bas#1 272s / bas#2 270s / asm 77s / smoke 40s / syntax 22s / compile 17s）**：TickStyle 四档不必等「VB6 真值」—— `LoadTypeLibEx(路径, REGKIND_NONE)` 直接读出 OCX 自带的类型库（枚举成员名与数值、属性 VARTYPE、dispid、事件签名全在里面），**这条门对本线其余所有「等 VB6 真值」的格子都开着**；能读的只有声明面，行为面（越界怎么办、事件时序）仍然读不到。顺带补掉第二条创建路对 Slider 一格样式都不挂那个洞（#83 的一半，另一半是容器子控件不发 `vb6_Slider_Init`，仍开着）；新开 #146（`SelLength`+`ClearSel`，VB6 的选区其实是 SelStart+SelLength）与 #147（`Text` 气泡+`TextPosition`）。｜上一轮: **C29-SL-g 已出并过门（门 #197，head `3c7d247`，attempt 1 即 11 job 全绿，wall 571s = 9m31s；逐 job Build 213s / vbp#1 283s / vbp#2 244s / vbp#3 255s / vbp#4 216s / bas#1 352s / bas#2 282s / asm 82s / smoke 43s / syntax 56s / compile 20s）⇒ 账 #141 的一半结**。这一格缺的不是 Slider 一条 arm，是一整片：`_GotFocus`/`_LostFocus` 那一句判据在仓里被抄成**三份表**（发 arm 的 subclass、装的 frame_menu、拆的 dispatch），三份写着同一份**四类白名单**（PictureBox/Frame/Label/Image），而另一批控件的焦点通知走父窗 `WM_COMMAND` 码（EN_=256/512、CBN_=1024/2048、LBN_=4/5、BN_=6/7）⇒ 两条路都没覆盖到的那一串（Slider、ListView、TreeView、DTPicker、MonthView、RichTextBox、两个滚动条、文件系统三件套）的焦点处理器是**编得过、永不触发**的死代码。改成 `controlFocusFromNativeNotify`（只排除已有 WM_COMMAND 来源的那六类，与 `controlClickFromNativeNotify` 同形状同处定义），**三份表一起改**。顺带补上 SL-d 漏在"拆"那份里的 `_Click` —— 后果不崩、只是少拆一枚，代价是两件存量工程（`Charts 2020/ucProgressCircular`、`czUI-main/czFormDemo`）各多出一条 `vb6_RemoveControlSubclass`，那是补齐对称不是行为回归。判据不伪造焦点消息：`vb6_Slider_SimStdEvent` 加一档 `kind=4`，它**不发消息**而是真 `SetFocus(控件)` —— 焦点那两条的原生来源就是窗口管理器自己发的，伪造就验不到"一次移动两边各发一次"；走 SL-d 那根已接好的支路，本批零新增方法接线。判据 `SG1-got`/`SG2-move`/`SG3-back`/`SG4-isolate` 四条在 BASE（HEAD 的干净 worktree 冷编 `.build/pre_slg_C3.exe`）逐条为 **N**，发码面钉 6 条（三档 arm + 那条 install + 那条 Remove + 那一档 SimStdEvent）；`sld3` 刻意只挂 GotFocus 一条处理器 ⇒ 它是"只靠焦点处理器也该被装/被拆"的证人。**一条老实记着不能红**：`SimStdEvent(…,4,0)` 的形状 BASE 也有（那里 kind=4 落到 default 当 KEYUP 用）—— 语义变了形状没变，它只当"调用点没退化"的证人。Label 留在"会发 arm"的集合里（STATIC 拿不到焦点 ⇒ arm 永不触发，与 VB6"Label 没有焦点事件"一致），不为它单开一张表。A/B 与爆炸半径都用**本批自己的 BASE**：118 件 `--emit-c` ⇒ `changed_files=3`、18 行、`UNEXPLAINED-LINES=0`；另跑全仓扫描 —— 带 `WM_SETFOCUS`/`WM_KILLFOCUS` arm 的工程只有 `ctrlslider`，且没有任何控件同时挂着子类 arm 与 `WM_COMMAND` 那条来源（排除表成立）。CI 收线：两份 `SlidApp.out`（x64/x86）逐字节相同、且与本地真跑那份相同（61 行），SG 四条齐，`.err` 各 0 字节。**Slider 这一页到此只剩等 VB6 真值的两条**（`TickStyle` 四档、`Min`/`Max` 超 16 位）；本仓另开 **账 #143**（`<控件>.SetFocus` 被吞成取属性，实测 `NewTab-test` 产物里是一条 `vb6_ComCall(裸 HWND, L"SetFocus")`，永不调用也无诊断），#141 剩下 `_DblClick`/`_Paint` 两半。
               # 上一条 = 2026-09-29T09:20:00+08:00   # 本轮追加：**C29-SL-f 已出并过门（门 #196，head `652a14b`，attempt 1 即 11 job 全绿，wall 596s = 9m56s；逐 job Build 208s / vbp#1 260s / vbp#2 236s / vbp#3 201s / vbp#4 170s / bas#1 346s / bas#2 379s / asm 101s / smoke 45s / syntax 26s / compile 22s）⇒ 账 #142 结**。这一格是上一格量出来的：夹具里 `.frm` 明明写了 `ToolTipText = "dtip"`，运行期读回来是空串 —— 属性行**解析到了、没人发射**（设计期那一趟只管样式/颜色/字体与 #125 那三条整数面，字符串属性不归它管）。补 `emitDesignerStringProps`，与 #125 同处、同样被**两条创建路**共用（顶层 `cgen_form_ctrl_style_apply.inc`、容器子控件 `cgen_form_frame_menu.inc` 各加一行）。四条口径都有量或推演支撑：① **只在写了非空值时才发**（空串与没写过读数同为 `""`，而那一步除了存窗口属性还要往共享 tooltip 控件注册一条工具 ⇒ 多发一条 = 改产物做语义上零的事；与 #125"反面才发"同纪律。夹具给 `sld1` 挂一条**显式** `Tag = ""` + 一条反向发码针钉住）；② **不转小写**（第一版顺手 `Symbol::toLower` ⇒ 量出来 `L"DTIP"`，读数当场翻红）；③ **只认 `FrmValueType::String`**（真实 VB6 会把资源引用写进这两条属性，`ToolTipText =   "frmTest.frx":0000` 解析成 `FrxReference` ⇒ 跳过；顺带挡住以后补 `.frx` 解析时把引用当文本灌进 `SysAllocString`）；④ 字面量转义反斜杠先走、双引号后走。判据：`ctrlslider` 加 SE10-dt=dtip / SE12-dttag=dtag / SE13-over=rt/ / SE14-emptytag=/end，`ctrlprop` 加 CP16=dtagL/ltt / CP17=dtxt/xtt / CP18=/2 —— 登记在通用段就要用非 Slider 的控件证（Label 挂设计期 Tag、TextBox 挂设计期 ToolTipText，CP18 顺带钉"清 ToolTipText 不牵连 Text"）；发码面 2 正向 + 1 反向。红证说实话：`SE10`/`SE12`/`CP16`/`CP17` 在 BASE 上逐条为空，`SE13`/`SE14`/`CP18` 两版同值、只当证人（它们的价值在配套的反向针与"双向都能动"这一面）。A/B（两份同为本地 Debug）：两件夹具产物与 BASE 的 `diff` 只落在 SE / CP 那几行，其余读数逐字节不动，`.err` 各 0 字节。**爆炸半径这次是直接测出来的**：118 件 `--emit-c` 全扫，带 `/* design ToolTipText|Tag */` 这行的工程**只有我自己那两件夹具**（各 2 行）⇒ 存量一枚不受影响（全仓没有第二处非空的设计期字符串属性写在标准控件上）；按行型分类护栏 `UNEXPLAINED-LINES=0`。注册仍然只往既有块追加串、零新增 `Test-*` 调用（分片归属不动，门 #191 那条波澜站稳之后连续四批一次过）。CI 收线：`test-logs-vbp-3`(x64) 与 `-4`(x86) 两份 `SlidApp.out` 逐字节相同、且与本地那份逐字节相同，`CtrlProp.out`(vbp#1) 三条齐。**Slider 这页到此只剩等口径的**：`TickStyle` 四档、`Min`/`Max` 超 16 位（都等 VB6 真值），另有 `GotFocus`/`LostFocus` 一档、#141（`_DblClick`/`_Paint` 没进"装不装"那趟）、#83（`WS_TABSTOP` + 主循环缺 `IsDialogMessage`）。
               # 上一条 = 2026-09-29T08:40:00+08:00   # 本轮追加：**C29-SL-e 已出并过门（门 #195，head `071c713`，attempt 1 那条红是与本批无关的旧脆弱 `frmevents`、attempt 2（`rerun --failed`）11 job 全绿，逐 job：Build 208s / vbp#1 248s / vbp#2 202s / vbp#3 215s / vbp#4 244s / bas#1 350s / bas#2 371s / asm 85s / smoke 41s / syntax 30s / compile 26s）**。这一格坏的**不是读写面**（`vb6_Get/SetToolTipText`、`vb6_Get/SetControlTag` 早就对所有可见控件登记着），而是**档位**：类型表里没有这两条，而两条 getter 的 C 返回型写 `void*`，装箱那一步是 C11 `_Generic`（`vb6rtl_variant.h` 的注释表自己写着「其他指针 (void*/class*/type*) → VariantObject」）⇒ 一条字符串被当**对象**装。实测（探针工程 `.build/sltt`，一枚 Slider + 一枚 Label、逐条表达式各一行）：`CStr(Slider1.ToolTipText)` 打**空**、`TypeName` 答 **"Object"**、`VarType` 答 **9**（VT_DISPATCH！）、`If Slider1.Tag = "tag1"` 答**假**；而同一枚属性的 `Len`(=3) / `InStr`(=1) / 赋值(="abc") 这几条**直接拿指针**的面是对的 —— 同一属性两种答案，就是"没登记"最典型的形状（不是崩、是半对半错）。还量到一条分叉：比较面上 `ToolTipText` 走 `vb6_StrCmp(getter, BSTR)`（对），`Tag` 走 `vb6_StrCmp(vb6_CStr(vb6_VariantFromValue(getter)), BSTR)`（拿 "" 去比）⇒ 分叉点在"按名字查档"那一层，不在发码处。修法两处：类型表**通用段**（与 #124 那两条同一段）加 `tooltiptext` / `tag` ⇒ String（登记之后 `CStr` 那一步被折掉，读侧就是裸 getter 调用，与 `Text` / `SelText` 同形）+ RTL 两条 getter 返回型 `void*` → **`wchar_t*`**（与已经对的 `vb6_GetControlText` 同型，这样万一还剩一条装箱路也落到 BSTR 那一档）。**不动** `vb6_GetMouseIcon` / `vb6_GetControlHwnd` —— 那两条本来就是对象，装成对象是对的（FlexGrid 的 `CellPicture` 同型：那边缺的是发码面的强转，不是档位）。判据两件夹具：`ctrlslider` 的 SE 段 10 条 + `ctrlprop` 的 CP12..CP15 —— **登记在通用段的意思必须用非 Slider 的控件来证**，那里是一枚 Label（`lblSingle.ToolTipText`）+ 一枚 TextBox（`txtOne.Tag`）+ 一枚 ListBox 的默认档；能红的是 SE1/SE2/SE3/SE4/SE6/SE11 与 CP12/CP13/CP14（BASE 逐条实测空串 / `Object` / `9` / `no`），SE5/SE7/SE8/SE9/CP15 是本来就直接拿指针、改之前也对的形状证人（诚实记着，不算进"负控证过能红"那一栏）。**SE10（`.frm` 里设计期那一条）刻意不登记成针** —— 它读回来是空的，登记等于把错的口径焊死 ⇒ 另开账 #142。发码面 2 条正向（读侧不再套 `CStr(VariantFromValue(...))`）+ 2 条反向钉住那两条旧装箱形状（反向在 BASE 上逐条命中 = 针能红）。注册仍然只往既有块追加串、零新增 `Test-*` 调用（分片归属不动）。A/B（两份同为本地 Debug）＋爆炸半径按**行型**分类：118 件输入 = 4 件变、65 行、`UNEXPLAINED-LINES=0`（其中 `Charts 2020/ucProgressCircular` 与 `czUI-main/czFormDemo` 两件是 SL-d 已解释过的存量接通）。CI 侧收线：`test-logs-vbp-3`(x64) 与 `-4`(x86) 两份 `SlidApp.out` 逐字节相同（53 行）、且与本地真跑那份逐字节相同；`CtrlProp.out`(vbp#1) 的 CP12..CP15 四条齐。Slider 这页剩下的：设计期那一条（#142）、`GotFocus`/`LostFocus` 那一档、`TickStyle` 四档与 `Min`/`Max` 超 16 位（两条都等 VB6 真值），另有 #141（子类化两份判据表剩下的不齐）与 #83（`WS_TABSTOP` / `IsDialogMessage`）两片。
               # 上一条 = 2026-09-29T07:25:00+08:00   # 本轮追加：**C29-SL-d 已出并过门（门 #194，head `c85a7a5`，attempt 1 唯一那条红是与本批无关的旧脆弱用例、attempt 2（`--failed`）11 job 全绿，wall 266s = 4m26s（attempt 2 只重跑那一个 job：23:12:15Z → 23:16:34Z）**。这一格缺的只有一条 arm，而且是**两份表不齐**：发 arm 那一趟（subclass）没有 `WM_LBUTTONUP` 档、装不装那一趟（frame_menu 的 `needsSubclass`）判据里没有 `_Click` —— 只补一处仍然是死的（实测：`ctrlfiles` 的 `fileList` 子类过程在产物里定义得好好的、一次都没被 install）。补上之后必须按 `controlClickFromNativeNotify` 把"Click 已由原生通知供着"的那批（BN_CLICKED / *_SELCHANGE / TCN_SELCHANGE / Toolbar / Menu）排除掉，否则同一次点击双发；**列表类 `DblClick` 的同型双路本批刻意没动**（记账）。判据 = 夹具 41 条里 SD 那 8 条（`SD7-install` 钉的正是"装不装那一趟"，`sld2` 除 Change 外只有 Click，而 Change 走的是父窗那条通道、与子类化无关）+ 发码面 10 条（8 正向、2 反向钉判据方法不许退化成取属性）；**注册只往既有块追加串、零新增 `Test-*` 调用**（分片归属不动，门 #191 那条波澜至此站稳）。A/B（两份同为本地 Debug）：产物 42 行只有 SD 那 7 行不同、BASE 七条全 N 且 `SD4-value` 读 50；爆炸半径按**行型**分类：118 件输入 = 3 件变、45 行、`UNEXPLAINED=0`，另两件（`Charts 2020/ucProgressCircular`、`czUI-main/czFormDemo`）是存量工程里一次真接通（它们的 `_Click` 以前是死代码）。CI 侧收线：`test-logs-vbp-3`（x64）与 `-4`（x86）两份 `SlidApp.out` 逐字节相同、且与本地真跑那份逐字节相同，两份 `.err` 各 0 字节。实测口径进 029 的 SL-d-0：轨道条在按下时自己就抢焦点 ⇒ 键那两条不靠 `WS_TABSTOP` 也到得了；裸的 UP/DBLCLK 不惊动父窗那条通道；`SetForegroundWindow` 默认被拒 ⇒ "没读数"不等于"没消息"。Slider 剩下的是 `ToolTipText`（要 `TTN_GETDISPINFO`）、`GotFocus`/`LostFocus` 那一档、以及两条等 VB6 真值的口径（`TickStyle` 四档、`Min/Max` 超 16 位）+ 账 #83 那片 `WS_TABSTOP`/`IsDialogMessage`。
               # 上一条 = 2026-09-29T05:50:00+08:00    # 本轮追加：**#128-b 已出并过门（门 #193，head `9cca28b`，attempt 1 即 11 job 全绿，wall 572s = 9m32s）⇒ 账 #128 收线**。**先把 RTL 搦干净再翻类型**：ListView 那六位（GridLines/FullRowSelect/MultiSelect/CheckBoxes/HideColumnHeaders/AllowColumnReorder）以前 getter **把写进去的原值照样回显**、设计期发的是 1，而且六位**根本没登记进类型表**（走兜底装箱）⇒ `CStr(ListView1.GridLines)` 打的是 **1**；现在存进去就归一化 -1/0 + 设计期改发 -1 + 才登 Boolean。OptionButton.Value 单开专桩（`vb6_GetOptionValue` 把 BST_CHECKED 映射成 -1 / `vb6_SetOptionValue` 把任意非 0 折回 1）—— 因为 CheckBox.Value 在 VB6 是**三态 Integer(0/1/2)**、两家原本共用一对桩；而且 **`BM_SETCHECK` 不吃 -1**，写侧那一折不能省。视图/排序那几位（View / SortKey / SortOrder / LabelEdit / Sorted）**刻意继续不登记**：前几个是枚举，`Sorted` 到底 Boolean 还是 ccSort* 本机拿不到 VB6 真值 ⇒ 押后等口径（不靠猜的枚举写代码，同 TickStyle 那条纪律）。**自己抓到的一处自己的错（本地自检抚住的，不是读代码读出来的）**：读写表里 `case CheckBox:` 与 `case OptionButton:` 是**共用同一段函数体**的，第一版把 `value` 那行直接改专桩 ⇒ **CheckBox 跟着被换掉**、灰态(2)被吃掉、`DS1` 从 `11` 变 `-1-1`。拆开两个 case 才回正。这是同一条旧纪律第三次：**共用 case 体的表格，改一个控件前先确认另一个想不想跟**。判据 = 五条新读数 + 五条发码针；**连带改别人夹具两条针按 VB6 口径改字面**（DS5 `01`→`FalseTrue`、LV14-GRID `1`→`True`）；注册**仍只往既有块里追加串、零新增 `Test-*` 调用**（SL-c 那轮量到的那条分片波澜至此站稳，本轮一次过）。诚实保留一条：`LV24`（比较面钉 `CheckBoxes = True`）在 BASE 上**没红**（装箱比较把非零当真）—— 它是行为针，不算进"负控证过能红"那一栏。A/B（两份同为本地 Debug）：新产物七件夹具针零缺失、十一块发码面零缺零命中；BASE(`pre_128`) 红在新读数与旧形状上。爆炸半径换成**按 hunk 分类**（旧符号集₆允许集、新符号集₆允许集、字面量不得变）：118 件输入 = 12 件变、107 行、`UNEXPLAINED=0`。**存量工程那五件是一次真修复**（与 #88 在 Charts 里同型）：`VBFlexGridDemo\UserEditingForm.frm:373` 写的是 `If Option1.Value = True Then`，修前发的是"先装箱再与 -1 比"而 getter 答 1 ⇒ **该分支恒假、从来没跑过**；该 demo 在 CI 只做"编译 + 起窗 3 秒不崩"（`tests_github/run_t2.ps1:285`，无 stdout 针）、而这些分支只在 Validate 事件里 ⇒ 不构成门风险，但变化写在这里、不藏在 diff 里。**本轮另一条工具链发现**（已进 memory）：本地 Debug 产物与 CI 的 Release 产物对**同一份源码**会发不同的 COM 属性读档案（各自都确定）⇒ A/B 两端必须同构建类型；第一版护栏因此虚报过 9 件 / 1347 行。Slider 一族（SL-a/b/c）与 #128 到此全部收线；下一步只剩等拍板的口径格（Slider `TickStyle`、`Min/Max` 超 16 位、ListView `Sorted`、`ToolTipText` 要 `TTN_GETDISPINFO`）。STATUS 保持 BUSY。
               # 上一档 = 2026-09-29T04:00:00+08:00   # 本轮追加：**#128-a 已出并过门（门 #192，head `32a20b8`，attempt 1 即 11 job 全绿，wall 551s = 9m11s）** = 把 #124 那一刀推平到其余控件：**先逐枚读 RTL getter 再决定谁能翻**（这格的工作量在分类）—— TreeView `CheckBoxes/HotTracking/HideSelection`（走 `TvBitAsVbBool`，HideSelection 还是反位）、MonthView `MultiSelect/ShowToday/ShowWeekNumbers`、DTPicker `CheckBox`、SSTab `WordWrap`、RichTextBox `ReadOnly/WordWrap`（后者 setter 里就已归化成 `on ? -1 : 0`）六枚确定答 -1/0 的归 Boolean。**两批刻意没动，理由是它们的 getter 不干净**（= 下一格 #128-b）：ListView 那族把写进去的原值**照样回显**（`v->multiSelect = (int)val`）⇒ 先翻类型会出玶"说是 Boolean、值是 1"，`X = True` 这类**比较恒假**；OptionButton.Value 与 CheckBox.Value **共用 `vb6_GetCheckValue`/`vb6_SetCheckValue`**（直接搬 BM_GETCHECK 的 0/1/2，而 CheckBox 的 Value 在 VB6 是**三态 Integer**）⇒ 正解是给 OptionButton 单开一对，注意 `BM_SETCHECK` 吃不了 -1。另排除两条看着像布尔其实是枚举的：ListView `LabelEdit`、`Sorted`。判据 = 五件既存夹具各加一条原始读数（`TV38=True/False/Boolean` 等五条，针面直接钉字面）+ 发码面四块 shape 各补一条正向、四块 absent 各补一条旧形状；**注册只往既有数组/块里追加串、一条 `Test-*` 调用都没新增**（上一格刚量清"插一条调用会重排后续所有用例的分片归属"，这样写就绕开，本轮门也确实一次过）。**连带改了别人夹具两条针**（#124 同型）：`ctrlsstab` 的 TS6/TS13 打的是裸 `& SSTab1.WordWrap` 拼接，归 Boolean 后拼出 False/True，针从 `=0`/`=-1` 改字面，同工程其余 34 条一字未动。A/B（**两份同构建类型**，SL-c 那格订正的口径）：新产物五件夹具针零缺失、八块发码面零缺零命中；BASE(`pre_128`) 恰红 8 条（5 新 + TS6/TS13）+ 缺 4 条正向针 + 命中 4 条旧形状，**其余针在 BASE 上照旧命中**；爆炸半径 = 118 件输入 `changed_files=5`（正是这五件夹具）、`changed_lines=14`、`UNCLASSIFIED=0`（每行差异必含这六枚 getter 名，按符号分类器判定）。下一格 = **#128-b**（RTL 先归化再翻类型：ListView 回显族、OptionButton.Value 专桩）。STATUS 保持 BUSY。
               # 上一档 = 2026-09-29T03:00:00+08:00   # 本轮追加：**C29-SL-c 已出并过门（门 #191，head `18a62b2`，attempt 2 全绿）** = Slider 的 `Change` / `Scroll` 两条事件接上了原生那条通道。**格子本身最值得记的是两件事**：① **分法是量出来的**（探针 `.build/slprobe/slmeasure7.c` + `8.c`，父窗 WndProc 逐条记 wParam）：Slider 与 ScrollBar **共用同一扇门** —— 横杆发 `WM_HSCROLL`、竖杆发 `WM_VSCROLL`，  值在 **高字**（一次真拖 = `5×N → 4 → 8`，高字 30,32,…,49 跟着控件走），`lParam` = 控件句柄；方向键是 `0 → 8`；**程序化 `TBM_SETPOS`/`SETRANGE` 一条都不发**（所以 `SC10/SC11` 是这条口径的哨兵，与 DT-c 的 `DT41` 同型）。  `Change` 因此按"**值真的变了**"发（现问 `GETPOS` 与自存基准比，拖拽中那一串 5 天然连续触发、落点 4 与收尾 8 不会多发），而不是按码表枚举 —— 两条**鉴别针**（`SC2` 同值的第二条 5、`SC11` SetValue 之后的同值通知）钉的就是这个区分；  `Scroll` 只认码 4/5，与仓里 HScrollBar/VScrollBar **已发货的口径同一档**（同一条原生通道不允许两套分法）。② **护栏第一版是错的，而且错得很像真回归**：我拿"本地 Debug 的 `.build/C3.exe`" 对 "CI 门里下来的 Release 产物" 逐行比 `--emit-c`，报出 **9 件工程 / 1347 行**差异，逐行看全是 `vb6_ComGetIntProp(oFont, L"Name")` vs `vb6_ComGetStringProp(...)` 这种 **COM 属性读法档案**（= 账 #82 那一族）。  两份产物各自都确定（同一 exe 跑两次逐字相同），所以差异来自 **构建类型本身**：门用 `-DCMAKE_BUILD_TYPE=Release`，`scripts/dev.ps1` 是 Debug。  改成"两份都是本地冷编 Release"（BASE = `git worktree` 到 HEAD 再 cmake Release，产物 `cp` 回主树保持输入路径一致）⇒ 数字立刻收敛到 **118 件输入里只 `ctrlslider` 一件变**。这条口径以后所有 A/B 都按它走。判据 = 夹具扩 11 条读数（SC1..SC11，事件断言只数增量）+ 新发码组 `sl_emitc_events` 7 正向 / 反向补 2；负控 = 门 #190 的产物编同一件 ⇒ 33 条里红恰好 11 条（SL/SB 一条没动，`SC7` 在那边读到 **20** = 没有一次事件动过值）、发码针缺 6、反向针命中 1；新产物 x64/x86 逐字相同、`.err` 0 字节。方法面按三处接（`sliderVars_` + 登记 + `com_bind` 改道，本线第五次碰同一坑）。**门的波澜**：attempt 1 唯一的红是 `frmevents` 的 `Got:` **整份空白**（#79 那条旧脆弱，不是部分读数不对），归因查了三件：该工程的 `--emit-c` 在 BASE/NEW 之间 **0 行差异**、本地用新 exe 真跑 17 行全齐 rc=0、而它落到这一片是**本批新增一条 `Test-EmitcShape` 把分片轮转位置挪走了**（#190 那轮它在 vbp #4）⇒ 记一条一般式：**往 `run_tests.ps1` 中间插一件用例会重排后续所有用例的分片归属**，时序敏感的用例由此换机器、换并发环境。`gh run rerun --failed` 只重跑那一个 job（272s）即绿，attempt 2 工件里 `SlidApp.out` 两架构 34 行逐字相同、33 条针零缺失。Slider 还欠的三格全部在等口径（`TickStyle` 四档、`Min/Max` 超 16 位、`ToolTipText` 要 `TTN_GETDISPINFO`）。下一格 = **#128**（其余控件的布尔属性仍在 Long 档：TreeView `CheckBoxes/HotTracking/HideSelection`、ListView `MultiSelect/CheckBoxes`、MonthView `MultiSelect/ShowToday/ShowWeekNumbers`、SSTab `WordWrap`、DTPicker `CheckBox`、RichTextBox `ReadOnly/WordWrap`、OptionButton `Value`，改法照 #124 那一刀）。STATUS 保持 BUSY。
               # 上一档 = 2026-09-29T01:43:32+08:00   # 本轮追加：**C29-SL-b 已出并过门（门 #190，head `b93e62d`，11 个 job 全绿，wall 569s = 9m29s）** = Slider 的值面（`Min/Max/Value` + `Small`/`LargeChange` + `SelStart/SelEnd/SelectRange`）。**值面全部直问直发控件，一格自存都没有**（唯一自存的还是 SL-a 那格 `TickFrequency` —— 原生问不出）；`vb6_Slider_Init` 从一参扩成十参（`-999` 一律"没写过"），读写表各补九格，`SelectRange` 按 #124 口径登记成 Boolean。**参数顺序是量出来的**（`slmeasure6.c`）：改 range 会连带重算三样（pos 顶到下限、`page` 从 20 变 18、`selstart` 跟到下限）⇒ Init 必须 range → line/page → pos → 刻度 → Sel，判据里 `SB3-changes=2/8` 就是这条顺序的证人。四条实测口径：默认档 `page=20`（VB6 文档写 5，本机 OCX 未注册拿不到真值 ⇒ **照原生答 20、不自作主张折成 5**，`SB9-defaults` 钉住这个选择，哪天有了 VB6 真值这条会翻红逼人来拍）；越界由控件钳位（150→100、5→10，我们不自己钳）；`Sel*` 没挂 `TBS_ENABLESELRANGE` 时 `TBM_SETSEL` 不生效、运行期改这一位也有效；`CLEARSEL` 之后 `GETSELSTART/END` 答 (UINT)-1 ⇒ getter 折成 0。Min/Max 是 Long 而原生只吃 16 位（实测 `MAKELONG(40000,50000)` 读回 -25536/-15536 = 截断，SDK 里**没有** `TBM_SETRANGE32`）⇒ 下发前钳到 ±32767、读回也是那一个钳过的数；VB6 超界怎么办拿不到真值 ⇒ 押后等口径。**自己这一格抓到两处自己的错，都是真跑当场翻出来的**：① SL-a 的 RTL 里手写的一串 `#ifndef TBM_xxx` 兜底编号中 `TBM_GETTHUMBRECT` 写成 `WM_USER+17`（头里那是 `TBM_GETSELSTART`，真值 +25）⇒ 整块删掉，消息名一律用 SDK 头，不再抄第二份编号表；② Init 扩参第一版把"要不要发 range"写成 `lo != min`，而没写时 `lo` 已被填成控件当前值、`min` 还是哨兵 -999 ⇒ 永远不等 ⇒ 每枚 Slider（包括什么都没写的）都被重发一次 range、连带重算 page、把刻度画了出来 —— 症状是 SL-a 已收线的 `SL7-nofreq=00` 当场变 `0-1`，**是上一格的判据把这一格兜住了**。判据 = `tests/ctrlslider` 扩到 22 条读数（x64 与 x86 各一臂、CI 两份工件 `cmp` 逐字节相同、`.err` 各 0 字节）+ 发码 12 正向针 / 5 反向针；BASE(`pre_124`) 负控 = 22 条缺 21、12 条发码针全缺、5 条反向针命中 3；爆炸半径 = 118 件 .vbp 输入里只有 `ctrlslider` 发出 `vb6_Slider_`（51 处）。下一格 = **SL-c**（`Change` / `Scroll` 两条事件，走 `WM_HSCROLL`/`WM_VSCROLL`；已用探针 `slmeasure7/8.c` 量出通道与负载布局）。STATUS 保持 BUSY。
               # 上一档 = 2026-09-29T01:10:00+08:00   # 本轮追加：**C29-SL-a 已出并过门（门 #189，head `d13b652`，11 个 job 全绿，wall 627s = 10m27s）** = Slider（`Begin MSComctlLib.Slider`）**从此有一枚真窗口**：登记之前它压根不创建（parser 那张按子串匹配的表里没有 "slider" ⇒ Unknown ⇒ `controlTypeToWin32Class` 给 nullptr），`vb6_hwnd_sld1` 恒 NULL，属性读写全落 `vb6_ComGetProp/vb6_ComSetProp(NULL, …)` ⇒ 真跑读数全空、写进去静默丢、退出码照旧 0（探针 `.build/slprobe/` 量的）。照 D6 走原生 `msctls_trackbar32`，登记面照 DTPicker 那格的清单数出来（parser 两处 + cgen 三处 + 新 RTL 文件**四处**登记：`c3rtl.rc` 221 / CMakeLists / rtl_embedded 两处 / `driver_link.cpp` 的 `opts.sourceFiles`）。**本格真正的收获是把自己第一发的结论订正掉**：那发拿 `TBM_GETCHANNELRECT` 的长短边判方向，得出"Orientation 运行期改不动"—— 后来看清这条消息的矩形**永远把行程长度放在 x 分量**（300x40 的横杆与 40x300 的竖杆都答 `(8,10)-(292,14)`），量它等于什么都没量；改问 `TBM_GETTHUMBRECT` 之后：创建水平再写 `TBS_VERT` + `SWP_FRAMECHANGED`，滑块矩形与"创建时就是竖"逐字相同，对照组只换帧不翻样式 ⇒ 不变 ⇒ **运行期翻向有效**。稳的证人是**滑块长短边**（横杆 11x22 的竖块、竖杆 22x11 的横块），而"推两端看位移轴"在 133x27 那种竖直行程被压成 0 的窗口上退化。另量到：**光挂 `TBS_AUTOTICKS` 不给 `TBM_SETTICFREQ` 是不画刻度的**。判据 = 新夹具 `tests/ctrlslider`（x64 与 x86 各 12 行、CI 两份工件 `cmp` **逐字节相同**）+ 发码 7 正向针 3 反向针；A/B = BASE(`pre_124`) 编同一件夹具 ⇒ 11 条读数针全红（针能红验过）；护栏 = 语料里零件工程用过 Slider 且改动全关在 `controlType == Slider` 里 ⇒ 存量发码逐字节不动。下一条 = **SL-b**（值面 `Min/Max/Value` + `Small`/`LargeChange` + `Sel*`；顺带要拍"原生范围只有 16 位而 VB6 的 Min/Max 是 Long"那条口径）。STATUS 保持 BUSY。
               # 上一档 = 2026-09-29T00:10:00+08:00   # 本轮 = **账 #124 收线（门 #188，head `ecc8333`，8/8 全绿，wall 524s = 8m44s）** = 控件的布尔属性 `Enabled` / `Visible` 在类型 oracle 里不是 Boolean：`CStr(cb.Enabled)` 打 **-1**、`TypeName` 给 **Long**、装箱 **VT_I4(3)**，VB6 三处是 True / Boolean / **VT_BOOL(11)**（`CStr(True)` 早就是对的，差的只是控件属性这一档）。根因一行：`controlPropType` 把 `left/top/width/height/visible/enabled` **一起**按 Long 登记 —— 那张表原本只为"按成员裸名查符号会被内置函数顶掉"而建，四个几何属性确实该是 Long，这两个跟着吃了一刀。改法 = 把这两个拆出来返回 Boolean（RTL getter 仍返回 int，需要窄化的两处本来自带 `(int16_t)`）。**判据**：`tests/ctrlstate` 扩四条、四个消费面各钉一面（CStr=DS2/DS3、TypeName+VarType=DS8、装箱两向=DS9/DS10）+ 发码 3 正向针 2 反向针；**装箱那两条形状不同是实测的**（赋值点 `wrapVariantValue` 不加窄化、实参点 `boxToVariant` 自带收窄）。**连带改了一枚别人的夹具**：`ctrlsstab` 的 TS25/26/27/28 打的也是 `.Visible`，针按 VB6 口径从 -1/0 改成 True/False，而 TS20/TS21/TS25B（TabVisible/Tabs 是 int）在 CI 工件里**逐字没动** —— 这半边才是"只翻该翻的"。护栏（只跑 .vbp 117 件，BASE = `pre_124`）= changed_lines 26 / touched 2 件、逐行可归类。另立 **#128**（其余控件的布尔属性仍在 Long 档：`OptionButton.Value`、SSTab `WordWrap`、TreeView/DTPicker/ListView 一批），理由与改法写在 029 本格末。下一条 = **C29-SL（Slider 滑杆）**，用户点名开工，规范 `ai\内置控件\Slider 控件（滑杆）.md`。STATUS 保持 BUSY。
               # 上一档 = 2026-09-28T23:40:00+08:00   # 本轮 = **账 #125 收线（门 #187，head `aa92fde`，8/8 全绿，wall 577s = 9m37s —— 对方刚把 vbp 拆 4 片；本批 `ec945c1` 是它的祖先，核过）** = 设计期 `Enabled` / `Visible` / CheckBox·OptionButton 的 `Value` 三件此前**两条创建路都没打到窗口上**（发码里 `BM_SETCHECK` / `EnableWindow` / `ShowWindow` 各 0 次）。修法 = 新增 `emitDesignerStateProps`（与 `emitDesignerStyleProps` 并排），顶层与容器子控件两条路各接一行，只在 .frm **显式写过**该项时才发（Enabled/Visible 只认 0 那档、CheckBox 的 Value 只认 1/2、OptionButton 非 0 折 `BST_CHECKED`；Timer/Menu/ImageList 直接 return）。**判据** = 新夹具 `tests/ctrlstate`（x64+x86 各一臂，Frame 子控件与顶层各一份同形状 + 两个"什么都没写"的选择性闸）：`DS1..DS7` 七条读数 + 发码 7 条正向针 + 3 条选择性反向针；负控 = 同一份判定代码跑 BASE 产物 ⇒ 红 `DS1/DS2/DS3/DS5` + 7 条发码针全缺（针能红验过）。护栏 = tests/ 全语料 459 件输入逐行比对：`changed_lines=29 / touched=7 件`，29 行**逐行可归类**（全是那三个 setter），其余零差异；合并后复量口径换成"只数 `/* design … */` 标记行" = 36 行 / 8 件（= 29 + 本批新夹具 7），证明合并没改本批的爆炸半径。**踩出一条判据纪律**：`Visible` 只能在窗体真显示之后问 —— `Form_Load` 里父窗还没 Show，`IsWindowVisible` 对任何控件都返回假（第一版探针因此把三条结论弄反过），读数一律放 Timer。**本批先跑的 #186（head `ec945c1`）红在 vbp 分片 #1/#2/#3 的三个接续用例**（`ctrlmanifest_builtin` / `balloon_manifest_is_user_supplied` / `frx_extract_content`），根因是对方上一版 CI 分片把"依赖基座产物的用例"拆散、与本批无关；`aa92fde` 按依赖分组修好后全绿。下一格 = **#124**（控件的布尔属性读回是 `int` 不是 Boolean ⇒ `CStr(cb.Enabled)` 打 -1 而 `CStr(True)` 已是对的 True）。STATUS 保持 BUSY。
               # 上一档 = 2026-09-28T19:05:00+08:00   # 本轮 = **账 #88 收线（门 #182，head `00f241c`，8/8 全绿，wall 973s = 16m13s；vbp 整格 PASS=126 FAIL=0 SKIP=1，本批只升级既有夹具判据、没新增用例）** = 比较上下文里的 COM 字符串成员被按数值解包：`tv1.Nodes(2).Text = "字面量"` **编得过、跑起来恒假**（发码 `vb6_CStrLong(vb6_ComGetIntProp(o, L"Text"))` = 拿 BSTR 指针当数字与字面量比）。根因一行：`visit(BinaryExpr&)` 只有 `Concat` 走 `resolveComValue()`（默认 BSTR），其余二元运算**一律** `resolveComValue("Long")` ⇒ 六条比较被顺带按数值解封。修法只看对侧（`hint88`）：Eq/Neq/Lt/Gt/Le/Ge 且对侧是串 ⇒ BSTR 档，算术与其余维持 Long 逐字节不动。**判据**：TV15 从"先取进局部 String 变量绕路"升级回**直接比**（BASE N / NEW Y）、新增 TV35 链式、TV36 必须仍然 N 的负向闸、TV37 钉**选择性**（同一条链上的 Index/Children 照旧 int 档）、TV34 保留旧路径；发码三根针（BSTR 正向 + int 档 Index 正向挡反向过度修正 + BASE 那条 CStrLong(GetIntProp) 负向）。A/B：NEW x64 与 x86 逐字相同；拿判据代码本身复跑 NEW `shape-missing=0/absent-hit=0/ALL MATCH`、BASE 缺两条正向针且命中那条负向针（#178 那轮红换来的纪律：结论必须过一遍真正判定它的代码）。护栏 = 116 件工程 `--emit-c` 逐行比对，**只有 VBFlexGridDemo 的 5 处变**且每处都是同一缺陷被修好，数值成员行逐字节不动。**同批订正两笔旧账**：#85（`CStr(集合成员)` 打空）实测已不复现 ⇒ 无代码改动作废；#120（用户过程与窗体属性同名）全仓 875 个源文件 **0 处**真碰撞 ⇒ 纸面缺口押后等口径。**另外用探针 `.build/p83/` 新记四格欠账并收窄 #83**（只记不修）：#124 控件布尔属性读回是 `int` 不是 Boolean（`CStr(cb.Enabled)` 打 **-1**，而 `CStr(True)` 已是对的 True —— 受害面 Print/CStr/TypeName/装箱四面，等值比较靠 -1 恰好对上所以抓不到）；#125 设计期 `Value = 1` 两条创建路都没发（发码里 `BM_SETCHECK` 0 次）；#126 设计期 `Enabled = 0` 两条路都没发（同形状 `Visible = 0` 倒是发了 ⇒ 不对称）；#127 Frame 的 `Visible` 读回 False 而窗口实际可见。⇒ #83 剩下的实证只有 **WS_TABSTOP**（同一枚 CheckBox 顶层 style `0x54000003`、子控件 `0x50000003`）与顶层那长串专有项没铺。下一格 = **#125 + #126 一起收**（同一个发射点，真跑可见的 VB6 齐平缺陷）。STATUS 保持 BUSY。
               # 上一档 = 2026-09-28T17:37:17+08:00   # 本轮 = **账 #123 收线（门 #181，head `77c251d`，8/8 全绿，wall 965s = 16m05s；bas 两分片 38+39 = 77/77、vbp 整格 `PASS=126 FAIL=0 SKIP=1 TOTAL=127`）**。这一格是用户一句「记忆有可能出错，实践才是检验真理的标准」把我顶回去实测才定性对的：根因是 `As Byte` 的 C 型 `uint8_t` 不匹配 local/param 的任何一条登记分支 ⇒ 局部 Byte 与 Byte 形参**进不了任何一张 `knownXxxVars_`**（全仓原先没有 `knownByteVars_`），`inferExprType` 答 Unknown、`isDefinitelyVariantExpr` 回退查符号表；模块级 `As Byte` 反而正常（走符号表）—— 同一类型两个作用域两套可见性。后果三件全部实测：① `s = 局部Byte` 发 `vb6_BSTR_Assign(&s, bt)` ⇒ 数字当 BSTR 指针存、跑起来 SIGSEGV rc=139；② **`(bt = 65)` 今天返回 False**（以前我只记成"读邻居栈的潜伏 UB"，低估了）—— 发的是 `vb6_VarCmpLongEq(&bt, 65)`，助手签名 `(vb6_VARIANT*, int32_t)`，拿 1 字节对象的地址当 16 字节 VARIANT 传；③ 装箱那条是**预存的反向错**：`wrapVariantValue` 的 Byte 档发 `vb6_VariantLong` ⇒ VT_I4=3，而 VB6 要 VT_UI1=17 —— 模块级 Byte 与 `CByte()` 早就中，局部 Byte 反倒因为不可见掉进 `_Generic` 而凑对 17，本批三形统一到 17。改动 7 处（state 声明 + inferExprType + isDefinitelyVariantExpr + localdecl 两处 + func/proc/prop 三处）。**两次我自己的错都是"跑一遍"抓住的、项目 diff 一声不吭**：第一版照 Boolean 先例把 Byte 塞进 `knownLongVars_` ⇒ `VarType` 17→3，而同一时刻 11 件存量输入 `--emit-c` 逐行零差异 ⇒ "存量零 diff"只证明存量没这形状，不证明改动惰性；第二版形参登记写进了 `else if (paramType == Long||Integer||Boolean)` 的**分支体内**，Byte 根本进不去 ⇒ 夹具 `B16` 读回 NE200 才暴露。判据 = `tests/test_byte_visible.bas`（x64+x86 各 26 条读数、逐字相同、`.err` 0 字节）+ 两枚发码断言（反向那枚钉 `vb6_VarCmpLongEq(&bt,` 一类形状不许回来）；A/B 喂 pre_123 ⇒ 跑到 B6 当场崩且崩之前 B1/B5 已是错答案、发码面 shape-missing=4 + absent-hit=2。波及半径只有 Charts 2020 差 4 行，其中 `Alpha < 128` 那一处是**真实工程里的同一症状被修掉**，本批随 vbp 那 126 枚一起跑绿。全程见 029 账 #123。
               # 上一档 = 2026-09-28T16:20:00+08:00   # 本轮 = **账 #122 收线（门 #180，head `30a9066`，8/8 全绿，wall 992s = 16m32s；bas 两分片 38+37 = 75/75、vbp 整格 `PASS=124 FAIL=0 SKIP=1 TOTAL=125`）**。从 #119 的测量里长出来的独立一格，根因不同：String 目标的赋值缺 VB6 的**隐式 CStr** —— `Dim s As String: s = 123` / `s = n`(Long) / `s = d`(Double) 这类今天原样交给 `vb6_BSTR_Assign`。后果分两种，**我在 #119 那格把它记成"响的"是记错了半边**，本轮一并订正：RTL 按 C 编（`cl /std:c11`），`int` 传给 `BSTR` 形参只是警告 ⇒ **整型那半今天编得过、跑起来才崩**（实测修复前二进制编 `s = n`：BUILD-RC=0、跑起来 SIGSEGV，bash 读 rc=139，stdout 一个字节都没有）；只有 Double/Date/Currency 那半才是响的 `error C2440 无法从 double 转换为 BSTR`。改一处 cgen：值的 `inferExprType` 落在既有 `isScalarImplicitStr159` 清单里时过一遍**现成的** `wrapToBSTR`（Fix 159-B 给 ByVal String 形参用的同一张折算表），折出来的新块顺手算进 #119 的「自有」侧（`valueOwned113 || wrapped122` ⇒ AssignMove，不新增泄漏）；没另开折算表、没动类型推断。**这一格最值钱的是护栏抓到我自己的回归**：7 枚存量工程 BASE vs NEW 逐行 diff 本应零改动（那种形状今天编不过 ⇒ 不可能存在于跑通的工程里），结果 Charts 2020 差 4 行，逐条定性 ——① `Assign(&vb6_ret_Caption, vb6_ByteArrayToString(me->m_Caption))` 被我套成 `vb6_CStrByte(vb6_ByteArrayToString(..))` = **真回归**：Fix 140 那条 Byte() 还原支已经把值转成字符串，但 AST 节点仍是 Byte 数组 ⇒ 只看类型的守卫重复折了一层；修法 = 那支自记 `wrappedByteArr140`，#122 见到就跳（与 Variant 那支同处理）。② `AssignMove(&sDiplay, vb6_CLng(me->m_Value))` 外面多了 `vb6_CStrLong(...)` = **不是回归，是这一格在它身上的真实修复**（VB6 `sDiplay = CLng(...)` 该得十进制串，修复前存的是 long）。修完复跑：6 枚逐字节相同，只剩 Charts 2020 那一行。**判据** `tests/test_str_cnum_assign.bas`（x64 + x86）19 条读数全取自真实输出、两架构逐字相同；日期不比字面文本（随区域设置变），只比同一折算路两个来源是否一致；四条负向钉「值本来就是字符串的一条都不许折」；针面不带方括号（#178 的教训）。**A/B**：同一枚夹具喂修复前二进制 ⇒ BUILD-RC=1；**负控**：同一套发码针面喂修复前二进制 ⇒ shape-missing=4。CI 那 120 枚 GUI 工程全绿 = 那行真实修复（`sDiplay`）运行期被读过也没红。**刻意没顺手做**：局部 `Dim bt As Byte` 到现在还崩 —— `inferExprType` 只有 Bstr/Single/Date/Double/Bool/Long/LongPtr/Variant 八张 `knownXxxVars_`，压根没有 Byte 那张（全仓无 `knownByteVars_`），局部 Byte 推不出类型 ⇒ 守卫不介入。那是类型可见性、动它会波及所有 `inferExprType` 调用点，另立任务 #123。全程见 029 账 #122。
               # 上一档 = 2026-09-28T15:35:00+08:00   # 本轮 = **账 #119 收线（门 #179，head `6e21ada`，8/8 全绿，wall 951s = 15m51s；bas 两分片 36+37 = 73/73、vbp 整格 `PASS=122 FAIL=0 SKIP=1 TOTAL=123`）**。这一格是 #113 的另一半：#113 为了少翻发码面，把新加的 `vb6_BSTR_AssignMove` 用 `bstrTargetViaTypeQuery` 那道闸**只**放行给模块级目标，局部/形参/UDT String 字段三条老路从此每句 `s = Left(x,2)` 都深拷贝一份、把刚 alloc 的临时串扔掉 ⇒ 纯漏。本批删掉那道闸（发码条件回到「值是自有临时 ⇒ Move」），并把「字面量算自有」限定成 `LiteralKind::String`（数值字面量赋给 String 目标发的是裸数字，cl `error C2440`，另立新账未修，不能顺手标成自有）。**铺开的底气 = 这一轮把 RTL 全扫了**：77 枚返回 `BSTR` 的函数里只有 3 枚发借来的指针（`vb6_ErrDescription` / `vb6_ErrSource` 直接回 `vb6_err` 那两字段、`vb6_ho_variantToBstr` 回 `v->bstrVal`），而三者都只能以成员访问形态到达，判定那一支（只认 `IdentifierExpr` callee）根本捡不到 ⇒ 不会误判自有。**判据**：`tests/test_str_leak.bas`（x64 + x86）读 `K32GetProcessMemoryInfo` 的 PagefileUsage **增量**，10 万次 `s = Left(gBig, 64)` 阈值 2MB，CI 实测 `delta=0` / `8192`；`api1=1 / api2=1` 是自带的失效负控（少了它，第二次读数失败会让 delta 自己减自己 = 0 = 假绿）；另有六条所有权语义护栏与两枚 `--emit-c` 断言。为此在 `vb6_di_stubs.c` 末尾手写一枚 `K32GetProcessMemoryInfo` 桩（按本仓惯例不重跑 `gen_di_stubs`）。**护栏**：7 枚存量工程 BASE vs NEW 的 `--emit-c` 全文逐行比 ⇒ **150 处改动全是 `Assign`→`AssignMove` 一对一翻转、其余一字未动**（Charts 2020 占 141 处），7 枚真构建全 RC=0。**门 #178 那一跑红了两枚，红的是我自己的针**：套件用 PowerShell `-like` 比针，`[ABC]` 在那里是**字符类** ⇒ 带方括号的读数永远匹配不上（全仓只有我这两条这么写）；产品读数逐条正确、其余 34 枚 bas 与 vbp 整格全绿 ⇒ 本批 150 处翻转没有造出任何 double-free。加固提交 `6e21ada` 去掉方括号后 #179 全绿。复盘见 029 账 #119 那一节。
               # 上一档 = 2026-09-28T14:05:00+08:00   # 本轮 = **账 #91 收线（门 #177，head `40b9e88`，8/8 全绿，wall 991s = 16m31s；bas 两分片 36+35 = 71/71、vbp 整格 `PASS=120 FAIL=0 SKIP=1 TOTAL=121`）**。用户问「#91 别人那分支应该修过了吧」—— 查下来两条读数都不成立、而且确实早已修好：① `DllGetVersion` 报 `E_INVALIDARG` 是**我自己的探针少写一参**（真签名 3 参 `(ByVal pVersion As DLLVERSIONINFO)`），补上 `dv.cbSize = Len(dv)` 后读回 `hr=0 / maj=6 min=16`，编译器无责；② `GetModuleFileNameA` 崩由别人的三刀修掉 —— `14c36a6`（A 版 API 的 String 出参回写 + 绕开 SDK 的 A/W 宏抢占）、`18ea80d`（定长 BSTR 语义，修出参缓冲堆越界）、`baea726`（出参回写须再验生成串是左值），且 `tests/declare_out/` 里六枚用例早已在册跑着，本轮实测两形读数都在。这一格**零产品码改动**，只给 `declare_gmn_path_out` 补一枚**定长缓冲**读数：`Dim bufF As String * 260` + `GetModuleFileName(0, bufF, 259)` → `nF=56`，与同程序里动态串那形的 `n=56` 同值，针名 `gmn-fixed-ok=Y`（`gmn-path-ok=Y` 那枚旧的继续在场）。订正全文见 029 §九 账 #91 那一节。
               # 上一档 = 2026-09-28T13:35:00+08:00   # 本轮 = **账 #107 收线（门 #176，head `6fe4b33`，8/8 全绿，wall 978s = 16m18s；vbp 整格 `PASS=120 FAIL=0 SKIP=1 TOTAL=121`）**。`BorderStyle = 0`（None）对所有走"存窗口属性"兜底的控件（ListBox / ComboBox / Frame / …，即非 Edit·RICHEDIT50W·非 Static 那一片）**设不上**：`SetPropW(hw, name, (HANDLE)0)` 的语义是 RemoveProp，写进去当场消失，`vb6_GetBorderStyle` 的 `GetPropW` 回 NULL 就落进"按类名/样式猜默认值"那段 —— ListBox 实测读回 **2**（`WS_BORDER` 被那把 `WS_OVERLAPPEDWINDOW` 尺当成 CAPTION 位）。读写两边统一改成 **`val+1` 存**（本仓同族第三处：Fix 187 给 BackColor 加了独立哨兵 `VB6_BackColorSet`、CommonDialog 那格注释里写着"整数一律存 val+1，因为 SetPropW(0) 与'从没设过'不可分辨"），Static 那格虽然自己翻 `WS_BORDER` 但**读**仍走属性，故一起改；设计期发射（`cgen_util_ctrl.cpp:937`）顺带修好。**同族逐处量过**：`VB6_ForeColor`（默认本就是黑）、`VB6_MultiLine`/`VB6_ScrollBars`（读侧从样式位推，往返自洽）、`VB6_PictureType`（0=none 也正好是默认）、`VB6_PrintX/Y`（默认 0）⇒ 只有 BorderStyle 这一处真的坏，但"其余只是巧合没坏"这条已写进 029 当口径。**判据**：`ctrlprop` 同一枚 ListBox 三次赋值 0/1/0 → `CP9=0` / `CP10=1` / `CP11=0`；**A/B**：BASE = `.build/fls/pre_113_C3.exe` 给 `2 / 1 / 2` ⇒ 两条 0 针修复前必红。这一格动 RTL ⇒ 按惯例 `touch src/driver/c3rtl.rc` 后走 `.build/rt_ninja.ps1`，门的 GUI 面正好当护栏。
               # 上一档 = 2026-09-28T06:20:00+08:00   # 本轮 = **C29-WS-e 已出（本地 x64 三跑 + x86 两跑 = 54/54；BASE = 门 #162 的工件负控是两截的）** = `GetData` / `PeekData` 的 **Byte 数组那一形**（`vbByteArray` = VT_ARRAY|VT_UI1 = 8192+17 = 8209）。五条读数（全部见 029 §九 C29-WS-e）：① VB6 的形态是**控件把那个变量的数组描述符换掉**，所以旧的那枚由 RTL 销毁，且销前先认 `signature == 0x5A1D`（Fix 082g 那枚 1D 魔数就是为这种场合备的）；② 它与字符串那一形的**唯一**区别是不过码页 ⇒ 判据载荷一律纯 ASCII，否则字节数随 CI 的 ACP 变（同账 #79 那族）；③ 没数据给 `count = 0` 的空数组（UBound=-1、LBound=0 ⇒ `UBound-LBound+1` 读 0），`maxLen` 只切前若干字节、剩下必须还在缓冲里（WS53/WS54 钉的就是分次取不丢不重）；③ **cgen 一行没改**：Byte 数组变量在 C 里就是`vb6_SafeArray1D* gBytes`，现有出参路已经发 `&gBytes`，与字符串那形 `&gGotB` 同形，槽位类型不同而已；⑤ 负控两截：修复前的编译器**根本不认 `vbByteArray`**（常量面是新开的，表里有 vbArray/vbByte 独缺这个合成值）；换成裸 `8209` 再喂它 ⇒ 旧 RTL 忽略 type、把 BSTR 写进数组槽 ⇒ 一读 `UBound` 就 **0xC0000005**（`RUN_EXIT=-1073741819`）⇒ 顺手钉住一条设计取舍：**分流必须按 `type`，不能让通用串路去猜实形类型**。另有一条工具性的：`vb6rtl_array.h` 里要 `vb6_VARIANT` ⇒ `#include "vb6rtl_variant.h"` 必须排在它前面（第一次编译就炸在这里：C2146/C2081/C2061 一片，看不出是包顺序问题）。判据 48→**54**、发码正 27 / 反 11 条；邻居 `ctrlrichtextbox` 89/89、`ctrldatetime` 42/42、`test_len_width` 12/12。前两跑：**门 #167（head `43ae6e5` = WS-d 加固）8/8 绿**，门 #168（head `e0501e9` = 账 #115 收线）在跑。**Winsock 这一族到此收口**，剩两格都是"量到但没做出"的：多客户端（第二条 `FD_ACCEPT`）与 UDP 广播。**门 = 下面这次 push。**
               # 上一档 = 2026-09-28T05:55:00+08:00   # 本轮 = **账 #115 收线：`Len(<裸标识符>)` 的存储宽度兜底桶不再吞字符串（新用例 x64 + x86 各真跑 = 12/12；BASE = 门 #162 的工件只红 LW2/LW4/LW11，其余九条逐字相同）**。先量清范围（用户拿 `Len("你好hello")` 质疑过，质疑对、范围比我原来记的窄）：字面量、局部 String、**String 形参（无论按值/按引用）**三形本来都对；错的只有**模块级 String 变量**与 `Len(vbCrLf)`（发 `(int32_t)sizeof(...)`，**错读数自己随架构变**：BASE 在 x64 读 8、x86 读 4 ⇒ 这类“宽度”缺陷必须两架构各真跑）。根因：判据看的是不是落在 `knownBstrVars_`/`knownVariantVars_`/`knownObjectVars_` 三张表里，而 **`knownBstrVars_` 每过程入口 clear（cgen_decl_func.cpp:43、cgen_decl_proc.cpp:51）、只由局部声明与形参填**，模块级变量永远不在表里。修法只在那个 `if` 上再挂两把尺、**不动数值分支**：`knownFixedStringLen_`（定长串模块级声明处也登记，C 型不是 BSTR，继续 sizeof）+ `inferExprType(...) == Vb6Type::String`（符号表认得的串回 `vb6_Len`）。零收割走了两条证法（以后凡"全局类型判定"的改动都照这样走）：① 30 件工程逐字节护栏（BASE 换门 #162 工件）→ 唯一 diff 是 `VBFlexGridDemo\Common\Common.bas` 的 `ComGetStringProp` → `ComGetIntProp`，逐字是别人 Fix 161d 的收获；② 扫全仓 417 个 `.bas/.frm/.cls`，`Len(<裸标识符>)` 落在模块级 String 上的位置**只有新用例自己**。新开账 #116：**模块级 `Dim gT As String * 8` + `Len(gT)` 让 C3.exe 直接 `abort()`**（门 #162 的工件也崩，退出码 3 / 0xC0000409、零诊断），而且它弹**前台模态**的 MSVC Debug Error 对话框（`SetErrorMode` 对它不生效）⇒ 探针别碰这一形；仓库原先零个用例用定长串（扫描零命中）所以这洞一直没露头。发码面正 6 / 反 2 条。**门 = Actions #168（run 36353753065，8/8，wall 938s；`test_len_width` 与 `_x86` 在 CI 里各跑到、读数逐字相同）。**
               # 上一档 = 2026-09-28T05:40:00+08:00   # 本轮 = **C29-WS-d 的夹具加固（本地 x64 + x86 各一跑 = 48/48）** = 把 `evtTimer_Timer` 顶上那道"等到前提成立再问"的闸**从 WS-d 的两个拍扩到全部九个到达步**，闸口径一律用字节总量 / 缓冲区读数（`wsB.BytesReceived >= 7`、`gTotB >= 18`、`gTotA >= 4`、`gReqS >= 1`、`gTotS >= 5`、`wsT.BytesReceived >= 5`、`gClsS >= 1`…）而不是事件条数，闸口上限 25 拍；另在第一个到达步后补一条证人 `W0=<事件数>/<字节合计>/<BytesReceived>`，下次再红能从 CI 日志直接分流。驱动这一格的读数：**门 #166（head `1549803` = WS-d）自己 8/8 绿，可前一跑 #165（head `52c85a3`，根本不含本批）却红在 `ctrlwinsock_x86`**，红的是 WS9/WS10/WS12/WS14…一把到达面，证人 `P=49741/49740/58897/1/0/1`（`gArrB=1` 而字节数不对 = 包还没到就问了）。两个头的 `--emit-c` 逐字相同（779 行）⇒ **不是别人那批把夹具改坏了，是同一份夹具在 CI 的 x86 作业上间歇性红** ⇒ 这是 WS-a 那条"事件条数不是判据"教训的变种：**判据也不能钉在拍号上**。另一格已准备好、本推不带（下一个门）：#115 的修法 + `tests/test_len_width.bas`（两架构各真跑 12/12；BASE = 门 #162 工件只红 LW2/LW4/LW11 三条，且 x64 读 8、x86 读 4 = 症状随架构变）；逐字节护栏（30 件工程基对）零收割。**门 = 下面这次 push。**
               # 上一档 = 2026-09-28T05:05:00+08:00   # 本轮 = **C29-WS-d 已出（本地 x64 + x86 各两跑 = 48/48 零红；BASE = 门 #162 的工件，只红 WS46 / WS47，其余 46 条照旧绿）** = Winsock 的发送面：TCP 的 `SendData` 改成**入队即返回 + 随 `FD_WRITE` 分块泵**（64 KiB 一块、上限 8 MiB，溢出报 `Error 10055` 不静默丢），每交一块报一次 `SendProgress`、排空报一次 `SendComplete`。三条读数（全部见 029 §九 C29-WS-d）：① **`FD_WRITE` 是一次性的边、不是电平**（只在“写阻塞转可写”那一下投一次）⇒ 队列非空期间绝不能重挂/摘掉那一位，否则剩下的内容永远敲在队里；② **异步 socket 上 `send` 允许只交一部分**（Winsock2 明写的，不是实现细节）⇒ 部分发送必须在 RTL 里吸收，交给用户判断 = 把数据损坏的口子留在产品里；③ **判据一律问总量、不问块数**（WS-a 那条教训的发送侧版本，报几截由内核缓决定，x64/x86/CI 三处不一样）⇒ 1 MiB 载荷钉四条总量尺。顺带把 WS-b 那条未解的第二条 `FD_ACCEPT` 又试了两种设法（显式重挂原消息号 / 换独立消息号，都在产物里），症状一字不变 ⇒ **不是“恢复方式没找对”**，已写进手册第 9 条。#115 范围量窄：`Len(字面量/局部 String/形参)` 都是对的，错的只有**模块级 String 变量**与 `Len(vbCrLf)`（发 `sizeof` ⇒ x64 读 8、x86 读 4），根因位点定到 `cgen_expr_call_builtin_pre.inc` 的兜底 sizeof 段；另新开账 #116：**模块级 `Dim gT As String * 8` + 赋值 + 打印让 C3.exe 直接 `abort()`**（退出码 3、零诊断，而且会弹 MSVC 调试错误对话框）。发码正 25 / 反 11 条。**门 = 下面这次 push。**
               # 上一档 = 2026-09-28T04:20:00+08:00   # 本轮 = **C29-WS-c 已出（本地 x64 三跑 + x86 两跑 = 41/41 零红；BASE = 门 #161 的工件，六条新针全红其余 35 条照旧绿）** = Winsock 的 `Error` 七参数事件 + 五条出错出口统一走 `vb6_WsFail` + **撤掉 `SO_REUSEADDR`**。三条读数每条都会再咬人（全部见 029 §九 C29-WS-c）：① **`closesocket` 会把 `WSAGetLastError()` 清成 0** ⇒ 任何 Winsock 调用一失败，先把编号抓进局部变量再谈清理（第一版没抓，事件照发、`Number` 是 0，现场零症状）；② **Windows 的 `SO_REUSEADDR` = 允许抢口**：两枚控件绑同一个 UDP 口，第二枚 `bind` 静默成功（wsprobe14 Q2），包只有一份落点 —— 老控件没这坑；撤掉之后撞口当场 10048，`Error` 才有确定性来源，而 Q3/Q4 证明撤它不赔快速重启；③ **把 String 形参赋给模块变量发的是裸指针拷贝**（`gErrD = (*Description);`，不发拷贝不加引用）⇒ 回调返回、调用方一释放原串就是悬垂指针，实测存下来的描述读出 `"WS38="` / `"Y"` 这种别人用过的堆块 —— **另立缺陷 #113**（影响面是所有带 String 参数的回调，不止 Winsock）；本批评据一律当场问长度（记进 Long），要留串就 `Description & ""` 复制走，WS40 钉这条绕法。判据 35→**41**、发码正 22 / 反 11 条（新三条里含那枚七参数原型 `void f(int16_t, BSTR*, int32_t, BSTR, BSTR, int32_t, int16_t*)` —— `ByVal ⇒ 按值、没写 ⇒ 按指针` 在七参数上同样成立，而这条只有 x86 真跑验得出来）。邻居 `ctrlrichtextbox` 89/89、`ctrldatetime` 42/42。**门 = Actions #162（run 36346774776，head `541e17b`，attempt 1，8/8 job 全绿，wall 979s = 16m19s；逐 job：Build 199s / vbp 778s / bas#2 295s / bas#1 294s / asm 81s / smoke 44s / syntax 25s / compile 18s）。**
               # 上一轮 = 2026-09-28T03:30:00+08:00   # 本轮 = **C29-WS-b 已出（本地 x64 三次 + x86 两次 = 35/35 零红；发码正 20 / 反 10 条；邻居 RT 89/89、DT 42/42）** = Winsock 第二档：TCP 的服务端与客户端互为前提那一整轮（`Listen` / `ConnectionRequest` / `Accept` / `Connect` + `Connect` 事件），仍走原生 Winsock2、没添文件。三个形状由读数定（全部见 029 §九 C29-WS-b）：① **监听面与数据面分开存**（`listenSock` + `sock` + 一条 FIFO `pending`），通知归属只能问 `wParam`（实测 = 触发它的 SOCKET；一枚控件最多同时挂两条 socket，都往同一个窗投同一个私有号）；② **`Connect` 同步逐条候选试到通** —— `getaddrinfo("localhost")` 先 `::1` 再 `127.0.0.1`，只试第一条对纯 IPv4 的监听端必然被拒；每条候选用"非阻塞 connect + `select(可写)` + `getsockopt(SO_ERROR)`"判成败，**通了才挂 `WSAAsyncSelect`**（挂着异步选择时 `select()` 问不出东西），副产品是判据可以紧跟在 `Connect` 那行之后；③ **`bytesReceived` 每次通知归零** —— TCP 的 `FD_CLOSE` 会先 drain 一次，留着上一笔计数就凭空多发一次 `DataArrival`（UDP 那一档暴露不出来）。**量到但没收的一格**：同一枚控件受理过一条连接之后，第二条的 `FD_ACCEPT` 在产品里再也送不到（裸码 trace：那扇窗此后一条 `WM_WS_NOTIFY` 都不收），四份探针（wsprobe10..13，把 park 清选择 / `FIONBIO` / 共用 (窗,号) / mask 分两次挂 / accept 在不在处理器内都各自排掉）全部复现不出来 ⇒ 控件数组多客户端与这条一起归 WS-c，没当既有行为发货。**本档第四次踩同一坑**：只接了 `withm.inc` + `cgen_call.cpp` 两码头，漏了第三处 `cgen_expr_member_form_builtin.inc` 的标记表 ⇒ `wsS.Listen` 落回 `vb6_ComCall(vb6_hwnd_wsS, L"Listen", NULL, 0)`，编得过、14 条判据一声不响地红（负控现场留档）。**门 = run #161 绿**（head `8f64562`，8/8 job 全绿，wall **960s = 16m00s**；逐 job：Build 209s / vbp 743s / bas#1 298s / bas#2 284s / asm 77s / smoke 43s / syntax 32s / compile 17s）⇒ GATE_BASELINE 从 #160 挪到 **#161**。
               # 上一轮 = 2026-09-28T02:35:00+08:00   # 本轮 = **C29-WS-a 已出（本地 x64 24/24、x86 24/24、x64 连跑五次 0 红；BASE 7 绿 / 17 红）** = Winsock 第一档：原生 Winsock2 的身份窗 + 状态机 + 属性面 + **UDP 一整轮**（不加载 MSWINSCK.OCX，同 D6）。开工前九轮探针把五条现代口径钉死（全部记 029 §九 C29-WS / C29-WS-a）：通知种类只认消息 `lParam` 低字（async-select 下 `WSAEnumNetworkEvents` 回空）、`FD_CLOSE` 会带着没读完的尾巴到 ⇒ 先 drain 再发 Close、未读数据会被重投 `FD_READ` ⇒ RTL 一次读干净（老控件那条「DataArrival 里必须 GetData」的根因就在这）、一枚双栈 socket 通吃 v4/v6、`getaddrinfo` 本机 `localhost` 是 `::1` 在前 ⇒ 必须逐条试。两条本档新踩的：`sin_port` 少了 `htons` 的症状是「发得出去、收不到、计数乱跳」（第一版 9 条红全在这）；等事件用「泵 N 次 DoEvents」必假红 —— 改成一步一个 Timer tick 才确定。顺带：`Close` 这种无实参方法走的是 `cgen_call.cpp` 那条**语句路**，与带实参那条（`withm.inc`）两码头都得接，只接一头就是本线踩过三次的那声不响（第一版红在 `vb6_ComCall(wsa, L"close")` + `wsa` 未声明）。TCP 的 Listen/Accept/Connect 在 RTL 里只留空壳、发码侧刻意**不接**（接了就是发一条什么都不做的调用）。**门 = 两次**：#159（head `886e5d9`）**红在一格** —— `ctrlwinsock_x86` 的 WS17/WS22，本机 x64 把"AB"+"CD"两笔合并成一次 `FD_READ`（gArrB=3）、CI 的 x86 分成两次（gArrB=4）⇒ **条数是时序不是不变量**，第一版拿一次读数当了身份（与 MV-d 那格"挂钟零头随时段翻脸"同一类错误）；判据改成问内容（`gGot4 = "ABCD"`、关掉之后再取是空串），WS18 的 `gArrA` 从 `=1` 松成 `>=1`，条数降级进 `P=` 证人行 = `f182ecd`，**门 #160 绿**（8/8，wall 893s = 14m53s）⇒ GATE_BASELINE 从 #158 挪到 **#160**。
               # 上一轮 = 2026-09-28T00:55:00+08:00   # 本轮 = **C29-RT-d 已出（本地 x64 89/89、x86 89/89、x64 连跑五次 0 红；BASE 63 绿 / 26 红）** = RichTextBox 的 `Change` / `SelChange` 两条事件，RichTextBox 这一族四格到此收口。两条与文档相反的读数：① `Change` 的码不是 TextBox 那格的 `EN_CHANGE`(768) —— Msftedit 压根不发 768，它发 `EN_UPDATE`(1024)，而且**排在重绘之后**（`Text =` 之后不泵消息就什么都收不到，第一批判据三条全红就是这么来的；VB6 是同步 raise，这条差异进手册）；② `ENM_SELCHANGE` = 0x00080000 **整枚在高 16 位**而 `EM_SETEVENTMASK` 只认 wParam 的低 16 位 ⇒ 高位必须走 lParam 那个指针，六格矩阵里只有这一格发得出来（不开掩码 = 一条不发），写在 `vb6_RTB_Init`。③ 更要紧的是**同一份产物连跑三次，选区通知 1 次走 WM_NOTIFY/1794、2 次走 WM_COMMAND/1815、还有一次一条都不发** ⇒ 两条 arm 都接（一条事件只落一条通道，双发会当场让计数变 2 翻红），而 SelChange 的判据不能用真属性写驱动，改用新的判据专用助手 `vb6_RTB_SimNotify` 各造一条真通知（DT-c / MV-c 同先例）。手法记账：C 探针给的分布与产物不一致，最后是把一次性 `vb6_RTB_TraceCmd` 编进产物自己的 WndProc 打裸码才看清要接几条（用完即删，本批没有它）。踩实一条老账：窗体模块里 `Left(s, n)` 被抢去当窗体的 Left 属性，运行期 0xC0000005（台账 #68 由『读到错值』升级成『会崩』）。判据 79→**89**、发码正 7 / 反 1 条，邻居 `ctrldatetime` 42/42、`ctrlmonthview` 40/40 无变化。全部读数见 029 §九 C29-RT-d。**门 = 下面这次 push。**
               # 上一轮 = 2026-09-27T22:10:00+08:00   # 本轮 = **C29-RT-c 已出（本地两架构 79/79 + BASE 61 绿 / 18 红）** = RichTextBox 的 `TextRTF` / `LoadFile` / `SaveFile` / `Find`（夹具 58→**79 条**、发码正 8 / 反 4 条针）。这一格真正的收获是 `EDITSTREAM` 回调的两条契约，各踩一次才量清：**字节数只走 `*pcb`（返回值非零 = 中止，而且那个值原样落进 `dwError`）** + **流尾要带结尾那个 NUL** —— 两条任一不满足，症状都是"赋值/LoadFile 成功、控件里一个字没有"（与"根本没调"同形）。另一条 cgen 侧的新坑：**标记必须在发实参之前清掉**，否则实参里的字面量被"COM 属性读"捡走（实测 `Find` 第三实参发成 `(-vb6_ComGetIntProp(rt4, L"find"))` ⇒ C2065）。还有 `Find` 的 flags 两边不同位（VB6 1/2 对原生 2/4）⇒ RTL 逐位折算，`RT75/RT76` 钉住。全部读数与理由见 029 §九 C29-RT-c 那一格。**门 = run #157 绿**（head `a3c4bc8`，8/8，wall 853s = 14m13s）⇒ GATE_BASELINE 从 #154 挪到 **#157**。
               # 上一轮 = 2026-09-27T18:52:00+08:00   # 本轮 = **C29-MV-d 已出（本地 40/40 两架构 + BASE 6 条全红）** = MonthView 的 Date 面归到"整天"。上一格那条 `OPEN_LOCAL_DIVERGENCE` 收到因由了，而且**不是** OS 版本差异、是**挂钟差异**：探针 `.build\mcsel6.c`（同一份源码编两份，一份挂 v6 manifest）打出逐字段读数 —— v6 下 `MCM_GETCURSEL` 把**当前挂钟时间**混进负载（发 43894 回 `2020-03-04 18:21:47.128` = 43894.765127，v5 那份回 43894.000000），**日从来没错过** ⇒ 三条红是 `CLng` 把零头**进位**出来的，于是"下午红、上午绿"：CI 那几轮都跑在上午（#151 = 08:32 UTC，零头 0.35 舍掉 ⇒ 绿），本机一下午就红了两天。修法 = 新增 `vb6_MvDaySerial`，读数前抹平四个时间字段，`Value` / `SelStart` / `DateClick` 负载三处一起用；`SelEnd` 顺带**推翻一条旧账**：实测两版都把止端存成**被选中的最后一天**（v5 回 46272.000000、v6 回 46272.999988 = 当日 23:59:59.9），旧注释那句"内部存半开区间、恒多一天"是把 `CLng` 进位后的读数当成了控件里的数 ⇒ 改成只抹时间、不折天（旧写法交出去的是 46270.999988 这种数，按数值比恒差一天）。判据 37→**40**：三条新针一律按数值比不取整（MV38/39/40），另在针之外打一行 `D=` & `CDbl(...)` 当原始证人。读数 **x64 与 x86 各 40/40**、两架构原始行逐字相同（`D=44562/46269/46271`）；BASE = 干净 worktree `.build\wt_pre`(`3f7591d`) 冷编那份，跑同一夹具 **34 绿 / 6 红**（老三条 + 新三条全红，它的原始行 `D=44562.7810300926/46269/46270.9999884259` 就是病灶）。同批用新 C3.exe 复跑邻居：`ctrlrichtextbox` 58/58、`ctrldatetime` 42/42（`W=143/64/95/121` 与台账逐字相同）⇒ Date 面无连带伤害。**门 = run #154 绿**（head `3c92a7a`，8/8 job 全绿，wall 958s = 15m58s，逐 job 见下面 GATE_BASELINE 那一格）⇒ GATE_BASELINE 从 #153 挪到 **#154**。上一格 C29-RT-dep 的门 = **#153 绿**（head `3f849bf`，8/8，wall 773s = 12m53s）。
               # 上一轮 = 2026-09-27T18:15:00+08:00   # 本轮 = **C29-RT-dep 已出（本地）+ MonthView 那格的分歧改口**。用户一句"Msftedit 是系统 dll，会不会有缺失啊"→ 探针 `.build\richcls.c`（x64+x86 各一份）量出：`riched20.dll` 只注册 `RichEdit20W`（问 `RICHEDIT50W` 回 err=1411），`Msftedit.dll` 才注册 `RICHEDIT50W` ⇒ RTL 里那句 `LoadLibrary(riched20)` 兜底是**假出口**（发码的类名是写死的字面量，换 dll 建不出窗口），删掉、换成分两支的 stderr 诊断。两支都用负控验过能红：exe 同目录放一份改名成 `Msftedit.dll` 的 riched20 ⇒ 打"类没注册"那条、夹具 **38 红 / 20 绿**；放一份不能加载的假文件 ⇒ 打"加载不了"那条、同样 38 红。正常机器改前改后都是 **x64 58/58、x86 58/58**。**同时把 MonthView 那格 `OPEN_LOCAL_DIVERGENCE` 改口**（见本文件上面那行）：两枚干净 worktree 冷编出来的 C3.exe 跑同一件夹具，读数与主树逐字相同 ⇒ 不是本地缓存；四组对照（挂/不挂 comctl v6 manifest × 加载/不加载 Msftedit）里 **A≡B**（不挂的两组全精确），挂上的两组三条"跨月"读数一律 +1 ⇒ 开关只有 **comctl v6 激活上下文**；（订正：当时写的"C≡D 逐字相同"过头了 —— 本月那一条 C 报 DIFFERENT、D 报 OK，未解释；"Msftedit 不参与"改由下一轮的 `mcsel6` 立住：那份探针不加载 Msftedit，v6 组照样带零头）`SETCURRENTVIEW` 先挪视图那条修法实测无效、已回退。STATUS 保持 BUSY。**本行的"门 = 下面那次 push"与"MonthView 一行没留"都已被下一轮（C29-MV-d，门 #154）替掉：改的是读回侧。**
               # 上一轮 = 2026-09-27T15:45:00+08:00   # 本轮 = **C29-RT-b 已出**（同一族接上 `Sel*` 格式面：字符四条效果 + 颜色 / 字体名 / 字号，段落对齐 + 三个缩进，共 11 个属性 22 个口）。夹具 `tests\ctrlrichtextbox` 从 32 条接到 **58 条**（RT33-RT58），**x64 与 x86 各 58/58**；BASE（本批之前的二进制，家族级负控）**20 绿 / 38 红**（新针里 16 条当场红，另 10 条是"应为 0 / 应为相等"那类反向针）。发码两面各加 13 / 3 条（新针全部对着真发码数过一遍，含 `vb6_RTB_SetSelFontSize(vb6_hwnd_rt2, 14)` 这条"裸数值不做装箱"）。
               # 这格真正的收获是**三态问法比切格时想的干净**：只问那一位时，`EM_GETCHARFORMAT` 返回的 `dwMask` 里该位被清 = 选区内不一致（本机读数：斜体全一致 `0xFFFFFFFF` / 跨边界 `0xFFFFFFFD`，字符面**返回值本身就等于那张掩码**；段落面同型但不能拿 rc 当掩码）。本项目没有 Null 可回 ⇒ 混合一律按"没有"那一头（`False` / `0` / 空串），针 RT35/RT42/RT47/RT56 钉这条口径。另三条折算与边界也都在判据里：字号原生单位 1/20 磅、对齐 VB6 0/1/2 对原生 `PFA_LEFT=1/CENTER=3/RIGHT=2`（不折算就把"设居中"读成"右对齐"）、悬挂缩进必须 `PFM_OFFSET` 取负（`PFM_OFFSETINDENT` 实测把整段推走，720 → 960）。头里没有 `PARAFORMAT2W` 这个名字（只有 `PARAFORMAT2`）—— 照记忆写 `*2W` 就是六个 C2065。
               # 护栏与归因：逐字节护栏漂移 48 → 57 行，多出来的全在 `ctrlprop\CtrlProp.vbp` 的 `WM_CTLCOLORSTATIC` / `VB6_BackColorSet` 那几行 = **合并进来的别人那批**（SSTabEx 颜色族），按 `vb6_RTB_` / `RICHEDIT` 过滤本批符号 = 0 命中；那份护栏的 BASE 仍是 C29-1a 的 `b21_C3.exe`，下次先换 BASE。**门 = run #151 绿**（head `45ade37` = 先合进别人的 `4623832`(Console) + `5476632`(COM ByRef/As New) 再叠本批，**8/8 job 全绿**：Build 200s / vbp 680s / bas#2 329s / bas#1 262s / asm 80s / smoke 51s / syntax 24s / compile 13s ⇒ GATE_BASELINE 挪到 **#151**。用例级那一层这次取不到日志（GitHub 的 jobs/{id}/logs 连着 502/504），所以本行只记 Jobs API 给的任务级结论，不编用例行。） 写这段之前曾误把 run #150 当成过本批的门：核对后 #150 的 head 是 `ab6d287`（= 别人那批：Console 内置对象 + COM ByRef/`As New` 两条修复的合并），而本批的 `261d51a` 当时还没推上去 —— 那条门里跑的 `rt_emitc_shape` 只有 RT-a 的 15 条正针，**不覆盖本批**。（发现过程与订正见下面 GATE_BASELINE 段那条注记。）
               # 归因教训一条（写死）：门号只认 `git ls-remote github dev` 与 `git log github/dev..HEAD` 两个都对上的那次 run —— 短 sha 前缀撞上别人刚推的 merge commit 就会像这样把别人的绿读成自己的。
               # 上一轮 = 2026-09-27T15:05:00+08:00   # 本轮 = **C29-RT-a 已出**（RichTextBox → 原生 `RICHEDIT50W`，不加载 RICHTX32.OCX）：新 RTL 文件 `vb6forms_richtextbox.c`（`.rc` **219**、四处登记齐）+ cgen 五处（枚举 / 三条映射 / 创建样式两路 / 读写表 / 设计期 `vb6_RTB_Init`）+ `vb6_ComCtl_Init` 里补一次 `LoadLibraryW(Msftedit.dll)`（这枚类不在 comctl32 的 `ICC_*` 体系里，是它与 DT/MV 唯一的通道差别）。夹具 `tests\ctrlrichtextbox` 四枚控件 **32 条判据：x64 与 x86 各 32/32**；BASE（`c298c_base_C3.exe`）同夹具 **22 红 / 10 绿**（留绿的十条全是"读回 0 / 空 / 反向判断"那类）。发码两面：`rt_emitc_shape` 15 条 + `rt_emitc_no_com_fallback` 5 条（含断 `vb6_RTB_SetScrollBars` **根本不许存在**）。代码 `2ba2392`，合并别人那批（`7ab2406`，SSTabEx 颜色族 + dbgdlg 夹具）后 `37f5404` 推 dev ⇒ **门 = run #148 绿**（head `37f5404`，8/8 job 全绿，wall **916s = 15m16s**；逐 job：Build 167s / vbp 744s / bas#2 219s / bas#1 203s / asm 90s / smoke 45s / syntax 27s / compile 19s。CI 的 vbp 那一格：`[VBP] ctrlrichtextbox ... PASS` 与 `ctrlrichtextbox_x86 ... PASS` + `rt_emitc_shape` / `rt_emitc_no_com_fallback` 两条，`Results: PASS=108 FAIL=0 SKIP=1` ⇒ 比上一轮多出的四条就是本批的）。  
               # 第三轮 C 探针（`.build\rtprobe3.c`）把切格时定的口径翻掉三条，判据形状是订正之后决定的：① **滚动条样式位会被控件自己抹掉**（创建时给了 `WS_VSCROLL`，空文本下 `GWL_STYLE` 读回 0，灌 60 行又自己回来）⇒ cgen 一并挂 `ES_DISABLENOSCROLL` 才谈得上"读回设计期那一位"，而事后 `SetWindowLong` 只有外观、量程停在 0..100（真高 1281）⇒ 那四位**没有写口**，判据另配控件侧证人 `VScrollRange`/`HScrollRange`；② 原生默认上限实测就是 **32767**（不是天文数字）⇒ "不限"只折 `==32767`，而 **`MaxLength` 管不住 `Text =` 赋值**（控件把上限抬到文本长度，`EM_REPLACESEL` 同样穿过去）；③ **`EM_SETTARGETDEVICE` 的 lParam 才是目标 DC** —— NULL = 折到本窗客户区宽（= VB6 的 True）、传 `GetDC(控件)` = 整屏宽（实际不折），而原生默认**不折**、与 VB6 相反 ⇒ 设计期没写也下发一次 True；`EM_SET/GETWRAPMODE` 在这枚上问不出也设不动（mode 恒读 0）⇒ 只能自存。  
               # 顺带修一条同族既有缺陷（不修就一直假绿）：`vb6_Get/SetBorderStyle` 只认类名 `"Edit"`，其余类落到"存窗口属性"兜底分支，而 **`SetPropW(hw, name, NULL)` 等于删属性** ⇒ `GetPropW` 回 NULL ⇒ 恒读默认 1，`BorderStyle = None` 对任何非 Edit / 非 Static 控件都设不上。本批只把 `RICHEDIT50W` 接进 Edit 那条 `WS_EX_CLIENTEDGE` 路，**其余类型的同一条缺陷新记一条账**（见 029 §九）。另记一条工具账：`.build\c29_emitc_guard.py` 的 BASE 还是 C29-1a 时代的 `b21_C3.exe`，二十来批的漂移全攒在里面（今天读来 48 行、集中在 Implements / ctor / cc_dll / Command 这些语言层工程，按 `vb6_RTB_`/`vb6_MV_`/`vb6_DTP_`/`RICHEDIT` 过滤 = 0 命中 ⇒ 与控件线无关）；下次动它先把 BASE 换成"本批之前那一次的 C3.exe 存档"。本线接下来：RT-b（`Sel*` 格式面）→ RT-c（`TextRTF` + `LoadFile`/`SaveFile` + `Find`，`EDITSTREAM` 回调四参）→ RT-d（`Change` 走 `WM_COMMAND`、`SelChange` 走 `WM_NOTIFY`，两条通道不同）。STATUS 保持 BUSY。
               # 上一轮 = 2026-09-27T12:25:00+08:00   # 本轮收口 = **C29-MV-c 过门 ⇒ MonthView 这一族收口**（代码 `8b2fe83`；DateClick = 原生 `MCN_SELCHANGE(-749)`，夹具 33→37 条、x64/x86 各 37/37、BASE 32 红 / 5 绿）。**门 = Actions run #143 绿**（head `a5bd656`，8/8 job 全绿，wall **936s = 15m36s**；逐 job：Build 223s / vbp 708s / bas#1 243s / bas#2 213s / asm 79s / smoke 44s / syntax 24s / compile 19s）⇒ GATE_BASELINE 挪到 **#142 → #143**。本轮真正的收获是那条只有 x86 暴露的形状错误：**`ByVal` 参数按值传**（规则在 `cgen_decl_proc.cpp:94`；`Form_MouseDown` 那批 `int16_t*/float*` 是因为 VB6 签名里没写 ByVal）。派发照抄指针形状时 x64 读数照旧正确、x86 事件参数读回 0 —— 又一次坐实“门只测默认架构 ⇒ 布局/ABI 类改动必须 x86 真编译真跑”那条纪律。**DTPicker（DT-a/b/c，#136/#139/#140）与 MonthView（MV-a/b/c，#141/#142/#143）两族到此全部收口**，本轮四格连出（DT-c + MV-a/b/c）、四扇门各自绿。本线接下来：`ai/内置控件/` 那张表里还剩 Slider / TabStrip / UpDown / RichTextBox / MaskEdBox / MSHFlexGrid / DataGrid / Adodc / MSComm / Winsock 没落原生；STATUS 保持 BUSY。
               # 上一轮 = 2026-09-27T12:20:00+08:00   # 本轮 = **ai/029 三格连出**：① **C29-MV-a 已出并过门 #141**（`c95ac75`，8/8，wall 905s）；② **C29-MV-b 已出并过门 #142**（`f0c8875` 那一批里，8/8，wall **615s = 10m15s**；逐 job：Build 153s / vbp 456s / bas#2 265s / bas#1 252s / asm 92s / smoke 40s / syntax 24s / compile 24s）⇒ GATE_BASELINE 挪到 **#142**；③ **C29-MV-c 已出并本地全绿**（DateClick，夹具 33→**37 条**、x64/x86 各 37/37、BASE 32 红 / 5 绿）⇒ **MonthView 这一族到此收口**。 MV-c 这格抓到一条只有 x86 才暴露的形状错误：**`ByVal` 参数按值传**（规则在 `cgen_decl_proc.cpp:94`，ByRef 才是指针 —— `Form_MouseDown` 那批 int16_t*/float* 是因为 VB6 签名里没写 ByVal）。派发头一版照抄了旁边的指针形状，症状是 **x64 读数照旧正确、x86 当场错值**（事件参数读回 0），两架构真跑才拦得下来 —— 又一次坐实"门只测默认架构 ⇒ 布局/ABI 类改动必须 x86 真编译真跑"那条纪律。另外量出：原生对程序化改选**不发** `MCN_SELCHANGE`（MV36，与 DT-c 的 Change 同型）。 本线接下来：ai/029 §一 那张表里还剩 Slider / TabStrip / RichTextBox / MSHFlexGrid / MaskEdBox / Winsock / MSComm / Adodc / DataGrid 等没落原生；STATUS 保持 BUSY。
               # 上一轮 = 2026-09-27T12:10:00+08:00   # 本轮 = **ai/029 两格**：① **C29-MV-a 已出并过门**（代码 `c95ac75`，夹具 22 条 x64/x86 各 22/22，BASE 19 红 / 3 绿）⇒ **门 #141 绿**（head `65dbbb1`，8/8，wall **905s = 15m05s**；逐 job：Build 191s / vbp 709s / bas#2 191s / bas#1 185s / asm 80s / smoke 42s / syntax 22s / compile 19s）⇒ GATE_BASELINE 挪到 #141。MonthView 的四条实测口径见 029 本格，最值钱的一条：**多月平铺不是开关而是窗口多大**（判据因此问 `MCM_GETCALENDARCOUNT`，不查我们自己的表）。② **C29-MV-b 已出并本地全绿**（Date 值面 Value / SelStart / SelEnd，夹具 22→**33 条**、x64/x86 各 33/33、BASE 28 红 / 5 绿）= 先拿一枚 C 探针问原生再写判据，量出三条：**CURSEL 与 SELRANGE 两张表互斥**（挂不挂 `MCS_MULTISELECT` 决定哪张能用，两条针一起钉）、`SETSELRANGE` 收闭区间而 `GETSELRANGE` 吐内部半开 ⇒ `GetSelEnd` 折回一天（针盯着，换机器不成立就红）、**范围只能在当前显示的那一个自然月里** ⇒ 判据一律以"本月 1 号"为基准推，不写死日期（写死会在 CI 未知日期上月初/月末随机红，同 022 账 #79 那条纪律）。DT-b 那两个日期换算助手挪进 `vb6forms_internal.h` 共享（MonthView 是第二个用户）。下一格 = **C29-MV-c**（`MCN_SELCHANGE(-749)` → VB6 的 `DateClick(DateSelected As Date)`，这条**带参数**、与 DT-c 那三条无参的不同形）。STATUS 保持 BUSY。
               # 上一轮 = 2026-09-27T11:20:00+08:00   # 本轮 = **ai/029 两格**：① **C29-DT-c 已出并过门**（代码 `adc69ad`）= DTPicker 的 Change / DropDown / CloseUp 接进父窗 `WM_NOTIFY`（-759 / -754 / -753，逐条对过本机 `CommCtrl.h`）；三条处理器在 VB6 **都没有参数** ⇒ 只认来源与码值、不按负载造对象。**门 = Actions run #140 绿**（head `adc69ad`，8/8，wall **834s = 13m54s**；逐 job：Build 131s / vbp 697s / bas#1 240s / bas#2 201s / asm 77s / smoke 52s / syntax 23s / compile 17s）⇒ GATE_BASELINE 挪到 #140。夹具 36→**42 条**、x64/x86 各 42/42；两条量出来的口径写进 029 本格：原生对**程序化** `Value` 赋值**不发** `DTN_DATETIMECHANGE`（照原生、不伪造，`DT41` 是这条口径的哨兵）、零参判据方法必须写成调用形 `dt1.SimChange()`（不带括号在语义层是属性读 ⇒ 整条语句发码为零，实测踩过）。**DTPicker 这一族到此收口**（`CallbackKeyDown` / `FutureDate` 两条刻意不做，理由在 029）。② **C29-MV-a 在跑** = MonthView 走原生 SysMonthCal32（新 RTL 文件 `vb6forms_monthview.c`，四处登记一次到位）。四条量出来的：多月平铺**不是开关而是窗口多大**（判据问 `MCM_GETCALENDARCOUNT`）、`ShowToday` 与 `MCS_NOTODAY` 是反的、`MaxSelCount` 真往返过控件而**没挂 MCS_MULTISELECT 的写不进去**（这比 GWL_STYLE 读数硬）、以及**与 DTPicker 相反**：MCS_ 三位运行期写有效（两枚控件的边界各量各的，不许外推）。夹具 `tests\ctrlmonthview` 22 条、x64/x86 各 22/22（原始读数两架构逐字相同），负控 = DT/MV 之前的二进制 19 红 / 3 绿。STATUS 保持 BUSY，下一格 = **C29-MV-b**（Date 三格）。
               # 上一轮 = 2026-09-27T10:20:00+08:00   # 本轮 = **ai/029 C29-DT-b 已出并过门**（代码 `082d7c4`）= DTPicker 的 Date 值面：`Value` / `MinDate` / `MaxDate` 走**裸 double 序列号**、不经任何装箱（`Dim d As Date` 在 C3 里就是 `double`，所以 `vb6_DTP_SetValue(hwnd, dReq)` 直接收变量）。本仓库第一次有控件属性的类型登记成 `Date` ⇒ `controlPropType` 那一格先量过再登记（漏登记的后果 = SSTab1.Tab 那族 AV）。**两条实测口径**：① 越出 MinDate/MaxDate 的赋值被控件**拒绝、值保持原样**，不是钳到边界；② 未勾（`GDT_NONE`）态下 `DTM_GETSYSTEMTIME` 照样回填一个内部日期（本机 36494）⇒ `GetValue` 必须认返回标志才回 0，而且『取消再勾回』要我们自己记住原来那天（存窗口属性），否则勾回来是控件的内部日期。读数：夹具 `tests\ctrldatetime` 24→**36 条**，x64 与 x86 各 36/36（`W=143/64/95/121` 两架构逐字相同）；负控 = DT-a 之前的二进制跑同一件夹具 32 红。发码 9 + 4 条针逐条对过真实 `DtfForm.c`。**门 = Actions run #139 绿**（head `082d7c4`，8/8 job 全绿，wall **944s = 15m44s**；逐 job：Build 196s / vbp 744s / bas#2 252s / bas#1 192s / asm 79s / smoke 49s / syntax 21s / compile 18s）⇒ GATE_BASELINE 挪到 #139。本轮同时把 **C29-DT-c（DTN_* 三条事件）出到本地全绿**：Change / DropDown / CloseUp 接进父窗 `WM_NOTIFY`（码值 -759 / -754 / -753，逐条对过本机 SDK 的 `CommCtrl.h`），三条处理器**都没有参数**（与 ListView/TreeView 那两条造对象的差别），夹具 36→42 条、x64/x86 各 42/42；两条新量：原生对**程序化** `Value` 赋值**不发** `DTN_DATETIMECHANGE`（DT41 钉这条口径，做 VB6 齐平时它必须翻红）、判据方法必须写 `dt1.SimChange()`（不带括号在语义层是属性读 ⇒ 发码一条都不发，实测踩过）。STATUS 保持 BUSY，下一格 = **C29-MV**（MonthView → SysMonthCal32）。
               # 上一轮 = 2026-09-27T09:25:00+08:00   # 本轮 = **控件线重开（用户点名 DTPicker / MonthView 没做）**，新格 **C29-DT-a 已出并本地全绿**：DTPicker 走原生 SysDateTimePick32（D6：不碰 MSCOMCT2.OCX —— 32 位 inproc 进不了 x64，而这枚今天整件扣在"第三方 OCX 按 COM 后期绑定"那一组 ⇒ 属性读回空、写进去静默丢、退出码照旧 0）。改动六处 + 新 RTL 文件 `vb6forms_dtpicker.c`。**两条实测口径**：① `DTS_SHOWNONE`/`DTS_UPDOWN` 是子窗口的创建参数，事后写 GWL_STYLE 会被控件抹回去 ⇒ 三条样式属性一律从 .frm 立进创建参数（重建窗口那条路今天走不通：setter 只拿到 HWND 值）；② 光读样式位会自洽地假绿（SDK 的 DTS_TIMEFORMAT=0x9 自带 bit0=UPDOWN），所以格式类判据一律配一条 `DTM_GETIDEALSIZE` 控件侧宽度 —— 它顺手翻出我自己的一个错判：运行期切 Format **有效**（95→121，我先前拿 143 当基线，那枚带着复选框当然更宽），运行期改 CheckBox 无效。读数：夹具 `tests\ctrldatetime` 四枚控件、x64 与 x86 各 **24/24**（宽度两架构逐字相同）；负控 = 修复前二进制跑同一件夹具 **19 红 / 2 绿**。发码两面钉（形状 6 条 + 反面 3 条）。**新记一条仓库事实**：新 RTL 文件要登记**四处** —— .rc / CMakeLists / rtl_embedded 那三处只管解包，`driver_link.cpp` 的 `opts.sourceFiles` 才管编译；少第四处的症状是解包成功、发码正确、18 条 `LNK2019: vb6_DTP_*`（本次实撞）。**门 = Actions run #136 绿**（head `bcd47c4`，8/8 job 全绿，wall 715s = 11m55s、vbp 511s；CI vbp 日志里 `ctrldatetime` 与 `ctrldatetime_x86` 各自 PASS、那格 FAIL=0）⇒ GATE_BASELINE 挪到 #136。本线还有三格排队：DT-b（Date 值面，要先量「控件属性 getter 返回 double」这条本仓库没人走过的类型路）、DT-c（DTN_* 三条事件）、C29-MV（MonthView → SysMonthCal32）。STATUS 保持 BUSY —— 这条线还在跑。
               # 上一轮 = 2026-09-27T07:05:00+08:00   # 本轮 = **只读复跑、零代码改动**：合并 `github/dev@ea681f3`（另一位作者的 gate/p20-44 那批）之后的 **dev 级复跑 #132 = completed/success**（head `ea681f3`，**8/8 job 全绿**，wall **751s = 12m31s**；逐 job：Build 210s / vbp 537s / bas#1 256s / bas#2 241s / asm 103s / smoke 40s / syntax 37s / compile 24s）。核过 `git merge-base --is-ancestor 56d5325 ea681f3` ⇒ **#132 的头里含本线 C29-M 那两笔**（`810984f` + `56d5325`），所以这一跑读的是"我的那一格与别人那批同树共存、无新增失败"。我随后推的 `3bd61a4` 相对 `ea681f3` **只差两份 .md**（`git diff --stat` = 2 files changed, +14/-5）⇒ 按"没有源码级变动就不占 Actions"的口径把被它触发的 **#133 取消**（HTTP 202）。STATUS 保持 ALL_DONE，本线批次表仍为空；欠账照旧（#91 / #82 / #83 / #85 / #88 / #76 / #77 / #79 / #64 / #68）。
               # 上一轮 = 2026-09-27T06:45:00+08:00   # 本轮收线 = **ai/029 C29-M 已出并过门**（判定 `810984f` + 修判据 `56d5325`）= 产物清单的让位判据从"给没给 .res"改成"**那份 .res 里真有 #1 清单**"（新文件 `src\driver\res_inventory.{hpp,cpp}` 走 RESFMT 资源目录；`name=2` ISOLATED 不算自带；解析不出结论时**保守让位** —— 两份 #1 会让加载器直接报错）。判据落在产物资源面（`Test-ProductManifest` 数清单份数 + 核特征串）：自带无清单 .res 的工程 = 1 份内置那份（BASE 编译器编同一件 = **0 份** ⇒ 假针能红）、BalloonTooltips = 1 份且是用户那份（`dpiAware` 在）；夹具 `tests\ctrlmanifest` x64 与 x86 各 3/3；41 件 `--emit-c` 护栏与上一格同一读数 ⇒ 发码没动。门 **#130 红在判据自身**（`Test-GuiVbp` 的产物在 `output\<用例名>\` 下，我按 `$OutDir\x.exe` 断，且那条针登记了两遍），修完门 **#131 绿**（head `56d5325`，8/8，wall 779s = 12m59s、vbp 607s）。**本线批次表到此清空**（1a/1b/3/5a/5b/5c/7/8a/8b/8c/9/9b/T/OLE/V6/M 全出）；新记一条待修账 **#91**（`Declare` 的 String/ByRef 出参两条读法都不通 ⇒ 运行期"版本/路径"类判据现在写不出来）。
               # 上一轮 = 2026-09-27T06:20:00+08:00   # 本轮 = **ai/029 C29-M（产物清单让位判据精确化）在跑****：① **C29-5c**（Toolbar 的 `ButtonClick` 走 WM_COMMAND、`ButtonMenuClick` 走 WM_NOTIFY 的 TBN_DROPDOWN(-710)，`9d32c36`）→ 门 **#126 绿**（8/8，wall 881s = 14m41s）；夹具 27→34 条、x64 与 x86 各 34/34，两条负控量过红在哪几条。② **C29-V6**（`b2674bc`）= 编译期把 comctl 的**声明面**统一到 v6（`C3_V6_VERSION_DEFS = /D_WIN32_IE=0x0600`，一处定三处用：普通路编译档 / 增量路编译档 / 增量缓存键）→ 门 **#128 绿**（8/8，wall 894s = 14m54s，vbp 675s）。V6 实测**行为影响为零**：五枚本线没碰过的 comctl 夹具拿 BASE/NEW 两个编译器各编各跑、输出逐字节相同，41 件 `--emit-c` 护栏读数不变。**推翻了自己 5c 那格的归因**：那两条"向控件现问 idCommand 问不出"跟 v5/v6 分家无关（v6 声明下 `sizeof(TBBUTTONINFOW)` 44→48 已证宏生效、重问读数逐字不变）；另一条负控：5a 的 `TB_SETBUTTONINFOW` 启用位补偿拿掉就 TB19 翻红 ⇒ 不是死代码。用户随后定了下一条口径："manifest 要兼容用户 —— 自带清单就用他的，否则用内置的"，实测现状是 `driver_link.cpp:763` 只看 `userResFile_.empty()`（给了 .res 就让位，不看那份 .res 里有没有清单）⇒ 拿 `VBMANLIB.RES`（17KB、零清单串）那类工程会**静默退回 v5**；C29-M 就此重裁成"`resFileHasManifest()` 精确让位 + 产物资源面判据"。本线台账里的"待做只剩 C29-M"照此改写。
               # 上一轮 = 2026-09-27T04:20:00+08:00   # 本轮 = **ai/029 两格连出**：① **C29-5c 已收**（Toolbar 的 `ButtonClick` 走 WM_COMMAND、`ButtonMenuClick` 走 WM_NOTIFY 的 TBN_DROPDOWN(-710)，`9d32c36`）→ 门 **#126 绿**（head `9d32c36`，8/8，wall 881s = 14m41s，vbp 675s）；夹具 27→**34 条**、x64 与 x86 各 34/34，两条负控量过红在哪几条（摘处理器 ⇒ 只红 TB28/29/31；Style 5 改 0 ⇒ 只红 TB32/33）。② **C29-V6 在跑**（用户口径"统一 v6，2026 年了还搞 v5 干毛"）：`msvc_driver.hpp` 新增 `C3_V6_VERSION_DEFS = /D_WIN32_IE=0x0600`，用在普通路编译档、增量路编译档、**增量缓存键** `flagsKeyFor` 三处；刻意不动 `WINVER`/`_WIN32_WINNT`（SDK 默认更新，写死会关掉 RTL 在用的声明）。读数：**行为影响 = 零** —— 五枚本线没碰过的 comctl 夹具拿 BASE/NEW 两个编译器各编各跑，输出**逐字节相同**（statusbar 42 / listview 20 / sstab 40 / imagelist 18 / progress 12 行），41 件 `--emit-c` 护栏与 5c 同一读数（发码一个字节没动），自己两枚 34/33 全绿。**推翻了自己上一格的归因**：5c 把"向控件现问 idCommand 问不出"记在"声明面 v5 / 运行面 v6 分家"账上，统一到 v6（`sizeof(TBBUTTONINFOW)` 44→48，证宏真生效）之后重问，两条路读数**逐字不变** ⇒ 那条 GET 压根不读回 idCommand，与版本无关；另一条负控：拿掉 5a 的"建完补 `TB_SETBUTTONINFOW` 打启用位"补偿 ⇒ **TB19 立刻翻红**，那条补偿不是死代码。门与收线读数随后补。
               # 上一轮 = 2026-09-27T03:25:00+08:00   # 本轮 = **ai/029 C29-5c（Toolbar 的两条按钮事件）在跑**（`ButtonClick` 走 WM_COMMAND、`ButtonMenuClick` 走 WM_NOTIFY 的 TBN_DROPDOWN(-710)；门与收线读数随后补）。
               # 上一轮 = 2026-09-27T03:10:00+08:00   # 本轮 = **ai/029 C29-8c 已出并过门**（代码 `9b42876`）= TreeView 的三条事件（`NodeClick` / `Expand` / `Collapse`）接进父窗 `WM_NOTIFY` 通道：RTL 只交 `vb6_TreeView_NotifyNodeIndex` + `NotifyExpanded`（action 是**三态**字段，按布尔猜会让一次 NodeClick 顺手把 Expand 也点掉）+ 两条 `Sim*` 判据助手，派发侧用 8b 就有的 `NodeAt` 造对象 ⇒ 处理器参数是**对象**。夹具 `tests\ctrltreeview` 27→**33 条**、x64 与 x86 各 33/33；两条负控各自量过红在哪条（Sim 下标换成不存在的 ⇒ **只红 TV28/TV29**；摘掉 `tv1_Expand` ⇒ 红 TV33/TV31/TV32 而 `NodeClick` 那三条照绿）。**判据触发点第三条铁律**：`Sim*` 必须放 **Timer**（`Form_Load` 被 block-events 拦、`Form_Activate` 在无头会话里永不来），且事件读数**只能问增量** —— 真控件在 `Expanded=True`/`EnsureVisible` 时自己就发通知（那不是缺陷，TV33 钉它）。**门 = Actions run #124 绿**（head `9b42876`，8/8 job 全绿，wall 917s = 15m17s、vbp 688s）⇒ #122 那格欠的"vbp 412→676 涨幅待核"**判为 runner 抖动、收掉**。路上把 #89 的口径改了一处：**Toolbar 的 ButtonClick 不在这条通道上**（按钮 `idCommand` 就是 1 基序号 ⇒ 走 `WM_COMMAND`，本次没验）。下一格 = **Toolbar 的 `ButtonClick`/`ButtonMenuClick`** + **C29-M manifest**。
               # 上一轮 = 2026-09-27T02:45:00+08:00   # 本轮 = **ai/029 C29-8c（TreeView 的三条事件派发）在跑**（NodeClick / Expand / Collapse 走父窗 WM_NOTIFY；门与收线读数随后补）。路上把 #89 的口径改了一处：**Toolbar 的 ButtonClick 不在这条通道上**（按钮 idCommand 就是 1 基序号，预计走 WM_COMMAND，本次没验），所以 #89 拆成 8c(TreeView 三事件) + 之后一格(Toolbar)。
               # 上一轮 = 2026-09-26T22:35:00+08:00   # 本轮 = **ai/029 两格连出**：① **C29-8b TreeView 的 Nodes/Node**（`3f1c4ab`）→ 门 **#120 绿**（8/8，wall 619s）；② **C29-5b Toolbar 的 Buttons/Button**（`ee6d769`）→ 门 **#122 绿**（8/8，wall 876s、vbp 676s 对照 #120 的 412s，涨幅本批解释不了，下一格再核）。两格都**复用** C29-3 那套真 IDispatch 成员对象机制（memberobj 只加枚举值与名字表，cgen 各动四处、零链式特例），5a 那条只改道 `Buttons.Count` 的特例钩子退休。8b 顺手拆掉 `wrapToBSTR` 两条未锚定 find（症状：返回 String 的项目函数被打成 `vb6_CStrLong(vb6_TF(...))` = 把 BSTR 指针当数字打）；5b 量到 `TB_ADDBUTTONSW` 不吃调用方给的 `fsState`（⇒ `Button.Enabled` 默认会是 False，建完补一条 `TB_SETBUTTONINFOW`）并清掉 `TB_RESET` 那条根本不在 SDK 头里的残留。新账 **#88**（比较左值的 COM 字符串成员被按数值解包，未修）；**#85** 复测多一个受害者（`CStr(tv1.Nodes.Count)` 打 0）。本线接下来 = **事件面那一格**（`ButtonClick`/`NodeClick`/`Expand`/`Collapse` 共用 WM_NOTIFY 通道）与 **C29-M manifest**。
               # 上一轮 = 2026-09-26T21:40:00+08:00   # 本轮 = **ai/029 C29-5b（Toolbar 的 Buttons/Button 成员对象面）在跑**（门与收线读数随后补）。
               # 上一轮 = 2026-09-26T21:25:00+08:00   # 本轮 = **ai/029 C29-8b 已出并过门**：TreeView 的 `Nodes` 集合与 `Node` 对象（复用 `vb6forms_memberobj.c` 那一族真 IDispatch 机制，cgen 只动四处、零链式特例），`tests\ctrltreeview` 夹具 12→27 条读数、x64 与 x86 都 27/27，两条负控各自实测过红在哪条 ⇒ 门 **#120 绿**（head `3f1c4ab`，8/8，wall 619s = 10m19s，vbp 412s）。同批顺手拆掉一条同族既有缺陷：`wrapToBSTR` 里两条**未锚定** `find("vb6_ComGetIntProp(")`（Fix 100a/100b/100d 那族的漏网）—— 返回 String 的项目函数只要实参含 COM 整数读就被发成 `vb6_CStrLong(vb6_TF(...))`，把 BSTR 指针当数字打（实测 `TV13=-998717560`）；41 件工程逐字节护栏 = 只有本批夹具差异（`.build/c298b_emitc_guard.py`，BASE 从 HEAD `9c78efc` 重建）。新账 **#88**（比较左值是 COM 字符串成员时按数值解包，`If nd.Text = "字面量"` 恒 False，夹具已绕开）；并给 **#85** 补一条复测：`CStr(tv1.Nodes.Count)` 打 0（Nodes 是第三个受害者）。下一格 = **C29-5b Toolbar 的 `Buttons` 对象面**（同一族机制，照 8b 那套接法复用）。
               # 上一轮 = 2026-09-26T21:05:00+08:00   # 本轮 = **ai/029 C29-8b（TreeView 的 Nodes/Node）在跑**：复用 vb6forms_memberobj.c 那套真 IDispatch 机制接上 Nodes 集合与 Node 对象（Add 五种 relationship / 下标与 Key 两条取法 / Text-Key-Tag-Checked-Expanded / Parent-Child-Children-Next-Previous-Root / Remove 带走子树 / Clear / For Each / EnsureVisible），夹具 tests\ctrltreeview 从 12 条扩到 27 条、x64 与 x86 都 27/27，逐字节护栏 41 件里只有夹具自己差异（GUARD OK），两条负控各自实测过红在哪条。顺手拆掉 wrapToBSTR 里那两条未锚定 find（与 Fix 100a/100b/100d 同族、被漏掉的两条）—— 症状：返回 String 的项目函数只要实参含 COM 整数读就被发成 vb6_CStrLong(vb6_TF(...))，把 BSTR 指针当数字打。门与收线读数见下面补的那一行。
               # 上一轮 = 2026-09-26T17:20:00+08:00   # 本轮 = **ai/029 两格**：① **C29-9b 收线**（CommonDialog 真弹框 + 起窗自关探针，`a561ddf`+`b4f1533`）→ 门 **#101 绿**（#99 那轮红在 `frmevents`，根因**是窗口级联偏移不是探针逻辑**，见 029 §九 与账 #79）；② **C29-8a TreeView 标量属性面**（`2339419`）→ 门 **#103 绿**（wall 10m27s）；③ **C29-5a Toolbar 的窗口 + 标量属性 + 设计期按钮**（`12394c2` + 合并 `a0546cd`）→ 门 **#110 绿**（wall 14m14s）。**用户定调：C29-4（WM_NOTIFY 通道 + StatusBar）移交另一位作者**，移交前三条前置读数 + 那条没拍下来的分叉写在 029 §九（探针留 `.build\sbprobe\`）。合并带进 **C29-3 的成员对象机制**（`vb6forms_memberobj.c`）与 **C29-7 ListView** ⇒ **C29-5b / C29-8b 的前置已清**（新账 #86）。本轮新记两条量出来的缺陷：#82（未登记属性按数值读会漏裸指针）、#85（`CStr(集合成员)` 绕开改道钩子）。 ⇒ **补一条读数纪律**：本线出货之后 dev 上又落了别人的两笔（StatusBar 事件面），其合并后跑 **#112 / #113 都在 `Tests (vbp)` 红**（#113 的头里我的部分只有合并 + 纯文档）。所以 GATE_BASELINE 仍停 **#110**（绿）；**深夜补齐**：那条红由 `dc4d008` 修掉，门 **#117 = success**，根因是**登记面错配**（`lv_emitc_events` 上挂着本该属于 StatusBar 的期望串，#116/#112/#113 红在同一条 `Tests (vbp)`） ⇒ 纪律：三人同改 `run_tests.ps1` 时，红的第一嫌疑是"谁的 needle 挂到了别人的工程上"，合并不止查重复登记，还要逐条核对期望串属于哪件工程；合并后本地复跑本线两枚夹具仍各 12/12，红的归因与"不代修别人半成品"的处置写在 029 §九；想查具体哪根针取不到 —— 作业日志端点这几轮一路 502/504，本地整组 vbp 是 74 PASS / 16 FAIL，其中 13 条的症状是 "no probe / no reader"（`COMACT`/`DISPATCH`/`TLB-*`/`VBP-DLL`/`ENC` 那几组，要靠外部探针与 `--keep-for-debug` 产物），而它们在 CI 同一 job 里绿（#110 的 8/8 就含这些组）⇒ 本地环境性红；但**合并前那次本地跑到 61 条我就停了，没拿到这些组的对照**，所以不能说"本地全绿"。
               # 上一轮 = 2026-09-26T14:43:56+08:00：**ai/029 控件线两格**（C29-9b 收线 + C29-8a TreeView，其原文保留在下面）：① C29-9b 收线（真弹框 + 起窗自关探针，`a561ddf` + 挪位处置 `b4f1533`）→ 门 **#101 绿**（#99 那轮红在 `frmevents`，根因是**窗口级联偏移**不是探针逻辑，见 029 §九 与本表 #79）；② **C29-8a TreeView 标量属性面**（新 RTL 文件 `vb6forms_treeview.c` + cgen 三面，`2339419`）→ 门 **#103 绿**，**wall 627s = 10m27s**、vbp 492s（基线 #96 = 730s/544s，本批还往 vbp 组加了 4 条用例）。TreeView 的 `Nodes`/`Node`/事件那半刻意留给 **C29-8b**：它吃 C29-3 的成员对象机制，而 C29-3/C29-7 在另一位作者手上（用户口径"listview 还没写完"）。新账两条 **#82**（未登记的控件属性按数值读会漏裸指针）/ **#83**（容器子控件那条第二创建路缺类型样式与设计期发射，ProgressBar/StatusBar/SSTab/三控件同样缺）。
               # 上一轮 = 2026-09-26T12:13:11+08:00：**ai/030 T30-B 接进 GA**：ci.yml 的回归步骤加 `-Incremental`（`c3739b5`）→ 门 **#94 红**（`test_delegate_x86` LNK1120：增量那条链接路直接叫 link.exe，没配 /MT 的 CRT）→ 修 `39ecaeb`（x86 链接补 `/NODEFAULTLIB:msvcrt.lib`）→ 门 **#96 绿**，**wall 19m19s → 12m10s（−37%）**，vbp 930→544 s、bas 344→219 / 374→235 s。本地入口 dev.ps1/test.bat 也已默认吃 缓存（`41ce0fd`）。新账 #77/#78 见 030 §10.5/§10.7。
               # 上一轮 = 2026-09-26T10:59:52+08:00（ai/030 两批：T30-A 内容寻址 obj store + 套件开关；其原注释链保留在下面）
               # 上一轮 = 2026-09-26T06:45:00+08:00（= 本轮之前的收线时刻，其原注释链保留在下面）
               # B21 交付一笔 = `2ddcc8b`（cgen 侧 15 文件 + 用例 32 条 + 手册 Boolean 页 + 分类护栏脚本
               # `.build\b21_emitc_guard.py`）。根因是两条不是一条、护栏 RED 一次的教训、以及顺带量出的 `Print #`
               # 那条，全在 **D70**；待拍板 5 就此收口，新撞出的一条记为待拍板 7（未拍板、未动手）。
               # 上一轮 = **B20（ai/028：反引号原始多行串 + 串内插值，run #82 全绿，head `c2f7317`）**：需求出自
               # `todo/vi.md`，**一切落在词法层**（无孔串折成与手写 `"..."` 同形的普通记号、有孔串展开成
               # `( "文" & CStr( 式 ) & ... )` ⇒ parser / AST / semantics / cgen / RTL 零改动）；交付三笔 =
               # `4c16f50` + `f8ec76e`（夹具按字节钉死 `-text`，防 CI 上四份塌成两份）+ `c2f7317`，读数见 **D68**/**D69**。
               # 再上一轮 = **B19（控制台/管道编码 + 错误出口 + 套件超时预算，run #73）**；再往前 = B18 收口批（run #68）。
               # 自动运行见本行不足 55 分钟请立即跳过。
               # 另开一条线（2026-09-26，用户指定）：**内置控件补全** = `ai\029-内置控件补全计划书.md`（11 份控件说明里
               # 缺的走 comctl32 原生窗口类，Data / OLE 容器明确不做）。进度与实测记在 029 §九，本表只挂这一行指针；
               # 那条线动共用文件（`src/project/frm_parser*`、`src/backend/detail/module/cgen_form_*`、`src/rtl/core/vb6forms/*`、
               # `tests/run_tests.ps1`）之前，先看本表 STATUS —— 若 BUSY 按同一套重入保护等静默，两条线不并行改同一文件。
LAST_COMMIT: 代码批 = 2ddcc8b(B21 布尔可见性+装箱)、c2f7317(B20 插值)、4c16f50+f8ec76e(B20 多行串/夹具钉字节)、8174219(B19 套件超时/读数)、d4e53c2(B19 错误出口)、796220d(B19 控制台编码)、732c1f8(B17 门后修堆损坏)、a5517fa(B17)、578faa4(B16)、941b6dc(B15)、0caa3c5(B15 门 head = 合并用户 Fix 192/193)、ddf4e9b(B14 测试批)、4534a83(B14 台账)、b10ec1a(B13e)、5d29a5c(B13e 台账)、9df23ba(B13d)、bd38798(B13c)、b1a8e58(B13b)、7b6570a+988c7cb(B13a；门 head = 合并 `9785f4f`)、7f829ee(B11/C05=B12)、b82a184+02d70fe(B11/C04)、c4aaa4c(B11/C03b)、e515d89(B11/C03a)、f0b820d(B11/C02)、e7c7a31(B11/C01)、3c5d8e6(B10)、9eb2ca7(B09c)、debb110(B09b)、02bac92(B09)   # **commit message 一律现写、不复用上批文本**；push 只推 `github/dev`（Actions 门），`origin`(gitcode) 与 `main` 不碰、**绝不建 MR**。
# （上一轮的 LAST_COMMIT，降为注释：本文件只允许一个 LAST_COMMIT 键） 代码批 = 0413bb9(B18)、732c1f8(B17 门后修堆损坏)、a5517fa(B17)、578faa4(B16)、941b6dc(B15)、0caa3c5(B15 门 head = 合并用户 Fix 192/193)、ddf4e9b(B14 测试批)、4534a83(B14 台账)、b10ec1a(B13e)、5d29a5c(B13e 台账)、9df23ba(B13d)、bd38798(B13c)、b1a8e58(B13b)、7b6570a+988c7cb(B13a；门 head = 合并 `9785f4f`)、7f829ee(B11/C05=B12)、b82a184+02d70fe(B11/C04)、c4aaa4c(B11/C03b)、e515d89(B11/C03a)、f0b820d(B11/C02)、e7c7a31(B11/C01)、3c5d8e6(B10)、9eb2ca7(B09c)、debb110(B09b)、02bac92(B09)   # **commit message 一律现写、不复用上批文本**；push 只推 `github/dev`（Actions 门），`origin`(gitcode) 与 `main` 不碰、**绝不建 MR**。
               含完整 COM）的收口文档 = `ai\027-接口继承CoClass实施收口.md`（交付总览表 + 语言/COM 两侧要点 +
               **v1 边界清单 12 条** + 怎么验 + 记录索引）；逐格过程与全部实测读数 = 本文件的设计记录 D1–D65。
               **若还要继续推进，下面是边界清单里值得单独立项的几条**（都不是本线的"未做完"，是明确划出去的）：
               1. `ComObj_Invoke` 的 invkind fallback（**D61-6**：dispid 配得上但 invkind 配不上时取"第一个同 dispid
                  表项"，属性 Get/Let 同名时会跨方向打到对方；收紧会改存量可观察行为 ⇒ 需要拍板）。
               2. canonical COM 返回形状（`HRESULT` + `[out, retval]`）与**跨世界身份合一**（外部客户拿包装器、
                  进程内是薄指针）—— B17 测量②已裁决"不做"，要做的话是一次对外 ABI 改动。
               3. `.bas` 里声明的新式接口当类型用在调用点认不出（D43③）、接口变量上的 `Property Let/Set` 与带
                  `Optional` 的接口槽调用点（B10 的三条 caller 侧洞）、接口值作实参/进 `Variant`（B06c）、
                  ⑮d、祖先 `Private` UDT 进方法签名（D40 末①）、`com_entry` 基类 extern 的 `void*` 返回、
                  `Class_Terminate` 在 EXE 里无触发点（D41）。
               4. **派生类自己 `Implements` 新式接口**（`VB3022`，B18 写示例时实测）与**经继承满足的接口契约**、
                  **继承来的 `Public` 字段对外 COM 暴露** —— 这三条要一起做（都在 stage 2.7 Pass D 与 3.4 的成员
                  合并那一片，028 类的工作量）。
               5. **【已收 = B21，读数见 D70】`CStr(布尔)` 与 `&` / `Print` 两套读数不一致**（B20 写插值用例时实测；A/B 确认与本批无关，改前改后同读数）：`CStr(True)`、`CStr(1 > 0)`、`CStr(CBool(True))` 全回 `"-1"`，而同一个值走 `True & ""` 回 `"True"`（VB6 两处都回 `"True"`）；顺带 `TypeName(True)` 回 `"Long"`，VB6 回 `"Boolean"` ⇒ 布尔的"字符串形状"在 C3 里有两条路，且 `TypeName` 这一条也不对。**这条会直接咬到新语法**：`${flag}` 插值降级成 `CStr(flag)`，于是插值里的布尔与 `Debug.Print flag` 读数不一致。→ B21 收完：根因是**两条**（登记表看不见布尔 + 装箱按 C 类型选到 `VT_I2`/`VT_I4`），两条一起修完读数才合一；`tests\test_bool_display.bas` 32 条钉死，含 `Integer`/`Long`/`Byte` 的反向护栏。
               6. **工程内类经 `CreateObject` 编译期改写后交给 `As Object`，按名点公有 `Function` 报 `vb6_ComCall: method "…" not found`**（早绑定 `Dim o As <类名>` 正常）—— B20 的 `tests\rawstr_proj\` 第一版就是这样写的，改成早绑定才通。与 B17 那条"契约成员是 Private ⇒ 默认面点不到"不是一回事：这里 `Note()` 是 `Public`。要么晚绑定那一面缺一块，要么改写出来的 VARIANT 类型标记不对 ⇒ 先量（外部注册 DLL 那条 `cc_dll_late_client` 是通的，所以缺口在"in-project 改写出来的对象"这一支）。
               7. **`Print #` / `Write #` 的实参不分类型**（B21 顺手量出，`.build\b21_out\probe_print.bas`）：非 BSTR 的实参一律先 `vb6_Str((int32_t)x)` 再落盘 ⇒ 布尔写成 `-1`（VB6 是 `True`）、`3.5` 被截成 `3`（VB6 是 ` 3.5`）、`Write #1, True` 写成 `"-1"`（VB6 是 `#TRUE#`）。这条不是布尔专属，是 `Print #` 那一族缺按 VB 类型分派的那一层；B21 只收了 `CStr` / `Format` / `&` / `String` 形参 / Variant 装箱 / `Debug.Print` 六处出口，落盘这一处另开一批（会改存量工程的落盘字节）。
               规则沿用：push 只推 `github/dev`；门跑 Actions（`.build/wait_run2.py <sha> <秒>` 盯）；`.build` 里的
               临时 `.ps1` 一律 ASCII only；用例文件按同目录邻居的编码/行尾（`.bas`/`.vbp` = UTF-8+CRLF，
               `tests\*.ps1` = BOM+CRLF）。
# （2026-09-27 的一次误读已订正：run #150 的 head `ab6d287` 是别人那批的合并门，不覆盖 C29-RT-b；本批真正的门是 #151。）
GATE_BASELINE: (Actions 门) c3test run **#230 [dev] = completed/success**（head `061d3b4` = 账 #164 `PictureBox 不进 tab 序`（排除表补一类），**attempt 1**，wall 571s = 9m31s；逐 job：Build 207s / vbp#1 251s / vbp#2 243s / vbp#3 248s / vbp#4 254s / bas#1 354s / bas#2 359s / asm 91s / smoke 44s / syntax 26s / compile 21s ⇒ GATE_BASELINE 从 **#227** 挪到 **#230**。**CI 工件对照取的是"同一片 shard、前后两轮"这一手**（比跟本地日志比更硬）：`test-logs-vbp-1`（49 个文件）门 #227 vs #230 **只有两份动** —— `TabWalkApp.out` 动五行（`TW-picstop` N→Y、`TW-seq` 少 `pic1` 一站、`TW-new` 非判据、两行按句柄数值打印的证人行；其余 11 行逐字节相同），`ModalApp.out` 动两行（`M1-startup=Y/second=N` 括号里的句柄数 + `busy=14→7`，**结论位一位没动**）；`test-logs-vbp-2`（39 个文件）只有 `TmApp.out` 的 `T11=Y/51→50`（拍数，账 #162 那一族）动，邻居 `BfApp.out` / `CbApp.out` / `TiApp.out` **逐字节相同**。本地补架构那一头：x64 与 x86 两份 tabwalk 日志抹掉两行句柄后**逐行相同**、与 CI 那份只差 `TW-new` 一格；CI 里 x86 那份根本取不到（两架构写同一个 `TabWalkApp.exe` 名，工件只剩一份 —— 同 #156 那条已知限制，不是本轮新问题）。`TabWalkApp.err` 与另两份 `.err` 均 0 字节。）
GATE_BASELINE: (Actions 门) c3test run **#227 [dev] = completed/success**（head `69845a9` = 账 #158 按钮那一半 `BS_NOTIFY` + 焦点码表接上第二条创建路那一格，**attempt 1**，wall 565s = 9m25s；逐 job：Build 211s / vbp#1 304s / vbp#2 246s / vbp#3 187s / vbp#4 252s / bas#1 260s / bas#2 350s / asm 79s / smoke 42s / syntax 20s / compile 16s；**CI 工件对照**：`test-logs-vbp-1`(x64) 与 `-2`(x86) 两份 `BfApp.out` 各 **122 字节、彼此逐字节相同、也与本地两份相同**（BF 六条 `1/1` + `both/noclick/each` 全 Y + `BTNFOCUS-DONE` 十条齐），两份 `.err` 各 0 字节；同两片里 `TabWalkApp.out` 与 `CbApp.out` 与上一轮（门 #224 工件）逐字节相同 ⇒ 邻居没被带坏）⇒ GATE_BASELINE 从 **#224** 挪到 **#227**。**上一轮 #225（head `db576da`）红一条 job**：红的不是行为，是本批自己把创建样式位挪了一格，而账 #83(a) 的 `cs_emitc_tabstop` 把旧的样式数逐字钉着 ⇒ 三串针按实测重算 + 补两条反面针，修完就是本轮的 `69845a9`。另：`gh run watch` 在这一轮**两次**以 `HTTP 504` 抛错退出（`WATCH-RC=1`）而门其实还在跑、最后是绿的 —— 看门不能只看 watcher 的退出码，要回 Runs API 按 `head_sha` 复核（本轮 #226 是别人 `gate/merge-dev` 的门，按 head 归到它自己名下，不计）。
GATE_BASELINE: (Actions 门) c3test run **#224 [dev] = completed/success**（head `a54193a` = 账 #158 Combo 焦点通知的码表订正那一格，**attempt 1**，wall 619s = 10m19s；逐 job：Build 239s / vbp#1 500s / vbp#2 490s / vbp#3 238s / vbp#4 514s / bas#1 347s / bas#2 338s / asm 113s / smoke 40s / syntax 21s / compile 24s；**CI 工件对照**：`test-logs-vbp-1`(x64) 与 `-2`(x86) 两份 `CbApp.out` 各 **88 字节、彼此逐字节相同**（`CF-cb1=1/1`、`CF-cb2=1/1/1`、`CF-lb=1/1`、`CF-tx=1/1`、`CF-each=Y`、`CF-cross=Y` + `COMBOFOCUS-DONE` 七条齐），两份 `.err` 各 0 字节）⇒ GATE_BASELINE 从 **#223** 挪到 **#224**。
GATE_BASELINE: (Actions 门) c3test run **#223 [dev] = completed/success**（head `eda3906` = 账 #160 设计期 `TabIndex` 下发那一格，**attempt 1**，wall 575s = 9m35s；**CI 工件对照**：`test-logs-vbp-1`(x64) 与 `-2`(x86) 两份 `TiApp.out` 各 **81 字节、彼此逐字节相同、也与本地两份相同**（TI 四条 + TABINDEX-DONE 齐），两份 `.err` 各 0 字节）⇒ GATE_BASELINE 从 **#222** 挪到 **#223**。
GATE_BASELINE: (Actions 门) c3test run **#222 [dev] = completed/success**（head `77c9d8c` = 账 #166 的产物名单一权威那一格（顶上叠着 #83(b2) 收线那笔纯文档提交 `0cbfd0e`，同一轮门里），**attempt 1 即 11 job 全绿**，wall 562s = 9m22s；这一轮 Runs API 的逐 job 时间戳同样返回 null（job 复用），只记 wall。**CI 工件对照**：`test-logs-vbp-4` 里新证人 `RenamedProbe.out` = **30 字节**、`NP-ok=Y/tick=1` + `EXENAME-DONE` 齐；四片工件里**没有任何 `NameProbe.*`** ⇒ 产物名权威只剩 `.vbp` 的 `ExeName32` 一个。（另记：本轮同期 #219/#221 两轮红是 **`gate/merge-dev` 分支**（别人的合并门），不是 dev —— 判门一律按 head_branch + head_sha 复核，`watch-gh-actions.ps1` 自己会提前 exit 0。）⇒ GATE_BASELINE 从 **#220** 挪到 **#222**。
GATE_BASELINE: (Actions 门) c3test run **#220 [dev] = completed/success**（head `5922033` = 账 #83(b2) 的夹具命名对齐那一格（`7be26c9` 那半刀产品改动叠在下面，同一轮门里），**attempt 1 即 11 job 全绿**，wall 563s = 9m23s；逐 job 时间戳这一轮 Runs API 返回 null（job 复用），只记 wall。**CI 工件对照**：`test-logs-vbp-1` 与 `-4` 两片各一份 `TabWalkApp.out`（**371 / 372 字节**，差的 1 字节是句柄十进制位数），两条架构读数逐行相同 —— `EX-fr=262144/f2=262144/pic=262144/form=256`、`TW-in1=Y/in2=N`、`TW-deep=N/inpic=N`、`TW-picstop=N`、`TW-top=Y`、`TW-shy=Y`、`TW-walked=Y`、`TABWALK-DONE`，两份 `.err` 各 0 字节 ⇒ 本批那三位**缺陷读数**在 CI 与本地一致，#165 落地时它们会一起翻红。（中间那一轮 **#218 = failure**，红 `tabwalk` / `tabwalk_x86` 两条 `FAIL (no exe)`，根因是夹具的 `ExeName32` 与 vbp 文件名不一致 = 两个权威管同一个名字，不是产品；单一权威那条改法开成 **账 #166**。）⇒ GATE_BASELINE 从 **#217** 挪到 **#220**。
GATE_BASELINE: (Actions 门) c3test run **#217 [dev] = completed/success**（head `35e27be` = 账 #83(b1) 的夹具加固那一格（`b6d6fcd` 那半刀产品改动叠在下面，同一轮门里），**attempt 1 即 11 job 全绿**，wall 610s = 10m10s；逐 job Build 200s / vbp#1 216s / vbp#2 267s / vbp#3 322s / vbp#4 232s / bas#1 323s / bas#2 407s / asm 79s / smoke 53s / syntax 24s / compile 12s。**CI 工件对照**：两片 `ModalApp.out` 各 **265 字节**、`MODAL-DONE` 齐、`.err` 各 0 字节，`MW-new=4/repeat=txtMain/hops=3` 与 `D1..D7` 七条两片一致；唯一差别就是在途 tick 被数出来那一条（一片 `ticks=1/busy=7`、一片 `ticks=1/busy=0`），与门 #215 同两个位置的 `ticks=21`｜`ticks=1` 对上 ⇒ **账 #162 就此结**。（中间那一轮 **#216 = failure**，红 `modal_x86` 一条 output mismatch，根因是夹具缺相位闸、不是产品 —— 加固后即 #217。）⇒ GATE_BASELINE 从 **#211** 挪到 **#217**（#212..#216 只记在 LAST_RUN 那几行里，没单独建行；本轮一并越过）。
GATE_BASELINE: (Actions 门) c3test run **#211 [dev] = completed/success**（head `4debd38` = C29-SL-r（账 #83 半 (a)）那一格，**attempt 1 即 11 个 job 全绿**，job 跨度 **587s = 9m47s**；逐 job：Build 208s / vbp#1 213s / vbp#2 231s / vbp#3 251s / vbp#4 224s / bas#1 358s / bas#2 377s / asm 86s / smoke 41s / syntax 25s / compile 24s。CI 工件 `test-logs-vbp-3`(x64) 与 `-4`(x86) 的 `CtrlState.out` 各 **250 字节**、与本地两份**逐字节相同**（`ST1`/`ST2` 齐、`.err` 0 字节）；中间那个 `#210` 是别人 workflow_dispatch 的一轮（head `9f36ae2`，失败），与本线无关 ⇒ 认 head 不认门号 ⇒ GATE_BASELINE 从 #209 挪到 **#211**。）
GATE_BASELINE: (Actions 门) c3test run **#209 [dev] = completed/success**（head `a87548f` = C29-SL-q 那一格（顶上叠着 SL-p 收线那笔纯文档提交），**attempt 1 即 11 个 job 全绿**，job 跨度 **573s = 9m33s**；逐 job：Build 207s / vbp#1 253s / vbp#2 210s / vbp#3 282s / vbp#4 229s / bas#1 251s / bas#2 364s / asm 91s / smoke 88s / syntax 28s / compile 21s。CI 工件 `test-logs-vbp-3`(x64) 与 `-4`(x86) 的 `SlidApp.out` 各 **1635 字节**、与本地两份**四份逐字节相同**（`SQ1/SQ2/SQ3` 三条齐、两份 `.err` 各 0 字节）⇒ GATE_BASELINE 从 #208 挪到 **#209**。）
GATE_BASELINE: (Actions 门) c3test run **#208 [dev] = completed/success**（head `1f7e634` = C29-SL-p 那一格（纯源码+夹具+针，没有叠文档），**attempt 1 即 11 个 job 全绿**，job 跨度 **520s = 8m40s**；逐 job：Build 128s / vbp#1 259s / vbp#2 236s / vbp#3 234s / vbp#4 251s / bas#1 348s / bas#2 390s / asm 164s / smoke 36s / syntax 23s / compile 21s。CI 工件 `test-logs-vbp-3`(x64) 与 `-4`(x86) 的 `SlidApp.out` 各 **1589 字节**、与本地两份**四份逐字节相同**（`SP1..SP3` 四条齐、两份 `.err` 各 0 字节）⇒ GATE_BASELINE 从 #207 挪到 **#208**。）
GATE_BASELINE: (Actions 门) c3test run **#207 [dev] = completed/success**（head `81e845c` = C29-SL-o 那一格（顶上叠着 SL-n 收线那笔纯文档提交），**attempt 1 即 11 个 job 全绿**，job 跨度 **595s = 9m55s**；逐 job：Build 213s / vbp#1 245s / vbp#2 233s / vbp#3 267s / vbp#4 231s / bas#1 351s / bas#2 379s / asm 89s / smoke 45s / syntax 26s / compile 17s。CI 工件 `test-logs-vbp-3`(x64) 与 `-4`(x86) 的 `SlidApp.out` 各 **1526 字节、彼此逐字节相同**，且与本地 `.build/sloa/run.out`、`.build/slob/run.out` 逐字节相同（`SO1/SO2/SO3` 三条齐、两份 `.err` 各 0 字节）⇒ GATE_BASELINE 从 #206 挪到 **#207**。）
GATE_BASELINE: (Actions 门) c3test run **#206 [dev] = completed/success**（head `36d4fba` = C29-SL-n 那一格（顶上叠着 SL-m 收线那笔纯文档提交），**attempt 1 即 11 个 job 全绿**，wall **539s = 8m59s**；逐 job：Build 208s / vbp#1 234s / vbp#2 254s / vbp#3 221s / vbp#4 227s / bas#1 324s / bas#2 322s / asm 80s / smoke 45s / syntax 23s / compile 27s。**CI 工件对照**：`test-logs-vbp-3`(x64) 与 `-4`(x86) 两份 `SlidApp.out` 各 **1460 字节、彼此逐字节相同**，且与本地 `.build/slla/run.out`、`.build/sllb/run.out` 也逐字节相同；SN 七条在 CI 产物里齐（含 `SN7-dblone=1/0`），两份 `.err` 各 0 字节。⇒ GATE_BASELINE 从 #205 挪到 **#206**。
GATE_BASELINE: (Actions 门) c3test run **#205 [dev] = completed/success**（head `69c79b7` = C29-SL-m 那一格（顶上叠着 SL-l 收线那笔纯文档提交），**attempt 1 即 11 个 job 全绿**，wall **542s = 9m02s**；逐 job：Build 177s / vbp#1 246s / vbp#2 180s / vbp#3 243s / vbp#4 228s / bas#1 352s / bas#2 350s / asm 69s / smoke 48s / syntax 22s / compile 26s。**CI 工件对照**：`test-logs-vbp-3`(x64) 与 `-4`(x86) 两份 `SlidApp.out` 各 **1444 字节、彼此逐字节相同**，且与本地 `.build/slla/run.out`、`.build/sllb/run.out` 也逐字节相同；SN 六条在 CI 产物里齐（含 `SN4-with-bare=Y`、`SN5-with-paren=Y/Y`、`SN6-with-clearsel=0/0`），两份 `.err` 各 0 字节（工件下载第一趟撞 502，重试才拿回；watch-gh-actions 又在 17:4x 被 504 打断一次，判定改用 Runs API）。⇒ GATE_BASELINE 从 #204 挪到 **#205**。
GATE_BASELINE: (Actions 门) c3test run **#204 [dev] = completed/success**（head `441b6f6` = C29-SL-l 那一格（顶上还叠着 #147 收线那三笔纯文档提交），**attempt 1 即 11 个 job 全绿**，wall **524s = 8m44s**；逐 job：Build 144s / vbp#1 278s / vbp#2 212s / vbp#3 239s / vbp#4 239s / bas#1 375s / bas#2 279s / asm 87s / smoke 44s / syntax 26s / compile 24s。**CI 工件对照**：`test-logs-vbp-3`(x64) 与 `-4`(x86) 两份 `SlidApp.out` 各 **1384 字节、彼此逐字节相同**，且与本地 `.build/slla/run.out`、`.build/sllb/run.out` 也逐字节相同；SN 三条在 CI 产物里就是 `SN1-bare=Y/0`、`SN2-paren=Y/Y`、`SN3-clearsel=0/0`，两份 `.err` 各 0 字节。（下载 `-4` 第一趟撞 HTTP 502，重试拿回；watch-gh-actions 那一轮则被 504 整个打断，判定改用 Runs API 的 job 级结论。）⇒ GATE_BASELINE 从 #203 挪到 **#204**。
GATE_BASELINE: (Actions 门) c3test run **#203 [dev] = completed/success**（head `0c99d8f` = C29-SL-k 那一格，**attempt 1 即 11 个 job 全绿**，wall **525s = 8m45s**；逐 job：Build 166s / vbp#1 250s / vbp#2 218s / vbp#3 258s / vbp#4 252s / bas#1 351s / bas#2 354s / asm 80s / smoke 41s / syntax 22s / compile 17s。**CI 工件对照**：`test-logs-vbp-3`(x64) 与 `-4`(x86) 两份 `SlidApp.out` 各 **1337 字节、彼此逐字节相同**，且与本地那两份（`.build/slka/run2.out` x64、`.build/slkb/run.out` x86）也**逐字节相同**；SK 七条在 CI 产物里就是 `SK1-text=VOL/String/8`…`SK7-end=N/0`，两份 `.err` 各 0 字节。⇒ GATE_BASELINE 从 #202 挪到 **#203**。
GATE_BASELINE: (Actions 门) c3test run **#202 [dev] = completed/success**（head `f7ec817` = C29-SL-j 那一格，**attempt 1 即 11 个 job 全绿**，wall **547s = 9m07s**；CI 工件已对照：`test-logs-vbp-3`(x64) 与 `-4`(x86) 逐字节相同、且与本地真跑那份也逐字节相同，`SJ1-reg=Y/Y/N` 齐、两份 `.err` 各 0 字节）
GATE_BASELINE: (Actions 门) c3test run **#201 [dev] = completed/success**（head `133175c` = C29-SL-i 那一格，**attempt 1 即 11 个 job 全绿**，wall **494s = 8m14s**。CI 工件对照同日补做：`test-logs-vbp-3`(x64) 与 `-4`(x86) 两份 `SlidApp.out` `cmp` 逐字节相同，且与本地真跑那一份也逐字节相同（81 行 / 1213 字节），两份 `.err` 各 0 字节）
GATE_BASELINE: (Actions 门) c3test run **#200 [dev] = completed/success**（head `cf7ab6b` = C29-SL-h 那一格，**attempt 1 即 11 个 job 全绿**，wall **475s = 7m55s**。CI 侧亲手复核：`test-logs-vbp-3`(x64) 与 `-4`(x86) 两份 `SlidApp.out` `cmp` 逐字节相同（71 行），且与本地真跑那一份也逐字节相同；`SH1..SH9` 九条齐、两份 `.err` 各 0 字节。另：#199/#198 是别人分支（`wip/vbeclipse-c1` / `-b2`）的 workflow_dispatch，那两个 sha 在本仓 `git cat-file` 查不到，与本线无关）
GATE_BASELINE: (Actions 门) c3test run **#197 [dev] = completed/success**（head `3c7d247` = C29-SL-g 那一条，**attempt 1 即 11 个 job 全绿**，wall 571s = 9m31s；逐 job：Build 213s / vbp#1 283s / vbp#2 244s / vbp#3 255s / vbp#4 216s / bas#1 352s / bas#2 282s / asm 82s / smoke 43s / syntax 56s / compile 20s。本批点名核过：`[EMITC-SHAPE] sl_emitc_native`/`sl_emitc_events`(vbp#1/#2)、`[EMITC-ABSENT] sl_emitc_no_com_fallback`(vbp#3)、`[VBP] ctrlslider`(vbp#3) 与 `ctrlslider_x86`(vbp#4) 全 PASS；工件里两份 `SlidApp.out`（vbp#3 = x64、vbp#4 = x86）逐字节相同、且与本地真跑那份逐字节相同（61 行），SG1-got / SG2-move / SG3-back / SG4-isolate 四条齐，两份 `.err` 各 0 字节。连着两批 attempt 1 一次过（#196/#197），`frmevents` 那格随机红这轮回没出现，账 #79 仍挂着）
               # 上一条基线 = #192 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#192 [dev] = completed/success**（head `32a20b8` = 账 #128-a 那一格，**attempt 1 即 11 个 job 全绿**，wall **551s = 9m11s**；逐 job：Build 183s / vbp#1 249s / vbp#2 221s / vbp#3 174s / vbp#4 248s / bas#1 316s / bas#2 361s / asm 83s / smoke 46s / syntax 28s / compile 27s。五条新读数从 CI 工件里逐字取到（x64/x86 两臂各一次、两臂相同）：`TV38=True/False/Boolean`、`DT43=True/Boolean/11`、`MV41=True/True/Boolean`、`RT90=True/False/Boolean`、`TS34=True/Boolean/11`，改字面的两条旧针在 CI 也是 `TS6-WRAP=False` / `TS13-WRAP=True`。上一轮红过的 `frmevents` 这轮随 vbp #1 照常绿（本批没新增 `Test-*` 调用 ⇒ 分片归属未动）。另记：门 #191 = 上一格 C29-SL-c 的门，**attempt 2 才绿**（attempt 1 红在 frmevents 的整份空白 stdout，`gh run rerun --failed` 重跑 vbp #1 即可，详情与归因见 029 那格）。
               # 上一条基线 = #191 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#191 [dev] = completed/success（attempt 2）**（head `18a62b2` = C29-SL-c 那一格；**attempt 1 的 11 job 里 10 绿、唯一的红是 vbp #1 的 `frmevents`（`Got:` 整份空白，旧脆弱 #79，且是本批新增一条发码针把它的分片从 #4 挪到 #1 才暴露的**，本批发码对该工程 0 行差异 + 本地新 exe 真跑 17 行全齐；`gh run rerun --failed` 只重跑 vbp #1（18:45:37Z→18:50:09Z = 272s）即绿，整跑 `completed/success`。attempt 2 工件：`test-logs-vbp-3` 与 `-4` 两份 `SlidApp.out` **逐字节相同**（34 行，登记的 33 条针零缺失）、`.err` 各 0 字节；发码针分别在 vbp #1/#2/#3 随分片绿）
               # 上一条基线 = #190 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#190 [dev] = completed/success**（head `b93e62d` = C29-SL-b 那一格，**11 个 job 全绿**，wall **569s = 9m29s**；逐 job：Build 202s / vbp#1 208s / vbp#2 263s / vbp#3 233s / vbp#4 286s / bas#1 359s / bas#2 364s / asm 91s / smoke 47s / syntax 23s / compile 23s。本批四条用例点名：`sl_emitc_native`(vbp#1) / `sl_emitc_no_com_fallback`(vbp#2) 随各自分片绿，`ctrlslider` / `ctrlslider_x86` 全 PASS；工件 `test-logs-vbp-3` 与 `-4` 两份 `SlidApp.out` 逐字节相同（`cmp` 过，22 条读数全在，SL-a 那 11 条一条没动）、`.err` 各 0 字节）
               # 上一条基线 = #189 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#189 [dev] = completed/success**（head `d13b652` = C29-SL-a 那一格，**11 个 job 全绿**，wall **627s = 10m27s**；逐 job：Build 196s / vbp#1 201s / vbp#2 218s / vbp#3 234s / vbp#4 270s / bas#1 359s / bas#2 426s / asm 123s / smoke 42s / syntax 29s / compile 18s。本批四条用例点名核过 `sl_emitc_native` / `sl_emitc_no_com_fallback` / `ctrlslider` / `ctrlslider_x86` 全 PASS；`test-logs-vbp-3` 与 `-4` 两份 `SlidApp.out` 逐字节相同、`.err` 各 0 字节）
               # 上一条基线 = #188 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#188 [dev] = completed/success**（head `ecc8333` = 账 #124 那一格 + 合并 `aa92fde`，**8/8 job 全绿**，wall **524s = 8m44s**；逐 job：Build 159s / vbp#1 200s / vbp#2 224s / vbp#3 191s / vbp#4 263s / bas#1 357s / bas#2 350s / asm 80s / smoke 45s / syntax 24s / compile 16s。本批四条用例点名核过：`cs_emitc_state`(vbp#1) / `cs_emitc_selectivity`(vbp#2) / `ctrlstate`+`ctrlsstab`(vbp#3) / `ctrlstate_x86`(vbp#4) 全 PASS，四个分片 FAIL=0；工件 `CtrlState.out` 与 `CtrlSSTab.out` 两份读数与本机逐字相同、`.err` 各 0 字节）
               # 上一条基线 = #187 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#187 [dev] = completed/success**（head `aa92fde` = 对方 CI 分片修复那一格，**含本批 `ec945c1`**（核过祖先关系），**8/8 job 全绿**，wall **577s = 9m37s**；逐 job：Build 216s / vbp#1 204s / vbp#2 224s / vbp#3 236s / vbp#4 270s / bas#1 260s / bas#2 358s / asm 87s / smoke 53s / syntax 33s / compile 16s。本批四条新用例点名核过：`[VBP] ctrlstate`（vbp#3）/ `ctrlstate_x86`（vbp#4）/ `[EMITC-SHAPE] cs_emitc_state`（vbp#1）/ `[EMITC-ABSENT] cs_emitc_selectivity`（vbp#2）全 PASS；工件 `test-logs-vbp-3` 里 `CtrlState.out` 八条读数 + `CTRLSTATE-DONE`、`.err` 0 字节）
               # 上一条基线 = #182 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#182 [dev] = completed/success**（head `00f241c` = 账 #88 那一格，**8/8 job 全绿**，wall **973s = 16m13s**；逐 job：Build 202s / vbp 765s / bas#1 341s / bas#2 275s / asm 81s / smoke 34s / syntax 25s / compile 16s。用例级从 `test-logs-vbp` 工件核：`TvfApp.out` = `TV15=Y TV35=Y TV36=N TV37=Y TV34=Y` + 36 行 `=Y` + `TREEVIEW-DONE`、`TvfApp.err` 0 字节；`ctrltreeview` / `ctrltreeview_x86` / `tv_emitc_shape` / `tv_emitc_no_com_fallback` 四条一起 PASS；vbp 整格 `Results: PASS=126 FAIL=0 SKIP=1 TOTAL=127`，与 #181 同数 —— 本批只升级既有夹具判据、没新增用例）
               # 上一条基线 = #181 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#181 [dev] = completed/success**（head `77c251d` = 账 #123 那一格，**8/8 job 全绿**，wall **965s = 16m05s**；逐 job：Build 200s / vbp 757s / bas#1 259s / bas#2 258s / asm 84s / smoke 45s / syntax 24s / compile 24s）。用例级从工件取到：`test_byte_visible.out` 与 `_x86` 各 26 行、两份 `.err` 0 字节、**两架构逐字相同**（比较面 局部/模块/Const 三条 + 赋 String 四条 + 装箱 VT_UI1=17 六条 + ByVal/ByRef 形参四条 + 不许被牵连的算术/Print/混算四条 + 2 万次装箱两条）；日志里 `[EMITC-SHAPE] byte_emitc_visible ... PASS` / `[EMITC-ABSENT] byte_emitc_no_variant_addr ... PASS`；bas 两分片 `PASS=38` + `PASS=39`（**77/77**，比 #180 +2），vbp 整格 `Results: PASS=126 FAIL=0 SKIP=1 TOTAL=127`。**本批最硬的一条不是发码形状而是真跑**：Charts 2020 里 `Dim Alpha As Byte` 的 `Alpha < 128` 以前发 `vb6_VarCmpLongLt(&Alpha, 128L)`（1 字节地址当 VARIANT），是同一个静默错答案在真实工程里的实例，现在走直接 C 比较且该工程随 126 枚一起绿。另记一条方法论：**"存量工程 `--emit-c` 零差异"不等于改动惰性** —— 本批两处自身错误（把 Byte 塞进 knownLongVars_ 令 VT_UI1 退成 VT_I4；形参登记写进分支体内导致 Byte 形参没登记）全都被专门探针抓住，而 11 件存量输入的逐行 diff 一声不吭。
               # 上一条基线 = #180 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#180 [dev] = completed/success**（head `30a9066` = 账 #122 那一格，**8/8 job 全绿**，wall **992s = 16m32s**；逐 job：Build 207s / vbp 776s / bas#1 346s / bas#2 347s / asm 79s / smoke 45s / syntax 29s / compile 24s）。用例级从工件取到：`test_str_cnum_assign.out` 与 `_x86` 各 19 行、两份 `.err` 0 字节、**两架构读数逐字相同**（`lit=123 / long=42 / dbl=2.5 / byte=65 / cbyte=65 / asc=65 / bool=True / expr=43 / len=4 / ubound=3 / date-ok=Y / str-left=ABC / str-ucase=XY / str-func=MMMM / func-num=21 / concat=n=42 d=2.5 / borrow=SHARED / loop=20042 / DONE`）；bas 两分片 `PASS=38` + `PASS=37`（**75/75**，比 #179 +2 = 本批这 pair），vbp 整格 `Results: PASS=124 FAIL=0 SKIP=1 TOTAL=125`，日志里 `[EMITC-SHAPE] cnum_emitc_wrapped ... PASS` / `[EMITC-ABSENT] cnum_emitc_no_overwrap ... PASS`。本批只动 cgen 一处、无 RTL 改动；**最硬的证据不是发码形状而是运行期**：护栏 diff 里那条被定性为「真实修复」的 Charts 2020 `sDiplay = CLng(...)`（修复前把 long 存进 BSTR 变量）跟着 120 枚 GUI 工程一起跑过、一枚没红。另记一条一般式：**RTL 按 C 编 ⇒ `int` 传 `BSTR` 形参只是警告，BUILD-RC=0 不代表类型对**，所以「这形状今天编不过、存量里不可能有」这类推断只对浮点那半成立，整型那半是潜伏的。
               # 上一条基线 = #179 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#179 [dev] = completed/success**（head `6e21ada` = 账 #119 **判据加固**那一格，**8/8 job 全绿**，wall **951s = 15m51s**；逐 job：Build 203s / vbp 742s / bas#1 368s / bas#2 323s / asm 76s / smoke 50s / syntax 28s / compile 16s）。用例级那一层从工件取到了：`job8/test_str_leak.out` = `api1=1 / api2=1 / delta=0 / leak-ok=Y / borrow=SHARED-ME / func-len=5 / concat=ABC / byref-len=3 head=C / chain=AAABBB / DONE`，`test_str_leak_x86.out` 同形（`delta=8192`），两份 `.err` 皆 0 字节 ⇒ 两架构都把 10 万次临时串赋值跑成私有提交零增长（阈值 2MB，修复前那 10 万只是纯漏，量级 15MB）。bas 两分片 `PASS=36` + `PASS=37`（**73/73**，比 #177 那跑 +2 = 本批两枚新用例），vbp 整格 `Results: PASS=122 FAIL=0 SKIP=1 TOTAL=123`，日志里两枚发码断言 `[EMITC-SHAPE] strleak_emitc_move ... PASS` / `[EMITC-ABSENT] strleak_emitc_no_move_borrowed ... PASS`。**120 枚 GUI 工程一枚没红 = 这批改了 150 处所有权没有造出 double-free**（本格真正的风险面）。**前一跑 #178（head `18491bc`，同一批产品码）是 failure**：只红 `Tests (bas #2)` 里本批那两枚新用例，而**产品读数逐条正确** —— 红因是针面写法：`Add-BasTest` 的比对走 PowerShell `-like`，`"S119-borrow=[SHARED-ME]"` 里的 `[SHARED-ME]` 被当**字符类**（匹配集合里的一个字符），对着带方括号的真输出必不匹配；反向同一枚针对着**不带括号**的输出反而能命中 ⇒ 这条坑既能假红也能假绿（全仓 3000 多行针面只有我这两条用了方括号，别人都躲着走）。加固 = 夹具与针面同步去掉方括号，并补做「跑承载判据的那段匹配代码」这一步（此前只手工看了 run.log 就下结论）。全程记 029 账 #119 那一节。
               # 上一条基线 = #177 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#177 [dev] = completed/success**（head `40b9e88` = 账 #91 收线那一格，**8/8 job 全绿**，wall **991s = 16m31s**；逐 job：Build 211s / vbp 774s / bas#2 326s / bas#1 267s / asm 82s / smoke 35s / syntax 17s / compile 17s）。这一格**零产品代码改动**（只把一枚读数钉进既有 bas 用例 ⇒ 门的意义是护栏而不是证据），用例级那一层照旧从工件取：`gh run download 36383298296 -n test-logs-bas-1` 里 `job14/test_declare_gmn_path_out.out` 给出 `nF=56 / gmn-fixed-ok=Y / n=56 / … / gmn-path-ok=Y`（同目录 `.err` 0 字节），bas 两分片 `PASS=36` + `PASS=35`（**71/71**），vbp 整格 `Results: PASS=120 FAIL=0 SKIP=1 TOTAL=121`。**#91 记为「他人分支已修 + 本批补钉」，不改产品码**：① 那条 `DllGetVersion` 的 `E_INVALIDARG` 出自我自己少写 `cbSize` 的探针，签名补正后 `hr=0 / maj=6 min=16`；② 那条 `GetModuleFileNameA` 崩早已被 `14c36a6` / `18ea80d` / `baea726` 修掉并在 `tests/declare_out/` 六枚用例里在册，账上挂的「崩」在 HEAD 复现不出。定长 Buffer 与动态串两形在同一枚 exe 里读回同一个 56 ⇒ 这条也算顺手给 #116 / #118 那片定长形挂了一枚运行期锚（**语义欠账仍在 #118，本行不代表它已修**）。
               # 上一条基线 = #176 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#176 [dev] = completed/success**（head `6fe4b33` = 账 #107 那一格，**8/8 job 全绿**，wall **978s = 16m18s**；逐 job：Build 210s / vbp 762s / bas#2 316s / bas#1 316s / asm 81s / smoke 38s / syntax 22s / compile 20s）。这一格动的是 **RTL**（`vb6forms_style.c` 一对读写 + `touch c3rtl.rc`），所以门的全量 GUI 面都算数：`[VBP] ctrlprop ... PASS`（本批新针 `CP9=0` / `CP10=1` / `CP11=0`，BASE 上是 `2 / 1 / 2`），其余控件工程一枚没红，vbp 整格 `Results: PASS=120 FAIL=0 SKIP=1 TOTAL=121`。
               # 上一条基线 = #175 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#175 [dev] = completed/success**（head `a322ba1` = 账 #108 那一格，**8/8 job 全绿**，wall **977s = 16m17s**；逐 job：Build 174s / vbp 799s / bas#1 320s / bas#2 314s / asm 81s / smoke 46s / syntax 20s / compile 24s）。vbp 那一格用例级证到：`[VBP] ctrlprop ... PASS`（本批新加的 `CP6=1` / `CP7=2` / `CP8=0` 三条读数在场，BASE 上是互换的 `2 / 1` ⇒ 修复前必红），整格 `Results: PASS=120 FAIL=0 SKIP=1 TOTAL=121`。
               # 上一条基线 = #174 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#174 [dev] = completed/success**（head `c4c0fde` = 账 #68 那一格，**8/8 job 全绿**，wall **924s = 15m24s**；逐 job：Build 173s / vbp 746s / bas#2 335s / bas#1 292s / asm 87s / smoke 41s / syntax 19s / compile 15s）。vbp 那一格用例级证到：`[VBP] ctrlfiles ... PASS` 与 `ctrlfiles_x86 ... PASS`（CF15/CF16 两条新读数在两架构上都到）、`[EMITC-SHAPE] cf_emitc_shape ... PASS`、`[EMITC-ABSENT] cf_emitc_no_stolen_left ... PASS`，整格 `Results: PASS=120 FAIL=0 SKIP=1 TOTAL=121`。
               # 上一条基线 = #173 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#173 [dev] = completed/success**（head `3d69179` = 账 #113 那一格，**8/8 job 全绿**，wall **989s = 16m29s**；逐 job：Build 204s / vbp 781s / bas#2 319s / bas#1 323s / asm 82s / smoke 46s / syntax 25s / compile 16s）。用例级证到了：bas 两个分片 `PASS=35` + `PASS=36`（**71/71**），工件里 `test_str_alias` 与 `test_str_alias_x86` 的 8 条读数双双逐字正确（`S113-A=PAYLOAD-0123456789` / `S113-C=SECOND-ONE` / `S113-DONE`，.err 皆 0 字节）；vbp 那一格 `Results: PASS=119 FAIL=0 SKIP=1 TOTAL=120`（`fixedstr_emitc_decl` / `fixedstr_emitc_no_ice` 继续 PASS）。
               # 上一条基线 = #172 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#172 [dev] = completed/success**（head `504a9d8` = 账 #116 那一格（含前一跑 #171 的订正），**8/8 job 全绿**，wall **994s = 16m34s**；逐 job：Build 220s / vbp 770s / bas#2 314s / bas#1 327s / asm 82s / smoke 39s / syntax 24s / compile 16s）。vbp 那一格用例级证到：`[EMITC-SHAPE] fixedstr_emitc_decl ... PASS`、`[EMITC-ABSENT] fixedstr_emitc_no_ice ... PASS`，以及被 #171 判红的两枚 GUI 存量工程 `VbQRCodegen` / `Charts2020` 双双回到 `PASS (compile, window, clean exit)`，整格 `Results: PASS=119 FAIL=0 SKIP=1 TOTAL=120`。**前一跑 #171（head `1ee2beb`）红在 vbp 的那两枚工程上（一摞 C2440：`vb6_VARIANT` ↔ `vb6_SafeArray1D*`）——红得值：那是本批把"非 SimpleTypeRef 的兜底"写成 Variant 的错，订正与复盘见 029 的账 #116 那一节。**
               # 上一条基线 = #170 那一格的原记录（#171 那一跑判红、已被本行取代），往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#154 [dev] = completed/success**（head `3c92a7a` = C29-MV-d 那一格，**8/8 job 全绿**，wall **958s = 15m58s**；逐 job：Build 210s / vbp 739s / bas#2 267s / bas#1 249s / asm 92s / smoke 44s / syntax 25s / compile 18s。用例级那一层照旧取不到（`actions/jobs/{id}/logs` 回 401），本行只记任务级结论。）
GATE_BASELINE: (Actions 级) c3test run **#157 [dev] = completed/success**（head `a3c4bc8` = C29-RT-c 那一格，**8/8 job 全绿**，wall **853s = 14m13s**；逐 job：Build 156s / vbp 693s / bas#1 303s / bas#2 288s / asm 63s / smoke 52s / syntax 24s / compile 14s。用例级那一层照旧取不到（logs 端点 401），本行只记任务级结论。）
GATE_BASELINE: (Actions 级) c3test run **#158 [dev] = completed/success（attempt 2）**（head `44db62b` = C29-RT-d 那一格，**8/8 job 全绿**，wall **737s = 12m17s**；逐 job：Build 208s / vbp 732s / bas#2 243s / bas#1 231s / asm 70s / smoke 65s / syntax 32s / compile 15s。**attempt 1 是 failure**：`Tests (vbp)` 里只红 `ctrldlg_probe`（`run timeout: 60s`，一条输出都没来得及打），同 job 的 `ctrldlg_probe_x86` 与其余 107 条全绿 —— 三条读数（DlApp 的 `--emit-c` 在 BASE/NEW 之间 md5 逐字节相同 + 本地新编译器连跑三次该用例全绿 + x86 同路径绿）判成该用例自带的卡死风险，重跑失败 job 后 8/8。全过程记 029 §九 C29-RT-d。另记一条工具进展：**用例级日志这次取到了** —— `gh run view --job <id> --log`（gh 走 keyring 那条凭据）能拿到每行 `[VBP] xxx ... PASS/FAIL`，而 gitcode 那个嵌入式 PAT 打 `/logs` 仍回 401。)
GATE_BASELINE: (Actions 级) c3test run **#160 [dev] = completed/success**（head `f182ecd` = C29-WS-a 的判据订正，**8/8 job 全绿**，wall **893s = 14m53s**；逐 job：Build 179s / vbp 710s / bas#1 294s / bas#2 217s / asm 66s / smoke 45s / syntax 30s / compile 16s。**上一跑 #159（head `886e5d9`）是 failure**：只红 `ctrlwinsock_x86` 一条（WS17 / WS22），x64 同批绿 —— 不是 flake，是判据把"x64 上两笔 UDP 合并成一次 FD_READ"这个**时序巧合**当成了不变量；x86 作业上分成两次到达就翻红。改判据（问内容不问条数）之后本地 x86 三跑 + x64 两跑零红，#160 绿 ⇒ GATE_BASELINE 从 #158 挪到 **#160**。）
GATE_BASELINE: (Actions 级) c3test run **#161 [dev] = completed/success**（head `8f64562` = C29-WS-b 那一格，**8/8 job 全绿**，wall **960s = 16m00s**；逐 job：Build 209s / vbp 743s / bas#1 298s / bas#2 284s / asm 77s / smoke 43s / syntax 32s / compile 17s。用例级日志这次照旧走 `gh run view --job`。）
GATE_BASELINE: (Actions 级) c3test run **#162 [dev] = completed/success**（head `541e17b` = C29-WS-c 那一格，**8/8 job 全绿**，wall **979s = 16m19s**；逐 job：Build 199s / vbp 778s / bas#2 295s / bas#1 294s / asm 81s / smoke 44s / syntax 25s / compile 18s。用例级那一层照旧取到（`gh run view --job 108697937890 --log`）：`[VBP] ctrlwinsock ... PASS` 与 `ctrlwinsock_x86 ... PASS`，`Results: PASS=112 FAIL=0 SKIP=1 TOTAL=113`。）
GATE_BASELINE: (Actions 级) c3test run **#166 [dev] = completed/success**（head `1549803` = C29-WS-d 那一格 + 合并别人 Fix 161d/161e 那批，**8/8 job 全绿**，wall **960s = 16m00s**；逐 job：Build 214s / vbp 743s / bas#2 308s / bas#1 299s / asm 75s / smoke 42s / syntax 18s / compile 17s。用例级：`[VBP] ctrlwinsock ... PASS` 与 `ctrlwinsock_x86 ... PASS`，整格 `Results: PASS=115 FAIL=0 SKIP=1 TOTAL=116`。**前一跑 #165（head `52c85a3`，不含本批）红在 `ctrlwinsock_x86`** —— 两个头的发码面逐字相同（`--emit-c` 各 779 行）⇒ 判成夹具自身的间歇性红，收法见 029 §九 C29-WS-d 的第 4 条与那段"门过了但浪费了一跑"。）
GATE_BASELINE: (Actions 级) c3test run **#167 [dev] = completed/success**（head `43ae6e5` = C29-WS-d 加固那一格，**8/8 job 全绿**，wall **965s = 16m05s**；逐 job：Build 195s / vbp 764s / bas#1 320s / bas#2 304s / asm 90s / smoke 42s / syntax 25s / compile 17s。这一跑的意义就是"九道闸在 CI 的 x86 作业上真起作用了"—— #165 红的那一把到达面在同一份发码下重跑全绿。）
GATE_BASELINE: (Actions 级) c3test run **#169 [dev] = completed/success**（head `506555e` = C29-WS-e 那一格，**8/8 job 全绿**，wall **957s = 15m57s**；逐 job：Build 200s / vbp 754s / bas#2 322s / bas#1 321s / asm 88s / smoke 44s / syntax 20s / compile 18s。
GATE_BASELINE: (Actions 级) c3test run **#170 [dev] = completed/success**（head `60f0241` = C29-WS-f 订正那一格，**8/8 job 全绿**，wall **984s = 16m24s**；逐 job：Build 212s / vbp 770s / bas#2 306s / bas#1 292s / asm 89s / smoke 47s / syntax 25s / compile 14s。用例级：`[VBP] ctrlwinsock ... PASS` 与 `ctrlwinsock_x86 ... PASS`（夹具已是 65 条判据），整格 `Results: PASS=117 FAIL=0 SKIP=1 TOTAL=118`。另一条工具收获：**用例级的真读数能下下来** —— `gh run download <run> -n test-logs-bas-2`，每个用例一对 `.out`/`.err`（#168 那跑里 `test_len_width` 与 `_x86` 的读数逐字相同，证了两架构都真跑到）。
GATE_BASELINE: (Actions 级) c3test run **#168 [dev] = completed/success**（head `e0501e9` = 账 #115 收线那一格，**8/8 job 全绿**，wall **938s = 15m38s**；逐 job：Build 193s / vbp 743s / bas#1 278s / bas#2 243s / asm 79s / smoke 48s / syntax 23s / compile 20s。**这一跑还顺扊把新用例在 CI 里跑到了**：`test-logs-bas-2` 那个工件里有 `test_len_width.out` 与 `test_len_width_x86.out`，两份读数逐字相同且两架构都是 LW2=5 / LW4=2 / LW11=0（修复前这三条在 x64 是 8、在 x86 是 4）。工具收获：**用例级的真读数可以直接下下来** —— `gh run download <run> -n test-logs-bas-2 -D <dir>`，每个用例一对 `.out` / `.err`；这比 `--log`（只有 PASS/FAIL 一行）与只能本地复跑都强。
GATE_BASELINE: (Actions 级) c3test run **#154 [dev] = completed/success**（head `3c92a7a` = C29-MV-d 那一格，**8/8 job 全绿**，wall **958s = 15m58s**；逐 job：Build 210s / vbp 739s / bas#1 249s / bas#2 267s / asm 92s / smoke 44s / syntax 25s / compile 18s）。
GATE_BASELINE: (Actions 级) c3test run **#153 [dev] = completed/success**（head `3f849bf` = C29-RT-dep 那一格，**8/8 job 全绿**，wall **773s = 12m53s**；逐 job：Build 157s / vbp 610s / bas#1 268s / bas#2 268s / asm 74s / smoke 41s / syntax 26s / compile 36s。用例级那一层照旧取不到 （`actions/jobs/{id}/logs` 回 401），本行只记任务级结论。）
GATE_BASELINE: (Actions 级) c3test run **#151 [dev] = completed/success**（head `45ade37` = C29-RT-b 那一格，**8/8 job 全绿**；逐 job：Build 200s / vbp 680s / bas#2 329s / bas#1 262s / asm 80s / smoke 51s / syntax 24s / compile 13s。用例级日志未取到（GitHub logs 端点 502/504），本行只记任务级结论）。
               # 上一条基线 = #148 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#148 [dev] = completed/success**（head `37f5404` = C29-RT-a 那一格（含合并别人 `7ab2406` 的 SSTabEx 那批），**8/8 job 全绿**，wall **916s = 15m16s**；逐 job：Build 167s / vbp 744s / bas#2 219s / bas#1 203s / asm 90s / smoke 45s / syntax 27s / compile 19s。vbp 那一格 `Results: PASS=108 FAIL=0 SKIP=1`，本批四条 `ctrlrichtextbox` / `ctrlrichtextbox_x86` / `rt_emitc_shape` / `rt_emitc_no_com_fallback` 全 PASS）。
               # 上一条基线 = #143 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#143 [dev] = completed/success**（head `a5bd656` = C29-MV-c 那一格，**8/8 job 全绿**，wall **936s = 15m36s**；逐 job：Build 223s / vbp 708s / bas#1 243s / bas#2 213s / asm 79s / smoke 44s / syntax 24s / compile 19s）。CI 的 vbp 那一格：`[VBP] ctrlmonthview ... PASS` 与 `ctrlmonthview_x86 ... PASS`。
               # 上一条基线 = #142 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#142 [dev] = completed/success**（head `f0c8875` = C29-MV-b 那一格，**8/8 job 全绿**，wall **615s = 10m15s**；逐 job：Build 153s / vbp 456s / bas#2 265s / bas#1 252s / asm 92s / smoke 40s / syntax 24s / compile 24s）。CI 的 vbp 那一格：`[VBP] ctrlmonthview ... PASS` 与 `ctrlmonthview_x86 ... PASS`。
               # 上一条基线 = #141 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#141 [dev] = completed/success**（head `65dbbb1` = C29-MV-a 那一格，**8/8 job 全绿**，wall **905s = 15m05s**；逐 job：Build 191s / vbp 709s / bas#2 191s / bas#1 185s / asm 80s / smoke 42s / syntax 22s / compile 19s）。CI 的 vbp 那一格直接对上号：`[VBP] ctrlmonthview ... PASS` 与 `ctrlmonthview_x86 ... PASS`（本轮新增的两枚夹具读数）。
               # 上一条基线 = #140 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#140 [dev] = completed/success**（head `adc69ad` = C29-DT-c 那一格，**8/8 job 全绿**，wall **834s = 13m54s**；逐 job：Build 131s / vbp 697s / bas#1 240s / bas#2 201s / asm 77s / smoke 52s / syntax 23s / compile 17s）。CI 的 vbp 那一格直接对上号：`[VBP] ctrldatetime ... PASS` 与 `ctrldatetime_x86 ... PASS`（那格里 DT37-DT42 这六条是本轮新针）。
               # 上一条基线 = #139 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#139 [dev] = completed/success**（head `082d7c4` = C29-DT-b 那一格，**8/8 job 全绿**，wall **944s = 15m44s**；逐 job：Build 196s / vbp 744s / bas#2 252s / bas#1 192s / asm 79s / smoke 49s / syntax 21s / compile 18s）。CI 的 vbp 那一格直接对上号：`[VBP] ctrldatetime ... PASS` 与 `ctrldatetime_x86 ... PASS`。
               # 上一条基线 = #136 那一格的原记录，往下顺移：
GATE_BASELINE: (Actions 级) c3test run **#136 [dev] = completed/success**（head `bcd47c4` = C29-DT-a 那一格，**8/8 job 全绿**，wall **715s = 11m55s**；逐 job：Build 199s / vbp 511s / bas#2 270s / bas#1 253s / asm 76s / smoke 39s / syntax 21s / compile 19s）。CI 的 vbp 那一格直接对上号：`[VBP] ctrldatetime ... PASS` 与 `ctrldatetime_x86 ... PASS`，`Results: PASS=100 FAIL=0 SKIP=1`。
               # 上一条基线 = #132 那一格的原记录，往下顺移：
               GATE_BASELINE: (Actions 级) c3test run **#132 [dev] = completed/success**（head `ea681f3` = 另一位作者把 `github/dev` 合进 gate/p20-44 后落下的那笔，**8/8 job 全绿**，wall **751s = 12m31s**；逐 job：Build 210s / vbp 537s / bas#1 256s / bas#2 241s / asm 103s / smoke 40s / syntax 37s / compile 24s）。**这一格算本线的门**：`git merge-base --is-ancestor 56d5325 ea681f3` 已核，该头里含 C29-M 的两笔源码 ⇒ 读数 = "本线那一格与别人那批同树共存、零新增失败"。
               # 上一条基线 = **#131**（head `56d5325` = C29-M 那一格，**8/8 job 全绿**，wall **779s = 12m59s**；逐 job：Build 166s / vbp 607s / bas#2 236s / bas#1 205s / asm 79s / smoke 45s / syntax 24s / compile 17s）。
               # 上一条基线 = **#128**（head `b2674bc` = C29-V6 那一格，**8/8 job 全绿**，wall **894s = 14m54s**；逐 job：Build 212s / vbp 675s / bas#2 205s / bas#1 184s / asm 82s / smoke 37s / syntax 23s / compile 16s）。
               # 上一条基线 = **#126**（head `9d32c36` = C29-5c 那一格，**8/8 job 全绿**，wall **881s = 14m41s**；逐 job：Build 199s / vbp 675s / bas#2 259s / bas#1 251s / asm 87s / smoke 39s / syntax 25s / compile 17s）。
               # 上一条基线 = **#124**（head `9b42876` = C29-8c 那一格，**8/8 job 全绿**，wall **917s = 15m17s**；逐 job：Build 205s / vbp 688s / bas#1 252s / bas#2 256s / asm 71s / smoke 36s / syntax 32s / compile 16s。本批只把 `ctrltreeview` 从 27 条扩到 33 条，vbp 相对 #122 的 676s 涨 12s ⇒ 与 runner 抖动同量级，#122 那格记下的"412s→676s 涨幅下一格再核"到此收口）。
               # 上一条基线 = **#122**（head `ee6d769` = C29-5b 那一格，**8/8 job 全绿**，wall **876s = 14m36s**；逐 job：Build 194s / vbp 676s / bas#1 266s / bas#2 263s / asm 135s / smoke 84s / syntax 23s / compile 17s）。
               # 上一条基线 = **#120**（head `3f1c4ab` = C29-8b 那一格，**8/8 job 全绿**，**wall 619s = 10m19s**；逐 job：Build 199s / vbp 412s / bas#1 256s / bas#2 243s / asm 95s / smoke 43s / syntax 21s / compile 17s。对照 #110 的 854s/544s：本批又往 vbp 组加了 15 条读数，still 更快）。
               # 上一条基线 = **#110**（head `a0546cd` = C29-5a 那批 `12394c2` + 合并 `github/dev@0cb1469`，**8/8 job 全绿**，**wall 854s = 14m14s**；逐 job：Build 190s、vbp 657s、bas#1 236s、bas#2 257s、asm 57s、smoke 44s、syntax 22s、compile 19s。vbp 那一格比 #96 的 544s 长，是因为本线三条控件判据（ctrltreeview / ctrltoolbar 各两臂 + 两件 emitc 断言）都往这组里加）。
               # 上一条基线 = **#103**（head `2339419`，C29-8a，wall 627s = 10m27s、vbp 492s）；再上一条 = **#96**（head `39ecaeb`，8/8 全绿，wall 730s = 12m10s，对照 #91 的 1159s = 19m19s：vbp 930→544s、bas #1 344→219s、bas #2 374→235s、asm 114→83s，compile/syntax 不动（那两组不构建，见 #78））。中间那笔 **#94 = failure** 不是回归，是本批换来的真证据：它抓出增量链接路缺 CRT 配平，已在 `39ecaeb` 修掉。
               # 上一条基线 = **#96**（head `39ecaeb`，8/8 全绿，wall 730s = 12m10s，对照 #91 的 1159s = 19m19s：vbp 930→544s、bas #1 344→219s、bas #2 374→235s、asm 114→83s，compile/syntax 不动（那两组不构建，见 #78））。中间那笔 **#94 = failure** 不是回归，是本批换来的真证据：它抓出增量链接路缺 CRT 配平，已在 `39ecaeb` 修掉。
               8/8 job 全绿 = Build C3.exe + Tests(smoke/syntax/vbp/compile/asm/bas#1/bas#2)，
               06:10→06:28 共 18.6 分钟；逐 job 用 `GET /runs/{id}/jobs` 核过）。
               **门只认 "VB6 C3 Regression" 那条 workflow 的 run 号**：从 `f7b1d2d` 这次推送起，dev 上
               多了另一位作者的 workflow "GitHub Tests T0+T1+T2"（它的 run #1 会同时被触发，编号体系
               完全不同，别把它当门 —— 本轮它有一条 T1 红，红因见上面 LAST_RUN 那条 5s 预算的读数）。
               本机同源读数：`-Category syntax` 129/0；30 件存量工程 `--emit-c` changed_lines=0；
               `tests\c29timer` 10 条 tick 区间读数 x64 与 x86 各 10/10；BASE 负控 10 条全翻红
               且每条对上症状（运行期开不起来=0、改 Interval 不生效=32、关掉还烧=16、精度地板=32）。
               精度对照（1 秒墙钟窗口的 tick 数，名义=1000/Interval）：Interval=20 → 改前 29 / 改后 49，
               =100 → 32 / 10，=5 → 32 / 198。
               基线相对上一版（run #86，head `e7f3352` = C29-9）新增的用面：`tmtimer[_x86]`。
               上一版基线 = run #86（C29-9）；再往前 = #85（C29-1b）、#84（C29-1a）。
```

> 重入保护：若运行开始时 STATUS=BUSY 且 LAST_RUN 距今不足 55 分钟，说明上一次运行可能仍在进行——本次**立即结束，不做任何修改**。

## 范围与验收硬边界（用户已确认，2026-09-23）

1. 深度：**含完整 COM 兼容**（P6 必做：IUnknown/IDispatch、类型库导出、DllGetClassObject/DllRegisterServer、CoCreateInstance/CreateObject 外部激活）。
2. 提交：每个批次过全量回归门后，仅暂存本批相关文件，commit 到当前分支（fan/dev），**永不 push**。
3. 收尾：全部批次完成后 STATUS=ALL_DONE；此后每次运行只复跑全量回归做只读验证并记一行日志，不改代码。

## 现状盘点（建表时已核实）

- VB6 式 `Implements` 已有：parser 语句收集（parser_module.cpp:119）、语义验证接口成员覆盖（semantic_analyzer.cpp:232-267）、`implementsNames`/`isInterface` 符号字段。
- COM 后端基础设施已存在：`src/backend/module/cgen_com.cpp`、`src/typelib/typelib_builder.*`、`src/com/typelib_parser_*`。
- `Inherits` / 显式 `Interface...End Interface` / `CoClass...End CoClass` 语句：**未实现**（grep 无语言层命中）。
- 语言风格先例：Delegate(a15c40b)、Overload(ea0591b)、Generics(7e2f9ed) 均为"tB式适配到本项目"，含 mangle 方案与护栏测试。

## 阶段计划（P1-P7，批次数可拆分但范围不得越界）

- **P0 设计细化**：结合 018 文档 + 现状，产出本项目语言映射设计并写入本文件"设计记录"节：Interface/CoClass 语句落在哪种模块宿主、成员 mangle 键（参照 $ov$/_g_ 先例）、vtable 内存布局与现有 cgen_com 的关系、GUID 来源（[InterfaceId]/[CoClassId] 或自动生成规则）、与既有 VB6 Implements 的共存策略、x86/x64 约束。设计完成后直接开工，不等人工批准（用户已授权自动化）。
- **P1 Interface 语句 + 编译期契约**：`Interface ... End Interface`（含 Extends 接口链、Sub/Function/Property Get/Let/Set 成员、无实现体检查）；Implements 完整性检查接入接口链；裸接口成员调用编译期解析。
- **P2 接口多态与生命周期**：接口变量 `Set`/`New`、按 vtable/表指针派发、引用计数与释放、接口↔类转换（TypeOf/隐式契约校验）。
- **P3 Inherits 类继承**：单继承、成员继承与遮蔽、Overridable/Overrides 虚钩子、Protected 可见性、MyBase 式显式基调用。
- **P4 Implements Via**：委托式实现（免手写转发），生成转调桩。
- **P5 CoClass 语句**：`CoClass ... End CoClass` + `[Default] Interface`、`[InterfaceId]/[CoClassId]`、组内 `New CoClass`/`CreateObject(ProgID)` 激活、默认接口派发。
- **P6 COM 兼容**：IUnknown 三件套运行时、IDispatch（GetIDsOfNames/Invoke）、类型库导出接入现有 typelib_builder、DllGetClassObject/DllRegisterServer/DllUnregisterServer/DllCanUnloadNow、外部（VB6/VBA/脚本）激活冒烟验收。
- **P7 终验收尾**：端到端示例工程 + 全量回归 + 文档（018 附录或新 0xx）+ STATUS=ALL_DONE。

每阶段要求：新增对应测试类别/用例；门 = `scripts/build.bat`（或 dev.ps1）构建 + `tests/run_tests.ps1 -Category all` 零新增失败；护栏参照先例（如零重载逐字节一致）。

## 批次清单（P0 细化后逐批登记，完成打勾）

| 批次 | 阶段 | 内容 | 状态 | Commit | Gate |
|---|---|---|---|---|---|
| B00 | P0 | 设计细化并写入本文件 | ☑ | 74ae1e7 | 纯文档批次，未构建（理由见运行日志） |
| B01 | P1 | 词法/AST/语法：`Interface…End Interface` 块 + `Extends` + 成员签名 + 无实现体检查 + 方括号属性行 | ☑ | 3add1ce | `Results: PASS=111 FAIL=0 SKIP=1 TOTAL=112`（gate_B01.log 全量）+ 7 条负例全绿 + x64/x86 双跑 + 8 文件（6 .bas/2 .vbp）emit-c 对无 B01 基线 exe 逐字节全同 |
| B02 | P1 | 语义：接口符号注册 + Extends 链 prepass(2.7) + 槽位表 + Implements 契约完整性/签名比对 | ☑ | beb75a7 | `Results: PASS=118 FAIL=0 SKIP=1 TOTAL=119`（gate_B02F.log；exe md5 13e99638 跑前后一致）+ 新增 7 条用例全绿 + legacy `test_implements` 仍 PASS。交付：stage 2.7 `runInterfacePrepass`/`IfaceRegistry`/Extends 链与槽表/五类诊断 + `checkNewStyleInterface` 严格契约比对（D15 记 1–6）。`SymbolKind::Interface` 按 D15-2 推迟到 B04。成员级子句拆给 B02b |
| B02b | P1 | 成员级 `Implements I.M[, I.N]` 尾子句（显式绑定优先于同名隐式匹配）+ 泛型模板内 Interface 的 3018 用例 | ☑ | dde7c32 | `Results: PASS=125 FAIL=0 SKIP=1 TOTAL=126`（gate_B02bF.log；exe md5 cdac040f 跑前后一致）+ 7 条新用例（itf_n14..n19 + itf_p02）全绿 + legacy `test_implements` 仍 PASS。交付：`parseTrailingImplementsClauses` + 三个过程节点的 `implementsClauses`（含克隆路径）+ 槽键兼容 `I.Name`/`I.get_Name` 与链上任一接口名 + `checkMemberImplementsClauses` 兜底（新 ID 3019）。详见 D16 |
| B03 | P1 | `.cls` 头行宿主形式 `Interface IFoo … End Interface`（1 文件 1 接口）+ `Module::isInterfaceModule` + 语法手册页/索引 | ☑ | c47cdce | `Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（gate_B03.log；exe md5 053150bf 跑前后一致）+ 3 条新用例（itf_p03 + itf_n20 + itf_xmod_writer）全绿 + legacy `test_implements` 仍 PASS。要点：宿主识别放 stage 2.7（parser 拿不到最终模块名）、VB3002 只豁免宿主自身、跨模块契约已端到端跑通（详见 D18） |
| B04 | P2 | 接口值代码生成：`vb6_ivtbl_<I>` COM 形态槽表 + 类侧实例 + 薄指针表示 + `As <Iface>` 变量登记 + 派发 | ☑ | 6bc97e8 | `Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（gate_B04.log；exe md5 ef1a7520 跑前后一致）+ 接口值派发端到端断言 IFV1/IFV2/IFV3 全绿 + legacy `test_implements` 仍 PASS。要点：新增 `cgen_iface_vtbl.cpp` 独立编译单元、`#ifndef VB6_IVTBL_<I>` 守卫替代 D19 设想的工程级去重表、`__iv_<I>` 紧跟 `__comObj`、槽键口径上提到 `interface_sig.hpp` 与语义层同源。B04a/B04b 合并成一批（理由见 D20-1）。逐字节 emit-c 护栏 8 文件对 pre-B04 基线全同 |
| B05 | P2 | 生命周期：实现类结构**前置** vtbl 指针数组（实测不可行，见 D21-1）+ refcount 头 + AddRef/Release + Set/Nothing/作用域释放 | ☑ | f644003 | `Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（gate_B05.log；exe md5 c615c657 跑前后一致）+ `itf_xmod_writer` 断言 5→11 条（LIFE1/2/3/9 + `TERM last=bye` + `TERM last=scoped`；实测该工程恰好 2 条 TERM，无误销毁、无重复释放）+ legacy `test_implements` 仍 PASS。要点：`__refcount` 只加在实现新式接口的类上（8 文件 emit-c 对 pre-B05 基线全同）；AddRef/Release 按 (类, 接口) 各一份；QI 仍占位到 B06；`__comObj` 非空时不归 0 销毁。详见 D21 |
| B06 | P2 | 转换与判定：接口↔类、多接口对象、`TypeOf … Is <接口>`、上/下行转换契约校验 | ☑ **B06a**（QI + 跨接口 Set + `TypeOf <接口变量> Is <接口>`）+ **B06b**（下行转换 + `TypeOf <类变量> Is <接口>` + 修 B05 的 Nothing 野地址）；**B06c 遗留**：接口值作实参 / 进 Variant → 随 B13/P6 处理 | 4dc6b7e | B06a：`Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（exe 7721bd72）+ 断言 11→18；B06b：`Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（gate_B06b.log；exe 4202570a 跑前后一致）+ 断言 18→25（DN0..DN3 + TOC1..TOC3）+ legacy `test_implements` 仍 PASS + 8 文件 emit-c 对 pre-B06b(@613d2b8) 全同 + **A/B 负控证明 D22-10 缺陷真实**（基线二进制 exit=139，本批 exit=0）。详见 D22/D23 |
| B07 | P3 | `Inherits` 语法 + 类链检测（单继承/环/深度）+ 继承成员合并与遮蔽 + 派生域 | ☑ **B07a**（语法 + stage 2.8 链检测/诊断 + 多文件负例通路）+ **B07b**（stage 3.4 成员合并 + 前缀布局 + 转发桩）；**B07 遗留**：裸名继承调用（要 `Me.`）、继承 `Public` 字段的 COM 对外暴露 → 分别归 B08+/P6 | b1c0050 | `Results: PASS=140 FAIL=0 SKIP=1 TOTAL=141`（gate_B07b.log；exe md5 50ce2e77 跑前后一致）+ `cls_inh` 三级链 12 条断言 + ci_n08/n09/n10 三条边界负例 + legacy `test_implements` 仍 PASS + 8 文件 emit-c 对 pre-B07b(@a056705) 全同。要点：并 8 张成员表（不是 11 张，理由在码内）、祖先私有字段**也复制进布局**、属性三向各一份桩、`_has_` 尾参必须转发、**封掉裸名继承调用的静默错代码**。详见 D27（B08 地图 = D28） |
| B08 | P3 | `Protected` 可见性 + `Overridable/Overrides/NotOverridable` + 类级虚表 `vb6_cvtbl_<Cls>` 与多态派发 | ☑ **B08a**（`Protected`：家族内经 `Me.` 可用，含跨 TU 与 Protected 字段）+ **B08b**（虚修饰符三件套语法 + `Overrides` 覆盖契约：槽键按方向配对、签名复用接口口径 + 把需要动态派发的调用点判死，避免静态绑回基类实现的假虚派发）+ **B08d**（类虚表 + 运行期真派发：3.4b 排每类有序虚槽、`const void* __cvtbl` 字段、表类型/实例/装载三点同源、两处类成员发码路按槽索引改写，并删掉 B08b 的 `Me.X` 拒绝）已交付；**B08c**（家族外访问 `Protected` 的拒绝，诊断 `VB3023`：判定落在 `visit(MemberAccessExpr)` = `obj.<成员>` 的唯一必经点；接收者→工程类靠新加的 `Symbol::srcTypeName`，认不出接收者或当前类未登记一律放过）已交付 → **B08 四条全出**（要点与踩坑见 D34） | 82b1b34(B08c)、df9806e(B08d)、05397be(B08b)、2117d1c(B08a) | `Results: PASS=149 FAIL=0 SKIP=1 TOTAL=150`（gate_B08d_v3.log；exe md5 83c4e49c 跑前后一致）+ `-Category syntax` 73→74（删 ci_n15、增 ci_n19/ci_n20）+ `Inh.vbp` 运行期断言 17→27 条（INH17..23 派发：`b/m/d.PickThru()` 分别 base/mid/mid = 绑最近覆盖者、INH20 叶类覆盖被基类体内看见、INH23 基类型变量持有派生实例不再切片；INH24..26 扇出）+ 8 文件 emit-c 对 pre-B08d(@fb6a254) **8/8 逐字节全同**。更早两轮证据：B08b = `148/0/1/149`（gate_B08b.log、exe 546265e7、syntax 66→73、断言 14→17）。要点：`ProcVirt` 四值枚举而非三 bool；契约检查落 2.8、槽表落 3.4b（3.4 之后分不清"谁声明的"，而"本类有没有入口"要读 3.4 的 inhProcs）；筛选集取**链根**的 dynamicKeys → 叶类也带字段；`Me.X` 不在"优先级2"那一批发码。详见 D31、D33（B08d 地图 = D32，其中 ②③ 已被 D33 修正） |
| B08e | P3 | 虚表线收尾：`resolveClassMemberCall` 其余 13 个消费点逐条接上派发或判死（站点地图与裁决见 D35）——**13 站已在 B08e-6 全部出完** | ☑ **B08e-1**（① `With w` 内 `.M()`）、**B08e-2**（⑤ `Me.<字段>.方法()`）、**B08e-3**（零代码：⑩/⑪ 裁决纠偏 + 找出先决条件⑭）、**B08e-4**（⑭ `suppressVirtDispatch_` + ⑪ `Me.<字段>.<属性>` 的读）、**B08e-5**（⑥⑦ 默认属性调用式 `m_up(9)`）、**B08e-6**（⑨ 接派发 + ⑩⑫ 判死 + `Test-CompileFail` 前置；⑧ 实测已被优先级2 判死、⑬ 改判无需改）已出；下一轮 **⑮**（UDT 字段调用位置的路由缺陷，D35-8）；②③ 判为探针无需改、④ 判为走不到 → **13 站在 B08e-6 后全部出完，⑮ 一出就开 B09（`MyBase`）** | e531d82(B08e-1)、8987386(B08e-2)、392a52d(B08e-4)、d9eca95(B08e-5)、40eea3f(B08e-6) | 最新门 = B08e-6（Actions 级，见状态头 GATE_BASELINE）+ `Inh.vbp` 断言 38→40（INH39 判别/INH40 对照）+ A/B（`pre_b08e6_C3.exe`：INH39 FAIL base、其余 39 条 OK；`ci_n24`/`ci_n25` 改前 `--emit-c` 退出码 0）+ 8 文件 emit-c 8/8 + `-Category syntax` 78→82 全绿。上一条本地门 = B08e-5：`Results: PASS=153 FAIL=0 SKIP=1 TOTAL=154`（gate_B08e5.log；exe md5 ac14cf25 跑前跑后一致）。上一批 B08e-4：同 153/0/1/154（gate_B08e4.log；exe f2547af6）。B08e-3 未跑门（零代码，先例 B00）。 |
| B08f | P3 | UDT 里放工程类对象字段的整条通路（⑮a 写方向 / ⑮b 调用方向 / ⑮c 派发 / ⑮d UDT 自己在第三个模块），实测与落点见 **D36**、真根因与实施见 **D37**；原挂在 B08e 的"站点⑮"，因与虚表无共同判据而单列 | ◐ **B08f-1 已出（`77ecef1`）= ⑮a+⑮b+⑮c 一起**（三处接线：语义层 Variant 分支也存类型名 / `udtFieldObjCType` 按当前符号表回判工程类 / Set 侧按真实 C 类型否决 Variant 容器判定；通路一通，D35 站点④ 才第一次可达，顺手接上派发）。**⑮d/⑮e 仍开**：UDT 声明在使用点之外的模块时同样坏；属性写穿过 UDT 字段两条路都发非法 C，卡在 `symTab_.lookupModule(UDT 名)` 的消费者可见性 —— 与本批不同判据，另批先量可见面再动 | 77ecef1 | 验收只能走**真编译**（`--emit-c` 对坏形状返回 0）：`Inh.vbp` 断言 40→43（INH41/42 判别、INH43 对照），改码前的二进制编不过这个工程；8 文件 emit-c 8/8、`-Category syntax` 82/0 |
| B09 | P3 | `MyBase.M(…)` 显式基调用（去虚化）+ 构造链顺序 + 无新语法逐字节护栏；前置 = `com_entry` 的 typedef 分块顺序（D35-9 ⑤） | ☑ **B09**（`02bac92`，记录见 **D38**）：`MyBase` 走发码层按接收者名字接管（precheck 的 `Err`/`VBA` 先例，lexer/parser 未动）；实现按“就近声明”取 `vb6_<owner>_<M>`，**不查 `__cvtbl`**；读 `Get>Function>Sub>Let>Set`、写只认 `Let/Set`；`VB3028` 判死三条形（无 Inherits / 基面无此名 / 基类 Private = C 层 static）；构造链根→叶 + `vb6_<基>_chain_init` 桥（仅“写了 Class_Initialize 且被谁继承”才发）。B09-0：`com_entry` 的类返回类型 extern 改 `struct vb6_cls_X*`。**仍开（另批）**：**B09b = x86 布局缺陷**（继承的 Private UDT 字段在派生 TU 退化成 `void*` → 前缀错位、`_New` 少分配 4 字节 → 写 `m_pt` 之后的基类字段越界；x64 巧合正确，所以门里看不见，实测见 **D39**）、`Class_Terminate` 反序链、`Set MyBase.<属性> = obj` 未验、基类自家 extern 的 `void*` 返回类型（Fix 184 只做了一半） | `02bac92` | `Inh.vbp` 断言 **43→52**（INH44/45 判别去虚化、INH49/50 构造链与桥、INH51 = B09-0 用例）；**改码前编不过这个工程**（C2143 + 5×C2065）；负例 `ci_n27`/`ci_n28` 走 `--emit-c` 断言 VB3028（改前退出码 0）；8 文件 emit-c 8/8、`-Category syntax` 82→84 全绿 |
| B09b | P3 | x86 布局缺陷：继承字段里 `Private` UDT 在派生 TU 解不出类型名 → 回落 `void*` → 前缀错位（D39 登记，本批治） | ☑ **B09b**（`debb110`，记录见 **D40**）：`runCrossModuleResolution()` 末尾沿 `Inherits` 链注入被继承字段用到的 UDT/Enum 类型符号（含 `udtMembers`；本地同名不抢）；判别实验 = 把 `Private Type` 改成 `Public` 看发射形状 + x86 能否跑起来；顺带清掉从 B07b 挂着的 INH35/INH36/INH39 三条 x86 红字；**门新增 `cls_inh_x86`**（同一份断言清单双架构跑） | `debb110` | `tests/cls_inh` x64 + x86 build+run 各 **52/52、0 FAIL**（改码前 x86 段错误）；8 文件 emit-c 对 `pre_b09_C3.exe` 8/8；`-Category syntax` 84/0 |
| B09c | P3 | 收尾 P3 三条小尾巴：`Set MyBase.<属性> = obj`、⑮e（属性写穿过 UDT 对象字段）、`Class_Terminate` 反序链 | ☑ **B09c**（`9eb2ca7`，三条的实测与处置见 **D41**）：① 需要改码并已修 （改码前只有接收者是裸 `MyBase` → C2065）；② **⑮e 早在 B08f-1 就通了**，D37 那条是在坏元数据上量的 → 只补断言（INH55/INH59）；③ **`Class_Terminate` 链不做** —— EXE 工程三种形状实测都不触发 terminate，发出去就是死代码 → 登记为生命周期/P6 的既有缺口 | `9eb2ca7` | `Inh.vbp` 断言 **52→59**、**x64 + x86 各 59/59、0 FAIL**（A/B：改码前编不过，1×C2065）；8 文件 emit-c 对 `pre_b09c_C3.exe` 8/8；`-Category syntax` 84/0；INH58 = B09b 最深用例（隔两级持有根的 Private UDT 字段）|
| B10 | P4 | `Implements IFace Via <holderVar>` 委托式实现：持有字段 + 自动转调桩 + 签名检查 | ☑ **B10**（`3c5d8e6`，实施与选型见 **D43**）：`Via` 软关键字四件套 → 语法 `ImplementsStmt::viaField` → **stage 2.7 Pass D** 裁决（`vias_`，一份来源同喂语义与发码；`VB3029`/`VB3030` 两类判死含链式 Via）→ 语义层免逐槽 `VB3012` → 发码层转调**持有对象的接口槽**（不直调 Private 成员：C 层 static 跨 TU 连不到）+ Nothing 字段退零值 | `3c5d8e6` | `tests/itf_via` 新工程 **VIA1..VIA9 x64 与 x86 各 9/9**；A/B 改码前 `VB2003`+`VB2002`；4 条负例（n21/n22/n23/n24）+ 1 条软关键字正例 p04；10 文件 `--emit-c` 对 `pre_b10_C3.exe` 全同 10/10；`-Category syntax` 84→89 |
| B11 | P5 | `CoClass…End CoClass` 语法 + `[CoClassId]/[Default] Interface/[ComCreatable]/[CoClassCustomConstructor]` + 契约聚合校验（实施依据换成 `ai/026`，其六节把 B11 拆成 C01–C05） | ◐ **C01 已出（`e7c7a31`）= 块语法 + 属性行落 AST**（零回归靠 `Module::coclasses` 不进 `declarations`；属性行归属按"名字+位置+同行"合判，见 D44/D45）；**C02 已出（`f0b820d`）= 身份求解唯一函数**（`src/semantics/coclass_identity.{hpp,cpp}` 纯函数 + stage 2.7 Pass E + `Driver::coclassIds_`；见 D46/D47）；**C03a 已出（`e515d89`）= 块形状与名字校验**（stage 2.7 Pass F 八条判据 + `VB3031/3032/3033`；宿主同名豁免见 D49-①）；**C03b 已出（`c4aaa4c`）= 契约聚合校验**（新 **stage 3.4c** `runCoClassContractCheck()`：按链倒着走、叶优先，缺槽 `VB3012`/签名不符 `VB3017`，`vias_` 命中的委托免逐槽；`VB3020` 文案收口。`As <CoClass>` 按 D50-② 整条并入 C05）。**C04 已出 = 存量头属性只读折算**（stage 2.7 新 **Pass E0**：四行 `Attribute VB_*` → 一条 `CoClassDecl` 进同一张表，Pass E 仍是唯一身份出口；折算记录**不参与判死**（Pass F/3.4c 跳过它），手写块优先，见 D52/D53）；**C05 已出（`7f829ee`）= 组内激活**（stage 2.7 新 **Pass G**：类型位点上的块名**就地改绑**`[Implementation]` 类 + `CreateObject(ProgID)` 换成 `New`，`VB3039` 挡住没有实现类的块名当类型用；实测 D54、实施 D55。发码侧零新分支，护栏 16/16。这一格同时把 022 的 **B12** 一起交付） | `7f829ee`(C05)、`c4aaa4c`(C03b)、`e515d89`(C03a)、`f0b820d`(C02)、`e7c7a31`(C01) | run **#23**（head 已核 = `4045b42`）+ `-Category syntax` 107→112 + 10 文件 `--emit-c` 对 `pre_b11c03b_C3.exe` 10/10 + A/B（n37/n38 零命中、n39 出旧句）+ `Id.vbp` 真编译真运行且发码逐字节未动 |
| B12 | P5 | 组内激活：`New <CoClass>` / `CreateObject("ProgID")` 编译期映射 + 默认接口派发 | ☑ **由 B11/C05 交付**（`7f829ee`）= Pass G 就地改绑实现类，`As`/`New`/`CreateObject` 三面同归一条工程类路；`As Object` 目标经 Fix 179a 拿到 IDispatch 包装。**口径偏差**：`As <块名>` 的成员面是实现类的公开成员（比默认接口宽），要窄视图写 `As IShape` —— 两条理由记 D54-⑤ | 实测 D54 / 实施 D55 | `-Category syntax` 116→118 + `tests/cc_act` CC1..CC9 **x64+x86 各 9/9** + 护栏 16/16 + A/B 三条坏读数（base 复现、new 消失）；门见状态头 |
| B13 | P6 | ~~IUnknown 三件套真实实现~~ **D56 推翻**：RTL 的 `ComObj_QueryInterface/AddRef/Release` 早是真实现（原子计数、归零销毁、`Class_Terminate` 在 DLL 侧有触发点）⇒ 本格真正的活 = **把 CoClass 块接进对外那一半**（身份出口、块名 ProgID、新式接口对外可调用） | ☑ **B13a（观测面）**：新助手 `Test-VbpDll` + 把现成的 `tests\test_activex_dll` 两份 DLL 工程接进回归（此前门内一次都没链接过 .dll）+ 新工程 `tests\cc_dll`；**B13b 已出**（`dll_entry` 的 CLSID / 默认接口 IID / 组名 ProgID 一律读 `coclassIds_`；`[ComCreatable(True)]` = 组名那一档 ProgID 的唯一开关，`legacyFolded` 继续走原路 = 折算隔离；legacy lambda 降级成「读不到块才用」的兜底，未并掉的两枚见 D57-4/5）、**B13c 已出**（D58）= 接口自己的 IID 也进唯一出口（`resolveIfaceIid` 与块内 `[Default]` 共用一个函数；Pass E 建 `Driver::ifaceIds_`），COM 服务器表 / 新式接口 vtable 的 QI / 类型库三条通道一律读这张表（BASE 实测同一个 `IProbe` 有**四枚** GUID、且 `.tlb` 广告给客户端的那枚服务器不应答 ⇒ D57-5 那条"未证"当场证伪并修好）；类型库的 coclass 默认接口开始跟随块的 `[Default]`（只延后需要重定向的那几行 ⇒ 没有手写块的工程连类型顺序都不动）；D56-5 的对外口径**拍板走 (b)**：`Private` 契约成员不发成 disp id，改打一条 `C3: …vtable interface: 0 Public member…` 信息行说明"注册了却点不到"（D58-5）；新增 `.tlb` 读数工具 `tests\tools\tlbprobe.cpp` + gated 用例 `cc_dll_tlb_matches_table`（助手在 `tests\tlb_identity.ps1`）。**B13d 已出**（D59）= 薄指针的规范 IUnknown：`vb6_iunk_<C>_<I>_QueryInterface` 不再把 `IID_IUnknown` 与本接口并成一个分支回 `self`（两个不同偏移的 `__iv_` 成员 ⇒ 同一对象两个身份），一律回**本类实现序第一个接口**的薄指针 —— 那是唯一一处偏移 0 就是 vtable、既能当身份又能被再次 QI/AddRef/Release 的合法指针，且零布局改动（D19 不入视野）、单接口类拿到的值与改前相同。用例两条：`itf_canonical_iunknown`（`--emit-c` 断形状 + 负控跑过 BASE）与真编译真跑的 `itf_xmod_writer`（QI1..QI4）。**本批最值钱的是顺手读出的一条危险**（D59-4）：RTL 包装器对表里的 `defaultIfaceIid`/`ifaceIids` 一律交回胖指针，而这两处自 B13b/c 起可能装着新式 vtable 接口的 IID 且 `.tlb` 把它当默认接口广告 ⇒ 早绑定客户 QI 成功后按虚表第 3 槽调用会打到 `GetTypeInfoCount`（静默调错函数）。两种收法（对外不发布 / 让包装器真返回薄指针 = B16）等 **B13e** 拍板 —— 本批刻意不动，因为它要把上一批刚立的用例翻面。**B13e 已出**（D60）= 对外默认接口的口径拍板 + 回退：**广告的那一枚必须就是应答的那一枚**。先证伪 D59-4 那条危险（`TypeLibBuilder` 只有 `TKIND_DISPATCH`/`TKIND_COCLASS`，客户从库里学不到接口虚表布局 ⇒ "第 3 槽打到 `GetTypeInfoCount`"不成立），再收同一次读码撞见的真问题：B13c 把 `.tlb` 的 coclass DEFAULT 引用重定向到 `_IProbe`（库里 0 成员），而服务器 `GetIDsOfNames`/`Invoke` 认的是类的公有成员那一档（`_CImpl`）⇒ 两边点不到。三处一起退回（表的 `defaultIfaceIid` 改由类型库回写值说话、`ifaceIids` 里 `[Default]` 同名接口那条覆盖分支删掉、类型库的 DEFAULT 重定向连同延后登记机器拆掉），coclass 的 CLSID 与接口自己的 IID **仍**取唯一出口 ⇒ 甲判据不动；`(i)` 的全量形态（滤掉 `ifaceIids` 里新式那几项）刻意没走，那会让表与 vtable 劈成两枚 GUID，正解归 B16（D60-4）。用例两条翻面/升级：`cc_dll_identity_single_source`（`IID_vb6def_CImpl` 回 `0x7CA8CD81`）、`cc_dll_tlb_matches_table`（甲 + 乙两条判据，负控在 BASE 上被乙判红）；那条 `…is a vtable interface…` 信息行随前提一起删除。**B13 五格到此出完（a/b/c/d/e）** | `9785f4f`(B13a)、`b1a8e58`(B13b)、`bd38798`(B13c)、`9df23ba`(B13d)、`b10ec1a`(B13e) |
| B14 | P6 | IDispatch 四件套接入新式接口（GetTypeInfo/GetIDsOfNames/Invoke + DispId 表） | ☑ **B14 已出，但范围被自己的测量重裁（D61）**：三件测量先行 —— ① **四件套本身是通的**（新探针 `tests\tools\disp_probe.c`：`LoadLibrary` + `DllGetClassObject` + `CreateInstance(IID_IDispatch)`，不查注册表；对真产物 `TestAXDLL.dll` 实测 `Add(2,40)=42`、`SetValue(7)`→`GetValue()=7`、`GetTypeInfo(0)` 回 coclass 那份（`kind=5`、`cFuncs`=0 是 coclass 常态）、未知名 `DISP_E_UNKNOWNNAME`）⇒ **缺的不是四件套，是成员面**；② 薄指针 `vb6_ivtbl_<I>` = `{QI,AddRef,Release,自有槽}`，**没有 IDispatch 那四槽** ⇒ 要接就是 dual/布局改动，且**必须排在 B15 之后**（先动它就重演 D60 的"广告 != 应答"）；③ `comDispid` 全工程只有一处生产者（`driver_codegen_dll_typelib.inc:186`，只收类模块 Public 成员）⇒ **契约成员从来不在这张表里**，要发只有"违口径 (b) 当公有发"（两边假绿）或"真接口 = B15"两条。本批交付 = 把 DLL 这条管线从"字节一致"升到"真调用得通"：探针 + 助手 `Test-DispatchInvoke`（`tests\disp_invoke.ps1`）+ 两条 gated 用例 `ax_dll_dispatch_invoke`（legacy 面真点通）与 `cc_dll_dispatch_iface_only`（只满足新式接口的类：**成员面必须为空** ⇒ `NAMES=ADD hr=0x80020006`；并把"接口 IID 由胖指针应答 `same=yes`"钉成实测事实 ⇒ B16 要改必须故意翻它）。**零编译器改动** ⇒ 产物逐字节不变（收线 exe md5 与开工同一枚），负控 = 换成不存在的 CLSID ⇒ `GETFACTORY hr=0x80040111` 判红（新助手先证明能红）。另登记一条没动的弱点（D61-6）：`ComObj_Invoke` 在 invkind 配不上时退回"第一个同 dispid 表项"，而 VB6 里 Get/Let 共享 dispid 是常态 ⇒ 收紧属口径题，等拍板 | `ddf4e9b` | |
| B15 | P6 | 类型库导出：新式接口在库里发成真接口（`TKIND_INTERFACE`）+ 修"接口宿主被登记成假 coclass"的形状 | ☑ **B15 已出（代码 `941b6dc`，门 head = 合并 `0caa3c5`，Actions run #59 全绿）＝ 只交付"形状与身份那一半"**，成员面按读数刻意押到 B16（理由 D62-1 mode 8 + D62-3，见 CURRENT_BATCH 第 1 条①③）。四件测量落 D62，其中两条推翻既有判断：**① 接口成员的签名从来没到发码现场**（026/B02 口径：Interface 块只进 `Module::interfaces` 不进 `declarations` ⇒ 宿主 Class 符号 `memberNames` 恒空 ⇒ `_IProbe` 的 `cFuncs=0` 是这条路本来不通，D61-3 那条"DispId 表缺一处生产者"的根因订正）；**② `CreateTypeLib2` 的位数 flag 有字节后果**（同一份 `oVft=24`，`SYS_WIN64` 建库读回 24、`SYS_WIN32` 读回 48，两枚 flag 出的 `.tlb` md5 不同）⇒ 本批不动它，登记成 B16 的前置裁决。成立条件实测：`TKIND_INTERFACE` 拒 `FUNC_DISPATCH`（`0x800288BD`）、要 `FUNC_PUREVIRTUAL` + `oVft`；`VT_PTR` 不给 `lptdesc` 会让建库进程当场崩；0 槽真接口可建（mode 9）；coclass 引用真接口并标 DEFAULT 走得通（mode 5）。交付：`TypeLibBuilder::addVtableInterface`（新）+ `driver_codegen_dll_typelib.inc` 那条循环加接口宿主分叉（判据 = Pass E 命中且 `ifaces_` 有同名键，legacy 那条一行不动）⇒ 库里 `coclass IProbe`（一枚 `generateUuid(类名)` 另 mint 的假 CLSID + 一句"可创建"）与 0 成员 `_IProbe` 一并撤掉，换成 `kind=interface name=IProbe` 用 Pass E 那枚 IID；types 4→3、1648→1484 字节。判据换读法不放宽：`Test-TlbIdentitySingleSource` 通道 3 改读真接口行 + 两条**反面**断言（那两行再现即红），负控实测 old=RED(三条全中)/new=GREEN。护栏：16 件 `--emit-c` 逐字节 0 变化、全产物 A/B 只有 `cc_dll` 的 `.tlb` 变、两件存量 DLL 的库一字未动。根因一句话：那条循环是唯一没有 `isInterface` 过滤的 COM 消费者（对照 `cgen_util_dllentry_collect.inc:26`、`driver_codegen_dll_sync.inc:8`）。 | `941b6dc` | 门 #59（8 job 全绿，`Tests (vbp)` 33/0/1）+ `b15_emitc_guard.py` 16/16 + `b15_ab_all.py` 分类 0 + `b15_negctl.ps1` 负控 + `b15_measure1/8.ps1` 七种 mode 建库实测 |
| B16 | P6 | **D62 重裁后的真身**：接口成员的对外调用契约（`vb6_ivtbl_<I>` canonical 化 → 库里那一档发成员）+ 类型库位数 flag/oVft 口径 + 薄指针 IDispatch 面与"胖应答瘦"的收法 + DllRegisterServer 一族对新式 CoClass/类工厂的接线与 x86/x64 双验 | ☑ **B16 已出（代码 `578faa4`，Actions run #62 全绿）** = ① **薄面整条 canonical 化**：`vb6_ivtbl_<I>` 的 IUnknown 前缀 + 契约槽 + 它们的实现函数在 x86 下一律 `__stdcall`（BASE 的 x86 产物 `grep -c __stdcall` = 0，NEW = 21；x64 上 MSVC 忽略该修饰 ⇒ 一份生产码两架构通用）；② **类型库那一档如实发契约成员**：`cFuncs` = 2（`Ping`/`Got`）、每成员 `oVft=(3+槽)*指针宽` / `callconv=stdcall` / 原生返回 vt、ByRef 建 `VT_PTR` 链；③ **位数 flag 跟 `--arch` 走**（`CreateTypeLib2(is64_ ? SYS_WIN64 : SYS_WIN32)`）—— 存量库默认架构逐字节不变、`--arch x86` 只差 32 个字节的位数布局字，x64/x86 两枚读端读数逐项相同（D63-3）；④ **薄面出入口** `vb6_iv_thin_<C>`/`vb6_iv_claim_<C>`：接口 IID 的 QI 交**薄指针**（B14 那条 `QI_EXTRA same=yes` 故意翻面）、包装器把底座引用交还**最后一个薄引用**（类工厂那种"QI 完就 Release 包装器"的规范姿势下交出去的指针不悬空）；⑤ **行里第 4 项（DllRegisterServer 一族接线）经读码 + A/B 判定无需新改动** —— 注册写入（`vb6comserver.c:62-171`）读的就是同一张服务器表、身份自 B13b 起走 `coclassIds_` 唯一出口，B16 的 A/B 里 `.def`/exports/`.rc` 一字未动；**真注册的外部端到端验收归 B17**（D63 与 B17 的 CURRENT_BATCH 都写明）。判据五条全对上读数：`cFuncs`/`oVft`/`callconv` 三样读数（新探针 `tests\tools\tlb_slots.cpp` 升格入库 + 两条 x64/x86 契约用例）+ x86 真跑 `VTBL_GET_AFTER=42` + 存量逐字节（17 件 `violations=0` / A/B `unclassified=0`）；canonical 返回形状与跨世界身份合一按"断不了先断形状"留 B17（D63-5，取舍看 B17 测量②）。 | `578faa4` | 门 #62（8 job 全绿）+ `b16_emitc_guard.py` 17 件 `violations=0` + `b16_ab_all.py` `unclassified=0`（只有 `cc_dll` 的 `.tlb` 1484→1624）+ 负控 `b16_negctl.ps1`（`pass=0 fail=2`，四条 needle 全 miss + 反面断言命中）+ 五条新用例 `b16_cases.ps1`（`pass=5 fail=0`）+ 迁移探针（默认架构 diffbytes=0 / x86 diffbytes=32） |
| B17 | P6 | 外部激活冒烟验收（CoCreateInstance 早绑定 + CreateObject/IDispatch 晚绑定 双路） | ☑ **B17 已出（代码 `a5517fa` + 门后修一条 `732c1f8`，Actions run #66 全绿）** = **对外那条管线第一次被真注册客户走通**，路上修掉两处一直存在的**静默**缺陷：① **`rc.exe` 的发现面太窄**（`driver_link.cpp`）—— 旧写法只认 `WindowsSdkDir` 环境变量与 `C:\Program Files (x86)\Windows Kits\10\bin`，SDK 装在别的盘（本机 = `D:\Windows Kits\10`，用户确认）就**静默不嵌**资源：实测 BASE 产出的 DLL 连 `.rsrc` 段都没有（数据目录 2 = `(0,0)`，NEW = `(4880)`/`(2504)`）⇒ 注册表里没有 TypeLib 项、外部工具按 LIBID 找不到库；改成与套件同一套探测（环境变量 → Program Files → 盘符 → PATH，目录内取版本号最大者，TypeLib 与版本信息两处共用），且"找不到 rc.exe"不再藏在 `--verbose` 后面；② **`vb6_UnregisterTypeLib` 的 `UnRegisterTypeLib` 实参顺序写反**（原型 `(libID, wVerMajor, wVerMinor, lcid, syskind)`，旧写法把 syskind 塞进 lcid 槽）⇒ 两个键都找错、函数失败，而本函数无条件返回 S_OK ⇒ **每次反注册都静默漏掉整棵 TypeLib 键**（实测：CLSID/ProgID 清干净、`TypeLib\{libid}` 还在）；按原型给对顺序并取库自己声明的 lcid/syskind（中性库记在 0x0409 下；x64 库 `syskind=3`、x86 库 `=1`，`.build/b17_tlbattr` 直接读 TLIBATTR 得到）。判据五条全对上读数：① 外部三条验收 —— `cc_dll_external_activate[_x86]`（注册四项 + `PROGID_LOOKUP same=yes` + `COCREATE_DISP ptr=OK` + `CALL=Twice result=42` + `COCREATE_IFACE` + `VTBL_GET_AFTER=42` + `CLEAN_*` 三类键 `gone`）与 `cc_dll_late_client`（**另一个进程**的 C3 客户 EXE：不引用 DLL，`CreateObject` + 晚绑定调用，`EXT1:OK`/`EXT2:OK`）；② 早绑定按库里 `oVft=(3+槽)*指针宽` 直调得 42，x86 + x64 双验；③ 注册表可清理（反注册后 CLSID/ProgID/TypeLib 三类键都不留，用例可重复跑 —— 修缺陷②之后才成立）；④ 存量行为不变（注册写同一张服务器表；`--emit-c` 17 件逐字节全同，全产物 A/B 只差 `rtl/vb6comserver.c` 的 10 行、全在反注册那段）；⑤ canonical 返回形状与跨世界身份合一按测量②**不做**。**晚绑定的边界**：契约成员是 `Private` ⇒ 类的默认面上按名点不到（`NAMES=Ping hr=0x80020006`，VB6 语义的应有读数，正面钉住）；为此 `CImpl.cls` 加了一个公有成员 `Twice`（B13b 那条 `methodCount` 针 0→1），于是同一类**两个面**并存：默认面走 `IDispatch` 晚绑定、接口面走契约槽早绑定。 | `a5517fa` + `732c1f8` | 门 #66（见 GATE_BASELINE）+ `.build/b17_ab_all.py`（A 段 `--emit-c` 17 件逐字节全同；B 段只差 `rtl/vb6comserver.c` 10 行；C 段资源目录：NEW 带不带 `WindowsSdkDir` 都嵌、BASE 不带就不嵌）+ 负控 `.build/b17_negctl.ps1` 喂 `pre_b17_C3.exe` ⇒ **pass=0 fail=2**（红的正是 `CLEAN_TYPELIB=STILL`）+ 三条新用例 `.build/b17_cases.ps1` ⇒ `pass=3 fail=0`（x64/x86 冒烟 + 外部客户）+ 门后补一条**既有潜伏缺陷**（`732c1f8`）：第一次门在既有用例 `ax_dll_dispatch_invoke` 上报 `exit=0xC0000374`（堆损坏），定位到 `ComObj_Invoke` 的 `coercedArgs`（`CoTaskMemAlloc` **不零**）收尾对每个元素 `VariantClear`，“原样传下去”那一支从没 `VariantInit` 过；A/B（把该数组填成 vt=VT_BSTR + 野指针的临时编译器）**恰好停在同一读数处**、清零版整趟干净；修为分配后整段清零（本机 12 次不复现 = 随堆状态偶发，非本批引入；同一次旧门的重跑又是绿的，也是这条的侧面证据） |
| B18 | P7 | 端到端示例工程 + 全量回归 + 设计文档归档（018 附录或新 023）+ STATUS=ALL_DONE | ☑ **B18 已出（代码 `0413bb9`，Actions run #68 全绿）—— 本线收口** = ① **端到端示例 `tests\cc_demo\`**：同一份源集合编两种形态（VB6 常规做法）——`DemoExe.vbp` 语言侧 `DEMO1..DEMO12`（组名当类型 `As Shape`/`New Shape`、契约成员**只能经接口变量**调、公有成员直调、`Overrides` 虚派发、`MyBase` 去虚化、继承来的公有成员、`Protected` 家族内可用、基类型变量持有派生实例不切片、`Via` 三个槽全转发、工程内 `CreateObject(ProgID)` 改写、`TypeOf`），`DemoDll.vbp` 对外侧由**另一个进程**的 C3 客户（`DemoClient.vbp`，不引用 DLL）`CreateObject("DemoDll.Shape")` 激活；② **归档 `ai\027-接口继承CoClass实施收口.md`**（018 保持原样）：交付总览表、语言/COM 两侧要点、**v1 边界 12 条**（逐条现状+影响）、怎么验（用例名/命令/探针与助手清单）、记录索引；手册 `CoClass 语句.md` 加示例指引；③ **写示例撞到并记下**两条 v1 边界（派生类自己 `Implements` 新式接口 = `VB3022`；EXE 的 CoClass 块不能 `[ComCreatable]` = `VB3033`）与一条易误读现象（`New <名字>` 不在工程内 ⇒ 静默按注册表创建、运行期 429 ⇒ `.vbp` 的 `Class=`/`Module=` 清单是唯一事实面），都进 **D65**；④ 四条新用例进 vbp 回归（`cc_demo_exe[_x86]` 走 `Test-Vbp`；`cc_demo_dll_external[_x86]` 走 `Test-ComActivateClient`，**注册与反注册成对**并断言三类键 `gone`，本机跑完实测注册表查无残留）；⑤ **本批零编译器改动**（`.build/C3.exe` 与 BASE 同 md5）⇒ 逐字节/A-B 护栏由"编译器没变"这条代替（先例 B14）。 | `0413bb9` | 门 #68（8 job 全绿）+ `.build/b18_cases.ps1` 四条 `pass=4 fail=0` + 注册表残留核查（DemoDll/CoDll 的 CLSID/ProgID/TypeLib 全 clean） |
| B19 | 维护 | 用户报障：`cmd` 里 C3 的报错信息与运行输出乱码（四语种通用：中/日/韩/英；cmd / PS 5.1 / pwsh 7 三 shell）+ CI bas#1 的 "run timeout 5s" 假红 | ☑ **B19 已出（代码 `796220d` + `d4e53c2`，套件侧 `8174219`，Actions run #73 全绿）** = ① **控制台/管道两套出口按句柄类型选路**——控制台（含 ConPTY）走 `WriteConsoleW`（与 `chcp` 无关）、管道/文件按**当时的控制台代码页**写字节、某行装不下就整行退 UTF-8（`lpUsedDefaultChar` 当判据），**全程不碰用户的代码页**；② 未处理运行期错误的出口按「有没有可写 std 句柄」判（旧判据 `GetConsoleWindow()` 在重定向下会把批处理卡在模态框上）；③ 三条用例进 vbp：读屏探针 `tests\tools\con_capture.c` 跑 x64/x86 两档 + `chcp 936` 重定向按**字节**断言；④ 套件运行预算 5s→`-RunTimeoutSec`（默认 60s）+ 超时读数（CPU 时间/进程状态/最后一行输出）。详见 **D66**/**D67** |
| B20 | 语言扩展 | **ai/028 两批一起发货**：反引号原始多行串（V1）+ 串内插值 `${expr}` / `${expr:fmt}`（V2）—— 需求出自 `todo/vi.md`（群友 Fan XiaoLei 2026-09-15），口径 = 现代语法糖，`&` 与普通 `"..."` 一字不动 | ☑ **B20 已出（V1 = `4c16f50` + 夹具钉字节 `f8ec76e`，V2 = `c2f7317`；门 = Actions run #82，head `c2f7317`，8/8 job 全绿）** = ① **一切落在词法层**：无孔的串在 `scanRawString` 出口折成与手写 `"..."` 逐字节同形的普通 `StringLiteral`；有孔的串展开成普通 token 链 `( "文本" & CStr( expr ) & ... )`（格式段 = `Format$(expr,"fmt")`）⇒ **parser / AST / semantics / cgen / driver / RTL 一字未动**，AST 里不存在"第二种字符串"（计划书 R4：`rawText` 全仓 57 处消费、至少四处各自剥引号折 `""`，形态位那条路每新增一处消费者就会漏一次）；② 语法口径：行界一律读成 CRLF（`source_manager.cpp` 已把源文件行尾抹成 LF ⇒ 源码行尾风格不可能影响串值）、起始反引号后紧跟的那一个换行裁掉、零转义、串内 `"` 原样、反引号双写、`$${` 是字面 `${` 的出口、裸 `{`/`}` 与冒号是文本（JSON 安全）；③ R4 四个落点各打一条真判据：`Const` 值 / `Declare … Lib` 的库名（DI 桩按 Lib 串选家族，折错直接 LNK2019）/ 工程内 `CreateObject` 的 ProgID（实测 1 次改写）/ 模块头 `Attribute`（折错则模块名对不上 `.vbp`）；④ "降级为真"有发码级证据：产物里是 `vb6_BSTR_Concat(vb6_BSTR_FromStr(L"n="), vb6_CStrLong(n))` —— 连"按实参类型改发专用 CStr"都一起继承；孔内未声明的名字照报既有 `VB3001`，且**第二个孔**报在 `(8,7)`（子扫描走 `[begin,end)` 窗口 + `getLocation` 播种 ⇒ 行列天然落在原文件，不需要事后平移 AST）；⑤ 新诊断三条全 ASCII 文案（D12）：`VB1007` 未闭合串 / `VB1008` 孔未闭合 / `VB1009` 空孔。读数与两条推翻计划书的订正见 **D68**（V1）与 **D69**（V2）。 | `4c16f50`+`f8ec76e`+`c2f7317` | 门 #82（8 job 全绿：Build + smoke + syntax + bas#1 + bas#2 + vbp + compile + asm；CI 分类逐条读数取不到 —— PAT 无 `actions:read` ⇒ job 日志端点 403，门结论以 run 级为准）。逐字节护栏 `.build\b20_emitc_guard.py`：BASE = 合并后、反引号之前的 `pre_b20_C3.exe` ⇒ **23 件存量工程 `--emit-c` 全同**；反向断言 **4/4**（V1 三件 + `test_interp.bas` 在 BASE 必失败、NEW 必成功）= 护栏能红的证明。用例：`test_rawstr`(+x86，18 条读数)、`rawstr_var`×4（GBK / UTF-8 BOM × CRLF / LF 同一内容四份源同读数）、`rawstr_proj`(+x86)、`test_interp`(+x86，21 条读数)、`rawstr_neg`×4 + `interp_neg`×4、`rs_emitc_shape` / `ri_emitc_shape`。本地读数：`-Category syntax` 两轮 128/0、129/0（第二轮含新增的 `ri_emitc_shape`）；`rawstr_proj` x64+x86 各 6 条 needle 全绿；四份变体各 4 条 needle 全绿。手册 `String 数据类型` 页加两节并标明非 VB6 原生（广告==应答）。**两条自己的坑**：(a) 四份变体第一次登记把 `"..." + $v + "..."` 直接写在参数位置 ⇒ PowerShell 把 `+` 当独立实参 ⇒ 四条 `FAIL (compile)`（本地 bas 跑抓到，门之前就修好并单独验过路径形状）；(b) 本机 `core.autocrlf=true` 把 `*.bas` 统统按 LF 存进索引 ⇒ 行尾/编码参与断言的夹具必须 `-text`，否则四份在 CI 上塌成两份、用例静默失效（见 `f8ec76e`，与仓库里 `.frx` 那段注释同源的道理）。 |
| B21 | 存量弱点 | **待拍板 5 收口**：布尔在「值→文本」与「装箱进 Variant」两条路上的读数合一（B20 撞见、A/B 确认与本批无关） | ☑ **B21 已出（代码 = `2ddcc8b`，门 = Actions run #83）** = 根因**是两条不是一条**，只修一条另一半仍在（实测：`CStr` 修好后 `VarType(b)` 依旧是 2）：① **可见性** —— `As Boolean` 与 `As Integer` 在 C 层同为 `int16_t`，cgen 的类型登记表按 **C 类型串**分派，于是 `inferExprType` 永远看不见布尔，`CStr(b)` / `b & ""` / `String` 形参收布尔实参 / 裸值的 `Debug.Print` 全落到整数分支（`-1`、`0`）；照 Fix 117c(Single)/Fix 175(Date) 的**并登记**口径加 `knownBoolVars_`，五个登记点（`Dim`/局部 `Const`/模块级变量/形参/函数返回值）各补一处、消费点一律先判布尔。② **装箱** —— `_Generic vb6_VariantFromValue` 按 C 类型选构造器，`int16_t` 命中 `short:` → `VT_I2`、布尔字面量的裸 `(-1)` 是 `int` → `VT_I4`，于是 `VarType(b)=2`、`TypeName(True)="Long"`、`Format` 走数字分支（`Format` 的布尔分支原本还**刻意**写着 `vb6_VariantInt((int16_t)x)`）；装箱点收进一个 `boxToVariant`，只有推断为布尔才换成 `vb6_VariantBool`、其余**原样**回退。RTL 侧本来就把 `VT_BOOL` 格式化得对（`vb6rtl_conv.c` 的 `True`/`False` 与 `TypeName="Boolean"`、`vb6rtl_format_extract.inc` 的 `Format`），所以这一半的活全在 cgen，不在运行时 —— 一开始按"RTL 不认布尔"去猜会走错方向。 | `2ddcc8b` | 门 #83（head `2ddcc8b`，**8/8 job 全绿**，逐 job 读数见 GATE_BASELINE）。逐字节护栏 `.build\b21_emitc_guard.py`：BASE = `pre_b21_C3.exe`（**HEAD 干净构建树**，先把 W1 改动导成 patch 再反向 `checkout --` 构建，构完 `git apply` 回来 ⇒ BASE 里不含本批任何一半）vs NEW = `C3.exe`，**30 件存量工程 `--emit-c` 只 4 行不同、且全在 `test_types.bas` 那几句 `Debug.Print` 布尔**（`-1` → `True`，正是要修的）；分类判据 = 新增侧必须带 `vb6_VariantBool(`/`vb6_CStrBool(`、两侧剥掉包装名与强制转换后逐字符相同。**本批最值钱的一条**：助手最初写成"实参已是 VARIANT 就直传"，看着更干净，护栏当场 RED —— 它把 VbQRCodegen 的 `vb6_VarType(vb6_VariantFromValue(VB6_SA_AT(...)))` 那层**恒等包装**一起删了（语义等价但存量可观察产物变了）⇒ 口径改成"非布尔原样回退"（读数见 D70-2）。判据用例 `tests\test_bool_display.bas` 32 条（B1-B8 转字符串四条路、B9-B14 类型标记、B15-B20 落进 Variant 的那一半、**B21-B27 反向护栏**钉 `Integer`/`Long`/`Byte` 的装箱读数、B28-B29 判定语义、B30 `${b}` 插值、B31-B32 裸值 `Debug.Print`）+ 登记 `test_bool_display[_x86]`。负控 = 同一份用例喂 BASE 二进制 ⇒ B1-B20 与 B30-B32 一起翻红（`B31-raw-1`、`B32-raw0`）。本地读数：x64 与 x86 各 32/32 通过（含走真 harness 的两条 shard）、`-Category syntax` 129/0、`test_types.bas` 现打 `Boolean=True`/`FalseVal=False`/`NotTrue=False`。**没顺带修的**：`Print #` / `Write #` 的实参不分类型一律 `vb6_Str((int32_t)x)` —— 同一批量出，已记待拍板 7。 |

> 批次可按实施中发现的耦合度合并/拆分，但**阶段范围不得越界**；每次运行只推进能各自独立过门的批。

## 设计记录（P0 / B00 产出，2026-09-23）

> 所有代码位置均为本次核实的现状坐标；实施时若已漂移以实际为准。

### D1 语法宿主（tB "twin 模块" 的本项目映射）

- 本项目**无 twin 模块**：模块种类只由扩展名在 `src/driver/driver_frontend.cpp:150-164` 决定，`.cls|.ctl|.pag → isClassModule`、`.frm → isFormModule`，`Module`（`src/ast/detail/ast_decl.hpp:275-306`）上没有 ModuleKind 枚举。
- **决策**：`Interface`/`CoClass` 作为**模块级声明块**，允许出现在 `.bas` 与 `.cls` 中（现有 `Enum`/`Type`/`Declare`/`Attribute` 在所有模块种类都无门控，故新增块不需门控改造）；名字**工程级唯一**，等价于 tB 的"工程级可见"。`.frm`/`.ctl` 不接受（头部分支有专用扫描逻辑）。
- 两种宿主形式：① **块形式** `Interface IShape … End Interface`，可与其他声明共存于 `.bas`（一个文件多个接口）；② **头行形式**（B03）`.cls` 首行 `Interface IFoo` + 尾行 `End Interface`，完全镜像既有 `Class Name(Of T)` 头行识别（`src/parser/parser_module.cpp:89-110`）与 `End Class` 消费（:105-110），使"一文件一接口"的 VB6 习惯成立。头行形式需要 `Module` 上新增 `isInterfaceModule/isCoClassModule` 子标记（该信息今天连 `.ctl/.pag` 都在 driver_frontend 局部丢失，正是同一处接缝）。
- 插入接缝：模块级主循环 `parser_module.cpp:82-165`，在 `Implements`(:120) 与 `Attribute`(:137) 分支之间加 `Interface`/`CoClass` 分支；块体解析模板取 `Enum…End Enum`（`parser_decl.cpp:260-290`，纯签名无体），成员含过程声明时取 `Property…End Property` 的 `parseBlockUntil` 式。
- **关键字策略（重要，零误伤）**：`Interface/CoClass/Inherits/Extends/Overridable/NotOverridable/Overrides/Protected/MyBase/Via` 全部按 Delegate(a15c40b) 的 4 处登记：`src/lexer/token.hpp` 枚举 + `src/lexer/token.cpp`（isKeyword/isStatementStart/kindToString）+ `src/lexer/lexer_keywords.cpp` 表 —— **并且同时登记进软关键字表 `src/parser/parser_helpers.cpp:14-63`**（先例：`Access`/`Default`/`Name` 等）。软 = 仍可在 `canBeName()` 位置当标识符用，块头部分支按 `kind` 精确匹配，两者不冲突。语料核查：`Interface` 一词在 tests/archive/publish 的 .bas/.cls/.frm 中仅出现于字符串与注释（6 文件，含 VBFlexGrid 的 `Attribute *.VB_Description`），其余 9 个词零命中；软登记是二道保险。
- **属性行 `[InterfaceId("{…}")]`**：lexer 现状是 `[` 直接扫到配对 `]` 产出**一个 Identifier token（文本含方括号）**（`src/lexer/lexer.cpp:231-244`），因为 `[My Type]` 名称引用语法依赖它。故**不改 lexer**：在模块级/声明前位置，若 Identifier token 文本以 `[` 开头且以 `]` 结尾，则用一个小解析器拆成 `属性名(+可选字符串实参)` 并挂到"下一条声明"上。今天这类行必然落入 `parser_module.cpp:160-164` 的 "unexpected token at module level" 错误分支，**错误→可解析，零回归风险**。
  - 现状缺口：`Attribute` 从不附着到声明节点（`Module::attributes` 是扁平列表，仅 `VB_Name`/`MultiUse` 被消费），也不存在 `Decl::attributes`。**决策**：给 `InterfaceDecl/CoClassDecl` 各自持有 `attributes` 字段（块级），成员级属性存进成员节点；不引入通用 `Decl::attributes`，避免全 AST 波及。
  - 支持集合（对齐 tB）：接口级 `[InterfaceId]`/`[Description]`/`[Hidden]`/`[Restricted]`/`[OleAutomation]`/`[ComImport]`/`[ComExtensible]`（后四类 P1 仅**接受+存档**，语义在 P6 才消费）；成员级 `[DispId]`/`[PreserveSig]`/`[Description]`；CoClass 级 `[CoClassId]`/`[ComCreatable]`/`[CoClassCustomConstructor]`。**不做** dispinterface 定义、`[Default, Source]` 事件源连接点（红线；仅语法与元数据接受）。

### D2 符号表与 mangle 键（沿用 $ov$ / _G_ 先例）

- 新增 `SymbolKind::Interface`、`SymbolKind::CoClass`（**追加**到 `src/semantics/symbol_table.hpp:23` 枚举末尾，不改既有值序；`Class/ComClass/ComInterface/Delegate` 已在此）。不复用 legacy 的 `Class + isInterface` 标记（`symbol_table.hpp:193-197`，`isInterface` 由 `semantic_analyzer.cpp:246` 与 `driver_compile.cpp:365-384` 的 3.6 阶段设置）——新式接口有显式符号，判定不再靠"被 Implements 才算接口"。
- 槽位（slot）身份键：`lower(<Iface>) + "." + lower(<slotName>)`，与 Overload 的 `name + "$ov$" + fp`（`symbol_table.hpp:329-332`）同族。存储键前缀先例 `$pg/$pl/$ps/$ev/$ty` → 新增 **`$itf$<ifaceLower>.<memberLower>`** 作为"实现映射"键（记录 *类成员 → 接口槽* 的对应，落在 `Symbol` 上的新 `implementsMap`，替代现在"什么都没有、靠字符串前缀猜"的局面：`interfaceMethodParams` 字段至今从未被写入，是死字段）。
- 槽名规范（属性拆三槽，COM 惯例）：`Property Get X` → `get_X`；`Property Let X` → `put_X`；`Property Set X` → `putref_X`；Sub/Function → 原名。C 标识符统一过 `CCodeGen::cIdent`（`src/backend/cgen_base_naming.cpp:49`）。
- 生成的 C 名字（**新前缀，与 legacy 完全隔离**）：
  | 产物 | 名字 | legacy 对照 |
  |---|---|---|
  | 接口槽表类型 | `vb6_ivtbl_<I>` | legacy `vb6_vtbl_<I>` |
  | 类侧槽表实例 | `vb6_ivtbl_<I>_for_<C>` | legacy `vb6_vtbl_<I>_for_<C>` |
  | 每 (类,接口,槽) 适配器 | `vb6_iimpl_<C>_<I>_<slot>` | 无（legacy 直连 `vb6_<C>_<I>_<M>`） |
  | 类级虚表（Inherits） | `vb6_cvtbl_<C>` / `vb6_cvtbl_<C>_impl` | 无 |
  | 对象内接口槽字段 | `__ivtbl[k]`（见 D3） | legacy `__comObj` 保留 |
- 接口内**禁止同名重载**（COM vtable 无重载概念）→ 编译期错误诊断，直接规避与 `$ov$` 体系的交叉复杂度。

### D3 内存布局与派发（P2 起即 COM 形态，避免二次改造）

- 现状：类实例 = `vb6_cls_<Name>`，**首字段固定 `void* __comObj`**（`src/backend/detail/base/cgen_base_generate_c_open.inc:73-112`），成员调用**全部去虚化直调** `vb6_<Class>_<Member>(me, …)`（`src/backend/cgen_util_classcall.cpp:17-228`）；唯一的间接派发是 legacy P6.4 的**胖对** `vb6_iface_<I>{vtbl,obj}` + `…_wrap()`（`src/backend/module/cgen_com.cpp:313-428`，派发在 `src/backend/detail/expr/cgen_expr_call_com_bind.inc:67-97`，赋值改写 `src/backend/detail/stmt/cgen_setlet_set_prop.inc:387-419`），且该 vtable **按成员名索引、无 IUnknown 前缀 → 不是 COM 表**。原生类实例今天**不计数**（裸指针，`_Destroy` 手工释放）。
- **决策 A：新式接口值 = 薄指针（单字 `void*`）**，指向对象内某个"vtable 指针字段"；这与 COM 完全一致，也是 `Set x = y` 语义与身份规则（同对象不同接口 → 不同地址、QI 往返地址一致）唯一自洽的表示。胖对 legacy 路径**原样保留不动**。
- **决策 B：槽表从第一天就是 COM 形态**：`vb6_ivtbl_<I>` = `[QueryInterface, AddRef, Release]` + （可调度时）`[GetTypeInfoCount, GetTypeInfo, GetIDsOfNames, Invoke]` + 展开后的自有槽。理由：内部调用与外部 COM 客户端共用同一张表，P6 不再改布局、不做双表；代价是 P2 就要生成本项目已有先例的适配桩（`vb6_disp_<cls>_<m>_invoke` 一族，见 `src/backend/detail/util/cgen_util_dllentry_collect.inc:245-344`）。`[PreserveSig]` 成员保留原生返回签名、不包 HRESULT。
- 槽序规则：`Extends` 链**深度优先、父先己后**，同层按声明序；IUnknown/IDispatch 前缀之上再排自有槽（即自有槽起始下标 3 或 7）。链上槽名冲突 → 编译错误（无影子槽）。
- 对象布局（仅对"实现≥1 个新式接口"的类改变，其余类逐字节不变 —— 这是护栏）：
  ```
  vb6_cls_Dog { const vb6_ivtbl_IUnknown_*  __ivtbl[k];  // 每接口一槽，声明序
                void* __comObj;  long __refcount;  <原字段…>;  struct vb6_events_X* events; }
  ```
  接口指针 = `&obj->__ivtbl[i]`；适配器拿到 `this` = 接口指针，**编译期已知偏移**做 container_of 减法得到 `vb6_cls_Dog*`，再直调 `vb6_Dog_<M>`。C 代码全按字段名访问（`me->__comObj`，见 `cgen_com_events.cpp:147` RaiseEvent），无裸偏移依赖，故前置新字段安全。
- 阶段接线：P2 先把 QI/IDispatch 前缀槽填**占位实现**（`E_NOTIMPL` / 简易计数），P6 换真实实现 —— 槽号因此从 P2 起就永久正确。

### D4 管线插入点

`src/driver/driver_compile.cpp` 现状阶段序：0 VBP → 1 lex → 1.5 pp → 2 parse → 2.5 typelib 导入 → 2.6 `runGenericsPrepass`(`driver_generics.cpp:91`) → 3 语义 → 3.5 `runCrossModuleResolution`(`driver_crossmod.cpp`) → 3.5b `runGenericsFixpoint`(`driver_generics.cpp:410`) → 3.6 接口标记 → 4 codegen → 5 link。
- **新增 stage 2.7 `runInterfacePrepass()`**（新文件 `src/driver/driver_interface.cpp`）：收集全部 `InterfaceDecl/CoClassDecl` → 建符号 → 解 `Extends` 链 → 产出**只读槽位表**（含序号、键、C 名、GUID）。必须在 stage 3 之前：跨模块引用接口名是常态。
- 跨模块延后（Overload 3.5 的先例）：*实现映射*（类成员 ↔ 接口槽）需要别的模块的成员表，放在 3.5 之后新增 **3.5c `resolveInterfaceContracts()`**；链式特化式的迭代需求用 3.5b 的 fixpoint 模式。
- 新式接口/CoClass 符号必须穿过克隆/替换通路（泛型 `src/ast/ast_clone.cpp:707-720` 与 `driver_generics.cpp:95-146`）—— v1 边界：**泛型类不得实现新式接口**，prepass 直接诊断拒绝（泛型 v1 也已经拒绝 `Implements`）。

### D5 与 legacy VB6 `Implements` 的共存策略

- legacy 三件套一律不动：模块级 `Implements X`（`parser_module.cpp:215-232`，点号拼成扁平串）、`IFace_M` 命名约定警告式覆盖检查（`semantic_analyzer.cpp:233-273`）、前缀扫描 vtable 生成（`cgen_com.cpp:333-375`）、`tests/test_implements_qi.bas` 与 ActiveX DLL 回归继续作为该路径的守卫。
- **分叉点**：`Implements <Name>` 解析后查符号，`kind==Interface`（新式）→ 走新路径（严格签名比对 + 槽表 + 薄指针）；`kind==Class`（legacy .cls 当接口用）→ 走旧路径。同名冲突（既有 `.cls` 叫 IFoo 又有 `Interface IFoo`）→ 编译错误。
- 成员级 `Implements I.M[, I.N]` 尾子句今天**不被解析**（`parser/parser_decl.cpp` 内无 `Implements` 处理）→ 属于"错误→可解析"的安全新增；组合满足多接口是 tB 明确特性，需支持逗号列表。
- 新式路径下 `IFace_M` 命名约定**不再必需**（显式子句或同名隐式匹配即可），但隐式匹配仍按大小写不敏感同名（含属性三槽）。

### D6 继承（P3）

- 语法：`Class Derived [Inherits Base]`（`.cls` 头行形式扩展，位置就在 `parser_module.cpp:89-104` 现有分支）；`MyBase.M(…)`；`Protected`；`Overridable/Overrides/NotOverridable`。
- 可见性：`src/common/types.hpp:54-59` 的 `AccessLevel{Public,Private,Friend}` **新增 `Protected`**（`Friend` 早已"保留未用"）；今天访问性**只在跨模块过滤**处生效（`symbol_table.cpp:383 getPublicSymbols()`、Private→C `static`），**调用点零检查、也没有派生域概念**（`ScopeKind` 只有 Module/Procedure/Block）。P3 需引入"当前类 + 其基链"上下文（语义分析器成员变量即可，勿扩 ScopeKind），并在成员解析处加 Protected 判定。
- 成员合并：派生类符号在 3.5c 之前做基链成员**继承合并**（`memberNames/memberReturnTypes/memberParams/memberProcKinds` 等表按 `Symbol` 现有分表结构逐项继承），派生域遮蔽规则 = 同名即遮蔽，签名不符的 `Overrides` → 错误。
- 虚派发：仅当类链中出现 `Overridable/Overrides` 时才生成 `vb6_cvtbl_<Cls>`（基先己后展平），`resolveClassMemberCall`（`cgen_util_classcall.cpp:17-228`）在"链上无虚成员"时保持今天的直调结果**逐字节不变**；`MyBase.M` 永远直调基实现（去虚化）。
- 构造链：沿用 `Class_Initialize/_Terminat` 现状，派生类初始化先跑基类（显式基限定调用按 tB 要求，v1 先自动链 + `MyBase.Class_Initialize` 支持）。

### D7 CoClass 与激活（P5）

- `CoClass Name … End CoClass`：块内 `[Default] Interface <I>`、`[Default, Source] Interface <E>`（**后者只接受并存档，不实现连接点**）、`[ComCreatable(True|False)]`、`[CoClassCustomConstructor]`（ByRef 实例入参 + 返回 HRESULT 的工厂过程）。
- 底层实现 = tB 口径：CoClass 是一张**契约聚合表**（可创建类 + 接口集合 + 默认接口），背后绑定一个私有 `.cls` 实现类（`Implements` 全部列出接口）；组内用户只见 `New <CoClass>` / `Dim x As <CoClass>`（等价其默认接口）。
- 激活：工程内 `New <CoClass>` 在编译期直接解析到实现类（不经注册表）；`CreateObject("ProgID")` 今天已走真实注册表链路（`src/rtl/core/vb6com/vb6com.c:509-560` 的 `CLSIDFromProgID + CoCreateInstance`，另有 `ComLib=` 免注册旁路 `vb6com.h:15-25`）→ P5 只需把 CoClass→CLSID→默认接口指针接进这条现成链路 + 编译期 ProgID 表。
- 工程侧既有元数据可复用：vbp 三段式 `Class=Name; x.cls; {CLSID}`（`src/project/vbp_parser.cpp:74-105`，消费于 `driver_compile.cpp:76-79` 的 `classClsidMap_`）、`Type=DLL|OleDll`→`ActiveXDLL`（`vbp_parser.cpp:284-286`）。

### D8 COM 兼容（P6）

- 已存在、优先复用不新造：运行时 `src/rtl/core/vb6comserver/`（`vb6comserver_obj.c` = 完整 IDispatch+QI+原子计数；`_factory.c` = IClassFactory；`_cp.c` = IConnectionPoint；`_pci.c` = IProvideClassInfo2；`vb6comserver.c:62-171` = 注册表写入），导出层 `src/backend/detail/util/cgen_util_dllentry_exports.inc:18-87`（DllGetClassObject/DllCanUnloadNow/DllRegisterServer/DllUnregisterServer/DllMain）、`activex_dll.def` 由 `driver_link.cpp:207-225` 生成、MSVC 链接常量已含 `ole32/oleaut32/uuid/advapi32`（`msvc_driver.cpp:229+`）。
- **真缺口**：`src/typelib/typelib_builder.cpp` 只会 `addDispInterface`（TKIND_DISPATCH）与 `addCoClass`，**没有 HK/双接口（TKIND_INTERFACE + 虚表）**，也不写枚举/UDT；`CreateTypeLib2(SYS_WIN64)` 写死（`typelib_builder.cpp:117`）→ x86 客户端读 64 位指针宽度语义，B15 要按 `-Arch` 传 `SYS_WIN32/SYS_WIN64` 并验证。
- GUID 来源：显式 `[InterfaceId]/[CoClassId]` > vbp 三段式 CLSID > **确定性 FNV-1a 生成**（现状 `cgen_util_dllentry_prelude.inc:39-70`，种子 `"iface:<progId>.<Iface>"`）。新式扩展到 `"itf:<Proj>.<Iface>"` / `"coc:<Proj>.<CoCls>"`。禁止随机：类型库与 .tlb 资源每次构建必须可复现。
- ABI 约束（先例已踩过的坑，写进验收）：x86 stdcall vs cdecl 与 `_`+`@n` 修饰；x64 统一调用约定但 `VARIANT.lVal` 截断指针的历史问题（Fix 093，`cgen_util_dllentry_collect.inc:304-343`）与结构体返回 ABI（Fix 184）。新式槽函数**一律 `STDMETHODCALLTYPE`**，参数按 COM 签名（[out,retval] / ByRef 指针 / propputref）。
- 验收（B17）：`tests/test_activex_dll/` 的"编译 DLL→注册→外部 CreateObject 晚绑定断言"模式（含 `tests/test_p613_typelib.bas`）复制一份新式接口版；再加一个原生 C++/PowerShell `CoCreateInstance` 早绑定 QI 冒烟。

### D9 测试与护栏（每批必做）

- 注册位置：`tests/run_tests.ps1` 内的 `Add-BasTest` / `Test-Vbp` / `Test-GuiVbp` 硬编码登记（无清单文件）。断言 = 内联子串数组匹配 stdout（`.bas` 用例）或 `XXX-N:OK` 标记（项目用例），**无期望文件、无 .c 快照**。
- 计划新增用例：`tests/test_interface.bas`、`tests/test_iface_chain.bas`（Extends 槽序/契约报错）、`tests/itf_xmod/`（跨模块接口）、`tests/test_inherit.bas`、`tests/cls_inherit/`、`tests/test_iface_via.bas`、`tests/test_coclass.bas`、`tests/ax_coclass/`（DLL + 外部激活）。诊断类批（B01/B02）用 `-Category syntax|compile`，断言错误文本。
- 门：`scripts/build.bat`（或 `scripts/dev.ps1`）+ `tests/run_tests.ps1 -Category all`，`Results: PASS=… FAIL=0 SKIP=… TOTAL=…` 原文记入 GATE_BASELINE；已知环境性 SKIP 只有 `test_vbman`（`-RequiresCom VBMANLIB.cVBMAN`）。
- **零新语法逐字节护栏**：沿用先例仪式（`.build\C3.exe <src> --emit-c` 与批次前输出 diff；Overload/Generics 分别以 5+3 与 16 文件验证），B04/B05/B08 三批改结构体/派发路径，必须各跑一次；注意 `git show HEAD:<file>` 是 LF，比较前先还 CRLF（`todo/ferock.md:94` 的教训）。
- 编码红线（CONTRIBUTING §）：`.bas/.cls/.frm/.vbp` = **GBK + CRLF**，`.md` = UTF-8，`.ps1/.bat` = GBK。新增测试用例文件必须以 GBK 写。

### D10 风险登记

| # | 风险 | 缓解 |
|---|---|---|
| R1 | 改 `vb6_cls_*` 结构（前置 vtbl/refcount）波及其它类布局 | 仅"实现新式接口的类"加字段；无新语法工程 emit-c 逐字节护栏 |
| R2 | refcount 与现状"裸指针手工 `_Destroy`"混用 → 双释放/泄漏 | 只对被接口接管的对象 Release；`Set x = Nothing` 与 `_Destroy` 的归属在 B05 明确并加测试 |
| R3 | P2 一次做 COM 形态槽表工作量大 | 槽表 COM 化 + 占位 IUnknown，把真实 COM 语义推到 P6，不返工即可 |
| R4 | 新关键字与存量标识符冲突 | 软关键字双登记；语料已核查零真实冲突 |
| R5 | typelib 只支持 dispinterface，dual 接口导出可能触及未知 API | B15 前先用最小 .tlb 手工验证 `ICreateTypeInfo::Layout` 路径；失败则回退为"仅 IDispatch 晚绑定可用"，并在总表记边界 |
| R6 | 共享工作树 3 写者 | 每批开始 `git status` + mtime 判定；他人半成品致构建失败即停并记录 |

### D11 v1 边界（明确不做）

dispinterface 定义；`[Default, Source]` 连接点实现；泛型类实现新式接口；接口内重载；接口成员含 ParamArray/UDT 参数（先警告，P6 视需要升级为错误）；事件在接口上的声明。

### D12 B01 实施中产生的设计校正（2026-09-23，写代码后回灌）

- **属性容器落点**：接口成员属性用 `InterfaceMember{ std::vector<InterfaceAttr> attributes; DeclPtr decl; }`（`src/ast/detail/ast_decl.hpp`），确认**不**引入通用 `Decl::attributes`；块级属性存 `InterfaceDecl::attributes`。
- **诊断文案必须用英文**（对 D1 的硬修正）：`tests/run_tests.ps1` 是无 BOM 的 GBK 脚本，Windows PowerShell 按 ANSI 码页读 `.ps1` → 脚本里的中文断言字面量必然乱码并静默匹配失败。解析器既有的 `expect` 类消息本来就是英文，故新增 4 条接口诊断统一英文 + 新 ID `ParseInvalidInterfaceMember=2011` / `ParseUnknownAttribute=2012`（`src/common/diagnostics.hpp`）。后续批次的契约诊断（B02）沿用此口径：**面向测试断言的错误文本 = ASCII**。
- **泛型拒绝前移到语法层**：`Interface X(Of T)` 与接口成员 `(Of T)` 在 parse 阶段即报 2011（原 D11 把它放在"prepass 直接诊断拒绝"），省掉一条跨阶段不变式。
- **软关键字边界**：`Interface`/`Extends` 进关键字表 + `isSoftKeyword`，但**不进** `Token::isStatementStart` —— 这样 `Dim Interface As Long` / `Interface = 7` / `Debug.Print ... + Extends` 全部照旧，接口块只在**模块级声明位**与**接口成员位**被识别（语句位永不识别）。`tests/test_interface.bas` 里有这段回归保险。
- **死循环守卫**：接口块成员循环遇到 `End` 且其后是 EOF 时必须 break，否则 `advance()` 在 EOF 不前进 → 编译器挂死（已加，B01 期间发现）。
- **属性行的接受位置**：只接受"模块级属性行 + Interface 声明"与"接口块成员位属性行"两处；其他位置（例如属性行后跟 `Sub`）报 `Attribute line must precede an Interface declaration` 并跳行 —— 这类行过去必然 VB2002，仍是"错误→可解析"。
- **B01 正例刻意不使用接口类型**：`Dim x As IShape` / `Implements IShape` 在 B01 还没有接口符号（`Module::interfaces` 不进符号表），用了就编译失败；契约与派生用例从 B02/B04 起补进同一文件。

### D13 与并发写者共处的工作节律（本轮实测）

- `scripts\build.bat` 从 agent 的 cmd 调用不可用（GBK 正文 + LF 行尾 → 错码刷屏后 exit 1，根本没跑 cmake）。构建一律 `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\dev.ps1 -SkipTest`。
- 锁窗口探测：`(Get-Process C3,cl,link,ninja | Measure-Object).Count`。**他人的一轮 `-Category run` 实测跑了 30 分钟以上**，其间 `.build\C3.exe` 被持有（重建必 LNK1168）且 `output/` 被反复改写。故：锁忙时只做编辑/文档/用例编写，不构建、不测量。
- 锁忙时的单文件预检（不产出目标文件、不碰 `.build`）：`cl /nologo /TP /Zs /EHsc /std:c++17 /utf-8 /Isrc <abs>`（包在 `call vcvarsall x64` 里）。B01 靠它提前抓到一个多余右括号。
- **基线纪律**：一次只跑一个回归，日志路径按批次唯一。本轮我自己先用 `Start-Process` 脱管跑了一个，又用后台任务跑了一个，两者写同一个 `.build\gate_base.log`，我从中读到的 `Results:` 行不可归因 → 记为无效；正确做法是跑前 `md5sum .build\C3.exe`、跑后再核一次，两次不同即作废重跑。
- 脱管进程不会被工具的 stop 回收（工具只杀自己的 wrapper），会留下孤儿继续锁 exe；识别办法是 `Get-CimInstance Win32_Process` 看命令行，只对**自己启动的** PID 做 `taskkill /T /F`。

### D14 B02 开工地图（2026-09-23 B01 收尾轮产出，行号为本轮实测）

- 符号登记：`SymbolKind`（`src/semantics/symbol_table.hpp:23-44`）现末尾值组是 ComGlobalNs 等，**在其后追加** `Interface`（再留 `CoClass` 给 P5），不动既有值序；`Symbol` 的新式字段直接加在 legacy 字段区 `isInterface/implementsNames/interfaceMethodParams/interfaceMethodNames`（:193-197，其中 `interfaceMethodParams` 仍是死字段，勿与其同名混用），建议 `slotTable`（vector：槽名/C 名/键/签名）+ `extendsChain` + `ifaceAttrs`。
- stage 2.7 接线：`src/driver/driver_compile.cpp` 阶段序现状——2.6 `runGenericsPrepass()`(:329) → 3 语义 → 3.5 `runCrossModuleResolution()`(:347) → 3.5b `runGenericsFixpoint()`(:358) → 3.6 接口标记(:365)。新 `runInterfacePrepass()` 插在 :329 与 :347 之间语义段之前；数据源=各 `Module::interfaces`（B01 已填）；泛型类宿主拒绝、`.cls` 同名冲突（既有 Class 叫 IFoo + Interface IFoo）在此报。
- Implements 分叉：`src/semantics/semantic_analyzer.cpp:232-273` 是 legacy 覆盖检查（warn 级、`IFace_M` 命名约定）；B02 在其**前面**按 `lookupModule(ifaceName)->kind` 分叉：`==Interface` → 新路径（error 级：缺槽/多槽按 D2 键、签名逐参比对、属性三槽 `get_/put_/putref_`），`==Class` → 原逻辑一行不动。legacy `tests/test_implements.vbp` 回归继续当旧路径守卫。
- 成员级 `Implements I.M` 尾子句：模块级 `parseImplements`（`parser_module.cpp:242`，声明头收集）已核实**不解析过程尾部**；B02 在 `parseSubOrFunction/Property` 尾部（`parser_decl.cpp` 的签名解析完成后、`expectEndOfStatement` 前）加可选 `Implements` 逗号列表，存进过程 Decl 新字段（如 `implementsClauses`）；`Implements` 本是硬关键字（TokenKind::Implements），无软登记问题，但需确认语句位恢复路径：`isStatementStart` 现状含 Implements，故过程尾扫到会立即在 `expectEndOfStatement` 报错——即"错误→可解析"安全新增（同 B01 属性行手法）。
- 诊断文案 = ASCII（D12 硬约束）；测试用例文件 = GBK+CRLF（D9 红线）；新用例登记在 `tests/run_tests.ps1` 的 itf_neg 数组（:864 起）与 bas 队列（:725 附近）。
- B02 门新增断言：契约报错类用 `Test-SyntaxFail`（编译期诊断）；正例契约满足类目前只能到 `--compile-only`/syntax 级（emit 要到 B04），`test_interface.bas` 里以"Implements IShape + 完整同名成员 → 无诊断"形式做 guard（现有 Add-BasTest 是 compile+run，接口类型变量还不可用，B02 正例改放 `.bas` 但**仅经 Test-Syntax 类通路**——注意 run 阶段若 `Set x = New Class` 未涉及接口类型则不受影响）。
- 逐字节护栏复用本轮做法：临时 `git worktree add D:\c3.<tag> --detach HEAD` + 同 cmake(Ninja/Debug 用主 .build 同款 cache 参数) 构建基线 exe，8 文件清单见本轮日志行，`--emit-c` 双路 `cmp`；用后 `git worktree remove --force + prune`。
- **GitHub Actions 全量 = 里程碑级，不是每批**（用户 2026-09-23 09:55 定调："这个测试不用每次都做，
  按之前的做就行，等大改全部实现完成之后再跑 github actions 这样的全量测试"）。所以：
  **每批仍走本机 `tests/run_tests.ps1 -Category all` 门 + 逐字节护栏 + 本地 commit**（本文件既有纪律，
  一字不改）；Actions 那一级留到**整条线（B04–B18 / STATUS=ALL_DONE）做完**跑一次，届时：
  仓库 = `https://github.com/fxl447098457/c3test`，主路径 `git push github HEAD:dev` → `ci.yml`
  （`regression` 矩阵 = run_tests.ps1 五段 smoke/bas×2/vbp/compile/syntax，`smoke` 本轮刚补进矩阵；
  CI 是 Release 构建，与本机 Debug 两个口径），体系级 T0/T1/T2 走 tag `github-test-NNN` 或
  workflow_dispatch → `ci_t0.yml`；监控用 `pwsh -File scripts/watch-gh-actions.ps1`
  （自取 remote 内嵌 PAT、输出打码，别手动 echo 那个 URL）。
  另两处事实供那一次参考：`tests_github/` 的用例是 2026-09-20 从 `tests/` **复制**的，本线新增的
  `itf_*` 在 run_t1/run_t2 里 0 命中（**他人正在修清单同步，本任务不动**）；`run_t2.ps1` 也不在
  `run_tests.ps1` 里。gitcode 那条 `origin` 与 Actions 无关。
  **分支名已定：远端 c3test 保持 `dev`，不改名为 `fan/dev`**（用户 2026-09-23 09:58 决定：
  "既然不改也不影响的话，那我就不改了"）。本地 `fan/dev` → 远端 `dev` 用 refspec 映射
  （`git push github HEAD:dev`）即可，全仓唯一引用分支名的 `.github/workflows/ci.yml:5`
  保持 `branches: [main, dev]` —— **不要再往里加 `fan/dev`**（本轮已有一次加了又撤）。
  **红线**：本机门不变 → 平时依旧"永不 push"；只有上面那一次里程碑全量允许推 `github` 远端
  （`dev` 分支或 `github-test-*` tag），不建 MR，且届时先跟用户确认一句。
- **触及类结构体布局 / Set / 派发 / COM 打包的批次，除全量门外必须另跑 T2**（B04–B06a 这三批此前漏了）。


### D15 B02 实施中产生的设计校正（2026-09-23，写代码后回灌）

1. **契约比对落在 stage 3（语义层）而不是 D4/D14 计划的 3.5c**。理由（实测后确认）：
   实现映射只需要"实现类自己模块的成员表"，而那份表在 stage 3 的模块 AST 里已完整；
   跨模块需要解决的其实是**接口名→契约**，这由 Driver 级登记表 `ifaces_` 一次建好即可，
   不必往每张模块符号表里注入 Interface 符号（那会牵动 `getPublicSymbols()`/3.5 跨模块注入，
   风险面大得多）。分叉点仍在 `semantic_analyzer.cpp` legacy Implements 循环的**开头**：
   命中登记表 → `checkNewStyleInterface()` + `continue`；未命中 → legacy 代码一行不执行改动。
2. **`SymbolKind::Interface` 本批不加**（与 D2/B02 批次描述偏离）：加了没有任何消费者
   （符号表不参与本批判定），等 B04 发码期与消费点同批落地。同理 `Symbol::slotTable`
   等 legacy 符号字段也不动——契约数据全在 `IfaceRegistry` 里。
3. **签名比对口径 = 源码签名**（`interface_sig.hpp`）：类型引用**原文小写** + 参数个数 +
   逐参 `ByVal/ByRef`/`Optional`/`ParamArray` + 返回类型原文。不用 `Vb6Type` 归一：
   Fix 047 同源问题（跨模块 Enum/UDT 在 stage 3 早期还没注入，归一后比会把正确实现判成不符）。
   `As String * N` 记作 `fixedstring`（宁可显式不匹配，也不伪装成 `String`）。
4. **新共享头 `src/semantics/interface_sig.hpp`（header-only）**：槽键规范
   （`get_/put_/putref_`，D2）与签名文本两侧只用这一份定义；`IfaceSlotView` 额外带
   `memberName`（声明原样名），否则诊断文本只能打印小写槽键。
5. **`--syntax-only` 实际会跑到 stage 3/3.6**（`driver_compile.cpp` 的 syntaxOnly 早退在 3.6 之后），
   所以 B02 的契约诊断**不需要 vbp 工程**就能测：单个 `.cls` 直接喂给 `C3.exe … --syntax-only`
   即可，`Test-SyntaxFail`（非空退出码 + ASCII 子串）完全够用。新增 6 条负例
   （n08 缺槽 / n09 签名不符 / n10 未知父 / n11 Extends 环 / n12 接口内同名 / n13 与模块重名）
   + 1 条正例 `tests/itf_pos/p01_contract_ok.cls`（`Test-Syntax` 通路，含 Extends 继承槽与
   属性三槽）→ 门总数 111→118。**这条对后续批次普遍适用**：凡"只报诊断、不发码"的批
   都用 Test-SyntaxFail/Test-Syntax，别急着造 vbp 工程。
6. **泛型模板双登记护栏（本轮补）**：模板 `.cls` 本体与其特化克隆都在 `modules_` 里，
   模板内的 `Interface` 块会被登记两次 → 莫名其妙的重名错。Pass A 直接拒绝
   （新 ID `SemInterfaceNotSupported=3018`）。该分支无用例（需要泛型工程），v1 边界足够。
7. **已知遗留（B02b 第①项）**：成员级 `Implements I.M[, I.N]` 尾子句未开工（需 `parser_decl.cpp`
   + `ProcDecl` 新字段 + 显式绑定优先于同名隐式匹配）。
8. **legacy 交互（P6 必修，现记档）**：新式接口名仍会进 `classSym->implementsNames`
   （收集发生在分叉之前），ActiveX DLL 的 `cgen_util_dllentry_tables.inc:88-130` 会按
   legacy 口径给它 mint 一个确定性 IID 并写进 `g_vb6_ifaceIids_*`。EXE 工程不发 dll_entry
   → 本批测试全绿不受影响；B13/B16 接线时必须在此处分叉（`ifaces_` 命中就走新式 IID）。
9. **构建/门实测节律**：本轮 03:07 开工 → 03:30 首门（25 分钟，PASS=118 FAIL=0 SKIP=1）
   → 复核源码时发现两处应修（见 6、及一处 move 后读键的诊断文本 nit）→ 03:59 重建 →
   04:02 复跑终门。**过门后若再改源码，必须重跑全量门**（否则提交的源码 ≠ 被测二进制）。

10. **重入保护的实测漏洞（本轮撞到）**：上一轮在 03:05 写 `STATUS=IDLE` 并提交 B01（3add1ce），
    **但它在 03:08:06 又提交了第二个 docs 提交 ea47b46**（回记哈希 + 产出 D14 地图），也就是说
    本轮 03:05–03:08 与其尾部重叠。本轮之所以没出事纯属顺序侥幸：它的提交只动总表，且我
    第一次 `git status`（03:05，看到 26 个脏文件）与第二次（03:07，树已 clean）之间正好夹着那次
    commit——中途 `git diff --stat` 返回空也一度让我以为总表在撒谎。**下一轮起加一条硬检查**：
    读总表后立刻 `git log -1 --format='%h %ct %s'`，若"最新提交时间 - LAST_RUN" 的绝对值 < 3 分钟，
    或工作树在一次 `git status` 里从脏变 clean，就 sleep 120s 再看一次，确认没有活的写者才动工具链。
    （`STATUS=BUSY 且 LAST_RUN<55min` 只挡住"提前收尾"的读法，挡不住"写完 IDLE 还在提交"。）

11. **总表的编辑会被对方的提交"顺手带走"**：ea47b46（03:08:06，上一轮的收尾 docs 提交）提交的是
    **工作树里当时已含本轮 BUSY 头**的那份文件 —— 即我 03:07:46 写的状态头被写进了*它的*提交。
    结果是自洽的（本轮末尾又在 beb75a7 里把该头改成 IDLE + 本轮记录），但要明白：**共享工作树里
    总表没有"我先写就是我的"这回事**。安全做法 = 本轮结束时把状态头**重写**一遍（本轮即如此），
    并且只用"批是否打勾 + CURRENT_BATCH 文本"判断进度，别用 LAST_COMMIT/STATUS 的字面值。

### D16 B02b（成员级 Implements 子句）实施记录与设计校正（2026-09-23）

1. **子句的语法位置**：`Sub` 在参数表之后、`Function/Property` 在 `As Type` 之后（tB 口径），
   三处共用 `Parser::parseTrailingImplementsClauses`（`parser_decl.cpp`）。接口块成员签名走
   `parseInterfaceMemberDecl`，**不调用**该函数 → 在 Interface 块里写子句仍然 VB2003，
   与"契约上不该有绑定"一致。畸形子句（只写 `Implements I`）在 parse 期报 2011 且**不入列表**，
   避免语义层再补一条级联诊断。
2. **显式绑定排斥隐式同名**（tB 口径，D5 的"显式优先"具体化）：写了任意子句的成员只按子句入座，
   不再参与同名隐式匹配（`checkNewStyleInterface` 里的 `explicitDecls` 集合）。否则一个成员会
   既按子句进 A 槽、又按同名被 B 接口认领，形成没人能预期的隐式双重认领。
3. **槽键解析要兼容两种写法 + 链上任一接口名**：`I.Name` 按实现成员的 Get/Let/Set 前缀推
   `get_/put_/putref_name`，同时也允许直接写槽名 `I.get_Name`；接口名可写 Extends 链上任意一层
   （`IfaceSlotView::ownerIface` 已带归属接口，父槽用父名或子名都能解）。属性 Set 槽 = `putref_`。
4. **"认领不到的子句"必须报错**，否则子句写成摆设还全绿：`checkNewStyleInterface` 每处理一个新式
   接口就把认领的子句记进 `boundClauses`（键 = (过程节点, 子句序号)），`analyze()` 末尾交给
   `checkMemberImplementsClauses` 兜底。四种成因分别给一句话（宿主非类模块 / 接口名不存在或是
   legacy 类 / 类未实现该接口 / 接口里没有这个成员）。新诊断 ID `SemInterfaceClauseUnbound=3019`。
   兜底同样落在 stage 3：子句解析只需工程级登记表 + 本模块成员表（D15-1 的结论对子句成立，
   不需要 3.5c）。
5. **单文件 `--syntax-only` 通路继续够用**（D15-5 第三次验证）：6 负 + 1 正全部 ASCII，无需 vbp。
   泛型模板内 Interface 的 3018 分支本轮补了用例 `itf_neg/n19_iface_in_generic.cls`（B02b-②）。
6. **`tests/run_tests.ps1` 是 UTF-8 无 BOM + CRLF**（4899 个非 ASCII 字节，`lf_only=0`），不是 D9
   写的 GBK。PowerShell 仍按 ANSI 码页读它 → 中文注释无害、中文字符串字面量必乱码（D12 的根因，
   本轮实测复现）。改它只能走二进制插入：写完核对 ①非 ASCII 字节数不变 ②`lf_only` 仍为 0
   ③`git diff --numstat` 只有预期的 +N/-1。本轮插入 11 行、n13 行加逗号。
7. **AST 新字段的克隆路径**：`SubDecl/FunctionDecl/PropertyDecl` 各加 `implementsClauses`，
   `ast_clone.cpp` 的 `cloneSubDecl/cloneFunctionDecl/clonePropertyDecl` 必须逐个补拷贝——
   泛型特化走的就是这三个函数，漏拷会让特化副本静默丢掉绑定（v1 已拒泛型实现接口，但克隆路径
   仍然要正确）。
8. **过门纪律的具体踩法**：本轮首门（05:01 起，exe 1f387137）跑到一半时又改了两处源码
   （去掉 `IfaceClauseRef` 的重复声明、畸形子句不入列表）→ 那次 125/0/1/126 只能当"存量项无回归"
   的参考，**不能作为提交依据**；必须重建 + 复跑全量门（D15-9 的再验证：门与二进制一一绑定）。
9. **已知不一致（留给 B03/B13，不在本批动）**：模块级 `Implements` 的分叉点用的是
   `ifaceReg_->find(Symbol::toLower(全名))`，即 `Implements Proj.IFoo` 这种带库/工程限定的写法
   **不会**命中新式路径（会掉进 legacy 分支）；成员级子句这边则接受限定名（末段回退，
   `ifaceLastSegment`）。二者对限定名的宽容度不对称。B03 起若要支持工程级接口引用，
   应把分叉点也改成同一套 `lookupWrittenIface`，并顺手补一条正例用例。

### D17 B03 开工地图（2026-09-23 B02b 收尾轮产出，行号为 dde7c32 上实测）

- **头行形式其实已经能解析**（探针实测，省掉一整块语法工作）：`.build/probe_itfhead.cls` =
  `VERSION/BEGIN…END/Attribute VB_Name = "IFoo"/Option Explicit` + `Interface IFoo … End Interface`，
  用 B02b 二进制跑 `--syntax-only` **只**报一条错误：
  `error VB3002: Interface name 'IFoo' collides with a module of the same name`
  （`src/driver/driver_interface.cpp:44-50`，moduleKeys 在 :26-27 预建）。也就是说 B01 的块解析器
  已经天然吃得下"一文件一接口"的头行写法，**B03 不是新语法批次，而是"宿主识别 + 重名放行"**。
- 三步落地：
  1. **宿主标记**：`Module` 加 `isInterfaceModule`（接缝就在 `src/ast/detail/ast_decl.hpp:332-333`
     的 `isClassModule`/`isFormModule` 旁，D1 早已点名这里丢子标记）。置位条件建议：
     `mod.isClassModule && mod.interfaces.size()==1 && ifaceLower(块名)==ifaceLower(mod.moduleName)`，
     在 `parseModuleBody` 的 Interface 块分支尾部扫一遍即可（`src/parser/parser_module.cpp:126-151`）。
     若要"整文件即该块"（块外不得再有声明），参照 Class 头行的 `classHeaderSeen_` 记法
     （`parser_module.cpp:89-104`）加一条收尾校验，报错文案保持 ASCII。
  2. **重名放行**：`runInterfacePrepass` Pass A 的 moduleKeys 集合把"宿主自己的模块名"剔除
     （只豁免它自己声明的那一个接口名；与其它模块/其它接口重名仍照旧报错）。
  3. **确认 legacy 侧不受污染**：头行宿主的成员是**签名节点**、不进 `mod.declarations`，所以它的
     Class 符号 `memberNames` 为空。别的模块写 `Implements IFoo` 时，分叉点
     （`semantic_analyzer.cpp:240-246`）会先命中登记表走新式路径——需要**实测**跨模块场景：
     `tests/itf_xmod/`（D9 已规划）建一个 vbp 工程 = 宿主 `IFoo.cls` + 实现类 `Foo.cls`，
     断言只有新式诊断、没有 legacy 的 `IFace_M` 命名约定警告。这一步也是 B04 发码的前置。
- **手册**：新增 `docs/vb6-manual/02-语句/Interface 语句.md`（模板 = 同目录 `Implements 语句.md`，
  UTF-8），并在 `docs/vb6-manual/README.md` 的语句清单第 157 行（`- [Implements 语句](…)`）之后插
  `- [Interface 语句](02-语句/Interface%20语句.md)`（注意 `%20` 转义；字母序 Implements < Interface < Input）。
- 门：基线 `125/0/1/126`。建议用例 = 1 正（宿主 .cls 单文件 `Test-Syntax`）+ 1 负（宿主名与另一
  接口冲突 → 仍 VB3002）+ 1 个 `itf_xmod` vbp 工程（`Test-Vbp`）。**逐字节护栏**：B03 不动发码，
  但改了 `Module` 结构与 prepass，按 D9 仍跑一次 `--emit-c` 对比（沿用 B01 的 8 文件清单）。

### D18 B03（`.cls` 头行宿主）实施记录与设计校正（2026-09-23）

1. **宿主识别必须放在 stage 2.7，不能放 parser**（D17 地图原建议"在 parseModuleBody 尾部扫一遍"
   实测不可行）：`Module::moduleName` 要到 `src/driver/driver_frontend.cpp:217-250` 才从
   `Attribute VB_Name` 定下来（parser 只见得到属性行），所以 `mod->isInterfaceModule` 由
   `runInterfacePrepass` Pass A 在泛型拒绝之后、登记接口名之前置位。识别条件
   = `isClassModule && interfaces.size()==1 && ifaceLower(块名)==ifaceLower(moduleName)`。
2. **VB3002 放行只豁免宿主自己的那一个块**：`hostOwnName = isInterfaceModule &&
   d.get()==interfaces.front().get()`。其它接口名与工程内模块名撞车仍照旧报错，
   所以 B02 的 `itf_n13_module_collision` 负例不受影响（本轮门内仍 PASS）。
3. **宿主文件的"1 文件 1 接口"约束**用 `SemInterfaceNotSupported`(3018) 报
   `Interface host module 'X' may contain only the Interface block`，且**只报第一条**
   （一条宿主违规背后往往是整批误用，级联文本没有信息量）。`options`/`attributes`
   不算声明——`.cls` 宿主必然带 `Option Explicit` 与 `Attribute VB_Name`。
4. **D17 的"legacy 污染待核点"结论 = 无污染，跨模块已端到端打通**：`tests/itf_xmod/`
   （`IWriter.cls` 头行宿主 + `CWriter.cls` 用 B02b 的成员级子句跨模块绑定三个槽 +
   `XMain.bas` 走具体类调用）编译成 exe 并跑出 `XMOD1:OK`/`XMOD2:OK`。机理：宿主的 Class
   符号仍在、`memberNames` 为空，但 `semantic_analyzer.cpp` 的 Implements 分叉先查登记表 →
   新式路径，legacy 的 `IFace_M` 命名约定一行未执行。**接口类型变量 `Dim s As IWriter` 本批
   刻意不做**（同名 Class 符号会先被类型解析吃掉）→ 归 B04。
5. **语法手册落点**：新页 `docs/vb6-manual/02-语句/Interface 语句.md` 首行用引用块标注
   "本项目扩展，非微软 VB6 原生语句"，与 MSDN 镜像正文区分；README 索引按字母序插在
   `Implements 语句` 与 `Input # 语句` 之间（链接用 `%20`，`#` 不转义）。
   该目录全部是 **CRLF**：Write 产出 LF 后要 `sed -i 's/\r*$/\r/'` 归一，且改 README 之后
   必须复查 `bare_lf==0`（本轮实测：README 524 CRLF / 0 bare-LF）。
6. **逐字节 emit-c 护栏本批未跑**（记录理由，免得下轮以为漏了）：只加了一个 bool 字段 +
   prepass 分支，未碰 `src/backend/**` 与发码路径；替代守卫是让宿主模块进真实工程
   编译+链接+运行（第 4 条）。按 D9 口径，B04/B05/B08 那三类改结构/派发的批次仍必须跑。
7. **用例通路新增第三条**：`Test-Vbp`（`run_tests.ps1` 的 all/run/vbp 块，M7Test 之后）。
   登记 = 插 3 处共 13 行、改 1 行（n19 补逗号），核对口径见 D16-6。

### D19 B04 开工地图（2026-09-23 B03 收尾轮产出；行号本轮实测，含对 D3 的硬修正）

- **⚠ 先改设计再做码：D3 的"前置新字段安全"结论是错的（本轮实测推翻）**。
  `src/rtl/core/vb6comserver/vb6comserver_obj.c:268-272`（`vb6_ComObject_Create` 回填反指针）与
  `:302-303`、`:318`（`vb6_ComObject_FromInstance` 复用/新建包装）都是 **裸偏移 0 读写**
  `void** ppComObj = (void**)instance;`，并且 `vb6comserver.h:126-128` 把它写成明文前置条件
  （"实例所在类的结构体首字段是 `__comObj`"）。而 `src/backend/detail/base/cgen_base_generate_c_open.inc:78`
  从 ExeComBridge 01 起**无条件**把 `void* __comObj` 发在 `vb6_cls_<Name>` 第一位。
  → **B04 落法**：`__comObj` 保持第 0 位不动，接口槽指针数组放在它**之后**
  （`__comObj; const void* __ivtbl[k]; <原字段…>`），container_of 减法按"字段名 + 实测偏移"算，
  不假设任何偏移；这样存量 COM 包装复用路径零改动。D3 里"接口指针 = `&obj->__ivtbl[i]`"的表述仍然成立。
- 结构体接缝现状（`cgen_base_generate_c_open.inc`）：`vb6_cls_<Name>` 开在 :73，`__comObj` :78，
  字段循环 :81-102，空结构补 `_placeholder` :104-106，`events` 指针 :108-111，收尾 :112。
  另外 `usedVb6IfaceTypes_`（`cgen_state.inc:262` 登记 → `cgen_base_generate_epilogue.inc:44-59` 出
  前置 typedef）已经在发 `vb6_vtbl_<I>` / `vb6_iface_<I>` 两个名字，**`vb6_ivtbl_` 前缀今天全仓零占用**，
  D2 的隔离命名可以放心用。
- **legacy 路径不发适配器**（与 D3 的预设有偏差）：`src/backend/module/cgen_com.cpp` 的
  `emitInterfaceVtable(Module&)` :313-428 只遍历 `module.implements` :316，成员靠**实现类自己声明的
  `IFoo_M` 前缀扫描**得到 :330-375，槽位**直接指向** `cProcName(IFoo_M)` 实现函数（:402-413），
  没有任何 thunk；`vb6_iface_<I>{vtbl,obj}` 胖对在 :394-398，`vb6_iface_<I>_wrap()` 在 :416-426。
  → B04 的 `vb6_iimpl_<C>_<I>_<slot>` 适配器是**新物种**，可抄的发码先例是 P6.5 事件包装
  （`cgen_base_generate_body_pass.inc:30-130`，`evtWrapperName`，签名拼装 :89-98）与
  P13.23 的 `emitComVtableSinks()`（`src/backend/module/cgen_com_events.cpp:288+`，声明 :256-270，
  钩子 `evt_decl.inc:76` / `evt_impl.inc:135`）。
- **两个必须新建的基础设施**（勘察实测，别当已存在）：
  1. `CCodeGen` **拿不到 `IfaceRegistry`**：登记表只在 `src/driver/driver.hpp:143 ifaces_`，
     消费者只有语义层（`semantic_analyzer.hpp:127`）。要按 legacy 的做法在模块循环里注入
     （先例：`ifaceImplementersMap` 产于 `src/driver/detail/driver_codegen_dup_module_vars.inc:10-25`，
     注入于 `driver_codegen_module_loop.inc:74`，API 在 `src/backend/detail/util/cgen_api.inc:217-221`）。
  2. **没有工程级"已发表"去重表**：`vb6_vtbl_<I>` 今天是**按实现类重复 typedef** 的。
     新式槽表若同样每个模块各发一遍，链接期必重复定义 → B04 需要一张工程级 set。
- 变量登记与派发落点（`As <Iface>` 走"薄指针"要动的四处 + 现值语义）：
  类型信息在 `Symbol::variableTypeName`（`src/semantics/semantic_analyzer_register.cpp:60-62`，
  `resolveTypeOrDefault`/`resolveTypeRef` 在 `semantic_analyzer_typeref.cpp:29`）；
  C 类型由 `src/backend/cgen_base_type.cpp:193-197` 决定，**当前是胖对 `vb6_iface_<I>` 按值**（门接条件是 legacy 的 `clsSym->isInterface` :194）
  → 薄指针 = 这里分叉（命中登记表走单字 `void*`）。登记 `knownIfaceVars_` 的四处：
  模块级 `src/backend/decl/cgen_decl_var.cpp:141-143`、局部 `src/backend/decl/cgen_localdecl.cpp:237-239`、
  跨模块 `src/backend/detail/base/cgen_base_generate_crossmod.inc:24`、形参
  `src/backend/decl/cgen_decl_func.cpp:109` / `cgen_decl_proc.cpp:116` / `cgen_decl_prop.cpp:117`。
  调用派发 = `src/backend/detail/expr/cgen_expr_call_com_bind.inc`（整体门控 `isComMarker_` :7；
  :67-97 查 `knownIfaceVars_` 后发 `x.vtbl->M(x.obj,…)` :91-93，标记来自
  `cgen_expr_member_obj_dispatch.inc:20-31`）；去虚化直调在 `src/backend/cgen_util_classcall.cpp:17-224`
  （`resolveClassMemberCall`，结果名 `vb6_<Cls>_<M>` 见 :115-118/:198/:223）；
  `Set x = y` 的接口改写 + `wrap()` 在 `src/backend/detail/stmt/cgen_setlet_set_prop.inc:387-435`
  （D3 记的 :419 已漂到 :435），`Set Nothing` :80-91。
- 宿主模块（B03 的 `isInterfaceModule`）**目前后端零消费**（`src/backend/**` 无人读它），
  `currentClassIsInterface`（`cgen_base_generate_decl_pass.inc:70-78`）只认 legacy 符号
  → B04 第一件事就是让宿主模块"只发槽表、不发类实例"，否则 `IWriter.cls` 会被当成普通类发码。
- 构建/管线接线成本很小：stage 4 入口 `src/driver/driver_compile.cpp:438-439` →
  `Driver::runCodeGeneration`（`src/driver/driver.cpp:111`，实现拆在 9 个
  `src/driver/detail/driver_codegen_*.inc`）；新增后端 `.cpp` 要进 `CMakeLists.txt` 的
  `vb6c3-cgen` 列表（:127-179，先例 `:163 src/backend/module/cgen_com.cpp`、
  `:173 src/backend/cgen_util_classcall.cpp`，注释行 :162）。
- 测试面（本轮再核实）：`Add-BasTest` 在 `tests/run_tests.ps1:644-652`（入队 `$basQueue`，
  由 :500 前结束的 `-Jobs` 并发 runner 消费；`test_interface` 的 x64/x86 双登记在 :727-728）、
  `Test-Vbp` :503-579（`itf_xmod` 已在 :800）、`Test-SyntaxFail` :584-601、`Test-Syntax` :603-617。
  **逐字节 emit-c 护栏**：`--emit-c` 开关在 `src/driver/driver_args.cpp:76`（`opts.emitC`，
  链接在 `src/driver/driver_link.cpp:57` 跳过）；8 文件清单与 worktree 基线做法见 D14-15 与
  B01 运行日志行（hello/test_rtl/test_array/test_error/test_ndarray/test_generics `.bas` +
  M6Test/test_implements `.vbp`）。B04 改结构体与派发路径 = **必跑**。
- 建议拆分：B04a = 槽表类型 + 宿主/类侧实例 + `isInterfaceModule` 后端接线 + 工程级去重表
  （只发码不派发，逐字节护栏先保住"无新语法工程零变化"）；B04b = 薄指针类型分叉 +
  `knownIfaceVars_` 四处登记 + 派发与 `Set`；B05 再管生命周期。两块各自过门。

### D20 B04（接口值代码生成）实施记录（2026-09-23）

1. **B04a/B04b 合并成一批**（D19 建议的拆分不成立）：只发槽表不派发没有任何**可观测行为**，
   过门等于没过门。实际节奏改成"先 emit-c 观察生成物 → 再接派发 → 一次过门"。
2. **落地后的内存布局**（`__comObj` 保第 0 位 = D19 硬修正的执行）：
   `vb6_cls_C { void* __comObj; vb6_ivref_<I> __iv_<I>; <原字段…>; }`，
   `vb6_ivref_<I>` 是**单词结构体** `{ const vb6_ivtbl_<I>* vt; }`，接口值 = `&obj->__iv_<I>`，
   适配器里 `offsetof(vb6_cls_C, __iv_<I>)` 做 container_of 再直调 `vb6_C_<M>(me, …)`。
3. **不需要 D19 说的"工程级已发表去重表"**：`vb6_ivtbl_<I>` / `vb6_ivref_<I>` 用
   `#ifndef VB6_IVTBL_<I>` 守卫，发在**每个模块头文件**里 → 同 TU 多次包含天然幂等，
   也不依赖 include 顺序；输出可复现靠"按小写接口名排序"（`unordered_map` 迭代顺序不稳定）。
   三个发射钩子在 `ivreg_` 为空时**零输出**，所以无新语法工程的生成物逐字节不变（已实测）。
4. **CCodeGen 接入方式**：`cgen.setInterfaceRegistry(&ifaces_)`，与 legacy 的
   `setInterfaceImplementers` 同一个注入点（`driver_codegen_module_loop.inc`）；
   新后端文件 `src/backend/module/cgen_iface_vtbl.cpp`（`CMakeLists.txt` 的 `vb6c3-cgen` 列表登记）。
5. **槽键口径只有一份**：成员名→槽键 = 先裸名（Sub/Function）再 `get_/put_/putref_` 前缀回退；
   `ifaceProcClauses` / `ifaceSlotPrefix` / `ifaceClauseSlotKey` 从 `semantic_analyzer_iface.cpp`
   上提到 `interface_sig.hpp`，语义层契约比对与后端绑定查找（`ivFindImplMember`）共用，
   包括"写了子句的成员不再参与同名隐式匹配"这条规则。
6. **实测生成物**（`tests/itf_xmod` 的 Main，`--emit-c`）：
   `vb6_ivref_IWriter* s = NULL;` → `s = &(w)->__iv_IWriter;  /* Set */` →
   `s->vt->emit(s, vb6_BSTR_FromStr(L"delta"));` / `s->vt->total(s)` / `s->vt->get_last(s)`；
   `Set s = Nothing` → `s = NULL`、`If s Is Nothing` 直接可用（薄指针语义免费送）。
   运行断言 `IFV1/IFV2/IFV3:OK` 已并入 `itf_xmod_writer` 的 `Test-Vbp`  needles。
7. **本批边界（后面批次补，别当已存在）**：
   ① 经接口变量**写**属性（`s.Title = v` → `put_` 槽）未接；
   ② 接口↔接口转换、`TypeOf … Is <接口>`（B06）；③ 真实 IUnknown/引用计数（三件套现在是
   `E_NOTIMPL` + 常量 1 的**占位**，槽号自此固定）（B05/B13）；
   ④ 宿主 `.cls` 仍按普通类发一个空结构体（无害，"宿主不发类实例"待 B05 一起做）；
   ⑤ `Set s = <Variant>` 右值推断没有 legacy 那样的唯一实现类兜底，直接报 ASCII 诊断
   `Interface binding: Set … needs a class instance or New …`，**该诊断暂无用例覆盖**。
8. **门数字不变**：129 项（本批只给 `itf_xmod_writer` 加了 3 条断言，没新增用例条目）；
   逐字节 emit-c 护栏 = 8 文件对 pre-B04 基线二进制（worktree @f8ca84e）**全同**。
9. **worktree 卫生**：基线 worktree 用完即 `git worktree remove --force` + `git worktree prune`，
   并确认 `git worktree list` 只剩主目录（历史上留过孤儿 worktree 迷惑后人）。

### D21 B05（生命周期：接口引用计数）实施记录（2026-09-23）

1. **落地形状**：`vb6_cls_C { void* __comObj; int32_t __refcount; vb6_ivref_<I> __iv_<I>; <原字段…>; }`，
   `_New()` 里 `__refcount = 1`。字段**只加在"实现新式接口的类"上**（`ivImplementedIfaces` 为空就一个
   字节都不发）→ 无新语法工程的生成物逐字节不变。这条门控是 D19「`__comObj` 必须第 0 字段」与
   「零回归逐字节护栏」同时成立的唯一解：既不能前置字段，也不能给所有类统一加计数头。
2. **AddRef/Release 必须按 (类, 接口) 各发一份**（B04 的占位三件套是按类一份）：`self` 是
   `&me->__iv_<I>`，container_of 的 `offsetof(vb6_cls_C, __iv_<I>)` 依赖具体接口，多接口类共用
   一份会算错实例地址。`QueryInterface` 仍 `E_NOTIMPL`（**槽号不动**），与 B06 的
   `TypeOf`/接口转换同批落地。
3. **归 0 销毁的守卫**：`if (me->__comObj != NULL) return 0UL;` —— 实例一旦被 COM 包装器接管，
   销毁权在包装器那一套计数（`vb6comserver_obj.c` wrapper refcount → `desc->destroyFunc`）。
   两套计数的统一是 P6/B13 的活；B05 的语义只在 EXE 直调路径上成立。
4. **引用语义三条**：
   - `Set <ivref> = <类变量 | New …>` 发成一个 C 块：存旧值 → 覆盖 → AddRef 新 → Release 旧
     （自引用安全）；RHS 是 `New` 时**不 AddRef**（`_New` 那 1 次初始引用直接移交）。
   - `Set <ivref> = Nothing`：先经槽 Release 再置 NULL。B04 之前这条落到 `vb6_ReleaseObject`，
     等于把薄指针当 `IDispatch*` 用（当时靠 `vb6_ComIsDispatchable` 拒绝才没崩）；现在是真释放。
   - `Dim … As <接口>` 局部变量在过程正常出口统一 Release：新 state 成员 `ivrefLocalsToRelease_`
     + `trackIvrefLocalForRelease()/emitIvrefScopeRelease()`，钩在 `ansiTempsToFree_` 的同一位置
     （Sub/Function/Property 三处出口发码、三处过程起点清空）。
5. **为什么不会误销毁（关键不变式，实测确认）**：类变量（`Dim w As C`）**从不**释放自己那份引用
   ——`vb6_cls_X_Destroy` 全项目只有两个调用点（COM wrapper 归零、UserControl terminate）。所以经
   类变量拿到的实例计数恒 ≥1，唯一能归 0 的路径是「`Set <接口变量> = New …` 且无别的接口变量共享」，
   这也正是本批能观测到 `Class_Terminate` 的唯一途径。（类变量泄漏是既有现状，不在本批处理。）
6. **`Exit Sub/Function` 走的裸 `return` 不清理**：与 ANSI 临时变量同一个既有缺口，本批按同一口径
   处理（不扩大），记为边界。
7. **只有裸标识符目标接计数**：`me-><字段>` 形式的接口变量（类模块成员）沿用 B04 的纯赋值改写、
   不做计数——`knownIvrefVars_` 按短名登记，`me->x` 命不中；留给后续批次。
8. **用例设计约束**：**不能**用 `Set q = p`（接口变量→接口变量）造"两个接口变量共享一个对象"的场景，
   B04 的诊断仍拒它（QI/转换要到 B06）→ 必须经类变量分发（`Set p1 = w2` / `Set p2 = w2`）。
   `CWriter.cls` 的 `Class_Terminate` 打 `TERM last=<m_last>`，用不同 `m_last` 让"哪个实例何时死"可辨：
   (a) `Set z = New` + `Set z = Nothing` → `TERM last=bye`（真归 0）；(b) 类变量 + p1/p2 全释放后
   `w2.Total()` 仍正确且无 TERM（不误销毁）；(c) `Sub ScopeExit` 不写 Nothing，纯靠过程出口 →
   `TERM last=scoped`。实测该工程**恰好 2 条 TERM**（`grep -c` 计数），顺序与预期一致。
   `itf_xmod_writer` 的 needles 由 5 条增至 11 条（本批 +6），**用例条目数不变**。
9. **护栏**：8 文件（hello/test_rtl/test_array/test_error/test_ndarray/test_generics .bas +
   M6Test/test_implements .vbp）`--emit-c` 对 pre-B05 基线二进制（worktree @6bc97e8，exe dc09fd9f）
   **全同**，`guard_fail=0`；worktree 用后即 `remove --force` + `prune`（D20-9 口径）。
10. **一处刻意留到 B06 的文案**：槽表 typedef 里的注释仍写
    `/* IUnknown prefix slots: placeholders in B04, real QI/refcount in B13/B05 */`。
    改它要重建二进制，而门与二进制一一绑定（D15-9），不值得为一条注释重跑 28 分钟全量门；
    B06 动 QI 时必然重编，那时一并更新（发射点是 `cgen_iface_vtbl.cpp` 的 `emitIfaceContractTypedefs`）。

### D22 B06a（QueryInterface / 跨接口 Set / `TypeOf … Is <接口>`）实施记录（2026-09-23）

1. **IID 只能按值比，且要按真 GUID 内存序发**：`vb6_iv_iid_<I>` 是 `#ifndef` 块里的
   `static const unsigned char[16]` → 同一接口在**每个编译单元各一份、地址不同**，比较必须走
   RTL 的 `vb6_IidEqual`（16 字节值比）。字节序 = Data1/2/3 小端 + Data4 原序，P6 交给真 COM
   时不必再翻。取值优先级：源码 `[InterfaceId("…")]`（`IfaceView.guid`，B01 起**只写不读**，
   本批起才有消费者）→ 否则 4 词 FNV-1a 从接口小写名派生（键 `"iviface:" + lower(名)`，与
   `cgen_util_dllentry_prelude.inc:56-70` 的 `generateIid` 同族；D8 禁随机）。
   IUnknown 用固定常量 `vb6_iv_iid_IUnknown`（自带 guard，全 TU 一份语义）。
2. **QI 与 AddRef/Release 同规格按 (类, 接口) 各一份**（D21-2 同一理由）：命中兄弟接口要
   `&me->__iv_<J>`，偏移依赖 J。为此 `emitIfaceImplTables` 顶部先发本类**全部 AddRef 的前向
   声明**——接口块的发射顺序与被调用顺序不一致，C 里用到必须先声明。语义：认 IUnknown / 本接口 /
   本类实现的其它接口，成功即对返回的那个指针 AddRef；其它 IID → `E_NOINTERFACE(0x80004002)`
   且 `*ppv=NULL`；`ppv`/`riid` 为空 → `E_POINTER(0x80070057)`。
3. **薄指针前 3 槽固定 = RTL 可以通用调用它**：RTL 新增 `vb6_ivtbl_prefix{QI,AddRef,Release}`
   + `vb6_IidEqual` + `vb6_IfaceSupports(ifacePtr, iid)`（声明在 `vb6rtl_class_com.h`、实现在
   `vb6rtl_com.c`；放 RTL 而非每个模块的 static，免 C4189/C4505 且只有一份）。
   `IfaceSupports` QI 成功后立刻 Release → 净效果只回答"支持不支持"，正是 VB6 `TypeOf … Is` 的口径。
4. **`vb6_TypeOf` 那个恒返 0 的桩刻意没碰**（`vb6rtl_conv.c:316-322`）：既有工程的
   `TypeOf x Is <类>` 一直恒假（`tests/BalloonTooltips/cTT.cls`、VBFlexGridDemo 在用），修它是
   **行为变更**，另批处理。本批只在 `visit(TypeOfExpr)` 开头加"右侧是新式接口"的分叉：左侧必须
   是接口变量，否则报 ASCII 诊断并返回 0（不静默给错答案）。该诊断没有 `--syntax-only` 级用例
   ——`Test-SyntaxFail` 只看语法/语义阶段，codegen 诊断够不着（同 D20-7⑤口径）。
5. **跨接口 `Set` 的引用口径**：发成一个 C 块 —— 存旧值 → `p->vt->QueryInterface(p,
   vb6_iv_iid_<目标>, &got)` → `q = (vb6_ivref_<目标>*)got` → Release 旧值。QI 成功已 AddRef，
   故 q 直接持有那份；失败得 NULL = Nothing（VB6 此处抛 438，本项目先按 Nothing 处理，
   错误码留给 P6/错误处理批次，已记边界）。
6. **用例形状**：`CWriter` 升级为双接口类（`IWriter` + 新增 `ILog`；`ILog` 带
   `Property Get Logged` → 顺带再验一次属性槽键 `get_Logged`）；另加**无实现类**的 `INope`
   （只进登记表、只出 IID）→ 负向 `TypeOf`/`E_NOINTERFACE` 有真实覆盖。新增 QI1..QI4 +
   TOF1..TOF3 共 7 条断言：TOF1 走"从 IWriter 指针 QI 到 ILog"、TOF2 走反向兄弟分支、
   TOF3 走 E_NOINTERFACE。实跑仍**恰好 2 条 TERM**（`last=bye` / `last=scoped`）→ 新增的
   QI/AddRef/Release 收支平衡，没有多一个少一个引用。
7. **本批边界**：① 下行转换 `Set <类变量> = <接口变量>` → B06b（见第 8 条）；②
   `TypeOf <类变量> Is <接口>`；③ 接口值进 Variant / 作实参传递的 QI 化（`cExprIsVariant`
   那条老路与薄指针不兼容，`vb6_ComIsDispatchable` 只查 7 槽，接口槽数少时会读到 vtable 之外）；
   ④ QI 认 IUnknown 时返回的是**本接口的薄指针**而非全对象统一的 IUnknown 指针（内部自洽；
   二进制 COM 兼容由 P6 的包装器负责）。
8. **下行转换（B06b）的真实障碍不是类型可见性，而是"身份验证"**（本条订正 Explore 报告的
   一处判断：`vb6_cls_<C>` 结构体是发在**类自己的 .h** 里的，别处 include 后就是完整类型 ——
   B04 生成的 `s = &(w)->__iv_IWriter;` 能编译过就是证据，所以使用点**看得见**
   `offsetof(vb6_cls_C, __iv_<I>)`，减偏移这件事哪个 TU 都能做）。真正的限制是：槽表实例
   `vb6_ivtbl_<I>_for_<C>` 是 C 的 .c 里的 `static const` 对象，别的 TU 拿不到它的地址，
   于是"这根薄指针到底是不是 C 的实例"没法在使用点判定。B06b 因此按类导出一个验明正身的
   助手（在 C 自己的 .c 里比较 `*(const void**)self == &vb6_ivtbl_<I>_for_<C>`，命中才减
   offsetof 返回实例，否则 NULL），声明放类的 .h 里跨模块可见；不要写成裸强转（对象不是那个类
   时会静默指到别处去）。上行方向（`TypeOf <类变量> Is <接口>`）同一套助手反着用即可。

9. **B06b 设计定稿（下一轮直接动工，不用再调研）**：按类在 `emitIfaceImplTables` 里多发两个
   跨模块助手（声明进类的 .h，实现进类的 .c，与槽表实例同 TU）：
   - `void* vb6_iv_from_iv_<C>(void* self)` —— 下行转换验身：比 `*(const void**)self` 与本类
     每个 `&vb6_ivtbl_<I>_for_<C>` 地址全等，命中才 `(char*)self - offsetof(vb6_cls_C, __iv_<I>)`，
     否则 NULL（= Nothing）。`Set <类变量 w> = <接口变量 s>` 发成
     `w = (vb6_cls_C*)vb6_iv_from_iv_C(s); if (w) s->vt->AddRef(s);`
     —— **必须 AddRef**：类变量那一遍引用永不释放（D21-5 不变式），不加码的话接口侧释放到 0
     会把对象从 w 脚下抽走。
   - `int32_t vb6_iv_test_iid_<C>(void* self, const void* riid)` —— `TypeOf <类变量> Is <接口>`
     用**静态** IID 归属判定（类变量的动态类型恒等于声明类型）：`self != NULL` 且 riid 命中
     IUnknown 或本类实现的任一接口 IID → 1。这条不需要减偏移，也就绕开了"字段是否存在"的
     编译期问题（`TypeOf w Is INope` 对不实现 INope 的类必须能编译并返回 False）。
   两者都只在 `ivImplementedIfaces(module)` 非空时发射 → 逐字节护栏照旧成立。
   静态契约（类不实现该接口时 `Set w = s` 要不要在编译期拒）留到 B13/P6 与 IID 表一起做，
   本批按运行期 NULL 处理并记边界。

10. **B05 留下的一个真缺陷，B06b 必须顺手修（已复现推理，未有用例）**：上转型发的是
    `s = &(w)->__iv_<I>;` —— `w` 是 Nothing（声明了但从没 `Set`）时，C 层面算的是
    `NULL + offsetof(...)` = 一个**非空野地址**，紧接着 `if (s) s->vt->AddRef(s)` 就解引用它 →
    崩溃。修法：发成 `s = (w) ? &(w)->__iv_<I> : NULL;`（三目即可，不需要临时变量）。
    配套负控用例：`Dim nz As CWriter`（永不 Set）→ `Set s = nz` → 期望 `s Is Nothing` 为真。
    `vb6_iv_from_iv_<C>` 同样要自守 `if (!self) return NULL;`。
    为什么 B05 的门没抓到：既有断言全部先 `Set w = New CWriter` 再上转，没有 Nothing 侧路径。

### D23 B06b（下行转换 / `TypeOf <类变量> Is <接口>` / 修 B05 野地址）实施记录（2026-09-23）

1. **按类导出两个跨模块助手**（`emitIfaceImplTables` 尾部发；声明进类的 `.h`、实现进类的 `.c`，
   与槽表实例同 TU）：
   - `void* vb6_iv_from_iv_<C>(void* self)` —— 薄指针 → 实例。先 `if (!self) return NULL;`，
     再取 `vt = *(const void**)self`，与本类每个 `&vb6_ivtbl_<I>_for_<C>` **地址全等**比对，
     命中才 `(char*)self - offsetof(vb6_cls_C, __iv_<I>)`。这就是 D22-8 说的"障碍是身份不是
     布局"的落地：表实例是 owning TU 的 `static const`，别处拿不到地址，所以判身只能在这里做。
   - `int32_t vb6_iv_test_iid_<C>(void* self, const void* riid)` —— 静态 IID 归属判定。
   两者都只在 `ivImplementedIfaces(module)` 非空时发射 → 无新语法工程生成物逐字节不变。
2. **下行转换**：`Set <类变量> = <接口变量>` 发成一个 C 块
   `void* __ivdown = vb6_iv_from_iv_C(s); w = (vb6_cls_C*)__ivdown; if (w) s->vt->AddRef(s);`
   —— **AddRef 是必需的**：类变量那一遍引用按 D21-5 的不变式永不释放，不给它加一次码，
   接口侧 `Release` 到 0 就会把对象从这个还活着的类变量脚下抽走（悬垂）。未命中得 NULL，
   即 VB6 的 `Set w = Nothing` 语义；VB6 这里其实抛 438，**静态契约拒绝留给 B13/P6**（记边界）。
3. **修掉 B05 的真缺陷（D22-10）**：上转型现在发成
   `s = ((w) ? &(w)->__iv_IWriter : NULL);`。守卫**只加在"裸类变量"这条路上** —— 给 `New`
   那条加三目会让 `_New()` 求值两次、凭空多创建一个实例（这个坑我在写的时候避开了，
   注释里写明原因）。配套负控用例 `DN0`：`Dim nz As CWriter` 永不 Set → `Set snz = nz` →
   `snz Is Nothing` 必须为真。**A/B 实测（同一份最小工程，只在 .build/negctl 里做，未入库）**：
   pre-B06b 基线二进制（@613d2b8，exe 67a28bfe）编译出的可执行文件 **段错误 exit=139**
   （= 0xC0000005，正是 `AddRef` 解引用 `NULL + offsetof` 那个非空野地址）；
   本批二进制（exe 4202570a）同一工程 exit=0 且打印 `NEGCTL:Nothing-OK`
   → 缺陷是真的、不是理论上的，且本批确实修掉了。`itf_xmod` 里的 DN0 就是这个 case 的常驻版本。
4. **`TypeOf <工程类变量> Is <新式接口>` 走静态 IID 归属**（`vb6_iv_test_iid_<C>`）：类变量的
   动态类型恒等于声明类型，所以这条**不需要减 offsetof**，也就不要求"这个类恰好实现该接口"
   才编译得过 —— 不实现时必须老实返回 False（用例 `TOC3` 用无实现类的 `INope` 打这一发）。
   左侧现在三类形态：接口变量（B06a，`vb6_IfaceSupports` 走真 QI）／类变量（B06b，静态判定）／
   其他（ASCII 诊断，仍无 `--syntax-only` 级用例覆盖）。`vb6_TypeOf` 那个恒返 0 的桩照旧一个字没动。
5. **两侧都只认裸标识符**：`me-><字段>` 形式的类字段／接口字段、属性目标、Variant 右值都不接
   （接口字段的声明侧 `Dim x As IWriter` 在类模块里会进 `knownIvrefVars_` 的短名，但 `Set me->x = …`
   的目标文本带 `me->` 前缀命不中 → 沿用 B04 的纯赋值路径，不计身份）。记为边界。
6. **实跑与负控**：`itf_xmod_writer` 断言 18→25 条（DN0..DN3 + TOC1..TOC3），实跑全中、
   仍然**恰好 2 条 TERM**（`last=bye` / `last=scoped`）→ 新增的下行转换 AddRef 与三处上转型
   守卫没有多一个或少一个引用。
7. **待回记**：门数字与逐字节护栏见总表 GATE_BASELINE（同批跑）。

### D24 B07 开工地图（2026-09-23 B06b 收尾轮产出；行号为 fa877ef/613d2b8 之后本轮 Explore 实测，含对 D2/D6 的硬修正）

- **`Inherits` 今天会被硬拒**（不是静默接受，好事）：`inherits` 词化成 `Identifier`，
  在 `src/parser/parser_module.cpp:187-191` 落到 fall-through，报
  `VB2002 unexpected token at module level: inherits`（`ParseUnexpectedToken=2002`,
  `src/common/diagnostics.hpp:49`），随后 `advance()+skipToNextLine()` → 整行丢弃。
  全仓（`src/` + `tests/` + `docs/vb6-manual/`）对 `Inherits`/`MyBase`/`Protected`/
  `Overridable`/`vb6_cvtbl_` 的命中数 = **0**，B07 起全是绿地。
- **语法接线 = 逐处镜像 `Extends`，四处**：`src/lexer/token.hpp:144-149`（枚举；
  `:149` 注释"仅接口域; 类继承用 Inherits, P3"就是本批的占位说明）、
  `src/lexer/lexer_keywords.cpp:51-54`、`src/parser/parser_helpers.cpp:56-57`（软关键字表，
  `canBeName()=Identifier||isSoftKeyword` 在 `:67-69`）、`src/lexer/token.cpp:31`
  （`isKeyword()` 显式列表）。枚举位置在 `TrueKeyword`(:24)~`GetObject`(:289) 区间内，
  所以 `token.cpp:7` 的区间判定自动覆盖。**切勿**进 `isStatementStart()`
  （`token.cpp:93-126` 连 `Interface`/`Extends` 都不在，D12 的"语句位永不识别"红线）。
  基类名解析照抄 `parseImplements()`（`parser_module.cpp:242-259`，Fix 083 的点号循环在 `:253-257`）。
- **D6 的头行接缝是错的，放松它=改既有 error path**：`parser_module.cpp:89-104` 那条
  `Class` 头行分支要求 `parseTypeParams()` 见到 `( Of`（`parser_decl_var.cpp:463-471` 否则返回空）
  → `:95-97` 直接报"泛型类头行需要 (Of T[,U])"。即今天 **`.cls` 首行写 `Class Foo` 本身就是错**。
  B07 若要支持 `Class Derived Inherits Base` 头行形式，必须放松该守卫 → 按 D9 跑逐字节护栏
  （`test_generics.bas` 在 8 文件清单内）并补一个泛型 `.cls` 用例。建议 v1 只做**独立子句行**
  `Inherits Base`（B01 的 `Interface…Extends` 也是两条路分开走的先例）。
- **AST 增量**：`Module` 加 `inheritsName`（`src/ast/detail/ast_decl.hpp:332-360` 字段区，
  `isClassModule:332`/`implements:349`/`interfaces:352` 旁），`ImplementsClause:40-44` 是"新结构体
  + clone 路径"的先例（B02b 动 3 处，`src/ast/ast_clone.cpp:707-720` 一带，见 D4）。
  `ast_printer_visitor_decl.inc` 既不印 `implements` 也不印 `interfaces`（grep 0 命中）→ 无需动打印器。
  下一个空闲**语义**诊断 ID = **3020**（`SemInterfaceClauseUnbound=3019`, `diagnostics.hpp:81`；4xxx 才是 codegen）。
- **类链求解照抄接口登记表的三趟结构**（`src/driver/driver_interface.cpp`）：Pass A 工程级唯一名 +
  与模块名冲突（`:29-84`，冲突判定 `:61-67`）、Pass B 父未知（`:90-94`）+ `seen` 集环检测
  （`:96-113`）、Pass C **父先序**展平 + `used` 键冲突诊断（`:115-167`，`chain.insert(begin)` 在 `:126`）。
  产出 = **每类一张只读"有效成员表"**（新 `ClassChainRegistry`，Driver 级，与 `IfaceRegistry` 平级），
  不改 AST、不动符号表枚举。理由与 D2 一致：每模块一张符号表而类名是工程级唯一。
  合并必须落在 **stage 3.5 之前**（`driver_compile.cpp`：2.6 泛型 :329 / 2.7 接口 :338 /
  3 语义 / 3.5 `runCrossModuleResolution` :356 / 3.5b :367 / 3.6 :372-387），否则外部工程的逐字段
  拷贝看不见派生域。单继承 =  arity 检查，另加深度上限。
- **D6 的合并字段清单不全**：`driver_crossmod.cpp:169-190` 是逐个字段手工拷贝外部 Class 符号，
  成员表实有 11 张（`symbol_table.hpp:116-197`）：`memberNames` `memberReturnTypes` `memberProcKinds`
  `memberParams` `memberFieldTypes` `memberFieldNames` `publicFieldNames` `memberFieldDispids`
  `memberLetParams` `memberSetParams` `eventNames`。漏一个 = 跨模块消费点瞎。**`interfaceMethodParams`
  是死字段**（D2 当时就记了从没被写入），别往里挂东西。
- **要"走基链"或吃合并表的函数（点名到行）**：`resolveClassMemberCall`
  （`cgen_util_classcall.cpp:17-224`；Fix 014 的 `classSym->memberNames` 兜底 `:161-200` 就是守门人，
  发出名在 `:115/:118/:198/:223`，注意它**不走** `cProcName`，是手拼 `vb6_<Cls>_<prefix><Member>`）、
  `findClassMemberCallParams`（`:237`，Phase A :248 / Phase B ~:407 / `memberReturnTypes` 兜底 :439）、
  `findClassMemberWriteParams`（消费点 `cgen_util_comwrite.cpp:275,:387`）、
  `getClassMethodReturnType`（`:400`）、`canonicalClassMemberName`/`canonicalClassFieldName`
  （`cgen_util_comwrite.cpp:178`）、`inferClassTypeOfExpr`（`cgen_util_classtype.cpp:15-124+`）、
  `knownClassVars_` 注册（`cgen_util.cpp:28`、`cgen_decl_{func,proc,prop}.cpp:104/110/112`）、
  driver 两张字段图扫描（`driver_codegen_typedfield_scan.inc:10-85`、`driver_codegen_voidfield_scan.inc:11-92`）、
  `classVariantMembers_` 扫描（`cgen_base_generate_state_scan.inc:56-114`）、
  `emitClassFieldAccessors`（声明 `cgen_helpers.inc:89`，调用 `cgen_base_generate_body_pass.inc:27`）、
  legacy Implements 覆盖检查（`semantic_analyzer.cpp:236-273`，扫 `memberNames` 在 `:271`
  → 继承来的实现**应当**满足契约，本批顺手让它走有效成员表）。
- **布局：不内嵌 `vb6_cls_Base`，改字段扁平复制**。内嵌一条就同时继承基类的 `__refcount` 与
  `__iv_<I>` 字段 → 双计数（撞 D21-1"一个类一个计数门禁"）、且 derived `_New` 要重复初始化继承槽的
  vtbl（`emitIfaceNewInit`, `cgen_iface_vtbl.cpp:353-365`）。B07 = 把基类私有字段按原序前置进派生
  struct 的用户字段区（发码点 `cgen_base_generate_c_open.inc:84-102`，字段 0 仍 `void* __comObj` `:76`，
  D19 不变）。
- **遮蔽判定必须大小写无关**（这是本轮发现的真实静默错字段风险）：`:84-102` 的字段循环**零去重、
  零归一**，而 `cIdent` 保留大小写（`cgen_base_naming.cpp:49-78`）、成员表键却是小写
  （`memberFieldNames`/`memberFieldTypes`）→ `m_X` 与 `M_x` 今天会发成**两个不同 C 成员**共享一个
  VB6 逻辑字段。合并时按小写键裁决胜者，struct 与所有表都用同一个拼写。
- **方法复用 = 转发桩，不做基类方法直调强转**。基方法符号是 `vb6_<Base>_<M>(vb6_cls_<Base>* me,…)`
  （属性 `vb6_<Base>_prop_get_<P>`），定义在基类 TU；`cProcName` 在
  `cgen_base_naming.cpp:249-265`（类方法一律带模块段），重载后缀 `_ov<fp>` 在 `:271-274`。
  `(vb6_cls_Base*)me` 强转有先例（`cgen_form.cpp:210,:237-250,:269` 的 `(vb6_cls_<ctl>*)me`），
  但它要求前缀布局逐字段一致，还得给每个调用点插桩；桩顺带绕过 `_ov` 后缀、B08 的 Overridable
  钩子、B09 的 MyBase 去虚化三件麻烦 → B07 发
  `vb6_<Derived>_<M>(vb6_cls_<Derived>* me,…){ vb6_<Base>_<M>((vb6_cls_<Base>*)me, …); }`
  只对**未被子类重名遮蔽**的继承成员生成。`me` 的类型来自 `classMeParam()`（`cgen_decl_prop.cpp:324-328`），
  `visit(MeExpr)` 在 `cgen_expr.cpp:396-411` 发裸 `me`。
- **B07 语义层有个真空**：全仓没有"项目类成员不存在"的诊断（`cgen_expr_member_precheck.inc` 只管
  `Err`/`VBA`/控件），今天写 `x.NoSuchMethod()` 会发成一个不存在的 C 调用 → **MSVC 编译错，不是 C3 诊断**。
  后果：负例断言只能挂在**本批新加的 3020 诊断文本**上，不能靠既有行为。
- **用例形态**：单 `.cls` 负例照 `tests/itf_neg/n08_missing_slot.cls`（`run_tests.ps1:866-890` 的
  `$itfNeg`，`Test-SyntaxFail` 在 `:584`，缺失文件→SKIP 不是 FAIL `:891-894`）——`Inherits NotThere`
  这类"单文件即可判定"的负例零成本。**但 shadow / 单继承 arity / 环 A↔B / 深度上限都是双文件用例**：
  `Test-SyntaxFail` 只接一个 `$Source`（`:590` 单引号参数）。两条路：给 helper 加文件列表
  （CLI 本身收多个位置参数：`driver_args.cpp:145` 逐个 push，`driver_frontend.cpp:149-163` 按扩展名定
  模块类型），或新写 `tests/cls_inh/*.vbp` + `Test-VbpFail`。B07 预算里含这个 helper。
  正例落地沿用 `tests/itf_xmod/` 的 vbp + `Test-Vbp` 断言（`:503`，只支持正向 needle、无 exclude）。
- **手册**：新建 `docs/vb6-manual/02-语句/Inherits 语句.md`（模板 = 同目录 `Interface 语句.md`，
  首行引用块标"本项目扩展"，末尾"实现状态"节，D18-5 的 CRLF 归一 + README `bare_lf==0` 复查照做），
  README 索引在 `:158`（`- [Interface 语句](…%20语句.md)`）旁按字母序插一行。
  **必须写明**：`Extends`=接口继承、`Inherits`=类继承，两者永不同义——018 自己 §二十一
  （`ai/讨论记录/018-接口继承与CoClass设计思路.md:1210-1224`）用 `Inherits` 写接口继承，
  而 `:1549-1551` 又把它划给类继承，实现按 `Extends` 落地，别照 018 的字面回头改。
- **D6 其余漂移**（按 D22-8 的规矩不原地改历史，只在此登记）：`getPublicSymbols` 实在
  `symbol_table.cpp:376`（D6 写 :383，7 行漂移）；D2 计划的 `SymbolKind::Interface` 至今没加
  （`symbol_table.hpp:23-45` 末位是 `ComGlobalNs`），B07 一切接口信息仍以 `IfaceRegistry` 为准。
- **护栏口径校正**：逐字节 `--emit-c` 对比是**每批手工仪式**，不是 CI/测试套属性
  （`grep emit-c tests/run_tests.ps1` = 0 命中），且其机理是**按 feature 早退**
  （`cgen_iface_vtbl.cpp:341-342/:354-356/:372-373`）。B07 的早退条件必须是
  "本类或其链用到 `Inherits`"，不是"工程里没有新语法"。8 文件清单里的
  `tests/test_implements.vbp` 是唯一带 `.cls` 的类工程输入，B07 开工时先**实测确认**它确实
  产出 class struct（否则"类布局未变"这句断言无覆盖）。
- **B07 边界（建议，超出的往后批）**：只做 `Inherits` 语法 + 链检测 + 有效成员表 +
  字段/方法继承的具体类调用；`Protected`/`Overridable`/`Overrides` = B08，`MyBase` = B09，
  接口实现经继承满足契约的**新式**路径与派生类 vtbl = B08+，CoClass = P5。

### D25 B07a（`Inherits` 语法 + 类链检测 stage 2.8）实施记录（2026-09-23）

1. **词法接线 = 逐处镜像 `Extends` 四处**，实测确认 D24 的判断：`token.hpp` 枚举插在 `Extends`
   之后 → 位置天然落在 `TrueKeyword..GetObject` 区间内，`token.cpp:7` 的区间判定自动覆盖，
   但仍照 `Interface`/`Extends` 的样子补进 `isKeyword()` 显式列表（保持一致性，不依赖区间）。
   软关键字表 `parser_helpers.cpp` 也补了 `Inherits` → `canBeName()` 仍认它，
   所以 `Dim inherits As Long`、`Implements inherits.X` 这类既有写法不受影响。
2. **零风险实证**：全仓 VB 语料（`*.bas`/`*.cls`/`*.frm`，含 demo 工程）里 `inherits` 这个词
   出现次数 = **0**（大小写不敏感、排除本轮新建的 `tests/cls_*`）→ 把它关键字化不改变任何
   存量输入的词形。护栏另用 8 文件 emit-c 独立证明（第 7 条）。
3. **只做独立子句行**（兑现 D24①）：`Inherits Base` 走模块级主循环新分支，位置就在 `Implements`
   之后；头行形式 `Class D Inherits B` **没有**动 —— 那条分支要求 `(Of T)`，放松它就是改既有
   error path。点号限定名（`Project.IBase`）与 `parseImplements` 同口径吃掉，登记表按整键 +
   末段两级解析（`ivLastSegment` 同思路）。
4. **`Module::inherits` 用值类型 `vector<InheritsStmt>`，不需要碰克隆路径**：`ASTCloner::cloneModule`
   本来就不拷 `implements`/`interfaces`（泛型特化副本不带契约），而泛型模板内的 `Inherits` 已在
   stage 2.8 拒绝 → 特化副本天然无继承，与克隆器现状自洽。B02b 那种"新结构体 + 3 处克隆"的工作量
   在这里省掉了，理由是**语义上特化类不该继承模板的继承关系**。
5. **单继承的 arity 检查不需要新代码**：一条子句里写 `Inherits A, B` 撞的是 parse 的
   `expectEndOfStatement()`（VB2003），"分两行各写一条"才由 2.8 报 3022 并只取第一条继续
   （不级联）。自环 `Inherits Self` 单文件即可判定 → 走了既有的 `Test-SyntaxFail` 通路。
6. **两条实测行为，都是刻意保留的**：
   - **一条环只报一份**：Pass B 按登记序解 `baseKey`，A↔B 互指时后解出的那个（PairB）才报
     `SemCircularDependency`，PairA 当时看到的 `baseKey` 还是空 → 不报。与 Extends 链同行为，
     不是漏报（`tests/cls_neg/ci_n06_*` 就是这一发的常驻用例）。
   - **接口宿主不能当基类**：`.cls` 里那一个同名 Interface 块不是实例类，Pass A 不登记它 →
     `Inherits IHost` 得到 3020 "must be a class module in this project"。文案说得过去
     （宿主确实不是可实例化的类），另开用例 `ci_n07_*`。
   深度上限 16（含自身）是**自保兜底**，VB6/tB 都没这个限制；链上每个节点自己数自己，
   所以报错的是最深的那个类。
7. **护栏口径补一个实测结论**（解除 D24 留的疑问）：8 文件清单里的 `tests/test_implements.vbp`
   确实产出 `typedef struct vb6_cls_CRectangle` / `vb6_cls_IShape` → **类布局在这份清单里是有覆盖的**，
   以后"没碰类结构体"这类断言可以拿它当证据。本轮 8 文件对 pre-B07a 基线（worktree @58f02fe，
   自建 Debug exe）`--emit-c` 输出逐字节全同。
8. **stage 2.8 的早退条件 = "工程里一条 `Inherits` 都没有"**（D24 末两条的兑现）：先扫
   `modules_` 再决定是否建表，因此非继承工程的生成物、诊断、阶段耗时无变化。
9. **用例通路扩到"多文件"**：新增 `Invoke-SyntaxProj` + `Test-SyntaxFailMulti` + `Test-SyntaxMulti`
   （C3 CLI 本来就收多个位置参数：`driver_args.cpp` 逐个 push、`driver_frontend.cpp` 按扩展名定
   模块类型）→ 双文件负例（互指环、基是接口宿主）与双文件正例第一次有了零构建通路。
   运行期正例 `cls_inh_pair`（`tests/cls_inh/Inh.vbp`）两条断言：`INH0:derived`（派生类自己能用）
   + `INH1:OK`（**基类自身的字段/方法发码没被派生类影响** —— 这条是对照组，不是新功能）。
   `run_tests.ps1` 登记为纯插入 +66/-0。条目基线 129 → 138（+8 语法 +1 vbp）。
10. **工具链教训（本轮连踩两次，务必记住）**：**用 bash heredoc / `printf` 写 `.bat` 会被 MSYS 改写**
    —— `>nul` 变成 `>/dev/null`、`\2019` 被当八进制转义、`\v` 变成垂直制表符 → `call vcvarsall` 静默
    失败打印"系统找不到指定的路径"，cmake 随后报 `No CMAKE_CXX_COMPILER could be found`，
    看起来像"VS 被卸了"。**基线构建 bat 一律用 python 以 r-string + CRLF 写**，写完断言
    文件里不含 `/dev/null`。
11. **`tests/run_tests.ps1` 是 UTF-8 带 BOM**（`git show HEAD:… | head -c 3` 实测，至少自 B01 起就是），
    项目记忆里"无 BOM 的 GBK"那条已经过期 → 编辑时**保留 BOM**（二进制读 `[3:]`、写回补
    `\xef\xbb\xbf`），"非 ASCII 字节数不变 + 无裸 LF + numstat 纯插入"三条断言照旧有效。
12. **B07b 待做**（本批刻意不碰，避免把合并与语法混在一个门里）：成员合并与遮蔽裁决（11 张成员表、
    大小写键）、继承字段进派生 struct（扁平复制，不内嵌）、继承方法转发桩、
    `resolveClassMemberCall` 等消费点走链、legacy Implements 覆盖检查吃合并表、
    手册页 `docs/vb6-manual/02-语句/Inherits 语句.md`（"实现状态"节要写真实进度，所以随 B07b 一起落）。
13. **门数字**：`Results: PASS=137 FAIL=0 SKIP=1 TOTAL=138`（`.build/gate_B07a.log`，11:25:47 起跑；exe md5 `031f993b` 跑前后一致 → 可归因）。条目 129→138（+8 语法 +1 vbp），零新增失败；legacy `test_implements` 仍 PASS。

### D26 B07b 开工地图（2026-09-23 B07a 收尾轮产出；行号本轮实测）

- **本轮已就位的东西**：`Module::inherits`（值类型 `vector<InheritsStmt>`）、stage 2.8
  `runClassChainPrepass` → `Driver::classes_`（`ClassChainView{name, mod, clause, baseKey,
  baseText, chain(父先己后的小写键), chainBroken}` + `classOrder_`），诊断 3020/3021/3022。
  **B07b 起 `classes_` 才有消费者**：注入点与发码点都要拿它，记得在 `driver_semantics.cpp`
  注给 analyzer（照 `ifaces_` 的做法）+ `CCodeGen` 侧加 `clsreg_` 指针（照 `ivreg_`）。
- **合并必须落在 stage 3 之后、3.5 之前**（B07a 的 2.8 只有"链"，没有"成员表"）：
  成员表是 `SemanticAnalyzer::analyze` 的类模块块（`semantic_analyzer.cpp:35-162`）在
  `symTab_.define` 之前逐声明填出来的，基类自己的表要等基类那轮 analyze 跑完才存在，而
  `modules_` 顺序不保证基先派后 → 合并做成**新的一轮 3.4**（遍历 `analyzers_`，
  按 `classes_[key].chain` 从根往叶把"父表"并进"子表"）。放在 3.5 之前才有意义的原因：
  `driver_crossmod.cpp:169-190` 是把 Class 符号的 11 张成员表**逐字段手工拷贝**给外部工程，
  合并晚于 3.5 就得再抄一遍外部副本。
- **并表规则**（v1）：键 = `Symbol::toLower(成员名)`；**子优先**（子已占的键父不再占），
  父先己后的顺序只在"两边都没写过"时决定 `memberNames` 的次序（对 legacy 契约检查
  `semantic_analyzer.cpp:236-273` 的扫描序可见，不影响新式接口槽序 —— 那是 `IfaceRegistry` 的事）。
  11 张表都要并：`memberNames` `memberReturnTypes` `memberProcKinds` `memberParams`
  `memberFieldTypes` `memberFieldNames` `publicFieldNames` `memberFieldDispids`
  `memberLetParams` `memberSetParams` `eventNames`（`symbol_table.hpp:116-197`；
  `interfaceMethodParams` 是死字段，别碰）。
- **发码侧要动的三处**（实测锚点）：
  1. 结构体字段：`cgen_base_generate_c_open.inc:84-102` 的字段循环只遍历
     `module.declarations` → 改成"先按链序发祖先的非遮蔽字段，再发本模块的"。
     `cIdent` **保留大小写** → 同名不同拼写要在合并阶段就裁决掉，发码只认胜者，
     否则一个 VB 字段发成两个 C 成员（D24③）。
  2. 转发桩：类过程体的发码点在 `cgen_base_generate_body_pass.inc`（`emitClassFieldAccessors`
     在 :27 被调），桩 = 对"链上祖先的、未被本类遮蔽的每个过程成员"生成
     `vb6_<D>_<M>(vb6_cls_<D>* me, …) { return vb6_<B>_<M>((vb6_cls_<B>*)me, …); }`；
     属性要按 `get_/put_/putref_` 三个方向分别发（`memberLetParams`/`memberSetParams` 已经带方向）。
  3. 成员查找：`resolveClassMemberCall` 的 Fix 014 兜底（`cgen_util_classcall.cpp:161-200`）
     读的就是 `classSym->memberNames` —— **合并进符号表之后这里一行都不用改**，
     `findClassMemberCallParams` / `getClassMethodReturnType` / `canonicalClassMemberName`
     同理（这是把合并做在符号表而不是做在发码层的最大收益）。
- **跨 TU 的可行性实测**（这条把 D24④ 的不安解除了一半）：桩里的
  `(vb6_cls_<B>*)me` 只要 `vb6_cls_<B>` 这个**类型名**可见即可 —— 多模块工程里
  `driver_codegen_module_loop.inc:47-59` 会把**所有**其它模块名塞进 `externalModules`，
  `cgen_base_generate_crossmod.inc:46-50` 于是给每个模块的 `.h` 都 `#include` 其它类头，
  而循环 include 时拿到的是 `cgen_base_generate_epilogue.inc:62-67` 那份
  `typedef struct vb6_cls_B vb6_cls_B;` **不完整类型** —— 指针转换对不完整类型合法，
  **成员访问才非法**。所以桩能编译 —— 最小 `cl` 探针（`typedef struct vb6_cls_B vb6_cls_B;` +
  `(vb6_cls_B*)me` 传给 `int vb6_B_M(vb6_cls_B*, int)`）零诊断通过；但 B07b 开工第一件事仍是
  在**真实管线**里实测这一条（最小工程：两个 .cls + 一条继承方法调用，看 cl 是否报 C2223/C2156）。
  本轮 `--emit-c` 观察（`tests/cls_inh`）：两个 `.h` **互相** `#include`（`InhBase.h` 里先
  `#include "InhDerived.h"` 再发自己的 struct），带 include 守卫时**谁先被包含谁就只拿到对方的
  前向 typedef** → 派生 TU 能否看到基类完整定义取决于包含序，不能只靠这条静态观察下结论。
  若实测发现基类不完整，改法是给派生模块的 `.c` 在头部**先**显式 `#include "<Base>.h"`（只此一条，
  不开泛化包含），或把桩发进基类 TU（代价：跨 TU 的符号可见性与 `_New` 顺序都要重看）。
- **v1 边界（写进 2.8 的 3022 里，别悄悄降级）**：
  ① 基类有 `EventDecl` → 拒绝（事件继承的语义 VB6/tB 都没定，且 `events` 字段位置一错就撞 D19 的偏移 0 不变式）；
  ② 基类或派生类实现**新式 `Interface`** → 拒绝（`__refcount`/`__iv_<I>` 前缀复制规则要与 D21-1 的
     "一个对象一个计数门禁"一起设计；`IfaceRegistry` 在 2.7 已就绪，判定条件现成）；
  ③ 基类是 legacy 接口类（被 `Implements` 当接口的 `.cls`）→ 允许（它的成员都是普通字段/过程）；
  ④ 泛型模板类的特化副本不参与继承（2.8 已拒模板内的 `Inherits`）。
- **前缀布局兼容**：v1 只有 `void* __comObj`（字段 0）+ 用户字段 → 派生类 = 祖先字段序 + 自身新字段，
  `(vb6_cls_<B>*)` 看到的偏移天然一致。②的拒绝把 `__refcount`/`__iv_` 的排布问题整体推到 B08+。
- **用例形态**：`tests/cls_inh/Inh.vbp` 已在门内（`INH0:derived` + `INH1:OK`）。B07b 往
  `InhDerived.cls` 加"经继承来的字段与方法"的断言（`INH2..INHk`），并补两条对照：
  遮蔽时子胜（子类自己写同名成员）、以及 `Set d = New InhDerived` 之后**基类自己的实例**
  仍走自己那份发码（防止桩把两个类的符号混起来）。负例走新加的
  `Test-SyntaxFailMulti`（①②两类都是双文件）。
- **手册**：`docs/vb6-manual/02-语句/Inherits 语句.md` 随 B07b 落（"实现状态"节要写真话），
  README 索引 `:158` 旁按字母序插一行；目录全 CRLF，写完归一并复查 `bare_lf==0`（D18-5）。
  页面上必须写清 `Extends`=接口、`Inherits`=类（018 §二十一自己把 `Inherits` 用作接口继承，
  实现走的是 `Extends`，别照字面回头改）。
- **逐字节护栏**：B07b 动结构体与发码 → 必须跑 8 文件清单（`test_implements.vbp` 已实测确认
  产出 `vb6_cls_CRectangle`/`vb6_cls_IShape`，类布局有覆盖，D25-7）。基线 = pre-B07b 的 worktree exe，
  **bat 用 python 写**（D25-10 的 MSYS `>nul` 坑）。

### D27 B07b（继承成员合并 stage 3.4 + 前缀布局 + 转发桩）实施记录（2026-09-23）

1. **合并落点 = 新 stage 3.4 `Driver::mergeInheritedMembers()`**（`driver_compile.cpp` 插在阶段3 与
   3.5 之间，理由照 D26：`driver_crossmod.cpp:169-190` 把 Class 符号成员表逐字段拷给外部工程副本）。
   消费的是 2.8 的 `chain`（父先己后），回填到 `ClassChainView::inhFields/inhProcs`，发码层只读这两张清单
   —— 这样"结构体字段"与"可调用成员"不可能各按一套规则走。
2. **v1 实际并 8 张表而不是 11 张**（过程面 6：`memberNames` `memberProcKinds` `memberReturnTypes`
   `memberParams` `memberLetParams` `memberSetParams`；字段面 2：`memberFieldNames` `memberFieldTypes`）。
   刻意不并的 3 张各有硬理由：`publicFieldNames` + `memberFieldDispids` 是 Fix 099 的 COM 对外暴露清单，
   并了就会让 `dll_entry` 去找派生 TU 根本不发的字段访问器（**链接期才炸**，比编译期难查）→ 归 P6/B13；
   `eventNames` 无需并（带事件的基类已在 2.8 判死）。字段**不进** `memberNames` 是 Fix 099 的硬规定
   （那张表被用来判定"成员访问是否为属性调用"，并字段会把 `Foo(obj.Field)` 改写成 `prop_get_` 调用）。
3. **同名 Property 的 Get/Let/Set 是"一个键、多个方向节点"**：第一版按成员键去重 `own.procs` →
   `Property Let` 的桩整个丢失（只有 `prop_get_Name` 发出来）。改成过程桶**不去重**、成员表按
   `mergedKeys` 每键只并一次、桩按节点逐方向发。这条只有跑真实工程才会露出来（INH9）。
4. **遮蔽裁决 = 近者（叶方向）优先抢键，落表仍按根→叶**：先"叶→根"扫描抢键（`taken` 集合），再按祖先下标
   `stable_sort` 落表，于是 `memberNames` 的次序是"根先、同层声明序"（对 legacy 契约检查的扫描序可见，
   不影响新式接口槽序）。层内多方向不互相遮蔽（`mine` 集合在整层扫完才并入 `taken`）——第一版在层内就
   `taken.insert` 导致第二个方向被自己挡住。
5. **布局 = 前缀复制，祖先的 `Private` 字段也复制**：派生 struct = `__comObj` + 祖先字段（根→叶）+ 自有字段。
   私有字段虽然对派生类不可见（不进成员表可见面），但**必须占布局**，否则基类实现自己的 `me->m_x` 全错位。
   `_New()` 的字段初始化与 `_Destroy()` 的 BSTR/数组释放同样改走 `structFieldDecls()`（继承在前），
   否则继承来的 String 字段既不初始化也不释放。
6. **转发桩**（`src/backend/module/cgen_inherit.cpp`，新独立编译单元）：
   `vb6_<D>_<M>(vb6_cls_<D>* me, …) { vb6_<B>_<M>((vb6_cls_<B>*)me, …); }`，属性的 C 名带
   `prop_get_/prop_let_/prop_set_` 前缀（与 `makePropertySignature`、`ivImplCName` 同口径，否则链接期找不到）。
   形参表与转调实参的口径直接照抄 `ivParamDecls`/`ivForwardArgs`（含 **Optional 的 `int _has_x` 尾参**
   ——不跟着转发基类的 `IsMissing` 就失真，INH6/7/8 专打这一发）。那两个函数在 `cgen_iface_vtbl.cpp` 的
   匿名 namespace 里、跨 TU 取不到 → 本文件重写了一份 5 行的 `paramsOf`，**没有**为此改 B04 的热点文件。
7. **跨 TU 可见性实测（解除 D24④/D26 的悬念）**：不给派生 `.c` 加任何显式 `#include "<Base>.h"`，
   现有"多模块工程互相 include 类头"就够 —— 桩只需要类型名（指针转换对不完整类型合法），而
   **继承来的 UDT 字段**要的完整类型也在基类 `.h` 里、经那道 include 可见（`_New` 发的
   `memset(&me->m_pt, 0, sizeof(me->m_pt))` 在派生 TU 编译通过 = 实证）。真实证据是运行期的
   `INH3`（`d.SetPt 5` → `d.PtSum() = 15`，走基类实现读写继承来的 `TPoint`）——编译通过不等于布局对。
   如果哪天包含关系变了，INH3 会先炸，这就是它当常驻用例的理由。
8. **本轮最有价值的发现：裸名调用继承成员是**静默错代码**，不是编译错误**。`Bump`（不带 `Me.`）在
   `Option Explicit` 下只给一条 VB3001 警告，`cl` 零诊断、exe 照生、**那段调用直接消失**
   （实测 `INH2:FAIL n=0`）。根因：合并只发生在 Class 符号的成员表上，模块作用域里没有这个过程的
   `Symbol`，发码侧认不出这个名字。已在 `visit(IdentifierExpr)` 的"未找到标识符"分支升格为
   VB3022 错误并要求写 `Me.` *成员名*（`SemanticAnalyzer::declaredByAncestor` 只看 2.8 的链与祖先声明，
   不复用 3.4 的回填 —— 判定发生在语义分析途中，那时合并还没跑）。手册"注意"节同步写明。
9. **v1 边界（2.8 Pass D，全部 3022，一个类只报第一条）**：① 基类带 `Event`（`events` sink 挂在结构体
   尾部，前缀复制不成立）；② 基类**或派生类**实现新式 `Interface`（`__refcount`/`__iv_<I>` 跟着复制 =
   一个对象两份计数，撞 D21-1 单门禁）；③ 基类有同名重载过程（桩要带 `$ov$` 变体后缀，与 B08 一起去虚化
   时一起设计）；④ 派生类重声明祖先同名字段（C 结构体容不下两个同名成员）。判据一律只看 AST。
   第一版 `hasOverloadedProcs` 把"同名 Property 的 Get+Let 两个节点"误判成重载，整条链被判死 ——
   与第 3 条同一个坑，两处都得按"成员键"算。
10. **用例与门**：`tests/cls_inh` 从 2 类扩成 3 类链（`InhBase ← InhMid ← InhDerived`），断言 2→12
    （INH0/1 原有；INH2 继承 Sub/Function+私有 Long；INH3 继承 UDT 字段；INH4/5 子胜遮蔽且基类实例不受影响；
    INH6/7/8 Optional 转发；INH9 属性 Get/Let 双向；INH10 中间层成员与祖父成员并存；INH11 实例隔离）；
    新增 3 条双文件负例 `ci_n08_bare_inherited_call` / `ci_n09_redeclared_field` / `ci_n10_event_base`，
    走 B07a 刚建的 `Test-SyntaxFailMulti`（`--syntax-only` 通路，无需 vbp）。`run_tests.ps1` 登记 +16/-1
    （那 1 行删除是把 `Test-Vbp` 的断言列表改成多行）。B07b 后语法类 65/0/0。
11. **工具链新坑（本轮实测）**：`scripts/build.bat` 是 **LF-only + GBK 注释**，从 agent 的 bash 里
    `cmd //c scripts\build.bat` 会被 cmd 在注释的 `）`/`(` 字节上截断 `if (...)` 块，产出一堆
    "'_BIN' 不是内部或外部命令" 式噪声、cmake 那侧看似 VS 丢失。**构建一律用
    `powershell -NoProfile -File scripts/dev.ps1 -SkipTest`**（它打 "All done!"，与既有门一致）。
12. **护栏**：8 文件 `--emit-c` 对 pre-B07b 基线（worktree @`a056705`，自建 Debug exe）**8/8 逐字节全同**；
    基线 worktree 用后即 `git worktree remove --force` + `prune`。
13. **门数字**：`Results: PASS=140 FAIL=0 SKIP=1 TOTAL=141`（`.build/gate_B07b.log`，12:47:44 起跑、13:1x 收；exe md5 `50ce2e77` 跑前后一致 → 可归因）。条目 138→141（三条 v1 边界负例），零新增失败；语法类单跑 65/0/0；legacy `test_implements` 与 `itf_xmod_writer`（25 断言）仍全绿。
14. **B08 的既有事实（本轮顺手实测）**：`AccessLevel` 只有 `Public/Private/Friend`（`types.hpp:54-59`）、
    全仓**没有任何"成员不可访问"的诊断**（`Protected`/`Overridable`/`Overrides` 三个词在 `src/` 里只出现在
    错误处理块的 `inProtectedBlock_`，与继承无关）→ B08 的可见性是**从零加检查 + 新诊断**，不是改口径。

### D28 B08 开工地图（2026-09-23 B07b 收尾轮产出；锚点本轮实测）

- **先拆批**（判据仍是 D20-1"每批都要有可观测行为"）：
  - **B08a = `Protected` 可见性**（语法 + 访问权限检查 + 新诊断）。可观测性天然成立：今天
    `Protected` 这个关键字**根本不存在**，写了就在 parse 期撞 VB2002，加了就能解析 + 报"不可访问"。
  - **B08b = `Overridable`/`Overrides`/`NotOverridable` + 类虚表 + 动态派发**。它要改结构体
    （多一个 `__cvtbl` 字段）与 `Me.` 调用点，必须与 B08a 分开过门、分开跑 8 文件护栏。
- **B08a 的四处接线 = 逐处镜像 `Friend`**（本轮实测锚点）：`src/lexer/lexer_keywords.cpp:31`
  （`{"friend", TokenKind::Friend}` 旁）、`src/lexer/token.hpp` 枚举、`src/lexer/token.cpp:25`
  （`isKeyword` 显式列表）、`src/parser/parser.cpp:288` 与 `src/parser/parser_decl.cpp:28-39`
  （访问修饰符 parse 分支）。`AccessLevel` 在 `src/common/types.hpp:54-59`，只有
  `Public=0/Private=1/Friend=2/Default=Public` → 新值取 `Protected = 3`（**别插在中间**，
  `Default = Public` 的别名与任何按数值的比较都会被插队影响）。
- **两个零成本的前车之鉴**（B07a/B07b 各踩一次）：① 新关键字**同时**进
  `src/parser/parser_helpers.cpp` 的软关键字表（`isSoftKeyword`/`canBeName`），并且**实测**全仓 VB 语料
  里该词出现次数 —— `inherits` 与 `protected` 本轮实测都是 **0 次**，`overridable`/`overrides` 待测；
  ② `isStatementStart()` 刻意不含这些词（D12"语句位永不识别"），别顺手加。
- **可见性检查是从零加、不是改口径**（D27-14）：全仓没有任何"成员不可访问"的诊断，`Private` 的
  "不可见"目前只体现为**跨模块注入时不收进去**（`driver_crossmod.cpp` 的 Public 过滤）。所以 B08a 要新加
  诊断 ID（`diagnostics.hpp` 现用最大 3022 → `SemMemberNotAccessible = 3023`、
  `SemOverridesMismatch = 3024` 预留），检查点候选：`resolveClassMemberCall`（成员访问）+
  字段访问的规范化点（`canonicalClassMemberName` 一族），两处都要能看到"访问者所在模块 vs 成员所属类"。
  `Protected` 的判定口径 = 当前模块是基类本身、或**在继承链上**（读 `Driver::classes_[key].chain`）→
  这也是 3.4 之外第二个消费 `chain` 的地方，注入路线照 `analyzer->setClassChainRegistry(&classes_)`
  （`driver_semantics.cpp:44`）。
- **B08b 的挂钩点早就备好了**：B07b 的转发桩（`cgen_inherit.cpp`）是"派生类调用祖先实现"的唯一出口，
  去虚化 = 让 `Overrides` 的桩改指本类实现、动态派发 = 让**基类体内**对 `Me.M` 的调用改走 `__cvtbl` 槽。
  要动的三处：① 结构体加 `__cvtbl`（**位置硬约束**：字段 0 仍是 `__comObj`（D19），`__iv_<I>` 紧跟其后
  是 B04 定死的 → 新指针只能排在 `emitIfaceClassFields` 之后，且**只给链上含 `Overridable` 的类**加
  —— 与 `__refcount` 同一套"按 feature 加字段"的保护栏手法，见 D21）；② `resolveClassMemberCall`
  （`cgen_util_classcall.cpp`）读 Class 符号判虚；③ 虚表实例按类一份、函数指针槽序稳定
  （根→叶、同层声明序 —— 就是 3.4 落表顺序，别另起一套）。
- **`Overrides` 的签名比对用源码签名口径**（与 D16 同一理由：跨模块 Enum/UDT 在 stage 3 早期尚未注入，
  归一成 `Vb6Type` 再比会误判）→ 直接复用 `src/semantics/interface_sig.hpp` 的 `ifaceSigFromDecl`，
  那里是签名文本的唯一出处。
- **继承与新式接口的互斥要重新审视**：B07b 在 2.8 Pass D 判死了"任一侧实现新式 Interface"（D27-9②），
  若 B08b 的虚表与 `__iv_<I>` 要在同一个类上共存，`__refcount`/`__cvtbl`/`__iv_` 三者的定序要一次定清
  （写进 D29，别在实现中途改）。
- **门与护栏**：B08a 不动发码 → 8 文件护栏可省（按 D18-6 的口径记录理由 + 用真实工程编译替代）；
  B08b 动结构体与派发 → **必须**跑，基线 exe 用 worktree（bat 用 python 以 r-string + CRLF 写，D25-10；
  构建本身用 `powershell -File scripts/dev.ps1 -SkipTest`，D27-11）。
- **用例形态**：B08a = `tests/cls_neg/ci_n11_protected_*`（跨模块访问 Protected → 3023）+
  `ci_pos_protected`（链内访问放行）；B08b = `tests/cls_inh` 再加一层（`InhBase` 的 `Overridable Sub Speak`、
  `InhDerived` `Overrides Speak`）+ 断言"基类指针调 `Speak` 也走派生实现"—— 这条是唯一能证明真虚派发的形状。

### D29 B08a（`Protected` 访问级别 + 家族内可见）实施记录（2026-09-23）

1. **拆批又拆了一层**：D28 原本把 B08a 定为"`Protected` 可见性"，做完才发现"可见性"有两半 ——
   "家族内可用"（本批交付）与"家族外越权要拒绝"（记为 **B08c**）。后者今天做不了：语义层
   `visit(MemberAccessExpr)` 根本不把 `obj` 解析成 Class（`semantic_analyzer_expr.cpp:141` 一路落到
   `lastExprType_ = Variant`），要拒就得给分析器补一套 obj→Class 解析；而"偷懒的做法"（在
   `driver_crossmod.cpp` 的逐字段拷贝里按访问级别过滤成员表）会让越权访问**退化成晚绑定 COM 调用**
   —— 从"太宽松"变成"运行期才炸"，比静默更糟。所以本批只做能诚实交付的那一半，边界写进手册。
2. **接线 = 逐处镜像 `Friend`，一共 13 处**（其中三处是静默陷阱，不写下来一定会踩）：
   `types.hpp` 枚举、`token.hpp` 枚举、`lexer_keywords.cpp`、`token.cpp::isKeyword`、
   `parser_helpers.cpp` 软关键字表、`parser.cpp::isDeclarationStart`、`parser_decl.cpp` 修饰符分派、
   `parser_decl_var.cpp` 行首修饰符消费、`parser_interface.cpp` 接口成员修饰符拒绝、
   `symbol_table.hpp` 新表、`semantic_analyzer.cpp` 填表、`driver_crossmod.cpp` 拷贝、
   `driver_classchain.cpp` 合并 + `cgen_base_generate_decl_pass.inc` 的 .h static 判定 + `accessStr` dump。
   - **陷阱①**：`parser_decl.cpp:30-37` 是 `case` 列表 + if/else，**`else` 兜底是 Private**。
     加了 `case TokenKind::Protected` 而忘加 if 分支 → `Protected Sub X` 静默变成 Private（不报错、行为相反）。
     同样形状的还有 `parseAccessDeclInBody`（过程体内 `Public/Private x As Long`，本批刻意**不**支持体内 Protected）。
   - **陷阱②**：`isDeclarationStart`/`isStatementStart` 的 `default: return false` → 漏加就是 parse 错误。
     注意 `isStatementStart` 里连 `Friend` 都没有，所以 Protected 也**不进**（过程体内它不是语句起点）。
   - **陷阱③**：`.h` 声明的 static 判定枚举的是 `Public || Friend`，而 `.c` 定义的判定是 `== Private`。
     只加访问级别不改 `.h` 那三处，就会得到"声明 `static`、定义非 static"的对不上 ——
     Protected 必须与 Public/Friend 同等（派生模块 TU 的转发桩要能看见它）。`.c` 侧本来就对（非 Private 即非 static）。
3. **新成员表 `Symbol::memberAccessLevels`（键小写 → AccessLevel）只有 3 个 touch point**：
   声明 + 填表 + 那两处拷贝（crossmod 与继承合并）。之所以这么少，是实测确认**全仓没有 Symbol 的
   序列化/反射**（`getPublicSymbols` 只看符号自己的 `access`，不看成员表）。填表按既有口径：
   属性只在"读上下文胜出"时写级别（与 `memberProcKinds` 同一处 `if (wins)`）；事件也记。
4. **Protected 的真正语义落点 = stage 3.4 的可见面过滤**：那张过滤今天只挡 `Private`
   （`driver_classchain.cpp:401`），于是 Protected 的基类成员**自动**被派生类继承并可经 `Me.` 使用，
   一行都不用改；而它不进 `publicFieldNames`（填表处只认 `AccessLevel::Public`，Fix 099）→
   天然不进 COM 对外暴露清单与 TypeLib 表（`driver_codegen_dll_typelib.inc` 判 `== Public`）✓ 这两处
   "什么都不做"正是对的，写下来免得后人以为漏了。
5. **可观测性（D20-1）**：`Protected` 今天根本不存在（写了撞 parse 错误）→ 本批的可观测面是三件事：
   家族内可用（`INH12` 经 `Me.SetSecret/Me.Secret` 用基类 Protected 方法、`INH13` 用基类 Protected 字段）、
   `.h` 非 static 带来的跨 TU 链接（`INH12` 能跑出来就是它通了）、以及 `Interface` 块内仍被拒（VB2011，
   `ci_n11_protected_in_interface`）。门 `Results: PASS=141 FAIL=0 SKIP=1 TOTAL=142`；语法类 66/0/0；
   **护栏 8/8 逐字节全同**（基线 = worktree @`b1c0050`，自建 Debug exe）—— 本批动了发码判定，按 D9 必须跑。
6. **语料实测**：`protected` / `overridable` / `overrides` / `mustinherit` / `notoverridable` 五个词在
   `tests/**` 的 `.bas/.cls/.frm/.ctl/.pag` 里出现 **0 次** → 关键字化对存量输入零影响；照 D25 的口径
   仍然一并进了软关键字表（`canBeName` 仍认它）。
7. **委托勘察的复核教训**：派 Explore 摸"访问级别都在哪儿被消费"是对的（它给出的
   `getPublicSymbols`/TypeLib/static 三组分类直接用上了），但它的两条锚点不准 ——
   `token.cpp:97-98` 其实是 `isStatementStart`（不是 token 名转字符串表，那张表不存在），
   `driver_crossmod.cpp` 的拷贝写的是 `extSym->memberSetParams = srcSym->...`（不是它报的 `dst->/src->`）。
   **动手前逐条 grep 复核锚点**只花几分钟，而错锚点在带 `assert count==1` 的打补丁脚本里会直接 no-op。
8. **B08b 待做**（本批刻意不碰）：`Overridable/Overrides` + 类虚表 + 动态派发（地图 D30）；
   B08c 待做：家族外越权访问 `Protected` 的拒绝（要先给分析器补 obj→Class 解析）。

### D30 B08b 开工地图（2026-09-23 B08a 收尾轮产出；锚点本轮实测）

- **本批要解决的唯一硬问题**：今天 `Me.M` 在**基类体内**是静态绑定（`resolveClassMemberCall` 拿到
  当前类符号 → 直接发 `vb6_<Base>_M`），所以"派生类覆盖了 `M`、基类的 `Talk` 里调 `M` 仍走派生实现"
  这条真虚派发**必须**有一个运行期入口。方案：给需要虚表的类发一份 `vb6_cvtbl_<Cls>` +
  struct 里一个 `__cvtbl` 字段，基类体内对可覆盖成员的调用改走槽位。
- **字段位置是硬约束**（D19 + D21 + B07b）：字段 0 恒为 `__comObj`，`__iv_<I>` 紧跟其后是 B04 定死的，
  继承前缀布局又要求"祖先字段先于自有字段" → `__cvtbl` 只能排在 `emitIfaceClassFields` 之后、
  用户字段之前，且**只给链上含 `Overridable` 的类加**（照 `__refcount` 的按 feature 加字段手法，
  这样无新语法的工程逐字节不变）。定序一次写死：`__comObj` → `__iv_<I>` → `__cvtbl` → 祖先字段 → 自有字段。
- **虚表槽序 = 3.4 的落表序**（根→叶、同层声明序、每键一份），别另起一套计数；否则同一个基类在
  不同派生类下的槽号会漂移。**新增一张派生表**存"哪些键是虚的"：`Symbol::memberVirtual`
  （小写键 → 槽号 + 是否 Overridable/MustOverride），touch point 与 D29-3 的 `memberAccessLevels` 完全一样
  （声明 + 填表 + `driver_crossmod.cpp:188` 那块 + `driver_classchain.cpp` 的 `if (firstTime && src)` 块）。
- **AST 侧要加标志位**：`SubDecl/FunctionDecl/PropertyDecl` 各加 `bool overridable/overrides/mustOverride`
  （现成先例：`PropertyDecl::propKind` 与 `implementsClauses`）。parse 顺序是
  `[访问修饰符] [Overridable|Overrides|NotOverridable] Sub|Function|Property` —— 修饰符在
  `parser_decl.cpp` 消费后要**再看一眼**第二个修饰符位，漏了就是 VB2002。
  v1 边界照 B07b 的口径写进 2.8（`SemInheritsNotSupported`）：泛型模板内、非类模块内出现这些词 = 拒。
- **`Overrides` 的校验**：基类同名键存在、且基类那份是 `Overridable`（或 `MustOverride` 时派生必须写）、
  签名一致 → 用 `interface_sig.hpp::ifaceSigFromDecl` 的**源码签名**口径（D16 的理由：跨模块 Enum/UDT 在
  stage 3 早期未注入，归一成 `Vb6Type` 会误判）。新诊断从 **3024** 起（3023 预留给 B08c 的越权拒绝）。
- **派发点只有两处**（其余调用形状继续沿用 B07b 的桩）：① 类体内 `Me.M` / 裸名 `M`（后者今天被 D27-8
  拒了，正好少一处）；② 入口函数本身 —— 建议 `vb6_<Cls>_M` 保持"具体实现的直调"，虚派发放到
  "**基类体内对可覆盖成员的调用**"这一处，别让每个外部调用点都查表（那会把 B09 的 `MyBase` 去虚化
  逼成另一套命名）。
- **用例形态（这是唯一能证明"真虚"的形状）**：`InhBase` 加
  `Public Overridable Function Speak() As String`（返回 "base"）+ `Public Function Talk() As String`
  （`Talk = Speak()`，走 `Me.`）；`InhDerived` 写 `Public Overrides Function Speak()`（"derived"）。
  断言：`d.Talk() = "derived"`（虚）而 `b.Talk() = "base"`（基类实例不受影响）、
  `d.Speak() = "derived"`、`b.Speak() = "base"`。负例：派生类 `Overrides` 一个基类没有的成员、
  以及基类成员没标 `Overridable` 却被 `Overrides` → 各一条 3024（双文件，走 `Test-SyntaxFailMulti`）。
- **与既有边界的冲突要先解**：2.8 Pass D 目前判死"基类/派生类实现新式 Interface"与"基类有重载成员"。
  B08b 若想让 `Overridable` 与新式接口共存，得先把 `__iv_<I>` 的计数/槽语义和虚表对齐（那是 B08+/P6 的活），
  v1 建议**保持这两条拒绝不变**，只在纯具体类链上做虚派发。
- **护栏**：本批改结构体 + 派发 → 8 文件 `--emit-c` 必须全同（基线 = pre-B08b worktree，exe 自建；
  bat 用 python r-string + CRLF 写，构建本身用 `dev.ps1 -SkipTest`，见 D25-10 / D27-11）。

### D31 B08b（`Overridable`/`Overrides`/`NotOverridable` 语法 + 覆盖契约 + 封掉假虚派发）实施记录（2026-09-23）

1. **拆批拆到第三层**：D30 把 B08b 定成"修饰符 + 类虚表 + 动态派发"，本轮交付的是**能诚实交付的那半**
   （语法 + 覆盖契约校验），类虚表/派发独立成 **B08d**（地图 D32）。判据不是工作量而是 D27-13 那条判例：
   只上语法时，"基类体内 `Me.M`"会**编得过、跑出基类实现**（静态绑定），后代的 `Overrides` 静默失效 ——
   "看着像虚调用其实不是"比编译失败危险得多。所以本批的交付里**必须包含把这条路判死**，
   而不是"先收语法、以后再修"。
2. **修饰位建模成四值枚举 `ProcVirt{None,Overridable,Overrides,NotOverridable}`，不是三个 bool**：
   三件套互斥，写成三个字段就有 6 种非法组合要在 parse 与语义两层各防一遍。`None` 与
   `NotOverridable` 运行语义相同但**不合并**：后者是显式意图，"覆盖一个 `NotOverridable` 成员"要报得准。
3. **AST 侧零签名改动**：三个过程 decl 用 `ProcVirt virt = ProcVirt::None;` 成员默认值 + parse 后赋值
   （与 `typeParams`/`implementsClauses` 同一手法），于是 `parser_interface.cpp`（接口成员也 new 这三种
   节点）与 `ast_clone.cpp` 的构造点**一个都不用改**。`ast_clone` 刻意**不拷** `virt` —— 泛型模板内的虚
   修饰符已被 E0 拒绝，特化副本永远不会带它（同 B07 `Inherits` 的取舍）。
4. **parse 的修饰符要"吃两次"**：VB6 里访问修饰符可省（`Overridable Sub X`），所以起手先吃一轮，
   进了 `Public/Private/…` 分支、消费访问修饰符**之后**再吃一轮（`Public Overridable Sub X`）。
   `isDeclarationStart` 三个 case 必加（漏了就是 parse 错误），`isStatementStart` **不加**（过程体内它
   不是语句起点，与 `Friend` 同口径 —— D29-2 陷阱②的重复确认）。这次没有再踩"else 兜底 Private"，
   因为新枚举是三态显式分派。
5. **契约检查落在 2.8 新增 `Driver::runVirtualContractChecks`，不是 3.4**：它要回答"祖先**自己声明**过
   这个槽吗、那份带没带 `Overridable`"。放 3.4 之后 Class 符号的成员表已被合并污染（含祖先条目与遮蔽
   结果），"谁声明的"就查不回来；而 2.8 已解好 `chain`，直接读各模块 decl 就够。属性按 `ifaceSlotKey`
   的**方向**配对（`get_/put_/putref_`），签名比对复用 `ifaceSigFromDecl`/`ifaceSigEqual` —— 与 Interface
   契约检查同一套函数，不会出现两套判据漂移（D16 的源码签名口径同理由）。
6. **`dynamicKeys` 存成员名小写、不存槽键**：语义层判定发生在 `visit(MemberAccessExpr)`，那里只有
   `memberName`、拿不到"这次访问用的是哪个方向"（`Me.X` 可能是 Get 也可能是 Let/Set）。用名字做键
   = 属性三向一起判，方向偏保守（可能多拒），但**不会漏**——漏就是静默绑错。B08d 发虚表时要另起一张
   按槽键的表，别复用这张（D32-2）。
7. **拒绝点在语义层，一共三处**：`visit(IdentifierExpr)` 的"查到符号"分支（裸名值引用/隐式调用）、
   `visit(IndexOrCallExpr)` 的**裸 callee `if (sym)` 分支**（它自己 `symTab_.lookup`、**不经过**
   `visit(IdentifierExpr)` → 少埋这一处就漏掉 `Speak(5)` 这种带实参调用），以及
   `visit(MemberAccessExpr)` 的 `Me.X` 形状（`Me.X(…)` 经 `analyzeExpr(callee)` 落在这里，一处覆盖两形态）。
   选语义层而不是发码层的理由是**可测性**：`--syntax-only` 就能断言（见 [[c3-build-test-hazards]] 的
   "`--syntax-only` 不是 parse-only"），不必为编译期诊断去搭 `.vbp` 负例工程。
8. **本轮唯一"差点误杀正例"的点**：`Speak = "base"` 是 VB6 的**返回值赋值**，标识符恰好就是函数名，
   与"类体内调用 `Speak`"在 `visit(IdentifierExpr)` 里长得一模一样 → 最早的写法把 `Overridable` 成员
   **自己的函数体**判死了（第一次编译正例就报 VB3027）。修法是豁免 `currentProc_` 同名的裸名引用。
   记在这里是因为它不报错、只是"最正面的用法突然不能用"，很容易被当成边界凑合过去。
9. **早退条件扩成 `anyClause || anyVirtual`，但 E0（位置合法性）只读 `modules_`、不建登记表**：若为了 E0
   也去建链登记表，"写了 `Overridable` 但一条 `Inherits` 都没有"的工程会第一次拿到非空 `classes_`，
   后端 `classChainOf()` 从 nullptr 变成有值 —— 那是给存量工程开了一条从没走过的路径，正是要避免的
   "顺带改了行为"。现在：没有 `Inherits` 就在 `checkVirtualPlacement()` 之后直接返回。
   8 文件 `--emit-c` 逐字节全同（@`2117d1c` 基线）就是这条的证据。
10. **工具链一条**：基线构建别再试 `scripts/dev.ps1`（worktree 里它只 `cmake --build`、不 configure →
    `is not a directory`），也不用回退到手写 bat —— `.build\base_build.ps1`（vcvars + cmake -B + --target c3，
    纯 ASCII）一次通过，比 D25-10/D29 那套"python 写 bat 防 MSYS 改写"省事，因为 PowerShell 不经 MSYS。
11. **可观测面（D20-1）**：正例 `INH14`（派生对象上直调被覆盖成员 = 派生实现）、`INH15`（基类实例不受
    影响）、`INH16`（`NotOverridable` 成员仍按继承走转发桩）；负例 7 条 = 契约 4 条（`ci_n12` 链上无目标
    3024 / `ci_n13` 目标未标 Overridable 3025 / `ci_n14` 签名不符 3026 / `ci_n15` 类体内调用被覆盖成员 3027）
    + 位置 3 条（`ci_n16` 无 `Inherits` 却写 `Overrides`、`ci_n17` 标准模块里写 `Overridable`、
    `ci_n18` `Interface` 块里写虚修饰符）。诊断 ID 按 D30 的预约从 **3024** 起，**3023 仍留给 B08c**。
    门 `Results: PASS=148 FAIL=0 SKIP=1 TOTAL=149`；`-Category syntax` 66→73；`Inh.vbp` 断言 14→17 条。

### D32 B08d 开工地图（类虚表 `vb6_cvtbl_<Cls>` + 动态派发；2026-09-23 B08b 收尾轮产出，锚点本轮实测）

- **本批唯一的硬问题**：把 B08b 用 `SemVirtualNotSupported`（VB3027）判死的"类体内调用被后代覆盖的
  成员"翻成真的按实例类型派发。三处拒绝点都在 `semantic_analyzer_expr.cpp`（搜 `vtblNeededMsg`），
  B08d 的正是要**删掉这三处判定**并让发码走槽位 —— 别留着判定再"另外"发虚表，那会变成永远到不了的码。
- **`dynamicKeys` 换成"槽表"**：B08b 存的是**成员名小写**（分析器拿不到属性方向，见 D31-5），发码需要的是
  `ifaceSlotKey` 级的**有序槽清单 + 每槽的实现选择**。所以在 `ClassChainView` 上另加一张
  `std::vector<VirtSlot>{ slotKey, nameKey, const Decl* 声明处, Module* owner }`（根→叶、每槽键一份，
  顺序就是 3.4 的落表序，别另起计数），`dynamicKeys` 保留做分析器侧的快路径（或改读新表）。
- **字段定序一次写死**（D19 偏移 0 + B04 的 `__iv_` + B07b 前缀布局）：
  `__comObj` → `__iv_<I>…` → `__cvtbl` → 祖先字段 → 自有字段。**只给链上任一类带过 `Overridable`/
  `Overrides` 的类加**（照 `__refcount` 的按 feature 加字段手法），否则 8 文件 `--emit-c` 护栏必挂。
  实现位置：`cgen_base_generate_c_open.inc` 的字段循环之前，与 `emitIfaceClassFields` 同一层。
- **类型要跨 TU 可见，表实例不用**（本条是读完 `cgen_inherit.cpp` 后改正的版本，原来写的
  "表与表项都必须非 static"是错的）：每个具体类只需要**一张**表实例，装它的是自己 TU 里的 `_New`
  → `static const vb6_cvtbl_<D> vb6_cvtbl_<D>_impl` 发在 `<D>.c` 就够；但**类型**
  `vb6_cvtbl_<X>` 与函数指针别名 `vb6_pfn_<X>_<slot>` 必须发在 `<X>.h`，因为祖先体内那句
  `Me.M()` 的发码在祖先 TU 里，它要按**自己那张视图**（`vb6_cvtbl_<祖先>`）去索引派生对象装进去的表。
- **多视图靠"字段类型 `void*` + 用点强转"统一**（这是 B08d 的关键设计，别改成 typed 字段）：
  A 声明 `Overridable` → 它的视图是 `{speak}`；B 又新声明一个 `Overridable` → B 的视图是 `{speak, extra}`。
  前缀布局要求 B 的 `__cvtbl` 与 A 的是**同一个字段**（同偏移），类型却不同 → 字段只能发成
  `void* __cvtbl;`，用点写成 `((const vb6_cvtbl_<本类>*)me->__cvtbl)-><槽字段>((vb6_cls_<本类>*)me, …)`。
  槽序定死为"链上虚槽、根→叶、每槽键一份"，于是**任何祖先视图都是派生表的前缀**，强转才成立。
- **装载点**：`vb6_cls_<D>_New`（`cgen_com.cpp` 的 `_New`/`_Destroy` 已经在 B07b 走过同一张
  `structFieldDecls`）→ `me->__cvtbl = (void*)&vb6_cvtbl_<D>_impl;`。**每个具体类只装自己那一张表**
  （同一祖先链、不同覆盖集 = 不同表实例），类型用本类视图 `vb6_cvtbl_<D>`、实例名 `vb6_cvtbl_<D>_impl`
  —— 定死这一对命名，`.h` 的类型、`.c` 的实例与 `_New` 的装载三处必须一致。
- **`Overrides` 的契约检查一行都不用改**（2.8 `runVirtualContractChecks`），它只保证"目标存在且可覆盖、
  签名一致"，与派发机制无关；批完后把 `dynamicKeys` 的**语义**从"要拒绝"改成"要发槽"即可。
- **用例翻转（这是本批的验收证据）**：`ci_n15_base.cls` 的 `Talk = Me.Pick()` 形状搬进
  `tests/cls_inh/InhBase.cls`，断言改成 `d.Talk() = "derived"`、`b.Talk() = "base"`、
  `m.Talk() = "mid"`（若 `InhMid` 也覆盖）；三级链里"中间类覆盖 + 叶类不覆盖"必须单独有一条断言，
  它才是"祖先体内调用绑到最近覆盖者"的证据。`ci_n15_*` 从负例清单里删掉（别留成 skip）。
- **与既有边界的冲突**：Pass D 仍判死"任一侧实现新式 Interface"与"基类有重载" → v1 的虚表只在纯具体类链上
  工作；`__iv_<I>` 与 `__cvtbl` 共存要等 P6 一起设计（`Inherits` 一个实现了接口的类目前根本进不来）。
  `MyBase.M`（B09）的**去虚化**正好依赖本批的表结构（`MyBase.M` = 直调 `vb6_<Base>_M`，不查表），
  所以表项别顺手做成"只能查表"的形态。
- **要复用的取名机器（`cgen_inherit.cpp` 实测，别另写一套）**：C 符号名 =
  `cProcName(procBaseName(decl), accessOf(decl), moduleName)`，属性名自带 `prop_get_/prop_let_/prop_set_` 前缀
  （与 `makePropertySignature`/`resolveClassMemberCall` 同源）；形参表 = `classMeParam()` + 逐个
  `makeParamCType(p, false)` + Optional 的 `int _has_<x>` 尾参；返回类型 = `inheritedRetType(decl)`。
  B08d 的函数指针别名与表项**必须**用这四件套拼，否则要么签名不一致、要么链接期找不到定义。
  填表时把派生实现强转成祖先视图的函数指针类型（前缀布局保证 ABI 一致），
  或者复用 B07b 已经发出来的转发桩做 entry —— 后者更稳，桩的签名天生就是"派生类的 me + 祖先的实现"。
- **护栏**：改结构体 + 改派发 → 8 文件 `--emit-c` 必须全同（基线 = pre-B08d worktree，自建 Debug exe）。
  本轮 B08b 用的 `base_build.ps1` + `byteguard_b08b.py` 已在 `.build\` 里，改两个路径就能复用
  （dev.ps1 在 worktree 里跑不通：它只 `cmake --build`、不 configure，见 D31-10）。

### D33 B08d（类虚表 `vb6_cvtbl_<Cls>` + 运行期动态派发）实施记录（2026-09-23）

1. **槽表落在 3.4b 而不是 2.8**（`Driver::buildVirtualSlotTables`，紧跟 `mergeInheritedMembers`）：
   判"这个槽在**本类面上**有没有可调用入口"要读 3.4 回填的 `inhProcs`（转发桩清单），在 2.8 判会漏掉
   遮蔽。诊断因此也是语义层的（`--syntax-only` 断言得到，`ci_n19`/`ci_n20` 就是这么测的）。
2. **筛选集取链根的 `dynamicKeys`，不是本类的**（D32② 的修正）：`virtSlots(X)` = 链根→X 之间
   **首次声明**为 `Overridable` 的槽 ∩ `dynamicKeys(chain[0])`。取根那份是关键 —— 叶类自己的
   `dynamicKeys` 恒为空（它没有后代），但**它必须带 `__cvtbl` 字段**，否则经基类型变量调 `x.Pick()`
   时基类视图会去读一个不存在的字段（= 别人第一个数据成员的地址，野指针）。前缀性质由"首次声明位置"
   定序保证，而链根集合对链上所有类相同 → 任何祖先的表都是后代表的前缀。另外**链根自己也要建表**：
   它的 `chain` 只有一个元素，照 `chain.size() < 2` 早退就会漏掉根（第一轮实测 `b.Speak()` 静默绑回
   基类实现，就是这个原因）。
3. **表类型只需本类 TU 可见，但必须发在跨模块 include 之后**（D32③ 的修正）：调用点用的永远是
   *接收者静态类型自己那张视图*，所以 `vb6_cvtbl_X` 发在 `X.h` 就够 —— 不过 `InhMain.c` 这类消费者
   TU 会 include `InhDerived.h`，视图类型必须在那批 include **之后**才落地，否则 C2440/C2100
   （`(const vb6_cvtbl_InhDerived)` 被当成非类型名）。定在类 struct 定义之后、`emitInterfaceVtable`
   之前；表实例 `extern const` 同处声明，定义在 `.c` 末尾。
4. **字段类型 `const void*`**（D32② 写的是 `void*`）：装载点写 `me->__cvtbl = &vb6_cvtbl_X_impl;`
   不带强转 → 不会出 discard-const 警告。表项的 me 形参一律用**本类**的 `vb6_cls_<X>*`（祖先视图与
   派生视图各自自足，两个视图之间从不转换指针）→ 填表、调用点都不需要强转，
   `vb6_cvtbl_A*`↔`vb6_cvtbl_B*` 这种不兼容转换在整个工程里根本不出现。
5. **表项 = 本类面上的入口声明**（`cvtblImplName` = `cProcName(procBaseName(impl), accessOf(impl),
   本类模块名)`）：impl 要么是本类自己的声明，要么是 B07b 判定要发桩的那份祖先声明 —— 两者的 C 名与
   签名由同一套取名机器拼，所以填表零强转、链接期不可能找不到定义。Private 过程在 `.c` 里是 `static`
   且**没有前置原型** → 表实例必须后置到 epilogue（与 `emitDelegateThunks()` 同一层）。
6. **`Me.X` 的发码不在"优先级2"那一支**：那一支（`cgen_expr_member_class_module.inc`）整块在
   `node.object` 是 `IdentifierExpr` 的分支里，而 `Me` 是 **MeExpr** → 真正发码的是
   `cgen_expr_member_class_fallback.inc`（它按 emit 出来的对象文本 `"me"` 查 `knownClassVars_`）。
   只改前者会得到"外部落点对了、类体内仍直调"的**假成功**（本轮第一轮就是这个形状：INH18/19/20/22
   全 FAIL 且全部返回 base，而 `InhMain.bas` 里的 `b.Speak()`/`d.Speak()` 已经是对的 —— 这种"半对"
   最容易误判为通过）。两处都得改写，且 fallback 那处必须派发。
7. **接收者必须是纯读表达式**（间接调用要把它的文本用两次：取表一次、me 实参一次）。判据取"文本里
   不存在紧跟标识符/`)`/`]` 的 `(`" —— 强转的 `(` 前面是运算符，所以 `(*c)`、`((vb6_cls_X*)b)`、
   `me->m_oSocket` 都算纯读；`vb6_cls_X_Default()`、`vb6_X_prop_get_pvSocket(me)` 不算。带调用的接收者
   里**默认实例那一路** `mustDispatch=false` 保持直调（它的动态类型恒等于静态类型，直调本来就正确）；
   其余用点**报错**，不退成静默直调（D27-13 判例）。
8. **两条新的 v1 判死**（都用 `SemVirtualNotSupported`＝VB3027，负例可断言）：`Overridable`/`Overrides`
   落在 `Property Let`/`Property Set` 上（写上下文发码在 `tryRewriteCOMLvalue` 那条路，本批不碰）；
   某槽在**本类面上没有入口**（祖先那份是 Private，或被另一个方向的同名成员遮蔽 → 3.4 不发桩），
   后者若放过就是表项指向一个不存在的函数。裸名（`Talk = Pick()`）仍按 B08b 的形状判死：那条路吃的是
   模块作用域的过程符号，与 B07b"继承成员必须写 `Me.`"同一条边界，文案已改成陈述这件事而非"虚表还没发"。
9. **验收形状**：`Inh.vbp` 的运行期断言 17→27 条（新增 INH17..INH26）。INH19（`d.PickThru()="mid"`：中间类覆盖、
   叶类不覆盖）是"绑到最近覆盖者"的直接证据；INH23（`Dim up As InhBase: Set up = d: up.Speak()`）是
   "基类型变量持有派生实例不再切片"的证据 —— 这条在 B07b/B08b 一直是**静默错**的，本批顺带修掉。
   INH24..INH26 是**扇出**（`InhSib` 与 `InhDerived` 同以一个基类分叉）：先探针验证过前缀性质
   （`base/sib2` 与 `base/dark2` 互不污染）才升格成常驻用例。`ci_n15_*` 从负例清单删除并删文件
   （形状升格为正例），新增 `ci_n19`（Let/Set）、`ci_n20`（裸名）。
10. **本批没做**：`With Me` 块内的接收者、属性返回对象作接收者（`pvSocket.Pick()` 那条 083c 通路）的
    派发 —— 两者都还在别的发码路上；家族外访问 `Protected` 的拒绝（B08c，诊断 3023 仍预留）。

### D34 B08c（家族外访问 `Protected` 的拒绝）实施记录（2026-09-23 手工续跑轮）

1. **判定落点**：`SemanticAnalyzer::visit(MemberAccessExpr)` 是 `obj.<成员>` 的**唯一必经点** —— 读、
   写（`AssignmentStmt`/`LetStmt`/`SetStmt` 都 `analyzeExpr(*node.target)`）、`CallStmt` 与
   `IndexOrCallExpr` 的 callee 全从这一处过，所以一个 `checkProtectedVisibility()` 调用就覆盖全部
   语句形状（探针实测 7 种接收者形状：模块级字段 / 局部 `Dim` / `ByVal` 形参 × 字段写 / `Sub` 带参
   调用 / `Function` 读 / `Property Get` 读，全中）。诊断 `SemProtectedOutsideFamily = 3023`
   （`diagnostics.hpp` 的预留兑现），文案 ASCII。
2. **接收者→类只能新加字段**：`Symbol::variableTypeName` 只在 `registerVariable`（模块级字段）里填，
   局部 `Dim w As C` **只在类型是委托时**才填、`ByVal w As C` 参数根本不填 → 第一版判定对局部变量和
   形参静默不响（探针实测）。为什么不把 `variableTypeName` 补满：它被后端十余处按"非空即类实例"
   消费（`inferClassTypeOfExpr` 的 084g 分支、`cgen_with`、`comwrite`、`dllentry_collect`…），补上就是
   **改发码**，直接破本批"发码零改动 + 8 文件全同"的验收口径 → 新增 `Symbol::srcTypeName`（四类登记点
   各留一份 `As <类型>` 原文：`registerVariable` + `visit(LocalDeclStmt)` + Sub/Function/Property 三处参数
   注册），目前全仓只有 B08c 一个读者。
3. **注册表按 `Protected` 也建**（本轮最大的设计修正）：开工地图假设"越权判定读 `ClassChainRegistry`
   就够了"，实测**只有 `Protected`、一条 `Inherits` 都没有**的合法工程会在 2.8 开头早退、`classes_`
   全空 → 判定整个静默失效。早退条件扩成 `anyClause || anyVirtual || anyProtected`
   （`moduleHasProtectedMember`），并让"有 Protected 无 Inherits"也走 Pass A/C。这类视图全是**单元素链**
   → 后端 `classChainOf()` 的 `chain.size() < 2` 守卫照旧返回 nullptr、3.4 合并与 3.4b 槽表在同一守卫上
   空转，所以这条只喂语义层，不给存量工程开任何新发码路（护栏实测 8/8 全同）。
   D30-⑦ 那条"不为 E0 建表"的理由**仍然成立**，别混淆：E0 只看 `modules_`，本来就不需要表。
4. **裁决顺序**：② 沿接收者链**叶优先**找最近声明者（与 3.4 的遮蔽裁决同向）—— 抢到键但不是
   Protected 就放过；③ 再看当前模块的链里有没有那个声明者（声明者=本类或本类祖先 → 家族内）。
   当前模块是类模块但**没登记**（泛型模板 / 接口宿主）→ 放过：证明不了越权就不报。
   标准模块 / 窗体不是任何类的家族 → 直接判。**漏报优先于误报**是这批的取舍（与 B08b 的"宁多拒"
   相反，因为这里多拒=把能编译的存量代码判死）。
5. **成员口径**：字段 / `Sub` / `Function` / `Property` 四类，与 driver 侧 `memberAccess()` 严格一致 ——
   两边认同同样的成员才不会出"表没建→判定静默失效"的缝（`Event` 因此不在判定内：带 Event 的基类
   在 2.8 已被判死，且 `memberAccess()` 不认它）。属性按**名字**取严（`MemberAccessExpr` 上拿不到
   Get/Let/Set 方向，与 B08b `dynamicKeys` 同一个限制）：任一方向 `Protected` 即按 `Protected` 论。
6. **不判的形状**（都写进手册，别当已完成）：数组元素 `w(1).X`、属性/函数返回对象 `pvSocket.X`、
   `With w` 内的 `.X`、嵌套接收者。CLR 那条"必须经 `Me` 或本类型更深实例访问"也没做 —— 家族内的
   基类型变量 `o.m_secret` 允许，正例 `ci_pos2_*` 就是钉这个决定的常驻证据。
7. **用例**：`-Category syntax` 74→78。负例三条各钉一种语句形状 —— `ci_n21`（家族外类里字段写）、
   `ci_n22`（标准模块里 `Protected Sub` 带实参调用）、`ci_n23`（家族外类里 `Property Get` 读）；
   正例 `ci_pos2_base/derived`（家族内经基类型变量与 `Me.` 访问，必须静默）。发码零改动，
   `Inh.vbp` 的 INH12（家族内 `Me.`）与 26 条运行期断言全不变。
8. **本轮环境事故（值得下一轮记住）**：全量门**连跑三次**才拿到 —— 前两次不是代码问题：
   (a) 20:50 另一个写入者把整个工作树复制到 `C:\Users\Administrator\Documents\c3.vb6.pro` 并从那份副本
   跑 `-Category all`，其构建在 20:50:23 重链了我 `.build\C3.exe`（对象文件全在，`[1/1] Linking` 即复原），
   我的门跑到 `test_generics_x86` 起连续 `FAIL (compile)` 后 powershell 以 exit=127 死掉；更早一次
   `.build\C3.exe` 直接**消失**（`FileNotFoundError`），当时 `tasklist` 里有 3 个 `cl.exe`。
   (b) 教训：**跑长门前先 `who_is_building.ps1` 看清有没有别的构建/套件在跑，跑完立刻核 exe md5**；
   门的 `Results` 只在"跑前后 md5 一致"时才可信。工具脚本 `.build/who_is_building.ps1` /
   `who_is_building2.ps1`（列 cmake/ninja/cl + 反查父进程与命令行）本轮留下复用。

### D35 B08e（把剩下的类成员发码路逐条接上派发或判死）实施记录 + 13 站地图（2026-09-24）

**本轮交付 = B08e-1：`With w` 块内的类成员调用接上虚表派发**（`cgen_expr_with.cpp:82` 那一站）。

1. **为什么这一站是真漏洞而不是收尾**：`With up As InhBase`（`up` 持 `InhDerived` 实例）里写
   `.Speak()`，改动前静默绑到 `vb6_InhBase_Speak` → 运行期答 `"base"`。这就是 B08d 用 INH23 修掉的
   同一个切片，只是换了一种写法，而 B08d 的地图（D32）没把 With 站列进必修两处里。
   **A/B 负控实测**：基线二进制（worktree `D:\.wt_b08e_base` @7f34a91，exe md5 `bc3eaa3b`）跑
   `Inh.vbp` → `INH27:FAIL base` / `INH28:FAIL hi bob (base)`；本批改后二进制（exe md5 `00a3a9b4`）
   → 两条 OK。INH29/30/31 两侧都绿，是"不许改坏"的对照：`With` 内**非槽**成员（`Property Let/Get`
   同名对、无括号裸 `Sub`）必须仍走原直调路。
2. **改法**：与 B08d 两处同口径 —— `resolveClassMemberCall` 的结果先留成 `directFnW`，再问
   `virtDispatchCallee(info.className, node.memberName, tempVar, /*mustDispatch=*/true, node.loc)`，
   认得出槽就换 callee，认不出（该类无表 / 该成员不是槽）保持原样。`With` 的接收者是块入口那个
   类实例 temp（`cgen_with.cpp` 声明成 `vb6_cls_<X>*` 的纯变量名）→ `cvtblObjIsPure` 恒真，
   所以 `mustDispatch=true` 不可能因"形状不支持"误报，纯粹是"不留静默直调"的表态。
   **两个方向坑**：① 属性写（`prop_let_`/`prop_set_`）不改写 —— 3.4b 根本不给 Let/Set 建槽
   （D33-7），拿 `node.memberName` 去查会命中同名的 `get_<X>` 槽，把写变成读；② `asCallCallee_`
   那一支仍走 `pendingChainObj_` 协议（Fix 090s），只是 `lastExpr_` 从裸函数名换成派发表达式，
   外层 `cgen_expr_call_com_bind.inc:343` / `cgen_call.cpp:331` 把它当 callee 拼参数 —— 派发表达式
   以 `->slot` 结尾不带右括号，不会误触发 Fix 083e 的"callee 已是 func(obj) 形式 → 拆开"路径。
3. **护栏**：8 文件 `--emit-c` 对基线 exe（7f34a91）逐字节全同 8/8（`.build/byteguard_b08e.py`，
   BASE/NEW 两个绝对路径改一下即可复用到下一批）。用例：`Inh.vbp` 运行期断言 26→31 条
   （`INH27..INH31`），`-Category syntax` 条数不变。
4. **本轮顺手做掉的可达性探针（下一轮别重做）**：`.build/probe_b08e/`（PB.vbp + PBBase/PBDerived/
   PBHolder，全在 `.build` 里、不入版本库，改完即弃）用**当前**二进制跑出的三条形状：

   | 形状 | 运行结果 | 发射出来的 C |
   |---|---|---|
   | `m_h.Speak()`（模块级字段，裸名接收者） | `derived` ✅ | `((const vb6_cvtbl_PBBase*)((vb6_cls_PBBase*)me->m_h)->__cvtbl)->speak(me->m_h)` |
   | `Me.m_h.Speak()`（同一个对象、同一个成员，多个 `Me.`） | **`base` ❌** | `vb6_PBBase_Speak((void*)me->m_h)` |
   | `With w : .Speak()` | `derived` ✅（本轮修的这条） | 派发表达式 |
   | `arr(1).Speak()`（类数组元素） | 编译不过 | `error C2224`（不是静默错码，是硬失败） |

   → **⑤ 站已被证实"可达且正在产出错代码"**，且接收者 `(void*)me->m_h` 是纯读、一行同口径改写即可
   （`thisArg` 与派发用的对象文本取**同一个串**，别拼两遍）；手册里"字段链 `me.m_oSocket.Pick()`
   正常派发"那句是**错的**，本轮已按此证据改正。②站（裸名模块级字段）证实**已在派发**，不用动。

5. **B08e-2（⑤ 站）实施记录**：`cgen_expr_member_voidptr_com.inc` 的 Fix 088b 分支里，
   `resolveClassMemberCall(fieldCls, …)` 的结果在 `thisArg` 算好之后交给
   `virtDispatchCallee(fieldCls, member, thisArg, /*mustDispatch=*/true, node.loc)`；
   `resolvedFn` 含 `_prop_let_`/`_prop_set_` 时跳过（3.4b 无该方向的槽，按成员名查会命中同名的
   `get_` 槽、把读表式子套到写上）。派发串与 this 实参共用 `thisArg` 一个串。
   **用例**：新类 `tests/cls_inh/InhHolder.cls`（`Private m_up As InhBase` + `Hold` + 三个成员），
   `Inh.vbp` 登记，断言 31→34 条。`INH32`（裸 `m_up.Speak()`）两侧都绿 = 对照；
   `INH33/INH34`（`Me.m_up.Speak()` / `Me.m_up.Greet(who)`）在**修复前二进制**
   （`.build/pre_b08e2_C3.exe`，md5 `00a3a9b4` = B08e-1 收线那颗）上实测
   `INH33:FAIL base` / `INH34:FAIL hi bob (base)`，修复后 OK；发射出来的 C：
   `((const vb6_cvtbl_InhBase*)((vb6_cls_InhBase*)(void*)me->m_up)->__cvtbl)->speak((void*)me->m_up)`。
   **顺带把"纯接收者"边界的第二条钉住了**：`pendingChainObj_`/`lastExpr_` 那套协议里换成派发表达式
   仍然正确（派发表达式以 `->slot` 结尾、不带右括号，不会误触发 Fix 083e 的"callee 已是 func(obj)"拆解）。
   **本批未做的半条**：`Me.m_up.Level`（属性**读**）实测不走 ⑤ 而走 ⑩，仍在静默切片 ——
   试写的 `INH35` 因此在修复后仍拿基类实现的 `5`，已连同 `Level` 属性对一起从用例里撤回（不把错行为
   钉成正例），证据与现成用例代码记在 ⑩ 行，下一轮照抄即可。
6. **B08e-3（本轮，零代码交付）= 把 ⑩/⑪ 两行的裁决钉死**，过程本身就是下一轮的施工图：
   - 先按上一版 ⑩ 行的判断在 `class_fallback.inc:224` 加派发 → `INH35` 仍拿 `5`（发射出来的 C 一字未变），   说明属性读**不走 ⑩**。把派发挪到 `m22_module.inc:120`（⑪）后 `InhHolder.c` 立刻炸出   `error C2039: "prop_let_level": 不是 "vb6_cvtbl_InhBase" 的成员` —— 落点找对了，但撞上了写路径。
   - **站点⑭（新增，是 ⑩/⑪ 的先决条件，不在这 13 个消费点里）**：`src/backend/cgen_util_comwrite.cpp` 的 Pattern C/D2（:216 起、:462 落文案   `/* Property Let via prop_get_ rewrite */`）处理 `obj.Prop = v` 的方式是**字符串级改写**——先按读上下文   发一遍 `vb6_<Cls>_prop_get_<P>(<接收者>)`，再把文本里的 `prop_get_` 换成 `prop_let_`/`prop_set_`。   callee 一旦是 `((const vb6_cvtbl_X*)(...)->__cvtbl)->prop_get_p` 这种派发表达式，盲换后缀就得到一张   **不存在的槽字段**（3.4b 明确不给 Let/Set 建槽 → `prop_let_level` 在表里没有）。
   - 三条可选出路（按破坏性从小到大，下一轮先选一条再动手）：(a) 在 ⑪ 只放行**只读属性**（`Property Get` 无同名 Let/Set 那份）——需要先给分析器加一个"该类面上这个成员有没有写方向"的查询，别用字符串猜；(b) 让 Pattern C/D2 认得派发表达式（识别 `__cvtbl)->` 前缀 → 该类属性**不参与**重写、改走一条真实的`prop_let_` 直调），代价是写方向永远不派发（与 3.4b 的判死一致，说得通）；(c) 把 ⑧⑨⑩⑪⑫⑬ 的"不纯接收者"与这条一起按 `VB3027` 判死，属性读继续切片。本轮**没选任何一条**：三条都要动 `cgen_util_comwrite.cpp` 或分析器查询，超出"一次一个改动"的门内余量。
7. **B08e-4（⑭ + ⑪）实施记录**：选型走的是**比 (a)(b)(c) 更小的一刀** —— 不动 `cgen_util_comwrite.cpp`，
   而是给发码层加一个左值标记：`cgen_state.inc` 新增 `bool suppressVirtDispatch_`，5 个
   `emitExpr(*node.target)`（`cgen_setlet.cpp` 两处、`cgen_assign_com_prop.inc` 一处、
   `cgen_setlet_set_prop.inc` 两处）求值左值期间置 true，①⑤⑪ 三处派发点见到它就直调。
   理由：Pattern C/D2 的字符串级改写**只消费左值文本**，所以只要左值永远不带派发表达式，它就永远看到
   自己认识形状；读上下文不受影响。**这条也解释了 B08e-3 没想到的事实**：裸 `m_up.Level = 5` 一直是对的，
   因为 `cgen_assign_prop_write.inc` 有一条**专门的属性写路径**（`/* Property Let */`），它在左值求值之前就
   把 `obj.Prop = v`（object 必须是 `IdentifierExpr`）直接发成 `vb6_<C>_prop_let_P(obj, v)`；
   带 `Me.` 前缀的写法落不进那条路径的 `IdentifierExpr` 判据，才退到 Pattern C/D2 —— 也就是说
   **⑭ 的触发条件恰好是"左值 + `Me.` 前缀"**，比"属性写"更窄。
   - ⑪ 的接线与 ①⑤ 同口径（`thisArg083d = "(void*)" + objExpr` 派发串与实参共用；`_prop_let_`/
     `_prop_set_` 与 `suppressVirtDispatch_` 双跳过）。发射形状（`InhHolder.ViaLevel`）：
     写 = `vb6_InhBase_prop_let_Level((void*)me->m_up  /* class var .m_up field */, 5);  /* Property Let via prop_get_ rewrite (Pattern C/D2) */`，
     读 = `((const vb6_cvtbl_InhBase*)((vb6_cls_InhBase*)(void*)me->m_up  /* class var .m_up field */)->__cvtbl)->prop_get_level(...)`。
   - 用例：`InhBase` 加 `Protected m_lvl` + `Overridable Property Get Level` + `Property Let Level`，
     `InhDerived` 加 `Overrides Property Get Level`（体内 `Me.m_lvl`），`InhHolder` 加 `ViaLevel()`/
     `ViaLevelBare()`，断言 34→36。**A/B**：基线（`.build/pre_b08e4_C3.exe` = B08e-2 那颗）上
     `INH35:FAIL 5`，本批 OK；`INH36`（裸接收者对照）两侧都绿。护栏 8/8；门 153/0/1/154（跑前后 exe md5 f2547af6）。
   - **仍未闭合的一半**：`Me.<字段>.<属性>` 的**写**永远直调（无槽，符合 3.4b），但若哪天给 Let/Set 建槽，
     `suppressVirtDispatch_` 这条标记要一起放开，否则写方向静默不派发。手册已按本轮事实更新。
   - 为什么本轮不过门：改动全部还原后 `git status --porcelain src tests` 为空 → 树与 8987386（已过门 153/0/1/154）   **逐字节一致**，再跑一次全量门不测任何新东西；只把 `.build/C3.exe` 重建回 HEAD 源码   （增量，`abcac1cd`），并用 8 文件 `--emit-c` 对 `pre_b08e3_C3.exe`（= B08e-2 那颗）全同 8/8 + `Inh.vbp` 跑回 34 条断言全绿，证明还原干净。纯文档批不跑门的先例见 B00。

8. **B08e-5（⑥⑦ 站：默认属性调用式 `m_up(9)`）实施记录 + ④ 判为走不到**：
   - 两处 Pattern L（`cgen_expr_call_callee_ident.inc:144`、`cgen_expr_call_callee_member.inc:224`，
     文本一模一样只差缩进）在 `objExpr` 之后加同口径的
     `virtDispatchCallee(fieldType, "Item", objExpr, /*mustDispatch=*/true, node.loc)`，带
     `suppressVirtDispatch_` 保护；派发串与实参共用同一个 `objExpr`。
   - **可达性是靠两条对照探针定下来的**（兑现 B08e-3 的教训）：同一实例、同一 `Item`，显式
     `m_up.Item(3)` 早已派发（发射 `((const vb6_cvtbl_PBBase*)(...)->__cvtbl)->prop_get_item`），
     默认式 `m_up(4)` 发的是 `vb6_PBBase_prop_get_Item(me->m_up, vb6_VariantFromValue(4))` →
     运行期答 `base4`。修复后两式一致答 `derived`。用例：`InhBase` 加
     `Public Overridable Property Get Item(ByVal v As Variant) As String`、`InhDerived` 加 `Overrides`、
     `InhHolder` 加 `ItemDefault()`/`ItemBare()`，断言 36→38（`INH37` 判别项、`INH38` 两侧都绿的对照）。
     **坑一条**：`Item` 的索引参数**必须声明成 `Variant`** —— 这一路固定用 `vb6_VariantFromValue` 打包，
     声明成 `Long` 时默认式撞 `error C2440: vb6_VARIANT → int32_t`（既有缺陷，与本批无关，也说明
     Pattern L 的适用范围只有 Variant 索引的默认属性）。
   - **④（UDT 对象字段）判为"派发无事可做"**：`u.h.Speak()` 连编译都过不去（C2039，发的是非法 C
     `u.h.Speak()`）—— UDT 对象字段在**调用位置**没拿到 `/* udt objfield */` marker，落到
     `class_fallback.inc` 末尾 `obj + "." + member` 的兜底。**登记为新站点⑮ = 路由缺陷**
     （不是派发缺陷）：先让 `u.h.M()` 编得过，再谈派发；它不在 `resolveClassMemberCall` 那 13 站里，
     别混进 B08e 的收尾计数。


**13 站地图（下一轮直接照此动工；站号 = `resolveClassMemberCall(` 的其余消费点）**

先记两条**判定前提**（本轮读码核实，省得下一轮重查）：
- 能否有槽 = 该类在不在 `ClassChainRegistry::classes_`：`driver_classchain.cpp:381` 的门槛是
  `mod->isClassModule && classTypeParams.empty() && !isInterfaceModule` → **.frm 永远无槽**、
  **泛型模板类（含其特化扁名）永远无槽**、接口宿主无槽；`.cls`/`.ctl`/`.pag` 都算类模块，
  所以 **UserControl 可以有槽**（B08e 不能把 `.ctl` 当外部形状跳过）。
- **cgen 期诊断 `--syntax-only` 抓不到**：`virtDispatchCallee` 的 `mustDispatch` 报错发生在发码，
  而 `driver_compile.cpp` 的 syntaxOnly 早退在 :426、3.4b 在 :376 → 3.4b 的判死（如 Let/Set 槽）
  能用 `Test-SyntaxFail` 断言，**B08e 要加的 VB3027 判死不能** → 需要先给 `run_tests.ps1` 补一个
  "编译必须失败 + 断言 stderr 文本"的 `Test-VbpFail`/`Test-CompileFail` 助手（`Test-SyntaxFail` 的
  `cmd /c` 合并 stderr 写法可照搬，把 `--syntax-only` 换成真编译）。

| 站 | 文件:行 | 接收者形状 | 裁决 |
|---|---|---|---|
| ① | `cgen_expr_with.cpp:82` | With temp（纯） | ☑ **已接派发**（本轮，见上） |
| ② | `cgen_expr_member_class_module.inc:20` | 只做"有没有这个成员"的探针，命中后落进已接派发的优先级2 | 判**无需改**（探针；上面探针表已证 `m_h.Speak()` 真在派发） |
| ③ | `cgen_expr_call_callee_member.inc:141` | 探针（`resolvedFn` 只测空，非空即交给已接派发的 MAE 路） | 判**无需改**（探针） |
| ④ | `cgen_expr_member_generic_access.inc:136` | `(void*)<UDT 对象字段链>` | **B08e-5 实测：这条现在走不到"静默直调"** —— `u.h.Speak()`（`Public Type … h As PBBase` 的 UDT 对象字段）编译期就炸：`error C2039: "Speak": 不是 "vb6_cls_PBBase" 的成员`，发出来的是 `u.h.Speak()` 这种非法 C（UDT 对象字段在**调用位置**没拿到 `/* udt objfield */` marker，落到 `class_fallback.inc` 末尾`obj + "." + member` 的兜底）。→ 派发在这里**无事可做**；真正的缺陷是**路由**，登记为新站点⑮（D35-8），不在这 13 个消费点里 |
| ⑤ | `cgen_expr_member_voidptr_com.inc:37`（Fix 088b typed 字段链） | `(void*)me->m_h`（纯读） | ☑ **B08e-2 已接派发**（`Me.m_h.Speak()`/带参的 `Greet` 两条断言钉住，见 D35-5）。**但只覆盖到 Sub/Function 的调用形状**：属性读走 ⑩，见该行的订正 |
| ⑥ | `cgen_expr_call_callee_ident.inc:144`（Pattern L，默认属性 `Item`） | `emitExpr(callee)` 出来的字段/变量（纯读） | ☑ **B08e-5 已接派发**（D35-8）。显式 `m_up.Item(9)` 早就对（走 ①/B08d），默认式 `m_up(9)` 此前静默切片 |
| ⑦ | `cgen_expr_call_callee_member.inc:224`（Pattern L 的另一半） | 同 ⑥ | ☑ **B08e-5 与 ⑥ 同批接上**（两处代码文本一模一样、只差缩进） |
| ⑧ | `cgen_expr_member_class_module.inc:140`（Fix 083c 属性返回对象） | `(void*)vb6_<Mod>_prop_get_<P>((void*)me)` → 含调用，**按构造即不纯** | **B08e-6 实测：这一支被优先级2 提前接走，对带槽成员不可达** —— `Own.Speak()`（Own 是本模块返回类对象的属性）在**改动前**就报 `VB3027`，报出它的是 `cgen_expr_member_class_module.inc:49`（优先级2 的 dispatch，`mustDispatch = defaultInstCls.empty()`），因为 Fix 090al 的 `inferClassTypeOfExpr` 能给属性标识符推出类名 → `itClassVar` 非空 → 根本走不到 :134 那个分支。`:140` 要命中必须"AST 推不出而符号推得出"，本轮没能构造出这种形状。→ **不改码**，用 `ci_n26_base/derived` 把这条既有契约钉住（判死口径与 D33-7 一致，只是先例已在） |
| ⑨ | `cgen_expr_member_obj_dispatch.inc:143`（083c/088d 同一族） | 属性返回对象，同 ⑧ | ☑ **B08e-6 接上派发（不是判死）** —— 地图把这一条当"同 ⑧"是**错的**：`Fix 088d` 这一支的接收者是 `(void*)vb6_ret_<函数名>`，即**本函数返回值的裸 C 局部**，纯读、`cvtblObjIsPure` 恒真 → 与 ⑤ 同法一行改写就派发。改前 `vb6_VtPos3Base_Speak((void*)vb6_ret_Maker)`、改后 `((const vb6_cvtbl_VtPos3Base*)((vb6_cls_VtPos3Base*)(void*)vb6_ret_Maker)->__cvtbl)->speak(...)`，运行期 INH39 由 `base` 翻成 `derived`（见 D35-9） |
| ⑩ | `cgen_expr_member_class_fallback.inc:224`（Fix 088c AST 兜底链） | `(void*)<任意已发码头>`，多为链式调用 | **上一版（B08e-2）把这条改成"与 ⑤ 同法可接"，本轮实测证否：`Me.<字段>.<属性>` 的读根本不走这一支**（本轮在此处加了派发，`INH35` 照旧拿基类实现的 `5`；把 ⑤ 的发射形状与这里对比后定位到 ⑪）。→ 裁决退回**判死**（与 ⑧⑨⑪⑫⑬ 同批处理：接收者按构造含调用时出 `VB3027`）。本轮改动已 `git checkout` 还原，`src/`+`tests/` 与 8987386 逐字节一致。**☑ B08e-6 已生效**（`ci_n25`）：`k.Own.Speak()` 改前发 `vb6_Vt25Base_Speak((void*)vb6_Vt25Base_prop_get_Own(k))`、运行期答 `base`，改前 `--emit-c` 退出码 0；改后出 `VB3027`、退出码 1。这一支的接收者是"纯读的已发码头"时（`me->m_x` 形状）会直接派发，所以接的是派发而不是判死 —— 判死只落在含调用的那一半 |

| ⑪(已交付) | `cgen_expr_member_m22_module.inc:120`（Fix 083d 非标识符 object）**← B08e-4 已接派发，见 D35-7**；原文如下 | | `"(void*)" + emitExpr(object)`：`Me.<字段>` 这类字段链发出来**是纯读**（地图上一版"按构造即不纯"说过头了） | **属性读的真正落点，但不能直接接派发** —— 本轮实测：这一支发出来的读文本会被`src/backend/cgen_util_comwrite.cpp` 的 **Pattern C/D2 字符串级重写**复用（`obj.Prop = v` 是先发一遍读、再把文本里的 `prop_get_X` 换成 `prop_let_X`），callee 换成派发表达式后它照样盲换 → `__cvtbl)->prop_get_level` 被改成 `->prop_let_level` → `error C2039: "prop_let_level" 不是 "vb6_cvtbl_InhBase" 的成员`。**先决条件 = 站点⑭**：让 Pattern C/D2 认得派发表达式（见 D35-6），否则这条只能连属性读一起搁着 |

| ⑫ | `cgen_expr_member_voidptr_com.inc:119`（Fix 015 方法链） | 内层调用返回实例 → 不纯 | ☑ **B08e-6 判死已生效**（`ci_n24`）：`k.Chain(2).Speak()` 改前发 `vb6_Vt24Base_Speak(vb6_Vt24Base_Chain(k, 2))` 并答 `base`，改后出 `VB3027`。**顺带一条实测**：内层 `Chain` 只要被 `Overrides` 过就会先派发（形状 `vb6_KBase_Speak(((const vb6_cvtbl_KBase*)...)->chain(k, 2))`），外层照旧切片 —— 也就是这一族的"内层对、外层错"是同一个洞的两个半边，判死把错的那半边变成响亮的那个 |
| ⑬ | `cgen_expr_member_form_builtin.inc:178`（`.ctl` 子控件） | `(vb6_cls_X*)vb6_UC_InstanceOf(<hwnd>)` → 含调用，不纯；而 `.ctl` **可以**有槽（见前提） | **B08e-6 改判：无需改**（原判"判死"作废）。接收者文本确实含调用，但 `vb6_UC_InstanceOf` 发的是 `r->me` —— 宿主创建该控件时登记的那个 `.ctl` 类的实例（`src/rtl/core/vb6forms/uc/uc_host.c:309`，登记点 `vb6_UC_HostCreate(typeName, …)`），**动态类型恒等于静态类型**，直调即正确。这与 D35 里"预声明实例 `X_Default()` 传 `mustDispatch=false`"是同一条理由；在这里放 `mustDispatch=true` 只会给**正确**的代码凭空造出 VB3027。VB6 也没有"把一个 `.ctl` 实例换成更深派生类"的途径（子控件不是运行时可赋值的对象变量） |

**D35-9 B08e-6（⑧⑨⑩⑫⑬ 判死批 + `Test-CompileFail` 前置）实施记录（2026-09-24 06:16–）**

1. **前置助手**：`tests/run_tests.ps1` 新增 `Invoke-CodegenProj` / `Test-CompileFail` / `Test-Compile`
   三个函数（`Test-SyntaxFailMulti` 的 `cmd /c ... 2>&1` 合并写法照搬，把 `--syntax-only` 换成
   **`--emit-c`**）。为什么是 `--emit-c` 而不是真编译：`--emit-c` 跑完前端 + 全部语义阶段 + 发码，
   正好越过 `driver_compile.cpp:426` 那个 syntaxOnly 早退点，又不调 cl.exe/link（负例一条几秒 vs
   真编译几十秒），且实测**不在源码目录落任何文件**（`tests/cls_neg` 跑前跑后 `ls` 全同）。
   判据仍是"退出码非 0 **且** 输出含指定 ASCII 文本"。
2. **开工前的实测把三条裁决改了**（地图上一版按"接收者形状"推，本轮按发射文本核）：
   - **⑧ 不是判死点，是已经死了**：`Own.Speak()`（Own = 本模块返回类对象的属性）用**改动前**的
     二进制就报 `VB3027`，报它的是优先级2 的 `class_module.inc:49`（`mustDispatch = defaultInstCls.empty()`），
     因为 Fix 090al 的 `inferClassTypeOfExpr` 给属性标识符也推得出类名 → `itClassVar` 命中 → 走不到 :134。
     `:140` 要命中需要"AST 推不出而符号推得出"，本轮构造不出这种形状 → **不改码**，`ci_n26` 钉住既有契约。
   - **⑨ 不是判死点，是能接**：`Fix 088d` 的接收者是 `(void*)vb6_ret_<函数名>`，本函数返回值的裸局部 →
     纯读 → 与 ⑤ 同法接派发（这是本轮唯一的行为改进，见 3）。
   - **⑬ 判死是错的**：`vb6_UC_InstanceOf(hwnd)` 发的是 `r->me`（`uc_host.c:309`），即宿主为该控件按
     `typeName` 创建的那个 `.ctl` 实例 → 动态类型恒等于静态类型 → 直调本来就正确；在这里放
     `mustDispatch=true` 等于给正确代码凭空造错。改成**无需改**（理由与 D35 里"预声明实例传 false"同源）。
3. **本轮交付**：⑨ 接派发、⑩⑫ 判死（`mustDispatch=true`，三处都带 `suppressVirtDispatch_` 与
   `_prop_let_`/`_prop_set_` 两道既有保护，派发串与 this 实参共用同一个串）。证据三条并排：
   - ⑨ `ci_pos3`：改前 `vb6_VtPos3Base_Speak((void*)vb6_ret_Maker)` → 改后
     `((const vb6_cvtbl_VtPos3Base*)((vb6_cls_VtPos3Base*)(void*)vb6_ret_Maker)->__cvtbl)->speak((void*)vb6_ret_Maker)`。
     运行期 `Inh.vbp` 断言 38→40：**INH39**（`RefOf.Speak()` 在派生实例上，改前 `FAIL base`、改后 OK）、
     **INH40**（基类实例上的对照，两侧都 OK）。A/B 用 `.build/pre_b08e6_C3.exe`（同一个工程：39 OK/1 FAIL）。
   - ⑩ `ci_n25_base/derived`：改前 `--emit-c` 退出码 **0**、发 `vb6_Vt25Base_Speak((void*)vb6_Vt25Base_prop_get_Own(k))`；
     改后退出码 **1** + `VB3027`。
   - ⑫ `ci_n24_base/derived`：改前退出码 0、发 `vb6_Vt24Base_Speak(vb6_Vt24Base_Chain(k, 2))`；改后 1 + `VB3027`。
   - `ci_n26`（⑧）：改前改后**都是**退出码 1 + `VB3027` —— 它的价值是防回归，不是新证据，别当新证据记。
4. **护栏与影响面**：8 文件 `--emit-c` 对 `pre_b08e6_C3.exe` 逐字节全同 8/8；`-Category syntax`
   本地 78→**82 条全绿**（+3 判死负例 +1 正例）。判死只在"该类有虚槽"时生效，而槽要求链上有
   `Overrides` → 全仓除 `tests/cls_inh`、`tests/cls_neg` 与 `.build/probe_*` 外没有任何工程用
   `Overridable`（`grep -rln Overridable --include=*.cls --include=*.bas --include=*.frm` 只剩探针三个文件），
   所以这三条改动的可达面就是本轮新用例本身。
5. **踩到一条既有缺陷（不在本批范围，登记）**：`com_entry.c` 给派生类的**覆盖**过程发 extern 时，
   用的是基类的 C 类型名（`extern vb6_cls_KBase* vb6_KDerived_Chain(...)`），而 `typedef struct vb6_cls_KBase`
   的前置声明排在它**后面** → C 编译期 `error C2143: 缺少"{"`（x64/Debug 实测，`Chain As KBase` +
   `Overrides Chain` 即触发；未被覆盖的继承形状不触发，因为走的是 `void*` 那条）。这属于 P6/B13 的
   "从基类继承的成员的 COM 对外暴露"一片，**未修**，本批探针改成不覆盖返回对象的方法绕开。
6. **剩余**：⑮（UDT 对象字段的调用位置路由，D35-8 登记）与 ②③④⑬ 的"判为探针/走不到/无需改"
   三条已在表里落定 → B08e 的 13 站到此全部出完，下一批 **B08e-7 = 站点⑮**，之后 **B09 = `MyBase`**。

建议次序：**④⑤（纯形状、真收益）→ ⑥⑦（默认属性）→ ⑧⑨⑩⑪⑫⑬（判死，一批一次做完，
共用同一个 `Test-CompileFail` 助手与同一条负例族）**。每站都要"最小用例先证可达"（④⑤⑥⑦ 若编不出
"静默直调"的用例就在总表记不可达并跳过，别凑数改码），⑧⑨⑩⑪⑫⑬ 的负例断言的是"必须报 VB3027"，
本身即证据。全部站号出完再开 B09（`MyBase`），两套口径不能并存。

### D36 站点⑮ 勘察（2026-09-24 07:16–07:45，本轮零代码改动）

> **本节的落点在 D37 被订正**：⑮a/⑮b 不是 cgen 的两处独立缺陷，而是同一个"UDT 成员的对象类型名在语义层就丢了"的两个症状；坏的范围只有"UDT 与类跨模块"。**改法看 D37，别照本节末尾的落点清单动手。**

**结论先说**：⑮ 不是"派发漏接的一站"，而是 **UDT 里放工程类对象字段（`Public Type T … h As <类>`）这条
通路本身不通**。派发是它的第三步，前两步不通就永远没有可运行期观测的第三条。因此 ⑮ 从 B08e
的收尾里**摘出来单列**，B08e 的 13 站在 D35-9 就已经全部出完了。

**实测证据（`.build/probe_b08e7/`，二进制 = `.build/pre_b08e7_C3.exe`，md5 `90907230` = 40eea3f 的产物）**：

| 形状 | `--emit-c` 发出来的 C | 真编译结果 |
|---|---|---|
| `Set u.h = d`（u 是 `Dim … As TWrap`，`h As U7Base`） | `u.h = vb6_VariantFromValue(d);  /* Set */` | **error C2440**：无法从 `vb6_VARIANT` 转 `vb6_cls_U7Base*`（结构体字段本来就是 `vb6_cls_U7Base* h;`） |
| `u.h.Speak()`（无参调用） | `… vb6_BSTR_Concat(L"U1:", u.h.Speak())` | **error C2039**："Speak" 不是 `vb6_cls_U7Base` 的成员（+ 级联 C2198） |
| `p.h.Tag(3)`（带参调用，p 是 UDT **形参**） | `vb6_ComCall(vb6_ComGetObjectProp(p, L"h"), L"Tag", …)` | **编得过、走的是外部 COM 晚绑定** —— 工程类实例不是 IDispatch，运行期必崩（D20 族的老坑） |
| `Set w = u.h` 之后 `w.Speak()`（同一个对象，先落到变量） | `((const vb6_cvtbl_U7Base*)((vb6_cls_U7Base*)w)->__cvtbl)->speak(w)` | ✅ 正确且已按实例派发 |
| `With p : .h.Tag(4)` | 同第三行的 COM 晚绑定 | 同第三行 |

**关键一条.** `--emit-c` 对第一、二行**返回 0、不报任何诊断** —— 这两个缺陷是 C 编译期才炸的，
所以 ⑮ 的负例不能走上一批刚铺好的 `Test-CompileFail`（它跑的是 `--emit-c`），必须走真编译
（`-Category compile`/`vbp`）或者干脆用正例（修好之后必须编得过 + 运行期断言）。别把 ⑮ 的证据
记成"`--emit-c` 退出码 0 = 没问题"。

**三条子缺陷与落点（下一批按这个次序做，每条自带验收）**：
- **⑮a 写方向**：`Set <UDT 变量>.<对象字段> = <对象>` 被当 Variant 字段发。包装发生在
  `src/backend/detail/stmt/cgen_setlet_set_rhs.inc`（:31/:61 那两条注释就是它的动机：`With .SourceFile = obj`
  需要 `void*→vb6_VARIANT`），但它没区分"目标其实是 `vb6_cls_X*` 结构体字段" → 判据现成：
  `udtFieldObjCType(udtCType, 字段)`（`src/backend/cgen_util_classtype.cpp:338`）返回 `vb6_cls_*` 时
  **不包装**、直发 `target = value`。验收 = `Set u.h = d` 编得过 + `Set w = u.h : w.Speak()` 答 `derived`。
- **⑮b 调用方向**：`u.h.M()` 没走"标记 → 类方法分发"那条既有通路。机制本来齐了
  （`appendUdtObjFieldMarker` 在 `cgen_util_classtype.cpp:442-448` 打标记，
  `cgen_expr_member_generic_access.inc:125-150` 消费标记），但 UDT 字段访问**在调用位置**这一支
  发的文本没有标记（无参 → `class_fallback` 末尾 `obj + "." + member` 兜底；带参 → 落 COM 晚绑定）。
  验收 = `u.h.Tag(3)` 发 `vb6_U7Base_Tag((void*)u.h, 3)` 且**不走** `vb6_ComCall`。
- **⑮c 派发**：⑮a/⑮b 通了之后，`thisArg = "(void*)u.h"` 是**纯读**（不含调用括号）→ 与 ⑤/⑨ 同法一行
  接 `virtDispatchCallee(..., mustDispatch=true, ...)`，`u.h.Speak()` 才按实例绑定。硬约束照旧：
  派发串与 this 实参共用同一个串、Let/Set 方向不碰、左值求值期间 `suppressVirtDispatch_` 优先。

**为什么本轮不动手**：⑮a 改的是 `Set` 发码主干（VBMAN 时代的 `With .X = obj` 形状全靠它），
⑮b 改的是成员访问的分支归属（`class_fallback` 与 generic_access 谁先命中），两条都是**存量兼容面
很大**的通路，与本线的"虚表派发"没有共同判据，合起来远超一个可过门的批次。按"宁少勿滥 + 一次一个
改动"的纪律，本轮只把勘察结果与落点钉下来，代码一行未改（因此也没有门可跑，先例 = B08e-3 零代码不跑门）。

### D37 ⑮ 的真根因：UDT 成员的对象类型名在**语义层**就丢了（2026-09-24 07:31–08:05，B08f-1 尝试轮）

D36 把 ⑮a/⑮b 的落点记在 cgen（`cgen_setlet_set_rhs.inc` 的 Variant 包装、`generic_access` 的标记消费），
本轮照那个落点改了一遍，**实测同一条语句一字不变** → 那条守卫是死代码，已 revert。真正的原因在下面：

1. **两套解析各说各话**。UDT 成员 `h As <项目类>` 走 `resolveTypeRef`（`semantic_analyzer_typeref.cpp:29`），
   它对认不出来的类型名**一律回退 `Vb6Type::Variant`**（同文件 Fix 040a/069/157 那一大段注释就是这条兜底的历史账单）。
   项目类名要走到 `symTab_.lookupModule(...)` 才判成 Object，而**跨模块的 Class 符号在这个模块被分析时还没注入**
   （注入在 stage 3.5）→ 于是 `mi.type = Variant`。
2. **名字也一起丢了**。登记处 `semantic_analyzer_decl_type.cpp:31-40` 只在
   `mi.type == UserDefinedType` 或 `== Object` 两个分支里存 `mi.typeRefName = stRef->name`；
   Variant 分支什么都不存 → **类名在 `Symbol::UdtMemberInfo` 里根本没有**，后面谁也补不回来。
3. **cgen 却看得懂同一个字段**。结构体发射器（`cgen_decl_type.cpp:74-80` 那条回退）跑在 stage 3.5 **之后**，
   `lookupModule(类名)->kind == Class` 成立 → 发出来的字段是 `vb6_cls_U7Base* h;`。
   于是同一个字段：**布局侧知道它是项目类，成员元数据侧只知道它是 Variant**。
4. **所有下游都挂在元数据上**。`udtFieldObjCType` 只在 `mi.type == Object` 时才认对象字段
   （`vb6_cls_X*` / `void*` 两分支），Variant 成员一律返回 "" → `appendUdtObjFieldMarker` 不打标记 →
   调用位置落 `class_fallback` 末尾的 `obj + "." + member`（C2039）或被 COM 晚绑定抢走（`vb6_ComCall`）；
   而 Set 侧 `inferUdtFieldVb6Type(...) == Variant` 为真 → 包 `vb6_VariantFromValue`（C2440）。
   **⑮a 与⑮b 是同一个缺口的两个症状，不是两条独立缺陷**（D36 的三分法在这里要收拢）。
5. **判别实验（决定性的一条）**：把同一个 UDT 声明在**类模块内**（`U7Base.cls` 里
   `Private Type TInCls … h As U7Base`），此时该模块自己有 Class 符号 → `mi.type = Object` +
   `typeRefName` 有值 → 实测两条都正确：`Set t.h = me;  /* Set */` 与
   `vb6_U7Base_Speak((void*)t.h)`（仍是直调，派发要等 ⑮c）。
   → **坏的范围只有"UDT 与类跨模块"这一种**，不是整条 UDT 对象字段通路都不通。

**B08f-1 的做法（本轮 08:20 起按此实施）= 两处小改 + 真编译验收**：
- `semantic_analyzer_decl_type.cpp`：SimpleTypeRef 的兜底分支里，**即使解析成 Variant 也把 `stRef->name` 存进
  `mi.typeRefName`**（只多存一个名字，不动 `mi.type`）。**读侧逐条核过（本轮实测 grep，共 9 处）**：
  `cgen_util_classtype.cpp:351`/`:354` 在 `mi.type == UserDefinedType` / `== Object` 分支内、`:537` 判
  `bt == UserDefinedType`；`cgen_util_type.cpp:686`、`cgen_expr_call_callee_member.inc:86`、
  `cgen_expr_call_callee_withm.inc:82` 三处都是 `mi.type == UserDefinedType && !typeRefName.empty()`；
  `cgen_util_classtype.cpp:234-237`/`:273-276` 两处只看"非空"，但拿到名字后还要
  `lookupModule(name)->kind == UserDefinedType` 才认 → **塞进去一个 Class 名，这九处一律看不见**，
  嵌套 UDT 判定也不会因此多认出一个成员。
- `udtFieldObjCType`：加一条 Variant 分支 —— `lookupModule(typeRefName)` 命中 `SymbolKind::Class` 时按
  **类的规范模块名**返回 `vb6_cls_<Module>*`（与结构体发射器同一口径，别用引用名）。
- 验收：`tests/cls_inh` 里加"跨模块 UDT 对象字段"的运行期断言（**必须走真编译**，`--emit-c` 对这两种坏形状返回 0）；
  `Set`/调用两条 + 8 文件逐字节护栏 + A/B。

**B08f-1 实施结果（2026-09-24 09:20，代码 `77ecef1`）**：上面三处接线一起做完，⑮a/⑮b 一次治好，
并且**通路一通，D35 站点④ 才第一次可达** → 顺手把 ⑮c 的派发也接上了（没另开一批：那 8 行不接就是
"编得过、静默绑基类实现"，正是 D27-13 那一类，留着比接上更危险）。实测三条并排：

| 形状 | 改码前（`.build/pre_b08e6_C3.exe`） | 改码后 |
|---|---|---|
| `Set u.h = d` | `u.h = vb6_VariantFromValue(d);` → **C2440** | `u.h  /* udt objfield vb6_cls_InhBase* */ = d;`（合法 C） |
| `u.h.Speak()`（无参） | `u.h.Speak()` → **C2039** | `((const vb6_cvtbl_InhBase*)((vb6_cls_InhBase*)(void*)u.h)->__cvtbl)->speak((void*)u.h)` |
| `u.h.Greet("bob")`（带参） | `vb6_ComCall(vb6_ComGetObjectProp(u, L"h"), L"Greet", …)`（编得过、运行期解 vtable 崩） | 同上形状走 `->greet((void*)u.h, …)` |

验收走真编译：`Inh.vbp` 运行期断言 **40→43**（INH41/INH42 判别、INH43 基类对照），**改码前的二进制
编不过这个工程**（同一份源码：C2440 + C2039）；8 文件 `--emit-c` 逐字节全同 8/8、`-Category syntax` 82/0。

**还开着一层（本轮实测，登记为 ⑮d）**：UDT **本身**声明在"使用它的模块"之外的第三个模块时
（`TWrap` 放 `InhUdt.bas`、`Sub Main` 放 `InhMain.bas`），同样的坏形状**原样复发** —— 这一次卡在
`udtFieldObjCType` 开头那句 `symTab_.lookupModule(udtName)`：消费者模块的符号表里没有那个 UDT 符号
（跟 `mi.type` 无关，是 UDT 名的跨模块可见性）。本轮把测试用例改成"UDT 与使用点同模块"（= 已实测
修好的那一类），**没有**顺手扩大改动面。⑮d 的判据要先量一下：`knownUdtVars_`/`inferUdtTypeOfExpr`
在跨模块 UDT 上的可见面有多窄（`cgen_localdecl.cpp:278` 只登记本模块看到的名字）。

**⑮e（本轮顺手实测，未改，不是本批引入的回归）**：**属性写穿过 UDT 对象字段**仍不通 ——
`u.h.Level = 5` 改前发 `u.h.Level = 5;`、改后发 `u.h->Level = 5;`，**两种都是非法 C**
（`Level` 是属性不是结构体字段，两条路都没接到 `prop_let_`）。要治得走 `cgen_assign_prop_write.inc`
那条专用属性写通路，而它要求 object 是 `IdentifierExpr`（D35-6 记过这条前提）—— 判据与 ⑮a/⑮b 不同，另批。
**同一轮的另一条行为变化**：站点④ 接上派发之后，`Set t.h = me` + `t.h.Speak()`（UDT 声明在类模块内、
改前就编得过的那一类）从 `vb6_U7Base_Speak((void*)t.h)` 变成按 `__cvtbl` 派发 —— 即本批除了修通路，
还把**存量**的"UDT 对象字段 + 覆盖成员"从静默切片改成按实例绑定。这是要的语义，但它作用在已经用这一形
状的老工程上，门里若有钉住旧行为的用例应按新语义改用例（本轮 `bas`/`vbp`/`compile` 全绿，没有这种用例）。

**一条小瑕疵（未改，记着）**：修好之后标记注释会漏进发射语句
（`u.h  /* udt objfield vb6_cls_InhBase* */ = d;`）—— C 语法合法（注释即空白），而且**这是既有通道
本来就有行为**（类模块内声明的 UDT 在改码前就发成这样），故本批不动它；要清就照
`cgen_expr_call_arg_emit.inc:583` 的"发射前剥离标记"办法，在赋值发射点统一剥。

### D38 B09 的实施记录：`MyBase` 去虚化 + 构造链（2026-09-24 08:32–，代码 `02bac92`）

**接管点选在发码层按名字匹配，没有加 token**（与状态头 ⑤ 那条"关键字策略"预告的不同，理由记在这里）：
`MyBase` 现在只是 `MemberAccessExpr(IdentifierExpr("MyBase"), M)`，`cgen_expr_member_precheck.inc`
里 `Err`/`LastError`/`VBA` 那一条**按小写接收者名最早接管**的通路就是为这种伪接收者准备的先例，
在它后面加一支 `if (_objLower == "mybase")` 就够 —— lexer/token 表、parser、软关键字表全不动，
于是"零新语法逐字节不变"这条护栏是**构造上成立**的，而不是靠 Early-return 保证的。
代价：源码里名叫 `MyBase` 的变量会在类模块里被当成关键字（语料核查：`tests/`+`archive/`+`publish/`
的 `.bas/.cls/.frm` 里 `MyBase` 零命中，与 D-关键字策略那条同一份证据）。语义层只补一处：
`mybase` 在类模块里不再报 `VB3001 未声明的标识符`（否则每条合法写法都配一条噪声）。

**取哪一份实现**：`findMyBaseProc(基类视图, 成员名, 读/写)` —— 先查基类**自己的声明**，
没有再沿基类的继承面（`inhProcs`）找，即"就近声明"。这条与 B07b 中转函数用的是同一份 owner，
所以不会出现"桩能连、`MyBase` 连不到"。C 名一律 `cProcName(procBaseName(decl), access, owner 模块名)`
（= 表项与桩的同一套拼名），接收者实参 `((vb6_cls_<owner>*)me)` —— 前缀布局让它必然指向同一偏移。
**不查 `__cvtbl`**：这就是"去虚化"，与 `virtDispatchCallee` 无任何交集，因此 D35 那 13 站的判据、
`mustDispatch`、`suppressVirtDispatch_` 在这里一律不适用。读写分档：读上下文按
`Get > Function > Sub > Let > Set`（与 `resolveClassMemberCall` 一致），写上下文只认 `Let`/`Set`
（拿 `Get` 去写就是 C2198 或值被丢掉）。写侧另有一条入口：`MyBase.X = v` 原来落到
`cgen_assign_stmt_special.inc` 的 `Module.var` 回退，发的是 `vb6_MyBase_X = v`（C2065），
现由 `tryEmitMyBaseAssign` 接管。

**判死（`VB3028`，新增诊断码）**：① 本类没有 `Inherits`；② 基类面上没有这个名字；
③ 目标是基类的 `Private` 成员。③ 不是洁癖：Private 过程在 C 层就是 `static`（Fix 089e），
派生 TU 连不到，放行只会把错误推到链接期。

**构造链**：`New` 派生类时祖先的 `Class_Initialize` 按**根→叶**先跑（`emitClassInitChain` 插在
`emitClassFactory` 里"自家那份初始化"之前，字段默认值已置好）。基类那份是 `static` → 加桥接
`vb6_<基>_chain_init`，**只在**"这个类自己写了 `Class_Initialize` **且**确实被谁继承"时发
（`classIsBaseOfSomething` 扫 `clsreg_` 的 `baseKey`）→ 零继承工程逐字节不变。
`MyBase.Class_Initialize` 复用同一座桥（INH50 钉的就是这条）。`Class_Terminate` 的反序链**没做**。

**B09-0（前置缺陷，D35-9 ⑤ 登记的那条）已修**：`com_entry.c` 的类方法 extern 用**另一个类**的
`vb6_cls_X*` 当返回类型，而本文件的 `typedef struct vb6_cls_X X;` 是按类分块发的 —— 派生类块排在
基类块之前时那一行就是 `error C2143`。改成 `struct vb6_cls_X*`（自带标签，与同一批 extern 里
`_New`/`_Destroy` 的写法同口径），顺序依赖消失。**没修**的另一半：同一个方法在基类自己的 extern 里
发的是 `void*`（`mapType` 把类返回看成 Variant，`variableTypeName` 在自家符号上为空），派生类那份才是
`vb6_cls_X*` —— 指针返回两者 ABI 相同，不崩，但 Fix 184 的意图只做到了派生侧；归 P6/B13 那片复核。

**实测形状**（同一份源码，`pre_b09_C3.exe` → `C3.exe`）：

| VB 写法 | 改码前 | 改码后 |
|---|---|---|
| `MyBase.Speak()`（本类已 `Overrides`） | `vb6_MyBase_Speak()` → C2065 | `vb6_InhBase_Speak(((vb6_cls_InhBase*)me))` |
| `MyBase.Cat("x","y")`（带可选实参） | 同上 | `vb6_InhBase_Cat(((vb6_cls_InhBase*)me), …, 1)` |
| `MyBase.Name = v` / `MyBase.Name` | `vb6_MyBase_Name = v` → C2065 | `vb6_InhBase_prop_let_Name(…)` / `…_prop_get_Name(…)` |
| `MyBase.Class_Initialize` | `vb6_MyBase_Class_Initialize()` → C2065 | `vb6_InhMid_chain_init(((vb6_cls_InhMid*)me))` |
| `New InhDerived` 的初始化 | 只跑自家那份 | `vb6_InhBase_chain_init(me); vb6_InhMid_chain_init(me);` 再跑自家 |

**验收**：`Inh.vbp` 真编译 + 运行期断言 **43→52**（INH44/INH45 判别去虚化 —— 同两个成员经
`Me.`/`obj.` 答案是 "derived"；INH46 就近遮蔽；INH47/INH48 属性写读 + `Public` 字段；
INH49/INH50 构造链与桥；INH51 返回工程类的 `Overrides`（B09-0 的用例）；INH52 中间类实例）。
**改码前的二进制编不过这个工程**：6 个 C 错，其中 `C2143` 正是 B09-0、`C2065` 是 `MyBase` 那几条。
负例 `ci_n27_mybase_no_inherits` / `ci_n28_mybase_no_such_member` 走 `--emit-c`（`Test-CompileFail`
数组现在自带 needle）：改后退出码 1 + `VB3028`，改前退出码 0、静默发未声明符号。
护栏：8 文件 `--emit-c` 对 `pre_b09_C3.exe` 全同 8/8、`-Category syntax` **82→84 全绿**。

### D39 B09 顺手挖出的 x86 布局缺陷：继承字段里的 Private UDT 在派生 TU 退化成 `void*`（2026-09-24 09:55，登记为 **B09b**）

**症状**：`Inh.vbp` 在 **x64 全绿（52 条断言）**，同一份源码 **x86 段错误**。定位过程：把
`INH48`（`MyBase.Label = v` 之后 `MyBase.Label` 读回）拆成"只写"/"只读"两个函数，x86 在
`M48a` 之后、`LabelWrite` 之前崩 —— 也就是**写一个位于 `m_pt` 之后的基类字段**就越界了。

**根因（发射形状，`--emit-c` 一眼可见）**：`InhBase` 里 `Private m_pt As TPoint`（`TPoint` = 两个
`Long`），派生类的结构体把这份**继承来的字段**发成了别的类型：

| 结构体 | `m_pt` 那一行 |
|---|---|
| `vb6_cls_InhBase` | `vb6_type_TPoint m_pt;` （8 字节） |
| `vb6_cls_InhMid` / `vb6_cls_InhDerived` / `vb6_cls_InhSib` | **`void* m_pt;`** （x86 4 字节 / x64 8 字节） |

于是"派生实例的前缀必须与祖先 struct 一字不差"这条 B07b 的核心前提，在**含 UDT 字段**的基类上：
x64 靠 `sizeof(void*) == sizeof(TPoint) == 8` **巧合成立**，x86 上 `void*` 只有 4 字节 —— 派生结构体
比祖先布局**短 4 字节**，`m_pt` 之后的每个字段（`m_name`/`m_lvl`/`Label`/`g_viaRet`/`g_init`）偏移全错，
而且 `_New()` 按 `sizeof(派生)` 分配，基类那份过程按自己的偏移写 → **写出堆外**。

**归属**：**不是 B09 引入的**。`structFieldDecls`（`cgen_inherit.cpp:75`）从 B07b 起就按
`mapTypeRef(var.asType)` 发继承字段类型，而 `TPoint` 是基类的 **Private UDT**，在派生 TU 的符号表里
看不见 → `mapTypeRef` 回退成 `void*`。证据：x86 上**改码前**的源码（`git show 546e7bb` 那一套 +
`pre_b09_C3.exe`）已经在越界写 —— `INH30`（`.Name`）与 `INH39/INH40`（`g_viaRet`）写的都是
`m_pt` 之后的字段，只是那几次的值没触发崩溃（三条 FAIL 而非崩溃）。B09 的 `g_init`/`Label`
把同一个洞撞成了段错误。

**为什么本批不顺手修**：修法是"让派生 TU 拿到祖先 Private UDT 的**定义**并按 `vb6_type_<X>` 发字段"，
需要（a）按**声明所在模块**解类型（与 D37 同一族：消费者模块的可见面），(b）把那份 UDT 定义注入消费者
.c（`VB6_TYPE_<X>_DEFINED` 那个 guard 说明按需注入的先例已有），(c）逐字节护栏要覆盖"字段偏移变化"
这一类**布局**改动 —— 与 B09 的判据（成员解析与派发）不同条线。且它只影响 `Inherits` 这条新线
（全仓除 `tests/cls_inh` 没有别的工程写 `Inherits`），不是存量 VB6 工程的回归。

**B09b 的验收必须包含 x86**：这一族的教训是"x64 巧合等价"掩盖了布局错误。最小用例 =
基类含 `Private <UDT 字段>` + 该字段之后还有别的字段 + 派生类读写它，**x86 真编译 + 真跑**。
`Inherits 语句.md` 的"注意"里那条"基类含 Event / 前缀复制不成立"旁边要补一句本版 x86 的 UDT 字段限制。

### D40 B09b 的实施记录：祖先 `Private` UDT 的**类型名**进派生模块作用域（2026-09-24 10:18–，代码 `debb110`）

**判据先钉住，再动手**（D39 留的三条里第一条就是别照抄 D37 那一版）。做法 = 拿同一个工程改一个词：
`Private Type TPoint` → `Public Type TPoint`，其余一字不动，`--emit-c` 一比：

| 基类里 UDT 的写法 | 派生结构体里那份继承字段 | x86 build+run | x64 build+run |
|---|---|---|---|
| `Private Type TPoint` | `void* m_pt;` | **段错误**（47 行后崩） | 52/52 全绿（假绿） |
| `Public Type TPoint` | `vb6_type_TPoint m_pt;` | 53/53 全绿 | 全绿 |

一张表就把洞定死了：**不缺类型定义、不缺 #include、不缺尺寸，只缺名字**。`visit(TypeDecl)` 发 UDT
定义时不分访问级别（还带 `VB6_TYPE_<X>_DEFINED` 守卫，重复发也安全），定义的文本一直在基类自己的
`.h` 里；而 `mapTypeRef` → `lookupTypeSymbol` 走的是**消费者模块**的符号表，`getPublicSymbols()`
（`driver_crossmod.cpp:33` 那份池子）按 `access != Private` 过滤 → 祖先的私有类型符号从来就没进过
派生模块的作用域 → 解不出就回落到 `void*`。所以这不需要 D39 里设想的“按需把 UDT 定义注进消费者 .h”
那条重活，落点在**注入阶段**，不是发码阶段。

**改了什么**：`runCrossModuleResolution()` 末尾加一条独立 pass（在 O3 重载补跑之前）：对每个
`ClassChainView` 有 `baseKey` 的模块，沿 `chain` 走到每个祖先模块，把它符号表里 `UserDefinedType` /
`EnumType` 中**被继承字段用到**的那些（键 = `inhFields` 里 `VariableDecl::asType` 是 `SimpleTypeRef`
的名字），以 `isExternal + sourceModule=<祖先模块名>` 复制进消费者表，`UserDefinedType` 连
`udtMembers` 一起复制（与既有跨模块 UDT 注入同一条口径）。三道护栏：① 只在 `!classes_.empty()` 时跑；
② 消费者已有同名符号一律不抢（`lookupModule` 命中就跳过）；③ 只沿 `Inherits` 链、只补被字段用到的名字
→ 不写 `Inherits` 的工程一个符号都不会多（8 文件 `--emit-c` 对 `pre_b09_C3.exe` 全同 8/8，且那一版
BASE 是 B09 之前 → 连 B09+B09b 两批一起证了没漂）。

**顺手清掉两条陈旧红字**：`Inh35`/`Inh36`/`Inh39` 从 B07b 起就在 x86 上失败（本批之前实测过：改码前的
源码 + 改码前的二进制，x86 三条 FAIL 且不崩）。它们读的都在 `m_pt` **之后**（`m_lvl` 经 `Level` 属性、
`g_viaRet`）—— 同一个错位洞，不是派发逻辑错。修完 x86 与 x64 各 52/52、0 FAIL。

**门里补的口径（本批真正的长期收益）**：`tests/run_tests.ps1` 把 `Inh.vbp` 的期望清单提成
`$inhExpected` 变量，同一份清单跑 **两遍**：`cls_inh_pair`（默认架构）+ 新增 `cls_inh_x86`
（`-Arch "x86"`）。**布局类改动从此自动双架构覆盖** —— 只测默认架构的门证明不了布局主张（D39 的教训）。
`-Category syntax` 84/0 不变（顺带证明脚本改动能解析）。

**没做的相邻两件事**（别顺手，都要先量）：① 祖先 `Private` UDT 出现在**方法签名**（参数/返回值）上时
仍是同一条回落路径（`inheritedRetType`/`makeParamCType` 也调 `mapTypeRef`）—— 本批只治“按值嵌进结构体”
这一类，签名那类是 ABI 问题、判据不同；② 泛型特化与 `Inherits` 的组合、以及⑮d（UDT 声明在第三个模块）
仍然各自卡在自己的可见性判据上，与本洞不同条线。

### D41 B09c 的三条实测与处置（2026-09-24 10:56–，代码 `9eb2ca7`）

领到的活是"收尾 P3 的三条小尾巴"。**先三条一起量，再决定动哪条** —— 结果只有一条需要改码，
一条本来就修好了，第三条根本不该做：

**① `Set MyBase.<属性> = obj`：确实坏，且只坏在接收者。** 改码前的发射形状（探针实测，非推测）：

```c
vb6_T9Base_prop_set_Peer(MyBase, me);  /* Set Property */
```

函数名那条通路自己查对了（基类的 `prop_set_`），实参顺序也对（this 在前、值在后），
**`MyBase` 被当裸标识符发出去** → C2065。所以修法就是在 `visit(SetStmt&)` 的
Nothing / COM / 链式写分支**之前**接管，复用 Let 侧那条 `tryEmitMyBaseAssign`，加 `forSet=true`：
只认基类面上的 `Property Set`（`procRankForWriteSet`）—— 拿 `Let` 槽接对象引用会静默丢引用，
所以命中 Let 时判死而不是将就。改码后：`vb6_InhBase_prop_set_Peer(((vb6_cls_InhBase*)me), me);`，
`Set MyBase.Peer = Nothing` 同一条路一并成立（值侧就是 NULL）。

**② ⑮e（属性写穿过 UDT 对象字段）：B08f-1 之后其实已经通了，本批只补断言、不改码。**
D37 那条“改后发 `u.h->Level = 5;`，两种都是非法 C”是**在类型名还没修好的元数据上量的**。
现在的实测形状（`h As F9Base`，`u.c.Level = 9`）：

```c
vb6_F9Base_prop_let_Level((void*)u.c, 9);  /* Property Let via prop_get_ rewrite (Pattern C/D2) */
... ((const vb6_cvtbl_F9Base*)((vb6_cls_F9Base*)(void*)u.c)->__cvtbl)->prop_get_level((void*)u.c) ...
```

写绑根的 `prop_let_`（3.4b 不给 Let 建槽，本就应该直调），读按实例派发 —— 运行期 `F9=109`，
**x64 与 x86 都是 109**。仓库里的同形断言 = INH55 / INH59。

**③ `Class_Terminate` 的继承链：不做，因为 EXE 工程里它根本没有触发点。** 探针一次跑三种形状
（`Set d = New X` + `Set d = Nothing`、`Dim d As New X`、过程内局部对象出作用域），
`Class_Terminate` 里的 `Debug.Print` **一条都没出来**。全仓 `_Destroy` 的调用点只有
`cgen_iface_vtbl.cpp:444`（接口引用 Release 归零）与窗体/控件销毁那条。也就是说
“发一条反序的 terminate 链”= 发死代码，而且要验证它得先把生命周期做出来 —— 与
`Inherits` 无关的一个**既有缺口**（普通工程类对象的终止时机），归 P6/生命周期那片另批登记，
不在 B09c 里顺手做。

**两条写错的期望（自己踩的，记下来免得再踩）**：
- INH55 一开始期望 `Level` 经 `InhSib` 的字段读回 54 —— 错在**忘了槽是按分支建的**：`Level` 只被
  `InhDerived` 覆盖，`InhSib` 这条支链上它没有槽，两个方向都该绑根（4 进 4 出）。这正是
  B08d“筛选集合取链根 dynamicKeys”的语义，测试写错了不是编译器错了。
- INH57 期望 `MyBase.Class_Initialize` 重跑后是 `base;sib;base;` —— 错在根类的初始化是
  **赋值** `g_init = "base;"` 不是追加，重跑合法地把叶类那段抹掉，正确答案是 `base;`。
  （判别性没丢：不重跑才是 `base;sib;`。）

**验收**：`Inh.vbp` 断言 **52→59**，**x64 与 x86 各 59/59、0 FAIL**；A/B = `pre_b09c_C3.exe`
编不过这个工程（就那一条 C2065）；8 文件 `--emit-c` 对 `pre_b09c_C3.exe` 全同 8/8；
`-Category syntax` 84/0。INH58 是 B09b 那条线的最深用例（`InhSib` 隔两级持有根的 `Private` UDT
字段 + `SetPt/PtSum` 走真偏移）。

**B09b 的一处加强（同一批用例里顺手拿到）**：`TPoint` 声明在 `InhBase`、被隔两级的 `InhSib` 持有，
而 `InhSib` 自己不加字段 —— 于是“派生链上任一级把 UDT 字段发成 `void*`”在 x86 上会以**破坏**而不是
**恰好重叠**的形式现形（`InhDerived` 那侧因为后面还有自有字段把空洞吃掉了，所以旧用例看不出问题）。

### D42 B10 前置实测（2026-09-24 11:40–，只量不改码；探针在 `.build/probe_b10/`）

**探针形状**：`V10Widget.cls`（一个普通类，成员 `Area`/`Describe`/`Name` 的 Get+Let/`Rename` 全齐）
+ `V10Square.cls`（同文件内声明 `Interface IShapeV10` 与 `Interface INamedV10 Extends IShapeV10`，
再写 `Implements INamedV10 Via m_w` + `Private m_w As V10Widget`）。两文件一起过
`--syntax-only` 与 `--emit-c`（跨模块形状，不建 vbp）。

1. **`Via` 今天连软关键字都不是**（本轮 grep 复核 D1 记的四个登记点：`src/lexer/token.hpp`、
   `src/lexer/token.cpp`、`src/lexer/lexer_keywords.cpp`、`src/parser/parser_helpers.cpp` —— **`Via` 零命中**）。
   实测后果落在**语法层**：`V10Square.cls(19,22): error VB2003: expected end of statement (newline or :)`
   ＋ `(19,26): error VB2002: unexpected token at module level: m_w`，两条路（syntax-only / emit-c）
   退出码都是 1。也就是说模块级 `Implements` 只吃到接口名，`Via` 被当标识符撞在语句结尾上。
   → **B10 的第一层是词法+语法**（四处登记 + `Via <字段>` 尾子句），且**语义层与发码层今天没有观测面**
   （编不过就拿不到发射形状 —— 别一上来就改 `cgen_com.cpp`）。
2. **要免掉的那份手写长什么样，量出来了**：同一份文件去掉 ` Via m_w` → 语义层按 **Extends 展开后的每个槽**
   逐条报 `VB3012 Implements INamedV10: member 'IShapeV10.Area' (Function Area() As double) is not
   implemented by class 'V10Bare'`（本例 5 条：`Area`/`Describe`/`Name` 的 Get/`Name` 的 Let/`Rename`）。
   → Via 的验收面 = **这 5 条 VB3012 全部消失**、且**不写任何一个转发成员**就能按 `IShapeV10`/`INamedV10`
   槽位调通运行期。
3. **顺序约束（先想清楚再动手）**：契约检查在语义层（stage 3，跑在 codegen 之前），转发桩是发码层产物，
   所以"已实现"的判定**必须走符号面**（查 `V10Widget` 的成员表，D2 预留的 `implementsMap`/$itf$ 键正是为此），
   不能等桩发出来再验 —— 反过来桩那侧要能拿到"槽 → 被委托成员"的同一份映射，两边共用一个来源，别各算一遍。
4. **本批边界**（越界就是范围红线）：`Via` 目标只认**本类的对象持有字段**；VB6 老形式
   `Implements x ByRef y As New T`、属性 `Set` 向的委托、`Extends` 深度 >2 的链、以及**B06c**（接口值作实参）
   都不在 B10 —— 先把一条 `Interface … Via <字段>` 的桩与契约同时打通。

### D43 B10 实施（2026-09-24 11:51–，代码 `3c5d8e6`）= `Implements <接口> Via <持有字段>`

**层序被实测走了一遍**：D42 说 Via 今天停在语法层。补完词法+语法之后语义层才露出面，
再补完裁决才有发码层 —— 一层量一层改，每层都有独立可观测面，没有"三层一起猜"。

**五个落点**
1. **词法**：`Via` 走软关键字四件套（`token.hpp` 枚举、`token.cpp::isKeyword` 链、
   `lexer_keywords.cpp` 表、`parser_helpers.cpp` 的 `canBeName` 软表）。语料核查：
   `tests`/`archive`/`publish` 的 `.bas/.cls/.frm/.ctl` 里 `\bvia\b` **零命中** → 登记零误伤；
   另加正例 `itf_p04_via_soft_ident`（`Dim Via As Long` 照样过）。
   `isStatementStart` **故意不加**（与 `Inherits` 同口径：它是子句中间词，不是语句开头）。
2. **语法**：`parseImplements` 吃完点号限定名后可选 `Via <名>` → `ImplementsStmt::viaField`。
   没写 Via 不进分支 = 存量路径一字不动。`ast_clone.cpp` 的 ImplementsStmt 分支跟着拷 viaField。
3. **裁决落在 stage 2.7 新增的 Pass D**（`runInterfacePrepass` 末尾），**不在语义层**：判定要看
   "字段类型那个类自己 Implements 了没有",而那些类的符号要到 3.5 才注入本模块作用域 ——
   只有 2.7 把整工程的模块表看全。产物 `Driver::vias_`（小写类模块名 → `ViaView{ifaceKey,
   fieldName, holderModule}`）**同时**喂语义层与发码层（一份来源，正是 D42-3 立的那条）。
   判死两条：`VB3029`（不在类模块 / 被委托名不是 Interface 块 / 目标不是本类的对象字段 /
   字段类型不是工程内的类）、`VB3030`（那个类没实现该接口，**含链式 Via** —— A Via f(f:B)、
   B Via g 运行期能构成无限回环，一次诊断只报一条）。
4. **语义**：`checkNewStyleInterface` 命中委托 → 该接口逐槽 `VB3012` 全免（D42-2 量出的那份手写
   就是这几条），但本类自家写过的成员**仍按接口槽校签名**。
5. **发码**：`emitIfaceImplTables` 的槽循环里 `ivFindImplMember` 取不到实现时，若本类委托了该接口，
   发一个转调"持有对象同名槽"的适配器：
   `vb6_ivref_<I>* h = me->m_h ? &me->m_h->__iv_<I> : NULL;` 然后 `h->vt-><slotKey>(h, ...)`。

**为什么绕接口表、不直调 `vb6_<持有类>_<成员>`**（本批最关键的一次选型）：实现接口的成员按 VB6
惯例写 `Private`，而 Private 过程发成 C 的 `static`、跨翻译单元连不到 —— 这正是 B09 对 `MyBase`
私有目标判死的那条限制，照直调只会把错误推到链接期。表项发在持有类自己的 TU 里，函数指针从
对象里读，天然可用；代价是一次间接调用，而这与 COM 客户端看到的形状本来就一致。
顺带白拿到**逐槽合成**：本类写了的成员照旧直调，只有没写的槽走表（用例 `CViaDeleg` 的
`Property Get Name` 自家的、其余五个槽委托，VIA5 钉住这条）。

**Nothing 持有字段**：`Sub` 槽 no-op，有返回值的槽回 `<T> __via0 = {0};` —— 标量、指针、聚合
三种返回类型同一个写法都合法（Variant 的 `{0}` 就是 VT_EMPTY），不必按类型分岔。

**顺手量出三条既有洞（与 Via 无关，登记、不在本批修）**
- 接口变量上的 **`Property Let`/`Set` 写**：`s.Name = "x"` 发成 `vb6_s_Name = ... /* Module.Name */`
  → `error C2065`。"经接口变量写属性"这条路从来没通过（`tests/itf_xmod` 的接口只有 Get，
  所以一直没暴露）。后果：Via 的 `put_*`/`putref_*` 槽目前只有**发射形状**可证、运行期到不了。
- **带 `Optional` 形参的接口槽**：调用点发 `s->vt->greet(s, <实参>)`，而表项签名带 `_has_` 尾参
  → `error C2198` 实参太少（B04 的调用点没接可选参数标志位）。
- 接口块声明在 **`.bas`** 里时，`Dim x As I` 的类型认得出（发了 `vb6_ivref_I*`），但成员调用落回
  模块变量形状（`vb6_x_M()`）→ 编不过。`tests/itf_via` 因此改用 B03 的头行宿主。
  D1 说的"块形式可与别的声明共存于 `.bas`"目前只在语法/契约层成立 —— 这条要单独立项。
三条同属"到达槽的那条调用路形状缺失"，改在 caller 侧。

**验收**：`tests/itf_via`（IViaShape/IViaNamed 头行宿主 + `CViaHolder` 全 Private 实现 +
`CViaBare` 纯委托 + `CViaDeleg` 一半自家）断言 VIA1..VIA9 + VIA-DONE，**x64 与 x86 各 9/9**
（布局类纪律：适配器要解引用别的类的结构体字段，D40 那条）；A/B = `.build/pre_b10_C3.exe`
对同一工程停在 `VB2003`+`VB2002`（D42 的原始读数）；负例 `itf_n21`/`itf_n23`/`itf_n24`
（单文件）与 `itf_n22_via_holder_not_impl`（双文件 VB3030）；逐字节护栏**扩到 10 文件**
（新增 `itf_xmod\XWriter.vbp`、`cls_inh\Inh.vbp` —— 本批动的是接口适配器发射器和 2.7，
这两个工程是最直接的压力点）10/10 全同；`-Category syntax` 84→89。

### D44 B11/C01 前置实测（2026-09-24 13:24–，只量不改码；探针在 `.build/probe_b11/P1..P4`）

**结论先行：`CoClass` 今天的可观测面只在语法层，且层与层之间断在第一行。** C01 的验收面因此是
`--syntax-only` + `--dump-ast`（D15-5 那条单文件通路），不是运行期。

四份探针（P1 = `.bas` 里整块含内联 Interface 定义、P2 = `.cls` 头行宿主、P3 = `As Circle`/`New Circle`
消费点、P4 = 畸形块）对 `.build/pre_b11_C3.exe` 的读数：

1. **`CoClass` 不是 token**（`src/lexer/` 零登记，只是 Identifier）→ 模块级主循环落进兜底分支，
   第一行报 `VB2002 unexpected token at module level: CoClass`。关键形状：**错误恢复只 `skipToNextLine`
   跳一行**，块内其余行于是各自按"模块级"重新解析 —— 所以 P1 报出 20 条互相无关的错、P2 只 2 条
   （13 行的 `[CoClassId]` 名字已在白名单 → 14 行 `Interface` 被当成**顶层接口块**吞掉、只剩 17 行
   `End CoClass` 报错）。一句话：现在的 CoClass 块不是"整块被拒"，而是**被拆散后各部分各自为政**，
   这比整块被拒更危险（P2 那份"看起来只差两行"其实把块内的接口声明偷运成了顶层声明）。
2. **属性白名单缺两词**：`itfKnownAttrNames`（`parser_interface.cpp:42-50`）收了 `coclassid`/
   `comcreatable`/`coclasscustomconstructor`/`default`/`source`，但 **`ProgId` 与 `Implementation` 没登记**
   → `VB2012 Unrecognized attribute line`。026 四节样本正好用了这两个词，所以 C01 不补就连样本都进不去。
3. **布尔实参不支持**：`[ComCreatable(True)]` → `VB2012 Attribute line argument must be a string or an
   integer: True`。属性行整体是一个 token，参数用 `strtoll` 现解析，`True` 在这里**不是 token** ——
   所以修法是解析器认 `"True"/"False"` 字面，不是改 lexer。
4. **同行"属性 + 声明"今天不通**：`[Default] Interface ICircle` → `VB2003 expected end of statement`（列 15）
   + `VB2002 Attribute line must precede an Interface declaration`，根因是 `parseBracketAttrLine` 末尾无条件
   `expectEndOfStatement()`。026 四节把 `[Default] Interface ICircle` 写在同一行 → 必须处理。**裁决：不动
   Interface 那条路**（`itf_n06` 负例正守着那句诊断，且改了等于给契约块开新语法），只给 CoClass 块内
   用一个 `requireOwnLine=false` 的变体：属性行吃掉后若同行还有内容，留给块体循环判定。
5. **`End CoClass`** 单独报 `VB2002 unexpected token at module level: End`（P3/P4 各一处）。
6. **语料核查**（026 二节那份的复核）：`tests/` + `archive/` 里 `coclass` 作标识符 **0 次**（唯一命中是
   `coclassid`/`coclassinfo` 之类）→ 软关键字登记安全。
7. **同名不同物，别撞车**：仓库里"coclass"三处既有设施全是**外部类型库侧**的（`typelib_builder_coclass.cpp`
   的 `addCoClass`、`symbol_table.hpp:41` 的 `ComClass`、`vbp_parser.hpp:84-85` 引用工程枚举 coclass 烘 ProgID 表），
   与新的语言块无关；`symbol_table.hpp:248-252` 已有一套"本工程类模块遮蔽类型库 coclass"的机制 —— **C03 做
   `As <CoClass>` 名解析时必须先说清"块名与 typelib 的 ComClass 同名谁赢"**，别顺手新造第三套优先级。

**C01 落点裁决**（严格守在 026 六节"不校验、不发码"那一格）：
- 词法四件套照 B10/Delegate：`token.hpp` 枚举 + `token.cpp::isKeyword` + `lexer_keywords.cpp` +
  `parser_helpers.cpp` 软表；`isStatementStart` 不加。
- AST 新增 `CoClassDecl`（含 `CoClassIfaceRef{ifaceName,isDefault,attributes}`），存进 **`Module::coclasses`** ——
  与 `interfaces` 同一个手法：**不进 `declarations`**，语义/发码层看不见它，所以"零回归"是结构性的而非测试性的。
  可观测面靠 `--dump-ast` 的 printer 一行。
- 块体语法：**只收属性行与 `Interface <名>` 引用行**（`Interface IShape` 在 026 样本里没有 `End Interface`
  = 引用已声明的接口，不是内联定义），`[Default]` 标默认。内联定义/字段/过程一律按非法行报错。
- **属性行归属要单独定规则**（实施时才暴露，026 四节没写）：样本把 `CoClassId/ProgId/ComCreatable/
  Implementation` 四条写在**首个契约条目之前**，而 `[Default]` 只能贴在条目上 —— 若一律"属性行留给下一条
  `Interface`"，那四条就挂到了第一个接口上（`--dump-ast` 一眼看穿：`Interface IShape (4 attrs)`）。
  定案 = 按"**名字 + 位置 + 是否同行**"三条合判：`[Default]` 永远归条目；同行紧跟 `Interface` 的归条目；
  其余在首个条目之前归块、之后归下一条目。**不靠空行**（空行不是语法）。
- 属性通路复用现成的 `InterfaceAttr` 与 `parseBracketAttrLine`（顺带补第 2、3 条：白名单加两词、
  布尔折成 `numValue` 1/0 —— 不加新字段，因为 C02 求解身份时 `True`/`1` 同义）。
- **C01 的代码进了 `parser_interface.cpp`，没另立 `parser_coclass.cpp`**（本条覆盖上面"新文件"的设想）：
  方括号属性行的全部语法（白名单、实参形态、`requireOwnLine`）都在那个文件里，CoClass 块要复用的正是它；
  分家就得把 `itfAsciiLower` 一类判定复制一份，等于把"一处属性行语法"拆成两处。CMakeLists 因此未动。
- **不新增诊断 ID**：畸形子句只用 `Parse*` 家族（2001/2002/2005/2012）；校验语义（宿主 `.frm/.ctl`、
  `[Implementation]` 指向、默认接口 ∈ 集合、契约聚合、拒绝清单四类）整片留给 C03，`As`/`New`/`CreateObject`
  改写留给 C05，身份求解唯一函数留给 C02。
- 于是这几条 C01 **明知不做**（P5 探针量过形状，都在 C03 账上）：同一块里 `Interface X` 写两遍不报重复、
  块名与别的 CoClass/类/接口重名不报、引用的接口名不存在不报、`[Default]` 标两条不报。
- `[Default, Source]` 这种**逗号并列**的属性名今天仍整串比对 → `VB2012` 不认（026 四节末那条"只接受并存档"
  要等真做元数据时才补拆解，不在 C01 顺手加）。
- 一处**顺带**的既有文案要跟着改：`parser_module.cpp:155` 那句 "Attribute line must precede an Interface
  declaration" 现在也要涵盖 CoClass（同批改 `itf_n06` 的 needle，否则负例假绿）。

### D45 B11/C01 实施（2026-09-24 13:24–，代码 `e7c7a31`）= `CoClass…End CoClass` 落到 AST

**落点**（15 文件，全在 `src/` 的语法侧 + 用例；CMakeLists 未动）：

- 词法四件套：`token.hpp` 枚举 + `token.cpp::isKeyword` + `lexer_keywords.cpp` + `parser_helpers.cpp` 软表。
- AST：`ast_enums.hpp` 的 `ASTNodeKind::CoClassDecl` + `ast_fwd.hpp` + `ast_decl.hpp`（`CoClassDecl` /
  `CoClassIfaceRef{ifaceName,isDefault,attributes,loc}` / `Module::coclasses`）+ `ast_visitor.hpp` 的 visit 钩子
  + printer 三处（`visit(CoClassDecl&)`、`visitDeclHelper` 的 case、`print(Module&)` 的循环）。
  **这一串就是"新增一个模块级块类型要登记的位置"清单**，C04（attribute 折算）与以后任何新块照它走，省一轮 grep。
- parser：`parser.hpp` 两个声明（`parseCoClassDecl`、`parseBracketAttrLine` 加 `requireOwnLine`，默认 true
  → Interface 那条路行为不变）+ `parser_interface.cpp` 末尾的 `parseCoClassDecl` +
  `parser_module.cpp` 主循环 Interface 分支旁多一个 CoClass 分支。
- 属性行通路：白名单补 `progid`/`implementation`；实参形态补 `True`/`False`（折进 `numValue`）。

**三条以后还用得上的判断**：

① **结构性零回归优于测试性零回归**。新块存进 `Module::coclasses` 而不是 `declarations`，语义层与发码层
没有任何一条路能看见它 —— 于是"10 文件 `--emit-c` 全同"是**必然**而不是运气（护栏照跑，只是它测的是
"我没有手滑"，不是"我猜到了所有影响面"）。这与 B01 给 `interfaces` 的做法同构，D44 把它写成了选型理由。
② **只到 AST 的批，观测面要两条腿**：`--syntax-only` 静默只证明"不报错"，证明"属性行进到了哪一层"得靠
`--dump-ast`。本批正是 printer 一行 `Interface IShape (4 attrs)` 暴露出"写在首个条目之前的身份四件套挂到了
第一个接口上"—— 光看 `--syntax-only` 退出码这条 bug 会一路带到 C02 的身份求解里才炸。
③ **块的属性归属规则要按"名字 + 位置 + 是否同行"合判，不能靠空行**（空行不是语法）。定案见 D44；
`[Default]` 折成 `isDefault` 布尔位、不再同时留在 `attributes` 里，免得 C03 面对两处真相。

**验收**：`-Category syntax` **89→95**（新用例 `itf_p05`/`itf_p06` + 负例 `itf_n25`..`n28`）；A/B =
`.build/pre_b11_C3.exe` 对 `n25` 停在 `VB2002 unexpected token at module level: CoClass`（D44 的原始读数）；
10 文件 `--emit-c` 对 `pre_b11_C3.exe` **10/10 逐字节全同**；另做一条**惰性证明**（比护栏更贴本批）：
同一份 `p05` 工程**删掉 CoClass 块**后 `--emit-c` 与保留块时逐字节相同（唯一差异是两个源文件名注释，
`.build/probe_b11/inert/{w,n}.c`）—— 块在发码层一行都不发；带块的 `p05` 还做了**真编译 + 真运行**
（打印 `ok`）。C01 没有任何语义可断，所以本批**不建 vbp 工程**（一次显式的"批粒度小于一个工程"取舍，
运行期断言从 C05 起才有承载面）。门 = 本次 push（`3c5d8e6..e7c7a31`）触发的 Actions run，编号与结论记在状态头。

**下一格 = C02**（026 六节）：身份求解唯一函数 `CLSID/IID/ProgID` 三优先级 + 两次构建可复现。C01 已经把
输入面备好了 —— `InterfaceAttr` 的字符串/整数/布尔三形态与 `CoClassDecl::attributes` 就是它的读取起点。

### D46 B11/C02 前置实测与裁决（2026-09-24 14:17–，只量不改码）

**这一格没有运行期可观测面，所以先得把"怎么验收"定下来。** 实测四件事：

1. **现成的 mint 底子是两个 `.inc` 内的局部 lambda，跨不出翻译单元**：
   `cgen_util_dllentry_prelude.inc:40-72` 的 `generateClsid` / `generateIid`（同一套 FNV-1a × 4 条种子链，
   只差 4 个初始常量）。它们是 `CCodeGen::generateDllEntry` 函数体的片段（`cgen_util.cpp:32-36` 把四段
   `.inc` 串成一个函数）→ **别的层想复用只能复制一份**，而那正是 026 三节"禁止三处各读一遍"要防的事。
2. **这条 mint 路今天没有任何测试覆盖**：全部 `tests/*.vbp` 都是 `Type=Exe`，`generateClsid/generateIid`
   只在 ActiveX DLL 的 coclass 表里跑。所以"把 legacy lambda 换成调用新函数"这件事**在 C02 里做不起**
   —— 改了也没有逐字节护栏能证明没改坏。裁决 = **C02 只新建唯一入口，legacy 两枚 lambda 一字不动**，
   替换动作按 026 七-1 / D15-8 归 **B13/B16 分叉**（届时先要有 DLL 形状的用例）。新函数用**不同的种子常量
   + 带前缀的 seed 串**，保证两套 mint 不可能撞车，并把这条写进代码注释。
3. **vbp 三段式与工程名来源都还在 026 引的位置**（复核无漂移）：解析在 `vbp_parser.cpp:66-112`
   （`Class=Name; x.cls; {CLSID}`，只有 `{...}` 形态才存进 `entry.clsidStr`），收集在
   `driver_compile.cpp:74-81` 的 `classClsidMap_[lower(moduleName)]`（**键是类模块名**，不是 CoClass 名）。
   工程名侧现成只有 `projectBaseName_`（`ExeName32` stem，否则 vbp 文件名 stem，单文件编译为空），
   vbp 的 `Name=` 字段只在 `driver_compile.cpp:117` 就地用来兜 DLL 的 `dllProgId`，**没存下来**。
4. **可观测面要找第三趟才对**：`--dump-ast` 在 `driver_compile.cpp:313` 就返回了，跑在 stage 2.7
   **之前**，拿不到工程名与 vbp 表；`--emit-c` 又不许在 C02 发码。本条原写"那就发一条 note 级诊断"，
   **实测是错的**：`Diagnostics::toString()` 确实无条件拼所有级别，但 Driver 只在**某个阶段失败**时
   才把它整体打印（`driver_compile.cpp` 里 10 处 `if (!runXxx()) { std::cerr << diag_->toString(); }`），
   所以 note 在成功的编译里根本看不见 —— 用它当验收面会得到一个"永远为空"的断言。
   定案 = 走 **stderr 的 `C3: ...` 信息行**（`driver_compile.cpp:63` 的 "C3: 加载工程" 是同族先例，
   而且只在真有 CoClass 块时才发，零新语法工程一字不多），并且**不新增诊断 ID**（那条 note 没有读者）。
   两条踩坑留档：① 信息行绝不能走 stdout —— `--emit-c` 的 stdout 就是 C 文本，混一行就毁产物；
   ② 这是 D44/D45 那条"每层的可观测面不一样"的第三次应验，而且这次是**先假设了一个面、量了才发现它不通**。

**C02 落点裁决**：

- 新单元 `src/semantics/coclass_identity.{hpp,cpp}`：`CoClassIdentity{clsid,iid,progId,defaultIface,impl,
  clsidSource,progIdSource}` + **唯一入口** `resolveCoClassIdentity(block, env)` + `mintGuid(seed, 常量组)`。
  放 `src/semantics/` 与 `interfaces_registry.hpp` 同族（plain 结构、只读视图、不拥有 AST）。
- 三档优先级按 026 三节实现，其中 **vbp 档的查表次序**要新定（026 没写）：`[Implementation("X")]` 的 X
  先查 `classClsidMap_`，查不到再用 CoClass 块名查（VB6 的三段式挂在类模块上，而 CoClass 名常与该模块同名）。
- `<Proj>` 取值 = vbp `Name=` > `projectBaseName_` > 字面量 `"VB6EXE"`（第三条兜法沿用
  `driver_codegen_dll_entry.inc:27` 已有的兜序，不新造）。为此在 `driver_compile.cpp` 存一个
  `vbpProjectName_`。**与 com_entry 那侧用 `projectBaseName_` 是分叉的**，理由 = VB6 的 ProgID 语义是
  `<工程名>.<类名>`，而工程名就是 `Name=` 字段；这条分叉要写进 B13/B16 的对账清单。
- IID 档 = 默认接口的 `[InterfaceId]`（`IfaceView::guid` 早在 B02 就存了）> `"itf:<Proj>.<接口名>"` mint；
  块里没有 `[Default]` 条目 → `iid` 留空并在 note 里标 `no-default`（存在性校验按 D44 归 C03）。
- 求解结果挂 **`Driver::coclassIds_`**（key = CoClass 名小写），与 `ifaces_`/`vias_` 同族；
  求解发生在 stage 2.7 的**新 Pass E**（那时接口登记表刚建好、工程名与 vbp 表都已就位）。
  注：**同名两个 CoClass 块 = 后一个不覆盖前一个**（`emplace` 首值胜），这是 C03 的重复名检查欠的账，不在 C02 报。
- 用例面：三条档位各一条 note 断言 + **可复现性两条**（同一输入跑两次逐字节相同；换一个工程名再跑，
  确定性档跟着变而显式档一字不动）。为此给 `run_tests.ps1` 加一个 `Test-SyntaxNote`
  （退出码 0 **且** 输出含 needle —— 现有 `Test-Syntax` 只看退出码，`Test-SyntaxFail` 只看非 0）。
  确定性档的期望串用 python 独立复算 FNV-1a 得到，**不拿编译器自己的输出当基线**（否则等于没测）。


### D47 B11/C02 实施（2026-09-24 14:17–，代码 `f0b820d`）= 身份求解唯一函数

**落点**：`src/semantics/coclass_identity.{hpp,cpp}`（新编译单元，进 `vb6c3-core`）+
`runInterfacePrepass` 末尾的 **Pass E** + `Driver::coclassIds_`（key = 块名小写）+
`driver_compile.cpp` 存一份 `vbpProjectName_`。发码层一行未动。

**唯一入口的形状**：`resolveCoClassIdentity(const CoClassDecl&, const CoClassEnv&)` 是纯函数
（不写诊断、不碰全局），`CoClassEnv{project, vbpClsids, ifaces}` 把"块外面的世界"当参数传进去
—— 这样 C05/B13/B15 的调用点各自给上下文，而**算法只有一份**。三档优先级按 026 三节，
其中三处 026 没写、这次必须定的口径：

- **vbp 档的查表次序**：三段式挂在**类模块名**上（`classClsidMap_` 的 key 就是模块名），而 CoClass
  块名不必等于任何模块名 → 先按 `[Implementation("X")]` 的 X 查，查不到再按块名查。
- **`<Proj>` 的兜序**：vbp `Name=` > `projectBaseName_` > 字面量 `"VB6EXE"`。第三条不是新造的，
  是 `driver_codegen_dll_entry.inc:27` 那段注释里已有的兜法。**与 com_entry 那侧用
  `projectBaseName_` 是分叉的**（VB6 的 ProgID 语义是 `<工程名>.<类名>`，工程名就是 `Name=`），
  这条要进 B13/B16 的对账清单。
- **seed 的大小写**：`coc:`/`itf:` 前缀 + 工程名与块名**一律小写**（VB 大小写不敏感，`Circle`
  与 `circle` 必须同一个 GUID）；而 ProgID 默认值保留块名**原样大小写**（它是给人看的名字）。

**两条判据级别的纪律**：

① **期望串必须由另一套实现算出来**。三条断言里的两个 GUID 是 python 独立复算 FNV-1a（同一
seed 串、同一常量组）得到的，不是拿编译器自己的输出抄回去 —— 否则用例只能"重述实现"，
mint 换算法也测不出来。用例还钉了一条更强的形状：**同一批源文件、只换 vbp 的 `Name=`**，
派生档整串跟着动、显式档一字不动（可复现性 = "只有声明决定身份"这句话的直接证伪面）。
② **legacy 那两枚 mint lambda 一字未动**（D46-2）：它们只在 ActiveX DLL 路径上跑，而全部
`tests/*.vbp` 都是 `Type=Exe` → 改了没有任何护栏能证明没改坏。新 mint 用**不同的常量组**，
并存期间不可能撞车；合并动作按 026 七-1 / D15-8 归 B13/B16。

**验收**：`-Category syntax` **96→99**（`cc_id_explicit_tier` 三条档位全串相等 +
`cc_id_project_name_scope` 换工程名 + `cc_id_repeatable` 同一输入跑两次逐字节相同，新助手
`Test-IdentityNote`/`Test-IdentityStable`）；A/B = `.build/pre_b11c02_C3.exe` 对 `Id.vbp` **一行身份
都不出**；10 文件 `--emit-c` 对 `pre_b11c02_C3.exe` **10/10**；`Id.vbp` **真编译真运行**（`Id.exe`
打出 `cc_id`，三条身份行照出）。门 = push `f660e25..f0b820d` 触发的 Actions run（编号见状态头）。

**C03 从这里接手**：块名重名 / 引用的接口名不存在 / 接口名重复 / `[Default]` 标两条这四条
"明知不做"，加上拒绝清单四类；其中"块名撞类型库 `ComClass`"要先定优先级（`symbol_table.hpp:248-252`
已有一套"工程类遮蔽类型库 coclass"的机制，别新造第三套）。**身份一律只许读 `Driver::coclassIds_`**，
不得再自己解析属性行 —— 这是本格留下的唯一入口约束。


### D48 B11/C03 前置实测与裁决（2026-09-24 15:07–，只量不改码；探针在 `.build/probe_c03/q1..q11`）

**结论先行：C03 要拒的九种形状里，八种今天一声不吭，只有一种已经被既有检查挡住了。**
每种都用 `.build/pre_b11c03_C3.exe`（= C02 收线二进制）跑 `--syntax-only` 量过：

| 形状 | 今天的读数 | 归谁 |
|---|---|---|
| 两个同名 `CoClass CCA` | 静默（`coclassIds_` 首值胜，第二个整块消失） | C03 |
| 块名撞**模块名** | 静默 | C03 |
| 同一块 `[Default]` 标两条 | 静默（求解取第一条） | C03 |
| 同一块 `Interface IOne` 写两遍 | 静默 | C03 |
| 条目引用不存在的接口 | 静默 | C03 |
| `[Implementation("NoSuch")]` 指向不存在 | 静默 | C03 |
| `[Implementation("IOne")]` 指向接口块 / `[Implementation("Q8bBase")]` 指向 `.bas` | 都静默 | C03 |
| `Inherits CCC`（CCC 是 CoClass 块名） | **已经报 `VB3020`**："inherits unknown base class 'CCC' (the target must be a class module in this project)" | 只需改文案 |
| 块列了 `Interface IShape`，实现类**根本没写** `Implements IShape` | **静默** | C03 的正菜 |
| 实现类写了 `Implements IShape` 但缺槽 | `VB3012`（B02 就有，逐槽一条） | 已覆盖 |

三条要记住的实测细节：

1. **`Inherits` 一个 CoClass 已经被挡**（`VB3020`），但理由是错的（"unknown base class"，而这个名字确实存在，
   只是不是类）。026 五-6 说"v1 拒，照 3022 那批同族处理" → 这条的活是**把文案改成指名"CoClass 块不能被继承"**，
   不是新造检查。文案一改，`tests/cls_inh`/`ci_n*` 里凡按 `VB3020` 原句断言的都要跟着核一遍。
2. **契约聚合是真缺口**：B02 的 `checkNewStyleInterface` 只对**写了 `Implements` 的类**跑；CoClass 块里的条目
   不会给实现类补上这份义务，所以"块说 CImpl 满足 IShape、而 CImpl 压根没 Implements"今天无人管 → 这正是
   C03 唯一的实质新增判定，也是 026 五-1"复用比对器"这句话的落点。
3. **聚合检查不能放 stage 2.7**：实现类的成员在 **3.4 `mergeInheritedMembers()` 之后**才带得上祖先实现
   （B07b 的转发桩那时才存在），而 2.7 只看得到本模块自有声明 —— 放 2.7 会把"基类实现了、派生类没重写"
   误判成缺失。裁决 = 新开 **stage 3.4c `runCoClassContractCheck()`**（3.4 之后、3.5 之前），
   读 `Driver::classes_` 的链与合并后的符号，比对器仍走 `interface_sig.hpp` 那套槽键与签名口径。

**C03 这一格切两半（026 六节把 C03 写成一格，但里面有两种性质的活）**：

- **C03a（本轮做）= 形状与名字校验**，全部能在 2.7 就地判定（`modules_` + `ifaces_` 都看得见，同 B10 Pass D
  的判据位置）：块名重名 / 撞模块名 / 撞接口名、条目重复、`[Default]` 多标、条目引用不存在的接口、
  `[Implementation]` 指向不存在或不是类、EXE 工程 `[ComCreatable(True)]`。
- **C03b（下一轮）= 契约聚合**（上面第 2/3 条）+ `As <CoClass>` = 默认接口视图（026 五-3）。
  `As` 实测今天也是**静默通过**（探针 q11：`Dim x As CCC` + `x.A` 一声不吭 → 未知类型名被当 Variant 晚绑定），
  所以它不是"从报错改成通过"而是"从静默错改成有类型"，和聚合一样要动语义层，别和 C03a 混一批。

**新增诊断按族给三个号，不混用**（沿用 3015-3018 的分法：一个号管一类判据）：
`SemCoClassEntryInvalid = 3031`（条目：未知接口名 / 重复条目 / `[Default]` 多标 / 默认接口不在集合里）、
`SemCoClassDuplicate = 3032`（名字撞车：块名重复、撞模块名、撞接口名）、
`SemCoClassNotSupported = 3033`（v1 边界：`[Implementation]` 不是类 / EXE 工程 `[ComCreatable(True)]`）。
文案一律 ASCII（D12）。

**一颗先拆掉的雷**：C02 的用例工程 `tests/cc_id/Id.vbp` 是 `Type=Exe` 且 `CCCircle` 写了
`[ComCreatable(True)]` —— 正是 3033 要判死的形状，而 `Test-IdentityNote` 要求退出码 0，
所以校验一落地，C02 那三条断言立刻变红。定案 = 走 026 的语义正道：**把 `cc_id` 的 `[ComCreatable(True)]`
改成 `False`**（needle 跟着改），EXE+ComCreatable 的拒绝另立 `cc_neg` 负例；不把 `cc_id` 换成 DLL 工程
（那会让身份用例依赖 `Type=DLL` 的 `Startup=`/tlb 行为，把两批的失败面搅在一起）。

### D49 B11/C03a 实施（2026-09-24 15:07–，代码 `e515d89`）= CoClass 块的形状与名字校验

**落点**：`driver_interface.cpp` 的 **stage 2.7 新增 Pass F**（判据位置与 Pass D 同一条理由：
只有此刻整工程模块表 + 接口登记表同时可见）+ 三个诊断号（3031 条目 / 3032 名字 / 3033 v1 边界）+
`runInterfacePrepass(const CompileOptions&)` 多收一个参数（只为 `isDll` 一项，`driver_compile.cpp`
一处调用点跟着改）。发码层零改动。

**八条判据**（D48 表格里"今天静默"的那八种，逐条对上探针）：块名重复 / 撞别的模块名 / 撞接口名（3032）、
条目引用不存在的接口 / 把类模块当接口列进集合 / 条目重复 / `[Default]` 多标（3031）、
`[Implementation]` 指向不存在或不是类模块 / EXE 工程 `[ComCreatable(True)]`（3033）。

**两条要留下的判断**：

① **块名 = 自己宿主模块名必须豁免**。第一版按"撞模块名就拒"写完，实测立刻把自己的正例 p07 打回 ——
而 VB6 最自然的写法就是 `Widget.cls` 里写 `CoClass Widget`（实现类即宿主）。豁免条件取
"撞的就是**装这个块的**那个模块"（与 B03 的接口宿主同一条理由），跨模块撞名照旧拒。
这条 026 没写，是实施时**被自己的用例逼出来的**（又一次"先量再写"，只是这次量的是自己的正例）。
② **上一批的用例是这一批的雷，而且要当面拆**（D48 已预告）：`cc_id`/`p05`/`p07` 三处正例分别声明了
不存在的实现类与 EXE+`[ComCreatable(True)]`，校验一落地就全红。定案不是放宽校验，而是把用例改成正面形状
（`cc_id` 补一个真的 `CircleImpl.cls`、`ComCreatable` 改 `False`、p07 改成宿主同名惯用法），
`True` 的那一面另立负例 `itf_n34`。**一个校验批的完成标志是"旧正例与新负例同时全绿"**，
只加负例不改正例的批都是把红灯留给下一轮。

**验收**：`-Category syntax` **99→107**（`itf_n29`..`n34` 单文件 + `itf_n35_coclass_legacy_cls`、
`itf_n36_coclass_vs_module` 双文件 + p05/p07 改造后仍静默 + `cc_id` 三条身份断言一字不差）；
A/B = `.build/pre_b11c03_C3.exe` 对 `n29`/`n34` **零命中**（这两条诊断本批之前根本不存在）；
10 文件 `--emit-c` 对 `pre_b11c03_C3.exe` **10/10 逐字节全同**（本批没碰发码，护栏测的是"没手滑"）；
`tests/cc_id/Id.vbp` 真编译真运行（三条身份行照出、`Id.exe` 打出 `cc_id`）。门 = push 后的 Actions run，
编号见状态头。

**下一格 = C03b**（D48 已切好边界）：`runCoClassContractCheck()` 放 **stage 3.4c**（3.4 成员合并之后、
3.5 之前），复用 `interface_sig.hpp` 的槽键与签名口径，判"块列出的每个接口，实现类（含祖先）必须满足"；
外加 `As <CoClass>` = 默认接口视图。`Inherits` 一个 CoClass 已被 `VB3020` 挡住，**这条只剩文案活**
（把"unknown base class"改成指名"CoClass 块不能被继承"），改文案时要顺带核 `tests/cls_inh` 与 `ci_n*`
里按原句断言的负例。


### D50 B11/C03b 前置实测与裁决（2026-09-24 15:38–，只量不改码；探针在 `.build/probe_c03b/pa..pf`）

**量的对象 = 契约聚合到底要看得见哪些成员表**（D48-3 留的那句"祖先实现了、派生类没重写 在 2.7
看不见"要落实）。三种供给路径分别探：

| 探针 | 形状 | 今天（C03a 之后）的读数 |
|---|---|---|
| pa | 派生类自己写 `Implements I` + `Inherits` 一个实现了 I 的基类 | `VB3022`：**派生类自己实现新式接口**就不许继承 |
| pb | 基类写 `Implements I`，派生类 `Inherits` 它、自己不写 | `VB3022`：**基类实现新式接口**就不许当基类 |
| pd | 基类**只声明成员、不认领接口**，派生 `Inherits` 它 | **全静默**（契约根本没人查） |
| pe | 绑定类缺一个槽 | 静默（本批之后 = `VB3012`） |
| pf | 绑定类的成员 `ByVal`、接口写 `ByRef` | 静默（本批之后 = `VB3017`） |
| pc | `Dim c As CCCircle`（把块名当类型） | **静默当 `Variant`**，连 `Set c = Nothing` 都不问 |

① **D48-3 的理由成立，但要说准它成立在哪一半**：可达的"祖先供给成员"只有一种形状 —— 祖先自己
声明成员而不写 `Implements`（pd）。带 `Implements` 的两种（pa/pb）今天被 B08f 的 v1 边界整条挡死，
所以"2.7 看不见祖先成员"不是假警报，而是**只有 pd 这一条路真的需要链**。结论不变：契约比对排在
链表（2.8）与成员合并（3.4）之后，也就是新阶段 3.4c。

② **`As <CoClass>` 本批整条推给 C05**（pc 实测是"静默 Variant"）：状态头那条"要么不做、要么只做到
认得这个名字是类型"的半开方案，量出来是**净负收益** —— 认得它是类型之后成员访问立刻开始报错，
而派发仍要到 C05，用户拿到的只是"从静默变成一堆说不清的错"。所以这一格与派发**同批**做。

③ 观测面复用现成的：`--syntax-only` 一路跑到阶段 3.6 之后才 return（`driver_compile.cpp:429`），
所以 3.4c 的诊断天然进得了 `Test-SyntaxFail`/`Test-SyntaxFailMulti` 这条负例通路，不必再找 D46
那种 stderr 信息行的替代面。

④ **号段沿用，不发新号**：`VB3012 = SemInterfaceNotImplemented`、`VB3017 = SemInterfaceSignatureMismatch`
（状态头说的"VB3012 族"就是这两个）。缺槽与签名不符在 `Implements` 那边已是这两个号，CoClass 这边
语义完全同族，只差主语 —— 多发一个号只会让人以为契约有两种算法。

⑤ **没写 `[Implementation]` 的块跳过**（p05 与 `cc_id` 的 `CCMint` 都是合法形状）：没绑实现类就无从
判起。顺手立"块必须绑实现类"是 C05 的事（那时才要拿它 `New`）。

### D51 B11/C03b 实施（2026-09-24 15:38–，代码 `c4aaa4c`）= 契约聚合校验 + VB3020 文案

**落点**：`driver_compile.cpp` 新增**阶段 3.4c**（3.4b 与 3.5 之间，且刻意放在 `modules_.size() > 1`
那个 if **之外** —— 单文件工程里的块同样要判）。实现 `Driver::runCoClassContractCheck()` 写在
`driver_interface.cpp` 末尾，与 Pass D/E/F 同一翻译单元：四条判据共用的前提是"整工程模块表 +
接口登记表同时可见"。零回归 = 工程里没有 CoClass 块立即 `return true`（与 3.4/3.4b 的早退同族）。

**槽表怎么建**（这一格的全部技术内容）：按 `ClassChainView::chain`（自根到叶）**倒着走**、槽键
首见者胜 —— 于是"派生遮蔽祖先"自动成立，且遮蔽规则与 3.4 的成员合并同源，不会两套裁决。
成员级 `Implements I.M` 子句照 B02b 的规矩：写了子句的成员**只**按子句入座，不回落同名隐式匹配。
访问级别不参与判定（B02 就是这个口径，这里另立一套只会让两个比对器互相打脸）。

**两条刻意的取舍**：

① **不重构 `checkNewStyleInterface` 来共用它那张 impl 表**。它的子句记账牵动 `VB3019`（未认领子句
的兜底诊断），改它 = 把 B02 那条已发货路径的回归面全展开；而这里多写的只有"按槽键建索引"二十来行。
两侧真正共用的是 `interface_sig.hpp` 那一套 inline 口径（槽键、签名抽取、签名相等），
所以"什么算同一个槽、什么算签名不符"仍只有一份定义 —— 这是"复用"的正确粒度（026 五-1）。
② 类链登记表只收"能当基类用"的模块，表里没有就当该类没有祖先（只用它自己那张声明表），
不为此扩表。泛型宿主（D11 边界）直接跳过。

**VB3020 文案**（D49 留的尾巴）：只在"基名撞到一个 CoClass 块名"时换成新句，指名那是组契约的块、
没有成员表可继承；其余情形保持原句。号不变，`ci_n01`（真不存在的基名）与 `ci_n07`（基名是接口宿主）
两条按原句断言的用例因此一字未动 —— 改前 grep 过，全仓按那句断言的就只有这两处。

**上一批的用例又一次变成这一批的红灯**（D49-② 的第二次应验，而且这次连"理由"都提前写好了）：
新校验一落地，`p07`/`cc_id/CircleImpl`/`cc_id/VbpImpl` 三条正例全红 —— 前两条是 `ByVal` 对上接口的
`ByRef`（签名口径算它俩不等），第三条是 `VbpImpl` 根本没有 `Move`/`Label`。定案仍是**改正例**：
按接口的写法去掉 `ByVal`、给 `VbpImpl` 补齐两槽，不放宽校验也不给 `ByVal` 开后门。

**验收**：`-Category syntax` **107→112**（`itf_n37` 缺槽 / `itf_n38` 签名不符 / `itf_n39` `Inherits`
块名 三条负例 + `itf_p08` 祖先供给 / `itf_p09` Via 委托供给 两条正例；p09 兼作 B10×B11 的接缝用例，
`vias_` 里查到的委托关系免逐槽 `VB3012`、签名不符照报）；A/B = `.build/pre_b11c03b_C3.exe` 对
`n37`/`n38` **零命中**，对 `n39` 出的是**旧那句** "unknown base class"（新句本批独有）；
10 文件 `--emit-c` 对 `pre_b11c03b_C3.exe` **10/10 逐字节全同**；`cc_id/Id.vbp` 真编译真运行、
且 `--emit-c` 与改码前逐字节相同（本批一克发码都不产），`cls_inh/Inh.vbp` 真编译 exit 0。
门 = push 后的 Actions run，编号见状态头。

**下一格 = C04**（存量 `Attribute VB_Creatable`/`Instancing` 只读折算成 CoClass 记录，`ai/026` C04）：
它独立且小，而 C05（组内激活）按 D50-② 要把 `As <CoClass>` 类型识别与派发**一并**做，开工前先量
"`As <未知类名>` 今天在哪一层吞掉"（`Variant` 兜底点）与"`New <CoClass>` 直调实现类工厂"的落点。

### D52 B11/C04 前置实测与裁决（2026-09-24 19:25–，只量不改码）

**本轮第一件事不是写码**：本地 `fan/dev` 落后 `github/dev` 三枚提交（Asm x64 + 静态库 + 包引用三线，
64 文件 +3409 行，`tests/run_tests.ps1` 也并进来 262 行）→ 先 ff 到 `9099e16`，**逐字节护栏的 BASE
必须用合并后的树重建**（旧 `pre_b11c04_C3.exe` 是 `00ee1b2` 树的产物，拿它比会把别人的改动记到本批头上）。
重建后 syntax 基线 = **112**（别的写者的用例并进来了，不再是 107）。

**① 存量 attribute 到底是什么**（`tests/` + `archive/` 全量 `.cls` 统计，"头属性 = 名字不带点"）：

| 属性行 | 条数 | 值域 |
|---|---|---|
| `Attribute VB_Name` | 252 | 字符串字面量（已被 stage 2 消费成模块名） |
| `Attribute VB_PredeclaredId` | 143 | `False` 142 / `True` 1 |
| `Attribute VB_GlobalNameSpace` | 143 | `False` 142 / `True` 1 |
| `Attribute VB_Creatable` | 143 | **`True` 134** / `False` 9 |
| `Attribute VB_Exposed` | 142 | `True` 115 / `False` 27 |
| 带点的成员级行 | 40+ 种名字 | `m_oSocket.VB_VarHelpID` 12、`Item.VB_UserMemId` 6、`*.VB_Description` … |

三条读法：值域里**只有布尔字面量**（`True`/`False`，无 `-1`/`0`/字符串形态）；`VB_Creatable = True` 是
绝对多数（134/143）；而**整仓 `Attribute Instancing` 0 条** —— VB6 的 instancing 写在 `BEGIN` 头里、
parse 期已进 `Module::instancing`（`parser_module.cpp:39-67`）。026 六节字面写着"折 `Instancing`"，
按语料它没有可折的载体 ⇒ **不折、也不复制**（`Module::instancing` 继续是唯一真相，B13 要读就读它）。

**② 今天谁消费它们**：全仓只有两处 —— `driver_frontend.cpp:222` 的 `VB_Name`、`driver_generics.cpp:108`
对泛型类拒收 `VB_PredeclaredId`/`VB_Exposed`。其余属性行 parse 完就躺在 `Module::attributes` 里无人读
⇒ 状态头那条"根本没被消费"成立，折算是一条全新的信息通道。

**③ 折算会打到谁**（门内 blast radius）：`tests/` 里带这四行的 `.cls` 恰好 **8 个**
（`BalloonTooltips/cTT`+`ISubclass`、`Charts 2020/ClsResizer`、`ctor/CtorPlain`+`CtorThing`、
`pkg_cls/…/PkgThing`、`VBFlexGridDemo/Builds/…/IVBFlexDataSource{,2}`），全部在真编译的门项目里。
其中 `cTT`/`ISubclass`/`ClsResizer`/`PkgThing` 写的是 `VB_Creatable = True` ⇒ **若折算记录参与 Pass F，
它们当场撞 C03a 新立的 `VB3033`**（EXE 工程不得声明可注册 COM 服务器）。这就是状态头预警的那颗雷。

**裁决**：

1. **落点 = stage 2.7 新增 Pass E0**（Pass D 之后、Pass E 之前），造一份 `CoClassDecl` 记进
   `Module::coclasses`。与 Pass D/F 同一条理由（此刻整工程模块表才可见），且 Pass E 因此**仍是唯一身份
   出口**、`coclassIds_` 仍是唯一读数（D47 之后这是硬规矩）。附带的观测面红利：`--dump-ast` 在 stage 2
   就返回 ⇒ 折算记录对 AST 打印**天然不可见**，"手写的"与"折出来的"在两处面上分得清清楚楚。
2. **折算只记信息、不开判死**：Pass F 与 3.4c 一律按 `CoClassDecl::foldedFromAttributes()` 跳过。
   理由按 ③ 的实测数字钉死（134 条 `VB_Creatable=True` 全在 EXE），不是"先宽松以后收紧"。
   3.4c 的早退条件同时改成"只认手写块"，否则每个存量工程都要为一条空契约清单白建一张表。
3. **`CoClassIdentity::legacyFolded` 是给 VB3020 文案用的**：折算记录的名字**恒等于**一个类模块名，
   而 `classes_` 恰好不收接口宿主 `.cls` 与泛型类（`driver_classchain.cpp:381-385`）⇒ 若不加这个位，
   `Inherits <接口宿主名>`（宿主又写了头属性）会收到一句凭空捏造的"那是个 CoClass 块"。
   口径 = 折算侧与登记侧用同一组条件筛（`isInterfaceModule` / `classTypeParams` / 后缀 `.cls`），
   再加 `legacyFolded` 兜底，双保险。
4. **触发条件** = 四行任一存在。只写 `VB_Name` 的模块不折 ⇒ 本工程所有 `.bas` 用例、以及
   `cc_id/Id.vbp` 那三个手写块工程的输出**一字不变**（实测该工程折出行数 = 0）。
5. **手写块优先**：同一模块两者都有 → 不折、只报一行信息（026 六节"以手写块为准"）。
   选"信息行"而不是"警告诊断"：警告计数是门基线的一部分，而这一条既不是缺陷、也不该沉默。

### D53 B11/C04 实施（2026-09-24 19:25–）= 存量 header attribute 只读折算

**三处落点**（全在语义层，发码零改动）：

- `ast_decl.hpp`：`CoClassDecl` 加 `std::vector<std::string> legacyFoldKeys` + `foldedFromAttributes()`。
  **一条字段而不是 bool + 清单两处** —— 后者迟早漂移，而"折了哪几行、各折成什么值"本来就是要打出来的东西
  （元素形如 `VB_Creatable=True`）。
- `driver_interface.cpp`：Pass E0 折算 + Pass F/3.4c 两处跳过 + Pass E 信息行尾追
  ` folded-from-legacy: K=V …`。折算记录写 `[Implementation("<类自己>")]`、`VB_Creatable` →
  `[ComCreatable(n)]`（身份求解读这一枚），另三枚按**原名**带着（今天无人读，B13/B15 的类型库标志位
  从这里取）。每个键**取首行**，与 Pass E 对同名属性行"首值胜"同一口径。
- `coclass_identity.{hpp,cpp}`：`CoClassIdentity::legacyFolded` 由纯函数按
  `block.foldedFromAttributes()` 置位 ⇒ "身份 = f(块)" 这条没被破坏。

**观测面选在 identity 那一行、不另起一行**（D46-① 那条教训的正用）：`Invoke-IdentityText` 今天就是按
`identity:` 过滤 stderr 的，折进同一行 ⇒ **零新 helper** 就能断言 vbp 档那条；同时"哪个类折了、折成了
什么"与它的三档身份天然同行，读日志不必在两行之间做 join。手写与折算在同一份输出里的对比因此是**一眼**的：
`... impl='' comCreatable=False`（手写）vs `... comCreatable=True folded-from-legacy: VB_Creatable=True …`。

**一条 harness 坑**（下次直接避）：`.build/fnv_oracle_c04.py` 以 `coc:proj.name` 形态收参时，MSYS 会把带
冒号的 argv 当路径转换 ⇒ python 里 `split(":")` 少一段。期望值改为**在 python 内联 import 该模块**算。
三条期望 GUID 全部由这枚独立实现算出（`FoldBase` = `{96466C30-E240-55A4-9434-24F1C884C2AB}`、
`FoldWins` = `{EA2B2FD6-E5C6-5192-D0C9-A13BC6FC7859}`），不是把编译器输出抄回去（D47-② 第二次执行）。

**验收**：`-Category syntax` **112→116**（`itf_p10_coclass_fold` = 折算正例，兼作"EXE + `VB_Creatable=True`
不判死"的防回归，并断言带点的成员级行**没有**被折进来、折出来的名字仍能被 `Inherits` 当基类用；
`itf_p11_coclass_block_wins` = 两种形状同处一模块时手写块胜；`cc_id_fold_vbp_tier` +
`cc_id_fold_repeatable` = 一个**从不写 CoClass** 的 `.cls` 靠 vbp 三段式拿到 `(vbp)` 档 CLSID 且整串可复现）；
新增 helper `Test-SyntaxNote`（exit 0 + 必须出现 + 必须不出现：折算这种"只加信息"的批第一次有了
可断言的正向面）；A/B = `.build/pre_b11c04_C3.exe` 对 p10/p11/IdFold 三条**零命中**；护栏 **16/16**
（10 文件严格逐字节 + 6 个"会折算"的工程 `--emit-c` 输出逐字节全同、stderr 增量只允许 `C3: CoClass`
那一条通道）；`ctor_ok.vbp` 真编译真运行（`amt:42`/`CNT:7` 与改码前一致）。门 = push 后的 Actions run。

**下一格 = C05**（组内激活）：`As <CoClass>` 类型识别 + 默认接口派发 + `New <CoClass>` +
`CreateObject(ProgID)` 编译期改写，按 D50-② **同一批**做；开工先量"`As <未知类名>` 今天在哪一层被吞成
`Variant`"。折算记录到 C05 才第一次有运行时后果 —— 届时必须回答"折出来的类能不能 `New`"，
而 `VB_Creatable = False` 那 9 条要给出可解释的差别（今天它们与 True 只差在记录里那一位）。



### D54 B11/C05 前置实测与裁决（2026-09-24 20:42–，只量不改码；探针在 `.build/probe_c05/M1..M6`）

**CURRENT_BATCH 给的三件事全部量到落点**（不是量到结论），另加一条顺带量的既有洞。

**① `As <未知类名>` 到底在哪一层被吞：两层各有一处兜底，而且互相不认识。** 语义层
`semantic_analyzer_typeref.cpp:135` —— `return Vb6Type::Variant;  // 未识别类型 → Variant (宽松策略)`；
它第 4 步查的是 `symTab_.lookupModule`，而**类符号要到 stage 3.5 才跨模块注入**
（`driver_crossmod.cpp:22`；这段历史就写在 `semantic_analyzer_decl_type.cpp:42-52` 的注释里，
所以"别的模块里 `As 真类名` 也先当 Variant、事后靠 `variableTypeName`/`srcTypeName` 捞回"）。
发码层 `cgen_base_type.cpp:293` —— `return "void*";`。⇒ **"认得这个名字是个类型"没有单一入口**，
而且按**名字串**比类的位点不止这两个（`memberFieldTypes`、`srcTypeName`、COM 解包表各比各的）。
这条直接钉死 D55 的选型。另外全仓**没有**"未知类型名"这个诊断号（3xxx 里最近的
`SemUndeclaredIdentifier=3001` 只当警告用在 `Implements` 上）⇒ 新规矩只能新起一号 = `VB3039`。

**② 项目类的新建与释放通道 —— 量完的结论是"C05 不新增任何释放面"。**
`Set x = New C` → `vb6_cls_C_New()`，返回**具体结构体指针**（`cgen_expr.cpp:303-322` +
`cgen_com.cpp:40-110` 的 `emitFactoryBody`）；类变量赋值是**裸 C 赋值**，`Set x = Nothing` 只发
`target = NULL;`（`cgen_setlet_set_prop.inc:111-115`，那里的注释自己写着"项目类没有引用计数、
`vb6_ReleaseObject` 会 AV"），`__refcount` 只给写了 `Implements` 的类生成
（`cgen_iface_vtbl.cpp:353-358`）；`_Destroy` 全仓只有两个调用点（接口 Release 归零、窗体
teardown）⇒ **D41"EXE 里 `Class_Terminate` 没有触发点"那条缺口不因本批扩大、也不在本批修**。
把组名改绑到实现类 = 用户拿到的就是这条既有不变式，不会以为"新语法带来了新生命周期"。

**③ 折算记录让生效面变大这件事，实测结论是"风险还没发生，但闸门现在就得下"。**
`coclassIds_` 里今天确实有 8 个与类模块同名的条目（门内），但**类型路径一个字都不读它** ——
`.progId` 只被 Pass E 打进 stderr，`implName`/`comCreatable` 只被 Pass F 与 3.4c 用于校验。
所以三条准入写死在改名的入口：(a) `legacyFolded` 不进别名表；(b) 块名被同名模块占住时
**类/模块赢** —— `As Widget` 今天的含义不许被一块新语法改掉，而且这条裁决早有先例：
`cgen_expr.cpp:326` 的 Fix 177b 写的就是"coclass 被本工程同名类遮蔽时，创建点改走工程类工厂"；
(c) 块名等于实现类名时不改（改了也无害，三道都留着，抢答这件事不该靠运气）。
M4 对照实测：`As StructImpl`（真类）与接口视图那条路在改前改后**一字不差**。

**④ 三条坏读数**（BASE = `pre_b11c05_C3.exe`，`--emit-c` 实发，`--syntax-only` 全部 rc=0 零诊断）
—— 这三条就是 026 五-3/4/5 选"编译期改写"的立项理由：
M1 `Dim a As PCStruct` → `void* a = 0;` + `vb6_ComCall(a, L"Move", ...)`（空指针上的晚绑定）；
M2 `Set a = New PCStruct` → `vb6_NewObject(L"PCStruct")`（运行期按名字查注册表，EXE 工程根本没注册）；
M3 `Set a = CreateObject("ProbeApp.PCStruct")` → `vb6_CreateObject(...)` 原样发。
M5（块没有 `[Implementation]`）与 M1 同样静默 ⇒ 这一位改判死（`VB3039`）：块没有实现类这件事
编译器已经知道，没有理由让用户去猜。

**⑤ 顺带量到一条与"默认接口视图"直接相关的既有洞**（不是本批引入；D43 第三条的现场复现）：
接口块声明在 `.bas` 时，`Dim iv As IShape : Set iv = New RealImpl : iv.Move 5` 发出的 C 是
`vb6_StructImpl_Move((&(int32_t){5}))` —— **接收者丢了、类还挑错了**（`RealImpl` 的成员发成了
`StructImpl_`）。⇒ 本批**不把 `As <块名>` 做成接口视图**的第二条理由：那条路自己还缺一半。
第一条理由是契约按 026 五-1 只要求"实现类**连同祖先满足**这些槽"，实现类完全可以不写
`Implements`（`p08` 就是合法形状），那时候对象身上根本没有那份接口槽表可指。

### D55 B11/C05 实施（2026-09-24 20:42–，代码 `7f829ee`）= stage 2.7 Pass G 组内名字激活

**选型：就地改 AST 的类型位点，不给语义层和 cgen 各开一个别名入口。** 理由 = D54-① 那句
"类名被读的地方不止两处"：别名要一处不漏，就得每个消费者都记得问一遍，而未来还会再加消费者。
先例是泛型单态化（`src/ast/ast_clone.hpp` 头注"语义/cgen 对泛型零感知"）。收益当场兑现：
**发码侧零新分支**（`As`/`New`/`CreateObject` 三面同归一条已经实测通了的工程类路），且
`--dump-ast` 在 stage 2 就返回 ⇒ 用例 dump 出来的仍是用户写的那个名字。

**三个落点**：
1. 新单元 `src/driver/coclass_activate.{hpp,cpp}`（一个 `ASTVisitor` 管 `New`/`TypeOf` 这两个
   字符串位点，自己写的语句递归管类型引用与 `CreateObject` 的节点替换）+ 新号
   `SemCoClassTypeUnbound = 3039`。覆盖的类型位点：模块字段、局部 `Dim`（含 `Dim x As New`）、
   形参、返回值、`ReDim ... As`、UDT 成员。**不动**的位点记成 v1 边界：`Declare`/`Delegate`/
   事件与接口声明里的类型、`Inherits`（块名照旧 `VB3020`，五-6 没让路）、`Implements`。
2. 接线 = stage 2.7 Pass F **之后**的 Pass G（`driver_interface.cpp`）：身份只从 Pass E 的
   `coclassIds_` 读、**不二次求解**（D47 的规矩），三道重名闸门见 D54-③。
3. `CreateObject` 的改写做成 **AST 节点替换**（`Set x = CreateObject("<已知 ProgID>")` 的整个
   右值换成 `NewExpr(<实现类>)`），不是在 cgen 里仿冒发码文本。理由是实测出来的：
   `cgen_setlet_set_rhs.inc:130-153`（Fix 179a）认 New 形状靠的是**文本前缀 `(vb6_cls_`**，
   换成节点之后 `As Object` 目标的 IDispatch 包装自动成立（`cc_act` 的 CC5 就是这条读数）。
   只认"整个右值就是这一枚调用"这一种形状，嵌套形态（`Foo(CreateObject("p.c"))`）照旧走注册表
   —— 边界写进手册，不静默。信息行只在真被用到时打：
   `C3: CoClass 'X' activated in-project: type name -> class 'Y' (n type reference(s), m CreateObject rewrite(s))`
   （`n` 含 `New`/`TypeOf` 那两处，所以 cc_act 的 Circle 是 10 处引用 + 2 处改写）。

**验收**：`-Category syntax` 116→**118**（`cc_act_group_names` 两块各一行激活 + ProgID 文本；
`itf_n40_coclass_type_no_impl` 断 `VB3039`；`p10` 加第三份文件 + 新 Absent 断言"折算名当类型用
不产生激活行"）；新端到端工程 `tests/cc_act/Act.vbp` 断言 **CC1..CC9 + CC-DONE，x64 与 x86 各
9/9**（CC6 = 派生 `Overrides` 经组名答话 ⇒ 虚表没被静态化；CC9 = 组名与实现类名指同一个实例；
CC2/CC3 走字段/形参/返回值三个位点；CC5 = Object 目标的 COM 包装）；A/B：BASE 对 p10 与 cc_act
**零激活行**、对 n40 **零诊断**，NEW 三条全兑现；护栏 **16/16 逐字节**，其中 `cc_id\Id.vbp`
声明三个手写块而从不当块名当类型用 ⇒ 连 stderr 都不多一个字，正是 D54-③ 要的"闸门有效"证据。

**下一格 = B13**（P6 对外那半的第一格：IUnknown 三件套真实实现 + 对象布局 COM 化收尾）。
本批给 B13 留两条必读：D54-⑤ 那条接口视图旧洞（"默认接口"四个字要写进对外文档之前得收掉）、
D54-② 那条"类变量永不 Release"不变式（对外一旦发 IDispatch/IUnknown，引用计数就得**从无到有**，
`cc_act` 的 CC5 是今天唯一一条包装读数）。折算的 `VB_Creatable = False` 那 9 条与 True 的差别
仍然只在记录里那一位，B13/B16 要给出可解释的处理。

**收线后追加的两条实测（同一个 exe，`.build/probe_c05/M7..M11`；写在这里是因为两条都推翻了我
自己在手册上先写的话）**：
- `TypeOf t7 Is Circle` 编译过、改写也发生（激活行的引用数从 10 涨到 13），运行期答 **False**；
  对照 `TypeOf raw Is ShapeAct`（一份**从不写 CoClass** 的形状）**同样答 False** ⇒ 缺的是
  `TypeOf … Is <类名>` 本身，与 CoClass 无关（`itf_xmod` 的 TOF1..TOF3 走的是接口名那条路，是好的）。
  同一份代码里 `t7 Is Nothing` 答得对。手册已把 `TypeOf` 从"已激活的位点"清单里摘出去、当面写明
  "改写照做、结果仍为否"。登记给后续：**对外那半要用到 `TypeOf`/QI 时这条必须先收**。
- `ReDim arr(1) As Circle` 在 C 层撞 `error C2224: '.Move' 左边必须是结构体/联合`；对照
  `ReDim arr(1) As ShapeAct`（普通类）**报同一处、同一条** ⇒ `ReDim ... As <工程类>` 今天是坏的
  （数组元素没发成类指针），不是本批的改写造成的。手册同步改口径。UDT 成员那一位反而是好的：
  `Private Type THeld : c As Circle : End Type` + `Set h.c = New Circle : h.c.Move 6 : h.c.Area()=14`
  实测 **U1:OK** ⇒ 手册把"UDT 成员"留在已交付清单里，把 `ReDim` 挪出去。

**给下一批的一条流程教训（自己撞的）**：`tests/run_tests.ps1` 在仓库里是 **UTF-8 带 BOM + LF**
（不是记忆里写的 CRLF）。把插入段按 CRLF 写进去 = 一次 29 行的登记变成 1624/1598 的整文件改动，
`git diff --numstat` 一眼就能看出来；插完必须断言 `count(b"\r\n") == 0`。另外 bash heredoc 里的
`"$Tests\itf_neg\n40_..."` 路径里紧跟着出现 `n`（例：`\itf_neg` 后面接 `n40_...`）时会被当成换行吃掉一行，带 `t` 同理 —— 反斜杠一律用 `chr(92)` 拼出来，别指望字符串里连写两个反斜杠。



**D56（B13 开工测量：对外那一半今天到底有什么；`.build/probe_b13/`、`.build/b13_probe.ps1`）**

0. **本格 CURRENT_BATCH 的第 0 步前提是错的，先记账**：`tests/` 下**不是**"全是 `Type=Exe`" ——
   `tests\test_activex_dll\` 里躺着两份 `Type=DLL` 工程（`test_activex_dll.vbp` / `test_event_dll.vbp`，
   5 个类、三段式 `Class=Name; file.cls; {CLSID}` 带显式 CLSID），**今天真编译 25 秒就产出
   259–269 KB 的 `.dll` + `TestAXDLL.tlb` + `activex_dll.def`（导出 DllGetClassObject / DllCanUnloadNow /
   DllRegisterServer / DllUnregisterServer / DllMain）**。它们的错处是**从没登记进 `run_tests.ps1`**
   （grep "activex" 零命中）⇒ 整条 ActiveX DLL 管线**在门内一次都没跑过**，此前每一句"对外还差什么"
   都是没测过的话。⇒ 本批第 0 步从"造工程"改成"**把现成工程接进回归** + 补一份 CoClass 块在 DLL 里的工程"。
1. **dll_entry 的唯一读数通道是"真编译 + 读回临时目录"**：生成的 C 只在 `--keep-for-debug` 时留下
   （stderr 打 `intermediates kept at: <dir>`），而 **`--emit-c` 不含 `dll_entry.c`**（实测 0 命中）。
   `g_vb6_coclasses[]` 每条 = `progId / clsidStr / classVariable / factoryFunc=vb6_cls_<C>_New /
   destroyFunc / methodCount + IDispatch 成员表 / ifaceCount + ifaceIids / defaultIfaceIid /
   sourceIfaceIid / events`。`test_activex_dll` 里 Calc 的四条公有成员都进了 disp 表
   （`L"SetValue", 1, 1`），`clsidStr` **就是 vbp 三段式那一枚** ⇒ CLSID 这一位在"显式写在 vbp 上"时两通道一致。
2. **RTL 侧的 IUnknown 三件套早就有真的**（本格按"要从零写"排期，是错的）：`vb6comserver_factory.c`
   的 `CF_CreateInstance` → `vb6_ComObject_Create` → QI；`vb6comserver_obj.c:24-72` 的
   `ComObj_QueryInterface` 认 IUnknown/IDispatch（返回 `self` = **规范指针**，符合 COM）、认
   `ifaceIids[i]`、认 `defaultIfaceIid`、按需挂 `IConnectionPointContainer` / `IProvideClassInfo2`，
   全不认才 `E_NOINTERFACE`；`AddRef/Release` 是 `InterlockedIncrement(&self->refCount)` + 全局
   `g_vb6_cRef`，**归零时调 `desc->destroyFunc` ⇒ `Class_Terminate` 在 DLL 侧有触发点**（D41 那条
   "没有触发点"只成立于 EXE，本批实测划清）。⇒ **"真实现 IUnknown 三件套"不是 B13 的活**。
3. **两套 mint 的分叉现在看得见，而且比 D46-2 说的更近**：同一份 `tests\cc_dll\CoDll.vbp`
   （新式接口 `IProbe` + `Implements IProbe` 的实现类 + `CoClass PG` 块）一次编译里 `IProbe` 拿到
   **两枚不同的 IID** —— stage 2.7（`coclass_identity.cpp` → `coclassIds_`）给
   `IID={F5CEF988-…} (minted)`、`ProgID=CoDll.PG (minted)`；dll_entry（`cgen_util_dllentry_prelude.inc`
   那两枚 lambda）给 `IID_vb6iface_IProbe = {0AD9CBC7-…}`、注册 ProgID `CoDll.CImpl`。并且
   **`coclassIds_` 的后端消费者为 0**（全仓 grep 只有 `driver_interface.cpp` 与 `driver_classchain.cpp`
   读它：校验、Pass G、VB3020 文案）⇒ **产品注册的从来不是 C02 求解出来的那一份**，
   D47 那句"唯一身份出口"目前只覆盖语义层，没覆盖产物。
4. **块名在对外这一侧完全不存在**：表按**类模块**逐条发（`classVariable="CImpl"`、`progId="CoDll.CImpl"`），
   `CoDll.PG` 在产物文本里 **0 次** ⇒ 组名 ProgID 进不了注册表，外部 `CreateObject("CoDll.PG")` 必失败；
   而组内那半（C05）认的正是 `CoDll.PG` ⇒ **两半不对称**。
5. **只实现新式接口的类对外不可调用**：`cc_dll` 那条 `methodCount=0 / methods=NULL` —— 新式契约的实现
   成员按 VB6 惯例写 `Private`（`p01`/`itf_xmod` 皆然），disp 表只收公有成员 ⇒ 外部 IDispatch 客户
   一个方法都点不到；而 `ifaceCount=1` 只把枚 IID 记进表、QI 命中后仍返回**同一个 IDispatch 指针**
   （`self`）⇒ 今天"新式接口对外"= 伪装成 dispinterface。这条要先定口径再动手（接口槽发 disp id？
   还是明确"新式接口不出 DLL"并给诊断？），不许顺手做。
6. **`comCreatable` / `legacyFolded` 在产品侧的读数（原测量点 ④）**：`[ComCreatable(True)]` 在 DLL 工程
   合法（EXE 才 `VB3033`），但两者后端消费者为 0 ⇒ 表里一条不少、注册一条不漏，
   `VB_Creatable = False` 那 9 条与 True 目前对产品**没有任何可观察差别**。
7. ⇒ **B13 按实测重切成三格**：**B13a（本批已出）= 观测面**（新助手 `Test-VbpDll` + 三条用例进门 +
   新工程 `tests/cc_dll`，零编译器代码改动）；**B13b = 把身份出口落到产物**（dll_entry 只从
   `coclassIds_` 取值、删第二套 lambda、块名 ProgID 上表、`comCreatable`/`legacyFolded` 给读数；
   `cc_dll_identity_two_channels` 就是这条的对照，现在钉分叉、合并后必须翻面）；
   **B13c = 新式接口对外可调用的口径 + 规范 IUnknown**（`vb6_iunk_<C>_<I>` 认 `IID_IUnknown`
   现返回**本接口薄指针**，D22-7④ 要在这里收）。
8. **一条通用教训**：**别把"没有用例"读成"没有实现"** —— 这次两个方向都错了：RTL 三件套比计划书假设的
   完整得多，而两份现成工程因为从未登记被当成"DLL 路径不存在"。测量阶段先花 25 秒真编译一次，
   比读码推断便宜得多，也比它可靠得多。


**D57（B13b 实施：dll_entry 的身份改读唯一出口）**

1. **接线照既有先例，不开新机制**：`CCodeGen` 上三个 setter（`setInterfaceRegistry` /
   `setViaRegistry` / `setClassChainRegistry`，`cgen_api.inc:227-232` + `cgen_state.inc` 成员 +
   `driver_codegen_module_loop.inc:75-77` 注入）已经是同一条路，本批加第四个
   `setCoClassIdentityRegistry(&coclassIds_)`。位置选在 `driver_codegen_dll_typelib.inc`
   （`dllCgen` 那一段）而**不是** `driver_codegen_module_loop.inc`，因为这张表只由
   `generateDllEntry` 消费。
2. **只接手写块，折算记录一律照旧**（`id.legacyFolded` 即跳过）。理由与 C04 同源而且是第二次生效：
   折算记录在没有 vbp 三段式时其 `clsid` 是 C02 mint，与 legacy `generateClsid` **种子不同** ⇒
   覆盖上去等于把"改码前编得过、注册表里躺着的存量 DLL 工程"的身份换掉 = 破逐字节护栏。
   副产物：**`legacyFolded` 第一次有产品级读数** —— 有它 = 不覆盖、且不会多出组名那一行
   （折算块名恒等于类模块名）。
3. **`[ComCreatable(True)]` 第一次有后果**：组名档 ProgID（`<工程>.<块名>`）**只在该位为真时发出**，
   表里就多一行 —— 同 CLSID、同类工厂、同 destroyFunc，只差 ProgID 文本，
   `g_vb6_coclassCount` 1→2。类模块名那一档**保留不删**：存量 DLL 客户照旧
   `CreateObject("<工程>.<类>")`，这是"两半不对称"的收法而不是替换。
   反过来说，一个手写块若不写 `[ComCreatable(True)]`，它的组名对外仍然不存在 —— 这是刻意的
   默认（对外可创建要显式表态），也给了 D56-6 那条"9 条 False 与 True 无差别"一个可解释的答案：
   差别现在存在于**手写块**这一侧，存量 attribute 那一侧继续无副作用（护栏）。
4. **IID 并掉的准确范围**：`iidMap` 建表时，接口名 == 该块 `[Default]` 接口 → 用出口值，
   于是 `IID_vb6iface_<I>` 与 `IID_vb6def_<C>` 对同一个接口第一次给同一个 GUID
   （cc_dll 实测两枚都成 `{F5CEF988-…}`，legacy 那枚 `0AD9CBC7` 从产物里消失）。
   **非默认接口仍走 legacy derive** —— 它要和 `cgen_iface_vtbl.cpp:84` 的 `ivDeriveIid`
   （key `"iviface:"+名`，同族不同 key）一起收，那是**第三通道**，归 B13c，别在这批半接。
5. **本批新量到的第三枚 mint 在类型库侧**（D56 没记，因为它不在 dll_entry 里）：
   `TypeLibBuilder::generateUuid`（`src/typelib/typelib_builder.cpp:24`），
   `driver_codegen_dll_typelib.inc:263-269` 的优先级是 `comClsidStr > classClsidMap_ > generateUuid`，
   **不读 `coclassIds_`**；且 typelib 构建会**回写** `comClsidStr` / `comDefaultIfaceIid` /
   `comSourceIfaceIid` 给 dll_entry 读（`generateDllEntry` 刻意排在 typelib 之后）。
    ⇒ 表内自洽了，但"表 vs `.tlb` 对同一个 coclass 是否同值"**本批未证**（需要能读 `.tlb` 的工具，
   `OleView`/`tlbimp` 不在依赖里）。登记给 B13c 顺手量 + 决定是否把 typelib 也接到出口上；
   026 七-1 说的"合并成一处"到今天为止是**两步里的第一步**。
6. **为什么 EXE 侧不动**（这是本批最重要的一条边界）：`com_entry.c` 由**同一个** `generateDllEntry`
   以 `includeDllExports=false` 生成，把注册表也喂给它 = 让每个带手写块的 EXE 工程换身份。
   所以 `comEntryCgen` 上刻意**不调** setter，EXE 继续整跑 legacy。护栏按构造成立，另加实测：
   `tests\cc_act\Act.vbp`（EXE、**带两个手写块**，是最危险的形状）BASE vs NEW 全产物 md5 对照
   （含 `com_entry.c`）。
7. **验收**：`cc_dll_identity_single_source`（原 `cc_dll_identity_two_channels` 翻面）断
   needle `"CoDll.PG"` + `const int g_vb6_coclassCount = 2;` + `0xF5CEF988`、absent `0AD9CBC7`；
   A/B = `.build/pre_b13b_C3.exe` 跑同一份工程，`CoDll.PG` 0 次、`0AD9CBC7` 1 次（缺陷在 BASE 侧复现）；
   四条 `Test-VbpDll` 全绿（含 x86）；16 件逐字节护栏全同；`.tlb` 与表的一致性**未**由本批证明。
8. **harness 坑（本轮自己撞的两条）**：① `src/` 下的 `.hpp/.inc/.cpp` 是 **CRLF**，而
   `ai/022` 与 `tests/run_tests.ps1` 的 **blob 是 LF**（工作树那份因 `core.autocrlf=true` 是 CRLF）
   ⇒ 任何"插入后断言行尾"的检查必须**按文件各自实测**，不能全仓一套；插入文本要先探测目标文件的
   `nl` 再拼（本批 `b13b_flip_case.py` 就是这么过的）。② Python 里
   `("A" + nl` 换行 `"B" + nl)` 是**语法错**（名字与字符串并置），必须写成 `("A" + nl + "B" + nl)`；
   症状是一句 `SyntaxError: Is this intended to be part of the string?`，跟引号无关，别去找引号。

**D58（B13c 读数与三通道合并：`IProbe` 在一次编译里曾有四枚 GUID）**

1. **先解决"没有读数工具"这件事**：仓里没有能读 `.tlb` 的东西（`OleView`/`tlbimp` 不在依赖里），
   所以自己写了 `tests\tools\tlbprobe.cpp`（`LoadTypeLib` + `GetTypeAttr` +
   `GetRefTypeOfImplType`/`GetImplTypeFlags`，约 150 行，打 LIBID / 每个类型的 GUID /
   coclass 引用了哪些接口且哪个是 DEFAULT / 方法表），并把"三通道同值"做成 gated 用例
   `Test-TlbIdentitySingleSource`（helper 单独放 `tests\tlb_identity.ps1`，`run_tests.ps1`
   里只多一行 dot-source + 一条登记）。D57-5 那条"至今未证"现在有证据了：**证的是它不成立**。
2. **BASE（`473784ed…`）实测 `tests\cc_dll`：一次编译里 `IProbe` 有四枚 GUID**
   - 语义层 = COM 服务器表 `{F5CEF988-…}`（B13b 已并）；
   - 类自己的 vtable QI `vb6_iv_iid_IProbe` = `{74B1BA3B-F40C-A0FE-9ED6-3795C52247D1}`
     （`ivDeriveIid`，key `"iviface:"+名`）；
   - 类型库 `dispinterface _IProbe` = `{1E34D82B-…}`（`generateUuid("_IProbe")`）；
   - 类型库 `coclass IProbe` = `{28B6B94B-…}`（`generateUuid("IProbe")`，接口模块被当类登记）。
   要紧的一条不是"值多"，而是**类型库广告给客户端的那枚，服务器根本不应答**：`dll_entry.c` 的
   `IID_vb6def_CImpl` 是 `F5CEF988`，而 `.tlb` 里 coclass `CImpl` 的 DEFAULT 引用是
   `_CImpl = {7CA8CD81-…}` ⇒ 早绑定客户端照库里的 IID 去 QI 必然 `E_NOINTERFACE`。
   这就是 D57-5 未证条目的真实答案，也是本格第 1 条的全部动机。
3. **合并的形式不是"再开一套 mint"，而是把接口自己的 IID 也搬进唯一出口**：
   `resolveIfaceIid(<Proj>, IfaceView)`（`[InterfaceId]` > `mintGuid(ifaceSeed)`，与
   `resolveCoClassIdentity` 里默认接口那条**共用同一个函数**，结构上不可能分叉）
   + Pass E 建 `Driver::ifaceIds_`（key = 接口名小写，源码序打印信息行）。三个消费者一律只读这张表：
   ① `ivDeriveIid`（vtable QI，命中即 `[InterfaceId]`/mint 都由表给）、② `iidMap`
   （dll_entry 的**非默认**接口，本格把 D57-4 欠的那条收掉）、③ 类型库（`_<接口>` 的 GUID
   + coclass 的 DEFAULT 引用）。**表里没有 = 这个接口不是新式接口** ⇒ 各自 legacy 派生照旧，
   存量工程逐字节不变（C04 折算隔离第三次生效，这次连"新式接口但没有块"也覆盖到了：
   `itf_xmod`/`cc_id` 两个工程没有 CoClass 块，它们的 iv 行照变）。
4. **类型库那半边还有一条结构修正**：coclass 的默认接口今天写死成 `<_类名>`（类自己那堆公有成员
   的 dispinterface），块里 `[Default] Interface IProbe` 是不算数的。本格把它接上：块的 `[Default]`
   指的若是新式接口，coclass 的 DEFAULT 引用改成 `_<接口>`。这里踩到一个登记顺序问题 ——
   符号表遍历序由 `unordered_map` 决定，`_IProbe` 可能排在 `CImpl` 之后，先引用后注册就报
   `TYPE_E_ELEMENTNOTFOUND`；解法是**只把需要重定向的那几行延后**到整轮之后登记，其余仍在原地登记。
   第一版无差别延后所有 coclass 时，`tests\test_activex_dll` 的 `TestAXDLL.tlb` 变了
   （3976→3976 字节，内容同、类型序变），是实测把它否决掉的 ⇒ 存量 `.tlb` 零变化。
5. **本格第 2 条拍板：走 (b)，新式接口的成员不发成 disp id。** 三条理由叠在一起：
   ① 契约成员按 VB6 惯例是 `Private`，塞进 IDispatch 表 = 换语义；② `[Default]` 指向新式接口的
   **声明含义**就是"这个类的门面是那个接口"，客户端因此看不见类的公有成员是这句话的应有之义；
   ③ (a) 其实什么都没解决 —— 类型库里的方法要有 dispid 才对客户端可见，而服务器的
   `GetIDsOfNames` 仍然只认公有成员，两边会假绿。所以今天对外**可调用**的是：类的公有成员
   （`_CImpl` 那一档）+ `[ComCreatable(True)]` 换来的组名 ProgID；新式接口要对外可达得走
   **真接口**（类型库 `TKIND_INTERFACE` + 契约成员进库 + QI 过去），那是 B15 类型库线的活，
   本格不越界。**给用户的读数**：编译时打一条
   `C3: CoClass 'CImpl' default interface 'IProbe' is a vtable interface: 0 Public member exported for IDispatch clients`
   （note 级诊断在成功的编译里根本打印不出来，所以仍走 stderr 的 `C3:` 信息行 —— 该通道第四次被用）。
6. **另一条只有读数能说清的旧账**：`.tlb` 里 `_IProbe` 的 `cFuncs` 是 **0** —— 新式接口模块写的
   `Sub Ping` / `Property Get Got` 从来没进类型库（typelib 那段按"类模块的公有成员"收集，
   接口模块的槽在 `IfaceRegistry` 里，两边从来没接）。这条连同第 5 条的"真接口"一起登记给 B15。
7. **护栏（BASE `473784ed…`）**：① `-Category syntax` **118/0**；② 16 件 `--emit-c` 逐字节护栏
   **升级成分类护栏** `.build/b13c_guard.py`：只允许
   `static const unsigned char vb6_iv_iid_<I>[16] = {…}` 那些行变（两个用新式接口的工程各 15 行），
   其余任何字节变化当场红 → **16/16 OK**；③ 全产物 A/B `.build/b13c_ab_all.py` 三个工程一次量：
   `cc_act`（EXE、带两个手写块）只有 5 个 `.h` 的 iv 行变、`com_entry.c` 等其余全同
   ⇒ D57-6 的"EXE 侧刻意不接"到今天还成立；`test_activex_dll`（存量 DLL）`.tlb` 与所有
   .c/.h/.def 全同（`.rc` 里只有 TYPELIB 的绝对临时路径天然不同）⇒ 折算隔离成立；
   `cc_dll`（手写块 + 新式接口）iv 行 + `.tlb` 变 = 本格的靶子。④ 新助手能红：
   假 CLSID 的负控 `FAIL`、正的 `PASS`。
8. **harness 坑（本轮两条）**：① `scripts\env.ps1` 是 **ANSI/GBK 且首行带 shebang**，
   `pwsh -Command ". scripts\env.ps1"` 会读成乱码并执行失败 ⇒ 自测脚本要 MSVC 环境就照
   run_tests.ps1 自己的办法 `cmd /c "call vcvarsall.bat x64 && set"` 灌进进程，别去 dot-source 它；
   ② 用 heredoc 往 python 里塞带反斜杠的 Windows 路径，`\2019`、`\M` 会被吃掉
   （`D:\Program Files (x86)\Microsoft Visual Studio\2019` 变成
   `…Visual StudioMicrosoft Visual Studio`），症状是一句"系统找不到指定的路径"——
   同一条老规矩第 N 次：反斜杠文本一律走 Write 工具。

**D59（B13d 规范 IUnknown + 读出来的一条危险：胖应答瘦的 IID）**

1. **收掉的东西**：新式接口薄指针的 `QueryInterface` 以前把 `IID_IUnknown` 与**本接口**的 IID
   并成一个分支、回 `*ppv = self`（BASE 实测 `tests\itf_xmod\XWriter.vbp`：
   `vb6_iunk_CWriter_IWriter_QueryInterface` 与 `..._ILog_QueryInterface` 各有
   `if (vb6_IidEqual(riid, vb6_iv_iid_IUnknown) || vb6_IidEqual(riid, vb6_iv_iid_<自己>)) { *ppv = self; … }`）。
   `__iv_IWriter` 与 `__iv_ILog` 是同一个结构体里两个不同偏移的成员 ⇒ 同一个对象从两个接口各问一次
   IUnknown 就是**两个地址**，COM 的"规范 IUnknown 恒等"不成立（D22-7④ 记的就是这个）。
   现在 IUnknown 单独一支，一律回**本类实现序第一个接口**的薄指针：
   `void* canon = &me->__iv_IWriter;  /* 规范指针 (ai/022 B13d): 本类实现的第一个接口 */`。
2. **为什么选"第一个接口"而不是别的落点**：① 零布局改动 —— `__comObj` 是裸回指指针、
   `vb6_cls_<C>` 偏移 0 不是 vtable，把 `me` 当 IUnknown 交出去会让客户对着 `__comObj` 取 vtable
   （= 直接踩飞），而任何 `__iv_<I>` 的地址偏移 0 就是 `vb6_ivtbl_*`，是这批里**唯一**既能当身份、
   又能被再次 QI/AddRef/Release 的合法对象指针；② 不新增字段 ⇒ D19 那条硬约束根本不进入视野，
   存量工程（没有 `__iv_` 字段）连形状都不变；③ "第一个"= `ivImplementedIfaces` 的源码声明序，
   也正是结构体里最靠前的那个 `__iv_`，两个"第一"天然重合，不需要新序规则。
   单接口类拿到的值与改之前**完全相同**（`canon == self`），只有多接口类的身份从 N 个收成 1 个。
3. **本批为什么只能断形状 + 真跑，不能断身份**：整条路径上**没有任何自己生成的调用点**去问
   `IID_IUnknown`（VB 侧没有 `QueryInterface` 面；接口值进 Variant/实参那条是 B06c 未开），
   真正的调用者是外部 COM 客户 ⇒ 端到端的身份断言属 B17。于是这批准据是两条：
   ① 新用例 `itf_canonical_iunknown`（`Test-CanonicalIUnknownShape`，走 B08e-6 那条
   `Invoke-CodegenProj`/`--emit-c` 通道）钉"每个 QI 都有一处规范指针分支（XWriter 期望 2 处）
   **且**旧的 `IID_IUnknown) ||` 合并分支已经不在"；② 真编译真跑的 `itf_xmod_writer`
   （QI1..QI4 + LIFE1..LIFE3 + TOF1..TOF3）保证本接口/兄弟接口两条分支与计数没被改坏。
   负控照例跑过：同一份工程在 BASE 上 `canon` 0 处、合并分支 2 处 ⇒ 用例能红。
4. **顺手读出来的一条危险，本批刻意没动**（这是本批最值钱的部分）：RTL 的
   `ComObj_QueryInterface`（`vb6comserver_obj.c:24-49`）对 `desc->ifaceIids[i]` 与
   `desc->defaultIfaceIid` 一律 `*ppv = self` —— 注释写的是 `// dispinterface: same IDispatch pointer`，
   在 legacy 世界里成立（那些 IID 本来就是 IDispatch 那档的）。但 **B13b/B13c 之后这两个字段里
   装的可能是新式 vtable 接口的 IID**，而 B13c 又让 `.tlb` 把同一个 IID 当 coclass 的默认接口广告出去
   ⇒ 早绑定客户端 QI 成功、拿到的是**胖包装器**指针，然后按接口 vtable 去调第 3 槽 —— 那位置在
   IDispatch 上是 `GetTypeInfoCount`。**调错函数比 `E_NOINTERFACE` 危险得多**（静默踩到别的方法）。
   三条出路，下一批拍板（别默认选）：
   (i) 对外**不发布**新式接口的 IID（表的 `defaultIfaceIid`/`ifaceIids` 滤掉 `iidreg_` 认识的项，
       `.tlb` 的 coclass 默认接口回到 `<_类名>`）= 小、安全、但要把 B13b/B13c 立的
       `cc_dll_identity_single_source` 与 `cc_dll_tlb_matches_table` 两条用例**翻面**；
   (ii) 让包装器**真的**应答新式 IID 并返回 `&me->__iv_<I>` = 要类侧发一张"实例→第 k 个接口指针"
       的映射表给 RTL 用（新增导出，`vb6_CoClassDesc` 加字段）= B16 的正路，也是唯一能让
       早绑定客户真用到接口的那条；
   (iii) 只让 `.tlb` 发布、表不应答 = 现状的反面，同样是客户拿不到，**排除**。
   本批不动它的理由不是偷懒：那是一条产品语义拍板（对外到底暴露不暴露 vtable 接口），而且会把
   上一批刚立的用例翻面 —— 同一批里既改身份口径又改对外发布口径，门就不干净了。
5. **护栏（BASE `759e9f51…`）**：① `-Category syntax` **119/0**（118 + 本批新用例）；
   ② `.build/b13d_guard.py` —— 判定从"只许某类行变"升级成**整段函数体豁免**：14 个不含新式接口的
   工程 `--emit-c` 逐字节全同；`itf_xmod`/`cc_id` 摘掉所有 `vb6_iunk_*_QueryInterface` 函数体后
   剩余行序列全同（QI 体 36→46 行），违例 0；③ `.build/b13d_ab_all.py` 三工程全产物 A/B：
   `cc_act`（EXE、带两个手写块）**一个文件都没变**、存量 `test_activex_dll` 只 `.rc` 那行绝对临时
   路径天然不同、`cc_dll` 只有 `CImpl.c` 的 QI 体变（**`.tlb` 与 `.def` 全同**）= 零未归类变化。
   注意 `cc_id` 一个 QI 体都没有（它声明接口但没人实现）⇒ "改动确实进了产物"这条自检要按**全局**
   行数判，按单工程判会被它误红 —— 本轮就是这么撞了一次。
6. **仍然没并的一条**（写清楚，别当已交付）：跨"包装器 ↔ 薄指针"两个世界的 IUnknown 身份还是两个
   （外部客户拿到 `vb6_ComObject*`，进程内薄指针 QI 拿到 `&me->__iv_<first>`）。把它接成一个 =
   上面 (ii) 那一半，归 B16/B17。

**D60（B13e 读数 + 对外口径拍板：广告的那一枚必须就是应答的那一枚；顺带把 D59-4 那条"危险"证伪）**

1. **本格开工第一枪打的是自己的前提**：D59-4 说"胖包装器应答瘦 IID ⇒ 早绑定客户按接口虚表调第 3 槽
   打到 `GetTypeInfoCount`，静默调错函数"。读码 + `tests\tools\tlbprobe.cpp` 实测**不成立**，两条理由：
   ① `TypeLibBuilder` 只会发 `TKIND_DISPATCH`（`addDispInterface`）与 `TKIND_COCLASS`（`addCoClass`），
   仓里**根本没有 `TKIND_INTERFACE` 那条发码路**（那正是 B15 的活）⇒ 客户从 `.tlb` 学到 `_IProbe`
   的 kind 是 `dispinterface`（本轮读数：`CoDll.tlb` 四个类型 = `_IProbe`/coclass `IProbe`/`_CImpl`/
   coclass `CImpl`，一个 `TKIND_INTERFACE` 都没有），按 IDispatch 走，不会去取接口虚表第 3 槽；
   ② 服务器侧 `ComObj_GetIDsOfNames`/`Invoke` 只查 `desc->methods`（= 类的 Public 成员），
   也从没有一条路把 vtable 接口的方法发出去。⇒ 那条"危险"要成立，客户得**从产物之外**自己知道
   接口虚表布局。D59-4 记的机制是错的，但它指的方向是对的 —— 同一次读码撞见了一条**真的**不同值。
2. **真问题（本批的靶子）**：B13c 把 `.tlb` 里 coclass 的 DEFAULT 引用重定向到块 `[Default]` 那个
   `_IProbe`，而 `_IProbe` 在库里的成员数是 **0**（D58-6 实测 `cFuncs`=0），服务器真正认账的成员面是
   `desc->methods` = 类的公有成员 = `_CImpl` 那一档。于是"广告"与"应答"是两份东西：早绑定客户照库
   编译 ⇒ 一个方法都点不到；晚绑定 `CreateObject` + 公有成员 ⇒ 反倒能用。B13c 那次改动把自己刚立的
   "表与 `.tlb` 逐值一致"从"默认接口 IID 一致"改成了"两边一起指向一枚点不到的 GUID"，一致性看着在，
   可用性掉了。教训：**"两条通道同值"这种判据必须同时钉住"值 + 那一档是谁"**，只比值会被自己骗。
3. **拍板（= D59-4 的 (i) 收窄版）：对外默认视图一律是 `<_类名>`，`[Default]` 只管语言层。**
   三处一起退回，不留中间态：① `cgen_util_dllentry_collect.inc` 不再拿出口的 `iid` 覆写
   `info.defaultIfaceIid`（改由类型库回写的 `comDefaultIfaceIid` 说话，也就是"谁应答谁上表"）；
   ② `cgen_util_dllentry_tables.inc` 删掉 B13b 那条"`ifaceIids` 里 `[Default]` 同名接口改用出口 IID"
   的分支（连带 `CoClassInfo::defaultIfaceName` 这个字段一起没用了）；③ `driver_codegen_dll_typelib.inc`
   删掉 DEFAULT 重定向 + 为它搭的"延后登记"机器（`TlbCoClassRow`/`tlbTypes`/`registerCoClass`）。
   保留的：`identityForClass`（coclass 的 CLSID 仍取唯一出口，与表同值）与 `ifaceIidFromMap`
   （接口模块自己那条 dispinterface 的 GUID 仍读 Pass E 的表）⇒ **甲判据一个字没松**。
   那条 `C3: … is a vtable interface: 0 Public member…` 信息行随之删除 —— 它的前提（库里广告的是
   零成员接口）已经不存在，留着一行只会误导。
4. **为什么没有把 `ifaceIids[]` 里"新式接口"那几项滤掉**（D59-4 的 (i) 全量形态）：滤掉就把甲判据
   从三条通道削成两条，而且表与 vtable 会**当场自相矛盾** —— 类自己的 `vb6_iv_iid_<I>` 还在应答那枚
   GUID，服务器表却装着另一枚。"胖应答瘦"的残余今天仍在一处：客户若已知接口 IID 并对着胖指针 QI，
   拿到的是包装器而非 `&me->__iv_<I>`。正解是 (ii) —— 让包装器真返回薄指针（`vb6_CoClassDesc` 加一张
   "实例→第 k 个接口指针"的映射，即 **B16**），那时这枚 IID 该应答成瘦指针、也就更不该滤。故本批
   刻意只收"广告档"，把这条写进 B16 的前置，不半接。
5. **判据落进回归**：`cc_dll_identity_single_source` 翻面（`IID_vb6def_CImpl` 回到 `0x7CA8CD81` 并
   新增为 needle，`IID_vb6iface_IProbe` 仍是 `0xF5CEF988`，撤掉那条日志断言）；
   `cc_dll_tlb_matches_table`（`tests\tlb_identity.ps1`）从一条升级成两条 —— 甲：表的
   `IID_vb6iface_<I>` == `<C>.h` 的 `vb6_iv_iid_<I>[16]` == 库的 `_<I>`；乙：表的 `IID_vb6def_<C>`
   == 库的 coclass DEFAULT 引用 == 库的 `_<C>` 那一档（再加 CLSID 与 `clsidStr` 对一次）。
   负控照例跑在 BASE 上：同一份助手喂 `pre_b13e_C3.exe` ⇒ 乙 判红（DEFAULT `{F5CEF988-…}` vs
   `_CImpl` `{7CA8CD81-…}`），喂本轮 exe ⇒ 全绿。D59-4 那句"要把两条用例翻面"也确实兑现了。
6. **护栏（BASE `905c9392cede13ab9d915f05b091ed3c`）**：① `-Category syntax` **119/0**；
   ② `.build/b13e_ab_all.py` 全产物 A/B，四工程一次量（比 B13d 多带 `test_event_dll`，因为
   `addCoClass` 从 lambda 搬回原地、事件源那条参数表达式被重写）：`cc_act`(EXE) **0 个文件变化**、
   两枚存量 DLL 连 `.tlb` 都**逐字节不变**（只 `.rc` 那行绝对临时路径天然不同）、`cc_dll` 只许
   `dll_entry.c` 的 `IID_vb6def_*` 那 1 行 + `.tlb` ⇒ 零未归类变化，且"改动确实进了产物"按全局
   行数判（`def_lines>=1`）；③ `-Category vbp` 见状态头门读数。
7. **一条 harness 坑**：PowerShell 的双引号串里写 `"…_$Var: …"` 会被解析成**作用/驱动器限定符**
   （`Variable reference is not valid`），整个测试文件 ParserError —— 消息文本要紧跟变量名时一律
   写 `${Var}`。本轮踩在 `tests\tlb_identity.ps1` 的一条中文消息上，被 `-Category syntax` 之外的
   dot-source 自测挡下，没进门。


**D61（B14 三件测量 + 范围重裁：DLL 那条对外管线第一次被真客户端走通）**

1. **测量① —— 胖包装器的 IDispatch 面（读码 + 真跑两头都量）**：
   `ComObj_QueryInterface`（`src\rtl\core\vb6comserver\vb6comserver_obj.c:24-72`）认 `IID_IUnknown`、
   `IID_IDispatch`、`desc->ifaceIids[i]`、`desc->defaultIfaceIid`，再加 CPC/PCI 两件 —— 前四处一律
   `*ppv = self`。真跑读数（新探针 `tests\tools\disp_probe.c`，进程内、不查注册表）：
   `DllGetClassObject` → `IClassFactory::CreateInstance(IID_IDispatch)` = S_OK；
   `QI(IUnknown)` `same=yes`（规范身份在胖侧成立）；**对 `IID_IProbe` 的 QI 也 S_OK 且 `same=yes`**
   ⇒ D59-4 / D60-4 那条"胖应答瘦"从推理变成实测事实（今天它仍然不可从产出的库到达：库里那一档是
   dispinterface，客户按 IDispatch 用它恰好是对的）。
   `GetTypeInfoCount` 恒 1；`GetTypeInfo(0)` 载入内嵌 `.tlb` 后**按 CLSID** 取 coclass 那份
   （实测 `TYPEINFO0_kind=5` = `TKIND_COCLASS`、`cFuncs`=0 —— coclass 的 TYPEATTR 本来就不带成员，
   成员在它引用的 implType 里，不是缺陷）；`GetTypeInfo(1)` = `DISP_E_BADINDEX`。
   `GetIDsOfNames` **完全不看 `riid`**，只在 `desc->methods` 里按名字 `wcscmp`；`Invoke` 同样不看 `riid`。
   未知名 → `DISP_E_UNKNOWNNAME`(0x80020006)，实测。
   **结论：四件套本身是通的** —— `Add(2,40)=42`、`SetValue(7)` 之后 `GetValue()=7`（跨两次 Invoke
   保住状态 ⇒ 同一个实例，包装器复用那条路也对）。所以 B14 原命题里"没接"的那半**不是**四件套缺失，
   而是**成员面为空**（见测量③）。
2. **测量② —— 薄指针那一侧没有 IDispatch 的位置**：`vb6_ivtbl_<I>` = `{ QI, AddRef, Release, <自有槽…> }`
   （`src\backend\module\cgen_iface_vtbl.cpp:297-330` 的头文件发射 + `:550-556` 的实例表）——
   IUnknown 形，**四件套那四槽不在里面**。给薄指针一张 IDispatch 面 = 在 `Release` 之后插四槽 ⇒
   接口自有成员的槽号全体后移 4。产品内部自洽（调用点全从同一份结构体发码），但这正是"真接口 / dual"
   的布局改动，而且**必须先有 B15 把接口成员发进库**，否则就是重演 D60：库里那一档 0 成员、
   服务器多应答一枚 IID，门内全绿、外面点不通。⇒ **B14 的布局那半排到 B15 之后**，本批刻意不做。
3. **测量③ —— DispId 表只有一处生产者，而且不产接口成员的**：全工程只有
   `src\driver\detail\driver_codegen_dll_typelib.inc:186` 写 `comDispid`，位置就在"类模块 Public 成员"
   那条收集循环里（`refs` 按 `access == AccessLevel::Public` 过滤）；消费者是
   `cgen_util_dllentry_tables.inc:49`（COM 表的 dispid）与类型库的 `<_类名>` dispinterface —— 两边同源
   （M29 那条回写）。**接口契约成员从来没进过这张表**。要发它只有两条路：把契约成员当公有发
   （口径 (b) 已禁止，且 `GetIDsOfNames` 只认公有成员 ⇒ 只会两边假绿）或走真接口（B15）。
4. **本批交付 = 把这条管线从"字节一致"升到"真调用得通"**：新增不查注册表的进程内客户端
   `tests\tools\disp_probe.c`（`LoadLibrary` + `DllGetClassObject` + `CreateInstance(IID_IDispatch)`，
   四件套各问一遍，另可追问第二枚 IID 看应答的是哪份指针）+ 助手 `Test-DispatchInvoke`
   （`tests\disp_invoke.ps1`，与 B13c 的 tlb 探针同一套形态：现编现用、`.obj` 落在产物目录）+ 两条
   gated 用例：`ax_dll_dispatch_invoke`（存量 `TestAXDLL.Calc` 真点通）与
   `cc_dll_dispatch_iface_only`（只满足新式接口的类：成员面**必须**为空 ⇒ `NAMES=ADD hr=0x80020006`，
   同时 `QI_EXTRA ... same=yes` 钉住今天这枚 IID 由胖指针应答）。**新助手先证明能红**：同一份用例喂一个
   不存在的 CLSID ⇒ `GETFACTORY hr=0x80040111 CLASS NOT REGISTERED` → 判红（实测）。
   这一格不改任何编译器源码 ⇒ 产物逐字节不变（BASE 与收线 exe 同一枚 md5），护栏就是这条本身。
5. **范围重裁（拍板）**：B14 收缩成本批这一格；"薄指针的 IDispatch 面 / 接口成员进库 / dual 布局"
   三条合并成 **B15 的前置问题**，顺序定死：**先把接口在库里发成真接口（`TKIND_INTERFACE` + 成员进库），
   再谈让 QI 交回薄指针（B16）**。理由就是 D60 教训的反向应用 —— 任何"让服务器多应答一枚 IID"的改动，
   都必须先满足"库里那一档的成员面 == 服务器应答的成员面"。
6. **读码撞见的一条弱点（登记，本批没动）**：`ComObj_Invoke`（`vb6comserver_obj.c:166-186`）先按
   dispid + invkind 精确配，**配不上就退回"第一个同 dispid 的表项"**。而 VB6 语义里 Property Get/Let/Set
   共享同一 dispid 是常态 ⇒ 一个只发成 Get 的属性，用 `DISPATCH_PROPERTYPUT` 问也会把 Get 那一项调出去。
   收紧（配不上即 `DISP_E_MEMBERNOTFOUND`）会改到存量工程的可观察行为，属口径题不是 bug 修 ⇒ 交下一批拍板。
**D62（B15 四件测量：一张 `TKIND_INTERFACE` 到底要什么、接口成员签名从哪来、位数 flag 有没有后果、今天的形状）**

1. **测量① = 成立条件**（`.build\tlb_iface_exp.cpp` + `.build\b15_measure1.ps1`，逐个 mode 真建真读；
   读回用新写的 `.build\tlb_slots.exe`：它打 oVft/callconv/funckind/返回 vt/参数 vt（含 `VT_PTR` 链内层）。
   `tests\tools\tlbprobe.cpp` 的格式是回归 needle 的来源，刻意没去动它）。七条读数：
   - `TKIND_INTERFACE` + `FUNC_DISPATCH` + `oVft=0`（= 照抄 `addDispInterface` 那套）⇒ **`AddFuncDesc` 就拒**
     （`0x800288BD`）。最少集合是四步：`FUNC_PUREVIRTUAL` + `oVft=(3+槽号)*指针宽` + `SetFuncAndParamNames`
     + `LayOut()`，之后 `SaveAllChanges` 出来的库读回 `kind=interface / FUNC 0 memid=1 name=DoIt invkind=1 nparams=1`。
   - 接口那一档**不吃** `SetTypeFlags(TYPEFLAG_FCANCREATE)`（同一个 `0x800288BD`，但库仍建成）⇒ 可创建位只属于 coclass。
   - `tdesc.vt = VT_I4|VT_PTR` 而 `lptdesc` 留空 ⇒ **建库进程当场崩、一行输出都没有**。ByRef 的正规建法实测可用：
     `VT_PTR` + `lptdesc` 指向内层 `TYPEDESC{VT_I4}`，读回 `vt=VT_PTR-> vt=0x0003`（mode 7）。
   - coclass 可以**引用真接口**并标 DEFAULT（`AddRefTypeInfo` + `AddImplType` + `SetImplTypeFlags`，mode 5 全 S_OK，
     读回 `REF 0 kind=interface … flags=0x1 DEFAULT`）⇒「类实现 IProbe」这一档在库里走得通，B16 不用先解这个结。
   - **0 槽的真接口建得起来**（mode 9：`CreateTypeInfo(INTERFACE)` + `SetGuid` + `LayOut` ⇒ 读回 `cFuncs=0`）
     ⇒ 本批「只修形状与身份、成员面押后」不是被建库逼的，是自选口径。
   - mode 8：把生成的 `vb6_ivtbl_IProbe` 逐字段照抄进库（两槽、`CC_CDECL`、`Got` 直接返回 `VT_I4`）**照样建成并原样读回**
     （`FUNC 0 Ping cdecl oVft=24 ret=VT_VOID` / `FUNC 1 Got invkind=2 cdecl oVft=32 ret=VT_I4`）
     ⇒ `CreateTypeLib2` **不校验 COM 规范**。这条既开门也开坑：想发什么发得出去，发错了没人拦 ⇒ 成员面的门槛
     不在库里，在我们自己的槽不是 canonical COM（见 5）。
2. **测量② = 接口成员签名的来源，并且订正 D61-3 的根因**：`Driver::ifaces_`（`IfaceView::slots`，每槽 `sig`
   指向 `SubDecl`/`FunctionDecl`/`PropertyDecl`，`body` 恒空）就是接口成员的唯一事实面，主 cgen 循环早就在用它
   （`driver_codegen_module_loop.inc:76` `setInterfaceRegistry(&ifaces_)`），而类型库这段在同一个函数体内 ⇒ 直接可读。
   **反面对照**：今天这段沿用的是**类符号**那条收集路（`clsSym.memberNames` + `lookupModuleByKind`），而 026/B02 的口径是
   Interface 块只进 `Module::interfaces`、**不进 `declarations`** ⇒ 接口模块的 Class 符号 `memberNames` 是空的
   ⇒ 实测 `_IProbe` 的 `cFuncs=0`。所以「接口成员点不通」不是 D61-3 说的「DispId 表少一处生产者」，是**成员压根没到发码现场**。
   类型 → `ELEMDESC` 的缺档（`typelib_builder.cpp:51-65` `mapVartype`，switch 只认精确枚举值）：`LongPtr`/`LongLong`/
   `UserDefinedType`/带位组合的 `Long|Array` 全落到 `default: VT_VARIANT`；`Object` 恒 `VT_DISPATCH`（工程类/接口指针
   该是 `VT_UNKNOWN` 或那枚接口的 alias）；ByRef 只写 `PARAMFLAG_FOUT`、**从不建 `VT_PTR` 链**（= 测量①那条崩溃的成因）。
   ⇒ 成员面要发，得连「参数类型保真」一起做，不是加一个 loop 就完事。
3. **测量③ = `CreateTypeLib2(SYS_WIN64)` 有后果，本批不动**：mode 2 与 mode 6 逐调用只差那枚 flag ⇒ `.tlb` **字节不同**
   （md5 `D700E512…` vs `92547F82…`，长度都是 1236），而且同一份 `oVft=24` 在 32 位库里读回 **48** ⇒ 位数声明会换算
   槽偏移，不是元数据摆设。**顺带一条流程教训**：③ 第一次跑出来是「两枚 md5 相同」的**假读数** —— cl 编译失败但旧
   `tlb_iface_exp.exe` 还在原地，脚本拿 `Test-Path exe` 当构建成功判据，于是把上一版二进制又跑了一遍。已改成先
   `Remove-Item` 旧 exe/tlb、再拿 cl 的 stdout 里有没有 `error C` 判成败（`.build\b15_measure8.ps1`）。
   ⇒ 改 flag = 全部存量 `.tlb` 重排 + oVft 语义变 ⇒ 登记为 **B16 的前置题**（真接口的 oVft 要按 `--arch` 给对，
   先得定 flag 怎么传）。另有一条今天的实测：同一工程 `--arch x86` 与 x64 的 `.tlb` **逐字节相同**
   （md5 `B0AB5567…`）—— 因为库里目前没有任何随指针宽度变的东西（`oVft=0`、无 `VT_PTR` 链）。
4. **测量④ = 今天的形状与根因**（`.build\b15_measure34.ps1` + `tlb_slots` 读 `CoDll.tlb`）：
   `TYPE 0 kind=dispinterface name=_IProbe guid={F5CEF988-…}`（Pass E 那枚，`cFuncs=0`）+
   `TYPE 1 kind=coclass name=IProbe guid={28B6B94B-B7A4-…}`（**第五枚 mint**：`TypeLibBuilder::generateUuid(clsSym.name)`，
   不在 COM 表、不在 `coclassIds_`、不在任何 QI 应答面）引用前者并标 DEFAULT。
   ⇒ 客户端导入后看到的「接口」= 一枚假 CLSID + 一张空 dispinterface + 一句「可创建」。
   根因一句话：`driver_codegen_dll_typelib.inc:84` 那条循环是**唯一没有 `isInterface` 过滤**的 COM 消费者
   （对照 `cgen_util_dllentry_collect.inc:26`、`driver_codegen_dll_sync.inc:8,35` 三处都有），而 stage 3.6
   （`driver_compile.cpp:623-641`）把被 `Implements` 点名的类一律标 `isInterface=true`（新式接口走的也是这条）
   ⇒ 接口模块在 COM 表里不存在、却在库里占一整行 coclass。
5. **范围拍板（按读数收窄，宁窄勿滥）**：B15 = **形状与身份那一半**：
   ① 新式接口宿主（判据与 B13c 同一条：`ifaceIds_` 命中 **且** `ifaces_` 有同名键）在库里发成 `TKIND_INTERFACE`，
   GUID 用 Pass E 那枚，行名用接口自己的名字（`IProbe`，不再 `_<名>`）；② 撤掉那枚假 coclass 行与它的 `generateUuid`
   mint；③ **成员面本批不发**，理由用实测表述：mode 8 证明建库什么都收 ⇒ 门槛在我们自己的槽不是 canonical COM
   （`CC_CDECL` + 原生返回值 + `vb6_ivtbl_<I>` 里没有 IDispatch 前缀 = D61-2），发成员必须先定 x86 calling convention
   与 oVft/位数 flag 口径（测量③），那是布局批（B16）的活，且按老规矩布局改动要 x86 真编译真跑，不与本批混；
   ④ 判据改点不放宽：`Test-TlbIdentitySingleSource` 通道 3 从「`kind=dispinterface name=_<接口>`」换成
   「`kind=interface name=<接口>`」，另加两条反面断言（库里不得再出现 `kind=coclass name=<接口>`、不得再出现
   `kind=dispinterface name=_<接口>`）⇒ 回退即红；⑤ 存量护栏：非新式接口宿主的每一档（`_CImpl`、coclass `CImpl`、
   stock `VBMANLIB`/`EventCalc`…）逐字节不变。

**D63（B16 三件测量 + 实施 + 判据与护栏：x86 调用约定、canonical 化的波及面、位数 flag/oVft 的裁决与迁移口径）**

1. **测量① = x86 调用约定今天不成立**：`--arch x86 --emit-c` 的产物里**一个约定修饰都没有** ——
   `vb6_ivtbl_IProbe` 的 IUnknown 前缀与契约槽全是裸声明（`long (*QueryInterface)(void* self, …);`
   = `__cdecl`），`grep -c __stdcall` = **BASE 0 / NEW 21**（`cc_dll`）。x64 是单一 ABI 看不出来，
   所以这条只能用 x86 产物说话。COM 规范要 `CC_STDCALL` ⇒ 整条薄面（前缀三槽 + 契约槽 + 它们的
   实现函数）一律加 `__stdcall`；x64 下 MSVC 接受并忽略该修饰（同 `cgen_delegate`/`cgen_com_events`
   先例）⇒ 一份生产码两架构通用。**差的另一样（canonical 返回形状）本批刻意不补**，理由见 5。
2. **测量② = canonical 化的波及面（分类护栏逐条逼出来的，不是猜的）**：三层，全部落进白名单：
   ① `__stdcall` 出现在"声明了/实现了新式接口"的工程的槽表与实现函数上（`cc_act`/`cc_dll`/`cc_id`/
   `itf_xmod` 的 `.h` 与 `.c`）；② 薄面出入口 `vb6_iv_thin_<C>`/`vb6_iv_claim_<C>` **只出现在有接口宿主**
   （类 `Implements` 了新式接口）的工程（`cc_act` 的 `com_entry.c` 6 行、`itf_xmod` 的 `CWriter.c` 15 行、
   `cc_dll` 的 `CImpl.c` 14 行 + `CImpl.h` 2 行）；③ COM 服务器表（`dll_entry.c`/`com_entry.c`）
   **每个 coclass 行**多两个字段 `ifaceThinPtr`/`instanceClaimRelease` —— 新式宿主填函数名、存量/无接口
   填 `NULL`（`ax_dll` 4 个 coclass 共 8 行、`ax_event` 5 个共 10 行）。**`cc_id` 只声明接口**
   （`Interface` + CoClass 折算）、没有宿主 ⇒ 产物里只有 `__stdcall`、没有薄面那一对。
3. **测量③ = 位数 flag 的裁决与迁移口径**：**flag 跟 `--arch` 走**（`CreateTypeLib2(is64_ ? SYS_WIN64 :
   SYS_WIN32)`）。理由：不发成员时不疼，一发成员 oVft 就是"客户端拿到的槽偏移"，恒 64 位会让 x86 客户端
   读到翻倍的偏移（D62-3 已实测换算）。**代价的实测边界**（这决定判据④怎么写）：默认架构（x64）下存量
   工程 `.tlb` **逐字节不变**（`ax_dll`/`ax_event` 在全产物 A/B 里无差异；单工程探针 md5 同 `87d34673`、
   diffbytes=0）；`--arch x86` 的存量 DLL 同尺寸（3976 字节）、**32 个字节不同** —— 首处 @20 是头部的
   位数声明字（x64 `43 00 00 00` / x86 `41 00 00 00`），其余 31 处是布局字（oVft 一族按 4/8 换算），
   且**两架构出的库恰好在同样这 32 处不同**（把 flag 的全部后果关在这一处），x64/x86 两枚读端读回的值
   **逐项相同** ⇒ 这是"声明位数如实"的必然修正，不是内容变化。判据④的口径因此写成
   "**默认架构一字不动 + x86 只有这 32 个字节的位数布局字**"。
4. **实施（按读数收窄，一次一个改动）**：`cgen_iface_vtbl.cpp` = 薄面整条 `__stdcall` + 发射
   `vb6_iv_thin_<C>`（按 IID 交薄指针）/`vb6_iv_claim_<C>`（包装器放手时退掉底座引用）；
   `cgen_util_dllentry_collect/prelude/tables.inc` = 服务器表行加 `ifaceThinPtr`/`instanceClaimRelease`；
   `vb6comserver.{h,c}` + `vb6rtl_class_com.h` = 包装器与类工厂按这两个字段接线（**接口 IID 的 `QI` 从
   "交胖指针"改成"交薄指针"**）；`typelib_builder.{hpp,cpp}` = `addVtableInterface` 改发**如实契约**
   （每成员 `FUNC_PUREVIRTUAL` + `oVft=(3+槽)*指针宽` + `CC_STDCALL` + 原生返回 vt，ByRef 建 `VT_PTR` 链），
   `CreateTypeLib2` 的 flag 跟 `is64_`；`driver_codegen_dll_typelib.inc` = 接口宿主分叉把
   `IfaceView::slots`（D62-2 认定的唯一事实面）逐槽喂进去。**零布局改动**（D19 不入视野），`__refcount` 头照旧。
5. **范围拍板（断不了的行为先断形状，写清哪半属 B17）**：① **canonical COM 的返回形状不做** —— 现在是
   "原生返回类型 + `__stdcall`"（库读数 `ret=0x0003` = `VT_I4` 直出，不是 `VT_HRESULT` + `[out, retval]`）；
   做它要改每个槽的签名与全工程调用点，且**是否必要取决于真早绑定客户能不能用**（B17 测量②）；
   ② **跨世界身份仍是两个值**：同一次 `QueryInterface` 问 `IID_IUnknown` 拿到包装器指针、问接口 IID 拿到
   薄指针（探针 `QI_IUNKNOWN … same=yes` 与 `QI_EXTRA … same=no` 并排就是这条读数）—— 合一属 B17；
   ③ 薄指针本身**可交付**：前 3 槽就是 IUnknown 三件套，客户端 `IUnknown_Release((IUnknown*)thin)` 是
   规范姿势（本批探针就这么收尾的）。
6. **判据（五条都对上读数）**：① `cFuncs=2` + 每成员的 `oVft`/`callconv`/`funckind`/返回 vt/参数 vt
   都有读数 —— 新探针 `tests\tools\tlb_slots.cpp`（自 `.build` 升格入库，与 `tlbprobe.cpp` 刻意分开：
   后者的打印格式是 identity 判据的 needle 来源）+ 助手 `tests\tlb_contract.ps1` + 两条 gated 用例
   `cc_dll_tlb_contract_x64/x86`（**x86 那条必须用 x86 读端**：同一份库 x64 读端会把 12 读成 24）。
   ② 真跑：`tests\tools\disp_probe.c` 长一条"按库里形状直调 vtable 槽"的支路 —— `CREATE_IFACE` 拿薄指针、
   `VTBL_GET_BEFORE=0` → 经槽 3 写 21 → `VTBL_GET_AFTER=42`（x64 与 **x86** 各一条 gated 用例，
   x86 那条是"布局改动只能靠真跑收"的那一半）。③ 「广告 == 应答」：`QI_EXTRA` 从 B14 钉的 `same=yes`
   **故意翻成 `same=no`**（= 交的是薄指针、不再是 IDispatch 包装器），且薄指针的槽 3/4 与库里那一档的
   `oVft=24/32` 一一对应（同一份形状）。④ 逐字节：`--emit-c` 17 件 `violations=0/17`
   （`cc_act`/`cc_dll`/`cc_id`/`itf_xmod` 四件按白名单分类，其余全同）+ 全产物 A/B `unclassified=0`
   （只有 `cc_dll` 的 `.tlb` 1484→1624 字节的**预期变化**，存量工程默认架构的 `.tlb` 与全部 `.def`/`.rc`
   一字不动）。⑤ 见 5。
7. **负控与"用例真能红"**：`.build\b16_negctl.ps1` 拿收线前的 `pre_b16_C3.exe` 跑两条契约用例 ⇒
   `pass=0 fail=2`、四条 needle 全 miss、且反面断言 `cFuncs=0` 命中；`.build\b16_cases.ps1` 拿新 exe 跑
   五条（契约 x64/x86 + 调度 x64/x86 + identity）= `pass=5 fail=0`。**注意 `.build` 里那两个 harness 是
   临时件（gitignored）**，入库的是 `tests\tlb_contract.ps1` + `tests\tools\tlb_slots.cpp` +
   `tests\tools\disp_probe.c` 的修改。
8. **流程教训（三条本批新踩的）**：① **套件在 CI 上用 pwsh 7 跑**（`ci.yml:121 shell: pwsh`），
   `Join-Path $a b c`（4 参 = `-AdditionalChildPath`）是 PS7 才有的形式，PS 5.1 下**静默给空串**
   ⇒ 本地复现套件必须用 `pwsh`，否则 x86 助手"建不出来"是假象（第一次跑 harness 就这么误红过）。
   ② **套件的 `$env:LIB` 是 x64 在前**，用 x86 的 `cl` 链接会把 x64 的 `ole32.lib` 挑走
   （实测 `LNK4272` + 10 个未解析外部）⇒ x86 助手/探针构建时临时 `$env:LIB = $msvc.LibX86`（两个助手都这么修）。
   ③ 探针里按槽直调要**先解一层**：薄指针是 `{ vt }`，`((probe_ivtbl*)p)->Got(p)` 是把 `p` 当槽表
   （实测读到对象后面的堆、当场崩）；正确写法 `*(probe_ivtbl**)p`。另：崩溃型探针必须
   `setvbuf(stdout, NULL, _IONBF, 0)`，否则块缓冲把已打印的读数全吞掉（B15 那条 `Test-Path exe` 教训的同族）。

**D64（B17 两件测量 + 两处静默缺陷的修复 + 外部激活判据：真注册、按 CLSID/接口 IID 的外部客户、晚绑定的边界）**

1. **测量① = 注册与外部激活今天通不通（逐级记读数）**：新探针 `tests\tools\com_act_probe.c`
   （自己调 `DllRegisterServer`，再让 COM 自己去找 `InprocServer32` 装载 DLL）跑出来的是
   **"三级都对、一级断"**：
   ① `DllRegisterServer` 写注册表这一级**通** —— `REG_CLSID_DEFAULT=OK CoDll.PG`、
   `REG_INPROC_PATH=OK <dll>`、`REG_THREADING=OK Apartment`、`REG_PROGID_CLSID=OK {…}`，
   CLSID/ProgID 两族键都落到了 HKLM\SOFTWARE\Classes（本机进程是管理员；非管理员会被 Windows
   重定向到 HKCU\Software\Classes，两侧对称所以用例免管理员也能跑）；
   ② `CLSIDFromProgID("CoDll.CImpl")`/`("CoDll.PG")` → 同一个 CLSID，`CoCreateInstance` 拿
   `IID_IDispatch` 得包装器 —— `CreateObject` 那条路**通**；
   ③ **断在类型库那一级**：`LoadTypeLibEx(<dll>)` 读不到内嵌库（`TYPELIB_LOAD=MISSING`），
   于是 `TypeLib` 项根本没写进注册表 —— 外部工具按 LIBID 找不到这份契约。查下去是**两条静默缺陷**：
   - **缺陷 A（link 期）**：`rc.exe` 的发现面太窄。`driver_link.cpp` 只认 `WindowsSdkDir`
     环境变量与 `C:\Program Files (x86)\Windows Kits\10\bin`，而本机 SDK 在 **`D:\Windows Kits\10`**
     ⇒ 找不到 rc.exe ⇒ `activex_dll_typelib.rc` 编不出 .res ⇒ **类型库与 VS_VERSION_INFO 都不嵌**。
     实测：BASE 产出的 DLL 连 `.rsrc` 段都没有（数据目录 2 = `(0,0)`），NEW 为 `(4880)`（ax_dll）/
     `(2504)`（cc_dll）。**旧写法在 CI 上看不出来**（GitHub runner 的 SDK 在默认盘），是本机才亮。
     顺带：旧代码在 bin 目录里取"枚举到的最后一个"版本目录（结果随枚举顺序变），且"找不到 rc.exe"
     只在 `--verbose` 下打印 —— 静默的质量损失，本轮改成始终打到 stderr。
   - **缺陷 B（RTL）**：`vb6_UnregisterTypeLib` 调 `UnRegisterTypeLib` 时**实参顺序写反** ——
     原型是 `(libID, wVerMajor, wVerMinor, lcid, syskind)`，旧写法写的是 `(…, SYS_WIN64, pAttr->lcid)`
     ⇒ 两个键都找错、函数失败，而本函数**无条件返回 S_OK** ⇒ **每次反注册都静默漏掉整棵 TypeLib 键**。
     实测（干净注册→反注册）：`CLSID\…`、`CoDll.PG\CLSID` 都清掉了，`TypeLib\{232C4D5F-…}` 还在。
     修法 = 按原型给对顺序，并取**库自己声明**的 lcid/syskind（中性库 RegisterTypeLib 记在 `0x0409`
     下 —— 实测 `pAttr->lcid = 0x0409`、x64 库 `syskind=3`、x86 库 `syskind=1`；另写一枚
     `.build/b17_tlbattr` 探针直接读 TLIBATTR 得到这三个值，不靠猜）。
2. **测量② = 早绑定那条路的外部读数，顺带回答 canonical 返回形状**：注册过的、由系统装载的 DLL
   上，`CoCreateInstance(CLSID, …, IID_IProbe)` 拿到**薄指针**，客户按库里 `oVft=(3+槽)*指针宽`
   直调契约槽 —— `VTBL_GET_BEFORE=0` → 槽 3 写 21 → `VTBL_GET_AFTER=42`，**x64 与 x86 各一遍**。
   ⇒ 答案：**canonical 返回形状（`HRESULT` + `[out, retval]`）不做** —— 真按库里 oVft 直调的客户
   已经能拿到值；那个形状只有"把接口当 dual 自动化接口用"才需要，与成员面（`GetIDsOfNames`）
   是同一件事，属 v1 边界。跨世界身份合一也不做（外部客户拿到的 `IDispatch` 仍是包装器，
   `IID_IUnknown` 与接口 IID 仍是两个指针值）—— 两条都写进手册。
3. **晚绑定的边界（VB6 语义，不是缺陷）**：契约成员按 VB6 惯例是 `Private`，所以**类的默认面上
   按名点不到它们**（`NAMES=Ping hr=0x80020006 DISP_E_UNKNOWNNAME`，`Got` 同）—— D61-3 已钉过
   这条口径（"契约成员不会变成对外可点的 disp id"）。为了让"外部 `CreateObject` + 按名调用"那一族
   （`test_p613_typelib` 的做法）有东西可点，`tests\cc_dll\CImpl.cls` 加了**一个公有成员**
   `Public Function Twice(ByVal n As Long) As Long`：同一类于是**两个面**并存 —— 默认面（`_CImpl`
   dispinterface，晚绑定按名走这条）与接口面（`IProbe` 的契约槽，早绑定走这条）。库与服务器表
   自动跟上（`_CImpl` 多一条 `FUNC 0 memid=1 name=Twice`，服务器表 `methodCount 0→1` ⇒ B13b 那条
   针跟着从 `0, /* methodCount */` 改成 `1, …`，广告==应答仍然成立）。
4. **实施（两处修复 + 夹具一处增长，零布局改动）**：`driver_link.cpp` 新增 `findRcExe()`
   （环境变量 → Program Files → 盘符扫描 → PATH，目录内取版本号最大者；TypeLib 与版本信息两处共用）
   + 找不到时的 stderr 提示；`vb6comserver.c` 的反注册按原型给对实参；`tests\cc_dll_client\`
   （客户 EXE 工程：`CreateObject` + 晚绑定调用，不引用 DLL）+ `tests\tools\com_act_probe.c`
   + `tests\com_activate.ps1`（两条助手：整趟冒烟 / 注册→跑外部客户→反注册）。**不改发送器**：
   `--emit-c` 17 件逐字节全同。
5. **判据（五条对上读数）**：① 外部那三条验收在新式 CoClass 上真绿 ——
   `cc_dll_external_activate[_x86]`（注册四项 + `PROGID_LOOKUP same=yes` + `COCREATE_DISP` +
   `CALL=Twice result=42` + `COCREATE_IFACE` + `VTBL_GET_AFTER=42` + 三类键 `gone`）与
   `cc_dll_late_client`（**另一个进程**的 C3 客户 EXE：`EXT1:OK`/`EXT2:OK`/`EXT-DONE`），
   且接口那一档的 IID/CLSID/ProgID 与 `.tlb`、服务器表三处同值（身份自 B13b 起单源，本批只是
   加了"外部队列"这一层读数）；② 早绑定按库里 oVft 直调拿到正确返回值，x86 + x64 双验；
   ③ **注册表可清理**：反注册后 CLSID/ProgID/TypeLib 三类键都不留，用例可重复跑（修缺陷 B 之后
   才成立）；④ 存量 DLL 的**注册与激活行为**不变（注册写的是同一张服务器表；`--emit-c` 与
   `.tlb`/`.def`/`.rc` 全部逐字节同，见 6）；⑤ canonical 返回形状与跨世界身份合一按 2 只断形状
   并写清属 v1 边界。
6. **护栏与负控**：`.build/b17_ab_all.py` 三段 —— A) `--emit-c` **17 件逐字节全同**（前端/cgen
   一行没动）；B) 全产物 A/B（双方都设 `WindowsSdkDir`，把这个变量按掉）：`.c/.h/.def/.rc/.tlb`
   全同，**只有 `rtl/vb6comserver.c` 差 10 行、且全在反注册那段**；C) rc.exe 发现面：NEW 带不带
   `WindowsSdkDir` **都嵌资源**（4880/2504），BASE 不带就不嵌（0）⇒ 缺陷 A 的因果链闭合。
   负控 `.build/b17_negctl.ps1`：同一套用例喂收线前的 `pre_b17_C3.exe`（并给它 `WindowsSdkDir`，
   把缺陷 A 按掉、只剩缺陷 B）⇒ **pass=0 fail=2**，红的正是 `CLEAN_TYPELIB=STILL`。
7. **三条教训（下一批直接用）**：① **产品二进制跨构建不可比** —— `link.exe` 每次都写 PE 时间戳，
   同一个 exe 连编两次 md5 都不同（实测 NEW-a/NEW-b），所以"产物字节 A/B"在这里没有区分力；
   要看的是**资源目录在不在**（数据目录 2）与 **emitted 文件**（那两样是确定的）。② **探针的模式词
   别写死在 `argv[5]`** —— `... <progid> reg` 这种少一个参数的写法会静默退化成"整趟跑"（注册完顺手
   反注册），调用方以为已经注册好，下一个进程再激活就 429；已改成从 `argv[4]` 起扫。③ **GUI 子系统
   exe 的未处理运行期错误会弹模态框**（RTL：有控制台走 stderr，没有就 `MessageBoxW`），套件里表现
   为"跑超时被杀"这种最难查的红 ⇒ 客户夹具要 `On Error Resume Next`，把失败落到"FAIL null"这类
   读数上让用例按缺 needle 判红（本机手工跑时真被这个框挡住过一次）。
8. **门抓到的一条既有潜伏缺陷（本批修掉，代码 `732c1f8`）**：第一次门的 vbp job 在**既有**用例
   `ax_dll_dispatch_invoke` 上 exit=**0xC0000374（STATUS_HEAP_CORRUPTION）**，读数正停在第一次
   `Invoke` 前一行（`NAMES=ADD … dispid=2`）。读码定位：`ComObj_Invoke`
   （`vb6comserver_obj.c`）用 `CoTaskMemAlloc` 拿 `coercedArgs`（**不是零**），收尾却对**每个**元素
   调 `VariantClear`；走"原样传下去"那一支（`args[i] = src`，`VT_I4` 等）的元素从没 `VariantInit` 过
   ⇒ `VariantClear` 读到垃圾 `vt`/指针，只要形状像 BSTR/Dispatch 就去 free 一个野地址。
   **机制证明（A/B）**：把该数组临时填成"`vt=VT_BSTR` + 野指针"的编译器所产 DLL，探针**恰好停在同一个
   读数处**（与 CI 现场一致）；清零版整趟干净。本机连跑 12 次不复现（新建堆页恰好为 0 时这条是良性的
   ⇒ 随堆状态偶发），所以它不是本批引入，是本批新增的外部激活动作把它逼出来的 —— 修法 = 分配后整段
   `memset(..., 0, argc * sizeof(VARIANT))`。这条也顺带解释了"同一批二进制在 CI 上偶发红"这类历史现象
   的一种来源：**未初始化内存 + 收尾清理**的组合。

**D65（B18 收口：端到端示例、写示例时撞到的两条 v1 边界、归档与验收面）**

1. **端到端示例的形状（`tests\cc_demo\`）**：**同一份源集合**编两种形态 —— 这是 VB6 的常规做法
   （同源多工程），也让"语言侧"与"对外侧"各有承载面：
   - `DemoExe.vbp`（Type=Exe，`Startup="Sub Main"`）：`IDemoShape`（接口宿主）+ `DemoShape`
     （`Implements IDemoShape`，含一个公有成员 `Twice`）+ `DemoHolder`（`Implements … Via m_h`）
     + `DemoBase`/`DemoDerived`（`Inherits` + `Overrides` + `Protected` + `MyBase`）+ `DemoBlocks.bas`
     （CoClass 块）→ `DEMO1..DEMO12` 各钉一个行为：组名当类型（`As Shape`/`New Shape`）、契约成员**只能经接口
     变量**调、公有成员可直调、覆盖后的虚派发、`MyBase` 去虚化、继承来的公有成员、`Protected` 家族内可用、
     基类型变量持有派生实例不切片、`Via` 委托三个槽全转发、工程内 `CreateObject(ProgID)` 改写、`TypeOf`。
   - `DemoDll.vbp`（Type=DLL，同源 + `DemoBlocksDll.bas` 的 `[ComCreatable(True)]`）→ 注册后由**另一个进程**的
     C3 客户 `DemoClient.vbp`（不引用 DLL）`CreateObject("DemoDll.Shape")` 调公有成员。
2. **写示例撞到的两条 v1 边界（都是"设计如此"，记下来免得下次再撞）**：
   ① **派生类自己 `Implements` 新式接口 → `VB3022` 拒绝**（第一版示例把 `Inherits` 与 `Implements` 写在同一个
   类上，编译期就挡住了）。所以示例把**继承链与接口链拆成两条**：`DemoBase/DemoDerived` 与
   `IDemoShape/DemoShape/DemoHolder` 互不相干。这条边界的另一半（经继承满足的接口契约、继承来的 `Public`
   字段对外暴露）此前已在 v1 边界清单里。
   ② **EXE 工程的 CoClass 块不能写 `[ComCreatable(True)]`**（`VB3033`：只有 ActiveX DLL 才注册 COM 服务器），
   所以块源按工程类型分两份 —— 这本身也是"EXE 只保留组内那一半"的正面证据（EXE 侧仍拿到
   `activated in-project: … 1 CreateObject rewrite(s)`）。
3. **一处容易误读的现象（记下来）**：`New <名字>` 里的名字若**不在本工程**（示例第一版把
   `DemoDerived.cls` 忘在 `.vbp` 外），编译器不报错，而是按 VB6 语义当成 `New <ProgID>` 走**注册表创建**
   ⇒ 运行期 `429`（弹模态框；套件里表现为超时）。**判据**：`.vbp` 的 `Class=`/`Module=` 清单是唯一事实面，
   漏登记 = 静默换语义。示例的四条用例（`cc_demo_exe[_x86]`、`cc_demo_dll_external[_x86]`）现在就是这条的护栏。
4. **归档**：`ai\027-接口继承CoClass实施收口.md`（新文档，`018` 保持"外部讨论存档"一字未改）——
   交付总览表（能力 → 用户写法 → 批次/代码）、语言/COM 两侧要点、**v1 边界清单**（12 条，逐条写现状与影响）、
   怎么验（用例名 + 命令 + 探针/助手清单）、记录索引（022 的 D1–D64 与手册四页）。手册 `CoClass 语句.md`
   末尾加了示例工程指引。
5. **验收面（四条新用例进 vbp 回归）**：`cc_demo_exe`(x64) / `cc_demo_exe_x86` 走 `Test-Vbp`（13 条 needle）；
   `cc_demo_dll_external`(x64) / `_x86` 走 `Test-ComActivateClient` —— **注册与反注册成对**（探针调
   `DllRegisterServer`/`DllUnregisterServer`，且断言 `CLEAN_CLSID/CLEAN_PROGID/CLEAN_TYPELIB=gone`），
   所以这几条在干净 runner 上可重复跑、不留键（本机跑完实测注册表三类键都查无）。
6. **本批零编译器改动**：`git status src/` 为空、`.build/C3.exe` 与 BASE 同 md5 ⇒ 逐字节/A-B 护栏由
   "编译器没变"这条代替（先例：B14 那一格）。

**D66（B19 控制台输出编码：按句柄类型选路 + 四语种实测 + cmd/PS5.1/pwsh7 三 shell 读数）**

1. **现场**：用户报"cmd 里用 C3 编译，报错信息乱码"，并追问"cmd 支持 UTF-16 输出吗"。答案是肯定的 ——
   控制台 API 的 `WriteConsoleW` 交 UTF-16，**渲染与 `chcp` 无关**；代码页只管**字节**那条路。
   两条缺陷各占一半：
   ① **编译器**：`main.cpp` 里有个 `ConsoleCodePageGuard`，进 C3 时把控制台代码页切成 65001、退出恢复。
      判据是 `GetConsoleWindow()` —— 在 **Windows Terminal / ConPTY** 下它返回 NULL ⇒ 守卫不触发，
      UTF-8 字节直接落到 936 控制台 = 乱码（用户踩的就是这个）；而且它切的是**整个控制台**，
      C3 随后拉起的 `cl.exe/link.exe` 吐 GBK 中文，在 65001 下反过来变乱码 —— 本来是好的也弄坏了。
   ② **运行库**：`Debug.Print` 走 `wprintf`。控制台上 CRT 会走宽字符路径（能显示），但**重定向/进管道**时
      按 C locale 转窄 ⇒ 中文直接丢成 `?`（实测 `?? 123`）。
2. **测量（先量再改）**：
   - **stdio 类型**（新探针 `.build/con_mode.c`）：`cmd` 里是**控制台句柄**；`powershell.exe`(5.1) 与
     `pwsh.exe`(7) 里都是**管道**（PowerShell 接走 native 输出）—— 所以"用户在哪个 shell 里跑"决定了
     走哪条路。
   - **shell 的解码**：两代 PowerShell 的 `[Console]::OutputEncoding` 默认都 = **启动时的控制台代码页**
     （本机 936/gb2312）；给它们 UTF-8 字节 = 乱码，给 GBK 字节 = 对。
   - **子进程改代码页没用**：让子进程 `SetConsoleOutputCP(65001)` 后再写 UTF-8，PS 解码结果**一字不变**
     （实测两组码点完全相同）⇒ 解码器在 shell 启动时就定了，程序影响不了它，只能顺着它写。
   - **控制台路径与代码页无关**：同一份四语种输出，控制台代码页 936/65001/437 三档读回**逐字相同**。
3. **实施**（两处同一套规则，**不碰用户的代码页**）：
   - 编译器：`main.cpp` 的守卫换成 `ConsoleUtf8Buf`（装到 `std::cout`/`std::cerr` 上，一个落点管全部
     `std::cout/cerr` 输出；按行刷，尾部半个 UTF-8 序列留到下次）。
   - 运行库：`vb6rtl_conv.c` 新增 `vb6_ConWriteHandle/OutW/ErrW`，`vb6_DebugPrintStr`/`DebugWriteBSTR`/
     `DebugWriteLong`/`DebugWriteDouble`/`DebugWriteNewline` 全改走它；`vb6rtl.c` 的"未处理运行期错误"
     也改走它，并把判据从 `GetConsoleWindow()` 换成判句柄（GUI 子系统 exe 从 cmd 里起时 std 句柄**就是**
     那份控制台，以前会弹模态框把批处理卡到有人点确定 —— 本会话早先就被它挡过一次）。
   - 规则：**句柄是控制台（含 ConPTY）→ UTF-8→UTF-16→`WriteConsoleW`**；**是管道/文件 → 按当时的
     控制台输出代码页写字节，某行装不下就整行退回 UTF-8**（`WideCharToMultiByte` 的 `lpUsedDefaultChar`
     当判据，不是硬编码任何语言/代码页）。文件重定向因此与 cmd/PowerShell 的默认解码一致
     （中文机 = GBK；`chcp 65001` = UTF-8）。
4. **四语种读数（"要通用"的验收）**：
   - **控制台**：中文/日本語 テスト/한국어 테스트/English 全对，且 936/65001/437 三档一致 ⇒ 任何语言、
     任何代码页都对（字体缺字形是另一回事）。
   - **管道路径**：默认 936 下 zh ✓ ja ✓（GBK 有假名与汉字）ko ✗（GBK 装不下 → 整行 UTF-8 → shell 按
     GBK 解码 = 乱码，这是**代码页的极限**，不是实现的）；把控制台先切 UTF-8（`chcp 65001`）再起 shell
     ⇒ 运行库看到 65001 就写 UTF-8、shell 也按 UTF-8 解码 ⇒ **中/日/韩/英四种同时成立**（pwsh 7 实测
     码点：`4E2D,6587` / `65E5,672C,8A9E,0020,30C6,30B9,30C8` / `D55C,AD6D,C5B4,0020,D14C,C2A4,D2B8` /
     `English test`）。这就是"任意语言通用"的姿势，写在手册 `Debug 对象` 页里。
5. **回归（三条用例进 vbp）**：`tests\cc_cn\CnMain.bas`（**UTF-8 BOM** 源 —— GBK 存不下韩文）+ 探针
   `tests\tools\con_capture.c`（真控制台里跑命令、**读屏幕缓冲**回 UTF-8；重定向只能量字节、量不到
   控制台那条路）+ 助手 `tests\console_enc.ps1`：`cc_cn_console`（x64）、`cc_cn_console_x86`（读屏断言
   四语种）、`cc_cn_redirect_gbk`（`chcp 936` + `> file`，**按字节**断言：GBK 装得下的行是 GBK 字节、
   韩文那行是 UTF-8 字节）。本机 3/3，且改完后既有 cc_demo(4/4)/cc_dll 外部激活(3/3)/接口·继承·CoClass
   六条仍全绿。
6. **三条方法论教训**（下次直接用）：① 读屏探针的 `LINE|` 是**屏幕行**，长行会折行，断言别指望"一眼一行"；
   探针本身也要按 UTF-8 显式读（`ProcessStartInfo.StandardOutputEncoding`）—— 否则 shell 按代码页解码
   反而把读数自己弄成乱码（本轮先踩了一次）。② **PowerShell 里 `Cp`、`Ls` 这类别名优先于同名函数**
   （`Cp` = Copy-Item），临时脚本的函数名要避开。③ 临时 `.ps1` 必须带 BOM 且只写 ASCII —— 无 BOM 的中文
   会被按 ANSI 读，本轮有一次假红就是这么来的。
7. **门后订正**：CI 那两条 `run timeout 5s` 与编码改动本身无关（同一份二进制本机 24/24 全绿地跑过、单跑 26ms）—— 是套件 5s 预算在 12 路并行编译下**误杀**，见 **D67**。

**D67（CI 的 "run timeout 5s" 假红：5s 墙钟在 12 路并行编译下分不清「饿着」与「真挂」⇒ 预算放宽到 60s + 超时读数）**

1. **现场**：B19 的两笔（`796220d` 控制台/管道编码、`d4e53c2` 未处理错误出口）推上 `github/dev` 后，
   门 **run#71** 的 bas#1 报 `test_rtl` / `test_compat` `run timeout 5s`，**run#72** 又只报 `test_compat`
   —— 每轮挂的用例都不同；而这两条在 #70（**同一镜像** `windows-2025-vs2026` 20260907.229、同一
   `-Jobs 20`、同一 toolset 14.51）里都是绿的，环境没变。
2. **先排除"用例真挂"**：
   - 同一份 `test_compat.exe` 单跑 **26ms**、40 行输出齐全；重定向 / 无控制台 / `DETACHED_PROCESS`
     三种起法都是 rc=0 ⇒ 程序本身没有会卡住的路径。
   - **本机按 CI 形状复跑**（4 核 + `-Jobs 20` ⇒ 12 worker）：24 条用例**时而 22/24 时而 24/24**，
     且每轮挂的用例不同（`hello` / `test_variant` / `test_compat`）⇒ 不是用例的错。
   - **轮询探针**（`.build/poll_stuck.ps1`，每 150ms 抓"活过 3s"的子进程）拿到现场：卡住的进程
     `state=Ready`（线程可运行却排不上）、**活到 4s 只烧了 15ms CPU**、模块表停在
     CRT/USER32 那一串 ⇒ 就是"被 12 路并行 `cl/link` 压住没被调度"。
   - **判据复核（是不是编码改动把 exe 弄慢了）**：新 exe 比旧 exe 只大 **3072 字节（1.2%）**、DLL 引用集
     **完全相同** ⇒ 启动成本没变；且**同一份二进制**本机 24/24 全绿地跑过 ⇒ **不是挂，是饿**。
   - 形状 A/B（旧 `pre_b18_C3.exe` vs 现 exe，各跑 1–3 轮）：旧 1/1 绿、新 2/3 绿 —— 样本太小，只能说明
     **这是边缘抖动**（同一二进制既能 24/24 也能 22/24），不能说明某一版引入了确定性行为。
3. **修**（`tests\run_tests.ps1`，`8174219`）：运行预算 5s → `-RunTimeoutSec`（默认 **60s**）；
   超时不再是干巴巴一句，改成报「**已烧 CPU 时间 + 进程状态 + 最后一行输出**」——
   `cpu≈0 且没输出` = 根本没被调度（机器太挤）、`cpu 不小还不退` = 真问题。真挂（模态框/死锁）照样被这条
   上限兜住，只是多等一会。串行 `Invoke-TestExe` 与并行 worker 两条路都改：`-Parallel` 的 runspace
   里**调不到脚本函数**，所以预算用 `$using:` 传、读数逻辑内联一份。
4. **判据**：`-RunTimeoutSec 0` 时两条路都把读数打出来（串行 `run timeout: 0s (cpu=16ms, alive,
   last='Sum = 5050')`、并行 `run timeout (cpu=16ms, alive, last='Count: 1')` —— 等于把"红"逼出来，
   助手能红）；默认预算下同两条用例 `PASS=2 FAIL=0`。门 = Actions run #73 全绿。
5. **三条教训**（下次直接用）：① **判"假死"要看 CPU 时间与线程状态，不能只看墙钟** ——
   `WaitForExit(5000)` 返回 false 只说明"还没退"，分不出"饿着"与"挂着"；`TotalProcessorTime` +
   `Threads[].ThreadState/WaitReason` 一起读才一眼可判（本轮 `15ms CPU + Ready` = 饿；换成
   `Waiting` + `WrUserRequest`，那就是模态框）。② **`ForEach-Object -Parallel` 的 runspace 是新的**：
   脚本函数（含 dot-source 进来的）**都调不到**，变量要 `$using:` 传 —— 超时读数因此在 worker 里内联了一份，
   不去赌 scriptblock 能不能跨 runspace 调。③ **探针的输出要显式按 UTF-8 读**
   （`ProcessStartInfo.StandardOutputEncoding`）：不指定时 shell 按 `[Console]::OutputEncoding`
   解码子进程输出，探针自己的读数会被解成乱码（`con_capture` 第一版就踩了，同 D66-6）。

### D68 V1（反引号原始多行串）落地后的三条订正（2026-09-25）

- **定界计数按 VB6 同规**：两枚相邻 = **空串**（原稿写"表达不出空串"，错），一枚字面反引号要 **四枚**（起始 + 双写 + 闭合），**三枚 = 未闭合** ⇒ `VB1007`。双写是贪心的（先看后一枚是否也成对），与 VB6 里四枚引号表示一枚引号完全同构 —— 少一条特例，用户心智也统一。用例 RS16/RS17/RS18 + 负例 `rs_n2_three_backticks`。
- **诊断文案 = ASCII，不是"与邻居同调的中文"**（计划书原稿那条被 **D12** 否掉）：断言钉 `VB1007` 号而不是文案。未闭合时**只吞掉起始那一枚就退回** —— 不退回来会把整份余文吃进一个假字面量，一条真错变成一片怪错。
- **`Declare` 那条落点的读数强度取决于桩的名字**：`Lib` 串不是"编得过就行"——DI 桩（`vb6_di_*`）按**家族**分文件、家族按 Lib 串选。第一版写 `GetTickCountRaw`（RTL 里没这枚桩）直接 `LNK2019`；换成已存在的 `GetTickCount` 之后，"名字必须折对"变成一条**自证**的判据（折错就没有家族 ⇒ 链接失败）。
- 附：`Const` 落点的发码读数 = `#define RS_CONST (vb6_BSTR_FromStr(L"k1\r\nk2 = \"v\""))` —— 行界以 `\r\n` 转义、字面量独占一行、非 ASCII 一律 `\uXXXX` ⇒ 与 `cl.exe` 的源码编码假设无关。
- 本批编译器侧净改动 = 4 个文件 `+71/-0`（纯加法）。

### D69 V2（串内插值）落地后的四条实测（2026-09-25）

- **降级点从 parser 换成词法展开（推翻计划书默认的路线 a）**，三条依据：① `peekToken2()` 的补料循环直接调 `scanToken()` ⇒ 注入必须收进**唯一取料口**（`takeScanned()`：先 `pending_` 后 `scanToken`），三处入口任何一处绕过就是**漏 token**；② parser 里按 `prevTok_.line == cur_.line` 做判定的只有 3 处（全在 `parser_expr_postfix.cpp`）且都要 `withDepth_ > 0` 或点号语境 ⇒ 注入 token 带真实行号不会牵动存量读数；③ 降级目标今天就编得过、读数也对（`("a" & CStr(n) & "b")`、`Format$(n, "#,##0")`、`Print (…& CStr(n))`）。收益：不新增 token kind、parser 两处字面量入口一处不改。
- **窗口的行列播种被实测证明是对的**：`in_n4` 故意让第二个孔里的表达式出错 ⇒ 未声明名字报在 `(8,7)`（原文件第八行第七格），而手写的 `s = "v=" & CStr(nopeHere)` 报的是**同一条** `VB3001`、同一个级别 ⇒ 计划书 §三.7 选"扫描窗口"而不是"子 buffer"的全部理由兑现。
- **一个致命的悬垂**：`content_` 是 `std::string_view`，构造里写 `content_ = buffer_->content().substr(begin, len)` 拿到的是**临时 `std::string` 的视图** ⇒ 所有源文件（连不带反引号的存量用例）在 `(1,1)` 报"意外字符"。修法：先套一层 `std::string_view(...)` 再 `substr`。**后面任何要开窗口的人都踩得到，写在这里。**
- **格式段不能沿用表达式位的扫描**：`${n:#,##0}` 里第一个 `#` 会被"跳过日期面量"的规则当成 `#…#` 的开头，一路吞到行尾 ⇒ 误报 `VB1008`。改成"顶层 `:` 之后按**纯文本**读到第一个 `}`"（计划书 §二 本来就是这个口径，是实现漏了）。同理 **坏孔要丢到本行末尾**再交 `Invalid`：不丢的话同一行尾那枚闭合反引号会被当成新串开头，一条真错配一串假错（含假的 `VB1007`）。
- 撞见两条**与本批无关**的存量弱点（A/B 确认改前改后同读数，未动，等拍板）：① `CStr(True)` 回 `"-1"`，VB6 回 `"True"`（`CStr(1 > 0)`、`CStr(t <> 0)` 同）⇒ 用例改用 `If … Then` 取读数；② 工程内类经 `CreateObject` 改写后交给 `Dim o As Object`，按名点公有 `Function` 报 `vb6_ComCall: method "…" not found`，早绑定（`As <类名>`）正常。
### D70 B21（布尔的可见性与装箱）落地后的四条读数（2026-09-25）

- **一条现象、两条根因，缺一半就只修一半**：`CStr(b)` 走的是"表达式→BSTR"的 `wrapToBSTR`/`CStr` 分支，`VarType(b)`、`Format(b,…)`、Variant 形参走的是"值→Variant 装箱"那一路。只登记 `knownBoolVars_`（可见性）之后实测 `CStr(b)=True` 但 `VarType(b)` 仍是 2、`Format(b,"G")` 仍是 `-1`；补上装箱侧才合一。**下次再撞"某个类型显示不对"，先问是哪一条路：登记表还是 `_Generic`。**
- **`_Generic vb6_VariantFromValue` 的分支表就是运行时的类型标记**：`short:` → `vb6_VariantInt`（`VT_I2`/2/"Integer"）、`int:` → `vb6_VariantLong`（`VT_I4`/3/"Long"）—— 而 `As Boolean` 在 C 层就是 `int16_t`、布尔字面量就是裸 `(-1)`，所以装箱必然落错。**同类问题这是第三次**（Single=Fix 117c、Date=Fix 175、Boolean=B21）：口径固定为「另开一张按 VB 声明登记的集合 + 消费点先判它」，不要去改 C 型（`int16_t` 与 `Integer` 同型是布局事实，动它波及 ABI 与 `.tlb`）。
- **护栏把"顺手改干净"抓了回来**：`boxToVariant` 第一版写成"实参已是 VARIANT 就直接传，不套恒等包装"，30 件工程对照当场 RED 两行 —— VbQRCodegen 的 `vb6_VarType(vb6_VariantFromValue(VB6_SA_AT(vb6_VARIANT, vParam, lIdx)))` 少了一层包装。语义上完全等价（`VariantFromValue` 对 `vb6_VARIANT` 是 identity），但它是**存量工程的可观察产物**，与本批主张无关 ⇒ 改成"非布尔一律原样回退"，对照回到只剩 4 行真该变的。**"更干净"不是这一批判据的合格理由，逐字节相同才是。**
- **BASE 的取法记一笔**：本批开工时工作树已含"可见性"那一半的未提交改动，直接 `cp` 现有 exe 当 BASE 会让护栏少测一半 ⇒ 先 `git diff > .build\w1_full.patch`，`git checkout --` 回 HEAD 构建 BASE（`pre_b21_C3.exe`），再 `git apply` 回来构建 NEW。窗口约 4 分钟、只碰自己那 13 个文件，共享树里用这种"导 patch—回退—重建—回灌"的写法而不是 stash。
- 顺带量出一条**不属于本批**的（记待拍板 7）：`Print #` / `Write #` 的非 BSTR 实参不分类型一律 `vb6_Str((int32_t)x)` ⇒ 布尔落盘 `-1`、`3.5` 被截成 `3`、`Write #1, True` 写成 `"-1"`（VB6 依次是 `True`、` 3.5`、`#TRUE#`）。
### D71 B22 开工地图（= 待拍板 6 的测量，2026-09-25 B21 收尾轮顺手做）**行号与读数均为本轮实测**

- **复现比登记的更小**：不需要 `CreateObject`。`.build\b22_out\W2App.vbp`（EXE，`W2Note.cls` 有公有 `Function Note()` 与 `Property Get Tag()`）里
  `Dim o As Object: Set o = New W2Note: o.Note()` ⇒ 空串 + stderr `vb6_ComCall: method "Note" not found`；`o.Tag` 同报
  `vb6_ComGetProp: property "Tag" not found` ⇒ **缺口在所有"工程内类实例交给 `As Object` 后按名点"这一片**，不是改写那一支。
- **改写那一支顺带量清了边界**：`swapCreateObject`（`src/driver/coclass_activate.cpp:87-99`）的 `byProgId` 只装 **CoClass 块的 ProgID**，
  所以 `CreateObject("W2App.W2Note")`（工程名 + 类名、无块）今天 429 = 设计内；补一个 `CoClass Note2 [Implementation("W2Note")]` 之后
  改写点 fire（`C3: CoClass 'Note2' activated in-project: ... 2 CreateObject rewrite(s)`），`As Object` 那一侧仍空、早绑定那一侧
  `[note-ok]` ⇒ 与 B20 登记的现象同一根。
- **发码读数**：`o = (void*)vb6_ComObject_FromInstance(vb6_FindCoClassDesc("W2Note"), (void*)vb6_cls_W2Note_New())`；类的 pack 出口
  是 `vb6_ComPack_W2Note` → `vb6_ComPackVB6InstanceRaw("W2Note", instance)`，而该函数按 `classVariable` 在 `g_vb6_coclasses` 里查 desc
  （`src/rtl/core/vb6comserver/vb6comserver_obj.c:449` 起，注释写明"未进 coclass 表的类没有 IDispatch 面 ⇒ 回 NULL"）。
  ⇒ 两条候选根因待分家：**(a)** `vb6_FindCoClassDesc("W2Note")` 回 NULL（那 `o` 干脆是 Nothing）；**(b)** desc 在、但它的成员名表为空
  （`GetIDsOfNames` 点不到 ⇒ 与 stderr 的"method not found"更合）。本轮未分家：`--emit-c` 里**看不到** server/coclass 表
  （只有 `--keep-for-debug` 的临时目录有），下一轮先从那里读 `g_vb6_coclasses` 的实参与成员表条数。
- **EXE 侧想借"对外那一档"补面是死路**：给 `Note2` 加 `[ComCreatable(True)]` ⇒ `VB3033`（EXE 工程不注册 COM 服务器，B18 已立的边界）。
  ⇒ 若走 (b)，成员表面得为"进程内晚绑定"单独发一张，不能复用注册表那一条。
- **本轮踩到的两条工具坑**（写在这里省下一轮的时间）：`.build` 里的临时 runner 用 `$ErrorActionPreference = "Stop"` 会把 C3 的
  **信息面 `C3:` 行（走 stderr）**当异常中断整个脚本 —— 这类 runner 一律 `"Continue"`；以及它报 `exit=1` 之后旧 exe 仍在原地、
  跑出来的读数全是**上一版的**（这条基线记忆里早就有："`Test-Path $exe` 不是构建成功的判据"，本轮又差一点中招）。
## 运行日志

- 2026-09-23 建表：范围确认（含完整COM）、规范文档 018 入库、现状盘点完成。
- 2026-09-23 00:07–00:20 **B00（P0 设计细化）完成**：4 路前端/符号/后端/工具链勘察 + 亲自核实关键接缝（`lexer.cpp:231-244` 方括号扫描、`parser_module.cpp:82-165` 模块级主循环与 Class 头行先例、`parser_helpers.cpp:14-63` 软关键字表、`types.hpp:54-59` AccessLevel、`cgen_com.cpp:313-428` legacy 胖对 vtable、`cgen_expr_call_com_bind.inc:67-97` 派发点），产出 D1-D11 设计记录与 B01-B18 批次表。**未构建**：一是本批纯文档无代码改动，二是工作树当时有活跃并发写者（`src/backend/detail/util/cgen_api.inc`、`src/backend/stmt/cgen_redim.cpp` 于 00:13 被改，`.build/C3.exe` 00:12:58 刚被他人重建），按纪律不得在其之上测量基线或重建 → 故 GATE_BASELINE 仍空缺，留给下次运行在静默树上建立。
- 2026-09-23 01:00–02:01 （上轮）**B01 代码完成**：`parser_interface.cpp` 新文件 + 17 文件登记（词法/AST/diagnostics/parser_module/run_tests 注册 + `test_interface.bas` + `itf_neg/` 7 负例），单文件 `cl /Zs` 预检通过，exe 01:18 已重建含 B01；但随后一轮来源不明的 `-Category all` 回归（孤儿 8204，01:59 起、父进程已死）持锁，未过门即结束。其表内"B01 ☑ + PASS=<FILL>"是**先打勾后补门**的违规写法，本轮已改实。
- 2026-09-23 02:16–03:10 **B01 过门并提交**：02:16 开工时孤儿全量回归（PID 8204，父进程已死，01:59 起）仍持 exe 锁 → 按纪律等其 02:27 自然退出（用户确认其他会话已结束）。核查 B01 全量 diff 无他人改动混入；ninja 报无活可干、`.build/C3.exe`(01:18, md5 5e9eb1cf) 已含 B01 且为最新。首跑 `-File … *> log` 因 `*>` 被当脚本参数传入致 `-Jobs` 转换失败（教训：重定向要在 `-Command` 内层）→ 改脱管 bat + 状态文件哨兵。**全量回归 02:30–02:59**：`Results: PASS=111 FAIL=0 SKIP=1 TOTAL=112`（SKIP=已知 test_vbman 环境项；跑前后 md5 一致，可归因）。B01 新用例 9 条全绿（test_interface x64/x86 + itf_n01..n07）。**逐字节护栏**：临时 worktree 建 HEAD(4fb3506，无 B01) 基线 exe，8 文件（hello/test_rtl/test_array/test_error/test_ndarray/test_generics .bas + M6Test/test_implements .vbp）`--emit-c` 双路 cmp 全同，worktree 用后即删。**GATE_BASELINE 自本行起正式建立**。B01 提交为 `3add1ce`。未开 B02（共享树里半批不可编译的代码会伤害另两位写者），改为产出 D14 开工地图供下轮直接动工。
- 2026-09-23 03:07–04:32 **B02（P1 语义层）主体过门并提交**：新 `src/driver/driver_interface.cpp`
  （stage 2.7 `runInterfacePrepass`：`Module::interfaces` → 工程级 `IfaceRegistry`，Extends 链求解 +
  槽表父先己后展平 + 五类诊断 VB3015-3018）、新 `src/semantics/semantic_analyzer_iface.cpp`
  （`checkNewStyleInterface`，在 legacy Implements 循环开头按登记表分叉，旧路径一行未改）、
  新只读头 `interfaces_registry.hpp` + `interface_sig.hpp`（槽键与签名文本的唯一出处）、
  `driver_compile.cpp` 插 2.7、`driver_semantics.cpp` 注登记表、`CMakeLists.txt` 两条源文件登记。
  用例 6 负 + 1 正（`itf_neg/n08..n13` + `itf_pos/p01_contract_ok.cls`，全 ASCII/GBK 安全的单文件通路）。
  **首门（03:30–03:53，exe b71a99a6）即 118/0/1/119**；随后复核源码发现两处应当修的地方：
  (a) 泛型模板 `.cls` 内的 `Interface` 块会因"模板本体 + 特化克隆都在 `modules_`"被登记两次，
  表现为莫名其妙的重名错 → Pass A 显式拒绝（新 ID 3018）；(b) `checkNewStyleInterface` 里
  `emplace(sig.slotKey, std::move(sig))` 实参求值顺序未定 + 失败分支读到已移空串 → 先把键取成
  局部 `const std::string key` 再用。**过门后再改源码即作废该次测量**，故 03:59 重建
  （md5 13e99638）、04:02 重跑全量终门（04:24 完成）：`Results: PASS=118 FAIL=0 SKIP=1 TOTAL=119`，
  跑前后 md5 一致、legacy `test_implements` 仍 PASS → 记为 GATE_BASELINE。成员级
  `Implements I.M` 子句未开工，登记为 B02b（详见 CURRENT_BATCH 与 D15）。
- 2026-09-23 04:42–05:56 **B02b（成员级 `Implements I.M` 子句 + 泛型 3018 用例）过门并提交 dde7c32**：
  开工按 D15-10 的硬检查等到 04:44 才确认上一轮收线（它 04:35 落代码后 04:37/04:40 还在追加 docs 提交）。
  改动 9 文件：`parser_decl.cpp` 新增 `parseTrailingImplementsClauses`（Sub 在参数表后、Function/Property 在
  `As Type` 后；点号拼接与模块级 Fix 083 对称，`A.B.C` → 接口 `A.B` + 成员 `C`）、`ast_decl.hpp` 加
  `ImplementsClause` 结构 + 三个过程节点的 `implementsClauses`、`ast_clone.cpp` 三个 clone 函数补拷贝
  （泛型特化走这条路，漏拷会静默丢绑定）、`semantic_analyzer_iface.cpp` 显式绑定入席（槽键兼容 `I.Name`
  与 `I.get_Name` 两种写法、Extends 链上父/子接口名任一；写了子句的成员**不再**参与同名隐式匹配）+ 新
  `checkMemberImplementsClauses` 兜底未认领子句（四种成因各给一句话，新 ID `SemInterfaceClauseUnbound=3019`）、
  `run_tests.ps1` 二进制插 11 行（该文件实为 **UTF-8 无 BOM + CRLF**、4899 个非 ASCII 字节，与 D9 写的
  GBK 不符，见 D16-6；改完核对非 ASCII 字节数不变 + `lf_only=0` + numstat 只 +11/-1）。用例 6 负
  （n14 无此成员 / n15 类未实现该接口 / n16 宿主非类模块 / n17 子句未限定 / n18 显式绑定下签名不符 /
  n19 = B02b-② 泛型模板内 Interface 的 3018 分支）+ 1 正（p02 覆盖"一成员认领两接口同名槽""继承槽用
  父名或子名""属性三槽 Get/Let""成员名与槽名完全无关"）。**首门（exe 1f387137，05:01–05:24）已
  125/0/1/126 零失败，但门后又改两处源码（去掉重复的 `IfaceClauseRef` 声明、畸形子句不入列表）→
  按 D15-9 那次测量不作提交依据**：05:24 重建（cdac040f）、单文件复验 7 条行为不变、05:25–05:48 复跑
  全量终门 `Results: PASS=125 FAIL=0 SKIP=1 TOTAL=126`（跑前后 md5 一致、legacy `test_implements` 仍
  PASS）→ 记为 GATE_BASELINE，B02/B02b 收口。未开 B03（余下时间不足一个"构建+25 分钟门"周期），
  改为产出 **D17 开工地图**：探针实测已把 B03 从"新语法批"降格为"宿主识别 + VB3002 重名放行 + 手册页"，
  并给出全部锚点行号与 legacy 污染待核点。
- 2026-09-23 06:08–06:46 **B03（P1 收尾：`.cls` 头行宿主形式）过门并提交 c47cdce**：人工续跑轮
  （用户"继续完成"），同轮先收 B02b（dde7c32）再做 B03。按 D17 地图开工，其中"在 parser 里识别宿主"
  一条实测不成立（`Module::moduleName` 要到 `driver_frontend.cpp:217-250` 才从 `Attribute VB_Name`
  定下来）→ 识别改放 stage 2.7 Pass A，置位新增的 `Module::isInterfaceModule`；VB3002 重名检查只豁免
  宿主自己那一个块（`hostOwnName`），`itf_n13` 撞车负例照旧报错；宿主文件里出现其它声明 → 3018 且只报
  第一条。**D17 留的"legacy 污染待核点"实测为无污染**：新用例 `tests/itf_xmod/`（头行宿主 `IWriter.cls`
  + `CWriter.cls` 用 B02b 的成员级子句跨模块绑定 Emit/Total/Last 三槽 + `XMain.bas` 具体类调用）编译
  链接成 exe 并跑出 `XMOD1:OK`/`XMOD2:OK`，登记走**第三个用例通路** `Test-Vbp`；接口类型变量
  `Dim s As IWriter` 本批刻意不做（同名 Class 符号会先被类型解析吃掉）→ 归 B04，已写进 CURRENT_BATCH。
  文档：新页 `docs/vb6-manual/02-语句/Interface 语句.md`（页首引用块标注"本项目扩展，非 MS 原生"）+
  README 索引按字母序插一行；该目录全 CRLF，Write 产出 LF 需 `sed -i 's/\r*$/\r/'` 归一并复查
  `bare_lf==0`（细节见 D18-5）。**门：06:21–06:45 全量
  `Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`**（exe md5 053150bf 跑前后一致；126→128 为新增
  p03/n20/itf_xmod_writer 三条，legacy `test_implements` 仍 PASS）。逐字节 emit-c 护栏本批未跑，
  理由与替代守卫记在 D18-6。P1 阶段（B01/B02/B02b/B03）至此**全部收口**，下一批进入 P2 发码期（B04）。
- 2026-09-23 06:46–07:00 **B03 收尾后追加 D19（B04 开工地图），未开 B04**：B04 是第一个改结构体与
  派发路径的发码批，按纪律不在共享树里留半批不可编译的后端代码。派一路只读勘察核实后端接缝，
  结果里有一条**推翻我自己 D3 的硬假设**并已亲自复核：`src/rtl/core/vb6comserver/vb6comserver_obj.c:268-272`
  与 `:302-303`、`:318` 用 `void** ppComObj = (void**)instance` 做**裸偏移 0** 读写，
  `vb6comserver.h:126-128` 还把"首字段是 `__comObj`"写成明文前置条件 →
  D3 说的"C 代码全按字段名访问、前置新字段安全"不成立，B04 必须让 `__comObj` 保持第 0 位、
  把 `__ivtbl[k]` 放在它之后。地图另记：legacy 槽表**不发适配器**（槽位直指 `IFoo_M` 实现函数，
  `cgen_com.cpp:402-413`）、`CCodeGen` 今天拿不到 `IfaceRegistry`（只有语义层注入）、
  **没有工程级"已发表"去重表**（`vb6_vtbl_<I>` 按实现类重复 typedef）、
  `isInterfaceModule` 后端零消费、胖对按值的门在 `cgen_base_type.cpp:193-197` 的 `clsSym->isInterface`、
  `vb6_ivtbl_` 前缀全仓零占用可用；建议拆 B04a（只发槽表 + 宿主不发类实例 + 逐字节护栏）
  与 B04b（薄指针 + 四处 `knownIfaceVars_` 登记 + 派发/`Set`）。**P1 阶段（B01/B02/B02b/B03）至此收口**，
  门基线 128/0/1/129（本轮未改代码，故未复跑）。

- 2026-09-23 07:05–08:04 **B04（P2 第一批：接口值代码生成）过门并提交 6bc97e8**：人工确认"没有其他写者"
  后直接开工。**D19 建议的 B04a/B04b 拆分未采纳**（只发槽表不派发没有任何可观测行为，过门等于没过门），
  节奏改成"先 `--emit-c` 观察生成物 → 再接派发 → 一次过门"。交付：新编译单元
  `src/backend/module/cgen_iface_vtbl.cpp`（14 个成员，`CMakeLists.txt` 的 `vb6c3-cgen` 登记）+
  三个发射钩子（头文件尾 / 类结构体 `__comObj` 之后 / `_New()` 里 `vt` 初始化）+
  `Set`/`MemberAccess`/`IndexOrCall` 三处接线 + `cgen_base_type.cpp` 在 Class 分支**之前**返回薄指针类型。
  两处对开工地图的修正：**不需要**工程级"已发表去重表"（`#ifndef VB6_IVTBL_<I>` 守卫 + 按小写接口名排序
  即可幂等且可复现）；槽键口径不必复制一份（`ifaceProcClauses`/`ifaceSlotPrefix`/`ifaceClauseSlotKey`
  从 `semantic_analyzer_iface.cpp` 上提到 `interface_sig.hpp`，语义比对与后端绑定查找同源）。
  实测生成物：`vb6_ivref_IWriter* s = NULL;` → `s = &(w)->__iv_IWriter;` → `s->vt->emit(s, …)` /
  `s->vt->total(s)` / `s->vt->get_last(s)`，`Set s = Nothing` 与 `If s Is Nothing` 白送（薄指针语义）。
  门 `Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（exe md5 ef1a7520 跑前后一致；用例条目数与 B03 持平，
  本批只加 3 条断言）+ 8 文件 `--emit-c` 对 pre-B04 worktree(@f8ca84e) 基线逐字节全同。
  基线 worktree 已 `remove --force` + `prune`，`git worktree list` 只剩主目录（D20-9）。

- 2026-09-23 08:08–08:59 **B05（P2 第二批：接口引用计数）过门并提交 f644003**：B04 收口后同轮续做。先派一路 Explore 专查"实例生命周期现状"，三个实测结论直接改写了方案
  （D21-1/4/5）：① 项目类实例**从来没有**作用域末尾释放——`vb6_cls_X_Destroy` 的调用点只有 COM wrapper
  归零（`vb6comserver_obj.c:83-87`）与 UserControl terminate（`cgen_form.cpp:250`）两处，所以"计数归 0"
  想可观测只有一条路：让 `New` 的那一次引用被接口变量直接收下（不 AddRef）；② 给所有类统一加计数头会
  改结构体布局、破逐字节护栏 → 字段只落在实现新式接口的类上，且在 `__comObj` 之后（D19 硬约束）；
  ③ `Set <薄指针> = Nothing` 原本落到 `vb6_ReleaseObject`，等于把薄指针当 `IDispatch*` 用（靠
  `vb6_ComIsDispatchable` 拒绝才没崩）→ 本批改成经槽真 Release。另：AddRef/Release 从 B04 的"按类一份"
  改成"按 (类, 接口) 一份"——`self` 是 `&me->__iv_<I>`，container_of 依赖具体接口的 offsetof，多接口类
  共用一份必然算错实例地址（本批开工即改）。门 `Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（exe c615c657），护栏 8 文件对 pre-B05 基线全同（worktree 用后即删）。

- 2026-09-23 09:00–09:46 **B06a（QI + 跨接口 Set + TypeOf Is 接口）过门并提交 613d2b8**：
  开工先派一路 Explore 摸 B06 接缝，三条实测结论改写了方案（D22-1/4/8）：
  ① `IfaceView.guid` 自 B01 起**只写不读**，IID 常量在生成侧一处都没有 → 本批起消费，且必须
  **按值比**（`#ifndef` 块发在每个模块头里，同接口常量各 TU 一份、地址不等）；按真 GUID 内存序
  发 16 字节，P6 交真 COM 时不用翻。② `vb6_TypeOf` 是恒返 0 的桩 —— 既有工程 `TypeOf x Is <类>`
  **一直恒假**（cTT.cls 在用），改它是行为变更 → 本批只加“右侧是新式接口”的分叉，桩一个字没动。
  ③ 跨 TU 强转薄指针不可能（类结构体只在自己 .c 里定义）→ 下行转换切给 B06b。
  实现上沿用 D21-2 的教训：QI 也按 (类, 接口) 各一份（命中兄弟接口要按各自 offsetof 取地址），
  并在类块顶部先发全部 AddRef 前向声明（发射顺序 ≠ 调用顺序）。RTL 加 `vb6_IidEqual` /
  `vb6_IfaceSupports`（QI 后立即 Release = VB6 TypeOf 的净效果），放 RTL 而不是每个模块的 static
  （免未引用警告、只一份）。用例把 CWriter 升成双接口类（IWriter+ILog，ILog 带 Property Get 再验一次
  属性槽键），并加**无实现类**的 INope 让 E_NOINTERFACE 分支真有覆盖。实跑 18 条断言全中、
  TERM 仍恰好 2 条（QI 收支平衡）。门 `Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（exe 7721bd72）；护栏 8 文件对 pre-B06a
  基线全同（worktree @f644003，用后即删）。

- 2026-09-23 09:52–10:37 **B06b（下行转换 + `TypeOf <类变量> Is <接口>` + 修 B05 的 Nothing 上转型野地址）过门并提交 4dc6b7e**：
  实现只有两件按类导出的助手（`vb6_iv_from_iv_<C>` 先比 `vt` 与自家 `&vb6_ivtbl_<I>_for_<C>` **地址全等**才减 `offsetof`；`vb6_iv_test_iid_<C>` 静态 IID 归属），
  这就是 D22-8 校正的落地——跨 TU 的障碍是身份验证不是布局（槽表实例是 owning TU 的 `static const`）。下行转换**必须 AddRef**：类变量那一遍引用按 D21-5 的不变式永不释放，
  不给它加一次码就会在接口侧 Release 到 0 时把对象从还活着的类变量脚下抽走。`TypeOf <类变量> Is <接口>` 走静态判定（类变量的动态类型恒等于声明类型 → 不需要 offsetof，
  也就不要求该类实现该接口，`TOC3` 用无实现类 `INope` 打这一发），`vb6_TypeOf` 恒返 0 的桩照旧未动。
  **必修项 D22-10 做了 A/B 实测**（最小工程只在 `.build/negctl` 里跑，未入库）：pre-B06b 基线二进制 exit=139（=0xC0000005，正是 `NULL + offsetof` 那个非空野地址被 AddRef 解引用），
  本批二进制 exit=0 且打印 `NEGCTL:Nothing-OK`；`itf_xmod` 的 `DN0` 是它的常驻版本。守卫只加在裸类变量路径——给 `New` 那条加三目会二次求值 `_New()`、凭空多创建一个实例。
  门 `Results: PASS=128 FAIL=0 SKIP=1 TOTAL=129`（exe 4202570a，10:06 起跑、最后一次源码改动在起跑前）；护栏 8 文件对 worktree @613d2b8 基线逐字节全同、用后即删；
  实跑 25 条断言全中且 **TERM 仍恰好 2 条** → 新增 AddRef 与三处上转型守卫没有多算也没有漏算引用。
  本轮另外两件事：①按用户指示把 CI 记账口径写进总表（**Actions 级全量是里程碑级**，每批仍跑本机门）；②派 Explore 出 **D24 = B07 开工地图**，其中含对 D2/D6 的 4 处硬修正
  （`SymbolKind::Interface` 从未加、`getPublicSymbols` 行号漂移、成员表清单缺 7 张、Class 头行分支要求 `(Of T)`）。边界照旧记录：`me->字段` 形式的接口/类字段不接、
  VB6 的 438 不抛（静态契约拒绝留 B13/P6）、B06c（接口值作实参/进 Variant）随 P6 处理。

- 2026-09-23 10:38–11:55 **B07a（`Inherits` 语法 + 类继承链 prepass stage 2.8）过门并提交 a056705**：
  词法接线逐处镜像 `Extends`（枚举位置天然落在 `isKeyword` 区间，仍补显式列表；软关键字表加 `Inherits` → `canBeName` 仍认它），
  并且实测**全仓 VB 语料里 `inherits` 这个词出现 0 次** → 关键字化对存量输入零影响，这一点又被 8 文件 emit-c 逐字节护栏独立证明（对 worktree @58f02fe 基线全同）。
  stage 2.8 `runClassChainPrepass` + Driver 级只读 `ClassChainRegistry`（Pass A/B/C 照抄 `driver_interface.cpp`）：登记/基类求解/环检测/链展开/深度 16，
  新诊断 3020/3021/3022；早退条件 = 工程里一条 `Inherits` 都没有（兑现 D24）。刻意保留的两条实测行为：**一条环只报一份**（与 Extends 同行为，
  `ci_n06_*` 常驻）、**接口宿主不能当基类**（宿主那一个 `.cls` 不是可实例化类 → 3020，`ci_n07_*`）。
  本批另开两条通路：`Test-SyntaxFail` 只接单源文件 → 新增 `Invoke-SyntaxProj`/`Test-SyntaxFailMulti`/`Test-SyntaxMulti`（C3 CLI 本就收多个位置参数），
  双文件负例第一次有零构建通路；运行期正例 `tests/cls_inh/Inh.vbp`（`INH0:derived` + `INH1:OK`，后者是“基类自身成员未被派生类影响”的对照组）。
  门 `Results: PASS=137 FAIL=0 SKIP=1 TOTAL=138`（exe 031f993b，11:25:47 起跑、最后一次源码改动在起跑前）。`run_tests.ps1` 登记为纯插入 +66/-0。
  本轮踩到的工具链坑（已进记忆）：heredoc/`printf` 写 `.bat` 会被 MSYS 改写（`>nul`→`>/dev/null`、`\2019` 当八进制）→ vcvars 静默失败、cmake 报找不到编译器，
  看起来像 VS 被卸；基线 bat 一律用 python r-string + CRLF 写并断言不含 `/dev/null`。另更正记忆：`run_tests.ps1` 是 **UTF-8 带 BOM**（自 B01 起就是），编辑时剥了要写回。
  B07b 地图 = D26（含“合并落 3.4”、11 张成员表、转发桩三处锚点、跨 TU 可见性待实测、v1 四类边界）。

- 2026-09-23 11:59–13:15 **B07b（继承成员合并 + 前缀布局 + 转发桩）过门并提交 b1c0050**：
  合并做成新的 **stage 3.4**（语义之后、跨模块之前），把祖先"自己声明"的成员并进派生类的 Class 符号 —— 实际并 **8 张**表而不是地图里说的 11 张：
  `publicFieldNames`/`memberFieldDispids` 并了会让 `dll_entry` 去找派生 TU 根本不发的字段访问器（**链接期才炸**）→ 推到 P6/B13，`eventNames` 无需并（带事件的基类已在 2.8 判死）；字段绝不进 `memberNames`（Fix 099 硬规定，会把字段访问改写成 `prop_get_` 调用）。
  发码走**前缀复制 + 按类转发桩**：派生 struct = `__comObj` + 祖先字段（根→叶，**含祖先私有字段**，否则基类实现的 `me->x` 全体错位）+ 自有字段，
  桩 `vb6_<D>_<M>(vb6_cls_<D>* me,…) { vb6_<B>_<M>((vb6_cls_<B>*)me,…); }` 用 `prop_get_/prop_let_/prop_set_` 三向各一份、Optional 的 `int _has_` 尾参一并转发（否则基类 `IsMissing` 失真）。
  D24④/D26 悬着的**跨 TU 可见性**这条实测解除：不用给派生 `.c` 显式 `#include "<Base>.h"`，现有互相 include 就给了桩所需的类型名，  而继承来的 `Private Type TPoint` 字段要的完整类型也在（`_New` 的 `memset(&me->m_pt,…,sizeof(me->m_pt))` 在派生 TU 编得过）；真正的证据是跑出来的 `INH3`（`d.SetPt 5` → `d.PtSum()=15`）——编译通过不等于布局对。
  **本轮最有价值的发现**：裸名调用继承成员**不是编译错误而是静默少一段代码**（`Option Explicit` 下只给 VB3001，exe 照生、`INH2` 实测 n=0）；  已按 `declaredByAncestor` 升格为 VB3022 并要求写 `Me.`，手册"注意"节同步。2.8 Pass D 另封四类边界（事件基类 / 任一侧新式接口 / 基类重载 / 重声明继承字段），
  两处踩到同一个坑：属性的 Get/Let/Set 是"一个成员键、多个方向节点"，按键去重会丢桩、按节点数判重载会误判整条链。
  门 `Results: PASS=140 FAIL=0 SKIP=1 TOTAL=141`（exe 50ce2e77，12:47:44 起跑）；护栏 8 文件对 worktree @a056705 基线 8/8 逐字节全同、用后即删。
  工具链两条：`scripts/build.bat`（LF-only + GBK 注释）在 agent 的 cmd 里会被注释中的括号字节截断 `if (...)` 块 → 构建改用 `dev.ps1 -SkipTest`；  `git checkout --` 在这个 autocrlf=true 的仓里会把工作树改成 CRLF（本轮差点用它"确认"，结果抹掉了我自己的 STATUS 改动）。
  本轮另外收了 B07a 的账（a056705 + 总表 296c556）。B08 地图 = D28（先拆 B08a=`Protected` / B08b=虚表，含四处镜像锚点与新诊断起点 3023）。

- 2026-09-23 13:22–14:06 **B08a**（P3 第三批第一半 = `Protected`）交付，代码 2117d1c。
  **拆批又拆了一层**：D28 把"Protected 可见性"当成一件事，做完发现它是两件事 —— 家族内可用（本批交付）与家族外越权要拒绝（→ 新登记 **B08c**）。后者今天做不了：`visit(MemberAccessExpr)` 不把 `obj` 解析成 Class，而"在跨模块逐字段拷贝里按访问级别过滤成员表"的偷懒做法会让越权访问**退化成运行期才炸的晚绑定 COM 调用**（比静默更糟）。手册页按这个边界如实写。
  落地面：`AccessLevel::Protected = 3`（末尾追加；`Default = Public` 是别名，插中间会撞值）、关键字接线 **13 处**（三处静默陷阱记在 D29-2：`parser_decl.cpp` 的 `else` 兜底是 Private、`isDeclarationStart` 漏 case 即 parse 错、`.h` 侧 static 判定与 `.c` 侧不同源）、新表 `Symbol::memberAccessLevels`（只有 3 个 touch point，实测全仓无 Symbol 序列化）、`.h` 声明三处放行。
  门 `Results: PASS=141 FAIL=0 SKIP=1 TOTAL=142`（exe 52bfaa97 跑前后一致）；`-Category syntax` 66/0/0；护栏 **8 文件对 worktree @b1c0050 基线 8/8 逐字节全同**（本批动了发码判定 → 按 D9 不省）；Inh.vbp 现在 14 条断言（INH12 = 跨 TU 用基类 Protected 方法、INH13 = Protected 字段读写回转）。
  两条工具/勘察教训：委托 Explore 摸消费点是划算的，但它两条锚点不准（`token.cpp:97` 其实是 `isStatementStart`；crossmod 的拷贝写的是 `extSym->/srcSym->`）→ **动手前逐条 grep 复核**，错锚点在带 `assert count==1` 的打补丁脚本里会直接 no-op。B08b 地图 = **D30**（`__cvtbl` 定序、`memberVirtual` 表、两处派发点、3024 起诊断、真虚断言形状）。

- 2026-09-23 14:12–15:30 **B08b**（虚修饰符语法 + 覆盖契约 + 封掉假虚派发）交付，代码 05397be。
  **拆批拆到第三层**：D30 把 B08b 定成"修饰符 + 类虚表 + 动态派发"，本轮只做前两件事**并把需要派发的调用点判死**，类虚表独立成 **B08d**（地图 D32）。驱动力是 D27-13 那条判例：只收语法的话，"基类体内 `Me.M`"会编得过、跑出基类实现，后代的 `Overrides` 静默失效 —— 那比编译失败危险。
  落地面：`ProcVirt` 四值枚举（**不是三个 bool**：三件套互斥，那样有 6 种非法组合要两层各防一遍）；AST 三个过程节点加 `virt` 成员（零构造函数签名改动，`ast_clone` 刻意不拷 → 模板内已拒）；parse 的修饰符**吃两次**（`Overridable Sub` 与 `Public Overridable Sub` 都收）；契约检查落在 **2.8 新增 `runVirtualContractChecks`**（放 3.4 之后 Class 符号成员表已被合并污染，分不清"谁声明的"），属性按 `ifaceSlotKey` 分方向配对、签名复用 `ifaceSigFromDecl`/`ifaceSigEqual`（与接口契约同一套函数，不留两套判据）。
  两处值得单独记：**① 拒绝点要三处**（`visit(IdentifierExpr)` 的查到符号分支、`visit(IndexOrCallExpr)` 的裸 callee 分支——它自己查符号、不经过前者，少埋一处就漏 `Speak(5)`；`visit(MemberAccessExpr)` 的 `Me.X`）；**② 返回值赋值差点被误杀**：`Speak = "base"` 在 `Speak` 自己体内长得和调用一模一样，要靠 `currentProc_` 同名豁免，否则最正面的用法第一刀就死。
  门 `Results: PASS=148 FAIL=0 SKIP=1 TOTAL=149`（exe 546265e7，14:56:16 起跑）；`-Category syntax` 66→73（新增 `ci_n12`..`ci_n18` 七条负例：无目标 / 未标 Overridable / 签名不符 / 需要动态派发 / 无 Inherits 写 Overrides / 标准模块写 Overridable / Interface 块写虚修饰符）；`Inh.vbp` 断言 14→17；护栏 8 文件对 worktree @2117d1c **8/8 逐字节全同**。
  工具链一条：worktree 基线构建用 `.build\base_build.ps1`（PowerShell 不经 MSYS，一条命令 configure+build 就过）；`scripts/dev.ps1` 在 worktree 里不行，它只 `cmake --build`、不 configure。诊断 ID 从 **3024** 起到 3027，**3023 仍留给 B08c**。下一批 **B08d**（地图 D32）。  收尾后为 B08d 摸了一遍 `cgen_inherit.cpp`，顺手改正 **D32 的一条**：表实例其实可以是 owning TU 的 `static const`（只有**类型**要跨 TU 发进 `.h`），并且多视图逼出 `__cvtbl` 只能是 `void*` + 用点强转 —— 原写法“表与表项都必须非 static”会让下一轮白改一遍发码；另把要复用的取名四件套（`cProcName`/`procBaseName`、`classMeParam`+`makeParamCType`、`inheritedRetType`）记进了 D32。
- 2026-09-23 16:15–19:45 **B08d 过门并提交（代码 `df9806e`）**：stage 3.4b `Driver::buildVirtualSlotTables`（每类有序虚槽表：按首次声明位置 根→叶、每槽键一份，筛选集取**链根**的 `dynamicKeys` → 祖先视图恒为派生表的前缀）+ 发码四件（`const void* __cvtbl` 字段 / `.h` 里本类视图的 `vb6_cvtbl_<X>` 类型 + `extern const` 表实例 / `_New()` 装载 / epilogue 末尾的表实例定义）+ 两处类成员发码路按槽索引改写 + 语义层按 D32① 删掉 `Me.X` 拒绝。
  三条**地图写错、实测才对**的教训（D32 的 ②③ 已由 D33 改正）：① `__cvtbl` 不能只给「链上带过虚修饰符的类」——**叶类也必须带**，否则基类型变量 `x.Pick()` 会让基类视图去读一个不存在的字段（= 别人第一个数据成员的地址，野指针）；判据换成「链根的 dynamicKeys 非空」就自然覆盖整条链。② 链根的 `chain` 只有一个元素，照 `chain.size() < 2` 早退会把**最该建表的那个类**漏掉（第一轮 `b.Speak()` 静默绑回 base 就是它）。③ 表类型不需要跨 TU，但**必须**发在 `<X>.h` 那批跨模块 include **之后**（消费者 TU 用的是接收者静态类型自己那张视图），且表实例要后置到 epilogue（Private 过程在 `.c` 里是 `static` 且没有前置原型）。
  一条**假成功**的形状值得记住：只改「优先级2」那一支之后，`InhMain.bas` 里的 `b.Speak()`/`d.Speak()` 已经派发对了、而类体内 `Me.Pick()` 仍是直调 → 靠 INH18/19/20/22 四条运行期 FAIL 才暴露；根因是 `Me` 是 **MeExpr**，进不了那支要求 `node.object` 为 IdentifierExpr 的分支，真正发码的地方是 `cgen_expr_member_class_fallback.inc`（它按 emit 出来的对象文本 `"me"` 查 `knownClassVars_`）。**「部分用例变绿」不等于机制接通** —— 运行期断言必须覆盖「类体内那条调用」。
  本轮跑了**三次全量门**：18:03–18:31（149/0/1/150，`gate_B08d.log`）、18:33–19:02（把扇出探针升格成常驻用例 `InhSib.cls` + INH24..26 之后重跑，同数）、19:06–19:34（发现自己这轮新写的一处代码注释把机制说反了 → 注释也参与编译、会改 exe，为保证「被测二进制 == 提交树」重建后第三次跑，最终 exe `83c4e49c`、`Results: PASS=149 FAIL=0 SKIP=1 TOTAL=150`）。教训：**升格用例与注释订正都该在起跑全量门之前做完**，一轮门就够。人工 19:35 问过「为什么一直在跑测试」，已解释，并提出可选口径（纯注释级改动以 vbp+syntax 两分类 + 8 文件护栏代证、不重跑全量）——**未拍板**，默认仍按「提交树 == 被测树」跑全量。
  其他证据：8 文件 `--emit-c` 对 worktree @fb6a254（自建 Debug exe）**8/8 逐字节全同**；`-Category syntax` 73→74（删 `ci_n15`、增 `ci_n19`/`ci_n20`）；`Inh.vbp` 运行期断言 17→27 条。扇出（一个基类两个分支）先在 `.build/probe_fan/` 探针实测（F0..F5 = `base/base2`、`bright/base2`、`bright/dark2`、上转型同值、`base/sib2`）再升格为常驻用例。工具链两条：`scripts/build.bat` 在这台机器上是 **LF-only** 的 .bat，cmd 解析 `for /f` 会碎掉（表现为一堆「不是内部或外部命令」）→ 主树构建走 `scripts/dev.ps1`，别去改别人的脚本；`cmd //c` 经 MSYS 会吞参数，PowerShell `-Command "& scripts\dev.ps1 ..."` 一条就够。下一批 **B08c**（家族外越权访问 `Protected` 的拒绝，诊断 3023 已预留，地图见 CURRENT_BATCH）。
- 2026-09-23 19:55 ~ 09-24 00:15 **B08c（`Protected` 家族外越权访问的拒绝）过门并提交 82b1b34**：
  判定落点选 `visit(MemberAccessExpr)` —— 读、写、`Set`/`Let`、`CallStmt` 与调用 callee 全都从这一处过，
  一个 `checkProtectedVisibility()` 就覆盖全部语句形状；沿 2.8 的类链按**叶优先**找最近声明者（与 3.4
  遮蔽裁决同向），再看当前模块的链里有没有那个声明者；认不出接收者、或当前类没登记（泛型模板 /
  接口宿主）一律**放过**（这里"多拒"= 把能编译的代码判死，与 B08b 的"宁多拒"取舍相反）。诊断 `VB3023`。
  两处开工地图没料到的前提：① 局部变量与参数**根本没有** `variableTypeName`（7 种接收者形状的探针实测
  有 2 种静默不响），而补满它会喂到后端十余处"非空即类实例"的消费点 = 改发码 → 新增只有本判定读的
  `Symbol::srcTypeName`（模块级字段 + 局部 + Sub/Function/Property 三处参数共 5 个登记点）；
  ② "只有 `Protected`、一条 `Inherits` 都没有"的合法工程在 2.8 开头早退 → `classes_` 全空 → 判定**整个
  静默失效** → 早退条件扩成 `anyClause || anyVirtual || anyProtected`（这类视图全是单元素链，后端
  `classChainOf()` 的 `chain.size() < 2` 守卫照旧返回 nullptr，所以只喂语义层）。
  用例：`-Category syntax` 74→78（`ci_n21` 家族外字段写 / `ci_n22` 标准模块里 `Protected Sub` 带实参调用 /
  `ci_n23` 家族外 `Property Get` 读，三条各钉一种语句形状；`ci_pos2_base|derived` 钉"家族内经基类型变量与
  `Me.` 访问必须静默"）。发码零改动 → 8 文件 `--emit-c` 对 pre-B08c(@fb6a254) **8/8 逐字节全同**；
  `Inh.vbp` 的 26 条运行期断言（含 INH12 家族内 `Me.`）与 legacy `test_implements`、`itf_xmod_writer` 全不变。
  **本轮最大的非代码收获 —— 全量门跑了四次**：前三次都被同一台机器上另一个写入者打断。它 20:41 把整个
  工作树复制到 `C:\Users\Administrator\Documents\c3.vb6.pro` 并从那份副本跑 `-Category all`，其间（20:50:23）
  还重链了我这边的 `.build\C3.exe` → v1 跑到 `test_generics_x86` 起连续 `FAIL (compile)`、powershell 以
  exit=127 死掉；更早一次 `.build\C3.exe` **直接消失**（对象文件全在，`[1/1] Linking` 即复原，`FileNotFoundError`
  不是代码问题）。v3 跑完了全程但留下 `test_softkeyword ... FAIL (compile)` —— 该用例一个类都没有、根本走不到
  本批判定，同 exe 单独复跑两次均 PASS，判为与第三方套件争抢 CPU/临时目录的瞬时失败，**不记数**；
  v4（23:26–00:07，153/0/1/154，跑前后 exe md5 33d68fc7 一致）才作为 GATE_BASELINE。
  固定动作从此加两条：跑长门前用 `.build/who_is_building2.ps1`（本轮新留：列 cmake/ninja/cl + 反查父进程与
  命令行）确认没有别的构建或套件在跑；跑完立刻核 exe md5，md5 变过的门一律重跑。
- 2026-09-24 00:55–01:45 **B08e-1（13 站地图的第 ① 站：`With w` 块内 `.M()` 接上类虚表派发）过门并提交 `e531d82`**：开工时 00:55 见上一轮（B08c）已把状态头改回 IDLE、门 153/0/1/154，按流程重占 BUSY 后从 `CURRENT_BATCH=B08e` 动工。改动一处（`cgen_expr_with.cpp` 的 ClassInstance 分支，+13/-1），口径照 B08d：`resolveClassMemberCall` 结果交给 `virtDispatchCallee(..., mustDispatch=true)` 改写，属性写方向显式排除（3.4b 无 Let/Set 槽，拿成员名去查会命中同名的 `get_` 槽、把写变成读）。**门一次过**：01:11 起跑（跑前 `who_is_building2.ps1` 静默）、01:37 收线 `Results: PASS=153 FAIL=0 SKIP=1 TOTAL=154`，跑前后 exe md5 均 `00a3a9b4` → 可归因；条目数与上一批相同，因为新证据 （`INH27..INH31`）全在既有用例 `cls_inh_pair` 内部。**A/B 负控**：worktree `D:\.wt_b08e_base` @7f34a91 建的基线二进制 `bc3eaa3b` 跑同一份用例 → `INH27:FAIL base` / `INH28:FAIL hi bob (base)`，本批两条 OK；**逐字节护栏** 8/8 全同。另外做了三件不改码的勘察（结论都在 D35）：① **13 站逐条裁决**（②③ 探针无需改、④⑤⑥⑦ 可接派发、⑧⑨⑩⑪⑫⑬ 只能判死）；② **可达性探针实证 ⑤ 站正在产出错代码** —— `Me.m_h.Speak()` 静默答 `base`，同一对象裸写 `m_h.Speak()` 已正确答 `derived`，`With w : .Speak()` 本批修好；类数组元素 `arr(1).Speak()` 是硬失败（C2224）不是静默；③ **cgen 期诊断 `--syntax-only` 抓不到**（`driver_compile.cpp` syntaxOnly 早退在 :426、3.4b 在 :376）→ 判死那一批要先给 `run_tests.ps1` 补 `Test-CompileFail` 助手。手册 `Inherits 语句.md` 顺手订正一句**把洞说成已支持**的错话（原文称字段链 `me.m_oSocket.Pick()` 正常派发，实为 ⑤ 的静默直调），并补上 `With` 已交付、`Me.<字段>.成员名` 未交付。**本轮环境事故（记给下一轮）**：给基线 worktree 建二进制时，后台 bash 任务在 ~8 分钟处被回收、把 ninja 打断在 79/114；随后 `base_build.ps1` 无条件重跑 `cmake -S` 触发 `ninja -t restat build.ninja: failed recompaction: Permission denied` →  configure 步骤假失败（对象文件其实全在，`ninja: no work to do` 即证明）。留下 `.build/base_build2.ps1`（`build.ninja` 存在就跳过 configure），下一轮建基线直接用它的 `-Wt` 形式。基线 worktree `D:\.wt_b08e_base` 本轮收线后已 `git worktree remove`。
- 2026-09-24 01:55–02:35 **B08e-2（⑤ 站：`Me.<字段>.方法()` 接上类虚表派发）过门并提交 `8987386`**：开工时 01:55 见上一轮（B08e-1）已收口为 IDLE，重占后按 `CURRENT_BATCH` 从 ⑤ 动工。改 `voidptr_com.inc` 一处（Fix 088b 分支，+10 行）：`thisArg` 算好后交给 `virtDispatchCallee(..., mustDispatch=true)`，`_prop_let_`/`_prop_set_` 跳过。新类 `tests/cls_inh/InhHolder.cls` + `Inh.vbp` 登记 + 断言 31→34。**门一次过**：02:03 起跑（跑前 `who_is_building2.ps1` 静默）、02:29 收线 `Results: PASS=153 FAIL=0 SKIP=1 TOTAL=154`，跑前后 exe md5 `62609690` 一致；日志里 43 处 "FAIL" 字样全是 `[SYNTAX-FAIL]` 用例名。**A/B 负控**：本轮起改了取基线的办法 —— 不再每批重建 worktree，而是**改码前** `cp .build/C3.exe .build/pre_b08e2_C3.exe`（= B08e-1 收线那颗 `00a3a9b4`）当"修复前"侧，省十分钟以上；用它跑同一份用例 → `INH33:FAIL base` / `INH34:FAIL hi bob (base)`，本批两条 OK，`INH32`（裸字段接收者）两侧都绿=对照。8 文件 `--emit-c` 对同一基线 8/8 全同。**本批最重要的一条是撤回**：顺手加的 `INH35`（`Me.m_up.Level` 属性读 + `Level` 的 Overridable/Overrides 对）在 ⑤ 改完后**仍拿基类实现的 `5`** —— 属性读根本不走 ⑤，而是落 ⑩ `class_fallback.inc:224`（发射出来的形状是 `vb6_InhBase_prop_get_Level((void*)me->m_up  /* class var .m_up field */)`）。据此把 ⑩ 的裁决从"统一判死"订正为"与 ⑤ 同法"（marker 只是 C 注释 → `cvtblObjIsPure` 判纯，可接派发），并**把 `Level` 属性对与 `INH35` 一起从用例里撤走**（不把错行为钉成正例），现成用例代码连同"错侧"证据写进 D35 的 ⑩ 行，下一轮直接照抄。两条踩过的坑记此：① `InhDerived` 里覆盖属性时裸写 `m_lvl` 会撞 VB3022（继承成员不许裸名），必须 `Me.m_lvl`；② 状态头那次"02:20 中途重占"写的是未来的时间（真实 02:03 才起跑门），收线时已改回事实。手册 `Inherits 语句.md` 同步：把"字段链已派发"写实、把 `Me.<字段>.<属性>` 的读列进未交付。
- 2026-09-24 02:55–03:15 **B08e-3：零代码交付（裁决纠偏 + 找到先决条件站点⑭），本轮不提交代码**：按上一轮 ⑩ 行写下的"与 ⑤ 同法可接"动工 —— 先在 `class_fallback.inc:224` 加派发，`INH35` 照旧拿基类实现的 `5`、发射出来的 C 一字未变 → ⑩ 不是属性读的落点（上一轮的推断错在这里）；换到 `m22_module.inc:120`（⑪）后立刻炸出 `error C2039: "prop_let_level": 不是 "vb6_cvtbl_InhBase" 的成员` → 落点找对了，但撞上写路径：`cgen_util_comwrite.cpp` 的 Pattern C/D2 处理 `obj.Prop = v` 是**字符串级改写**（先发一遍读再把 `prop_get_` 换成 `prop_let_`），callee 变成派发表达式后被盲换成一个 3.4b 故意不建的槽字段。**据此把 ⑩ 行退回"判死"、⑪ 行改成"真落点但被⑭卡住"，新增站点⑭（D35-6，含三条出路，建议选 (b)：让 Pattern C/D2 认得 `__cvtbl)->` 前缀）**，并把 `INH35`/`Level` 属性对与两处发码改动一起还原 —— `git status --porcelain src tests` 为空即树与 8987386 逐字节一致，**故本批不跑门**（先例 B00：纯文档批次测不到新东西），只把 `.build/C3.exe` 增量重建回 HEAD 源码并用 8 文件 `--emit-c`（对 `.build/pre_b08e3_C3.exe` 全同 8/8）+ `Inh.vbp` 34 条断言全绿来证明还原干净。**记一条流程教训**：上一轮的 ⑩ 行写着"实测订正"，但那次"实测"只看了运行结果、没核对发射出来的 C 出自哪一支，于是把错误的落点判断连用例配方一起传给了下一轮 —— 从此 B08e 系列的负控一律要附**发射形状对比**（本轮就是靠对比 ⑤/⑩/⑪ 三处的输出形状在 5 分钟内定位的）。另记：改码前 `cp .build/C3.exe .build/pre_<批>_C3.exe` 当"修复前"侧，比每批重建 worktree 省十分以上（本轮沿用）。
- 2026-09-24 03:19–04:28 **B08e-4（站点⑭ + ⑪：`Me.<字段>.<属性>` 的读接上类虚表派发）过门并提交 `392a52d`**：上一轮 D35-6 给了三条出路，本轮实际走了**第四条**——不动 `cgen_util_comwrite.cpp`，改在发码层加`bool suppressVirtDispatch_`（`cgen_state.inc`），5 个 `emitExpr(*node.target)` 求值左值期间置位，①⑤⑪ 三处派发点见到就直调：Pattern C/D2 只消费**左值文本**，只要左值永远不带派发表达式，它看到的仍是自己认识的老形状，读上下文照常派发。顺带挖出 B08e-3 没想到的事实：裸 `m_up.Level = 5` 一直正确，是因为`cgen_assign_prop_write.inc` 有一条专门的属性写路径（要求 object 是 `IdentifierExpr`）会在左值求值之前直接发`vb6_<C>_prop_let_P(obj, v)`；带 `Me.` 前缀才落进该判据之外、退到字符串级重写 —— **⑭ 的真实触发条件是"左值 + `Me.` 前缀"**，比"属性写"窄得多。⑪ 接线同 ①⑤（`thisArg083d` 一份两用、`_prop_let_`/`_prop_set_` 与 `suppressVirtDispatch_` 双跳过）。用例：`InhBase` 加 `Protected m_lvl` + `Overridable Property Get Level`/`Property Let Level`、`InhDerived` 加 `Overrides Property Get Level`（体内 `Me.m_lvl`）、`InhHolder` 加 `ViaLevel()`/`ViaLevelBare()`，断言 34→36；**A/B**：基线 `.build/pre_b08e4_C3.exe` 上 `INH35:FAIL 5`（基类 getter），本批 OK，`INH36`（裸接收者对照）两侧都绿；8 文件 `--emit-c` 对基线 8/8 全同。**门 03:29–04:20**：`Results: PASS=153 FAIL=0 SKIP=1 TOTAL=154`（跑前后 exe md5 f2547af6 一致才记账）—— 起跑那一刻 `who_is_building2.ps1` 抓到另一个写入者正从 `C:\Users\Administrator\Documents\c3.vb6.pro` 那份副本跑 `Inh.vbp`（正是 D34-8 记过的那棵树），本轮按纪律**不杀不动**，靠 md5 前后一致判定测量仍可归因；门比平常慢约一倍。**两条记过的失误**：① 提交标题从上一批复制串了行，写成"With 块内…（B08e-1）"而内容是 B08e-4 —— 共享树里不改历史，已在状态头按哈希声明清楚；② 往 `.cls` 里用 `sed -i` 插含 `&` 的行会被替换式吃掉（探针文件被弄坏一次，改用 Write 工具重写），沿用本仓"别拿 shell 内联改文件"的教训。手册 `Inherits 语句.md` 同步：`Me.<字段>.<属性>` 的读从"未交付"移到"已交付"，并写明写方向仍直调。
- 2026-09-24 04:31–06:00 **B08e-5（⑥⑦ 站：默认属性调用式 `m_up(9)` 接上类虚表派发）过门并提交 `d9eca95`**：原判的"先做 ④"被实测改写了：**④ 现在走不到静默直调** —— `u.h.Speak()`（UDT 里的对象字段）发的是非法 C `u.h.Speak()`，编译期就 `error C2039`，属**路由缺陷**，另登为站点⑮（不在 13 站计数里）。真正可达的是 ⑥⑦：同一实例、同一 `Item`，显式 `m_up.Item(3)` 早已派发、默认式 `m_up(4)` 发 `vb6_PBBase_prop_get_Item(me->m_up, vb6_VariantFromValue(4))` → 运行期答 `base4`。两处 Pattern L（`call_callee_ident.inc:144`、`call_callee_member.inc:224`，文本一模一样只差缩进）加同口径派发，带 `suppressVirtDispatch_` 保护。用例：`InhBase` 加 `Overridable Property Get Item(ByVal v As Variant)` + `InhDerived` 的 `Overrides` + `InhHolder` 的 `ItemDefault()`/`ItemBare()`，断言 36→38；A/B 用 `.build/pre_b08e5_C3.exe`（INH37 修复前 FAIL base9）；8 文件 `--emit-c` 全同 8/8；修复前后的发射形状两行 C 一起抄进 D35-8（B08e-3 的教训兑现）。**踩到的一条既有限制**（不是本批引入）：Pattern L 固定用 `vb6_VariantFromValue` 打包索引，`Item` 声明成 `ByVal i As Long` 时默认式撞 `error C2440: vb6_VARIANT → int32_t` → 本用例的 `Item` 必须是 Variant 索引。**门 04:49–05:50（61 分钟）**：`Results: PASS=153 FAIL=0 SKIP=1 TOTAL=154`、跑前后 exe md5 `ac14cf25` 一致；慢的原因是另一棵树（`C:\Users\Administrator\Documents\c3.vb6.pro`）同时在做大构建（一眼数到 43 个 C3.exe / 20 cl.exe / 37 link.exe），按纪律不碰它，只在 05:16 中途重占过一次锁。**流程变更（用户 05:36 指示）**：全量回归改到 GitHub Actions —— 推 `github` remote 的 `dev` 分支触发 `.github/workflows/ci.yml`（build + 分片 regression 矩阵），用 `scripts/watch-gh-actions.ps1` 盯；本地 `all` 门降级为里程碑级，批次内只做 `-Category syntax` + 目标用例单跑 + `--emit-c` 护栏。本批仍按本地门记账（起跑时还不知此指示）。**失误一条**：本批 commit 标题又从上一批复制串了行（连续第二次），这次在推送前用 `--amend` 改回 d9eca95 的正确标题，并用 `git diff a81d779..HEAD --stat` 校验只含本批 6 个文件、内容零变化；教训写进状态头：message 一律现写。
- 2026-09-24 05:58–06:10 **门流程切到 GitHub Actions（用户 05:36 指示）**：`git push github HEAD:dev`（快进 `a227c7a..f8e3a1a`，含本批 `d9eca95` + `f8e3a1a`）→ `.github/workflows/ci.yml` run **#10 [dev] = completed/success**，05:58 触发、06:08 收线，**约 11 分钟**跑完 7 个 job（build / smoke / bas#1 / bas#2 / syntax / compile / vbp；CI 是 **Release** 构建，与本机 Debug 门两个口径，记账时分开写）。对比：本批那轮本地全量门 04:49–05:50 = **61 分钟**（另两棵树在抢 CPU 时更慢），一小时一轮的自动化根本放不下 → **从此批内不做本地 `-Category all`**，只做三项快检（`-Category syntax` 或目标用例单跑 + 目标 vbp 单跑 + 8 文件 `--emit-c` 逐字节护栏）+ A/B 负控，全量门推 Actions 后用 `watch-gh-actions.ps1` 盯。**脚本必须用 pwsh 7 跑**：`"C:/Program Files/PowerShell/7/pwsh" -NoProfile -File scripts/watch-gh-actions.ps1 [-Once]` —— Windows PowerShell 5.1 会把这个 BOM-less UTF-8 脚本按 ANSI 读，中文注释直接炸成 `ParserError: UnexpectedToken`，看着像脚本坏了其实是用错 shell。红线不变：只推 `github` 的 `dev`，不建 MR、不碰 main、不推 `origin`(gitcode)。
- 2026-09-24 06:16– **B08e-6（⑨ 接派发 + ⑩⑫ 判死 + `Test-CompileFail` 前置）提交 `40eea3f`**：开工前按发射文本核了一遍 ⑧⑨⑬，**三条里两条裁决要改**（D35-9 ②）：⑧ 的 `Own.Speak()` 用**改动前**的二进制就报 `VB3027`（报它的是优先级2 `class_module.inc:49`，Fix 090al 的 `inferClassTypeOfExpr` 给属性标识符也推得出类名 → `itClassVar` 命中，:140 那一支对带槽成员根本走不到）→ 不改码，只用 `ci_n26` 把既有契约钉住，并写明它是防回归不是新证据；⑨ 的 `Fix 088d` 接收者是 `(void*)vb6_ret_<函数名>`（本函数返回值的裸局部，纯读）→ **能接派发**，成了本轮唯一的行为改进；⑬ 的 `vb6_UC_InstanceOf(hwnd)` 发的是 `r->me`（`uc_host.c:309`，宿主按 `typeName` 建的那个 `.ctl` 实例）→ 动态类型恒等于静态类型，直调本来就正确，**判死是给正确代码凭空造错**，改判无需改。真正落地的：⑨ 派发（`Inh.vbp` 断言 38→40，**INH39** 改前 `FAIL base`/改后 OK，**INH40** 两侧都绿当对照）、⑩⑫ `mustDispatch=true` 判死（`ci_n25`/`ci_n24` 改前 `--emit-c` 退出码 0 + 静默直调的 C 各一行，改后 1 + `VB3027`，形状都抄进 D35-9③）。**前置助手**：`tests/run_tests.ps1` 加 `Invoke-CodegenProj`/`Test-CompileFail`/`Test-Compile`，跑的是 **`--emit-c`** 而不是真编译 —— 它跑完前端+语义+发码（正好越过 `driver_compile.cpp:426` 的 syntaxOnly 早退），又不用付 cl.exe/link 的钱，实测不在源码目录落文件。护栏：8 文件 `--emit-c` 对 `pre_b08e6_C3.exe` 全同 8/8、`-Category syntax` 78→**82 全绿**；判死只在链上有 `Overrides` 时生效，而全仓除 `tests/cls_inh`/`tests/cls_neg`/`.build/probe_*` 没有任何工程写 `Overridable`（一条 `grep -rln` 就是证据），可达面即本批用例本身。**登记一条既有缺陷（未修）**：`com_entry.c` 给派生类的**覆盖**过程发 extern 时用了基类 C 类型名，而 `typedef struct vb6_cls_<基>` 的前置声明排在它后面 → `error C2143`（`Overrides` 一个返回工程类类型的方法即触发）；属 P6/B13 那片，本批探针改成不覆盖返回对象的方法绕开。**全量门 = Actions**（推 `github/dev` 触发 `ci.yml`，`watch-gh-actions.ps1` 用 pwsh 7 盯）：结果记在状态头 GATE_BASELINE。至此 13 站全部出完，下一批 **B08e-7 = 站点⑮**（UDT 对象字段的调用位置路由），⑮ 一出就开 **B09 = `MyBase` + 构造链顺序（那处要顺手把 `com_entry` 的 `Overrides` 返回类型缺陷一起处理，B09 的 `MyBase` 转发桩会先撞上它）**。
- 2026-09-24 07:16–07:50 **B08e-7 勘察轮（零代码改动，不跑门，先例 B08e-3/B00）**：按状态头领的活是"站点⑮ = UDT 对象字段的调用位置路由"，实测把它**推翻成三条不相干的缺陷**，并且**它根本不属于 B08e**（B08e 的 13 站在上一轮就出完了）：① `Set u.h = d` 发成 `u.h = vb6_VariantFromValue(d);  /* Set */` → 结构体字段本来就是 `vb6_cls_U7Base* h;` → C 编译期 **C2440**；② `u.h.Speak()`（无参）发成 `u.h.Speak()` → **C2039**；②' 更坏的一条：`p.h.Tag(3)`（带参、p 是 UDT 形参）居然**编得过**，因为发的是 `vb6_ComCall(vb6_ComGetObjectProp(p, L"h"), L"Tag", …)` —— 把工程类实例当 IDispatch 晚绑定，运行期必崩（D20 那一族的老坑），`With p : .h.Tag(4)` 同一条路；③ 对照组 `Set w = u.h : w.Speak()` 现在就正确（发射形状已是 `((const vb6_cvtbl_U7Base*)…)->speak(w)`）。**方法论一条**：`--emit-c` 对 ①② 返回 **0**、什么都不报（这两个是 C 编译期才炸的），所以上一批刚铺的 `Test-CompileFail` 帮不上 ⑮ 的忙 —— ⑮ 的验收只能走真编译（`-Category compile`/`vbp`）或正例运行期断言，别把"`--emit-c` 退出码 0"当成"这形状没问题"的证据。三条与落点（`cgen_setlet_set_rhs.inc:31/61` 的 Variant 包装没区分 `vb6_cls_X*` 结构体字段；`appendUdtObjFieldMarker`/`generic_access:125-150` 的标记通路在调用位置没被用上；判据 `udtFieldObjCType` 现成）连同实测表写进 **D36**，批次表新增 **B08f**（⑮a 写方向 / ⑮b 调用方向 / ⑮c 派发），下一轮 **B08f-1 = ⑮a**；手册 `Inherits 语句.md` 的未交付那条从"连编译都过不去，属路由缺陷"改成实测的三种行为。本轮**未动任何代码**（`.build/probe_b08e7/` 是探针，gitignore 内），因此没有门可跑、也没有可提交的构建。
- 2026-09-24 07:31–09:30 **B08f-1（跨模块 UDT 对象字段：⑮a+⑮b+⑮c 一次做完）提交 `77ecef1`**：开头按 D36 的落点去改 Set 发码（新增 `udtFieldCTypeOfTarget` + 在 `cgen_setlet_set_rhs.inc` 否决 Variant 判定），改完实测**同一条语句一字不变** → 那条守卫当时确实是死代码，revert；顺着"为什么改不到"再挖一层才碰到真根因（**D37**）：`As <项目类>` 跨模块引用时 `resolveTypeRef` 认不出来（Class 符号要到 stage 3.5 才注入本模块作用域）→ 一律回退 `Variant`，而 UDT 成员登记只在 `UserDefinedType`/`Object` 两个分支里存 `typeRefName` → **类名连名字都没留下**；同一个字段因此有两套口径：结构体发射器（跑在 3.5 之后）按名字查得到类、发 `vb6_cls_X* h;`，成员元数据说是 Variant。**判别实验**钉住结论：把同一个 UDT 声明在类模块内，`Set t.h = me` 与 `t.h.Speak()` 两条**改前就正确**（metadata 全对）。三处接线一起改：语义层 Variant 分支也存类型名（只加元数据、不动 `mi.type`；`typeRefName` 全仓 9 处读侧逐条核过，要么限定 `mi.type`、要么查到名字后还要 `kind == UserDefinedType` 才认 → 塞一个 Class 名它们全都看不见）、`udtFieldObjCType` 按当前符号表回判、Set 侧按真实 C 类型否决 Variant 容器判定（`void*` 的 COM 字段维持 Fix 084n 原样）。**通路一通就把派发接上**（D35 站点④ 至此第一次可达，那 8 行不接就正好是 D27-13 判过的"编得过、调用静默退化"，留着比接上更危险）：`u.h.Speak()` 从非法 C 变成 `((const vb6_cvtbl_InhBase*)((vb6_cls_InhBase*)(void*)u.h)->__cvtbl)->speak(...)`。**验收只能走真编译**（`--emit-c` 对这些坏形状全返回 0 —— 上一轮记的那条方法论这轮用上了）：`Inh.vbp` 断言 40→43（INH41/INH42 判别、INH43 基类对照），**改码前的二进制编不过这个工程**（C2440 + C2039，A/B 用 `pre_b08e6_C3.exe`）；8 文件 `--emit-c` 全同 8/8、`-Category syntax` 82/0。测试用例写成"UDT 与使用点同模块"，因为另两层本批没碰：**⑮d**（UDT 声明在第三个模块 → 消费者模块 `lookupModule(UDT 名)` 查不到，同一坏形状原样复发）、**⑮e**（属性写穿过 UDT 对象字段：改前 `u.h.Level = 5;`、改后 `u.h->Level = 5;`，两种都是非法 C，不是本批引入的回归）。另记一条小瑕疵：标记注释会漏进发射语句（`u.h  /* udt objfield … */ = d;`，C 语法合法、且类模块那一形改前就这样），未动。**门 = GitHub Actions**（推 `github/dev` 的 `40eea3f..77ecef1`，pwsh 7 跑 `watch-gh-actions.ps1` 盯），结论与 run#head_sha 核对记在状态头。下一批 **B09 = `MyBase` 显式基调用 + 构造链顺序**。
- 2026-09-24 08:32– **B09（`MyBase` 去虚化基调用 + 构造链根→叶 + 前置 B09-0）提交 `02bac92`**：`MyBase` 走**发码层按接收者名字接管**（precheck 的 `Err`/`VBA` 先例），lexer/parser 一行未动 → 零新语法逐字节护栏构造成立；实现取“就近声明”（基类自己 > 基类的 `inhProcs`），C 名与 B07b 中转函数、虚表表项同一套拼名，**不查 `__cvtbl`**；读按 `Get>Function>Sub>Let>Set`、写只认 `Let/Set`；`MyBase.X = v` 从 `Module.var` 回退里抢回来。判死三条形 `VB3028`（无 Inherits / 基面无此名 / 基类 Private 成员 = C 层 static）。构造链 `emitClassInitChain` + 桥接 `vb6_<基>_chain_init`（仅“写了 Class_Initialize 且被谁继承”才发），`MyBase.Class_Initialize` 复用同一座桥；`Class_Terminate` 反序链未做。前置 B09-0：`com_entry.c` 的类返回类型 extern 改 `struct vb6_cls_X*`，去掉 typedef 分块顺序依赖（`Overrides` 一个返回工程类的方法即 C2143）。`Inh.vbp` 断言 43→52（**改码前编不过这个工程**：C2143 + 5×C2065）、负例 `ci_n27`/`ci_n28` 走 `--emit-c` 断言 VB3028（改前退出码 0）、8 文件 emit-c 8/8、`-Category syntax` 82→84 全绿。门 = GitHub Actions（推 `github/dev`，pwsh 7 跑 `watch-gh-actions.ps1`），结论与 run#head_sha 核对记在状态头。下一批 **B10 = `Implements Via`（委托式实现，免手写转发桩）**；仍开：`Class_Terminate` 继承链、`Set MyBase.<属性> = obj` 未验、基类自家 extern 的 `void*` 返回类型。
- 2026-09-24 09:43 **B09 的门 = GitHub Actions run #13 [dev] = completed/success**（head_sha 已核 = `02bac92`，本机 `git rev-parse HEAD` 同值；build + smoke / bas#1 / bas#2 / syntax / compile / vbp 全绿，Release 配置）。收线前另做了一项目前不在门里的检查：**x86 侧 `Inh.vbp` 段错误**，拆用例定位到`m_pt` 之后的基类字段写入越界，根因 = 继承的 **Private UDT 字段**在派生结构体里发成 `void*`（x64 8 字节巧合对齐、x86 4 字节 → 前缀布局错位、`_New` 少分配 4 字节）。**B07b 起就有，不是本批引入**（改码前的源码在 x86 已经越界写 `g_viaRet`/`m_name`，只是没崩）。登记为 **B09b**（实测与验收要求见 **D39**），下一批先修它再开 B10。本轮零代码改动（代码已在 `02bac92`，门已过）。
- 2026-09-24 10:18– **B09b（祖先 `Private` UDT 的类型名进派生模块作用域 = x86 继承字段错位）提交 `debb110`**：先用“改一个词”的判别实验钉死洞的性质（同一工程把 `Private Type TPoint` 写成 `Public`，派生结构体立刻从 `void* m_pt;` 变 `vb6_type_TPoint m_pt;` 且 x86 从段错误变全绿）→ 缺的**只是名字**，不是定义、不是 include、不是尺寸：`getPublicSymbols()` 按 `access != Private` 过滤，祖先的私有类型符号从来没进过消费者作用域，`mapTypeRef` 于是回落 `void*`。落点在 `runCrossModuleResolution()` 末尾加一条pass（沿 `Inherits` 链、只补被 `inhFields` 用到的 UDT/Enum 名、本地同名不抢、`udtMembers` 一并复制）。**顺手清掉三条从 B07b 就挂着的 x86 红字**（INH35/INH36/INH39 读的都排在 `m_pt` 之后 = 同一个洞，不是派发逻辑错）。**门里新增 `cls_inh_x86`**（同一份 `$inhExpected` 清单跑 x64+x86 两遍）→ 布局类改动从此自动双架构覆盖。实测：`tests/cls_inh` 两架构 build+run 各 **52/52、0 FAIL**（改码前 x86 崩）；8 文件 `--emit-c` 对 `pre_b09_C3.exe` 全同 8/8（这一版 BASE 在 B09 之前，连 B09 一起证没漂）；`-Category syntax` 84/0。门 = GitHub Actions（推 `02bac92..debb110`，pwsh 7 盯 `watch-gh-actions.ps1`），结论与 run#head_sha 核对记在状态头。下一批 **B10 = `Implements Via`（委托式实现，免手写转发桩）**。
- 2026-09-24 10:56– **B09c（`Set MyBase.<属性>` 目标侧 + P3 收尾两条实测）提交 `9eb2ca7`**：三条候选先一起量再动手 —— 只有 ① 需要改码（改码前发 `vb6_T9Base_prop_set_Peer(MyBase, me)`，函数名与实参序都对、**错的只是接收者** → 在 SetStmt 各分支之前接管，`tryEmitMyBaseAssign(forSet=true)` 只认 Property Set，命中 Let 判死）；② ⑮e 早在 B08f-1 就通了（D37 那条是在坏元数据上量的，现测：写绑根 `prop_let_`、读按实例派发、运行期 109 两架构一致）→ 只补断言；③ `Class_Terminate` 继承链**不做** —— EXE 工程三种形状实测都不触发 terminate（`_Destroy` 只有接口 Release 归零与窗体销毁两个调用点），现在发就是发死代码，登记为生命周期/P6 那片的一个既有缺口。自己写错的两条期望也记进 D41（`Level` 在 `InhSib` 分支上没有槽 = 按分支建表是对的；根类 `Class_Initialize` 是赋值不是追加）。验收：`Inh.vbp` 断言 **52→59**、**x64 与 x86 各 59/59**、A/B 改码前编不过（1×C2065）、8 文件 emit-c 对 `pre_b09c_C3.exe` 8/8、`-Category syntax` 84/0。门 = GitHub Actions（推 `debb110..9eb2ca7`，pwsh 7 盯 watcher）。**P3（继承线）到此收口**：B07a/B07b/B08a-d/B08e 13 站/B08f-1/B09/B09b/B09c 全出，下一批 **B10 = `Implements Via`（委托式实现，免手写转发桩）**。
- 2026-09-24 11:40– **B09c 收线（只读 + 文档，未改编译器代码）**：Actions run **#15 [dev] = completed/success**， 收线后核 **head_sha = `9eb2ca7`** = 本机 `git rev-parse HEAD`，7 个 job（Build + smoke/bas#1/bas#2/syntax/compile/vbp） 全 success；vbp 分片日志里 `cls_inh_pair ... PASS` 与 `cls_inh_x86 ... PASS` 逐条可见，该分片 `PASS=18 FAIL=0 SKIP=1 TOTAL=19`。 状态头按此记账（`STATUS=IDLE`、CURRENT_BATCH 交出 **B10**、BASE 换 `pre_b10_C3.exe`，本轮 exe md5 `4e3c5cb9`）， 手册 `Inherits 语句.md` 补三处：`MyBase` 的 `Set` 向（含"基类那个名字只有 Let"→ `VB3028`）、 `Class_Terminate` 反序链**为什么不做**（EXE 里没有触发点，实测三种形状）、属性写穿过 UDT 对象字段 从"未交付"移到"已实测通过"（D41 订正）。**顺手把 B10 的第 0 步量掉了**（记录见 **D42**）： `Implements I Via m_f` 今天停在**语法层**（`VB2003` + `VB2002`，因为 `Via` 连软关键字都没登记 —— 四个登记点零命中）， 而去掉 `Via …` 的那份手写就是 **5 条 `VB3012`**（Extends 展开后逐槽报）→ Via 的验收面 = 这 5 条消失且零转发成员。 **P3（继承线）至此完全收口。**
- 2026-09-24 11:51– **B10（`Implements <接口> Via <持有字段>` 委托式实现）提交 `3c5d8e6`**：四个落点一次接完 —— `Via` 软关键字登记、`ImplementsStmt::viaField`、**stage 2.7 新增 Pass D** 把"字段是不是本类对象字段 / 字段类型那个类实现没实现该接口"裁决成 `vias_`（判定必须在这里做：那些类的符号 3.5 才注入，只有 2.7 看得见整工程模块表），语义层据此免掉逐槽 `VB3012`、发码层据此给没写的槽转发到 **持有对象自己的接口槽** `h->vt-><slot>(h, …)`。**关键选型**：不直调 `vb6_<持有类>_<成员>` —— 接口实现按惯例是 `Private` = C 层 `static`，跨翻译单元连不到（B09 就是这条判死了 `MyBase` 的私有目标），绕表白拿"逐槽合成"。新工程 `tests/itf_via`（纯委托 / 一半自家 / Nothing 字段 / 引用计数 / `TypeOf`）**x64+x86 各 9/9**；A/B 是 D42 那条原始读数（改码前 `VB2003`+`VB2002`）；护栏扩到 10 文件（把两个接口/继承工程也纳入逐字节对照）10/10；`-Category syntax` 84→89。**顺带量出三条既有洞并登记**（见 D43 末）：接口变量的 `Property Let/Set` 写、带 `Optional` 的接口槽调用点少发 `_has_`、接口块放 `.bas` 时调用点认不出 —— 三条都是"到达槽的 caller 形状缺失"，不是委托分支错，另批处理。下一批 **B11 = P5 `CoClass…End CoClass` 语法 + 属性折算 + 契约聚合校验**。
- 2026-09-24 13:15– **B10 收线（门 = Actions run #16 全绿，本轮只推送+盯门+记账，未改代码）**：上一次推 `github/dev` 连撞 GitHub 502/504（`api.github.com` 也 502），按"门没跑就不收 IDLE"的规矩把 STATUS 留在 BUSY 并把下一步写死在 CURRENT_BATCH 里；本轮 `-c http.version=HTTP/1.1` 推成 `9eb2ca7..3c5d8e6`，watcher 收线后核 **run#16 head_sha = 3c5d8e6**、7 个 job 全 success，vbp 分片日志里 `itf_via_pair`/`itf_via_x86` 逐条 PASS（该分片 20/0/1/21）。GATE_BASELINE 换成本批，P4 的 Via 正式记成已交付，下一批 **B11 = P5 CoClass**（实施依据 `ai/026` 的 C01→C04，开工先量 `CoClass` 块的今天行为并写 D44）。
- 2026-09-24 13:24– **B11/C01（`CoClass…End CoClass` 块语法 + 属性行落 AST）提交 `e7c7a31`**：按 022 清单第 1 条先量后写（D44 四条读数：`CoClass` 不是 token 故只第一行报 VB2002、`ProgId`/`Implementation` 不在属性白名单、`[ComCreatable(True)]` 布尔实参被 `strtoll` 判死、`[Default] Interface X` 同行写法不通），再落 C01 的三个落点（词法软关键字四件套 / AST 六处登记含 printer / parser 在 `parser_interface.cpp` 复用属性行 machinery 并加 `requireOwnLine`）。实施中暴露一条 026 没写的规则：**属性行归属要按"名字+位置+同行"合判**（`--dump-ast` 抓到身份四件套挂到了第一个接口上），顺手把 `itf_n06` 的 needle 跟着改成涵盖 CoClass。验收 = `-Category syntax` 89→95（`itf_p05`/`p06` + `itf_n25`..`n28`）+ A/B（`pre_b11_C3.exe` 停在 D44 原始读数）+ 10 文件 `--emit-c` 10/10 + 带块工程真编译真运行；C01 无语义可断，故本批**不建 vbp 工程**（运行期断言从 C05 起才有承载面）。门 = 本次 push 触发的 Actions run，结论见状态头。
- 2026-09-24 14:17– **B11/C02（身份求解唯一函数 CLSID/IID/ProgID + 可复现性）提交 `f0b820d`**：先量四件事再动手（D46）—— 现成 mint 是 `.inc` 里的局部 lambda、跨不出翻译单元；legacy 那两枚 lambda 没有任何用例覆盖（全部 `tests/*.vbp` 都是 Type=Exe）所以本批一字不动；vbp 三段式与工程名的行号复核无漂移；**"发一条 note 诊断当验收面"这个设想实测不通**（Driver 只在阶段失败时整体打印诊断，note 在成功的编译里看不见）→ 改走 stderr 的 `C3:` 信息行，且不能走 stdout，因为 `--emit-c` 的 stdout 就是 C 文本。落点 = 新单元 `src/semantics/coclass_identity.{hpp,cpp}`（纯函数唯一入口；026 没写的三条口径这次定了：vbp 查表次序 = 先 `[Implementation]` 再块名、`<Proj>` 兜序 = `Name=` > 工程基名 > "VB6EXE"、seed 一律小写而 ProgID 保留原大小写）+ stage 2.7 Pass E + `Driver::coclassIds_`。验收 = `-Category syntax` 96→99（三条断言的期望 GUID 由 python 独立复算 FNV-1a 得到，不拿编译器自己的输出当基线；含"换 vbp `Name=` 只有派生档动、显式档一字不动"与"同一输入跑两次逐字节相同"）、A/B（`pre_b11c02_C3.exe` 一行身份都不出）、10 文件 `--emit-c` 10/10、`Id.vbp` 真编译真运行（`Id.exe` 打出 cc_id）。门 = push `f660e25..f0b820d` 触发的 run#19，结论见状态头（GitHub 连撞 502/504，第 4 次才推上去，且**必须用 `ls-remote` 核实**：第 3 次那句 "Everything up-to-date" 是假象，远端当时还停在 f660e25）。另有两条工具性教训：`watch-gh-actions.ps1` 在 502 上会整脚本抛错退出（`ErrorActionPreference=Stop`），不能当可靠哨兵；CJK 长串经 `python - <<EOF` 的 stdin 会被按 cp936 解码，改用 UTF-8 脚本文件。）
- 2026-09-24 15:07– **B11/C03a（CoClass 块形状与名字校验，stage 2.7 Pass F）提交 `e515d89`**：D48 量出 C03 要拒的九种形状里**八种今天一声不吭**（唯一已挡住的是 `Inherits` 一个 CoClass 名 → `VB3020`，只是理由文案说成 unknown base class，这条只剩改文案）。本批按判据族给三个号：`VB3031` 条目（未知接口 / 把类模块当接口 / 条目重复 / `[Default]` 多标）、`VB3032` 块名撞车、`VB3033` v1 边界（`[Implementation]` 不是类模块 / EXE 工程 `[ComCreatable(True)]`）。两条实施教训：**块名等于自己宿主模块名必须豁免**（第一版把正例 p07 打回了 —— VB6 最自然的写法就是 `Widget.cls` 里写 `CoClass Widget`，与 B03 接口宿主同一条理由）；**上一批的用例是这一批的雷**——`cc_id`/`p05`/`p07` 三处正例分别声明了不存在的实现类与 EXE+ComCreatable，校验一落地全红，定案是改用例为正形状（补 `CircleImpl.cls`、ComCreatable 改 False、p07 换成宿主同名惯用法）而不是放宽校验，`True` 那一面另立 `itf_n34` 负例。验收 = `-Category syntax` 99→107（n29..n34 单文件 + n35/n36 双文件 + 旧正例全绿 + cc_id 三条身份断言一字不差）、A/B（`pre_b11c03_C3.exe` 对 n29/n34 零命中）、10 文件 `--emit-c` 10/10、`Id.vbp` 真编译真运行。门 = push 后的 Actions run，结论见状态头。**下一格 = C03b**：`runCoClassContractCheck()` 放 stage 3.4c（2.7 判不准祖先实现，D48-3）+ `As <CoClass>` 默认接口视图 + `VB3020` 文案。
- 2026-09-24 15:38– **B11/C03b（CoClass 契约聚合校验 = 新阶段 3.4c）提交 `c4aaa4c`**：D50 量出可达的"祖先供给成员"只有"祖先声明成员而不写 `Implements`"这一种形状（带 `Implements` 的两种被 B08f 的 `VB3022` 挡死），所以 D48-3 那条"2.7 看不见祖先成员"要说准成立在哪一半 —— 但结论不变，比对排在链表与成员合并之后；同批量到 `As <CoClass>` 今天**静默当 Variant**，于是把 026 五-3 整条推给 C05 与派发同批做（只做"认得是类型"的半开会把静默变成说不清的错）。实施：`runCoClassContractCheck()` 按链倒着走、槽键首见者胜（派生遮蔽祖先，与 3.4 同源），成员级子句沿用 B02b 规矩，缺槽/签名不符复用 `VB3012`/`VB3017` 两个号（主语换成 CoClass 块）；**刻意不重构** `checkNewStyleInterface` 来共用它的 impl 表 —— 那牵动 `VB3019` 的子句记账，而两侧共用的 `interface_sig.hpp` 已经保证"槽"与"签名相等"只有一份定义。`VB3020` 的文案活按 D49 收掉：只有"基名是个 CoClass 块"换新句，`ci_n01`/`ci_n07` 原句不动。D49-② 那条教训第二次应验 —— 新校验把自己的三条旧正例（`p07`/`CircleImpl`/`VbpImpl`）打红，两处 `ByVal` 撞 `ByRef`、一处压根没实现，定案照旧是改正例。验收 = `-Category syntax` 107→112（n37/n38/n39 + p08/p09）、A/B 零命中（n39 出旧句，新句本批独有）、10 文件 `--emit-c` 10/10、`Id.vbp` 真编译真运行且发码逐字节未动。**下一格 = C04**（存量 attribute 只读折算），C05 连带 `As <CoClass>` 一起做。
- 2026-09-24 19:25– **B11/C04（存量 header attribute 只读折算）**：开工先同步 —— 本地落后 `github/dev` 三枚提交，ff 到 `9099e16`（Asm x64 / 静态库 / 包引用三线并入），BASE 用合并后的树重建（`pre_b11c04_C3.exe` md5 `dfb2fbfd`），syntax 基线因此 107→**112**。实测 D52（四行头属性各 143 条、`VB_Creatable=True` 占 134 且全在 EXE 工程、`Attribute Instancing` 整仓 **0** 条、门内 8 个 `.cls` 会触发折算）⇒ 裁决"折算只记信息、不开判死"，并用 `legacyFolded` 守住 VB3020 那句文案。实施 D53（Pass E0 + `CoClassDecl::legacyFoldKeys` + identity 行尾追 `folded-from-legacy:`）。验收 = 112→116、A/B 三条零命中、护栏 16/16（含 6 个会折算工程的 `--emit-c` 逐字节对照）、`ctor_ok.vbp` 真编译真运行。门 = 本次 push 触发的 Actions run，编号见状态头。
- 2026-09-24 20:37– **B11/C04 收线（门 = Actions run #31 全绿，本轮只盯门 + 记账，未改编译器代码）**：
  收线前核 **head_sha = `c3c00e0`** = 本机 `git rev-parse HEAD`（该 head = 本批 `b82a184`/`458d9bc`/`02d70fe`
  三枚 + 并入上游 Asm 线 `3c8b204` 的合并），**8 个 job 全 success**。本批过程中 `github/dev` 被他人推进两次
  （三枚到 `9099e16`、一枚到 `3c8b204`），两次都先把本地 ff/merge 上去、再以**合并后的树**重算基线：
  syntax 107→112→**116**、逐字节护栏 10→**16**、门 job 数 7→**8** —— 下一批开工别看错数；共享分支上
  "开工先同步、收线前再同步一次"从现在起是本条线的固定动作。GATE_BASELINE 换成本批，
  BASE = `.build/pre_b11c05_C3.exe`（本轮收线 exe md5 `0bc91f03`，已核与 `.build/C3.exe` 同 md5），
  状态头收 `STATUS=IDLE`、CURRENT_BATCH 交出 **C05（组内激活）**，并把 C04 新引入的连带风险
  （`coclassIds_` 现在为每个带 header attribute 的存量类模块存着一个**与模块同名**的条目，门内 8 个）
  写进 C05 的第 0 步测量点 ③。
- 2026-09-24 20:42– **B11/C05（组内激活，同批交付 022 的 B12）提交 `7f829ee`**：先按 CURRENT_BATCH 给的三点量（D54）—— ① 吞点在**两层**且互相不认识（语义 `semantic_analyzer_typeref.cpp:135` 回退 `Variant`、发码 `cgen_base_type.cpp:293` 回退 `void*`），类符号还要 3.5 才注入 ⇒ 没有单一入口，别名表这条路要 N 个消费者都记得问；② 项目类**没有引用计数**、`Set x = Nothing` 只发 `target = NULL;`、`_Destroy` 全仓两个调用点 ⇒ 本批不新增释放面、也不碰 D41；③ 折算名今天**根本不被类型路径读**（`coclassIds_` 只喂校验与 stderr）⇒ 闸门现在下：折算记录不进表、同名模块占位时类/模块赢（先例 = Fix 177b 那句遮蔽裁决）、块名等于实现类不改。④ 三条 BASE 坏读数（`void* a = 0` + 晚绑定 / `vb6_NewObject(L"块名")` / 注册表 `vb6_CreateObject`，全 rc=0 零诊断）就是 026 五-3/4/5 的立项理由；⑤ 顺带复现 D43 第三条：`.bas` 里声明的接口当类型用时 `iv.Move 5` 发成 `vb6_StructImpl_Move((&(int32_t){5}))`（接收者丢了、类挑错了）⇒ v1 不把组名做成接口视图的第二条理由。实施（D55）：新单元 `src/driver/coclass_activate.{hpp,cpp}` 做 **stage 2.7 Pass G 就地改名**（选型照泛型单态化"语义/cgen 零感知"，兑现的收益是发码侧零新分支），`CreateObject` 走 **AST 节点替换**而非 cgen 仿冒文本 —— Fix 179a 按发码文本前缀认 New，替换成节点后 `As Object` 的 IDispatch 包装自动成立。新号 `VB3039`。验收 = syntax 116→**118**、`tests/cc_act` CC1..CC9 **x64+x86 各 9/9**、A/B 三条坏读数 base 复现 new 消失、护栏 **16/16**（`cc_id` 三个块零激活行 = 闸门证据）。门 = 本次 push 触发的 Actions run，编号见状态头。

- 2026-09-24 22:05– **B11/C05 收线（门 = Actions run #37 全绿，本轮只盯门 + 记账，未改编译器代码）**：
  收线前核 **head_sha = `0dca0d0`** = 本机 `git rev-parse HEAD`（该 head = 代码 `7f829ee` + 两处文档订正
  `df3f737`/`0dca0d0`），**8 个 job 全 success**；同一批里 `df3f737` 的 run #34 也已全绿（收线门取订正后的 head）。
  vbp 分片含本批新增的 `cc_act_pair` / `cc_act_x86`（脚本末行 `if ($script:fail -gt 0) { exit 1 }` ⇒
  job 绿就等于这两条真跑真过，不是被 Test-Path 跳过）。本批共享分支无并发：开工与收线都是
  本地 = `github/dev`（开工 `cb33ab3`、收线代码 head `0dca0d0`），没有出现 C04 那种被他人抢先推的情况。
  GATE_BASELINE 换成本批，BASE = `.build/pre_b11c06_C3.exe`（本轮收线 exe md5 `d010be65`，已核与
  `.build/C3.exe` 同 md5），状态头收 `STATUS=IDLE`、CURRENT_BATCH 交出 **B13**（P6 第一格：IUnknown
  三件套真实现 + 对象布局 COM 化收尾；第 0 步 = 先造一份 ActiveX DLL 形状的能编能跑用例），并把这批
  量到的三条既有洞（`.bas` 里声明的新式接口在调用点认不出、`TypeOf x Is <工程类名>` 恒 False、
  `ReDim a(1) As <工程类>` 撞 C2224）留进 B13 的"仍开"段供后续裁决。

- 2026-09-24 23:40– **B13a（DLL 管线的观测面）提交 `7b6570a`（测试与工程）+ `988c7cb`（D56 与手册），收线前并入 Asm 线后门 head = 合并提交 `9785f4f`**：开工第一件事就是把 CURRENT_BATCH 的前提送去实测，两枚都炸（D56）—— ① `tests/` 下**不是**全是 `Type=Exe`：`tests\test_activex_dll\` 躺着两份 `Type=DLL` 工程，真编译 25 秒产出 266 KB 的 `.dll` + `TestAXDLL.tlb` + `activex_dll.def`（五个导出函数齐全），它们的错处只是**从没登记进 `run_tests.ps1`** ⇒ 整条对外管线在门内一次都没跑过；② 「IUnknown 三件套要从零写」也是错的：`vb6comserver_obj.c:24-99` 的 QI/AddRef/Release 是真实现（原子计数、归零调 `destroyFunc` ⇒ `Class_Terminate` 在 DLL 侧**有**触发点，D41 那条只成立于 EXE）。于是本批改交付「让这一半第一次可被断言」：新助手 `Test-VbpDll`（DLL 没有 stdout ⇒ 断产物存在 + `dll_entry.c`/`activex_dll.def` 内容 + stderr 身份行；`--emit-c` **不含** `dll_entry.c`，唯一读数通道是 `--keep-for-debug` 打出的临时目录）+ 三条用例（两份现成工程 + 新写的 `tests/cc_dll/CoDll.vbp`）。D56 量出对外那半真正缺的四条、并且能指名道姓：**块名 ProgID 在产物里 0 次**、同一份编译里 `IProbe` 拿到**两枚 IID**（语义层 `{F5CEF988-…}` vs 产物 `{0AD9CBC7-…}`，因为 `coclassIds_` 后端消费者为 0）、只实现新式接口的类 `methodCount=0` 对外点不到任何方法、`comCreatable`/`legacyFolded` 对产品零影响 ⇒ B13 重切成 a/b/c 三格。本批自身零编译器代码改动；收线前并入 Asm 线 8 枚（合并提交 `9785f4f`）⇒ BASE 与基线数字一律按**合并后的树**重算（收线 exe md5 见状态头，`-Category syntax` 与逐字节护栏都在那棵树上复量）。另记一条 harness 坑：本仓 `core.autocrlf=true` ⇒ `tests/run_tests.ps1` 的**工作树** checkout 后是 CRLF（仓库里的 blob 仍是 LF+BOM），所以「插入后断言 count(CRLF)==0」这类自检要从「文件字节」改成「committed blob」（`git show HEAD:<path>`），否则合并完第一次 splice 就会被自己的断言挡下。本批真正咬人的是另一条老坑复发：注册文本经 bash heredoc 落盘时 `"$Tests\test_activex_dll"` 里的 `\t` 被吃掉成 TAB，症状是编译器报「`.vbp文件中没有源文件: D:\c3.vb6.pro\tests est_activex_dll …`」，而 `Test-VbpDll` 把它报成 `FAIL (compile)` —— **看起来像 x86 工具链问题或并发争用**（我第一反应就是后者，还「复现失败」了一次才去查字节）。⇒ 带反斜杠的 PowerShell 文本一律走 Write 工具或 `chr(92)`；新助手第一次红，先 `hex()` 看路径字节，再怀疑工具链。门 = 本次 push 触发的 Actions run，编号见状态头。

- 2026-09-25 00:58– **B13b（`dll_entry` 改读唯一出口）提交 `b1a8e58`**：接线照 `ivreg_`/`ivviareg_`/`clsreg_` 的先例加第四个只读 setter，注入点选在 `dllCgen`（`driver_codegen_dll_typelib.inc`）而**不是**模块循环 —— 这张表只有 `generateDllEntry` 消费。三条裁决：① **只接手写块**，`legacyFolded` 一律跳过（折算的 clsid 在没 vbp 三段式时是 C02 mint，种子与 legacy 不同 ⇒ 覆盖 = 换掉存量 DLL 的注册身份 = 破护栏；C04 的隔离原则第二次生效，顺带给 `legacyFolded` 第一个产品读数）；② **`[ComCreatable(True)]` 是组名档 ProgID 的唯一开关** —— 写它表里多一行（同 CLSID 同工厂，count 1→2），`<工程>.<类模块名>` 那档保留不动，存量 DLL 客户照旧 `CreateObject`；③ 块内 `[Default]` 接口的 IID 取出口值 ⇒ `IID_vb6iface_<I>` 与 `IID_vb6def_<C>` 第一次对同一个接口同值（实测两枚都成 `F5CEF988`，legacy 那枚 `0AD9CBC7` 从产物消失）。**验收证据 = 用例翻面**：`cc_dll_identity_two_channels`（B13a 钉的就是那枚分叉）改名 `cc_dll_identity_single_source`，needle/absent 整个反过来 —— 这条按 D56-7 预设的方式兑现。**没做完的两件事点名留给 B13c**（不是漏的）：非默认接口仍走 legacy derive（它要和 `ivDeriveIid` 一起收，那是第三通道），以及本批新量到的**类型库侧第四枚** `TypeLibBuilder::generateUuid`（它回写的值 dll_entry 在读，表与 `.tlb` 是否同值今天仍未证，D57-5）。**EXE 侧刻意不接**（`comEntryCgen` 不调 setter），并实测到位：`cc_act`（EXE、带两个手写块）BASE vs NEW 全产物 12 文件零 DIFF。护栏 16/16、syntax 118/0、DLL 用例 4/4（含 x86）。本轮另记两条 harness 坑（D57-8）：`src/` 是 CRLF 而 `ai/022`/`run_tests.ps1` 的 blob 是 LF（工作树因 autocrlf 又是 CRLF）⇒ 行尾断言必须按文件实测；Python 里 `("A" + nl` 换行接字符串是语法错，症状却报 「Is this intended to be part of the string?」，别去查引号。门 = 本次 push 触发的 Actions run，编号见状态头。

- 2026-09-25 01:24– **B13c（三通道 IID 合并 + 新式接口对外口径）代码提交 `bd38798`**：开工第一件事照例是把 CURRENT_BATCH 的前提送去实测，这一轮的读数是**类型库从来没有和服务器对过话** —— D57-5 说"未证"，实测答案是"**证伪**"：BASE（`473784ed…`）编 `tests\cc_dll`，同一个 `IProbe` 拿到**四枚** GUID（出口/表 `{F5CEF988-…}`、vtable QI `{74B1BA3B-…}`、`.tlb` 的 `_IProbe` `{1E34D82B-…}`、`.tlb` 的 coclass `IProbe` `{28B6B94B-…}`），而 `.tlb` 里 coclass `CImpl` 的 DEFAULT 引用是 `_CImpl = {7CA8CD81-…}`，服务器手里那枚是 `F5CEF988` ⇒ **早绑定客户端照类型库去 QI 必然 `E_NOINTERFACE`**。为了把这个"没工具"的死角变成可断言的东西，新写 `tests\tools\tlbprobe.cpp`（`LoadTypeLib`+`GetTypeAttr`+`GetRefTypeOfImplType`，151 行）与助手 `Test-TlbIdentitySingleSource`（`tests\tlb_identity.ps1`，`run_tests.ps1` 只多一行 dot-source）。合并的做法仍是不开第二套 mint：**把接口自己的 IID 也搬进唯一出口**（`resolveIfaceIid` 与块内 `[Default]` 那条共用一个函数；Pass E 建 `Driver::ifaceIds_`），三个消费者一律读它，`ivDeriveIid`/`generateIid`/`generateUuid` 降级成"表里没有（= 不是新式接口）才用"的兜底。类型库另外补了一条**结构**修正：coclass 的默认接口以前写死 `<_类名>`，现在块的 `[Default]` 指新式接口时跟随它；踩到 `unordered_map` 遍历序导致的 `TYPE_E_ELEMENTNOTFOUND`，解法刻意收窄成**只延后需要重定向的那几行** —— 第一版无差别延后所有 coclass 时存量 `TestAXDLL.tlb` 变了 3976 字节（内容同、类型序变），是实测把它否决掉的。第 2 条口径**拍板走 (b)**：契约成员（`Private`）不发成 disp id（那是换语义，而且服务器的 `GetIDsOfNames` 只认公有成员，(a) 只会两边假绿），改成编译时打一条 `C3: CoClass … is a vtable interface: 0 Public member exported for IDispatch clients`；真接口那条正路（`TKIND_INTERFACE` + 成员进库）连同实测到的新洞（`_IProbe` 的 `cFuncs` 恒 0 —— 接口模块成员从来没进类型库）一起归 B15。护栏两层都升级：16 件 `--emit-c` 改成**分类**护栏（`.build/b13c_guard.py`，只许 `vb6_iv_iid_*` 行变）→ 16/16；全产物 A/B（`.build/b13c_ab_all.py`）三工程一次量 → `cc_act` 只有 5 个 `.h` 的 iv 行变且 `com_entry.c` 全同（D57-6 的 EXE 不接仍然成立）、存量 `test_activex_dll` 含 `.tlb` **全同**、`cc_dll` 允许 `.tlb` 变 = 本批靶子，零未归类变化。门内读数：`-Category syntax` 118/0、`-Category vbp` 27/0/1（`test_vbman` SKIP 是本机没注册 VBMAN）、DLL 侧 5 条用例全绿（含 x86 与新用例）。另记两条 harness 坑（D58-8）：`scripts\env.ps1` 是 ANSI 且首行带 shebang，pwsh dot-source 会读成乱码 ⇒ 自测脚本要 MSVC 环境就自己 `cmd /c "call vcvarsall.bat x64 && set"`；heredoc 塞 Windows 路径第 N 次咬人（`\2019` 被吃掉成"系统找不到指定的路径"）⇒ 反斜杠文本一律走 Write 工具。门 = 本次 push 触发的 Actions run，编号见状态头。

- 2026-09-25 02:33– **B13d（规范 IUnknown）代码提交 `9df23ba`**：单文件改动（`cgen_iface_vtbl.cpp` 的 QI 发射器）。读数先行的老规矩又兑现一次：BASE 的 `itf_xmod\XWriter.vbp`（`CWriter` 同时实现 `IWriter`+`ILog` = 最小两接口形状）实测两处 QI 都是 `if (IidEqual(riid, IID_IUnknown) || IidEqual(riid, IID_<自己>)) *ppv = self` —— 而 `__iv_IWriter` 与 `__iv_ILog` 是同一结构体两个不同偏移的成员，所以"同一对象问 IUnknown"按入口接口不同给出两个地址（D22-7④ 从记录变成可复现的读数）。收法选**"本类实现序第一个接口的薄指针"当规范指针**：① 它是唯一一处偏移 0 就是 vtable 的合法 COM 对象指针（`vb6_cls_<C>` 偏移 0 是裸回指 `__comObj`，把 `me` 交出去等于让客户对着指针取 vtable ⇒ 直接踩飞）；② 不加字段 ⇒ D19 那条布局约束根本不进入视野，存量工程（没有 `__iv_`）形状不动；③ "实现序第一个"与"结构体最靠前的 `__iv_`"天然重合（同一份 `ivImplementedIfaces` 喂两边），不需要新序规则。单接口类 `canon == self`，值一字未变。 **本批真正的新料是顺手读出的一条危险**（D59-4）：`ComObj_QueryInterface` 对 `desc->defaultIfaceIid`/`ifaceIids[i]` 一律 `*ppv = self`（注释还写着 "dispinterface: same IDispatch pointer" —— legacy 世界里是对的），可这两个字段自 B13b/B13c 起装的**可能是新式 vtable 接口**的 IID，而 `.tlb` 自 B13c 起把同一个 IID 当 coclass 默认接口广告 ⇒ 早绑定客户 QI 成功、按接口第 3 槽调用打到 IDispatch 的 `GetTypeInfoCount` = **静默调错函数**，比 `E_NOINTERFACE` 危险。两种收法（对外不发布新式 IID / 让包装器真返回 `&me->__iv_<I>` = 实质是 B16）留 **B13e** 拍板，本批刻意不动 —— 动它要把上一批刚立的 `cc_dll_identity_single_source` 与 `cc_dll_tlb_matches_table` 两条用例翻面，同一批里既改身份口径又改对外发布口径，门就不干净了。 断言只能断形状：全链路**没有任何自生成调用点**会去问 `IID_IUnknown`（VB 没有 QI 面，接口值进 Variant/实参是未开的 B06c），端到端身份属 B17；于是新用例 `itf_canonical_iunknown` 走 B08e-6 那条 `--emit-c` 通道钉"两处 QI 各有一处规范指针分支 + 旧的 `IID_IUnknown) ||` 合并分支已消失"，负控在 BASE 上跑过（`canon` 0 处、合并分支 2 处 ⇒ 用例能红），行为那半靠真编译真跑的 `itf_xmod_writer`（QI1..QI4 + LIFE + TOF）兜。 护栏：① `-Category syntax` **119/0**（118 + 新用例）；② `.build/b13d_guard.py` 把豁免粒度从"某类行"升到"**整段函数体**"——14 个不含新式接口的工程 `--emit-c` 逐字节全同，`itf_xmod` 摘掉所有 `vb6_iunk_*_QueryInterface` 后剩余行序列全同（QI 体 36→46 行），`cc_id` 一个 QI 体都没有；③ `.build/b13d_ab_all.py` 三工程全产物 A/B：`cc_act`（EXE、带两个手写块）**零文件变化**、存量 `test_activex_dll` 只 `.rc` 那行绝对临时路径天然不同、`cc_dll` 只 `CImpl.c` 的 QI 体变（**`.tlb` 与 `.def` 一字未动**）= 零未归类变化；④ `-Category vbp` **27 PASS / 0 FAIL / 1 SKIP**（SKIP = `test_vbman`，已知基线）。手册两条：`Interface 语句.md` 补规范 IUnknown 已交付 + 跨包装器/薄指针两个世界的身份还没接成一个；`CoClass 语句.md` 补一条用户能踩的已知边界（拍板前别把"默认接口是新式接口"的 CoClass 暴露给早绑定客户）。harness 坑一条："`改动确实进了产物`"这类自检要按**全局**行数判 —— 先按单工程判时被 `cc_id`（声明接口但无人实现 ⇒ QI 体 0 处）误红了一次。 门 = 本次 push 触发的 Actions run，门 = run **#50**（8 job 全绿、head 已核 `2c7c5e2`）。

- 2026-09-25 03:23– **B13e（对外默认接口的口径：广告 == 应答）代码提交 `b10ec1a`**：本格的前提是上一批自己写的，所以第一枪照例打自己 —— D59-4 那条"胖包装器应答瘦 IID ⇒ 早绑定客户按接口虚表调第 3 槽打到 `GetTypeInfoCount`、静默调错函数"**实测不成立**：`TypeLibBuilder` 只会发 `TKIND_DISPATCH`/`TKIND_COCLASS`（本轮 `tlbprobe` 读数：`CoDll.tlb` 四个类型，`TKIND_INTERFACE` 0 个，`_IProbe` 的 kind 是 dispinterface），客户从库里学不到接口虚表布局；服务器侧 `GetIDsOfNames`/`Invoke` 也只查 `desc->methods`。方向对、机制错，真问题在旁边那条：B13c 把 `.tlb` 里 coclass 的 DEFAULT 引用重定向到块 `[Default]` 那个 `_IProbe`，而 `_IProbe` 在库里成员数是 **0**（D58-6），服务器认账的成员面是类的公有成员 = `_CImpl` 那一档 ⇒ 广告与应答是两份东西，**B13c 那次"让两条通道同值"把自己立的判据做成了空壳**（值一致、可用性没了）。拍板走 D59-4 的 (i) 收窄版：对外默认视图一律 `<_类名>`，`[Default]` 只管语言层，三处一起退回不留中间态 —— `collect.inc` 不再拿出口 `iid` 覆写 `info.defaultIfaceIid`（改由类型库回写值说话）、`tables.inc` 删掉 B13b 那条"`ifaceIids` 里 `[Default]` 同名接口改用出口 IID"的分支（`CoClassInfo::defaultIfaceName` 字段随之删）、`driver_codegen_dll_typelib.inc` 拆掉 DEFAULT 重定向连同为它搭的延后登记机器（`TlbCoClassRow`/`tlbTypes`/`registerCoClass` 与那条 `…is a vtable interface: 0 Public member…` 信息行）。**保留** `identityForClass`（coclass 的 CLSID 仍取出口）与 `ifaceIidFromMap`（接口模块那条 dispinterface 的 GUID 仍读 Pass E 的表）⇒ 甲判据"一个接口一次编译一枚 GUID"一个字没松。没走 (i) 全量（把 `ifaceIids[]` 里新式那几项滤掉）的理由写进 D60-4：那会把表与 vtable 劈成两枚 GUID、当场自相矛盾，而"胖应答瘦"的正解是 B16 让包装器真返回 `&me->__iv_<I>`。用例两条：`cc_dll_identity_single_source` 翻面（`IID_vb6def_CImpl` 回 `0x7CA8CD81` 并新增为 needle、`IID_vb6iface_IProbe` 仍 `0xF5CEF988`、撤掉那条日志断言），`cc_dll_tlb_matches_table` 从一条判据升成两条（乙 = 表的 `IID_vb6def_<C>` == 库的 DEFAULT 引用 == 库的 `_<C>`），负控跑在 BASE `905c9392…` 上 ⇒ 乙 判红（`{F5CEF988}` vs `{7CA8CD81}`）、喂新 exe 全绿。护栏：① `-Category syntax` **119/0**；② `.build/b13e_ab_all.py` 全产物 A/B 四工程（比 B13d 多带 `test_event_dll`，因为 `addCoClass` 从 lambda 搬回原地、事件源那条参数表达式重写了）：`cc_act`(EXE) **0 个文件变化**、两枚存量 DLL 连 `.tlb` 都逐字节不变（只 `.rc` 那行绝对临时路径天然不同）、`cc_dll` 只 `dll_entry.c` 的 `IID_vb6def_*` **1 行** + `.tlb` = 零未归类变化，且按全局 `def_lines>=1` 自证"回退确实进了产物"；③ `-Category vbp` **27 PASS / 0 FAIL / 1 SKIP**（SKIP 还是 `test_vbman`），收线 exe md5 `1c09fa371ad0882d2e1229fc4ebc809c`（跑测试前后一字不差）。手册一条订正：`CoClass 语句.md` 里 B13c 写的"类型库的 coclass 默认接口现在跟随块"与那条"已知边界"（第 3 槽打到 `GetTypeInfoCount`）**都是错的**，改成"对外广告的那个视图 == 服务器应答的那个视图" + 把证伪过程留在 D60-1。harness 坑一条：中文消息里紧跟变量名写 `"…_$Var: …"` 会被 PowerShell 当成作用限定符 ⇒ 整个 `tests\tlb_identity.ps1` ParserError，一律写 `${Var}`。门 = 本次 push 触发的 Actions run，编号见状态头。


- 2026-09-25 04:10– **B14（IDispatch 那一半：第一次被真客户端走通）代码提交 `ddf4e9b`（纯测试批，零编译器改动）**：照例先送测量，三件里有两件把自己的排期打了。① **四件套本身是通的**：新写 `tests\tools\disp_probe.c`（不查注册表的进程内客户端：`LoadLibrary` + `DllGetClassObject` + `CreateInstance(IID_IDispatch)`，然后把 `GetTypeInfoCount`/`GetTypeInfo`/`GetIDsOfNames`/`Invoke` 各问一遍），对真编译出来的 `TestAXDLL.dll` 实测 `Add(2,40)=42`、`SetValue(7)` → `GetValue()=7`（跨两次 Invoke 保住状态 ⇒ 同一实例、包装器复用那条路也对）、未知名 `DISP_E_UNKNOWNNAME`、`GetTypeInfo(0)` 回的是 **coclass** 那份（`kind=5`、`cFuncs`=0 是 coclass 的常态，不是缺陷）、`GetTypeInfo(1)` = `DISP_E_BADINDEX`。所以"IDispatch 四件套没接新式接口"这句话里，**缺的不是四件套，是成员面**。② 薄指针 `vb6_ivtbl_<I>` = `{QI, AddRef, Release, 自有槽…}`，**四件套那四槽不在里面** ⇒ 要给薄指针一张 IDispatch 面就得插四槽、接口自有成员槽号全体后移，这是 dual/真接口的布局活，而且**必须先有 B15 把接口成员发进库**，否则就是重演 D60（库里 0 成员、服务器多应答一枚 IID，门内全绿外面点不通）。③ 全工程只有一处生产者写 `comDispid`（`driver_codegen_dll_typelib.inc:186`，在"类模块 Public 成员"那条收集循环里），**接口契约成员从来没进过这张表** ⇒ 要么违口径 (b) 把契约成员当公有发（`GetIDsOfNames` 只认公有成员，那是两边假绿），要么走真接口 = B15。 **范围重裁（D61-5 拍板）**：B14 收缩成"把这条管线从字节级升到真跑级"这一格；"薄指针的 IDispatch 面 / 接口成员进库 / dual 布局"合并为 **B15 的前置**，顺序定死 B15 → B16（先让库里那一档有成员面，再让 QI 多交一枚指针）。交付 = 探针 + 助手 `Test-DispatchInvoke`（`tests\disp_invoke.ps1`，与 B13c 的 tlb 探针同一形态：现编现用、`.obj` 落产物目录）+ 两条 gated 用例：`ax_dll_dispatch_invoke`（存量 legacy 面真点通）与 `cc_dll_dispatch_iface_only`（只满足新式接口的类：成员面**必须**为空 ⇒ `NAMES=ADD hr=0x80020006`；同时 `QI_EXTRA ... same=yes` 把"今天这枚接口 IID 由胖指针应答"钉成实测事实，B16 要改就得故意翻它）。新助手先证明能红：喂一个不存在的 CLSID ⇒ `GETFACTORY hr=0x80040111` → 判红。零编译器改动 ⇒ 产物逐字节不变（开工/收线 exe 同一枚 md5 `1c09fa37…`），这条本身就是护栏。 **读码顺手登记的一条弱点（D61-6，本批没动）**：`ComObj_Invoke` 先按 dispid + invkind 精配、**配不上就退回"第一个同 dispid 的表项"** —— 而 VB6 里 Property Get/Let/Set 共享同一 dispid 是常态，所以一个只发成 Get 的属性用 `DISPATCH_PROPERTYPUT` 问也会把 Get 调出去；收紧会改到存量工程的可观察行为 ⇒ 属口径题，等拍板。harness 坑一条：拿 `'` 当 PowerShell 的行连接符（真身是反引号）会让整个 `run_tests.ps1` 变 ParserError，而 `-File` 那条包装只要末尾有 `echo` 就仍回 exit 0 ⇒ 本地读数要看日志内容不能只看退出码；本轮 `-Category vbp` 第一次就是这么"跑绿"的。 门 = 本次 push 触发的 Actions run，编号见状态头。


- 2026-09-25 04:51–07:00 **B15 一格出完并过门（Actions run #59，head `0caa3c5`）**：开工先做四件测量（读数与两条订正落 **D62**），按 D62-5 把范围收窄成"类型库的形状与身份那一半"后动工 —— `TypeLibBuilder::addVtableInterface`（`TKIND_INTERFACE` + Pass E 那枚 IID + 接口自己的名字，不收 `FCANCREATE`）+ `driver_codegen_dll_typelib.inc` 那条循环加接口宿主分叉，**撤掉** `coclass IProbe`（连带那枚谁也不认的 `generateUuid(类名)` mint 与"可创建"这句假广告）和 0 成员的 `_IProbe`。库里 types 4→3。判据换读法不放宽：`Test-TlbIdentitySingleSource` 通道 3 改读真接口行 + 两条反面断言，负控（`.build/b15_negctl.ps1`）实测旧库 RED 三条全中、新库 GREEN；`--emit-c` 16 件全同、全产物 A/B 只 `cc_dll` 的 `.tlb` 变。成员面（`cFuncs` 仍 0）连同 `CreateTypeLib2` 位数 flag 排 B16，理由都在 D62-1/D62-3 里有读数。代码提交 `941b6dc`；**门 head 是合并后的 `0caa3c5`**（用户在 `github/dev` 上的 Fix 192/193 —— `a4e8eb3` 元素类型为项目类的数组、`7e9b6bc` `TypeOf … Is <项目类>` —— merge 进来无冲突，两条已从本表"仍开"清单撤下并记在 CURRENT_BATCH 末）。收线 exe（merged tree）md5 `a7e41e5d…` 已存成 `.build/pre_b16_C3.exe`。 **本轮四条流程教训（都是自己的锅，记下来给后续批次）**：① **`Test-Path exe` 不是构建成功的判据** —— 一次 `cl` 失败但旧实验 exe 还在原地，脚本照样跑完，量出来的"③ flag 无后果"是**上一版二进制的读数**（重测后结论翻转，见 D62-3）；改成先删旧产物、再看 cl 输出里有没有 `error C`。② **PS 5.1 按 ANSI 读无 BOM 的 `.ps1` 时，行尾的中文字节会把换行吃掉** —— `.build` 里一个临时脚本的 `param()` 因此静默失效（参数全空、脚本空跑），临时脚本一律 ASCII only。③ **bash 传给 `powershell -File` 的 Windows 路径里反斜杠会被 MSYS 吃掉** —— `-OutputDirectory D:\c3…` 变成仓库根下一个 `c3.vb6.pro.build…` 目录；要么写 `"D:\\…"` 要么用正斜杠，且**别把带 `-File` 的调用当可以直接信 exit code**（同 B14 那条）。④ **本地全量/分类门这一轮两次作废**：第一次是自己 `rm -rf` 了正在被自己那次跑使用的产物目录（把 M7Test 与三条 DLL 用例弄成假红），第二次是默认 `output/` 与另两位写者的并发套件互踩（22 条假红）⇒ 本地只记快检（emit-c 护栏 / A/B / 负控 / 单工程 tlbprobe 读数），门数一律以 Actions 为准，这正是"门跑 Actions"的理由。

- 2026-09-25 07:20–08:34 **B16 一格出完并过门（Actions run #62，head `578faa4`）**：开工先做三件测量（读数与裁决落 **D63**）—— ① **x86 调用约定今天不成立**：`--arch x86 --emit-c` 的产物里一个约定修饰都没有（`grep -c __stdcall` **BASE 0 / NEW 21**，`cc_dll`），而 COM 规范是 `CC_STDCALL` ⇒ 薄面整条 canonical 化；② **canonical 化的波及面**靠分类护栏逐条逼出来（`__stdcall` 只出现在声明/实现新式接口的工程、薄面那一对 `vb6_iv_thin_/vb6_iv_claim_` 只出现在**有接口宿主**的工程、服务器表每个 coclass 行多 `ifaceThinPtr`/`instanceClaimRelease` 两个字段 = 新式宿主填函数名、存量填 `NULL`），`cc_id` 只声明接口没有宿主 ⇒ 只有 `__stdcall`（护栏因此把 `EXPECT_FIRING` 与 `EXPECT_THIN_HOSTS` 分开写死）；③ **位数 flag 拍板跟 `--arch` 走**并把代价量到字节：默认架构存量库**逐字节不变**（md5 同 `87d34673`、diffbytes=0），`--arch x86` 只差 **32 个字节**（@20 头部位数声明字 `41`/`43` + 31 处布局字），且两架构的差集与 BASE→NEW 的差集**同址**、x64/x86 两枚读端读回的值逐项相同。实施**零布局改动**（D19 不入视野）：`cgen_iface_vtbl.cpp`（`__stdcall` + 薄面出入口）、`cgen_util_dllentry_collect/prelude/tables.inc`（表行两个新字段）、`vb6comserver.{h,c}` + `vb6rtl_class_com.h`（接口 IID 的 QI 交薄指针、包装器把底座引用交还最后一个薄引用）、`typelib_builder.{hpp,cpp}`（如实发契约 + flag 跟位数）、`driver_codegen_dll_typelib.inc`（喂 `IfaceView::slots`，= D62-2 认定的唯一事实面）。判据五条全上读数：新探针 `tests\tools\tlb_slots.cpp` 升格入库 + 助手 `tests\tlb_contract.ps1` 两条契约用例（x64 `oVft=24/32`、x86 `oVft=12/16`，**x86 那条必须用 x86 读端**）+ `tests\tools\disp_probe.c` 长一条按库里形状**直调 vtable 槽**的支路（`VTBL_GET_BEFORE=0` → 槽 3 写 21 → `VTBL_GET_AFTER=42`，x64 与 x86 各一条真跑用例）+ `QI_EXTRA same=yes→no` 的**故意翻面**（接口 IID 现在交薄指针）。护栏：`--emit-c` 17 件 `violations=0`、全产物 A/B `unclassified=0`（只有 `cc_dll` 的 `.tlb` 1484→1624 的预期变化）、负控在 `pre_b16_C3.exe` 上 `pass=0 fail=2`（四条 needle 全 miss + 反面断言 `cFuncs=0` 命中）。**三条流程教训**（写进 D63-8）：套件在 CI 上用 pwsh 7（`Join-Path` 4 参形式在 PS 5.1 下静默给空串 ⇒ 本地复现也得用 pwsh）、套件 `$env:LIB` 是 x64 在前（x86 助手链接要临时换 `$msvc.LibX86`，否则 LNK4272 + 10 个未解析外部）、探针直调薄指针的槽要**先解一层**（`*(probe_ivtbl**)p`；`((probe_ivtbl*)p)->Got(p)` 是拿指针当槽表，实测读到堆外当场崩）——顺带：崩溃型探针必须 `setvbuf(stdout, NULL, _IONBF, 0)`，否则块缓冲把已打印的读数全吞掉。行里第 4 项（DllRegisterServer 一族接线）经读码 + A/B 判定**无需新改动**（注册写入读同一张服务器表、身份自 B13b 起走唯一出口；`.def`/exports/`.rc` 一字未动），真注册的外部端到端验收归 **B17**。收线 exe（= 门 head 的树）md5 `90129adbb5dbb3c5fca14fcc872dc0f2`，已存成 `.build/pre_b17_C3.exe` 作下一批 BASE。

- 2026-09-25 08:40–10:20 **B17 一格出完并过门（Actions run #66，head `732c1f8`）**：开工按批次要求先做两件测量（读数与裁决落 **D64**）—— ① **注册与外部激活逐级读数**：`DllRegisterServer` 那一级通（CLSID/ProgID/InprocServer32/ThreadingModel 四项都在），`CLSIDFromProgID` + `CoCreateInstance(IID_IDispatch)`（= `CreateObject` 那条路）通，**断在类型库那一级**（`LoadTypeLibEx` 读不到内嵌库 ⇒ 注册表里没有 TypeLib 项）；挖下去是**两条静默缺陷**：`rc.exe` 的发现面太窄（旧写法只认 `WindowsSdkDir` 与 `C:\Program Files (x86)\…`，而本机 SDK 在 `D:\Windows Kits\10` ⇒ 类型库与版本信息**都不嵌**，BASE 产物连 `.rsrc` 段都没有）与 `UnRegisterTypeLib` 的**实参顺序写反**（`lcid`/`syskind` 互换 ⇒ 反注册静默漏掉整棵 TypeLib 键、而 RTL 无条件返回 S_OK）。② **早绑定外部读数**：注册过的、由系统装载的 DLL 上按接口 IID 激活拿薄指针、按库里 `oVft` 直调契约槽得 42（x64 + x86）⇒ 顺带裁决 **canonical 返回形状不做**（真按 oVft 直调的客户已经能用）。实施 = 两处修复 + `tests\tools\com_act_probe.c`（升格入库）+ 助手 `tests\com_activate.ps1`（整趟冒烟 / 注册→跑外部客户→反注册）+ 夹具 `tests\cc_dll_client\`（C3 编出来的**外部**客户 EXE，走真实注册表链路）+ `CImpl.cls` 加一个公有成员 `Twice`（晚绑定要有东西可点；B13b 的 `methodCount` 针 0→1）。三条新用例进 vbp 回归（x64/x86 冒烟 + 外部客户）。护栏：`.build/b17_ab_all.py` 三段 —— `--emit-c` **17 件逐字节全同**、全产物 A/B 只差 `rtl/vb6comserver.c` 的 10 行（全在反注册那段）、**资源目录读数**（NEW 带不带 `WindowsSdkDir` 都嵌 = 修好了；BASE 不带就不嵌 = 缺陷现场）；负控 `.build/b17_negctl.ps1` 喂收线前的 `pre_b17_C3.exe` ⇒ **pass=0 fail=2**，红的正是 `CLEAN_TYPELIB=STILL`。**三条教训写进 D64-7**：产品二进制跨构建不可比（`link.exe` 写 PE 时间戳，同一个 exe 连编两次 md5 都不同 ⇒ 判据要看资源目录/emitted 文件）；探针的模式词别写死在 `argv[5]`（`… <progid> reg` 会静默退化成整趟跑，注册完顺手反注册 ⇒ 下一个进程再激活就 `429`）；GUI 子系统 exe 的未处理运行期错误会弹**模态框**（套件里表现为"跑超时被杀"）⇒ 客户夹具要 `On Error Resume Next` 把失败落到读数上（本机手工跑时真被这个框挡过一次）。**门后补一刀**：第一次门（#65）的 vbp job 在既有用例 `ax_dll_dispatch_invoke` 上报 exit=0xC0000374（堆损坏），读码 + A/B 把它定位成 `ComObj_Invoke` 的 `coercedArgs`未初始化就被 `VariantClear`（本机 12 次不复现 = 随堆状态偶发，非本批引入），修为整段清零（`732c1f8`）后重跑门 #66 全绿。收线 exe（含该修复）md5 `38f247e30ed5ac3513366c9733bf4dda`，已存成 `.build/pre_b18_C3.exe` 作下一批 BASE。

- 2026-09-25 10:30–11:35 **B18 收口批出完并过门（Actions run #68，head `0413bb9`）—— 本线全部交付，`STATUS = ALL_DONE`**：本批**零编译器改动**（`git status src/` 空、`.build/C3.exe` 与 BASE 同 md5 ⇒ 逐字节/A-B 护栏由"编译器没变"这条代替，先例 B14），交付三样：① **端到端示例 `tests\cc_demo\`** —— 同一份源集合编两种形态（VB6 常规做法）：`DemoExe.vbp` 语言侧把整条线用一遍（`IDemoShape` 接口宿主 + `DemoShape` 的 `Implements` + `DemoHolder` 的 `Via` 委托 + `DemoBase`/`DemoDerived` 的 `Inherits`/`Overrides`/`Protected`/`MyBase` + `CoClass Shape` 块 + `As Shape`/`New Shape`/工程内 `CreateObject` 改写/`TypeOf`），`DEMO1..DEMO12` 各钉一个行为、x64+x86 双跑；`DemoDll.vbp` 对外侧（同源 + 块里 `[ComCreatable(True)]`）注册后由**另一个进程**的 C3 客户 `DemoClient.vbp`（不引用 DLL）`CreateObject("DemoDll.Shape")` 调公有成员。② **归档 `ai\027-接口继承CoClass实施收口.md`**（`018` 保持"外部讨论存档"一字未改）：交付总览表、语言/COM 两侧要点、**v1 边界 12 条**（逐条现状 + 影响）、怎么验（用例名/命令/探针与助手清单）、记录索引；手册 `CoClass 语句.md` 加示例指引。③ **写示例撞到的三条读数进 D65**：派生类自己 `Implements` 新式接口 = `VB3022`（示例因此把继承链与接口链拆成两条）、EXE 的 CoClass 块不能写 `[ComCreatable(True)]` = `VB3033`（块源按工程类型分两份，EXE 侧仍拿到 in-project 激活）、`New <名字>` 不在工程内时编译器**不报错**而是按 VB6 语义当 `New <ProgID>` 走注册表创建 ⇒ 运行期 429（第一版示例把 `DemoDerived.cls` 忘在 `.vbp` 外就踩到了；`.vbp` 的 `Class=`/`Module=` 清单是唯一事实面）。验收：四条新用例进 vbp 回归 —— `cc_demo_exe[_x86]`（`Test-Vbp`，13 条 needle）+ `cc_demo_dll_external[_x86]`（`Test-ComActivateClient`：注册 → 跑外部客户 → 反注册，并断言 `CLEAN_CLSID/CLEAN_PROGID/CLEAN_TYPELIB=gone`），本机 `.build/b18_cases.ps1` 四条 `pass=4 fail=0`，跑完实测注册表里 DemoDll 与 CoDll 的 CLSID/ProgID/TypeLib 三类键**全部查无**（可重复跑、不脏机器）。收线 exe md5 `38f247e30ed5ac3513366c9733bf4dda`（= BASE）。
- 2026-09-25 11:05–11:30 **同步 gitcode：`fan/dev` 推送 + MR !53 合并进 `main`（门 = Actions run #70 全绿，head `e38fbeb`）**：gitcode 侧 `fan/dev` 停在 B01 时代（`4fb3506`）而 `main` 有 14 个这条线没有的提交（Fix 188/195/196/197 + `!52 merge ferock/0.10.7`）⇒ 先把 `fan/dev` 推上去（`4fb3506..f785fda`，服务端钩子 PASSED），再合 main。**冲突只有一处文件** `src/rtl/core/vb6comserver/vb6comserver_obj.c`，两段都解在语义并集上：① `ComObj_Release` 归零分支 = **Fix 188 的 `ownsInstance` 总闸 + B16 的 claim 分叉**（借用型包装只清 `__comObj` 回填；拥有型里有 claim 的走 `instanceClaimRelease`、其余走 `destroyFunc`）；② `ComObj_Invoke` 实参打包 **取 main 侧的 `vb6_VarSlot`(24B) 槽设计** —— 它每槽 `memset` 清零 + “槽内是副本、资源归调用方、不得 VariantClear”，把 B17 那条堆损坏的根因整块拿掉，我的 `memset(coercedArgs, 0, argc*sizeof(VARIANT))` 既多余、在 24B 槽下 x86 又量错，故只留一份。合并后本机验：示例 4/4、外部激活 3/3、接口/继承/CoClass 六条 + `itf_xmod`/`test_implements` 全绿；然后 `e38fbeb` 推 `github/dev` 跑全量门 = run **#70 8/8 job 绿、`Tests (vbp)` 43/0/1/44**（与合并前的 #68 同读数）。MR !53 已合并（`merged:true`），gitcode `main` 推进到 `68fb713d`；`github/dev` 留在已验证的 `e38fbeb`。
- 2026-09-25 12:20–14:52 **B19（用户报障）出完并过门（Actions run #73 全绿，head `8174219`）**：用户先报"在 cmd 里用 C3 编译，报错信息乱码"，再追问"cmd 支持 UTF-16 输出吗"、"PowerShell 5.1 / pwsh 里是否也正常"、"**要走通用的，不能光针对中文环境，韩/日/英都要通用**"。三问都先量再改：① **量**（读数进 D66-2）—— `cmd` 里 std 是**控制台句柄**、两代 PowerShell 里都是**管道**（PowerShell 接走 native 输出）；两代 PS 的 `[Console]::OutputEncoding` 默认 = **启动时**的控制台代码页（本机 936）；**子进程改代码页对 shell 的解码没有影响**（实测两组码点一字不变）；同一份四语种输出在控制台代码页 936/65001/437 三档下**逐字相同**（`WriteConsoleW` 交 UTF-16，与 `chcp` 无关）。② **改**（`796220d`）—— 编译器 `main.cpp` 的 `ConsoleCodePageGuard`（判 `GetConsoleWindow()`、切的是**整个控制台**；ConPTY 下根本不触发 ⇒ 用户踩的就是它；且会把 C3 拉起的 `cl/link` 的 GBK 输出反过来弄乱）换成 `ConsoleUtf8Buf`（装到 `std::cout`/`std::cerr`，**不碰用户代码页**）；运行库 `vb6rtl_conv.c` 新增 `vb6_ConWriteHandle/OutW/ErrW`，`Debug.Print` 一族与"未处理运行期错误"全改走它（控制台 → UTF-8→UTF-16→`WriteConsoleW`；管道/文件 → 按**当时的控制台代码页**写字节、整行装不下就退 UTF-8，判据是 `lpUsedDefaultChar`，不硬编码任何语言/代码页）。③ **判据** —— 新探针 `tests\tools\con_capture.c`（**真控制台**里跑命令、**读屏幕缓冲**回 UTF-8；重定向只能量字节、量不到控制台那条路）+ 助手 `tests\console_enc.ps1` 三条进 vbp（`cc_cn_console` / `cc_cn_console_x86` 读屏断言 + `cc_cn_redirect_gbk` 按**字节**断言），源是**四语种**夹具 `tests\cc_cn\CnMain.bas`（UTF-8 BOM，GBK 存不下韩文）；手册 `Debug 对象` 页记下"任意语言通用"的姿势（先 `chcp 65001` 再起 shell ⇒ 中/日/韩/英同时成立）。④ **门后**：CI bas#1 连报两轮 `run timeout 5s`（#71 连 `test_rtl` 一起挂）—— 其中 `test_rtl` 那条是**真挂**，修在 `d4e53c2`（未处理错误的出口旧判据在 std 被重定向时会弹模态框把自己卡死：实测改前 5024ms 被超时杀、stderr 空；改后 25ms + `Unhandled VB6 Error #429: …` 落到 stderr）；剩下的两条是**饿**（读数与判据进 **D67**：12 路并行 `cl/link` 把 4 vCPU runner（本机同为 4 核）压满，健康小 exe 也 >5s 才跑完；本机 4 核复现"同一二进制时而 22/24 时而 24/24"，卡住的进程 4 秒只烧 15ms CPU、线程 `Ready`）⇒ 套件运行预算 5s→60s（`-RunTimeoutSec`）+ 超时读数（CPU/状态/最后一行输出）（`8174219`）。门重跑 = Actions run #73 全绿 全绿（8/8 job；bas#1 24/0/0、vbp 46/0/1/47，含三条新 `[ENC]` 用例）。

- 2026-09-25 19:10–21:55 **B20（ai/028 两批：多行字符串 + 字符串插值）出完并过门（Actions run #82，head `c2f7317`，8/8 job 全绿）**：用户先要求"先写计划书"（`ai/028`，写之前把锚点全部实测复核过），拍板后按 V1→V2 开工。
  **开工第一件不是写码，是把 `github/dev` 合回来**（本地 ahead 1 / behind 7：他人 Fix 190/194–197 已在 CI 分支上，门 #81 后来也绿）——合并基线 `8174219`，两侧改动不相交 ⇒ 无冲突，且合并树相对 `github/dev` 只差 `ai/022` 一个文件 ⇒ 代码面与已被门验证的 `a93d97e` 逐字节相同，合并本身不再跑门。**本轮用户新定调：没有源码级变动就不要跑 Actions**（"不然纯属浪费"），已写进流程与记忆。
  **V1（`4c16f50` + `f8ec76e`）**：`scanRawString` 在词法出口归一（行界 CRLF / 串内 `"` 重新双写 / 反引号双写折叠 / 零转义），产物就是普通 `StringLiteral` ⇒ 计划书 R4 那四处独立折叠（cgen、Const 注册、`CreateObject` 的 `unquote`、`Declare` 的 `stripQuotes`）一行都不用改，且四个落点各有一条真判据。**实测推翻原稿两条**：定界计数与 VB6 同规（两枚 = 空串、四枚 = 一枚字面反引号、三枚 = 未闭合），以及诊断文案必须 ASCII（D12 硬约束，原稿写"与邻居同调的中文"）。编译器侧净改动 `+71/-0`，全在词法层。
  **V2（`c2f7317`）**：降级点从计划书默认的 parser 路线换成**词法展开**（三条实测依据见 D69），于是 token kind / parser 两处字面量入口一处都没动；孔内表达式交给一个 `[begin,end)` 窗口的子 `Lexer`，行列用现成的 `getLocation` 播种 ⇒ 第二个孔的报错落在 `(8,7)`，与手写同形报的是同一条 `VB3001`；发码里看到 `vb6_BSTR_Concat(vb6_BSTR_FromStr(L"n="), vb6_CStrLong(n))` ⇒ 插值真的降成了真记号链。诊断只加两条（`VB1008`/`VB1009`），原稿第三条 `2017 ParseInterpTrailingJunk` 作废——孔里的坏形状交给真 parser 报既有语法错。
  **路上三次自己抓自己**（全在 D69）：① `content_` 是 `string_view`，写成 `buffer_->content().substr(...)` ⇒ 临时串悬垂，**所有**文件在 `(1,1)` 报"意外字符"（连存量用例全红，一眼看出是地基被弄坏而不是特性问题）；② 格式段沿用表达式位的扫描 ⇒ `${n:#,##0}` 的 `#` 被当日期面量一路吞到行尾（误报 1008）；③ 四份编码变体登记时用 `"..." + $v + "..."` 直接写在参数位置 ⇒ PowerShell 把 `+` 当独立实参 ⇒ 四条假红 —— **这一条是本地 bas 跑抓到的**，门之前就修好并单独验过路径形状。
  **仓库级坑（值得所有后续批次知道）**：本机 `core.autocrlf=true` 把 `*.bas` 统统按 LF 存进索引（`git ls-files --eol` 显示 `i/lf`）⇒"同一内容存成 CRLF / LF 两份源"这类判据在 commit 之后两份塌成一份，用例在 CI 上静默退化、看着绿其实没量。修法：给字节参与断言的夹具加 `-text` + `git add --renormalize`，再核 `git cat-file blob` 的 CR 字节数（0/22/0/22 才叫真分开）——见 `f8ec76e`，与仓库里 `.frx` 那段注释讲的是同一件事。
  **护栏与门**：`.build\b20_emitc_guard.py` 量 BASE = `pre_b20_C3.exe`（合并后、反引号之前）vs NEW ⇒ **23 件存量工程 `--emit-c` 逐字节全同**；反向断言 **4/4**（含反引号的源在 BASE 必失败、NEW 必成功）= 护栏能红的证明。门 = run #82，8 个 job 全绿；CI 分类逐条读数这一轮取不到（PAT 无 `actions:read`，job 日志端点 403），已在 GATE_BASELINE 里写明并按 run 级记账。手册 `String 数据类型` 页加两节，标明是本项目扩展、真 VB6 编不过（广告==应答）。

- 2026-09-25 22:05–22:20 **同步 gitcode：`fan/dev` 推送 + MR !54 合并进 `main`**（门 = Actions run #82 全绿，head `c2f7317`）：先把 `origin/main`（MR !53 的服务端合并提交 `68fb713`）合回 `fan/dev` —— 无冲突（内容早已在 `e38fbeb` 那一支里合过），合并树 = 门验证过的那棵树；推送前专门核了一次"相对 main 删掉了什么"，因为删除行集中在 `vb6rtl_file.c` / `driver/main.cpp` / `vb6forms.c` 这些**别人也在改**的文件上：实测方向是反的 —— main 上还是 `vb6rtl_file.c` 的 482 行旧版（= `e38fbeb` 那份 md5 `8a00315f`），我这边是 Fix 197 的 662 行新版（md5 `180c3dce`，就是 `github/dev` 上 `c905707` 那一支）⇒ 这次 MR 是**推进** main，不是覆盖别人的新工作。随后 `fan/dev` 推上去（`4675438..3304236`，服务端钩子 PASSED）→ 建 MR !54（84 文件 +4030/−344）→ 按常设默认（`mergeable: true` 且 `base.sha` == 实时 `main` 才合，不再问）合并 → `main` 移到 `696733d`「!54 merge fan/dev into main」。**合完核对**：`git diff origin/main HEAD` 空、两侧 tree 同一枚 oid ⇒ main 现在就是被门验证过的那棵树。台账本身是文档，按本轮新定调**不推 CI**。
- 2026-09-25 22:26–23:20 **B21（待拍板 5 收口 = 布尔的两套读数）出完**：开工先做 BASE —— 工作树里已有本批前一半（可见性登记）的未提交改动，`cp` 现成 exe 当基线会让护栏少测一半 ⇒ `git diff > .build\w1_full.patch`、`git checkout --` 那 13 个文件回 HEAD、构建 `pre_b21_C3.exe`（md5 `70f5b98d`）、`git apply` 回灌、再构建 NEW（md5 `9456577b`，收线快照留在 `.build\b21_C3.exe`）。共享树里不用 stash，导 patch 的窗口约 4 分钟、只碰自己的文件。
  **两条根因**：① 可见性（`As Boolean` 与 `Integer` 同为 `int16_t`，登记表按 C 类型串分派 ⇒ `inferExprType` 看不见布尔，`CStr` / `&` / `String` 形参 / 裸值 `Debug.Print` 全落整数分支）—— 加 `knownBoolVars_`，五个登记点各补一处、消费点先判它；② 装箱（`_Generic vb6_VariantFromValue` 把 `int16_t` 送到 `short:` → `VT_I2`、把布尔字面量的裸 `(-1)` 送到 `int:` → `VT_I4`）—— 装箱点收进 `boxToVariant`，只有推断为布尔才换 `vb6_VariantBool`。RTL 侧对 `VT_BOOL` 本来就格式化得对，所以本批全在 cgen。
  **护栏 RED 一次、按读数改口径**：助手第一版写成“实参已是 VARIANT 就直传”，30 件工程对照当场差 2 行 —— VbQRCodegen 的`vb6_VarType(vb6_VariantFromValue(VB6_SA_AT(...)))` 少了一层**恒等**包装（语义等价，但那是存量可观察产物）⇒ 改成“非布尔一律原样回退”，对照回到只剩 4 行真该变的（全在 `test_types.bas` 的 `Debug.Print` 布尔那几句，`-1` → `True`）。分类判据脚本 `.build\b21_emitc_guard.py`（清单 = B20 那 23 件 + 7 件含 `As Boolean` 的存量源）。
  **用例与读数**：`tests\test_bool_display.bas` 32 条（含 `Integer`/`Long`/`Byte` 装箱读数的反向护栏、`${b}` 插值、裸值 `Debug.Print`）+ 登记 `test_bool_display[_x86]`；负控 = 同一份用例喂 BASE 二进制 ⇒ B1-B20 与 B30-B32 一起翻红（`B31-raw-1`、`B32-raw0`）；两条用例走真 harness（`-Category bas -BasShard 22/60 -BasShardTotal 100`）各 PASS；x64/x86 各 32/32；`-Category syntax` 129/0；`test_types.bas` 现打 `Boolean=True`/`FalseVal=False`/`NotTrue=False`。手册 `Boolean 数据类型` 页加“本项目的实现口径”一节（含未对齐的那条边界）。
  **顺带量出、未动手**：`Print #` / `Write #` 的非 BSTR 实参不分类型一律 `vb6_Str((int32_t)x)`\u21d2 布尔落盘 `-1`、`3.5` 截成 `3`、`Write #1, True` 写成 `"-1"`（VB6 依次是 True、空格 3.5、`#TRUE#`）—— 记为待拍板 7。
  **门 = Actions run #83**（head `2ddcc8b`，23:05 到 23:19，14.6 分钟）**8/8 job 全绿** = Build C3.exe + Tests(smoke / syntax / vbp / compile / asm / bas#1 / bas#2)。这一轮的门读数**逐 job 核过**：PAT 无 `actions:read` 仍旧挡住 job 日志端点（403），但 `GET /repos/…/actions/runs/36151881443/jobs` 不要那个权限域，于是 run 级一把绿升级成八条 `status/conclusion` + `head_sha` 对表 —— 以后都按这条取数，别再拿 run 级当结论。
  **B22（= 待拍板 6）本轮只量不做**，复现缩到最小形（`New` 一个工程内类交给 `As Object` 就会空，不需要 `CreateObject`），两条候选根因的分家、探针位置与两条工具坑全在 **D71**。
  **本会话撞到的一条**：一次工具返回里夹了段伪装成“security notice”的提示注入，要我去 `curl` 一个本机监听端口并把它称作“默认目标”。没执行（本机那条端口探测根本没有任何 LISTENING，日志里也没有那个端口号），只登记不照办。
- 2026-09-26 00:05-00:30 （用户指定，另起一线）**内置控件补全：只做测量与计划书，未动一行代码** —— 计划书 = `ai\029-内置控件补全计划书.md`（11 份 `ai\内置控件\*.md` 说明逐条定性；走 comctl32 「原生窗口类」，不引 MSCOMCTL.OCX；Data 与 OLE 容器明确不做并写死理由）。三条实测把这条线的形状定下来了：
  ① **Shape / Line / Drive / Dir / File 这五个是「写好一半、创建那一刀没接上」** —— `controlTypeToWin32Class`（`src\project\frm_parser_util.cpp:184`）对它们全部 `default: return nullptr`，于是 `cgen_form_create_controls.inc:436` 一路落到 `:770` 的「不可见控件跳过」；而 RTL 那半边是真在的（`vb6forms_shape.c:172` 自注册 `VB6_SHAPE`/`VB6_LINE`、`cgen_util_ctrl.cpp:151/307` 属性表、`vb6forms_webview.c:196` 盘/目录/文件填充）。实测法 = 在 `tests\test_form\Form1.frm` 上加这五个控件再 `--emit-c`（`.build\c29_out\m.c`）：五个 `vb6_hwnd_xxx` 声明都在、`vb6_RegisterShapeLineClasses` 也发了，但 `vb6_CreateControl(` 仍旧只有 3 次 ⇒ **句柄永远 NULL、屏幕上什么都没有**。⇒ 本线第一批（C29-1）就是接这一刀，最便宜。
  ② **全仓库零 `WM_NOTIFY`**（`WM_NOTIFY` / `NM_` / `LVN_` / `TVN_` / `TBN_` 四类关键字零命中）⇒ ListView / TreeView / Toolbar / StatusBar / SSTab 这五个通知型控件共用一条还没建的通道，硬前置排在 C29-4。
  ③ 两条「顺手就答掉」的待量：**产物不带 application manifest**（`Microsoft.Windows.Common-Controls` 全仓零命中 ⇒ 原生控件只能拿 comctl32 v5.82 行为，是否补一份 `.rc` manifest 变成决策 D5，建议独立成 C29-M 不混进控件批次）；**`For Each` 今天只吃数组**（`cgen_control.cpp:151` 起三类判定，`As Collection` 有声明没有枚举）⇒ 集合对象的 `For Each n In TreeView1.Nodes` 用「按显示序返回对象数组快照」就能通，不必做 `_NewEnum`（顺带记下「Collection 也不能 For Each」这条既有缺口）。
  **要用户拍的板共五条**（D1 成员对象表示法 / D2 事件参数类型名收不收 `MSComctlLib.` 前缀 / D3 GUI 判据手法与探针入册 / D4 Data + OLE 不做 / D5 comctl 版本），批次表 C29-0..C29-9 一批一扇门。本轮零源码改动 ⇒ 按「没源码不进 CI」不推 `github/dev`。
- 2026-09-26 01:05–02:40 （用户重申目标后开工）**029 内置控件线第一批 C29-1a 出完并过门 = 代码 `5b37b4e`，门 Actions run #84（8/8 job 全绿）**：把「RTL 写好一半、创建那一刀没接上」的 Shape / Line 接上 —— `controlTypeToWin32Class` 缺那两格 ⇒ 控件一路落到「不可见控件跳过」，句柄永远 NULL，屏幕上什么都没有而且一声不响（取证法：往 `tests\test_form\Form1.frm` 上摆五个这类控件再 `--emit-c`，`vb6_CreateControl(` 仍旧只有原有那 3 次）。开工先按 022/B21 那条口径取 BASE（工作树当时已含本批一部分改动 ⇒ 导 patch、回退自己那几个文件、构建 BASE、回灌、构建 NEW），BASE 复用 `b21_C3.exe`（md5 9456577b = HEAD 那一版）。
  **同批堵掉的三条同族洞**（都是「两套代码各写一遍、只有一套更新」）：① 顶层控件与容器子控件是两条创建路，子控件那条一条外观属性都不发；控件数组（`lamp(0)`/`lamp(1)`，Shape 手册里「指示灯」就是这种用法）的设计期初始化硬写 `vb6_hwnd_<名>`，会打到空句柄上 ⇒ 抽成 `emitShapeLineProps` + `lineRectFromEndpoints` + `ctrlHwndExprForInit` 两条路共用；② `ctrl.Left` 被 VB 内置函数名 `Left` 带偏成 String（Fix 081i 给 UDT 记过同一个坑），于是 `If Line1.Left = 600` 生成 `vb6_StrCmp(整数, BSTR)` ⇒ 把 600 当指针解引用，**实测就是段错误**，不是读数偏差；③ `borderstyle` 在属性表里被通用分支先抢走（通用那条查在类型 switch 之前）⇒ `Line1.BorderStyle` 读的是窗口边框样式(0/1)而不是画笔线型(0..6)。另两条纯运行时的：`SetPropW` 存 `(HANDLE)0` 与「从没设过」不可分辨（`FillStyle=0`/`BorderStyle=0` 读回变默认 1）⇒ 三处改存 val+1；旧的 `if (borderS <= 0) borderS = 1` 让 `PS_NULL` 分支成死码，一并去掉。Line 的 X1/Y1/X2/Y2 定口径为**容器缇值**、窗口矩形 = 四端点包围盒、赋端点连窗口一起搬。
  **判据与护栏**：`tests\ctrlshape\`（CsApp.vbp + CsForm.frm，21 条读数：几何落位 / 设计期整数属性 / 运行期读写回路 / 端点搬窗口 / BorderWidth 强制实线 / 容器子控件 / 控件数组按槽位），窗体 `Form_Load` 打完 `Unload Me` 自退 ⇒ 走现成 `Test-Vbp` 拿 stdout 针，登记 `ctrlshape[_x86]`；x64 与 x86 各 21/21，负控喂 BASE 二进制 21 条全翻红，`.build\c29_emitc_guard.py` 30 件存量工程 `--emit-c` 与 BASE 逐字节全同（changed_lines=0）。手册 `Shape 控件` / `Line 控件` 两页各加一节「本项目的实现口径」（缇与像素的分界、无事件仍然成立、以及三项没落地的属性）。029 §九记全量读数。
  **两处本轮踩到的工具/流程坑（都已写进记忆与 029）**：① **RTL 的 .c 改动必须重建 C3.exe 才生效** —— RTL 是以资源嵌在编译器里的（`src\driver\rtl_embedded.hpp` + `c3rtl.rc`，链接期解出到临时目录再编），只改 RTL 就跑测试读到的是旧运行时，症状正好是「代码写了不出现」；② 自己起的本地 `-Category vbp` 把 `C3.exe` 锁住 ⇒ 随后的 `build.bat` 撞 `LNK1168`（exit 1168），按纪律等套件跑完再构建，没动任何进程。另外记一条**本机 vbp 噪声**：`-Category vbp` 本机跑 PASS=42 FAIL=15，那 15 条是探针/读手没建成加两条要注册表/非 ASCII 控制台的用例，灌过 vcvars 的第二次跑一字不差复现，而 CI 的 vbp job 全绿 ⇒ 门读数一律以 CI 为准。
  **计划书被实测推翻的一条**：C29-0（外部子窗口探针入册）第一批就判定**不必做** —— 「程序自己读几何属性 + 打完自退出」这条判据既够强（BASE 下 21 条全红）又零新基础设施。同时读出两条留给后续的事实：控件 `.hwnd` 不是数值面（`void*` 装箱成 `VT_DISPATCH`，`<> 0` 恒不成立），以及 `As Collection` 能 `Add`/`Count` 但**不能 `For Each`**（`For Each` 只吃数组）—— 后者正是 §三 D1 选 (a) 时「集合靠数组快照可枚举」的第二个理由。下一批 = C29-1b（文件系统三控件），开工地图已写进 029 §九。

- 2026-09-26 03:05–04:10 （同一目标续跑）**029 内置控件线第二批 C29-1b 出完并过门 = 代码 `440129c`，门 Actions run #85（head `440129c`，8/8 job 全绿，18.6 分钟）**：把文件系统三控件接上原生窗口 —— Drive = `COMBOBOX`(`CBS_DROPDOWNLIST`)、Dir / File = `LISTBOX`(`LBS_NOTIFY`)，同一批补上样式、设计期初值（`.frm` 的 `Drive`/`Path`/`Pattern` 以前**根本读不到**，那里硬写 `GetCurrentDirectoryW`）、属性读写表、`WM_COMMAND` 三条派发（Drive `Change`/`Click`、Dir `Change`(下钻后)/`Click`、File `Click`/`DblClick`）。RTL 那半边补两条 VB6 口径：Dir 的每一项列成 `[名字]`（原来是裸名字，既看不出层级也没法下钻）+ 新增 `vb6_DirListBoxDescendSelected`（双击下钻，`Path` 真变了才回发 `Change`）；`Drive` 出口去掉尾反斜杠（列表项仍是 `C:\`，属性读数是 `C:`）。
  **本批最值钱的一条是"定义有、声明无"**：那 12 个 RTL 入口躺了很久但任何 `.h` 里都没有声明 —— 因为三控件在 `controlTypeToWin32Class` 缺格 ⇒ 句柄永远 NULL ⇒ 从没有调用点。接上第一枚真句柄，生成代码按"返回 `int`"的隐式原型编译、字符串句柄被截成 32 位，**真跑段错误在 OLEAUT32**（`--emit-c` 与 syntax 分类都看不出，只有真编译真跑能暴露）。`vb6forms_*.h` 不引 `oleauto.h`，故声明一律写 `wchar_t*`（`OLECHAR` = `wchar_t`，与 `.c` 里的 `BSTR` 同一类型、不冲突）。
  **另两条同族洞**：① 可索引属性 `List(i)` 的发码点在 `cgen_expr_call_callee_withm.inc`(P13.3b) 而**不在**成员访问侧(P13.3)，两处各自硬编码 `{ListBox, ComboBox}` —— 只放宽一处照旧 `C2198: vb6_GetListItem 用于调用的参数太少`（发成 `vb6_GetListItem(hwnd)(0)`）；② 控件的字符串属性没登记进 `inferExprType` ⇒ `File1.FileName = File1.List(0)` 两侧判成 Variant 走 `vb6_VarCmpEq`，而 `vb6_GetListItem` 声明是 `void*` ⇒ 装箱落 `default:` 成 `VT_UNKNOWN` —— **同一条读数 x64 为真、x86 为假**。中途先在用例里绕（赋值给一枚 String 临时变量，两架构都绿），撤了：那是把编译器的问题挪进测试。修在推断层（C29-1a 加 `kNumericFc` 的同一处）。`Fix 194` 的注释早就点过这个 `void*` 坑，但只补了 `CStr` 那一路。
  **未修、记下的一条既有缺口**：`Left(字符串, n)` 写在**窗体模块**里被控件属性抢走（编成 `vb6_GetControlLeft(vb6_hwnd_..., ..., 1)`），同一句在 `.bas` 里正常 ⇒ 与本线无关，本批判据改用 `Mid(s,1,1)` 绕开。
  **判据与护栏**：`tests\ctrlfiles\`（CfApp.vbp + CfForm.frm，14 条读数：三控件都有窗口且填进去过 / `[名字]` 约定 / 设计期 `Path`+`Pattern` / 改 `Pattern` 立刻重刷 / `ListIndex`↔`FileName` 回路 / 目录→文件与盘→目录两条联动 / 与原生 `ListBox` 口径一致），登记 `ctrlfiles[_x86]`；x64 与 x86 各 14/14，负控喂 BASE(`b21_C3.exe`) 12 条翻红（另两条只走原生 `ListBox`，本来就不测本批能力）。通知接线在无头环境点不了 ⇒ 另登记 `cf_emitc_shape` 断**发码形状**。30 件存量工程 `--emit-c` 逐字节全同（BASE 一颗盖住 C29-1a + C29-1b 两批，changed_lines=0）、`-Category syntax` 129/0、本机 `-Category vbp` PASS=45（比上一版 42 恰好多本批新登记的三条），红的 15 条是门前沿用同一批环境噪声。手册三页各加一节「本项目的实现口径」。029 §九记全量读数、批次表与状态头同步。
- 2026-09-26 04:12–04:45 （用户给的两个选项里选 CommonDialog）**本轮零源码改动**：C29-2 移交后换批做 C29-9， 按 §八 的"先测量再动码"量了五件事，读数全在 029 §九 那一格。要点三条：
  ① **CommonDialog 今天是静默空转**：它早就挂在"ActiveX 按 COM 后期绑定"那一组（`CoCreateInstance` 真 OCX `MSComDlg.CommonDialog` + `vb6_ComCall`/`vb6_ComGetStringProp`）。本机 HKCR 有登记但 OCX 只有 32 位 ⇒ x64 建不出来，探针六条读数（`DialogTitle`/`Filter`/`Flags`/`CancelError`/`ShowOpen` 后 `FileName`/`ShowColor` 后 `Color`）**全空**、`ShowOpen` 不出现、程序照旧打 `PROBE-DONE` 且退出码 0。这答掉了 §八-V3 那条待量（"CI 上绿是真空转还是静默不存在也绿" ⇒ 后者）。设计期那四行也一行都没落到控件；`Flags`(数值)/`CancelError`(布尔) 还错走字符串取器，`CD1.Color = 255` 先报一条 VB4001 警告。
  ② **comdlg32 不是新依赖**：DI 层已有 `vb6_di_GetOpenFileNameW`/`vb6_di_ChooseColorA` 一族，`vb6_di_com_stubs.c` 里 `#pragma comment(lib, "comdlg32.lib")` 也在；按同文件 GDI+ 那条"flat API 走 GetProcAddress"的现例实现 ⇒ §六-3 的工具链依赖面不动。
  ③ **卡住的是判据不是实现**：D3 那条"模态框自己起 Timer 自关"走不通，根因在更前面 —— **`VB.Timer` 运行期根本不触发**，两条独立成因（都实测）：注册那一刀压根没发（`cgen_form_wndproc_create.inc` 那段循环遍历 `frmDesc.formControl.children`，窗体的 Timer 不在里面 ⇒ 产物里一个 `vb6_SetTimer(` 都没有，而 `case WM_TIMER: vb6_DispatchTimer(wParam)` 那半边是好的）；Timer 是无窗口控件 ⇒ `vb6_hwnd_<timer>` 恒 NULL，`Enabled`/`Interval` 写进 `SetPropW(NULL, ...)` 静默丢。探针读数：1.2 秒 DoEvents 轮询 ticks=0（删掉设计期 `Enabled` 行照旧 0）。
  **⇒ 本轮停在一条要用户点头的岔口**：建议先立一小批 **C29-T（Timer 运行期触发）**——它本身是内置控件线上的一条真实缺陷，且 C29-9 六个 `Show*` 的自关闭探针依赖它；通了再动 CommonDialog 本体。按纪律 STATUS 保持 BUSY、不收线（本轮没有可过门的源码改动），下一轮第一件事 = 按用户对岔口的答复开工。探针工程留在 `.build\c299_probe\`、`.build\c299_ac\`，不进回归。
- 2026-09-26 04:12–05:45 **029 线两件事**：先把 C29-2（ProgressBar）**移交另一位作者**（用户指定；本线那 8 个文件的半成品从共享树退干净、判据工程与 WIP 留在 `.build\c292_...` 下、重建过 C3.exe），期间用户两次定调 => 新决策 **D6：内置控件一律原生实现，默认路绝不走外部 OCX**（理由 = MSCOMCTL/MSComDlg 这些 OCX 只有 32 位，x64 里 `CoCreateInstance` 直接失败，"门绿"跟"控件在"是两件事）。
  然后按 D6 做 **C29-9 CommonDialog = 已出并过门（`e7f3352`，门 run #86 8/8 全绿）**。四处 OCX 形状一起摘（变量声明 / `CoCreateInstance` / 成员访问 COM 派发 / **无括号方法的语句路**，最后一处是照着 `List1.Clear` 那个先例的位置接的，只改 callee 侧不够）；属性宿主换成自注册的不可见子窗口 `VB6_COMMONDIALOG`（0x0、清 WS_VISIBLE），这样 cgen 的 `readFn(hwnd)`/`writeFn(hwnd,v)` 形状不用特判，`GetParent` 顺手就是模态父窗；六个 `Show*` 直调 comdlg32 且只经 `LoadLibrary`+`GetProcAddress`（照 GDI+ 那条例子）⇒ import lib 依赖面零增长。`Filter` 对外仍是竖线串，原生 `描述\0模式\0…` 那张表只在 `Show*` 那一刻折；取消按 `CancelError` 报 32755 且不改已有读数。
  **同批堵掉一条通用属性抢占**（第三次撞到"两套名单只更新一套"这一族）：CommonDialog 的 `FontName`/`FontSize` 是 `ChooseFont` 字段，而通用那族 `vb6_SetControlFontName` 查在类型 switch **之前**，不挡就把赋值静默落到控件字体上、读回空。另按 `SetPropW` 存 0 那条老坑给布尔 normalize（VB6 `True` = -1，直存 `val+1` 会变 0 跟"从没设过"撞车）。
  **判据**：`tests\ctrldlg`（10 条读数 x64/x86 各 10/10，恒假守卫把六个 `Show*` 留在源码里——真弹框的判据另立 C29-9b，否则用例会在没人点"取消"的地方把门卡死）；发码两面都钉：`dl_emitc_shape` 断原生入口在、`dl_emitc_no_ocx` 断 `CoCreateInstance`/`vb6_com_<名>` 不再在（为此新加助手 `Test-EmitcAbsent`，并按纪律双向验过：真针 PASS、在场形状报红）。负控喂 BASE 直接编不过（那条路上没有属性宿主），"改之前的症状"由本轮开工测量给（探针六条读数全空、退出码 0）。护栏：syntax 129/0、30 件工程逐字节全同。
  **本轮工具/流程踩坑（已写进记忆）**：python 里 `"ai\029-..."` 会被当 `\02` 八进制转义吃掉，落盘成一个裸控制字符 + `9-`（本轮在 022 状态头里真发生了一次，已修并全文扫控制字符）→ 写档脚本里的反斜杠一律 `chr(92)` 拼；`io.open(p).read()` 是文本模式会把 CRLF 折成 LF，按行尾改写必须 `'rb'` + 手工 split `chr(13)+chr(10)`；给 PS 5.1 跑的临时脚本要么 ASCII-only 要么带 BOM。
- 2026-09-26 05:44–06:45 **029 线 C29-T 出完并过门 + 一次"超时"归因**：
  ① C29-9 收线（门 #86，head `e7f3352`，8/8 全绿）后，用户追加一句"把 timer 的精度做高一点"，于是把 C29-9 测量时撞见的那条独立缺陷正式立成 **C29-T** 并做完：Timer 换成自注册的不可见窗口 `VB6_TIMER` 当身份（与 C29-9 的 `VB6_COMMONDIALOG` 同法，cgen 的属性读写形状不用特判）、注册不再看设计期 `Enabled` 的脸色（有事件处理器就挂表，设计期值只决定起不起）、底层从 `SetTimer`（~15.6 ms 地板）换成 winmm `timeSetEvent(wResolution=1)`，到期回调只把消息投回窗体、仍走原 `case WM_TIMER` 派发口。winmm 经 `LoadLibrary`+`GetProcAddress` 取 → 不新增 import lib。
  ② 两条老坑各撞一次，值得记：布尔 `Enabled` 存 `val+1` 时 VB6 的 True=-1 会变成 0（与"从没设过"不可分辨）→ 单独 normalize；**`TIME_PERIODIC` 手抄成 0x02（真值 1）会让 `timeSetEvent` 直接失败并静默退回 SetTimer** —— 不报错、功能也对，只是"精度没上去"，极难归因 → 改成显式引 `mmsystem.h`（它只是被 `WIN32_LEAN_AND_MEAN` 排除）。判据因此必须带量化区间：20 ms 名义 50 给 [40,60]，`Interval=5` 名义 200 而地板只有 ~64（阈值 100 才分得开）。
  ③ 推送后门 #88 全绿；同时 dev 上被触发了另一位作者的 workflow（"GitHub Tests T0+T1+T2" run #1），用户贴来三条超时（`test_ndarray_x86` / `test_compat` / `test_types_x86`，都是 `run timeout 5s`）。本地按同一口径复跑（x64+x86 各一遍，含针校验）：三条**全绿**，**运行阶段只有 84~1127 ms**，慢的是编译阶段（22~43 s）→ 结论是那条 5 s 预算在 `-Jobs 20` 的 4 核 runner 上被并发cl/link 饿死（正是主套件 B19 把它改成 60s 的原因，见 `8174219`），不是程序问题、也不是本批改慢。那份 `tests_github` 副本的清单与预算**按既有约定不去碰**（有人在修），本轮只在台账记归因。
### D72 ai/030 两批落地后的读数与两条新账（2026-09-26）

- **读数**：暖 obj store 下一次构建 2.8 s（冷编 22.0 s）；15 例 bas 分片一趟 282 s -> 36 s；
  同一入口 `dev.ps1 -TestCategory smoke` 带开关 7 s / `-NoObjCache` 21 s。整组 vbp 带开关
  `PASS=67 FAIL=0 SKIP=1 TOTAL=68`，与门 #91（不带开关）同口径逐字相同。
- **`-O 2` 那臂 RTL 仍 65/65 命中、用户码 0/2**：RTL 的编译档与用户优化档确实解耦了；
  `--arch x86` 独立一格，首编全 miss、再编全命中，6 条 x86 vbp 无回退链。
- **新账 #78**：`Test-Compile` 在 run_tests.ps1 里定义了两次（190 真构建 / 762 只做 --emit-c），
  后者覆盖前者 => `-Category compile` 根本不构建（5 s、10 例全 PASS、零产物）。本批没顺手改名，
  因为会动到 ai/028 那批的断言语义。
- **新账 #77**：`findClExe()` 返回字面量 `cl.exe`，所以 toolsetTag 在"非 dev shell + 没预灌 env"
  时退化成常量 `"cl.exe|"` => "cl 版本进键"这句目前没有内容撑着。随包带 obj（T30-D）之前必须补。
### D73 把 -Incremental 接进 GA 的三条读数（2026-09-26）

- **门确实变快**：#91（不带开关）wall 19m19s → #96（带开关）wall **12m10s，−37%**。分组：
  vbp 930→544 s、bas #1 344→219 s、bas #2 374→235 s、asm 114→83 s；smoke/compile/syntax 基本
  不动（后两组本就不构建 —— 见 #78）。没到本机那个 7.8× 是因为每个 job 仍付一次冷编 RTL
  (~20 s)，加上跑阶段/链接/20 路争 4 vCPU。
- **#94 的红是这批最值钱的一步**：`test_delegate_x86` LNK1120。根因不在缓存正确性，而在增量
  那条路**直接叫 link.exe**（编译/链接拆两步，`cl /link` 用不了），linker 按默认 /MD 配 msvcrt，
  与 obj 的 `LIBCMT`（x86 走 /MT，P24-09）打架，x86 独有的 `__except_handler4_common` 落空。
  `dumpbin /directives` 定位；修 = 镜像编译侧判据补 `/NODEFAULTLIB:msvcrt.lib`。这洞潜伏多久
  无从得知，因为 `--incremental` 在接进 GA 之前从没被任何回归跑过。
- **教训入账**：新路径"本地跑绿"不等于覆盖到 —— 我本地跑了 15 例 bas + 整组 vbp 都没碰到
  delegate 的 x86×msvcrt 组合。判据要靠让门去跑新路，不是靠扩大本地抽查。
