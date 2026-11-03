# check_app_title_authority.ps1 -- 账 #305 第320刀 (账 #592): App.Title 与 MsgBox 默认题头只问一格
#
# 病灶三条 (同一位主人的三个面, 改前各错各的):
#   ① 读: 发码侧把 `App.Title` 折成 `vb6_App_EXEName()` —— VB6 答 .vbp 的 Title;
#   ② 写: `App.Title = v` 拿上面那条折算当左值 => `vb6_App_EXEName() = ...` (cl C2106, 整棵工程编不过);
#   ③ 题头: MsgBox 省略 title 时递空串, VB6 递 App.Title (对话框少一栏字)。
# 口径 (唯一主人): 运行期真值只住 src/rtl/core/vb6rtl/vb6rtl_system.c 的 g_vb6AppTitle 一格。
#   生成的入口点在 vb6_Init() 之后按 .vbp 的 Title (空则 Name) 播种; 读/写各一处出口;
#   MsgBox 的默认题头借用同一格。未播种 (单文件编译没有 .vbp) 时读取退 EXEName —— 与改动前一致;
#   显式设成空串则答空串 (两档靠 g_vb6AppTitleSet 分开)。
#
# 规则 (每条都对应一个会让本哨兵红的假改动, 见 .local/k320/nc_320.py):
#   R1 三个出口 (setter/getter/borrow) 在 RTL 各只定义一次, 伞头各只声明一次;
#   R2 那一格的语义四件: setter 先退旧 / 标记"已播种" / 未播种退 EXEName / 显式空串答空;
#   R3 MsgBox 的默认题头借那一格 (省略 title 不再递空串);
#   R4 播种 lambda 定义一次, 每个入口模板各调一次 (调用数 == vb6_Init() 数);
#   R5 驱动按 Title 优先、exe 名兜底、最后才是工程 Name 的读法只有一处, 字段与 setter 各一枚;
#   R6 读侧折叠恰一处, 旧的 EXE 名形已撤 (EXEName 自己那一臂仍在);
#   R7 写侧拦截恰一处, 且局部同名优先 (knownLocalVars_ 同一判据);
#   R8 认这枚名字的 BSTR 返回名单一份都没漏 —— 第二步之后 Print#/Write# 那两份
#      **已删**（分流改问 inferExprType）；本刀把"App.Title 仍是字符串路径"改由
#      B600 运行断言兜住（R8b），名单总数随之 6 -> 4。
#   R9 夹具 / 期望 / 发码针 / 本哨兵自己都已在 run_tests.ps1 登记;
# 账 #305 第330刀 (账 #605) 又给这条线加了两腿 (VB6 实测在 ActiveX DLL 里 App 答的是这枚产物
# 自己 —— 宿主 = 32 位 cscript: P1 答本工程 Title、P3 答 DLL 模块名、P4 答 DLL 目录):
#   R10 播种那一站在 DLL 入口也得有 (Exe 五站之外), 且种子算法只许 appTitleSeed() 一处;
#   R11 App.Path / App.EXEName 问「这枚产物住在哪」只许一位主人, 不许再问宿主进程。
# 输出 ASCII (控制台 GBK, 中文读数重定向后不可 grep)。文件必须 UTF-8 BOM + CRLF。
#
# 用法: powershell -File scripts\check_app_title_authority.ps1 [-RepoRoot <树根>]   (PASS = exit 0)

param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$fail = @()
$SQ = [char]39

$sep = [IO.Path]::DirectorySeparatorChar
function Load([string]$rel) {
    $p = Join-Path $RepoRoot ($rel.Replace('/', $sep))
    if (-not (Test-Path -LiteralPath $p)) {
        $script:fail += ('missing file ' + $rel)
        return ''
    }
    return [IO.File]::ReadAllText($p)
}
function C([string]$lit, [string]$txt) {
    return ([regex]::Matches($txt, [regex]::Escape($lit))).Count
}
function Want([string]$rule, [int]$got, [int]$expect, [string]$why) {
    if ($got -ne $expect) {
        $script:fail += ($rule + ' = ' + $got + ' (want ' + $expect + '): ' + $why)
    }
}

$SYS  = Load 'src/rtl/core/vb6rtl/vb6rtl_system.c'
$CONV = Load 'src/rtl/core/vb6rtl/vb6rtl_conv.c'
$HDR  = Load 'src/rtl/core/vb6rtl/vb6rtl_builtin.h'
$ENT  = Load 'src/backend/detail/base/cgen_base_generate_entry.inc'
$MEM  = Load 'src/backend/detail/expr/cgen_expr_member_form_builtin.inc'
$ASG  = Load 'src/backend/stmt/cgen_assign.cpp'
$DRV  = Load 'src/driver/detail/driver_codegen_module_loop.inc'
$STATE = Load 'src/backend/detail/util/cgen_state.inc'
$API  = Load 'src/backend/detail/util/cgen_api.inc'
$RUN  = Load 'tests/run_tests.ps1'
$DI   = Load 'tests/disp_invoke.ps1'
$HPP  = Load 'src/driver/driver.hpp'
$DLLE = Load 'src/backend/detail/util/cgen_util_dllentry_exports.inc'
$DLLT = Load 'src/driver/detail/driver_codegen_dll_typelib.inc'
$FX1  = Load 'tests/b592_apptitle/APT.bas'
$FX2  = Load 'tests/b592_apptitle_fallback.bas'
$FX3  = Load 'tests/b592_apptitle/B592.vbp'
$FX4  = Load 'tests/b605_dllapp/B605k.vbp'
$FX5  = Load 'tests/b605_dllapp/Probe.cls'

# ---- R1 出口三枚: RTL 定义各一次, 伞头声明各一次 ----
Want 'R1 setter definition' (C 'void vb6_SetAppTitle(const wchar_t* title) {' $SYS) 1 'App.Title 的写侧出口只许有一处定义'
Want 'R1 getter definition' (C 'BSTR vb6_App_Title(void) {' $SYS) 1 'App.Title 的读侧出口只许有一处定义'
Want 'R1 borrow definition' (C 'const wchar_t* vb6_AppTitleBorrow(void) {' $SYS) 1 'MsgBox 借题头的出口只许有一处定义'
Want 'R1 setter declaration' (C 'void vb6_SetAppTitle(const wchar_t* title);' $HDR) 1 '伞头必须声明 setter (§B202: 生成代码里出现的名字都要有声明)'
Want 'R1 getter declaration' (C 'BSTR vb6_App_Title(void);' $HDR) 1 '伞头必须声明读出口'
Want 'R1 borrow declaration' (C 'const wchar_t* vb6_AppTitleBorrow(void);' $HDR) 1 '伞头必须声明借用出口'
if ($SYS.Length -lt 4000) {
    $fail += ('R1 scan floor: vb6rtl_system.c is ' + $SYS.Length + ' chars (want >=4000) -- 扫描瞎了却报干净')
}

# ---- R2 那一格的语义 ----
Want 'R2 cell storage' (C 'static BSTR g_vb6AppTitle = NULL;' $SYS) 1 '真值只住这一格 (另立一份 = 第二个主人)'
Want 'R2 seeded flag' (C 'static int  g_vb6AppTitleSet = 0;' $SYS) 1 '「从没播种」与「显式空串」必须分得开'
Want 'R2 setter frees old' (C 'if (g_vb6AppTitle) { SysFreeString(g_vb6AppTitle); g_vb6AppTitle = NULL; }' $SYS) 1 '改写题头要先退旧那一枚 (否则每改一次漏一枚 BSTR)'
Want 'R2 setter marks seeded' (C 'g_vb6AppTitleSet = 1;' $SYS) 1 'setter 必须留下「已播种」这一笔'
Want 'R2 empty leaves NULL' (C 'if (title && *title) g_vb6AppTitle = vb6_BSTR_FromStr(title);' $SYS) 1 '空串不许铸一枚空 BSTR 冒充「有值」'
Want 'R2 unseeded falls back' (C 'if (!g_vb6AppTitleSet) return vb6_App_EXEName();' $SYS) 1 '单文件编译 (无 .vbp) 时的旧口径只在「从没播种」这一档'
Want 'R2 explicit empty answers empty' (C 'return vb6_BSTR_FromStr(g_vb6AppTitle ? g_vb6AppTitle : L"");' $SYS) 1 '读出口交副本, 显式空串答空串'
Want 'R2 borrow hands the cell' (C 'return g_vb6AppTitle;' $SYS) 1 '借用出口只交指针, 不许顺手复制'

# ---- R3 MsgBox 的默认题头 ----
Want 'R3 MsgBox borrows cell' (C 'vb6_AppTitleBorrow()' $CONV) 1 '省略 title 时问那一格, 不许另立第二份真值'
Want 'R3 MsgBox default arm' (C 'title ? title : (dfltTitle ? dfltTitle : L"")' $CONV) 1 '显式 title 仍优先; 省略时落 App.Title, 两档都不许递空串兜底'
Want 'R3 MessageBox call kept' (C 'MessageBoxW(NULL, prompt ? prompt : L"",' $CONV) 1 '真正的弹框调用必须还在 (改回探针自己的 MessageBox 就不算修)'

# ---- R4 入口点播种 ----
Want 'R4 seed lambda defined once' (C 'auto emitAppTitleInit = [&]() {' $ENT) 1 '播种只许有一处定义'
Want 'R4 seed skips empty title' (C 'if (projectTitle_.empty()) return;' $ENT) 1 '无 .vbp (或标题为空) 时一行不发'
Want 'R4 seed emits setter once' (C 'c_.emitLine("vb6_SetAppTitle(L\"' $ENT) 1 '播种那一行只许拼一处'
$nInit = C 'c_.emitLine("vb6_Init();");' $ENT
$nSeed = C 'emitAppTitleInit();' $ENT
if ($nInit -lt 1) { $fail += 'R4 scan floor: no vb6_Init() line in the entry template (模板被搬走或改名了)' }
Want 'R4 seed called per entry' $nSeed $nInit '每个入口模板都要在 vb6_Init() 之后播种一次 (少一处 = 那条入口的 App.Title 永远是 EXE 名)'
# 账 #605: ActiveX DLL 那条路压根不进上面这五个入口模板 —— 它的入口是 dll_entry.c 的 DllMain。
Want 'R4 dll seed guard' (C 'if (!projectTitle_.empty()) {' $DLLE) 1 'DLL 入口的播种也要「种子为空就不发」那道闸 (无 .vbp 时零回归)'
Want 'R4 dll seed emits setter' (C 'entry.emitLine("vb6_SetAppTitle(L\"' $DLLE) 1 'DLL 入口只许拼一处播种行 (与 Exe 同形, 不另立第二枚 setter)'
Want 'R4 dll init kept' (C 'entry.emitLine("vb6_Init();");' $DLLE) 1 'DllMain 里 vb6_Init() 那一行必须在 (播种的相对位置靠它)'
$iInit = $DLLE.IndexOf('entry.emitLine("vb6_Init();");')
$iSeed = $DLLE.IndexOf('entry.emitLine("vb6_SetAppTitle(L\"')
$iDetach = $DLLE.IndexOf('DLL_PROCESS_DETACH')
if ($iInit -lt 0 -or $iSeed -lt 0 -or $iDetach -lt 0 -or -not ($iInit -lt $iSeed -and $iSeed -lt $iDetach)) {
    $fail += ('R4 dll seed order: init=' + $iInit + ' seed=' + $iSeed + ' detach=' + $iDetach +
              ' (播种必须发在 vb6_Init() 之后、PROCESS_DETACH 之前)')
}

# ---- R5 驱动那一问 ----
Want 'R5 driver seeds once' (C 'cgen.setProjectTitle(' $DRV) 1 '工程标题只许递进去一处'
# 账 #305 第三百五十九刀改的是**口径**不是调用数: 无 Title= 时先退 exe 名 (ExeName32 的 stem),
# 再退工程 Name。产物侧为什么现在是对的 —— 参考二进制的持久化键实测就是 exe 名:
#   .vbp: ExeName32="GUISample.exe", Name="Project1", 无 Title=;
#   参考运行写下的 HKCU\Software\VB and VBA Program Settings\GUISample\Settings 还在
#   (MainLeft/MainTop/MainWidth/MainHeight 四项), 键名 = GUISample = exe 名, 不是 Project1。
# 账 #605 (第330刀) 又把这句算法收进 appTitleSeed() 一处 (Exe 与 DLL 递同一枚种子)。
# 所以本哨兵钉「算法体内的三档次序」(空白/换行不敏感), 而不是钉调用点的字面量:
# 谁把 exe 名那一档撤掉/挪到 Name 之后, 这里就红; 旧两档/就地抄形回潮各单独钉 0。
$r5ThreeTier = '(?s)std::string appTitleSeed\(\) const \{\s*return !verTitle_\.empty\(\)\s*?\?\s*verTitle_\s*:\s*\(' +
                '\s*!projectBaseName_\.empty\(\)\s*?\?\s*projectBaseName_\s*:\s*vbpProjectName_\s*\)\s*;\s*\}'
Want 'R5 Title then exe name then Name' ([regex]::Matches($HPP, $r5ThreeTier).Count) 1 '口径 = .vbp 的 Title, 空则 exe 名 (ExeName32 stem, 参考持久化键), 最后才是工程 Name'
Want 'R5 station exe' (C 'cgen.setProjectTitle(appTitleSeed());' $DRV) 1 'Exe 那一路逐模块 cgen 递种子'
Want 'R5 station dll' (C 'dllCgen.setProjectTitle(appTitleSeed());' $DLLT) 1 'DLL 那一路也递同一枚种子 —— 账 #605 的本体'
$r5OldTwoTier = '(?s)setProjectTitle\(!verTitle_\.empty\(\)'
Want 'R5 old inline fold gone exe' ([regex]::Matches($DRV, $r5OldTwoTier).Count) 0 '就地抄那句口径不许回潮 (Exe 路)'
Want 'R5 old inline fold gone dll' ([regex]::Matches($DLLT, $r5OldTwoTier).Count) 0 '就地抄那句口径不许回潮 (DLL 路)'
Want 'R5 title field once' (C 'std::string projectTitle_;' $STATE) 1 '携带位只许有一枚'
Want 'R5 setter api once' (C 'void setProjectTitle(std::string t) { projectTitle_ = std::move(t); }' $API) 1 'API 只许有一处, 不许各自写字段'

# ---- R6 读侧折叠 ----
Want 'R6 read folds to the cell' (C 'if (memLower == "title") { lastExpr_ = "vb6_App_Title()"; return; }' $MEM) 1 'App.Title 读取只许折到那一格'
Want 'R6 old EXE-name fold gone' (C 'if (memLower == "title") { lastExpr_ = "vb6_App_EXEName()"; return; }' $MEM) 0 '旧的「App.Title = EXE 名」折算不许回潮'
Want 'R6 EXEName arm kept' (C 'memLower == "exename"' $MEM) 1 'App.EXEName 自己那一臂不许被顺手删掉'

# ---- R7 写侧拦截 ----
Want 'R7 write intercept once' (C 'vb6_SetAppTitle(" + rhsAt + ");' $ASG) 1 'App.Title 的写只许有一处出口'
Want 'R7 intercept guard' (C 'objAt == "app" && memAt == "title" && !knownLocalVars_.count(objAt)' $ASG) 1 '局部同名变量优先 (与读取侧同一判据)'
Want 'R7 rhs wraps to BSTR' (C 'wrapToBSTR(std::move(lastExpr_), *node.value)' $ASG) 1 '右值按 BSTR 交给 setter (数值/串各走既有折算)'

# ---- R8 BSTR 返回名单 (第二步后 6 -> 4) ----
# Print#/Write# 的两份名字表已删，分流改问 inferExprType；那两处不再需要认 App.Title。
# 剩下四处（拼接/调用/builtin 前置）仍按名字认，继续钉。
$lists = [ordered]@{
    'src/backend/detail/expr/cgen_expr_call_builtin_pre.inc' = 1
    'src/backend/expr/cgen_expr_binary.cpp'                  = 1
    'src/backend/expr/cgen_expr_binary_util.cpp'             = 1
    'src/backend/stmt/cgen_call.cpp'                         = 1
    # 第二步已删：cgen_file_io.cpp 的两份 bstrFuncs 名单（此处登记 0，回潮即红）
    'src/backend/stmt/cgen_file_io.cpp'                      = 0
}
$listSum = 0
foreach ($k in $lists.Keys) {
    $t = Load $k
    $got = C '"vb6_App_Title' $t
    Want ('R8 list ' + [IO.Path]::GetFileName($k)) $got $lists[$k] ('这份 BSTR 返回名单漏了 App.Title (没登记的出口 = 每调用漏一枚 BSTR / 拼接形状错)')
    $listSum += $got
}
Want 'R8 list total' $listSum 4 '四处仍在的名单站都要认这枚名字 (file_io 那两处已删)'
# R8b: 删表之后 App.Title 走字符串路径的**运行期**证据 (类型由 inferExprType 答 String)
# 第三刀 (账 #608/第331刀) 把「这枚实参是不是 BSTR」收成一位主人 fileOutArgIsBstr:
# 类型那一问只许在主人里出现一次; 两条分流 (Print#/Write#) 与窗体面 Print 必须都转调它 ——
# 它们各自留一份拷贝就是第二个判据 (名字/形状会重新变成主判据)。
$FIO = Load 'src/backend/stmt/cgen_file_io.cpp'
Want 'R8b file_io owner defined' (C 'bool CCodeGen::fileOutArgIsBstr(const std::string& expr, Expr& argNode) const {' $FIO) 1 '「这枚实参是不是 BSTR」只许一位主人'
Want 'R8b file_io asks type once' (C 'inferExprType(argNode) == Vb6Type::String' $FIO) 1 '类型那一问只住主人里 (两处各抄一份 = 名字判据回潮)'
Want 'R8b file_io both sites call owner' (C 'if (fileOutArgIsBstr(val, *expr)) {' $FIO) 2 'Print# 与 Write# 都必须转调主人'
Want 'R8b file_io form print owner' (C 'fileOutArgBstrExpr(val, *expr)' $FIO) 1 '窗体面 Print 也必须转调同一主人'
Want 'R8b b600 prints App.Title' (C 'Print #f, App.Title' (Load 'tests/b600_print_string_dispatch/B600Main.bas')) 1 'B600 必须有一条 App.Title 的 Print #'
Want 'R8b b600 title seeded' (C 'Title="B600 Title"' (Load 'tests/b600_print_string_dispatch/B600.vbp')) 1 'B600 的 .vbp 要有 Title (否则读回 EXE 名, 与类型路径无关)'
Want 'R8b b600 needle' (C '"[B600 Title]"' $RUN) 1 '期望表里要钉 App.Title 写盘行 (数值路径会写成指针整数)'

# ---- R10 DLL 入口那一站 (账 #605 的脸①) ----  (腿在 R4 里, 这里只钉「两站同形」)
Want 'R10 two seed stations' ((C 'vb6_SetAppTitle(L\"' $ENT) + (C 'vb6_SetAppTitle(L\"' $DLLE)) 2 'Exe 入口模板与 DLL 入口各一枚, 多一枚就是另立形状'

# ---- R11 App.Path / App.EXEName 答这枚产物自己 (账 #605 的脸②) ----
Want 'R11 owner defined once' (C 'static const wchar_t* vb6_SelfModulePath(void) {' $SYS) 1 '「这枚产物住在哪个模块」只许一处实现'
Want 'R11 owner asks its own module' (C 'GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,' $SYS) 1 '取法 = 从自己的地址问模块 (静态链进谁就答谁)'
Want 'R11 stations' (C 'vb6_SelfModulePath()' $SYS) 2 'App.Path 与 App.EXEName 各问一次 (第三处 = 有人另立主人)'
Want 'R11 host fold gone' (C 'GetModuleFileNameW(NULL, buf, 1024)' $SYS) 1 '只剩 App.PrevInstance 还问进程 (无 VB6 读数, 刻意在册)'
Want 'R11 hInstance kept' (C 'return (intptr_t)(uintptr_t)GetModuleHandleW(NULL);' $SYS) 1 'App.hInstance 本刀刻意不动 (同上没有读数)'
Want 'R11 never frees self' (C 'FreeLibrary(' $SYS) 0 '给自己那枚模块调 FreeLibrary = 正在执行中的模块被卸 (崩)'

# ---- R9 登记面 ----
Want 'R9 sentinel registered' (C 'check_app_title_authority.ps1' $RUN) 1 '本哨兵在 run_tests.ps1 里被引用一次'
Want 'R9 sentinel function' (C 'function Test-AppTitleAuthority' $RUN) 1 '那道静态检查自己有一处定义'
Want 'R9 sentinel called' ([regex]::Matches($RUN, '(?m)^\s+(?:if \(Enter-StaticShard\) \{ )?Test-AppTitleAuthority\s*(?: \})?\s*$').Count) 1 '裸调用自己有一行 (只有定义没调用 = 门永远不跑)'
Want 'R9 vbp cells' (C 'Test-Vbp "b592_app_title' $RUN) 2 '两台各一片 (x64 + x86)'
Want 'R9 bas cells' (C 'Add-BasTest "b592_app_title_fallback' $RUN) 2 '两个架构各一片'
Want 'R9 emit shape cell' (C 'Test-EmitcShape "b592_app_title_shape"' $RUN) 1 '正针: 发码里要有播种 + 读写两形'
Want 'R9 emit absent cell' (C 'Test-EmitcAbsent "b592_app_title_lesion"' $RUN) 1 '反针: 旧左值形本体不许在场'
Want 'R9 stray escaped arch' (C '-Arch "x86\' $RUN) 0 '行尾多一个反斜杠会把 C3 的参数打成 x86\ (第九十八刀真栽过)'
Want 'R9 vbp needle seed' (C '"APT-seed=B592 Title"' $RUN) 1 '期望表里要钉播种值 (读的确实是 .vbp 的 Title, 不是 Name)'
Want 'R9 vbp needle set' (C '"APT-set=renamed"' $RUN) 1 '期望表里要钉运行期改写'
Want 'R9 vbp needle empty' (C '"APT-empty=[]"' $RUN) 1 '期望表里要钉「显式空串答空」'
Want 'R9 fallback needle' (C '"APT-fallback=b592_apptitle_fallback"' $RUN) 1 '单文件编译的旧口径 (退 EXEName) 要在期望表里'
Want 'R9 seed needle shape' (C 'vb6_SetAppTitle(L"B592 Title");' $RUN) 1 '发码正针要钉播种那一行'
Want 'R9 lesion needle shape' ([regex]::Matches($RUN, '(?m)^\s+' + [regex]::Escape($SQ + 'vb6_App_EXEName() = ' + $SQ) + '\s*$').Count) 1 '发码反针要钉旧左值形'
# 账 #605 那一格: 形状 (dll_entry.c) 与行为 (进程内一点) 同趟编译各读一次, 尺住在 disp_invoke.ps1。
Want 'R9 dll harness defined' (C 'function Test-DllAppSelf' $DI) 1 '那把进程内的尺有一处定义'
Want 'R9 dll case registered' (C 'Test-DllAppSelf "b605_dll_app_self"' $RUN) 1 'b605 那一格在 runner 里登记'
Want 'R9 dll seed needle' (C 'vb6_SetAppTitle(L"B605 Dll Title");' $RUN) 1 '那一格要钉 dll_entry.c 的播种行'
Want 'R9 dll title needle' (C '"READ_ANS=605-title=[B605 Dll Title]",' $RUN) 1 '期望里要钉运行期读到的 Title (注释 + 针)'
Want 'R9 dll exe needle' (C '"exe=[B605k]",' $RUN) 1 '期望里要钉 DLL 自己的模块名'
Want 'R9 dll path needle' (C '"tailOk=True") @(' $RUN) 1 '期望里要钉 App.Path 落在 DLL 自己那个目录'
Want 'R9 dll absent host' (C '"exe=[disp_probe",' $RUN) 1 '反面证人: 宿主名字顶回来 = 又退回 GetModuleFileNameW(NULL) (注释 + 针)'
Want 'R9 dll absent half' (C '"title=[B605k]")' $RUN) 1 '反面证人: 只修读取那一脸、播种没落也过不了'

# ---- 夹具本体 ----
Want 'FX vbp Title' (C 'Title="B592 Title"' $FX3) 1 '夹具的 .vbp 要有 Title (否则播种那一行根本不发)'
Want 'FX vbp ExeName32' (C 'ExeName32="B592AppTitle.exe"' $FX3) 1 'Test-Vbp 按 ExeName32 找产物 (账 #166)'
Want 'FX read/write' (C 'App.Title' $FX1) 5 '夹具要覆盖 读(3 处) + 写(2 处) 五种形'
Want 'FX write set' (C 'App.Title = "renamed"' $FX1) 1 '夹具要覆盖运行期改写'
Want 'FX write empty' (C 'App.Title = ""' $FX1) 1 '夹具要覆盖显式空串'
Want 'FX single-file read' (C 'App.Title' $FX2) 1 '单文件夹具要读 (证未播种时退 EXEName)'
Want 'FX605 vbp Type' (C 'Type=OleDll' $FX4) 1 '夹具必须是 ActiveX DLL —— 本刀那一路的入口'
Want 'FX605 vbp Title' (C 'Title="B605 Dll Title"' $FX4) 1 '播种值要有来处 (与运行期读数同源)'
Want 'FX605 vbp ExeName32' (C 'ExeName32="B605k.dll"' $FX4) 1 '产物名按 ExeName32 找 (账 #166 那条口径)'
Want 'FX605 read title' (C '& App.Title &' $FX5) 1 '夹具要在 DLL 里读 App.Title'
Want 'FX605 read exename' (C '& App.EXEName &' $FX5) 1 '夹具要读 App.EXEName (答 DLL 还是宿主的那一问)'
Want 'FX605 read path' (C 'App.Path, "b605_self"' $FX5) 1 '夹具要拿 App.Path 与 DLL 自己的目录标记对表'

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Output ('FAIL ' + $m) }
    Write-Output ('FAIL check_app_title_authority: ' + $fail.Count + ' rule(s) broken')
    exit 1
}
Write-Output ('PASS check_app_title_authority: entry seed ' + $nSeed + '/' + $nInit +
    ' lists=' + $listSum + ' cells=' + (C 'Test-Vbp "b592_app_title' $RUN) +
    ' read=' + (C 'if (memLower == "title") { lastExpr_ = "vb6_App_Title()"; return; }' $MEM) +
    ' write=' + (C 'vb6_SetAppTitle(" + rhsAt + ");' $ASG) +
    ' msgbox=' + (C 'vb6_AppTitleBorrow()' $CONV) +
    ' dllseed=' + (C 'vb6_SetAppTitle(L\"' $DLLE) +
    ' selfmod=' + (C 'vb6_SelfModulePath()' $SYS))
exit 0
