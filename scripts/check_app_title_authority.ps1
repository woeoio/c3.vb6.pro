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
#   R5 驱动按 Title 优先、Name 兜底的读法只有一处, 字段与 setter 各一枚;
#   R6 读侧折叠恰一处, 旧的 EXE 名形已撤 (EXEName 自己那一臂仍在);
#   R7 写侧拦截恰一处, 且局部同名优先 (knownLocalVars_ 同一判据);
#   R8 认这枚名字的六份 BSTR 返回名单一份都没漏;
#   R9 夹具 / 期望 / 发码针 / 本哨兵自己都已在 run_tests.ps1 登记。
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
$FX1  = Load 'tests/b592_apptitle/APT.bas'
$FX2  = Load 'tests/b592_apptitle_fallback.bas'
$FX3  = Load 'tests/b592_apptitle/B592.vbp'

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

# ---- R5 驱动那一问 ----
Want 'R5 driver seeds once' (C 'cgen.setProjectTitle(' $DRV) 1 '工程标题只许递进去一处'
Want 'R5 Title first Name fallback' (C 'cgen.setProjectTitle(!verTitle_.empty() ? verTitle_ : vbpProjectName_);' $DRV) 1 '口径 = .vbp 的 Title, 空则 Name (VB6 的默认题头同理)'
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

# ---- R8 六份 BSTR 返回名单 ----
$lists = [ordered]@{
    'src/backend/detail/expr/cgen_expr_call_builtin_pre.inc' = 1
    'src/backend/expr/cgen_expr_binary.cpp'                  = 1
    'src/backend/expr/cgen_expr_binary_util.cpp'             = 1
    'src/backend/stmt/cgen_call.cpp'                         = 1
    'src/backend/stmt/cgen_file_io.cpp'                      = 2
}
$listSum = 0
foreach ($k in $lists.Keys) {
    $t = Load $k
    $got = C '"vb6_App_Title' $t
    Want ('R8 list ' + [IO.Path]::GetFileName($k)) $got $lists[$k] ('这份 BSTR 返回名单漏了 App.Title (没登记的出口 = 每调用漏一枚 BSTR / 拼接形状错)')
    $listSum += $got
}
Want 'R8 list total' $listSum 6 '六处名单站都要认这枚名字'

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

# ---- 夹具本体 ----
Want 'FX vbp Title' (C 'Title="B592 Title"' $FX3) 1 '夹具的 .vbp 要有 Title (否则播种那一行根本不发)'
Want 'FX vbp ExeName32' (C 'ExeName32="B592AppTitle.exe"' $FX3) 1 'Test-Vbp 按 ExeName32 找产物 (账 #166)'
Want 'FX read/write' (C 'App.Title' $FX1) 5 '夹具要覆盖 读(3 处) + 写(2 处) 五种形'
Want 'FX write set' (C 'App.Title = "renamed"' $FX1) 1 '夹具要覆盖运行期改写'
Want 'FX write empty' (C 'App.Title = ""' $FX1) 1 '夹具要覆盖显式空串'
Want 'FX single-file read' (C 'App.Title' $FX2) 1 '单文件夹具要读 (证未播种时退 EXEName)'

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Output ('FAIL ' + $m) }
    Write-Output ('FAIL check_app_title_authority: ' + $fail.Count + ' rule(s) broken')
    exit 1
}
Write-Output ('PASS check_app_title_authority: entry seed ' + $nSeed + '/' + $nInit +
    ' lists=' + $listSum + ' cells=' + (C 'Test-Vbp "b592_app_title' $RUN) +
    ' read=' + (C 'if (memLower == "title") { lastExpr_ = "vb6_App_Title()"; return; }' $MEM) +
    ' write=' + (C 'vb6_SetAppTitle(" + rhsAt + ");' $ASG) +
    ' msgbox=' + (C 'vb6_AppTitleBorrow()' $CONV))
exit 0
