# check_test_helper_integrity.ps1 - 账 #78 的结构哨兵 (第 40 道；只扫 tests/run_tests.ps1, 不起 cl)
#
# 为什么要这一道: #78 那一格的形状是「两个 helper 同名, 后定义的把先定义的遮掉了」。
# PowerShell 里后定义者赢, 于是 tests/run_tests.ps1 里那个**真构建**的 Test-Compile (旧 303 行)
# 从来没跑过 —— 打印的还是 "[COMPILE] ..."，注释写的是「编译+运行」，门里 10 份夹具
# (test_comprehensive[2].bas + 8 枚 .frm) 于是**一条 cl 也没编过**，而且没有任何地方报红。
# 这一族与 #166 (exe 名两个权威)、#231 (属性名单三份)、#245 (类型来源) 同一个形状，
# 而它比它们更阴：被遮掉的那半根本不产生噪声，它只是安静地不跑。
#
# 规则 (改坏了会红):
#   H1  tests/run_tests.ps1 里 `function` 名字**不许重复** (按出现次数数, 一个名字出现两次 = 遮蔽)。
#   H2  每个 `function Test-*` 助手必须在自己的定义之外还被**提到**至少一次
#       (写在 `Where-Object { Test-X ... }` / `if (Test-X ...)` 里也算 —— 覆盖面是"调不调得到"，
#        与第 37 道哨兵的 R1b 同一口径: 「提没提」不算，「跑不跑」才算)。
#   H3  两个助手的身份必须钉住: 唯一那枚 `Test-Compile` 的函数体里必须有 `--output-dir` (真构建)
#       且**不许**出现 `Invoke-CodegenProj`；codegen 那一枚必须叫 `Test-CodegenOk`，体里必须有
#       `Invoke-CodegenProj`。⇒ 谁再把只发码的那枚改回同名，这一格当场红。
#   H4  --emit-c 那一族的判据不许冒充构建: `Test-CodegenOk` / `Test-EmitcShape` / `Test-EmitcAbsent`
#       三者的标签必须互不相同且都不是 `[COMPILE]` (日志上读得出"这格只发了码")。
#   H5  `Test-DispatchInvoke` 的**调用面**要用 AST 钉住绑定: 每个调用点的「位置实参个数 / 数组实参
#       个数 / Needles 与 Absent 各含几条针」都必须等于名单里的数。为什么要 AST —— 往多实参助手的
#       数组里插针时，插错的那一笔**语法仍然对**，变的只是绑定: 门 #532 上第三十八刀插进 Needles 的
#       三条针吃掉了收尾的 `") @(`, 于是 `"CALL=ADD"` 从 Absent 变成 Needles、ExtraIid 往后挪了一格,
#       探针少了第三个参数 ⇒ 整段薄指针/契约槽读数(EXTRA_IID/QI_EXTRA/CREATE_IFACE/VTBL_*)安静不出现,
#       而 `[Parser]::ParseFile` 报 0 错。这一格盯的就是那一种红。
#
#   H6  注册语句不许被**行首注释**吃掉 (账 #305 §B163, 第五十一刀的实测坑): 把新写的注释拼到最后一行时如果把语句一起拼在 `#` 那一行的行尾, 它整条变成注释 —— 而 `[Parser]::ParseFile` 仍报 0 错 (与门 #532 那次同一族: 红不在语法, 在绑定)。
#       判据 = AST 给的每个注册调用点必须是它那行第一个非空 token（行前只许空白）——不是数出现次数，因为注释里合法可以提这个名字。
# 用法: powershell -File scripts/check_test_helper_integrity.ps1   (PASS => exit 0)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$rt = Join-Path $root 'tests\run_tests.ps1'
if (-not (Test-Path $rt)) { Write-Host "FAIL missing $rt"; exit 1 }
$lines = Get-Content -LiteralPath $rt -Encoding UTF8

$fails = @()

# ---- 定义表：名字 -> 行号列表 (只认顶格的 function) ----
$defs = @{}
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^function\s+([A-Za-z][A-Za-z0-9_-]*)') {
        $n = $Matches[1]
        if (-not $defs.ContainsKey($n)) { $defs[$n] = @() }
        $defs[$n] += ($i + 1)
    }
}

# ---- H1: 重名遮蔽 ----
foreach ($n in ($defs.Keys | Sort-Object)) {
    if ($defs[$n].Count -gt 1) {
        $fails += ('H1 function name ' + $n + ' is defined ' + $defs[$n].Count + ' times at lines ' +
                   ($defs[$n] -join ',') + ' - the later one silently shadows the earlier one' +
                   ' (that is exactly how the [COMPILE] group stopped building anything)')
    }
}

# ---- H2: 每个 Test-* 助手都要被引用到 (定义之外) ----
# 只数**代码行**: 整行以 # 开头的是注释, 注释里提一句不算引用。这一条不能松 ——
# 把调用点注释掉正是"覆盖面安静消失"的第二种写法 (与 #78 的重名遮蔽同一个后果),
# 按全文数的话它永远绿。注: 代码行**尾部**的行内注释仍会被数到 (偏保守, 不会假红)。
$codeLines = @()
foreach ($l in $lines) {
    $t = $l.TrimStart()
    if ($t.StartsWith('#')) { continue }
    $codeLines += $l
}
$raw = $codeLines -join "`n"
foreach ($n in ($defs.Keys | Where-Object { $_ -like 'Test-*' } | Sort-Object)) {
    $all = [regex]::Matches($raw, '(?<![\w-])' + [regex]::Escape($n) + '(?![\w-])').Count
    $own = $defs[$n].Count                      # 每处定义本身占一次
    if ($all -le $own) {
        $fails += ('H2 helper ' + $n + ' is never referenced outside its definition (dead cell: ' +
                   'a judgement that cannot run is not a judgement)')
    }
}

# ---- H3: 两枚助手的身份 ----
function Body-of([string]$name) {
    if (-not $defs.ContainsKey($name)) { return $null }
    $start = $defs[$name][0] - 1
    $depth = 0; $seen = $false
    for ($j = $start; $j -lt $lines.Count; $j++) {
        foreach ($ch in $lines[$j].ToCharArray()) {
            if ($ch -eq '{') { $depth++; $seen = $true }
            elseif ($ch -eq '}') { $depth-- }
        }
        if ($seen -and $depth -le 0) { return ($lines[$start..$j] -join "`n") }
    }
    return $null
}
$bCompile = Body-of 'Test-Compile'
if (-not $bCompile) {
    $fails += 'H3 Test-Compile is missing'
} else {
    if ($bCompile -notmatch '--output-dir') {
        $fails += 'H3 Test-Compile no longer invokes a real build (no --output-dir) - the [COMPILE] label would lie'
    }
    if ($bCompile -match 'Invoke-CodegenProj') {
        $fails += 'H3 Test-Compile went back to codegen-only (Invoke-CodegenProj) - that is the #78 shadowing bug reborn'
    }
}
$bCodegen = Body-of 'Test-CodegenOk'
if (-not $bCodegen) {
    $fails += 'H3 Test-CodegenOk is missing (the codegen-only helper must carry a name that says so)'
} elseif ($bCodegen -notmatch 'Invoke-CodegenProj') {
    $fails += 'H3 Test-CodegenOk no longer runs the codegen path it is named for'
}

# ---- H4: 标签不许串味 ----
$labels = @{}
foreach ($pair in @(@('Test-Compile', '[COMPILE]'), @('Test-CodegenOk', '[CODEGEN-OK]'),
                   @('Test-EmitcShape', '[EMITC-SHAPE]'), @('Test-EmitcAbsent', '[EMITC-ABSENT]'))) {
    $body = Body-of $pair[0]
    if ($null -eq $body) { continue }          # H3 已经报过
    if ($body -notmatch [regex]::Escape($pair[1])) {
        $fails += ('H4 ' + $pair[0] + ' must print the label ' + $pair[1] + ' (the log has to say which kind of check ran)')
    }
    if ($pair[0] -ne 'Test-Compile' -and $body -match '\[COMPILE\]') {
        $fails += ('H4 ' + $pair[0] + ' prints [COMPILE] although it does not build anything')
    }
}

# ---- H5: 多数组实参助手的绑定形状 (账 #305 SB144, 门 #532 的实证 ----
$tok5 = $null; $err5 = $null
$ast5 = [System.Management.Automation.Language.Parser]::ParseFile($rt, [ref]$tok5, [ref]$err5)
if ($err5.Count -gt 0) {
    $fails += ('H5 parse errors in run_tests.ps1: ' + $err5.Count)
    $ast5 = $null
}
if ($ast5) {
    $wantD = @{
    # 名字 = @(element数-助手名, Needles 针数, Absent 针数, -Tail 针数) —— 加针/加调用点要同时改这张名单
    # 账 #305 §B159/§B162 (第五十刀): 两台各加一条 CALL=ECHOARR (真 SAFEARRAY 进服务端、再交回客户端)。
    'ax_dll_dispatch_invoke'         = @(6, 6, 1, 0)
    'cc_dll_dispatch_iface_only'     = @(7, 11, 1, 0)
    'cc_dll_dispatch_iface_only_x86' = @(8, 11, 1, 0)
    # 账 #305 §B185 (第七十七刀): IDispatch 服务端进门那枚错误边界, 两台各一条真跑用例。
    # 针面各写一份(见 run_tests.ps1 那两条的理由); x86 那格多出的两枚元素是 `-Arch "x86"`
    # 这一对**命名**实参 —— AST 里它们各算一格 CommandElement, 所以是 8 而不是 7。
    'b185_server_error_frame'        = @(6, 12, 1, 0)
    'b185_server_error_frame_x86'    = @(8, 12, 1, 0)
    # 账 #305 §B186 (第七十八刀): 类工厂那辙门的错误边界, 两台各一条真跑用例。
    # 第四格 = -Tail 那枚数组里的针数 (没有这一档就写 0): 这一族用例给助手新增了一个
    # 命名数组实参, 「数组实参恰好两枚」那句话从此只管**位置**那两档。
    'b186_factory_error_frame'         = @(8, 5, 2, 1)
    'b186_factory_error_frame_x86'    = @(10, 5, 2, 1)
    # 账 #305 §B188 (第八十刀判据): 出厂那一趟的格子, 两枚一类各一台; 位置面 6 (+2 = -Arch),
    # Needles 6 / Absent 1 (「Unhandled error」当反面证人) / -Tail 1 枚 CREATE= 臂
    'b188_terminate_error_frame'          = @(8, 6, 1, 1)
    'b188_terminate_error_frame_x86'     = @(10, 6, 1, 1)
    # 账 #305 §B187 第二腿 + 第四腿 + 第五腿 (八十三刀装箱 / 八十四刀实参那一站 / 第八十五刀 ByRef 那座桥 / 八十六刀 TypeName 认包装器): 两架构真跑格。
    # 位置面 8 (+2 = -Arch), Needles 10 / Absent 6 (装箱腿 / 实参站 / ByRef 桥 / 类名那一问 / Nothing 那一问 / 声明档折叠那一问 六处各留一枚"没落地就回来"的数)
    # / -Tail 1 枚 CREATE= 臂 (7 个 CLSID 写在同一条里)
    'b187_boxed_object_slot'                    = @(8, 10, 6, 1)
    'b187_boxed_object_slot_x86'               = @(10, 10, 6, 1)
}
$ds = @($ast5.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.CommandAst] -and
    $n.GetCommandName() -eq 'Test-DispatchInvoke' }, $true))
if ($ds.Count -ne $wantD.Count) {
    $fails += ('H5 Test-DispatchInvoke call sites ' + $ds.Count + ' != ' + $wantD.Count +
               ' (名单要跟着调用点一起改)')
}
foreach ($c in $ds) {
    $els = @($c.CommandElements)
    $nm = ($els[1].Extent.Text -replace '"', '')
    # 账 #305 §B186 (第七十八刀): 数组实参分两档。紧跟在**命名参数**后面的那档不是针面,
    # 是给助手的参数据的 (今天只认 -Tail 这一枚); 位置那两档才是 Needles 与 Absent。
    # 之前那句「数组实参恰好 2 枚」把两档混在一起数, 加了 -Tail 就会永远红 ——
    # 要钉的其实是「位置两档各自的针数 + 命名那档叫什么、里面几条」。
    $pos = @(); $named = @()
    for ($e = 0; $e -lt $els.Count; $e++) {
        if ($els[$e].GetType().Name -ne 'ArrayExpressionAst') { continue }
        if ($e -gt 0 -and $els[$e - 1].GetType().Name -eq 'CommandParameterAst') {
            $named += , @($els[$e - 1].Extent.Text, $els[$e])
        } else {
            $pos += $els[$e]
        }
    }
    # 注释行里也写引号(那几行是给人读的)，先丢掉再数针的条数
    $counts = @($pos | ForEach-Object {
        $clean = ($_.Extent.Text -split '\r?\n' | Where-Object { $_.TrimStart() -notmatch '^#' }) -join "`n"
        ([regex]::Matches($clean, '"[^"\r\n]*"')).Count
    })
    $tailNames = @($named | ForEach-Object { $_[0] })
    $tailCounts = @($named | ForEach-Object {
        $clean = ($_[1].Extent.Text -split '\r?\n' | Where-Object { $_.TrimStart() -notmatch '^#' }) -join "`n"
        ([regex]::Matches($clean, '"[^"\r\n]*"')).Count
    })
    if (-not $wantD.ContainsKey($nm)) { $fails += ('H5 没登记过的调用点 ' + $nm); continue }
    $w = $wantD[$nm]
    if (($els.Count - 1) -ne $w[0]) {
        $fails += ('H5 ' + $nm + ' 位置实参 ' + ($els.Count - 1) + ' != ' + $w[0] +
                   ' —— 数组收尾被吃掉时后面的实参会整体往前挪')
    }
    if ($pos.Count -ne 2) {
        $fails += ('H5 ' + $nm + ' 位置数组实参 ' + $pos.Count + ' != 2 (Needles + Absent)')
    }
    if ($tailNames.Count -gt 1 -or ($tailNames.Count -eq 1 -and $tailNames[0] -ne '-Tail')) {
        $fails += ('H5 ' + $nm + ' 不认得的命名数组实参: ' + ($tailNames -join ',') +
                   ' —— 名单只认 -Tail 这一档')
    }
    $tailN = 0
    if ($tailCounts.Count -eq 1) { $tailN = $tailCounts[0] }
    if ($w.Count -lt 4) {
        $fails += ('H5 名单 ' + $nm + ' 只有 ' + $w.Count + ' 个数 —— 第四格是 -Tail 的针数 (没有写 0)')
    } elseif ($tailN -ne $w[3]) {
        $fails += ('H5 ' + $nm + ' -Tail 针数 ' + $tailN + ' != ' + $w[3])
    }
    if ($counts.Count -eq 2 -and ($counts[0] -ne $w[1] -or $counts[1] -ne $w[2])) {
        $fails += ('H5 ' + $nm + ' 针数 ' + ($counts -join '/') + ' != ' + $w[1] + '/' + $w[2])
    }
}

}
# ---- H6: 注册语句必须自成一行，也不许整条变成注释 (账 #305 §B163, 第五十一刀的实测坑) ----
# 形状: 把新写的注释拼到最后一行时，若连语句一起拼在 `#` 那行的行尾，它整条变成注释;
#       [Parser]::ParseFile 仍报 0 错，那枚夹具安静不跑（与门 #532 同一族: 红不在语法，在绑定）。
# 两条判据: ① 剥掉注释后的 token 文本里「名字 + 空白 + 引号或括号」的调用形出现几次，AST 就必须找到几个调用点
#           (少了 = 有一条被注释吃掉了，那种情形 AST 里根本不存在，只靠②盯不住);
#           ② 每个调用点必须是它那一行第一个非空 token（行前只许空白）。
if ($ast5) {
    $txt6 = [System.IO.File]::ReadAllText($rt)
    $code6 = (($tok5 | Where-Object { $_.GetType().Name -ne 'CommentToken' } | ForEach-Object { $_.Text }) -join ' ')
    foreach ($fn6 in @('Add-BasTest', 'Test-DispatchInvoke')) {
        $txt6n = @([regex]::Matches($code6, ($fn6 + '[ \t]+["\(]'))).Count
        $cs6 = @($ast5.FindAll({ param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq $fn6 }, $true))
        if ($txt6n -ne $cs6.Count) {
            $fails += ('H6 ' + $fn6 + ': 剥注释后的调用形 ' + $txt6n + ' 处, AST 调用点 ' + $cs6.Count +
                       ' 处 —— 差的那几条整行变成了注释（语法不响，夹具安静不跑）')
        }
        foreach ($c6 in $cs6) {
            $st6 = $c6.Extent.StartOffset
            $ls6 = $txt6.LastIndexOf("`n", $st6 - 1)
            if ($ls6 -lt 0) { $ls6 = -1 }
            if ($txt6.Substring($ls6 + 1, $st6 - $ls6 - 1) -notmatch '^[ \t]*$') {
                $ln6 = [regex]::Matches($txt6.Substring(0, $st6), "`n").Count + 1
                $fails += ('H6 ' + $fn6 + ' 第 ' + $ln6 + ' 行的注册不是那行的开头'+ ' —— 排在注释行行尾的语句只是注释，夹具安静不跑')
            }
        }
    }
}

# ---- H7: 这套 harness 自己（与全部哨兵）必须带 UTF-8 BOM ----
#   为什么这道门非有不可: pwsh 7 读无 BOM 的 UTF-8 也认中文 ⇒ **本机永远绿**，只有 CI 的
#   Windows PowerShell 5.1 会按 ANSI(cp936) 读 ⇒ 中文注释被解成乱码，乱码里带出的引号/反引号
#   把语法**炸掉**。账 #438 实测: `tests/run_tests.ps1` 的 BOM 被 `[IO.File]::WriteAllText` 顺手
#   抹掉（那一句是为了插一个调用点），本机 `pwsh` 一路绿，CI 上
#   `[STATIC] ps51_joinpath` + `[STATIC] test_helper_integrity` 双双红成
#   「Unexpected token '\Include'」/「parse errors: 72」—— 而报错指的地方跟动手的地方
#   **隔着一千多行**，因为真正炸的是**注释**。
#   作用域只取「CI 会去 parse 的那批」= harness + scripts/check_*.ps1：仓库里另有 10 份
#   无 BOM 的 .ps1（含 6 份带中文的构建/开发脚本），它们不由这套 harness parse，
#   拿「全仓库都要 BOM」当判据会平白判红 6 条。
if (-not $fails.Contains('H7x')) {
    $needBom = @($rt) + @(Get-ChildItem (Join-Path $PSScriptRoot 'check_*.ps1') | ForEach-Object { $_.FullName })
    foreach ($p7 in $needBom) {
        $b7 = [System.IO.File]::ReadAllBytes($p7)
        if ($b7.Length -lt 3 -or $b7[0] -ne 0xEF -or $b7[1] -ne 0xBB -or $b7[2] -ne 0xBF) {
            $fails += ('H7 ' + (Split-Path -Leaf $p7) + ' 没有 UTF-8 BOM —— pwsh 7 照样绿，只有 CI 的 5.1 会按 ' +
                       'ANSI 读，把中文注释解成乱码再炸语法；要改这个文件请用会保留 BOM 的写法')
        }
    }
}

if ($fails.Count -gt 0) {
    $fails | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host ('OK check_test_helper_integrity: H1..H7 (functions=' + $defs.Count +
             ' Test-* helpers=' + @($defs.Keys | Where-Object { $_ -like 'Test-*' }).Count +
             ' no shadowing, all reachable, harness+sentinels all BOM-ed)')
exit 0
