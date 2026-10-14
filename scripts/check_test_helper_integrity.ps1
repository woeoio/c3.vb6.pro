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
    # 名字 = @(位置实参个数, Needles 针数, Absent 针数) —— 加针/加调用点要同时改这张名单
    # 账 #305 §B159/§B162 (第五十刀): 两台各加一条 CALL=ECHOARR (真 SAFEARRAY 进服务端、再交回客户端)。
    'ax_dll_dispatch_invoke'         = @(6, 6, 1)
    'cc_dll_dispatch_iface_only'     = @(7, 11, 1)
    'cc_dll_dispatch_iface_only_x86' = @(8, 11, 1)
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
    $arrays = @($els | Where-Object { $_.GetType().Name -eq 'ArrayExpressionAst' })
    # 注释行里也写引号(那几行是给人读的)，先丢掉再数针的条数
    $counts = @($arrays | ForEach-Object {
        $clean = ($_.Extent.Text -split '\r?\n' | Where-Object { $_.TrimStart() -notmatch '^#' }) -join "`n"
        ([regex]::Matches($clean, '"[^"\r\n]*"')).Count
    })
    if (-not $wantD.ContainsKey($nm)) { $fails += ('H5 没登记过的调用点 ' + $nm); continue }
    $w = $wantD[$nm]
    if (($els.Count - 1) -ne $w[0]) {
        $fails += ('H5 ' + $nm + ' 位置实参 ' + ($els.Count - 1) + ' != ' + $w[0] +
                   ' —— 数组收尾被吃掉时后面的实参会整体往前挪')
    }
    if ($arrays.Count -ne 2) {
        $fails += ('H5 ' + $nm + ' 数组实参 ' + $arrays.Count + ' != 2 (Needles + Absent)')
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

if ($fails.Count -gt 0) {
    $fails | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host ('OK check_test_helper_integrity: H1..H6 (functions=' + $defs.Count +
             ' Test-* helpers=' + @($defs.Keys | Where-Object { $_ -like 'Test-*' }).Count +
             ' no shadowing, all reachable)')
exit 0
