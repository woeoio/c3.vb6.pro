# check_owned_temp_keeps.ps1 - 账 #454 §B164 形①（第一百三十五刀 = 「谁存住了传进去的带权值」那张 KEEPS 表）
#
# 这一刀**只做表，不动行为**：把「一条语句里现造的带权临时量传进去之后，谁把它存住了」
# 这件事做成一张**有出口权威的表**，并把「出口面」从 grep 名单改成**权威签名表**推出来的。
#
# 为什么必须先有表：上一版（9546dbfe / 6bf5989 / be5d6abf）把 §B164 里那条
# 「KEEPS = 0」的问法当成了「一切实参位上的现造串都没人存住」—— 那一句只覆盖
# 「vb6_VariantToString 第一个实参位上那 3578 处 mint 的消费者」，判的是 RTL 那一侧。
# 命题换错 ⇒ 门把 5 份既有夹具的运行期读数改了（注册表那一族整条断掉等）⇒ 那一版已撤回。
# 教训：**先证明「知道谁不能回收」，再谈「能回收」**。
#
# 三分类（这一版的口径，**只看得指针轴**）：
#   keeps_pointer = 0  确定只在调用内读，指针/槽位/引用**没有**带出调用 ⇒ 允许回收
#   keeps_pointer = 1  确定把指针/槽位/引用**保留到调用外**（存进结构体·堆·全局、
#                      原样返回出去，或经由它写出去）⇒ 禁止回收
#   keeps_pointer = U  证不出（静态分析或权威签名表答不出来）
# 逃逸通道三条（keeps_pointer=1 的「为什么」那一栏，缺一条就漏一族）：
#   store   存进结构体/堆/全局（指针活着越过调用）
#   ret     原样返回（借出去 ⇒ 调用方拿到的是同一枚指针）
#   write   经由它写出去（形参是 BSTR* / VARIANT* 那种可写槽）
#
# **第二条轴：persists_data**（数据持久化）—— 与指针轴**分开**，**不参与回收**。
# 「这条形参带进去的**数据**有没有被落到一个比调用活得久的地方」（注册表 / 文件 / 剪贴板）。
# 注册表那一族在**这条轴上是 1**（值真的落进了注册表），在**指针轴上是 0**（RegSetValueExW
# 与 vb6_RegBuildKey 都复制 ⇒ 指针不留存）—— 两轴各答各的，回收只看指针轴。
#
# 第三列 gate = open / closed：closed 的行**不进提名面**。
# 注册表那一族挂 closed 的理由**不是**它不安全（它指针轴上是 0），而是
# **它那一族的运行期破坏至今未归因**（CI 上 R2-vt=1 而不是 8204，连注册表键都没建起来）
# ⇒ 未归因 ⇒ 不许提名（纪律：未登记差异 ⇒ 停机）。
#
# 出口面（**不许用 grep 名单当权威**）：权威 = RTL 的那 18 个头
# （src/rtl/core/vb6rtl/*.h + src/rtl/core/vb6forms/*.h，也就是生成代码 include 的
# 那两张伞头 vb6rtl.h / vb6forms.h 拉进来的全部子头）。表里每一行都必须能从那 18 个头
# 推出来；推不出来 = unclassified ≠ 0 ⇒ 停机。
#
# **提名面**（下一刀回收时要用的那一块）：keeps_pointer = 0 ∧ gate = open。
# 硬形状（那一刀自己的判据）不在本表里；本表只答「谁安全、谁不安全、谁不许动」。
#
# 分类的**可证性**（这一版最要紧的一条）：只**证得出**的才给 0/1，证不出的给 U。
#   证得出 KEEPS=1：形参被**整枚**赋给一个结构体成员/文件级全局，或被**原样**返回，
#                   或取了它的地址
#   证得出 KEEPS=0：形参从不被整枚赋出去、不被原样返回、不取地址，且它流进去的每一枚
#                   vb6_* 出口自己就是 KEEPS=0（转手 ⇒ 下游说了算；成环 ⇒ U）
#   其余（赋给局部但那个局部的去向答不出、流进答不出的出口、找不到体）⇒ U
# ⚠ 「赋给局部」**不许**直接判 store：`BSTR copy = source;` 是**复制**，不是存指针。
#   那一族（复制）要靠「局部自己有没有再逃出去」来判，判不出就 U。
#
# 用法:  pwsh -File scripts/check_owned_temp_keeps.ps1            # 判
#        pwsh -File scripts/check_owned_temp_keeps.ps1 -Bless    # 重新登记快照
[CmdletBinding()]
param(
    [string]$Root = "",
    [switch]$Bless,
    [string]$TableFile = ""
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }
# ⚠ 参数**不许**叫 $Table：PowerShell 变量名大小写不敏感，而下面有一枚同名的小写 hashtable
# （$rows 那一枚）—— 两者是同一枚变量，[string] 形参会把 hashtable 折成字符串，
# 然后 `$rows[$k]` 报「Unable to index into an object of type System.String」。
if (-not $TableFile) { $TableFile = Join-Path $Root 'scripts\owned_temp_keeps_164.txt' }

$RtlHdrDir = Join-Path $Root 'src\rtl\core\vb6rtl'
$RtlFrmDir = Join-Path $Root 'src\rtl\core\vb6forms'
$bad = @()

# ---------------------------------------------------------------- 权威出口面
$hdrFiles = @()
foreach ($dir in @($RtlHdrDir, $RtlFrmDir)) {
    if (-not (Test-Path -LiteralPath $dir)) { continue }
    $hdrFiles += @(Get-ChildItem -LiteralPath $dir -Recurse -File -Filter *.h)
}
# ⚠ 枚举顺序**必须定死**（排序）：同名出口的原型取「第一次见到的那份」，而
# Get-ChildItem 的顺序在两次运行之间会变 ⇒ 表会变（踩过一次：快照漂 396 行，
# 而且两次漂的名单还不一样 ⇒ 表根本不可登记）。
$hdrFiles = @($hdrFiles | Sort-Object FullName)
$authorityHeaders = @($hdrFiles | ForEach-Object { $_.Name } | Sort-Object -Unique)

function Strip-Cmt([string]$S) {
    $t = ($S -replace '/\*.*?\*/', '') -replace '//.*$', ''
    $ix = $t.IndexOf('/*')
    if ($ix -ge 0) { $t = $t.Substring(0, $ix) }
    return $t
}
function Get-Balance([string]$S) {
    $n = 0
    foreach ($ch in $S.ToCharArray()) {
        if ($ch -eq '(' -or $ch -eq '[') { $n += 1 } elseif ($ch -eq ')' -or $ch -eq ']') { $n -= 1 }
    }
    return $n
}
function Split-Params([string]$Params) {
    $p = ($Params -replace '\s+', ' ').Trim()
    if ($p -eq '' -or $p -eq 'void') { return @() }
    $depth = 0; $cur = ''; $items = @()
    foreach ($ch in $p.ToCharArray()) {
        if ($ch -eq '(' -or $ch -eq '[') { $depth += 1 } elseif ($ch -eq ')' -or $ch -eq ']') { $depth -= 1 }
        if ($ch -eq ',' -and $depth -eq 0) { $items += $cur; $cur = '' } else { $cur += $ch }
    }
    $items += $cur
    $out = @()
    foreach ($it in $items) {
        $t = (($it -replace '\bconst\s+', ' ') -replace '\b(unsigned|signed)\s+', ' ').Trim()
        if ($t -eq '' -or $t -eq 'void' -or $t -eq '...') { continue }
        $out += $t
    }
    # ⚠ 逗号一元运算符：`return $out` 在**只有一个形参**时会被 PowerShell 展平成一条字符串
    # ⇒ 外面 `$params.Count` 变成**字符串长度**，形参位按字符扫过去，整枚出口掉出表外。
    return ,$out
}
# 形参「类型 + 名字」→ 类型记号。⚠ 先把**形参名**摘掉（"BSTR s" ⇒ "BSTR"），
# 摘不干净的那一格按「类型就是它自己」处理。
function Formal-Kind([string]$T) {
    $t = ($T -replace '\s+', ' ').Trim()
    $ty = ($t -replace '\b[A-Za-z_][A-Za-z0-9_]*\s*$', '').Trim()
    if ($ty -eq '') { $ty = $t }
    if ($ty -match '\bBSTR\s*\*$') { return 'BSTRPTR' }
    if ($ty -match '\bBSTR\s*$') { return 'BSTR' }
    if ($ty -match '\bvb6_VARIANT\s*\*$') { return 'VARIANTPTR' }
    if ($ty -match '\bvb6_VARIANT\s*$') { return 'VARIANT' }
    return ''
}

$protos = @{}          # 出口名 -> 形参类型数组
$retKind = @{}         # 出口名 -> 返回档（'BSTR' / 'VARIANT' / ''）
$surface = @{}         # "出口名|形参序号" -> 记号
foreach ($f in $hdrFiles) {
    $lines = [System.IO.File]::ReadAllLines($f.FullName)
    $i = 0
    while ($i -lt $lines.Count) {
        $ln = $lines[$i]
        if ($ln -eq '' -or $ln -match '^[ \t#/]') { $i += 1; continue }
        $buf = Strip-Cmt $ln
        $k = $i
        while ((Get-Balance $buf) -gt 0 -and ($k + 1 -lt $lines.Count)) {
            $k += 1
            $buf += ' ' + (Strip-Cmt $lines[$k]).Trim()
        }
        $open = $buf.IndexOf('(')
        $close = $buf.LastIndexOf(')')
        if ($open -lt 0 -or $close -le $open) { $i = $k + 1; continue }
        $head = $buf.Substring(0, $open).Trim()
        $m = [regex]::Match($head, '([A-Za-z_][A-Za-z0-9_]*)\s*$')
        if (-not $m.Success) { $i = $k + 1; continue }
        $name = $m.Groups[1].Value
        if ($name -notlike 'vb6_*') { $i = $k + 1; continue }
        $params = Split-Params $buf.Substring($open + 1, $close - $open - 1)
        if ($params.Count -eq 0 -and $head -notmatch '\bBSTR\b' -and $head -notmatch '\bvb6_VARIANT\b') {
            $i = $k + 1; continue
        }
        if (-not $protos.ContainsKey($name)) {
            $protos[$name] = $params
            # 返回档 = 头里那枚类型（BSTR / vb6_VARIANT）—— 「返回侧所有权」那一轴靠它取面
            $h2 = ($head -replace '\b(static|inline|__inline|extern)\b', '').Trim()
            $h2 = ($h2 -replace '\b[A-Za-z_][A-Za-z0-9_]*\s*$', '').Trim()
            if ($h2 -match '\bBSTR\s*$') { $retKind[$name] = 'BSTR' }
            elseif ($h2 -match '\bvb6_VARIANT\s*$') { $retKind[$name] = 'VARIANT' }
            else { $retKind[$name] = '' }
            for ($pi = 0; $pi -lt $params.Count; $pi++) {
                $kind = Formal-Kind $params[$pi]
                if ($kind -eq 'BSTR' -or $kind -eq 'VARIANT') { $surface["$name|$pi"] = $kind }
            }
        }
        $i = $k + 1
    }
}
Write-Host ("KEEPS 权威出口面: " + $authorityHeaders.Count + " 个头 / " + $protos.Count +
            " 个原型 / " + $surface.Count + " 枚按值带权形参位")

# ---------------------------------------------------------------- RTL 体
$bodyOf = @{}
$globalsOf = @{}       # 文件级变量名（每枚出口记它所在文件的文件级变量）
foreach ($f in (Get-ChildItem -LiteralPath (Join-Path $Root 'src\rtl') -Recurse -File -Include *.c, *.h, *.inc | Sort-Object FullName)) {
    $txt = [System.IO.File]::ReadAllText($f.FullName)
    # 文件级变量：不在任何函数体内的 `static`/普通全局声明
    $gvars = @{}
    foreach ($gm in [regex]::Matches($txt, '(?m)^\s*(?:static\s+)?(?:const\s+)?[A-Za-z_][A-Za-z0-9_ \*]*\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\[[^\]]*\])?\s*(?:=[^;]*)?;')) {
        $gvars[$gm.Groups[1].Value] = $true
    }
    foreach ($mm in [regex]::Matches($txt, '(?s)(?:^|\n)\s*(?:static\s+)?[A-Za-z_][A-Za-z0-9_ \*]*\b(vb6_[A-Za-z0-9_]+)\s*\(([^;{]*)\)\s*\{')) {
        $nm = $mm.Groups[1].Value
        if ($bodyOf.ContainsKey($nm)) { continue }
        $start = $mm.Index + $mm.Length
        $depth = 1
        $j = $start
        while ($j -lt $txt.Length -and $depth -gt 0) {
            if ($txt[$j] -eq '{') { $depth += 1 } elseif ($txt[$j] -eq '}') { $depth -= 1 }
            $j += 1
        }
        $bodyOf[$nm] = $txt.Substring($start, $j - $start)
        $globalsOf[$nm] = $gvars
        # **内部静态助手的形参表**（不在公开头里 ⇒ 不进出口面，但**转手判定**要用：
        # 下游是谁、它那一格是带权位不是，答不出就把本枚降成 U）
        if (-not $protos.ContainsKey($nm)) {
            $dp = Split-Params $mm.Groups[2].Value
            if ($dp.Count -gt 0) { $protos[$nm] = $dp }
        }
    }
}

# ---------------------------------------------------------------- 分类
$script:clsClass = ''
$script:clsEscape = ''
$script:clsMemo = @{}
$script:clsStack = @{}

function Get-LocalNames([string]$body) {
    $names = @{}
    foreach ($lm in [regex]::Matches($body, '(?m)^\s*(?:const\s+)?[A-Za-z_][A-Za-z0-9_]*(?:\s*\*)?(?:\s+[A-Za-z_][A-Za-z0-9_]*)*\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:[=;\[])' )) {
        $names[$lm.Groups[1].Value] = $true
    }
    return $names
}

# 数据持久化的**落点**（第二条轴的 sink）：这些 API 一调，数据就落到一个比调用活得久的地方。
# 注册表写 / 文件写 / 剪贴板。**这一轴不参与回收**（指针轴才是回收看的那一轴）。
$script:sinkRx = '(RegSetValueExW|RegCreateKeyExW|RegSetKeyValueW|WriteFile|WriteConsoleW|fwrite|fputs|fputws|SetClipboardData)'

function Classify([string]$name, [int]$idx) {
    $key = "$name|$idx"
    if ($script:clsMemo.ContainsKey($key)) {
        $m = $script:clsMemo[$key]
        $script:clsClass = $m[0]; $script:clsEscape = $m[1]
        return
    }
    if ($script:clsStack.ContainsKey($key)) {   # 成环 ⇒ 答不出（**并把它记成 U**：
        # 不记的话，这枚的答案就取决于「从环的哪一头进来」⇒ 同一份源码两次算两样）
        $script:clsClass = 'U'; $script:clsEscape = 'cycle'
        $script:clsMemo[$key] = @('U', 'cycle')
        return
    }
    $script:clsStack[$key] = $true
    $script:clsClass = 'U'; $script:clsEscape = 'unknown'

    if (-not $bodyOf.ContainsKey($name)) { $script:clsEscape = 'nobody' }
    elseif (-not $protos.ContainsKey($name)) { $script:clsEscape = 'noproto' }
    else {
        $params = $protos[$name]
        if ($idx -ge $params.Count) { $script:clsEscape = 'noproto' }
        else {
            $nm = [regex]::Match($params[$idx], '([A-Za-z_][A-Za-z0-9_]*)\s*$')
            if (-not $nm.Success) { $script:clsEscape = 'noname' }
            else {
                $arg = $nm.Groups[1].Value
                $body = $bodyOf[$name]
                $locals = Get-LocalNames $body
                $gvars = $globalsOf[$name]
                $esc = [System.Collections.Generic.List[string]]::new()
                $why = [System.Collections.Generic.List[string]]::new()

                # (1) 整枚赋给一个**非局部**的左值（结构体成员 / 文件级全局）⇒ store
                foreach ($mm in [regex]::Matches($body, '(?m)([A-Za-z_][A-Za-z0-9_]*(?:\s*(?:->|\.)\s*[A-Za-z_][A-Za-z0-9_]*)+)\s*=\s*\(?\s*' + [regex]::Escape($arg) + '\s*\)?\s*;')) {
                    $esc.Add('store') | Out-Null; $why.Add('field') | Out-Null; break
                }
                if ($esc.Count -eq 0) {
                    foreach ($mm in [regex]::Matches($body, '(?m)(?<![*\w])([A-Za-z_][A-Za-z0-9_]*)\s*=\s*\(?\s*' + [regex]::Escape($arg) + '\s*\)?\s*;')) {
                        $lhs = $mm.Groups[1].Value
                        if ($lhs -eq $arg) { continue }
                        if ($locals.ContainsKey($lhs)) {
                            # 赋给局部：那一枚局部自己逃不逃？（复制 vs 存指针）
                            if ([regex]::IsMatch($body, '(?m)return\s+\(?\s*' + [regex]::Escape($lhs) + '\s*\)?\s*;')) {
                                $esc.Add('ret') | Out-Null; $why.Add('local-returned') | Out-Null
                            } elseif ([regex]::IsMatch($body, '(?m)&' + [regex]::Escape($lhs) + '\b')) {
                                $esc.Add('write') | Out-Null; $why.Add('local-address') | Out-Null
                            } else {
                                $why.Add('local-copy') | Out-Null
                            }
                        } elseif ($gvars.ContainsKey($lhs)) {
                            $esc.Add('store') | Out-Null; $why.Add('global') | Out-Null
                        } else {
                            $esc.Add('U') | Out-Null; $why.Add('unknown-lvalue') | Out-Null
                        }
                        break
                    }
                }
                # (2) 原样返回（借出去，不是副本）
                if ([regex]::IsMatch($body, '(?m)return\s+\(?\s*' + [regex]::Escape($arg) + '\s*\)?\s*;')) {
                    if (-not $esc.Contains('ret')) { $esc.Add('ret') | Out-Null; $why.Add('returned') | Out-Null }
                }
                # (3) 取地址（经它写出去），或**经解引用写出去**（`*slot = arg` ⇒ 形参
                #     那一格是可写槽，调用方拿到的就是这枚指针）
                if ([regex]::IsMatch($body, '&' + [regex]::Escape($arg) + '\b')) {
                    if (-not $esc.Contains('write')) { $esc.Add('write') | Out-Null; $why.Add('address') | Out-Null }
                }
                if ([regex]::IsMatch($body, '\*\s*[A-Za-z_][A-Za-z0-9_]*\s*=\s*\(?\s*' + [regex]::Escape($arg) + '\s*\)?\s*;')) {
                    if (-not $esc.Contains('write')) { $esc.Add('write') | Out-Null; $why.Add('deref-write') | Out-Null }
                }
                # (4) 转手给另一枚 vb6_* 出口 ⇒ 下游说了算
                if ($esc.Count -eq 0 -or $esc.Contains('U')) {
                    foreach ($mm in [regex]::Matches($body, '\b(vb6_[A-Za-z0-9_]+)\s*\(([^;]*?)\)')) {
                        $callee = $mm.Groups[1].Value
                        if ($callee -eq $name) { continue }
                        $callArgs = Split-Params $mm.Groups[2].Value
                        # 这一枚形参在**调用点的第几格**（按位置，不按名字 —— 下游的形参名
                        # 与上游的形参名常常不是同一个字，按名字配会配不上）
                        $didx = -1
                        for ($q = 0; $q -lt $callArgs.Count; $q++) {
                            if ([regex]::IsMatch($callArgs[$q], '\b' + [regex]::Escape($arg) + '\b')) { $didx = $q; break }
                        }
                        if ($didx -lt 0) { continue }   # 这一枚形参没进这通调用
                        if (-not $bodyOf.ContainsKey($callee)) {
                            $esc.Add('U') | Out-Null; $why.Add("handoff-unknown:$callee") | Out-Null; break
                        }
                        if (-not $protos.ContainsKey($callee) -or $didx -ge $protos[$callee].Count) {
                            $esc.Add('U') | Out-Null; $why.Add("handoff-unknown-slot:$callee") | Out-Null; break
                        }
                        $downKind = Formal-Kind $protos[$callee][$didx]
                        if ($downKind -ne 'BSTR' -and $downKind -ne 'VARIANT') {
                            # 下游那一格不是带权位（转手过去的是别的类型）⇒ 本枚不必管
                            $why.Add("handoff-to-plain:$callee") | Out-Null; break
                        }
                        Classify $callee $didx
                        if ($script:clsClass -eq '1') { $esc.Add('store') | Out-Null; $why.Add("handoff-keeps:$callee") | Out-Null }
                        elseif ($script:clsClass -eq 'U') { $esc.Add('U') | Out-Null; $why.Add("handoff-unknown:$callee") | Out-Null }
                        else { $why.Add("handoff-reads:$callee") | Out-Null }
                        break
                    }
                }
                if ($esc.Contains('store') -or $esc.Contains('ret') -or $esc.Contains('write')) {
                    $script:clsClass = '1'
                    $script:clsEscape = (($esc | Sort-Object -Unique) -join '+')
                } elseif ($esc.Contains('U')) {
                    $script:clsClass = 'U'
                    $script:clsEscape = (($why | Sort-Object -Unique) -join ';')
                } else {
                    $script:clsClass = '0'
                    $script:clsEscape = (($why | Sort-Object -Unique) -join ';')
                }
            }
        }
    }
    $script:clsMemo[$key] = @($script:clsClass, $script:clsEscape)
    $script:clsStack.Remove($key)
}

# ---------------------------------------------------------------- 借用轴（返回侧）
# 这一枚出口**交出去的指针是借来的**（不归调用方所有 ⇒ 调用方不许放它）。它不是形参位的性质，
# 而是**出口**的性质 ⇒ 单独一节，不塞进上面那张按形参位排的表。
# 权威 = 编译器侧那枚借用名单 `kBorrowedBstrExports167`（src/backend/cgen_util_ctrl.cpp），
# 那份名单本身由 `check_bstr_borrow_slot.ps1` 的 Q6 从 RTL **双向推导**钉住（导出面全、无遗漏）。
# 为什么这一节必须在：撤回那一版（9546dbfe / 6bf5989 / be5d6abf）**只问了 mint 那一侧**
# （「这一格是不是现造的」），没问 borrow 那一侧（「这一格是借来的吗」）⇒ 它把
# `vb6_BSTR_Concat` 换成 `vb6_BSTR_ConcatFree`，于是**放掉了借来的 `vb6_ErrDescription()`**
# （RTL 全局 vb6_err 的指针）⇒ 全局悬空 ⇒ 读数变堆垃圾；又对**借来的变体**
# （`vb6_VariantArrayGetND` 的元素，指向数组里那块）发 `vb6_VariantClear` ⇒ 数组被清坏。
# 本机实测：那一版在 `test_bstr_borrow_slot` 上 W8b/W9a/W9b 读出堆垃圾，在
# `test_host_decimal`(x86) 上直接 0xC0000374（堆损坏）⇒ **借用族一律不许进提名面**。
$borrowAuthority = Join-Path $Root 'src\backend\cgen_util_ctrl.cpp'
$borrowedExports = @()
if (Test-Path -LiteralPath $borrowAuthority) {
    $authTxt = [System.IO.File]::ReadAllText($borrowAuthority)
    $am = [regex]::Match($authTxt, 'kBorrowedBstrExports167\[\]\s*=\s*\{([\s\S]*?)\}')
    if ($am.Success) {
        foreach ($nm in [regex]::Matches($am.Groups[1].Value, '"(vb6_[A-Za-z0-9_]+)"')) {
            $borrowedExports += $nm.Groups[1].Value
        }
    }
}
if ($borrowedExports.Count -eq 0) {
    $bad += 'K11 编译器侧的借用名单 kBorrowedBstrExports167 读不出来 —— 借用族答不了 ⇒ 停机'
}

# ---------------------------------------------------------------- 返回侧的所有权（账 #456 §B227 第五格）
# 「这一枚出口交出去的是不是**新造且归调用方所有**」——
#   owned = 出口自己新造的（调用方可以放它）   —— 返回表达式一路落到一枚 **mint**
#   alias = 交出去的是**别名**（指着调用方的内存或一个全局）—— 调用方**不许**放它
#   U      = 答不出（安全默认：当作不是新造的 ⇒ 不许回收）
# 为什么必须**逐枚**答：编译器侧那三枚借用名单（kBorrowedBstrExports167）**只覆盖 BSTR 返回**，
# 盖不到数组元素读那一族（`vb6_VariantArrayGet`/`GetND`/`GetVal` 交出去的是变体里那格
# **指着数组里那一枚**的 BSTR）⇒ 拿「三枚名单」当全部，就是撤回那一版的病根。
$script:mintRx = '(SysAllocString|SysAllocStringLen|vb6_BSTR_FromStr|vb6_BSTR_Concat|vb6_BSTR_ConcatFree|vb6_BSTR_FromBSTR|vb6_BSTR_Empty|vb6_VariantCopy|vb6_VariantFromValue|vb6_VariantEmpty|vb6_Format)'
$script:aliasRx = '([A-Za-z_][A-Za-z0-9_]*)\s*(->|\.)\s*[A-Za-z0-9_]+'

function Get-ReturnExpr([string]$name) {
    if (-not $bodyOf.ContainsKey($name)) { return '' }
    $body = $bodyOf[$name]
    $ms = [regex]::Matches($body, '(?m)return\s+([^;]+);')
    if ($ms.Count -eq 0) { return '' }
    return $ms[$ms.Count - 1].Groups[1].Value.Trim()
}

# 逐枚分类（带传递闭包：返回的是另一枚出口 ⇒ 继承那枚的分类；成环 ⇒ U）
$script:ownMemo = @{}
function Classify-Return([string]$name, [System.Collections.Generic.HashSet[string]]$seen) {
    if ($script:ownMemo.ContainsKey($name)) {
        $m = $script:ownMemo[$name]
        return @($m[0], $m[1])
    }
    if ($seen.Contains($name)) { return @('U', 'cycle') }          # 成环 ⇒ 答不出
    $seen.Add($name) | Out-Null
    $kind = 'U'
    $why = 'unknown'
    $expr = Get-ReturnExpr $name
    if ($expr -eq '') {
        $kind = 'U'; $why = 'no-return'
    } elseif ($expr -match ('^' + $script:mintRx + '\s*\(')) {
        $kind = 'owned'; $why = 'mint'
    } elseif ($expr -match ('\b' + $script:mintRx + '\s*\(')) {
        $kind = 'U'; $why = 'mixed'          # 返回表达式里混着 mint 与别的东西 ⇒ 答不出
    } elseif ($expr -match $script:aliasRx) {
        $kind = 'alias'; $why = 'alias'
    } elseif ($expr -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*$' -and $bodyOf.ContainsKey($name)) {
        # 返回的是**本地结构体** ⇒ 逐格看它的字段：指针那一格是从 mint 来的（新造，归我们所有），
        # 还是从**形参**读来的（别名！）—— 数组元素读那一族就在这里：
        #   result.bstrVal = ((BSTR*)data)[off];   ← 不复制，指着数组里那一枚 ⇒ alias
        #   result = ((vb6_VARIANT*)data)[off];     ← 浅拷，BSTR 那格仍共享 ⇒ alias
        $local = $Matches[1]
        $pnames = @()
        foreach ($pp in $protos[$name]) {
            $mm2 = [regex]::Match($pp, '([A-Za-z_][A-Za-z0-9_]*)\s*$')
            if ($mm2.Success) { $pnames += $mm2.Groups[1].Value }
        }
        # 账 #463 §B164 研究（台账 #462/#463 两行）：这一支原来只有「逐格看字段」一问，**漏了三形** ——
        #   ① 返回的就是**形参本身**（vb6_VariantIdentity: `return v;`）⇒ 别名
        #   ② 返回的局部量被**整份**赋成形参（vb6_CDec 的 DECIMAL 那一支: `result = v;`）⇒ 别名
        #      （那一支交出去的 DECIMAL 指着**形参**那一份 ⇒ 放掉它 = 放掉别人手里的东西）
        #   ③ 返回的局部量是**在形参之间挑一个**（vb6_IIfBSTR: `result = cond ? truepart : falsepart;`）⇒ 别名
        # 这三形以前一律落到 'struct-mint' ⇒ 判成 owned ⇒ [own] 表里多了四枚别名
        #   （vb6_IIfBSTR / vb6_IIfVariant / vb6_CDec / vb6_VariantIdentity）
        # ⇒ 回收刀会把「别人手里的那一枚」当现造的放掉 ⇒ 悬空 / 双放。判据见台账 #463。
        $aliasKind = ''
        if ($pnames -contains $local) {
            $aliasKind = 'return-param'
        } else {
            $wholeHit = $false
            foreach ($pn in $pnames) {
                if ([regex]::IsMatch($bodyOf[$name],
                        ('(?m)\b' + [regex]::Escape($local) + '\s*=\s*&?\s*' + [regex]::Escape($pn) + '\s*;'))) {
                    $wholeHit = $true; break
                }
            }
            if ($wholeHit) {
                $aliasKind = 'assigned-from-param'
            } else {
                # ③ 在形参之间挑一个：右值提到形参、**没有调用**、且去掉形参与运算符后**不剩东西**
                foreach ($am in [regex]::Matches($bodyOf[$name],
                            '(?m)\b' + [regex]::Escape($local) + '\s*=\s*([^;]+);')) {
                    $rhs = $am.Groups[1].Value
                    if ($rhs -match '\(') { continue }
                    $mentions = $false
                    $rest = $rhs
                    foreach ($pn in $pnames) {
                        if ($rhs -match ('\b' + [regex]::Escape($pn) + '\b')) { $mentions = $true }
                        $rest = [regex]::Replace($rest, '\b' + [regex]::Escape($pn) + '\b', '')
                    }
                    if (-not $mentions) { continue }
                    if ((($rest -replace '[\s?:&|!,]', '')) -eq '') { $aliasKind = 'selected-from-param'; break }
                }
            }
        }
        if ($aliasKind -ne '') {
            $kind = 'alias'; $why = $aliasKind
        } else {
        $aliasHit = $false; $unknownHit = $false
        foreach ($am in [regex]::Matches($bodyOf[$name], '(?m)\b' + [regex]::Escape($local) + '\s*\.\s*([A-Za-z0-9_]+)\s*=\s*([^;]+);')) {
            $field = $am.Groups[1].Value
            $rhs = $am.Groups[2].Value
            if ($field -notmatch 'bstrVal|parray|pvarVal|pdispVal|piface|punkVal|pszVal|psafeArray|pgv') { continue }
            if ($rhs -match ('\b' + $script:mintRx + '\s*\(')) { continue }
            $hit = $false
            foreach ($pn in $pnames) {
                if ($rhs -match ('\b' + [regex]::Escape($pn) + '\b')) { $hit = $true; break }
            }
            if ($hit) { $aliasHit = $true } else { $unknownHit = $true }
        }
        if ($aliasHit) { $kind = 'alias'; $why = 'field-of-param' }
        elseif ($unknownHit) { $kind = 'U'; $why = 'field-unknown' }
        else { $kind = 'owned'; $why = 'struct-mint' }
        }
    } else {
        $cm = [regex]::Match($expr, '^\s*(vb6_[A-Za-z0-9_]+)\s*\(')
        if ($cm.Success -and $bodyOf.ContainsKey($cm.Groups[1].Value)) {
            $sub = Classify-Return $cm.Groups[1].Value $seen
            $kind = $sub[0]; $why = "via:" + $cm.Groups[1].Value
        } else {
            $kind = 'U'; $why = 'unknown'
        }
    }
    $script:ownMemo[$name] = @($kind, $why)
    return @($kind, $why)
}

$ownOwned = @(); $ownAlias = @(); $ownUnknown = @(); $ownClass = @{}
$ownVariant = @()
foreach ($n in ($retKind.Keys | Sort-Object)) {
    if ($retKind[$n] -ne 'BSTR' -and $retKind[$n] -ne 'VARIANT') { continue }
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $cls = Classify-Return $n $seen
    $ownClass[$n] = $cls[0]
    if ($cls[0] -eq 'owned') { $ownOwned += $n } elseif ($cls[0] -eq 'alias') { $ownAlias += $n } else { $ownUnknown += $n }
}
$ownVariant = @($ownOwned | Where-Object { $retKind[$_] -eq 'VARIANT' })
$aliasExits = $ownAlias
$closedExits = @('vb6_SaveSetting', 'vb6_GetSetting', 'vb6_DeleteSetting', 'vb6_GetAllSettings')

# ---------------------------------------------------------------- 消费轴（账 #463 §B164 研究新增）
# 「消费者会不会**自己放掉**自己的形参」—— 与 keeps_pointer（指针有没有**带出**调用）是**另一轴**：
#     keeps_pointer = '1' ⇒ 消费者把指针留在外面 ⇒ 调用方不能放（否则双放/悬空）
#     consumes      = 1   ⇒ 消费者自己放掉这一格 ⇒ 调用方**也不能**再放（否则**双放** ⇒ 堆损坏）
# 判据 = 函数体里对自己的形参调 vb6_BSTR_Free / SysFreeString / vb6_VariantClear。
# 「条件性」= 那一放**看运行期分支**（if/else 的条件不是那个形参本身）⇒ 调用方**无从知道**
#     ⇒ 那一格**不许**提名（IIf 那一族：没选中的那一支被放掉，而调用方不知道会放哪一支）。
$consumesRows = @{}
foreach ($nm in $bodyOf.Keys) {
    $plist = @()
    if ($protos.ContainsKey($nm)) { $plist = $protos[$nm] }
    $pnames = @()
    foreach ($pp in $plist) {
        $mm2 = [regex]::Match($pp, '([A-Za-z_][A-Za-z0-9_]*)\s*$')
        if ($mm2.Success) { $pnames += $mm2.Groups[1].Value }
    }
    if ($pnames.Count -eq 0) { continue }
    $body = $bodyOf[$nm]
    for ($pi = 0; $pi -lt $pnames.Count; $pi++) {
        $pn = $pnames[$pi]
        $rx = '(vb6_BSTR_Free|SysFreeString|vb6_VariantClear)\s*\(\s*&?\s*' + [regex]::Escape($pn) + '\s*\)'
        foreach ($fm in [regex]::Matches($body, $rx)) {
            $lineStart = $body.LastIndexOf("`n", [Math]::Max(0, $fm.Index - 1))
            $line = $body.Substring($lineStart + 1, $fm.Index - $lineStart)
            $cond = '0'
            $im = [regex]::Match($line, '\b(?:if|else)\b')
            if ($im.Success) {
                $cm2 = [regex]::Match($line, '\bif\s*\(([^)]*)\)')
                if ($cm2.Success) {
                    # 条件**就是**那个形参（空指针检查）⇒ 与分支无关 ⇒ 不算条件性
                    if ($cm2.Groups[1].Value.Trim() -ne $pn) { $cond = '1' }
                } else {
                    $cond = '1'      # else 分支 ⇒ 条件性
                }
            }
            $key = $nm + '|' + $pi
            if (-not $consumesRows.ContainsKey($key)) {
                $consumesRows[$key] = @($cond, $fm.Groups[1].Value, $pn)
            }
        }
    }
}
Write-Host ("消费轴（消费者自己放掉形参的那几格）: " + $consumesRows.Count + " 格")
$ownedTempExceptions = @(
    foreach ($k in ($rows.Keys | Sort-Object)) {
        if ($rows[$k][0] -ne '0') {
            $ps = $k -split '\|'
            [pscustomobject]@{ name = $ps[0]; slot = [int]$ps[1]; keeps = $rows[$k][0] }
        }
    }
)
Write-Host ("返回侧所有权: owned " + $ownOwned.Count + " / alias " + $ownAlias.Count + " / U " + $ownUnknown.Count)

# ---------------------------------------------------------------- 返回侧的两个锚（K14）
# 正控 = **新造**那一族（调用方可以放它）· 负控 = **别名**那一族（调用方不许放它）
$ownAnchors = @(
    @{ name = 'vb6_BSTR_FromStr';      want = 'owned'; why = 'mint: SysAllocString' },
    @{ name = 'vb6_BSTR_Concat';       want = 'owned'; why = 'mint: 新造串' },
    @{ name = 'vb6_CStr';              want = 'owned'; why = 'mint: 新造串' },
    @{ name = 'vb6_CStrLong';          want = 'owned'; why = 'mint: 新造串' },
    @{ name = 'vb6_VariantToString';   want = 'owned'; why = 'mint: 新造串' },
    @{ name = 'vb6_ErrDescription';    want = 'alias'; why = '全局 vb6_err 的指针' },
    @{ name = 'vb6_ErrSource';         want = 'alias'; why = '全局 vb6_err 的指针' },
    @{ name = 'vb6_VariantString';     want = 'alias'; why = '变体那格 BSTR 就是形参本身' },
    @{ name = 'vb6_VariantArrayGet';   want = 'alias'; why = 'bstrVal 指着数组里那一枚' },
    @{ name = 'vb6_VariantArrayGetND'; want = 'alias'; why = 'bstrVal 指着数组里那一枚' },
    @{ name = 'vb6_VariantArrayGetVal';want = 'alias'; why = 'bstrVal 指着数组里那一枚' }
)
foreach ($a in $ownAnchors) {
    if (-not $ownClass.ContainsKey($a.name)) {
        $bad += ("K14 返回侧锚点缺失（表里没有这一枚）: " + $a.name + " —— " + $a.why)
        continue
    }
    if ($ownClass[$a.name] -ne $a.want) {
        $bad += ("K14 返回侧锚点 " + $a.name + " 判成 " + $ownClass[$a.name] + "，期望 " + $a.want + "（" + $a.why + "）")
    }
}
# K14b **档号**：mint 交出来的是 BSTR 还是 VARIANT —— 只由出口的返回档定
#     （源表达式的 VB 档不作数：`Len(v)` 的源档是 Variant、交出去的是 BSTR ⇒ 照源档声明就 C2440）。
foreach ($mk in @(@{ name = 'vb6_BSTR_FromStr';    want = 'B'; why = 'BSTR 档的 mint' },
                   @{ name = 'vb6_VariantToString'; want = 'B'; why = '名字带 Variant、交出的却是 BSTR' },
                   @{ name = 'vb6_VariantLong';     want = 'V'; why = 'VARIANT 档的 mint' },
                   @{ name = 'vb6_CDec';            want = 'V'; why = 'VARIANT 档的 mint' })) {
    $k = $retKind[$mk.name]
    if (-not $k) {
        $bad += ("K14b 档号锚点缺失（表里没有这一枚）: " + $mk.name)
        continue
    }
    $want = if ($k -eq 'BSTR') { 'B' } else { 'V' }
    if ($want -ne $mk.want) {
        $bad += ("K14b 档号锚点 " + $mk.name + " 的返回档是 " + $k + "，期望 " + $mk.want + "（" + $mk.why + "）")
    }
}
# K15 两条轴**不许互相顶替**：借用族三枚必须是 alias（不是 owned）——
#     「这一枚交出去的是借来的」与「这一枚交出去的是新造的」是同一件事的两面，
#     两面判成一样 ⇒ 有一面判错了（撤回那一版就是只判了 mint 那一面）。
foreach ($b in $borrowedExports) {
    if ($ownClass.ContainsKey($b) -and $ownClass[$b] -eq 'owned') {
        $bad += ("K15 借用族 " + $b + " 在返回侧被判成 owned（交出去的是借来的指针，不许当新造）")
    }
}
# K16 两条轴**不许互相顶替**：alias 那一族与 owned 那一族必须**不相交**
#     （同一枚出口不能既是「新造」又是「别名」—— 判成一样就说明有一面判错了）
$overlapOwn = @($ownOwned | Where-Object { $ownAlias -contains $_ })
if ($overlapOwn.Count -gt 0) {
    $bad += ("K16 " + $overlapOwn.Count + " 枚出口同时被判成 owned 与 alias（两轴判成一样 = 有一面判错了）: " +
             (($overlapOwn | Select-Object -First 5) -join ', '))
}

$rows = @{}
# ⚠ 遍历顺序必须**定死**（排序）：分类里有成环判定，顺序不同会得出不同的表
# （踩过一次： hashtable 的 Keys 顺序变一次，快照就漂 396 行 ⇒ 表变成不可登记的）。
foreach ($key in ($surface.Keys | Sort-Object)) {
    $parts = $key -split '\|'
    $nm = $parts[0]
    $idx = [int]$parts[1]
    Classify $nm $idx
    $kp = $script:clsClass
    $esc = $script:clsEscape
    # ---- 第二条轴：persists_data（数据持久化，**不参与回收**）----
    # 这枚形参带进去的**数据**有没有落到一个比调用活得久的地方。判据 = 它的数据有没有
    # 流进一枚 sink（注册表写 / 文件写 / 剪贴板）。证不出就 U（这一轴不卡任何东西）。
    $persists = '0'
    if ($bodyOf.ContainsKey($nm) -and $protos.ContainsKey($nm) -and $idx -lt $protos[$nm].Count) {
        $arg = [regex]::Match($protos[$nm][$idx], '([A-Za-z_][A-Za-z0-9_]*)\s*$').Groups[1].Value
        $body = $bodyOf[$nm]
        $hasSink = [regex]::IsMatch($body, $script:sinkRx)
        $argInSink = $false
        if ($hasSink) {
            # ⚠ 落点的**实参表**要取到「那一整条语句」为止：非贪婪的 `[^;]*?\)` 会在
            # `(const BYTE*)setting` 的那个 ')' 处停下 ⇒ 形参名压根不在捕获里（实测
            # SaveSetting 的四格因此全落成 U）。改成「从 sink 那一枚起，往后看同一条语句」。
            foreach ($sm in [regex]::Matches($body, $script:sinkRx)) {
                $tail = $body.Substring($sm.Index, [Math]::Min(400, $body.Length - $sm.Index))
                $stop = $tail.IndexOf(';')
                if ($stop -ge 0) { $tail = $tail.Substring(0, $stop) }
                if ([regex]::IsMatch($tail, '\b' + [regex]::Escape($arg) + '\b')) { $argInSink = $true; break }
            }
        }
        if ($argInSink) { $persists = '1' } elseif ($hasSink) { $persists = 'U' }
    }
    # ---- 第三列 gate：closed 的行**不进提名面** ----
    # 注册表那一族挂 closed 的理由是**它那一族的运行期破坏未归因**（CI 上 R2-vt=1 而不是
    # 8204，连注册表键都没建起来）⇒ 未归因 ⇒ 不许提名。**不是**因为它指针轴上不安全
    # （它在指针轴上是 0：RegSetValueExW 与 vb6_RegBuildKey 都复制）。
    $gate = 'open'
    if ($nm -in @('vb6_SaveSetting', 'vb6_GetSetting', 'vb6_DeleteSetting', 'vb6_GetAllSettings')) {
        $gate = 'closed'
    }
    $rows[$key] = @($kp, $esc, $persists, $gate, $surface[$key])
}
$k0 = @($rows.Values | Where-Object { $_[0] -eq '0' }).Count
$k1 = @($rows.Values | Where-Object { $_[0] -eq '1' }).Count
$ku = @($rows.Values | Where-Object { $_[0] -eq 'U' }).Count
# 提名面 = keeps_pointer 0 ∧ gate open（**不看 persists_data**：数据持久化那条轴不参与回收）
$nomRows = @($rows.Keys | Where-Object { $rows[$_][0] -eq '0' -and $rows[$_][3] -eq 'open' })
$closedRows = @($rows.Keys | Where-Object { $rows[$_][3] -eq 'closed' })
Write-Host ("KEEPS 表（指针轴）: keeps_pointer=0 $k0 / =1 $k1 / U $ku （共 " + $rows.Count + " 行）")
Write-Host ("提名面（keeps_pointer=0 ∧ gate=open）: " + $nomRows.Count + " 行 ; closed " + $closedRows.Count + " 行")

# ---------------------------------------------------------------- 三锚（按裁决）
# 正控 = **指针轴**上真存指针的四族（用户裁决）：vb6_ErrRaise / vb6_RaiseError（存进全局
#   vb6_err）、vb6_ho_setVariantBstr（存进 VARIANT 槽）、vb6_BSTR_AssignMove（经解引用写出去）、
#   vb6_ArraySetVariant / vb6_LoadResData / vb6_IIfVariant（经可写槽接手）⇒ keeps_pointer = 1
# 负控 = 只读那一枚（vb6_Len）⇒ keeps_pointer = 0
# 灰控 = 证不出那一枚（vb6_DecSetTag 第 0 格）⇒ keeps_pointer = U（**必须保持 U**）
$anchors = @(
    @{ key = 'vb6_ErrRaise|1';         want = '1'; why = 'store: 存进全局 vb6_err' },
    @{ key = 'vb6_ErrRaise|2';         want = '1'; why = 'store: 存进全局 vb6_err' },
    @{ key = 'vb6_RaiseError|1';       want = '1'; why = 'store: 存进全局 vb6_err' },
    @{ key = 'vb6_ho_setVariantBstr|1'; want = '1'; why = 'store: 存进 VARIANT 槽' },
    @{ key = 'vb6_BSTR_AssignMove|1';  want = '1'; why = 'write: 经解引用写出去' },
    @{ key = 'vb6_ArraySetVariant|2';  want = '1'; why = 'write: 经可写槽接手' },
    @{ key = 'vb6_LoadResData|0';      want = '1'; why = 'write: 经可写槽接手' },
    @{ key = 'vb6_IIfVariant|1';        want = '1'; why = 'ret+write' },
    @{ key = 'vb6_Len|0';              want = '0'; why = 'read-only: 只在调用内读' },
    @{ key = 'vb6_DecSetTag|0';        want = 'U'; why = '证不出: 转手进 vb6_DEC_VT' }
)
foreach ($a in $anchors) {
    if (-not $rows.ContainsKey($a.key)) {
        $bad += ("K1 锚点缺失（表里没有这一行）: " + $a.key + " —— " + $a.why)
        continue
    }
    if ($rows[$a.key][0] -ne $a.want) {
        $bad += ("K1 锚点 " + $a.key + " keeps_pointer = " + $rows[$a.key][0] + "，期望 " + $a.want +
                 "（" + $a.why + "；实测逃逸通道 = " + $rows[$a.key][1] + "）")
    }
}
# K8 注册表族：指针轴上是 0（复制证据完整）＋ **gate=closed**（CI 红未归因 ⇒ 不进提名面）
$regRows = @($rows.Keys | Where-Object { ($_ -split '\|')[0] -in @('vb6_SaveSetting', 'vb6_GetSetting', 'vb6_DeleteSetting', 'vb6_GetAllSettings') })
if ($regRows.Count -eq 0) {
    $bad += 'K8 表里没有注册表那一族 —— 那一族必须显式在册（哪怕是 closed）'
} else {
    foreach ($r in $regRows) {
        if ($rows[$r][0] -ne '0') {
            $bad += ("K8 注册表族 " + $r + " 的 keeps_pointer = " + $rows[$r][0] +
                     "，期望 0（RegSetValueExW 与 vb6_RegBuildKey 都复制 ⇒ 指针不留存）")
        }
        if ($rows[$r][3] -ne 'closed') {
            $bad += ("K8 注册表族 " + $r + " 的 gate = " + $rows[$r][3] +
                     "，期望 closed（CI 红未归因 ⇒ 那一族不进提名面）")
        }
    }
}
# K9 提名面：closed 的行**一律不许**出现在里面（这条是「停机」纪律的机械形式）
$leak = @($closedRows | Where-Object { $rows[$_][0] -eq '0' -and $nomRows -contains $_ })
if ($leak.Count -gt 0) {
    $bad += ("K9 提名面里混进了 closed 的行 " + $leak.Count + " 处: " + (($leak | Select-Object -First 4) -join ', '))
}
# K10 提名面**只看指针轴**：一条 persists_data=1 而 keeps_pointer=0 的行，必须是 closed
#     （数据持久化那条轴不参与回收 ⇒ 落在 open 集合里就是「按数据轴禁回收」= 走错轴）
$misaxis = @($rows.Keys | Where-Object { $rows[$_][2] -eq '1' -and $rows[$_][0] -eq '0' -and $rows[$_][3] -ne 'closed' })
if ($misaxis.Count -gt 0) {
    $bad += ("K10 " + $misaxis.Count + " 行 persists_data=1 却落在 open 集合里（数据轴不参与回收 ⇒ " +
             "按数据轴禁回收就是走错轴）: " + (($misaxis | Select-Object -First 4) -join ', '))
}
$grayRules = @('user-proc', 'late-bound-com')

# ---------------------------------------------------------------- 快照对账
$snap = @{}
$snapBorrow = @()
if (Test-Path -LiteralPath $TableFile) {
    foreach ($ln in [System.IO.File]::ReadAllLines($TableFile)) {
        if ($ln -eq '' -or $ln.StartsWith('#')) { continue }
        if ($ln.StartsWith('[borrow]')) {
            $snapBorrow += $ln.Substring(8).Trim()
            continue
        }
        # ⚠ 字段分隔**不许用 '|'**：键自己带一个 '|'（出口名|形参序号），用 '|' 分隔
        # 就把键切成两半 ⇒ 快照对账全表漂（踩过一次：漂 396 行，名单每跑一次还不一样）。
        # 用 TAB：键与逃逸通道里都不可能出现 TAB。
        $p = $ln -split "`t", 6
        if ($p.Count -lt 6) { continue }
        $snap[$p[0]] = @($p[1], $p[2], $p[3], $p[4], $p[5])
    }
}
if ($Bless) {
    $out = @()
    $out += '# owned_temp_keeps_164.txt —— 账 #454 §B164 形① 的 KEEPS 表（出口面出自 RTL 的权威头）'
    $out += '# 字段（TAB 分隔）: 出口名|形参序号 <TAB> keeps_pointer(0/1/U) <TAB> 逃逸通道 <TAB> persists_data(1/0/U) <TAB> gate(open/closed) <TAB> 形参档'
    $out += '# keeps_pointer = 指针/槽位/引用有没有带出调用（**回收只看这一轴**）'
    $out += '# persists_data = 这条形参的数据有没有落到比调用活得久的地方（**不参与回收**）'
    $out += '# gate=closed = **不进提名面**（注册表那一族：CI 红未归因 ⇒ 停机）'
    $out += '# 用户过程与晚绑定 COM 不在表里（不是具名出口）⇒ 按规则归 U，见 check_owned_temp_keeps.ps1'
    foreach ($k in ($rows.Keys | Sort-Object)) {
        $v = $rows[$k]
        $out += ($k + "`t" + $v[0] + "`t" + $v[1] + "`t" + $v[2] + "`t" + $v[3] + "`t" + $v[4])
    }
    $out += '# --- 借用轴（返回侧）：这一枚出口交出去的指针是借来的，不归调用方所有 ⇒ 不许放它 ---'
    foreach ($b in ($borrowedExports | Sort-Object)) { $out += ('[borrow] ' + $b) }
    $out += '# --- 返回侧所有权：owned = 这一枚交出去的是**新造且归调用方所有**（调用方可以放它）---'
    foreach ($o in ($ownOwned | Sort-Object)) { $out += ('[own] ' + $o + ' ' + $retKind[$o]) }
    $out += '# --- 返回侧所有权：alias = 交出去的是**别名**（指着调用方的内存或一个全局）⇒ 不许放它 ---'
    foreach ($a in ($ownAlias | Sort-Object)) { $out += ('[alias] ' + $a + ' ' + $retKind[$a]) }
    $out += '# --- 消费轴（账 #463 §B164 研究新增）：消费者**自己放掉**自己的形参的那几格 ---'
    $out += '# 字段: 出口名|形参序号 <TAB> 条件性(0/1) <TAB> 放掉它的调用 <TAB> 形参名'
    $out += '# 条件性=1 ⇒ 放不放**看运行期分支**（IIf 那一族）⇒ 调用方无从知道 ⇒ 不许提名'
    foreach ($k in ($consumesRows.Keys | Sort-Object)) {
        $v = $consumesRows[$k]
        $out += ($k + "`t" + $v[0] + "`t" + $v[1] + "`t" + $v[2])
    }
    [System.IO.File]::WriteAllLines($TableFile, $out, (New-Object System.Text.UTF8Encoding($false)))
    # ⚠ **只登记这张表**（推导出来的**证据**）；发码侧那份（src/common/owned_temp_keeps.hpp）
    #   是**手写**的出口，由哨兵 **K13** 双向对账 —— 与仓里 `kBorrowedBstrExports167`（手写）
    #   + `check_bstr_borrow_slot` Q6（推导后对账）同一形状。**不由本脚本改写那份**：
    #   两边都由本脚本生成的话，对账就成了自己跟自己比，K13 会永远绿。
    Write-Host ("已登记快照: " + $TableFile + " （" + $rows.Count + " 行 + 借用 " + $borrowedExports.Count +
                " 枚 + owned " + $ownOwned.Count + " 枚 + alias " + $ownAlias.Count + " 枚）")
    exit 0
}

# K11 借用轴：① 表里那一节必须与**编译器侧权威名单**逐字相同（双向）；
#      ② 借用族与按形参位排的这张表**不相交** —— 借来的那一枚是**出口**（交出去的东西不归
#      调用方所有），它**不可能**同时是「按值带权形参位」的接收者；两处若重了，说明有人把一枚
#      借来的出口当成了现造的临时量来回收（撤回那一版的病根）。
#      ③ **回收点那一侧必须再问一次借用权威**：`ConcatFree` 放掉左操作数、`VariantClear` 清一枚
#      变体，都只在「那一枚归调用方所有」时才成立 —— 这是回收刀那一侧的判据，本表只把名单钉住。
#      ⚠ 这里**故意没有**「借用族不许进提名面」那一格：借用族是**出口**（它们没有按值的带权形参位）
#      ⇒ 那一格在今天这张表上**恒为空**，写进去就是一条**永远不会红的判据**（负控一跑就露馅：
#      把借用族的 gate 打开，那一格照样绿）。它属于**回收点**那一侧，由回收刀自己的哨兵钉。
if ($borrowedExports.Count -gt 0) {
    $missingB = @($borrowedExports | Where-Object { $snapBorrow -notcontains $_ })
    $extraB = @($snapBorrow | Where-Object { $borrowedExports -notcontains $_ })
    if ($missingB.Count -gt 0) {
        $bad += ("K11 表的借用轴少了 " + $missingB.Count + " 枚（权威名单里有，表里没有）: " + ($missingB -join ', '))
    }
    if ($extraB.Count -gt 0) {
        $bad += ("K11 表的借用轴多了 " + $extraB.Count + " 枚（表里有，权威名单里没有）: " + ($extraB -join ', '))
    }
    # ② 不相交：借用族是**出口**，不该同时出现在「按值带权形参位」这张表里
    $overlap = @($rows.Keys | Where-Object { ($_ -split '\|')[0] -in $borrowedExports })
    if ($overlap.Count -gt 0) {
        $bad += ("K11 借用族 " + $overlap.Count + " 枚同时出现在按值带权形参位表里（借来的出口不是接收者）: " +
                 (($overlap | Select-Object -First 4) -join ', '))
    }
}
# ⚠ 「借用族不许进提名面」那一格**故意不写在这里**：借用族是**出口**，今天这张表里根本没有它们的行
#     ⇒ 那一格恒为空 = 一条**永远不会红的判据**（负控一跑就露馅：把借用族的 gate 打开，照样绿）。
#     它属于**回收点**那一侧（`ConcatFree` 放掉左操作数 / `VariantClear` 清一枚变体 ⇒ 只在
#     「那一枚归调用方所有」时才成立），由回收刀自己的哨兵钉；本表只把借用名单钉住（K11①）
#     与不相交性（K11②）。

# K12 **发码侧那份**（src/common/owned_temp_keeps.hpp）与本表**双向对齐**：
#     例外行（KEEPS=1 / U）、closed 那一族、借用族，三样都要**逐行相同** ——
#     表是**推导**出来的（从 RTL 的权威头），头是**给发码侧问的那一份**；
#     少一行（表里有、头里没有 ⇒ 发码侧会提名一枚不许提名的）都红。
$hdrFile = Join-Path $Root 'src\common\owned_temp_keeps.hpp'
if (-not (Test-Path -LiteralPath $hdrFile)) {
    $bad += 'K12 src/common/owned_temp_keeps.hpp 不在 —— 发码侧那份权威没了 ⇒ 停机'
} else {
    $hdrTxt = [System.IO.File]::ReadAllText($hdrFile)
    # 头里的例外行
    $hdrRows = @{}
    $em = [regex]::Match($hdrTxt, 'kRows\[\]\s*=\s*\{([\s\S]*?)\n    \};')
    if (-not $em.Success) {
        $bad += 'K12 owned_temp_keeps.hpp 里的 kRows 读不出来'
    } else {
        foreach ($rm in [regex]::Matches($em.Groups[1].Value, '\{\s*"(vb6_[A-Za-z0-9_]+)"\s*,\s*(\d+)\s*,\s*''(.)''\s*,\s*"([^"]*)"\s*\}')) {
            $hdrRows[$rm.Groups[1].Value + '|' + $rm.Groups[2].Value] = @($rm.Groups[3].Value, $rm.Groups[4].Value)
        }
        # 表里的例外行（KEEPS=1 / U）
        $tblRows = @{}
        foreach ($k in $rows.Keys) {
            if ($rows[$k][0] -ne '0') { $tblRows[$k] = @($rows[$k][0], $rows[$k][1]) }
        }
        $missH = @($tblRows.Keys | Where-Object { -not $hdrRows.ContainsKey($_) })
        $missT = @($hdrRows.Keys | Where-Object { -not $tblRows.ContainsKey($_) })
        if ($missH.Count -gt 0) {
            $bad += ("K12 表里的例外行 " + $missH.Count + " 枚在发码侧那份里没有（发码侧会提名一枚不许提名的）: " +
                     (($missH | Select-Object -First 5) -join ', '))
        }
        if ($missT.Count -gt 0) {
            $bad += ("K12 发码侧那份里多出 " + $missT.Count + " 枚例外行（表里没有 ⇒ 那是发码侧自己加的）: " +
                     (($missT | Select-Object -First 5) -join ', '))
        }
        foreach ($k in $tblRows.Keys) {
            if ($hdrRows.ContainsKey($k)) {
                if ($hdrRows[$k][0] -ne $tblRows[$k][0]) {
                    $bad += ("K12 " + $k + " 的分类在两处不同：表 " + $tblRows[$k][0] + " / 头 " + $hdrRows[$k][0])
                } elseif ($hdrRows[$k][1] -ne $tblRows[$k][1]) {
                    $bad += ("K12 " + $k + " 的逃逸通道在两处不同：表 " + $tblRows[$k][1] + " / 头 " + $hdrRows[$k][1])
                }
            }
        }
    }
    # closed 那一族
    $cm = [regex]::Match($hdrTxt, 'kClosed\[\]\s*=\s*\{([\s\S]*?)\n    \};')
    if (-not $cm.Success) {
        $bad += 'K12 owned_temp_keeps.hpp 里的 kClosed 读不出来'
    } else {
        $hdrClosed = @([regex]::Matches($cm.Groups[1].Value, '"(vb6_[A-Za-z0-9_]+)"') | ForEach-Object { $_.Groups[1].Value })
        $tblClosed = @($closedRows | ForEach-Object { ($_ -split '\|')[0] } | Sort-Object -Unique)
        $d1 = @($tblClosed | Where-Object { $hdrClosed -notcontains $_ })
        $d2 = @($hdrClosed | Where-Object { $tblClosed -notcontains $_ })
        if ($d1.Count -gt 0) { $bad += ("K12 closed 那一族：表里有 " + ($d1 -join ', ') + "，发码侧那份里没有") }
        if ($d2.Count -gt 0) { $bad += ("K12 closed 那一族：发码侧那份里多了 " + ($d2 -join ', ') + "，表里没有") }
    }
    # 借用族：发码侧那份 ≡ 编译器侧 kBorrowedBstrExports167 ≡ 本表的 [borrow] 节
    $bm = [regex]::Match($hdrTxt, 'kBorrowed\[\]\s*=\s*\{([\s\S]*?)\n    \};')
    if (-not $bm.Success) {
        $bad += 'K12 owned_temp_keeps.hpp 里的 kBorrowed 读不出来'
    } else {
        $hdrBorrow = @([regex]::Matches($bm.Groups[1].Value, '"(vb6_[A-Za-z0-9_]+)"') | ForEach-Object { $_.Groups[1].Value })
        $d1 = @($borrowedExports | Where-Object { $hdrBorrow -notcontains $_ })
        $d2 = @($hdrBorrow | Where-Object { $borrowedExports -notcontains $_ })
        if ($d1.Count -gt 0) { $bad += ("K12 借用族：编译器侧名单里有 " + ($d1 -join ', ') + "，发码侧那份里没有") }
        if ($d2.Count -gt 0) { $bad += ("K12 借用族：发码侧那份里多了 " + ($d2 -join ', ') + "，编译器侧名单里没有") }
    }
    # K13 **返回侧所有权**：发码侧那份的 owned 名单必须与本表推导出来的 owned 名单**逐行相同**（双向）
    #     —— 少一行 = 发码侧会放掉一枚它证明是借来的/答不出的值；多一行 = 那是发码侧自己加的。
    $om = [regex]::Match($hdrTxt, '(?s)GENERATED-OWNED-EXITS <<<(.*?)// >>> END-GENERATED-OWNED-EXITS')
    if (-not $om.Success) {
        $bad += 'K13 owned_temp_keeps.hpp 里的 owned 名单读不出来（标记不在？）'
    } else {
        $hdrOwned = @([regex]::Matches($om.Groups[1].Value, '"(vb6_[A-Za-z0-9_]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $d1 = @($ownOwned | Where-Object { $hdrOwned -notcontains $_ })
        $d2 = @($hdrOwned | Where-Object { $ownOwned -notcontains $_ })
        if ($d1.Count -gt 0) {
            $bad += ("K13 推导出来的 owned " + $d1.Count + " 枚发码侧那份里没有（发码侧会放掉一枚它证明不是新造的值）: " +
                     (($d1 | Select-Object -First 5) -join ', '))
        }
        if ($d2.Count -gt 0) {
            $bad += ("K13 发码侧那份 owned 多出 " + $d2.Count + " 枚（推导里没有 ⇒ 那是发码侧自己加的）: " +
                     (($d2 | Select-Object -First 5) -join ', '))
        }
    }
    # K13c **owned 名单里那些 VARIANT 档**（其余 owned 交出 BSTR）—— 档号**只由出口的返回档定**
    #     （源表达式的 VB 档不作数：`Len(v)` 的源档是 Variant、交出去的是 BSTR）。这张名单是
    #     owned 那张的**子集** ⇒ 少一枚/多一枚/不是子集，三样都红。
    $ownVariant = @($ownOwned | Where-Object { $retKind[$_] -eq 'VARIANT' })
    $vm = [regex]::Match($hdrTxt, 'kOwnedVariant\[\]\s*=\s*\{([\s\S]*?)\n    \};')
    if (-not $vm.Success) {
        $bad += 'K13c owned_temp_keeps.hpp 里的 kOwnedVariant 读不出来'
    } else {
        $hdrVar = @([regex]::Matches($vm.Groups[1].Value, '"(vb6_[A-Za-z0-9_]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $d1 = @($ownVariant | Where-Object { $hdrVar -notcontains $_ })
        $d2 = @($hdrVar | Where-Object { $ownVariant -notcontains $_ })
        if ($d1.Count -gt 0) {
            $bad += ("K13c VARIANT 档的 owned " + $d1.Count + " 枚发码侧那份里没有: " + (($d1 | Select-Object -First 5) -join ', '))
        }
        if ($d2.Count -gt 0) {
            $bad += ("K13c 发码侧那份的 VARIANT 档 owned 多出 " + $d2.Count + " 枚（推导里不是 VARIANT 档）: " + (($d2 | Select-Object -First 5) -join ', '))
        }
    }
    # K13d **消费轴**（账 #463 §B164 研究新增）：发码侧那份的 consumes 名单必须与本表推导出来的
    #     **逐格相同**（双向）—— 少一格 = 发码侧会把「消费者自己放掉」的那一枚**再放一次**（双放）。
    $cm2 = [regex]::Match($hdrTxt, 'kConsumes\[\]\s*=\s*\{([\s\S]*?)\n    \};')
    if (-not $cm2.Success) {
        $bad += 'K13d owned_temp_keeps.hpp 里的 kConsumes 读不出来'
    } else {
        $hdrConsumes = @{}
        foreach ($xm in [regex]::Matches($cm2.Groups[1].Value, '\{\s*"(vb6_[A-Za-z0-9_]+)"\s*,\s*(\d+)\s*,\s*''(.)''\s*\}')) {
            $hdrConsumes[$xm.Groups[1].Value + '|' + $xm.Groups[2].Value] = $xm.Groups[3].Value
        }
        $d1 = @($consumesRows.Keys | Where-Object { -not $hdrConsumes.ContainsKey($_) })
        $d2 = @($hdrConsumes.Keys | Where-Object { -not $consumesRows.ContainsKey($_) })
        if ($d1.Count -gt 0) {
            $bad += ("K13d 推导出来的消费格 " + $d1.Count + " 格发码侧那份里没有（发码侧会双放）: " +
                     (($d1 | Select-Object -First 6) -join ', '))
        }
        if ($d2.Count -gt 0) {
            $bad += ("K13d 发码侧那份的消费格多出 " + $d2.Count + " 格（推导里没有）: " +
                     (($d2 | Select-Object -First 6) -join ', '))
        }
        foreach ($k in $consumesRows.Keys) {
            if ($hdrConsumes.ContainsKey($k) -and $hdrConsumes[$k] -ne $consumesRows[$k][0]) {
                $bad += ("K13d " + $k + " 的条件性在两处不同：表 " + $consumesRows[$k][0] + " / 头 " + $hdrConsumes[$k])
            }
        }
    }
    # K13b 每张表末尾必须有**空哨兵**（遍历写法是 `for (...; r->exitName; ++r)`）——
    #     少了它就**走出数组**（本刀实测：编得过、跑起来 0xC0000374 堆损坏 ⇒ 正是这一类）。
    #     这条是**文本**判据，但它挡的正是「表编得过、跑起来炸」那一族。
    foreach ($arr in @(@{ name = 'kRows';      pat = 'static const OwnedTempException kRows\[\] = \{' },
                       @{ name = 'kClosed';   pat = 'static const char\* const kClosed\[\] = \{' },
                       @{ name = 'kBorrowed'; pat = 'static const char\* const kBorrowed\[\] = \{' },
                       @{ name = 'kOwned';    pat = 'static const char\* const kOwned\[\] = \{' },
                       @{ name = 'kOwnedVariant'; pat = 'static const char\* const kOwnedVariant\[\] = \{' },
                       @{ name = 'kConsumes'; pat = 'static const OwnedTempConsumes kConsumes\[\] = \{' })) {
        $am2 = [regex]::Match($hdrTxt, '(?s)' + $arr.pat + '(.*?)\n    \};')
        if (-not $am2.Success) {
            $bad += ("K13b " + $arr.name + " 读不出来（头被改坏了？）")
            continue
        }
        $body2 = $am2.Groups[1].Value
        $tail2 = (($body2 -split "`r?`n") | Where-Object { $_.Trim() -ne '' } | Select-Object -Last 1)
        if ($tail2 -notmatch 'nullptr' -and $tail2 -notmatch '^\s*0\s*,') {
            $bad += ("K13b " + $arr.name + " 末尾没有空哨兵（遍历会走出数组 ⇒ 0xC0000374 那一族）")
        }
    }
}

# K17 **下一刀的预测站点集合**（账 #458 §B228 交付的第二件事）：把「哪些形状会被提名」
#     列成一张表，逐条**问那位主人**（头里的 ownedTempNominationVerdict），再证明
#     **预测集合与所有 gate=closed 行交集为空** —— 这一条是机械的，不是断言。
#     形状表里的三行是**陷阱**（借来的 / 别名 / closed 消费者），必须**不被提名**；
#     另有一行是「顶层不是调用」（裸字面量）⇒ 连 mint 都不是。
$shapes = @(
    @{ id = 'concat-arg';   consumer = 'vb6_Len';        slot = 0; mint = 'vb6_BSTR_Concat';    want = 'B'; why = '拼接结果进 Len：现造的 BSTR + Len 不留指针' },
    @{ id = 'concat-arg2';  consumer = 'vb6_Len';        slot = 0; mint = 'vb6_BSTR_Concat';    want = 'B'; why = '同上（数组元素那一形）' },
    @{ id = 'varianttoStr'; consumer = 'vb6_Len';        slot = 0; mint = 'vb6_VariantToString'; want = 'B'; why = 'Len(v)：交出去的是 BSTR 档（名字带 Variant 不算数）' },
    @{ id = 'variant-arg';  consumer = 'vb6_VarType';     slot = 0; mint = 'vb6_VariantFromValue'; want = '0'; why = '那一枚是 _Generic 宏（不在权威出口面上）⇒ 答不出 ⇒ 不提名' },
    @{ id = 'TRAP-cdec';    consumer = 'vb6_VarAddV';     slot = 0; mint = 'vb6_CDec';             want = '0'; why = 'CDec 的 DECIMAL 那一支交回**形参本身** ⇒ 别名（账 #463 补上的那一形）⇒ 不许放' },
    @{ id = 'TRAP-iifbstr'; consumer = 'vb6_IIfBSTR';     slot = 1; mint = 'vb6_BSTR_FromStr';     want = '0'; why = 'IIf 那一族**条件性**地放掉没选中的那一支 ⇒ 调用方无从知道 ⇒ 不许提名' },
    @{ id = 'TRAP-consumes'; consumer = 'vb6_BSTR_ConcatFree'; slot = 0; mint = 'vb6_BSTR_FromStr'; want = '0'; why = 'ConcatFree 放掉**左**形参 ⇒ 调用方再放就是双放 ⇒ 不许提名' },
    @{ id = 'TRAP-borrowed'; consumer = 'vb6_Len';       slot = 0; mint = 'vb6_ErrDescription'; want = '0'; why = '借来的那一枚（全局 vb6_err 的指针）⇒ 不许放' },
    @{ id = 'TRAP-alias';   consumer = 'vb6_Len';        slot = 0; mint = 'vb6_VariantArrayGetND'; want = '0'; why = '数组元素读交出去的是别名 ⇒ 不许放' },
    @{ id = 'TRAP-closed'; consumer = 'vb6_SaveSetting'; slot = 0; mint = 'vb6_BSTR_FromStr'; want = '0'; why = 'closed 那一族（未归因 ⇒ 不许提名）' },
    @{ id = 'TRAP-comfour'; consumer = 'vb6_Len';        slot = 0; mint = 'vb6_ComCall';       want = '0'; why = 'COM 四兄弟交的是 void*，不是带权值' },
    @{ id = 'TRAP-notmint'; consumer = 'vb6_Len';        slot = 0; mint = '';                    want = '0'; why = '顶层不是调用（裸字面量）⇒ 连 mint 都不是' }
)
$predicted = @()
foreach ($sh in $shapes) {
    $v = 0
    if ($sh.mint -ne '') {
        # 头里那位主人的判定在 C++ 侧；这里**按同一条规则**在 PowerShell 侧重演一遍
        # （规则只有三道闸，逐条对应头里的三个问句 ⇒ 两侧同源；K17b 用真编的探针校准过）
        $kind = 0
        if ($ownOwned -contains $sh.mint) { $kind = 'B' }
        if ($ownVariant -contains $sh.mint) { $kind = 'V' }
        if ($kind -eq '0') { $v = 0 }
        elseif ($closedExits -contains $sh.consumer) { $v = 0 }
        else {
            $keeps = '0'
            foreach ($ex in $ownedTempExceptions) {
                if ($ex.name -eq $sh.consumer -and $ex.slot -eq $sh.slot) { $keeps = $ex.keeps }
            }
            # ④ 消费轴（账 #463）：消费者自己放掉这一格 ⇒ 不许再放（双放）
            $consumes = $consumesRows.ContainsKey($sh.consumer + '|' + $sh.slot)
            if ($keeps -ne '0' -or $consumes) { $v = 0 } else { $v = $kind }
        }
    }
    if ($v -ne $sh.want) {
        $bad += ("K17 形状 " + $sh.id + " 的判定是 '" + $v + "'，期望 '" + $sh.want + "'（" + $sh.why + "）")
    }
    if ($v -ne '0') { $predicted += $sh }
}
# 预测集合 ∩ closed = ∅（机械证明：逐条查预测集合里的消费者）
foreach ($p in $predicted) {
    if ($closedExits -contains $p.consumer) {
        $bad += ("K17 预测集合里混进了 closed 消费者 " + $p.consumer + "（" + $p.id + "）")
    }
}
# 另：预测集合里**不许**出现借用族 / 别名族（那两族的出口名在 owned 名单里就不该有）
foreach ($p in $predicted) {
    if ($aliasExits -contains $p.mint) {
        $bad += ("K17 预测集合里混进了别名族出口 " + $p.mint + "（" + $p.id + "）")
    }
}

# K2: unclassified == 0 —— 出口面里每一枚按值带权形参位都必须有行
$missing = @($surface.Keys | Where-Object { -not $rows.ContainsKey($_) })
if ($missing.Count -gt 0) {
    $bad += ("K2 出口面里有 " + $missing.Count + " 枚带权形参位在表里没有行（unclassified ≠ 0）: " +
             (($missing | Select-Object -First 6) -join ', '))
}
# K3: 表里不许有**出口面之外**的行（表比权威面大 = 表自己造了出口）
$extra = @($rows.Keys | Where-Object { -not $surface.ContainsKey($_) })
if ($extra.Count -gt 0) {
    $bad += ("K3 表里有 " + $extra.Count + " 行不在权威出口面里（表自己造了出口）: " +
             (($extra | Select-Object -First 6) -join ', '))
}
# K4: U 那一类必须**显式存在且非空**（不可判定不许被静默归 0）
if ($ku -eq 0) {
    $bad += 'K4 U（不可判定）那一类是空的 —— 答不出的出口必须显式归 U，不许静默变成 KEEPS=0'
}
# K5: 快照 == 现算（漂移即红）
if ($snap.Count -eq 0) {
    $bad += 'K5 快照 ' + $TableFile + ' 不存在或读不出 —— 先跑一次 -Bless 登记'
} else {
    $drift = @()
    foreach ($k in $rows.Keys) {
        if (-not $snap.ContainsKey($k)) { $drift += ('+' + $k) } else {
            $s = $snap[$k]; $v = $rows[$k]
            if ($s[0] -ne $v[0] -or $s[1] -ne $v[1] -or $s[2] -ne $v[2]) { $drift += ('~' + $k) }
        }
    }
    foreach ($k in $snap.Keys) { if (-not $rows.ContainsKey($k)) { $drift += ('-' + $k) } }
    if ($drift.Count -gt 0) {
        $bad += ("K5 表与快照漂移 " + $drift.Count + " 行（先归因，再 -Bless）: " +
                 (($drift | Select-Object -First 8) -join ', '))
    }
}
# K6: 灰控两条规则必须在案（用户过程 / 晚绑定 COM 归 U）
if ($grayRules.Count -ne 2) { $bad += 'K6 灰控规则条数不对（用户过程 + 晚绑定 COM，两条都要在案）' }
# K7: 负证人登记 —— 撤回那一版弄坏过 5 份既有夹具，它们从「受害者」升级成**预先登记的证人**
$witnesses = @(
    'tests\test_registry_settings.bas',
    'tests\test_bstr_borrow_slot.bas',
    'tests\test_host_decimal.bas'
)
$missingW = @($witnesses | Where-Object { -not (Test-Path -LiteralPath (Join-Path $Root $_)) })
if ($missingW.Count -gt 0) {
    $bad += ("K7 负证人不存在: " + ($missingW -join ', '))
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS owned temp keeps: authority headers " + $authorityHeaders.Count +
                " / surface " + $surface.Count + " / rows " + $rows.Count +
                " | keeps_pointer 0 " + $k0 + " 1 " + $k1 + " U " + $ku +
                " | closed " + $closedRows.Count + " | nomination " + $nomRows.Count +
                " | unclassified 0 | 15 anchors OK (K14 + K14b) | K13/K13b/K13c two-way | gray rules 2 | witnesses 3")
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
