# check_com_sig_field_defaults.ps1 - 账 §B73/§B94 的结构哨兵 (第 44 道；钉「签名载体的字段必须带默认初值」)
#
# 为什么要这一道：`ComMemberInfo::returnType` 曾经没有初值，而 parseVarDesc 只在 VAR_PERINSTANCE
# 那一支给它赋值 ⇒ dual 接口的属性（VARDESC kind=VAR_PROPERTY，如 stdole.StdFont.Name）交出去的是一枚
# **没写过的 16 位字段**。实测同一枚二进制连跑三次读到 29620/6971/50156，换一台工具链编出来的二进制就
# 稳定落进另一个档 ⇒ 「同一笔提交，CI 与本机发码分家」（泛型 `vb6_ComGetProp` ⇄ 带类型的 `vb6_ComGetIntProp`，
# 而 Name 是 BSTR ⇒ 后者压根是错的答案）。那条老账（§B73）排查时把嫌疑一路收到「只剩 C3.exe 自己的工具链」，
# 真凶其实就是这枚未初始化字段 —— 账 #245 给它补了 `= Vb6Type::Variant`（=「不知道」档）之后，
# 两台工具链 × 全语料的 --emit-c 才第一次逐字节对上。
#
# 这一道钉的不是「值对不对」而是「有没有答案」：这类结构体是**跨阶段递信息的载体**，
# 少一个初值 = 编译器按寄存器/栈残值答题，而这种缺陷在任何一台机器上都编得过、跑得动。
#
# 规则 (改坏了会红):
#   F1 名单里每枚 struct 都必须在指定文件里**找得到**（改名/搬走 = 哨兵空转）。
#   F2 每枚**值类型或裸指针**字段必须带默认初值（std::string / std::vector / std::unique_ptr 这些自己会初始化，跳过）。
#   F3 检查到的字段数不许退化（>=15，今天 19），且 `ComMemberInfo` 与 `ComMethodSig` 里必须都出现 `returnType`
#      —— 那枚元凶字段被摘掉或改名，也要红。
#
# 用法: powershell -File scripts/check_com_sig_field_defaults.ps1 [-Root <检出根>]   (PASS => exit 0)

param([string]$Root = "")
$ErrorActionPreference = 'Stop'
$root = if ($Root) { (Resolve-Path $Root).Path } else { Split-Path -Parent $PSScriptRoot }

$specs = @(
    [pscustomobject]@{ rel = 'src\com\typelib_parser.hpp'; pin = 'returnType';
                       structs = @('ComParamInfo', 'ComMemberInfo', 'ComInterfaceInfo', 'ComCoClassInfo',
                                   'ComEnumInfo', 'ComEnumMemberInfo', 'TypeLibResult') },
    [pscustomobject]@{ rel = 'src\semantics\symbol_table.hpp'; pin = 'returnType';
                       structs = @('ComMethodSig') }
)
$fails = @()
$checked = 0
$valueTypeRx = '^(bool|char|wchar_t|short|int|long|float|double|unsigned|signed|size_t|ptrdiff_t|uint\d+_t|int\d+_t|intptr_t|uintptr_t)$'
$userTypeRx  = '^[A-Z][A-Za-z0-9_]*$'

foreach ($spec in $specs) {
    $path = Join-Path $root $spec.rel
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Host ('FAIL missing ' + $spec.rel); exit 1
    }
    $text = [System.IO.File]::ReadAllText($path)
    foreach ($nm in $spec.structs) {
        # 必须匹配到 `struct <名字> {` 那一行 —— 只 IndexOf('struct ' + $nm) 会被**前缀**骗过去
        # （把 ComMethodSig 改名成 ComMethodSigRenamed，子串照样命中 ⇒ F1 静默绿，实测栽过一次）
        $rxDef = [regex]('struct\s+' + [regex]::Escape($nm) + '\s*\{')
        $m0 = $rxDef.Match($text)
        if (-not $m0.Success) { $fails += ('F1 ' + $nm + ' not found in ' + $spec.rel); continue }
        $j = $m0.Index + $m0.Length - 1
        # 体尾配到**第一个顶格的 `};`**（不许定长窗口 —— 别人的字段一多就把固定长度读穿）
        $m = [regex]::Match($text.Substring($j), "(?m)^\s*\};")
        if (-not $m.Success) { $fails += ('F1 ' + $nm + ' body never closes with };'); continue }
        $bodyLines = @($text.Substring($j + 1, $m.Index) -split "\r?\n")
        $sawPin = $false
        foreach ($raw in $bodyLines) {
            $s = $raw.Trim()
            if (-not $s) { continue }
            if ($s.StartsWith('//') -or $s.StartsWith('/*') -or $s.StartsWith('*')) { continue }
            $s = ($s -replace '//.*$', '').Trim()
            if (-not $s.EndsWith(';')) { continue }
            $decl = $s.Substring(0, $s.Length - 1).Trim()
            if ($decl.StartsWith('return')) { continue }
            $lhs = $decl
            $hasInit = $false
            $eq = $decl.IndexOf('=')
            if ($eq -gt 0) { $lhs = $decl.Substring(0, $eq).Trim(); $hasInit = $true }
            if ($lhs.Contains('(')) { continue }          # 内联函数/工厂，不是字段
            $parts = @($lhs -split '\s+', 2)
            if ($parts.Count -lt 2) { continue }
            $type = $parts[0]
            $name = $parts[1].Trim()
            if ($type.StartsWith('std::')) { continue }
            if ($type -in @('using', 'typedef', 'enum', 'static', 'constexpr', 'friend', 'public', 'private')) { continue }
            $isValue = ($type -match $valueTypeRx) -or ($type -match $userTypeRx)
            $isPtr = $lhs.Contains('*')
            if (-not $isValue -and -not $isPtr) { continue }
            $checked++
            if ($name -eq $spec.pin) { $sawPin = $true }
            if (-not $hasInit) {
                $fails += ('F2 ' + $nm + '.' + $name + ' (type ' + $type + ') 没有默认初值 —— ' +
                           '这类字段会被下游当「这成员是什么型/哪一档」的答案读；没有初值就是拿栈上残值答题')
            }
        }
        if ($nm -eq 'ComMemberInfo' -or $nm -eq 'ComMethodSig') {
            if (-not $sawPin) { $fails += ('F3 ' + $nm + ' 里没有 ' + $spec.pin + ' 字段（被摘掉或改名 = 本哨兵的靶子消失）') }
        }
    }
}

if ($checked -lt 15) {
    $fails += ('F3 只检查到 ' + $checked + ' 枚值类型字段（地板 15）—— 名单里的 struct 空了或解析没吃到字段')
}
if ($fails.Count -gt 0) {
    foreach ($f in $fails) { Write-Host ('FAIL ' + $f) }
    exit 1
}
Write-Host ('OK check_com_sig_field_defaults: F1..F3 (' + $checked + ' value/pointer fields all default-initialized ' +
            'across 8 signature-carrier structs; returnType pinned in both ComMemberInfo and ComMethodSig)')
exit 0
