# check_com_prop_type_authority.ps1 - 账 #245 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 这一格的缺陷形状是"答案的来源没被守住"。类型库里的属性是 VARDESC,
# parseVarDesc 以前只在 VAR_PERINSTANCE 那一支给 member.returnType 赋值, 而 dual 接口的
# 属性 (stdole 的 StdFont.Name/Size/Bold/Charset ...) 是 VAR_PROPERTY(3) ⇒ 交出去的是
# **没赋过值的 16 位字段**。实测同一枚二进制连跑三次读到 29620 / 6971 / 50156, 换一档
# MSVC 编出来的二进制就稳定落在另一个值上 ⇒ 下游 resolveComValue 把这份垃圾当
# "这枚 COM 属性是什么型"的答案用, 于是 BSTR 属性被发成 vb6_ComGetIntProp, 而 CI 与本机
# 对同一笔提交交出两种发码 (门 #410 的 emit-samples 对差实测 13 行)。编译不响、链接不响、
# 语料 A/B 也只响一半, 所以要结构钉。
#
# 规则 (改坏了会红, 不是装饰):
#   C1  VARDESC 的类型无条件取: parseVarDesc 体内必须有 `member.returnType = mapTypeDesc(&pVD->elemdescVar.tdesc, pTI);`,
#       且该函数里**不许**再用 `if (pVD->varkind == VAR_PERINSTANCE)` 把它包回去 (那正是本次的现场)。
#   C2  交出去之前就有初值: typelib_parser.hpp 里 ComParamInfo / ComMemberInfo 两个 struct 的每个
#       **标量数据成员**声明行都必须带 `=` (std::string / std::vector 豁免)。没有初值的枚举字段
#       一旦被复制进 Symbol::ComMethodSig, 就是"每进程自洽、跨构建不同"的那一格。
#   C3  口径住在唯一出口里: comGetterExpr (属性 getter 的档位选择) 必须
#       ①先看 sv.isGet (put/method 那份签名不许当属性读来解包),
#       ②类型未知时退回 `vb6_VariantFromComResult(vb6_ComGetProp(`，且**整个函数体里不许出现
#       unpackType / packFnHint** —— 那两枚是「按调用上下文猜一档」的旧答案来源，同族另一处
#       (cgen_expr_call_arg_emit.inc) 本来就是退 VARIANT。读表那处 comSigViewOf 必须仍交出 isPropertyGet。
#   C4  census: `comMethods.find(` 在 src\backend 里只许出现一次，且就在 cgen_util_com.cpp 的
#       comSigViewOf 里；以前那五处消费者（三份 .inc + cgen_util_com.cpp 的另一支）必须仍逐处
#       问 comSigViewOf —— 名单是「谁在问」，不是「谁自己查表」。名单外再开一处读法必须先回来把口径并进来再登记。
#
# 用法: powershell -File scripts/check_com_prop_type_authority.ps1   (PASS ⇒ exit 0)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

$descPath = Join-Path $root 'src\com\typelib_parser_desc.cpp'
$hdrPath  = Join-Path $root 'src\com\typelib_parser.hpp'
$comPath  = Join-Path $root 'src\backend\cgen_util_com.cpp'

foreach ($p in @($descPath, $hdrPath, $comPath)) {
    if (-not (Test-Path $p)) { Write-Host "FAIL missing file: $p"; exit 1 }
}

$desc = Get-Content -Raw -Encoding UTF8 $descPath
$hdr  = Get-Content -Raw -Encoding UTF8 $hdrPath
$com  = Get-Content -Raw -Encoding UTF8 $comPath

$fails = @()

# ---- C1: parseVarDesc 无条件给 returnType 赋值 ----
$vIdx = $desc.IndexOf('ComMemberInfo TypeLibParser::parseVarDesc')
if ($vIdx -lt 0) {
    $fails += 'C1 parseVarDesc not found in typelib_parser_desc.cpp'
} else {
    $nIdx = $desc.IndexOf('ComMemberInfo TypeLibParser::parseFuncDesc', $vIdx)
    $endIdx = if ($nIdx -gt 0) { $nIdx } else { $desc.Length }
    $body = $desc.Substring($vIdx, $endIdx - $vIdx)
    if ($body -notmatch 'member\.returnType\s*=\s*mapTypeDesc\(\s*&pVD->elemdescVar\.tdesc') {
        $fails += 'C1 parseVarDesc no longer assigns member.returnType from elemdescVar.tdesc'
    }
    if ($body -match 'if\s*\(\s*pVD->varkind\s*==\s*VAR_PERINSTANCE\s*\)') {
        $fails += 'C1 parseVarDesc guards returnType on VAR_PERINSTANCE again (VAR_PROPERTY/VAR_CONST would emit an uninitialized field)'
    }
}

# ---- C2: 两个 struct 的标量成员都要有初值 ----
foreach ($st in @('ComParamInfo', 'ComMemberInfo')) {
    $sIdx = $hdr.IndexOf("struct $st {")
    if ($sIdx -lt 0) { $fails += "C2 struct $st not found"; continue }
    $eIdx = $hdr.IndexOf('};', $sIdx)
    $blk = $hdr.Substring($sIdx, $eIdx - $sIdx)
    foreach ($line in ($blk -split "`r?`n")) {
        $t = ($line -replace '//.*$', '').Trim()      # 先脱行尾注释: 声明行都带中文注释, 不脱就永远看不见它们
        if ($t -eq '' -or -not $t.EndsWith(';')) { continue }
        if ($t -match '^std::(string|vector)') { continue }
        if ($t -notmatch '=') {
            $fails += "C2 struct $st has an uninitialized scalar member: $t"
        }
    }
}

# ---- C3: 属性 getter 的档位口径（账 #245/§B94 之后住在唯一出口里）----
$geIdx = $com.IndexOf('std::string CCodeGen::comGetterExpr(')
if ($geIdx -lt 0) {
    $fails += 'C3 comGetterExpr (the single typed-getter exit) is gone from cgen_util_com.cpp'
} else {
    $geTail = $com.IndexOf("`n}", $geIdx)
    $geBlk = if ($geTail -gt 0) { $com.Substring($geIdx, $geTail - $geIdx) } else { $com.Substring($geIdx) }
    # 判据前先剥 //… 与 /*…*/: 函数上方的说明里就写着 "unpackType"，不剥会把一句说明读成一次猜档。
    $geCode = ($geBlk -replace '//[^\r\n]*', '') -replace '/\*[\s\S]*?\*/', ''
    if ($geCode -notmatch 'vb6_VariantFromComResult\(vb6_ComGetProp\(') {
        $fails += 'C3 unknown return type no longer falls back to the generic VARIANT read'
    }
    if ($geCode -notmatch '!sv\.isGet') {
        $fails += 'C3 the getter exit stopped gating on the confirmed property-getter flag (a put/method signature must not be unpacked as a property read)'
    }
    if ($geCode -match 'unpackType|packFnHint') {
        $fails += 'C3 the getter exit guesses a shape from the call context again (that is how BSTR props became ComGetIntProp)'
    }
}
$svIdx = $com.IndexOf('CCodeGen::ComSigView CCodeGen::comSigViewOf(')
if ($svIdx -lt 0) {
    $fails += 'C3 comSigViewOf (the only reader of Symbol::comMethods) is gone'
} else {
    $svTail = $com.IndexOf("`n}", $svIdx)
    $svBlk = $com.Substring($svIdx, $svTail - $svIdx)
    if ($svBlk -notmatch 'isGet = it->second\.isPropertyGet') {
        $fails += 'C3 comSigViewOf no longer hands out the isPropertyGet answer'
    }
}
$ebIdx = $com.IndexOf('if (isEarlyBoundCom_ && earlyBoundSym_)')
if ($ebIdx -lt 0) {
    $fails += 'C3 resolveComValue early-bound branch not found'
} else {
    $ebTail = $com.IndexOf("`n    }", $ebIdx)
    $ebBlk = if ($ebTail -gt 0) { $com.Substring($ebIdx, $ebTail - $ebIdx) } else { $com.Substring($ebIdx) }
    if ($ebBlk -notmatch 'comGetterExpr\(comSigViewOf\(') {
        $fails += 'C3 resolveComValue no longer routes its early-bound read through the single exit'
    }
}

# ---- C4: census —— 读这张表的地方只剩一处，问它的地方一处不少 ----
$backendDir = Join-Path $root 'src\backend'
$readers = @((Get-ChildItem -Path $backendDir -Recurse -File -Include *.cpp,*.inc,*.hpp |
        Where-Object { (Get-Content -Raw -Encoding UTF8 $_.FullName) -match 'comMethods\.find\(' } |
        ForEach-Object { $_.Name }))
$authority = 'cgen_util_com.cpp'
if ($readers.Count -ne 1 -or $readers[0] -ne $authority) {
    $fails += ('C4 Symbol::comMethods must be read only in ' + $authority +
               ' (comSigViewOf); found readers: ' + ($readers -join ', '))
}
$findCount = ([regex]::Matches($com, 'comMethods\.find\(')).Count
if ($findCount -ne 1) {
    $fails += ('C4 comMethods.find( appears ' + $findCount +
               ' times in ' + $authority + '; exactly the one inside comSigViewOf is allowed')
}
$askers = @{
    'cgen_util_com.cpp'               = $comPath
    'cgen_expr_call_arg_emit.inc'    = (Join-Path $backendDir 'detail\expr\cgen_expr_call_arg_emit.inc')
    'cgen_expr_call_com_bind.inc'    = (Join-Path $backendDir 'detail\expr\cgen_expr_call_com_bind.inc')
    'cgen_expr_call_prelude.inc'     = (Join-Path $backendDir 'detail\expr\cgen_expr_call_prelude.inc')
}
foreach ($k in ($askers.Keys | Sort-Object)) {
    $fp = $askers[$k]
    if (-not (Test-Path $fp)) { $fails += ('C4 missing former consumer file: ' + $k); continue }
    if ((Get-Content -Raw -Encoding UTF8 $fp) -notmatch 'comSigViewOf\(') {
        $fails += ('C4 ' + $k + ' no longer asks comSigViewOf (a reader must route through the exit, not vanish)')
    }
}

if ($fails.Count -gt 0) {
    $fails | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host 'OK check_com_prop_type_authority: C1..C4'
exit 0
