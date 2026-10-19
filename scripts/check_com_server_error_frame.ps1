# ============================================================
# check_com_server_error_frame.ps1 — 账 #305 §B185 (第七十七刀) 的结构性哨兵
#
# 为什么单开这一道：这一格的缺陷形状是「服务端进门那趟压根没有出口」，而且它的后果不在
# 服务端里 —— `ComObj_Invoke` 收下 pExcepInfo 却从不碰、`invokeFunc` 返回 void、末尾无条件
# `return S_OK`，于是 VB 体里没人接住的一条运行期错误一路走到 `vb6_RaiseError`，那里没帧
# ⇒ 打一行 stderr 再 `exit(错误号)` ⇒ **死的是调用它的那个进程**。两台实测(改前)：越界写
# Public Integer 字段 RC=6、体里一句 Err.Raise 5 RC=5，连 `cc_dll` 那批夹具都只能靠
# 「探针没走到那一臂」侥幸绿着 ⇒ 发码、链接、语料 A/B 全都看不见这一格，只有真跑看得见。
# 行为判据住在 tests\b185_srv + Test-DispatchInvoke(两台)；这里钉的是**别再退回两遍形状**。
#
# 规则：
#   R1  那一处只许有一枚帧：vb6comserver_obj.c 的 ComObj_Invoke 体内 `setjmp(` 恰好 1 次、
#       `vb6_SaveErrState();` 与 `vb6_RestoreErrState();` 各恰好 1 次，而且 `setjmp` 那一行
#       必须排在 `method->invokeFunc(` 之前(帧要先登记再调进去，反过来就是「抛错时还没人接」)。
#   R2  登记顺序照 cgen 的 On Error GoTo 协议：`setjmp` 之后必须先登记 ptr/set/active 再调，
#       即 `vb6_error_jmp_ptr = &srvJmp` 那一行在 `setjmp(` 之后、`invokeFunc(` 之前。
#   R3  交回形状只住一处：`0x800A0000` 这个数在全 RTL 只许出现 1 次(就在 vb6_ComFailExcep
#       的函数体里)，且 `coll_fail` 体内不许再出现 `scode =`(那枚内置 Collection 的旧主人
#       必须仍是转调)。这一枚出口自己：声明 1 + 定义 1 + 调用点 2(coll_fail 与 ComObj_Invoke)。
#   R4  失败路径不许跳过收尾：ComObj_Invoke 末尾那两行无条件收尾(`if (args)` / `if (coercedArgs)`
#       各 CoTaskMemFree)必须还在，且 `vb6_ComFailExcep(` 那一次调用排在它们之后
#       (错误路径与成功路径走同一条收尾，否则每次抛错漏两批缓冲)。
#   R5  账 #305 §B186 (第七十八刀)：服务端进门的**第二站** = `vb6_ComObject_Create` 里
#       `desc->factoryFunc()` 那一句(它走进 Class_Initialize)。同一道决定不许只补一半：
#       这一格里 `setjmp` / Save / Restore / `factoryFunc` 各恰好 1 次，且登记排在调进去之前。
#   R6  同那一刀的**分档**：这一条通道没有 EXCEPINFO，失败只能编进 HRESULT ——
#       `vb6_ComObject_Create` 里交回 0x800A0000|n 的那一句必须恰好 1 次、E_OUTOFMEMORY 那两格
#       (包装器分不到 / 工厂交回 NULL)必须还在，而 `CF_CreateInstance` 里**不许**再出现
#       `return E_OUTOFMEMORY`(它过去把 NULL 一律说成内存不够，正是这一格的分档缺失)；
#       客户端三条 create 腿一律问 `vb6_ComFailCreate`，那张网里不许再出现裸的 `vb6_RaiseError(429`。
#
# 用法: powershell -File scripts\check_com_server_error_frame.ps1 [-Root <目录>]   (PASS ⇒ exit 0)
# 负控在 src 的副本上跑(-Root)，共享工作树一格不动。

param([string]$Root = "")

$ErrorActionPreference = 'Stop'
if ($Root -eq "") { $Root = Split-Path -Parent $PSScriptRoot }

$objPath  = Join-Path $Root 'src\rtl\core\vb6comserver\vb6comserver_obj.c'
$comPath  = Join-Path $Root 'src\rtl\core\vb6com\vb6com.c'
$colPath  = Join-Path $Root 'src\rtl\core\vb6com\vb6com_collection.c'
$hPath    = Join-Path $Root 'src\rtl\core\vb6com\vb6com_internal.h'
$facPath  = Join-Path $Root 'src\rtl\core\vb6comserver\vb6comserver_factory.c'

$fail = @()

# 同一文件里那条 R3 的尺子只量代码：注释里引那句形状是这一仓的常规做法(§B182 的口径
# 「谁在提这句文本都算」反过来用 —— 不许把注释改哑来迁就尺子)，所以先剥注释再数。
function Strip-Comments185([string]$text) {
    $out = New-Object System.Text.StringBuilder
    $i = 0
    $n = $text.Length
    $inBlock = $false
    $inLine = $false
    $inStr = $false
    while ($i -lt $n) {
        $c = $text[$i]
        $d = if ($i + 1 -lt $n) { $text[$i + 1] } else { [char]0 }
        if ($inLine) { if ($c -eq "`n") { $inLine = $false; [void]$out.Append($c) }; $i++; continue }
        if ($inBlock) { if ($c -eq "*" -and $d -eq "/") { $inBlock = $false; $i += 2; continue }
                        if ($c -eq "`n") { [void]$out.Append($c) }; $i++; continue }
        if ($inStr) { [void]$out.Append($c); if ($c -eq "\") { if ($i + 1 -lt $n) { [void]$out.Append($text[$i+1]) }; $i += 2; continue }
                      if ($c -eq '"') { $inStr = $false }; $i++; continue }
        if ($c -eq "/" -and $d -eq "*") { $inBlock = $true; $i += 2; continue }
        if ($c -eq "/" -and $d -eq "/") { $inLine = $true; $i += 2; continue }
        if ($c -eq '"') { $inStr = $true }
        [void]$out.Append($c)
        $i++
    }
    return $out.ToString()
}

function Read-Lines([string]$p) {
    $raw = [System.IO.File]::ReadAllText($p)
    return ,($raw -split "`r?`n")
}

# 只取函数体：从函数头那一行到下一个顶格 '}' (这仓的 C 文件都用顶格收尾)
function Get-Body([string[]]$lines, [string]$headRx) {
    $s = -1
    for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match $headRx) { $s = $i; break } }
    if ($s -lt 0) { return @() }
    for ($j = $s; $j -lt $lines.Count; $j++) {
        if ($j -gt $s -and $lines[$j] -match '^\}') { return $lines[$s..($j - 1)] }
    }
    return $lines[$s..($lines.Count - 1)]
}

function Count-In([string[]]$body, [string]$needle) {
    $n = 0
    foreach ($l in $body) { if ($l.Contains($needle)) { $n++ } }
    return $n
}
function Index-Of([string[]]$body, [string]$needle) {
    for ($i = 0; $i -lt $body.Count; $i++) { if ($body[$i].Contains($needle)) { return $i } }
    return -1
}

$obj  = Read-Lines $objPath
$invokeBody = Get-Body $obj 'static HRESULT STDMETHODCALLTYPE ComObj_Invoke'
if ($invokeBody.Count -lt 10) { Write-Host 'FAIL R1 cannot locate ComObj_Invoke body'; exit 1 }

# ---- R1 ----
$nSet  = Count-In $invokeBody 'setjmp('
$nSave = Count-In $invokeBody 'vb6_SaveErrState();'
$nRest = Count-In $invokeBody 'vb6_RestoreErrState();'
$nCall = Count-In $invokeBody 'method->invokeFunc('
if ($nSet -ne 1)  { $fail += "R1 setjmp in ComObj_Invoke = $nSet (want exactly 1)" }
if ($nSave -ne 1) { $fail += "R1 vb6_SaveErrState in ComObj_Invoke = $nSave (want exactly 1)" }
if ($nRest -ne 1) { $fail += "R1 vb6_RestoreErrState in ComObj_Invoke = $nRest (want exactly 1)" }
if ($nCall -ne 1) { $fail += "R1 method->invokeFunc in ComObj_Invoke = $nCall (want exactly 1)" }

# ---- R2 ----
$iSet  = Index-Of $invokeBody 'setjmp('
$iPtr  = Index-Of $invokeBody 'vb6_error_jmp_ptr = &srvJmp'
$iCall = Index-Of $invokeBody 'method->invokeFunc('
if ($iSet -lt 0 -or $iPtr -lt 0 -or $iCall -lt 0) {
    $fail += "R2 missing one of setjmp/jmp_ptr-registration/invokeFunc (set=$iSet ptr=$iPtr call=$iCall)" }
elseif (-not ($iSet -lt $iPtr -and $iPtr -lt $iCall)) {
    $fail += "R2 order is setjmp($iSet) -> jmp_ptr($iPtr) -> invokeFunc($iCall); must register the frame before calling in" }

# ---- R3 ----
$allRtl = @(Get-ChildItem -Path (Join-Path $Root 'src\rtl') -Recurse -Include *.c, *.h, *.inc)
$nShape = 0
foreach ($f in $allRtl) {
    $t = Strip-Comments185 ([System.IO.File]::ReadAllText($f.FullName))
    $nShape += ([regex]::Matches($t, '0x800A0000')).Count
}
if ($nShape -ne 1) { $fail += "R3 the 0x800A0000 shape appears $nShape times in src\rtl (must live in exactly one exit)" }

$col    = Read-Lines $colPath
$failBody = Get-Body $col '^static HRESULT coll_fail'
if ((Count-In $failBody 'scode =') -ne 0) {
    $fail += "R3 coll_fail assigns scode again - it must stay a thin shell over vb6_ComFailExcep" }
if ((Count-In $failBody 'vb6_ComFailExcep(') -ne 1) {
    $fail += "R3 coll_fail does not call vb6_ComFailExcep exactly once" }

$com  = Read-Lines $comPath
$defB = Get-Body $com '^HRESULT vb6_ComFailExcep'
if ($defB.Count -lt 5) { $fail += "R3 vb6_ComFailExcep definition not found in vb6com.c" }
# 声明按**签名形状**数，且先剥注释 (账 #305 §B186 的实测: 头文件里那句「造: vb6_ComFailExcep 填
# scode」是给人读的说明，被整行 Contains 数成了第二枚声明 —— 与 R3 那条数 0x800A 的尺子同形，
# 这一族的做法是**把尺子的形状识别改对**，不是把说明文字改哑)。
$hCode = Strip-Comments185 ((Read-Lines $hPath) -join "`n")
$decls = ([regex]::Matches($hCode, 'HRESULT vb6_ComFailExcep\s*\(')).Count
if ($decls -ne 1) { $fail += "R3 declarations of vb6_ComFailExcep in vb6com_internal.h = $decls (want 1)" }
# 调用点 = 真的把那枚出口**用出去**的两处；本文件顶部那句 extern 声明不算(否则这一计数会随
# 跨族声明的写法漂，而漂的不是这道决定)。
$objCode = Strip-Comments185 (($obj -join "`n"))
$colCode = Strip-Comments185 ((Read-Lines $colPath) -join "`n")
$comCode = Strip-Comments185 ((Read-Lines $comPath) -join "`n")
$callsites = ([regex]::Matches($objCode, 'return vb6_ComFailExcep\s*\(')).Count + ([regex]::Matches($colCode, 'return vb6_ComFailExcep\s*\(')).Count
if ($callsites -ne 2) { $fail += "R3 callers of vb6_ComFailExcep = $callsites (want 2: coll_fail + ComObj_Invoke)" }
$externs = ([regex]::Matches($objCode, 'extern HRESULT vb6_ComFailExcep')).Count
if ($externs -ne 1) { $fail += "R3 cross-family extern of vb6_ComFailExcep in vb6comserver_obj.c = $externs (want 1)" }
$nDef = ([regex]::Matches($comCode, 'HRESULT vb6_ComFailExcep\s*\(')).Count
if ($nDef -ne 1) { $fail += "R3 definitions of vb6_ComFailExcep in vb6com.c = $nDef (want exactly 1)" }

# ---- R4 ----
# 收尾那一趟的两行(四空格那版; 早退分支里 12 空格的同名片段不算)必须还在, 且失败交回排在它们
# 之后 —— 否则每次抛错都漏 args/coercedArgs 两批缓冲, 连带 §B162 那枚数组那一臂。
function Index-Of-Re([string[]]$body, [string]$rx) {
    for ($i = 0; $i -lt $body.Count; $i++) { if ($body[$i] -match $rx) { return $i } }
    return -1
}
$iFreeA = Index-Of-Re $invokeBody '^    if \(args\) CoTaskMemFree\(args\);$'
$iFreeC = Index-Of-Re $invokeBody '^    if \(coercedArgs\) CoTaskMemFree\(coercedArgs\);$'
$iFailc = Index-Of $invokeBody 'return vb6_ComFailExcep('
if ($iFreeA -lt 0 -or $iFreeC -lt 0) {
    $fail += "R4 the unconditional cleanup tail (if (args)/if (coercedArgs) CoTaskMemFree) is gone from ComObj_Invoke" }
elseif ($iFailc -lt 0) {
    $fail += "R4 no failure return through vb6_ComFailExcep inside ComObj_Invoke" }
elseif (-not ($iFreeA -lt $iFailc -and $iFreeC -lt $iFailc)) {
    $fail += "R4 failure return($iFailc) must come after the cleanup tail args=$iFreeA coercedArgs=$iFreeC, or every raise leaks both buffers" }

# ---- R5 (账 #305 §B186: 第二站那一格同样只许一枚帧，且登记先于调进去) ----
$createBody = Get-Body $obj '^vb6_ComObject\* vb6_ComObject_Create'
if ($createBody.Count -lt 10) { Write-Host 'FAIL R5 cannot locate vb6_ComObject_Create body'; exit 1 }
$cSet  = Count-In $createBody 'setjmp('
$cSave = Count-In $createBody 'vb6_SaveErrState();'
$cRest = Count-In $createBody 'vb6_RestoreErrState();'
$cCall = Count-In $createBody 'desc->factoryFunc()'
if ($cSet -ne 1)  { $fail += "R5 setjmp in vb6_ComObject_Create = $cSet (want exactly 1)" }
if ($cSave -ne 1) { $fail += "R5 vb6_SaveErrState in vb6_ComObject_Create = $cSave (want exactly 1)" }
if ($cRest -ne 1) { $fail += "R5 vb6_RestoreErrState in vb6_ComObject_Create = $cRest (want exactly 1)" }
if ($cCall -ne 1) { $fail += "R5 desc->factoryFunc in vb6_ComObject_Create = $cCall (want exactly 1)" }
$cSetI  = Index-Of $createBody 'setjmp('
$cPtrI  = Index-Of $createBody 'vb6_error_jmp_ptr = &srvJmpC'
$cCallI = Index-Of $createBody 'desc->factoryFunc()'
if ($cSetI -lt 0 -or $cPtrI -lt 0 -or $cCallI -lt 0) {
    $fail += "R5 missing one of setjmp/jmp_ptr-registration/factoryFunc (set=$cSetI ptr=$cPtrI call=$cCallI)" }
elseif (-not ($cSetI -lt $cPtrI -and $cPtrI -lt $cCallI)) {
    $fail += "R5 order is setjmp($cSetI) -> jmp_ptr($cPtrI) -> factoryFunc($cCallI); must register the frame before calling in" }

# ---- R6 (账 #305 §B186: 这一条通道的失败要分得开，客户端三条腿只问一枚出口) ----
$cVb = Count-In $createBody '*pFailHr = vb6_ComVbHresult('
$cOom = Count-In $createBody '*pFailHr = E_OUTOFMEMORY;'
if ($cVb -ne 1)  { $fail += "R6 vb6_ComVbHresult out-param in vb6_ComObject_Create = $cVb (want exactly 1)" }
if ($cOom -ne 2) { $fail += "R6 E_OUTOFMEMORY out-param in vb6_ComObject_Create = $cOom (want 2: wrapper alloc + factory NULL)" }
$cFree = Count-In $createBody 'CoTaskMemFree(obj);'
if ($cFree -lt 2) { $fail += "R6 the two failure paths must both free obj before returning NULL (found $cFree)" }

$fac  = Read-Lines $facPath
$cfBody = Get-Body $fac 'static HRESULT STDMETHODCALLTYPE CF_CreateInstance'
if ($cfBody.Count -lt 5) { $fail += "R6 CF_CreateInstance body not found" }
if ((Count-In $cfBody 'vb6_ComObject_Create(self->desc, &failHr)') -ne 1) {
    $fail += "R6 CF_CreateInstance does not ask vb6_ComObject_Create for the failure shape" }
if ((Count-In $cfBody 'return E_OUTOFMEMORY;') -ne 0) {
    $fail += "R6 CF_CreateInstance blankets NULL into E_OUTOFMEMORY again - the two questions must stay separable" }

$createCalls = ([regex]::Matches($objCode, 'vb6_ComObject\* vb6_ComObject_Create\s*\(')).Count
if ($createCalls -ne 1) { $fail += "R6 definitions of vb6_ComObject_Create = $createCalls (want 1)" }
$mkCalls = ([regex]::Matches($comCode, 'vb6_ComVbHresult\s*\(')).Count
if ($mkCalls -ne 2) { $fail += "R6 uses of vb6_ComVbHresult in vb6com.c = $mkCalls (want 2: the definition + the scode fill)" }
$failCreateDefs = ([regex]::Matches($comCode, 'static void vb6_ComFailCreate\s*\(')).Count
$failCreateCalls = ([regex]::Matches($comCode, 'vb6_ComFailCreate\s*\(')).Count
if ($failCreateDefs -ne 1) { $fail += "R6 vb6_ComFailCreate definitions = $failCreateDefs (want 1)" }
if ($failCreateCalls -ne 4) { $fail += "R6 vb6_ComFailCreate mentions in vb6com.c = $failCreateCalls (want 4: 1 def + 3 legs)" }
$blanket429 = ([regex]::Matches($comCode, 'vb6_RaiseError\s*\(\s*429')).Count
if ($blanket429 -ne 0) { $fail += "R6 a create leg raises a blanket 429 again ($blanket429) - it must go through vb6_ComFailCreate" }
$externVb = ([regex]::Matches($objCode, 'extern HRESULT vb6_ComVbHresult')).Count
if ($externVb -ne 1) { $fail += "R6 cross-family extern of vb6_ComVbHresult in vb6comserver_obj.c = $externVb (want 1)" }
$hVb = ([regex]::Matches((Strip-Comments185 ((Read-Lines $hPath) -join "`n")), 'HRESULT vb6_ComVbHresult\s*\(')).Count
if ($hVb -ne 1) { $fail += "R6 declaration of vb6_ComVbHresult in vb6com_internal.h = $hVb (want 1)" }

# ---- R7 (账 #305 §B188, 第八十刀): 服务端进门的**第三站** = 出厂那一趟也要问过同一枚帧 ----
# ComObj_Release 走到零时那两条腿 (desc->instanceClaimRelease / desc->destroyFunc) 都踏进 VB 代码
# (Class_Terminate), 而这条通道没有带号的载体 (Release 交回 ULONG, 也没有 EXCEPINFO) => 底线是
# 「不许打死客户端」: 体内必须有一枚帧把两腿一起包住, 且两者都排在 setjmp 之后。
$relStart = -1
$relEnd = -1
for ($i = 0; $i -lt $obj.Count; $i++) {
    if ($relStart -lt 0 -and $obj[$i] -match 'static ULONG STDMETHODCALLTYPE ComObj_Release\(') { $relStart = $i }
    elseif ($relStart -ge 0 -and $obj[$i] -match '^}$') { $relEnd = $i; break }
}
if ($relStart -lt 0 -or $relEnd -lt 0) {
    $fail += ("R7 cannot locate the ComObj_Release body in vb6comserver_obj.c - the way-out station " +
              "was renamed or moved; §B188's reading lives there")
} else {
    $relBody = Strip-Comments185 (($obj[$relStart..($relEnd - 1)]) -join "`n")
    $relSetjmp = ([regex]::Matches($relBody, 'setjmp\s*\(')).Count
    $relClaim = ([regex]::Matches($relBody, 'desc->instanceClaimRelease\s*\(')).Count
    $relDestroy = ([regex]::Matches($relBody, 'desc->destroyFunc\s*\(')).Count
    $relSave = ([regex]::Matches($relBody, 'vb6_SaveErrState\s*\(')).Count
    $relRestore = ([regex]::Matches($relBody, 'vb6_RestoreErrState\s*\(')).Count
    if ($relSetjmp -ne 1) {
        $fail += ("R7 setjmp() inside ComObj_Release = " + $relSetjmp + " (want 1) - an unframed " +
                  "Class_Terminate reaches vb6_RaiseError with no frame => ExitProcess(number) kills the client")
    }
    $relReg = ([regex]::Matches($relBody, 'vb6_error_jmp_ptr = &')).Count
    if ($relReg -ne 1) {
        $fail += ("R7 frame registered inside ComObj_Release = " + $relReg + " (want 1) - a setjmp that nobody " +
                  "registers as the active handler cannot catch the raise: vb6_RaiseError still exits the process")
    }
    if ($relSave -ne 1 -or $relRestore -ne 1) {
        $fail += ("R7 Save/RestoreErrState in ComObj_Release = " + $relSave + "/" + $relRestore + " (want 1/1)")
    }
    if ($relClaim -ne 1 -or $relDestroy -ne 1) {
        $fail += ("R7 way-out legs in ComObj_Release = claim:" + $relClaim + " destroy:" + $relDestroy +
                  " (want 1/1) - both legs must sit inside the same one frame")
    }
    $iSj = $relBody.IndexOf('setjmp')
    $iDe = $relBody.IndexOf('desc->destroyFunc')
    if ($iSj -ge 0 -and $iDe -ge 0 -and $iDe -lt $iSj) {
        $fail += "R7 destroyFunc is called BEFORE setjmp in ComObj_Release - outside the frame, so it still kills the client"
    }
}

# ---- R8 (账 #305 §B188 第九十刀): 出厂那一腿的**尾巴**也要走完 ----
# vb6_RunVbProtected = 「把一截 VB 体跑在自己带的帧里, 交回抛出来的号」这一问的唯一答案;
# 写者只许一个 = 发码里 `_Destroy` 那一行, 且顺序必须 帧 < vb6_Free(me) < 把号交出去 ——
# 号先走、尾巴后拆 与 尾巴没拆 是同一枚缺陷的两种拼法, 所以这一格钉的是顺序不是存在。
$rtlPath = Join-Path $Root 'src\rtl\core\vb6rtl\vb6rtl.c'
$rtlHdr  = Join-Path $Root 'src\rtl\core\vb6rtl\vb6rtl_array.h'
$cgenP   = Join-Path $Root 'src\backend\module\cgen_com.cpp'
$rtPath  = Join-Path $Root 'tests\run_tests.ps1'
$r8Missing = @($rtlPath, $rtlHdr, $cgenP, $rtPath) | Where-Object { -not (Test-Path $_) }
if ($r8Missing) {
    foreach ($m in $r8Missing) { $fail += ("R8 missing file " + $m) }
} else {
    $rtlCode = Strip-Comments185 ((Read-Lines $rtlPath) -join "`n")
    $hdrCode = Strip-Comments185 ((Read-Lines $rtlHdr) -join "`n")
    $cgenCode = Strip-Comments185 ((Read-Lines $cgenP) -join "`n")
    $rtCode = Strip-Comments185 ((Read-Lines $rtPath) -join "`n")
    $r8Def = ([regex]::Matches($rtlCode, 'int32_t\s+vb6_RunVbProtected\s*\(')).Count
    $r8Dec = ([regex]::Matches($hdrCode, 'vb6_RunVbProtected')).Count
    $r8Emit = ([regex]::Matches($cgenCode, 'vb6_RunVbProtected')).Count
    $r8Raise = ([regex]::Matches($cgenCode, 'vb6_RaiseError\(vb6_termErr')).Count
    $r8Free = ([regex]::Matches($cgenCode, 'vb6_Free\(me\);')).Count
    $r8CellA = ([regex]::Matches($rtCode, 'b188_destroy_tail_is_reachable')).Count
    $r8CellB = ([regex]::Matches($rtCode, 'b188_destroy_keeps_no_bare_terminate_call')).Count
    if ($r8Def -ne 1)   { $fail += ("R8 vb6_RunVbProtected definition in vb6rtl.c = " + $r8Def + " (want 1)") }
    if ($r8Dec -ne 1)   { $fail += ("R8 declaration in vb6rtl_array.h = " + $r8Dec + " (want 1)") }
    if ($r8Emit -ne 1)  { $fail += ("R8 emit sites of vb6_RunVbProtected in cgen_com.cpp = " + $r8Emit +
                  " (want 1) - 一截尾巴有两处写法就等于没有那一处") }
    if ($r8Raise -ne 1) { $fail += ("R8 emit site handing the number back = " + $r8Raise + " (want 1)") }
    if ($r8Free -ne 1)  { $fail += ("R8 emit site vb6_Free(me) = " + $r8Free + " (want 1)") }
    if ($r8CellA -ne 1) { $fail += ("R8 run_tests cell b188_destroy_tail_is_reachable = " + $r8CellA + " (want 1)") }
    if ($r8CellB -ne 1) { $fail += ("R8 run_tests cell b188_destroy_keeps_no_bare_terminate_call = " + $r8CellB + " (want 1)") }
    # 帧那一半的形状必须齐: 一枚 setjmp + 一次登记 + Save/Restore 各一 + 转调 body(me) + 读号
    $r8Sj   = ([regex]::Matches($rtlCode, 'setjmp\s*\(frame\)')).Count
    $r8Reg = ([regex]::Matches($rtlCode, 'vb6_error_jmp_ptr = &frame')).Count
    $r8Call = ([regex]::Matches($rtlCode, 'body\s*\(\s*me\s*\)\s*;')).Count
    if ($r8Sj -ne 1 -or $r8Reg -ne 1 -or $r8Call -ne 1) {
        $fail += ("R8 helper internals setjmp/register/call = " + $r8Sj + "/" + $r8Reg + "/" + $r8Call +
                  " (want 1/1/1) - 只 setjmp 不登记就接不住 raise, 不转调体就是空帧")
    }
    $i1 = $cgenCode.IndexOf('vb6_RunVbProtected')
    $i2 = $cgenCode.IndexOf('vb6_Free(me);')
    $i3 = $cgenCode.IndexOf('vb6_RaiseError(vb6_termErr')
    if ($i1 -ge 0 -and $i2 -ge 0 -and $i3 -ge 0 -and -not ($i1 -lt $i2 -and $i2 -lt $i3)) {
        $fail += ("R8 ordering in cgen_com.cpp = frame:" + $i1 + " free:" + $i2 + " hand-back:" + $i3 +
                  " (want frame < free < hand-back) - 号先交出去, 尾巴照旧没人拆 = 本刀没落地")
    }
}
# ---- R9 (账 #305 §B188 第九十一刀): 运行期那一头进了回归, 且那条诊断不再每枚实例打一行 ----
# 第八十九/九十刀把「尾巴走完」钉在发码形状上, 形状针证不了"实例真被拆掉": 判据升格 =
# tests\tools\leak_probe.c (不查注册表的进程内客户, 量 N 轮 create/release 窗口里本进程 commit 的涨量)
# + disp_invoke.ps1 里的 Test-InstanceLeak + run_tests.ps1 里两台各一格。
# 读数吃工具自己给的 verdict (容差一枚页) 而不是字节数: 改前 x64 208896 / x86 90112 字节 = 50 倍/22 倍
# 的带余量, 而 CI 的堆噪声在一页以内。两头同批钉, 因为**判据会随尺漂**——工具改了打印顺序,
# 针就断在这里, 而不是等门红才发现。
# 同一趟还钉节流那一格: [C3_COM] 那一行只有**一枚出口**, 且它的写条件带着 vb6_TermErrClaim ——
# 裸 `if (termErrNum) {` 复原 = 5000 轮 5000 行 stderr, 把现场盖掉的正是 §B181 那条老教训。
$leakTool = Join-Path $Root 'tests\tools\leak_probe.c'
$dispHelp = Join-Path $Root 'tests\disp_invoke.ps1'
$objPath  = Join-Path $Root 'src\rtl\core\vb6comserver\vb6comserver_obj.c'
$r9Missing = @($leakTool, $dispHelp, $objPath) | Where-Object { -not (Test-Path $_) }
if ($r9Missing) {
    foreach ($m in $r9Missing) { $fail += ("R9 missing file " + $m) }
} else {
    $toolCode = Strip-Comments185 ((Read-Lines $leakTool) -join "`n")
    $helpCode = Strip-Comments185 ((Read-Lines $dispHelp) -join "`n")
    $objCode  = Strip-Comments185 ((Read-Lines $objPath) -join "`n")

    $r9Tol = ([regex]::Matches($toolCode, 'LEAK_TOLERANCE_BYTES')).Count
    $r9Warm = ([regex]::Matches($toolCode, 'LEAK_WARMUP_CYCLES')).Count
    $r9Verdicts = ([regex]::Matches($toolCode, 'verdict=(OK|LEAK|BROKEN)')).Count
    if ($r9Tol -ne 3)  { $fail += ("R9 leak_probe.c LEAK_TOLERANCE_BYTES sites = " + $r9Tol +
                  " (want 3: define + the two compares) - 判据不带容差就是让 CI 的堆噪声当裁判") }
    if ($r9Warm -ne 2) { $fail += ("R9 leak_probe.c LEAK_WARMUP_CYCLES sites = " + $r9Warm +
                  " (want 2: define + the pre-window loop) - 窗口里混进首次一次性分配就会把改好的构建判成 LEAK") }
    if ($r9Verdicts -lt 3) { $fail += ("R9 leak_probe.c verdict literals = " + $r9Verdicts +
                  " (want >=3) - tag 与 verdict 必须**相邻**印, run_tests 的针读的就是那一串") }

    $r9Func = ([regex]::Matches($helpCode, 'function Test-InstanceLeak')).Count
    $r9Get  = ([regex]::Matches($helpCode, 'Get-LeakProbeExe')).Count
    $r9Cell = ([regex]::Matches($rtCode, 'b188_way_out_tail_releases\b')).Count
    $r9CellX = ([regex]::Matches($rtCode, 'b188_way_out_tail_releases_x86\b')).Count
    if ($r9Func -ne 1 -or $r9Get -ne 2) {
        $fail += ("R9 helper Test-InstanceLeak/Get-LeakProbeExe = " + $r9Func + "/" + $r9Get +
                  " (want 1/2) - 尺搬进回归之前它只是本机的一次性读数")
    }
    if ($r9Cell -ne 1 -or $r9CellX -ne 1) {
        $fail += ("R9 run_tests cells (x64/x86) = " + $r9Cell + "/" + $r9CellX +
                  " (want 1/1) - 漏一份实例资源是**按指针宽度**算的, 单架构判据不够")
    }

    $r9Claim = ([regex]::Matches($objCode, 'vb6_TermErrClaim')).Count
    $r9Max = ([regex]::Matches($objCode, 'VB6_TERM_ERR_LOG_MAX')).Count
    $r9Log = ([regex]::Matches($objCode, 'fprintf\(stderr, "\[C3_COM\]')).Count
    $r9Bare = ([regex]::Matches($objCode, 'if\s*\(\s*termErrNum\s*\)\s*\{')).Count
    if ($r9Claim -ne 2) { $fail += ("R9 vb6_TermErrClaim sites = " + $r9Claim + " (want 2: definition + the one guard)") }
    if ($r9Max -ne 3)   { $fail += ("R9 VB6_TERM_ERR_LOG_MAX sites = " + $r9Max + " (want 3: define + table + bound)") }
    if ($r9Log -ne 1)   { $fail += ("R9 [C3_COM] diagnostic outlets = " + $r9Log + " (want 1) - 第二枚出口必定是条没节流的流水") }
    if ($r9Bare -ne 0)  { $fail += ("R9 un-throttled `if (termErrNum)` sites = " + $r9Bare +
                  " (want 0) - 每出厂一枚实例打一行, 5000 轮的判据会把真现场盖掉") }
}
if ($fail.Count -gt 0) {
    foreach ($f in $fail) { Write-Host ("FAIL " + $f) }
    exit 1
}
Write-Host ("PASS check_com_server_error_frame: frames=3 way-out-framed=1 registration-ordered shape=1 callers=2 cleanup-before-return factory-shapes-split destroy-tail-framed=1 tail-order=frame<free<hand-back leak-cell-framed=2 term-diag-throttled=1")
exit 0
