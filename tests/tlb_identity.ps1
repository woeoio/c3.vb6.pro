# ============================================================
# ai/022 B13c/B13e: 类型库探针用例（run_tests.ps1 用 `. $PSScriptRoot\tlb_identity.ps1` 引入）
#
# 钉的是两条判据，都从**产物**里读，不读编译器内部状态：
#
# 甲（B13c，读法在 B15 换过一次）：**一个接口在一次编译里只许一枚 GUID**。三条通道各自读一遍：
#   1) `dll_entry.c` 的 `IID_vb6iface_<接口>`            —— COM 服务器表的接口 IID 数组
#   2) `<实现类>.h` 的 `vb6_iv_iid_<接口>[16]`           —— 新式接口 vtable 的 QueryInterface
#   3) `<工程>.tlb` 里**那一档真接口**（`kind=interface name=<接口>`）—— 客户端导入后看到的那一枚
#      B15 之前这里读的是 `kind=dispinterface name=_<接口>`：库里给接口的位置挂着一张 0 成员的
#      dispinterface，外加一行冒充 coclass 的假身份 ⇒ 通道 3 换读法，同时加两条反面断言
#      （见下面的 `又被登记成`），谁把形状改回去这条用例就红。
#
# 乙（B13e）：**广告 == 应答**。类型库给 coclass 标的 DEFAULT 接口，必须就是服务器真正
# 应答成员面的那一枚（`desc->methods` = 类的 Public 成员 = `_<类名>`）：
#   4) `dll_entry.c` 的 `IID_vb6def_<实现类>`  ==  5) `.tlb` 里 coclass 的 DEFAULT 引用
#   且该引用指向 `_<类名>` 本身；coclass 的 CLSID 再与表里的 `clsidStr` 对一次。
#   B13c 曾把 5) 改成跟块里的 `[Default]`（`_<接口>`），实测把两头劈开了 ⇒ 本批回退。
#
# 为什么非要自己写探针：D57-5 记的就是「表与 .tlb 是否同值」到今天没人证过 —— 类型库是
# `CreateTypeLib2` 写出来的二进制，`--emit-c` 不含它，仓里也没有 OleView/tlbimp 依赖。
# `tests\tools\tlbprobe.cpp` 就是那 15 行的读数工具：LoadTypeLib + GetTypeAttr。
# ============================================================

function GuidFromIidInitializer {
    param([string]$Text)
    # C 的 GUID 初始化器 {0xF5CEF988,0x3217,0x6173,{0x94,...}} -> 可读文本 {F5CEF988-...}
    if ($Text -notmatch '\{0x([0-9A-Fa-f]+),0x([0-9A-Fa-f]+),0x([0-9A-Fa-f]+),\{([^}]*)\}\}') { return "" }
    $d1 = $matches[1].ToUpper().PadLeft(8, '0')
    $d2 = $matches[2].ToUpper().PadLeft(4, '0')
    $d3 = $matches[3].ToUpper().PadLeft(4, '0')
    $bytes = @($matches[4] -split ',' | ForEach-Object { $_.Trim().Substring(2).ToUpper() })
    if ($bytes.Count -ne 8) { return "" }
    return "{" + $d1 + "-" + $d2 + "-" + $d3 + "-" + ($bytes[0] + $bytes[1]) + "-" + (($bytes[2..7]) -join "") + "}"
}

function GuidFromIvBytes {
    param([string]$Text)
    # 16 个原始字节 = COM 的落地序（Data1/2/3 小端、Data4 大端），交给 [Guid] 自己解释
    $ms = [regex]::Matches($Text, '0x([0-9A-Fa-f]{2})')
    if ($ms.Count -lt 16) { return "" }
    $arr = New-Object byte[] 16
    for ($i = 0; $i -lt 16; $i++) { $arr[$i] = [Convert]::ToByte($ms[$i].Groups[1].Value, 16) }
    return ([Guid]$arr).ToString('B').ToUpper()
}

function Get-TlbProbeExe {
    $src = Join-Path $Tests "tools\tlbprobe.cpp"
    $exe = Join-Path $OutDir "tlbprobe.exe"
    if ((Test-Path $exe) -and ((Get-Item $src).LastWriteTime -le (Get-Item $exe).LastWriteTime)) { return $exe }
    $prev = (Get-Location).Path
    Set-Location $OutDir          # cl 把 .obj 落在当前目录，跟着产物走就不会脏工作树
    & cl /nologo /EHsc /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib *> $null
    Set-Location $prev
    if (-not (Test-Path $exe)) { return $null }
    return $exe
}

function Test-TlbIdentitySingleSource {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string]$TlbName,       # .tlb 的文件名主体（= 工程输出名）
        [string]$ClassName,     # 实现类模块名
        [string]$IfaceName,     # 块里 [Default] 指的那个新式接口名
        [string]$ExpectedClsid  # 期望的 CLSID（带花括号、大写）
    )
    # vbp 分片: 本片不跑这例 (闸门在 run_tests.ps1 里定义)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [TLB-ID] $Name ... "

    $out = & $C3 $VbpFile --output-dir $OutDir --keep-for-debug 2>&1
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        return
    }
    $genDir = $null
    foreach ($line in $out) {
        $t = "$line"
        $at = $t.IndexOf("intermediates kept at: ")
        if ($at -ge 0) { $genDir = $t.Substring($at + 23).Trim(); break }
    }
    if (-not $genDir -or -not (Test-Path $genDir)) {
        $script:fail++
        Write-Host "FAIL (no intermediates)" -ForegroundColor Red
        return
    }
    $probe = Get-TlbProbeExe
    if (-not $probe) {
        $script:fail++
        Write-Host "FAIL (no probe)" -ForegroundColor Red
        return
    }
    $tlbPath = Join-Path $genDir "$TlbName.tlb"
    if (-not (Test-Path $tlbPath)) {
        $script:fail++
        Write-Host "FAIL (no tlb)" -ForegroundColor Red
        return
    }
    $tlbText = ((& $probe $tlbPath) | Out-String)
    $entry = ""
    $hdr = ""
    $p = Join-Path $genDir "dll_entry.c"
    if (Test-Path $p) { $entry = Get-Content $p -Raw }
    $p = Join-Path $genDir "$ClassName.h"
    if (Test-Path $p) { $hdr = Get-Content $p -Raw }

    # 通道 1a：COM 服务器表里"接口自己的 IID"
    $tableIfaceIid = ""
    if ($entry -match ("static const IID IID_vb6iface_" + $IfaceName + " = (\{[^;]*\})")) {
        $tableIfaceIid = GuidFromIidInitializer $matches[1]
    }
    # 通道 1b：COM 服务器表里"默认接口的 IID"（服务器按它应答早绑定 QI）
    $tableIid = ""
    if ($entry -match ("static const IID IID_vb6def_" + $ClassName + " = (\{[^;]*\})")) {
        $tableIid = GuidFromIidInitializer $matches[1]
    }
    # 通道 2：类自己的 vtable QI
    $ivIid = ""
    if ($hdr -match ("vb6_iv_iid_" + $IfaceName + "\[16\] = (\{[^}]*\})")) {
        $ivIid = GuidFromIvBytes $matches[1]
    }
    # 通道 3：类型库里那一档**真接口**（ai/022 B15：TKIND_INTERFACE，名字 = 接口自己的名字）
    $tlbIface = ""
    if ($tlbText -match ("kind=interface name=" + $IfaceName + " guid=(\{[0-9A-F-]{36}\})")) {
        $tlbIface = $matches[1]
    }
    $tlbClass = ""
    if ($tlbText -match ("kind=dispinterface name=_" + $ClassName + " guid=(\{[0-9A-F-]{36}\})")) {
        $tlbClass = $matches[1]
    }
    $coclassGuid = ""
    $defaultRef = ""
    $tlbLines = $tlbText -split "`r?`n"
    for ($i = 0; $i -lt $tlbLines.Count; $i++) {
        if ($tlbLines[$i] -notmatch ("kind=coclass name=" + $ClassName + " guid=(\{[0-9A-F-]{36}\})")) { continue }
        $coclassGuid = $matches[1]
        for ($j = $i + 1; $j -lt $tlbLines.Count; $j++) {
            if ($tlbLines[$j] -notmatch '^\s+REF') { break }
            if ($tlbLines[$j] -match ' DEFAULT' -and $tlbLines[$j] -notmatch 'SOURCE') {
                if ($tlbLines[$j] -match 'guid=(\{[0-9A-F-]{36}\})') { $defaultRef = $matches[1] }
            }
        }
        break
    }

    $detail = @()
    if (-not $tableIfaceIid) { $detail += "dll_entry.c 里没有 IID_vb6iface_$IfaceName" }
    if (-not $tableIid)   { $detail += "dll_entry.c 里没有 IID_vb6def_$ClassName" }
    if (-not $ivIid)      { $detail += "$ClassName.h 里没有 vb6_iv_iid_$IfaceName" }
    if (-not $tlbIface)   { $detail += "类型库里读不到真接口 $IfaceName 那一档的 GUID" }
    if (-not $tlbClass)   { $detail += "类型库里读不到 _$ClassName 的 GUID" }
    if (-not $coclassGuid){ $detail += "类型库里读不到 coclass $ClassName" }
    if (-not $defaultRef) { $detail += "coclass $ClassName 没有 DEFAULT 接口引用" }
    # 甲：一个接口在一次编译里只许一枚 GUID（vtable QI / 表 / 类型库三条通道同值）
    if ($tableIfaceIid -and $ivIid -and $tableIfaceIid -ne $ivIid) {
        $detail += "表的接口 IID 与 vtable 不同值: $tableIfaceIid vs $ivIid"
    }
    if ($tableIfaceIid -and $tlbIface -and $tableIfaceIid -ne $tlbIface) {
        $detail += "表的接口 IID 与类型库真接口 $IfaceName 不同值: $tableIfaceIid vs $tlbIface"
    }
    # 甲的两条反面断言（ai/022 B15）：接口宿主那一档曾经是一张 0 成员的 `_<接口>` dispinterface
    # 外加一行冒充 coclass 的假 CLSID。形状改回去 = 本批白做，所以这里直接判红。
    if ($tlbText -match ("kind=coclass name=" + $IfaceName + " guid=")) {
        $detail += "库里 $IfaceName 又占了一行 coclass（一枚假 CLSID + 一句可创建）"
    }
    if ($tlbText -match ("kind=dispinterface name=_" + $IfaceName + " guid=")) {
        $detail += "库里 $IfaceName 又有一张 0 成员的 _$IfaceName dispinterface"
    }
    # 乙：广告的那一枚 == 服务器应答的那一枚（B13e）—— 类型库说谁是默认接口，
    # 服务器的成员面与早绑定 QI 就必须是谁。`_<类名>` 正是 `desc->methods` 那一档。
    if ($tableIid -and $defaultRef -and $tableIid -ne $defaultRef) {
        $detail += "表的默认接口 IID 不等于类型库的 DEFAULT 引用: $tableIid vs $defaultRef"
    }
    if ($defaultRef -and $tlbClass -and $defaultRef -ne $tlbClass) {
        $detail += "类型库的 DEFAULT 引用不是 ${ClassName}: $defaultRef vs $tlbClass"
    }
    if ($coclassGuid -ne $ExpectedClsid) { $detail += "类型库 coclass = $coclassGuid，期望 $ExpectedClsid" }
    if ($entry -notmatch ('"' + [regex]::Escape($ExpectedClsid) + '",\s*/\*\s*clsidStr')) {
        $detail += "dll_entry.c 的 clsidStr 不是 $ExpectedClsid"
    }

    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (one IID per interface, stored as a real interface; advertised default == answered default)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { foreach ($d in $detail) { Write-Host "        $d" -ForegroundColor DarkGray } }
        else { Write-Host ("        " + ($detail -join '; ')) -ForegroundColor DarkGray }
    }
}

# ============================================================
# ai/022 B13d: 规范 IUnknown 的形状断言
# ============================================================
# COM 的规矩是"同一个对象不管从哪个接口问 IUnknown，拿到的必须是同一个指针"。
# 薄指针那条路以前把 `IID_IUnknown` 与本接口的 IID 并成一个分支、回 `*ppv = self`，
# 于是两个接口各问一次会拿到两个不同地址（`__iv_A` 与 `__iv_B` 是同一个结构体里
# 两个不同偏移的成员，这一点没有含糊空间）。本批改成一律回"本类实现序第一个接口"
# 的薄指针 —— 那既是稳定身份，也是一份偏移 0 就是 vtable 的合法对象指针。
#
# 为什么断形状而不是断行为：整条代码路径上**没有任何自己生成的调用点**去问
# `IID_IUnknown`（VB 侧没有 QueryInterface 面，接口值也不能进 Variant = B06c），
# 真正的调用者是外部 COM 客户 ⇒ 那是 B17 的读数。这里钉的是"每个 QI 都回同一个
# 规范指针、且旧的合并分支已经不在"，配一条真编译真跑（`itf_xmod_writer` 的
# QI1..QI4）保证兄弟接口/本接口两条分支没有被改坏。
function Test-CanonicalIUnknownShape {
    param(
        [string]$Name,
        [array]$Sources,
        [string]$CanonNeedle,     # 例: "void* canon = &me->__iv_IWriter;"
        [int]$CanonCount          # 该工程里实现新式接口的类的 QI 份数
    )
    $script:total++
    Write-Host -NoNewline "  [IUNK-SHAPE] $Name ... "
    $text = Invoke-CodegenProj $Sources
    $detail = @()
    if ($script:codegenProjExit -ne 0) { $detail += "codegen exit=$script:codegenProjExit" }
    $hits = ([regex]::Matches($text, [regex]::Escape($CanonNeedle))).Count
    if ($hits -ne $CanonCount) { $detail += "规范指针分支 $hits 处，期望 $CanonCount 处" }
    # 旧形状：IUnknown 与本接口并成一个分支、回 self —— 每个接口各回各的，身份不恒等
    if ($text.Contains("vb6_IidEqual(riid, vb6_iv_iid_IUnknown) ||")) {
        $detail += "还有把 IID_IUnknown 与本接口并在一起的分支（回 self = 每个接口一个身份）"
    }
    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS ($CanonCount QI 都回同一枚规范指针)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host ("        " + ($detail -join '; ')) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}
