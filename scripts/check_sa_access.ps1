# 账 #209 + 账 #214 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 一维动态数组的**元素**访问在 RTL 里只有一个入口 —— VB6_SA_AT。
# 这一刀之前它是裸指针算术 (`((type*)((arr)->data))[idx - (arr)->lBound]`), 描述符是 NULL
# (动态数组从未 ReDim / 已被 Erase) 或下标越界时它一声不吭地读 NULL+0xc → 原生 0xC0000005。
# 实测 Charts 2020 ucChartBar demo 点 Random: 三次崩在同一偏移 0x1abf2, 符号化落在
# 生成代码 ucChartBar.c:591 的 `With m_Serie(Index)`。VB6 在这一条是运行时错误 9, 和
# vb6_UBound/vb6_LBound 的 rev2 同一族 —— 那两处已经抛 9, 这一处漏了。
#
# 现在两支都收在同一形状的出口上: 一维走 vb6_SaElemPtr, 多维走 vb6_SaNdElemPtr -- (各一条 inline 热路径 + 一条冷路径)
# 账 #214 是这一族的下一半: VB6_SA_ND_AT1/2/3 当初与 #209 修法前的一维一模一样 (描述符在不在、
# 每一维在不在范围, 一个都不问), 而 4 秩以上另有一条 `_ndoff_` 兜底 —— 那条更糟: 发码侧把实参写成
# `(int[]){i0, i1}` 只塞两枚下标却按实际秩数交出去, 于是 `Dim a4(1 To 2,1 To 2,1 To 2,1 To 2)`
# **在范围内也** 0xC0000005 (x86/x64 同形)。两支合起来存量 1468 处 / 4 份工程, 大头是 VBFlexGridDemo。
# 所以 A4 现在改成“一条都不许有”, 另加 A6 (多维那一支的唯一入口) 与 A7 (发码侧的形状)。
#
# 规则 (改坏了会红, 不是装饰):
#   A1  VB6_SA_AT 的宏定义恰好 1 处, 且宏体走 vb6_SaElemPtr (不许再退回裸算术)
#   A2  vb6_SaElemPtr 的 inline 定义恰好 1 处, 体内两条比较(下界/上界)与 NULL 那条都还在
#   A3  vb6_SaElemFail 声明 1 + 定义 1, 定义体里 vb6_ErrRaise(9 那条还在
#   A4  src/backend 里非注释的 `->data` 元素寻址行 = 0 (两支都不许再手算元素地址)
#   A5  src/rtl 与 src/backend 里 `->data))[` 这种"宏里手算下标"的形状 = 0
#   A6  多维那一支: AT1/AT2/AT3/ATN 四条宏各 1 处且宏体都经 vb6_SaNdElemPtr;
#       vb6_SaNdElemPtr 定义 1 处, 体内形状问句(dimCount)+两条上下界比较+冷路径调用都在;
#       vb6_SaNdElemFail 声明 1 + 定义 1, 定义体真的抛 9
#   A7  发码侧 4 秩以上那一条: 发 VB6_SA_ND_ATN, 下标实参按 `((const int32_t[]){` 拼(最外层括号
#       不是装饰, 少了它就是 cl C4002), 且不再手拼 `(int[]){` / 不再直接发 vb6_SafeArrayND_Offset(

# --- A8..A11（账 #305 §B154，第四十二刀）：一枚被 ReDim 重建的描述符「按哪一档分配」只许有一处答案 ---
# 旧形状 = cgen_redim.cpp 只问 ReDim 语句自己的 As 子句，没写 As 就回落 vb6_sa_variant；而**访问侧**
# （VB6_SA_AT / VB6_SA_ND_AT2）问的是声明表 arrayElemTypes_。同一枚描述符两个答案：造出来 24 字节
# 一步、写进去 8 字节一步，直接读写自己一致所以看着是好的，而凡信描述符自己那一档的消费方
# （Destroy / Erase / 装箱后的元素读）全走错步长 —— tests/test_redim_elem.bas 的 R3/R4 两枚 box-cell
# 在修前读回**空串**（两台同形，读数 .build/b1373_fix/{rnb6,rf6,rf3}.txt）。
#   A8  唯一出口 resolveReDimElemType：定义 1、声明 1、在 cgen_redim.cpp 被叫 1 次以上，函数体两头都
#       要问（As 子句 + arrayElemTypes_）；cgen_redim.cpp 里不许再出现「声明处只问 As」那一行。
#   A9  src/backend 非注释里 vb6_sa_variant 的字面量恰好 3 处：分档表的 Variant 档、它的 default 档、
#       以及 Variant 槽那一枚**刻意写死**的 Preserve（那一支的档是权威，不是猜）。第四处 = 又有人在
#       某个调用点手写兜底档。
#   A10 多维描述符投进 _Generic 装箱：装箱里那一层投影恰好 1 处；反向的「强转盖住装箱」恰好 0 处
#       （那是 BASE 上 cl C2440 的形状；裸递 ND 描述符则掉进 _Generic 的 default = vb6_VariantObject，
#        槽的 vt 成了 DISPATCH，Set/GetND 只认 vt & vtArray ⇒ 一声不响地既不写也不读）。
#   A11 声明表自己不许空转：写档站点 >= 8、.find 站点 >= 3、.erase 站点 == 1（过程开始清表那一条还在；
#       撤掉它，上一个过程的元素档会漏给下一个过程）。
# --- A12/A13（账 #305 §B155 格④，第四十三刀）：多维 Preserve 那枚出口自己也得收档 ---
# vb6_SafeArrayReDimPreserveND 的形参表里原本压根没有 elemType，于是 arr==NULL（首次分配就走 Preserve）
# 只能按 vb6_sa_empty 建 —— 尺寸表对 empty 答 default 4 字节，而访问侧按声明档走步长（x64：String 8、
# Variant 24），写出去就落在只有 1/2 到 1/6 大小的块外面。这一族是**响**的：探针 z.bas 修前实测
# 0xC0000374 与 0xC0000005 两种死法、另有一次侥幸跑完（读数 .build/b1382_preserve/ 与 .build/b1395_k43/），
# 而**直接读与直接写用的是同一枚错步长**，所以两者一致、看着是好的 —— 能分辨修前的只有「信描述符自己
# 那一档的消费方」（装箱后的元素读）那一条，判据就按那条钉。
#   A12 这一枚出口的头与体都收 elemType（各恰好 1；头追不上体在 runner 上是 cl C2197，见账 #240）；
#       体内首次分配那一条必须按调用方给的档建，vb6_sa_empty 那枚兜底不许再出现，udt 那一挡也还得在
#       （尺寸表对 udt 答 0 ⇒ calloc(n, 0) 给一枚非空的零长指针，写进去就是纯堆越界）。
#   A13 后端发这一枚的行恰好 1 处，且那一处把 saElemType 递了进去（0 处 = RTL 有形参而后端不递：C 里
#       实参不足只报 warning 照编，于是「运行时自己猜档」原地复活）。
# --- A14/A15（账 #305 §B156，第四十四刀）：多维 ReDim 的界数组临时量只许一处拼法、且按语句唯一 ---
# 旧形状 = 两条发码路各按**名字**裸拼一条 C 局部数组（visit(ReDimStmt) 用 _redim_bounds_<数组名>，
# emitReDimComplexTarget 用 _redim_bounds<净化后的左值>）。名字里没有语句序号，于是同一过程里
# 第二条指向同一目标的多维 ReDim 就在同一个 C 函数里把 `vb6_SafeArrayBound <名>[] = {…};` 发两遍
# ⇒ cl C2374「重定义；多次初始化」（两台同形、比第四十二/四十三两刀都老，读数 .build/b1395_k43/）。
# 真语料那一形 43 份文件起步（按文件计的上界），门禁语料两份**恰好都分在不同过程** ⇒ 门一直不红。
#   A14 拼法只许一处：nextNDBoundsTemp 声明 1（cgen_helpers.inc）/ 定义 1（cgen_redim.cpp）/ 被叫恰好 2
#       （两条发码路都在）；`"_redim_bounds` 这枚字面量在 src/backend 非注释里恰好 1（就在出口里），
#       旧的 `std::string boundsVar = "_redim_bounds` 裸拼形状恰好 0。
#   A15 序号与 tempCounter_ 同生同灭：ndBoundsSeq_ 声明 1、`ndBoundsSeq_ = 0` 恰好 1（在 prologue 里，
#       与 tempCounter_ 那一处同处）；固定数组那条 `_bounds_` 仍恰好 2 处 —— 它按声明名，而 VB 不许
#       同过程重名声明，所以**不在**这条权威里（把它卷进来 = 第二份答案）。
# --- A16（账 #305 §B158，第四十五刀）：多维数组「第几格」只许一套算术 ---
# 访问侧那两套（头里的 vb6_SaNdElemPtr 与体里的 vb6_SafeArrayND_Offset）写的是同一件事：
# off = Σ (idx[d]-lb[d])*stride[d]，stride[0]=1、stride[d]=stride[d-1]*cnt[d-1] —— **第 0 维最快**。
# 而两枚 Preserve(多维) 搬旧元素时各自另写了一份 stride[d]=stride[d+1]*cnt[d+1]（最后一维最快），
# 两套算术一致的那一格只有 (1,1)；其余格在「长一格」时被搬到别处（实测：2×3 长成 2×6 之后
# POST-1-3 读回空、r2c2 出现在 POST-1-4、r1c3 出现在 POST-2-4，.build/b1403_k44probe/rpn4_pre45.txt）。
#   A16 收成一枚 static 出口 vb6_SaNdSlot：定义恰好 1、出现（定义+调用）恰好 5（两枚 Preserve 各 2
#       次调用：旧格位与新格位）；oldStrides/newStrides 那两份表各恰好 0（它们就是第二套算术）；
#       丢弃旧元素那一段的下标枚举回到降序（`for (d = arr->dimCount - 1; ... d--)`）就报红 ——
#       枚举序就是线性格位序，降序枚举算出来的 indices 与 linear 那枚格位不是同一套对应关系。
# --- A17（账 #305 §B155 格①②，第四十六刀）：「这枚数组是哪一档」只许一处答 ---
# 改前「这枚数组是哪一档」说了四遍，每一遍都不问描述符：vb6_VariantArray 把 vt 写死成
# vb6_vtArray|vb6_vtVariant（于是一切 typed 数组装箱后 VarType 一律 8204、TypeName 一律
# "Variant"，而 VB6 是 8200/"String()" 这一族）、vb6_VariantFromUdtBytes 写死 BYTE、
# vb6rtl_com.c 那两条 VT_ARRAY 进口各自折一枚三元表达式（BSTR/VARIANT 之外一律答 vtLong ⇒
# VT_ARRAY|VT_I2 进来标签答 Long，而载体是 int16 —— 标签与描述符两个答案，拿着错标签交回
# COM 就是错档）。TypeName 那边压根没有数组分支，8204 落 default 答 "Variant"。
#   A17 新表 vb6_SaVtFromElemType 声明 1 / 定义 1 / 调用恰好 5（array.h 的 UDT 字节装箱 1 + com.c 那
#       条进口 2 + array.c 自己那枚认载体的出口 2 + 出站白名单 1，第五十刀）；反向表 vb6_SaElemTypeFromVt
#       声明 1 / 定义 1 / 调用恰好 3（com.c 那两条 VT_ARRAY 进口 + §B162 入站那一臂），手抄的元素 switch 回到 0 处；认载体那一问 vb6_SaVtOfArrayPtr 声明 1 / 定义 1 /
#       调用恰好 2（装箱 + GetAllSettings）；写死的 vb6_vtArray|vb6_vtVariant 恰好 2 且**只能在**
#       vb6_SaVtOfArrayPtr 的两条兜底里（NULL 与认不出载体）；旧的三元折档形状 `vb6_vtArray | (baseVt`
#       恰好 0。TypeName 那一张标量名单只许一份：`case vb6_vtEmpty:` 在 vb6rtl_conv.c 恰好 1，
#   A19（§B159 第五十刀）出站那一半：白名单 vb6_ComSaVtOut 里批准的每一档，反向表必须认得（否则
#       写进去真数组、读回来是别家档 = 半程往返，第五十刀的 Byte 那一档就是这么现形的）；真数组
#       只许一处造（vb6_ComPackArrayOut 定义 1 / 调用 2 —— 第三十刀起多了一位消费者：数组形参的出站回写），数组那一问在 PackBorrow 的 switch 之前 1 处；
#       入站那一臂同样只许一处认（vb6_VariantArrayIntoSlot 定义 1 / 调用 1），它造的 descriptor
#       只许一处收（vb6_VariantSlotRelease 在服务端 free 槽阵列之前调一次）。
#       抽出来的 vb6_TypeNameScalar 定义 1 / 调用恰好 2（数组那一支 + 标量那一条）。
# --- A18（账 #305 §B160/§B161，第四十八刀 + 第四十九刀 + 第五十刀）：数组递给 ByRef Variant 形参时
#     「这是不是一枚数组」那一问只许一处答, 而且必须答在按 VB 类型分档的那条 switch 之前 ---
# 改前那一支只认两种实参：晚绑定对象 (void*) 与工程类实例，其余一律落 `&arr` ——
# 数组变量的 C 表示是描述符指针，`&arr` 就是「描述符指针的地址」，而被调方按
# vb6_VARIANT* 解析 ⇒ 静默垃圾（真跑读数 VarType 40320 / TypeName "Variant"，VB6 交
# 8195 / "Long()"；cl 只给一条不兼容指针的警告，BUILD-RC=0）。第四十八刀把装箱补在
# switch 的 default 里, 只打到了过程内 Dim 的数组; **模块级**数组被 inferExprType 答成元素档,
# 走的是标量那几臂 —— 实测 m-long=3/"Long"、m-byte=17/"Byte"、m-str=8/"String",
# 元素读静默答 evt=0 val=0（探针矩阵 .build/k48probe/P49Main.bas 八条形状全录在案）。
# 第四十九刀把这一问提到 switch 之前, 判据吃数组表 knownArrays_（与元素访问 VB6_SA_AT 同一张）,
# 装箱仍走外面那一层 vb6_VariantFromValue。
#   A18 谓词 arrArg92y 定义恰好 1 / 使用恰好 1（那条 if）；问数组表 knownArrays_.count(arrLs92y)
#       恰好 1；临时量前缀 "_vb6_arrarg_" 恰好 1；装箱吃的那枚原样名 arrRaw92y 恰好 2（定义 + 装箱）；
#       第四十八刀那条住在 switch 里的 argIsArr92y 恰好 0（同一问不许住两处）；
#       2026-10-26 第二百三十九刀（账 #305 §B224 第二半）加了第四位消费者 comByRefArgExpr：按引用发那一问
#       也要先问「这是不是数组载体」（数组不许取址，否则 `&arr` 就是描述符指针的地址）。
#       顺序两条：定义行在 switch (argVbType92y) **之前**，装箱那一条在它**之后**。
#       站点计数全在 src/backend 里数（只剥行注释；后端 .inc 的字面量里有 /* */，
#       剥块注释会把真代码整段吃掉 —— 第四十八刀那条假红就是这么来的）。

#
# 用法:  pwsh -File scripts\check_sa_access.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function Lines([string]$path) {
    ([System.IO.File]::ReadAllText($path) -split "`r?`n")
}

$rtlDir = Join-Path $root "src\rtl"
$beDir = Join-Path $root "src\backend"
$arrH = Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_array.h"
$arrC = Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_array.c"
foreach ($p in @($arrH, $arrC)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host ("FAIL A0 file missing: " + $p) -ForegroundColor Red
        exit 1
    }
}

# A1: 宏只有定义点这一处写法, 而且必须经 vb6_SaElemPtr
$macroDefs = @()
$macroBodyCallsHelper = $false
foreach ($f in (Get-ChildItem -LiteralPath $rtlDir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h" })) {
    $ln = 0
    foreach ($line in (Lines $f.FullName)) {
        $ln++
        if ($line.Trim() -match '^#\s*define\s+VB6_SA_AT\(') {
            $macroDefs += ($f.Name + ":" + $ln)
            $body = (Lines $f.FullName)[$ln..([Math]::Min($ln + 2, (Lines $f.FullName).Count - 1))] -join " "
            if ($body -match "vb6_SaElemPtr") { $macroBodyCallsHelper = $true }
        }
    }
}
if ($macroDefs.Count -ne 1) {
    $bad += ("A1 VB6_SA_AT macro definitions = " + $macroDefs.Count + " (want exactly 1) -> " + ($macroDefs -join " | "))
}
if ($macroDefs.Count -eq 1 -and -not $macroBodyCallsHelper) {
    $bad += "A1 VB6_SA_AT body no longer routes through vb6_SaElemPtr (the check moved back to raw math)"
}

# A2: 那个唯一的 inline, 三条检查一条不许哑
$h = [System.IO.File]::ReadAllText($arrH)
$defs = @([regex]::Matches($h, 'static\s+inline\s+void\*\s+vb6_SaElemPtr\s*\('))
if ($defs.Count -ne 1) {
    $bad += ("A2 vb6_SaElemPtr defined " + $defs.Count + " times (want 1)")
} else {
    $blk = [regex]::Match($h, 'static\s+inline\s+void\*\s+vb6_SaElemPtr[\s\S]{0,800}?\r?\n\}')
    if ($blk.Value -notmatch '!arr') { $bad += "A2 NULL-descriptor check is gone from the helper" }
    if ($blk.Value -notmatch 'idx\s*<\s*arr->lBound') { $bad += "A2 lower-bound compare is gone from the helper" }
    if ($blk.Value -notmatch 'idx\s*>\s*arr->uBound') { $bad += "A2 upper-bound compare is gone from the helper" }
    if ($blk.Value -notmatch 'vb6_SaElemFail') { $bad += "A2 the cold-path call is gone from the helper" }
}

# A3: 冷路径声明 + 定义各 1, 且真的抛 9
$decl = @([regex]::Matches($h, 'void\s+vb6_SaElemFail\s*\(')).Count
$csrc = [System.IO.File]::ReadAllText($arrC)
$defn = @([regex]::Matches($csrc, '\r?\nvoid\s+vb6_SaElemFail\s*\(')).Count
if ($decl -ne 1) { $bad += ("A3 vb6_SaElemFail 声明 = " + $decl + " (want 1)") }
if ($defn -ne 1) { $bad += ("A3 vb6_SaElemFail 定义 = " + $defn + " (want 1)") }
$cblk = [regex]::Match($csrc, 'void\s+vb6_SaElemFail[\s\S]{0,1200}?\r?\n\}')
if (-not $cblk.Success -or $cblk.Value -notmatch 'vb6_ErrRaise\(\s*9\b') {
    $bad += "A3 cold path no longer raises runtime error 9 (silent return would spare sites without On Error)"
}

# A4: 后端不再许有任何一处手算元素地址的行 (两支都收进出口了)
$rawData = @()
foreach ($f in (Get-ChildItem -LiteralPath $beDir -Recurse -File | Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
    $ln = 0
    foreach ($line in (Lines $f.FullName)) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($line.Contains('->data')) { $rawData += ($f.Name + ":" + $ln) }
    }
}
if ($rawData.Count -ne 0) {
    $bad += ("A4 后端手算元素地址的行 = " + $rawData.Count +
             " (want 0: 一维走 VB6_SA_AT, 多维走 VB6_SA_ND_AT*, 都不该再碰 ->data) -> " + ($rawData -join " | "))
}

# A5: "宏里手算下标"这个形状在整个仓里不许复活
$shape = 0
foreach ($dir in @($rtlDir, $beDir)) {
    foreach ($f in (Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h", ".cpp", ".inc", ".hpp" })) {
        foreach ($line in (Lines $f.FullName)) {
            if ($line.Trim().StartsWith("//")) { continue }
            if ($line.Contains('->data))[')) { $shape++ }
        }
    }
}
if ($shape -ne 0) {
    $bad += ("A5 裸 `->data))[` 元素寻址 = " + $shape + " 处 (want 0; 一维那支必须只有宏这一个入口)")
}

# A6: 多维那一支的入口与热路径问句
$ndMacros = @()
$ndViaHelper = 0
foreach ($f in (Get-ChildItem -LiteralPath $rtlDir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h" })) {
    $fl = Lines $f.FullName
    $ln = 0
    foreach ($line in $fl) {
        $ln++
        if ($line -match '^\s*#\s*define\s+(VB6_SA_ND_AT\d|VB6_SA_ND_ATN)\b') {
            $ndMacros += ($Matches[1] + "@" + $f.Name + ":" + $ln)
            $hi = [Math]::Min($ln + 2, $fl.Count - 1)
            $body = $fl[$ln..$hi] -join " "
            if ($body -match "vb6_SaNdElemPtr") { $ndViaHelper++ }
        }
    }
}
if ($ndMacros.Count -ne 4) {
    $bad += ("A6 多维元素访问宏 = " + $ndMacros.Count + " 处 (want 4: AT1/AT2/AT3/ATN) -> " + ($ndMacros -join " | "))
}
if ($ndMacros.Count -eq 4 -and $ndViaHelper -ne 4) {
    $bad += ("A6 宏体经 vb6_SaNdElemPtr 的条数 = " + $ndViaHelper + " (want 4; 有一支退回裸算术了)")
}
$ndDefs = @([regex]::Matches($h, 'static\s+inline\s+void\*\s+vb6_SaNdElemPtr\s*\('))
if ($ndDefs.Count -ne 1) {
    $bad += ("A6 vb6_SaNdElemPtr defined " + $ndDefs.Count + " times (want 1)")
} else {
    $ndblk = [regex]::Match($h, 'static\s+inline\s+void\*\s+vb6_SaNdElemPtr[\s\S]{0,1200}?\r?\n\}')
    if ($ndblk.Value -notmatch '!a') { $bad += "A6 NULL-descriptor check is gone from the ND helper" }
    if ($ndblk.Value -notmatch 'dimCount') { $bad += "A6 the ND helper no longer asks the shape (rank vs dimCount)" }
    if ($ndblk.Value -notmatch 'idx\[d\]\s*<\s*lb') { $bad += "A6 lower-bound compare is gone from the ND helper" }
    if ($ndblk.Value -notmatch 'idx\[d\]\s*>=\s*lb\s*\+\s*cnt') { $bad += "A6 upper-bound compare is gone from the ND helper" }
    if ($ndblk.Value -notmatch 'vb6_SaNdElemFail') { $bad += "A6 the cold-path call is gone from the ND helper" }
}
$ndDecl = @([regex]::Matches($h, 'void\s+vb6_SaNdElemFail\s*\(')).Count
$ndDefn = @([regex]::Matches($csrc, '\r?\nvoid\s+vb6_SaNdElemFail\s*\(')).Count
if ($ndDecl -ne 1) { $bad += ("A6 vb6_SaNdElemFail 声明 = " + $ndDecl + " (want 1)") }
if ($ndDefn -ne 1) { $bad += ("A6 vb6_SaNdElemFail 定义 = " + $ndDefn + " (want 1)") }
$ndCblk = [regex]::Match($csrc, '\r?\nvoid\s+vb6_SaNdElemFail[\s\S]{0,1400}?\r?\n\}')
if (-not $ndCblk.Success -or $ndCblk.Value -notmatch 'vb6_ErrRaise\(\s*9\b') {
    $bad += "A6 ND cold path no longer raises runtime error 9 (silent return would spare sites without On Error)"
}

# A7: 发码侧 4 秩以上那一支的形状
$atnEmit = 0; $atnLiteral = 0; $handIdxList = 0; $offsetEmit = 0
foreach ($f in (Get-ChildItem -LiteralPath $beDir -Recurse -File | Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
    $ln = 0
    foreach ($line in (Lines $f.FullName)) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($t.Contains("VB6_SA_ND_ATN(")) { $atnEmit++ }
        if ($t.Contains('((const int32_t[])')) { $atnLiteral++ }
        if ($t.Contains('(int[]){')) { $handIdxList++ }
        if ($t.Contains('vb6_SafeArrayND_Offset(')) { $offsetEmit++ }
    }
}
if ($atnEmit -lt 1) { $bad += "A7 后端不再发 VB6_SA_ND_ATN (4 秩以上那支又回到手算偏移了)" }
if ($atnLiteral -ne 1) {
    $bad += ("A7 ATN 的下标实参不是按 ((const int32_t[]){ 拼的 = " + $atnLiteral +
             " 处 (want 1; 少最外层那对括号 ⇒ cl C4002 参数过多)")
}
if ($handIdxList -ne 0) {
    $bad += ("A7 手拼 (int[]){ 的下标列表又出现 " + $handIdxList +
             " 处 (那正是「只塞两枚下标却按实际秩数交出去」那个 bug 的形状)")
}
if ($offsetEmit -ne 0) {
    $bad += ("A7 后端直接发 vb6_SafeArrayND_Offset( 的行 = " + $offsetEmit +
             " (want 0: 元素访问只许走 AT* 那几个出口)")
}


# A8: ReDim 的元素档只问一处出口
$beFiles = @(Get-ChildItem -LiteralPath $beDir -Recurse -File | Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })
$rdDef = 0; $rdDecl = 0; $rdCall = 0; $rdAsOnly = 0; $rdBodyBoth = 0; $rdBodyMember = 0
foreach ($f in $beFiles) {
    $txt = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($line in (Lines $f.FullName)) {
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($t -match '^Vb6Type\s+CCodeGen::resolveReDimElemType\s*\(') { $rdDef++ }
        if ($f.Name -eq "cgen_helpers.inc" -and $t -match '^Vb6Type\s+resolveReDimElemType\s*\(') { $rdDecl++ }
        if ($f.Name -eq "cgen_redim.cpp") {
            if ($t.Contains("resolveReDimElemType(")) { $rdCall++ }
            if ($t -match '^Vb6Type\s+elemType\s*=\s*resolveArrayElemType\s*\(') { $rdAsOnly++ }
        }
    }
    if ($f.Name -eq "cgen_expr_array.cpp") {
        # 账 #475 §B220: 体里现在三头都在 —— As 子句 / arrayElemTypes_ / 点号链 UDT 成员
        # (resolveUdtMemberElemInfo)。窗口不设上限: 函数体首个 col-0 的 `}` 就是收尾，
        # 从前那个 900 字上限会在加了几行之后就量不到 arrayElemTypes_ (假红)。
        $blk = [regex]::Match($txt, 'Vb6Type\s+CCodeGen::resolveReDimElemType[\s\S]*?\r?\n\}')
        if ($blk.Value -match 'resolveArrayElemType\(\s*typeRef\s*\)' -and $blk.Value -match 'arrayElemTypes_\.find') { $rdBodyBoth++ }
        if ($blk.Value -match 'resolveUdtMemberElemInfo\s*\(') { $rdBodyMember++ }
    }
}
if ($rdDef -ne 1) { $bad += ("A8 resolveReDimElemType 定义 = " + $rdDef + " (要求 1；两处定义 = 第二份答案)") }
if ($rdDecl -ne 1) { $bad += ("A8 resolveReDimElemType 声明(cgen_helpers.inc) = " + $rdDecl + " (要求 1)") }
if ($rdCall -lt 1) { $bad += "A8 cgen_redim.cpp 不再调 resolveReDimElemType（ReDim 的元素档又回到只问 As 子句）" }
if ($rdAsOnly -ne 0) { $bad += ("A8 cgen_redim.cpp 里「声明处只问 As」的老形状 = " + $rdAsOnly + " 处 (要求 0；那正是本刀改掉的那一行)") }
if ($rdDef -eq 1 -and $rdBodyBoth -ne 1) { $bad += "A8 resolveReDimElemType 的体里没同时问 As 子句与 arrayElemTypes_（只剩一头 = 修前的形状）" }
if ($rdDef -eq 1 -and $rdBodyMember -ne 1) { $bad += "A8 resolveReDimElemType 不再问点号链 UDT 成员（成员动态数组的元素档又回落 Variant ⇒ 账 #475 的 16/4 字节步长错位回来）" }

# A20: UDT 成员动态数组的元素档只许一位主人 (账 #475 §B220, 第一百七十九刀)
# 修前: `ReDim t.Items(1 To 3)` 走裸名回退, 分配 vb6_sa_variant, 访问 Pattern B 按
# `VB6_SA_AT(int32_t, t.Items, i)` 4 字节步 —— 16 vs 4。修后成员表那一问收进
# resolveUdtMemberElemInfo, 复杂目标与点号链目标各问一次, 谁也不许自己抄一份 udtMembers 循环。
$memDecl = 0; $memDef = 0; $memAll = 0
foreach ($f in $beFiles) {
    foreach ($line in (Lines $f.FullName)) {
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($f.Name -eq "cgen_helpers.inc" -and $t -match '^bool\s+resolveUdtMemberElemInfo\s*\(') { $memDecl++ }
        if ($f.Name -eq "cgen_redim.cpp" -and $t -match '^bool\s+CCodeGen::resolveUdtMemberElemInfo\s*\(') { $memDef++ }
        if ($t -match '\bresolveUdtMemberElemInfo\s*\(' -and $t -notmatch '^bool\s+(CCodeGen::)?resolveUdtMemberElemInfo') { $memAll++ }
    }
}
if ($memAll -ne 2) { $bad += ("A20 resolveUdtMemberElemInfo 调用点 = " + $memAll + " (恰 2: cgen_expr_array.cpp 点号链 + cgen_redim.cpp 复杂目标转调)") }
if ($memDecl -ne 1) { $bad += ("A20 resolveUdtMemberElemInfo 的声明(cgen_helpers.inc) = " + $memDecl + " (恰好 1；0 = 点号链那一头又问不到成员表)") }
if ($memDef -ne 1)  { $bad += ("A20 resolveUdtMemberElemInfo 的定义(cgen_redim.cpp) = " + $memDef + " (恰好 1；>1 = 第二份成员循环)") }


# A9: 字面量兜底档的名单恰好三条
$saVariant = @()
foreach ($f in $beFiles) {
    $ln = 0
    foreach ($line in (Lines $f.FullName)) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($line.Contains("vb6_sa_variant")) { $saVariant += ($f.Name + ":" + $ln) }
    }
}
if ($saVariant.Count -ne 3) {
    $bad += ("A9 src/backend 非注释里的 vb6_sa_variant 字面量 = " + $saVariant.Count +
             " (恰好 3: 分档表 Variant 档 + default 档 + Variant 槽的 Preserve) -> " + ($saVariant -join " | "))
}
$slotKeep = 0
foreach ($s in $saVariant) { if ($s.StartsWith("cgen_redim.cpp")) { $slotKeep++ } }
if ($slotKeep -ne 1) { $bad += ("A9 Variant 槽那枚刻意写死的档不在 cgen_redim.cpp 里了 = " + $slotKeep + " (要求 1)") }

# A10: 多维描述符装箱的投影方向
$boxCast = 0; $castBox = 0
foreach ($f in $beFiles) {
    foreach ($line in (Lines $f.FullName)) {
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($line.Contains("vb6_VariantFromValue((vb6_SafeArray1D*)")) { $boxCast++ }
        if ($line.Contains("(vb6_SafeArray1D*)vb6_VariantFromValue")) { $castBox++ }
    }
}
if ($boxCast -ne 1) { $bad += ("A10 装箱里那一层投影 = " + $boxCast + " 处 (恰好 1；0 = 又把 ND 描述符裸递给 _Generic，它落 default 的 vb6_VariantObject)") }
if ($castBox -ne 0) { $bad += ("A10 强转盖住装箱的写法 = " + $castBox + " 处 (要求 0；BASE 上这一形实测 cl C2440)") }

# A11: 声明表的供给面与清理面
$tblWrite = 0; $tblFind = 0; $tblErase = 0
foreach ($f in $beFiles) {
    foreach ($line in (Lines $f.FullName)) {
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($line.Contains("arrayElemTypes_[")) { $tblWrite++ }
        if ($line.Contains("arrayElemTypes_.find")) { $tblFind++ }
        if ($line.Contains("arrayElemTypes_.erase")) { $tblErase++ }
    }
}
if ($tblWrite -lt 8) { $bad += ("A11 往 arrayElemTypes_ 写档的站点 = " + $tblWrite + " (下限 8；声明路被摘掉就没人给 ReDim 答档)") }
if ($tblFind -lt 3) { $bad += ("A11 问 arrayElemTypes_ 的站点 = " + $tblFind + " (下限 3：访问侧/类型汇总/ReDim 至少各一处)") }
if ($tblErase -ne 1) { $bad += ("A11 arrayElemTypes_.erase = " + $tblErase + " (恰好 1：过程开始的清表只许一处)") }


# A14/A15: 多维 ReDim 界数组临时量的唯一出口 (账 #305 §B156, 第四十四刀)
$nbDecl = 0; $nbDef = 0; $nbCall = 0; $nbLit = 0; $nbOld = 0; $nbSeq = 0; $nbReset = 0; $fxBounds = 0
foreach ($f in $beFiles) {
    $ln = 0
    foreach ($line in (Lines $f.FullName)) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($f.Name -eq "cgen_helpers.inc" -and $t -match '^std::string\s+nextNDBoundsTemp\s*\(') { $nbDecl++ }
        if ($f.Name -eq "cgen_redim.cpp" -and $t -match '^std::string\s+CCodeGen::nextNDBoundsTemp\s*\(') { $nbDef++ }
        if ($f.Name -eq "cgen_redim.cpp" -and $t.Contains("nextNDBoundsTemp(") -and $t -notmatch '^std::string\s+CCodeGen::') { $nbCall++ }
        if ($t.Contains('"_redim_bounds')) { $nbLit++ }
        if ($t.Contains('std::string boundsVar = "_redim_bounds')) { $nbOld++ }
        if ($t.Contains("ndBoundsSeq_")) { $nbSeq++ }
        if ($t -match '^ndBoundsSeq_\s*=\s*0;') { $nbReset++ }   # 声明行(int ...)不算复位
        if ($t.Contains('"_bounds_"')) { $fxBounds++ }
    }
}
if ($nbDecl -ne 1) { $bad += ("A14 nextNDBoundsTemp 的声明(cgen_helpers.inc) = " + $nbDecl + " (恰好 1；0 = 出口撤了，两条路又各拼一份)") }
if ($nbDef -ne 1)  { $bad += ("A14 nextNDBoundsTemp 的定义(cgen_redim.cpp) = " + $nbDef + " (恰好 1；>1 = 第二份答案)") }
if ($nbCall -ne 2) { $bad += ("A14 两条发码路叫这一枚出口 = " + $nbCall + " (恰好 2；1 = 另一条路回到裸拼，那条就是 C2374 的现场)") }
if ($nbLit -ne 1)  { $bad += ("A14 后端非注释里 _redim_bounds 那枚前缀字面量 = " + $nbLit + " (恰好 1；>1 = 名字格式有了第二处写法)") }
if ($nbOld -ne 0)  { $bad += ("A14 旧的 boundsVar 裸拼形状 = " + $nbOld + " 处 (要求 0；那正是本刀撤掉的两行)") }
if ($nbReset -ne 1){ $bad += ("A15 ndBoundsSeq_ 的复位点 = " + $nbReset + " (恰好 1；0 = 序号跨模块漂，产物就不确定了)") }
if ($nbSeq -lt 3)  { $bad += ("A15 提到 ndBoundsSeq_ 的行 = " + $nbSeq + " (下限 3：声明/复位/使用各一)") }
if ($fxBounds -ne 2){ $bad += ("A15 固定数组那条的 _bounds_ 前缀字面量 = " + $fxBounds + " (恰好 2：两条声明路；把它卷进这枚出口=第二份答案)") }

# A16: 多维 Preserve 的下标算术只许访问侧那一套 (账 #305 §B158, 第四十五刀)
$slotDef = @([regex]::Matches($csrc, 'static int32_t vb6_SaNdSlot\s*\(')).Count
$slotAll = @([regex]::Matches($csrc, 'vb6_SaNdSlot\s*\(')).Count
$oldStr = @([regex]::Matches($csrc, 'oldStrides')).Count
$newStr = @([regex]::Matches($csrc, 'newStrides')).Count
$descIdx = @([regex]::Matches($csrc, 'for\s*\(\s*d\s*=\s*arr->dimCount - 1;\s*d >= 0;\s*d--\)')).Count
if ($slotDef -ne 1) { $bad += ("A16 vb6_SaNdSlot 的定义 = " + $slotDef + " (恰好 1；0 = 两枚 Preserve 又各自写一遍下标算术)") }
if ($slotAll -ne 5) { $bad += ("A16 vb6_SaNdSlot 的定义+调用 = " + $slotAll + " 处 (恰好 5：1 定义 + 两枚 Preserve 各 2 次)") }
if ($oldStr + $newStr -ne 0) { $bad += ("A16 oldStrides/newStrides = " + ($oldStr + $newStr) + " 处 (要求 0：那两份表就是与访问侧相反的第二套算术)") }
if ($descIdx -ne 0) { $bad += ("A16 丢弃旧元素那一段的下标枚举是降序 = " + $descIdx + " 处 (要求 0：枚举序即线性格位序，降序算出的 indices 与 linear 不是同一套对应)") }
# 第一百五十四刀: ND 那位 Preserve 的本体改名 _Tr (多收一枚元素交回回调), 老名字退成转调
    # ⇒ A16 读函数体的那两枚钉跟着搬到 _Tr 上; 老名字与它的 NULL 转调由 A12b 钉住。
foreach ($sigA16 in @('vb6_SafeArrayND\* vb6_SafeArrayReDimPreserveND_Tr\(vb6_safearray_elemtype',
                      'vb6_SafeArrayND\* vb6_SafeArrayReDimPreserveND_Udt\(int32_t elemSize')) {
    $mA16 = [regex]::Match($csrc, $sigA16 + '[\s\S]{0,3000}')
    if (-not $mA16.Success) { $bad += ("A16 找不到这一枚 Preserve 的定义：" + $sigA16); continue }
    $useA16 = @([regex]::Matches($mA16.Value, 'vb6_SaNdSlot\s*\(')).Count
    if ($useA16 -ne 2) { $bad += ("A16 这一枚 Preserve 里调用唯一出口 = " + $useA16 + " 处 (恰好 2：旧格位与新格位各一次)") }
}

# A12: 多维 Preserve 这一枚自己收档 (账 #305 §B155 格④, 第四十三刀)
$ndPresDeclH = @([regex]::Matches($h, 'vb6_SafeArrayReDimPreserveND\s*\(\s*vb6_safearray_elemtype\s+elemType')).Count
$ndPresDefC = @([regex]::Matches($csrc, 'vb6_SafeArrayND\*\s+vb6_SafeArrayReDimPreserveND\s*\(\s*vb6_safearray_elemtype\s+elemType')).Count
if ($ndPresDeclH -ne 1) { $bad += ("A12 vb6_SafeArrayReDimPreserveND 的声明收 elemType = " + $ndPresDeclH + " 处 (恰好 1；0 = 形参表又回到没有档，NULL 描述符只能猜)") }
if ($ndPresDefC -ne 1) { $bad += ("A12 vb6_SafeArrayReDimPreserveND 的定义收 elemType = " + $ndPresDefC + " 处 (恰好 1；头/体个数不一致在 runner 上是 cl C2197)") }
# A12b: 第一百五十四刀 —— 两个名字都要钉住 (本体 _Tr 收档又收回调; 老名字只剩一句转调)
$ndTrDeclH = @([regex]::Matches($h, 'vb6_SafeArrayReDimPreserveND_Tr\s*\(\s*vb6_safearray_elemtype\s+elemType')).Count
$ndTrDefC = @([regex]::Matches($csrc, 'vb6_SafeArrayND\*\s+vb6_SafeArrayReDimPreserveND_Tr\s*\(\s*vb6_safearray_elemtype\s+elemType')).Count
$ndTrNull = @([regex]::Matches($csrc, 'vb6_SafeArrayReDimPreserveND_Tr\s*\(elemType,\s*arr,\s*dimCount,\s*newBounds,\s*NULL\)')).Count
if ($ndTrDeclH -ne 1) { $bad += ("A12b _Tr 的声明收 elemType = " + $ndTrDeclH + " (恰好 1; 0 = 兄弟出口在头上没有声明)") }
if ($ndTrDefC -ne 1) { $bad += ("A12b _Tr 的定义收 elemType = " + $ndTrDefC + " (恰好 1; 头/体不一致在 runner 上是 cl C2197)") }
if ($ndTrNull -ne 1) { $bad += ("A12b 老名字转调传 NULL = " + $ndTrNull + " (恰好 1; 0 = 转调丢了, 非类档那一发也开始交回元素)") }

$ndPresHead = [regex]::Match($csrc, 'vb6_SafeArrayND\*\s+vb6_SafeArrayReDimPreserveND_Tr\s*\(\s*vb6_safearray_elemtype[\s\S]{0,900}')
if ($ndPresHead.Value -notmatch 'vb6_SafeArrayReDimND\(\s*elemType') { $bad += "A12 首次分配那一条不再按调用方给的档建（回到 empty = 尺寸表 default 4 字节）" }
if ($ndPresHead.Value -match 'vb6_sa_empty') { $bad += "A12 多维 Preserve 里又出现 vb6_sa_empty 兜底（那正是本刀撤掉的那一行）" }
if ($ndPresHead.Value -notmatch 'elemType == vb6_sa_udt') { $bad += "A12 多维 Preserve 不再挡 udt 档（尺寸表对 udt 答 0 ⇒ calloc(n,0) 是非空零长指针）" }

# A13: 发码侧把档递进去
$ndPresEmit = 0; $ndPresTyped = 0
foreach ($f in $beFiles) {
    foreach ($line in (Lines $f.FullName)) {
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($t.Contains("vb6_SafeArrayReDimPreserveND(")) {
            $ndPresEmit++
            if ($t -match 'vb6_SafeArrayReDimPreserveND\("\s*\+\s*saElemType') { $ndPresTyped++ }
        }
    }
}
if ($ndPresEmit -ne 1) { $bad += ("A13 后端发 vb6_SafeArrayReDimPreserveND 的行 = " + $ndPresEmit + " (恰好 1)") }
if ($ndPresTyped -ne 1) { $bad += ("A13 那一行把 saElemType 递进去的写法 = " + $ndPresTyped + " 处 (恰好 1；0 = RTL 有形参而后端不递)") }

# A17: 「这枚数组是哪一档」只许一处答 (账 #305 §B155 格①②, 第四十六刀)
# 先剥注释再数 —— 这一族的名字太容易被注释里的引用数成一枚站点（§B145 那一刀的哨兵就是被自己
# 写进注释的一句 `switch (v.vt)` 打红的），而这里比的是**名字**，注释里提一次就多一次。
function Strip-Cmt([string]$s) {
    $s = [regex]::Replace($s, '/\*[\s\S]*?\*/', '')
    $s = [regex]::Replace($s, '(?m)//[^\r\n]*', '')
    return $s
}
$vH = Strip-Cmt ([System.IO.File]::ReadAllText($arrH))
$vC = Strip-Cmt ([System.IO.File]::ReadAllText($arrC))
$vCom = Strip-Cmt ([System.IO.File]::ReadAllText((Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_com.c")))
$vConv = Strip-Cmt ([System.IO.File]::ReadAllText((Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_conv.c")))
$vVar = Strip-Cmt ([System.IO.File]::ReadAllText((Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_variant.h")))
$vReg = Strip-Cmt ([System.IO.File]::ReadAllText((Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_registry.c")))

$vtfDecl = @([regex]::Matches($vH, '(?m)^int32_t\s+vb6_SaVtFromElemType\s*\(')).Count
$vtfDef  = @([regex]::Matches($vC, '(?m)^int32_t\s+vb6_SaVtFromElemType\s*\(')).Count
$vtfAll  = @([regex]::Matches($vH + $vC + $vCom, 'vb6_SaVtFromElemType\s*\(')).Count
if ($vtfDecl -ne 1) { $bad += ("A17 vb6_SaVtFromElemType 的声明(array.h) = " + $vtfDecl + " (恰好 1；0 = 表还在但没人能问)") }
if ($vtfDef -ne 1)  { $bad += ("A17 vb6_SaVtFromElemType 的定义(array.c) = " + $vtfDef + " (恰好 1；>1 = 第二张档表)") }
if ($vtfAll - $vtfDecl - $vtfDef -ne 5) {
    $bad += ("A17 问这张表的地方 = " + ($vtfAll - $vtfDecl - $vtfDef) + " 处 (恰好 5：UDT 字节装箱 1 + COM 进口 1 + 认载体出口 2 + 出站白名单 1；少一处 = 那一位又自己折一档。第五十三刀 §B142 把 COM 那**两条进口**并成一枚核，所以 6 -> 5 —— 是少了一张手抄表，不是少了一次问)")
}

$vtPackDecl = @([regex]::Matches($vVar, '(?m)^int32_t\s+vb6_SaVtOfArrayPtr\s*\(')).Count
$vtPackDef  = @([regex]::Matches($vC, '(?m)^int32_t\s+vb6_SaVtOfArrayPtr\s*\(')).Count
$vtPackAll  = @([regex]::Matches($vC + $vVar + $vReg, 'vb6_SaVtOfArrayPtr\s*\(')).Count
if ($vtPackDecl -ne 1) { $bad += ("A17 vb6_SaVtOfArrayPtr 的声明(variant.h) = " + $vtPackDecl + " (恰好 1)") }
if ($vtPackDef -ne 1)  { $bad += ("A17 vb6_SaVtOfArrayPtr 的定义(array.c) = " + $vtPackDef + " (恰好 1)") }
if ($vtPackAll - $vtPackDecl - $vtPackDef -ne 2) {
    $bad += ("A17 手里只有裸指针时问这一枚的地方 = " + ($vtPackAll - $vtPackDecl - $vtPackDef) + " 处 (恰好 2：装箱 + GetAllSettings)")
}

$hardTag = @([regex]::Matches($vH + $vC + $vCom + $vVar + $vReg, 'vb6_vtArray\s*\|\s*vb6_vtVariant')).Count
if ($hardTag -ne 2) { $bad += ("A17 写死的 vb6_vtArray|vb6_vtVariant = " + $hardTag + " 处 (恰好 2：只许在 vb6_SaVtOfArrayPtr 的 NULL 与认不出载体那两条兜底里；多一处 = 有人又开始替数组猜档)") }
$ternTag = @([regex]::Matches($vCom, 'vb6_vtArray\s*\|\s*\(\s*baseVt')).Count
if ($ternTag -ne 0) { $bad += ("A17 旧的三元折档形状 vb6_vtArray|(baseVt …) = " + $ternTag + " 处 (要求 0：它把 I2/R8/R4 进来一律答成 Long)") }

# A17 续（账 #305 §B159 第五十刀）：反向表 vb6_SaElemTypeFromVt 也只许一处答，
#   两条进口的手抄元素 switch 一旦回来（那是同一个问句的第二份、第三份答案），这里报红。
$stiDecl = @([regex]::Matches($vH, '(?m)^vb6_safearray_elemtype\s+vb6_SaElemTypeFromVt\s*\(')).Count
$stiDef  = @([regex]::Matches($vC, '(?m)^vb6_safearray_elemtype\s+vb6_SaElemTypeFromVt\s*\(')).Count
$stiAll  = @([regex]::Matches($vH + $vC + $vCom, 'vb6_SaElemTypeFromVt\s*\(')).Count
if ($stiDecl -ne 1) { $bad += ("A17 vb6_SaElemTypeFromVt 的声明(array.h) = " + $stiDecl + " (恰好 1；0 = 反向表还在但没人能问)") }
if ($stiDef -ne 1)  { $bad += ("A17 vb6_SaElemTypeFromVt 的定义(array.c) = " + $stiDef + " (恰好 1；>1 = 第二张反向表)") }
if ($stiAll - $stiDecl - $stiDef -ne 2) {
    $bad += ("A17 问反向表的地方 = " + ($stiAll - $stiDecl - $stiDef) + " 处 (恰好 2：COM 进口那一枚核 1 + 入站那一臂 (§B162) 1；第五十三刀 §B142 把 com.c 那两条进口并成一枚核，所以 3 -> 2。多一处 = 又有人自己抄了一份元素 switch)")
}
$stiHand = @([regex]::Matches($vCom, 'case\s+VT_BSTR\s*:\s*et\s*=')).Count
if ($stiHand -ne 0) { $bad += ("A17 手抄的元素 switch (case VT_BSTR: et =) = " + $stiHand + " 处 (要求 0：第五十刀把它并进反向表；回来就是两份答案)") }

$tnEmpty = @([regex]::Matches($vConv, 'case\s+vb6_vtEmpty\s*:')).Count
$tnScalarDef = @([regex]::Matches($vConv, '(?m)^static const wchar_t\*\s+vb6_TypeNameScalar\s*\(')).Count
$tnScalarAll = @([regex]::Matches($vConv, 'vb6_TypeNameScalar\s*\(')).Count
if ($tnEmpty -ne 1) { $bad += ("A17 TypeName 的标量名单里 case vb6_vtEmpty: = " + $tnEmpty + " 处 (恰好 1；2 = 名单被抄了两份，改一档就会漏另一档)") }
if ($tnScalarDef -ne 1) { $bad += ("A17 vb6_TypeNameScalar 的定义 = " + $tnScalarDef + " (恰好 1)") }
if ($tnScalarAll - $tnScalarDef -ne 2) {
    $bad += ("A17 问那张名单的地方 = " + ($tnScalarAll - $tnScalarDef) + " 处 (恰好 2：数组那一支的元素名 + 标量那一条；数组分支被摘掉就只剩 1)")
}
# A19（账 #305 §B159，第五十刀）：出站白名单与反向表必须对得上，真数组只许一处造
$blkOut = $vCom.Substring($vCom.IndexOf('static VARTYPE vb6_ComSaVtBlittable'))
$blkOut = $blkOut.Substring(0, $blkOut.IndexOf('' + [char]10 + '}'))
$invTxt = $vC.Substring($vC.IndexOf('vb6_safearray_elemtype vb6_SaElemTypeFromVt'))
$invTxt = $invTxt.Substring(0, $invTxt.IndexOf('' + [char]10 + '}'))
$whitelist = @([regex]::Matches($blkOut, 'case\s+(vb6_sa_\w+)\s*:') | ForEach-Object { $_.Groups[1].Value })
$inverse  = @([regex]::Matches($invTxt, 'return\s+(vb6_sa_\w+)\s*;') | ForEach-Object { $_.Groups[1].Value })
if ($whitelist.Count -lt 6) { $bad += ("A19 出站白名单只数出 " + $whitelist.Count + " 档 (下限 6：bool/byte/int/long/single/double) —— 正则或函数体形状变了，这条规则断电") }
foreach ($w in $whitelist) {
    if ($inverse -notcontains $w) {
        $bad += ("A19 出站白名单批准了 " + $w + "，而反向表 vb6_SaElemTypeFromVt 不认它 —— 写出去是真数组、读回来是别家档（半程往返，Byte 那一档实测答 8204/Variant() 并按 24 字节步长越界读）")
    }
}
$packDef = @([regex]::Matches($vCom, '(?m)^static void\s+vb6_ComPackArrayOut\s*\(')).Count
$packAll = @([regex]::Matches($vCom, 'vb6_ComPackArrayOut\s*\(')).Count
if ($packDef -ne 1) { $bad += ("A19 造真数组那枚出口的定义 = " + $packDef + " (恰好 1)") }
if ($packAll - $packDef -ne 2) { $bad += ("A19 调用它的地方 = " + ($packAll - $packDef) + " 处 (恰好 2：PackBorrow 的数组那一臂 + 数组形参的出站回写；再多 = 第三个人也开始现造真数组)") }
$tagAsk = @([regex]::Matches($vCom, 'if\s*\(\s*v\.vt\s*&\s*vb6_vtArray\s*\)')).Count
if ($tagAsk -ne 1) { $bad += ("A19 PackBorrow 里数组那一问 = " + $tagAsk + " 处 (恰好 1；0 = 数组又落 default 静默交 VT_EMPTY)") }
$inDef = @([regex]::Matches($vCom, '(?m)^static void\s+vb6_VariantArrayIntoSlot\s*\(')).Count
$inAll = @([regex]::Matches($vCom, 'vb6_VariantArrayIntoSlot\s*\(')).Count
if ($inDef -ne 1) { $bad += ("A19 入站那一臂的定义 = " + $inDef + " (恰好 1)") }
if ($inAll - $inDef -ne 1) { $bad += ("A19 入站那一臂被调的地方 = " + ($inAll - $inDef) + " 处 (恰好 1：BorrowIntoSlot；多一处 = 第二个人也开始自己折真数组)") }
$relDef = @([regex]::Matches($vCom, '(?m)^void\s+vb6_VariantSlotRelease\s*\(')).Count
$relSrv = @([regex]::Matches((Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6comserver\vb6comserver_obj.c')), 'vb6_VariantSlotRelease\s*\(')).Count
if ($relDef -ne 1) { $bad += ("A19 回收出口的定义 = " + $relDef + " (恰好 1)") }
if ($relSrv -ne 1) { $bad += ("A19 服务端调回收出口 = " + $relSrv + " 处 (恰好 1；0 = 入站那一臂造的描述符每调一次漏一枚)") }
$iAsk = $vCom.IndexOf('if (v.vt & vb6_vtArray)')
$iSw  = $vCom.IndexOf('void vb6_VariantPackBorrow')
$iSw  = $vCom.IndexOf('switch (v.vt)', $iSw)
if ($iAsk -lt 0 -or $iSw -lt 0) { $bad += "A19 锚点找不到（数组那一问或那张 switch 被改形了）" }
elseif ($iAsk -gt $iSw) { $bad += "A19 数组那一问跑到那张 switch 之后了 —— 标签对不上任何 case，它会静默落 default" }
# A18: 数组 -> ByRef Variant 形参的装箱只许一处答, 且必须问在分档 switch 之前 (账 #305 §B161, 第四十九刀)
# $beAll 只剥**行注释**：正则里的 \r\n 必须是字面反斜杠（.NET 自己解释），写成**真换行**会
# 把注释行的下一整行一起吃掉 —— 后端这些文件是 CRLF，那正是上一版 Strip-LineCmt 把站点数
# 打成 0 的原因。块注释更剥不得：后端 .inc 里到处是发码字符串字面量中的 /* */。
$beAll = ''
function Strip-LineCmt([string]$s) { return [regex]::Replace($s, '(?m)//[^\r\n]*', '') }
foreach ($f in (Get-ChildItem -LiteralPath $beDir -Recurse -File | Where-Object { $_.Extension -in '.cpp', '.inc', '.hpp' })) {
    $beAll += (Strip-LineCmt ([System.IO.File]::ReadAllText($f.FullName)))
}
$a18Txt = Strip-LineCmt ([System.IO.File]::ReadAllText((Join-Path $beDir 'detail\expr\cgen_expr_call_arg_emit.inc')))
$a18ComTxt = Strip-LineCmt ([System.IO.File]::ReadAllText((Join-Path $beDir 'cgen_util_com.cpp')))
$a18Decl = @([regex]::Matches($beAll, '(?m)const bool\s+arrArg92y\s*=')).Count
$a18All  = @([regex]::Matches($beAll, 'arrArg92y')).Count
# 第五十刀: 那一问本身搬进了唯一出口 isArrayCarrierArgName —— 判据改成「定义 1 / 声明 1 / 两个消费者」,
#   而「这张表在哪儿被问」(knownArrays_.count(lowerName)) 仍然恰好 1 处, 就住在那条出口里。
$a18Known = @([regex]::Matches($beAll, 'knownArrays_\s*\.\s*count\s*\(\s*lowerName')).Count
$a18AuthDef = @([regex]::Matches($beAll, '(?m)^bool CCodeGen::isArrayCarrierArgName\s*\(')).Count
$a18AuthAll = @([regex]::Matches($beAll, 'isArrayCarrierArgName\s*\(')).Count
$a18AuthDecl = @([regex]::Matches($beAll, '(?m)^\s*bool\s+isArrayCarrierArgName\s*\(')).Count
$a18AsnTxt = Strip-LineCmt ([System.IO.File]::ReadAllText((Join-Path $beDir 'cgen_util_ctrl.cpp')))
$a18AsnQ = $a18AsnTxt.IndexOf('isArrayCarrierArgName(lowerB163, cExpr)')
$a18AsnT = $a18AsnTxt.IndexOf('Vb6Type vtype = inferExprType(static_cast<Expr&>(*valueNode));')
$a18ComQ = $a18ComTxt.IndexOf('isArrayCarrierArgName(lower, id.name)')
$a18ComV = $a18ComTxt.IndexOf('if (knownVariantVars_.count(lower)) return')
$a18Temp  = @([regex]::Matches($beAll, '"_vb6_arrarg_"')).Count
$a18Raw   = @([regex]::Matches($beAll, 'arrRaw92y')).Count
$a18K48   = @([regex]::Matches($beAll, 'argIsArr92y')).Count
$a18Sw    = $a18Txt.IndexOf('switch (argVbType92y)')
$a18DefI  = $a18Txt.IndexOf('const bool arrArg92y')
$a18BoxI  = $a18Txt.IndexOf('if (arrArg92y) {')
if ($a18Sw -lt 0 -or $a18DefI -lt 0 -or $a18BoxI -lt 0) { $bad += ("A18 那三行锚点找不到 (switch=" + $a18Sw + " def=" + $a18DefI + " box=" + $a18BoxI + ") —— 分档那条 switch 或这一问被改形了") }
if ($a18Decl -ne 1) { $bad += ("A18 arrArg92y 的定义 = " + $a18Decl + " (恰好 1；0 = 装箱那一档又没人问了)") }
if ($a18All - $a18Decl -ne 1) {
    $bad += ("A18 arrArg92y 的使用 = " + ($a18All - $a18Decl) + " 处 (恰好 1：那条 if；多一处 = 判据又搬回 switch 里面，或抄了第二份)")
}
if ($a18Known -ne 1) { $bad += ("A18 出口里问数组表 knownArrays_ 的地方 = " + $a18Known + " 处 (恰好 1，且必须住在 isArrayCarrierArgName 里；0 = 判据换成了别的表，2 = 又抄了一份)") }
if ($a18AuthDef -ne 1) { $bad += ("A18 那条唯一出口的定义 = " + $a18AuthDef + " (恰好 1)") }
if ($a18AuthAll - $a18AuthDef - $a18AuthDecl -ne 4) { $bad += ("A18 问那条出口的地方 = " + ($a18AuthAll - $a18AuthDef - $a18AuthDecl) + " 处 (恰好 4：ByRef Variant 实参 + COM 实参打包 + 第二百三十九刀的 comByRefArgExpr（数组载体不许取址）；少一处 = 有一个消费者又开始自己答)") }
if ($a18ComQ -lt 0 -or $a18ComV -lt 0) { $bad += "A18 COM 打包那一格的锚点找不到（数组那一问或 knownVariantVars_ 那条链被改形了）" }
if ($a18AsnQ -lt 0 -or $a18AsnT -lt 0) {
    $bad += ("A18 赋值站（wrapVariantValue）那两行锚点找不到 (asn=" + $a18AsnQ + " tbl=" + $a18AsnT + ") —— 第五十一刀那一问或那张类型表被改形了")
}
elseif ($a18AsnQ -gt $a18AsnT) {
    $bad += ("A18 赋值站的数组那一问排到了 inferExprType 那张表之后 —— 模块级数组与撞内在函数名的数组会先被折成标量档（账 #305 §B163，第五十一刀）")
}
elseif ($a18ComQ -gt $a18ComV) { $bad += "A18 COM 实参的数组那一问排到了已知变量型那串之后 —— 模块级数组会先被 knownLongVars_/knownBstrVars_ 折成标量档（实测交出去是 3/Long、8/String）" }
if ($a18Temp -ne 1) { $bad += ("A18 临时量前缀 _vb6_arrarg_ = " + $a18Temp + " 处 (恰好 1)") }
if ($a18Raw -ne 2) { $bad += ("A18 arrRaw92y (装箱吃的那枚原样名) = " + $a18Raw + " 处 (恰好 2：定义 + 装箱表达式；1 = 装箱又去吃分档之后的 argVal)") }
if ($a18K48 -ne 0) { $bad += ("A18 第四十八刀那条住在 switch 里的 argIsArr92y = " + $a18K48 + " 处 (要求 0：同一问不许住两处，而且它在 default 里打不到模块级数组)") }
if ($a18Sw -ge 0 -and $a18DefI -ge 0 -and ($a18DefI -gt $a18Sw)) { $bad += "A18 「是不是数组」那一问跑到分档 switch 之后了 —— 模块级数组会先被按元素档分箱（§B161 的现场）" }
if ($a18Sw -ge 0 -and $a18BoxI -ge 0 -and ($a18BoxI -lt $a18Sw)) { $bad += "A18 数组装箱那条跑到分档 switch 之前了 —— 它得覆盖分档那一趟的结果" }

if ($bad.Count -eq 0) {
    Write-Host ("PASS SA element access: macro " + $macroDefs.Count +
                " / inline " + $defs.Count +
                " / cold-path decl " + $decl + " def " + $defn +
                " / backend raw-data lines " + $rawData.Count +
                " / raw-index shape " + $shape + " / ND macros " + $ndMacros.Count + " / ND inline " + $ndDefs.Count + " / ATN emit " + $atnEmit + " / ATN literal " + $atnLiteral + " / redim helper def " + $rdDef + " call " + $rdCall + " / sa_variant lit " + $saVariant.Count + " / box cast " + $boxCast + " / table w " + $tblWrite + " f " + $tblFind + " e " + $tblErase + " / ndPres decl " + $ndPresDeclH + " def " + $ndPresDefC + " emit " + $ndPresEmit + " typed " + $ndPresTyped + " / nb call " + $nbCall + " lit " + $nbLit + " / seq " + $nbSeq + " fix " + $fxBounds + " / slot def " + $slotDef + " all " + $slotAll + " strideTab " + ($oldStr + $newStr) + " desc " + $descIdx +
                " / ndPres_Tr decl " + $ndTrDeclH + " def " + $ndTrDefC + " fwd-NULL " + $ndTrNull +
                " / vt-flavor decl " + $vtfDecl + " def " + $vtfDef + " call " + ($vtfAll - $vtfDecl - $vtfDef) +
                " / pack-of-ptr " + ($vtPackAll - $vtPackDecl - $vtPackDef) + " hard tag " + $hardTag + " tern " + $ternTag +
                " / TypeName list " + $tnEmpty + " scalar " + ($tnScalarAll - $tnScalarDef) +
                " / a19 out-wl " + $whitelist.Count + " inv " + $inverse.Count +
                " pack-call " + ($packAll - $packDef) + " array-ask " + $tagAsk + " in-call " + ($inAll - $inDef) + " rel-srv " + $relSrv +
                " / sti-call " + ($stiAll - $stiDecl - $stiDef) + " hand " + $stiHand +
                " / a18 arr-box decl " + $a18Decl + " uses " + ($a18All - $a18Decl) +
                " known " + $a18Known + " raw " + $a18Raw + " k48 " + $a18K48 +
                " order " + (($a18DefI -lt $a18Sw) -and ($a18BoxI -gt $a18Sw))) -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
