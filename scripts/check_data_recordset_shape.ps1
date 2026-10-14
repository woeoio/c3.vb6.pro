# check_data_recordset_shape.ps1 —— Data 控件 recordset 成员面的发码形状只许有一套答案（账 #278 §B124）
#
# 立账的原委：C29-Data 曾给 `Data1.Recordset.<成员>` 写过一整套"透传形态直译"—— 认 `vb6_Data_Self(`
# 前缀，发 `vb6_Data_Refresh((void*)vb6_hwnd_Data1)` 那种独立出口；而控件属性分支实际交出的是
# `vb6_Data_RecordsetObj(` 的真 IDispatch（成员名表在 vb6forms_memberobj.c:136-137），成员走 vb6_Com*。
# 两套答案同时住在 src/backend 里，而 §B124 的九形探针量到：**那五格前缀判断没有一格被走到**
# （产物里直译调用 0 处）—— 第十五刀撤掉死码，recordset 成员面只剩 memberobj 那一套。
# 这个哨兵钉住撤完之后的形状，防的正是"再给某一枚单独成员补一格直译"—— 那恰是本族长第二份答案的方式。
#
# 判据面：
#   R1 产物：tests/ctrlzero/RsForm.frm 的十形都发成 vb6_Data_RecordsetObj( + vb6_Com* —— 宿主表达式
#      交出的次数恰好 10，逐成员名字面量的条数钉死（Refresh 2 / Move* 各 1 / Fields 2 / BOF 1 /
#      RecordCount 1 / Value 1），且撤掉的那七枚直译出口在产物里必须**还是** 0 处（出现任意一枚就红）。
#   R2 源码：那七枚出口名在 src/backend 的非注释行里再出现成字符串字面量 ⇒ 红；被撤的抠槽助手
#      dataSelfHwndExpr 同名标识符同罪。
#   R3 RTL 那头不许留 phantom 出口（账 #278 §B127）：名单前两枚（Self / FieldValueStr）在 src/rtl
#      去注释后各恰好 0 次 —— 第十五刀撤的是**读者**，这两枚是它们的**被读者**；留一枚没人叫的导出，
#      下一轮就有人把它当成"能用的入口"再补一格直译（§B124 那份清单的成因正是这一对）。
#   R4 防空转：扫到的 .cpp/.inc/.hpp 份数 >= 100、RTL 的 .c/.h 份数 >= 100、产物字节数 >= 4000、
#      两张名单自己的长度也钉死（名单被清空就是把哨兵改成了绿着的空转）。
#   名单外的一切都不碰（不误伤存量活路）：Data 控件自身的属性面 DatabaseName / Connect / RecordSource /
#   BOF / EOF / RecordCount 由 cgen_util_ctrl.cpp 的属性表答（那是 `Data1.BOF` 那一形，不是 recordset 成员），
#   创建期的 vb6_Data_Init / vb6_Data_Bind / vb6_Data_SetRepositionHandler 同理。
#
# 出口只走 --emit-c（只到发码，不起 cl），所以这条判据是秒级的；负控用 -Root 把整棵树指到副本上跑。

param(
    [string]$Root = ''      # 负控用: 把整棵树指到副本上跑
)
$ErrorActionPreference = 'Stop'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }

$viol = @()

# 第十五刀撤掉的那七枚直译出口（recordset 成员面唯一该有的答案是 memberobj 的真 IDispatch）
$DeadExits = @(
    'vb6_Data_Self',
    'vb6_Data_FieldValueStr',
    'vb6_Data_Refresh',
    'vb6_Data_MoveNext',
    'vb6_Data_MovePrevious',
    'vb6_Data_MoveFirst',
    'vb6_Data_MoveLast'
)
# 夹具十形按成员名在产物里数出的字面量条数（2026-10-10 实测）
$MemberCount = [ordered]@{
    'Refresh'      = 2
    'MoveFirst'    = 1
    'MoveLast'     = 1
    'MoveNext'     = 1
    'MovePrevious' = 1
    'Fields'       = 2
    'BOF'          = 1
    'RecordCount'  = 1
    'Value'        = 1
}
$HostObjWant = 10

if ($DeadExits.Count -ne 7)   { $viol += ('名单自废: DeadExits = ' + $DeadExits.Count + ' 枚 (恰好 7)') }
if ($MemberCount.Count -ne 9) { $viol += ('名单自废: MemberCount = ' + $MemberCount.Count + ' 键 (恰好 9)') }

function Count-Of([string]$Hay, [string]$Needle) {
    $n = 0
    $i = 0
    while ($true) {
        $k = $Hay.IndexOf($Needle, $i)
        if ($k -lt 0) { break }
        $n += 1
        $i = $k + 1
    }
    return $n
}

# ---------- R1: 产物那一头 ----------
$c3 = Join-Path $Root '.build\C3.exe'
$fx = Join-Path $Root 'tests\ctrlzero\RsForm.frm'
$bytes = 0
if (-not (Test-Path -LiteralPath $c3)) {
    $viol += ('EMIT: 没有编译器 ' + $c3 + ' (这一面读的是发码)')
} elseif (-not (Test-Path -LiteralPath $fx)) {
    $viol += 'EMIT: 夹具 tests/ctrlzero/RsForm.frm 不在'
} else {
    $lat = [System.Text.Encoding]::GetEncoding(28591)
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('rsshape_' + [guid]::NewGuid().ToString('N') + '.c')
    $tmpE = [IO.Path]::ChangeExtension($tmp, '.err')
    $pr = Start-Process -FilePath $c3 -ArgumentList @('"' + $fx + '"', '--emit-c') -WorkingDirectory $Root -RedirectStandardOutput $tmp -RedirectStandardError $tmpE -NoNewWindow -Wait -PassThru
    $emit = ''
    if (Test-Path -LiteralPath $tmp) { $emit = $lat.GetString([IO.File]::ReadAllBytes($tmp)) }
    Remove-Item $tmp, $tmpE -Force -ErrorAction SilentlyContinue
    if ($pr.ExitCode -ne 0) { $viol += ('EMIT: 夹具退出码 ' + $pr.ExitCode) }
    $bytes = $emit.Length

    $gotHost = Count-Of $emit 'vb6_Data_RecordsetObj(vb6_hwnd_DataZ'
    if ($gotHost -ne $HostObjWant) {
        $viol += ('SHAPE: recordset 宿主交出 ' + $gotHost + ' 次 (恰好 ' + $HostObjWant + ') ⇒ 夹具的十形或 memberobj 那一档被改了')
    }
    foreach ($nm in $DeadExits) {
        $c = Count-Of $emit ($nm + '(')
        if ($c -ne 0) {
            $viol += ('DEAD-RETURNED: 产物里又出现 ' + $c + ' 枚 ' + $nm + '( ⇒ 有人给单枚成员补了第二套直译')
        }
    }
    foreach ($k in $MemberCount.Keys) {
        $c = Count-Of $emit ('L"' + $k + '"')
        if ($c -ne $MemberCount[$k]) {
            $viol += ('MEMBER: L"' + $k + '" 在产物里 ' + $c + ' 条 (恰好 ' + $MemberCount[$k] + ') ⇒ 这一形少发/多发，或换了通道')
        }
    }
}

# ---------- R2: 源码那一头 ----------
$beDir = Join-Path $Root 'src\backend'
$beFiles = @()
if (-not (Test-Path -LiteralPath $beDir)) {
    $viol += ('SRC: src/backend 不在: ' + $beDir)
} else {
    $beFiles = @(Get-ChildItem -LiteralPath $beDir -Recurse -File -Include *.cpp, *.inc, *.hpp)
    $pat = '"(' + ($DeadExits -join '|') + ')[("]'
    $spell = @($beFiles | Select-String -CaseSensitive -Pattern $pat |
               Where-Object { $_.Line -notmatch '^\s*(//|\*)' })
    foreach ($s in $spell) {
        $viol += ('NAME-COPIED: ' + $s.Path.Substring($Root.Length + 1) + ':' + $s.LineNumber)
    }
    $helper = @($beFiles | Select-String -CaseSensitive -Pattern 'dataSelfHwndExpr' |
                Where-Object { $_.Line -notmatch '^\s*(//|\*)' })
    foreach ($s in $helper) {
        $viol += ('HELPER-RETURNED: ' + $s.Path.Substring($Root.Length + 1) + ':' + $s.LineNumber)
    }
}

# ---------- R3 (账 #278 §B127): RTL 那头不许留 phantom 出口 ----------
# 名单的头两枚是被读者：撤了发码侧的读者之后，RTL 里那两枚定义/声明也一起撤了 ⇒ 这里钉住"不许回来"。
$Phantom = @('vb6_Data_Self', 'vb6_Data_FieldValueStr')
$rtlDir = Join-Path $Root 'src\rtl'
$rtlFiles = @()
if (-not (Test-Path -LiteralPath $rtlDir)) {
    $viol += ('RTL: src/rtl 不在: ' + $rtlDir)
} else {
    $rtlFiles = @(Get-ChildItem -LiteralPath $rtlDir -Recurse -File -Include *.c, *.h)
    foreach ($f in $rtlFiles) {
        $ln = 0
        foreach ($raw in [IO.File]::ReadAllLines($f.FullName)) {
            $ln += 1
            $body = $raw
            $cut = $body.IndexOf('//')
            if ($cut -ge 0) { $body = $body.Substring(0, $cut) }
            # 行内块注释整段去掉；剩下的 /* 是跨行块的开头 ⇒ 从它切（同行代码在块里，不算调用）
            $body = [regex]::Replace($body, '/\*.*?\*/', '')
            $bx = $body.IndexOf('/*')
            if ($bx -ge 0) { $body = $body.Substring(0, $bx) }
            foreach ($nm in $Phantom) {
                if ($body -match ('\b' + $nm + '\b')) {
                    $viol += ('RTL-PHANTOM: ' + $f.FullName.Substring($Root.Length + 1) + ':' + $ln +
                              ' 又出现 ' + $nm + ' ⇒ 一枚没人叫的出口回来了')
                }
            }
        }
    }
}

# ---------- R4: 防空转 ----------
if ($beFiles.Count -lt 100) { $viol += ('SCANNED-FILES=' + $beFiles.Count + ' (<100) ⇒ 路径或通配被改坏') }
if ($rtlFiles.Count -lt 100) { $viol += ('RTL-FILES=' + $rtlFiles.Count + ' (<100) ⇒ RTL 那一户没扫到') }
if ($bytes -gt 0 -and $bytes -lt 4000) { $viol += ('EMIT-BYTES=' + $bytes + ' (<4000) ⇒ 夹具产物空了') }

Write-Host ('backend_files=' + $beFiles.Count + ' rtl_files=' + $rtlFiles.Count +
            ' emit_bytes=' + $bytes + ' dead_exits=' + $DeadExits.Count +
            ' member_keys=' + $MemberCount.Count)
if ($viol.Count -gt 0) {
    Write-Host ('FAIL: recordset 发码形状不齐 {0} 处' -f $viol.Count)
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    exit 1
}
Write-Host 'PASS: recordset 成员面只有一套答案 (memberobj 真 IDispatch, RTL 无 phantom 出口)'
exit 0
