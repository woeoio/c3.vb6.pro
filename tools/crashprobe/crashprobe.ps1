# crashprobe.ps1 - crashprobe 的薄包装: 读目标 PE 头的 machine 字段, 自动选对应位宽的
# crashprobe 二进制再调用. 用户只调这一个脚本, 不要手挑 exe (选错位宽 = dump/栈两边都不对).
#
# 用法:
#   powershell -File tools\crashprobe\crashprobe.ps1 -Target .\B5b.exe
#   powershell -File tools\crashprobe\crashprobe.ps1 -Target .\B5b.exe -DumpPath .\B5b.dmp
#   pwsh -File tools\crashprobe\crashprobe.ps1 .\B5b.exe            # 位置参数也认
#
# 位宽规则 (只认这两种, 别的 PE 直接拒绝而不是猜):
#   machine=0x014C (IMAGE_FILE_MACHINE_I386)  -> crashprobe32.exe
#   machine=0x8664 (IMAGE_FILE_MACHINE_AMD64) -> crashprobe.exe
#
# 退出码:
#   0   目标正常退出 (EXIT_CODE=0)
#   2   目标文件不存在
#   3   不是能识别的 PE / PE 头损坏
#   4   缺对应位宽的工具二进制 (先跑 tools\crashprobe\build.bat)
#   其它 = crashprobe 的退出码 = 目标进程退出码 (= 异常码, 如 0xC0000005).
#          **这个非零值本身就是取证结果**, 不要当普通失败处理.
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Target,
    [Parameter(Position = 1)][string]$DumpPath = ""
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

function Fail([string]$message, [int]$code) {
    [Console]::Error.WriteLine('[crashprobe] ' + $message)
    exit $code
}

if (-not (Test-Path -LiteralPath $Target -PathType Leaf)) {
    Fail ('目标不存在: ' + $Target) 2
}
$targetFull = (Resolve-Path -LiteralPath $Target).Path
if (-not $DumpPath) {
    $DumpPath = [System.IO.Path]::ChangeExtension($targetFull, '.dmp')
}

# ---- 读 PE 头: e_lfanew @0x3C, IMAGE_FILE_HEADER.Machine @e_lfanew+4 ----
$machine = -1
try {
    $fs = [System.IO.File]::Open($targetFull, 'Open', 'Read', 'ReadWrite')
    try {
        if ($fs.Length -lt 0x40) { Fail ('不是 PE: 文件太短 (' + $fs.Length + ' 字节)') 3 }
        $br = New-Object System.IO.BinaryReader($fs)
        if ($br.ReadUInt16() -ne 0x5A4D) { Fail '不是 PE: 缺 MZ 头' 3 }   # 'MZ'
        $fs.Position = 0x3C
        $lfanew = $br.ReadInt32()
        if ($lfanew -lt 0x40 -or $lfanew + 24 -gt $fs.Length) { Fail ('PE 头损坏: e_lfanew=0x' + ('{0:X}' -f $lfanew)) 3 }
        $fs.Position = $lfanew
        if ($br.ReadUInt32() -ne 0x00004550) { Fail '不是 PE: 缺 PE\0\0 签名' 3 }
        $machine = $br.ReadUInt16()
    } finally { $fs.Dispose() }
} catch {
    Fail ('读不出目标 PE: ' + $_.Exception.Message) 3
}

switch ($machine) {
    0x014C { $toolName = 'crashprobe32.exe'; $bits = 'x86' }
    0x8664 { $toolName = 'crashprobe.exe';   $bits = 'x64' }
    default {
        Fail ("不认识的 PE machine=0x{0:X4} (只支持 0x014C=x86 / 0x8664=x64)" -f $machine) 3
    }
}

$tool = Join-Path $here $toolName
if (-not (Test-Path -LiteralPath $tool -PathType Leaf)) {
    Fail ('缺对应位宽的工具: ' + $tool + ' (先跑 tools\crashprobe\build.bat)') 4
}

Write-Host ("[crashprobe] target=" + $targetFull + " bits=" + $bits + " tool=" + $toolName)
& $tool $targetFull $DumpPath
exit $LASTEXITCODE
