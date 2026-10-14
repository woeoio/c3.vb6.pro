#查找visual studio的安装位置
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"

if (Test-Path $vswhere) {
    $installPath = & $vswhere -latest -products * -property installationPath

    if ($installPath) {
        Write-Output "Visual Studio 安装路径: $installPath"

        $devenv = Join-Path $installPath "Common7\IDE\devenv.exe"
        if (Test-Path $devenv) {
            Write-Output "devenv.exe: $devenv"
        }
    } else {
        Write-Output "未找到已安装的 Visual Studio。"
    }
} else {
    Write-Output "未找到 vswhere.exe，可能没有安装 Visual Studio Installer。"
}