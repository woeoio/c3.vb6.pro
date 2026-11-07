param(
    [int]$Port = 8765,
    [string]$Root
)
# 账 #417: urlpic_local 的最小 loopback HTTP 服务 (TcpListener, 不需要管理员/URL ACL)。
# 只服务 GET, 一次一连接; 端口被占或绑不上由调用方记 SKIP(asset/local-http), 不许当 PASS。
$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $Port)
$listener.Start()
Write-Output "LISTENING $Port"
while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
        $stream = $client.GetStream()
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::ASCII, $false, 1024, $true)
        $reqLine = $reader.ReadLine()
        while ($true) {
            $h = $reader.ReadLine()
            if ($null -eq $h -or $h -eq '') { break }
        }
        $path = ''
        if ($reqLine) {
            $parts = $reqLine.Split(' ')
            if ($parts.Length -ge 2) { $path = $parts[1].TrimStart('/') }
        }
        $file = Join-Path $Root $path
        if ($path -and (Test-Path $file -PathType Leaf)) {
            $body = [System.IO.File]::ReadAllBytes($file)
            $head = "HTTP/1.1 200 OK`r`nContent-Length: $($body.Length)`r`nContent-Type: image/png`r`nConnection: close`r`n`r`n"
            $hb = [System.Text.Encoding]::ASCII.GetBytes($head)
            $stream.Write($hb, 0, $hb.Length)
            $stream.Write($body, 0, $body.Length)
        } else {
            $head = "HTTP/1.1 404 Not Found`r`nContent-Length: 0`r`nConnection: close`r`n`r`n"
            $hb = [System.Text.Encoding]::ASCII.GetBytes($head)
            $stream.Write($hb, 0, $hb.Length)
        }
        $stream.Flush()
    } catch {
        # 单次请求失败不影响服务
    } finally {
        $client.Close()
    }
}
