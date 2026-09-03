$port = 5173
$root = $PSScriptRoot

$listener = [System.Net.Sockets.TcpListener]::new(
    [System.Net.IPAddress]::Loopback,
    $port
)

$listener.Start()

Write-Host "Serving: $root"
Write-Host "URL:     http://127.0.0.1:$port"
Write-Host "Stop:    Ctrl+C"

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()

        try {
            $stream = $client.GetStream()
            $reader = [System.IO.StreamReader]::new($stream)

            $requestLine = $reader.ReadLine()

            # HTTP headers を読み捨てる
            while ($true) {
                $line = $reader.ReadLine()
                if ([string]::IsNullOrEmpty($line)) {
                    break
                }
            }

            # GET /xxx HTTP/1.1 からパスを取得
            $path = "index.html"

            if ($requestLine -match '^GET\s+([^\s]+)') {
                $urlPath = $Matches[1].Split('?')[0]

                if ($urlPath -ne "/") {
                    $path = [Uri]::UnescapeDataString(
                        $urlPath.TrimStart('/')
                    )
                }
            }

            $file = Join-Path $root $path

            if (Test-Path $file -PathType Leaf) {
                $bytes = [System.IO.File]::ReadAllBytes($file)

                $extension = [System.IO.Path]::GetExtension($file).ToLower()

                $contentType = switch ($extension) {
                    ".html" { "text/html; charset=utf-8" }
                    ".js"   { "text/javascript; charset=utf-8" }
                    ".css"  { "text/css; charset=utf-8" }
                    ".json" { "application/json; charset=utf-8" }
                    ".png"  { "image/png" }
                    ".jpg"  { "image/jpeg" }
                    ".jpeg" { "image/jpeg" }
                    ".svg"  { "image/svg+xml" }
                    default { "application/octet-stream" }
                }

                $header =
                    "HTTP/1.1 200 OK`r`n" +
                    "Content-Type: $contentType`r`n" +
                    "Content-Length: $($bytes.Length)`r`n" +
                    "Connection: close`r`n" +
                    "Cache-Control: no-store`r`n" +
                    "`r`n"

                $headerBytes =
                    [System.Text.Encoding]::ASCII.GetBytes($header)

                $stream.Write(
                    $headerBytes,
                    0,
                    $headerBytes.Length
                )

                $stream.Write(
                    $bytes,
                    0,
                    $bytes.Length
                )
            }
            else {
                $body =
                    [System.Text.Encoding]::UTF8.GetBytes(
                        "404 Not Found"
                    )

                $header =
                    "HTTP/1.1 404 Not Found`r`n" +
                    "Content-Type: text/plain; charset=utf-8`r`n" +
                    "Content-Length: $($body.Length)`r`n" +
                    "Connection: close`r`n" +
                    "`r`n"

                $headerBytes =
                    [System.Text.Encoding]::ASCII.GetBytes($header)

                $stream.Write(
                    $headerBytes,
                    0,
                    $headerBytes.Length
                )

                $stream.Write(
                    $body,
                    0,
                    $body.Length
                )
            }
        }
        catch [System.IO.IOException] {
            # ブラウザ/WebView2が途中で接続を切ることがある。
            # PoCでは正常扱いして次の接続を待つ。
        }
        catch {
            Write-Warning $_.Exception.Message
        }
        finally {
            $client.Close()
        }
    }
}
finally {
    $listener.Stop()
}