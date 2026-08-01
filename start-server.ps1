$ErrorActionPreference = "Stop"

$Port = 8765
$Root = Join-Path $env:LOCALAPPDATA "OllamaOutlookAddin"
if (-not (Test-Path (Join-Path $Root "manifest.xml"))) {
  $Root = Split-Path -Parent $MyInvocation.MyCommand.Path
}

function Get-LocalhostCertificate {
  $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {
    $_.Subject -eq "CN=localhost" -and $_.FriendlyName -eq "CoOllama Localhost Dev" -and $_.NotAfter -gt (Get-Date)
  } | Sort-Object NotAfter -Descending | Select-Object -First 1

  if (-not $cert) {
    $cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "Cert:\CurrentUser\My" -FriendlyName "CoOllama Localhost Dev" -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(3)
  }

  $trusted = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $cert.Thumbprint } | Select-Object -First 1
  if (-not $trusted) {
    $tmp = Join-Path $env:TEMP "ask-ollama-localhost.cer"
    Export-Certificate -Cert $cert -FilePath $tmp | Out-Null
    Import-Certificate -FilePath $tmp -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null
    Remove-Item $tmp -Force
  }

  return $cert
}

function Get-ContentType($Path) {
  switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
    ".html" { "text/html; charset=utf-8" }
    ".js" { "application/javascript; charset=utf-8" }
    ".css" { "text/css; charset=utf-8" }
    ".png" { "image/png" }
    ".xml" { "application/xml; charset=utf-8" }
    default { "application/octet-stream" }
  }
}

function Send-Response($Stream, [int]$Status, [string]$Reason, [byte[]]$Body, [string]$ContentType) {
  $header = "HTTP/1.1 $Status $Reason`r`nContent-Length: $($Body.Length)`r`nContent-Type: $ContentType`r`nCache-Control: no-store, no-cache, must-revalidate, max-age=0`r`nPragma: no-cache`r`nExpires: 0`r`nConnection: close`r`n`r`n"
  $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
  $Stream.Write($headerBytes, 0, $headerBytes.Length)
  if ($Body.Length -gt 0) {
    $Stream.Write($Body, 0, $Body.Length)
  }
}

$cert = Get-LocalhostCertificate
$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $Port)
$listener.Start()

Write-Host "Serving CoOllama from $Root"
Write-Host "Keep this window open while using the Outlook add-in."
Write-Host "URL: https://localhost:$Port/pane/taskpane.html"
Write-Host "Certificate: $($cert.Thumbprint)"
Write-Host ""

try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    $client.ReceiveTimeout = 10000
    $client.SendTimeout = 10000
    try {
      $ssl = [Net.Security.SslStream]::new($client.GetStream(), $false)
      $ssl.ReadTimeout = 10000
      $ssl.WriteTimeout = 10000
      $ssl.AuthenticateAsServer($cert, $false, [Security.Authentication.SslProtocols]::Tls12, $false)

      $reader = [IO.StreamReader]::new($ssl, [Text.Encoding]::ASCII, $false, 1024, $true)
      $requestLine = $reader.ReadLine()
      if (-not $requestLine) { continue }
      while (($line = $reader.ReadLine()) -ne $null -and $line -ne "") {}

      $parts = $requestLine.Split(" ")
      if ($parts.Length -lt 2 -or $parts[0] -ne "GET") {
        Send-Response $ssl 405 "Method Not Allowed" ([Text.Encoding]::UTF8.GetBytes("Method not allowed")) "text/plain; charset=utf-8"
        continue
      }

      $rawPath = [Uri]::UnescapeDataString($parts[1].Split("?")[0])
      if ($rawPath -eq "/") { $rawPath = "/pane/taskpane.html" }
      $relativePath = $rawPath.TrimStart("/").Replace("/", [IO.Path]::DirectorySeparatorChar)
      $fullPath = [IO.Path]::GetFullPath((Join-Path $Root $relativePath))
      $rootFull = [IO.Path]::GetFullPath($Root)

      if (-not $fullPath.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) {
        Send-Response $ssl 403 "Forbidden" ([Text.Encoding]::UTF8.GetBytes("Forbidden")) "text/plain; charset=utf-8"
        continue
      }

      if (-not (Test-Path $fullPath -PathType Leaf)) {
        Send-Response $ssl 404 "Not Found" ([Text.Encoding]::UTF8.GetBytes("Not found")) "text/plain; charset=utf-8"
        continue
      }

      $bytes = [IO.File]::ReadAllBytes($fullPath)
      Send-Response $ssl 200 "OK" $bytes (Get-ContentType $fullPath)
    } catch {
      Write-Host "Request failed: $($_.Exception.Message)"
    } finally {
      if ($reader) { $reader.Dispose() }
      if ($ssl) { $ssl.Dispose() }
      $client.Close()
    }
  }
} finally {
  $listener.Stop()
}





