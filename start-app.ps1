$appUrl = "http://127.0.0.1:4173/"
$apiUrl = "http://127.0.0.1:4174/api/meta/health"
$projectPath = Split-Path -Parent $MyInvocation.MyCommand.Path

$serverReady = $false
$apiReady = $false
try {
  Invoke-WebRequest -Uri $appUrl -UseBasicParsing -TimeoutSec 2 | Out-Null
  $serverReady = $true
} catch {
  $serverReady = $false
}

try {
  Invoke-WebRequest -Uri $apiUrl -UseBasicParsing -TimeoutSec 2 | Out-Null
  $apiReady = $true
} catch {
  $apiReady = $false
}

if (-not $apiReady) {
  Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "start /b node server.js > meta-api.out.log 2> meta-api.err.log" -WorkingDirectory $projectPath

  for ($i = 0; $i -lt 10 -and -not $apiReady; $i++) {
    Start-Sleep -Seconds 1
    try {
      Invoke-WebRequest -Uri $apiUrl -UseBasicParsing -TimeoutSec 2 | Out-Null
      $apiReady = $true
    } catch {
    }
  }
}

if (-not $serverReady) {
  Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "start /b npm.cmd run dev -- --host 127.0.0.1 --port 4173 > vite-dev.out.log 2> vite-dev.err.log" -WorkingDirectory $projectPath

  for ($i = 0; $i -lt 15 -and -not $serverReady; $i++) {
    Start-Sleep -Seconds 1
    try {
      Invoke-WebRequest -Uri $appUrl -UseBasicParsing -TimeoutSec 2 | Out-Null
      $serverReady = $true
    } catch {
    }
  }
}

Start-Process $appUrl
