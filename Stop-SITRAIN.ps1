$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pidFile = Join-Path $projectRoot "sitrain_data\server.pid"
if (-not (Test-Path -LiteralPath $pidFile)) {
    Write-Host "SITRAIN server PID file not found."
    exit 0
}
$serverPid = [int](Get-Content -LiteralPath $pidFile -Raw)
$process = Get-Process -Id $serverPid -ErrorAction SilentlyContinue
if ($process) {
    Stop-Process -Id $serverPid
    Write-Host "SITRAIN server stopped."
}
Remove-Item -LiteralPath $pidFile
