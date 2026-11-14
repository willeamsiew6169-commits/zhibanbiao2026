# Windows WhatsApp sender watchdog. Restart ONLY on explicit safe exit code 20.
$ErrorActionPreference = 'Stop'
$Root = 'C:\gyt_system\zhiban_cloud_system'
$Script = Join-Path $Root 'whatsapp_agent\whatsapp_sender.py'
$Python = 'C:\Users\wille\anaconda3\python.exe'
$LogDir = Join-Path $Root 'whatsapp_agent\watchdog_logs'
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$Log = Join-Path $LogDir 'watchdog.log'
function Log([string]$message) {
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $message
    Write-Host $line
    Add-Content -LiteralPath $Log -Value $line -Encoding UTF8
}
if (-not (Test-Path -LiteralPath $Python)) { Log "STOP: Python not found: $Python"; exit 2 }
if (-not (Test-Path -LiteralPath $Script)) { Log "STOP: Sender not found: $Script"; exit 2 }
Log 'Watchdog started. Only safe browser-failure exit 20 will be restarted.'
while ($true) {
    try {
        & $Python -u $Script
        $code = $LASTEXITCODE
    } catch {
        Log "STOP: Failed to launch sender: $_"
        exit 3
    }
    if ($code -eq 20) {
        Log 'Browser connection failed before claiming a job; restart in 30 seconds.'
        Start-Sleep -Seconds 30
        continue
    }
    Log "STOP: Sender exit code=$code. No automatic restart (safety/manual/unknown exit)."
    exit $code
}
