# Run once in PowerShell to create a startup shortcut for CURRENT Windows user.
$ErrorActionPreference = 'Stop'
$Root = 'C:\gyt_system\zhiban_cloud_system'
$Watchdog = Join-Path $Root 'whatsapp_agent\启动WhatsApp守护程序.ps1'
if (-not (Test-Path -LiteralPath $Watchdog)) { throw "Missing watchdog script: $Watchdog" }
$startup = [Environment]::GetFolderPath('Startup')
$link = Join-Path $startup '观音堂WhatsApp自动守护.lnk'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($link)
$shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$shortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $Watchdog + '"'
$shortcut.WorkingDirectory = Join-Path $Root 'whatsapp_agent'
$shortcut.Description = '观音堂 WhatsApp 自动通知发送端守护程序（登录 Windows 后启动）'
$shortcut.Save()
Write-Host "Installed startup shortcut: $link"
Write-Host 'It will launch at next Windows sign-in. To start now, run the watchdog script once.'
