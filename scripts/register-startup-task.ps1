# Registers a Windows Scheduled Task that starts the cvetopt server hidden,
# at logon of the given Windows account, so it keeps running in the background
# even if nobody has cvetopt.exe open. This is what makes config.yaml's
# schedule: (auto-run of the report buttons) fire unattended.
#
# Requires:
# - Run as Administrator on the Windows server.
# - The account in -UserId must be able to log on and run Excel COM (interactive
#   session - do NOT use SYSTEM, see README_WIN.md section 5).
# - For the server to survive an unattended reboot, that account also needs
#   Sysinternals Autologon configured (README_WIN.md section 5) - otherwise this
#   task only fires once someone actually logs into that account (RDP etc).
#
# Run: powershell -ExecutionPolicy Bypass -File scripts\register-startup-task.ps1 -UserId "SERVER\invoice"
param(
    [Parameter(Mandatory = $true)]
    [string]$UserId,
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [string]$TaskName = "cvetopt-autostart"
)

$ErrorActionPreference = "Stop"

$vbs = Join-Path $ProjectRoot "cvetopt-hidden.vbs"
if (-not (Test-Path $vbs)) {
    Write-Error "Missing $vbs"
}

$action = New-ScheduledTaskAction `
    -Execute "$env:SystemRoot\System32\wscript.exe" `
    -Argument "`"$vbs`"" `
    -WorkingDirectory $ProjectRoot
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $UserId
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel Highest

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null

Write-Host "OK: task '$TaskName' - starts the cvetopt server (hidden, no browser) at logon of $UserId."
Write-Host "Server keeps running in the background; cvetopt.exe just attaches to it and"
Write-Host "won't stop it when its window is closed (only stops a server it started itself)."
Write-Host ""
Write-Host "For this to also survive an unattended reboot, configure Autologon for $UserId"
Write-Host "(README_WIN.md, section 5)."
Write-Host "To stop the background server manually: cvetopt-stop.bat"
