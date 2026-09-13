# Registers a Windows Scheduled Task that runs the cvetopt server, so config.yaml's
# schedule: (auto-run of the report buttons) fires without anyone using the app.
#
# Default: starts at logon of -UserId in a real interactive session, and locks the
#   screen right after (no input simulation in this app, so locking is safe).
#   Pair it with Sysinternals Autologon so a reboot needs no human - see README_WIN.md §5.
#   Excel COM is only supported in an interactive session, which is why this is the default.
#
# -Unattended: starts at boot under stored credentials, no logon and no window at all.
#   Excel then runs in session 0, where Microsoft does not support it: a modal Excel
#   dialog (repair/locked file) would hang invisibly and silently stop every later run.
#   Also breaks the "выбрать папку/файл" buttons in Настройки. Use only as a fallback.
#
# Requires Administrator. -UserId must own the project files and be able to run Excel.
#
# Run: powershell -ExecutionPolicy Bypass -File scripts\register-startup-task.ps1 -UserId "SERVER\invoice"
param(
    [Parameter(Mandatory = $true)]
    [string]$UserId,
    [switch]$Unattended,
    [System.Security.SecureString]$Password,
    [switch]$NoLock,
    [string]$ProjectRoot,
    [string]$TaskName = "cvetopt-autostart"
)

$ErrorActionPreference = "Stop"

if (-not $ProjectRoot) {
    $ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$bat = Join-Path $ProjectRoot "cvetopt.bat"
if (-not (Test-Path $bat)) {
    Write-Error "Missing $bat"
}

$lockTaskName = "$TaskName-lock"

# ExecutionTimeLimit 0 - without it Task Scheduler kills the server after 3 days.
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit ([TimeSpan]::Zero)

if ($Unattended) {
    if (-not $Password) {
        $Password = Read-Host "Windows password for $UserId" -AsSecureString
    }
    $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))

    # No interactive desktop here: skip the browser tab and the pause-on-error prompts.
    $action = New-ScheduledTaskAction `
        -Execute "$env:SystemRoot\System32\cmd.exe" `
        -Argument "/c set CVETOPT_HIDDEN=1&& set CVETOPT_NO_BROWSER=1&& `"$bat`"" `
        -WorkingDirectory $ProjectRoot
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $trigger.Delay = "PT1M"
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
        -Settings $settings -User $UserId -Password $plain -RunLevel Highest -Force | Out-Null
    Unregister-ScheduledTask -TaskName $lockTaskName -Confirm:$false -ErrorAction SilentlyContinue

    Write-Host "OK: task '$TaskName' - starts the cvetopt server at boot as $UserId."
    Write-Host "No console window, no logon needed - but Excel now runs in session 0."
    Write-Host "Run one Excel step from the UI to confirm it works, and re-register without"
    Write-Host "-Unattended if anything hangs."
}
else {
    $action = New-ScheduledTaskAction -Execute $bat -WorkingDirectory $ProjectRoot
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $UserId
    $principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel Highest
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
        -Settings $settings -Principal $principal -Force | Out-Null

    Write-Host "OK: task '$TaskName' - starts cvetopt.bat at logon of $UserId."

    if ($NoLock) {
        Unregister-ScheduledTask -TaskName $lockTaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Screen lock task not registered (-NoLock)."
    }
    else {
        $lockAction = New-ScheduledTaskAction `
            -Execute "$env:SystemRoot\System32\rundll32.exe" `
            -Argument "user32.dll,LockWorkStation"
        $lockTrigger = New-ScheduledTaskTrigger -AtLogOn -User $UserId
        $lockTrigger.Delay = "PT2M"
        $lockPrincipal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive
        Register-ScheduledTask -TaskName $lockTaskName -Action $lockAction -Trigger $lockTrigger `
            -Settings $settings -Principal $lockPrincipal -Force | Out-Null
        Write-Host "OK: task '$lockTaskName' - locks the screen 2 min after logon."
        Write-Host "The server keeps running while locked (this app drives Excel over COM only)."
    }

    Write-Host ""
    Write-Host "Next: set up Sysinternals Autologon for $UserId so a reboot needs no human."
    Write-Host "Never LOG OFF that account - the server dies with the session. Disconnect instead."
}

Write-Host ""
Write-Host "cvetopt.exe attaches to this server and won't stop it when its window closes."
Write-Host "To stop the server manually: cvetopt-stop.bat"
