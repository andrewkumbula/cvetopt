# Registers a Windows Scheduled Task that runs the cvetopt server, so config.yaml's
# schedule: (auto-run of the report buttons) fires without anyone using the app.
#
# Default (unattended): starts at system boot, no console window, nobody has to log in.
#   Excel COM is not officially supported without an interactive session - if the Excel
#   steps hang or fail, re-register with -Interactive and set up Autologon instead.
#   Side effect: the "выбрать папку/файл" buttons in Настройки stop working (the dialog
#   would open on the invisible session-0 desktop) - type paths by hand there.
#
# -Interactive: starts at logon of -UserId with a visible console window. Needs someone
#   (or Sysinternals Autologon) to actually log into that account. See README_WIN.md §5.
#
# Requires Administrator. -UserId must own the project files and be able to run Excel.
#
# Run: powershell -ExecutionPolicy Bypass -File scripts\register-startup-task.ps1 -UserId "SERVER\invoice"
param(
    [Parameter(Mandatory = $true)]
    [string]$UserId,
    [System.Security.SecureString]$Password,
    [switch]$Interactive,
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

# ExecutionTimeLimit 0 - without it Task Scheduler kills the server after 3 days.
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit ([TimeSpan]::Zero)

if ($Interactive) {
    $action = New-ScheduledTaskAction -Execute $bat -WorkingDirectory $ProjectRoot
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $UserId
    $principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel Highest
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
        -Settings $settings -Principal $principal -Force | Out-Null

    Write-Host "OK: task '$TaskName' - starts cvetopt.bat at logon of $UserId."
    Write-Host "Console window + browser tab open automatically. Do not close that window."
    Write-Host "Needs Autologon for $UserId to survive a reboot (README_WIN.md, section 5)."
}
else {
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

    Write-Host "OK: task '$TaskName' - starts the cvetopt server at boot as $UserId."
    Write-Host "No console window, no logon needed. Test it now: schtasks /Run /TN `"$TaskName`""
    Write-Host ""
    Write-Host "IMPORTANT: run one Excel step from the UI once (e.g. «Шаблон -> копия на"
    Write-Host "сегодняшнюю дату») to confirm Excel works without an interactive session."
    Write-Host "If it hangs: re-run this script with -Interactive and set up Autologon."
}

Write-Host ""
Write-Host "cvetopt.exe attaches to this server and won't stop it when its window closes."
Write-Host "To stop the server manually: cvetopt-stop.bat"
