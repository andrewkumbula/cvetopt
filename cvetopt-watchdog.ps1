# Keeps the cvetopt server alive. Registered by scripts\register-startup-task.ps1 as a
# task that repeats every few minutes (see cvetopt-autostart-watchdog), so a crashed or
# accidentally-closed server restarts on its own - without waiting for the next logon
# or reboot, which the plain "At logon" trigger cannot do.
#
# Does nothing if the server is already up, already starting, or a cvetopt console
# window is already open (even a stuck one - piling up more never helps and once
# silently produced ~30 stuck windows overnight when the underlying start hung).
param(
    [switch]$Hidden,
    [string]$ProjectRoot = $PSScriptRoot
)

function Test-ServerUp {
    try {
        Invoke-RestMethod -Uri "http://127.0.0.1:8000/api/state" -TimeoutSec 3 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

if (Get-Process -Name python -ErrorAction SilentlyContinue) {
    exit 0
}
if (Test-ServerUp) {
    exit 0
}
if (Get-Process -Name cmd -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -like "cvetopt*" }) {
    exit 0
}

if ($Hidden) {
    Start-Process -FilePath "$env:SystemRoot\System32\wscript.exe" `
        -ArgumentList "`"$(Join-Path $ProjectRoot 'cvetopt-hidden.vbs')`"" `
        -WorkingDirectory $ProjectRoot
}
else {
    Start-Process -FilePath (Join-Path $ProjectRoot "cvetopt.bat") -WorkingDirectory $ProjectRoot
}
