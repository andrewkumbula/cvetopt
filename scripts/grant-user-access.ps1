# Grants Modify rights on cvetopt folder (and Invoice/sklad) to a Windows user.
# Run as Administrator:
#   powershell -ExecutionPolicy Bypass -File C:\Apps\cvetopt\scripts\grant-user-access.ps1 -User ".\Ilya"
param(
    [Parameter(Mandatory = $true)]
    [string]$User,

    [string]$ProjectRoot = "",

    [string[]]$ExtraPaths = @(
        "C:\Invoice",
        "C:\Инвойсы склад"
    )
)

$ErrorActionPreference = "Stop"

if (-not $ProjectRoot) {
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    if (-not $scriptDir) {
        $scriptDir = "C:\Apps\cvetopt\scripts"
    }
    $ProjectRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path
}

function Grant-Modify([string]$Path, [string]$Account) {
    if (-not (Test-Path $Path)) {
        Write-Warning "Skip (not found): $Path"
        return
    }
    Write-Host "==> icacls $Path -> $Account :(OI)(CI)(M)"
    & icacls $Path /grant "${Account}:(OI)(CI)(M)" /T | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "icacls failed for $Path (exit $LASTEXITCODE)"
    }
}

Write-Host "Project: $ProjectRoot"
Write-Host "User:    $User"
Write-Host ""

Grant-Modify $ProjectRoot $User

foreach ($p in $ExtraPaths) {
    Grant-Modify $p $User
}

$userName = ($User -split '\\')[-1]
$userDesktop = "C:\Users\$userName\Desktop"
$exe = Join-Path $ProjectRoot "cvetopt.exe"
$vbs = Join-Path $ProjectRoot "cvetopt-launcher.vbs"
$lnkPath = Join-Path $userDesktop "cvetopt.lnk"

if (Test-Path $userDesktop) {
    $wsh = New-Object -ComObject WScript.Shell
    $shortcut = $wsh.CreateShortcut($lnkPath)
    if (Test-Path $exe) {
        $shortcut.TargetPath = $exe
        $shortcut.Arguments = ""
    } elseif (Test-Path $vbs) {
        $shortcut.TargetPath = "$env:SystemRoot\System32\wscript.exe"
        $shortcut.Arguments = "`"$vbs`""
    } else {
        Write-Warning "No cvetopt.exe / cvetopt-launcher.vbs - shortcut skipped"
        $shortcut = $null
    }
    if ($null -ne $shortcut) {
        $shortcut.WorkingDirectory = $ProjectRoot
        $shortcut.Description = "cvetopt"
        $shortcut.Save()
        Write-Host "Shortcut: $lnkPath"
    }
} else {
    Write-Warning "User desktop not found: $userDesktop (user never logged in?)"
}

Write-Host ""
Write-Host "Done. User should start Desktop\cvetopt shortcut under their own account."
Write-Host "Do NOT keep cvetopt running under the admin account while the user works."
