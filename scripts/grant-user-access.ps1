# Grants Modify rights on cvetopt folder (and Invoice/sklad) to a Windows user.
# Run as Administrator.
#
# First find the exact login:
#   net user
#   Get-LocalUser
#
# Then:
#   powershell -NoProfile -ExecutionPolicy Bypass -File C:\Apps\cvetopt\scripts\grant-user-access.ps1 -User "Ilya" -ProjectRoot "C:\Apps\cvetopt"
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

function Resolve-AccountName([string]$Name) {
    $raw = $Name.Trim()
    if ($raw -match '^\\.\\(.+)$') {
        $raw = $Matches[1]
    }
    if ($raw -match '^[^\\]+\\(.+)$') {
        $candidate = $Matches[1]
    } else {
        $candidate = $raw
    }

    # Prefer local account by name (case-insensitive).
    try {
        $local = Get-LocalUser -ErrorAction Stop | Where-Object {
            $_.Name -ieq $candidate -or $_.FullName -ieq $candidate
        } | Select-Object -First 1
        if ($local) {
            return "$env:COMPUTERNAME\$($local.Name)"
        }
    } catch {
        # Get-LocalUser may be missing on older Server editions.
    }

    # Profile folder under C:\Users
    $profile = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ieq $candidate } |
        Select-Object -First 1
    if ($profile) {
        return "$env:COMPUTERNAME\$($profile.Name)"
    }

    # Last resort: as typed (domain\user etc.)
    if ($Name -match '\\') {
        return $Name
    }
    return "$env:COMPUTERNAME\$candidate"
}

function Grant-Modify([string]$Path, [string]$Account) {
    if (-not (Test-Path $Path)) {
        Write-Warning "Skip (not found): $Path"
        return
    }
    Write-Host "==> icacls $Path -> $Account :(OI)(CI)(M)"
    & icacls $Path /grant "${Account}:(OI)(CI)(M)" /T
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Account not found or icacls failed (exit $LASTEXITCODE)."
        Write-Host "List local users:"
        try {
            Get-LocalUser | ForEach-Object { Write-Host ("  - " + $_.Name + " (" + $_.FullName + ")") }
        } catch {
            Write-Host "  (run: net user)"
        }
        throw "icacls failed for $Path (exit $LASTEXITCODE). Use exact login from the list above."
    }
}

$Account = Resolve-AccountName $User
Write-Host "Project: $ProjectRoot"
Write-Host "User arg: $User"
Write-Host "Resolved: $Account"
Write-Host "Computer: $env:COMPUTERNAME"
Write-Host ""

Grant-Modify $ProjectRoot $Account

foreach ($p in $ExtraPaths) {
    Grant-Modify $p $Account
}

$userName = ($Account -split '\\')[-1]
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
    Write-Warning "User desktop not found: $userDesktop"
    Write-Host "Existing profiles:"
    Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
        ForEach-Object { Write-Host ("  - " + $_.Name) }
}

Write-Host ""
Write-Host "Done. User should start Desktop\cvetopt under their own account."
Write-Host "Do NOT keep cvetopt running under the admin account while the user works."
