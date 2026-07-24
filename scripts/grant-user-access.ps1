# Выдаёт пользователю права на папку cvetopt (и опционально на Invoice/склад).
# Запуск ОТ АДМИНИСТРАТОРА:
#   powershell -ExecutionPolicy Bypass -File scripts\grant-user-access.ps1 -User "SERVER\ИмяПользователя"
#
# Пример:
#   powershell -ExecutionPolicy Bypass -File scripts\grant-user-access.ps1 -User "BananaMan\invoice"
param(
    [Parameter(Mandatory = $true)]
    [string]$User,

    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,

    # Доп. папки, куда Excel/почта пишут файлы (подставьте свои, если отличаются).
    [string[]]$ExtraPaths = @(
        "C:\Invoice",
        "C:\Инвойсы склад"
    )
)

$ErrorActionPreference = "Stop"

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

# Ярлык на рабочий стол пользователя (если профиль уже есть).
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
    if ($shortcut -ne $null) {
        $shortcut.WorkingDirectory = $ProjectRoot
        $shortcut.Description = "cvetopt"
        $shortcut.Save()
        Write-Host "Shortcut: $lnkPath"
    }
} else {
    Write-Warning "User desktop not found: $userDesktop (user never logged in?)"
}

Write-Host ""
Write-Host "Done. Ask the user to log off/on, then double-click their Desktop\cvetopt shortcut."
Write-Host "Do NOT start cvetopt under the admin account while the user works."
