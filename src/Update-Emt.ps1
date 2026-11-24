#Requires -Version 5.1
<#
.SYNOPSIS
    Installs a new version of Exchange Message Trail (EMT). EMT starts this script, then closes.

.DESCRIPTION
    EMT copies this script to a temporary folder and starts it: the EMT folder changes during the update.
    The steps:
      1. Wait until the EMT process closes.
      2. Keep a copy of the EMT folder (-Backup). Only the copy of the last update stays.
      3. Copy the new files. The folders src, gui and help are replaced completely.
         The folder Lang gets the new files: a language file of the user stays.
         Other files in the EMT folder (for example, the shortcut) stay.
      4. Unblock the files (the mark "downloaded from the Internet").
      5. Start EMT.
    If step 3 or 4 fails, the script restores the copy and starts the earlier version.
    The result goes to -ResultPath (JSON). At the next start, EMT writes it to the log and to audit.csv.
#>
param(
    [Parameter(Mandatory)] [int] $ProcessId,
    [Parameter(Mandatory)] [string] $Source,
    [Parameter(Mandatory)] [string] $Target,
    [Parameter(Mandatory)] [string] $Backup,
    [Parameter(Mandatory)] [string] $ResultPath,
    [string] $FromVersion = '',
    [string] $ToVersion = '',
    [string] $LogPath = '',
    [int] $WaitSeconds = 60,
    [switch] $NoStart
)

$ErrorActionPreference = 'Stop'

# The folders that the package owns: they are replaced completely.
$owned = @('src', 'gui', 'help')
# Files of earlier versions that are not in the package now.
$obsolete = @('EMT-Classic.cmd', 'ExchangeMessageTrail.Classic.ps1')

function Write-UpdateLog([string] $Text) {
    if (-not $LogPath) { return }
    try { Add-Content -LiteralPath $LogPath -Value ('{0:yyyy-MM-dd HH:mm:ss} {1}' -f (Get-Date), $Text) -Encoding UTF8 } catch { }
}

function Invoke-Robocopy([string] $From, [string] $To, [string] $Mode) {
    <# Robocopy: exit codes 0-7 are success, 8 and more are errors. #>
    if (-not (Test-Path -LiteralPath $From)) { throw "The folder is missing: $From" }
    $out = & robocopy.exe $From $To $Mode /R:2 /W:1 /NFL /NDL /NJH /NJS /NP 2>&1
    if ($LASTEXITCODE -ge 8) { throw "robocopy $Mode '$From' -> '$To': error $LASTEXITCODE. $(($out | Where-Object { $_ }) -join ' ')" }
}

# Backup is empty until the copy is made: then EMT knows that the files did not change.
$result = [ordered]@{
    Result = 'Failed'; From = $FromVersion; To = $ToVersion; Backup = ''; Error = ''
    Time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
}
$changed = $false
$startEmt = -not $NoStart
try {
    Write-UpdateLog "Update $FromVersion -> $ToVersion. Source: $Source. Target: $Target."

    # 1. EMT must be closed: Windows does not replace the files of a running program in all cases.
    $p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($p -and -not $p.WaitForExit($WaitSeconds * 1000)) {
        # EMT still runs: do not start a second EMT.
        $startEmt = $false
        throw "EMT did not close in $WaitSeconds s."
    }

    # 2. The copy of the current version. Earlier copies are removed.
    $backupRoot = Split-Path -Parent $Backup
    if (Test-Path -LiteralPath $backupRoot) {
        Get-ChildItem -LiteralPath $backupRoot -Directory -Filter 'EMT-*' | Remove-Item -Recurse -Force
    }
    Invoke-Robocopy $Target $Backup '/E'
    $result.Backup = $Backup
    Write-UpdateLog "The copy of the earlier version: $Backup"

    # 3. The new files.
    $changed = $true
    foreach ($d in $owned) { Invoke-Robocopy (Join-Path $Source $d) (Join-Path $Target $d) '/MIR' }
    Invoke-Robocopy (Join-Path $Source 'Lang') (Join-Path $Target 'Lang') '/E'
    Get-ChildItem -LiteralPath $Source -File | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $Target -Force }
    foreach ($f in $obsolete) {
        $path = Join-Path $Target $f
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }

    # 4. Unblock.
    Get-ChildItem -LiteralPath $Target -Recurse -File | Unblock-File

    $result.Result = 'OK'
    Write-UpdateLog 'The update is installed.'
} catch {
    $result.Error = $_.Exception.Message
    Write-UpdateLog "ERROR: $($result.Error)"
    if ($changed) {
        try {
            # The folder gets the files of the copy again. New files of the package are removed.
            Invoke-Robocopy $Backup $Target '/MIR'
            $result.Result = 'RolledBack'
            Write-UpdateLog 'The earlier version is restored.'
        } catch {
            $result.Error += " The restore failed: $($_.Exception.Message)"
            Write-UpdateLog "ERROR: the restore failed: $($_.Exception.Message)"
        }
    }
}

try {
    $dir = Split-Path -Parent $ResultPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [pscustomobject]$result | ConvertTo-Json | Set-Content -LiteralPath $ResultPath -Encoding UTF8
} catch { Write-UpdateLog "ERROR: the result file: $($_.Exception.Message)" }

if ($startEmt) {
    $cmd = Join-Path $Target 'EMT.cmd'
    if (Test-Path -LiteralPath $cmd) { Start-Process -FilePath $cmd -WorkingDirectory $Target -WindowStyle Hidden }
}
