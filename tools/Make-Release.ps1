#Requires -Version 5.1
<#
.SYNOPSIS
    Makes the ZIP file of a release (ExchangeMessageTrail-<version>.zip).

.DESCRIPTION
    The script reads the version from the first section of CHANGELOG.md.
    It takes the files from git, not from the working folder. Changes that are not committed do not go into the ZIP file.
    The ZIP file contains only the files that EMT needs to run: the launchers, the scripts, src, gui, Lang, help, LICENSE, README.md and CHANGELOG.md.
    The ZIP file does not contain tests, tools, docs and the git settings.
    All files are in one folder ExchangeMessageTrail-<version>.
    Attach the ZIP file to the GitHub release.

.PARAMETER Ref
    The commit or the tag to use. Default: HEAD.

.PARAMETER OutputFolder
    The folder for the ZIP file. Default: the temporary folder of the user.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\tools\Make-Release.ps1 -Ref v2.1.0
#>
[CmdletBinding()]
param(
    [string] $Ref = 'HEAD',
    [string] $OutputFolder = [IO.Path]::GetTempPath()
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot

# The files and the folders of the ZIP file.
$include = @(
    'EMT.cmd', 'EMT-Classic.cmd',
    'ExchangeMessageTrail.ps1', 'ExchangeMessageTrail.Classic.ps1',
    'src', 'gui', 'Lang', 'help',
    'LICENSE', 'README.md', 'CHANGELOG.md'
)

# The version is the first heading "## X.Y.Z" in CHANGELOG.md of the commit.
$changelog = & git -C $repo show "$($Ref):CHANGELOG.md"
if ($LASTEXITCODE -ne 0) { throw "git cannot read CHANGELOG.md in '$Ref'." }
$heading = $changelog | Where-Object { $_ -match '^##\s+(\d+\.\d+\.\d+)' } | Select-Object -First 1
if (-not $heading) { throw "CHANGELOG.md has no heading '## X.Y.Z'." }
$version = ([regex]::Match($heading, '\d+\.\d+\.\d+')).Value

$name = "ExchangeMessageTrail-$version"
$zip = Join-Path $OutputFolder "$name.zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }

# git archive applies .gitattributes: the .ps1, .cmd and .lng files get CRLF.
& git -C $repo archive --format=zip "--prefix=$name/" -o $zip $Ref -- $include
if ($LASTEXITCODE -ne 0) { throw "git archive failed." }

$count = (& git -C $repo ls-tree -r --name-only $Ref -- $include).Count
$size = [math]::Round((Get-Item -LiteralPath $zip).Length / 1KB)
Write-Host "Version: $version"
Write-Host "Files:   $count"
Write-Host "Size:    $size KB"
Write-Host "ZIP:     $zip"
