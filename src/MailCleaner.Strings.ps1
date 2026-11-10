<#
    MailCleaner.Strings.ps1 — user interface languages.

    The texts are in the folder "Lang" next to "src": one file <code>.lng for each language
    (Lang\ru.lng, Lang\en.lng, ...). To add a language, put a new .lng file in this folder.
    No code change is necessary. The format is in docs/TRANSLATING.md.

    Use:  T 'key'            -> text in the current language
          T 'key' $a $b      -> text with {0}, {1} replaced (the -f operator)

    The current language is $global:McSync.Lang (the window and the background runspace
    share it), or $global:McLang, or 'en'.
    A key that is not in the current language comes from English, then from Russian.
    A key that is not in any file returns the key itself.
#>

# The folder is found from the location of this file. A global variable: the core and the
# runspace load this file in different scopes.
$global:McLangDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'Lang'

function Read-McLanguageFile {
    <#
        Reads one .lng file. Returns Code, Name, Author, File, Texts (key -> text) and Problems.
        Line types:  # comment  |  ; comment  |  empty  |  @meta = value  |  key = text
    #>
    param([Parameter(Mandatory)] [string] $Path)
    $texts = @{}
    $meta = @{}
    $problems = New-Object System.Collections.Generic.List[string]
    $n = 0
    foreach ($raw in [IO.File]::ReadAllLines($Path, [Text.Encoding]::UTF8)) {
        $n++
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith('#') -or $line.StartsWith(';')) { continue }
        $eq = $line.IndexOf('=')
        if ($eq -lt 1) { $problems.Add("line ${n}: no 'key = text'"); continue }
        $key = $line.Substring(0, $eq).Trim()
        $val = $line.Substring($eq + 1).Trim()
        if ($key.StartsWith('@')) { $meta[$key.Substring(1).ToLowerInvariant()] = $val; continue }
        if ($key -notmatch '^[A-Za-z0-9_.]+$') { $problems.Add("line ${n}: bad key '$key'"); continue }
        if ($texts.ContainsKey($key)) { $problems.Add("line ${n}: duplicate key '$key'") }
        $texts[$key] = $val
    }
    $code = if ($meta.ContainsKey('code') -and $meta['code']) { $meta['code'].ToLowerInvariant() }
            else { [IO.Path]::GetFileNameWithoutExtension($Path).ToLowerInvariant() }
    [pscustomobject]@{
        Code     = $code
        Name     = $(if ($meta.ContainsKey('name') -and $meta['name']) { $meta['name'] } else { $code })
        Author   = $(if ($meta.ContainsKey('author')) { $meta['author'] } else { '' })
        File     = $Path
        Texts    = $texts
        Problems = @($problems)
    }
}

function Get-McStringTable {
    <# All languages from the Lang folder: code -> language object. Read once, then cached. #>
    param([switch] $Reload)
    if (-not $Reload -and (Get-Variable -Name McStringTable -Scope Global -ErrorAction SilentlyContinue) -and $global:McStringTable) {
        return $global:McStringTable
    }
    $table = @{}
    if (Test-Path $global:McLangDir) {
        foreach ($f in Get-ChildItem -Path $global:McLangDir -Filter '*.lng' -File | Sort-Object Name) {
            try {
                $l = Read-McLanguageFile -Path $f.FullName
                $table[$l.Code] = $l
            } catch {
                Write-Verbose "Language file $($f.Name) is not readable: $($_.Exception.Message)"
            }
        }
    }
    $global:McStringTable = $table
    $table
}

function Get-McLanguages {
    <# The available languages for the language list: Code, Name, Author. #>
    $table = Get-McStringTable
    @($table.Values | Sort-Object Name | Select-Object Code, Name, Author)
}

function Get-McLanguage {
    if ((Get-Variable -Name McSync -Scope Global -ErrorAction SilentlyContinue) -and $global:McSync -and
        $global:McSync.ContainsKey('Lang') -and $global:McSync.Lang) { return [string]$global:McSync.Lang }
    if ((Get-Variable -Name McLang -Scope Global -ErrorAction SilentlyContinue) -and $global:McLang) { return [string]$global:McLang }
    'en'
}

function Get-McDefaultLanguage {
    <# The Windows UI language if a file for it exists; Russian for Ukrainian and Belarusian Windows; else English. #>
    $table = Get-McStringTable
    $win = [Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName.ToLowerInvariant()
    if ($table.ContainsKey($win)) { return $win }
    if ($win -in 'uk', 'be' -and $table.ContainsKey('ru')) { return 'ru' }
    'en'
}

function T {
    <# Text for the key in the current language. Extra arguments replace {0}, {1}, ... #>
    param(
        [Parameter(Mandatory, Position = 0)] [string] $Key,
        [Parameter(Position = 1, ValueFromRemainingArguments)] [object[]] $Arguments
    )
    $table = Get-McStringTable
    $text = $null
    foreach ($code in @((Get-McLanguage), 'en', 'ru')) {
        if ($table.ContainsKey($code) -and $table[$code].Texts.ContainsKey($Key)) { $text = $table[$code].Texts[$Key]; break }
    }
    if ($null -eq $text) { return $Key }
    if ($Arguments -and $Arguments.Count) {
        try { return ($text -f $Arguments) } catch { return $text }   # a bad {n} in a translation must not stop the tool
    }
    $text
}

function Test-McLanguageFile {
    <#
        Examines a language file against the reference file (English).
        Returns the problems as text lines. No lines = the file is correct.
        Translators can run it:
          . .\src\MailCleaner.Strings.ps1
          Test-McLanguageFile .\Lang\de.lng
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [string] $ReferencePath = (Join-Path $global:McLangDir 'en.lng'),
        [switch] $AllowMissing
    )
    $l = Read-McLanguageFile -Path $Path
    $ref = Read-McLanguageFile -Path $ReferencePath
    foreach ($p in $l.Problems) { $p }
    if (-not $l.Name -or $l.Name -eq $l.Code) { 'no @name line' }
    foreach ($k in $l.Texts.Keys | Sort-Object) {
        if (-not $ref.Texts.ContainsKey($k)) { "unknown key '$k'"; continue }
        $want = @([regex]::Matches($ref.Texts[$k], '\{(\d+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $have = @([regex]::Matches($l.Texts[$k], '\{(\d+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        if (($want -join ',') -ne ($have -join ',')) { "key '$k': placeholders {$($have -join '},{')} must be {$($want -join '},{')}" }
    }
    if (-not $AllowMissing) {
        foreach ($k in $ref.Texts.Keys | Sort-Object) {
            if (-not $l.Texts.ContainsKey($k)) { "missing key '$k' (English text is used)" }
        }
    }
}
