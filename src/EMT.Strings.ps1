<#
    EMT.Strings.ps1 — user interface languages.

    The texts are in the folder "Lang" next to "src": one file <code>.lng for each language
    (Lang\ru.lng, Lang\en.lng, ...). To add a language, put a new .lng file in this folder.
    No code change is necessary. The format is in docs/TRANSLATING.md.

    Use:  T 'key'            -> text in the current language
          T 'key' $a $b      -> text with {0}, {1} replaced (the -f operator)

    The current language is $global:EmtSync.Lang (the window and the background runspace
    share it), or $global:EmtLang, or 'en'.
    A key that is not in the current language comes from English, then from Russian.
    A key that is not in any file returns the key itself.
#>

# The folder is found from the location of this file. A global variable: the core and the
# runspace load this file in different scopes.
$global:EmtLangDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'Lang'

function Read-EmtLanguageFile {
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

function Get-EmtStringTable {
    <# All languages from the Lang folder: code -> language object. Read once, then cached. #>
    param([switch] $Reload)
    # PSVariable.GetValue, not Get-Variable: T calls this function for each text, and a cmdlet call is slow.
    # GetValue gives $null for a missing variable (no StrictMode error).
    $cached = $ExecutionContext.SessionState.PSVariable.GetValue('global:EmtStringTable')
    if (-not $Reload -and $cached) { return $cached }
    $table = @{}
    if (Test-Path $global:EmtLangDir) {
        foreach ($f in Get-ChildItem -Path $global:EmtLangDir -Filter '*.lng' -File | Sort-Object Name) {
            try {
                $l = Read-EmtLanguageFile -Path $f.FullName
                $table[$l.Code] = $l
            } catch {
                Write-Verbose "Language file $($f.Name) is not readable: $($_.Exception.Message)"
            }
        }
    }
    $global:EmtStringTable = $table
    $table
}

function Get-EmtLanguages {
    <# The available languages for the language list: Code, Name, Author. #>
    $table = Get-EmtStringTable
    @($table.Values | Sort-Object Name | Select-Object Code, Name, Author)
}

function Get-EmtLanguage {
    # PSVariable.GetValue: fast, and $null for a missing variable (see Get-EmtStringTable).
    $vars = $ExecutionContext.SessionState.PSVariable
    $sync = $vars.GetValue('global:EmtSync')
    if ($sync -and $sync.ContainsKey('Lang') -and $sync.Lang) { return [string]$sync.Lang }
    $lang = $vars.GetValue('global:EmtLang')
    if ($lang) { return [string]$lang }
    'en'
}

function Get-EmtDefaultLanguage {
    <# The Windows UI language if a file for it exists; Russian for Ukrainian and Belarusian Windows; else English. #>
    $table = Get-EmtStringTable
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
    $table = Get-EmtStringTable
    $text = $null
    foreach ($code in @((Get-EmtLanguage), 'en', 'ru')) {
        if ($table.ContainsKey($code) -and $table[$code].Texts.ContainsKey($Key)) { $text = $table[$code].Texts[$Key]; break }
    }
    if ($null -eq $text) { return $Key }
    # Not "if ($Arguments)": PowerShell reads @(0) and @('') as false, and the value 0 must replace {0}.
    if ($null -ne $Arguments -and $Arguments.Count -gt 0) {
        try { return ($text -f $Arguments) } catch { return $text }   # a bad {n} in a translation must not stop the tool
    }
    $text
}

function Test-EmtLanguageFile {
    <#
        Examines a language file against the reference file (English).
        Returns the problems as text lines. No lines = the file is correct.
        Translators can run it:
          . .\src\EMT.Strings.ps1
          Test-EmtLanguageFile .\Lang\de.lng
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [string] $ReferencePath = (Join-Path $global:EmtLangDir 'en.lng'),
        [switch] $AllowMissing
    )
    $l = Read-EmtLanguageFile -Path $Path
    $ref = Read-EmtLanguageFile -Path $ReferencePath
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
