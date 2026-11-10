<#
    MailCleaner.Core.ps1 — UI-independent logic for MailCleaner.

    Loaded (dot-sourced) by:
      * MailCleaner.ps1 (the GUI) inside its background runspace;
      * tests\MailCleaner.Core.Tests.ps1 (Pester 5).

    Pure helpers (no Exchange calls) come first and are fully unit-tested.
    Exchange-dependent operations follow; they expect the Exchange cmdlets to be
    imported into the session (see Connect-McExchange).
#>

Set-StrictMode -Version 2.0

# User interface texts (ru/en) and the T function.
. (Join-Path $PSScriptRoot 'MailCleaner.Strings.ps1')

function Get-McSearchPrefix { 'MC_' }

# Up to this number of mailboxes, the search uses Search-Mailbox (seconds) instead of Compliance Search (minutes).
function Get-McQuickSearchLimit { 10 }

function Get-McExchangeCommandNames {
    @(
        'Get-TransportService', 'Get-MessageTrackingLog', 'Get-Recipient', 'Get-AcceptedDomain',
        'New-ComplianceSearch', 'Start-ComplianceSearch', 'Stop-ComplianceSearch',
        'Get-ComplianceSearch', 'Remove-ComplianceSearch',
        'New-ComplianceSearchAction', 'Get-ComplianceSearchAction', 'Remove-ComplianceSearchAction',
        'Search-Mailbox', 'Get-MailboxSearch', 'Remove-MailboxSearch'
    )
}

#region ---------- Logging / progress / cancellation (shared with the UI via $global:McSync) ----------

function Write-McLog {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('Info', 'Ok', 'Warn', 'Error')] [string] $Level = 'Info'
    )
    $line = '{0:HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level.ToUpperInvariant().PadRight(5), $Message
    if ((Get-Variable -Name McSync -Scope Global -ErrorAction SilentlyContinue) -and $global:McSync) {
        $global:McSync.Log.Enqueue($line)
    } else {
        Write-Verbose $line
    }
}

function Set-McProgress {
    param([int] $Percent, [string] $Status = '')
    if ((Get-Variable -Name McSync -Scope Global -ErrorAction SilentlyContinue) -and $global:McSync) {
        $global:McSync.Progress = [math]::Max(0, [math]::Min(100, $Percent))
        $global:McSync.Status   = $Status
    }
}

function Test-McCancelRequested {
    [bool]((Get-Variable -Name McSync -Scope Global -ErrorAction SilentlyContinue) -and
           $global:McSync -and $global:McSync.Cancel)
}

function Test-McCancel {
    if (Test-McCancelRequested) {
        throw [System.OperationCanceledException]::new((T 'core.cancelled'))
    }
}

function Wait-McSeconds {
    <# Waits and checks for cancel each second. #>
    param([int] $Seconds)
    for ($i = 0; $i -lt $Seconds; $i++) { Test-McCancel; Start-Sleep -Seconds 1 }
}

#endregion

#region ---------- Pure helpers ----------

function Split-McAddressList {
    <#
        "a@x; b@x, c@x" -> lower-case unique array.
        The separators are the comma, the semicolon and the line break.
        A space separates only full addresses ("a@x.com b@x.com"). A name with a space
        ("Ivanov Ivan") stays one item: the address book search (ANR) needs the full name.
    #>
    param([string] $Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $items = foreach ($part in ($Text -split '[,;\r\n]+')) {
        $p = $part.Trim()
        if (-not $p) { continue }
        $words = @($p -split '\s+' | Where-Object { $_ })
        if ($words.Count -gt 1 -and @($words | Where-Object { Test-McFullAddress $_ }).Count -eq $words.Count) { $words }
        else { $p -replace '\s+', ' ' }
    }
    @($items | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
}

function Test-McFullAddress {
    <# True for a full SMTP address: name@domain.tld. A part of an address or a name is not full. #>
    param([string] $Text)
    [bool]($Text -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Get-McKqlDateRange {
    <#
        KQL filters by whole days only. Returns From (inclusive) and To (exclusive) as yyyy-MM-dd.
        An End exactly at midnight does not pull in the next day.

        Exchange can read KQL dates as local time or as UTC. The range covers both:
        it is the union of the local day range and the UTC day range. The range can grow by
        one day at most. The exact time filter is the message tracking log.
    #>
    param(
        [Parameter(Mandatory)] [datetime] $Start, [Parameter(Mandatory)] [datetime] $End,
        [Nullable[timespan]] $UtcOffset = $null
    )
    if ($End -le $Start) { throw (T 'crit.endBeforeStart') }
    if ($null -eq $UtcOffset) { $UtcOffset = [TimeZoneInfo]::Local.GetUtcOffset($Start) }
    $ceilDay = { param([datetime] $d) if ($d -eq $d.Date) { $d.Date } else { $d.Date.AddDays(1) } }
    $from = @($Start.Date, ($Start - $UtcOffset).Date) | Sort-Object | Select-Object -First 1
    $to   = @((& $ceilDay $End), (& $ceilDay ($End - $UtcOffset))) | Sort-Object | Select-Object -Last 1
    [pscustomobject]@{
        From = $from.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        To   = $to.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    }
}

function Get-McBaseSubject {
    <# Subject without reply/forward prefixes (RE:, FW:, Fwd:, AW:, WG:, Отв:, Пересл: …), lower case. #>
    param([string] $Subject)
    $rx = '^\s*((re|fw|fwd|aw|wg|tr|rv|sv|vs|отв|ответ|пересл|переслано)\s*(\[\d+\])?\s*:\s*)+'
    (([string]$Subject) -replace $rx, '').Trim().ToLowerInvariant()
}

function ConvertTo-McKqlPhrase {
    param([string] $Text)
    '"' + ($Text -replace '"', '').Trim() + '"'
}

function New-McSearchQuery {
    <#
        Builds the KQL for Compliance Search / Search-Mailbox.
        Refuses to build a query that is limited by dates only (it would match everything).
    #>
    param(
        [string[]] $Senders = @(),
        [string]   $Subject,
        [string]   $Attachment,
        [string]   $Keywords,
        [Parameter(Mandatory)] [datetime] $Start,
        [Parameter(Mandatory)] [datetime] $End,
        [Nullable[timespan]] $UtcOffset = $null,
        [string[]] $Recipients = @()
    )
    $parts = New-Object System.Collections.Generic.List[string]
    $senders = @($Senders | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    if ($senders.Count) { $parts.Add('(' + (($senders | ForEach-Object { "from:$_" }) -join ' OR ') + ')') }
    if ($Subject)    { $parts.Add('(subject:' + (ConvertTo-McKqlPhrase $Subject) + ')') }
    if ($Attachment) { $parts.Add('(attachment:' + (ConvertTo-McKqlPhrase $Attachment) + ')') }
    if ($Keywords)   { $parts.Add('(' + (ConvertTo-McKqlPhrase $Keywords) + ')') }
    if ($parts.Count -eq 0) {
        throw (T 'crit.datesOnly')
    }
    # The recipient limits the query: only messages that this address received (To, Cc or Bcc).
    # Without it, "from:x" finds all messages of the sender to all people.
    $rcpts = @($Recipients | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    if ($rcpts.Count) {
        $parts.Add('(' + (($rcpts | ForEach-Object { "to:$_ OR cc:$_ OR bcc:$_" }) -join ' OR ') + ')')
    }
    $range = Get-McKqlDateRange -Start $Start -End $End -UtcOffset $UtcOffset
    $parts.Add("(sent>=$($range.From))")
    $parts.Add("(sent<$($range.To))")
    $parts -join ' AND '
}

function Complete-McCriteria {
    <# Returns a copy of the criteria with all keys present (StrictMode refuses missing hashtable keys). #>
    param([Parameter(Mandatory)] [hashtable] $Criteria)
    $c = @{
        Senders = @(); Subject = ''; Recipient = ''; MessageId = ''; Attachment = ''; Keywords = ''
        Start = [datetime]::MinValue; End = [datetime]::MinValue; IncludeForwards = $true; AllMailboxes = $false
        OnlyRecipient = $false
    }
    foreach ($k in @($Criteria.Keys)) { $c[$k] = $Criteria[$k] }
    $c.Senders = @($c.Senders | Where-Object { $_ })
    $c
}

function Test-McCriteria {
    <# Returns a list of problems; empty list = criteria are usable. #>
    param([Parameter(Mandatory)] [hashtable] $Criteria)
    $Criteria = Complete-McCriteria $Criteria
    $problems = @()
    $hasKey = @($Criteria.Senders).Count -or $Criteria.Subject -or $Criteria.Attachment -or
              $Criteria.Keywords -or $Criteria.MessageId
    if (-not $hasKey) { $problems += (T 'crit.noKey') }
    if ($Criteria.End -le $Criteria.Start) { $problems += (T 'crit.endBeforeStart') }
    if ($Criteria.OnlyRecipient -and -not $Criteria.Recipient) { $problems += (T 'crit.onlyNeedsRecipient') }
    if ($Criteria.OnlyRecipient -and $Criteria.AllMailboxes)   { $problems += (T 'crit.onlyAndAll') }
    if (($Criteria.End - $Criteria.Start).TotalDays -gt 31)    { $problems += (T 'crit.periodTooLong') }
    $problems
}

function Get-McSearchHits {
    <# Parses ComplianceSearch.SuccessResults: "Location: x, Item count: N, Total size: S" (separators , or ;). #>
    param([string] $SuccessResults)
    $rx = 'Location:\s*(?<loc>[^,;}\r\n]+)[,;]\s*Item count:\s*(?<cnt>\d+)[,;]\s*Total size:\s*(?<size>\d+)'
    foreach ($m in [regex]::Matches([string]$SuccessResults, $rx)) {
        [pscustomobject]@{
            Mailbox = $m.Groups['loc'].Value.Trim().ToLowerInvariant()
            Items   = [int]$m.Groups['cnt'].Value
            SizeKB  = [math]::Round([int64]$m.Groups['size'].Value / 1KB, 1)
        }
    }
}

function Get-McSum {
    <#
        Sum of a numeric property. 0 for an empty list.
        Not "Measure-Object -Sum": with an empty list it has no Sum property, and StrictMode stops.
    #>
    param([object[]] $Items = @(), [Parameter(Mandatory)] [string] $Property)
    $sum = 0
    foreach ($i in $Items) { if ($null -ne $i -and $null -ne $i.$Property) { $sum += $i.$Property } }
    $sum
}

function Test-McInternalDomain {
    <#
        True when the address belongs to an accepted domain.
        Accepted domain names can have the form "contoso.com" or "*.contoso.com".
    #>
    param([string] $Address, [string[]] $AcceptedDomains = @())
    if ($Address -notmatch '@(?<d>[^@]+)$') { return $false }
    $d = $Matches['d'].ToLowerInvariant()
    foreach ($a in $AcceptedDomains) {
        $a = ([string]$a).ToLowerInvariant()
        if ($a -eq $d) { return $true }
        if ($a.StartsWith('*.') -and ($d -eq $a.Substring(2) -or $d.EndsWith($a.Substring(1)))) { return $true }
    }
    $false
}

function ConvertFrom-McSizeText {
    <# Search-Mailbox gives the size as text, e.g. "1.25 MB (1,310,720 bytes)". Returns KB. #>
    param([string] $Text)
    if ($Text -match '\((?<b>[\d,\.\s ]+)\s*bytes\)') {
        $digits = $Matches['b'] -replace '[^\d]', ''
        if ($digits) { return [math]::Round([int64]$digits / 1KB, 1) }
    }
    0
}

function Test-McServerGeneratedId {
    <#
        True when a MessageId was generated by an Exchange server itself
        (host part = a transport server name or *.local). Typical for
        forwarding/redirect inbox rules and mailbox forwarding (no copy in Sent Items).
    #>
    param([string] $MessageId, [string[]] $ServerNames = @())
    if ($MessageId -notmatch '@(?<host>[^>]+)>?$') { return $false }
    $msgHost = $Matches['host'].ToLowerInvariant()
    if ($msgHost.EndsWith('.local')) { return $true }
    foreach ($s in $ServerNames) {
        if ($s -and ($msgHost -eq $s.ToLowerInvariant() -or $msgHost.StartsWith($s.ToLowerInvariant() + '.'))) {
            return $true
        }
    }
    $false
}

function Get-McMessageSummary {
    <# One row per MessageId: first submit/receive time, sender, subject, recipient count, server-generated flag. #>
    param([object[]] $Events = @(), [string[]] $ServerNames = @())
    # An event without MessageId (for example NOTIFYMAPI) is not a message: no row for it.
    $Events | Where-Object { $_ -and $_.MessageId } | Group-Object MessageId | ForEach-Object {
        $g      = @($_.Group | Sort-Object Timestamp)
        $first  = ($g | Where-Object { $_.EventId -in 'RECEIVE', 'SUBMIT' } | Select-Object -First 1)
        if (-not $first) { $first = $g[0] }
        $rcpt   = @($g | ForEach-Object { $_.Recipients } | Where-Object { $_ } |
                    ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
        $auto   = Test-McServerGeneratedId -MessageId $_.Name -ServerNames $ServerNames
        [pscustomobject]@{
            Time       = $first.Timestamp
            Sender     = ([string]$first.Sender).ToLowerInvariant()
            Subject    = $first.MessageSubject
            Recipients = $rcpt.Count
            ServerGenerated = [bool]$auto
            Note       = $(if ($auto) { T 'note.serverGenerated' } else { '' })
            MessageId  = $_.Name
        }
    } | Sort-Object Time
}

function Get-McMissingMailboxes {
    <# Locations that were searched but returned no items. #>
    param([string[]] $Locations = @(), [object[]] $Hits = @())
    if ($Locations -contains 'All') { return @() }
    $found = @($Hits | Where-Object { $_.Items -gt 0 } | ForEach-Object { $_.Mailbox.ToLowerInvariant() })
    @($Locations | ForEach-Object { $_.ToLowerInvariant() } | Where-Object { $_ -notin $found } | Sort-Object -Unique)
}

function Get-McLargeDeleteWarning {
    <# Text of the warning for a large delete. Empty text = no warning. #>
    param([int] $Items, [int] $Messages, [int] $MaxItems = 100, [int] $MaxMessages = 5)
    $reasons = @()
    if ($Items -gt $MaxItems)       { $reasons += (T 'large.items' $MaxItems $Items) }
    if ($Messages -gt $MaxMessages) { $reasons += (T 'large.messages' $MaxMessages $Messages) }
    if ($reasons.Count) { T 'large.warning' ($reasons -join ', ') } else { '' }
}

function Test-McSearchName {
    <# True only for the names that MailCleaner makes: MC_yyyyMMdd_HHmmss, with or without "-shadow". #>
    param([string] $Name)
    $Name -cmatch '^MC_\d{8}_\d{6}(-shadow)?$'
}

function Get-McSearchNameTime {
    <# The creation time from a MailCleaner search name (MC_yyyyMMdd_HHmmss[-shadow]), or $null. #>
    param([string] $Name)
    if ($Name -notmatch '^MC_(?<t>\d{8}_\d{6})') { return $null }
    $t = [datetime]::MinValue
    if ([datetime]::TryParseExact($Matches['t'], 'yyyyMMdd_HHmmss', [Globalization.CultureInfo]::InvariantCulture,
                                  [Globalization.DateTimeStyles]::None, [ref]$t)) { return $t }
    $null
}

function New-McSearchName {
    param([datetime] $Now = (Get-Date))
    '{0}{1:yyyyMMdd_HHmmss}' -f (Get-McSearchPrefix), $Now
}

function Write-McAudit {
    <# Appends one row to the CSV audit trail. #>
    param([Parameter(Mandatory)] [string] $Path, [Parameter(Mandatory)] [hashtable] $Record)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $row = [ordered]@{
        Timestamp  = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Operator   = "$env:USERDOMAIN\$env:USERNAME"
        Computer   = $env:COMPUTERNAME
        Server     = ''
        Action     = ''
        SearchName = ''
        Query      = ''
        Mailboxes  = ''
        Items      = ''
        Result     = ''
        Details    = ''
    }
    foreach ($k in @($Record.Keys)) { if ($row.Contains($k)) { $row[$k] = [string]$Record[$k] } }
    [pscustomobject]$row | Export-Csv -Path $Path -Append -NoTypeInformation -Encoding UTF8
}

#endregion

#region ---------- Exchange operations ----------

function Connect-McExchange {
    <# Opens an Exchange remote session and imports the needed cmdlets globally. Returns capabilities. #>
    param([Parameter(Mandatory)] [string] $Server)
    if ((Get-Variable -Name McExSession -Scope Global -ErrorAction SilentlyContinue) -and $global:McExSession) {
        Remove-PSSession $global:McExSession -ErrorAction SilentlyContinue
        $global:McExSession = $null
    }
    Write-McLog (T 'log.connecting' $Server)
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/PowerShell/" `
                   -Authentication Kerberos -ErrorAction Stop
    $module = Import-PSSession $session -CommandName (Get-McExchangeCommandNames) -DisableNameChecking `
                  -AllowClobber -ErrorAction Stop -WarningAction SilentlyContinue
    Import-Module $module -Global -DisableNameChecking -WarningAction SilentlyContinue
    $global:McExSession = $session

    # Accepted domains: an address in another domain is external without a Get-Recipient call.
    $global:McAcceptedDomains = @()
    if (Get-Command Get-AcceptedDomain -ErrorAction SilentlyContinue) {
        try {
            $global:McAcceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop |
                                          ForEach-Object { ([string]$_.DomainName).ToLowerInvariant() })
            Write-McLog (T 'log.domains' ($global:McAcceptedDomains -join ', '))
        } catch {
            Write-McLog (T 'log.domainsFailed' $_.Exception.Message) 'Warn'
        }
    }
    $caps = Get-McCapabilities
    Write-McLog (T 'log.connected' $Server) 'Ok'
    $caps
}

function Get-McAcceptedDomains {
    if ((Get-Variable -Name McAcceptedDomains -Scope Global -ErrorAction SilentlyContinue) -and $global:McAcceptedDomains) {
        return @($global:McAcceptedDomains)
    }
    @()
}

function Get-McCapabilities {
    $has = { param($n) [bool](Get-Command $n -ErrorAction SilentlyContinue) }
    $sm  = Get-Command Search-Mailbox -ErrorAction SilentlyContinue
    [pscustomobject]@{
        Tracking   = (& $has 'Get-MessageTrackingLog') -and (& $has 'Get-TransportService')
        Search     = (& $has 'New-ComplianceSearch') -and (& $has 'Start-ComplianceSearch') -and (& $has 'Get-ComplianceSearch')
        Verify     = [bool]$sm
        SoftDelete = (& $has 'New-ComplianceSearchAction') -and (& $has 'Get-ComplianceSearchAction')
        HardDelete = [bool]($sm -and $sm.Parameters.ContainsKey('DeleteContent'))
    }
}

function Find-McTrackingEvents {
    <#
        Message tracking, tolerant to one wrong filter:
          a) by each sender (no subject), b) by subject (no sender), c) by MessageId, d) by recipient only.
        Keeps messages whose subject contains $Subject (if given).

        All events of one message have the same sender and subject. Thus the sender and subject
        queries already return all events (RECEIVE, EXPAND, DELIVER, SEND) of the message.
        A second pass by MessageId is necessary only with a recipient filter: that filter
        returns only the events of that recipient. Use -CompleteRecipients for the second pass.
    #>
    param(
        [string[]] $Senders = @(), [string] $Subject, [string] $Recipient, [string] $MessageId,
        [Parameter(Mandatory)] [datetime] $Start, [Parameter(Mandatory)] [datetime] $End,
        [switch] $IncludeForwards, [switch] $CompleteRecipients
    )
    $servers = @(Get-TransportService | ForEach-Object { [string]$_.Name })
    Write-McLog (T 'log.trackServers' $servers.Count)
    $common = @{ Start = $Start; End = $End; ResultSize = 'Unlimited' }
    if ($Recipient) { $common.Recipients = $Recipient }
    $trackErrors = New-Object System.Collections.Generic.List[string]

    $queries = @()
    if ($MessageId) { $queries += , @{ MessageId = $MessageId } }
    else {
        foreach ($s in $Senders) { $queries += , @{ Sender = $s } }
        if ($Subject)            { $queries += , @{ MessageSubject = $Subject } }
        if (-not $queries.Count -and $Recipient) { $queries += , @{} }
    }

    $raw = New-Object System.Collections.Generic.List[object]
    $total = [math]::Max(1, $queries.Count * $servers.Count); $n = 0
    foreach ($q in $queries) {
        foreach ($srv in $servers) {
            Test-McCancel
            $n++; Set-McProgress ([int](30 * $n / $total)) (T 'progress.tracking' $srv)
            $p = $common + $q
            $ev = $null
            foreach ($e in @(Get-MessageTrackingLog -Server $srv @p -ErrorAction SilentlyContinue -ErrorVariable ev)) {
                if ($e) { $raw.Add($e) }
            }
            foreach ($x in @($ev)) { if ($x) { $trackErrors.Add("${srv}: $($x.Exception.Message)") } }
        }
    }

    # The subject is text, not a wildcard pattern: "[External] Report" must match itself.
    $pattern = '*' + [System.Management.Automation.WildcardPattern]::Escape([string]$Subject) + '*'
    $hits = @($raw | Where-Object { -not $Subject -or ([string]$_.MessageSubject) -like $pattern })

    if ($Senders.Count) {
        $orig = @($hits | Where-Object { ([string]$_.Sender).ToLowerInvariant() -in $Senders })
        if ($orig.Count) {
            if ($IncludeForwards) {
                # A forward is a message from another sender with the same base subject
                # (subject without RE:/FW: prefixes). A different subject that only contains the text is not a forward.
                $bases = @($orig | ForEach-Object { Get-McBaseSubject $_.MessageSubject } | Sort-Object -Unique)
                $fw = @($hits | Where-Object {
                    ([string]$_.Sender).ToLowerInvariant() -notin $Senders -and (Get-McBaseSubject $_.MessageSubject) -in $bases
                })
                $hits = @($orig) + @($fw)
            } else {
                $hits = $orig
            }
        } elseif ($IncludeForwards -and $Subject) {
            Write-McLog (T 'log.noOriginal') 'Warn'
        } else {
            $hits = @()
        }
    }
    # Events without MessageId (for example NOTIFYMAPI) are not messages. They stay in the event list only.
    $ids   = @($hits | ForEach-Object { [string]$_.MessageId } | Where-Object { $_ } | Sort-Object -Unique)
    $blank = @($hits | Where-Object { -not $_.MessageId })
    Write-McLog (T 'log.trackCount' $ids.Count)

    $source = $raw
    if ($Recipient -and $CompleteRecipients -and $ids.Count) {
        Write-McLog (T 'log.secondPass')
        $source = New-Object System.Collections.Generic.List[object]
        $total = [math]::Max(1, $ids.Count * $servers.Count); $n = 0
        foreach ($id in $ids) {
            foreach ($srv in $servers) {
                Test-McCancel
                $n++; Set-McProgress (30 + [int](5 * $n / $total)) (T 'progress.trackingId' $srv)
                $ev = $null
                foreach ($e in @(Get-MessageTrackingLog -Server $srv -MessageId $id -Start $Start -End $End.AddDays(1) `
                                     -ResultSize Unlimited -ErrorAction SilentlyContinue -ErrorVariable ev)) {
                    if ($e) { $source.Add($e) }
                }
                foreach ($x in @($ev)) { if ($x) { $trackErrors.Add("${srv}: $($x.Exception.Message)") } }
            }
        }
    }
    $errs = @($trackErrors | Sort-Object -Unique)
    foreach ($t in $errs) { Write-McLog (T 'log.trackError' $t) 'Error' }
    $events = @(Select-McUniqueEvents -Events $source -MessageIds $ids) + @(Select-McUniqueEvents -Events $blank -KeepBlank)
    [pscustomobject]@{
        Events  = @($events | Sort-Object Timestamp)
        Servers = $servers
        Errors  = $errs
    }
}

function Select-McUniqueEvents {
    <#
        Keeps the events of the given MessageIds. Removes duplicates (one event can match two queries).
        -KeepBlank also keeps the events without MessageId.
    #>
    param([object[]] $Events = @(), [string[]] $MessageIds = @(), [switch] $KeepBlank)
    $seen = @{}
    $Events | Where-Object { $_ -and (($_.MessageId -and $_.MessageId -in $MessageIds) -or ($KeepBlank -and -not $_.MessageId)) } | ForEach-Object {
        $key = '{0}|{1:o}|{2}|{3}|{4}' -f $_.MessageId, $_.Timestamp, $_.EventId, $_.ServerHostname, (@($_.Recipients) -join ',')
        if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $_ }
    } | Sort-Object Timestamp
}

function Resolve-McRecipients {
    <#
        Classifies addresses: Internal mailbox / External / Group / Other.
        Addresses outside the accepted domains are external without a call.
        The other addresses are resolved in batches of 20 (one Get-Recipient -Filter call per batch).
        If a batch call fails, the batch is resolved address by address.
    #>
    param([string[]] $Addresses = @(), [string[]] $AcceptedDomains = @(), [int] $BatchSize = 20)
    $list = @($Addresses | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $lookup = @($list | Where-Object { -not $AcceptedDomains.Count -or (Test-McInternalDomain $_ $AcceptedDomains) })
    $byAddress = @{}

    for ($b = 0; $b -lt $lookup.Count; $b += $BatchSize) {
        Test-McCancel
        $batch = @($lookup[$b..([math]::Min($b + $BatchSize, $lookup.Count) - 1)])
        Set-McProgress (35 + [int](5 * ($b + $batch.Count) / [math]::Max(1, $lookup.Count))) (T 'progress.recipients' ($b + $batch.Count) $lookup.Count)
        $filter = ($batch | ForEach-Object { "EmailAddresses -eq 'smtp:$($_ -replace "'", "''")'" }) -join ' -or '
        try {
            foreach ($r in @(Get-Recipient -Filter $filter -ResultSize Unlimited -ErrorAction Stop)) {
                $proxies = @($r.EmailAddresses | ForEach-Object { ([string]$_) -replace '^smtp:', '' } |
                             ForEach-Object { $_.ToLowerInvariant() })
                foreach ($a in $batch) { if ($a -in $proxies) { $byAddress[$a] = $r } }
            }
        } catch {
            Write-McLog (T 'log.batchFailed' $_.Exception.Message) 'Warn'
            foreach ($a in $batch) {
                try { $byAddress[$a] = Get-Recipient -Identity $a -ErrorAction Stop } catch { }
            }
        }
    }

    foreach ($a in $list) {
        $r = $byAddress[$a]
        $type = if ($r) { [string]$r.RecipientType } else { '' }
        $kind = if (-not $r) { 'External' }
                elseif ($type -eq 'UserMailbox') { 'Internal' }
                elseif ($type -in 'MailContact', 'MailUser') { 'External' }
                elseif ($type -like '*Group*') { 'Group' }
                else { 'Other' }
        [pscustomobject]@{
            Address = $a
            Kind    = $kind
            Mailbox = $(if ($kind -eq 'Internal') { ([string]$r.PrimarySmtpAddress).ToLowerInvariant() } else { '' })
            Type    = $type
        }
    }
}

function Find-McAddressCandidates {
    <#
        Address book search for a partial address or a name (Get-Recipient -Anr). ANR compares
        the text with the alias, the display name, the first name, the last name and the start
        of the e-mail addresses. Returns Input, Address (primary SMTP), Name, Type.
    #>
    param([Parameter(Mandatory)] [string] $Text, [int] $Max = 25)
    $found = @()
    try {
        $found = @(Get-Recipient -Anr $Text -ResultSize $Max -ErrorAction Stop)
    } catch {
        Write-McLog (T 'log.resolveFailed' $Text $_.Exception.Message) 'Warn'
    }
    @($found | ForEach-Object {
        [pscustomobject]@{
            Input   = $Text
            Address = ([string]$_.PrimarySmtpAddress).ToLowerInvariant()
            Name    = [string]$_.DisplayName
            Type    = $(if ($_.PSObject.Properties['RecipientTypeDetails'] -and $_.RecipientTypeDetails) { [string]$_.RecipientTypeDetails } else { [string]$_.RecipientType })
        }
    } | Where-Object { $_.Address } | Sort-Object Name, Address)
}

function Resolve-McAddressInputs {
    <# For each partial address: Input and Candidates (see Find-McAddressCandidates). #>
    param([Parameter(Mandatory)] [string[]] $Inputs, [int] $Max = 25)
    Write-McLog (T 'log.resolveStart' ($Inputs -join ', '))
    $i = 0
    foreach ($t in $Inputs) {
        Test-McCancel
        $i++; Set-McProgress ([int](100 * $i / $Inputs.Count)) (T 'status.resolve')
        [pscustomobject]@{ Input = $t; Candidates = @(Find-McAddressCandidates -Text $t -Max $Max); Max = $Max }
    }
}

function Wait-McComplianceSearch {
    param([Parameter(Mandatory)] [string] $Name, [int] $TimeoutMinutes = 60, [int] $ProgressFrom = 40, [int] $ProgressTo = 95)
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ($true) {
        Start-Sleep -Seconds 3
        if ((Get-Variable -Name McSync -Scope Global -ErrorAction SilentlyContinue) -and $global:McSync -and $global:McSync.Cancel) {
            Write-McLog (T 'log.csStopping' $Name) 'Warn'
            $stopErr = $null
            Stop-ComplianceSearch -Identity $Name -Confirm:$false -ErrorAction SilentlyContinue -ErrorVariable stopErr
            if ($stopErr) { Write-McLog (T 'log.stopError' $Name (@($stopErr)[0].Exception.Message)) 'Warn' }
            # Exchange on-premises can continue the search after Stop-ComplianceSearch until it completes.
            Write-McLog (T 'log.stopNote' $Name) 'Warn'
            Test-McCancel
        }
        $s = Get-ComplianceSearch -Identity $Name
        $pct = [int]$s.JobProgress
        Set-McProgress ($ProgressFrom + [int](($ProgressTo - $ProgressFrom) * $pct / 100)) (T 'progress.cs' $s.Status $pct)
        if ($s.Status -in 'Completed', 'Stopped', 'Failed') { break }
        if ((Get-Date) -gt $deadline) { throw (T 'err.csTimeout' $Name $TimeoutMinutes) }
    }
    if ($s.Status -ne 'Completed') { throw (T 'err.csStatus' $Name $s.Status) }
    $s
}

function Invoke-McSearch {
    <# Full pipeline: tracking -> recipients -> KQL -> Compliance Search -> hits. #>
    param([Parameter(Mandatory)] [hashtable] $Criteria, [int] $TimeoutMinutes = 60)

    $Criteria = Complete-McCriteria $Criteria
    $problems = @(Test-McCriteria $Criteria)
    if ($problems.Count) { throw ($problems -join ' ') }

    $senders = @($Criteria.Senders)
    $caps = Get-McCapabilities
    if (-not $caps.Search -and -not $caps.Verify) {
        throw (T 'err.noSearchRights')
    }
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $stage = { param($text) Write-McLog (T 'log.stage' $text $timer.Elapsed.TotalSeconds) }

    # ---- 0. Only the recipient mailbox ----
    $onlyMailbox = ''
    if ($Criteria.OnlyRecipient) {
        $r = @(Resolve-McRecipients -Addresses $Criteria.Recipient)
        if (-not $r.Count -or $r[0].Kind -ne 'Internal') {
            throw (T 'err.notInternal' $Criteria.Recipient)
        }
        $onlyMailbox = $r[0].Mailbox
        Write-McLog (T 'log.onlyMode' $onlyMailbox)
    }

    # ---- 1. Message tracking ----
    $tracking = [pscustomobject]@{ Events = @(); Servers = @(); Errors = @() }
    if ($caps.Tracking) {
        $tracking = Find-McTrackingEvents -Senders $senders -Subject $Criteria.Subject -Recipient $Criteria.Recipient `
                        -MessageId $Criteria.MessageId -Start $Criteria.Start -End $Criteria.End `
                        -IncludeForwards:([bool]$Criteria.IncludeForwards) -CompleteRecipients:(-not $onlyMailbox)
        & $stage (T 'stage.tracking')
    } else {
        Write-McLog (T 'log.noTrackRights') 'Warn'
    }
    $events  = @($tracking.Events)
    $trackErrors = @($tracking.Errors)
    $summary = @(Get-McMessageSummary -Events $events -ServerNames $tracking.Servers)
    if (-not $events.Count) { Write-McLog (T 'log.trackEmpty') 'Warn' }
    foreach ($m in @($summary | Where-Object ServerGenerated)) { Write-McLog (T 'log.serverGenerated' $m.MessageId $m.Sender) 'Warn' }

    # ---- 2. Recipients ----
    $addresses = @($events | ForEach-Object { $_.Recipients } | Where-Object { $_ } |
                   ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $trackSenders = @($events | ForEach-Object { [string]$_.Sender } | Where-Object { $_ } |
                      ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $recipients = @(Resolve-McRecipients -Addresses (@($addresses) + @($trackSenders) | Sort-Object -Unique) `
                        -AcceptedDomains (Get-McAcceptedDomains))
    $internal = @($recipients | Where-Object Kind -eq 'Internal' | ForEach-Object Mailbox | Sort-Object -Unique)
    $external = @($recipients | Where-Object { $_.Kind -eq 'External' -and $_.Address -in $addresses })
    if ($external.Count) { Write-McLog (T 'log.external' $external.Count) 'Warn' }
    & $stage (T 'stage.recipients' $recipients.Count)

    # ---- 3. KQL ----
    $kqlSenders = @($senders)
    if (($Criteria.IncludeForwards -or -not $kqlSenders.Count) -and $trackSenders.Count) { $kqlSenders += $trackSenders }
    $kqlSubject = $Criteria.Subject
    if (-not $kqlSubject -and $Criteria.MessageId -and $summary.Count) { $kqlSubject = $summary[0].Subject }
    $kqlRecipients = @()
    if ($Criteria.Recipient) { $kqlRecipients = @($Criteria.Recipient) }
    $query = New-McSearchQuery -Senders $kqlSenders -Subject $kqlSubject -Attachment $Criteria.Attachment `
                 -Keywords $Criteria.Keywords -Start $Criteria.Start -End $Criteria.End -Recipients $kqlRecipients

    $locations = @($internal)
    if ($onlyMailbox) {
        $locations = @($onlyMailbox)
    } elseif ($Criteria.AllMailboxes) {
        $locations = @('All')
    } elseif (-not $locations.Count) {
        # No recipient mailboxes from the tracking log. With a recipient or senders in the criteria,
        # search only their mailboxes: a search of all mailboxes is slow and not necessary.
        $own = @(Resolve-McRecipients -Addresses (@($kqlRecipients) + @($senders)) -AcceptedDomains (Get-McAcceptedDomains) |
                 Where-Object Kind -eq 'Internal' | ForEach-Object Mailbox | Sort-Object -Unique)
        if ($kqlRecipients.Count -and $own.Count) {
            $locations = $own
            Write-McLog (T 'log.recipientScope' ($own -join ', ')) 'Warn'
        } else {
            Write-McLog (T 'log.searchAll') 'Warn'
            $locations = @('All')
        }
    }
    Write-McLog "KQL: $query"

    # ---- 4. Mailbox search ----
    Test-McCancel
    $isAll = $locations[0] -eq 'All'
    $quick = (-not $isAll) -and $caps.Verify -and ($locations.Count -le (Get-McQuickSearchLimit) -or -not $caps.Search)
    if ($isAll -and -not $caps.Search) { throw (T 'err.allNeedsCs') }

    if ($quick) {
        Write-McLog (T 'log.quick' $locations.Count)
        $hits = @(Invoke-McQuickSearch -Mailboxes $locations -Query $query)
        $name      = ''
        $items     = [int](Get-McSum -Items @($hits | Where-Object { $_.Items -gt 0 }) -Property Items)
        $sizeKB    = [double](Get-McSum -Items @($hits | Where-Object { $_.Items -gt 0 }) -Property SizeKB)
        $unindexed = -1
        $errors    = (@($hits | Where-Object { $_.Error } | ForEach-Object { "$($_.Mailbox): $($_.Error)" }) -join '; ')
    } else {
        $name = New-McSearchName
        Write-McLog (T 'log.cs' $name $(if ($isAll) { T 'common.all' } else { $locations.Count }))
        $s = New-McComplianceSearchRun -Name $name -Locations $locations -Query $query -TimeoutMinutes $TimeoutMinutes
        $hits      = @(Get-McSearchHits -SuccessResults ([string]$s.SuccessResults))
        $items     = [int]$s.Items
        $sizeKB    = [math]::Round($s.Size / 1KB, 1)
        $unindexed = [int]$s.UnindexedItems
        $errors    = [string]$s.Errors
    }
    & $stage (T 'stage.search')

    # Incomplete tracking = incomplete mailbox list. Block the delete unless all mailboxes were searched.
    if ($trackErrors.Count -and -not $isAll) {
        $msg = T 'err.trackIncomplete' $trackErrors.Count
        $errors = (@($errors, $msg) | Where-Object { $_ }) -join '; '
    }

    $found   = @($hits | Where-Object { $_.Items -gt 0 } | Sort-Object Mailbox)
    # "Not found in a mailbox" is useful only when the message is in other mailboxes.
    # When no mailbox has it, the message probably does not exist: no list.
    $missing = if ($found.Count) { @(Get-McMissingMailboxes -Locations $locations -Hits $hits) } else { @() }
    Write-McLog (T 'log.found' $items $found.Count $sizeKB) 'Ok'
    if ($errors) { Write-McLog (T 'log.searchErrors' $errors) 'Error' }
    if ($unindexed -gt 0) { Write-McLog (T 'log.unindexed' $unindexed) 'Warn' }
    # The search found items, but the tracking log has no message: nothing confirms that these are the
    # messages the user wants. The window shows this and asks for the typed word before a delete.
    $unconfirmed = ($found.Count -gt 0) -and ($summary.Count -eq 0)
    if ($unconfirmed) { Write-McLog (T 'log.unconfirmed' $items) 'Warn' }
    Set-McProgress 100 (T 'progress.done')

    [pscustomobject]@{
        Method     = $(if ($quick) { 'SearchMailbox' } else { 'ComplianceSearch' })
        SearchName = $name
        Query      = $query
        Locations  = $locations
        Events     = $events
        Messages   = $summary
        Recipients = $recipients
        External   = $external
        Found      = $found
        Missing    = $missing
        Items      = $items
        SizeKB     = [math]::Round($sizeKB, 1)
        Unindexed  = $unindexed
        Errors     = $errors
        TrackingErrors = $trackErrors
        Unconfirmed = [bool]$unconfirmed
        Seconds    = [int]$timer.Elapsed.TotalSeconds
    }
}

function New-McComplianceSearchRun {
    <# Makes, starts and waits for one Compliance Search. Returns the search object. #>
    param(
        [Parameter(Mandatory)] [string] $Name, [Parameter(Mandatory)] [string[]] $Locations,
        [Parameter(Mandatory)] [string] $Query, [int] $TimeoutMinutes = 60
    )
    New-ComplianceSearch -Name $Name -ExchangeLocation $Locations -ContentMatchQuery $Query | Out-Null
    Start-ComplianceSearch -Identity $Name
    Wait-McComplianceSearch -Name $Name -TimeoutMinutes $TimeoutMinutes
}

function Invoke-McQuickSearch {
    <#
        Search-Mailbox -EstimateResultOnly for each mailbox. For a few mailboxes this takes
        seconds, a Compliance Search takes minutes. The count includes Recoverable Items,
        the same as a Compliance Search.
    #>
    param([Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query)
    $i = 0
    foreach ($m in $Mailboxes) {
        Test-McCancel
        $i++; Set-McProgress (40 + [int](55 * $i / $Mailboxes.Count)) (T 'progress.quick' $m)
        try {
            $r = Search-Mailbox -Identity $m -SearchQuery $Query -EstimateResultOnly `
                     -WarningAction SilentlyContinue -ErrorAction Stop
            [pscustomobject]@{
                Mailbox = $m.ToLowerInvariant(); Items = [int]$r.ResultItemsCount
                SizeKB = ConvertFrom-McSizeText ([string]$r.ResultItemsSize); Error = ''
            }
        } catch {
            [pscustomobject]@{ Mailbox = $m.ToLowerInvariant(); Items = 0; SizeKB = 0; Error = $_.Exception.Message }
        }
    }
}

function Invoke-McSoftDelete {
    <# Purge (SoftDelete) of an existing Compliance Search: items go to Recoverable Items. #>
    param([Parameter(Mandatory)] [string] $SearchName, [int] $TimeoutMinutes = 60)
    Write-McLog (T 'log.softStart' $SearchName)
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete -Confirm:$false | Out-Null
    $actionName = "${SearchName}_Purge"
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    do {
        Start-Sleep -Seconds 3
        $a = Get-ComplianceSearchAction -Identity $actionName
        Set-McProgress 50 (T 'progress.soft' $a.Status)
        if ((Get-Date) -gt $deadline) { throw (T 'err.softTimeout' $TimeoutMinutes) }
    } while ($a.Status -notin 'Completed', 'PartiallySucceeded', 'Failed')
    if ($a.Errors) { Write-McLog (T 'log.softErrors' $a.Errors) 'Error' }
    Write-McLog (T 'log.softStatus' $a.Status) $(if ($a.Status -eq 'Completed') { 'Ok' } else { 'Warn' })
    [pscustomobject]@{ Status = [string]$a.Status; Errors = [string]$a.Errors }
}

function Test-McHardDeleteScope {
    <#
        Safety check before a hard delete. Search-Mailbox must not find more items than the search did.
        More items = the query reads differently in Search-Mailbox (for example, the date format)
        or new messages arrived. Returns the mailboxes that fail the check.
    #>
    param([Parameter(Mandatory)] [object[]] $Status, [Parameter(Mandatory)] [hashtable] $Expected)
    foreach ($s in $Status) {
        $exp = if ($Expected.ContainsKey($s.Mailbox)) { [int]$Expected[$s.Mailbox] } else { 0 }
        if ($s.Total -lt 0) {
            [pscustomobject]@{ Mailbox = $s.Mailbox; Expected = $exp; Actual = $s.Total; Reason = (T 'scope.checkFailed' $s.Note) }
        } elseif ($s.Total -gt $exp) {
            [pscustomobject]@{ Mailbox = $s.Mailbox; Expected = $exp; Actual = $s.Total; Reason = (T 'scope.more') }
        }
    }
}

function Invoke-McHardDelete {
    <#
        Permanent delete with Search-Mailbox -DeleteContent (incl. Recoverable Items), mailbox by mailbox.
        On cancel, it stops before the next mailbox and returns the partial result (Cancelled = true).
    #>
    param([Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query)
    $results = New-Object System.Collections.Generic.List[object]
    $cancelled = $false
    $i = 0
    foreach ($m in $Mailboxes) {
        if (Test-McCancelRequested) { $cancelled = $true; break }
        $i++; Set-McProgress ([int](60 * $i / $Mailboxes.Count)) (T 'progress.hard' $m)
        try {
            $r = Search-Mailbox -Identity $m -SearchQuery $Query -SearchDumpster -DeleteContent -Force `
                     -Confirm:$false -WarningAction SilentlyContinue -ErrorAction Stop
            Write-McLog (T 'log.hardDeleted' $m $r.ResultItemsCount) 'Ok'
            $results.Add([pscustomobject]@{ Mailbox = $m; Deleted = [int]$r.ResultItemsCount; Error = '' })
        } catch {
            Write-McLog (T 'log.hardError' $m $_.Exception.Message) 'Error'
            $results.Add([pscustomobject]@{ Mailbox = $m; Deleted = 0; Error = $_.Exception.Message })
        }
    }
    if ($cancelled) { Write-McLog (T 'log.hardStopped' $results.Count $Mailboxes.Count) 'Warn' }
    [pscustomobject]@{ Results = @($results); Cancelled = $cancelled }
}

function Get-McCopyStatus {
    <#
        Verification via Search-Mailbox (Compliance Search also sees Recoverable Items and would
        report the same numbers after a SoftDelete). Two counts per mailbox:
        Visible = folders the user sees; Total = including Recoverable Items.
    #>
    param([Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query, [switch] $TotalOnly)
    $i = 0
    foreach ($m in $Mailboxes) {
        Test-McCancel
        $i++; Set-McProgress ([int](100 * $i / $Mailboxes.Count)) (T 'progress.check' $m)
        try {
            $t = Search-Mailbox -Identity $m -SearchQuery $Query -SearchDumpster -EstimateResultOnly `
                     -WarningAction SilentlyContinue -ErrorAction Stop
            $visible = -1
            if (-not $TotalOnly) {
                $v = Search-Mailbox -Identity $m -SearchQuery $Query -SearchDumpster:$false -EstimateResultOnly `
                         -WarningAction SilentlyContinue -ErrorAction Stop
                $visible = [int]$v.ResultItemsCount
            }
            [pscustomobject]@{ Mailbox = $m.ToLowerInvariant(); Visible = $visible; Total = [int]$t.ResultItemsCount; Note = '' }
        } catch {
            [pscustomobject]@{ Mailbox = $m.ToLowerInvariant(); Visible = -1; Total = -1; Note = $_.Exception.Message }
        }
    }
}

function Get-McSettledCopyStatus {
    <#
        Verification after a delete. The search index can lag behind the delete for some seconds.
        If the user can still see the message in a mailbox, it waits and examines that mailbox again.
    #>
    param(
        [Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query,
        [int] $Retries = 3, [int] $DelaySeconds = 15
    )
    $check = @(Get-McCopyStatus -Mailboxes $Mailboxes -Query $Query)
    for ($n = 1; $n -le $Retries; $n++) {
        $left = @($check | Where-Object { $_.Visible -gt 0 } | ForEach-Object Mailbox)
        if (-not $left.Count) { break }
        Write-McLog (T 'log.checkRetry' $left.Count $DelaySeconds $n $Retries)
        Wait-McSeconds $DelaySeconds
        $again = @(Get-McCopyStatus -Mailboxes $left -Query $Query)
        $check = @($check | Where-Object { $_.Mailbox -notin $left }) + $again
    }
    @($check | Sort-Object Mailbox)
}

function Get-McOldSearches {
    <#
        MailCleaner searches older than the limit:
          * Compliance Searches MC_yyyyMMdd_HHmmss;
          * the shadow In-Place eDiscovery searches MC_yyyyMMdd_HHmmss-shadow that Exchange makes
            for each Compliance Search. Exchange can keep them after the Compliance Search is removed.
        The age comes from the name. Kind = ComplianceSearch or Shadow.
    #>
    param([int] $OlderThanDays = 7)
    $limit = (Get-Date).AddDays(-$OlderThanDays)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($s in @(Get-ComplianceSearch | Where-Object { Test-McSearchName $_.Name })) {
        $when = Get-McSearchNameTime $s.Name
        if ($when -and $when -lt $limit) { $list.Add([pscustomobject]@{ Name = $s.Name; When = $when; Kind = 'ComplianceSearch' }) }
    }
    if (Get-Command Get-MailboxSearch -ErrorAction SilentlyContinue) {
        foreach ($m in @(Get-MailboxSearch -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*-shadow' -and (Test-McSearchName $_.Name) })) {
            $when = Get-McSearchNameTime $m.Name
            if ($when -and $when -lt $limit) { $list.Add([pscustomobject]@{ Name = $m.Name; When = $when; Kind = 'Shadow' }) }
        }
    } else {
        Write-McLog (T 'log.noShadowCmd') 'Warn'
    }
    @($list | Sort-Object When, Name)
}

function Test-McNotFoundText {
    <# True for an Exchange "not found" message. "couldn.t": Exchange can write a typographic apostrophe (’). #>
    param([string] $Text)
    [bool]($Text -match 'couldn.t be found|could not be found|not found|не найден')
}

function Get-McComplianceSearchInfo {
    <#
        Reads one Compliance Search by -Identity. Returns Exists ($true, $false or $null = unknown),
        Status and Error.

        IMPORTANT: no "-ErrorAction Stop" for the *-ComplianceSearch* cmdlets. Tested on Exchange 2019:
        with Stop, these cmdlets fail with "Object reference not set" instead of the real answer.
        The errors come in -ErrorVariable.
    #>
    param([Parameter(Mandatory)] [string] $Name)
    $ge = $null
    $s = $null
    try {
        $s = Get-ComplianceSearch -Identity $Name -ErrorAction SilentlyContinue -ErrorVariable ge
    } catch {
        $ge = @($_)
    }
    $err = (@($ge) | Where-Object { $_ } | ForEach-Object { [string]$_.Exception.Message }) -join ' '
    $exists = if ($s) { $true } elseif ($err -and -not (Test-McNotFoundText $err)) { $null } else { $false }
    [pscustomobject]@{
        Exists = $exists
        Status = $(if ($s -and @($s)[0].PSObject.Properties['Status']) { [string]@($s)[0].Status } else { '' })
        Error  = $err
    }
}

function Get-McComplianceSearchStatus {
    <# The status of a Compliance Search (Completed, InProgress, ...), or '' if it is not found or not readable. #>
    param([Parameter(Mandatory)] [string] $Name)
    (Get-McComplianceSearchInfo $Name).Status
}

function Test-McComplianceSearchExists {
    <#
        True if the Compliance Search exists. Uses -Identity: the full list of searches is not reliable
        on all servers. "Not found" = false. Another error = true (unknown: do not report a removal).
    #>
    param([Parameter(Mandatory)] [string] $Name)
    $i = Get-McComplianceSearchInfo $Name
    $i.Exists -ne $false
}

function Remove-McSearches {
    <#
        Removes MailCleaner searches. A name with "-shadow" is a shadow In-Place eDiscovery search.
        For a Compliance Search, it also removes its actions and its shadow search, if Exchange kept it.
        Names that MailCleaner does not make are skipped.
    #>
    # Exchange 2019 removes a Compliance Search in the background: it can stay in the list for more
    # than 10 seconds after a remove without errors. Thus up to 60 seconds of checks.
    param([Parameter(Mandatory)] [string[]] $Names, [int] $RecheckTimes = 12, [int] $RecheckSeconds = 5)
    $done = @{}
    $hasShadowCmd = [bool](Get-Command Remove-MailboxSearch -ErrorAction SilentlyContinue)
    # Remove-ComplianceSearch usually removes the shadow search too. Thus the shadow search
    # can be gone when its turn comes: examine it before the remove.
    $shadowExists = { param($name) [bool](Get-MailboxSearch -Identity $name -ErrorAction SilentlyContinue) }

    # The Compliance Searches first, then the shadow searches that are still there.
    $ordered = @(@($Names | Where-Object { $_ -notlike '*-shadow' }) + @($Names | Where-Object { $_ -like '*-shadow' }))
    foreach ($n in $ordered) {
        if ($done.ContainsKey($n)) { continue }
        $done[$n] = $true
        if (-not (Test-McSearchName $n)) { Write-McLog (T 'log.cleanupSkip' $n) 'Warn'; continue }
        Test-McCancel
        # One failed search must not stop the cleanup of the others.
        try {
            if ($n -like '*-shadow') {
                if (-not $hasShadowCmd) { Write-McLog (T 'log.noShadowCmd') 'Warn'; continue }
                if (-not (& $shadowExists $n)) { Write-McLog (T 'log.shadowGone' $n); continue }
                Remove-MailboxSearch -Identity $n -Confirm:$false -ErrorAction Stop
                Write-McLog (T 'log.shadowRemoved' $n) 'Ok'
                continue
            }
            # The action first. Do NOT list all actions (Get-ComplianceSearchAction without -Identity):
            # on some servers one damaged action makes the full list fail with "Object reference not set".
            # MailCleaner gives its purge action a known name: <search name>_Purge.
            # A search that still runs cannot be removed (Stop-ComplianceSearch does not always stop it).
            # Do not touch it or its shadow search: the cleanup can remove it later.
            $status = Get-McComplianceSearchStatus $n
            if ($status -in 'Starting', 'InProgress', 'Stopping', 'NotStarted') {
                Write-McLog (T 'log.stillRunning' $n $status) 'Warn'
                $done["$n-shadow"] = $true
                continue
            }

            # IMPORTANT: no "-ErrorAction Stop" for the Remove-Compliance* cmdlets. Tested on Exchange 2019:
            # the cmdlet writes an internal error ("Object reference not set"). With Stop, this error
            # stops the remove in the middle and the search stays. Without Stop, the remove completes.
            $action = "${n}_Purge"
            try {
                if (Get-ComplianceSearchAction -Identity $action -ErrorAction SilentlyContinue) {
                    Remove-ComplianceSearchAction -Identity $action -Confirm:$false 2>$null
                }
            } catch { }
            # The shadow search first. With the shadow search, Remove-ComplianceSearch removes the
            # shadow search, fails with "Object reference not set" and keeps the Compliance Search.
            # Without the shadow search, Remove-ComplianceSearch works (tested on Exchange 2019).
            $shadow = "$n-shadow"
            if ($hasShadowCmd -and -not $done.ContainsKey($shadow)) {
                $done[$shadow] = $true
                if (& $shadowExists $shadow) {
                    try {
                        Remove-MailboxSearch -Identity $shadow -Confirm:$false -ErrorAction Stop
                        Write-McLog (T 'log.shadowRemoved' $shadow) 'Ok'
                    } catch {
                        Write-McLog (T 'log.cleanupError' $shadow $_.Exception.Message) 'Warn'
                    }
                } elseif ($shadow -in $Names) {
                    Write-McLog (T 'log.shadowGone' $shadow)
                }
            }

            # The Compliance Search. Without "-ErrorAction Stop" (see above). The result is the check
            # by -Identity, not the error stream: the internal error also occurs when the remove works.
            $removed = $false
            $lastErr = $null
            for ($attempt = 1; $attempt -le 2 -and -not $removed; $attempt++) {
                $rmErr = $null
                try {
                    Remove-ComplianceSearch -Identity $n -Confirm:$false -ErrorVariable rmErr 2>$null
                } catch {
                    $rmErr = @($_)
                }
                if (-not (Test-McComplianceSearchExists $n)) { $removed = $true }
                elseif ($rmErr) { $lastErr = @($rmErr)[0] }
            }
            if (-not $removed) {
                # Exchange removes the search in the background. Wait for it.
                for ($k = 1; $k -le [math]::Max(1, $RecheckTimes) -and -not $removed; $k++) {
                    Set-McProgress ([int](100 * $k / [math]::Max(1, $RecheckTimes))) (T 'progress.removeWait' $n)
                    if (-not (Test-McComplianceSearchExists $n)) { $removed = $true; break }
                    if ($k -lt $RecheckTimes) { Wait-McSeconds $RecheckSeconds }
                }
            }
            if (-not $removed) {
                $info = Get-McComplianceSearchInfo $n
                $seen = if ($info.Status) { $info.Status } elseif ($info.Error) { $info.Error } else { '?' }
                if ($lastErr) {
                    # Exchange returned an error and the search is still there: a real error.
                    throw (T 'err.stillThere' $lastErr.Exception.Message ($RecheckTimes * $RecheckSeconds))
                }
                # No error from Exchange: the remove is probably still running in the background.
                Write-McLog (T 'log.removePending' $n ($RecheckTimes * $RecheckSeconds) $seen) 'Warn'
                continue
            }
            Write-McLog (T 'log.cleanupRemoved' $n) 'Ok'
        } catch {
            Write-McLog (T 'log.cleanupError' $n $_.Exception.Message) 'Error'
        }
    }
}

function Invoke-McOperation {
    <# Single entry point used by the GUI's background runspace. #>
    param([Parameter(Mandatory)] [string] $Name, [hashtable] $Params = @{})
    switch ($Name) {
        'Connect'      { Connect-McExchange -Server $Params.Server }
        'Resolve'      { @(Resolve-McAddressInputs -Inputs $Params.Inputs) }
        'Search'       { Invoke-McSearch -Criteria $Params.Criteria -TimeoutMinutes $Params.TimeoutMinutes }
        'SoftDelete'   { $searchName = $Params.SearchName
                         if (-not $searchName) {
                             # The quick search has no Compliance Search. Purge needs one: make it now.
                             $searchName = New-McSearchName
                             Write-McLog (T 'log.softSearch' $searchName @($Params.Mailboxes).Count)
                             [void](New-McComplianceSearchRun -Name $searchName -Locations $Params.Mailboxes `
                                        -Query $Params.Query -TimeoutMinutes $Params.TimeoutMinutes)
                         }
                         $r = Invoke-McSoftDelete -SearchName $searchName -TimeoutMinutes $Params.TimeoutMinutes
                         [pscustomobject]@{
                             Delete = $r; SearchName = $searchName
                             Check = @(Get-McSettledCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query)
                         } }
        'HardDelete'   { Write-McLog (T 'log.precheck')
                         $pre = @(Get-McCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query -TotalOnly)
                         $bad = @(Test-McHardDeleteScope -Status $pre -Expected $Params.Expected)
                         if ($bad.Count) {
                             $list = ($bad | ForEach-Object { T 'scope.item' $_.Mailbox $_.Expected $_.Actual $_.Reason }) -join '; '
                             throw (T 'err.precheck' $list)
                         }
                         Write-McLog (T 'log.precheckOk') 'Ok'
                         $d = Invoke-McHardDelete -Mailboxes $Params.Mailboxes -Query $Params.Query
                         $check = if ($d.Cancelled) { @() } else { @(Get-McSettledCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query) }
                         [pscustomobject]@{ Delete = @($d.Results); Cancelled = $d.Cancelled; Check = $check } }
        'Verify'       { [pscustomobject]@{ Check = @(Get-McCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query) } }
        'ListOld'      { Get-McOldSearches -OlderThanDays $Params.Days }
        'RemoveOld'    { Remove-McSearches -Names $Params.Names }
        default        { throw (T 'err.unknownOp' $Name) }
    }
}

#endregion
