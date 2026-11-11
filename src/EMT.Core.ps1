<#
    EMT.Core.ps1 — UI-independent logic for Exchange Message Trail (EMT).

    Loaded (dot-sourced) by:
      * ExchangeMessageTrail.ps1 (the GUI) inside its background runspace;
      * tests\EMT.Core.Tests.ps1 (Pester 5).

    Pure helpers (no Exchange calls) come first and are fully unit-tested.
    Exchange-dependent operations follow; they expect the Exchange cmdlets to be
    imported into the session (see Connect-EmtExchange).
#>

Set-StrictMode -Version 2.0

# User interface texts (ru/en) and the T function.
. (Join-Path $PSScriptRoot 'EMT.Strings.ps1')

# New searches use EMT_. The older version (MailCleaner) used MC_: the cleanup accepts both.
function Get-EmtSearchPrefix { 'EMT_' }

# Up to this number of mailboxes, the search uses Search-Mailbox (seconds) instead of Compliance Search (minutes).
function Get-EmtQuickSearchLimit { 10 }

function Get-EmtExchangeCommandNames {
    @(
        'Get-TransportService', 'Get-MessageTrackingLog', 'Get-Recipient', 'Get-AcceptedDomain',
        'New-ComplianceSearch', 'Start-ComplianceSearch', 'Stop-ComplianceSearch',
        'Get-ComplianceSearch', 'Remove-ComplianceSearch',
        'New-ComplianceSearchAction', 'Get-ComplianceSearchAction', 'Remove-ComplianceSearchAction',
        'Search-Mailbox', 'Get-MailboxSearch', 'Remove-MailboxSearch',
        'Get-InboxRule', 'Get-MailboxJunkEmailConfiguration', 'Get-Mailbox', 'Get-MailboxStatistics'
    )
}

#region ---------- Logging / progress / cancellation (shared with the UI via $global:EmtSync) ----------

function Write-EmtLog {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('Info', 'Ok', 'Warn', 'Error')] [string] $Level = 'Info'
    )
    $line = '{0:HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level.ToUpperInvariant().PadRight(5), $Message
    if ((Get-Variable -Name EmtSync -Scope Global -ErrorAction SilentlyContinue) -and $global:EmtSync) {
        $global:EmtSync.Log.Enqueue($line)
    } else {
        Write-Verbose $line
    }
}

function Set-EmtProgress {
    param([int] $Percent, [string] $Status = '')
    if ((Get-Variable -Name EmtSync -Scope Global -ErrorAction SilentlyContinue) -and $global:EmtSync) {
        $global:EmtSync.Progress = [math]::Max(0, [math]::Min(100, $Percent))
        $global:EmtSync.Status   = $Status
    }
}

function Test-EmtCancelRequested {
    [bool]((Get-Variable -Name EmtSync -Scope Global -ErrorAction SilentlyContinue) -and
           $global:EmtSync -and $global:EmtSync.Cancel)
}

function Test-EmtCancel {
    if (Test-EmtCancelRequested) {
        throw [System.OperationCanceledException]::new((T 'core.cancelled'))
    }
}

function Wait-EmtSeconds {
    <# Waits and checks for cancel each second. #>
    param([int] $Seconds)
    for ($i = 0; $i -lt $Seconds; $i++) { Test-EmtCancel; Start-Sleep -Seconds 1 }
}

#endregion

#region ---------- Pure helpers ----------

function Split-EmtAddressList {
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
        if ($words.Count -gt 1 -and @($words | Where-Object { Test-EmtFullAddress $_ }).Count -eq $words.Count) { $words }
        else { $p -replace '\s+', ' ' }
    }
    @($items | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
}

function Test-EmtFullAddress {
    <# True for a full SMTP address: name@domain.tld. A part of an address or a name is not full. #>
    param([string] $Text)
    [bool]($Text -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Get-EmtPeriodEnd {
    <#
        The end of the period includes its minute: 23:59 is 23:59:59.999, 14:30 is 14:30:59.999.
        The time fields have minutes only. The tracking log uses the exact time: without this, a message
        in the last minute of the period is not found.
    #>
    param([Parameter(Mandatory)] [datetime] $Value)
    $m = New-Object DateTime($Value.Year, $Value.Month, $Value.Day, $Value.Hour, $Value.Minute, 0, $Value.Kind)
    $m.AddMinutes(1).AddTicks(-1)
}
function Get-EmtQuickPeriodDays {
    <# The quick period buttons: the first day and the day after the last day, from today. #>
    [ordered]@{ Today = @(0, 1); Yesterday = @(-1, 0); Week = @(-6, 1); Days14 = @(-13, 1) }
}
function Get-EmtQuickPeriod {
    <#
        The quick period button that has the same period as the fields: Today, Yesterday, Week, Days14,
        or an empty string. The start is 0:00 of the first day, the end is 23:59 of the last day (to the minute).
        The window marks this button. A manual change of the period removes the mark.
    #>
    param([Parameter(Mandatory)] [datetime] $Start, [Parameter(Mandatory)] [datetime] $End, [datetime] $Now = (Get-Date))
    $s = New-Object DateTime($Start.Year, $Start.Month, $Start.Day, $Start.Hour, $Start.Minute, 0)
    $e = New-Object DateTime($End.Year, $End.Month, $End.Day, $End.Hour, $End.Minute, 0)
    $days = Get-EmtQuickPeriodDays
    foreach ($k in $days.Keys) {
        if ($s -eq $Now.Date.AddDays($days[$k][0]) -and $e -eq $Now.Date.AddDays($days[$k][1]).AddMinutes(-1)) { return $k }
    }
    ''
}
function Get-EmtKqlDateRange {
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

function Get-EmtBaseSubject {
    <# Subject without reply/forward prefixes (RE:, FW:, Fwd:, AW:, WG:, Отв:, Пересл: …), lower case. #>
    param([string] $Subject)
    $rx = '^\s*((re|fw|fwd|aw|wg|tr|rv|sv|vs|отв|ответ|пересл|переслано)\s*(\[\d+\])?\s*:\s*)+'
    (([string]$Subject) -replace $rx, '').Trim().ToLowerInvariant()
}

function ConvertTo-EmtKqlPhrase {
    param([string] $Text)
    '"' + ($Text -replace '"', '').Trim() + '"'
}

function New-EmtSearchQuery {
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
    if ($Subject)    { $parts.Add('(subject:' + (ConvertTo-EmtKqlPhrase $Subject) + ')') }
    if ($Attachment) { $parts.Add('(attachment:' + (ConvertTo-EmtKqlPhrase $Attachment) + ')') }
    if ($Keywords)   { $parts.Add('(' + (ConvertTo-EmtKqlPhrase $Keywords) + ')') }
    if ($parts.Count -eq 0) {
        throw (T 'crit.datesOnly')
    }
    # The recipient limits the query: only messages that this address received (To, Cc or Bcc).
    # Without it, "from:x" finds all messages of the sender to all people.
    $rcpts = @($Recipients | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    if ($rcpts.Count) {
        $parts.Add('(' + (($rcpts | ForEach-Object { "to:$_ OR cc:$_ OR bcc:$_" }) -join ' OR ') + ')')
    }
    $range = Get-EmtKqlDateRange -Start $Start -End $End -UtcOffset $UtcOffset
    $parts.Add("(sent>=$($range.From))")
    $parts.Add("(sent<$($range.To))")
    $parts -join ' AND '
}

function Complete-EmtCriteria {
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

function Test-EmtCriteria {
    <#
        Returns a list of problems; empty list = criteria are usable.
        -ForTrace: the trace only reads the tracking log. A recipient alone is a usable criterion there
        ("what did this address receive"). The delete needs more: a query by dates and recipient only is too wide.
    #>
    param([Parameter(Mandatory)] [hashtable] $Criteria, [switch] $ForTrace)
    $Criteria = Complete-EmtCriteria $Criteria
    $problems = @()
    $hasKey = @($Criteria.Senders).Count -or $Criteria.Subject -or $Criteria.Attachment -or
              $Criteria.Keywords -or $Criteria.MessageId -or ($ForTrace -and $Criteria.Recipient)
    if (-not $hasKey) { $problems += (T 'crit.noKey') }
    if ($Criteria.End -le $Criteria.Start) { $problems += (T 'crit.endBeforeStart') }
    if ($Criteria.OnlyRecipient -and -not $Criteria.Recipient) { $problems += (T 'crit.onlyNeedsRecipient') }
    if ($Criteria.OnlyRecipient -and $Criteria.AllMailboxes)   { $problems += (T 'crit.onlyAndAll') }
    if (($Criteria.End - $Criteria.Start).TotalDays -gt 31)    { $problems += (T 'crit.periodTooLong') }
    $problems
}

function Get-EmtSearchHits {
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

function Get-EmtSum {
    <#
        Sum of a numeric property. 0 for an empty list.
        Not "Measure-Object -Sum": with an empty list it has no Sum property, and StrictMode stops.
    #>
    param([object[]] $Items = @(), [Parameter(Mandatory)] [string] $Property)
    $sum = 0
    foreach ($i in $Items) { if ($null -ne $i -and $null -ne $i.$Property) { $sum += $i.$Property } }
    $sum
}

function Test-EmtInternalDomain {
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

function ConvertFrom-EmtSizeText {
    <# Search-Mailbox gives the size as text, e.g. "1.25 MB (1,310,720 bytes)". Returns KB. #>
    param([string] $Text)
    if ($Text -match '\((?<b>[\d,\.\s ]+)\s*bytes\)') {
        $digits = $Matches['b'] -replace '[^\d]', ''
        if ($digits) { return [math]::Round([int64]$digits / 1KB, 1) }
    }
    0
}

function Test-EmtServerGeneratedId {
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

function Get-EmtMessageSummary {
    <# One row per MessageId: first submit/receive time, sender, subject, recipient count, server-generated flag. #>
    param([object[]] $Events = @(), [string[]] $ServerNames = @())
    # An event without MessageId (for example NOTIFYMAPI) is not a message: no row for it.
    $Events | Where-Object { $_ -and $_.MessageId } | Group-Object MessageId | ForEach-Object {
        $g      = @($_.Group | Sort-Object Timestamp)
        $first  = ($g | Where-Object { $_.EventId -in 'RECEIVE', 'SUBMIT' } | Select-Object -First 1)
        if (-not $first) { $first = $g[0] }
        $rcpt   = @($g | ForEach-Object { $_.Recipients } | Where-Object { $_ } |
                    ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
        $auto   = Test-EmtServerGeneratedId -MessageId $_.Name -ServerNames $ServerNames
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

function Get-EmtMissingMailboxes {
    <# Locations that were searched but returned no items. #>
    param([string[]] $Locations = @(), [object[]] $Hits = @())
    if ($Locations -contains 'All') { return @() }
    $found = @($Hits | Where-Object { $_.Items -gt 0 } | ForEach-Object { $_.Mailbox.ToLowerInvariant() })
    @($Locations | ForEach-Object { $_.ToLowerInvariant() } | Where-Object { $_ -notin $found } | Sort-Object -Unique)
}

function Get-EmtExtraItemMailboxes {
    <#
        The mailboxes with more found items than messages in the tracking log.
        The query (KQL) cannot use the message ID: it finds all items with the sender, the subject and the days.
        More items can be other messages with the same subject, or copies in Recoverable Items.
        No tracking messages (the result is not confirmed): no list, the warning "not confirmed" tells it.
    #>
    param([object[]] $Found = @(), [int] $Messages)
    if ($Messages -le 0) { return }
    @($Found | Where-Object { $_ -and [int](Get-EmtProp $_ 'Items') -gt $Messages } | ForEach-Object { [string]$_.Mailbox })
}
function Get-EmtLargeDeleteWarning {
    <# Text of the warning for a large delete. Empty text = no warning. #>
    param([int] $Items, [int] $Messages, [int] $MaxItems = 100, [int] $MaxMessages = 5)
    $reasons = @()
    if ($Items -gt $MaxItems)       { $reasons += (T 'large.items' $MaxItems $Items) }
    if ($Messages -gt $MaxMessages) { $reasons += (T 'large.messages' $MaxMessages $Messages) }
    if ($reasons.Count) { T 'large.warning' ($reasons -join ', ') } else { '' }
}

function Test-EmtSearchName {
    <# True only for the names that EMT makes: EMT_yyyyMMdd_HHmmss (or the older MC_...), with or without "-shadow". #>
    param([string] $Name)
    $Name -cmatch '^(EMT|MC)_\d{8}_\d{6}(-shadow)?$'
}

function Get-EmtSearchNameTime {
    <# The creation time from an EMT search name (EMT_ or MC_ + yyyyMMdd_HHmmss[-shadow]), or $null. #>
    param([string] $Name)
    if ($Name -cnotmatch '^(EMT|MC)_(?<t>\d{8}_\d{6})') { return $null }
    $t = [datetime]::MinValue
    if ([datetime]::TryParseExact($Matches['t'], 'yyyyMMdd_HHmmss', [Globalization.CultureInfo]::InvariantCulture,
                                  [Globalization.DateTimeStyles]::None, [ref]$t)) { return $t }
    $null
}

function New-EmtSearchName {
    param([datetime] $Now = (Get-Date))
    '{0}{1:yyyyMMdd_HHmmss}' -f (Get-EmtSearchPrefix), $Now
}

function Write-EmtAudit {
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

function Connect-EmtExchange {
    <# Opens an Exchange remote session and imports the needed cmdlets globally. Returns capabilities. #>
    param([Parameter(Mandatory)] [string] $Server)
    if ((Get-Variable -Name EmtExSession -Scope Global -ErrorAction SilentlyContinue) -and $global:EmtExSession) {
        Remove-PSSession $global:EmtExSession -ErrorAction SilentlyContinue
        $global:EmtExSession = $null
    }
    Write-EmtLog (T 'log.connecting' $Server)
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/PowerShell/" `
                   -Authentication Kerberos -ErrorAction Stop
    $module = Import-PSSession $session -CommandName (Get-EmtExchangeCommandNames) -DisableNameChecking `
                  -AllowClobber -ErrorAction Stop -WarningAction SilentlyContinue
    Import-Module $module -Global -DisableNameChecking -WarningAction SilentlyContinue
    $global:EmtExSession = $session

    # Accepted domains: an address in another domain is external without a Get-Recipient call.
    $global:EmtAcceptedDomains = @()
    if (Get-Command Get-AcceptedDomain -ErrorAction SilentlyContinue) {
        try {
            $global:EmtAcceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop |
                                          ForEach-Object { ([string]$_.DomainName).ToLowerInvariant() })
            Write-EmtLog (T 'log.domains' ($global:EmtAcceptedDomains -join ', '))
        } catch {
            Write-EmtLog (T 'log.domainsFailed' $_.Exception.Message) 'Warn'
        }
    }
    $caps = Get-EmtCapabilities
    Write-EmtLog (T 'log.connected' $Server) 'Ok'
    $caps
}

function Get-EmtAcceptedDomains {
    if ((Get-Variable -Name EmtAcceptedDomains -Scope Global -ErrorAction SilentlyContinue) -and $global:EmtAcceptedDomains) {
        return @($global:EmtAcceptedDomains)
    }
    @()
}

function Get-EmtCapabilities {
    $has = { param($n) [bool](Get-Command $n -ErrorAction SilentlyContinue) }
    $sm  = Get-Command Search-Mailbox -ErrorAction SilentlyContinue
    [pscustomobject]@{
        Tracking   = (& $has 'Get-MessageTrackingLog') -and (& $has 'Get-TransportService')
        Search     = (& $has 'New-ComplianceSearch') -and (& $has 'Start-ComplianceSearch') -and (& $has 'Get-ComplianceSearch')
        Verify     = [bool]$sm
        SoftDelete = (& $has 'New-ComplianceSearchAction') -and (& $has 'Get-ComplianceSearchAction')
        HardDelete = [bool]($sm -and $sm.Parameters.ContainsKey('DeleteContent'))
        MailboxCheck = (& $has 'Get-Mailbox') -and (& $has 'Get-InboxRule')
    }
}

function Find-EmtTrackingEvents {
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
    Write-EmtLog (T 'log.trackServers' $servers.Count)
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
            Test-EmtCancel
            $n++; Set-EmtProgress ([int](30 * $n / $total)) (T 'progress.tracking' $srv)
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
                $bases = @($orig | ForEach-Object { Get-EmtBaseSubject $_.MessageSubject } | Sort-Object -Unique)
                $fw = @($hits | Where-Object {
                    ([string]$_.Sender).ToLowerInvariant() -notin $Senders -and (Get-EmtBaseSubject $_.MessageSubject) -in $bases
                })
                $hits = @($orig) + @($fw)
            } else {
                $hits = $orig
            }
        } elseif ($IncludeForwards -and $Subject) {
            Write-EmtLog (T 'log.noOriginal') 'Warn'
        } else {
            $hits = @()
        }
    }
    # Events without MessageId (for example NOTIFYMAPI) are not messages. They stay in the event list only.
    $ids   = @($hits | ForEach-Object { [string]$_.MessageId } | Where-Object { $_ } | Sort-Object -Unique)
    $blank = @($hits | Where-Object { -not $_.MessageId })
    Write-EmtLog (T 'log.trackCount' $ids.Count)

    $source = $raw
    if ($Recipient -and $CompleteRecipients -and $ids.Count) {
        Write-EmtLog (T 'log.secondPass')
        $pass = Get-EmtEventsByMessageId -MessageIds $ids -Start $Start -End $End -Servers $servers -ProgressFrom 30 -ProgressTo 35
        $source = $pass.Events
        foreach ($x in $pass.Errors) { $trackErrors.Add($x) }
    }
    $errs = @($trackErrors | Sort-Object -Unique)
    foreach ($t in $errs) { Write-EmtLog (T 'log.trackError' $t) 'Error' }
    $events = @(Select-EmtUniqueEvents -Events $source -MessageIds $ids) + @(Select-EmtUniqueEvents -Events $blank -KeepBlank)
    [pscustomobject]@{
        Events  = @($events | Sort-Object Timestamp)
        Servers = $servers
        Errors  = $errs
    }
}

function Get-EmtEventsByMessageId {
    <#
        All tracking events of each message ID on each transport server.
        Get-MessageTrackingLog accepts one message ID: one call for each message ID and each server.
        The end is one day later: a delivery can occur after the end of the period. Returns Events and Errors.
    #>
    param(
        [Parameter(Mandatory)] [string[]] $MessageIds, [Parameter(Mandatory)] [datetime] $Start,
        [Parameter(Mandatory)] [datetime] $End, [Parameter(Mandatory)] [string[]] $Servers,
        [int] $ProgressFrom = 0, [int] $ProgressTo = 100
    )
    $events = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[string]
    $total = [math]::Max(1, $MessageIds.Count * $Servers.Count); $n = 0
    foreach ($id in $MessageIds) {
        foreach ($srv in $Servers) {
            Test-EmtCancel
            $n++; Set-EmtProgress ($ProgressFrom + [int](($ProgressTo - $ProgressFrom) * $n / $total)) (T 'progress.trackingId' $srv)
            $ev = $null
            foreach ($e in @(Get-MessageTrackingLog -Server $srv -MessageId $id -Start $Start -End $End.AddDays(1) `
                                 -ResultSize Unlimited -ErrorAction SilentlyContinue -ErrorVariable ev)) {
                if ($e) { $events.Add($e) }
            }
            foreach ($x in @($ev)) { if ($x) { $errors.Add("${srv}: $($x.Exception.Message)") } }
        }
    }
    # ToArray(), not @(): PowerShell 5.1 can fail on @() of a generic List.
    [pscustomobject]@{ Events = $events.ToArray(); Errors = $errors.ToArray() }
}

function Select-EmtUniqueEvents {
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

function Resolve-EmtRecipients {
    <#
        Classifies addresses: Internal mailbox / External / Group / Other.
        Addresses outside the accepted domains are external without a call.
        The other addresses are resolved in batches of 20 (one Get-Recipient -Filter call per batch).
        If a batch call fails, the batch is resolved address by address.
    #>
    param([string[]] $Addresses = @(), [string[]] $AcceptedDomains = @(), [int] $BatchSize = 20)
    $list = @($Addresses | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $lookup = @($list | Where-Object { -not $AcceptedDomains.Count -or (Test-EmtInternalDomain $_ $AcceptedDomains) })
    $byAddress = @{}

    for ($b = 0; $b -lt $lookup.Count; $b += $BatchSize) {
        Test-EmtCancel
        $batch = @($lookup[$b..([math]::Min($b + $BatchSize, $lookup.Count) - 1)])
        Set-EmtProgress (35 + [int](5 * ($b + $batch.Count) / [math]::Max(1, $lookup.Count))) (T 'progress.recipients' ($b + $batch.Count) $lookup.Count)
        $filter = ($batch | ForEach-Object { "EmailAddresses -eq 'smtp:$($_ -replace "'", "''")'" }) -join ' -or '
        try {
            foreach ($r in @(Get-Recipient -Filter $filter -ResultSize Unlimited -ErrorAction Stop)) {
                $proxies = @($r.EmailAddresses | ForEach-Object { ([string]$_) -replace '^smtp:', '' } |
                             ForEach-Object { $_.ToLowerInvariant() })
                foreach ($a in $batch) { if ($a -in $proxies) { $byAddress[$a] = $r } }
            }
        } catch {
            Write-EmtLog (T 'log.batchFailed' $_.Exception.Message) 'Warn'
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

function Find-EmtAddressCandidates {
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
        Write-EmtLog (T 'log.resolveFailed' $Text $_.Exception.Message) 'Warn'
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

function Resolve-EmtAddressInputs {
    <# For each partial address: Input and Candidates (see Find-EmtAddressCandidates). #>
    param([Parameter(Mandatory)] [string[]] $Inputs, [int] $Max = 25)
    Write-EmtLog (T 'log.resolveStart' ($Inputs -join ', '))
    $i = 0
    foreach ($t in $Inputs) {
        Test-EmtCancel
        $i++; Set-EmtProgress ([int](100 * $i / $Inputs.Count)) (T 'status.resolve')
        [pscustomobject]@{ Input = $t; Candidates = @(Find-EmtAddressCandidates -Text $t -Max $Max); Max = $Max }
    }
}

function Wait-EmtComplianceSearch {
    param([Parameter(Mandatory)] [string] $Name, [int] $TimeoutMinutes = 60, [int] $ProgressFrom = 40, [int] $ProgressTo = 95)
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ($true) {
        Start-Sleep -Seconds 3
        if ((Get-Variable -Name EmtSync -Scope Global -ErrorAction SilentlyContinue) -and $global:EmtSync -and $global:EmtSync.Cancel) {
            Write-EmtLog (T 'log.csStopping' $Name) 'Warn'
            $stopErr = $null
            Stop-ComplianceSearch -Identity $Name -Confirm:$false -ErrorAction SilentlyContinue -ErrorVariable stopErr
            if ($stopErr) { Write-EmtLog (T 'log.stopError' $Name (@($stopErr)[0].Exception.Message)) 'Warn' }
            # Exchange on-premises can continue the search after Stop-ComplianceSearch until it completes.
            Write-EmtLog (T 'log.stopNote' $Name) 'Warn'
            Test-EmtCancel
        }
        $s = Get-ComplianceSearch -Identity $Name
        $pct = [int]$s.JobProgress
        Set-EmtProgress ($ProgressFrom + [int](($ProgressTo - $ProgressFrom) * $pct / 100)) (T 'progress.cs' $s.Status $pct)
        if ($s.Status -in 'Completed', 'Stopped', 'Failed') { break }
        if ((Get-Date) -gt $deadline) { throw (T 'err.csTimeout' $Name $TimeoutMinutes) }
    }
    if ($s.Status -ne 'Completed') { throw (T 'err.csStatus' $Name $s.Status) }
    $s
}

function Invoke-EmtSearch {
    <# Full pipeline: tracking -> recipients -> KQL -> Compliance Search -> hits. #>
    param([Parameter(Mandatory)] [hashtable] $Criteria, [int] $TimeoutMinutes = 60)

    $Criteria = Complete-EmtCriteria $Criteria
    $problems = @(Test-EmtCriteria $Criteria)
    if ($problems.Count) { throw ($problems -join ' ') }

    $senders = @($Criteria.Senders)
    $caps = Get-EmtCapabilities
    if (-not $caps.Search -and -not $caps.Verify) {
        throw (T 'err.noSearchRights')
    }
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $stage = { param($text) Write-EmtLog (T 'log.stage' $text $timer.Elapsed.TotalSeconds) }

    # ---- 0. Only the recipient mailbox ----
    $onlyMailbox = ''
    if ($Criteria.OnlyRecipient) {
        $r = @(Resolve-EmtRecipients -Addresses $Criteria.Recipient)
        if (-not $r.Count -or $r[0].Kind -ne 'Internal') {
            throw (T 'err.notInternal' $Criteria.Recipient)
        }
        $onlyMailbox = $r[0].Mailbox
        Write-EmtLog (T 'log.onlyMode' $onlyMailbox)
    }

    # ---- 1. Message tracking ----
    $tracking = [pscustomobject]@{ Events = @(); Servers = @(); Errors = @() }
    if ($caps.Tracking) {
        $tracking = Find-EmtTrackingEvents -Senders $senders -Subject $Criteria.Subject -Recipient $Criteria.Recipient `
                        -MessageId $Criteria.MessageId -Start $Criteria.Start -End $Criteria.End `
                        -IncludeForwards:([bool]$Criteria.IncludeForwards) -CompleteRecipients:(-not $onlyMailbox)
        & $stage (T 'stage.tracking')
    } else {
        Write-EmtLog (T 'log.noTrackRights') 'Warn'
    }
    $events  = @($tracking.Events)
    $trackErrors = @($tracking.Errors)
    $summary = @(Get-EmtMessageSummary -Events $events -ServerNames $tracking.Servers)
    if (-not $events.Count) { Write-EmtLog (T 'log.trackEmpty') 'Warn' }
    foreach ($m in @($summary | Where-Object ServerGenerated)) { Write-EmtLog (T 'log.serverGenerated' $m.MessageId $m.Sender) 'Warn' }

    # ---- 2. Recipients ----
    $addresses = @($events | ForEach-Object { $_.Recipients } | Where-Object { $_ } |
                   ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $trackSenders = @($events | ForEach-Object { [string]$_.Sender } | Where-Object { $_ } |
                      ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $recipients = @(Resolve-EmtRecipients -Addresses (@($addresses) + @($trackSenders) | Sort-Object -Unique) `
                        -AcceptedDomains (Get-EmtAcceptedDomains))
    $internal = @($recipients | Where-Object Kind -eq 'Internal' | ForEach-Object Mailbox | Sort-Object -Unique)
    $external = @($recipients | Where-Object { $_.Kind -eq 'External' -and $_.Address -in $addresses })
    if ($external.Count) { Write-EmtLog (T 'log.external' $external.Count) 'Warn' }
    & $stage (T 'stage.recipients' $recipients.Count)

    # ---- 3. KQL ----
    $kqlSenders = @($senders)
    if (($Criteria.IncludeForwards -or -not $kqlSenders.Count) -and $trackSenders.Count) { $kqlSenders += $trackSenders }
    $kqlSubject = $Criteria.Subject
    if (-not $kqlSubject -and $Criteria.MessageId -and $summary.Count) { $kqlSubject = $summary[0].Subject }
    $kqlRecipients = @()
    if ($Criteria.Recipient) { $kqlRecipients = @($Criteria.Recipient) }
    $query = New-EmtSearchQuery -Senders $kqlSenders -Subject $kqlSubject -Attachment $Criteria.Attachment `
                 -Keywords $Criteria.Keywords -Start $Criteria.Start -End $Criteria.End -Recipients $kqlRecipients

    $locations = @($internal)
    if ($onlyMailbox) {
        $locations = @($onlyMailbox)
    } elseif ($Criteria.AllMailboxes) {
        $locations = @('All')
    } elseif (-not $locations.Count) {
        # No recipient mailboxes from the tracking log. With a recipient or senders in the criteria,
        # search only their mailboxes: a search of all mailboxes is slow and not necessary.
        $own = @(Resolve-EmtRecipients -Addresses (@($kqlRecipients) + @($senders)) -AcceptedDomains (Get-EmtAcceptedDomains) |
                 Where-Object Kind -eq 'Internal' | ForEach-Object Mailbox | Sort-Object -Unique)
        if ($kqlRecipients.Count -and $own.Count) {
            $locations = $own
            Write-EmtLog (T 'log.recipientScope' ($own -join ', ')) 'Warn'
        } else {
            Write-EmtLog (T 'log.searchAll') 'Warn'
            $locations = @('All')
        }
    }
    Write-EmtLog "KQL: $query"

    # ---- 4. Mailbox search ----
    Test-EmtCancel
    $isAll = $locations[0] -eq 'All'
    $quick = (-not $isAll) -and $caps.Verify -and ($locations.Count -le (Get-EmtQuickSearchLimit) -or -not $caps.Search)
    if ($isAll -and -not $caps.Search) { throw (T 'err.allNeedsCs') }

    if ($quick) {
        Write-EmtLog (T 'log.quick' $locations.Count)
        $hits = @(Invoke-EmtQuickSearch -Mailboxes $locations -Query $query)
        $name      = ''
        $items     = [int](Get-EmtSum -Items @($hits | Where-Object { $_.Items -gt 0 }) -Property Items)
        $sizeKB    = [double](Get-EmtSum -Items @($hits | Where-Object { $_.Items -gt 0 }) -Property SizeKB)
        $unindexed = -1
        $errors    = (@($hits | Where-Object { $_.Error } | ForEach-Object { "$($_.Mailbox): $($_.Error)" }) -join '; ')
    } else {
        $name = New-EmtSearchName
        Write-EmtLog (T 'log.cs' $name $(if ($isAll) { T 'common.all' } else { $locations.Count }))
        $s = New-EmtComplianceSearchRun -Name $name -Locations $locations -Query $query -TimeoutMinutes $TimeoutMinutes
        $hits      = @(Get-EmtSearchHits -SuccessResults ([string]$s.SuccessResults))
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
    $missing = if ($found.Count) { @(Get-EmtMissingMailboxes -Locations $locations -Hits $hits) } else { @() }
    Write-EmtLog (T 'log.found' $items $found.Count $sizeKB) 'Ok'
    if ($errors) { Write-EmtLog (T 'log.searchErrors' $errors) 'Error' }
    if ($unindexed -gt 0) { Write-EmtLog (T 'log.unindexed' $unindexed) 'Warn' }
    # The search found items, but the tracking log has no message: nothing confirms that these are the
    # messages the user wants. The window shows this and asks for the typed word before a delete.
    $unconfirmed = ($found.Count -gt 0) -and ($summary.Count -eq 0)
    if ($unconfirmed) { Write-EmtLog (T 'log.unconfirmed' $items) 'Warn' }
    Set-EmtProgress 100 (T 'progress.done')

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

function New-EmtComplianceSearchRun {
    <# Makes, starts and waits for one Compliance Search. Returns the search object. #>
    param(
        [Parameter(Mandatory)] [string] $Name, [Parameter(Mandatory)] [string[]] $Locations,
        [Parameter(Mandatory)] [string] $Query, [int] $TimeoutMinutes = 60
    )
    New-ComplianceSearch -Name $Name -ExchangeLocation $Locations -ContentMatchQuery $Query | Out-Null
    Start-ComplianceSearch -Identity $Name
    Wait-EmtComplianceSearch -Name $Name -TimeoutMinutes $TimeoutMinutes
}

function Invoke-EmtQuickSearch {
    <#
        Search-Mailbox -EstimateResultOnly for each mailbox. For a few mailboxes this takes
        seconds, a Compliance Search takes minutes. The count includes Recoverable Items,
        the same as a Compliance Search.
    #>
    param([Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query)
    $i = 0
    foreach ($m in $Mailboxes) {
        Test-EmtCancel
        $i++; Set-EmtProgress (40 + [int](55 * $i / $Mailboxes.Count)) (T 'progress.quick' $m)
        try {
            $r = Search-Mailbox -Identity $m -SearchQuery $Query -EstimateResultOnly `
                     -WarningAction SilentlyContinue -ErrorAction Stop
            [pscustomobject]@{
                Mailbox = $m.ToLowerInvariant(); Items = [int]$r.ResultItemsCount
                SizeKB = ConvertFrom-EmtSizeText ([string]$r.ResultItemsSize); Error = ''
            }
        } catch {
            [pscustomobject]@{ Mailbox = $m.ToLowerInvariant(); Items = 0; SizeKB = 0; Error = $_.Exception.Message }
        }
    }
}

function Invoke-EmtSoftDelete {
    <# Purge (SoftDelete) of an existing Compliance Search: items go to Recoverable Items. #>
    param([Parameter(Mandatory)] [string] $SearchName, [int] $TimeoutMinutes = 60)
    Write-EmtLog (T 'log.softStart' $SearchName)
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete -Confirm:$false | Out-Null
    $actionName = "${SearchName}_Purge"
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    do {
        Start-Sleep -Seconds 3
        $a = Get-ComplianceSearchAction -Identity $actionName
        Set-EmtProgress 50 (T 'progress.soft' $a.Status)
        if ((Get-Date) -gt $deadline) { throw (T 'err.softTimeout' $TimeoutMinutes) }
    } while ($a.Status -notin 'Completed', 'PartiallySucceeded', 'Failed')
    if ($a.Errors) { Write-EmtLog (T 'log.softErrors' $a.Errors) 'Error' }
    Write-EmtLog (T 'log.softStatus' $a.Status) $(if ($a.Status -eq 'Completed') { 'Ok' } else { 'Warn' })
    [pscustomobject]@{ Status = [string]$a.Status; Errors = [string]$a.Errors }
}

function Test-EmtHardDeleteScope {
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

function Invoke-EmtHardDelete {
    <#
        Permanent delete with Search-Mailbox -DeleteContent (incl. Recoverable Items), mailbox by mailbox.
        On cancel, it stops before the next mailbox and returns the partial result (Cancelled = true).
    #>
    param([Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query)
    $results = New-Object System.Collections.Generic.List[object]
    $cancelled = $false
    $i = 0
    foreach ($m in $Mailboxes) {
        if (Test-EmtCancelRequested) { $cancelled = $true; break }
        $i++; Set-EmtProgress ([int](60 * $i / $Mailboxes.Count)) (T 'progress.hard' $m)
        try {
            $r = Search-Mailbox -Identity $m -SearchQuery $Query -SearchDumpster -DeleteContent -Force `
                     -Confirm:$false -WarningAction SilentlyContinue -ErrorAction Stop
            Write-EmtLog (T 'log.hardDeleted' $m $r.ResultItemsCount) 'Ok'
            $results.Add([pscustomobject]@{ Mailbox = $m; Deleted = [int]$r.ResultItemsCount; Error = '' })
        } catch {
            Write-EmtLog (T 'log.hardError' $m $_.Exception.Message) 'Error'
            $results.Add([pscustomobject]@{ Mailbox = $m; Deleted = 0; Error = $_.Exception.Message })
        }
    }
    if ($cancelled) { Write-EmtLog (T 'log.hardStopped' $results.Count $Mailboxes.Count) 'Warn' }
    # ToArray(), not @(): PowerShell 5.1 can fail on @() of a generic List.
    [pscustomobject]@{ Results = $results.ToArray(); Cancelled = $cancelled }
}

function Get-EmtCopyStatus {
    <#
        Verification via Search-Mailbox (Compliance Search also sees Recoverable Items and would
        report the same numbers after a SoftDelete). Two counts per mailbox:
        Visible = folders the user sees; Total = including Recoverable Items.
    #>
    param([Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query, [switch] $TotalOnly)
    $i = 0
    foreach ($m in $Mailboxes) {
        Test-EmtCancel
        $i++; Set-EmtProgress ([int](100 * $i / $Mailboxes.Count)) (T 'progress.check' $m)
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

function Get-EmtSettledCopyStatus {
    <#
        Verification after a delete. The search index can lag behind the delete for some seconds.
        If the user can still see the message in a mailbox, it waits and examines that mailbox again.
    #>
    param(
        [Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query,
        [int] $Retries = 3, [int] $DelaySeconds = 15
    )
    $check = @(Get-EmtCopyStatus -Mailboxes $Mailboxes -Query $Query)
    for ($n = 1; $n -le $Retries; $n++) {
        $left = @($check | Where-Object { $_.Visible -gt 0 } | ForEach-Object Mailbox)
        if (-not $left.Count) { break }
        Write-EmtLog (T 'log.checkRetry' $left.Count $DelaySeconds $n $Retries)
        Wait-EmtSeconds $DelaySeconds
        $again = @(Get-EmtCopyStatus -Mailboxes $left -Query $Query)
        $check = @($check | Where-Object { $_.Mailbox -notin $left }) + $again
    }
    @($check | Sort-Object Mailbox)
}

function Get-EmtOldSearches {
    <#
        EMT searches older than the limit (EMT_ and the older MC_):
          * Compliance Searches EMT_yyyyMMdd_HHmmss;
          * the shadow In-Place eDiscovery searches EMT_yyyyMMdd_HHmmss-shadow that Exchange makes
            for each Compliance Search. Exchange can keep them after the Compliance Search is removed.
        The age comes from the name. Kind = ComplianceSearch or Shadow.
    #>
    param([int] $OlderThanDays = 7)
    $limit = (Get-Date).AddDays(-$OlderThanDays)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($s in @(Get-ComplianceSearch | Where-Object { Test-EmtSearchName $_.Name })) {
        $when = Get-EmtSearchNameTime $s.Name
        if ($when -and $when -lt $limit) { $list.Add([pscustomobject]@{ Name = $s.Name; When = $when; Kind = 'ComplianceSearch' }) }
    }
    if (Get-Command Get-MailboxSearch -ErrorAction SilentlyContinue) {
        foreach ($m in @(Get-MailboxSearch -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*-shadow' -and (Test-EmtSearchName $_.Name) })) {
            $when = Get-EmtSearchNameTime $m.Name
            if ($when -and $when -lt $limit) { $list.Add([pscustomobject]@{ Name = $m.Name; When = $when; Kind = 'Shadow' }) }
        }
    } else {
        Write-EmtLog (T 'log.noShadowCmd') 'Warn'
    }
    @($list | Sort-Object When, Name)
}

function Test-EmtNotFoundText {
    <# True for an Exchange "not found" message. "couldn.t": Exchange can write a typographic apostrophe (’). #>
    param([string] $Text)
    [bool]($Text -match 'couldn.t be found|could not be found|not found|не найден')
}

function Get-EmtComplianceSearchInfo {
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
    $exists = if ($s) { $true } elseif ($err -and -not (Test-EmtNotFoundText $err)) { $null } else { $false }
    [pscustomobject]@{
        Exists = $exists
        Status = $(if ($s -and @($s)[0].PSObject.Properties['Status']) { [string]@($s)[0].Status } else { '' })
        Error  = $err
    }
}

function Get-EmtComplianceSearchStatus {
    <# The status of a Compliance Search (Completed, InProgress, ...), or '' if it is not found or not readable. #>
    param([Parameter(Mandatory)] [string] $Name)
    (Get-EmtComplianceSearchInfo $Name).Status
}

function Test-EmtComplianceSearchExists {
    <#
        True if the Compliance Search exists. Uses -Identity: the full list of searches is not reliable
        on all servers. "Not found" = false. Another error = true (unknown: do not report a removal).
    #>
    param([Parameter(Mandatory)] [string] $Name)
    $i = Get-EmtComplianceSearchInfo $Name
    $i.Exists -ne $false
}

function Remove-EmtSearches {
    <#
        Removes EMT searches. A name with "-shadow" is a shadow In-Place eDiscovery search.
        For a Compliance Search, it also removes its actions and its shadow search, if Exchange kept it.
        Names that EMT does not make are skipped.
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
        if (-not (Test-EmtSearchName $n)) { Write-EmtLog (T 'log.cleanupSkip' $n) 'Warn'; continue }
        Test-EmtCancel
        # One failed search must not stop the cleanup of the others.
        try {
            if ($n -like '*-shadow') {
                if (-not $hasShadowCmd) { Write-EmtLog (T 'log.noShadowCmd') 'Warn'; continue }
                if (-not (& $shadowExists $n)) { Write-EmtLog (T 'log.shadowGone' $n); continue }
                Remove-MailboxSearch -Identity $n -Confirm:$false -ErrorAction Stop
                Write-EmtLog (T 'log.shadowRemoved' $n) 'Ok'
                continue
            }
            # The action first. Do NOT list all actions (Get-ComplianceSearchAction without -Identity):
            # on some servers one damaged action makes the full list fail with "Object reference not set".
            # EMT gives its purge action a known name: <search name>_Purge.
            # A search that still runs cannot be removed (Stop-ComplianceSearch does not always stop it).
            # Do not touch it or its shadow search: the cleanup can remove it later.
            $status = Get-EmtComplianceSearchStatus $n
            if ($status -in 'Starting', 'InProgress', 'Stopping', 'NotStarted') {
                Write-EmtLog (T 'log.stillRunning' $n $status) 'Warn'
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
                        Write-EmtLog (T 'log.shadowRemoved' $shadow) 'Ok'
                    } catch {
                        Write-EmtLog (T 'log.cleanupError' $shadow $_.Exception.Message) 'Warn'
                    }
                } elseif ($shadow -in $Names) {
                    Write-EmtLog (T 'log.shadowGone' $shadow)
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
                if (-not (Test-EmtComplianceSearchExists $n)) { $removed = $true }
                elseif ($rmErr) { $lastErr = @($rmErr)[0] }
            }
            if (-not $removed) {
                # Exchange removes the search in the background. Wait for it.
                for ($k = 1; $k -le [math]::Max(1, $RecheckTimes) -and -not $removed; $k++) {
                    Set-EmtProgress ([int](100 * $k / [math]::Max(1, $RecheckTimes))) (T 'progress.removeWait' $n)
                    if (-not (Test-EmtComplianceSearchExists $n)) { $removed = $true; break }
                    if ($k -lt $RecheckTimes) { Wait-EmtSeconds $RecheckSeconds }
                }
            }
            if (-not $removed) {
                $info = Get-EmtComplianceSearchInfo $n
                $seen = if ($info.Status) { $info.Status } elseif ($info.Error) { $info.Error } else { '?' }
                if ($lastErr) {
                    # Exchange returned an error and the search is still there: a real error.
                    throw (T 'err.stillThere' $lastErr.Exception.Message ($RecheckTimes * $RecheckSeconds))
                }
                # No error from Exchange: the remove is probably still running in the background.
                Write-EmtLog (T 'log.removePending' $n ($RecheckTimes * $RecheckSeconds) $seen) 'Warn'
                continue
            }
            Write-EmtLog (T 'log.cleanupRemoved' $n) 'Ok'
        } catch {
            Write-EmtLog (T 'log.cleanupError' $n $_.Exception.Message) 'Error'
        }
    }
}

#endregion

#region ---------- Trace: delivery status for each recipient ----------

function Get-EmtProp {
    <# The value of a property, or $null. StrictMode stops on a missing property: this function does not. #>
    param($Object, [Parameter(Mandatory)] [string] $Name)
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { $p.Value } else { $null }
}

function Get-EmtStatusRank {
    <# The order of the delivery statuses. A higher status replaces a lower one for the same recipient. #>
    param([string] $Status)
    switch ($Status) {
        'Delivered' { 10 } 'External' { 9 } 'Forwarded' { 8 } 'Failed' { 7 } 'GroupBlocked' { 6 } 'Dropped' { 5 }
        'Suppressed' { 4 } 'Deferred' { 3 } 'InTransit' { 2 } 'Group' { 1 } default { 0 }
    }
}

function Get-EmtErrorStatuses {
    <# The statuses that are a non-delivery: the recipient did not get the message, and Exchange did not plan it. #>
    'Failed', 'Dropped', 'GroupBlocked'
}

function Get-EmtDropStatus {
    <#
        The status of a DROP event. The server answer tells the cause.
        Exchange writes DROP also for usual steps with a success code (250 2.1.5):
          RESOLVER.GRP.Expanded, RESOLVER.GRP.NotRedirected - a group: the members get the message (Group);
          RESOLVER.FWD.Forwarded - the mailbox has forwarding: the message went to the forwarding address (Forwarded);
          CAT.DuplicateRecipient - a second copy for the same recipient: no status, the other events give it ($null);
          RESOLVER.GRP.BlockedMessage - the group did not accept the message (GroupBlocked, a non-delivery);
          other 2.x.x codes and agent texts - Exchange did not send this copy by a setting (Suppressed):
            RESOLVER.MSGTYPE.AF (automatic forward), RESOLVER.OOF.* (automatic reply), RESOLVER.MSGTYPE.NDR.
        A DROP with a 4.x.x or 5.x.x code, or without a cause, is a lost message (Dropped).
    #>
    param([string] $Detail)
    switch -Regex ($Detail) {
        'RESOLVER\.GRP\.(Expanded|NotRedirected)' { return 'Group' }
        'RESOLVER\.FWD\.Forwarded'                { return 'Forwarded' }
        'CAT\.DuplicateRecipient'                  { return $null }
        'RESOLVER\.GRP\.BlockedMessage'           { return 'GroupBlocked' }
    }
    if ($Detail -match '\b[45]\.\d{1,3}\.\d{1,3}\b') { return 'Dropped' }
    if ($Detail -match '\b2\.\d{1,3}\.\d{1,3}\b' -or $Detail -match '\[Agent:') { return 'Suppressed' }
    'Dropped'
}

function Get-EmtEventStatus {
    <#
        The delivery status that one tracking event gives to one recipient, or $null (no status).
        Kind = Internal / External / Group / Other (from Resolve-EmtRecipients).
        SEND to an internal recipient is a transfer to another Exchange server, not the end of the delivery.
        SENDEXTERNAL is a send through a Send connector to a server outside the organization
        (for example a hidden copy that a transport rule adds). This is the end of the delivery for Exchange.
        DUPLICATEDELIVER is a delivery of a second copy to a mailbox that has the message.
        DROP: Get-EmtDropStatus reads the server answer ($Detail). Many DROP events are not errors.
    #>
    param([string] $EventId, [string] $Kind, [string] $Detail)
    switch ($EventId.ToUpperInvariant()) {
        'DELIVER'          { 'Delivered' }
        'DUPLICATEDELIVER' { 'Delivered' }
        'SEND'          { if ($Kind -eq 'Internal') { 'InTransit' } else { 'External' } }
        'SENDEXTERNAL'  { 'External' }
        'FAIL'          { 'Failed' }
        'DROP'          { Get-EmtDropStatus $Detail }
        'POISONMESSAGE' { 'Dropped' }
        'DEFER'         { 'Deferred' }
        'EXPAND'        { 'InTransit' }
        'REDIRECT'      { 'InTransit' }
        'HAREDIRECT'    { 'InTransit' }
        'TRANSFER'      { 'InTransit' }
        'RESOLVE'       { 'InTransit' }
        'RECEIVE'       { 'Unknown' }
        'SUBMIT'        { 'Unknown' }
        default         { $null }
    }
}

function Get-EmtStatusHint {
    <#
        A short explanation of an SMTP status text (enhanced status code x.y.z). Empty text: no explanation.
        Example: "550 5.1.1 RESOLVER.ADR.RecipNotFound; not found" -> "The address does not exist".
    #>
    param([string] $Text)
    if (-not $Text) { return '' }
    # The causes of a DROP event with a success code (Get-EmtDropStatus).
    $code = if ($Text -match '\b([245]\.\d{1,3}\.\d{1,3})\b') { $Matches[1] } else { '' }
    $special = switch -Regex ($Text) {
        'RESOLVER\.GRP\.(Expanded|NotRedirected)' { 'hint.grpExpanded'; break }
        'RESOLVER\.FWD\.Forwarded'                { 'hint.forwarded'; break }
        'RESOLVER\.MSGTYPE\.AF'                   { 'hint.autoForward'; break }
        'RESOLVER\.OOF\.'                         { 'hint.oof'; break }
        'RESOLVER\.MSGTYPE\.NDR'                  { 'hint.ndrSuppressed'; break }
        'CAT\.DuplicateRecipient'                  { 'hint.duplicate'; break }
        'RESOLVER\.GRP\.BlockedMessage'           { 'hint.grpBlocked'; break }
        '\[Agent:'                                 { 'hint.suppressed'; break }
    }
    # A text of an agent has no code: the hint starts without "code:".
    if ($special) { return ((T $special $code) -replace '^:\s*', '') }
    if ($Text -notmatch '\b(?<c>[245])\.(?<s>\d{1,3})\.(?<d>\d{1,3})\b') { return '' }
    $code = "$($Matches['c']).$($Matches['s']).$($Matches['d'])"
    $key = switch -Regex ($code) {
        '^2\.'               { ''; break }
        '^5\.1\.(0|1|10)$'   { 'hint.noAddress'; break }
        '^5\.1\.8$'          { 'hint.badSender'; break }
        '^5\.2\.2$'          { 'hint.full'; break }
        '^5\.(2\.3|3\.4)$'   { 'hint.tooBig'; break }
        '^[45]\.4\.7$'       { 'hint.expired'; break }
        '^5\.4\.(1|4|6|14)$' { 'hint.route'; break }
        '^4\.4\.\d+$'        { 'hint.noConnect'; break }
        '^5\.7\.\d+$'        { 'hint.denied'; break }
        '^4\.3\.\d+$'        { 'hint.resources'; break }
        '^5\.6\.\d+$'        { 'hint.content'; break }
        '^4\.7\.\d+$'        { 'hint.tempPolicy'; break }
        '^5\.'               { 'hint.permanent'; break }
        '^4\.'               { 'hint.temporary'; break }
        default              { '' }
    }
    if ($key) { T $key $code } else { '' }
}

function Get-EmtDeliveryFolder {
    <#
        The folder of a delivery, or an empty text. Exchange writes the folder name in the recipient status
        of a DELIVER event when the message goes to a folder other than Inbox (usually an inbox rule).
        Only the name, not the path. An SMTP answer (a code x.y.z, "[{LED=...}]") is not a folder.
    #>
    param([string] $EventId, [string] $Detail)
    if ($EventId.ToUpperInvariant() -notin 'DELIVER', 'DUPLICATEDELIVER') { return '' }
    $d = ([string]$Detail).Trim()
    if (-not $d) { return '' }
    if ($d -match '\b[245]\.\d{1,3}\.\d{1,3}\b' -or $d.StartsWith('[{') -or $d -match '^\d{3}\b') { return '' }
    $d
}

function Get-EmtDeliveryReport {
    <#
        One row for each message and recipient: the final delivery status from the tracking events.
        The status with the highest rank wins (DELIVER after DEFER = Delivered). With the same rank, the later event wins.
        Kinds: address -> Internal / External / Group / Other.
        Folder: the folder of the delivery (Get-EmtDeliveryFolder). Then Detail is empty: the folder is not a server answer.
    #>
    param([object[]] $Events = @(), [hashtable] $Kinds = @{})
    $state = [ordered]@{}
    foreach ($e in @($Events | Where-Object { $_ -and $_.MessageId } | Sort-Object Timestamp)) {
        $rcpts = @(Get-EmtProp $e 'Recipients')
        $stats = @(Get-EmtProp $e 'RecipientStatus')
        for ($i = 0; $i -lt $rcpts.Count; $i++) {
            $a = ([string]$rcpts[$i]).Trim().ToLowerInvariant()
            if (-not $a) { continue }
            $kind = if ($Kinds.ContainsKey($a)) { [string]$Kinds[$a] } else { '' }
            $detail = if ($i -lt $stats.Count) { [string]$stats[$i] } else { '' }
            $code = Get-EmtEventStatus -EventId ([string]$e.EventId) -Kind $kind -Detail $detail
            if (-not $code) { continue }
            $key = "$($e.MessageId)|$a"
            $old = if ($state.Contains($key)) { $state[$key] } else { $null }
            if ($old -and (Get-EmtStatusRank $code) -lt (Get-EmtStatusRank $old.Status)) { continue }
            $folder = Get-EmtDeliveryFolder -EventId ([string]$e.EventId) -Detail $detail
            $state[$key] = [pscustomobject]@{
                MessageId = [string]$e.MessageId
                Recipient = $a
                Kind      = $kind
                Status    = $code
                Time      = $e.Timestamp
                Server    = [string](Get-EmtProp $e 'ServerHostname')
                Event     = [string]$e.EventId
                Folder    = $folder
                Detail    = $(if ($folder) { '' } else { $detail })
            }
        }
    }
    foreach ($row in $state.Values) {
        # A group address has no own delivery: the members get the message.
        if ($row.Kind -eq 'Group' -and (Get-EmtStatusRank $row.Status) -le (Get-EmtStatusRank 'InTransit')) { $row.Status = 'Group' }
        $row | Add-Member -NotePropertyName Hint -NotePropertyValue (Get-EmtStatusHint $row.Detail) -PassThru
    }
}

function Get-EmtStatusOrder {
    <# The order of the statuses in the status line of a message. #>
    'Delivered', 'External', 'Forwarded', 'Failed', 'GroupBlocked', 'Dropped', 'Suppressed', 'Deferred', 'InTransit', 'Group', 'Unknown'
}

function Get-EmtMessageStatusLine {
    <# The status line of one message from its report rows: "Delivered: 5, External: 1, Failed: 1". #>
    param([object[]] $Rows = @())
    $order = Get-EmtStatusOrder
    $parts = foreach ($s in $order) {
        $n = @($Rows | Where-Object { $_.Status -eq $s }).Count
        if ($n) { T "line.$s" $n }
    }
    if ($parts) { @($parts) -join ', ' } else { T 'line.none' }
}

function Get-EmtNoResultHints {
    <# What to examine when the trace finds nothing. #>
    param([Parameter(Mandatory)] [hashtable] $Criteria, [datetime] $Now = (Get-Date))
    $c = Complete-EmtCriteria $Criteria
    $hints = @()
    if ($c.Start -lt $Now.AddDays(-30)) { $hints += (T 'nores.tooOld') } else { $hints += (T 'nores.period') }
    if (@($c.Senders).Count) { $hints += (T 'nores.sender') }
    if ($c.Recipient)        { $hints += (T 'nores.recipient') }
    if ($c.Subject)          { $hints += (T 'nores.subject') }
    $hints += (T 'nores.notSent')
    $hints
}

function Invoke-EmtTrace {
    <#
        Trace: only the message tracking log, no mailbox search. Seconds, not minutes.
        Needs only the Message Tracking role. Returns the messages, the delivery report and the events.
        -Partial with a recipient: no second pass by message ID. The search is fast for many messages.
        The report has only the rows of this recipient. Get-EmtMessageDetail gives the other recipients of a message.
    #>
    param([Parameter(Mandatory)] [hashtable] $Criteria, [switch] $Partial)
    $Criteria = Complete-EmtCriteria $Criteria
    $problems = @(Test-EmtCriteria $Criteria -ForTrace)
    if ($problems.Count) { throw ($problems -join ' ') }
    $caps = Get-EmtCapabilities
    if (-not $caps.Tracking) { throw (T 'log.noTrackRole') }
    # A partial search only with a recipient filter: without it, the search has no second pass.
    $partialSearch = [bool]($Partial -and $Criteria.Recipient)
    if ($partialSearch) { Write-EmtLog (T 'log.tracePartial' $Criteria.Recipient) }
    $timer = [System.Diagnostics.Stopwatch]::StartNew()

    $tracking = Find-EmtTrackingEvents -Senders @($Criteria.Senders) -Subject $Criteria.Subject -Recipient $Criteria.Recipient `
                    -MessageId $Criteria.MessageId -Start $Criteria.Start -End $Criteria.End -CompleteRecipients:(-not $partialSearch)
    $events = @($tracking.Events)
    $summary = @(Get-EmtMessageSummary -Events $events -ServerNames $tracking.Servers)

    $addresses = @($events | ForEach-Object { @(Get-EmtProp $_ 'Recipients') } | Where-Object { $_ } |
                   ForEach-Object { ([string]$_).ToLowerInvariant() } | Sort-Object -Unique)
    $recipients = @(Resolve-EmtRecipients -Addresses $addresses -AcceptedDomains (Get-EmtAcceptedDomains))
    $kinds = @{}
    foreach ($x in $recipients) { $kinds[$x.Address] = $x.Kind }
    $report = @(Get-EmtDeliveryReport -Events $events -Kinds $kinds)
    # A partial search knows the full delivery of the searched recipient only. The other rows are not complete.
    if ($partialSearch) { $report = @($report | Where-Object { $_.Recipient -eq ([string]$Criteria.Recipient).ToLowerInvariant() }) }

    # The recipients of each message in one pass: a search in the full report for each message is slow.
    $byMsg = @{}
    foreach ($r in $report) {
        if (-not $byMsg.ContainsKey($r.MessageId)) { $byMsg[$r.MessageId] = New-Object System.Collections.ArrayList }
        [void]$byMsg[$r.MessageId].Add($r)
    }
    $errorStatuses = @(Get-EmtErrorStatuses)
    $messages = @(foreach ($m in $summary) {
        # @( ) keeps an array: PowerShell makes one row of an array a single object, and it has no Count in strict mode.
        $rows = @(if ($byMsg.ContainsKey($m.MessageId)) { $byMsg[$m.MessageId].ToArray() })
        [pscustomobject]@{
            Time            = $m.Time
            Sender          = $m.Sender
            Subject         = $m.Subject
            Recipients      = $rows.Count
            Problems        = @($rows | Where-Object { $_.Status -in $errorStatuses }).Count
            Pending         = @($rows | Where-Object { $_.Status -in 'Deferred', 'InTransit', 'Unknown' }).Count
            StatusLine      = Get-EmtMessageStatusLine $rows
            ServerGenerated = $m.ServerGenerated
            MessageId       = $m.MessageId
        }
    })
    $failed = @($report | Where-Object { $_.Status -in $errorStatuses }).Count
    Write-EmtLog (T 'log.traceDone' $messages.Count $report.Count $failed ([int]$timer.Elapsed.TotalSeconds)) 'Ok'
    Set-EmtProgress 100 (T 'progress.done')
    [pscustomobject]@{
        Criteria       = $Criteria
        Messages       = $messages
        Report         = $report
        Events         = $events
        Recipients     = $recipients
        TrackingErrors = @($tracking.Errors)
        NoResultHints  = $(if ($messages.Count) { @() } else { @(Get-EmtNoResultHints $Criteria) })
        Seconds        = [int]$timer.Elapsed.TotalSeconds
        Partial        = $partialSearch
    }
}

function Get-EmtMessageDetail {
    <#
        The full data of messages after a partial message search (Invoke-EmtTrace -Partial):
        all events, all recipients and the delivery report. Two calls for each message with two servers.
    #>
    param(
        [Parameter(Mandatory)] [string[]] $MessageIds, [Parameter(Mandatory)] [datetime] $Start,
        [Parameter(Mandatory)] [datetime] $End
    )
    $servers = @(Get-TransportService | ForEach-Object { [string]$_.Name })
    $pass = Get-EmtEventsByMessageId -MessageIds $MessageIds -Start $Start -End $End -Servers $servers -ProgressFrom 5 -ProgressTo 80
    $errs = @($pass.Errors | Sort-Object -Unique)
    foreach ($t in $errs) { Write-EmtLog (T 'log.trackError' $t) 'Error' }
    $events = @(Select-EmtUniqueEvents -Events $pass.Events -MessageIds $MessageIds)
    $addresses = @($events | ForEach-Object { @(Get-EmtProp $_ 'Recipients') } | Where-Object { $_ } |
                   ForEach-Object { ([string]$_).ToLowerInvariant() } | Sort-Object -Unique)
    $recipients = @(Resolve-EmtRecipients -Addresses $addresses -AcceptedDomains (Get-EmtAcceptedDomains))
    $kinds = @{}
    foreach ($x in $recipients) { $kinds[$x.Address] = $x.Kind }
    $report = @(Get-EmtDeliveryReport -Events $events -Kinds $kinds)
    Set-EmtProgress 100 (T 'progress.done')
    [pscustomobject]@{ MessageIds = @($MessageIds); Events = $events; Report = $report; Recipients = $recipients; Errors = $errs }
}

#endregion

#region ---------- Mailbox check (level 1: why a delivered message can be missing) ----------

function Invoke-EmtReadOnly {
    <#
        Runs one read-only Exchange cmdlet. Returns Value and Error. No "-ErrorAction Stop" (see the Exchange 2019 rule).
        A missing cmdlet (no role) gives the error "no permission".
    #>
    param([Parameter(Mandatory)] [string] $Command, [Parameter(Mandatory)] [hashtable] $Params)
    if (-not (Get-Command $Command -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{ Value = $null; Error = (T 'find.noCmd' $Command) }
    }
    $ev = $null
    $val = $null
    try {
        $val = & $Command @Params -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ErrorVariable ev
    } catch { $ev = @($_) }
    $err = (@($ev) | Where-Object { $_ } | ForEach-Object { [string]$_.Exception.Message }) -join ' '
    [pscustomobject]@{ Value = $val; Error = $err }
}

function Test-EmtAddressInList {
    <# True if the address or its domain is in a list of senders and domains (junk e-mail lists). #>
    param([string] $Address, [object[]] $List = @())
    if (-not $Address) { return $false }
    $a = $Address.ToLowerInvariant()
    $d = ($a -split '@')[-1]
    foreach ($x in $List) {
        $v = ([string]$x).Trim().ToLowerInvariant()
        if ($v -and ($v -eq $a -or $v -eq $d -or $v -eq "@$d")) { return $true }
    }
    $false
}

function Get-EmtRuleMatch {
    <#
        Does an inbox rule apply to the message? Returns All (no conditions), Yes, No or Unknown
        (only conditions that EMT cannot examine, for example words in the body).
    #>
    param($Rule, [string] $Sender, [string] $Subject)
    $s = ([string]$Sender).ToLowerInvariant()
    $subj = ([string]$Subject).ToLowerInvariant()
    $from      = @(@(Get-EmtProp $Rule 'From') | Where-Object { $_ })
    $fromWords = @(@(Get-EmtProp $Rule 'FromAddressContainsWords') | Where-Object { $_ })
    $subWords  = @(@(@(Get-EmtProp $Rule 'SubjectContainsWords') + @(Get-EmtProp $Rule 'SubjectOrBodyContainsWords')) | Where-Object { $_ })
    $bodyWords = @(@(Get-EmtProp $Rule 'BodyContainsWords') | Where-Object { $_ })
    $known = $from.Count + $fromWords.Count + $subWords.Count
    if (-not $known -and -not $bodyWords.Count) { return 'All' }
    foreach ($f in $from)      { if ($s -and ([string]$f).ToLowerInvariant().Contains($s)) { return 'Yes' } }
    foreach ($w in $fromWords) { if ($s -and $s.Contains(([string]$w).ToLowerInvariant())) { return 'Yes' } }
    foreach ($w in $subWords)  { if ($subj -and $subj.Contains(([string]$w).ToLowerInvariant())) { return 'Yes' } }
    if (-not $known) { return 'Unknown' }
    'No'
}

function Get-EmtMailboxFindings {
    <#
        The findings of the mailbox check. Data: Rules, Junk, Box (Get-Mailbox), Stats (Get-MailboxStatistics)
        and an error text for each. Returns Severity (Alert / Warn / Info), Area and Text.
    #>
    param([Parameter(Mandatory)] $Data, [string] $Sender, [string] $Subject)
    $out = New-Object System.Collections.Generic.List[object]
    $add = { param($sev, $area, $text) $out.Add([pscustomobject]@{ Severity = $sev; Area = $area; Text = $text }) }

    # ---- Inbox rules ----
    if ($Data.RulesError) { & $add 'Warn' 'Rules' (T 'find.notChecked' $Data.RulesError) }
    foreach ($r in @($Data.Rules | Where-Object { $_ })) {
        if ((Get-EmtProp $r 'Enabled') -eq $false) { continue }
        $name = [string](Get-EmtProp $r 'Name')
        $move = [string](Get-EmtProp $r 'MoveToFolder')
        $del  = (Get-EmtProp $r 'DeleteMessage') -eq $true
        $fwd  = @(@(@(Get-EmtProp $r 'ForwardTo') + @(Get-EmtProp $r 'RedirectTo') + @(Get-EmtProp $r 'ForwardAsAttachmentTo')) |
                  Where-Object { $_ } | ForEach-Object { [string]$_ })
        if (-not $move -and -not $del -and -not $fwd.Count) { continue }
        $match = Get-EmtRuleMatch -Rule $r -Sender $Sender -Subject $Subject
        $sev = if ($match -in 'All', 'Yes') { 'Alert' } elseif ($match -eq 'Unknown') { 'Warn' } else { 'Info' }
        $suffix = T "find.match$match"
        if ($del)       { & $add $sev 'Rules' ((T 'find.ruleDelete' $name) + ' ' + $suffix) }
        if ($move)      { & $add $sev 'Rules' ((T 'find.ruleMove' $name ($move -replace '^[^:]*:\\', '')) + ' ' + $suffix) }
        if ($fwd.Count) { & $add $sev 'Rules' ((T 'find.ruleForward' $name ($fwd -join ', ')) + ' ' + $suffix) }
    }

    # ---- Junk e-mail ----
    if ($Data.JunkError) { & $add 'Warn' 'Junk' (T 'find.notChecked' $Data.JunkError) }
    elseif ($Data.Junk) {
        if ((Get-EmtProp $Data.Junk 'Enabled') -eq $false) { & $add 'Info' 'Junk' (T 'find.junkOff') }
        if (Test-EmtAddressInList $Sender @(Get-EmtProp $Data.Junk 'BlockedSendersAndDomains')) { & $add 'Alert' 'Junk' (T 'find.junkBlocked' $Sender) }
        if (Test-EmtAddressInList $Sender @(Get-EmtProp $Data.Junk 'TrustedSendersAndDomains')) { & $add 'Info' 'Junk' (T 'find.junkTrusted' $Sender) }
    }

    # ---- Mailbox forwarding ----
    if ($Data.BoxError) { & $add 'Warn' 'Forwarding' (T 'find.notChecked' $Data.BoxError) }
    elseif ($Data.Box) {
        $target = [string](Get-EmtProp $Data.Box 'ForwardingSmtpAddress')
        if (-not $target) { $target = [string](Get-EmtProp $Data.Box 'ForwardingAddress') }
        $target = $target -replace '^smtp:', ''
        if ($target) {
            if ((Get-EmtProp $Data.Box 'DeliverToMailboxAndForward') -eq $true) { & $add 'Warn' 'Forwarding' (T 'find.fwdKeep' $target) }
            else { & $add 'Alert' 'Forwarding' (T 'find.fwdNoCopy' $target) }
        }
    }

    # ---- Quota ----
    if ($Data.StatsError) { & $add 'Warn' 'Quota' (T 'find.notChecked' $Data.StatsError) }
    elseif ($Data.Stats -and $Data.Box) {
        $usedKB = ConvertFrom-EmtSizeText ([string](Get-EmtProp $Data.Stats 'TotalItemSize'))
        $usedText = '{0:N1} MB' -f ($usedKB / 1KB)
        $quota = [string](Get-EmtProp $Data.Box 'ProhibitSendReceiveQuota')
        if ((Get-EmtProp $Data.Box 'UseDatabaseQuotaDefaults') -eq $true) { & $add 'Info' 'Quota' (T 'find.quotaDb' $usedText) }
        elseif (-not $quota -or $quota -match 'Unlimited') { & $add 'Info' 'Quota' (T 'find.quotaNone' $usedText) }
        else {
            $quotaKB = ConvertFrom-EmtSizeText $quota
            if ($quotaKB -gt 0) {
                $pct = [math]::Round(100 * $usedKB / $quotaKB)
                if ($pct -ge 100)    { & $add 'Alert' 'Quota' (T 'find.quotaFull' $usedText) }
                elseif ($pct -ge 90) { & $add 'Warn' 'Quota' (T 'find.quotaHigh' $pct $usedText) }
                else                 { & $add 'Info' 'Quota' (T 'find.quotaOk' $pct $usedText) }
            }
        }
    }

    if (-not @($out | Where-Object { $_.Severity -in 'Alert', 'Warn' }).Count) { & $add 'Info' 'Note' (T 'find.nothing') }
    & $add 'Info' 'Note' (T 'find.note')
    # Not "@($out)": PowerShell 5.1 fails on @() of a generic List ("Argument types do not match").
    $out.ToArray()
}

function Invoke-EmtMailboxCheck {
    <# Reads the rules, the junk e-mail settings, the forwarding and the quota of one mailbox. Read-only. #>
    param([Parameter(Mandatory)] [string] $Mailbox, [string] $Sender, [string] $Subject)
    Write-EmtLog (T 'log.checkMailbox' $Mailbox)
    $rules = Invoke-EmtReadOnly 'Get-InboxRule' @{ Mailbox = $Mailbox }
    Set-EmtProgress 25 (T 'progress.checkMailbox' $Mailbox)
    $junk  = Invoke-EmtReadOnly 'Get-MailboxJunkEmailConfiguration' @{ Identity = $Mailbox }
    Set-EmtProgress 50 (T 'progress.checkMailbox' $Mailbox)
    $box   = Invoke-EmtReadOnly 'Get-Mailbox' @{ Identity = $Mailbox }
    Set-EmtProgress 75 (T 'progress.checkMailbox' $Mailbox)
    $stats = Invoke-EmtReadOnly 'Get-MailboxStatistics' @{ Identity = $Mailbox }
    $data = [pscustomobject]@{
        Rules = @($rules.Value); RulesError = $rules.Error
        Junk  = @($junk.Value)[0]; JunkError = $junk.Error
        Box   = @($box.Value)[0];  BoxError = $box.Error
        Stats = @($stats.Value)[0]; StatsError = $stats.Error
    }
    Set-EmtProgress 100 (T 'progress.done')
    $findings = @(Get-EmtMailboxFindings -Data $data -Sender $Sender -Subject $Subject)
    $alerts = @($findings | Where-Object Severity -eq 'Alert').Count
    Write-EmtLog (T 'log.checkMailboxDone' $Mailbox $alerts) $(if ($alerts) { 'Warn' } else { 'Ok' })
    # Data, Sender and Subject: the window makes the findings again after a language change.
    [pscustomobject]@{ Mailbox = $Mailbox; Findings = $findings; Data = $data; Sender = $Sender; Subject = $Subject }
}

#endregion

#region ---------- Entry point ----------

function Invoke-EmtOperation {
    <# Single entry point used by the GUI's background runspace. #>
    param([Parameter(Mandatory)] [string] $Name, [hashtable] $Params = @{})
    switch ($Name) {
        'Connect'      { Connect-EmtExchange -Server $Params.Server }
        'Trace'        { Invoke-EmtTrace -Criteria $Params.Criteria -Partial:([bool]$Params.Partial) }
        'TraceDetail'  { Get-EmtMessageDetail -MessageIds $Params.MessageIds -Start $Params.Start -End $Params.End }
        'CheckMailbox' { Invoke-EmtMailboxCheck -Mailbox $Params.Mailbox -Sender $Params.Sender -Subject $Params.Subject }
        'Resolve'      { @(Resolve-EmtAddressInputs -Inputs $Params.Inputs) }
        'Search'       { Invoke-EmtSearch -Criteria $Params.Criteria -TimeoutMinutes $Params.TimeoutMinutes }
        'SoftDelete'   { $searchName = $Params.SearchName
                         if (-not $searchName) {
                             # The quick search has no Compliance Search. Purge needs one: make it now.
                             $searchName = New-EmtSearchName
                             Write-EmtLog (T 'log.softSearch' $searchName @($Params.Mailboxes).Count)
                             [void](New-EmtComplianceSearchRun -Name $searchName -Locations $Params.Mailboxes `
                                        -Query $Params.Query -TimeoutMinutes $Params.TimeoutMinutes)
                         }
                         $r = Invoke-EmtSoftDelete -SearchName $searchName -TimeoutMinutes $Params.TimeoutMinutes
                         [pscustomobject]@{
                             Delete = $r; SearchName = $searchName
                             Check = @(Get-EmtSettledCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query)
                         } }
        'HardDelete'   { Write-EmtLog (T 'log.precheck')
                         $pre = @(Get-EmtCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query -TotalOnly)
                         $bad = @(Test-EmtHardDeleteScope -Status $pre -Expected $Params.Expected)
                         if ($bad.Count) {
                             $list = ($bad | ForEach-Object { T 'scope.item' $_.Mailbox $_.Expected $_.Actual $_.Reason }) -join '; '
                             throw (T 'err.precheck' $list)
                         }
                         Write-EmtLog (T 'log.precheckOk') 'Ok'
                         $d = Invoke-EmtHardDelete -Mailboxes $Params.Mailboxes -Query $Params.Query
                         $check = if ($d.Cancelled) { @() } else { @(Get-EmtSettledCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query) }
                         [pscustomobject]@{ Delete = @($d.Results); Cancelled = $d.Cancelled; Check = $check } }
        'Verify'       { [pscustomobject]@{ Check = @(Get-EmtCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query) } }
        'ListOld'      { Get-EmtOldSearches -OlderThanDays $Params.Days }
        'RemoveOld'    { Remove-EmtSearches -Names $Params.Names }
        default        { throw (T 'err.unknownOp' $Name) }
    }
}

#endregion
