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

function Get-McSearchPrefix { 'MC_' }

function Get-McExchangeCommandNames {
    @(
        'Get-TransportService', 'Get-MessageTrackingLog', 'Get-Recipient',
        'New-ComplianceSearch', 'Start-ComplianceSearch', 'Stop-ComplianceSearch',
        'Get-ComplianceSearch', 'Remove-ComplianceSearch',
        'New-ComplianceSearchAction', 'Get-ComplianceSearchAction', 'Remove-ComplianceSearchAction',
        'Search-Mailbox'
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

function Test-McCancel {
    if ((Get-Variable -Name McSync -Scope Global -ErrorAction SilentlyContinue) -and
        $global:McSync -and $global:McSync.Cancel) {
        throw [System.OperationCanceledException]::new('Операция отменена пользователем.')
    }
}

#endregion

#region ---------- Pure helpers ----------

function Split-McAddressList {
    <# "a@x; b@x, c@x" -> lower-case unique array. #>
    param([string] $Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    @($Text -split '[,;\s]+' | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLowerInvariant() } |
        Sort-Object -Unique)
}

function Get-McKqlDateRange {
    <#
        KQL in Compliance Search filters by whole days only.
        Returns From (inclusive) and To (exclusive) as yyyy-MM-dd.
        An End exactly at midnight does not pull in the next day.
    #>
    param([Parameter(Mandatory)] [datetime] $Start, [Parameter(Mandatory)] [datetime] $End)
    if ($End -le $Start) { throw 'Конец периода должен быть позже начала.' }
    $endDay = if ($End -eq $End.Date) { $End.Date } else { $End.Date.AddDays(1) }
    [pscustomobject]@{
        From = $Start.Date.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        To   = $endDay.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    }
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
        [Parameter(Mandatory)] [datetime] $End
    )
    $parts = New-Object System.Collections.Generic.List[string]
    $senders = @($Senders | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    if ($senders.Count) { $parts.Add('(' + (($senders | ForEach-Object { "from:$_" }) -join ' OR ') + ')') }
    if ($Subject)    { $parts.Add('(subject:' + (ConvertTo-McKqlPhrase $Subject) + ')') }
    if ($Attachment) { $parts.Add('(attachment:' + (ConvertTo-McKqlPhrase $Attachment) + ')') }
    if ($Keywords)   { $parts.Add('(' + (ConvertTo-McKqlPhrase $Keywords) + ')') }
    if ($parts.Count -eq 0) {
        throw 'Запрос без отправителя, темы, вложения и слов совпал бы со всеми письмами за период — задайте хотя бы один критерий.'
    }
    $range = Get-McKqlDateRange -Start $Start -End $End
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
    if (-not $hasKey) { $problems += 'Укажите отправителя, тему, вложение, слова или MessageId.' }
    if ($Criteria.End -le $Criteria.Start) { $problems += 'Конец периода должен быть позже начала.' }
    if (($Criteria.End - $Criteria.Start).TotalDays -gt 31) {
        $problems += 'Период больше 31 дня — сузьте его (журнал доставки по умолчанию хранится 30 дней).'
    }
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
    $Events | Where-Object { $_ } | Group-Object MessageId | ForEach-Object {
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
            Note       = $(if ($auto) { 'создано сервером — вероятно, правило/автопересылка' } else { '' })
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
    Write-McLog "Подключение к http://$Server/PowerShell/ (Kerberos)..."
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/PowerShell/" `
                   -Authentication Kerberos -ErrorAction Stop
    $module = Import-PSSession $session -CommandName (Get-McExchangeCommandNames) -DisableNameChecking `
                  -AllowClobber -ErrorAction Stop -WarningAction SilentlyContinue
    Import-Module $module -Global -DisableNameChecking -WarningAction SilentlyContinue
    $global:McExSession = $session
    $caps = Get-McCapabilities
    Write-McLog "Подключено к $Server." 'Ok'
    $caps
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
        Keeps messages whose subject contains $Subject (if given), then pulls every event of those MessageIds.
    #>
    param(
        [string[]] $Senders = @(), [string] $Subject, [string] $Recipient, [string] $MessageId,
        [Parameter(Mandatory)] [datetime] $Start, [Parameter(Mandatory)] [datetime] $End,
        [switch] $IncludeForwards
    )
    $servers = @(Get-TransportService | ForEach-Object { [string]$_.Name })
    Write-McLog "Журнал доставки: серверов $($servers.Count)."
    $common = @{ Start = $Start; End = $End; ResultSize = 'Unlimited'; ErrorAction = 'Continue' }
    if ($Recipient) { $common.Recipients = $Recipient }

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
            $n++; Set-McProgress ([int](30 * $n / $total)) "Журнал доставки: $srv"
            $p = $common + $q
            foreach ($e in @(Get-MessageTrackingLog -Server $srv @p)) { if ($e) { $raw.Add($e) } }
        }
    }

    $hits = @($raw | Where-Object { -not $Subject -or $_.MessageSubject -like "*$Subject*" })
    if (-not $IncludeForwards -and $Senders.Count) {
        $hits = @($hits | Where-Object { ([string]$_.Sender).ToLowerInvariant() -in $Senders })
    }
    $ids = @($hits | Select-Object -ExpandProperty MessageId -Unique)
    Write-McLog "Писем (MessageId) в журнале: $($ids.Count)."

    $all = New-Object System.Collections.Generic.List[object]
    foreach ($id in $ids) {
        foreach ($srv in $servers) {
            Test-McCancel
            foreach ($e in @(Get-MessageTrackingLog -Server $srv -MessageId $id -Start $Start -End $End.AddDays(1) `
                                 -ResultSize Unlimited -ErrorAction Continue)) { if ($e) { $all.Add($e) } }
        }
    }
    [pscustomobject]@{ Events = @($all | Sort-Object Timestamp); Servers = $servers }
}

function Resolve-McRecipients {
    <# Classifies addresses: Internal mailbox / External / Group / Other. #>
    param([string[]] $Addresses = @())
    $i = 0
    foreach ($a in ($Addresses | Sort-Object -Unique)) {
        Test-McCancel
        $i++; Set-McProgress (30 + [int](10 * $i / [math]::Max(1, $Addresses.Count))) "Получатели: $a"
        $r = $null
        try { $r = Get-Recipient -Identity $a -ErrorAction Stop } catch { $r = $null }
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

function Wait-McComplianceSearch {
    param([Parameter(Mandatory)] [string] $Name, [int] $TimeoutMinutes = 60, [int] $ProgressFrom = 40, [int] $ProgressTo = 95)
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ($true) {
        Start-Sleep -Seconds 5
        if ((Get-Variable -Name McSync -Scope Global -ErrorAction SilentlyContinue) -and $global:McSync -and $global:McSync.Cancel) {
            Write-McLog "Остановка поиска $Name..." 'Warn'
            Stop-ComplianceSearch -Identity $Name -Confirm:$false -ErrorAction SilentlyContinue
            Test-McCancel
        }
        $s = Get-ComplianceSearch -Identity $Name
        $pct = [int]$s.JobProgress
        Set-McProgress ($ProgressFrom + [int](($ProgressTo - $ProgressFrom) * $pct / 100)) "Поиск в ящиках: $($s.Status), $pct%"
        if ($s.Status -in 'Completed', 'Stopped', 'Failed') { break }
        if ((Get-Date) -gt $deadline) { throw "Поиск $Name не завершился за $TimeoutMinutes мин." }
    }
    if ($s.Status -ne 'Completed') { throw "Поиск $Name завершился со статусом $($s.Status)." }
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
    if (-not $caps.Search) { throw 'Нет прав на Compliance Search (роль Mailbox Search / Discovery Management).' }

    # ---- 1. Message tracking ----
    $tracking = [pscustomobject]@{ Events = @(); Servers = @() }
    if ($caps.Tracking) {
        $tracking = Find-McTrackingEvents -Senders $senders -Subject $Criteria.Subject -Recipient $Criteria.Recipient `
                        -MessageId $Criteria.MessageId -Start $Criteria.Start -End $Criteria.End `
                        -IncludeForwards:([bool]$Criteria.IncludeForwards)
    } else {
        Write-McLog 'Нет прав на журнал доставки (Get-MessageTrackingLog) — шаг пропущен, поиск будет по всем ящикам.' 'Warn'
    }
    $events  = @($tracking.Events)
    $summary = @(Get-McMessageSummary -Events $events -ServerNames $tracking.Servers)
    if (-not $events.Count) { Write-McLog 'В журнале доставки ничего не найдено.' 'Warn' }
    foreach ($m in @($summary | Where-Object Note)) { Write-McLog "Письмо $($m.MessageId) от $($m.Sender): $($m.Note)." 'Warn' }

    # ---- 2. Recipients ----
    $addresses = @($events | ForEach-Object { $_.Recipients } | Where-Object { $_ } |
                   ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $trackSenders = @($events | ForEach-Object { [string]$_.Sender } | Where-Object { $_ } |
                      ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $recipients = @(Resolve-McRecipients -Addresses (@($addresses) + @($trackSenders) | Sort-Object -Unique))
    $internal = @($recipients | Where-Object Kind -eq 'Internal' | ForEach-Object Mailbox | Sort-Object -Unique)
    $external = @($recipients | Where-Object { $_.Kind -eq 'External' -and $_.Address -in $addresses })
    if ($external.Count) { Write-McLog "Внешних получателей: $($external.Count) — у них письмо удалить нельзя." 'Warn' }

    # ---- 3. KQL ----
    $kqlSenders = @($senders)
    if (($Criteria.IncludeForwards -or -not $kqlSenders.Count) -and $trackSenders.Count) { $kqlSenders += $trackSenders }
    $kqlSubject = $Criteria.Subject
    if (-not $kqlSubject -and $Criteria.MessageId -and $summary.Count) { $kqlSubject = $summary[0].Subject }
    $query = New-McSearchQuery -Senders $kqlSenders -Subject $kqlSubject -Attachment $Criteria.Attachment `
                 -Keywords $Criteria.Keywords -Start $Criteria.Start -End $Criteria.End

    $locations = @($internal)
    if ($Criteria.AllMailboxes -or -not $locations.Count) {
        if (-not $Criteria.AllMailboxes) { Write-McLog 'Ящики получателей не определены — поиск по ВСЕМ ящикам (дольше).' 'Warn' }
        $locations = @('All')
    }

    # ---- 4. Compliance Search ----
    Test-McCancel
    $name = New-McSearchName
    Write-McLog "Поиск $name, ящиков: $(if ($locations[0] -eq 'All') { 'все' } else { $locations.Count })."
    Write-McLog "KQL: $query"
    New-ComplianceSearch -Name $name -ExchangeLocation $locations -ContentMatchQuery $query | Out-Null
    Start-ComplianceSearch -Identity $name
    $s = Wait-McComplianceSearch -Name $name -TimeoutMinutes $TimeoutMinutes

    $hits    = @(Get-McSearchHits -SuccessResults ([string]$s.SuccessResults))
    $found   = @($hits | Where-Object Items -gt 0 | Sort-Object Mailbox)
    $missing = @(Get-McMissingMailboxes -Locations $locations -Hits $hits)
    Write-McLog ("Найдено писем: {0} в {1} ящиках, {2:N1} КБ." -f $s.Items, $found.Count, ($s.Size / 1KB)) 'Ok'
    if ($s.Errors) { Write-McLog "Ошибки поиска: $($s.Errors)" 'Error' }
    if ([int]$s.UnindexedItems -gt 0) {
        Write-McLog "Неиндексированных элементов в этих ящиках: $($s.UnindexedItems) — они не входят в результат и не удаляются." 'Warn'
    }
    Set-McProgress 100 'Готово'

    [pscustomobject]@{
        SearchName = $name
        Query      = $query
        Locations  = $locations
        Events     = $events
        Messages   = $summary
        Recipients = $recipients
        External   = $external
        Found      = $found
        Missing    = $missing
        Items      = [int]$s.Items
        SizeKB     = [math]::Round($s.Size / 1KB, 1)
        Unindexed  = [int]$s.UnindexedItems
        Errors     = [string]$s.Errors
    }
}

function Invoke-McSoftDelete {
    <# Purge (SoftDelete) of an existing Compliance Search: items go to Recoverable Items. #>
    param([Parameter(Mandatory)] [string] $SearchName, [int] $TimeoutMinutes = 60)
    Write-McLog "Мягкое удаление по поиску $SearchName..."
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete -Confirm:$false | Out-Null
    $actionName = "${SearchName}_Purge"
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    do {
        Start-Sleep -Seconds 5
        $a = Get-ComplianceSearchAction -Identity $actionName
        Set-McProgress 50 "Удаление: $($a.Status)"
        if ((Get-Date) -gt $deadline) { throw "Удаление не завершилось за $TimeoutMinutes мин." }
    } while ($a.Status -notin 'Completed', 'PartiallySucceeded', 'Failed')
    if ($a.Errors) { Write-McLog "Ошибки удаления: $($a.Errors)" 'Error' }
    Write-McLog "Статус удаления: $($a.Status). Письма перенесены в «Элементы для восстановления»." $(if ($a.Status -eq 'Completed') { 'Ok' } else { 'Warn' })
    [pscustomobject]@{ Status = [string]$a.Status; Errors = [string]$a.Errors }
}

function Invoke-McHardDelete {
    <# Permanent delete with Search-Mailbox -DeleteContent (incl. Recoverable Items), mailbox by mailbox. #>
    param([Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query)
    $i = 0
    foreach ($m in $Mailboxes) {
        Test-McCancel
        $i++; Set-McProgress ([int](80 * $i / $Mailboxes.Count)) "Безвозвратное удаление: $m"
        try {
            $r = Search-Mailbox -Identity $m -SearchQuery $Query -SearchDumpster -DeleteContent -Force `
                     -Confirm:$false -WarningAction SilentlyContinue -ErrorAction Stop
            Write-McLog ("{0}: удалено {1}" -f $m, $r.ResultItemsCount) 'Ok'
            [pscustomobject]@{ Mailbox = $m; Deleted = [int]$r.ResultItemsCount; Error = '' }
        } catch {
            Write-McLog "${m}: ошибка — $($_.Exception.Message)" 'Error'
            [pscustomobject]@{ Mailbox = $m; Deleted = 0; Error = $_.Exception.Message }
        }
    }
}

function Get-McCopyStatus {
    <#
        Verification via Search-Mailbox (Compliance Search also sees Recoverable Items and would
        report the same numbers after a SoftDelete). Two counts per mailbox:
        Visible = folders the user sees; Total = including Recoverable Items.
    #>
    param([Parameter(Mandatory)] [string[]] $Mailboxes, [Parameter(Mandatory)] [string] $Query)
    $i = 0
    foreach ($m in $Mailboxes) {
        Test-McCancel
        $i++; Set-McProgress ([int](100 * $i / $Mailboxes.Count)) "Проверка: $m"
        try {
            $v = Search-Mailbox -Identity $m -SearchQuery $Query -SearchDumpster:$false -EstimateResultOnly `
                     -WarningAction SilentlyContinue -ErrorAction Stop
            $t = Search-Mailbox -Identity $m -SearchQuery $Query -SearchDumpster -EstimateResultOnly `
                     -WarningAction SilentlyContinue -ErrorAction Stop
            [pscustomobject]@{ Mailbox = $m; Visible = [int]$v.ResultItemsCount; Total = [int]$t.ResultItemsCount; Note = '' }
        } catch {
            [pscustomobject]@{ Mailbox = $m; Visible = -1; Total = -1; Note = $_.Exception.Message }
        }
    }
}

function Get-McOldSearches {
    param([int] $OlderThanDays = 7)
    $limit = (Get-Date).AddDays(-$OlderThanDays)
    @(Get-ComplianceSearch | Where-Object { $_.Name -like "$(Get-McSearchPrefix)*" } | ForEach-Object {
        $when = if ($_.JobEndTime) { [datetime]$_.JobEndTime } elseif ($_.LastModifiedTime) { [datetime]$_.LastModifiedTime } else { $null }
        if ($when -and $when -lt $limit) { [pscustomobject]@{ Name = $_.Name; When = $when } }
    } | Sort-Object When)
}

function Remove-McSearches {
    param([Parameter(Mandatory)] [string[]] $Names)
    foreach ($n in $Names) {
        if ($n -notlike "$(Get-McSearchPrefix)*") { Write-McLog "Пропуск $n — не создан MailCleaner." 'Warn'; continue }
        Test-McCancel
        Get-ComplianceSearchAction -ErrorAction SilentlyContinue | Where-Object { $_.SearchName -eq $n } |
            ForEach-Object { Remove-ComplianceSearchAction -Identity $_.Identity -Confirm:$false -ErrorAction SilentlyContinue }
        Remove-ComplianceSearch -Identity $n -Confirm:$false
        Write-McLog "Удалён поиск $n." 'Ok'
    }
}

function Invoke-McOperation {
    <# Single entry point used by the GUI's background runspace. #>
    param([Parameter(Mandatory)] [string] $Name, [hashtable] $Params = @{})
    switch ($Name) {
        'Connect'      { Connect-McExchange -Server $Params.Server }
        'Search'       { Invoke-McSearch -Criteria $Params.Criteria -TimeoutMinutes $Params.TimeoutMinutes }
        'SoftDelete'   { $r = Invoke-McSoftDelete -SearchName $Params.SearchName -TimeoutMinutes $Params.TimeoutMinutes
                         [pscustomobject]@{ Delete = $r; Check = @(Get-McCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query) } }
        'HardDelete'   { $r = @(Invoke-McHardDelete -Mailboxes $Params.Mailboxes -Query $Params.Query)
                         [pscustomobject]@{ Delete = $r; Check = @(Get-McCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query) } }
        'Verify'       { [pscustomobject]@{ Check = @(Get-McCopyStatus -Mailboxes $Params.Mailboxes -Query $Params.Query) } }
        'ListOld'      { Get-McOldSearches -OlderThanDays $Params.Days }
        'RemoveOld'    { Remove-McSearches -Names $Params.Names }
        default        { throw "Неизвестная операция: $Name" }
    }
}

#endregion
