#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for src\MailCleaner.Core.ps1. Exchange is not necessary:
    Exchange cmdlets are replaced with stub functions and Pester mocks.

    Run:  Invoke-Pester -Path .\tests
#>

BeforeAll {
    . (Join-Path $PSScriptRoot '..\src\MailCleaner.Core.ps1')
    # Some tests examine Russian texts. The language tests set other languages themselves.
    $global:McLang = 'ru'

    # Stubs: Pester can mock only commands that exist.
    function Get-TransportService { [CmdletBinding()] param() }
    function Get-MessageTrackingLog {
        [CmdletBinding()]
        param($Server, $Sender, $MessageSubject, $MessageId, $Recipients, $Start, $End, $ResultSize)
    }
    function Get-Recipient { [CmdletBinding()] param($Identity, $Filter, $Anr, $ResultSize) }
    function Get-ComplianceSearch { [CmdletBinding()] param($Identity) }
    function Remove-ComplianceSearch { [CmdletBinding(SupportsShouldProcess)] param($Identity) }
    function Get-ComplianceSearchAction { [CmdletBinding()] param($Identity) }
    function Remove-ComplianceSearchAction { [CmdletBinding(SupportsShouldProcess)] param($Identity) }
    function Get-MailboxSearch { [CmdletBinding()] param($Identity) }
    function Remove-MailboxSearch { [CmdletBinding(SupportsShouldProcess)] param($Identity) }

    # Directory for the Get-Recipient mock: address -> recipient object.
    function New-Rcpt($Type, $Primary, [string[]]$Proxies) {
        [pscustomobject]@{ RecipientType = $Type; PrimarySmtpAddress = $Primary; EmailAddresses = $Proxies }
    }

    function New-Ev($Id, $From, $Subject, $Kind, [string[]]$To, [datetime]$Time) {
        [pscustomobject]@{
            MessageId = $Id; Sender = $From; MessageSubject = $Subject; EventId = $Kind
            Recipients = $To; Timestamp = $Time; ServerHostname = 'EXCH1'; Source = 'SMTP'
        }
    }
}

Describe 'Split-McAddressList' {
    It 'splits on comma, semicolon and spaces, lower-cases and removes duplicates' {
        Split-McAddressList 'B@x.com; a@x.com, A@X.com  c@x.com' | Should -Be @('a@x.com', 'b@x.com', 'c@x.com')
    }
    It 'returns an empty array for empty text' {
        @(Split-McAddressList '  ').Count | Should -Be 0
    }
}

Describe 'Get-McKqlDateRange' {
    It 'does not include the next day when End is at midnight' {
        $r = Get-McKqlDateRange -Start ([datetime]'2026-11-09') -End ([datetime]'2026-11-10') -UtcOffset 0
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-10'
    }
    It 'includes the End day when End is not at midnight' {
        $r = Get-McKqlDateRange -Start ([datetime]'2026-11-09 15:00') -End ([datetime]'2026-11-09 16:00') -UtcOffset 0
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-10'
    }
    It 'refuses End before Start' {
        { Get-McKqlDateRange -Start ([datetime]'2026-11-10') -End ([datetime]'2026-11-09') } | Should -Throw
    }
    It 'covers the local and the UTC day range (UTC+3, full local day)' {
        # Local 09.11 00:00 – 10.11 00:00 = UTC 08.11 21:00 – 09.11 21:00.
        $r = Get-McKqlDateRange -Start ([datetime]'2026-11-09') -End ([datetime]'2026-11-10') -UtcOffset ([timespan]'03:00')
        $r.From | Should -Be '2026-11-08'
        $r.To   | Should -Be '2026-11-10'
    }
    It 'covers the local and the UTC day range (UTC+3, end after local midnight)' {
        # Local 09.11 22:00 – 10.11 01:00 = UTC 09.11 19:00 – 09.11 22:00. The local range needs 10.11.
        $r = Get-McKqlDateRange -Start ([datetime]'2026-11-09 22:00') -End ([datetime]'2026-11-10 01:00') -UtcOffset ([timespan]'03:00')
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-11'
    }
}

Describe 'Get-McBaseSubject' {
    It 'removes reply and forward prefixes' {
        Get-McBaseSubject 'FW: RE: Отчёт' | Should -Be 'отчёт'
        Get-McBaseSubject 'Пересл: Отв: Отчёт' | Should -Be 'отчёт'
        Get-McBaseSubject 'Fwd:Report' | Should -Be 'report'
    }
    It 'keeps a tag in square brackets' {
        Get-McBaseSubject 'FW: [EXTERNAL] Report' | Should -Be '[external] report'
    }
}

Describe 'Get-McLargeDeleteWarning' {
    It 'gives no warning for a small delete' { Get-McLargeDeleteWarning -Items 50 -Messages 4 | Should -Be '' }
    It 'warns for many items'    { Get-McLargeDeleteWarning -Items 101 -Messages 1 | Should -Match 'БОЛЬШОЙ ОБЪЁМ' }
    It 'warns for many messages' { Get-McLargeDeleteWarning -Items 10 -Messages 6 | Should -Match 'разных писем' }
}

Describe 'Test-McHardDeleteScope' {
    It 'passes when Search-Mailbox finds the same or fewer items' {
        $st = @([pscustomobject]@{ Mailbox = 'a@x.com'; Total = 2; Note = '' }, [pscustomobject]@{ Mailbox = 'b@x.com'; Total = 0; Note = '' })
        @(Test-McHardDeleteScope -Status $st -Expected @{ 'a@x.com' = 2; 'b@x.com' = 1 }).Count | Should -Be 0
    }
    It 'fails when Search-Mailbox finds more items' {
        $st = @([pscustomobject]@{ Mailbox = 'a@x.com'; Total = 5; Note = '' })
        $bad = @(Test-McHardDeleteScope -Status $st -Expected @{ 'a@x.com' = 2 })
        $bad.Count | Should -Be 1
        $bad[0].Actual | Should -Be 5
    }
    It 'fails when the check did not run' {
        $st = @([pscustomobject]@{ Mailbox = 'a@x.com'; Total = -1; Note = 'access denied' })
        @(Test-McHardDeleteScope -Status $st -Expected @{ 'a@x.com' = 1 }).Count | Should -Be 1
    }
}

Describe 'New-McSearchQuery' {
    BeforeAll {
        $s = [datetime]'2026-11-09'
        $e = [datetime]'2026-11-10'
    }
    It 'joins senders with OR and adds subject and dates' {
        New-McSearchQuery -Senders 'b@x.com', 'A@x.com' -Subject 'Отчёт за квартал' -Start $s -End $e -UtcOffset 0 |
            Should -Be '(from:a@x.com OR from:b@x.com) AND (subject:"Отчёт за квартал") AND (sent>=2026-11-09) AND (sent<2026-11-10)'
    }
    It 'removes quotes from phrases' {
        New-McSearchQuery -Subject 'a "b" c' -Start $s -End $e -UtcOffset 0 | Should -Match 'subject:"a b c"'
    }
    It 'adds attachment and keywords' {
        $q = New-McSearchQuery -Attachment 'report.xlsx' -Keywords 'salary' -Start $s -End $e -UtcOffset 0
        $q | Should -Match '\(attachment:"report\.xlsx"\)'
        $q | Should -Match '\("salary"\)'
    }
    It 'refuses a query that has only dates' {
        { New-McSearchQuery -Start $s -End $e -UtcOffset 0 } | Should -Throw
    }
}

Describe 'Test-McCriteria' {
    It 'accepts a subject only' {
        @(Test-McCriteria @{ Subject = 'x'; Start = [datetime]'2026-11-09'; End = [datetime]'2026-11-10' }).Count | Should -Be 0
    }
    It 'refuses criteria without a key field' {
        @(Test-McCriteria @{ Recipient = 'a@x.com'; Start = [datetime]'2026-11-09'; End = [datetime]'2026-11-10' }).Count | Should -Be 1
    }
    It 'refuses a period longer than 31 days' {
        @(Test-McCriteria @{ Subject = 'x'; Start = [datetime]'2026-09-01'; End = [datetime]'2026-11-10' }).Count | Should -Be 1
    }
}

Describe 'Complete-McCriteria' {
    It 'adds missing keys and keeps given values' {
        $c = Complete-McCriteria @{ Subject = 'x'; Senders = @('a@x.com', '') }
        $c.Subject         | Should -Be 'x'
        $c.Senders         | Should -Be @('a@x.com')
        $c.MessageId       | Should -Be ''
        $c.IncludeForwards | Should -BeTrue
    }
}

Describe 'Get-McSearchHits' {
    It 'parses the comma format' {
        $h = @(Get-McSearchHits '{Location: A@x.com, Item count: 2, Total size: 2048, Location: b@x.com, Item count: 0, Total size: 0}')
        $h.Count | Should -Be 2
        $h[0].Mailbox | Should -Be 'a@x.com'
        $h[0].Items   | Should -Be 2
        $h[0].SizeKB  | Should -Be 2
    }
    It 'parses the semicolon format' {
        $h = @(Get-McSearchHits "Location: a@x.com; Item count: 1; Total size: 1024`r`nLocation: c@x.com; Item count: 3; Total size: 3072")
        $h.Count | Should -Be 2
        $h[1].Items | Should -Be 3
    }
    It 'returns nothing for empty text' {
        @(Get-McSearchHits '').Count | Should -Be 0
    }
}

Describe 'Test-McServerGeneratedId' {
    It 'marks a .local host' {
        Test-McServerGeneratedId '<5f3c@EXCH01.Contoso.local>' | Should -BeTrue
    }
    It 'marks a transport server name' {
        Test-McServerGeneratedId '<abc@exch1.contoso.com>' -ServerNames 'EXCH1' | Should -BeTrue
    }
    It 'does not mark a normal client MessageId' {
        Test-McServerGeneratedId '<abc@contoso.com>' -ServerNames 'EXCH1' | Should -BeFalse
    }
}

Describe 'Get-McMessageSummary' {
    It 'makes one row per MessageId with the first RECEIVE and all recipients' {
        $t = [datetime]'2026-11-09 16:53'
        $ev = @(
            New-Ev '<1@contoso.com>' 'A@x.com' 'Report' 'RECEIVE' @('b@x.com', 'c@x.com') $t
            New-Ev '<1@contoso.com>' 'a@x.com' 'Report' 'DELIVER' @('b@x.com') $t.AddSeconds(2)
            New-Ev '<2@EXCH1.contoso.local>' 'b@x.com' 'FW: Report' 'RECEIVE' @('d@y.com') $t.AddSeconds(20)
        )
        $s = @(Get-McMessageSummary -Events $ev)
        $s.Count | Should -Be 2
        $s[0].Sender     | Should -Be 'a@x.com'
        $s[0].Recipients | Should -Be 2
        $s[0].Note       | Should -BeNullOrEmpty
        $s[1].Note       | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-McMissingMailboxes' {
    It 'returns searched mailboxes without items' {
        $hits = @([pscustomobject]@{ Mailbox = 'a@x.com'; Items = 1 }, [pscustomobject]@{ Mailbox = 'b@x.com'; Items = 0 })
        Get-McMissingMailboxes -Locations 'A@x.com', 'b@x.com', 'c@x.com' -Hits $hits | Should -Be @('b@x.com', 'c@x.com')
    }
    It 'returns nothing for an All search' {
        @(Get-McMissingMailboxes -Locations 'All' -Hits @()).Count | Should -Be 0
    }
}

Describe 'New-McSearchName' {
    It 'uses the MC_ prefix and a timestamp' {
        New-McSearchName -Now ([datetime]'2026-11-10 11:03:45') | Should -Be 'MC_20261110_110345'
    }
}

Describe 'Write-McAudit' {
    It 'appends rows with fixed columns' {
        $p = Join-Path $TestDrive 'audit\audit.csv'
        Write-McAudit -Path $p -Record @{ Action = 'Search'; Items = 5; Unknown = 'ignored' }
        Write-McAudit -Path $p -Record @{ Action = 'SoftDelete'; Result = 'Clean' }
        $rows = @(Import-Csv $p)
        $rows.Count | Should -Be 2
        $rows[0].Action | Should -Be 'Search'
        $rows[0].Items  | Should -Be '5'
        $rows[1].Result | Should -Be 'Clean'
        $rows[0].PSObject.Properties.Name | Should -Not -Contain 'Unknown'
    }
}

Describe 'Resolve-McRecipients' {
    BeforeAll {
        $script:Directory = @(
            New-Rcpt 'UserMailbox' 'U@x.com' @('SMTP:U@x.com', 'smtp:alias@x.com')
            New-Rcpt 'MailContact' 'c@y.com' @('SMTP:c@y.com')
            New-Rcpt 'MailUniversalDistributionGroup' 'g@x.com' @('SMTP:g@x.com')
        )
        Mock Get-Recipient {
            if (-not $Filter) { throw 'not found' }
            $wanted = @([regex]::Matches($Filter, "smtp:([^']+)") | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() })
            $script:Directory | Where-Object {
                $p = @($_.EmailAddresses | ForEach-Object { ($_ -replace '^smtp:', '').ToLowerInvariant() })
                @($wanted | Where-Object { $_ -in $p }).Count
            }
        }
    }
    It 'classifies mailbox, alias, contact, group and unknown addresses' {
        $r = @(Resolve-McRecipients -Addresses 'u@x.com', 'alias@x.com', 'c@y.com', 'g@x.com', 'z@ext.org')
        ($r | Where-Object Address -eq 'u@x.com').Kind        | Should -Be 'Internal'
        ($r | Where-Object Address -eq 'u@x.com').Mailbox     | Should -Be 'u@x.com'
        ($r | Where-Object Address -eq 'alias@x.com').Mailbox | Should -Be 'u@x.com'
        ($r | Where-Object Address -eq 'c@y.com').Kind        | Should -Be 'External'
        ($r | Where-Object Address -eq 'g@x.com').Kind        | Should -Be 'Group'
        ($r | Where-Object Address -eq 'z@ext.org').Kind      | Should -Be 'External'
    }
    It 'uses one call per batch' {
        $null = Resolve-McRecipients -Addresses 'u@x.com', 'c@y.com', 'g@x.com' -BatchSize 2
        Should -Invoke Get-Recipient -Times 2 -Exactly
    }
    It 'does not look up addresses outside the accepted domains' {
        $r = @(Resolve-McRecipients -Addresses 'u@x.com', 'z@ext.org' -AcceptedDomains 'x.com')
        Should -Invoke Get-Recipient -Times 1 -Exactly
        ($r | Where-Object Address -eq 'z@ext.org').Kind | Should -Be 'External'
    }
}

Describe 'Test-McInternalDomain' {
    It 'matches an exact domain' { Test-McInternalDomain 'a@Contoso.com' 'contoso.com' | Should -BeTrue }
    It 'matches a subdomain of a wildcard domain' { Test-McInternalDomain 'a@mail.contoso.com' '*.contoso.com' | Should -BeTrue }
    It 'does not match another domain' { Test-McInternalDomain 'a@outlook.com' 'contoso.com', '*.contoso.com' | Should -BeFalse }
}

Describe 'ConvertFrom-McSizeText' {
    It 'reads the bytes in parentheses' { ConvertFrom-McSizeText '1.25 MB (1,310,720 bytes)' | Should -Be 1280 }
    It 'reads bytes with spaces' { ConvertFrom-McSizeText '2 KB (2 048 bytes)' | Should -Be 2 }
    It 'returns 0 for unknown text' { ConvertFrom-McSizeText '' | Should -Be 0 }
}

Describe 'Select-McUniqueEvents' {
    It 'removes duplicate events and keeps only the given MessageIds' {
        $t = [datetime]'2026-11-09 10:00'
        $a = New-Ev '<1@x>' 'a@x.com' 'S' 'DELIVER' @('b@x.com') $t
        $b = New-Ev '<2@x>' 'a@x.com' 'S' 'DELIVER' @('b@x.com') $t
        @(Select-McUniqueEvents -Events @($a, $a, $b) -MessageIds '<1@x>').Count | Should -Be 1
    }
}

Describe 'Complete-McCriteria and Test-McCriteria: only the recipient mailbox' {
    It 'sets OnlyRecipient to false by default' {
        (Complete-McCriteria @{}).OnlyRecipient | Should -BeFalse
    }
    It 'refuses OnlyRecipient without a recipient' {
        @(Test-McCriteria @{ Subject = 'x'; OnlyRecipient = $true; Start = [datetime]'2026-11-09'; End = [datetime]'2026-11-10' }).Count |
            Should -Be 1
    }
    It 'refuses OnlyRecipient together with AllMailboxes' {
        @(Test-McCriteria @{ Subject = 'x'; Recipient = 'a@x.com'; OnlyRecipient = $true; AllMailboxes = $true
                             Start = [datetime]'2026-11-09'; End = [datetime]'2026-11-10' }).Count | Should -Be 1
    }
}

Describe 'Find-McTrackingEvents' {
    BeforeAll {
        $t = [datetime]'2026-11-09 16:53'
        Mock Get-TransportService { [pscustomobject]@{ Name = 'EXCH1' } }
        Mock Get-MessageTrackingLog {
            $all = @(
                New-Ev '<1@x.com>' 'a@x.com' 'Отчёт за квартал'     'RECEIVE' @('b@x.com') $t
                New-Ev '<2@x.com>' 'b@x.com' 'FW: Отчёт за квартал' 'RECEIVE' @('c@x.com') $t.AddMinutes(10)
                New-Ev '<3@x.com>' 'a@x.com' 'Другое письмо'    'RECEIVE' @('d@x.com') $t.AddMinutes(20)
            )
            if ($MessageId)      { return $all | Where-Object MessageId -eq $MessageId }
            if ($Sender)         { return $all | Where-Object Sender -eq $Sender }
            if ($MessageSubject) { return $all | Where-Object { $_.MessageSubject -like "*$MessageSubject*" } }
            $all
        }
    }
    It 'finds the original and the forward when forwards are included' {
        $r = Find-McTrackingEvents -Senders 'a@x.com' -Subject 'Отчёт за квартал' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards
        @($r.Events | Select-Object -ExpandProperty MessageId -Unique) | Should -Be @('<1@x.com>', '<2@x.com>')
    }
    It 'finds only the original when forwards are excluded' {
        $r = Find-McTrackingEvents -Senders 'a@x.com' -Subject 'Отчёт за квартал' -Start $t.Date -End $t.Date.AddDays(1)
        @($r.Events | Select-Object -ExpandProperty MessageId -Unique) | Should -Be @('<1@x.com>')
    }
    It 'finds the message by subject when the sender is wrong' {
        $r = Find-McTrackingEvents -Senders 'wrong@x.com' -Subject 'Отчёт за квартал' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards
        @($r.Events | Select-Object -ExpandProperty MessageId -Unique).Count | Should -Be 2
    }
    It 'does not do a second pass by MessageId without a recipient filter' {
        $null = Find-McTrackingEvents -Senders 'a@x.com' -Subject 'Отчёт за квартал' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards -CompleteRecipients
        # One sender query + one subject query on one server.
        Should -Invoke Get-MessageTrackingLog -Times 2 -Exactly
    }
    It 'does a second pass by MessageId with a recipient filter' {
        $null = Find-McTrackingEvents -Subject 'Отчёт за квартал' -Recipient 'b@x.com' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards -CompleteRecipients
        # One subject query + one query for each of the two messages.
        Should -Invoke Get-MessageTrackingLog -Times 3 -Exactly
    }
    It 'does not take a message from another sender with a different subject as a forward' {
        Mock Get-MessageTrackingLog {
            @(
                New-Ev '<1@x.com>' 'a@x.com' 'Итоги'             'RECEIVE' @('b@x.com') $t
                New-Ev '<2@x.com>' 'b@x.com' 'FW: Итоги'         'RECEIVE' @('c@x.com') $t
                New-Ev '<9@x.com>' 'z@x.com' 'Итоги года для HR' 'RECEIVE' @('d@x.com') $t
            ) | Where-Object { -not $Sender -or $_.Sender -eq $Sender }
        }
        $r = Find-McTrackingEvents -Senders 'a@x.com' -Subject 'Итоги' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards
        @($r.Events | Select-Object -ExpandProperty MessageId -Unique | Sort-Object) | Should -Be @('<1@x.com>', '<2@x.com>')
    }
    It 'matches a subject with square brackets as text' {
        Mock Get-MessageTrackingLog {
            @(New-Ev '<1@x.com>' 'a@x.com' '[EXTERNAL] Отчёт' 'RECEIVE' @('b@x.com') $t)
        }
        $r = Find-McTrackingEvents -Senders 'a@x.com' -Subject '[EXTERNAL] Отчёт' -Start $t.Date -End $t.Date.AddDays(1)
        @($r.Events).Count | Should -Be 1
        { Find-McTrackingEvents -Subject '[EXTERNAL Отчёт' -Start $t.Date -End $t.Date.AddDays(1) } | Should -Not -Throw
    }
    It 'returns the errors of a server that fails' {
        Mock Get-TransportService { [pscustomobject]@{ Name = 'EXCH1' }; [pscustomobject]@{ Name = 'EXCH2' } }
        Mock Get-MessageTrackingLog {
            if ($Server -eq 'EXCH2') { Write-Error 'server is not available' ; return }
            @(New-Ev '<1@x.com>' 'a@x.com' 'Отчёт' 'RECEIVE' @('b@x.com') $t)
        }
        $r = Find-McTrackingEvents -Senders 'a@x.com' -Start $t.Date -End $t.Date.AddDays(1)
        @($r.Errors).Count | Should -Be 1
        $r.Errors[0] | Should -Match 'EXCH2'
        @($r.Events).Count | Should -Be 1
    }
    It 'does not do a second pass without -CompleteRecipients' {
        $null = Find-McTrackingEvents -Subject 'Отчёт за квартал' -Recipient 'b@x.com' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards
        Should -Invoke Get-MessageTrackingLog -Times 1 -Exactly
    }
}

Describe 'Test-McCancel' {
    AfterEach { Remove-Variable McSync -Scope Global -ErrorAction SilentlyContinue }
    It 'throws when cancel is requested' {
        $global:McSync = @{ Cancel = $true; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        { Test-McCancel } | Should -Throw
    }
    It 'does nothing without cancel' {
        $global:McSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        { Test-McCancel } | Should -Not -Throw
    }
}

Describe 'Language files (Lang\*.lng)' {
    BeforeAll {
        $root = Join-Path $PSScriptRoot '..'
        $langDir = Join-Path $root 'Lang'
        $en = Read-McLanguageFile (Join-Path $langDir 'en.lng')
        $ru = Read-McLanguageFile (Join-Path $langDir 'ru.lng')
    }
    AfterEach { $global:McLang = 'ru' }

    It 'has the bundled files ru.lng and en.lng with names' {
        $en.Code | Should -Be 'en'
        $ru.Code | Should -Be 'ru'
        $en.Name | Should -Be 'English'
        $ru.Name | Should -Not -BeNullOrEmpty
    }
    It 'ru.lng and en.lng are complete and correct' {
        @(Test-McLanguageFile (Join-Path $langDir 'ru.lng')) | Should -BeNullOrEmpty
        @(Test-McLanguageFile (Join-Path $langDir 'en.lng')) | Should -BeNullOrEmpty
    }
    It 'all other .lng files have no unknown keys and correct placeholders' {
        foreach ($f in Get-ChildItem $langDir -Filter '*.lng' | Where-Object Name -notin 'ru.lng', 'en.lng') {
            @(Test-McLanguageFile $f.FullName -AllowMissing) | Should -BeNullOrEmpty -Because $f.Name
        }
    }
    It 'every text key used in the code is in en.lng' {
        $files = @(Join-Path $root 'MailCleaner.ps1'; Join-Path $root 'src\MailCleaner.Core.ps1')
        $used = foreach ($f in $files) {
            $src = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
            [regex]::Matches($src, "\bT\s+'([A-Za-z0-9_.]+)'") | ForEach-Object { $_.Groups[1].Value }
            # Keys given to helpers that call T: New-McButton 'act.soft', Set-McLoc $x 'crit.group', ...
            [regex]::Matches($src, "(?:New-McLabel|New-McButton|New-McCheckBox|Set-McLoc\s+\S+|Set-McTip\s+\S+)\s+'([A-Za-z0-9_.]+)'") |
                ForEach-Object { $_.Groups[1].Value }
        }
        # Keys made at run time: "tab.$k" and "kind.$($_.Kind)".
        $used = @($used) + @('messages', 'tracking', 'recipients', 'found', 'missing', 'check' | ForEach-Object { "tab.$_" }) +
                @('Internal', 'External', 'Group', 'Other' | ForEach-Object { "kind.$_" })
        $missing = @($used | Sort-Object -Unique | Where-Object { $_ -and -not $en.Texts.ContainsKey($_) })
        $missing | Should -BeNullOrEmpty
    }
    It 'gives the text in the current language and fills the placeholders' {
        $global:McLang = 'en'
        T 'log.found' 3 2 1.5 | Should -Match '^Messages found: 3 in 2 mailboxes'
        $global:McLang = 'ru'
        T 'log.found' 3 2 1.5 | Should -Match '^Найдено писем: 3 в 2 ящиках'
    }
    It 'uses English for a key that a language does not have, and the key for an unknown key' {
        $global:McLang = 'xx'
        T 'conn.connect' | Should -Be 'Connect'
        T 'no.such.key' | Should -Be 'no.such.key'
    }
}

Describe 'Read-McLanguageFile and Test-McLanguageFile' {
    BeforeAll {
        $ref = Join-Path $TestDrive 'ref.lng'
        Set-Content -Path $ref -Encoding UTF8 -Value @(
            '@code = en', '@name = English', 'a.one = One {0}', 'a.two = Two')
    }
    It 'reads meta lines, comments, and an equals sign in the text' {
        $f = Join-Path $TestDrive 'de.lng'
        Set-Content -Path $f -Encoding UTF8 -Value @(
            '# comment', '; comment', '', '@code = de', '@name = Deutsch', '@author = Test',
            'a.one = Eins {0} = gleich', '  a.two   =   Zwei  ')
        $l = Read-McLanguageFile $f
        $l.Code | Should -Be 'de'
        $l.Name | Should -Be 'Deutsch'
        $l.Author | Should -Be 'Test'
        $l.Texts['a.one'] | Should -Be 'Eins {0} = gleich'
        $l.Texts['a.two'] | Should -Be 'Zwei'
        @(Test-McLanguageFile $f -ReferencePath $ref) | Should -BeNullOrEmpty
    }
    It 'uses the file name when there is no @code line' {
        $f = Join-Path $TestDrive 'pl.lng'
        Set-Content -Path $f -Encoding UTF8 -Value @('a.two = Dwa')
        (Read-McLanguageFile $f).Code | Should -Be 'pl'
    }
    It 'finds unknown keys, missing keys, wrong placeholders, duplicates and bad lines' {
        $f = Join-Path $TestDrive 'bad.lng'
        Set-Content -Path $f -Encoding UTF8 -Value @(
            '@code = bad', '@name = Bad', 'a.one = Eins', 'a.one = Eins {0}', 'a.three = Drei', 'no equals sign')
        $p = @(Test-McLanguageFile $f -ReferencePath $ref)
        ($p -join "`n") | Should -Match "duplicate key 'a.one'"
        ($p -join "`n") | Should -Match "unknown key 'a.three'"
        ($p -join "`n") | Should -Match "missing key 'a.two'"
        ($p -join "`n") | Should -Match 'no .key = text.'
    }
    It 'finds a wrong placeholder' {
        $f = Join-Path $TestDrive 'ph.lng'
        Set-Content -Path $f -Encoding UTF8 -Value @('@name = Ph', 'a.one = Eins {1}', 'a.two = Zwei')
        (@(Test-McLanguageFile $f -ReferencePath $ref) -join "`n") | Should -Match "key 'a.one': placeholders"
    }
}
Describe 'Events without MessageId' {
    It 'are not a row in the message summary' {
        $t = [datetime]'2026-11-10 08:59:45'
        $ev = @(
            New-Ev ''          'a@x.com' ''          'NOTIFYMAPI' @() $t
            New-Ev '<1@x.com>' 'a@x.com' 'RE: Topic' 'RECEIVE'    @('b@x.com') $t
        )
        $s = @(Get-McMessageSummary -Events $ev)
        $s.Count | Should -Be 1
        $s[0].MessageId | Should -Be '<1@x.com>'
    }
    It 'stay in the event list, but are not counted as messages' {
        $t = [datetime]'2026-11-10 08:59:45'
        Mock Get-TransportService { [pscustomobject]@{ Name = 'EXCH1' } }
        Mock Get-MessageTrackingLog {
            @(
                New-Ev ''          'a@x.com' ''          'NOTIFYMAPI' @() $t
                New-Ev '<1@x.com>' 'a@x.com' 'RE: Topic' 'RECEIVE'    @('b@x.com') $t
            )
        }
        $r = Find-McTrackingEvents -Senders 'a@x.com' -Start $t.Date -End $t.Date.AddDays(1)
        @($r.Events).Count | Should -Be 2
        @($r.Events | Where-Object { $_.MessageId } | Select-Object -ExpandProperty MessageId -Unique) | Should -Be @('<1@x.com>')
    }
}

Describe 'MailCleaner search names' {
    It 'accepts only the names that MailCleaner makes' {
        Test-McSearchName 'MC_20261110_124718'        | Should -BeTrue
        Test-McSearchName 'MC_20261110_124718-shadow' | Should -BeTrue
        Test-McSearchName 'MC_quick_20261110_124718'  | Should -BeFalse
        Test-McSearchName 'ContentSearch_for_Delete1' | Should -BeFalse
        Test-McSearchName 'mc_20261110_124718'        | Should -BeFalse
    }
    It 'reads the time from the name' {
        Get-McSearchNameTime 'MC_20261110_124718-shadow' | Should -Be ([datetime]'2026-11-10 12:47:18')
        Get-McSearchNameTime 'Other' | Should -BeNullOrEmpty
    }
}

Describe 'Get-McOldSearches and Remove-McSearches' {
    BeforeAll {
        $old = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        $new = (Get-Date).AddMinutes(-5).ToString('yyyyMMdd_HHmmss')
        Mock Get-ComplianceSearch { @([pscustomobject]@{ Name = "MC_$old" }, [pscustomobject]@{ Name = "MC_$new" },
                                      [pscustomobject]@{ Name = 'ContentSearch_for_Delete1' }) }
        Mock Get-MailboxSearch {
            $all = @([pscustomobject]@{ Name = "MC_$old-shadow" }, [pscustomobject]@{ Name = "MC_$new-shadow" },
                     [pscustomobject]@{ Name = 'Legal hold 2025' })
            if ($Identity) { $all | Where-Object Name -eq $Identity } else { $all }
        }
        Mock Get-ComplianceSearchAction { }
        Mock Remove-ComplianceSearch { }
        Mock Remove-MailboxSearch { }
    }
    It 'lists old searches and old shadow searches, and not the other searches' {
        $l = @(Get-McOldSearches -OlderThanDays 7)
        @($l | ForEach-Object Name | Sort-Object) | Should -Be @("MC_$old", "MC_$old-shadow")
        ($l | Where-Object Name -eq "MC_$old-shadow").Kind | Should -Be 'Shadow'
    }
    It 'lists today''s shadow search with 0 days' {
        @(Get-McOldSearches -OlderThanDays 0 | ForEach-Object Name) | Should -Contain "MC_$new-shadow"
    }
    It 'removes the search and its shadow once, and skips other names' {
        Remove-McSearches -Names "MC_$old", "MC_$old-shadow", 'Legal hold 2025'
        Should -Invoke Remove-ComplianceSearch -Times 1 -Exactly -ParameterFilter { $Identity -eq "MC_$old" }
        Should -Invoke Remove-MailboxSearch -Times 1 -Exactly -ParameterFilter { $Identity -eq "MC_$old-shadow" }
        Should -Invoke Remove-MailboxSearch -Times 0 -Exactly -ParameterFilter { $Identity -eq 'Legal hold 2025' }
    }
}
Describe 'Partial addresses' {
    It 'keeps a name with spaces as one item, and splits full addresses at spaces' {
        Split-McAddressList 'Ivanov Ivan; a@x.com b@x.com' | Should -Be @('a@x.com', 'b@x.com', 'ivanov ivan')
        Split-McAddressList "jsmi`r`nc@x.com" | Should -Be @('c@x.com', 'jsmi')
    }
    It 'knows a full address' {
        Test-McFullAddress 'a.b@contoso.com' | Should -BeTrue
        Test-McFullAddress 'jsmith'    | Should -BeFalse
        Test-McFullAddress 'jsmi@contoso'  | Should -BeFalse
        Test-McFullAddress 'Ivanov Ivan'     | Should -BeFalse
    }
    It 'finds the candidates in the address book (ANR)' {
        Mock Get-Recipient {
            if ($Anr -eq 'ivan') {
                [pscustomobject]@{ PrimarySmtpAddress = 'I.Ivanov@x.com'; DisplayName = 'Ivanov Ivan'; RecipientTypeDetails = 'UserMailbox' }
                [pscustomobject]@{ PrimarySmtpAddress = 'ivanova@x.com'; DisplayName = 'Ivanova Anna'; RecipientTypeDetails = 'UserMailbox' }
            }
        }
        $c = @(Find-McAddressCandidates -Text 'ivan')
        $c.Count | Should -Be 2
        $c[0].Address | Should -Be 'i.ivanov@x.com'
        $c[0].Name    | Should -Be 'Ivanov Ivan'
        $c[0].Type    | Should -Be 'UserMailbox'
        @(Find-McAddressCandidates -Text 'nobody').Count | Should -Be 0
    }
    It 'returns no candidates when the address book search fails' {
        Mock Get-Recipient { throw 'access denied' }
        @(Find-McAddressCandidates -Text 'ivan').Count | Should -Be 0
    }
    It 'gives one result for each input' {
        Mock Get-Recipient { if ($Anr -eq 'jsmi') { [pscustomobject]@{ PrimarySmtpAddress = 'jsmith@x.com'; DisplayName = 'V S'; RecipientType = 'UserMailbox' } } }
        $r = @(Resolve-McAddressInputs -Inputs 'jsmi', 'nobody')
        $r.Count | Should -Be 2
        @($r[0].Candidates).Count | Should -Be 1
        $r[0].Candidates[0].Address | Should -Be 'jsmith@x.com'
        @($r[1].Candidates).Count | Should -Be 0
    }
}
Describe 'Remove-McSearches when Exchange removes the shadow search itself' {
    BeforeAll {
        $stamp = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        $stamp2 = (Get-Date).AddDays(-9).ToString('yyyyMMdd_HHmmss')
        Mock Get-ComplianceSearchAction { }
        Mock Remove-MailboxSearch { }
        # Remove-ComplianceSearch also removed the shadow: Get-MailboxSearch finds nothing.
        Mock Get-MailboxSearch { }
    }
    It 'does not fail when the shadow search is already gone (the case from the log)' {
        Mock Remove-ComplianceSearch { }
        { Remove-McSearches -Names "MC_$stamp-shadow", "MC_$stamp" } | Should -Not -Throw
        Should -Invoke Remove-ComplianceSearch -Times 1 -Exactly
        Should -Invoke Remove-MailboxSearch -Times 0 -Exactly
    }
    It 'continues with the next search after an error' {
        Mock Remove-ComplianceSearch { if ($Identity -eq "MC_$stamp") { throw 'busy' } }
        { Remove-McSearches -Names "MC_$stamp", "MC_$stamp2" } | Should -Not -Throw
        Should -Invoke Remove-ComplianceSearch -Times 1 -Exactly -ParameterFilter { $Identity -eq "MC_$stamp2" }
    }
}
Describe 'Remove-McSearches when Remove-ComplianceSearch returns an error' {
    BeforeAll {
        $stamp = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        Mock Get-MailboxSearch { }
        Mock Remove-MailboxSearch { }
        Mock Remove-ComplianceSearch { throw [System.NullReferenceException]::new('Object reference not set to an instance of an object.') }
    }
    It 'reports a removal when the search is gone after the error (the case from the log)' {
        Mock Get-ComplianceSearch { }
        Mock Get-ComplianceSearchAction { throw [System.NullReferenceException]::new('Object reference not set to an instance of an object.') }
        $global:McSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        try {
            Remove-McSearches -Names "MC_$stamp"
            $lines = @($global:McSync.Log.ToArray())
        } finally { Remove-Variable McSync -Scope Global -ErrorAction SilentlyContinue }
        ($lines -join "`n") | Should -Match '\[WARN \].*MC_'
        ($lines -join "`n") | Should -Not -Match '\[ERROR\]'
    }
    It 'reports an error when the search is still there after the error' {
        Mock Get-ComplianceSearch { [pscustomobject]@{ Name = "MC_$stamp" } }
        Mock Get-ComplianceSearchAction { }
        $global:McSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        try {
            Remove-McSearches -Names "MC_$stamp"
            $lines = @($global:McSync.Log.ToArray())
        } finally { Remove-Variable McSync -Scope Global -ErrorAction SilentlyContinue }
        ($lines -join "`n") | Should -Match '\[ERROR\].*MC_'
    }
}