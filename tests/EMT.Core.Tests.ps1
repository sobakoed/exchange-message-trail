#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for src\EMT.Core.ps1. Exchange is not necessary:
    Exchange cmdlets are replaced with stub functions and Pester mocks.

    Run:  Invoke-Pester -Path .\tests
#>

BeforeAll {
    . (Join-Path $PSScriptRoot '..\src\EMT.Core.ps1')
    # Some tests examine Russian texts. The language tests set other languages themselves.
    $global:EmtLang = 'ru'

    # Stubs: Pester can mock only commands that exist.
    function Get-TransportService { [CmdletBinding()] param() }
    function Get-MessageTrackingLog {
        [CmdletBinding()]
        param($Server, $Sender, $MessageSubject, $MessageId, $Recipients, $Start, $End, $ResultSize)
    }
    function Get-Recipient { [CmdletBinding()] param($Identity, $Filter, $Anr, $ResultSize, $RecipientType) }
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

Describe 'Split-EmtAddressList' {
    It 'splits on comma, semicolon and spaces, lower-cases and removes duplicates' {
        Split-EmtAddressList 'B@x.com; a@x.com, A@X.com  c@x.com' | Should -Be @('a@x.com', 'b@x.com', 'c@x.com')
    }
    It 'returns an empty array for empty text' {
        @(Split-EmtAddressList '  ').Count | Should -Be 0
    }
}

Describe 'Get-EmtKqlDateRange' {
    It 'does not include the next day when End is at midnight' {
        $r = Get-EmtKqlDateRange -Start ([datetime]'2026-11-09') -End ([datetime]'2026-11-10') -UtcOffset 0
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-10'
    }
    It 'includes the End day when End is not at midnight' {
        $r = Get-EmtKqlDateRange -Start ([datetime]'2026-11-09 15:00') -End ([datetime]'2026-11-09 16:00') -UtcOffset 0
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-10'
    }
    It 'refuses End before Start' {
        { Get-EmtKqlDateRange -Start ([datetime]'2026-11-10') -End ([datetime]'2026-11-09') } | Should -Throw
    }
    It 'covers the local and the UTC day range (UTC+3, full local day)' {
        # Local 09.11 00:00 – 10.11 00:00 = UTC 08.11 21:00 – 09.11 21:00.
        $r = Get-EmtKqlDateRange -Start ([datetime]'2026-11-09') -End ([datetime]'2026-11-10') -UtcOffset ([timespan]'03:00')
        $r.From | Should -Be '2026-11-08'
        $r.To   | Should -Be '2026-11-10'
    }
    It 'covers the local and the UTC day range (UTC+3, end after local midnight)' {
        # Local 09.11 22:00 – 10.11 01:00 = UTC 09.11 19:00 – 09.11 22:00. The local range needs 10.11.
        $r = Get-EmtKqlDateRange -Start ([datetime]'2026-11-09 22:00') -End ([datetime]'2026-11-10 01:00') -UtcOffset ([timespan]'03:00')
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-11'
    }
}

Describe 'Get-EmtBaseSubject' {
    It 'removes reply and forward prefixes' {
        Get-EmtBaseSubject 'FW: RE: Отчёт' | Should -Be 'отчёт'
        Get-EmtBaseSubject 'Пересл: Отв: Отчёт' | Should -Be 'отчёт'
        Get-EmtBaseSubject 'Fwd:Report' | Should -Be 'report'
    }
    It 'keeps a tag in square brackets' {
        Get-EmtBaseSubject 'FW: [EXTERNAL] Report' | Should -Be '[external] report'
    }
}

Describe 'Get-EmtLargeDeleteWarning' {
    It 'gives no warning for a small delete' { Get-EmtLargeDeleteWarning -Items 50 -Messages 4 | Should -Be '' }
    It 'warns for many items'    { Get-EmtLargeDeleteWarning -Items 101 -Messages 1 | Should -Match 'БОЛЬШОЙ ОБЪЁМ' }
    It 'warns for many messages' { Get-EmtLargeDeleteWarning -Items 10 -Messages 6 | Should -Match 'разных писем' }
}

Describe 'Test-EmtHardDeleteScope' {
    It 'passes when Search-Mailbox finds the same or fewer items' {
        $st = @([pscustomobject]@{ Mailbox = 'a@x.com'; Total = 2; Note = '' }, [pscustomobject]@{ Mailbox = 'b@x.com'; Total = 0; Note = '' })
        @(Test-EmtHardDeleteScope -Status $st -Expected @{ 'a@x.com' = 2; 'b@x.com' = 1 }).Count | Should -Be 0
    }
    It 'fails when Search-Mailbox finds more items' {
        $st = @([pscustomobject]@{ Mailbox = 'a@x.com'; Total = 5; Note = '' })
        $bad = @(Test-EmtHardDeleteScope -Status $st -Expected @{ 'a@x.com' = 2 })
        $bad.Count | Should -Be 1
        $bad[0].Actual | Should -Be 5
    }
    It 'fails when the check did not run' {
        $st = @([pscustomobject]@{ Mailbox = 'a@x.com'; Total = -1; Note = 'access denied' })
        @(Test-EmtHardDeleteScope -Status $st -Expected @{ 'a@x.com' = 1 }).Count | Should -Be 1
    }
}

Describe 'New-EmtSearchQuery' {
    BeforeAll {
        $s = [datetime]'2026-11-09'
        $e = [datetime]'2026-11-10'
    }
    It 'joins senders with OR and adds subject and dates' {
        New-EmtSearchQuery -Senders 'b@x.com', 'A@x.com' -Subject 'Отчёт за квартал' -Start $s -End $e -UtcOffset 0 |
            Should -Be '(from:a@x.com OR from:b@x.com) AND (subject:"Отчёт за квартал") AND (sent>=2026-11-09) AND (sent<2026-11-10)'
    }
    It 'removes quotes from phrases' {
        New-EmtSearchQuery -Subject 'a "b" c' -Start $s -End $e -UtcOffset 0 | Should -Match 'subject:"a b c"'
    }
    It 'adds attachment and keywords' {
        $q = New-EmtSearchQuery -Attachment 'report.xlsx' -Keywords 'salary' -Start $s -End $e -UtcOffset 0
        $q | Should -Match '\(attachment:"report\.xlsx"\)'
        $q | Should -Match '\("salary"\)'
    }
    It 'refuses a query that has only dates' {
        { New-EmtSearchQuery -Start $s -End $e -UtcOffset 0 } | Should -Throw
    }
}

Describe 'Test-EmtCriteria' {
    It 'accepts a subject only' {
        @(Test-EmtCriteria @{ Subject = 'x'; Start = [datetime]'2026-11-09'; End = [datetime]'2026-11-10' }).Count | Should -Be 0
    }
    It 'refuses criteria without a key field' {
        @(Test-EmtCriteria @{ Recipient = 'a@x.com'; Start = [datetime]'2026-11-09'; End = [datetime]'2026-11-10' }).Count | Should -Be 1
    }
    It 'refuses a period longer than 31 days' {
        @(Test-EmtCriteria @{ Subject = 'x'; Start = [datetime]'2026-09-01'; End = [datetime]'2026-11-10' }).Count | Should -Be 1
    }
}

Describe 'Complete-EmtCriteria' {
    It 'adds missing keys and keeps given values' {
        $c = Complete-EmtCriteria @{ Subject = 'x'; Senders = @('a@x.com', '') }
        $c.Subject         | Should -Be 'x'
        $c.Senders         | Should -Be @('a@x.com')
        $c.MessageId       | Should -Be ''
        $c.IncludeForwards | Should -BeTrue
    }
}

Describe 'Get-EmtSearchHits' {
    It 'parses the comma format' {
        $h = @(Get-EmtSearchHits '{Location: A@x.com, Item count: 2, Total size: 2048, Location: b@x.com, Item count: 0, Total size: 0}')
        $h.Count | Should -Be 2
        $h[0].Mailbox | Should -Be 'a@x.com'
        $h[0].Items   | Should -Be 2
        $h[0].SizeKB  | Should -Be 2
    }
    It 'parses the semicolon format' {
        $h = @(Get-EmtSearchHits "Location: a@x.com; Item count: 1; Total size: 1024`r`nLocation: c@x.com; Item count: 3; Total size: 3072")
        $h.Count | Should -Be 2
        $h[1].Items | Should -Be 3
    }
    It 'returns nothing for empty text' {
        @(Get-EmtSearchHits '').Count | Should -Be 0
    }
}

Describe 'Test-EmtServerGeneratedId' {
    It 'marks a .local host' {
        Test-EmtServerGeneratedId '<5f3c@EXCH01.Contoso.local>' | Should -BeTrue
    }
    It 'marks a transport server name' {
        Test-EmtServerGeneratedId '<abc@exch1.contoso.com>' -ServerNames 'EXCH1' | Should -BeTrue
    }
    It 'does not mark a normal client MessageId' {
        Test-EmtServerGeneratedId '<abc@contoso.com>' -ServerNames 'EXCH1' | Should -BeFalse
    }
}

Describe 'Get-EmtMessageSummary' {
    It 'makes one row per MessageId with the first RECEIVE and all recipients' {
        $t = [datetime]'2026-11-09 16:53'
        $ev = @(
            New-Ev '<1@contoso.com>' 'A@x.com' 'Report' 'RECEIVE' @('b@x.com', 'c@x.com') $t
            New-Ev '<1@contoso.com>' 'a@x.com' 'Report' 'DELIVER' @('b@x.com') $t.AddSeconds(2)
            New-Ev '<2@EXCH1.contoso.local>' 'b@x.com' 'FW: Report' 'RECEIVE' @('d@y.com') $t.AddSeconds(20)
        )
        $s = @(Get-EmtMessageSummary -Events $ev)
        $s.Count | Should -Be 2
        $s[0].Sender     | Should -Be 'a@x.com'
        $s[0].Recipients | Should -Be 2
        $s[0].Note       | Should -BeNullOrEmpty
        $s[1].Note       | Should -Not -BeNullOrEmpty
        $s[0].Expanded   | Should -Be 0
    }
    It 'adds the RecipientCount of the EXPAND events (a message to a group)' {
        $t = [datetime]'2026-11-09 16:53'
        $exp1 = New-Ev 'm1' 'a@x.com' 'News' 'EXPAND' @('b@x.com') $t.AddSeconds(1)
        $exp1 | Add-Member RecipientCount 670
        $exp2 = New-Ev 'm1' 'a@x.com' 'News' 'EXPAND' @('b@x.com') $t.AddSeconds(2)
        $exp2 | Add-Member RecipientCount 12
        $agent = New-Ev 'm1' 'a@x.com' 'News' 'AGENTINFO' @('b@x.com') $t.AddSeconds(3)
        $agent | Add-Member RecipientCount 673
        $ev = @((New-Ev 'm1' 'a@x.com' 'News' 'RECEIVE' @('all@x.com') $t), $exp1, $exp2, $agent)
        $s = @(Get-EmtMessageSummary -Events $ev)
        $s[0].Expanded | Should -Be 682
    }
}

Describe 'Get-EmtMissingMailboxes' {
    It 'returns searched mailboxes without items' {
        $hits = @([pscustomobject]@{ Mailbox = 'a@x.com'; Items = 1 }, [pscustomobject]@{ Mailbox = 'b@x.com'; Items = 0 })
        Get-EmtMissingMailboxes -Locations 'A@x.com', 'b@x.com', 'c@x.com' -Hits $hits | Should -Be @('b@x.com', 'c@x.com')
    }
    It 'returns nothing for an All search' {
        @(Get-EmtMissingMailboxes -Locations 'All' -Hits @()).Count | Should -Be 0
    }
}

Describe 'New-EmtSearchName' {
    It 'uses the EMT_ prefix and a timestamp' {
        New-EmtSearchName -Now ([datetime]'2026-11-10 11:03:45') | Should -Be 'EMT_20261110_110345'
    }
}

Describe 'Write-EmtAudit' {
    It 'appends rows with fixed columns' {
        $p = Join-Path $TestDrive 'audit\audit.csv'
        Write-EmtAudit -Path $p -Record @{ Action = 'Search'; Items = 5; Unknown = 'ignored' }
        Write-EmtAudit -Path $p -Record @{ Action = 'SoftDelete'; Result = 'Clean' }
        $rows = @(Import-Csv $p)
        $rows.Count | Should -Be 2
        $rows[0].Action | Should -Be 'Search'
        $rows[0].Items  | Should -Be '5'
        $rows[1].Result | Should -Be 'Clean'
        $rows[0].PSObject.Properties.Name | Should -Not -Contain 'Unknown'
    }
}

Describe 'Resolve-EmtRecipients' {
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
        $r = @(Resolve-EmtRecipients -Addresses 'u@x.com', 'alias@x.com', 'c@y.com', 'g@x.com', 'z@ext.org')
        ($r | Where-Object Address -eq 'u@x.com').Kind        | Should -Be 'Internal'
        ($r | Where-Object Address -eq 'u@x.com').Mailbox     | Should -Be 'u@x.com'
        ($r | Where-Object Address -eq 'alias@x.com').Mailbox | Should -Be 'u@x.com'
        ($r | Where-Object Address -eq 'c@y.com').Kind        | Should -Be 'External'
        ($r | Where-Object Address -eq 'g@x.com').Kind        | Should -Be 'Group'
        ($r | Where-Object Address -eq 'z@ext.org').Kind      | Should -Be 'External'
    }
    It 'uses one call per batch' {
        $null = Resolve-EmtRecipients -Addresses 'u@x.com', 'c@y.com', 'g@x.com' -BatchSize 2
        Should -Invoke Get-Recipient -Times 2 -Exactly
    }
    It 'does not look up addresses outside the accepted domains' {
        $r = @(Resolve-EmtRecipients -Addresses 'u@x.com', 'z@ext.org' -AcceptedDomains 'x.com')
        Should -Invoke Get-Recipient -Times 1 -Exactly
        ($r | Where-Object Address -eq 'z@ext.org').Kind | Should -Be 'External'
    }
}

Describe 'Resolve-EmtRecipients with DELIVER addresses' {
    BeforeEach {
        $global:EmtPublicFolders = $null
        $script:Directory = @(
            New-Rcpt 'UserMailbox' 'u@x.com' @('SMTP:u@x.com')
            New-Rcpt 'UserMailbox' 'v@x.com' @('SMTP:v@x.com')
            New-Rcpt 'PublicFolder' 'pf@x.com' @('SMTP:pf@x.com', 'smtp:pf2@x.com')
            New-Rcpt 'MailUniversalDistributionGroup' 'all@x.com' @('SMTP:all@x.com')
        )
        Mock Get-Recipient {
            if ($RecipientType -eq 'PublicFolder') { return @($script:Directory | Where-Object RecipientType -eq 'PublicFolder') }
            $wanted = @([regex]::Matches($Filter, "smtp:([^']+)") | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() })
            $script:Directory | Where-Object {
                $p = @($_.EmailAddresses | ForEach-Object { ($_ -replace '^smtp:', '').ToLowerInvariant() })
                @($wanted | Where-Object { $_ -in $p }).Count
            }
        }
    }
    AfterAll { $global:EmtPublicFolders = $null }
    It 'takes a delivered address as a mailbox without a lookup' {
        $r = @(Resolve-EmtRecipients -Addresses 'u@x.com', 'v@x.com', 'all@x.com' -AcceptedDomains 'x.com' -Delivered 'U@x.com', 'v@x.com')
        ($r | Where-Object Address -eq 'u@x.com').Kind    | Should -Be 'Internal'
        ($r | Where-Object Address -eq 'u@x.com').Mailbox | Should -Be 'u@x.com'
        ($r | Where-Object Address -eq 'v@x.com').Type    | Should -Be 'UserMailbox'
        ($r | Where-Object Address -eq 'all@x.com').Kind  | Should -Be 'Group'
        # One call for the public folders, one call for the group: no call for the two mailboxes.
        Should -Invoke Get-Recipient -Times 1 -Exactly -ParameterFilter { $RecipientType -eq 'PublicFolder' }
        Should -Invoke Get-Recipient -Times 1 -Exactly -ParameterFilter { $Filter -and $Filter -notmatch 'u@x\.com|v@x\.com' }
        Should -Invoke Get-Recipient -Times 2 -Exactly
    }
    It 'looks up a delivered address of a public folder' {
        $r = @(Resolve-EmtRecipients -Addresses 'u@x.com', 'pf2@x.com' -AcceptedDomains 'x.com' -Delivered 'u@x.com', 'pf2@x.com')
        ($r | Where-Object Address -eq 'pf2@x.com').Kind    | Should -Be 'Other'
        ($r | Where-Object Address -eq 'pf2@x.com').Mailbox | Should -Be ''
        Should -Invoke Get-Recipient -Times 1 -Exactly -ParameterFilter { $Filter -match 'pf2@x\.com' -and $Filter -notmatch 'u@x\.com' }
    }
    It 'gets the public folder list one time for each connection' {
        $null = Resolve-EmtRecipients -Addresses 'u@x.com' -Delivered 'u@x.com'
        $null = Resolve-EmtRecipients -Addresses 'v@x.com' -Delivered 'v@x.com'
        Should -Invoke Get-Recipient -Times 1 -Exactly -ParameterFilter { $RecipientType -eq 'PublicFolder' }
    }
    It 'looks up all addresses if the public folder list is not available' {
        Mock Get-Recipient { throw 'access denied' } -ParameterFilter { $RecipientType -eq 'PublicFolder' }
        $r = @(Resolve-EmtRecipients -Addresses 'u@x.com', 'v@x.com' -AcceptedDomains 'x.com' -Delivered 'u@x.com', 'v@x.com')
        ($r | Where-Object Address -eq 'u@x.com').Kind | Should -Be 'Internal'
        Should -Invoke Get-Recipient -Times 1 -Exactly -ParameterFilter { $Filter -match 'u@x\.com' -and $Filter -match 'v@x\.com' }
        # The failure stays until the next connection: no second try.
        $null = Resolve-EmtRecipients -Addresses 'u@x.com' -Delivered 'u@x.com'
        Should -Invoke Get-Recipient -Times 1 -Exactly -ParameterFilter { $RecipientType -eq 'PublicFolder' }
    }
    It 'does not change an address outside the accepted domains' {
        $r = @(Resolve-EmtRecipients -Addresses 'z@ext.org' -AcceptedDomains 'x.com' -Delivered 'z@ext.org')
        ($r | Where-Object Address -eq 'z@ext.org').Kind | Should -Be 'External'
        Should -Invoke Get-Recipient -Times 0 -Exactly
    }
}

Describe 'Get-EmtEventAddresses -EventId' {
    It 'takes the addresses of one event type only' {
        $t = [datetime]'2026-11-10'
        $ev = @(
            New-Ev 'm1' 'a@x.com' 's' 'RECEIVE' @('B@x.com', 'c@x.com') $t
            New-Ev 'm1' 'a@x.com' 's' 'DELIVER' @('B@x.com') $t
            [pscustomobject]@{ MessageId = 'm1'; Recipients = @('d@x.com') }
        )
        @(Get-EmtEventAddresses $ev -EventId 'DELIVER') | Should -Be @('b@x.com')
        @(Get-EmtEventAddresses $ev).Count | Should -Be 3
    }
}

Describe 'Test-EmtInternalDomain' {
    It 'matches an exact domain' { Test-EmtInternalDomain 'a@Contoso.com' 'contoso.com' | Should -BeTrue }
    It 'matches a subdomain of a wildcard domain' { Test-EmtInternalDomain 'a@mail.contoso.com' '*.contoso.com' | Should -BeTrue }
    It 'does not match another domain' { Test-EmtInternalDomain 'a@outlook.com' 'contoso.com', '*.contoso.com' | Should -BeFalse }
}

Describe 'ConvertFrom-EmtSizeText' {
    It 'reads the bytes in parentheses' { ConvertFrom-EmtSizeText '1.25 MB (1,310,720 bytes)' | Should -Be 1280 }
    It 'reads bytes with spaces' { ConvertFrom-EmtSizeText '2 KB (2 048 bytes)' | Should -Be 2 }
    It 'returns 0 for unknown text' { ConvertFrom-EmtSizeText '' | Should -Be 0 }
}

Describe 'Select-EmtUniqueEvents' {
    It 'removes duplicate events and keeps only the given MessageIds' {
        $t = [datetime]'2026-11-09 10:00'
        $a = New-Ev '<1@x>' 'a@x.com' 'S' 'DELIVER' @('b@x.com') $t
        $b = New-Ev '<2@x>' 'a@x.com' 'S' 'DELIVER' @('b@x.com') $t
        @(Select-EmtUniqueEvents -Events @($a, $a, $b) -MessageIds '<1@x>').Count | Should -Be 1
    }
}

Describe 'Complete-EmtCriteria and Test-EmtCriteria: only the recipient mailbox' {
    It 'sets OnlyRecipient to false by default' {
        (Complete-EmtCriteria @{}).OnlyRecipient | Should -BeFalse
    }
    It 'refuses OnlyRecipient without a recipient' {
        @(Test-EmtCriteria @{ Subject = 'x'; OnlyRecipient = $true; Start = [datetime]'2026-11-09'; End = [datetime]'2026-11-10' }).Count |
            Should -Be 1
    }
    It 'refuses OnlyRecipient together with AllMailboxes' {
        @(Test-EmtCriteria @{ Subject = 'x'; Recipient = 'a@x.com'; OnlyRecipient = $true; AllMailboxes = $true
                             Start = [datetime]'2026-11-09'; End = [datetime]'2026-11-10' }).Count | Should -Be 1
    }
}

Describe 'Find-EmtTrackingEvents' {
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
        $r = Find-EmtTrackingEvents -Senders 'a@x.com' -Subject 'Отчёт за квартал' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards
        @($r.Events | Select-Object -ExpandProperty MessageId -Unique) | Should -Be @('<1@x.com>', '<2@x.com>')
    }
    It 'finds only the original when forwards are excluded' {
        $r = Find-EmtTrackingEvents -Senders 'a@x.com' -Subject 'Отчёт за квартал' -Start $t.Date -End $t.Date.AddDays(1)
        @($r.Events | Select-Object -ExpandProperty MessageId -Unique) | Should -Be @('<1@x.com>')
    }
    It 'finds the message by subject when the sender is wrong' {
        $r = Find-EmtTrackingEvents -Senders 'wrong@x.com' -Subject 'Отчёт за квартал' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards
        @($r.Events | Select-Object -ExpandProperty MessageId -Unique).Count | Should -Be 2
    }
    It 'does not do a second pass by MessageId without a recipient filter' {
        $null = Find-EmtTrackingEvents -Senders 'a@x.com' -Subject 'Отчёт за квартал' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards -CompleteRecipients
        # One sender query + one subject query on one server.
        Should -Invoke Get-MessageTrackingLog -Times 2 -Exactly
    }
    It 'does a second pass by MessageId with a recipient filter' {
        $null = Find-EmtTrackingEvents -Subject 'Отчёт за квартал' -Recipient 'b@x.com' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards -CompleteRecipients
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
        $r = Find-EmtTrackingEvents -Senders 'a@x.com' -Subject 'Итоги' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards
        @($r.Events | Select-Object -ExpandProperty MessageId -Unique | Sort-Object) | Should -Be @('<1@x.com>', '<2@x.com>')
    }
    It 'matches a subject with square brackets as text' {
        Mock Get-MessageTrackingLog {
            @(New-Ev '<1@x.com>' 'a@x.com' '[EXTERNAL] Отчёт' 'RECEIVE' @('b@x.com') $t)
        }
        $r = Find-EmtTrackingEvents -Senders 'a@x.com' -Subject '[EXTERNAL] Отчёт' -Start $t.Date -End $t.Date.AddDays(1)
        @($r.Events).Count | Should -Be 1
        { Find-EmtTrackingEvents -Subject '[EXTERNAL Отчёт' -Start $t.Date -End $t.Date.AddDays(1) } | Should -Not -Throw
    }
    It 'returns the errors of a server that fails' {
        Mock Get-TransportService { [pscustomobject]@{ Name = 'EXCH1' }; [pscustomobject]@{ Name = 'EXCH2' } }
        Mock Get-MessageTrackingLog {
            if ($Server -eq 'EXCH2') { Write-Error 'server is not available' ; return }
            @(New-Ev '<1@x.com>' 'a@x.com' 'Отчёт' 'RECEIVE' @('b@x.com') $t)
        }
        $r = Find-EmtTrackingEvents -Senders 'a@x.com' -Start $t.Date -End $t.Date.AddDays(1)
        @($r.Errors).Count | Should -Be 1
        $r.Errors[0] | Should -Match 'EXCH2'
        @($r.Events).Count | Should -Be 1
    }
    It 'does not do a second pass without -CompleteRecipients' {
        $null = Find-EmtTrackingEvents -Subject 'Отчёт за квартал' -Recipient 'b@x.com' -Start $t.Date -End $t.Date.AddDays(1) -IncludeForwards
        Should -Invoke Get-MessageTrackingLog -Times 1 -Exactly
    }
}

Describe 'Test-EmtCancel' {
    AfterEach { Remove-Variable EmtSync -Scope Global -ErrorAction SilentlyContinue }
    It 'throws when cancel is requested' {
        $global:EmtSync = @{ Cancel = $true; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        { Test-EmtCancel } | Should -Throw
    }
    It 'does nothing without cancel' {
        $global:EmtSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        { Test-EmtCancel } | Should -Not -Throw
    }
}

Describe 'Language files (Lang\*.lng)' {
    BeforeAll {
        $root = Join-Path $PSScriptRoot '..'
        $langDir = Join-Path $root 'Lang'
        $en = Read-EmtLanguageFile (Join-Path $langDir 'en.lng')
        $ru = Read-EmtLanguageFile (Join-Path $langDir 'ru.lng')
    }
    AfterEach { $global:EmtLang = 'ru' }

    It 'has the bundled files ru.lng and en.lng with names' {
        $en.Code | Should -Be 'en'
        $ru.Code | Should -Be 'ru'
        $en.Name | Should -Be 'English'
        $ru.Name | Should -Not -BeNullOrEmpty
    }
    It 'ru.lng and en.lng are complete and correct' {
        @(Test-EmtLanguageFile (Join-Path $langDir 'ru.lng')) | Should -BeNullOrEmpty
        @(Test-EmtLanguageFile (Join-Path $langDir 'en.lng')) | Should -BeNullOrEmpty
    }
    It 'all other .lng files have no unknown keys and correct placeholders' {
        foreach ($f in Get-ChildItem $langDir -Filter '*.lng' | Where-Object Name -notin 'ru.lng', 'en.lng') {
            @(Test-EmtLanguageFile $f.FullName -AllowMissing) | Should -BeNullOrEmpty -Because $f.Name
        }
    }
    It 'every text key used in the code and in the window files is in en.lng' {
        $files = @(Join-Path $root 'ExchangeMessageTrail.ps1'; Join-Path $root 'ExchangeMessageTrail.Classic.ps1'
                   Join-Path $root 'src\EMT.Core.ps1')
        $used = foreach ($f in $files) {
            $src = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
            [regex]::Matches($src, "\bT\s+'([A-Za-z0-9_.]+)'") | ForEach-Object { $_.Groups[1].Value }
            # Keys given to helpers that call T: New-EmtButton 'act.soft', Set-EmtLoc $x 'crit.group', ...
            [regex]::Matches($src, "(?:New-EmtLabel|New-EmtButton|New-EmtCheckBox|Set-EmtLoc\s+\S+|Set-EmtTip\s+\S+)\s+'([A-Za-z0-9_.]+)'") |
                ForEach-Object { $_.Groups[1].Value }
        }
        # The window files: Tag="key" of the texts.
        foreach ($x in Get-ChildItem (Join-Path $root 'gui') -Filter '*.xaml') {
            $xml = [IO.File]::ReadAllText($x.FullName, [Text.Encoding]::UTF8)
            $used = @($used) + @([regex]::Matches($xml, '\bTag="([A-Za-z0-9_.]+)"') | ForEach-Object { $_.Groups[1].Value })
        }
        # Keys made at run time: "tab.$k", "trtab.$k", "kind.$($_.Kind)" and "theme.$t".
        $used = @($used) + @('messages', 'tracking', 'recipients', 'found', 'missing', 'check' | ForEach-Object { "tab.$_" }) +
                @('recipients', 'events', 'mailbox' | ForEach-Object { "trtab.$_" }) +
                @('Internal', 'External', 'Group', 'Other' | ForEach-Object { "kind.$_" }) +
                @('Dark', 'Light' | ForEach-Object { "theme.$_" })
        $missing = @($used | Sort-Object -Unique | Where-Object { $_ -and -not $en.Texts.ContainsKey($_) })
        $missing | Should -BeNullOrEmpty
    }
    It 'gives the text in the current language and fills the placeholders' {
        $global:EmtLang = 'en'
        T 'log.found' 3 2 1.5 | Should -Match '^Messages found: 3 in 2 mailboxes'
        $global:EmtLang = 'ru'
        T 'log.found' 3 2 1.5 | Should -Match '^Найдено писем: 3 в 2 ящиках'
    }
    It 'uses English for a key that a language does not have, and the key for an unknown key' {
        $global:EmtLang = 'xx'
        T 'conn.connect' | Should -Be 'Connect'
        T 'no.such.key' | Should -Be 'no.such.key'
    }
}

Describe 'Read-EmtLanguageFile and Test-EmtLanguageFile' {
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
        $l = Read-EmtLanguageFile $f
        $l.Code | Should -Be 'de'
        $l.Name | Should -Be 'Deutsch'
        $l.Author | Should -Be 'Test'
        $l.Texts['a.one'] | Should -Be 'Eins {0} = gleich'
        $l.Texts['a.two'] | Should -Be 'Zwei'
        @(Test-EmtLanguageFile $f -ReferencePath $ref) | Should -BeNullOrEmpty
    }
    It 'uses the file name when there is no @code line' {
        $f = Join-Path $TestDrive 'pl.lng'
        Set-Content -Path $f -Encoding UTF8 -Value @('a.two = Dwa')
        (Read-EmtLanguageFile $f).Code | Should -Be 'pl'
    }
    It 'finds unknown keys, missing keys, wrong placeholders, duplicates and bad lines' {
        $f = Join-Path $TestDrive 'bad.lng'
        Set-Content -Path $f -Encoding UTF8 -Value @(
            '@code = bad', '@name = Bad', 'a.one = Eins', 'a.one = Eins {0}', 'a.three = Drei', 'no equals sign')
        $p = @(Test-EmtLanguageFile $f -ReferencePath $ref)
        ($p -join "`n") | Should -Match "duplicate key 'a.one'"
        ($p -join "`n") | Should -Match "unknown key 'a.three'"
        ($p -join "`n") | Should -Match "missing key 'a.two'"
        ($p -join "`n") | Should -Match 'no .key = text.'
    }
    It 'finds a wrong placeholder' {
        $f = Join-Path $TestDrive 'ph.lng'
        Set-Content -Path $f -Encoding UTF8 -Value @('@name = Ph', 'a.one = Eins {1}', 'a.two = Zwei')
        (@(Test-EmtLanguageFile $f -ReferencePath $ref) -join "`n") | Should -Match "key 'a.one': placeholders"
    }
}
Describe 'Events without MessageId' {
    It 'are not a row in the message summary' {
        $t = [datetime]'2026-11-10 08:59:45'
        $ev = @(
            New-Ev ''          'a@x.com' ''          'NOTIFYMAPI' @() $t
            New-Ev '<1@x.com>' 'a@x.com' 'RE: Topic' 'RECEIVE'    @('b@x.com') $t
        )
        $s = @(Get-EmtMessageSummary -Events $ev)
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
        $r = Find-EmtTrackingEvents -Senders 'a@x.com' -Start $t.Date -End $t.Date.AddDays(1)
        @($r.Events).Count | Should -Be 2
        @($r.Events | Where-Object { $_.MessageId } | Select-Object -ExpandProperty MessageId -Unique) | Should -Be @('<1@x.com>')
    }
}

Describe 'EMT search names' {
    It 'accepts only the names that EMT makes (EMT_, and the older MC_)' {
        Test-EmtSearchName 'EMT_20261110_124718'       | Should -BeTrue
        Test-EmtSearchName 'EMT_20261110_124718-shadow' | Should -BeTrue
        Test-EmtSearchName 'MC_20261110_124718'        | Should -BeTrue
        Test-EmtSearchName 'MC_20261110_124718-shadow' | Should -BeTrue
        Test-EmtSearchName 'MC_quick_20261110_124718'  | Should -BeFalse
        Test-EmtSearchName 'ContentSearch_for_Delete1' | Should -BeFalse
        Test-EmtSearchName 'mc_20261110_124718'        | Should -BeFalse
    }
    It 'reads the time from the name' {
        Get-EmtSearchNameTime 'MC_20261110_124718-shadow' | Should -Be ([datetime]'2026-11-10 12:47:18')
        Get-EmtSearchNameTime 'EMT_20261110_124718' | Should -Be ([datetime]'2026-11-10 12:47:18')
        Get-EmtSearchNameTime 'Other' | Should -BeNullOrEmpty
    }
}

Describe 'Get-EmtOldSearches and Remove-EmtSearches' {
    BeforeAll {
        $old = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        $new = (Get-Date).AddMinutes(-5).ToString('yyyyMMdd_HHmmss')
        # The full list for Get-EmtOldSearches. By -Identity: "not found" (the search is removed).
        Mock Get-ComplianceSearch {
            if ($Identity) { throw "The operation couldn't be performed because '$Identity' couldn't be found." }
            @([pscustomobject]@{ Name = "MC_$old" }, [pscustomobject]@{ Name = "MC_$new" },
              [pscustomobject]@{ Name = 'ContentSearch_for_Delete1' })
        }
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
        $l = @(Get-EmtOldSearches -OlderThanDays 7)
        @($l | ForEach-Object Name | Sort-Object) | Should -Be @("MC_$old", "MC_$old-shadow")
        ($l | Where-Object Name -eq "MC_$old-shadow").Kind | Should -Be 'Shadow'
    }
    It 'lists today''s shadow search with 0 days' {
        @(Get-EmtOldSearches -OlderThanDays 0 | ForEach-Object Name) | Should -Contain "MC_$new-shadow"
    }
    It 'removes the search and its shadow once, and skips other names' {
        Remove-EmtSearches -Names "MC_$old", "MC_$old-shadow", 'Legal hold 2025' -RecheckSeconds 0
        Should -Invoke Remove-ComplianceSearch -Times 1 -Exactly -ParameterFilter { $Identity -eq "MC_$old" }
        Should -Invoke Remove-MailboxSearch -Times 1 -Exactly -ParameterFilter { $Identity -eq "MC_$old-shadow" }
        Should -Invoke Remove-MailboxSearch -Times 0 -Exactly -ParameterFilter { $Identity -eq 'Legal hold 2025' }
    }
}
Describe 'Partial addresses' {
    It 'keeps a name with spaces as one item, and splits full addresses at spaces' {
        Split-EmtAddressList 'Ivanov Ivan; a@x.com b@x.com' | Should -Be @('a@x.com', 'b@x.com', 'ivanov ivan')
        Split-EmtAddressList "jsmi`r`nc@x.com" | Should -Be @('c@x.com', 'jsmi')
    }
    It 'knows a full address' {
        Test-EmtFullAddress 'a.b@contoso.com' | Should -BeTrue
        Test-EmtFullAddress 'jsmith'    | Should -BeFalse
        Test-EmtFullAddress 'jsmi@contoso'  | Should -BeFalse
        Test-EmtFullAddress 'Ivanov Ivan'     | Should -BeFalse
    }
    It 'finds the candidates in the address book (ANR)' {
        Mock Get-Recipient {
            if ($Anr -eq 'ivan') {
                [pscustomobject]@{ PrimarySmtpAddress = 'I.Ivanov@x.com'; DisplayName = 'Ivanov Ivan'; RecipientTypeDetails = 'UserMailbox' }
                [pscustomobject]@{ PrimarySmtpAddress = 'ivanova@x.com'; DisplayName = 'Ivanova Anna'; RecipientTypeDetails = 'UserMailbox' }
            }
        }
        $c = @(Find-EmtAddressCandidates -Text 'ivan')
        $c.Count | Should -Be 2
        $c[0].Address | Should -Be 'i.ivanov@x.com'
        $c[0].Name    | Should -Be 'Ivanov Ivan'
        $c[0].Type    | Should -Be 'UserMailbox'
        @(Find-EmtAddressCandidates -Text 'nobody').Count | Should -Be 0
    }
    It 'returns no candidates when the address book search fails' {
        Mock Get-Recipient { throw 'access denied' }
        @(Find-EmtAddressCandidates -Text 'ivan').Count | Should -Be 0
    }
    It 'gives one result for each input' {
        Mock Get-Recipient { if ($Anr -eq 'jsmi') { [pscustomobject]@{ PrimarySmtpAddress = 'jsmith@x.com'; DisplayName = 'V S'; RecipientType = 'UserMailbox' } } }
        $r = @(Resolve-EmtAddressInputs -Inputs 'jsmi', 'nobody')
        $r.Count | Should -Be 2
        @($r[0].Candidates).Count | Should -Be 1
        $r[0].Candidates[0].Address | Should -Be 'jsmith@x.com'
        @($r[1].Candidates).Count | Should -Be 0
    }
}
Describe 'Remove-EmtSearches when Exchange removes the shadow search itself' {
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
        { Remove-EmtSearches -Names "MC_$stamp-shadow", "MC_$stamp" -RecheckSeconds 0 } | Should -Not -Throw
        Should -Invoke Remove-ComplianceSearch -Times 1 -Exactly
        Should -Invoke Remove-MailboxSearch -Times 0 -Exactly
    }
    It 'continues with the next search after an error' {
        Mock Remove-ComplianceSearch { if ($Identity -eq "MC_$stamp") { throw 'busy' } }
        { Remove-EmtSearches -Names "MC_$stamp", "MC_$stamp2" -RecheckSeconds 0 } | Should -Not -Throw
        Should -Invoke Remove-ComplianceSearch -Times 1 -Exactly -ParameterFilter { $Identity -eq "MC_$stamp2" }
    }
}
Describe 'Remove-EmtSearches when Remove-ComplianceSearch returns an error' {
    BeforeAll {
        $stamp = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        Mock Get-MailboxSearch { }
        Mock Remove-MailboxSearch { }
        Mock Remove-ComplianceSearch { throw [System.NullReferenceException]::new('Object reference not set to an instance of an object.') }
    }
    It 'reports a removal (green) when the search is gone after the internal error' {
        Mock Get-ComplianceSearch { }
        Mock Get-ComplianceSearchAction { throw [System.NullReferenceException]::new('Object reference not set to an instance of an object.') }
        $global:EmtSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        try {
            Remove-EmtSearches -Names "MC_$stamp" -RecheckSeconds 0
            $lines = @($global:EmtSync.Log.ToArray())
        } finally { Remove-Variable EmtSync -Scope Global -ErrorAction SilentlyContinue }
        ($lines -join "`n") | Should -Match '\[OK   \].*MC_'
        ($lines -join "`n") | Should -Not -Match '\[ERROR\]'
    }
    It 'reports an error when the search is still there after the error' {
        Mock Get-ComplianceSearch { [pscustomobject]@{ Name = "MC_$stamp" } }
        Mock Get-ComplianceSearchAction { }
        $global:EmtSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        try {
            Remove-EmtSearches -Names "MC_$stamp" -RecheckSeconds 0
            $lines = @($global:EmtSync.Log.ToArray())
        } finally { Remove-Variable EmtSync -Scope Global -ErrorAction SilentlyContinue }
        ($lines -join "`n") | Should -Match '\[ERROR\].*MC_'
    }
}
Describe 'T with the values 0, empty text and $false' {
    AfterEach { $global:EmtLang = 'ru' }
    It 'puts 0 in the text (the case from the log: "Писем (MessageId) в журнале: {0}.")' {
        $global:EmtLang = 'ru'
        T 'log.trackCount' 0 | Should -Be 'Писем (MessageId) в журнале: 0.'
        $global:EmtLang = 'en'
        T 'stage.recipients' 0 | Should -Be 'Addresses resolved: 0'
    }
    It 'puts an empty text and $false in the text' {
        $global:EmtLang = 'en'
        T 'log.connected' '' | Should -Be 'Connected to .'
        T 'log.connected' $false | Should -Be 'Connected to False.'
    }
}
Describe 'Remove-EmtSearches when Exchange removes the search some seconds after the error' {
    It 'reports a removal when the search goes from the list on a later check (the case from the log)' {
        $stamp = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        $global:EmtTestChecks = 0
        Mock Get-MailboxSearch { }
        Mock Get-ComplianceSearchAction { }
        Mock Remove-ComplianceSearch { throw [System.NullReferenceException]::new('Object reference not set to an instance of an object.') }
        # The first two checks still see the search, the third does not.
        Mock Get-ComplianceSearch { $global:EmtTestChecks++; if ($global:EmtTestChecks -le 2) { [pscustomobject]@{ Name = "MC_$stamp" } } }
        $global:EmtSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        try {
            Remove-EmtSearches -Names "MC_$stamp" -RecheckSeconds 0
            $lines = @($global:EmtSync.Log.ToArray())
        } finally {
            Remove-Variable EmtSync, EmtTestChecks -Scope Global -ErrorAction SilentlyContinue
        }
        ($lines -join "`n") | Should -Match '\[OK   \].*MC_'
        ($lines -join "`n") | Should -Not -Match '\[ERROR\]'
    }
}
Describe 'Remove-EmtSearches does not list all search actions' {
    It 'asks only for the action SEARCHNAME_Purge (the full list fails on some servers)' {
        $stamp = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        Mock Get-MailboxSearch { }
        Mock Remove-ComplianceSearch { }
        Mock Remove-ComplianceSearchAction { }
        Mock Get-ComplianceSearchAction {
            if (-not $Identity) { throw [System.NullReferenceException]::new('Object reference not set to an instance of an object.') }
            if ($Identity -eq "MC_${stamp}_Purge") { [pscustomobject]@{ Name = $Identity } }
        }
        { Remove-EmtSearches -Names "MC_$stamp" -RecheckSeconds 0 } | Should -Not -Throw
        Should -Invoke Get-ComplianceSearchAction -Times 0 -Exactly -ParameterFilter { -not $Identity }
        Should -Invoke Remove-ComplianceSearchAction -Times 1 -Exactly -ParameterFilter { $Identity -eq "MC_${stamp}_Purge" }
        Should -Invoke Remove-ComplianceSearch -Times 1 -Exactly
    }
}
Describe 'Remove-EmtSearches: the order and the second attempt (tested on Exchange 2019)' {
    BeforeAll {
        $stamp = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        Mock Get-ComplianceSearchAction { }
    }
    It 'removes the shadow search before the Compliance Search' {
        $global:EmtTestOrder = New-Object System.Collections.Generic.List[string]
        Mock Get-MailboxSearch { if ($Identity -eq "MC_$stamp-shadow") { [pscustomobject]@{ Name = $Identity } } }
        Mock Remove-MailboxSearch { $global:EmtTestOrder.Add("shadow:$Identity") }
        Mock Remove-ComplianceSearch { $global:EmtTestOrder.Add("search:$Identity") }
        try {
            Remove-EmtSearches -Names "MC_$stamp", "MC_$stamp-shadow" -RecheckSeconds 0
            @($global:EmtTestOrder) | Should -Be @("shadow:MC_$stamp-shadow", "search:MC_$stamp")
        } finally { Remove-Variable EmtTestOrder -Scope Global -ErrorAction SilentlyContinue }
    }
    It 'removes the search with a second attempt when the first fails and the search is still there' {
        $global:EmtTestRemoves = 0
        Mock Get-MailboxSearch { }
        Mock Remove-ComplianceSearch {
            $global:EmtTestRemoves++
            if ($global:EmtTestRemoves -eq 1) { throw [System.NullReferenceException]::new('Object reference not set to an instance of an object.') }
        }
        # Exists after the first attempt; the second attempt works.
        Mock Get-ComplianceSearch { if ($global:EmtTestRemoves -lt 2) { [pscustomobject]@{ Name = "MC_$stamp" } } else { throw "The operation couldn't be performed because 'MC_$stamp' couldn't be found." } }
        $global:EmtSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        try {
            Remove-EmtSearches -Names "MC_$stamp" -RecheckSeconds 0
            $lines = @($global:EmtSync.Log.ToArray())
        } finally { Remove-Variable EmtSync, EmtTestRemoves -Scope Global -ErrorAction SilentlyContinue }
        Should -Invoke Remove-ComplianceSearch -Times 2 -Exactly
        ($lines -join "`n") | Should -Not -Match '\[ERROR\]'
    }
}

Describe 'Test-EmtComplianceSearchExists' {
    It 'is false for "not found" and true for another error' {
        Mock Get-ComplianceSearch { throw "The operation couldn't be performed because 'X' couldn't be found." }
        Test-EmtComplianceSearchExists 'X' | Should -BeFalse
        Mock Get-ComplianceSearch { throw 'The server is not available.' }
        Test-EmtComplianceSearchExists 'X' | Should -BeTrue
        Mock Get-ComplianceSearch { [pscustomobject]@{ Name = 'X' } }
        Test-EmtComplianceSearchExists 'X' | Should -BeTrue
    }
}
Describe 'Remove-EmtSearches: lessons from Exchange 2019' {
    BeforeAll {
        $stamp = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        Mock Get-ComplianceSearchAction { }
        Mock Remove-MailboxSearch { }
        Mock Get-MailboxSearch { if ($Identity) { [pscustomobject]@{ Name = $Identity } } }
    }
    It 'does not touch a search that is still running, nor its shadow search' {
        Mock Get-ComplianceSearch { [pscustomobject]@{ Name = "MC_$stamp"; Status = 'InProgress' } }
        Mock Remove-ComplianceSearch { }
        Remove-EmtSearches -Names "MC_$stamp", "MC_$stamp-shadow" -RecheckSeconds 0
        Should -Invoke Remove-ComplianceSearch -Times 0 -Exactly
        Should -Invoke Remove-MailboxSearch -Times 0 -Exactly
    }
    It 'never calls Remove-ComplianceSearch with -ErrorAction Stop (it stops the remove in the middle)' {
        $global:EmtTestGone = $false
        Mock Get-ComplianceSearch { if (-not $global:EmtTestGone) { [pscustomobject]@{ Name = "MC_$stamp"; Status = 'Completed' } } else { throw "'MC_$stamp' couldn't be found." } }
        Mock Remove-ComplianceSearch { $global:EmtTestGone = $true }
        try {
            Remove-EmtSearches -Names "MC_$stamp" -RecheckSeconds 0
        } finally { Remove-Variable EmtTestGone -Scope Global -ErrorAction SilentlyContinue }
        Should -Invoke Remove-ComplianceSearch -Times 1 -Exactly
        Should -Invoke Remove-ComplianceSearch -Times 0 -Exactly -ParameterFilter { $PSBoundParameters['ErrorAction'] -eq 'Stop' }
    }
}
Describe 'Recipient in the query and in the search scope' {
    It 'adds the recipient to the KQL query (To, Cc, Bcc)' {
        $q = New-EmtSearchQuery -Senders 'a@x.com' -Recipients 'B@x.com' -Start ([datetime]'2026-11-10') -End ([datetime]'2026-11-11') -UtcOffset 0
        $q | Should -Match '\(to:b@x\.com OR cc:b@x\.com OR bcc:b@x\.com\)'
        $q | Should -Match '\(from:a@x\.com\)'
    }
    It 'searches only the recipient and sender mailboxes when the tracking log is empty (the case from the screenshot)' {
        Mock Get-EmtCapabilities { [pscustomobject]@{ Tracking = $true; Search = $true; Verify = $true; SoftDelete = $true; HardDelete = $true } }
        Mock Find-EmtTrackingEvents { [pscustomobject]@{ Events = @(); Servers = @('EXCH1'); Errors = @() } }
        Mock Get-EmtAcceptedDomains { @('x.com') }
        Mock Resolve-EmtRecipients {
            @($Addresses | ForEach-Object { [pscustomobject]@{ Address = $_; Kind = 'Internal'; Mailbox = $_; Type = 'UserMailbox' } })
        }
        Mock Invoke-EmtQuickSearch { [pscustomobject]@{ Mailbox = 'b@x.com'; Items = 0; SizeKB = 0; Error = '' } }
        Mock New-EmtComplianceSearchRun { throw 'must not search all mailboxes' }
        $r = Invoke-EmtSearch -Criteria @{
            Senders = @('a@x.com'); Recipient = 'b@x.com'; Start = [datetime]'2026-11-10'; End = [datetime]'2026-11-11'; IncludeForwards = $false
        }
        @($r.Locations | Sort-Object) | Should -Be @('a@x.com', 'b@x.com')
        # Nothing found anywhere: no list of mailboxes without the message.
        @($r.Missing).Count | Should -Be 0
        $r.Query | Should -Match 'to:b@x\.com'
        Should -Invoke New-EmtComplianceSearchRun -Times 0 -Exactly
    }
    It 'marks a result that the tracking log does not confirm' {
        Mock Get-EmtCapabilities { [pscustomobject]@{ Tracking = $true; Search = $true; Verify = $true; SoftDelete = $true; HardDelete = $true } }
        Mock Find-EmtTrackingEvents { [pscustomobject]@{ Events = @(); Servers = @('EXCH1'); Errors = @() } }
        Mock Get-EmtAcceptedDomains { @('x.com') }
        Mock Resolve-EmtRecipients { @($Addresses | ForEach-Object { [pscustomobject]@{ Address = $_; Kind = 'Internal'; Mailbox = $_; Type = 'UserMailbox' } }) }
        Mock Invoke-EmtQuickSearch { [pscustomobject]@{ Mailbox = 'a@x.com'; Items = 5; SizeKB = 10; Error = '' } }
        $r = Invoke-EmtSearch -Criteria @{
            Senders = @('a@x.com'); Recipient = 'b@x.com'; Start = [datetime]'2026-11-10'; End = [datetime]'2026-11-11'
        }
        $r.Unconfirmed | Should -BeTrue
    }
}
Describe 'Get-EmtSum' {
    It 'gives 0 for an empty list and the sum for a list' {
        Get-EmtSum -Items @() -Property Items | Should -Be 0
        Get-EmtSum -Items @([pscustomobject]@{ Items = 2 }, [pscustomobject]@{ Items = 3 }) -Property Items | Should -Be 5
        Get-EmtSum -Items @([pscustomobject]@{ SizeKB = 1.5 }, $null) -Property SizeKB | Should -Be 1.5
    }
}
Describe 'Test-EmtComplianceSearchExists with a typographic apostrophe' {
    It 'is false for not found with the apostrophe U+2019 (the case from the log)' {
        $msg = "The operation couldn$([char]0x2019)t be performed because 'X' couldn$([char]0x2019)t be found."
        Mock Get-ComplianceSearch { throw $msg }
        Test-EmtComplianceSearchExists 'X' | Should -BeFalse
    }
}
Describe 'Remove-EmtSearches when Exchange removes the search in the background' {
    It 'gives a warning, not an error, when the remove returned no error but the search is still visible' {
        $stamp = (Get-Date).AddDays(-10).ToString('yyyyMMdd_HHmmss')
        Mock Get-MailboxSearch { }
        Mock Get-ComplianceSearchAction { }
        Mock Remove-ComplianceSearch { }
        Mock Get-ComplianceSearch { [pscustomobject]@{ Name = "MC_$stamp"; Status = 'Completed' } }
        $global:EmtSync = @{ Cancel = $false; Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        try {
            Remove-EmtSearches -Names "MC_$stamp" -RecheckSeconds 0
            $lines = @($global:EmtSync.Log.ToArray())
        } finally { Remove-Variable EmtSync -Scope Global -ErrorAction SilentlyContinue }
        ($lines -join "`n") | Should -Match '\[WARN \].*MC_.*Completed'
        ($lines -join "`n") | Should -Not -Match '\[ERROR\]'
    }
}
Describe 'No -ErrorAction Stop for the Compliance Search cmdlets (lesson from Exchange 2019)' {
    It 'the core has no *-ComplianceSearch* call with -ErrorAction Stop' {
        $src = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\src\EMT.Core.ps1'), [Text.Encoding]::UTF8)
        $bad = @($src -split "`r?`n" | Where-Object { $_ -notmatch '^\s*#' -and $_ -match '\b\w+-ComplianceSearch\w*\b[^#]*-ErrorAction\s+Stop' })
        $bad | Should -BeNullOrEmpty
    }
    It 'Get-EmtComplianceSearchInfo: object = exists, "not found" = gone, another error = unknown' {
        Mock Get-ComplianceSearch { [pscustomobject]@{ Name = 'X'; Status = 'Completed' } }
        $i = Get-EmtComplianceSearchInfo 'X'
        $i.Exists | Should -BeTrue
        $i.Status | Should -Be 'Completed'
        Mock Get-ComplianceSearch { Write-Error "The operation couldn$([char]0x2019)t be performed because 'X' couldn$([char]0x2019)t be found." }
        (Get-EmtComplianceSearchInfo 'X').Exists | Should -BeFalse
        Mock Get-ComplianceSearch { Write-Error 'Object reference not set to an instance of an object.' }
        $i = Get-EmtComplianceSearchInfo 'X'
        $i.Exists | Should -BeNullOrEmpty
        $i.Error  | Should -Match 'Object reference'
        Should -Invoke Get-ComplianceSearch -Times 0 -ParameterFilter { $PSBoundParameters['ErrorAction'] -eq 'Stop' }
    }
}
Describe 'Trace: delivery report' {
    BeforeAll {
        $t = [datetime]'2026-11-10 10:00'
        function New-TEv($Id, $Kind, [string[]]$To, [string[]]$Status, $Time, $Server = 'EXCH1') {
            [pscustomobject]@{ MessageId = $Id; EventId = $Kind; Recipients = $To; RecipientStatus = $Status
                               Timestamp = $Time; ServerHostname = $Server; Source = 'X'; Sender = 'a@x.com'; MessageSubject = 'S' }
        }
    }
    AfterEach { $global:EmtLang = 'ru' }
    It 'gives Delivered after DEFER, and the later decisive event' {
        $ev = @(
            New-TEv '<1>' 'RECEIVE' @('b@x.com') @('') $t
            New-TEv '<1>' 'DEFER'   @('b@x.com') @('451 4.4.0 Primary target IP address responded') $t.AddMinutes(1)
            New-TEv '<1>' 'DELIVER' @('b@x.com') @('') $t.AddMinutes(5) 'EXCH2'
        )
        $r = @(Get-EmtDeliveryReport -Events $ev -Kinds @{ 'b@x.com' = 'Internal' })
        $r.Count | Should -Be 1
        $r[0].Status | Should -Be 'Delivered'
        $r[0].Server | Should -Be 'EXCH2'
    }
    It 'gives Failed with the server answer and a hint for 5.1.1' {
        $global:EmtLang = 'en'
        $ev = @(New-TEv '<1>' 'FAIL' @('nobody@x.com') @('550 5.1.1 RESOLVER.ADR.RecipNotFound; not found') $t)
        $r = @(Get-EmtDeliveryReport -Events $ev)
        $r[0].Status | Should -Be 'Failed'
        $r[0].Detail | Should -Match '5\.1\.1'
        $r[0].Hint   | Should -Match 'does not exist'
    }
    It 'gives External for SEND to an external address, InTransit for SEND to an internal one' {
        $ev = @(New-TEv '<1>' 'SEND' @('ext@outlook.com', 'b@x.com') @('250 2.6.0 Queued', '250 2.6.0 Queued') $t)
        $r = @(Get-EmtDeliveryReport -Events $ev -Kinds @{ 'ext@outlook.com' = 'External'; 'b@x.com' = 'Internal' })
        ($r | Where-Object Recipient -eq 'ext@outlook.com').Status | Should -Be 'External'
        ($r | Where-Object Recipient -eq 'b@x.com').Status | Should -Be 'InTransit'
    }
    It 'gives External for SENDEXTERNAL after TRANSFER (a hidden copy that a transport rule adds)' {
        $ev = @(
            New-TEv '<1>' 'RECEIVE'      @('monitor@security.example') @('') $t
            New-TEv '<1>' 'TRANSFER'     @('monitor@security.example') @('') $t.AddSeconds(1)
            New-TEv '<1>' 'SENDEXTERNAL' @('monitor@security.example') @('250 2.1.5 Ok') $t.AddSeconds(2)
        )
        $r = @(Get-EmtDeliveryReport -Events $ev -Kinds @{ 'monitor@security.example' = 'External' })
        $r.Count | Should -Be 1
        $r[0].Status | Should -Be 'External'
        $r[0].Detail | Should -Be '250 2.1.5 Ok'
    }
    It 'gives External for SENDEXTERNAL also for an internal kind: the message left the organization' {
        $ev = @(New-TEv '<1>' 'SENDEXTERNAL' @('b@x.com') @('250 OK') $t)
        $r = @(Get-EmtDeliveryReport -Events $ev -Kinds @{ 'b@x.com' = 'Internal' })
        $r[0].Status | Should -Be 'External'
    }
    It 'gives Delivered for DUPLICATEDELIVER' {
        $ev = @(
            New-TEv '<1>' 'TRANSFER'         @('b@x.com') @('') $t
            New-TEv '<1>' 'DUPLICATEDELIVER' @('b@x.com') @('') $t.AddSeconds(1)
        )
        $r = @(Get-EmtDeliveryReport -Events $ev -Kinds @{ 'b@x.com' = 'Internal' })
        $r[0].Status | Should -Be 'Delivered'
    }
    It 'gives Group for a group address, and Unknown for RECEIVE only' {
        $ev = @(New-TEv '<1>' 'RECEIVE' @('team@x.com', 'c@x.com') @('', '') $t)
        $r = @(Get-EmtDeliveryReport -Events $ev -Kinds @{ 'team@x.com' = 'Group'; 'c@x.com' = 'Internal' })
        ($r | Where-Object Recipient -eq 'team@x.com').Status | Should -Be 'Group'
        ($r | Where-Object Recipient -eq 'c@x.com').Status | Should -Be 'Unknown'
    }
    It 'keeps each message separate and ignores events without MessageId' {
        $ev = @(
            New-TEv '<1>' 'DELIVER' @('b@x.com') @('') $t
            New-TEv '<2>' 'FAIL'    @('b@x.com') @('550 5.2.2 mailbox full') $t
            New-TEv ''    'NOTIFYMAPI' @('b@x.com') @('') $t
        )
        $r = @(Get-EmtDeliveryReport -Events $ev)
        $r.Count | Should -Be 2
        ($r | Where-Object MessageId -eq '<2>').Status | Should -Be 'Failed'
    }
}

Describe 'Trace: hints and status lines' {
    AfterEach { $global:EmtLang = 'ru' }
    It 'explains the common SMTP codes' {
        $global:EmtLang = 'en'
        Get-EmtStatusHint '550 5.2.2 STOREDRV.Deliver; mailbox full' | Should -Match 'full'
        Get-EmtStatusHint '500 4.4.7 Message delayed' | Should -Match 'expired'
        Get-EmtStatusHint '550 5.7.1 RESOLVER.RST.NotAuthorized' | Should -Match 'rejected'
        Get-EmtStatusHint '451 4.4.0 DNS query failed' | Should -Match 'connection'
        Get-EmtStatusHint '554 5.9.9 something new' | Should -Match 'permanent'
    }
    It 'gives no hint for a success code or for no code' {
        Get-EmtStatusHint '250 2.6.0 Queued mail for delivery' | Should -Be ''
        Get-EmtStatusHint '' | Should -Be ''
        Get-EmtStatusHint 'no code here' | Should -Be ''
    }
    It 'makes the status line of a message' {
        $global:EmtLang = 'en'
        $rows = @(
            [pscustomobject]@{ Status = 'Delivered' }, [pscustomobject]@{ Status = 'Delivered' },
            [pscustomobject]@{ Status = 'Failed' }, [pscustomobject]@{ Status = 'External' }
        )
        Get-EmtMessageStatusLine $rows | Should -Be 'delivered: 2, external: 1, failed: 1'
        Get-EmtMessageStatusLine @() | Should -Be 'no recipients in the tracking log'
    }
    It 'gives hints for an empty trace result' {
        $global:EmtLang = 'en'
        $now = [datetime]'2026-11-10'
        $h = @(Get-EmtNoResultHints -Criteria @{ Senders = @('a@x.com'); Recipient = 'b@x.com'; Start = $now.AddDays(-40); End = $now } -Now $now)
        ($h -join "`n") | Should -Match '30 days'
        ($h -join "`n") | Should -Match 'Sender'
        ($h -join "`n") | Should -Match 'Recipient'
        ($h -join "`n") | Should -Match 'Sent Items'
    }
    It 'accepts a recipient alone for a trace, not for a delete' {
        $c = @{ Recipient = 'b@x.com'; Start = [datetime]'2026-11-10'; End = [datetime]'2026-11-11' }
        @(Test-EmtCriteria $c -ForTrace).Count | Should -Be 0
        @(Test-EmtCriteria $c).Count | Should -Be 1
    }
}

Describe 'Mailbox check: findings' {
    BeforeAll {
        function New-Data {
            param($Rules = @(), $Junk = $null, $Box = $null, $Stats = $null)
            [pscustomobject]@{ Rules = $Rules; RulesError = ''; Junk = $Junk; JunkError = ''; Box = $Box; BoxError = ''; Stats = $Stats; StatsError = '' }
        }
    }
    AfterEach { $global:EmtLang = 'ru' }
    It 'finds a rule that moves messages from this sender (Alert)' {
        $global:EmtLang = 'en'
        $rule = [pscustomobject]@{ Name = 'Archive'; Enabled = $true; From = @('Ivanov Ivan [SMTP:a@x.com]'); MoveToFolder = 'jsmith:\Archive'; DeleteMessage = $false }
        $f = @(Get-EmtMailboxFindings -Data (New-Data -Rules @($rule)) -Sender 'a@x.com' -Subject 'Report')
        $a = @($f | Where-Object Severity -eq 'Alert')
        $a.Count | Should -Be 1
        $a[0].Text | Should -Match 'Archive'
        $a[0].Text | Should -Match 'applies to this message'
    }
    It 'finds a rule with the sender in the Outlook form [EX:...] (no SMTP address)' {
        $dn = '/o=Contoso/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=abc123-Ivanov'
        $rule = [pscustomobject]@{ Name = 'Ivanov'; Enabled = $true; From = @("""Ivanov Ivan"" [EX:$dn]"); MoveToFolder = 'jsmith:\Ivanov'; DeleteMessage = $false }
        $data = New-Data -Rules @($rule)
        # Without the DN: as before, the condition does not apply.
        @(Get-EmtMailboxFindings -Data $data -Sender 'a@x.com')[0].Severity | Should -Be 'Info'
        $data | Add-Member SenderDn $dn.ToUpperInvariant()
        @(Get-EmtMailboxFindings -Data $data -Sender 'a@x.com')[0].Severity | Should -Be 'Alert'
        $data.SenderDn = '/o=Contoso/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=other'
        @(Get-EmtMailboxFindings -Data $data -Sender 'b@x.com')[0].Severity | Should -Be 'Info'
    }
    It 'reads the LegacyExchangeDN of the sender in the mailbox check' {
        Mock Invoke-EmtReadOnly {
            if ($Command -eq 'Get-Mailbox' -and $Params.Identity -eq 'a@x.com') {
                return [pscustomobject]@{ Value = [pscustomobject]@{ LegacyExchangeDN = '/o=C/cn=Recipients/cn=a' }; Error = '' }
            }
            [pscustomobject]@{ Value = $null; Error = '' }
        }
        (Invoke-EmtMailboxCheck -Mailbox 'j@x.com' -Sender 'a@x.com').Data.SenderDn | Should -Be '/o=C/cn=Recipients/cn=a'
        (Invoke-EmtMailboxCheck -Mailbox 'j@x.com').Data.SenderDn | Should -Be ''
        Should -Invoke Invoke-EmtReadOnly -Times 1 -Exactly -ParameterFilter { $Command -eq 'Get-Mailbox' -and $Params.Identity -eq 'a@x.com' }
    }
    It 'gives Info for a rule that does not apply, and ignores a disabled rule' {
        $r1 = [pscustomobject]@{ Name = 'Other'; Enabled = $true; SubjectContainsWords = @('invoice'); DeleteMessage = $true }
        $r2 = [pscustomobject]@{ Name = 'Off'; Enabled = $false; DeleteMessage = $true }
        $f = @(Get-EmtMailboxFindings -Data (New-Data -Rules @($r1, $r2)) -Sender 'a@x.com' -Subject 'Report')
        @($f | Where-Object Area -eq 'Rules').Count | Should -Be 1
        @($f | Where-Object Area -eq 'Rules')[0].Severity | Should -Be 'Info'
    }
    It 'gives Alert for a blocked sender or domain' {
        $junk = [pscustomobject]@{ Enabled = $true; BlockedSendersAndDomains = @('x.com'); TrustedSendersAndDomains = @() }
        $f = @(Get-EmtMailboxFindings -Data (New-Data -Junk $junk) -Sender 'a@x.com')
        @($f | Where-Object { $_.Area -eq 'Junk' -and $_.Severity -eq 'Alert' }).Count | Should -Be 1
    }
    It 'gives Alert for forwarding without a copy, Warn with a copy' {
        $box1 = [pscustomobject]@{ ForwardingSmtpAddress = 'smtp:ext@gmail.com'; DeliverToMailboxAndForward = $false }
        $box2 = [pscustomobject]@{ ForwardingSmtpAddress = 'smtp:ext@gmail.com'; DeliverToMailboxAndForward = $true }
        @(Get-EmtMailboxFindings -Data (New-Data -Box $box1) | Where-Object { $_.Area -eq 'Forwarding' })[0].Severity | Should -Be 'Alert'
        @(Get-EmtMailboxFindings -Data (New-Data -Box $box2) | Where-Object { $_.Area -eq 'Forwarding' })[0].Severity | Should -Be 'Warn'
    }
    It 'gives Warn for a mailbox that is 95% full' {
        $box = [pscustomobject]@{ UseDatabaseQuotaDefaults = $false; ProhibitSendReceiveQuota = '10 GB (10,737,418,240 bytes)' }
        $stats = [pscustomobject]@{ TotalItemSize = '9.5 GB (10,200,547,328 bytes)' }
        $q = @(Get-EmtMailboxFindings -Data (New-Data -Box $box -Stats $stats) | Where-Object Area -eq 'Quota')
        $q[0].Severity | Should -Be 'Warn'
    }
    It 'reports a check that did not run, and says that nothing explains the problem' {
        $global:EmtLang = 'en'
        $d = New-Data
        $d.RulesError = 'no cmdlet Get-InboxRule (no permission)'
        $f = @(Get-EmtMailboxFindings -Data $d)
        ($f | Where-Object Area -eq 'Rules')[0].Text | Should -Match 'Not examined'
        $clean = @(Get-EmtMailboxFindings -Data (New-Data))
        ($clean.Text -join "`n") | Should -Match 'No rule, block'
    }
    It 'knows the rule match types' {
        Get-EmtRuleMatch -Rule ([pscustomobject]@{}) -Sender 'a@x.com' -Subject 'S' | Should -Be 'All'
        Get-EmtRuleMatch -Rule ([pscustomobject]@{ BodyContainsWords = @('x') }) -Sender 'a@x.com' -Subject 'S' | Should -Be 'Unknown'
        Get-EmtRuleMatch -Rule ([pscustomobject]@{ SubjectContainsWords = @('rep') }) -Sender 'a@x.com' -Subject 'Report' | Should -Be 'Yes'
        Get-EmtRuleMatch -Rule ([pscustomobject]@{ FromAddressContainsWords = @('zzz') }) -Sender 'a@x.com' -Subject 'S' | Should -Be 'No'
    }
}

Describe 'Window files (gui\*.xaml)' {
    BeforeAll {
        $root = Join-Path $PSScriptRoot '..'
        $gui = Join-Path $root 'gui'
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
        function Read-Text([string] $Name) { [IO.File]::ReadAllText((Join-Path $gui $Name), [Text.Encoding]::UTF8) }
        function Get-Keys([string] $Xml) { @([regex]::Matches($Xml, 'x:Key="([^"{]+)"') | ForEach-Object { $_.Groups[1].Value }) }
    }
    It 'reads each theme and the styles without errors' {
        foreach ($n in 'Theme.Dark.xaml', 'Theme.Light.xaml', 'Styles.xaml') {
            { [void][System.Windows.Markup.XamlReader]::Parse((Read-Text $n)) } | Should -Not -Throw -Because $n
        }
    }
    It 'reads the main window without errors' -Skip:([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        { [void][System.Windows.Markup.XamlReader]::Parse((Read-Text 'MainWindow.xaml')) } | Should -Not -Throw
    }
    It 'has the same color keys in the dark and in the light theme' {
        $dark = Get-Keys (Read-Text 'Theme.Dark.xaml') | Sort-Object
        $light = Get-Keys (Read-Text 'Theme.Light.xaml') | Sort-Object
        Compare-Object $dark $light | Should -BeNullOrEmpty
    }
    It 'uses only resources that the theme or the styles have' {
        $known = @(Get-Keys (Read-Text 'Theme.Dark.xaml')) + @(Get-Keys (Read-Text 'Styles.xaml'))
        foreach ($n in 'Styles.xaml', 'MainWindow.xaml') {
            $used = @([regex]::Matches((Read-Text $n), '\{(?:Dynamic|Static)Resource ([A-Za-z0-9_.]+)\}') | ForEach-Object { $_.Groups[1].Value })
            @($used | Sort-Object -Unique | Where-Object { $_ -notin $known }) | Should -BeNullOrEmpty -Because $n
        }
        # The script gives some resources by name (cell templates, styles of the dialog buttons).
        $src = [IO.File]::ReadAllText((Join-Path $root 'ExchangeMessageTrail.ps1'), [Text.Encoding]::UTF8)
        $byName = @([regex]::Matches($src, "FindResource\('([A-Za-z0-9_.]+)'\)|-OkStyle '([A-Za-z0-9_.]+)'") |
                    ForEach-Object { if ($_.Groups[1].Value) { $_.Groups[1].Value } else { $_.Groups[2].Value } })
        @($byName | Where-Object { $_ -notin $known }) | Should -BeNullOrEmpty
    }
    It 'has an element with x:Name for each control variable of the script' {
        $names = @([regex]::Matches((Read-Text 'MainWindow.xaml'), 'x:Name="([A-Za-z0-9_]+)"') | ForEach-Object { $_.Groups[1].Value })
        $src = [IO.File]::ReadAllText((Join-Path $root 'ExchangeMessageTrail.ps1'), [Text.Encoding]::UTF8)
        $vars = @([regex]::Matches($src, '\$((?:tb|btn|chk|dt|grid|tab|lbl|cmb|pg|card|dot)[A-Z][A-Za-z0-9]*)') |
                  ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        @($vars | Where-Object { $_ -notin $names }) | Should -BeNullOrEmpty
    }
    It 'has the EMT icon with the sizes 16 to 256' {
        $ico = Join-Path $root 'gui\emt.ico'
        $dec = New-Object System.Windows.Media.Imaging.IconBitmapDecoder((New-Object Uri $ico), 'None', 'OnLoad')
        $sizes = @($dec.Frames | ForEach-Object { $_.PixelWidth } | Sort-Object)
        $sizes | Should -Contain 16
        $sizes | Should -Contain 32
        $sizes | Should -Contain 256
        Add-Type -AssemblyName System.Drawing
        { (New-Object System.Drawing.Icon($ico, 16, 16)).Dispose() } | Should -Not -Throw
    }
    It 'makes a shortcut with the EMT icon and the application ID of the window' {
        $dir = Join-Path $env:TEMP ('EMT-lnk-' + [guid]::NewGuid().ToString('n'))
        try {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'src\New-EmtShortcut.ps1') -Path $dir | Out-Null
            $lnk = Join-Path $dir 'Exchange Message Trail.lnk'
            Test-Path $lnk | Should -BeTrue
            $sc = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk)
            # The shortcut opens EMT.cmd: no PowerShell command line (an antivirus can delete such a shortcut).
            $sc.TargetPath | Should -Be ([IO.Path]::GetFullPath((Join-Path $root 'EMT.cmd')))
            $sc.Arguments | Should -BeNullOrEmpty
            $sc.IconLocation | Should -Match 'gui\\emt\.ico'
            $item = (New-Object -ComObject Shell.Application).NameSpace($dir).ParseName('Exchange Message Trail.lnk')
            $item.ExtendedProperty('System.AppUserModel.ID') | Should -Be 'ExchangeMessageTrail'
        } finally { Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    It 'gives the EMT icon and the application ID to both windows' {
        foreach ($f in 'ExchangeMessageTrail.ps1', 'ExchangeMessageTrail.Classic.ps1') {
            $src = [IO.File]::ReadAllText((Join-Path $root $f), [Text.Encoding]::UTF8)
            $src | Should -Match 'emt\.ico' -Because $f
            $src | Should -Match "SetCurrentProcessExplicitAppUserModelID\('ExchangeMessageTrail'\)" -Because $f
        }
    }
}

Describe 'Trace: the causes of a DROP event (tested on Exchange 2019)' {
    BeforeAll {
        $t = [datetime]'2026-11-10 10:00'
        function New-TEv($Id, $Kind, [string[]]$To, [string[]]$Status, $Time) {
            [pscustomobject]@{ MessageId = $Id; EventId = $Kind; Recipients = $To; RecipientStatus = $Status
                               Timestamp = $Time; ServerHostname = 'EXCH1'; Source = 'ROUTING'; Sender = 'a@x.com'; MessageSubject = 'S' }
        }
        # The server answer as the tracking log of Exchange 2019 writes it.
        function New-Led([string] $Text) { "[{LED=250 2.1.5 $Text};{MSG=};{FQDN=};{IP=};{LRT=}]" }
        function Get-OneStatus([string] $Detail, [string] $Kind = 'Internal') {
            $r = @(Get-EmtDeliveryReport -Events @(New-TEv '<1>' 'DROP' @('r@x.com') @($Detail) $t) -Kinds @{ 'r@x.com' = $Kind })
            if ($r.Count) { $r[0].Status } else { $null }
        }
    }
    AfterEach { $global:EmtLang = 'ru' }

    It 'gives Group for an expanded group, not Dropped' {
        Get-OneStatus (New-Led 'RESOLVER.GRP.Expanded; distribution list expanded') 'Group' | Should -Be 'Group'
        Get-OneStatus (New-Led 'RESOLVER.GRP.NotRedirected; not redirected') 'Group' | Should -Be 'Group'
        Get-OneStatus (New-Led 'RESOLVER.GRP.Expanded; distribution list expanded') 'Other' | Should -Be 'Group'
    }
    It 'gives Forwarded for a mailbox with forwarding' {
        Get-OneStatus (New-Led 'RESOLVER.FWD.Forwarded; recipient forwarded') | Should -Be 'Forwarded'
    }
    It 'gives Suppressed for a copy that Exchange does not send by a setting' {
        foreach ($c in 'RESOLVER.MSGTYPE.AF; Suppressed Auto Forward', 'RESOLVER.OOF.IntToExt; suppressed',
                       'RESOLVER.OOF.ExtToInt; suppressed', 'RESOLVER.MSGTYPE.NDR; suppressed NDR') {
            Get-OneStatus (New-Led $c) | Should -Be 'Suppressed' -Because $c
        }
        Get-OneStatus '[Stage: OnPromotedEvent][Agent: Shared Mailbox Sent Item Agent]' | Should -Be 'Suppressed'
    }
    It 'gives GroupBlocked for a group that does not accept the message, and this is a non-delivery' {
        Get-OneStatus (New-Led 'RESOLVER.GRP.BlockedMessage; blocked') 'Group' | Should -Be 'GroupBlocked'
        'GroupBlocked' | Should -BeIn @(Get-EmtErrorStatuses)
        'Suppressed' | Should -Not -BeIn @(Get-EmtErrorStatuses)
        'Forwarded' | Should -Not -BeIn @(Get-EmtErrorStatuses)
    }
    It 'gives no status for a duplicate recipient: the other events give the status' {
        Get-OneStatus (New-Led 'CAT.DuplicateRecipient; duplicate') | Should -BeNullOrEmpty
        $ev = @(
            New-TEv '<1>' 'DELIVER' @('r@x.com') @('') $t
            New-TEv '<1>' 'DROP'    @('r@x.com') @((New-Led 'CAT.DuplicateRecipient; duplicate')) $t.AddSeconds(1)
        )
        @(Get-EmtDeliveryReport -Events $ev -Kinds @{ 'r@x.com' = 'Internal' })[0].Status | Should -Be 'Delivered'
    }
    It 'keeps Dropped for a DROP with an error code or without a cause' {
        Get-OneStatus '550 5.7.1 dropped by a policy' | Should -Be 'Dropped'
        Get-OneStatus '' | Should -Be 'Dropped'
        @(Get-EmtDeliveryReport -Events @(New-TEv '<1>' 'POISONMESSAGE' @('r@x.com') @('') $t))[0].Status | Should -Be 'Dropped'
    }
    It 'shows the new statuses in the status line in the fixed order' {
        $global:EmtLang = 'en'
        $rows = @('Delivered', 'Suppressed', 'Forwarded', 'GroupBlocked' | ForEach-Object { [pscustomobject]@{ Status = $_ } })
        Get-EmtMessageStatusLine $rows | Should -Be 'delivered: 1, forwarded: 1, rejected by a group: 1, not sent by a setting: 1'
    }
    It 'explains each cause in the column Explanation' {
        $global:EmtLang = 'en'
        Get-EmtStatusHint (New-Led 'RESOLVER.GRP.Expanded; x') | Should -Match '^2\.1\.5: the group was expanded'
        Get-EmtStatusHint (New-Led 'RESOLVER.FWD.Forwarded; x') | Should -Match 'forwarding'
        Get-EmtStatusHint (New-Led 'RESOLVER.MSGTYPE.AF; x') | Should -Match 'automatic forward'
        Get-EmtStatusHint (New-Led 'RESOLVER.OOF.IntToExt; x') | Should -Match 'Out of Office'
        Get-EmtStatusHint (New-Led 'RESOLVER.GRP.BlockedMessage; x') | Should -Match 'restriction'
        Get-EmtStatusHint '[Stage: OnPromotedEvent][Agent: Shared Mailbox Sent Item Agent]' | Should -Match '^Exchange did not send'
    }
    It 'counts only the non-deliveries as errors of a message' {
        $ev = @(
            New-TEv '<1>' 'DELIVER' @('a1@x.com') @('') $t
            New-TEv '<1>' 'DROP'    @('team@x.com') @((New-Led 'RESOLVER.GRP.Expanded; x')) $t
            New-TEv '<1>' 'DROP'    @('fw@x.com') @((New-Led 'RESOLVER.FWD.Forwarded; x')) $t
        )
        $r = @(Get-EmtDeliveryReport -Events $ev -Kinds @{ 'a1@x.com' = 'Internal'; 'team@x.com' = 'Group'; 'fw@x.com' = 'Internal' })
        @($r | Where-Object { $_.Status -in (Get-EmtErrorStatuses) }).Count | Should -Be 0
    }
}

Describe 'Invoke-EmtTrace: messages with one recipient, more recipients and no report rows' {
    BeforeAll {
        $t = [datetime]'2026-11-10 10:00'
        function New-TEv($Id, $Kind, [string[]]$To, $Time) {
            [pscustomobject]@{ MessageId = $Id; EventId = $Kind; Recipients = $To; RecipientStatus = @($To | ForEach-Object { '' })
                               Timestamp = $Time; ServerHostname = 'EXCH1'; Source = 'SMTP'; Sender = 'a@x.com'; MessageSubject = 'S' }
        }
        Mock Get-EmtCapabilities { [pscustomobject]@{ Tracking = $true } }
        Mock Get-EmtAcceptedDomains { @('x.com') }
        Mock Resolve-EmtRecipients { @($Addresses | ForEach-Object { [pscustomobject]@{ Address = $_; Kind = 'Internal'; Mailbox = $_; Type = 'UserMailbox' } }) }
        Mock Find-EmtTrackingEvents {
            [pscustomobject]@{
                Servers = @('EXCH1'); Errors = @()
                Events  = @(
                    New-TEv '<one>' 'RECEIVE' @('b@x.com') $t
                    New-TEv '<one>' 'DELIVER' @('b@x.com') $t.AddSeconds(1)
                    New-TEv '<two>' 'RECEIVE' @('b@x.com', 'c@x.com') $t
                    New-TEv '<two>' 'DELIVER' @('b@x.com', 'c@x.com') $t.AddSeconds(1)
                    # An event without recipients: the message has no report row.
                    New-TEv '<none>' 'SUBMIT' @() $t
                )
            }
        }
    }
    It 'keeps each message, also with one recipient (lesson from Exchange 2019: 61 IDs gave 59 messages)' {
        $r = Invoke-EmtTrace -Criteria @{ Senders = @('a@x.com'); Start = $t.AddDays(-1); End = $t.AddDays(1) }
        @($r.Messages).Count | Should -Be 3
        (@($r.Messages) | Where-Object MessageId -eq '<one>').Recipients | Should -Be 1
        (@($r.Messages) | Where-Object MessageId -eq '<two>').Recipients | Should -Be 2
        (@($r.Messages) | Where-Object MessageId -eq '<none>').Recipients | Should -Be 0
    }
}

Describe 'Get-EmtExtraItemMailboxes: more items than messages in the tracking log' {
    It 'gives the mailboxes with more items than messages' {
        $found = @(
            [pscustomobject]@{ Mailbox = 'a@x.com'; Items = 2 }
            [pscustomobject]@{ Mailbox = 'b@x.com'; Items = 1 }
            [pscustomobject]@{ Mailbox = 'c@x.com'; Items = 4 }
        )
        @(Get-EmtExtraItemMailboxes -Found $found -Messages 1) | Should -Be @('a@x.com', 'c@x.com')
        @(Get-EmtExtraItemMailboxes -Found $found -Messages 4) | Should -BeNullOrEmpty
    }
    It 'gives nothing when the tracking log has no message (the warning "not confirmed" tells it)' {
        @(Get-EmtExtraItemMailboxes -Found @([pscustomobject]@{ Mailbox = 'a@x.com'; Items = 3 }) -Messages 0) | Should -BeNullOrEmpty
    }
}

Describe 'The period of "Delete this message": one day of KQL' {
    It 'gives one day from some minutes before the message to the next midnight (UTC+3)' {
        $sent = [datetime]'2026-11-10 14:03'
        $r = Get-EmtKqlDateRange -Start $sent.AddMinutes(-5) -End $sent.Date.AddDays(1) -UtcOffset ([timespan]'03:00')
        $r.From | Should -Be '2026-11-10'
        $r.To   | Should -Be '2026-11-11'
    }
    It 'gives two days for a message before 03:00 local time (UTC+3): the UTC day is the day before' {
        $sent = [datetime]'2026-11-10 01:30'
        $r = Get-EmtKqlDateRange -Start $sent.AddMinutes(-5) -End $sent.Date.AddDays(1) -UtcOffset ([timespan]'03:00')
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-11'
    }
}

Describe 'Trace: the folder of a delivery (an inbox rule)' {
    BeforeAll {
        $t = [datetime]'2026-11-10 10:00'
        function Get-Row($EventId, $Detail) {
            $e = [pscustomobject]@{ MessageId = '<1>'; EventId = $EventId; Recipients = @('r@x.com'); RecipientStatus = @($Detail)
                                    Timestamp = $t; ServerHostname = 'EXCH1'; Source = 'STOREDRIVER'; Sender = 'a@x.com'; MessageSubject = 'S' }
            @(Get-EmtDeliveryReport -Events @($e) -Kinds @{ 'r@x.com' = 'Internal' })[0]
        }
    }
    It 'gives the folder of a DELIVER event, and the server answer is empty' {
        $r = Get-Row 'DELIVER' 'Projects'
        $r.Folder | Should -Be 'Projects'
        $r.Detail | Should -Be ''
        (Get-Row 'DELIVER' 'ОМО').Folder | Should -Be 'ОМО'
    }
    It 'gives no folder for Inbox (empty status), for an SMTP answer and for other events' {
        (Get-Row 'DELIVER' '').Folder | Should -Be ''
        $r = Get-Row 'DELIVER' '250 2.0.0 OK'
        $r.Folder | Should -Be ''
        $r.Detail | Should -Be '250 2.0.0 OK'
        (Get-Row 'SEND' 'Projects').Folder | Should -Be ''
        Get-EmtDeliveryFolder -EventId 'DELIVER' -Detail '[{LED=250 2.1.5 RESOLVER.GRP.Expanded};{MSG=}]' | Should -Be ''
    }
}

Describe 'Get-EmtQuickPeriod: the button of the current period' {
    BeforeAll { $now = [datetime]'2026-11-11 15:20' }
    It 'finds each quick period' {
        Get-EmtQuickPeriod -Start '2026-11-11 00:00' -End '2026-11-11 23:59' -Now $now | Should -Be 'Today'
        Get-EmtQuickPeriod -Start '2026-11-10 00:00' -End '2026-11-10 23:59' -Now $now | Should -Be 'Yesterday'
        Get-EmtQuickPeriod -Start '2026-11-05 00:00' -End '2026-11-11 23:59' -Now $now | Should -Be 'Week'
        Get-EmtQuickPeriod -Start '2026-10-29 00:00' -End '2026-11-11 23:59' -Now $now | Should -Be 'Days14'
    }
    It 'ignores the seconds: the fields have minutes only' {
        Get-EmtQuickPeriod -Start '2026-11-11 00:00:00' -End '2026-11-11 23:59:59.999' -Now $now | Should -Be 'Today'
    }
    It 'gives no button for a manual period' {
        Get-EmtQuickPeriod -Start '2026-11-11 08:00' -End '2026-11-11 23:59' -Now $now | Should -Be ''
        Get-EmtQuickPeriod -Start '2026-11-11 00:00' -End '2026-11-11 18:00' -Now $now | Should -Be ''
        Get-EmtQuickPeriod -Start '2026-11-04 00:00' -End '2026-11-11 23:59' -Now $now | Should -Be ''
    }
    It 'gives no button after midnight: "today" is a new day' {
        Get-EmtQuickPeriod -Start '2026-11-11 00:00' -End '2026-11-11 23:59' -Now ([datetime]'2026-11-12 00:01') | Should -Be 'Yesterday'
        Get-EmtQuickPeriod -Start '2026-11-05 00:00' -End '2026-11-11 23:59' -Now ([datetime]'2026-11-12 00:01') | Should -Be ''
    }
    It 'has a button with this key for each quick period, in the two windows' {
        $keys = @((Get-EmtQuickPeriodDays).Keys)
        $keys | Should -Be @('Today', 'Yesterday', 'Week', 'Days14')
    }
}

Describe 'Get-EmtPeriodEnd: the end of the period includes its minute' {
    It 'gives the last moment of the minute' {
        (Get-EmtPeriodEnd ([datetime]'2026-11-10 23:59')) | Should -Be ([datetime]'2026-11-10 23:59:59.9999999')
        (Get-EmtPeriodEnd ([datetime]'2026-11-10 14:30:20')) | Should -Be ([datetime]'2026-11-10 14:30:59.9999999')
    }
    It 'includes a message in the last minute of the day' {
        $end = Get-EmtPeriodEnd ([datetime]'2026-11-10 23:59')
        ([datetime]'2026-11-10 23:59:30') -le $end | Should -BeTrue
        ([datetime]'2026-11-11 00:00:00') -le $end | Should -BeFalse
    }
    It 'gives the same KQL days as the end at the next midnight' {
        $a = Get-EmtKqlDateRange -Start ([datetime]'2026-11-10') -End (Get-EmtPeriodEnd ([datetime]'2026-11-10 23:59')) -UtcOffset 0
        $b = Get-EmtKqlDateRange -Start ([datetime]'2026-11-10') -End ([datetime]'2026-11-11') -UtcOffset 0
        $a.From | Should -Be $b.From
        $a.To   | Should -Be $b.To
    }
}

Describe 'Message search by a recipient: partial search and the details of a message' {
    BeforeAll {
        $t = [datetime]'2026-11-10 10:00'
        function New-TEv($Id, $Kind, [string[]]$To, $Time) {
            [pscustomobject]@{ MessageId = $Id; EventId = $Kind; Recipients = $To; RecipientStatus = @($To | ForEach-Object { '' })
                               Timestamp = $Time; ServerHostname = 'EXCH1'; Source = 'SMTP'; Sender = 'a@x.com'; MessageSubject = 'S' }
        }
        Mock Get-EmtCapabilities { [pscustomobject]@{ Tracking = $true } }
        Mock Get-EmtAcceptedDomains { @('x.com') }
        Mock Resolve-EmtRecipients { @($Addresses | ForEach-Object { [pscustomobject]@{ Address = $_; Kind = 'Internal'; Mailbox = $_; Type = 'UserMailbox' } }) }
        # The first pass with a recipient filter: the events of the recipient b@x.com only.
        # RECEIVE has all recipients; the delivery to c@x.com is on another server and is not in the result.
        Mock Find-EmtTrackingEvents {
            [pscustomobject]@{ Servers = @('EXCH1', 'EXCH2'); Errors = @()
                Events = @(
                    New-TEv '<m1>' 'RECEIVE' @('b@x.com', 'c@x.com') $t
                    New-TEv '<m1>' 'DELIVER' @('b@x.com') $t.AddSeconds(1)
                ) }
        }
        $crit = @{ Recipient = 'b@x.com'; Start = $t.AddDays(-1); End = $t.AddDays(1) }
    }
    It 'makes no second pass and keeps only the rows of the searched recipient' {
        $r = Invoke-EmtTrace -Criteria $crit -Partial
        $r.Partial | Should -BeTrue
        Should -Invoke Find-EmtTrackingEvents -Times 1 -Exactly -ParameterFilter { -not $CompleteRecipients }
        @($r.Report).Count | Should -Be 1
        @($r.Report)[0].Recipient | Should -Be 'b@x.com'
        @($r.Report)[0].Status | Should -Be 'Delivered'
    }
    It 'gives the group size from the EXPAND event in the message rows' {
        $exp = New-TEv '<m1>' 'EXPAND' @('b@x.com') $t.AddSeconds(1)
        $exp | Add-Member RecipientCount 670
        Mock Find-EmtTrackingEvents {
            [pscustomobject]@{ Servers = @('EXCH1'); Errors = @()
                Events = @((New-TEv '<m1>' 'RECEIVE' @('all@x.com') $t), $exp, (New-TEv '<m1>' 'DELIVER' @('b@x.com') $t.AddSeconds(2))) }
        }
        $r = Invoke-EmtTrace -Criteria $crit -Partial
        @($r.Messages)[0].Expanded | Should -Be 670
        @($r.Messages)[0].ToGroup  | Should -BeFalse
    }
    It 'marks a message of a search by a group address' {
        Mock Find-EmtTrackingEvents {
            [pscustomobject]@{ Servers = @('EXCH1'); Errors = @()
                Events = @(New-TEv '<m1>' 'RECEIVE' @('all@x.com') $t) }
        }
        Mock Resolve-EmtRecipients { @($Addresses | ForEach-Object { [pscustomobject]@{ Address = $_; Kind = 'Group'; Mailbox = ''; Type = 'MailUniversalDistributionGroup' } }) }
        $r = Invoke-EmtTrace -Criteria @{ Recipient = 'All@x.com'; Start = $t.AddDays(-1); End = $t.AddDays(1) } -Partial
        @($r.Messages)[0].ToGroup | Should -BeTrue
        (Invoke-EmtTrace -Criteria @{ Recipient = 'All@x.com'; Start = $t.AddDays(-1); End = $t.AddDays(1) }).Messages[0].ToGroup | Should -BeFalse
    }
    It 'makes the second pass without -Partial (the classic window, the search of copies)' {
        $r = Invoke-EmtTrace -Criteria $crit
        $r.Partial | Should -BeFalse
        Should -Invoke Find-EmtTrackingEvents -Times 1 -Exactly -ParameterFilter { $CompleteRecipients }
    }
    It 'is not partial without a recipient: such a search has no second pass' {
        (Invoke-EmtTrace -Criteria @{ Senders = @('a@x.com'); Start = $t.AddDays(-1); End = $t.AddDays(1) } -Partial).Partial | Should -BeFalse
    }
    It 'loads all recipients of a message: one call for each server' {
        Mock Get-TransportService { @([pscustomobject]@{ Name = 'EXCH1' }, [pscustomobject]@{ Name = 'EXCH2' }) }
        Mock Get-MessageTrackingLog {
            if ($Server -eq 'EXCH1') { @(New-TEv '<m1>' 'RECEIVE' @('b@x.com', 'c@x.com') $t; New-TEv '<m1>' 'DELIVER' @('b@x.com') $t.AddSeconds(1)) }
            else { @(New-TEv '<m1>' 'DELIVER' @('c@x.com') $t.AddSeconds(2)) }
        }
        $d = Get-EmtMessageDetail -MessageIds @('<m1>') -Start $t.AddDays(-1) -End $t.AddDays(1)
        Should -Invoke Get-MessageTrackingLog -Times 2 -Exactly
        @($d.MessageIds) | Should -Be @('<m1>')
        @($d.Report).Count | Should -Be 2
        @($d.Report | Where-Object Status -eq 'Delivered').Count | Should -Be 2
    }
    It 'collects the errors of the servers and continues' {
        Mock Get-MessageTrackingLog { Write-Error 'RPC server is unavailable' }
        $p = Get-EmtEventsByMessageId -MessageIds @('<m1>', '<m2>') -Start $t -End $t.AddDays(1) -Servers @('EXCH1')
        @($p.Errors).Count | Should -Be 2
        @($p.Events).Count | Should -Be 0
    }
}

Describe 'Shadow copy (Shadow Redundancy, tested on Exchange 2019)' {
    # A script sends a message without Message-ID. Server 1 gives one MessageId and delivers the message.
    # Server 2 keeps a shadow copy with its own MessageId: only HARECEIVE and HADISCARD.
    BeforeAll {
        $t = [datetime]'2026-11-13 10:03:19'
        function New-SEv($Id, $Kind, [string[]]$To, $Time, $Server) {
            [pscustomobject]@{ MessageId = $Id; EventId = $Kind; Recipients = $To; RecipientStatus = @($To | ForEach-Object { '' })
                               Timestamp = $Time; ServerHostname = $Server; Source = 'SMTP'; Sender = 'script@x.com'; MessageSubject = 'Sync' }
        }
        $script:ShadowEvents = @(
            New-SEv '<s1@EXCH02.x.local>' 'HARECEIVE' @('b@x.com', 'c@x.com') $t 'EXCH02'
            New-SEv '<m1@EXCH01.x.local>' 'HAREDIRECT' @('b@x.com', 'c@x.com') $t 'EXCH01'
            New-SEv '<m1@EXCH01.x.local>' 'RECEIVE' @('b@x.com', 'c@x.com') $t 'EXCH01'
            New-SEv '<m1@EXCH01.x.local>' 'DELIVER' @('b@x.com', 'c@x.com') $t.AddSeconds(1) 'EXCH01'
            New-SEv '<s1@EXCH02.x.local>' 'HADISCARD' @('b@x.com', 'c@x.com') $t.AddSeconds(24) 'EXCH02'
        )
    }
    It 'Get-EmtMessageSummary: the copy with only HARECEIVE and HADISCARD is a shadow copy, not a message from a rule' {
        $s = @(Get-EmtMessageSummary -Events $script:ShadowEvents -ServerNames @('EXCH01', 'EXCH02'))
        $copy = $s | Where-Object MessageId -eq '<s1@EXCH02.x.local>'
        $copy.Shadow | Should -BeTrue
        $copy.ServerGenerated | Should -BeFalse
        $copy.Note | Should -Be (T 'note.shadow')
        # HAREDIRECT on the first server does not make the original message a shadow copy.
        ($s | Where-Object MessageId -eq '<m1@EXCH01.x.local>').Shadow | Should -BeFalse
    }
    It 'Get-EmtMessageStatusLine -Shadow: "shadow copy", not "no recipients"' {
        Get-EmtMessageStatusLine @() -Shadow | Should -Be 'теневая копия'
        Get-EmtMessageStatusLine @() | Should -Be (T 'line.none')
    }
    It 'Get-EmtMessageCountText: a shadow copy is not counted as a message' {
        $m = { param($shadow) [pscustomobject]@{ Shadow = $shadow } }
        Get-EmtMessageCountText @((& $m $false), (& $m $true), (& $m $false)) | Should -Be '2 (+ теневых копий: 1)'
        Get-EmtMessageCountText @((& $m $false), (& $m $false)) | Should -Be '2'
        Get-EmtMessageCountText @([pscustomobject]@{ MessageId = 'x' }) | Should -Be '1'
        Get-EmtMessageCountText @() | Should -Be '0'
    }
    It 'Invoke-EmtTrace: the shadow copy has the status "shadow copy", and the log counts one message' {
        Mock Get-EmtCapabilities { [pscustomobject]@{ Tracking = $true } }
        Mock Get-EmtAcceptedDomains { @('x.com') }
        Mock Resolve-EmtRecipients { @($Addresses | ForEach-Object { [pscustomobject]@{ Address = $_; Kind = 'Internal'; Mailbox = $_; Type = 'UserMailbox' } }) }
        Mock Find-EmtTrackingEvents { [pscustomobject]@{ Servers = @('EXCH01', 'EXCH02'); Errors = @(); Events = $script:ShadowEvents } }
        Mock Write-EmtLog { }
        $r = Invoke-EmtTrace -Criteria @{ Senders = @('script@x.com'); Start = $t.AddDays(-1); End = $t.AddDays(1) }
        $copy = @($r.Messages) | Where-Object MessageId -eq '<s1@EXCH02.x.local>'
        $copy.Shadow | Should -BeTrue
        $copy.StatusLine | Should -Be 'теневая копия'
        (@($r.Messages) | Where-Object MessageId -eq '<m1@EXCH01.x.local>').StatusLine | Should -Not -Be 'теневая копия'
        Should -Invoke Write-EmtLog -ParameterFilter { $Message -eq (T 'log.traceDone' 1 2 0 0) }
    }
}

Describe 'Invoke-EmtOperation' {
    It 'Trace: the classic window does not send Partial, and the search runs (not partial)' {
        # The core runs in strict mode. A missing key of the parameters must not stop the operation.
        Mock Invoke-EmtTrace { [bool]$Partial }
        Set-StrictMode -Version 2.0
        Invoke-EmtOperation -Name 'Trace' -Params @{ Criteria = @{} } | Should -BeFalse
        Invoke-EmtOperation -Name 'Trace' -Params @{ Criteria = @{}; Partial = $true } | Should -BeTrue
    }
}

Describe 'Parameters of the operations: the windows and the core' {
    # The core runs in strict mode: $Params.Key with a missing key stops the operation.
    # The test reads the code (AST), it does not run the windows.
    BeforeAll {
        $root = Join-Path $PSScriptRoot '..'
        function Get-EmtAst([string] $File) {
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root $File), [ref]$tokens, [ref]$errors)
            if ($errors.Count) { throw "$File has parse errors: $($errors[0].Message)" }
            $ast
        }

        # The core: for each operation, the keys that it reads as $Params.Key (required keys).
        # A key read as $Params['Key'] is optional: it gives $null when it is missing.
        $op = (Get-EmtAst 'src\EMT.Core.ps1').Find({ param($a)
            $a -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $a.Name -eq 'Invoke-EmtOperation' }, $true)
        $switch = $op.Find({ param($a) $a -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
        $script:Required = @{}
        foreach ($clause in $switch.Clauses) {
            $keys = @($clause.Item2.FindAll({ param($a)
                $a -is [System.Management.Automation.Language.MemberExpressionAst] -and
                $a.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $a.Expression.VariablePath.UserPath -eq 'Params' -and
                $a.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true) |
                ForEach-Object { $_.Member.Value } | Sort-Object -Unique)
            $script:Required[[string]$clause.Item1.Value] = $keys
        }

        # A window: each call "Start-EmtOperation -Name '...' -Params @{...}" and its keys.
        function Get-EmtOperationCalls([string] $File) {
            $calls = (Get-EmtAst $File).FindAll({ param($a)
                $a -is [System.Management.Automation.Language.CommandAst] -and $a.GetCommandName() -eq 'Start-EmtOperation' }, $true)
            foreach ($c in $calls) {
                $name = $null; $keys = @()
                $el = $c.CommandElements
                for ($i = 1; $i -lt $el.Count - 1; $i++) {
                    if ($el[$i] -isnot [System.Management.Automation.Language.CommandParameterAst]) { continue }
                    switch ($el[$i].ParameterName) {
                        'Name'   { $name = [string]$el[$i + 1].Value }
                        'Params' {
                            $h = $el[$i + 1]
                            if ($h -isnot [System.Management.Automation.Language.HashtableAst]) { throw "${File}:$($c.Extent.StartLineNumber): -Params is not @{...}." }
                            $keys = @($h.KeyValuePairs | ForEach-Object { [string]$_.Item1.Value })
                        }
                    }
                }
                [pscustomobject]@{ Line = $c.Extent.StartLineNumber; Name = $name; Keys = $keys }
            }
        }
    }

    It 'the core has the switch of the operations' {
        $script:Required.Count | Should -BeGreaterOrEqual 10
        $script:Required['Trace'] | Should -Be @('Criteria')
    }

    It '<File>: each operation exists in the core and gets all required keys' -ForEach @(
        @{ File = 'ExchangeMessageTrail.ps1' }, @{ File = 'ExchangeMessageTrail.Classic.ps1' }
    ) {
        $calls = @(Get-EmtOperationCalls $File)
        # The parse must find the calls. Without them, the test passes without a check.
        $calls.Count | Should -BeGreaterOrEqual 10
        $problems = @(foreach ($c in $calls) {
            if (-not $c.Name -or -not $script:Required.ContainsKey($c.Name)) { "line $($c.Line): the core has no operation '$($c.Name)'"; continue }
            foreach ($k in $script:Required[$c.Name]) {
                if ($k -notin $c.Keys) { "line $($c.Line): operation '$($c.Name)' does not get the key '$k'" }
            }
        })
        $problems | Should -BeNullOrEmpty -Because ($problems -join '; ')
    }
}

Describe 'User guide (help\*.html)' {
    BeforeAll {
        $root = Join-Path $PSScriptRoot '..'
        $help = Join-Path $root 'help'
        function Read-Page([string] $Name) { [IO.File]::ReadAllText((Join-Path $help $Name), [Text.Encoding]::UTF8) }
        # The last version in CHANGELOG.md: the first line "## x.y.z".
        $changelog = [IO.File]::ReadAllText((Join-Path $root 'CHANGELOG.md'), [Text.Encoding]::UTF8)
        $version = [regex]::Match($changelog, '(?m)^## (\d+\.\d+\.\d+)').Groups[1].Value
    }
    It 'has the version of the last CHANGELOG entry on each page (update the guide with each version)' {
        $version | Should -Not -BeNullOrEmpty
        foreach ($p in 'index.html', 'en.html', 'ru.html') {
            [regex]::Match((Read-Page $p), '<meta name="emt-version" content="([^"]+)"').Groups[1].Value | Should -Be $version -Because $p
        }
    }
    It 'has each image that a page shows' {
        foreach ($p in 'en.html', 'ru.html') {
            foreach ($m in [regex]::Matches((Read-Page $p), '<img src="([^"]+)"')) {
                Test-Path (Join-Path $help $m.Groups[1].Value) | Should -BeTrue -Because "$p shows $($m.Groups[1].Value)"
            }
        }
    }
    It 'has a section for each link of the table of contents' {
        foreach ($p in 'en.html', 'ru.html') {
            $page = Read-Page $p
            $ids = @([regex]::Matches($page, '\sid="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
            foreach ($m in [regex]::Matches($page, 'href="#([^"]+)"')) {
                $m.Groups[1].Value | Should -BeIn $ids -Because "$p links to #$($m.Groups[1].Value)"
            }
        }
    }
    It 'has the same sections in English and in Russian' {
        $en = @([regex]::Matches((Read-Page 'en.html'), '<(?:section|h3) id="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
        $ru = @([regex]::Matches((Read-Page 'ru.html'), '<(?:section|h3) id="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
        $en | Should -Be $ru
    }
    It 'uses no external files: the guide operates without the Internet' {
        # A link (<a href>) to a web page is allowed: the user clicks it. Images, scripts and styles must be local.
        foreach ($p in 'index.html', 'en.html', 'ru.html') {
            $page = Read-Page $p
            $page | Should -Not -Match 'src="(https?:)?//' -Because $p
            $page | Should -Not -Match '<link[^>]+href="(https?:)?//' -Because $p
            $page | Should -Not -Match 'url\((https?:)?//' -Because $p
        }
        (Read-Page 'help.css') | Should -Not -Match '(url\(|@import)\s*["'']?(https?:)?//'
    }
    It 'opens each external link in a new tab, without access to the guide' {
        foreach ($p in 'en.html', 'ru.html') {
            foreach ($m in [regex]::Matches((Read-Page $p), '<a [^>]*href="https?://[^"]*"[^>]*>')) {
                $m.Value | Should -Match 'target="_blank"' -Because $m.Value
                $m.Value | Should -Match 'rel="noopener"' -Because $m.Value
            }
        }
    }
}

Describe 'Tables and defaults of the windows (2.2.2)' {
    BeforeAll {
        $root = Split-Path -Parent $PSScriptRoot
        $main = [IO.File]::ReadAllText((Join-Path $root 'ExchangeMessageTrail.ps1'), [Text.Encoding]::UTF8)
        $classic = [IO.File]::ReadAllText((Join-Path $root 'ExchangeMessageTrail.Classic.ps1'), [Text.Encoding]::UTF8)
        $core = [IO.File]::ReadAllText((Join-Path $root 'src\EMT.Core.ps1'), [Text.Encoding]::UTF8)
    }
    It 'a double-click on the table "Messages" opens the events only from a row, not from the header' {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($main, [ref]$null, [ref]$null)
        $calls = @($ast.FindAll({ param($n)
            $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
            $n.Member.Value -eq 'Add_MouseDoubleClick' -and $n.Expression.Extent.Text -eq '$gridTrMessages' }, $true))
        $calls.Count | Should -Be 1
        $calls[0].Arguments[0].Extent.Text | Should -Match 'DataGridRow'
    }
    It 'the main window has no star column: a drag changes only one column' {
        $main | Should -Not -Match 'DataGridLengthUnitType\]::Star'
    }
    It 'the default timeout is 30 minutes in both windows and in the core' {
        $main | Should -Match '\$script:LastTimeout = 30\b'
        $main | Should -Match '\{ \$n \} else \{ 30 \}'
        $classic | Should -Match '\$nudTimeout\.Value = 30;'
        [regex]::Matches($core, '\[int\] \$TimeoutMinutes = (\d+)') | ForEach-Object { $_.Groups[1].Value | Should -Be '30' }
    }
}

Describe 'The window "About" (2.3.0)' {
    BeforeAll {
        $root = Split-Path -Parent $PSScriptRoot
        $readme = [IO.File]::ReadAllText((Join-Path $root 'README.md'), [Text.Encoding]::UTF8)
        $changelog = [IO.File]::ReadAllText((Join-Path $root 'CHANGELOG.md'), [Text.Encoding]::UTF8)
    }
    It 'Get-EmtVersion is the first version in CHANGELOG.md' {
        Get-EmtVersion | Should -Be ([regex]::Match($changelog, '(?m)^## (\d+\.\d+\.\d+)').Groups[1].Value)
    }
    It 'the links are https pages of GitHub and Telegram' {
        $l = Get-EmtLinks
        $l.GitHub | Should -Be 'https://github.com/sobakoed/exchange-message-trail'
        $l.Releases | Should -Be 'https://github.com/sobakoed/exchange-message-trail/releases/latest'
        $l.Telegram | Should -Match '^https://t\.me/[A-Za-z0-9_]+$'
    }
    It 'the TON address is the address of README.md and of the user guide' {
        $ton = (Get-EmtLinks).Ton
        $ton | Should -Match '^[A-Za-z0-9_-]{48}$'
        $readme | Should -Match ([regex]::Escape($ton))
        foreach ($p in 'ru.html', 'en.html') {
            [IO.File]::ReadAllText((Join-Path $root "help\$p"), [Text.Encoding]::UTF8) | Should -Match ([regex]::Escape($ton)) -Because $p
        }
    }
    It 'the QR code is in the folder help' {
        Join-Path $root 'help\img\donate-ton.png' | Should -Exist
    }
    It 'both windows have the menu of the button "?" and the key F1' {
        foreach ($n in 'ExchangeMessageTrail.ps1', 'ExchangeMessageTrail.Classic.ps1') {
            $src = [IO.File]::ReadAllText((Join-Path $root $n), [Text.Encoding]::UTF8)
            $src | Should -Match '\$btnHelp\.Add_Click\(\{ Show-EmtHelpMenu \}\)' -Because $n
            $src | Should -Match 'function Show-EmtAbout' -Because $n
            $src | Should -Match "T 'menu\.about'" -Because $n
            $src | Should -Match 'F1' -Because $n
        }
    }
}
