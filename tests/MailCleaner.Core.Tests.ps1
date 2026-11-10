#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for src\MailCleaner.Core.ps1. Exchange is not necessary:
    Exchange cmdlets are replaced with stub functions and Pester mocks.

    Run:  Invoke-Pester -Path .\tests
#>

BeforeAll {
    . (Join-Path $PSScriptRoot '..\src\MailCleaner.Core.ps1')

    # Stubs: Pester can mock only commands that exist.
    function Get-TransportService { [CmdletBinding()] param() }
    function Get-MessageTrackingLog {
        [CmdletBinding()]
        param($Server, $Sender, $MessageSubject, $MessageId, $Recipients, $Start, $End, $ResultSize)
    }
    function Get-Recipient { [CmdletBinding()] param($Identity) }

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
        $r = Get-McKqlDateRange -Start ([datetime]'2026-11-09') -End ([datetime]'2026-11-10')
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-10'
    }
    It 'includes the End day when End is not at midnight' {
        $r = Get-McKqlDateRange -Start ([datetime]'2026-11-09 15:00') -End ([datetime]'2026-11-09 16:00')
        $r.From | Should -Be '2026-11-09'
        $r.To   | Should -Be '2026-11-10'
    }
    It 'refuses End before Start' {
        { Get-McKqlDateRange -Start ([datetime]'2026-11-10') -End ([datetime]'2026-11-09') } | Should -Throw
    }
}

Describe 'New-McSearchQuery' {
    BeforeAll {
        $s = [datetime]'2026-11-09'
        $e = [datetime]'2026-11-10'
    }
    It 'joins senders with OR and adds subject and dates' {
        New-McSearchQuery -Senders 'b@x.com', 'A@x.com' -Subject 'Отчёт за квартал' -Start $s -End $e |
            Should -Be '(from:a@x.com OR from:b@x.com) AND (subject:"Отчёт за квартал") AND (sent>=2026-11-09) AND (sent<2026-11-10)'
    }
    It 'removes quotes from phrases' {
        New-McSearchQuery -Subject 'a "b" c' -Start $s -End $e | Should -Match 'subject:"a b c"'
    }
    It 'adds attachment and keywords' {
        $q = New-McSearchQuery -Attachment 'report.xlsx' -Keywords 'salary' -Start $s -End $e
        $q | Should -Match '\(attachment:"report\.xlsx"\)'
        $q | Should -Match '\("salary"\)'
    }
    It 'refuses a query that has only dates' {
        { New-McSearchQuery -Start $s -End $e } | Should -Throw
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
    It 'classifies mailbox, contact, group and unknown addresses' {
        Mock Get-Recipient {
            switch ($Identity) {
                'u@x.com' { [pscustomobject]@{ RecipientType = 'UserMailbox'; PrimarySmtpAddress = 'U@x.com' } }
                'c@y.com' { [pscustomobject]@{ RecipientType = 'MailContact'; PrimarySmtpAddress = 'c@y.com' } }
                'g@x.com' { [pscustomobject]@{ RecipientType = 'MailUniversalDistributionGroup'; PrimarySmtpAddress = 'g@x.com' } }
                default   { throw 'not found' }
            }
        }
        $r = @(Resolve-McRecipients -Addresses 'u@x.com', 'c@y.com', 'g@x.com', 'z@ext.org')
        ($r | Where-Object Address -eq 'u@x.com').Kind    | Should -Be 'Internal'
        ($r | Where-Object Address -eq 'u@x.com').Mailbox | Should -Be 'u@x.com'
        ($r | Where-Object Address -eq 'c@y.com').Kind    | Should -Be 'External'
        ($r | Where-Object Address -eq 'g@x.com').Kind    | Should -Be 'Group'
        ($r | Where-Object Address -eq 'z@ext.org').Kind  | Should -Be 'External'
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
