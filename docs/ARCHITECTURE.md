# Architecture

This document gives the structure of the Exchange Message Trail (EMT) code.

## Files

| File | Content |
|---|---|
| `ExchangeMessageTrail.ps1` | The window, the background runspace and the event handlers |
| `src/EMT.Core.ps1` | The logic without the user interface |
| `src/EMT.Strings.ps1` | The language loader, the function `T` and the language file check |
| `Lang/*.lng` | The user interface texts, one file for each language |
| `EMT.cmd` | The launcher (STA mode, hidden console) |
| `tests/EMT.Core.Tests.ps1` | The Pester 5 unit tests for the core |

All `.ps1` files use UTF-8 with BOM and CRLF.
Windows PowerShell 5.1 reads a file without BOM in the ANSI code page. This changes the Cyrillic text.

## Threads

```
UI thread (STA)                                 Background runspace
──────────────────────────────                  ────────────────────────────────────
Form, controls, event handlers                  Remote session to Exchange
        │ Start-EmtOperation ── BeginInvoke ──▶ . EMT.Core.ps1
        │                                       Invoke-EmtOperation -Name <operation>
Timer (400 ms) ◀── $Sync.Log (queue) ────────── Write-EmtLog
        │      ◀── $Sync.Progress / Status ──── Set-EmtProgress
        │      ─── $Sync.Cancel ──────────────▶ Test-EmtCancel
        │      ─── $Sync.Lang ────────────────▶ T 'key'
        └ Complete-EmtOperation ◀── EndInvoke ─ result object
```

- The window never calls Exchange. The background runspace does all Exchange work.
- The runspace keeps one remote session (`$global:EmtExSession`) for all operations.
- `$Sync` (`$global:EmtSync`) is a synchronized hashtable. The UI thread and the runspace share it.
- The timer reads the log queue and the progress. It calls `Complete-EmtOperation` at the end of an operation.
- The timer stops while a result handler shows a dialog. This prevents a second call.

## Connection

`Connect-EmtExchange` does these steps:

1. It makes a session with `New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Kerberos`.
2. It imports only the necessary cmdlets with `Import-PSSession -CommandName` (`Get-EmtExchangeCommandNames`).
3. It imports the module into the global scope of the runspace.
4. It reads the accepted domains.
5. It returns the capabilities (`Get-EmtCapabilities`): Tracking, MailboxCheck, Search, Verify, SoftDelete, HardDelete.

The Exchange session is a NoLanguage session. It does not run script blocks.
For this reason, the logic runs in the runspace. The runspace calls the imported cmdlets.
The runspace uses the execution policy `Bypass`: `Import-PSSession` makes a temporary script module.

## Operations

`Invoke-EmtOperation` is the only entry point for the window.

| Operation | Tab | Function | Result |
|---|---|---|---|
| `Connect` | — | `Connect-EmtExchange` | Capabilities |
| `Resolve` | both | `Resolve-EmtAddressInputs` (`Get-Recipient -Anr`) | Candidates for each partial address |
| `Trace` | Message search | `Invoke-EmtTrace` | Messages, delivery report, events, hints |
| `CheckMailbox` | Message search | `Invoke-EmtMailboxCheck` | Findings and the data of the checks |
| `Search` | Delete | `Invoke-EmtSearch` | Search result |
| `SoftDelete` | Delete | `Invoke-EmtSoftDelete`, then `Get-EmtSettledCopyStatus` | Delete status and verification |
| `HardDelete` | Delete | `Test-EmtHardDeleteScope`, `Invoke-EmtHardDelete`, `Get-EmtSettledCopyStatus` | Delete counts and verification |
| `Verify` | Delete | `Get-EmtCopyStatus` | Verification |
| `ListOld` | Delete | `Get-EmtOldSearches` | Old `EMT_*` and `MC_*` searches |
| `RemoveOld` | Delete | `Remove-EmtSearches` | None |

## Trace (tab Message search)

`Invoke-EmtTrace` reads only the tracking log. It needs only the Message Tracking role.

1. `Test-EmtCriteria -ForTrace` examines the criteria. A recipient alone is a usable criterion.
2. `Find-EmtTrackingEvents` reads the tracking log of all transport servers.
   With a recipient filter, a second pass by message ID gets all recipients.
3. `Get-EmtMessageSummary` makes one row for each message ID. Events without a message ID are not messages.
4. `Resolve-EmtRecipients` classifies each address: internal, external, group or other.
5. `Get-EmtDeliveryReport` makes one row for each message and recipient.
6. When there is no message, `Get-EmtNoResultHints` gives the hints.

### Delivery report

`Get-EmtEventStatus` gives the status of one event for one recipient. `Get-EmtStatusRank` orders the statuses.
For each message and recipient, the status with the highest rank wins. With the same rank, the later event wins.

| Rank | Status | Events |
|---|---|---|
| 7 | Delivered | `DELIVER` |
| 6 | External | `SEND` to an external address |
| 5 | Failed | `FAIL` |
| 4 | Dropped | `DROP`, `POISONMESSAGE` |
| 3 | Deferred | `DEFER` |
| 2 | InTransit | `SEND` to an internal address, `EXPAND`, `REDIRECT`, `HAREDIRECT`, `TRANSFER`, `RESOLVE` |
| 1 | Group | A group address without its own delivery |
| 0 | Unknown | `RECEIVE`, `SUBMIT` |

`SEND` to an internal recipient is a transfer to another Exchange server, not the end of the delivery.
`Get-EmtStatusHint` reads the enhanced status code (`x.y.z`) in the server answer and gives the explanation.

## Mailbox check

`Invoke-EmtMailboxCheck` runs four read-only cmdlets with `Invoke-EmtReadOnly`:
`Get-InboxRule`, `Get-MailboxJunkEmailConfiguration`, `Get-Mailbox`, `Get-MailboxStatistics`.
`Invoke-EmtReadOnly` never uses `-ErrorAction Stop`. It returns the value and the error text.
A missing cmdlet gives the error "no permission".

`Get-EmtMailboxFindings` makes the findings (Severity, Area, Text) from the data.
`Get-EmtRuleMatch` decides if a rule applies to the message: All (no conditions), Yes, No, Unknown (body words).
The result also contains the data. Thus the window makes the findings again after a language change.

## Delete

`Invoke-EmtSearch` does these steps:

1. `Complete-EmtCriteria` adds the missing keys. `Test-EmtCriteria` examines the criteria.
2. With **Recipient mailbox only**, `Resolve-EmtRecipients` finds the recipient mailbox.
3. `Find-EmtTrackingEvents` reads the tracking log.
4. `Resolve-EmtRecipients` classifies each address.
5. `New-EmtSearchQuery` makes the KQL query. The recipient goes into the query as `to:` / `cc:` / `bcc:`.
6. EMT selects the search locations and the method (see the table below).
7. `Get-EmtMissingMailboxes` finds the mailboxes without the message (only when other mailboxes have it).
8. The result is **Unconfirmed** when the search found items, but the tracking log has no message.

| Condition | Locations |
|---|---|
| Recipient mailbox only | The recipient mailbox |
| Search all mailboxes | `All` |
| The tracking log found internal recipients | Internal recipient and sender mailboxes |
| No internal recipients, a recipient is set | The recipient and sender mailboxes |
| No internal recipients, no recipient | `All` |

| Locations | Method | Function |
|---|---|---|
| 1 to `Get-EmtQuickSearchLimit` (10) | Quick | `Invoke-EmtQuickSearch` |
| More, or `All` | Compliance Search | `New-EmtComplianceSearchRun`, `Wait-EmtComplianceSearch`, `Get-EmtSearchHits` |

A quick search does not make a Compliance Search. The `SoftDelete` operation makes one on the found mailboxes first:
a purge action needs a Compliance Search.

### Cleanup and Exchange 2019 behavior

`Remove-EmtSearches` obeys the Exchange 2019 behavior (see the README):

1. It skips a search that still runs (`Get-EmtComplianceSearchStatus`).
2. It removes the action `<search name>_Purge` by name. It never lists all actions.
3. It removes the shadow search `<search name>-shadow`.
4. It removes the Compliance Search without `-ErrorAction Stop`. It tries two times.
5. It examines the search with `Get-EmtComplianceSearchInfo` (by `-Identity`) for up to 60 seconds.

`Get-EmtComplianceSearchInfo` gives Exists = true, false or unknown. "Not found" (`Test-EmtNotFoundText`) is false.
Another error is unknown: EMT does not report a removal.

`Test-EmtSearchName` accepts only `EMT_yyyyMMdd_HHmmss` and the older `MC_yyyyMMdd_HHmmss`, with or without `-shadow`.

## Window

The window has a common top line (connection, timeout, language), two tabs and a common bottom part
(log, connection, status, progress, **Cancel**).

`Update-EmtConnIndicator` sets the indicator of the bottom line: green = connected, grey = not connected.
The missing roles are not in the window: the log shows one warning for each function that is not available.

| Tab | Main controls |
|---|---|
| Message search | `$grpTr` criteria, `$gridTrMessages`, detail tabs `$trDetail` (recipients, events, mailbox), actions |
| Delete | `$grpCrit` criteria, `$lblDelSource`, `$lblSummary`, result tabs `$tabs`, actions |

`Set-UiState` sets the state of each control from the current data:

| Control | Enabled when |
|---|---|
| Find (Message search) | Connected, Tracking |
| Check the recipient mailbox | A message and an internal recipient are selected, MailboxCheck |
| Delete this message… | A message is selected, Search or Verify |
| Find copies | Connected, Search or Verify |
| Delete (soft) | A clean result, no delete yet, SoftDelete + Search + Verify |
| Delete permanently | A clean result, no hard delete yet, HardDelete |
| Cancel | An operation runs |

`Open-EmtDeleteForMessage` fills the tab **Delete** from the selected message (sender, subject, message ID, period)
and sets `$script:DelFromTrace`. It does not copy the recipients: the search of copies uses all found mailboxes.
`Invoke-EmtInvalidate` clears the delete result and `$script:DelFromTrace` when the user changes a delete criterion.

### Criteria layout

`New-EmtCriteriaTable` makes a table with two pairs "label + field" in each row. The two field columns have 50% each.
`New-EmtPeriodPanel` makes the line **Period**: start — end and the buttons **Today**, **Yesterday**, **7 days**.
`New-EmtActionRow` puts the main button of the tab at the right edge of the fields.
`New-EmtCriteriaGroup` makes the group box. Its height comes from the content.
`Update-EmtCriteriaLabels` gives the label columns the same width on both tabs.
Call it after the labels are on the form (they need the font of the form) and after each language change.

### Start and close

| Step | Code |
|---|---|
| Start error | A `trap` at the top of the script shows a message box. The console is hidden and cannot show the error. |
| Console | `EMT.cmd` starts PowerShell with `-WindowStyle Hidden`. The script also hides a console that has only its own process (`GetConsoleProcessList`) and a `-File` start. |
| Front | The `Shown` handler sets `TopMost` for one moment and activates the window. |
| Size | `Save-EmtSettings` keeps the normal bounds (`RestoreBounds` when maximized) and the maximized state. At the start, EMT uses them only if the window is on a screen. |

### Tables

`Set-EmtGridData` binds a `DataTable`. The columns have the property names. `-Hidden` columns give data for the row
colors only. `Update-EmtGridHeaders` sets the header texts in the current language.
`Resize-EmtGridColumns` sets the initial widths once, when the grid is visible. The column mode is `None`:
the user can change the widths, and a sort keeps them.
`New-EmtGrid` fixes the row height and the header height.

`Set-EmtAccentButton` gives the main button of each tab a light green background.
A disabled button gets the usual grey background.

### Languages

`EMT.Core.ps1` loads `EMT.Strings.ps1`. Thus the window, the runspace and the tests have `T`.
The runspace reads `$global:EmtSync.Lang`. Thus a language change also changes the log of the next operation.

The window registers each control with a text key (`Set-EmtLoc`, `Set-EmtTip`).
`Update-EmtLanguage` applies the texts again and shows the result views again.
The views make all texts at show time (statuses, hints, status lines, findings). They do not keep texts of the old language.
The audit file uses fixed English values. The language does not change them.

## Code rules

- Use `Set-StrictMode -Version 2.0` in the core. Use `Get-EmtProp` to read a property that can be missing.
- Do not use `$script:` variables in the core. The core runs in different scopes.
- Use `Write-EmtLog` for messages. Do not use `Write-Host` in the core.
- Call `Test-EmtCancel` in each loop that calls Exchange.
- Do not use `-ErrorAction Stop` with a `*-ComplianceSearch*` cmdlet. A unit test examines the source code.
- Do not write `@($list)` for a `System.Collections.Generic.List[object]`: PowerShell 5.1 fails with
  "Argument types do not match". Use `$list.ToArray()`.
- Do not use `Measure-Object -Sum` for a list that can be empty. Use `Get-EmtSum`.
- Do not put user interface text in the code. Add a key to `Lang/en.lng` and `Lang/ru.lng`, and use `T 'key'`.
  A unit test finds keys in the code that are not in `en.lng`.
- Keep the core free of Windows Forms types.
