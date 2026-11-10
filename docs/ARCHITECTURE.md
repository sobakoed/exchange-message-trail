# Architecture

This document gives the structure of the MailCleaner code.

## Files

| File | Content |
|---|---|
| `MailCleaner.ps1` | The window, the background runspace and the event handlers |
| `src/MailCleaner.Core.ps1` | The logic without the user interface |
| `src/MailCleaner.Strings.ps1` | The language loader, the function `T` and the language file check |
| `Lang/*.lng` | The user interface texts, one file for each language |
| `MailCleaner.cmd` | The launcher (STA mode, minimized console) |
| `tests/MailCleaner.Core.Tests.ps1` | The Pester 5 unit tests for the core |

All `.ps1` files use UTF-8 with BOM and CRLF.
Windows PowerShell 5.1 reads a file without BOM in the ANSI code page.
This changes the Cyrillic text.

## Threads

```
UI thread (STA)                                 Background runspace
──────────────────────────────                  ────────────────────────────────────
Form, controls, event handlers                  Remote session to Exchange
        │ Start-McOperation ─── BeginInvoke ──▶ . MailCleaner.Core.ps1
        │                                       Invoke-McOperation -Name <operation>
Timer (400 ms) ◀── $Sync.Log (queue) ────────── Write-McLog
        │      ◀── $Sync.Progress / Status ──── Set-McProgress
        │      ─── $Sync.Cancel ──────────────▶ Test-McCancel
        └ Complete-McOperation ◀── EndInvoke ── result object
```

- The window never calls Exchange. The background runspace does all Exchange work.
- The runspace keeps one remote session (`$global:McExSession`) for all operations.
- `$Sync` is a synchronized hashtable. The UI thread and the runspace share it.
- The timer reads the log queue and the progress. It calls `Complete-McOperation` at the end of an operation.
- The timer stops while a result handler shows a dialog. This prevents a second call.

## Connection

`Connect-McExchange` does these steps:

1. It makes a session with `New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Kerberos`.
2. It imports only the necessary cmdlets with `Import-PSSession -CommandName`.
3. It imports the module into the global scope of the runspace.
4. It returns the capabilities (`Get-McCapabilities`).

The Exchange session is a NoLanguage session. It does not run script blocks.
For this reason, the logic runs in the runspace. The runspace calls the imported cmdlets.

The runspace uses the execution policy `Bypass`.
`Import-PSSession` makes a temporary script module. A strict policy can block it.

## Operations

`Invoke-McOperation` is the only entry point for the window.

| Operation | Function | Result |
|---|---|---|
| `Connect` | `Connect-McExchange` | Capabilities |
| `Resolve` | `Resolve-McAddressInputs` (`Get-Recipient -Anr`) | Candidates for each partial address |
| `Search` | `Invoke-McSearch` | Search result |
| `SoftDelete` | `Invoke-McSoftDelete`, then `Get-McCopyStatus` | Delete status and verification |
| `HardDelete` | `Invoke-McHardDelete`, then `Get-McCopyStatus` | Delete counts and verification |
| `Verify` | `Get-McCopyStatus` | Verification |
| `ListOld` | `Get-McOldSearches` | Old `MC_*` searches |
| `RemoveOld` | `Remove-McSearches` | None |

## Search pipeline

Before the search, the window examines the senders and the recipient (`Test-McFullAddress`).
For a part of an address or a name, it starts the operation `Resolve`.
`Complete-McAddressResolve` puts the full addresses in the fields (automatic for one candidate,
`Show-McPick` for more). Then `Start-McSearch` starts the search. Thus `Invoke-McSearch` gets only full addresses.

`Invoke-McSearch` does these steps:

1. `Complete-McCriteria` adds the missing keys. `Test-McCriteria` examines the criteria.
2. With **Только ящик получателя**, `Resolve-McRecipients` finds the recipient mailbox.
3. `Find-McTrackingEvents` reads the tracking log of all transport servers.
4. `Get-McMessageSummary` makes one row for each message ID.
5. `Resolve-McRecipients` classifies each address: internal, external, group or other.
6. `New-McSearchQuery` makes the KQL query.
7. MailCleaner selects the search method (see below).
8. `Get-McMissingMailboxes` finds the mailboxes without the message.

The log shows the time of each stage. The result contains `Method` and `Seconds`.

### Search locations

| Mode | Locations |
|---|---|
| Только ящик получателя | The recipient mailbox only |
| Искать во всех ящиках | `All` |
| Default | Internal recipient mailboxes and internal sender mailboxes. `All` if there are none. |

### Search method

| Condition | Method | Function |
|---|---|---|
| 1 to `Get-McQuickSearchLimit` (10) mailboxes, `Search-Mailbox` available | Quick | `Invoke-McQuickSearch` |
| More mailboxes, or `All` | Compliance Search | `New-McComplianceSearchRun`, `Wait-McComplianceSearch`, `Get-McSearchHits` |

A quick search does not make a Compliance Search. `SearchName` is empty.
The `SoftDelete` operation makes a Compliance Search on the found mailboxes first.
A purge action needs a Compliance Search.

### Tracking log

All events of one message have the same sender and subject.
Thus the sender query and the subject query return all events of the message.
`Find-McTrackingEvents` does a second pass by message ID only with `-Recipient` and `-CompleteRecipients`.
`Select-McUniqueEvents` removes the duplicate events of the two queries.

### Recipients

`Connect-McExchange` reads the accepted domains (`Get-AcceptedDomain`).
`Resolve-McRecipients` does not look up the addresses outside these domains.
It looks up the other addresses in batches: one `Get-Recipient -Filter` call for 20 addresses.
The filter compares `EmailAddresses`, so aliases also match.
If a batch call fails, it looks up the batch address by address.

## Pure functions

These functions do not call Exchange. The unit tests examine them.

| Function | Task |
|---|---|
| `Split-McAddressList` | Splits the address text |
| `Get-McKqlDateRange` | Gives the KQL date range (the end date is exclusive) |
| `Get-McBaseSubject` | Gives the subject without reply and forward prefixes |
| `Get-McLargeDeleteWarning` | Gives the warning text for a large delete |
| `Test-McHardDeleteScope` | Compares the `Search-Mailbox` counts with the search counts before a hard delete |
| `Test-McInternalDomain` | Finds an address in the accepted domains |
| `ConvertFrom-McSizeText` | Reads the size text of `Search-Mailbox` |
| `Select-McUniqueEvents` | Removes duplicate tracking events |
| `New-McSearchQuery` | Makes the KQL query. It refuses a query with only dates. |
| `Complete-McCriteria`, `Test-McCriteria` | Prepare and examine the criteria |
| `Get-McSearchHits` | Reads `SuccessResults` |
| `Test-McServerGeneratedId` | Finds a message ID that Exchange made |
| `Get-McMessageSummary` | Makes the message rows |
| `Get-McMissingMailboxes` | Finds the mailboxes without hits |
| `New-McSearchName` | Gives the search name |
| `Write-McAudit` | Writes one audit row |

## Window states

`Set-UiState` sets the state of each control from the current data:

| Condition | Enabled controls |
|---|---|
| Not connected | Server field, Connect, criteria |
| Connected, no result | Criteria, Find, Cleanup |
| Operation runs | Cancel only |
| Result with messages, no search errors | Delete buttons (if the role is available), Verify, Export |
| After a soft delete | Hard delete, Verify, Export |
| After a hard delete | Verify, Export |

`$script:Deleted` keeps the last delete on the current result.
A change of a criterion clears the result (`Invoke-McInvalidate`).

### Delete operations

| Operation | Steps |
|---|---|
| `SoftDelete` | Compliance Search (only after a quick search), purge, `Get-McSettledCopyStatus` |
| `HardDelete` | `Get-McCopyStatus -TotalOnly`, `Test-McHardDeleteScope`, `Invoke-McHardDelete`, `Get-McSettledCopyStatus` |

`Invoke-McHardDelete` stops before the next mailbox on cancel. It returns the partial result with `Cancelled = true`.
`Complete-McOperation` writes an audit row (`Cancelled` or `Failed`) when a delete operation stops with an error.
`Get-McSettledCopyStatus` examines again the mailboxes where the user can still see the message (index lag).

### Tables

`Set-McGridData` binds a `DataTable` and marks the grid in `$script:GridNeedsSize`.
`Resize-McGridColumns` sets the initial widths once, when the grid is visible.
The column mode is `None`: the user can change the widths, and a sort keeps them.

## Languages

`MailCleaner.Core.ps1` loads `MailCleaner.Strings.ps1`. Thus the window, the runspace and the tests have `T`.

| Item | Function |
|---|---|
| `Read-McLanguageFile` | Reads one `.lng` file: code, name, author, texts, problems |
| `Get-McStringTable` | Reads all files in `Lang` once. Cache: `$global:McStringTable` |
| `Get-McLanguages` | The list for the language box: code, name, author |
| `Get-McLanguage` | The current language: `$global:McSync.Lang`, then `$global:McLang`, then `en` |
| `Get-McDefaultLanguage` | The Windows language if a file exists, Russian for `uk` and `be`, else English |
| `T 'key' $a $b` | The text in the current language. Fallback: English, Russian, then the key itself. A bad placeholder in a translation does not stop the tool. |
| `Test-McLanguageFile` | Compares a file with `en.lng`: unknown and missing keys, placeholders, duplicates, bad lines |

The runspace reads `$global:McSync.Lang`. Thus a language change also changes the log and the errors of the next operation.

The window registers each control with a text key (`Set-McLoc`, `Set-McTip`).
`Update-McLanguage` applies the texts again, and shows the result views again.
The grid columns have the property names. `Update-McGridHeaders` sets the header texts.
The row colors use the property names. Thus they do not depend on the language.

The audit file uses fixed English values (`Search`, `Clean`, `Cancelled` …). The language does not change them.

## Code rules

- Use `Set-StrictMode -Version 2.0` in the core.
- Do not use `$script:` variables in the core. The core runs in different scopes.
- Use `Write-McLog` for messages. Do not use `Write-Host` in the core.
- Call `Test-McCancel` in each loop that calls Exchange.
- Keep the core free of Windows Forms types.
- Do not put user interface text in the code. Add a key to `Lang/en.lng` and `Lang/ru.lng`, and use `T 'key'`.
  A unit test finds keys in the code that are not in `en.lng`.
