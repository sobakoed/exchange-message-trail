# Architecture

This document gives the structure of the MailCleaner code.

## Files

| File | Content |
|---|---|
| `MailCleaner.ps1` | The window, the background runspace and the event handlers |
| `src/MailCleaner.Core.ps1` | The logic without the user interface |
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
| `Search` | `Invoke-McSearch` | Search result |
| `SoftDelete` | `Invoke-McSoftDelete`, then `Get-McCopyStatus` | Delete status and verification |
| `HardDelete` | `Invoke-McHardDelete`, then `Get-McCopyStatus` | Delete counts and verification |
| `Verify` | `Get-McCopyStatus` | Verification |
| `ListOld` | `Get-McOldSearches` | Old `MC_*` searches |
| `RemoveOld` | `Remove-McSearches` | None |

## Search pipeline

`Invoke-McSearch` does these steps:

1. `Complete-McCriteria` adds the missing keys. `Test-McCriteria` examines the criteria.
2. `Find-McTrackingEvents` reads the tracking log of all transport servers.
3. `Get-McMessageSummary` makes one row for each message ID.
4. `Resolve-McRecipients` classifies each address: internal, external, group or other.
5. `New-McSearchQuery` makes the KQL query.
6. MailCleaner makes and starts a Compliance Search `MC_<date>_<time>`.
7. `Wait-McComplianceSearch` waits for the end. It stops the search on cancel.
8. `Get-McSearchHits` reads `SuccessResults`. `Get-McMissingMailboxes` finds the mailboxes without the message.

The search locations are the internal recipient mailboxes and the internal sender mailboxes.
If there are no internal recipients, the location is `All`.

## Pure functions

These functions do not call Exchange. The unit tests examine them.

| Function | Task |
|---|---|
| `Split-McAddressList` | Splits the address text |
| `Get-McKqlDateRange` | Gives the KQL date range (the end date is exclusive) |
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
| Not connected | Server field, Connect |
| Connected, no result | Criteria, Find, Cleanup |
| Operation runs | Cancel only |
| Result with messages, no search errors | Delete buttons (if the role is available), Verify, Export |

A change of a criterion clears the result (`Invoke-McInvalidate`).

## Code rules

- Use `Set-StrictMode -Version 2.0` in the core.
- Do not use `$script:` variables in the core. The core runs in different scopes.
- Use `Write-McLog` for messages. Do not use `Write-Host` in the core.
- Call `Test-McCancel` in each loop that calls Exchange.
- Keep the core free of Windows Forms types.
