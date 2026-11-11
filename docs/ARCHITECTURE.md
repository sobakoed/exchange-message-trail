# Architecture

This document gives the structure of the Exchange Message Trail (EMT) code.

## Files

| File | Content |
|---|---|
| `ExchangeMessageTrail.ps1` | The window (WPF), the background runspace and the event handlers |
| `gui/MainWindow.xaml` | The layout of the window |
| `gui/Styles.xaml` | The styles of all controls |
| `gui/Theme.Dark.xaml`, `gui/Theme.Light.xaml` | The colors of the dark and the light theme |
| `src/EMT.Core.ps1` | The logic without the user interface |
| `src/EMT.Strings.ps1` | The language loader, the function `T` and the language file check |
| `Lang/*.lng` | The user interface texts, one file for each language |
| `EMT.cmd` | The launcher (STA mode, hidden console) |
| `ExchangeMessageTrail.Classic.ps1`, `EMT-Classic.cmd` | The classic window (Windows Forms): a fallback for one version |
| `tests/EMT.Core.Tests.ps1` | The Pester 5 unit tests for the core, the language files and the window files |

All `.ps1` files use UTF-8 with BOM and CRLF. The `.xaml` files use UTF-8 and CRLF.
Windows PowerShell 5.1 reads a file without BOM in the ANSI code page. This changes the Cyrillic text.

## Threads

```
UI thread (STA)                                 Background runspace
──────────────────────────────                  ────────────────────────────────────
Window, controls, event handlers                Remote session to Exchange
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
| 10 | Delivered | `DELIVER`, `DUPLICATEDELIVER` |
| 9 | External | `SEND` to an external address, `SENDEXTERNAL` |
| 8 | Forwarded | `DROP` with `RESOLVER.FWD.Forwarded` |
| 7 | Failed | `FAIL` |
| 6 | GroupBlocked | `DROP` with `RESOLVER.GRP.BlockedMessage` |
| 5 | Dropped | `DROP` with a `4.x.x` / `5.x.x` code or without a cause, `POISONMESSAGE` |
| 4 | Suppressed | `DROP` with another `2.x.x` code or an agent text (`RESOLVER.MSGTYPE.AF`, `RESOLVER.OOF.*`, `RESOLVER.MSGTYPE.NDR`) |
| 3 | Deferred | `DEFER` |
| 2 | InTransit | `SEND` to an internal address, `EXPAND`, `REDIRECT`, `HAREDIRECT`, `TRANSFER`, `RESOLVE` |
| 1 | Group | A group address without its own delivery; `DROP` with `RESOLVER.GRP.Expanded` or `RESOLVER.GRP.NotRedirected` |
| 0 | Unknown | `RECEIVE`, `SUBMIT` |

`SEND` to an internal recipient is a transfer to another Exchange server, not the end of the delivery.
`SENDEXTERNAL` is a send through a Send connector, for example a hidden copy that a transport rule adds.
`Get-EmtDropStatus` reads the server answer of a `DROP` event. On Exchange 2019, most `DROP` events have
the success code `250 2.1.5`: they are usual steps, not errors. A `DROP` with `CAT.DuplicateRecipient` gives no status.
`Get-EmtErrorStatuses` gives the statuses that are errors: Failed, GroupBlocked, Dropped.
`Get-EmtStatusOrder` gives the order of the statuses in the status line of a message.
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

The window is WPF. The layout is in `gui/MainWindow.xaml`, the styles of all controls are in `gui/Styles.xaml`,
the colors are in `gui/Theme.Dark.xaml` and `gui/Theme.Light.xaml`.
The window has a common top line (server, **Connect**, timeout, language, theme), two tabs and a common bottom part
(log, connection, status, progress, **Cancel**).

| Item | How it operates |
|---|---|
| Elements | The script makes a variable for each element with `x:Name`: `x:Name="btnSoft"` gives `$btnSoft`. |
| Texts | `Tag="key"` of an element is a text key. `Update-EmtTexts` sets the texts (TextBlock, button, check box, tab). |
| Tooltips | `Set-EmtTip` registers a control and a key. |
| Theme | `Add-EmtTheme` adds the theme dictionary (first) and the styles. `Set-EmtTheme` replaces the theme dictionary. All styles use `DynamicResource`: the window changes at once. |
| System colors | `Get-EmtThemeDictionary` adds the system color keys (`SystemColors.*BrushKey`) with the theme colors. One XAML file cannot have more than one system key. |
| Title bar | `Set-EmtTitleBar` calls `DwmSetWindowAttribute` (dark title bar for the dark theme). Older Windows versions ignore it. |
| Dialogs | `New-EmtDialog` makes a window with the theme. `Show-EmtMessage`, `Show-EmtPick`, `Show-EmtConfirm` (`New-EmtConfirmWindow`), `Show-EmtInput`. The folder selection of the export is the standard Windows dialog. |
| Errors | `Dispatcher.UnhandledException` catches an error of an event handler. The window stays open, the log and a message show the error. |

`Update-EmtConnIndicator` sets the indicator of the bottom line: green = connected, grey = not connected.
The missing roles are not in the window: the log shows one warning for each function that is not available.

| Tab | Main controls |
|---|---|
| Message search | `$cardTr` criteria, `$gridTrMessages`, detail tabs `$trDetail` (recipients, events, mailbox), actions |
| Delete | `$cardDel` criteria, `$lblDelSource`, `$lblSummary`, result tabs `$tabs`, actions |

`Set-UiState` sets the state of each control from the current data:

| Control | Enabled when |
|---|---|
| Find (Message search) | Connected, Tracking |
| Check the recipient mailbox | A message and an internal recipient are selected, MailboxCheck |
| Delete this message… | A message is selected, Search or Verify |
| Find copies | Connected, Search or Verify |
| Delete (soft) | A clean result, no delete yet, SoftDelete + Search + Verify |
| Delete permanently | A clean result, no hard delete yet, HardDelete |
| Verify | Copies found, Verify |
| Clean up old searches | Connected, Search |
| Cancel | An operation runs |

`Open-EmtDeleteForMessage` fills the tab **Delete** from the selected message (sender, subject, message ID, period)
and sets `$script:DelFromTrace`. It does not copy the recipients: the search of copies uses all found mailboxes.
`Invoke-EmtInvalidate` clears the delete result and `$script:DelFromTrace` when the user changes a delete criterion.
`$script:Loading` stops the invalidation while the program fills the fields itself.

### Criteria layout

The criteria of each tab are a grid with two pairs "label + field" in each row. The two field columns have the same width.
The label columns have `SharedSizeGroup` (`Lbl1`, `Lbl2`). Thus the fields have the same size and position on both tabs,
also after a language change. The main button of the tab is at the right edge of the fields.
Enter starts the main button of the open tab (`IsDefault`).

### Dates

`Set-EmtDateCulture` gives the window thread a copy of the Windows regional settings:
the date format and the first day of the week are the same as in the other programs.
The names of the months and the week days come from the interface language.
The window does not set its property `Language`: then the date pickers use the culture of the thread.
Each date picker has a time field. `Get-EmtPeriodValue` gives the date and the time, `Set-EmtPeriodValue` sets them.
The time uses the short time format of Windows (`Get-EmtTimeFormat`). `ConvertTo-EmtTime` reads the format of Windows
and the 24-hour format. A wrong time goes back to the last correct time (`$script:LastTime`).
PowerShell gives a `DateTime` for a `DateTime?`: the property `Value` does not exist.
`Get-EmtPeriodEnd` makes the end of the period include its minute: 23:59 is 23:59:59.999. The quick periods end at 23:59.
`Open-EmtDeleteForMessage` sets the period from 5 minutes before the message to 23:59 of its day:
the KQL range then has one day (`Get-EmtKqlDateRange`).
`Get-EmtExtraItemMailboxes` gives the mailboxes with more items than the tracking log has messages: the summary line
shows an orange warning, and the delete confirmation needs the typed word.

### Start and close

| Step | Code |
|---|---|
| Start error | A `trap` at the top of the script shows a message box. The console is hidden and cannot show the error. |
| Console | `EMT.cmd` starts PowerShell with `-WindowStyle Hidden`. The script also hides a console that has only its own process (`GetConsoleProcessList`) and a `-File` start. |
| Front | The `ContentRendered` handler sets `Topmost` for one moment and activates the window. |
| Size | `Save-EmtSettings` keeps the normal bounds (`RestoreBounds` when maximized) and the maximized state in the key `WpfWindow` (WPF units). At the start, EMT uses them only if the window is on the screen. |
| Settings | `Save-EmtSettings` keeps the other keys of `settings.json`. The classic window keeps its size in the key `Window` (pixels). |

### Tables

`Set-EmtGridData` binds a `DataTable`. The columns have the property names. A property with an empty header key
is data only. The header texts come from the current language.
`Resize-EmtGrids` sets the initial widths once, when the grid is visible (60–450 px, the last column gets the free space).
Then the widths are fixed: a sort keeps them. The row height is fixed (`CanUserResizeRows = False`).

The rows have no fill color. Two parameters of `Set-EmtGridData` give the colors:

| Parameter | Result | Template |
|---|---|---|
| `-Mark` | The color of the left edge of a row: Green, Blue, Red, Yellow, Grey or none | `EmtMarkCell` |
| `-PartsColumn`, `-Parts` | A dot and a colored text for each part of a status | `EmtPartsCell` |

| Table | Left edge |
|---|---|
| Messages (Message search) | `Get-EmtStatusMark`: the worst status of the recipients |
| Recipients (Message search) | `Get-EmtStatusTone`: the status of the recipient |
| Events, Recipient mailbox | Red for an error, yellow for a delay or a warning |
| Delete: messages, recipients, verification | Yellow for a notification of Exchange; red for an external recipient; red / yellow / green for the verification |

`Set-EmtSummaryParts` makes a summary line from the parts of the template (`{0}`, `{1:N1}`, ...).
Only the number of errors and the warnings before the line are red.

### Languages

`EMT.Core.ps1` loads `EMT.Strings.ps1`. Thus the window, the runspace and the tests have `T`.
The runspace reads `$global:EmtSync.Lang`. Thus a language change also changes the log of the next operation.

`Update-EmtLanguage` applies the texts again and shows the result views again.
The views make all texts at show time (statuses, hints, status lines, findings). They do not keep texts of the old language.
The audit file uses fixed English values. The language does not change them.

### Classic window

`ExchangeMessageTrail.Classic.ps1` is the earlier window (Windows Forms). It has the same functions and the same core.
`EMT-Classic.cmd` starts it. It stays for one version as a fallback. A later version removes it.

## Code rules

- Use `Set-StrictMode -Version 2.0` in the core. Use `Get-EmtProp` to read a property that can be missing.
- Do not use `$script:` variables in the core. The core runs in different scopes.
- Use `Write-EmtLog` for messages. Do not use `Write-Host` in the core.
- Call `Test-EmtCancel` in each loop that calls Exchange.
- Do not use `-ErrorAction Stop` with a `*-ComplianceSearch*` cmdlet. A unit test examines the source code.
- Do not write `@($list)` for a `System.Collections.Generic.List[object]`: PowerShell 5.1 fails with
  "Argument types do not match". Use `$list.ToArray()`.
- Do not use `Measure-Object -Sum` for a list that can be empty. Use `Get-EmtSum`.
- Do not put user interface text in the code or in XAML. Add a key to `Lang/en.lng` and `Lang/ru.lng`,
  and use `T 'key'` or `Tag="key"`. A unit test finds keys in the code and in `gui/*.xaml` that are not in `en.lng`.
- Keep the core free of window types (WPF, Windows Forms).
- Use only `DynamicResource` for the colors. A unit test finds resources that the theme or the styles do not have.
- Give an element an `x:Name` before the script uses it. A unit test finds variables without an element.
- A script block with `GetNewClosure()` does not see the functions of the script. Use it only for local variables
  of a dialog (for example `$w.DialogResult = $true`).
- Do not open a modal dialog in an automatic test. Make the window with `New-EmtConfirmWindow` and show it outside the screen.
