# Exchange Message Trail (EMT)

Exchange Message Trail is a Windows tool for Exchange Server 2016 and 2019 (on-premises).
It answers these questions:

- Did a user send a message?
- Did the message arrive?
- Why did the message not arrive? At which step did the delivery stop?

When it is necessary, EMT also deletes a message from the mailboxes.

The window has two tabs:

| Tab | Function | Speed | Changes data |
|---|---|---|---|
| **Message search** | Finds a message in the message tracking log. Shows the delivery status of each recipient. Examines the mailbox of a recipient. | Seconds | No |
| **Delete** | Finds the copies of a message in the mailboxes. Deletes them. Verifies the result. | Seconds to minutes | Yes |

The user interface is available in English and Russian. You can add a language: see [Languages](#languages).

## Contents

- [Requirements](#requirements)
- [Permissions](#permissions)
- [Install](#install)
- [Start](#start)
- [Window](#window)
- [Message search](#message-search)
- [Mailbox check](#mailbox-check)
- [Delete](#delete)
- [Files](#files)
- [Languages](#languages)
- [Limits](#limits)
- [Troubleshooting](#troubleshooting)
- [More documents](#more-documents)

## Requirements

| Item | Requirement |
|---|---|
| Exchange | Exchange Server 2016 or 2019, on-premises |
| Computer | The Exchange server, or a Windows computer in the same domain |
| PowerShell | Windows PowerShell 5.1 |
| .NET | .NET Framework 4.5 or later |
| Network | HTTP access to `http://<server>/PowerShell/` |
| Authentication | Kerberos (domain account) |

You do not have to install the Exchange Management Tools.
EMT connects to Exchange through remote PowerShell.

## Permissions

EMT examines the permissions of your account when it connects. The log shows a warning for each function that is not available.
A button stays disabled if your account does not have the necessary role.

**Message search only (for example, the help desk):**

| Function | Cmdlets | Role group |
|---|---|---|
| Message search | `Get-MessageTrackingLog`, `Get-TransportService`, `Get-Recipient` | Organization Management, Records Management or Compliance Management |
| Mailbox check | `Get-InboxRule`, `Get-MailboxJunkEmailConfiguration`, `Get-Mailbox`, `Get-MailboxStatistics` | Organization Management or Recipient Management |
| Accepted domains | `Get-AcceptedDomain` (optional, for speed) | Organization Management or View-Only Organization Management |

**Delete:**

| Function | Cmdlets | Role group |
|---|---|---|
| Search | `New-ComplianceSearch`, `Start-ComplianceSearch`, `Get-ComplianceSearch` | Discovery Management |
| Verify | `Search-Mailbox` | Discovery Management |
| Soft delete | `New-ComplianceSearchAction` | Discovery Management |
| Hard delete | `Search-Mailbox -DeleteContent` | Mailbox Import Export role (assign it manually) |
| Cleanup of shadow searches | `Get-MailboxSearch`, `Remove-MailboxSearch` | Discovery Management |

To assign the Mailbox Import Export role, do this step in the Exchange Management Shell:

```powershell
New-ManagementRoleAssignment -Role "Mailbox Import Export" -User "<account>"
```

Close EMT and start it again after a role change.

## Install

1. Copy the repository folder to the computer.
2. Keep the folder structure. `ExchangeMessageTrail.ps1` needs the folders `src` and `Lang`.
3. If you downloaded a ZIP file, unblock the files:

```powershell
Get-ChildItem -Recurse | Unblock-File
```

## Start

Use one of these procedures.

**Procedure A: the launcher**

1. Double-click `EMT.cmd`.

**Procedure B: PowerShell**

1. Open Windows PowerShell.
2. Go to the EMT folder.
3. Type this command:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\ExchangeMessageTrail.ps1 -Server exch01.contoso.local
```

| Parameter | Function |
|---|---|
| `-Server` | The FQDN of the Exchange server. Optional. EMT keeps the last server name. On an Exchange server, the default value is the name of that server. |
| `-Language` | The language code, for example `en` or `ru`. Optional. See [Languages](#languages). |

**Connect:**

1. In the field **Exchange server (FQDN)**, type the full server name.
2. Click **Connect**.
3. Make sure that the bottom line shows a green indicator and "Connected".
4. Make sure that the log shows no warning about a missing role.

## Window

| Area | Controls |
|---|---|
| Top line | Server, **Connect**, operation timeout, language |
| Tab **Message search** | Criteria, **Find**, summary line, messages, details of the selected message, actions |
| Tab **Delete** | Criteria, **Find copies**, source line, summary line, result tabs, delete actions |
| Log | The messages of the session (common for both tabs) |
| Bottom left | The connection: a green indicator and the server name, or a grey indicator and "Not connected" |
| Bottom right | The running operation: status, progress, **Cancel** |

- Push Enter to start the search of the open tab.
- To stop an operation, click **Cancel**. EMT stops after the current step.
- **Tables:** Drag a column border to change the width. Click a column header to sort.
  Drag a column header to move the column. Push Ctrl+C to copy the selected rows with headers.
- The tab name shows the row count, for example **Recipients (4)**.

## Message search

The trace reads only the message tracking log. It does not open the mailboxes. It does not change data.

### Procedure

1. Open the tab **Message search**.
2. Type one or more criteria (see the table below).
3. Set the period. Use **Today**, **Yesterday** or **7 days**.
4. Click **Find**.
5. In the list of messages, select a message.
6. Examine the tab **Recipients**: the delivery status of each recipient.
7. If a recipient did not get the message, read the columns **Explanation** and **Server answer**.
8. Double-click a message to see the tab **Events**: all tracking events of the message.

### Criteria

| Field | Use |
|---|---|
| Senders | One or more addresses, parts of addresses or names. Separator: comma or semicolon. |
| Recipient | One address, a part of an address or a name. A recipient alone is permitted: EMT shows all messages to this address. |
| Subject contains | A part of the subject |
| Period | Start and end: date and time. The tracking log uses the exact time. |
| Message ID | The internet message ID |

Set one criterion or more. A period alone is not permitted.

**Partial addresses and names:** EMT finds the full address in the Exchange address book (`Get-Recipient -Anr`).
ANR compares the text with the alias, the display name, the first name, the last name and the start of the e-mail addresses.

| Matches | Result |
|---|---|
| One | EMT puts the full address in the field. The log shows the change. |
| More than one | A window shows the list. Select the address. For senders, you can select more than one (Ctrl+click). |
| None | An error window shows the text. The search does not start. |

### Delivery status

| Status | Color | Meaning |
|---|---|---|
| Delivered | Green | Exchange put the message in the mailbox (event `DELIVER`) |
| Sent to an external system | Blue | Exchange sent the message to an external server (event `SEND`) |
| Delivery failed | Red | Exchange did not deliver the message (event `FAIL`) |
| Dropped | Red | Exchange dropped the message (event `DROP` or `POISONMESSAGE`) |
| Deferred | Yellow | The delivery is delayed. Exchange tries again (event `DEFER`). |
| In transit | Grey | The message is between Exchange servers |
| Group: sent to the members | Grey | The address is a group. The members get the message. |
| No delivery data | Grey | The tracking log has only the receive event |

The status of a recipient comes from the most important event. For example, `DELIVER` after `DEFER` gives **Delivered**.

The column **Explanation** explains the SMTP code in the column **Server answer**:

| Code | Explanation |
|---|---|
| `5.1.1`, `5.1.10` | The address does not exist |
| `5.2.2` | The recipient mailbox is full |
| `5.2.3`, `5.3.4` | The message is too large |
| `4.4.7`, `5.4.7` | The delivery time expired |
| `5.4.1`, `5.4.4`, `5.4.6`, `5.4.14` | No route: the domain or the server is not available, or a loop |
| `4.4.x` | No connection to the recipient server |
| `5.7.x` | Rejected: no permission, a delivery restriction, a policy or a filter |
| Other `5.x.x` / `4.x.x` | A permanent / temporary error |

**Message list:**

| Row color | Meaning |
|---|---|
| Red | One or more recipients did not get the message |
| Yellow | One or more deliveries are not complete |

The column **Note** marks a message that Exchange made, not a user: for example, an inbox rule or mailbox forwarding.

### Nothing found

When the trace finds no message, the log shows what to examine:

- the period: the message can be earlier or later;
- the addresses: the message can come from another address or go to a group;
- the subject;
- the folders Outbox and Sent Items of the sender: the message can be not sent.

The tracking log keeps data for 30 days by default.

## Mailbox check

The tracking log ends when Exchange puts the message in the mailbox.
After that, an inbox rule, the junk e-mail filter or mailbox forwarding can move the message.
The mailbox check finds these causes. It does not change data.

**Procedure:**

1. On the tab **Message search**, select a message.
2. On the tab **Recipients**, select an internal recipient.
3. Click **Check the recipient mailbox**.
4. Examine the tab **Recipient mailbox**.

| Check | Finding |
|---|---|
| Rules | An enabled rule that moves, deletes or forwards messages. **Important** if the rule applies to this message (sender or subject), or if the rule has no conditions. |
| Junk e-mail | The sender or the domain is in the blocked list: the message is in the folder Junk E-mail. The filter is off. The sender is in the trusted list. |
| Forwarding | Mailbox forwarding to another address, with or without a copy in the mailbox |
| Quota | The mailbox size and the quota. **Caution** at 90% and more. **Important** at 100%. |

| Severity | Color |
|---|---|
| Important | Red |
| Caution | Yellow |
| Information | None |

If no check finds a cause, the user can have deleted or moved the message.

## Delete

Use the tab **Delete** to find the copies of a message in the mailboxes and to delete them.

> **WARNING:** A hard delete is permanent. The user cannot recover the message.
> Do a soft delete first if you are not sure.

### Procedure from a traced message (recommended)

1. On the tab **Message search**, select the message.
2. Click **Delete this message…**.
3. EMT opens the tab **Delete** with the data of the message: sender, subject, message ID, period.
   The green line says **Data from the tracking log**.
4. Click **Find copies**.
5. Examine the result tabs.
6. Click **Delete (soft)** or **Delete permanently**.
7. Read the summary in the confirmation window.
8. Confirm the delete.
9. Examine the tab **Verification**.

### Procedure with your own criteria

1. Open the tab **Delete**.
2. Type the criteria. The orange line says **Data typed by the user**.
3. Do steps 4–9 of the procedure above.

Use this procedure, for example, for a message that is older than the tracking log.

> **CAUTION:** If you change a criterion after a search, EMT clears the result.
> The line **Data from the tracking log** also changes to **Data typed by the user**.

### Criteria

| Field | Use | Tracking log | Compliance Search (KQL) |
|---|---|---|---|
| Senders | Addresses, parts of addresses or names | `-Sender` | `from:` (OR) |
| Subject contains | A part of the subject | `-MessageSubject` | `subject:"…"` |
| Recipient | One address | `-Recipients` | `to:` / `cc:` / `bcc:` |
| Period | Date and time | exact time | full days only |
| Message ID | The internet message ID | `-MessageId` | subject and sender of the message |
| Attachment name | A file name | not used | `attachment:"…"` |
| Words in the text | A phrase | not used | `"…"` |

| Check box | Function |
|---|---|
| Include forwards (FW:) | Adds the forwarded copies from other senders |
| Recipient mailbox only | Searches and deletes only in the mailbox in **Recipient** |
| Search all mailboxes | Searches all mailboxes |

The last two check boxes exclude each other.

- **Forwarded copy:** a message from another sender with the same base subject.
  The base subject is the subject without `RE:`, `FW:`, `Fwd:`, `AW:`, `WG:`, `Отв:`, `Пересл:` and similar prefixes.
- **Subject text:** EMT reads the subject as text. Square brackets, `*` and `?` have no special meaning.
- **Recipient:** EMT adds the recipient to the KQL query. Without the recipient, `from:sender` finds all messages of the sender.
- **Period:** Compliance Search uses full days only. The KQL range covers the local days and the UTC days.
  Thus the range can be one day larger than the period.
- **Tracking errors:** if a transport server does not answer, the mailbox list can be incomplete.
  The summary line is red and the delete is not available. Select **Search all mailboxes** and do the search again.

### Search scope and speed

| Condition | Mailboxes |
|---|---|
| **Recipient mailbox only** | The recipient mailbox |
| **Search all mailboxes** | All mailboxes |
| The tracking log found the message | The internal recipient and sender mailboxes |
| The tracking log is empty, a recipient is set | The recipient and sender mailboxes |
| The tracking log is empty, no recipient | All mailboxes |

| Mailboxes | Method | Usual time |
|---|---|---|
| 1 to 10 | `Search-Mailbox -EstimateResultOnly` | Seconds |
| More than 10, or all | Compliance Search | 1 to 3 minutes, or more for all mailboxes |

The quick method does not give the number of unindexed items. The summary line shows `n/a`.
A soft delete after a quick search makes a Compliance Search first. This adds 1 to 3 minutes.

### Results

| Tab | Content |
|---|---|
| Messages | One row for each message in the tracking log |
| Tracking log | All tracking events |
| Recipients | Internal mailboxes and external addresses. A red row is an external address. |
| Found in mailboxes | Mailboxes with the message: count and size |
| Not found | Searched mailboxes without the message (only when other mailboxes have it) |
| Verification | The result of the last verification |

**NOT CONFIRMED BY THE TRACKING LOG:** the search found items, but the tracking log has no message.
The summary line is red. Each delete then needs the typed word `DELETE`.

**External recipients:** EMT cannot delete a message from an external mailbox.
Tell your security team about each external recipient.

### Delete types

| | Soft delete | Hard delete |
|---|---|---|
| Button | **Delete (soft)** | **Delete permanently** |
| Cmdlet | `New-ComplianceSearchAction -Purge -PurgeType SoftDelete` | `Search-Mailbox -SearchDumpster -DeleteContent` |
| Result | The message moves to Recoverable Items | The message is deleted from all folders and from Recoverable Items |
| User can recover | Yes, with "Recover Deleted Items" in Outlook | No |
| Confirmation | Summary window. The word `DELETE` for a large or not confirmed delete. | Summary window and the word `DELETE` |
| Safety check | — | `Search-Mailbox` must not find more items than the search |

- **Large delete:** more than 100 items, or more than 5 different messages in the tracking log.
- **Safety check before a hard delete:** if a `Search-Mailbox` count is more than the search count, EMT stops the delete.
- **Holds:** a mailbox on Litigation Hold or In-Place Hold keeps a copy after a hard delete in a hidden folder.

| Last delete | Available |
|---|---|
| None | Soft delete, hard delete, verify |
| Soft delete | Hard delete, verify |
| Hard delete | Verify |

### Verification

A Compliance Search also searches Recoverable Items. After a soft delete, a new Compliance Search shows the same count.
For this reason, EMT verifies with `Search-Mailbox`:

| Column | Meaning |
|---|---|
| In visible folders | Messages that the user can see |
| Total (with Recoverable Items) | Messages in all folders |

| Row color | Meaning |
|---|---|
| Green | The mailbox does not contain the message |
| Red | The user can see the message |
| Yellow | The verification failed for this mailbox |

The search index can show a deleted message for some seconds.
If the user can still see the message, EMT waits 15 seconds and examines the mailbox again (3 times maximum).

### Cleanup of old searches

Each search of the tab **Delete** makes a Compliance Search `EMT_<date>_<time>`.
For each Compliance Search, Exchange also makes a shadow In-Place eDiscovery search `EMT_<date>_<time>-shadow`.

Click **Clean up old searches** to remove the old searches. EMT removes:

- the Compliance Searches `EMT_<date>_<time>` and their actions;
- the shadow searches `EMT_<date>_<time>-shadow`;
- the searches `MC_<date>_<time>` of the older version (MailCleaner).

The age comes from the date and time in the name. To remove the searches of today, type `0` days.
EMT does not remove searches with other names. EMT does not touch a search that still runs.

## Files

| File | Content |
|---|---|
| `%USERPROFILE%\ExchangeMessageTrail\audit.csv` | One row for each trace, mailbox check, search, delete, verification and cleanup |
| `%USERPROFILE%\ExchangeMessageTrail\logs\EMT_<date>.log` | The full log of one session |
| `%APPDATA%\ExchangeMessageTrail\settings.json` | The last server, the language, the options and the timeout. Senders and subject are not kept. |

At the first start, EMT copies `settings.json` and `audit.csv` of the older version (MailCleaner). The log shows the copy.

Audit columns: `Timestamp`, `Operator`, `Computer`, `Server`, `Action`, `SearchName`, `Query`, `Mailboxes`, `Items`, `Result`, `Details`.

| `Action` | `Result` values |
|---|---|
| Trace | `OK`, `Errors` |
| MailboxCheck | `OK`. `Items` = the number of important findings. |
| Search | `OK`, `Errors` |
| SoftDelete, HardDelete, Verify | `Clean`, `RecoverableOnly`, `Visible`, `CheckFailed` |
| SoftDelete, HardDelete (stopped) | `Cancelled`, `Failed` |
| CleanupSearches | `OK` |

**Export:** Click **Export to CSV** on a tab to save its data as CSV files.

> **CAUTION:** The log, audit and export files contain mailbox addresses.
> Keep them in a safe location.

## Languages

The languages are files in the folder `Lang`: `en.lng` (English) and `ru.lng` (Russian).

- Select the language in the list **Language** in the top line. The change occurs immediately.
- EMT keeps the selection in `settings.json`.
- At the first start, EMT uses the Windows language if a file for it exists. Else it uses English.
- The log, the confirmation windows and the error messages also use the selected language.
  The audit file uses fixed English values.

To add a language, put a new `.lng` file in the folder `Lang`. No code change is necessary.
See [docs/TRANSLATING.md](docs/TRANSLATING.md).

## Limits

- The tracking log keeps data for 30 days by default. The maximum period is 31 days.
- The tracking log does not show what occurs in the mailbox after the delivery. Use the [mailbox check](#mailbox-check).
- EMT does not delete unindexed items. The summary line shows their number.
- EMT does not find a forwarded copy with a changed subject.
- `Search-Mailbox` is not available in Exchange Online. EMT is for on-premises Exchange only.

## Troubleshooting

| Problem | Cause | Action |
|---|---|---|
| `Start Exchange Message Trail in STA mode` | PowerShell runs in MTA mode | Use `EMT.cmd` or the `-STA` parameter |
| Kerberos error on connect | The computer is not in the domain, or the server name is not an FQDN | Use the FQDN. Use a domain computer. |
| A button is disabled | Your account does not have the role | See [Permissions](#permissions) |
| The trace finds nothing | Wrong period, wrong address, or the log data is too old | Read the hints in the log. Make the period larger. Make sure that the date is in the last 30 days. |
| Delivered, but the user does not see the message | A rule, the junk e-mail filter or forwarding | Use the [mailbox check](#mailbox-check) |
| Found 0 copies, but the message exists | The message is unindexed, or the subject is different | Use **Search all mailboxes**. Make the subject text shorter. |
| The search of copies is slow | Many mailboxes: EMT uses Compliance Search | Examine the stage times in the log. For one mailbox, use **Recipient mailbox only**. |
| The same count after a soft delete | Compliance Search reads Recoverable Items | Use the tab **Verification**, not a new search |
| A search stays after **Clean up old searches** | The search still runs | Wait for the end of the search. Do the cleanup again. |

### Exchange 2019 behavior

These facts come from tests on Exchange Server 2019. EMT obeys them. Obey them also in your own scripts.

| Behavior | Effect | Rule |
|---|---|---|
| A `*-ComplianceSearch*` cmdlet with `-ErrorAction Stop` | The cmdlet fails with "Object reference not set to an instance of an object". `Remove-ComplianceSearch` stops in the middle, and the search stays. | Do not use `-ErrorAction Stop` with these cmdlets. Use `-ErrorVariable`. |
| `Get-ComplianceSearchAction` without `-Identity` | The full list fails with "Object reference not set to an instance of an object" when one action is damaged | Ask for one action: `-Identity <search name>_Purge` |
| `Remove-ComplianceSearch` | Exchange removes the search in the background. The search can stay in the list for some seconds. | After the remove, examine the search again for up to 60 seconds |
| `Remove-ComplianceSearch` and the shadow search | The cmdlet also removes the shadow search `<search name>-shadow` | Remove the shadow search first. Then remove the search. |
| `Stop-ComplianceSearch` | The search does not stop. It continues until it is complete (1–3 minutes). | Do not remove a running search. Wait for the status `Completed`. |
| A running search | `Remove-ComplianceSearch` cannot remove it | Examine the status before the remove |
| The text "couldn’t be found" | The message has a typographic apostrophe (U+2019), not `'` | Accept any character at this position: `couldn.t be found` |
| The Exchange admin center | **In-place eDiscovery & hold** shows only the shadow searches. It does not show the Compliance Searches. | Use `Get-ComplianceSearch` to see the Compliance Searches |

## More documents

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): the structure of the code.
- [docs/TESTING.md](docs/TESTING.md): the unit tests and the test procedure on Exchange.
- [docs/TRANSLATING.md](docs/TRANSLATING.md): how to add a language.
- [docs/TODO.md](docs/TODO.md): possible future work.
- [CHANGELOG.md](CHANGELOG.md): the list of changes.
