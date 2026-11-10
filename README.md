# MailCleaner

MailCleaner is a Windows GUI for Exchange Server 2016 and 2019 (on-premises).
Use it to find a misdirected message, see its recipients, and delete it from mailboxes.

The tool does these steps:

1. It finds the message in the message tracking log.
2. It finds the forwarded copies (`FW:`) of the message.
3. It finds the internal and the external recipients.
4. It searches the mailboxes with a Compliance Search.
5. It deletes the message with a soft delete or a hard delete.
6. It verifies each mailbox after the delete.
7. It writes each action to an audit file.

The user interface is available in English and Russian. You can add a language: see [Languages](#languages).
The documentation is in English.

## Contents

- [Requirements](#requirements)
- [Permissions](#permissions)
- [Install](#install)
- [Start](#start)
- [Procedure](#procedure)
- [Search criteria](#search-criteria)
- [Results](#results)
- [Delete types](#delete-types)
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
MailCleaner connects to Exchange through remote PowerShell.

## Permissions

MailCleaner checks the permissions of your account when it connects.
A button stays disabled if your account does not have the necessary role.

| Function | Cmdlets | Role group |
|---|---|---|
| Message tracking | `Get-MessageTrackingLog`, `Get-TransportService` | Organization Management, Records Management or Compliance Management |
| Search | `New-ComplianceSearch`, `Start-ComplianceSearch`, `Get-ComplianceSearch` | Discovery Management |
| Accepted domains | `Get-AcceptedDomain` (optional, for speed) | Organization Management or View-Only Organization Management |
| Cleanup of shadow searches | `Get-MailboxSearch`, `Remove-MailboxSearch` | Discovery Management |
| Verify | `Search-Mailbox` | Discovery Management |
| Soft delete | `New-ComplianceSearchAction` | Discovery Management |
| Hard delete | `Search-Mailbox -DeleteContent` | Mailbox Import Export role (assign it manually) |

To assign the Mailbox Import Export role, do this procedure in the Exchange Management Shell:

```powershell
New-ManagementRoleAssignment -Role "Mailbox Import Export" -User "<account>"
```

Close MailCleaner and start it again after a role change.

## Install

1. Copy the repository folder to the computer.
2. Keep the folder structure. `MailCleaner.ps1` needs the folders `src` and `Lang`.
3. If you downloaded a ZIP file, unblock the files:

```powershell
Get-ChildItem -Recurse | Unblock-File
```

## Start

Use one of these procedures.

**Procedure A: the launcher**

1. Double-click `MailCleaner.cmd`.

**Procedure B: PowerShell**

1. Open Windows PowerShell.
2. Go to the MailCleaner folder.
3. Type this command:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\MailCleaner.ps1 -Server exch01.contoso.local
```

The `-Server` parameter is optional. MailCleaner keeps the last server name.
On an Exchange server, the default value is the name of that server.

## Procedure

1. In the field **Сервер Exchange (FQDN)**, type the full server name.
2. Click **Подключиться**.
3. Make sure that the status line shows the roles (`✓`).
4. Type the search criteria (see [Search criteria](#search-criteria)).
5. Set the period. Use **Сегодня** or **Вчера** for a full day.
6. Click **Найти**, or push Enter.
7. Wait for the end of the search. The progress bar shows the status.
8. Examine the tabs (see [Results](#results)).
9. Make sure that the result contains only the necessary message.
10. Click **Удалить (мягко)** or **Удалить безвозвратно**.
11. Read the summary in the confirmation window.
12. Confirm the delete.
13. Examine the tab **Проверка**.

To stop an operation, click **Отмена** (bottom right).
MailCleaner stops after the current step.

**Window layout:**

| Area | Controls |
|---|---|
| Top line | Server, Connect, operation timeout, cleanup of old searches, language, roles |
| Критерии поиска | Senders, subject, recipient, period with **Сегодня** and **Вчера**, message ID, attachment, words, options, **Найти** |
| Summary line | Counts, method, time. A red line shows errors: the delete is not available. |
| Tabs | The results. The tab name shows the row count. |
| Log | The messages of the session |
| Bottom left | Actions on the result: soft delete, hard delete, verify, export |
| Bottom right | The running operation: status, progress, cancel |

**Tables:** Drag a column border to change the width. Click a column header to sort.
Drag a column header to move the column. Push Ctrl+C to copy the selected rows with headers.

**Button states after a delete:**

| Last delete | Available |
|---|---|
| None | Soft delete, hard delete, verify |
| Soft delete | Hard delete, verify |
| Hard delete | Verify |

A new search resets the state.

> **WARNING:** A hard delete is permanent. The user cannot recover the message.
> Do a soft delete first if you are not sure.

> **CAUTION:** If you change a criterion after a search, MailCleaner clears the result.
> Do the search again before you delete.

## Search criteria

Type one or more of these criteria. A query with only a period is not permitted.

| Field | Use | Message tracking | Compliance Search (KQL) |
|---|---|---|---|
| Отправители | One or more sender addresses, parts of addresses or names | `-Sender` | `from:` (OR) |
| Тема содержит | Part of the subject | `-MessageSubject` | `subject:"…"` |
| Получатель | One recipient address, a part of an address or a name | `-Recipients` | not used (see **Только ящик получателя**) |
| Начало / Конец периода | Date and time | exact time | full days only |
| MessageId | Internet message ID | `-MessageId` | subject and sender of the message |
| Имя вложения | Attachment file name | not used | `attachment:"…"` |
| Слова в тексте | A phrase in the message | not used | `"…"` |

| Check box | Function |
|---|---|
| Включать пересылки (FW:) | Adds the senders of forwarded copies to the KQL query |
| Только ящик получателя | Searches and deletes the message only in the mailbox in **Получатель** |
| Искать во всех ящиках | Searches all mailboxes, not only the recipient mailboxes |

The last two check boxes exclude each other.

**Partial addresses and names:** In **Отправители** and **Получатель**, you can type a part of an address
(`jsmi`) or a name (`Ivanov Ivan`). When you click **Найти**, MailCleaner finds the full address
in the Exchange address book (`Get-Recipient -Anr`). ANR compares the text with the alias, the display name,
the first name, the last name and the start of the e-mail addresses.

| Matches | Result |
|---|---|
| One | MailCleaner puts the full address in the field. The log shows the change. |
| More than one | A window shows the list. Select the address. For senders, you can select more than one (Ctrl+click). |
| None | An error window shows the text. The search does not start. |

The tracking log and the KQL query always get full addresses. MailCleaner does not use wildcards:
a wildcard can find messages of other users.

**Separators:** Separate the senders with a comma or a semicolon.
A space separates only full addresses. A name with a space stays one item.

**Forwarded copy:** a message from another sender with the same base subject.
The base subject is the subject without `RE:`, `FW:`, `Fwd:`, `AW:`, `WG:`, `Отв:`, `Пересл:` and similar prefixes.
A message with a different subject that only contains the text is not a forwarded copy.
If the tracking log has no message from the senders, MailCleaner uses all messages with the subject text.
The log shows a warning. Examine the tab **Письма**.

**Subject text:** MailCleaner reads the subject as text. Square brackets, `*` and `?` have no special meaning.

**Получатель without "Только ящик получателя":** MailCleaner finds the messages that this address received.
Then it searches all recipient mailboxes of these messages.
Use this mode to delete the message from all recipients.

**Получатель with "Только ящик получателя":** MailCleaner searches only this mailbox.
The delete affects only this mailbox. This is the fastest mode.

**How MailCleaner finds the message:**

1. It sends one tracking query for each sender.
2. It sends one tracking query for the subject.
3. It keeps the messages with the subject text.

All events of one message have the same sender and the same subject.
Thus these queries return all events of the message: receive, expansion, delivery and send.
With a recipient filter, the tracking log returns only the events of that recipient.
In this case, MailCleaner reads all events of each message again by message ID.

A wrong sender or a wrong time does not stop the search.
The subject query finds the message.

**Period:** Compliance Search uses full days only.
An end time of 00:00 does not include the next day.
Exchange can read the KQL dates as local time or as UTC.
For this reason, the KQL range covers the local days and the UTC days.
The range can be one day larger than the period. The tracking log uses the exact time.

**Tracking errors:** if a transport server does not answer, the mailbox list can be incomplete.
The summary line is red and the delete is not available.
Select **Искать во всех ящиках** and do the search again.

If the tracking log does not find internal recipients, MailCleaner searches all mailboxes.

## Search speed

MailCleaner selects the search method from the number of mailboxes:

| Mailboxes | Method | Usual time |
|---|---|---|
| 1 to 10 | `Search-Mailbox -EstimateResultOnly` for each mailbox | Seconds |
| More than 10, or all mailboxes | Compliance Search | 1 to 3 minutes, or more for all mailboxes |

The quick method does not give the number of unindexed items. The summary line shows `н/д`.
A soft delete after a quick search makes a Compliance Search first. This adds 1 to 3 minutes.

**Other speed measures:**

- MailCleaner reads the tracking log once. A second pass occurs only with a recipient filter.
- MailCleaner does not look up addresses outside the accepted domains. These addresses are external.
- MailCleaner looks up the other addresses in batches of 20 addresses.

The log shows the time of each stage. The summary line shows the method and the full time.

## Results

| Tab | Content |
|---|---|
| Письма | One row for each message. A yellow row is a message that Exchange made (an inbox rule or forwarding). Events without a message ID (for example `NOTIFYMAPI`) are not messages: they are only in the tab **Журнал доставки**. |
| Журнал доставки | All tracking events |
| Получатели | Internal mailboxes and external addresses. A red row is an external address. |
| Найдено в ящиках | Mailboxes with the message: count and size |
| Не найдено | Searched mailboxes without the message |
| Проверка | The result of the last verification |

The summary line shows these values:

- the number of messages and mailboxes;
- the number of mailboxes without the message;
- the number of external recipients;
- the number of unindexed items.

**External recipients:** MailCleaner cannot delete a message from an external mailbox.
Tell your security team about each external recipient.

**Yellow row in "Письма":** Exchange made this message, not a user.
An inbox rule or mailbox forwarding can send messages out of the organization.
Examine the forwarding settings of that mailbox:

```powershell
Get-InboxRule -Mailbox <user> | Where-Object { $_.ForwardTo -or $_.RedirectTo -or $_.ForwardAsAttachmentTo }
Get-Mailbox <user> | Format-List ForwardingSmtpAddress, DeliverToMailboxAndForward
```

## Delete types

| | Soft delete | Hard delete |
|---|---|---|
| Button | Удалить (мягко) | Удалить безвозвратно |
| Cmdlet | `New-ComplianceSearchAction -Purge -PurgeType SoftDelete` | `Search-Mailbox -SearchDumpster -DeleteContent` |
| Result | The message moves to Recoverable Items | The message is deleted from all folders and from Recoverable Items |
| User can recover | Yes, with "Recover Deleted Items" in Outlook | No |
| Confirmation | Summary window. The word `УДАЛИТЬ` for a large delete. | Summary window and the word `УДАЛИТЬ` |
| Safety check before the delete | — | `Search-Mailbox` must not find more items than the search |

**Summary window:** it shows the counts, the KQL query, the messages from the tracking log
(time, sender, subject) and the mailboxes.

**Large delete:** more than 100 items, or more than 5 different messages in the tracking log.
The summary window shows a warning. A soft delete then also needs the word `УДАЛИТЬ`.

**Safety check before a hard delete:** MailCleaner counts the items in each mailbox with `Search-Mailbox`.
If a count is more than the search count, MailCleaner stops the delete.
A larger count can occur when new messages arrived, or when `Search-Mailbox` reads the query differently.
Do the search again.

**Verification:** A Compliance Search also searches Recoverable Items.
After a soft delete, a new Compliance Search shows the same count.
For this reason, MailCleaner verifies with `Search-Mailbox`.
It gives two counts for each mailbox:

| Column | Meaning |
|---|---|
| В видимых папках | Messages that the user can see |
| Всего | Messages in all folders, including Recoverable Items |

| Row color | Meaning |
|---|---|
| Green | The mailbox does not contain the message |
| Red | The user can see the message |
| Yellow | The verification failed for this mailbox |

The search index can show a deleted message for some seconds.
If the user can still see the message, MailCleaner waits 15 seconds and examines the mailbox again (3 times maximum).

**Holds:** A mailbox on Litigation Hold or In-Place Hold keeps a copy after a hard delete.
Exchange keeps this copy in a hidden folder. The user cannot see this copy.

## Files

MailCleaner writes these files:

| File | Content |
|---|---|
| `%USERPROFILE%\MailCleaner\audit.csv` | One row for each search, delete, verification and cleanup |
| `%USERPROFILE%\MailCleaner\logs\MailCleaner_<date>.log` | The full log of one session |
| `%APPDATA%\MailCleaner\settings.json` | The last server, the language, the options and the timeout. Senders and subject are not kept. |

Audit columns: `Timestamp`, `Operator`, `Computer`, `Server`, `Action`, `SearchName`,
`Query`, `Mailboxes`, `Items`, `Result`, `Details`.

| `Action` | `Result` values |
|---|---|
| Search | `OK`, `Errors` |
| SoftDelete, HardDelete, Verify | `Clean`, `RecoverableOnly`, `Visible`, `CheckFailed` |
| SoftDelete, HardDelete (stopped) | `Cancelled`, `Failed`. `Details` gives the cause and, for a hard delete, the number of deleted items. |
| CleanupSearches | `OK` |

**Export:** Click **Экспорт в CSV** to save all tabs as CSV files.
The file `*-errors.csv` contains the tracking errors and the search errors.

**Cleanup:** Each search makes a Compliance Search with the name `MC_<date>_<time>`.
For each Compliance Search, Exchange also makes a shadow In-Place eDiscovery search
`MC_<date>_<time>-shadow`. The Exchange admin center shows it in **compliance management** >
**in-place eDiscovery & hold**. Exchange can keep it after the Compliance Search is removed.

Click **Очистить старые поиски** to remove old searches. MailCleaner removes:

- the Compliance Searches `MC_<date>_<time>` and their actions;
- the shadow searches `MC_<date>_<time>-shadow`.

The age comes from the date and time in the name. To remove the searches of today, type `0` days.
MailCleaner does not remove searches with other names.
The shadow searches need the cmdlets `Get-MailboxSearch` and `Remove-MailboxSearch` (Discovery Management).

> **CAUTION:** The log, audit and export files contain mailbox addresses.
> Keep them in a safe location.

## Languages

The languages are files in the folder `Lang`: `en.lng` (English) and `ru.lng` (Russian).

- Select the language in the list **Language** in the top line. The change occurs immediately.
- MailCleaner keeps the selection in `settings.json`.
- At the first start, MailCleaner uses the Windows language if a file for it exists. Else it uses English.
- To start with a specified language, use the parameter `-Language`: `.\MailCleaner.ps1 -Language en`.
- The log, the confirmation windows and the error messages also use the selected language.
  The audit file uses fixed English values.

To add a language, put a new `.lng` file in the folder `Lang`. No code change is necessary.
See [docs/TRANSLATING.md](docs/TRANSLATING.md).

## Limits

- MailCleaner does not delete unindexed items. The summary line shows their number.
- MailCleaner does not find a forwarded copy with a changed subject.
- The message tracking log keeps data for 30 days by default. The maximum period is 31 days.
- `Search-Mailbox` is deprecated in Exchange Online. MailCleaner is for on-premises Exchange only.

## Troubleshooting

| Problem | Cause | Action |
|---|---|---|
| `Start MailCleaner in STA mode` | PowerShell runs in MTA mode | Use `MailCleaner.cmd` or the `-STA` parameter |
| Kerberos error on connect | The computer is not in the domain, or the server name is not an FQDN | Use the FQDN. Use a domain computer. |
| A button is disabled | Your account does not have the role | See [Permissions](#permissions) |
| The tracking log is empty | Wrong period, or the log data is too old | Set a full day. Make sure that the date is in the last 30 days. |
| Found 0, but the message exists | The message is unindexed, or the subject is different | Use **Искать во всех ящиках**. Make the subject text shorter. |
| The search is slow | Many mailboxes: MailCleaner uses Compliance Search | Examine the stage times in the log. For one mailbox, use **Только ящик получателя**. |
| The same count after a soft delete | Compliance Search reads Recoverable Items | Use the tab **Проверка**, not a new search |

## More documents

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): the structure of the code.
- [docs/TESTING.md](docs/TESTING.md): the unit tests and the test procedure on Exchange.
- [docs/TRANSLATING.md](docs/TRANSLATING.md): how to add a language.
- [CHANGELOG.md](CHANGELOG.md): the list of changes.
