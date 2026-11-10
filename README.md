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

The user interface is in Russian. The documentation is in English.

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
2. Keep the folder structure. `MailCleaner.ps1` needs `src\MailCleaner.Core.ps1`.
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
6. Click **Найти**.
7. Wait for the end of the search. The progress bar shows the status.
8. Examine the tabs (see [Results](#results)).
9. Make sure that the result contains only the necessary message.
10. Click **Удалить (мягко)** or **Удалить безвозвратно**.
11. Read the summary in the confirmation window.
12. Confirm the delete.
13. Examine the tab **Проверка**.

To stop an operation, click **Отмена**.
MailCleaner stops after the current step.

> **WARNING:** A hard delete is permanent. The user cannot recover the message.
> Do a soft delete first if you are not sure.

> **CAUTION:** If you change a criterion after a search, MailCleaner clears the result.
> Do the search again before you delete.

## Search criteria

Type one or more of these criteria. A query with only a period is not permitted.

| Field | Use | Message tracking | Compliance Search (KQL) |
|---|---|---|---|
| Отправители | One or more sender addresses | `-Sender` | `from:` (OR) |
| Тема содержит | Part of the subject | `-MessageSubject` | `subject:"…"` |
| Получатель | One recipient address | `-Recipients` | not used |
| Начало / Конец периода | Date and time | exact time | full days only |
| MessageId | Internet message ID | `-MessageId` | subject and sender of the message |
| Имя вложения | Attachment file name | not used | `attachment:"…"` |
| Слова в тексте | A phrase in the message | not used | `"…"` |

| Check box | Function |
|---|---|
| Включать пересылки (FW:) | Adds the senders of forwarded copies to the KQL query |
| Искать во всех ящиках | Searches all mailboxes, not only the recipient mailboxes |

**How MailCleaner finds the message:**

1. It sends one tracking query for each sender.
2. It sends one tracking query for the subject.
3. It keeps the messages with the subject text.
4. It reads all events of these messages (delivery, expansion, send).

A wrong sender or a wrong time does not stop the search.
The subject query finds the message.

**Period:** Compliance Search uses full days only.
An end time of 00:00 does not include the next day.

If the tracking log does not find internal recipients, MailCleaner searches all mailboxes.

## Results

| Tab | Content |
|---|---|
| Письма | One row for each message. A yellow row is a message that Exchange made (an inbox rule or forwarding). |
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
| Confirmation | Summary window | Summary window and the word `УДАЛИТЬ` |

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

**Holds:** A mailbox on Litigation Hold or In-Place Hold keeps a copy after a hard delete.
Exchange keeps this copy in a hidden folder. The user cannot see this copy.

## Files

MailCleaner writes these files:

| File | Content |
|---|---|
| `%USERPROFILE%\MailCleaner\audit.csv` | One row for each search, delete, verification and cleanup |
| `%USERPROFILE%\MailCleaner\logs\MailCleaner_<date>.log` | The full log of one session |
| `%APPDATA%\MailCleaner\settings.json` | The last server, senders, subject and options |

Audit columns: `Timestamp`, `Operator`, `Computer`, `Server`, `Action`, `SearchName`,
`Query`, `Mailboxes`, `Items`, `Result`, `Details`.

| `Action` | `Result` values |
|---|---|
| Search | `OK`, `Errors` |
| SoftDelete, HardDelete, Verify | `Clean`, `RecoverableOnly`, `Visible`, `CheckFailed` |
| CleanupSearches | `OK` |

**Export:** Click **Экспорт в CSV** to save all tabs as CSV files.

**Cleanup:** Each search makes a Compliance Search with the name `MC_<date>_<time>`.
Click **Очистить старые поиски** to delete old `MC_*` searches.
MailCleaner does not delete searches with other names.

> **CAUTION:** The log, audit and export files contain mailbox addresses.
> Keep them in a safe location.

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
| The same count after a soft delete | Compliance Search reads Recoverable Items | Use the tab **Проверка**, not a new search |

## More documents

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): the structure of the code.
- [docs/TESTING.md](docs/TESTING.md): the unit tests and the test procedure on Exchange.
- [CHANGELOG.md](CHANGELOG.md): the list of changes.
