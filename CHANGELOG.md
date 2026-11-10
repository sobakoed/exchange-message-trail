# Changelog

## 1.0.2 — 2026-11-10

Window audit.

**Criteria**

- The criteria have two columns "label + field" on both tabs. All text fields have the same width.
  The fields have the same width and position on the tabs **Message search** and **Delete**, also after a language change.
- The same order on both tabs: senders and recipient, subject and message ID, then the period.
- One line **Period**: start — end, **Today**, **Yesterday**, **7 days**. The tab **Delete** also has **7 days** now.
  The labels "Period start" and "Period end" are removed.
- At the minimum window size, all buttons are fully visible. Before, **7 days** was cut.
- **Find** and **Find copies** are at the right edge of the fields, in the last line of the criteria.
- The criteria area gets its height from the content.

**Connection**

- The connection state moved from the top line to the bottom line: a green indicator and the server name,
  or a grey indicator and "Not connected".
- The window does not show the roles. The log shows one warning for each function that is not available.
  New warnings: search, verification, soft delete.
- A long status text of an operation is cut. The tooltip shows the full text.

**Start and close**

- The window opens in front of the other windows.
- `EMT.cmd` starts PowerShell with a hidden console. EMT also hides a console that it does not share
  with other programs. A console from which the user started EMT stays visible.
- An error at the start shows a window with the cause. The console is hidden and cannot show it.
- The window size, the position and the maximized state are kept in `settings.json`.
  EMT uses them only if the window is on a screen.

**Russian texts**

- The tab "Трассировка" is now "Поиск писем". The window title, the status and the log use the same name.

**English texts**

- The tab **Trace** is now **Message search**. The window title, the status and the log use the same name.
  The audit value `Trace` does not change.

## 1.0.1 — 2026-11-10

Window.

- The row height and the header height of the tables are fixed. The user can change only the column width.
- The buttons **Find** (tab **Trace**) and **Find copies** (tab **Delete**) have a light green background.
  A disabled button has the usual grey background.

## 1.0.0 — 2026-11-10

Trace first, delete when necessary. New name: Exchange Message Trail (EMT).

**New name**

- The program is now Exchange Message Trail (EMT). The older name was MailCleaner.
- New file names: `ExchangeMessageTrail.ps1`, `EMT.cmd`, `src/EMT.Core.ps1`, `src/EMT.Strings.ps1`,
  `tests/EMT.Core.Tests.ps1`.
- New prefix of the functions: `Emt` (for example `Invoke-EmtSearch`). The older prefix was `Mc`.
- New prefix of the searches on the server: `EMT_yyyyMMdd_HHmmss`. The cleanup also removes the older `MC_*` searches.
- New data folders: `%USERPROFILE%\ExchangeMessageTrail` and `%APPDATA%\ExchangeMessageTrail`.
  At the first start, EMT copies `audit.csv` and `settings.json` from the MailCleaner folders. The log shows each copy.
- The function for translators is now `Test-EmtLanguageFile`.

**Window**

- Two tabs in one window: **Trace** and **Delete**. The connection line, the log, the progress and **Cancel** are common.
- The tab **Trace** opens at the start. The tab **Delete** is always available.
  Its buttons are disabled without the necessary roles.
- Enter starts the operation of the open tab.

**Trace**

- New operation: find messages in the tracking log only. It needs only the Message Tracking role.
  A trace takes some seconds. It does not make a Compliance Search.
- Criteria: senders, recipient, subject, message ID, period. A recipient alone is a usable criterion.
- Buttons for the period: **Today**, **Yesterday**, **7 days**.
- The message table shows a delivery status for each message, for example "2 delivered, 1 failed".
  Messages with an error are red. Messages in transit or deferred are yellow.
- The tab **Recipients** shows the status of each recipient: Delivered, Sent to an external system, Delivery failed,
  Dropped, Deferred, In transit, Group, No delivery data. Each status has a color.
- The column **Explanation** tells the cause of an error from the enhanced status code
  (for example `5.1.1`: the address does not exist).
- The tab **Events** shows all tracking events of the message in time order.
- No result: the log shows hints (period, 30 days of the tracking log, sender, recipient, subject, Sent Items).
- Export of the trace to CSV.

**Mailbox check**

- New button **Check the recipient mailbox** for a delivered message and an internal recipient.
- The check reads the inbox rules, the junk e-mail settings, the mailbox forwarding and the quota. It changes nothing.
- The findings tell if a rule moves, deletes or forwards the message, if the sender is blocked,
  if the mailbox forwards the messages without a copy, and if the mailbox is almost full.
- A check without permission gives the finding "not examined: no permission".
- The tab **Recipient mailbox** shows the findings. Important findings are red, cautions are yellow.

**Delete**

- New button **Delete this message…** on the tab **Trace**. It opens the tab **Delete** and fills the criteria
  from the selected message: sender, subject, message ID, recipients, period.
- A source line on the tab **Delete**: "Data from the tracking log" (green) or "Data typed by the user" (orange).
  A change of a criterion clears the result and sets the orange line.
- The search button on the tab **Delete** is now **Find copies**.

**Audit**

- New actions in `audit.csv`: `Trace` and `MailboxCheck`.

**Texts**

- 106 new keys in `Lang/en.lng` and `Lang/ru.lng`.
- After a language change, the window makes all texts again: statuses, explanations, status lines and findings.

**Documentation**

- README: the work order (trace, mailbox check, delete), the roles for the help desk, the statuses and the codes.
- New document `docs/TODO.md`: possible future work.
- `docs/ARCHITECTURE.md`, `docs/TESTING.md` and `docs/TRANSLATING.md` are updated for EMT.

## 0.6.0 — 2026-11-10

Recipient and safety.

- The recipient is now in the KQL query: `(to:x OR cc:x OR bcc:x)`. Before, `from:sender` with a recipient
  found all messages of the sender to all people.
- Tracking log without messages and a recipient in the criteria: the search uses only the mailboxes
  of the recipient and the senders, not all mailboxes.
- A result that the tracking log does not confirm (items found, no message in the tracking log):
  red summary line, a warning in the confirmation window, and the typed word for each delete.
- Quick search with no found items no longer stops with an error (`Measure-Object` with an empty list).
- The prefixes of the summary line ("ERRORS", "NOT CONFIRMED") have a space before the text.
- No message in any mailbox: the tab "Not found" stays empty, and the summary line starts with "No messages found".
- Cleanup: the remove worked, but the check said "still there". The Exchange message "couldn’t be found" can
  have a typographic apostrophe (U+2019). The check now accepts any character at this place.
- Cleanup: Exchange 2019 removes a Compliance Search in the background. It can stay in the list for more than
  10 seconds. MailCleaner now waits up to 60 seconds (a check each 5 seconds) and shows the wait in the progress line.
  If `Remove-ComplianceSearch` returned no error, a search that is still visible is a warning with its status, not an error.
- Cleanup, second root cause: `Get-ComplianceSearch -Identity ... -ErrorAction Stop` also fails with a wrong error
  instead of "not found". Thus the check said "still there" for a removed search.
  No `*-ComplianceSearch*` cmdlet uses `-ErrorAction Stop` now. The errors come in `-ErrorVariable`.
  A unit test examines the source code for this rule.

## 0.5.2 — 2026-11-10

- Texts: the value 0 now replaces `{0}` (for example "Messages in the tracking log: 0"). Before, the text showed `{0}`.
- Cleanup: after an error of `Remove-ComplianceSearch`, Exchange can remove the search some seconds later.
  MailCleaner examines the list of searches up to 5 times with 2 seconds between the checks.
  Only a search that is still there after 10 seconds is an error.
- Cleanup: MailCleaner does not list all search actions any more (`Get-ComplianceSearchAction` without `-Identity`).
  On some servers, one damaged action makes the full list fail with "Object reference not set", and the cleanup failed.
  MailCleaner now asks only for its own action `<search name>_Purge`.
- Cleanup: tested on Exchange 2019. If the shadow search exists, `Remove-ComplianceSearch` removes the shadow search,
  fails with "Object reference not set" and keeps the Compliance Search. A second remove works.
  MailCleaner now removes the shadow search first, then the Compliance Search. After an error it tries once more.
- Cleanup: the check "does the search exist" uses `Get-ComplianceSearch -Identity`, not the full list.
- Cleanup, the root cause (tested on Exchange 2019): `Remove-ComplianceSearch -ErrorAction Stop` fails with
  "Object reference not set" and the search stays. Without `-ErrorAction Stop`, the same command removes the search.
  MailCleaner no longer uses `-ErrorAction Stop` for `Remove-ComplianceSearch` and `Remove-ComplianceSearchAction`.
  The result comes from the check by `-Identity`.
- Cleanup: a search that is still running is not touched (also not its shadow search). The log tells to remove it later.
- Cancel: `Stop-ComplianceSearch` does not always stop a search on Exchange on-premises. The log tells this.

## 0.5.1 — 2026-11-10

- Cleanup: `Remove-ComplianceSearch` also removes the shadow search. MailCleaner now examines a shadow search
  before the remove. A shadow search that is already gone is not an error.
- Cleanup: an error on one search does not stop the cleanup of the other searches. The log shows the error.
- Cleanup: Exchange on-premises can return an error ("Object reference not set…") and still remove the search.
  After an error, MailCleaner examines the list of searches. If the search is gone, the log shows a warning, not an error.
- Cleanup: an error when MailCleaner removes the actions of a search does not stop the removal of the search.

## 0.5.0 — 2026-11-10

Partial addresses.

- **Отправители** and **Получатель** accept a part of an address or a name.
  MailCleaner finds the full address in the Exchange address book (`Get-Recipient -Anr`) before the search.
- One match: the full address goes in the field. More matches: a selection window. No match: an error, no search.
- The tracking log and the KQL query get only full addresses. No wildcards.
- Separators of the senders: comma, semicolon, line break. A space separates only full addresses.

## 0.4.1 — 2026-11-10

- Tracking events without a message ID (for example `NOTIFYMAPI`) are not a row in the tab of messages.
  They stay in the tab of the tracking log.
- The cleanup of old searches also removes the shadow In-Place eDiscovery searches `MC_*-shadow`.
- The cleanup removes only the names `MC_yyyyMMdd_HHmmss` and `MC_yyyyMMdd_HHmmss-shadow`. The age comes from the name.

## 0.4.0 — 2026-11-10

Languages.

- The user interface is available in English and Russian.
- The texts are in the files `Lang/*.lng`. To add a language, put a new file in this folder. No code change is necessary.
- The language list in the top line. The change occurs immediately. The selection is kept in `settings.json`.
- At the first start, the Windows language is used if a file for it exists. Else English.
- New parameter `-Language`.
- The log, the progress, the confirmation windows and the error messages use the selected language.
- A missing key in a translation uses the English text.
- New check `Test-McLanguageFile` for translators. New document `docs/TRANSLATING.md`.
- Buttons get wider for a long translation.
- Unit tests: complete bundled files, no unknown keys or wrong placeholders in any file, all keys of the code in `en.lng`.

## 0.3.0 — 2026-11-10

Code audit and window audit.

**Correctness and safety**

- The subject is text in the tracking filter. Square brackets, `*` and `?` have no special meaning.
- A forwarded copy must have the same base subject. A message with a different subject is not added.
- The KQL date range covers local time and UTC.
- Tracking errors of a transport server go to the result. The delete is not available until all mailboxes are searched.
- Safety check before a hard delete: `Search-Mailbox` must not find more items than the search.
- A stopped delete writes an audit row with `Cancelled` or `Failed`. A cancelled hard delete gives the partial counts.
- Verification after a delete examines the mailboxes again when the index lags.
- The confirmation window shows the messages. A large delete shows a warning and needs the word `УДАЛИТЬ`.
- Senders and subject are not kept in `settings.json`.

**Window**

- Table columns: the user can change the width and the order. A sort keeps the widths. Ctrl+C copies rows.
- Initial column widths from the content, also for hidden tabs.
- The tab name shows the row count.
- New layout: timeout and cleanup in the top line; **Сегодня** and **Вчера** next to the end date;
  actions bottom left; status, progress and **Отмена** bottom right.
- The criteria can be typed before the connection. Enter starts the search.
- **Только ящик получателя** and **Искать во всех ящиках** exclude each other.
- Button states after a delete. The Connect button shows **Переподключиться**.
- The summary line is red with errors. Its tooltip shows the full text.
- The search button needs only `Search-Mailbox`. The soft delete needs the Compliance Search role.
- Export: a file name for the quick search, and a new file `*-errors.csv`.
- Closing the window during an operation does not freeze it.

## 0.2.0 — 2026-11-10

Search speed.

- New check box **Только ящик получателя**: the search and the delete use only the recipient mailbox.
- Quick search with `Search-Mailbox` for 1 to 10 mailboxes. A Compliance Search is made only for a soft delete.
- The tracking log is read once. The second pass by message ID occurs only with a recipient filter.
- Duplicate tracking events are removed.
- Addresses outside the accepted domains are external without a lookup.
- Address lookup in batches of 20 (`Get-Recipient -Filter`). Aliases match.
- The log shows the time of each stage. The summary line shows the method and the full time.
- The status poll interval is 3 seconds (was 5 seconds).

## 0.1.0 — 2026-11-10

First version.

- Windows Forms GUI for Exchange Server 2016 and 2019 (on-premises).
- Background runspace with its own Exchange remote session. The window does not freeze.
- Message tracking with separate queries for the sender and for the subject.
- Forwarded copies (`FW:`) from other senders.
- Mark for messages that Exchange made (inbox rules, mailbox forwarding).
- Internal and external recipients.
- Compliance Search on the recipient mailboxes, or on all mailboxes.
- Soft delete (`New-ComplianceSearchAction -Purge -PurgeType SoftDelete`).
- Hard delete (`Search-Mailbox -DeleteContent`) with a typed confirmation.
- Verification with `Search-Mailbox`: visible folders and Recoverable Items.
- Role check at connect. Buttons without a role stay disabled.
- Audit file, session log, CSV export, cleanup of old `MC_*` searches.
- Pester 5 unit tests for the core.
