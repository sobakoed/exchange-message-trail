# Testing

This document gives the unit tests and the test procedure on Exchange.

## Unit tests

The unit tests examine `src/EMT.Core.ps1`, `src/EMT.Strings.ps1` and the language files. They do not need Exchange.
Stub functions and Pester mocks replace the Exchange cmdlets.

**Requirements:** Pester 5.0 or later.

To install Pester 5, type this command:

```powershell
Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser -Force -SkipPublisherCheck
```

To run the tests, do these steps:

1. Open Windows PowerShell 5.1.
2. Go to the repository folder.
3. Type this command:

```powershell
Invoke-Pester -Path .\tests -Output Detailed
```

All tests must pass.

## Syntax check

To examine the syntax of all files, type this command:

```powershell
Get-ChildItem -Recurse -Filter *.ps1 | ForEach-Object {
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$null, [ref]$errors)
    if ($errors) { $errors } else { "OK  $($_.Name)" }
}
```

## Test procedure on Exchange

Do this procedure on test mailboxes before you use EMT in production.
The procedure uses the English interface. Select **English** in the language list.

### Preparation

1. Get three test mailboxes: `test1`, `test2`, `test3`.
2. Get one external address that you can read, for example `test@example.net`.
3. From `test1`, send a message to `test2`, `test3` and the external address. Use a unique subject.
4. From `test1`, send a second message to `nobody@<your domain>`. This address must not exist.
5. From `test2`, forward the first message to `test3`.
6. In the mailbox of `test3`, make an inbox rule: messages from `test1` go to a folder `Test`.
7. From `test1`, send a third message to `test3`. Use a new unique subject.
8. Wait 5 minutes for the index.

### Connection

1. Start EMT with `EMT.cmd`.
2. Type the server name.
3. Click **Connect**.
4. Make sure that the bottom line shows a green indicator and "Connected".
5. Make sure that the log shows no warning about a missing role.
6. Make sure that the tab **Message search** is open.

### Message search: delivered message

1. In **Senders**, type the address of `test1`.
2. In **Subject contains**, type the subject of the first message.
3. Click **Today**.
4. Click **Find**.
5. Make sure that the time in the summary line is a few seconds.
6. Make sure that the message table shows the message.
7. Select the message.
8. Make sure that the tab **Recipients** shows `test2` and `test3` with the status **Delivered** (green).
9. Make sure that the external address has the status **Sent to an external system** (blue).
10. Open the tab **Events**.
11. Make sure that the tab shows the events of the message in time order.

### Message search: failed message

1. In **Subject contains**, type the subject of the second message.
2. Click **Find**.
3. Select the message.
4. Make sure that the recipient row is red.
5. Make sure that **Explanation** tells that the address does not exist (code `5.1.1`).
6. Make sure that **Server answer** shows the text of the server.

### Message search: no result

1. In **Subject contains**, type a text that no message has.
2. Click **Find**.
3. Make sure that the summary line starts with "No messages found".
4. Make sure that the log shows the hints: period, sender, recipient, subject.
5. In **Period**, set the start to 35 days ago.
6. In **Period**, set the end to 34 days ago.
7. Click **Find**.
8. Make sure that the log shows the hint about the 30 days of the tracking log.

### Message search: recipient only

1. Clear **Senders**.
2. Clear **Subject contains**.
3. Click **Today**.
4. In **Recipient**, type the address of `test3`.
5. Click **Find**.
6. Make sure that the table shows all messages to `test3` for the period.

### Mailbox check

1. Find the third message.
2. Select the message.
3. On the tab **Recipients**, select the row of `test3`.
4. Click **Check the recipient mailbox**.
5. Make sure that the tab **Recipient mailbox: test3** opens.
6. Make sure that a finding tells about the rule and the folder `Test`.
7. Make sure that the finding tells that the rule applies to this message.
8. Select the first message.
9. On the tab **Recipients**, select the row of the external address.
10. Make sure that **Check the recipient mailbox** is disabled.

### Partial address

1. In **Senders**, type a part of the alias of `test1`.
2. Click **Find**.
3. Make sure that the field shows the full address of `test1`.
4. Make sure that the log shows the change.
5. In **Senders**, type a part of a name that more than one user has.
6. Click **Find**.
7. Make sure that the selection window shows the users.
8. In **Senders**, type a text that no user has.
9. Click **Find**.
10. Make sure that an error window shows the text.
11. Make sure that the message search does not start.

### Delete from the message search

1. On the tab **Message search**, find the first message.
2. Select the message.
3. Click **Delete this message…**.
4. Make sure that the tab **Delete** opens.
5. Make sure that the criteria contain the sender, the subject, the message ID and the period of the message.
6. Make sure that the source line is green and starts with "Data from the tracking log".
7. Click **Find copies**.
8. Make sure that the tab **Found in mailboxes** shows `test1`, `test2` and `test3`.
9. Change one letter in **Subject contains**.
10. Make sure that the result is cleared.
11. Make sure that the source line is orange and starts with "Data typed by the user".

### Search speed

1. On the tab **Message search**, select the first message again.
2. Click **Delete this message…**.
3. Click **Find copies**.
4. Make sure that the summary line shows the quick search. There are 3 mailboxes.
5. Make sure that the log shows the time of each stage.
6. In **Recipient**, type the address of `test3`.
7. Select **Recipient mailbox only**.
8. Click **Find copies**.
9. Make sure that the tab **Found in mailboxes** shows only `test3`.

### Soft delete

1. Clear **Recipient**.
2. Click **Find copies**.
3. Click **Delete (soft)**.
4. Confirm the delete.
5. Make sure that the log shows a new Compliance Search before the delete.
6. Make sure that the tab **Verification** shows `0` in **In visible folders** for each mailbox.
7. Make sure that **Total (with Recoverable Items)** is more than `0`. The messages are in Recoverable Items.

### Hard delete

1. Click **Find copies**.
2. Click **Delete permanently**.
3. Type `DELETE`.
4. Confirm the delete.
5. Make sure that the tab **Verification** shows `0` in both columns.
6. If a mailbox is on hold, make sure that the log shows the hold warning.

### Cleanup of old searches

1. Click **Clean up old searches**.
2. Make sure that the list contains the `EMT_*` searches of this test.
3. If the server has old `MC_*` searches, make sure that the list contains them.
4. Confirm the cleanup.
5. Make sure that the log shows "Shadow search … removed" and "Search … removed" for each search.
6. A removal can take up to 60 seconds. Make sure that the progress line shows the wait.
7. EMT does not remove a search that still runs. Make sure that the log tells you to remove it later.

### Window

1. Drag the border of a column in a table. Make sure that the width changes.
2. Click the column header. Make sure that the sort keeps the width.
3. On the tab **Delete**, select **Search all mailboxes**.
4. Select **Recipient mailbox only**. Make sure that **Search all mailboxes** is clear.
5. Select **Русский** in the language list. Make sure that all texts change, also in the tables.
6. Select **English** again.
7. Make the window as small as possible. Make sure that all buttons are fully visible.
8. Make sure that all text fields of the criteria have the same width.
9. Open the other tab. Make sure that the fields have the same width and position.
10. Drag the border of a table row. Make sure that the row height does not change.

### Start and close

1. Open another program, for example a browser, in full screen.
2. Start EMT with `EMT.cmd`.
3. Make sure that the EMT window is in front of the other program.
4. Make sure that there is no PowerShell console window and no console button on the taskbar.
5. Change the size and the position of the EMT window.
6. Close EMT.
7. Start EMT again. Make sure that the window has the same size and position.
8. Maximize the window. Close EMT. Start EMT again. Make sure that the window is maximized.

### Cancel

1. On the tab **Delete**, select **Search all mailboxes**.
2. Click **Find copies**.
3. Click **Cancel** during the search.
4. Make sure that the log shows the cancel.
5. Make sure that the window operates.

### Audit

1. Open `%USERPROFILE%\ExchangeMessageTrail\audit.csv`.
2. Make sure that there is one row for each message search, mailbox check, search of copies, delete and verification.

### Upgrade from MailCleaner

Do this test only on a computer with the older version (MailCleaner).

1. Make sure that `%USERPROFILE%\MailCleaner\audit.csv` exists.
2. Start EMT.
3. Make sure that `%USERPROFILE%\ExchangeMessageTrail\audit.csv` exists.
4. Make sure that the log shows "A file of the older version was copied".

### Permissions

1. Start EMT with an account that has only the Message Tracking role.
2. Connect to the server.
3. Make sure that **Find** on the tab **Message search** operates.
4. Make sure that **Find copies** is disabled.
5. Make sure that **Delete this message…** is disabled.
6. Make sure that the log shows a warning for each function that is not available:
   the mailbox check, the search of copies, the verification, the soft delete and the hard delete.
7. Close EMT.
8. Start EMT with an account without the Mailbox Import Export role.
9. Connect to the server.
10. Make sure that **Delete permanently** is disabled.
11. Make sure that the log shows the warning about the Mailbox Import Export role.

### Errors

1. Type a server name that does not exist.
2. Click **Connect**.
3. Make sure that an error window shows the cause.
4. Make sure that the window operates.
