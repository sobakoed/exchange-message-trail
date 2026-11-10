# Testing

This document gives the unit tests and the test procedure on Exchange.

## Unit tests

The unit tests examine `src/MailCleaner.Core.ps1`. They do not need Exchange.
Stub functions and Pester mocks replace the Exchange cmdlets.

**Requirements:** Pester 5.0 or later.

To install Pester 5, do this step:

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

Do this procedure on a test mailbox set before you use MailCleaner in production.

**Preparation:**

1. Get three test mailboxes: `test1`, `test2`, `test3`.
2. From `test1`, send a message to `test2` and `test3`. Use a unique subject.
3. From `test2`, forward the message to `test3`.
4. Wait 5 minutes for the index.

**Search:**

1. Start MailCleaner.
2. Connect to the server.
3. Make sure that the status line shows the roles.
4. In **Отправители**, type the address of `test1`.
5. In **Тема содержит**, type the unique subject.
6. Click **Сегодня**.
7. Click **Найти**.
8. Make sure that the tab **Письма** shows the message and the forward.
9. Make sure that the tab **Найдено в ящиках** shows `test1`, `test2` and `test3`.

**Search speed:**

1. Do the search again.
2. Make sure that the summary line shows `быстрый поиск`. There are 3 mailboxes.
3. Make sure that the log shows the time of each stage.
4. In **Получатель**, type the address of `test3`.
5. Select **Только ящик получателя**.
6. Click **Найти**.
7. Make sure that the tab **Найдено в ящиках** shows only `test3`.

**Window:**

1. Drag the border of a column in a table. Make sure that the width changes.
2. Click the column header. Make sure that the sort keeps the width.
3. Select **Искать во всех ящиках**. Then select **Только ящик получателя**. Make sure that the first check box is clear.
4. Make the window as small as possible. Make sure that all buttons are fully visible.

**Partial address:**

1. In **Отправители**, type a part of the alias of `test1`.
2. Click **Найти**.
3. Make sure that the field shows the full address of `test1` and the log shows the change.
4. Type a part of a name that more than one user has. Click **Найти**.
5. Make sure that the selection window shows the users.
6. Type a text that no user has. Click **Найти**.
7. Make sure that an error window shows the text and the search does not start.

**Subject with brackets:**

1. Send a test message with the subject `[TEST] Unique subject`.
2. Search with this subject.
3. Make sure that the tab **Письма** shows the message.

**Cancel:**

1. Click **Найти**.
2. Click **Отмена** during the search.
3. Make sure that the log shows the cancel.
4. Make sure that the window operates.

**Soft delete:**

1. Clear **Получатель**.
2. Do the search again.
3. Click **Удалить (мягко)**.
4. Confirm the delete.
5. Make sure that the log shows a new Compliance Search before the delete.
6. Make sure that the tab **Проверка** shows `0` in **В видимых папках** for each mailbox.
7. Make sure that **Всего** is more than `0`. The messages are in Recoverable Items.

**Hard delete:**

1. Do the search again.
2. Click **Удалить безвозвратно**.
3. Type `УДАЛИТЬ`.
4. Confirm the delete.
5. Make sure that the tab **Проверка** shows `0` in both columns.
6. If a mailbox is on hold, make sure that the log shows the hold warning.

**Audit:**

1. Open `%USERPROFILE%\MailCleaner\audit.csv`.
2. Make sure that there is one row for each search, delete and verification.

**Permissions:**

1. Start MailCleaner with an account without the Mailbox Import Export role.
2. Connect to the server.
3. Make sure that **Удалить безвозвратно** is disabled.
4. Make sure that the status line shows `Безвозвратное ✗`.

**Errors:**

1. Type a server name that does not exist.
2. Click **Подключиться**.
3. Make sure that an error window shows the cause.
4. Make sure that the window operates.
