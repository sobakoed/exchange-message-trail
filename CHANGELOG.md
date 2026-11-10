# Changelog

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
