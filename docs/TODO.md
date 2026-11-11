# To-do

Possible future work. These items are not in the plan of a version.

## Level 2 of the mailbox check: "Where is the message"

The mailbox check (level 1) finds causes: rules, junk e-mail, forwarding, quota.
It does not tell the folder of the message.

Level 2 finds the folder of the message in the recipient mailbox by the message ID:

- the folder (for example Inbox, Junk E-mail, Deleted Items, Projects/2026);
- read or not read;
- in Recoverable Items or not.

| Item | Decision |
|---|---|
| Interface | EWS (Exchange Web Services), with SOAP requests and Kerberos. No external libraries. |
| Search | `FindFolder` (Deep) and `FindItem` with a filter on `PR_INTERNET_MESSAGE_ID` |
| Permission | The role ApplicationImpersonation. Get the approval of the security team first. |
| Place in the window | A new button on the tab **Recipients**, next to **Check the recipient mailbox** |

## Exchange server

- Find and remove the damaged search action on the server. The full list `Get-ComplianceSearchAction` fails.
- Examine if `Stop-ComplianceSearch` makes a search complete faster.

## Classic window

- Remove `ExchangeMessageTrail.Classic.ps1` and `EMT-Classic.cmd` when the WPF window operates on the servers without problems.
  Remove also the key `Window` (classic window size) from `settings.json` and the classic file from the language key test.
