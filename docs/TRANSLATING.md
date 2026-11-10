# Translating Exchange Message Trail

This document tells you how to add a language to the user interface of Exchange Message Trail (EMT).
You do not have to change the code.

## Contents

- [How languages work](#how-languages-work)
- [File format](#file-format)
- [Procedure: add a language](#procedure-add-a-language)
- [Procedure: examine your file](#procedure-examine-your-file)
- [Rules for the texts](#rules-for-the-texts)
- [Procedure: send your translation](#procedure-send-your-translation)

## How languages work

- Each language is one file in the folder `Lang`: `Lang\<code>.lng`.
- EMT reads all `.lng` files when it starts.
- The language list in the top line of the window shows all languages that it found.
- The language changes immediately. EMT keeps the selection for the next start.
- If a key is not in your file, EMT uses the English text for this key.
  Thus a partial translation is possible.

The bundled files are:

| File | Language | Status |
|---|---|---|
| `Lang\en.lng` | English | Reference file. All keys. |
| `Lang\ru.lng` | Russian | All keys. |

## File format

```
# EMT language file
@code = de
@name = Deutsch
@author = Your Name

# ---- Window: frame and connection ----
conn.connect = Verbinden
log.connected = Verbunden mit {0}.
```

| Line | Meaning |
|---|---|
| `# text` or `; text` | Comment. EMT ignores it. |
| Empty line | EMT ignores it. |
| `@code = de` | The language code (ISO 639-1, two letters). If this line is missing, the file name is the code. |
| `@name = Deutsch` | The name in the language list. Write it in your language. |
| `@author = …` | Your name or your team. Optional. |
| `key = text` | One text. The key is before the first `=`. The text is after it. |

- The encoding is UTF-8. A BOM is permitted.
- The text can contain `=`. Only the first `=` separates the key and the text.
- EMT removes the spaces at the start and at the end of the key and of the text.
- One text is one line. A line break in a text is not possible.

## Procedure: add a language

1. Copy `Lang\en.lng`.
2. Give the copy the name `<code>.lng`. For example: `de.lng`.
3. Open the copy in a text editor that saves UTF-8 (for example, Visual Studio Code or Notepad).
4. Change the line `@code`.
5. Change the line `@name`.
6. Change the line `@author`.
7. Translate the text after `=` in each line.
8. Do not change the keys (the text before `=`).
9. Save the file in UTF-8.
10. Start EMT.
11. Select your language in the language list.
12. Examine each part of the window (see [Rules for the texts](#rules-for-the-texts)).

## Procedure: examine your file

EMT has a check for language files. Do these steps:

1. Open Windows PowerShell 5.1.
2. Go to the EMT folder.
3. Type these commands:

```powershell
. .\src\EMT.Strings.ps1
Test-EmtLanguageFile .\Lang\de.lng
```

No output = the file is correct.

The check finds these problems:

| Message | Cause | Action |
|---|---|---|
| `unknown key '…'` | The key is not in `en.lng` | Correct the key, or remove the line |
| `missing key '…'` | The key is not in your file | Add the line. Until then, EMT uses the English text. |
| `key '…': placeholders …` | The `{0}`, `{1}` in your text are not the same as in `en.lng` | Use the same placeholders |
| `duplicate key '…'` | The key occurs two times | Remove one line |
| `line N: no 'key = text'` | The line has no `=` | Correct the line, or make it a comment |
| `no @name line` | The file has no language name | Add `@name = …` |

To accept a partial translation, add `-AllowMissing`:

```powershell
Test-EmtLanguageFile .\Lang\de.lng -AllowMissing
```

## Rules for the texts

**Placeholders.** `{0}`, `{1}`, `{2}` are values that EMT puts in the text:
a server name, a count, an address. Keep each placeholder. You can change the order in the sentence.
A format in a placeholder, for example `{2:N1}` or `{1:N0}`, must stay the same.

**Fixed words.** Do not translate these words:

- Cmdlet and product names: `Search-Mailbox`, `Compliance Search`, `SoftDelete`, `Exchange`, `Outlook`, `KQL`, `MessageId`, `FQDN`.
- Role names: `Mailbox Search`, `Mailbox Import Export`, `Discovery Management`, `Message Tracking`.
- The search name prefixes `EMT_*` and `MC_*`.

**Date formats.** The keys `fmt.picker`, `fmt.datetime` and `fmt.date` are .NET date formats
(for example `dd.MM.yyyy HH:mm`). Use the usual format of your language.

**The confirmation word.** The key `confirm.word` is the word that the user types to confirm a delete
(`DELETE` in English). Use one word in capital letters.

**Length.** Buttons get wider for a long text. Labels in the criteria area have a limited space.
Keep the texts short. Make sure that the window shows the full text at the minimum window size.

**Style.** Use short sentences. Use one term for one thing in all texts.
Use the same tone as `en.lng`: clear instructions, no slang.

## Procedure: send your translation

1. Run the check (see [Procedure: examine your file](#procedure-examine-your-file)).
2. Run the unit tests: `Invoke-Pester -Path .\tests`.
   The tests examine all `.lng` files for unknown keys and placeholders.
3. Make a pull request with the new file `Lang\<code>.lng`.
4. In the pull request, write the language and the Windows version that you used for the test.
