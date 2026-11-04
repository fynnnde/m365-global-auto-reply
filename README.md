<div align="center">

# m365-global-auto-reply

**One message for every mailbox.** See and manage the automatic replies of your whole Microsoft 365 tenant from a single window.

![Status](https://img.shields.io/badge/status-alpha-orange)
![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D6)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1-5391FE)
![Exchange Online](https://img.shields.io/badge/Exchange%20Online-supported-2e9e5b)
![Built with Claude](https://img.shields.io/badge/built%20with-Claude-D97757)

</div>

> [!WARNING]
> **This is an early alpha release (v0.1.0-alpha).**
> It was mostly tested against a simulated tenant, so **bugs are very likely**, especially in the template editor.
> Please try it on one or two **test mailboxes** first and check the result in Outlook before you run it on everyone.
> Found a bug? Open an issue and attach the lines from the Activity log.

---

## Why

Setting a company-wide out-of-office for holidays or an office move usually means a PowerShell script, a copy of the HTML text and a lot of hoping.
This tool puts a proper interface on top of it: sign in, pick a template, pick a period, apply. Nobody is left on read.

## Features

| | |
|---|---|
| **Overview** | Every user mailbox with its current state, period and the message that is set right now |
| **Bulk apply** | Create an auto-reply for the mailboxes you checked, or for everyone, with template, period and options |
| **Per mailbox** | Enable, disable or edit the auto-reply of a single mailbox straight from the list |
| **Safe by default** | Mailboxes that already have an active auto-reply are skipped unless you choose to overwrite |
| **Guard rails** | Typed confirmation for changes to 10 or more mailboxes, one failing mailbox never stops the run, automatic retries when Exchange throttles |
| **Template editor** | Fonts, sizes, colors, bold, italic, underline, strikethrough, alignment, lists, links, images, plus raw HTML view |
| **Import** | Open your existing `.html` or `.md` templates |
| **Languages** | English and German interface |

## Quick start

### Option A: the executable

1. Download `m365-global-auto-reply-v0.1.0-alpha.zip` from the [releases page](../../releases) and unzip it anywhere.
2. Start `m365-global-auto-reply.exe` and click **Sign in with Microsoft**.

That is all. The exe is a single file. On the first start it unpacks its icon and font to `%LOCALAPPDATA%\m365-global-auto-reply`, and if the Exchange Online PowerShell module is missing it **offers to install it for you** (current user only, no admin rights needed).

> [!NOTE]
> Windows SmartScreen or your antivirus may warn about the unsigned executable. This is a common false positive for scripts packed with PS2EXE. If you prefer, run it from source (Option B).

### Option B: from source

```powershell
git clone https://github.com/fynnnde/m365-global-auto-reply.git
cd m365-global-auto-reply
powershell -STA -ExecutionPolicy Bypass -File .\src\m365-global-auto-reply.ps1
```

## Requirements

- Windows 10 or 11 with **Windows PowerShell 5.1** (already included)
- The [`ExchangeOnlineManagement`](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell) module (the app offers to install it when it is missing)
- A Microsoft 365 account that is allowed to read and change the auto-reply settings of other mailboxes, for example an Exchange administrator
- Sign-in works with MFA. It opens your browser, no password is typed into the app

## How to use it

1. **Sign in.** The mailbox list loads, then the app reads the auto-reply of every mailbox in the background. The progress bar and the Activity log show where it is.
2. **Read the table.** A green blinking dot means the auto-reply is on, a red dot means it is off.
3. **Create an auto-reply.** Click **New auto-reply**, choose a template and a period, decide whether to overwrite existing replies, and pick *Checked mailboxes* or *All mailboxes*.
4. **Single mailbox.** Use **Edit** or **Enable / Disable** on a row.
5. **Disable many at once.** Check mailboxes, click **Disable checked**.
6. **Templates.** Open the **Templates** tab to write and save your own.

### Periods

| Option | Behavior |
|---|---|
| Until I turn it off | Stays active until you disable it |
| From now until an end date | Starts immediately, ends at the given time |
| Between a start and an end date | Scheduled window |

Dates use the format `dd.MM.yyyy HH:mm`, for example `24.12.2026 08:00`.

## Templates

Templates are plain HTML files.

| Where | What |
|---|---|
| `templates\` in the repository, or a built-in default inside the exe | Built-in templates, read-only |
| `%LOCALAPPDATA%\m365-global-auto-reply\templates` | Your own templates |

A template you save under the name of a built-in one replaces it. Deleting yours brings the built-in back.

### Images

Exchange only accepts the HTML body of an auto-reply. Images cannot be attached, and Outlook does not display images that are embedded in the message.
So images **must be hosted online**:

1. Upload the image to a service such as [imgur](https://imgur.com).
2. Insert the **direct link**. It has to start with `https://` and end with `.png`, `.jpg`, `.jpeg` or `.gif`.

External recipients may still have to click *Download pictures* in their mail client. No tool can change that.

## Known issues (alpha)

- Mostly tested against a mock tenant, real tenants may behave differently
- The editor is built on the old browser control that ships with Windows. Some formatting combinations behave oddly and undo drops the selection
- Reading the status takes about one second per mailbox, large tenants take a while (the list is usable meanwhile)
- Free positioning of text or images is not possible, mail clients do not support it. Images resize in steps of 10 px, paragraphs indent in steps of 40 px
- The executable is not code-signed

## Security and privacy

- No password and no token is stored. Sign-in is handled by the Exchange Online module
- Scripts, frames, forms and event handlers are stripped from templates before they are previewed, saved or sent
- Logs are written to `%LOCALAPPDATA%\m365-global-auto-reply\logs` and **contain mailbox addresses**. Do not post them publicly without removing those
- The app talks to Exchange Online only. There is no telemetry

## Build it yourself

```powershell
.\build.ps1
```

The script first checks that the app script parses and all windows load, then installs [PS2EXE](https://github.com/MScholtes/PS2EXE) for the current user if needed, bakes the icon and font into the script and writes one `dist\m365-global-auto-reply.exe`.

### Try it without a tenant

A mock mode simulates 60 mailboxes, one of them always failing:

```powershell
$env:M365AR_MOCK = '1'
powershell -STA -File .\src\m365-global-auto-reply.ps1
```

## Project structure

```text
m365-global-auto-reply
├── src/        the application (single PowerShell script, WPF)
├── templates/  built-in templates
├── assets/     app icon and the Font Awesome font
└── build.ps1   builds dist/m365-global-auto-reply.exe
```

## Built with AI

This project was written together with **Claude** (Anthropic): design, code and this README came out of a long back-and-forth between a human and an AI assistant. The human decided what to build, tried every version and reported what was broken. Please review the code before you trust it with your tenant.

## Credits and disclaimer

- Icons: [Font Awesome Free](https://fontawesome.com/license/free), fonts under SIL OFL 1.1, icons under CC BY 4.0
- Executable packaging: [PS2EXE](https://github.com/MScholtes/PS2EXE)
- This is an independent project. It is **not affiliated with, endorsed by or sponsored by Microsoft**. Microsoft 365, Exchange Online and Outlook are trademarks of Microsoft Corporation
- Provided as is, without warranty. You are responsible for what you send to your mailboxes

## License

No license has been chosen yet. Until one is added, all rights are reserved.
