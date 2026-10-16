# eM Client Universal Migrator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-lightgrey.svg)](#supported-platforms)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)

A complete, free, open-source tool with a **friendly graphical user interface (GUI)** to migrate **Accounts, Emails, Subfolders, Calendars, and Contacts** from **eM Client** to:
1. **Mozilla Thunderbird** (Automated account configuration, native MBOX stores, Unified Folders)
2. **Microsoft Outlook** (Native Outlook PST generation or live Profile COM automation)
3. **Standalone Files** (Standard MBOX, EML, ICS calendar, and VCF vCard files)

---

## 📸 Friendly Graphical Interface (UI)

Simply run `python main.py` to launch the interactive desktop interface:

```text
+---------------------------------------------------------------------------------+
|  eM Client Universal Migrator                                                   |
|  Transfer Accounts, Emails, Folders, Calendars & Contacts to Thunderbird/Outlook|
+---------------------------------------------------------------------------------+
| 1. Migration Paths & Destination                                                |
|    eM Client Data:     [ C:\Users\user\AppData\Roaming\eM Client       ] [Browse] |
|    Target Application: (o) Mozilla Thunderbird  ( ) Microsoft Outlook ( ) Export|
|    Thunderbird Profile:[ Profiles/xyz.default-release                 ] [Browse] |
+---------------------------------------------------------------------------------+
| 2. Components to Transfer                                                       |
|    [x] Account Settings (OAuth2/IMAP/SMTP)   [x] Emails & Folders               |
|    [x] Calendar Events (.ics)                [x] Contacts / Address Book (.vcf) |
+---------------------------------------------------------------------------------+
| 3. Select Accounts and Folders to Migrate                                       |
|    [Select All]  [Deselect All]  [Refresh Tree]                                 |
|    Account / Folder Name                                               Items    |
|    -----------------------------------------------------------------------------|
|    [-] [x] work@company.com                                            4,520    |
|        [x] INBOX                                                       3,810    |
|        [x] Sent                                                          690    |
|        [x] Archive                                                        20    |
|    [-] [x] personal@gmail.com                                         12,410    |
|        [x] [Gmail]/All Mail                                           11,900    |
|        [x] [Gmail]/Sent Mail                                             510    |
+---------------------------------------------------------------------------------+
|    [========================================] 100%                              |
|    Status: Migration completed successfully!                                    |
|    [Log Console: Migrated 16,930 messages to native stores.]                    |
|                                                            [ ▶ Start Migration ]|
+---------------------------------------------------------------------------------+
```

---

## ✨ Features

- **🖥️ Friendly Graphical User Interface (GUI)**:
  - Interactive treeview to select/deselect specific accounts and individual subfolders.
  - Component toggles: choose whether to include accounts, emails, calendars, or contacts.
  - Live progress bar and real-time log console.
  - Runs in background threads so the UI never hangs.
- **⚡ Zero External Dependencies for Core & Thunderbird**:
  - Pure Python 3 standard library (`tkinter`, `sqlite3`, `email`, etc.).
  - No `pip install` required for standard migrations.
- **🦅 Target: Mozilla Thunderbird**:
  - Injects accounts and identities into `prefs.js`.
  - Configures incoming IMAP/POP3 servers and outgoing SMTP servers.
  - Automatically enables modern **OAuth2** for Microsoft and Google accounts.
  - Reassembles full RFC 822 MIME message trees and writes MboxRD with `X-Mozilla-Status` flags (read/unread).
  - Updates `virtualFolders.dat` and `folderCache.json` so accounts appear in **Unified Folders** (Unified Inbox, Sent, Trash, Junk).
- **💼 Target: Microsoft Outlook**:
  - Automatically creates separate `.pst` data files for each account so they appear as first-level peers in Outlook's sidebar.
  - Maps nested folder paths to Outlook's native folders (e.g. `Εισερχόμενα` / `Inbox`, `Απεσταλμένα` / `Sent`, `Διαγραμμένα` / `Deleted`, `Ανεπιθύμητα` / `Junk`).
  - Decodes RFC 2047 encoded email subjects (`=?UTF-8?...`).
  - Detects and preserves HTML bodies and attachments.
  - *(Requires Windows + `pip install pywin32` for Outlook COM automation)*.
- **📅 Calendars & 📇 Contacts**:
  - Master `eMClient_All_Calendars.ics` and per-account `.ics` files.
  - Master `eMClient_All_Contacts.vcf` (vCard 3.0) and per-account `.vcf` files.
- **🛡️ Safety First**:
  - Automatic timestamped backups of `prefs.js`.
  - Stale process lock checking.
  - Non-destructive read-only access to eM Client SQLite databases.

---

## 🚀 Quick Start

### 1. Launch the Graphical Interface
```bash
python main.py
```
*(Or double-click `main.py` on Windows).*

### 2. Command Line (Headless / Automation)
You can also run the tool directly from the terminal without the GUI:

```bash
# Preview what would be migrated
python main.py --dry-run

# Migrate to Thunderbird (default)
python main.py --target thunderbird

# Migrate to Microsoft Outlook (PST files)
python main.py --target outlook

# Export emails to standalone EML files
python main.py --target export
```

---

## 🛠️ CLI Options

```text
usage: main.py [-h] [--gui] [--target {thunderbird,outlook,export}]
               [--emclient-path PATH] [--tb-profile PATH] [--dry-run]
               [--mails-only] [--calendar-only] [--contacts-only]

eM Client Universal Migrator (Thunderbird & Outlook)

options:
  -h, --help            Show this help message and exit
  --gui                 Launch the graphical user interface
  --target {thunderbird,outlook,export}
                        Migration target (default: thunderbird)
  --emclient-path PATH  Custom path to eM Client data folder
  --tb-profile PATH     Custom path to Thunderbird profile folder
  --dry-run             Preview migration without modifying files
  --mails-only          Only migrate emails
  --calendar-only       Only export calendar events
  --contacts-only       Only export contacts
```

---

## 📖 Importing Calendars & Contacts

After running the migrator, calendar and contact files are stored in:
`Documents/Thunderbird Migration/`

- **Thunderbird Calendar**: Switch to Calendar tab (`Ctrl+Shift+C`) ➔ `Events and Tasks` ➔ `Import...` ➔ select `eMClient_All_Calendars.ics`.
- **Thunderbird Contacts**: Open Address Book (`Ctrl+Shift+B`) ➔ `Tools` ➔ `Import` ➔ choose **vCard file (.vcf)** ➔ select `eMClient_All_Contacts.vcf`.
- **Outlook**: `File` ➔ `Open & Export` ➔ `Open Calendar (.ics)` or import vCards.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
Feel free to use, modify, and distribute it!
