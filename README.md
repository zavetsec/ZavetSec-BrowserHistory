<div align="center">

# `>_` ZavetSec-BrowserHistory

**Forensic browser history extractor for Windows — all users, all browsers, one report**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)](https://github.com/zavetsec)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078d4?logo=windows&logoColor=white)](https://github.com/zavetsec)
[![No Dependencies](https://img.shields.io/badge/Dependencies-none%20required-00d4ff)](https://github.com/zavetsec)
[![License: MIT](https://img.shields.io/badge/License-MIT-green)](https://github.com/zavetsec)

Bypasses file locks via VSS Shadow Copy · 16 browsers · interactive HTML report with search and filters

[Quick Start](#-quick-start) · [Parameters](#-parameters) · [Report](#-html-report) · [Requirements](#-requirements) · [FAQ](#-faq)

</div>

---

## What it is

`ZavetSec-BrowserHistory.ps1` is a standalone PowerShell script for forensic analysis. It extracts browser history from **all local user profiles** on a Windows machine and generates a self-contained interactive HTML report.

Used during incident response, insider threat investigations, and workstation reviews — no third-party tools required.

---

## Features

- **16 browsers** — Chrome, Edge, Firefox, Brave, Opera, Opera GX, Yandex, Vivaldi, Tor Browser, Thunderbird, Waterfox, LibreWolf, Pale Moon, Epic, Comodo Dragon, OneDrive WebView
- **All users** — enumerates all profiles via the registry (`ProfileList`), not just the current user
- **File lock bypass** — VSS Shadow Copy allows reading `History` / `places.sqlite` from running browsers
- **Dual parse mode** — full SQL via `sqlite3.exe` (titles, visit counts, precise timestamps) or regex fallback with zero dependencies
- **Interactive HTML report** — search, filter by user / browser / date, column sorting, in-browser CSV export
- **Optional CSV export** — save `.csv` alongside the HTML for SIEM ingestion or Excel
- **Date filtering** — at collection level (`-DateFrom` / `-DateTo`) and inside the report itself

---

## ⚡ Quick Start

```powershell
# Minimal run (Administrator required)
.\ZavetSec-BrowserHistory.ps1

# Open the report immediately after generation
.\ZavetSec-BrowserHistory.ps1 -OpenReport

# Full run with CSV and date filter
.\ZavetSec-BrowserHistory.ps1 -OpenReport -CsvExport -DateFrom 2025-01-01 -DateTo 2025-06-30

# Specify an explicit output path
.\ZavetSec-BrowserHistory.ps1 -OutputPath "C:\IR\host42_history.html" -OpenReport

# Skip archiving — save report as plain HTML (useful if 7-Zip is unavailable)
.\ZavetSec-BrowserHistory.ps1 -NoArchive -OpenReport
```

> Without arguments the report is saved to `.\Reports\<HOSTNAME>_<TIMESTAMP>.html`

---

## 📋 Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-OutputPath` | String | auto | Path for the HTML report. Auto-generated in `.\Reports\` if omitted |
| `-MaxRecordsPerBrowser` | Int | `5000` | Maximum records to extract per browser profile |
| `-OpenReport` | Switch | — | Open the report in the default browser after generation |
| `-CsvExport` | Switch | — | Save results as CSV alongside the HTML |
| `-DateFrom` | String | — | Collect records starting from this date (`yyyy-MM-dd`) |
| `-DateTo` | String | — | Collect records up to and including this date (`yyyy-MM-dd`) |
| `-NoArchive` | Switch | — | Skip ZIP archiving — save HTML (and CSV) as plain files |

---

## 📊 HTML Report

The report is a single self-contained `.html` file — no internet connection or server required.

```
┌──────────────┬──────────────┬───────────────┬──────────────────┐
│ Total Records│ Users Scanned│ Unique Domains│  Extraction Mode │
│    12 847    │      3       │     1 204     │  VSS Shadow Copy │
└──────────────┴──────────────┴───────────────┴──────────────────┘
```

**Side panels:**
- Records by User — horizontal bars with record counts per user
- Top Domains — top 15 domains by visit frequency

**Records table:**

| # | User | Browser | Domain | Title / URL | Visits | Last Visit |
|---|------|---------|--------|-------------|--------|------------|
| 1 | john | Edge | github.com | GitHub · Build... | ▬▬▬ 42 | 2025-06-14 09:31 |

**Filters:**
- User and browser toggle buttons
- Full-text search across URL / title / domain
- Date range filter (From / To) with native date pickers
- `⬓ CSV` button — exports only **visible** (filtered) rows, BOM UTF-8 for Excel compatibility
- Click any column header to sort ascending / descending

---

## 🔧 Requirements

| | |
|---|---|
| **PowerShell** | 5.1+ (built into Windows 8.1 / Server 2012 R2 and later) |
| **Privileges** | Local Administrator — required for VSS and reading other users' profiles |
| **7z.exe** | Optional but recommended. Required for AES-256 encrypted ZIP output |
| **sqlite3.exe** | Optional. Without it — regex fallback (URLs only, no titles or visit counts) |

### Installing sqlite3.exe (recommended)

1. Download `sqlite-tools-win-x64-*.zip` from [sqlite.org/download.html](https://sqlite.org/download.html)
2. Place `sqlite3.exe` next to the script or in a `sqlite3\` subfolder

### Installing 7z.exe (recommended)

1. Download and install [7-Zip](https://www.7-zip.org/) — or copy `7z.exe` from an existing installation
2. Place `7z.exe` next to the script, or ensure it is available system-wide in `PATH`

```
ZavetSec-BrowserHistory.ps1
sqlite3.exe              ← or sqlite3\sqlite3.exe
Reports\
```

---

## 🌐 Supported Browsers

| Engine | Browsers |
|---|---|
| **Chromium** | Google Chrome, Microsoft Edge, Brave, Yandex Browser, Opera, Opera GX, Vivaldi, Epic Privacy Browser, Comodo Dragon, OneDrive WebView, Win WebExperience |
| **Firefox / Gecko** | Mozilla Firefox, Tor Browser, Thunderbird, Waterfox, LibreWolf, Pale Moon |

---

## 🛡️ Extraction Modes

```
Running as Administrator
        │
        ├─ VSS available → Shadow Copy of C:\
        │      Reads History / places.sqlite from running browsers
        │      (files are not locked inside the shadow copy)
        │
        └─ No privileges / VSS unavailable → Direct Copy
               Attempts ReadAllBytes / Copy-Item / xcopy fallback chain
               (may miss profiles with currently open browser sessions)

sqlite3.exe found
        │
        ├─ Yes → SQL query: URL + Title + VisitCount + Timestamp
        │
        └─ No  → Regex fallback: URL only, Visits=1, Time=unknown
```

---

## 📁 Output Structure

**Default (with archiving):**
```
.\Reports\
└── HOSTNAME_20250614_093100.zip   ← AES-256 encrypted archive
```
The unencrypted `.html` (and `.csv` if `-CsvExport`) are deleted after archiving.

**With `-NoArchive`:**
```
.\Reports\
├── HOSTNAME_20250614_093100.html
└── HOSTNAME_20250614_093100.csv   ← with -CsvExport
```

**CSV columns:** `UserName, Browser, Domain, Title, URL, Visits, LastVisit`

---

## 🔑 Archive Password

A 16-character random password is generated on each run and printed to the console on completion:

```
  ================================================================
   REPORT READY
  ================================================================

   Archive  : .\Reports\HOSTNAME_20250614_093100.zip

   Password : aB3$mK9#Xv2!pQnZ

   Save this password — it will not be shown again.
  ================================================================
```

The password is not stored anywhere. Copy it before closing the terminal.

**Encryption requires `7z.exe`** (7-Zip). If not found, the script falls back to a standard unencrypted ZIP with a warning. Place `7z.exe` next to the script or install [7-Zip](https://www.7-zip.org/) system-wide.

```
ZavetSec-BrowserHistory.ps1
7z.exe                   ← optional but recommended
sqlite3.exe              ← or sqlite3\sqlite3.exe
Reports\
```

---

## ❓ FAQ

**Can it run without Administrator rights?**
Yes — the script will run and read the current user's profile. VSS will be unavailable and other users' profiles will not be accessible. Full collection requires Administrator.

**The browser is running and the file is locked — what happens?**
If running as Admin, the script automatically creates a VSS shadow copy and reads the database files from it — no lock issues.

**Are portable browser installations supported?**
No — the script looks for profiles at standard `AppData` paths. Portable installs with non-standard locations are not detected.

**How does this relate to ZavetSec Triage?**
`Invoke-ZavetSecTriage.ps1` includes browser history collection as module #14. `ZavetSec-BrowserHistory.ps1` is a standalone tool for when you need a focused, detailed browser-only report.

---

## 📋 Changelog

### v1.0 — Initial release
- VSS Shadow Copy — file lock bypass for running browsers
- 16 browsers, all local user profiles via registry enumeration
- Dual parse mode: full SQL (`sqlite3.exe`) and regex fallback
- `-DateFrom` / `-DateTo` — date range filtering at collection level
- `-CsvExport` — parallel CSV output alongside the HTML report
- `[CmdletBinding()]` + `Write-Progress` — `-Verbose` support and console progress bar
- AES-256 encrypted ZIP output via `7z.exe` — 16-char random password printed to console
- `-NoArchive` — skip archiving, save report as plain HTML for environments without 7-Zip
- HTML: interactive report with search, user / browser / date filters, column sorting
- HTML: in-browser CSV export of visible rows (BOM UTF-8, Excel-compatible)
- HTML: visit count mini-bar, URL and domain tooltips, browser favicon

---

## `>_` disclaimer

> This tool is intended for legitimate forensic analysis within authorized IR investigations on systems you have permission to access. The author assumes no responsibility for use outside of those boundaries.

---

<div align="center">

**ZavetSec** — security tooling for those who read logs at 2am

[github.com/zavetsec](https://github.com/zavetsec) · MIT License

</div>
