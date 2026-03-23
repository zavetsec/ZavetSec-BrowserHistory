<div align="center">

# `>_` ZavetSec-BrowserHistory

**Browser history acquisition for Windows IR — all users, all browsers, one encrypted report**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)](https://github.com/zavetsec)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078d4?logo=windows&logoColor=white)](https://github.com/zavetsec)
[![No Dependencies](https://img.shields.io/badge/Dependencies-none%20required-00d4ff)](https://github.com/zavetsec)
[![License: MIT](https://img.shields.io/badge/License-MIT-green)](https://github.com/zavetsec)

Bypasses file locks via VSS Shadow Copy · 16 browsers · AES-256 encrypted output · interactive HTML report

[When to Use](#-when-to-use) · [Quick Start](#-quick-start) · [Parameters](#-parameters) · [Report](#-html-report) · [Limitations](#-limitations)

</div>

---

## What it is

A standalone PowerShell script for rapid browser history acquisition during incident response. Runs as a single file, requires no installation, and produces a self-contained interactive HTML report packed into an AES-256 encrypted ZIP.

Designed for the scenario where you are on a live Windows machine, need browser evidence fast, and cannot install tools.

---

## When to Use

**Incident Response / DFIR**
A host is suspected of compromise. You need to know what sites were visited, when, and by which user — before the machine is reimaged or isolated. Run the script, get the encrypted archive, take it off the machine.

**Insider Threat Investigation**
An account shows anomalous behavior. Browser history correlated across all local profiles can confirm or rule out data exfiltration via web uploads, webmail, or cloud storage.

**Workstation Audit**
Pre-termination review, policy compliance check, or post-incident reconstruction. Covers all user profiles in a single run without touching each account manually.

**When not to use**
This is a live-acquisition triage tool, not a forensic imager. It does not preserve original file metadata, does not produce chain-of-custody artifacts, and should not be the sole evidence source in legal proceedings.

---

## Features

- **16 browsers** — Chrome, Edge, Firefox, Brave, Opera, Opera GX, Yandex, Vivaldi, Tor Browser, Thunderbird, Waterfox, LibreWolf, Pale Moon, Epic, Comodo Dragon, OneDrive WebView
- **All users** — enumerates every local profile via the registry (`ProfileList`), not just the current session
- **File lock bypass** — VSS Shadow Copy reads `History` / `places.sqlite` from browsers that are currently running
- **Dual parse mode** — full SQL via `sqlite3.exe` (titles, visit counts, precise timestamps) or regex fallback with zero dependencies
- **Encrypted output** — AES-256 ZIP via `7z.exe`, 16-char random password printed to console once
- **Interactive HTML report** — search, filter by user / browser / date, column sorting, in-browser CSV export
- **Date filtering** — at collection level (`-DateFrom` / `-DateTo`) and interactively inside the report

---

## ⚡ Quick Start

```powershell
# Standard run — produces encrypted ZIP + prints password (Administrator required)
.\ZavetSec-BrowserHistory.ps1

# Open report immediately after generation
.\ZavetSec-BrowserHistory.ps1 -OpenReport

# With CSV export and date range
.\ZavetSec-BrowserHistory.ps1 -OpenReport -CsvExport -DateFrom 2025-01-01 -DateTo 2025-06-30

# No archive — plain HTML output (when 7-Zip is unavailable on the host)
.\ZavetSec-BrowserHistory.ps1 -NoArchive -OpenReport

# Custom output path
.\ZavetSec-BrowserHistory.ps1 -OutputPath "C:\IR\host42_history.html"
```

> Without arguments the report is saved to `.\Reports\<HOSTNAME>_<TIMESTAMP>.zip`

---

## 📋 Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-OutputPath` | String | auto | Path for the HTML report. Auto-generated in `.\Reports\` if omitted |
| `-MaxRecordsPerBrowser` | Int | `5000` | Maximum records per browser profile |
| `-OpenReport` | Switch | — | Open the report after generation |
| `-CsvExport` | Switch | — | Save results as CSV alongside the HTML |
| `-DateFrom` | String | — | Collect records from this date (`yyyy-MM-dd`) |
| `-DateTo` | String | — | Collect records up to this date (`yyyy-MM-dd`) |
| `-NoArchive` | Switch | — | Skip ZIP — save HTML (and CSV) as plain files |

---

## 📊 HTML Report

Single self-contained `.html` file — no server, no internet required. Packed into an encrypted ZIP by default.

```
┌──────────────┬──────────────┬───────────────┬──────────────────┐
│ Total Records│ Users Scanned│ Unique Domains│  Extraction Mode │
│    12 847    │      3       │     1 204     │  VSS Shadow Copy │
└──────────────┴──────────────┴───────────────┴──────────────────┘
```

**Side panels:** Records by User · Top 15 Domains by visit count

**Records table:**

| # | User | Browser | Domain | Title / URL | Visits | Last Visit |
|---|------|---------|--------|-------------|--------|------------|
| 1 | john | Edge | github.com | GitHub · Build... | ▬▬▬ 42 | 2025-06-14 09:31 |

**Filters:** user · browser · date range · full-text search · `⬓ CSV` export of visible rows (BOM UTF-8) · column sort

---

## 🔑 Archive Password

Generated fresh on every run using a cryptographically secure RNG. Printed once to console on completion — not stored anywhere.

```
  ================================================================
   REPORT READY
  ================================================================

   Archive  : .\Reports\HOSTNAME_20250614_093100.zip
   Password : aB3$mK9#Xv2!pQnZ

   Save this password - it will not be shown again.
  ================================================================
```

If `7z.exe` is not found next to the script or in `PATH`, the script falls back to an unencrypted ZIP (via built-in .NET) and prints a warning. Use `-NoArchive` to skip archiving entirely.

---

## 🔧 Setup

The script runs with no external files. Optional tools improve output quality:

| File | Effect |
|---|---|
| `sqlite3.exe` | Full SQL parse: titles + visit counts + timestamps. Without it: URL-only regex fallback |
| `7z.exe` | AES-256 encrypted ZIP output. Without it: unencrypted ZIP fallback |

Both files are included in the [release archive](https://github.com/zavetsec/ZavetSec-BrowserHistory/releases/latest). Download the release ZIP, extract everything to a folder — the script will find them automatically.

Alternatively, download them from official sources and place next to the script:
- `sqlite3.exe` — [sqlite.org/download.html](https://sqlite.org/download.html) (`sqlite-tools-win-x64-*.zip`)
- `7z.exe` — [7-zip.org](https://www.7-zip.org/) (copy from any existing 7-Zip installation)

```
ZavetSec-BrowserHistory.ps1
sqlite3.exe        ← included in release archive
7z.exe             ← included in release archive
Reports\           ← created automatically
```

---

## 🛡️ How It Works

```
Running as Administrator?
        │
        ├─ Yes → VSS Shadow Copy of C:\
        │        Reads locked database files from the shadow
        │
        └─ No  → Direct copy fallback
                 (ReadAllBytes → Copy-Item → xcopy)
                 Running browsers may have locked files

sqlite3.exe present?
        │
        ├─ Yes → SQL: URL + Title + VisitCount + Timestamp
        │
        └─ No  → Regex scan of raw DB bytes: URL only

7z.exe present?
        │
        ├─ Yes → AES-256 encrypted ZIP, password to console
        │
        └─ No  → Unencrypted ZIP (built-in .NET compression)
```

---

## 🌐 Supported Browsers

| Engine | Browsers |
|---|---|
| **Chromium** | Google Chrome, Microsoft Edge, Brave, Yandex Browser, Opera, Opera GX, Vivaldi, Epic Privacy Browser, Comodo Dragon, OneDrive WebView, Win WebExperience |
| **Firefox / Gecko** | Mozilla Firefox, Tor Browser, Thunderbird, Waterfox, LibreWolf, Pale Moon |

---

## 📁 Output

**Default:**
```
.\Reports\
└── HOSTNAME_20250614_093100.zip    ← AES-256 encrypted
```
Source `.html` / `.csv` are deleted after successful archiving.

**With `-NoArchive`:**
```
.\Reports\
├── HOSTNAME_20250614_093100.html
└── HOSTNAME_20250614_093100.csv    ← with -CsvExport
```

**CSV columns:** `UserName, Browser, Domain, Title, URL, Visits, LastVisit`

---

## ⚠️ Limitations

- **Live acquisition only** — reads from a running system, does not image or hash original database files
- **No chain of custody** — not a substitute for forensic imaging tools (FTK, dd, KAPE) in legal contexts
- **Standard profile paths only** — portable browser installs or custom profile locations are not detected
- **C:\ only** — VSS shadow copy targets the system drive; profiles on other drives use direct copy
- **No history deletion detection** — does not identify gaps, cleared history, or anti-forensic activity
- **Regex fallback is approximate** — without `sqlite3.exe`, results may include duplicate or partial URLs

---

## ❓ FAQ

**Can it run without Administrator rights?**
Yes — reads the current user's profile only. VSS is unavailable, other profiles are not accessible.

**The browser is running and its database is locked — does it still work?**
Yes, if running as Administrator. VSS shadow copy bypasses file locks entirely.

**How does this relate to ZavetSec Triage?**
`Invoke-ZavetSecTriage.ps1` includes browser history as module #14 alongside 17 other collection modules. This script is the standalone version — use it when you need browser evidence only, or when triage output is too broad.

---

## 📋 Changelog

### v1.0 — Initial release
- VSS Shadow Copy — file lock bypass for running browsers
- 16 browsers, all local user profiles via registry enumeration
- Dual parse mode: full SQL (`sqlite3.exe`) and regex fallback
- AES-256 encrypted ZIP via `7z.exe` — 16-char cryptographic password to console
- `-NoArchive` — plain file output for environments without 7-Zip
- `-DateFrom` / `-DateTo` — date range filtering at collection level
- `-CsvExport` — parallel CSV output
- `Write-Progress` + `[CmdletBinding()]` — console progress bar, `-Verbose` support
- HTML: interactive report with search, filters, column sorting
- HTML: in-browser CSV export (BOM UTF-8), visit count bar, tooltips, favicon

---

## `>_` disclaimer

> Intended for authorized forensic analysis and incident response on systems you have explicit permission to access. The author assumes no responsibility for use outside these boundaries.

---

<div align="center">

**ZavetSec** — security tooling for those who read logs at 2am

[github.com/zavetsec](https://github.com/zavetsec) · MIT License

</div>
