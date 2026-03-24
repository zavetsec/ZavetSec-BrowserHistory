<div align="center">

# `>_` ZavetSec-BrowserHistory

**Browser history acquisition for Windows IR — all users, all browsers, one encrypted report**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)](https://github.com/zavetsec)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078d4?logo=windows&logoColor=white)](https://github.com/zavetsec)
[![No Dependencies](https://img.shields.io/badge/Dependencies-none%20required-00d4ff)](https://github.com/zavetsec)
[![License: MIT](https://img.shields.io/badge/License-MIT-green)](https://github.com/zavetsec)

Bypasses file locks via VSS Shadow Copy · 16 browsers · AES-256 encrypted output · interactive HTML report

[When to Use](#-when-to-use) · [Quick Start](#-quick-start) · [Parameters](#-parameters) · [Report](#-html-report) · [Remote Execution](#-remote-execution) · [Limitations](#-limitations) · [Roadmap](#-roadmap)

</div>

---

## TL;DR

One-command browser history acquisition for live IR on Windows:

```powershell
.\ZavetSec-BrowserHistory.ps1
```

- Bypasses locked browser databases via VSS Shadow Copy — browser stays running
- Collects all local user profiles in a single pass
- Produces an encrypted, analyst-ready HTML report with search and filters
- Zero installation — single `.ps1` file, PowerShell 5.1, any Windows endpoint

> 🚧 **Upcoming:** inline suspicious domain tagging — C2s, phishing domains, paste sites, and cloud exfil targets flagged directly in the report

---

## Why this tool

There are several browser history tools out there. This one is built specifically for the IR scenario — not casual browsing analysis or digital parenting tools.

The differences that matter in the field:

- **VSS Shadow Copy** — reads `History` and `places.sqlite` directly from a running browser without terminating it or waiting for a clean close. Most tools fail silently or return partial results when the browser holds a lock.
- **All users in one run** — enumerates every local profile via the registry (`ProfileList`), covering domain accounts, local accounts, and accounts that have never logged in during this session. One command, one machine, complete picture.
- **Encrypted evidence packaging** — output is packed into an AES-256 ZIP with a one-time cryptographically random password printed once to console. Safe to copy over a jump host, drop on a USB, or attach to a ticket without exposing the contents in transit.
- **Zero installation** — a single `.ps1` file. Runs on any Windows endpoint with PowerShell 5.1 — every corporate machine since Windows 7. No MSI, no NuGet, no setup.
- **Built-in fallback chain** — no `sqlite3.exe`? Opportunistically recovers URLs from raw SQLite pages. No `7z.exe`? Built-in .NET compression still produces a ZIP. No Admin rights? Current user profile is still collected.
- **Get-Help native** — fully documented via PowerShell comment-based help. `Get-Help .\ZavetSec-BrowserHistory.ps1 -Full` works out of the box.

**Where it fits in an IR workflow:**
Deploy early — before or alongside a KAPE collection, and before imaging. It handles the browser artifact layer that KAPE targets frequently miss on live systems due to file locking. The structured CSV output ingests directly into a SIEM or timeline tool. The encrypted archive can be handed off at any point without breaking evidence handling procedures.

---

## When to Use

**Incident Response / DFIR**
Host is suspected of compromise. You need browser evidence — what was visited, when, by which account — before reimaging or network isolation cuts off access. Run the script, collect the encrypted archive, proceed.

**Insider Threat Investigation**
An account shows anomalous behavior. Browser history across all local profiles can confirm or rule out data staging via webmail, cloud storage uploads, or paste sites — evidence that process logs alone won't surface.

**Workstation Audit**
Pre-termination review, compliance check, or post-incident reconstruction. Covers all profiles in a single run without manual per-account access.

**When not to use**
This is a live-acquisition triage tool. It does not preserve original file metadata, does not compute cryptographic hashes of source database files, and does not produce chain-of-custody documentation. Do not use as the sole evidence source in formal legal proceedings — pair it with a proper forensic imager.

---

## Features

- **16 browsers** — covers major Chromium and Gecko forks commonly observed in enterprise environments and threat actor tradecraft: Chrome, Edge, Firefox, Brave, Opera, Opera GX, Yandex, Vivaldi, Tor Browser, Thunderbird, Waterfox, LibreWolf, Pale Moon, Epic, Comodo Dragon, OneDrive WebView
- **All users** — full registry-based profile enumeration, not just the active session
- **File lock bypass** — VSS Shadow Copy reads locked databases from running browsers
- **Dual parse mode** — full SQL via `sqlite3.exe` (titles, visit counts, accurate timestamps) or regex fallback with zero dependencies
- **Accurate timestamps** — correct epoch conversion for both Chromium (1601-01-01 base) and Firefox (Unix epoch) timestamp formats
- **Encrypted output** — AES-256 ZIP via `7z.exe` + `7z.dll`, 16-char cryptographic password to console once
- **Interactive HTML report** — search, filter by user / browser / date range, column sorting, in-browser CSV export
- **Remote-execution ready** — safe `$ScriptDir` resolution for PsExec / WinRM / SYSTEM context; `-OpenReport` auto-suppressed when no interactive desktop is available
- **Native help** — full `Get-Help` support with examples

---

## ⚡ Quick Start

```powershell
# View built-in help
Get-Help .\ZavetSec-BrowserHistory.ps1 -Full
.\ZavetSec-BrowserHistory.ps1 -Help

# Standard run — encrypted ZIP + password to console (Administrator required)
.\ZavetSec-BrowserHistory.ps1

# Open report immediately after generation
.\ZavetSec-BrowserHistory.ps1 -OpenReport

# Date-scoped collection with CSV export
.\ZavetSec-BrowserHistory.ps1 -CsvExport -DateFrom 2025-01-01 -DateTo 2025-06-30

# No archive — plain HTML (use when 7z.exe + 7z.dll are not available)
.\ZavetSec-BrowserHistory.ps1 -NoArchive -OpenReport

# Collect all records, no limit
.\ZavetSec-BrowserHistory.ps1 -MaxRecordsPerBrowser 999999

# Save directly to a network share
.\ZavetSec-BrowserHistory.ps1 -NoArchive -OutputPath "\\server\IR\host42_history.html"
```

> Without arguments the report is saved to `.\Reports\<HOSTNAME>_<TIMESTAMP>.zip`

---

## 📋 Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-OutputPath` | String | auto | Path for the HTML report. Auto-generated in `.\Reports\` if omitted |
| `-MaxRecordsPerBrowser` | Int | `5000` | Maximum records per browser profile. Use `999999` for unlimited |
| `-OpenReport` | Switch | — | Open the report after generation. Skipped silently in remote sessions |
| `-CsvExport` | Switch | — | Save results as CSV alongside the HTML |
| `-DateFrom` | String | — | Collect records from this date (`yyyy-MM-dd`) |
| `-DateTo` | String | — | Collect records up to this date (`yyyy-MM-dd`) |
| `-NoArchive` | Switch | — | Skip ZIP — save HTML (and CSV) as plain files |
| `-Help` | Switch | — | Show help and exit |

---

## 📊 HTML Report

Single self-contained `.html` file — no server, no CDN, no internet required. Opens on an air-gapped analyst workstation.

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
| 1 | john | Edge | github.com | GitHub · Build... | ▬▬▬ 42 | 2026-03-24 09:31 |

**Filters:** user · browser · date range (From / To) · full-text search · `⬓ CSV` export of visible rows (BOM UTF-8, Excel-compatible) · click any column header to sort

---

## 🔑 Archive Password

Generated fresh on every run via `System.Security.Cryptography.RandomNumberGenerator`. Printed once to console on completion — never written to disk.

```
  ================================================================
   REPORT READY
  ================================================================

   Archive  : .\Reports\HOSTNAME_20260324_093100.zip
   Encrypt  : AES-256

   Password : TJewuHDDC9phPTYN

   Save this password - it will not be shown again.
  ================================================================
```

Password alphabet is alphanumeric only — no special characters that could cause shell escaping issues when passing the password to `7z.exe`.

If `7z.exe` + `7z.dll` are not found, the script falls back to an unencrypted ZIP with a warning. Use `-NoArchive` to skip archiving entirely.

---

## 🔧 Setup

The script runs with no external files. Optional tools improve output quality:

| File | Effect |
|---|---|
| `sqlite3.exe` | Full SQL parse: titles + visit counts + accurate timestamps. Without it: URL-only regex fallback |
| `7z.exe` + `7z.dll` | AES-256 encrypted ZIP. Both files required — `7z.exe` alone will not work |

Both are included in the [release archive](https://github.com/zavetsec/ZavetSec-BrowserHistory/releases/latest). Download, extract everything to one folder, run.

To obtain manually:
- `sqlite3.exe` — [sqlite.org/download.html](https://sqlite.org/download.html) → `sqlite-tools-win-x64-*.zip`
- `7z.exe` + `7z.dll` — copy both from an existing 7-Zip installation (`C:\Program Files\7-Zip\`)

```
ZavetSec-BrowserHistory.ps1
sqlite3.exe        ← included in release archive
7z.exe             ← included in release archive
7z.dll             ← included in release archive (required — 7z.exe will not work without it)
Reports\           ← created automatically
```

---

## 🛡️ How It Works

```
Running as Administrator?
        │
        ├─ Yes → VSS Shadow Copy of C:\
        │        Reads locked History / places.sqlite from the shadow
        │        Browser stays running, no interruption
        │
        └─ No  → Direct copy fallback
                 ReadAllBytes → Copy-Item → xcopy
                 Running browsers may have their DB files locked

sqlite3.exe present?
        │
        ├─ Yes → SQL query: URL + Title + VisitCount + Timestamp
        │        Chromium: microseconds since 1601-01-01 (correct epoch)
        │        Firefox:  microseconds since 1970-01-01 (Unix epoch)
        │
        └─ No  → Regex scan of raw database bytes: URL only

7z.exe + 7z.dll present?
        │
        ├─ Yes → AES-256 encrypted ZIP, 16-char password to console once
        │
        └─ No  → Unencrypted ZIP via built-in .NET System.IO.Compression
```

---

## 🌐 Supported Browsers

| Engine | Browsers |
|---|---|
| **Chromium** | Google Chrome, Microsoft Edge, Brave, Yandex Browser, Opera, Opera GX, Vivaldi, Epic Privacy Browser, Comodo Dragon, OneDrive WebView, Win WebExperience |
| **Firefox / Gecko** | Mozilla Firefox, Tor Browser, Thunderbird, Waterfox, LibreWolf, Pale Moon |

---

## 🖥️ Remote Execution

The script is designed to run safely in non-interactive contexts.

**PsExec (as SYSTEM):**
```powershell
psexec \\TARGET -s powershell.exe -ExecutionPolicy Bypass -File "\\share\ZavetSec-BrowserHistory.ps1" -NoArchive -OutputPath "\\share\output\TARGET_history.html"
```

**WinRM / Invoke-Command:**
```powershell
Invoke-Command -ComputerName TARGET -FilePath .\ZavetSec-BrowserHistory.ps1 -ArgumentList @{NoArchive=$true; OutputPath="\\share\output\TARGET_history.html"}
```

**Remote execution notes:**
- `-OpenReport` is automatically suppressed when no interactive desktop is detected (PsExec SYSTEM, WinRM) — no error, just a logged skip
- `$ScriptDir` resolves correctly in all contexts: local, PsExec, WinRM, scheduled task
- For remote runs, use `-NoArchive` with a UNC `OutputPath`, or collect the ZIP and password separately

---

## 📁 Output

**Default (with archiving):**
```
.\Reports\
└── HOSTNAME_20260324_093100.zip    ← AES-256 encrypted
```
Source `.html` and `.csv` are deleted after successful archiving.

**With `-NoArchive`:**
```
.\Reports\
├── HOSTNAME_20260324_093100.html
└── HOSTNAME_20260324_093100.csv    ← with -CsvExport
```

**CSV columns:** `UserName, Browser, Domain, Title, URL, Visits, LastVisit`

---

## ⚠️ Limitations

- **Live acquisition only** — reads from a running system; does not image or hash source database files
- **No chain of custody** — not a substitute for forensic imaging (FTK, dd, KAPE) in legal contexts
- **Standard profile paths only** — portable browser installs or non-default profile locations are not detected
- **C:\ drive only** — VSS shadow copy targets the system drive; profiles on other volumes fall back to direct copy
- **No deletion detection** — does not identify history gaps, SQLite free pages, or anti-forensic clearing activity
- **Regex fallback is approximate** — without `sqlite3.exe`, results are opportunistically recovered from raw SQLite pages and may contain duplicates or partial URLs from internal database structures

---

## ❓ FAQ

**Can it run without Administrator rights?**
Yes — the script runs and collects the current user's profile. VSS is unavailable, other user profiles are inaccessible. For full collection, Administrator is required.

**The browser is running and the database is locked — does it still work?**
Yes, when running as Administrator. VSS shadow copy reads the database from a point-in-time snapshot; the lock is irrelevant.

**7z.exe is in the folder but encryption still fails?**
Both `7z.exe` and `7z.dll` are required. `7z.dll` provides the codec support — without it, 7-Zip cannot process ZIP format and exits with an error. Copy both files from `C:\Program Files\7-Zip\`.

**Why is `-MaxRecordsPerBrowser` limited to 5000 by default?**
Conservative default to keep report generation fast and file sizes manageable on machines with years of history. Use `-MaxRecordsPerBrowser 999999` to collect everything.

**How does this relate to ZavetSec Triage?**
`Invoke-ZavetSecTriage.ps1` includes browser history as module #14 alongside 17 other collection modules. This script is the standalone version — use it when you need browser evidence only, or when running the full triage is not practical.

**Does it work alongside KAPE?**
Yes. Run before or alongside KAPE collection. This script handles the file-lock problem that KAPE browser targets frequently encounter on live systems. The `-CsvExport` output ingests directly into most timeline tools.

---

## 🗺️ Roadmap

- [ ] **🚧 Suspicious domain tagging** ← *next* — flag known C2 infrastructure, phishing domains, paste sites, and cloud exfil targets inline in the report without leaving the analyst workstation
- [ ] **IOC export** — one-click export of all collected domains and URLs as a flat IOC list for TIP / SIEM ingestion
- [ ] **Tor / anonymizer detection** — surface visits to .onion proxies, anonymizer services, and VPN provider pages
- [ ] **Timeline view** — unified chronological view across all users and browsers in a single scrollable timeline
- [ ] **Multi-drive VSS** — extend shadow copy acquisition beyond C:\
- [ ] **Download history** — collect Chromium download records alongside browsing history

---

## 🤝 Contributing

Most useful contributions:

- **New browser paths** — if a browser is not detected, open an issue with the path to its `History` or `places.sqlite`
- **Bug reports** — unexpected behavior on specific Windows versions, domain configurations, or profile setups
- **Suspicious domain lists** — curated C2, phishing, and exfil domain lists for the planned tagging feature

Requirements for PRs: PowerShell 5.1 compatible, tested on real Windows, zero-dependency guarantee preserved for the core collection path.

[Open an issue](https://github.com/zavetsec/ZavetSec-BrowserHistory/issues)

---

## 📋 Changelog

### v1.0 — Initial release
- VSS Shadow Copy — file lock bypass for running browsers
- 16 browsers, all local user profiles via registry enumeration
- Dual parse mode: full SQL (`sqlite3.exe`) and regex fallback
- Accurate timestamp conversion for both Chromium (1601 epoch) and Firefox (Unix epoch)
- AES-256 encrypted ZIP via `7z.exe` + `7z.dll` — 16-char cryptographic password to console
- `-NoArchive` — plain file output for environments without 7-Zip
- `-DateFrom` / `-DateTo` — date range filtering at collection level and in report
- `-CsvExport` — parallel CSV output alongside HTML
- `-Help` switch + full `Get-Help` comment-based documentation
- Remote execution support: robust `$ScriptDir` resolution, `-OpenReport` auto-suppressed in non-interactive sessions
- `Write-Progress` + `[CmdletBinding()]` — console progress, `-Verbose` support
- HTML: interactive report — search, filter, sort, date range picker, in-browser CSV export
- HTML: visit count mini-bar, URL and domain tooltips, browser favicon, improved contrast palette

---

## `>_` disclaimer

> Intended for authorized forensic analysis and incident response on systems you have explicit permission to access. The author assumes no responsibility for use outside these boundaries.

---

<div align="center">

**ZavetSec** — security tooling for those who read logs at 2am

[github.com/zavetsec](https://github.com/zavetsec) · MIT License

</div>
