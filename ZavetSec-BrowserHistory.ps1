# ============================================================
#  ZavetSec-BrowserHistory.ps1  v1.0
#  Extracts browser history for ALL local user profiles
#  via VSS shadow copy (bypasses file locks)
#
#  Requires: Run as Administrator
#            7z.exe + 7z.dll in same folder (optional, enables AES-256 ZIP)
#            sqlite3.exe in same folder (optional, enables full browser parse)
#
#  Parameters:
#    -OutputPath            Path for the HTML report (auto-generated if omitted)
#    -MaxRecordsPerBrowser  Max records to extract per browser profile (default: 5000)
#    -OpenReport            Open the HTML report in browser after generation
#    -CsvExport             Also save results as CSV alongside the HTML
#    -DateFrom              Filter records from this date (yyyy-MM-dd)
#    -DateTo                Filter records up to this date  (yyyy-MM-dd)
#    -NoArchive             Skip ZIP archiving - save HTML (and CSV) as plain files
#
#  Examples:
#    .\ZavetSec-BrowserHistory.ps1
#    .\ZavetSec-BrowserHistory.ps1 -OpenReport -CsvExport
#    .\ZavetSec-BrowserHistory.ps1 -DateFrom 2025-01-01 -DateTo 2025-06-30 -CsvExport
#    .\ZavetSec-BrowserHistory.ps1 -NoArchive -OpenReport
#
#  https://github.com/zavetsec
# ============================================================

#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputPath          = "",
    [int]$MaxRecordsPerBrowser   = 5000,
    [switch]$OpenReport,
    # Optional: also save results as CSV alongside the HTML report
    [switch]$CsvExport,
    # Optional: filter records by date range (ISO format: 2025-01-01)
    [string]$DateFrom            = "",
    [string]$DateTo              = "",
    # Optional: skip ZIP archiving, save HTML (and CSV) as plain files
    [switch]$NoArchive
)

$ErrorActionPreference = "SilentlyContinue"

# Validate date parameters early
$filterDateFrom = $null
$filterDateTo   = $null
if ($DateFrom -ne "") {
    try { $filterDateFrom = [datetime]::Parse($DateFrom) }
    catch { Write-Warning "Invalid -DateFrom '$DateFrom'. Use format: yyyy-MM-dd. Filter ignored." }
}
if ($DateTo -ne "") {
    try { $filterDateTo = [datetime]::Parse($DateTo).AddDays(1).AddSeconds(-1) }
    catch { Write-Warning "Invalid -DateTo '$DateTo'. Use format: yyyy-MM-dd. Filter ignored." }
}
Add-Type -AssemblyName System.Web

# --- Resolve script directory ---
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# --- Resolve output path: Reports subfolder next to script ---
if ([string]::IsNullOrEmpty($OutputPath)) {
    $ReportsDir = Join-Path $ScriptDir "Reports"
    if (-not (Test-Path $ReportsDir)) {
        New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null
    }
    $TimeStamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $HostTag    = $env:COMPUTERNAME
    $OutputPath = Join-Path $ReportsDir "${HostTag}_${TimeStamp}.html"
}

# --- Resolve sqlite3.exe: same folder as script OR PATH ---
$Sqlite3 = $null
$candidates = @(
    (Join-Path $ScriptDir "sqlite3.exe"),
    (Join-Path $ScriptDir "sqlite3\sqlite3.exe")
)
$env:PATH -split ';' | Where-Object { $_ } | ForEach-Object {
    $candidates += (Join-Path $_ "sqlite3.exe")
}
foreach ($c in $candidates) {
    if ($c -and (Test-Path $c)) { $Sqlite3 = (Resolve-Path $c).Path; break }
}

Write-Host ""
Write-Host "  [*] Browser History Extractor  [ALL USERS]" -ForegroundColor Cyan
Write-Host "  [*] Script dir : $ScriptDir"                -ForegroundColor DarkGray
if ($Sqlite3) {
    Write-Host "  [+] sqlite3.exe : $Sqlite3" -ForegroundColor Green
} else {
    Write-Host "  [!] sqlite3.exe : NOT FOUND - falling back to regex extraction" -ForegroundColor Yellow
    Write-Host "      Place sqlite3.exe in: $ScriptDir"   -ForegroundColor DarkGray
}
Write-Host ""

# ============================================================
#  VSS SHADOW COPY - bypass file locks
# ============================================================
$ShadowPath = $null
$ShadowLink = "$env:TEMP\BHE_Shadow_$([System.IO.Path]::GetRandomFileName())"
$VssCleanup = $false

function New-ShadowCopy {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "  [!] Not running as Administrator - VSS unavailable, using direct copy" -ForegroundColor Yellow
        return $false
    }

    Write-Host "  [*] Creating VSS shadow copy of C:\ ..." -ForegroundColor Cyan
    try {
        $shadow = (Get-WmiObject -List Win32_ShadowCopy).Create("C:\", "ClientAccessible")
        if ($shadow.ReturnValue -ne 0) {
            Write-Host "  [!] VSS creation failed (code $($shadow.ReturnValue))" -ForegroundColor Yellow
            return $false
        }
        $id         = $shadow.ShadowID
        $sc         = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $id }
        $devicePath = $sc.DeviceObject + "\"
        cmd /c "mklink /d `"$ShadowLink`" `"$devicePath`"" 2>$null | Out-Null
        if (Test-Path $ShadowLink) {
            $script:ShadowPath = $ShadowLink
            $script:VssCleanup = $true
            $script:ShadowId   = $id
            Write-Host "  [+] Shadow copy mounted: $ShadowLink" -ForegroundColor Green
            return $true
        }
    } catch {
        Write-Host "  [!] VSS error: $_" -ForegroundColor Yellow
    }
    return $false
}

function Remove-ShadowCopy {
    if (-not $script:VssCleanup) { return }
    try {
        cmd /c "rmdir `"$ShadowLink`"" 2>$null | Out-Null
        $sc = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $script:ShadowId }
        if ($sc) { $sc.Delete() }
        Write-Host "  [*] VSS shadow copy removed" -ForegroundColor DarkGray
    } catch {}
}

function Get-ShadowedPath {
    param([string]$Path)
    if ($script:ShadowPath) {
        $rel = $Path -replace '^[A-Za-z]:\\', ''
        return Join-Path $script:ShadowPath $rel
    }
    return $Path
}

# ============================================================
#  FILE COPY - try multiple methods
# ============================================================
function Copy-DbFile {
    param([string]$Source)

    $realSource = Get-ShadowedPath $Source
    if (-not (Test-Path $realSource)) { return $null }

    $tmp = [System.IO.Path]::GetTempFileName() + ".db"

    try {
        $bytes = [System.IO.File]::ReadAllBytes($realSource)
        [System.IO.File]::WriteAllBytes($tmp, $bytes)
        if ((Get-Item $tmp).Length -gt 0) { return $tmp }
    } catch {}

    try {
        Copy-Item -Path $realSource -Destination $tmp -Force
        if (Test-Path $tmp) { return $tmp }
    } catch {}

    try {
        $dir   = Split-Path $tmp
        $fname = Split-Path $tmp -Leaf
        cmd /c "xcopy /Y /Q `"$realSource`" `"$dir\`"" 2>$null | Out-Null
        $copied = Join-Path $dir (Split-Path $realSource -Leaf)
        if (Test-Path $copied) {
            Rename-Item $copied $fname -Force
            if (Test-Path $tmp) { return $tmp }
        }
    } catch {}

    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    return $null
}

# ============================================================
#  QUERY SQLITE - explicit UTF-8 to preserve Cyrillic
# ============================================================
function Invoke-Sqlite {
    param([string]$DbPath, [string]$Query)
    if (-not $script:Sqlite3) { return $null }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = $script:Sqlite3
        $psi.Arguments              = "`"$DbPath`" `"$Query`""
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8

        $proc = [System.Diagnostics.Process]::Start($psi)
        $out  = $proc.StandardOutput.ReadToEnd()
        $proc.WaitForExit()

        return $out -split "`n" | Where-Object { $_ -ne "" }
    } catch {
        return $null
    }
}

# ============================================================
#  CHROMIUM HISTORY
# ============================================================
function Get-ChromiumHistory {
    param([string]$ProfilePath, [string]$BrowserName, [string]$UserName, [int]$Limit)

    $histFile = Join-Path $ProfilePath "History"
    if (-not (Test-Path (Get-ShadowedPath $histFile)) -and -not (Test-Path $histFile)) { return @() }

    $tmp = Copy-DbFile $histFile
    if (-not $tmp) {
        Write-Host "  [!] Could not copy: $histFile" -ForegroundColor Red
        return @()
    }

    $records = @()
    try {
        if ($script:Sqlite3) {
            $query = "SELECT url, title, visit_count, last_visit_time FROM urls ORDER BY last_visit_time DESC LIMIT $Limit;"
            $rows  = Invoke-Sqlite $tmp $query
            foreach ($row in $rows) {
                if ([string]::IsNullOrWhiteSpace($row)) { continue }
                $p  = $row -split '\|'
                if ($p.Count -lt 2) { continue }
                $ts = if ($p.Count -ge 4 -and $p[3]) { try { [int64]$p[3] } catch { 0 } } else { 0 }
                $dt = if ($ts -gt 0) {
                    try { [datetime]::FromFileTimeUtc(($ts - 11644473600000000) * 10) } catch { [datetime]::MinValue }
                } else { [datetime]::MinValue }
                $records += [PSCustomObject]@{
                    UserName  = $UserName
                    Browser   = $BrowserName
                    URL       = $p[0]
                    Title     = if ($p.Count -ge 2 -and $p[1]) { $p[1] } else { $p[0] }
                    Visits    = if ($p.Count -ge 3) { try { [int]$p[2] } catch { 1 } } else { 1 }
                    LastVisit = $dt
                    Domain    = try { ([System.Uri]$p[0]).Host } catch { "unknown" }
                }
            }
        }

        if ($records.Count -eq 0) {
            $content  = [System.IO.File]::ReadAllText($tmp, [System.Text.Encoding]::Latin1)
            $matches_ = [regex]::Matches($content, 'https?://[^\x00-\x1F\x7F"<> ]{5,500}')
            $seen = @{}
            foreach ($m in $matches_) {
                $url = $m.Value.TrimEnd(".,;)'`"\\")
                if ($seen[$url] -or $records.Count -ge $Limit) { continue }
                $seen[$url] = $true
                $records += [PSCustomObject]@{
                    UserName  = $UserName
                    Browser   = $BrowserName
                    URL       = $url
                    Title     = $url
                    Visits    = 1
                    LastVisit = [datetime]::MinValue
                    Domain    = try { ([System.Uri]$url).Host } catch { "unknown" }
                }
            }
        }
    } finally {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
    return $records
}

# ============================================================
#  FIREFOX HISTORY
# ============================================================
function Get-FirefoxHistory {
    param([string]$ProfilePath, [string]$BrowserName, [string]$UserName, [int]$Limit)

    $histFile = Join-Path $ProfilePath "places.sqlite"
    if (-not (Test-Path (Get-ShadowedPath $histFile)) -and -not (Test-Path $histFile)) { return @() }

    $tmp = Copy-DbFile $histFile
    if (-not $tmp) { return @() }

    $records = @()
    try {
        if ($script:Sqlite3) {
            $query = "SELECT p.url, COALESCE(p.title,''), p.visit_count, MAX(h.visit_date) FROM moz_places p LEFT JOIN moz_historyvisits h ON p.id=h.place_id WHERE p.hidden=0 GROUP BY p.id ORDER BY MAX(h.visit_date) DESC LIMIT $Limit;"
            $rows  = Invoke-Sqlite $tmp $query
            foreach ($row in $rows) {
                if ([string]::IsNullOrWhiteSpace($row)) { continue }
                $p  = $row -split '\|'
                if ($p.Count -lt 1) { continue }
                $ts = if ($p.Count -ge 4 -and $p[3]) { try { [int64]$p[3] } catch { 0 } } else { 0 }
                $dt = if ($ts -gt 0) {
                    try { [datetime]::UnixEpoch.AddMicroseconds($ts) } catch { [datetime]::MinValue }
                } else { [datetime]::MinValue }
                $records += [PSCustomObject]@{
                    UserName  = $UserName
                    Browser   = $BrowserName
                    URL       = $p[0]
                    Title     = if ($p.Count -ge 2 -and $p[1]) { $p[1] } else { $p[0] }
                    Visits    = if ($p.Count -ge 3) { try { [int]$p[2] } catch { 1 } } else { 1 }
                    LastVisit = $dt
                    Domain    = try { ([System.Uri]$p[0]).Host } catch { "unknown" }
                }
            }
        }

        if ($records.Count -eq 0) {
            $content  = [System.IO.File]::ReadAllText($tmp, [System.Text.Encoding]::Latin1)
            $matches_ = [regex]::Matches($content, 'https?://[^\x00-\x1F\x7F"<> ]{5,500}')
            $seen = @{}
            foreach ($m in $matches_) {
                $url = $m.Value.TrimEnd(".,;)'`"\\")
                if ($seen[$url] -or $records.Count -ge $Limit) { continue }
                $seen[$url] = $true
                $records += [PSCustomObject]@{
                    UserName  = $UserName
                    Browser   = $BrowserName
                    URL       = $url
                    Title     = $url
                    Visits    = 1
                    LastVisit = [datetime]::MinValue
                    Domain    = try { ([System.Uri]$url).Host } catch { "unknown" }
                }
            }
        }
    } finally {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
    return $records
}

# ============================================================
#  BROWSER DEFINITIONS
#  Tokens: {L}=AppData\Local  {R}=AppData\Roaming  {H}=profile root
# ============================================================
$browserDefs = @(
    @{ Name="Chromium";          Type="Chromium"; Rel=@("{L}\Chromium\User Data\Default","{L}\Chromium\User Data\Profile 1") },
    @{ Name="Google Chrome";     Type="Chromium"; Rel=@("{L}\Google\Chrome\User Data\Default","{L}\Google\Chrome\User Data\Profile 1","{L}\Google\Chrome\User Data\Profile 2") },
    @{ Name="Microsoft Edge";    Type="Chromium"; Rel=@("{L}\Microsoft\Edge\User Data\Default","{L}\Microsoft\Edge\User Data\Profile 1") },
    @{ Name="Brave";             Type="Chromium"; Rel=@("{L}\BraveSoftware\Brave-Browser\User Data\Default") },
    @{ Name="Yandex Browser";    Type="Chromium"; Rel=@("{L}\Yandex\YandexBrowser\User Data\Default") },
    @{ Name="Opera";             Type="Chromium"; Rel=@("{R}\Opera Software\Opera Stable") },
    @{ Name="Opera GX";          Type="Chromium"; Rel=@("{R}\Opera Software\Opera GX Stable") },
    @{ Name="Vivaldi";           Type="Chromium"; Rel=@("{L}\Vivaldi\User Data\Default") },
    @{ Name="Epic Browser";      Type="Chromium"; Rel=@("{L}\Epic Privacy Browser\User Data\Default") },
    @{ Name="Comodo Dragon";     Type="Chromium"; Rel=@("{L}\Comodo\Dragon\User Data\Default") },
    @{ Name="OneDrive WebView";  Type="Chromium"; Rel=@("{L}\Microsoft\OneDrive\EBWebView\Default") },
    @{ Name="Win WebExperience"; Type="Chromium"; Rel=@("{L}\Packages\MicrosoftWindows.Client.WebExperience_cw5n1h2txyewy\LocalState\EBWebView\Default") },
    @{ Name="Mozilla Firefox";   Type="Firefox";  Rel=@("{R}\Mozilla\Firefox\Profiles") },
    @{ Name="Thunderbird";       Type="Firefox";  Rel=@("{R}\Thunderbird\Profiles") },
    @{ Name="Tor Browser";       Type="Firefox";  Rel=@("{R}\Tor Browser\Browser\TorBrowser\Data\Browser\profile.default","{H}\Desktop\Tor Browser\Browser\TorBrowser\Data\Browser\profile.default") },
    @{ Name="Waterfox";          Type="Firefox";  Rel=@("{R}\Waterfox\Profiles") },
    @{ Name="LibreWolf";         Type="Firefox";  Rel=@("{R}\LibreWolf\Profiles") },
    @{ Name="Pale Moon";         Type="Firefox";  Rel=@("{R}\Moonchild Productions\Pale Moon\Profiles") }
)

# ============================================================
#  ENUMERATE LOCAL USER PROFILES
# ============================================================
$systemNames = @('Public','Default','Default User','All Users','defaultuser0','desktop.ini')

function Get-UserProfiles {
    $profiles = @()

    # Primary: registry ProfileList (covers domain users and roaming profiles)
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $keys = Get-ChildItem $regPath -ErrorAction SilentlyContinue
    foreach ($key in $keys) {
        $profPath = (Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
        if (-not $profPath) { continue }
        $uname = Split-Path $profPath -Leaf
        if ($systemNames -contains $uname) { continue }
        if ($uname -match '^(SYSTEM|NETWORK SERVICE|LOCAL SERVICE)$') { continue }
        if (Test-Path $profPath) {
            $profiles += [PSCustomObject]@{
                UserName    = $uname
                ProfilePath = $profPath
                Local       = Join-Path $profPath "AppData\Local"
                Roaming     = Join-Path $profPath "AppData\Roaming"
            }
        }
    }

    # Fallback: scan C:\Users directly
    if ($profiles.Count -eq 0) {
        $usersRoot = "$env:SystemDrive\Users"
        Get-ChildItem $usersRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            if ($systemNames -contains $_.Name) { return }
            $loc = Join-Path $_.FullName "AppData\Local"
            $rom = Join-Path $_.FullName "AppData\Roaming"
            if (Test-Path $loc) {
                $profiles += [PSCustomObject]@{
                    UserName    = $_.Name
                    ProfilePath = $_.FullName
                    Local       = $loc
                    Roaming     = $rom
                }
            }
        }
    }

    return $profiles
}

# ============================================================
#  MAIN COLLECTION
# ============================================================
$vssOk = New-ShadowCopy

$userProfiles = Get-UserProfiles
Write-Host "  [*] User profiles found: $($userProfiles.Count)" -ForegroundColor Cyan
foreach ($up in $userProfiles) {
    Write-Host "      -> $($up.UserName)  ($($up.ProfilePath))" -ForegroundColor DarkGray
}
Write-Host ""

$allRecords = @()
$userStats  = @{}   # key: username -> total records
$uIdx = 0

foreach ($up in $userProfiles) {

    $uIdx++
    Write-Progress -Activity "Collecting browser history" `
        -Status "User: $($up.UserName)  ($uIdx / $($userProfiles.Count))" `
        -PercentComplete ([int]($uIdx / [Math]::Max($userProfiles.Count,1) * 100))

    Write-Host "  [USER] $($up.UserName)" -ForegroundColor Yellow
    $userTotal = 0

    foreach ($def in $browserDefs) {

        # Substitute tokens with actual paths for this user
        $resolvedPaths = @()
        foreach ($rel in $def.Rel) {
            $resolvedPaths += $rel `
                -replace '\{L\}', $up.Local `
                -replace '\{R\}', $up.Roaming `
                -replace '\{H\}', $up.ProfilePath
        }

        $collected = @()

        if ($def.Type -eq "Firefox") {
            foreach ($basePath in $resolvedPaths) {
                $shadowBase = Get-ShadowedPath $basePath
                $testPath   = if (Test-Path $shadowBase) { $shadowBase } `
                              elseif (Test-Path $basePath) { $basePath } `
                              else { $null }
                if (-not $testPath) { continue }

                $profileDirs = if ($testPath -match "Profiles$") {
                    Get-ChildItem $testPath -Directory -ErrorAction SilentlyContinue
                } else {
                    @([System.IO.DirectoryInfo]$testPath)
                }
                foreach ($dir in $profileDirs) {
                    $origDir = $dir.FullName
                    if ($script:ShadowPath) {
                        $origDir = $dir.FullName -replace [regex]::Escape($script:ShadowPath + "\"), "C:\"
                    }
                    $recs = Get-FirefoxHistory -ProfilePath $origDir -BrowserName $def.Name -UserName $up.UserName -Limit $MaxRecordsPerBrowser
                    $collected += $recs
                }
            }
        } else {
            foreach ($profilePath in $resolvedPaths) {
                $recs = Get-ChromiumHistory -ProfilePath $profilePath -BrowserName $def.Name -UserName $up.UserName -Limit $MaxRecordsPerBrowser
                $collected += $recs
            }
        }

        $unique = @($collected | Sort-Object URL -Unique)
        $count  = $unique.Count

        if ($count -gt 0) {
            Write-Host "    [+] $($def.Name.PadRight(22)) $count records" -ForegroundColor Green
            $allRecords += $unique
            $userTotal  += $count
        } else {
            Write-Host "    [-] $($def.Name.PadRight(22)) not found / empty" -ForegroundColor DarkGray
        }
    }

    if ($userTotal -gt 0) {
        $userStats[$up.UserName] = $userTotal
        Write-Host "    [=] Total for $($up.UserName): $userTotal records" -ForegroundColor Cyan
    }
    Write-Host ""
}

Remove-ShadowCopy

Write-Progress -Activity "Collecting browser history" -Completed

# Apply optional date filter
if ($filterDateFrom -or $filterDateTo) {
    $before = $allRecords.Count
    $allRecords = @($allRecords | Where-Object {
        if ($_.LastVisit -eq [datetime]::MinValue) { return $true }  # keep undated
        $dt = $_.LastVisit
        $okFrom = -not $filterDateFrom -or $dt -ge $filterDateFrom
        $okTo   = -not $filterDateTo   -or $dt -le $filterDateTo
        $okFrom -and $okTo
    })
    Write-Host "  [*] Date filter applied: $before -> $($allRecords.Count) records" -ForegroundColor DarkGray
}

Write-Host "  [*] Total records collected: $($allRecords.Count)" -ForegroundColor Yellow
Write-Host "  [*] Building HTML report..."                        -ForegroundColor Cyan

# ============================================================
#  HTML GENERATION
# ============================================================

$rowsHtml = ""
$idx = 1
$sorted = $allRecords | Sort-Object LastVisit -Descending

# Pre-calculate max visits for bar scaling
$maxVisits = ($allRecords | Measure-Object -Property Visits -Maximum).Maximum
if (-not $maxVisits -or $maxVisits -eq 0) { $maxVisits = 1 }

foreach ($r in $sorted) {
    $dateStr     = if ($r.LastVisit -ne [datetime]::MinValue) { $r.LastVisit.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "-" }
    $dateIso     = if ($r.LastVisit -ne [datetime]::MinValue) { $r.LastVisit.ToLocalTime().ToString("yyyy-MM-dd") } else { "" }
    $safeUser    = [System.Web.HttpUtility]::HtmlEncode($r.UserName)
    $safeTitle   = [System.Web.HttpUtility]::HtmlEncode($r.Title)
    $safeUrl     = [System.Web.HttpUtility]::HtmlEncode($r.URL)
    $safeDomain  = [System.Web.HttpUtility]::HtmlEncode($r.Domain)
    $safeBrowser = [System.Web.HttpUtility]::HtmlEncode($r.Browser)
    $bClass      = ($r.Browser -replace '[^a-zA-Z]','').ToLower()
    $visitBarW   = [int]([Math]::Max(3, [Math]::Min(60, ($r.Visits / $maxVisits) * 60)))

    $rowsHtml += "<tr class=`"row`" data-user=`"$safeUser`" data-browser=`"$safeBrowser`" data-domain=`"$safeDomain`" data-url=`"$safeUrl`" data-title=`"$safeTitle`" data-date=`"$dateIso`">`n"
    $rowsHtml += "<td class=`"idx`">$idx</td>`n"
    $rowsHtml += "<td class=`"uc`"><span class=`"ub`">$safeUser</span></td>`n"
    $rowsHtml += "<td><span class=`"bb b-$bClass`">$safeBrowser</span></td>`n"
    $rowsHtml += "<td class=`"dc`" title=`"$safeDomain`">$safeDomain</td>`n"
    $rowsHtml += "<td class=`"tc`"><a href=`"$($r.URL)`" target=`"_blank`" rel=`"noopener`" title=`"$safeUrl`">$safeTitle</a></td>`n"
    $rowsHtml += "<td class=`"vc`"><div class=`"vw`"><div class=`"vb`" style=`"width:${visitBarW}px`"></div><span class=`"vn`">$($r.Visits)</span></div></td>`n"
    $rowsHtml += "<td class=`"dtc`">$dateStr</td>`n"
    $rowsHtml += "</tr>`n"
    $idx++
}

# Panel: records by user
$userStatsHtml = ""
$maxUserStat   = ($userStats.Values | Measure-Object -Maximum).Maximum
if (-not $maxUserStat -or $maxUserStat -eq 0) { $maxUserStat = 1 }
foreach ($kvp in ($userStats.GetEnumerator() | Sort-Object Value -Descending)) {
    $pct = [int](($kvp.Value / $maxUserStat) * 100)
    $sn  = [System.Web.HttpUtility]::HtmlEncode($kvp.Key)
    $userStatsHtml += "<div class=`"sr`"><span class=`"sn`">$sn</span><div class=`"sbw`"><div class=`"sb us`" style=`"width:$pct%`"></div></div><span class=`"sc`">$($kvp.Value)</span></div>`n"
}

# Panel: top 15 domains
$topDomainsHtml = ""
$topDomains = $allRecords | Where-Object { $_.Domain -and $_.Domain -ne "unknown" } |
    Group-Object Domain | Sort-Object Count -Descending | Select-Object -First 15
foreach ($dg in $topDomains) {
    $pct = [int](($dg.Count / [Math]::Max($allRecords.Count,1)) * 100)
    $dn  = [System.Web.HttpUtility]::HtmlEncode($dg.Name)
    $topDomainsHtml += "<div class=`"sr`"><span class=`"sn dn`">$dn</span><div class=`"sbw`"><div class=`"sb db`" style=`"width:$([Math]::Max($pct,2))%`"></div></div><span class=`"sc`">$($dg.Count)</span></div>`n"
}

$reportTime   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$hostName     = $env:COMPUTERNAME
$runAs        = $env:USERNAME
$totalRecords = $allRecords.Count
$userCount    = $userStats.Count
$vssStatus    = if ($vssOk) { "VSS Shadow Copy" } else { "Direct / Regex" }

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>BrowserHistory :: $hostName</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='6' fill='%2307090e'/><circle cx='16' cy='16' r='11' fill='none' stroke='%2300d4ff' stroke-width='2'/><line x1='16' y1='5' x2='16' y2='27' stroke='%2300d4ff' stroke-width='1.5'/><line x1='5' y1='16' x2='27' y2='16' stroke='%2300d4ff' stroke-width='1.5'/><ellipse cx='16' cy='16' rx='5.5' ry='11' fill='none' stroke='%230090ff' stroke-width='1.5'/></svg>"/>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;600;700&family=Rajdhani:wght@400;600;700&display=swap" rel="stylesheet"/>
<style>
:root{--bg:#07090e;--bg2:#0c1018;--bg3:#101520;--bg4:#141b24;--bd:#1a2a3a;--bd2:#203040;--ac:#00d4ff;--ac2:#0090ff;--ac3:#ff3060;--gr:#00ff88;--yw:#ffd700;--tx:#c0d0e0;--mt:#405060;--mt2:#253545;--pu:#9966ff}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'JetBrains Mono',monospace;font-size:13px;line-height:1.5;min-height:100vh}
body::before{content:'';position:fixed;inset:0;background:repeating-linear-gradient(0deg,transparent,transparent 3px,rgba(0,0,0,.08) 3px,rgba(0,0,0,.08) 4px);pointer-events:none;z-index:9998}
body::after{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(0,180,255,.02) 1px,transparent 1px),linear-gradient(90deg,rgba(0,180,255,.02) 1px,transparent 1px);background-size:36px 36px;pointer-events:none;z-index:0}
.wrap{position:relative;z-index:1;max-width:1600px;margin:0 auto;padding:0 24px 60px}
.hdr{padding:36px 0 28px;border-bottom:1px solid var(--bd);margin-bottom:32px;position:relative}
.hdr::after{content:'';position:absolute;bottom:-1px;left:0;width:180px;height:2px;background:linear-gradient(90deg,var(--ac),transparent)}
.hdr-row{display:flex;justify-content:space-between;flex-wrap:wrap;gap:16px;align-items:flex-start}
.logo-line{display:flex;align-items:center;gap:10px}
.lb{color:var(--ac);font-family:'Rajdhani',sans-serif;font-size:30px;font-weight:700}
.lt{font-family:'Rajdhani',sans-serif;font-size:28px;font-weight:700;letter-spacing:5px;text-transform:uppercase;background:linear-gradient(130deg,#fff 0%,var(--ac) 60%,var(--ac2) 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.ls{font-size:10px;color:var(--mt);letter-spacing:3px;text-transform:uppercase;margin-top:6px}
.hdr-meta{display:flex;flex-direction:column;gap:5px;align-items:flex-end}
.mi{font-size:11px;color:var(--mt);letter-spacing:1px}
.mi span{color:var(--ac)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:32px}
.card{background:var(--bg2);border:1px solid var(--bd);border-radius:4px;padding:18px 20px;position:relative;overflow:hidden;transition:border-color .2s,transform .2s}
.card:hover{border-color:var(--ac);transform:translateY(-2px)}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--ac2),var(--ac))}
.cl{font-size:9px;color:var(--mt);letter-spacing:2px;text-transform:uppercase;margin-bottom:8px}
.cv{font-family:'Rajdhani',sans-serif;font-size:36px;font-weight:700;color:#fff;line-height:1}
.cv.a{color:var(--ac)}.cv.g{color:var(--gr)}.cv.y{color:var(--yw)}.cv.p{color:var(--pu)}.cv.r{color:var(--ac3)}
.panels{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:32px}
@media(max-width:860px){.panels{grid-template-columns:1fr}}
.panel{background:var(--bg2);border:1px solid var(--bd);border-radius:4px;overflow:hidden}
.ph{padding:12px 18px;border-bottom:1px solid var(--bd);background:var(--bg3);display:flex;align-items:center;gap:8px}
.pt{font-family:'Rajdhani',sans-serif;font-size:12px;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:var(--ac)}
.pb{padding:18px}
.sr{display:flex;align-items:center;gap:10px;margin-bottom:9px}
.sn{width:140px;font-size:11px;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:0}
.sbw{flex:1;height:5px;background:var(--bg4);border-radius:3px;overflow:hidden}
.sb{height:100%;border-radius:3px;background:linear-gradient(90deg,var(--ac2),var(--ac))}
.us{background:linear-gradient(90deg,#5522bb,#9966ff)!important}
.sc{font-size:11px;color:var(--mt);width:46px;text-align:right;flex-shrink:0}
.dn{font-size:10px}.db{background:linear-gradient(90deg,var(--ac3),#ff7040)!important}
.b-chromium{background:linear-gradient(90deg,#1a52a8,#6ba3ff)!important}
.b-googlechrome{background:linear-gradient(90deg,#1050aa,#4285f4)!important}
.b-microsoftedge{background:linear-gradient(90deg,#004880,#0078d4)!important}
.b-brave,.b-bravebrowser{background:linear-gradient(90deg,#b03818,#fb542b)!important}
.b-yandexbrowser{background:linear-gradient(90deg,#a07800,#ffc800)!important}
.b-opera{background:linear-gradient(90deg,#900010,#ff1b2d)!important}
.b-operagx{background:linear-gradient(90deg,#900010,#ff3d5a)!important}
.b-vivaldi{background:linear-gradient(90deg,#901818,#ef3939)!important}
.b-torbrowser{background:linear-gradient(90deg,#3a1558,#7d4698)!important}
.b-waterfox{background:linear-gradient(90deg,#007090,#00c8ff)!important}
.b-librewolf{background:linear-gradient(90deg,#006070,#00acc1)!important}
.b-mozillafirefox{background:linear-gradient(90deg,#a03818,#ff7139)!important}
.b-thunderbird{background:linear-gradient(90deg,#003880,#0a84ff)!important}
.b-onedrivewebview,.b-winwebexperience{background:linear-gradient(90deg,#004060,#0078d4)!important}
.tb{background:var(--bg2);border:1px solid var(--bd);border-radius:4px;padding:12px 18px;margin-bottom:14px;display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.tbl{font-size:9px;color:var(--mt);letter-spacing:2px;text-transform:uppercase}
.sw{position:relative;flex:1;min-width:200px}
#si{width:100%;background:var(--bg3);border:1px solid var(--bd2);border-radius:3px;padding:7px 12px;color:var(--tx);font-family:'JetBrains Mono',monospace;font-size:12px;outline:none;transition:border-color .2s}
#si:focus{border-color:var(--ac)}
#si::placeholder{color:var(--mt)}
.fg{display:flex;gap:7px;flex-wrap:wrap}
.fb{background:var(--bg3);border:1px solid var(--bd2);border-radius:3px;padding:5px 11px;color:var(--mt);font-family:'JetBrains Mono',monospace;font-size:10px;cursor:pointer;transition:all .15s;white-space:nowrap}
.fb:hover,.fb.active{border-color:var(--ac);color:var(--ac);background:rgba(0,212,255,.05)}
.fu{background:var(--bg3);border:1px solid var(--bd2);border-radius:3px;padding:5px 11px;color:var(--mt);font-family:'JetBrains Mono',monospace;font-size:10px;cursor:pointer;transition:all .15s;white-space:nowrap}
.fu:hover,.fu.active{border-color:var(--pu);color:var(--pu);background:rgba(153,102,255,.05)}
.ctr{font-size:11px;color:var(--mt);margin-left:auto}.ctr span{color:var(--ac)}
.df{display:flex;align-items:center;gap:5px;flex-shrink:0}
.di{background:var(--bg3);border:1px solid var(--bd2);border-radius:3px;padding:5px 8px;color:var(--tx);font-family:'JetBrains Mono',monospace;font-size:10px;outline:none;transition:border-color .2s;width:120px}
.di:focus{border-color:var(--ac)}
.di::-webkit-calendar-picker-indicator{filter:invert(0.5) sepia(1) hue-rotate(180deg)}
.xbtn{background:none;border:1px solid var(--bd2);border-radius:3px;padding:4px 7px;color:var(--mt);font-size:11px;cursor:pointer;transition:all .15s;flex-shrink:0}
.xbtn:hover{border-color:var(--ac3);color:var(--ac3)}
.ebtn{background:var(--bg3);border:1px solid var(--bd2);border-radius:3px;padding:5px 12px;color:var(--gr);font-family:'JetBrains Mono',monospace;font-size:10px;cursor:pointer;transition:all .15s;white-space:nowrap;flex-shrink:0}
.ebtn:hover{border-color:var(--gr);background:rgba(0,255,136,.06)}
.vw{display:flex;align-items:center;gap:6px}
.vb{height:4px;border-radius:2px;background:linear-gradient(90deg,var(--ac2),var(--gr));min-width:3px;max-width:60px;flex-shrink:0}
.vn{color:var(--gr);font-size:12px;min-width:24px}
.tw{background:var(--bg2);border:1px solid var(--bd);border-radius:4px;overflow:hidden}
.ti{overflow-x:auto}
table{width:100%;border-collapse:collapse}
thead{background:var(--bg3);position:sticky;top:0;z-index:10}
th{padding:11px 13px;text-align:left;font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--ac);border-bottom:1px solid var(--bd);font-weight:600;cursor:pointer;user-select:none;white-space:nowrap}
th:hover{color:#fff}
th::after{content:' \21C5';opacity:.25;font-size:8px}
th.sa::after{content:' \2191';opacity:1}th.sd::after{content:' \2193';opacity:1}
tbody tr{border-bottom:1px solid var(--bd);transition:background .1s}
tbody tr:hover{background:rgba(0,212,255,.03)}
tbody tr.hidden{display:none}
td{padding:8px 13px;vertical-align:middle}
.idx{color:var(--mt);font-size:10px;width:46px}
.uc{max-width:120px}
.ub{display:inline-block;padding:2px 7px;border-radius:3px;font-size:9px;font-weight:600;letter-spacing:.5px;color:#fff;white-space:nowrap;background:linear-gradient(90deg,#4422aa,#9966ff)}
.bb{display:inline-block;padding:2px 7px;border-radius:3px;font-size:9px;font-weight:600;letter-spacing:.5px;color:#fff;white-space:nowrap;background:var(--mt2)}
.dc{color:var(--ac2);font-size:11px;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tc{max-width:340px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tc a{color:var(--tx);text-decoration:none;transition:color .15s}.tc a:hover{color:var(--ac)}
.vc{text-align:center;color:var(--gr);font-size:12px}
.dtc{color:var(--mt);font-size:11px;white-space:nowrap}
.ftr{margin-top:36px;padding-top:18px;border-top:1px solid var(--bd);display:flex;justify-content:space-between;flex-wrap:wrap;gap:10px}
.ft{font-size:10px;color:var(--mt);letter-spacing:1px}.fa{color:var(--ac)}
.fgl{color:var(--mt);text-decoration:none;transition:color .2s}.fgl:hover{color:var(--ac)}
.nr{text-align:center;padding:36px;color:var(--mt);font-size:12px;display:none}
.nr.v{display:block}
</style>
</head>
<body>
<div class="wrap">
<header class="hdr">
  <div class="hdr-row">
    <div>
      <div class="logo-line"><span class="lb">[</span><span class="lt">BrowserHistory</span><span class="lb">]</span></div>
      <div class="ls">Forensic Extraction Report &nbsp;//&nbsp; All Users &nbsp;//&nbsp; $vssStatus</div>
    </div>
    <div class="hdr-meta">
      <div class="mi">HOST <span>$hostName</span></div>
      <div class="mi">RUN AS <span>$runAs</span></div>
      <div class="mi">TIME <span>$reportTime</span></div>
    </div>
  </div>
</header>
<div class="cards">
  <div class="card"><div class="cl">Total Records</div><div class="cv a">$totalRecords</div></div>
  <div class="card"><div class="cl">Users Scanned</div><div class="cv p">$userCount</div></div>
  <div class="card"><div class="cl">Unique Domains</div><div class="cv y" id="dc">-</div></div>
  <div class="card"><div class="cl">Extraction Mode</div><div class="cv" style="font-size:14px;color:var(--mt);margin-top:4px">$vssStatus</div></div>
</div>
<div class="panels">
  <div class="panel"><div class="ph"><span class="pt">Records by User</span></div><div class="pb">$userStatsHtml</div></div>
  <div class="panel"><div class="ph"><span class="pt">Top Domains</span></div><div class="pb">$topDomainsHtml</div></div>
</div>
<div class="tb">
  <span class="tbl">User</span>
  <div class="fg" id="uf"><button class="fu active" data-u="all" onclick="suf('all',this)">All</button></div>
  <span class="tbl" style="margin-left:8px">Browser</span>
  <div class="fg" id="bf"><button class="fb active" data-b="all" onclick="sbf('all',this)">All</button></div>
  <div class="sw"><input type="text" id="si" placeholder="Search URL, title, domain..." oninput="ft()"/></div>
  <div class="df">
    <span class="tbl">From</span>
    <input type="date" id="df" class="di" onchange="ft()"/>
    <span class="tbl">To</span>
    <input type="date" id="dt" class="di" onchange="ft()"/>
    <button class="xbtn" onclick="clrDates()" title="Clear date filter">&#x2715;</button>
  </div>
  <button class="ebtn" onclick="exportCsv()" title="Export visible rows to CSV">&#x2913; CSV</button>
  <div class="ctr">Showing <span id="sc">$totalRecords</span> / $totalRecords</div>
</div>
<div class="tw"><div class="ti">
<table>
<thead><tr>
  <th onclick="st(0)">#</th>
  <th onclick="st(1)">User</th>
  <th onclick="st(2)">Browser</th>
  <th onclick="st(3)">Domain</th>
  <th onclick="st(4)">Title / URL</th>
  <th onclick="st(5)">Visits</th>
  <th onclick="st(6)">Last Visit</th>
</tr></thead>
<tbody id="tb">
$rowsHtml
</tbody>
</table>
</div>
<div class="nr" id="nr">No matching records.</div>
</div>
<div class="ftr">
  <div class="ft">Generated by <span class="fa">ZavetSec-BrowserHistory.ps1 v1.0</span> &nbsp;&#x2022;&nbsp; Extraction: <span class="fa">$vssStatus</span> &nbsp;&#x2022;&nbsp; <a href="https://github.com/zavetsec" target="_blank" rel="noopener" class="fgl">github.com/zavetsec</a></div>
  <div class="ft"><span class="fa">$hostName</span> &nbsp;&#x2022;&nbsp; $runAs &nbsp;&#x2022;&nbsp; $reportTime</div>
</div>
</div>
<script>
(function(){
  var r=document.querySelectorAll('#tb tr'),d=new Set();
  r.forEach(function(x){var v=x.dataset.domain;if(v&&v!=='unknown')d.add(v);});
  document.getElementById('dc').textContent=d.size;
})();
(function(){
  var r=document.querySelectorAll('#tb tr'),u=new Set(),b=new Set();
  r.forEach(function(x){u.add(x.dataset.user);b.add(x.dataset.browser);});
  var wu=document.getElementById('uf');
  u.forEach(function(v){
    if(!v)return;
    var btn=document.createElement('button');
    btn.className='fu';btn.textContent=v;btn.dataset.u=v;
    btn.onclick=function(){suf(v,btn);};
    wu.appendChild(btn);
  });
  var wb=document.getElementById('bf');
  b.forEach(function(v){
    if(!v)return;
    var btn=document.createElement('button');
    btn.className='fb';btn.textContent=v;btn.dataset.b=v;
    btn.onclick=function(){sbf(v,btn);};
    wb.appendChild(btn);
  });
})();
var au='all',ab='all',as_='';
function suf(n,el){au=n;document.querySelectorAll('.fu').forEach(function(b){b.classList.remove('active');});el.classList.add('active');ft();}
function sbf(n,el){ab=n;document.querySelectorAll('.fb').forEach(function(b){b.classList.remove('active');});el.classList.add('active');ft();}
function clrDates(){document.getElementById('df').value='';document.getElementById('dt').value='';ft();}
function ft(){
  as_=document.getElementById('si').value.toLowerCase();
  var dfv=document.getElementById('df').value;
  var dtv=document.getElementById('dt').value;
  var r=document.querySelectorAll('#tb tr'),s=0;
  r.forEach(function(x){
    var mu=au==='all'||x.dataset.user===au;
    var mb=ab==='all'||x.dataset.browser===ab;
    var h=(x.dataset.url+' '+x.dataset.title+' '+x.dataset.domain).toLowerCase();
    var ms=!as_||h.indexOf(as_)>=0;
    var md=true;
    if(dfv||dtv){
      var xd=x.dataset.date;
      if(xd){
        if(dfv&&xd<dfv)md=false;
        if(dtv&&xd>dtv)md=false;
      }
    }
    var v=mu&&mb&&ms&&md;x.classList.toggle('hidden',!v);if(v)s++;
  });
  document.getElementById('sc').textContent=s;
  document.getElementById('nr').classList.toggle('v',s===0);
}
var sc_=-1,sa=true;
function st(c){
  var tb=document.getElementById('tb'),r=Array.from(tb.querySelectorAll('tr')),hs=document.querySelectorAll('th');
  if(sc_===c){sa=!sa;}else{sc_=c;sa=true;}
  hs.forEach(function(h,i){h.classList.remove('sa','sd');if(i===c)h.classList.add(sa?'sa':'sd');});
  r.sort(function(a,b){
    var av=a.cells[c]?a.cells[c].textContent.trim():'',bv=b.cells[c]?b.cells[c].textContent.trim():'';
    var n=parseFloat(av)-parseFloat(bv);if(!isNaN(n))return sa?n:-n;
    return sa?av.localeCompare(bv):bv.localeCompare(av);
  });
  r.forEach(function(x){tb.appendChild(x);});
}
function exportCsv(){
  var rows=document.querySelectorAll('#tb tr:not(.hidden)');
  var lines=['#,User,Browser,Domain,Title,URL,Visits,LastVisit'];
  rows.forEach(function(x,i){
    function q(s){return '"'+s.replace(/"/g,'""')+'"';}
    var c=x.cells;
    if(!c||c.length<7)return;
    var a=x.querySelector('a');
    lines.push([
      i+1,
      q(x.dataset.user||''),
      q(x.dataset.browser||''),
      q(x.dataset.domain||''),
      q(x.dataset.title||''),
      q(a?a.href:x.dataset.url||''),
      c[5].textContent.trim().replace(/[^0-9]/g,'')||'1',
      q(c[6].textContent.trim())
    ].join(','));
  });
  var blob=new Blob(['\uFEFF'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});
  var a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='BrowserHistory_export.csv';
  a.click();
}
</script>
</body>
</html>
"@

$html | Set-Content -Path $OutputPath -Encoding UTF8

# Optional CSV export
if ($CsvExport -and $allRecords.Count -gt 0) {
    $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, ".csv")
    $allRecords | Select-Object UserName, Browser, Domain, Title, URL, Visits,
        @{ Name="LastVisit"; Expression={ if ($_.LastVisit -ne [datetime]::MinValue) { $_.LastVisit.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "" } } } |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
}

# ============================================================
#  ZIP WITH PASSWORD  (skipped when -NoArchive is set)
# ============================================================

$zipOk       = $false
$zipPath     = $null
$zipPassword = $null
$sevenZip    = $null

if ($NoArchive) {

    Write-Host "  [*] -NoArchive set - skipping ZIP, files saved as-is" -ForegroundColor Yellow

} else {

    # Generate a 16-character random password: mixed case + digits, no special chars
    # (special chars like # ^ % ! break 7z command line argument parsing)
    $pwChars  = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789'.ToCharArray()
    $rng      = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $pwBytes  = New-Object byte[] 16
    $rng.GetBytes($pwBytes)
    $zipPassword = -join ($pwBytes | ForEach-Object { $pwChars[$_ % $pwChars.Length] })
    $rng.Dispose()

    $zipPath = [System.IO.Path]::ChangeExtension($OutputPath, ".zip")

    # Collect files to pack
    $filesToPack = @($OutputPath)
    if ($CsvExport -and (Test-Path ([System.IO.Path]::ChangeExtension($OutputPath, ".csv")))) {
        $filesToPack += [System.IO.Path]::ChangeExtension($OutputPath, ".csv")
    }

    # Locate 7-Zip — always resolve to full absolute path to avoid PS execution warnings
    $sevenZipCandidates = @(
        (Join-Path $ScriptDir "7z.exe"),
        (Join-Path $ScriptDir "7-zip\7z.exe"),
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )
    # Check PATH entries manually without invoking Get-Command (avoids PS warning on local exe)
    $env:PATH -split ';' | Where-Object { $_ } | ForEach-Object {
        $sevenZipCandidates += (Join-Path $_ "7z.exe")
    }
    foreach ($c in $sevenZipCandidates) {
        if ($c -and (Test-Path $c)) { $sevenZip = (Resolve-Path $c).Path; break }
    }

    if ($sevenZip) {
        # AES-256 encrypted ZIP via direct ProcessStartInfo — bypasses all PS/cmd escaping
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName               = $sevenZip
            $psi.UseShellExecute        = $false
            $psi.CreateNoWindow         = $true
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true

            $argParts = @("a", "-tzip", "-mem=AES256", "-p$zipPassword", "`"$zipPath`"")
            foreach ($f in $filesToPack) { $argParts += "`"$f`"" }
            $psi.Arguments = $argParts -join ' '

            $proc = [System.Diagnostics.Process]::Start($psi)
            $stdout = $proc.StandardOutput.ReadToEnd()
            $stderr = $proc.StandardError.ReadToEnd()
            $proc.WaitForExit()

            if ($proc.ExitCode -eq 0 -and (Test-Path $zipPath)) {
                $zipOk = $true
                Write-Host "  [+] Encrypted ZIP : $zipPath" -ForegroundColor Green
                Write-Host "      Method         : AES-256 via 7-Zip" -ForegroundColor DarkGray
            } else {
                Write-Host "  [!] 7-Zip failed (exit $($proc.ExitCode))" -ForegroundColor Yellow
                if ($stderr) { Write-Host "      $($stderr.Trim())" -ForegroundColor DarkGray }
            }
        } catch {
            Write-Host "  [!] 7-Zip error: $_" -ForegroundColor Yellow
        }
    }

    if (-not $zipOk) {
        # Fallback: standard ZIP without encryption + warning
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
            $zip = [System.IO.Compression.ZipFile]::Open($zipPath, 'Create')
            foreach ($f in $filesToPack) {
                $entryName = Split-Path $f -Leaf
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f, $entryName, 'Optimal') | Out-Null
            }
            $zip.Dispose()
            $zipOk = $true
            Write-Host "  [+] ZIP saved      : $zipPath" -ForegroundColor Green
            Write-Host "  [!] WARNING        : 7-Zip failed - ZIP created WITHOUT encryption" -ForegroundColor Yellow
            Write-Host "      Required files   : 7z.exe + 7z.dll (both from 7-Zip installation)" -ForegroundColor DarkGray
            Write-Host "      Copy both to     : $ScriptDir" -ForegroundColor DarkGray
        } catch {
            Write-Host "  [!] ZIP creation failed: $_" -ForegroundColor Red
        }
    }

    # Remove unencrypted source files after successful archiving
    if ($zipOk) {
        foreach ($f in $filesToPack) {
            Remove-Item $f -Force -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================
#  FINAL SUMMARY
# ============================================================
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor DarkGray
Write-Host "   REPORT READY" -ForegroundColor Cyan
Write-Host "  ================================================================" -ForegroundColor DarkGray
Write-Host ""

if ($NoArchive) {
    Write-Host "   Report   : $OutputPath" -ForegroundColor White
    if ($CsvExport) {
        $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, ".csv")
        if (Test-Path $csvPath) {
            Write-Host "   CSV      : $csvPath" -ForegroundColor White
        }
    }
    Write-Host ""
    Write-Host "   Mode     : plain files (no archive, no password)" -ForegroundColor Yellow
} elseif ($zipOk) {
    Write-Host "   Archive  : $zipPath" -ForegroundColor White
    Write-Host ""
    if ($sevenZip) {
        Write-Host "   Password : " -ForegroundColor DarkGray -NoNewline
        Write-Host $zipPassword -ForegroundColor Yellow
        Write-Host ""
        Write-Host "   Save this password - it will not be shown again." -ForegroundColor DarkGray
    } else {
        Write-Host "   Password : [not set - 7-Zip unavailable]" -ForegroundColor Yellow
        Write-Host "   Protect the ZIP manually if transferring over untrusted channels." -ForegroundColor DarkGray
    }
} else {
    Write-Host "   Report   : $OutputPath" -ForegroundColor White
    Write-Host ""
    Write-Host "   [!] ZIP archiving failed - report left as plain file." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  ================================================================" -ForegroundColor DarkGray
Write-Host ""

if ($OpenReport) {
    if ($NoArchive -or (-not $zipOk)) {
        Start-Process $OutputPath
    } elseif ($zipOk -and $sevenZip) {
        $tmpDir  = Join-Path $env:TEMP ("BHE_View_" + [System.IO.Path]::GetRandomFileName())
        New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
        $htmlLeaf = Split-Path $OutputPath -Leaf
        try {
            $psi2 = New-Object System.Diagnostics.ProcessStartInfo
            $psi2.FileName               = $sevenZip
            $psi2.UseShellExecute        = $false
            $psi2.CreateNoWindow         = $true
            $psi2.RedirectStandardOutput = $true
            $psi2.RedirectStandardError  = $true
            $psi2.Arguments              = "e `"$zipPath`" `"-o$tmpDir`" -p$zipPassword -y"
            $proc2 = [System.Diagnostics.Process]::Start($psi2)
            $proc2.WaitForExit()
        } catch {}
        $tmpHtml = Join-Path $tmpDir $htmlLeaf
        if (Test-Path $tmpHtml) { Start-Process $tmpHtml }
    }
}
