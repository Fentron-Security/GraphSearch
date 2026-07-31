# Copyright 2026 Fentron Security Solutions
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

<#
.SYNOPSIS
    Stage 2 of GraphSearch. Takes a Stage 1 keyword report, downloads each hit
    file, extracts its text, and runs format+checksum validators to confirm
    whether a real identifier (not just the keyword) is present - with context.

.DESCRIPTION
    The Microsoft Search index answers "does this word appear in the file."
    It cannot answer "is this a real Sozialversicherungsnummer, and in what
    context." This script closes that gap: it reads the actual bytes of each
    hit file locally, so full regex + checksum validation is possible.

    Two-stage flow:
      1. Invoke-GraphKeywordDiscovery.ps1  -> report CSV (locations)
      2. Invoke-ContextValidation.ps1 -ReportCsv <that csv>  -> validated CSV

    Text extraction is native for: txt, csv, tsv, log, xml, html, json, md,
    docx, xlsx, pptx, msg. PDF needs pdftotext (poppler). Legacy binary
    doc/xls/ppt and zip archives are handled if LibreOffice (soffice) is
    available; zips are recursed into. Missing helpers are reported once at
    startup, not per file. Use -DryRun to plan a run (types, sizes, coverage)
    without authenticating or downloading anything.

    Output is two reports: a full occurrence-level CSV, and a DISTINCT report
    that collapses repeated numbers to one row per unique identifier with file
    and occurrence counts - the client-ready view.

.EXAMPLE
    .\Invoke-ContextValidation.ps1 `
        -TenantId $env:GRAPH_TENANT_ID -ClientId $env:GRAPH_CLIENT_ID `
        -ClientSecret $env:GRAPH_CLIENT_SECRET `
        -ReportCsv .\report\keyword-discovery-tenant-20260730-120000.csv `
        -Validators DE-RVNR,CH-AHV

.NOTES
    Requires Files.Read.All (to download file content). Same app registration
    as Stage 1. Downloaded files are written to a temp dir and deleted after
    extraction unless -KeepFiles is set.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$ClientSecret,

    # Stage 1 output CSV. Must contain DriveId and ItemId columns.
    [Parameter(Mandatory)][string]$ReportCsv,

    # Select validators to run by any combination of the three. With none set,
    # ALL validators run (noisy for broad no-checksum patterns - prefer scoping
    # by -Country or -Category for a real engagement).
    #   -Validators DE-RVNR,CH-AHV        by name
    #   -Country    DE,CH,US              by jurisdiction
    #   -Category   payment,bank,national-id,employee-id
    [string[]]$Validators,
    [string[]]$Country,
    [string[]]$Category,

    # Characters of surrounding text to capture per match.
    [int]$Context = 60,

    # Cap on file size to download (MB). Larger files are skipped. Raise for
    # tenants with big HR exports, e.g. -MaxFileMB 100.
    [int]$MaxFileMB = 25,

    # Max entries to extract from a single zip archive (guards against zip bombs).
    [int]$MaxZipEntries = 200,

    # Skip the distinct/summary report and emit only the full occurrence CSV.
    [switch]$NoDistinctReport,

    # Plan only: classify what a run WOULD download (types, sizes, coverage) and
    # write a manifest, without authenticating or pulling any file content.
    [switch]$DryRun,

    [string]$OutDir = ".\report",
    [switch]$KeepFiles
)

$ErrorActionPreference = "Stop"

Import-Module (Join-Path $PSScriptRoot 'GraphSearch.Validators.psm1') -Force

# ---------- Preflight: external helpers ----------
# Detect optional converters once, up front, so the operator sees what's
# missing (and how to install it) before a long run instead of mid-stream.
$script:HasPdftotext = [bool](Get-Command pdftotext -ErrorAction SilentlyContinue)
$script:HasSoffice   = [bool](Get-Command soffice   -ErrorAction SilentlyContinue)
if (-not $script:HasSoffice) {
    # LibreOffice often isn't on PATH even when installed; probe common locations.
    $sofficePaths = @(
        "$env:ProgramFiles\LibreOffice\program\soffice.exe",
        "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    )
    $found = $sofficePaths | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if ($found) { $script:SofficePath = $found; $script:HasSoffice = $true }
}

Write-Host "Helper availability:" -ForegroundColor Yellow
Write-Host ("  pdftotext (PDF)          : {0}" -f ($(if($script:HasPdftotext){'yes'}else{'NO'}))) -ForegroundColor $(if($script:HasPdftotext){'Green'}else{'DarkYellow'})
Write-Host ("  soffice (doc/xls/ppt,zip): {0}" -f ($(if($script:HasSoffice){'yes'}else{'NO'}))) -ForegroundColor $(if($script:HasSoffice){'Green'}else{'DarkYellow'})
if (-not $script:HasPdftotext) { Write-Host "    -> PDFs will be skipped. Install: winget install -e --id oschwartz10612.Poppler" -ForegroundColor DarkGray }
if (-not $script:HasSoffice)   { Write-Host "    -> Legacy Office (.doc/.xls/.ppt) and zip contents will be skipped. Install: winget install -e --id TheDocumentFoundation.LibreOffice" -ForegroundColor DarkGray }
Write-Host ""

if (-not (Test-Path $ReportCsv)) { throw "Report CSV not found: $ReportCsv" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("gs-ctx-" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

# ---------- Auth ----------
function Get-GraphToken {
    param($TenantId,$ClientId,$ClientSecret)
    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }
    (Invoke-RestMethod -Method POST `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -ContentType "application/x-www-form-urlencoded" -Body $body).access_token
}
if (-not $DryRun) {
    $script:Token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
}

function Invoke-GraphDownload {
    param([string]$DriveId,[string]$ItemId,[string]$OutFile)
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/content"
    for ($i=0; $i -lt 5; $i++) {
        try {
            Invoke-WebRequest -Uri $uri -Headers @{ Authorization = "Bearer $script:Token" } `
                -OutFile $OutFile -ErrorAction Stop
            return $true
        } catch {
            $code = $_.Exception.Response.StatusCode.value__
            if ($code -eq 401) {
                $script:Token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
                continue
            }
            if ($code -eq 429 -or $code -ge 500) {
                $wait = [int]($_.Exception.Response.Headers["Retry-After"]); if (-not $wait) { $wait = [math]::Pow(2,$i) }
                Start-Sleep -Seconds $wait; continue
            }
            Write-Warning "download failed ($code): $ItemId"
            return $false
        }
    }
    return $false
}

# ---------- Text extraction ----------
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ZipEntryText {
    param([string]$Zip,[string]$EntryPattern)
    $sb = New-Object System.Text.StringBuilder
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Zip)
    try {
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -match $EntryPattern) {
                $reader = New-Object System.IO.StreamReader($entry.Open())
                try { [void]$sb.AppendLine($reader.ReadToEnd()) } finally { $reader.Dispose() }
            }
        }
    } finally { $archive.Dispose() }
    return $sb.ToString()
}

function Convert-XmlToText { param([string]$Xml)
    # strip tags, decode the handful of XML entities that matter for matching
    $t = $Xml -replace '<[^>]+>',' '
    $t = $t -replace '&amp;','&' -replace '&lt;','<' -replace '&gt;','>' -replace '&quot;','"' -replace '&#39;',"'"
    return $t
}

function Invoke-Soffice {
    param([string]$Source,[string]$ConvertTo,[string]$OutDir)
    $exe = if ($script:SofficePath) { $script:SofficePath } else { 'soffice' }
    # headless convert; soffice writes <basename>.<ext> into OutDir
    & $exe --headless --norestore --convert-to $ConvertTo --outdir $OutDir $Source *> $null
    $target = Join-Path $OutDir ([System.IO.Path]::GetFileNameWithoutExtension($Source) + '.' + ($ConvertTo -split ':')[0])
    if (Test-Path $target) { return $target }
    return $null
}

function Get-MsgText {
    # Outlook .msg is an OLE2 compound file; the body/subject/headers live in
    # UTF-16 (and sometimes ASCII) property streams. Rather than pull in an
    # OLE2 parser, extract printable UTF-16LE and ASCII runs directly - crude,
    # but it reliably surfaces the identifiers and surrounding words we need.
    # Note: binary attachments inside the .msg (e.g. an embedded PDF) are not
    # decoded here; they'd need separate extraction.
    param([string]$Path)
    try { $bytes = [System.IO.File]::ReadAllBytes($Path) } catch { return $null }
    $sb = New-Object System.Text.StringBuilder
    # UTF-16LE at both byte alignments (stream may not be even-aligned once
    # we're scanning the whole container rather than a single sector).
    foreach ($off in 0,1) {
        if ($off -ge $bytes.Length) { continue }
        $u = [System.Text.Encoding]::Unicode.GetString($bytes, $off, $bytes.Length - $off)
        foreach ($m in [regex]::Matches($u, '[\x20-\x7E\u00A0-\u024F]{4,}')) { [void]$sb.AppendLine($m.Value) }
    }
    $a = [System.Text.Encoding]::ASCII.GetString($bytes)
    foreach ($m in [regex]::Matches($a, '[\x20-\x7E]{4,}')) { [void]$sb.AppendLine($m.Value) }
    return $sb.ToString()
}

function Convert-LegacyOffice {
    # Convert .doc/.xls/.ppt to their modern equivalents via LibreOffice, then
    # reuse the native extractor. Returns extracted text or $null.
    param([string]$Path,[string]$Extension,[string]$WorkDir)
    if (-not $script:HasSoffice) {
        Write-Warning "  soffice not available; skipping legacy $Extension $([IO.Path]::GetFileName($Path))"
        return $null
    }
    $map = @{ '.doc'='docx'; '.xls'='xlsx'; '.ppt'='pptx' }
    $to  = $map[$Extension.ToLower()]
    if (-not $to) { return $null }
    $converted = Invoke-Soffice -Source $Path -ConvertTo $to -OutDir $WorkDir
    if (-not $converted) { Write-Warning "  conversion failed: $([IO.Path]::GetFileName($Path))"; return $null }
    $text = Get-FileText -Path $converted -Extension ".$to" -Depth 1
    Remove-Item $converted -Force -ErrorAction SilentlyContinue
    return $text
}

function Get-ArchiveText {
    # Recurse one level into a zip: extract supported entries and concatenate
    # their text. Depth-guarded so nested zips don't recurse without bound.
    param([string]$Path,[int]$Depth)
    if ($Depth -ge 2) { return $null }   # don't descend into zips-in-zips-in-zips
    $sb = New-Object System.Text.StringBuilder
    $ex = Join-Path ([System.IO.Path]::GetDirectoryName($Path)) ([guid]::NewGuid().ToString('N').Substring(0,8))
    New-Item -ItemType Directory -Force -Path $ex | Out-Null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $count = 0
        try {
            foreach ($entry in $archive.Entries) {
                if ([string]::IsNullOrEmpty($entry.Name)) { continue }   # directory
                if ($count -ge $MaxZipEntries) { Write-Warning "  zip entry cap ($MaxZipEntries) reached"; break }
                $count++
                $dest = Join-Path $ex ("{0:D4}{1}" -f $count, [System.IO.Path]::GetExtension($entry.Name))
                try {
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
                    $t = Get-FileText -Path $dest -Extension ([System.IO.Path]::GetExtension($entry.Name)) -Depth ($Depth + 1)
                    if ($t) { [void]$sb.AppendLine($t) }
                } catch { } finally { Remove-Item $dest -Force -ErrorAction SilentlyContinue }
            }
        } finally { $archive.Dispose() }
    } finally { Remove-Item $ex -Recurse -Force -ErrorAction SilentlyContinue }
    return $sb.ToString()
}

function Get-FileText {
    param([string]$Path,[string]$Extension,[int]$Depth = 0)
    switch -Regex ($Extension.ToLower()) {
        '^\.(txt|csv|tsv|log|md|json)$' {
            return Get-Content -Path $Path -Raw -Encoding UTF8
        }
        '^\.(xml|html?|xhtml)$' {
            return Convert-XmlToText (Get-Content -Path $Path -Raw -Encoding UTF8)
        }
        '^\.docx$' {
            return Convert-XmlToText (Get-ZipEntryText -Zip $Path -EntryPattern '^word/document\.xml$')
        }
        '^\.xlsx$' {
            $strings = Get-ZipEntryText -Zip $Path -EntryPattern '^xl/sharedStrings\.xml$'
            $sheets  = Get-ZipEntryText -Zip $Path -EntryPattern '^xl/worksheets/sheet\d+\.xml$'
            return (Convert-XmlToText $strings) + ' ' + (Convert-XmlToText $sheets)
        }
        '^\.pptx$' {
            return Convert-XmlToText (Get-ZipEntryText -Zip $Path -EntryPattern '^ppt/slides/slide\d+\.xml$')
        }
        '^\.(doc|xls|ppt)$' {
            return Convert-LegacyOffice -Path $Path -Extension $Extension -WorkDir ([System.IO.Path]::GetDirectoryName($Path))
        }
        '^\.zip$' {
            return Get-ArchiveText -Path $Path -Depth $Depth
        }
        '^\.msg$' {
            return Get-MsgText -Path $Path
        }
        '^\.pdf$' {
            if (-not $script:HasPdftotext) { return $null }   # already warned at startup
            $txt = "$Path.txt"
            & pdftotext -q -enc UTF-8 $Path $txt 2>$null
            if (Test-Path $txt) { $c = Get-Content $txt -Raw -Encoding UTF8; Remove-Item $txt -Force; return $c }
            return $null
        }
        default {
            Write-Warning "unsupported type $Extension; skipping $([IO.Path]::GetFileName($Path))"
            return $null
        }
    }
}

# ---------- Main ----------
$rows = Import-Csv -Path $ReportCsv
if (-not ($rows | Get-Member -Name DriveId) -or -not ($rows | Get-Member -Name ItemId)) {
    throw "Report CSV must contain DriveId and ItemId columns (produced by Stage 1)."
}

# de-dupe files: one download per unique DriveId+ItemId even if multiple keyword hits
$files = $rows | Sort-Object DriveId,ItemId -Unique

# Classify an extension by how Stage 2 will handle it.
function Get-Disposition {
    param([string]$Ext,[double]$SizeKB)
    if ($SizeKB -and ($SizeKB / 1024) -gt $MaxFileMB) { return 'oversized' }
    switch -Regex ($Ext.ToLower()) {
        '^\.(txt|csv|tsv|log|md|json|xml|html?|xhtml|docx|xlsx|pptx)$' { 'native' }
        '^\.pdf$'            { if ($script:HasPdftotext) { 'pdf' } else { 'pdf (helper missing)' } }
        '^\.(doc|xls|ppt)$'  { if ($script:HasSoffice)   { 'legacy-office' } else { 'legacy-office (helper missing)' } }
        '^\.zip$'            { 'zip' }
        '^\.msg$'            { 'msg' }
        default              { 'unsupported' }
    }
}

# ---------- Dry run: plan only, no auth, no download ----------
if ($DryRun) {
    $plan = foreach ($f in $files) {
        $ext = [System.IO.Path]::GetExtension($f.FileName)
        [pscustomobject]@{
            FileName    = $f.FileName
            Extension   = $ext.ToLower()
            SizeKB      = $f.SizeKB
            Disposition = Get-Disposition -Ext $ext -SizeKB ([double]($f.SizeKB))
            Owner       = $f.Owner
            WebUrl      = $f.WebUrl
        }
    }
    $stamp = Get-Date -Format yyyyMMdd-HHmmss
    $manifest = Join-Path $OutDir "context-validation-$stamp-dryrun.csv"
    $plan | Export-Csv -Path $manifest -NoTypeInformation -Encoding UTF8

    $totalMB = [math]::Round((($files | Measure-Object -Property SizeKB -Sum).Sum / 1024), 1)
    Write-Host "DRY RUN - no content downloaded." -ForegroundColor Yellow
    Write-Host "$($files.Count) unique file(s), ~${totalMB}MB to download for a real run." -ForegroundColor Yellow
    Write-Host "Manifest -> $manifest`n" -ForegroundColor Green
    Write-Host "By planned disposition:" -ForegroundColor Yellow
    $plan | Group-Object Disposition | Sort-Object Count -Descending |
        Select-Object @{n='Disposition';e={$_.Name}}, Count | Format-Table -AutoSize
    Write-Host "By extension:" -ForegroundColor Yellow
    $plan | Group-Object Extension | Sort-Object Count -Descending |
        Select-Object @{n='Extension';e={$_.Name}}, Count | Format-Table -AutoSize
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    return
}

# Resolve selection once for logging.
$selected = Get-ValidatorDefinitions -Name $Validators -Country $Country -Category $Category
if (-not $selected) { throw "No validators matched the selection (Validators/Country/Category)." }
Write-Host "Validating $($files.Count) unique file(s) with: $((@($selected.Name) -join ', '))" -ForegroundColor Yellow

$findings = New-Object System.Collections.Generic.List[object]
$n = 0
foreach ($f in $files) {
    $n++
    $name = $f.FileName
    $ext  = [System.IO.Path]::GetExtension($name)
    Write-Host "[$n/$($files.Count)] $name" -ForegroundColor Cyan

    if ($f.SizeKB -and ([double]$f.SizeKB / 1024) -gt $MaxFileMB) {
        Write-Warning "  over ${MaxFileMB}MB; skipped"; continue
    }

    $local = Join-Path $tmp ("{0:D5}{1}" -f $n, $ext)
    if (-not (Invoke-GraphDownload -DriveId $f.DriveId -ItemId $f.ItemId -OutFile $local)) { continue }

    try   { $text = Get-FileText -Path $local -Extension $ext }
    catch { Write-Warning "  extraction failed: $($_.Exception.Message)"; $text = $null }

    if ($text) {
        foreach ($hit in (Invoke-Validators -Text $text -Name $Validators -Country $Country -Category $Category -Context $Context)) {
            $findings.Add([pscustomobject]@{
                FileName      = $name
                WebUrl        = $f.WebUrl
                Owner         = $f.Owner
                Type          = $hit.Type
                IdCountry     = $hit.Country
                Category      = $hit.Category
                Match         = $hit.Match
                Normalized    = ($hit.Match -replace '[\s.\-]','').ToUpper()
                ChecksumValid = $hit.ChecksumValid
                Offset        = $hit.Offset
                Context       = $hit.Context
                DriveId       = $f.DriveId
                ItemId        = $f.ItemId
            })
        }
    }

    if (-not $KeepFiles) { Remove-Item $local -Force -ErrorAction SilentlyContinue }
}

if (-not $KeepFiles) { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }
else { Write-Host "Downloaded files kept in $tmp" -ForegroundColor DarkGray }

# ---------- Report: full occurrence-level ----------
$stamp = Get-Date -Format yyyyMMdd-HHmmss
$csv = Join-Path $OutDir "context-validation-$stamp.csv"
$findings | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-Host "`n$($findings.Count) occurrence(s) -> $csv" -ForegroundColor Green

$hasExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)
if ($hasExcel) {
    $xlsx = $csv -replace '\.csv$','.xlsx'
    $findings | Export-Excel -Path $xlsx -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName "Occurrences"
    Write-Host "XLSX -> $xlsx" -ForegroundColor Green
}

# ---------- Report: distinct (client-ready) ----------
# Collapse repeated numbers: one row per unique identifier, with how many files
# and occurrences it appeared in. This is the count that goes to a client -
# "N distinct AHV numbers", not "half a million cell hits".
if (-not $NoDistinctReport) {
    $distinct = $findings | Group-Object Type,Normalized | ForEach-Object {
        $g = $_.Group
        $first = $g[0]
        [pscustomobject]@{
            Type          = $first.Type
            IdCountry     = $first.IdCountry
            Category      = $first.Category
            Match         = $first.Match
            ChecksumValid = $first.ChecksumValid
            FileCount     = (@($g | Select-Object -ExpandProperty WebUrl -Unique)).Count
            Occurrences   = $g.Count
            SampleFile    = $first.FileName
            SampleUrl     = $first.WebUrl
            SampleContext = $first.Context
        }
    } | Sort-Object Type, @{Expression='FileCount';Descending=$true}

    $dcsv = Join-Path $OutDir "context-validation-$stamp-distinct.csv"
    $distinct | Export-Csv -Path $dcsv -NoTypeInformation -Encoding UTF8
    Write-Host "$($distinct.Count) distinct identifier(s) -> $dcsv" -ForegroundColor Green
    if ($hasExcel) {
        $dxlsx = $dcsv -replace '\.csv$','.xlsx'
        $distinct | Export-Excel -Path $dxlsx -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName "Distinct"
        Write-Host "XLSX -> $dxlsx" -ForegroundColor Green
    }
}

# ---------- Summary ----------
$confirmed = @($findings | Where-Object { $_.ChecksumValid -eq $true }).Count
$invalid   = @($findings | Where-Object { $_.ChecksumValid -eq $false }).Count
$unknown   = @($findings | Where-Object { $null -eq $_.ChecksumValid }).Count
Write-Host ""
Write-Host "Occurrences" -ForegroundColor Yellow
Write-Host "  checksum-valid                    : $confirmed" -ForegroundColor Green
Write-Host "  format-match, checksum-invalid    : $invalid" -ForegroundColor DarkYellow
Write-Host "  no checksum (e.g. Personalnummer) : $unknown" -ForegroundColor DarkGray

if (-not $NoDistinctReport) {
    Write-Host "`nDistinct identifiers by type (checksum-valid only)" -ForegroundColor Yellow
    $findings |
        Where-Object { $_.ChecksumValid -eq $true } |
        Group-Object Type |
        ForEach-Object {
            [pscustomobject]@{
                Type     = $_.Name
                Distinct = (@($_.Group | Select-Object -ExpandProperty Normalized -Unique)).Count
                Files    = (@($_.Group | Select-Object -ExpandProperty WebUrl -Unique)).Count
            }
        } | Sort-Object Type | Format-Table -AutoSize
}
