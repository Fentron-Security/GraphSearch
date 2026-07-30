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
    docx, xlsx, pptx. PDF requires pdftotext (poppler) on PATH; if absent, PDFs
    are skipped with a warning. Legacy binary doc/xls/ppt are not supported.

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

    # Cap on file size to download (MB). Larger files are skipped.
    [int]$MaxFileMB = 25,

    [string]$OutDir = ".\report",
    [switch]$KeepFiles
)

$ErrorActionPreference = "Stop"

Import-Module (Join-Path $PSScriptRoot 'GraphSearch.Validators.psm1') -Force

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
$script:Token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

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

function Get-FileText {
    param([string]$Path,[string]$Extension)
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
        '^\.pdf$' {
            $pdftotext = Get-Command pdftotext -ErrorAction SilentlyContinue
            if (-not $pdftotext) { Write-Warning "pdftotext not found; skipping PDF $([IO.Path]::GetFileName($Path))"; return $null }
            $txt = "$Path.txt"
            & $pdftotext.Source -q -enc UTF-8 $Path $txt 2>$null
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

# ---------- Report ----------
$stamp = Get-Date -Format yyyyMMdd-HHmmss
$csv = Join-Path $OutDir "context-validation-$stamp.csv"
$findings | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-Host "`n$($findings.Count) finding(s) -> $csv" -ForegroundColor Green

if (Get-Module -ListAvailable -Name ImportExcel) {
    $xlsx = $csv -replace '\.csv$','.xlsx'
    $findings | Export-Excel -Path $xlsx -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName "Validated"
    Write-Host "XLSX -> $xlsx" -ForegroundColor Green
}

# summary: confirmed vs keyword-only
$confirmed = @($findings | Where-Object { $_.ChecksumValid -eq $true }).Count
$invalid   = @($findings | Where-Object { $_.ChecksumValid -eq $false }).Count
$unknown   = @($findings | Where-Object { $null -eq $_.ChecksumValid }).Count
Write-Host ""
Write-Host "  checksum-valid : $confirmed" -ForegroundColor Green
Write-Host "  format-match, checksum-invalid : $invalid" -ForegroundColor DarkYellow
Write-Host "  no checksum (e.g. Personalnummer) : $unknown" -ForegroundColor DarkGray
$findings | Group-Object Type | Select-Object Name,Count | Format-Table -AutoSize
