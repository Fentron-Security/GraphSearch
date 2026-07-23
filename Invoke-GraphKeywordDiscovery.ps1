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
    Enterprise-wide keyword discovery across OneDrive + SharePoint via Microsoft Graph.
    App-only (client credentials). Outputs a location report (CSV + XLSX).

.NOTES
    Entra app registration requires APPLICATION permissions (admin consent):
        Files.Read.All, Sites.Read.All
    For the /search/query app-only path you MUST pass a Region in the request body.

.EXAMPLE
    .\Invoke-GraphKeywordDiscovery.ps1 -TenantId <guid> -ClientId <guid> `
        -ClientSecret <secret> -Keywords "confidential","SSN","passport" -Region NAM
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$ClientSecret,
    [Parameter(Mandatory)][string[]]$Keywords,

    # Graph app-only search region: NAM, EUR, APC, etc. Required for app-only /search/query.
    [string]$Region = "NAM",

    # search = index-based (content + metadata, fast, enterprise-wide).
    # enumerate = walk every drive and per-drive search (slower, exhaustive, good fallback).
    [ValidateSet("search","enumerate")][string]$Mode = "search",

    [string]$OutDir = ".\report",
    [int]$PageSize = 200
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# ---------- Auth (client credentials) ----------
function Get-GraphToken {
    param($TenantId,$ClientId,$ClientSecret)
    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }
    $resp = Invoke-RestMethod -Method POST `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -ContentType "application/x-www-form-urlencoded" -Body $body
    return $resp.access_token
}

$script:Token   = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
$script:Headers = @{ Authorization = "Bearer $script:Token"; "Content-Type" = "application/json" }

function Invoke-Graph {
    param([string]$Uri,[string]$Method="GET",$Body)
    for ($i=0; $i -lt 6; $i++) {
        try {
            if ($Body) {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $script:Headers `
                    -Body ($Body | ConvertTo-Json -Depth 12)
            } else {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $script:Headers
            }
        } catch {
            $code = $_.Exception.Response.StatusCode.value__
            if ($code -eq 401) {
                # token refresh
                $script:Token   = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
                $script:Headers = @{ Authorization = "Bearer $script:Token"; "Content-Type" = "application/json" }
                continue
            }
            if ($code -eq 429 -or $code -ge 500) {
                $wait = [int]($_.Exception.Response.Headers["Retry-After"]); if (-not $wait) { $wait = [math]::Pow(2,$i) }
                Write-Warning "Throttled/err $code — retry in $wait s"; Start-Sleep -Seconds $wait; continue
            }
            throw
        }
    }
    throw "Graph call failed after retries: $Uri"
}

$results = New-Object System.Collections.Generic.List[object]

function Add-Hit {
    param($item,$keyword)
    $pr = $item.parentReference
    $results.Add([pscustomobject]@{
        Keyword      = $keyword
        FileName     = $item.name
        WebUrl       = $item.webUrl
        SiteId       = $pr.siteId
        DriveId      = $pr.driveId
        Path         = ($pr.path -replace "^/drives/[^/]+/root:","")
        SizeKB       = if ($item.size) { [math]::Round($item.size/1KB,1) } else { $null }
        LastModified = $item.lastModifiedDateTime
        ModifiedBy   = $item.lastModifiedBy.user.displayName
        CreatedBy    = $item.createdBy.user.displayName
        ItemId       = $item.id
    })
}

# ---------- Mode 1: index-based /search/query (enterprise-wide) ----------
function Search-IndexWide {
    param([string[]]$Keywords,[string]$Region,[int]$PageSize)
    foreach ($kw in $Keywords) {
        Write-Host "[search] '$kw'" -ForegroundColor Cyan
        $from = 0
        do {
            $body = @{
                requests = @(@{
                    entityTypes = @("driveItem")
                    query       = @{ queryString = $kw }
                    from        = $from
                    size        = $PageSize
                    region      = $Region      # REQUIRED for app-only
                    fields      = @("name","webUrl","parentReference","lastModifiedDateTime",
                                    "lastModifiedBy","createdBy","size","id")
                })
            }
            $resp = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/search/query" -Method POST -Body $body
            $container = $resp.value[0].hitsContainers[0]
            foreach ($hit in $container.hits) { Add-Hit -item $hit.resource -keyword $kw }
            $more = $container.moreResultsAvailable
            $from += $PageSize
        } while ($more)
    }
}

# ---------- Mode 2: enumerate every drive, per-drive search (exhaustive) ----------
function Get-AllDrives {
    $drives = New-Object System.Collections.Generic.List[object]

    # SharePoint sites
    $uri = "https://graph.microsoft.com/v1.0/sites?search=*&`$top=100"
    do {
        $r = Invoke-Graph -Uri $uri
        foreach ($s in $r.value) {
            try {
                $d = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/sites/$($s.id)/drives?`$top=100"
                foreach ($drv in $d.value) { $drives.Add([pscustomobject]@{ Id=$drv.id; Owner=$s.webUrl }) }
            } catch { Write-Warning "site $($s.webUrl): $($_.Exception.Message)" }
        }
        $uri = $r.'@odata.nextLink'
    } while ($uri)

    # OneDrive per user
    $uri = "https://graph.microsoft.com/v1.0/users?`$top=100&`$select=id,userPrincipalName"
    do {
        $r = Invoke-Graph -Uri $uri
        foreach ($u in $r.value) {
            try {
                $d = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$($u.id)/drive"
                $drives.Add([pscustomobject]@{ Id=$d.id; Owner=$u.userPrincipalName })
            } catch { } # user may have no OneDrive provisioned
        }
        $uri = $r.'@odata.nextLink'
    } while ($uri)

    return $drives
}

function Search-Enumerate {
    param([string[]]$Keywords)
    $drives = Get-AllDrives
    Write-Host "[enumerate] $($drives.Count) drives" -ForegroundColor Cyan
    foreach ($drv in $drives) {
        foreach ($kw in $Keywords) {
            $uri = "https://graph.microsoft.com/v1.0/drives/$($drv.Id)/root/search(q='$kw')?`$top=200"
            do {
                try { $r = Invoke-Graph -Uri $uri } catch { break }
                foreach ($item in $r.value) { Add-Hit -item $item -keyword $kw }
                $uri = $r.'@odata.nextLink'
            } while ($uri)
        }
    }
}

# ---------- Run ----------
if ($Mode -eq "search") { Search-IndexWide -Keywords $Keywords -Region $Region -PageSize $PageSize }
else                    { Search-Enumerate -Keywords $Keywords }

# de-dupe (same file can hit on multiple keywords — keep all, but drop exact dupes per keyword)
$results = $results | Sort-Object Keyword,WebUrl -Unique

$csv = Join-Path $OutDir "keyword-discovery-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
$results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-Host "`n$($results.Count) hits -> $csv" -ForegroundColor Green

# Optional XLSX if ImportExcel is present
if (Get-Module -ListAvailable -Name ImportExcel) {
    $xlsx = $csv -replace "\.csv$",".xlsx"
    $results | Export-Excel -Path $xlsx -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName "Discovery"
    Write-Host "XLSX -> $xlsx" -ForegroundColor Green
}

# quick summary by keyword
$results | Group-Object Keyword | Select-Object Name,Count | Format-Table -AutoSize
