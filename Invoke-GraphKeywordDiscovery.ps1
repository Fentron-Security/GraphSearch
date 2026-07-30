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
    Keyword discovery across OneDrive and SharePoint via Microsoft Graph.
    App-only (client credentials). Outputs a location report (CSV + XLSX).

.DESCRIPTION
    Scope controls the blast radius of the search:

      Tenant  (default) Every SharePoint site and every user's OneDrive.
      User              Only the OneDrive(s) of the named user(s).
      Site              Only the named SharePoint site(s).

    Mode controls how a Tenant-scoped search is executed. User and Site
    scope always use per-drive search, since the drive set is already known.

.NOTES
    Entra app registration requires APPLICATION permissions (admin consent):

      Scope=Tenant   Files.Read.All, Sites.Read.All, User.Read.All (enumerate mode)
      Scope=User     Files.Read.All, User.Read.All
      Scope=Site     Files.Read.All, Sites.Read.All  (or Sites.Selected + per-site grant)

    Narrower scope needs fewer permissions. Prefer the smallest scope the
    engagement allows.

.EXAMPLE
    # Whole tenant
    .\Invoke-GraphKeywordDiscovery.ps1 -TenantId $env:GRAPH_TENANT_ID `
        -ClientId $env:GRAPH_CLIENT_ID -ClientSecret $env:GRAPH_CLIENT_SECRET `
        -Keywords "confidential","SSN"

.EXAMPLE
    # Two specific users
    .\Invoke-GraphKeywordDiscovery.ps1 -TenantId $env:GRAPH_TENANT_ID `
        -ClientId $env:GRAPH_CLIENT_ID -ClientSecret $env:GRAPH_CLIENT_SECRET `
        -Scope User -Users "jdoe@contoso.com","asmith@contoso.com" `
        -Keywords "severance"

.EXAMPLE
    # One SharePoint site
    .\Invoke-GraphKeywordDiscovery.ps1 -TenantId $env:GRAPH_TENANT_ID `
        -ClientId $env:GRAPH_CLIENT_ID -ClientSecret $env:GRAPH_CLIENT_SECRET `
        -Scope Site -Sites "https://contoso.sharepoint.com/sites/HR" `
        -Keywords "SSN"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$ClientSecret,
    [Parameter(Mandatory)][string[]]$Keywords,

    # Tenant = everything. User = named users' OneDrive only. Site = named sites only.
    [ValidateSet("Tenant","User","Site")][string]$Scope = "Tenant",

    # Required when -Scope User. UPN or object ID. Accepts pipeline-friendly arrays.
    [string[]]$Users,

    # Required when -Scope Site. Full site URL, or hostname:/sites/name form.
    [string[]]$Sites,

    # Graph app-only search region. Required for Tenant + search mode.
    [string]$Region = "NAM",

    # Applies to Scope=Tenant only.
    #   search    = index-based, fast, tenant-wide
    #   enumerate = walk every drive, exhaustive, slow
    [ValidateSet("search","enumerate")][string]$Mode = "search",

    # By default a keyword containing a space is wrapped in quotes and searched
    # as an exact phrase. Without this, KQL treats "foo bar" as foo AND bar
    # anywhere in the document - noisy and inconsistent. Set this switch to
    # keep the old AND behavior.
    [switch]$NoPhraseQuoting,

    [string]$OutDir = ".\report",

    # Larger pages = fewer round trips = faster on high-hit keywords. Max 500.
    [int]$PageSize = 500
)

$ErrorActionPreference = "Stop"

# ---------- Parameter validation ----------
if ($Scope -eq "User" -and -not $Users) {
    throw "-Scope User requires -Users (one or more UPNs or object IDs)."
}
if ($Scope -eq "Site" -and -not $Sites) {
    throw "-Scope Site requires -Sites (one or more site URLs)."
}
if ($Scope -ne "Tenant" -and $PSBoundParameters.ContainsKey('Mode')) {
    Write-Warning "-Mode applies to -Scope Tenant only; ignored for Scope=$Scope."
}
if ($PageSize -gt 500) { Write-Warning "PageSize capped at 500 by Graph."; $PageSize = 500 }

# ---------- Keyword normalization ----------
# Multi-word keyword -> exact phrase, unless the caller already quoted it or
# used KQL operators, or opted out with -NoPhraseQuoting.
function Format-Keyword {
    param([string]$Keyword)
    if ($NoPhraseQuoting) { return $Keyword }
    $k = $Keyword.Trim()
    if ($k -match '\s' -and $k -notmatch '["():]' -and $k -notmatch '\b(AND|OR|NOT)\b') {
        return '"{0}"' -f $k
    }
    return $k
}

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
                $script:Token   = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
                $script:Headers = @{ Authorization = "Bearer $script:Token"; "Content-Type" = "application/json" }
                continue
            }
            if ($code -eq 429 -or $code -ge 500) {
                $wait = [int]($_.Exception.Response.Headers["Retry-After"]); if (-not $wait) { $wait = [math]::Pow(2,$i) }
                Write-Warning "Throttled/err $code - retry in $wait s"; Start-Sleep -Seconds $wait; continue
            }
            throw
        }
    }
    throw "Graph call failed after retries: $Uri"
}

$results = New-Object System.Collections.Generic.List[object]

function Add-Hit {
    param($item,$keyword,$owner)
    $pr = $item.parentReference
    $results.Add([pscustomobject]@{
        Keyword      = $keyword
        Owner        = $owner
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

# ---------- Drive resolution ----------
function Resolve-UserDrives {
    param([string[]]$Users)
    $drives = New-Object System.Collections.Generic.List[object]
    foreach ($u in $Users) {
        try {
            $d = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$u/drive"
            $drives.Add([pscustomobject]@{ Id=$d.id; Owner=$u })
            Write-Host "  resolved $u" -ForegroundColor DarkGray
        } catch {
            Write-Warning "$u - no OneDrive provisioned or not found; skipped."
        }
    }
    return $drives
}

function Resolve-SiteDrives {
    param([string[]]$Sites)
    $drives = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Sites) {
        # Accept full URL or hostname:/sites/name
        if ($s -match '^https?://([^/]+)(/.*)$') {
            $siteRef = "$($Matches[1]):$($Matches[2])"
        } else {
            $siteRef = $s
        }
        try {
            $site = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/sites/$siteRef"
            $d = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$top=100"
            foreach ($drv in $d.value) { $drives.Add([pscustomobject]@{ Id=$drv.id; Owner=$site.webUrl }) }
            Write-Host "  resolved $($site.webUrl) ($($d.value.Count) drive(s))" -ForegroundColor DarkGray
        } catch {
            Write-Warning "$s - site not found or no access; skipped."
        }
    }
    return $drives
}

function Get-AllTenantDrives {
    $drives = New-Object System.Collections.Generic.List[object]

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

    $uri = "https://graph.microsoft.com/v1.0/users?`$top=100&`$select=id,userPrincipalName"
    do {
        $r = Invoke-Graph -Uri $uri
        foreach ($u in $r.value) {
            try {
                $d = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$($u.id)/drive"
                $drives.Add([pscustomobject]@{ Id=$d.id; Owner=$u.userPrincipalName })
            } catch { }
        }
        $uri = $r.'@odata.nextLink'
    } while ($uri)

    return $drives
}

# ---------- Search implementations ----------
function Search-IndexWide {
    param([string[]]$Keywords,[string]$Region,[int]$PageSize)
    foreach ($kw in $Keywords) {
        $q = Format-Keyword $kw
        Write-Host "[search] $q" -ForegroundColor Cyan
        $from = 0
        do {
            $body = @{
                requests = @(@{
                    entityTypes = @("driveItem")
                    query       = @{ queryString = $q }
                    from        = $from
                    size        = $PageSize
                    region      = $Region
                    fields      = @("name","webUrl","parentReference","lastModifiedDateTime",
                                    "lastModifiedBy","createdBy","size","id")
                })
            }
            $resp = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/search/query" -Method POST -Body $body
            $container = $resp.value[0].hitsContainers[0]
            foreach ($hit in $container.hits) { Add-Hit -item $hit.resource -keyword $kw -owner "(tenant)" }
            $more = $container.moreResultsAvailable
            $from += $PageSize
        } while ($more)
    }
}

function Search-Drives {
    param($Drives,[string[]]$Keywords)
    $n = 0
    foreach ($drv in $Drives) {
        $n++
        Write-Host "[$n/$($Drives.Count)] $($drv.Owner)" -ForegroundColor Cyan
        foreach ($kw in $Keywords) {
            $q = Format-Keyword $kw
            $qEnc = [uri]::EscapeDataString($q)
            $uri = "https://graph.microsoft.com/v1.0/drives/$($drv.Id)/root/search(q='$qEnc')?`$top=200"
            do {
                try { $r = Invoke-Graph -Uri $uri } catch { break }
                foreach ($item in $r.value) { Add-Hit -item $item -keyword $kw -owner $drv.Owner }
                $uri = $r.'@odata.nextLink'
            } while ($uri)
        }
    }
}

# ---------- Dispatch ----------
Write-Host "Scope: $Scope" -ForegroundColor Yellow

switch ($Scope) {
    "User" {
        Write-Host "Resolving $($Users.Count) user drive(s)..." -ForegroundColor Yellow
        $drives = Resolve-UserDrives -Users $Users
        if (-not $drives.Count) { throw "No drives resolved. Check the UPNs and Files.Read.All / User.Read.All consent." }
        Search-Drives -Drives $drives -Keywords $Keywords
    }
    "Site" {
        Write-Host "Resolving $($Sites.Count) site(s)..." -ForegroundColor Yellow
        $drives = Resolve-SiteDrives -Sites $Sites
        if (-not $drives.Count) { throw "No drives resolved. Check the site URLs and Sites.Read.All / Sites.Selected consent." }
        Search-Drives -Drives $drives -Keywords $Keywords
    }
    "Tenant" {
        if ($Mode -eq "search") {
            Search-IndexWide -Keywords $Keywords -Region $Region -PageSize $PageSize
        } else {
            Write-Host "Enumerating all tenant drives..." -ForegroundColor Yellow
            $drives = Get-AllTenantDrives
            Write-Host "$($drives.Count) drives" -ForegroundColor Yellow
            Search-Drives -Drives $drives -Keywords $Keywords
        }
    }
}

# ---------- Report ----------
$results = $results | Sort-Object Keyword,WebUrl -Unique

$stamp = Get-Date -Format yyyyMMdd-HHmmss
$csv = Join-Path $OutDir "keyword-discovery-$($Scope.ToLower())-$stamp.csv"
$results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-Host "`n$($results.Count) hits -> $csv" -ForegroundColor Green

if (Get-Module -ListAvailable -Name ImportExcel) {
    $xlsx = $csv -replace "\.csv$",".xlsx"
    $results | Export-Excel -Path $xlsx -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName "Discovery"
    Write-Host "XLSX -> $xlsx" -ForegroundColor Green
}

$results | Group-Object Keyword | Select-Object Name,Count | Format-Table -AutoSize
