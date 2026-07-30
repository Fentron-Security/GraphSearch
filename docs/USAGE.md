# Usage Guide

Full reference for `Invoke-GraphKeywordDiscovery.ps1`. For setup, see [APP-REGISTRATION.md](APP-REGISTRATION.md); for handling rules, see [SECURITY.md](../SECURITY.md).

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ or 7.x | 7.x recommended |
| Entra ID app registration | App-only — see [APP-REGISTRATION.md](APP-REGISTRATION.md) |
| Admin consent | Granted on `Files.Read.All`, `Sites.Read.All` |
| `ImportExcel` module | Optional — enables XLSX output |

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

---

## Usage

### Tenant-wide (default)

```powershell
.\Invoke-GraphKeywordDiscovery.ps1 `
    -TenantId     $env:GRAPH_TENANT_ID `
    -ClientId     $env:GRAPH_CLIENT_ID `
    -ClientSecret $env:GRAPH_CLIENT_SECRET `
    -Keywords     "confidential","SSN","passport" `
    -Region       NAM
```

### Specific users

```powershell
.\Invoke-GraphKeywordDiscovery.ps1 `
    -TenantId $env:GRAPH_TENANT_ID -ClientId $env:GRAPH_CLIENT_ID `
    -ClientSecret $env:GRAPH_CLIENT_SECRET `
    -Scope    User `
    -Users    "jdoe@contoso.com","asmith@contoso.com" `
    -Keywords "severance","offer letter"
```

Feed a list from a file:

```powershell
$targets = Get-Content .\departing-users.txt
.\Invoke-GraphKeywordDiscovery.ps1 ... -Scope User -Users $targets -Keywords "confidential"
```

### Specific SharePoint sites

```powershell
.\Invoke-GraphKeywordDiscovery.ps1 `
    -TenantId $env:GRAPH_TENANT_ID -ClientId $env:GRAPH_CLIENT_ID `
    -ClientSecret $env:GRAPH_CLIENT_SECRET `
    -Scope    Site `
    -Sites    "https://contoso.sharepoint.com/sites/HR","https://contoso.sharepoint.com/sites/Legal" `
    -Keywords "SSN"
```

---

## Scope

`-Scope` sets the blast radius. Use the narrowest one the engagement allows.

| Scope | Searches | Required permissions |
|---|---|---|
| `Tenant` *(default)* | Every SharePoint site and every user's OneDrive | `Files.Read.All`, `Sites.Read.All`, `User.Read.All` *(enumerate mode)* |
| `User` | Only the named users' OneDrive | `Files.Read.All`, `User.Read.All` |
| `Site` | Only the named SharePoint sites | `Files.Read.All`, `Sites.Read.All` — or `Sites.Selected` + per-site grant |

`User` and `Site` scope always run per-drive search; `-Mode` applies to `Tenant` only and is ignored (with a warning) otherwise.

**Why it matters.** `Tenant` scope needs tenant-wide read across every mailbox owner's private OneDrive. `User` scope with `Sites.Selected` avoids `Sites.Read.All` entirely — much easier to get a client to sign off on, and a much smaller footprint if the credential is ever compromised.

**Site reference formats.** Both work:

```
https://contoso.sharepoint.com/sites/HR
contoso.sharepoint.com:/sites/HR
```

Users can be UPN or object ID. Users without a provisioned OneDrive are warned and skipped, not fatal.

---

### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-TenantId` | yes | — | Directory (tenant) ID |
| `-ClientId` | yes | — | Application (client) ID |
| `-ClientSecret` | yes | — | Client secret value |
| `-Keywords` | yes | — | One or more search terms (KQL supported). Multi-word terms are auto-quoted as exact phrases |
| `-Scope` | no | `Tenant` | `Tenant`, `User`, or `Site` |
| `-Users` | if `Scope User` | — | UPNs or object IDs |
| `-Sites` | if `Scope Site` | — | Site URLs |
| `-Region` | no | `NAM` | Graph search region — **required for tenant `search` mode**. `NAM`, `EUR`, `APC`, `GBR`, `AUS`, `CAN`, `IND`, `JPN` |
| `-Mode` | no | `search` | `search` or `enumerate` — `Tenant` scope only |
| `-NoPhraseQuoting` | no | off | Disable auto-quoting; treat multi-word terms as `AND` anywhere |
| `-OutDir` | no | `.\report` | Output directory |
| `-PageSize` | no | `500` | Results per page (max 500). Higher = fewer round trips = faster |

---

## Modes

Applies to `-Scope Tenant` only.

### `search` — index-based (default)

Uses `POST /search/query`. One call per keyword, tenant-wide, served from the Microsoft Search index.

- **Fast** — minutes, not hours
- Searches **file content and metadata**, not just filenames
- Best for most engagements

**Limits:** only indexed content is returned. Very recent uploads, unsupported file types, and content excluded from the index will be missed. Deep result paging is capped, so very broad terms should be narrowed with KQL.

### `enumerate` — drive-by-drive

Walks every SharePoint site drive and every user's OneDrive, running a per-drive search on each.

- **Exhaustive and deterministic** — you know exactly which drives were covered
- **Slow** — hours at large tenant scale, and heavily throttled
- Use as a fallback, for validation, or when defensible coverage matters (litigation hold, audit)

```powershell
.\Invoke-GraphKeywordDiscovery.ps1 -TenantId ... -Mode enumerate -Keywords "SSN"
```

---

## Multi-word keywords and performance

**Phrases.** A keyword with a space is wrapped in quotes and searched as an exact phrase by default. Without this, KQL reads `foo bar` as `foo AND bar` scattered anywhere in a document — the noisy, inconsistent behavior seen in early testing. `-NoPhraseQuoting` restores the old behavior when you actually want the AND semantics.

**Why some keywords are slow and others instant.** Runtime scales with hit count, not query complexity. A common term (`Personalnummer` in a German HR estate) returns thousands of hits across many pages, each page a round trip; a rare term finishes in one. Two levers:

- `-PageSize 500` (now the default) cuts the number of round trips.
- Narrow broad terms with KQL — `filetype`, `lastModifiedTime`, `path` — so you page through fewer junk hits. See the table below.

**Keyword ≠ context.** Even a fast, precise keyword hit only proves the *word* is present. To confirm a real identifier and see how it was used, run [Stage 2 context validation](CONTEXT-VALIDATION.md).

## KQL Query Refinement

The `-Keywords` values are passed to Microsoft Search as KQL. Use this to cut noise on broad terms.

| Goal | Example |
|---|---|
| File type | `"SSN filetype:xlsx"` |
| Multiple types | `"confidential (filetype:docx OR filetype:pdf)"` |
| Exact phrase | `'"employee roster"'` |
| Date range | `"merger lastModifiedTime>=2025-01-01"` |
| Single site | `"SSN path:https://contoso.sharepoint.com/sites/HR"` |
| Author | `"budget author:\"Jane Doe\""` |
| Exclusion | `"passport -template"` |
| Combined | `"SSN filetype:xlsx lastModifiedTime>=2024-01-01"` |

---

## Output

`report\keyword-discovery-<scope>-YYYYMMDD-HHMMSS.csv` (and `.xlsx` if `ImportExcel` is installed)

| Column | Description |
|---|---|
| `Keyword` | Which search term matched |
| `Owner` | Drive owner — UPN for OneDrive, site URL for SharePoint. `(tenant)` for index-based search |
| `FileName` | File name |
| `WebUrl` | Direct link to the file |
| `SiteId` | SharePoint site GUID |
| `DriveId` | Drive GUID |
| `Path` | Folder path within the drive |
| `SizeKB` | File size |
| `LastModified` | Last modified timestamp (UTC) |
| `ModifiedBy` | Last modifier display name |
| `CreatedBy` | Creator display name |
| `ItemId` | driveItem ID (for follow-up Graph calls) |

A per-keyword hit-count summary prints to console at the end.

---

## Operational Notes

**Throttling.** Graph returns 429s at tenant scale. The script honors `Retry-After` with exponential backoff. Expect long-running `enumerate` jobs to pause repeatedly — this is normal, not a failure.

**Token expiry.** Access tokens last ~60 min. The script auto-refreshes on 401, so long runs are safe.

**Duplicates.** A file matching several keywords appears once per keyword. De-dupe on `WebUrl` in Excel if you want a unique file list.

**Permissions on hits.** The report shows *where* files are, not *who can see them*. For exposure analysis, follow up with `GET /drives/{driveId}/items/{itemId}/permissions` using the `DriveId` and `ItemId` columns.

**Secret handling.** Don't hardcode the client secret. Prefer a certificate credential, Azure Key Vault, or environment variables. Delete the app registration when the engagement closes.

---

## Limitations

- **Keyword ≠ pattern matching.** This finds the literal string `SSN`, not an unlabeled 9-digit number inside a spreadsheet. For true sensitive-data-type detection (SSN, PAN, PHI patterns), layer **Microsoft Purview** sensitive info types / Content Explorer on top.
- **No OCR.** Scanned PDFs and images are only found if the tenant has OCR indexing enabled.
- **Encrypted / password-protected files** are not indexed and will not match.
- **Teams chat, Exchange, and Planner** are out of scope — this covers `driveItem` only. Change `entityTypes` to extend.

---

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| `403 Forbidden` on `/search/query` | Admin consent not granted, or `Region` missing from the request body |
| `401 Unauthorized` immediately | Wrong tenant/client ID, or expired secret |
| Zero results, no error | Content not indexed yet; try `-Mode enumerate` to confirm |
| `400 Bad Request` on search | Malformed KQL — check quoting and escaping |
| Constant 429s | Reduce `-PageSize`, run off-hours, or split keywords across runs |
| `enumerate` skips users | Users without a provisioned OneDrive are silently skipped (expected) |
| `-Scope User requires -Users` | Pass one or more UPNs with `-Users` |
| `No drives resolved` on User scope | UPN typo, user has no OneDrive, or `User.Read.All` not consented |
| `No drives resolved` on Site scope | Site URL wrong, or `Sites.Selected` granted without a per-site permission entry |
