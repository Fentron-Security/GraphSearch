<h1 align="center">GraphSearch</h1>

<p align="center">
  Enterprise-wide keyword discovery across OneDrive and SharePoint via Microsoft Graph.
</p>

<p align="center">
  <a href="LICENSE"><img alt="License: Apache 2.0" src="https://img.shields.io/badge/license-Apache%202.0-blue.svg"></a>
  <img alt="PowerShell 5.1+" src="https://img.shields.io/badge/PowerShell-5.1%2B-5391FE.svg">
  <img alt="Read-only" src="https://img.shields.io/badge/access-read--only-brightgreen.svg">
  <a href="../../actions"><img alt="CI" src="../../actions/workflows/ci.yml/badge.svg"></a>
</p>

---

Search every OneDrive and SharePoint site in a Microsoft 365 tenant for keywords, then export a report of **where matching files live and who owns them**.

Built for data discovery, DLP pre-assessments, eDiscovery scoping, and compliance sweeps.

```powershell
.\Invoke-GraphKeywordDiscovery.ps1 `
    -TenantId     $env:GRAPH_TENANT_ID `
    -ClientId     $env:GRAPH_CLIENT_ID `
    -ClientSecret $env:GRAPH_CLIENT_SECRET `
    -Keywords     "confidential","SSN","passport" `
    -Region       NAM
```

## Features

- **Tenant-wide** — every SharePoint site and every user's OneDrive in a single run
- **Content search, not filename search** — matches text *inside* documents via the Microsoft Search index
- **Two modes** — fast index-based search, or exhaustive drive-by-drive enumeration for defensible coverage
- **KQL support** — narrow by file type, date range, author, or site
- **Location + ownership report** — CSV always, XLSX when `ImportExcel` is installed
- **Read-only by design** — the tool never writes to, modifies, or deletes tenant content
- **Throttle-tolerant** — token auto-refresh and `Retry-After` backoff for long tenant-scale runs

## Quick start

**1. Register an Entra ID app** with these **Application** permissions, then grant admin consent:

| Permission | Purpose |
| --- | --- |
| `Files.Read.All` | Read files across all drives |
| `Sites.Read.All` | Read all SharePoint sites |
| `User.Read.All` | `enumerate` mode only — resolve users to their OneDrive |

Full walkthrough, including certificate auth: **[docs/APP-REGISTRATION.md](docs/APP-REGISTRATION.md)**

**2. Configure credentials**

```powershell
Copy-Item .env.example .env   # then fill it in — .env is gitignored
```

**3. Optional — XLSX output**

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

**4. Run.** Full parameter reference, KQL syntax, and troubleshooting: **[docs/USAGE.md](docs/USAGE.md)**

## Modes

| Mode | Speed | Coverage | Use when |
| --- | --- | --- | --- |
| `search` *(default)* | Minutes | Indexed content only | Most engagements |
| `enumerate` | Hours | Every drive, deterministic | Litigation hold, audit, validating index gaps |

## Output

CSV (and XLSX) written to `report/keyword-discovery-YYYYMMDD-HHMMSS.csv`:

`Keyword` · `FileName` · `WebUrl` · `SiteId` · `DriveId` · `Path` · `SizeKB` · `LastModified` · `ModifiedBy` · `CreatedBy` · `ItemId`

Sample: [examples/sample-report.csv](examples/sample-report.csv)

## Limitations

- **Keyword ≠ pattern matching.** Finds the literal string `SSN`, not an unlabeled 9-digit number. For real sensitive-data-type detection, layer **Microsoft Purview** sensitive info types on top.
- **No OCR** unless the tenant has OCR indexing enabled.
- **Encrypted / password-protected files** are not indexed and will not match.
- **Scope is `driveItem` only** — Exchange, Teams chat, and Planner are out of scope.

## Security

This tool can read **every file in a tenant**. Get written authorization before running it against an environment you do not own, and read **[SECURITY.md](SECURITY.md)** first.

Never commit `.env`, certificates, or generated reports — the included `.gitignore` blocks all three.

## Documentation

| Doc | Contents |
| --- | --- |
| [docs/USAGE.md](docs/USAGE.md) | Parameters, modes, KQL reference, output schema, troubleshooting |
| [docs/APP-REGISTRATION.md](docs/APP-REGISTRATION.md) | Entra ID setup, certificate and secret auth, decommissioning |
| [SECURITY.md](SECURITY.md) | Authorization, credential handling, output handling, disclosure |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Style, linting, PR scope |

## License

Apache License 2.0 — see [LICENSE](LICENSE) and [NOTICE](NOTICE).
