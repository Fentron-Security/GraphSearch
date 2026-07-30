# Context Validation (Stage 2)

## The problem this solves

The Microsoft Search index can only tell you **a keyword appears in a file**. It cannot tell you:

- whether `Sozialversicherungsnummer` sits next to an actual social-insurance number or just the word in a policy document,
- whether a number that looks like an ID is real or a coincidence,
- what the surrounding context was.

Microsoft Search has no regex and no checksum awareness, so a keyword sweep over a German/Swiss tenant returns a large pile of hits with no way to separate real exposure from noise. That was the core of the field feedback.

## The two-stage design

```
Stage 1  Invoke-GraphKeywordDiscovery.ps1   index sweep -> file locations (fast, broad)
Stage 2  Invoke-ContextValidation.ps1       download hits -> regex + checksum -> context (precise)
```

Stage 1 narrows millions of files to the few hundred that mention your terms. Stage 2 downloads only those, reads the actual bytes locally, and runs full pattern matching where the index couldn't. You get the index's speed on the first pass and regex precision on the second, without reading every file in the tenant.

## Running it

```powershell
# 1. Locate
.\Invoke-GraphKeywordDiscovery.ps1 `
    -TenantId $env:GRAPH_TENANT_ID -ClientId $env:GRAPH_CLIENT_ID `
    -ClientSecret $env:GRAPH_CLIENT_SECRET `
    -Keywords "Sozialversicherungsnummer","Personalnummer"

# 2. Validate the hits from that report
.\Invoke-ContextValidation.ps1 `
    -TenantId $env:GRAPH_TENANT_ID -ClientId $env:GRAPH_CLIENT_ID `
    -ClientSecret $env:GRAPH_CLIENT_SECRET `
    -ReportCsv .\report\keyword-discovery-tenant-20260730-120000.csv `
    -Country DE,CH,NL,US
```

### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-ReportCsv` | yes | — | A Stage 1 output CSV. Must have `DriveId` and `ItemId` columns |
| `-Validators` | no | — | Select validators by name |
| `-Country` | no | — | Select by jurisdiction (DE, CH, US, UK, NL, ES, CA, Global) |
| `-Category` | no | — | Select by data class (payment, bank, national-id, employee-id) |

With none of the three set, all validators run.
| `-Context` | no | `60` | Characters of surrounding text captured per match |
| `-MaxFileMB` | no | `25` | Skip files larger than this |
| `-OutDir` | no | `.\report` | Output directory |
| `-KeepFiles` | no | off | Keep downloaded files instead of deleting after extraction |

## Validators

The tool assumes a **global** use case. Validators are data-driven definitions, each tagged with a country and a category, so a run can be scoped by name, jurisdiction, or data class.

| Name | Country | Category | Identifier | Validation |
|---|---|---|---|---|
| `CreditCard` | Global | payment | Payment card (PAN) | Luhn + length 13-19 |
| `IBAN` | Global | bank | Bank account (~80 countries) | mod-97 |
| `US-SSN` | US | national-id | Social Security Number | format + excluded ranges |
| `UK-NINo` | UK | national-id | National Insurance Number | format + prefix rules |
| `DE-RVNR` | DE | national-id | Renten-/Sozialversicherungsnummer | Modulo-10 checksum |
| `CH-AHV` | CH | national-id | AHV/AVS number | EAN-13 checksum |
| `NL-BSN` | NL | national-id | Burgerservicenummer | 11-proof |
| `ES-DNI` | ES | national-id | DNI / NIE | mod-23 letter |
| `CA-SIN` | CA | national-id | Social Insurance Number | Luhn |
| `DE-Personalnummer` | DE | employee-id | Employee number | **heuristic — no checksum** |

Every checksum was verified against a known-valid reference number before shipping. `US-SSN` and `UK-NINo` have no national check digit, so they validate on structure plus the officially excluded ranges/prefixes.

This is a **framework, not a claim to cover every country**. Adding a jurisdiction (France NIR, Italy Codice Fiscale, Poland PESEL, Brazil CPF, etc.) is one entry in `$script:Definitions` plus, if its checksum isn't already present, one primitive. The reusable primitives shipped — Luhn, IBAN mod-97, 11-proof, EAN-13 — already cover a large share of the world's national IDs and all payment/bank data.

### Selecting what to run

Scope a run three ways, combinable:

```powershell
-Country DE,CH,NL,US          # by name
-Country    DE,CH,NL                 # by jurisdiction
-Category   payment,bank             # by data class
```

With no selector, **all** validators run. Prefer scoping a real engagement — see the false-positive note below.

### False positives on checksum-only identifiers

Some identifiers are a bare run of digits distinguished only by a checksum — `NL-BSN` (any 9 digits passing the 11-proof), `CreditCard` (any 13-19 digit Luhn-valid run), `CA-SIN`. A checksum cuts the noise but doesn't eliminate it: an unrelated order reference or GUID fragment can pass by chance. This is inherent, not a defect.

Two mitigations, both built in: `ChecksumValid` lets you filter to checksum-passing matches, and the `Context` column shows the surrounding text so an analyst can confirm the number is what it looks like. For checksum-only types, **always review Context before reporting a hit as real** — scope by country/category to keep the review set small.

### Verified reference numbers

| Validator | Reference | Result |
|---|---|---|
| `DE-RVNR` | `65121267H008` | valid (check 8) |
| `CH-AHV` | `756.1234.5678.97` | valid (check 7) |
| `NL-BSN` | `111222333` | valid |
| `ES-DNI` | `12345678Z` | valid |
| `CreditCard` | `4111111111111111` | valid (Luhn) |
| `IBAN` | `GB82 WEST 1234 5698 7654 32` | valid (mod-97) |

Published "example" national IDs frequently use fabricated check digits and will correctly fail (e.g. the German `65 170439 K 001`). That is the validator working, not a bug.

### Personalnummer — why there is no validator

A Personalnummer is an **employer-assigned** number with no national format or checksum. `DE-Personalnummer` runs a label-adjacent heuristic only, returning candidates with a blank `ChecksumValid`. Real detection needs a per-client pattern, added as a custom validator.

## Output

`report\context-validation-YYYYMMDD-HHMMSS.csv`:

| Column | Description |
|---|---|
| `FileName` | File the match was found in |
| `WebUrl` | Direct link |
| `Owner` | Drive owner (UPN or site URL) carried from Stage 1 |
| `Type` | Which validator matched |
| `IdCountry` | Country the identifier belongs to |
| `Category` | `payment`, `bank`, `national-id`, or `employee-id` |
| `Match` | The matched string |
| `ChecksumValid` | `True` = format **and** checksum valid; `False` = format matched but checksum failed; blank = no checksum exists |
| `Offset` | Character offset of the match within the extracted text |
| `Context` | Surrounding text (`±Context` chars) so an analyst can judge relevance |
| `DriveId` / `ItemId` | For follow-up Graph calls |

The console summary splits findings three ways: checksum-valid, format-match-but-checksum-invalid, and no-checksum. **Checksum-valid is your real-exposure count** — the number worth reporting to a client.

## Supported file types

Native extraction: `txt`, `csv`, `tsv`, `log`, `md`, `json`, `xml`, `html`, `docx`, `xlsx`, `pptx`.

**PDF** requires [`pdftotext`](https://poppler.freedesktop.org/) (poppler) on `PATH`. Without it, PDFs are skipped with a warning. On Windows: `winget install oschwartz10612.Poppler` or add a poppler build to `PATH`.

**Not supported:** legacy binary `doc` / `xls` / `ppt`, and anything encrypted or password-protected.

## Adding a custom validator

Validators live in `GraphSearch.Validators.psm1` as entries in `$script:Definitions`. Each is a `[pscustomobject]` with `Name`, `Country`, `Category`, `Description`, a `Pattern` (regex to find candidates), and a `Validate` scriptblock returning `$true`/`$false`. If the identifier uses a checksum already implemented (Luhn, mod-97, 11-proof, EAN-13), reuse that primitive; otherwise add one alongside the others.

```powershell
[pscustomobject]@{
    Name='FR-NIR'; Country='FR'; Category='national-id'
    Description='French social security number (mod-97)'
    Pattern='\b[12]\s?\d{2}\s?\d{2}\s?\d{2,3}\s?\d{3}\s?\d{3}\s?\d{2}\b'
    Validate={ param($v) <# 97 - (first13 mod 97) == last2 #> }
}
```

No other file changes are needed — selection by name, country, and category picks it up automatically.

## Privacy note

Stage 2 downloads real file content, which for a positive result means real personal data lands on the machine running the script. Files go to a temp directory and are deleted after extraction unless `-KeepFiles` is set. The output report itself contains matched identifiers and context snippets — treat it as regulated client data and handle it per `SECURITY.md`.
