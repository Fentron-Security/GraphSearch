# Contributing

## Before you open a PR

1. Run PSScriptAnalyzer — CI enforces it:
   ```powershell
   Install-Module PSScriptAnalyzer -Scope CurrentUser
   Invoke-ScriptAnalyzer -Path .\Invoke-GraphKeywordDiscovery.ps1 -Severity Error,Warning
   ```
2. Confirm no secrets, tenant IDs, real domains, client names, or report files are in the diff.
3. Test against a dev tenant. Do not test against a client environment.

## Style

- Approved PowerShell verbs (`Get-Verb`)
- `[CmdletBinding()]` on all functions
- No `Write-Host` for data output — use the pipeline; `Write-Host` is for status only
- Handle 429 and 401 on every new Graph call, or route it through `Invoke-Graph`

## Scope

In scope: new entity types, output formats, auth methods, KQL helpers, throttling improvements.

Out of scope: anything that writes to, modifies, or deletes tenant content. This tool is read-only by design and stays that way.

## Reporting bugs

Include: PowerShell version, mode used, redacted error output, and whether the tenant is multi-geo. Never paste real tenant IDs, URLs, or file paths into an issue.
