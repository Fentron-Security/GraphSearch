# Security

## What this tool can do

With `Files.Read.All` and `Sites.Read.All`, this tool can read **every file in the tenant** — every user's OneDrive, every SharePoint site, including HR, legal, and executive content. Treat the app registration as a privileged asset.

## Operating rules

**Authorization.** Get written authorization before running against any environment you do not own. Tenant-wide file access is in scope-of-work territory, not "just a script."

**Credentials.**
- Prefer **certificate credentials** over client secrets.
- Never hardcode credentials. Use environment variables, Azure Key Vault, or a certificate store.
- Never commit `.env`, `.pfx`, `.pem`, or `.key` files. The `.gitignore` blocks them — verify with `git status`.

**Lifecycle.** Delete the app registration when the engagement closes. A dormant tenant-wide-read app registration is a standing risk.

**Least privilege.** For narrower scope, use [`Sites.Selected`](https://learn.microsoft.com/en-us/graph/permissions-reference) with per-site grants instead of `Sites.Read.All`.

**Output handling.** Report files contain filenames, paths, and user names from the client environment — this is client confidential data and may be regulated. Store it per your engagement's data-handling requirements, and never commit it to source control.

**Auditing.** Graph calls are logged in the tenant's Entra sign-in and audit logs. Assume the client's SOC will see this activity — brief them in advance to avoid a false-positive incident.

## Reporting a vulnerability

Open a private security advisory via GitHub's **Security → Report a vulnerability**, or email the maintainer. Please do not open a public issue for security problems.
