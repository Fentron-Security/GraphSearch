# Entra ID App Registration

## 1. Create the registration

Entra admin center → **Applications → App registrations → New registration**

| Field | Value |
|---|---|
| Name | `Graph Keyword Discovery` |
| Supported account types | Accounts in this organizational directory only (single tenant) |
| Redirect URI | *(leave blank — app-only flow)* |

Record the **Application (client) ID** and **Directory (tenant) ID** from the Overview blade.

## 2. Grant API permissions

**API permissions → Add a permission → Microsoft Graph → Application permissions**

| Permission | Required for |
|---|---|
| `Files.Read.All` | All modes |
| `Sites.Read.All` | All modes |
| `User.Read.All` | `enumerate` mode only |

Then click **Grant admin consent for \<tenant\>**. Without consent every call returns `403`.

> **Narrower alternative:** replace `Sites.Read.All` with `Sites.Selected` and grant per-site access via `PUT /sites/{siteId}/permissions`. Reduces blast radius when the client will only authorize specific sites.

## 3. Credentials

### Option A — Certificate (recommended)

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=GraphKeywordDiscovery" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -NotAfter (Get-Date).AddMonths(6)

Export-Certificate -Cert $cert -FilePath .\graphkd-public.cer
$cert.Thumbprint
```

Upload `graphkd-public.cer` under **Certificates & secrets → Certificates → Upload certificate**. Keep the private key in the local cert store — never export it into the repo.

### Option B — Client secret

**Certificates & secrets → New client secret.** Copy the **Value** immediately; it is not shown again. Set the shortest expiry that covers the engagement.

## 4. Verify

```powershell
$body = @{
    client_id     = $env:GRAPH_CLIENT_ID
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $env:GRAPH_CLIENT_SECRET
    grant_type    = "client_credentials"
}
$token = (Invoke-RestMethod -Method POST `
    -Uri "https://login.microsoftonline.com/$env:GRAPH_TENANT_ID/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" -Body $body).access_token

Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?search=*&`$top=1" `
    -Headers @{ Authorization = "Bearer $token" }
```

A site object means auth and consent are working.

## 5. Decommission

When the engagement closes: delete the app registration, revoke the certificate, and confirm removal in the tenant's Enterprise applications list. A dormant tenant-wide-read registration is a standing risk.
