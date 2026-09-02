# Exchange Online empty archive folder cleanup

Safely report and remove empty user-created folders from an Exchange Online archive mailbox with the EWS Managed API.

The script runs in report-only mode unless you add `-Delete`. Before deleting a folder, it checks the folder again, protects Exchange system folders, and processes eligible folders from deepest to shallowest.

> [!WARNING]
> Microsoft begins disabling EWS in Exchange Online in October 2026. This script is intended for short-term, one-off migration cleanup. Don't use it as the foundation for a new long-term service.

## Safety controls

The script:

- Reports candidates by default
- Requires `-Delete` for removal
- Supports `-WhatIf` and `-Confirm`
- Protects archive roots and well-known folders
- Excludes hidden folders
- Excludes Deleted Items and Recoverable Items hierarchies
- Checks normal and associated item counts
- Revalidates each folder immediately before deletion
- Hard-deletes eligible folders from deepest to shallowest
- Records candidates, attempts, skips, successes, and failures in an optional CSV log

## Prerequisites

- Windows PowerShell 5.1 or later
- An Exchange Online archive mailbox
- The EWS Managed API 2.2 assembly, `Microsoft.Exchange.WebServices.dll`
- The [MSAL.PS](https://www.powershellgallery.com/packages/MSAL.PS) module
- A Microsoft Entra app registration
- A certificate with a private key
- Exchange Online application permission `full_access_as_app`
- Administrator consent for the application permission

## Configure authentication

### 1. Create a certificate

Create or obtain a certificate suitable for application authentication. Install the certificate, including its private key, in one of these stores:

- `Cert:\CurrentUser\My`
- `Cert:\LocalMachine\My`

Record its thumbprint.

### 2. Register the application

In the Microsoft Entra admin center:

1. Go to **Identity > Applications > App registrations**.
2. Select **New registration**.
3. Enter a name, such as `EWS archive folder cleanup`.
4. Create the registration and record its application (client) ID and directory (tenant) ID.
5. Open **Certificates & secrets > Certificates**.
6. Upload the public certificate.

### 3. Grant Exchange Online permission

1. Open **API permissions** for the app registration.
2. Select **Add a permission > APIs my organization uses**.
3. Find **Office 365 Exchange Online**.
4. Select **Application permissions**.
5. Add `full_access_as_app`.
6. Select **Grant admin consent**.

This permission grants broad mailbox access. Restrict the app to the required mailbox or mail-enabled security group with an Exchange Application Access Policy where supported.

```powershell
Connect-ExchangeOnline

New-ApplicationAccessPolicy `
  -AppId '<app-id>' `
  -PolicyScopeGroupId 'ews-archive-cleanup@contoso.com' `
  -AccessRight RestrictAccess `
  -Description 'Restrict EWS archive cleanup access'

Test-ApplicationAccessPolicy `
  -AppId '<app-id>' `
  -Identity 'user@contoso.com'
```

Policy changes can take time to apply.

## Install dependencies

Install MSAL.PS:

```powershell
Install-Module MSAL.PS -Scope CurrentUser
```

Obtain the EWS Managed API 2.2 assembly and note the path to either its installation directory or `Microsoft.Exchange.WebServices.dll`.

## Run a report

Start with report-only mode:

```powershell
.\Remove-EmptyArchiveFolders.ps1 `
  -Mailbox 'user@contoso.com' `
  -TenantId '<tenant-id>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<certificate-thumbprint>' `
  -EwsManagedApiPath 'C:\Path\To\Microsoft.Exchange.WebServices.dll' `
  -LogPath '.\archive-folder-report.csv'
```

Review the `Candidate` rows in the console output or CSV file.

## Preview deletion

Use `-Delete -WhatIf` to preview the deletion workflow without changing the mailbox:

```powershell
.\Remove-EmptyArchiveFolders.ps1 `
  -Mailbox 'user@contoso.com' `
  -TenantId '<tenant-id>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<certificate-thumbprint>' `
  -EwsManagedApiPath 'C:\Path\To\Microsoft.Exchange.WebServices.dll' `
  -Delete `
  -WhatIf
```

## Delete eligible folders

After reviewing the report and preview, run:

```powershell
.\Remove-EmptyArchiveFolders.ps1 `
  -Mailbox 'user@contoso.com' `
  -TenantId '<tenant-id>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<certificate-thumbprint>' `
  -EwsManagedApiPath 'C:\Path\To\Microsoft.Exchange.WebServices.dll' `
  -Delete `
  -LogPath '.\archive-folder-deletion.csv'
```

PowerShell requests confirmation because deletion has a high impact. Add `-Confirm:$false` only when you intentionally want unattended deletion.

## Auto-expanding archive limitation

Folders can reside in auxiliary archive mailboxes when auto-expanding archiving is enabled. The script doesn't follow cross-archive redirects automatically.

Exchange Online also prevents deletion of folders beneath Deleted Items when those folders are hosted in an auto-expanding archive. This behavior is by design.

## Troubleshooting

- **Archive root unavailable:** Confirm that the mailbox has an archive and that the app can access it.
- **Access denied:** Check `full_access_as_app`, administrator consent, the certificate, and any Application Access Policy.
- **Folder skipped:** Review the `Reason` field. The folder might contain items, associated data, child folders, or protected content.
- **Redirect or moved error:** The folder might be in an auxiliary auto-expanding archive. The script stops rather than attempting unsafe routing.
- **Module or assembly missing:** Confirm that MSAL.PS is installed and that `-EwsManagedApiPath` points to the EWS DLL or its directory.

## License

Licensed under the MIT License. See [LICENSE](LICENSE).
